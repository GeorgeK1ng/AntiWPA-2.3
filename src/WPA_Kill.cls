VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WPA_Kill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'//////////////////////////////////////////////////////
'/// WPA-Kill - Crypts SelfCheckBlocks
'//
'// Structure of a selfcheckblock:
'//   CrypBlockStub           ;Contains ponter to DecryptionDataHead
'//   CryData                 ;Contains crypted functiondata
'//      CryDataBlock[0..x]   ;Contains crypted functiondata
'//   DecryptionDataHeads
'//      Key1 , Key2          ;Key for Decryption
'//      RelocChunk [0..x]    ;information for relocation Fixup

Private NOP$
Private JMP_EAX$
Private PUSH_XXX$
Private CALL_XXX$
Private OFFSET_RECOGNITION_PATTERN$
Private DATA_TO_KILL_CHECKSUMFUNTION$

Private Possible_Patch_Offset    As New VA_TextSectionPtr
Private isSP2         As Boolean
Private ptrToDecryptionDataHeads As New VA_TextSectionPtr
Private ptrToCheckSumFunction    As New VA_TextSectionPtr
Private ptrExecFunction    As New VA_TextSectionPtr
Private Start_Of_Text_Section    As New VA_TextSectionPtr
Private input_buffer As New StringReader

Private LengthOfFirstChecksumFunctionCall&
Private LengthOfExecFunctionCall&

Dim Key1&, Key2&
Dim CryDataBlock As New StringReader
'Dim bCheckSumFunctionSet As Boolean
Dim src&, dest&
Dim CheckSumFunctionCounter&
Dim RelocTranslation&

Dim CryptblockCount&
Dim MapFile As New FileStream
Private Sub MapFileOpen()
   On Error Resume Next
   If frmMain.mi_MAP.Checked Then
      MapFile.Create Filename.Path & Filename.Name & ".map", True
   End If
End Sub
Public Sub MapFileClose()
   On Error Resume Next
   If frmMain.mi_MAP.Checked Then
      MapFile.CloseFile
      If CryptblockCount Then
         frmMain.displn "Å Added " & CryptblockCount & " symbols to Map-file:" 'CurDir$ & "\" &
         frmMain.displn MapFile.Filename
      End If
   End If
End Sub

Public Sub MapFileAddEntry(address, Name$)
   On Error Resume Next
   If frmMain.mi_MAP.Checked Then
      MapFile.FixedString(-1) = " 0001:" & H32(address) & Space(7) & Name & vbCrLf
   End If
End Sub



'/////////////////////////////////////////////////////
'/// ScanForCrypBlockStub - Scans for next CrypBlock
'//
Private Function ScanForCrypBlockStub() As Boolean
   ' ----------- Search for Offset ------------------
   '
   'Example: Decryption Block Stub (SP2 build 2082..)
   '0104C31C  9C            PUSHFD            <- P_GENERAL_ORG Searchstring
   '0104C31D  60            PUSHAD            <- P_GENERAL_ORG Searchstring
   '0104C31E  51            PUSH      ECX
   '0104C31F  6800C70401    PUSH      0104C700   ;end of decrpytiondata; Start of Decrpytiondataheads
   '0104C324  6A01          PUSH      01
   '0104C326  E82635FCFF    CALL      0100F851
   '0104C32B  FFE0          JMP       EAX     <- additional_check Searchstring
   '0104C32D  90            NOP                  ;Start of decrpytiondata
   
   '....
        
   'Example: Decryption Block Stub (SP1 build ..1106):
   '9C            PUSHFD
   '60            PUSHAD
   '                                <- no (PUSH  ECX) here (this was used in version 2.0 to distinguish versions - changed because winLH was also detected as SP2
   '68 59F90301   PUSH    0103F959     it's important to distinguish versions when calculating relocation Fixups
   '68 01000000   PUSH    1
   'E8 E3E7FCFF   CALL    0100DF5A
   '83C4 04       ADD     ESP, 4
   'FFE0          JMP     EAX
   '??                              this is currently used to distinguish versions
   '                                it's important to distinguish versions when calculating relocation Fixups



   With input_buffer

   
     
    ' Find Possible_Patch_Offset
      .DisableAutoMove = True
         Possible_Patch_Offset.TEXT_Rel_Offset = .Findstring(P_GENERAL_ORG)
      .DisableAutoMove = False
      '0104231C  9C           * PUSHFD       <- current position
      '0104231D  60             PUSHAD
      '0104231E  51             PUSH    ECX
   
   
    ' Look for already Patched string between new and old Block
      Dim tmp_AlreadyPatchedOffset&
      tmp_AlreadyPatchedOffset = .Findstring(P_GENERAL_CRK, IIf(Possible_Patch_Offset.TEXT_Rel_Offset, Possible_Patch_Offset.TEXT_Rel_Offset, .Length) - .Position)
      '4907C: 33C0               xor         eax,eax     <- tmp1
      '4907E: C22C00             retn        0002C ;" ,
      '49081: 0401               add         al,001 ;"?  <- current position
      '49083: 6801000000         push        000000001
      '49088: E85E76FCFF         call        0000106EB
      '4908D: 83C404             add         esp,004 ;"
      '49090: FFE0               jmp         eax
      
      
    ' if tmp1=0 -means Nothing found- then bIsPatched = FALSE
      Dim bIsPatched As Boolean: bIsPatched = tmp_AlreadyPatchedOffset
      If bIsPatched Then
       ' Set Position to new Block
         Possible_Patch_Offset.TEXT_Rel_Offset = tmp_AlreadyPatchedOffset
      Else
   
        ' if no more Possible_Patch_Offsets found exit loop
          If Possible_Patch_Offset.TEXT_Rel_Offset = 0 Then err.Raise FoundVersion.UNKNOWN_VERSION
   '-> EXIT: UNKNOWN_VERSION
      
        ' Restore old position
         .Position = Possible_Patch_Offset.TEXT_Rel_Offset + Len(P_GENERAL_ORG)
      End If
     
   ' Note: If File is already Patched the following Data can't be read because they
   '       were overwritten by the patch - so skip this...
     If bIsPatched = False Then
     
        ' Test if the next command is (51 PUSH ECX; AND 68 PUSH_XXX )or 68 PUSH_XXX
         Select Case .int8
            Case &H51
               If .int8 <> &H68 Then Exit Function
               ' Continue Exec...
            Case &H68
               ' Continue Exec...
'               Stop
            Case Else
               Exit Function
         End Select
               
         'Get pointer To ptrToDecryptionDataHeads and Check for integrity(by property let)
         '-> Must point to an address inside the .text section
          ptrToDecryptionDataHeads = .int32
         
      End If
      
     
     '0104C31F  6800C70401          PUSH      0104C700
     '0104C324  6A01                PUSH      01       <- current position
     '0104C326  E82635FCFF          CALL      0100F851
     '0104C32B  FFE0                JMP       EAX
     '0104C32D  90  /8F (SP1)       NOP
     'Seek over call
      .Findstring CALL_XXX, 10: .Move (4)
   
     
     'do additional check ("JMP EAX Filter)
      If .Findstring(JMP_EAX, 8) Then
        
       ' If next command is 90 NOP then it is Set SP2-Flag
         isSP2 = .int8 = &H90
    
         If bIsPatched Then err.Raise FoundVersion.ALREADY_PATCHED
         ScanForCrypBlockStub = True
      End If
      
   End With
End Function


'//////////////////////////////////////////////////////////////////////////////////
'//// SeekForOffset - Search for Pattern and set offset
'//
'// DebugWatch forCurrent Position in file:
'   Hex(input_buffer.Position + Raw_Start_Of_Text_Section)
Public Sub SeekForOffset()
   
   Dim isReadyToPatch As Boolean

  '----------- Initialisation ------------------
   On Error GoTo SeekForOffset_err
   
   frmMain.displn "Å Scanning for byte pattern: "
   
   
 ' change to appDir - necessary to find & load AntiWPA_Crypt.dll later
   Dim tmpDirNow$: tmpDirNow = CurDir
   ChDrive App.Path: ChDir App.Path
   
   MapFileOpen
   
 ' Process-whole-.text-section-Loop    (input_buffer=.text-section)
   With input_buffer
      
      Dim LoopCounter&
      For LoopCounter = 1 To &H7FFFFFFF
         
         If (LoopCounter And &HF) = 0 Then frmMain.disp "."
         If ScanForCrypBlockStub Then
            
            
            'frmmain.Dispv "O"
          
            '---- Read header of first Block ----
            .Position = ptrToDecryptionDataHeads.TEXT_Rel_Offset
            Key1 = .int32
            Key2 = .int32
            
'          Stop
'q: DoEvents: GoTo q
          
            Dim CryDataEnd&: CryDataEnd = .int32
            
            ' There might be an difference between pointer the VA-pointers In DecryptionDataHeads and
            ' the actually VA (Prefered Loadaddress) i.e. if the it's an dll which was rebased
            '(RAW  )    Real_VA     VA_in_DecryptionDataHeads
            '(342e0)    384ee0      434EE0
            ' so address may need to translated...
              Dim RelocTranslation&
            ' CryDataEnd ends at position where DecryptionDataHeads starts so...
              RelocTranslation = ptrToDecryptionDataHeads.VA - CryDataEnd
            
            CryDataEnd = CryDataEnd + RelocTranslation
            
            Dim cryDataStart As New VA_TextSectionPtr
            cryDataStart = .int32 + RelocTranslation
            
            Dim CryDataSize:   CryDataSize = CryDataEnd - cryDataStart
          
            
            Dim CryData As New StringReader
            
           ' ----- Copy Crypted Data into 'CryData' Buffer ----
                           .StorePos
                      .Position = cryDataStart.TEXT_Rel_Offset
            CryData = .FixedString(CryDataSize)
                           .RestorePos
            
         ' VERBOSE output
           frmMain.DispLnV ""
           frmMain.DispLnV "CryDataBlock: " & H32(cryDataStart.TEXT_RAW_Offset) & _
                  "  Size: " & H16(CLng(CryDataSize))
           
                  
                  
          '---- Decrypt first and next Block(s) ----
            
            Dim Counter_Chunks&
          
          ' VERBOSE output
          ' Set 4 chars Space before Text
            frmMain.Fillchars = Space(4)
            
            Dim CryData_ReaderPos&: CryData_ReaderPos = 0
            Dim CryData_WritePos&:  CryData_WritePos = 0
            
            For Counter_Chunks = 0 To &H7FFFFFFF
            
             ' Read data from header
               Dim CryDataBlockSize&: CryDataBlockSize = .int16
             
             ' VERBOSE output
               frmMain.DispLnV "Chunk " & Counter_Chunks & " Size: " & H16(CryDataBlockSize)
               
             ' Check for end of Header
                  If (CryDataBlockSize = 65535) And _
                  (.int16 = 65535) Then Exit For
             ' rewind RelocMarker read
               .Move -2
                
'             ' verify Key1
'             ' Note: the RelocMarker should alway be &HFFFF (65535)
'             '       the RelocMarker is XOR (crypted) with Key1
'               Dim relocKey&
'               relocKey = ((RelocMarker Xor Key1) And 65535)
'               If relocKey <> 65535 Then
'
'                   err.Raise vbObjectError, "SeekForOffset@Decrypt~", _
'                   "Keys at Offset " & Hex(input_buffer.Position + Start_Of_Text_Section.TEXT_RAW_Offset) & " are different! relocKey=" & H16(RelocMarker) & _
'                   " Key1=" & H16(Key1)
'               Else
'
'               End If
             
             
             Dim Lastreloc: Lastreloc = 0
             CryDataBlock = ""
             CryDataBlock.DisableAutoMove = False
             
             ' DeCrypt DataBlock and skip possible reloc addresses (applied only to dll's)
             ' Note: 'CryData' contains all the DataBlocks
             ' CryDataBlock' contains only the current DataBlock
               Do
                ' Get & Decrypt RelocChunk
                  Dim reloc&: reloc = .int16            'get reloc
                  reloc = reloc Xor (Key1 And 65535)    'Decrypt reloc
                  
                  Dim tmp_CryData$
                
                ' Is Reloc a Start/Stop Marker ?
                  If reloc = 65535 Then
                   'Decrypt rest of data length= reloc-.position
                 ' Extract CryData to tmp string for API-Call and store it in CryDataBlock
                   tmp_CryData = CryData.FixedString(CryDataBlockSize - Lastreloc)
                   DeCrypt tmp_CryData, tmp_CryData, Len(tmp_CryData), Key1, Key2
                   CryDataBlock.FixedString = tmp_CryData

                   
                   ' Exit if it's the second Start_Stop_Marker
                     Exit Do
                   ' ...Repeat loop
                  Else
                 ' Decrypt data length= reloc-.position
                 
                 ' Info: Locations with pointer that are managed by the .reloc section
                 '       are excluded from encryption
                 ' The following part decrypts these crypted parts...
                  
                 ' Extract CryData to tmp string for API-Call and store it in CryDataBlock
                   tmp_CryData = CryData.FixedString(reloc - Lastreloc)
                   DeCrypt tmp_CryData, tmp_CryData, Len(tmp_CryData), Key1, Key2
                   CryDataBlock.FixedString = tmp_CryData
                 
                 ' ... and that one simply copy the 32bit unencrypted pointer
                   CryDataBlock.int32 = CryData.int32

                  End If
                  Lastreloc = reloc + 4
               Loop While True
             
               CryData.Position = Swap(CryData.Position, CryData_ReaderPos)   ' "CryData.Position =  ..." is need will not store inCryData.Position
               
               CryDataBlock.DisableAutoMove = True
   
               
              ' = Start of current CryData block
               src = cryDataStart + CryData.Position
               
               CryData.Position = CryData_WritePos
              
              'Depending if mi_Cry_Remove is Checked
              ' set destination to overwriting CryptStub or to source (since we didn't move the decrypted data)
               dest = IIf(frmMain.mi_Cry_Remove.Checked, Possible_Patch_Offset + CryData.Position, src)

              '---- Apply Relocation Fixup on CryDataBlock ---
               ApplyFixUps
               
              '---- Store CryDataBlock in CryData ---
              'Cut off first checksumfunction at the beginning
               CryDataBlock.Position = LengthOfFirstChecksumFunctionCall
              
              'Write Decrypted & Relocated data into CryData
               CryData.FixedString = CryDataBlock.FixedString
              
              'Set new writepos
               CryData_WritePos = CryData.Position - LengthOfExecFunctionCall
               
              'Writing is finished - Set CryData Pos pointer to next Read Pos
               CryData.Position = CryData_ReaderPos
               
              '---- Get new Key from CryDataHead ---
              'Skip Fillbytes (or better Skip Fillwords)
               Do: Loop While .int16 = 0
               .Move -2
                              
              'Read new keys
               Key1 = (.int32 Xor Key1) Or &H10001
               Key2 = (.int32 Xor Key2) Or &H101
               'Calculating the Bitmask
               'these Bit's have to be set: 0,2,4,5
               'Offset    0  1  2  3    4  5  6  7
               'Data     10 20 30 40   50 60 70 80
               'To_dword    40302010      80706050
               '                Key1          Key2
               'Bitmask  01 00 01 00   01 01 00 00    ;Versions 2600.{0,1106,
               'To_dword    00010001      00000101

            Next 'loop for all Crypted DataBlock's
            
          ' Append/move remain data in crydata
            CryData.CopyData CryData_ReaderPos, CryData_WritePos
          ' update CryData_WritePos value
            CryData_WritePos = CryData.Position
               
          ' right after the
            Dim EndOfDecryptionDataHeads As New VA_TextSectionPtr
            EndOfDecryptionDataHeads.TEXT_Rel_Offset = .Position

          ' Overheadsize = CryStubSize + (DecryptionDataHeadSize with free space in CryData)
            Dim Overheadsize&
            Overheadsize = (cryDataStart - Possible_Patch_Offset) + _
                           (EndOfDecryptionDataHeads - cryDataStart - CryData_WritePos)
          
          ' STATISTICS: Wasted bytes by CrypBlockOverhead
            Dim TotalOverhead&
            TotalOverhead = TotalOverhead + Overheadsize
                   
          ' STATISTICS: Wasted bytes by CrypBlockOverhead
            Dim TotalCryptedData&
            TotalCryptedData = TotalCryptedData + (EndOfDecryptionDataHeads - Possible_Patch_Offset)
                   
            
          ' If mi_Cry_Remove is Checked
          ' Fill DecryptionDataHead and Free Space in CryData with CC
            If frmMain.mi_Cry_Remove.Checked Then
               
             ' Write 'CC'-String with overhead size at end of decrypted code into CryData
               CryData.Position = CryData_WritePos
               CryData.FixedString = String(Overheadsize, Chr(&HCC))
               
            End If

             
             
          ' Set 2 chars Space before TextOutput
            frmMain.Fillchars = Space(2)
            
              
            ' For Debug purpose: Write every CryData to a new file on HD
            If frmMain.mi_Debug.Checked Then
               Dim binfile As New FileStream
               binfile.Create Filename.Path & Hex(Possible_Patch_Offset.TEXT_RAW_Offset) & ".bin", True, False
               binfile.FixedString(-1) = CryDataBlock 'CryData
               binfile.CloseFile
            End If


'           --- write deCryData into file ---
'           Seek to .text section or to decrypt stub <- note: this will overwrite the
            
            file.Position = IIf(frmMain.mi_Cry_Remove.Checked, _
                                Possible_Patch_Offset.TEXT_RAW_Offset, _
                                cryDataStart.TEXT_RAW_Offset)
                     
            If frmMain.mi_Debug_StoreInExe.Checked Then
               'write whole .text section to winlogon.exe
                file.FixedString(-1) = CryData
            End If
'
             
             
           MapFileAddEntry Possible_Patch_Offset.TEXT_Rel_Offset, "Cryptblock_" & H16(CryptblockCount)
           CryptblockCount = CryptblockCount + 1
                
                  
             
              
              
          ' Scan inside decrypted data for OFFSET_RECOGNITION_PATTERN
            CryData.Position = 0
            If 0 <> CryData.Findstring(OFFSET_RECOGNITION_PATTERN) Then '"D7 04 07 80 0F" ;[Cmp] 800704D7,JZ...
              
             ' --- Right Offset Found !!! ---
             ' Set Offset as patch offset and exit loop
               frmMain.P_offset = Possible_Patch_Offset.TEXT_RAW_Offset
               frmMain.P_data = P_GENERAL_CRK
               
   ' -> Exit Loop
              'If Debug option enabled continue scan
               If frmMain.bDoFullScan Then
                  
                  frmMain.disp "X"
                  
                  isReadyToPatch = True
               Else
                  frmMain.displn
                  frmMain.displn "PatchOffset: " & H32(frmMain.P_offset) & "  -  Ready to Patch !"
                  err.Raise FoundVersion.READY_TO_PATCH
               End If

            End If
       End If

    Next
 
 End With

    
err.Clear
SeekForOffset_err:
'*Catch err*
  With frmMain
         
      Select Case err
         
   ''     ' General Error -> show Error and contiune loop
   '       Case vbObjectError
   '         frmMain.displn "ERROR: " & err.Description
   '
   '       ' VB only allows resume or Resume Next
   '       ' This don't work - because Errorhandler will not installed
   ''         On Error GoTo SeekForOffset_err
   ''         'Resume Next
   ''         GoTo NextItem
         
         Case FoundVersion.ALREADY_PATCHED
            If .bDoFullScan Then
               .displn "DEBUG: Can't continue processing."
               .displn "DEBUG: You must use an unpatched version as decryption input!"
            End If
         
         Case FoundVersion.READY_TO_PATCH
            err.Clear
            
         Case FoundVersion.UNKNOWN_VERSION
      
      End Select
    '*Finally*
       
       
   ' Kill Checksumfunction
     If frmMain.mi_Cry_Remove.Checked And (ptrToCheckSumFunction <> 0) Then 'And _
        'frmMain.mi_NopOutSCB.Checked = True Then
   
           ' Seek to Checksumfunction & Fix it
             file.Position = ptrToCheckSumFunction.TEXT_RAW_Offset
             frmMain.displn "CheckSumFunctionPatch applied at: 0x" & Hex(file.Position)
             file.FixedString(-1) = DATA_TO_KILL_CHECKSUMFUNTION
   
     End If
     
     
     
   
    'restore currentDir
     ChDrive tmpDirNow
     ChDir tmpDirNow
     
   
     If isReadyToPatch Then ' And (frmmain.mi_Debug.Checked Or frmmain.mi_Debug_StoreInExe.Checked)) = False
      
         If .mi_Cry_Remove.Checked And .mi_NopOutSCB.Checked Then
         
          ' Skip CrypInit at Start
          ' 09 0C        PUSHFD            <- Old EntryPoint
          ' 60           PUSHAD
          ' E8 65AAFCFF  CALL    0000BC2F
          ' 61           POPAD
          ' 9D           POPFD
          ' 6A 70        PUSH    70        <- New EntryPoint
          ' ...          ...
         
          ' Seek to entrypoint & Search for 'POPAD, POPFD' (in a region around 0x10 bytes)
            Dim tmp_skipSize&
            file.Position = PE_Header.EntryPointRVA - PE_Header.BaseofCode + PE_Header.HeaderSize
            tmp_skipSize = InStr(1, file.FixedString(&H10), HexvaluesToString("61 9D"))
          
          ' If found...
            If tmp_skipSize Then
            
              .displn
              .disp "EntryPoint Correction: " & H32(PE_Header.EntryPointRVA) & " -> "
            
            ' ...Write New EntryPoint to File
              PE_Header.EntryPointRVA = PE_Header.EntryPointRVA + tmp_skipSize + 1
              PE_info.WriteHeader
              
              .disp H32(PE_Header.EntryPointRVA)
              .displn
            End If
         End If
        .displn
        .displn "Total crypted data : " & BlockAlign_l(TotalCryptedData \ 1024, 4) & " KB"
        .displn "Total overhead data: " & BlockAlign_l(TotalOverhead \ 1024, 4) & " KB"
        err.Clear
      
      ' Show statisitics
        If .mi_Cry_Remove.Checked Then .displn: .displn "Crypt blocks overhead has been removed."
        If .mi_Debug_StoreInExe.Checked Then .displn: .displn "Uncrypted data has been writing to:": .displn file.Filename
        If .mi_Debug.Checked Then .displn: .displn "*.bin files have been create."
        
     Else
      'Throw any other unhandled error to caller Function
       If err Then err.Raise err, err.Source, err.Description      'Throw to caller Function
     End If
  End With
End Sub

'//////////////////////////////////////////////////////////////////////////////////
'//// ApplyFixUps - Fixes relocated addresses in an uncrypted data chunk
'//
'// SRC         Current VA_Currentaddress
'// Dest                VA_Newaddress
'// RelocMarker Current reloc Start/Stop Marker
Private Function ApplyFixUps()


 ' if 'Remove crypt blocks' is Checked -
 ' Get startoffset of 'Checksum' And 'ExecFunction'
   If frmMain.mi_Cry_Remove.Checked Then GetPointers
   
   With input_buffer
    ' VERBOSE output
    ' Set 6 chars Space before Text
      frmMain.Fillchars = Space(6)
      
    ' Init 'Keys' StringReader
      Dim Keys As New StringReader
      Keys.EOS = False
      Keys.int32 = Key1
      Keys.int32 = Key2
      
      
    ' Process all Relocs
      Dim Counter_Relocs&
      For Counter_Relocs = 0 To &H7FFFFFFF
      
       ' Get Key for current Reloc
         If Keys.EOS Then Keys.EOS = False ' set position to first if end Of String is reached
         Dim key&                          ' if it is SP delete last 4 bit via AND &HFFF0
         key = Keys.int16 And IIf(isSP2, &HFFF0, 65535)
      
       ' Get RelocChunk & exit if Start/Stop Marker
         Do
          ' Get & Decrypt RelocChunk
            Dim reloc&: reloc = .int16            'get reloc
            reloc = reloc Xor (Key1 And 65535)    'Decrypt reloc
         
          ' Is Reloc a Start/Stop Marker ?
            If reloc = 65535 Then
               Exit For
              ' ...Repeat loop
            Else
               If frmMain.mi_Cry_Remove.Checked Then Exit Do  ' ...Exit loop
            End If
         Loop While True

       
       ' Apply Fixup
         With CryDataBlock
            
          ' Move to call xxx to apply fixup
            .Position = reloc
            
          ' Calculate reloc fixUp
            Dim fixUp&
            fixUp = src - dest
          
          ' VERBOSE output
            frmMain.DispV "Reloc: +" & H16(reloc) & " " & H32(.int32) & " +" & H16(fixUp)
          
            .int32 = .int32 + fixUp - key
          
          ' Calculate and validate (via .VA property let) FixedAddress
            Dim FixedAddress As New VA_TextSectionPtr
            FixedAddress.VA = .int32 + dest + .Position + 4
          
          ' VERBOSE output
            frmMain.DispV "=>" & H32(FixedAddress) & " XorKey" & Keys.Position \ 2 & ": " & H16(key) & vbCrLf
          
          ' if 'Remove crypt blocks' is Checked -
          ' Get startoffset of 'Checksum' And 'ExecFunction'
            If frmMain.mi_Cry_Remove.Checked Then NopOutChecksum_Exec_Func FixedAddress
         
         End With
      
      Next
      
      CheckSumFunctionCounter = 0
       
      ' VERBOSE output
      ' Set 4 chars Space before Text
      frmMain.Fillchars = Space(4)
      frmMain.DispLnV ""
   End With
End Function

Private Sub NopOutChecksum_Exec_Func(FixedAddress As VA_TextSectionPtr)
  
  With CryDataBlock
     Dim Nopstart&
     Dim NopEnd&: NopEnd = .Position + 4
     
     
     .bSearchBackward = True
   
   ' Nopout CheckSumFunction FunctionCall
     If FixedAddress = ptrToCheckSumFunction Then
      
        CheckSumFunctionCounter = CheckSumFunctionCounter + 1
      
      ' Test if this is the first Checksumfunction in the Chunk
        If CheckSumFunctionCounter = 1 Then
          
        ' Store pos for Cutoff First CheckSumFunction later
          LengthOfFirstChecksumFunctionCall = NopEnd '-0     <- NopStartPos

        Else
        
        ' set on E8 -Call Opcode
          .Move -1
          GetStartOfCheckSumFunctionCall (8)
        
        ' Do NopOut
          Nopstart = .Position
          If frmMain.mi_NopOutSCB.Checked Then
            .FixedString = Nops(NopEnd - Nopstart)
          Else
            .Move NopEnd - Nopstart
          End If
          
        End If
        
   ' Nopout ExecFunction FunctionCall
     ElseIf FixedAddress = ptrExecFunction Then
      ' Nopout
      
        .Move -2
'            frmMain.DispLnV H16(CryDataBlock.Position) & vbTab & H16(PrevMaxReloc)
  
'            If frmMain.mi_Cry_Remove.Checked Then
        
      ' Test if this is a FUNCTION_CONTINUE Call
      ' $-E     9C              PUSHFD                     $-E      9C              PUSHFD
      ' $-D     60              PUSHAD                     $-D      60              PUSHAD
      ' $-C     68 08000000     PUSH    8                  $-C      56              PUSH    ESI
      '            ^- ReadPosition
      ' $-7     E8 106B7800     CALL    winlogon.0100A366  $-B      57              PUSH    EDI
      ' $-2     FFE0            JMP     EAX                $-A      6A 08           PUSH    8
      ' $ ==>   0000            ADD     [EAX], AL          $-8      E8 9825FCFF     CALL    0100E0A0
      '                                                    $-3      FFE0            JMP     EAX
      '                                                    $-1      90              NOP
      ' ----------------- SP1 -------------------------    ---------------- SP2 ------------------
        Dim ExecMode&
        ExecMode = .int8
        If ExecMode = 0 Then: .Move -3: ExecMode = .int32

        
        
        If ExecMode = &H8 Then
           
           Nopstart = .Findstring(P_GENERAL_ORG)
           .bSearchBackward = False
           NopEnd = .Findstring(HexvaluesToString("FF E0")) + 2 'range 3
           
           LengthOfExecFunctionCall = NopEnd - Nopstart
      
      
      
      ' Test if this is a FUNCTION_BREAK Call
      ' C20800     retn        00008
      ' 0C         pushfd                  <-P_GENERAL_ORG
      ' 60         pushad
      ' 56         push esi
      ' 57         push edi
      ' 6A06       push        006
      '   ^- ReadPosition
      ' E8AB82FDFF call        0FFFD8AD0
      ' C3         retn
      
      '(90         NOP
      ' 5E         pop         esi
      ' C9         leave
      ' C3         retn)
      
      
        ElseIf ExecMode = &H6 Then
'           Stop
'           Do: DoEvents: Loop While True
           Nopstart = .Findstring(P_GENERAL_ORG)
           
'                      .bSearchBackward = False
 '          .Position = .Findstring(HexvaluesToString("E8")) 'range 5
  '         .Move 5
   '        NopEnd = .Position
           
    '       LengthOfExecFunctionCall = .Findstring(HexvaluesToString("C3")) - Nopstart

           
           LengthOfExecFunctionCall = (NopEnd - Nopstart) + 1
        
      
      
      ' Test if this is a FUNCTION_RET Call
      ' 60         pushad                  <- 60 PUSHAD!
      ' 56         push esi
      ' 57         push edi
      ' 6A0A       push        00A
      '   ^- ReadPosition
      ' E8A782FDFF call        0FFFD82B0   <-NopOut
        ElseIf ExecMode = &HA Then
           Nopstart = .FindByte(&H60) '60 PUSHAD
           LengthOfExecFunctionCall = 0
        Else
           err.Raise vbObjectError, "ApplyFixUps", "Unexpected EXEC_DECRYPT mode value: " & Hex(.int8) & " Or " & H16(.int32)
        End If

        
        .Position = Nopstart
        .FixedString = Nops(NopEnd - Nopstart)
     End If
  End With
End Sub
'///////////////////////////////////////////////////////////////
'// GetPointers - Get StartOffset of ptrToCheckSumFunction
'//               and ptrExecFunction
Private Sub GetPointers()

   With input_buffer
      .StorePos
      ' VERBOSE output
      ' Set 6 chars Space before Text
      frmMain.Fillchars = Space(6)
      
      Dim Keys As New StringReader
      Keys.EOS = False
      Keys.int32 = Key1
      Keys.int32 = Key2
      
      Dim Counter_Relocs&
      Dim MinReloc&: MinReloc = 65535
      Dim MaxReloc&: MaxReloc = 0
      '               Dim PrevMaxReloc&: PrevMaxReloc = 0
      
      For Counter_Relocs = 0 To &H7FFFFFFF
      
       ' Get Key for Reloc
         If Keys.EOS Then Keys.EOS = False ' set position to first if end Of String is reached
         Dim key&                          ' if it is SP delete last 4 bit via AND &HFFF0
         key = Keys.int16 And IIf(isSP2, &HFFF0, 65535)
      
       ' Get RelocChunk & exit if Start/Stop Marker
         Do
          ' Get & Decrypt RelocChunk
            Dim reloc&: reloc = .int16            'get reloc
            reloc = reloc Xor (Key1 And 65535)    'Decrypt reloc
            
          ' Is Reloc a Start/Stop Marker ?
            If reloc = 65535 Then
             ' Exit if it's the second Start_Stop_Marker
               Exit For
             ' ...Repeat loop
            Else
             ' ...Exit loop
               Exit Do
            End If
         Loop While True
         
       ' Apply Fixup
         With CryDataBlock

          ' Move to call xxx to apply fixup
            .Position = reloc

'          ' VERBOSE output
'            frmMain.DispV "Reloc: +" & H16(reloc) & " " & H32(.int32)

          ' Calculate and validate FixUpValue
            Dim FixedAddress As New VA_TextSectionPtr

            FixedAddress.VA = .int32 - key + src + .Position + 4


'          ' VERBOSE output
'            frmMain.DispV "=>" & H32(FixedAddress) & " XorKey" & Keys.Position \ 2 & ": " & H16(key) & vbCrLf

          ' Get ptrExecFunction
          ' Assumption:  Call with highest relocaddress is calls to ExecFunction
            If reloc > MaxReloc Then
               MaxReloc = reloc
               ptrExecFunction = FixedAddress
'               frmMain.DispLnV "set ExecFunction = Call " & H32(ptrExecFunction.VA)
            End If

          ' Get ptrToCheckSumFunction
          ' Assumption:  Call with lowest relocaddress is calls to Checksumfunction
            If reloc < MinReloc Then
               MinReloc = reloc
               ptrToCheckSumFunction = FixedAddress
'               frmMain.DispLnV "set CheckSumFunction = Call " & H32(ptrToCheckSumFunction.VA)
            End If

         End With
         
         Next
      .RestorePos
      End With
                          
    ' Set new Destination for Reloc
      dest = dest - MinReloc - 4
End Sub



Private Sub Class_Initialize()
 ' Initialise Vars
   JMP_EAX = HexvaluesToString("FF E0")
   PUSH_XXX = Chr(&H68)
   CALL_XXX = Chr(&HE8)
   NOP = Chr(&H90)
   OFFSET_RECOGNITION_PATTERN = HexvaluesToString("D7 4 7 80 F")
   'This is [Cmp] 800704D7,JZ...  ; 800704D7 seems to be some errorcode or what else but
   'I hope this will stay the same in futur version ! ;)
   DATA_TO_KILL_CHECKSUMFUNTION = HexvaluesToString("C2 8 0")
   
 ' Init Start_Of_Text_Section with 0 (= start of textsection)
   Start_Of_Text_Section.TEXT_Rel_Offset = 0
   
'  Reopen file
   file.Filename = Filename.Filename
   file.CloseFile
   PE_info.Create
   
 ' Read whole .text section into buffer
   file.Position = Start_Of_Text_Section.TEXT_RAW_Offset     'Seek to .text section
   input_buffer = file.FixedString(Start_Of_Text_Section.TEXT_SECTION_RawSize)

End Sub
'///////////////////////////////////////////////////////////////////////
'// GetStartOfCheckSumFunctionCall  - Gets offset where the first push
'//                                   for the CheckSumFunction Call is
'//
'// Note: this is a recursiv Function
'// For proper function CryDataBlock.position must be at the Call (E8)
Private Function GetStartOfCheckSumFunctionCall(StackCounter%, Optional StackCounterDeincrease& = 0) As Long

  'Set Stackounter value
   StackCounter = StackCounter - StackCounterDeincrease
   
    ' If StackCounter is 0 all Pushes are processed
      If StackCounter = 0 Then
      '   GetStartOfCheckSumFunctionCall = .Position
         Exit Function
      ElseIf StackCounter < 0 Then
         Stop
      End If
   
   
   With CryDataBlock
    '
      Dim tmpByte_1&, _
          tmpByte_2&, _
          tmpByte_3&, _
          tmpByte_4&, _
          tmpByte_5&

      .Move -5
'      Debug.Assert .Position <> 0
      .DisableAutoMove = False
      tmpByte_5 = .int8
      tmpByte_4 = .int8
      tmpByte_3 = .int8
      tmpByte_2 = .int8
      tmpByte_1 = .int8
      .DisableAutoMove = True

      
    ' Detect all possible pushes

'      Dim isStackCounterBE_2 As Boolean: isStackCounterBE_2 = StackCounter >= 2
      Dim PUSH_R32 As Boolean
      Dim PUSH_V32_08 As Boolean
      Dim PUSH_V32 As Boolean
      
'      Dim isStackCounterBE_4 As Boolean: isStackCounterBE_4 = StackCounter >= 4
      Dim PUSH_R16 As Boolean
      Dim PUSH_V16_8 As Boolean
      Dim PUSH_V16_16 As Boolean


      If StackCounter >= 4 Then
         PUSH_R32 = ((tmpByte_1 And &HF8) = &H50)
         PUSH_V32_08 = (tmpByte_2 = &H6A) And Not ((tmpByte_4 = &H68) And (tmpByte_5 = &H66))
         PUSH_V32 = (tmpByte_5 = &H68)
      End If
      
      If StackCounter >= 2 Then
         PUSH_R16 = (tmpByte_2 = &H66) And ((tmpByte_1 And &HF8) = &H50) And Not ((tmpByte_4 = &H68) And (tmpByte_5 = &H66))
         PUSH_V16_8 = (tmpByte_3 = &H66) And (tmpByte_2 = &H6A)
         PUSH_V16_16 = (tmpByte_4 = &H66) And (tmpByte_3 = &H68)
      End If


' 68 66506653     PUSH    53665066  ' 66:50           PUSH    AX
                                    ' 66:53           PUSH    BX
' 66:53           PUSH    BX
' 66:51           PUSH    CX
' E8 454DFCFF     CALL    0100E034


' 66:53           PUSH    BX
' 68 94680000     PUSH    6894
' 66:53           PUSH    BX
' E8 BE43FCFF     CALL    0100E034


'  66:68 406A      PUSH    6A40
'  52              PUSH    EDX
'  66:51           PUSH    CX
'  E8 D782FCFF     CALL    0100FE92



      If PUSH_R32 And Not (PUSH_V32_08 Or PUSH_R16 Or PUSH_V16_8 Or PUSH_V16_16) Then
          ' 50                 PUSH EAX
          Debug.Assert (tmpByte_1 And &HFC) = &H50
           .Move -1
           GetStartOfCheckSumFunctionCall StackCounter, 4
   
       ElseIf PUSH_V32_08 And Not (PUSH_V16_8 Or PUSH_V16_16) Then
           '         6A 02     PUSH    2
           .Move -2
           GetStartOfCheckSumFunctionCall StackCounter, 4
            
       ElseIf PUSH_R16 And Not (PUSH_V16_8 Or PUSH_V16_16) Then
           '         66 53     PUSH    BX
          Debug.Assert (tmpByte_1 And &HFC) = &H50
           .Move -2
           GetStartOfCheckSumFunctionCall StackCounter, 2
       
       ElseIf PUSH_V16_8 And Not (PUSH_V16_16) Then
           '      66 6A C8     PUSH    C8
           .Move -3
           GetStartOfCheckSumFunctionCall StackCounter, 2
       
       ElseIf PUSH_V16_16 Then
           '   66 68 C8:4D     PUSH  4DC8
           .Move -4
           GetStartOfCheckSumFunctionCall StackCounter, 2
       
       ElseIf PUSH_V32 Then
           '68 E8:A2:00:00     PUSH 0A2E8
           .Move -5
'           Debug.Assert tmpByte_1 = 0 And (tmpByte_2 = 0 Or tmpByte_2 = 1)
           GetStartOfCheckSumFunctionCall StackCounter, 4
       Else
          err.Raise vbObjectError, "GetStartOfCheckSumFunctionCall", "Detection StartOfCheckSumFunctionCall failure. Maybe there is an other command between the Push's"
       End If
   
   End With
   
   
End Function


Private Function Nops(size) As String
   Nops = String(size, NOP)
End Function

Private Sub Class_Terminate()
  On Error Resume Next
  MapFileClose
End Sub
