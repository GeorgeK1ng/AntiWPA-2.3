VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6495
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   7800
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer_OleDrag 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton Cmd_RestoreBackup 
      Caption         =   "Restore Backup"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmd_OffsetLocator 
      Appearance      =   0  'Flat
      Caption         =   "Offset Locator"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   5880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmd_cancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Quit"
      Height          =   495
      Left            =   4335
      TabIndex        =   3
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton Cmd_apply 
      Appearance      =   0  'Flat
      Default         =   -1  'True
      Height          =   495
      Left            =   6030
      TabIndex        =   2
      ToolTipText     =   "Click right to choose and left to apply"
      Top             =   5880
      Width           =   1695
   End
   Begin VB.TextBox Txt_Console 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A0D0A0&
      Height          =   5655
      Left            =   120
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "frmMain.frx":030A
      Top             =   120
      Width           =   7575
   End
   Begin VB.Label lbl_Email 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "crackware2k@freenet.de"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   120
      MouseIcon       =   "frmMain.frx":062C
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   6165
      Width           =   1965
   End
   Begin VB.Menu static 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mi_open 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu static 
      Caption         =   "&Options"
      Index           =   1
      Begin VB.Menu mi_OOBE 
         Caption         =   "Apply &OOBE Fix"
         Checked         =   -1  'True
      End
      Begin VB.Menu mi_wpa 
         Caption         =   "Apply &WPA Fix"
         Checked         =   -1  'True
      End
      Begin VB.Menu mi_MsOOBE_Overwrite 
         Caption         =   "Replace Msoobe.exe with AntiWPA"
      End
      Begin VB.Menu Seperator 
         Caption         =   "-"
      End
      Begin VB.Menu mi_NopOutSCB 
         Caption         =   "Remove selfcheck blocks"
      End
      Begin VB.Menu mi_Cry_Remove 
         Caption         =   "Remove crypt blocks"
      End
      Begin VB.Menu Seperator3 
         Caption         =   "-"
      End
      Begin VB.Menu mi_Debug 
         Caption         =   "Debug: Save decrypted code to *.bin"
      End
      Begin VB.Menu mi_Debug_StoreInExe 
         Caption         =   "Debug: Save decrypted code to exe"
      End
      Begin VB.Menu mi_MAP 
         Caption         =   "Debug: &Create MAP-File"
      End
      Begin VB.Menu Seperator4 
         Caption         =   "-"
      End
      Begin VB.Menu mi_verbose 
         Caption         =   "Debug: Verbose Mode"
      End
   End
   Begin VB.Menu static 
      Caption         =   "&Help"
      Index           =   3
      Begin VB.Menu mi_update 
         Caption         =   "Check for Update"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////
'// Frm_Main - Main Module
'//
'// Does Contains GUI stuff and PatchLogic
'//
Option Explicit

Private Const MAX_CHARS = 59  ' For Textbox Output (used in displn)
Public Fillchars$  ' For Textbox Output usually "  "


'Private Const SB_HORZ As Long = 0
'Private Const SB_VERT As Long = 1
'Private Declare Function SetScrollRange Lib "user32.dll" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
'
'Private Const SB_BOTH As Long = 3
'Private Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long

Private Const WM_CHAR = &H102
Private Const WM_PASTE = &H302
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type
Private Const OFN_READONLY As Long = &H1
Private Const OFN_OVERWRITEPROMPT            As Long = &H2
Private Const OFN_HIDEREADONLY               As Long = &H4
Private Const OFN_NOCHANGEDIR                As Long = &H8
Private Const OFN_SHOWHELP                   As Long = &H10
Private Const OFN_ENABLEHOOK                 As Long = &H20
Private Const OFN_ENABLETEMPLATE             As Long = &H40
Private Const OFN_ENABLETEMPLATEHANDLE       As Long = &H80
Private Const OFN_NOVALIDATE                 As Long = &H100
Private Const OFN_ALLOWMULTISELECT           As Long = &H200
Private Const OFN_EXTENSIONDIFFERENT         As Long = &H400
Private Const OFN_PATHMUSTEXIST              As Long = &H800
Private Const OFN_FILEMUSTEXIST              As Long = &H1000
Private Const OFN_CREATEPROMPT               As Long = &H2000
Private Const OFN_SHAREAWARE                 As Long = &H4000
Private Const OFN_NOREADONLYRETURN           As Long = &H8000
Private Const OFN_NOTESTFILECREATE           As Long = &H10000
Private Const OFN_NONETWORKBUTTON            As Long = &H20000
Private Const OFN_NOLONGNAMES                As Long = &H40000
Private Const OFN_EXPLORER                   As Long = &H80000
Private Const OFN_NODEREFERENCELINKS         As Long = &H100000
Private Const OFN_LONGNAMES                  As Long = &H200000
Private Const OFN_ENABLEINCLUDENOTIFY        As Long = &H400000
'Private Const OFN_ENABLESIZING               As Long = &H800000
Private Const OFN_DONTADDTORECENT            As Long = &H2000000
'Private Const OFN_FORCESHOWHIDDEN            As Long = &H10000000
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long


'Private Const INF_STYLE_CACHE_DISABLE As Long = &H20
'Private Const INF_STYLE_NONE As Long = &H0
'Private Const INF_STYLE_CACHE_ENABLE As Long = &H10
'Private Const INF_STYLE_OLDNT As Long = &H1
'Private Const INF_STYLE_WIN4 As Long = &H2
'
'Private Declare Function SetupOpenInfFile Lib "setupapi.dll" Alias "SetupOpenInfFileA" (ByVal FileName As String, ByVal fClass As String, ByVal fStyle As Long, ByRef ErrorLine As Long) As Long
'Private Declare Function SetupInstallFromInfSection Lib "setupapi.dll" Alias "SetupInstallFromInfSectionA" (ByVal Owner As Long, ByRef fHandle As Long, ByVal SectionName As String, ByVal Flags As Long, ByVal RelativeKeyRoot As Long, ByVal SourceRootPath As String, ByVal CopyFlags As Long, ByRef MsgHandler As Long, ByVal Context As Long, ByRef DeviceInfoSet As Long, ByRef DeviceInfoData As Long) As Long
'Private Declare Function SetupCloseInfFile Lib "setupapi.dll" (ByReffHandle As Long) As Long

Public P_data$
Public P_offset&

Public SCBFillOption As SCBFillOptions

Private mvarIsAlreadyPatched As Boolean 'lokale Kopie
Private MSOOBE_Filename As New ClsFilename
Private MSOOBE_Backup As New ClsFilename
Private SFC_Blocker As ClsSFC_Blocker

Public Property Let IsAlreadyPatched(ByVal vData As Boolean)

    Cmd_apply.Caption = IIf(vData, "", "Apply / ") & "Browse"
    
    mvarIsAlreadyPatched = vData
End Property
Public Property Get IsAlreadyPatched() As Boolean
    IsAlreadyPatched = mvarIsAlreadyPatched
End Property



Public Sub CheckForUpdate()
     ShellExecute 0, "open", _
               "http://antiwpa11.tk", _
               "", "", 0
  ' & App.Major & "." & App.Minor & App.Revision,

End Sub


Private Sub output(Text$)
   
'Dim strEnd&, strStart&
'strStart = 1
'strEnd = Len(Text)
'
'Do
'
'   strEnd = InStr(strStart, Text, vbCrLf) + 2
'   If strEnd = 2 Then
'      List1.List(List1.ListCount - 1) = List1.List(List1.ListCount - 1) & _
'        Mid(Text, strStart, Len(Text) - strStart + 1)
'      Exit Do
'   Else
'      List1.AddItem Mid(Text, strStart, strEnd - strStart - 2)
'      List1.Selected(List1.ListCount - 1) = True
'
'      If List1.ListCount = 27 Then
'         SetScrollRange List1.hwnd, SB_VERT, 0, 0, 1
'      End If
'   End If
'
'   strStart = strEnd
'   DoEvents
'Loop While True
'Exit Sub



  'Crop Textboxcontent if it's near the Textboxlimit
   If Len(Txt_Console) > 65000 Then _
      Txt_Console = "<Snip>" & Right(Txt_Console, 50000)
   
   
   Txt_Console.SelStart = Len(Txt_Console)
'
'''1. Version
'   Dim oldtext
'   oldtext = Clipboard.GetText
'
'      Clipboard.SetText ByVal Text
'      PostMessage Txt_Console.hwnd, WM_PASTE, 0, 0
'      DoEvents
'      Clipboard.SetText (oldtext)


'2. Version

   Dim i
   For i = 1 To Len(Text)
      Dim char$
      char = Mid(Text, i, 1)
      If char <> vbLf Then
         PostMessage Txt_Console.hwnd, &H102, _
         Asc(char), 0
      End If
   Next
   DoEvents

'4. Version (this is too slow/inefficent)

'   Txt_Console = Txt_Console & Text
'   Txt_Console.SelStart = Len(Txt_Console)

End Sub


'Show text in console Textbox and add newline
Public Sub displn(ParamArray Text())
'   On Error Resume Next
   
  'Join array to a string
   Dim tmp As New StringReader
   tmp = Join(Text, vbTab)
   

   If Len(tmp) = 0 Then
      output vbCrLf
   Else
      
     'Output next lines
      Dim i&, steps&
      steps = MAX_CHARS ' + 2
   
      i = 1
      Do While i <= Len(tmp)
        'Output line
         If Left(tmp, i) <> "Å" Then
            steps = MAX_CHARS - Len(Fillchars)
            output Fillchars & Mid(tmp, i, steps) & vbCrLf
         Else
            steps = MAX_CHARS ' + 2
            output vbCrLf & Mid(tmp, i, steps) & vbCrLf
         End If
         i = i + steps
      Loop
   End If

End Sub

'Show text in console Textbox
Public Sub disp(tmp$) 'ParamArray Text()
'   Dim tmp$
'   tmp = Join(Text, vbTab)
   
   Dim LengthOfLine
'   LengthOfLine = Len(List1.List(List1.ListCount - 1))
   LengthOfLine = Len(Txt_Console) - InStrRev(Txt_Console, vbCrLf)
   
   If LengthOfLine > MAX_CHARS + 2 Then
      output vbCrLf & Fillchars & tmp
   ElseIf LengthOfLine = 1 Then
      If Left(tmp, 1) <> "Å" Then
         output Fillchars & tmp
      Else
         output tmp
      End If
   Else
     output tmp
   End If
         

End Sub
Private Function CreateStringPatter(ByRef Patter$, ByVal Repeat&)
  CreateStringPatter = Space(Len(Patter) * Repeat)
  Dim i
  For i = 0 To Repeat - 1
      Mid(CreateStringPatter, i * Len(Patter) + 1) = Patter
  Next
End Function

' Creates a Asmfillpattern with 'long' ASM codes to
' makes the fillcode th take as less as possible lines in the disassembling
Private Function SCB_FillCode(size As Long) As String
 
 'Check it a short jmp will work
  Debug.Assert size <= &H7F + 2
  
 'Calc how often fillpatter shall be repeated
  Dim Num_10erBlocks%, Num_5erBlocks%, Num_1erBlocks%
  Num_1erBlocks = (size - 2)
  If SCBFillOption = SCBFillOptions.FILL_WITH_LONG_ASM Then
     Num_10erBlocks = Num_1erBlocks \ 10: Num_1erBlocks = Num_1erBlocks Mod 10
     Num_5erBlocks = Num_1erBlocks \ 5
     Num_1erBlocks = Num_1erBlocks Mod 5
  End If
   
  
 'Create & return fillpattern
  SCB_FillCode = Chr(&HEB) & ByteToBin(size - 2) & _
                 CreateStringPatter(HexvaluesToString("C7 05 FF FF FF FF FF FF FF FF"), Num_10erBlocks) & _
                 CreateStringPatter(HexvaluesToString("A3 FF FF FF FF"), Num_5erBlocks) & _
                 String(Num_1erBlocks, Chr(&H90))
  
End Function

Private Function ByteToBin(ByRef value&) As String
   Dim tmp$
   tmp = Space(1)
   MemCopyLngToStr tmp, value, 1
   ByteToBin = tmp
End Function


'//////////////////////////////////////////////////////////////////////////////////
'// RemoveSCB - Apply the Patch and Remove the Security Check Block in winlogon.exe
'//
'// Disable Selfcheckblock in Winlogon.exe to make it possible to patch it.
'// There are about 400 selfcheckblock all over the program (.text section)
'// If in one selfcheckblock is noticed that something in Winlogon.exe has changed
'// windows will crash/reboot
'//
'//
'// Example for Winlogon SP1 Selfcheckblock
'//
'01031F91: FF2514330701                 jmp         d,[01073314] ;points to
'01031F97: 9C                           pushfd         ;<-  01031F97
'01031F98: 60                           pushad
'01031F99: FF742414                     push        d,[esp][14]
'01031F9D: FF742410                     push        d,[esp][10]
'01031FA1: FF74240C                     push        d,[esp][0C]
'01031FA5: FF742408                     push        d,[esp][08]
'01031FA9: 680E000000                   push        00000000E
'01031FAE: 685EE20601                   push        00106E25E
'01031FB3: 6800000000                   push        000000000
'01031FB8: 6860000000                   push        0000000600
'01031FBD: FF35C0A70001                 push        d,[0100A7C0]
'01031FC3: 68D4370701                   push        0010737D4
'01031FC8: 68E4190001                   push        0010019E4
'01031FCD: 680AEC0601                   push        00106EC0A
'01031FD2: E8A79BFDFF                   call       .00100BB7E  ;Set Bpx on 01031FDF here!
'   Skipped (Not executed)
'   01031FD7: 81EC0C040000                 sub         esp,00000040C
'   01031FDD: 61                           popad
'   01031FDE: 9D                           popfd
'
'01031FDF: 83C430                       add         esp,030
'01031FE2: 8D1DF01F0301                 lea         ebx,[01031FF0]  ;=Normal Programm
'01031FE8: 891D14330701                 mov         [01073314],ebx  ;jmp [01073314] will skip the check
'01031FEE: 61                           popad
'01031FEF: 9D                           popfd
'
'Normal Programm
'01031FF0: 57                           push        edi
'01031FF1: 56                           push        esi
'01031FF2: E8D8830000                   call       .00103A3CF
'01031FF7: 83FF02                       cmp         edi,002 ;
'01031FFA: 746C                         je         .001032068
'...
'//////////////////////////////////////////////////////
'// Read PE-Header to get all information(BaseofCode, ImageBase...) to calculate the
'// File offset from a Virtual Address(=VA)
'//
'// Seek to .data section in exe file and search for the Security Check Block (SCB) table
'// All VA-Address in SCB table will be overwritten with 'good' ones.
'// Result when executing winlogon - all SCB are skipped.
'//
Private Sub RemoveSCB(Filename$)
  
  '--- Disable winlogon Anti-Crack-Security ---
   On Error GoTo RemoveSCB_err
   
   displn "Å Disabling winlogon Anti-Crack-Security: "
   With file
     '--- Open Files ---
      .Create Filename
      file_readonly.Create Filename, , , True
      
'     'Don't forget to reinitalise in case you changed the file
'      file.create (filename)
'      PE_info.create


     '--- seek to .data -section---
     'note: in the .data-section is a table with pointers to all
     '      SelfCheckBlocks(=SCB's) in the .text section
      .Position = PE_Header.Sections(DATA_SECTION).PointertoRawData

'     Calculate VirtualAddres_to_RAW_Offset value (for later use)
'      Dim VA_to_RAW_Offset&
'      VA_to_RAW_Offset = PE_Header.BaseofCode + PE_Header.ImageBase - PE_Header.HeaderSize

      
      Dim counter&, wasted&

'      If mi_NopOutSCB.Checked Then
'         Dim Checkproc&, CheckprocMin&, CheckprocMax&
'         CheckprocMin = &H7FFFFFFF
'      End If


      Dim bad_offset As New VA_TextSectionPtr
      bad_offset.RaiseErrorIfInvalid = False
      
      For counter = 0 To &H7FFFFFF
        
        'read Long Value
         bad_offset = .longValue
'         On Error GoTo RemoveSCB_err
        
        'do Loop while 'bad_offset' is a valid VA-Adress
        'valid means: Offset must be in the range  of the image and
        '             Offset maybe have the value 0
         If ((RangeCheck(bad_offset, PE_Header.ImageBase + PE_Header.ImageSize, PE_Header.ImageBase)) Or _
                 (bad_offset = 0)) = False Then Exit For
                 
         If bad_offset = 0 Then counter = 0

                 
        'seek to location in .text-section 'bad_offset' points to
         file_readonly.Position = bad_offset.TEXT_RAW_Offset
        
        '--- find SCB-Table ---
        'Filter valid SCB (Opcode: 9C pushfd;  60 pushad)
         If file_readonly.intValue = &H609C Then
            
           '--- located good offset ---
           '61 popad; 9D popfd
            file_readonly.FindBytes &H61, &H9D
            
'            If mi_NopOutSCB.Checked Then
'             ' - Collect data for ripout all checkfunctions later -
'             ' Assumption: there no other (important and useful) functions between checkfunctions
'             ' - if there any -between CheckprocMin and CheckprocMax- they will delete
'               file_readonly.Position = file_readonly.Position - 16
'               file_readonly.FindBytes &HE8
'
'             ' Get checkproc statistic data
'               Checkproc = file_readonly.Position - Not (file_readonly.longValue) + 3
'               CheckprocMin = Min(CheckprocMin, Checkproc)
'               CheckprocMax = Max(CheckprocMax, Checkproc)
'               Debug.Assert bad_offset <> &H101F1D7
'               displn H32(Checkproc), H32(bad_offset)
'            End If
            
            
            
           'lea         ebx,xxx
            file_readonly.FindBytes &H8D, &H1D
               
           'file_readonly.Position = file_readonly.Position + 5
            Dim good_offset&
            good_offset = file_readonly.longValue


          ' Statistics: Calculate SCB_size
            Dim SCB_size&
            SCB_size = good_offset - bad_offset ' + VA_to_RAW_Offset)
            
          ' --- Nop Out ---
            If mi_NopOutSCB.Checked Then
              'store old filepos
               Dim tmp As Long: tmp = file.Position
                
              'Nop out call to Checkproc
             ' Rewind to jmp [xxx] (6 Bytes)
               file.Position = bad_offset.TEXT_RAW_Offset - 6

               If SCBFillOption = NOP_OUT_COMPLETE Then
                  file.FixedString(-1) = String(SCB_size + 6, Chr(&H90))
               Else
                  file.FixedString(-1) = SCB_FillCode(SCB_size + 6)
               End If
                
              'restore old filepos
               file.Position = tmp
            End If
            
          ' --- Display progress ---
            If mi_verbose.Checked Then
            
             ' Statistics: Calculate max_SCB_size
               Dim max_SCB_size&
               max_SCB_size = Max(max_SCB_size, SCB_size)
               
             ' Statistics: Calculate wasted bytes through Security Check Blocks
               wasted = wasted + SCB_size + 5
            
             ' Output

               If counter = 1 Then
                  DispLnV "SCB" & vbTab & "Offset" & vbTab & "Bad_VA" & vbTab & "Good_VA" & vbTab & "SCB_size"
                  DispLnV String(38, "-")
               End If
               DispLnV counter & vbTab & Hex(.Position - 4) & vbTab & Hex(bad_offset) & vbTab & Hex(good_offset) & vbTab & SCB_size
            Else
               If (counter And &HF) = 0 Then disp "."
            End If
            
'              'SecurityCheck - make sure that 'good_offset' is valid
'              '61 popad; 9D popfd
'               file_readonly.FindBytes  &H61, &H9D
'
'               If file_readonly.Position <> (good_offset - VA_to_RAW_Offset) Then
'                  Err.Raise vbObjectError + 10, , "Can't find 'good_offset'!"
'               End If

           '--- write 'good_offset' ---
            file.Move -4
            file.longValue = good_offset
   
     
         End If   'of Filter scb-blocks
   
      Next
      
'      If mi_NopOutSCB.Checked Then
'         displn Hex(CheckprocMin), Hex(CheckprocMax)
'         file.Position = CheckprocMin
'         file.FixedString(-1) = String(CheckprocMax - CheckprocMin, Chr(&HCC))
'      End If
      
' Show statistics
'      displn "Chunks: " & Chunks, "max_SCB_size: " & max_SCB_size, "wasted bytes: " & wasted
      
      displn
      If mi_NopOutSCB.Checked Then
         displn "Additionally self check blocks have been overwritten using option: " & _
         Switch(SCBFillOption = SCBFillOptions.NOP_OUT_COMPLETE, "NopOut_Complete", _
                SCBFillOption = SCBFillOptions.NOP_OUT_NORMAL, "NopOut_Normal", _
                SCBFillOption = SCBFillOptions.FILL_WITH_LONG_ASM, "Long_ASM_Fill")
      End If

      DispLnV "Approx. wasted bytes by checkblocks: " & wasted


    'Apply patch
     If P_offset And mi_wpa.Checked Then
         displn "Å Applying Anti-WPA Patch at offset:  " & Hex(P_offset)
         .Position = P_offset
         .FixedString(-1) = P_data
         
'          displn "Setting File patched Marker at offset:  " & Hex(P_MARKER_OFFSET)
'         .Position = P_MARKER_OFFSET
'         .char = P_MARK_AS_PATCHED_VALUE
         
      Else
         displn "Å Skipping Anti-WPA Patch..."
      End If
      
      
      .CloseFile
      file_readonly.CloseFile
      
    ' Update PE_Checksum
      Dim RetVal&
      disp "PE_Checksum update"
      RetVal = PE_info.UpdateChecksum
      If RetVal = 0 Then disp "d!" Else disp " failed. Error: " & RetVal
      displn
         
   End With
  
err.Clear
RemoveSCB_err:
Select Case err
   Case 0:
   Case Else
    ' Store errordata
      Dim num&, src$, descr$
      num = err.Number
      src$ = err.Source
      descr = err.Description
      
      displn "CRITICAL ERROR!"
   
End Select
   
  'trigger SCFMessageBox if file was modified
   On Error Resume Next
   If P_offset Then SFC_Blocker.WaitForSFCMessagebox = True

  'raise error
   If num <> 0 Then err.Raise num, src, descr
   
End Sub

Private Sub createBackup()
   With Filename

      displn "Å Preparing patch..."
     
     'Prepare FileNames
      Dim FileExe$, FileBak$
      FileExe = .Name & .Ext
      FileBak = .Name & ".bak"
      
     'Set Workingdir
      On Error Resume Next
      ChDrive .Path
      ChDir .Path
      On Error GoTo 0
      
     'Delete dllcache\winlogon.exe
      myFileDelete "dllcache\" & FileExe
      
     'Delete old winlogon.bak
      myFileDelete FileBak
     
      
     'Better we close the file before renaming...
     'in short what later will cause problems:
     'the openfilehandle which refered to winlogon.exe will
     'refered to winlogon.bak after renaming, but the File.-objekt
     'still thinks the openfilehandle belongs to winlogon.exe
      file.CloseFile
      
     'Rename winlogon.exe to winlogon.bak
      myFileRename FileExe, FileBak
     
     'copy winlogon.bak to winlogon.exe
      myFileCopy FileBak, FileExe
      
     'Remove readonly attrib & Test if .FileName exists => raise.Err 53
      SetAttr .Filename, vbNormal
   
   End With
End Sub
Private Sub OOBE_Fix()
    ' OOBE NeedActivation Registryfix (LicDll.dll!GetExpirationInfo)
   displn "Å Applying OOBE NeedActivation Registryfix..."
   Dim reg As New Registry
   With reg
      .Create HKEY_LOCAL_MACHINE, REG_PATH_WPAEVENTS, True, REG_DATA_OOBETimer
      .RegValueDataTypeForCreateNew = REG_BINARY
      .Regdata = HexvaluesToString(REG_VALUE_OOBETimer)
   End With
   displn REG_DATA_OOBETimer & " = " & REG_VALUE_OOBETimer
   
   displn "Å Removing activation shortcut..."
   displn "Executing: 'syssetup.inf!DEL_OOBE_ACTIVATE'..."
   ShellExecute Me.hwnd, "open", "rundll32", "setupapi,InstallHinfSection DEL_OOBE_ACTIVATE 132 syssetup.inf", 0, vbHide
'      Dim hInf&, RetVal&
'      hInf = SetupOpenInfFile("syssetup.inf", vbNullString, INF_STYLE_WIN4, 0)
'      RetVal = SetupInstallFromInfSection(Me.hwnd, hInf, "DEL_OOBE_ACTIVATE", &H100, 0, 0, 0, 0, 0, 0, 0)
'      RetVal = SetupCloseInfFile(hInf)


End Sub
Private Sub MsOOBE_Overwrite()
   displn "Å Replacing MSOOBE.exe ..."
'   myFileDelete MSOOBE_Backup.Filename
   myFileRename MSOOBE_PATH & MSOOBE_EXE, MSOOBE_Backup.Filename
   myFileCopy App.Path & "\" & App.EXEName & ".exe", MSOOBE_PATH & MSOOBE_EXE
   myFileCopy App.Path & "\" & CRYPT_DLL, MSOOBE_PATH & CRYPT_DLL
End Sub




Private Sub Cmd_apply_Click()
   If IsAlreadyPatched Then BrowseForFile: Exit Sub
    On Error Resume Next
   
   'Disable all buttons during work
    ButtonsEnable False
   
    On Error GoTo Cmd_apply_Click_err
   
   'Backup winlogon.exe
    createBackup
   
  
  'Test if it's the winlogon in system32 dir
   If FullPathName(Filename.Filename) Like GetSystemDirectory & "*" Then
   
     '--- Temporary disable Windows System File Protection ---
     'Note: 'set SFC_Blocker = Nothing' will remove it
      Set SFC_Blocker = New ClsSFC_Blocker
      SFC_Blocker.Create
   End If
   
   'Replace MsOOBE.exe
    If mi_MsOOBE_Overwrite.Checked Then MsOOBE_Overwrite
   
   'Patch winlogon.exe
    RemoveSCB Filename.Filename
   
   'Remove SFC_Blocker (only takes effect if set)
    Set SFC_Blocker = Nothing
    
   'Apply OOBE-Fix if checked
    If mi_OOBE.Checked Then OOBE_Fix
  
   'Finish - Show success MSG
    displn "Å Congratulations '" & Filename.Name & Filename.Ext & "' was patched successfully."
    displn
    displn "Remember: You must run this patch again after you have"
    displn "          installed a servicepack !"
    displn
    displn "ReRun this patch and check if it says 'already applied'"
    displn "to ensure that it's still active/wasn't undone by the"
    displn "Windows Systemfile Protection."
    
               
    err.Clear
Cmd_apply_Click_err:
     
   'Enable Buttons
    ButtonsEnable Enabled
    
   'Disable Buttons
    Cmd_apply.Enabled = False
    cmd_OffsetLocator.Visible = False
    cmd_cancel.SetFocus
    
    Select Case err
      Case 0
      
      Case Else
         displn "=> PATCH ABORTED - " & err.Description
        'Remove SFC_Blocker in case there was an error (only takes effect if set)
         Set SFC_Blocker = Nothing
    End Select

    

End Sub

Private Sub Cmd_apply_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   If Button = vbRightButton Then BrowseForFile
   
End Sub


Private Sub BrowseForFile()
   Dim File_Dlg_Struct As OPENFILENAME
   
   disp "Å Browse for file ..."

   With File_Dlg_Struct
      .lStructSize = 19 * 4
      .hwndOwner = Me.hwnd
      .lpstrFilter = "Winlogon.exe" & vbNullChar & "*.*" & vbNullChar & vbNullChar

      .nMaxFile = &HFF
      .lpstrFile = Filename.Filename & vbNullChar & Space(&HFF)

      .nMaxFileTitle = &HFF
      .lpstrFileTitle = Space(20)
      .lpstrTitle = "Please choose another Winlogon.exe for patching" & vbNullChar
      .Flags = OFN_DONTADDTORECENT Or OFN_FILEMUSTEXIST ' Or OFN_HIDEREADONLY

   
      Dim RetVal&
      RetVal = GetOpenFileName(File_Dlg_Struct)
      If RetVal Then
         displn
         
         'Convert ZeroString to VB String
         .lpstrFile = Left(.lpstrFile, InStr(.nFileExtension, .lpstrFile, vbNullChar) - 1)
         
         'Set as new FileName
         Filename.Filename = .lpstrFile
         
'         If mi_verbose.Checked = False Then mi_verbose_Click
         If mi_OOBE.Checked Then mi_oobe_Click
         
         'Check for known version
         Check_FileVersion
      Else
         displn "ABORTED!"
      End If
   
   End With
End Sub


Private Sub cmd_cancel_Click()
   End
End Sub

Private Sub cmd_cancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbRightButton Then
      displn "Forcing unlock all buttons..."
      displn "I hope you know what you're doing!"
      Cmd_apply.Enabled = True
      cmd_OffsetLocator.Visible = True
      Cmd_RestoreBackup.Visible = Not (Cmd_RestoreBackup.Visible)
'      mi_Debug_StoreInExe.Enabled = True
   End If
End Sub

Private Sub cmd_OffsetLocator_Click()
        Frm_Locator.Show vbModal
End Sub
'Private Function IsModified(Filename$) As Boolean
'   Dim PE_data As New PE_info
'   PE_info.Create
'
'   Dim myfile As New FileStream
'   myfile.Create Filename, , , True
'
'   IsModified = Not ((myfile.LastAccessDate + 1061900000) = PE_Header.TimeDateStamp)
'End Function
Private Sub MsOOBE_Restore()
   displn "Restoring " & MSOOBE_EXE & " ..."
   
  'Delete AntiWPA_Crypt.dll
   myFileDelete MSOOBE_PATH & CRYPT_DLL
  
  'Delete WPA_Kill.exe(which has name msoobe.exe)
  '...is done via rename in cause it is running (already in use)
   myFileDelete MSOOBE_PATH & MSOOBE_EXE & ".Del" 'to a sure rename will succeed (rename_dest must not exists)
   myFileRename MSOOBE_PATH & MSOOBE_EXE, MSOOBE_PATH & MSOOBE_EXE & ".Del"
  'Finally try to delete WPA_Kill.exe (whose name is now MSOOBE.EXE.del)
  'If MSOOBE.EXE.del is currently executed it will fail, but the major thing is that...
   myFileDelete MSOOBE_PATH & MSOOBE_EXE & ".Del"
  '...now the filename MSOOBE.EXE free for use (does not exists anymore)
   
  'Rename MSOOBE.com to MSOOBE.exe
   myFileRename MSOOBE_Backup.Filename, MSOOBE_PATH & MSOOBE_EXE
   
End Sub

Private Sub Cmd_RestoreBackup_Click()
   
   ButtonsEnable False
   
   On Error GoTo Cmd_RestoreBackup_err
   With Filename
      displn
      displn "Å Starting Restore..."
      
      If mi_wpa.Checked Then
         displn "Å Menu item 'Apply WPA Fix' is checked..."
      
        'Better we close the file before renaming...
         file.CloseFile
         
         displn "-> Preparing undo patch ..."
        
        'Prepare FileNames
         Dim FileExe$, FileBak$, FileDel$
         FileExe = .Name & .Ext
         FileBak = .Name & ".bak"
         FileDel = .Name & ".Del"
         
        'Set Workingdir
         If .Path <> "" Then
            ChDrive .Path
            ChDir .Path
         End If
        
        'Test if .FileName exists => raise.Err 53 and if they have the same size
         displn "Checking Backupfile " & FileBak ' & " ..."
         If FileLen(FileExe) <> FileLen(FileBak) Then _
         err.Raise vbObjectError, , FileExe & " has different size."
         
      On Error Resume Next
        
        'Delete winlogon.exe
         If myFileDelete(FileExe) = False Then
        
           'Delete dllcache\winlogon.exe
            myFileDelete "dllcache\" & FileExe
            
           'Delete old winlogon.del
            myFileDelete FileDel
         
           
            displn "Restoring Backup ..."
             
           'Rename winlogon.exe to winlogon.del
            myFileRename FileExe, FileDel
         
         End If
         
        'Rename winlogon.bak to winlogon.exe
         If myFileRename(FileBak, FileExe) = False Then
            displn "" & FileExe & " was restored successfully."
         Else
            displn "Restore of " & FileExe & " FAILED!"
         End If
         
        'Security check
         If Not FileExists(FileExe) Then
            MsgBox FileExe & " doesn't exists anymore." & vbCrLf & _
            "Please restore 'winlogon.exe'(usually in C:\windows\system32) manually." & vbCrLf & _
            "" & vbCrLf & _
            "Without winlogon.exe ya windows will start anymore !", vbCritical, "Very critical error"
         End If

      On Error GoTo Cmd_RestoreBackup_err
      
      End If
      
      If mi_OOBE.Checked Then
         displn "Å Menuitem 'Apply OOBE Fix' is checked..."
    
       ' Undo OOBE NeedActivation Registryfix (LicDll.dll!GetExpirationInfo)
         displn "-> Setting OOBE NeedActivation Registryfix invalid..."
         displn REG_DATA_OOBETimer & " = 00" ' & REG_VALUE_OOBETimer
                
         Dim reg As New Registry
         With reg
            .Create HKEY_LOCAL_MACHINE, REG_PATH_WPAEVENTS, True, REG_DATA_OOBETimer
            .RegValueDataTypeForCreateNew = REG_BINARY
            .Regdata = HexvaluesToString(0)
         End With
         
         
         displn "Restoring activation shortcut..."
         displn "Executing: 'syssetup.inf!RESTORE_OOBE_ACTIVATE'..."
         ShellExecute Me.hwnd, "open", "rundll32", "setupapi,InstallHinfSection RESTORE_OOBE_ACTIVATE 132 syssetup.inf", 0, vbHide
         
         displn "OOBE-state was successfully restored."
         
      End If

     'Restore msoobe.exe if Checked
      If mi_MsOOBE_Overwrite.Checked Then
         displn "Å Menuitem 'Replace Msoobe.exe with AntiWPA' is checked..."
         MsOOBE_Restore
      End If

'   If IsModified(Filename.Filename) Then
'      displn
'      displn "Attentions: Possibly the restored file is not the real original " & _
'             "(Last modifikation date and PE-Timestamp doesn't match)."
'      displn "To really make sure the original winlogon.exe is in restore it manually !"
'
'   End If
   
   End With
   
   IsAlreadyPatched = False
    
   Cmd_RestoreBackup.Visible = False
    
   err.Clear
Cmd_RestoreBackup_err:
    
    ButtonsEnable True
    
    Select Case err
      Case 0
         Cmd_apply.Enabled = False
      
      Case Else
         displn "=> UNDO ABORTED - " & err.Description
    End Select
   
   
   
   
End Sub

Private Sub Form_Activate()
'run this only once
   Static bAlready_executed As Boolean
   If bAlready_executed Then Exit Sub
   bAlready_executed = True
   
   Check_FileVersion
   

End Sub

Private Sub Form_Initialize()
   On Error Resume Next
   Fillchars = Space(2)
   GetFileName
   
   MSOOBE_Filename.Filename = Filename.Path & MSOOBE_PATH & MSOOBE_EXE
   MSOOBE_Backup.Filename = MSOOBE_Filename.Path & MSOOBE_Filename.Name & ".COM"

End Sub

Private Sub Form_Load()
'   Dim reg As New Registry
'   reg.Create HKEY_LOCAL_MACHINE, REG_PATH_WPAEVENTS
'
'   reg.RegValue = REG_DATA_OOBETimer
'   reg.Regdata
   
   
   
   
   Dim RetVal
   P_GENERAL_ORG = HexvaluesToString("9C 60") '&H68")
   P_GENERAL_CRK = HexvaluesToString("33 C0 C2 2C 0")
   
   Me.Caption = App.ProductName & " " & App.Major & "." & App.Minor & "." & App.Revision & App.LegalCopyright
   Txt_Console = Replace( _
                         Replace(Txt_Console, "{DATE}", App.Comments), _
                  "{VER}", App.Major & "." & App.Minor)
   
  'If launched with command '/a' it's launched instead of msoobe
   'mi_MsOOBE_Overwrite.Checked =
   If InStr(1, Command, "/", vbTextCompare) Then
      On Error Resume Next
      displn "Å Launching " & " " & MSOOBE_Backup & Command
      RetVal = ShellExecute(Me.hwnd, "open", MSOOBE_Backup, Command, "", 1)
      If RetVal <= 32 Then
         displn "=> ERROR " & RetVal
         If RetVal = SE_ERR_ACCESSDENIED Then
            displn "ACCESS DENIED - A job with an ActiveProcessLimit was " & _
                   "assigned to this process by winlogon."
            displn "However you can 'misuse' File/Open to do some " & _
                   "file rename/copy/delete operation via right click and " & _
                   "copy & paste..."
         End If

      End If
      
   End If
   
          
End Sub



Private Sub TestVersion(offset&, ORG_Data$, CRK_Data$, NameOfVersion$)
      
   With file
      'Read & compare Orignal data
       .Position = offset
       
       If .FixedString(Len(ORG_Data)) = ORG_Data Then

        ' Set Offsets
          P_offset = offset
          P_data = CRK_Data

        ' Found version
          err.Raise FoundVersion.READY_TO_PATCH, , NameOfVersion
       Else
          .Position = offset
          
          If .FixedString(Len(CRK_Data)) = CRK_Data Then
         
            'Already Patched
             err.Raise FoundVersion.ALREADY_PATCHED, , NameOfVersion
          End If
       End If
   End With
End Sub

Private Function IsCommandlineSet() As Boolean
IsCommandlineSet = Command <> ""
End Function

Private Sub GetFileName()
   
  'Get and cut systemdir and add "winlogon.exe"
   Filename.Filename = GetSystemDirectory & "\winlogon.exe"
   
  'Overwrite Filename with valid Filename from commandline
   If IsCommandlineSet Then
      displn "Å Started with Commandline: " & Command
      
      If myFileExists(Command) Then
         Filename.Filename = Command
'         displn "File Found! FullPath:""" & FullPathName(Command) & """"
         
         frmMain.mi_OOBE.Checked = False
      Else
         displn "File not Found: """ & Command & """"
         
         displn "Current Dir: """ & CurDir$ & """"
      End If
   
   End If
   
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
   End
End Sub

Private Sub lbl_Email_Click()
  ShellExecute 0, "open", _
               "mailto:" & lbl_Email.Caption & _
               "?subject=Anti WPA Patch " & App.Major & "." & App.Minor & App.Revision, _
               "", "", 0
End Sub

'//////////////////////////////////////////////////////////////////////////////////
'//// Check_FileVersion - enable/disable patcher + initialise
'//
'// Check if already patched
'//   + Check for marker
'//     -> finds version already with this version
'//
'//   + SeekTo & Check all know patch locations
'//     -> finds version patched with previous versions
'//
'//  Finds & Set Offset
'//
Private Sub Check_FileVersion()
   On Error GoTo Result
   
   ButtonsEnable False
   
   IsAlreadyPatched = False
   cmd_OffsetLocator.Visible = False
   Cmd_RestoreBackup.Visible = False
   P_offset = 0
   P_data = ""

   With file
      
      If Filename = vbNullString Then err.Raise ERR_NO_FILENAME
      
      displn "Å Opening " & FullPathName(Filename.Filename) & " ..."
      
   ' if mi_Debug_StoreInExe=false open file in readonlymode
      .Create Filename.Filename, Readonly:=(mi_Debug_StoreInExe.Checked Or mi_Cry_Remove.Checked) = False
      
    ' Get PE-infomation for use later
      PE_info.Create

#If DEV_MODE <> 1 Then
      If (mi_Debug.Checked Or mi_Debug_StoreInExe.Checked) = False Then
   '      .Position = P_MARKER_OFFSET
   '      If .Char = P_MARK_AS_PATCHED_VALUE Then err.Raise FoundVersion.ALREADY_PATCHED
   '
        ' Retail ------------------------------------
   
         TestVersion P_XP_Retail_OFFSET, _
                     P_GENERAL_ORG, _
                     P_GENERAL_CRK, _
                     "XP(5.1) 2600.0 (Retail)"
   
         TestVersion P_2K3_Retail_OFFSET, _
                     P_GENERAL_ORG, _
                     P_GENERAL_CRK, _
                     "2K3(5.2) 3790.0 (Retail)"
   
        ' Servicepacks ------------------------------------
   
         TestVersion P_XP_SP3_OFFSET, _
                     P_GENERAL_ORG, _
                     P_GENERAL_CRK, _
                     "XP(5.1) 2600.5512 (SP3)"
         
         TestVersion P_XP_SP2_RC2_OFFSET, _
                     P_GENERAL_ORG, _
                     P_GENERAL_CRK, _
                     "XP(5.1) 2600.2180 (SP2 RTM)"
   
   
         TestVersion P_XP_SP2_RC1B_OFFSET, _
                     P_GENERAL_ORG, _
                     P_GENERAL_CRK, _
                     "XP(5.1) 2600.2142 (SP2 RC1)"
   
         TestVersion P_XP_SP2_RC1A_OFFSET, _
                     P_GENERAL_ORG, _
                     P_GENERAL_CRK, _
                     "XP(5.1) 2600.2120 (SP2 RC1)"
   
         TestVersion P_XP_SP2_BETA_OFFSET, _
                     P_GENERAL_ORG, _
                     P_GENERAL_CRK, _
                     "XP(5.1) 2600.2096 (SP2 BETA)"
   
         TestVersion P_XP_SP1_OFFSET, _
                     P_GENERAL_ORG, _
                     P_GENERAL_CRK, _
                     "XP(5.1) 2600.1106 (SP1)"
      End If
#End If
     Dim WPA_Kill As New WPA_Kill
     WPA_Kill.SeekForOffset
      

     ' err.Raise FoundVersion.UNKNOWN_VERSION
   err.Clear
Result:
  ButtonsEnable True
  
  Select Case err
   Case 0   'No error
   
   Case vbObjectError + 4
        displn "ERROR: " & err.Description
        displn "Antiwpa2 does only support 32 bit files."
        displn "Please use Antiwpa3 64 bit instead."
        

   Case FoundVersion.READY_TO_PATCH
        displn "Found: Windows " & err.Description
        
        
   Case FoundVersion.ALREADY_PATCHED
'        displn "Found: Windows " & err.Description
        displn
        displn "Patch already applied!"
        
        IsAlreadyPatched = True
        Cmd_RestoreBackup.Visible = True

        
   Case FoundVersion.UNKNOWN_VERSION
        displn
        displn "Can't disable Product Activation because the"
        displn "WPA-Byte pattern was not found."
        displn ""
        displn "However hit the apply button to disable all"
        displn "selfcheckblock in " & Filename.Name & Filename.Ext & " ."
        displn "If you'd opened some other file ignore that error."
        displn "And hit apply to continue."

'        displn "Sorry I don't know how to patch this one."
'        displn "You may try to find the right offset with the offset"
'        displn "Locator. Or just hit the apply button to disable all"
'        displn "selfcheckblock in " & Filename.Name & Filename.Ext & " . But this will not"
'        displn "disable the Product Activation - it will just make"
'        displn "it execute some msec's faster and ""opens"""
'        displn "it for further patching. ;)"
'        Frm_Locator.Show vbModal
        
        cmd_OffsetLocator.Visible = True

      Case ERR_NO_FILENAME
         IsAlreadyPatched = True
      
      Case Else
        displn vbCrLf & "ERROR: " & err.Description
        IsAlreadyPatched = True
      
   
  End Select
    
  .CloseFile
  End With

End Sub


Private Sub mi_Cry_Remove_Click()
 ' Toggle Menu Item
   mi_Cry_Remove.Checked = Not (mi_Cry_Remove.Checked)
 
 ' also Enable 'Debug: Save decrypted code to exe'
   If mi_Debug_StoreInExe.Checked <> mi_Cry_Remove.Checked Then mi_Debug_StoreInExe_Click
  
   
 ' Output change in Logwindow
   Output_mi_Checkstate mi_Cry_Remove
End Sub

Private Sub mi_MsOOBE_Overwrite_Click()
   mi_MsOOBE_Overwrite.Checked = Not (mi_MsOOBE_Overwrite.Checked)
End Sub

Private Sub mi_NopOutSCB_Click()
 ' Toggle value
   mi_NopOutSCB.Checked = Not (mi_NopOutSCB.Checked)
   
 ' show Form if checked
   If mi_NopOutSCB.Checked Then
      frm_Remove_SCB_Options.Show vbModal
   End If
 
 ' Output change in Logwindow
   Output_mi_Checkstate mi_NopOutSCB
      
End Sub

'--- Menue File ---
Private Sub mi_open_Click()
   BrowseForFile
End Sub



'--- Menue Options ---
Private Sub mi_oobe_Click()
   mi_OOBE.Checked = Not (mi_OOBE.Checked)
   Output_mi_Checkstate mi_OOBE
End Sub

Private Sub mi_wpa_Click()
   mi_wpa.Checked = Not (mi_wpa.Checked)
   Output_mi_Checkstate mi_wpa
End Sub

Private Sub mi_Debug_Click()
   mi_Debug.Checked = Not (mi_Debug.Checked)
   Output_mi_Checkstate mi_Debug
End Sub

Private Sub mi_Debug_StoreInExe_Click()
   If (mi_Debug_StoreInExe.Checked = False) Then
      
      If (mi_Cry_Remove.Checked = False) Then _
         If (vbYes <> _
         MsgBox("Attention this setting will CRASH your winlogon.exe for sure." & vbCrLf & _
         "Don't use this for your active Winlogon.exe in C:\windows\system32 !" & vbCrLf & _
         "If you additionally check 'Remove Crypt Blocks' the exe may work." & vbCrLf & _
         "" & vbCrLf & _
         "Enable this setting before you disassemble your winlogon.exe." & vbCrLf & _
         "Now you will be able to see the also the protected functions." & vbCrLf & _
         "" & vbCrLf & _
         "Do you really want to enable this setting" & vbCrLf & _
         "", vbExclamation + vbYesNoCancel, "Attention Dangerous Setting !!! ")) Then _
            Exit Sub
         
      MsgBox "The next file you open will be decrypted.", vbInformation, "Debug_StoreInExe Enabled"
      mi_Debug_StoreInExe.Checked = True
      Output_mi_Checkstate mi_Debug_StoreInExe

   Else
      mi_Debug_StoreInExe.Checked = False
      Output_mi_Checkstate mi_Debug_StoreInExe
   End If
End Sub

Private Sub mi_MAP_Click()
   mi_MAP.Checked = Not (mi_MAP.Checked)
   Output_mi_Checkstate mi_MAP
End Sub


Private Sub mi_verbose_Click()
   mi_verbose.Checked = Not (mi_verbose.Checked)
   Output_mi_Checkstate mi_verbose
End Sub


'--- Menue Info ---
Private Sub mi_update_Click()
   CheckForUpdate
End Sub



Private Sub Output_mi_Checkstate(mi As Menu)
   displn "Option '" & Replace(mi.Caption, "&", "") & "' = " & IIf(mi.Checked, "En", "Dis") & "abled"
End Sub



Public Sub DispLnV(Text$)
   If mi_verbose.Checked Then displn Text
End Sub


Public Sub DispV(Text$)
   If mi_verbose.Checked Then disp Text
End Sub

Private Sub ButtonsEnable(state As Boolean)
   Cmd_apply.Enabled = state
   cmd_cancel.Enabled = state
   Cmd_RestoreBackup.Enabled = state
End Sub



Public Property Get bDoFullScan() As Variant
   bDoFullScan = mi_Debug.Checked Or _
                 mi_Debug_StoreInExe.Checked Or _
                 mi_Cry_Remove.Checked
End Property




Function myFileCopy(SourceFileName$, destinationFileName$)
         On Error Resume Next
         displn "Copying: " & SourceFileName & " -> " & destinationFileName
         
         Dim Err_Description$
         myFileCopy = FileCopy(SourceFileName, destinationFileName, Err_Description)
         If myFileCopy = False Then
            displn "=> FAILED - " & Err_Description
         End If
        
End Function
Function myFileRename(SourceFileName$, destinationFileName$) As Boolean

      Dim Error As FileRename_error
      displn "Renaming: " & SourceFileName & " -> " & destinationFileName
      
      myFileRename = FileRename(SourceFileName$, destinationFileName$, Error) <> True
      If myFileRename Then
         Select Case Error
            Case ERR_FileRename_Source_Missing
               displn "=> FAILED - Can't open source file!"
               
            Case ERR_FileRename_Dest_Already_Exists
               displn "=> FAILED - destination file already exists!"
               
            Case ERR_FileRename_Source_In_Use
               displn "=> FAILED - source file is in use!"
               
         End Select
      End If

End Function

Function myFileDelete(SourceFileName$) As Boolean
   
   displn "Deleting: " & SourceFileName
   
   Dim Err_Description$
   myFileDelete = FileDelete(SourceFileName, Err_Description$)
   
   If myFileDelete = False Then displn "=> FAILED - " & err.Description
  
End Function

Function myFileExists(SourceFileName$) As Boolean
   
   On Error Resume Next
   DispLnV "Checking existence of: " & SourceFileName
   
   Dim Err_Description$
   myFileExists = FileExists(SourceFileName, Err_Description$)
'   myFileExists = GetAttr(SourceFileName)
   
'   If myFileExists = False Then displn "=> FAILED - " & Err_Description
  
End Function

Private Sub Timer_OleDrag_Timer()
   On Error Resume Next
   Timer_OleDrag.Enabled = False

'        If mi_verbose.Checked = False Then mi_verbose_Click
         If mi_OOBE.Checked Then mi_oobe_Click
         
        'Check for known version
         Check_FileVersion
 End Sub


Private Sub Txt_Console_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo Txt_Console_OLEDragDrop_err
   
  'Set as new FileName
   Filename.Filename = Data.Files(1)
   
   Timer_OleDrag.Enabled = True
   

Txt_Console_OLEDragDrop_err:
Select Case err
Case 0

Case Else
   displn "-->Drop'n'Drag ERR: " & err.Description

End Select
End Sub
