VERSION 5.00
Begin VB.Form Frm_Locator 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Offset Finder"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7590
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "Frm_Offset_Locator.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.CommandButton cmd_Quit 
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   2.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   60
      Left            =   7800
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Width           =   15
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2190
      Left            =   3360
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "Frm_Offset_Locator.frx":0F12
      Top             =   2280
      Width           =   3975
   End
   Begin VB.CommandButton cmd_softiceinfo 
      Caption         =   "Finding the right offset with Numega Softice"
      Height          =   450
      Left            =   3360
      TabIndex        =   8
      Top             =   4560
      Width           =   3975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Offsets Hints"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   3975
      Begin VB.ListBox List_Offset_Hints 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         ItemData        =   "Frm_Offset_Locator.frx":1023
         Left            =   120
         List            =   "Frm_Offset_Locator.frx":1025
         TabIndex        =   4
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version           Offset   Diff       Diff_Next"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3690
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Possible Patch Offsets"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox List_Offsets 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         ItemData        =   "Frm_Offset_Locator.frx":1027
         Left            =   120
         List            =   "Frm_Offset_Locator.frx":1029
         TabIndex        =   1
         Top             =   600
         Width           =   2880
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Num  Offset   Diff"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1590
      End
   End
   Begin VB.TextBox txt_devWinlogonPath 
      Height          =   360
      Left            =   120
      TabIndex        =   13
      Top             =   6720
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.TextBox txt_hiew 
      Height          =   375
      Left            =   3720
      TabIndex        =   14
      Text            =   "HIEW32.EXE"
      Top             =   5040
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton Cmd_Update 
      Caption         =   "Check for Update"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4560
      Width           =   2295
   End
   Begin VB.ListBox List_dev_WinlogonFilename 
      Height          =   2220
      ItemData        =   "Frm_Offset_Locator.frx":102B
      Left            =   120
      List            =   "Frm_Offset_Locator.frx":1050
      TabIndex        =   12
      Top             =   4440
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   $"Frm_Offset_Locator.frx":117D
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   5280
      Width           =   7215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Attention:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   4920
      Width           =   1695
   End
End
Attribute VB_Name = "Frm_Locator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Start of decryptionblock
'001B:0104C31C  9C                  PUSHFD      <- P_GENERAL_ORG Searchstring
'001B:0104C31D  60                  PUSHAD      <- P_GENERAL_ORG Searchstring
'001B:0104C31E  51                  PUSH      ECX
'001B:0104C31F  6800C70401          PUSH      0104C700   ;end of decrpytiondata; Start of Decrpytiondataheads
'001B:0104C324  6A01                PUSH      01
'001B:0104C326  E82635FCFF          CALL      0100F851
'001B:0104C32B  FFE0                JMP       EAX  <-additional_check Searchstring
'001B:0104C32D  90  /8F (SP1)       NOP                  ;Start of decrpytiondata
'winlogon SP2

'Winlogon retail
'0103DC36: 9C                           pushfd
'0103DC37: 60                           pushad
'0103DC38: 6843E90301                   push        00103E943  -----? (1)
'0103DC3D: 6801000000                   push        000000001 ;"   ?"
'0103DC42: E855FAFCFF                   call       .00100D69C  -----? (2)
'0103DC47: 83C404                       add         esp,004 ;"?"
'0103DC4A: FFE0                         jmp         eax

Private Const List_Offsets_Seperator$ = "  "
Private Const List_Offset_Hints_Seperator$ = vbTab

Private List_Offsets_bNoUpdate As Boolean

#If DEV_MODE = 1 Then
Private ProcID&
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As INPUT_TYPE, ByVal cbSize As Long) As Long
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long

 
Private Type KEYBDINPUT
   wVk As Integer
   wScan As Integer
   dwFlags As Long
   time As Long
   dwExtraInfo As Long
End Type
 
 
Private Type INPUT_TYPE
   dwType As Long
   xi As KEYBDINPUT
End Type
 
 
 
'KEYBDINPUT dwFlags-Konstanten
Private Const KEYEVENTF_EXTENDEDKEY = &H1 'Der Scancode hat das Präfix &HE0
Private Const KEYEVENTF_KEYUP = &H2 'Die angegebene Taste wird losgelassen
Private Const KEYEVENTF_UNICODE = &H4 'Benutzt ein Unicode Buchstaben der nicht von einen der Tastaturcodes stammt welcher eine Tastatureingabe Simuliert
 
'INPUT_TYPE dwType-Konstanten
Private Const INPUT_MOUSE = 0 'Mauseingabe
Private Const INPUT_KEYBOARD = 1 'Tastatureingabe
Private Const INPUT_HARDWARE = 2 'Hardwarenachricht
 
Private Sub SendKey(char As Byte)
   Dim IT As INPUT_TYPE
   
'   With IT
'      'KeyDown
'       .dwType = INPUT_KEYBOARD
'       .xi.wVk = Char
'       .xi.wScan = MapVirtualKey(Char, 0)
'       .xi.dwFlags = 0
'       If SendInput(1&, IT, 28&) = 0 Then Debug.Print "Sendingkeys failed."
'
'      'KeyUp
'       .xi.dwFlags = KEYEVENTF_KEYUP
'       If SendInput(1&, IT, 28&) = 0 Then Debug.Print "Sendingkeys failed."
'    End With
'
'    Sleep (10)
'   'Process Windows messages
''    DoEvents
      
End Sub

Private Sub SendKeys(vbKeys$)
   Dim i As Integer
   For i = 1 To Len(vbKeys)
      SendKey (Asc(Mid(vbKeys, i)))
   Next
End Sub

#End If

Private Sub cmd_Quit_Click()
   Unload Me
End Sub

Private Sub cmd_softiceinfo_Click()
   Text1.Visible = True
End Sub

Private Sub Cmd_Update_Click()
   frmMain.CheckForUpdate
End Sub

Private Sub Form_Load()
   'On Error Resume Next
   file.Filename = Filename.Filename
   file.CloseFile
   
   PE_info.Create
   
#If DEV_MODE = 1 Then
   List_dev_WinlogonFilename.Visible = True
   txt_devWinlogonPath.Visible = True
   txt_hiew.Visible = True
#End If



   
'--- Fill Possible_Patch_Offsets List ----
   With PE_Header.Sections(TEXT_SECTION)
     
     'Calculate VirtualAddres_to_RAW_Offset value (for later use)
      Dim VA_to_TEXT_Rel_Offset&
      VA_to_TEXT_Rel_Offset = PE_Header.BaseofCode + PE_Header.ImageBase ' - PE_Header.HeaderSize

      
      Dim Start_Of_Text_Section&
      Start_Of_Text_Section = .PointertoRawData
     
     'Seek to .text section
      file.Position = Start_Of_Text_Section
      
     'Read whole .text section into buffer
      Dim input_buffer As New StringReader
      input_buffer = file.FixedString(.RawDataSize)
      
            
   End With
   
    Dim Possible_Patch_Offset&, Last_Offset&
    
    
   'JMP EAX Filter
    Dim JMP_EAX$
    JMP_EAX = HexvaluesToString("FF E0") ' & Chr(&H90)
    
    Dim PUSH_XXX$
    PUSH_XXX = Chr(&H68)
    
    Dim CALL_XXX$
    CALL_XXX = Chr(&HE8)
    
   'Clear List
    List_Offsets.Clear
    
    With input_buffer
    Do
  
        .StorePos
        'Find Possible_Patch_Offset
         Possible_Patch_Offset = .Findstring(P_GENERAL_ORG)
         
         .RestorePos
         Dim tmp1&
         tmp1 = .Findstring(P_GENERAL_CRK, IIf(Possible_Patch_Offset, Possible_Patch_Offset, .Length) - .Position)
         
         Dim bIsPatched As Boolean
         bIsPatched = tmp1
         If bIsPatched Then Possible_Patch_Offset = tmp1 + 2
         .Position = Possible_Patch_Offset + 2
        
        'if no more Possible_Patch_Offsets found exit loop
         If Possible_Patch_Offset = 0 Then Exit Do
        
'        Debug.Assert bIsPatched = False
        If bIsPatched = False Then
        
            'Seek to PUSH <ptrToDecryptionDataHeads>
             Dim ptrToDecryptionDataHeads
             ptrToDecryptionDataHeads = .Findstring(PUSH_XXX, 2)
             If ptrToDecryptionDataHeads = 0 Then GoTo NextItem 'err.Raise vbObjectError + 1, "", "ptrToDecryptionDataHeads not found"
            
            'Get pointer To ptrToDecryptionDataHeads
             ptrToDecryptionDataHeads = .int32 ' - VA_to_TEXT_Rel_Offset + 1
            
            'Check ptrToDecryptionDataHeads for integrity
             If (ptrToDecryptionDataHeads < VA_to_TEXT_Rel_Offset) Or (ptrToDecryptionDataHeads > VA_to_TEXT_Rel_Offset + PE_Header.Sections(TEXT_SECTION).VirtualSize) Then GoTo NextItem
             ptrToDecryptionDataHeads = ptrToDecryptionDataHeads - VA_to_TEXT_Rel_Offset
         End If


        'Seek over call
         .Findstring CALL_XXX, 8: .Move (4)
        
        'additional check ("JMP EAX Filter)
         If .Findstring(JMP_EAX, 8) Then
       
'         '---- Decrypt first Block ----
'         .Position = ptrToDecryptionDataHeads
'         Dim Key1: Key1 = .int32
'         Dim Key2: Key2 = .int32
'
'       ' Copy Crypted Data into CryptedData Buffer
'         Dim CryData As New StringReader
'         Dim CryDataEnd:   CryDataEnd = .int32
'         Dim CryDataStart:   CryDataStart = .int32
'         Dim CryDataSize:  CryDataSize = CryDataEnd - CryDataStart
'               .StorePos
'         .Position = CryDataStart - VA_to_TEXT_Rel_Offset + 1
'         CryData = .FixedString(CryDataSize)
'               .RestorePos
'
'        '---- Decrypt first and next Block(s) ----
'         Do
'
'            Dim CryDataBlockSize&: CryDataBlockSize = .int16
'            Dim RelocMarker: RelocMarker = .int16
'
'         '  Check for end of Header
'            If (CryDataBlockSize = 65535) And (RelocMarker = 65535) Then Exit Do
''Debug.Print Hex(Not (RelocMarker)), Hex((Key1 And &HFFFF))
'
'           ' DeCrypt data
'             CryData.StorePos
'             Dim tmp$: tmp = CryData.FixedString(CryDataBlockSize)
'             CryData.RestorePos
'
''Debug.Print "Cry:", Hex(Asc(tmp))
'             DeCrypt tmp, tmp, Len(tmp), Key1, Key2
''Debug.Print "unCry:", Hex(Asc(tmp))
'
'           ' copy tmp-data to CryDataBlock
'             Dim CryDataBlock As New StringReader: CryDataBlock = tmp
'             CryDataBlock.DisableAutoMove = True
'
''          ' Skip RelocChunk
''            .FindInt (RelocMarker)
'
'           '---- Apply Relocation Fixup ----
'            Dim Keys As New StringReader
'            Keys.EOS = False
'            Keys.int32 = Key1
'            Keys.int32 = Key2
'
'           'Set source
'            Dim Src&: Src = CryDataStart + (CryData.Position - 1)    ' = Start of current CryData block
'
'           'Set srcination
'            Dim src&: src = Possible_Patch_Offset + VA_to_TEXT_Rel_Offset
'
'            Do
'
'              ' Get Key for Reloc
'                If Keys.EOS Then Keys.EOS = False     'set position to first if end Of String is reached
'                Dim key&: key = Keys.int16 ' And &HFFF0
'
'
'
'
'              ' Calculate reloc fixUp
'                Dim fixUp&
'                fixUp = src - Src - key
'
'
'              ' Get RelocChunk & Exit if Reloc Start/Stop Marker
'                Dim reloc&: reloc = .int16           'get reloc
'                If reloc = RelocMarker Then Exit Do  'is Start/Stop Marker
'                reloc = reloc Xor (Key1 And 65535)   'Decrypt reloc
'
'              ' Apply Fixup
'                With CryDataBlock
'                  .Position = reloc + 1
'Debug.Print Hex(key), Hex(.int32), Hex(reloc)
'                  .int32 = .int32 + fixUp
'                End With
'
'
'
'            Loop While True
'
'            'Write Decrypted and Relocated data
'             CryData.FixedString = CryDataBlock
'
'
'           ' Skip Fillbytes
'             Do: Loop While .int16 = 0
'             .Move -2
'
'            'Read new keys
'             Key1 = (.int32 Xor Key1) Or 1
'             Key2 = .int32 Xor Key2
'
'           Loop While True
'
'DeCrypt_Done:
'
'
'
'
''           Dim binfile As New FileStream
''           binfile.Create FileName.Path & Hex(Possible_Patch_Offset - 1 + Start_Of_Text_Section) & ".bin", True, False
''           binfile.FixedString(-1) = CryData
''           binfile.CloseFile
''
''     'Seek to .text section
''      file.Position = Possible_Patch_Offset - 1 + Start_Of_Text_Section
''
''     'write whole .text section to winlogon.exe
''      file.FixedString(-1) = CryData
'
'
'
'
'           If 0 <> InStr(1, CryData, HexvaluesToString("D7 04 07 80 0f")) Then '3 40 0 80
           
                 Dim diff&
                 diff = Possible_Patch_Offset - Last_Offset
                 
               ' Output Possible_Patch_Offset
                 List_Offsets.AddItem Format(List_Offsets.ListCount, "000") & List_Offsets_Seperator & _
                               Hex(Possible_Patch_Offset _
                                   + Start_Of_Text_Section _
                                   - 2) & IIf(bIsPatched, "*", "") & List_Offsets_Seperator & _
                                   IIf(List_Offsets.ListCount = 0, "-", Hex(Possible_Patch_Offset - Last_Offset) & List_Offsets_Seperator)
                 
                 Last_Offset = Possible_Patch_Offset
'            End If
'       Else
'           Dim badcounter
'           badcounter = badcounter + 1
'           Debug.Print Hex(Possible_Patch_Offset)
       End If
NextItem:
    Loop While True
    End With
'    Debug.Print "filtered:", badcounter
    
    With List_Offsets
      .Enabled = .ListCount
      List_Offset_Hints.Enabled = .Enabled
       If .Enabled = False Then
         .AddItem "No possible patch"
         .AddItem "locations found."
         .AddItem ""
         .AddItem "Maybe file is too "
         .AddItem "different!"
       Else
         HeuristicScan
       End If
    End With
    

    
    
   List_Offset_Hints.Clear
    
    '--- Fill Hint_Offsets List ----
    fill_List_Offset_Hints _
      "XP Retail", P_XP_Retail_OFFSET, "1EC", "E31", _
      "XP SP1", P_XP_SP1_OFFSET, "3CC", "3EB", _
      "XP SP2 Beta", P_XP_SP2_BETA_OFFSET, "464", "488", _
      "XP SP2 RC1A", P_XP_SP2_RC1A_OFFSET, "444", "4B0", _
      "XP SP2 RC1B", P_XP_SP2_RC1B_OFFSET, "444", "478", _
      "XP SP2 RC2", P_XP_SP2_RC2_OFFSET, "435", "4A0", _
      "2K3 Retail", P_2K3_Retail_OFFSET, "3DB", "465"

End Sub
Private Sub HeuristicScan()
   On Error GoTo HeuristicScan_err
   
   Dim SelectedListitems As New Collection
   
   With List_Offsets
      
      Dim i&, diff&
      For i = 1 To .ListCount
         diff = CLng("&h" & Split(.List(i), List_Offsets_Seperator)(2))
         If (&HE00 <= diff) And (diff <= &H1A00) Then
            diff = CLng("&h" & Split(.List(i + 1), List_Offsets_Seperator)(2))
            
            If (&H300 <= diff) And (diff <= &H500) Then
            ' highlight SelectThisItem in List_Offsets
               HighlightAndSelectItem (i + 1)
               Exit For
            End If
         End If
      Next
   
   
   End With


HeuristicScan_err:
End Sub

Private Sub fill_List_Offset_Hints(ParamArray TextAndOffset())
   On Error Resume Next
   Dim i&
   For i = LBound(TextAndOffset) To UBound(TextAndOffset) Step 4
      
      List_Offset_Hints.AddItem _
         "Win" & TextAndOffset(i) & List_Offset_Hints_Seperator & _
         Hex(TextAndOffset(i + 1)) & List_Offset_Hints_Seperator & _
         TextAndOffset(i + 2) & List_Offset_Hints_Seperator & _
         TextAndOffset(i + 3)
   
   Next
End Sub



Private Sub Form_Unload(Cancel As Integer)
#If DEV_MODE = 1 Then
      Shell Environ("systemroot") & "\System32\Taskkill.exe /pid " & ProcID, vbNormalFocus
#End If
End Sub

Private Sub List_Offset_Hints_Click()
   On Error Resume Next
   
   Dim offset&
   offset = "&h" & Split(List_Offset_Hints, List_Offset_Hints_Seperator)(1)
   
   Dim item
   For item = 1 To List_Offsets.ListCount
      If offset <= CLng("&h" & Split(List_Offsets.List(item), List_Offsets_Seperator)(1)) Then
         
       ' highlight SelectThisItem in List_Offsets
         HighlightAndSelectItem (item)
         Exit For
      End If
   Next
   
End Sub

' highlight Select an Item in List_Offsets
Private Sub HighlightAndSelectItem(ItemIndex)
         List_Offsets_bNoUpdate = True
         With List_Offsets
            .Selected(.ListCount - 1) = True
            .Selected(ItemIndex - 6) = True
            List_Offsets_bNoUpdate = False
            .Selected(ItemIndex) = True
         End With
End Sub


#If DEV_MODE = 1 Then

Private Sub List_dev_WinlogonFileName_Click()
  On Error Resume Next
   
   
   Filename.Filename = txt_devWinlogonPath & List_dev_WinlogonFilename
   Form_Load
'   Form_Unload False ' Shell Environ("systemroot") & "\System32\Taskkill.exe /pid " & ProcID, vbNormalFocus
   
'   Dim tmp$
'   tmp = txt_hiew.Text & " """ & FileName.FileName & """"
'   Debug.Print "Executing:", tmp
'   Debug.Assert (Len(tmp) < 150) 'Shell() fails if commandline is longer than 150 bytes
'
'  'Please check dos-box settings:Properties/Layout/windowspuffersize/height=26
'   ProcID = Shell(tmp, vbNormalFocus)
'
'   Sleep (500)
'   SendKeys Chr(vbKeyReturn) & Chr(vbKeyReturn)
'   Me.SetFocus
   
  
End Sub

Private Sub List_Offsets_Click()
#Else
Private Sub List_Offsets_DblClick()
#End If

   On Error GoTo List_Offsets_Click_err
   
   If List_Offsets_bNoUpdate Then Exit Sub
   
  'set selected Offset as new patch offset
   frmMain.P_offset = "&h" & Split(List_Offsets.Text, List_Offsets_Seperator)(1)
   
  'Set Patchdata
  frmMain.P_data = P_GENERAL_CRK


#If DEV_MODE = 1 Then
'   AppActivate "HIEW:"
'   Sleep (100)
'   SendKeys Chr(vbKeyF5) & _
'            Hex(frmMain.P_offset) & Chr(vbKeyReturn)
'   Me.SetFocus
'
'
'
#Else
   
   frmMain.displn "!!! Attention. User has set new patch offset to: " & Hex(frmMain.P_offset) & " !!!"
   
   cmd_Quit_Click
#End If

List_Offsets_Click_err:
   
End Sub




Private Sub Text1_DblClick()
   Text1.Visible = False
End Sub

