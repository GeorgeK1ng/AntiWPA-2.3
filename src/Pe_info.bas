Attribute VB_Name = "Pe_info_bas"

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SE_ERR_ACCESSDENIED As Long = 5

Public Type Section
    SectionName          As String * 8
    VirtualSize          As Long
    RVAOffset            As Long
    RawDataSize          As Long
    PointertoRawData     As Long
    PointertoRelocs      As Long
    PointertoLineNumbers As Long
    NumberofRelocs       As Integer
    NumberofLineNumbers  As Integer
    SectionFlags         As Long
End Type

Public Type PE_Header
  PESignature                    As Long
  Machine                        As Integer
  NumberofSections               As Integer
  TimeDateStamp                  As Long
  PointertoSymbolTable           As Long
  NumberofSymbols                As Long
  OptionalHeaderSize             As Integer
  Characteristics                As Integer
  Magic                          As Integer
  MajorVersionNumber             As Byte
  MinorVersionNumber             As Byte
  SizeofCodeSection              As Long
  InitializedDataSize            As Long
  UninitializedDataSize          As Long
  EntryPointRVA                  As Long
  BaseofCode                     As Long
  BaseofData                     As Long

' extra NT stuff
  ImageBase                      As Long
  SectionAlignment               As Long
  FileAlignment                  As Long
  OSMajorVersion                 As Integer
  OSMinorVersion                 As Integer
  UserMajorVersion               As Integer
  UserMinorVersion               As Integer
  SubSysMajorVersion             As Integer
  SubSysMinorVersion             As Integer
  RESERVED                       As Long
  ImageSize                      As Long
  HeaderSize                     As Long
  FileChecksum                   As Long
  SubSystem                      As Integer
  DLLFlags                       As Integer
  StackReservedSize              As Long
  StackCommitSize                As Long
  HeapReserveSize                As Long
  HeapCommitSize                 As Long
  LoaderFlags                    As Long
  NumberofDataDirectories        As Long
'end of NTOPT Header
  ExportTableAddress             As Long
  ExportTableAddressSize         As Long
  ImportTableAddress             As Long
  ImportTableAddressSize         As Long
  ResourceTableAddress           As Long
  ResourceTableAddressSize       As Long
  ExceptionTableAddress          As Long
  ExceptionTableAddressSize      As Long
  SecurityTableAddress           As Long
  SecurityTableAddressSize       As Long
  BaseRelocationTableAddress     As Long
  BaseRelocationTableAddressSize As Long
  DebugDataAddress               As Long
  DebugDataAddressSize           As Long
  CopyrightDataAddress           As Long
  CopyrightDataAddressSize       As Long
  GlobalPtr                      As Long
  GlobalPtrSize                  As Long
  TLSTableAddress                As Long
  TLSTableAddressSize            As Long
  LoadConfigTableAddress         As Long
  LoadConfigTableAddressSize     As Long
  Gap                            As String * &H28&
  Sections(32)                   As Section
End Type


Public Type MSCrypt_Header
  Key1                           As Long
  Key2                           As Long
  ptrHead                        As Long
  ptrData                        As Long
  DataSize                       As Integer
  FirstRelocChunk                As Integer
End Type



Public Const REG_PATH_WPAEVENTS$ = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\WPAEvents"
Public Const REG_DATA_OOBETimer$ = "OOBETimer"
Public Const REG_VALUE_OOBETimer$ = "7F 63 3E BE EC 25 8E 19 BE A7 92 C6"

'assumption the .text Sections ist the first and .data Section the second in pe_header
Public Const TEXT_SECTION& = 0
Public Const DATA_SECTION& = 1

' Crack & original data for all Versions(XP, SP1, SP2, 2K3)
Public P_GENERAL_ORG$
Public P_GENERAL_CRK$

Public Const P_XP_Retail_OFFSET& = &H3D236
Public Const P_XP_SP1_OFFSET& = &H4907C
Public Const P_XP_SP2_BETA_OFFSET& = &H4B71C
Public Const P_XP_SP2_RC1A_OFFSET& = &H4BDB4
Public Const P_XP_SP2_RC1B_OFFSET& = &H4BFE4
Public Const P_XP_SP2_RC2_OFFSET& = &H4AE64
Public Const P_XP_SP3_OFFSET& = &H4C13C


Public Const P_2K3_Retail_OFFSET& = &H4F108

Public Const P_MARKER_OFFSET& = &H74
Public Const P_MARK_AS_PATCHED_VALUE$ = "!"

Public Const MSOOBE_PATH$ = "OOBE\"
Public Const MSOOBE_EXE$ = "MSOOBE.EXE"
Public Const CRYPT_DLL$ = "AntiWPA_Crypt.dll"

Public PE_info As New PE_info
Public PE_Header As PE_Header
Public file As New FileStream
Public file_readonly As New FileStream
Public Filename As New ClsFilename

Public dbg_file As New FileStream

Public Enum FoundVersion
   READY_TO_PATCH& = vbObjectError + &H10
   ALREADY_PATCHED& = vbObjectError + &H20
   UNKNOWN_VERSION& = vbObjectError + &H40
End Enum

Public Const ERR_NO_FILENAME& = vbObjectError + &H100

Public Enum SCBFillOptions
   NOP_OUT_COMPLETE
   NOP_OUT_NORMAL
   FILL_WITH_LONG_ASM
End Enum

