VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PE_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private PEStart&
'Private Const CHECKSUM_OPEN_FAILURE As Long = 1
'Private Const CHECKSUM_MAP_FAILURE As Long = 2
'Private Const CHECKSUM_MAPVIEW_FAILURE As Long = 3
'Private Const CHECKSUM_UNICODE_FAILURE As Long = 4
Private Declare Function MapFileAndCheckSum Lib "imagehlp.dll" Alias "MapFileAndCheckSumA" (ByVal Filename As String, ByRef HeaderSum As Long, ByRef CheckSum As Long) As Long

Public Sub Create()
'     '--- find PE-signature ---
'     'Get First 0x400 Bytes
'      Dim tmpstr$
'      file.Position = 0
'      tmpstr = file.FixedString(&H400)
'
'     'Locate start of PE-header
'      PEStart = InStr(1, tmpstr, "PE" & vbNullChar & vbNullChar, vbBinaryCompare)
'      If PEStart = 0 Then err.Raise vbObjectError + 1, , "No PE-Header Found"
    
     '--- find PE-signature ---
     'Check DOS Header
      Dim tmpstr$
      file.Position = 0
     
     'MZ DOS-Header->e_magic
      If file.intValue <> &H5A4D Then err.Raise vbObjectError + 1, , "No ExeFile DOS-Header.e_magic<>""MZ"""

     
     'Locate & Validate PE-header
      file.Position = &H3C '   DOS-Header->e_lfanew
      PEStart = file.longValue
      PEStart = PEStart + 1
      If file.longValue = &H4550 Then err.Raise vbObjectError + 2, , "No ExeFile 'PE-Header.Signature<>""PE"""
    
    '  --- get PE_Header  ---
      Dim hFile&
      hFile = FreeFile
      Open file.Filename For Binary Access Read As #hFile
        Get hFile, PEStart, PE_Header
      Close hFile
      
    ' Validate Maschine Type
      If PE_Header.Machine <> &H14C Then
         If PE_Header.Machine = &H8664 Then
            err.Raise vbObjectError + 4, , "PE-Header.Signature=HDR64_MAGIC!"
         Else
           err.Raise vbObjectError + 3, , "Unsupported PE-Header.Signature<>I386(0x14C)."
         End If
      End If
     
'    ' Validate Optionaldata Type
'      If PE_Header.Magic <> &H10B Then
'         err.Raise vbObjectError + 5, , "PE_Header.Magic <> PE32!"
'      End If
     
'      If PE_Header.NumberofDataDirectories > &HE Then
'         err.Raise vbObjectError + 5, , "PE_Header.NumberofDataDirectories must be greater than 14"
'      End If
'imagever 3..5 <-winXP / 6 vista
     
End Sub

Public Sub WriteHeader()
   If PEStart = 0 Then err.Raise vbObjectError, , "You must call PE_info::Create first!"
    
    '  --- get PE_Header  ---
      Dim hFile&
      hFile = FreeFile
      Open file.Filename For Binary Access Write As #hFile
        Put hFile, PEStart, PE_Header
      Close hFile

End Sub

Public Function UpdateChecksum&() '(Optional ByRef Output_ChecksumCurrent&, Optional ByRef Output_ChecksumNew&)
   If PEStart = 0 Then err.Raise vbObjectError, , "You must call PE_info::Create first!"
   UpdateChecksum = MapFileAndCheckSum(file.Filename, 0, PE_Header.FileChecksum)
   WriteHeader
End Function


