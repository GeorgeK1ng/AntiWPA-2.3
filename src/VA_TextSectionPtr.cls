VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VA_TextSectionPtr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public RaiseErrorIfInvalid As Boolean

Private VA_Start_Of_Text_Section&
Private VA_End_Of_Text_Section&
Private VA_to_RAW_Offset&

Private mvarValue As Long

Public Property Get TEXT_SECTION_RawSize() As Long
    TEXT_SECTION_RawSize = PE_Header.Sections(TEXT_SECTION).RawDataSize
End Property




'Set an Offset which is identic with fileoffset in the exe
'Note: Only relieable for offsetes inside the text section

Public Property Let TEXT_RAW_Offset(ByVal vData As Long)
   If (vData > PE_Header.SizeofCodeSection) And RaiseErrorIfInvalid Then _
      err.Raise vbObjectError, , Hex(vData) & " is bigger than PE_Header.SizeofCodeSection."
   mvarValue = vData + VA_to_RAW_Offset
End Property
Public Property Get TEXT_RAW_Offset() As Long
    TEXT_RAW_Offset = mvarValue - VA_to_RAW_Offset
End Property



'Set an Offset which is relative to the text section
'like 0004 (which will point to the 4.th byte in the text.section)
Public Property Let TEXT_Rel_Offset(ByVal vData As Long)
   If (vData > PE_Header.SizeofCodeSection) And RaiseErrorIfInvalid Then _
      err.Raise vbObjectError, , Hex(vData) & " is bigger than PE_Header.SizeofCodeSection."
   mvarValue = vData + VA_Start_Of_Text_Section
End Property
Public Property Get TEXT_Rel_Offset() As Long
    TEXT_Rel_Offset = mvarValue - VA_Start_Of_Text_Section
End Property




'set VirtualAddress like 01002C00 which must be inside the exe
Public Property Let VA(ByVal vData As Long)
   If RaiseErrorIfInvalid Then _
      RangeCheck vData, VA_Start_Of_Text_Section + PE_Header.ImageSize, _
      VA_Start_Of_Text_Section, H32(vData) & " is an invalid VA-Address."
      
   mvarValue = vData
End Property
Public Property Get VA() As Long
    VA = mvarValue
End Property



'set VirtualAddress like 01002C00 which must be inside the textsection of the exe
Public Property Let value(ByVal vData As Long)
    If RaiseErrorIfInvalid Then
   
       mvarValue = vData And RangeCheck(vData, VA_End_Of_Text_Section, VA_Start_Of_Text_Section)
    
       If mvarValue = False Then _
       err.Raise ERR_VA_TEXT_PTR, , Hex(vData) & " is a invalid VA_Address. Should be between " & Hex(VA_Start_Of_Text_Section) & " and " & Hex(VA_End_Of_Text_Section)
       
   Else
      mvarValue = vData
   End If
    
End Property
Public Property Get value() As Long
Attribute value.VB_UserMemId = 0
    value = mvarValue
End Property





Private Sub Class_Initialize()

      VA_Start_Of_Text_Section = PE_Header.BaseofCode + PE_Header.ImageBase
      VA_End_Of_Text_Section = VA_Start_Of_Text_Section + PE_Header.Sections(TEXT_SECTION).VirtualSize

      VA_to_RAW_Offset = VA_Start_Of_Text_Section - PE_Header.HeaderSize

      RaiseErrorIfInvalid = True
End Sub
