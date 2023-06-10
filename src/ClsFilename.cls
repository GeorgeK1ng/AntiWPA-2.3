VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsFilename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Private
Public mvarFileName$


Public Property Get Filename() As String
Attribute Filename.VB_UserMemId = 0
   Filename = mvarFileName
End Property

Public Property Let Filename(ByVal vNewValue As String)
   mvarFileName = vNewValue
End Property

'         FileName.WholeName = .lpstrFile
'         FileName.Path = Left(.lpstrFile, .nFileOffset)
'         FileName.FileName = Mid(.lpstrFile, .nFileOffset, .nFileExtension - .nFileOffset)
'         FileName.Ext = Mid(.lpstrFile, .nFileExtension, -1)


Public Property Get Path() As String
   Path = Left(mvarFileName, InStrRev(mvarFileName, "\"))
End Property

Public Property Let Path(ByVal vNewValue As String)
   If Right(vNewValue, 1) = "\" Then
      mvarFileName = vNewValue & Name & Ext
   Else
      mvarFileName = vNewValue & "\" & Name & Ext
   End If
End Property


Public Property Get Name() As String
On Error Resume Next
   Name = Split(Mid(mvarFileName, Len(Path) + 1), ".")(0)
End Property

Public Property Let Name(ByVal vNewValue As String)
   mvarFileName = Path & vNewValue & Ext
End Property



Public Property Get Ext() As String
On Error Resume Next
   Ext = Mid(mvarFileName, InStrRev(mvarFileName, "."))
End Property

Public Property Let Ext(ByVal vNewValue As String)
      
   mvarFileName = Path & Name & IIf(Left(vNewValue, 1) = ".", "", ".") & vNewValue
End Property
