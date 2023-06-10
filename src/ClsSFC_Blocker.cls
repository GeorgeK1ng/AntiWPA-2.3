VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSFC_Blocker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private reg As New Registry, reg2 As New Registry
Private orgValue_ServicePackSourcePath$
Private orgValue_SourcePath$
Public WaitForSFCMessagebox As Boolean

Public Sub Create()
   With frmMain
      .displn "Å Disabling Windows System File Protection temporary:"
     
     'open Registry
     'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\\ServicePackSourcePath
      reg.Create HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Setup", False, "ServicePackSourcePath"
      
     'test if "ServicePackSourcePath" is a valid path
      On Error Resume Next
      GetAttr reg.Regdata
      If Err = 0 Then
      
         ' save ServicePackSourcePath Value
           orgValue_ServicePackSourcePath = reg.Regdata
           
           .displn "RegDelete: ServicePackSourcePath="""""
         
         ' delete ServicePackSourcePath Value
           reg.Regdata = ""
      End If
   
      On Error Resume Next
      reg2.Create HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Setup", False, "SourcePath"
      GetAttr reg2.Regdata
      If Err = 0 Then
         ' save Value
           orgValue_SourcePath = reg2.Regdata
         
           .displn "RegDelete: SourcePath="""""
         
         ' delete Value
           reg2.Regdata = ""
   
      End If
   End With
   
   WaitForSFCMessagebox = False
End Sub



Private Sub Class_Terminate()
   With frmMain
      If WaitForSFCMessagebox Then
         .displn "Waiting for ""Windows Systemfile Protection"" dialogbox"
         #If DEV_MODE Then
         #Else
         MsgBox "Now the ""Windows Systemfile Protection"" dialogbox should popup." & vbCrLf & _
                "Click on ""Abort"" and then click on ""Yes""." & vbCrLf & vbCrLf & _
                "Note: On Serversystems just wait about 5 Sec and then click OK.", vbInformation, _
               "Please wait for the ""Windows Systemfile Protection"" dialogbox"
         #End If
      End If
     
     
     'Restore original settings
      .displn "Å Restoring Windows System File Protection data:"
      
      If orgValue_ServicePackSourcePath <> "" Then
         .displn "RegSetValue: ServicePackSourcePath="
         .displn "             """ & orgValue_ServicePackSourcePath & """"
         reg.Regdata = orgValue_ServicePackSourcePath
      End If
      
      If orgValue_SourcePath <> "" Then
         .displn "RegSetValue: SourcePath="
         .displn "             """ & orgValue_SourcePath & """"
         reg2.Regdata = orgValue_SourcePath
      End If
   End With
End Sub
