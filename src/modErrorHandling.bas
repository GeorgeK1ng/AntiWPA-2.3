Attribute VB_Name = "modErrorHandling"
' Definieren Sie hier Ihre benutzerdefinierten Fehler. Verwenden
' Sie auf jeden Fall Nummern, die größer als 512 sind, um Konflikte
' mit OLE-Fehlernummern zu vermeiden.

'BitMasken zum Herfiltern von Modul/Funktion/Fehlen aus einer Fehlernummer 'z.b. &H130100 AND ERR_BITMASK_FUNKTION -> &H130000
 Public Const ERR_BITMASK_MODUL = &H7F000000
 Public Const ERR_BITMASK_FUNKTION = &HFF0000
 Public Const ERR_BITMASK_ERROR = &HFF00
 Public Const ERR_BITMASK_ERRORGROUP = &HFF
' Public Const ERR_BITMASK_FUNCTION_AND_MODUL = ERR_BITMASK_FUNKTION And ERR_BITMASK_MODUL

'FehlerGruppierung für unkomplizierte aber auch ungenaue Fehlerbehandlung
 Public Const ERR_GROUP_REGISTRY_GENERAL_ERROR = &H0
 Public Const ERR_GROUP_REGISTRY_NOT_FOUND = &H1
 Public Const ERR_GROUP_REGISTRY_WRITE_ERROR = &H2


'Basis für FehlerNummer dieses Moduls   40xxxxxx
 Public Const ERR_BASE_REGISTRY& = &H40000000

'Fehlerdeklarationen
 Public Const ERR_REGISTRY_GET_HKEY_INVALID_REGHANDLE = &H10100 Or ERR_GROUP_REGISTRY_NOT_FOUND _
                                                                                                      + vbObjectError + ERR_BASE_REGISTRY
 Public Const ERR_REGISTRY_LET_HKEY_RegCloseKey_FAILED = &H20100 Or ERR_GROUP_REGISTRY_GENERAL_ERROR _
                                                                                                      + vbObjectError + ERR_BASE_REGISTRY
 Public Const ERR_REGISTRY_LET_REGKEY_RegCreateKeyEx_FAILED = &H50100 Or ERR_GROUP_REGISTRY_WRITE_ERROR _
                                                                                                      + vbObjectError + ERR_BASE_REGISTRY
 Public Const ERR_REGISTRY_LET_REGKEY_RegOpenKeyEx_FAILED = &H50200 Or ERR_GROUP_REGISTRY_NOT_FOUND _
                                                                                                      + vbObjectError + ERR_BASE_REGISTRY
 Public Const ERR_REGISTRY_SET_REGKEY_RegDeleteKey_FAILED = &H60100 Or ERR_GROUP_REGISTRY_NOT_FOUND _
                                                                                                      + vbObjectError + ERR_BASE_REGISTRY
 Public Const ERR_REGISTRY_SET_REGVALUE_RegDeleteValue_FAILED = &H90100 Or ERR_GROUP_REGISTRY_NOT_FOUND _
                                                                                                      + vbObjectError + ERR_BASE_REGISTRY
 Public Const ERR_REGISTRY_GET_REGDATA_RegQueryValueExNULL_FAILED = &H100100 Or ERR_GROUP_REGISTRY_NOT_FOUND _
                                                                                                      + vbObjectError + ERR_BASE_REGISTRY
 Public Const ERR_REGISTRY_GET_REGDATA_RegQueryValueExString_FAILED = &H100200 Or ERR_GROUP_REGISTRY_NOT_FOUND _
                                                                                                      + vbObjectError + ERR_BASE_REGISTRY
 Public Const ERR_REGISTRY_GET_REGDATA_RegQueryValueExLong_FAILED = &H100300 Or ERR_GROUP_REGISTRY_NOT_FOUND _
                                                                                                      + vbObjectError + ERR_BASE_REGISTRY
 Public Const ERR_REGISTRY_GET_REGDATA_REGDATATYPE_NOT_SET = &H100400 Or ERR_GROUP_REGISTRY_GENERAL_ERROR _
                                                                                                      + vbObjectError + ERR_BASE_REGISTRY
 Public Const ERR_REGISTRY_GET_REGDATA_REGDATATYPE_NOT_SUPPORTED = &H100500 Or ERR_GROUP_REGISTRY_GENERAL_ERROR _
                                                                                                      + vbObjectError + ERR_BASE_REGISTRY
 Public Const ERR_REGISTRY_LET_REGDATA_RegSetValueEx_FAILED = &H110100 Or ERR_GROUP_REGISTRY_WRITE_ERROR _
                                                                                                      + vbObjectError + ERR_BASE_REGISTRY
 Public Const ERR_REGISTRY_REGGETALLKEYS_RegQueryInfoKey_FAILED = &H120100 Or ERR_GROUP_REGISTRY_GENERAL_ERROR _
                                                                                                      + vbObjectError + ERR_BASE_REGISTRY
 Public Const ERR_REGISTRY_REGGETALLKEYS_RegEnumKeyEx_FAILED = &H130100 Or ERR_GROUP_REGISTRY_NOT_FOUND _
                                                                                                      + vbObjectError + ERR_BASE_REGISTRY



Public Const ERR_VA_TEXT_PTR& = &H10100 Or vbObjectError
                                                                                                      
Public Function GetSimpleError() As Long
  'Bei VisualBasic Fehlern - 1 zuzückgeben
   If err.Number < 0 Then
      GetSimpleError = -1
   Else
      GetSimpleError = err.Number And ERR_BITMASK_ERRORGROUP
   End If
End Function

   
Public Sub ShowError()
   
   With err
    ' Testen ob Objekterror (vbObjectError = -2147221504)
    ' Alle positiven Fehlernummern(> 0) sind Visual Basic Fehlermeldungen
    ' Alle negativen Fehlernummern(< 0) sind benutzerdefinierte Fehlermeldungen
      If .Number <= 0 Then

         MsgBox .Source & vbCrLf & _
                .Description, _
                vbCritical, _
                "Unbehandelter Fehler: " & Format(Hex(.Number - vbObjectError), "0000 0000")
      Else
      
         MsgBox .Source & vbCrLf & _
                .Description, _
                vbCritical, _
                "Unerwarteter Fehler: " & .Number
      
      End If
   End With

End Sub
