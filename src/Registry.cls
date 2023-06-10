VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'///////////////////////////////////////////////////////////////
'//  R e g i s t r y . c l s
'//
'//  Aufgabe: Zugriff auf die Registry mittels API's
'//
'//  Hinweis: Konstantendeklationen  für Fehlernummern
'//           stehen in modErrorHandling.bas !

'Registry Errorcodes
 'Private Const ERROR_NONE = 0
 Private Const ERROR_BADDB = 1
 Private Const ERROR_BADKEY = 2
 Private Const ERROR_CANTOPEN = 3
 Private Const ERROR_CANTREAD = 4
 Private Const ERROR_CANTWRITE = 5
 Private Const ERROR_OUTOFMEMORY = 6
 Private Const ERROR_INVALID_PARAMETER = 7
 Private Const ERROR_ACCESS_DENIED = 8
 Private Const ERROR_INVALID_PARAMETERS = 87
 Private Const ERROR_NO_MORE_ITEMS = 259

'Registry AccessOptions
 Private Const READ_CONTROL = &H20000
 Private Const KEY_QUERY_VALUE = &H1
 Private Const KEY_SET_VALUE = &H2
 Private Const KEY_CREATE_SUB_KEY = &H4
 Private Const KEY_ENUMERATE_SUB_KEYS = &H8
 Private Const KEY_NOTIFY = &H10
 Private Const KEY_CREATE_LINK = &H20
'Private Const KEY_ALL_ACCESS = &H3F
 Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

 Private Const REG_OPTION_NON_VOLATILE = 0  ' VOLATILE = flüchtig; RegDaten auf Festplatte Speichern

 Private Const RESERVED = 0

 Private Const DEFAULT_SECURITY_DESCRIPTOR = 0

 Private Const REG_CREATED_NEW_KEY = &H1                      ' New Registry Key created
 Private Const REG_OPENED_EXISTING_KEY = &H0 'Vermutlich 0 - nicht getested



'--------------------------------------------------------------------------------
' Daten zur Addressierung
'--------------------------------------------------------------------------------
Public hRootKey As Long   'Vordefinierter RegistrySchlüssel (hRootKey) _
                           z.b. hRootKey_LOCAL_MACHINE
                           
Private mvarRegkey As String  'Registrypfad _
                           z.b. "Software\AbraSax\AbraSax-Updater"
                           
Private mvarRegValue As String 'RegistryAttributName _
                           z.b. "AppPath"

Private mvarRegdata As Variant 'Registrydaten _
                           z.b. "C:\Programme\app\appliation.exe"

'--------------------------------------------------------------------------------
' Zusätzliche Daten
'--------------------------------------------------------------------------------

Private mvarhKey As Long ' speichert Zugriffsnummer (Handle), _
                           welche bei Öffnen eines Registrypfades zurückgeben wird _
                           und zum Lesen/Schreiben von Daten benötigt wird

Private mvarRegValueDataType As Long 'DatenTyp des zuletzt gelesen/geschrieben _
                           Registrywertes

Public CreateIfNotExist As Boolean  'TRUE = Neuen Registrykey beim _
                           Setzen von RegData erstellen

Public RegValueDataTypeForCreateNew As Byte  'legt DatenTyp des für Registryattribut fest _
                           Default : REG_SZ

Public RetVal As Long 'Speichert den Rückgabewerte einer Funktion


'Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

'---------------------------------------------------------------
'- API-Deklarationen der Registrierung...
'---------------------------------------------------------------
'Access
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hRootKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hRootKey As Long, ByVal lpSubKey As String, ByVal RESERVED As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hRootKey As Long) As Long

'Read
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hRootKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hRootKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hRootKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long

'Write
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hRootKey As Long, ByVal lpValueName As String, ByVal RESERVED As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hRootKey As Long, ByVal lpValueName As String, ByVal RESERVED As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hRootKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hRootKey As Long, ByVal lpSubKey As String) As Long

'Browse
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, _
                        ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, _
                        ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, _
                        lpftLastWriteTime As FileTime) As Long
Private Type FileTime
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FileTime) As Long

''Vordefinierte RegistrySchlüssel (hRootKey)
'Public Property Get HKEY_CLASSES_ROOT(): HKEY_CLASSES_ROOT = &H80000000: End Property   'Public Const HKEY_CLASSES_ROOT = &H80000000
'Public Property Get HKEY_CURRENT_USER(): HKEY_CURRENT_USER = &H80000001: End Property   'Public Const HKEY_CURRENT_USER = &H80000001
'Public Property Get HKEY_LOCAL_MACHINE(): HKEY_LOCAL_MACHINE = &H80000002: End Property 'Public Const HKEY_LOCAL_MACHINE = &H80000002
'Public Property Get HKEY_USERS(): HKEY_USERS = &H80000003: End Property                 'Public Const HKEY_USERS = &H80000003
'
''Registrierungsdatentypen
Public Property Get REG_SZ(): REG_SZ = 1: End Property          ' String            'Public Const REG_SZ As Long = 1
Public Property Get REG_BINARY(): REG_BINARY = 3: End Property  ' Binär Zeichenfolge'Public Const REG_BINARY As Long = 3
Public Property Get REG_DWORD(): REG_DWORD = 4: End Property    ' 32-Bit-Zahl       'Public Const REG_DWORD As Long = 4

'///////////////////////////////////////////////////////////////////////////////////////
'////  Property hKey (Rx,w) - Aktuelles RegistryHandle zurückgeben
'//
'// 01 GET ...
'//
 Public Property Get hKey() As Long
Attribute hKey.VB_Description = "attrib"
   On Error GoTo err
   
 ' Bei ungültigen Handle Laufzeitfehler auslösen
   If mvarhKey = 0 Then
      err.Raise ERR_REGISTRY_GET_HKEY_INVALID_REGHANDLE, _
               , "No valid handle set."
   Else
   
    ' Handle in MemberVariable speichern
      hKey = mvarhKey
      
   End If
   Exit Property
   
err:
    LocalErrorHandler "hKey Property Get"
End Property
 
'//
'// 02 LET ... hKey speichern und altes Handle ggf. schliessen
'//         Beachte: Wenn eine API-Funktion(z.b. RegOpenKeyEx) _
                     einen Wert in hKey schreibt, _
                     wird hKey Property Let nicht ausgelöst)
 Private Property Let hKey(ByVal vData As Long)
     
'' Schnellversion weil - die Laufzeitfehler stören beim Debuggen _
'                      - Sehr unwahrscheinlich, das RegCloseKey fehlschlägt
'     On Error GoTo hKeyLetErr
'     RegCloseKey mvarhKey
    
    
  ' Wenn kein gültiges Handle geöffnet ist wird Laufzeitfehler ERR_REGISTRY_INVALID_REGHANDLE ausgelöst
  ' und RegCloseKey nicht aufgerufen
    On Error Resume Next

    RetVal = RegCloseKey(hKey)

  ' unerwartete Fehler weiterleiten - ERR_REGISTRY_GET_HKEY_INVALID_REGHANDLE ignorieren
    If (err.Number <> ERROR_NONE) And (err.Number <> ERR_REGISTRY_GET_HKEY_INVALID_REGHANDLE) Then GoTo hKeyLetErr:

   ' Fehlerbehandlung aktivieren - Err-Objekt wird zurückgesetzt
     On Error GoTo hKeyLetErr

   ' Wenn RegCloseKey nicht erfolgreich war Laufzeitfehler ERR_REGISTRY_LET_HKEY_RegCloseKey_FAILED
     If RetVal <> ERROR_NONE Then _
        err.Raise ERR_REGISTRY_LET_HKEY_RegCloseKey_FAILED, , _
                  "RegCloseKey (" & Hex(mvarhKey) & " [" & Regkey & "] ) failed." & vbCrLf & _
                  "RetVal= " & Hex(RetVal) & "h"

     mvarhKey = vData
     
     Exit Property
   
hKeyLetErr:
    LocalErrorHandler "hKey Property Let"
End Property




'///////////////////////////////////////////////////////////////////////////////////////
'////  Property RegValueDataType (R) - Aktuellen RegistryDatenType zurückgeben
'//
'// 03 GET ...
'//
Public Property Get RegValueDataType() As String
   
   RegValueDataType = mvarRegValueDataType

End Property



'///////////////////////////////////////////////////////////////////////////////////////
'////  Property Regkey (R,Wx,Sx) - Registrypfad Zurückgeben/Öffnen/Löschen
'//
'// 04 GET ... - Regkey Zurückgeben
'//
 Public Property Get Regkey() As String
   On Error GoTo RegkeyGetErr
   
      Regkey = mvarRegkey
   
Exit Property

RegkeyGetErr:
   LocalErrorHandler "Regkey Property Get"
End Property
'//
'// 05 LET ... - Regkey Öffnen/Neuanlegen
'//
 Public Property Let Regkey(ByVal vData As String)
   On Error GoTo RegkeyLetErr
   
    ' Registrypfad öffnen  ... (und RegHandle hKey Speichern _
      altes RegHandle wird in hKey Property Get geschlossen)
      
      Dim tmp_hKey As Long  'Nimmt hKey von API entgegen
      If CreateIfNotExist Then
      
       ' Neuen Registrykey anlegen
         Dim KeyAlreadyExisted As Long 'KeyAlreadyExisted = REG_CREATED_NEW_KEY -> es wurde ein neuer Key angelegt
         RetVal = RegCreateKeyEx(hRootKey, vData, RESERVED, vbNullString, _
                                     REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, DEFAULT_SECURITY_DESCRIPTOR, _
                                     tmp_hKey, KeyAlreadyExisted)
       ' Wenn RegCreateKeyEx nicht erfolgreich war Laufzeitfehler ERR_REGISTRY_LET_REGKEY_RegCreateKeyEx_FAILED
         If RetVal <> ERROR_NONE Then _
            err.Raise ERR_REGISTRY_LET_REGKEY_RegCreateKeyEx_FAILED, , _
                      "RegCreateKeyEx (" & vData & ") failed." & vbCrLf & _
                      "RetVal= " & Hex(RetVal) & "h"
      Else
      
       ' Nur vorhandenen Registrykey öffnen
         RetVal = RegOpenKeyEx(hRootKey, vData, REG_OPTION_NON_VOLATILE, _
                                   KEY_ALL_ACCESS, tmp_hKey)
                                   
       ' Wenn RegOpenKeyEx nicht erfolgreich war Laufzeitfehler ERR_REGISTRY_LET_REGKEY_RegOpenKeyEx_FAILED
         If RetVal <> ERROR_NONE Then _
            err.Raise ERR_REGISTRY_LET_REGKEY_RegOpenKeyEx_FAILED, , _
                      "RegOpenKeyEx (" & vData & ") failed." & vbCrLf & _
                      "RetVal= " & Hex(RetVal) & "h"



      End If
      
    ' hKey Property Let auslösen _
     (Nötig weil: Wenn eine API-Funktion(z.b. RegOpenKeyEx) einen Wert in hKey schreibt, _
      wird hKey Property Let nicht ausgelöst)
      hKey = tmp_hKey

    ' aktuell geöffneten Regkey in mvarRegkey speichern
      mvarRegkey = vData
   Exit Property
RegkeyLetErr:
   LocalErrorHandler "Regkey Property Let"
End Property
'//
'// 06 SET .. Regkey Löschen
'//
 Public Property Set Regkey(obj)
  
   On Error GoTo RegkeySetErr
   
 ' Wenn obj = Nothing wird Regvalue gelöscht
   If obj Is Nothing Then
      RetVal = RegDeleteKey(hRootKey, mvarRegkey)
     
    ' Wenn RegDeleteKey nicht erfolgreich war Laufzeitfehler ERR_REGISTRY_SET_REGKEY_RegDeleteKey_FAILED
      If RetVal <> ERROR_NONE Then _
         err.Raise ERR_REGISTRY_SET_REGKEY_RegDeleteKey_FAILED, , _
                   "RegDeletevalue(" & GetFullRegPath & ") failed." & vbCrLf & _
                     "RetVal= " & Hex(RetVal) & "h"
      
     'Registrykey schliessen & Handle Löschen -> RegValue Löschen ->RegKey Löschen
      hKey = Empty:                              RegValue = Empty:  mvarRegkey = Empty
     
   End If
   
  Exit Property
RegkeySetErr:
   LocalErrorHandler "Regkey Property Set"
 End Property



'///////////////////////////////////////////////////////////////////////////////////////
'////  Property RegValue (Rx/Wx) - RegistryAttributName festlegen
'//
'// 07 GET ...
'//
 Public Property Get RegValue() As String
   On Error GoTo RegValueGetErr
   
 ' Registrywert lesen
   RegValue = mvarRegValue
   
 Exit Property
RegValueGetErr:
   LocalErrorHandler "RegValue Property Get"
 End Property
'//
'// 08 LET ...
'//
 Public Property Let RegValue(ByVal vData As String)
    On Error GoTo RegValueLetErr
 
   'Registry Wert Speicheren
    mvarRegValue = vData
 
 Exit Property
 
RegValueLetErr:
    LocalErrorHandler "RegValue Property Let"
End Property
'//
'// 09 SET ..
'//
 Public Property Set RegValue(obj)
   
   On Error GoTo RegValueSetErr
    
 ' Wenn obj = Nothing wird Regvalue gelöscht
   If obj Is Nothing Then
      RetVal = RegDeleteValue(hKey, RegValue)
      
    
    ' Wenn RegOpenKeyEx nicht erfolgreich war Laufzeitfehler ERR_REGISTRY_SET_REGVALUE_RegDeleteValue_FAILED
      If RetVal <> ERROR_NONE Then _
         err.Raise ERR_REGISTRY_SET_REGVALUE_RegDeleteValue_FAILED, , _
                   "RegDeleteValue (" & GetFullRegPath & ") failed." & vbCrLf & _
                   "RetVal= " & Hex(RetVal) & "h"
      
     'RegValue Löschen
'      RegValue = Empty
      
   End If
    Exit Property
RegValueSetErr:
   LocalErrorHandler "RegValue Property Set"
End Property


'///////////////////////////////////////////////////////////////////////////////////////
'////  Property Regdata (Rx/Wx) - Registrydaten Lesen/Schreiben
'//
'// 10 GET ...
'//
 Public Property Get Regdata() As Variant
   
   On Error GoTo RegdataGetErr
   
 ' - RegistryWert Lesen -
  
 ' Grösse(Buffersize) und Art(mvarRegdata) der zu lesenden Daten bestimmen
   Dim Buffer As String
   Dim BufferSize As Long
  
 ' mvarRegValueDataType auf einen ungültigen Wert setzen um später zu Prüfen _
   ob RegQueryValueExNULL auch wirklich einen neuen Wert geschrieben hat
   mvarRegValueDataType = -1
   
   RetVal = RegQueryValueExNULL(hKey, RegValue, RESERVED, mvarRegValueDataType, 0, BufferSize)
 
 ' Wenn RegQueryValueExNULL nicht erfolgreich war Laufzeitfehler ERR_REGISTRY_GET_REGDATA_RegQueryValueExNULL_FAILED
   If RetVal <> ERROR_NONE Then _
      err.Raise ERR_REGISTRY_GET_REGDATA_RegQueryValueExNULL_FAILED, , _
                "RegQueryValueExNULL (" & GetFullRegPath & ") failed." & vbCrLf & _
                "RetVal= " & Hex(RetVal) & "h"
 
 ' Unterschiedliche Verarbeitung je nach Type
   Select Case RegValueDataType
      Case REG_SZ:         'Datentyp String

         ' StringBuffer zum Empfangen der Daten erzeugen
           Buffer = String(BufferSize - 1, 0)    '-> Warnung: "Buffersize - 1" Verursacht u.U einen Buffer Overflow (es kommt 1-Byte(00) zuviel zurück)
           RetVal = RegQueryValueExString(hKey, RegValue, RESERVED, 0, Buffer, BufferSize)

          ' Wenn RegQueryValueExString nicht erfolgreich war Laufzeitfehler ERR_REGISTRY_GET_REGDATA_RegQueryValueExString_FAILED
            If RetVal <> ERROR_NONE Then _
            err.Raise ERR_REGISTRY_GET_REGDATA_RegQueryValueExString_FAILED, , _
                      "RegQueryValueExString (" & GetFullRegPath & ") failed." & vbCrLf & _
                      "RetVal= " & Hex(RetVal) & "h"
           
          ' Inhalt von buffer zurückgeben
            Regdata = Buffer


      Case REG_BINARY         'Datentyp Binary
      
         ' StringBuffer zum Empfangen der Daten erzeugen
           Buffer = String(BufferSize - 1, 0)    '-> Warnung: "Buffersize - 1" Verursacht u.U einen Buffer Overflow (es kommt 1-Byte(00) zuviel zurück)
           RetVal = RegQueryValueExString(hKey, RegValue, RESERVED, 0, Buffer, BufferSize)

          ' Wenn RegQueryValueExString nicht erfolgreich war Laufzeitfehler ERR_REGISTRY_GET_REGDATA_RegQueryValueExString_FAILED
            If RetVal <> ERROR_NONE Then _
            err.Raise ERR_REGISTRY_GET_REGDATA_RegQueryValueExString_FAILED, , _
                      "RegQueryValueExString (" & GetFullRegPath & ") failed." & vbCrLf & _
                      "RetVal= " & Hex(RetVal) & "h"
                      
         ' Inhalt von buffer zurückgeben
           Regdata = Buffer

'         ' in HexString(wie z.b. "12 0a 0d") konvertieren
'           Dim nLoop As Long
'           sBinaryString = ""
'           For nLoop = 1 To Len(DataInputBuffer)
'               sBinaryString = sBinaryString & Format$(Hex(Asc(Mid$(Value, nLoop, 1))), "00") & " "
'           Next
'           Value = sBinaryString

      Case REG_DWORD:

           RetVal = RegQueryValueExLong(hKey, RegValue, RESERVED, 0, Regdata, 4)

          ' Wenn RegQueryValueExLong nicht erfolgreich war Laufzeitfehler ERR_REGISTRY_GET_REGDATA_RegQueryValueExLong_FAILED
            If RetVal <> ERROR_NONE Then _
            err.Raise ERR_REGISTRY_GET_REGDATA_RegQueryValueExLong_FAILED, , _
                      "RegQueryValueExLong (" & GetFullRegPath & ") failed." & vbCrLf & _
                      "RetVal= " & Hex(RetVal) & "h"
                      
      Case -1
          ' Fehler: RegQueryValueExNULL hat keinen Datentyp zurückgegeben
            err.Raise ERR_REGISTRY_GET_REGDATA_REGDATATYPE_NOT_SET, , _
                      "DateType was not set. RegQueryValueExNULL failed."

      Case Else
            ' Fehler: RegQueryValueExNULL hat keinen Datentyp zurückgegeben
            err.Raise ERR_REGISTRY_GET_REGDATA_REGDATATYPE_NOT_SUPPORTED, , _
            "DataType(" & RegValueDataType & ") of Registrypath " & GetFullRegPath & ") is not supported."
   End Select

 Exit Property

RegdataGetErr:
   LocalErrorHandler "Regdata Property Get"
 End Property
'//
'// 11 LET Regdata Schreiben
'//
 Public Property Let Regdata(ByVal value As Variant)
   On Error GoTo RegdataLetErr
'
   Select Case RegValueDataTypeForCreateNew
   
      Case REG_SZ
           RetVal = RegSetValueExString(hKey, RegValue, RESERVED, RegValueDataTypeForCreateNew, value & Chr$(0), Len(value) + 1)
           
      Case REG_DWORD
           RetVal = RegSetValueExLong(hKey, RegValue, RESERVED, RegValueDataTypeForCreateNew, value, 4)
      
      Case REG_BINARY
           RetVal = RegSetValueExString(hKey, RegValue, RESERVED, RegValueDataTypeForCreateNew, value, Len(value))
   End Select
   
 ' Wenn RegSetValueEx nicht erfolgreich war Laufzeitfehler ERR_REGISTRY_LET_REGKEY_RegSetValueEx_FAILED

   If RetVal <> ERROR_NONE Then _
            err.Raise ERR_REGISTRY_LET_REGDATA_RegSetValueEx_FAILED, , _
                      "RegSetValueEx (" & Regkey & ") failed." & vbCrLf & _
                      "RetVal= " & Hex(RetVal) & "h"

   
  'Wert in mvarRegdata speichern
   mvarRegdata = value
   
 Exit Property
 
RegdataLetErr:
   LocalErrorHandler "Regdata Property Let"
 End Property


'/////////////////////////////////////////////////////////////////////////////////////////////
'// 12 RegGetAllKeys - Alle RegistrySubs ermitteln und als neue Collection zurückgeben
 Public Function RegGetAllKeys() As Collection
   On Error GoTo RegGetAllKeysErr

    ' alle RegistrySubKeys ermitteln und als neue Collection zurückgeben
      Set RegGetAllKeys = New Collection  ' Neue Collection erzeugen und RegGetAllKeys zuweisen
      Dim FileTime As FileTime            ' Dummy FileTime Type erzeugen
      Dim NumberOfKeys As Long            ' Nimmt Anzahl der SubKeys auf
      Dim MaxBufferSize&, BufferSize&     ' StringBuffer Size
    
    ' Anzahl der Keys(NumberOfKeys) und längsten Keyname(MaxBufferSize) ermitteln
      RetVal = RegQueryInfoKey(hKey, 0, 0, RESERVED, NumberOfKeys, MaxBufferSize, 0, 0, 0, 0, 0, FileTime)
    
    ' Wenn RegQueryInfoKey nicht erfolgreich war Laufzeitfehler ERR_REGISTRY_REGGETALLKEYS_RegQueryInfoKey_FAILED
      If RetVal <> ERROR_NONE Then _
         err.Raise ERR_REGISTRY_REGGETALLKEYS_RegQueryInfoKey_FAILED, , _
                   "RegQueryInfoKey(" & Regkey & ") failed." & vbCrLf & _
                     "RetVal= " & Hex(RetVal) & "h"
   
    ' Hinreichend grossen Stringbuffer erzeugen
      Dim strBuffer As String
      strBuffer = Space(MaxBufferSize)
   
    ' Für alle Keys...
      Dim keyIndex As Long
      For keyIndex = 0 To NumberOfKeys - 1
         
         BufferSize = MaxBufferSize
         
       ' Keyname(strBuffer) abrufen
         RetVal = RegEnumKeyEx(hKey, keyIndex, strBuffer, BufferSize, RESERVED, 0, 0, FileTime)
       
       ' Wenn RegEnumKeyEx nicht erfolgreich war Laufzeitfehler ERR_REGISTRY_REGGETALLKEYS_RegEnumKeyEx_FAILED
         If (RetVal <> ERROR_NONE) And _
            (RetVal <> ERROR_NO_MORE_ITEMS) And _
            (RetVal <> &HEA) Then _
            err.Raise ERR_REGISTRY_REGGETALLKEYS_RegEnumKeyEx_FAILED, , _
                      "RegEnumKeyEx(" & Regkey & ") failed." & vbCrLf & _
                      "RetVal= " & Hex(RetVal) & "h"
      
       ' String nach Stringendezeichen abschneiden
         RegGetAllKeys.Add Left$(strBuffer, BufferSize)
         
      Next
 
Exit Function
 
RegGetAllKeysErr:
   LocalErrorHandler "RegGetAllKeys()"
End Function


'///////////////////////////////////////////////////////////////////////////////////////
'//  GetFullRegPath - gibt den Aktuellen Registrypfad als String zurück
'//  Beispiel         "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0\PromptForProjectOnStartup"
Public Function GetFullRegPath() As String
On Error Resume Next
   GetFullRegPath = Split("HKEY_CLASSES_ROOT,HKEY_CURRENT_USER,HKEY_LOCAL_MACHINE,HKEY_USERS")(mvarhKey And &HFFFF) & "\" & _
                    mvarRegkey & "\" & RegValue
End Function


'///////////////////////////////////////////////////////////////////////////////////////
'//  Create - Initialisiert Objekt
Public Sub Create(hRootKey&, Regkey$, Optional CreateIfNotExist As Boolean = False, Optional RegValue)
   Me.hRootKey = hRootKey
   Me.CreateIfNotExist = CreateIfNotExist
   Me.Regkey = Regkey
   If IsMissing(RegValue) = False Then Me.RegValue = RegValue
End Sub


'---------------------------------------------------------------------------------------
'Lokale Fehlerbehandlung
'---------------------------------------------------------------------------------------
'///////////////////////////////////////////////////////////////////////////////////////
'//  LocalErrorHandler - Reicht Fehler an betreffendes Objekt weiter
'
Private Sub LocalErrorHandler(Optional Location$ = "")
   Static ProjectName$           'Name des VB-Projekts
   Static IsProjectNameInitialised As Boolean           'Name des VB-Projekts
   
   With err
    ' Ermittelt AktuellenProjektname
  '-> Unzuverlässiger Code ! Bei err.raise zum Erzeugen des Fehlers darf Location nicht angeben werden
      If IsProjectNameInitialised = False Then ProjectName = .Source: IsProjectNameInitialised = True
      
   
    ' Fehler Weiterreichen
      'WICHTIGER HINWEIS FÜR ENTWICKLER:
      'Wenn VB6 beim Debuggen hier unterbricht
      'Rechtsklick in das QuellcodeFenster
      'Umschalten/"Bei nicht bearbeiteten Fehlern" wählen !
      .Raise .Number, _
             "Registry :: " & Location & IIf(err.Source <> ProjectName, _
                                             " ->" & vbCrLf & .Source, ""), _
             .Description, _
             .HelpFile, .HelpContext
   End With
End Sub



Private Sub Class_Initialize()
   
 ' Eigenschaften mit DefaultWerten initialisieren
   RegValueDataTypeForCreateNew = REG_SZ
   
End Sub

Private Sub Class_Terminate()

'   On Error Resume Next
   
 ' Handle schliessen
   If mvarhKey <> 0 Then hKey = Empty
End Sub
