Attribute VB_Name = "modAscosiate"
Option Explicit

Private Declare Function RegSetValue& Lib "advapi32.dll" Alias "RegSetValueA" (ByVal Hkey&, ByVal lpszSubKey$, ByVal fdwType&, ByVal lpszValue$, ByVal dwLength&)
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal Hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal Hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Enum HKeyTypes
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum

Private Const ERROR_BADDB = 1&
Private Const ERROR_BADKEY = 2&
Private Const ERROR_CANTOPEN = 3&
Private Const ERROR_CANTREAD = 4&
Private Const ERROR_CANTWRITE = 5&
Private Const ERROR_OUTOFMEMORY = 6&
Private Const ERROR_INVALID_PARAMETER = 7&
Private Const ERROR_ACCESS_DENIED = 8&
Private Const MAX_PATH = 256&
'Private Const HKEY_CLASSES_ROOT = &H80000000
'Private Const HKEY_CURRENT_USER = &H80000001
'Private Const HKEY_LOCAL_MACHINE = &H80000002
'Private Const HKEY_USERS = &H80000003
'Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_DYN_DATA = &H80000006
Private Const REG_SZ = 1
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const ERROR_SUCCESS = 0&
Public Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
Public Function GetLongFilename(ByVal sShortFilename As String) As String
    Dim lRet As Long
    Dim sLongFilename As String
    sLongFilename = String$(1024, " ")
    lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    If lRet > Len(sLongFilename) Then
        sLongFilename = String$(lRet + 1, " ")
        lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    End If
    If lRet > 0 Then
        GetLongFilename = Left$(sLongFilename, lRet)
    End If
End Function
Public Function Associate(ByVal apPath As String, ByVal EXT As String, ByVal Action As String) As Boolean
'Exit Function
    Dim sKeyValue As String, ret&, lphKey&, apTitle As String, sKeyName As String
    Dim OldVal As String, UsePath As String
    
    'Set up vars
    apTitle = ParseName(apPath)
    If InStr(EXT, ".") = 0 Then EXT = "." & EXT
    sKeyName = EXT
    sKeyValue = apTitle
    
    OldVal = GetString(HKEY_CLASSES_ROOT, sKeyName, "")
    
    'create key "tsEDIT"
    If OldVal = "" Then
        ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
        If ret& <> 0 Then GoTo AssocFailed
        ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)
        If ret& <> 0 Then GoTo AssocFailed
        sKeyName = apTitle
        sKeyValue = apPath & " %1"
        ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
        If ret& <> 0 Then GoTo AssocFailed
        ret& = RegSetValue&(lphKey&, "shell\Edit_NN\command", REG_SZ, sKeyValue, MAX_PATH)
        If ret& <> 0 Then GoTo AssocFailed
          sKeyValue = apPath
        ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
        If ret& <> 0 Then GoTo AssocFailed
        ret& = RegSetValue&(lphKey&, "DefaultIcon", REG_SZ, sKeyValue, MAX_PATH)
        If ret& <> 0 Then GoTo AssocFailed
    Else
        UsePath = OldVal & "\shell\" & IIf(LCase(Action) = "o", "Open", "Edit With tsEDIT")
        OldVal = GetString(HKEY_CLASSES_ROOT, UsePath, "")
        
    End If
    
   Associate = True
  Exit Function
AssocFailed:
  Associate = False
End Function

Public Function ParseName(ByVal sPath As String) As String
    Dim intX As Integer: intX = Len(sPath)
    Do Until InStr(intX, sPath, "\") <> 0:      intX = intX - 1:        Loop
    ParseName = Mid(sPath, intX + 1)
End Function



'#######################


Public Sub SaveKey(Hkey As HKeyTypes, strPath As String)
    Dim keyhand&
    Call RegCreateKey(Hkey, strPath, keyhand&)
    Call RegCloseKey(keyhand&)
End Sub


Public Function GetString(Hkey As HKeyTypes, strPath As String, strValue As String)
    'EXAMPLE:
    '
    'text1.text = getstring(HKEY_CURRENT_USE
    '     R, "Software\VBW\Registry", "String")
    '
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    Dim lValueType As Long
    Dim r As Long
    Call RegOpenKey(Hkey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)


    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)


        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))


            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
End Function


Public Sub SaveString(Hkey As HKeyTypes, strPath As String, strValue As String, strdata As String)
    'EXAMPLE:
    '
    'Call savestring(HKEY_CURRENT_USER, "Sof
    '     tware\VBW\Registry", "String", text1.tex
    '     t)
    '
    Dim keyhand As Long
    Dim r As Long
    Call RegCreateKey(Hkey, strPath, keyhand)
    Call RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    Call RegCloseKey(keyhand)
End Sub


Public Function DeleteValue(ByVal Hkey As HKeyTypes, ByVal strPath As String, ByVal strValue As String)
    'EXAMPLE:
    '
    'Call DeleteValue(HKEY_CURRENT_USER, "So
    '     ftware\VBW\Registry", "Dword")
    '
    Dim keyhand As Long
    Call RegOpenKey(Hkey, strPath, keyhand)
    Call RegDeleteValue(keyhand, strValue)
    Call RegCloseKey(keyhand)
End Function


Public Function DeleteKey(ByVal Hkey As HKeyTypes, ByVal strPath As String)
    'EXAMPLE:
    '
    'Call DeleteKey(HKEY_CURRENT_USER, "Soft
    '     ware\VBW\Registry")
    '
    Dim keyhand As Long
    Call RegDeleteKey(Hkey, strPath)
End Function


