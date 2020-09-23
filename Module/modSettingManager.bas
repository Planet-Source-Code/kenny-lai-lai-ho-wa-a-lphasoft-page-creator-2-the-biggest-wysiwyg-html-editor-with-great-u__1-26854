Attribute VB_Name = "modSettingManager"
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004


Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long


Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long


Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long


Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    Const KEY_ALL_ACCESS = &HF003F
    Const HKEY_DYN_DATA = &H80000006
    Const REG_BINARY = 3
    Const REG_DWORD = 4
    Const REG_DWORD_BIG_ENDIAN = 5
    Const REG_DWORD_LITTLE_ENDIAN = 4
    Const REG_EXPAND_SZ = 2
    Const REG_LINK = 6
    Const REG_MULTI_SZ = 7
    Const REG_NONE = 0
    Const REG_RESOURCE_LIST = 8
    Const REG_SZ = 1


Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String)
    Dim lResult As Long
    Dim lValueType As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    On Error GoTo 0
    lResult = RegQueryValueEx(hKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufSize)


    If lResult = ERROR_SUCCESS Then


        If lValueType = REG_SZ Then
            strBuf = String(lDataBufSize, " ")
            lResult = RegQueryValueEx(hKey, strValueName, 0&, 0&, ByVal strBuf, lDataBufSize)


            If lResult = ERROR_SUCCESS Then
                RegQueryStringValue = StripTerminator(strBuf)
            End If
        End If
    End If
End Function


Public Function GetString(hKey As Long, strpath As String, strvalue As String)
    Dim keyhand&
    Dim datatype&
    R = RegOpenKey(hKey, strpath, keyhand&)
    GetString = RegQueryStringValue(keyhand&, strvalue)
    R = RegCloseKey(keyhand&)
End Function





Public Sub savestring(hKey As Long, strpath As String, strvalue As String, strdata As String)
    Dim keyhand&
    R = RegCreateKey(hKey, strpath, keyhand&)
    R = RegSetValueEx(keyhand&, strvalue, 0, REG_SZ, ByVal strdata, Len(strdata))
    R = RegCloseKey(keyhand&)
End Sub


Public Sub Delstring(hKey As Long, strpath As String, sKey As String)
    Dim keyhand&
    R = RegOpenKey(hKey, strpath, keyhand&)
    R = RegDeleteValue(keyhand&, sKey)
    R = RegCloseKey(keyhand&)
End Sub


Public Sub SaveSet(AppName As String, Section As String, Key As Variant, Value As Variant)
    savestring HKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & AppName & "\" & Section, CStr(Key), CStr(Value)
End Sub


Public Function GetSet(AppName As String, Section As String, Key As Variant, Optional Default As Variant) As Variant
    GetSet = GetString(HKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & AppName & "\" & Section, CStr(Key))
    If GetSet = "" Then GetSet = Default
End Function


Public Function DelSet(AppName As String, Section As String, Key As Variant) As Variant
    Delstring HKEY_CURRENT_USER, "Software\" & App.CompanyName & "\" & AppName & "\" & Section, CStr(Key)
End Function


