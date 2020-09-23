Attribute VB_Name = "modReg"
Public Declare Function GetShortPathName Lib "Kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Const REG_SZ = 1
Public Const REG_DWORD = 4

Public Sub savekey(Hkey As Long, strPath As String)
    Dim keyhand&
    r = RegCreateKey(Hkey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)
End Sub

Public Function GetShortFileName(ByVal sFileName As String) As String
On Error Resume Next
Dim lLen As Long, sPath As String

    Const MAX_LEN& = 164

sPath$ = String$(MAX_LEN + 1, 0)
lLen& = GetShortPathName(sFileName$, sPath$, MAX_LEN)

    GetShortFileName$ = Left$(sPath$, lLen&)

End Function

Public Function getstring(Hkey As Long, strPath As String, strValue As String)
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    r = RegOpenKey(Hkey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)


    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)


        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))


            If intZeroPos > 0 Then
                getstring = Left$(strBuf, intZeroPos - 1)
            Else
                getstring = strBuf
            End If
        End If
    End If
End Function


Public Sub savestring(Hkey As Long, strPath As String, strValue As String, StrData As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(Hkey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal StrData, Len(StrData))
    r = RegCloseKey(keyhand)
End Sub


Function getdword(ByVal Hkey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    Dim r As Long
    Dim keyhand As Long
    r = RegOpenKey(Hkey, strPath, keyhand)
    lDataBufSize = 4
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)


    If lResult = ERROR_SUCCESS Then


        If lValueType = REG_DWORD Then
            getdword = lBuf
        End If
    End If
    r = RegCloseKey(keyhand)
End Function


Function SaveDword(ByVal Hkey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(Hkey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    r = RegCloseKey(keyhand)
End Function


Public Function DeleteKey(ByVal Hkey As Long, ByVal strKey As String)
    Dim r As Long
    r = RegDeleteKey(Hkey, strKey)
End Function


Public Function DeleteValue(ByVal Hkey As Long, ByVal strPath As String, ByVal strValue As String)
    Dim keyhand As Long
    r = RegOpenKey(Hkey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function

Public Sub assFile(sProgram As String, sIcon As String, sFile As String, sDescription As String, Optional ByVal sCon As String = "text/plain")
    Call savestring(HKEY_CLASSES_ROOT, "\." & sFile$, "", sFile$ & "file")
    'content type
    Call savestring(HKEY_CLASSES_ROOT, "\." & sFile$, "Content Type", sCon$)
    'name
    Call savestring(HKEY_CLASSES_ROOT, "\" & sFile$ & "file", "", sDescription$)
    'edit flags
    Call SaveDword(HKEY_CLASSES_ROOT, "\" & sFile$ & "file", "EditFlags", "0000")
    'file's icon (can be an icon file, or an
    '     icon located within a dll file)
    Call savestring(HKEY_CLASSES_ROOT, "\" & sFile$ & "file\DefaultIcon", "", GetShortFileName$(sIcon$))
    'Shell
    Call savestring(HKEY_CLASSES_ROOT, "\" & sFile$ & "file\Shell", "", "")
    'Shell Open
    Call savestring(HKEY_CLASSES_ROOT, "\" & sFile$ & "file\Shell\Open", "", "")
    'Shell open command
    Call savestring(HKEY_CLASSES_ROOT, "\" & sFile$ & "file\Shell\Open\command", "", GetShortFileName$(sProgram$) & " %1")
End Sub
