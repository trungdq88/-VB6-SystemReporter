Attribute VB_Name = "modRegistry"
Option Explicit


Public Enum RegistryKeys
  HKEY_CLASSES_ROOT = &H80000000
  HKEY_CURRENT_USER = &H80000001
  HKEY_LOCAL_MACHINE = &H80000002
  HKEY_USERS = &H80000003
  HKEY_CURRENT_CONFIG = &H80000005
  HKEY_DYN_DATA = &H80000006
End Enum

Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1
Public Const REG_DWORD = 4

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long





Public Function DeleteValue(ByVal hKey As RegistryKeys, ByVal strPath As String, ByVal strValue As String)
On Error Resume Next

  Dim KeyHand As Long
  
  RegOpenKey hKey, strPath, KeyHand
  RegDeleteValue KeyHand, strValue
  RegCloseKey KeyHand

End Function

Public Function GetString(ByVal hKey As RegistryKeys, ByVal strPath As String, ByVal strValue As String) As String
On Error Resume Next

  Dim KeyHand As Long
  Dim datatype As Long
  Dim lResult As Long
  Dim strBuf As String
  Dim lDataBufSize As Long
  Dim intZeroPos As Integer
  Dim lValueType As Long
  
  RegOpenKey hKey, strPath, KeyHand
  lResult = RegQueryValueEx(KeyHand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
  If lValueType = REG_SZ Then
    strBuf = String(lDataBufSize, " ")
    lResult = RegQueryValueEx(KeyHand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
      intZeroPos = InStr(strBuf, Chr(0))
      If intZeroPos > 0 Then
        GetString = Left(strBuf, intZeroPos - 1)
      Else
        GetString = strBuf
      End If
    End If
  End If
    
End Function

Public Sub SaveString(ByVal hKey As RegistryKeys, ByVal strPath As String, ByVal strValue As String, ByVal strData As String)
On Error Resume Next

  Dim KeyHand As Long
  
  RegCreateKey hKey, strPath, KeyHand
  RegSetValueEx KeyHand, strValue, 0, REG_SZ, ByVal strData, Len(strData)
  RegCloseKey KeyHand

End Sub



