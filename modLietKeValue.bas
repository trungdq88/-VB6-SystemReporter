Attribute VB_Name = "modLietKeValue"
Option Explicit
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_CURRENT_USER = &H80000001
Private Const KEY_ALL_ACCESS = &HF003F
Private Const REG_SZ = 1
Private Const REG_BINARY = 3                     ' Free form binary
Private Const REG_DWORD = 4                      ' 32-bit number
Private Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Private Const REG_MULTI_SZ = 7                   ' Multiple Unicode strings
Dim RetVal As Long
Dim hKey As Long
Dim NameKey As String
Dim lpType As Long
Dim LenName As Long
Dim Data(0 To 255) As Byte
Dim DataLen As Long
Dim DataString As String
Dim index As Long
Dim i As Long
Dim KetQua As String
Public xTotalStartUp
Public Function GetKeyValue(FullKeyName)
frmMain.txt.Text = frmMain.txt.Text & " [" & FullKeyName & "]" & vbCrLf
xTotalStartUp = 0
Dim Key1, Key2, i, Ua
Ua = 10
DoEvents
For i = 1 To Len(FullKeyName)
DoEvents
    If Mid(FullKeyName, i, 1) = "\" Then
        Ua = Ua + 10
        If Ua = 20 Then
DoEvents
            Key1 = Left(FullKeyName, i - 1)
            Key2 = Right(FullKeyName, Len(FullKeyName) - i)
        End If
    End If
Next i
'frmMain.Cls
If Key1 = "HKEY_LOCAL_MACHINE" Then
DoEvents
RetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, Key2, 0, KEY_ALL_ACCESS, hKey)
ElseIf Key1 = "HKEY_CURRENT_USER" Then
RetVal = RegOpenKeyEx(HKEY_CURRENT_USER, Key2, 0, KEY_ALL_ACCESS, hKey)
End If

index = 0
Do While RetVal = 0
    NameKey = Space(255)
    DataString = Space(255)
    LenName = 255
    DataLen = 255
DoEvents
    RetVal = RegEnumValue(hKey, index, NameKey, LenName, ByVal 0, lpType, Data(0), DataLen)
    If RetVal = 0 Then
DoEvents
        NameKey = Left(NameKey, LenName) 'Rút b? kho?n tr?ng th?a
        DataString = ""
' X? lý thông tin theo ki?u c?a nó và ??a vào bi?n DataString
        Select Case lpType
             Case REG_SZ
                For i = 0 To DataLen - 1
                    DataString = DataString & Chr(Data(i)) ' N?i các ch? cái thành chu?i
                Next
             Case REG_BINARY
DoEvents
                For i = 0 To DataLen - 1
                    Dim temp As String
                    temp = Hex(Data(i))
                    If Len(temp) < 2 Then temp = String(2 - Len(temp), "0") & temp
                    DataString = DataString & temp & " "
 ' N?i các c?p s? nh? phân l?i v?i nhau
                Next
            Case REG_DWORD
DoEvents
                For i = DataLen - 1 To 0 Step -1
                    DataString = DataString & Hex(Data(i)) 'N?i các sô hexa v?i nhau
                Next
            Case REG_MULTI_SZ
DoEvents
                For i = 0 To DataLen - 1
                    DataString = DataString & Chr(Data(i))
    'N?i các ký t? bao g?m ký t? vbNullChar (?? cách dòng) thành m?t chu?i, b?n có th? s? d?ng m?t m?ng g?m nhi?u string thay vì là m?t
                Next
            Case REG_EXPAND_SZ
DoEvents
                For i = 0 To DataLen - 2
                    DataString = DataString & Chr(Data(i))
    'N?i các ký t? l?i v?i nhau, b? ký t? NULL cu?i cùng
                Next
            Case Else
DoEvents
                DataString = " ??? "
        ' Trên ?ây là 5 ki?u có trên WinXP
        End Select
    End If
    If Left(Left(NameKey, LenName), 1) <> " " Then
DoEvents
    '///////////////////
    'Form1.List1.AddItem DataString
    Dim sX As String
    sX = Left(NameKey, LenName) & " = " & DataString

DoEvents
    frmMain.txt.Text = frmMain.txt.Text & "    [+] " & sX
    frmMain.txt.Text = frmMain.txt.Text & vbCrLf
    
    '///////////////
    End If
    index = index + 1
    'frmMain.Print Left(NameKey, LenName) & "=" & DataString
Loop
RetVal = RegCloseKey(hKey)
DoEvents

frmMain.txt.Text = frmMain.txt.Text & vbCrLf
End Function

Public Function GetFileName(ByVal sPath As String) As String
GetFileName = Mid(sPath, InStrRev(sPath, "\") + 1)
End Function
Public Function GetFolderPath(ByVal sPath As String) As String
GetFolderPath = Left(sPath, InStrRev(sPath, "\") - 1)
End Function

Public Sub GetFolderStartUp()
With frmMain
    Dim j


    frmMain.txt.Text = frmMain.txt.Text & " [C:\Documents and Settings\All Users\Start Menu\Programs\Startup]" & vbCrLf
    .File1.Path = "C:\Documents and Settings\All Users\Start Menu\Programs\Startup"
    For j = 0 To .File1.ListCount - 1
        frmMain.txt.Text = frmMain.txt.Text & "    [+] " & .File1.List(j) & vbCrLf
    Next j
    frmMain.txt.Text = frmMain.txt.Text & vbCrLf
    frmMain.txt.Text = frmMain.txt.Text & " [C:\Documents and Settings\" & Environ$("USERNAME") & "\Start Menu\Programs\Startup]" & vbCrLf
    .File1.Path = "C:\Documents and Settings\" & Environ$("USERNAME") & "\Start Menu\Programs\Startup"
    For j = 0 To .File1.ListCount - 1
        frmMain.txt.Text = frmMain.txt.Text & "    [+] " & .File1.List(j) & vbCrLf
    Next j
    frmMain.txt.Text = frmMain.txt.Text & vbCrLf
End With
End Sub









