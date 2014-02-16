Attribute VB_Name = "modFile"
Public Function WriteFileUni(FileName As String, Unistr As String)
Dim FSO As Object 'tao 1 file mo'i rôi mo'i ghi vào
      Set FSO = CreateObject("Scripting.FileSystemObject").CreateTextFile(FileName, True)
      Set FSO = Nothing
      Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(FileName, 2, , -1)
          FSO.Write Unistr
      Set FSO = Nothing
End Function
Public Function ReadFileUni(FileName As String) As String
    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject").OpenTextFile(FileName, 1, , -2)
    ReadFileUni = FSO.Readall
    Set FSO = Nothing
End Function
Public Function FileExists(sFile As String) As Boolean
On Error Resume Next
FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function

