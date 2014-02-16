Attribute VB_Name = "mMain"
Sub Main()
Shell "regsvr32 /s UniControls_v2.0.ocx"
frmMain.Show

End Sub

'Public Function FileExists(sFile As String) As Boolean
'On Error Resume Next
'FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
'End Function

