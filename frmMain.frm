VERSION 5.00
Object = "{D1ECD258-D329-4388-AB83-DEC261A66B86}#1.0#0"; "UniControls_v2.0.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perfect System Reporter"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11205
   FillColor       =   &H00FFC0C0&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFC0C0&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   11205
   StartUpPosition =   2  'CenterScreen
   Begin UniControls.UniCommonDialog Dialog1 
      Left            =   8040
      Top             =   480
      _ExtentX        =   714
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin UniControls.UniButton cmdSave 
      Height          =   375
      Left            =   9600
      TabIndex        =   4
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Icon            =   "frmMain.frx":27A2
      Style           =   2
      Caption         =   "Lu7u la5i"
      IconAlign       =   3
      iNonThemeStyle  =   2
      Enabled         =   0   'False
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
      ShowFocusRectangle=   0   'False
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   5880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6960
      Top             =   1200
   End
   Begin UniControls.ProgressBar Bar1 
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   16711680
      Scrolling       =   2
      Value           =   0
   End
   Begin UniControls.UniLabel UniLabel2 
      Height          =   255
      Left            =   0
      Top             =   720
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "Nha61n va2o nu1t 'Ba81t D9a62u Kie63m Tra'"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   33023
   End
   Begin UniControls.UniButton cmdStart 
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   960
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      Icon            =   "frmMain.frx":27BE
      Style           =   2
      Caption         =   "Ba81t D9a62u Kie63m Tra"
      IconAlign       =   3
      iNonThemeStyle  =   2
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedBordersByTheme=   0   'False
      ShowFocusRectangle=   0   'False
   End
   Begin UniControls.UniTextBox txt 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   10610
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Text            =   ""
      MultiLine       =   -1  'True
      Locked          =   -1  'True
      BorderStyle     =   2
      Scrollbar       =   3
   End
   Begin UniControls.UniLabel UniLabel1 
      Height          =   375
      Left            =   120
      Top             =   240
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   661
      Alignment       =   1
      Caption         =   "Kie63m tra he65 tho61ng"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Dim sConnType As String * 255
Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
Private memInfo As MEMORYSTATUS
Dim memoryInfo As MEMORYSTATUS
Dim lastpcent As Single, lastTot As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" _
   (lpBuffer As MEMORYSTATUS)
Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Sub cmdSave_Click()
Dialog1.FileName = ""
Dialog1.Filter = "Text File (*.txt)|*.txt|"
Dialog1.ShowSave
Dim sPa As String
sPa = Dialog1.FileName
If sPa <> "" Then
    If Right(sPa, 3) <> ".txt" Then sPa = sPa & ".txt"
    WriteFileUni sPa, txt.Text
    UniMsgBox "D9a4 lu7u xong!", vbOKOnly + vbInformation, "OK"
    Shell "notepad " & ChrW(34) & sPa & ChrW(34), vbNormalFocus
End If
End Sub

Private Sub cmdStart_Click()
Bar1.Value = 0
Timer1.Enabled = True
txt.Visible = False
cmdStart.Enabled = False
cmdSave.Enabled = False
CheckNow
End Sub




Private Sub Form_Load()

File1.Archive = True
File1.System = True
File1.Hidden = True
File1.ReadOnly = True
End Sub

Function GetOS()
Dim strComputer, strWMIOS
strComputer = "."
Dim objWmiService: Set objWmiService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Dim strOsQuery: strOsQuery = "Select * from Win32_OperatingSystem"
Dim colOperatingSystems: Set colOperatingSystems = objWmiService.ExecQuery(strOsQuery)
Dim objOs
Dim strOsVer

    For Each objOs In colOperatingSystems
        strWMIOS = objOs.Caption & " " & objOs.Version
    Next
GetOS = strWMIOS
End Function

Private Sub Timer1_Timer()

Bar1.Value = Bar1.Value + 3
If Bar1.Value > Bar1.Max - 1 Then Bar1.Value = 0

End Sub

Function GetComputer()
    Dim dwlen As Long
    Dim strString As String
    dwlen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwlen, "X")
    GetComputerName strString, dwlen
    strString = Left(strString, dwlen)
    GetComputer = strString
End Function

Public Function GetRAMTotal() As String
   Call GlobalMemoryStatus(memInfo)
        GetRAMTotal = Round(memInfo.dwTotalPhys / 1024 / 1024, 3) & " MB"
End Function
Function GetMemoryInfo()

  DoEvents
  GlobalMemoryStatus memoryInfo
    Dim Totp1
    Dim Availp1
    Dim pcent
    Dim lastpcent
    Dim lastTot
  Totp1 = Int(memoryInfo.dwTotalPhys / 1044032 * 10 + 0.5) / 10
  Availp1 = Int(memoryInfo.dwAvailPhys / 1044032 * 10 + 0.5) / 10
  pcent = Int(Availp1 / Totp1 * 100)
  
  lastpcent = pcent
  lastTot = memoryInfo.dwMemoryLoad
  
  GetMemoryInfo = Format(lastpcent)

End Function

Private Sub CheckNow()
DoEvents
txt.Text = ToUnicode("" _
& " Ma64u ba1o ca1o ti2nh tra5ng ma1y ti1nh hie65n ta5i." & vbCrLf _
& " Thu75c hie65n bo73i chu7o7ng tri2nh: Perfect Antivirus 2009." & vbCrLf _
& " Tho72i gian: " & Time & " - " & Date & vbCrLf _
& " - Tho6ng tin ma1y ti1nh:" & vbCrLf)

txt.Text = txt.Text & ToUnicode("    + He65 d9ie62u ha2nh: ") & GetOS & vbCrLf _
& ToUnicode("    + Te6n ngu7o72i su73 du5ng: ") & Environ$("username") & vbCrLf _
& ToUnicode("    + Te6n ma1y ti1nh: ") & GetComputer & vbCrLf _
& ToUnicode("    + Dung lu7o75ng bo65 nho71 RAM: ") & GetRAMTotal & vbCrLf




txt.Text = txt.Text & "===============================================================================" & vbCrLf _
& vbCrLf _
& vbCrLf _
& ToUnicode(" [1] - Ca1c chu7o7ng tri2nh d9ang cha5y trong bo65 nho71:" & vbCrLf & vbCrLf)
'/////////////////////////////////////
DoEvents
Dim ColItems
Dim ObjItem
Set ColItems = GetObject("winmgmts:\root\CIMV2").ExecQuery("SELECT * FROM Win32_Process")
DoEvents
For Each ObjItem In ColItems
   'frmMain.lblStatus.Caption = ObjItem.ExecutablePath
   If IsNull(ObjItem.ExecutablePath) = False Then txt.Text = txt.Text & " " & ObjItem.ExecutablePath & " : " & ObjItem.ProcessID & vbCrLf
DoEvents
Next
Set ColItems = Nothing
Set ObjItem = Nothing
'/////////////////////////////
GachXx
txt.Text = txt.Text & vbCrLf & vbCrLf & ToUnicode(" [2] - Ca1c chu7o7ng tri2nh d9u7o75c na5p lu1c kho73i d9o65ng:") & vbCrLf & vbCrLf
DoEvents
GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
DoEvents
GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
DoEvents
GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
DoEvents
GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
DoEvents
GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx"
DoEvents
GetFolderStartUp

txt.Text = txt.Text & vbCrLf
GachXx
txt.Text = txt.Text & vbCrLf & ToUnicode(" [3] - Gia1 tri5 cu3a ca1c Key quan tro5ng trong kho1a Winlogon:" & vbCrLf & vbCrLf & vbCrLf)
GetKeyValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Winlogon"
GachXx
DoEvents
txt.Text = txt.Text & vbCrLf & ToUnicode(" [4] - No65i dung ta65p tin Hosts:" & vbCrLf & vbCrLf & vbCrLf)
If FileExists("C:\WINDOWS\system32\drivers\etc\hosts") = False Then
    DoEvents
    txt.Text = txt.Text & ToUnicode(" (Kho6ng ti2m tha61y ta65p tin Hosts)") & vbCrLf & vbCrLf
    GachXx
Else
    DoEvents
    txt.Text = txt.Text & "-------------------------------------------------------------------" & vbCrLf & ReadFileUni("C:\WINDOWS\system32\drivers\etc\hosts") & "-------------------------------------------------------------------" & vbCrLf & vbCrLf
    GachXx
End If

txt.Text = txt.Text & vbCrLf & ToUnicode(" [5] - Ca1c tho6ng so61 ca2i d9a85t cu3a Internet Explorer:" & vbCrLf & vbCrLf & vbCrLf)
DoEvents
GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main"
DoEvents
GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Main"
DoEvents
GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Search"
GachXx
txt.Text = txt.Text & vbCrLf & ToUnicode(" [6] - Kho1a d9a8ng ky1 cu3a ca1c ta65p tin thu75c thi:" & vbCrLf & vbCrLf)
DoEvents
txt.Text = txt.Text & " [HKEY_CLASSES_ROOT\exefile\shell\open\command]" & vbCrLf _
& " (Default) = " & GetString(HKEY_CLASSES_ROOT, "exefile\shell\open\command", "") & vbCrLf & vbCrLf _
& " [HKEY_CLASSES_ROOT\comfile\shell\open\command]" & vbCrLf _
& " (Default) = " & GetString(HKEY_CLASSES_ROOT, "comfile\shell\open\command", "") & vbCrLf & vbCrLf _
& " [HKEY_CLASSES_ROOT\batfile\shell\open\command]" & vbCrLf _
& " (Default) = " & GetString(HKEY_CLASSES_ROOT, "batfile\shell\open\command", "") & vbCrLf & vbCrLf _
& " [HKEY_CLASSES_ROOT\piffile\shell\open\command]" & vbCrLf _
& " (Default) = " & GetString(HKEY_CLASSES_ROOT, "piffile\shell\open\command", "") & vbCrLf & vbCrLf
GachXx

txt.Text = txt.Text & vbCrLf & ToUnicode(" [7] - Ti2nh tra5ng Kho1a/Mo73 ca1c chu71c na8ng cu3a Windows" & vbCrLf & vbCrLf & vbCrLf)
DoEvents
GetKeyValue ("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies")

DoEvents
GetKeyValue ("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System")
DoEvents
GetKeyValue ("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer")

DoEvents
GetKeyValue ("HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\System")
txt.Text = txt.Text & vbCrLf
GachXx

txt.Text = txt.Text & vbCrLf & ToUnicode(" [8] - Ca1c ta65p tin Autorun.inf trong o63 d9i4a:" & vbCrLf & vbCrLf & vbCrLf)

'///////////////////////////////////////////////
On Error Resume Next
DoEvents
    Dim Str
    Dim str2
    Dim FSO  As New FileSystemObject
    Dim drv  As Drive
    Dim drvs As Drives
    DoEvents
    Set drvs = FSO.Drives
    DoEvents
    For Each drv In drvs
        If UCase(drv.DriveLetter) <> "A" Then
        DoEvents
            If FileExists(drv.DriveLetter & ":\autorun.inf") = False Then
                txt.Text = txt.Text & " " & drv.VolumeName & " [" & drv.DriveLetter & ":\]" & ToUnicode(" - Kho6ng pha1t hie65n Autorun.") & vbCrLf
            Else
                txt.Text = txt.Text & " " & drv.VolumeName & " [" & drv.DriveLetter & ":\]" & ToUnicode(" - Pha1t hie65n Autorun!") & vbCrLf _
                & "-------------------------------------" & vbCrLf _
                & ReadFileUni(drv.DriveLetter & ":\autorun.inf") & vbCrLf _
                & "-------------------------------------" & vbCrLf
            End If
        End If
    Next
    DoEvents
    Set FSO = Nothing
    Set drv = Nothing
    Set drvs = Nothing

    DoEvents
'//////////////////////////////////////





txt.Text = txt.Text & vbCrLf & vbCrLf & vbCrLf & "==============================================================================" & vbCrLf & vbCrLf
txt.Text = txt.Text & vbCrLf _
& ToUnicode(" Hoa2n ta61t ba1o ca1o." & vbCrLf _
& " Tu72 ba3n ba1o ca1o na2y, ba5n co1 the63 tu75 pha6n ti1ch va2 ti2m hie63u d9e63 loa5i bo3 ca1c loa5i Virus thu7o72ng ga85p tre6n ma1y." & vbCrLf _
& " Ne61u pha1t hie65n ma1y ti1nh ba5n ga85p va61n d9e62 qua File Log o73 tre6n, ba5n co1 the63 ba65t ta61t ca3 ca1c chu71c na8ng tu75 d9o65ng ba3o ve65 cu3a PAV 2009, chu7o7ng tri2nh se4 tu75 d9o65ng su73a chu74a ta61t ca3." & vbCrLf _
& " Ne61u kho6ng bie61t ca1ch pha6n ti1ch, ba5n co1 the63 gu73i File Log na2y d9e61n nhu74ng ngu7o72i co1 chuye6n mo6n nho72 ho5 gia3i d9a1p va2 d9u7a ra lo72i khuye6n cho ma1y ti1nh cu3a ba5n." & vbCrLf _
& " Hoa85c cu4ng co1 the63 Post File Log na2y le6n 1 so61 die64n d9a2n nho72 pha6n ti1ch nhu7:  http://truongton.net (Mu5c 'Be65nh Vie65n Ma1y Ti1nh'), http://virusvn.com, http://benhvientinhoc.com" & vbCrLf _
& " " & vbCrLf _
& " " & vbCrLf _
& " --------------------------------------- End --------------------------------------" & vbCrLf _
& " Copyright © Perfect Antivirus 2009")

txt.Visible = True
Timer1.Enabled = False
cmdStart.Enabled = True
cmdSave.Enabled = True
Bar1.Value = 0

UniMsgBox "D9a4 Kie63m Tra Xong!" & vbCrLf & "D9a4 copy file log va2o bo65 nho71, ba5n chi3 ca62n nha61n Ctrl + V", vbOKOnly + vbInformation, "OK"

'Software\Policies\Microsoft\Internet Explorer\Control Panel
End Sub



Private Sub GachXx()
txt.Text = txt.Text & vbCrLf & " ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf
End Sub
