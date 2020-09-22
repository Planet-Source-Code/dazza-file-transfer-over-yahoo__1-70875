VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FILEXFER2"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Text            =   "Target User"
      Top             =   1560
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   480
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Send File"
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin XfeR2.YFT YFT1 
      Left            =   600
      Top             =   4440
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Accept File"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":003D
      TabIndex        =   5
      Text            =   "mcs.msg.yahoo.com"
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Logout"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      Caption         =   "Invisible Login?"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Password"
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Yahoo ID"
      Top             =   120
      Width           =   2055
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   480
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock WsSSL 
      Left            =   480
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents SSL1 As SSL
Attribute SSL1.VB_VarHelpID = -1
Private yChallenge As String, SendingFile As Boolean
Private YIDFrom As String, YmsgFileToken As String, Filename As String, Filesize As Long, HttpFileSession As String
Private RelayServer As String, myFileTile As String, myFilePath As String, FileToken As String, SavePath As String

Private Sub Command3_Click()
On Error Resume Next
SavePath = OpenFolder("Save File To:")
If SavePath = "" Then Exit Sub
Command3.Visible = False
SendingFile = False
Winsock1.SendData AcceptFile(Text1.Text, YIDFrom, YmsgFileToken)
End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim SendSize As Long
CD1.DialogTitle = "Select File To Send."
CD1.ShowOpen
If CD1.FileTitle = "" Then Exit Sub
If CD1.Filename = "" Then Exit Sub
Text4.Text = CD1.FileTitle
myFilePath = CD1.Filename
SendingFile = True
FileToken = YMSG15(Text1.Text, Date & Time)
SendSize = FileLen(myFilePath)
Winsock1.SendData SendFile(Text1.Text, Text3.Text, FileToken, Text4.Text, SendSize)
End Sub

Private Sub Form_Load()
On Error Resume Next
Text1.Text = GetSetting(App.EXEName, "Options", "YID", "Yahoo ID")
Text2.Text = GetSetting(App.EXEName, "Options", "YPW", "Password")
Text3.Text = GetSetting(App.EXEName, "Options", "TGT", "Target User")
End Sub

Private Sub SSL1_SSLTokenData(sToken As String)
On Error Resume Next
If sToken = "invalid" Then
Me.Caption = "invalid login"
Exit Sub
Else
Set SSL1 = New SSL
SSL1.GetAuth sToken
End If
End Sub

Private Sub SSL1_SSLAuthData(yCookie As String, tCookie As String, sCrumb As String)
On Error Resume Next
Winsock1.SendData LoginYMSG(Text1.Text, YMSG15(sCrumb, yChallenge), yCookie, tCookie, Check2.Value)
End Sub

Private Sub Command1_Click()
On Error Resume Next
Winsock1.Close
Winsock1.Connect Combo1.Text, "5050"
End Sub

Private Sub Command2_Click()
On Error Resume Next
Winsock1.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
SaveSetting App.EXEName, "Options", "YID", Text1.Text
SaveSetting App.EXEName, "Options", "YPW", Text2.Text
SaveSetting App.EXEName, "Options", "TGT", Text3.Text
End
End Sub

Private Sub Winsock1_Connect()
On Error Resume Next
Winsock1.SendData PollCapacity
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String, DataLength As Integer, TmpData As String, HeaderLength As Integer
HeaderLength = 20
With Winsock1
While .BytesReceived >= HeaderLength
Call .PeekData(Data, vbString, HeaderLength)
DataLength = (256 * Asc(Mid(Data, 9, 1)) + Asc(Mid(Data, 10, 1))) + HeaderLength
If DataLength <= .BytesReceived Then
Call .GetData(TmpData, vbString, DataLength)
If Mid(TmpData, 1, 4) = "YMSG" Then
ParseYmsg TmpData
Else
TmpData = ""
End If
Else
Exit Sub
End If
DoEvents
Wend
End With
End Sub

Private Sub ParseYmsg(Data As String)
On Error Resume Next
Dim PckTyp As Long, YChal1() As String, YChal2() As String
Dim FTDat1() As String, FTDat2() As String, FTDat3() As String, FTDat4() As String, FTDat5() As String, FTDat6() As String, FTDat7() As String, FTDat8() As String, FTDat9() As String, FTDat10() As String
PckTyp = (256 * Asc(Mid$(Data, 11, 1)) + Asc(Mid$(Data, 12, 1)))
Select Case PckTyp
Case Is = 76
Winsock1.SendData GetChallenge(Text1.Text)
Case Is = 85
Me.Caption = "Logged in"
If InStr(1, Data, "Y" & Chr(9) & "v=") Then
Dim CookieStart As Integer
Dim CookieEnd As Integer
Dim datatype As Integer
Dim Cookie1Start As String, Cookie2Start As String
Dim Cookie1End As String, Cookie2End As String
Dim Cookie1 As String, Cookie2 As String
Cookie1Start = InStr(1, Data, "Y" & Chr(9) & "v=")
Cookie1End = InStr(Cookie1Start + 1, Data, "À€") + 1
Cookie1 = Mid$(Data, Cookie1Start, Cookie1End - Cookie1Start)
Cookie2Start = InStr(1, Data, "T" & Chr(9) & "z=")
Cookie2End = InStr(Cookie2Start + 1, Data, "À€")
Cookie2 = Mid$(Data, Cookie2Start, Cookie2End - Cookie2Start)
MyCookie = Cookie1 & " " & Cookie2
MyCookie = Replace(MyCookie, Chr(9), "=")
MyCookie = Replace(MyCookie, "À", ";")
End If
YmsgID = Mid(Data, 17, 4)
Case Is = 84
If Mid(Data, 13, 4) = String(4, 255) Then Me.Caption = "Invalid Login"
Case Is = 87
YmsgID = Mid(Data, 17, 4)
YChal1 = Split(Data, "À€94À€")
YChal2 = Split(YChal1(1), "À€")
yChallenge = YChal2(0)
Set SSL1 = New SSL
SSL1.GetToken Text1.Text, Text2.Text, yChallenge
Case Is = 220
FTDat1 = Split(Data, "4À€")
FTDat2 = Split(FTDat1(1), "À€")
YIDFrom = FTDat2(0)
FTDat3 = Split(Data, "À€265À€")
FTDat4 = Split(FTDat3(1), "À€")
YmsgFileToken = FTDat4(0)
FTDat5 = Split(Data, "À€27À€")
FTDat6 = Split(FTDat5(1), "À€")
Filename = FTDat6(0)
FTDat7 = Split(Data, "À€28À€")
FTDat8 = Split(FTDat7(1), "À€")
Filesize = FTDat8(0)
If SendingFile = True Then
Winsock1.SendData SendFileSession(Text1.Text, Text3.Text, FileToken, Text4.Text, "66.94.230.124")
Else
Command3.Visible = True
Label1.Caption = YIDFrom & " is sending you " & Filename & " - " & FormatFileSize(Filesize)
End If
Case Is = 221
FTDat1 = Split(Data, "4À€")
FTDat2 = Split(FTDat1(1), "À€")
YIDFrom = FTDat2(0)
FTDat3 = Split(Data, "À€265À€")
FTDat4 = Split(FTDat3(1), "À€")
YmsgFileToken = FTDat4(0)
FTDat5 = Split(Data, "À€27À€")
FTDat6 = Split(FTDat5(1), "À€")
Filename = FTDat6(0)
FTDat7 = Split(Data, "À€251À€")
FTDat8 = Split(FTDat7(1), "À€")
HttpFileSession = FTDat8(0)
FTDat9 = Split(Data, "À€250À€")
FTDat10 = Split(FTDat9(1), "À€")
RelayServer = FTDat10(0)
If SendingFile = False Then
YFT1.RecvFile Text1.Text, YIDFrom, Replace(HttpFileSession, Chr(2), "%02"), SavePath & "\" & Filename, RelayServer
End If
Case Is = 222
FTDat1 = Split(Data, "4À€")
FTDat2 = Split(FTDat1(1), "À€")
YIDFrom = FTDat2(0)
FTDat3 = Split(Data, "À€265À€")
FTDat4 = Split(FTDat3(1), "À€")
YmsgFileToken = FTDat4(0)
FTDat5 = Split(Data, "À€27À€")
FTDat6 = Split(FTDat5(1), "À€")
Filename = FTDat6(0)
FTDat7 = Split(Data, "À€251À€")
FTDat8 = Split(FTDat7(1), "À€")
HttpFileSession = FTDat8(0)
FTDat9 = Split(Data, "À€250À€")
FTDat10 = Split(FTDat9(1), "À€")
RelayServer = FTDat10(0)
If SendingFile = True Then
YFT1.SendFile Text1.Text, YIDFrom, Replace(HttpFileSession, Chr(2), "%02"), myFilePath, "66.94.230.124"
End If
End Select
End Sub

Private Sub YFT1_FileReceived()
On Error Resume Next
Label1.Caption = "File Downloaded."
PB1.Value = 0
End Sub

Private Sub YFT1_FileRecvConnected()
On Error Resume Next
Winsock1.SendData AcceptFileSession(Text1.Text, YIDFrom, YmsgFileToken, Filename, HttpFileSession)
End Sub

Private Sub YFT1_FileRecvError()
On Error Resume Next
Label1.Caption = "File RECV Error"
End Sub

Private Sub YFT1_FileRecvProgress(BytesRecv As Long, bytesTotal As Long)
On Error Resume Next
PB1.Max = bytesTotal
PB1.Value = BytesRecv
Label1.Caption = "Downloading: " & FormatFileSize(BytesRecv) & " / " & FormatFileSize(bytesTotal)
End Sub

Private Sub YFT1_FileSendBufferProgress(BytesSent As Long, bytesTotal As Long)
PB1.Max = bytesTotal
PB1.Value = BytesSent
Label1.Caption = "Loading Buffer: " & FormatFileSize(BytesSent) & " / " & FormatFileSize(bytesTotal)
End Sub

Private Sub YFT1_FileSendError()
On Error Resume Next
Label1.Caption = "File SEND Error"
SendingFile = False
PB1.Value = 0
End Sub

Private Sub YFT1_FileSendProgress(BytesSent As Long, bytesTotal As Long)
On Error Resume Next
PB1.Max = bytesTotal
PB1.Value = BytesSent
Label1.Caption = "Uploading: " & FormatFileSize(BytesSent) & " / " & FormatFileSize(bytesTotal)
End Sub

Private Sub YFT1_FileSent()
On Error Resume Next
Label1.Caption = "File Sent."
SendingFile = False
PB1.Value = 0
End Sub
