VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl YFT 
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1050
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   345
   ScaleWidth      =   1050
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   840
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   240
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "YFT 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "YFT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event FileSendError()
Public Event FileSent()
Public Event FileSendProgress(BytesSent As Long, bytesTotal As Long)
Public Event FileSendBufferProgress(BytesSent As Long, bytesTotal As Long)
Public Event FileRecvError()
Public Event FileReceived()
Public Event FileRecvProgress(BytesRecv As Long, bytesTotal As Long)
Public Event FileRecvConnected()
Private MyYID As String, MyTarget As String, MyFileToken As String, myFilePath As String, MyDestFile As String, MyHostip As String, MyFileLen As Long
Private FileContent As Long, FirstPckt As Boolean, ContLen As Long, Fhandle1 As Long, BeginCounter As Boolean, TmpCnt As Long

Private Function FTGetfile(Sender As String, Recver As String, Token As String, yCookie As String, FtServer As String) As String
On Error Resume Next
Dim StrFT As String
StrFT = "GET /relay?token=" & Token & "&sender=" & Sender & "&recver=" & Recver & " HTTP/1.1" & vbCrLf
StrFT = StrFT & "Connection: Close" & vbCrLf
StrFT = StrFT & "Cookie: " & yCookie & vbCrLf
StrFT = StrFT & "Host: " & FtServer & vbCrLf
StrFT = StrFT & "User-Agent: Mozilla/5.0" & vbCrLf
StrFT = StrFT & "Cache-Control: no-cache" & vbCrLf & vbCrLf
FTGetfile = StrFT
End Function

Private Sub FTSetFile(Sender As String, Recver As String, Token As String, yCookie As String, FtServer As String, FilePath As String)
On Error Resume Next
Dim StrFT As String, Data
BeginCounter = False
MyFileLen = FileLen(FilePath)
StrFT = "POST /relay?token=" & Token & "&sender=" & Sender & "&recver=" & Recver & " HTTP/1.1" & vbCrLf
StrFT = StrFT & "Cache-Control: no-cache" & vbCrLf
StrFT = StrFT & "Cookie: " & yCookie & vbCrLf
StrFT = StrFT & "Host: " & FtServer & vbCrLf
StrFT = StrFT & "Content-Length: " & MyFileLen & vbCrLf
StrFT = StrFT & "User-Agent: Mozilla/5.0" & vbCrLf
StrFT = StrFT & "Connection: Close" & vbCrLf & vbCrLf
Winsock2.SendData StrFT
DoEvents
TmpCnt = 0
Open FilePath For Binary As #1
Do While Not EOF(1)
Data = Input(8192, #1)
If Winsock2.State = sckConnected Then
Winsock2.SendData Data
TmpCnt = TmpCnt + Len(Data)
RaiseEvent FileSendBufferProgress(TmpCnt, MyFileLen)
End If
DoEvents
Loop
BeginCounter = True
End Sub

Public Sub SendFile(Username As String, UserTo As String, SessionKey As String, FilePath As String, Hostip As String)
On Error Resume Next
MyTarget = UserTo
MyFileToken = SessionKey
myFilePath = FilePath
MyYID = Username
MyHostip = Hostip
Winsock2.Close
Winsock2.RemoteHost = MyHostip
Winsock2.RemotePort = "80"
Winsock2.Connect
End Sub

Public Sub RecvFile(Username As String, UserTo As String, SessionKey As String, DestFile As String, Hostip As String)
On Error Resume Next
MyTarget = UserTo
MyFileToken = SessionKey
MyDestFile = DestFile
MyYID = Username
MyHostip = Hostip
FirstPckt = True
ContLen = 0
Winsock3.Close
Winsock3.RemoteHost = MyHostip
Winsock3.RemotePort = "80"
Winsock3.Connect
End Sub

Private Sub Winsock2_Connect()
On Error Resume Next
Call FTSetFile(MyYID, MyTarget, MyFileToken, MyCookie, MyHostip, myFilePath)
End Sub

Private Sub Winsock2_SendComplete()
On Error Resume Next
If BeginCounter = False Then Exit Sub
RaiseEvent FileSent
End Sub

Private Sub Winsock2_SendProgress(ByVal BytesSent As Long, ByVal BytesRemaining As Long)
On Error Resume Next
If BeginCounter = False Then Exit Sub
RaiseEvent FileSendProgress(MyFileLen - BytesRemaining, MyFileLen)
End Sub

Private Sub Winsock3_Connect()
On Error Resume Next
Winsock3.SendData FTGetfile(MyYID, MyTarget, MyFileToken, MyCookie, MyHostip)
RaiseEvent FileRecvConnected
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String, HeadStr() As String, HeadLen As Integer
Winsock2.GetData Data
HeadStr = Split(Data, vbCrLf & vbCrLf)
HeadLen = Len(HeadStr(0))
If InStr(1, HeadStr(0), "404 Not Found") Then Winsock2.Close: RaiseEvent FileSendError: Exit Sub
If InStr(1, HeadStr(0), "500 Service Error") Then Winsock2.Close: RaiseEvent FileSendError: Exit Sub
End Sub

Private Sub Winsock3_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String, HeadStr() As String, HeadLen As Integer, DatL1() As String, DatL2() As String
Winsock3.GetData Data
If FirstPckt = True Then
FirstPckt = False
HeadStr = Split(Data, vbCrLf & vbCrLf)
HeadLen = Len(HeadStr(0))
If InStr(1, HeadStr(0), "404 Not Found") Then Winsock3.Close: RaiseEvent FileRecvError: Exit Sub
If InStr(1, HeadStr(0), "500 Service Error") Then Winsock3.Close: RaiseEvent FileRecvError: Exit Sub
DatL1 = Split(HeadStr(0), "Content-Length: ")
DatL2 = Split(DatL1(1), vbCrLf)
ContLen = DatL2(0)
FileContent = Mid(Data, HeadLen + 5)
Kill MyDestFile
Fhandle1 = FreeFile
Open MyDestFile For Binary Access Write As #Fhandle1
Else
Put #Fhandle1, , Data
FileContent = FileContent + Len(Data)
RaiseEvent FileRecvProgress(FileContent, ContLen)
If FileContent = ContLen Then
Winsock3.Close
Close #Fhandle1
RaiseEvent FileReceived
End If
End If
End Sub

Private Sub Winsock2_Close()
On Error Resume Next
Winsock2.Close
End Sub

Private Sub Winsock3_Close()
On Error Resume Next
Winsock3.Close
End Sub

Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Winsock2.Close
End Sub

Private Sub Winsock3_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Winsock3.Close
End Sub
