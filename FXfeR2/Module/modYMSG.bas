Attribute VB_Name = "modYMSG"
Option Explicit
Public YmsgID As String, MyCookie As String

Private Function YMSGPacket(YMSGPacketData As String, YMSGStatus As String, YMSGKey As String, YMSGCommand As Integer) As String
YMSGPacket = "YMSG" & Chr(0) & Chr(15) & String(2, 0) & Chr(Int(Len(YMSGPacketData) / 256)) & Chr(Int(Len(YMSGPacketData) Mod 256)) & Chr(0) & Chr(YMSGCommand) & YMSGStatus & YMSGKey & YMSGPacketData
End Function

Public Function LoginYMSG(YID As String, Y_Hash As String, Y_Cookie As String, T_Cookie As String, Optional Invisible As Boolean = False) As String
Dim InVTyp As String
If Invisible = True Then
InVTyp = Chr(&HC)
Else
InVTyp = Chr(0)
End If
LoginYMSG = YMSGPacket("277À€" & Y_Cookie & "À€278À€" & T_Cookie & "À€307À€" & Y_Hash & "À€0À€" & YID & "À€2À€" & YID & "À€1À€" & YID & "À€192À€0À€2À€0À€244À€4194239À€135À€ym8.1.0.209À€148À€300À€", String(3, 0) & InVTyp, YmsgID, 84)
End Function

Public Function GetChallenge(YID As String) As String
GetChallenge = YMSGPacket("1À€" & YID & "À€", String(4, 0), String(4, 0), 87)
End Function

Public Function SendPm(YID As String, TargetUser As String, PmTxt As String) As String
SendPm = YMSGPacket("1À€" & YID & "À€5À€" & TargetUser & "À€14À€" & PmTxt & "À€", String(4, 0), YmsgID, 6)
End Function

Public Function PollCapacity() As String
PollCapacity = YMSGPacket("", String(4, 0), String(4, 0), 76)
End Function

Public Function AcceptFile(YID As String, TargetUser As String, YmsgFileToken As String) As String
AcceptFile = YMSGPacket("1À€" & YID & "À€5À€" & TargetUser & "À€265À€" & YmsgFileToken & "À€222À€3À€", String(4, 0), YmsgID, 220)
End Function

Public Function AcceptFileSession(YID As String, TargetUser As String, YmsgFileToken As String, Filename As String, Filesession As String) As String
AcceptFileSession = YMSGPacket("1À€" & YID & "À€5À€" & TargetUser & "À€265À€" & YmsgFileToken & "À€27À€" & Filename & "À€249À€3À€251À€" & Filesession & "À€", String(4, 0), YmsgID, 222)
End Function

Public Function SendFile(YID As String, TargetUser As String, YmsgFileToken As String, Filename As String, Filesize As Long) As String
SendFile = YMSGPacket("1À€" & YID & "À€5À€" & TargetUser & "À€265À€" & YmsgFileToken & "À€222À€1À€266À€1À€302À€268À€300À€268À€27À€" & Filename & "À€28À€" & Filesize & "À€301À€268À€303À€268À€", String(4, 0), YmsgID, 220)
End Function
'

Public Function SendFileSession(YID As String, TargetUser As String, YmsgFileToken As String, Filename As String, FileHostIP As String) As String
SendFileSession = YMSGPacket("1À€" & YID & "À€5À€" & TargetUser & "À€265À€" & YmsgFileToken & "À€27À€" & Filename & "À€249À€3À€250À€" & FileHostIP & "À€", String(4, 0), YmsgID, 221)
End Function
