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
LoginYMSG = YMSGPacket("277��" & Y_Cookie & "��278��" & T_Cookie & "��307��" & Y_Hash & "��0��" & YID & "��2��" & YID & "��1��" & YID & "��192��0��2��0��244��4194239��135��ym8.1.0.209��148��300��", String(3, 0) & InVTyp, YmsgID, 84)
End Function

Public Function GetChallenge(YID As String) As String
GetChallenge = YMSGPacket("1��" & YID & "��", String(4, 0), String(4, 0), 87)
End Function

Public Function SendPm(YID As String, TargetUser As String, PmTxt As String) As String
SendPm = YMSGPacket("1��" & YID & "��5��" & TargetUser & "��14��" & PmTxt & "��", String(4, 0), YmsgID, 6)
End Function

Public Function PollCapacity() As String
PollCapacity = YMSGPacket("", String(4, 0), String(4, 0), 76)
End Function

Public Function AcceptFile(YID As String, TargetUser As String, YmsgFileToken As String) As String
AcceptFile = YMSGPacket("1��" & YID & "��5��" & TargetUser & "��265��" & YmsgFileToken & "��222��3��", String(4, 0), YmsgID, 220)
End Function

Public Function AcceptFileSession(YID As String, TargetUser As String, YmsgFileToken As String, Filename As String, Filesession As String) As String
AcceptFileSession = YMSGPacket("1��" & YID & "��5��" & TargetUser & "��265��" & YmsgFileToken & "��27��" & Filename & "��249��3��251��" & Filesession & "��", String(4, 0), YmsgID, 222)
End Function

Public Function SendFile(YID As String, TargetUser As String, YmsgFileToken As String, Filename As String, Filesize As Long) As String
SendFile = YMSGPacket("1��" & YID & "��5��" & TargetUser & "��265��" & YmsgFileToken & "��222��1��266��1��302��268��300��268��27��" & Filename & "��28��" & Filesize & "��301��268��303��268��", String(4, 0), YmsgID, 220)
End Function
'

Public Function SendFileSession(YID As String, TargetUser As String, YmsgFileToken As String, Filename As String, FileHostIP As String) As String
SendFileSession = YMSGPacket("1��" & YID & "��5��" & TargetUser & "��265��" & YmsgFileToken & "��27��" & Filename & "��249��3��250��" & FileHostIP & "��", String(4, 0), YmsgID, 221)
End Function
