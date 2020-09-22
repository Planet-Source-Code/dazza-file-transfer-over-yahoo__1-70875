Attribute VB_Name = "modApi"
Option Explicit
Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" (lpBI As BrowseInfo) As Long
Private Declare Function lstrcat Lib "kernel32.dll" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function StrFormatByteSize Lib "shlwapi.dll" Alias "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByRef cchBuf As Long) As String
Private Type BrowseInfo
hWndOwner As Long
pIDLRoot As Long
pszDisplayName As Long
lpszTitle As Long
ulFlags As Long
lpfnCallback As Long
lParam As Long
iImage As Long
End Type

Public Function OpenFolder(Prompt As String) As String
On Error Resume Next
Dim n As Integer
Dim IDList As Long
Dim Result As Long
Dim ThePath As String
Dim BI As BrowseInfo
With BI
.hWndOwner = GetActiveWindow()
.lpszTitle = lstrcat(Prompt, "")
.ulFlags = &H1
End With
IDList = SHBrowseForFolder(BI)
If IDList Then
ThePath = String$(260&, 0)
Result = SHGetPathFromIDList(IDList, ThePath)
Call CoTaskMemFree(IDList)
n = InStr(ThePath, vbNullChar)
If n Then ThePath = Left$(ThePath, n - 1)
End If
OpenFolder = ThePath
End Function

Public Function FormatFileSize(ByVal Amount As Long) As String
On Error Resume Next
Dim Buffer As String
Dim Result As String
Buffer = Space$(255)
Result = StrFormatByteSize(Amount, Buffer, Len(Buffer))
If InStr(Result, vbNullChar) > 1 Then FormatFileSize = Left$(Result, InStr(Result, vbNullChar) - 1)
End Function
