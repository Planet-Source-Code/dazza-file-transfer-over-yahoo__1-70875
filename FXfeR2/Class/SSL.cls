VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SSL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents WsSSL As Winsock
Attribute WsSSL.VB_VarHelpID = -1
Public Event SSLTokenData(sToken As String)
Public Event SSLAuthData(yCookie As String, tCookie As String, sCrumb As String)
Private sFirst As Boolean
Private SSLTMP As String
Private iLayer As Integer
Private InBuffer As String
Private iSeekLen As Integer
Private sBuffer As String
Private MASTER_KEY As String
Private CLIENT_READ_KEY As String
Private CLIENT_WRITE_KEY As String
Private Private_KEY As String
Private ENCODED_CERT As String
Private CONNECTION_ID As String
Private SEND_SEQUENCE_NUMBER As Double
Private RECV_SEQUENCE_NUMBER As Double
Private CLIENT_HELLO As String
Private CHALLENGE_DATA As String
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hSessionKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, ByRef hSessionKey As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hSessionKey As Long) As Long
Private Declare Function CryptImportKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal hPubKey As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
Private Declare Function CryptExportKey Lib "advapi32.dll" (ByVal hSessionKey As Long, ByVal hExpKey As Long, ByVal dwBlobType As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long) As Long
Private Declare Function CryptGenRandom Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwLen As Long, ByVal pbBuffer As String) As Long
Private Const SERVICE_PROVIDER As String = "Microsoft Enhanced Cryptographic Provider v1.0" & vbNullChar
Private Const KEY_CONTAINER As String = "GCN SSL Container" & vbNullChar
Private Const PROV_RSA_FULL As Long = 1
Private Const CRYPT_NEWKEYSET As Long = 8
Private Const CRYPT_EXPORTABLE As Long = 1
Private Const CALG_MD5 As Long = 32771
Private Const CALG_RC4 As Long = 26625
Private Const HP_HASHVAL As Long = 2
Private Const SIMPLEBLOB As Long = 1
Private Const GEN_KEY_BITS As Long = &H800000
Private hCryptProv As Long
Private hClientWriteKey As Long
Private hClientReadKey As Long
Private hMasterKey As Long
Private lngType As Long
Private bRaisedConnected As Boolean

Private Function ExportKeyBlob(ByRef StrMasterKey As String, ByRef StrReadKey As String, ByRef StrWriteKey As String, ByVal StrChallenge As String, ByVal StrConnectionID As String, ByVal StrPrivateKey As String) As String
On Error Resume Next
Dim lngReturnValue As Long
Dim lngLength As Long
Dim rgbBlob As String
Dim hPrivateKey As Long
Call CreateKey(hMasterKey, StrMasterKey)
StrMasterKey = MD5_Hash(StrMasterKey)
Call CreateKey(hClientReadKey, StrMasterKey & "0" & StrChallenge & StrConnectionID)
Call CreateKey(hClientWriteKey, StrMasterKey & "1" & StrChallenge & StrConnectionID)
StrReadKey = MD5_Hash(StrMasterKey & "0" & StrChallenge & StrConnectionID)
StrWriteKey = MD5_Hash(StrMasterKey & "1" & StrChallenge & StrConnectionID)
lngReturnValue = CryptImportKey(hCryptProv, StrPrivateKey, Len(StrPrivateKey), 0, 0, hPrivateKey)
lngReturnValue = CryptExportKey(hMasterKey, hPrivateKey, SIMPLEBLOB, 0, vbNull, lngLength)
rgbBlob = String(lngLength, 0)
lngReturnValue = CryptExportKey(hMasterKey, hPrivateKey, SIMPLEBLOB, 0, rgbBlob, lngLength)
If hPrivateKey <> 0 Then CryptDestroyKey hPrivateKey
If hMasterKey <> 0 Then CryptDestroyKey hMasterKey
ExportKeyBlob = ReverseString(Right(rgbBlob, 128))
End Function

Private Sub CreateKey(ByRef KeyName As Long, ByVal HashData As String)
On Error Resume Next
Dim lngParams As Long
Dim lngReturnValue As Long
Dim lngHashLen As Long
Dim hHash As Long
lngReturnValue = CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not create a Hash Object (CryptCreateHash API)"
lngReturnValue = CryptHashData(hHash, HashData, Len(HashData), 0)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not calculate a Hash Value (CryptHashData API)"
lngParams = GEN_KEY_BITS Or CRYPT_EXPORTABLE
lngReturnValue = CryptDeriveKey(hCryptProv, CALG_RC4, hHash, lngParams, KeyName)
If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "Could not create a session key (CryptDeriveKey API)"
If hHash <> 0 Then CryptDestroyHash hHash
End Sub

Private Function RC4_Encrypt(ByVal Plaintext As String) As String
On Error Resume Next
Dim lngLength As Long
Dim lngReturnValue As Long
lngLength = Len(Plaintext)
lngReturnValue = CryptEncrypt(hClientWriteKey, 0, False, 0, Plaintext, lngLength, lngLength)
RC4_Encrypt = Plaintext
End Function

Private Function RC4_Decrypt(ByVal Ciphertext As String) As String
On Error Resume Next
Dim lngLength As Long
Dim lngReturnValue As Long
lngLength = Len(Ciphertext)
lngReturnValue = CryptDecrypt(hClientReadKey, 0, False, 0, Ciphertext, lngLength)
RC4_Decrypt = Ciphertext
End Function

Private Function GenerateRandomBytes(ByVal Length As Long, ByRef TheString As String) As Boolean
On Error Resume Next
Dim i As Integer
Randomize
TheString = ""
For i = 1 To Length
TheString = TheString & Chr(Int(Rnd * 256))
Next
GenerateRandomBytes = CryptGenRandom(hCryptProv, Length, TheString)
End Function

Private Function MD5_Hash(ByVal TheString As String) As String
On Error Resume Next
Dim lngReturnValue As Long
Dim strHash As String
Dim hHash As Long
Dim lngHashLen As Long
lngReturnValue = CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash)
lngReturnValue = CryptHashData(hHash, TheString, Len(TheString), 0)
lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, vbNull, lngHashLen, 0)
strHash = String(lngHashLen, vbNullChar)
lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, strHash, lngHashLen, 0)
If hHash <> 0 Then CryptDestroyHash hHash
MD5_Hash = strHash
End Function

Private Sub ConnectSSL(sHost As String)
On Error Resume Next
WsSSL.Close
WsSSL.Connect sHost, 443
End Sub

Private Sub CloseSSL()
On Error Resume Next
WsSSL.Close
End Sub

Private Sub CertToPrivateKey()
On Error Resume Next
Const lPbkLen As Long = 1024
Dim lOffset As Long
Dim lStart As Long
Dim sBlkLen As String
Dim sRevKey As String
Dim ASNStart As Long
Dim ASNKEY As String
lOffset = CLng(lPbkLen \ 8)
lStart = 5 + (lOffset \ 128) * 2
ASNStart = InStr(1, ENCODED_CERT, Chr(48) & Chr(129) & Chr(137) & Chr(2) & Chr(129) & Chr(129) & Chr(0)) + lStart
ASNKEY = Mid(ENCODED_CERT, ASNStart, 128)
sRevKey = ReverseString(ASNKEY)
sBlkLen = CStr(Hex(lPbkLen \ 256))
If Len(sBlkLen) = 1 Then sBlkLen = "0" & sBlkLen
Private_KEY = (HexToBin("06020000" & "00A40000" & "52534131" & "00" & sBlkLen & "0000" & "01000100") & sRevKey)
End Sub

Private Function VerifyMAC(ByVal DecryptedRecord As String) As Boolean
On Error Resume Next
Dim PrependedMAC As String
Dim RecordData As String
Dim CalculatedMAC As String
PrependedMAC = Mid(DecryptedRecord, 1, 16)
RecordData = Mid(DecryptedRecord, 17)
CalculatedMAC = MD5_Hash(CLIENT_READ_KEY & RecordData & RecvSequence)
Call IncrementRecv
If CalculatedMAC = PrependedMAC Then
VerifyMAC = True
Else
VerifyMAC = False
End If
End Function

Private Function SendSequence() As String
On Error Resume Next
Dim TempString As String
Dim TempSequence As Double
Dim TempByte As Double
Dim i As Integer
TempSequence = SEND_SEQUENCE_NUMBER
For i = 1 To 4
TempByte = 256 * ((TempSequence / 256) - Int(TempSequence / 256))
TempSequence = Int(TempSequence / 256)
TempString = Chr(TempByte) & TempString
Next
SendSequence = TempString
End Function

Private Function RecvSequence() As String
On Error Resume Next
Dim TempString As String
Dim TempSequence As Double
Dim TempByte As Double
Dim i As Integer
TempSequence = RECV_SEQUENCE_NUMBER
For i = 1 To 4
TempByte = 256 * ((TempSequence / 256) - Int(TempSequence / 256))
TempSequence = Int(TempSequence / 256)
TempString = Chr(TempByte) & TempString
Next
RecvSequence = TempString
End Function

Private Sub SendClientHello(ByRef Socket As Winsock)
On Error Resume Next
iLayer = 0
bRaisedConnected = False
Call GenerateRandomBytes(16, CHALLENGE_DATA)
SEND_SEQUENCE_NUMBER = 0
RECV_SEQUENCE_NUMBER = 0
CLIENT_HELLO = Chr(1) & Chr(0) & Chr(2) & Chr(0) & Chr(3) & Chr(0) & Chr(0) & Chr(0) & Chr(Len(CHALLENGE_DATA)) & Chr(1) & Chr(0) & Chr(128) & CHALLENGE_DATA
If Socket.State = 7 Then Socket.SendData AddRecordHeader(CLIENT_HELLO)
End Sub

Private Sub SendMasterKey(ByRef Socket As Winsock)
On Error Resume Next
iLayer = 1
Call GenerateRandomBytes(32, MASTER_KEY)
Call CertToPrivateKey
Socket.SendData AddRecordHeader(Chr(2) & Chr(1) & Chr(0) & Chr(128) & Chr(0) & Chr(0) & Chr(0) & Chr(128) & Chr(0) & Chr(0) & _
ExportKeyBlob(MASTER_KEY, CLIENT_READ_KEY, CLIENT_WRITE_KEY, CHALLENGE_DATA, CONNECTION_ID, Private_KEY))
End Sub

Private Sub SendClientFinish(ByRef Socket As Winsock)
On Error Resume Next
iLayer = 2
Call SSLSend(Socket, Chr(3) & CONNECTION_ID)
End Sub

Private Sub SSLSend(ByRef Socket As Winsock, ByVal Plaintext As String)
On Error Resume Next
Dim SSLRecord As String
Dim OtherPart As String
Dim SendAnother As Boolean
If Len(Plaintext) > 32751 Then
SendAnother = True
Plaintext = Mid(Plaintext, 1, 32751)
OtherPart = Mid(Plaintext, 32752)
Else
SendAnother = False
End If
SSLRecord = AddMACData(Plaintext)
SSLRecord = RC4_Encrypt(SSLRecord)
SSLRecord = AddRecordHeader(SSLRecord)
Socket.SendData SSLRecord
If SendAnother = True Then
Call SSLSend(Socket, OtherPart)
End If
End Sub

Private Function AddMACData(ByVal Plaintext As String) As String
On Error Resume Next
AddMACData = MD5_Hash(CLIENT_WRITE_KEY & Plaintext & SendSequence) & Plaintext
End Function

Private Function AddRecordHeader(ByVal RecordData As String) As String
On Error Resume Next
Dim FirstChar As String
Dim LastChar As String
Dim TheLen As Long
TheLen = Len(RecordData)
FirstChar = Chr(128 + (TheLen \ 256))
LastChar = Chr(TheLen Mod 256)
AddRecordHeader = FirstChar & LastChar & RecordData
Call IncrementSend
End Function

Private Sub IncrementSend()
On Error Resume Next
SEND_SEQUENCE_NUMBER = SEND_SEQUENCE_NUMBER + 1
If SEND_SEQUENCE_NUMBER = 4294967296# Then SEND_SEQUENCE_NUMBER = 0
End Sub

Private Sub IncrementRecv()
On Error Resume Next
RECV_SEQUENCE_NUMBER = RECV_SEQUENCE_NUMBER + 1
If RECV_SEQUENCE_NUMBER = 4294967296# Then RECV_SEQUENCE_NUMBER = 0
End Sub

Private Function BytesToLen(ByVal TwoBytes As String) As Long
On Error Resume Next
Dim FirstByteVal As Long
FirstByteVal = Asc(Left(TwoBytes, 1))
If FirstByteVal >= 128 Then FirstByteVal = FirstByteVal - 128
BytesToLen = 256 * FirstByteVal + Asc(Right(TwoBytes, 1))
End Function

Private Function HexToBin(ByVal HexString As String) As String
On Error Resume Next
Dim BinString As String
BinString = ""
Dim i As Integer
For i = 1 To Len(HexString) Step 2
BinString = BinString & Chr(Val("&H" & Mid(HexString, i, 2)))
Next i
HexToBin = BinString
End Function

Private Sub wsSSL_Close()
On Error Resume Next
iLayer = 0
bRaisedConnected = False
WsSSL.Close
End Sub

Private Sub wsSSL_Connect()
On Error Resume Next
Call SendClientHello(WsSSL)
End Sub

Private Sub wsSSL_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim sdata As String
Dim lReachLen As Long
Do
If iSeekLen = 0 Then
If bytesTotal >= 2 Then
WsSSL.GetData sdata, vbString, 2
iSeekLen = BytesToLen(sdata)
bytesTotal = bytesTotal - 2
Else
Exit Sub
End If
End If
If bytesTotal >= iSeekLen Then
WsSSL.GetData sdata, vbString, iSeekLen
bytesTotal = bytesTotal - iSeekLen
Else
Exit Sub
End If
Select Case iLayer
Case 0:
ENCODED_CERT = Mid(sdata, 12, BytesToLen(Mid(sdata, 6, 2)))
CONNECTION_ID = Right(sdata, BytesToLen(Mid(sdata, 10, 2)))
Call IncrementRecv
Call SendMasterKey(WsSSL)
Case 1:
sdata = RC4_Decrypt(sdata)
If Right(sdata, Len(CHALLENGE_DATA)) = CHALLENGE_DATA Then
If VerifyMAC(sdata) Then Call SendClientFinish(WsSSL)
Else
WsSSL.Close
End If
Case 2:
sdata = RC4_Decrypt(sdata)
If VerifyMAC(sdata) = False Then WsSSL.Close
iLayer = 3
Case 3:
sdata = RC4_Decrypt(sdata)
If VerifyMAC(sdata) Then
sBuffer = sBuffer & Mid(sdata, 17)
End If
End Select
iSeekLen = 0
Loop Until bytesTotal = 0
If iLayer = 3 Then
If bRaisedConnected = False Then
bRaisedConnected = True
GoTo 1
End If
End If
If sBuffer <> "" Then
If sFirst = True Then
If InStr(1, sBuffer, "ymsgr=") Then
RaiseEvent SSLTokenData(SplitString(sBuffer, "ymsgr=", vbCrLf))
Else
RaiseEvent SSLTokenData("invalid")
CloseSSL
End If
Else
If InStr(1, sBuffer, "crumb=") Then
CloseSSL
DoEvents
RaiseEvent SSLAuthData(SplitString(sBuffer, "Y=", ";"), SplitString(sBuffer, "T=", ";"), SplitString(sBuffer, "crumb=", vbCrLf))
End If
End If
End If
Exit Sub
1
SendSSL SSLTMP
End Sub

Private Function ReverseString(ByVal TheString As String) As String
On Error Resume Next
Dim Reversed As String
Dim i As Integer
For i = Len(TheString) To 1 Step -1
Reversed = Reversed & Mid(TheString, i, 1)
Next i
ReverseString = Reversed
End Function

Private Function SendSSL(sdata As String)
On Error Resume Next
If WsSSL.State = sckConnected Then
Call SSLSend(WsSSL, sdata)
Else
Exit Function
End If
End Function

Public Sub GetToken(YID As String, YPASS As String, Chaln As String, Optional sHost As String)
Dim Str As String, Pck As String
sFirst = True
If sHost = "" Then sHost = "login.yahoo.com"
Pck = "src=ymsgr&login=" & YID & "&passwd=" & YPASS & "&chal=" & Chaln
Str = "POST /config/pwtoken_get HTTP/1.1" & vbCrLf
Str = Str + "Referer: https://" & sHost & vbCrLf
Str = Str + "Accept-Language: en-us" & vbCrLf
Str = Str + "Content-Type: application/x-www-form-urlencoded" & vbCrLf
Str = Str + "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1)" & vbCrLf
Str = Str + "Host: " & sHost & vbCrLf
Str = Str + "Content-Length: " & Len(Pck) & vbCrLf
Str = Str + "Connection: Close" & vbCrLf
SSLTMP = Str + "Cache-Control: no-cache" & vbCrLf & vbCrLf & Pck
ConnectSSL sHost
End Sub

Public Sub GetAuth(Tokn As String, Optional sHost As String)
Dim Str As String, Pck As String
sFirst = False
If sHost = "" Then sHost = "login.yahoo.com"
Pck = "src=ymsgr&token=" & Tokn
Str = "POST /config/pwtoken_login HTTP/1.1" & vbCrLf
Str = Str + "Referer: https://" & sHost & vbCrLf
Str = Str + "Accept-Language: en-us" & vbCrLf
Str = Str + "Content-Type: application/x-www-form-urlencoded" & vbCrLf
Str = Str + "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1)" & vbCrLf
Str = Str + "Host: " & sHost & vbCrLf
Str = Str + "Content-Length: " & Len(Pck) & vbCrLf
Str = Str + "Connection: Close" & vbCrLf
SSLTMP = Str + "Cache-Control: no-cache" & vbCrLf & vbCrLf & Pck
ConnectSSL sHost
End Sub

Private Sub Class_Initialize()
On Error Resume Next
Dim lngReturnValue As Long
Dim TheAnswer As Long
Set WsSSL = Form1.WsSSL
lngReturnValue = CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, CRYPT_NEWKEYSET)
If lngReturnValue = 0 Then
lngReturnValue = CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, 0)
If lngReturnValue = 0 Then TheAnswer = MsgBox("you do not have the required High Encryption Pack installed." & vbCrLf & "Would like to download this pack from Microsoft's website?", 16 + vbYesNo)
End If
If TheAnswer = vbYes Then
Call Shell("START http://www.microsoft.com/windows/ie/downloads/recommended/128bit/default.asp", vbHide)
End If
If TheAnswer = vbNo Then
End If
End Sub

Private Sub Class_Terminate()
On Error Resume Next
If hClientWriteKey <> 0 Then CryptDestroyKey hClientWriteKey
If hClientReadKey <> 0 Then CryptDestroyKey hClientReadKey
If hCryptProv <> 0 Then CryptReleaseContext hCryptProv, 0
End Sub

Private Function SplitString(ByVal TheString As String, ByVal StrStart As String, ByVal StrEnd As String) As String
On Error Resume Next
Dim X As Integer
X = InStr(1, TheString, StrStart) + Len(StrStart)
SplitString = Mid(TheString, X, InStr(X, TheString, StrEnd) - X)
End Function
