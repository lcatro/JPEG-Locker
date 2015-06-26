Attribute VB_Name = "InternetIO"
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, _
                         ByVal lAccessType As Long, ByVal sProxyName As String, _
                         ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
                         
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
                         
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" ( _
                         ByVal hInternetSession As Long, ByVal sUrl As String, _
                         ByVal sHeaders As String, ByVal lHeadersLength As Long, _
                         ByVal lFlags As Long, ByVal lContext As Long) As Long
                         
Private Const INTERNET_FLAG_RELOAD = &H80000000
                         
Private Declare Function InternetReadFile Lib "wininet.dll" ( _
                         ByVal hFile As Long, ByVal sBuffer As String, _
                         ByVal lNumBytesToRead As Long, _
                         lNumberOfBytesRead As Long) As Integer
                         
Private Declare Function InternetReadFileByte Lib "wininet.dll" Alias "InternetReadFile" ( _
                         ByVal hFile As Long, _
                         ByRef sBuffer As Byte, _
                         ByVal lNumberOfBytesToRead As Long, _
                         lNumberOfBytesRead As Long) As Integer
                                    
Private Declare Function InternetQueryDataAvailable Lib "wininet.dll" (ByVal hInet As Long, dwAvail As Long, _
                         ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
                         
Private Declare Function InternetSetFilePointer Lib "wininet.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, _
                         lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long, _
                         ByVal dwContext As Long) As Long
                         
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer

Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal _
hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As _
String, ByVal lOptionalLength As Long) As Integer

Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" _
(ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, _
ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer

Private Const HTTP_QUERY_CONTENT_TYPE = 1
Private Const HTTP_QUERY_CONTENT_LENGTH = 5
Private Const HTTP_QUERY_EXPIRES = 10
Private Const HTTP_QUERY_LAST_MODIFIED = 11
Private Const HTTP_QUERY_PRAGMA = 17
Private Const HTTP_QUERY_VERSION = 18
Private Const HTTP_QUERY_STATUS_CODE = 19
Private Const HTTP_QUERY_STATUS_TEXT = 20
Private Const HTTP_QUERY_RAW_HEADERS = 21
Private Const HTTP_QUERY_RAW_HEADERS_CRLF = 22
Private Const HTTP_QUERY_FORWARDED = 30
Private Const HTTP_QUERY_SERVER = 37
Private Const HTTP_QUERY_USER_AGENT = 39
Private Const HTTP_QUERY_SET_COOKIE = 43
Private Const HTTP_QUERY_REQUEST_METHOD = 45
Private Const HTTP_STATUS_DENIED = 401
Private Const HTTP_STATUS_PROXY_AUTH_REQ = 407
 
Private Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000
Private Const HTTP_QUERY_FLAG_NUMBER = &H20000000

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Dim InternetHwnd As Long
Dim URLHwnd As Long

Function DateAvailableSize() As Long
Dim Buf As Long
InternetQueryDataAvailable URLHwnd, Buf, 0, 0
DateSize = Buf
End Function

Sub InitWinInet()
InternetHwnd = InternetOpen(App.EXEName, INTERNET_OPEN_TYPE_PRECONFIG, 0, 0, 0)
End Sub

Function OpenURL(ByVal URL As String) As Boolean
URLHwnd = InternetOpenUrl(InternetHwnd, URL, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
OpenURL = IIf(URLHwnd <> 0, True, False)
End Function

Function GetFileSize() As Long
Dim PointBuf As String * 16
If HttpQueryInfo(URLHwnd, HTTP_QUERY_CONTENT_LENGTH, ByVal PointBuf, 16, 0) Then
GetFileSize = CLng(PointBuf)
End If
End Function

Sub SetFilePoint(ByVal Point As Long)
InternetSetFilePointer URLHwnd, Point, 0, 0, 0
End Sub

Function ReadData(ByVal Size As Long) As String
Dim ReadBuf As String
ReadBuf = Space(Size)
InternetReadFile URLHwnd, ReadBuf, Size, 0
ReadData = ReadBuf
End Function

Function ReadDataByte(ByVal Size As Long) As Variant
Dim ReadBuf() As Byte
ReDim ReadBuf(Size)
InternetReadFileByte URLHwnd, ReadBuf(0), Size, 0
For I = 0 To Size
ReadDataByte(I) = ReadBuf(I)
Next
End Function

Function ReadDataBit() As String
Dim ReadBuf As String
ReadBuf = Space(1)
InternetReadFile URLHwnd, ReadBuf, 1, 0
ReadDataBit = ReadBuf
End Function

Sub CloseURL()
InternetCloseHandle URLHwnd
URLHwnd = 0
End Sub

Sub CloseWinInet()
InternetCloseHandle InternetHwnd
InetHwnd = 0
End Sub

