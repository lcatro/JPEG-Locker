Attribute VB_Name = "Update"

'<Ver>3.2.0</Ver>
'<Download Main>http://ic2012.cn/Locker/Locker.rar</Download Main>
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Const BINDF_GETNEWESTVERSION As Long = &H10

Private Const STR_VER_LEFT = "<Ver>"
Private Const STR_VER_RIGHT = "</Ver>"
Private Const STR_DOWNLOADMAIN_LEFT = "<Download Main>"
Private Const STR_DOWNLOADMAIN_RIGHT = "</Download Main>"
Private Const STR_DOWNLOADOTHERFILE_LEFT = "<Download Other File>"
Private Const STR_DOWNLOADOTHERFILE_RIGHT = "</Download Other File>"
Private Const STR_DOWNLOADOTHERFILE_NAME_LEFT = "<Download Other File Name>"
Private Const STR_DOWNLOADOTHERFILE_NAME_RIGHT = "</Download Other File Name>"
Private Const STR_DOWNLOADOTHERFILE_URL_LEFT = "<Download Other File URL>"
Private Const STR_DOWNLOADOTHERFILE_URL_RIGHT = "</Download Other File URL>"

Private Const OK = 0

Sub FilterString(Inputstr As String, ByVal FilterString As String, Optional FilterFromLeft As Boolean = True)
If FilterFromLeft Then
    Inputstr = Mid(Inputstr, InStr(Inputstr, FilterString) + Len(FilterString))
Else
    Inputstr = Left(Inputstr, InStr(Inputstr, FilterString) - 1)
End If
End Sub

Private Sub ResoltVersionString(ByVal VersionString As String, Optional MainVer As Long, Optional MinorVer As Long, Optional RevisionVer As Long)
'VersionString -> "Ver=x.x.x"
On Error Resume Next
MainVer = CLng(Left(VersionString, InStr(VersionString, ".") - 1))
MinorVer = CLng(Left(Right(VersionString, InStr(VersionString, ".") + 1), InStr(VersionString, ".") - 1))
RevisionVer = CLng(Right(Right(VersionString, InStr(VersionString, ".") + 1), InStr(VersionString, ".") + 1))
End Sub

Function DownLoadFile(ByVal URL As String, ByVal SavePath As String) As Boolean
DownLoadFile = IIf(URLDownloadToFile(0, URL, SavePath, BINDF_GETNEWESTVERSION, 0) = OK, True, flase)
End Function

Function CheckUpdate(ByVal UpdateVersionFileURL As String, DownloadFileURL As String, Optional DownloadOtherFiles As Long = 0) As Boolean
On Error GoTo ERR
If Not InternetIO.OpenURL(UpdateVersionFileURL) Then GoTo ERR

Dim Buffer As String
Buffer = InternetIO.ReadData(InternetIO.GetFileSize)
If Buffer = "" Then GoTo ERR
InternetIO.CloseURL

Dim StrList() As String
StrList = Split(Buffer, vbCrLf)

If InStr(StrList(0), STR_VER_LEFT) > 0 Then
    Dim Ver As String
    Ver = StrList(0)
    
    FilterString Ver, STR_VER_LEFT
    FilterString Ver, STR_VER_RIGHT, False
    
    Dim Main As Long
    Dim Minor As Long
    Dim Revision As Long
    
    ResoltVersionString Ver, Main, Minor, Revision
 
    If Main > App.Major Or Minor > App.Minor Or Revision > App.Revision Then
        DownloadFileURL = StrList(1)
        FilterString DownloadFileURL, STR_DOWNLOADMAIN_LEFT
        FilterString DownloadFileURL, STR_DOWNLOADMAIN_RIGHT, False
        
        If UBound(StrList) - 2 > 1 Then
            FilterString StrList(2), STR_DOWNLOADOTHERFILE_LEFT
            FilterString StrList(2), STR_DOWNLOADOTHERFILE_RIGHT, False
            DownloadOtherFiles = CLng(StrList(2))
        End If
        
        CheckUpdate = True
        Exit Function
    End If
End If

ERR:
CheckUpdate = False
End Function

Function GetOtherFilesDownloadURL(ByVal UpdateVersionFileURL As String, DownloadOtherFiles As Long, DownloadOtherFileName() As String, DownloadOtherFileURL() As String) As Boolean
On Error GoTo ERR
If Not InternetIO.OpenURL(UpdateVersionFileURL) Then GoTo ERR

Dim Buffer As String
Buffer = InternetIO.ReadData(InternetIO.GetFileSize)
If Buffer = "" Then GoTo ERR
InternetIO.CloseURL

Dim StrList() As String
StrList = Split(Buffer, vbCrLf)

If UBound(StrList) <= 2 And Not (UBound(StrList) Mod 2) = 0 Then GoTo ERR

FilterString StrList(2), STR_DOWNLOADOTHERFILE_LEFT
FilterString StrList(2), STR_DOWNLOADOTHERFILE_RIGHT, False
DownloadOtherFiles = CLng(StrList(2))

For I = 3 To UBound(StrList) Step 2
    FilterString StrList(I), STR_DOWNLOADOTHERFILE_NAME_LEFT
    FilterString StrList(I), STR_DOWNLOADOTHERFILE_NAME_RIGHT, False
    DownloadOtherFileName(I - 3) = StrList(I)
    FilterString StrList(I + 1), STR_DOWNLOADOTHERFILE_URL_LEFT
    FilterString StrList(I + 1), STR_DOWNLOADOTHERFILE_URL_RIGHT, False
    DownloadOtherFileURL(I - 3) = StrList(I + 1)
Next

GetOtherFilesDownloadURL = True
Exit Function

ERR:
GetOtherFilesDownloadURL = False
End Function
