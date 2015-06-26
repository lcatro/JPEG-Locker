VERSION 5.00
Begin VB.Form Menu 
   Caption         =   "MenuForm"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu Optional 
      Caption         =   "选项"
      Visible         =   0   'False
      Begin VB.Menu OptionalSave 
         Caption         =   "设置存放目录"
      End
      Begin VB.Menu OtherSetting 
         Caption         =   "其它设置"
         Begin VB.Menu DeleteSorcSetting 
            Caption         =   "是否替换源文件"
            Begin VB.Menu NeedDeleteSorcSetting 
               Caption         =   "是"
            End
            Begin VB.Menu NoDeleteSorcSetting 
               Caption         =   "否"
            End
         End
         Begin VB.Menu BadSetting 
            Caption         =   "处理失败时是否退出处理"
            Begin VB.Menu NeedBadSetting 
               Caption         =   "需要"
            End
            Begin VB.Menu NoBadSetting 
               Caption         =   "不需要"
            End
         End
         Begin VB.Menu PasswordSetting 
            Caption         =   "启用明文密码输入"
            Begin VB.Menu NeedPasswordSetting 
               Caption         =   "启用"
            End
            Begin VB.Menu NoPasswordSetting 
               Caption         =   "不启用"
            End
         End
      End
      Begin VB.Menu Nothing1 
         Caption         =   "-"
      End
      Begin VB.Menu TryToUpdate 
         Caption         =   "检查更新"
      End
      Begin VB.Menu Nothing2 
         Caption         =   "-"
      End
      Begin VB.Menu GiveMeAdvice 
         Caption         =   "给产品提意见"
      End
      Begin VB.Menu ShareIt 
         Caption         =   "分享程序"
      End
      Begin VB.Menu About 
         Caption         =   "关于我们"
      End
   End
   Begin VB.Menu FileListOptional 
      Caption         =   "文件列表选项"
      Visible         =   0   'False
      Begin VB.Menu ClearThis 
         Caption         =   "清除该项"
      End
      Begin VB.Menu ClearAll 
         Caption         =   "清除所有项"
      End
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Sub sapiCoTaskMemFree Lib "ole32" Alias "CoTaskMemFree" (ByVal pv As Long)
                        
Private Const BIF_RETURNONLYFSDIRS = &H1

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Const URL_UPDATE_VERSION = "http://www.ic2012.cn/software/LockerUpdateVersion.html"

Private Const URL_PRODUCE_SHARE = "http://www.ic2012.cn/thread-htm-fid-53.html"
Private Const URL_IC2012 = "http://www.ic2012.cn"

Private Const STR_STAT_DELETESORC = "自动替换目标文件"

Private Function BrowseFolder(szDialogTitle As String, hwnd As Long) As String
  Dim x As Long, bi As BROWSEINFO, dwIList As Long
  Dim szPath As String, wPos As Integer
  
    With bi
        .hOwner = hwnd
        .lpszTitle = szDialogTitle
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    
    dwIList = SHBrowseForFolder(bi)
    szPath = Space$(512)
    x = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
    
    If x Then
        wPos = InStr(szPath, Chr(0))
        BrowseFolder = Left$(szPath, wPos - 1)
        Call sapiCoTaskMemFree(dwIList)
    Else
        BrowseFolder = vbNullString
    End If
End Function

Private Sub About_Click()
MsgBox "本程序由LCatro 开发,感谢你的使用.." & vbCrLf & vbCrLf & "特别鸣谢:" & vbCrLf & "Yenter(提交漏洞与建议)" & vbCrLf & "傻X(提交漏洞与建议)" & vbCrLf & "小火柴(提交建议)", vbOKOnly, "关于我们"
End Sub

Private Sub ClearAll_Click()
Stat = WaitSelect
Main.SetStartButtonStyleToEnter
Main.PrintPictureOnButton Main.StartButton.Left + 1, Main.StartButton.Top + 1
Main.SelectFileAll

Main.FileList.Clear
End Sub

Private Sub ClearThis_Click()
Main.FileList.RemoveItem Main.FileList.ListIndex

If Main.FileList.ListCount = 0 Then
Stat = WaitSelect
Main.SetStartButtonStyleToEnter
Main.PrintPictureOnButton Main.StartButton.Left, Main.StartButton.Top
End If
End Sub

Sub NeedDeleteSorcSetting_Click()
Main.FilePath.Caption = STR_STAT_DELETESORC
DeleteSorc = True
NeedDeleteSorcSetting.Enabled = False
NoDeleteSorcSetting.Enabled = True
End Sub

Sub NeedBadSetting_Click()
BadExit = True
NeedBadSetting.Enabled = False
NoBadSetting.Enabled = True
End Sub

Sub NeedPasswordSetting_Click()
PasswordCanSee = True
NeedPasswordSetting.Enabled = False
NoPasswordSetting.Enabled = True
End Sub

Sub NoDeleteSorcSetting_Click()
Main.FilePath.Caption = "存放目录:" & SavePath & "\"
DeleteSorc = False
NeedDeleteSorcSetting.Enabled = True
NoDeleteSorcSetting.Enabled = False
End Sub

Sub NoBadSetting_Click()
BadExit = False
NeedBadSetting.Enabled = True
NoBadSetting.Enabled = False
End Sub

Sub NoPasswordSetting_Click()
PasswordCanSee = False
NeedPasswordSetting.Enabled = True
NoPasswordSetting.Enabled = False
End Sub

Private Sub OptionalSave_Click()
Dim rtn As String
rtn = BrowseFolder("设置存放路径", Me.hwnd)

If Not rtn = "" Then
    SavePath = rtn
    Main.FilePath.Caption = "存放目录:" & SavePath & "\"
End If
End Sub

Private Sub GiveMeAdvice_Click()
GiveAdvice.Show
End Sub

Private Sub ShareIt_Click()
ShellExecute Me.hwnd, "Open", URL_PRODUCE_SHARE, 0, 0, 0
End Sub

Sub GotoIC2012()
ShellExecute Me.hwnd, "Open", URL_IC2012, 0, 0, 0
End Sub

Sub TryToUpdate_Click()
On Error GoTo ERR:
Dim DownloadFileURL As String
If Update.CheckUpdate(URL_UPDATE_VERSION, DownloadFileURL) Then
    If (MsgBox("程序可以更新,需要吗?", vbYesNo + vbDefaultButton2 + vbQuestion, "更新程序") = vbYes) Then
        If DownloadFileURL = "" Then
            MsgBox "程序更新失败,原因:" & vbCrLf & "获取更新文件地址失败", vbCritical, , "更新程序"
            Exit Sub
        End If
        
        If Update.DownLoadFile(DownloadFileURL, App.Path & "\Update.rar") Then
            MsgBox "程序更新失败,原因:" & vbCrLf & "下载文件失败", vbCritical, , "更新程序"
            Exit Sub
        End If
        RAR.RARExecute OP_EXTRACT, App.Path & "\Update.rar", App.Path
        
        Kill App.Path & "\Update.rar"
        
        If Shell(App.Path & "\HelpUpdate.exe " & App.Path & " " & App.EXEName & " " & GetCurrentProcessId, vbMinimizedNoFocus) = 0 Then MsgBox "程序更新失败,原因:" & vbCrLf & "缺少HelpUpdate.exe 程序" & vbCrLf & "请手动更新,最新版本的程序已经下载到当前目录(Update.exe)", vbCritical, , "更新程序"
        End
    End If
Else
    MsgBox "当前已是最新版本", vbYes, "更新程序"
End If
Exit Sub

ERR:
MsgBox "更新失败", vbCritical, "更新程序"
End Sub
