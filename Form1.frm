VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Main 
   BorderStyle     =   0  'None
   Caption         =   "JPEG Locker"
   ClientHeight    =   8250
   ClientLeft      =   -45
   ClientTop       =   -375
   ClientWidth     =   12000
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   12000
   StartUpPosition =   2  '屏幕中心
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   5040
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.ListBox FileList 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000001&
      Height          =   750
      ItemData        =   "Form1.frx":19862
      Left            =   3000
      List            =   "Form1.frx":19864
      MultiSelect     =   2  'Extended
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      Top             =   3480
      Width           =   4950
   End
   Begin MSComctlLib.ProgressBar Progress 
      Height          =   210
      Left            =   1860
      TabIndex        =   1
      Top             =   7870
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label WEB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "POWERED BY IC2012.CN"
      BeginProperty Font 
         Name            =   "@宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FDB313&
      Height          =   180
      Left            =   9600
      TabIndex        =   5
      Top             =   7870
      Width           =   1800
   End
   Begin VB.Image MiniButton 
      Height          =   270
      Left            =   11205
      Top             =   120
      Width           =   285
   End
   Begin VB.Image SelectButton 
      Height          =   750
      Left            =   7950
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JPEG Locker"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   165
      Width           =   990
   End
   Begin VB.Image ExitButton 
      Height          =   270
      Left            =   11535
      ToolTipText     =   "退出"
      Top             =   120
      Width           =   285
   End
   Begin VB.Image OptionalButton 
      Height          =   270
      Left            =   10875
      ToolTipText     =   "设置"
      Top             =   120
      Width           =   285
   End
   Begin VB.Label FilePath 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "存放目录:"
      ForeColor       =   &H00666666&
      Height          =   180
      Left            =   3000
      TabIndex        =   2
      Top             =   4800
      Width           =   810
   End
   Begin VB.Image StartButton 
      Height          =   1110
      Left            =   4680
      Top             =   5760
      Width           =   2685
   End
   Begin VB.Label Version 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "程序版本:"
      ForeColor       =   &H00666666&
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   7870
      Width           =   810
   End
   Begin VB.Image Image1 
      Height          =   8250
      Left            =   0
      OLEDropMode     =   1  'Manual
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const KEY_SAVEPATH = "SavePath"
Private Const KEY_DELETESORC = "DeleteSorc"
Private Const KEY_BADEXIT = "BadExit"
Private Const KEY_PASSWORDCANSEE = "PasswordCanSee"

Private Const Extension_Filename_LOCK = ".lck"
Private Const Extension_Filename_JPEG = ".jpg"
Private Const Extension_Filename_JPEG_1 = ".jpeg"

Private Const PIXEL_BACKGOUND_HEIGHT = 550
Private Const PIXEL_BACKGOUND_WIDTH = 800
Private Const PIXEL_SELECTBUTTON_HEIGHT = 48
Private Const PIXEL_SELECTBUTTON_WIDTH = 89

Private Const PIXEL_OFFSET_WEB_LEFT = 25
Private Const PIXEL_OFFSET_WEB_RIGHT = 15
Private Const PIXEL_CHANGE_PROGRESS_LEN_HIGHT = 15
Private Const PIXEL_CHANGE_PROGRESS_LEN_LOW = 15

Private Const STR_HELP = "  点击文件夹图标加载加密/解密的图片,然后开始"

Dim SelectButtonStyle As STYLE_BUTTON
Dim StartButtonStyle As STYLE_BUTTON
Dim OptionalButtonStyle As STYLE_BUTTON
Dim MiniButtonStyle As STYLE_BUTTON
Dim ExitButtonStyle As STYLE_BUTTON

Private Sub ShowHelp()
FileList.Clear
FileList.AddItem ""
FileList.AddItem STR_HELP
FileList.Enabled = False
End Sub

Sub SetStartButtonStyleToEnter()
StartButtonStyle = MOUSE_ENTER
End Sub

Sub ShowSelectFile()
FileList.Clear
FileList.Enabled = True
End Sub

Sub SelectFileAll()
FileDialog.DialogTitle = "选择文件"
FileDialog.Filter = "JPEG文件(*.jpg)|*.jpg|上锁图片(*.lck)|*.lck"
End Sub

Private Sub SelectFileJPEG()
FileDialog.DialogTitle = "选择文件"
FileDialog.Filter = "JPEG文件(*.jpg)|*.jpg"
End Sub

Private Sub SelectFileLock()
FileDialog.DialogTitle = "选择文件"
FileDialog.Filter = "上锁图片文件(*.lck)|*.lck"
End Sub

Private Function IsMixExtensionFileName(ByVal filename As String) As Boolean
Dim Cache() As String
Cache = Split(filename, ".")

If UBound(Cache) > 1 Then
    IsMixExtensionFileName = True
Else
    IsMixExtensionFileName = False
End If
End Function

Private Sub Form_Load()
On Error Resume Next
InternetIO.InitWinInet
'Menu.TryToUpdate_Click

SelectButtonStyle = MOUSE_ENTER
StartButtonStyle = MOUSE_ENTER
OptionalButtonStyle = MOUSE_ENTER
MiniButtonStyle = MOUSE_ENTER
ExitButtonStyle = MOUSE_ENTER

Stat = WaitSelect

FileDialog.Flags = cdlOFNHideReadOnly Or cdlOFNAllowMultiselect Or cdlOFNExplorer

Version.Caption = Version.Caption & App.Major & "." & App.Minor & "." & App.Revision

SavePath = SaveProcessData.GetData(KEY_SAVEPATH)
SavePath = IIf(SavePath = "", App.Path, SavePath)
FilePath.Caption = "存放目录:" & SavePath & "\"
DeleteSorc = SaveProcessData.GetDataBoolean(KEY_DELETESORC)
BadExit = SaveProcessData.GetDataBoolean(KEY_BADEXIT)
PasswordCanSee = SaveProcessData.GetDataBoolean(KEY_PASSWORDCANSEE)
If DeleteSorc Then
    Menu.NeedDeleteSorcSetting_Click
Else
    Menu.NoDeleteSorcSetting_Click
End If
If BadExit Then
    Menu.NeedBadSetting_Click
Else
    Menu.NoBadSetting_Click
End If
If PasswordCanSee Then
    Menu.NeedPasswordSetting_Click
Else
    Menu.NoPasswordSetting_Click
End If

SelectFileAll
ShowHelp

Me.Picture = LoadResPicture(101, vbResBitmap)

If (PublicData.DPI_BIG = PublicData.GetSystemDPI) Then
    PublicData.ResetFormSizeForHeightDPI Me, PIXEL_BACKGOUND_HEIGHT, PIXEL_BACKGOUND_WIDTH
    Dim Ctrl As Control
    For Each Ctrl In Me
        Ctrl.Top = CalcuDPIBigPixel(Ctrl.Top / DPI_SMALL_TIWP)
        Ctrl.Left = CalcuDPIBigPixel(Ctrl.Left / DPI_SMALL_TIWP)
    Next
    
    FileList.Width = SelectButton.Left - FileList.Left
    SelectButton.Height = FileList.Height
    SelectButton.Width = CLng(SelectButton.Height / PIXEL_SELECTBUTTON_HEIGHT) * PIXEL_SELECTBUTTON_WIDTH
    
    Progress.Width = CalcuDPIBigPixel(Progress.Width / PublicData.DPI_SMALL_TIWP) - CalcuDPIBigPixel(PIXEL_CHANGE_PROGRESS_LEN_HIGHT)
    WEB.Left = WEB.Left - CalcuDPIBigPixel(PIXEL_OFFSET_WEB_LEFT)
Else
    Progress.Left = Progress.Left + CalcuDPIBigPixel(PIXEL_CHANGE_PROGRESS_LEN_LOW)
    WEB.Left = WEB.Left + CalcuDPIBigPixel(PIXEL_OFFSET_WEB_RIGHT)
End If

PrintPictureOnButton 0, 0
End Sub

Private Sub Form_Terminate()
SaveProcessData.SaveData KEY_SAVEPATH, SavePath
SaveProcessData.SaveDataBoolean KEY_DELETESORC, DeleteSorc
SaveProcessData.SaveDataBoolean KEY_BADEXIT, BadExit
SaveProcessData.SaveDataBoolean KEY_PASSWORDCANSEE, PasswordCanSee

InternetIO.CloseWinInet
End Sub

Private Sub FileList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = MOUSE_RIGHT Then Me.PopupMenu Menu.FileListOptional
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
PublicData.MoveForm Me
End Sub

Sub PrintPictureOnButton(ByVal x As Long, ByVal y As Long)
If ((SelectButton.Left <= x Or x <= SelectButton.Left + SelectButton.Width) And (SelectButton.Top <= y Or y <= SelectButton.Top + SelectButton.Height)) And SelectButtonStyle = MOUSE_ENTER Then
    SelectButton.Picture = LoadResPicture(107, vbResBitmap)
    SelectButtonStyle = MOUSE_LEAVE
End If
If ((StartButton.Left <= x Or x <= StartButton.Left + StartButton.Width) And (StartButton.Top <= y Or y <= StartButton.Top + StartButton.Height)) And StartButtonStyle = MOUSE_ENTER Then
    If Stat = WaitSelect Then
        StartButton.Picture = LoadResPicture(109, vbResBitmap)
    ElseIf Stat = IsLock Then
        StartButton.Picture = LoadResPicture(103, vbResBitmap)
    ElseIf Stat = IsUnlock Then
        StartButton.Picture = LoadResPicture(105, vbResBitmap)
    End If
    StartButtonStyle = MOUSE_LEAVE
End If
If ((OptionalButton.Left <= x Or x <= OptionalButton.Left + OptionalButton.Width) And (OptionalButton.Top <= y Or y <= OptionalButton.Top + OptionalButton.Height)) And OptionalButtonStyle = MOUSE_ENTER Then
    OptionalButton.Picture = LoadResPicture(115, vbResBitmap)
    OptionalButtonStyle = MOUSE_LEAVE
End If
If ((MiniButton.Left <= x Or x <= MiniButton.Left + MiniButton.Width) And (MiniButton.Top <= y Or y <= MiniButton.Top + MiniButton.Height)) And MiniButtonStyle = MOUSE_ENTER Then
    MiniButton.Picture = LoadResPicture(113, vbResBitmap)
    MiniButtonStyle = MOUSE_LEAVE
End If
If ((ExitButton.Left <= x Or x <= ExitButton.Left + ExitButton.Width) And (ExitButton.Top <= y Or y <= ExitButton.Top + ExitButton.Height)) And ExitButtonStyle = MOUSE_ENTER Then
    ExitButton.Picture = LoadResPicture(111, vbResBitmap)
    ExitButtonStyle = MOUSE_LEAVE
End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
PrintPictureOnButton x, y
End Sub

Private Sub FileList_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Image1_OLEDragDrop Data, Effect, Button, Shift, x, y
End Sub

Private Sub Image1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim filename() As String
filename = Split(Data.Files.Item(1), "/")

If IsMixExtensionFileName(filename(UBound(filename))) Then
    PublicData.PrintErrMessage "你选取的文件中含有混合拓展名"
    Exit Sub
End If

If FileList.List(1) = STR_HELP Then FileList.Clear

If Stat = WaitSelect Then
    If InStr(Data.Files.Item(1), Extension_Filename_JPEG) Or InStr(Data.Files.Item(1), Extension_Filename_JPEG_1) Then
        Stat = IsLock
        SelectFileJPEG
    ElseIf InStr(Data.Files.Item(1), Extension_Filename_LOCK) Then
        Stat = IsUnlock
        SelectFileLock
    End If
    
    ShowSelectFile
    StartButton_MouseMove 0, 0, 0, 0
End If

For I = 1 To Data.Files.Count
    If (Stat = IsLock And (InStr(Data.Files.Item(I), Extension_Filename_JPEG) Or InStr(Data.Files.Item(I), Extension_Filename_JPEG_1))) Or (Stat = IsUnlock And (InStr(Data.Files.Item(I), Extension_Filename_LOCK))) Then FileList.AddItem Data.Files.Item(I)
Next
End Sub

Private Sub SelectButton_Click()
On Error GoTo Cancel

If FileList.List(1) = STR_HELP Then ShowSelectFile

FileDialog.filename = ""
FileDialog.ShowOpen

Dim FilePath As String
Dim List() As String
FilePath = FileDialog.filename

If InStr(1, FilePath, Chr(0)) > 0 Then
    List = Split(FilePath, Chr(0))
    If InStr(List(1), Extension_Filename_JPEG) > 0 Or InStr(List(1), Extension_Filename_JPEG_1) > 0 And (Stat = WaitSelect Or Stat = IsLock) Then
        Stat = IsLock
        SelectFileJPEG
    ElseIf InStr(List(1), Extension_Filename_LOCK) > 0 And (Stat = WaitSelect Or Stat = IsUnlock) Then
        Stat = IsUnlock
        SelectFileLock
    End If
    
    For I = 1 To UBound(List)
        If (Stat = IsLock And (InStr(List(I), Extension_Filename_JPEG) Or InStr(List(I), Extension_Filename_JPEG_1))) Or (Stat = IsUnlock And InStr(List(I), Extension_Filename_LOCK)) Then
            If Not IsMixExtensionFileName(List(1)) Then
                FileList.AddItem List(0) & "\" & List(I)
            Else
                PublicData.PrintErrMessage "你选取的文件中含有混合拓展名"
                Exit Sub
            End If
        End If
    Next
Else
    Dim filename() As String
    filename = Split(FilePath, "/")
    
    If IsMixExtensionFileName(filename(UBound(filename))) Then
        PublicData.PrintErrMessage "你选取的文件中含有混合拓展名"
        Exit Sub
    End If

    If (InStr(FilePath, Extension_Filename_JPEG) > 0 Or InStr(FilePath, Extension_Filename_JPEG_1) > 0) And (Stat = WaitSelect Or Stat = IsLock) Then
        Stat = IsLock
        SelectFileJPEG
    ElseIf InStr(FilePath, Extension_Filename_LOCK) > 0 And (Stat = WaitSelect Or Stat = IsUnlock) Then
        Stat = IsUnlock
        SelectFileLock
    End If

    FileList.AddItem FilePath
End If
StartButtonStyle = MOUSE_ENTER
PrintPictureOnButton StartButton.Left - 1, 0
Exit Sub

Cancel:
End Sub

Private Function PrintErrFileList(ByVal ErrSums As Long, ErrFileList() As Long) As String
Dim Report As String

For I = 0 To ErrSums - 1
    If I = ErrSums Then
        Report = Report & FileList.List(ErrFileList(I))
        Exit For
    End If
    Report = Report & FileList.List(ErrFileList(I)) & vbCrLf
Next

PrintErrFileList = Report
End Function

Private Function GetPath(ByVal UpBound As Long, FilePath() As String) As String
Dim rtn As String
For I = 0 To UpBound - 1
    rtn = rtn & FilePath(I) & "\"
Next

GetPath = Left(rtn, Len(rtn) - 1)
End Function

Private Sub StartButton_Click()
On Error GoTo ERR:
If FileList.ListCount = 0 Or FileList.List(1) = STR_HELP Then
    PublicData.PrintErrMessage "请先选择图片"
    Exit Sub
End If

Dim BadSums As Long
Dim BadIndex() As Long
ReDim BadIndex(FileList.ListCount)

Progress.Max = FileList.ListCount
Progress.Value = Progress.Min

Dim InputPassword As String
InputPassword = GetPassword
If InputPassword = "" Then Exit Sub

GetTimeTick.StartTimeTick

For I = 0 To FileList.ListCount - 1
    Dim FilePath() As String
    FilePath = Split(FileList.List(I), "\")
    Dim filename() As String
    filename = Split(FilePath(UBound(FilePath)), ".")
    
    If PictureLock.IsLockFile(FileList.List(I)) Then
        If Not PictureLock.UnlockPicture(FileList.List(I), IIf(DeleteSorc, GetPath(UBound(FilePath), FilePath) & "\" & filename(0) & Extension_Filename_JPEG, SavePath & "\" & filename(0) & Extension_Filename_JPEG), InputPassword) Then
            If Not BadExit Then
                BadIndex(BadSums) = I
                BadSums = BadSums + 1
            Else
                PublicData.PrintErrDetail "1个密码错误!", FileList.List(I)
                Exit Sub
            End If
        End If
        
        If I = FileList.ListCount - 1 And BadSums > 0 Then
            PublicData.PrintErrDetail BadSums & "个密码错误!", PrintErrFileList(BadSums, BadIndex)
            Exit Sub
        End If
    Else
        PictureLock.LockPicture FileList.List(I), IIf(DeleteSorc, GetPath(UBound(FilePath), FilePath) & "\" & filename(0) & Extension_Filename_LOCK, SavePath & "\" & filename(0) & Extension_Filename_LOCK), InputPassword
    End If
    If DeleteSorc Then Kill FileList.List(I)
    
    Progress.Value = I + 1
Next

PrintUseTime GetTimeTick.GetCurrentTimeTick, IIf(Stat = IsLock, lck, unlck)

FileList.Clear
SelectFileAll
Stat = WaitSelect
StartButton_MouseMove 0, 0, 0, 0

Exit Sub
ERR:
PublicData.PrintErrMessage "您输入了错误的文件路径"
FileList.Clear
SelectFileAll
Stat = WaitSelect
Image1_MouseMove 0, 0, StartButton.Left - 1, 0
End Sub

Private Sub OptionalButton_Click()
Me.PopupMenu Menu.Optional
End Sub

Private Sub MiniButton_Click()
Me.WindowState = 1
End Sub

Private Sub ExitButton_Click()
Form_Terminate
End
End Sub

Private Sub SelectButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
SelectButton.Picture = LoadResPicture(106, vbResBitmap)
SelectButtonStyle = MOUSE_ENTER
End Sub

Private Sub StartButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Stat = WaitSelect Then
    StartButton.Picture = LoadResPicture(108, vbResBitmap)
ElseIf Stat = IsLock Then
    StartButton.Picture = LoadResPicture(102, vbResBitmap)
ElseIf Stat = IsUnlock Then
    StartButton.Picture = LoadResPicture(104, vbResBitmap)
End If
StartButtonStyle = MOUSE_ENTER
End Sub

Private Sub OptionalButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
OptionalButton.Picture = LoadResPicture(114, vbResBitmap)
OptionalButtonStyle = MOUSE_ENTER
End Sub

Private Sub MiniButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
MiniButton.Picture = LoadResPicture(112, vbResBitmap)
MiniButtonStyle = MOUSE_ENTER
End Sub

Private Sub ExitButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ExitButton.Picture = LoadResPicture(110, vbResBitmap)
ExitButtonStyle = MOUSE_ENTER
End Sub

Private Sub WEB_Click()
Menu.GotoIC2012
End Sub
