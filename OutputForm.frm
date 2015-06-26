VERSION 5.00
Begin VB.Form OutputForm 
   BorderStyle     =   0  'None
   ClientHeight    =   2700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   2700
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.TextBox AboutDetail 
      Appearance      =   0  'Flat
      Height          =   3615
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3120
      Width           =   6375
   End
   Begin VB.Timer TickTime 
      Enabled         =   0   'False
      Left            =   1680
      Top             =   1320
   End
   Begin VB.Image ShowDetail 
      Height          =   255
      Left            =   0
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label PrintStr 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PrintErr"
      ForeColor       =   &H8000000E&
      Height          =   180
      Left            =   3870
      TabIndex        =   0
      Top             =   1700
      Width           =   720
   End
   Begin VB.Image BackGround 
      Appearance      =   0  'Flat
      Height          =   2775
      Left            =   0
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "OutputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const LEN_PIXEL_BACKROUND_HEIGHT = 180
Private Const LEN_PIXEL_BACKROUND_WIDTH = 480

Private Const LEN_PIXEL_MOVE_DOWN = 300
Private Const LEN_PIXEL_MOVE_UP = LEN_PIXEL_MOVE_DOWN
Private Const LEN_PIXEL_MOVE_SHAKE = 40

Private Const CODE_COLOR_BACKGROUND = &HFDB313

Private Const TIME_CHANGE_WINDOW = 120
Private Const TIME_CHANGE_SHAKE = 20
Private Const TIME_CHANGE_WAIT = 20

Private Enum SHOW_STAT
OnlyMessage = 0
Detail
End Enum

Private Enum CHANGE_STAT
No = 0
Down
Up
ShakeDown
ShakeUp
End Enum

Dim ShowStr As String
Dim ShowStyle As PRINTERR_STYLE
Dim ShowStat As SHOW_STAT

Dim ChangeStat As CHANGE_STAT
Dim ChangeLen As Long
Dim ChangeIndex As Long
Dim ChangeStep As Long
Dim ChangeStepLen As Long

Sub SetShowDetail(ByVal Data As String)
AboutDetail.Text = Data
End Sub

Sub SetShowStr(ByVal ERR As String)
ShowStr = ERR
End Sub

Sub SetFormStyle(ByVal Style As PRINTERR_STYLE)
ShowStyle = Style
End Sub

Private Sub BackGround_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If MOUSE_RIGHT = Button Then PublicData.MoveForm Me
End Sub

Private Sub BackGround_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If MOUSE_LEFT = Button Then Unload Me
End Sub

Private Sub Form_Load()
If PublicData.DPI_BIG = PublicData.GetSystemDPI Then
    ResetFormSizeForHeightDPI Me, LEN_PIXEL_BACKROUND_HEIGHT, LEN_PIXEL_BACKROUND_WIDTH
    PrintStr.Top = CalcuDPIBigPixel(PrintStr.Top / DPI_SMALL_TIWP)
    PrintStr.Left = CalcuDPIBigPixel(PrintStr.Left / DPI_SMALL_TIWP)
    ShowDetail.Top = CalcuDPIBigPixel(ShowDetail.Top / DPI_SMALL_TIWP)
    ShowDetail.Width = CalcuDPIBigPixel(ShowDetail.Width / DPI_SMALL_TIWP)
    AboutDetail.Top = CalcuDPIBigPixel(AboutDetail.Top / DPI_SMALL_TIWP)
    AboutDetail.Width = CalcuDPIBigPixel(AboutDetail.Width / DPI_SMALL_TIWP)
    AboutDetail.Height = CalcuDPIBigPixel(AboutDetail.Height / DPI_SMALL_TIWP)
End If

Me.BackColor = CODE_COLOR_BACKGROUND
Me.Picture = LoadResPicture(401, vbResBitmap)
ChangeStat = No
If ShowStyle = PrintMessage Then
    ShowDetail.Visible = False
Else
    ShowDetail.Picture = LoadResPicture(402, vbResBitmap)
    ShowDetail.Visible = True
End If
PrintStr.Caption = ShowStr
End Sub

Private Sub Form_Unload(Cancel As Integer)
ShowStr = ""
PrintStr.Caption = ""
AboutDetail.Text = ""
ShowStat = OnlyMessage
End Sub

Private Sub TickTime_Timer()
If ChangeStat = Down And ChangeIndex < ChangeStep Then
    Me.Height = Me.Height + ChangeStepLen
    ShowDetail.Top = Me.Height - ShowDetail.Height

    ChangeIndex = ChangeIndex + 1
ElseIf ChangeStat = Up And ChangeIndex < ChangeStep Then
    Me.Height = Me.Height - ChangeStepLen
    ShowDetail.Top = Me.Height - ShowDetail.Height

    ChangeIndex = ChangeIndex + 1
ElseIf ChangeStat = ShakeDown And ChangeIndex < ChangeStep Then
    Me.Height = Me.Height + ChangeStepLen
    
    ChangeIndex = ChangeIndex + 1
ElseIf ChangeStat = ShakeUp And ChangeIndex < ChangeStep Then
    Me.Height = Me.Height - ChangeStepLen
    
    ChangeIndex = ChangeIndex + 1
ElseIf ChangeIndex >= ChangeStep Then
    If ChangeStat = Down Or ChangeStat = Up Then
        ChangeStat = ShakeDown
    
        ChangeLen = IIf(PublicData.DPI_BIG = PublicData.GetSystemDPI, PublicData.CalcuDPIBigPixel(LEN_PIXEL_MOVE_SHAKE), PublicData.CalcuDPISmallPixel(LEN_PIXEL_MOVE_SHAKE))
        ChangeIndex = 0
        ChangeStep = 1
        ChangeStepLen = LEN_PIXEL_MOVE_SHAKE
        
        TickTime.Interval = TIME_CHANGE_SHAKE
    ElseIf ChangeStat = ShakeDown Then
        ChangeStat = ShakeUp
        ChangeIndex = 0
    Else
        ChangeStat = No
        TickTime.Enabled = False
    End If
End If
End Sub

Private Sub ChangeWindowDown()
ChangeStat = Down
ChangeLen = IIf(PublicData.DPI_BIG = PublicData.GetSystemDPI, PublicData.CalcuDPIBigPixel(LEN_PIXEL_MOVE_DOWN), PublicData.CalcuDPISmallPixel(LEN_PIXEL_MOVE_DOWN))
ChangeIndex = 0
ChangeStep = TIME_CHANGE_WINDOW / TIME_CHANGE_WAIT
ChangeStepLen = ChangeLen / ChangeStep

TickTime.Interval = TIME_CHANGE_WAIT
TickTime.Enabled = True
End Sub

Private Sub ChangeWindowUp()
ChangeStat = Up
ChangeLen = IIf(PublicData.DPI_BIG = PublicData.GetSystemDPI, PublicData.CalcuDPIBigPixel(LEN_PIXEL_MOVE_UP), PublicData.CalcuDPISmallPixel(LEN_PIXEL_MOVE_UP))
ChangeIndex = 0
ChangeStep = TIME_CHANGE_WINDOW / TIME_CHANGE_WAIT
ChangeStepLen = ChangeLen / ChangeStep

TickTime.Interval = TIME_CHANGE_WAIT
TickTime.Enabled = True
End Sub

Private Sub ShowDetail_Click()
If ShowStat = OnlyMessage Then
    ChangeWindowDown
    ShowDetail.Picture = LoadResPicture(403, vbResBitmap)
    
    ShowStat = Detail
Else
    ChangeWindowUp
    ShowDetail.Picture = LoadResPicture(402, vbResBitmap)
    
    ShowStat = OnlyMessage
End If
End Sub
