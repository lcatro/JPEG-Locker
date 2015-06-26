VERSION 5.00
Begin VB.Form Finish 
   BorderStyle     =   0  'None
   ClientHeight    =   2700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '∆¡ƒª÷––ƒ
   Begin VB.Label UseTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.000"
      ForeColor       =   &H8000000E&
      Height          =   180
      Left            =   3960
      TabIndex        =   0
      Top             =   1700
      Width           =   450
   End
End
Attribute VB_Name = "Finish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const LEN_PIXEL_BACKGROUNG_HEIGHT = 180
Private Const LEN_PIXEL_BACKGROUNG_WIDTH = 480

Dim WindowsStyle As FINISH_STYLE
Dim WaitTime As Long
Dim TheTime As Long

Sub SetShowStyle(ByVal Style As FINISH_STYLE)
WindowsStyle = Style
End Sub

Sub InputTime(ByVal MicroSecond As Long)
TheTime = MicroSecond
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If MOUSE_LEFT = Button Then Unload Me
End Sub

Private Sub Form_Load()
WaitTime = 0
TheTime = TheTime / 100
If WindowsStyle = FINISH_STYLE.lck Then
    Me.Picture = LoadResPicture(201, vbResBitmap)
    UseTime.Caption = IIf(TheTime = 0 Or TheTime >= 1, TheTime, "0" & TheTime) & "√Î"
ElseIf WindowsStyle = FINISH_STYLE.unlck Then
    Me.Picture = LoadResPicture(202, vbResBitmap)
    UseTime.Caption = IIf(TheTime = 0 Or TheTime >= 1, TheTime, "0" & TheTime) & "√Î"
End If

If PublicData.GetSystemDPI = PublicData.DPI_BIG Then
    PublicData.ResetFormSizeForHeightDPI Me, LEN_PIXEL_BACKGROUNG_HEIGHT, LEN_PIXEL_BACKGROUNG_WIDTH
    UseTime.Top = PublicData.CalcuDPIBigPixel(UseTime.Top / PublicData.DPI_SMALL_TIWP)
Else
    Me.Height = PublicData.CalcuDPISmallPixel(LEN_PIXEL_BACKGROUNG_HEIGHT)
    Me.Width = PublicData.CalcuDPISmallPixel(LEN_PIXEL_BACKGROUNG_WIDTH)
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
UseTime.Caption = ""
TheTime = 0
End Sub
