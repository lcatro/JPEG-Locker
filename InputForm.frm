VERSION 5.00
Begin VB.Form InputForm 
   BorderStyle     =   0  'None
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '∆¡ƒª÷––ƒ
   Begin VB.TextBox InputText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   1560
      TabIndex        =   0
      Top             =   1330
      Width           =   3735
   End
   Begin VB.Image CloseButton 
      Height          =   255
      Left            =   6840
      Top             =   120
      Width           =   255
   End
   Begin VB.Image Subit 
      Height          =   465
      Left            =   5400
      Top             =   1090
      Width           =   615
   End
End
Attribute VB_Name = "InputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const DISTANCE_PIXEL_TOP = 50
Private Const DISTANCE_PIXEL_BOTTOM_LABEL_TO_TEXT = 10
Private Const DISTANCE_PIXEL_BOTTOM_TEXT_TO_FORM = 10
Private Const DISTANCE_PIXEL_LEFT = 50
Private Const DISTANCE_PIXEL_RIGHT = 50

Private Const LEN_PIXEL_BACKGROUNG_HEIGHT = 180
Private Const LEN_PIXEL_BACKGROUNG_WIDTH = 480

Private Const STR_HELP = "«Î ‰»Î√‹¬Î"
Private Const STR_PASSWORD_CHAR = "*"

Private Const KEY_ENTER = 13

Dim CloseButtonStyle As STYLE_BUTTON

Private Sub CloseButton_Click()
Unload Me
End Sub

Private Sub CloseButton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If CloseButtonStyle = MOUSE_LEAVE Then
    CloseButton.Picture = LoadResPicture(110, vbResBitmap)
    CloseButtonStyle = MOUSE_ENTER
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If ((CloseButton.Left <= x Or x < CloseButton.Left + CloseButton.Width) And (CloseButton.Top <= y Or y < CloseButton.Top + CloseButton.Height)) And CloseButtonStyle = MOUSE_ENTER Then
    CloseButton.Picture = LoadResPicture(111, vbResBitmap)
    CloseButtonStyle = MOUSE_LEAVE
End If
End Sub

Private Sub Form_Load()
CloseButtonStyle = MOUSE_LEAVE
'InputText.Text = STR_HELP
InputText_Click

Me.Picture = LoadResPicture(301, vbResBitmap)
Subit.Picture = LoadResPicture(302, vbResBitmap)
CloseButton.Picture = LoadResPicture(111, vbResBitmap)

If PublicData.DPI_BIG = PublicData.GetSystemDPI Then
    ResetFormSizeForHeightDPI Me, LEN_PIXEL_BACKGROUNG_HEIGHT, LEN_PIXEL_BACKGROUNG_WIDTH
    Dim Ctrl As Control
    For Each Ctrl In Me
        Ctrl.Top = CalcuDPIBigPixel(Ctrl.Top / DPI_SMALL_TIWP)
        Ctrl.Left = CalcuDPIBigPixel(Ctrl.Left / DPI_SMALL_TIWP)
    Next
    
    InputText.Width = Subit.Left - InputText.Left
End If
End Sub

Private Sub InputText_Click()
'If InputText.Text = STR_HELP Then InputText.Text = ""

If PasswordCanSee Then
    InputText.PasswordChar = ""
Else
    InputText.PasswordChar = STR_PASSWORD_CHAR
End If
End Sub

Private Sub InputText_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = KEY_ENTER Then Subit_Click
End Sub

Private Sub Subit_Click()
If Not InputText.Text = "" Then
    PublicData.InputData = InputText.Text
Else
    PublicData.PrintErrMessage STR_HELP
End If

Unload Me
End Sub

Private Sub Cancel_Click()
PublicData.InputData = ""
Unload Me
End Sub

