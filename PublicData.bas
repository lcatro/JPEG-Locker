Attribute VB_Name = "PublicData"

Public Const MOUSE_LEFT = 1
Public Const MOUSE_RIGHT = 2

Enum PRINTERR_STYLE
PrintMessage = 0
PrintList
End Enum

Enum FINISH_STYLE
lck = 0
unlck
End Enum

Enum STYLE_BUTTON
MOUSE_ENTER = 0
MOUSE_LEAVE
End Enum

Enum PRO_STATE
WaitSelect = 0
IsLock
IsUnlock
End Enum

Public Stat As PRO_STATE
Public SavePath As String
Public DeleteSorc As Boolean
Public BadExit As Boolean
Public PasswordCanSee As Boolean

Public InputData As String

Public Const DPI_BIG = 120
Public Const DPI_BIG_TIWP = 12
Public Const DPI_SMALL = 96
Public Const DPI_SMALL_TIWP = 15

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Const LOGPIXELSX = 88

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub MoveForm(ByVal Frm As Form)
ReleaseCapture
SendMessage Frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Function GetSystemDPI() As Long
GetSystemDPI = GetDeviceCaps(GetDC(0), LOGPIXELSX)
End Function

Function CalcuDPIBigPixel(ByVal Pixel As Long) As Long
CalcuDPIBigPixel = Pixel * DPI_BIG_TIWP
End Function

Function CalcuDPISmallPixel(ByVal Pixel As Long) As Long
CalcuDPISmallPixel = Pixel * DPI_SMALL_TIWP
End Function

Sub ResetFormSizeForHeightDPI(ByVal Frm As Form, ByVal BackgroundHeight_PIXEL As Long, ByVal BackgroundWidth_PIXEL As Long)
Frm.Height = CalcuDPIBigPixel(BackgroundHeight_PIXEL)
Frm.Width = CalcuDPIBigPixel(BackgroundWidth_PIXEL)
End Sub

Sub PrintUseTime(ByVal UseTime As Long, Optional ByVal Style As FINISH_STYLE = lck)
Finish.SetShowStyle Style
Finish.InputTime UseTime
Finish.Show 1
End Sub

Sub PrintErrMessage(ByVal ErrData As String)
OutputForm.SetFormStyle PrintMessage
OutputForm.SetShowStr ErrData
OutputForm.Show 1
End Sub
 
Sub PrintErrDetail(ByVal ErrData As String, ByVal ErrDetail As String)
OutputForm.SetFormStyle PrintList
OutputForm.SetShowStr ErrData
OutputForm.SetShowDetail ErrDetail
OutputForm.Show 1
End Sub

Function GetPassword() As String
InputData = ""

InputForm.Show 1

GetPassword = InputData
End Function

