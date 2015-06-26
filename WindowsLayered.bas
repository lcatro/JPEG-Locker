Attribute VB_Name = "WindowsLayered"
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1

Public Const AW_HOR_POSITIVE = &H1         'Animates   the   window   from   left   to   right.   This   flag   can   be   used   with   roll   or   slide   animation.
Public Const AW_HOR_NEGATIVE = &H2         'Animates   the   window   from   right   to   left.   This   flag   can   be   used   with   roll   or   slide   animation.
Public Const AW_VER_POSITIVE = &H4         'Animates   the   window   from   top   to   bottom.   This   flag   can   be   used   with   roll   or   slide   animation.
Public Const AW_VER_NEGATIVE = &H8         'Animates   the   window   from   bottom   to   top.   This   flag   can   be   used   with   roll   or   slide   animation.
Public Const AW_CENTER = &H10         'Makes   the   window   appear   to   collapse   inward   if   AW_HIDE   is   used   or   expand   outward   if   the   AW_HIDE   is   not   used.
Public Const AW_HIDE = &H10000         'Hides   the   window.   By   default,   the   window   is   shown.
Public Const AW_ACTIVATE = &H20000         'Activates   the   window.
Public Const AW_SLIDE = &H40000         'Uses   slide   animation.   By   default,   roll   animation   is   used.
Public Const AW_BLEND = &H80000         'Uses   a   fade   effect.   This   flag   can   be   used   only   if   hwnd   is   a   top-level   window.

Private Declare Function AnimateWindow Lib "user32 " (ByVal hwnd As Long, _
                                                            ByVal dwTime As Long, _
                                                            ByVal dwFlags As Long) As Boolean
'                     hwnd,         //   要进行特效显示的窗体的句柄
'                     dwTime,     //   动画持续时间，以毫秒为单位
'                     dwFlags     //   动画类型

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

''  Windows 的窗口透明度设置与窗口淡出淡入
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

''  Windows 的不规则窗口

Private Type GdiplusStartupInput
  GdiplusVersion As Long ' Must be 1 for GDI+ v1.0, the current version as of this writing.
  DebugEventCallback As Long ' Ignored on free builds
  SuppressBackgroundThread As Long ' FALSE unless you're prepared to call
  ' the hook/unhook functions properly
  SuppressExternalCodecs As Long ' FALSE unless you want GDI+ only to use
  ' its internal image codecs.
End Type
Private Enum GpStatus ' aka Status
  Ok = 0
  GenericError = 1
  InvalidParameter = 2
  OutOfMemory = 3
  ObjectBusy = 4
  InsufficientBuffer = 5
  NotImplemented = 6
  Win32Error = 7
  WrongState = 8
  Aborted = 9
  FileNotFound = 10
  ValueOverflow = 11
  AccessDenied = 12
  UnknownImageFormat = 13
  FontFamilyNotFound = 14
  FontStyleNotFound = 15
  NotTrueTypeFont = 16
  UnsupportedGdiplusVersion = 17
  GdiplusNotInitialized = 18
  PropertyNotFound = 19
  PropertyNotSupported = 20
End Enum
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As GpStatus
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As GpStatus
Private Declare Function GdipDrawImage Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal x As Single, ByVal y As Single) As GpStatus
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, Graphics As Long) As GpStatus
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal Graphics As Long) As GpStatus
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As String, Image As Long) As GpStatus
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As GpStatus
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As GpStatus
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As GpStatus

Dim GDIToken As Long


Sub SetWindowsLayered(ByVal WindowHwnd As Long, ByVal Value As Long, ByVal Flag As Long)
        Dim rtn     As Long
        rtn = GetWindowLong(WindowHwnd, GWL_EXSTYLE)
        rtn = rtn Or WS_EX_LAYERED
        SetWindowLong WindowHwnd, GWL_EXSTYLE, rtn
        SetLayeredWindowAttributes WindowHwnd, 0, Value, Flag
End Sub
Sub AnimateWindowIn(ByVal WindowsHwnd As Long, Optional RunTime As Long = 500)
AnimateWindow WindowsHwnd, RunTime, AW_BLEND Or AW_ACTIVATE
End Sub
Sub AnimateWindowOut(ByVal WindowsHwnd As Long, Optional RunTime As Long = 500)
AnimateWindow WindowsHwnd, RunTime, AW_BLEND Or AW_HIDE
End Sub

Function LoadImagePNG(ByVal WindowsHDC As Long, ByVal ImagePath As String) As Boolean
Dim Graphics As Long
If GdipCreateFromHDC(WindowsHDC, Graphics) = Ok Then
Dim Img As Long
GdipLoadImageFromFile StrConv(ImagePath, vbUnicode), Img
If GdipDrawImage(Graphics, Img, 0, 0) = Ok Then
GdipDisposeImage Img
GdipDeleteGraphics Graphics
LoadImagePNG = True
Exit Function
End If
End If

GdipDisposeImage Img
GdipDeleteGraphics Graphics
LoadImagePNG = False
End Function

Sub GDIInit()
Dim GpInput As GdiplusStartupInput
GpInput.GdiplusVersion = 1
GdiplusStartup GDIToken, GpInput
End Sub

Sub GDIClean()
GdiplusShutdown GDIToken
End Sub

Function GetImageSize(ByVal ImagePath As String, Width As Long, Height As Long) As Boolean
Dim Img As Long
If GdipLoadImageFromFile(StrConv(ImagePath, vbUnicode), Img) = Ok Then
GdipGetImageWidth Img, Width
GdipGetImageHeight Img, Height
GetImageSize = True
Exit Function
End If
Width = 0
Height = 0

GetImageSize = False
End Function

Sub RoundRectRgnWindow(ByVal Win As Form, Optional Value As Long = 20)
Dim Rgn As Long
Rgn = CreateRoundRectRgn(0, 0, Win.ScaleX(Win.Width, vbTwips, vbPixels), Win.ScaleY(Win.Height, vbTwips, vbPixels), Value, Value)
SetWindowRgn Win.hwnd, Rgn, True
DeleteObject Rgn
End Sub






