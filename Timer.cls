VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, nIDEvent As Long, ByVal Interval As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Boolean

Dim TimerID As Long
Dim WinHwnd As Long
Dim Interval As Long
Dim RunSubAddr As Long

Sub RunTimer(ByVal WindowHwnd As Long, ByVal SubAddress As Long, Optional TimeInterval As Long = 10)
TimerID = SetTimer(WindowHwnd, 0, TimeInterval, SubAddress)
WinHwnd = WindowHwnd
Interval = TimeInterval
RunSubAddr = SubAddress
End Sub

Sub ExitTimer()
KillTimer WinHwnd, TimerID
End Sub

Sub RestartTimer()
TimerID = SetTimer(WindowHwnd, TimerID, TimeInterval, RunSubAddr)
End Sub
