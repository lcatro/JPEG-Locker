Attribute VB_Name = "GetTimeTick"
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim TimeTick As Long
Sub StartTimeTick()
TimeTick = GetTickCount
End Sub
Function GetCurrentTimeTick() As Long
GetCurrentTimeTick = IIf(TimeTick > 0, GetTickCount - TimeTick, 0)
End Function
