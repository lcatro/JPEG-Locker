Attribute VB_Name = "MainProc"

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Const PROCESS_ALL_ACCESS = &H1F0FFF

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Const WAIT_BLOCK = -1

Sub Main()
On Error GoTo ERR
Dim Argv() As String
Argv = Split(Command(), " ")

If Not UBound(Argv) = 2 Then
    MsgBox "传入参数有误:" & UBound(Argv), vbCritical
    End
End If

WaitForSingleObject OpenProcess(PROCESS_ALL_ACCESS, 0, CLng(Argv(2))), WAIT_BLOCK

Kill Argv(0) & "\" & Argv(1) & ".exe"

Name Argv(0) & "\Update.exe" As Argv(0) & "\" & Argv(1) & ".exe"

Shell Argv(0) & "\" & Argv(1) & ".exe", vbNormalFocus
Exit Sub
ERR:
MsgBox "0;" & Argv(0) & vbCrLf & "1;" & Argv(1) & vbCrLf & "2;" & Argv(2) & vbCrLf & vbCrLf & "错误号:" & ERR.Number & "(" & ERR.Description & ")"
End Sub
