Attribute VB_Name = "DataRank"

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Const STR_SPACE = 32

Private Function GetString(ByVal Str As String, ByVal Point As Long) As String  ''  获取字符串个某个位置的符号
If Point = 0 Or Point > Len(Str) Then Exit Function

GetString = Mid(Left(Str, Point), Point)
End Function

Sub Encode(InputData() As Byte, ByVal InputDataLen As Long, OutputData() As Byte, OutputDataLen As Long)   ''  分三段打乱数据
Dim StrLen As String
Dim Step As Long
StrLen = InputDataLen
For Step = 0 To 2
    If (StrLen Mod 3) = 0 Then
        OutputDataLen = StrLen
        ReDim OutputData(OutputDataLen)
        
        If Step = 1 Then
            OutputData((OutputDataLen / 3) * 2 - 1) = STR_SPACE
        ElseIf Step = 2 Then
            OutputData((OutputDataLen / 3) * 2 - 1) = STR_SPACE
            OutputData((OutputDataLen / 3) * 2 - 2) = STR_SPACE
        End If
        Exit For
    End If
    
    StrLen = StrLen + 1
Next

''  123->231
CopyMemory OutputData(0), InputData(OutputDataLen / 3), OutputDataLen / 3
CopyMemory OutputData(OutputDataLen / 3), InputData((OutputDataLen / 3) * 2), OutputDataLen / 3 - IIf(Step = 0, 0, IIf(Step = 1, 1, IIf(Step = 2, 2, 1)))
CopyMemory OutputData((OutputDataLen / 3) * 2), InputData(0), OutputDataLen / 3
End Sub


Sub Decode(InputData() As Byte, ByVal InputDataLen As Long, OutputData() As Byte, OutputDataLen As Long)
If Not (InputDataLen Mod 3) = 0 Then Exit Sub

''  231->123
If InputData((InputDataLen / 3) * 2 - 1) = STR_SPACE And InputData((InputDataLen / 3) * 2 - 2) = STR_SPACE Then
    OutputDataLen = InputDataLen - 2
    ReDim OutputData(OutputDataLen)
    CopyMemory OutputData((InputDataLen / 3) * 2), InputData(InputDataLen / 3), InputDataLen / 3 - 2
ElseIf InputData((InputDataLen / 3) * 2 - 1) = STR_SPACE Then
    OutputDataLen = InputDataLen - 1
    ReDim OutputData(OutputDataLen)
    CopyMemory OutputData((InputDataLen / 3) * 2), InputData(InputDataLen / 3), InputDataLen / 3 - 1
Else
    OutputDataLen = InputDataLen
    ReDim OutputData(OutputDataLen)
    CopyMemory OutputData((InputDataLen / 3) * 2), InputData(InputDataLen / 3), InputDataLen / 3
End If

CopyMemory OutputData(0), InputData((InputDataLen / 3) * 2), InputDataLen / 3
CopyMemory OutputData(InputDataLen / 3), InputData(0), InputDataLen / 3
End Sub
