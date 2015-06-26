Attribute VB_Name = "SaveProcessData"

Sub SaveData(ByVal Key As String, ByVal Data As String)
SaveSetting App.ProductName, "ProcessData", Key, Data
End Sub

Sub SaveDataLong(ByVal Key As String, ByVal Data As Long)
SaveSetting App.ProductName, "ProcessData", Key, Data
End Sub

Sub SaveDataBoolean(ByVal Key As String, ByVal Data As Boolean)
SaveSetting App.ProductName, "ProcessData", Key, IIf(Data, 1, 0)
End Sub

Function GetData(ByVal Key As String) As String
GetData = GetSetting(App.ProductName, "ProcessData", Key, "")
End Function

Function GetDataLong(ByVal Key As String) As Long
GetDataLong = GetSetting(App.ProductName, "ProcessData", Key, -1)
End Function

Function GetDataBoolean(ByVal Key As String) As Boolean
On Error Resume Next
GetDataBoolean = IIf(Not GetSetting(App.ProductName, "ProcessData", Key, 0) = 0, 1, 0)
End Function
