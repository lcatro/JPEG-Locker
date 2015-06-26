Attribute VB_Name = "PictureLock"

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Const RND_RANGE = 256  ''  随机数生成范围

Private Const LEN_ALLOWDATA = 256  ''  结构头偏移数据长度范围

Private Const LEN_LONG = 4  ''  Long 结构长度

Private Enum RANK  ''  排列方式
No = 0
Yes
End Enum

Private Enum FLAG_LOCK_V1_0  ''  乱写你懂的 1.0 版本加密
FLAGRANK = 4                 ''  四个随机锁标志
FLAG1 = &HFFFF0208           ''  锁标志
FLAG2 = &H6F852D0E
FLAG3 = &HC7522AD0
FLAG4 = &H291A6EF5
End Enum

Private Function GetLockFlag(ByVal Index As Long) As Long  ''  由索引号取锁标志
Select Case Index
    Case 1:
        GetLockFlag = FLAG_LOCK_V1_0.FLAG1
    Case 2:
        GetLockFlag = FLAG_LOCK_V1_0.FLAG2
    Case 3:
        GetLockFlag = FLAG_LOCK_V1_0.FLAG3
    Case 4:
        GetLockFlag = FLAG_LOCK_V1_0.FLAG4
    Case Else
        GetLockFlag = FLAG_LOCK_V1_0.FLAG1
End Select
End Function

Private Function ChackLockFlag(ByVal Number As Long) As Boolean  ''  判断锁标志
Select Case Number
    Case FLAG_LOCK_V1_0.FLAG1:
        ChackLockFlag = True
    Case FLAG_LOCK_V1_0.FLAG2:
        ChackLockFlag = True
    Case FLAG_LOCK_V1_0.FLAG3:
        ChackLockFlag = True
    Case FLAG_LOCK_V1_0.FLAG4:
        ChackLockFlag = True
    Case Else
        ChackLockFlag = False
End Select
End Function

Private Function CreateRnd(ByVal Range As Long) As Long  ''  生成一个随机数(传入取值范围,返回[0,Range) [也有可能取到Range ])
Randomize
CreateRnd = CLng(Range * Rnd)  ''  注意VB 的四舍五入取值法
End Function

Private Function CreateString(ByVal StringLength As Long) As String  ''  生成一串随机字符串(传入字符长度)
Dim rtn As String
For I = 1 To StringLength
    rtn = rtn & Chr(CreateRnd(RND_RANGE))
Next

CreateString = rtn
End Function

Private Function CreateLockFlag() As Long  ''  随机生成锁标志
CreateLockFlag = GetLockFlag(CreateRnd(FLAG_LOCK_V1_0.FLAGRANK) + 1)
End Function

Private Function GetString(ByVal Str As String, ByVal Point As Long) As String  ''  获取字符串个某个位置的符号
If Point = 0 Or Point > Len(Str) Then Exit Function

GetString = Mid(Left(Str, Point), Point)
End Function

Private Sub StringToByte(ByVal InString As String, ByVal LenString As Long, InByte() As Byte)  ''  String 数据变量转Byte
For I = 1 To LenString
InByte(I - 1) = CByte(Asc(GetString(InString, I)))
Next
End Sub

Function IsLockFile(ByVal FilePath As String) As Boolean  ''  判断文件是否已经被上锁
Dim Data(LEN_LONG - 1) As Byte

Open FilePath For Binary As #1
Get #1, , Data
Close

Dim Number As String
Open FilePath For Binary As #1
Get #1, , Data
Close
Number = Hex(Data(3))
Number = Number & IIf(Data(2) <= &HF, "0" & Hex(Data(2)), Hex(Data(2)))
Number = Number & IIf(Data(1) <= &HF, "0" & Hex(Data(1)), Hex(Data(1)))
Number = Number & IIf(Data(0) <= &HF, "0" & Hex(Data(0)), Hex(Data(0)))

IsLockFile = ChackLockFlag(Val("&H" & Number))
End Function

''LockPicture FILE_PATH_OPEN, FILE_PATH_SAVE, FILE_NUM

Sub LockPicture(ByVal FilePathOpen As String, ByVal FilePathSave As String, ByVal LockString As String)  ''  加锁图片
If IsLockFile(FilePathOpen) = True Then Exit Sub  ''  处理到已经上锁了的图片就立即退出,防止别人调试

Dim Data() As Byte  ''  源图片文件数据
Dim Exchange() As Byte  ''  加密后的图片数据
Dim SorcFileLength As Long  ''  源图片文件长度

Dim LockFlag As Long  ''  锁标志
Dim DataPoint As Long  ''  偏移数据指针
Dim AllocData As String  ''  填充数据
Dim PasswordLength As Long  ''  密码长度
Dim Password As String  ''  密码
Dim RankNum As RANK  ''  排列方式
LockFlag = CreateLockFlag()
DataPoint = CreateRnd(LEN_ALLOWDATA)  ''  随机生成填充数据长度
AllocData = CreateString(DataPoint)  ''  生成填充数据
Password = Base64Encode(LockString)  ''  加密密码
PasswordLength = Len(Password)  ''  获取加密后的密码长度
RankNum = Yes  ''  选择排列

Dim AllocPasswordBlockSize As Long
AllocPasswordBlockSize = LEN_LONG * 4 + DataPoint + PasswordLength  ''  计算密码块的大小(有三个Long 型数据和填充数据和密码数据)

Open FilePathOpen For Binary As #1  ''  读取源图片数据
    SorcFileLength = FileLen(FilePathOpen)  ''  获取源图片大小
    ReDim Data(SorcFileLength - 1)  ''  在重定义数组时,由于VB 数组的最后一位可以不是NULL 终止字符串,所以可以根据源文件长度减一使Data 里面的数据充满..
    Get #1, , Data  ''  获取数据
Close

ReDim Exchange(AllocPasswordBlockSize + SorcFileLength - 1)  ''  加密后的数据长度包含了密码块和源图片数据

CopyMemory Exchange(0), LockFlag, LEN_LONG  ''  填充密码块
CopyMemory Exchange(LEN_LONG), DataPoint, LEN_LONG
CopyMemory Exchange(LEN_LONG * 2), AllocData, DataPoint
CopyMemory Exchange(LEN_LONG * 2 + DataPoint), PasswordLength, LEN_LONG
For I = LEN_LONG * 3 + DataPoint To LEN_LONG * 3 + DataPoint + PasswordLength - 1  ''  写密码到密码块中[不知道为什么不能使用CopyMemory ,请高手指点] -- LCatro  2013.8.23
    Exchange(I) = Asc(GetString(Password, I - (LEN_LONG * 3 + DataPoint) + 1))
Next
CopyMemory Exchange(LEN_LONG * 3 + DataPoint + PasswordLength), RankNum, LEN_LONG

Dim DataCache() As Byte
Dim DataCacheLen As Long
Encode Data, SorcFileLength, DataCache, DataCacheLen  ''  打乱数据
ReDim Preserve Exchange(DataCacheLen + AllocPasswordBlockSize - 1)  ''  重新设置交换数据缓冲大小(保留原数据)

CopyMemory Exchange(AllocPasswordBlockSize), DataCache(0), DataCacheLen     ''  将源图片数据复制到加密文件缓存

Open FilePathSave For Binary As #1  ''  写数据
    Put #1, , Exchange
Close
End Sub

''UnlockPicture FILE_PATH_SAVE, FILE_PATH_SAVETEST

Function UnlockPicture(ByVal FilePathOpen As String, ByVal FilePathSave As String, ByVal UnlockString As String) As Boolean
If IsLockFile(FilePathOpen) = False Then
    UnlockPicture = False
    Exit Function
End If

Dim Data() As Byte
Dim Exchange() As Byte
Dim SorcFileLength As Long
Dim DataPoint As Long
Dim AllocData As String
Dim PasswordLength As Long
Dim Password As String
Dim RankNum As RANK

Open FilePathOpen For Binary As #1
    SorcFileLength = FileLen(FilePathOpen)
    ReDim Data(SorcFileLength - 1)
    Get #1, , Data
Close

CopyMemory DataPoint, Data(LEN_LONG), LEN_LONG  ''  读取填充数据长度
CopyMemory PasswordLength, Data(LEN_LONG * 2 + DataPoint), LEN_LONG ''  读取密码长度

For I = 0 To PasswordLength - 1  ''  密码读取
    Password = Password & Chr(Data(LEN_LONG * 3 + DataPoint + I))
Next
Password = Base64.Base64Decode(Password)  ''  密码解密

CopyMemory RankNum, Data(LEN_LONG * 3 + DataPoint + PasswordLength), LEN_LONG  ''  读取排列方式

If Not UnlockString = Password Then  ''  密码对比
    UnlockPicture = False
    Exit Function
End If

Dim AllocPasswordBlockSize As Long
AllocPasswordBlockSize = LEN_LONG * 4 + DataPoint + PasswordLength

ReDim Exchange(SorcFileLength - AllocPasswordBlockSize - 1)

CopyMemory Exchange(0), Data(AllocPasswordBlockSize), SorcFileLength - AllocPasswordBlockSize

Dim DataCache() As Byte
Dim DataCacheLen As Long
Decode Exchange, SorcFileLength - AllocPasswordBlockSize, DataCache, DataCacheLen
ReDim Exchange(DataCacheLen - 1)

CopyMemory Exchange(0), DataCache(0), DataCacheLen  ''  将源图片数据复制到加密文件缓存

Open FilePathSave For Binary As #1
    Put #1, , Exchange
Close

UnlockPicture = True
End Function
