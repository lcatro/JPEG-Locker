Attribute VB_Name = "RAR"
Option Explicit
'
'
'******************************************************************
'模块用途：
'   解压 RAR 格式的压缩包。
'
'
'******************************************************************
'注：
'   1. 要使用本模块，必须附带 UnRAR.dll 文件。当前版本为：3.90.100.227
'   2. 本模块改编自 WinRAR 官方下载的 UnRARDLL.exe 压缩包里面的例子中的 VBasic Sample 1。
'       该 UnRARDLL.rar 是在 http://www.rarlab.com/rar_add.htm 中下载 UnRAR.dll (UnRAR dynamic library for Windows software developers.)
'       直接下载链接：http://www.rarlab.com/rar/UnRARDLL.exe    (在 2009-10-20 测试有效)
'   3. 关于 UnRAR.dll 的导出函数释义，
'           看这里有更加详细的中文说明，甚至有Unicode版本的导出函数：http://baike.baidu.com/view/697654.htm，
'           比中文版本说明更详细的原版的英文说明，见 UnRARDLL.exe 压缩包中的 unrardll.txt 文件。
'
'
'******************************************************************
'用法示例：
'    '要把经过加密的RAR包 C:\压缩样本.rar 文件解压到 C:\Temp 目录下，密码是123456
'    lngResult = RARExecute(OP_EXTRACT, "C:\压缩样本.rar", "C:\Temp\", "123456")
'    If lngResult = 0 Then
'        MsgBox "解压成功"
'    Else
'        MsgBox "解压失败，返回的错误代码是[" & lngResult & "]"
'    End If
'
'用法参数说明：
'   Mode        - 运行模式：（OP_EXTRACT = 解压；OP_TEST = 测试；OP_LIST = 查看）
'   RarFile     - RAR文件名
'   SaveTo      - 解压目录，为空表示在当前目录解压。
'   Password    - 密码
'
'******************************************************************
'修改者信息：
'   修改者：ZhongWei
'   QQ：1124091881
'   Email：1124091881@qq.com
'   修改时间：2009-10-21
'
'******************************************************************
'   此处为官方 UnRARDLL.exe 压缩包中的 VBasic Sample 1 的作者信息。
'
'   Ported to Visual Basic by Pedro Lamas
'
'E-mail:  sniper@hotpop.com
'HomePage (dedicated to VB):  www.terravista.pt/portosanto/3723/
'
'******************************************************************

'--------------------------------------------------------------------------------------------
'       常量定义
'--------------------------------------------------------------------------------------------

Const ERAR_END_ARCHIVE = 10
Const ERAR_NO_MEMORY = 11           '内存不足。
Const ERAR_BAD_DATA = 12            '数据错误，可能是压缩包文件头丢失，或者文件CRC校验错误。
Const ERAR_BAD_ARCHIVE = 13         '卷错误。
Const ERAR_UNKNOWN_FORMAT = 14      '未知的压缩包格式。
Const ERAR_EOPEN = 15               '打开卷失败。
Const ERAR_ECREATE = 16             '创建文件失败。
Const ERAR_ECLOSE = 17              '关闭文件失败。
Const ERAR_EREAD = 18               '读错误。
Const ERAR_EWRITE = 19              '写错误。
Const ERAR_SMALL_BUF = 20           '缓冲区过小。
 
Const RAR_OM_LIST = 0       '查看操作，仅用于 RAROpenArchiveData 结构。只为读取文件头而打开压缩包。
Const RAR_OM_EXTRACT = 1    '解压操作，仅用于 RAROpenArchiveData 结构。为检测或者解压缩而打开压缩包。
 
Const RAR_SKIP = 0      '跳过，仅用于 RARProcessFile 的 Operation 参数。
Const RAR_TEST = 1      '测试，仅用于 RARProcessFile 的 Operation 参数。
Const RAR_EXTRACT = 2   '解压，仅用于 RARProcessFile 的 Operation 参数。
 
Const RAR_VOL_ASK = 0
Const RAR_VOL_NOTIFY = 1

Enum RarOperations
    OP_EXTRACT = 0      '解压
    OP_TEST = 1         '测试
    OP_LIST = 2         '查看
End Enum
 
' Flags 标志可取值：
'        0x01 - file continued from previous volume 前述卷的继续。
'        0x02 - file continued on next volume 下一个卷还有该文件的部分
'        0x04 - file encrypted with password 文件已加密
'        0x08 - file comment present 文件存在注释
'        0x10 - compression of previous files is used (solid flag) 此文件压缩同前面的文件有关（固实标志）
'                  Bits   7 6 5
'                         0 0 0 - 目录大小为 64 Kb
'                         0 0 1 - 目录大小为 128 Kb
'                         0 1 0 - 目录大小为 256 Kb
'                         0 1 1 - 目录大小为 512 Kb
'                         1 0 0 - 目录大小为 1024 Kb
'                         1 0 1 - 目录大小为 2048 KB
'                         1 1 0 - 目录大小为 4096 KB
'                         1 1 1 - 文件就是目录
'                         其余字节保留

Private Type RARHeaderData
    ArcName As String * 260     '输出压缩文件名，以0结束的字符串。 也可以是当前卷名称。
    FileName As String * 260    '目录名或者文件名，包含压缩包内的路径，以0结束的字符串，以OEM (DOS)编码方式给出。
    Flags As Long               '输出文件标志。
    PackSize As Long            '输出压缩文件的分包大小或者文件切割大小。
    UnpSize As Long             '解压后的文件大小。
    HostOS As Long              '压缩文件的宿主操作系统。0 - MS DOS；1 - OS/2；2 - Win32；3 - Unix。
    FileCRC As Long             '压缩之前文件的CRC值。如果文件被分割到不同的卷中，将不会在卷中给出。（后面这段话意思好像是说，如果你将一个文件压缩到多个包中，每个分卷包不会存放部分文件的CRC。我试验切割一个文件到几个卷，然后将其中的几个卷拷贝到其他目录再使用WinRar打开，发现其中CRC值不同。）输出压缩之前文件的CRC值。如果文件被分割到不同的卷中，将不会在卷中给出。（后面这段话意思好像是说，如果你将一个文件压缩到多个包中，每个分卷包不会存放部分文件的CRC。我试验切割一个文件到几个卷，然后将其中的几个卷拷贝到其他目录再使用WinRar打开，发现其中CRC值不同。）
    FileTime As Long            '按照MS DOS格式输出的日期和时间。
    UnpVer As Long              '解压需要的Rar版本。按照10 * Major version + minor version格式给出。
    Method As Long              '压缩方式。
    FileAttr As Long            '文件属性。
    CmtBuf As String            '文件注释缓冲区，(据说)，在这个版本的Dll还没有实现，CmtState 始终为0。
    CmtBufSize As Long          '注释的缓冲区大小。最大的注释长度为64KB。
    CmtSize As Long             '读取到缓冲区的实际注释大小，不能超过CmtBufSize。
    CmtState As Long            '注释状态。见 RAROpenArchiveData 结构的说明。
End Type

Private Type RAROpenArchiveData
    ArcName As String           '压缩包文件名，全路径，以'\0'作为结尾的字符串。
    OpenMode As Long            '操作类型，可用 RAR_OM_LIST 或者  RAR_OM_EXTRACT 常量。
    OpenResult As Long          '打开文件的结果，返回的是错误代码。0 - 打开成功，无错误。
    CmtBuf As String            '指向一个用来存放注释的缓冲区。最大的注释长度为64KB。注释是以0结尾的字符串。如果注释文本的长度超过缓冲区大小，注释文本将被截断。如果 CmtBuf 为 null，将不会读取注释。
    CmtBufSize As Long          '注释的缓冲区大小。
    CmtSize As Long             '读取到缓冲区的实际注释大小，不能超过CmtBufSize。
    CmtState As Long            '注释状态：1-有注释。ERAR_NO_MEMORY - 内存不足，无法释放注释。ERAR_BAD_DATA - 注释损坏。ERAR_UNKNOWN_FORMAT - 注释格式无效。ERAR_SMALL_BUF - 缓冲区过小，无法读取全部注释。
End Type


'--------------------------------------------------------------------------------------------
'       API 定义
'--------------------------------------------------------------------------------------------

'-----------------------
'作用：
'   打开Rar文件并为使用的结构体分配空间
'参数：
'   ArchiveData     - 指向 RAROpenArchiveData 这个结构体。
'返回值：
'   返回压缩包文件的 handle ,出错时返回 null
Private Declare Function RAROpenArchive Lib "unrar.dll" (ByRef ArchiveData As RAROpenArchiveData) As Long

'-----------------------
'作用：
'   关闭打开的压缩包并释放分配的内存。只有当处理压缩文件的过程结束后才可以调用这个过程，如果处理压缩文件的过程只是停止，使用这个过程将会引起错误。
'参数：
'   hArcData     - 这个参数存放从 RAROpenArchive 函数获得的压缩包文件的句柄。
'返回值：
'   0           - 成功。
'   ERAR_ECLOSE - 关闭压缩文件时发生错误。
Private Declare Function RARCloseArchive Lib "unrar.dll" (ByVal hArcData As Long) As Long

'-----------------------
'作用：
'   读取压缩包的头部。
'参数：
'   hArcData    - 这个参数存放从 RAROpenArchive 函数获得的压缩包文件的句柄。
'   HeaderData  - 指向 RARHeaderData 结构。
'返回值：
'   0                   - 成功。
'   ERAR_END_ARCHIVE    - 文档结束。End of archive
'   ERAR_BAD_DATA       - 文件头损坏。File header broken
Private Declare Function RARReadHeader Lib "unrar.dll" (ByVal hArcData As Long, ByRef HeaderData As RARHeaderData) As Long

'-----------------------
'作用：
'   执行动作，然后指向下一个文件。
'   执行时，将会根据 RAR_OM_EXTRACT 确定释放还是测试当前文件。
'   如果设置了 RAR_OM_LIST 给出模式，那么调用这个函数将会忽略当前文件直接指向下一个文件。
'参数：
'   hArcData    - 这个参数存放从 RAROpenArchive 函数获得的压缩包文件的句柄。
'   Operation   - 文件操作。有以下的选择：
'                　　RAR_SKIP     - 指向压缩包中的下一个文件。如果压缩包是固定， 并且RAR_OM_EXTRACT 已经设置，那么会处理当前文件 ---操作比简单的查找要慢。
'                　　RAR_TEST     - 检测当前文件，然后移动到压缩包中的下一个文件。如果 RAR_OM_LIST 已经设置了打开模式，那么操作同RAR_SKIP一样?
'                　　RAR_EXTRACT  - 解压当前文件，然后指向下一个文件，如果
'                　　RAR_OM_LIST  - 已经设置了打开模式，那么操作同RAR_SKIP一样。
'   DestPath    - 解压文件的目录，这是一个以0结尾的字符串。如果 DestPath 为 null，表示解压到当前目录下。只有 DestName 为null时，这个参数才有意义。
'   DestName    - 指向一个包含完整路径和名称的以0结尾的字符串，默认为null。如果 DestName 有定义（也就是不是 Null）将会用它来替换压缩包中的原始文件名和路径。
'返回值：
'    0                      - 成功。
'    ERAR_BAD_DATA          - 文件CRC错误
'    ERAR_BAD_ARCHIVE       - 卷不是有效的Rar文件
'    ERAR_UNKNOWN_FORMAT    - 未知的格式
'    ERAR_EOPEN             - 卷打开错误
'    ERAR_ECREATE           - 文件建立错误
'    ERAR_ECLOSE            - 文件关闭错误
'    ERAR_EREAD             - 读取错误
'    ERAR_EWRITE            - 写入错误
'注意：
'   如果你希望放弃解当前的解压缩操作，请在处理 UCM_PROCESSDATA 回调函数，返回-1。
Private Declare Function RARProcessFile Lib "unrar.dll" (ByVal hArcData As Long, ByVal Operation As Long, ByVal DestPath As String, ByVal DestName As String) As Long

'-----------------------
'作用：
'   给未加密的压缩包上设置一个密码。
'参数：
'   hArcData    - 这个参数存放从 RAROpenArchive 函数获得的压缩包文件的句柄。
'   Password    - 密码字符串，以 vbNull 为结尾。
Private Declare Sub RARSetPassword Lib "unrar.dll" (ByVal hArcData As Long, ByVal Password As String)

'-----------------------
'作用：
'   绝对函数，使用 RARSetCallback 函数替换。
'   RARSetCallback 函数是一个回调函数。见在 http://baike.baidu.com/view/697654.htm 的说明。
Private Declare Sub RARSetChangeVolProc Lib "unrar.dll" (ByVal hArcData As Long, ByVal Mode As Long)

'-----------------------
'作用：
'   返回 API 版本。
'注意：
'    返回当前UnRar.DLL中API的版本，在 unrar.h 中由 RAR_DLL_VERSION 定义。只有当 UnRar.DLL中的API升级时，才会提高版本号。不要将这个版本同UnRar.Dll的编译版本弄混，编译版本在每一次编译的时候都会变化。
'    如果 RARGetDllVersion() 返回值低于你软件需要的版本，就表示你使用的DLL版本太低。
'    在老的Unrar.dll中没有提供这个功能，所以最好在使用时要先用LoadLibrary 和 GetProcAddress 检查一下是否有这个功能。
'Private Declare Sub RARGetDllVersion Lib "unrar.dll" ()



'--------------------------------------------------------------------------------------------
'       函数 定义
'--------------------------------------------------------------------------------------------

' 作用：
'   从RAR文件中解压文件。
' 参数：
'   Mode        = 对 RAR 文档的操作类型
'   RARFile     = RAR 文件名
'   SaveTo      = 解压后的保存位置。
'   sPassword   = 解压密码 (可选)
' 返回值：
'   0   - 操作完成，无错误。
'   111 - 打开文件错误，内存不足。(ERAR_NO_MEMORY)
'   112 - 打开文件错误，压缩包文件头丢失。(ERAR_BAD_DATA)
'   113 - 打开文件错误，不是一个有效的RAR压缩包。(ERAR_BAD_ARCHIVE)
'   115 - 打开文件错误，无法打开文件。(ERAR_EOPEN)
'   212 - 操作过程中出错，CRC校验错误。(ERAR_BAD_DATA)
'   213 - 操作过程中出错，卷错误。(ERAR_BAD_ARCHIVE)
'   214 - 操作过程中出错，未知的压缩包格式。(ERAR_UNKNOWN_FORMAT)
'   215 - 操作过程中出错，打开卷失败。(ERAR_EOPEN)
'   216 - 操作过程中出错，创建文件失败。(ERAR_ECREATE)
'   217 - 操作过程中出错，文件关闭失败。(ERAR_ECLOSE)
'   218 - 操作过程中出错，读错误。(ERAR_EREAD)
'   219 - 操作过程中出错，写错误。(ERAR_EWRITE)
Public Function RARExecute(ByVal Mode As RarOperations, ByVal RarFile As String, ByVal SaveTo As String, Optional Password As String) As Long
    Dim lHandle As Long
    Dim iStatus As Integer
    Dim uRAR As RAROpenArchiveData
    Dim uHeader As RARHeaderData
    Dim sStat As String, Ret As Long
     
    uRAR.ArcName = RarFile
    uRAR.CmtBuf = Space(16384)      '初始化注释的缓冲区
    uRAR.CmtBufSize = 16384
    
    '设置操作类型
    If Mode = OP_LIST Then
        uRAR.OpenMode = RAR_OM_LIST
    Else
        uRAR.OpenMode = RAR_OM_EXTRACT
    End If
    
    '打开压缩包
    lHandle = RAROpenArchive(uRAR)
    If uRAR.OpenResult <> 0 Then        '打开RAR文件失败
        '在这里，uRAR.OpenResult 返回的可能有如下的几个错误：
        '   11     - 打开文件错误，内存不足。(ERAR_NO_MEMORY)
        '   12     - 打开文件错误，压缩包文件头丢失。(ERAR_BAD_DATA)
        '   13     - 打开文件错误，不是一个有效的RAR压缩包。(ERAR_BAD_ARCHIVE)
        '   15     - 打开文件错误，无法打开文件。(ERAR_EOPEN)
        RARExecute = 100 + uRAR.OpenResult
        Exit Function
    End If
 
    If Password <> "" Then RARSetPassword lHandle, Password     '不为空的时候，设置RAR的密码。
    
    '若注释存在则显示RAR注释
    'If (uRAR.CmtState = 1) Then MsgBox uRAR.CmtBuf, vbApplicationModal + vbInformation, "注释"
    
    '循环显示压缩包内的每个文件。
    iStatus = 0       '赋个零值，以便进入循环
    Do Until iStatus <> 0
        '读入压缩包内的文件头
        iStatus = RARReadHeader(lHandle, uHeader)
        
        If iStatus = ERAR_BAD_DATA Then     '读取文件头错误
            RARExecute = 212
            Exit Function
        End If
        
        sStat = Left(uHeader.FileName, InStr(1, uHeader.FileName, vbNullChar) - 1)  '压缩包中的每一项的名字(目录名或文件名，包含路径)
        
        '根据不同的操作方式，对文件进行处理
        Select Case Mode
            Case RarOperations.OP_EXTRACT
                'Ret = RARProcessFile(lHandle, RAR_EXTRACT, "", uHeader.FileName)
                Ret = RARProcessFile(lHandle, RAR_EXTRACT, SaveTo, "")
            Case RarOperations.OP_TEST
                Ret = RARProcessFile(lHandle, RAR_TEST, "", uHeader.FileName)
            Case RarOperations.OP_LIST
                Ret = RARProcessFile(lHandle, RAR_SKIP, "", "")
        End Select
        
        If Ret > 0 Then
            '操作失败。在这里，Ret 返回的可能有如下的几个错误：
            '   12 - CRC校验错误。(ERAR_BAD_DATA)
            '   13 - 卷错误。(ERAR_BAD_ARCHIVE)
            '   14 - 未知的压缩包格式。(ERAR_UNKNOWN_FORMAT)
            '   15 - 打开卷失败。(ERAR_EOPEN)
            '   16 - 创建文件失败。(ERAR_ECREATE)
            '   17 - 文件关闭失败。(ERAR_ECLOSE)
            '   18 - 读错误。(ERAR_EREAD)
            '   19 - 写错误。(ERAR_EWRITE)
            RARExecute = 200 + Ret
            Exit Function
        End If
        
        'iStatus = RARReadHeader(lHandle, uHeader)   '读入下一个文件的文件头
        'Refresh        '用于在界面中刷新显示
    Loop
    
    '关闭压缩包句柄
    RARCloseArchive lHandle
    
    RARExecute = 0
End Function

