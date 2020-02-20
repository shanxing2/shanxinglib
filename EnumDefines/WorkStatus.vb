Namespace ShanXingTech
    ''' <summary>
    ''' 工作状态
    ''' </summary>
    Public Enum WorkStatus
        [Start]
        [Pause]
        [Continue]
        Ending
        [End]
        [Init]
        DataOutPut
        DataOutPutting
        DataOutPuted
        DataImportting
        DataImportted
        ''' <summary>
        ''' 正在更新数据库中的数据
        ''' </summary>
        DatabaseUpdating
        ''' <summary>
        ''' 已更新数据库中的数据
        ''' </summary>
        DatabaseUpdated
        ''' <summary>
        ''' 文件扫描中
        ''' </summary>
        Scanning
        ''' <summary>
        ''' 文件扫描结束
        ''' </summary>
        Scanned
        ''' <summary>
        ''' 采集前
        ''' </summary>
        Colltect
        ''' <summary>
        ''' 采集后
        ''' </summary>
        Colltected
        ''' <summary>
        ''' 数据保存到数据库之后
        ''' </summary>
        DataSaved
        ''' <summary>
        ''' 登录成功
        ''' </summary>
        Logined
        ''' <summary>
        ''' 搜索中
        ''' </summary>
        Searching
        ''' <summary>
        ''' 搜索完毕
        ''' </summary>
        Searched
        ''' <summary>
        ''' 进行中
        ''' </summary>
        Doing
        ''' <summary>
        ''' 已完成
        ''' </summary>
        Done
    End Enum
End Namespace


