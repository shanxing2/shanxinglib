Imports ShanXingTech.Win32API.UnsafeNativeMethods

Namespace ShanXingTech.Media2
    ''' <summary>
    ''' MP3相关操作类
    ''' </summary>
    Partial Public Class Mp3Player
#Region "属性区"
        Private _type As MediaType
        ''' <summary>
        ''' 多媒体文件类型
        ''' </summary>
        ''' <returns></returns>
        Public Property Type() As MediaType
            Get
                Return _type
            End Get
            Set(ByVal value As MediaType)
                _type = value
            End Set
        End Property

        Private _fileName As String
        ''' <summary>
        ''' 多媒体文件路径
        ''' </summary>
        ''' <returns></returns>
        Public Property FileName() As String
            Get
                Return _fileName
            End Get
            Set(ByVal value As String)
                _fileName = value
            End Set
        End Property

        Private _status As MciStatus
        ''' <summary>
        ''' 多文件文件播放状态
        ''' </summary>
        ''' <returns></returns>
        Public Property Status() As MciStatus
            Get
                Return _status
            End Get
            Set(ByVal value As MciStatus)
                _status = value
            End Set
        End Property
#End Region

#Region "构造函数区"
        Public Sub New()
            ' 
        End Sub

        Public Sub New(ByVal mediaFileName As String, ByVal mediaType As MediaType)
            _fileName = mediaFileName
            _type = mediaType
        End Sub
#End Region

#Region "函数区"
        ''' <summary>
        ''' 播放媒体文件
        ''' </summary>
        Private Sub Open()
            MciExcute($"open ""{_fileName}"" type {_type.ToString()} alias newmedia")
            _status = MciStatus.Open
        End Sub

        ''' <summary>
        ''' 执行 Mci 命令
        ''' </summary>
        ''' <param name="lpstrCommand"></param>
        ''' <returns></returns>
        Private Function MciExcute(ByVal lpstrCommand As String) As Integer
            Dim funcRst = MciSendString(lpstrCommand, Nothing, 0, IntPtr.Zero)

            Return funcRst
        End Function

        ''' <summary>
        ''' 播放媒体文件
        ''' </summary>
        Public Sub Play()
            Open()
            MciExcute("play newmedia")
            _status = MciStatus.Playing
        End Sub

        ''' <summary>
        ''' 重复播放媒体文件
        ''' </summary>
        Public Sub PlayRepeat()
            Open()
            MciExcute("play newmedia repeat")
            _status = MciStatus.Playing
        End Sub

        ''' <summary>
        ''' 停止播放媒体文件
        ''' </summary>
        Public Sub Close()
            MciExcute("close newmedia")
            _status = MciStatus.Close
        End Sub

        ''' <summary>
        ''' 停止设备的播放或记录
        ''' </summary>
        Public Sub [Stop]()
            MciExcute("stop newmedia")
            _status = MciStatus.Stopped
        End Sub
#End Region

    End Class

End Namespace

