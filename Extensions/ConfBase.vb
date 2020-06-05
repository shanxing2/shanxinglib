Namespace ShanXingTech
    ''' <summary>
    ''' 配置信息辅助基类。由于VB.NET自带的 <see cref="My.Settings"/> 工具处于调试模式时，可能不按照预期运行（不保存用户设置值），因此造此基类以替代。
    ''' 需要添加配置项时可往子类中添加属性，同时，可以给配置项设置默认属性（建议在子类构造函数里面延迟设置属性值）
    ''' 20191217
    ''' </summary>
    Public Class ConfBase
#Region "属性区"
        ''' <summary>
        ''' 产品名
        ''' </summary>
        ''' <returns></returns>
        Public Property ProductName As String
#End Region
    End Class
End Namespace

