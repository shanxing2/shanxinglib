Imports System.IO
Imports System.Net
Imports System.Net.Http.Headers

Namespace ShanXingTech.Net2
    ''' <summary>
    ''' 专用于Head请求的响应
    ''' </summary>
    Public Class HeadResponse
        ''' <summary>
        ''' 是否成功
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Success As Boolean
        ''' <summary>
        ''' 请求状态
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property StatusCode As HttpStatusCode
        ''' <summary>
        ''' 响应头
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Header As HttpResponseHeaders

        Public Sub New(success As Boolean, statusCode As HttpStatusCode, header As HttpResponseHeaders)
            Me.Success = success
            Me.StatusCode = statusCode
            Me.Header = header
        End Sub
    End Class
End Namespace

