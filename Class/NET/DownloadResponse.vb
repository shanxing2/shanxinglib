Imports System.IO
Imports System.Net

Namespace ShanXingTech.Net2
    ''' <summary>
    ''' 专用于下载请求的响应
    ''' </summary>
    Public Class DownloadResponse
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
        ''' 下载请求返回的内存流
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property MemoryStream As MemoryStream

        Public Sub New(success As Boolean, statusCode As HttpStatusCode, memoryStream As MemoryStream)
            Me.Success = success
            Me.StatusCode = statusCode
            Me.MemoryStream = memoryStream
        End Sub
    End Class
End Namespace

