Imports System.Net.Http.Headers

Namespace ShanXingTech.Net2
    Public Class UploadInfo
        ''' <summary>
        ''' 请求链接
        ''' </summary>
        ''' <returns></returns>
        Public Property RequestUrl As String
        ''' <summary>
        ''' 本地文件的绝对路径
        ''' </summary>
        ''' <returns></returns>
        Public Property FileFullPath As String
        ''' <summary>
        ''' 设置 HTTP 响应上的 Content-Type 内容标头值
        ''' DispositionType 的 值必须要跟抓包的一致
        ''' ContentDispositionHeaderValue 的 name 值必须要跟抓包的一致
        ''' FileName 的值没有要求, 可以随便写, 建议写成文件的名称（带后缀）
        ''' </summary>
        ''' <returns></returns>
        Public Property ContentDisposition As ContentDispositionHeaderValue
    End Class
End Namespace