Imports System.Net.Http
Imports System.Net.Http.Headers

Namespace ShanXingTech.Net2
    Public Class UploadInfo
        ''' <summary>
        ''' 请求链接
        ''' </summary>
        ''' <returns></returns>
        Public Property RequestUrl As String
        ''' <summary>
        ''' 设置 HTTP 请求的数据源
        ''' </summary>
        ''' <returns></returns>
        Public Property HttpContents As List(Of HttpContentDetail)

        ''' <summary>
        ''' 字符串参数实例 ：New HttpContentDetail With {.Content = New StringContent("pic_common_upload"), .Name = "api"}
        ''' 文件参数实例 ：New HttpContentDetail With {.Content = New ByteArrayContent(IO.File.ReadAllBytes(fileFullPath)), .Name = "filedata", .FileName = fileName}
        ''' </summary>
        Public Class HttpContentDetail
            Public Property Content As HttpContent
            ''' <summary>
            ''' <see cref="Content"/> 的 Name，需要注意大小写，建议是和抓包一致。
            ''' </summary>
            ''' <returns></returns>
            Public Property Name As String
            ''' <summary>
            ''' 文件名。如果服务器返回获取不到文件，则必须传入文件名。<see cref="StringContent"/> 不需要设置 此属性
            ''' </summary>
            ''' <returns></returns>
            Public Property FileName As String
        End Class
    End Class
End Namespace