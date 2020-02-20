Imports System.ComponentModel

Namespace ShanXingTech.Net2
    ''' <summary>
    ''' 下载完成事件类
    ''' </summary>
    Public Class DownloadFileCompletedEventArgs
        Inherits AsyncCompletedEventArgs

        Public Sub New([error] As Exception, cancelled As Boolean, userState As Object)
            MyBase.New([error], cancelled, userState)
        End Sub
    End Class
End Namespace