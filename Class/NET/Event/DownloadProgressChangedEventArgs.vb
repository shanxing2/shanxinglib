Imports System.Net.Http.Handlers

Namespace ShanXingTech.Net2
    ''' <summary>
    ''' 下载进度改变事件类
    ''' </summary>
    Public Class DownloadProgressChangedEventArgs
        Inherits HttpProgressEventArgs

        Public Sub New(progressPercentage As Integer, userToken As Object, bytesTransferred As Long, totalBytes As Long?)
            MyBase.New(progressPercentage, userToken, bytesTransferred, totalBytes)
        End Sub
    End Class
End Namespace

