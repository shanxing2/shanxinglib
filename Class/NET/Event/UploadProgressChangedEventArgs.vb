
Namespace ShanXingTech.Net2
    ''' <summary>
    ''' 上传进度改变事件类
    ''' </summary>
    Public Class UploadProgressChangedEventArgs
        Inherits Net.Http.Handlers.HttpProgressEventArgs

        Public Sub New(progressPercentage As Integer, userState As Object, bytesTransferred As Long, totalBytes As Long?)
            MyBase.New(progressPercentage, userState, bytesTransferred, totalBytes)
        End Sub
    End Class
End Namespace