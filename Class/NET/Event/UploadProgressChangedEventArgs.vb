
Namespace ShanXingTech.Net2
    ''' <summary>
    ''' 上传进度改变事件类
    ''' </summary>
    Public Class UploadProgressChangedEventArgs
        Inherits Net.Http.Handlers.HttpProgressEventArgs

        Public Sub New(progressPercentage As Integer, userToken As Object, bytesTransferred As Long, totalBytes As Long?)
            MyBase.New(progressPercentage, userToken, bytesTransferred, totalBytes)
        End Sub
    End Class
End Namespace