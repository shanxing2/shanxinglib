Namespace ShanXingTech.Net2
    ''' <summary>
    ''' 上传完成事件类
    ''' </summary>
    Public Class UploadFileCompletedEventArgs
        Inherits ComponentModel.AsyncCompletedEventArgs

        Public Sub New([error] As Exception, cancelled As Boolean, userState As Object)
            MyBase.New([error], cancelled, userState)
        End Sub
    End Class
End Namespace