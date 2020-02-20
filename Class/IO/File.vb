Imports System.IO
Imports System.Text

Namespace ShanXingTech.IO2
    Public Class File
        ''' <summary>
        ''' 获取文件的MD5值
        ''' </summary>
        ''' <param name="fileFullPath">文件完整路径，如C:\test.txt</param>
        ''' <returns></returns>
        Public Shared Function GetMD5Value(fileFullPath As String) As String
            If Not IO.File.Exists(fileFullPath) Then
                Throw New FileNotFoundException(fileFullPath)
            End If

            Dim retVal As Byte()

            Try
                Dim fileStream As New FileStream(fileFullPath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, True)
                Using md5 = New System.Security.Cryptography.MD5CryptoServiceProvider()
                    retVal = md5.ComputeHash(fileStream)
                End Using
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try

            Dim funRst = retVal.ToHexString(UpperLowerCase.Lower)

            Return funRst
        End Function
    End Class

End Namespace
