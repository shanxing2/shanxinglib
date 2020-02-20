Imports ShanXingTech.Windows2

Namespace ShanXingTech.IO2
    Public Class Directory
        ''' <summary>
        ''' 创建文件夹,如果该文件夹已存在，则直接返回，不引发任何异常。
        ''' 成功返回true,失败返回false
        ''' </summary>
        ''' <param name="fileName">可以是具体的文件路径，也可以是文件夹路径</param>
        ''' <returns></returns>
        Public Shared Function Create(ByVal fileName As String) As Boolean
            If fileName.IsNullOrEmpty Then
                Throw New ArgumentException(NameOf(fileName))
            End If

            Dim funcRst As Boolean

            ' 创建文件夹
            ' 如果该文件夹已存在，直接再创建不会引发文件夹已存在异常。
            ' 如果 获取到的后缀名为空，说明是一个文件夹路径
            ' 否则就需要获取文件夹名
            Dim path = fileName
            Dim extension = System.IO.Path.GetExtension(fileName)
            If extension.Length > 0 Then
                path = System.IO.Path.GetDirectoryName(fileName)
            End If

            ' 如果文件夹已经存在则不需要再次创建
            If System.IO.Directory.Exists(path) Then
                Return True
            End If

            System.IO.Directory.CreateDirectory(path)

            Do Until System.IO.Directory.Exists(path)
                '
            Loop

            funcRst = True

            Return funcRst
        End Function

        ''' <summary>
        ''' 从本地物理路径删除文件夹
        ''' </summary>
        ''' <param name="directory"></param>
        ''' <returns>删除成功返回True,找不到文件夹或者删除失败返回False</returns>
        Public Shared Function Delete(ByVal directory As String) As Boolean
            Dim funcRst As Boolean

            Try
                If System.IO.Directory.Exists(directory) Then
                    System.IO.Directory.Delete(directory, True)
                    ' DeleteDirectory方法好像是异步的 需要等待清空完成 然后再新建
                    While System.IO.Directory.Exists(directory)
                       Windows2.Delay(1)
                    End While
                    Debug.Print(Logger.MakeDebugString("删除本地文件夹成功 " & directory))

                    funcRst = True
                Else
                    Debug.Print(Logger.MakeDebugString("删除本地文件夹失败或者找不到文件夹 " & directory))
                End If
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try

            Return funcRst
        End Function

        ''' <summary>
        ''' 获取某目录下的所有文件
        ''' 可以避免权限拒绝的异常
        ''' </summary>
        ''' <param name="filePath"></param>
        ''' <param name="filter">文件筛选器</param>
        ''' <param name="fileNameMode">指明返回的数据是文件的绝对路径还是只包含文件名以及后缀的路径</param>
        ''' <returns></returns>
        Public Shared Function GetFiles(ByVal filePath As String,
                                         ByVal filter As String(),
                                         ByVal fileNameMode As FileNameMode) As List(Of String)

            Dim files As New List(Of String)(64)

            Try

                Dim directory = My.Computer.FileSystem.GetDirectoryInfo(filePath)

                If fileNameMode = FileNameMode.FileName Then
#Region "只获取文件名和后缀"
                    '先获取主目录文件
                    For Each file In My.Computer.FileSystem.GetFiles(directory.FullName,
                                             FileIO.SearchOption.SearchTopLevelOnly,
                                                 filter)
                        files.Add(System.IO.Path.GetFileName(file))
                    Next

                    '再获取子目录文件
                    For Each subDirectory In directory.GetDirectories()
                        Try
                            For Each file In My.Computer.FileSystem.GetFiles(subDirectory.FullName, FileIO.SearchOption.SearchAllSubDirectories, filter)
                                files.Add(System.IO.Path.GetFileName(file))
                            Next
                        Catch ex As Exception
                            Continue For
                        End Try
                    Next
#End Region
                Else
#Region "获取文件全路径"
                    '先获取主目录文件
                    For Each file In My.Computer.FileSystem.GetFiles(directory.FullName,
                                             FileIO.SearchOption.SearchTopLevelOnly,
                                                 filter)
                        files.Add(System.IO.Path.GetFullPath(file))
                    Next

                    '再获取子目录文件
                    For Each subDirectory In directory.GetDirectories()
                        Try
                            For Each file In My.Computer.FileSystem.GetFiles(subDirectory.FullName, FileIO.SearchOption.SearchAllSubDirectories, filter)
                                files.Add(System.IO.Path.GetFullPath(file))
                            Next
                        Catch ex As Exception
                            Continue For
                        End Try
                    Next
#End Region
                End If
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try

            Return files
        End Function

        ''' <summary>
        ''' 获取某目录下的所有文件,可以避免权限拒绝的异常。默认只获取文件名（包含后缀）
        ''' 可以避免权限拒绝的异常
        ''' </summary>
        ''' <param name="filePath"></param>
        ''' <param name="filter">文件筛选器</param>
        ''' <returns></returns>
        Public Shared Function GetFiles(ByVal filePath As String,
                                         ByVal filter As String()) As List(Of String)
            Return GetFiles(filePath, filter, FileNameMode.FileName)
        End Function
    End Class

End Namespace
