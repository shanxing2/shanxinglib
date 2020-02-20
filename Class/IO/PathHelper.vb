Imports System.IO
Imports System.Windows.Forms
Imports ShanXingTech.Text2
Imports ShanXingTech.Win32API

Namespace ShanXingTech.IO2
    Public Class PathHelper

#Region "函数区"
        ''' <summary>
        ''' 打开文件所在的文件夹，并且定位到文件(选中文件)
        ''' </summary>
        ''' <param name="fileName"></param>
        Public Shared Sub LocateFile(ByVal fileName As String)
            Process.Start("Explorer.exe", "/select," & fileName)
        End Sub

        ''' <summary>
        ''' 获取选定的文件名
        ''' </summary>
        ''' <param name="filter">文件筛选器</param>
        ''' <param name="dialogTitle"></param>
        ''' <returns></returns>
        Public Shared Function GetFileName(ByVal filter As String, ByVal dialogTitle As String) As (FileName As String, Success As Boolean)

            Dim filename = String.Empty
            Dim success As Boolean
            Dim funcRst = (filename, success)

            Try
                Using openFileDlg As New OpenFileDialog()
                    ' 如果有上次选择的路径这初始化路径设置为上次的
                    If Not Conf.Instance.LastBrowseFolder.IsNullOrEmpty AndAlso
                        System.IO.Directory.Exists(Conf.Instance.LastBrowseFolder) Then
                        openFileDlg.InitialDirectory = Conf.Instance.LastBrowseFolder
                    Else
                        openFileDlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    End If

                    openFileDlg.Title = dialogTitle
                    openFileDlg.Filter = filter
                    openFileDlg.FileName = String.Empty

                    If openFileDlg.ShowDialog() = DialogResult.OK Then
                        filename = openFileDlg.FileName

                        If filename.Length = 0 Then Return funcRst

                        ' 记住这次选择的路径 以备下次使用
                        Conf.Instance.LastBrowseFolder = Path.GetDirectoryName(filename)
                        Conf.Instance.Save()

                        funcRst = (filename, True)
                    End If
                End Using
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try

            Return funcRst
        End Function

        ''' <summary>
        ''' 获取选定的文件名
        ''' </summary>
        ''' <param name="filter"></param>
        ''' <param name="dialogTitle"></param>
        ''' <returns></returns>
        Public Shared Function GetFileName(ByVal filter As FileFilter, ByVal dialogTitle As String) As (FileName As String, Success As Boolean)
            Dim getRst = GetFileName(EnsureHandledFileFilter(filter), dialogTitle)

            Return getRst
        End Function

        ''' <summary>
        ''' 获取要选择的Excel文件名
        ''' 成功返回文件名，否则返回空
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function GetExcelFileName() As String

            Dim funcRst As String = String.Empty
            GetExcelFileName(Nothing, funcRst)

            Return funcRst
        End Function

        ''' <summary>
        ''' 获取要选择的Excel文件名
        ''' 成功返回true,取消返回false
        ''' </summary>
        ''' <param name="textbox">显示选中的文件短路径</param>
        ''' <param name="filename">返回选中的文件完整路径，取消返回空</param>
        ''' <returns></returns>
        Public Shared Function GetExcelFileName(ByVal textbox As TextBox， ByRef filename As String) As Boolean
            Dim getRst = GetFileName(FileFilter.EXCEL, "选择文件")
            Dim funcRst = getRst.Success
            filename = getRst.FileName

            ' 转换成短路径
            If textbox IsNot Nothing Then
                textbox.Text = GetShortPath(filename)
            End If

            Return funcRst
        End Function

        ''' <summary>
        ''' 获取要选择的Excel文件名
        ''' </summary>
        ''' <param name="textbox">显示选中的文件短路径；可为Nothing，不接收返回信息</param>
        ''' <returns>成功 Success=true,取消 Success=false;FileName返回选中的文件完整路径，取消返回空</returns>
        Public Shared Function GetExcelFileName(ByVal textbox As TextBox) As (FileName As String, Success As Boolean)
            Dim filename As String = Nothing
            Dim Success = GetExcelFileName(textbox, filename)

            Return (filename, Success)
        End Function

        ''' <summary>
        ''' 获取要选择的文件名
        ''' </summary>
        ''' <param name="sourceFileName">选中的文件名</param>
        ''' <param name="dlgTitle"></param>
        ''' <param name="dlgFilter"></param>
        ''' <returns>转换成功的话函数返回选中文件短路径，否则返回选中文件全路径</returns>
        Public Shared Function GetShortFileName(ByRef sourceFileName As String, ByVal dlgTitle As String, ByVal dlgFilter As String) As String
            Dim funcRst As String = String.Empty
            sourceFileName = String.Empty

            Try
                Dim openFileDlg As New OpenFileDialog
                ' 如果有上次选择的路径这初始化路径设置为上次的
                If Not Conf.Instance.LastBrowseFolder.IsNullOrEmpty Then
                    openFileDlg.InitialDirectory = Conf.Instance.LastBrowseFolder
                Else
                    openFileDlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                End If

                openFileDlg.Title = dlgTitle
                openFileDlg.Filter = dlgFilter
                openFileDlg.FileName = String.Empty
                If openFileDlg.ShowDialog = DialogResult.OK Then
                    sourceFileName = openFileDlg.FileName
                    funcRst = sourceFileName

                    If funcRst.Length = 0 Then Return funcRst

                    ' 记住这次选择的路径 以备下次使用
                    Conf.Instance.LastBrowseFolder = Path.GetDirectoryName(funcRst)
                    Conf.Instance.Save()

                    ' 转换成短路径
                    funcRst = GetShortPath(funcRst)
                End If
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try

            Return funcRst
        End Function

        ''' <summary>
        ''' 输入全路径（文件夹路径或者文件路径）
        ''' 成功返回短路径，失败返回原路径
        ''' </summary>
        ''' <param name="fullPath"></param>
        ''' <returns></returns>
        Public Shared Function GetShortPath(ByVal fullPath As String) As String
            Dim funcRst As String = fullPath

            ' 显示短路径
            If fullPath.Length > 0 Then
                ' sb的初始capacity必须2倍的字符串长度
                Dim sb As Text.StringBuilder = StringBuilderCache.Acquire(fullPath.Length * 2)
                Dim shortPathLen As Integer = UnsafeNativeMethods.GetShortPathName(fullPath, sb, sb.Capacity)

                If shortPathLen > 0 Then
                    sb.Remove(shortPathLen, sb.Length - shortPathLen)
                    funcRst = StringBuilderCache.GetStringAndReleaseBuilder(sb)
                End If
            End If

            Return funcRst
        End Function

        ''' <summary>
        ''' 获取用户选择的文件夹路径
        ''' </summary>
        ''' <param name="textBox">要输出的textbox控件名称，输出路径</param>
        ''' <returns>成功返回true，其余返回false</returns>
        Private Shared Function GetSelectPath(textBox As TextBox， ByVal dlgTips As String， ByVal showNewFolderButton As Boolean) As (FolderName As String, Success As Boolean)
            Dim funcRst As Boolean
            Dim tempFolder As String

            Try
                Using folderBrowserDlg As New FolderBrowserDialog
                    ' 如果有上次选择的路径这初始化路径设置为上次的
                    '否则设置为 C:\ProgramData\ShanXingTech\具体程序集名称\具体程序集版本
                    If Not Conf.Instance.LastBrowseFolder.IsNullOrEmpty Then
                        ' 默认选中上次选择的文件夹
                        folderBrowserDlg.SelectedPath = Conf.Instance.LastBrowseFolder
                    Else
                        folderBrowserDlg.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    End If

                    folderBrowserDlg.Description = dlgTips
                    ' 显示 新建文件夹 按钮
                    folderBrowserDlg.ShowNewFolderButton = showNewFolderButton

                    ' 保存文件夹 路径
                    If folderBrowserDlg.ShowDialog() = DialogResult.OK Then
                        tempFolder = folderBrowserDlg.SelectedPath
                        If Not tempFolder.EndsWith("\") Then tempFolder += "\"

                        ' 创建一个新的文件夹
                        System.IO.Directory.CreateDirectory(tempFolder)

                        If tempFolder.Length = 0 Then Return (tempFolder, False)

                        ' 记住这次选择的路径 以备下次使用
                        Conf.Instance.LastBrowseFolder = Path.GetDirectoryName(tempFolder)
                        Conf.Instance.Save()

                        textBox.Text = tempFolder
                        funcRst = True
                    End If
                End Using
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try

            Return (tempFolder, funcRst)
        End Function

        ''' <summary>
        ''' 获取文件要存放的文件夹路径（由用户选择）
        ''' </summary>
        ''' <param name="textBox">要输出的textbox控件名称，输出路径</param>
        ''' <returns>成功返回true，其余返回false</returns>
        Public Shared Function SetFileSavePath(textBox As TextBox， ByVal dlgTips As String) As (FolderName As String, Success As Boolean)
            Return GetSelectPath(textBox, dlgTips, True)
        End Function

        ''' <summary>
        ''' 获取文件要存放的文件夹路径（由用户选择）
        ''' </summary>
        ''' <param name="textBox">要输出的textbox控件名称，输出短路径</param>
        ''' <returns>成功返回true，其余返回false</returns>
        Public Shared Function SetFileSaveShortPath(textBox As TextBox， ByVal dlgTips As String) As (FolderName As String, Success As Boolean)
            Dim setRst = SetFileSavePath(textBox, dlgTips)
            ' 如果获取成功，那就返回段路径
            If setRst.Success Then
                textBox.Text = GetShortPath(setRst.FolderName)
            End If

            Return setRst
        End Function

        ''' <summary>
        ''' 获取文件要存放的文件夹路径（由用户选择）
        ''' </summary>
        ''' <param name="textBox">要输出的textbox控件名称，输出短路径</param>
        ''' <param name="folder">返回选择的文件夹路径</param>
        ''' <param name="dlgTips">需要显示的提示</param>
        ''' <returns>成功返回true，其余返回false</returns>
        Public Shared Function SetFileSaveShortPath(textBox As TextBox， ByRef folder As String， ByVal dlgTips As String) As Boolean
            Dim setRst = SetFileSavePath(textBox, dlgTips)
            ' 如果获取成功，那就返回段路径
            If setRst.Success Then
                textBox.Text = GetShortPath(setRst.FolderName)
            End If
            folder = setRst.FolderName

            Return setRst.Success
        End Function

        ''' <summary>
        ''' 获取文件存储的文件夹短路径（由用户选择）
        ''' </summary>
        ''' <param name="textBox">要输出的textbox控件名称，输出短路径</param>
        ''' <param name="dlgTips">需要显示的提示</param>
        ''' <returns>成功返回true，其余返回false</returns>
        Public Shared Function GetFileStoreShortPath(textBox As TextBox， ByVal dlgTips As String) As (FolderName As String, Success As Boolean)
            Dim setRst = GetFileStorePath(textBox, dlgTips)
            ' 如果获取成功，那就返回段路径
            If setRst.Success Then
                textBox.Text = GetShortPath(setRst.FolderName)
            End If

            Return setRst
        End Function

        ''' <summary>
        ''' 获取文件存储的文件夹路径（由用户选择）
        ''' </summary>
        ''' <param name="textBox">要输出的textbox控件名称，输出路径</param>
        ''' <param name="dlgTips">需要显示的提示</param>
        ''' <returns>成功返回true，其余返回false</returns>
        Public Shared Function GetFileStorePath(textBox As TextBox， ByVal dlgTips As String) As (FolderName As String, Success As Boolean)
            Return GetSelectPath(textBox, dlgTips, False)
        End Function


        ''' <summary>
        ''' 获取文件存储的文件夹路径（由用户选择）
        ''' </summary>
        ''' <param name="textBox">要输出的textbox控件名称，输出路径</param>
        ''' <param name="folder">返回选择的文件夹路径</param>
        ''' <returns>成功返回true，其余返回false</returns>
        Public Shared Function GetFileStorePath(textBox As TextBox， ByRef folder As String， ByVal dlgTips As String) As Boolean
            Dim setRst = GetFileStorePath(textBox, dlgTips)
            folder = setRst.FolderName

            Return setRst.Success
        End Function

        ''' <summary>
        ''' 获取要新建的文件名
        ''' </summary>
        ''' <param name="filter">筛选器</param>
        ''' <returns>成功返回true，其余返回false</returns>
        Public Shared Function SetSaveFileName(ByVal filter As String) As (FileName As String, Success As Boolean)
            Dim funcRst = SetSaveFileName(filter, "另存为")

            Return funcRst
        End Function

        ''' <summary>
        ''' 获取要新建的文件名
        ''' </summary>
        ''' <param name="filter">筛选器</param>
        ''' <returns>成功返回true，其余返回false</returns>
        Public Shared Function SetSaveFileName(ByVal filter As FileFilter) As (FileName As String, Success As Boolean)
            Return SetSaveFileName(EnsureHandledFileFilter(filter), "另存为")
        End Function

        ''' <summary>
        ''' 获取要新建的文件名
        ''' </summary>
        ''' <param name="filter">筛选器</param>
        ''' <returns>成功返回true，其余返回false</returns>
        Public Shared Function SetSaveFileName(ByVal filter As FileFilter, ByVal dialogTitle As String) As (FileName As String, Success As Boolean)
            Return SetSaveFileName(EnsureHandledFileFilter(filter), dialogTitle)
        End Function

        Private Shared Function EnsureHandledFileFilter(ByVal filter As FileFilter) As String
            Dim filterString = filter.GetDescriptions
            If filterString.IndexOf(", ") > -1 Then
                filterString = filterString.Replace(", ", "|")
            End If

            Return filterString
        End Function


        ''' <summary>
        ''' 获取要新建的文件名
        ''' </summary>
        ''' <param name="filter">筛选器</param>
        ''' <returns>成功返回true，其余返回false</returns>
        Public Shared Function SetSaveFileName(ByVal filter As String, ByVal dialogTitle As String) As (FileName As String, Success As Boolean)
            Dim funcRst As Boolean
            Dim fileName = String.Empty

            Try
                Using saveFileDlg As New SaveFileDialog
                    ' 如果有上次选择的路径这初始化路径设置为上次的
                    If Not Conf.Instance.LastBrowseFolder.IsNullOrEmpty AndAlso System.IO.Directory.Exists(Conf.Instance.LastBrowseFolder) Then
                        saveFileDlg.InitialDirectory = Conf.Instance.LastBrowseFolder
                    Else
                        saveFileDlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    End If

                    saveFileDlg.Title = dialogTitle
                    saveFileDlg.Filter = filter

                    saveFileDlg.FileName = String.Empty

                    If saveFileDlg.ShowDialog() = DialogResult.OK Then
                        fileName = saveFileDlg.FileName

                        ' 记住这次选择的路径 以备下次使用
                        Conf.Instance.LastBrowseFolder = Path.GetDirectoryName(fileName)
                        Conf.Instance.Save()

                        funcRst = True
                    End If
                End Using
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try

            Return (fileName, funcRst)
        End Function

        ''' <summary>
        ''' 判断参数 <paramref name="fileName"/> 是否以 ""c开头
        ''' </summary>
        ''' <param name="fileName"></param>
        ''' <returns></returns>
        Private Shared Function CopyStartWithEmptyChar(ByVal fileName As String) As Boolean
            Return "‪"c = fileName.Chars(0)
        End Function

        ''' <summary>
        ''' 确保参数 <paramref name="fileName"/> 不以 ""c开头
        ''' </summary>
        ''' <param name="fileName"></param>
        Public Shared Sub EnsureNoStartWithEmptyChar(ByVal fileName As String)
            If CopyStartWithEmptyChar(fileName) Then
                Throw New ArgumentException($"参数{NameOf(fileName)}是以一个'""""'字符开头的字符串，该字符不能用作文件路径，此错误可能由于用户直接从‘文件属性——安全——对象属性’中复制引起。")
            End If
        End Sub
#End Region
    End Class
End Namespace

