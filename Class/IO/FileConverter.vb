Imports System.IO
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports ShanXingTech.Text2

Namespace ShanXingTech.IO2
    Public Class FileConverter


        ''' <summary>
        ''' 方法之一txt导入到Excel
        ''' 另外还可以txt导入mdb,然后再直接查询到Excel
        ''' sheetName 工作表名称,如 Sheet1
        ''' txtPath 完整的路径名,如 "C:\test.txt"
        ''' excelFileName 文件名不需要带后缀 程序会自动检查,然后直接用适用于当前Excel版本的后缀,如"C:\test.txt"
        ''' 注意！文件名不能包含特殊字符
        ''' </summary>
        ''' <param name="params"></param>
        ''' <returns>成功 Success 返回True,ExcelFileName 返回文件名；失败 Success 返回False,ExcelFileName 返回错误信息</returns>
        Public Shared Function Txt2Excel(ByVal params As ExportExcelFileInfo) As (ExcelFileName As String, Success As Boolean)
            Dim success As Boolean

            ' 获取参数
            Dim sheetName As String
            Dim txtPath As String
            Dim excelFileName As String

            Try
                ' 获取参数
                sheetName = params.SheetName
                txtPath = params.TxtCachePath
                excelFileName = params.ExcelFileName

                ' 文件名不能包含特殊字符,否则不能导出
                ' 特殊字符转换为 "asc值" 的形式，值 为相应的特殊字符asc值
                Dim directoryName = Path.GetDirectoryName(excelFileName)
                directoryName = If(directoryName.EndsWith("\"c), directoryName, directoryName & "\")
                Dim tempExcelFileName = excelFileName.Replace(directoryName, "")
                For Each invalidFileNameChar In Path.GetInvalidFileNameChars
                    tempExcelFileName = tempExcelFileName.Replace(invalidFileNameChar.ToString, Asc(invalidFileNameChar).ToString.Insert(0, "ASC"))
                Next

                excelFileName = directoryName & tempExcelFileName

                PathHelper.EnsureNoStartWithEmptyChar(excelFileName)

                If System.IO.File.Exists(txtPath) = False Then
                    Debug.Print(Logger.MakeDebugString("数据源不存在！"))
                    Return ("数据源不存在！", False)
                End If
            Catch ex As IndexOutOfRangeException
                Logger.Debug(ex)
            Catch ex As ArgumentException
                Logger.Debug(ex)
            End Try

            Dim xlBook As Object = Nothing
            Dim xlSheet As Object = Nothing

            ' 后期绑定Excel对象 不需要知道系统安装的是哪个版本的Excel
            ' 不需要引用Excel
            If Not ExcelEngine.Excel.Exists Then
                Return ("本机未安装WPS或者Excel，导出失败！", False)
            End If

            xlBook = ExcelEngine.Excel.Instance.Workbooks.Add
            xlSheet = xlBook.Worksheets(sheetName)
            ' xlApp.Visible = True
            Try
                ' 先检查本地是否有文件
                ' 有则直接删除
                If System.IO.File.Exists(excelFileName & ".xlsx") Then
                    System.IO.File.Delete(excelFileName & ".xlsx")
                ElseIf System.IO.File.Exists(excelFileName & ".xls") Then
                    System.IO.File.Delete(excelFileName & ".xls")
                End If


                ' txt导入到Excel
                ' With xlSheet.QueryTables.Add(Connection:="TEXT;" & TxtPath, Destination:=xlSheet.Range("$A$" & DateRows))
                With xlSheet.QueryTables.Add(Connection:="TEXT;" & txtPath, Destination:=xlSheet.Range("$A$1"))
                    .Name = System.IO.Path.GetFileNameWithoutExtension(txtPath)
                    .FieldNames = True
                    .RowNumbers = False
                    .FillAdjacentFormulas = False
                    .PreserveFormatting = True
                    .RefreshOnFileOpen = False
                    .RefreshStyle = 1   ' xlInsertDeleteCells
                    .SavePassword = False
                    .SaveData = True
                    .AdjustColumnWidth = True
                    .RefreshPeriod = 0
                    .TextFilePromptOnRefresh = False
                    ' 返回或设置正向查询表中导入的文本文件的编码格式
                    ' .TextFilePlatform = 936     'GB2312/ANSI
                    ' .TextFilePlatform = 950     'BIG5
                    .TextFilePlatform = params.TextFilePlatform     'UTF-8 Unicode编码
                    '.TextFilePlatform = GetCodePage(txtPath)
                    .TextFileStartRow = 1
                    .TextFileParseType = 1  ' xlDelimited
                    .TextFileTextQualifier = 1  ' xlTextQualifierDoubleQuote
                    'True if consecutive delimiters are treated as a single delimiter when you import a text file into a query table. The default value is False. 
                    ' 将连续分隔符看作是一个分隔符
                    .TextFileConsecutiveDelimiter = True
                    ' Tab当作分隔符
                    .TextFileTabDelimiter = False
                    ' 分号当作分隔符
                    .TextFileSemicolonDelimiter = False
                    ' 空格 当作分隔符
                    .TextFileSpaceDelimiter = False
                    ' 逗号 当作分隔符
                    .TextFileCommaDelimiter = False
                    ' 自定义分隔符为 “@” 自定义分隔符 不能跟系统已经定义好的一样
                    .TextFileOtherDelimiter = If(params.TextFileOtherDelimiter, "")
                    ' 设置导出之后每列数据的格式（数组长度可以多于数据列数,2为常规，系统自动根据前八行数据设置）
                    'xlGeneralFormat     常规
                    'xlTextFormat    文本
                    'xlSkipColumn    跳过列
                    'xlDMYFormat     “日-月-年”日期格式
                    'xlDYMFormat     “日-年-月”日期格式
                    'xlEMDFormat     EMD 日期
                    'xlMDYFormat     “月-日-年”日期格式
                    'xlMYDFormat     “月-年-日”日期格式
                    'xlYDMFormat     “年-日-月”日期格式
                    'xlYMDFormat     “年-月-日”日期格式
                    ' 默认常量是xlGeneral
                    .TextFileColumnDataTypes = params.TextFileColumnsDataType.ToArray
                    '.TextFileColumnsDataType = SetColumnDataTypes(txtPath)

                    .TextFileTrailingMinusNumbers = True
                    .Refresh(BackgroundQuery:=False)
                End With
                '51 = xlOpenXMLWorkbook (without macro's in 2007-2013, xlsx)
                '52 = xlOpenXMLWorkbookMacroEnabled (with or without macro's in 2007-2013, xlsm)
                '50 = xlExcel12 (EXCEL Binary Workbook in 2007-2013 with or without macro's, xlsb)
                '56 = xlExcel8 (97-2003 format in EXCEL 2007-2013, xls)
                '8.0,              9.0,       10.0,          11.0,      12.0,   14.0,15.0分别对应于
                'EXCEL office97、office2000、officeXP(2002)、office2003、office2007、2010、2013

                If Not IO.File.Exists(excelFileName) Then
                    ' 保存为xls格式
                    ' xlBook.SaveAs ExcelFileName, 56      
                    ' 有些电脑会有“类 Workbook 的 SaveAs 方法无效”错误(wps和excel同时安装）
                    ' 去掉后面的FileFormat参数，然后根据当前wps或者excel版本号保存为相应的格式即可
                    ' 无论是excel或者wps，都是版本号为11以上的才默认保存为xlsx格式
                    If ExcelEngine.Excel.Version > 11.0# Then
                        excelFileName = excelFileName & ".xlsx"
                    ElseIf ExcelEngine.Excel.Version <= 11.0# Then
                        excelFileName = excelFileName & ".xls"
                    End If

                    If System.IO.File.Exists(excelFileName) Then
                        ' 如果本地已经存在xlsFileName用SaveAs保存时会有提示
                        ' 屏蔽提示
                        ExcelEngine.Excel.Instance.DisplayAlerts = False
                        Call xlBook.SaveAs(excelFileName)
                        xlBook.Close(True)
                    Else
                        Call xlBook.SaveAs(excelFileName)
                        xlBook.Close(True)
                    End If
                Else
                    xlBook.save
                End If


                ' 导出成功,返回true
                success = True
                Debug.Print(Logger.MakeDebugString("保存完成"))
                ' 发生数据源不存在、EXCEL 找不到文本文件来可刷新该外部数据区域  等错误
                ' Catch ex As Runtime.InteropServices.COMException
            Catch ex As Exception
                Logger.Debug(ex)

                excelFileName = ex.Message
            Finally
                If xlBook IsNot Nothing Then
                    ' 关闭提示是否保存提示框（不提示是否保存,直接保存文件）
                    'xlBook.Saved = True
                    ' 不显示 提示信息 包括保存更改
                    'xlApp.DisplayAlerts = False        
                    ExcelEngine.Excel.Instance.Workbooks.Close
                    'OfficeEngine.Office.Instance.Quit

                    Runtime.InteropServices.Marshal.FinalReleaseComObject(xlSheet)
                    Runtime.InteropServices.Marshal.FinalReleaseComObject(xlBook)
                    xlSheet = Nothing
                    xlBook = Nothing

                    GC.Collect()
                    ' 挂起所有线程 确保在垃圾回收完成之前，其他线程不会调用null对象
                    GC.WaitForPendingFinalizers()
                End If
            End Try

            Return (excelFileName, success)
        End Function

        ''' <summary>
        ''' 从 <paramref name="value"/> 缓存导出到Excel,需要传入Excel文件保存路径
        ''' </summary>
        ''' <param name="value"></param>
        ''' <param name="columnNames">列头数组，以<paramref name="columnDelimiter"/>作为分隔符分割;如果传入空数组或者Nothing，则不导出列头</param>
        ''' <param name="columnsDataType">格式值列表，列表项数与列数对应,传入空列表或者Nothing将自动转换</param>
        ''' <param name="columnDelimiter">列分隔符,注：连续的分隔符会自动合并</param>
        ''' <param name="excelFileFullPath">将存储Excel文件的绝对路径；如果带有后缀名，函数会自动去掉传入的后缀名，然后根据系统环境使用合适的后缀名</param>
        ''' <returns>Success指示是否导出成功，ExcelFileName为Excel文件的保存路径</returns>
        Public Shared Function ExportToExcel(ByRef value As String, ByRef columnNames As String(), ByRef columnsDataType As List(Of ExcelColumnDataType), ByRef columnDelimiter As String, ByRef excelFileFullPath As String) As (ExcelFileName As String, Success As Boolean)
            Dim funcRst As Boolean

            Try
                Dim excelFilePath = Path.GetDirectoryName(excelFileFullPath)

                If Not IO.Directory.Exists(excelFilePath) Then
                    Directory.Create(excelFilePath)
                End If

                Dim cacheTxtFilePath = $"{excelFilePath}\ShanXingTechCache{Date.Now.ToTimeStampString(TimePrecision.Millisecond)}.txt"

                ' 如果存在缓存文件 则删除
                If IO.File.Exists(cacheTxtFilePath) Then
                    IO.File.Delete(cacheTxtFilePath)
                End If

                ' 把列头信息存到本地缓存 以备导出
                ' 如果没有传入列头信息，那就不需要导出列头
                If columnNames?.Length > 0 AndAlso columnDelimiter IsNot Nothing Then
                    Dim tempColumnNames = String.Join(columnDelimiter, columnNames)
                    Writer.WriteText(cacheTxtFilePath, tempColumnNames)
                End If

                ' 把没有数据的列替换成空格，这样导出就不会乱序
                Dim doubleDelimiter = columnDelimiter & columnDelimiter
                While value.Contains(doubleDelimiter & doubleDelimiter)
                    value = value.Replace(doubleDelimiter & doubleDelimiter, doubleDelimiter & " " & doubleDelimiter)
                End While

                ' 把具体信息存到本地缓存 以备导出
                Writer.WriteText(cacheTxtFilePath, value)

                Dim convertRst = ExportToExcel(cacheTxtFilePath, columnsDataType, columnDelimiter, excelFileFullPath)

                ' 如果导出失败 会返回false
                funcRst = convertRst.Success

                ' 把返回的文件名或者错误信息保存起来
                excelFileFullPath = convertRst.ExcelFileName
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try

            Return (excelFileFullPath, funcRst)
        End Function

        ''' <summary>
        ''' 从缓存<paramref name="cacheTxtFilePath"/>导出到Excel,需要传入Excel文件保存路径
        ''' </summary>
        ''' <param name="columnsDataType">格式值列表，列表项数与列数对应,传入空列表或者Nothing将自动转换</param>
        ''' <param name="columnDelimiter">列分隔符,注：连续的分隔符会自动合并</param>
        ''' <param name="excelFileFullPath">将存储Excel文件的绝对路径；如果带有后缀名，函数会自动去掉传入的后缀名，然后根据系统环境使用合适的后缀名</param>
        ''' <returns>Success指示是否导出成功，ExcelFileName为Excel文件的保存路径</returns>
        Public Shared Function ExportToExcel(ByVal cacheTxtFilePath As String, ByRef columnsDataType As List(Of ExcelColumnDataType), ByRef columnDelimiter As String, ByRef excelFileFullPath As String) As (ExcelFileName As String, Success As Boolean)
            Dim funcRst As Boolean

            Try
                ' 去掉后缀，函数会自己根据系统安装的excel文档选择合适的后缀
                Dim fileExtention = ".xls"
                If excelFileFullPath.EndsWith(fileExtention, StringComparison.OrdinalIgnoreCase) Then
                    excelFileFullPath = excelFileFullPath.Remove(excelFileFullPath.Length - fileExtention.Length)
                End If
                fileExtention = ".xlsx"
                If excelFileFullPath.EndsWith(fileExtention, StringComparison.OrdinalIgnoreCase) Then
                    excelFileFullPath = excelFileFullPath.Remove(excelFileFullPath.Length - fileExtention.Length)
                End If

                ' 导出缓存到exce线程(文件名不能包含特殊字符)
                Dim excelFileInfo As New ExportExcelFileInfo With {
                    .SheetName = "Sheet1",
                    .TxtCachePath = cacheTxtFilePath,
                    .ExcelFileName = excelFileFullPath,
                    .TextFileOtherDelimiter = columnDelimiter,
                    .TextFilePlatform = CodePage.UTF8
                }

                ' 如果调用者没有传入列数据类型，那么就是用默认的列数据类型
                If columnsDataType IsNot Nothing Then
                    excelFileInfo.TextFileColumnsDataType = columnsDataType
                End If

                Dim convertOperate As (ExcelFileName As String, Success As Boolean) = Nothing
                Task.Run(
                Sub()
                    convertOperate = Txt2Excel(excelFileInfo)

                    ' 删除缓存文件
                    If System.IO.File.Exists(cacheTxtFilePath) Then
                        System.IO.File.Delete(cacheTxtFilePath)
                    End If
                End Sub).GetAwaiter.GetResult()

                ' 如果导出失败 会返回false
                funcRst = convertOperate.Success

                ' 把返回的文件名或者错误信息保存起来
                excelFileFullPath = convertOperate.ExcelFileName
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try

            Return (excelFileFullPath, funcRst)
        End Function

        ''' <summary>
        ''' 从StringBuilder缓存导出到Excel,需要传入Excel文件保存路径
        ''' </summary>
        ''' <param name="sb"></param>
        ''' <param name="columnNames">列头数组，以<paramref name="columnDelimiter"/>作为分隔符分割;如果传入空数组或者Nothing，则不导出列头</param>
        ''' <param name="columnsDataType">格式值列表，列表项数与列数对应,传入空列表或者Nothing将自动转换</param>
        ''' <param name="columnDelimiter">列分隔符,注：连续的分隔符会自动合并</param>
        ''' <param name="excelFileFullPath">将存储Excel文件的绝对路径；如果带有后缀名，函数会自动去掉传入的后缀名，然后根据系统环境使用合适的后缀名</param>
        ''' <returns>Success指示是否导出成功，ExcelFileName为Excel文件的保存路径</returns>
        Public Shared Function ExportToExcel(ByRef sb As StringBuilder, ByRef columnNames As String(), ByRef columnsDataType As List(Of ExcelColumnDataType), ByRef columnDelimiter As String, ByRef excelFileFullPath As String) As (ExcelFileName As String, Success As Boolean)
            Return ExportToExcel(sb.ToString, columnNames, columnsDataType, columnDelimiter, excelFileFullPath)
        End Function

        ''' <summary>
        ''' 从StringBuilder缓存导出到Excel，不需要传入Excel保存路径，会有弹框要求选择Excel路径
        ''' </summary>
        ''' <param name="sb"></param>
        ''' <param name="columnNames">列头数组，以<paramref name="columnDelimiter"/>作为分隔符分割;如果传入空数组或者Nothing，则不导出列头</param>
        ''' <param name="columnsDataType">格式值列表，列表项数与列数对应,传入空列表或者Nothing将自动转换</param>
        ''' <param name="columnDelimiter">列分隔符</param>
        ''' <returns>Success指示是否导出成功，Message为Excel文件的保存路径或者是错误提示,取消导出时Message为空，Success为False</returns>
        Public Shared Function ExportToExcel(ByVal sb As StringBuilder, ByVal columnNames As String(), ByVal columnsDataType As List(Of ExcelColumnDataType), ByVal columnDelimiter As String) As (Message As String, Success As Boolean)

            Dim excelFileFullPath As String = String.Empty
            Dim funcRst As Boolean

            Dim rst = PathHelper.SetSaveFileName(FileFilter.EXCEL)
            If rst.Success Then
                excelFileFullPath = rst.FileName
                Return ExportToExcel(sb, columnNames, columnsDataType, columnDelimiter, excelFileFullPath)
            End If

            Return (excelFileFullPath, funcRst)
        End Function

        ''' <summary>
        ''' 从StringBuilder缓存导出到Excel，不需要传入Excel保存路径，会有弹框要求选择Excel路径
        ''' </summary>
        ''' <param name="sb">存储单列数据的StringBuilder</param>
        ''' <returns>Success指示是否导出成功，Message为Excel文件的保存路径或者是错误提示,取消导出时Message为空，Success为False</returns>
        Public Shared Function ExportToExcel(ByVal sb As StringBuilder) As (Message As String, Success As Boolean)
            Return ExportToExcel(sb, Nothing, Nothing, Nothing)
        End Function

        ''' <summary>
        ''' <paramref name="dgv"/> 整表导出到Excel，不需要传入Excel保存路径，会有弹框要求选择Excel路径
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <param name="columnNames">列头数组，以<paramref name="columnDelimiter"/>作为分隔符分割;如果传入空数组或者Nothing，则不导出列头</param>
        ''' <param name="columnDelimiter">列分隔符</param>
        ''' <param name="columnsDataType">格式值列表，列表项数与列数对应,传入空列表或者Nothing将自动转换</param>
        ''' <returns>Success指示是否导出成功，Message为Excel文件的保存路径或者是错误提示,取消导出时Message为空，Success为False</returns>
        Public Shared Function ExportToExcel(ByVal dgv As DataGridView, ByVal columnNames As String(), ByVal columnDelimiter As String, ByVal columnsDataType As List(Of ExcelColumnDataType)) As (Message As String, Success As Boolean)

            Dim excelFileFullPath As String = String.Empty
            Dim funcRst As Boolean

            Dim rst = PathHelper.SetSaveFileName(FileFilter.EXCEL)
            If rst.Success Then
                excelFileFullPath = rst.FileName

                ' 从dgv 获取整表数据到缓存sb
                Dim sb = New StringBuilder(100 * dgv.RowCount)
                For Each row As DataGridViewRow In dgv.Rows
                    For colIndex As Integer = 0 To dgv.ColumnCount - 1
                        sb.Append(row.Cells.Item(colIndex).FormattedValue).Append("@@")
                    Next
                    ' 去除最后两个@@ 并且添加换行符
                    sb.Remove(sb.Length - 2, 2).AppendLine()
                Next
                ' 把没有数据的列替换成空格，这样导出就不会乱序
                sb.Replace("@@@@", "@@ @@")

                Return ExportToExcel(sb, columnNames, columnsDataType, columnDelimiter, excelFileFullPath)
            End If

            Return (excelFileFullPath, funcRst)
        End Function

        ''' <summary>
        ''' <paramref name="dgv"/> 整表导出到Excel，不需要传入Excel保存路径，会有弹框要求选择Excel路径
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <param name="columnsDataType">格式值列表，列表项数与列数对应,传入空列表或者Nothing将自动转换</param>
        ''' <returns>Success指示是否导出成功，Message为Excel文件的保存路径或者是错误提示,取消导出时Message为空，Success为False</returns>
        Public Shared Function ExportToExcel(ByVal dgv As DataGridView, ByVal columnsDataType As List(Of ExcelColumnDataType)) As (Message As String, Success As Boolean)

            Dim excelFileFullPath As String = String.Empty
            Dim funcRst As Boolean

            Dim rst = PathHelper.SetSaveFileName(FileFilter.EXCEL)
            If rst.Success Then
                excelFileFullPath = rst.FileName

                ' 从dgv 获取整表数据到缓存sb
                Dim sb = New StringBuilder(100 * dgv.RowCount)
                For Each row As DataGridViewRow In dgv.Rows
                    For colIndex As Integer = 0 To dgv.ColumnCount - 1
                        sb.Append(row.Cells.Item(colIndex).FormattedValue).Append("@@")
                    Next
                    ' 去除最后两个@@ 并且添加换行符
                    sb.Remove(sb.Length - 2, 2).AppendLine()
                Next
                ' 把没有数据的列替换成空格，这样导出就不会乱序
                sb.Replace("@@@@", "@@ @@")

                ' 获取dgv所有列的列头作为将要导出数据的列头
                Dim columnNames(dgv.ColumnCount - 1) As String
                For colIndex = 0 To dgv.ColumnCount - 1
                    columnNames(colIndex) = dgv.Columns(colIndex).Name
                Next

                Dim columnDelimiter = "@@"
                Return ExportToExcel(sb, columnNames, columnsDataType, columnDelimiter, excelFileFullPath)
            End If

            Return (excelFileFullPath, funcRst)
        End Function
        Public Shared Function CsvToExcel(ByVal csvFileFullPath As String, ByVal codePage As CodePage) As String
            Dim csvContext = IO2.Reader.ReadFile(csvFileFullPath, Encoding.GetEncoding(codePage))
            If csvContext.Length = 0 Then
                Return ""
            End If

            Dim cacheExcelPath = Application.StartupPath
            If "\"c <> cacheExcelPath.Chars(cacheExcelPath.Length - 1) Then
                cacheExcelPath += "\"
            End If
            Dim cacheTxtFilePath = $"{cacheExcelPath}ShanXingTechCache{Date.Now.ToTimeStampString(TimePrecision.Millisecond)}"

            Dim expRst = ExportToExcel(csvContext, Nothing, Nothing, ",", cacheTxtFilePath)
            Return expRst.ExcelFileName
        End Function

        ''' <summary>
        ''' 返回或设置正向查询表中导入的文本文件的编码格式
        ''' </summary>
        ''' <param name="fileName"></param>
        ''' <returns></returns>
        Private Function GetCodePage(filename As String) As Integer
            Dim codePage As Integer = System.Text.Encoding.Default.CodePage

            Using fs As FileStream = New System.IO.FileStream(filename, System.IO.FileMode.Open, System.IO.FileAccess.Read)
                Dim br As BinaryReader = New System.IO.BinaryReader(fs)
                Dim buffer As Byte() = br.ReadBytes(2)

                '文件前两字节指定文件编码格式
                If buffer(0) = &HEF AndAlso buffer(1) = &HBB Then
                    codePage = System.Text.Encoding.UTF8.CodePage
                ElseIf buffer(0) = &HFE AndAlso buffer(1) = &HFF Then
                    codePage = System.Text.Encoding.BigEndianUnicode.CodePage
                ElseIf buffer(0) = &HFF AndAlso buffer(1) = &HFE Then
                    codePage = System.Text.Encoding.Unicode.CodePage
                Else
                    codePage = System.Text.Encoding.Default.CodePage
                End If
            End Using
            Return codePage
        End Function

        ''' <summary>
        ''' 设置每列字段的数据类型
        ''' </summary>
        ''' <param name="fileName"></param>
        ''' <returns></returns>
        Private Function SetColumnDataTypes(fileName As String) As Integer()
            Dim LineStr As String = String.Empty
            ' 字段分隔符
            Dim delimiter As Char() = {"@"c}

            ' 获取数据源的第一行 从此行可以得到数据的字段列 个数
            Using sr As New StreamReader(fileName)
                LineStr = sr.ReadLine()
            End Using

            Dim columnArray = LineStr.Split(delimiter)
            Dim ColDataTypes(columnArray.Count) As Integer
            ' 填充数组，所有列设置为 文本类型
            For i = 0 To columnArray.Count - 1
                ColDataTypes(i) = ExcelColumnDataType.XlTextFormat
            Next

            Return ColDataTypes
        End Function

    End Class
End Namespace

