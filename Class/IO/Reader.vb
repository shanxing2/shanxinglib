Imports System.Data.Common
Imports System.Data.OleDb
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports ShanXingTech.Exception2

Namespace ShanXingTech.IO2
	Public NotInheritable Class Reader
#Region "常量区"
		''' <summary>
		''' 以默认列名Fn(n表示大于1的整数，F1表示默认第一列)查询Sheet1表格中全部数据
		''' </summary>
		Public Const DefaultExcelQuerySql = "select * from [Sheet1$]"
#End Region

#Region "字段区"

#End Region

#Region "函数区"
		''' <summary>
		''' 共享方式读取文件的全部内容
		''' </summary>
		''' <param name="path">文件路径</param>
		''' <param name="encoding">字符编码</param>
		''' <returns>返回读取到的字符串</returns>
		Public Shared Function ReadFile(ByVal path As String， ByVal encoding As Text.Encoding) As String
			Dim funcRst As String = String.Empty

			Try
				' 只读共享模式，可以解决 正由另一进程使用，因此该进程无法访问此文件 问题
				Dim fs As New FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
				' 声明数据流文件读方法  
				Using sr As New StreamReader(fs, encoding)
					funcRst = sr.ReadToEnd()
				End Using
			Catch ex As Exception
				Logger.WriteLine(ex)
			End Try

			Return funcRst
		End Function

		''' <summary>
		''' 获取Csv文件的所有行数据
		''' </summary>
		''' <param name="csvFilePath"></param>
		''' <param name="encoding"></param>
		''' <returns></returns>
		Public Shared Function GetAllCsvRows(ByVal csvFilePath As String， ByVal encoding As Text.Encoding) As String()
			Dim csvText = ReadFile(csvFilePath, encoding)
			If csvText.Length = 0 Then Return {}

			Dim csvRows = csvText.Split({vbCrLf, vbCr, vbLf, Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

			Return csvRows
		End Function

		''' <summary>
		''' 设置自动序号
		''' </summary>
		''' <param name="sourceDt"></param>
		''' <returns></returns>
		Private Shared Function SetAutoSerialNumber(ByVal sourceDt As DataTable， ByVal rowNumberColumnName As String) As DataTable
			Dim dataTable = New DataTable()
			' AutoIncrement  获取或设置一个值，该值指示对于添加到该表中的新行，列是否将列的值自动递增  
			Dim column As New DataColumn() With {
				.AutoIncrement = True,
				.ColumnName = rowNumberColumnName,
				.AutoIncrementSeed = 1,
				.AutoIncrementStep = 1
			}
			dataTable.Columns.Add(column)
			' Merge合并DataTable  
			' table.Merge(dataTable)
			dataTable.Merge(sourceDt)

			Return dataTable
		End Function

		''' <summary>
		''' 设置DataGridView的数据源
		''' </summary>
		''' <param name="dgv"></param>
		''' <param name="dt"></param>
		Private Shared Sub SetDataToDataGridView(ByVal dgv As DataGridView, ByVal dt As DataTable)
			If dgv.InvokeRequired Then
				dgv.Invoke(Sub() dgv.DataSource = dt)
			Else
				dgv.DataSource = dt
			End If
		End Sub

		''' <summary>
		''' 从excel表中查找符合条件的信息并且返回
		''' </summary>
		''' <param name="excelFullPath"></param>
		''' <param name="sql">SQL查询语句</param>
		''' <param name="hdr">第一行是否作为标题</param>
		''' <returns></returns>
		Public Shared Function GetOneDataFromExcel(ByVal excelFullPath As String, ByVal sql As String, ByVal hdr As HDRMode) As String
			Dim funcRst = String.Empty

			Try
				Using reader As OleDbDataReader = IO2.Reader.GetDataFromExcel(excelFullPath, sql, hdr)
					If reader.Read Then
						funcRst = If(DBNull.Value.Equals(reader(0)), String.Empty, CStr(reader(0)))
					End If
				End Using
			Catch ex As Exception
				Logger.WriteLine(ex, sql,,,)
			End Try

			Return funcRst
		End Function

		''' <summary>
		''' ADO方式读取Excel文件里的数据
		''' jet/ace 数据驱动引擎会根据前8行数据设置每列的数据类型
		''' </summary>
		''' <param name="excelFullPath">Excel文件的全路径</param>
		''' <param name="sql">SQL查询语句</param>
		''' <param name="table">存储缓存数据的DataTable</param>
		''' <param name="hdr">第一行是否作为标题</param>
		''' <param name="showRowNumber">显示行号</param>
		Public Shared Sub GetDataFromExcel(ByVal excelFullPath As String, ByVal sql As String, ByRef table As DataTable, ByVal hdr As HDRMode, ByVal showRowNumber As Boolean)
			Try
				Using csvReader = GetDataFromExcel(excelFullPath, sql, hdr)
					' 把读取到的数据按照传入的参数决定是否要加行号列，然后再保存到缓存表
					table = DataReader(csvReader)
					If showRowNumber Then
						table = SetAutoSerialNumber(table, "序号")
					End If
				End Using
			Catch ex As Exception
				Logger.WriteLine(ex, sql,,,)
			End Try
		End Sub

		''' <summary>
		''' ADO方式读取Excel文件里的数据
		''' jet/ace 数据驱动引擎会根据前8行数据设置每列的数据类型
		''' </summary>
		''' <param name="excelFullPath">Excel文件的全路径</param>
		''' <param name="table">存储缓存数据的DataTable</param>
		''' <param name="hdr">第一行是否作为标题</param>
		''' <param name="showRowNumber">显示行号</param>
		Public Shared Sub GetDataFromExcel(ByVal excelFullPath As String, ByRef table As DataTable, ByVal hdr As HDRMode, ByVal showRowNumber As Boolean)
			GetDataFromExcel(excelFullPath, DefaultExcelQuerySql, table, hdr， showRowNumber)
		End Sub

		''' <summary>
		''' ADO方式读取Excel文件里的数据
		''' jet/ace 数据驱动引擎会根据前8行数据设置每列的数据类型
		''' </summary>
		''' <param name="excelFullPath">Excel文件的全路径</param>
		''' <param name="sql">SQL查询语句</param>
		''' <param name="dgv">读取出来的数据要绑定的<paramref name="dgv"/></param>
		''' <param name="hdr">第一行是否作为标题</param>
		''' <param name="showRowNumber">显示行号</param>
		Public Shared Sub GetDataFromExcel(ByVal excelFullPath As String, ByVal sql As String, ByRef dgv As DataGridView, ByVal hdr As HDRMode, ByVal showRowNumber As Boolean)
			Dim dataTable As DataTable = Nothing
			GetDataFromExcel(excelFullPath, sql, dataTable, hdr, showRowNumber)
			SetDataToDataGridView(dgv, dataTable)
		End Sub

		''' <summary>
		''' ADO方式读取Excel文件里的数据
		''' jet/ace 数据驱动引擎会根据前8行数据设置每列的数据类型
		''' </summary>
		''' <param name="excelFullPath">Excel文件的全路径</param>
		''' <param name="dgv">读取出来的数据要绑定的<paramref name="dgv"/></param>
		''' <param name="showRowNumber">显示行号</param>
		Public Shared Sub GetDataFromExcel(ByVal excelFullPath As String, ByRef dgv As DataGridView, ByVal showRowNumber As Boolean)
			GetDataFromExcel(excelFullPath, DefaultExcelQuerySql, dgv, HDRMode.Yes, showRowNumber)
		End Sub

		''' <summary>
		''' ADO方式读取Excel文件里的数据
		''' jet/ace 数据驱动引擎会根据前8行数据设置每列的数据类型
		''' </summary>
		''' <param name="excelFullPath">Excel文件的全路径</param>
		''' <param name="dgv">读取出来的数据要绑定的<paramref name="dgv"/></param>
		''' <param name="hdr">第一行是否作为标题</param>
		''' <param name="showRowNumber">显示行号</param>
		Public Shared Sub GetDataFromExcel(ByVal excelFullPath As String, ByRef dgv As DataGridView, ByVal hdr As HDRMode, ByVal showRowNumber As Boolean)
			GetDataFromExcel(excelFullPath, DefaultExcelQuerySql, dgv, hdr， showRowNumber)
		End Sub

		''' <summary>
		''' ADO方式读取Excel文件里的数据
		''' jet/ace 数据驱动引擎会根据前8行数据设置每列的数据类型
		''' </summary>
		''' <param name="excelFullPath">Excel文件的全路径</param>
		''' <param name="sql">SQL查询语句</param>
		''' <param name="dgv">读取出来的数据要绑定的<paramref name="dgv"/></param>
		''' <param name="showRowNumber">显示行号</param>
		Public Shared Sub GetDataFromExcel(ByVal excelFullPath As String, ByVal sql As String, ByRef dgv As DataGridView, ByVal showRowNumber As Boolean)
			GetDataFromExcel(excelFullPath, sql, dgv, HDRMode.Yes, showRowNumber)
		End Sub

        ''' <summary>
        ''' ADO方式读取Excel文件里的数据
        ''' jet/ace 数据驱动引擎会根据前8行数据设置每列的数据类型
        ''' 建议使用using语句调用,以确保使用完之后关闭reader,并且要在用完Reader之后 End Using 语句之前调用 Reader的CloseDbConnection扩展方法,确保相应的数据库连接在关闭
        ''' 注：如果出现 “未在本地计算机上注册“Microsoft.Jet.OLEDB.4.0”提供程序” 类似错误
        ''' 在解决方案上右键-属性-选择“编译”选项卡--点击 “高级编译选项”-弹出 高级编译设置窗口   ------在最下面的 “选择目标 CPU”---选择X86就可以 了
        ''' 另外，如果机器上装的是WPS,而且确实没有这个驱动的话，改成X86也依然有这个提示，这个时候就需要安装这个驱动了
        ''' 下载地址 https://www.microsoft.com/zh-CN/download/details.aspx?id=13255
        ''' 选择相应的驱动器
        ''' 最新测试，不管是xlsx还是xls，Microsoft.ACE.OLEDB.12.0驱动都可以读取，
        ''' 不过有个问题，装有office或者wps的电脑，不一定会有Microsoft.ACE.OLEDB.12.0这个驱动，但是一定会有
        ''' Microsoft.Jet.OLEDB.4.0 驱动，所以，需要根据实际情况，决定用哪个驱动
        ''' 如果office是32位那就装 AccessDatabaseEngine.exe 驱动
        ''' 如果是64位，那就装 AccessDatabaseEngine_X64.exe 驱动
        ''' 项目属性——编译——选择AnyCpu，勾选首选32位
        ''' </summary>
        ''' <param name="excelFullPath">Excel文件的全路径</param>
        ''' <param name="sql">SQL查询语句</param>
        ''' <param name="hdr">第一行是否作为标题</param>
        ''' <returns>返回所有数据的Reader，可以用索引或者列名读取</returns>
        Public Shared Function GetDataFromExcel(ByVal excelFullPath As String, ByVal sql As String, ByVal hdr As HDRMode) As OleDbDataReader
            PathHelper.EnsureNoStartWithEmptyChar(excelFullPath)

            Dim reader As OleDbDataReader

#Region "ADO.NET方式读取Excel数据"
            Try
                ' 20180426
                ' HDR=Yes 第一行作为标题
                Dim oleCn As New OleDb.OleDbConnection With {
                    .ConnectionString = If(excelFullPath.EndsWith(".xlsx"),
                    $"Provider={ExcelEngine.Excel.Provider};Data Source={excelFullPath};Extended Properties='{ExcelEngine.Excel.ISAM};HDR={hdr.ToString()};IMEX=2'",
                    $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={excelFullPath};Extended Properties='EXCEL 8.0;HDR={hdr.ToString()};IMEX=2'")
                }

                If oleCn.State = ConnectionState.Closed Then
                    oleCn.Open()
                End If

                Using sqlCmd = New OleDbCommand() With {
                    .CommandText = sql,
                    .Connection = oleCn
                }
                    reader = sqlCmd.ExecuteReader(CommandBehavior.CloseConnection)
                End Using
#End Region
            Catch ex As InvalidOperationException
                Logger.WriteLine(ex)
                If ex.Message.IndexOf(".OLEDB.", StringComparison.OrdinalIgnoreCase) > -1 Then
                    Throw New EngineNotFoundExcption(ex.Message, ex)
                Else
                    Throw
                End If
            Catch ex As OleDb.OleDbException
                Throw
            Catch ex As Exception
                '微软已知bug,遇到可跳过
                'If Err.Number = -2147467259 And Err.Description = "系统不支持所选择的排序。" Then
                '    Application.DoEvents()
                'End If
                Logger.WriteLine(ex)

                Throw
            End Try

            Return reader
        End Function

        ''' <summary>
        ''' ADO方式读取Excel文件里的数据
        ''' jet/ace 数据驱动引擎会根据前8行数据设置每列的数据类型
        ''' 建议使用using语句调用,以确保使用完之后关闭reader,并且要在用完Reader之后 End Using 语句之前调用 Reader的CloseDbConnection扩展方法,确保相应的数据库连接在关闭
        ''' </summary>
        ''' <param name="excelFullPath">Excel文件的全路径</param>
        ''' <param name="hdr">第一行是否作为标题</param>
        ''' <returns>返回所有数据的Reader，可以用索引或者列名读取</returns>
        Public Shared Function GetDataFromExcel(ByVal excelFullPath As String, ByVal hdr As HDRMode) As OleDbDataReader
            Return GetDataFromExcel(excelFullPath, DefaultExcelQuerySql, hdr)
        End Function

        ''' <summary>
        ''' ADO方式读取CSV(逗号分隔文件,逗号分隔的txt文件也可以)里的数据
        ''' jet/ace 数据驱动引擎会根据前8行数据设置每列的数据类型
        ''' </summary>
        ''' <param name="csvFileFullPath">CSV文件(逗号分隔文件,逗号分隔的txt文件也可以)的全路径</param>
        ''' <param name="codePage">文件的编码页</param>
        ''' <param name="sql">SQL查询语句</param>
        ''' <param name="dt">读取出来的数据要绑定的<paramref name="dt"/></param>
        ''' <param name="hdr">第一行是否作为标题</param>
        ''' <param name="showRowNumber">显示行号</param>
        ''' <param name="rowNumberColumnName">行号的名称(前面的 <paramref name="showRowNumber"/> 设置为 True，此参数才起作用)</param>
        Private Shared Sub InternalGetDataFromCsv(ByVal csvFileFullPath As String, ByVal codePage As CodePage, ByVal sql As String, ByRef dt As DataTable, ByVal hdr As HDRMode, ByVal showRowNumber As Boolean， ByVal rowNumberColumnName As String)
            PathHelper.EnsureNoStartWithEmptyChar(csvFileFullPath)

            Using reader = GetDataReaderFromCsv(csvFileFullPath, codePage, sql, hdr)
                ' 把读取到的数据绑定到dgv上
                dt = DataReader(reader)
                If showRowNumber Then
                    dt = SetAutoSerialNumber(dt, rowNumberColumnName)
                End If
            End Using
        End Sub

        ''' <summary>
        ''' ADO方式读取CSV(逗号分隔文件,逗号分隔的txt文件也可以)里的数据
        ''' jet/ace 数据驱动引擎会根据前8行数据设置每列的数据类型
        ''' </summary>
        ''' <param name="csvFileFullPath">CSV文件(逗号分隔文件,逗号分隔的txt文件也可以)的全路径</param>
        ''' <param name="codePage">文件的编码页</param>
        ''' <param name="sql">SQL查询语句</param>
        ''' <param name="dt">读取出来的数据要绑定的<paramref name="dt"/></param>
        ''' <param name="hdr">第一行是否作为标题</param>
        ''' <param name="showRowNumber">显示行号</param>
        ''' <param name="rowNumberColumnName">行号的名称(前面的 <paramref name="showRowNumber"/> 设置为 True，此参数才起作用)</param>
        Public Shared Sub GetDataFromCsv(ByVal csvFileFullPath As String, ByVal codePage As CodePage, ByVal sql As String, ByRef dt As DataTable, ByVal hdr As HDRMode, ByVal showRowNumber As Boolean， ByVal rowNumberColumnName As String)
            ' 如果sql里面包含 除法运算，那么得先把csv格式转成excel格式，再从excel查询
            ' 这是一条极其不严谨的正则（验证通过有除法参与的运算之后生成的新列）
            Dim pattern = ".*?\.*? as \w+"
            If Regex.IsMatch(sql, pattern, RegexOptions.IgnoreCase Or RegexOptions.Compiled) Then
                Dim cacheExcelFile As String
                Try
                    cacheExcelFile = FileConverter.CsvToExcel(csvFileFullPath, codePage)
                    If cacheExcelFile.Length = 0 Then
                        Throw New InvalidCastException("从CSV到Excel的转换失败")
                    End If
                    GetDataFromExcel(cacheExcelFile, sql, dt, HDRMode.Yes, showRowNumber)
                Catch ex As Exception
                    Logger.WriteLine(ex)
                Finally
                    ' 删除缓存文件
                    If cacheExcelFile.Length > 0 AndAlso IO.File.Exists(cacheExcelFile) Then
                        IO.File.Delete(cacheExcelFile)
                    End If
                End Try
            Else
                InternalGetDataFromCsv(csvFileFullPath, codePage, sql, dt, hdr, showRowNumber, rowNumberColumnName)
            End If
        End Sub

        ''' <summary>
        ''' ADO方式读取CSV(逗号分隔文件,逗号分隔的txt文件也可以)里的数据
        ''' jet/ace 数据驱动引擎会根据前8行数据设置每列的数据类型
        ''' </summary>
        ''' <param name="csvFileFullPath">CSV文件(逗号分隔文件,逗号分隔的txt文件也可以)的全路径</param>
        ''' <param name="codePage">文件的编码页</param>
        ''' <param name="sql">SQL查询语句</param>
        ''' <param name="dgv">读取出来的数据要绑定的<paramref name="dgv"/></param>
        ''' <param name="hdr">第一行是否作为标题</param>
        ''' <param name="showRowNumber">显示行号</param>
        ''' <param name="rowNumberColumnName">行号的名称(前面的 <paramref name="showRowNumber"/> 设置为 True，此参数才起作用)</param>
        Public Shared Sub GetDataFromCsv(ByVal csvFileFullPath As String, ByVal codePage As CodePage, ByVal sql As String, ByRef dgv As DataGridView, ByVal hdr As HDRMode, ByVal showRowNumber As Boolean， ByVal rowNumberColumnName As String)
			Dim dt As DataTable = Nothing
			GetDataFromCsv(csvFileFullPath, codePage, sql, dt, hdr, showRowNumber, rowNumberColumnName)
			SetDataToDataGridView(dgv, dt)
		End Sub

		''' <summary>
		''' ADO方式读取CSV(逗号分隔文件,逗号分隔的txt文件也可以)里的数据
		''' jet/ace 数据驱动引擎会根据前8行数据设置每列的数据类型
		''' </summary>
		''' <param name="csvFileFullPath">CSV文件(逗号分隔文件,逗号分隔的txt文件也可以)的全路径</param>
		''' <param name="codePage">文件的编码页</param>
		''' <param name="dgv">读取出来的数据要绑定的<paramref name="dgv"/></param>
		''' <param name="hdr">第一行是否作为标题</param>
		''' <param name="showRowNumber">显示行号</param>
		''' <param name="rowNumberColumnName">行号的名称(前面的 <paramref name="showRowNumber"/> 设置为 True，此参数才起作用)</param>
		Public Shared Sub GetDataFromCsv(ByVal csvFileFullPath As String, ByVal codePage As CodePage, ByRef dgv As DataGridView, ByVal hdr As HDRMode, ByVal showRowNumber As Boolean， ByVal rowNumberColumnName As String)
			Dim csvFileName = Path.GetFileName(csvFileFullPath)
			Dim sql = $"select * from [{csvFileName}]"
			GetDataFromCsv(csvFileFullPath, codePage, sql, dgv, hdr, showRowNumber, rowNumberColumnName)
		End Sub

		''' <summary>
		''' ADO方式读取CSV(逗号分隔文件,逗号分隔的txt文件也可以)里的数据
		''' jet/ace 数据驱动引擎会根据前8行数据设置每列的数据类型
		''' </summary>
		''' <param name="csvFileFullPath">CSV文件(逗号分隔文件,逗号分隔的txt文件也可以)的全路径（默认使用UTF-8编码读取文件）</param>
		''' <param name="sql">SQL查询语句</param>
		''' <param name="dgv">读取出来的数据要绑定的<paramref name="dgv"/></param>
		''' <param name="hdr">第一行是否作为标题</param>
		''' <param name="showRowNumber">显示行号(行号列名默认为 "序号")</param>
		Public Shared Sub GetDataFromCsv(ByVal csvFileFullPath As String, ByVal sql As String, ByRef dgv As DataGridView, ByVal hdr As HDRMode, ByVal showRowNumber As Boolean)
			Call GetDataFromCsv(csvFileFullPath, CodePage.UTF8, sql, dgv, hdr, showRowNumber, "序号")
		End Sub

		''' <summary>
		''' ADO方式读取CSV(逗号分隔文件,逗号分隔的txt文件也可以)里的数据
		''' jet/ace 数据驱动引擎会根据前8行数据设置每列的数据类型
		''' </summary>
		''' <param name="csvFileFullPath">CSV文件(逗号分隔文件,逗号分隔的txt文件也可以)的全路径</param>
		''' <param name="codePage">文件的编码页</param>
		''' <param name="dgv">读取出来的数据要绑定的<paramref name="dgv"/></param>
		''' <param name="hdr">第一行是否作为标题</param>
		''' <param name="showRowNumber">显示行号(行号列名默认为 "序号")</param>
		Public Shared Sub GetDataFromCsv(ByVal csvFileFullPath As String, ByVal codePage As CodePage, ByRef dgv As DataGridView, ByVal hdr As HDRMode, ByVal showRowNumber As Boolean)
			Dim csvFileName = Path.GetFileName(csvFileFullPath)
			Dim sql = $"select * from [{csvFileName}]"
			Call GetDataFromCsv(csvFileFullPath, codePage, sql, dgv, hdr, showRowNumber, "序号")
		End Sub

		''' <summary>
		''' ADO方式读取CSV(逗号分隔文件,逗号分隔的txt文件也可以)里的数据
		''' jet/ace 数据驱动引擎会根据前8行数据设置每列的数据类型
		''' </summary>
		''' <param name="csvFileFullPath">CSV文件(逗号分隔文件,逗号分隔的txt文件也可以)的全路径（默认使用UTF-8编码读取文件）</param>
		''' <param name="dgv">读取出来的数据要绑定的<paramref name="dgv"/></param>
		''' <param name="hdr">第一行是否作为标题</param>
		''' <param name="showRowNumber">显示行号(行号列名默认为 "序号")</param>
		Public Shared Sub GetDataFromCsv(ByVal csvFileFullPath As String, ByRef dgv As DataGridView, ByVal hdr As HDRMode, ByVal showRowNumber As Boolean)
			Dim csvFileName = Path.GetFileName(csvFileFullPath)
			Dim sql = $"select * from [{csvFileName}]"
			Call GetDataFromCsv(csvFileFullPath, CodePage.UTF8, sql, dgv, hdr, showRowNumber, "序号")
		End Sub

		''' <summary>
		''' ADO方式读取CSV文件里的数据
		''' jet/ace 数据驱动引擎会根据前8行数据设置每列的数据类型
		''' </summary>
		''' <param name="csvFileFullPath">CSV文件的全路径</param>
		''' <param name="codePage">文件<paramref name="csvFileFullPath"/>的编码</param>
		''' <param name="sql">SQL查询语句</param>
		''' <param name="dgv">读取出来的数据要绑定的<paramref name="dgv"/></param>
		''' <param name="hdr">第一行是否作为标题</param>
		''' <param name="showRowNumber">显示行号</param>
		Public Shared Sub GetDataFromCsv(ByVal csvFileFullPath As String, ByVal codePage As CodePage, ByVal sql As String, ByRef dgv As DataGridView, ByVal hdr As HDRMode, ByVal showRowNumber As Boolean)
			GetDataFromCsv(csvFileFullPath, codePage, sql, dgv, hdr, showRowNumber, "序号")
		End Sub

		''' <summary>
		''' ADO方式读取CSV文件里的数据
		''' jet/ace 数据驱动引擎会根据前8行数据设置每列的数据类型
		''' </summary>
		''' <param name="csvFileFullPath">CSV文件的全路径</param>
		''' <param name="codePage">文件<paramref name="csvFileFullPath"/>的编码</param>
		''' <param name="dt">读取出来的数据要绑定的<paramref name="dt"/></param>
		''' <param name="hdr">第一行是否作为标题</param>
		''' <param name="showRowNumber">显示行号</param>
		Public Shared Sub GetDataFromCsv(ByVal csvFileFullPath As String, ByVal codePage As CodePage, ByRef dt As DataTable, ByVal hdr As HDRMode, ByVal showRowNumber As Boolean)
			GetDataFromCsv(csvFileFullPath, codePage, GetDelimitedFileQuerySql(csvFileFullPath), dt, hdr, showRowNumber)
		End Sub

		''' <summary>
		''' ADO方式读取CSV文件里的数据
		''' jet/ace 数据驱动引擎会根据前8行数据设置每列的数据类型
		''' </summary>
		''' <param name="csvFileFullPath">CSV文件的全路径</param>
		''' <param name="codePage">文件<paramref name="csvFileFullPath"/>的编码</param>
		''' <param name="sql">SQL查询语句</param>
		''' <param name="dt">读取出来的数据要绑定的<paramref name="dt"/></param>
		''' <param name="hdr">第一行是否作为标题</param>
		''' <param name="showRowNumber">显示行号</param>
		Public Shared Sub GetDataFromCsv(ByVal csvFileFullPath As String, ByVal codePage As CodePage, ByVal sql As String, ByRef dt As DataTable, ByVal hdr As HDRMode, ByVal showRowNumber As Boolean)
			GetDataFromCsv(csvFileFullPath, codePage, sql, dt, hdr, showRowNumber, "序号")
		End Sub

        ''' <summary>
        ''' 使用DataReader效率比DataAdapter高
        ''' </summary>
        ''' <param name="oleCmd"></param>
        ''' <returns></returns>
        Private Shared Function DataReader(ByVal oleCmd As OleDbCommand) As DataTable
            Dim funcRst As DataTable

            Try
                Using reader As OleDbDataReader = oleCmd.ExecuteReader(CommandBehavior.CloseConnection)
                    funcRst = DataReader(reader)
                End Using
            Catch ex As Exception
                Logger.WriteLine(ex)
                Throw
            End Try

            Return If(funcRst, New DataTable)
        End Function

        ''' <summary>
        ''' 使用DataReader效率比DataAdapter高
        ''' </summary>
        ''' <param name="dbReader"></param>
        ''' <returns></returns>
        Private Shared Function DataReader(ByVal dbReader As DbDataReader) As DataTable
            Dim funcRst As New DataTable

            Try
                Using reader = dbReader
                    Dim col As DataColumn
                    'Dim row As DataRow

                    ' 获取列名已经列数据类型信息
                    ' jet/ace 数据驱动引擎会根据前8行数据设置每列的数据类型
                    For i As Integer = 0 To reader.FieldCount - 1
                        col = New DataColumn With {
                            .ColumnName = reader.GetName(i),
                            .DataType = reader.GetFieldType(i)
                        }

                        funcRst.Columns.Add(col)
                    Next

                    funcRst.BeginLoadData()

                    ' 获取行数据
                    While reader.Read
                        'row = funcRst.NewRow
                        Dim row = New Object(reader.FieldCount - 1) {}
                        For i As Integer = 0 To reader.FieldCount - 1
                            row(i) = reader(i)
                        Next
                        funcRst.LoadDataRow(row, True)
                        'funcRst.Rows.Add(row)
                    End While

                    funcRst.EndLoadData()
                End Using

            Catch ex As Exception
                Logger.WriteLine(ex)
                Throw
            End Try

            Return funcRst
        End Function

        ''' <summary>
        ''' ADO方式读取CSV(逗号分隔文件,逗号分隔的txt文件也可以)里的数据
        ''' jet/ace 数据驱动引擎会根据前8行数据设置每列的数据类型
        ''' 注：建议使用using语句调用,以确保使用完之后关闭reader,并且要在用完Reader之后 End Using 语句之前调用 Reader的CloseDbConnection扩展方法,确保相应的数据库连接在关闭
        ''' </summary>
        ''' <param name="csvFileFullPath">CSV文件(逗号分隔文件,逗号分隔的txt文件也可以)的全路径</param>
        ''' <param name="codePage">文件的编码页</param>
        ''' <param name="sql">SQL查询语句</param>
        ''' <param name="hdr">第一行是否作为标题</param>
        Public Shared Function GetDataReaderFromCsv(ByVal csvFileFullPath As String, ByVal codePage As CodePage, ByVal sql As String, ByVal hdr As HDRMode) As DbDataReader
            PathHelper.EnsureNoStartWithEmptyChar(csvFileFullPath)

            Dim reader As DbDataReader

#Region "ADO方式读取Csv(逗号分隔文件,逗号分隔的txt文件也可以)里的数据"
            Dim oleCn As New OleDb.OleDbConnection

			Try
#Region "帮助信息"
				' 如果出现 “未在本地计算机上注册“Microsoft.Jet.OLEDB.4.0”提供程序” 类似错误
				' 在解决方案上右键-属性-选择“编译”选项卡--点击 “高级编译选项”-弹出 高级编译设置窗口   ------在最下面的 “选择目标 CPU”---选择X86就可以 了
				' 另外，如果机器上装的是WPS,而且确实没有这个驱动的话，改成X86也依然有这个提示，这个时候就需要安装这个驱动了
				' 下载地址 https://www.microsoft.com/zh-CN/download/details.aspx?id=13255
				' 选择相应的驱动器
				' 最新测试，不管是xlsx还是xls，Microsoft.ACE.OLEDB.12.0驱动都可以读取，
				' 如果office是32位那就装 AccessDatabaseEngine.exe 驱动
				' 如果是64位，那就装 AccessDatabaseEngine_X64.exe 驱动
				' 项目属性——编译——选择AnyCpu，勾选首选32位
				' 20180426

				'    '用ADO查询
				'    '################################################
				'    '方法二
				'    '条件一：文本之间得有某种分隔符,FMT=Delimited(,)这里指定
				'    '条件二：HDR=NO,这里指定不以第一行为字段名,适合文件目录下“test.txt”这个文档的格式
				'    '其他参数,请去各种搜索引擎获知
				'    'Data Source 到文件夹即可 不需要具体文件名
				'    CharacterSet=65001 解决读取UTF-8 文件乱码问题
				'    '################################################
#End Region
				' HDR=Yes 第一行作为标题
				Dim csvFilePath = Path.GetDirectoryName(csvFileFullPath)
				oleCn.ConnectionString = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={csvFilePath};Extended Properties='text;HDR={hdr.ToString()};IMEX=1;FMT=Delimited(,);CharacterSet={CStr(codePage)}'"

                Using oleCmd As New OleDbCommand(sql, oleCn)
                    ' 确保连接打开
                    If oleCn.State = ConnectionState.Closed Then
                        oleCn.Open()
                    End If

                    reader = oleCmd.ExecuteReader(CommandBehavior.CloseConnection)
                End Using
#End Region
            Catch ex As InvalidOperationException
				Logger.WriteLine(ex)
				If ex.Message.IndexOf(".OLEDB.", StringComparison.OrdinalIgnoreCase) > -1 Then
					Throw New EngineNotFoundExcption(ex.Message, ex)
				Else
					Throw
				End If
			Catch ex As Exception
				reader?.Close()
				'微软已知bug,遇到可跳过
				'If Err.Number = -2147467259 And Err.Description = "系统不支持所选择的排序。" Then
				'    Application.DoEvents()
				'End If
				Logger.WriteLine(ex)

				Throw
			End Try

			Return reader
		End Function

		''' <summary>
		''' ADO方式读取CSV(逗号分隔文件,逗号分隔的txt文件也可以)里的数据
		''' jet/ace 数据驱动引擎会根据前8行数据设置每列的数据类型
		''' 注：建议使用using语句调用,以确保使用完之后关闭reader,并且要在用完Reader之后 End Using 语句之前调用 Reader的CloseDbConnection扩展方法,确保相应的数据库连接在关闭
		''' </summary>
		''' <param name="csvFileFullPath">CSV文件(逗号分隔文件,逗号分隔的txt文件也可以)的全路径</param>
		''' <param name="codePage">文件的编码页</param>
		''' <param name="hdr">第一行是否作为标题</param>
		Public Shared Function GetDataReaderFromCsv(ByVal csvFileFullPath As String, ByVal codePage As CodePage, ByVal hdr As HDRMode) As DbDataReader
			Return GetDataReaderFromCsv(csvFileFullPath, codePage, GetDelimitedFileQuerySql(csvFileFullPath), hdr)
		End Function

		Private Shared Function GetDelimitedFileQuerySql(ByVal fileFullPath As String) As String
			Dim ext = IO.Path.GetExtension(fileFullPath)
			ext = If(ext.Length > 0 AndAlso "."c = ext.Chars(0),
				ext.Substring(1, ext.Length - 1),
				"csv")
			Dim sql = $"select * from [{IO.Path.GetFileNameWithoutExtension(fileFullPath)}#{ext}]"
			Return sql
		End Function
#End Region

	End Class
End Namespace