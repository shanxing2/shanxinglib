Imports System.IO
Imports System.Runtime.CompilerServices

Namespace ShanXingTech
	Public NotInheritable Class Logger
#Region "字段区"
		''' <summary>
		''' 上一次使用Log类的日期
		''' </summary>
		Private Shared m_LastDate As Date = Date.Now
		Private Shared ReadOnly m_DateFormat As String = "yyyy-MM-dd HH:mm:ss"
		Private Shared m_LogFile As String = $"C:\ShanXingTech\Log\log{Date.Now.ToString("MMdd")}.log"

#End Region

#Region "属性区"
		''' <summary>
		''' Logger启动状态。默认为 <see cref="LoggerStatus.On"/>,输出log信息。
		''' </summary>
		''' <returns></returns>
		Public Shared Property Status As LoggerStatus = LoggerStatus.On
#End Region

		''' <summary>
		''' 生成调试信息函数
		''' </summary>
		''' <param name="logString">要输出的信息</param>
		''' <param name="callerMemberName"></param>
		''' <param name="callerFilePath"></param>
		''' <param name="callerLineNumber"></param>
		''' <returns></returns>
		Public Shared Function MakeDebugString(ByVal logString As String, <CallerMemberName> Optional callerMemberName As String = Nothing, <CallerFilePath> Optional callerFilePath As String = Nothing, <CallerLineNumber()> Optional callerLineNumber As Integer = 0) As String
			Dim debugString = $"Time：{Date.Now.ToString(m_DateFormat)} {My.Resources.CallSubName}{If(callerFilePath, String.Empty)} {callerMemberName} {My.Resources.FileLineNumber}{callerLineNumber}    {logString}"

			Return debugString
		End Function

		''' <summary>
		''' 生成调试信息函数
		''' </summary>
		''' <param name="exception">异常对象</param>
		''' <param name="callerMemberName"></param>
		''' <param name="callerFilePath"></param>
		''' <param name="callerLineNumber"></param>
		Public Shared Function MakeDebugString(Of T As Exception)(ByRef exception As T, <CallerMemberName> Optional callerMemberName As String = Nothing, <CallerFilePath> Optional callerFilePath As String = Nothing, <CallerLineNumber()> Optional callerLineNumber As Integer = 0) As String
			' 构造输出日记信息
			Dim logWriteString = MakeLogString(exception.Message, callerMemberName, callerFilePath, callerLineNumber)

			logWriteString = $"{logWriteString}{My.Resources.ProblemDetail}{Environment.NewLine}{exception.StackTrace}{Environment.NewLine}"
			If exception.InnerException IsNot Nothing Then
				logWriteString = $"{logWriteString}{My.Resources.ProblemDescription}{exception.InnerException.Message}{Environment.NewLine}{My.Resources.ProblemDetail}{Environment.NewLine}{exception.InnerException.StackTrace}{Environment.NewLine}"
			End If

			Return logWriteString
		End Function

		Private Shared Function MakeLogString(ByVal logString As String, <CallerMemberName> Optional callerMemberName As String = Nothing, <CallerFilePath> Optional callerFilePath As String = Nothing, <CallerLineNumber()> Optional callerLineNumber As Integer = 0) As String
			' 构造输出日记信息
			Dim stackInfo = $"   Time：{Date.Now.ToString(m_DateFormat)}{Environment.NewLine}{My.Resources.CallSubName.PadLeftByByte(9)}{callerFilePath} {callerMemberName} {My.Resources.FileLineNumber}{callerLineNumber}"
			Dim logWriteString = $"{stackInfo}{Environment.NewLine}{My.Resources.ProblemDescription}{logString}{Environment.NewLine}"

			Return logWriteString
		End Function

		Private Shared Function MakeLogString(Of T As Exception)(ByRef exception As T, <CallerMemberName> Optional callerMemberName As String = Nothing, <CallerFilePath> Optional callerFilePath As String = Nothing, <CallerLineNumber()> Optional callerLineNumber As Integer = 0) As String
			' 构造输出日记信息
			Dim logWriteString = MakeLogString(exception.Message, callerMemberName, callerFilePath, callerLineNumber)

			logWriteString = $"{logWriteString}{My.Resources.ProblemDetail}{Environment.NewLine}{exception.StackTrace}{Environment.NewLine}"
			If exception.InnerException IsNot Nothing Then
				logWriteString = $"{logWriteString}{My.Resources.ProblemDescription}{Environment.NewLine}{exception.InnerException.Message}{Environment.NewLine}{My.Resources.ProblemDetail}{Environment.NewLine}{exception.InnerException.StackTrace}{Environment.NewLine}"
			End If

			Return logWriteString
		End Function

		''' <summary>
		''' 调试输出函数
		''' </summary>
		''' <param name="exception"></param>
		''' <param name="callerMemberName"></param>
		''' <param name="callerFilePath"></param>
		''' <param name="callerLineNumber"></param>
		Public Shared Sub Debug(Of T As Exception)(ByRef exception As T, <CallerMemberName> Optional callerMemberName As String = Nothing, <CallerFilePath> Optional callerFilePath As String = Nothing, <CallerLineNumber()> Optional callerLineNumber As Integer = 0)
			If Status = LoggerStatus.Off Then Return

			Diagnostics.Debug.Print(MakeDebugString(exception, callerMemberName, callerFilePath, callerLineNumber))
		End Sub

		''' <summary>
		''' 调试输出函数
		''' </summary>
		''' <param name="logString">要输出的信息</param>
		''' <param name="callerMemberName"></param>
		''' <param name="callerFilePath"></param>
		''' <param name="callerLineNumber"></param>
		Public Shared Sub Debug(ByVal logString As String, <CallerMemberName> Optional callerMemberName As String = Nothing, <CallerFilePath> Optional callerFilePath As String = Nothing, <CallerLineNumber()> Optional callerLineNumber As Integer = 0)
			If Status = LoggerStatus.Off Then Return

			' 构造输出日记信息
			Diagnostics.Debug.Print(MakeDebugString(logString, callerMemberName, callerFilePath, callerLineNumber))
		End Sub

		''' <summary>
		''' 写日志函数
		''' </summary>
		''' <param name="exception">异常对象</param>
		''' <param name="callerMemberName"></param>
		''' <param name="callerFilePath"></param>
		''' <param name="callerLineNumber"></param>
		Public Shared Sub WriteLine(Of T As Exception)(ByRef exception As T, <CallerMemberName> Optional callerMemberName As String = Nothing, <CallerFilePath> Optional callerFilePath As String = Nothing, <CallerLineNumber()> Optional callerLineNumber As Integer = 0)
			If Status = LoggerStatus.Off Then Return

			WriteLine(MakeLogString(exception, callerMemberName, callerFilePath, callerLineNumber))
		End Sub

		''' <summary>
		''' 写日志函数
		''' </summary>
		''' <param name="logString">要输出的信息</param>
		''' <param name="callerMemberName"></param>
		''' <param name="callerFilePath"></param>
		''' <param name="callerLineNumber"></param>
		Public Shared Sub WriteLine(ByVal logString As String, <CallerMemberName> Optional callerMemberName As String = Nothing, <CallerFilePath> Optional callerFilePath As String = Nothing, <CallerLineNumber()> Optional callerLineNumber As Integer = 0)
			If Status = LoggerStatus.Off Then Return

			' 构造输出日记信息
			Dim logWriteString = MakeLogString(logString, callerMemberName, callerFilePath, callerLineNumber)

			Diagnostics.Debug.Print(logWriteString)
			WriteLine(logWriteString)
		End Sub

		''' <summary>
		''' 写日志函数
		''' </summary>
		''' <param name="exception">异常对象</param>
		''' <param name="callerMemberName"></param>
		''' <param name="callerFilePath"></param>
		''' <param name="callerLineNumber"></param>
		Public Shared Sub WriteLine(Of T As Exception)(ByRef exception As T, ByVal logString As String, <CallerMemberName> Optional callerMemberName As String = Nothing, <CallerFilePath> Optional callerFilePath As String = Nothing, <CallerLineNumber()> Optional callerLineNumber As Integer = 0)
			If Status = LoggerStatus.Off Then Return

			' 构造输出日记信息
			Dim logWriteString = MakeLogString(exception, callerMemberName, callerFilePath, callerLineNumber)
			logWriteString = String.Concat(logWriteString, logString， Environment.NewLine， Environment.NewLine)

			Diagnostics.Debug.Print(logWriteString)
			WriteLine(logWriteString)
		End Sub

		''' <summary>
		''' 写日志函数
		''' </summary>
		''' <param name="logWriteString">要输出的信息</param>
		Private Shared Sub WriteLine(ByVal logWriteString As String)
			Try
				MakeNowDayLogFileName()

				' 文件存在则覆盖
				' 文件不存在则创建
				' 创建文件之前 必须保证目录存在，否则会产生目录不存在错误
				If Not Directory.Exists(Path.GetDirectoryName(m_LogFile)) Then
					Directory.CreateDirectory(Path.GetDirectoryName(m_LogFile))
				End If

				IO2.Writer.WriteText(m_LogFile, logWriteString, FileMode.Append, IO2.CodePage.UTF8)
			Catch ex As Exception
				Debug(ex)
				Throw
			End Try
		End Sub

		''' <summary>
		''' 生成当日Log文件路径名
		''' </summary>
		Private Shared Sub MakeNowDayLogFileName()
			If m_LastDate.Day < Date.Now.Day OrElse
				m_LastDate.Month < Date.Now.Month Then
				m_LastDate = Date.Now
				m_LogFile = $"C:\ShanXingTech\Log\log{Date.Now.ToString("MMdd")}.log"
			End If
		End Sub
	End Class
End Namespace
