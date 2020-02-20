Imports System.IO
Imports ShanXingTech.Text2

Namespace ShanXingTech.IO2
	Public NotInheritable Class Writer
		''' <summary>
		''' 把数据写入txt
		''' </summary>
		''' <param name="path">txt路径</param>
		''' <param name="value">具体值</param>
		''' <param name="access">写文件的方式（默认追加方式）</param>
		''' <param name="codepage">代码页标识符</param>
		Public Shared Sub WriteText(ByVal path As String， ByVal value As String, ByVal access As FileMode, ByVal codepage As CodePage)
			' 确保存在缓存目录 以便写出文件头 
			If Not System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(path)) Then
				Throw New DirectoryNotFoundException("缓存路径不存在,数据无法正常导出！  文件名:" & path)
			End If

			' 去除最后的空行
			' 应直接写文件，不更改文件内容，弃用20190712
			'If value.Length > 1 AndAlso
			'    value.EndsWith(Environment.NewLine) Then
			'    value = value.Remove(value.Length - Environment.NewLine.Length, Environment.NewLine.Length)
			'End If

			' 只写共享模式，可以解决 正由另一进程使用，因此该进程无法访问此文件 问题
			' 不需要写两个Using块，fs的内存会在Using sw块结束之后自动被释放
			' vs的代码分析 和clr via C# 都有提到介个问题
			Dim fs As New FileStream(path, access, FileAccess.Write, FileShare.Write)
			'声明数据流文件写入方法  
			Using sw As New StreamWriter(fs, Text.Encoding.GetEncoding(codepage))
				sw.WriteLine(value)
			End Using
		End Sub

		''' <summary>
		''' 把数据写入txt，默认UTF8编码写文件
		''' </summary>
		''' <param name="path">txt路径</param>
		''' <param name="value">具体值</param>
		''' <param name="access">写文件的方式（默认追加方式）</param>
		Public Shared Sub WriteText(ByVal path As String， ByVal value As String, ByVal access As FileMode)
			WriteText(path, value, access, CodePage.UTF8)
		End Sub

		''' <summary>
		''' 把数据写入txt，默认UTF8编码写文件
		''' </summary>
		''' <param name="path">txt路径</param>
		''' <param name="value">具体值</param>
		Public Shared Sub WriteText(ByVal path As String， ByVal value As String)
			WriteText(path, value, FileMode.Append, CodePage.UTF8)
		End Sub

		''' <summary>
		''' 把数据写入txt
		''' </summary>
		''' <param name="path">txt路径</param>
		''' <param name="sb">StringBuilder</param>
		''' <param name="access">写文件的方式（默认追加方式）</param>
		''' <param name="codepage">代码页标识符</param>
		Public Shared Sub WriteText(ByVal path As String， ByVal sb As Text.StringBuilder, ByVal access As FileMode, ByVal codepage As CodePage)
			' 确保存在缓存目录 以便写出文件头 
			If Not System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(path)) Then
				Throw New DirectoryNotFoundException("缓存路径不存在,数据无法正常导出！  文件名:" & path)
			End If

			Dim value = StringBuilderCache.GetStringAndReleaseBuilder(sb)
			WriteText(path, value, access, CodePage.UTF8)
		End Sub

		''' <summary>
		''' 把数据写入txt，默认UTF8编码写文件
		''' </summary>
		''' <param name="path">txt路径</param>
		''' <param name="sb">StringBuilder</param>
		''' <param name="access">写文件的方式（默认追加方式）</param>
		Public Shared Sub WriteText(ByVal path As String， ByVal sb As Text.StringBuilder, ByVal access As FileMode)
			WriteText(path, sb, access, CodePage.UTF8)
		End Sub

		''' <summary>
		''' 把数据写入txt，默认UTF8编码写文件
		''' </summary>
		''' <param name="path">txt路径</param>
		''' <param name="sb">StringBuilder</param>
		Public Shared Sub WriteText(ByVal path As String， ByVal sb As Text.StringBuilder)
			WriteText(path, sb, FileMode.Append, CodePage.UTF8)
		End Sub
	End Class
End Namespace

