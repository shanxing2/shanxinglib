Imports System.IO.Compression
Imports System.Text

Namespace ShanXingTech.IO2.Compression
	''' <summary>
	''' 这个人很懒，什么都没写
	''' </summary>
	Public Class Zip
		' 明文：神即道, 道法自然, 如来|闪星网络信息科技 ShanXingTech Q2287190283
		' 算法：古典密码中的有密钥换位密码 密钥：ShanXingTech
		Public Const ShanXingTechQ2287190283 = "神闪X7,SQB道信T2道网N9来A2D如H2C然技HA即星I1|N8E法息E8,络G0自科C3"

		''' <summary>
		''' 如果压缩或者解压缩传入的参数不合法则抛出异常
		''' </summary>
		''' <param name="directoryName">文件夹名</param>
		''' <param name="zipFileFullPath">压缩文件名</param>
		Private Shared Sub ThrowIfParamsIllegal(ByRef directoryName As String, ByRef zipFileFullPath As String)
			If directoryName Is Nothing Then
				Throw New ArgumentNullException(String.Format(My.Resources.NullReference, NameOf(directoryName)))
			End If

			If zipFileFullPath Is Nothing Then
				Throw New ArgumentNullException(String.Format(My.Resources.NullReference, NameOf(zipFileFullPath)))
			End If

			If directoryName.Length = 0 Then
				Throw New ArgumentException("文件夹名无效")
			End If

			If zipFileFullPath.Length = 0 Then
				Throw New ArgumentException("文件名无效")
			End If

			Try
				' 如果得到的文件名没有包含后缀，则自动补上后缀 “.zip”
				Dim tempFileName = IO.Path.GetFileName(zipFileFullPath)
				If tempFileName.IndexOf(".") = -1 Then
					zipFileFullPath += ".zip"
				End If
			Catch ex As Exception
				Throw
			End Try
		End Sub

		''' <summary>
		''' 解压文件夹
		''' 使用中的文件不能被覆盖，而且会引发错误
		''' </summary>
		''' <param name="sourceFileFullPath">zip文件名</param>
		''' <param name="destinationDirectoryName">输出的解压文件夹名</param>
		''' <param name="encoding">在压缩文件中读取项名时使用的编码</param>
		Public Shared Sub UnZip(ByVal sourceFileFullPath As String, ByVal destinationDirectoryName As String, ByVal encoding As Encoding)
			ThrowIfParamsIllegal(destinationDirectoryName, sourceFileFullPath)

			Try
				' System.Text.Encoding.Default解决中文文件名解压后乱码问题
				IO.Compression.ZipFile.ExtractToDirectory(sourceFileFullPath, destinationDirectoryName, encoding)
			Catch ex As Exception
				Throw
			End Try
		End Sub

		''' <summary>
		''' 解压文件夹
		''' 使用中的文件不能被覆盖，而且会引发错误
		''' </summary>
		''' <param name="sourceFileFullPath">zip文件名</param>
		''' <param name="destinationDirectoryName">输出的解压文件夹名</param>
		Public Shared Sub UnZip(ByVal sourceFileFullPath As String, ByVal destinationDirectoryName As String)
			UnZip(sourceFileFullPath, destinationDirectoryName, System.Text.Encoding.Default)
		End Sub

		''' <summary>
		''' 压缩文件夹
		''' </summary>
		''' <param name="sourceDirectoryName">需要压缩的文件夹名</param>
		''' <param name="destinationFileFullPath">输出的压缩文件名</param>
		''' <param name="encoding">在压缩文件中写入项名时使用的编码</param>
		Public Shared Sub ZipDirectory(ByVal sourceDirectoryName As String, ByVal destinationFileFullPath As String, ByVal encoding As Encoding)
			ThrowIfParamsIllegal(sourceDirectoryName, destinationFileFullPath)

			' 如果已经存到压缩文件，则直接抛出异常，而不是由 CreateFromDirectory 函数抛出异常
			' 因为 CreateFromDirectory 会创建一个压缩文件，然后再抛出异常
			Dim isExists = IO.File.Exists(destinationFileFullPath)
			If isExists Then
				Throw New IO.IOException("压缩文件已经存在")
			End If

			Try
				' System.Text.Encoding.Default解决中文文件名解压后乱码问题
				IO.Compression.ZipFile.CreateFromDirectory(sourceDirectoryName, destinationFileFullPath, CompressionLevel.Optimal, False, encoding)
			Catch ex As Exception
				Throw
			End Try
		End Sub

		''' <summary>
		''' 压缩文件夹
		''' </summary>
		''' <param name="sourceDirectoryName">需要压缩的文件夹名</param>
		''' <param name="destinationFileFullPath">输出的压缩文件名</param>
		Public Shared Sub ZipDirectory(ByVal sourceDirectoryName As String, ByVal destinationFileFullPath As String)
			ZipDirectory(sourceDirectoryName, destinationFileFullPath, Text.Encoding.Default)
		End Sub

		''' <summary>
		''' 压缩文件
		''' </summary>
		''' <param name="sourceFileFullPath">需要压缩的文件名</param>
		''' <param name="destinationFileFullPath">输出的压缩文件名</param>
		''' <param name="encoding">在压缩文件中写入项名时使用的编码</param>
		Public Shared Sub ZipFile(ByVal sourceFileFullPath As String, ByVal destinationFileFullPath As String, ByVal encoding As Encoding)
			' 如果已经存到压缩文件，则直接抛出异常，而不是由 CreateFromDirectory 函数抛出异常
			' 因为 CreateFromDirectory 会创建一个压缩文件，然后再抛出异常
			Dim isExists = IO.File.Exists(destinationFileFullPath)
			If isExists Then
				Throw New IO.IOException("压缩文件已经存在")
			End If

			Try
				Dim fileName = IO.Path.GetFileName(sourceFileFullPath)
				Using archive = IO.Compression.ZipFile.Open(destinationFileFullPath, ZipArchiveMode.Create, encoding)
					archive.CreateEntryFromFile(sourceFileFullPath, fileName, CompressionLevel.Optimal)
				End Using
			Catch ex As Exception
				Throw
			End Try
		End Sub

		''' <summary>
		''' 确定指定的文件是否存在
		''' </summary>
		''' <param name="fileName">文件名（带后缀，如 a.txt）</param>
		''' <param name="destinationFileFullPath">已经存在的压缩文件名</param>
		''' <param name="encoding">在压缩文件中写入项名时使用的编码</param>
		''' <returns></returns>
		Public Shared Function Exists(ByVal fileName As String, ByVal destinationFileFullPath As String, ByVal encoding As Encoding) As Boolean
			Dim funcRst As Boolean

			Using archive = IO.Compression.ZipFile.Open(destinationFileFullPath, ZipArchiveMode.Read, encoding)
				' 如果压缩文件内没有文件，直接返回false
				If archive.Entries.Count > 0 Then
					Dim entry = archive.GetEntry(fileName)
					funcRst = entry IsNot Nothing
				End If
			End Using

			Return funcRst
		End Function

		''' <summary>
		''' 把文件添加到已经存在的压缩文件中
		''' </summary>
		''' <param name="sourceFileFullPath">需要加进压缩文件中的文件名</param>
		''' <param name="destinationFileFullPath">已经存在的压缩文件名</param>
		''' <param name="encoding">在压缩文件中写入项名时使用的编码</param>
		Public Shared Sub Add(ByVal sourceFileFullPath As String, ByVal destinationFileFullPath As String, ByVal encoding As Encoding)
			' 如果压缩文件中已经有了同样的文件，则不再添加
			Dim fileName = IO.Path.GetFileName(sourceFileFullPath)
			If Exists(fileName, destinationFileFullPath, encoding) Then
				Throw New IO.IOException($"压缩文件内已经有同名文件 '{fileName}'")
			End If

			Using archive = IO.Compression.ZipFile.Open(destinationFileFullPath, ZipArchiveMode.Update, encoding)
				archive.CreateEntryFromFile(sourceFileFullPath, fileName, CompressionLevel.Optimal)
			End Using
		End Sub

		''' <summary>
		''' 把文件从已经存在的压缩文件中移除
		''' </summary>
		''' <param name="findFileFullName">需要从压缩文件中移除的文件名</param>
		''' <param name="destinationFileFullPath">已经存在的压缩文件名</param>
		''' <param name="encoding">在压缩文件中写入项名时使用的编码</param>
		Public Shared Sub Remove(ByVal findFileFullName As String, ByVal destinationFileFullPath As String, ByVal encoding As Encoding)
			' 如果压缩文件中没有同样的文件，则不移除
			Dim fileName = IO.Path.GetFileName(findFileFullName)
			If Not Exists(fileName, destinationFileFullPath, encoding) Then
				Throw New IO.FileNotFoundException("压缩文件内没有同名文件", fileName)
			End If

			Using archive = IO.Compression.ZipFile.Open(destinationFileFullPath, ZipArchiveMode.Update, encoding)
				Dim entry = archive.GetEntry(fileName)
				entry.Delete()
			End Using
		End Sub

		''' <summary>
		''' 获取压缩文件中所有文件
		''' </summary>
		''' <param name="zipFileFullPath">压缩文件名</param>
		''' <param name="encoding">在压缩文件中读取项名时使用的编码</param>
		Public Shared Iterator Function GetFiles(ByVal zipFileFullPath As String, ByVal encoding As Encoding) As IEnumerable(Of FileInfo)
			Using archive = IO.Compression.ZipFile.Open(zipFileFullPath, ZipArchiveMode.Read, encoding)
				For Each entry In archive.Entries
					Dim fileInfo As New FileInfo With {
					.CompressedLength = entry.CompressedLength,
					.FullName = entry.FullName,
					.LastWriteTime = entry.LastWriteTime,
					.Length = entry.Length,
					.FileName = entry.Name
				}

					Yield fileInfo
				Next
			End Using
		End Function
		''' <summary>
		''' 解压到当前文件夹
		''' 使用中的文件不能被覆盖，而且会引发错误
		''' </summary>
		''' <param name="sourceFileFullPath">zip文件名</param>
		''' <param name="destinationDirectory">输出的解压文件夹,此文件夹作为当前文件夹</param>
		''' <param name="encoding">在压缩文件中读取项名时使用的编码</param>
		Public Shared Sub UnzipToCurrentDirectory(ByVal sourceFileFullPath As String, ByVal destinationDirectory As String, ByVal encoding As Encoding)
			ThrowIfParamsIllegal(destinationDirectory, sourceFileFullPath)

			Using archive = IO.Compression.ZipFile.Open(sourceFileFullPath, ZipArchiveMode.Read, encoding)
				Dim topDirectory As String
				Dim topDirectoryLength As Integer
				' 如果第一个 entry.Name.Length = 0 说明这个entry是顶层文件夹，往后所有文件都先去掉顶层文件夹，然后再解压
				Dim firstEntry = archive.Entries(0)
				If firstEntry.Name.Length = 0 Then
					topDirectory = firstEntry.FullName
					topDirectoryLength = topDirectory.Length
				End If

				For Each entry In archive.Entries
					Dim fileName = IO.Path.Combine(destinationDirectory, If(topDirectoryLength > 0, entry.FullName.Substring(topDirectoryLength), entry.FullName))
					'  如果文件夹，则需要创建文件夹,不需要写文件
					If entry.Length = 0 AndAlso entry.Name.Length = 0 Then
						If Not EnsureExistsDirectory(fileName) Then
							Exit For
						End If
						Continue For
					End If
					Using steam = entry.Open
						StreamToFile(steam, fileName)
					End Using
				Next
			End Using
		End Sub

		''' <summary>
		''' 文件IO流转换成文件
		''' </summary>
		''' <param name="stream"></param>
		''' <param name="fileName"></param>
		Private Shared Sub StreamToFile(ByVal stream As IO.Stream, ByVal fileName As String)
			Using fileStream As New IO.FileStream(fileName,
								  IO.FileMode.Create,
								  IO.FileAccess.ReadWrite,
								  IO.FileShare.ReadWrite,
								  80 * 1024,
								  True)

				stream.CopyTo(fileStream)
			End Using
		End Sub


		''' <summary>
		''' 提取文件
		''' </summary>
		''' <param name="sourceFileFullPath">zip文件路径</param>
		''' <param name="fileName">要提取的文件名，如果不包含后缀则视为提取文件夹，最后一个‘.’做为后缀分隔符</param>
		''' <param name="destinationDirectory">存储提取文件<paramref name="fileName"/>的目录</param>
		''' <param name="extractPath">是否提取路径，如果为True，则同时创建文件<paramref name="fileName"/>所在的目录（如果有），如果为False,则只提取文件</param>
		''' <param name="encoding"></param>
		Public Shared Sub ExtractFile(ByVal sourceFileFullPath As String, ByVal fileName As String, ByVal destinationDirectory As String, extractPath As Boolean, ByVal encoding As Encoding)
			ThrowIfParamsIllegal(destinationDirectory, sourceFileFullPath)

			Using archive = IO.Compression.ZipFile.Open(sourceFileFullPath, ZipArchiveMode.Read, encoding)
				For Each entry In archive.Entries
					' 命中文件
					If entry.Name = fileName Then
						Dim fileFullPath = IO.Path.Combine(destinationDirectory, If(extractPath, entry.FullName, entry.Name))
						' 确保文件夹存在
						If Not EnsureExistsDirectory(fileFullPath) Then Exit For
						Using steam = entry.Open
							StreamToFile(steam, fileFullPath)
						End Using
						' 提取文件后退出循环
						Exit For
					End If
				Next
			End Using
		End Sub

		''' <summary>
		''' 提取文件夹（包含文件夹内的子文件夹以及文件）
		''' </summary>
		''' <param name="sourceFileFullPath">zip文件路径</param>
		''' <param name="directory">要提取的文件夹名</param>
		''' <param name="destinationDirectory">存储提取文件夹<paramref name="directory"/>的目录</param>
		''' <param name="extractPath">是否提取路径，如果为True，则同时创建文件<paramref name="directory"/>所在的目录（如果有），如果为False,则只提取文件</param>
		''' <param name="encoding"></param>
		Public Shared Sub ExtractDirectory(ByVal sourceFileFullPath As String, ByVal directory As String, ByVal destinationDirectory As String, extractPath As Boolean, ByVal encoding As Encoding)
			ThrowIfParamsIllegal(directory, sourceFileFullPath)

			If Not directory.EndsWith("/") Then
				directory += "/"
			End If
			Using archive = IO.Compression.ZipFile.Open(sourceFileFullPath, ZipArchiveMode.Read, encoding)
				For Each entry In archive.Entries
					' 命中文件夹
					If entry.Name.Length > 0 AndAlso entry.FullName.Contains(directory) Then
						Dim fileFullPath = IO.Path.Combine(destinationDirectory,
													   If(extractPath,
													   entry.FullName,
													   entry.FullName.Substring(entry.FullName.IndexOf(directory))))
						' 确保文件夹存在
						If Not EnsureExistsDirectory(fileFullPath) Then Exit For
						Using steam = entry.Open
							StreamToFile(steam, fileFullPath)
						End Using
					End If
				Next
			End Using
		End Sub

		''' <summary>
		''' 确保文件夹存在
		''' </summary>
		''' <returns></returns>
		Private Shared Function EnsureExistsDirectory(ByVal fileFullPath As String) As Boolean
			If Not IO.Directory.Exists(fileFullPath) Then
				Dim creatSuccess = IO2.Directory.Create(fileFullPath)
				If Not creatSuccess Then
					Debug.Print(Logger.MakeDebugString("创建目录失败"))
					Return False
				End If
			End If

			Return True
		End Function

		Public Class FileInfo
			''' <summary>
			''' 获取在 zip 存档中的项的压缩大小
			''' </summary>
			''' <returns></returns>
			Public Property CompressedLength As Long
			''' <summary>
			''' 获取 zip 存档中的项的相对路径
			''' </summary>
			''' <returns></returns>
			Public Property FullName As String
			''' <summary>
			''' 获取或设置最近一次更改 zip 存档中的项的时间
			''' </summary>
			''' <returns></returns>
			Public Property LastWriteTime As System.DateTimeOffset
			''' <summary>
			''' 获取 zip 存档中的项的未压缩大小
			''' </summary>
			''' <returns></returns>
			Public Property Length As Long
			''' <summary>
			''' 获取在 zip 存档中的项的文件名
			''' </summary>
			''' <returns></returns>
			Public Property FileName As String
		End Class
	End Class

End Namespace
