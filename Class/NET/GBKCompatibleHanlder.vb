Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks

Namespace ShanXingTech.Net2
	Partial Class HttpAsync
		''' <summary>
		''' 兼容GBK编码网页，代替默认的HttpClientHandler
		''' </summary>
		Private Class GBKCompatibleHanlder
			Inherits HttpClientHandler
#Region "字段区"
			Private ReadOnly m_CharSet As String
#End Region

			Public Sub New()
				'
			End Sub

			''' <summary>
			''' 
			''' </summary>
			''' <param name="responseCharSet">解码响应文本使用的字符集，设置错误会导致乱码。设置之前，确保访问的每个网页的字符集都是一样的，否则建议使用无参数的构造函数，程序内部自动检查字符集，当然也会牺牲一点效率。默认使用的字符集是 <see cref="DefaultCharSet"/></param>
			Public Sub New(ByVal responseCharSet As String)
				m_CharSet = responseCharSet
			End Sub

			''' <summary>
			''' 重写SendAsync函数，解决部分中文网页Content.Headers.ContentType没有编码信息而导致乱码问题
			''' </summary>
			''' <param name="request"></param>
			''' <param name="cancellationToken"></param>
			''' <returns></returns>
			Protected Overrides Async Function SendAsync(request As HttpRequestMessage, cancellationToken As CancellationToken) As Task(Of HttpResponseMessage)
				Dim response = Await MyBase.SendAsync(request, cancellationToken)
				Dim contentType = response.Content.Headers.ContentType
				' 如果设置了 allowAutoRedirect =false,有时候返回的是302，没有文本，所以不需要获取文本编码
				If contentType Is Nothing Then Return response

				If String.IsNullOrEmpty(contentType.CharSet) Then
					contentType.CharSet = Await GetCharSetAsync(contentType, response.Content)
				End If

				Return response
			End Function

			''' <summary>
			''' 获取网页的CharSet；因为是从网页源码中获取CharSet，比较低效，所以要求高效的时候，尽量不要使用
			''' </summary>
			''' <param name="contentType"></param>
			''' <returns></returns>
			Private Async Function GetCharSetAsync(ByVal contentType As MediaTypeHeaderValue, ByVal content As HttpContent) As Task(Of String)
				' 获取顺序 调用者设置>meta标签>内部默认（DefaultCharSet）
				If Not m_CharSet.IsNullOrEmpty Then
					Return m_CharSet
				End If

                If content Is Nothing Then Return m_CharSet

                Dim response = Await content.ReadAsStringAsync
                If response.IsNullOrEmpty Then Return m_CharSet

                Dim match = Regex.Match(response, "<meta.*?charset=""?(\w+-?\w+)""?", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
				Dim charSet = match.Groups(1).Value

				' 如果没法从返回到的文本中获取编码方式，那就尝试根据 MediaType 决定 charset 20180627
				If charSet.Length = 0 AndAlso
					(String.Equals("application/json", contentType.MediaType, StringComparison.OrdinalIgnoreCase) OrElse
					String.Equals("text/html", contentType.MediaType, StringComparison.OrdinalIgnoreCase)) Then
					charSet = "utf-8"
				End If

				' 如果获取到的charSet为空的话，那就设置为默认的charSet
				Return If(charSet, DefaultCharSet)
			End Function
		End Class
    End Class
End Namespace
