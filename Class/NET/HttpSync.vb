Imports System.IO
Imports System.IO.Compression
Imports System.Net
Imports System.Net.Security
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Security
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Text.RegularExpressions
Imports ShanXingTech.Text2

Namespace ShanXingTech.Net2
	''' <summary>
	''' 这个人很懒，什么都没写
	''' </summary>
	Public NotInheritable Class HttpSync

		' 明文：神即道, 道法自然, 如来|闪星网络信息科技 ShanXingTech Q2287190283
		' 算法：古典密码中的有密钥换位密码 密钥：ShanXingTech
		Public Const ShanXingTechQ2287190283 = "神闪X7,SQB道信T2道网N9来A2D如H2C然技HA即星I1|N8E法息E8,络G0自科C3"

#Region "枚举区"

#End Region

#Region "常量区"
		''' <summary>
		''' 80K 的buffer
		''' </summary>
		Public Const BufferSize As Integer = 1024 * 80
		Public Const AcceptEncoding As String = "GZIP"
		'访问网络错误提示（断线 超时 url无效等）
		Public Const NetWorkError As String = "网络访问有误"
		' 红米上的UC请求头
		' SetHeaderValue(request.Headers, "User-Agent", "Mozilla/5.0 (Linux; U; Android 5.0.2; zh-CN; Redmi Note 2 Build/LRX22G;) AppleWebKit/537.36 (KHTML,like Gecko) Version/4.0 Chrome/40.0.2214.89 UCBrowser/11.3.0.907 Mobile Safari/537.36 AliApp(TUnionSDK/0.2.8)")
		Public Const GBKContentType As String = "text/html;charset=GBK"
		Public Const UTF8ContentType As String = "text/html;charset=UTF-8"
		Public Const GB2312ContentType As String = "text/html;charset=GB2312"
#End Region

#Region "字段区/共享属性区"
		Private Shared s_DefaultEncoding As Text.Encoding
		Public Shared ReadOnly Property DefaultEncoding() As Text.Encoding
			Get
				' 只在用到的时候才初始化
				If s_DefaultEncoding Is Nothing Then
					s_DefaultEncoding = Text.Encoding.UTF8
				End If

				Return s_DefaultEncoding
			End Get
		End Property

		''' <summary>
		''' 用以实现自动获取编码方式
		''' 存储上一次请求的Host
		''' 如果这次请求的host不同于上一次，则重新获取编码方式
		''' </summary>
		''' <returns></returns>
		Private Shared Property s_PreRequestHost As String

		''' <summary>
		''' 指示类是否初始化
		''' </summary>
		''' <returns></returns>
		Public Shared Property IsInit As Boolean

		Private Shared s_ConnectTimeout As Integer = 5000
		Public Shared Property ConnectTimeout() As Integer
			Get
				Return s_ConnectTimeout
			End Get
			Set(value As Integer)
				s_ConnectTimeout = value
			End Set
		End Property

		Private Shared s_ReadWriteTimeout As Integer = 8000
		Public Shared Property ReadWriteTimeout() As Integer
			Get
				Return s_ReadWriteTimeout
			End Get
			Set(value As Integer)
				s_ReadWriteTimeout = value
			End Set
		End Property

		Private Shared s_IgnoreSSLCheck As Boolean = True
		Public Shared Property IgnoreSSLCheck() As Boolean
			Get
				Return s_IgnoreSSLCheck
			End Get
			Set(value As Boolean)
				s_IgnoreSSLCheck = value
			End Set
		End Property

		Public Shared ReadOnly Property UserAgent() As String
			Get
				Return GetRandUserAgent()
			End Get
		End Property

		Private Shared _userAgentOfPc As String()
		''' <summary>
		''' PC端的UA数组
		''' </summary>
		Private Shared ReadOnly Property UserAgentOfPc() As String()
			Get
				If _userAgentOfPc Is Nothing Then
					_userAgentOfPc = {
						"Mozilla/5.0 (Macintosh; U; Intel Mac OS X 10_6_8; en-us) AppleWebKit/534.50 (KHTML, like Gecko) Version/5.1 Safari/534.50",
						"Mozilla/5.0 (Windows; U; Windows NT 6.1; en-us) AppleWebKit/534.50 (KHTML, like Gecko) Version/5.1 Safari/534.50",
						"Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0;",
						"Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)",
						"Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0)",
						"Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)",
						"Mozilla/5.0 (Macintosh; Intel Mac OS X 10.6; rv:2.0.1) Gecko/20100101 Firefox/4.0.1",
						"Mozilla/5.0 (Windows NT 6.1; rv:2.0.1) Gecko/20100101 Firefox/4.0.1",
						"Opera/9.80 (Macintosh; Intel Mac OS X 10.6.8; U; en) Presto/2.8.131 Version/11.11",
						"Opera/9.80 (Windows NT 6.1; U; en) Presto/2.8.131 Version/11.11",
						"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_0) AppleWebKit/535.11 (KHTML, like Gecko) Chrome/17.0.963.56 Safari/535.11",
						"Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; Maxthon 2.0)",
						"Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; TencentTraveler 4.0)",
						"Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1)",
						"Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; The World)",
						"Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; Trident/4.0; SE 2.X MetaSr 1.0; SE 2.X MetaSr 1.0; .NET CLR 2.0.50727; SE 2.X MetaSr 1.0)",
						"Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; 360SE)",
						"Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1; Avant Browser)",
						"Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 5.1)",
						"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.116 UBrowser/5.6.11466.7 Safari/537.36",
						"Mozilla/5.0 (Windows NT 6.1; WOW64; rv:54.0) Gecko/20100101 Firefox/54.0"
					}
				End If

				Return _userAgentOfPc
			End Get
		End Property

		Private Shared _rand As Random
		Private Shared ReadOnly Property rand() As Random
			Get
				If _rand Is Nothing Then
					_rand = New Random
				End If
				Return _rand
			End Get
		End Property

		Public Shared ReadOnly DefualtHttpHeadersParam As Dictionary(Of HttpRequestHeader, String)
#End Region

#Region "类型构造函数"
		''' <summary>
		''' 类构造函数
		''' 类之内的任意一个静态方法第一次调用时调用此构造函数
		''' 而且程序生命周期内仅调用一次
		''' </summary>
		Shared Sub New()
			DefualtHttpHeadersParam = New Dictionary(Of HttpRequestHeader, String) From {
				{HttpRequestHeader.AcceptEncoding, "gzip"},
				{HttpRequestHeader.UserAgent, UserAgent},
				{HttpRequestHeader.CacheControl, "no-cache"}
			}

			Call Init()
			Call SetAllowUnsafeHeaderParsing20(True)
		End Sub
#End Region

#Region "实例构造函数"
		Public Sub New()
			'
		End Sub
#End Region

		''' <summary>
		''' 设置WebRequest类一些全局的属性，可以加速请求
		''' 这个方法一般调用一次即可
		''' </summary>
		Private Shared Sub Init()
			' 对ServicePointManager的设置将会影响到webrequest从而影响httpwebrequest
			' 这个值最好不要超过1024
			ServicePointManager.DefaultConnectionLimit = 48
			' 去掉“Expect: 100-Continue”请求头，不然会引起post（417） expectation failed
			ServicePointManager.Expect100Continue = False
			ServicePointManager.DnsRefreshTimeout = s_ConnectTimeout
			ServicePointManager.UseNagleAlgorithm = True
			' HttpWebRequest 的请求因为网络问题导致连接没有被释放则会占用连接池中的连接个数，导致并发连接数量减少
			ServicePointManager.SetTcpKeepAlive(True, 1000 * 30, 2)
			' 不使用缓存 使用缓存可能会得到错误的结果
			WebRequest.DefaultCachePolicy = New Cache.RequestCachePolicy(Cache.RequestCacheLevel.NoCacheNoStore)
			' 解决第一次请求时很慢的问题 不使用代理
			WebRequest.DefaultWebProxy = Nothing

			IsInit = True
		End Sub

		''' <summary>
		''' 
		''' </summary>
		''' <param name="useUnsafe"></param>
		''' <returns></returns>
		Private Shared Function SetAllowUnsafeHeaderParsing20(useUnsafe As Boolean) As Boolean
			'Get the assembly that contains the internal class
			Dim aNetAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetAssembly(GetType(System.Net.Configuration.SettingsSection))
			If aNetAssembly Is Nothing Then Return False
			'Use the assembly in order to get the internal type for the internal class
			Dim aSettingsType As Type = aNetAssembly.[GetType]("System.Net.Configuration.SettingsSectionInternal")
			If aSettingsType Is Nothing Then Return False
			'Use the internal static property to get an instance of the internal settings class.
			'If the static instance isn't created allready the property will create it for us.
			Dim anInstance As Object = aSettingsType.InvokeMember("Section", System.Reflection.BindingFlags.[Static] Or System.Reflection.BindingFlags.GetProperty Or System.Reflection.BindingFlags.NonPublic, Nothing, Nothing, Array.Empty(Of Object)())

			If anInstance Is Nothing Then Return False
			'Locate the private bool field that tells the framework is unsafe header parsing should be allowed or not
			Dim aUseUnsafeHeaderParsing As System.Reflection.FieldInfo = aSettingsType.GetField("useUnsafeHeaderParsing", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
			If aUseUnsafeHeaderParsing Is Nothing Then Return False
			aUseUnsafeHeaderParsing.SetValue(anInstance, useUnsafe)
			Return True
		End Function


		''' <summary>
		''' 发送请求
		''' </summary>
		''' <param name="url">请求链接</param>
		''' <param name="method">请求方法</param>
		''' <param name="httpHeadersParam">需要设置的请求头集合</param>
		''' <param name="cookies">传入传出的Cookies</param>
		''' <param name="setSomeRequestHeader">如果设置为True,则自动设置Accept-Encoding、Cache-Control、User-Agent三个请求头为工具内部默认值</param>
		''' <returns></returns>
		Private Shared Function GetHttpWebRequest(url As String, method As HttpMethod, ByRef cookies As CookieContainer, httpHeadersParam As Dictionary(Of HttpRequestHeader, String), ByVal setSomeRequestHeader As Boolean) As HttpWebRequest

			Dim request As HttpWebRequest = Nothing
			Try
				If url.StartsWith("https".TryGetIntern, StringComparison.OrdinalIgnoreCase) Then
					If s_IgnoreSSLCheck Then
						'ServicePointManager.ServerCertificateValidationCallback = AddressOf TrustAllValidationCallback
						ServicePointManager.ServerCertificateValidationCallback = Function()
																					  Return True
																				  End Function
					End If
					request = DirectCast(WebRequest.CreateDefault(New Uri(url)), HttpWebRequest)
					request.ProtocolVersion = HttpVersion.Version11
				Else
					request = DirectCast(WebRequest.Create(url), HttpWebRequest)
				End If

#Region "设置cookie"
				If cookies Is Nothing Then
					' 如果传入的cookie是Nothing并且httpHeadersParam中包含cookie头，说明cookie通过httpHeadersParam来设置
					' 通过httpHeadersParam来设置cookie可以无视domain
					' httpHeadersParam中包含cookie头 将会在下面的 设置请求头 代码块设置到类库中
					If Not httpHeadersParam?.ContainsKey(HttpRequestHeader.Cookie) Then
						request.CookieContainer = New CookieContainer()
						cookies = request.CookieContainer
					End If
				Else
					request.CookieContainer = cookies
				End If
#End Region

#Region "设置请求头"
				SetRequestHeader(request, httpHeadersParam, method, setSomeRequestHeader)
#End Region

			Catch ex As ArgumentNullException
				Logger.WriteLine(ex)
			Catch ex As UriFormatException
				Logger.WriteLine(ex)
			Catch ex As SecurityException
				Logger.WriteLine(ex)
			Catch ex As InvalidCastException
				Logger.WriteLine(ex)
			Catch ex As Exception
				Logger.Debug(ex)
			End Try

			Return request
		End Function

		Private Shared Sub SetRequestHeader(ByRef request As HttpWebRequest, ByRef httpHeadersParam As Dictionary(Of HttpRequestHeader, String), ByVal method As HttpMethod, ByVal setSomeRequestHeader As Boolean)
			' 根据用户传入的请求头设置httpWebRequest对象
			If httpHeadersParam?.Count > 0 Then
				Dim tempDic = CloneDic(httpHeadersParam)

				For Each key In tempDic.Keys
					' User-Agent、Content-Type、Referer 必须使用介种方式添加 不能使用默认的属性去设置
					Select Case key
						Case HttpRequestHeader.Accept
							request.Accept = tempDic(key)
						Case HttpRequestHeader.Connection
							'request.Headers.Set(key, s_DefualttempDic(key))
							'request.Connection = s_DefualttempDic(key)
							NetHelper.SetHeaderValue(request.Headers, "Connection", tempDic(key))
						Case HttpRequestHeader.Expect
							request.Expect = tempDic(key)
						Case HttpRequestHeader.Date
							request.Date = CDate(tempDic(key))
						Case HttpRequestHeader.Host
							request.Host = tempDic(key)
						Case HttpRequestHeader.Range
							request.AddRange(CLng(tempDic(key)))
						Case HttpRequestHeader.Referer
							request.Referer = tempDic(key)
						Case HttpRequestHeader.ContentType
							request.ContentType = tempDic(key)
						Case HttpRequestHeader.IfModifiedSince
							request.IfModifiedSince = CDate(tempDic(key))
						Case HttpRequestHeader.TransferEncoding
							request.TransferEncoding = tempDic(key)
						Case HttpRequestHeader.UserAgent
							request.UserAgent = tempDic(key)
						Case Else
							request.Headers.Set(key, tempDic(key))
					End Select
				Next
			End If

			request.Method = method.ToString()
			request.Timeout = s_ConnectTimeout
			request.ReadWriteTimeout = s_ReadWriteTimeout

			' 如果用户没有设置请求头，并且允许使用一些默认请求头，则自动设置
			If setSomeRequestHeader Then
				If String.IsNullOrEmpty(request.Headers.Get("Accept-Encoding")) Then
					request.Headers.Set(HttpRequestHeader.AcceptEncoding, AcceptEncoding)
				End If
				If String.IsNullOrEmpty(request.Headers.Get("Cache-Control")) Then
					request.Headers.Set(HttpRequestHeader.CacheControl, "no-cache".TryGetIntern)
				End If
				If String.IsNullOrEmpty(request.Headers.Get("User-Agent")) Then
					request.UserAgent = UserAgent
				End If
			End If
		End Sub


		Private Shared Function CloneDic(ByVal dic As Dictionary(Of HttpRequestHeader, String)) As Dictionary(Of HttpRequestHeader, String)
			Dim obj As New Object
			Dim newDic As Dictionary(Of HttpRequestHeader, String)
			SyncLock obj
				Dim memoryStream As New MemoryStream()
				Dim formatter As New BinaryFormatter()
				formatter.Serialize(memoryStream, dic)
				memoryStream.Position = 0
				newDic = DirectCast(formatter.Deserialize(memoryStream), Dictionary(Of HttpRequestHeader, String))
			End SyncLock

			Return newDic
		End Function

		''' <summary>
		''' 发送请求
		''' </summary>
		''' <param name="url">请求链接</param>
		''' <param name="method">请求方法</param>
		''' <param name="userAgent">需要模拟的浏览器UA</param>
		''' <param name="cookies">传入传出的Cookies</param>
		''' <returns></returns>
		Private Shared Function GetHttpWebRequest(ByVal url As String, ByVal method As HttpMethod, ByVal userAgent As String, ByRef cookies As CookieContainer) As HttpWebRequest
			' 核心代码来自于 topsdk Top.Api.Reader.WebUtils
			Dim request As HttpWebRequest
			If userAgent IsNot Nothing Then
				Dim tempHttpHeaders = New Dictionary(Of HttpRequestHeader, String)(DefualtHttpHeadersParam)
				tempHttpHeaders(HttpRequestHeader.UserAgent) = userAgent
				request = GetHttpWebRequest(url:=url, method:=method, cookies:=cookies, httpHeadersParam:=tempHttpHeaders, setSomeRequestHeader:=False)
			Else
				request = GetHttpWebRequest(url:=url, method:=method, cookies:=cookies, httpHeadersParam:=DefualtHttpHeadersParam, setSomeRequestHeader:=False)
			End If

			Return request
		End Function

		''' <summary>
		''' 发送请求
		''' </summary>
		''' <param name="url">请求链接</param>
		''' <param name="method">请求方法</param>
		''' <param name="httpHeadersParam">需要设置的请求头集合</param>
		''' <param name="cookies">传入传出的Cookies</param>
		''' <returns></returns>
		Private Shared Function GetHttpWebRequest(url As String, method As HttpMethod, ByRef cookies As CookieContainer, httpHeadersParam As Dictionary(Of HttpRequestHeader, String)) As HttpWebRequest

			Dim request As HttpWebRequest = GetHttpWebRequest(url:=url, method:=method, cookies:=cookies, httpHeadersParam:=httpHeadersParam, setSomeRequestHeader:=False)

			Return request
		End Function

		''' <summary>
		''' 发送请求
		''' </summary>
		''' <param name="url">请求链接</param>
		''' <param name="method">请求方法</param>
		''' <param name="httpHeadersParam">需要设置的请求头集合</param>
		''' <returns></returns>
		Private Shared Function GetHttpWebRequest(url As String, method As HttpMethod, httpHeadersParam As Dictionary(Of HttpRequestHeader, String)) As HttpWebRequest

			' 向指定地址发送请求
			Dim request As HttpWebRequest = GetHttpWebRequest(url:=url, method:=method, cookies:=Nothing, httpHeadersParam:=httpHeadersParam, setSomeRequestHeader:=False)

			Return request
		End Function

		''' <summary>
		''' 从响应中获取网页源码以及返回cookies
		''' </summary>
		''' <param name="request"></param>
		''' <param name="encoding"></param>
		''' <param name="cookies"></param>
		''' <returns></returns>
		Private Shared Function GetResponseAsString(ByVal request As HttpWebRequest, ByVal encoding As Text.Encoding, ByRef cookies As CookieContainer) As String
			Dim responseContent As String
			Dim response As HttpWebResponse = Nothing
			Dim stream As Stream = Nothing

			Try
				response = DirectCast(request.GetResponse, HttpWebResponse)
				If response Is Nothing Then Return responseContent

				stream = response.GetResponseStream()
				' 如果流需要解压缩则先解压缩
				If response.Headers.ToString.IndexOf(AcceptEncoding, StringComparison.OrdinalIgnoreCase) > -1 Then
					' 不需要写两个Using块，stream的内存会在Using reader块结束之后自动被释放
					' vs的代码分析 和clr via C# 都有提到介个问题
					stream = New GZipStream(stream, CompressionMode.Decompress)
				End If
				responseContent = StringBuilderCache.GetStringAndReleaseBuilderSuper(stream.ReadToEndExt(encoding))

				' 获取cookie
				cookies?.Add(response.Cookies)
			Catch ex As Net.WebException
				Logger.Debug(ex)
			Catch ex As Sockets.SocketException
				Logger.Debug(ex)
			Catch ex As IOException
				Logger.Debug(ex)
			Catch ex As Exception
				Logger.WriteLine(ex)
			Finally
				If stream IsNot Nothing Then
					stream.Close()
				End If
				If response IsNot Nothing Then
					response.Close()
				End If
			End Try

			Return responseContent
		End Function


		''' <summary>
		''' 从响应中获取网页源码,不处理cookies,尝试自动识别编码,无法识别的话使用UTF-8作为默认编码
		''' </summary>
		''' <param name="request"></param>
		''' <returns></returns>
		Private Shared Function GetResponseAsString(ByVal request As HttpWebRequest) As String
			Return GetResponseAsString(request, cookies:=Nothing)
		End Function

		''' <summary>
		''' 从响应中获取网页源码,不处理cookies,尝试自动识别编码,无法识别的话使用UTF-8作为默认编码
		''' </summary>
		''' <param name="request"></param>
		''' <returns></returns>
		Private Shared Function GetResponseAsString(ByVal request As HttpWebRequest, ByRef cookies As CookieContainer) As String
			Dim responseContent As String
			Dim response As HttpWebResponse = Nothing
			Dim stream As Stream = Nothing

			Try
				response = DirectCast(request.GetResponse, HttpWebResponse)
				If response Is Nothing Then Return responseContent

				stream = response.GetResponseStream()
				' 如果流需要解压缩则先解压缩
				If response.Headers.ToString.IndexOf(AcceptEncoding, StringComparison.OrdinalIgnoreCase) > -1 Then
					' 不需要写两个Using块，stream的内存会在Using reader块结束之后自动被释放
					' vs的代码分析 和clr via C# 都有提到介个问题
					stream = New GZipStream(stream, CompressionMode.Decompress)
				End If

				' 从返回的响应头中动态获取编码来解码
				' 没有则启用默认编码
				Dim encoding As Encoding
				Dim contentType = response.ContentType
				Dim pattern = "charset=([\w-]+)".TryGetIntern
				Dim encodingName = Regex.Match(contentType, pattern, RegexOptions.IgnoreCase Or RegexOptions.Compiled).Groups(1).Value
				encoding = If(encodingName.Length > 0,
					Text.Encoding.GetEncoding(encodingName),
					Text.Encoding.UTF8)

				responseContent = StringBuilderCache.GetStringAndReleaseBuilderSuper(stream.ReadToEndExt(encoding))

				' 获取cookie
				cookies?.Add(response.Cookies)
			Catch ex As Net.WebException
				Logger.Debug(ex)
			Catch ex As Sockets.SocketException
				Logger.Debug(ex)
			Catch ex As IOException
				Logger.Debug(ex)
			Catch ex As Exception
				Logger.WriteLine(ex)
			Finally
				If stream IsNot Nothing Then
					stream.Close()
				End If
				If response IsNot Nothing Then
					response.Close()
				End If
			End Try

			Return responseContent
		End Function


		''' <summary>
		''' 从响应中获取网页源码不处理cookies
		''' </summary>
		''' <param name="request"></param>
		''' <param name="encoding"></param>
		''' <returns></returns>
		Private Shared Function GetResponseAsString(ByVal request As HttpWebRequest, ByVal encoding As Text.Encoding) As String
			Return GetResponseAsString(request, encoding, Nothing)
		End Function

		Private Shared Function TrustAllValidationCallback(sender As Object, certificate As X509Certificate, chain As X509Chain, errors As SslPolicyErrors) As Boolean
			' 核心代码来自于 topsdk Top.Api.Reader
			Return True
		End Function

		''' <summary>
		''' 
		''' </summary>
		''' <param name="url"></param>
		''' <param name="postData">可为为Nothing</param>
		''' <param name="cookies"></param>
		''' <param name="encoding">必须指定具体Encoding，不能为Nothing</param>
		''' <param name="httpHeadersParam"></param>
		''' <returns></returns>
		Public Shared Function Post(ByVal url As String, ByVal postData As String, ByRef cookies As CookieContainer, ByVal encoding As Text.Encoding, ByVal httpHeadersParam As Dictionary(Of HttpRequestHeader, String)) As String
			Dim webRequest As HttpWebRequest = GetHttpWebRequest(url:=url, method:=HttpMethod.POST, cookies:=cookies, httpHeadersParam:=httpHeadersParam, setSomeRequestHeader:=False)

			If Not String.IsNullOrEmpty(postData) Then
				Dim body As Byte() = {}
				If postData.Length > 0 Then
					body = encoding.GetBytes(postData)
					webRequest.ContentLength = body.Length
				End If

				If body.Length > 0 Then
					Try
						Using requestStream As Stream = webRequest.GetRequestStream()
							requestStream.Write(body, 0, body.Length)
						End Using
					Catch ex As Exception
						Logger.WriteLine(ex)

						Return NetWorkError
					End Try
				End If
			End If

			Dim responseContent = GetResponseAsString(webRequest, encoding, cookies)

			Return responseContent
		End Function

		''' <summary>
		''' 发送POST请求(最多尝试三次)
		''' </summary>
		''' <param name="url"></param>
		''' <param name="postData"></param>
		''' <param name="encoding"></param>
		''' <param name="cookies"></param>
		''' <param name="httpHeadersParam"></param>
		''' <returns></returns>
		Public Shared Function PostThreeTimeIfError(ByVal url As String, ByVal postData As String, ByRef cookies As CookieContainer, ByVal encoding As Text.Encoding, ByVal httpHeadersParam As Dictionary(Of HttpRequestHeader, String)) As String

			Dim responseContent As String

			' 最少执行一次，最多执行三次 一定程度上确保操作能成功
			Dim getTime As Integer
			Do
				responseContent = Post(url, postData, cookies, encoding, httpHeadersParam)

				getTime += 1

				' 没有联网并且获取网络信息错误的时候 马上退出 
				' 不再浪费时间尝试
				' 非 获取网络信息错误的时候也马上退出，不再进行第二次尝试
				If (responseContent Is NetWorkError.TryGetIntern AndAlso Not NetHelper.IsConnectedToInternet()) OrElse
				Not responseContent Is NetWorkError.TryGetIntern Then
					Exit Do
				ElseIf (responseContent Is NetWorkError.TryGetIntern AndAlso NetHelper.IsConnectedToInternet()) OrElse responseContent.Length = 0 Then
					Continue Do
				End If
			Loop While responseContent Is NetWorkError.TryGetIntern AndAlso getTime < 3

			Return responseContent
		End Function

		Public Shared Function UpLoadFile1(ByVal url As String, ByVal fileName As String, ByRef cookies As CookieContainer, ByVal encoding As Text.Encoding, ByVal httpHeadersParam As Dictionary(Of HttpRequestHeader, String)) As String
			' 向指定地址发送请求
			Dim request As HttpWebRequest = GetHttpWebRequest(url:=url, method:=HttpMethod.OPTIONS, cookies:=cookies, httpHeadersParam:=httpHeadersParam, setSomeRequestHeader:=False)

			request.Headers.Set("Access-Control-Request-Method", HttpMethod.POST.ToString)
			request.Headers.Set("Origin", "https://pan.baidu.com")

			Dim responseContent = GetResponseAsString(request, encoding)

			Return responseContent
		End Function


		Public Shared Function UpLoadFile(ByVal url As String, ByVal fileName As String, ByRef cookies As CookieContainer, ByVal encoding As Text.Encoding, ByVal httpHeadersParam As Dictionary(Of HttpRequestHeader, String)) As (Success As Boolean, Message As String)
			Dim request As HttpWebRequest

			'请求响应类
			Dim response As HttpWebResponse = Nothing
			'文件流
			Dim file_stream As FileStream = Nothing
			'响应结果读取类
			Dim reader As StreamReader = Nothing

			If url.StartsWith("https".TryGetIntern, StringComparison.OrdinalIgnoreCase) Then
				If s_IgnoreSSLCheck Then
					'ServicePointManager.ServerCertificateValidationCallback = AddressOf TrustAllValidationCallback
					ServicePointManager.ServerCertificateValidationCallback = Function()
																				  Return True
																			  End Function
				End If
				request = DirectCast(WebRequest.CreateDefault(New Uri(url)), HttpWebRequest)
				request.ProtocolVersion = HttpVersion.Version11
			Else
				request = DirectCast(WebRequest.Create(url), HttpWebRequest)
			End If

#Region "设置请求头"
			request.Method = HttpMethod.POST.ToString

#Region "设置cookie"
			If cookies Is Nothing Then
				' 如果传入的cookie是Nothing并且httpHeadersParam中包含cookie头，说明cookie通过httpHeadersParam来设置
				' 通过httpHeadersParam来设置cookie可以无视domain
				If Not httpHeadersParam.ContainsKey(HttpRequestHeader.Cookie) Then
					request.CookieContainer = New CookieContainer()
					cookies = request.CookieContainer
				End If
			Else
				request.CookieContainer = cookies
			End If
#End Region

			' 根据用户传入的请求头设置httpWebRequest对象
			If httpHeadersParam IsNot Nothing AndAlso httpHeadersParam.Count > 0 Then
				For Each key In httpHeadersParam.Keys
					' User-Agent、Content-Type、Referer 必须使用介种方式添加 不能使用默认的属性去设置
					Select Case key
						Case HttpRequestHeader.Accept
							request.Accept = httpHeadersParam(key)
						Case HttpRequestHeader.Connection
							'request.Headers.Set(key, s_DefualtHttpHeadersParam(key))
							'request.Connection = s_DefualtHttpHeadersParam(key)
							NetHelper.SetHeaderValue(request.Headers, "Connection", httpHeadersParam(key))
						Case HttpRequestHeader.Expect
							request.Expect = httpHeadersParam(key)
						Case HttpRequestHeader.Date
							request.Date = CDate(httpHeadersParam(key))
						Case HttpRequestHeader.Host
							request.Host = httpHeadersParam(key)
						Case HttpRequestHeader.Range
							request.AddRange(CLng(httpHeadersParam(key)))
						Case HttpRequestHeader.Referer
							request.Referer = httpHeadersParam(key)
						Case HttpRequestHeader.ContentType
							request.ContentType = httpHeadersParam(key)
						Case HttpRequestHeader.IfModifiedSince
							request.IfModifiedSince = CDate(httpHeadersParam(key))
						Case HttpRequestHeader.TransferEncoding
							request.TransferEncoding = httpHeadersParam(key)
						Case HttpRequestHeader.UserAgent
							request.UserAgent = httpHeadersParam(key)
						Case Else
							request.Headers.Set(key, httpHeadersParam(key))
					End Select
				Next
			End If
#End Region
			request.ContentType = "multipart/form-data; boundary=---------------------------101843053213124"

			'包体填充类
			'设置请求体
			Dim mem_stream As MemoryStream = New MemoryStream()

			'边界符
			Dim boundary As String = "-----------------------------12046247044215"
			Dim begin_boundary As Byte() = Encoding.UTF8.GetBytes(boundary & vbCrLf)
			'开头
			mem_stream.Write(begin_boundary, 0, begin_boundary.Length)

			Dim file_string As String = $"Content-Disposition: form-data; name=""file""; filename=""blob""{Environment.NewLine}Content-Type: application/octet-stream{Environment.NewLine}{Environment.NewLine}"
			Dim file_byte As Byte() = Encoding.UTF8.GetBytes(file_string)
			mem_stream.Write(file_byte, 0, file_byte.Length)

			' 文件格式
			Dim fileFilter = Encoding.UTF8.GetBytes(ChrW(137) & "PNG" & vbCrLf & ChrW(26))
			mem_stream.Write(fileFilter, 0, fileFilter.Length)

			Dim buffer As Byte() = New Byte(4096) {}
			Dim bytes As Integer
			'文件内容
			file_stream = New FileStream(fileName, FileMode.Open, FileAccess.Read)

			bytes = file_stream.Read(buffer, 0, buffer.Length)
			While bytes <> 0
				mem_stream.Write(buffer, 0, bytes)
				bytes = file_stream.Read(buffer, 0, buffer.Length)
			End While

			'结尾
			Dim end_boundary As Byte() = Encoding.UTF8.GetBytes(boundary & "--" & vbCrLf)
			mem_stream.Write(end_boundary, 0, end_boundary.Length)

			request.ContentLength = mem_stream.Length
			Dim req_stream As Stream = request.GetRequestStream()
			mem_stream.Position = 0
			Dim temp_buffer As Byte() = New Byte(CInt(mem_stream.Length) - 1) {}
			mem_stream.Read(temp_buffer, 0, temp_buffer.Length)
			req_stream.Write(temp_buffer, 0, temp_buffer.Length)

			Dim responseContent = GetResponseAsString(request, encoding, cookies)

			Return (True, responseContent)
		End Function


#Region "获取网页编码"
		''' <summary>
		''' 从响应头中获取网页编码 
		''' 储存到模块级的encoding成员中，其他不传入encoding参数的函数都是直接调用这个值
		''' 没有的话返回Text.Encoding.Default
		''' </summary>
		''' <param name="url"></param>
		''' <param name="userAgent"></param>
		''' <param name="cookies"></param>
		Public Shared Sub GetEncoding（ByVal url As String, ByVal userAgent As String, ByRef cookies As CookieContainer）
			' 向指定地址发送请求
			Dim request As HttpWebRequest = GetHttpWebRequest(url:=url, method:=HttpMethod.HEAD, userAgent:=userAgent, cookies:=cookies)

			Try
#Region "从数据流中异步获取网页源码"
				Using response As HttpWebResponse = DirectCast(request.GetResponse(), HttpWebResponse)
					If response Is Nothing Then Exit Try

					' 从返回的响应头中动态获取编码来解码
					' 没有则启用默认编码
					' 确保返回头里面有host 或者大于0
					Dim headers = request.Headers.ToString
					If request.Headers.Count = 0 OrElse
					Not headers.IndexOf("host".TryGetIntern, StringComparison.OrdinalIgnoreCase) > -1 OrElse
					request.Headers.Get("Host".TryGetIntern).Equals(s_PreRequestHost, StringComparison.OrdinalIgnoreCase) Then
						Exit Try
					End If
					s_PreRequestHost = request.Headers.Get("Host".TryGetIntern)

					' 从返回的响应头中动态获取编码来解码
					' 没有则启用默认编码
					headers = response.ContentType
					Dim pattern = "charset=([\w-]+)".TryGetIntern
					Dim encodingName = Regex.Match(headers, pattern, RegexOptions.IgnoreCase Or RegexOptions.Compiled).Groups(1).Value
					s_DefaultEncoding = If(encodingName.Length > 0,
						Text.Encoding.GetEncoding(encodingName),
						Text.Encoding.UTF8)

					' 获取cookie
					If cookies IsNot Nothing AndAlso response.Cookies.Count > 0 Then
						cookies.Add(response.Cookies)
					End If
				End Using
#End Region
			Catch ex As Exception
				Logger.Debug(ex)
			End Try
		End Sub

		''' <summary>
		''' 从响应头中获取网页编码 
		''' 储存到模块级的encoding成员中，其他不传入encoding参数的函数都是直接调用这个值
		''' 没有的话返回Text.Encoding.Default
		''' </summary>
		''' <param name="url"></param>
		''' <param name="userAgent"></param>
		''' <param name="cookies"></param>
		Public Shared Function GetEncodingAsync（ByVal url As String, ByVal userAgent As String, ByRef cookies As CookieContainer） As IAsyncResult
			Dim asyncResult As IAsyncResult = Nothing

			' 向指定地址发送请求
			Dim request As HttpWebRequest = GetHttpWebRequest(url:=url, method:=HttpMethod.HEAD, userAgent:=userAgent, cookies:=cookies)

			Try
#Region "从数据流中异步获取网页源码"
				asyncResult = request.BeginGetResponse(AddressOf GetEncodingResponseCallback, request)
				Debug.Print(Logger.MakeDebugString(String.Concat(vbTab, "等待异步")))
#End Region
			Catch ex As Net.WebException
				Logger.Debug(ex)
			Catch ex As Exception
				Logger.Debug(ex)
			End Try

			Return asyncResult
		End Function

		Private Shared Sub GetEncodingResponseCallback(asynchronousResult As IAsyncResult)
			Try
				Dim request As HttpWebRequest = DirectCast(asynchronousResult.AsyncState, HttpWebRequest)

				If request Is Nothing Then
					s_DefaultEncoding = Text.Encoding.UTF8
					Return
				End If

				Using response As HttpWebResponse = DirectCast(request.EndGetResponse(asynchronousResult), HttpWebResponse)
					If response Is Nothing Then
						s_DefaultEncoding = Text.Encoding.UTF8
						Return
					End If

					' 从返回的响应头中动态获取编码来解码
					' 没有则启用默认编码
					' 确保返回头里面有host 或者大于0
					Dim headers = request.Headers.ToString
					If request.Headers.Count = 0 OrElse
						Not headers.IndexOf("host".TryGetIntern, StringComparison.OrdinalIgnoreCase) > -1 OrElse
						request.Headers.Get("Host".TryGetIntern).Equals(s_PreRequestHost, StringComparison.OrdinalIgnoreCase) Then
						Return
					End If
					s_PreRequestHost = request.Headers.Get("Host".TryGetIntern)

					' 从返回的响应头中动态获取编码来解码
					' 没有则启用默认编码
					headers = response.ContentType
					Dim pattern = "charset=([\w-]+)".TryGetIntern
					Dim encodingName = Regex.Match(headers, pattern, RegexOptions.IgnoreCase Or RegexOptions.Compiled).Groups(1).Value
					s_DefaultEncoding = If(encodingName.Length > 0,
						Text.Encoding.GetEncoding(encodingName),
						Text.Encoding.UTF8)
				End Using

				Debug.Print(Logger.MakeDebugString(String.Concat(vbTab, "异步完毕")))
			Catch ex As Exception
				Logger.Debug(ex)
			End Try
		End Sub

		''' <summary>
		''' 从响应头中获取网页编码 
		''' 没有的话返回Text.Encoding.UTF8
		''' </summary>
		''' <param name="url"></param>
		Public Shared Sub GetEncoding（ByVal url As String）
			GetEncoding(url, UserAgent, Nothing)
		End Sub
#End Region


		''' <summary>
		''' 获取网页源码(最多尝试三次)
		''' </summary>
		''' <param name="url">链接</param>
		''' <param name="encoding">网页编码</param>
		''' <param name="cookies">传入传出的Cookies</param>
		''' <param name="httpHeadersParam">各种请求头，cookie头也可以在这个参数里面设置</param>
		''' <returns></returns>
		Public Shared Function GetHtmlThreeTimeIfError(ByVal url As String, ByVal encoding As Encoding, ByRef cookies As CookieContainer, ByVal httpHeadersParam As Dictionary(Of HttpRequestHeader, String)) As String

			Dim responseContent As String

			' 最少执行一次，最多执行三次 一定程度上确保操作能成功
			Dim getTime As Integer
			Do

#Region "Get操作主体"
				' 向指定地址发送请求
				Dim request As HttpWebRequest = GetHttpWebRequest(url:=url, method:=HttpMethod.GET, cookies:=cookies, httpHeadersParam:=httpHeadersParam, setSomeRequestHeader:=False)

				' 任何时候都返回 响应头中的cookie
				responseContent = GetResponseAsString(request, encoding, cookies)
#End Region

				getTime += 1

				' 没有联网并且获取网络信息错误的时候 马上退出 
				' 不再浪费时间尝试
				' 非 获取网络信息错误的时候也马上退出，不再进行第二次尝试
				If (responseContent Is NetWorkError.TryGetIntern AndAlso Not NetHelper.IsConnectedToInternet()) OrElse
				Not responseContent Is NetWorkError.TryGetIntern Then
					Exit Do
				ElseIf (responseContent Is NetWorkError.TryGetIntern AndAlso NetHelper.IsConnectedToInternet()) OrElse responseContent.Length = 0 Then
					Continue Do
				End If
			Loop While responseContent Is NetWorkError.TryGetIntern AndAlso getTime < 3

			Return responseContent
		End Function

		''' <summary>
		''' 获取网页源码(最多尝试三次)；尝试自动获取网页编码，失败则使用UTF-8编码
		''' </summary>
		''' <param name="url">链接</param>
		''' <param name="cookies">传入传出的Cookies</param>
		''' <param name="httpHeadersParam">各种请求头，cookie头也可以在这个参数里面设置</param>
		''' <param name="setSomeRequestHeader">如果设置为True,则自动设置Accept-Encoding、Cache-Control、User-Agent三个请求头为工具内部默认值</param>
		''' <returns></returns>
		Public Shared Function GetHtmlThreeTimeIfError(ByVal url As String, ByRef cookies As CookieContainer, ByVal httpHeadersParam As Dictionary(Of HttpRequestHeader, String), ByVal setSomeRequestHeader As Boolean) As String
			Dim responseContent As String
			Dim request As HttpWebRequest

			' 最少执行一次，最多执行三次 一定程度上确保操作能成功
			Dim getTime As Integer
			Do

#Region "Get操作主体"
				' 向指定地址发送请求
				request = GetHttpWebRequest(url:=url, method:=HttpMethod.GET, cookies:=cookies, httpHeadersParam:=httpHeadersParam, setSomeRequestHeader:=setSomeRequestHeader)

				responseContent = GetResponseAsString(request, cookies)
#End Region

				getTime += 1

				' 没有联网并且获取网络信息错误的时候 马上退出 
				' 不再浪费时间尝试
				' 非 获取网络信息错误的时候也马上退出，不再进行第二次尝试
				If (responseContent Is NetWorkError.TryGetIntern AndAlso Not NetHelper.IsConnectedToInternet()) OrElse
				Not responseContent Is NetWorkError.TryGetIntern Then
					Exit Do
				ElseIf (responseContent Is NetWorkError.TryGetIntern AndAlso NetHelper.IsConnectedToInternet()) OrElse responseContent.Length = 0 Then
					Continue Do
				End If
			Loop While responseContent Is NetWorkError.TryGetIntern AndAlso getTime < 3

			Return responseContent
		End Function

		''' <summary>
		''' 获取网页源码(最多尝试三次)；尝试自动获取网页编码，失败则使用UTF-8编码
		''' </summary>
		''' <param name="url">链接</param>
		''' <param name="httpHeadersParam">各种请求头，cookie头也可以在这个参数里面设置</param>
		''' <param name="setSomeRequestHeader">如果设置为True,则自动设置Accept-Encoding、Cache-Control、User-Agent三个请求头为工具内部默认值</param>
		''' <returns></returns>
		Public Shared Function GetHtmlThreeTimeIfError(ByVal url As String, ByVal httpHeadersParam As Dictionary(Of HttpRequestHeader, String), ByVal setSomeRequestHeader As Boolean) As String
			Return GetHtmlThreeTimeIfError(url, Nothing, httpHeadersParam, setSomeRequestHeader)
		End Function

		''' <summary>
		''' 获取网页源码(最多尝试三次)；尝试自动获取网页编码，失败则使用UTF-8编码
		''' </summary>
		''' <param name="url">链接</param>
		''' <param name="httpHeadersParam">各种请求头，cookie头也可以在这个参数里面设置</param>
		''' <returns></returns>
		Public Shared Function GetHtmlThreeTimeIfError(ByVal url As String, ByVal httpHeadersParam As Dictionary(Of HttpRequestHeader, String)) As String
			Return GetHtmlThreeTimeIfError(url, Nothing, httpHeadersParam, False)
		End Function

		''' <summary>
		''' 获取网页源码(最多尝试三次)
		''' </summary>
		''' <param name="url">链接</param>
		''' <param name="encoding">网页编码</param>
		''' <param name="httpHeadersParam">各种请求头，cookie头也可以在这个参数里面设置</param>
		''' <returns></returns>
		Public Shared Function GetHtmlThreeTimeIfError(ByVal url As String, ByVal encoding As Encoding, ByVal httpHeadersParam As Dictionary(Of HttpRequestHeader, String)) As String
			Dim responseContent As String

			' 最少执行一次，最多执行三次 一定程度上确保操作能成功
			Dim getTime As Integer
			Do

#Region "Get操作主体"
				' 向指定地址发送请求
				Dim request As HttpWebRequest = GetHttpWebRequest(url:=url, method:=HttpMethod.GET, cookies:=Nothing, httpHeadersParam:=httpHeadersParam, setSomeRequestHeader:=False)

				responseContent = GetResponseAsString(request, encoding, Nothing)
#End Region
				getTime += 1

				' 没有联网并且获取网络信息错误的时候 马上退出 
				' 不再浪费时间尝试
				' 非 获取网络信息错误的时候也马上退出，不再进行第二次尝试
				If (responseContent Is NetWorkError.TryGetIntern AndAlso Not NetHelper.IsConnectedToInternet()) OrElse
				Not responseContent Is NetWorkError.TryGetIntern Then
					Exit Do
				ElseIf (responseContent Is NetWorkError.TryGetIntern AndAlso NetHelper.IsConnectedToInternet()) OrElse responseContent.Length = 0 Then
					Continue Do
				End If
			Loop While responseContent Is NetWorkError.TryGetIntern AndAlso getTime < 3

			Return responseContent
		End Function

		''' <summary>
		''' 获取网页源码(最多尝试三次)
		''' </summary>
		''' <param name="url">链接</param>
		''' <param name="encoding">网页编码</param>
		''' <returns></returns>
		Public Shared Function GetHtmlThreeTimeIfError(ByVal url As String, ByVal encoding As Text.Encoding) As String
			Dim responseContent As String

			' 最少执行一次，最多执行三次 一定程度上确保操作能成功
			Dim getTime As Integer
			Do

#Region "Get操作主体"
				' 向指定地址发送请求
				Dim request As HttpWebRequest = GetHttpWebRequest(url:=url, method:=HttpMethod.GET, userAgent:=UserAgent, cookies:=Nothing)

				responseContent = GetResponseAsString(request, encoding)
#End Region

				getTime += 1

				' 没有联网并且获取网络信息错误的时候 马上退出 
				' 不再浪费时间尝试
				' 非 获取网络信息错误的时候也马上退出，不再进行第二次尝试
				If (responseContent Is NetWorkError.TryGetIntern AndAlso Not NetHelper.IsConnectedToInternet()) OrElse
				Not responseContent Is NetWorkError.TryGetIntern Then
					Exit Do
				ElseIf (responseContent Is NetWorkError.TryGetIntern AndAlso NetHelper.IsConnectedToInternet()) OrElse responseContent.Length = 0 Then
					Continue Do
				End If
			Loop While responseContent Is NetWorkError.TryGetIntern AndAlso getTime < 3

			Return responseContent
		End Function

		''' <summary>
		''' 获取网页源码(最多尝试三次);尝试自动获取网页编码，失败则使用UTF-8编码,使用内部UA
		''' </summary>
		''' <param name="url">链接</param>
		''' <returns></returns>
		Public Shared Function GetHtmlThreeTimeIfError(ByVal url As String) As String
			Dim responseContent As String

			' 最少执行一次，最多执行三次 一定程度上确保操作能成功
			Dim getTime As Integer
			Do

#Region "Get操作主体"
				' 向指定地址发送请求
				Dim request As HttpWebRequest = GetHttpWebRequest(url:=url, method:=HttpMethod.GET, userAgent:=UserAgent, cookies:=Nothing)

				responseContent = GetResponseAsString(request, cookies:=Nothing)
#End Region

				getTime += 1

				' 没有联网并且获取网络信息错误的时候 马上退出 
				' 不再浪费时间尝试
				' 非 获取网络信息错误的时候也马上退出，不再进行第二次尝试
				If (responseContent Is NetWorkError.TryGetIntern AndAlso Not NetHelper.IsConnectedToInternet()) OrElse
				Not responseContent Is NetWorkError.TryGetIntern Then
					Exit Do
				ElseIf (responseContent Is NetWorkError.TryGetIntern AndAlso NetHelper.IsConnectedToInternet()) OrElse responseContent.Length = 0 Then
					Continue Do
				End If
			Loop While responseContent Is NetWorkError.TryGetIntern AndAlso getTime < 3

			Return responseContent
		End Function

		''' <summary>
		''' 通过url获取网页源码(返回值可能为html json xml...)
		''' 不处理Cookie
		''' <para> 使用默认UserAgent</para>
		''' </summary>
		''' <param name="url">请求链接</param>
		''' <param name="encoding">解码方式</param>
		''' <returns></returns>
		Public Shared Function GetHtml(ByVal url As String, ByVal encoding As Text.Encoding) As String
			Dim responseContent As String

			' 向指定地址发送请求
			Dim request As HttpWebRequest = GetHttpWebRequest(url:=url, method:=HttpMethod.GET, userAgent:=UserAgent, cookies:=Nothing)

			responseContent = GetResponseAsString(request, encoding)

			Return responseContent
		End Function

		''' <summary>
		''' 通过url获取网页源码(返回值可能为html json xml...)
		''' <para> 不处理Cookie</para>
		''' <para> 使用默认UserAgent</para>
		''' <para> 注意：使用前请调用GetEncoding方法先获取网页编码</para>
		''' </summary>
		''' <param name="url">请求链接</param>
		''' <returns></returns>
		Public Shared Function GetHtml(ByVal url As String) As String
			Dim responseContent As String

			' 向指定地址发送请求
			Dim request As HttpWebRequest = GetHttpWebRequest(url:=url, method:=HttpMethod.GET, userAgent:=UserAgent, cookies:=Nothing)

			responseContent = GetResponseAsString(request, DefaultEncoding, Nothing)

			Return responseContent
		End Function


		''' <summary>
		''' 通过url获取网页源码(返回值可能为html json xml...)
		''' <para> 处理Cookie</para>
		''' <para> 使用默认UserAgent</para>
		''' <para> 注意：使用前请调用GetEncoding方法先获取网页编码</para>
		''' </summary>
		''' <param name="url">请求链接</param>
		''' <param name="cookies">传入传出的Cookies</param>
		''' <returns></returns>
		Public Shared Function GetHtml(ByVal url As String， ByRef cookies As CookieContainer) As String
			Dim responseContent As String

			' 向指定地址发送请求
			Dim request As HttpWebRequest = GetHttpWebRequest(url:=url, method:=HttpMethod.GET, userAgent:=UserAgent, cookies:=cookies)

			responseContent = GetResponseAsString(request, DefaultEncoding, cookies)

			Return responseContent
		End Function


		''' <summary>
		''' 通过url获取网页源码(返回值可能为html json xml...)
		''' <para> 处理Cookie</para>
		''' <para> 使用默认UserAgent</para>
		''' </summary>
		''' <param name="url">请求链接</param>
		''' <param name="cookies">传入传出的Cookies</param>
		''' <param name="encoding">解码方式</param>
		''' <returns></returns>
		Public Shared Function GetHtml(ByVal url As String， ByRef cookies As CookieContainer, ByVal encoding As Text.Encoding) As String
			' 向指定地址发送请求
			Dim request As HttpWebRequest = GetHttpWebRequest(url:=url, method:=HttpMethod.GET, userAgent:=UserAgent, cookies:=cookies)

			Dim responseContent = GetResponseAsString(request, encoding, cookies)

			Return responseContent
		End Function

		''' <summary>
		''' 从UA数组中随机获取一个ua作为某次请求ua
		''' </summary>
		''' <returns></returns>
		Private Shared Function GetRandUserAgent() As String
			Dim uaIndex = rand.Next(0, UserAgentOfPc.Length)

			Dim ua = UserAgentOfPc(uaIndex)

			Return ua
		End Function
	End Class

End Namespace