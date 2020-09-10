Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Handlers
Imports System.Net.Http.Headers
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks

Imports ShanXingTech
Imports ShanXingTech.Exception2
Imports ShanXingTech.IO2
Imports ShanXingTech.Text2

Namespace ShanXingTech.Net2
    Public Class HttpAsync
        Implements IDisposable

#Region "常量区"
        Public Const DefaulMediaType As String = "application/x-www-form-urlencoded"

#End Region

#Region "事件区"
        Public Event DownloadProgressChanged As EventHandler(Of DownloadProgressChangedEventArgs)
        Public Event DownloadFileCompleted As EventHandler(Of DownloadFileCompletedEventArgs)
        Public Event UploadProgressChanged As EventHandler(Of UploadProgressChangedEventArgs)
        Public Event UploadFileCompleted As EventHandler(Of UploadFileCompletedEventArgs)
#End Region

#Region "实例属性区"
        Public ReadOnly Property Cookies() As CookieContainer
            Get
                Return m_HttpClientHandler.CookieContainer
            End Get
        End Property

        Public ReadOnly Property RequestHeaders() As HttpRequestHeaders
            Get
                Return m_HttpClient.DefaultRequestHeaders
            End Get
        End Property

        ''' <summary>
        ''' 默认的编码字符集为GBK
        ''' </summary>
        ''' <returns></returns>
        Public Property DefaultCharSet As String

        Private m_AllowAutoRedirect As Boolean
        ''' <summary>
        ''' 指示请求是否启用自动重定向
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property AllowAutoRedirect() As Boolean
            Get
                Return m_AllowAutoRedirect
            End Get
        End Property

        Private m_DefaulTimeoutMilliseconds As Integer
        ''' <summary>
        ''' 默认的超时时间
        ''' </summary>
        ''' <returns></returns>
        Public Property DefaulTimeoutMilliseconds As Integer
            Set(value As Integer)
                m_DefaulTimeoutMilliseconds = value

                ' 如果已经实例化，需要同时更新 Timeout 字段以使设置生效
                If m_HttpClient Is Nothing Then Return
                ReInit(Cookies, m_AllowAutoRedirect, Proxy, DefaultCharSet)
            End Set
            Get
                Return m_DefaulTimeoutMilliseconds
            End Get
        End Property

        ''' <summary>
        ''' 获取处理程序使用的代理信息
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Proxy() As IWebProxy
            Get
                Return m_HttpClientHandler.Proxy
            End Get
        End Property

        ''' <summary>
        ''' 获取是否使用的代理信息
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property UseProxy() As Boolean
            Get
                Return m_HttpClientHandler.UseProxy
            End Get
        End Property
#End Region

#Region "字段区"
        '<ThreadStatic>
        'Private Shared m_CancellationTokenSource As CancellationTokenSource
        '<ThreadStatic>
        'Private Shared m_CancellationToken As CancellationToken
        Public Shared ReadOnly Instance As HttpAsync
        Private m_HttpClient As HttpClient
        Private m_HttpClientHandler As HttpClientHandler
        Private m_ProgressMessageHandler As ProgressMessageHandler
#End Region

#Region "构造函数区"

        ''' <summary>
        ''' 类构造函数
        ''' 类之内的任意一个静态方法第一次调用时调用此构造函数
        ''' 而且程序生命周期内仅调用一次
        ''' </summary>
        Shared Sub New()
            Instance = New HttpAsync
        End Sub

        ''' <summary>
        ''' 不传入cookies，自动处理重定向
        ''' </summary>
        Sub New()
            HttpPublicConfig()

            InitInternal(Nothing, True, Nothing, Nothing)
        End Sub

        ''' <summary>
        ''' 传入cookies，自动处理重定向
        ''' </summary>
        ''' <param name="cookies"></param>

        Sub New(ByRef cookies As CookieContainer)
            HttpPublicConfig()

            InitInternal(cookies, True, Nothing, Nothing)
        End Sub

        ''' <summary>
        ''' 不传入cookies，重定向由参数 <paramref name="allowAutoRedirect"/> 决定
        ''' </summary>
        ''' <param name="allowAutoRedirect">指示处理程序是否应跟随重定向响应,默认为True</param>
        Sub New(ByVal allowAutoRedirect As Boolean)
            HttpPublicConfig()

            m_AllowAutoRedirect = allowAutoRedirect
            InitInternal(Nothing, allowAutoRedirect, Nothing, Nothing)
        End Sub

        ''' <summary>
        ''' 不传入cookies，自动重定向，代理信息由参数 <paramref name="proxy"/> 决定
        ''' </summary>
        ''' <param name="proxy">指示处理程序是否应设置代理信息,默认为Nothing</param>
        Sub New(ByRef proxy As IWebProxy)
            HttpPublicConfig()

            InitInternal(Nothing, True, proxy, Nothing)
        End Sub

        ''' <summary>
        ''' 传入cookies，重定向由参数 <paramref name="allowAutoRedirect"/> 决定
        ''' </summary>
        ''' <param name="cookies"></param>
        ''' <param name="allowAutoRedirect">指示处理程序是否应跟随重定向响应,默认为True</param>
        Sub New(ByRef cookies As CookieContainer, ByVal allowAutoRedirect As Boolean)
            HttpPublicConfig()

            m_AllowAutoRedirect = allowAutoRedirect
            InitInternal(cookies, allowAutoRedirect, Nothing, Nothing)
        End Sub

        ''' <summary>
        ''' 传入cookies，重定向由参数 <paramref name="allowAutoRedirect"/> 决定，代理信息由参数 <paramref name="proxy"/> 决定
        ''' </summary>
        ''' <param name="cookies"></param>
        ''' <param name="allowAutoRedirect">指示处理程序是否应跟随重定向响应,默认为True</param>
        ''' <param name="proxy">指示处理程序是否应设置代理信息,默认为Nothing</param>
        Sub New(ByRef cookies As CookieContainer, ByVal allowAutoRedirect As Boolean, ByRef proxy As IWebProxy)
            HttpPublicConfig()

            m_AllowAutoRedirect = allowAutoRedirect
            InitInternal(cookies, allowAutoRedirect, proxy, Nothing)
        End Sub

        ''' <summary>
        ''' 传入cookies，重定向由参数 <paramref name="allowAutoRedirect"/> 决定，代理信息由参数 <paramref name="proxy"/> 决定
        ''' </summary>
        ''' <param name="cookies"></param>
        ''' <param name="allowAutoRedirect">指示处理程序是否应跟随重定向响应,默认为True</param>
        ''' <param name="proxy">指示处理程序是否应设置代理信息,默认为Nothing</param>
        ''' <param name="charSet">指示应对返回文本采用何种编码</param>
        Sub New(ByRef cookies As CookieContainer, ByVal allowAutoRedirect As Boolean, ByRef proxy As IWebProxy, ByVal charSet As String)
            HttpPublicConfig()

            m_AllowAutoRedirect = allowAutoRedirect
            InitInternal(cookies, allowAutoRedirect, proxy, charSet)
        End Sub
#End Region

#Region "非公开的内部实现函数，公共方法等"
        ''' <summary>
        ''' 用于重新初始化类，比如更换账号等涉及到重新传入cookie的情况
        ''' </summary>
        ''' <param name="cookies">如果为Nothing，则只重新初始化类，不设置cookie</param>
        ''' <param name="allowAutoRedirect">指示处理程序是否应跟随重定向响应,默认为True</param>
        ''' <param name="proxy">指示处理程序是否应设置代理信息,默认为Nothing</param>
        ''' <param name="charSet">解码响应文本使用的字符集，设置错误会导致乱码。设置之前，确保访问的每个网页的字符集都是一样的，否则建议使用无参数的构造函数，程序内部自动检查字符集，当然也会牺牲一点效率。</param>
        Private Sub InitInternal(ByVal cookies As CookieContainer, ByVal allowAutoRedirect As Boolean, ByRef proxy As IWebProxy, ByVal charSet As String)
            DefaultCharSet = "GBK"
            m_DefaulTimeoutMilliseconds = 10000

            m_HttpClientHandler = New GBKCompatibleHanlder(charSet) With {
                .UseProxy = proxy IsNot Nothing,
                .Proxy = proxy,
                .AutomaticDecompression = DecompressionMethods.GZip Or DecompressionMethods.Deflate,
                .AllowAutoRedirect = allowAutoRedirect
            }
            ' 把传入的cookie装盒并且设置到请求头
            ' cookie只能是这样设置.而且httpClient内部会自动管理cookies
            ' 不能多次设置，只能是设置一次
            If cookies IsNot Nothing Then
                m_HttpClientHandler.CookieContainer = cookies
            End If

            m_ProgressMessageHandler = New ProgressMessageHandler(m_HttpClientHandler)
            m_HttpClient = New HttpClient(m_ProgressMessageHandler, False) With {
                .Timeout = TimeSpan.FromMilliseconds(m_DefaulTimeoutMilliseconds)
            }
            m_HttpClient.DefaultRequestHeaders.ExpectContinue = False
        End Sub

        ''' <summary>
        ''' 用于重新初始化类，比如更换账号等涉及到重新传入cookie的情况
        ''' </summary>
        ''' <param name="cookies">如果为Nothing，则只重新初始化类，不设置cookie</param>
        Public Sub ReInit(ByRef cookies As CookieContainer)
            ReInit(cookies, True)
        End Sub

        ''' <summary>
        ''' 用于重新初始化类，比如更换账号等涉及到重新传入cookie的情况
        ''' </summary>
        ''' <param name="cookies">如果为Nothing，则只重新初始化类，不设置cookie</param>
        ''' <param name="allowAutoRedirect">指示处理程序是否应跟随重定向响应,默认为True</param>
        Public Sub ReInit(ByRef cookies As CookieContainer, ByVal allowAutoRedirect As Boolean)
            m_AllowAutoRedirect = allowAutoRedirect
            InitInternal(cookies, allowAutoRedirect, Nothing, Nothing)
        End Sub

        ''' <summary>
        ''' 用于重新初始化类，比如更换账号等涉及到重新传入cookie的情况
        ''' </summary>
        ''' <param name="cookies">如果为Nothing，则只重新初始化类，不设置cookie</param>
        ''' <param name="allowAutoRedirect">指示处理程序是否应跟随重定向响应,默认为True</param>
        ''' <param name="proxy">指示处理程序是否应设置代理信息,默认为Nothing</param>
        Public Sub ReInit(ByRef cookies As CookieContainer, ByVal allowAutoRedirect As Boolean, ByRef proxy As IWebProxy)
            m_AllowAutoRedirect = allowAutoRedirect
            InitInternal(cookies, allowAutoRedirect, proxy, Nothing)
        End Sub
        ''' <summary>
        ''' 用于重新初始化类，比如更换账号等涉及到重新传入cookie的情况
        ''' </summary>
        ''' <param name="cookies">如果为Nothing，则只重新初始化类，不设置cookie</param>
        ''' <param name="allowAutoRedirect">指示处理程序是否应跟随重定向响应,默认为True</param>
        ''' <param name="proxy">指示处理程序是否应设置代理信息,默认为Nothing</param>
        ''' <param name="charSet">解码响应文本使用的字符集，设置错误会导致乱码。设置之前，确保访问的每个网页的字符集都是一样的，否则建议使用无参数的构造函数，程序内部自动检查字符集，当然也会牺牲一点效率。</param>
        Private Sub ReInit(ByRef cookies As CookieContainer, ByVal allowAutoRedirect As Boolean, ByRef proxy As IWebProxy, ByVal charSet As String)
            m_AllowAutoRedirect = allowAutoRedirect
            InitInternal(cookies, allowAutoRedirect, proxy, charSet)
        End Sub

        ''' <summary>
        ''' 预热，缩短第一次正式请求的时间
        ''' </summary>
        Public Async Function PreHeatAsync() As Task
            Dim httpRequestMessage As New HttpRequestMessage With {
                .RequestUri = New Uri("https://t.alicdn.com/t/gettime"),
                .Method = New Http.HttpMethod("HEAD")
            }

            Try
                Dim rst = Await m_HttpClient.SendAsync(httpRequestMessage)
                rst.EnsureSuccessStatusCode()
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try
        End Function

        ''' <summary>
        ''' 发送HEAD请求获取响应头
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns>成功 Success 返回True，其他情况都返回False；
        ''' 成功ResponseHeaders返回获取到的响应头，其他情况都返回具体错误信息；
        ''' HttpStatusCode返回相应的http请求状态码</returns>
        Private Async Function InternalHeadAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of HeadResponse)

            Dim success As Boolean
            Dim statusCode As HttpStatusCode
            Dim responseHeaders As Headers.HttpResponseHeaders
            Dim cts As CancellationTokenSource

            Try
                cts = New CancellationTokenSource
                cts.CancelAfter(m_DefaulTimeoutMilliseconds)
                Dim ct = cts.Token
                ct.ThrowIfCancellationRequested()

                ' 初始化uri类实例
                Dim baseAddress As New Uri(url)

                ' 构造请求消息体
                Dim requestMessage As New HttpRequestMessage(Http.HttpMethod.Head, baseAddress)
                requestMessage.SetRequestHeaders(requestHeaders, HttpMethod.HEAD)

                ' 发送请求
                Dim headResponse = Await m_HttpClient.SendAsync(requestMessage, ct)

                success = (headResponse.StatusCode = HttpStatusCode.OK)
                statusCode = headResponse.StatusCode
                responseHeaders = headResponse.Headers
            Catch ex As UriFormatException
                statusCode = HttpStatusCode.BadRequest
            Catch ex As TaskCanceledException
                statusCode = HttpStatusCode.RequestTimeout
            Catch ex As Exception
                Logger.WriteLine(ex)
            Finally
                If cts IsNot Nothing Then
                    cts.Dispose()
                End If
            End Try

            Return New HeadResponse(success, statusCode, responseHeaders)
        End Function

        ''' <summary>
        ''' 获取请求返回的信息
        ''' </summary>
        ''' <param name="response"></param>
        ''' <returns></returns>
        Private Async Function InternalGetResponseStringAsync(ByVal response As HttpResponseMessage) As Task(Of HttpResponse)
            Dim success As Boolean
            Dim statusCode As HttpStatusCode
            Dim responseContent As String
            Dim header As HttpResponseHeaders

            Try
                ' 检查请求状态
                header = response.Headers
                statusCode = response.StatusCode
                If response.IsSuccessStatusCode Then
                    ' 获取响应文本
                    responseContent = Await ReadAsStringAsync(response)

                    statusCode = HttpStatusCode.OK
                    success = True
                ElseIf response.StatusCode = HttpStatusCode.Found Then
                    responseContent = HttpStatusCode.Found.ToString

                    ' 如果是302的话 responseContent 返回头信息
                    success = True
                Else
                    responseContent = response.ReasonPhrase
                End If
            Catch ex As TaskCanceledException
                statusCode = HttpStatusCode.RequestTimeout
                responseContent = ex.Message
            Catch ex As Exception
                statusCode = response.StatusCode
                responseContent = ex.Message

                Logger.WriteLine(ex)
            Finally
                If response IsNot Nothing Then
                    response.Dispose()
                End If
            End Try

            Return New HttpResponse(success, statusCode, responseContent, header)
        End Function

        Private Async Function ReadAsStringAsync(ByVal response As HttpResponseMessage) As Task(Of String)
            ' 修复 ReadAsStringAsync 引发异常 “utf8”不是支持的编码名。有关定义自定义编码的信息，请参阅关于 Encoding.RegisterProvider 方法的文档 
            ' 某些网页不规范，charset设置的是 charset=UTF8 而不是 charset=UTF-8，20200327
            Dim charset = response.Content.Headers.ContentType.CharSet
            If "utf8".Equals(charset, StringComparison.OrdinalIgnoreCase) Then
                response.Content.Headers.ContentType.CharSet = "utf-8"
            End If

            Return Await response.Content.ReadAsStringAsync
        End Function

        ''' <summary>
        ''' 获取网页源码
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns>成功 Success 返回True，其他情况都返回False；
        ''' 成功Message返回获取到的源码，其他情况都返回具体错误信息；
        ''' HttpStatusCode返回相应的http请求状态码</returns>
        Private Async Function InternalGetAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse
            Dim cts As CancellationTokenSource
            Try
                cts = New CancellationTokenSource
                cts.CancelAfter(m_DefaulTimeoutMilliseconds)
                Dim ct = cts.Token
                ct.ThrowIfCancellationRequested()

                ' 初始化uri类实例
                Dim baseAddress As New Uri(url)
                ' 构造请求消息体
                Dim requestMessage As New HttpRequestMessage(Http.HttpMethod.Get, baseAddress)
                requestMessage.SetRequestHeaders(requestHeaders, HttpMethod.GET)
                ct.ThrowIfCancellationRequested()
                ' 发送请求
                Dim response = Await m_HttpClient.SendAsync(requestMessage, ct)
                httpResponse = Await InternalGetResponseStringAsync(response)
            Catch ex As UriFormatException
                httpResponse = New HttpResponse(False, HttpStatusCode.BadRequest, ex.Message)
            Catch ex As TaskCanceledException
                httpResponse = New HttpResponse(False, HttpStatusCode.RequestTimeout, ex.Message)
            Catch ex As Exception
                httpResponse = New HttpResponse(False, HttpStatusCode.BadRequest, ex.ToString)
            Finally
                If cts IsNot Nothing Then
                    cts.Dispose()
                End If
            End Try

            Return httpResponse
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContentKvp">请求主体键值对集合，不需要编码，直接原字符串传入</param>
        ''' <param name="postContentEncoding">请求主体的编码方式，必须跟抓包的一致，否则可能会导致乱码</param>
        ''' <returns>注：如果同样的包，其他工具返回结果正常，本工具返回异常，请检查编码以及cookie的域</returns>
        Public Async Function InternalPostAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContentKvp As IEnumerable(Of KeyValuePair(Of String, String)), ByVal postContentEncoding As Text.Encoding) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse
            Dim cts As CancellationTokenSource

            Try
                cts = New CancellationTokenSource
                cts.CancelAfter(m_DefaulTimeoutMilliseconds)
                Dim ct = cts.Token
                ct.ThrowIfCancellationRequested()

                Dim content = GetHttpContent(requestHeaders, postContentKvp, postContentEncoding)

                ct.ThrowIfCancellationRequested()

                ' 初始化uri类实例
                Dim baseAddress As New Uri(url)
                ' 构造请求消息体
                Dim requestMessage As New HttpRequestMessage(Http.HttpMethod.Post, baseAddress) With {
                    .Content = content
                }
                requestMessage.SetRequestHeaders(requestHeaders, HttpMethod.POST)

                ct.ThrowIfCancellationRequested()

                ' 发送请求
                Dim response = Await m_HttpClient.SendAsync(requestMessage, ct)
                httpResponse = Await InternalGetResponseStringAsync(response)
            Catch ex As UriFormatException
                httpResponse = New HttpResponse(False, HttpStatusCode.BadRequest, ex.Message)
            Catch ex As TaskCanceledException
                httpResponse = New HttpResponse(False, HttpStatusCode.RequestTimeout, ex.Message)
            Catch ex As Exception
                httpResponse = New HttpResponse(False, HttpStatusCode.BadRequest, ex.Message)

                Logger.WriteLine(ex)
            Finally
                If cts IsNot Nothing Then
                    cts.Dispose()
                End If
            End Try

            Return httpResponse
        End Function

        Private Function GetHttpContent(ByVal requestHeaders As Dictionary(Of String, String), ByVal postContentKvp As IEnumerable(Of KeyValuePair(Of String, String)), ByVal postContentEncoding As Text.Encoding) As HttpContent
            Dim content As HttpContent
            Dim mediaType As String
            ' 设置请求头(如果请求头包含content-type，那么得先获取到)
            ' 如果能确保 postContent 里面没有特殊字符，比如 + = & 这些，那就可以用 StringContent，否则建议用 FormUrlEncodedContent（内部会对特殊字符转换）
            If Not requestHeaders.TryGetValue("content-type", mediaType) Then
                mediaType = DefaulMediaType
            End If
            ' 若 post结果不符合预期,比如乱码、丢字，可以尝试修改此处
            If postContentEncoding.CodePage = Text.Encoding.UTF8.CodePage Then
                content = New FormUrlEncodedContent(postContentKvp)
            Else
                Dim postContent = EncodeContent(postContentKvp, postContentEncoding)
                content = New StringContent(postContent, postContentEncoding)
            End If
            Dim mediaTypeHeaderValue = New MediaTypeHeaderValue(mediaType)
            If postContentEncoding IsNot Nothing Then
                mediaTypeHeaderValue.CharSet = postContentEncoding.WebName
            End If
            content.Headers.ContentType = mediaTypeHeaderValue

            Return content
        End Function

        ''' <summary>
        ''' 对请求文本进行UrlEncode
        ''' </summary>
        ''' <param name="postContentKvp"></param>
        ''' <param name="postContentEncoding"></param>
        ''' <returns></returns>
        Private Function EncodeContent(ByVal postContentKvp As IEnumerable(Of KeyValuePair(Of String, String)), ByVal postContentEncoding As Text.Encoding) As String
            Dim sb = New StringBuilder(361)
            For Each kvp As KeyValuePair(Of String, String) In postContentKvp
                If sb.Length > 0 Then
                    sb.Append("&"c)
                End If
                sb.Append(kvp.Key)
                sb.Append("="c)
                sb.Append(kvp.Value.UrlEncode(postContentEncoding))
            Next
            Dim postContent = sb.ToString
            Return postContent
        End Function

        ''' <summary>
        ''' 异步下载文件
        ''' </summary>
        ''' <param name="url">下载链接</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns></returns>
        Private Async Function InternalDownloadFileAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of DownloadResponse)

            Dim success As Boolean
            Dim statusCode As HttpStatusCode
            Dim memoryStream As New MemoryStream
            Dim cts As CancellationTokenSource

            Dim registDownloadProgressChangedEventHandler As Boolean
            Try
                ' 初始化uri类实例
                Dim baseAddress As New Uri(url)

                ' 构造请求消息体
                Dim requestMessage As New HttpRequestMessage(Http.HttpMethod.Get, baseAddress)
                requestMessage.SetRequestHeaders(requestHeaders, HttpMethod.GET)

#Region "下载文件主体"
                ' 注册进度变化事件
                AddHandler m_ProgressMessageHandler.HttpReceiveProgress, AddressOf DownloadProgressChangedEventHandler
                registDownloadProgressChangedEventHandler = True

                cts = New CancellationTokenSource
                cts.CancelAfter(Timeout.InfiniteTimeSpan)
                Dim ct = cts.Token

                ' 发送请求
                Dim response = Await m_HttpClient.SendAsync(requestMessage, ct)
                ct.ThrowIfCancellationRequested()

                Using responseStream = Await response.Content.ReadAsStreamAsync
                    If responseStream IsNot Nothing Then
                        Await responseStream.CopyToAsync(memoryStream)
                    End If
                End Using
                ' 报告下载完成
                RaiseEvent DownloadFileCompleted(Nothing, New DownloadFileCompletedEventArgs(Nothing, False, Nothing))
#End Region

                success = True
                statusCode = HttpStatusCode.OK
            Catch ex As UriFormatException
                statusCode = HttpStatusCode.BadRequest

                ' 报告下载取消
                RaiseEvent DownloadFileCompleted(Nothing, New DownloadFileCompletedEventArgs(ex, True, Nothing))
            Catch ex As TaskCanceledException
                statusCode = HttpStatusCode.RequestTimeout

                ' 报告下载取消
                RaiseEvent DownloadFileCompleted(Nothing, New DownloadFileCompletedEventArgs(ex, True, Nothing))
            Catch ex As Exception
                ' 报告下载取消
                RaiseEvent DownloadFileCompleted(Nothing, New DownloadFileCompletedEventArgs(ex, True, Nothing))

                If ex.Message.IndexOf("404") = -1 Then
                    Logger.WriteLine(ex)
                Else
                    statusCode = HttpStatusCode.NotFound
                End If
            Finally
                If registDownloadProgressChangedEventHandler Then
                    RemoveHandler m_ProgressMessageHandler.HttpReceiveProgress, AddressOf DownloadProgressChangedEventHandler
                End If
                If cts IsNot Nothing Then
                    cts.Dispose()
                End If
            End Try

            Return New DownloadResponse(success, statusCode, memoryStream)
        End Function

        ''' <summary>
        ''' 异步下载文件
        ''' </summary>
        ''' <param name="url">下载链接</param>
        ''' <param name="fileFullPath">文件存储路径</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns></returns>
        Private Async Function InternalDownloadFileAsync(ByVal url As String, ByVal fileFullPath As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of HttpResponse)
#Region "防呆检查"
            ' 判断传入的路径是否为文件（没有后缀名就视为非文件）
            Dim isFile = Path.GetExtension(fileFullPath).Length > 0
            If Not isFile Then
                Return New HttpResponse(False, HttpStatusCode.BadRequest, $"{NameOf(fileFullPath)}({fileFullPath}) is not a file")
            End If

            ' 判断文件夹是否存在
            Dim isDirectoryExists = IO.Directory.Exists(Path.GetDirectoryName(fileFullPath))
            If Not isDirectoryExists Then
                Return New HttpResponse(False, HttpStatusCode.BadRequest, $"Directory not found of {NameOf(fileFullPath)}({fileFullPath})")
            End If
#End Region
            Dim success As Boolean
            Dim statusCode As HttpStatusCode
            Dim responseContent As String
            Dim cts As CancellationTokenSource
            Dim response As HttpResponseMessage

            Dim registDownloadProgressChangedEventHandler As Boolean
            Try
                ' 初始化uri类实例
                Dim baseAddress As New Uri(url)

                ' 构造请求消息体
                Dim requestMessage As New HttpRequestMessage(Http.HttpMethod.Get, baseAddress)
                requestMessage.SetRequestHeaders(requestHeaders, HttpMethod.GET)

#Region "下载文件主体"
                ' 注册进度变化事件
                AddHandler m_ProgressMessageHandler.HttpReceiveProgress, AddressOf DownloadProgressChangedEventHandler
                registDownloadProgressChangedEventHandler = True

                cts = New CancellationTokenSource
                cts.CancelAfter(Timeout.InfiniteTimeSpan)
                Dim ct = cts.Token

                ' 发送请求
                response = Await m_HttpClient.SendAsync(requestMessage)
                ct.ThrowIfCancellationRequested()

                Dim responseStream = Await response.Content.ReadAsStreamAsync
                Dim bufferSize = 80 * 1024
                Using fileStream As New FileStream(fileFullPath,
                                                          FileMode.Create,
                                                          FileAccess.ReadWrite,
                                                          FileShare.ReadWrite,
                                                          bufferSize,
                                                          True)
                    ' 可以用 CopyToAsync 方法一步写文件，但是不能做进度控制，没法向用户反馈
                    ' 所以还是用下面的方法
                    ' 关于效率：测试的时候 两种方法效率很接近(只测试过几最大M的文件)
                    Await responseStream.CopyToAsync(fileStream)

                    ' 报告下载完成
                    RaiseEvent DownloadFileCompleted(Nothing, New DownloadFileCompletedEventArgs(Nothing, False, Nothing))
                End Using

                If responseStream IsNot Nothing Then
                    responseStream.Dispose()
                End If
#End Region

                success = True
                statusCode = HttpStatusCode.OK
                responseContent = fileFullPath
            Catch ex As UriFormatException
                statusCode = HttpStatusCode.BadRequest
                responseContent = ex.Message

                ' 报告下载取消
                RaiseEvent DownloadFileCompleted(Nothing, New DownloadFileCompletedEventArgs(ex, True, Nothing))
            Catch ex As TaskCanceledException
                statusCode = HttpStatusCode.RequestTimeout
                responseContent = ex.Message

                ' 报告下载取消
                RaiseEvent DownloadFileCompleted(Nothing, New DownloadFileCompletedEventArgs(ex, True, Nothing))
            Catch ex As Exception
                responseContent = ex.Message

                ' 报告下载取消
                RaiseEvent DownloadFileCompleted(Nothing, New DownloadFileCompletedEventArgs(ex, True, Nothing))

                If ex.Message.IndexOf("404") = -1 Then
                    Logger.WriteLine(ex)
                Else
                    statusCode = HttpStatusCode.NotFound
                End If
            Finally
                If registDownloadProgressChangedEventHandler Then
                    RemoveHandler m_ProgressMessageHandler.HttpReceiveProgress, AddressOf DownloadProgressChangedEventHandler
                End If
                If cts IsNot Nothing Then
                    cts.Dispose()
                End If
            End Try

            Return New HttpResponse(success, statusCode, responseContent, response.Headers)
        End Function

        Private Sub DownloadProgressChangedEventHandler(sender As Object, e As HttpProgressEventArgs)
            RaiseEvent DownloadProgressChanged(sender, New DownloadProgressChangedEventArgs(e.ProgressPercentage, e.UserState, e.BytesTransferred, e.TotalBytes))
        End Sub

        Private Sub UploadProgressChangedEventHandler(sender As Object, e As HttpProgressEventArgs)
            RaiseEvent UploadProgressChanged(sender, New UploadProgressChangedEventArgs(e.ProgressPercentage, e.UserState, e.BytesTransferred, e.TotalBytes))
        End Sub

        ''' <summary>
        ''' 异步上传文件
        ''' </summary>
        ''' <param name="uploadInfo">上传文件需要的相关信息</param>
        ''' <param name="requestHeaders"></param>
        ''' <returns></returns>
        Private Async Function InternalUploadFileAsync(ByVal uploadInfo As UploadInfo, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of HttpResponse)
            Dim success As Boolean
            Dim statusCode As HttpStatusCode
            Dim responseContent As String
            Dim response As HttpResponseMessage

            Using content = New MultipartFormDataContent()
                ' 添加数据源（没有特殊情况，一般不需要设置 boundary,类库里面会自动生成）
                For Each ct In uploadInfo.HttpContents
                    If TypeOf ct.Content Is StringContent Then
                        content.Add(ct.Content, ct.Name)
                    Else
                        If ct.Content.Headers.ContentType Is Nothing Then
                            ' 某些服务器要求ContentType 必须符合上传文件类型，否则会上传失败，返回 “只支持jpg或png格式的图片”
                            Dim mediaType = FileContentTypeHelper.GetMimeType(ct.FileName)
                            ct.Content.Headers.ContentType = New MediaTypeHeaderValue(mediaType)
                        End If
                        content.Add(ct.Content, ct.Name, ct.FileName)
                    End If
                Next
                ' 查看数据源 Await content.ReadAsStringAsync

                Dim registUploadProgressChangedEventHandler As Boolean
                Try
                    ' 初始化uri类实例
                    Dim baseAddress As New Uri(uploadInfo.RequestUrl)
                    ' 构造请求消息体
                    Dim requestMessage As New HttpRequestMessage(Http.HttpMethod.Post, baseAddress) With {
                        .Content = content
                    }
                    requestMessage.SetRequestHeaders(requestHeaders, HttpMethod.POST)

                    ' 注册进度变化事件
                    AddHandler m_ProgressMessageHandler.HttpSendProgress, AddressOf UploadProgressChangedEventHandler
                    registUploadProgressChangedEventHandler = True

                    ' 开始上传数据
                    response = Await m_HttpClient.SendAsync(requestMessage)
                    responseContent = Await ReadAsStringAsync(response)

                    ' 报告上传完成
                    RaiseEvent UploadFileCompleted(Nothing, New UploadFileCompletedEventArgs(Nothing, False, Nothing))

                    success = True
                    statusCode = HttpStatusCode.OK
                Catch ex As UriFormatException
                    statusCode = HttpStatusCode.BadRequest
                    responseContent = ex.Message

                    ' 报告上传取消
                    RaiseEvent UploadFileCompleted(Nothing, New UploadFileCompletedEventArgs(ex, True, Nothing))
                Catch ex As TaskCanceledException
                    statusCode = HttpStatusCode.RequestTimeout
                    responseContent = ex.Message

                    ' 报告上传取消
                    RaiseEvent UploadFileCompleted(Nothing, New UploadFileCompletedEventArgs(ex, True, Nothing))
                Catch ex As Exception
                    responseContent = ex.Message

                    ' 报告上传取消
                    RaiseEvent UploadFileCompleted(Nothing, New UploadFileCompletedEventArgs(ex, True, Nothing))

                    statusCode = HttpStatusCode.BadRequest
                Finally
                    If registUploadProgressChangedEventHandler Then
                        RemoveHandler m_ProgressMessageHandler.HttpSendProgress, AddressOf UploadProgressChangedEventHandler
                    End If
                End Try

                Return New HttpResponse(success, statusCode, responseContent, response.Headers)
            End Using
        End Function

        ''' <summary>
        ''' Http请求公共设置，作用于整个类
        ''' </summary>
        Private Sub HttpPublicConfig()
            ' 对ServicePointManager的设置将会影响到webrequest从而影响httpwebrequest
            ' 这个值最好不要超过1024
            ServicePointManager.DefaultConnectionLimit = 128
            ' 去掉“Expect: 100-Continue”请求头，不然会引起post（417） expectation failed
            ServicePointManager.Expect100Continue = False
            ServicePointManager.DnsRefreshTimeout = 5000
            ServicePointManager.UseNagleAlgorithm = True
            ' HttpWebRequest 的请求因为网络问题导致连接没有被释放则会占用连接池中的连接个数，导致并发连接数量减少
            ServicePointManager.SetTcpKeepAlive(True, 1000 * 30, 2)
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 Or
                SecurityProtocolType.Tls Or
                SecurityProtocolType.Tls11 Or
                SecurityProtocolType.Tls12
            '' 不使用缓存 使用缓存可能会得到错误的结果
            'WebRequest.DefaultCachePolicy = New Cache.RequestCachePolicy(Cache.RequestCacheLevel.NoCacheNoStore)
            ServicePointManager.ServerCertificateValidationCallback = AddressOf TrustAllValidationCallback
        End Sub

        Private Function TrustAllValidationCallback(sender As Object, certificate As X509Certificate, chain As X509Chain, errors As SslPolicyErrors) As Boolean
            Return True
        End Function

        ''' <summary>
        ''' 设置默认的 CharSet
        ''' </summary>
        ''' <param name="charSetName"></param>
        Public Sub ModifyDefaultCharSet(ByVal charSetName As String)
            DefaultCharSet = charSetName
        End Sub

        ''' <summary>
        ''' 设置请求头
        ''' </summary>
        ''' <param name="requestHeaders">请求头</param>
        ''' <param name="method">请求方法</param>
        Public Sub SetRequestHeaders(ByRef requestHeaders As Dictionary(Of String, String), ByVal method As HttpMethod)
            m_HttpClient.SetRequestHeaders(requestHeaders, method)
        End Sub
#End Region

#Region "函数区"
        ''' <summary>
        ''' 发送HEAD请求获取响应头 
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="requestHeaders"></param>
        ''' <returns>成功 Success 返回True，其他情况都返回False；
        ''' 成功ResponseHeaders返回获取到的响应头，其他情况都返回具体错误信息；
        ''' HttpStatusCode返回相应的http请求状态码</returns>
        Public Async Function HeadAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of HeadResponse)
            Return Await InternalHeadAsync(url, requestHeaders)
        End Function

        ''' <summary>
        ''' 获取网页源码
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns>成功 Success 返回True，其他情况都返回False；
        ''' 成功Message返回获取到的源码，其他情况都返回具体错误信息；
        ''' HttpStatusCode返回相应的http请求状态码</returns>
        Public Async Function GetAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of HttpResponse)
            Return Await InternalGetAsync(url, requestHeaders)
        End Function

        ''' <summary>
        ''' 获取网页源码
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <returns></returns>
        Public Async Function GetAsync(ByVal url As String) As Task(Of HttpResponse)
            Return Await InternalGetAsync(url, Nothing)
        End Function

        ''' <summary>
        ''' 尝试获取网页源码直到成功，最多尝试三次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <returns></returns>
        Public Async Function TryGetThreeTimeIfErrorAsync(ByVal url As String) As Task(Of HttpResponse)
            Return Await TryGetAsync(url, 3)
        End Function

        ''' <summary>
        ''' 尝试获取网页源码直到成功，最多尝试 <paramref name="tryTime"/> 次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Async Function TryGetAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), tryTime As Integer) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse

            Do
                httpResponse = Await InternalGetAsync(url, requestHeaders)
                tryTime -= 1
            Loop Until httpResponse.Success OrElse tryTime <= 0

            Return httpResponse
        End Function

        ''' <summary>
        ''' 尝试获取网页源码直到成功，最多尝试三次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <returns></returns>
        Public Async Function TryGetAsync(ByVal url As String, ByVal tryTime As Integer) As Task(Of HttpResponse)
            Dim header As Dictionary(Of String, String)
            Return Await TryGetAsync(url, header, tryTime)
        End Function

        ''' <summary>
        ''' 执行Get请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <returns></returns>
        Public Async Function TryGetAsync(ByVal url As String, ByVal retureIfContain As String, ByVal tryTime As Integer) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse

            Do
                httpResponse = Await TryGetAsync(url, 1)
                tryTime -= 1
            Loop Until (httpResponse.Success AndAlso httpResponse.Message.IndexOf(retureIfContain) > -1) OrElse tryTime <= 0
            If httpResponse.Success Then
                httpResponse = New HttpResponse(httpResponse.Message.IndexOf(retureIfContain) > -1 OrElse tryTime > 0, httpResponse.StatusCode, httpResponse.Message, httpResponse.Header)
            End If

            Return httpResponse
        End Function

        ''' <summary>
        ''' 执行Get请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <returns></returns>
        Public Async Function TryGetAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal retureIfContain As String, ByVal tryTime As Integer) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse

            Do
                httpResponse = Await TryGetAsync(url, requestHeaders, 1)
                tryTime -= 1
            Loop Until (httpResponse.Success AndAlso httpResponse.Message.IndexOf(retureIfContain) > -1) OrElse tryTime <= 0
            If httpResponse.Success Then
                httpResponse = New HttpResponse(httpResponse.Message.IndexOf(retureIfContain) > -1 OrElse tryTime > 0, httpResponse.StatusCode, httpResponse.Message, httpResponse.Header)
            End If

            Return httpResponse
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入。如果传入有包含编码的数据，可能会导致乱码或者数据丢失</param>
        ''' <param name="postContentEncoding">>请求主体的编码方式，必须跟抓包的一致，否则可能会导致乱码</param>
        ''' <returns>注：如果同样的包，其他工具返回结果正常，本工具返回异常，请检查编码以及cookie的域</returns>
        Public Async Function PostAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String, ByVal postContentEncoding As Text.Encoding) As Task(Of HttpResponse)
            Return Await InternalPostAsync(url, requestHeaders, postContent.ToKeyValuePairs, postContentEncoding)
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入；内部默认使用UTF-8编码，如果抓包得到的不是UTF-8编码，请调用可以自定义编码方式的重载函数</param>
        ''' <returns>注：如果同样的包，其他工具返回结果正常，本工具返回异常，请检查编码以及cookie的域</returns>
        Public Async Function PostAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String) As Task(Of HttpResponse)
            Return Await PostAsync(url, requestHeaders, postContent, Text.Encoding.UTF8)
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入。如果传入有包含编码的数据，可能会导致乱码或者数据丢失</param>
        ''' <param name="postContentEncoding">请求主体的编码方式，必须跟抓包的一致，否则可能会导致乱码</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param> 
        ''' <returns>注：如果同样的包，其他工具返回结果正常，本工具返回异常，请检查编码以及cookie的域</returns>
        Public Async Function TryPostAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String, ByVal postContentEncoding As Text.Encoding, tryTime As Integer) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse

            Do
                httpResponse = Await PostAsync(url, requestHeaders, postContent, postContentEncoding)
                tryTime -= 1
            Loop Until httpResponse.Success OrElse tryTime <= 0

            Return httpResponse
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入；内部默认使用UTF-8编码，如果抓包得到的不是UTF-8编码，请调用可以自定义编码方式的重载函数</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns>注：如果同样的包，其他工具返回结果正常，本工具返回异常，请检查编码以及cookie的域</returns>
        Public Async Function TryPostAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String, tryTime As Integer) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse

            Do
                httpResponse = Await PostAsync(url, requestHeaders, postContent, Text.Encoding.UTF8)
                tryTime -= 1
            Loop Until httpResponse.Success OrElse tryTime <= 0

            Return httpResponse
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入。如果传入有包含编码的数据，可能会导致乱码或者数据丢失</param>
        ''' <param name="postContentEncoding">请求主体的编码方式，必须跟抓包的一致，否则可能会导致乱码</param>
        ''' <returns>注：如果同样的包，其他工具返回结果正常，本工具返回异常，请检查编码以及cookie的域</returns>
        Public Async Function TryPostThreeTimeIfErrorAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String, ByVal postContentEncoding As Text.Encoding) As Task(Of HttpResponse)
            Return Await TryPostAsync(url, requestHeaders, postContent, postContentEncoding, 3)
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入；内部默认使用UTF-8编码，如果抓包得到的不是UTF-8编码，请调用可以自定义编码方式的重载函数</param>
        ''' <returns>注：如果同样的包，其他工具返回结果正常，本工具返回异常，请检查编码以及cookie的域</returns>
        Public Async Function TryPostThreeTimeIfErrorAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String) As Task(Of HttpResponse)
            Return Await TryPostAsync(url, requestHeaders, postContent, 3)
        End Function

        ''' <summary>
        ''' 异步下载文件
        ''' </summary>
        ''' <param name="url">下载链接</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns></returns>
        Public Async Function DownloadFileAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of DownloadResponse)
            Return Await InternalDownloadFileAsync(url, requestHeaders)
        End Function

        ''' <summary>
        ''' 尝试下载文件直到成功，最多尝试 <paramref name="tryTime"/> 次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Async Function TryDownloadFileAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), tryTime As Integer) As Task(Of DownloadResponse)
            Dim httpResponse As DownloadResponse

            Do
                httpResponse = Await InternalDownloadFileAsync(url, requestHeaders)
                tryTime -= 1
            Loop Until httpResponse.Success OrElse tryTime <= 0

            Return httpResponse
        End Function

        ''' <summary>
        ''' 尝试下载文件直到成功，最多尝试三次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <returns></returns>
        Public Async Function TryDownloadFileThreeTimeIfErrorAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of DownloadResponse)
            Return Await TryDownloadFileAsync(url, requestHeaders, 3)
        End Function

        ''' <summary>
        ''' 异步下载文件
        ''' </summary>
        ''' <param name="url">下载链接</param>
        ''' <param name="fileFullPath">文件存储路径</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns></returns>
        Public Async Function DownloadFileAsync(ByVal url As String, ByVal fileFullPath As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of HttpResponse)
            Return Await InternalDownloadFileAsync(url, fileFullPath, requestHeaders)
        End Function

        ''' <summary>
        ''' 尝试下载文件直到成功，最多尝试 <paramref name="tryTime"/> 次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Async Function TryDownloadFileAsync(ByVal url As String, ByVal fileFullPath As String, tryTime As Integer) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse

            Do
                httpResponse = Await InternalDownloadFileAsync(url, fileFullPath, Nothing)
                tryTime -= 1
            Loop Until httpResponse.Success OrElse tryTime <= 0

            Return httpResponse
        End Function

        ''' <summary>
        ''' 尝试下载文件直到成功，最多尝试 <paramref name="tryTime"/> 次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Async Function TryDownloadFileAsync(ByVal url As String, ByVal fileFullPath As String, ByVal requestHeaders As Dictionary(Of String, String), tryTime As Integer) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse

            Do
                httpResponse = Await InternalDownloadFileAsync(url, fileFullPath, requestHeaders)
                tryTime -= 1
            Loop Until httpResponse.Success OrElse tryTime <= 0

            Return httpResponse
        End Function

        ''' <summary>
        ''' 尝试下载文件直到成功，最多尝试三次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <returns></returns>
        Public Async Function TryDownloadFileThreeTimeIfErrorAsync(ByVal url As String, ByVal fileFullPath As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of HttpResponse)
            Return Await TryDownloadFileAsync(url, fileFullPath, requestHeaders, 3)
        End Function

        ''' <summary>
        ''' 异步上传文件
        ''' </summary>
        ''' <param name="uploadInfo">上传文件需要的相关信息</param>
        ''' <param name="requestHeaders"></param>
        ''' <returns></returns>
        Public Async Function UploadFileAsync(ByVal uploadInfo As UploadInfo, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of HttpResponse)
            Return Await InternalUploadFileAsync(uploadInfo, requestHeaders)
        End Function


        ''' <summary>
        ''' 尝试上传文件直到成功，最多尝试 <paramref name="tryTime"/> 次
        ''' </summary>
        ''' <param name="uploadInfo">上传文件需要的相关信息</param>
        ''' <param name="requestHeaders"></param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Async Function TryUploadFileAsync(ByVal uploadInfo As UploadInfo, ByVal requestHeaders As Dictionary(Of String, String), tryTime As Integer) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse

            Do
                httpResponse = Await InternalUploadFileAsync(uploadInfo, requestHeaders)
                tryTime -= 1
            Loop Until httpResponse.Success OrElse tryTime <= 0

            Return httpResponse
        End Function

        ''' <summary>
        ''' 尝试上传文件直到成功，最多尝试三次
        ''' </summary>
        ''' <param name="uploadInfo">上传文件需要的相关信息</param>
        ''' <param name="requestHeaders"></param>
        ''' <returns></returns>
        Public Async Function TryUploadFileThreeTimeIfErrorAsync(ByVal uploadInfo As UploadInfo, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of HttpResponse)
            Return Await TryUploadFileAsync(uploadInfo, requestHeaders, 3)
        End Function
#End Region


#Region "IDisposable Support"
        ' 要检测冗余调用
        Private disposedValue As Boolean = False

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If disposedValue Then Return

            If disposing Then
                ' TODO: 释放托管状态(托管对象)。
                If m_ProgressMessageHandler IsNot Nothing Then
                    m_ProgressMessageHandler.Dispose()
                    m_ProgressMessageHandler = Nothing
                End If

                If m_HttpClientHandler IsNot Nothing Then
                    m_HttpClientHandler.Dispose()
                    m_HttpClientHandler = Nothing
                End If

                If m_HttpClient IsNot Nothing Then
                    m_HttpClient.Dispose()
                    m_HttpClient = Nothing
                End If
            End If

            ' TODO: 释放未托管资源(未托管对象)并在以下内容中替代 Finalize()。
            ' TODO: 将大型字段设置为 null。


            disposedValue = True
        End Sub

        ' TODO: 仅当以上 Dispose(disposing As Boolean)拥有用于释放未托管资源的代码时才替代 Finalize()。
        'Protected Overrides Sub Finalize()
        '    ' 请勿更改此代码。将清理代码放入以上 Dispose(disposing As Boolean)中。
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' Visual Basic 添加此代码以正确实现可释放模式。
        Public Sub Dispose() Implements IDisposable.Dispose
            ' 请勿更改此代码。将清理代码放入以上 Dispose(disposing As Boolean)中。
            Dispose(True)
            ' TODO: 如果在以上内容中替代了 Finalize()，则取消注释以下行。
            ' GC.SuppressFinalize(Me)
        End Sub
#End Region
    End Class

End Namespace
