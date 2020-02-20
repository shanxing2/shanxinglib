Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Handlers
Imports System.Net.Http.Headers
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
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
        Public Shared Event DownloadProgressChanged As EventHandler(Of DownloadProgressChangedEventArgs)
        Public Shared Event DownloadFileCompleted As EventHandler(Of DownloadFileCompletedEventArgs)
        Public Shared Event UploadProgressChanged As EventHandler(Of UploadProgressChangedEventArgs)
        Public Shared Event UploadFileCompleted As EventHandler(Of UploadFileCompletedEventArgs)
#End Region

#Region "实例属性区"

#End Region

#Region "共享属性区"
        Private Shared s_IsInitialized As Boolean
        ''' <summary>
        ''' 指示此类是否已经初始化并且可用
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property IsInitialized() As Boolean
            Get
                Return s_IsInitialized
            End Get
        End Property

        Public Shared ReadOnly Property Cookies() As CookieContainer
            Get
                Return s_HttpClientHandler.CookieContainer
            End Get
        End Property

        Public Shared ReadOnly Property RequestHeaders() As HttpRequestHeaders
            Get
                Return s_HttpClient.DefaultRequestHeaders
            End Get
        End Property

        ''' <summary>
        ''' 默认的编码字符集为GBK
        ''' </summary>
        ''' <returns></returns>
        Public Shared Property DefaultCharSet As String

        Private Shared s_AllowAutoRedirect As Boolean
        ''' <summary>
        ''' 指示请求是否启用自动重定向
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property AllowAutoRedirect() As Boolean
            Get
                Return s_AllowAutoRedirect
            End Get
        End Property

        Private Shared s_DefaulTimeoutMilliseconds As Integer
        ''' <summary>
        ''' 默认的超时时间
        ''' </summary>
        ''' <returns></returns>
        Public Shared Property DefaulTimeoutMilliseconds As Integer
            Set(value As Integer)
                s_DefaulTimeoutMilliseconds = value

                ' 如果已经实例化，需要同时更新 Timeout 字段以使设置生效
                If s_HttpClient Is Nothing Then Return
                ReInit(s_HttpClient, s_HttpClientHandler, s_ProgressMessageHandler, Cookies, s_AllowAutoRedirect, Proxy, DefaultCharSet)
            End Set
            Get
                Return s_DefaulTimeoutMilliseconds
            End Get
        End Property

        ''' <summary>
        ''' 获取处理程序使用的代理信息
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property Proxy() As IWebProxy
            Get
                Return s_HttpClientHandler.Proxy
            End Get
        End Property

        ''' <summary>
        ''' 获取是否使用的代理信息
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property UseProxy() As Boolean
            Get
                Return s_HttpClientHandler.UseProxy
            End Get
        End Property
#End Region

#Region "字段区"
        '<ThreadStatic>
        'Private Shared m_CancellationTokenSource As CancellationTokenSource
        '<ThreadStatic>
        'Private Shared m_CancellationToken As CancellationToken
        Private m_HttpClient As HttpClient
        Private m_HttpClientHandler As HttpClientHandler
        Private m_ProgressMessageHandler As ProgressMessageHandler
        Private Shared s_HttpClient As HttpClient
        Private Shared s_HttpClientHandler As HttpClientHandler
        Private Shared s_ProgressMessageHandler As ProgressMessageHandler
#End Region

#Region "构造函数区"
        ''' <summary>
        ''' 类构造函数
        ''' 类之内的任意一个静态方法第一次调用时调用此构造函数
        ''' 而且程序生命周期内仅调用一次
        ''' </summary>
        Shared Sub New()
            HttpPublicConfig()

            DefaultCharSet = "GBK"
            s_DefaulTimeoutMilliseconds = 10000
        End Sub

        ''' <summary>
        ''' 不传入cookies，自动处理重定向
        ''' </summary>
        Sub New()
            HttpPublicConfig()

            InitInternal(m_HttpClient, m_HttpClientHandler, m_ProgressMessageHandler, Nothing, True, Nothing, Nothing)
            PreHeat(m_HttpClient)

            s_DefaulTimeoutMilliseconds = 10000
        End Sub

        ''' <summary>
        ''' 传入cookies，自动处理重定向
        ''' </summary>
        ''' <param name="cookies"></param>

        Sub New(ByRef cookies As CookieContainer)
            HttpPublicConfig()

            InitInternal(m_HttpClient, m_HttpClientHandler, m_ProgressMessageHandler, cookies, True, Nothing, Nothing)
            PreHeat(m_HttpClient)
        End Sub

        ''' <summary>
        ''' 不传入cookies，重定向由参数 <paramref name="allowAutoRedirect"/> 决定
        ''' </summary>
        ''' <param name="allowAutoRedirect">指示处理程序是否应跟随重定向响应,默认为True</param>
        Sub New(ByVal allowAutoRedirect As Boolean)
            HttpPublicConfig()

            s_AllowAutoRedirect = allowAutoRedirect
            InitInternal(m_HttpClient, m_HttpClientHandler, m_ProgressMessageHandler, Nothing, allowAutoRedirect, Nothing, Nothing)
            PreHeat(m_HttpClient)
        End Sub

        ''' <summary>
        ''' 不传入cookies，自动重定向，代理信息由参数 <paramref name="proxy"/> 决定
        ''' </summary>
        ''' <param name="proxy">指示处理程序是否应设置代理信息,默认为Nothing</param>
        Sub New(ByRef proxy As IWebProxy)
            HttpPublicConfig()

            InitInternal(m_HttpClient, m_HttpClientHandler, m_ProgressMessageHandler, Nothing, True, proxy, Nothing)
            PreHeat(m_HttpClient)
        End Sub

        ''' <summary>
        ''' 传入cookies，重定向由参数 <paramref name="allowAutoRedirect"/> 决定
        ''' </summary>
        ''' <param name="cookies"></param>
        ''' <param name="allowAutoRedirect">指示处理程序是否应跟随重定向响应,默认为True</param>
        Sub New(ByRef cookies As CookieContainer, ByVal allowAutoRedirect As Boolean)
            HttpPublicConfig()

            s_AllowAutoRedirect = allowAutoRedirect
            InitInternal(m_HttpClient, m_HttpClientHandler, m_ProgressMessageHandler, cookies, allowAutoRedirect, Nothing, Nothing)
            PreHeat(m_HttpClient)
        End Sub

        ''' <summary>
        ''' 传入cookies，重定向由参数 <paramref name="allowAutoRedirect"/> 决定，代理信息由参数 <paramref name="proxy"/> 决定
        ''' </summary>
        ''' <param name="cookies"></param>
        ''' <param name="allowAutoRedirect">指示处理程序是否应跟随重定向响应,默认为True</param>
        ''' <param name="proxy">指示处理程序是否应设置代理信息,默认为Nothing</param>
        Sub New(ByRef cookies As CookieContainer, ByVal allowAutoRedirect As Boolean, ByRef proxy As IWebProxy)
            HttpPublicConfig()

            s_AllowAutoRedirect = allowAutoRedirect
            InitInternal(m_HttpClient, m_HttpClientHandler, m_ProgressMessageHandler, cookies, allowAutoRedirect, proxy, Nothing)
            PreHeat(m_HttpClient)
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

            s_AllowAutoRedirect = allowAutoRedirect
            InitInternal(m_HttpClient, m_HttpClientHandler, m_ProgressMessageHandler, cookies, allowAutoRedirect, proxy, charSet)
            PreHeat(m_HttpClient)
        End Sub
#End Region

#Region "共享函数 实例函数共用，这个区块内不应该使用共享字段、属性等"
        ''' <summary>
        ''' 预热
        ''' </summary>
        ''' <param name="httpClient"></param>
        Private Shared Sub PreHeat(ByRef httpClient As HttpClient)
            Dim httpRequestMessage As New HttpRequestMessage With {
                .RequestUri = New Uri("https://t.alicdn.com/t/gettime"),
                .Method = New Http.HttpMethod("HEAD")
            }

            ' 因为可能需要在构造函数里面调用预热函数，所以不能 用  Await/Async 模式处理异步
            Dim tmpTask = httpClient.SendAsync(httpRequestMessage)
            tmpTask.ContinueWith(
                Sub(taskCompletion As Task(Of HttpResponseMessage))
                    ' ###########################################################
                    ' # 在ContinueWith块中，调试命中断点的情况时，Intellisense无法正常工作
                    ' # 表现为 鼠标停留到变量上时返回的是变量的默认值而不是实时值
                    ' # 自动窗口&局部变量窗口&即时窗口工作正常
                    ' ###########################################################
                    ' TODO something
                    Try
                        ' 适用于Try语句块内包含异步并行或者任务的场景。
                        ' 需要执行的异步操作
                        taskCompletion.Result.EnsureSuccessStatusCode()
                    Catch ex As AggregateException
                        For Each innerEx As Exception In ex.InnerExceptions
                            Logger.WriteLine(innerEx)
                        Next
                    Catch ex As Exception
                        Logger.WriteLine(ex)
                    End Try
                End Sub, TaskContinuationOptions.None)

            ' 等待异步操作完成
            Do
                Windows2.Delay(100)
            Loop Until tmpTask.IsCanceled Or tmpTask.IsCompleted Or tmpTask.IsFaulted
        End Sub

        ''' <summary>
        ''' 发送HEAD请求获取响应头
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns>成功 Success 返回True，其他情况都返回False；
        ''' 成功ResponseHeaders返回获取到的响应头，其他情况都返回具体错误信息；
        ''' HttpStatusCode返回相应的http请求状态码</returns>
        Private Shared Async Function InternalHeadAsync(ByVal httpClient As HttpClient, ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of HeadResponse)

            Dim success As Boolean
            Dim statusCode As HttpStatusCode
            Dim responseHeaders As Headers.HttpResponseHeaders
            Dim cts As CancellationTokenSource

            Try
                If Not s_IsInitialized Then
                    Throw New HttpAsyncUnInitializeException("未初始化，请先调用 'HttpAsync.Init()' 或者 'HttpAsync.ReInit()' 初始化")
                End If

                cts = New CancellationTokenSource
                cts.CancelAfter(s_DefaulTimeoutMilliseconds)
                Dim ct = cts.Token
                ct.ThrowIfCancellationRequested()

                ' 初始化uri类实例
                Dim baseAddress As New Uri(url)

                ' 构造请求消息体
                Dim requestMessage As New HttpRequestMessage(Http.HttpMethod.Head, baseAddress)
                requestMessage.SetRequestHeaders(requestHeaders, HttpMethod.HEAD)

                ' 发送请求
                Dim headResponse = Await httpClient.SendAsync(requestMessage, ct)

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
        Private Shared Async Function InternalGetResponseStringAsync(ByVal response As HttpResponseMessage) As Task(Of HttpResponse)
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
                    responseContent = Await response.Content.ReadAsStringAsync

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

        ''' <summary>
        ''' 获取网页源码
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns>成功 Success 返回True，其他情况都返回False；
        ''' 成功Message返回获取到的源码，其他情况都返回具体错误信息；
        ''' HttpStatusCode返回相应的http请求状态码</returns>
        Private Shared Async Function InternalGetAsync(ByVal httpClient As HttpClient, ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse
            Dim cts As CancellationTokenSource
            Try
                If Not s_IsInitialized Then
                    Throw New HttpAsyncUnInitializeException("未初始化，请先调用 'HttpAsync.Init()' 或者 'HttpAsync.ReInit()' 初始化")
                End If

                cts = New CancellationTokenSource
                cts.CancelAfter(s_DefaulTimeoutMilliseconds)
                Dim ct = cts.Token
                ct.ThrowIfCancellationRequested()

                ' 初始化uri类实例
                Dim baseAddress As New Uri(url)
                ' 构造请求消息体
                Dim requestMessage As New HttpRequestMessage(Http.HttpMethod.Get, baseAddress)
                requestMessage.SetRequestHeaders(requestHeaders, HttpMethod.GET)
                ct.ThrowIfCancellationRequested()
                ' 发送请求
                Dim response = Await httpClient.SendAsync(requestMessage, ct)
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
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入。如果传入有包含编码的数据，可能会导致乱码或者数据丢失</param>
        ''' <param name="postContentEncoding">请求主体的编码方式，必须跟抓包的一致，否则可能会导致乱码</param>
        ''' <returns></returns>
        Public Shared Async Function PostAsync(ByVal httpClient As HttpClient, ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String, ByVal postContentEncoding As Text.Encoding) As Task(Of HttpResponse)
            Return Await InternalPostAsync(httpClient, url, requestHeaders, postContent.ToKeyValuePairs(False), postContentEncoding)
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContentKvp">请求主体键值对集合，不需要编码，直接原字符串传入</param>
        ''' <param name="postContentEncoding">请求主体的编码方式，必须跟抓包的一致，否则可能会导致乱码</param>
        ''' <returns></returns>
        Public Shared Async Function InternalPostAsync(ByVal httpClient As HttpClient, ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContentKvp As IEnumerable(Of KeyValuePair(Of String, String)), ByVal postContentEncoding As Text.Encoding) As Task(Of HttpResponse)

            Dim httpResponse As HttpResponse
            Dim cts As CancellationTokenSource

            Try
                If Not s_IsInitialized Then
                    Throw New HttpAsyncUnInitializeException("未初始化，请先调用 'HttpAsync.Init()' 或者 'HttpAsync.ReInit()' 初始化")
                End If

                cts = New CancellationTokenSource
                cts.CancelAfter(s_DefaulTimeoutMilliseconds)
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
                Dim response = Await httpClient.SendAsync(requestMessage, ct)
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

        Private Shared Function GetHttpContent(ByVal requestHeaders As Dictionary(Of String, String), ByVal postContentKvp As IEnumerable(Of KeyValuePair(Of String, String)), ByVal postContentEncoding As Text.Encoding) As HttpContent
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
        Private Shared Function EncodeContent(ByVal postContentKvp As IEnumerable(Of KeyValuePair(Of String, String)), ByVal postContentEncoding As Text.Encoding) As String
            Dim sb = StringBuilderCache.AcquireSuper(361)
            For Each kvp As KeyValuePair(Of String, String) In postContentKvp
                If sb.Length > 0 Then
                    sb.Append("&"c)
                End If
                sb.Append(kvp.Key)
                sb.Append("="c)
                sb.Append(kvp.Value.UrlEncode(postContentEncoding))
            Next
            Dim postContent = StringBuilderCache.GetStringAndReleaseBuilderSuper(sb)
            Return postContent
        End Function

        ''' <summary>
        ''' 异步下载文件
        ''' </summary>
        ''' <param name="httpClient"></param>
        ''' <param name="progressMessageHandler"></param>
        ''' <param name="url">下载链接</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns></returns>
        Private Shared Async Function InternalDownloadFileAsync(ByVal httpClient As HttpClient, ByVal progressMessageHandler As ProgressMessageHandler, ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of DownloadResponse)

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
                AddHandler progressMessageHandler.HttpReceiveProgress, AddressOf DownloadProgressChangedEventHandler
                registDownloadProgressChangedEventHandler = True

                cts = New CancellationTokenSource
                cts.CancelAfter(Timeout.InfiniteTimeSpan)
                Dim ct = cts.Token

                ' 发送请求
                Dim response = Await httpClient.SendAsync(requestMessage, ct)
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
                    RemoveHandler progressMessageHandler.HttpReceiveProgress, AddressOf DownloadProgressChangedEventHandler
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
        Private Shared Async Function InternalDownloadFileAsync(ByVal httpClient As HttpClient, ByVal progressMessageHandler As ProgressMessageHandler, ByVal url As String, ByVal fileFullPath As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of HttpResponse)
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
            Dim responseContent = String.Empty
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
                AddHandler progressMessageHandler.HttpReceiveProgress, AddressOf DownloadProgressChangedEventHandler
                registDownloadProgressChangedEventHandler = True

                cts = New CancellationTokenSource
                cts.CancelAfter(Timeout.InfiniteTimeSpan)
                Dim ct = cts.Token

                ' 发送请求
                Dim response = Await httpClient.SendAsync(requestMessage)
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
                    RemoveHandler progressMessageHandler.HttpReceiveProgress, AddressOf DownloadProgressChangedEventHandler
                End If
                If cts IsNot Nothing Then
                    cts.Dispose()
                End If
            End Try

            Return New HttpResponse(success, statusCode, responseContent)
        End Function

        Private Shared Sub DownloadProgressChangedEventHandler(sender As Object, e As HttpProgressEventArgs)
            RaiseEvent DownloadProgressChanged(sender, New DownloadProgressChangedEventArgs(e.ProgressPercentage, e.UserState, e.BytesTransferred, e.TotalBytes))
        End Sub

        Private Shared Sub UploadProgressChangedEventHandler(sender As Object, e As HttpProgressEventArgs)
            RaiseEvent UploadProgressChanged(sender, New UploadProgressChangedEventArgs(e.ProgressPercentage, e.UserState, e.BytesTransferred, e.TotalBytes))
        End Sub

        ''' <summary>
        ''' 异步上传文件
        ''' </summary>
        ''' <param name="uploadInfo">上传文件需要的相关信息</param>
        ''' <param name="requestHeaders"></param>
        ''' <returns></returns>
        Private Shared Async Function InternalUploadFileAsync(ByVal httpClient As HttpClient, ByVal progressMessageHandler As ProgressMessageHandler, ByVal uploadInfo As UploadInfo, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of HttpResponse)
#Region "防呆检查"
            ' 判断传入的路径是否为文件（没有后缀名就视为非文件）
            Dim isFile = Path.GetExtension(uploadInfo.FileFullPath).Length > 0
            If Not isFile Then
                Return New HttpResponse(False, HttpStatusCode.BadRequest, "uploadInfo.FileFullPath is not a legal file")
            End If

            If Not s_IsInitialized Then
                Throw New HttpAsyncUnInitializeException("未初始化，请先调用 'HttpAsync.Init()' 或者 'HttpAsync.ReInit()' 初始化")
            End If
#End Region

            Dim success As Boolean
            Dim statusCode As HttpStatusCode
            Dim responseContent As String

            Dim fileName = Path.GetFileName(uploadInfo.FileFullPath)
            Dim mediaType = FileContentTypeHelper.GetMimeType(fileName)

            Using content = New MultipartFormDataContent()
                Dim fileContent = New ByteArrayContent(IO.File.ReadAllBytes(uploadInfo.FileFullPath))
                ' ContentDispositionHeaderValue 的 name 值必须要跟抓包的一致
                ' FileName 的值没有要求，可以随便写，建议写成文件的名称（带后缀）
                fileContent.Headers.ContentDisposition = uploadInfo.ContentDisposition

                ' ContentType 必须符合上传文件类型，否则会上传失败，返回 “只支持jpg或png格式的图片”
                fileContent.Headers.ContentType = New MediaTypeHeaderValue(mediaType)

                content.Add(fileContent)

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
                    AddHandler progressMessageHandler.HttpSendProgress, AddressOf UploadProgressChangedEventHandler
                    registUploadProgressChangedEventHandler = True


                    ' 开始上传数据
                    Dim response = Await httpClient.SendAsync(requestMessage)
                    responseContent = Await response.Content.ReadAsStringAsync()

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
                        RemoveHandler progressMessageHandler.HttpSendProgress, AddressOf UploadProgressChangedEventHandler
                    End If
                End Try

                Return New HttpResponse(success, statusCode, responseContent)
            End Using
        End Function
#End Region

#Region "共享函数区"
        ''' <summary>
        ''' Http请求公共设置，作用于整个类
        ''' </summary>
        Private Shared Sub HttpPublicConfig()
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

        Private Shared Function TrustAllValidationCallback(sender As Object, certificate As X509Certificate, chain As X509Chain, errors As SslPolicyErrors) As Boolean
            Return True
        End Function

        ''' <summary>
        ''' 用于重新初始化类，比如更换账号等涉及到重新传入cookie的情况
        ''' </summary>
        ''' <param name="cookies">如果为Nothing，则只重新初始化类，不设置cookie</param>
        Public Shared Sub Init(ByRef cookies As CookieContainer)
            Init(cookies, True)
        End Sub

        ''' <summary>
        ''' 用于重新初始化类，比如更换账号等涉及到重新传入cookie的情况
        ''' </summary>
        ''' <param name="cookies">如果为Nothing，则只重新初始化类，不设置cookie</param>
        Public Shared Sub ReInit(ByRef cookies As CookieContainer)
            ReInit(cookies, True)
        End Sub

        ''' <summary>
        ''' 用于重新初始化类，比如更换账号等涉及到重新传入cookie的情况
        ''' </summary>
        ''' <param name="cookies">如果为Nothing，则只重新初始化类，不设置cookie</param>
        ''' <param name="allowAutoRedirect">指示处理程序是否应跟随重定向响应,默认为True</param>
        Public Shared Sub Init(ByRef cookies As CookieContainer, ByVal allowAutoRedirect As Boolean)
            s_AllowAutoRedirect = allowAutoRedirect
            InitInternal(s_HttpClient, s_HttpClientHandler, s_ProgressMessageHandler, cookies, allowAutoRedirect, Nothing, Nothing)
            PreHeat(s_HttpClient)
        End Sub

        ''' <summary>
        ''' 用于重新初始化类，比如更换账号等涉及到重新传入cookie的情况
        ''' </summary>
        ''' <param name="cookies">如果为Nothing，则只重新初始化类，不设置cookie</param>
        ''' <param name="allowAutoRedirect">指示处理程序是否应跟随重定向响应,默认为True</param>
        ''' <param name="proxy">指示处理程序是否应设置代理信息,默认为Nothing</param>
        Public Shared Sub Init(ByRef cookies As CookieContainer, ByVal allowAutoRedirect As Boolean, ByRef proxy As IWebProxy)
            s_AllowAutoRedirect = allowAutoRedirect
            InitInternal(s_HttpClient, s_HttpClientHandler, s_ProgressMessageHandler, cookies, allowAutoRedirect, proxy, Nothing)
            PreHeat(s_HttpClient)
        End Sub

        ''' <summary>
        ''' 用于重新初始化类，比如更换账号等涉及到重新传入cookie的情况
        ''' </summary>
        ''' <param name="cookies">如果为Nothing，则只重新初始化类，不设置cookie</param>
        ''' <param name="allowAutoRedirect">指示处理程序是否应跟随重定向响应,默认为True</param>
        ''' <param name="proxy">指示处理程序是否应设置代理信息,默认为Nothing</param>
        ''' <param name="charSet">解码响应文本使用的字符集，设置错误会导致乱码。设置之前，确保访问的每个网页的字符集都是一样的，否则建议使用无参数的构造函数，程序内部自动检查字符集，当然也会牺牲一点效率。</param>
        Public Shared Sub Init(ByRef cookies As CookieContainer, ByVal allowAutoRedirect As Boolean, ByRef proxy As IWebProxy, ByVal charSet As String)
            s_AllowAutoRedirect = allowAutoRedirect
            InitInternal(s_HttpClient, s_HttpClientHandler, s_ProgressMessageHandler, cookies, allowAutoRedirect, proxy, charSet)
            PreHeat(s_HttpClient)
        End Sub

        ''' <summary>
        ''' 用于重新初始化类，比如更换账号等涉及到重新传入cookie的情况
        ''' </summary>
        ''' <param name="cookies">如果为Nothing，则只重新初始化类，不设置cookie</param>
        ''' <param name="allowAutoRedirect">指示处理程序是否应跟随重定向响应,默认为True</param>
        Public Shared Sub ReInit(ByRef cookies As CookieContainer, ByVal allowAutoRedirect As Boolean)
            s_AllowAutoRedirect = allowAutoRedirect
            InitInternal(s_HttpClient, s_HttpClientHandler, s_ProgressMessageHandler, cookies, allowAutoRedirect, Nothing, Nothing)
            PreHeat(s_HttpClient)
        End Sub

        ''' <summary>
        ''' 用于重新初始化类，比如更换账号等涉及到重新传入cookie的情况
        ''' </summary>
        ''' <param name="cookies">如果为Nothing，则只重新初始化类，不设置cookie</param>
        ''' <param name="allowAutoRedirect">指示处理程序是否应跟随重定向响应,默认为True</param>
        ''' <param name="proxy">指示处理程序是否应设置代理信息,默认为Nothing</param>
        Public Shared Sub ReInit(ByRef cookies As CookieContainer, ByVal allowAutoRedirect As Boolean, ByRef proxy As IWebProxy)
            s_AllowAutoRedirect = allowAutoRedirect
            InitInternal(s_HttpClient, s_HttpClientHandler, s_ProgressMessageHandler, cookies, allowAutoRedirect, proxy, Nothing)
            PreHeat(s_HttpClient)
        End Sub

        ''' <summary>
        ''' 用于重新初始化类，比如更换账号等涉及到重新传入cookie的情况
        ''' </summary>
        ''' <param name="cookies">如果为Nothing，则只重新初始化类，不设置cookie</param>
        ''' <param name="allowAutoRedirect">指示处理程序是否应跟随重定向响应,默认为True</param>
        ''' <param name="proxy">指示处理程序是否应设置代理信息,默认为Nothing</param>
        ''' <param name="charSet">解码响应文本使用的字符集，设置错误会导致乱码。设置之前，确保访问的每个网页的字符集都是一样的，否则建议使用无参数的构造函数，程序内部自动检查字符集，当然也会牺牲一点效率。</param>
        Private Shared Sub ReInit(ByRef httpClient As HttpClient, ByRef httpClientHandler As HttpClientHandler, ByRef processMessageHander As ProgressMessageHandler, ByRef cookies As CookieContainer, ByVal allowAutoRedirect As Boolean, ByRef proxy As IWebProxy, ByVal charSet As String)
            s_AllowAutoRedirect = allowAutoRedirect
            InitInternal(httpClient, httpClientHandler, processMessageHander, cookies, allowAutoRedirect, proxy, charSet)
            PreHeat(httpClient)
        End Sub

        ''' <summary>
        ''' 用于重新初始化类，比如更换账号等涉及到重新传入cookie的情况
        ''' </summary>
        ''' <param name="cookies">如果为Nothing，则只重新初始化类，不设置cookie</param>
        ''' <param name="allowAutoRedirect">指示处理程序是否应跟随重定向响应,默认为True</param>
        ''' <param name="proxy">指示处理程序是否应设置代理信息,默认为Nothing</param>
        ''' <param name="charSet">解码响应文本使用的字符集，设置错误会导致乱码。设置之前，确保访问的每个网页的字符集都是一样的，否则建议使用无参数的构造函数，程序内部自动检查字符集，当然也会牺牲一点效率。</param>
        Private Shared Sub InitInternal(ByRef httpClient As HttpClient, ByRef httpClientHandler As HttpClientHandler, ByRef processMessageHander As ProgressMessageHandler, ByRef cookies As CookieContainer, ByVal allowAutoRedirect As Boolean, ByRef proxy As IWebProxy, ByVal charSet As String)
            s_IsInitialized = False

            httpClientHandler = New GBKCompatibleHanlder(charSet) With {
                .UseProxy = If(proxy Is Nothing, False, True),
                .Proxy = proxy,
                .AutomaticDecompression = DecompressionMethods.GZip Or DecompressionMethods.Deflate,
                .AllowAutoRedirect = allowAutoRedirect
            }
            ' 把传入的cookie装盒并且设置到请求头
            ' cookie只能是这样设置.而且httpClient内部会自动管理cookies
            ' 不能多次设置，只能是设置一次
            If cookies IsNot Nothing Then
                httpClientHandler.CookieContainer = cookies
            End If

            processMessageHander = New ProgressMessageHandler(httpClientHandler)
            httpClient = New HttpClient(processMessageHander, False) With {
                .Timeout = TimeSpan.FromMilliseconds(s_DefaulTimeoutMilliseconds)
            }
            httpClient.DefaultRequestHeaders.ExpectContinue = False

            s_IsInitialized = True
        End Sub

        ''' <summary>
        ''' 设置请求头
        ''' </summary>
        ''' <param name="requestHeaders">请求头</param>
        ''' <param name="method">请求方法</param>
        Public Shared Sub SetRequestHeaders(ByRef requestHeaders As Dictionary(Of String, String), ByVal method As HttpMethod)
            s_HttpClient.SetRequestHeaders(requestHeaders, method)
        End Sub

        ''' <summary>
        ''' 发送HEAD请求获取响应头
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns>成功 Success 返回True，其他情况都返回False；
        ''' 成功ResponseHeaders返回获取到的响应头，其他情况都返回具体错误信息；
        ''' HttpStatusCode返回相应的http请求状态码</returns>
        Public Shared Async Function HeadAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of HeadResponse)
            Return Await InternalHeadAsync(s_HttpClient, url, requestHeaders)
        End Function

        ''' <summary>
        ''' 获取网页源码
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns>成功 Success 返回True，其他情况都返回False；
        ''' 成功Message返回获取到的源码，其他情况都返回具体错误信息；
        ''' HttpStatusCode返回相应的http请求状态码</returns>
        Public Shared Async Function GetAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of HttpResponse)
            Return Await InternalGetAsync(s_HttpClient, url, requestHeaders)
        End Function

        ''' <summary>
        ''' 获取网页源码
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <returns></returns>
        Public Shared Async Function GetAsync(ByVal url As String) As Task(Of HttpResponse)
            Return Await InternalGetAsync(s_HttpClient, url, Nothing)
        End Function

        ''' <summary>
        ''' 尝试获取网页源码直到成功，最多尝试三次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <returns></returns>
        Public Shared Async Function TryGetThreeTimeIfErrorAsync(ByVal url As String) As Task(Of HttpResponse)
            Return Await TryGetAsync(url, 3)
        End Function

        ''' <summary>
        ''' 尝试获取网页源码直到成功，最多尝试 <paramref name="tryTime"/> 次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Shared Async Function TryGetAsync(ByVal url As String, tryTime As Integer) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse

            Do
                httpResponse = Await InternalGetAsync(s_HttpClient, url, Nothing)
                tryTime -= 1
            Loop Until httpResponse.Success OrElse tryTime <= 0

            Return httpResponse
        End Function

        ''' <summary>
        ''' 尝试获取网页源码直到成功，最多尝试 <paramref name="tryTime"/> 次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Shared Async Function TryGetAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), tryTime As Integer) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse

            Do
                httpResponse = Await InternalGetAsync(s_HttpClient, url, requestHeaders)
                tryTime -= 1
            Loop Until httpResponse.Success OrElse tryTime <= 0

            Return httpResponse
        End Function

        ''' <summary>
        ''' 执行Get请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <returns></returns>
        Public Shared Async Function TryGetAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal retureIfContain As String, ByVal tryTime As Integer) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse

            Do
                httpResponse = Await TryGetAsync(url, requestHeaders, 1)
                tryTime -= 1
            Loop Until (httpResponse.Success AndAlso httpResponse.Message.IndexOf(retureIfContain) > -1) OrElse tryTime <= 0
            If httpResponse.Success Then
                httpResponse = New HttpResponse(httpResponse.Message.IndexOf(retureIfContain) > -1 OrElse tryTime > 0, httpResponse.StatusCode, httpResponse.Message)
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
        ''' <returns></returns>
        Public Shared Async Function PostAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String, ByVal postContentEncoding As Text.Encoding) As Task(Of HttpResponse)
            Return Await PostAsync(s_HttpClient, url, requestHeaders, postContent, postContentEncoding)
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入；内部默认使用UTF-8编码，如果抓包得到的不是UTF-8编码，请调用可以自定义编码方式的重载函数</param>
        ''' <returns></returns>
        Public Shared Async Function PostAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String) As Task(Of HttpResponse)
            Return Await PostAsync(s_HttpClient, url, requestHeaders, postContent, Text.Encoding.UTF8)
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入。如果传入有包含编码的数据，可能会导致乱码或者数据丢失</param>
        ''' <param name="postContentEncoding">>请求主体的编码方式，必须跟抓包的一致，否则可能会导致乱码</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param> 
        ''' <returns></returns>
        Public Shared Async Function TryPostAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String, ByVal postContentEncoding As Text.Encoding, tryTime As Integer) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse

            Do
                httpResponse = Await PostAsync(s_HttpClient, url, requestHeaders, postContent, postContentEncoding)
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
        ''' <returns></returns>
        Public Shared Async Function TryPostAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String, tryTime As Integer) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse

            Do

                httpResponse = Await PostAsync(s_HttpClient, url, requestHeaders, postContent, Text.Encoding.UTF8)
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
        ''' <returns></returns>
        Public Shared Async Function TryPostThreeTimeIfErrorAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String, ByVal postContentEncoding As Text.Encoding) As Task(Of HttpResponse)
            Return Await TryPostAsync(url, requestHeaders, postContent, postContentEncoding, 3)
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入；内部默认使用UTF-8编码，如果抓包得到的不是UTF-8编码，请调用可以自定义编码方式的重载函数</param>
        ''' <returns></returns>
        Public Shared Async Function TryPostThreeTimeIfErrorAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String) As Task(Of HttpResponse)
            Return Await TryPostAsync(url, requestHeaders, postContent, 3)
        End Function

        ''' <summary>
        ''' 异步下载文件
        ''' </summary>
        ''' <param name="url">下载链接</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns></returns>
        Public Shared Async Function DownloadFileAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of DownloadResponse)
            Return Await InternalDownloadFileAsync(s_HttpClient, s_ProgressMessageHandler, url, requestHeaders)
        End Function

        ''' <summary>
        ''' 尝试下载文件直到成功，最多尝试 <paramref name="tryTime"/> 次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Shared Async Function TryDownloadFileAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), tryTime As Integer) As Task(Of DownloadResponse)
            Dim httpResponse As DownloadResponse

            Do
                httpResponse = Await InternalDownloadFileAsync(s_HttpClient, s_ProgressMessageHandler, url, requestHeaders)
                tryTime -= 1
            Loop Until httpResponse.Success OrElse tryTime <= 0

            Return httpResponse
        End Function

        ''' <summary>
        ''' 尝试下载文件直到成功，最多尝试三次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <returns></returns>
        Public Shared Async Function TryDownloadFileThreeTimeIfErrorAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of DownloadResponse)
            Return Await TryDownloadFileAsync(url, requestHeaders, 3)
        End Function


        '        ''' <summary>
        '        ''' 异步下载文件 （旧版通过自己计算的下载反馈进度 20171002）
        '        ''' </summary>
        '        ''' <param name="url">下载链接</param>
        '        ''' <param name="fileFullPath">文件存储路径</param>
        '        ''' <param name="requestHeaders">请求头</param>
        '        ''' <returns></returns>
        '        Public Shared Async Function DownloadFileAsync(ByVal url As String, ByVal fileFullPath As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of HttpResponse)
        '#Region "防呆检查"
        '            ' 判断传入的路径是否为文件（没有后缀名就视为非文件）
        '            Dim isFile = System.ShanXingTech.Path.GetExtension(fileFullPath).Length > 0
        '            If Not isFile Then
        '                Return (False, HttpStatusCode.BadRequest, "FileFullPath isnot a file")
        '            End If

        '            ' 判断文件夹是否存在
        '            Dim isDirectoryExists = System.ShanXingTech.Directory.Exists(System.ShanXingTech.Path.GetDirectoryName(fileFullPath))
        '            If Not isDirectoryExists Then
        '                Return (False, HttpStatusCode.BadRequest, "Directory not found of fileFullPath")
        '            End If
        '#End Region

        '            Dim success As Boolean
        '            Dim statusCode As HttpStatusCode
        '            Dim responseContent As String

        '            'Dim source As New CancellationTokenSource()
        '            'Dim token As CancellationToken = source.Token
        '            ''source.CancelAfter(30000)

        '            ' 设置请求头
        '            s_HttpClient.SetRequestHeaders(requestHeaders)

        '            Try
        '                ' 初始化uri类实例
        '                Dim baseAddress As New Uri(url)
        '#Region "先发送HEAD请求获取文件长度"
        '                Dim httpRequestMessage As New HttpRequestMessage With {
        '                .RequestUri = baseAddress,
        '                .Method = New Http.HttpMethod("HEAD")
        '                }
        '                Dim headResponse = Await s_HttpClient.SendAsync(httpRequestMessage)
        '                Dim totalLength As Long?
        '                If headResponse.StatusCode = HttpStatusCode.OK Then
        '                    totalLength = headResponse.Content.Headers.ContentLength
        '                Else
        '                    Return (False, headResponse.StatusCode, headResponse.ReasonPhrase)
        '                End If
        '#End Region

        '#Region "下载文件主体"
        '                ' 发送请求
        '                Using responseStream = Await s_HttpClient.GetStreamAsync(baseAddress)
        '                    Dim bufferSize = 80 * 1024
        '                    Dim buffer As Byte() = New Byte(bufferSize - 1) {}

        '                    Dim totalBytesReceived As Long

        '                    Using fileStream As New System.ShanXingTech.FileStream(fileFullPath,
        '                                                          System.ShanXingTech.FileMode.Create,
        '                                                          System.ShanXingTech.FileAccess.ReadWrite,
        '                                                          System.ShanXingTech.FileShare.ReadWrite,
        '                                                          bufferSize,
        '                                                          True)
        '                        ' 可以用 CopyToAsync 方法一步写文件，但是不能做进度控制，没法向用户反馈
        '                        ' 所以还是用下面的方法
        '                        ' 关于效率：测试的时候 两种方法效率很接近(只测试过几最大M的文件)
        '                        Dim lengthRead As Integer
        '                        Do
        '                            Await fileStream.WriteAsync(buffer, 0, lengthRead)
        '                            lengthRead = Await responseStream.ReadAsync(buffer, 0, bufferSize)

        '                            ' 报告进度
        '                            totalBytesReceived += lengthRead
        '                            RaiseEvent DownloadProgressChanged(s_HttpClient, New DownloadProgressChangedEventArgs(CInt(totalBytesReceived / totalLength * 100), fileFullPath))
        '                        Loop While lengthRead <> 0

        '                        ' 报告下载完成
        '                        RaiseEvent DownloadFileCompleted(s_HttpClient, New DownloadFileCompletedEventArgs(Nothing, False, Nothing))
        '                    End Using
        '                End Using
        '#End Region

        '                Success = True
        '                statusCode = HttpStatusCode.OK
        '            Catch ex As UriFormatException
        '                statusCode = HttpStatusCode.BadRequest
        '                responseContent = ex.Message

        '                ' 报告下载取消
        '                RaiseEvent DownloadFileCompleted(s_HttpClient, New DownloadFileCompletedEventArgs(ex, True, Nothing))
        '            Catch ex As TaskCanceledException
        '                statusCode = HttpStatusCode.RequestTimeout
        '                responseContent = ex.Message

        '                ' 报告下载取消
        '                RaiseEvent DownloadFileCompleted(s_HttpClient, New DownloadFileCompletedEventArgs(ex, True, Nothing))

        '                Logger.WriteLine(ex)
        '            Catch ex As Exception
        '                responseContent = ex.Message

        '                ' 报告下载取消
        '                RaiseEvent DownloadFileCompleted(s_HttpClient, New DownloadFileCompletedEventArgs(ex, True, Nothing))

        '                If ex.Message.IndexOf("404") = -1 Then
        '                    Logger.WriteLine(ex)
        '                Else
        '                    statusCode = HttpStatusCode.NotFound
        '                End If
        '            End Try

        '            Return (Success, statusCode, responseContent)
        '        End Function

        ''' <summary>
        ''' 异步下载文件
        ''' </summary>
        ''' <param name="url">下载链接</param>
        ''' <param name="fileFullPath">文件存储路径</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns></returns>
        Public Shared Async Function DownloadFileAsync(ByVal url As String, ByVal fileFullPath As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of HttpResponse)
            Return Await InternalDownloadFileAsync(s_HttpClient, s_ProgressMessageHandler, url, fileFullPath, requestHeaders)
        End Function

        ''' <summary>
        ''' 尝试下载文件直到成功，最多尝试 <paramref name="tryTime"/> 次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Shared Async Function TryDownloadFileAsync(ByVal url As String, ByVal fileFullPath As String, ByVal requestHeaders As Dictionary(Of String, String), tryTime As Integer) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse

            Do
                httpResponse = Await InternalDownloadFileAsync(s_HttpClient, s_ProgressMessageHandler, url, fileFullPath, requestHeaders)
                tryTime -= 1
            Loop Until httpResponse.Success OrElse tryTime <= 0

            Return httpResponse
        End Function

        ''' <summary>
        ''' 尝试下载文件直到成功，最多尝试三次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <returns></returns>
        Public Shared Async Function TryDownloadFileThreeTimeIfErrorAsync(ByVal url As String, ByVal fileFullPath As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of HttpResponse)
            Return Await TryDownloadFileAsync(url, fileFullPath, requestHeaders, 3)
        End Function

        ''' <summary>
        ''' 异步上传文件
        ''' </summary>
        ''' <param name="uploadInfo">上传文件需要的相关信息</param>
        ''' <param name="requestHeaders"></param>
        ''' <returns></returns>
        Public Shared Async Function UploadFileAsync(ByVal uploadInfo As UploadInfo, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of HttpResponse)
            Return Await InternalUploadFileAsync(s_HttpClient, s_ProgressMessageHandler, uploadInfo, requestHeaders)
        End Function


        ''' <summary>
        ''' 尝试上传文件直到成功，最多尝试 <paramref name="tryTime"/> 次
        ''' </summary>
        ''' <param name="uploadInfo">上传文件需要的相关信息</param>
        ''' <param name="requestHeaders"></param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Shared Async Function TryUploadFileAsync(ByVal uploadInfo As UploadInfo, ByVal requestHeaders As Dictionary(Of String, String), tryTime As Integer) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse

            Do
                httpResponse = Await InternalUploadFileAsync(s_HttpClient, s_ProgressMessageHandler, uploadInfo, requestHeaders)
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
        Public Shared Async Function TryUploadFileThreeTimeIfErrorAsync(ByVal uploadInfo As UploadInfo, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of HttpResponse)
            Return Await TryUploadFileAsync(uploadInfo, requestHeaders, 3)
        End Function
#End Region


#Region "实例函数区"
        Public Shared Sub SetDefaultCharSet(ByVal charSetName As String)
            HttpAsync.DefaultCharSet = charSetName
        End Sub

        ''' <summary>
        ''' 设置请求头
        ''' </summary>
        ''' <param name="requestHeaders">请求头</param>
        ''' <param name="method">请求方法</param>
        ''' <param name="noUse">此参数无实际意义，仅用作占位</param>
        Public Sub SetRequestHeaders(ByRef requestHeaders As Dictionary(Of String, String), ByVal method As HttpMethod, ByVal noUse As Integer)
            m_HttpClient.SetRequestHeaders(requestHeaders, method)
        End Sub

        ''' <summary>
        ''' 发送HEAD请求获取响应头 
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="requestHeaders"></param>
        ''' <param name="noUse">因为vb.net暂时不能实现实例方法跟共享方法同名同参数同返回值，所以只能用一个参数来占位以表示不同</param>
        ''' <returns>成功 Success 返回True，其他情况都返回False；
        ''' 成功ResponseHeaders返回获取到的响应头，其他情况都返回具体错误信息；
        ''' HttpStatusCode返回相应的http请求状态码</returns>
        Public Async Function HeadAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal noUse As Integer) As Task(Of HeadResponse)
            Return Await InternalHeadAsync(m_HttpClient, url, requestHeaders)
        End Function

        ''' <summary>
        ''' 获取网页源码
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns>成功 Success 返回True，其他情况都返回False；
        ''' 成功Message返回获取到的源码，其他情况都返回具体错误信息；
        ''' HttpStatusCode返回相应的http请求状态码</returns>
        Public Async Function GetAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal noUse As Integer) As Task(Of HttpResponse)

            Return Await InternalGetAsync(m_HttpClient, url, requestHeaders)
        End Function

        ''' <summary>
        ''' 获取网页源码
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <returns></returns>
        Public Async Function GetAsync(ByVal url As String, ByVal noUse As Integer) As Task(Of HttpResponse)

            Return Await InternalGetAsync(m_HttpClient, url, Nothing)
        End Function

        ''' <summary>
        ''' 尝试获取网页源码直到成功，最多尝试三次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <returns></returns>
        Public Async Function TryGetThreeTimeIfErrorAsync(ByVal url As String, ByVal noUse As Integer) As Task(Of HttpResponse)
            Return Await TryGetAsync(url, 3, 0)
        End Function

        ''' <summary>
        ''' 尝试获取网页源码直到成功，最多尝试 <paramref name="tryTime"/> 次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Async Function TryGetAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), tryTime As Integer, ByVal noUse As Integer) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse

            Do

                httpResponse = Await InternalGetAsync(m_HttpClient, url, requestHeaders)
                tryTime -= 1
            Loop Until httpResponse.Success OrElse tryTime <= 0

            Return httpResponse
        End Function

        ''' <summary>
        ''' 尝试获取网页源码直到成功，最多尝试三次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <returns></returns>
        Public Async Function TryGetAsync(ByVal url As String, ByVal tryTime As Integer, ByVal noUse As Integer) As Task(Of HttpResponse)
            Dim header As Dictionary(Of String, String)
            Return Await TryGetAsync(url, header, tryTime, 0)
        End Function

        ''' <summary>
        ''' 执行Get请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <returns></returns>
        Public Async Function TryGetAsync(ByVal url As String, ByVal retureIfContain As String, ByVal tryTime As Integer, ByVal noUse As Integer) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse

            Do
                httpResponse = Await HttpAsync.TryGetAsync(url, 3)
                tryTime -= 1
            Loop Until (httpResponse.Success AndAlso httpResponse.Message.IndexOf(retureIfContain) > -1) OrElse tryTime <= 0
            If httpResponse.Success Then
                httpResponse = New HttpResponse(httpResponse.Message.IndexOf(retureIfContain) > -1 OrElse tryTime > 0, httpResponse.StatusCode, httpResponse.Message)
            End If

            Return httpResponse
        End Function

        ''' <summary>
        ''' 执行Get请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <returns></returns>
        Public Async Function TryGetAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal referer As String, ByVal retureIfContain As String, ByVal tryTime As Integer, ByVal noUse As Integer) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse

            Do
                httpResponse = Await TryGetAsync(url, requestHeaders, 1)
                tryTime -= 1
            Loop Until (httpResponse.Success AndAlso httpResponse.Message.IndexOf(retureIfContain) > -1) OrElse tryTime <= 0
            If httpResponse.Success Then
                httpResponse = New HttpResponse(httpResponse.Message.IndexOf(retureIfContain) > -1 OrElse tryTime > 0, httpResponse.StatusCode, httpResponse.Message)
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
        ''' <returns></returns>
        Public Async Function PostAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String, ByVal postContentEncoding As Text.Encoding, ByVal noUse As Integer) As Task(Of HttpResponse)
            Return Await PostAsync(m_HttpClient, url, requestHeaders, postContent, postContentEncoding)
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入；内部默认使用UTF-8编码，如果抓包得到的不是UTF-8编码，请调用可以自定义编码方式的重载函数</param>
        ''' <returns></returns>
        Public Async Function PostAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String, ByVal noUse As Integer) As Task(Of HttpResponse)
            Return Await PostAsync(m_HttpClient, url, requestHeaders, postContent, Text.Encoding.UTF8)
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入。如果传入有包含编码的数据，可能会导致乱码或者数据丢失</param>
        ''' <param name="postContentEncoding">请求主体的编码方式，必须跟抓包的一致，否则可能会导致乱码</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param> 
        ''' <returns></returns>
        Public Async Function TryPostAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String, ByVal postContentEncoding As Text.Encoding, tryTime As Integer, ByVal noUse As Integer) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse

            Do
                httpResponse = Await PostAsync(m_HttpClient, url, requestHeaders, postContent, postContentEncoding)
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
        ''' <returns></returns>
        Public Async Function TryPostAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String, tryTime As Integer, ByVal noUse As Integer) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse

            Do
                httpResponse = Await PostAsync(m_HttpClient, url, requestHeaders, postContent, Text.Encoding.UTF8)
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
        ''' <returns></returns>
        Public Async Function TryPostThreeTimeIfErrorAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String, ByVal postContentEncoding As Text.Encoding, ByVal noUse As Integer) As Task(Of HttpResponse)
            Return Await TryPostAsync(url, requestHeaders, postContent, postContentEncoding, 3, 0)
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入；内部默认使用UTF-8编码，如果抓包得到的不是UTF-8编码，请调用可以自定义编码方式的重载函数</param>
        ''' <returns></returns>
        Public Async Function TryPostThreeTimeIfErrorAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String, ByVal noUse As Integer) As Task(Of HttpResponse)
            Return Await TryPostAsync(url, requestHeaders, postContent, 3, 0)
        End Function

        ''' <summary>
        ''' 异步下载文件
        ''' </summary>
        ''' <param name="url">下载链接</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns></returns>
        Public Async Function DownloadFileAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal noUse As Integer) As Task(Of DownloadResponse)
            Return Await InternalDownloadFileAsync(m_HttpClient, m_ProgressMessageHandler, url, requestHeaders)
        End Function

        ''' <summary>
        ''' 尝试下载文件直到成功，最多尝试 <paramref name="tryTime"/> 次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Async Function TryDownloadFileAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), tryTime As Integer, ByVal noUse As Integer) As Task(Of DownloadResponse)
            Dim httpResponse As DownloadResponse

            Do
                httpResponse = Await InternalDownloadFileAsync(m_HttpClient, m_ProgressMessageHandler, url, requestHeaders)
                tryTime -= 1
            Loop Until httpResponse.Success OrElse tryTime <= 0

            Return httpResponse
        End Function

        ''' <summary>
        ''' 尝试下载文件直到成功，最多尝试三次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <returns></returns>
        Public Async Function TryDownloadFileThreeTimeIfErrorAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal noUse As Integer) As Task(Of DownloadResponse)
            Return Await TryDownloadFileAsync(url, requestHeaders, 3, 0)
        End Function

        ''' <summary>
        ''' 异步下载文件
        ''' </summary>
        ''' <param name="url">下载链接</param>
        ''' <param name="fileFullPath">文件存储路径</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns></returns>
        Public Async Function DownloadFileAsync(ByVal url As String, ByVal fileFullPath As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal noUse As Integer) As Task(Of HttpResponse)
            Return Await InternalDownloadFileAsync(m_HttpClient, m_ProgressMessageHandler, url, fileFullPath, requestHeaders)
        End Function

        ''' <summary>
        ''' 尝试下载文件直到成功，最多尝试 <paramref name="tryTime"/> 次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Async Function TryDownloadFileAsync(ByVal url As String, ByVal fileFullPath As String, ByVal requestHeaders As Dictionary(Of String, String), tryTime As Integer, ByVal noUse As Integer) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse

            Do
                httpResponse = Await InternalDownloadFileAsync(m_HttpClient, m_ProgressMessageHandler, url, fileFullPath, requestHeaders)
                tryTime -= 1
            Loop Until httpResponse.Success OrElse tryTime <= 0

            Return httpResponse
        End Function

        ''' <summary>
        ''' 尝试下载文件直到成功，最多尝试三次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <returns></returns>
        Public Async Function TryDownloadFileThreeTimeIfErrorAsync(ByVal url As String, ByVal fileFullPath As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal noUse As Integer) As Task(Of HttpResponse)
            Return Await TryDownloadFileAsync(url, fileFullPath, requestHeaders, 3, 0)
        End Function

        ''' <summary>
        ''' 异步上传文件
        ''' </summary>
        ''' <param name="uploadInfo">上传文件需要的相关信息</param>
        ''' <param name="requestHeaders"></param>
        ''' <returns></returns>
        Public Async Function UploadFileAsync(ByVal uploadInfo As UploadInfo, ByVal requestHeaders As Dictionary(Of String, String), ByVal noUse As Integer) As Task(Of HttpResponse)
            Return Await InternalUploadFileAsync(m_HttpClient, m_ProgressMessageHandler, uploadInfo, requestHeaders)
        End Function


        ''' <summary>
        ''' 尝试上传文件直到成功，最多尝试 <paramref name="tryTime"/> 次
        ''' </summary>
        ''' <param name="uploadInfo">上传文件需要的相关信息</param>
        ''' <param name="requestHeaders"></param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Async Function TryUploadFileAsync(ByVal uploadInfo As UploadInfo, ByVal requestHeaders As Dictionary(Of String, String), tryTime As Integer, ByVal noUse As Integer) As Task(Of HttpResponse)
            Dim httpResponse As HttpResponse

            Do
                httpResponse = Await InternalUploadFileAsync(m_HttpClient, m_ProgressMessageHandler, uploadInfo, requestHeaders)
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
        Public Async Function TryUploadFileThreeTimeIfErrorAsync(ByVal uploadInfo As UploadInfo, ByVal requestHeaders As Dictionary(Of String, String), ByVal noUse As Integer) As Task(Of HttpResponse)
            Return Await TryUploadFileAsync(uploadInfo, requestHeaders, 3, 0)
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

            End If

            ' TODO: 释放未托管资源(未托管对象)并在以下内容中替代 Finalize()。
            ' TODO: 将大型字段设置为 null。

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

            If s_ProgressMessageHandler IsNot Nothing Then
                s_ProgressMessageHandler.Dispose()
                s_ProgressMessageHandler = Nothing
            End If

            If s_HttpClientHandler IsNot Nothing Then
                s_HttpClientHandler.Dispose()
                s_HttpClientHandler = Nothing
            End If

            If s_HttpClient IsNot Nothing Then
                s_HttpClient.Dispose()
                s_HttpClient = Nothing
            End If
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
