Imports System.ComponentModel
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Handlers
Imports System.Net.Http.Headers
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports ShanXingTech
Imports ShanXingTech.IO2
Imports ShanXingTech.Text2

Namespace ShanXingTech.Net2
    Public Class HttpAsync
        Implements IDisposable

#Region "内部类"
        ''' <summary>
        ''' 下载进度改变事件类
        ''' </summary>
        Public Class DownloadProgressChangedEventArgs
            Inherits HttpProgressEventArgs

            Public Sub New(progressPercentage As Integer, userToken As Object, bytesTransferred As Long, totalBytes As Long?)
                MyBase.New(progressPercentage, userToken, bytesTransferred, totalBytes)
            End Sub
        End Class

        ''' <summary>
        ''' 下载完成事件类
        ''' </summary>
        Public Class DownloadFileCompletedEventArgs
            Inherits AsyncCompletedEventArgs

            Public Sub New([error] As Exception, cancelled As Boolean, userState As Object)
                MyBase.New([error], cancelled, userState)
            End Sub
        End Class

        ''' <summary>
        ''' 上传进度改变事件类
        ''' </summary>
        Public Class UploadProgressChangedEventArgs
            Inherits HttpProgressEventArgs

            Public Sub New(progressPercentage As Integer, userToken As Object, bytesTransferred As Long, totalBytes As Long?)
                MyBase.New(progressPercentage, userToken, bytesTransferred, totalBytes)
            End Sub
        End Class


        ''' <summary>
        ''' 上传完成事件类
        ''' </summary>
        Public Class UploadFileCompletedEventArgs
            Inherits AsyncCompletedEventArgs

            Public Sub New([error] As Exception, cancelled As Boolean, userState As Object)
                MyBase.New([error], cancelled, userState)
            End Sub
        End Class

        ''' <summary>
        ''' 兼容GBK编码网页，代替默认的HttpClientHandler
        ''' </summary>
        Private Class GBKCompatibleHanlder
            Inherits HttpClientHandler
#Region "字段区"
            Private m_CharSet As String
#End Region

            Public Sub New()

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
                    contentType.CharSet = GetCharSet(contentType)
                End If

                Return response
            End Function

            ''' <summary>
            ''' 获取网页的CharSet；因为是从网页源码中获取CharSet，比较低效，所有要求高效的时候，尽量不要使用
            ''' </summary>
            ''' <param name="contentType"></param>
            ''' <returns></returns>
            Private Function GetCharSet(ByVal contentType As MediaTypeHeaderValue) As String
                ' 获取顺序 调用者设置>meta标签>内部默认（DefaultCharSet）
                If Not m_CharSet.IsNullOrEmpty Then
                    Return m_CharSet
                End If

                'Dim response = Await HttpContent.ReadAsStringAsync
                'Dim match = Regex.Match(response, "<meta.*?charset=""?(\w+-?\w+)""?", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
                'Dim charSet = match.Groups(1).Value

                '' 如果没法从返回到的文本中获取编码方式，那就尝试根据 MediaType 决定 charset 20180627
                'If charSet.Length = 0 AndAlso
                '    (String.Equals("application/json", HttpContent.Headers.ContentType.MediaType, StringComparison.OrdinalIgnoreCase) OrElse
                '    String.Equals("text/html", HttpContent.Headers.ContentType.MediaType, StringComparison.OrdinalIgnoreCase)) Then
                '    charSet = "utf-8"
                'End If
                Dim charSet As String = Nothing
                If String.Equals("application/json", contentType.MediaType, StringComparison.OrdinalIgnoreCase) OrElse
                        String.Equals("text/html", contentType.MediaType, StringComparison.OrdinalIgnoreCase) Then
                    charSet = "utf-8"
                End If

                ' 如果获取到的charSet为空的话，那就设置为默认的charSet
                Return If(charSet, DefaultCharSet)
            End Function
        End Class
#End Region

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
        Private Property instHttpClient As HttpClient
        Private Property instHttpClientHandler As HttpClientHandler
        Private Property instProgressMessageHandler As ProgressMessageHandler
#End Region

#Region "共享属性区"
        Private Shared Property s_HttpClient As HttpClient
        Private Shared Property s_HttpClientHandler As HttpClientHandler
        Private Shared Property s_ProgressMessageHandler As ProgressMessageHandler

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
        ''' <summary>
        ''' 默认的超时时间
        ''' </summary>
        ''' <returns></returns>
        Public Shared Property DefaulTimeoutMilliseconds As Integer

#End Region

#Region "字段区"
        <ThreadStatic>
        Private Shared m_CancellationTokenSource As CancellationTokenSource
        <ThreadStatic>
        Private Shared m_CancellationToken As CancellationToken
#End Region

#Region "构造函数区"
        ''' <summary>
        ''' 类构造函数
        ''' 类之内的任意一个静态方法第一次调用时调用此构造函数
        ''' 而且程序生命周期内仅调用一次
        ''' </summary>
        Shared Sub New()
            HttpPublicConfig()

            InitInternal(s_HttpClient, s_HttpClientHandler, s_ProgressMessageHandler)

            s_IsInitialized = True

            DefaultCharSet = "GBK"
            DefaulTimeoutMilliseconds = 10000
        End Sub

        ''' <summary>
        ''' 不传入cookies，自动处理重定向
        ''' </summary>
        Sub New()
            HttpPublicConfig()

            InitInternal(instHttpClient, instHttpClientHandler, instProgressMessageHandler)

            s_IsInitialized = True
            DefaulTimeoutMilliseconds = 10000
        End Sub

        ''' <summary>
        ''' 传入cookies，自动处理重定向
        ''' </summary>
        ''' <param name="cookies"></param>

        Sub New(ByRef cookies As CookieContainer)
            HttpPublicConfig()

            InitInternal(instHttpClient, instHttpClientHandler, instProgressMessageHandler, True)
            InitInternal(instHttpClientHandler, cookies)
        End Sub

        ''' <summary>
        ''' 不传入cookies，重定向由参数 <paramref name="allowAutoRedirect"/> 决定
        ''' </summary>
        ''' <param name="allowAutoRedirect">指示处理程序是否应跟随重定向响应,默认为True</param>
        Sub New(ByVal allowAutoRedirect As Boolean)
            HttpPublicConfig()

            InitInternal(instHttpClient, instHttpClientHandler, instProgressMessageHandler, allowAutoRedirect)
            InitInternal(instHttpClientHandler, Nothing)
        End Sub

        ''' <summary>
        ''' 传入cookies，重定向由参数 <paramref name="allowAutoRedirect"/> 决定
        ''' </summary>
        ''' <param name="cookies"></param>
        ''' <param name="allowAutoRedirect">指示处理程序是否应跟随重定向响应,默认为True</param>
        Sub New(ByRef cookies As CookieContainer, ByVal allowAutoRedirect As Boolean)
            HttpPublicConfig()

            InitInternal(instHttpClient, instHttpClientHandler, instProgressMessageHandler, allowAutoRedirect)
            InitInternal(instHttpClientHandler, cookies)
        End Sub
#End Region

#Region "共享函数 实例函数共用，这个区块内不应该使用共享字段、属性等"
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="httpClient"></param>
        ''' <param name="httpClientHandler"></param>
        ''' <param name="processMessageHander"></param>
        ''' <param name="allowAutoRedirect">指示处理程序是否应跟随重定向响应,默认为True</param>
        Private Shared Sub InitInternal(ByRef httpClient As HttpClient, ByRef httpClientHandler As HttpClientHandler, ByRef processMessageHander As ProgressMessageHandler, ByVal allowAutoRedirect As Boolean)
            httpClient = New HttpClient() With {.Timeout = TimeSpan.FromMilliseconds(1618)}

            httpClient.DefaultRequestHeaders.ExpectContinue = False

            ' 预热
            Dim httpRequestMessage As New HttpRequestMessage With {
                .RequestUri = New Uri("https://t.alicdn.com/t/gettime"),
                .Method = New Http.HttpMethod("HEAD")}

            Try
                ' 适用于Try语句块内包含异步并行或者任务的场景。
                ' 需要执行的异步操作
                httpClient.SendAsync(httpRequestMessage).Result.EnsureSuccessStatusCode()
            Catch ex As AggregateException
                For Each innerEx As Exception In ex.InnerExceptions
                    Logger.WriteLine(innerEx)
                Next
            Catch ex As HttpRequestException
                '
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try

            ' 预热完再重新实例化会有效嘛？
            httpClientHandler = New GBKCompatibleHanlder With {
                .UseProxy = False,
                .Proxy = Nothing,
                .AutomaticDecompression = DecompressionMethods.GZip Or DecompressionMethods.Deflate,
                .AllowAutoRedirect = allowAutoRedirect
            }

            processMessageHander = New ProgressMessageHandler(httpClientHandler)

            httpClient = New HttpClient(processMessageHander, False) With {
                .Timeout = TimeSpan.FromSeconds(10)
            }
            httpClient.DefaultRequestHeaders.ExpectContinue = False
        End Sub

        Private Shared Sub InitInternal(ByRef httpClient As HttpClient, ByRef httpClientHandler As HttpClientHandler, ByRef processMessageHander As ProgressMessageHandler)
            InitInternal(httpClient, httpClientHandler, processMessageHander, True)
        End Sub

        ''' <summary>
        ''' 用于初始化类，如果操作需要带上cookies的话一定要在此函数传入；如果不需要传入cookies，直接调用相应函数即可
        ''' </summary>
        ''' <param name="cookies"></param>
        Private Shared Sub InitInternal(ByRef httpClientHandler As HttpClientHandler， ByRef cookies As CookieContainer)
            ' 把传入的cookie装盒并且设置到请求头
            ' cookie只能是这样设置.而且s_HttpClientHandler会内部会自动管理cookies
            ' 不能多次设置，只能是设置一次
            If cookies Is Nothing OrElse httpClientHandler.CookieContainer.Count > 0 Then Return
            httpClientHandler.CookieContainer = cookies

            s_IsInitialized = True
        End Sub

        ''' <summary>
        ''' 发送HEAD请求获取响应头
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns>成功 Success 返回True，其他情况都返回False；
        ''' 成功ResponseHeaders返回获取到的响应头，其他情况都返回具体错误信息；
        ''' HttpStatusCode返回相应的http请求状态码</returns>
        Private Shared Async Function InternalHeadAsync(ByVal httpClient As HttpClient, ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal cancellationToken As CancellationToken) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, ResponseHeaders As Headers.HttpResponseHeaders))

            Dim success As Boolean
            Dim statusCode As HttpStatusCode
            Dim responseHeaders As Headers.HttpResponseHeaders

            Try
                cancellationToken.ThrowIfCancellationRequested()

                ' 初始化uri类实例
                Dim baseAddress As New Uri(url)

                ' 构造请求消息体
                Dim requestMessage As New HttpRequestMessage(Http.HttpMethod.Head, baseAddress)
                requestMessage.SetRequestHeaders(requestHeaders, HttpMethod.HEAD)

                ' 发送请求
                Dim headResponse = Await httpClient.SendAsync(requestMessage, cancellationToken)
                If headResponse.StatusCode = HttpStatusCode.OK Then
                    responseHeaders = headResponse.Headers
                    statusCode = HttpStatusCode.OK
                    success = True
                Else
                    statusCode = headResponse.StatusCode
                    Return (False, headResponse.StatusCode, headResponse.Headers)
                End If
            Catch ex As UriFormatException
                statusCode = HttpStatusCode.BadRequest
            Catch ex As TaskCanceledException
                statusCode = HttpStatusCode.RequestTimeout
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try

            Return (success, statusCode, responseHeaders)
        End Function
        ''' <summary>
        ''' 获取请求返回的信息
        ''' </summary>
        ''' <param name="response"></param>
        ''' <returns></returns>
        Private Shared Async Function InternalGetResponseStringAsync(ByVal response As HttpResponseMessage) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Dim success As Boolean
            Dim statusCode As HttpStatusCode
            Dim responseContent As String

            Try
                ' 检查请求状态
                If response.IsSuccessStatusCode Then
                    ' 获取响应文本
                    responseContent = Await response.Content.ReadAsStringAsync

                    statusCode = HttpStatusCode.OK
                    success = True
                ElseIf response.StatusCode = HttpStatusCode.Found Then
                    responseContent = response.Headers.ToString

                    ' 如果是302的话 responseContent 返回头信息
                    statusCode = response.StatusCode
                    success = True
                Else
                    statusCode = response.StatusCode
                    responseContent = response.ReasonPhrase
                End If
            Catch ex As TaskCanceledException
                statusCode = HttpStatusCode.RequestTimeout
                responseContent = ex.Message
            Catch ex As Exception
                statusCode = response.StatusCode
                responseContent = ex.Message

                Logger.WriteLine(ex)
            End Try

            Return (success, statusCode, responseContent)
        End Function

        ''' <summary>
        ''' 获取网页源码
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns>成功 Success 返回True，其他情况都返回False；
        ''' 成功Message返回获取到的源码，其他情况都返回具体错误信息；
        ''' HttpStatusCode返回相应的http请求状态码</returns>
        Private Shared Async Function InternalGetAsync(ByVal httpClient As HttpClient, ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal cancellationToken As CancellationToken) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Dim funcRst As (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String)
            Try
                cancellationToken.ThrowIfCancellationRequested()
                ' 初始化uri类实例
                Dim baseAddress As New Uri(url)
                ' 构造请求消息体
                Dim requestMessage As New HttpRequestMessage(Http.HttpMethod.Get, baseAddress)
                requestMessage.SetRequestHeaders(requestHeaders, HttpMethod.GET)
                cancellationToken.ThrowIfCancellationRequested()
                ' 发送请求
                Dim response = Await httpClient.SendAsync(requestMessage， cancellationToken)
                funcRst = Await InternalGetResponseStringAsync(response)
            Catch ex As UriFormatException
                funcRst.HttpStatusCode = HttpStatusCode.BadRequest
                funcRst.Message = ex.Message
            Catch ex As TaskCanceledException
                funcRst.HttpStatusCode = HttpStatusCode.RequestTimeout
                funcRst.Message = ex.Message
            Catch ex As Exception
                funcRst.Message = ex.Message

                Logger.WriteLine(ex)
            End Try

            Return funcRst
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入。如果传入有包含编码的数据，可能会导致乱码或者数据丢失</param>
        ''' <param name="postContentEncoding">请求主体的编码方式，必须跟抓包的一致，否则可能会导致乱码</param>
        ''' <returns></returns>
        Public Shared Async Function PostAsync(ByVal httpClient As HttpClient, ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String， ByVal postContentEncoding As Text.Encoding, ByVal cancellationToken As CancellationToken) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Return Await InternalPostAsync(httpClient, url, requestHeaders, postContent.ToKeyValuePairs(False), postContentEncoding, cancellationToken)
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContentKvp">请求主体键值对集合，不需要编码，直接原字符串传入</param>
        ''' <param name="postContentEncoding">请求主体的编码方式，必须跟抓包的一致，否则可能会导致乱码</param>
        ''' <param name="cancellationToken">取消令牌</param>
        ''' <returns></returns>
        Public Shared Async Function InternalPostAsync(ByVal httpClient As HttpClient, ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContentKvp As IEnumerable(Of KeyValuePair(Of String, String))， ByVal postContentEncoding As Text.Encoding, ByVal cancellationToken As CancellationToken) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))

            Dim funcRst As (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String)

            Try
                cancellationToken.ThrowIfCancellationRequested()

                Dim content = GetHttpContent(requestHeaders, postContentKvp, postContentEncoding)

                cancellationToken.ThrowIfCancellationRequested()

                ' 初始化uri类实例
                Dim baseAddress As New Uri(url)
                ' 构造请求消息体
                Dim requestMessage As New HttpRequestMessage(Http.HttpMethod.Post, baseAddress) With {
                    .Content = content
                }
                requestMessage.SetRequestHeaders(requestHeaders, HttpMethod.POST)

                cancellationToken.ThrowIfCancellationRequested()

                ' 发送请求
                Dim response = Await httpClient.SendAsync(requestMessage， cancellationToken)
                funcRst = Await InternalGetResponseStringAsync(response)
            Catch ex As UriFormatException
                funcRst.HttpStatusCode = HttpStatusCode.BadRequest
                funcRst.Message = ex.Message
            Catch ex As TaskCanceledException
                funcRst.HttpStatusCode = HttpStatusCode.RequestTimeout
                funcRst.Message = ex.Message
            Catch ex As Exception
                funcRst.Message = ex.Message

                Logger.WriteLine(ex)
            End Try

            Return funcRst
        End Function
        Private Shared Function GetHttpContent(ByVal requestHeaders As Dictionary(Of String, String), ByVal postContentKvp As IEnumerable(Of KeyValuePair(Of String, String))， ByVal postContentEncoding As Text.Encoding) As HttpContent
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
                Dim postContent = ContentEncode(postContentKvp, postContentEncoding)
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
        Private Shared Function ContentEncode(ByVal postContentKvp As IEnumerable(Of KeyValuePair(Of String, String))， ByVal postContentEncoding As Text.Encoding) As String
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
        Private Shared Async Function DownloadFileAsyncInternal(ByVal httpClient As HttpClient, ByVal progressMessageHandler As ProgressMessageHandler, ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Stream As Stream))

            Dim success As Boolean
            Dim statusCode As HttpStatusCode
            Dim responseStream As Stream

            Try
                ' 初始化uri类实例
                Dim baseAddress As New Uri(url)

                ' 构造请求消息体
                Dim requestMessage As New HttpRequestMessage(Http.HttpMethod.Get, baseAddress)
                requestMessage.SetRequestHeaders(requestHeaders, HttpMethod.GET)

#Region "下载文件主体"
                ' 注册进度变化事件
                AddHandler progressMessageHandler.HttpReceiveProgress, AddressOf DownloadProgressChangedEventHandler

                ' 发送请求
                Dim response = Await httpClient.SendAsync(requestMessage)
                responseStream = Await response.Content.ReadAsStreamAsync
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
            End Try

            Return (success, statusCode, responseStream)
        End Function

        ''' <summary>
        ''' 异步下载文件
        ''' </summary>
        ''' <param name="url">下载链接</param>
        ''' <param name="fileFullPath">文件存储路径</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns></returns>
        Private Shared Async Function DownloadFileAsyncInternal(ByVal httpClient As HttpClient, ByVal progressMessageHandler As ProgressMessageHandler, ByVal url As String, ByVal fileFullPath As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
#Region "防呆检查"
            ' 判断传入的路径是否为文件（没有后缀名就视为非文件）
            Dim isFile = System.IO.Path.GetExtension(fileFullPath).Length > 0
            If Not isFile Then
                Return (False, HttpStatusCode.BadRequest, $"{NameOf(fileFullPath)}({fileFullPath}) is not a file")
            End If

            ' 判断文件夹是否存在
            Dim isDirectoryExists = System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(fileFullPath))
            If Not isDirectoryExists Then
                Return (False, HttpStatusCode.BadRequest, $"Directory not found of {NameOf(fileFullPath)}({fileFullPath})")
            End If
#End Region

            Dim success As Boolean
            Dim statusCode As HttpStatusCode
            Dim responseContent As String

            Try
                ' 初始化uri类实例
                Dim baseAddress As New Uri(url)

                ' 构造请求消息体
                Dim requestMessage As New HttpRequestMessage(Http.HttpMethod.Get, baseAddress)
                requestMessage.SetRequestHeaders(requestHeaders, HttpMethod.GET)

#Region "下载文件主体"
                ' 注册进度变化事件
                AddHandler progressMessageHandler.HttpReceiveProgress, AddressOf DownloadProgressChangedEventHandler

                ' 发送请求
                Dim response = Await httpClient.SendAsync(requestMessage)
                Using responseStream = Await response.Content.ReadAsStreamAsync
                    Dim bufferSize = 80 * 1024
                    Using fileStream As New System.IO.FileStream(fileFullPath,
                                                          System.IO.FileMode.Create,
                                                          System.IO.FileAccess.ReadWrite,
                                                          System.IO.FileShare.ReadWrite,
                                                          bufferSize,
                                                          True)
                        ' 可以用 CopyToAsync 方法一步写文件，但是不能做进度控制，没法向用户反馈
                        ' 所以还是用下面的方法
                        ' 关于效率：测试的时候 两种方法效率很接近(只测试过几最大M的文件)
                        Await responseStream.CopyToAsync(fileStream)

                        ' 报告下载完成
                        RaiseEvent DownloadFileCompleted(Nothing, New DownloadFileCompletedEventArgs(Nothing, False, Nothing))
                    End Using
                End Using
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
            End Try

            Return (success, statusCode, responseContent)
        End Function

        Private Shared Sub DownloadProgressChangedEventHandler(sender As Object, e As HttpProgressEventArgs)
            Dim progressMessageHandler = DirectCast(sender, ProgressMessageHandler)
            RemoveHandler progressMessageHandler.HttpReceiveProgress, AddressOf DownloadProgressChangedEventHandler

            RaiseEvent DownloadProgressChanged(sender, New DownloadProgressChangedEventArgs(e.ProgressPercentage, e.UserState, e.BytesTransferred, e.TotalBytes))
        End Sub

        Private Shared Sub UploadProgressChangedEventHandler(sender As Object, e As HttpProgressEventArgs)
            Dim progressMessageHandler = DirectCast(sender, ProgressMessageHandler)
            RemoveHandler progressMessageHandler.HttpSendProgress, AddressOf UploadProgressChangedEventHandler

            RaiseEvent UploadProgressChanged(sender, New UploadProgressChangedEventArgs(e.ProgressPercentage, e.UserState, e.BytesTransferred, e.TotalBytes))
        End Sub

        ''' <summary>
        ''' 异步上传文件
        ''' </summary>
        ''' <param name="uploadInfo">上传文件需要的相关信息</param>
        ''' <param name="requestHeaders"></param>
        ''' <returns></returns>
        Private Shared Async Function InternalUploadFileAsync(ByVal httpClient As HttpClient, ByVal progressMessageHandler As ProgressMessageHandler, ByVal uploadInfo As UploadInfo, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
#Region "防呆检查"
            ' 判断传入的路径是否为文件（没有后缀名就视为非文件）
            Dim isFile = System.IO.Path.GetExtension(uploadInfo.FileFullPath).Length > 0
            If Not isFile Then
                Return (False, HttpStatusCode.BadRequest, "uploadInfo.FileFullPath is not a legal file")
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
                End Try

                Return (success, statusCode, responseContent)
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
            '' 不使用缓存 使用缓存可能会得到错误的结果
            'WebRequest.DefaultCachePolicy = New Cache.RequestCachePolicy(Cache.RequestCacheLevel.NoCacheNoStore)
        End Sub

        ''' <summary>
        ''' 用于初始化类，如果操作需要带上cookies的话一定要在此函数传入；如果不需要传入cookies，直接调用相应函数即可
        ''' </summary>
        ''' <param name="cookies"></param>
        Public Shared Sub Init(ByRef cookies As CookieContainer)
            InitInternal(s_HttpClientHandler, cookies)
        End Sub

        ''' <summary>
        ''' 用于重新初始化类，比如更换账号等涉及到重新传入cookie的情况
        ''' </summary>
        ''' <param name="cookies">如果为Nothing，则只重新初始化类，不设置cookie</param>
        Public Shared Sub ReInit(ByRef cookies As CookieContainer)
            ReInit(cookies, True)
        End Sub

        ''' <summary>
        ''' 用于初始化类，如果操作需要带上cookies的话一定要在此函数传入；如果不需要传入cookies，直接调用相应函数即可
        ''' </summary>
        ''' <param name="cookies"></param>
        ''' <param name="allowAutoRedirect">指示处理程序是否应跟随重定向响应,默认为True</param>
        Public Shared Sub Init(ByRef cookies As CookieContainer, ByVal allowAutoRedirect As Boolean)
            InitInternal(s_HttpClientHandler, cookies)
        End Sub

        ''' <summary>
        ''' 用于重新初始化类，比如更换账号等涉及到重新传入cookie的情况
        ''' </summary>
        ''' <param name="cookies">如果为Nothing，则只重新初始化类，不设置cookie</param>
        ''' <param name="allowAutoRedirect">指示处理程序是否应跟随重定向响应,默认为True</param>
        Public Shared Sub ReInit(ByRef cookies As CookieContainer, ByVal allowAutoRedirect As Boolean)
            ' 把传入的cookie装盒并且设置到请求头
            ' cookie只能是这样设置.而且s_HttpClientHandler会内部会自动管理cookies
            ' 如果需要多次设置cookie或者是重新传入cookie，只能再次初始化 s_HttpClientHandler 实例
            InitInternal(s_HttpClient, s_HttpClientHandler, s_ProgressMessageHandler, allowAutoRedirect)

            If cookies Is Nothing Then
                s_IsInitialized = True
                Return
            End If

            s_HttpClientHandler.CookieContainer = cookies

            s_IsInitialized = True
        End Sub

        ''' <summary>
        ''' 设置请求头
        ''' </summary>
        ''' <param name="requestHeaders">请求头</param>
        ''' <param name="method">请求方法</param>
        Public Shared Sub SetRequestHeaders(ByRef requestHeaders As Dictionary(Of String, String)， ByVal method As HttpMethod)
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
        Public Shared Async Function HeadAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, ResponseHeaders As Headers.HttpResponseHeaders))
            m_CancellationTokenSource = New CancellationTokenSource
            m_CancellationToken = m_CancellationTokenSource.Token
            Return Await InternalHeadAsync(s_HttpClient, url, requestHeaders, m_CancellationToken)
        End Function

        ''' <summary>
        ''' 获取网页源码
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns>成功 Success 返回True，其他情况都返回False；
        ''' 成功Message返回获取到的源码，其他情况都返回具体错误信息；
        ''' HttpStatusCode返回相应的http请求状态码</returns>
        Public Shared Async Function GetAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            m_CancellationTokenSource = New CancellationTokenSource
            m_CancellationToken = m_CancellationTokenSource.Token
            Return Await InternalGetAsync(s_HttpClient, url, requestHeaders, m_CancellationToken)
        End Function

        ''' <summary>
        ''' 获取网页源码
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <returns></returns>
        Public Shared Async Function GetAsync(ByVal url As String) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            m_CancellationTokenSource = New CancellationTokenSource
            m_CancellationToken = m_CancellationTokenSource.Token
            Return Await InternalGetAsync(s_HttpClient, url, Nothing, m_CancellationToken)
        End Function

        ''' <summary>
        ''' 尝试获取网页源码直到成功，最多尝试三次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <returns></returns>
        Public Shared Async Function TryGetThreeTimeIfErrorAsync(ByVal url As String) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Return Await TryGetAsync(url， 3)
        End Function

        ''' <summary>
        ''' 尝试获取网页源码直到成功，最多尝试 <paramref name="tryTime"/> 次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Shared Async Function TryGetAsync(ByVal url As String, tryTime As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Dim funcRst As (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String) = (False, HttpStatusCode.Unused, String.Empty)

            Do
                m_CancellationTokenSource = New CancellationTokenSource
                m_CancellationToken = m_CancellationTokenSource.Token

                funcRst = Await InternalGetAsync(s_HttpClient, url, Nothing, m_CancellationToken)
                tryTime -= 1
            Loop Until funcRst.Success OrElse tryTime <= 0

            Return funcRst
        End Function

        ''' <summary>
        ''' 尝试获取网页源码直到成功，最多尝试 <paramref name="tryTime"/> 次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Shared Async Function TryGetAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), tryTime As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Dim funcRst As (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String) = (False, HttpStatusCode.Unused, String.Empty)

            Do
                m_CancellationTokenSource = New CancellationTokenSource
                m_CancellationToken = m_CancellationTokenSource.Token
                funcRst = Await InternalGetAsync(s_HttpClient, url, requestHeaders, m_CancellationToken)
                tryTime -= 1
            Loop Until funcRst.Success OrElse tryTime <= 0

            Return funcRst
        End Function

        ''' <summary>
        ''' 执行Get请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <returns></returns>
        Public Shared Async Function TryGetAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal retureIfContain As String, ByVal tryTime As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Dim funcRst As (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String) = (False, HttpStatusCode.Unused, String.Empty)

            Do
                funcRst = Await TryGetAsync(url, requestHeaders, 3)
                tryTime -= 1
            Loop Until (funcRst.Success AndAlso funcRst.Message.IndexOf(retureIfContain) > -1) OrElse tryTime <= 0
            If funcRst.Success Then
                funcRst.Success = funcRst.Message.IndexOf(retureIfContain) > -1 OrElse tryTime > 0
            End If

            Return funcRst
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入。如果传入有包含编码的数据，可能会导致乱码或者数据丢失</param>
        ''' <param name="postContentEncoding">>请求主体的编码方式，必须跟抓包的一致，否则可能会导致乱码</param>
        ''' <returns></returns>
        Public Shared Async Function PostAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String， ByVal postContentEncoding As Text.Encoding) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            m_CancellationTokenSource = New CancellationTokenSource
            m_CancellationToken = m_CancellationTokenSource.Token
            Return Await PostAsync(s_HttpClient, url, requestHeaders, postContent, postContentEncoding, m_CancellationToken)
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入；内部默认使用UTF-8编码，如果抓包得到的不是UTF-8编码，请调用可以自定义编码方式的重载函数</param>
        ''' <returns></returns>
        Public Shared Async Function PostAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            m_CancellationTokenSource = New CancellationTokenSource
            m_CancellationToken = m_CancellationTokenSource.Token
            Return Await PostAsync(s_HttpClient, url, requestHeaders, postContent, Text.Encoding.UTF8, m_CancellationToken)
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
        Public Shared Async Function TryPostAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String， ByVal postContentEncoding As Text.Encoding， tryTime As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Dim funcRst As (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String) = (False, HttpStatusCode.Unused, String.Empty)

            Do
                m_CancellationTokenSource = New CancellationTokenSource
                m_CancellationToken = m_CancellationTokenSource.Token
                funcRst = Await PostAsync(s_HttpClient, url, requestHeaders， postContent, postContentEncoding, m_CancellationToken)
                tryTime -= 1
            Loop Until funcRst.Success OrElse tryTime <= 0

            Return funcRst
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入；内部默认使用UTF-8编码，如果抓包得到的不是UTF-8编码，请调用可以自定义编码方式的重载函数</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Shared Async Function TryPostAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String， tryTime As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Dim funcRst As (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String) = (False, HttpStatusCode.Unused, String.Empty)

            Do
                m_CancellationTokenSource = New CancellationTokenSource
                m_CancellationToken = m_CancellationTokenSource.Token
                funcRst = Await PostAsync(s_HttpClient, url, requestHeaders， postContent, Text.Encoding.UTF8, m_CancellationToken)
                tryTime -= 1
            Loop Until funcRst.Success OrElse tryTime <= 0

            Return funcRst
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入。如果传入有包含编码的数据，可能会导致乱码或者数据丢失</param>
        ''' <param name="postContentEncoding">请求主体的编码方式，必须跟抓包的一致，否则可能会导致乱码</param>
        ''' <returns></returns>
        Public Shared Async Function TryPostThreeTimeIfErrorAsync(ByVal url As String， ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String， ByVal postContentEncoding As Text.Encoding) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Return Await TryPostAsync(url, requestHeaders, postContent, postContentEncoding, 3)
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入；内部默认使用UTF-8编码，如果抓包得到的不是UTF-8编码，请调用可以自定义编码方式的重载函数</param>
        ''' <returns></returns>
        Public Shared Async Function TryPostThreeTimeIfErrorAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Return Await TryPostAsync(url, requestHeaders, postContent, 3)
        End Function

        ''' <summary>
        ''' 异步下载文件
        ''' </summary>
        ''' <param name="url">下载链接</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns></returns>
        Public Shared Async Function DownloadFileAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Stream As Stream))
            Return Await DownloadFileAsyncInternal(s_HttpClient, s_ProgressMessageHandler, url, requestHeaders)
        End Function

        ''' <summary>
        ''' 尝试下载文件直到成功，最多尝试 <paramref name="tryTime"/> 次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Shared Async Function TryDownloadFileAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String)， tryTime As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Stream As Stream))
            Dim funcRst As (Success As Boolean, HttpStatusCode As HttpStatusCode, Stream As Stream) = (False, 0, Nothing)

            Do
                funcRst = Await DownloadFileAsyncInternal(s_HttpClient, s_ProgressMessageHandler, url, requestHeaders)
                tryTime -= 1
            Loop Until funcRst.Success OrElse tryTime <= 0

            Return funcRst
        End Function

        ''' <summary>
        ''' 尝试下载文件直到成功，最多尝试三次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <returns></returns>
        Public Shared Async Function TryDownloadFileThreeTimeIfErrorAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Stream As Stream))
            Return Await TryDownloadFileAsync(url, requestHeaders, 3)
        End Function


        '        ''' <summary>
        '        ''' 异步下载文件 （旧版通过自己计算的反馈进度 20171002）
        '        ''' </summary>
        '        ''' <param name="url">下载链接</param>
        '        ''' <param name="fileFullPath">文件存储路径</param>
        '        ''' <param name="requestHeaders">请求头</param>
        '        ''' <returns></returns>
        '        Public Shared Async Function DownloadFileAsync(ByVal url As String, ByVal fileFullPath As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
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
        Public Shared Async Function DownloadFileAsync(ByVal url As String, ByVal fileFullPath As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Return Await DownloadFileAsyncInternal(s_HttpClient, s_ProgressMessageHandler, url, fileFullPath, requestHeaders)
        End Function

        ''' <summary>
        ''' 尝试下载文件直到成功，最多尝试 <paramref name="tryTime"/> 次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Shared Async Function TryDownloadFileAsync(ByVal url As String, ByVal fileFullPath As String, ByVal requestHeaders As Dictionary(Of String, String)， tryTime As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Dim funcRst As (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String) = (False, HttpStatusCode.Unused, String.Empty)

            Do
                funcRst = Await DownloadFileAsyncInternal(s_HttpClient, s_ProgressMessageHandler, url, fileFullPath, requestHeaders)
                tryTime -= 1
            Loop Until funcRst.Success OrElse tryTime <= 0

            Return funcRst
        End Function

        ''' <summary>
        ''' 尝试下载文件直到成功，最多尝试三次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <returns></returns>
        Public Shared Async Function TryDownloadFileThreeTimeIfErrorAsync(ByVal url As String, ByVal fileFullPath As String, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Return Await TryDownloadFileAsync(url, fileFullPath, requestHeaders, 3)
        End Function

        ''' <summary>
        ''' 异步上传文件
        ''' </summary>
        ''' <param name="uploadInfo">上传文件需要的相关信息</param>
        ''' <param name="requestHeaders"></param>
        ''' <returns></returns>
        Public Shared Async Function UploadFileAsync(ByVal uploadInfo As UploadInfo, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Return Await InternalUploadFileAsync(s_HttpClient, s_ProgressMessageHandler, uploadInfo, requestHeaders)
        End Function


        ''' <summary>
        ''' 尝试上传文件直到成功，最多尝试 <paramref name="tryTime"/> 次
        ''' </summary>
        ''' <param name="uploadInfo">上传文件需要的相关信息</param>
        ''' <param name="requestHeaders"></param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Shared Async Function TryUploadFileAsync(ByVal uploadInfo As UploadInfo, ByVal requestHeaders As Dictionary(Of String, String)， tryTime As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Dim funcRst As (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String) = (False, HttpStatusCode.Unused, String.Empty)

            Do
                funcRst = Await InternalUploadFileAsync(s_HttpClient, s_ProgressMessageHandler, uploadInfo, requestHeaders)
                tryTime -= 1
            Loop Until funcRst.Success OrElse tryTime <= 0

            Return funcRst
        End Function

        ''' <summary>
        ''' 尝试上传文件直到成功，最多尝试三次
        ''' </summary>
        ''' <param name="uploadInfo">上传文件需要的相关信息</param>
        ''' <param name="requestHeaders"></param>
        ''' <returns></returns>
        Public Shared Async Function TryUploadFileThreeTimeIfErrorAsync(ByVal uploadInfo As UploadInfo, ByVal requestHeaders As Dictionary(Of String, String)) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Return Await TryUploadFileAsync(uploadInfo, requestHeaders, 3)
        End Function
#End Region


#Region "实例函数区"
        Public Sub SetDefaultCharSet(ByVal charSetName As String)
            HttpAsync.DefaultCharSet = charSetName
        End Sub

        ''' <summary>
        ''' 设置请求头
        ''' </summary>
        ''' <param name="requestHeaders">请求头</param>
        ''' <param name="method">请求方法</param>
        ''' <param name="noUse">此参数无实际意义，仅用作占位</param>
        Public Sub SetRequestHeaders(ByRef requestHeaders As Dictionary(Of String, String)， ByVal method As Net2.HttpMethod, ByVal noUse As Integer)
            instHttpClient.SetRequestHeaders(requestHeaders, method)
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
        Public Async Function HeadAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal noUse As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, ResponseHeaders As Headers.HttpResponseHeaders))
            m_CancellationTokenSource = New CancellationTokenSource
            m_CancellationToken = m_CancellationTokenSource.Token
            Return Await InternalHeadAsync(instHttpClient, url, requestHeaders, m_CancellationToken)
        End Function

        ''' <summary>
        ''' 获取网页源码
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns>成功 Success 返回True，其他情况都返回False；
        ''' 成功Message返回获取到的源码，其他情况都返回具体错误信息；
        ''' HttpStatusCode返回相应的http请求状态码</returns>
        Public Async Function GetAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal noUse As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            m_CancellationTokenSource = New CancellationTokenSource
            m_CancellationToken = m_CancellationTokenSource.Token
            Return Await InternalGetAsync(instHttpClient, url, requestHeaders, m_CancellationToken)
        End Function

        ''' <summary>
        ''' 获取网页源码
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <returns></returns>
        Public Async Function GetAsync(ByVal url As String, ByVal noUse As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            m_CancellationTokenSource = New CancellationTokenSource
            m_CancellationToken = m_CancellationTokenSource.Token
            Return Await InternalGetAsync(instHttpClient, url, Nothing, m_CancellationToken)
        End Function

        ''' <summary>
        ''' 尝试获取网页源码直到成功，最多尝试三次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <returns></returns>
        Public Async Function TryGetThreeTimeIfErrorAsync(ByVal url As String, ByVal noUse As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Return Await TryGetAsync(url, 3, 0)
        End Function

        ''' <summary>
        ''' 尝试获取网页源码直到成功，最多尝试 <paramref name="tryTime"/> 次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Async Function TryGetAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), tryTime As Integer, ByVal noUse As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Dim funcRst As (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String) = (False, HttpStatusCode.Unused, String.Empty)

            Do
                m_CancellationTokenSource = New CancellationTokenSource
                m_CancellationToken = m_CancellationTokenSource.Token
                funcRst = Await InternalGetAsync(instHttpClient, url, requestHeaders, m_CancellationToken)
                tryTime -= 1
            Loop Until funcRst.Success OrElse tryTime <= 0

            Return funcRst
        End Function

        ''' <summary>
        ''' 尝试获取网页源码直到成功，最多尝试三次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <returns></returns>
        Public Async Function TryGetAsync(ByVal url As String, ByVal tryTime As Integer, ByVal noUse As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Dim header As Dictionary(Of String, String)
            Return Await TryGetAsync(url, header， tryTime, 0)
        End Function

        ''' <summary>
        ''' 执行Get请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <returns></returns>
        Public Async Function TryGetAsync(ByVal url As String, ByVal retureIfContain As String, ByVal tryTime As Integer, ByVal noUse As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Dim funcRst As (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String) = (False, HttpStatusCode.Unused, String.Empty)

            Do
                funcRst = Await HttpAsync.TryGetAsync(url, 3)
                tryTime -= 1
            Loop Until (funcRst.Success AndAlso funcRst.Message.IndexOf(retureIfContain) > -1) OrElse tryTime <= 0
            If funcRst.Success Then
                funcRst.Success = funcRst.Message.IndexOf(retureIfContain) > -1 OrElse tryTime > 0
            End If

            Return funcRst
        End Function

        ''' <summary>
        ''' 执行Get请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <returns></returns>
        Public Async Function TryGetAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal referer As String, ByVal retureIfContain As String, ByVal tryTime As Integer, ByVal noUse As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Dim funcRst As (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String) = (False, HttpStatusCode.Unused, String.Empty)

            Do
                funcRst = Await TryGetAsync(url, requestHeaders, 3)
                tryTime -= 1
            Loop Until (funcRst.Success AndAlso funcRst.Message.IndexOf(retureIfContain) > -1) OrElse tryTime <= 0
            If funcRst.Success Then
                funcRst.Success = funcRst.Message.IndexOf(retureIfContain) > -1 OrElse tryTime > 0
            End If

            Return funcRst
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入。如果传入有包含编码的数据，可能会导致乱码或者数据丢失</param>
        ''' <param name="postContentEncoding">>请求主体的编码方式，必须跟抓包的一致，否则可能会导致乱码</param>
        ''' <returns></returns>
        Public Async Function PostAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String， ByVal postContentEncoding As Text.Encoding, ByVal noUse As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            m_CancellationTokenSource = New CancellationTokenSource
            m_CancellationToken = m_CancellationTokenSource.Token
            Return Await PostAsync(instHttpClient, url, requestHeaders, postContent, postContentEncoding, m_CancellationToken)
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入；内部默认使用UTF-8编码，如果抓包得到的不是UTF-8编码，请调用可以自定义编码方式的重载函数</param>
        ''' <returns></returns>
        Public Async Function PostAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String, ByVal noUse As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            m_CancellationTokenSource = New CancellationTokenSource
            m_CancellationToken = m_CancellationTokenSource.Token
            Return Await PostAsync(instHttpClient, url, requestHeaders, postContent, Text.Encoding.UTF8, m_CancellationToken)
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
        Public Async Function TryPostAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String， ByVal postContentEncoding As Text.Encoding， tryTime As Integer, ByVal noUse As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Dim funcRst As (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String) = (False, HttpStatusCode.Unused, String.Empty)

            Do
                m_CancellationTokenSource = New CancellationTokenSource
                m_CancellationToken = m_CancellationTokenSource.Token
                funcRst = Await PostAsync(instHttpClient, url, requestHeaders， postContent, postContentEncoding, m_CancellationToken)
                tryTime -= 1
            Loop Until funcRst.Success OrElse tryTime <= 0

            Return funcRst
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入；内部默认使用UTF-8编码，如果抓包得到的不是UTF-8编码，请调用可以自定义编码方式的重载函数</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Async Function TryPostAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String， tryTime As Integer, ByVal noUse As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Dim funcRst As (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String) = (False, HttpStatusCode.Unused, String.Empty)

            Do
                m_CancellationTokenSource = New CancellationTokenSource
                m_CancellationToken = m_CancellationTokenSource.Token

                funcRst = Await PostAsync(instHttpClient, url, requestHeaders， postContent, Text.Encoding.UTF8, m_CancellationToken)
                tryTime -= 1
            Loop Until funcRst.Success OrElse tryTime <= 0

            Return funcRst
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入。如果传入有包含编码的数据，可能会导致乱码或者数据丢失</param>
        ''' <param name="postContentEncoding">请求主体的编码方式，必须跟抓包的一致，否则可能会导致乱码</param>
        ''' <returns></returns>
        Public Async Function TryPostThreeTimeIfErrorAsync(ByVal url As String， ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String， ByVal postContentEncoding As Text.Encoding, ByVal noUse As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Return Await TryPostAsync(url, requestHeaders, postContent, postContentEncoding, 3, 0)
        End Function

        ''' <summary>
        ''' 发送POST请求
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="requestHeaders">请求头集合。类库内部默认的 ContentType 头值为 <see cref="DefaulMediaType"/> 。若请求头不区分大小写，请使用带有 <see cref="IEqualityComparer(Of T)"/> 的重载来实例化集合。</param>
        ''' <param name="postContent">请求主体，不需要编码，直接原字符串传入；内部默认使用UTF-8编码，如果抓包得到的不是UTF-8编码，请调用可以自定义编码方式的重载函数</param>
        ''' <returns></returns>
        Public Async Function TryPostThreeTimeIfErrorAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal postContent As String, ByVal noUse As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Return Await TryPostAsync(url, requestHeaders, postContent, 3, 0)
        End Function

        ''' <summary>
        ''' 异步下载文件
        ''' </summary>
        ''' <param name="url">下载链接</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns></returns>
        Public Async Function DownloadFileAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal noUse As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Stream As Stream))
            Return Await DownloadFileAsyncInternal(instHttpClient, instProgressMessageHandler, url, requestHeaders)
        End Function

        ''' <summary>
        ''' 尝试下载文件直到成功，最多尝试 <paramref name="tryTime"/> 次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Async Function TryDownloadFileAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String)， tryTime As Integer, ByVal noUse As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Stream As Stream))
            Dim funcRst As (Success As Boolean, HttpStatusCode As HttpStatusCode, Stream As Stream) = (False, 0, Nothing)

            Do
                funcRst = Await DownloadFileAsyncInternal(instHttpClient, instProgressMessageHandler, url, requestHeaders)
                tryTime -= 1
            Loop Until funcRst.Success OrElse tryTime <= 0

            Return funcRst
        End Function

        ''' <summary>
        ''' 尝试下载文件直到成功，最多尝试三次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <returns></returns>
        Public Async Function TryDownloadFileThreeTimeIfErrorAsync(ByVal url As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal noUse As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Stream As Stream))
            Return Await TryDownloadFileAsync(url, requestHeaders, 3， 0)
        End Function

        ''' <summary>
        ''' 异步下载文件
        ''' </summary>
        ''' <param name="url">下载链接</param>
        ''' <param name="fileFullPath">文件存储路径</param>
        ''' <param name="requestHeaders">请求头</param>
        ''' <returns></returns>
        Public Async Function DownloadFileAsync(ByVal url As String, ByVal fileFullPath As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal noUse As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Return Await DownloadFileAsyncInternal(instHttpClient, instProgressMessageHandler, url, fileFullPath, requestHeaders)
        End Function

        ''' <summary>
        ''' 尝试下载文件直到成功，最多尝试 <paramref name="tryTime"/> 次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Async Function TryDownloadFileAsync(ByVal url As String, ByVal fileFullPath As String, ByVal requestHeaders As Dictionary(Of String, String)， tryTime As Integer, ByVal noUse As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Dim funcRst As (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String) = (False, HttpStatusCode.Unused, String.Empty)

            Do
                funcRst = Await DownloadFileAsyncInternal(instHttpClient, instProgressMessageHandler, url, fileFullPath, requestHeaders)
                tryTime -= 1
            Loop Until funcRst.Success OrElse tryTime <= 0

            Return funcRst
        End Function

        ''' <summary>
        ''' 尝试下载文件直到成功，最多尝试三次
        ''' </summary>
        ''' <param name="url">网页链接</param>
        ''' <returns></returns>
        Public Async Function TryDownloadFileThreeTimeIfErrorAsync(ByVal url As String, ByVal fileFullPath As String, ByVal requestHeaders As Dictionary(Of String, String), ByVal noUse As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Return Await TryDownloadFileAsync(url, fileFullPath, requestHeaders, 3, 0)
        End Function

        ''' <summary>
        ''' 异步上传文件
        ''' </summary>
        ''' <param name="uploadInfo">上传文件需要的相关信息</param>
        ''' <param name="requestHeaders"></param>
        ''' <returns></returns>
        Public Async Function UploadFileAsync(ByVal uploadInfo As UploadInfo, ByVal requestHeaders As Dictionary(Of String, String), ByVal noUse As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Return Await InternalUploadFileAsync(instHttpClient, instProgressMessageHandler, uploadInfo, requestHeaders)
        End Function


        ''' <summary>
        ''' 尝试上传文件直到成功，最多尝试 <paramref name="tryTime"/> 次
        ''' </summary>
        ''' <param name="uploadInfo">上传文件需要的相关信息</param>
        ''' <param name="requestHeaders"></param>
        ''' <param name="tryTime">尝试次数。成功会立刻返回，失败会继续尝试直到用完尝试次数</param>
        ''' <returns></returns>
        Public Async Function TryUploadFileAsync(ByVal uploadInfo As UploadInfo, ByVal requestHeaders As Dictionary(Of String, String)， tryTime As Integer, ByVal noUse As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Dim funcRst As (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String) = (False, HttpStatusCode.Unused, String.Empty)

            Do
                funcRst = Await InternalUploadFileAsync(instHttpClient, instProgressMessageHandler, uploadInfo, requestHeaders)
                tryTime -= 1
            Loop Until funcRst.Success OrElse tryTime <= 0

            Return funcRst
        End Function

        ''' <summary>
        ''' 尝试上传文件直到成功，最多尝试三次
        ''' </summary>
        ''' <param name="uploadInfo">上传文件需要的相关信息</param>
        ''' <param name="requestHeaders"></param>
        ''' <returns></returns>
        Public Async Function TryUploadFileThreeTimeIfErrorAsync(ByVal uploadInfo As UploadInfo, ByVal requestHeaders As Dictionary(Of String, String), ByVal noUse As Integer) As Task(Of (Success As Boolean, HttpStatusCode As HttpStatusCode, Message As String))
            Return Await TryUploadFileAsync(uploadInfo, requestHeaders, 3, 0)
        End Function
#End Region

#Region "内部类"
        Public Class UploadInfo
            ''' <summary>
            ''' 请求链接
            ''' </summary>
            ''' <returns></returns>
            Public Property RequestUrl As String
            ''' <summary>
            ''' 本地文件的绝对路径
            ''' </summary>
            ''' <returns></returns>
            Public Property FileFullPath As String
            ''' <summary>
            ''' 设置 HTTP 响应上的 Content-Type 内容标头值
            ''' DispositionType 的 值必须要跟抓包的一致
            ''' ContentDispositionHeaderValue 的 name 值必须要跟抓包的一致
            ''' FileName 的值没有要求， 可以随便写， 建议写成文件的名称（带后缀）
            ''' </summary>
            ''' <returns></returns>
            Public Property ContentDisposition As ContentDispositionHeaderValue
        End Class
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

            If instProgressMessageHandler IsNot Nothing Then
                instProgressMessageHandler.Dispose()
                instProgressMessageHandler = Nothing
            End If

            If instHttpClientHandler IsNot Nothing Then
                instHttpClientHandler.Dispose()
                instHttpClientHandler = Nothing
            End If

            If instHttpClient IsNot Nothing Then
                instHttpClient.Dispose()
                instHttpClient = Nothing
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

            If m_CancellationTokenSource IsNot Nothing Then
                m_CancellationTokenSource.Dispose()
                m_CancellationTokenSource = Nothing
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
