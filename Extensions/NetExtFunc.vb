Imports System.Drawing
Imports System.Net
Imports System.Net.Http
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows.Forms
Imports Microsoft.Win32
Imports ShanXingTech
Imports ShanXingTech.Text2
Imports ShanXingTech.Win32API
Imports ShanXingTech.Windows2

Namespace ShanXingTech

    Partial Public Module ExtensionFunc
        ''' <summary>
        ''' 获取 htmlTag 相对于浏览器左上角的位置
        ''' </summary>
        ''' <param name="htmlTag"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function OffsetPoint(ByVal htmlTag As HtmlElement) As (Point As Point, TopParent As HtmlElement)
            Dim topParent = htmlTag
            Dim oldTopParent = topParent
            Dim pointOffset As New Point()
            Do
                pointOffset.X += topParent.OffsetRectangle.Left
                pointOffset.Y += topParent.OffsetRectangle.Top
                oldTopParent = topParent
                topParent = topParent.OffsetParent
            Loop Until topParent Is Nothing

            Return (pointOffset, oldTopParent)
        End Function

        ''' <summary>
        ''' 获取 htmlTag 的相对于滚动条的位置
        ''' </summary>
        ''' <param name="htmlTag"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function OffsetScrollPoint(ByVal htmlTag As HtmlElement) As Point
            Dim offset = OffsetPoint(htmlTag)

            Dim bodyTag = offset.TopParent.Document.GetElementsByTagName("HTML")(0)
            Dim point As New Point(offset.Point.X - bodyTag.ScrollLeft, offset.Point.Y - bodyTag.ScrollTop)

            Return point
        End Function

        ''' <summary>
        ''' 鼠标从浏览器左上角到移动到某个标签元素中心处（后台操作，不影响用户鼠标）
        ''' <para>每移动一个点，随机延时1-2毫秒</para>
        ''' </summary>
        ''' <param name="htmlTag"></param>
        ''' <param name="browser"></param>
        ''' <param name="backgroundOperate">后台操作</param>
        <Extension()>
        Public Sub MouseMoveTo(ByRef htmlTag As HtmlElement, ByRef browser As WebBrowser, ByVal backgroundOperate As Boolean)
            Dim randN As New Random()
            ' 模拟鼠标在浏览器内移动（从浏览器左上角到 htmlTag 的中间 ）

            ' 获取 htmlTag 的相对于滚动条的位置
            Dim htmlTagPointOfScrollX = htmlTag.OffsetScrollPoint.X + htmlTag.OffsetRectangle.Width \ 2
            Dim htmlTagPointOfScrollY = htmlTag.OffsetScrollPoint.Y + htmlTag.OffsetRectangle.Height \ 2

            ' 由点 browserPointOfScreen 向 htmlTag 元素的中间连续移动
            ' 获取browser相对于屏幕的位置
            Dim browserPointOfScreen = browser.PointToScreen(New Point(0, 0))
            ' y = kx
            Dim k = (htmlTagPointOfScrollY + browserPointOfScreen.Y) / (browserPointOfScreen.X + htmlTagPointOfScrollX)
            Dim y = 0
            Dim hwndOfBrowser = browser.GetBrowserReallyHwnd
            For x = browserPointOfScreen.X To browserPointOfScreen.X + htmlTagPointOfScrollX
                y = CInt(x * k)
                If backgroundOperate Then
                    ' 后台操作
                    SendMessage(hwndOfBrowser, MouseEventFlags.WM_MOUSEMOVE, 0, x + (y << 16))
                    Windows2.Delay(randN.Next(1, 2))
                Else
                    ' 前台操作
                    SetCursorPos(x, y)
                End If
            Next
        End Sub

        ''' <summary>
        ''' 鼠标从浏览器左上角到移动到某个标签元素中心处（默认后台操作，不影响用户鼠标）
        ''' <para>每移动一个点，随机延时1-2毫秒</para>
        ''' </summary>
        ''' <param name="htmlTag"></param>
        ''' <param name="browser"></param>
        <Extension()>
        Public Sub MouseMoveTo(ByRef htmlTag As HtmlElement, ByRef browser As WebBrowser)
            MouseMoveTo(htmlTag， browser， True)
        End Sub

        ''' <summary>
        ''' 开启或者关闭默认浏览器webbrowser加载网页的提示音（嘟嘟声???）
        ''' </summary>
        <Extension()>
        Public Function DisableNavigationSounds(ByRef browser As WebBrowser, ByVal disable As Boolean) As Integer
            Return CoInternetSetFeatureEnabled(INTERNETFEATURELIST.FEATURE_DISABLE_NAVIGATION_SOUNDS, SET_FEATURE_ON_PROCESS, disable)
        End Function

        ''' <summary>
        ''' 改变传入WebBrowser默认的UA
        ''' </summary>
        ''' <param name="browser"></param>
        ''' <param name="userAgent"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ChangeUserAgent(ByRef browser As WebBrowser, ByVal userAgent As String) As Integer
            Return UrlMkSetSessionOption(URLMON_OPTION_USERAGENT, userAgent, userAgent.Length, 0)
        End Function

        ''' <summary>
        ''' 把传入的WebBrowser的UA设置回默认值
        ''' </summary>
        ''' <param name="browser"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function RefreshUserAgent(ByRef browser As WebBrowser) As Integer
            Return UrlMkSetSessionOption(URLMON_OPTION_USERAGENT_REFRESH, Nothing, 0, 0)
        End Function

        ''' <summary>
        ''' 获取webbrowser登录成功后的cookie，需要带上登录成功后的URL。
        ''' 如果处理过的登录成功后cookie不能用于其他页面，请查看cookie的Domain是否异常。
        ''' <para>最多可以获取 m_Cookies.PerDomainCapacity 个cookies。如果传入的cookies为Nothing,则内部会初始化PerDomainCapacity为66。如果cookies.PerDomainCapacity小于66，那么将会先把cookies.PerDomainCapacity的值设置为66，然后再获取cookies</para>
        ''' <para>webBrowser1.Document.Cookie无法获取HttpOnly属性的Cookie</para>
        ''' <para>可以获取IE8及以上的cookie，其他版本IE需要自己去读取cookie文件处理</para>
        ''' </summary>
        ''' <param name="cookies"></param>
        ''' <param name="url">操作成功之后的Url</param>
        <Extension()>
        Public Sub GetFromUrl(ByRef cookies As CookieContainer, ByVal url As String)
            If cookies Is Nothing Then
                cookies = New CookieContainer() With {.PerDomainCapacity = 66}
            End If

            Try
                Dim cookieKeyValuePairs = Net2.NetHelper.GetCookie(url)
                ParseCookie(cookies， cookieKeyValuePairs， url)
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try
        End Sub


        ''' <summary>
        ''' KeyValue形式的cookie字符串转换成Cookie并存入CookieContainer。
        ''' <para>最多可以获取 m_Cookies.PerDomainCapacity 个cookies。如果传入的cookies为Nothing,则内部会初始化PerDomainCapacity为66。如果cookies.PerDomainCapacity小于66，那么将会先把cookies.PerDomainCapacity的值设置为66，然后再获取cookies</para>
        ''' </summary>
        ''' <param name="cookies"></param>
        ''' <param name="cookieKeyValuePairs"></param>
        ''' <param name="domain"></param>
        <Extension()>
        Public Sub GetFromKeyValuePairs(ByRef cookies As CookieContainer, ByVal cookieKeyValuePairs As String， ByVal domain As String)
            If cookies Is Nothing Then
                cookies = New CookieContainer() With {.PerDomainCapacity = 66}
            Else
                If cookies.PerDomainCapacity < 66 Then cookies.PerDomainCapacity = 66
            End If

            ParseCookie(cookies， cookieKeyValuePairs， domain)
        End Sub

        ''' <summary>
        ''' 解析字符串格式的cookie，保存到传入的<paramref name="Cookies"/>参数中
        ''' </summary>
        ''' <param name="cookies"></param>
        ''' <param name="cookieKeyValuePairs"></param>
        ''' <param name="domain"></param>
        Private Sub ParseCookie(ByRef Cookies As CookieContainer, ByVal cookieKeyValuePairs As String， ByVal domain As String)
            ' ##########################20170330########################
            ' 获取传入的domain的根域名（主域名）
            ' 如https://myseller.taobao.com，根域名是.taobao.com
            ' 暂时没有找到方法确定某个cookie所属的domain，暂时设定为传入url的根域名
            ' ##########################20170704########################
            domain = Net2.NetHelper.GetRootDomain(domain)

            Dim cookieKeyValuePairArray As String() = cookieKeyValuePairs.Split({";"}, StringSplitOptions.RemoveEmptyEntries)
            Dim name As String
            Dim value As String
            Dim cookie As New Cookie With {
                .Domain = domain
            }


            For Each cookieKeyValuePair In cookieKeyValuePairArray
                Dim equalSymbolPostion = cookieKeyValuePair.IndexOf("=")
                ' 有些cookie只有key而没有value,此处需要特殊处理
                If equalSymbolPostion = -1 Then
                    name = cookieKeyValuePair.Trim()
                    value = String.Empty
                Else
                    name = cookieKeyValuePair.Substring(0, equalSymbolPostion).Trim()
                    value = cookieKeyValuePair.Substring(equalSymbolPostion + "=".Length).Trim()
#Region "注意信息"
                    ' 有时候会有比较奇怪的cookie，比如1688的某一个cookie
                    ' $Version=1; _tmp_ck_0="pcJxPh5%2B7lEEYZ%2F5hRcrgIj5hksWozhFCoqTDNHQYbf6Jp2qnuz1d3QVBrGnlupaIw1Nx0lGykurR5oci3YKyX9ReiC8%2F%2Bv9Wisk4H9dS75FnxfcRBFap7lOuoBa26wFBwqHNLOAJeBtEDez%2B%2BLNEJZm7g93D4Yhf9PBHuVTfh3fLZZP9hVIv5TalkceJi3emEE1Z3%2BB61cypPReOmMHlwWCXUWwmbHgH5NHI1BX%2FPlkbuw28EsX6iS%2BZ40N6XXDjFRWa8dH%2FIDK%2FviOP6oEUS%2Bm%2FvZ92hGMaXsYZt0yIQGlntd6eNZAE9Xq0ckLMCOPbH9JoHRpFABUrTnQxH5VetwAsOCYKcnMNzcLShnC%2BnWkqtre1BUm2CwD48qK0O0taq6gm0BzL1w%3D"; $Path=/; $Domain=.1688.com
                    ' 获取的时候我们自取 
                    ' name = _tmp_ck_0
                    ' value = pcJxPh5%2B7lEEYZ%2F5hRcrgIj5hksWozhFCoqTDNHQYbf6Jp2qnuz1d3QVBrGnlupaIw1Nx0lGykurR5oci3YKyX9ReiC8%2F%2Bv9Wisk4H9dS75FnxfcRBFap7lOuoBa26wFBwqHNLOAJeBtEDez%2B%2BLNEJZm7g93D4Yhf9PBHuVTfh3fLZZP9hVIv5TalkceJi3emEE1Z3%2BB61cypPReOmMHlwWCXUWwmbHgH5NHI1BX%2FPlkbuw28EsX6iS%2BZ40N6XXDjFRWa8dH%2FIDK%2FviOP6oEUS%2Bm%2FvZ92hGMaXsYZt0yIQGlntd6eNZAE9Xq0ckLMCOPbH9JoHRpFABUrTnQxH5VetwAsOCYKcnMNzcLShnC%2BnWkqtre1BUm2CwD48qK0O0taq6gm0BzL1w%3D
#End Region

                    ' 去掉开头 和 后面的 所有 "
                    While value.StartsWith("""")
                        value = value.Remove(0, 1)
                    End While

                    While value.EndsWith("""")
                        value = value.Remove(value.Length - 1, 1)
                    End While
                End If

                ' Cookie 的值。value 参数不能包含分号 (;) 或逗号 (,)，除非它们包含在转义的双引号中。
                ' 有时候有些cookie比较特殊，导致无法成功添加进去
                ' 比如 百度网盘上传的cookie中有一个键值队为Hm_lvt_7a3960b6f067eb0085b7f96ff5e660b0=1503060813,1503309277,1503577240,1503654236
                ' 后面的值有逗号会导致无法添加
                If value.IndexOf(",") > -1 Then
                    value = value.Replace(",", "%2C")
                End If
                If value.IndexOf(";") > -1 Then
                    value = value.Replace(",", "%3B")
                End If

                cookie.Name = name
                cookie.Value = value

                Try
                    If domain.Length > 0 Then
                        Cookies.Add(cookie)
                    Else
                        Cookies.Add(New Uri(domain), cookie)
                    End If
                Catch ex As Exception
                    Logger.WriteLine(ex)

                    Continue For
                End Try
            Next
        End Sub

        ''' <summary>
        ''' CookieContainer转换成 'Key=Value;' 形式
        ''' </summary>
        ''' <param name="cookies"></param>
        ''' <param name="exceptExpired">为True时，已经失效/过期的cookie将不会输出</param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToKeyValuePairs(ByRef cookies As CookieContainer, ByVal exceptExpired As Boolean) As String
            If cookies Is Nothing Then
                Throw New ArgumentNullException(String.Format(My.Resources.NullReference， NameOf(cookies)))
            End If

            Dim sb = New StringBuilder(360)
            Dim cookiesName As New List(Of String)

            For Each c In cookies.GetEnumerator
                If exceptExpired AndAlso c.Expired Then Continue For

                ' 不添加重复值
                If cookiesName.Contains(c.Name) Then
                    Continue For
                Else
                    cookiesName.Add(c.Name)
                End If

                ' 有时候会有比较奇怪的cookie，比如1688的某一个cookie
                ' $Version=1; _tmp_ck_0="pcJxPh5%2B7lEEYZ%2F5hRcrgIj5hksWozhFCoqTDNHQYbf6Jp2qnuz1d3QVBrGnlupaIw1Nx0lGykurR5oci3YKyX9ReiC8%2F%2Bv9Wisk4H9dS75FnxfcRBFap7lOuoBa26wFBwqHNLOAJeBtEDez%2B%2BLNEJZm7g93D4Yhf9PBHuVTfh3fLZZP9hVIv5TalkceJi3emEE1Z3%2BB61cypPReOmMHlwWCXUWwmbHgH5NHI1BX%2FPlkbuw28EsX6iS%2BZ40N6XXDjFRWa8dH%2FIDK%2FviOP6oEUS%2Bm%2FvZ92hGMaXsYZt0yIQGlntd6eNZAE9Xq0ckLMCOPbH9JoHRpFABUrTnQxH5VetwAsOCYKcnMNzcLShnC%2BnWkqtre1BUm2CwD48qK0O0taq6gm0BzL1w%3D"; $Path=/; $Domain=.1688.com
                ' 获取的时候我们自取 
                ' name = _tmp_ck_0
                ' value = pcJxPh5%2B7lEEYZ%2F5hRcrgIj5hksWozhFCoqTDNHQYbf6Jp2qnuz1d3QVBrGnlupaIw1Nx0lGykurR5oci3YKyX9ReiC8%2F%2Bv9Wisk4H9dS75FnxfcRBFap7lOuoBa26wFBwqHNLOAJeBtEDez%2B%2BLNEJZm7g93D4Yhf9PBHuVTfh3fLZZP9hVIv5TalkceJi3emEE1Z3%2BB61cypPReOmMHlwWCXUWwmbHgH5NHI1BX%2FPlkbuw28EsX6iS%2BZ40N6XXDjFRWa8dH%2FIDK%2FviOP6oEUS%2Bm%2FvZ92hGMaXsYZt0yIQGlntd6eNZAE9Xq0ckLMCOPbH9JoHRpFABUrTnQxH5VetwAsOCYKcnMNzcLShnC%2BnWkqtre1BUm2CwD48qK0O0taq6gm0BzL1w%3D
                ' 去掉开头 和 后面的 所有 "
                Dim value = c.Value
                While """"c = value.Chars(0)
                    value = value.Remove(0, 1)
                End While

                While """"c = value.Chars(value.Length - 1)
                    value = value.Remove(value.Length - 1, 1)
                End While

                sb.Append(c.Name).Append("="c).Append(value).Append(";"c)
            Next

            Dim funcRst As String = StringBuilderCache.GetStringAndReleaseBuilder(sb)

            Return funcRst
        End Function

        ''' <summary>
        ''' CookieContainer转换成Key=Value;形式,已经失效/过期的cookie将不会输出
        ''' </summary>
        ''' <param name="cookies"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToKeyValuePairs(ByRef cookies As CookieContainer) As String
            Return ToKeyValuePairs(cookies, True)
        End Function

        ''' <summary>
        ''' 使 键为 <paramref name="key"/> 的Cookie无效/过期（Expires = Now.AddDays(-1)）。往后的Http请求将不会再带上这个Cookie，除非其再次生效。
        ''' </summary>
        ''' <param name="cookies"></param>
        ''' <param name="key"></param>
        ''' <returns>修改成功返回True,无修改或者失败返回False</returns>
        <Extension()>
        Public Function TryExpire(ByRef cookies As CookieContainer, ByVal key As String) As Boolean
            If cookies Is Nothing OrElse key.IsNullOrEmpty Then
                Return False
            End If

            Dim expired As Boolean
            Try
                For Each c In cookies.GetEnumerator
                    If key = c.Name Then
                        c.Expires = Now.AddDays(-1)
                        expired = c.Expired
                        Exit For
                    End If
                Next
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try

            Return expired
        End Function

        ''' <summary>
        ''' 使 键为 <paramref name="key"/> 的Cookie无效/过期（Expires = Now.AddDays(-1)）。往后的Http请求将不会再带上这个Cookie，除非其再次生效。
        ''' </summary>
        ''' <param name="cookies"></param>
        ''' <param name="key"></param>
        ''' <returns>修改成功返回True,无修改或者失败返回False</returns>
        <Extension()>
        Public Function TryExpire(ByRef cookies As CookieContainer, ByVal key() As String) As Boolean
            If cookies Is Nothing OrElse
                key Is Nothing OrElse
                key.Length = 0 Then
                Return False
            End If

            Dim expired As Boolean
            Try
                For Each c In cookies.GetEnumerator
                    If key.Contains(c.Name) Then
                        c.Expires = Now.AddDays(-1)
                        expired = c.Expired
                    End If
                Next
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try

            Return expired
        End Function

        ''' <summary>
        ''' 检查Cookie实例中是否包含某个键
        ''' </summary>
        ''' <param name="cookies"></param>
        ''' <param name="key"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ContainKey(ByRef cookies As CookieContainer, ByVal key As String) As Boolean
            If cookies Is Nothing Then
                Throw New ArgumentNullException(String.Format(My.Resources.NullReference， NameOf(cookies)))
            End If

            If key Is Nothing Then
                Throw New ArgumentNullException(String.Format(My.Resources.NullReference， NameOf(key)))
            End If

            If key.Length = 0 Then
                Return False
            End If

            For Each c In cookies.GetEnumerator
                If c.Name = key Then
                    Return True
                End If
            Next

            Return False
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="cookies"></param>
        ''' <returns></returns>
        <Extension()>
        Public Iterator Function GetEnumerator(ByVal cookies As CookieContainer) As IEnumerable(Of Cookie)
            If cookies Is Nothing Then
                Throw New ArgumentNullException(String.Format(My.Resources.NullReference， NameOf(cookies)))
            End If

            ' 从非公共成员中获取cookie到表
            Dim table As Hashtable = DirectCast(cookies.GetType().InvokeMember("m_domainTable", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.GetField Or Reflection.BindingFlags.Instance, Nothing, cookies, Array.Empty(Of Object)()), Hashtable)

            If table Is Nothing Then Return

            ' 从非公共成员中获取cookie到表 遍历
            For Each pathList In table.Values
                Dim lstCookieCol As SortedList = DirectCast(pathList.[GetType]().InvokeMember("m_list", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.GetField Or Reflection.BindingFlags.Instance, Nothing, pathList, Array.Empty(Of Object)()), SortedList)
                For Each colCookies As ICollection In lstCookieCol.Values
                    For Each c As Cookie In colCookies
                        Yield c
                    Next
                Next
            Next
        End Function

        ''' <summary>
        ''' 更改程序内部IE浏览器默认的版本号（IE7）为 <paramref name="browserEmulationMode"/>
        ''' <para>需要重新打开程序更改才会生效</para>
        ''' </summary>
        ''' <param name="browser"></param>
        ''' <param name="browserEmulationMode">程序内部IE浏览器的仿真版本</param>
        ''' <param name="appName">app.exe and app.vshost.exe</param>
        ''' <returns>设置成功或者无效设置返回True，设置失败返回False</returns>
        <Extension()>
        Public Function SetVersionEmulation(ByRef browser As WebBrowser, ByVal browserEmulationMode As BrowserEmulationMode, ByVal appName As String) As (Success As Boolean, NeedToReOpen As Boolean)
            Dim funcRst As Boolean
            Dim needToReOpen As Boolean

            Try
                ' 不存在则创建
                Using browserEmulationKey As RegistryKey = If(Registry.CurrentUser.OpenSubKey(BROWSER_EMULATION_KEY, RegistryKeyPermissionCheck.ReadWriteSubTree), Registry.CurrentUser.CreateSubKey(BROWSER_EMULATION_KEY))

                    If browserEmulationKey Is Nothing Then Return (False, False)

                    ' 如果未设置或者小于需要设置的值，则执行设置
                    Dim value = browserEmulationKey.GetValue(appName)
                    If value Is Nothing OrElse CInt(value) < browserEmulationMode Then
                        browserEmulationKey.SetValue(appName, browserEmulationMode, RegistryValueKind.DWord)
                        ' 新设置或者是重新设置 Emulation值后需要重新打开软件
                        needToReOpen = True
                    End If

                    funcRst = True
                End Using
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try

            Return (funcRst, needToReOpen)
        End Function

        ''' <summary>
        ''' 禁止程序内部IE浏览器某个默认行为（IE8及以上）
        ''' </summary>
        ''' <param name="browser"></param>
        ''' <param name="behavior"></param>
        <Extension()>
        Public Function SuppressWininetBehavior(ByRef browser As WebBrowser， ByVal behavior As Integer) As Boolean

            ' SOURCE: http://msdn.microsoft.com/en-us/library/windows/desktop/aa385328%28v=vs.85%29.aspx
            '    * INTERNET_OPTION_SUPPRESS_BEHAVIOR (81):
            '    *      A general purpose option that is used to suppress behaviors on a process-wide basis. 
            '    *      The lpBuffer parameter of the function must be a pointer to a DWORD containing the specific behavior to suppress. 
            '    *      This option cannot be queried with InternetQueryOption. 
            '    *      
            '    * INTERNET_SUPPRESS_COOKIE_PERSIST (3):
            '    *      Suppresses the persistence of m_Cookies, even if the server has specified them as persistent.
            '    *      Version:  Requires Internet Explorer 8.0 or later.
            '    

            ' 禁止Cookie保留，也就是每个窗口具有独立Cookie
            Dim lpBuffer = New IntPtr(behavior)
            Dim success = InternetSetOption(IntPtr.Zero, INTERNET_OPTION_SUPPRESS_BEHAVIOR, lpBuffer, Marshal.SizeOf(behavior.GetType))
            IntPtrFree(lpBuffer)

            Return success
        End Function

        ''' <summary>
        ''' <see cref="WebBrowser"/> 执行JavaScript并返回结果
        ''' </summary>
        ''' <param name="webBrowser"></param>
        ''' <param name="js">完整的js函数实现</param>
        ''' <param name="jsFuncName"><paramref name="js"/> 中函数的名称</param>
        ''' <returns></returns>
        <Extension()>
        Public Function RunJs(ByRef webBrowser As WebBrowser, ByVal js As String, ByVal jsFuncName As String) As Object
            If webBrowser?.Document Is Nothing Then Return Nothing

            If js.IsNullOrEmpty Then
                Throw New ArgumentNullException(String.Format(My.Resources.NullReference， NameOf(js)))
            End If

            If jsFuncName.IsNullOrEmpty Then
                Throw New ArgumentNullException(String.Format(My.Resources.NullReference， NameOf(jsFuncName)))
            End If

            Dim script = webBrowser.Document.CreateElement("script")
            If script Is Nothing Then Return Nothing

            script.SetAttribute("type", "text/javascript")
            script.SetAttribute("id", jsFuncName)
            script.SetAttribute("text", js)

            ' 函数名做id,不重复添加
            Dim existsChildren As Boolean
            For Each element2 As HtmlElement In webBrowser.Document.Body.Children.AsParallel
                If jsFuncName = element2.Id Then
                    existsChildren = True
                    Exit For
                End If
            Next

            If Not existsChildren Then
                Dim element = webBrowser.Document.Body.AppendChild(script)
                If element Is Nothing Then Return Nothing
            End If

            Return webBrowser.Document.InvokeScript(jsFuncName)
        End Function

        ''' <summary>
        ''' <see cref="WebBrowser"/> 执行JavaScript并返回结果
        ''' </summary>
        ''' <param name="webBrowser"></param>
        ''' <param name="jsFuncName">Html页面中JS函数的名称，如果此函数不在Html页面中，请使用 <seealso cref="RunJs(ByRef WebBrowser, String, String)"/> </param>
        ''' <returns></returns>
        <Extension()>
        Public Function RunJs(ByRef webBrowser As WebBrowser, ByVal jsFuncName As String) As Object
            If webBrowser?.Document Is Nothing Then Return Nothing

            If jsFuncName.IsNullOrEmpty Then
                Throw New ArgumentNullException(String.Format(My.Resources.NullReference， NameOf(jsFuncName)))
            End If

            Return webBrowser.Document.InvokeScript(jsFuncName)
        End Function

        ''' <summary>
        ''' 设置请求头
        ''' </summary>
        ''' <param name="client"></param>
        ''' <param name="requestHeaders"></param>
        <Extension()>
        Public Sub SetRequestHeaders(ByRef client As HttpClient, ByRef requestHeaders As Dictionary(Of String, String)， ByVal method As Net2.HttpMethod)
            ' 先清空所有头再按需添加，因为 DefaultRequestHeaders 没有提供方法来替换某个请求头原来的值
            'client.DefaultRequestHeaders.Clear()

            If requestHeaders IsNot Nothing Then
                ' 如果强求头里面包含 m_Cookies ,那么就抛出异常，提示用户使用 Init 或者 ReInit方法传入 m_Cookies
                If requestHeaders.ContainsKey("cookie") Then
                    Throw New NotSupportedException("不支持在请求头传入Cookie，请使用 Init 或者 ReInit方法设置Cookie")
                End If

                For Each header In requestHeaders
                    ' Post时不应该在此处设置content-type请求头
                    If method = Net2.HttpMethod.POST AndAlso
                        header.Key.Equals("content-type", StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If

                    ' 没有这个请求头或者是已有的请求头跟现在的不同时，才添加
                    If Not client.DefaultRequestHeaders.Contains(header.Key) Then
                        client.DefaultRequestHeaders.TryAddWithoutValidation(header.Key, header.Value)
                    End If
                Next
            End If
        End Sub

        ''' <summary>
        ''' 设置请求头
        ''' </summary>
        ''' <param name="requestMessage"></param>
        ''' <param name="requestHeaders"></param>
        <Extension()>
        Public Sub SetRequestHeaders(ByRef requestMessage As HttpRequestMessage, ByRef requestHeaders As Dictionary(Of String, String)， ByVal method As Net2.HttpMethod)
            If requestHeaders Is Nothing Then Return

            ' 如果强求头里面包含 m_Cookies ,那么就抛出异常，提示用户使用 Init 或者 ReInit方法传入 m_Cookies
            If requestHeaders.ContainsKey("cookie") Then
                    Throw New NotSupportedException("不支持传入Cookie，请使用 Init 或者 ReInit方法设置Cookies")
                End If

                Dim keys = requestHeaders.Keys.ToArray
                For Each key In keys
                    ' Post时不应该在此处设置content-type请求头
                    If method = Net2.HttpMethod.POST AndAlso key.Equals("content-type", StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If

                    ' 没有这个请求头时才添加
                    If Not requestMessage.Headers.Contains(key) Then
                        requestMessage.Headers.TryAddWithoutValidation(key, requestHeaders(key))
                    End If
                Next
            'For Each header In requestHeaders
            '                ' Post时不应该在此处设置content-type请求头
            '                If method = Net2.HttpMethod.POST AndAlso header.Key.Equals("content-type", StringComparison.OrdinalIgnoreCase) Then
            '                    Continue For
            '                End If

            '                ' 没有这个请求头时才添加
            '                If Not requestMessage.Headers.Contains(header.Key) Then
            '                    requestMessage.Headers.TryAddWithoutValidation(header.Key, header.Value)
            '                End If
            '            Next
        End Sub

        ''' <summary>
        ''' 获取内置 <see cref="WebBrowser"/> 控件的真正句柄（<see cref="WebBrowser.Handle"/> 获取到的句柄是经过封装的外层句柄，并不能用来操作 <see cref="WebBrowser"/> 内的内容，比如弹窗）。
        ''' 注：必须等 <see cref="WebBrowser"/> 加载完毕之后才能获取。
        ''' </summary>
        ''' <param name="browser"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function GetBrowserReallyHwnd(ByRef browser As WebBrowser) As IntPtr
            ' 方法一
            'Dim sbClassName As New StringBuilder(256)
            'Dim childHandle = GetWindow(browser.Handle, WindowType.GW_CHILD)

            'Do While childHandle <> IntPtr.Zero
            '    GetClassName(childHandle, sbClassName, sbClassName.Capacity)

            '    If "Internet Explorer_Server" = sbClassName.ToString Then
            '        Return childHandle
            '    End If

            '    childHandle = GetWindow(childHandle, WindowType.GW_CHILD)
            'Loop
            ' 方法二
            Dim childHandle = FindWindowEx(browser.Handle, IntPtr.Zero, "Shell Embedding", Nothing)
            If IntPtr.Zero = childHandle Then Return IntPtr.Zero
            childHandle = FindWindowEx(childHandle, IntPtr.Zero, "Shell DocObject View", Nothing)
            If IntPtr.Zero = childHandle Then Return IntPtr.Zero
            childHandle = FindWindowEx(childHandle, IntPtr.Zero, "Internet Explorer_Server", Nothing)

            Return childHandle
        End Function
    End Module
End Namespace