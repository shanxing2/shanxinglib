Imports System.IO
Imports System.Net
Imports System.Security
Imports System.Text
Imports System.Text.RegularExpressions
Imports NETWORKLIST
Imports ShanXingTech.Text2
Imports ShanXingTech.Win32API

Namespace ShanXingTech.Net2
    Public NotInheritable Class NetHelper
        ''' <summary>
        ''' 获取网络连接状态(跟桌面右下角网络图标同步)
        ''' </summary>
        ''' <returns>已联网——True,未联网——False</returns>
        Public Shared Function IsConnectedToInternet() As Boolean
            Dim funcRst As Boolean

            ' https://docs.microsoft.com/zh-cn/windows/win32/api/_nla/
            Dim networkManager = New NetworkListManagerClass
            Dim networks = networkManager.GetNetworks(NLM_ENUM_NETWORK.NLM_ENUM_NETWORK_CONNECTED).Cast(Of INetwork)
            For Each nw In networks
                'Debug.WriteLine($"Name:{nw.GetName()} {NameOf(nw.IsConnectedToInternet)}:{nw.IsConnectedToInternet}")
                If nw.IsConnectedToInternet Then
                    funcRst = True
                    Exit For
                End If
            Next

            Return funcRst
        End Function

        ''' <summary>
        ''' 设置请求标头 此方法可以解决“此标头必须使用适当的属性进行修改” 问题
        ''' </summary>
        ''' <param name="headers"></param>
        ''' <param name="name"></param>
        ''' <param name="value"></param>
        Public Shared Sub SetHeaderValue(headers As WebHeaderCollection, name As String, value As String)
            Dim [property] = GetType(WebHeaderCollection).GetProperty("InnerCollection".TryGetIntern, System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.NonPublic)

            If [property] Is Nothing Then Return
            Dim collection = DirectCast([property].GetValue(headers, Nothing), Specialized.NameValueCollection)
            collection(name) = value
        End Sub

        ''' <summary>
        ''' 获取一个Url的根域名
        ''' </summary>
        ''' <param name="url">如果为空或者Nothing直接返回<paramref name="url"/>的值</param>
        ''' <returns>如https://myseller.taobao.com，返回 .taobao.com</returns>
        Public Shared Function GetRootDomain(ByVal url As String) As String
            If String.IsNullOrEmpty(url) Then Return url

            ' 获取传入的domain的根域名（主域名）
            ' 如https://myseller.taobao.com，根域名是.taobao.com
            ' 如http://xiaomi.tmall.com，根域名是.tmall.com
            Dim pattern = "(\.\w+)?\.\w+\.(com.cn|com|net.cn|net|org.cn|name|org|gov.cn|gov|cn|com.hk|mobi|me|info|name|biz|cc|tv|asia|hk|网络|公司|中国|ac.cn|bj.cn|sh.cn|tj.cn|cq.cn|he.cn|sx.cn|nm.cn|ln.cn|jl.cn|hl.cn|js.cn|zj.cn|ah.cn|fj.cn|jx.cn|sd.cn|ha.cn|hb.cn|hn.cn|gd.cn|gx.cn|hi.cn|sc.cn|gz.cn|yn.cn|xz.cn|sn.cn|gs.cn|qh.cn|nx.cn|xj.cn|tw.cn|hk.cn|mo.cn|travel|tw|com.tw|la|sh|ac|io|ws|us|tm|vc|ag|bz|in|mn|me|sc|co|org.tw|jobs|tel)\b"
            Dim match = Regex.Match(url, pattern， RegexOptions.IgnoreCase Or RegexOptions.Compiled)
            Dim domain = match.Groups(0).Value
            Return domain
        End Function

        ''' <summary>
        ''' 获取一个Url的域名
        ''' </summary>
        ''' <param name="url"></param>
        ''' <returns></returns>
        Public Shared Function GetDomain(ByVal url As String) As String
            ' 如果传入的链接不包含协议http或者htpps，那就直接返回，不需要做后续操作
            If url.Length = 0 OrElse Not url.StartsWith("http".TryGetIntern) Then Return url

            ' 新方法
            Dim domain2 = url
            Try
                Dim uri = New Uri(url)
                domain2 = String.Concat(uri.Scheme, "://".TryGetIntern, uri.Host)
            Catch ex As UriFormatException
                '
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try

            Return domain2
        End Function

        ''' <summary>
        ''' 字符串型的cookie转换成cookie型的cookiecollection
        ''' </summary>
        ''' <param name="cookieStr"></param>
        ''' <param name="cookie"></param>
        ''' <param name="domain"></param>
        Public Shared Sub GetCookieFromString(ByVal cookieStr As String, ByRef cookie As CookieContainer, ByVal domain As String)
            '如果cookie集合未初始化 则进行初始化
            If cookie Is Nothing Then
                cookie = New CookieContainer()
            End If
            'Debug.Print(String.Concat(My.Resources.CallSubName, Reflection.MethodBase.GetCurrentMethod.Name, " 转换cookie开始"))
            Try

                Dim cookiesKeyValuePair As String() = cookieStr.Split({";"}, StringSplitOptions.RemoveEmptyEntries)
                Dim cookieName As String = String.Empty
                Dim cookieValue As String = String.Empty

                For Each cookieKeyValuePair As String In cookiesKeyValuePair
                    Dim equalSymbolPostion = cookieKeyValuePair.IndexOf("=")
                    cookieName = cookieKeyValuePair.Substring(0, equalSymbolPostion).Trim()
                    cookieValue = cookieKeyValuePair.Substring(equalSymbolPostion + "=".Length).Trim()

                    'Debug.Print(cookieName & "=" & cookieValue)
                    Dim ck As New Cookie(cookieName, cookieValue) With {
                        .Domain = domain
                    }
                    cookie.Add(ck)
                Next

                'Debug.Print(String.Concat(My.Resources.CallSubName, Reflection.MethodBase.GetCurrentMethod.Name, " 转换cookie成功"))
            Catch ex As Exception
                Logger.WriteLine(ex)
            Finally
                ' Debug.Print(String.Concat(My.Resources.CallSubName, Reflection.MethodBase.GetCurrentMethod.Name, " 转换cookie结束"))
            End Try
        End Sub

#Region "获取webbrowser登录成功后的cookie，需要带上登录成功后的URL"
        ' &H2000=8192 浏览器版本
        Private Const InternetCookieHttpOnly As Integer = &H2000

        ''' <summary>
        ''' 获取webbrowser登录成功后的cookie，需要带上登录成功后的URL
        ''' webBrowser1.Document.Cookie无法获取HttpOnly属性的Cookie
        ''' 可以获取IE8及以上的cookie，其他版本IE需要自己去读取cookie文件处理
        ''' </summary>
        ''' <param name="url"></param>
        ''' <returns></returns>
        <SecurityCritical()>
        Public Shared Function GetCookie(url As String) As String
            Dim funcRst As String = String.Empty

            Dim size As Integer

            ' 想获取数据大小 双字节 /2之后就是数据的长度
            ' 第三个参数为非 Nothing或者为StringBuilder时 会一直返回false
            ' 因此第三个参数应该设置为nothing
            If InternetGetCookieEx(url, vbNullString, Nothing, size, InternetCookieHttpOnly, IntPtr.Zero) Then
                If size <= 0 Then
                    Return String.Empty
                End If

                ' 然后根据取得的数据大小初始化缓存，以及再次传入刚才返回的size
                Dim sb = New StringBuilder(size + 1)
                If Not InternetGetCookieEx(url, vbNullString, sb, size, InternetCookieHttpOnly, IntPtr.Zero) Then
                    Return String.Empty
                End If

                funcRst = sb.ToString
            End If

            'Dim lastErrorCode = Marshal.GetLastWin32Error '<-- 259
            Return funcRst
        End Function
#End Region

        ''' <summary>
        ''' 删除跟传入的链接 <paramref name="url"/> 的domain有关的所有cookies 
        ''' </summary>
        ''' <param name="url">一个链接或者可以直接传入domain</param>
        ''' <returns>成功返回删除个数，失败或者没有相关cookies可以删除则返回0</returns>
        Public Shared Function DeleteCookiesAboutDomain(ByVal url As String) As Integer
            ' 通过传入的url获取根域名
            Dim domain = GetRootDomain(url)

            ' 去掉域名前面的 ”.“
            If "."c = domain.Chars(0) Then
                domain = domain.Remove(0, 1)
            End If

            ' 获取ie cookie目录下的所有cookie文件
            Dim cookieFiles = My.Computer.FileSystem.GetFiles(Environment.GetFolderPath(Environment.SpecialFolder.Cookies))

            Dim deleteCount As Integer

            For Each file In cookieFiles.AsParallel
                Dim cookieStr = String.Empty
                Try
                    Using sr As New StreamReader(file)
                        cookieStr = sr.ReadToEnd()
                    End Using

                    ' 删除属于此domain的cookie
                    If cookieStr.IndexOf(domain, StringComparison.OrdinalIgnoreCase) > -1 Then
                        System.IO.File.Delete(file)

                        deleteCount += 1
                    End If
                Catch ex As Exception
                    Continue For
                End Try
            Next

            Return deleteCount
        End Function
    End Class
End Namespace
