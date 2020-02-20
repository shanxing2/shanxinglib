Imports System.Net
Imports System.Net.Sockets


Namespace ShanXingTech.Net2
    ''' <summary>
    ''' 这个人很懒，什么都没写
    ''' </summary>
    Public Class IPHelper
        ' 明文：神即道, 道法自然, 如来|闪星网络信息科技 ShanXingTech Q2287190283
        ' 算法：古典密码中的有密钥换位密码 密钥：ShanXingTech
        Public Const ShanXingTechQ2287190283 = "神闪X7,SQB道信T2道网N9来A2D如H2C然技HA即星I1|N8E法息E8,络G0自科C3"

        ''' <summary>
        ''' 获取公网IP
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function GetPublicIP() As String
            Dim result As String = GetPublicIPFromChinaz()

			' 用第一种方法获取失败 并且网络已连接
			If result.Length = 0 AndAlso NetHelper.IsConnected() Then
				result = GetPublicIPFromWanWang()
			End If

			' 如果用上面的方法都获取不了，使用以下具有不确定性的方法
			' Dns.GetHostEntry方法取得的IP可能会有几个
			' 目前暂时米有准确的方法获取到正确的公网IP
			If result.Length = 0 Then
                result = GetPublicIPFromIPHostEntry(AddressFamily.InterNetwork)
            End If

            Return result
        End Function

        ''' <summary>
        ''' 获取本地IPV4版本的地址
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function GetLocalIPV4() As String
            Dim result = GetPublicIPFromIPHostEntry(AddressFamily.InterNetwork)

            Return result
        End Function

        ''' <summary>
        ''' 获取本地IPV6版本的地址
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function GetLocalIPV6() As String
            Dim result = GetPublicIPFromIPHostEntry(AddressFamily.InterNetworkV6)

            Return result
        End Function

        ''' <summary>
        ''' 站长之家获取本地公网IP地址接口（百度搜ip调用的接口之一）
        ''' </summary>
        ''' <returns></returns>
        Private Shared Function GetPublicIPFromChinaz() As String
            Dim result As String = String.Empty
            Dim html As String = String.Empty

            Try
                Using web As New WebClient
                    ' 站长之家获取本地公网IP地址接口（百度搜ip调用的接口之一）
                    html = web.DownloadString("http://ip.chinaz.com/getip.aspx")
                    result = Text.RegularExpressions.Regex.Match(html, "\b\d{1,3}(\.\d{1,3}){3}\b").Value
                End Using
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 万网获取本地公网IP地址接口
        ''' </summary>
        ''' <returns></returns>
        Private Shared Function GetPublicIPFromWanWang() As String
            Dim result As String = String.Empty
            Dim html As String = String.Empty

            Try
                Using web As New WebClient
                    ' 万网获取本地公网IP地址接口
                    html = web.DownloadString("http://www.net.cn/static/customercare/yourip.asp")
                    result = Text.RegularExpressions.Regex.Match(html, "\b\d{1,3}(\.\d{1,3}){3}\b").Value
                End Using
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 获取本地IP地址
        ''' </summary>
        ''' <returns></returns>
        Private Shared Function GetPublicIPFromIPHostEntry(ByVal addressFamily As AddressFamily) As String
            Dim result As String = String.Empty

            Try
                Dim myEntry As IPHostEntry = Dns.GetHostEntry(Dns.GetHostName())
                result = myEntry.AddressList.FirstOrDefault(Function(c) c.AddressFamily = addressFamily).ToString()
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try

            Return result
        End Function


    End Class

End Namespace
