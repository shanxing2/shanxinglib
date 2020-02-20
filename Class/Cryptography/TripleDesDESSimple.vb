Imports System.Security.Cryptography

Namespace ShanXingTech.Cryptography
    ''' <summary>
    ''' 加密解密
    ''' </summary>
    Public NotInheritable Class TripleDesDESSimple
        Implements IDisposable

        ' 明文：神即道, 道法自然, 如来|闪星网络信息科技 ShanXingTech Q2287190283
        ' 算法：古典密码中的有密钥换位密码 密钥：ShanXingTech
        Public Const ShanXingTechQ2287190283 = "神闪X7,SQB道信T2道网N9来A2D如H2C然技HA即星I1|N8E法息E8,络G0自科C3"

        ''' <summary>
        ''' 默认编码方式为Unicode
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Encoding As Text.Encoding

        Private ReadOnly m_TripleDes As TripleDESCryptoServiceProvider
        Private Const HexKeyLength As Integer = 4
        Private Const ValidDateLength As Integer = 10
        Private Const ValidDateFormat As String = "yyyy-MM-dd"
        Private Const DefaultValidDate As String = "2099-12-31"
        Public Const OperateError As String = "Error"

        ''' <summary>
        ''' 用来初始化 3DES 加密服务提供程序的构造函数。
        ''' </summary>
        ''' <param name="key">密钥 控制 EncryptString 和 DecryptString 方法。</param>
        Sub New(ByVal key As String)
            ' Initialize the crypto provider.
            m_TripleDes = New TripleDESCryptoServiceProvider
            m_TripleDes.Key = TruncateHash(key, m_TripleDes.KeySize \ 8)
            m_TripleDes.IV = TruncateHash(String.Empty, m_TripleDes.BlockSize \ 8)

            Me.Encoding = Text.Encoding.Unicode
        End Sub

        Sub New(ByVal key As String， ByVal encoding As Text.Encoding)
            ' Initialize the crypto provider.
            m_TripleDes = New TripleDESCryptoServiceProvider
            m_TripleDes.Key = TruncateHash(key, m_TripleDes.KeySize \ 8)
            m_TripleDes.IV = TruncateHash(String.Empty, m_TripleDes.BlockSize \ 8)

            Me.Encoding = encoding
        End Sub

        ''' <summary>
        ''' 将从指定密钥的哈希创建指定长度的字节数组
        ''' </summary>
        ''' <param name="key"></param>
        ''' <param name="length"></param>
        ''' <returns></returns>
        Private Function TruncateHash(ByVal key As String, ByVal length As Integer) As Byte()
            Dim hash As Byte()
            Try
                Using sha1 As New SHA1CryptoServiceProvider
                    ' Hash the key.
                    Dim keyBytes() As Byte = Encoding.GetBytes(key)
                    hash = sha1.ComputeHash(keyBytes)
                End Using
            Catch ex As Exception
                Return Nothing
            End Try


            ' Truncate(缩短) or pad（填充） the hash.
            ReDim Preserve hash(length - 1)

            Return hash
        End Function

        ''' <summary>
        ''' 加密字符串
        ''' </summary>
        ''' <param name="plaintext">明文</param>
        ''' <returns></returns>
        Public Function EncryptString(ByVal plaintext As String) As String
            ' Convert the plaintext string to a byte array.
            Dim plaintextBytes() As Byte = Encoding.GetBytes(plaintext)

            Try
                ' Create the stream.
                Dim ms As New System.IO.MemoryStream
                ' Create the encoder to write to the stream.
                Using encStream As New CryptoStream(ms,
                                                  m_TripleDes.CreateEncryptor(),
                                                  CryptoStreamMode.Write)

                    ' Use the crypto stream to write the byte array to the stream.
                    encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
                    encStream.FlushFinalBlock()

                    ' Convert the encrypted stream to a printable string.
                    Return Convert.ToBase64String(ms.ToArray)
                End Using
            Catch ex As Exception
                Return String.Empty
            End Try
        End Function

        ''' <summary>
        ''' 解密字符串
        ''' </summary>
        ''' <param name="encryptedtext">密文</param>
        ''' <returns></returns>
        Public Function DecryptString(ByVal encryptedtext As String) As String
            Try
                If encryptedtext.Length = 0 Then Return String.Empty

                ' Convert the encrypted text string to a byte array.
                Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedtext)

                ' Create the stream.
                Dim ms As New System.IO.MemoryStream
                ' Create the decoder to write to the stream.
                Using decStream As New CryptoStream(ms,
                                                  m_TripleDes.CreateDecryptor(),
                                                  CryptoStreamMode.Write)


                    ' Use the crypto stream to write the byte array to the stream.
                    decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
                    decStream.FlushFinalBlock()

                    ' Convert the plaintext stream to a string.
                    Return Encoding.GetString(ms.ToArray)
                End Using
            Catch ex As Exception
                Return String.Empty
            End Try
        End Function

        ''' <summary>
        ''' 解密 (如果加密的时候，没有传入有效期，那么默认有效期的值为 2099-12-31 )
        ''' </summary>
        ''' <param name="cipherText">密文</param>
        ''' <returns></returns>
        Public Shared Function Decrypt(ByVal cipherText As String, ByVal appName As String) As (ValidDate As Date, PlainText As String)
            Dim plainText = String.Empty
            Dim key = NameOf(ShanXingTechQ2287190283)
            Dim company = key

            Dim validDate As Date = Now

            '如果密文长度小于等于 HexKeyLength+ValidDateLength，则直接不需要解密，返回error
            If cipherText.Length <= (HexKeyLength + ValidDateLength) Then Return (Now, OperateError)

            Try
                Using aes As New Cryptography.TripleDesDESSimple(key)
                    plainText = aes.DecryptString(cipherText)
                End Using

                If plainText.Length = 0 Then Return (Now, OperateError)

                ' 从字符串中获取有效日期(后5位就是有效日期)
                validDate = CDate(plainText.Substring(plainText.Length - ValidDateLength))

                ' 从字符串中获取密钥(后四位就是密钥长度的16进制形式)
                Dim keyLength = Convert.ToInt32(plainText.Substring(plainText.Length - HexKeyLength - ValidDateLength, HexKeyLength), 16)
                key = plainText.Substring(plainText.Length - keyLength - HexKeyLength - ValidDateLength, keyLength)

                ' 去掉后面的 常量ShanXingTech的常量名 然后再解密
                plainText = plainText.Substring(0, plainText.Length - key.Length - HexKeyLength - ValidDateLength)
                Using aes As New Cryptography.TripleDesDESSimple(key)
                    plainText = aes.DecryptString(plainText)
                End Using

                ' 如果密文里面没有包含软件名称相关信息，也是无效
                Dim exeName = plainText.Substring(plainText.IndexOf(company) + company.Length)
                If appName.IndexOf(exeName, StringComparison.OrdinalIgnoreCase) = -1 Then Return (Now, OperateError)

                ' 去掉后面的 常量ShanXingTech的常量名以及软件名 然后再解密
                plainText = plainText.Substring(0, plainText.IndexOf(company))
                Using aes As New Cryptography.TripleDesDESSimple(key)
                    plainText = aes.DecryptString(plainText)
                End Using
            Catch ex As Exception
                plainText = OperateError
            End Try

            Return (validDate, plainText)
        End Function

        ''' <summary>
        ''' 加密
        ''' </summary>
        ''' <param name="plainText">明文</param>
        ''' <param name="key">密钥</param>
        ''' <param name="validDate">有效天数,建议不要大于99年的天数总和</param>
        ''' <returns></returns>
        Public Shared Function Encrypt(ByVal plainText As String, ByVal key As String, ByVal validDate As Date) As String
            If plainText.Length = 0 Then Return OperateError

            Dim cipherText = String.Empty

            Try
                Using aes As New Cryptography.TripleDesDESSimple(key)
                    ' 加密一次之后再把加密后的字符串跟 常量ShanXingTech的常量名已经软件名连接，再进行加密
                    cipherText = aes.EncryptString(aes.EncryptString(plainText) & NameOf(ShanXingTechQ2287190283) & key)
                End Using

                ' 用常量 ShanXingTech 的名称当做新密钥再去加密一次
                Using aes As New Cryptography.TripleDesDESSimple(NameOf(ShanXingTechQ2287190283))
                    ' 把 原始密钥的值和密钥长度值的16进制值(4个字符串长度)以及5位长度的有效日期拼接起来，放到加密字符串的最后
                    cipherText = aes.EncryptString(cipherText & key & key.Length.ToString("X4") & validDate.ToString(ValidDateFormat))
                End Using
            Catch ex As Exception
                plainText = OperateError
            End Try

            Return cipherText
        End Function

        ''' <summary>
        ''' 加密 (默认密文有效期到 2099-12-31 )
        ''' </summary>
        ''' <param name="plainText">明文</param>
        ''' <param name="key">密钥</param>
        ''' <returns></returns>
        Public Shared Function Encrypt(ByVal plainText As String, ByVal key As String) As String
            If plainText.Length = 0 Then Return OperateError

            Dim cipherText = Encrypt(plainText, key, CDate(DefaultValidDate))

            Return cipherText
        End Function

#Region "IDisposable Support"
        ' 要检测冗余调用
        Private isDisposed2 As Boolean

        ' IDisposable
        Protected Sub Dispose(disposing As Boolean)
            ' 窗体内的控件调用Close或者Dispose方法时，isDisposed2的值为True
            If isDisposed2 Then Return

            If disposing Then
                ' TODO: 释放托管资源(托管对象)。
                If m_TripleDes IsNot Nothing Then
                    m_TripleDes.Dispose()
                End If
            End If

            ' TODO: 释放未托管资源(未托管对象)并在以下内容中替代 Finalize()。
            ' TODO: 将大型字段设置为 null。

            isDisposed2 = True
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