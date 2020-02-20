Imports System.IO
Imports System.Security.Cryptography

Namespace ShanXingTech.Cryptography
    Public NotInheritable Class AesSimple
        Implements IDisposable

        ' 明文：神即道, 道法自然, 如来|闪星网络信息科技 ShanXingTech Q2287190283
        ' 算法：古典密码中的有密钥换位密码 密钥：ShanXingTech
        Public Const ShanXingTechQ2287190283 = "神闪X7,SQB道信T2道网N9来A2D如H2C然技HA即星I1|N8E法息E8,络G0自科C3"

#Region "枚举区"
        Enum InputOutputMode
            Hex
            Base64
        End Enum
#End Region

#Region "字段区"
        Private m_Aes As Aes
        Private decryptor As ICryptoTransform
        Private encryptor As ICryptoTransform

#End Region

#Region "属性区"
        ''' <summary>
        ''' 对称算法的密钥
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Key As Byte()
        ''' <summary>
        ''' 对称算法的初始化向量(偏移量)
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property IV As Byte()
#End Region

#Region "构造函数"

#End Region
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="key">密钥</param>
        ''' <param name="iv">初始化向量(偏移量)</param>
        Public Sub New(ByVal key As Byte(), ByVal iv As Byte())
            If key Is Nothing OrElse key.Length = 0 Then
                Throw New ArgumentNullException(String.Format(My.Resources.NullReference, NameOf(key)))
            End If
            If iv Is Nothing OrElse iv.Length = 0 Then
                Throw New ArgumentNullException(String.Format(My.Resources.NullReference, NameOf(iv)))
            End If

            Me.Key = key
            Me.IV = iv

            m_Aes = Aes.Create()
            m_Aes.Key = Me.Key
            m_Aes.IV = Me.IV
        End Sub

        ''' <summary>
        ''' 加密
        ''' </summary>
        ''' <param name="plainText"></param>
        ''' <returns>返回加密字符串</returns>
        ''' <param name="outputMode">指示应该将参数 <paramref name="plainText"/> 加密成何种格式。默认 <see cref="InputOutputMode.Base64"/></param>
        Public Function EncryptString(ByVal plainText As String, Optional ByVal outputMode As InputOutputMode = InputOutputMode.Base64) As String
            ' Check arguments.
            If plainText Is Nothing OrElse plainText.Length <= 0 Then
                Throw New ArgumentNullException(String.Format(My.Resources.NullReference, NameOf(plainText)))
            End If

            Dim encryptor As ICryptoTransform
            Try
                Dim encrypted() As Byte
                ' Create the streams used for encryption.
                Dim msEncrypt As New MemoryStream()
                encryptor = m_Aes.CreateEncryptor(Key, IV)
                Dim csEncrypt As New CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write)
                Using swEncrypt As New StreamWriter(csEncrypt)
                    'Write all data to the stream.
                    swEncrypt.Write(plainText)
                End Using
                encrypted = msEncrypt.ToArray()
                msEncrypt = Nothing

                ' Return the encrypted bytes from the memory stream.
                Return If(outputMode = InputOutputMode.Hex,
                    Convert.FromBase64String(Convert.ToBase64String(encrypted)).ToHexString,
                    Convert.ToBase64String(encrypted))
            Catch ex As Exception
                Return String.Empty
            Finally
                If encryptor IsNot Nothing Then
                    encryptor.Dispose()
                    encryptor = Nothing
                End If
            End Try


        End Function

        ''' <summary>
        ''' 解密
        ''' </summary>
        ''' <param name="cipherText">已加密的字符串</param>
        ''' <param name="inputMode">指示参数 <paramref name="cipherText"/> 目前采用何种格式展现。默认 <see cref="InputOutputMode.Base64"/></param>
        ''' <returns></returns>
        Public Function DecryptString(ByVal cipherText As String, Optional ByVal inputMode As InputOutputMode = InputOutputMode.Base64) As String
            ' Check arguments.
            If cipherText Is Nothing OrElse cipherText.Length = 0 Then
                Throw New ArgumentNullException(String.Format(My.Resources.NullReference, NameOf(cipherText)))
            End If

            ' Declare the string used to hold
            ' the decrypted text.
            Dim plaintext As String = Nothing
            Dim decryptor As ICryptoTransform

            Try
                Dim tempCipherText = If(inputMode = InputOutputMode.Hex,
                    Convert.ToBase64String(cipherText.HexStringToBytes),
                    cipherText)

                ' Create the streams used for decryption.
                Dim msDecrypt As New MemoryStream(Convert.FromBase64String(tempCipherText))
                decryptor = m_Aes.CreateDecryptor(Key, IV)
                Dim csDecrypt As New CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read)
                Using srDecrypt As New StreamReader(csDecrypt)
                    ' Read the decrypted bytes from the decrypting stream
                    ' and place them in a string.
                    plaintext = srDecrypt.ReadToEnd()
                End Using

                Return plaintext
            Catch ex As Exception
                Return String.Empty
            Finally
                If decryptor IsNot Nothing Then
                    decryptor.Dispose()
                    decryptor = Nothing
                End If
            End Try
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
                If encryptor IsNot Nothing Then
                    encryptor.Dispose()
                    encryptor = Nothing
                End If

                If decryptor IsNot Nothing Then
                    decryptor.Dispose()
                    decryptor = Nothing
                End If

                If m_Aes IsNot Nothing Then
                    m_Aes.Dispose()
                    m_Aes = Nothing
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