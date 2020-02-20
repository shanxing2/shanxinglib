
Namespace ShanXingTech
    Partial Public Module ExtensionFunc
#Region "字段区"
        Private m_GBKEncoding As Text.Encoding
        Private ReadOnly Property gbkEncoding() As Text.Encoding
            Get
                If m_GBKEncoding Is Nothing Then
                    m_GBKEncoding = Text.Encoding.GetEncoding("GBK")
                End If

                Return m_GBKEncoding
            End Get
        End Property
#End Region

#Region "常量区"
        Private ReadOnly BROWSER_EMULATION_KEY As String = "Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION"
        Private ReadOnly BufferSize As Integer = 1024 * 80
#End Region
    End Module

End Namespace
