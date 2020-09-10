Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports ShanXingTech
Imports ShanXingTech.Text2

Namespace ShanXingTech
    Partial Public Module ExtensionFunc
        ''' <summary>
        ''' 获取字符串的32位MD5值
        ''' </summary>
        ''' <param name="str">需要执行MD5加密的字符串</param>
        ''' <returns></returns>
        <Extension()>
        Public Function GetMD5Value(ByVal str As String) As String
            ' 空字符串的md5是 d41d8cd98f00b204e9800998ecf8427e
            ' 所以不需要处理空字符串的情况

            If str Is Nothing Then Throw New NullReferenceException(String.Format(My.Resources.NullReference, NameOf(str)))

            Dim retVal As Byte()
            Using md5 = New System.Security.Cryptography.MD5CryptoServiceProvider()
                Try
                    retVal = md5.ComputeHash(Encoding.UTF8.GetBytes(str))
                Catch ex As Exception
                    Logger.WriteLine(ex)
                End Try
            End Using

            Dim funRst = retVal.ToHexString(UpperLowerCase.Lower)

            Return funRst
        End Function

        ''' <summary>
        ''' 获取字符串 <paramref name="sourceString"/> GBK编码的字节数
        ''' </summary>
        ''' <param name="sourceString"></param>
        ''' <returns></returns>
        <Extension()>
        Private Function GetDoubleByteCharCount(ByVal sourceString As String) As Integer
            Dim doubleByteCount As Integer

            ' 如果sb的长度小于1w，那就用同步的方式获取字节数
            If sourceString.Length < 10000 Then
                For Each c In sourceString
                    If gbkEncoding.GetByteCount(CStr(c)) = 2 Then
                        doubleByteCount += 1
                    End If
                Next
            Else
                Parallel.ForEach(
                    Concurrent.Partitioner.Create(0, sourceString.Length),
                    Sub(range)
                        For index = range.Item1 To range.Item2
                            If gbkEncoding.GetByteCount(CStr(sourceString.Chars(index))) = 2 Then
                                Threading.Interlocked.Increment(doubleByteCount)
                            End If
                        Next
                    End Sub)
            End If

            Return doubleByteCount
        End Function

        ''' <summary>
        ''' 获取字符串 <paramref name="sourceString"/> GBK编码的字节数
        ''' </summary>
        ''' <param name="sourceString"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function GetByteCount(ByVal sourceString As String) As Integer
            Dim byteCount As Integer

            ' 如果sb的长度小于1w，那就用同步的方式获取字节数
            If sourceString.Length < 10000 Then
                For i = 0 To sourceString.Length - 1
                    byteCount += gbkEncoding.GetByteCount(CStr(sourceString.Chars(i)))
                Next
            Else
                Parallel.ForEach(
                    Concurrent.Partitioner.Create(0, sourceString.Length),
                    Sub(range)
                        For index = range.Item1 To range.Item2
                            If gbkEncoding.GetByteCount(CStr(sourceString.Chars(index))) = 2 Then
                                Threading.Interlocked.Increment(byteCount)
                                Threading.Interlocked.Increment(byteCount)
                            Else
                                Threading.Interlocked.Increment(byteCount)
                            End If
                        Next
                    End Sub)
            End If

            Return byteCount
        End Function

        ''' <summary>
        ''' 获取字符串 <paramref name="sourceString"/> 的长度，以 “2个英文或者数字占一个长度，1个汉字占一个长度” 为计算标准
        ''' </summary>
        ''' <param name="sourceString"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function GetLengthExt(ByVal sourceString As String) As Double
            Dim doubleByteCount = sourceString.GetDoubleByteCharCount
            Dim byteCount = (sourceString.Length - doubleByteCount) / 2 + doubleByteCount

            Return byteCount
        End Function

        ''' <summary>
        ''' 获取StringBuilder <paramref name="sb"/> 的长度，以 “2个英文或者数字占一个长度，1个汉字占一个长度” 为计算标准
        ''' </summary>
        ''' <param name="sb"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function GetLengthExt(ByVal sb As StringBuilder) As Double
            ' 一个汉字2个字节，1个数字或者英文一个字节
            Dim doubleByteCharCount = sb.GetDoubleByteCharCount
            Dim sbLength = sb.Length

            ' 如果全部由英文或者数字组成，那就直接返回 sb自身的长度
            If sbLength = doubleByteCharCount Then
                Return sbLength
            Else
                ' 多出多少就是有多少个数字或者英文字符
                Dim singleByteCharCount = sb.Length - doubleByteCharCount
                Dim sbTotalLength = singleByteCharCount / 2 + doubleByteCharCount

                Return sbTotalLength
            End If
        End Function

        ''' <summary>
        ''' 获取StringBuilder <paramref name="sb"/> GBK编码的字节数
        ''' </summary>
        ''' <param name="sb"></param>
        ''' <returns></returns>
        <Extension()>
        Private Function GetDoubleByteCharCount(ByVal sb As StringBuilder) As Integer
            Dim doubleByteCount As Integer

            ' 如果sb的长度小于1w，那就用同步的方式获取字节数
            If sb.Length < 10000 Then
                For i = 0 To sb.Length - 1
                    If gbkEncoding.GetByteCount(CStr(sb.Chars(i))) = 2 Then
                        doubleByteCount += 1
                    End If
                Next
            Else
                Parallel.ForEach(
                    Concurrent.Partitioner.Create(0, sb.Length),
                    Sub(range)
                        For index = range.Item1 To range.Item2
                            If gbkEncoding.GetByteCount(CStr(sb.Chars(index))) = 2 Then
                                Threading.Interlocked.Increment(doubleByteCount)
                            End If
                        Next
                    End Sub)
            End If

            Return doubleByteCount
        End Function

        ''' <summary>
        ''' 获取StringBuilder <paramref name="sb"/> GBK编码的字节数
        ''' </summary>
        ''' <param name="sb"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function GetByteCount(ByVal sb As StringBuilder) As Integer
            Dim byteCount As Integer

            ' 如果sb的长度小于1w，那就用同步的方式获取字节数
            If sb.Length < 10000 Then
                For i = 0 To sb.Length - 1
                    byteCount += gbkEncoding.GetByteCount(CStr(sb.Chars(i)))
                Next
            Else
                Parallel.ForEach(
                    Concurrent.Partitioner.Create(0, sb.Length),
                    Sub(range)
                        For index = range.Item1 To range.Item2
                            Threading.Interlocked.Increment(gbkEncoding.GetByteCount(CStr(sb.Chars(index))))
                        Next
                    End Sub)
            End If


            Return byteCount
        End Function

        ''' <summary>
        ''' 返回一个新字符串，该字符串通过在此字符串中的字符右侧填充空格来达到指定的总长度，从而使这些字符左对齐。
        ''' 适用于字符串中包含中英文的情况
        ''' </summary>
        ''' <param name="sourceString">原字符串</param>
        ''' <param name="totalByteWidth">需要构造的字符串总长度</param>
        ''' <param name="paddingChar">填充的字符，默认为 一个空格</param>
        ''' <returns></returns>
        <Extension()>
        Public Function PadRightByByte(ByVal sourceString As String, ByVal totalByteWidth As Integer， paddingChar As Char) As String
            Dim totalByteCount = gbkEncoding.GetByteCount(sourceString)
            If totalByteWidth < totalByteCount Then Return sourceString
            Dim doubleByteCount = sourceString.GetDoubleByteCharCount

            Return sourceString.PadRight(totalByteWidth - doubleByteCount, paddingChar)
        End Function


        ''' <summary>
        ''' 返回一个新字符串，该字符串通过在此字符串中的字符右侧填充空格来达到指定的总长度，从而使这些字符左对齐。
        ''' 适用于字符串中包含中英文的情况
        ''' </summary>
        ''' <param name="sourceString">原字符串</param>
        ''' <param name="totalByteWidth">需要构造的字符串总长度</param>
        ''' <returns></returns>
        <Extension()>
        Public Function PadRightByByte(ByVal sourceString As String, ByVal totalByteWidth As Integer) As String
            Return sourceString.PadRightByByte(totalByteWidth, " "c)
        End Function

        ''' <summary>
        ''' 返回一个新字符串，该字符串通过在此字符串中的字符左侧填充空格来达到指定的总长度，从而使这些字符右对齐。
        ''' 适用于字符串中包含中英文的情况
        ''' </summary>
        ''' <param name="sourceString">原字符串</param>
        ''' <param name="totalByteWidth">需要构造的字符串总长度</param>
        ''' <param name="paddingChar">填充的字符</param>
        ''' <returns></returns>
        <Extension()>
        Public Function PadLeftByByte(ByVal sourceString As String, ByVal totalByteWidth As Integer， ByVal paddingChar As Char) As String
            Dim totalByteCount = gbkEncoding.GetByteCount(sourceString)
            If totalByteWidth < totalByteCount Then Return sourceString
            Dim doubleByteCount = sourceString.GetDoubleByteCharCount

            Return sourceString.PadLeft(totalByteWidth - doubleByteCount, paddingChar)
        End Function

        ''' <summary>
        ''' 返回一个新字符串，该字符串通过在此字符串中的字符左侧填充空格来达到指定的总长度，从而使这些字符右对齐。
        ''' 适用于字符串中包含中英文的情况。填充字符默认为 一个空格
        ''' </summary>
        ''' <param name="sourceString">原字符串</param>
        ''' <param name="totalByteWidth">需要构造的字符串总长度</param>
        ''' <returns></returns>
        <Extension()>
        Public Function PadLeftByByte(ByVal sourceString As String, ByVal totalByteWidth As Integer) As String
            Return PadLeftByByte(sourceString, totalByteWidth, " "c)
        End Function

        ''' <summary>
        ''' 尝试去除 <paramref name="sourceString"/> 中的换行符
        ''' </summary>
        ''' <param name="sourceString"></param>
        <Extension()>
        Public Function TryRemoveNewLine(ByVal sourceString As String) As String
            If sourceString.IsNullOrEmpty Then Return sourceString

            Dim lastpos As Integer
            Dim i As Integer
            Dim sb = New StringBuilder(sourceString.Length)
            While i < sourceString.Length
                If sourceString(i) = ControlChars.Cr OrElse sourceString(i) = ControlChars.Lf Then
                    ' 找到第一个 cr 或者是 lf的位置,然后把前面的字符加到sb
                    sb.Append(sourceString, lastpos, i)
                    Do
                        i += 1

                        ' 从上一个位置开始，找到到下一个不为 cr 或者 不为 lf的字符
                        While i < sourceString.Length AndAlso
                        (sourceString(i) = ControlChars.Cr OrElse sourceString(i) = ControlChars.Lf)
                            i += 1
                        End While
                        lastpos = i

                        i += 1
                        ' 从上一个不为 cr 或者 不为 lf的字符开始，找到下一个为 cr 或者 为 lf的字符,然后把中间的字符加到sb
                        While i < sourceString.Length AndAlso
                        sourceString(i) <> ControlChars.Cr AndAlso
                        sourceString(i) <> ControlChars.Lf
                            i += 1
                        End While

                        If i >= sourceString.Length Then Exit While
                        sb.Append(sourceString, lastpos, i - lastpos)
                    Loop While i < sourceString.Length
                End If

                i += 1
            End While

            ' 可能源字符串没有包含换行符,因此里面没有任何数据，直接返回源字符串就好
            If sb.Length = 0 Then
                sb.ToString()
            Else
                sourceString = sb.ToString
            End If

            Return sourceString
        End Function


		''' <summary>
		''' unicode转换成中文
		''' 如：\u5929\u5929\u7279\u4EF7 转换之后是 天天特价,https:\/\/live.bilibili.com 转换之后是 https://live.bilibili.com
		'''<para>详细请看函数 <seealso cref="Regex.Unescape(String)"/></para>
		''' </summary>
		''' <param name="sourceString">原字符串</param>
		''' <returns></returns>
		<Extension()>
		Public Function TryUnescape(ByVal sourceString As String) As String
			If sourceString.IsNullOrEmpty Then Return sourceString

			sourceString = Regex.Unescape(sourceString)
			Return sourceString
		End Function

		''' <summary>
		''' 原始字符串转换为16进制字符串(unicode)
		''' 如：天天特价 ,传入分隔符\u转换之后是 \u5929\u5929\u7279\u4EF7
		''' </summary>
		''' <param name="sourceString">原字符串</param>
		''' <param name="lUCase">生成大写或者小写形式的16进制字符串</param>
		''' <param name="containDoubleByte"><paramref name="sourceString"/>是否包含双字节字符串</param>
		''' <param name="separator">每个字符转换之后需要带上的分隔符，可为空</param>
		''' <returns></returns>
		<Extension()>
		Public Function TryToUnicode(ByVal sourceString As String, ByVal lUCase As UpperLowerCase, ByVal containDoubleByte As Boolean, ByVal separator As String) As String
			If String.IsNullOrEmpty(sourceString) Then Return sourceString

			Dim fotmat As String = "x"
			If lUCase = UpperLowerCase.Upper Then
				fotmat = "X"
			End If

			Dim subStringLength As Integer
			If containDoubleByte Then
				fotmat += "4"
				subStringLength = 4
			Else
				subStringLength = 2
				fotmat += "2"
			End If

            Dim sb = New StringBuilder(sourceString.Length * subStringLength)

            Dim i As Integer
			Dim needSeparator = Not String.IsNullOrEmpty(separator)
			While i < sourceString.Length
				If needSeparator Then
					sb.Append(separator)
				End If
				sb.Append(Convert.ToUInt32(sourceString(i)).ToString(fotmat))
				i += 1
			End While

            Return sb.ToString
        End Function

		''' <summary>
		''' 判断是否包含汉字字符串
		''' 有效范围 [\u4e00-\u9fa5]
		''' </summary>
		''' <param name="sourceString"></param>
		''' <returns></returns>
		<Extension()>
        Public Function ContainChineseChar(ByVal sourceString As String) As Boolean
            If sourceString.IsNullOrEmpty Then Return False

			Return Regex.IsMatch(sourceString, "[\u4e00-\u9fa5]")
		End Function

        ''' <summary>
        ''' 使用 <paramref name="encoding"/> 编码方式 将传入的字符串编码。默认的编码结果为小写形式。
        ''' </summary>
        ''' <param name="sourceString"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function UrlEncode(ByVal sourceString As String, ByVal encoding As Encoding) As String
            Return UrlEncode(sourceString, encoding, UpperLowerCase.Lower)
        End Function

        ''' <summary>
        ''' 使用 <paramref name="encoding"/> 编码方式 将传入的字符串编码
        ''' </summary>
        ''' <param name="sourceString"></param>
        ''' <param name="encoding"></param>
        ''' <param name="lUCase">编码结果为大写还是小写</param>
        ''' <returns></returns>
        <Extension()>
        Public Function UrlEncode(ByVal sourceString As String, ByVal encoding As Encoding, ByVal lUCase As UpperLowerCase) As String
            If sourceString.IsNullOrEmpty Then Return String.Empty

            If lUCase = UpperLowerCase.Lower Then
                Return Web.HttpUtility.UrlEncode(sourceString, encoding)
            Else
                Dim funcRst = Web.HttpUtility.UrlEncode(sourceString, encoding)
                Dim sb = New StringBuilder(sourceString.Length * 3)
                For Each c In sourceString
                    Dim ue = Web.HttpUtility.UrlEncode(c, encoding)
                    If c = ue Then
                        sb.Append(c)
                    Else
                        sb.Append(ue.ToUpper)
                    End If
                Next

                Return sb.ToString
            End If
        End Function

        '''' <summary>
        '''' 使用 <paramref name="gbkEncoding"/> 编码方式 将传入的字符串解码
        '''' </summary>
        '''' <param name="sb"></param>
        '''' <param name="gbkEncoding"></param>
        '''' <returns></returns>
        '<Extension()>
        'Public Function UrlDecode(ByVal sb As String, ByVal gbkEncoding As Encoding) As String
        '    Dim funcRst As String = String.Empty

        '    ' 因为需要解码的字符串里面可能含有未编码的字符串，所以用这种算法确定字节数组的大小不太精确
        '    ' 最终字节数组的大小可能大于实际编码之后的字节数组大小
        '    Dim bLen = sb.Length \ 3 + (sb.Length Mod 3) - 1
        '    Dim bStr(bLen) As Byte

        '    Dim index As Integer
        '    Dim bIndex As Integer
        '    Dim hexStr As String = String.Empty
        '    Dim chChar As Char

        '    ' 未编码的字符跟编码过的字符需要分开处理
        '    While index < sb.Length
        '        chChar = sb.Chars(index)
        '        If chChar = "%"c Then
        '            hexStr = String.Concat("&h".TryGetIntern & sb.Substring(index + 1, 2))
        '            ' 因为是专门用来处理url解码的，所以不需要在此判断是否 hexStr 时候为数字 以提高效率
        '            bStr(bIndex) = CByte(hexStr)
        '            bIndex += 1
        '            index += 3
        '        Else
        '            ' 如果是没有经过编码的 需要再编码一次
        '            ' 同时 bStr 数组长度需要增加 N 个字节长度 以保证 bStr 数组有足够的长度容纳增加的字节
        '            ' 然后复制到 bStr 数组里面
        '            ' N 根据字符类型来确定
        '            Dim tempByte = gbkEncoding.GetBytes(chChar)
        '            ReDim Preserve bStr(bStr.Length - 1 + tempByte.Length)
        '            tempByte.CopyTo(bStr, bIndex)

        '            bIndex += tempByte.Length
        '            index += tempByte.Length
        '        End If
        '    End While

        '    funcRst = gbkEncoding.GetString(bStr)

        '    Return funcRst
        'End Function

        ''' <summary>
        ''' 使用name将传入的字符串解码
        ''' </summary>
        ''' <param name="sourceString"></param>
        ''' <param name="encoding"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function UrlDecode(ByVal sourceString As String, ByVal encoding As Encoding) As String
            Dim funcRst = Web.HttpUtility.UrlDecode(sourceString, encoding)

            Return funcRst
        End Function

        ''' <summary>
        ''' 将字符串中包含的可转义字符转义，比如 '·' 转换为 '&amp;middot;'
        ''' </summary>
        ''' <param name="sourceString">原字符串</param>
        ''' <returns></returns>
        <Extension()>
        Public Function HtmlEncode(ByVal sourceString As String) As String
            Dim funcRst = Web.HttpUtility.HtmlEncode(sourceString)

            Return funcRst
        End Function

        ''' <summary>
        ''' 将字符串中包含的转义字符还原，比如 '&amp;middot;' 还原为 '·'
        ''' </summary>
        ''' <param name="sourceString">原字符串</param>
        ''' <returns></returns>
        <Extension()>
        Public Function HtmlDecode(ByVal sourceString As String) As String
            Dim funcRst As String = String.Empty

            funcRst = Web.HttpUtility.HtmlDecode(sourceString)

            Return funcRst
        End Function

        ''' <summary>
        ''' 尝试从暂存池中获取该字符串的相同引用，并将它分配给变量，如果没有则向暂存池添加对 <paramref name="str"/> 的引用，然后返回该引用。
        ''' 如果为Nothing或者空，则直接返回 <paramref name="str"/>
        ''' 如果应用程序经常对字符串进行区分大小写的，序号式的比较，或者事先知道许多字符串对象都有相同的值，就可以利用CLR的字符串留用机制来显著提高性能。
        ''' </summary>
        ''' <param name="str">需要从暂存池检索的字符串</param>
        ''' <returns></returns>
        <Extension()>
        Public Function TryGetIntern(ByVal str As String) As String
            If String.IsNullOrEmpty(str) Then Return str

            Return String.Intern(str)
        End Function

        ''' <summary>
        ''' String.IsNullOrEmpty的封装
        ''' </summary>
        ''' <param name="str"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function IsNullOrEmpty(ByVal str As String) As Boolean
            Return str Is Nothing OrElse str.Length = 0
        End Function

        ''' <summary>
        ''' 报告指定字符串在此实例中的第一个匹配项的从零开始的索引
        ''' </summary>
        ''' <param name="sb"></param>
        ''' <param name="findString">要搜寻的字符串</param>
        ''' <returns></returns>
        <Extension()>
        Public Function IndexOf(ByVal sb As StringBuilder, ByVal findString As String) As Integer
#Region "NothingOrEmpty"
            ' 如果字符串为Nothing或者空，则 Return -1
            If String.IsNullOrEmpty(findString) Then
                'Debug.Print(Logger.MakeDebugString("findString长度为0或者为Nothing,不需要查找"))
                Return -1
            End If
#End Region

            Dim findStringlength = findString.Length
            Dim sblength = sb.Length

#Region "当 sb 长度为0时使用此方法"
            If sblength = 0 Then
                'Debug.Print(Logger.MakeDebugString("sb 长度为0,不需要查找"))
                Return -1
            End If
#End Region

#Region "当 findString 长度 大于 sb 长度时使用此方法"
            If findStringlength > sblength Then
                'Debug.Print(Logger.MakeDebugString("findString 长度 大于 sb 长度,不需要查找"))
                Return -1
            End If
#End Region

#Region "当 findString 长度为1时使用此方法查找"
            If findStringlength = 1 Then
                For sbIndex = 0 To sblength - 1
                    If findString.Equals(sb.Chars(sbIndex), StringComparison.OrdinalIgnoreCase) Then
                        Return sbIndex
                    End If
                Next

                'Debug.Print(Logger.MakeDebugString( $"   全部查找完毕，没有找到 {findString}"))
                Return -1
            End If
#End Region

#Region "当 findString 长度大于等于2时使用此方法查找"
            Dim startIndex As Integer
            While startIndex + findStringlength <= sb.Length
                If sb(startIndex) = findString(0) Then
                    If ScanRight(sb, findString, startIndex) Then
                        Return startIndex
                    Else
                        startIndex += 1
                    End If
                Else
                    startIndex += 1
                End If
            End While
#End Region

            Return -1
        End Function

        ''' <summary>
        ''' 从 <paramref name="sb"/> 的 <paramref name="startIndex"/> 位置开始或者 <paramref name="length"/> 长度字符串
        ''' </summary>
        ''' <param name="startIndex">提取数据的开始位置(从0算起)</param>
        ''' <param name="length">需要提取数据的长度</param>
        ''' <returns></returns>
        <Extension()>
        Public Function Substring(ByRef sb As StringBuilder, ByVal startIndex As Integer, ByVal length As Integer) As String
            Dim funcRst = String.Empty

            If sb.Length = 0 Then Return funcRst
            If length = 1 Then Return sb.Chars(startIndex)

            Dim tempSb = New StringBuilder(length)

            For charIndex = startIndex To startIndex + length - 1
                tempSb.Append(sb.Chars(charIndex))
            Next

            Return tempSb.ToString
        End Function


        ''' <summary>
        ''' 移除 <paramref name="sb"/> 中的所有 <paramref name="findChar"/>
        ''' </summary>
        ''' <param name="sb"></param>
        ''' <param name="findChar">要移除的字符，可为 Char.MinValue 或者 Nothing</param>
        ''' <returns></returns>
        <Extension()>
        Public Function TryRemove(ByRef sb As StringBuilder, ByVal findChar As Char) As Boolean
            Dim buidlerIndex = sb.Length - 1
            If buidlerIndex < 0 Then Return False

            While buidlerIndex > 0
                If findChar = sb(buidlerIndex) Then
                    sb.Remove(buidlerIndex, 1)
                End If
                buidlerIndex -= 1
            End While

            Return True
        End Function

        ''' <summary>
        ''' 移除 <paramref name="sb"/> 中的所有 <paramref name="findString"/>
        ''' </summary>
        ''' <param name="sb"></param>
        ''' <param name="findString">要移除的字符串，可为 Char.MinValue 或者 Nothing</param>
        ''' <returns></returns>
        <Extension()>
        Public Function TryRemove(ByRef sb As StringBuilder, ByVal findString As String) As Boolean
            If findString Is Nothing Or findString.Length = 0 Then Return False

            Dim buidlerIndex = sb.Length - 1
            If buidlerIndex < 0 Then Return False

            Dim findStringLength = findString.Length

            While buidlerIndex > 0
                If findString(0) = sb(buidlerIndex) AndAlso
                    ScanRight(sb, findString, buidlerIndex) Then
                    sb.Remove(buidlerIndex, findStringLength)
                End If
                buidlerIndex -= 1
            End While

            Return True
        End Function

        ''' <summary>
        ''' 去掉 <paramref name="fileName"/> 中可能会存在的不能用做文件名的特殊字符
        ''' </summary>
        ''' <param name="fileName"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function TryRemoveInvalidFileNameChar(ByRef fileName As String) As Boolean
            If fileName Is Nothing OrElse fileName.Length = 0 Then Return False

            Dim sb = New StringBuilder(fileName.Length)
            Dim invalidFileNameChars = IO.Path.GetInvalidFileNameChars
            For i = 0 To fileName.Length - 1
                If invalidFileNameChars.AsParallel.Contains(fileName(i)) Then
                    Continue For
                End If
                sb.Append(fileName(i))
            Next
            fileName = sb.ToString

            Return True
        End Function

        ''' <summary>
        ''' 返回一个布尔值，改值指示指定的字符串是否出现在此实例中
        ''' </summary>
        ''' <param name="sb"></param>
        ''' <param name="findString">要搜寻的字符串</param>
        ''' <returns></returns>
        <Extension()>
        Public Function Contains(ByVal sb As StringBuilder, ByVal findString As String) As Boolean
            Return sb.IndexOf(findString) > -1
        End Function

        ''' <summary>
        ''' StreamReader实例的ReadToEnd增强版 比自带的高效
        ''' </summary>
        ''' <param name="stream"></param>
        ''' <param name="encoding"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ReadToEndExt(ByVal stream As Stream, ByVal encoding As System.Text.Encoding) As StringBuilder
            Dim funcRst = New StringBuilder(BufferSize)
            If Not stream.CanRead Then Return funcRst

            ' StreamReader实例的Read方法默认的缓存大小为4K=4096字节
            ' 如果构造流时未指定内部缓冲区大小， 则缓冲区的默认大小为 4 KB（4096 字节）。 
            ' 第三个参数detectEncodingFromByteOrderMarks 为false,表示不自动识别编码，由用户传入
            Using reader As New System.IO.StreamReader(stream, encoding, False, BufferSize)
                ' 使用 Read 方法时，较为高效的方法是使用与流的内部缓冲区大小一致的缓冲区，其中内部缓冲区设置为您想要的块大小，并始终读取小于此块大小的内容。
                Dim buffer As Char() = New Char(BufferSize - 1) {}
                ' 注意 
                ' 如果这一次返回的数据长度小于readLen
                ' 则buffer中readLen前面的数据为本次需要读取的数据
                ' 后面的数据为上一次的数据
                ' 不过，有个很奇怪的事儿，有时候reader里面一次只有很少的数据可以读出来，实际上，数据应该很多的
                ' 比如stream里面应该有16576字节的数据，但是就算给的buffer大于16576字节，一样也只能读出3154（貌似是随机）个字节
                ' 必须多读几次
                ' 所以，这个时候用Peek方法是不能正常读取的，只能像下面这样判断Read方法的返回值
                Dim readCharLen = reader.Read(buffer, 0, BufferSize)
                While readCharLen > 0
                    ' 根据上面的注释 只从buffer里面读取本次读到的数据
                    funcRst.Append(buffer, 0, readCharLen)

                    readCharLen = reader.Read(buffer, 0, BufferSize)
                End While
            End Using

            Return funcRst
        End Function

        ''' <summary>
        ''' StreamReader实例的ReadToEnd增强版 比自带的高效
        ''' </summary>
        ''' <param name="stream"></param>
        ''' <param name="encoding"></param>
        ''' <returns></returns>
        <Extension()>
        Public Async Function ReadToEndExtAsync(ByVal stream As Stream, ByVal encoding As System.Text.Encoding) As Task(Of StringBuilder)
            Dim funcRst = New StringBuilder(BufferSize)
            If Not stream.CanRead Then Return funcRst

            ' StreamReader实例的Read方法默认的缓存大小为4K=4096字节
            ' 如果构造流时未指定内部缓冲区大小， 则缓冲区的默认大小为 4 KB（4096 字节）。 
            ' 第三个参数detectEncodingFromByteOrderMarks 为false,表示不自动识别编码，由用户传入
            Using reader As New System.IO.StreamReader(stream, encoding, False, BufferSize)
                ' 使用 Read 方法时，较为高效的方法是使用与流的内部缓冲区大小一致的缓冲区，其中内部缓冲区设置为您想要的块大小，并始终读取小于此块大小的内容。
                Dim buffer As Char() = New Char(BufferSize - 1) {}
                ' 注意 
                ' 如果这一次返回的数据长度小于readLen
                ' 则buffer中readLen前面的数据为本次需要读取的数据
                ' 后面的数据为上一次的数据
                ' 不过，有个很奇怪的事儿，有时候reader里面一次只有很少的数据可以读出来，实际上，数据应该很多的
                ' 比如stream里面应该有16576字节的数据，但是就算给的buffer大于16576字节，一样也只能读出3154（貌似是随机）个字节
                ' 必须多读几次
                ' 所以，这个时候用Peek方法是不能正常读取的，只能像下面这样判断Read方法的返回值
                Dim readCharLen = Await reader.ReadAsync(buffer, 0, BufferSize)
                While readCharLen > 0
                    ' 根据上面的注释 只从buffer里面读取本次读到的数据
                    funcRst.Append(buffer, 0, readCharLen)

                    readCharLen = Await reader.ReadAsync(buffer, 0, BufferSize)
                End While
            End Using

            Return funcRst
        End Function

        ''' <summary>
        ''' 检查某个字符串是否全部由标点符号或者数学符号或者货币符号等字符组成
        ''' </summary>
        ''' <param name="str">需要比较的字符串</param>
        ''' <returns></returns>
        <Extension()>
        Public Function EqualPunctuationOrSymbol(ByVal str As String) As Boolean
            Dim match = Regex.Match(str, "^[\p{P}\p{S}]+$", RegularExpressions.RegexOptions.IgnoreCase Or RegexOptions.Compiled)
            Return match.Success
        End Function

        ''' <summary>
        ''' 获取第一个匹配组的值
        ''' </summary>
        ''' <param name="input"></param>
        ''' <param name="pattern"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function GetFirstMatchValue(ByVal input As String, ByVal pattern As String, ByVal options As RegexOptions) As String
            If input.IsNullOrEmpty Then
                Return String.Empty
            End If

            Dim match = Regex.Match(input, pattern, options)
            If match.Success Then
                Dim funcRst = match.Groups(1).Value
                Return funcRst
            Else
                Return String.Empty
            End If
        End Function

        ''' <summary>
        ''' 获取第一个匹配组的值,匹配选项默认为 <see cref="RegexOptions.Compiled"/> Or <see cref="RegexOptions.IgnoreCase"/>
        ''' </summary>
        ''' <param name="input"></param>
        ''' <param name="pattern"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function GetFirstMatchValue(ByVal input As String, ByVal pattern As String) As String
            Return GetFirstMatchValue(input, pattern, RegexOptions.IgnoreCase Or RegexOptions.Compiled)
        End Function


        ''' <summary>
        ''' 正则检测字符串 <paramref name="input"/> 是否匹配 <paramref name="pattern"/>
        ''' </summary>
        ''' <param name="input">待检测字符串</param>
        ''' <param name="pattern">匹配表达式</param>
        ''' <param name="options">匹配选项</param>
        ''' <returns></returns>
        <Extension()>
        Public Function IsMatch(ByVal input As String, ByVal pattern As String, ByVal options As RegexOptions) As Boolean
            If input.IsNullOrEmpty Then
                Return False
            End If

            Return Regex.IsMatch(input, pattern, options)
        End Function

        ''' <summary>
        ''' 正则检测字符串 <paramref name="input"/> 是否匹配 <paramref name="pattern"/>,匹配选项默认为 <see cref="RegexOptions.Compiled"/> Or <see cref="RegexOptions.IgnoreCase"/>
        ''' </summary>
        ''' <param name="input">待检测字符串</param>
        ''' <param name="pattern">匹配表达式</param>
        ''' <returns></returns>
        <Extension()>
        Public Function IsMatch(ByVal input As String, ByVal pattern As String) As Boolean
            Return IsMatch(input, pattern, RegexOptions.IgnoreCase Or RegexOptions.Compiled)
        End Function

        ''' <summary>
        ''' 从包含键值对的字符串中获取键 <paramref name="key"/> 的值
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="key"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function GetQueryParamValue(ByVal url As String, ByVal key As Char) As String
            If url.IsNullOrEmpty Then
                Throw New ArgumentException(String.Format(My.Resources.NullReference, NameOf(url)))
            End If

            Dim urlLength = url.Length
            Dim i As Integer
            Dim valueIndex As Integer = -1
            Dim nextKeyIndex As Integer = -1

            Dim sb = New StringBuilder(urlLength)
            While i < urlLength
                If IsCharEqual(key, url(i), True) AndAlso
                    "="c = url(i + 1) AndAlso
                    ("&"c = url(i - 1) OrElse "?"c = url(i - 1)) Then
                    valueIndex = i + 1 + 1
                    i += 1 + 1
                    Continue While
                End If

                If valueIndex > -1 Then
                    If "&"c = url(i) Then
                        nextKeyIndex = i
                        Exit While
                    Else
                        sb.Append(url(i))
                    End If
                End If

                i += 1
            End While

            Dim keyword = sb.ToString
            Return keyword
        End Function


        ''' <summary>
        ''' 从包含键值对的字符串中获取键 <paramref name="key"/> 的值
        ''' </summary>
        ''' <param name="url"></param>
        ''' <param name="key"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function GetQueryParamValue(ByVal url As String, ByVal key As String) As String
            If url.IsNullOrEmpty Then
                Throw New ArgumentException(String.Format(My.Resources.NullReference, NameOf(url)))
            End If
            If key.IsNullOrEmpty Then
                Throw New ArgumentException(String.Format(My.Resources.NullReference, NameOf(key)))
            End If

            Dim urlLength = url.Length
            Dim i As Integer
            Dim valueIndex As Integer = -1
            Dim nextKeyIndex As Integer = -1
            ' "(\w+)=([\w+|-|/|+|=]*)"
            Dim sb = New StringBuilder(urlLength)
            While i < urlLength
                ' 是 key的第一个字符，并且往后都一样，并且前面一个字符是&或者？
                If IsCharEqual(key(0), url(i), True) AndAlso
                    ScanRight(url, key, i) AndAlso
                    "="c = url(i + key.Length) AndAlso
                    (i = 0 OrElse ("&"c = url(i - 1) OrElse "?"c = url(i - 1))) Then
                    valueIndex = i + key.Length + 1
                    i = valueIndex
                    Continue While
                End If

                If valueIndex > -1 Then
                    If "&"c = url(i) Then
                        nextKeyIndex = i
                        Exit While
                    Else
                        sb.Append(url(i))
                    End If
                End If

                i += 1
            End While

            If sb.Length = 0 Then Return String.Empty

            Dim keyword = sb.ToString
            Return keyword
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="firstChar"></param>
        ''' <param name="secondChar"></param>
        ''' <param name="ignoreCase">忽略大小写</param>
        ''' <returns></returns>
        Private Function IsCharEqual(ByVal firstChar As Char, ByVal secondChar As Char, ByVal ignoreCase As Boolean) As Boolean
            If ignoreCase Then
                ' 大小写相同，或者是经过大小写转换之后相同
                Return firstChar = secondChar OrElse
                    (Char.IsLetter(firstChar) AndAlso Char.IsLetter(secondChar) AndAlso (firstChar.CompareTo(secondChar) = 32 OrElse
                    firstChar.CompareTo(secondChar) = -32))
            Else
                Return firstChar = secondChar
            End If
        End Function

        ''' <summary>
        ''' 从 <paramref name="compareString"/> 的索引 <paramref name="startIndex"/> 处开始向右检查是否包含 <paramref name="findString"/>
        ''' </summary>
        ''' <param name="compareString"></param>
        ''' <param name="findString"></param>
        ''' <param name="startIndex"></param>
        ''' <param name="ignoreCase">忽略大小写</param>
        ''' <returns></returns>
        <Extension()>
        Public Function ScanRight(ByVal compareString As String, ByVal findString As String, ByVal startIndex As Integer, Optional ByVal ignoreCase As Boolean = True) As Boolean
            If findString.Length > compareString.Length - startIndex Then
                Return False
            End If

            ' 特殊处理 1-2个长度的
            If findString.Length = 1 Then
                Return True
            ElseIf findString.Length = 2 Then
                Return IsCharEqual(findString(1), compareString(startIndex), ignoreCase)
            End If

            ' 循环处理3个长度以上的
            Dim i As Integer
            While i < findString.Length
                If IsCharEqual(findString(i), compareString(startIndex), ignoreCase) Then
                    startIndex += 1
                    i += 1
                Else
                    Return False
                End If
            End While

            Return True
        End Function

        ''' <summary>
        ''' 从 <paramref name="compareBuilder"/> 的索引 <paramref name="startIndex"/> 处开始向右检查是否包含 <paramref name="findString"/>
        ''' </summary>
        ''' <param name="compareBuilder"></param>
        ''' <param name="findString"></param>
        ''' <param name="startIndex"></param>
        ''' <param name="ignoreCase">忽略大小写</param>
        ''' <returns></returns>
        <Extension()>
        Public Function ScanRight(ByVal compareBuilder As StringBuilder, ByVal findString As String, ByVal startIndex As Integer, Optional ByVal ignoreCase As Boolean = True) As Boolean
            If findString.Length > compareBuilder.Length - startIndex Then
                Return False
            End If

            ' 特殊处理 1-2个长度的
            If findString.Length = 1 Then
                Return True
            ElseIf findString.Length = 2 Then
                Return IsCharEqual(findString(1), compareBuilder(startIndex + 1), ignoreCase)
            End If

            ' 循环处理3个长度以上的
            Dim i As Integer
            While i < findString.Length
                If IsCharEqual(findString(i), compareBuilder(startIndex), ignoreCase) Then
                    startIndex += 1
                    i += 1
                Else
                    Return False
                End If
            End While

            Return True
        End Function

        ''' <summary>
        ''' 转义 字符串中包含 Like操作符(<seealso cref="DataTable.Select(String)"/>) 支持的通配符。
        ''' Operator LIKE is used to include only values that match a pattern with wildcards. Wildcard character is * or %, it can be at the beginning of a pattern '*value', at the end 'value*', or at both '*value*'. Wildcard in the middle of a patern 'va*lue' is not allowed.
        ''' If a pattern in a LIKE clause contains any of these special characters * % [ ], those characters must be escaped in brackets [ ] like this [*], [%], [[] or []].
        ''' see https://www.csharp-examples.net/dataview-rowfilter/
        ''' </summary>
        ''' <param name="valueWithWildcards"></param>
        <Extension()>
        Public Function TryEscapeLikeWildcards(ByVal valueWithWildcards As String) As String
            If valueWithWildcards.IsNullOrEmpty Then Return String.Empty

            Dim sb = New StringBuilder(360)

            Dim i As Integer
            While i < valueWithWildcards.Length
                Dim c As Char = valueWithWildcards(i)
                If c = "*"c OrElse c = "%"c OrElse c = "["c OrElse c = "]"c Then
                    sb.Append("[").Append(c).Append("]")
                ElseIf c = "'"c Then
                    sb.Append("''")
                Else
                    sb.Append(c)
                End If
                i += 1
            End While

            Return sb.ToString
        End Function
    End Module
End Namespace
