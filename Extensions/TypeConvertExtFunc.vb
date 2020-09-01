Imports System.Globalization
Imports System.IO
Imports System.IO.Compression
Imports System.Runtime.CompilerServices
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Text
Imports System.Threading.Tasks
Imports System.Web
Imports System.Web.Script.Serialization

Imports ShanXingTech.Text2

Namespace ShanXingTech
    Partial Public Module ExtensionFunc
#Region "属性区"
        ''' <summary>
        ''' 序列化器，可用于对象实例（如某某实体）的序列化或者字符串（如Json字符串）到实体的反序列化
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property MSJsSerializer As JavaScriptSerializer
            Get
                Return If(MSJsSerializer, New JavaScriptSerializer)
            End Get
        End Property

        ''' <summary>
        ''' 合法的16进制字符表
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property HexChars As Char()
            Get
                Return If(HexChars, "0123456789abcdefABCDEF".ToCharArray)
            End Get
        End Property
#End Region
#Region "常量区"

#End Region

        ''' <summary>
        ''' 16进制字符串转换为直接数组
        ''' </summary>
        ''' <param name="hexString"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function HexStringToInteger(ByVal hexString As String) As Integer
            If String.IsNullOrEmpty(hexString) Then
                Throw New NullReferenceException(String.Format(My.Resources.NullReference, NameOf(hexString)))
            End If

            Return CInt(hexString.Insert(0, "&H"))
        End Function

        ''' <summary>
        ''' 16进制字符串转换为字节数组
        ''' </summary>
        ''' <param name="hexString"></param>
        ''' <param name="containDoubleByte"><paramref name="hexString"/>源是否包含双字节字符</param>
        ''' <returns></returns>
        <Extension()>
        Public Function HexStringToBytes(ByVal hexString As String, ByVal containDoubleByte As Boolean) As Byte()
            ' 去掉所有换行符
            hexString = hexString.TryRemoveNewLine()

            Dim subStringLength = If(containDoubleByte, 4, 2)

            Dim decimalBytes As Byte() = New Byte(hexString.Length \ subStringLength - 1) {}
            Dim hIndex As Integer = 0
            For i = 0 To (hexString.Length - subStringLength) Step subStringLength
                'decimalBytes(hIndex) =  CByte(hexString.Substring(i, subStringLength).Insert(0, "&H"))
                decimalBytes(hIndex) = Convert.ToByte(hexString.Substring(i, subStringLength), 16)
                hIndex += 1
            Next

            Return decimalBytes
        End Function

        ''' <summary>
        ''' 16进制字符串转换为直接数组。默认把<paramref name="hexString"/>当做全部由单字节字符组成的字符串来处理
        ''' </summary>
        ''' <param name="hexString"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function HexStringToBytes(ByVal hexString As String) As Byte()
            Return HexStringToBytes(hexString, False)
        End Function

        ''' <summary>
        ''' 原始字符串转换为16进制字符串
        ''' </summary>
        ''' <param name="sourceString"></param>
        ''' <param name="lUCase">生成大写或者小写形式的16进制字符串</param>
        ''' <param name="containDoubleByte"><paramref name="sourceString"/>是否包含双字节字符串</param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToHexString(ByVal sourceString As String, ByVal lUCase As UpperLowerCase, ByVal containDoubleByte As Boolean) As String
            Return TryToUnicode(sourceString, lUCase, containDoubleByte, String.Empty)
        End Function

        ''' <summary>
        ''' 原始字符串转换为16进制字符串。默认把<paramref name="sourceString"/>当做全部由单字节字符组成的字符串来处理
        ''' </summary>
        ''' <param name="sourceString"></param>
        ''' <param name="lUCase">生成大写或者小写形式的16进制字符串</param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToHexString(ByVal sourceString As String, ByVal lUCase As UpperLowerCase) As String
            Return ToHexString(sourceString, lUCase, False)
        End Function

        ''' <summary>
        ''' 16进制字符串转换为原始字符串，如果包含非16进制字符将截断返回
        ''' </summary>
        ''' <param name="sourceString"></param>
        ''' <param name="containDoubleByte"><paramref name="sourceString"/>是否包含双字节字符串</param>
        ''' <returns></returns>
        <Extension()>
        Public Function FromHexString(ByVal sourceString As String, ByVal containDoubleByte As Boolean) As String
            If sourceString.IsNullOrEmpty Then Return String.Empty

            Dim byteLength = If(containDoubleByte, 4, 2)
            Dim sb = StringBuilderCache.AcquireSuper(sourceString.Length \ byteLength)

            Try
                For i = 0 To (sourceString.Length - byteLength) Step byteLength
                    Dim subString = sourceString.Substring(i, byteLength)
                    If Not ValidateHexChar(subString) Then
                        Exit For
                    End If
                    sb.Append(Convert.ToChar(Integer.Parse(subString, NumberStyles.HexNumber)))
                Next
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try

            Return StringBuilderCache.GetStringAndReleaseBuilderSuper(sb)
        End Function

        ''' <summary>
        ''' 验证 <paramref name="data"/> 中所有的字符是否为合法的十六进制字符
        ''' </summary>
        ''' <param name="data"></param>
        ''' <returns></returns>
        Private Function ValidateHexChar(ByVal data As String) As Boolean
            Dim i = 0
            While i < data.Length
                If Not HexChars.Contains(data(i)) Then
                    Return False
                End If

                i += 1
            End While

            Return True
        End Function

        ''' <summary>
        ''' 16进制字符串转换为原始字符串。默认把<paramref name="sourceString"/>当做全部由单字节字符组成的字符串来处理
        ''' </summary>
        ''' <param name="sourceString"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function FromHexString(ByVal sourceString As String) As String
            Return FromHexString（sourceString, False）
        End Function

        ''' <summary>
        ''' Deflate算法压缩
        ''' </summary>
        ''' <param name="buffer"></param>
        ''' <returns></returns>
        <Extension()>
        Public Async Function CompressAsync(buffer As Byte()) As Task(Of Byte())
            If buffer Is Nothing OrElse buffer.Length = 0 Then Return buffer

            Dim compressBytes() As Byte
            Dim dms As New MemoryStream(buffer)
            Dim cms As New MemoryStream()

            Try
                Using cs As New DeflateStream(cms, CompressionMode.Compress, False)
                    Await dms.CopyToAsync(cs, 8192)
                End Using

                compressBytes = cms.ToArray()
            Catch
                Throw
            Finally
                If cms IsNot Nothing Then
                    cms.Dispose()
                End If

                If dms IsNot Nothing Then
                    dms.Dispose()
                End If
            End Try

            Return If(compressBytes, buffer)
        End Function

        ''' <summary>
        ''' Deflate算法解压缩
        ''' </summary>
        ''' <param name="buffer"></param>
        ''' <returns></returns>
        <Extension()>
        Public Async Function DeCompressAsync(buffer As Byte()) As Task(Of Byte())
            If buffer Is Nothing OrElse buffer.Length = 0 Then Return buffer

            If buffer.Length < 2 Then
                Return buffer
            End If

            Dim deCompressBytes() As Byte
            Dim cs As New MemoryStream(buffer)
            Dim dms As New MemoryStream()

            Try
                ' Deflate 算法压缩之后的数据，第一个字节是 78h（120b），第二个字节是 DAh(218b)
                If buffer(0) = 120 AndAlso buffer(1) = 218 Then
                    ' 先读取前两个deflate压缩算法标识字节，然后才能用deflateStream解压
                    ' 这个行为与 zlib库、sharpZiplib库等不同（这些库都是直接传入解压）
                    cs.ReadByte()
                    cs.ReadByte()
                End If

                Using ds As New DeflateStream(cs, CompressionMode.Decompress, False)
                    Await ds.CopyToAsync(dms, 8192)
                End Using

                deCompressBytes = dms.ToArray()
            Catch
                Throw
            Finally
                If dms IsNot Nothing Then
                    dms.Dispose()
                End If

                If cs IsNot Nothing Then
                    cs.Dispose()
                End If
            End Try

            Return If(deCompressBytes, buffer)
        End Function

        ''' <summary>
        ''' Deflate算法压缩
        ''' </summary>
        ''' <param name="buffer"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function Compress(buffer As Byte()) As Byte()
            If buffer Is Nothing OrElse buffer.Length = 0 Then Return buffer

            Dim compressBytes() As Byte
            Dim dms As New MemoryStream(buffer)
            Dim cms As New MemoryStream()

            Try
                Using cs As New DeflateStream(cms, CompressionMode.Compress, False)
                    dms.CopyTo(cs, 8192)
                End Using

                compressBytes = cms.ToArray()
            Catch
                Throw
            Finally
                If cms IsNot Nothing Then
                    cms.Dispose()
                End If

                If dms IsNot Nothing Then
                    dms.Dispose()
                End If
            End Try

            Return If(compressBytes, buffer)
        End Function

        ''' <summary>
        ''' Deflate算法解压缩
        ''' </summary>
        ''' <param name="buffer"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function DeCompress(buffer As Byte()) As Byte()
            If buffer Is Nothing OrElse buffer.Length = 0 Then Return buffer

            If buffer.Length < 2 Then
                Return buffer
            End If

            Dim cs As New MemoryStream(buffer)
            Dim dms As New MemoryStream()
            Dim deCompressBytes() As Byte

            Try

                ' Deflate 算法压缩之后的数据，第一个字节是 78h（120b），第二个字节是 DAh(218b)
                If buffer(0) = 120 AndAlso buffer(1) = 218 Then
                    ' 先读取前两个deflate压缩算法标识字节，然后才能用deflateStream解压
                    ' 这个行为与 zlib库、sharpZiplib库等不同（这些库都是直接传入解压）
                    cs.ReadByte()
                    cs.ReadByte()
                End If


                Using ds As New DeflateStream(cs, CompressionMode.Decompress, False)
                    ds.CopyTo(dms, 8192)
                End Using

                deCompressBytes = dms.ToArray()
            Catch
                Throw
            Finally
                If dms IsNot Nothing Then
                    dms.Dispose()
                End If

                If cs IsNot Nothing Then
                    cs.Dispose()
                End If
            End Try

            Return If(deCompressBytes, buffer)
        End Function

        ''' <summary>
        ''' Byte类型数组转换为16进制字符串
        ''' </summary>
        ''' <param name="sourceByte"></param>
        ''' <param name="upperLowerCase">生成大写或者小写形式的16进制字符串</param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToHexString(ByVal sourceByte() As Byte, ByVal upperLowerCase As UpperLowerCase) As String
            Dim i As Integer
            Dim sb = If(sourceByte.Length > 180,
                StringBuilderCache.AcquireSuper(sourceByte.Length * 2),
                StringBuilderCache.Acquire(sourceByte.Length * 2))
            Dim fotmat As String = "x2"
            If upperLowerCase = UpperLowerCase.Upper Then
                fotmat = "X2"
            End If

            While i < sourceByte.Length
                sb.Append(sourceByte(i).ToString(fotmat))
                i += 1
            End While

            Return If(sourceByte.Length > 180,
                StringBuilderCache.GetStringAndReleaseBuilderSuper(sb),
                StringBuilderCache.GetStringAndReleaseBuilder(sb))
        End Function

        ''' <summary>
        ''' Byte类型数组转换为16进制字符串小写形式
        ''' </summary>
        ''' <param name="sourceByte"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToHexString(ByVal sourceByte() As Byte) As String
            Return ToHexString(sourceByte, UpperLowerCase.Lower)
        End Function

        ''' <summary>
        ''' 通过反射得到单个枚举项的描述
        ''' 如果获取多个枚举项的描述，需要用 <see cref="GetDescriptions([Enum])"/>，否则无法正确获取
        ''' </summary>
        ''' <param name="enumItem"></param>
        ''' <returns></returns>       
        <Extension()>
        Public Function GetDescription(ByVal enumItem As System.Enum) As String
            Dim funcRst = String.Empty

            Dim enumItemType = enumItem.GetType()
            Dim enumItemName = enumItem.ToString()
            Dim fieldinfo = enumItemType.GetField(enumItemName)
            ' 如果enumItemName是多态枚举值或者是未找到符合指定要求的字段的对象，将返回nothing
            If fieldinfo Is Nothing Then
                Return funcRst
            End If

            Dim obj = fieldinfo.GetCustomAttributes(GetType(System.ComponentModel.DescriptionAttribute), False)

            If obj.Length = 0 Then
                Return funcRst
            Else
                Dim descriptionAttribute = DirectCast(obj.First, System.ComponentModel.DescriptionAttribute)
                If descriptionAttribute Is Nothing Then
                    Return funcRst
                End If

                funcRst = descriptionAttribute.Description
            End If

            Return funcRst
        End Function

        ''' <summary>
        ''' 利用反射获取多个枚举项的描述
        ''' 如果获取单个枚举项的描述，用Single版本会更高效
        ''' 如果枚举类没有按照建议的格式来定义，则返回空值
        ''' 建议None定义为0,与All同时被定义到枚举项中
        ''' </summary>
        ''' <param name="enumItems">如果为Nothing，则返回空值</param>
        ''' <returns>多个返回值之间用 ", " 分开</returns>
        <Extension()>
        Public Function GetDescriptions(ByVal enumItems As System.Enum) As String
            ' 如果枚举类没有按照建议的格式来定义，则返回空值
            ' 建议None定义为0,与All同时被定义到枚举项中
            If Not System.Enum.IsDefined(enumItems.GetType, "None") OrElse
                Not System.Enum.IsDefined(enumItems.GetType, "All") Then
                Return String.Empty
            End If

            ' 快速判断 None跟All
            If "None".Equals(enumItems.ToString, StringComparison.OrdinalIgnoreCase) OrElse
                "All".Equals(enumItems.ToString, StringComparison.OrdinalIgnoreCase) Then
                Return enumItems.GetDescription
            End If

            ' 利用反射获取枚举项的描述
            ' 先把enumItems中包含的所有枚举项提取出来到数组中
            ' 如果有更高效的集合类，可以替换数组，因为用数组要不断ReDim Preserve
            Dim enumValues = System.Enum.GetValues(enumItems.GetType())
            Dim tempEnumItems(enumValues.Length - 1) As System.Enum
            Dim itemCount = 0

            For i = 0 To enumValues.Length - 1
                Dim tempEnum = TryCast(enumValues(i), System.Enum)
                If tempEnum Is Nothing Then Return String.Empty
                If enumItems.HasFlag(tempEnum) Then
                    tempEnumItems(itemCount) = tempEnum
                    itemCount += 1
                End If
            Next

            ' 然后判断数组的长度，如果大于1，那就是enumItems中不包含枚举值None项和All项（通常None定义为0,一般会与All同时被定义到枚举项中）
            Dim sb = StringBuilderCache.Acquire(100)
            Dim splitString = ", "
            If itemCount > 0 Then
                For i = 0 To itemCount - 1
                    If tempEnumItems(i).Equals(System.Enum.Parse(enumItems.GetType, "None", True)) OrElse
                        tempEnumItems(i).Equals(System.Enum.Parse(enumItems.GetType, "All", True)) Then
                        Continue For
                    End If

                    ' 过滤空 Desc
                    Dim desc = tempEnumItems(i).GetDescription()
                    If desc.Length = 0 Then Continue For

                    sb.Append(String.Concat(desc, splitString))
                Next
            Else
                Dim desc = tempEnumItems(0).GetDescription()
                ' 过滤空 Desc
                If desc.Length > 0 Then
                    sb.Append(String.Concat(desc, splitString))
                End If
            End If

            ' 去掉最后的", " 从后面往前面判断
            If sb.Length > 0 AndAlso
                sb.Chars(sb.Length - 2).Equals(splitString.Chars(0)) AndAlso
                sb.Chars(sb.Length - 1).Equals(splitString.Chars(1)) Then
                sb.Remove(sb.Length - splitString.Length, splitString.Length)
            End If

            Return StringBuilderCache.GetStringAndReleaseBuilder(sb)
        End Function

        ''' <summary>
        ''' 科学计数法数字转化为一般的数字
        ''' </summary>
        ''' <param name="source"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToDefaultDecimal(ByVal source As Object) As Decimal
            Return CDec(CDbl(source))
        End Function

        ''' <summary>
        ''' 科学计数法数字转化为一般的数字
        ''' </summary>
        ''' <param name="source"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToDefaultDecimal(ByVal source As Integer) As Decimal
            Return CDec(CDbl(source))
        End Function

        ''' <summary>
        ''' 科学计数法数字转化为一般的数字
        ''' </summary>
        ''' <param name="source"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToDefaultDecimal(ByVal source As Long) As Decimal
            Return CDec(CDbl(source))
        End Function

        ''' <summary>
        ''' 科学计数法数字转化为一般的数字
        ''' </summary>
        ''' <param name="source"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToDefaultDecimal(ByVal source As String) As Decimal
            Return CDec(CDbl(source))
        End Function


        '''' <summary>
        '''' 成功返回相应的32位（4 个字节）整数
        '''' 失败抛出异常
        '''' 如果 <paramref name="value"/> 包含小数，采用“四舍六入五成双”法，其目的是弥补在将许多这样的数字相加时可能会累积的偏量。例如，将 0.5 舍入为 0，并同时将 1.5 和 2.5 舍入为 2。
        '''' </summary>
        '''' <param name="value">数值的字符串形式</param>
        '''' <returns></returns>
        '<Extension()>
        'Public Function ToIntegerOfCulture(Of T)(ByVal value As String) As T
        '    ' 为空或者为nothging直接抛出异常
        '    If String.IsNullOrEmpty(value) Then
        '        Throw New ArgumentNullException(String.Format(My.Resources.NullReference， NameOf(value)))
        '    End If

        '    ' 如果不包含小数点“.”，而且字符串长度大于integer最大值的长度(包含负数)，抛出异常
        '    ' 此处只能从长度简单判断，如果不符合上面的条件，有可能会抛出异常
        '    If value.IndexOf(".") = -1 AndAlso value.Length > CStr(T.MaxValue).Length + 1 Then
        '        Throw New ArgumentNullException(String.Format(My.Resources.NullReference， NameOf(value)))
        '    End If

        '    Return CT(value)
        'End Function

        ''' <summary>
        ''' 成功返回相应的32位（4 个字节）整数
        ''' 失败抛出异常
        ''' 如果 <paramref name="value"/> 包含小数，采用“四舍六入五成双”法，其目的是弥补在将许多这样的数字相加时可能会累积的偏量。例如，将 0.5 舍入为 0，并同时将 1.5 和 2.5 舍入为 2。
        ''' </summary>
        ''' <param name="value">数值的字符串形式</param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToIntegerOfCulture(ByVal value As String) As Integer
            ' 为空或者为nothging直接抛出异常
            If String.IsNullOrEmpty(value) Then
                Throw New ArgumentNullException(String.Format(My.Resources.NullReference， NameOf(value)))
            End If

            ' 如果不包含小数点“.”，而且字符串长度大于integer最大值的长度(包含负数)，抛出异常
            ' 此处只能从长度简单判断，如果不符合上面的条件，有可能会抛出异常
            If value.IndexOf(".") = -1 AndAlso value.Length > CStr(Integer.MaxValue).Length + 1 Then
                Throw New ArgumentNullException(String.Format(My.Resources.NullReference， NameOf(value)))
            End If

            Return CInt(value)
        End Function

#Disable Warning BC40027 ' 函数的返回类型不符合 CLS
        ''' <summary>
        ''' 成功返回相应的32位（4 个字节）整数
        ''' 失败抛出异常
        ''' 如果 <paramref name="value"/> 包含小数，采用“四舍六入五成双”法，其目的是弥补在将许多这样的数字相加时可能会累积的偏量。例如，将 0.5 舍入为 0，并同时将 1.5 和 2.5 舍入为 2。
        ''' </summary>
        ''' <param name="value">数值的字符串形式</param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToUIntegerOfCulture(ByVal value As String) As UInteger
#Enable Warning BC40027 ' 函数的返回类型不符合 CLS
            ' 为空或者为nothging直接抛出异常
            If String.IsNullOrEmpty(value) Then
                Throw New ArgumentNullException(String.Format(My.Resources.NullReference， NameOf(value)))
            End If

            ' 如果不包含小数点“.”，而且字符串长度大于integer最大值的长度(包含负数)，抛出异常
            ' 此处只能从长度简单判断，如果不符合上面的条件，有可能会抛出异常
            If value.IndexOf(".") = -1 AndAlso value.Length > CStr(Integer.MaxValue).Length + 1 Then
                Throw New ArgumentNullException(String.Format(My.Resources.NullReference， NameOf(value)))
            End If

            Return CUInt(value)
        End Function

        ''' <summary>
        ''' 成功返回相应的64位（8 个字节）整数
        ''' 失败抛出异常
        ''' 如果 <paramref name="value"/> 包含小数，采用“四舍六入五成双”法，其目的是弥补在将许多这样的数字相加时可能会累积的偏量。例如，将 0.5 舍入为 0，并同时将 1.5 和 2.5 舍入为 2。
        ''' </summary>
        ''' <param name="value">数值的字符串形式</param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToLongOfCulture(ByVal value As String) As Long
            ' 为空或者为nothging直接抛出异常
            If String.IsNullOrEmpty(value) Then
                Throw New ArgumentNullException(String.Format(My.Resources.NullReference， NameOf(value)))
            End If

            ' 如果不包含小数点“.”，而且字符串长度大于integer最大值的长度(包含负数)，抛出异常
            ' 此处只能从长度简单判断，如果不符合上面的条件，有可能会抛出异常
            If value.IndexOf(".") = -1 AndAlso value.Length > CStr(Long.MaxValue).Length + 1 Then
                Throw New ArgumentNullException(String.Format(My.Resources.NullReference， NameOf(value)))
            End If

            Return CLng(value)
        End Function

#Disable Warning BC40027 ' 函数的返回类型不符合 CLS
        ''' <summary>
        ''' 成功返回相应的64位（8 个字节）整数
        ''' 失败抛出异常
        ''' 如果 <paramref name="value"/> 包含小数，采用“四舍六入五成双”法，其目的是弥补在将许多这样的数字相加时可能会累积的偏量。例如，将 0.5 舍入为 0，并同时将 1.5 和 2.5 舍入为 2。
        ''' </summary>
        ''' <param name="value">数值的字符串形式</param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToULongOfCulture(ByVal value As String) As ULong
#Enable Warning BC40027 ' 函数的返回类型不符合 CLS
            ' 为空或者为nothging直接抛出异常
            If String.IsNullOrEmpty(value) Then
                Throw New ArgumentNullException(String.Format(My.Resources.NullReference， NameOf(value)))
            End If

            ' 如果不包含小数点“.”，而且字符串长度大于integer最大值的长度(包含负数)，抛出异常
            ' 此处只能从长度简单判断，如果不符合上面的条件，有可能会抛出异常
            If value.IndexOf(".") = -1 AndAlso value.Length > CStr(ULong.MaxValue).Length + 1 Then
                Throw New ArgumentNullException(String.Format(My.Resources.NullReference， NameOf(value)))
            End If

            Return CULng(value)
        End Function

        ''' <summary>
        ''' 成功返回相应带符号的 IEEE 32 位（4 个字节）双精度浮点数
        ''' 失败返回 -1.0 或者抛出异常
        ''' </summary>
        ''' <param name="value">数值的字符串形式</param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToSingleOfCulture(ByVal value As String) As Single
            ' 为空或者为nothging直接抛出异常
            If String.IsNullOrEmpty(value) Then
                Throw New ArgumentNullException(String.Format(My.Resources.NullReference， NameOf(value)))
            End If

            ' 如果不包含小数点“.”，而且字符串长度大于integer最大值的长度，抛出异常.0
            ' 此处只能从长度简单判断，如果不符合上面的条件，有可能会抛出异常
            If value.IndexOf(".") = -1 AndAlso value.Length > CStr(Single.MaxValue).Length Then
                Throw New ArgumentNullException(String.Format(My.Resources.NullReference， NameOf(value)))
            End If

            Return CSng(value)
        End Function

        ''' <summary>
        ''' 成功返回相应带符号的 IEEE 64 位（8 个字节）双精度浮点数
        ''' 失败返回 -1.0 或者抛出异常
        ''' </summary>
        ''' <param name="value">数值的字符串形式</param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToDoubleOfCulture(ByVal value As String) As Double
            ' 为空或者为nothging直接抛出异常
            If String.IsNullOrEmpty(value) Then
                Throw New ArgumentNullException(String.Format(My.Resources.NullReference， NameOf(value)))
            End If

            ' 如果不包含小数点“.”，而且字符串长度大于integer最大值的长度，抛出异常.0
            ' 此处只能从长度简单判断，如果不符合上面的条件，有可能会抛出异常
            If value.IndexOf(".") = -1 AndAlso value.Length > CStr(Double.MaxValue).Length Then
                Throw New ArgumentNullException(String.Format(My.Resources.NullReference， NameOf(value)))
            End If

            Return CDbl(value)
        End Function

        '''' <summary>
        '''' 以符合当前区域设置来转换
        '''' 成功返回相应字符串表现形式
        '''' 失败返回 String.Empty 常量 或者抛出异常
        '''' </summary>
        '''' <param name="value">Integer 类型的整数</param>
        '''' <returns></returns>
        '<Extension()>
        'Public Function ToStringOfCulture(Of T)(ByVal value As T) As String
        '    Return CStr(value)
        'End Function

        ''' <summary>
        ''' 以符合当前区域设置来转换
        ''' 成功返回相应字符串表现形式
        ''' 失败返回 String.Empty 常量 或者抛出异常
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToStringOfCulture(ByVal value As Short) As String
            Return CStr(value)
        End Function

#Disable Warning BC40028 ' 参数的类型不符合 CLS
        ''' <summary>
        ''' 以符合当前区域设置来转换
        ''' 成功返回相应字符串表现形式
        ''' 失败返回 String.Empty 常量 或者抛出异常
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToStringOfCulture(ByVal value As UShort) As String
#Enable Warning BC40028 ' 参数的类型不符合 CLS
            Return CStr(value)
        End Function

        ''' <summary>
        ''' 以符合当前区域设置来转换
        ''' 成功返回相应字符串表现形式
        ''' 失败返回 String.Empty 常量 或者抛出异常
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToStringOfCulture(ByVal value As Integer) As String
            Return CStr(value)
        End Function

#Disable Warning BC40028 ' 参数的类型不符合 CLS
        ''' <summary>
        ''' 以符合当前区域设置来转换
        ''' 成功返回相应字符串表现形式
        ''' 失败返回 String.Empty 常量 或者抛出异常
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToStringOfCulture(ByVal value As UInteger) As String
#Enable Warning BC40028 ' 参数的类型不符合 CLS
            Return CStr(value)
        End Function

        ''' <summary>
        ''' 以符合当前区域设置来转换
        ''' 成功返回相应字符串表现形式
        ''' 失败返回 String.Empty 常量 或者抛出异常
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToStringOfCulture(ByVal value As Long) As String
            Return CStr(value)
        End Function

#Disable Warning BC40028 ' 参数的类型不符合 CLS
        ''' <summary>
        ''' 以符合当前区域设置来转换
        ''' 成功返回相应字符串表现形式
        ''' 失败返回 String.Empty 常量 或者抛出异常
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToStringOfCulture(ByVal value As ULong) As String
#Enable Warning BC40028 ' 参数的类型不符合 CLS
            Return CStr(value)
        End Function

        ''' <summary>
        ''' 以符合当前区域设置来转换
        ''' 成功返回相应字符串表现形式
        ''' 失败返回 String.Empty 常量 或者抛出异常
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToStringOfCulture(ByVal value As Single) As String
            Return CStr(value)
        End Function

        ''' <summary>
        ''' 以符合当前区域设置来转换
        ''' 成功返回相应字符串表现形式
        ''' 失败返回 String.Empty 常量 或者抛出异常
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToStringOfCulture(ByVal value As Double) As String
            Return CStr(value)
        End Function

        ''' <summary>
        ''' 以符合当前区域设置来转换
        ''' 成功返回相应字符串表现形式
        ''' 失败返回 String.Empty 常量 或者抛出异常
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToStringOfCulture(ByVal value As Date) As String
            Return CStr(value)
        End Function

        ''' <summary>
        ''' 以符合当前区域设置来转换
        ''' 成功返回相应字符串表现形式
        ''' 失败返回 String.Empty 常量 或者抛出异常
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToStringOfCulture(ByVal value As Boolean) As String
            Return CStr(value)
        End Function

        ''' <summary>
        ''' 以符合当前区域设置来转换
        ''' 成功返回相应字符串表现形式
        ''' 失败返回 String.Empty 常量 或者抛出异常
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToStringOfCulture(ByVal value As Object) As String
            Return CStr(value)
        End Function

        ''' <summary>
        ''' 将字符串解析为 Integer 类型数值
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ParseToInteger(value As String) As (Success As Boolean, Result As Integer)
            Dim number As Integer
            Return (Integer.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, number), number)
        End Function

#Disable Warning BC40041 ' 类型不符合 CLS
        ''' <summary>
        ''' 将字符串解析为 UInteger 类型数值
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ParseToUInteger(value As String) As (Success As Boolean, Result As UInteger)
#Enable Warning BC40041 ' 类型不符合 CLS
            Dim number As UInteger
            Return (UInteger.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, number), number)
        End Function

        ''' <summary>
        ''' 将字符串解析为 Long 类型数值
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ParseToLong(value As String) As (Success As Boolean, Result As Long)
            Dim number As Long
            Return (Long.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, number), number)
        End Function


#Disable Warning BC40041 ' 类型不符合 CLS
        ''' <summary>
        ''' 将字符串解析为 ULong 类型数值
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ParseToULong(value As String) As (Success As Boolean, Result As ULong)
#Enable Warning BC40041 ' 类型不符合 CLS
            Dim number As ULong
            Return (ULong.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, number), number)
        End Function

        ''' <summary>
        ''' 将字符串解析为 Decimal 类型数值
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ParseToDecimal(value As String) As (Success As Boolean, Result As Decimal)
            Dim number As Decimal
            Return (Decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, number), number)
        End Function

        ''' <summary>
        ''' 将字符串解析为 Double 类型数值
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ParseToDouble(value As String) As (Success As Boolean, Result As Double)
            Dim number As Double
            Return (Double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, number), number)
        End Function

        ''' <summary>
        ''' 将字符串解析为 Double 类型数值
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ParseToBoolean(value As String) As (Success As Boolean, Result As Boolean)
            Dim bool As Boolean
            Return (Boolean.TryParse(value, bool), bool)
        End Function

        ''' <summary>
        ''' 使用当前区域设置进行转换
        ''' </summary>
        ''' <param name="source"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToIntegerOfCulture(ByVal source As Decimal) As Integer
            Return CInt(source)
        End Function

        ''' <summary>
        ''' 反序列化
        ''' </summary>
        ''' <param name="input"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function Deserialize(Of T As Class)(ByVal input As String) As T
            Return MSJsSerializer.Deserialize(Of T)(input)
        End Function

        ''' <summary>
        ''' 序列化
        ''' </summary>
        ''' <param name="source"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function Serialize(Of T As Class)(ByVal source As T) As String
            Return MSJsSerializer.Serialize(source)
        End Function

        ''' <summary>
        ''' 字符串形式的键值对转换为集合。此函数与类库自带的函数 <see cref="Web.HttpUtility.ParseQueryString"/> 不同，不需要提前编码；如果值域包含特殊字符如“=、+”等，类库自带的函数 <see cref="Web.HttpUtility.ParseQueryString"/> 无法正常解析
        ''' </summary>
        ''' <param name="kvString"></param>
        ''' <param name="urlEncoded"><paramref name="kvString"/> 是否已经编码</param>
        ''' <param name="encoding">指示 <paramref name="kvString"/> 采用的是何种编码，如果 <paramref name="kvString"/> 未编码，可传入 Nothing</param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToKeyValuePairs(kvString As String, urlEncoded As Boolean, encoding As Encoding) As IEnumerable(Of KeyValuePair(Of String, String))
            ' 如果已经编码过，那就先解码
            If urlEncoded Then kvString = HttpUtility.UrlDecode(kvString, encoding)

            Dim list As New List(Of KeyValuePair(Of String, String))(8)
            Dim l = If(kvString <> Nothing, kvString.Length, 0)
            Dim i = 0

            While i < l
                ' find next & while noting first = on the way (and if there are more)
                Dim si As Integer = i
                Dim ti As Integer = -1

                While i < l
                    Dim ch As Char = kvString(i)

                    If ch = "="c Then
                        If ti < 0 Then
                            ti = i
                        End If
                    ElseIf ch = "&"c AndAlso kvString.IsLetterNext(i) Then
                        Exit While
                    End If

                    i += 1
                End While

                ' extract the name / value pair
                Dim name As String = Nothing
                Dim value As String

                If ti >= 0 Then
                    name = kvString.Substring(si, ti - si)
                    value = kvString.Substring(ti + 1, i - ti - 1)
                Else
                    value = kvString.Substring(si, i - si)
                End If

                ' add name / value pair to the collection
                list.Add(New KeyValuePair(Of String, String)(name, value))

                ' trailing '&'
                If i = l - 1 AndAlso kvString(i) = "&"c Then
                    list.Add(New KeyValuePair(Of String, String)(Nothing, String.Empty))
                End If

                i += 1
            End While

            Return list
        End Function

        ''' <summary>
        ''' 字符串形式的键值对转换为集合。此函数与类库自带的函数 <see cref="Web.HttpUtility.ParseQueryString"/> 不同，不需要提前编码；如果值域包含特殊字符如“=、+”等，类库自带的函数 <see cref="Web.HttpUtility.ParseQueryString"/> 无法正常解析
        ''' </summary>
        ''' <param name="kvString"></param>
        ''' <param name="urlEncoded"><paramref name="kvString"/> 是否已经编码</param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToKeyValuePairs(kvString As String, urlEncoded As Boolean) As IEnumerable(Of KeyValuePair(Of String, String))
            Return ToKeyValuePairs(kvString, urlEncoded, Nothing)
        End Function

        ''' <summary>
        ''' 字符串形式的键值对转换为集合。此函数与类库自带的函数 <see cref="Web.HttpUtility.ParseQueryString"/> 不同，不需要提前编码；如果值域包含特殊字符如“=、+”等，类库自带的函数 <see cref="Web.HttpUtility.ParseQueryString"/> 无法正常解析。
        ''' <para>注：<paramref name="kvString"/> 必须是未执行过 <see cref="HttpUtility.UrlEncode(String, Encoding)"/> 的字符串</para>
        ''' </summary>
        ''' <param name="kvString"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToKeyValuePairs(kvString As String) As IEnumerable(Of KeyValuePair(Of String, String))
            Return ToKeyValuePairs(kvString, False, Nothing)
        End Function

        ''' <summary>
        ''' 判断下一个字符是否为字母（大写或者是小写英文字母）
        ''' </summary>
        ''' <param name="sourceString">源字符串</param>
        ''' <param name="offset">当前索引</param>
        ''' <returns></returns>
        <Extension()>
        Public Function IsLetterNext(ByVal sourceString As String, ByVal offset As Integer) As Boolean
            Dim nextOffset = offset + 1
            Return sourceString.Length > nextOffset AndAlso (Char.IsLower(sourceString.Chars(nextOffset)) OrElse
            Char.IsUpper(sourceString.Chars(nextOffset)) OrElse
            "_"c = sourceString.Chars(nextOffset))
        End Function

        Private Sub GetKvp(ByRef isValue As Boolean, ByRef KeyValuePairs As Dictionary(Of String, String), ByRef sbKey As StringBuilder, ByRef sbValue As StringBuilder)
            ' 如果只有 ‘&’ 而找不到配对的 ‘=’ 说明 ‘&’是属于值的一部分
            ' 需要把下一个键值对之前的内容追加到上一个键值对的值中
            If Not isValue Then
                Dim lastItem = KeyValuePairs.Last
                sbKey.Insert(0, "&"c).Insert(0, lastItem.Value)
                Dim value = StringBuilderCache.GetStringAndReleaseBuilder(sbKey)
                KeyValuePairs(lastItem.Key) = value
            Else
                Dim key = StringBuilderCache.GetStringAndReleaseBuilder(sbKey)
                Dim value = StringBuilderCache.GetStringAndReleaseBuilder(sbValue)
                KeyValuePairs.Add(key, value)
            End If
            sbKey.Length = 0
            sbValue.Length = 0
            isValue = False
        End Sub

        ''' <summary>
        ''' 加密Cookies
        ''' </summary>
        ''' <param name="cookies"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function EncryptCookies(ByVal cookies As Net.CookieContainer) As String
            Dim cookiesKvp = cookies.ToKeyValuePairs
            Dim hexCookiesKvp = cookiesKvp.ToHexString(UpperLowerCase.Lower)
            Dim bytes = hexCookiesKvp.HexStringToBytes
            bytes = bytes.Compress
            hexCookiesKvp = bytes.ToHexString(UpperLowerCase.Lower)

            Return hexCookiesKvp
        End Function

        ''' <summary>
        ''' 解密Cookies字符串，一般与 <see cref="EncryptCookies(Net.CookieContainer)"/> 配套使用
        ''' </summary>
        ''' <param name="hexCookiesKvp"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function DecryptCookies(ByVal hexCookiesKvp As String) As String
            Dim bytes = hexCookiesKvp.HexStringToBytes
            bytes = bytes.DeCompress
            hexCookiesKvp = bytes.ToHexString(UpperLowerCase.Lower)
            Dim sourceCookiesKvp = hexCookiesKvp.FromHexString

            Return sourceCookiesKvp
        End Function

        ''' <summary>
        ''' 将对象序列化为二进制
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="instance"></param>
        ''' <param name="nameOfInstance">对象名，可用 NamoOf 获取</param>
        ''' <returns>成功返回二进制文件存储路径,失败返回 <see cref="String.Empty"/></returns>
        <Extension()>
        Public Function BinarySerialize(Of T As Class)(ByVal instance As T, ByVal nameOfInstance As String) As String
            Try
                If instance Is Nothing Then Return String.Empty

                Dim objType = instance.GetType
                Dim filePath As String = $"{nameOfInstance}_{objType.Name}.bs"
                Using fs As New FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)
                    Dim bf As New BinaryFormatter()
                    bf.Serialize(fs, instance)
                End Using

                Return filePath
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try

            Return String.Empty
        End Function

        ''' <summary>
        ''' 将二进制反序列化为对象
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="outInstance">接收反序列化结果的对象，可传入未初始化的对象，但不能传入 Nothing</param>
        ''' <param name="storePath">二进制文件的储存路径</param>
        <Extension()>
        Public Sub BinaryDeserialize(Of T As Class)(ByRef outInstance As T, ByVal storePath As String)
            Try
                Using fs As New FileStream(storePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                    Dim bf As New BinaryFormatter()
                    Dim obj = bf.Deserialize(fs)
                    If obj Is Nothing Then Return

                    outInstance = DirectCast(obj, T)
                End Using
            Catch ex As Exception
                '
            End Try
        End Sub
    End Module
End Namespace
