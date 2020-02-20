Imports ShanXingTech.Exception2
Imports System.Runtime.CompilerServices


Namespace ShanXingTech
    Partial Public Module ExtensionFunc
        ''' <summary>
        ''' 生成时间戳，无小数。
        ''' </summary>
        ''' <param name="sourceDate"></param>
        ''' <param name="TimePrecision">输出的时间戳精度。精度毫秒(13个字符),精度秒(10个字符)</param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToTimeStamp(ByVal sourceDate As Date, ByVal timePrecision As TimePrecision) As Long
            ' 不需要进行此步判断，因为TimePrecision参数是枚举类型，如果在使用的时候
            ' 输入了一个非枚举类型地址，编译器会自动报错的
            'If timePrecision <> timePrecision.Millisecond AndAlso timePrecision <> timePrecision.Second Then Return String.Empty

            Dim precision = 10000
            If timePrecision = TimePrecision.Second Then
                precision = 10000000
            End If
            ' New DateTime(1970, 1, 1).Ticks = 621355968000000000
            Dim funcRst = CLng((sourceDate.ToUniversalTime.Ticks - 621355968000000000) / precision)
            Return funcRst
        End Function

        ''' <summary>
        ''' 生成时间戳，无小数。
        ''' </summary>
        ''' <param name="sourceDate"></param>
        ''' <param name="TimePrecision">输出的时间戳精度。精度毫秒(13个字符),精度秒(10个字符)</param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToTimeStampString(ByVal sourceDate As Date, ByVal timePrecision As TimePrecision) As String
            ' 不需要进行此步判断，因为TimePrecision参数是枚举类型，如果在使用的时候
            ' 输入了一个非枚举类型地址，编译器会自动报错的
            'If timePrecision <> timePrecision.Millisecond AndAlso timePrecision <> timePrecision.Second Then Return String.Empty
            Dim funcRst = ToTimeStamp(sourceDate， timePrecision).ToStringOfCulture
            Return funcRst
        End Function

        ''' <summary>
        ''' 生成时间戳，无小数。默认精度毫秒(13个字符)
        ''' </summary>
        ''' <param name="sourceDate"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToTimeStampString(ByVal sourceDate As Date) As String
            Return ToTimeStampString(sourceDate, TimePrecision.Millisecond)
        End Function

        ''' <summary>
        ''' 时间戳转换成日期
        ''' </summary>
        ''' <param name="timeStampValue"></param>
        ''' <param name="TimePrecision">输出的时间戳精度。精度毫秒(13个字符),精度秒(10个字符)</param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToTimeStampTime(ByVal timeStampValue As Long, ByVal timePrecision As TimePrecision) As Date
            If timeStampValue < 0 Then Throw New TimeStampException(String.Format(My.Resources.ArgumentOutOfRange, NameOf(timeStampValue), "(0,Long.MaxValue]"))

            Dim dtStart = TimeZone.CurrentTimeZone.ToLocalTime(New DateTime(1970, 1, 1))
            Dim precision = 10000L
            If timePrecision = TimePrecision.Second Then
                precision = 10000000
            End If
            Dim ticks = timeStampValue * precision
            Dim dtTime = dtStart.AddTicks(ticks)

            Return dtTime
        End Function

        ''' <summary>
        ''' 毫秒级(13个字符,Long 类型)时间戳转换成日期。
        ''' </summary>
        ''' <param name="timeStampValue"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToTimeStampTime(ByVal timeStampValue As Long) As Date
            Return ToTimeStampTime(timeStampValue, TimePrecision.Millisecond)
        End Function

        ''' <summary>
        ''' 秒级时间戳转换成日期
        ''' </summary>
        ''' <param name="timeStampValue"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToTimeStampTime(ByVal timeStampValue As Integer) As Date
            ' 因为 Integer 针对32位处理器进行过优化，所以直接在实现一个 Integer 的版本，而不是直接调用 Long 的版本
            ' 又因为 timeStamp * 10000000 会导致运算溢出，所以还是直接调用 Long 版本吧
            ' 20180627
            Return ToTimeStampTime(timeStampValue, TimePrecision.Second)
        End Function

        ''' <summary>
        ''' 时间戳转换成日期
        ''' </summary>
        ''' <param name="timeStampValue"></param>
        ''' <param name="TimePrecision">输出的时间戳精度。精度毫秒(13个字符),精度秒(10个字符)</param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToTimeStampString(ByVal timeStampValue As Long, ByVal timePrecision As TimePrecision, ByVal format As String) As String
            Return ToTimeStampTime(timeStampValue, timePrecision).ToString(format)
        End Function

        ''' <summary>
        ''' 毫秒级时间戳转换成日期。
        ''' </summary>
        ''' <param name="timeStampValue"></param>
        ''' <returns>默认生成字符串格式为 yyyy-MM-dd HH:mm:ss </returns>
        <Extension()>
        Public Function ToTimeStampString(ByVal timeStampValue As Long) As String
            Return ToTimeStampString(timeStampValue, TimePrecision.Millisecond, "yyyy-MM-dd HH:mm:ss")
        End Function

        ''' <summary>
        ''' 秒级时间戳转换成日期
        ''' </summary>
        ''' <param name="timeStampValue"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function ToTimeStampString(ByVal timeStampValue As Integer, ByVal format As String) As String
            Return ToTimeStampTime(timeStampValue, TimePrecision.Second).ToString(format)
        End Function


        ''' <summary>
        ''' 秒级时间戳转换成日期
        ''' </summary>
        ''' <param name="timeStampValue"></param>
        ''' <returns>默认生成字符串格式为 yyyy-MM-dd HH:mm:ss </returns>
        <Extension()>
        Public Function ToTimeStampString(ByVal timeStampValue As Integer) As String
            Return ToTimeStampTime(timeStampValue, TimePrecision.Second).ToString("yyyy-MM-dd HH:mm:ss")
        End Function
    End Module
End Namespace
