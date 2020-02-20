Imports System.IO
Imports System.Text

Namespace ShanXingTech.IO2

    ''' <summary>
    ''' 获取文件编码类
    ''' </summary>
    Public Class EncodingHelper
        ''' <summary> 
        ''' 给定文件的路径，读取文件的二进制数据，判断文件的编码类型 
        ''' </summary> 
        ''' <param name="fileName">文件路径</param> 
        ''' <returns>文件的编码类型</returns> 
        Public Shared Function GetFileEncoding(ByVal fileName As String) As System.Text.Encoding
            Dim r As Encoding

            Using fs As New FileStream(fileName, FileMode.Open, FileAccess.Read)
                r = GetFileEncoding(fs)
            End Using

            Return r
        End Function

        ''' <summary> 
        ''' 通过给定的文件流，判断文件的编码类型 
        ''' </summary> 
        ''' <param name="fs">文件流</param> 
        ''' <returns>文件的编码类型</returns> 
        Public Shared Function GetFileEncoding(ByVal fs As FileStream) As System.Text.Encoding
            Dim unicode = Encoding.Unicode.GetPreamble
            Dim bigEndianUnicode = Encoding.BigEndianUnicode.GetPreamble
            Dim utf8 = Encoding.UTF8.GetPreamble
            Dim utf32 = Encoding.UTF32.GetPreamble

            Dim funcRst As Encoding = Encoding.Default
            '文件的字符集在Windows下有两种， 一种是ANSI， 一种Unicode。
            '对于Unicode， Windows支持了它的三种编码方式， 一种是小尾编码（unicode)， 一种是大尾编码(BigEndianUnicode)， 一种是UTF - 8编码。
            '我们可以从文件的头部来区分一个文件是属于哪种编码。当头部开始的两个字节为 FF FE时，是Unicode的小尾编码；当头部的两个字节为FE FF时，是Unicode的大尾编码；当头部两个字节为EF BB时，是Unicode的UTF-8编码；当头部三个字节为EF BB BF时，是带bom的Unicode的UTF-8编码；当它不为这些时，则是ANSI编码。

            Using r As New BinaryReader(fs, System.Text.Encoding.Default)
                Dim i As Integer = CInt(fs.Length)

                Dim ss As Byte() = r.ReadBytes(i)
                If IsUTF8Bytes(ss) OrElse (ss(0) = utf8(0) AndAlso ss(1) = utf8(1) AndAlso ss(2) = utf8(2)) Then
                    funcRst = Encoding.UTF8
                ElseIf IsUTF8Bytes(ss) OrElse (ss(0) = utf32(0) AndAlso ss(1) = utf32(1) AndAlso ss(2) = utf32(2)) AndAlso ss(3) = utf32(3) Then
                    funcRst = Encoding.UTF32
                ElseIf ss(0) = bigEndianUnicode(0) AndAlso ss(1) = bigEndianUnicode(1) Then
                    funcRst = Encoding.BigEndianUnicode
                ElseIf ss(0) = unicode(0) AndAlso ss(1) = unicode(1) Then
                    funcRst = Encoding.Unicode
                End If
            End Using

            Return funcRst
        End Function

        ''' <summary> 
        ''' 判断是否是不带 BOM 的 utf8 格式 
        ''' </summary> 
        ''' <param name="data"></param> 
        ''' <returns></returns> 
        Private Shared Function IsUTF8Bytes(data As Byte()) As Boolean
            Dim charByteCounter As Integer = 1
            ' 计算当前正分析的字符应还有的字节数 
            Dim curByte As Byte
            ' 当前分析的字节. 
            Dim i As Integer = 0
            While i < data.Length
                curByte = data(i)
                If charByteCounter = 1 Then
                    If curByte >= &H80 Then
                        ' 判断当前 
                        While (curByte And &H80) <> 0
                            curByte = curByte << 1
                            charByteCounter += 1
                        End While
                        ' 标记位首位若为非0 则至少以2个1开始 如:110XXXXX...........1111110X　 
                        If charByteCounter = 1 OrElse charByteCounter > 6 Then
                            Return False
                        End If
                    End If
                Else
                    ' 若是UTF-8 此时第一位必须为1 
                    If (curByte And &HC0) <> &H80 Then
                        Return False
                    End If
                    charByteCounter -= 1
                End If
                i += 1
            End While

            If charByteCounter > 1 Then
                Throw New Exception("非预期的文件格式")
            End If

            Return True
        End Function
    End Class
End Namespace
