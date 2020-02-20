Namespace ShanXingTech.IO2
    ''' <summary>
    ''' 用以构建TXT文件导出为EXCEL文件的辅助类
    ''' </summary>
    Public Class ExportExcelFileInfo
        Public Sub New()
            ' 设置前10列默认数据类型为自动根据数据的前8行转换
            TextFileColumnsDataType = New List(Of ExcelColumnDataType) From {
                ExcelColumnDataType.XlGeneralFormat,
                ExcelColumnDataType.XlGeneralFormat,
                ExcelColumnDataType.XlGeneralFormat,
                ExcelColumnDataType.XlGeneralFormat,
                ExcelColumnDataType.XlGeneralFormat,
                ExcelColumnDataType.XlGeneralFormat,
                ExcelColumnDataType.XlGeneralFormat,
                ExcelColumnDataType.XlGeneralFormat,
                ExcelColumnDataType.XlGeneralFormat,
                ExcelColumnDataType.XlGeneralFormat
            }

            TextFilePlatform = 65001
        End Sub

        ''' <summary>
        ''' 【必须】工作薄名
        ''' </summary>
        ''' <returns></returns>
        Public Property SheetName As String
        ''' <summary>
        ''' 【必须】txt缓存文件名
        ''' </summary>
        ''' <returns></returns>
        Public Property TxtCachePath As String
        ''' <summary>
        ''' 【必须】需要导出的Excel文件名,不需要带后缀
        ''' </summary>
        ''' <returns></returns>
        Public Property ExcelFileName As String
        ''' <summary>
        ''' 文件编码格式
        ''' 默认UTF-8编码(65001)
        ''' </summary>
        ''' <returns></returns>
        Public Property TextFilePlatform As Integer
        ''' <summary>
        ''' 【必须】自定义字段分隔符，缓存数据分裂的依据； 不能跟系统已经定义的一样（分别有：Tab、';'、' '、','）。
        ''' 将连续分隔符看作是一个分隔符
        ''' </summary>
        ''' <returns></returns>
        Public Property TextFileOtherDelimiter As String
        ''' <summary>
        ''' 字段数据类型列表
        ''' 默认设置10列，数据类型全部为 ExcelColumnDataType.XlGeneralFormat
        ''' </summary>
        ''' <returns></returns>
        Public Property TextFileColumnsDataType As List(Of ExcelColumnDataType)
    End Class
End Namespace