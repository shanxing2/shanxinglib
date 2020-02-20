Imports System.ComponentModel

Namespace ShanXingTech.IO2

    ''' <summary>
    ''' 从Txt到Excel导出时，列行为以及列数据类型的控制
    ''' </summary>
    Public Enum ExcelColumnDataType
        ''' <summary>
        ''' 自动转换为合适的数据类型（根据前8行数据。默认的数据列类型）
        ''' </summary>
        <Description("常规")>
        XlGeneralFormat = 1

        ''' <summary>
        ''' 自动转换为文本
        ''' </summary>
        <Description("文本")>
        XlTextFormat
        ''' <summary>
        ''' '月-日-年'日期格式
        ''' </summary>
        <Description("""月-日-年""日期格式")>
        XLMDYFormat
        ''' <summary>
        ''' '日-月-年'日期格式
        ''' </summary>
        <Description("""日-月-年""日期格式")>
        XLDMYFormat
        ''' <summary>
        ''' '年-月-日'日期格式
        ''' </summary>
        <Description("""年-月-日""日期格式")>
        XLYMDFormat
        ''' <summary>
        ''' '月-年-日'日期格式
        ''' </summary>
        <Description("""月-年-日""日期格式")>
        XLMYDFormat
        ''' <summary>
        ''' '日-年-月'日期格式
        ''' </summary>
        <Description("""日-年-月""日期格式")>
        XLDYMFormat
        ''' <summary>
        ''' '年-日-月'日期格式
        ''' </summary>
        <Description("""年-日-月""日期格式")>
        XLYDMFormat

        ''' <summary>
        ''' 跳过列，不导出
        ''' </summary>
        <Description("跳过列")>
        XlSkipColumn

        ''' <summary>
        ''' 只有在安装并选定了台湾地区语言时才可以使用 XlEMDFormat。XlEMDFormat 常量指定使用台湾地区纪元日期。
        ''' </summary>
        <Description("EMD 日期")>
        XLEMDFormat
    End Enum
End Namespace
