Imports System.ComponentModel


Namespace ShanXingTech

    ''' <summary>
    ''' 文件过滤器
    ''' </summary>
    <Flags>
    Public Enum FileFilter
        ''' <summary>
        ''' 无，一般不使用
        ''' </summary>
        <Description("")>
        None = 0
        ''' <summary>
        ''' EXCEL 工作簿(*.xls;*.xlsx)|*.xls;*.xlsx
        ''' </summary>
        <Description("EXCEL 工作簿(*.xls;*.xlsx)|*.xls;*.xlsx")>
        EXCEL = 1

        ''' <summary>
        ''' 文本文档(*.txt)|*.txt
        ''' </summary>
        <Description("文本文档(*.txt)|*.txt")>
        TXT = 2

        ''' <summary>
        ''' CSV (逗号分隔)(*.csv)|*.CSV
        ''' </summary>
        <Description("CSV (逗号分隔)文件(*.csv)|*.csv")>
        CSV = 4
        ''' <summary>
        ''' Access 数据库(*.mdb;*.mde;*.accdb;*.accde)|*.mdb;*.mde;*.accdb;*.accde
        ''' </summary>
        <Description("Access 数据库(*.mdb;*.mde;*.accdb;*.accde)|*.mdb;*.mde;*.accdb;*.accde")>
        Access = 8
        ''' <summary>
        ''' 可执行程序(*.exe)|*.exe
        ''' </summary>
        <Description("可执行程序(*.exe)|*.exe")>
        EXE = 16
        ''' <summary>
        ''' 所有文件(*.*)|*.*
        ''' </summary>
        <Description("所有文件(*.*)|*.*")>
        All = EXCEL Or TXT Or CSV Or Access Or EXE
    End Enum

End Namespace
