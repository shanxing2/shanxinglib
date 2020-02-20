Imports System.ComponentModel

Namespace ShanXingTech
    ''' <summary>
    ''' 浏览器仿真版本
    ''' 注：
    ''' <para>1.标准模式是指浏览器的默认模式，而 !DOCTYPE 控制的是文档模式</para> 
    ''' <para>2.所有应用程序中使用的WebBrowser控件默认的模式是IE7标准模式</para>
    ''' </summary>
    Public Enum BrowserEmulationMode
        ''' <summary>
        ''' 以IE7的标准模式按照 !DOCTYPE 指令来展现网页
        ''' </summary>
        IE7 = 7000
        ''' <summary>
        ''' 以IE8的标准模式按照 !DOCTYPE 指令来展现网页
        ''' </summary>
        IE8 = 8000
        ''' <summary>
        ''' 强制以IE8的标准模式展现，忽略 !DOCTYPE 指令
        ''' </summary>
        IE8Standard = 8888
        ''' <summary>
        ''' 以IE9的标准模式按照 !DOCTYPE 指令来展现网页
        ''' </summary>
        IE9 = 9000
        ''' <summary>
        ''' 强制以IE9的标准模式展现，忽略 !DOCTYPE 指令
        ''' </summary>
        IE9Standard = 9999
        ''' <summary>
        ''' 以IE10的标准模式按照 !DOCTYPE 指令来展现网页
        ''' </summary>
        IE10 = 10000
        ''' <summary>
        ''' 强制以IE10的标准模式展现，忽略 !DOCTYPE 指令
        ''' </summary>
        IE10Standard = 10001
        ''' <summary>
        ''' 以IE11的标准模式按照 !DOCTYPE 指令来展现网页
        ''' </summary>
        IE11 = 11000
        ''' <summary>
        ''' 强制以IE11的标准模式展现，忽略 !DOCTYPE 指令
        ''' </summary>
        IE11Standard = 11001
    End Enum

    ''' <summary>
    ''' 鼠标移出窗体后执行的动作
    ''' </summary>
    Public Enum MouseLeaveAction As Integer
        ''' <summary>
        ''' 什么都不做
        ''' </summary>
        None
        ''' <summary>
        ''' 关闭窗体
        ''' </summary>
        Close
        ''' <summary>
        ''' 隐藏窗体
        ''' </summary>
        Hide
    End Enum

    ''' <summary>
    ''' 时间精度的
    ''' </summary>
    Public Enum TimePrecision As Integer
        ''' <summary>
        ''' 精度：秒 10个字符长度
        ''' </summary>
        <Description("秒")>
        Second
        ''' <summary>
        ''' 精度：毫秒 13个字符长度
        ''' </summary>
        <Description("毫秒")>
        Millisecond
    End Enum

    Public Enum UpperLowerCase As Integer
        Upper
        Lower
    End Enum
End Namespace
