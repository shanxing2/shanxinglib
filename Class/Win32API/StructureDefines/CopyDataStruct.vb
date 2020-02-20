Imports System.Runtime.InteropServices

Namespace ShanXingTech.Win32API
    ''' <summary>
    ''' 进程间通信数据载体结构
    ''' </summary>
    <StructLayout(LayoutKind.Sequential)>
    Public Structure CopyDataStruct
        ''' <summary>
        ''' 32位的自定义数据
        ''' </summary>
        Public Type As IntPtr
        ''' <summary>
        ''' lpData指针指向数据的大小（字节数）
        ''' </summary>
        Public ByteLength As Integer
        ''' <summary>
        ''' 指向数据的指针
        ''' </summary>
        Public Data As String
        'Public Data As IntPtr
    End Structure
End Namespace
