Imports System.Windows.Forms
Imports ShanXingTech.Win32API
Imports ShanXingTech.Win32API.UnsafeNativeMethods

Namespace ShanXingTech
    ''' <summary>
    ''' 进程间通讯相关类
    ''' </summary>
    Public NotInheritable Class Processer
        Private Shared s_Encoding As Text.Encoding
        Shared Sub New()
            s_Encoding = Text.Encoding.UTF8
        End Sub

#Region "函数区"
        '''' <summary>
        '''' 向进程发送消息
        '''' </summary>
        '''' <param name="hWnd"></param>
        '''' <param name="message"></param>
        '''' <returns></returns>
        'Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal message As String) As Integer
        '    Dim buffer As IntPtr = IntPtrAlloc(message)

        '    Dim copyData As New CopyDataStruct With {
        '        .Type = IntPtr.Zero,
        '        .Data = buffer,
        '        .ByteLength = message.Length
        '    }
        '    Dim copyDataBuff As IntPtr = IntPtrAlloc(copyData)

        '    Dim rst = Win32API.UnsafeNativeMethods.SendMessage(hWnd, WM_COPYDATA, 0, copyDataBuff)
        '    IntPtrFree(copyDataBuff)
        '    IntPtrFree(buffer)

        '    Return rst
        'End Function

        ''' <summary>
        ''' 向进程发送消息
        ''' </summary>
        ''' <param name="hWnd"></param>
        ''' <param name="message"></param>
        ''' <returns></returns>
        Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal message As String) As Integer
            Dim msgByteArr As Byte() = s_Encoding.GetBytes(message)
            Dim len As Integer = msgByteArr.Length
            Dim copyData As CopyDataStruct
            copyData.Type = New IntPtr(100)
            copyData.Data = message
            copyData.ByteLength = len + 1

            Dim rst = Win32API.UnsafeNativeMethods.SendMessage(hWnd, WM_COPYDATA, 0, copyData)

            IntPtrFree(copyData.Type)

            Return rst
        End Function

        ''' <summary>
        ''' 向进程发送消息
        ''' </summary>
        ''' <param name="hWnd"></param>
        ''' <param name="message"></param>
        ''' <returns></returns>
        Public Shared Function PostMessage(ByVal hWnd As IntPtr, ByVal message As String) As Boolean
            Dim msgByteArr As Byte() = s_Encoding.GetBytes(message)
            Dim len As Integer = msgByteArr.Length
            Dim copyData As CopyDataStruct
            copyData.Type = New IntPtr(100)
            copyData.Data = message
            copyData.ByteLength = len + 1

            Dim rst = Win32API.UnsafeNativeMethods.PostMessage(hWnd, WM_COPYDATA, 0, copyData)

            Return rst
        End Function

        '''' <summary>
        '''' 处理来自进程的消息
        '''' </summary>
        '''' <param name="m"></param>
        '''' <returns></returns>
        'Public Shared Function GetMessage(ByRef m As Message) As String
        '    Dim copyData As New CopyDataStruct()
        '    Dim copyDataType As Type = copyData.GetType
        '    copyData = DirectCast(m.GetLParam(copyDataType), CopyDataStruct)

        '    Dim rst = Marshal.PtrToStringAnsi(copyData.Data)
        '    IntPtrFree(copyData.Data)

        '    Return rst
        'End Function

        ''' <summary>
        ''' 处理来自进程的消息
        ''' </summary>
        ''' <param name="m"></param>
        ''' <returns></returns>
        Public Shared Function GetMessage(ByRef m As Message) As String
            Dim copyData As New CopyDataStruct()
            copyData = DirectCast(m.GetLParam(copyData.GetType), CopyDataStruct)

            IntPtrFree(copyData.Type)

            Return copyData.Data
        End Function

        Public Shared Sub ModifyEncoding(ByVal encoding As Text.Encoding)
            s_Encoding = encoding
        End Sub
#End Region
    End Class

End Namespace
