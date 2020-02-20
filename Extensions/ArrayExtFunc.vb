Imports System.Runtime.CompilerServices

Namespace ShanXingTech
    Partial Public Module ExtensionFunc
        ''' <summary>
        ''' 从数组中截取返回一个子数组
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="array">原始数组</param>
        ''' <param name="startIndex"> <paramref name="array"/> 中的偏移量，从0开始</param>
        ''' <param name="length">要截取的字节数</param>
        ''' <returns></returns>
        <Extension>
        Public Function SubArray(Of T)(ByRef array() As T, ByVal startIndex As Integer, ByVal length As Integer) As T()
            Dim dst(length - 1) As T
            Buffer.BlockCopy(array, startIndex, dst, 0, length)
            Return dst
        End Function
    End Module
End Namespace