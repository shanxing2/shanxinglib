Imports System.Runtime.InteropServices

Public Module InPtrHelper
    ' 声明为Public的Module是在根命名空间下的 ，外部可以直接访问到

    ''' <summary>
    ''' Allocate a pointer to an arbitrary structure on the global heap.
    ''' 使用完之后必须使用<see cref="IntPtrFree(ByRef IntPtr)"/>释放非托管内存
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="param"></param>
    ''' <returns></returns>
    Public Function IntPtrAlloc(Of T)(param As T) As IntPtr
        ' 申请跟参数param一样大小的内存 
        ' 如果param是string，应该用Marshal.AllocHGlobal(param.Length)
        Dim retval As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(param))
        Marshal.StructureToPtr(param, retval, False)

        Return retval
    End Function

    ''' <summary>
    ''' 从全局堆中释放指向任意结构的指针(或非托管内存)，多与<see cref="IntPtrAlloc(Of T)(T)"/>函数配对使用
    ''' </summary>
    ''' <param name="preAllocated"></param>
    Public Sub IntPtrFree(ByRef preAllocated As IntPtr)
        If IntPtr.Zero = preAllocated Then
            Return
        End If

        Try
            ' 可能会引发句柄无效异常
            Marshal.FreeHGlobal(preAllocated)
        Catch ex As Exception
            Throw
        Finally
            preAllocated = IntPtr.Zero
        End Try
    End Sub

End Module