Imports System.Runtime.CompilerServices
Imports System.Threading.Tasks

Namespace ShanXingTech
    Partial Public Module ExtensionFunc
        ''' <summary>
        ''' 尝试释放<paramref name="bc"/>，不会导致 <see cref="Concurrent.BlockingCollection(Of T).GetConsumingEnumerable()"/> 引发 <see cref="System.ArgumentNullException"/> 异常
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="bc"></param>
        ''' <param name="loopTask"><see cref="Concurrent.BlockingCollection(Of T).GetConsumingEnumerable()"/>所在的Task，类似 'm_DoSomethingTask = Task.Run(AddressOf TryDoSomething)',如果TryDoSomething是异步的，那么必须为'Task.Run(Sub() TryDoSomethingAsync.GetAwaiter.GetResult())'这种形式，如果使用AddressOf会变成同步，使用Await关键字会使任务未按预期运行而导致<see cref="Concurrent.BlockingCollection(Of T).GetConsumingEnumerable()"/>引发 <see cref="ArgumentNullException"/>异常</param>
        ''' <returns></returns>
        <Extension()>
        Public Function TryRelease(Of T)(ByRef bc As Concurrent.BlockingCollection(Of T), ByRef loopTask As Task) As Boolean
            Try
                If loopTask IsNot Nothing Then
                    If bc IsNot Nothing AndAlso Not bc.IsAddingCompleted Then
                        bc.CompleteAdding()
                    End If

                    While Not bc.IsAddingCompleted OrElse
                          Not loopTask.IsCompleted
                        Windows2.Delay(100)
                    End While
                    loopTask.Dispose()
                    loopTask = Nothing

                    bc.Dispose()
                    bc = Nothing
                End If
            Catch ex As Exception
                Return False
            End Try

            Return True
        End Function
    End Module
End Namespace
