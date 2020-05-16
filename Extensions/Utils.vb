Imports System.Runtime.CompilerServices
Imports System.Threading.Tasks

Namespace ShanXingTech
    Partial Public Module ExtensionFunc
        ''' <summary>
        ''' 尝试释放<paramref name="bc"/>，不会导致 <see cref="Concurrent.BlockingCollection(Of T).GetConsumingEnumerable()"/> 引发 <see cref="System.ArgumentNullException"/> 异常
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="bc"></param>
        ''' <param name="loopTask"><see cref="Concurrent.BlockingCollection(Of T).GetConsumingEnumerable()"/>所在的Task，类似 'm_DoSomethingTask = Task.Run(AddressOf TryDoSomething)'</param>
        ''' <returns></returns>
        <Extension()>
        Public Function TryRelease(Of T)(ByRef bc As Concurrent.BlockingCollection(Of T), ByRef loopTask As Task) As Boolean
            Try
                If loopTask IsNot Nothing Then
                    If bc IsNot Nothing AndAlso Not bc.IsAddingCompleted Then
                        bc.CompleteAdding()
                    End If

                    While Not loopTask.IsCompleted
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
