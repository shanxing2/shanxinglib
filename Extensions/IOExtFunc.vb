Imports System.Data.Common
Imports System.Runtime.CompilerServices

Namespace ShanXingTech
    Partial Public Module ExtensionFunc
#Region "DataBase"
        ''' <summary>
        ''' 关闭Reader对应的数据库连接
        ''' </summary>
        ''' <param name="reader"></param>
        <Extension()>
        <Obsolete("已用CommandBehavior.CloseConnection解决DbDataReader的关闭已经对应数据库连接的关闭问题，此方法弃用")>
        Public Sub CloseDbConnectionOld(ByVal reader As DbDataReader)
            If reader Is Nothing Then Return

            ' 通过反射获取数据库软连接对象,并且关闭连接，释放对象占用的资源
            Dim dbCommand As DbCommand = TryCast(reader.GetType().InvokeMember("_command", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.GetField Or Reflection.BindingFlags.Instance, Nothing, reader, Array.Empty(Of Object)()), DbCommand)

            If dbCommand IsNot Nothing Then
                Using dbConnection = dbCommand.Connection
                    dbCommand.Dispose()
                End Using

                dbCommand = Nothing
            End If

            reader.Close()
        End Sub
#End Region
    End Module
End Namespace
