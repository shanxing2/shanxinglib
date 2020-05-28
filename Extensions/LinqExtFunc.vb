Imports System.Linq.Expressions
Imports System.Runtime.CompilerServices

Namespace ShanXingTech
    Partial Public Module ExtensionFunc

        ''' <summary>
        ''' 按字段名为 <paramref name="propertyName"/> 的字段升序排序
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="query"></param>
        ''' <param name="propertyName">需要排序的字段名称</param>
        ''' <returns></returns>
        <Extension>
        Public Function OrderBy(Of T)(query As IQueryable(Of T), propertyName As String) As IOrderedQueryable(Of T)
            Return OrderBy(Of T)(query, propertyName, False)
        End Function

        ''' <summary>
        ''' 按字段名为 <paramref name="propertyName"/> 的字段降序排序
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="query"></param>
        ''' <param name="propertyName">需要排序的字段名称</param>
        ''' <returns></returns>
        <Extension>
        Public Function OrderByDescending(Of T)(query As IQueryable(Of T), propertyName As String) As IOrderedQueryable(Of T)
            Return OrderBy(Of T)(query, propertyName, True)
        End Function

        ''' <summary>
        ''' 按字段名为 <paramref name="propertyName"/> 的字段排序
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="query"></param>
        ''' <param name="propertyName">需要排序的字段名称</param>
        ''' <param name="isDesc">是否降序</param>
        ''' <returns></returns>
        Function OrderBy(Of T)(query As IQueryable(Of T), propertyName As String, isDesc As Boolean) As IOrderedQueryable(Of T)
            Dim methodname = If(isDesc, NameOf(OrderByDescendingInternal), NameOf(OrderByInternal))
            Dim memberProp = GetType(T).GetProperty(propertyName)
            If memberProp Is Nothing Then
                Throw New ArgumentOutOfRangeException($"'{propertyName}'不属于'{query.ToString}'的成员")
            End If
            Dim method = GetType(ExtensionFunc).GetMethod(methodname).MakeGenericMethod(GetType(T), memberProp.PropertyType)

            Return CType(method.Invoke(Nothing, New Object() {query, memberProp}), IOrderedQueryable(Of T))
        End Function

        Private Function OrderByInternal(Of T, TProp)(query As IQueryable(Of T), memberProperty As System.Reflection.PropertyInfo) As IOrderedQueryable(Of T)
            Return query.OrderBy(GetLamba(Of T, TProp)(memberProperty))
        End Function

        Private Function OrderByDescendingInternal(Of T, TProp)(query As IQueryable(Of T), memberProperty As System.Reflection.PropertyInfo) As IOrderedQueryable(Of T)
            Return query.OrderByDescending(GetLamba(Of T, TProp)(memberProperty))
        End Function

        Private Function GetLamba(Of T, TProp)(memberProperty As System.Reflection.PropertyInfo) As Expression(Of Func(Of T, TProp))
            If memberProperty.PropertyType <> GetType(TProp) Then
                Throw New InvalidCastException("类型不匹配，无法转换")
            End If

            Dim thisArg = Expression.Parameter(GetType(T))
            Dim lamba = Expression.Lambda(Of Func(Of T, TProp))(Expression.[Property](thisArg, memberProperty), thisArg)

            Return lamba
        End Function
    End Module
End Namespace