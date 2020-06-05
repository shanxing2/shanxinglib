Imports System.Runtime.CompilerServices

Imports ShanXingTech

Namespace ShanXingTech

    Partial Public Module ExtensionFunc
        ' 下面的扩展函数都是为 Conf 类服务的
#Region "函数区"
        ''' <summary>
        ''' 保存配置（持久化配置信息）。不需要每个属性更改都执行一次，建议并且需要保证在主窗体关闭的时候调用一次。
        ''' </summary>
        ''' <param name="instance">继承<see cref="ConfBase"/>类的实例，可以为未初始化状态。</param>
        <Extension>
        Public Sub Save(ByVal instance As ConfBase)
            If instance Is Nothing Then Return

            Try
                Dim json = instance.Serialize.ToHexString(UpperLowerCase.Upper, True)
                If json.Length = 0 Then Return

                Dim productInfo = GetProductInfo(instance)
                Dim confPath = GetProductConfPath(productInfo.ProductName, productInfo.IsCallByOwn)
                IO2.Writer.WriteText(confPath, json, IO.FileMode.Create, IO2.CodePage.UTF8)
            Catch ex As Exception
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' 尝试获取 <paramref name="instance"/> 存储在本地的实例序列化字符串
        ''' </summary>
        ''' <param name="instance">继承<see cref="ConfBase"/>类的实例，可以为未初始化状态。</param>
        ''' <returns></returns>
        <Extension>
        Public Function GetLocalInstanceSerialization(ByVal instance As ConfBase) As String
            Dim productInfo = GetProductInfo(instance)
            Dim confPath = GetProductConfPath(productInfo.ProductName, productInfo.IsCallByOwn)

            If Not IO.Directory.Exists(confPath) Then
                IO2.Directory.Create(confPath)
            End If

            Dim json = If(IO.File.Exists(confPath),
            IO2.Reader.ReadFile(confPath, Text.Encoding.UTF8).FromHexString(True),
            New ConfBase With {.ProductName = productInfo.ProductName}.Serialize)

            Return json
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="productName"></param>
        ''' <param name="isCallByOwn">是否自己调用自己，或者是其他程序集调用</param>
        ''' <returns></returns>
        Private Function GetProductConfPath(ByVal productName As String, ByVal isCallByOwn As Boolean) As String
            ' 最终路径 C:\ProgramData\ShanXingTech\{产品名}\conf.json
            Dim confPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)
            confPath = If("\"c = confPath.Chars(confPath.Length - 1),
           $"{confPath}ShanXingTech{productName}{If(isCallByOwn, "_", "\")}conf.conf",
           $"{confPath}\ShanXingTech\{productName}{If(isCallByOwn, "_", "\")}conf.conf")

            Return confPath
        End Function

        Private Function GetProductInfo(ByVal instance As ConfBase) As (ProductName As String, IsCallByOwn As Boolean)
            Dim currentAssemblyName As String
            Dim topCallingAssemblyName As String
            ' 实例化前后的堆栈信息不一样，所以不能用一样的方法获取
            If instance Is Nothing Then
                Dim st = New StackTrace
                If st.FrameCount < 2 Then
                    Throw New TypeInitializationException(NameOf(instance), New Exception($"不支持的继承深度，只支持从 {NameOf(ConfBase)} 继承的类调用此辅助工具。"))
                End If
                currentAssemblyName = st.GetFrames(0).GetMethod.DeclaringType.Assembly.GetName.Name
                topCallingAssemblyName = st.GetFrames(2).GetMethod.DeclaringType.Assembly.GetName.Name
            Else
                Dim t = instance.GetType
                currentAssemblyName = Reflection.Assembly.GetExecutingAssembly.GetName.Name
                topCallingAssemblyName = t.Assembly.GetName.Name
            End If

            ' 如果程序集名一样，说明是类库自身 Conf 调用,需要加上入口程序集的名称，这样每个调用类库的程序都会有一个属于自己的conf文件，而不是共享
            ' 如果需要所有程序共享类库设置，那就直接返回 currentAssemblyName 即可
            Dim isCallByOwn = currentAssemblyName.Equals(topCallingAssemblyName, StringComparison.OrdinalIgnoreCase)
            If isCallByOwn Then
                Return ($"{Reflection.Assembly.GetEntryAssembly.GetName.Name}\{currentAssemblyName}", isCallByOwn)
            Else
                Return (topCallingAssemblyName, isCallByOwn)
            End If
        End Function

        ''' <summary>
        ''' 移除入口程序（主程序，调用验证模块的软件）本地Conf文件
        ''' </summary>
        ''' <param name="instance">继承<see cref="ConfBase"/>类的实例，可以为未初始化状态。</param>
        ''' <returns></returns>
        <Extension>
        Public Function TryRemoveLocalConf(ByVal instance As ConfBase) As Boolean
            Try
                Dim productInfo = GetProductInfo(instance)
                Dim confPath = GetProductConfPath(productInfo.ProductName, productInfo.IsCallByOwn)
                If IO.File.Exists(confPath) Then
                    IO.File.Delete(confPath)
                End If
            Catch ex As Exception
                Logger.WriteLine(ex)
                Return False
            End Try

            Return True
        End Function
#End Region
    End Module
End Namespace
