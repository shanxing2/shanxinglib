Namespace ShanXingTech.IO2
	''' <summary>
	''' 注：调试模式时强制中断 或者重新启动会因为实例不能被释放导致内存泄露（实质是异常退出时不会自动调用 <see cref="Dispose"/> 释放资源）
	''' </summary>
	Public NotInheritable Class ExcelEngine
		Implements IDisposable

#Region "字段区"
		Private _Instance As Object
		Private _Exists As Boolean
		Private _ISAM As String
		Private _Filter As String
		Private _Version As Double
		Private _Provider As String
		Private Shared m_Excel As ExcelEngine
#End Region

#Region "属性区"
		''' <summary>
		''' Excel实例，创建后一直存在（任务管理器上可以看到一个Excel.exe进程），直到软件关闭后才销毁
		''' </summary>
		''' <returns></returns>
		Public ReadOnly Property Instance As Object
			Get
				Return _Instance
			End Get
		End Property

		''' <summary>
		''' 存在Excel引擎
		''' </summary>
		''' <returns></returns>
		Public ReadOnly Property Exists As Boolean
			Get
				Return _Exists
			End Get
		End Property

		''' <summary>
		''' 
		''' </summary>
		''' <returns></returns>
		Public ReadOnly Property ISAM As String
			Get
				Return _ISAM
			End Get
		End Property

		''' <summary>
		''' 后缀
		''' </summary>
		''' <returns></returns>
		Public ReadOnly Property Filter As String
			Get
				Return _Filter
			End Get
		End Property

		''' <summary>
		''' 版本
		''' </summary>
		''' <returns></returns>
		Public ReadOnly Property Version As Double
			Get
				Return _Version
			End Get
		End Property

		''' <summary>
		''' 驱动引擎
		''' </summary>
		''' <returns></returns>
		Public ReadOnly Property Provider As String
			Get
				Return _Provider
			End Get
		End Property

		''' <summary>
		''' 生命周期内可重用的Excel实例
		''' </summary>
		''' <returns></returns>
		Public Shared ReadOnly Property Excel As ExcelEngine
			Get
				If m_Excel Is Nothing Then
					m_Excel = New ExcelEngine
				End If

				Return m_Excel
			End Get
		End Property

#End Region

#Region "构造函数区"
		''' <summary>
		''' 类构造函数
		''' 类之内的任意一个静态方法第一次调用时调用此构造函数
		''' 而且程序生命周期内仅调用一次
		''' </summary>
		Sub New()
			GetExcelEngine()
		End Sub
#End Region


#Region "函数区"
		''' <summary>
		''' 获取Excel对象
		''' </summary>
		Private Sub GetExcelEngine()
			' 因为用 CreateObject 的方式尝试创建不存在的Com对象会引发异常，比较耗时，所以改用 Type.GetTypeFromProgID() 方法，不引发异常，先获取对象的类型，如果能获取到说明存在相应的对象
			' 这个时候再用 Activator.CreateInstance 方法去创建对象即可 20170903

			' 因为其他重载都是直接调用三个参数的重载，所以直接用三个参数的重载就好了
			Dim xlAppType = Type.GetTypeFromProgID("ET.Application", Nothing, False)

			If xlAppType Is Nothing Then
				Debug.Print(Logger.MakeDebugString("获取WPS ET对象失败"))

				xlAppType = Type.GetTypeFromProgID("KET.Application", Nothing, False)
			End If

			If xlAppType Is Nothing Then
				Debug.Print(Logger.MakeDebugString("获取WPS KET对象失败"))

				xlAppType = Type.GetTypeFromProgID("EXCEL.Application", Nothing, False)
			End If

			If xlAppType Is Nothing Then
				Debug.Print(Logger.MakeDebugString("获取WPS对象或者Excel对象失败"))
				Return
			Else
				Debug.Print(Logger.MakeDebugString("获取WPS对象或者Excel对象成功"))

				Try
					_Instance = Activator.CreateInstance(xlAppType)
					xlAppType = Nothing
					If _Instance Is Nothing Then Return

					' 到达此步说明能创建WPS对象或者Excel对象成功， 继续执行后续操作
					' 可以使用反射的方式去操作excel对象，也可以像平时引用excel对象一样操作，不过因为是Object，所以不会有相关方法属性等智能提示
					_Instance.Visible = False
					_Version = CDbl(_Instance.Version)
					Select Case _Version
						Case <= 11.0
							_Filter = ".XLS"
							_ISAM = "EXCEL 8.0"
							_Provider = "Microsoft.Jet.OLEDB.4.0"
							_Exists = True
						Case > 11.0
							_Filter = ".XLSX"
							_ISAM = "EXCEL 12.0 Xml"
							_Provider = "Microsoft.ACE.OLEDB.12.0"
							_Exists = True
						Case Else
							_Filter = ".XLS"
							_ISAM = "EXCEL 8.0"
							_Provider = "Microsoft.Jet.OLEDB.4.0"
							_Exists = False
							Debug.Print(Logger.MakeDebugString("获取Excel或者WPS对象失败"))
					End Select
				Catch ex As Exception
					Logger.WriteLine(ex)
				End Try
			End If
		End Sub

#Region "IDisposable Support"
		Private disposedValue As Boolean ' 要检测冗余调用

		' IDisposable
		Protected Sub Dispose(disposing As Boolean)
			If Not disposedValue Then
				If disposing Then
					' TODO: 释放托管状态(托管对象)。

				End If

                ' TODO: 释放未托管资源(未托管对象)并在以下内容中替代 Finalize()。
                ' TODO: 将大型字段设置为 null。
                If _Instance IsNot Nothing Then
                    Try
                        _Instance.Quit
                    Catch ex As Exception
                        Logger.WriteLine(ex)
                    Finally
                        _Instance = Nothing
                    End Try
                End If
            End If
            disposedValue = True
        End Sub

        ' TODO: 仅当以上 Dispose(disposing As Boolean)拥有用于释放未托管资源的代码时才替代 Finalize()。
        Protected Overrides Sub Finalize()
            ' 请勿更改此代码。将清理代码放入以上 Dispose(disposing As Boolean)中。
            Dispose(False)
            MyBase.Finalize()
        End Sub

        ' Visual Basic 添加此代码以正确实现可释放模式。
        Public Sub Dispose() Implements IDisposable.Dispose
            ' 请勿更改此代码。将清理代码放入以上 Dispose(disposing As Boolean)中。
            Dispose(True)
            ' TODO: 如果在以上内容中替代了 Finalize()，则取消注释以下行。
            GC.SuppressFinalize(Me)
        End Sub
#End Region
#End Region
    End Class

End Namespace
