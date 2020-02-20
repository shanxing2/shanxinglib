Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms

Imports ShanXingTech.Win32API
Imports ShanXingTech.Win32API.UnsafeNativeMethods

Namespace ShanXingTech
	''' <summary>
	''' 这个人很懒，什么都没写
	''' </summary>
	Public NotInheritable Class Windows2
		' 明文：神即道, 道法自然, 如来|闪星网络信息科技 ShanXingTech Q2287190283
		' 算法：古典密码中的有密钥换位密码 密钥：ShanXingTech
		Public Const ShanXingTechQ2287190283 = "神闪X7,SQB道信T2道网N9来A2D如H2C然技HA即星I1|N8E法息E8,络G0自科C3"

		''' <summary>
		''' a combobox is made up of three controls, a button, a list and textbox; 
		''' not support 64 bit
		''' </summary>
		''' <param name="control"></param>
		''' <returns></returns>
		Public Shared Function GetComboBoxInfo(control As Control) As ComBoBoxInfo
			Dim info As ComBoBoxInfo
			info.cbSize = Marshal.SizeOf(info)
			UnsafeNativeMethods.GetComboBoxInfo(control.Handle, info)

			Return info
		End Function

		''' <summary>
		''' <para>绘制居中显示的提示，这是一个玄学函数</para> 
		''' </summary>
		''' <param name="parentCtrl">父容器</param>
		''' <param name="tips">提示文字</param>
		''' <param name="closeMillisecond">延时关闭时间，单位毫秒</param>
		''' <param name="operateSucceed">操作是否成功(默认成功)   
		''' <para>操作成功——绿底白字（淡绿色）</para> 
		''' <para>操作失败——红底白字(淡珊瑚色)</para> </param>
		''' <param name="enabledSound">是否播放提示音</param>
		Public Shared Sub DrawTipsTask(ByVal parentCtrl As Control, ByVal tips As String, ByVal closeMillisecond As Integer, ByVal operateSucceed As Boolean, ByVal enabledSound As Boolean)
			' 线程安全操作 创建lbl必须要在invoke里面 否则parentCtl就操作不了 其他线程创建的控件
			If parentCtrl Is Nothing OrElse
			parentCtrl.IsDisposed OrElse
			Not parentCtrl.Visible Then Return

			parentCtrl.Invoke(Sub() DrawTipsAction(parentCtrl, tips, closeMillisecond, operateSucceed, enabledSound))
		End Sub

		Private Shared Sub DrawTipsAction(ByVal parentCtrl As Control, ByVal tips As String, ByVal closeMillisecond As Integer, ByVal operateSucceed As Boolean, ByVal enabledSound As Boolean)
			Dim lbl As New Label
			Dim fontFamily As New FontFamily("微软雅黑")
			Dim font As New Font(fontFamily, 20)
			Dim sizeOfString As New SizeF()
			Using g As Graphics = lbl.CreateGraphics
				' 测量字体宽度，汉字和英文的高度是一样的
				sizeOfString = g.MeasureString(tips, font)
			End Using

			parentCtrl.SuspendLayout()
			'lbl.AutoSize = True
			' 超过标签宽度的文本 不显示出来 用...表示 
			' 然后鼠标移到标签上会tip的方式显示全部文本
			lbl.AutoEllipsis = True
			' 根据传入的 操作结果 选择 提示底色
			' 同时发出相应的提示音
			If operateSucceed Then
				lbl.BackColor = Color.LightGreen
				If enabledSound Then My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Asterisk)
			Else
				lbl.BackColor = Color.LightCoral
				If enabledSound Then My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Hand)
			End If
			' 动态添加 并设置标签属性
			lbl.Font = font
			fontFamily.Dispose()
			font.Dispose()
			lbl.ForeColor = Color.White

			' 根据父容器宽度 控制数据显示行数（全部数据显示出来）
			' 如果 字符串长度小于窗体宽度 则居中显示
			If sizeOfString.Width > parentCtrl.Width Then
				Dim lineOfString As Single = 0F
				' 垂直居中显示
				' 计算相对于当前父容器的宽度，如果完全显示出Tips在label上需要多少行
				lineOfString = sizeOfString.Width / (parentCtrl.ClientRectangle.Width - (parentCtrl.Width - parentCtrl.ClientRectangle.Width) * 2)
				' 标签的宽度=父容器的宽度，标签的高度=字符串的高度*字符串行数
				Dim totalHeightOfString = CInt(sizeOfString.Height) * CInt(lineOfString)
				' 如果字符串的总高度超过了父容器的总高度，则设置label的大小为父窗体容器客户区的大小
				' 由于用了 lbl.AutoEllipsis = True 属性，所有后面多出的字符串会显示成...
				' 鼠标移动上去会有提示
				' 否则，label的宽度=父容器的客户区宽度，label的高度=字符串的总高度
				If totalHeightOfString > parentCtrl.ClientRectangle.Height Then
					lbl.Size = parentCtrl.ClientRectangle.Size
					lbl.Location = New Point(0, 0)
				Else
					lbl.Size = New Size(parentCtrl.ClientRectangle.Width, totalHeightOfString)
					lbl.Location = New Point(0, CInt(parentCtrl.ClientRectangle.Height / 2 - lbl.Height / 2))
				End If
			Else
				lbl.AutoSize = True

				' 居父容器中显示
				lbl.Location = New Point(CInt(parentCtrl.ClientRectangle.Width / 2 - sizeOfString.Width / 2), CInt(parentCtrl.ClientRectangle.Height / 2 - sizeOfString.Height / 2))
			End If

			lbl.Name = "drawTips"
			lbl.Text = tips

			' 添加到窗体
			parentCtrl.Controls.Add(lbl)
			' 置顶功能必须是添加控件到窗体之后 才能设置
			lbl.BringToFront()
			parentCtrl.ResumeLayout()
			parentCtrl.Refresh()

			Using cts As New CancellationTokenSource
				Dim token As CancellationToken = cts.Token
				cts.CancelAfter(closeMillisecond)

				Dim tsk = Task.Delay(closeMillisecond, token)
				tsk.ContinueWith(
				Sub(taskCompletion As Task)
					' ###########################################################
					' # 在ContinueWith块中，调试命中断点的情况时，Intellisense无法正常工作
					' # 表现为 鼠标停留到变量上时返回的是变量的默认值而不是实时值
					' # 自动窗口&局部变量窗口&即时窗口工作正常
					' ###########################################################
					' TODO something
					If Not parentCtrl.IsHandleCreated OrElse
						parentCtrl.Disposing OrElse
						parentCtrl.IsDisposed OrElse
						cts.IsCancellationRequested Then
						Return
					End If

					parentCtrl.Invoke(
					Sub()
						parentCtrl.Controls.Remove(lbl)
						If lbl.Visible Then lbl.Hide()
						parentCtrl?.Refresh()
						lbl.Dispose()
					End Sub)
				End Sub, TaskContinuationOptions.None)
			End Using
		End Sub

		''' <summary>
		''' <para>绘制居中显示的提示，这是一个玄学函数</para> 
		''' </summary>
		''' <param name="parentCtrl">父容器</param>
		''' <param name="tips">提示文字</param>
		''' <param name="closeMillisecond">延时关闭时间，单位毫秒</param>
		''' <param name="operateSucceed">操作是否成功(默认成功)   
		''' <para>操作成功——绿底白字（淡绿色）</para> 
		''' <para>操作失败——红底白字(淡珊瑚色)</para> </param>
		Public Shared Sub DrawTipsTask(ByVal parentCtrl As Control, ByVal tips As String, ByVal closeMillisecond As Integer, ByVal operateSucceed As Boolean)
			DrawTipsTask(parentCtrl, tips, closeMillisecond, operateSucceed, True)
		End Sub

		''' <summary>
		''' 参数 <paramref name="markColor"/> 的值为 True时，行号为 <paramref name="rowIndex"/>，列号为 <paramref name="colIndex"/> 的单元格背景设置成HotPink
		''' </summary>
		''' <param name="dgv"></param>
		''' <param name="rowIndex"></param>
		''' <param name="colIndex"></param>
		''' <param name="markColor"></param>
		Public Shared Sub SetCellsBackColor(ByVal dgv As DataGridView, ByVal rowIndex As Integer, ByVal colIndex As Integer, ByVal markColor As Boolean)
			If markColor Then
				dgv.Rows(rowIndex).Cells(colIndex).Style.BackColor = Color.HotPink
				dgv.Rows(rowIndex).Cells(colIndex).Style.ForeColor = Color.White
			Else
				dgv.Rows(rowIndex).Cells(colIndex).Style.BackColor = Color.White
				dgv.Rows(rowIndex).Cells(colIndex).Style.ForeColor = Color.Black
			End If
		End Sub


		''' <summary>
		''' 行号为 <paramref name="rowIndex"/>  的单元格值小于等于 <paramref name="compareValue"/> 时，单元格背景设置成HotPink
		''' </summary>
		''' <param name="dgv"></param>
		''' <param name="rowIndex"></param>
		''' <param name="colIndex"></param>
		''' <param name="compareValue"></param>
		Public Shared Sub SetCellsBackColor(ByVal dgv As DataGridView, ByVal rowIndex As Integer, ByVal colIndex As Integer， ByVal compareValue As Integer)
			If (dgv.RowCount = 1 AndAlso dgv.Rows(0).Cells(1).Value Is Nothing) OrElse dgv.RowCount = 0 Then
				'
			Else
				Dim value = dgv.Rows(rowIndex).Cells(colIndex).Value
				If value IsNot Nothing AndAlso Not DBNull.Value.Equals(value) AndAlso
					CInt(value) <= compareValue Then
					dgv.Rows(rowIndex).Cells(colIndex).Style.BackColor = Color.HotPink
					dgv.Rows(rowIndex).Cells(colIndex).Style.ForeColor = Color.White
				Else
					dgv.Rows(rowIndex).Cells(colIndex).Style.BackColor = Color.White
					dgv.Rows(rowIndex).Cells(colIndex).Style.ForeColor = Color.Black
				End If
			End If
		End Sub

		Public Delegate Function CompareValue() As Boolean

		''' <summary>
		''' 当<paramref name="compare"/> 为条件成立时，行号为 <paramref name="rowIndex"/>，列号为 <paramref name="colIndex"/> 的单元格背景设置成HotPink
		''' </summary>
		''' <param name="dgv"></param>
		''' <param name="colIndex"></param>
		''' <param name="compare"></param>
		Public Shared Sub SetCellsBackColor(ByVal dgv As DataGridView, ByVal rowIndex As Integer, ByVal colIndex As Integer， ByVal compare As CompareValue)
			If (dgv.RowCount = 1 AndAlso dgv.Rows(0).Cells(1).Value Is Nothing) OrElse dgv.RowCount = 0 Then
				'
			Else
				If compare() Then
					dgv.Rows(rowIndex).Cells(colIndex).Style.BackColor = Color.HotPink
					dgv.Rows(rowIndex).Cells(colIndex).Style.ForeColor = Color.White
				Else
					dgv.Rows(rowIndex).Cells(colIndex).Style.BackColor = Color.White
					dgv.Rows(rowIndex).Cells(colIndex).Style.ForeColor = Color.Black
				End If
			End If
		End Sub

		''' <summary>
		''' 所有行，列号为 <paramref name="colIndex"/> 的单元格值等于-1时，单元格背景设置成HotPink
		''' </summary>
		''' <param name="dgv"></param>
		''' <param name="colIndex"></param>
		Public Shared Sub SetCellsBackColor(ByVal dgv As DataGridView, ByVal colIndex As Integer)
			If (dgv.RowCount = 1 AndAlso dgv.Rows(0).Cells(1).Value Is Nothing) OrElse dgv.RowCount = 0 Then
				'
			Else
				Dim rowIndex = 0
				While rowIndex < dgv.RowCount
					Dim value = dgv.Rows(rowIndex).Cells(colIndex).Value
					If value IsNot Nothing AndAlso Not DBNull.Value.Equals(value) AndAlso
						CStr(value) = "-1" Then
						dgv.Rows(rowIndex).Cells(colIndex).Style.BackColor = Color.HotPink
						dgv.Rows(rowIndex).Cells(colIndex).Style.ForeColor = Color.White
					Else
						dgv.Rows(rowIndex).Cells(colIndex).Style.BackColor = Color.White
						dgv.Rows(rowIndex).Cells(colIndex).Style.ForeColor = Color.Black
					End If

					rowIndex += 1
				End While
			End If
		End Sub

		''' <summary>
		''' 所有行，列号为 <paramref name="colIndex"/> 的单元格值小于等于 <paramref name="compareValue"/> 时，单元格背景设置成HotPink
		''' </summary>
		''' <param name="dgv"></param>
		''' <param name="colIndex"></param>
		Public Shared Sub SetCellsBackColor(ByVal dgv As DataGridView, ByVal colIndex As Integer， ByVal compareValue As Integer)
			If (dgv.RowCount = 1 AndAlso dgv.Rows(0).Cells(1).Value Is Nothing) OrElse dgv.RowCount = 0 Then
				Return
			Else
				Dim rowIndex = 0
				While rowIndex < dgv.RowCount
					Dim value = dgv.Rows(rowIndex).Cells(colIndex).Value
					If value IsNot Nothing AndAlso Not DBNull.Value.Equals(value) AndAlso
						CInt(value) <= compareValue Then
						dgv.Rows(rowIndex).Cells(colIndex).Style.BackColor = Color.HotPink
						dgv.Rows(rowIndex).Cells(colIndex).Style.ForeColor = Color.White
					Else
						dgv.Rows(rowIndex).Cells(colIndex).Style.BackColor = Color.White
						dgv.Rows(rowIndex).Cells(colIndex).Style.ForeColor = Color.Black
					End If

					rowIndex += 1
				End While
			End If
		End Sub

		''' <summary>
		''' 所有行，列号为 <paramref name="colIndex"/> 的单元格值等于 <paramref name="compareValue"/> 时，单元格背景设置成HotPink
		''' </summary>
		''' <typeparam name="T"></typeparam>
		''' <param name="dgv"></param>
		''' <param name="colIndex"></param>
		''' <param name="compareValue"></param>
		Public Shared Sub SetCellsBackColor(Of T)(ByVal dgv As DataGridView, ByVal colIndex As Integer， ByVal compareValue As T)
			If (dgv.RowCount = 1 AndAlso dgv.Rows(0).Cells(1).Value Is Nothing) OrElse dgv.RowCount = 0 Then
				'
			Else
				Dim rowIndex = 0
				While rowIndex < dgv.RowCount
					Dim value = dgv.Rows(rowIndex).Cells(colIndex).Value
					If value IsNot Nothing AndAlso Not DBNull.Value.Equals(value) AndAlso
						CType(value, T).Equals(compareValue) Then
						dgv.Rows(rowIndex).Cells(colIndex).Style.BackColor = Color.HotPink
						dgv.Rows(rowIndex).Cells(colIndex).Style.ForeColor = Color.White
					Else
						dgv.Rows(rowIndex).Cells(colIndex).Style.BackColor = Color.White
						dgv.Rows(rowIndex).Cells(colIndex).Style.ForeColor = Color.Black
					End If

					rowIndex += 1
				End While
			End If
		End Sub

        ''' <summary>
        ''' 高精度延时函数，优点：没有时间限制，几乎不占CPU。
        ''' 如果是在多线程的函数中调用，优先考虑使用 <see cref="Task.Delay"/>
        ''' </summary>
        ''' <param name="milliseconds"></param>
        Public Shared Sub Delay(ByVal milliseconds As Integer)
            If 0 >= milliseconds Then Return

            Dim timeBegin As Integer = TimeGetTime()

            Do
                Application.DoEvents()
                Sleep(1)
            Loop Until milliseconds < (TimeGetTime() - timeBegin)
        End Sub

        ''' <summary>
        ''' 可取消的高精度延时函数，优点：没有时间限制，几乎不占CPU。
        ''' </summary>
        ''' <param name="milliseconds"></param>
        ''' <param name="token"></param>
        Public Shared Sub Delay(ByVal milliseconds As Integer, ByVal token As CancellationToken)
            If 0 >= milliseconds Then Return

            Dim timeBegin As Integer = TimeGetTime()

            Do
                If token.IsCancellationRequested Then Return
                Application.DoEvents()
                Sleep(1)
            Loop Until milliseconds < (TimeGetTime() - timeBegin)
        End Sub

        ''' <summary>
        ''' 随机延时 <paramref name="minValue"/> - <paramref name="maxValue"/>
        ''' 精度根据传入的 <paramref name="accuracy"/> 来确定
        ''' </summary>
        ''' <param name="minValue">最小延时值</param>
        ''' <param name="maxValue">最大延时值</param>
        ''' <param name="accuracy">计时精度，毫秒或者秒</param>
        Public Shared Sub RandDelay(ByVal minValue As Integer, ByVal maxValue As Integer, ByVal accuracy As TimePrecision)
			Dim valueDelay = GetRandMilliSeconds(minValue, maxValue, accuracy)
            Delay(valueDelay)
        End Sub

		''' <summary>
		''' 随机延时 <paramref name="minValue"/> - <paramref name="maxValue"/>
		''' 精度根据传入的 <paramref name="accuracy"/> 来确定
		''' </summary>
		''' <param name="minValue">最小延时值</param>
		''' <param name="maxValue">最大延时值</param>
		''' <param name="accuracy">计时精度，毫秒或者秒</param>
		Public Shared Sub RandDelay(ByVal minValue As Integer, ByVal maxValue As Integer, ByVal accuracy As TimePrecision, ByVal token As CancellationToken)
			Dim valueDelay = GetRandMilliSeconds(minValue, maxValue, accuracy)
            Delay(valueDelay, token)
        End Sub

		''' <summary>
		''' 随机延时 <paramref name="minValue"/> - <paramref name="maxValue"/>
		''' 精度根据传入的 <paramref name="accuracy"/> 来确定
		''' </summary>
		''' <param name="minValue">最小延时值</param>
		''' <param name="maxValue">最大延时值</param>
		''' <param name="accuracy">计时精度，毫秒或者秒</param>
		Public Shared Sub RandDelay(ByVal minValue As Decimal, ByVal maxValue As Decimal, ByVal accuracy As TimePrecision)
			Dim valueDelay = GetRandMilliSeconds(CInt(minValue), CInt(maxValue), accuracy)
            Delay(valueDelay)
        End Sub

		''' <summary>
		''' 随机延时 <paramref name="minValue"/> - <paramref name="maxValue"/>
		''' 精度根据传入的 <paramref name="accuracy"/> 来确定
		''' </summary>
		''' <param name="minValue">最小延时值</param>
		''' <param name="maxValue">最大延时值</param>
		''' <param name="accuracy">计时精度，毫秒或者秒</param>
		Public Shared Sub RandDelay(ByVal minValue As Decimal, ByVal maxValue As Decimal, ByVal accuracy As TimePrecision, ByVal token As CancellationToken)
			Dim valueDelay = GetRandMilliSeconds(CInt(minValue), CInt(maxValue), accuracy)
            Delay(valueDelay, token)
        End Sub

		Private Shared Function GetRandMilliSeconds(ByVal minValue As Integer, ByVal maxValue As Integer, ByVal accuracy As TimePrecision) As Integer
			If minValue = 0 AndAlso maxValue = 0 OrElse (minValue <= 0 OrElse maxValue <= 0) Then
				Return 0
			End If

			Dim rand = New Random(Date.Now.Millisecond)
			Dim valueDelay = rand.Next(minValue, maxValue + 1)
			' 如果传入的单位是秒，则先把延时的秒数转换成毫秒
			If accuracy = TimePrecision.Second Then
				valueDelay *= 1000
			End If

			Return valueDelay
		End Function

		''' <summary>
		''' Flashes a window（Not control） until the window comes to the foreground
		''' Receives the form that will flash.
		''' </summary>
		''' <param name="hWnd">The handle to the window to flash</param>
		''' <returns>whether or not the window needed flashing</returns>
		Public Shared Function FlashWindowEx(hWnd As IntPtr) As Boolean
			Return FlashWindowEx(hWnd, FlashWindow.FLASHW_ALL Or FlashWindow.FLASHW_TIMERNOFG)
		End Function

		''' <summary>
		''' Flashes a window（Not control） until the window comes to the foreground
		''' Receives the form that will flash.
		''' FLASHWINFO.uCount default value is UInteger.MaxValue.
		''' </summary>
		''' <param name="hWnd">The handle to the window to flash</param>
		''' <param name="dwFlags">The flash status of the window</param>
		''' <returns>whether or not the window needed flashing</returns>
		Public Shared Function FlashWindowEx(hWnd As IntPtr, ByVal dwFlags As FlashWindow) As Boolean
			Return FlashWindowEx(hWnd, dwFlags, Integer.MaxValue)
		End Function

		''' <summary>
		''' Flashes a window（Not control） until the window comes to the foreground
		''' Receives the form that will flash
		''' </summary>
		''' <param name="hWnd">The handle to the window to flash</param>
		''' <param name="dwFlags">The flash status of the window</param>
		''' <param name="nCount">闪烁窗口的次数。如果<paramref name="dwFlags"/>设置为<see cref="FlashWindow.FLASHW_STOP"/>的同时设置 <paramref name="nCount"/>为0，则窗口会恢复为初始状态，否则将继续保持橙色状态。</param>
		''' <returns>whether or not the window needed flashing</returns>
		Public Shared Function FlashWindowEx(hWnd As IntPtr, ByVal dwFlags As FlashWindow， ByVal nCount As Integer) As Boolean
			Dim fInfo As New FLASHWINFO()

			fInfo.cbSize = CUInt(Marshal.SizeOf(fInfo))
			fInfo.hwnd = hWnd
			fInfo.dwFlags = dwFlags
			fInfo.uCount = CUInt(nCount)
			fInfo.dwTimeout = 0

			Return Win32API.FlashWindowEx(fInfo)
		End Function
	End Class
End Namespace


