Imports System.Drawing
Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

Imports ShanXingTech.Text2
Imports ShanXingTech.Win32API

Namespace ShanXingTech
    Partial Public Module ExtensionFunc

#Region "字段区"
        Private m_EventDic As New Concurrent.ConcurrentDictionary(Of String, MouseLeaveAction)
        Private m_toolTip As ToolTip
#End Region

#Region "构造函数"
        Sub New()
            m_toolTip = New ToolTip
        End Sub
#End Region

#Region "函数区"
        <Extension()>
        Public Function GetNames(ByVal columns As DataColumnCollection) As String()
            Dim cols(columns.Count - 1) As String
            For i = 0 To columns.Count - 1
                cols(i) = columns(i).ColumnName
            Next

            Return cols
        End Function

        ''' <summary>
        ''' 从 DataTable 中按条件找到某行
        ''' </summary>
        ''' <param name="table">一个有数据的DataTable</param>
        ''' <param name="columnIndex">列索引</param>
        ''' <param name="cellValue">单元格值</param>
        ''' <returns></returns>
        <Extension()>
        Public Function GetRow(Of T)(ByVal table As DataTable, ByVal columnIndex As Integer, ByVal cellValue As T) As DataRow
            Dim row As DataRow = Nothing

            ' 由于函数会被用于并行，所以不能用for each 否则会引发
            ' 集合已修改；枚举操作可能无法执行
            If table.Rows.Count = 0 Then Return Nothing

            For rowIndex = table.Rows.Count - 1 To 0 Step -1
                ' 统一转换成String再比较，否则内部会自动转换为相应的类型，而不是转换为Object
                ' 比如 如果 cellValue 是传入Double类型的，那么内部会自动把table.Rows(rowIndex).Item(columnIndex)转换为Double类型
                ' 如果不能转换的话就会引发转换无效异常
                If cellValue = table.Rows(rowIndex).Item(columnIndex) Then
                    row = table.Rows(rowIndex)
                    Exit For
                End If
            Next

            Return row
        End Function

        ''' <summary>
        ''' 从 DataTable 中按条件找到某行
        ''' </summary>
        ''' <param name="table">一个有数据的DataTable</param>
        ''' <param name="columnName">列名</param>
        ''' <param name="cellValue">单元格值</param>
        ''' <returns></returns>
        <Extension()>
        Public Function GetRow(Of T)(ByVal table As DataTable, ByVal columnName As String, ByVal cellValue As T) As DataRow
            Dim row As DataRow = Nothing

            ' 先通过传入的列名 columnName 找到 列索引 columnIndex，再去操作（用列索引比用列名高效）
            Dim columnIndex = -1
            Dim tempColumnIndex = 0
            For Each col As DataColumn In table.Columns
                If col.ColumnName = columnName Then
                    columnIndex = tempColumnIndex
                    Exit For
                End If

                tempColumnIndex += 1
            Next

            ' 如果找不到对应的列索引，则返回 Nothing
            If columnIndex = -1 Then Return Nothing

            row = GetRow(table， columnIndex， cellValue)

            Return row
        End Function

        ''' <summary>
        ''' 根据列<paramref name="colIndex"/>的<paramref name="findCellValue"/>值获取所在的行Index
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="dgv"></param>
        ''' <param name="colIndex">需要查找的列</param>
        ''' <param name="findCellValue">需要查找的值</param>
        ''' <returns></returns>
        <Extension()>
        Public Function GetRowIndex(Of T)(ByVal dgv As DataGridView, ByVal colIndex As Integer, ByVal findCellValue As T) As Integer
            Dim funcRst As Integer

            ' 由于函数会被用于并行，所以不能用for each 否则会引发
            ' 集合已修改；枚举操作可能无法执行
            If dgv.Rows.Count = 0 Then Return -1

            For rowIndex = dgv.Rows.Count - 1 To 0 Step -1
                If dgv.Rows(rowIndex).Cells(colIndex).Value = findCellValue Then
                    funcRst = dgv.Rows(rowIndex).Index
                    Exit For
                End If
            Next

            Return funcRst
        End Function

        '''' <summary>
        '''' 从 DataTable 中按条件找到某行
        '''' </summary>
        '''' <param name="table">一个有数据的DataTable</param>
        '''' <param name="columnIndex">列索引</param>
        '''' <param name="cellValue">单元格值</param>
        '''' <returns></returns>
        '<Extension()>
        'Public Function GetRow(ByVal table As DataTable, ByVal columnIndex As Integer, ByVal cellValue As Object) As DataRow
        '    Dim row As DataRow = Nothing

        '    ' 由于函数会被用于并行，所以不能用for each 否则会引发
        '    ' 集合已修改；枚举操作可能无法执行

        '    If table.Rows.Count = 0 Then Return Nothing

        '    For rowIndex = table.Rows.Count - 1 To 0 Step -1
        '        ' 同一转换成String再比较，否则内部会自动转换为相应的类型，而不是转换为Object
        '        ' 比如 如果 cellValue 是传入Double类型的，那么内部会自动把table.Rows(rowIndex).Item(columnIndex)转换为Double类型
        '        ' 如果不能转换的话就会引发转换无效异常
        '        If CStr(cellValue) = CStr(table.Rows(rowIndex).Item(columnIndex)) Then
        '            row = table.Rows(rowIndex)
        '            Exit For
        '        End If
        '    Next

        '    Return row
        'End Function

        '''' <summary>
        '''' 从 DataTable 中按条件找到某行
        '''' </summary>
        '''' <param name="table">一个有数据的DataTable</param>
        '''' <param name="columnName">列名</param>
        '''' <param name="cellValue">单元格值</param>
        '''' <returns></returns>
        '<Extension()>
        'Public Function GetRow(ByVal table As DataTable, ByVal columnName As String, ByVal cellValue As Object) As DataRow
        '    Dim row As DataRow = Nothing

        '    ' 先通过传入的列名 columnName 找到 列索引 columnIndex，再去操作（用列索引比用列名高效）
        '    Dim columnIndex = -1
        '    Dim tempColumnIndex = 0
        '    For Each col As DataColumn In table.Columns
        '        If col.ColumnName = columnName Then
        '            columnIndex = tempColumnIndex
        '            Exit For
        '        End If

        '        tempColumnIndex += 1
        '    Next

        '    ' 如果找不到对应的列索引，则返回 Nothing
        '    If columnIndex = -1 Then Return Nothing

        '    row = GetRow(table， columnIndex， cellValue)

        '    Return row
        'End Function

        ''' <summary>
        ''' 删除当前选中的行以及对应的数据源（DataSource）中的行
        ''' </summary>
        ''' <param name="dgv">当前操作的DataGridView</param>
        ''' <param name="columnIndex">具有唯一值的列（无重复值）。<para>注：如果使用 DataGridView.Columns.Add 或者 DataGridView.Columns.Insert方法动态添加列，建议调用 <see cref="DeleteRowAfterSortOrNot(ByRef DataGridView, String)"/> 版本，除非你能确定索引 <paramref name="columnIndex"/> 对应的列就是你需要的列 </para> </param>
        <Extension()>
        Public Sub DeleteRowAfterSortOrNot(ByRef dgv As DataGridView, ByVal columnIndex As Integer)
            Dim dt = TryCast(dgv.DataSource, DataTable)
            If dt Is Nothing Then
                Throw New NullReferenceException(String.Format(My.Resources.NullReference, NameOf(dgv.DataSource)))
            End If

            Dim findCellValue = dgv.Rows(dgv.CurrentRow.Index).Cells(columnIndex).Value
            DeleteRowAfterSortOrNot(dt, columnIndex, findCellValue)
        End Sub


        ''' <summary>
        ''' 删除当前选中的行以及对应的数据源（DataSource）中的行
        ''' </summary>
        ''' <param name="dgv">当前操作的DataGridView</param>
        ''' <param name="columnIndex">具有唯一值的列（无重复值）。<para>注：如果使用 DataGridView.Columns.Add 或者 DataGridView.Columns.Insert方法动态添加列，建议调用 <see cref="DeleteRowAfterSortOrNot(ByRef DataGridView, String)"/> 版本，除非你能确定索引 <paramref name="columnIndex"/> 对应的列就是你需要的列 </para> </param>
        <Extension()>
        Public Sub DeleteRowAfterSortOrNot(Of T)(ByRef dgv As DataGridView, ByVal columnIndex As Integer, ByVal findCellValue As T)
            Dim dt = TryCast(dgv.DataSource, DataTable)
            If dt Is Nothing Then
                Throw New NullReferenceException(String.Format(My.Resources.NullReference, NameOf(dgv.DataSource)))
            End If

            DeleteRowAfterSortOrNot(dt, columnIndex, findCellValue)
        End Sub

        ''' <summary>
        ''' 删除当前选中的行以及对应的数据源（DataSource）中的行
        ''' </summary>
        ''' <param name="dgv">当前操作的DataGridView</param>
        ''' <param name="columnName">具有唯一值的列（无重复值）</param>
        <Extension()>
        Public Sub DeleteRowAfterSortOrNot(ByRef dgv As DataGridView, ByVal columnName As String)
            Dim dt = TryCast(dgv.DataSource, DataTable)
            If dt Is Nothing Then
                Throw New NullReferenceException(String.Format(My.Resources.NullReference, NameOf(dgv.DataSource)))
            End If

            ' 有时候用户会使用动态添加列的方式添加到dgv，不管用 DataGridView.Columns.Add 方法，还是 DataGridView.Columns.Insert 明确设置添加的列索引，
            ' 添加完之后，刚添加的Index都是0
            ' 所以会导致用户设置的Index可能会与实际的不一样，我们需要分别获取列 columnName 在Dt跟Dgv中实际的列索引
            ' 20190101
            Dim columnIndexInDt = GetColumnIndex(dt, columnName)
            Dim columnIndexInDgv = GetColumnIndex(dgv, columnName)
            Dim findCellValue = dgv.Rows(dgv.CurrentRow.Index).Cells(columnIndexInDgv).Value
            DeleteRowAfterSortOrNot(dt, columnIndexInDt, findCellValue)
        End Sub

        ''' <summary>
        ''' 删除当前选中的行以及对应的数据源（DataSource）中的行
        ''' </summary>
        ''' <param name="dgv">当前操作的DataGridView</param>
        ''' <param name="columnName">具有唯一值的列（无重复值）</param>
        ''' <param name="findCellValue">要查找的值</param>
        <Extension()>
        Public Sub DeleteRowAfterSortOrNot(Of T)(ByRef dgv As DataGridView, ByVal columnName As String, ByVal findCellValue As T)
            Dim dt = TryCast(dgv.DataSource, DataTable)
            If dt Is Nothing Then
                Throw New NullReferenceException(String.Format(My.Resources.NullReference, NameOf(dgv.DataSource)))
            End If

            ' 有时候用户会使用动态添加列的方式添加到dgv，不管用 DataGridView.Columns.Add 方法，还是 DataGridView.Columns.Insert 明确设置添加的列索引，
            ' 添加完之后，刚添加的Index都是0
            ' 所以会导致用户设置的Index可能会与实际的不一样，我们需要分别获取列 columnName 在Dt跟Dgv中实际的列索引
            ' 20190101
            Dim columnIndexInDt = GetColumnIndex(dt, columnName)
            DeleteRowAfterSortOrNot(dt, columnIndexInDt, findCellValue)
        End Sub

        ''' <summary>
        ''' 删除行
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="colIndex"></param>
        ''' <param name="findCellValue"></param>
        <Extension()>
        Public Sub DeleteRowAfterSortOrNot(ByRef dt As DataTable, ByVal colIndex As Integer, ByVal findCellValue As Object)
            DeleteRowAfterSortOrNot(dt, colIndex, findCellValue)
        End Sub

        ''' <summary>
        ''' 删除行
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="colIndex"></param>
        ''' <param name="findCellValue"></param>
        <Extension()>
        Public Sub DeleteRowAfterSortOrNot(Of T)(ByRef dt As DataTable, ByVal colIndex As Integer, ByVal findCellValue As T)
            Dim row = dt.GetRow(colIndex, findCellValue)
            If row Is Nothing Then
                Throw New NullReferenceException(String.Format(My.Resources.NullReference, NameOf(row)))
            End If
            If row.RowState <> DataRowState.Deleted AndAlso row.RowState <> DataRowState.Detached Then
                row.Delete()
            End If
            If row.RowState <> DataRowState.Detached Then
                row.AcceptChanges()
            End If
        End Sub

        ''' <summary>
        ''' 删除行
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="columnName"></param>
        ''' <param name="findCellValue"></param>
        <Extension()>
        Public Sub DeleteRowAfterSortOrNot(ByRef dt As DataTable, ByVal columnName As String, ByVal findCellValue As Object)
            DeleteRowAfterSortOrNot(dt, columnName, findCellValue)
        End Sub

        ''' <summary>
        ''' 删除行
        ''' </summary>
        ''' <param name="dt"></param>
        ''' <param name="columnName"></param>
        ''' <param name="findCellValue"></param>
        <Extension()>
        Public Sub DeleteRowAfterSortOrNot(Of T)(ByRef dt As DataTable, ByVal columnName As String, ByVal findCellValue As T)
            Dim row = dt.GetRow(columnName, findCellValue)
            If row Is Nothing Then
                Throw New NullReferenceException(String.Format(My.Resources.NullReference, NameOf(row)))
            End If
            If row.RowState <> DataRowState.Deleted AndAlso row.RowState <> DataRowState.Detached Then
                row.Delete()
            End If
            If row.RowState <> DataRowState.Detached Then
                row.AcceptChanges()
            End If
        End Sub

        <Extension()>
        Public Function GetColumnIndex(ByRef dgv As DataGridView, ByVal columnName As String) As Integer
            Dim funcRst As Integer

            For i = 0 To dgv.ColumnCount - 1
                If dgv.Columns(i).Name = columnName Then
                    funcRst = dgv.Columns(i).Index
                    Exit For
                End If
            Next

            Return funcRst
        End Function

        <Extension()>
        Public Function GetColumnIndex(ByRef dt As DataTable, ByVal columnName As String) As Integer
            Dim colIndex As Integer = 0
            For i = 0 To dt.Columns.Count - 1
                If dt.Columns(i).ColumnName = columnName Then
                    Exit For
                End If

                colIndex += 1
            Next

            If colIndex > dt.Columns.Count Then
                colIndex = -1
            End If

            Return colIndex
        End Function

        ''' <summary>
        ''' 为类TextBox控件设置水印(Textbox、Combobox(样式为DropDownList时无效)等)
        ''' 暂时不支持多行文本框
        ''' </summary>
        ''' <param name="control"></param>
        ''' <param name="text"></param>
        <Extension()>
        Public Sub SetCueBanner(ByVal control As Control, ByVal text As String)
            If TypeOf control Is ComboBox Then
                ' a combobox is made up of three controls, a button, a list and textbox; 
                ' we want the textbox 
                Dim comBoBoxInfo = Windows2.GetComboBoxInfo(control)
                If IntPtr.Zero = comBoBoxInfo.hwndItem Then Return
                SendMessage(comBoBoxInfo.hwndItem, EM_SETCUEBANNER, 0, text)
            Else
                SendMessage(control.Handle, EM_SETCUEBANNER, 0, text)
            End If
        End Sub

        ''' <summary>
        ''' 在鼠标光标位置处显示窗体
        ''' </summary>
        ''' <param name="childForm">要打开的窗体</param>
        ''' <param name="actionWhileMouseLeave">鼠标移出窗体时，需要执行的动作</param>
        ''' <param name="yOffset">要打开的窗体的初始位置y坐标相对于鼠标y坐标的偏移</param>
        ''' <param name="topMost">是否为模式窗体</param>
        <Extension()>
        Public Sub ShowFollowMousePosition(Of T As Form)(ByRef childForm As T, ByVal actionWhileMouseLeave As MouseLeaveAction, ByVal yOffset As Integer, ByVal topMost As Boolean)
            ' 计算鼠标相对于屏幕的位置，窗体出现的位置以 mousePoint 为参考点
            ' 如果 子窗体的‘下、右’任一部超出了屏幕，则自动向‘上、左’调整到能显示完整窗体为止
            Dim mousePoint = childForm.PointToClient(Cursor.Position)
            Dim childFormX = mousePoint.X
            Dim childFormY = mousePoint.Y + yOffset

            Dim currentWorkingArea = Screen.FromPoint(mousePoint).WorkingArea
            Dim currentWorkingAreaRectangle = New Rectangle(currentWorkingArea.X, currentWorkingArea.Y, currentWorkingArea.Width + Math.Abs(currentWorkingArea.X), currentWorkingArea.Height + Math.Abs(currentWorkingArea.Y))
            ' 下中右 依次检查,上左不用检查
            If childFormY + childForm.Height > currentWorkingAreaRectangle.Height Then
                childFormY = currentWorkingAreaRectangle.Height - childForm.Height
            End If

            ' 主屏幕
            If childFormX > 0 Then
                If childFormX + childForm.Width > currentWorkingAreaRectangle.Width Then
                    childFormX = currentWorkingAreaRectangle.Width - childForm.Width
                End If
            Else ' 次屏幕
                If childForm.Width - Math.Abs(childFormX) > 0 Then
                    childFormX = -childForm.Width
                End If
            End If

            Dim point = New Point With {
                .X = childFormX,
                .Y = childFormY
            }
            childForm.Location = point
            childForm.TopMost = topMost

            HandleMouseLeaveEvent(childForm， actionWhileMouseLeave)

            If topMost Then
                childForm.ShowDialog()
            Else
                childForm.Show()
            End If
        End Sub

        ''' <summary>
        ''' 在鼠标光标位置处显示窗体。
        ''' 鼠标移出窗体时，需要执行的动作，默认为 隐藏<paramref name="childForm"/>。
        ''' 要打开的窗体的初始位置y坐标相对于鼠标y坐标的偏移,默认为 32(Cursor.Size.Height)。
        ''' 是否为模式窗体，默认为 False
        ''' </summary>
        ''' <param name="childForm">要打开的窗体</param>
        <Extension()>
        Public Sub ShowOnMouseCursorPoint(Of T As Form)(ByRef childForm As T)
            ShowFollowMousePosition(childForm, MouseLeaveAction.Hide, 32, False)
        End Sub

        Private Sub HandleMouseLeaveEvent(Of T As Form)(ByRef childForm As T, ByVal actionWhileMouseLeave As MouseLeaveAction)
            Dim eventFulllName = childForm.Name & ".MouseLeave"
            Dim tempActionWhileMouseLeave As MouseLeaveAction
            If Not m_EventDic.TryGetValue(eventFulllName, tempActionWhileMouseLeave) Then
                m_EventDic.TryAdd(eventFulllName, actionWhileMouseLeave)
            Else
                m_EventDic.TryUpdate(eventFulllName, actionWhileMouseLeave, actionWhileMouseLeave)
            End If

            AddHandler childForm.MouseLeave, AddressOf Form_MouseLeaveHandle
        End Sub

        Private Sub Form_MouseLeaveHandle(sender As Object, e As EventArgs)
            Dim form = DirectCast(sender, Form)
            If form Is Nothing Then Return

            ' 当鼠标移出窗口3-5毫秒之后关闭窗口
            Dim mouseFromPoint = form.PointToClient(Control.MousePosition)
            If mouseFromPoint.X < 0 OrElse
                mouseFromPoint.Y < 0 OrElse
                mouseFromPoint.X >= form.Size.Width OrElse
                mouseFromPoint.Y >= form.Size.Height Then
                Dim eventFulllName = form.Name & ".MouseLeave"
                Dim actionWhileMouseLeave As MouseLeaveAction
                If m_EventDic.TryGetValue(eventFulllName, actionWhileMouseLeave) Then
                    RemoveHandler form.MouseLeave, AddressOf Form_MouseLeaveHandle

                    If actionWhileMouseLeave = MouseLeaveAction.Hide Then
                        form.Visible = False
                    ElseIf actionWhileMouseLeave = MouseLeaveAction.Close Then
                        form.Close()
                    End If
                End If
            End If
        End Sub

        ''' <summary>
        ''' 中心对齐
        ''' </summary>
        ''' <typeparam name="T1"></typeparam>
        ''' <typeparam name="T2"></typeparam>
        ''' <param name="control">需要对齐的控件</param>
        ''' <param name="alignControl">参照控件</param>
        ''' <param name="locationX"><paramref name="control"/>新位置的X坐标</param>
        <Extension()>
        Public Sub LocationCenterAlign(Of T1 As Control, T2 As Control)(ByRef control As T1, ByRef alignControl As T2, ByVal locationX As Integer)
            control.Location = New Point(locationX, alignControl.Height \ 2 + alignControl.Top - control.Height \ 2)
        End Sub

        ''' <summary>
        ''' 在控件 <paramref name="control"/> 下方距离一个光标高度位置处显示工具提示文本
        ''' 当鼠标移出控件 <paramref name="control"/> 后会自动删除工具提示文本
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="control"></param>
        ''' <param name="cursor">设置当鼠标指针位于控件上时显示的光标</param>
        ''' <param name="tips">需要显示的提示文本</param>
        <Extension()>
        Public Sub ShowTips(Of T As Control)(ByVal control As T, ByVal cursor As Cursor, ByVal tips As String)
            Dim toolTip As New ToolTip
            control.Cursor = control.Cursor
            toolTip.Show(tips, control， 0, control.Cursor.Size.Height)

            AddHandler control.MouseLeave, AddressOf ToolTipControl_MouseLeaveHandle
            AddHandler control.LostFocus, AddressOf ToolTipControl_LostFocusHandle
        End Sub

        ''' <summary>
        ''' 在控件 <paramref name="control"/> 下方距离一个光标高度位置处显示工具提示文本
        ''' 当鼠标移出控件 <paramref name="control"/> 后会自动删除工具提示文本
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="control"></param>
        ''' <param name="tips">需要显示的提示文本</param>
        <Extension()>
        Public Sub ShowTips(Of T As Control)(ByVal control As T, ByVal tips As String)
            m_toolTip.Show(tips, control， 0, control.Cursor.Size.Height)

            AddHandler control.MouseLeave, AddressOf ToolTipControl_MouseLeaveHandle
        End Sub

        Private Sub ToolTipRemoveEventHandlerHandler(sender As Object, e As EventArgs)
            Dim control = DirectCast(sender, Control)
            RemoveHandler control.MouseLeave, AddressOf ToolTipControl_MouseLeaveHandle
            m_toolTip.Hide(control)
        End Sub

        Private Sub ToolTipControl_MouseLeaveHandle(sender As Object, e As EventArgs)
            ToolTipRemoveEventHandlerHandler(sender, e)
        End Sub

        Private Sub ToolTipControl_LostFocusHandle(sender As Object, e As EventArgs)
            ToolTipRemoveEventHandlerHandler(sender, e)
        End Sub

        ''' <summary>
        ''' 网页滚动条是否滚到到最底下
        ''' </summary>
        ''' <param name="scroll"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function IsOnBottom(ByVal scroll As HtmlElement) As Boolean
            Return ((scroll.OffsetRectangle.Height + scroll.ScrollTop) = scroll.ScrollRectangle.Height)
        End Function

        ''' <summary>
        ''' 获取当前选择行的 某个列的值
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <param name="columnName"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function GetCurrentRowCellValue(Of T As DataGridView)(ByRef dgv As T, ByVal columnName As String) As String
            If dgv.RowCount = 0 OrElse (dgv.RowCount = 1 AndAlso dgv.Item(columnName, 0).Value Is Nothing) Then
                Return String.Empty
            End If

            Dim result As String = String.Empty

            Try
                ' 获取fieldName对应的列index
                ' 用colIndex去获取对应列的数据比fieldName效率高
                ' 有时候dgv.CurrentRow.Cells(fieldName)会找不到fieldName对应的列
                Dim colIndex = -1
                For i = 0 To dgv.Columns.Count - 1
                    If dgv.Columns(i).HeaderText.Equals(columnName, StringComparison.OrdinalIgnoreCase) Then
                        colIndex = dgv.Columns(i).Index
                        Exit For
                    End If
                Next

                If colIndex <> -1 Then
                    result = dgv.CurrentRow.Cells(colIndex).Value.ToString()
                Else
                    result = dgv.CurrentRow.Cells(columnName).Value.ToString()
                End If
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 获取当前选择行的前一行 某个列的值
        ''' 此函数只适用于  数据显示窗体删除操作处
        ''' 其他处请使用 返回string类型数值版本
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <param name="columnName"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function GetCurrentRowCellValue(Of T As DataGridView)(ByRef dgv As T, ByVal columnName As String, ByVal noUseParam As Integer) As Object
            If dgv.RowCount = 0 OrElse (dgv.RowCount = 1 AndAlso dgv.Item(columnName, 0).Value Is Nothing) Then
                Return Nothing
            End If

            Dim result As Object = Nothing

            Try
                ' 获取fieldName对应的列index
                ' 用colIndex去获取对应列的数据比fieldName效率高
                ' 有时候dgv.CurrentRow.Cells(fieldName)会找不到fieldName对应的列
                Dim colIndex = -1
                For i = 0 To dgv.Columns.Count - 1
                    If dgv.Columns(i).HeaderText.Equals(columnName, StringComparison.OrdinalIgnoreCase) Then
                        colIndex = dgv.Columns(i).Index
                        Exit For
                    End If
                Next

                If colIndex <> -1 Then
                    If dgv.CurrentRow.Index = 0 Then
                        result = dgv.Rows(0).Cells(colIndex).Value
                    Else
                        result = dgv.Rows(dgv.CurrentRow.Index - 1).Cells(colIndex).Value
                    End If
                Else
                    If dgv.CurrentRow.Index = 0 Then
                        result = dgv.Rows(0).Cells(columnName).Value
                    Else
                        result = dgv.Rows(dgv.CurrentRow.Index - 1).Cells(columnName).Value
                    End If
                End If
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 调整DataGridView控件 
        ''' 1.添加列标题 
        ''' 2.标题居中 
        ''' 4.不允许用户手动编辑 及新增删除行 只读
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <param name="showLastColumn">是否显示最后一列</param>
        ''' <param name="autoSizeColumnsMode">设置列宽自动调整模式</param>
        ''' <param name="appendColumns">要动态添加的列</param> 
        ''' <returns></returns>
        <Extension()>
        Public Function AdjustDgv(Of T As DataGridView)(ByRef dgv As T， ByVal showLastColumn As Boolean, ByVal autoSizeColumnsMode As DataGridViewAutoSizeColumnsMode, ByVal appendColumns As DataGridViewColumn(), Optional ByVal sortMode As DataGridViewColumnSortMode = DataGridViewColumnSortMode.NotSortable) As Boolean
            Try
                ' 如果已经调整过 就不需要再次调整
                If dgv.ReadOnly Then
                    Return True
                End If

                With dgv
                    .SuspendLayout()

                    If appendColumns?.Length > 0 Then
                        dgv.Columns.AddRange(appendColumns)
                    End If

                    For Each col As DataGridViewColumn In .Columns
                        col.SortMode = sortMode
                    Next col

                    ' 是否隐藏最后一列
                    dgv.Columns.Item(dgv.Columns.Count - 1).Visible = showLastColumn

                    ' 标题剧中对齐
                    .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                    .AutoSizeColumnsMode = autoSizeColumnsMode
                    ' 标题不换行
                    .ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False

                    .AllowDrop = False
                    .AllowUserToAddRows = False
                    .AllowUserToDeleteRows = False
                    .ReadOnly = True

                    .ResumeLayout()
                End With

                Return True
            Catch ex As Exception
                Logger.WriteLine(ex)

                Return False
            End Try
        End Function

        ''' <summary>
        ''' 调整DataGridView控件 
        ''' 1.添加列标题 
        ''' 2.标题居中 
        ''' 4.不允许用户手动编辑 及新增删除行 只读
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <param name="showLastColumn">是否显示最后一列</param>
        ''' <param name="autoSizeColumnsMode">设置列宽自动调整模式</param>
        ''' <param name="appendColumn">要动态添加的列</param> 
        ''' <returns></returns>
        <Extension()>
        Public Function AdjustDgv(Of T As DataGridView)(ByRef dgv As T， ByVal showLastColumn As Boolean, ByVal autoSizeColumnsMode As DataGridViewAutoSizeColumnsMode, ByVal appendColumn As DataGridViewColumn, Optional ByVal sortMode As DataGridViewColumnSortMode = DataGridViewColumnSortMode.NotSortable) As Boolean
            Dim newColumn(0) As DataGridViewColumn
            newColumn(0) = appendColumn
            Return AdjustDgv(dgv, showLastColumn, autoSizeColumnsMode, newColumn, sortMode)
        End Function

        ''' <summary>
        ''' 调整DataGridView控件 
        ''' 1.添加列标题 
        ''' 2.标题居中 
        ''' 4.不允许用户手动编辑 及新增删除行 只读
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <param name="showLastColumn">是否显示最后一列</param>
        ''' <param name="autoSizeColumnsMode">设置列宽自动调整模式</param>
        ''' <returns></returns>
        <Extension()>
        Public Function AdjustDgv(Of T As DataGridView)(ByRef dgv As T， ByVal showLastColumn As Boolean, ByVal autoSizeColumnsMode As DataGridViewAutoSizeColumnsMode) As Boolean
            Dim emptyColumn() As DataGridViewColumn
            Return AdjustDgv(dgv, showLastColumn, autoSizeColumnsMode, emptyColumn, DataGridViewColumnSortMode.NotSortable)
        End Function

        ''' <summary>
        ''' 调整DataGridView控件 
        ''' 1.添加列标题 
        ''' 2.标题居中 
        ''' 3.设置列宽自动调整模为 <see cref="DataGridViewAutoSizeColumnMode.DisplayedCells"/> 列宽调整到适合位于屏幕上当前显示的行中的列的所有单元格（包括标头单元格）的内容。
        ''' 4.不允许用户手动编辑 及新增删除行 只读
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <param name="showLastColumn">是否显示最后一列</param>
        ''' <returns></returns>
        <Extension()>
        Public Function AdjustDgv(Of T As DataGridView)(ByRef dgv As T， ByVal showLastColumn As Boolean) As Boolean
            Return AdjustDgv(dgv, showLastColumn, DataGridViewAutoSizeColumnsMode.DisplayedCells)
        End Function

        ''' <summary>
        ''' 调整DataGridView控件 
        ''' 1.添加列标题 
        ''' 2.标题居中 
        ''' 3.设置列宽自动调整模为 AllCells 列宽调整到适合位于屏幕上当前显示的行中的列的所有单元格（包括标头单元格）的内容。
        ''' 4.不允许用户手动编辑 及新增删除行 只读
        ''' </summary>
        ''' <param name="dgv"></param>
        <Extension()>
        Public Function AdjustDgv(Of T As DataGridView)(ByRef dgv As T) As Boolean
            Return AdjustDgv(dgv, True)
        End Function

        ''' <summary>
        ''' Flashes a window（Not control） until the window comes to the foreground
        ''' Receives the form that will flash
        ''' </summary>
        ''' <param name="form">the window to flash</param>
        ''' <returns>whether or not the window needed flashing</returns>
        <Extension()>
        Public Function FlashWindowEx(ByVal form As Form) As Boolean
            Return FlashWindowEx(form, FlashWindow.FLASHW_ALL Or FlashWindow.FLASHW_TIMERNOFG)
        End Function

        ''' <summary>
        ''' Flashes a window（Not control） until the window comes to the foreground
        ''' Receives the form that will flash.
        ''' FLASHWINFO.uCount default value is UInteger.MaxValue.
        ''' </summary>
        ''' <param name="form">the window to flash</param>
        ''' <param name="dwFlags">The flash status of the window</param>
        ''' <returns>whether or not the window needed flashing</returns>
        <Extension()>
        Public Function FlashWindowEx(ByVal form As Form, ByVal dwFlags As FlashWindow) As Boolean
            Return FlashWindowEx(form, dwFlags, Integer.MaxValue)
        End Function

        ''' <summary>
        ''' Flashes a window（Not control） until the window comes to the foreground
        ''' Receives the form that will flash
        ''' </summary>
        ''' <param name="form">the window to flash</param>
        ''' <param name="dwFlags">The flash status of the window</param>
        ''' <param name="nCount">闪烁窗口的次数。如果<paramref name="dwFlags"/>设置为<see cref="FlashWindow.FLASHW_STOP"/>的同时设置 <paramref name="nCount"/>为0，则窗口会恢复为初始状态，否则将继续保持橙色状态。</param>
        ''' <returns>whether or not the window needed flashing</returns>
        <Extension()>
        Public Function FlashWindowEx(ByVal form As Form, ByVal dwFlags As FlashWindow， ByVal nCount As Integer) As Boolean
            If form Is Nothing Then Return False
            If Not form.IsHandleCreated Then Return False
            Dim hwnd As IntPtr

            If form.InvokeRequired Then
                form.Invoke(Sub() hwnd = form.Handle)
            Else
                hwnd = form.Handle
            End If
            Return Windows2.FlashWindowEx(hwnd, dwFlags, nCount)
        End Function

        ''' <summary>
        ''' 是否点击了TreeNode的复选框
        ''' </summary>
        ''' <param name="e"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function IsClickedCheckbox(ByVal e As TreeNodeMouseClickEventArgs) As Boolean
            Return e.Location.X >= 22 AndAlso e.Location.X <= 34
        End Function

        ''' <summary>
        ''' 是否点击了TreeNode的加减按钮
        ''' </summary>
        ''' <param name="e"></param>
        ''' <returns></returns>
        <Extension()>
        Public Function IsClickedPlusMinus(ByVal e As TreeNodeMouseClickEventArgs) As Boolean
            Return e.Location.X >= 6 AndAlso e.Location.X <= 21
        End Function

        ''' <summary>
        ''' 参数 <paramref name="markColor"/> 的值为 True时，行号为 <paramref name="rowIndex"/>，列号为 <paramref name="colIndex"/> 的单元格背景设置成HotPink
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <param name="rowIndex"></param>
        ''' <param name="colIndex"></param>
        ''' <param name="markColor"></param>
        <Extension()>
        Public Sub SetCellsBackColor(ByVal dgv As DataGridView, ByVal rowIndex As Integer, ByVal colIndex As Integer, ByVal markColor As Boolean)
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
        <Extension()>
        Public Sub SetCellsBackColor(ByVal dgv As DataGridView, ByVal rowIndex As Integer, ByVal colIndex As Integer， ByVal compareValue As Integer)
            If (dgv.RowCount = 1 AndAlso dgv.Rows(0).Cells(1).Value Is Nothing) OrElse dgv.RowCount = 0 Then
                '
            Else
                Dim value = dgv.Rows(rowIndex).Cells(colIndex).Value
                If value IsNot DBNull.Value AndAlso
                    CInt(value) <= compareValue Then
                    dgv.Rows(rowIndex).Cells(colIndex).Style.BackColor = Color.HotPink
                    dgv.Rows(rowIndex).Cells(colIndex).Style.ForeColor = Color.White
                Else
                    dgv.Rows(rowIndex).Cells(colIndex).Style.BackColor = Color.White
                    dgv.Rows(rowIndex).Cells(colIndex).Style.ForeColor = Color.Black
                End If
            End If
        End Sub

        ''' <summary>
        ''' 当<paramref name="compare"/> 为条件成立时，列号为 <paramref name="colIndex"/> 的所有行的单元格背景设置成HotPink
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <param name="colIndex"></param>
        ''' <param name="compare"></param>
        <Extension()>
        Public Sub SetCellsBackColor(Of T)(ByVal dgv As DataGridView, ByVal colIndex As Integer， ByVal compare As Func(Of Object, T, Boolean), ByVal compareValue As T)
            If (dgv.RowCount = 1 AndAlso dgv.Rows(0).Cells(1).Value Is Nothing) OrElse dgv.RowCount = 0 Then
                '
            Else
                Dim rowIndex = 0
                While rowIndex < dgv.RowCount
                    Dim value = dgv.Rows(rowIndex).Cells(colIndex).Value

                    If value IsNot DBNull.Value AndAlso compare(value, compareValue) Then
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
        ''' 所有行，列号为 <paramref name="colIndex"/> 的单元格值等于-1时，单元格背景设置成HotPink
        ''' </summary>
        ''' <param name="dgv"></param>
        ''' <param name="colIndex"></param>
        <Extension()>
        Public Sub SetCellsBackColor(ByVal dgv As DataGridView, ByVal colIndex As Integer)
            If (dgv.RowCount = 1 AndAlso dgv.Rows(0).Cells(1).Value Is Nothing) OrElse dgv.RowCount = 0 Then
                '
            Else
                Dim rowIndex = 0
                While rowIndex < dgv.RowCount
                    Dim value = dgv.Rows(rowIndex).Cells(colIndex).Value
                    If value IsNot DBNull.Value AndAlso
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
        ''' 所有行，列号为 <paramref name="colIndex"/> 的单元格值等于 <paramref name="compareValue"/> 时，单元格背景设置成HotPink
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="dgv"></param>
        ''' <param name="colIndex"></param>
        ''' <param name="compareValue"></param>
        <Extension()>
        Public Sub SetCellsBackColor(Of T)(ByVal dgv As DataGridView, ByVal colIndex As Integer， ByVal compareValue As T)
            If (dgv.RowCount = 1 AndAlso dgv.Rows(0).Cells(1).Value Is Nothing) OrElse dgv.RowCount = 0 Then
                '
            Else
                Dim rowIndex = 0
                While rowIndex < dgv.RowCount
                    Dim value = dgv.Rows(rowIndex).Cells(colIndex).Value
                    If value IsNot DBNull.Value AndAlso
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

#End Region

    End Module


End Namespace
