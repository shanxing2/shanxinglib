
Imports System.Runtime.InteropServices

Namespace ShanXingTech.Win32API
    <StructLayout(LayoutKind.Sequential)>
    Public Structure ComBoBoxInfo
        Public cbSize As Integer
        Public rcItem As Rect
        Public rcButton As Rect
        Public stateButton As ComboBoxButtonState
        Public hwndCombo As Integer
        Public hwndItem As IntPtr
        Public hwndList As IntPtr

		Public Overrides Function Equals(obj As Object) As Boolean
			Dim comBoBoxInfoObj = CType(obj, ComBoBoxInfo)

			Return hwndCombo = comBoBoxInfoObj.hwndCombo
		End Function

		Public Overrides Function GetHashCode() As Integer
			Throw New NotImplementedException()
		End Function

		Public Shared Operator =(left As ComBoBoxInfo, right As ComBoBoxInfo) As Boolean
			Return left.Equals(right)
		End Operator

		Public Shared Operator <>(left As ComBoBoxInfo, right As ComBoBoxInfo) As Boolean
			Return Not left = right
		End Operator
	End Structure

	Public Enum ComboBoxButtonState
        STATE_SYSTEM_NONE = 0
        STATE_SYSTEM_INVISIBLE = &H8000
        STATE_SYSTEM_PRESSED = &H8
    End Enum
End Namespace