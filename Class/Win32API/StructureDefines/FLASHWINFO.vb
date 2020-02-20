Imports System.Runtime.InteropServices

Namespace ShanXingTech.Win32API
	<StructLayout(LayoutKind.Sequential)>
	Public Structure FLASHWINFO
#Disable Warning bc40025
		''' <summary>
		''' 该结构的字节大小
		''' </summary>
		Public cbSize As UInteger
		''' <summary>
		''' 需要闪烁的窗口的句柄，该窗口可以是打开的或最小化的
		''' </summary>
		Public hwnd As IntPtr
		''' <summary>
		''' 闪烁的状态,可以是下面取值之一或组合<see cref="FlashWindow"/>
		''' </summary>
		Public dwFlags As UInteger
		''' <summary>
		''' 闪烁窗口的次数
		''' </summary>
		Public uCount As UInteger
		''' <summary>
		''' 窗口闪烁的频度，毫秒为单位；若该值为0，则为默认图标的闪烁频度
		''' </summary>
		Public dwTimeout As UInteger
#Enable Warning bc40025
	End Structure
End Namespace