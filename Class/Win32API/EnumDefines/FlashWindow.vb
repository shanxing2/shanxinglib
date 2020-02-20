Namespace ShanXingTech.Win32API
#Disable Warning bc40032
	''' <summary>
	''' The flash status of the window
	''' </summary>
	Public Enum FlashWindow As UInteger
		''' <summary>
		''' 停止闪烁，系统将重置窗口到其初始状态,如果<see cref="FLASHWINFO.uCount"/>不同时设置为0，橙色状态会继续保持。
		''' </summary>    
		FLASHW_STOP = 0

		''' <summary>
		''' 闪烁窗口的标题。
		''' </summary>
		FLASHW_CAPTION = 1

		''' <summary>
		''' 闪烁窗口的任务栏按钮。
		''' </summary>
		FLASHW_TRAY = 2

		''' <summary>
		''' 同时闪烁窗口标题和窗口的任务栏按钮。
		''' This is equivalent to setting the FLASHW_CAPTION | FLASHW_TRAY flags. 
		''' </summary>
		FLASHW_ALL = 3

		''' <summary>
		''' 不停地闪烁，直到FLASHW_STOP标志被设置。
		''' </summary>
		FLASHW_TIMER = 4

		''' <summary>
		''' 不停地闪烁，直到窗口前端显示。
		''' </summary>
		FLASHW_TIMERNOFG = 12
	End Enum
#Enable Warning bc40032
End Namespace
