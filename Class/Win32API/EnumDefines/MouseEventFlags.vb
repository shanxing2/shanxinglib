
Namespace ShanXingTech.Win32API
    <Flags()>
    Public Enum MouseEventFlags As Integer
        MOUSEEVENTF_MOVE = &H1
        MOUSEEVENTF_LEFTDOWN = &H2
        MOUSEEVENTF_LEFTUP = &H4
        MOUSEEVENTF_RIGHTDOWN = &H8
        MOUSEEVENTF_RIGHTUP = &H10
        MOUSEEVENTF_MIDDLEDOWN = &H20
        MOUSEEVENTF_MIDDLEUP = &H40
        MOUSEEVENTF_XDOWN = &H80
        MOUSEEVENTF_XUP = &H100
        MOUSEEVENTF_WHEEL = &H800
        MOUSEEVENTF_HWHEEL = &H1000
        ''' <summary>
        ''' dX和dY参数含有规范化的绝对坐标。如果不设置，这些参数含有相对数据：相对于上次位置的改动位置。此标志可设置，也可不设置，不管鼠标的类型或与系统相连的类似于鼠标的设备的类型如何。要得到关于相对鼠标动作的信息，参见<see cref="Mouse_Event(MouseEventFlags, Integer, Integer, Integer, Integer)"/>参数部分。
        ''' </summary>
        MOUSEEVENTF_ABSOLUTE = &H8000
        WM_MOUSEMOVE = &H200
        WM_LBUTTONDOWN = &H201
        WM_LBUTTONUP = &H202
        WM_LBUTTONDBLCLK = &H203
        WM_RBUTTONDOWN = &H204
        WM_RBUTTONUP = &H205
        WM_RBUTTONDBLCLK = &H206
        WM_MBUTTONDOWN = &H207
        WM_MBUTTONUP = &H208
        WM_MBUTTONDBLCLK = &H209
        WM_MOUSEWHEEL = &H20A
    End Enum
End Namespace
