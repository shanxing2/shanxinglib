Imports System.Diagnostics.CodeAnalysis
Imports System.Runtime.InteropServices
Imports System.Runtime.Versioning
Imports System.Security
Imports System.Security.Permissions
Imports System.Text
Imports Microsoft.Win32.SafeHandles

Namespace ShanXingTech.Win32API
    ''' <summary>
    ''' 这个人很懒，什么都没写
    ''' </summary>
    <SuppressUnmanagedCodeSecurity()>
    <HostProtection(SecurityAction.LinkDemand, MayLeakOnAbort:=True)>
    Partial Public Module UnsafeNativeMethods
        'Public NotInheritable Class UnsafeNativeMethods
        ' 明文：神即道, 道法自然, 如来|闪星网络信息科技 ShanXingTech Q2287190283
        ' 算法：古典密码中的有密钥换位密码 密钥：ShanXingTech
        Public Const ShanXingTechQ2287190283 = "神闪X7,SQB道信T2道网N9来A2D如H2C然技HA即星I1|N8E法息E8,络G0自科C3"

#Region "函数声明区"
		''' <summary>
		''' 判断文件夹是否为空
		''' </summary>
		''' <param name="pszPath"></param>
		''' <returns>Returns TRUE if pszPath is an empty directory. Returns FALSE if pszPath is not a directory, or if it contains at least one file other than "."c or "..".</returns>
		<DllImport(ExternDll.Shlwapi, SetLastError:=True, CharSet:=CharSet.Unicode)>
		Public Function PathIsDirectoryEmptyA(ByVal pszPath As String) As Boolean
        End Function

		''' <summary>
		''' 成功返回<paramref name="hWndChild"/>的句柄
		''' 失败返回0
		''' </summary>
		''' <param name="hWndChild"></param>
		''' <param name="hWndNewParent"></param>
		''' <returns></returns>
		<DllImport(ExternDll.User32, SetLastError:=True, CharSet:=CharSet.Auto)>
		Public Function SetParent(ByVal hWndChild As IntPtr, ByVal hWndNewParent As IntPtr) As IntPtr
		End Function

		''' <summary>
		''' 获取传入路径的段路径形式
		''' </summary>
		''' <param name="longPath">原始文件长路径</param>
		''' <param name="ShortPath">储存文件短路径的缓存器</param>
		''' <param name="bufferSize"> 参数 <paramref name="ShortPath"/> 长度的2倍，即 <paramref name="ShortPath"/>.Length*2</param>
		''' <returns></returns>
		<DllImport(ExternDll.Kernel32, SetLastError:=True, CharSet:=CharSet.Unicode)>
		Public Function GetShortPathName(ByVal longPath As String,
				 <MarshalAs(UnmanagedType.LPTStr)> ByVal ShortPath As System.Text.StringBuilder,
				 <MarshalAs(Runtime.InteropServices.UnmanagedType.U4)> ByVal bufferSize As Integer) As Integer
		End Function

		''' <summary>
		''' 获取cookies
		''' </summary>
		''' <param name="pchURL"></param>
		''' <param name="pchCookieName"></param>
		''' <param name="pchCookieData"></param>
		''' <param name="pcchCookieData"></param>
		''' <param name="dwFlags"></param>
		''' <param name="lpReserved"></param>
		''' <returns></returns>
		<DllImport(ExternDll.Wininet, CharSet:=CharSet.Unicode, SetLastError:=True)>
		Public Function InternetGetCookieEx(pchURL As String, pchCookieName As String, pchCookieData As StringBuilder, ByRef pcchCookieData As Integer, dwFlags As Integer, lpReserved As IntPtr) As Boolean
		End Function

		''' <summary>
		''' 设置Cookies
		''' </summary>
		''' <param name="lpszUrlName"></param>
		''' <param name="lbszCookieName"></param>
		''' <param name="lpszCookieData"></param>
		''' <returns></returns>
		<DllImport(ExternDll.Wininet, CharSet:=CharSet.Auto, SetLastError:=True)>
		Public Function InternetSetCookie(lpszUrlName As String, lbszCookieName As String, lpszCookieData As String) As Boolean
		End Function

		<DllImport(ExternDll.Urlmon, CharSet:=CharSet.Ansi)>
		Public Function UrlMkSetSessionOption(dwOption As Integer, pBuffer As String, dwBufferLength As Integer, dwReserved As Integer) As Integer
		End Function

		''' <summary>
		''' 设置浏览器行为
		''' </summary>
		''' <param name="FeatureEntry"></param>
		''' <param name="dwFlags"></param>
		''' <param name="fEnable"></param>
		''' <returns></returns>
		<DllImport(ExternDll.Urlmon, CharSet:=CharSet.Auto, SetLastError:=True)>
		Public Function CoInternetSetFeatureEnabled(ByVal FeatureEntry As INTERNETFEATURELIST,
			   ByVal dwFlags As Integer,
			   ByVal fEnable As Boolean) As Integer
		End Function

		''' <summary>
		''' 获取本地系统的网络连接状态。
		''' </summary>
		''' <param name="dwFlag">指向一个变量，该变量接收连接描述内容。该参数在函数返回FALSE时仍可以返回一个有效的标记。该参数可以为下列值的一个或多个。</param>
		''' <para>值					含义
		''' INTERNET_CONNECTION_CONFIGURED	0x40	Local system has a valid connection to the Internet, but it might Or might Not be currently connected.
		''' INTERNET_CONNECTION_LAN 		0x02	Local system uses a local area network to connect to the Internet.
		''' INTERNET_CONNECTION_MODEM		0x01	Local system uses a modem to connect to the Internet.
		''' INTERNET_CONNECTION_MODEM_BUSY	0x08	No longer used.
		''' INTERNET_CONNECTION_OFFLINE 	0x20	Local system Is in offline mode.
		''' INTERNET_CONNECTION_PROXY		0x0		Local system uses a proxy server to connect to the Internet.
		''' INTERNET_RAS_INSTALLED			0x10	Local system has RAS installed.
		''' </para>
		''' <param name="dwReserved">保留值,必须为0。</param>
		''' <returns>当存在一个modem或一个LAN连接时，返回TRUE，当不存在internet连接或所有的连接当前未被激活时，返回false。
		''' <para>当该函数返回false时，程序可以调用GetLastError来接收错误代码</para></returns>
		''' <remarks>该函数如果返回TRUE，表明至少有一个连接是有效的。它并不能保证这个有效的连接是连向一个指定的主机。程序应该经常检查利用API连接到服务器的返回错误代码，用以判断连接状态。使用InternetCheckConnection函数可以判断一个连接到指定主机的连接是否建立。
		''' <para>返回值为TRUE也表明一个modem连接处于激活状态或一个LAN连接处于激活状态。而FALSE代表modem和LAN均不处于连接状态。如果返回FALSE，INTERNET_CONNECTION_CONFIGURED 标识将被设置，以表明自动拨号被设置为“总是拨号”，但当前不处于激活状态。如果自动拨号未被设置，函数返回FALSE。</para></remarks>
		<DllImport(ExternDll.Wininet, CharSet:=CharSet.Auto, SetLastError:=True)>
		<Obsolete("此函数获取状态有延时，而且不准确，不应该再调用此函数判断网络状态")>
		Public Function InternetGetConnectedState(ByRef dwFlag As Integer, ByVal dwReserved As Integer) As Boolean
		End Function

		''' <summary>
		''' 该函数将指定的消息发送到一个或多个窗口。此函数为指定的窗口调用窗口程序，直到窗口程序处理完消息再返回。而和函数 <see cref="PostMessage(IntPtr, Integer, IntPtr, String)"/> 不同，<see cref="PostMessage(IntPtr, Integer, IntPtr, String)"/> 是将一个消息寄送到一个线程的消息队列后就立即返回。
		''' </summary>
		''' <param name="hWnd">其窗口程序将接收消息的窗口的句柄。如果此参数为HWND_BROADCAST，则消息将被发送到系统中所有顶层窗口，包括无效或不可见的非自身拥有的窗口、被覆盖的窗口和弹出式窗口，但消息不被发送到子窗口。</param>
		''' <param name="msg">指定被发送的消息。</param>
		''' <param name="wParam">指定附加的消息特定信息。</param>
		''' <param name="lParam">指定附加的消息特定信息。</param>
		''' <returns>返回值指定消息处理的结果，依赖于所发送的消息。</returns>
		''' <remarks>需要用HWND_BROADCAST通信的应用程序应当使用函数RegisterWindowMessage来为应用程序间的通信取得一个唯一的消息。
		''' 如果指定的窗口是由正在调用的线程创建的，则窗口程序立即作为子程序调用。如果指定的窗口是由不同线程创建的，则系统切换到该线程并调用恰当 窗口程序。线程间的消息只有在线程执行消息检索代码时才被处理。发送线程被阻塞直到接收线程处理完消息为止。</remarks> 
		<DllImport(ExternDll.User32, CharSet:=CharSet.Unicode)>
		Public Function SendMessage(hWnd As IntPtr, ByVal msg As Integer, wParam As Integer, <MarshalAs(UnmanagedType.LPWStr)> lParam As String) As Integer
		End Function
		''' <summary>
		''' 该函数将指定的消息发送到一个或多个窗口。此函数为指定的窗口调用窗口程序，直到窗口程序处理完消息再返回。而和函数 <see cref="PostMessage(IntPtr, Integer, Integer,ByRef CopyDataStruct)"/> 不同，<see cref="PostMessage(IntPtr, Integer, Integer, ByRef CopyDataStruct)"/> 是将一个消息寄送到一个线程的消息队列后就立即返回。
		''' </summary>
		''' <param name="hWnd">其窗口程序将接收消息的窗口的句柄。如果此参数为HWND_BROADCAST，则消息将被发送到系统中所有顶层窗口，包括无效或不可见的非自身拥有的窗口、被覆盖的窗口和弹出式窗口，但消息不被发送到子窗口。</param>
		''' <param name="msg">指定被发送的消息。</param>
		''' <param name="wParam">指定附加的消息特定信息。</param>
		''' <param name="lParam">指定附加的消息特定信息。</param>
		''' <returns>返回值指定消息处理的结果，依赖于所发送的消息。</returns>
		''' <remarks>需要用HWND_BROADCAST通信的应用程序应当使用函数RegisterWindowMessage来为应用程序间的通信取得一个唯一的消息。
		''' 如果指定的窗口是由正在调用的线程创建的，则窗口程序立即作为子程序调用。如果指定的窗口是由不同线程创建的，则系统切换到该线程并调用恰当 窗口程序。线程间的消息只有在线程执行消息检索代码时才被处理。发送线程被阻塞直到接收线程处理完消息为止。</remarks> 
		<DllImport(ExternDll.User32, EntryPoint:="SendMessage")>
		Public Function SendMessage(hWnd As IntPtr, ByVal msg As Integer, wParam As Integer, ByRef lParam As CopyDataStruct) As Integer
		End Function
		''' <summary>
		''' 该函数将指定的消息发送到一个或多个窗口。此函数为指定的窗口调用窗口程序，直到窗口程序处理完消息再返回。而和函数 <see cref="PostMessage(IntPtr, Integer, Integer, Integer)"/> 不同，<see cref="PostMessage(IntPtr, Integer, Integer, Integer)"/> 是将一个消息寄送到一个线程的消息队列后就立即返回。
		''' </summary>
		''' <param name="hWnd">其窗口程序将接收消息的窗口的句柄。如果此参数为HWND_BROADCAST，则消息将被发送到系统中所有顶层窗口，包括无效或不可见的非自身拥有的窗口、被覆盖的窗口和弹出式窗口，但消息不被发送到子窗口。</param>
		''' <param name="msg">指定被发送的消息。</param>
		''' <param name="wParam">指定附加的消息特定信息。</param>
		''' <param name="lParam">指定附加的消息特定信息。</param>
		''' <returns>返回值指定消息处理的结果，依赖于所发送的消息。</returns>
		''' <remarks>需要用HWND_BROADCAST通信的应用程序应当使用函数RegisterWindowMessage来为应用程序间的通信取得一个唯一的消息。
		''' 如果指定的窗口是由正在调用的线程创建的，则窗口程序立即作为子程序调用。如果指定的窗口是由不同线程创建的，则系统切换到该线程并调用恰当 窗口程序。线程间的消息只有在线程执行消息检索代码时才被处理。发送线程被阻塞直到接收线程处理完消息为止。</remarks> 
		<DllImport(ExternDll.User32, EntryPoint:="SendMessage")>
		Public Function SendMessage(hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
		End Function
		''' <summary>
		''' 该函数将指定的消息发送到一个或多个窗口。此函数为指定的窗口调用窗口程序，直到窗口程序处理完消息再返回。而和函数 <see cref="PostMessage(IntPtr, Integer, Integer, Integer)"/> 不同，<see cref="PostMessage(IntPtr, Integer, Integer, Integer)"/> 是将一个消息寄送到一个线程的消息队列后就立即返回。
		''' </summary>
		''' <param name="hWnd">其窗口程序将接收消息的窗口的句柄。如果此参数为HWND_BROADCAST，则消息将被发送到系统中所有顶层窗口，包括无效或不可见的非自身拥有的窗口、被覆盖的窗口和弹出式窗口，但消息不被发送到子窗口。</param>
		''' <param name="msg">指定被发送的消息。</param>
		''' <param name="wParam">指定附加的消息特定信息。</param>
		''' <param name="lParam">指定附加的消息特定信息。</param>
		''' <returns>返回值指定消息处理的结果，依赖于所发送的消息。</returns>
		''' <remarks>需要用HWND_BROADCAST通信的应用程序应当使用函数RegisterWindowMessage来为应用程序间的通信取得一个唯一的消息。
		''' 如果指定的窗口是由正在调用的线程创建的，则窗口程序立即作为子程序调用。如果指定的窗口是由不同线程创建的，则系统切换到该线程并调用恰当 窗口程序。线程间的消息只有在线程执行消息检索代码时才被处理。发送线程被阻塞直到接收线程处理完消息为止。</remarks> 
		<DllImport(ExternDll.User32, EntryPoint:="SendMessage")>
		Public Function SendMessage(hWnd As IntPtr, ByVal msg As Integer, wParam As Integer, ByVal lParam As IntPtr) As Integer
		End Function

		<DllImport(ExternDll.User32, SetLastError:=True, CharSet:=CharSet.Auto)>
		Public Function PostMessage(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> lParam As String) As Boolean
		End Function

		<DllImport(ExternDll.User32, SetLastError:=True, CharSet:=CharSet.Auto)>
		Public Function PostMessage(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByRef lParam As CopyDataStruct) As Boolean
		End Function

		<DllImport(ExternDll.User32, SetLastError:=True, CharSet:=CharSet.Auto)>
		Public Function PostMessage(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As Integer, ByRef lParam As CopyDataStruct) As Boolean
		End Function

		<DllImport(ExternDll.User32, SetLastError:=True, CharSet:=CharSet.Auto)>
		Public Function PostMessage(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Boolean
		End Function

		<DllImport(ExternDll.User32, CharSet:=CharSet.Auto)>
		Public Function ReleaseCapture() As Boolean
		End Function

		''' <summary>
		''' 多媒体播放
		''' </summary>
		''' <param name="lpszSoundName"></param>
		''' <param name="uFlags"></param>
		''' <returns></returns>
		<DllImport(ExternDll.WinMM, EntryPoint:="sndPlaySound", CharSet:=CharSet.Unicode, SetLastError:=True)>
		Public Function SndPlaySound(ByVal lpszSoundName As String, ByVal uFlags As Integer) As Integer
		End Function

		''' <summary>
		''' 用来播放多媒体文件的API指令，可以播放MPEG,AVI,WAV,MP3,等等。命令串中不区分大小写
		''' </summary>
		''' <param name="command"></param>
		''' <param name="buffer"></param>
		''' <param name="bufferSize"></param>
		''' <param name="hwndCallback"></param>
		''' <returns></returns>
		<DllImport(ExternDll.WinMM, EntryPoint:="mciSendString", CharSet:=CharSet.Unicode, SetLastError:=True)>
		Public Function MciSendString(ByVal command As String, ByVal buffer As StringBuilder, ByVal bufferSize As Integer, ByVal hwndCallback As IntPtr) As Integer
		End Function

		''' <summary>
		''' 返回系统开启算起所经过的时间,单位毫秒。
		''' </summary>
		''' <returns></returns>
		<DllImport(ExternDll.WinMM, EntryPoint:="timeGetTime")>
		Public Function TimeGetTime() As Integer
		End Function

		''' <summary>
		''' 当前操作挂起指定时间
		''' </summary>
		''' <param name="dwMilliseconds"></param>
		<DllImport(ExternDll.Kernel32, EntryPoint:="Sleep")>
		Public Sub Sleep(ByVal dwMilliseconds As Integer)
		End Sub

		''' <summary>
		''' 设置浏览器相关选项
		''' </summary>
		''' <param name="hInternet"></param>
		''' <param name="dwOption"></param>
		''' <param name="lpBuffer"></param>
		''' <param name="dwBufferLength"></param>
		''' <returns></returns>
		<DllImport(ExternDll.Wininet, SetLastError:=True, CharSet:=CharSet.Auto)>
		Public Function InternetSetOption(hInternet As IntPtr, dwOption As Integer, ByRef lpBuffer As IntPtr, dwBufferLength As Integer) As Boolean
		End Function

		''' <summary>
		''' 查找窗口句柄
		''' </summary>
		''' <param name="lpClassName">窗体类名</param>
		''' <param name="windowTitle">窗体标题</param>
		''' <returns></returns>
		<DllImport(ExternDll.User32, SetLastError:=True, CharSet:=CharSet.Auto)>
		Public Function FindWindow(ByVal lpClassName As String,
										 ByVal windowTitle As String) As IntPtr
		End Function

		''' <summary>
		''' 在窗口列表中寻找与指定条件相符的第一个子窗口
		''' </summary>
		''' <param name="parentHandle">要查找的子窗口所在的父窗口的句柄（如果设置了 <paramref name="parentHandle"/>，则表示从这个 <paramref name="parentHandle"/>指向的父窗口中搜索子窗口）。</param>
		''' <param name="childAfter">子窗体句柄。查找从在Z序中的下一个子窗口开始。子窗口必须为 <paramref name="parentHandle"/> 窗口的直接子窗口而非后代窗口。如果<paramref name="childAfter"/> 为0，查找从 <paramref name="parentHandle"/> 的第一个子窗口开始。如果  <paramref name="parentHandle"/> 和 <paramref name="childAfter"/> 同时为0，则函数查找所有的顶层窗口及消息窗口。</param>
		''' <param name="lpClassName">指向一个指定了类名的空结束字符串，或一个标识类名字符串的成员的指针。如果该参数为一个成员，则它必须为前次调用theGlobaIAddAtom函数产生的全局成员。该成员为16位，必须位于<paramref name="lpClassName"/> 的低16位，高位必须为0。</param>
		''' <param name="windowTitle">指向一个指定了窗口名（窗口标题）的空结束字符串。如果该参数为 Nothing，则为所有窗口全匹配。</param>
		''' <returns>成功返回句柄，失败返回0</returns>
		<DllImport(ExternDll.User32, SetLastError:=True, CharSet:=CharSet.Auto)>
		Public Function FindWindowEx(ByVal parentHandle As IntPtr,
										 ByVal childAfter As IntPtr,
										 ByVal lpClassName As String,
										 ByVal windowTitle As String) As IntPtr
		End Function

		''' <summary>
		''' 返回父窗口句柄
		''' </summary>
		''' <param name="childHWnd"></param>
		''' <returns></returns>
		<DllImport(ExternDll.User32, ExactSpelling:=True, CharSet:=CharSet.Auto)>
		Public Function GetParent(ByVal childHWnd As IntPtr) As IntPtr
		End Function

		''' <summary>
		''' 将创建指定窗口的线程设置到前台，并且激活该窗口。键盘输入转向该窗口，并为用户改各种可视的记号。系统给创建前台窗口的线程分配的权限稍高于其他线程。
		''' </summary>
		''' <param name="hWnd"></param>
		''' <returns></returns>
		<DllImport(ExternDll.User32, ExactSpelling:=True, CharSet:=CharSet.Auto)>
		Public Function SetForegroundWindow(ByVal hWnd As IntPtr) As Integer
		End Function

		''' <summary>
		''' 异步激活窗口
		''' </summary>
		''' <param name="hWnd"></param>
		''' <param name="nCmdShow"></param>
		''' <returns></returns>
		<DllImport(ExternDll.User32, SetLastError:=True)>
		Public Function ShowWindowAsync(hWnd As IntPtr, <MarshalAs(UnmanagedType.I4)> nCmdShow As WindowCommand) As <MarshalAs(UnmanagedType.Bool)> Boolean
		End Function

		''' <summary>
		''' 激活窗口
		''' </summary>
		''' <param name="hWnd"></param>
		''' <returns></returns>
		<DllImport(ExternDll.User32, SetLastError:=True)>
		Public Function SetActiveWindow(ByVal hWnd As IntPtr) As IntPtr
		End Function

		''' <summary>
		''' 设置窗口位置
		''' </summary>
		''' <param name="hWnd"></param>
		''' <param name="hWndInsertAfter"></param>
		''' <param name="X"></param>
		''' <param name="Y"></param>
		''' <param name="cx"></param>
		''' <param name="cy"></param>
		''' <param name="uFlags"></param>
		''' <returns></returns>
		<DllImport(ExternDll.User32, SetLastError:=True)>
		Public Function SetWindowPos(ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As WindowPositions) As Boolean
		End Function

		''' <summary>
		''' 设置窗口位置
		''' </summary>
		''' <param name="hWnd"></param>
		''' <param name="hWndInsertAfter"></param>
		''' <param name="X"></param>
		''' <param name="Y"></param>
		''' <param name="cx"></param>
		''' <param name="cy"></param>
		''' <param name="uFlags"></param>
		''' <returns></returns>
		<DllImport(ExternDll.User32, SetLastError:=True)>
		Public Function SetWindowPos(ByVal hWnd As IntPtr, ByVal hWndInsertAfter As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As WindowPositions) As Boolean
		End Function

		''' <summary>
		''' 该函数用来改变指定窗口的属性。
		''' </summary>
		''' <param name="hWndChild"></param>
		''' <param name="nIndex"></param>
		''' <param name="hWndParent"></param>
		''' <returns></returns>
		Public Function SetWindowLong(hWndChild As IntPtr, <MarshalAs(UnmanagedType.I4)> nIndex As WindowLong, hWndParent As IntPtr) As IntPtr
			Return If(IntPtr.Size = 4,
				IntPtr.op_Explicit(SetWindowLongPtr32(hWndChild, nIndex, hWndParent)),
				SetWindowLongPtr64(hWndChild, nIndex, hWndParent))
		End Function

		Public Declare Auto Function SetWindowLongPtr32 Lib "user32.dll" Alias "SetWindowLong" (hWndChild As IntPtr, <MarshalAs(UnmanagedType.I4)> nIndex As WindowLong, hWndParent As IntPtr) As Integer

		Public Declare Auto Function SetWindowLongPtr64 Lib "user32.dll" Alias "SetWindowLongPtr" (hWndChild As IntPtr, <MarshalAs(UnmanagedType.I4)> nIndex As WindowLong, hWndParent As IntPtr) As IntPtr

		''' <summary>
		''' 窗口是否可见
		''' </summary>
		''' <param name="hWnd"></param>
		''' <returns></returns>
		<DllImport(ExternDll.User32, SetLastError:=True)>
		Public Function IsWindowVisible(ByVal hWnd As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
		End Function

		''' <summary>
		''' 该函数确定给定的窗口句柄是否识别一个已存在的窗口。
		''' </summary>
		''' <param name="hWnd">被测试窗口的句柄。</param>
		''' <returns>如果窗口句柄标识了一个已存在的窗口，返回值为非零；如果窗口句柄未标识一个已存在窗口，返回值为零。</returns>
		<DllImport(ExternDll.User32, SetLastError:=True)>
		Public Function IsWindow(ByVal hWnd As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
		End Function

		''' <summary>
		''' 把光标移到屏幕的指定位置。如果新位置不在由 ClipCursor函数设置的屏幕矩形区域之内，则系统自动调整坐标，使得光标在矩形之内。（此函数为前台操作，后台移动鼠标可用<see cref="SendMessage(IntPtr, Integer, Integer,  Integer)"/>或<see cref="PostMessage(IntPtr, Integer, IntPtr, String)"/>）
		''' </summary>
		''' <param name="x">指定光标的新的X坐标，以屏幕坐标表示。</param>
		''' <param name="y">指定光标的新的Y坐标，以屏幕坐标表示。</param>
		''' <returns></returns>
		<DllImport(ExternDll.User32, CharSet:=CharSet.Auto)>
		Public Function SetCursorPos(ByVal x As Integer, ByVal y As Integer) As Boolean
		End Function

		''' <summary>
		''' 鼠标移动和按钮点击。如果鼠标被移动，用设置MOUSEEVENTF_MOVE来表明，dX和dy保留移动的信息
		''' </summary>
		''' <param name="dwFlags">标志位集，指定点击按钮和鼠标动作的多种情况,可以是 <see cref="MouseEventFlags"/> 中的任意组合</param>
		''' <param name="dx">指定鼠标沿x轴的绝对位置或者从上次鼠标事件产生以来移动的数量，依赖于MOUSEEVENTF_ABSOLUTE的设置。给出的绝对数据作为鼠标的实际X坐标；给出的相对数据作为移动的mickeys数。一个mickey表示鼠标移动的数量，表明鼠标已经移动。</param>
		''' <param name="dy">指定鼠标沿y轴的绝对位置或者从上次鼠标事件产生以来移动的数量，依赖于MOUSEEVENTF_ABSOLUTE的设置。给出的绝对数据作为鼠标的实际y坐标，给出的相对数据作为移动的mickeys数。</param>
		''' <param name="dwData">如果dwFlags为MOUSEEVENTF_WHEEL，则dwData指定鼠标轮移动的数量。正值表明鼠标轮向前转动，即远离用户的方向；负值表明鼠标轮向后转动，即朝向用户。一个轮击定义为WHEEL_DELTA，即120。如果dwFlagsS不是MOUSEEVENTF_WHEEL，则dWData应为零。</param>
		''' <param name="dwExtraInfo">指定与鼠标事件相关的附加32位值。应用程序调用函数GetMessageExtraInfo来获得此附加信息。</param>
		<DllImport(ExternDll.User32, EntryPoint:="mouse_event"， CharSet:=CharSet.Auto)>
		Public Sub Mouse_Event(dwFlags As MouseEventFlags, dx As Integer, dy As Integer, dwData As Integer, dwExtraInfo As Integer)
		End Sub

		''' <summary>
		''' 将你打开的APP中客户区的坐标点信息转换为整个屏幕的坐标，其中：所有的坐标（无论是屏幕坐标还是客户区坐标）其坐标原点都是左上角为（0，0）。
		''' 其中：屏幕坐标是指你的显示器的左上角（0， 0）开始的两条坐标轴，而客户区坐标是指你的应用程序打开后除了标题栏、工具栏、菜单栏后的剩下区域，在这个区域中，左上角为坐标的原点（0，0），以上两个坐标都是从左到右为正、从上到下为正,一般用来在鼠标右键的编程中
		''' </summary>
		''' <param name="hWnd">用户区域用于转换的窗口句柄。</param>
		''' <param name="lpPoint">指向一个含有要转换的用户坐标的结构的指针，如果函数调用成功，新屏幕坐标复制到此结构。</param>
		''' <returns>如果函数调用成功，返回值为非零值，否则为零。</returns>
		''' <remarks>函数用屏幕坐标取代POINT结构中的用户坐标，屏幕坐标与屏幕左上角相关联。</remarks>
		<DllImport(ExternDll.User32, CharSet:=CharSet.Auto)>
		Public Function ClientToScreen(ByVal hWnd As IntPtr, ByRef lpPoint As Drawing.Point) As Boolean
		End Function

		''' <summary>
		''' a combobox is made up of three controls, a button, a list and textbox; 
		''' not support 64 bit
		''' </summary>
		''' <param name="hWnd"></param>
		''' <param name="pcbi"></param>
		''' <returns></returns>
		<DllImport(ExternDll.User32, CharSet:=CharSet.Auto, SetLastError:=True)>
		Public Function GetComboBoxInfo(hWnd As IntPtr, ByRef pcbi As ComBoBoxInfo) As Boolean
		End Function

		''' <summary>
		''' 注册全局热键
		''' </summary>
		''' <param name="hwnd">接收热键产生WM_HOTKEY消息的窗口句柄。若该参数NULL，传递给调用线程的WM_HOTKEY消息必须在消息循环中进行处理。</param>
		''' <param name="hotkeyId">定义热键的标识符。调用线程中的其他热键，不能使用同样的标识符。应用程序必须定义一个0X0000-0xBFFF范围的值。一个共享的动态链接库（DLL）必须定义一个范围为0xC000-0xFFFF的值(GlobalAddAtom函数返回该范围）。为了避免与其他动态链接库定义的热键冲突，一个DLL必须使用GlobalAddAtom函数获得热键的标识符。</param>
		''' <param name="fsModifiers">定义为了产生WM_HOTKEY消息而必须与由nVirtKey参数定义的键一起按下的键。</param>
		''' <param name="vk">定义热键的虚拟键码。</param>
		''' <returns>若函数调用成功，返回一个非0值。若函数调用失败，则返回值为0。若要获得更多的错误信息，可以调用GetLastError函数。</returns>
		<DllImport(ExternDll.User32, CharSet:=CharSet.Auto)>
		Public Function RegisterHotKey(ByVal hwnd As IntPtr, ByVal hotkeyId As Integer, ByVal fsModifiers As FsModifiers, ByVal vk As Windows.Forms.Keys) As Boolean
		End Function

		''' <summary>
		''' 注销全局热键
		''' </summary>
		''' <param name="hwnd"></param>
		''' <param name="hotkeyId"></param>
		''' <returns></returns>
		<DllImport(ExternDll.User32, EntryPoint:="UnregisterHotKey", CharSet:=CharSet.Auto)>
		Public Function UnRegisterHotKey(ByVal hwnd As IntPtr, ByVal hotkeyId As Integer) As Boolean
		End Function

		''' <summary>
		''' 向全局原子表添加一个字符串，并返回这个字符串的唯一标识符（原子ATOM）。注：全局原子不会在应用程序终止时自动删除。每次调用<see cref="GlobalAddAtom(String)"/>>函数，必须相应的调用<see cref="GlobalDeleteAtom(Integer)"/>>函数删除原子。
		''' </summary>
		''' <param name="lpString"></param>
		''' <returns></returns>
		<DllImport(ExternDll.Kernel32, CallingConvention:=CallingConvention.StdCall, CharSet:=CharSet.Unicode)>
		Public Function GlobalAddAtom(lpString As String) As Integer
		End Function

		<DllImport(ExternDll.Kernel32, CallingConvention:=CallingConvention.StdCall, CharSet:=CharSet.Unicode)>
		Public Function GlobalDeleteAtom(atom As Integer) As Integer
		End Function

		''' <summary>
		''' 确定调用进程是否由用户模式的调试器调试。
		''' </summary>
		''' <returns>如果当前进程运行在调试器的上下文，返回值为非零值。否则，返回值是零。</returns>
		<DllImport(ExternDll.Kernel32)>
		Public Function IsDebuggerPresent() As Boolean
		End Function

		''' <summary>
		'''     Copies the text of the specified window's title bar (if it has one) into a buffer. If the specified window is a
		'''     control, the text of the control is copied. However, GetWindowText cannot retrieve the text of a control in another
		'''     application.
		'''     <para>
		'''         Go to https://msdn.microsoft.com/en-us/library/windows/desktop/ms633520%28v=vs.85%29.aspx  for more
		'''         information
		'''     </para>
		''' </summary>
		''' <param name="hWnd">
		'''     C++ ( hWnd [in]. Type: HWND )<br />A <see cref="IntPtr" /> handle to the window or control containing the text.
		''' </param>
		''' <param name="lpString">
		'''     C++ ( lpString [out]. Type: LPTSTR )<br />The <see cref="StringBuilder" /> buffer that will receive the text. If
		'''     the string is as long or longer than the buffer, the string is truncated and terminated with a null character.
		''' </param>
		''' <param name="nMaxSize">
		'''     C++ ( nMaxCount [in]. Type: int )<br /> Should be equivalent to
		'''     <see cref="StringBuilder.Length" /> after call returns. The <see cref="int" /> maximum number of characters to copy
		'''     to the buffer, including the null character. If the text exceeds this limit, it is truncated.
		''' </param>
		''' <returns>
		'''     If the function succeeds, the return value is the length, in characters, of the copied string, not including
		'''     the terminating null character. If the window has no title bar or text, if the title bar is empty, or if the window
		'''     or control handle is invalid, the return value is zero. To get extended error information, call GetLastError.<br />
		'''     This function cannot retrieve the text of an edit control in another application.
		''' </returns>
		''' <remarks>
		'''     If the target window is owned by the current process, GetWindowText causes a WM_GETTEXT message to be sent to the
		'''     specified window or control. If the target window is owned by another process and has a caption, GetWindowText
		'''     retrieves the window caption text. If the window does not have a caption, the return value is a null string. This
		'''     behavior is by design. It allows applications to call GetWindowText without becoming unresponsive if the process
		'''     that owns the target window is not responding. However, if the target window is not responding and it belongs to
		'''     the calling application, GetWindowText will cause the calling application to become unresponsive. To retrieve the
		'''     text of a control in another process, send a WM_GETTEXT message directly instead of calling GetWindowText.<br />For
		'''     an example go to
		'''     <see cref="!:https://msdn.microsoft.com/en-us/library/windows/desktop/ms644928%28v=vs.85%29.aspx#sending">
		'''         Sending a
		'''         Message.
		'''     </see>
		''' </remarks>
		<DllImport(ExternDll.User32, CharSet:=CharSet.Auto, SetLastError:=True)>
		Public Function GetWindowText(hWnd As IntPtr, lpString As StringBuilder, nMaxSize As Integer) As Integer
		End Function

		''' <summary>
		''' <see cref="EnumWindows(ByVal EnumWindowsProc, IntPtr)"/> 专用的回调函数
		''' </summary>
		''' <param name="hWnd">找到的窗口句柄</param>
		''' <param name="lParam"><see cref="EnumWindows(ByVal EnumWindowsProc, IntPtr)"/> 传给的参数; 因为它是指针, 可传入, 但一般用作传出数据</param>
		''' <returns>函数返回 False 时, 调用它的 <see cref="EnumWindows(ByVal EnumWindowsProc, IntPtr)"/> 将停止遍历并返回 False</returns>
		Public Delegate Function EnumWindowsProc(ByVal hWnd As IntPtr, ByRef lParam As IntPtr) As Boolean
		''' <summary>
		''' 枚举所有顶层窗口
		''' </summary>
		''' <param name="lpEnumFunc">回调函数委托</param>
		''' <param name="lParam">给回调函数的参数, 它对应回调函数的第二个参数</param>
		''' <returns>成功与否, 其实是返回了回调函数的返回值</returns>
		<DllImport(ExternDll.User32, SetLastError:=True, CharSet:=CharSet.Auto)>
		Public Function EnumWindows(
			ByVal lpEnumFunc As EnumWindowsProc,
			ByVal lParam As IntPtr) As Boolean
		End Function

		''' <summary>
		''' 该函数用于判断指定的窗口是否允许接受键盘或鼠标输入。
		''' </summary>
		''' <param name="hWnd"></param>
		''' <returns>若窗口允许接受键盘或鼠标输入，则返回非0值，若窗口不允许接受键盘或鼠标输入，则返回值为0。</returns>
		''' <remarks>子窗口只有在被允许并且可见时才可接受输入。 </remarks>  
		<DllImport(ExternDll.User32, SetLastError:=True)>
		Public Function IsWindowEnabled(ByVal hWnd As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
		End Function

		''' <summary>
		''' Retrieves a handle to a window that has the specified relationship (Z-Order or owner) to the specified window.
		''' </summary>
		''' <remarks>The EnumChildWindows function is more reliable than calling GetWindow in a loop. An application that
		''' calls GetWindow to perform this task risks being caught in an infinite loop or referencing a handle to a window
		''' that has been destroyed.</remarks>
		''' <param name="hWnd">A handle to a window. The window handle retrieved is relative to this window, based on the
		''' value of the uCmd parameter.</param>
		''' <param name="uCmd">The relationship between the specified window and the window whose handle is to be
		''' retrieved.</param>
		''' <returns>
		''' If the function succeeds, the return value is a window handle. If no window exists with the specified relationship
		''' to the specified window, the return value is NULL. To get extended error information, call GetLastError.
		''' </returns>
		<DllImport(ExternDll.User32, SetLastError:=True)>
		Public Function GetWindow(hWnd As IntPtr, uCmd As WindowType) As IntPtr
		End Function

		<DllImport(ExternDll.User32, CharSet:=CharSet.Auto)>
		Public Function GetClassName(ByVal hWnd As IntPtr, ByVal lpClassName As StringBuilder, ByVal nMaxCount As Integer) As Integer
			' Leave function empty    
		End Function


		''' <summary>
		''' 获取包含指定模块的文件的完全限定路径。
		''' </summary>
		''' <param name="hProcess">进程的句柄</param>
		''' <param name="hModule">模块的句柄。可以是一个DLL模块，或者是一个应用程序的实例句柄。如果该参数为NULL，该函数返回该应用程序全路径。</param>
		''' <param name="lpFileName">指定一个字串缓冲区，要在其中容纳文件的用NULL字符中止的路径名，<paramref name="hModule"/> 模块就是从这个文件装载进来的</param>
		''' <param name="nMaxSize">装载到缓冲区 <paramref name="lpFileName"/> 的最大字符数量</param>
		''' <returns>如执行成功，返回复制到<paramref name="lpFileName"/>的实际字符数量；
		''' 零表示失败。使用GetLastError可以打印错误信息。</returns>
		<DllImport(ExternDll.Psapi, SetLastError:=True)>
		Public Function GetModuleFileNameEx(ByVal hProcess As IntPtr, ByVal hModule As IntPtr, <Out()> ByVal lpFileName As StringBuilder, <[In]()> <MarshalAs(UnmanagedType.U4)> ByVal nMaxSize As Integer) As Integer
		End Function

		''' <summary>
		''' The return value is the identifier of the thread that created the window.
		''' 返回线程号，注意，<paramref name="lpdwProcessId"/> 是存放进程号的变量。返回值是线程号，<paramref name="lpdwProcessId"/> 是进程号存放处。
		''' </summary>
		''' <param name="hwnd">[in] （向函数提供的）被查找窗口的句柄.</param>
		''' <param name="lpdwProcessId"><paramref name="lpdwProcessId"/>[out] 进程号的存放地址（变量地址） Pointer to a variable that receives the process identifier. If this parameter is not NULL, <see cref="GetWindowThreadProcessId"/> copies the identifier of the process to the variable; otherwise, it does not. （如果参数不为NULL，即提供了存放处--变量，那么本函数把进程标志拷贝到存放处，否则不动作。）</param>
		''' <returns></returns>
		<DllImport(ExternDll.User32, SetLastError:=True)>
		Public Function GetWindowThreadProcessId(ByVal hwnd As IntPtr,
													 ByRef lpdwProcessId As Integer) As Integer
		End Function

		<DllImport(ExternDll.Kernel32, SetLastError:=True)>
		Public Function CreateToolhelp32Snapshot(ByVal dwFlags As SnapshotFlags, ByVal th32ProcessID As Integer) As IntPtr
		End Function

		<DllImport(ExternDll.Kernel32, SetLastError:=True)>
		Public Function Module32First(ByVal hSnapshot As IntPtr, ByVal lpme As MODULEENTRY32) As Boolean
		End Function

		<DllImport(ExternDll.Kernel32, SetLastError:=True)>
		Public Function Module32First(ByVal hSnapshot As HandleRef, ByVal lpme As IntPtr) As Boolean
		End Function

		<DllImport(ExternDll.Kernel32, CharSet:=CharSet.Auto, SetLastError:=True)>
		<ResourceExposure(ResourceScope.None)>
		Public Function Module32Next(handle As HandleRef, entry As IntPtr) As Boolean
		End Function

		<DllImport(ExternDll.Kernel32, ExactSpelling:=True, CharSet:=CharSet.Auto, SetLastError:=True)>
		Public Function CloseHandle(handle As IntPtr) As Boolean
        End Function

        <DllImport(ExternDll.Kernel32, CharSet:=CharSet.Auto, SetLastError:=True)>
        <ResourceExposure(ResourceScope.None)>
        Public Function OpenProcess(access As Integer, inherit As Boolean, processId As Integer) As SafeProcessHandle
        End Function

        <DllImport(ExternDll.Psapi, CharSet:=CharSet.Auto, SetLastError:=True)>
        <ResourceExposure(ResourceScope.None)>
        Public Function EnumProcessModules(handle As SafeProcessHandle, modules As IntPtr, size As Integer, ByRef needed As Integer) As Boolean
        End Function

        <DllImport(ExternDll.Psapi, CharSet:=CharSet.Auto, SetLastError:=True)>
        <ResourceExposure(ResourceScope.Machine)>
        Public Function EnumProcesses(processIds As Integer(), size As Integer, ByRef needed As Integer) As Boolean
        End Function

        <SuppressMessage("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")>
        <DllImport(ExternDll.Psapi, CharSet:=CharSet.Auto, SetLastError:=True, BestFitMapping:=False)>
        <ResourceExposure(ResourceScope.Machine)>
        Public Function GetModuleFileNameEx(processHandle As HandleRef, moduleHandle As HandleRef, baseName As StringBuilder, size As Integer) As Integer
        End Function

        <DllImport(ExternDll.Psapi, CharSet:=CharSet.Auto, SetLastError:=True)>
        <ResourceExposure(ResourceScope.Process)>
        Public Function GetModuleInformation(processHandle As SafeProcessHandle, moduleHandle As HandleRef, ntModuleInfo As NtModuleInfo, size As Integer) As Boolean
        End Function

        <DllImport(ExternDll.Psapi, CharSet:=CharSet.Auto, SetLastError:=True, BestFitMapping:=False)>
        <ResourceExposure(ResourceScope.Machine)>
        Public Function GetModuleBaseName(processHandle As SafeProcessHandle, moduleHandle As HandleRef, baseName As StringBuilder, size As Integer) As Integer
        End Function

        <DllImport(ExternDll.Kernel32, SetLastError:=True)>
        <ResourceExposure(ResourceScope.None)>
        Public Function IsWow64Process(hProcess As SafeProcessHandle, ByRef wow64Process As Boolean) As Boolean
        End Function

        <DllImport(ExternDll.Kernel32, CharSet:=CharSet.Auto)>
        <ResourceExposure(ResourceScope.Process)>
        Public Function GetCurrentProcessId() As Integer
        End Function

        <DllImport(ExternDll.Kernel32, CharSet:=CharSet.Ansi, SetLastError:=True)>
        <ResourceExposure(ResourceScope.Process)>
        Public Function GetCurrentProcess() As IntPtr
        End Function

		''' <summary>
		''' 从注册表中检索本地计算机的NetBIOS名称
		''' </summary>
		''' <param name="lpBuffer">缓冲器</param>
		''' <param name="nMaxSize">缓冲区大小，应该传入 <see cref="StringBuilder.Capacity"/> * 2 </param>
		''' <returns></returns>
		<DllImport(ExternDll.Kernel32, CharSet:=CharSet.Auto, BestFitMapping:=False)>
		<ResourceExposure(ResourceScope.None)>
		Public Function GetComputerName(lpBuffer As StringBuilder, nMaxSize As Integer) As Boolean
		End Function

		''' <summary>
		''' 闪烁指定的窗口。它不会更改窗口的激活状态。
		''' </summary>
		''' <param name="pwfi">指向 FLASHWINFO 结构的指针。</param>
		''' <returns>返回调用 FlashWindowEx 函数之前指定窗口状态。如果调用之前窗口标题是活动的，返回值为非零值。</returns>
		<DllImport(ExternDll.User32)>
		Public Function FlashWindowEx(ByRef pwfi As FLASHWINFO) As <MarshalAs(UnmanagedType.Bool)> Boolean
		End Function
#End Region
	End Module

End Namespace