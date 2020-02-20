Namespace ShanXingTech.Win32API
    <Flags>
    Public Enum WindowPositions As Integer
        ''' <summary>If the calling thread and the thread that owns the window are attached to different input queues, 
        ''' the system posts the request to the thread that owns the window. This prevents the calling thread from 
        ''' blocking its execution while other threads process the request.</summary>
        ''' <remarks>SWP_ASYNCWINDOWPOS</remarks>
        SynchronousWindowPosition = &H4000
        ''' <summary>Prevents generation of the WM_SYNCPAINT message.</summary>
        ''' <remarks>SWP_DEFERERASE</remarks>
        DeferErase = &H2000
        ''' <summary>Draws a frame (defined in the window's class description) around the window.</summary>
        ''' <remarks>SWP_DRAWFRAME</remarks>
        DrawFrame = &H20
        ''' <summary>Applies new frame styles set using the SetWindowLong function. Sends a WM_NCCALCSIZE message to 
        ''' the window, even if the window's size is not being changed. If this flag is not specified, WM_NCCALCSIZE 
        ''' is sent only when the window's size is being changed.</summary>
        ''' <remarks>SWP_FRAMECHANGED</remarks>
        FrameChanged = &H20
        ''' <summary>Hides the window.</summary>
        ''' <remarks>SWP_HIDEWINDOW</remarks>
        HideWindow = &H80
        ''' <summary>Does not activate the window. If this flag is not set, the window is activated and moved to the 
        ''' top of either the topmost or non-topmost group (depending on the setting of the hWndInsertAfter 
        ''' parameter).</summary>
        ''' <remarks>SWP_NOACTIVATE</remarks>
        DoNotActivate = &H10
        ''' <summary>Discards the entire contents of the client area. If this flag is not specified, the valid 
        ''' contents of the client area are saved and copied back into the client area after the window is sized or 
        ''' repositioned.</summary>
        ''' <remarks>SWP_NOCOPYBITS</remarks>
        DoNotCopyBits = &H100
        ''' <summary>Retains the current position (ignores X and Y parameters).</summary>
        ''' <remarks>SWP_NOMOVE</remarks>
        IgnoreMove = &H2
        ''' <summary>Does not change the owner window's position in the Z order.</summary>
        ''' <remarks>SWP_NOOWNERZORDER</remarks>
        DoNotChangeOwnerZOrder = &H200
        ''' <summary>Does not redraw changes. If this flag is set, no repainting of any kind occurs. This applies to 
        ''' the client area, the nonclient area (including the title bar and scroll bars), and any part of the parent 
        ''' window uncovered as a result of the window being moved. When this flag is set, the application must 
        ''' explicitly invalidate or redraw any parts of the window and parent window that need redrawing.</summary>
        ''' <remarks>SWP_NOREDRAW</remarks>
        DoNotRedraw = &H8
        ''' <summary>Same as the SWP_NOOWNERZORDER flag.</summary>
        ''' <remarks>SWP_NOREPOSITION</remarks>
        DoNotReposition = &H200
        ''' <summary>Prevents the window from receiving the WM_WINDOWPOSCHANGING message.</summary>
        ''' <remarks>SWP_NOSENDCHANGING</remarks>
        DoNotSendChangingEvent = &H400
        ''' <summary>Retains the current size (ignores the cx and cy parameters).</summary>
        ''' <remarks>SWP_NOSIZE</remarks>
        IgnoreResize = &H1
        ''' <summary>Retains the current Z order (ignores the hWndInsertAfter parameter).</summary>
        ''' <remarks>SWP_NOZORDER</remarks>
        IgnoreZOrder = &H4
        ''' <summary>Displays the window.</summary>
        ''' <remarks>SWP_SHOWWINDOW</remarks>
        ShowWindow = &H40
    End Enum
End Namespace


