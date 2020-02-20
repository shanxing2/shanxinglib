Imports System.Runtime.InteropServices

Namespace ShanXingTech.Win32API
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Public Structure MODULEENTRY32
        Public dwSize As Integer
        Public th32ModuleID As Integer
        Public th32ProcessID As Integer
        Public GlblcntUsage As Integer
        Public ProccntUsage As Integer
        Public modBaseAddr As IntPtr
        Public modBaseSize As Integer
        Public hModule As IntPtr
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)>
        Public szModule As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
        Public szExePath As String
    End Structure
End Namespace