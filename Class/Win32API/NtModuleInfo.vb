Imports System.Runtime.InteropServices

Namespace ShanXingTech.Win32API
    <StructLayout(LayoutKind.Sequential)>
    Public Structure NtModuleInfo
        Public BaseOfDll As IntPtr
        Public SizeOfImage As Integer
        Public EntryPoint As IntPtr
    End Structure
End Namespace