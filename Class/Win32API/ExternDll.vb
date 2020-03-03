'------------------------------------------------------------------------------
' <copyright file="ExternDll.cs" company="Microsoft">
'     Copyright (c) Microsoft Corporation.  All rights reserved.
' </copyright>
'------------------------------------------------------------------------------

Namespace ShanXingTech
    Friend NotInheritable Class ExternDll
        Private Sub New()
        End Sub

#If FEATURE_PAL AndAlso Not SILVERLIGHT Then

		#If Not PLATFORM_UNIX Then
		Friend Const DLLPREFIX As String = ""
		Friend Const DLLSUFFIX As String = ".dll"
		#Else
		#If __APPLE__ Then
		Friend Const DLLPREFIX As String = "lib"
		Friend Const DLLSUFFIX As String = ".dylib"
		#Elif _AIX Then
		Friend Const DLLPREFIX As String = "lib"
		Friend Const DLLSUFFIX As String = ".a"
		#Elif __hppa__ OrElse IA64 Then
		Friend Const DLLPREFIX As String = "lib"
		Friend Const DLLSUFFIX As String = ".sl"
		#Else
		Friend Const DLLPREFIX As String = "lib"
		Friend Const DLLSUFFIX As String = ".so"
		#End If
		#End If

		Public Const Kernel32 As String = DLLPREFIX + "rotor_pal" + DLLSUFFIX
		Public Const User32 As String = DLLPREFIX + "rotor_pal" + DLLSUFFIX
		Public Const Mscoree As String = DLLPREFIX + "sscoree" + DLLSUFFIX

		#Elif FEATURE_PAL AndAlso SILVERLIGHT Then

		Public Const Kernel32 As String = "coreclr"
		Public Const User32 As String = "coreclr"


#Else
        Public Const Activeds As String = "activeds.dll"
        Public Const Advapi32 As String = "advapi32.dll"
        Public Const Comctl32 As String = "comctl32.dll"
        Public Const Comdlg32 As String = "comdlg32.dll"
        Public Const Gdi32 As String = "gdi32.dll"
        Public Const Gdiplus As String = "gdiplus.dll"
        Public Const Hhctrl As String = "hhctrl.ocx"
        Public Const Imm32 As String = "imm32.dll"
        Public Const Kernel32 As String = "kernel32.dll"
        Public Const Loadperf As String = "Loadperf.dll"
        Public Const Mscoree As String = "mscoree.dll"
        Public Const Clr As String = "clr.dll"
        Public Const Msi As String = "msi.dll"
        Public Const Mqrt As String = "mqrt.dll"
        Public Const Ntdll As String = "ntdll.dll"
        Public Const Ole32 As String = "ole32.dll"
        Public Const Oleacc As String = "oleacc.dll"
        Public Const Oleaut32 As String = "oleaut32.dll"
        Public Const Olepro32 As String = "olepro32.dll"
        Public Const PerfCounter As String = "perfcounter.dll"
        Public Const Powrprof As String = "Powrprof.dll"
        Public Const Psapi As String = "psapi.dll"
        Public Const Shell32 As String = "shell32.dll"
        Public Const User32 As String = "user32.dll"
        Public Const Uxtheme As String = "uxtheme.dll"
        Public Const WinMM As String = "winmm.dll"
        Public Const Winspool As String = "winspool.drv"
        Public Const Wtsapi32 As String = "wtsapi32.dll"
        Public Const Version As String = "version.dll"
        Public Const Vsassert As String = "vsassert.dll"
        Public Const Fxassert As String = "Fxassert.dll"
        Public Const Shlwapi As String = "shlwapi.dll"
        Public Const Crypt32 As String = "crypt32.dll"
        Public Const ShCore As String = "SHCore.dll"
        Public Const Wldp As String = "wldp.dll"
        Public Const Wininet As String = "wininet.dll"
        Public Const Urlmon As String = "urlmon.dll"

        ' system.data specific
        Friend Const Odbc32 As String = "odbc32.dll"
        Friend Const SNI As String = "System.Data.dll"

        ' system.data.oracleclient specific
        Friend Const OciDll As String = "oci.dll"
        Friend Const OraMtsDll As String = "oramts.dll"
#End If
    End Class
End Namespace