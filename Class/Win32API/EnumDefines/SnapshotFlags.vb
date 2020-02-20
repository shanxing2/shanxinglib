Namespace ShanXingTech.Win32API
    <Flags()>
    Public Enum SnapshotFlags As Integer
        HeapList = &H1
        Process = &H2
        Thread = &H4
        [Module] = &H8
        Module32 = &H10
        Inherit = &H80000000
        All = &HF
        NoHeaps = &H40000000
    End Enum
End Namespace
