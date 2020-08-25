Imports System.IO
Imports System.Runtime.InteropServices

Namespace ShanXingTech.Cryptography
    ''' <summary>
    ''' 一个简单的注册类
    ''' </summary>
    Public NotInheritable Class RegistSimple
        ' 注册信息保存路径
        Private Shared s_RegistInfoFolder As String
        Private Shared s_RegistInfoFileName As String

        Public Shared Function GetRegistInfo(ByVal appName As String) As (MachineCode As String, RegistrationCode As String)
            Dim machineCode As String = String.Empty
            Dim registrationCode As String = String.Empty

            ' 初始化注册信息保存路径

            ' 先获取主板UUID
            Dim uuid = Windows2.CmdRun("wmic csproduct get UUID")
            ' 注，有些电脑会返回 ’03000200-0400-0500-0006-000700080009‘，这个可能是重复的，
            ' 没有uuid的全部是F

            ' 再获取COM GUID
            Dim guidId As String
            ' 可能有些设备没法获取到uuid，所以使用guidId作为uuid
            If uuid.IsNullOrEmpty Then
                Dim asm = My.Application.GetType.Assembly
                Dim guidType = GetType(GuidAttribute)
                For Each guidAttr As GuidAttribute In asm.GetCustomAttributes(guidType, False)
                    ' 文件名是"程序集的COM GUID-MachineName"
                    guidId = guidAttr.Value
                    Exit For
                Next
            Else
                guidId = uuid.Replace("UUID", String.Empty).Trim
            End If

            ' uuid后再添加 计算机名
            guidId &= "_" & Environment.MachineName.GetHashCode

            s_RegistInfoFolder = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)
            If Not "\"c = s_RegistInfoFolder.Chars(s_RegistInfoFolder.Length - 1) Then
                s_RegistInfoFolder += "\"
            End If
            ' 为了兼容旧版本，需要添加以下代码，把之前的注册文件移动到新位置
            Dim oldRegistInfoFileName = $"{s_RegistInfoFolder}{guidId}-{appName}.dat"

            s_RegistInfoFolder += My.Application.Info.CompanyName & "\"
            s_RegistInfoFileName = $"{s_RegistInfoFolder}{guidId}-{appName}.dat"

            ' 如果本地没有相关注册信息 那就生成一个注册码保存到本地
            If Not Directory.Exists(s_RegistInfoFolder) Then
                ' 创建保存目录
                IO2.Directory.Create(s_RegistInfoFolder)
            End If

            ' 如果有旧的注册信息，则移动到新文件夹内
            If File.Exists(oldRegistInfoFileName) Then
                Directory.Move(oldRegistInfoFileName, s_RegistInfoFileName)
            End If

            If File.Exists(s_RegistInfoFileName) Then
                ' 如果有 则读取
                Dim fileContext = My.Computer.FileSystem.ReadAllText(s_RegistInfoFileName, System.Text.Encoding.UTF8)
                If Not fileContext.IndexOf("=") > -1 Then
                    Return (machineCode, registrationCode)
                End If

                machineCode = fileContext.Substring(0, fileContext.IndexOf("="))
                registrationCode = fileContext.Substring(fileContext.IndexOf("=") + 1)

                Return (machineCode, registrationCode)
            Else
                ' 没有注册信息则生成全球唯一注册码并保存
                machineCode = If(guidId.IsNullOrEmpty, Guid.NewGuid.ToString(), guidId)

                WriteRegistInfoToFile(machineCode, registrationCode)

                Return (machineCode, registrationCode)
            End If
        End Function

        ''' <summary>
        ''' 保存注册信息到本地
        ''' </summary>
        ''' <param name="machineCode"></param>
        ''' <param name="registrationCode"></param>
        Public Shared Sub WriteRegistInfoToFile(ByVal machineCode As String, ByVal registrationCode As String)
            ' 写文件
            Try
                Using stream As New StreamWriter(s_RegistInfoFileName, False, System.Text.Encoding.UTF8)
                    stream.Write($"{machineCode}={registrationCode}")
                End Using
                Debug.Print(Logger.MakeDebugString("注册信息保存完成"))
            Catch ex As Exception
                Logger.WriteLine(ex)
            End Try
        End Sub
    End Class
End Namespace


