'//==++==
'// 
'//   Copyright (c) Microsoft Corporation.  All rights reserved.
'// 
'//==--==
'/*============================================================
'**
'** Class:  StringBuilderCache
'**
'** Purpose: provide a cached reusable instance Of stringbuilder
'**          per thread  it's an optimisation that reduces the 
'**          number of instances constructed And collected.
'**
'**  Acquire - Is used to get a string builder to use of a 
'**            particular size.  It can be called any number of 
'**            times, if a stringbuilder Is in the cache then
'**            it will be returned And the cache emptied.
'**            subsequent calls will return a New stringbuilder.
'**
'**            A StringBuilder instance Is cached in 
'**            Thread Local Storage And so there Is one per thread
'**
'**  Release - Place the specified builder in the cache if it Is 
'**            Not too big.
'**            The stringbuilder should Not be used after it has 
'**            been released.
'**            Unbalanced Releases are perfectly acceptable.  It
'**            will merely cause the runtime to create a New 
'**            stringbuilder next time Acquire Is called.
'**
'**  GetStringAndRelease
'**          - ToString() the stringbuilder, Release it to the 
'**            cache And return the resulting string
'**
'===========================================================*/
'https://referencesource.microsoft.com/#mscorlib/system/text/stringbuildercache.cs,a6dbe82674916ac0

Namespace ShanXingTech.Text2
    ''' <summary>
    ''' 缓存StringBuilder
    ''' <para>使用缓存技术可以避免New StringBuilder（非第一次）时的内存申请消耗</para>
    ''' <para>以及ToString时产生的字符串副本没有被及时回收的情况，</para>
    ''' <para>如果StringBuilder的文本内容（即ToString之后的字符串）小于等于360字节，</para>
    ''' <para>适合使用此缓存技术</para> 
    ''' <para>特别素需要多次使用StringBuilder时效果更加明显。</para>
    ''' <para>如果大于等于361字节，可以用带有Super后缀的版本</para> 
    ''' <para>此代码来源于ILSpy 下的System.Text的StringBuilderCache，</para>
    ''' <para>另外，VS IDE现在的版本Roslyn的产品经理Bill Chiles 在谈及Roslyn编辑器优化的时候，</para>
    ''' <para>亦有说到此种缓存技术。</para>
    ''' <para>文章Performance Facts and .NET Framework Tips</para>
    ''' <para>英文原文链接：http://download-codeplex.sec.s-msft.com/Download?ProjectName=roslyn&amp;DownloadId=838017Essential </para>
    ''' <para>译文链接：http://www.cnblogs.com/yangecnu/p/Essential-DotNet-Framework-Performance-Truths-and-Tips.html </para>
    ''' <para>另外在MSDN “编写大型的响应式 .NET Framework 应用” 一文中也可以看到</para>
    ''' <para>注：本类非线程安全</para>
    ''' </summary>
    Public Class StringBuilderCache
        ' 明文：神即道, 道法自然, 如来|闪星网络信息科技 ShanXingTech Q2287190283
        ' 算法：古典密码中的有密钥换位密码 密钥：ShanXingTech
        Public Const ShanXingTechQ2287190283 = "神闪X7,SQB道信T2道网N9来A2D如H2C然技HA即星I1|N8E法息E8,络G0自科C3"

        ''' <summary>
        ''' 缓存大小360字节
        ''' </summary>
        Public Const MAX_BUILDER_SIZE As Integer = 360
        ''' <summary>
        ''' .net中大对象的定义为85000B=83K
        ''' 大对象每次都回收内存
        ''' </summary>
        Public Const MAX_BUILDER_SIZE_SUPER As Integer = 1024 * 80
        ''' <summary>
        ''' 线程独立，每个线程都有一个唯一的实例
        ''' </summary>
        <ThreadStatic()>
        Private Shared s_CachedStringBuilder As Text.StringBuilder
        ''' <summary>
        ''' 线程独立，每个线程都有一个唯一的实例
        ''' </summary>
        <ThreadStatic()>
        Private Shared s_CachedStringBuilderSuper As Text.StringBuilder

        ''' <summary>
        ''' 获取StringBuilder对象的缓存实例。默认缓存最大大小为 <see cref="MAX_BUILDER_SIZE"/>
        ''' </summary>
        ''' <param name="capacity"></param>
        ''' <returns></returns>
        Public Shared Function Acquire(ByVal capacity As Integer) As Text.StringBuilder
            If capacity <= MAX_BUILDER_SIZE Then
                Dim sb = s_CachedStringBuilder

                If sb Is Nothing Then
                    Return New Text.StringBuilder(capacity)
                End If

                sb.Clear()
                s_CachedStringBuilder = Nothing

                ' 已经有了缓存（使用过一次ToString方法之后）
                ' 返回缓存实例
                If capacity <= sb.Capacity Then
                    Return sb
                End If
            End If

            Return New Text.StringBuilder(capacity)
        End Function

        ''' <summary>
        ''' 获取StringBuilder对象的缓存实例。默认实例大小为16
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function Acquire() As Text.StringBuilder
            Return Acquire(16)
        End Function

        ''' <summary>
        ''' 获取StringBuilder对象的缓存实例。 默认缓存最大大小为 <see cref="MAX_BUILDER_SIZE_SUPER"/>
        ''' </summary>
        ''' <param name="capacity"></param>
        ''' <returns></returns>
        Public Shared Function AcquireSuper(ByVal capacity As Integer) As Text.StringBuilder
            If capacity <= MAX_BUILDER_SIZE_SUPER Then
                Dim sb = s_CachedStringBuilderSuper

                If sb Is Nothing Then
                    Return New Text.StringBuilder(capacity)
                End If

                sb.Clear()
                s_CachedStringBuilderSuper = Nothing

                ' 已经有了缓存（使用过一次ToString方法之后）
                ' 返回缓存实例
                If capacity <= sb.Capacity Then
                    Return sb
                End If
            End If

            Return New Text.StringBuilder(capacity)
        End Function

        ''' <summary>
        ''' 获取StringBuilder对象的缓存实例。默认实例大小为 <see cref="MAX_BUILDER_SIZE"/> +1
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function AcquireSuper() As Text.StringBuilder
            Return AcquireSuper(MAX_BUILDER_SIZE + 1)
        End Function

        ''' <summary>
        ''' 释放内存
        ''' 如果大小小于等于<paramref name="maxBuilderSize"/>>则缓存下来以备下次使用
        ''' </summary>
        ''' <param name="sb"></param>
        ''' <param name="maxBuilderSize"></param>
        Private Shared Sub Release(ByRef sb As Text.StringBuilder, ByRef cacheStringBuilder As Text.StringBuilder, ByVal maxBuilderSize As Integer)
            If sb.Capacity <= maxBuilderSize Then
                cacheStringBuilder = sb
            End If
            sb.Clear()
        End Sub

        ''' <summary>
        ''' 获取实例字符串内容并且释放内存
        ''' </summary>
        ''' <param name="sb"></param>
        ''' <returns></returns>
        Public Shared Function GetStringAndReleaseBuilder(ByRef sb As Text.StringBuilder) As String
            Dim funcRst As String = sb.ToString()
            Release(sb, s_CachedStringBuilder, MAX_BUILDER_SIZE)

            Return funcRst
        End Function

        ''' <summary>
        ''' 获取实例字符串内容并且释放内存
        ''' </summary>
        ''' <param name="sb"></param>
        ''' <returns></returns>
        Public Shared Function GetStringAndReleaseBuilderSuper(ByRef sb As Text.StringBuilder) As String
            Dim funcRst As String = sb.ToString()
            Release(sb, s_CachedStringBuilderSuper, MAX_BUILDER_SIZE_SUPER)

            Return funcRst
        End Function

        Private Shared Sub ReleaseStringBuilderCache(ByRef sb As Text.StringBuilder, ByRef cacheStringBuilder As Text.StringBuilder)
            sb.Length = 0

            If cacheStringBuilder Is Nothing Then Return

            cacheStringBuilder.Length = 0
        End Sub

        Public Shared Sub ReleaseStringBuilderCache(ByRef sb As Text.StringBuilder)
            ReleaseStringBuilderCache(sb, s_CachedStringBuilder)
        End Sub

        Public Shared Sub ReleaseStringBuilderSuperCache(ByRef sb As Text.StringBuilder)
            ReleaseStringBuilderCache(sb, s_CachedStringBuilderSuper)
        End Sub

    End Class

    Public Class GBKEncoding
        Inherits Text.Encoding

        ''' <summary>
        ''' 获取 GBK 格式的编码
        ''' </summary>
        ''' <returns></returns>
        Public Shared ReadOnly Property GBK() As Text.Encoding = Text.Encoding.GetEncoding(936)

        Public Overrides Function GetByteCount(chars() As Char, index As Integer, count As Integer) As Integer
            Throw New NotImplementedException()
        End Function

        Public Overrides Function GetBytes(chars() As Char, charIndex As Integer, charCount As Integer, bytes() As Byte, byteIndex As Integer) As Integer
            Throw New NotImplementedException()
        End Function

        Public Overrides Function GetCharCount(bytes() As Byte, index As Integer, count As Integer) As Integer
            Throw New NotImplementedException()
        End Function

        Public Overrides Function GetChars(bytes() As Byte, byteIndex As Integer, byteCount As Integer, chars() As Char, charIndex As Integer) As Integer
            Throw New NotImplementedException()
        End Function

        Public Overrides Function GetMaxByteCount(charCount As Integer) As Integer
            Throw New NotImplementedException()
        End Function

        Public Overrides Function GetMaxCharCount(byteCount As Integer) As Integer
            Throw New NotImplementedException()
        End Function
    End Class
End Namespace

