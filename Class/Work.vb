
Namespace ShanXingTech.Work
    ''' <summary>
    ''' 这个人很懒，什么都没写
    ''' </summary>
    Public Class Utils
        ' 明文：神即道, 道法自然, 如来|闪星网络信息科技 ShanXingTech Q2287190283
        ' 算法：古典密码中的有密钥换位密码 密钥：ShanXingTech
        Public Const ShanXingTechQ2287190283 = "神闪X7,SQB道信T2道网N9来A2D如H2C然技HA即星I1|N8E法息E8,络G0自科C3"

        ''' <summary>
        ''' 关键词
        ''' </summary>
        ''' <returns></returns>
        Public Property KeyWord As String
        ''' <summary>
        ''' 店铺类型
        ''' </summary>
        ''' <returns></returns>
        Public Property ShopType As String

        ''' <summary>
        ''' 重写Equals 这样ArrayList的contain方法就可以比较类了
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        Public Overrides Function Equals(obj As Object) As Boolean
            Return KeyWord.Equals(DirectCast(obj, Utils).KeyWord) AndAlso ShopType.Equals(DirectCast(obj, Utils).ShopType)
        End Function

    End Class
End Namespace