Namespace ShanXingTech.Net2
    Public Enum UploadMode
        ''' <summary>
        ''' 通过MD5比较，如果文件存在，则不上传；如果文件不存在，则上传
        ''' </summary>
        Create
        ''' <summary>
        ''' 通过MD5比较，如果文件存在，则上传之后会自动采用 “增量重命名” 方法重命名，格式为 “原文件名(数字)” ；如果文件不存在，则上传
        ''' </summary>
        RenameIncrementallyOrCreate
    End Enum
End Namespace