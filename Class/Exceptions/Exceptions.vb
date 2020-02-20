Imports System.Runtime.Serialization


Namespace ShanXingTech.Exception2
	<Serializable>
	Public Class TimeStampException
		Inherits Exception

		Public Sub New(message As String)
			MyBase.New(message)
		End Sub

		Public Overrides Sub GetObjectData(info As SerializationInfo, context As StreamingContext)
			MyBase.GetObjectData(info, context)
		End Sub

		Public Sub New()
		End Sub

		Public Sub New(message As String, innerException As Exception)
			MyBase.New(message, innerException)
		End Sub

		Protected Sub New(serializationInfo As SerializationInfo, streamingContext As StreamingContext)
			MyBase.New(serializationInfo, streamingContext)
		End Sub
	End Class

	<Serializable>
	Public Class EngineNotFoundExcption
		Inherits Exception

		Public Sub New()
		End Sub

		Public Sub New(message As String)
			MyBase.New(message)
		End Sub

		Public Sub New(message As String, innerException As Exception)
			MyBase.New(message, innerException)
		End Sub

		Protected Sub New(info As SerializationInfo, context As StreamingContext)
			MyBase.New(info, context)
		End Sub
	End Class

	<Serializable>
	Public Class HttpAsyncUnInitializeException
		Inherits Exception

		Public Sub New(message As String)
			MyBase.New(message)
		End Sub

		Public Overrides Sub GetObjectData(info As SerializationInfo, context As StreamingContext)
			MyBase.GetObjectData(info, context)
		End Sub

		Public Sub New()
		End Sub

		Public Sub New(message As String, innerException As Exception)
			MyBase.New(message, innerException)
		End Sub

		Protected Sub New(serializationInfo As SerializationInfo, streamingContext As StreamingContext)
			MyBase.New(serializationInfo, streamingContext)
		End Sub
	End Class
End Namespace
