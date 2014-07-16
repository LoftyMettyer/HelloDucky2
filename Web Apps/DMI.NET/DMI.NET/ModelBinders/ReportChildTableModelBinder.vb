
Imports System.Web.Mvc
Imports System.IO
Imports System.Web.Script.Serialization
Imports DMI.NET.Classes

Public Class ReportChildTableModelBinder
	Implements IModelBinder

	Public Function BindModel(controllerContext As ControllerContext, bindingContext As ModelBindingContext) As Object Implements IModelBinder.BindModel
		Dim contentType = controllerContext.HttpContext.Request.ContentType
		If Not contentType.StartsWith("application/json", StringComparison.OrdinalIgnoreCase) Then
			Return (Nothing)
		End If

		Dim bodyText As String

		Using stream = controllerContext.HttpContext.Request.InputStream
			stream.Seek(0, SeekOrigin.Begin)
			Using reader = New StreamReader(stream)
				bodyText = reader.ReadToEnd()
			End Using
		End Using

		If String.IsNullOrEmpty(bodyText) Then
			Return (Nothing)
		End If

		Dim tweet = New JavaScriptSerializer().Deserialize(Of ReportChildTables)(bodyText)

		Return (tweet)
	End Function

End Class
