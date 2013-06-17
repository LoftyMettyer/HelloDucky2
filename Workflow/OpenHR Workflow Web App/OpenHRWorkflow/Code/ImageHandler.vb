﻿Imports System.Collections
Imports System.Web

Public Class ImageHandler
   Implements IHttpHandler, IReadOnlySessionState

	Private Shared ReadOnly ContentMap As New Hashtable(StringComparer.CurrentCultureIgnoreCase)

	Shared Sub New()
		ContentMap.Add(".png", "image/png")
		ContentMap.Add(".bmp", "image/bmp")
		ContentMap.Add(".gif", "image/gif")
		ContentMap.Add(".ico", "image/x-icon")
		ContentMap.Add(".jpeg", "image/jpeg")
		ContentMap.Add(".*", "image/x-unknown")
	End Sub

	Public Sub ProcessRequest(context As HttpContext) Implements IHttpHandler.ProcessRequest

		Dim request = context.Request, response = context.Response, id As Integer

		Integer.TryParse(request.QueryString("id"), id)

		If Not context.Request.Headers("If-Modified-Since").IsNullOrEmpty() Then
			response.StatusCode = 304
			response.StatusDescription = "Not Modified"
			Return
		End If

		'Which database do we want to get the picture from?
		Dim db As Database

		If request.QueryString("s") IsNot Nothing Then

			'for workflows the database info is store in the session and comes from the workflow url
			Dim url As WorkflowUrl = CType(HttpContext.Current.Session("workflowUrl"), WorkflowUrl)
			db = New Database(Database.CreateConnectionString(url.Server, url.Database, url.User, url.Password))
		Else
			db = New Database(App.Config.ConnectionString)
		End If

		'Stream the image from the database straight into the Response stream
		Dim picture = db.GetPicture(id)

		If picture Is Nothing Then
			Throw New HttpException(404, "Image not found")
		End If

		response.ContentType = GetContentType(picture.Name)
		response.OutputStream.Write(picture.Image, 0, picture.Image.Length)
		response.Cache.SetCacheability(HttpCacheability.[Public])
		response.Cache.SetExpires(DateTime.Now.AddYears(1))
		response.Cache.SetLastModified(DateTime.Now.AddYears(-1))
		response.End()
	End Sub

	Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
		Get
			Return False
		End Get
	End Property

	Private Shared Function GetContentType(fileName As String) As String
		Dim index As Integer = fileName.LastIndexOf("."c)
		Dim str As String = Nothing

		If index > 0 Then
			str = DirectCast(ContentMap(fileName.Substring(index)), String)
		End If
		If str Is Nothing Then
			str = DirectCast(ContentMap(".*"), String)
		End If

		Return str
	End Function
End Class
