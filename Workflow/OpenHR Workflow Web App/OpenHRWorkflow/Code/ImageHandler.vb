
Public Class ImageHandler
	Implements IHttpHandler, IReadOnlySessionState

	Private Shared ReadOnly ContentMap As New Hashtable(StringComparer.InvariantCultureIgnoreCase)

	Shared Sub New()
		ContentMap.Add(".png", "image/png")
		ContentMap.Add(".bmp", "image/bmp")
		ContentMap.Add(".gif", "image/gif")
		ContentMap.Add(".ico", "image/x-icon")
		ContentMap.Add(".jpeg", "image/jpeg")
		ContentMap.Add(".jpg", "image/jpeg")
		ContentMap.Add(".*", "image/x-unknown")
	End Sub

	Public Sub ProcessRequest(context As HttpContext) Implements IHttpHandler.ProcessRequest

		Dim request = context.Request,
			response = context.Response,
			id As Integer

		Integer.TryParse(request.QueryString("id"), id)

		'Which database do we want to get the picture from?
		Dim conn As String
		conn = App.Config.ConnectionString

		Dim db As New Database(conn)
		Dim picture = db.GetPicture(id)

		If picture Is Nothing Then
			Throw New HttpException(404, "Image not found")
		End If

		response.ContentType = GetContentType(picture.Name)
		response.OutputStream.Write(picture.Image, 0, picture.Image.Length)
		response.Cache.SetCacheability(HttpCacheability.Public)
		response.Cache.SetMaxAge(TimeSpan.FromDays(90))
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
