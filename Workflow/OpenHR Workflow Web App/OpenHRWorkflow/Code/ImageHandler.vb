Imports System.Collections
Imports System.Web

Public Class ImageHandler
   Implements IHttpHandler, IReadOnlySessionState

   Private Shared ReadOnly _contentMap As New Hashtable(StringComparer.CurrentCultureIgnoreCase)

   Shared Sub New()
      _contentMap.Add(".png", "image/png")
      _contentMap.Add(".bmp", "image/bmp")
      _contentMap.Add(".gif", "image/gif")
      _contentMap.Add(".ico", "image/x-icon")
      _contentMap.Add(".jpeg", "image/jpeg")
      _contentMap.Add(".*", "image/x-unknown")
   End Sub

   Public Sub ProcessRequest(context As HttpContext) Implements IHttpHandler.ProcessRequest

      Dim request = context.Request, response = context.Response, id As Integer

      If Not Integer.TryParse(request.QueryString("id"), id) Then
         'TODO return not found
         response.StatusCode = 304
         response.StatusDescription = "Not Modified"
         Return
      End If

      If Not [String].IsNullOrEmpty(context.Request.Headers("If-Modified-Since")) Then
         response.StatusCode = 304
         response.StatusDescription = "Not Modified"
         Return
      End If

      'Which database do we want to get the picture from?
      Dim db As Database
      If request.QueryString("s") IsNot Nothing Then
         'for workflows the database info is store in the session and comes from the workflow url
         Dim url As WorkflowUrl = CType(HttpContext.Current.Session("workflowUrl"), WorkflowUrl)
         db = New Database(Database.GetConnectionString(url.Server, url.Database, url.User, url.Password))
      Else
         db = New Database(Configuration.ConnectionString)
      End If

      'Stream the image from the database straight into the Response stream
      Dim fileName As String = Nothing
      db.GetPicture(id, response.OutputStream, fileName)

      response.ContentType = GetContentType(fileName)
      'response.OutputStream.Write(image, 0, image.Length)
      response.Cache.SetCacheability(HttpCacheability.[Public])
      response.Cache.SetExpires(DateTime.Now.AddYears(1))
      response.Cache.SetLastModified(DateTime.Now.AddYears(-1))
      response.End()
   End Sub

   Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
      Get
         Return True
      End Get
   End Property

   Private Shared Function GetContentType(fileName As String) As String
      Dim index As Integer = fileName.LastIndexOf("."c)
      Dim str As String = Nothing

      If index > 0 Then
         str = DirectCast(_contentMap(fileName.Substring(index)), String)
      End If
      If str Is Nothing Then
         str = DirectCast(_contentMap(".*"), String)
      End If

      Return str
   End Function
End Class
