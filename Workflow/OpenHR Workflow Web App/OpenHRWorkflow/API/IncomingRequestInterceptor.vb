Imports System.IO
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Web.Hosting

Namespace API
	''' <summary>
	''' Intercept incoming API requests
	''' </summary>
	Public Class IncomingRequestInterceptor
		Inherits DelegatingHandler

		Protected Overrides Function SendAsync(request As HttpRequestMessage, cancellationToken As CancellationToken) As Task(Of HttpResponseMessage)
			'Log request
			'File.AppendAllText(HttpContext.Current.Server.MapPath("~/WorkflowAPITraffic.txt"),
			'					 "[" + DateTime.Now.ToString() + "] Request: " +
			'					request.RequestUri.ToString +
			'					Environment.NewLine
			'					)

			'Try to get the WorkspaceUserId using two different methods:
			General.Global_WorkspaceUserId = ""

			'Method 1: Using a cookie
			Const workspaceCookieName = "iplanetdirectorypro" 'Name of workspace cookie returned after a user has authenticated
			If Not request.Headers.GetCookies(workspaceCookieName).FirstOrDefault Is Nothing Then
				Try
					Dim cookie As CookieHeaderValue = request.Headers.GetCookies(workspaceCookieName).FirstOrDefault()
					Dim workspaceTokenId = cookie(workspaceCookieName).Value
					General.Global_WorkspaceUserId = OpenAmRestCalls.GetIdFromSession(workspaceTokenId)
				Catch ex As Exception

				End Try
			End If

			'Method 2: Using a parameter in the querystring
			If General.Global_WorkspaceUserId = "" Then
				Const queryStringParameterWithUsername = "ADV_feedUser" 'Name of query string parameter containing the user name
				Dim queryString = request.GetQueryNameValuePairs()
				General.Global_WorkspaceUserId = queryString.FirstOrDefault(Function(k) k.Key = queryStringParameterWithUsername).Value
				If General.Global_WorkspaceUserId = Nothing Then
					General.Global_WorkspaceUserId = ""
				End If
			End If

			'Return the response
			Return MyBase.SendAsync(request, cancellationToken)

			'Return MyBase.SendAsync(request, cancellationToken).ContinueWith(
			'		Function(response)
			'			Return response.Result
			'		End Function
			')
		End Function
	End Class
End Namespace