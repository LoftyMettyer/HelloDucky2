Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Threading
Imports System.Threading.Tasks

Namespace API
   ''' <summary>
   ''' Intercept incoming API requests
   ''' </summary>
   Public Class IncomingRequestInterceptor
      Inherits DelegatingHandler

      Protected Overrides Function SendAsync(request As HttpRequestMessage, cancellationToken As CancellationToken) As Task(Of HttpResponseMessage)
         Const workspaceCookieName = "iplanetdirectorypro" 'Name of workspace cookie returned after a user has authenticated

         If Not request.Headers.GetCookies(workspaceCookieName).FirstOrDefault Is Nothing Then
            Dim cookie As CookieHeaderValue = request.Headers.GetCookies(workspaceCookieName).FirstOrDefault()
            Dim workspaceTokenId = cookie(workspaceCookieName).Value
            Try
               General.Global_WorkspaceUserId = OpenAmRestCalls.GetIdFromSession(workspaceTokenId)
            Catch ex As Exception
               General.Global_WorkspaceUserId = ""
            End Try
         End If

         Return MyBase.SendAsync(request, cancellationToken)
      End Function

   End Class
End Namespace