Imports System.Web.Http
Imports Newtonsoft.Json
Imports OpenHRWorkflow.Classes.Workspace

Namespace API.Workspace

   Public Class ManageController
      Inherits ApiController

      Public Class FeedParticipantPingResponse
         Inherits FeedParticipantResponseBase

         Public data As String
      End Class

      Public Class FeedParticipantPingResponseWrapper
         Public FeedParticipantPingResponse As FeedParticipantPingResponse
      End Class

      <HttpGet>
      Public Function Ping() As FeedParticipantPingResponseWrapper
         Dim FeedParticipantPingResponse As New FeedParticipantPingResponse With {.result = "SUCCESS", .message = "Ping Response", .data = Nothing}

         Return New FeedParticipantPingResponseWrapper() With {.FeedParticipantPingResponse = FeedParticipantPingResponse}
      End Function
   End Class
End Namespace