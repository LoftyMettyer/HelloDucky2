Imports System.Web.Http
Imports Newtonsoft.Json
Imports OpenHRWorkflow.Classes.Workspace

Namespace API.Workspace

   Public Class ManageController
      Inherits ApiController

      Private Class FeedParticipantPingResponse
         Inherits FeedParticipantResponseBase

         Public data As String
      End Class

      <HttpGet>
      Function Ping() As String
         Dim FeedParticipantResponse As New FeedParticipantPingResponse With {.result = "SUCCESS", .message = "Ping Response", .data = Nothing}

         Return JsonConvert.SerializeObject(New With {FeedParticipantResponse})
      End Function
   End Class
 End Namespace