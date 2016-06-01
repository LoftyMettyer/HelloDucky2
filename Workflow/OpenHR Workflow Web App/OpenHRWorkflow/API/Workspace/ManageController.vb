Imports System.Web.Http
Imports Newtonsoft.Json
Imports OpenHRWorkflow.Classes.Workspace

Namespace API.Workspace

   Public Class ManageController
      Inherits ApiController

      <HttpGet>
      Function Ping() As String
         Dim FeedParticipantResponse As New FeedParticipantResponse With {.result = "SUCCESS", .message = "Ping Response", .data = Nothing}

         Return JsonConvert.SerializeObject(New With {FeedParticipantResponse})
      End Function
   End Class
End NameSpace