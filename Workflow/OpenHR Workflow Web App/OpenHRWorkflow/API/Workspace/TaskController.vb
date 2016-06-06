Imports System.Collections.ObjectModel
Imports System.Data.SqlClient
Imports System.Web.Http
Imports Newtonsoft.Json
Imports OpenHRWorkflow.Classes.Workspace

Namespace API.Workspace
   Public Class TaskController
      Inherits ApiController

      Private Class GetTaskListResponse
         Inherits FeedParticipantResponseBase

         Public responseCode As FeedParticipantResponseCodeEnum
         Public data As Collection(Of FeedUserTask)
      End Class

      <HttpGet>
      Function GetTasklist() As String

         Dim FeedParticipantResponse As New GetTaskListResponse With {
            .responseCode = FeedParticipantResponseCodeEnum.Success,
            .message = FeedParticipantResponseCodeEnum.Success.ToString,
            .data = New Collection(Of FeedUserTask)
         }

         If String.IsNullOrEmpty(General.Global_WorkspaceUserId) Then
            FeedParticipantResponse.result = "FAIL"
            FeedParticipantResponse.message = "Invalid WorkspaceUserId"
            FeedParticipantResponse.responseCode = FeedParticipantResponseCodeEnum.InvalidUser

            Return JsonConvert.SerializeObject(New With {FeedParticipantResponse})
         End If

         Try
            Dim db As New Database(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)
            Dim ds As DataSet = db.GetDataSet("spASRWorkspaceCheckPendingWorkflowSteps",
             CommandType.StoredProcedure,
             New SqlParameter("@username", General.Global_WorkspaceUserId))

            ' I wish I had entity framework at my disposal here!
            If ds.Tables.Count > 0 Then
               For Each dr As DataRow In ds.Tables(0).Rows
                  FeedParticipantResponse.data.Add(New FeedUserTask With {
                     .id = dr("instanceStepID").ToString(),
                     .category = "OpenHR Workflow",
                     .header1 = dr("name").ToString(),
                     .header2 = dr("description").ToString()})
               Next
            End If
         Catch ex As Exception
            FeedParticipantResponse.result = "FAIL"
            FeedParticipantResponse.message = ex.Message
            FeedParticipantResponse.responseCode = FeedParticipantResponseCodeEnum.CodeException
         End Try

         FeedParticipantResponse.result = "SUCCESS"
         Return JsonConvert.SerializeObject(New With {FeedParticipantResponse})

      End Function
   End Class
End Namespace