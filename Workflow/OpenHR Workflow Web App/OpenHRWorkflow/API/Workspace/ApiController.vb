Imports System.Collections.ObjectModel
Imports System.Data.SqlClient
Imports System.Web.Mvc
Imports System.Web.Script.Serialization
Imports Newtonsoft.Json
Imports OpenHRWorkflow.Classes.Workspace

Namespace API.Workspace
	Public Class ApiController
		Inherits Http.ApiController

		Public Class FeedParticipantResponse
			Inherits FeedParticipantResponseBase

			Public responseCode As FeedParticipantResponseCodeEnum
			Public data As List(Of Task)
		End Class

		Public Class FeedParticipantResponseWrapper
			Public FeedParticipantResponse As FeedParticipantResponse
		End Class

		<Http.HttpGet>
		Public Function Task() As FeedParticipantResponseWrapper
			'Uncomment the bit below to generate test tasks without needing a database connection
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'Dim FeedParticipantResponseTest As New FeedParticipantResponse With {
			'	 .responseCode = FeedParticipantResponseCodeEnum.Success,
			'	 .message = FeedParticipantResponseCodeEnum.Success.ToString,
			'	 .data = CreateTestTasks()
			'}
			'FeedParticipantResponseTest.result = "SUCCESS"
			'Return New FeedParticipantResponseWrapper() With {.FeedParticipantResponse = FeedParticipantResponseTest}
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

			Dim FeedParticipantResponse As New FeedParticipantResponse With {
						.responseCode = FeedParticipantResponseCodeEnum.Success,
						.message = FeedParticipantResponseCodeEnum.Success.ToString,
						.data = New List(Of Task)
				 }

			'Validation
			If String.IsNullOrEmpty(General.Global_WorkspaceUserId) Then
				FeedParticipantResponse.result = "FAIL"
				FeedParticipantResponse.message = "Invalid WorkspaceUserId"
				FeedParticipantResponse.responseCode = FeedParticipantResponseCodeEnum.InvalidUser

				Return New FeedParticipantResponseWrapper() With {.FeedParticipantResponse = FeedParticipantResponse}
			End If

			Try
				Dim db As New Database(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)
				Dim ds As DataSet = db.GetDataSet("spASRWorkspaceCheckPendingWorkflowSteps", CommandType.StoredProcedure, New SqlParameter("@username", General.Global_WorkspaceUserId))

				If ds.Tables.Count > 0 Then
					Dim taskAction As New TaskAction
					taskAction.name = "Open"
					taskAction.label = "Open"
					taskAction.showOnListScreen = "Y"
					taskAction.style = "default"
					For Each dr As DataRow In ds.Tables(0).Rows
						taskAction.url = dr("URL").ToString
						Dim t As New Task With {
						 .id = dr("instanceStepID").ToString(),
						 .category = "OpenHR Workflows",
						 .subCategory = dr("name").ToString(),
						 .status = "Pending",
						 .sourceName = "SourceSystem",
						 .sourceId = "SourceSystemId",
						 .createdBy = "User",
						 .priority = 2
						 }

						t.headers.Add(dr("description").ToString())

						t.actions.Add(taskAction)

						FeedParticipantResponse.data.Add(t)
					Next
				End If
			Catch ex As Exception
				FeedParticipantResponse.result = "FAIL"
				FeedParticipantResponse.message = ex.Message
				FeedParticipantResponse.responseCode = FeedParticipantResponseCodeEnum.CodeException
			End Try

			FeedParticipantResponse.result = "SUCCESS"
			Return New FeedParticipantResponseWrapper() With {.FeedParticipantResponse = FeedParticipantResponse}

		End Function

#Region "Test tasks creation"

		'The tasks are built as follows:
		Private Function CreateTestTasks() As List(Of Task)
			Dim tasks As New List(Of Task)()
			Dim now As New [DateTime]()
			Dim actionBefore As [DateTime] = now.AddHours(3)

			'actions
			Dim actions As New List(Of TaskAction)
			Dim taskAction As New TaskAction '
			taskAction.name = "Open"
			taskAction.label = "Open"
			taskAction.url = "http://www.google.co.uk"
			taskAction.showOnListScreen = "Y"
			taskAction.style = ""
			actions.Add(taskAction)

			'Notification 1
			Dim task As New Task
			task.id = "111"
			task.category = "OpenHR Workflows"
			task.subCategory = "External Requisition"
			task.created = now
			task.createdBy = "Richard Beney"
			task.actionBefore = actionBefore
			task.priority = 1
			task.status = "Unread"
			task.sourceName = "SourceSystem"
			task.sourceId = "SourceSystemId"
			tasks.Add(task)
			task.headers.Add("Header 1")
			task.actions = actions
			task.images = New List(Of TaskImage)

			'Notification 2
			task = New Task
			task.id = "222"
			task.category = "OpenHR Workflows"
			task.subCategory = "External Requisition"
			task.created = now
			task.createdBy = "Richard Beney"
			task.actionBefore = actionBefore
			task.priority = 1
			task.status = "Unread"
			task.sourceName = "SourceSystem"
			task.sourceId = "SourceSystemId"
			tasks.Add(task)
			task.headers.Add("Header 2")
			task.actions = actions
			task.images = New List(Of TaskImage)

			Return tasks
		End Function
#End Region
	End Class
End Namespace