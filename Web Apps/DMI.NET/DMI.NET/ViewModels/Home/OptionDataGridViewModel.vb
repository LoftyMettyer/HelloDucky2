Imports System.ComponentModel.DataAnnotations
Imports System.Web.HttpContext
Imports System.Collections.ObjectModel
Imports HR.Intranet.Server
Imports System.Data.SqlClient

Namespace ViewModels

	Public Class OptionDataGridViewModel

		<Display(Name:="View :")>
		Public Property Views As New Collection(Of SelectListItem)
		<Display(Name:="Order :")>
		Public Property Orders As New Collection(Of SelectListItem)

		Public Property OrderId() As Integer
		Public Property ViewId() As Integer

		Public Property TableID As Integer
		Public Property CourseTitle As String
		Public Property RecordID As Integer

		Public Property DataFrameSource As String
		Public Property OptionAction As String
		Public Property GotoOptionActionSelect As String
		Public Property GotoOptionActionCancel As String
		Public Property PageTitle As String
		Public Property SubmitAction As String

		Public Sub New(GotoOptionPage As String)

			Views = getViews()
			Orders = getOrders()

			OrderId = CInt(Current.Session("TB_BulkBookingDefaultOrderID"))	' CInt(Current.Session("optionLinkOrderID"))
			ViewId = CInt(Current.Session("TB_BulkBookingDefaultViewID"))

			TableID = NullSafeInteger(Current.Session("optionLinkTableID"))
			CourseTitle = NullSafeString(Current.Session("optionCourseTitle"))
			RecordID = NullSafeInteger(Current.Session("optionRecordID"))

			Select Case GotoOptionPage
				Case "tbTransferCourseFind"
					DataFrameSource = "TBTRANSFERCOURSEFIND"
					OptionAction = "LOADTRANSFERCOURSE"
					GotoOptionActionSelect = OptionActionType.SELECTTRANSFERCOURSE
					GotoOptionActionCancel = OptionActionType.SELECTTRANSFERCOURSE
					PageTitle = "Find Course Record"
					SubmitAction = "tbTransferCourseFind_Submit"

				Case "tbTransferBookingFind"
					DataFrameSource = "TBTRANSFERBOOKINGFIND"
					OptionAction = "LOADTRANSFERBOOKING"
					GotoOptionActionSelect = OptionActionType.SELECTTRANSFERBOOKING_1
					GotoOptionActionCancel = OptionActionType.CANCEL
					PageTitle = "Transfer Booking"
					SubmitAction = "tbTransferBookingFind_Submit"

				Case "tbBookCourseFind"
					DataFrameSource = "TBBOOKCOURSEFIND"
					OptionAction = "LOADBOOKCOURSE"
					GotoOptionActionSelect = OptionActionType.SELECTBOOKCOURSE_1
					GotoOptionActionCancel = OptionActionType.CANCEL
					PageTitle = "Book Course"
					SubmitAction = "tbBookCourseFind_Submit"

				Case "tbAddFromWaitingListFind"
					DataFrameSource = "TBADDFROMWAITINGLISTFIND"
					OptionAction = "LOADADDFROMWAITINGLIST"
					GotoOptionActionSelect = OptionActionType.SELECTADDFROMWAITINGLIST_1
					GotoOptionActionCancel = OptionActionType.CANCEL
					PageTitle = "Add From Waiting List"
					SubmitAction = "tbAddFromWaitingListFind_Submit"


			End Select

		End Sub


		Private Function getViews() As Collection(Of SelectListItem)

			Dim objDatabase As Database = CType(Current.Session("DatabaseFunctions"), Database)
			Dim objItems As New Collection(Of SelectListItem)

			Dim prmDfltOrderID As New SqlParameter("plngDfltOrderID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim rstViewRecords = objDatabase.DB.GetDataTable("sp_ASRIntGetLinkViews", CommandType.StoredProcedure _
					, New SqlParameter("plngTableID", SqlDbType.Int) With {.Value = CleanNumeric(Current.Session("optionLinkTableID"))} _
					, prmDfltOrderID)

			For Each objRow As DataRow In rstViewRecords.Rows

				Dim objRowItem As New SelectListItem() With {.Value = CStr(objRow(0)), .Text = Replace(objRow(1).ToString(), "_", " ")}
				objItems.Add(objRowItem)

			Next

			If Current.Session("optionLinkOrderID") <= 0 Then
				Current.Session("optionLinkOrderID") = prmDfltOrderID.Value
			End If


			Return objItems


		End Function

		Private Function getOrders() As Collection(Of SelectListItem)

			Dim objDatabase As Database = CType(Current.Session("DatabaseFunctions"), Database)
			Dim objItems As New Collection(Of SelectListItem)


			Dim rstOrderRecords = objDatabase.GetTableOrders(CInt(CleanNumeric(Current.Session("optionLinkTableID"))), 0)
			For Each objRow As DataRow In rstOrderRecords.Rows

				Dim objRowItem As New SelectListItem() With {.Value = CStr(objRow(1)), .Text = Replace(objRow(0).ToString(), "_", " ")}
				objItems.Add(objRowItem)

			Next

			Return objItems

		End Function



	End Class


End Namespace
