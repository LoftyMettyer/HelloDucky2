Imports System.ComponentModel.DataAnnotations
Imports System.Collections.ObjectModel
Imports System.Web.HttpContext
Imports HR.Intranet.Server
Imports System.Data.SqlClient

Namespace ViewModels.Home

	Public Class BulkBookingSelectionViewModel

		Public Property Views As New Collection(Of SelectListItem)
		Public Property Orders As New Collection(Of SelectListItem)

		Public Property TableID As Integer

		Public Property FirstRecPos As Integer
		Public Property CurrentRecCount As Integer
		Public Property PageAction As String

		<Display(Name:="View :")> _
		Public Property View As String()

		<Display(Name:="Order :")> _
		Public Property Order As String()



		Public Sub New()

			Current.Session("optionLinkViewID") = Current.Session("TB_BulkBookingDefaultViewID")
			Current.Session("optionLinkOrderID") = 0

			Views = getViews()
			Orders = getOrders()

			TableID = NullSafeInteger(Current.Session("TB_EmpTableID"))

			FirstRecPos = 1
			CurrentRecCount = 0
			PageAction = "LOAD"

		End Sub

		Private Function getViews() As Collection(Of SelectListItem)

			Dim objDatabase As Database = CType(Current.Session("DatabaseFunctions"), Database)
			Dim objItems As New Collection(Of SelectListItem)

			Dim prmDfltOrderID As New SqlParameter("plngDfltOrderID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim rstViewRecords = objDatabase.DB.GetDataTable("sp_ASRIntGetLinkViews", CommandType.StoredProcedure _
					, New SqlParameter("plngTableID", SqlDbType.Int) With {.Value = CleanNumeric(Current.Session("TB_EmpTableID"))} _
					, prmDfltOrderID)

			For Each objRow As DataRow In rstViewRecords.Rows

				Dim objRowItem As New SelectListItem() With {.Value = CStr(objRow(0)), .Text = Replace(objRow(1).ToString(), "_", " "), .Selected = (CInt(objRow(0)) = CInt(Current.Session("optionLinkViewID")))}
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


			Dim rstOrderRecords = objDatabase.GetTableOrders(CInt(CInt(CleanNumeric(Current.Session("TB_EmpTableID")))), 0)
			For Each objRow As DataRow In rstOrderRecords.Rows

				Dim objRowItem As New SelectListItem() With {.Value = CStr(objRow(1)), .Text = Replace(objRow(0).ToString(), "_", " "), .Selected = (objRow(1) = CInt(Current.Session("optionLinkOrderID")))}
				objItems.Add(objRowItem)

			Next

			Return objItems

		End Function
	End Class




End Namespace