Imports System.Collections.ObjectModel
Imports System.ComponentModel.DataAnnotations
Imports HR.Intranet.Server
Imports System.Data.SqlClient

Namespace ViewModels.Home
	' ReSharper disable once InconsistentNaming
	Public Class BulkBookingViewModel


		Public Property BookingStatuses As New Collection(Of SelectListItem)

		Public Property TableID() As Integer
		Public Property txt1000SepCols() As String
		Public Property CourseRecordID() As Integer
		Public Property TbStatusPExists() As String

		<Display(Name:="Booking Status :")> _
		Public Property BookingStatus As String()


		Public Sub New()

			BookingStatuses = getStatuses()
			TableID = HttpContext.Current.Session("TB_EmpTableID")
			txt1000SepCols = Get1000SepColumns()
			TbStatusPExists = HttpContext.Current.Session("TB_TBStatusPExists")
			CourseRecordID = HttpContext.Current.Session("optionRecordID")


		End Sub



		Private Function getStatuses() As Collection(Of SelectListItem)

			Dim objItems As New Collection(Of SelectListItem)

			Dim objRowItem As New SelectListItem() With {.Value = "B", .Text = "Booked", .Selected = True}
			objItems.Add(objRowItem)
			objRowItem = New SelectListItem() With {.Value = "P", .Text = "Provisional"}
			objItems.Add(objRowItem)

			Return objItems


		End Function


		Private Function Get1000SepColumns()
			Dim objDataAccess As clsDataAccess = CType(HttpContext.Current.Session("DatabaseAccess"), clsDataAccess)
			Dim prmErrorMsg As New SqlParameter("psErrorMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prm1000SepCols As New SqlParameter("ps1000SeparatorCols", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

			Dim rstFindRecords = objDataAccess.GetFromSP("sp_ASRIntGetTBEmployeeColumns" _
														, prmErrorMsg _
														, prm1000SepCols)

			Return prm1000SepCols.Value.ToString()

		End Function

	End Class
End Namespace