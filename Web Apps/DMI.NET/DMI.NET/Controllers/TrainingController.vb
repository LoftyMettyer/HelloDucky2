Option Strict On
Option Explicit On

Imports System.Data.SqlClient
Imports DMI.NET.Models.Responses
Imports HR.Intranet.Server

Namespace Controllers
	Public Class TrainingController

		Private ReadOnly Property objDataAccess As clsDataAccess
			Get
				Return CType(HttpContext.Current.Session("DatabaseAccess"), clsDataAccess)
			End Get
		End Property

		Private ReadOnly Property TB_TBTableID As Integer
			Get
				Return CInt(HttpContext.Current.Session("TB_TBTableID"))
			End Get
		End Property

		Private ReadOnly Property TB_CourseTableID As Integer
			Get
				Return CInt(HttpContext.Current.Session("TB_CourseTableID"))
			End Get
		End Property

		Private ReadOnly Property TB_TBStatusColumnID As Integer
			Get
				Return CInt(HttpContext.Current.Session("TB_TBStatusColumnID"))
			End Get
		End Property

		Private ReadOnly Property TB_EmpTableID As Integer
			Get
				Return CInt(HttpContext.Current.Session("TB_EmpTableID"))
			End Get
		End Property

		Private ReadOnly Property TB_CourseCancelDateColumnID As Integer
			Get
				Return CInt(HttpContext.Current.Session("TB_CourseCancelDateColumnID"))
			End Get
		End Property

		Private ReadOnly Property tbCourseTitle As String
			Get
				Return HttpContext.Current.Session("tbCourseTitle").ToString()
			End Get
		End Property

		Private ReadOnly Property TB_TBTableInsert As Boolean
			Get
				Return CBool(HttpContext.Current.Session("TB_TBTableInsert"))
			End Get
		End Property

		Private ReadOnly Property TB_TBCancelDateColumnID As Integer
			Get
				Return CInt(HttpContext.Current.Session("TB_TBCancelDateColumnID"))
			End Get
		End Property

		Private ReadOnly Property TB_WaitListTableID As Integer
			Get
				Return CInt(HttpContext.Current.Session("TB_WaitListTableID"))
			End Get
		End Property

		Private ReadOnly Property TB_WaitListTableInsert As String
			Get
				Return HttpContext.Current.Session("TB_WaitListTableInsert").ToString()
			End Get
		End Property

		Private ReadOnly Property TB_WaitListCourseTitleColumnID As Integer
			Get
				Return CInt(HttpContext.Current.Session("TB_WaitListCourseTitleColumnID"))
			End Get
		End Property

		Private ReadOnly Property TB_WaitListCourseTitleColumnUpdate As String
			Get
				Return HttpContext.Current.Session("TB_WaitListCourseTitleColumnUpdate").ToString()
			End Get
		End Property

		Public Function data_submit_CancelCourse(lngRecordID As Integer, sRealSource As String) As TrainingBookingResponse

			Dim result As New TrainingBookingResponse
			Dim sAction As String

			Try

				' Check number of bookings made.
				Dim prmNumberOfBookings = New SqlParameter("piNumberOfBookings", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmErrorMessage = New SqlParameter("psErrorMessage", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmCourseTitle = New SqlParameter("psCourseTitle", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

				objDataAccess.ExecuteSP("sp_ASRIntCancelCourse" _
					, prmNumberOfBookings _
					, New SqlParameter("piCourseRecordID", SqlDbType.Int) With {.Value = CleanNumeric(lngRecordID)} _
					, New SqlParameter("piTrainBookTableID", SqlDbType.Int) With {.Value = TB_TBTableID} _
					, New SqlParameter("piCourseTableID", SqlDbType.Int) With {.Value = TB_CourseTableID} _
					, New SqlParameter("piTrainBookStatusColumnID", SqlDbType.Int) With {.Value = TB_TBStatusColumnID} _
					, New SqlParameter("psCourseRealSource", SqlDbType.VarChar, -1) With {.Value = sRealSource} _
				, prmErrorMessage _
				, prmCourseTitle)

				sAction = "CANCELCOURSE_1"
				result.NumberOfBookings = CInt(prmNumberOfBookings.Value)
				result.Message = prmErrorMessage.Value.ToString()
				result.CourseTitle = prmCourseTitle.Value.ToString()

			Catch ex As Exception
				result.Message = "Error cancelling the course." & vbCrLf & ex.Message.RemoveSensitive()
				sAction = "SAVEERROR"

			End Try

			result.Action = sAction


			Return result

		End Function

		Public Function data_submit_CancelCourse2(lngRecordID As Integer, sRealSource As String, iTBCourseRecordID As Integer, bCreateWLRecords As Boolean) As TrainingBookingResponse

			Dim result As New TrainingBookingResponse
			Dim sAction As String
			Try

				Dim prmErrorMessage = New SqlParameter("psErrorMessage", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

				objDataAccess.ExecuteSP("sp_ASRIntCancelCoursePart2" _
					, New SqlParameter("piEmployeeTableID", SqlDbType.Int) With {.Value = TB_EmpTableID} _
					, New SqlParameter("piCourseTableID", SqlDbType.Int) With {.Value = TB_CourseTableID} _
					, New SqlParameter("psCourseRealSource", SqlDbType.VarChar, -1) With {.Value = sRealSource} _
					, New SqlParameter("piCourseRecordID", SqlDbType.Int) With {.Value = CleanNumeric(lngRecordID)} _
					, New SqlParameter("piTransferCourseRecordID", SqlDbType.Int) With {.Value = iTBCourseRecordID} _
					, New SqlParameter("piCourseCancelDateColumnID", SqlDbType.Int) With {.Value = TB_CourseCancelDateColumnID} _
					, New SqlParameter("psCourseTitle", SqlDbType.VarChar, -1) With {.Value = tbCourseTitle} _
					, New SqlParameter("piTrainBookTableID", SqlDbType.Int) With {.Value = TB_TBTableID} _
					, New SqlParameter("pfTrainBookTableInsert", SqlDbType.Bit) With {.Value = TB_TBTableInsert} _
					, New SqlParameter("piTrainBookStatusColumnID", SqlDbType.Int) With {.Value = TB_TBStatusColumnID} _
					, New SqlParameter("piTrainBookCancelDateColumnID", SqlDbType.Int) With {.Value = TB_TBCancelDateColumnID} _
					, New SqlParameter("piWaitListTableID", SqlDbType.Int) With {.Value = TB_WaitListTableID} _
					, New SqlParameter("pfWaitListTableInsert", SqlDbType.Bit) With {.Value = TB_WaitListTableInsert} _
					, New SqlParameter("piWaitListCourseTitleColumnID", SqlDbType.Int) With {.Value = TB_WaitListCourseTitleColumnID} _
					, New SqlParameter("pfWaitListCourseTitleColumnUpdate", SqlDbType.Bit) With {.Value = TB_WaitListCourseTitleColumnUpdate} _
					, New SqlParameter("pfCreateWaitListRecords", SqlDbType.Bit) With {.Value = bCreateWLRecords} _
					, prmErrorMessage)

				result.Message = prmErrorMessage.Value.ToString()

				If Len(result.Message) > 0 Then
					sAction = "SAVEERROR"
				Else
					sAction = "LOAD"
				End If

			Catch ex As Exception
				result.Message = "Error cancelling the course." & vbCrLf & FormatError(ex.Message)
				sAction = "SAVEERROR"

			End Try

			result.Action = sAction
			Return result

		End Function

		Public Function data_submit_CancelBooking(transferToWaitingList As Boolean, lngRecordID As Integer) As TrainingBookingResponse

			Dim result As New TrainingBookingResponse

			Try

				Dim prmErrorMessage = New SqlParameter("psErrorMessage", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

				objDataAccess.ExecuteSP("sp_ASRIntCancelBooking" _
					, New SqlParameter("pfTransferBookings", SqlDbType.Bit) With {.Value = transferToWaitingList} _
					, New SqlParameter("piTBRecordID", SqlDbType.Int) With {.Value = lngRecordID} _
					, prmErrorMessage)

				If Len(prmErrorMessage.Value.ToString()) > 0 Then
					result.Message = prmErrorMessage.Value.ToString()
					result.Action = "SAVEERROR"
				Else
					result.Action = "CANCELBOOKING_1"
				End If

			Catch ex As Exception
				result.Message = "Error cancelling the booking." & vbCrLf & ex.Message.RemoveSensitive()
				result.Action = "SAVEERROR"

			End Try

			Return result

		End Function

	End Class
End Namespace