Option Strict On
Option Explicit On

Imports System.Data.SqlClient
Imports DMI.NET.Classes
Imports DMI.NET.Models.Responses
Imports HR.Intranet.Server

Namespace Controllers
	Public Class RecordController

		Private ReadOnly Property objDataAccess As clsDataAccess
			Get
				Return CType(HttpContext.Current.Session("DatabaseAccess"), clsDataAccess)
			End Get
		End Property

		Private ReadOnly Property _userName As String
			Get
				Return HttpContext.Current.Session("Username").ToString()
			End Get
		End Property

		Private ReadOnly Property TB_TBTableID As Integer
			Get
				Return CInt(HttpContext.Current.Session("TB_TBTableID"))
			End Get
		End Property

		Public Function data_submit_DELETE(lngTableID As Integer, sRealSource As String, lngRecordID As Integer, sReaction As String) As PostResponse

			Dim result As New PostResponse

			Try

				Dim prmResult As New SqlParameter("piResult", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				objDataAccess.ExecuteSP("sp_ASRDeleteRecord" _
						, prmResult _
						, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = lngTableID} _
						, New SqlParameter("psRealSource", SqlDbType.VarChar, 255) With {.Value = sRealSource} _
						, New SqlParameter("piID", SqlDbType.Int) With {.Value = lngRecordID})

				Select Case CType(prmResult.Value, SaveResponse)
					Case SaveResponse.AmendedByAnotherUser
						result.Message = "The record has been amended by another user and will be refreshed."
				End Select

				result.RecordID = 0

				If Len(sReaction) > 0 Then
					result.Action = sReaction
				Else
					result.Action = "LOAD"
				End If

				objDataAccess.ExecuteSP("spASREmailImmediate" _
						, New SqlParameter("@Username", SqlDbType.VarChar, 255) With {.Value = _userName})

			Catch ex As Exception
				result.Message = "The record could not be deleted." & vbCrLf & ex.Message.RemoveSensitive()
				result.Action = "SAVEERROR"

			End Try

			Return result

		End Function

		Public Function data_submit_SAVE(lngTableID As Integer, lngRecordID As Integer, sReaction As String, fTBOverride As Boolean _
			, iTBEmployeeRecordID As Integer, iTBCourseRecordID As Integer, sTBBookingStatusValue As String, sInsertUpdateDef As String _
			, sRealSource As String, iTimestamp As Integer, lngOriginalRecordID As Integer) As SaveRecordRepsonse

			Dim result As New SaveRecordRepsonse

			Dim sErrorMsg = ""
			Dim sAction As String
			Dim sCourseOverbooked As String
			Dim fOK As Boolean
			Dim fWarning As Boolean

			Dim sTBErrorMsg = ""
			Dim sTBWarningMsg = ""
			Dim sCode = ""
			Dim sTBResultCode As String

			If Not fTBOverride AndAlso (NullSafeInteger(lngTableID) = NullSafeInteger(TB_TBTableID) AndAlso Licence.IsModuleLicenced(SoftwareModule.Training)) Then
				' Training Booking check.
				Try
					Dim prmResult = New SqlParameter("@piResultCode", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
					Dim prmCourseOverbooked = New SqlParameter("@psCourseOverbooked", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

					objDataAccess.ExecuteSP("sp_ASRIntValidateTrainingBooking" _
						, prmResult _
						, New SqlParameter("piEmpRecID", SqlDbType.Int) With {.Value = iTBEmployeeRecordID} _
						, New SqlParameter("piCourseRecID", SqlDbType.Int) With {.Value = iTBCourseRecordID} _
						, New SqlParameter("psBookingStatus", SqlDbType.VarChar, -1) With {.Value = sTBBookingStatusValue} _
						, New SqlParameter("piTBRecID", SqlDbType.Int) With {.Value = lngRecordID} _
						, prmCourseOverbooked)

					sTBResultCode = prmResult.Value.ToString()
					sCourseOverbooked = prmCourseOverbooked.Value.ToString()
				Catch ex As Exception
					sErrorMsg = "Error validating training booking." & vbCrLf & ex.Message.RemoveSensitive()
				End Try

				If Len(sErrorMsg) = 0 Then
					If CInt(sTBResultCode) > 0 Then
						If Len(sTBResultCode) = 4 Then
							' Get the overbooking check code.
							sCode = Left(sTBResultCode, 1)
							If sCode = "1" Then
								sTBErrorMsg = "The course is already fully booked. Unable to make the booking."
							Else
								If sCode = "2" Then
									sTBWarningMsg = "The course is already fully booked. Unable to make the booking."
								End If
							End If
						End If

						If Len(sTBResultCode) >= 3 Then
							' Get the pre-requisite check code.
							sCode = Mid(sTBResultCode, Len(sTBResultCode) - 2, 1)
							If sCode = "1" Then
								If Len(sTBErrorMsg) > 0 Then
									sTBErrorMsg = sTBErrorMsg & vbCrLf
								End If
								sTBErrorMsg = sTBErrorMsg & "The delegate has not met the pre-requisites for the course. Unable to make the booking."
							Else
								If sCode = "2" Then
									If Len(sTBWarningMsg) > 0 Then
										sTBWarningMsg = sTBWarningMsg & vbCrLf
									End If
									sTBWarningMsg = sTBWarningMsg & "The delegate has not met the pre-requisites for the course."
								End If
							End If
						End If

						If Len(sTBResultCode) >= 2 Then
							' Get the availability check code.
							sCode = Mid(sTBResultCode, Len(sTBResultCode) - 1, 1)
							If sCode = "1" Then
								If Len(sTBErrorMsg) > 0 Then
									sTBErrorMsg = sTBErrorMsg & vbCrLf
								End If
								sTBErrorMsg = sTBErrorMsg & "The delegate is unavailable for the course."
							Else
								If sCode = "2" Then
									If Len(sTBWarningMsg) > 0 Then
										sTBWarningMsg = sTBWarningMsg & vbCrLf
									End If
									sTBWarningMsg = sTBWarningMsg & "The delegate is unavailable for the course."
								End If
							End If
						End If

						If Len(sTBResultCode) >= 1 Then
							' Get the Overlapped Booking check code.
							sCode = Mid(sTBResultCode, Len(sTBResultCode), 1)
							If sCode = "1" Then
								If Len(sTBErrorMsg) > 0 Then
									sTBErrorMsg = sTBErrorMsg & vbCrLf
								End If
								sTBErrorMsg = sTBErrorMsg & "The delegate is already booked on a course that overlaps with this course. Unable to make the booking."
							Else
								If sCode = "2" Then
									If Len(sTBWarningMsg) > 0 Then
										sTBWarningMsg = sTBWarningMsg & vbCrLf
									End If
									sTBWarningMsg = sTBWarningMsg & "The delegate is already booked on a course that overlaps with this course. Unable to make the booking."
								End If
							End If
						End If
					End If
				End If
			End If

			If Len(sTBErrorMsg) > 0 Then
				' Training Booking validation failure.	
				sErrorMsg = sTBErrorMsg
				sAction = "SAVEERROR"
			Else
				If Len(sTBWarningMsg) > 0 Then
					sErrorMsg = sTBWarningMsg
					sAction = sReaction
					fWarning = True
				Else
					' Check if we're inserting or updating.
					If lngRecordID = 0 Then
						' Inserting.
						Try

							Dim prmRecordID As New SqlParameter("piNewRecordID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
							Dim prmErrorMessage As New SqlParameter("errorMessage", SqlDbType.NVarChar, -1) With {.Direction = ParameterDirection.Output}

							objDataAccess.ExecuteSP("spASRIntInsertNewRecord" _
								, prmRecordID _
								, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = lngTableID} _
								, New SqlParameter("FromRecordID", SqlDbType.Int) With {.Value = lngOriginalRecordID} _
								, New SqlParameter("psInsertDef", SqlDbType.VarChar, -1) With {.Value = sInsertUpdateDef} _
								, prmErrorMessage)

							If prmErrorMessage.Value.ToString().Length > 0 Then
								sAction = "SAVEERROR"
								sErrorMsg = prmErrorMessage.Value.ToString()
							Else

								lngRecordID = CInt(prmRecordID.Value)

								' This was a copied record - ensure that OLE columns are also copied
								If lngOriginalRecordID > 0 Then

									objDataAccess.ExecuteSP("spasrIntCopyRecordPostSave" _
										, New SqlParameter("tableID", SqlDbType.Int) With {.Value = lngTableID} _
										, New SqlParameter("FromRecordID", SqlDbType.Int) With {.Value = lngOriginalRecordID} _
										, New SqlParameter("ToRecordID", SqlDbType.Int) With {.Value = lngRecordID})
								End If

								If Len(sReaction) > 0 Then
									sAction = sReaction
								Else
									sAction = "LOAD"
								End If

								objDataAccess.ExecuteSP("spASREmailImmediate", New SqlParameter("@Username", SqlDbType.VarChar, 255) With {.Value = _userName})
							End If

						Catch ex As SqlException
							If ex.Number.Equals(50000) Then
								If ex.Message.IndexOf("The transaction ended in the trigger", StringComparison.Ordinal) > 0 Then
									sErrorMsg = Trim(Mid(ex.Message, 1, (InStr(ex.Message, "The transaction ended in the trigger")) - 1))
								ElseIf ex.Message.IndexOf("Invalid object name", StringComparison.Ordinal) > 0 Then
									sErrorMsg = Trim(Mid(ex.Message, 1, (InStr(ex.Message, "Invalid object name")) - 1))
								Else
									sErrorMsg = sErrorMsg & FormatError(ex.Message)
								End If
							Else
								sErrorMsg = sErrorMsg & FormatError(ex.Message)
							End If

							fOK = False

							Dim sRecDescExists = ""
							If Mid(sErrorMsg, 3, 5) <> "-----" Then
								sRecDescExists = vbCrLf
							End If

							sErrorMsg = "The new record could not be created." & sRecDescExists & sErrorMsg
							sAction = "SAVEERROR"
						End Try
					Else
						' Updating.
						Try
							Dim prmResult As New SqlParameter("piResult", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
							Dim prmErrorMessage As New SqlParameter("errorMessage", SqlDbType.NVarChar, -1) With {.Direction = ParameterDirection.Output}

							objDataAccess.ExecuteSP("spASRIntUpdateRecord" _
								, prmResult _
								, New SqlParameter("psUpdateDef", SqlDbType.VarChar, -1) With {.Value = sInsertUpdateDef} _
								, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = NullSafeInteger(CleanNumeric(lngTableID))} _
								, New SqlParameter("psRealSource", SqlDbType.VarChar, 255) With {.Value = sRealSource} _
								, New SqlParameter("piID", SqlDbType.Int) With {.Value = CleanNumeric(lngRecordID)} _
								, New SqlParameter("piTimestamp", SqlDbType.Int) With {.Value = CleanNumeric(iTimestamp)} _
								, prmErrorMessage)

							If prmErrorMessage.Value.ToString().Length > 0 Then
								sAction = "SAVEERROR"
								sErrorMsg = prmErrorMessage.Value.ToString()
							Else

								Select Case CType(prmResult.Value, SaveResponse)
									Case SaveResponse.NoLongerInView
										sErrorMsg = "The record has been amended by another user and will be refreshed."
									Case SaveResponse.AmendedByAnotherUser
										sErrorMsg = "The record has been amended by another user and will be refreshed."
									Case SaveResponse.DeletedByAnotherUser
										sErrorMsg = "The record has been deleted by another user."
								End Select

								If Len(sReaction) > 0 Then
									sAction = sReaction
								Else
									sAction = "LOAD"
								End If

								objDataAccess.ExecuteSP("spASREmailImmediate", _
										New SqlParameter("@Username", SqlDbType.VarChar, 255) With {.Value = _userName})

							End If

						Catch ex As SqlException
							If ex.Number.Equals(50000) Then
								If InStr(ex.Message, "The transaction ended in the trigger") > 0 Then
									sErrorMsg = sErrorMsg & FormatError(Trim(Mid(ex.Message, 1, (InStr(ex.Message, "The transaction ended in the trigger")) - 1)))
								ElseIf InStr(ex.Message, "Invalid object name") > 0 Then
									sErrorMsg = sErrorMsg & FormatError(Trim(Mid(ex.Message, 1, (InStr(ex.Message, "Invalid object name")) - 1)))
								Else
									sErrorMsg = sErrorMsg & FormatError(ex.Message)
								End If
							Else
								sErrorMsg = sErrorMsg & FormatError(ex.Message)
							End If

							fOK = False

							Dim sRecDescExists = ""
							If Mid(sErrorMsg, 3, 5) <> "-----" Then
								sRecDescExists = vbCrLf
							End If

							sErrorMsg = "The record could not be updated." & sRecDescExists & sErrorMsg
							sAction = "SAVEERROR"

						End Try

					End If
				End If
			End If

			result.RecordID = lngRecordID
			result.Action = sAction
			result.Message = sErrorMsg
			result.TBResultCode = sTBResultCode
			result.CourseOverbooked = sCourseOverbooked
			result.Warning = fWarning
			result.OK = fOK

			Return result

		End Function


	End Class
End Namespace