Imports System.Web.Mvc
Imports System.IO

Namespace Controllers

	Public Class HomeController
		Inherits Controller

		'
		' GET: /Home
		Function Main() As ActionResult
			Return View()
		End Function

		Function Find() As ActionResult
			Return View()
		End Function

		<HttpPost()>
		Function Default_Submit()

			' Save the required table/view and screen IDs in session variables.
			Session("action") = Request.Form("txtAction")
			Session("tableID") = Request.Form("txtGotoTableID")
			Session("viewID") = Request.Form("txtGotoViewID")
			Session("screenID") = Request.Form("txtGotoScreenID")
			Session("orderID") = Request.Form("txtGotoOrderID")
			Session("recordID") = Request.Form("txtGotoRecordID")
			Session("parentTableID") = Request.Form("txtGotoParentTableID")
			Session("parentRecordID") = Request.Form("txtGotoParentRecordID")
			Session("realSource") = Request.Form("txtGotoRealSource")
			Session("filterDef") = Request.Form("txtGotoFilterDef")
			Session("filterSQL") = Request.Form("txtGotoFilterSQL")
			Session("lineage") = Request.Form("txtGotoLineage")
			Session("defseltype") = Request.Form("txtGotoDefSelType")
			Session("utilID") = Request.Form("txtGotoUtilID")
			Session("locateValue") = Request.Form("txtGotoLocateValue")
			Session("firstRecPos") = Request.Form("txtGotoFirstRecPos")
			Session("currentRecCount") = Request.Form("txtGotoCurrentRecCount")
			Session("fromMenu") = Request.Form("txtGotoFromMenu")
			Session("reset") = Request.Form("txtReset")

			Session("reloadMenu") = Request.Form("txtReloadMenu")

			Session("StandardReport_Type") = Request.Form("txtStandardReportType")
			Session("optionRecordID") = "0"
			Session("optionAction") = ""

			' Go to the requested page.
			Return RedirectToAction(Request.Form("txtGotoPage").Replace(".asp", ""))

		End Function

		Function DefSel() As ActionResult
			Return View()
		End Function

		<HttpPost()>
		Function DefSel(value As FormCollection)
			Return View()
		End Function

		<HttpPost()>
		Function DefSel_Submit(value As FormCollection)

			' Set some session variables used by all the util pages
			Session("utiltype") = Request.Form("utiltype")
			Session("utilid") = Request.Form("utilid")
			Session("utilname") = Request.Form("utilname")
			Session("action") = Request.Form("action")
			Session("utiltableid") = Request.Form("txtTableID")

			' Now examine what we are doing and redirect as appropriate
			If (Session("action") = "new") Or _
			 (Session("action") = "edit") Or _
			 (Session("action") = "view") Or _
			 (Session("action") = "copy") Then
				Select Case Session("utiltype")
					Case 1 ' CROSS TABS
						Return RedirectToAction("util_def_crosstabs")
					Case 2 ' CUSTOM REPORTS
						Return RedirectToAction("util_def_customreports")
					Case 9 ' MAIL MERGE
						Return RedirectToAction("util_def_mailmerge")
					Case 10 ' PICKLISTS
						Return RedirectToAction("util_def_picklist")
					Case 11 ' FILTERS
						Return RedirectToAction("util_def_expression")
					Case 12 ' CALCULATIONS
						Return RedirectToAction("util_def_expression")
					Case 17 ' CALENDAR REPORTS
						Return RedirectToAction("util_def_calendarreport")
				End Select

			ElseIf Session("action") = "delete" Then
				Select Case Session("utiltype")
					Case 1	' CROSS TABS
						Session("reaction") = "CROSSTABS"
					Case 2	' CUSTOM REPORTS
						Session("reaction") = "CUSTOMREPORTS"
					Case 9	' MAIL MERGE
						Session("reaction") = "MAILMERGE"
					Case 10 ' PICKLISTS
						Session("reaction") = "PICKLISTS"
					Case 11 ' FILTERS
						Session("reaction") = "FILTERS"
					Case 12 ' CALCULATIONS
						Session("reaction") = "CALCULATIONS"
					Case 17 ' CALENDAR REPORTS
						Session("reaction") = "CALENDARREPORTS"
				End Select
				Return RedirectToAction("checkforusage")
			End If

		End Function

		Function DefSelProperties() As ActionResult
			Return View()
		End Function

		Function Util_Def_CustomReports() As ActionResult
			Return View()
		End Function

		Function CheckForUsage() As ActionResult
			Return View()
		End Function

		Function Data() As ActionResult
			Return View()
		End Function

		Function Data_Submit() As ActionResult

			On Error Resume Next

			Const DEADLOCK_ERRORNUMBER = -2147467259
			Const DEADLOCK_MESSAGESTART = "YOUR TRANSACTION (PROCESS ID #"
			Const DEADLOCK_MESSAGEEND = ") WAS DEADLOCKED WITH ANOTHER PROCESS AND HAS BEEN CHOSEN AS THE DEADLOCK VICTIM. RERUN YOUR TRANSACTION."
			Const DEADLOCK2_MESSAGESTART = "TRANSACTION (PROCESS ID "
			Const DEADLOCK2_MESSAGEEND = ") WAS DEADLOCKED ON "
			Const SQLMAILNOTSTARTEDMESSAGE = "SQL MAIL SESSION IS NOT STARTED."

			Dim iRETRIES = 5
			Dim iRetryCount = 0
			Dim sErrorMsg = "", sErrMsg = ""
			Dim fWarning = False
			Dim fOk = False
			Dim fTBOverride = False

			' Read the information from the calling form.
			Dim sRealSource = Request.Form("txtRealSource")
			Dim lngTableID = Request.Form("txtCurrentTableID")
			Dim lngScreenID = Request.Form("txtCurrentScreenID")
			Dim lngViewID = Request.Form("txtCurrentViewID")
			Dim lngRecordID = Request.Form("txtRecordID")
			Dim sAction = Request.Form("txtAction")
			Dim sReaction = Request.Form("txtReaction")
			Dim sInsertUpdateDef = Request.Form("txtInsertUpdateDef")
			Dim iTimestamp = Request.Form("txtTimestamp")
			Dim iTBEmployeeRecordID = Request.Form("txtTBEmployeeRecordID")
			Dim iTBCourseRecordID = Request.Form("txtTBCourseRecordID")
			Dim sTBBookingStatusValue = Request.Form("txtTBBookingStatusValue")
			Dim fUserChoice = Request.Form("txtUserChoice")

			If Request.Form("txtTBOverride") = "" Then
				fTBOverride = False
			Else
				fTBOverride = CBool(Request.Form("txtTBOverride"))
			End If

			If sAction = "SAVE" Then
				Dim sTBErrorMsg = ""
				Dim sTBWarningMsg = ""
				Dim iTBResultCode = 0
				Dim sCode = ""

				If (Not fTBOverride) And (CLng(lngTableID) = CLng(Session("TB_TBTableID"))) Then
					' Training Booking check.
					Dim cmdTBCheck = Server.CreateObject("ADODB.Command")
					cmdTBCheck.CommandText = "sp_ASRIntValidateTrainingBooking"
					cmdTBCheck.CommandType = 4	' Stored procedure
					cmdTBCheck.ActiveConnection = Session("databaseConnection")

					Dim prmResult = cmdTBCheck.CreateParameter("resultCode", 3, 2)	' 3=integer, 2=output
					cmdTBCheck.Parameters.Append(prmResult)

					Dim prmTBEmployeeRecordID = cmdTBCheck.CreateParameter("empRecID", 3, 1)	'3=integer, 1=input
					cmdTBCheck.Parameters.Append(prmTBEmployeeRecordID)
					prmTBEmployeeRecordID.value = CleanNumeric(iTBEmployeeRecordID)

					Dim prmTBCourseRecordID = cmdTBCheck.CreateParameter("courseRecID", 3, 1) '3=integer, 1=input
					cmdTBCheck.Parameters.Append(prmTBCourseRecordID)
					prmTBCourseRecordID.value = CleanNumeric(iTBCourseRecordID)

					Dim prmTBStatus = cmdTBCheck.CreateParameter("tbStatus", 200, 1, 8000) '200=varchar, 1=input, 8000=size
					cmdTBCheck.Parameters.Append(prmTBStatus)
					prmTBStatus.value = sTBBookingStatusValue

					Dim prmTBRecordID = cmdTBCheck.CreateParameter("tbRecID", 3, 1)	'3=integer, 1=input
					cmdTBCheck.Parameters.Append(prmTBRecordID)
					prmTBRecordID.value = CleanNumeric(lngRecordID)

					Err.Clear()
					cmdTBCheck.Execute()
					If (Err.Number <> 0) Then
						sErrorMsg = "Error validating training booking." & vbCrLf & FormatError(Err.Description)
					End If

					If Len(sErrorMsg) = 0 Then
						iTBResultCode = cmdTBCheck.Parameters("resultCode").Value
					End If
					cmdTBCheck = Nothing

					If Len(sErrorMsg) = 0 Then
						If iTBResultCode > 0 Then
							Dim sTBResultCode = CStr(iTBResultCode)
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

							' The required stored procedure exists, so run it.
							Dim cmdInsertRecord = Server.CreateObject("ADODB.Command")
							cmdInsertRecord.CommandText = "spASRIntInsertNewRecord"
							cmdInsertRecord.CommandType = 4 ' Stored procedure
							cmdInsertRecord.CommandTimeout = 180
							cmdInsertRecord.ActiveConnection = Session("databaseConnection")

							Dim prmNewID = cmdInsertRecord.CreateParameter("newID", 3, 2)
							cmdInsertRecord.Parameters.Append(prmNewID)

							Dim prmInsertSQL = cmdInsertRecord.CreateParameter("insertSQL", 201, 1, 2147483646)
							cmdInsertRecord.Parameters.Append(prmInsertSQL)
							prmInsertSQL.value = sInsertUpdateDef

							Dim fDeadlock = True
							Do While fDeadlock
								fDeadlock = False

								cmdInsertRecord.ActiveConnection.Errors.Clear()

								' Run the insert stored procedure.
								cmdInsertRecord.Execute()

								If cmdInsertRecord.ActiveConnection.Errors.Count > 0 Then
									For iLoop = 1 To cmdInsertRecord.ActiveConnection.Errors.Count
										sErrMsg = FormatError(cmdInsertRecord.ActiveConnection.Errors.Item(iLoop - 1).Description)

										If (cmdInsertRecord.ActiveConnection.Errors.Item(iLoop - 1).Number = DEADLOCK_ERRORNUMBER) And _
										 (((UCase(Left(sErrMsg, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
										(UCase(Right(sErrMsg, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
										 ((UCase(Left(sErrMsg, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
										  (InStr(UCase(sErrMsg), DEADLOCK2_MESSAGEEND) > 0))) Then
											' The error is for a deadlock.
											' Sorry about having to use the err.description to trap the error but the err.number
											' is not specific and MSDN suggests using the err.description.
											If (iRetryCount < iRETRIES) And (cmdInsertRecord.ActiveConnection.Errors.Count = 1) Then
												iRetryCount = iRetryCount + 1
												fDeadlock = True
											Else
												If Len(sErrorMsg) > 0 Then
													sErrorMsg = sErrorMsg & vbCrLf
												End If
												sErrorMsg = sErrorMsg & "Another user is deadlocking the database. Try saving again."
												fOk = False
											End If
										ElseIf UCase(cmdInsertRecord.ActiveConnection.Errors.Item(iLoop - 1).Description) = SQLMAILNOTSTARTEDMESSAGE Then
											'"SQL Mail session is not started."
											'Ignore this error
											'ElseIf (cmdInsertRecord.ActiveConnection.Errors.Item(iloop - 1).Number = XP_SENDMAIL_ERRORNUMBER) And _
											'	(UCase(Left(cmdInsertRecord.ActiveConnection.Errors.Item(iloop - 1).Description, Len(XP_SENDMAIL_MESSAGE))) = XP_SENDMAIL_MESSAGE) Then
											'"EXECUTE permission denied on object 'xp_sendmail'"
											'Ignore this error
										ElseIf cmdInsertRecord.ActiveConnection.Errors.Item(iLoop - 1).nativeerror = 3609 Then
											' Ignore the follow on message that says "The transaction ended in the trigger."
										Else
											'NHRD 18082011 HRPRO-1572 Removed extra carriage return for this error msg
											'sErrorMsg = sErrorMsg & vbcrlf & _
											sErrorMsg = sErrorMsg & _
											 FormatError(cmdInsertRecord.ActiveConnection.Errors.Item(iLoop - 1).Description)
											fOk = False
										End If
									Next

									cmdInsertRecord.ActiveConnection.Errors.Clear()

									If Not fOk Then
										'JPD 20110705 HRPRO-1572
										' Now get validation failure message prefixed woth <record description> and <line of hyphens>.
										' Only add extra carriage return if required (ie. if there is a record description).
										Dim sRecDescExists = ""
										If Mid(sErrorMsg, 3, 5) <> "-----" Then
											sRecDescExists = vbCrLf
										End If

										sErrorMsg = "The new record could not be created." & sRecDescExists & sErrorMsg
										sAction = "SAVEERROR"
									End If
								Else
									lngRecordID = cmdInsertRecord.Parameters("newID").Value

									If Len(sReaction) > 0 Then
										sAction = sReaction
									Else
										sAction = "LOAD"
									End If
								End If
							Loop
							cmdInsertRecord = Nothing


							'MH20001017 Immediate email stuff to go in v1.9.0
							Dim cmdInsertRecord2 = Server.CreateObject("ADODB.Command")
							cmdInsertRecord2.CommandText = "spASREmailImmediate"
							cmdInsertRecord2.CommandType = 4	' Stored procedure
							cmdInsertRecord2.CommandTimeout = 180
							cmdInsertRecord2.ActiveConnection = Session("databaseConnection")

							Dim prmInsertSQL2 = cmdInsertRecord2.CreateParameter("Username", 200, 1, 255)	' 200=varchar, 1=input, 255=size
							cmdInsertRecord2.Parameters.Append(prmInsertSQL2)
							prmInsertSQL2.value = Session("Username")

							Err.Clear()
							cmdInsertRecord2.Execute()
							cmdInsertRecord2 = Nothing
						Else
							' Updating.

							' The required stored procedure exists, so run it.
							Dim cmdUpdateRecord = Server.CreateObject("ADODB.Command")
							cmdUpdateRecord.CommandText = "spASRIntUpdateRecord"
							cmdUpdateRecord.CommandType = 4 ' Stored procedure
							cmdUpdateRecord.CommandTimeout = 180
							cmdUpdateRecord.ActiveConnection = Session("databaseConnection")

							Dim prmResultCode = cmdUpdateRecord.CreateParameter("resultCode", 3, 2)	' 3=integer, 2=output
							cmdUpdateRecord.Parameters.Append(prmResultCode)

							Dim prmUpdateSQL = cmdUpdateRecord.CreateParameter("updateSQL", 201, 1, 2147483646)
							cmdUpdateRecord.Parameters.Append(prmUpdateSQL)
							prmUpdateSQL.value = sInsertUpdateDef

							Dim prmTableID = cmdUpdateRecord.CreateParameter("tableID", 3, 1)
							cmdUpdateRecord.Parameters.Append(prmTableID)
							prmTableID.value = CLng(CleanNumeric(lngTableID))

							Dim prmRealSource = cmdUpdateRecord.CreateParameter("realSource", 200, 1, 255)
							cmdUpdateRecord.Parameters.Append(prmRealSource)
							prmRealSource.value = sRealSource

							Dim prmID = cmdUpdateRecord.CreateParameter("id", 3, 1) ' 3=integer, 1=input
							cmdUpdateRecord.Parameters.Append(prmID)
							prmID.value = CleanNumeric(lngRecordID)

							Dim prmTimestamp = cmdUpdateRecord.CreateParameter("timestamp", 3, 1)	' 3=integer, 1=input
							cmdUpdateRecord.Parameters.Append(prmTimestamp)
							prmTimestamp.value = CleanNumeric(iTimestamp)

							Dim fDeadlock = True
							Do While fDeadlock
								fDeadlock = False

								cmdUpdateRecord.ActiveConnection.Errors.Clear()

								' Run the update stored procedure.
								cmdUpdateRecord.Execute()

								If cmdUpdateRecord.ActiveConnection.Errors.Count > 0 Then
									For iLoop = 1 To cmdUpdateRecord.ActiveConnection.Errors.Count
										sErrMsg = FormatError(cmdUpdateRecord.ActiveConnection.Errors.Item(iLoop - 1).Description)

										If (cmdUpdateRecord.ActiveConnection.Errors.Item(iLoop - 1).Number = DEADLOCK_ERRORNUMBER) And _
										 (((UCase(Left(sErrMsg, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
										  (UCase(Right(sErrMsg, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
										 ((UCase(Left(sErrMsg, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
										 (InStr(UCase(sErrMsg), DEADLOCK2_MESSAGEEND) > 0))) Then
											' The error is for a deadlock.
											' Sorry about having to use the err.description to trap the error but the err.number
											' is not specific and MSDN suggests using the err.description.
											If (iRetryCount < iRETRIES) And (cmdUpdateRecord.ActiveConnection.Errors.Count = 1) Then
												iRetryCount = iRetryCount + 1
												fDeadlock = True
											Else
												If Len(sErrorMsg) > 0 Then
													sErrorMsg = sErrorMsg & vbCrLf
												End If
												sErrorMsg = sErrorMsg & "Another user is deadlocking the database. Try saving again."
												fOk = False
											End If
										ElseIf UCase(cmdUpdateRecord.ActiveConnection.Errors.Item(iLoop - 1).Description) = SQLMAILNOTSTARTEDMESSAGE Then
											'"SQL Mail session is not started."
											'Ignore this error
										ElseIf cmdUpdateRecord.ActiveConnection.Errors.Item(iLoop - 1).nativeerror = 3609 Then
											' Ignore the follow on message that says "The transaction ended in the trigger."
										Else
											sErrorMsg = sErrorMsg & vbCrLf & _
											 FormatError(cmdUpdateRecord.ActiveConnection.Errors.Item(iLoop - 1).Description)
											fOk = False
										End If
									Next

									cmdUpdateRecord.ActiveConnection.Errors.Clear()

									If Not fOk Then
										'JPD 20110705 HRPRO-1572
										' Now get validation failure message prefixed with <record description> and <line of hyphens>.
										' Only add extra carriage return if required (ie. if there is a record description).
										Dim sRecDescExists = ""
										If Mid(sErrorMsg, 3, 5) <> "-----" Then
											sRecDescExists = vbCrLf
										End If

										sErrorMsg = "The record could not be updated." & sRecDescExists & sErrorMsg
										sAction = "SAVEERROR"
									End If
								Else
									Select Case cmdUpdateRecord.Parameters("resultCode").Value
										Case 1 ' Record changed by another user, and still in the current table/view.
											sErrorMsg = "The record has been amended by another user and will be refreshed."
										Case 2 ' Record changed by another user, and is no longer in the current table/view.
											sErrorMsg = "The record has been amended by another user and is no longer in the current view."
										Case 3 ' Record deleted by another user.
											sErrorMsg = "The record has been deleted by another user."
									End Select

									If Len(sReaction) > 0 Then
										sAction = sReaction
									Else
										sAction = "LOAD"
									End If
								End If
							Loop
							cmdUpdateRecord = Nothing

							'MH20001017 Immediate email stuff to go in v1.9.0
							cmdUpdateRecord = Server.CreateObject("ADODB.Command")
							cmdUpdateRecord.CommandText = "spASREmailImmediate"
							cmdUpdateRecord.CommandType = 4 ' Stored procedure
							cmdUpdateRecord.CommandTimeout = 180
							cmdUpdateRecord.ActiveConnection = Session("databaseConnection")

							prmUpdateSQL = cmdUpdateRecord.CreateParameter("Username", 200, 1, 255)	' 200=varchar, 1=input, 255=size
							cmdUpdateRecord.Parameters.Append(prmUpdateSQL)
							prmUpdateSQL.value = Session("Username")

							Err.Clear()
							cmdUpdateRecord.Execute()
							cmdUpdateRecord = Nothing
						End If
					End If
				End If
			ElseIf sAction = "DELETE" Then
				' Deleting.

				' The required stored procedure exists, so run it.
				Dim cmdDeleteRecord = Server.CreateObject("ADODB.Command")
				cmdDeleteRecord.CommandText = "sp_ASRDeleteRecord"
				cmdDeleteRecord.CommandType = 4 ' Stored procedure
				cmdDeleteRecord.ActiveConnection = Session("databaseConnection")

				Dim prmResultCode = cmdDeleteRecord.CreateParameter("resultCode", 3, 2)
				cmdDeleteRecord.Parameters.Append(prmResultCode)

				Dim prmTableID = cmdDeleteRecord.CreateParameter("tableID", 3, 1)
				cmdDeleteRecord.Parameters.Append(prmTableID)
				prmTableID.value = CLng(CleanNumeric(lngTableID))

				Dim prmRealSource = cmdDeleteRecord.CreateParameter("realSource", 200, 1, 8000)
				cmdDeleteRecord.Parameters.Append(prmRealSource)
				prmRealSource.value = CleanString(sRealSource)

				Dim prmID = cmdDeleteRecord.CreateParameter("id", 3, 1)
				cmdDeleteRecord.Parameters.Append(prmID)
				prmID.value = CleanNumeric(lngRecordID)

				Dim fDeadlock = True
				Do While fDeadlock
					fDeadlock = False

					cmdDeleteRecord.ActiveConnection.Errors.Clear()

					' Run the delete stored procedure.
					cmdDeleteRecord.Execute()

					If cmdDeleteRecord.ActiveConnection.Errors.Count > 0 Then
						For iLoop = 1 To cmdDeleteRecord.ActiveConnection.Errors.Count
							sErrMsg = FormatError(cmdDeleteRecord.ActiveConnection.Errors.Item(iLoop - 1).Description)

							If (cmdDeleteRecord.ActiveConnection.Errors.Item(iLoop - 1).Number = DEADLOCK_ERRORNUMBER) And _
							 (((UCase(Left(sErrMsg, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
							  (UCase(Right(sErrMsg, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
							 ((UCase(Left(sErrMsg, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
							 (InStr(UCase(sErrMsg), DEADLOCK2_MESSAGEEND) > 0))) Then

								' The error is for a deadlock.
								' Sorry about having to use the err.description to trap the error but the err.number
								' is not specific and MSDN suggests using the err.description.
								If (iRetryCount < iRETRIES) And (cmdDeleteRecord.ActiveConnection.Errors.Count = 1) Then
									iRetryCount = iRetryCount + 1
									fDeadlock = True
								Else
									If Len(sErrorMsg) > 0 Then
										sErrorMsg = sErrorMsg & vbCrLf
									End If
									sErrorMsg = sErrorMsg & "Another user is deadlocking the database. Try saving again."
									fOk = False
								End If
							ElseIf cmdDeleteRecord.ActiveConnection.Errors.Item(iLoop - 1).nativeerror = 3609 Then
								' Ignore the follow on message that says "The transaction ended in the trigger."
							Else
								sErrorMsg = sErrorMsg & vbCrLf & _
								 FormatError(cmdDeleteRecord.ActiveConnection.Errors.Item(iLoop - 1).Description)
								fOk = False
							End If
						Next

						cmdDeleteRecord.ActiveConnection.Errors.Clear()

						If Not fOk Then
							sErrorMsg = "The record could not be deleted." & vbCrLf & sErrorMsg
							sAction = "SAVEERROR"
						End If
					Else
						Select Case cmdDeleteRecord.Parameters("resultCode").Value
							Case 2 ' Record changed by another user, and is no longer in the current table/view.
								sErrorMsg = "The record has been amended by another user and is no longer in the current view."
						End Select

						lngRecordID = 0

						If Len(sReaction) > 0 Then
							sAction = sReaction
						Else
							sAction = "LOAD"
						End If
					End If
				Loop
				cmdDeleteRecord = Nothing

				'MH20100609
				Dim cmdInsertRecord = Server.CreateObject("ADODB.Command")
				cmdInsertRecord.CommandText = "spASREmailImmediate"
				cmdInsertRecord.CommandType = 4 ' Stored procedure
				cmdInsertRecord.CommandTimeout = 180
				cmdInsertRecord.ActiveConnection = Session("databaseConnection")

				Dim prmInsertSQL = cmdInsertRecord.CreateParameter("Username", 200, 1, 255)	' 200=varchar, 1=input, 255=size
				cmdInsertRecord.Parameters.Append(prmInsertSQL)
				prmInsertSQL.value = Session("Username")

				Err.Clear()
				cmdInsertRecord.Execute()
				cmdInsertRecord = Nothing

			ElseIf sAction = "CANCELCOURSE" Then
				' Check number of bookings made.
				Dim cmdCancelCourse = Server.CreateObject("ADODB.Command")
				cmdCancelCourse.CommandText = "sp_ASRIntCancelCourse"
				cmdCancelCourse.CommandType = 4 ' Stored procedure
				cmdCancelCourse.ActiveConnection = Session("databaseConnection")

				Dim prmNumberOfBookings = cmdCancelCourse.CreateParameter("numberOfBookings", 3, 2)	' 3=integer, 2=output
				cmdCancelCourse.Parameters.Append(prmNumberOfBookings)

				Dim prmCourseRecordID = cmdCancelCourse.CreateParameter("courseRecordID", 3, 1)
				cmdCancelCourse.Parameters.Append(prmCourseRecordID)
				prmCourseRecordID.value = CleanNumeric(lngRecordID)

				Dim prmTBTableID = cmdCancelCourse.CreateParameter("tbTableID", 3, 1)
				cmdCancelCourse.Parameters.Append(prmTBTableID)
				prmTBTableID.value = CleanNumeric(Session("TB_TBTableID"))

				Dim prmCourseTableID = cmdCancelCourse.CreateParameter("courseTableID", 3, 1)
				cmdCancelCourse.Parameters.Append(prmCourseTableID)
				prmCourseTableID.value = CleanNumeric(Session("TB_CourseTableID"))

				Dim prmTrainBookStatusColumnID = cmdCancelCourse.CreateParameter("tbStatusColumnID", 3, 1)
				cmdCancelCourse.Parameters.Append(prmTrainBookStatusColumnID)
				prmTrainBookStatusColumnID.value = CleanNumeric(Session("TB_TBStatusColumnID"))

				Dim prmCourseRealSource = cmdCancelCourse.CreateParameter("courseRealSource", 200, 1, 8000)
				cmdCancelCourse.Parameters.Append(prmCourseRealSource)
				prmCourseRealSource.value = sRealSource

				Dim prmErrorMessage = cmdCancelCourse.CreateParameter("errorMessage", 200, 2, 8000)
				cmdCancelCourse.Parameters.Append(prmErrorMessage)

				Dim prmCourseTitle = cmdCancelCourse.CreateParameter("courseTitle", 200, 2, 8000)
				cmdCancelCourse.Parameters.Append(prmCourseTitle)

				Err.Clear()
				cmdCancelCourse.Execute()
				If (Err.Number <> 0) Then
					sErrorMsg = "Error cancelling the course." & vbCrLf & FormatError(Err.Description)
					sAction = "SAVEERROR"
				Else
					sAction = "CANCELCOURSE_1"
					Session("numberOfBookings") = cmdCancelCourse.Parameters("numberOfBookings").Value
					Session("tbErrorMessage") = cmdCancelCourse.Parameters("errorMessage").Value
					Session("tbCourseTitle") = cmdCancelCourse.Parameters("courseTitle").Value
				End If

				cmdCancelCourse = Nothing
			ElseIf sAction = "CANCELCOURSE_2" Then
				Dim cmdCancelCourse = Server.CreateObject("ADODB.Command")
				cmdCancelCourse.CommandText = "sp_ASRIntCancelCoursePart2"
				cmdCancelCourse.CommandType = 4 ' Stored procedure
				cmdCancelCourse.ActiveConnection = Session("databaseConnection")

				Dim prmEmployeeTableID = cmdCancelCourse.CreateParameter("employeeTableID", 3, 1)
				cmdCancelCourse.Parameters.Append(prmEmployeeTableID)
				prmEmployeeTableID.value = CleanNumeric(Session("TB_EmpTableID"))

				Dim prmCourseTableID = cmdCancelCourse.CreateParameter("courseTableID", 3, 1)
				cmdCancelCourse.Parameters.Append(prmCourseTableID)
				prmCourseTableID.value = CleanNumeric(Session("TB_CourseTableID"))

				Dim prmCourseRealSource = cmdCancelCourse.CreateParameter("courseRealSource", 200, 1, 8000)
				cmdCancelCourse.Parameters.Append(prmCourseRealSource)
				prmCourseRealSource.value = sRealSource

				Dim prmCourseRecordID = cmdCancelCourse.CreateParameter("courseRecordID", 3, 1)
				cmdCancelCourse.Parameters.Append(prmCourseRecordID)
				prmCourseRecordID.value = CleanNumeric(lngRecordID)

				Dim prmNewCourseRecordID = cmdCancelCourse.CreateParameter("newCourseRecordID", 3, 1)
				cmdCancelCourse.Parameters.Append(prmNewCourseRecordID)
				prmNewCourseRecordID.value = CleanNumeric(iTBCourseRecordID)

				Dim prmCourseCancelDateColumnID = cmdCancelCourse.CreateParameter("courseCancelDateColumnID", 3, 1)
				cmdCancelCourse.Parameters.Append(prmCourseCancelDateColumnID)
				prmCourseCancelDateColumnID.value = CleanNumeric(Session("TB_CourseCancelDateColumnID"))

				Dim prmCourseTitle = cmdCancelCourse.CreateParameter("courseTitle", 200, 1, 8000)
				cmdCancelCourse.Parameters.Append(prmCourseTitle)
				prmCourseTitle.value = Session("tbCourseTitle")

				Dim prmTBTableID = cmdCancelCourse.CreateParameter("tbTableID", 3, 1)
				cmdCancelCourse.Parameters.Append(prmTBTableID)
				prmTBTableID.value = CleanNumeric(Session("TB_TBTableID"))

				Dim prmTBTableInsert = cmdCancelCourse.CreateParameter("tbTableInsert", 11, 1)	' 11=boolean, 1=input
				cmdCancelCourse.Parameters.Append(prmTBTableInsert)
				prmTBTableInsert.value = CleanBoolean(Session("TB_TBTableInsert"))

				Dim prmTrainBookStatusColumnID = cmdCancelCourse.CreateParameter("tbStatusColumnID", 3, 1)
				cmdCancelCourse.Parameters.Append(prmTrainBookStatusColumnID)
				prmTrainBookStatusColumnID.value = CleanNumeric(Session("TB_TBStatusColumnID"))

				Dim prmTrainBookCancelDateColumnID = cmdCancelCourse.CreateParameter("tbCancelDateColumnID", 3, 1)
				cmdCancelCourse.Parameters.Append(prmTrainBookCancelDateColumnID)
				prmTrainBookCancelDateColumnID.value = CleanNumeric(Session("TB_TBCancelDateColumnID"))

				Dim prmWLTableID = cmdCancelCourse.CreateParameter("wlTableID", 3, 1)
				cmdCancelCourse.Parameters.Append(prmWLTableID)
				prmWLTableID.value = CleanNumeric(Session("TB_WaitListTableID"))

				Dim prmWLTableInsert = cmdCancelCourse.CreateParameter("wlTableInsert", 11, 1)	' 11=boolean, 1=input
				cmdCancelCourse.Parameters.Append(prmWLTableInsert)
				prmWLTableInsert.value = CleanBoolean(Session("TB_WaitListTableInsert"))

				Dim prmWLCourseTitleColumnID = cmdCancelCourse.CreateParameter("wlCourseTitleColumnID", 3, 1)
				cmdCancelCourse.Parameters.Append(prmWLCourseTitleColumnID)
				prmWLCourseTitleColumnID.value = CleanNumeric(Session("TB_WaitListCourseTitleColumnID"))

				Dim prmWLCourseTitleColumnUpdate = cmdCancelCourse.CreateParameter("wlCourseTitleColumnUpdate", 11, 1)	' 11=boolean, 1=input
				cmdCancelCourse.Parameters.Append(prmWLCourseTitleColumnUpdate)
				prmWLCourseTitleColumnUpdate.value = CleanBoolean(Session("TB_WaitListCourseTitleColumnUpdate"))

				Dim prmCreateWaitListRecords = cmdCancelCourse.CreateParameter("createWaitListRecords", 11, 1) ' 11=boolean, 1=input
				cmdCancelCourse.Parameters.Append(prmCreateWaitListRecords)
				prmCreateWaitListRecords.value = CleanBoolean(Request.Form("txtTBCreateWLRecords"))

				Dim prmErrorMessage = cmdCancelCourse.CreateParameter("errorMessage", 200, 2, 8000)
				cmdCancelCourse.Parameters.Append(prmErrorMessage)

				Err.Clear()
				cmdCancelCourse.Execute()

				If (Err.Number <> 0) Then
					sErrorMsg = "Error cancelling the course." & vbCrLf & FormatError(Err.Description)
					sAction = "SAVEERROR"
				Else
					sErrorMsg = cmdCancelCourse.Parameters("errorMessage").Value

					If Len(sErrorMsg) > 0 Then
						sAction = "SAVEERROR"
					Else
						sAction = "LOAD"
					End If
				End If

				cmdCancelCourse = Nothing

			ElseIf sAction = "CANCELBOOKING" Then
				Dim cmdCancelBooking = Server.CreateObject("ADODB.Command")
				cmdCancelBooking.CommandText = "sp_ASRIntCancelBooking"
				cmdCancelBooking.CommandType = 4	' Stored procedure
				cmdCancelBooking.ActiveConnection = Session("databaseConnection")

				Dim prmTransferBooking = cmdCancelBooking.CreateParameter("transferBooking", 11, 1)	'11=boolean, 1=input
				cmdCancelBooking.Parameters.Append(prmTransferBooking)
				prmTransferBooking.value = CleanBoolean(fUserChoice)

				Dim prmTBRecordID = cmdCancelBooking.CreateParameter("tbRecordID", 3, 1)	'3=integer, 1=input
				cmdCancelBooking.Parameters.Append(prmTBRecordID)
				prmTBRecordID.value = CleanNumeric(lngRecordID)

				Dim prmErrorMessage = cmdCancelBooking.CreateParameter("errorMessage", 200, 2, 8000)	'2=varchar, 2=output, 8000=size
				cmdCancelBooking.Parameters.Append(prmErrorMessage)

				Err.Clear()
				cmdCancelBooking.Execute()
				If (Err.Number <> 0) Then
					sErrorMsg = "Error cancelling the booking." & vbCrLf & FormatError(Err.Description)
					sAction = "SAVEERROR"
				Else
					sErrorMsg = cmdCancelBooking.Parameters("errorMessage").Value

					If Len(sErrorMsg) > 0 Then
						sAction = "SAVEERROR"
					Else
						sAction = "CANCELBOOKING_1"
					End If
				End If

				cmdCancelBooking = Nothing
			End If

			Session("selectSQL") = Request.Form("txtSelectSQL")
			Session("fromDef") = Request.Form("txtFromDef")
			Session("filterSQL") = Request.Form("txtFilterSQL")
			Session("filterDef") = Request.Form("txtFilterDef")
			Session("realSource") = sRealSource
			Session("tableID") = lngTableID
			Session("screenID") = lngScreenID
			Session("viewID") = lngViewID
			Session("recordID") = lngRecordID
			Session("action") = sAction
			Session("reaction") = ""
			Session("warningFlag") = fWarning
			Session("parentTableID") = Request.Form("txtParentTableID")
			Session("parentRecordID") = Request.Form("txtParentRecordID")
			Session("defaultCalcColumns") = Request.Form("txtDefaultCalcCols")
			Session("insertUpdateDef") = sInsertUpdateDef
			Session("errorMessage") = sErrorMsg
			Session("ReportBaseTableID") = Request.Form("txtReportBaseTableID")
			Session("ReportParent1TableID") = Request.Form("txtReportParent1TableID")
			Session("ReportParent2TableID") = Request.Form("txtReportParent2TableID")
			Session("ReportChildTableID") = Request.Form("txtReportChildTableID")
			Session("Param1") = Request.Form("txtParam1")

			'JDM - 24/07/02 - Fault 3917 - Reset year for absence calendar
			Session("stdrpt_AbsenceCalendar_StartYear") = Year(DateTime.Now())

			'JDM - 10/10/02 - Fault 4534 - Reset start month for absence calendar
			Session("stdrpt_AbsenceCalendar_StartMonth") = ""

			'TM - 05/09/02 - Store the event log parameters in session vaiables.
			Session("ELFilterUser") = Request.Form("txtELFilterUser")
			Session("ELFilterType") = Request.Form("txtELFilterType")
			Session("ELFilterStatus") = Request.Form("txtELFilterStatus")
			Session("ELFilterMode") = Request.Form("txtELFilterMode")
			Session("ELOrderColumn") = Request.Form("txtELOrderColumn")
			Session("ELOrderOrder") = Request.Form("txtELOrderOrder")

			Session("ELAction") = Request.Form("txtELAction")

			Session("ELCurrentRecCount") = Request.Form("txtELCurrRecCount")
			If Session("ELCurrentRecCount") < 1 Or Len(Session("ELCurrentRecCount")) < 1 Then
				Session("ELCurrentRecCount") = 0
			End If

			Session("ELFirstRecPos") = Request.Form("txtEL1stRecPos")
			If Session("ELFirstRecPos") < 1 Or Len(Session("ELFirstRecPos")) < 1 Then
				Session("ELFirstRecPos") = 1
			End If

			' Go to the requested page.
			RedirectToAction("Data")

		End Function

		Function Util_RecordSelection() As ActionResult
			Return View()
		End Function

		Function Util_CustomReportChilds() As ActionResult
			Return View()
		End Function

		Function Util_EmailSelection() As ActionResult
			Return View()
		End Function

		Function Util_SortOrderSelection() As ActionResult
			Return View()
		End Function

		Function EventLog() As ActionResult
			Return View()
		End Function


		Function LinksMain() As ActionResult
			If Session("objButtonInfo") Is Nothing Or Session("objHypertextInfo") Is Nothing Or Session("objDropdownInfo") Is Nothing Then
				Return RedirectToAction("Login", "Account")
			End If

			Dim objHypertextInfo As VBA.Collection = Session("objHypertextInfo")
			Dim objButtonInfo As VBA.Collection = Session("objButtonInfo")
			Dim objDropdownInfo As VBA.Collection = Session("objDropdownInfo")

			Dim lstButtonInfo = (From collectionItem As Object In objHypertextInfo Select New navigationLink(collectionItem.ID, collectionItem.DrillDownHidden, collectionItem.LinkType, collectionItem.LinkOrder, collectionItem.Text, collectionItem.Text1, collectionItem.Text2, collectionItem.Prompt, collectionItem.ScreenID, collectionItem.TableID, collectionItem.ViewID, collectionItem.PageTitle, collectionItem.URL, collectionItem.UtilityType, collectionItem.UtilityID, collectionItem.NewWindow, collectionItem.BaseTable, collectionItem.LinkToFind, collectionItem.SingleRecord, collectionItem.PrimarySequence, collectionItem.SecondarySequence, collectionItem.FindPage, collectionItem.EmailAddress, collectionItem.EmailSubject, collectionItem.AppFilePath, collectionItem.AppParameters, collectionItem.DocumentFilePath, collectionItem.DisplayDocumentHyperlink, collectionItem.IsSeparator, collectionItem.Element_Type, collectionItem.SeparatorOrientation, collectionItem.PictureID, collectionItem.Chart_ShowLegend, collectionItem.Chart_Type, collectionItem.Chart_ShowGrid, collectionItem.Chart_StackSeries, collectionItem.Chart_ShowValues, collectionItem.Chart_ViewID, collectionItem.Chart_TableID, collectionItem.Chart_ColumnID, collectionItem.Chart_FilterID, collectionItem.Chart_AggregateType, collectionItem.Chart_ColumnName, collectionItem.Chart_ColumnName_2, collectionItem.UseFormatting, collectionItem.Formatting_DecimalPlaces, collectionItem.Formatting_Use1000Separator, collectionItem.Formatting_Prefix, collectionItem.Formatting_Suffix, collectionItem.UseConditionalFormatting, collectionItem.ConditionalFormatting_Operator_1, collectionItem.ConditionalFormatting_Value_1, collectionItem.ConditionalFormatting_Style_1, collectionItem.ConditionalFormatting_Colour_1, collectionItem.ConditionalFormatting_Operator_2, collectionItem.ConditionalFormatting_Value_2, collectionItem.ConditionalFormatting_Style_2, collectionItem.ConditionalFormatting_Colour_2, collectionItem.ConditionalFormatting_Operator_3, collectionItem.ConditionalFormatting_Value_3, collectionItem.ConditionalFormatting_Style_3, collectionItem.ConditionalFormatting_Colour_3, collectionItem.SeparatorColour, collectionItem.InitialDisplayMode, collectionItem.Chart_TableID_2, collectionItem.Chart_ColumnID_2, collectionItem.Chart_TableID_3, collectionItem.Chart_ColumnID_3, collectionItem.Chart_SortOrderID, collectionItem.Chart_SortDirection, collectionItem.Chart_ColourID, collectionItem.Chart_ShowPercentages)).ToList()
			lstButtonInfo.AddRange(From collectionItem As Object In objButtonInfo Select New navigationLink(collectionItem.ID, collectionItem.DrillDownHidden, collectionItem.LinkType, collectionItem.LinkOrder, collectionItem.Text, collectionItem.Text1, collectionItem.Text2, collectionItem.Prompt, collectionItem.ScreenID, collectionItem.TableID, collectionItem.ViewID, collectionItem.PageTitle, collectionItem.URL, collectionItem.UtilityType, collectionItem.UtilityID, collectionItem.NewWindow, collectionItem.BaseTable, collectionItem.LinkToFind, collectionItem.SingleRecord, collectionItem.PrimarySequence, collectionItem.SecondarySequence, collectionItem.FindPage, collectionItem.EmailAddress, collectionItem.EmailSubject, collectionItem.AppFilePath, collectionItem.AppParameters, collectionItem.DocumentFilePath, collectionItem.DisplayDocumentHyperlink, collectionItem.IsSeparator, collectionItem.Element_Type, collectionItem.SeparatorOrientation, collectionItem.PictureID, collectionItem.Chart_ShowLegend, collectionItem.Chart_Type, collectionItem.Chart_ShowGrid, collectionItem.Chart_StackSeries, collectionItem.Chart_ShowValues, collectionItem.Chart_ViewID, collectionItem.Chart_TableID, collectionItem.Chart_ColumnID, collectionItem.Chart_FilterID, collectionItem.Chart_AggregateType, collectionItem.Chart_ColumnName, collectionItem.Chart_ColumnName_2, collectionItem.UseFormatting, collectionItem.Formatting_DecimalPlaces, collectionItem.Formatting_Use1000Separator, collectionItem.Formatting_Prefix, collectionItem.Formatting_Suffix, collectionItem.UseConditionalFormatting, collectionItem.ConditionalFormatting_Operator_1, collectionItem.ConditionalFormatting_Value_1, collectionItem.ConditionalFormatting_Style_1, collectionItem.ConditionalFormatting_Colour_1, collectionItem.ConditionalFormatting_Operator_2, collectionItem.ConditionalFormatting_Value_2, collectionItem.ConditionalFormatting_Style_2, collectionItem.ConditionalFormatting_Colour_2, collectionItem.ConditionalFormatting_Operator_3, collectionItem.ConditionalFormatting_Value_3, collectionItem.ConditionalFormatting_Style_3, collectionItem.ConditionalFormatting_Colour_3, collectionItem.SeparatorColour, collectionItem.InitialDisplayMode, collectionItem.Chart_TableID_2, collectionItem.Chart_ColumnID_2, collectionItem.Chart_TableID_3, collectionItem.Chart_ColumnID_3, collectionItem.Chart_SortOrderID, collectionItem.Chart_SortDirection, collectionItem.Chart_ColourID, collectionItem.Chart_ShowPercentages))
			lstButtonInfo.AddRange(From collectionItem As Object In objDropdownInfo Select New navigationLink(collectionItem.ID, collectionItem.DrillDownHidden, collectionItem.LinkType, collectionItem.LinkOrder, collectionItem.Text, collectionItem.Text1, collectionItem.Text2, collectionItem.Prompt, collectionItem.ScreenID, collectionItem.TableID, collectionItem.ViewID, collectionItem.PageTitle, collectionItem.URL, collectionItem.UtilityType, collectionItem.UtilityID, collectionItem.NewWindow, collectionItem.BaseTable, collectionItem.LinkToFind, collectionItem.SingleRecord, collectionItem.PrimarySequence, collectionItem.SecondarySequence, collectionItem.FindPage, collectionItem.EmailAddress, collectionItem.EmailSubject, collectionItem.AppFilePath, collectionItem.AppParameters, collectionItem.DocumentFilePath, collectionItem.DisplayDocumentHyperlink, collectionItem.IsSeparator, collectionItem.Element_Type, collectionItem.SeparatorOrientation, collectionItem.PictureID, collectionItem.Chart_ShowLegend, collectionItem.Chart_Type, collectionItem.Chart_ShowGrid, collectionItem.Chart_StackSeries, collectionItem.Chart_ShowValues, collectionItem.Chart_ViewID, collectionItem.Chart_TableID, collectionItem.Chart_ColumnID, collectionItem.Chart_FilterID, collectionItem.Chart_AggregateType, collectionItem.Chart_ColumnName, collectionItem.Chart_ColumnName_2, collectionItem.UseFormatting, collectionItem.Formatting_DecimalPlaces, collectionItem.Formatting_Use1000Separator, collectionItem.Formatting_Prefix, collectionItem.Formatting_Suffix, collectionItem.UseConditionalFormatting, collectionItem.ConditionalFormatting_Operator_1, collectionItem.ConditionalFormatting_Value_1, collectionItem.ConditionalFormatting_Style_1, collectionItem.ConditionalFormatting_Colour_1, collectionItem.ConditionalFormatting_Operator_2, collectionItem.ConditionalFormatting_Value_2, collectionItem.ConditionalFormatting_Style_2, collectionItem.ConditionalFormatting_Colour_2, collectionItem.ConditionalFormatting_Operator_3, collectionItem.ConditionalFormatting_Value_3, collectionItem.ConditionalFormatting_Style_3, collectionItem.ConditionalFormatting_Colour_3, collectionItem.SeparatorColour, collectionItem.InitialDisplayMode, collectionItem.Chart_TableID_2, collectionItem.Chart_ColumnID_2, collectionItem.Chart_TableID_3, collectionItem.Chart_ColumnID_3, collectionItem.Chart_SortOrderID, collectionItem.Chart_SortDirection, collectionItem.Chart_ColourID, collectionItem.Chart_ShowPercentages))

			Dim viewModel = New NavLinksViewModel With {.NavigationLinks = lstButtonInfo, .NumberOfLinks = objDropdownInfo.Count}

			Return View(viewModel)
		End Function


		Public Sub ShowPhoto(imageName As String)
			'TODO fetch path from registry
			Const localImagesPath As String = "\\abs16090\hrprotemp\"

			'TODO fetch imagename from db
			Dim file = localImagesPath & imageName
			Dim fStream As New FileStream(file, FileMode.Open, FileAccess.Read)
			Dim br As New BinaryReader(fStream)

			' Show the number of bytes in the array.
			br.Close()
			fStream.Close()

			Response.ContentType = "image/png"
			Response.WriteFile(file)

		End Sub

		Function LogOff()
			Session("databaseConnection") = Nothing
			Return RedirectToAction("Login", "Account")
		End Function

	End Class


End Namespace




