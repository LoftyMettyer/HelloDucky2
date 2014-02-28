Option Explicit On
Option Strict Off

Imports System.Web.Mvc
Imports System.Web.UI.DataVisualization.Charting
Imports System.IO
Imports System.Web
Imports ADODB
Imports System.Drawing
Imports DMI.NET.Code
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server
Imports DMI.NET.Models
Imports System.Data.SqlClient


Namespace Controllers
	Public Class HomeController
		Inherits Controller

		Private MultiAxisChart As New Chart

		Private Structure MultiAxisChartVertical
			Public Vertical_ID As Integer
			Public Vertical As String
		End Structure
		Private Structure MultiAxisChartHorizontal
			Public Horizontal_ID As Integer
			Public Horizontal As String
			Public Colour As Integer
		End Structure

#Region "Configuration"

		Function Configuration() As ActionResult
			Return View()
		End Function

		<HttpPost()>
		Function Configuration_Submit(value As FormCollection)
			'If (Request.Form("txtPrimaryStartMode") = "") Then
			'    Return View()
			'End If
			On Error Resume Next

			Dim sTemp
			Dim sType = ""
			Dim sControlName

			If (Request.Form("txtPrimaryStartMode") <> "") Then


				' Save the user configuration settings.
				Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)

				objDatabase.SaveUserSetting("RecordEditing", "Primary", Request.Form("txtPrimaryStartMode"))
				objDatabase.SaveUserSetting("RecordEditing", "History", Request.Form("txtHistoryStartMode"))
				objDatabase.SaveUserSetting("RecordEditing", "LookUp", Request.Form("txtLookupStartMode"))
				objDatabase.SaveUserSetting("RecordEditing", "QuickAccess", Request.Form("txtQuickAccessStartMode"))
				objDatabase.SaveUserSetting("ExpressionBuilder", "ViewColours", Request.Form("txtExprColourMode"))
				objDatabase.SaveUserSetting("ExpressionBuilder", "NodeSize", Request.Form("txtExprNodeMode"))
				objDatabase.SaveUserSetting("IntranetFindWindow", "BlockSize", Request.Form("txtFindSize"))

				Session("PrimaryStartMode") = Request.Form("txtPrimaryStartMode")
				Session("HistoryStartMode") = Request.Form("txtHistoryStartMode")
				Session("LookupStartMode") = Request.Form("txtLookupStartMode")
				Session("QuickAccessStartMode") = Request.Form("txtQuickAccessStartMode")
				Session("ExprColourMode") = Request.Form("txtExprColourMode")
				Session("ExprNodeMode") = Request.Form("txtExprNodeMode")
				Session("FindRecords") = Request.Form("txtFindSize")

				'--------------------------------------------
				' Save the DefSel 'only mine' settings.
				'--------------------------------------------
				For i = 0 To 20
					Select Case i
						Case 0
							sType = "BatchJobs"
						Case 1
							sType = "Calculations"
						Case 2
							sType = "CrossTabs"
						Case 3
							sType = "CustomReports"
						Case 4
							sType = "DataTransfer"
						Case 5
							sType = "Export"
						Case 6
							sType = "Filters"
						Case 7
							sType = "GlobalAdd"
						Case 8
							sType = "GlobalUpdate"
						Case 9
							sType = "GlobalDelete"
						Case 10
							sType = "Import"
						Case 11
							sType = "MailMerge"
						Case 12
							sType = "Picklists"
						Case 13
							sType = "CalendarReports"
						Case 14
							sType = "Labels"
						Case 15
							sType = "LabelDefinition"
						Case 16
							sType = "MatchReports"
						Case 17
							sType = "CareerProgression"
						Case 18
							sType = "EmailGroups"
						Case 19
							sType = "RecordProfile"
						Case 20
							sType = "SuccessionPlanning"
					End Select

					sControlName = "txtOwner_" & sType
					sTemp = "onlymine " & sType

					objDatabase.SaveUserSetting("defsel", sTemp, Request.Form(sControlName))

				Next

				'--------------------------------------------
				' Save the Utility Warning settings.
				'--------------------------------------------
				For i = 0 To 4
					Select Case i
						Case 0
							sType = "DataTransfer"
						Case 1
							sType = "GlobalAdd"
						Case 2
							sType = "GlobalUpdate"
						Case 3
							sType = "GlobalDelete"
						Case 4
							sType = "Import"
					End Select

					sControlName = "txtWarn_" & sType
					sTemp = "warning " & sType

					objDatabase.SaveUserSetting("warningmsg", sTemp, Request.Form(sControlName))

				Next

				'--------------------------------------------
				' Redirect to the save confirmation page.
				'--------------------------------------------
				'Session("confirmtext") = "User Configuration has been saved successfully."
				'Session("confirmtitle") = "User Configuration"
				'Session("followpage") = "default"
				'Session("reaction") = Request.Form("txtReaction")
			End If

			Return RedirectToAction("CONFIGURATION")

		End Function

		Function PcConfiguration() As ActionResult
			Return View()
		End Function

#End Region

		<HttpPost()>
		Function util_def_crosstabs_submit(value As FormCollection)

			Try

				Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
				Dim prmID = New SqlParameter("piId", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = CleanNumeric(Request.Form("txtSend_ID"))}

				objDataAccess.ExecuteSP("sp_ASRIntSaveCrossTab", _
						New SqlParameter("psName", SqlDbType.VarChar, 255) With {.Value = Request.Form("txtSend_name")}, _
						New SqlParameter("psDescription", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_description")}, _
						New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_baseTable"))}, _
						New SqlParameter("piSelection", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_allRecords"))}, _
						New SqlParameter("piPicklistID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_picklist"))}, _
						New SqlParameter("piFilterID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_filter"))}, _
						New SqlParameter("pfPrintFilter", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_printFilter"))}, _
						New SqlParameter("psUserName", SqlDbType.VarChar, 255) With {.Value = Request.Form("txtSend_userName")}, _
						New SqlParameter("piHColID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_HColID"))}, _
						New SqlParameter("psHStart", SqlDbType.VarChar, 100) With {.Value = Request.Form("txtSend_HStart")}, _
						New SqlParameter("psHStop", SqlDbType.VarChar, 100) With {.Value = Request.Form("txtSend_HStop")}, _
						New SqlParameter("psHStep", SqlDbType.VarChar, 100) With {.Value = Request.Form("txtSend_HStep")}, _
						New SqlParameter("piVColID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_VColID"))}, _
						New SqlParameter("psVStart", SqlDbType.VarChar, 100) With {.Value = Request.Form("txtSend_VStart")}, _
						New SqlParameter("psVStop", SqlDbType.VarChar, 100) With {.Value = Request.Form("txtSend_VStop")}, _
						New SqlParameter("psVStep", SqlDbType.VarChar, 100) With {.Value = Request.Form("txtSend_VStep")}, _
						New SqlParameter("piPColID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_PColID"))}, _
						New SqlParameter("psPStart", SqlDbType.VarChar, 100) With {.Value = Request.Form("txtSend_PStart")}, _
						New SqlParameter("psPStop", SqlDbType.VarChar, 100) With {.Value = Request.Form("txtSend_PStop")}, _
						New SqlParameter("psPStep", SqlDbType.VarChar, 100) With {.Value = Request.Form("txtSend_PStep")}, _
						New SqlParameter("piIType", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_IType"))}, _
						New SqlParameter("piIColID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_IColID"))}, _
						New SqlParameter("pfPercentage", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_Percentage"))}, _
						New SqlParameter("pfPerPage", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_PerPage"))}, _
						New SqlParameter("pfSuppress", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_Suppress"))}, _
						New SqlParameter("pfUse1000Separator", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_Use1000Separator"))}, _
						New SqlParameter("pfOutputPreview", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_OutputPreview"))}, _
						New SqlParameter("piOutputFormat", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_OutputFormat"))}, _
						New SqlParameter("pfOutputScreen", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_OutputScreen"))}, _
						New SqlParameter("pfOutputPrinter", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_OutputPrinter"))}, _
						New SqlParameter("psOutputPrinterName", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_OutputPrinterName")}, _
						New SqlParameter("pfOutputSave", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_OutputSave"))}, _
						New SqlParameter("piOutputSaveExisting", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_OutputSaveExisting"))}, _
						New SqlParameter("pfOutputEmail", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_OutputEmail"))}, _
						New SqlParameter("piOutputEmailAddr", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_OutputEmailAddr"))}, _
						New SqlParameter("psOutputEmailSubject", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_OutputEmailSubject")}, _
						New SqlParameter("psOutputEmailAttachAs", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_OutputEmailAttachAs")}, _
						New SqlParameter("psOutputFilename", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_OutputFilename")}, _
						New SqlParameter("psAccess", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_access")}, _
						New SqlParameter("psJobsToHide", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_jobsToHide")}, _
						New SqlParameter("psJobsToHideGroups", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_jobsToHideGroups")}, _
						prmID)

				Session("confirmtext") = "Cross tab has been saved successfully"
				Session("confirmtitle") = "Cross Tabs"
				Session("followpage") = "defsel.asp"
				Session("reaction") = Request.Form("txtSend_reaction")
				Session("utilid") = prmID.Value

				Return RedirectToAction("ConfirmOK")

			Catch ex As Exception

				Response.Write("<html>" & vbCrLf)
				Response.Write("	<head>" & vbCrLf)
				Response.Write("		<meta name=""GENERATOR"" Content=""Microsoft Visual Studio 6.0"">" & vbCrLf)
				Response.Write("		<link href=""OpenHR.css"" rel=stylesheet type=text/css >" & vbCrLf)
				Response.Write("		<title>" & vbCrLf)
				Response.Write("			OpenHR Intranet" & vbCrLf)
				Response.Write("		</title>" & vbCrLf)
				Response.Write("  <!--#INCLUDE FILE=""include/ctl_SetStyles.txt"" -->")
				Response.Write("	</head>" & vbCrLf)
				Response.Write("	<body id='bdyMainBody' name='bdyMainBody' " & Session("BodyTag") & ">" & vbCrLf)

				Response.Write("	<table align='center' class='outline' cellPadding='5' cellSpacing='0'>" & vbCrLf)
				Response.Write("		<tr>" & vbCrLf)
				Response.Write("			<td>" & vbCrLf)
				Response.Write("				<table class='invisible' cellspacing='0' cellpadding='0'>" & vbCrLf)
				Response.Write("				  <tr> " & vbCrLf)
				Response.Write("				    <td colspan='3' height='10'></td>" & vbCrLf)
				Response.Write("				  </tr>" & vbCrLf)
				Response.Write("				  <tr> " & vbCrLf)
				Response.Write("				    <td colspan='3' align='center'> " & vbCrLf)
				Response.Write("							<h3>Error</h3>" & vbCrLf)
				Response.Write("				    </td>" & vbCrLf)
				Response.Write("				  </tr>" & vbCrLf)
				Response.Write("				  <tr> " & vbCrLf)
				Response.Write("				    <td width='20' height='10'></td> " & vbCrLf)
				Response.Write("				    <td> " & vbCrLf)
				Response.Write("							<h4>Error saving cross tab</h4>" & vbCrLf)
				Response.Write("				    </td>" & vbCrLf)
				Response.Write("				    <td width='20'></td> " & vbCrLf)
				Response.Write("				  </tr>" & vbCrLf)
				Response.Write("				  <tr> " & vbCrLf)
				Response.Write("				    <td width='20' height='10'></td> " & vbCrLf)
				Response.Write("				    <td> " & vbCrLf)
				Response.Write(ex.Message & vbCrLf)
				Response.Write("			    </td>" & vbCrLf)
				Response.Write("			    <td width='20'></td> " & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr> " & vbCrLf)
				Response.Write("			    <td colspan='3' height='20'></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr> " & vbCrLf)
				Response.Write("			    <td colspan='3' height='10' align='center'>" & vbCrLf)
				Response.Write("						<input type='button' value='Retry' name='GoBack' class='btn' OnClick='window.history.back(1)' style='width: 80px' id='cmdGoBack' />" & vbCrLf)
				Response.Write("			    </td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr>" & vbCrLf)
				Response.Write("			    <td colspan='3' height='10'></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			</table>" & vbCrLf)
				Response.Write("    </td>" & vbCrLf)
				Response.Write("  </tr>" & vbCrLf)
				Response.Write("</table>" & vbCrLf)
				Response.Write("	</body>" & vbCrLf)
				Response.Write("</html>" & vbCrLf)
			End Try

		End Function

		<HttpPost()>
		Function newUser_Submit(value As FormCollection) As JsonResult

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			Dim fSubmitNewUser = (Len(Request.Form("txtGotoPage")) = 0)

			If fSubmitNewUser Then
				' Read the Password details from the Password form.
				Dim sNewUserLogin = Request.Form("selNewUser")

				' Create an OpenHR user associated with the
				' given SQL Server login.

				Try
					objDataAccess.ExecuteSP("sp_ASRIntNewUser", _
							New SqlParameter("@psUserName", SqlDbType.VarChar, 128) With {.Value = sNewUserLogin})


				Catch ex As Exception
					Session("ErrorTitle") = "New User Page"
					Session("ErrorText") = "You could not add the user because of the following error:<p>" & FormatError(ex.Message)
					Dim data1 = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = ""}
					Return Json(data1, JsonRequestBehavior.AllowGet)

				End Try

				Session("ErrorTitle") = "New User Page"
				Session("ErrorText") = "User added successfully."
				Dim data = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = ""}
				Return Json(data, JsonRequestBehavior.AllowGet)

			Else
				' Read the information from the calling form.
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

				' Go to the requested page.
				' Response.Redirect(Request.Form("txtGotoPage"))
				Session("txtGotoPage") = Request.Form("txtGotoPage")
			End If

		End Function

		<HttpPost()>
		Function passwordChange_Submit(value As FormCollection) As JsonResult

			On Error Resume Next

			Dim sReferringPage = ""
			Dim fSubmitPasswordChange = ""
			Dim sErrorText = ""
			Dim fRedirectToSSI As Boolean

			If True Then
				fSubmitPasswordChange = (Len(Request.Form("txtGotoPage")) = 0)

				If fSubmitPasswordChange Then
					' Force password change only if there are no other users logged in with the same name.
					Dim iUserSessionCount As Integer = ASRFunctions.GetCurrentUsersCountOnServer(Session("Username"))

					' variables to help select which main screen we return to after change or cancel
					fRedirectToSSI = CleanBoolean(Request.Form("txtRedirectToSSI"))
					Dim sMainRedirect = IIf(fRedirectToSSI, "Main?SSIMode=True", "main")

					If iUserSessionCount < 2 Then
						' Read the Password details from the Password form.
						Dim sCurrentPassword = Request.Form("txtCurrentPassword")
						Dim sNewPassword = Request.Form("txtPassword1")



						' Attempt to change the password on the SQL Server.
						Dim cmdChangePassword = CreateObject("ADODB.Command")
						cmdChangePassword.CommandText = "sp_password"
						cmdChangePassword.CommandType = 4	' Stored Procedure
						cmdChangePassword.ActiveConnection = Session("databaseConnection")

						Dim prmCurrentPassword = cmdChangePassword.CreateParameter("currentPassword", 200, 1, 255)
						cmdChangePassword.Parameters.Append(prmCurrentPassword)
						If Len(sCurrentPassword) > 0 Then
							prmCurrentPassword.value = sCurrentPassword
						Else
							prmCurrentPassword.value = DBNull.Value
						End If

						Dim prmNewPassword = cmdChangePassword.CreateParameter("newPassword", 200, 1, 255)
						cmdChangePassword.Parameters.Append(prmNewPassword)
						If Len(sNewPassword) > 0 Then
							prmNewPassword.value = sNewPassword
						Else
							prmNewPassword.value = DBNull.Value
						End If

						Err.Clear()
						cmdChangePassword.Execute()

						' Release the ADO command object.
						cmdChangePassword = Nothing

						If Err.Number <> 0 Then
							Session("ErrorTitle") = "Change Password Page"
							Session("ErrorText") = "You could not change your password because of the following error:<p>" & FormatError(Err.Description)
							Dim data = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = sMainRedirect}
							Return Json(data, JsonRequestBehavior.AllowGet)
							' Return RedirectToAction("error", "home")
						Else
							' Password changed okay. Update the appropriate record in the ASRSysPasswords table.
							Dim cmdPasswordOK = CreateObject("ADODB.Command")
							cmdPasswordOK.CommandText = "sp_ASRIntPasswordOK"
							cmdPasswordOK.CommandType = 4	' Stored Procedure
							cmdPasswordOK.ActiveConnection = Session("databaseConnection")

							Err.Clear()
							cmdPasswordOK.Execute()
							If Err.Number <> 0 Then
								Session("ErrorTitle") = "Change Password Page"
								Session("ErrorText") = "You could not change your password because of the following error:<p>" & FormatError(Err.Description)
								Dim data1 = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = sMainRedirect}
								Return Json(data1, JsonRequestBehavior.AllowGet)
								' Return RedirectToAction("error", "Account")
							End If

							' Release the ADO command object.
							cmdPasswordOK = Nothing

							' Close and reopen the connection object.
							Dim conX = Session("databaseConnection")
							Dim sConnString = conX.connectionString

							Dim iPos1 = InStr(UCase(sConnString), UCase(";PWD=" & sCurrentPassword))
							If iPos1 > 0 Then
								conX.close()
								conX = Nothing
								Session("databaseConnection") = ""


								Dim sNewConnString = Left(sConnString, iPos1 + 4) & sNewPassword & Mid(sConnString, iPos1 + 5 + Len(sCurrentPassword))
								' Open a connection to the database.
								conX = CreateObject("ADODB.Connection")
								conX.open(sNewConnString)

								If Err.Number <> 0 Then
									Session("ErrorTitle") = "Change Password Page"
									Session("ErrorText") = "You could not change your password because of the following error:<p>" & FormatError(Err.Description)
									Dim data1 = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = sMainRedirect}
									Return Json(data1, JsonRequestBehavior.AllowGet)
									' Return RedirectToAction("error", "Account")
								End If

								Session("databaseConnection") = conX

							End If

							' Tell the user that the password was changed okay.
							Session("ErrorTitle") = "Change Password Page"
							Session("ErrorText") = "Password changed successfully."

							Dim data = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = sMainRedirect}
							Return Json(data, JsonRequestBehavior.AllowGet)
							' Return RedirectToAction("message", "Account")
						End If
					Else
						Session("ErrorTitle") = "Change Password Page"
						sErrorText = "You could not change your password.<p>The account is currently being used by "
						If iUserSessionCount > 2 Then
							sErrorText = sErrorText & iUserSessionCount & " users"
						Else
							sErrorText = sErrorText & "another user"
						End If
						sErrorText = sErrorText & " in the system."
						Session("ErrorText") = sErrorText

						' Return RedirectToAction("Loginerror", "Account")
					End If
				Else
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

					' Go to the requested page.
					' Return RedirectToAction(Request.Form("txtGotoPage"))
					Session("txtGotoPage") = Request.Form("txtGotoPage")
				End If
			End If
		End Function

		Function ConfirmOK() As ActionResult
			Return View()
		End Function

		' GET: /Home
		Function Main(Optional SSIMode As Boolean = vbFalse) As ActionResult

			'Dim iSingleRecordViewID As Integer = CleanNumeric(Session("SingleRecordViewID"))

			'Dim prmRecordID = New SqlParameter("piRecordID", SqlDbType.Int)
			'prmRecordID.Direction = ParameterDirection.Output

			'Dim prmRecordCount = New SqlParameter("piRecordCount", SqlDbType.Int)
			'prmRecordCount.Direction = ParameterDirection.Output

			'clsDataAccess.GetDataSet("spASRIntGetSelfServiceRecordID", prmRecordID, prmRecordCount _
			'											, New SqlParameter("piViewID", iSingleRecordViewID))

			'' Reload the toplevelrecid session variable as linksMain may have reset it.
			'Dim sErrorDescription = ""

			'If (Err.Number <> 0) Then
			'	sErrorDescription = "Unable to get the personnel record ID." & vbCrLf & FormatError(Err.Description)
			'End If

			'If Len(sErrorDescription) = 0 Then
			'	If prmRecordCount.Value = 1 Then
			'		' Only one record.
			'		Session("TopLevelRecID") = CLng(prmRecordID.Value)
			'	Else
			'		If prmRecordCount.Value = 0 Then
			'			' No personnel record. 
			'			Session("TopLevelRecID") = 0
			'		Else
			'			' More than one personnel record.
			'			sErrorDescription = "You have access to more than one record in the defined Single-record view."

			'			Session("ErrorTitle") = "Login Page"
			'			Session("ErrorText") =
			'			 "You could not login to the OpenHR database because of the following reason:" & sErrorDescription & "<p>" & vbCrLf

			'			Response.Redirect("FormError")

			'			' Return RedirectToAction("Loginerror", "Account")
			'		End If
			'	End If
			'Else
			'	Session("ErrorTitle") = "Login Page"
			'	Session("ErrorText") =
			'	 "You could not login to the OpenHR database because of the following reason:" & vbCrLf & sErrorDescription & "<p>" & vbCrLf
			'	Response.Redirect("FormError")
			'	' Return RedirectToAction("Loginerror", "Account")
			'End If

			''	cmdSSRecord = Nothing

			ResetSessionVars()

			Session("selectSQL") = ""
			ViewBag.SSIMode = SSIMode

			Return View()
		End Function

		Function Find(Optional sParameters As String = "") As ActionResult
			'Data access variables
			Dim objSession As SessionInfo = CType(Session("SessionContext"), SessionInfo)	'Set session info
			Dim objDataAccess As New clsDataAccess(objSession.LoginInfo) 'Instantiate DataAccess class
			Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)

			Dim SPParameters() As SqlParameter

			' Additional controller actions for SSI view. Only SSI calls to this action have parameters.
			If sParameters.Length > 0 Then
				' =========================
				' Self-service Find request
				' =========================
				Dim lngTopLevelRecordID As Integer
				Dim sTableName As String
				Dim sViewName As String

				Dim objUser As New clsSettings
				objUser.SessionInfo = CType(Session("SessionContext"), SessionInfo)
				Dim sErrorDescription = ""

				Dim iRealTableID = 0
				Dim iRealViewID = 0

				lngTopLevelRecordID = Session("TopLevelRecID")

				If NullSafeInteger(Session("tableType")) <> 2 Then
					' Top Level table.
					'Response.Write "#<FONT COLOR='Red'><B>Top Level table.</B></FONT>#<BR>"

					Session("recordID") = lngTopLevelRecordID
					Session("parentTableID") = 0
					Session("parentRecordID") = 0
				Else
					' Child table.
					' Response.Write "#<FONT COLOR='Red'><B>Child table.</B></FONT>#<BR>"

					iRealTableID = Session("SSILinkTableID")
					iRealViewID = Session("SSILinkViewID")
					'session("tableID") = 0 
					Session("viewID") = 0
					Session("parentTableID") = Session("SSILinkTableID")
					Session("parentRecordID") = lngTopLevelRecordID
				End If

				' Read the screen info from the query string.			

				'Response.Write "#<FONT COLOR='Red'><B>sParameters = " & sParameters & "</B></FONT>#<BR>"
				'Response.Write "#<FONT COLOR='Red'><B>parentTableID = " & session("parentTableID") & "</B></FONT>#<BR>"
				'Response.Write "#<FONT COLOR='Red'><B>parentRecordID = " & session("parentRecordID") & "</B></FONT>#<BR>"

				Session("action") = Left(sParameters, InStr(sParameters, "_") - 1)
				sParameters = Mid(sParameters, InStr(sParameters, "_") + 1)
				Session("firstRecPos") = Left(sParameters, InStr(sParameters, "_") - 1)
				sParameters = Mid(sParameters, InStr(sParameters, "_") + 1)
				Session("currentRecCount") = Left(sParameters, InStr(sParameters, "_") - 1)
				Session("locateValue") = Mid(sParameters, InStr(sParameters, "_") + 1)

				' Flag an error if there is no current table or view is specified.
				If (Session("tableID") <= 0) Then
					'and (session("viewID") <= 0) then

					sErrorDescription = "The find page could not be loaded." & vbCrLf & "No table or view specified."
				End If


				If Len(sErrorDescription) = 0 Then
					' Flag an error if there is no current screen is specified.
					If (Session("linkType") <> "multifind") And _
						(Session("screenID") <= 0) Then
						sErrorDescription = "The find page could not be loaded." & vbCrLf & "No screen specified."
					End If
				End If

				If Len(sErrorDescription) = 0 Then
					If (Session("linkType") = "multifind") Then
						Dim prm_piOrderID As New SqlParameter("@piOrderID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
						SPParameters = New SqlParameter() { _
								New SqlParameter("@piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("tableID"))}, _
								prm_piOrderID _
						}

						Try
							objDataAccess.ExecuteSP("spASRIntGetDefaultOrder", SPParameters)
						Catch ex As Exception
							sErrorDescription = "The find page could not be loaded." & vbCrLf & "The default order for the table could not be determined :" & vbCrLf & FormatError(ex.Message)
						End Try

						Session("orderID") = prm_piOrderID.Value
					Else
						' Get the screen's default order if none is already specified.
						Dim prm_plngOrderID As New SqlParameter("@plngOrderID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
						SPParameters = New SqlParameter() { _
								New SqlParameter("@plngScreenID", SqlDbType.Int) With {.Value = CleanNumeric(Session("screenID"))}, _
								prm_plngOrderID _
						}

						Try
							objDataAccess.ExecuteSP("sp_ASRIntGetScreenOrder", SPParameters)
						Catch ex As Exception
							sErrorDescription = "The find page could not be loaded." & vbCrLf & "The default order for the table could not be determined :" & vbCrLf & FormatError(ex.Message)
						End Try

						Session("orderID") = prm_plngOrderID.Value
					End If
				End If

				'Response.Write "#<FONT COLOR='Red'>session(SSILinkViewID) = <B>" & session("SSILinkViewID") & "</B></FONT>#<BR>"
				'Response.Write "#<FONT COLOR='Red'>session(SSILinkTableID) = <B>" & session("SSILinkTableID") & "</B></FONT>#<BR>"
				'Response.Write "#<FONT COLOR='Red'>session(PersonnelTableID) = <B>" & session("PersonnelTableID") & "</B></FONT>#<BR>"
				'Response.Write "#<FONT COLOR='Red'>session(TopLevelRecID) = <B>" & session("TopLevelRecID") & "</B></FONT>#<BR>"
				'Response.Write "#<FONT COLOR='Red'>session(SingleRecordViewID) = <B>" & session("SingleRecordViewID") & "</B></FONT>#<BR>"
				'Response.Write "#<FONT COLOR='Red'>session(tableID) = <B>" & session("tableID") & "</B></FONT>#<BR>"
				'Response.Write "#<FONT COLOR='Red'>session(viewID) = <B>" & session("viewID") & "</B></FONT>#<BR>"

				If Len(sErrorDescription) = 0 Then

					If NullSafeInteger(Session("SSILinkViewID")) = NullSafeInteger(Session("SingleRecordViewID")) Then
						lngTopLevelRecordID = Session("TopLevelRecID")
					End If

					If NullSafeInteger(Session("tableType")) <> 2 Then
						' Top Level table.
						Session("recordID") = 0	'  lngPersonnelRecordID			' never set???
						Session("parentTableID") = 0
						Session("parentRecordID") = 0
					Else
						' Child table.
						Session("parentTableID") = Session("SSILinkTableID")
						Session("parentRecordID") = lngTopLevelRecordID
					End If

					' Enable response buffering as we may redirect the response further down this page.
					Response.Buffer = True
				End If

				Dim sRecDesc = ""
				If NullSafeInteger(Session("SSILinkViewID")) <> NullSafeInteger(Session("SingleRecordViewID")) And (Len(sErrorDescription) = 0) Then

					Try

						Dim prmRecordDesc As New SqlParameter("psRecDesc", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
						Dim prmErrorMessage As New SqlParameter("psErrorMessage", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

						objDataAccess.ExecuteSP("spASRIntGetRecordDescriptionInView" _
								, New SqlParameter("piViewID", SqlDbType.Int) With {.Value = CleanNumeric(Session("SSILinkViewID"))} _
								, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("tableID"))} _
								, New SqlParameter("piRecordID", SqlDbType.Int) With {.Value = 0} _
								, New SqlParameter("piParentTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("parentTableID"))} _
								, New SqlParameter("piParentRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("parentRecordID"))} _
								, prmRecordDesc _
								, prmErrorMessage)

						If prmErrorMessage.Value.ToString().Length > 0 Then
							sErrorDescription = "Unable to get the record description." & vbCrLf & prmErrorMessage.Value.ToString()
						Else
							sRecDesc = prmRecordDesc.Value.ToString()
						End If

					Catch ex As Exception
						sErrorDescription = "Unable to get the record description." & vbCrLf & ex.Message

					End Try


				End If

				If (Len(sErrorDescription) = 0) Then
					Dim sTitle As String = ""

					If (Session("linkType") <> "multifind") Then

						sTableName = Replace(objDatabase.GetTableName(CInt(Session("tableID"))), "_", " ")

						sTitle = "Select the required "

						If Len(sTableName) > 0 Then
							sTitle = sTitle & sTableName & " "
						End If

						sTitle = sTitle & "record"

						If Len(sRecDesc) > 0 Then
							sTitle = sTitle & " for " & sRecDesc
						End If
					Else

						Try

							Dim prmPageTitle = New SqlParameter("psPageTitle", SqlDbType.VarChar, 200) With {.Direction = ParameterDirection.Output}

							objDataAccess.ExecuteSP("spASRIntGetPageTitle" _
								, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("TableID"))} _
								, New SqlParameter("piViewID", SqlDbType.Int) With {.Value = CleanNumeric(Session("ViewID"))} _
								, prmPageTitle)

							sTitle = Replace(prmPageTitle.Value.ToString(), "_", " ")

						Catch ex As Exception
							sErrorDescription = "Error getting the page title." & vbCrLf & FormatError(ex.Message)

						End Try
					End If

					ViewBag.pageTitle = sTitle
				End If

				If (Len(sErrorDescription) = 0) Then

					If NullSafeInteger(Session("SSILinkViewID")) > -1 Then

						Try
							sViewName = ""

							Dim prmViewName = New SqlParameter("psViewName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
							Dim prmViewID = New SqlParameter("piViewID", SqlDbType.Int)

							If NullSafeInteger(Session("SSILinkViewID")) <> NullSafeInteger(Session("SingleRecordViewID")) And (Session("linkType") <> "multifind") Then
								prmViewID.Value = CleanNumeric(Session("SSILinkViewID"))
							Else
								prmViewID.Value = CleanNumeric(Session("SingleRecordViewID"))
							End If

							objDataAccess.ExecuteSP("spASRIntGetViewName", prmViewID, prmViewName)

							If Not IsDBNull(prmViewName.Value) Then
								sViewName = Replace(prmViewName.Value.ToString(), "_", " ")
							End If

						Catch ex As Exception
							sErrorDescription = "Error getting the link view name." & vbCrLf & FormatError(ex.Message)

						End Try

					Else

						sTableName = Replace(objDatabase.GetTableName(CInt(Session("SSILinkTableID"))), "_", " ")

					End If

					If (NullSafeInteger(Session("SSILinkViewID")) = NullSafeInteger(Session("SingleRecordViewID")) Or _
						(Session("linkType") = "multifind")) And _
						Session("SingleRecordViewID") = 0 Then

						sViewName = "single record"
					End If
				End If
			Else

				' Flag an error if there is no current table or view is specified.
				If (Session("tableID") <= 0) And _
				 (Session("viewID") <= 0) Then

					Session("ErrorTitle") = "Find Page"
					Session("ErrorText") = "No table or view specified."
					Response.Redirect("FormError")
				End If

				' Flag an error if there is no current screen is specified.
				If Session("screenID") <= 0 Then
					Session("ErrorTitle") = "Find Page"
					Session("ErrorText") = "No screen specified."
					Response.Redirect("FormError")
				End If

				' Get the screen's default order if none is already specified.
				If Session("orderID") <= 0 Then

					Try

						Dim prmOrder As New SqlParameter("plngOrderID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
						SPParameters = New SqlParameter() { _
								New SqlParameter("plngScreenID", SqlDbType.Int) With {.Value = CleanNumeric(Session("screenID"))}, _
								prmOrder}
						objDataAccess.ExecuteSP("sp_ASRIntGetScreenOrder", SPParameters)

						Session("orderID") = prmOrder.Value.ToString()

					Catch ex As Exception
						Session("ErrorTitle") = "Find Page"
						Session("ErrorText") = "The default order for the screen could not be determined :<p>" & vbNewLine & ex.Message

					End Try

				End If

				' Enable response buffering as we may redirect the response further down this page.
				Response.Buffer = True

				ViewBag.pageTitle = ""
			End If

			Return View()

		End Function

		Function _default() As ActionResult
			Return View()
		End Function


		<HttpPost()>
		Function default_Submit()

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
			Session("singleRecordID") = CInt(Request.Form("txtGotoOptionDefSelRecordID"))
			Session("optionRecordID") = 0
			Session("optionAction") = ""

			' Go to the requested page.
			Return RedirectToAction(Request.Form("txtGotoPage").Replace(".asp", ""))

		End Function


		<HttpPost()>
		Function emptyoption_Submit()

			On Error Resume Next

			' Save the required information in session variables.
			Session("optionScreenID") = Request.Form("txtGotoOptionScreenID")
			Session("optionTableID") = Request.Form("txtGotoOptionTableID")
			Session("optionViewID") = Request.Form("txtGotoOptionViewID")
			Session("optionOrderID") = Request.Form("txtGotoOptionOrderID")
			Session("optionRecordID") = Request.Form("txtGotoOptionRecordID")
			Session("optionFilterDef") = Request.Form("txtGotoOptionFilterDef")
			Session("optionFilterSQL") = Request.Form("txtGotoOptionFilterSQL")
			Session("optionValue") = Request.Form("txtGotoOptionValue")
			Session("optionLinkTableID") = Request.Form("txtGotoOptionLinkTableID")
			Session("optionLinkOrderID") = Request.Form("txtGotoOptionLinkOrderID")
			Session("optionLinkViewID") = Request.Form("txtGotoOptionLinkViewID")
			Session("optionLinkRecordID") = Request.Form("txtGotoOptionLinkRecordID")
			Session("optionColumnID") = Request.Form("txtGotoOptionColumnID")
			Session("optionLookupColumnID") = Request.Form("txtGotoOptionLookupColumnID")
			Session("optionLookupMandatory") = Request.Form("txtGotoOptionLookupMandatory")
			Session("optionLookupValue") = Request.Form("txtGotoOptionLookupValue")
			Session("optionLookupFilterValue") = Request.Form("txtGotoOptionLookupFilterValue")
			Session("optionFile") = Request.Form("txtGotoOptionFile")
			Session("optionExtension") = Request.Form("txtGotoOptionExtension")
			'Session("optionOLEOnServer") = Request.Form("txtGotoOptionOLEOnServer")
			Session("optionAction") = Request.Form("txtGotoOptionAction")
			Session("optionPageAction") = Request.Form("txtGotoOptionPageAction")
			Session("optionCourseTitle") = Request.Form("txtGotoOptionCourseTitle")
			Session("optionFirstRecPos") = Request.Form("txtGotoOptionFirstRecPos")
			Session("optionCurrentRecCount") = Request.Form("txtGotoOptionCurrentRecCount")
			Session("optionExprType") = Request.Form("txtGotoOptionExprType")
			Session("optionExprID") = Request.Form("txtGotoOptionExprID")
			Session("optionFunctionID") = Request.Form("txtGotoOptionFunctionID")
			Session("optionParameterIndex") = Request.Form("txtGotoOptionParameterIndex")
			Session("OptionRealsource") = Request.Form("txtGotoOptionRealsource")
			Session("StandardReport_Type") = Request.Form("txtStandardReportType")
			Session("optionDefSelType") = Request.Form("txtGotoOptionDefSelType")
			Session("singleRecordID") = CInt(Request.Form("txtGotoOptionDefSelRecordID"))
			Session("optionOLEType") = Request.Form("txtGotoOptionOLEType")
			Session("optionOLEMaxEmbedSize") = Request.Form("txtGotoOptionOLEMaxEmbedSize")
			Session("optionOLEReadOnly") = Request.Form("txtGotoOptionOLEReadOnly")
			Session("optionIsPhoto") = Request.Form("txtGotoOptionIsPhoto")
			Session("optionOnlyNumerics") = Request.Form("txtOptionOnlyNumerics")
			Session("StandardReport_Type") = Request.Form("txtStandardReportType")

			' Go to the requested page.
			Return RedirectToAction(Request.Form("txtGotoOptionPage"))

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
					Case 10	' PICKLISTS
						Return RedirectToAction("util_def_picklist")
					Case 11	' FILTERS
						Return RedirectToAction("util_def_expression")
					Case 12	' CALCULATIONS
						Return RedirectToAction("util_def_expression")
					Case 17	' CALENDAR REPORTS
						Return RedirectToAction("util_def_calendarreport")
						'Case 25	' WORKFLOW 
						'Return RedirectToAction("util_run_workflow")
				End Select

			ElseIf Session("action") = "delete" Then
				Select Case Session("utiltype")
					Case 1	' CROSS TABS
						Session("reaction") = "CROSSTABS"
					Case 2	' CUSTOM REPORTS
						Session("reaction") = "CUSTOMREPORTS"
					Case 9	' MAIL MERGE
						Session("reaction") = "MAILMERGE"
					Case 10	' PICKLISTS
						Session("reaction") = "PICKLISTS"
					Case 11	' FILTERS
						Session("reaction") = "FILTERS"
					Case 12	' CALCULATIONS
						Session("reaction") = "CALCULATIONS"
					Case 17	' CALENDAR REPORTS
						Session("reaction") = "CALENDARREPORTS"
						'Case 25	' WORKFLOW 
						'	Session("reaction") = "WORKFLOWS"
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

		Function util_def_crosstabs() As ActionResult
			Return View()
		End Function

		Function CheckForUsage() As ActionResult
			Return View()
		End Function

		Function util_delete() As ActionResult
			Return View()
		End Function

		Function Data() As ActionResult
			Return View()
		End Function

		Function OptionData() As ActionResult
			Return View()
		End Function

		Function optionData_Submit() As ActionResult

			On Error Resume Next

			' Read the information from the calling form.
			Session("optionAction") = Request.Form("txtOptionAction")
			Session("optionTableID") = Request.Form("txtOptionTableID")
			Session("optionViewID") = Request.Form("txtOptionViewID")
			Session("optionOrderID") = Request.Form("txtOptionOrderID")
			Session("optionColumnID") = Request.Form("txtOptionColumnID")
			Session("optionPageAction") = Request.Form("txtOptionPageAction")
			Session("optionFirstRecPos") = Request.Form("txtOptionFirstRecPos")
			Session("optionCurrentRecCount") = Request.Form("txtOptionCurrentRecCount")
			Session("optionLocateValue") = Request.Form("txtGotoLocateValue")
			Session("optionCourseTitle") = Request.Form("txtOptionCourseTitle")
			Session("optionRecordID") = Request.Form("txtOptionRecordID")
			Session("optionLinkRecordID") = Request.Form("txtOptionLinkRecordID")
			Session("optionValue") = Request.Form("txtOptionValue")
			Session("optionSQL") = Request.Form("txtOptionSQL")
			Session("optionPromptSQL") = Request.Form("txtOptionPromptSQL")
			Session("optionOnlyNumerics") = Request.Form("txtOptionOnlyNumerics")
			Session("optionLookupColumnID") = Request.Form("txtOptionLookupColumnID")
			Session("optionFilterValue") = Request.Form("txtOptionLookupFilterValue")
			Session("IsLookupTable") = Request.Form("txtOptionIsLookupTable")
			Session("optionParentTableID") = Request.Form("txtOptionParentTableID")
			Session("optionParentRecordID") = Request.Form("txtOptionParentRecordID")
			Session("option1000SepCols") = Request.Form("txtOption1000SepCols")

			Session("StandardReport_Type") = Request.Form("txtStandardReportType")

			' Go to the requested page.
			Return RedirectToAction("OptionData")

		End Function

		Function Data_Submit() As ActionResult
			Dim iRETRIES = 5
			Dim iRetryCount = 0
			Dim sErrorMsg = "", sErrMsg = ""
			Dim fWarning = False
			Dim fOk = False
			Dim fTBOverride = False

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

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

				If (Not fTBOverride) And (NullSafeInteger(lngTableID) = NullSafeInteger(Session("TB_TBTableID"))) Then
					' Training Booking check.

					Try

						Dim prmResult = New SqlParameter("@piResultCode", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

						objDataAccess.ExecuteSP("sp_ASRIntValidateTrainingBooking" _
							, prmResult _
							, New SqlParameter("piEmpRecID", SqlDbType.Int) With {.Value = CleanNumeric(iTBEmployeeRecordID)} _
							, New SqlParameter("piCourseRecID", SqlDbType.Int) With {.Value = CleanNumeric(iTBCourseRecordID)} _
							, New SqlParameter("psBookingStatus", SqlDbType.VarChar, -1) With {.Value = sTBBookingStatusValue} _
							, New SqlParameter("piTBRecID", SqlDbType.Int) With {.Value = CleanNumeric(lngRecordID)})

						iTBResultCode = prmResult.Value

					Catch ex As Exception
						sErrorMsg = "Error validating training booking." & vbCrLf & FormatError(ex.Message)

					End Try


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
							Try
								Dim prmRecordID As New SqlParameter("piNewRecordID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
								objDataAccess.ExecuteSP("spASRIntInsertNewRecord" _
									, prmRecordID _
									, New SqlParameter("psInsertDef", SqlDbType.VarChar, -1) With {.Value = sInsertUpdateDef})

								lngRecordID = prmRecordID.Value

								If Len(sReaction) > 0 Then
									sAction = sReaction
								Else
									sAction = "LOAD"
								End If

								objDataAccess.ExecuteSP("spASREmailImmediate", New SqlParameter("@Username", SqlDbType.VarChar, 255) With {.Value = Session("Username")})

							Catch ex As SqlException
								If ex.Number.Equals(50000) Then
									sErrorMsg = Trim(Mid(ex.Message, 1, (InStr(ex.Message, "The transaction ended in the trigger")) - 1))
								Else
									sErrorMsg = sErrorMsg & FormatError(ex.Message)
								End If

								fOk = False

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
								objDataAccess.ExecuteSP("spASRIntUpdateRecord" _
									, prmResult _
									, New SqlParameter("psUpdateDef", SqlDbType.VarChar, -1) With {.Value = sInsertUpdateDef} _
									, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = NullSafeInteger(CleanNumeric(lngTableID))} _
									, New SqlParameter("psRealSource", SqlDbType.VarChar, 255) With {.Value = sRealSource} _
									, New SqlParameter("piID", SqlDbType.Int) With {.Value = CleanNumeric(lngRecordID)} _
									, New SqlParameter("piTimestamp", SqlDbType.Int) With {.Value = CleanNumeric(iTimestamp)})

								Select Case prmResult.Value
									Case 1 ' Record changed by another user, and is no longer in the current table/view.
										sErrorMsg = "The record has been amended by another user and will be refreshed."
									Case 2 ' Record changed by another user, and still in the current table/view.
										sErrorMsg = "The record has been amended by another user and will be refreshed."
									Case 3 ' Record deleted by another user.
										sErrorMsg = "The record has been deleted by another user."
								End Select

								If Len(sReaction) > 0 Then
									sAction = sReaction
								Else
									sAction = "LOAD"
								End If

								objDataAccess.ExecuteSP("spASREmailImmediate", _
										New SqlParameter("@Username", SqlDbType.VarChar, 255) With {.Value = Session("Username")})


							Catch ex As Exception

								sErrorMsg = sErrorMsg & FormatError(ex.Message)
								fOk = False

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
			ElseIf sAction = "DELETE" Then
				' Deleting.

				Try

					Dim prmResult As New SqlParameter("piResult", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
					objDataAccess.ExecuteSP("sp_ASRDeleteRecord" _
							, prmResult _
							, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = NullSafeInteger(CleanNumeric(lngTableID))} _
							, New SqlParameter("psRealSource", SqlDbType.VarChar, 255) With {.Value = sRealSource} _
							, New SqlParameter("piID", SqlDbType.Int) With {.Value = CleanNumeric(lngRecordID)})

					Select Case prmResult.Value
						Case 2 ' Record changed by another user, and is no longer in the current table/view.
							sErrorMsg = "The record has been amended by another user and will be refreshed."
					End Select

					lngRecordID = 0

					If Len(sReaction) > 0 Then
						sAction = sReaction
					Else
						sAction = "LOAD"
					End If

					objDataAccess.ExecuteSP("spASREmailImmediate" _
							, New SqlParameter("@Username", SqlDbType.VarChar, 255) With {.Value = Session("Username")})


				Catch ex As Exception
					sErrorMsg = "The record could not be deleted." & vbCrLf & FormatError(ex.Message)
					sAction = "SAVEERROR"


				End Try


			ElseIf sAction = "CANCELCOURSE" Then

				Try

					' Check number of bookings made.
					Dim prmNumberOfBookings = New SqlParameter("piNumberOfBookings", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
					Dim prmErrorMessage = New SqlParameter("psErrorMessage", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
					Dim prmCourseTitle = New SqlParameter("psCourseTitle", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

					objDataAccess.ExecuteSP("sp_ASRIntCancelCourse" _
						, prmNumberOfBookings _
						, New SqlParameter("piCourseRecordID", SqlDbType.Int) With {.Value = CleanNumeric(lngRecordID)} _
						, New SqlParameter("piTrainBookTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("TB_TBTableID"))} _
						, New SqlParameter("piCourseTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("TB_CourseTableID"))} _
						, New SqlParameter("piTrainBookStatusColumnID", SqlDbType.Int) With {.Value = CleanNumeric(Session("TB_TBStatusColumnID"))} _
						, New SqlParameter("psCourseRealSource", SqlDbType.VarChar, -1) With {.Value = sRealSource} _
					, prmErrorMessage _
					, prmCourseTitle)

					sAction = "CANCELCOURSE_1"
					Session("numberOfBookings") = prmNumberOfBookings.Value
					Session("tbErrorMessage") = prmErrorMessage.Value
					Session("tbCourseTitle") = prmCourseTitle.Value

				Catch ex As Exception
					sErrorMsg = "Error cancelling the course." & vbCrLf & FormatError(ex.Message)
					sAction = "SAVEERROR"

				End Try

			ElseIf sAction = "CANCELCOURSE_2" Then

				Try

					Dim prmErrorMessage = New SqlParameter("psErrorMessage", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

					objDataAccess.ExecuteSP("sp_ASRIntCancelCoursePart2" _
						, New SqlParameter("piEmployeeTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("TB_EmpTableID"))} _
						, New SqlParameter("piCourseTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("TB_CourseTableID"))} _
						, New SqlParameter("psCourseRealSource", SqlDbType.VarChar, -1) With {.Value = sRealSource} _
						, New SqlParameter("piCourseRecordID", SqlDbType.Int) With {.Value = CleanNumeric(lngRecordID)} _
						, New SqlParameter("piTransferCourseRecordID", SqlDbType.Int) With {.Value = CleanNumeric(iTBCourseRecordID)} _
						, New SqlParameter("piCourseCancelDateColumnID", SqlDbType.Int) With {.Value = CleanNumeric(Session("TB_CourseCancelDateColumnID"))} _
						, New SqlParameter("psCourseTitle", SqlDbType.VarChar, -1) With {.Value = Session("tbCourseTitle")} _
						, New SqlParameter("piTrainBookTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("TB_TBTableID"))} _
						, New SqlParameter("pfTrainBookTableInsert", SqlDbType.Bit) With {.Value = CleanBoolean(Session("TB_TBTableInsert"))} _
						, New SqlParameter("piTrainBookStatusColumnID", SqlDbType.Int) With {.Value = CleanNumeric(Session("TB_TBStatusColumnID"))} _
						, New SqlParameter("piTrainBookCancelDateColumnID", SqlDbType.Int) With {.Value = CleanNumeric(Session("TB_TBCancelDateColumnID"))} _
						, New SqlParameter("piWaitListTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("TB_WaitListTableID"))} _
						, New SqlParameter("pfWaitListTableInsert", SqlDbType.Bit) With {.Value = CleanBoolean(Session("TB_WaitListTableInsert"))} _
						, New SqlParameter("piWaitListCourseTitleColumnID", SqlDbType.Int) With {.Value = CleanNumeric(Session("TB_WaitListCourseTitleColumnID"))} _
						, New SqlParameter("pfWaitListCourseTitleColumnUpdate", SqlDbType.Bit) With {.Value = CleanBoolean(Session("TB_WaitListCourseTitleColumnUpdate"))} _
						, New SqlParameter("pfCreateWaitListRecords", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtTBCreateWLRecords"))} _
						, prmErrorMessage)

					sErrorMsg = prmErrorMessage.Value.ToString()

					If Len(sErrorMsg) > 0 Then
						sAction = "SAVEERROR"
					Else
						sAction = "LOAD"
					End If

				Catch ex As Exception
					sErrorMsg = "Error cancelling the course." & vbCrLf & FormatError(ex.Message)
					sAction = "SAVEERROR"

				End Try


			ElseIf sAction = "CANCELBOOKING" Then

				Try

					Dim prmErrorMessage = New SqlParameter("psErrorMessage", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

					objDataAccess.ExecuteSP("sp_ASRIntCancelBooking" _
						, New SqlParameter("pfTransferBookings", SqlDbType.Bit) With {.Value = CleanBoolean(fUserChoice)} _
						, New SqlParameter("piTBRecordID", SqlDbType.Int) With {.Value = CleanNumeric(lngRecordID)} _
						, prmErrorMessage)

					If Len(prmErrorMessage.Value.ToString()) > 0 Then
						sAction = "SAVEERROR"
					Else
						sAction = "CANCELBOOKING_1"
					End If

				Catch ex As Exception
					sErrorMsg = "Error cancelling the booking." & vbCrLf & FormatError(ex.Message)
					sAction = "SAVEERROR"

				End Try

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
			Return RedirectToAction("Data", "Home")

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

		Function Util_CalcSelection() As ActionResult
			Return View()
		End Function

		Function Util_SortOrderSelection() As ActionResult
			Return View()
		End Function

		<HttpPost()>
		Function LinksMain(Optional psScreenInfo As String = "") As ActionResult
			' Get dashboard items
			Dim sParameters As String = psScreenInfo

			Dim objSession = CType(Session("SessionContext"), SessionInfo)
			Dim objDataAccess As New clsDataAccess(objSession.LoginInfo)

			If sParameters.Length > 0 Then

				ResetSessionVars()

				Session("SSILinkTableID") = NullSafeInteger(Left(sParameters, InStr(1, sParameters, "!") - 1))
				Session("SSILinkViewID") = NullSafeInteger(Mid(sParameters, InStr(sParameters, "!") + 1, (InStr(sParameters, "_") - 1) - (InStr(sParameters, "!"))))

				If Mid(sParameters, InStr(sParameters, "_") + 1) = "" Then
					Session("TopLevelRecID") = 0
				Else
					Session("TopLevelRecID") = NullSafeInteger(Mid(sParameters, InStr(sParameters, "_") + 1))
				End If

			End If


			If (NullSafeInteger(Session("SSILinkTableID")) = NullSafeInteger(Session("SingleRecordTableID"))) _
				And (NullSafeInteger(Session("SSILinkViewID")) = NullSafeInteger(Session("SingleRecordViewID"))) Then

				' Ripped from AcctController
				Try
					' grab some more info for the dashboard						
					Dim sErrorDescription = ""

					' Get the self-service record ID.

					Dim prmRecordID = New SqlParameter("piRecordID", SqlDbType.Int)
					prmRecordID.Direction = ParameterDirection.Output

					Dim prmRecordCount = New SqlParameter("piRecordCount", SqlDbType.Int)
					prmRecordCount.Direction = ParameterDirection.Output

					objDataAccess.ExecuteSP("spASRIntGetSelfServiceRecordID", prmRecordID, prmRecordCount _
																, New SqlParameter("piViewID", CleanNumeric(Session("SingleRecordViewID"))))


					If prmRecordCount.Value = 1 Then
						' Only one record.
						Session("TopLevelRecID") = NullSafeInteger(prmRecordID.Value)
					Else
						If prmRecordCount.Value = 0 Then
							' No personnel record. 
							Session("TopLevelRecID") = 0
						Else
							' More than one personnel record.
							sErrorDescription = "You have access to more than one record in the defined Single-record view."

							Session("ErrorTitle") = "Login Page"
							Session("ErrorText") =
							 "You could not login to the OpenHR database because of the following reason:" & sErrorDescription & "<p>" & vbCrLf

							Response.Redirect("FormError")

						End If
					End If

					Err.Clear()

					' Are we displaying the Workflow Out of Office Hyperlink for this view?
					Dim lngSSILinkTableID As Short = Convert.ToInt16(Session("SingleRecordTableID"))
					Dim lngSSILinkViewID As Short = Convert.ToInt16(Session("SingleRecordViewID"))
					Dim fShowOOOHyperlink As Boolean = False

					Dim prmTableID2 = New SqlParameter("piTableID", SqlDbType.Int)
					prmTableID2.Value = lngSSILinkTableID

					Dim prmViewID2 = New SqlParameter("piViewID", SqlDbType.Int)
					prmViewID2.Value = lngSSILinkViewID

					Dim prmDisplayHyperlink = New SqlParameter("pfDisplayHyperlink", SqlDbType.Bit)
					prmDisplayHyperlink.Direction = ParameterDirection.Output

					objDataAccess.ExecuteSP("spASRIntShowOutOfOfficeHyperlink", prmTableID2, prmViewID2, prmDisplayHyperlink)

					If (Err.Number() <> 0) Then
						sErrorDescription = "Error getting the Workflow Out of Office hyperlink setting." & vbCrLf & FormatError(Err.Description)
					Else
						fShowOOOHyperlink = prmDisplayHyperlink.Value
					End If

					Session("WF_ShowOutOfOffice") = fShowOOOHyperlink

				Catch ex As Exception

					Session("ErrorTitle") = "Login Page"
					Session("ErrorText") =
					 "You could not login to the OpenHR database because of the following reason:" & vbCrLf & ex.Message & "<p>" & vbCrLf
					Response.Redirect("FormError")

				End Try
				' End Ripped
			End If

			Dim sViewDescription As String
			Dim sViewName As String = ""

			' For SSI, subordinate views
			If NullSafeInteger(Session("SSILinkViewID")) <> NullSafeInteger(Session("SingleRecordViewID")) Then

				Try

					' Get the record description.
					Dim prmRecordDesc = New SqlParameter("psRecDesc", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

					objDataAccess.ExecuteSP("sp_ASRIntGetRecordDescription" _
							, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("TableID"))} _
							, New SqlParameter("piRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("TopLevelRecID"))} _
							, New SqlParameter("piParentTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("parentTableID"))} _
							, New SqlParameter("piParentRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("parentRecordID"))} _
							, prmRecordDesc)

					sViewDescription = prmRecordDesc.Value

					If sViewDescription.Length > 0 Then
						sViewDescription = " view - " & prmRecordDesc.Value
					End If

					Dim rowViewName = objDataAccess.GetDataTable("SELECT viewname FROM asrsysviews WHERE viewid = " & Session("SSILinkViewID"), CommandType.Text)
					If rowViewName.Rows.Count > 0 Then
						sViewName = rowViewName(0)(0).ToString()
					End If

					' get the view name, and append it.
					If sViewName.Length > 0 Then sViewDescription = sViewName.Replace("_", " ") & sViewDescription

					Session("ViewDescription") = sViewDescription

				Catch ex As Exception
					Throw

				End Try

			Else
				Session("ViewDescription") = "My Dashboard"
			End If


			Dim objNavigation = New HR.Intranet.Server.clsNavigationLinks
			objNavigation.SessionInfo = CType(Session("SessionContext"), SessionInfo)

			objNavigation.ClearLinks()

			objNavigation.SSITableID = Session("SSILinkTableID")
			objNavigation.SSIViewID = Session("SSILinkViewID")
			objNavigation.LoadLinks()
			objNavigation.LoadNavigationLinks()

			Dim viewModel = New NavLinksViewModel With {.NavigationLinks = objNavigation.GetAllLinks, .NumberOfLinks = objNavigation.GetAllLinks.Count}

			Return View(viewModel)
		End Function

		' TODO
		Public Sub ShowPhoto(imageName As String)
			'TODO fetch path from registry
			Dim localImagesPath As String = HttpContext.Server.MapPath("~/pictures/profilephotos/")

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

		Public Sub ShowImageFromDb(imageID As String)

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
			Dim objRs As DataTable
			Dim image(-1) As Byte

			Try
				objRs = objDataAccess.GetFromSP("spASRIntGetPicture", _
						 New SqlParameter("piPictureID", SqlDbType.Int) With {.Value = CleanNumeric(imageID)})

				Dim objRow = objRs.Rows(0)

				image = CType(objRow(1), Byte())

				If image Is Nothing Then
					Throw New HttpException(404, "Image not found")
				End If

				' Check file extension to ensure correct MIME type.
				Dim imageExtension As String = Path.GetExtension(objRow(0).ToString()).ToLowerInvariant()
				Select Case imageExtension
					Case ".ico"
						Response.ContentType = "image/x-icon"
						Response.OutputStream.Write(image, 0, image.Length)

					Case ".bmp"
						Response.ContentType = "image/bmp"
						Response.OutputStream.Write(image, 0, image.Length)

					Case ".gif"
						Response.ContentType = "image/gif"
						Response.OutputStream.Write(image, 0, image.Length)

					Case ".jpg", ".jpeg"
						Response.ContentType = "image/jpeg"
						Response.OutputStream.Write(image, 0, image.Length)

					Case Else
						Response.ContentType = "image/bmp"
						Response.OutputStream.Write(image, 0, image.Length)

				End Select

			Catch ex As Exception
				' um...
			End Try


		End Sub

		Function GetChart(height As Long,
											width As Long,
											showLegend As Boolean,
											dottedGrid As Boolean,
											showValues As Boolean,
											stack As Boolean,
											showPercent As Boolean,
											chartType As Long,
											tableID As Long,
											columnID As Long,
											filterID As Long,
											aggregateType As Long,
											elementType As ElementType,
											sortOrderID As Long,
											sortDirection As Long,
											colourID As Long,
											title As String,
											showLabels As Boolean) As FileContentResult

			Err.Clear()

			Dim mrstChartData As DataTable
			Dim sErrorDescription As String

			Dim objChart = New HR.Intranet.Server.clsChart
			objChart.SessionInfo = CType(Session("SessionContext"), SessionInfo)

			mrstChartData = objChart.GetChartData(tableID, columnID, filterID, aggregateType, elementType, 0, 0, 0, 0, sortOrderID, sortDirection, colourID)

			If Err.Number <> 0 Then
				sErrorDescription = "The Chart field values could not be retrieved." & vbCrLf & FormatError(Err.Description)
			Else
				sErrorDescription = ""
			End If

			If Not mrstChartData Is Nothing Then
				If mrstChartData.Rows.Count > 500 Then mrstChartData = Nothing ' limit to 500 rows as get row buffer limit exceeded error.
			End If

			If Len(sErrorDescription) = 0 And Not mrstChartData Is Nothing Then

				If mrstChartData.Rows.Count > 0 Then
					Dim objRow1 = mrstChartData.Rows(0)

					If objRow1(0).ToString() <> "No Access" Then
						If objRow1(0).ToString() <> "No Data" Then

							Dim chart1 As New Chart()

							chart1.Width = Unit.Pixel(width)
							chart1.Height = Unit.Pixel(height)

							' Set Legend's visual attributes
							If showLegend = True Then
								chart1.Legends.Add("Default")
								chart1.Legends("Default").Enabled = True
								chart1.Legends("Default").BackColor = Color.Transparent
								chart1.Legends("Default").ShadowOffset = 2
								chart1.Legends("Default").BackColor = ColorTranslator.FromHtml("#D3DFF0")
							End If

							If Not String.IsNullOrEmpty(title) Then
								chart1.Titles.Add("MainTitle")
								chart1.Titles(0).Text = title
								chart1.Titles(0).Font = New Font(chart1.Titles(0).Font.Name, 20) 'Set the font size without changing the font family
							End If

							chart1.ChartAreas.Add("ChartArea1")

							chart1.ChartAreas("ChartArea1").BackColor = Color.FromArgb(64, 211, 211, 211)
							chart1.ChartAreas("ChartArea1").BackSecondaryColor = Color.Transparent
							chart1.ChartAreas("ChartArea1").ShadowColor = Color.Transparent

							chart1.ChartAreas("ChartArea1").AxisY.LineColor = Color.FromArgb(64, 64, 64, 64)
							chart1.ChartAreas("ChartArea1").AxisY.MajorGrid.LineColor = Color.FromArgb(64, 64, 64, 64)
							chart1.ChartAreas("ChartArea1").AxisX.LineColor = Color.FromArgb(64, 64, 64, 64)
							chart1.ChartAreas("ChartArea1").AxisX.MajorGrid.LineColor = Color.FromArgb(64, 64, 64, 64)

							chart1.ChartAreas("ChartArea1").AxisX.LabelStyle.Enabled = showLabels
							chart1.ChartAreas("ChartArea1").AxisY.LabelStyle.Enabled = showLabels

							' Gridlines
							If dottedGrid = True Then
								chart1.ChartAreas("ChartArea1").AxisX.LineDashStyle = ChartDashStyle.Dot
								chart1.ChartAreas("ChartArea1").AxisY.LineDashStyle = ChartDashStyle.Dot
								chart1.ChartAreas("ChartArea1").AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dot
								chart1.ChartAreas("ChartArea1").AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot
							Else
								chart1.ChartAreas("ChartArea1").AxisX.LineDashStyle = ChartDashStyle.NotSet
								chart1.ChartAreas("ChartArea1").AxisY.LineDashStyle = ChartDashStyle.NotSet
								chart1.ChartAreas("ChartArea1").AxisX.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
								chart1.ChartAreas("ChartArea1").AxisY.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
							End If


							If chartType = 0 Or chartType = 2 Or chartType = 4 Or chartType = 6 Or chartType = 14 Then
								' 3D Settings
								chart1.ChartAreas("ChartArea1").Area3DStyle.Enable3D = True
								chart1.ChartAreas("ChartArea1").Area3DStyle.Perspective = 10
								chart1.ChartAreas("ChartArea1").Area3DStyle.Inclination = 15
								chart1.ChartAreas("ChartArea1").Area3DStyle.Rotation = 10
								chart1.ChartAreas("ChartArea1").Area3DStyle.IsRightAngleAxes = False
								chart1.ChartAreas("ChartArea1").Area3DStyle.WallWidth = 0
								chart1.ChartAreas("ChartArea1").Area3DStyle.IsClustered = False
							End If

							' Series - just one series as multiaxis = false.
							chart1.Series.Add("Default")
							chart1.Series("Default").IsVisibleInLegend = False
							chart1.Series("Default").BorderColor = Color.FromArgb(180, 26, 59, 105)
							chart1.Series("Default").Color = Color.FromArgb(220, 65, 140, 240)

							' Show Values/Percentages
							If showValues = True Then
								chart1.Series("Default")("LabelStyle") = "Top"
								chart1.Series("Default").IsValueShownAsLabel = True

								If showPercent = True Then
									chart1.Series("Default").Label = "#PERCENT{P2}"
								End If
							End If

							Select Case chartType
								Case 0, 1
									If stack = True Then
										chart1.Series("Default").ChartType = SeriesChartType.StackedColumn
									Else
										chart1.Series("Default").ChartType = SeriesChartType.Column
									End If

								Case 2, 3
									chart1.Series("Default").ChartType = SeriesChartType.Line
								Case 4, 5
									If stack = True Then
										chart1.Series("Default").ChartType = SeriesChartType.StackedArea
									Else
										chart1.Series("Default").ChartType = SeriesChartType.Area
									End If

								Case 6, 7
									chart1.Series("Default").ChartType = SeriesChartType.StepLine
								Case 14
									chart1.Series("Default").ChartType = SeriesChartType.Pie

									chart1.ChartAreas("ChartArea1").BackColor = Color.Transparent
									chart1.ChartAreas("ChartArea1").BackSecondaryColor = Color.Transparent
									chart1.ChartAreas("ChartArea1").ShadowColor = Color.Transparent

							End Select

							'See Color Palette details here:http://blogs.msdn.com/b/alexgor/archive/2009/10/06/setting-chart-series-colors.aspx
							Dim brightPastelColorPalette As Integer() = {15764545, 4306172, 671968, 9593861, 12566463, 6896410, 8578047, 14523410, 4942794, 14375936, 8966899, 8479568, 11057649, 689120, 12489592}
							Dim pointNum As Integer

							For Each objRow As DataRow In mrstChartData.Rows
								If objRow(0).ToString() <> "No Access" And objRow(0).ToString() <> "No Data" Then

									Dim pointBackColor As Color
									If objRow(2) = 16777215 Then
										pointBackColor = ColorTranslator.FromWin32(brightPastelColorPalette(pointNum Mod 15))
									Else
										Try
											pointBackColor = ColorTranslator.FromWin32(objRow(2))
										Catch ex As Exception
											pointBackColor = ColorTranslator.FromWin32(brightPastelColorPalette(pointNum Mod 15))
										End Try
									End If

									If showLabels Then
										chart1.Series("Default").Points.Add(New DataPoint() With {.AxisLabel = objRow(0), .YValues = New Double() {objRow(1)}, .Color = pointBackColor})
									Else
										chart1.Series("Default").Points.Add(New DataPoint() With {.Label = " ", .YValues = New Double() {objRow(1)}, .Color = pointBackColor})
									End If

									If showLegend = True Then
										chart1.Legends("Default").CustomItems.Add(New LegendItem(objRow(0), pointBackColor, ""))
									End If
								End If

								pointNum += 1
							Next

							Using ms = New MemoryStream()
								chart1.SaveImage(ms, ChartImageFormat.Png)
								ms.Seek(0, SeekOrigin.Begin)

								Return File(ms.ToArray(), "image/png", "mychart.png")
							End Using
						Else
							' No Data						
						End If
					Else
						' No Access
					End If

				End If

			End If

		End Function



		Function GetMultiAxisChart(height As Long,
											width As Long,
											showLegend As Boolean,
											dottedGrid As Boolean,
											showValues As Boolean,
											stack As Boolean,
											showPercent As Boolean,
											chartType As Long,
											tableID As Long,
											columnID As Long,
											filterID As Long,
											aggregateType As Long,
											elementType As ElementType,
											tableID_2 As Long,
											columnID_2 As Long,
											tableID_3 As Long,
											columnID_3 As Long,
											sortOrderID As Long,
											sortDirection As Long,
											colourID As Long,
											title As String,
											showLabels As Boolean) As FileContentResult

			Err.Clear()

			Dim RotateX As Integer = HttpContext.Request.QueryString("rotateX")
			If RotateX = 0 Then RotateX = 15
			Dim RotateY As Integer = HttpContext.Request.QueryString("rotateY")
			If RotateY = 0 Then RotateY = 10

			Dim mrstChartData As DataTable
			Dim sErrorDescription As String

			Dim objChart = New HR.Intranet.Server.clsMultiAxisChart
			objChart.SessionInfo = CType(Session("SessionContext"), SessionInfo)

			' Pass required info to the DLL
			mrstChartData = objChart.GetChartData(tableID, columnID, filterID, aggregateType, elementType, tableID_2, columnID_2, tableID_3, columnID_3, sortOrderID, sortDirection, colourID)

			If Err.Number <> 0 Then
				sErrorDescription = "The Chart field values could not be retrieved." & vbCrLf & FormatError(Err.Description)
			Else
				sErrorDescription = ""
			End If

			If Not mrstChartData Is Nothing Then
				If mrstChartData.Rows.Count > 500 Then mrstChartData = Nothing ' limit to 500 rows as get row buffer limit exceeded error.
			End If

			If Len(sErrorDescription) = 0 And Not mrstChartData Is Nothing Then
				Dim seriesName As String

				If mrstChartData.Rows.Count > 0 Then

					Dim objRow1 = mrstChartData.Rows(0)

					If TryCast(objRow1(0), String) <> "No Access" Then
						If TryCast(objRow1(0), String) <> "No Data" Then
							MultiAxisChart.Width = Unit.Pixel(width)
							MultiAxisChart.Height = Unit.Pixel(height)

							' Set Legend's visual attributes
							If showLegend = True Then
								MultiAxisChart.Legends.Add("Default")
								MultiAxisChart.Legends("Default").Enabled = True
								MultiAxisChart.Legends("Default").BackColor = Color.Transparent
								MultiAxisChart.Legends("Default").ShadowOffset = 2
								MultiAxisChart.Legends("Default").BackColor = ColorTranslator.FromHtml("#D3DFF0")
							End If

							If Not String.IsNullOrEmpty(title) Then
								MultiAxisChart.Titles.Add("MainTitle")
								MultiAxisChart.Titles(0).Text = title
								MultiAxisChart.Titles(0).Font = New Font(MultiAxisChart.Titles(0).Font.Name, 20) 'Set the font size without changing the font family
							End If

							seriesName = "Default"

							MultiAxisChart.ChartAreas.Add(seriesName)

							MultiAxisChart.ChartAreas(seriesName).BackColor = Color.FromArgb(64, 211, 211, 211)
							MultiAxisChart.ChartAreas(seriesName).BackSecondaryColor = Color.Transparent
							MultiAxisChart.ChartAreas(seriesName).ShadowColor = Color.Transparent
							MultiAxisChart.ChartAreas(seriesName).AxisY.LineColor = Color.FromArgb(64, 64, 64, 64)
							MultiAxisChart.ChartAreas(seriesName).AxisY.MajorGrid.LineColor = Color.FromArgb(64, 64, 64, 64)
							MultiAxisChart.ChartAreas(seriesName).AxisX.LineColor = Color.FromArgb(64, 64, 64, 64)
							MultiAxisChart.ChartAreas(seriesName).AxisX.MajorGrid.LineColor = Color.FromArgb(64, 64, 64, 64)

							MultiAxisChart.ChartAreas(seriesName).AxisX.LabelStyle.Enabled = showLabels
							MultiAxisChart.ChartAreas(seriesName).AxisY.LabelStyle.Enabled = showLabels

							' Gridlines
							If dottedGrid = True Then
								MultiAxisChart.ChartAreas(seriesName).AxisX.LineDashStyle = ChartDashStyle.Dot
								MultiAxisChart.ChartAreas(seriesName).AxisY.LineDashStyle = ChartDashStyle.Dot
								MultiAxisChart.ChartAreas(seriesName).AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dot
								MultiAxisChart.ChartAreas(seriesName).AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot
							Else
								MultiAxisChart.ChartAreas(seriesName).AxisX.LineDashStyle = ChartDashStyle.NotSet
								MultiAxisChart.ChartAreas(seriesName).AxisY.LineDashStyle = ChartDashStyle.NotSet
								MultiAxisChart.ChartAreas(seriesName).AxisX.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
								MultiAxisChart.ChartAreas(seriesName).AxisY.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
							End If

							' 3D Settings
							If chartType = 0 Or chartType = 2 Or chartType = 4 Or chartType = 6 Or chartType = 14 Then
								MultiAxisChart.ChartAreas(seriesName).Area3DStyle.Enable3D = True
								MultiAxisChart.ChartAreas(seriesName).Area3DStyle.Perspective = 10
								MultiAxisChart.ChartAreas(seriesName).Area3DStyle.Inclination = RotateX
								MultiAxisChart.ChartAreas(seriesName).Area3DStyle.Rotation = RotateY
								MultiAxisChart.ChartAreas(seriesName).Area3DStyle.IsRightAngleAxes = False
								MultiAxisChart.ChartAreas(seriesName).Area3DStyle.WallWidth = 0
								MultiAxisChart.ChartAreas(seriesName).Area3DStyle.IsClustered = False
							End If

							Dim seriesNames As String = ""
							'See Color Palette details here:http://blogs.msdn.com/b/alexgor/archive/2009/10/06/setting-chart-series-colors.aspx
							Dim brightPastelColorPalette As Integer() = {15764545, 4306172, 671968, 9593861, 12566463, 6896410, 8578047, 14523410, 4942794, 14375936, 8966899, 8479568, 11057649, 689120, 12489592}
							Dim pointNum As Integer

							'Fill missing data
							Dim i As Integer
							Dim j As Integer
							Dim r As DataRow

							'Determine the verticals and horizontals we have in the datatable; once we get them, we can fill the missing gaps in the data

							'Verticals
							Dim MultiAxisChartVerticals As New List(Of MultiAxisChartVertical)
							Dim MinVerticalID As Integer = Convert.ToInt32(mrstChartData.Compute("min(VERTICAL_ID)", String.Empty))	'Get the minimum vertical ID
							Dim MaxVerticalID As Integer = Convert.ToInt32(mrstChartData.Compute("max(VERTICAL_ID)", String.Empty))	'Get the maximum vertical ID
							For i = MinVerticalID To MaxVerticalID
								r = mrstChartData.Select("VERTICAL_ID = " & i).FirstOrDefault
								MultiAxisChartVerticals.Add(New MultiAxisChartVertical With {.Vertical_ID = r("VERTICAL_ID"), .Vertical = r("VERTICAL")})
							Next

							'Horizontals
							Dim MultiAxisChartHorizontals As New List(Of MultiAxisChartHorizontal)
							Dim MinHorizontalID As Integer = Convert.ToInt32(mrstChartData.Compute("min(HORIZONTAL_ID)", String.Empty))	'Get the minimum horizontal ID
							Dim MaxHorizontalID As Integer = Convert.ToInt32(mrstChartData.Compute("max(HORIZONTAL_ID)", String.Empty))	'Get the maximum horizontal ID
							For i = MinHorizontalID To MaxHorizontalID
								r = mrstChartData.Select("HORIZONTAL_ID = " & i).FirstOrDefault
								MultiAxisChartHorizontals.Add(New MultiAxisChartHorizontal With {.Horizontal_ID = r("HORIZONTAL_ID"), .Horizontal = r("HORIZONTAL"), .Colour = r("COLOUR")})
							Next

							'Compare and fill the gaps
							Dim newRow As DataRow
							Dim Vertical As MultiAxisChartVertical
							Dim Horizontal As MultiAxisChartHorizontal
							For i = 0 To MultiAxisChartHorizontals.Count - 1
								Horizontal = MultiAxisChartHorizontals(i)
								For j = 0 To MultiAxisChartVerticals.Count - 1
									Vertical = MultiAxisChartVerticals(j)
									If mrstChartData.Select("HORIZONTAL_ID = " & Horizontal.Horizontal_ID & " AND VERTICAL_ID = " & Vertical.Vertical_ID).FirstOrDefault Is Nothing Then 'This combination doesn't exist...
										'Insert a new row
										newRow = mrstChartData.NewRow
										newRow("HORIZONTAL_ID") = Horizontal.Horizontal_ID
										newRow("HORIZONTAL") = Horizontal.Horizontal
										newRow("VERTICAL_ID") = Vertical.Vertical_ID
										newRow("VERTICAL") = Vertical.Vertical
										newRow("Aggregate") = 0
										newRow("COLOUR") = Horizontal.Colour
										mrstChartData.Rows.Add(newRow)
									End If
								Next
							Next

							Dim dv As DataView = mrstChartData.AsDataView	'Copy the datatable to a dataview so we can sort it
							dv.Sort = "VERTICAL_ID DESC, HORIZONTAL_ID ASC"	'Sort
							For Each objRow As DataRow In dv.ToTable.Rows	'Loop over the dataview's rows
								If TryCast(objRow("HORIZONTAL_ID"), String) <> "No Access" And TryCast(objRow("HORIZONTAL_ID"), String) <> "No Data" Then
									seriesName = objRow("VERTICAL").ToString()
									If seriesName = "" Then
										seriesName = "(No name)"
									End If
									Dim columnName As String = objRow("HORIZONTAL").ToString()
									Dim yVal As Integer = CInt(objRow("Aggregate"))
									Dim pointBackColor As Color
									If objRow("COLOUR") = 16777215 Then
										pointBackColor = ColorTranslator.FromWin32(brightPastelColorPalette(pointNum Mod 15))
									Else
										Try
											pointBackColor = ColorTranslator.FromWin32(objRow("COLOUR"))
										Catch ex As Exception
											pointBackColor = ColorTranslator.FromWin32(brightPastelColorPalette(pointNum Mod 15))
										End Try
									End If

									If Not seriesNames.Contains("<" & seriesName & ">") Then
										' Add the series - ONLY if not already added.
										MultiAxisChart.Series.Add(seriesName)
										MultiAxisChart.Series(seriesName).IsVisibleInLegend = False

										seriesNames &= "<" & seriesName & ">"

										' Show Values/Percentages
										If showValues = True Then
											MultiAxisChart.Series(seriesName)("LabelStyle") = "Top"
											MultiAxisChart.Series(seriesName).IsValueShownAsLabel = True

											If showPercent = True Then
												MultiAxisChart.Series(seriesName).Label = "#PERCENT{P2}"
											End If
										End If

										Select Case chartType
											Case 0, 1
												If stack = True Then
													MultiAxisChart.Series(seriesName).ChartType = SeriesChartType.StackedColumn
												Else
													MultiAxisChart.Series(seriesName).ChartType = SeriesChartType.Column
												End If

											Case 2, 3
												MultiAxisChart.Series(seriesName).ChartType = SeriesChartType.Line
											Case 4, 5
												If stack = True Then
													MultiAxisChart.Series(seriesName).ChartType = SeriesChartType.StackedArea
												Else
													MultiAxisChart.Series(seriesName).ChartType = SeriesChartType.Area
												End If

											Case 6, 7
												MultiAxisChart.Series(seriesName).ChartType = SeriesChartType.StepLine
											Case 14
												MultiAxisChart.Series(seriesName).ChartType = SeriesChartType.Pie
										End Select
									End If

									If showLabels Then
										MultiAxisChart.Series(seriesName).Points.Add(New DataPoint() With {
																																											 .AxisLabel = columnName,
																																											 .YValues = New Double() {yVal},
																																											 .Color = pointBackColor,
																																											 .IsEmpty = (yVal = 0)
																																											 })
									Else
										MultiAxisChart.Series(seriesName).Points.Add(New DataPoint() With {
																																										 .Label = " ",
																																										 .YValues = New Double() {yVal},
																																										 .Color = pointBackColor,
																																										 .IsEmpty = (yVal = 0)
																																										 })
									End If

									If showLegend = True Then
										Dim legendAdded As Boolean = False
										For Each legItem As LegendItem In MultiAxisChart.Legends("Default").CustomItems
											If legItem.Name = objRow("HORIZONTAL") Then legendAdded = True
										Next

										If Not legendAdded Then
											MultiAxisChart.Legends("Default").CustomItems.Add(New LegendItem(objRow("HORIZONTAL"), pointBackColor, ""))
										End If
									End If

								End If
								pointNum += 1
							Next

							If showLabels Then
								MultiAxisChart.ChartAreas("Default").AxisX.Interval = 1	'Show all X axis legends (labels?)
								MultiAxisChart.AlignDataPointsByAxisLabel()
							End If

							'Make all the datapoints semi-transparent							
							MultiAxisChart.ApplyPaletteColors()
							For Each series As Series In MultiAxisChart.Series
								For Each point As DataPoint In series.Points
									point.Color = Color.FromArgb(180, point.Color)
								Next
							Next

							Using ms = New MemoryStream()
								MultiAxisChart.SaveImage(ms, ChartImageFormat.Png)
								ms.Seek(0, SeekOrigin.Begin)

								Return File(ms.ToArray(), "image/png", "mychart.png")
							End Using
						Else
							' No Data
						End If
					Else
						' No access
					End If
				End If
			End If

		End Function

		Function PasswordChange() As ActionResult
			Return View()
		End Function

		Function NewUser() As ActionResult
			Return View()
		End Function

		'Function ForcePasswordChange() As ActionResult
		'    Return View()
		'End Function

		Function Poll() As PartialViewResult
			'Return PartialView()
		End Function

#Region "Event Log Forms"

		Function emailSelection() As ActionResult
			Return View()
		End Function

		Function EventLog() As ActionResult
			Return View()
		End Function

		Function EventLogDetails() As ActionResult
			Return View()
		End Function

		Function EventLogPurge() As ActionResult
			Return View()
		End Function

		Function EventLogSelection() As ActionResult
			Return View()
		End Function

#End Region

#Region "Running Reports"

		Function util_run_crosstabsMain() As ActionResult
			Return PartialView()
		End Function

		Function util_run_crosstabsData() As ActionResult
			Return PartialView()
		End Function

		Function util_run_crosstabsBreakdown() As ActionResult
			Return PartialView()
		End Function

		Function util_run_crosstabs() As ActionResult
			Return PartialView()
		End Function

		<HttpPost()>
		Function util_run_crosstabsDataSubmit()

			On Error Resume Next

			Session("CT_Mode") = Request.Form("txtMode")
			Session("CT_EmailGroupID") = Request.Form("txtEmailGroupID")
			Session("CT_EmailGroupAddr") = Request.Form("txtEmailGroupAddr")
			Session("CT_UtilID") = Request.Form("txtUtilID")

			If Session("CT_Mode") = "BREAKDOWN" Then
				Session("CT_Hor") = Request.Form("txtHor")
				Session("CT_Ver") = Request.Form("txtVer")
				Session("CT_Pgb") = Request.Form("txtPgb")
				Session("CT_IntersectionType") = Request.Form("txtIntersectionType")
				Session("CT_CellValue") = Request.Form("txtCellValue")
				Session("CT_Use1000") = Request.Form("txtUse1000")
			Else
				Session("CT_PageNumber") = Request.Form("txtPageNumber")
				Session("CT_IntersectionType") = Request.Form("txtIntersectionType")
				Session("CT_ShowPercentage") = Request.Form("txtShowPercentage")
				Session("CT_PercentageOfPage") = Request.Form("txtPercentageOfPage")
				Session("CT_SuppressZeros") = Request.Form("txtSuppressZeros")
				Session("CT_Use1000") = Request.Form("txtUse1000")
			End If

			' Go to the requested page.
			Return RedirectToAction("util_run_crosstabsData")

		End Function

		<ValidateInput(False)>
		Function util_run_promptedvalues() As ActionResult

			Session("utiltype") = Request.Form("utiltype")
			Session("utilid") = Request.Form("utilid")
			Session("utilname") = Request.Form("utilname")
			Session("action") = Request.Form("action")
			Session("MailMerge_Template") = Nothing

			Return View()
		End Function

		<HttpPost()>
		Function util_run_uploadtemplate(TemplateFile As HttpPostedFileBase) As ActionResult

			Try

				If Not TemplateFile Is Nothing Then
					' Read input stream from request
					Dim Buffer = New Byte(TemplateFile.InputStream.Length - 1) {}
					Dim offset As Integer = 0
					Dim cnt As Integer = 0
					While (InlineAssignHelper(cnt, TemplateFile.InputStream.Read(Buffer, offset, 10))) > 0
						offset += cnt
					End While

					Session("MailMerge_Template") = New MemoryStream(Buffer)

				End If

			Catch ex As Exception
				Session("ErrorTitle") = "File upload"
				Session("ErrorText") = "You could not upload the template file because of the following error:<p>" & FormatError(ex.Message)
			End Try

			Return Content("hello ducky")

		End Function

		<HttpPost()>
		Function util_run_promptedvalues_submit(TemplateFile As HttpPostedFileBase) As ActionResult

			'		Try

			'For Each ob As HttpPostedFile In Request.Files

			'	Dim helloducky = ob

			'Next


			'	Dim blah = CType(Request.Form("TemplateFile"), HttpPostedFile)

			'	If Not TemplateFile Is Nothing Then
			'		' Read input stream from request
			'		Dim Buffer = New Byte(TemplateFile.InputStream.Length - 1) {}
			'		Dim offset As Integer = 0
			'		Dim cnt As Integer = 0
			'		While (InlineAssignHelper(cnt, TemplateFile.InputStream.Read(Buffer, offset, 10))) > 0
			'			offset += cnt
			'		End While

			'		Session("MailMerge_Template") = New MemoryStream(Buffer)

			'	End If

			'Catch ex As Exception
			'	Session("ErrorTitle") = "File upload"
			'	Session("ErrorText") = "You could not upload the template file because of the following error:<p>" & FormatError(ex.Message)
			'End Try

			Return View("util_run")
		End Function

		<ValidateInput(False)>
		Function util_run() As ActionResult
			Return PartialView()
		End Function

		<ValidateInput(False)>
		Function util_run_customreports() As ActionResult
			Return PartialView()
		End Function

		Function util_run_calendarreport_main() As ActionResult
			Return PartialView()
		End Function

		Public Function util_run_crosstab_downloadoutput() As FilePathResult

			Dim lngFormat As OutputFormats = Request("txtFormat")
			Dim blnScreen As Boolean = False
			Dim blnPrinter As Boolean = Request("txtPrinter")
			Dim strPrinterName As String = Request("txtPrinterName")
			Dim blnSave As Boolean = Request("txtSave")
			Dim lngSaveExisting As Long = Request("txtSaveExisting")
			Dim blnEmail As Boolean = Request("txtEmail")
			Dim lngEmailGroupID As Integer = Request("txtEmailGroupID")
			Dim strEmailSubject As String = Request("txtEmailSubject")
			Dim strEmailAttachAs As String = Request("txtEmailAttachAs")
			Dim strDownloadFileName As String = Request("txtFilename")
			Dim strDownloadExtension As String
			Dim strInterSectionType As String

			Dim lngLoopMin As Long
			Dim lngLoopMax As Long

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
			Dim objCrossTab As CrossTab = CType(Session("objCrossTab" & Session("UtilID")), CrossTab)

			Dim ClientDLL As New HR.Intranet.Server.clsOutputRun
			ClientDLL.SessionInfo = CType(Session("SessionContext"), SessionInfo)

			Dim objUser As New HR.Intranet.Server.clsSettings
			objUser.SessionInfo = CType(Session("SessionContext"), SessionInfo)

			ClientDLL.SaveAsValues = Session("OfficeSaveAsValues").ToString()

			ClientDLL.SettingOptions(objUser.GetUserSetting("Output", "WordTemplate", "").ToString() _
				, objUser.GetUserSetting("Output", "ExcelTemplate", "").ToString() _
				, CBool(objUser.GetUserSetting("Output", "ExcelGridlines", False)) _
				, CBool(objUser.GetUserSetting("Output", "ExcelHeaders", False)) _
				, CBool(objUser.GetUserSetting("Output", "ExcelOmitSpacerRow", False)) _
				, CBool(objUser.GetUserSetting("Output", "ExcelOmitSpacerCol", False)) _
				, CBool(objUser.GetUserSetting("Output", "AutoFitCols", True)) _
				, CBool(objUser.GetUserSetting("Output", "Landscape", True)) _
				, False)

			ClientDLL.SettingLocations(CInt(objUser.GetUserSetting("Output", "TitleCol", 3)) _
				, CInt(objUser.GetUserSetting("Output", "TitleRow", 2)) _
				, CInt(objUser.GetUserSetting("Output", "DataCol", 2)) _
				, CInt(objUser.GetUserSetting("Output", "DataRow", 4)))

			ClientDLL.SettingTitle(CBool(objUser.GetUserSetting("Output", "TitleGridLines", False)) _
				, CBool(objUser.GetUserSetting("Output", "TitleBold", True)) _
				, CBool(objUser.GetUserSetting("Output", "TitleUnderline", False)) _
				, CInt(objUser.GetUserSetting("Output", "TitleBackcolour", "16777215")) _
				, CInt(objUser.GetUserSetting("Output", "TitleForecolour", "6697779")) _
				, objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "TitleBackcolour", "16777215"))) _
				, objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "TitleForecolour", "6697779"))))

			ClientDLL.SettingHeading(CBool(objUser.GetUserSetting("Output", "HeadingGridLines", True)) _
				, CBool(objUser.GetUserSetting("Output", "HeadingBold", True)) _
				, CBool(objUser.GetUserSetting("Output", "HeadingUnderline", False)) _
				, CInt(objUser.GetUserSetting("Output", "HeadingBackcolour", 16248553)) _
				, CInt(objUser.GetUserSetting("Output", "HeadingForecolour", 6697779)) _
				, CInt(objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "HeadingBackcolour", 16248553)))) _
				, CInt(objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "HeadingForecolour", 6697779)))))

			ClientDLL.SettingData(CBool(objUser.GetUserSetting("Output", "DataGridLines", True)) _
				, CBool(objUser.GetUserSetting("Output", "DataBold", False)) _
				, CBool(objUser.GetUserSetting("Output", "DataUnderline", False)) _
				, CInt(objUser.GetUserSetting("Output", "DataBackcolour", 15988214)) _
				, CInt(objUser.GetUserSetting("Output", "DataForecolour", 6697779)) _
				, CInt(objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "DataBackcolour", 15988214)))) _
				, CInt(objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "DataForecolour", 6697779)))))

			ClientDLL.InitialiseStyles()
			ClientDLL.HeaderCols = 1

			'Set Options
			If Not objCrossTab.OutputPreview Then
				lngFormat = objCrossTab.OutputFormat
				blnScreen = objCrossTab.OutputScreen
				blnPrinter = objCrossTab.OutputPrinter
				strPrinterName = objCrossTab.OutputPrinterName
				blnSave = objCrossTab.OutputSave
				lngSaveExisting = objCrossTab.OutputSaveExisting
				blnEmail = objCrossTab.OutputEmail
				lngEmailGroupID = objCrossTab.OutputEmailID
				strEmailSubject = objCrossTab.OutputEmailSubject
				strEmailAttachAs = objCrossTab.OutputEmailAttachAs
				strDownloadFileName = objCrossTab.DownloadFileName
			End If

			If strDownloadFileName.Length = 0 Then
				objCrossTab.OutputFormat = lngFormat
				objCrossTab.OutputFilename = ""
				strDownloadFileName = objCrossTab.DownloadFileName
			End If

			strDownloadExtension = Path.GetExtension(strDownloadFileName)

			Dim fOK = ClientDLL.SetOptions(False, lngFormat, False, False, strPrinterName, True, lngSaveExisting _
				, blnEmail, lngEmailGroupID, strEmailSubject, strEmailAttachAs, strDownloadExtension)

			If fOK Then
				If ClientDLL.GetFile() Then
					If lngFormat = OutputFormats.fmtDataOnly Then

					ElseIf lngFormat = OutputFormats.fmtExcelPivotTable Then

						'Response.Write("  ClientDLL.PivotSuppressBlanks = (window.chkSuppressZeros.checked == true);" & vbCrLf)
						'Response.Write("  ClientDLL.PivotDataFunction = window.cboIntersectionType.options[window.cboIntersectionType.selectedIndex].text;" & vbCrLf)

						ClientDLL.AddColumn(" ", SQLDataType.sqlVarChar, 0, False)
						For intCount = 0 To objCrossTab.ColumnHeadingUbound(0)
							ClientDLL.AddColumn(objCrossTab.ColumnHeading(0, intCount), SQLDataType.sqlVarChar, objCrossTab.IntersectionDecimals, objCrossTab.Use1000Separator)
						Next

						'Response.Write("  ClientDLL.AddColumn(window.cboIntersectionType.options[window.cboIntersectionType.selectedIndex].text, 2, " & objCrossTab.IntersectionDecimals & "," & LCase(objCrossTab.Use1000Separator) & ");" & vbCrLf)
						strInterSectionType = "TODO Intersection type"
						ClientDLL.AddColumn(strInterSectionType, SQLDataType.sqlInteger, objCrossTab.IntersectionDecimals, objCrossTab.Use1000Separator)

						Dim rsPivot As DataTable
						Dim strSQL As String

						Dim strOutput(,) As String
						Dim strPageValue As String = ""
						Dim lngGroupNum As Integer
						Dim lngCol As Integer
						Dim lngRow As Integer


						strSQL = "SELECT HOR as 'Horizontal', VER as 'Vertical'" & IIf(objCrossTab.PageBreakColumn, ", PGB as 'Page Break'", vbNullString) & ", RecDesc as 'Record Description'" & IIf(objCrossTab.IntersectionColumn, ", Ins as 'Intersection'", vbNullString) & IIf(objCrossTab.CrossTabType = CrossTabType.cttAbsenceBreakdown, ", Value as 'Duration'", vbNullString) & " FROM " & objCrossTab.mstrTempTableName

						If lngFormat = CrossTabType.cttAbsenceBreakdown Then
							strSQL = strSQL & " WHERE NOT HOR IN ('Total','Count','Average')"

						ElseIf objCrossTab.PageBreakColumn Then
							strSQL = strSQL & " ORDER BY PGB"
						End If

						rsPivot = objDataAccess.GetDataTable(strSQL)

						'------------

						With rsPivot

							'			ReDim mstrOutputPivotArray(0)

							If Not objCrossTab.PageBreakColumn Then
								lngRow = 1
								ReDim strOutput(.Columns.Count, 0)
								For lngCol = 0 To .Columns.Count - 1
									strOutput(lngCol, 0) = rsPivot.Columns(lngCol).ColumnName
								Next
							End If

							For Each objRow As DataRow In rsPivot.Rows

								If objCrossTab.PageBreakColumn Then
									If strPageValue <> objRow("Page Break") Then

										If strPageValue <> vbNullString Then

											ClientDLL.AddPage(objCrossTab.Name, strPageValue)
											ClientDLL.ArrayDim(UBound(strOutput, 1), UBound(strOutput, 2))
											For lngCol = 0 To UBound(strOutput, 1)
												For lngRow = 0 To UBound(strOutput, 2)
													ClientDLL.ArrayAddTo(lngCol, lngRow, strOutput(lngCol, lngRow))
												Next
											Next

											ClientDLL.DataArray()

										End If
										strPageValue = objRow("Page Break").ToString()

										lngRow = 1
										ReDim strOutput(.Columns.Count - 1, 0)
										For lngCol = 0 To .Columns.Count - 1
											strOutput(lngCol, 0) = objRow(lngCol).ColumnName
										Next

									End If
								Else
									strPageValue = objCrossTab.BaseTableName

								End If

								ReDim Preserve strOutput(.Columns.Count, lngRow)
								For lngCol = 0 To .Columns.Count - 1

									If lngCol < 2 Or (lngCol = 2 And objCrossTab.PageBreakColumn) Then

										'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
										'UPGRADE_WARNING: Couldn't resolve default property of object GetGroupNumber(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										lngGroupNum = objCrossTab.GetGroupNumber(CStr(IIf(IsDBNull(objRow(lngCol)), vbNullString, objRow(lngCol))), CShort(lngCol))
										'UPGRADE_WARNING: Couldn't resolve default property of object mvarHeadings()(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										strOutput(lngCol, lngRow) = objCrossTab.ColumnHeading(lngCol, lngGroupNum)
									Else
										strOutput(lngCol, lngRow) = objRow(lngCol)
									End If
								Next
								lngRow += 1
							Next
						End With

						ClientDLL.AddPage(objCrossTab.Name, strPageValue)

						ClientDLL.ArrayDim(UBound(strOutput, 1), UBound(strOutput, 2))
						For lngCol = 0 To UBound(strOutput, 1)
							For lngRow = 0 To UBound(strOutput, 2) - 1
								ClientDLL.ArrayAddTo(lngCol, lngRow, strOutput(lngCol, lngRow))
							Next
						Next

						ClientDLL.DataArray()
						ClientDLL.Complete()

					Else


						''MH20040219
						'Response.Write("  var lngExcelDataType;")
						'Response.Write("  if (window.chkPercentType.checked == true) {" & vbCrLf)
						'Response.Write("    lngExcelDataType = 0;" & vbCrLf)		 'sqlNumeric
						'Response.Write("  }" & vbCrLf)
						'Response.Write("  else {" & vbCrLf)
						'Response.Write("    lngExcelDataType = 2;" & vbCrLf)		 'sqlUnknown
						'Response.Write("  }" & vbCrLf)

						Dim lngExcelDataType = 0 '?????

						ClientDLL.AddColumn(" ", 12, 0, False)
						For intCount = 0 To objCrossTab.ColumnHeadingUbound(0)
							ClientDLL.AddColumn(Left(objCrossTab.ColumnHeading(0, intCount), 255), lngExcelDataType, objCrossTab.IntersectionDecimals _
								, LCase(objCrossTab.Use1000Separator))
						Next

						strInterSectionType = "TODO Intersection type"
						ClientDLL.AddColumn(strInterSectionType, lngExcelDataType, objCrossTab.IntersectionDecimals, objCrossTab.Use1000Separator)


						If objCrossTab.PageBreakColumn = True Then
							lngLoopMin = 0
							lngLoopMax = objCrossTab.ColumnHeadingUbound(2)
						Else
							lngLoopMin = 0
							lngLoopMax = 0
						End If

						Dim sOutputGridCaption As String = objCrossTab.CrossTabName

						For lngCount = lngLoopMin To lngLoopMax
							If objCrossTab.PageBreakColumn = True Then
								ClientDLL.AddPage(sOutputGridCaption, Left(objCrossTab.ColumnHeading(2, lngCount), 255))
							Else
								If objCrossTab.CrossTabType = CrossTabType.cttAbsenceBreakdown Then
									ClientDLL.AddPage(sOutputGridCaption, "Absence Breakdown")
								Else
									ClientDLL.AddPage(sOutputGridCaption, objCrossTab.BaseTableName)
								End If
							End If

							objCrossTab.BuildOutputStrings(lngCount)
							ClientDLL.ArrayDim(objCrossTab.DataArrayCols, objCrossTab.DataArrayRows)
							For intCol = 0 To objCrossTab.DataArrayCols
								For intRow = 0 To objCrossTab.DataArrayRows
									ClientDLL.ArrayAddTo(intCol, intRow, HttpUtility.HtmlDecode(Left(objCrossTab.DataArray(intCol, intRow), 255)))
								Next
							Next

							ClientDLL.DataArray()
						Next

						ClientDLL.Complete()

					End If

				End If
			End If

			' Only send output if not email
			If Not (objCrossTab.OutputEmail And objCrossTab.OutputEmailID > 0) Then
				If IO.File.Exists(ClientDLL.GeneratedFile) Then
					Try
						Response.ClearContent()
						Response.AddHeader("Content-Disposition", "attachment; filename=" + strDownloadFileName)
						Response.TransmitFile(ClientDLL.GeneratedFile)
						Response.Flush()
					Catch ex As Exception
					Finally
						IO.File.Delete(ClientDLL.GeneratedFile)
					End Try
				End If
			End If

		End Function


		Public Function util_run_customreport_downloadoutput() As FilePathResult

			'Session("CT_Mode") = Request("txtMode")
			Session("OutputOptions_Format") = Request("txtFormat")
			Session("OutputOptions_Screen") = False	' Request("txtScreen")
			Session("OutputOptions_Printer") = Request("txtPrinter")
			Session("OutputOptions_PrinterName") = Request("txtPrinterName")
			Session("OutputOptions_Save") = Request("txtSave")
			Session("OutputOptions_SaveExisting") = Request("txtSaveExisting")
			Session("OutputOptions_Email") = Request("txtEmail")
			Session("OutputOptions_EmailGroupID") = Request("txtEmailGroupID")
			Session("OutputOptions_EmailGroup") = Request("txtEmailGroup")
			Session("OutputOptions_EmailSubject") = Request("txtEmailSubject")
			Session("OutputOptions_EmailAttachAs") = Request("txtEmailAttachAs")
			Session("OutputOptions_Filename") = Request("txtFilename")
			Session("utiltype") = Request.Form("txtUtilType")

			Dim objReport As HR.Intranet.Server.Report = Session("CustomReport")
			Dim ClientDLL As New HR.Intranet.Server.clsOutputRun
			ClientDLL.SessionInfo = CType(Session("SessionContext"), SessionInfo)
			Dim objUser As New HR.Intranet.Server.clsSettings
			objUser.SessionInfo = CType(Session("SessionContext"), SessionInfo)

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			Dim fOK As Boolean
			Dim bBradfordFactor As Boolean

			ClientDLL.ResetColumns()
			ClientDLL.ResetStyles()
			ClientDLL.SaveAsValues = Session("OfficeSaveAsValues").ToString()

			ClientDLL.SettingLocations(CInt(objUser.GetUserSetting("Output", "TitleCol", 3)) _
				, CInt(objUser.GetUserSetting("Output", "TitleRow", 2)) _
				, CInt(objUser.GetUserSetting("Output", "DataCol", 2)) _
				, CInt(objUser.GetUserSetting("Output", "DataRow", 4)))

			ClientDLL.SettingTitle(CBool(objUser.GetUserSetting("Output", "TitleGridLines", False)) _
				, CBool(objUser.GetUserSetting("Output", "TitleBold", True)) _
				, CBool(objUser.GetUserSetting("Output", "TitleUnderline", False)) _
				, CInt(objUser.GetUserSetting("Output", "TitleBackcolour", "16777215")) _
				, CInt(objUser.GetUserSetting("Output", "TitleForecolour", "6697779")) _
				, objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "TitleBackcolour", "16777215"))) _
				, objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "TitleForecolour", "6697779"))))

			ClientDLL.SettingHeading(CBool(objUser.GetUserSetting("Output", "HeadingGridLines", True)) _
				, CBool(objUser.GetUserSetting("Output", "HeadingBold", True)) _
				, CBool(objUser.GetUserSetting("Output", "HeadingUnderline", False)) _
				, CInt(objUser.GetUserSetting("Output", "HeadingBackcolour", 16248553)) _
				, CInt(objUser.GetUserSetting("Output", "HeadingForecolour", 6697779)) _
				, CInt(objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "HeadingBackcolour", 16248553)))) _
				, CInt(objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "HeadingForecolour", 6697779)))))

			ClientDLL.SettingData(CBool(objUser.GetUserSetting("Output", "DataGridLines", True)) _
				, CBool(objUser.GetUserSetting("Output", "DataBold", False)) _
				, CBool(objUser.GetUserSetting("Output", "DataUnderline", False)) _
				, CInt(objUser.GetUserSetting("Output", "DataBackcolour", 15988214)) _
				, CInt(objUser.GetUserSetting("Output", "DataForecolour", 6697779)) _
				, CInt(objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "DataBackcolour", 15988214)))) _
				, CInt(objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "DataForecolour", 6697779)))))

			ClientDLL.InitialiseStyles()

			ClientDLL.SettingOptions(objUser.GetUserSetting("Output", "WordTemplate", "").ToString() _
				, objUser.GetUserSetting("Output", "ExcelTemplate", "").ToString() _
				, CBool(objUser.GetUserSetting("Output", "ExcelGridlines", False)) _
				, CBool(objUser.GetUserSetting("Output", "ExcelHeaders", False)) _
				, CBool(objUser.GetUserSetting("Output", "ExcelOmitSpacerRow", False)) _
				, CBool(objUser.GetUserSetting("Output", "ExcelOmitSpacerCol", False)) _
				, CBool(objUser.GetUserSetting("Output", "AutoFitCols", True)) _
				, CBool(objUser.GetUserSetting("Output", "Landscape", True)) _
				, False) 'emailnotimplementedyet


			Dim lngFormat As OutputFormats
			Dim blnScreen As Boolean
			Dim blnPrinter As Boolean
			Dim strPrinterName As String
			Dim blnSave As Boolean
			Dim lngSaveExisting As Long
			Dim blnEmail As Boolean
			Dim lngEmailGroupID As Long
			Dim strEmailSubject As String
			Dim strEmailAttachAs As String
			'	Dim strFileName As String

			Dim arrayColumnsDefinition() As String
			Dim arrayPageBreakValues
			Dim arrayVisibleColumns
			Dim sEmailAddresses As String = ""
			Dim sErrorDescription As String = ""

			Dim strDownloadFileName As String
			Dim strDownloadExtension As String

			'Set Options
			If objReport.OutputPreview Then
				lngFormat = NullSafeInteger(Session("OutputOptions_Format"))
				blnScreen = False
				blnPrinter = False
				strPrinterName = ""
				blnSave = False
				lngSaveExisting = False
				blnEmail = Session("OutputOptions_Email")
				lngEmailGroupID = CLng(Session("OutputOptions_EmailGroupID"))
				strEmailSubject = Session("OutputOptions_EmailSubject")
				strEmailAttachAs = Session("OutputOptions_EmailAttachAs")
				strDownloadFileName = Request("txtFilename")

			Else
				lngFormat = objReport.OutputFormat
				blnScreen = objReport.OutputScreen
				blnPrinter = objReport.OutputPrinter
				strPrinterName = objReport.OutputPrinterName
				blnSave = objReport.OutputSave
				lngSaveExisting = objReport.OutputSaveExisting
				blnEmail = objReport.OutputEmail
				lngEmailGroupID = CLng(objReport.OutputEmailID)
				strEmailSubject = objReport.OutputEmailSubject
				strEmailAttachAs = objReport.OutputEmailAttachAs
				strDownloadFileName = objReport.DownloadFileName
			End If

			If strDownloadFileName.Length = 0 Then
				objReport.OutputFormat = lngFormat
				objReport.OutputFilename = ""
				strDownloadFileName = objReport.DownloadFileName
			End If

			strDownloadExtension = Path.GetExtension(strDownloadFileName)

			If (objReport.OutputEmail) And (objReport.OutputEmailID > 0) Then

				Try
					Dim rstEmailAddr = objDataAccess.GetDataTable("spASRIntGetEmailGroupAddresses", CommandType.StoredProcedure _
								, New SqlParameter("EmailGroupID", SqlDbType.Int) With {.Value = CleanNumeric(lngEmailGroupID)})

					If Not rstEmailAddr Is Nothing Then
						For Each objRow In rstEmailAddr.Rows
							sEmailAddresses = sEmailAddresses & objRow(0) & ";"
						Next
					End If

				Catch ex As Exception
					sErrorDescription = "Error getting the email addresses for group." & vbCrLf & FormatError(ex.Message)
				End Try

				fOK = ClientDLL.SetOptions(False, lngFormat, blnScreen, blnPrinter, strPrinterName, blnSave, lngSaveExisting, blnEmail, sEmailAddresses _
					, strEmailSubject, strEmailAttachAs, strDownloadExtension)

			Else

				fOK = ClientDLL.SetOptions(False, lngFormat, Session("OutputOptions_Screen"), Session("OutputOptions_Printer") _
				, Session("OutputOptions_PrinterName"), True, Session("OutputOptions_SaveExisting") _
				, Session("OutputOptions_Email"), Session("OutputOptions_EmailGroupID"), Session("OutputOptions_EmailSubject") _
				, Session("OutputOptions_EmailAttachAs"), strDownloadExtension)

			End If

			arrayColumnsDefinition = objReport.OutputArray_Columns
			arrayPageBreakValues = objReport.OutputArray_PageBreakValues
			arrayVisibleColumns = objReport.OutputArray_VisibleColumns


			ClientDLL.SizeColumnsIndependently = True

			Dim sColHeading As String
			Dim iColDataType As Integer
			Dim iColDecimals As Integer
			Dim sBreakValue As String
			Dim blnBreakCheck As Boolean
			Dim bIsCol1000 As Boolean
			Dim lngActualRow As Integer
			Dim lngCol As Integer
			Dim lngRow As Integer

			ClientDLL.ArrayDim(UBound(arrayVisibleColumns, 2), 0)

			If lngFormat = 0 Then	'Session("OutputOptions_Format") = 0 Then
				If Session("OutputOptions_Printer") = True Then
					ClientDLL.SetPrinter()
				End If
			Else
				ClientDLL.HeaderRows = 1
				If ClientDLL.GetFile() = True Then

					If objReport.ReportHasPageBreak Then

						ClientDLL.ArrayDim(UBound(arrayVisibleColumns, 2), 0)
						lngActualRow = 0
						lngRow = 1

						For Each objRow As DataRow In objReport.datCustomReportOutput.Rows

							lngRow += 1
							lngActualRow += 1
							If lngActualRow = objReport.datCustomReportOutput.Rows.Count Then

								If objReport.ReportHasSummaryInfo Then
									sBreakValue = "Grand Totals"
								Else
									sBreakValue = arrayPageBreakValues(lngActualRow)
								End If

								If (lngActualRow > 0) Then
									If bBradfordFactor = True Then
										ClientDLL.AddPage(objReport.ReportCaption, "Bradford Factor")
									Else
										ClientDLL.AddPage(objReport.ReportCaption, Replace(sBreakValue, "&&", "&"))
									End If

									For lngCol = 0 To UBound(arrayVisibleColumns, 2)
										sColHeading = arrayVisibleColumns(0, lngCol)
										iColDataType = arrayVisibleColumns(1, lngCol)
										iColDecimals = arrayVisibleColumns(2, lngCol)
										ClientDLL.AddColumn(sColHeading, iColDataType, iColDecimals, False)
										ClientDLL.ArrayAddTo(lngCol, 0, sColHeading)
									Next

									ClientDLL.DataArray()
									lngActualRow = 0
									blnBreakCheck = True
									sBreakValue = ""

								End If

							ElseIf objRow(1).ToString() = "*" And Not blnBreakCheck Then
								sBreakValue = arrayPageBreakValues(lngRow)

								If bBradfordFactor = True Then
									ClientDLL.AddPage(objReport.ReportCaption, "Bradford Factor")
								Else
									ClientDLL.AddPage(objReport.ReportCaption, sBreakValue)
								End If

								For lngCol = 0 To UBound(arrayVisibleColumns, 2)
									sColHeading = arrayVisibleColumns(0, lngCol)
									iColDataType = arrayVisibleColumns(1, lngCol)
									iColDecimals = arrayVisibleColumns(2, lngCol)
									bIsCol1000 = arrayVisibleColumns(3, lngCol)
									ClientDLL.AddColumn(sColHeading, iColDataType, iColDecimals, bIsCol1000)
									ClientDLL.ArrayAddTo(lngCol, 0, sColHeading)
								Next

								ClientDLL.DataArray()
								ClientDLL.ArrayDim(UBound(arrayVisibleColumns, 2), 0)
								lngActualRow = 0
								blnBreakCheck = True
								ClientDLL.ResetColumns()
								ClientDLL.ResetStyles()

							ElseIf Not objRow(0).ToString() = "*" Then
								blnBreakCheck = False
								lngCol = 0

								ClientDLL.ArrayReDim()

								For lngCount = 0 To UBound(arrayVisibleColumns, 2)
									ClientDLL.ArrayAddTo(lngCol, lngActualRow, objRow.Item(lngCount + 1).ToString())
									lngCol += 1
								Next

							End If

						Next

					Else ' no page break

						ClientDLL.ArrayDim(UBound(arrayVisibleColumns, 2), objReport.datCustomReportOutput.Rows.Count + 1)

						If bBradfordFactor = True Then
							ClientDLL.PageTitles = False
							ClientDLL.AddPage("Bradford Factor", "Bradford Factor")
						Else
							ClientDLL.AddPage(objReport.ReportCaption, Replace(objReport.BaseTableName, "&&", "&"))
						End If

						For lngCol = 0 To UBound(arrayVisibleColumns, 2)
							sColHeading = arrayVisibleColumns(0, lngCol)
							iColDataType = arrayVisibleColumns(1, lngCol)
							iColDecimals = arrayVisibleColumns(2, lngCol)
							bIsCol1000 = arrayVisibleColumns(3, lngCol)
							ClientDLL.AddColumn(sColHeading, iColDataType, iColDecimals, bIsCol1000)
							ClientDLL.ArrayAddTo(lngCol, 0, sColHeading)
						Next


						lngRow = 1
						For Each objRow As DataRow In objReport.datCustomReportOutput.Rows

							For iCountColumns = 1 To UBound(arrayVisibleColumns, 2) + 1
								If objReport.ReportHasSummaryInfo Then
									ClientDLL.ArrayAddTo(iCountColumns - 1, lngRow, objRow(iCountColumns).ToString())
								Else
									ClientDLL.ArrayAddTo(iCountColumns - 1, lngRow, objRow(iCountColumns + 1).ToString())
								End If
							Next

							lngRow += 1

						Next

					End If

					ClientDLL.DataArray()
				End If

			End If

			ClientDLL.Complete()

			' Only send output if not email
			If Not (objReport.OutputEmail And objReport.OutputEmailID > 0) Then

				If IO.File.Exists(ClientDLL.GeneratedFile) Then
					Try
						Response.ClearContent()
						Response.AddHeader("Content-Disposition", "attachment; filename=" + strDownloadFileName)
						Response.TransmitFile(ClientDLL.GeneratedFile)
						Response.Flush()
					Catch ex As Exception
					Finally
						IO.File.Delete(ClientDLL.GeneratedFile)
					End Try
				End If
			End If

		End Function

		Public Function util_run_calendarreport_data() As ActionResult
			Return View()
		End Function

		Public Function util_run_calendarreport_breakdown() As ActionResult

			Dim objCalendarEvent As New Models.CalendarEvent
			Dim sSQL As String

			Dim objCalendar As HR.Intranet.Server.CalendarReport = CType(Session("objCalendar" & Session("CalRepUtilID")), HR.Intranet.Server.CalendarReport)
			Dim intEventID As Int32 = Request.Form("txtBaseIndex").ToString()

			Dim datEvent As DataTable = objCalendar.EventsRecordset

			sSQL = "ID =" & intEventID
			Dim objRow As DataRow = datEvent.Select(sSQL).FirstOrDefault()

			objCalendarEvent.BaseID = objRow.Item("BaseID").ToString()
			objCalendarEvent.Description = objCalendar.ConvertDescription(objRow("description1").ToString(), objRow("description2").ToString(), objRow("descriptionExpr").ToString())
			objCalendarEvent.EventName = objRow.Item("Name").ToString()
			objCalendarEvent.StartDate = objRow.Item("StartDate").ToString()
			objCalendarEvent.StartSession = objRow.Item("StartSession").ToString()
			objCalendarEvent.EndDate = objRow.Item("EndDate").ToString()
			objCalendarEvent.EndSession = objRow.Item("EndSession").ToString()
			objCalendarEvent.Duration = objRow.Item("Duration").ToString()
			objCalendarEvent.Reason = objRow.Item("EventDescription1").ToString()
			objCalendarEvent.Region = objRow.Item("Region").ToString()
			objCalendarEvent.CalendarCode = objRow.Item("Legend").ToString()

			Dim datWorkingPatterns As DataTable = objCalendar.rsCareerChange
			If Not datWorkingPatterns Is Nothing Then
				sSQL = String.Format("BaseID = {0} AND [WP_Date] <= '{1}'", objCalendarEvent.BaseID, objCalendarEvent.StartDate)
				objRow = datWorkingPatterns.Select(sSQL, "[WP_Date]").FirstOrDefault()

				If Not objRow Is Nothing Then
					objCalendarEvent.WorkingPattern = objRow.Item("WP_Pattern").ToString()
				End If
			End If

			Return View(objCalendarEvent)

		End Function

		Function util_run_calendarreport_download() As FileStreamResult

			Dim strDownloadFileName As String = Request("txtFilename")
			Dim objCalendar = CType(Session("objCalendar" & Session("UtilID")), CalendarReport)

			Dim objOutput As New CalendarOutput
			objOutput.ReportData = objCalendar.Events
			objOutput.Calendar = objCalendar

			If strDownloadFileName = "" Then
				strDownloadFileName = objCalendar.CalendarReportName + ".xlsx"
			End If


			objOutput.DownloadFileName = strDownloadFileName
			objOutput.Generate(objCalendar.OutputFormat)

			If IO.File.Exists(objOutput.GeneratedFile) Then
				Try
					Response.ClearContent()
					Response.AddHeader("Content-Disposition", "attachment; filename=" + strDownloadFileName)
					Response.TransmitFile(objOutput.GeneratedFile)
					Response.Flush()
				Catch ex As Exception
				Finally
					IO.File.Delete(objOutput.GeneratedFile)
				End Try
			End If


		End Function



		Function util_run_calendarreport_data_submit() As ActionResult

			On Error Resume Next

			Session("CALREP_Action") = Request.Form("txtAction")
			Session("CALREP_Month") = Request.Form("txtMonth")
			Session("CALREP_Year") = Request.Form("txtYear")
			Session("CALREP_VisibleStartDate") = Request.Form("txtVisibleStartDate")
			Session("CALREP_VisibleEndDate") = Request.Form("txtVisibleEndDate")
			Session("CalRep_Mode") = Request.Form("txtMode")
			Session("EmailGroupID") = Request.Form("txtEmailGroupID")
			Session("CALREP_firstLoad") = Request.Form("txtLoadCount")

			Session("CALREP_IncludeBankHolidays") = Request.Form("txtIncludeBankHolidays")
			Session("CALREP_IncludeWorkingDaysOnly") = Request.Form("txtIncludeWorkingDaysOnly")
			Session("CALREP_ShowBankHolidays") = Request.Form("txtShowBankHolidays")
			Session("CALREP_ShowCaptions") = Request.Form("txtShowCaptions")
			Session("CALREP_ShowWeekends") = Request.Form("txtShowWeekends")
			Session("CALREP_ChangeOptions") = Request.Form("txtChangeOptions")

			' Go to the requested page.
			Return RedirectToAction("util_run_calendarreport_data")

		End Function

		<ValidateInput(False)>
		Function util_run_workflow() As ActionResult
			Return PartialView()
		End Function

		<ValidateInput(False)>
		Function WorkflowPendingSteps() As ActionResult
			Return PartialView()
		End Function

		<ValidateInput(False)>
		Function util_run_customreportsMain() As ActionResult
			Return PartialView()
		End Function

		Function Progress() As ActionResult
			Return PartialView()
		End Function

		Function Refresh() As ActionResult
			Return View()
		End Function

		'  Function util_run_promptedvaluessubmit() As ActionResult
		'     Return RedirectToAction("util_run")
		'    End Function

#End Region

#Region "Defining Reports"

		Function util_def_calendarreportdates_data_submit()
			Session("CalendarAction") = Request.Form("txtCalendarAction")
			Session("CalendarBaseTableID") = Request.Form("txtCalendarBaseTableID")
			Session("CalendarEventTableID") = Request.Form("txtCalendarEventTableID")
			Session("CalendarLookupTableID") = Request.Form("txtCalendarLookupTableID")

			'Response.Redirect("util_def_calendarreportdates_data")
			Return RedirectToAction("util_def_calendarreportdates_data")
		End Function

		Function util_def_calendarreport_submit()

			Try

				Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
				Dim prmID = New SqlParameter("piId", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = CleanNumeric(Request.Form("txtSend_ID"))}

				Dim prmFixedStart = New SqlParameter("psFixedStart", SqlDbType.VarChar, 100)
				If Len(Request.Form("txtSend_FixedStart")) > 0 Then
					prmFixedStart.Value = Request.Form("txtSend_FixedStart")
				Else
					prmFixedStart.Value = ""
				End If

				Dim prmFixedEnd = New SqlParameter("psFixedEnd", SqlDbType.VarChar, 100)
				If Len(Request.Form("txtSend_FixedEnd")) > 0 Then
					prmFixedEnd.Value = Request.Form("txtSend_FixedEnd")
				Else
					prmFixedEnd.Value = ""
				End If

				objDataAccess.ExecuteSP("spASRIntSaveCalendarReport", _
				New SqlParameter("psName", SqlDbType.VarChar, 255) With {.Value = Request.Form("txtSend_name")}, _
					New SqlParameter("psDescription", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_description")}, _
					New SqlParameter("piBaseTable", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_baseTable"))}, _
					New SqlParameter("pfAllRecords", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_allRecords"))}, _
					New SqlParameter("piPicklist", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_picklist"))}, _
					New SqlParameter("piFilter", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_filter"))}, _
					New SqlParameter("pfPrintFilterHeader", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_printFilterHeader"))}, _
					New SqlParameter("psUserName", SqlDbType.VarChar, 255) With {.Value = Request.Form("txtSend_userName")}, _
					New SqlParameter("piDescription1", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_desc1"))}, _
					New SqlParameter("piDescription2", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_desc2"))}, _
					New SqlParameter("piDescriptionExpr", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_descExpr"))}, _
					New SqlParameter("piRegion", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_region"))}, _
					New SqlParameter("pfGroupByDesc", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_groupbydesc"))}, _
					New SqlParameter("psDescSeparator", SqlDbType.VarChar, 100) With {.Value = Request.Form("txtSend_descseparator")}, _
					New SqlParameter("piStartType", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_StartType"))}, _
					prmFixedStart, _
					New SqlParameter("piStartFrequency", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_StartFrequency"))}, _
					New SqlParameter("piStartPeriod", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_StartPeriod"))}, _
					New SqlParameter("piStartDateExpr", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_CustomStart"))}, _
					New SqlParameter("piEndType", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_EndType"))}, _
					prmFixedEnd, _
					New SqlParameter("piEndFrequency", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_EndFrequency"))}, _
					New SqlParameter("piEndPeriod", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_EndPeriod"))}, _
					New SqlParameter("piEndDateExpr", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_CustomEnd"))}, _
					New SqlParameter("pfShowBankHols", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_ShadeBHols"))}, _
					New SqlParameter("pfShowCaptions", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_Captions"))}, _
					New SqlParameter("pfShowWeekends", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_ShadeWeekends"))}, _
					New SqlParameter("pfStartOnCurrentMonth", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_StartOnCurrentMonth"))}, _
					New SqlParameter("pfIncludeWorkdays", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_IncludeWorkingDaysOnly"))}, _
					New SqlParameter("pfIncludeBankHols", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_IncludeBHols"))}, _
					New SqlParameter("pfOutputPreview", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_OutputPreview"))}, _
					New SqlParameter("piOutputFormat", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_OutputFormat"))}, _
					New SqlParameter("pfOutputScreen", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_OutputScreen"))}, _
					New SqlParameter("pfOutputPrinter", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_OutputPrinter"))}, _
					New SqlParameter("psOutputPrinterName", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_OutputPrinterName")}, _
					New SqlParameter("pfOutputSave", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_OutputSave"))}, _
					New SqlParameter("piOutputSaveExisting", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_OutputSaveExisting"))}, _
					New SqlParameter("pfOutputEmail", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_OutputEmail"))}, _
					New SqlParameter("pfOutputEmailAddr", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_OutputEmailAddr"))}, _
					New SqlParameter("psOutputEmailSubject", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_OutputEmailSubject")}, _
					New SqlParameter("psOutputEmailAttachAs", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_OutputEmailAttachAs")}, _
					New SqlParameter("psOutputFilename", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_OutputFilename")}, _
					New SqlParameter("psAccess", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_access")}, _
					New SqlParameter("psJobsToHide", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_jobsToHide")}, _
					New SqlParameter("psJobsToHideGroups", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_jobsToHideGroups")}, _
					New SqlParameter("psEvents", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_Events")}, _
					New SqlParameter("psEvents2", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_Events2")}, _
					New SqlParameter("psOrderString", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_OrderString")}, _
					prmID)

				Session("confirmtext") = "Report has been saved successfully"
				Session("confirmtitle") = "Calendar Reports"
				Session("followpage") = "defsel"
				Session("reaction") = Request.Form("txtSend_reaction")
				Session("utilid") = prmID.Value

			Catch ex As Exception
				Response.Write("<HTML>" & vbCrLf)
				Response.Write("	<HEAD>" & vbCrLf)
				Response.Write("		<META NAME=""GENERATOR"" Content=""Microsoft Visual Studio 6.0"">" & vbCrLf)
				Response.Write("		<LINK href=""AutoBG.css"" rel=stylesheet type=text/css >" & vbCrLf)
				Response.Write("		<TITLE>" & vbCrLf)
				Response.Write("			OpenHR Intranet" & vbCrLf)
				Response.Write("		</TITLE>" & vbCrLf)
				Response.Write("		<meta http-equiv=""X-UA-Compatible"" content=""IE=5"">" & vbCrLf)
				Response.Write("	</HEAD>" & vbCrLf)
				Response.Write("	<BODY id=bdyMainBody name=""bdyMainBody"" " & Session("BodyTag") & ">" & vbCrLf)

				Response.Write("	<table align=center border=1 cellPadding=5 cellSpacing=0>" & vbCrLf)
				Response.Write("		<TR>" & vbCrLf)
				Response.Write("			<TD bgcolor=threedface>" & vbCrLf)
				Response.Write("				<table border=0 cellspacing=0 cellpadding=0>" & vbCrLf)
				Response.Write("				  <tr> " & vbCrLf)
				Response.Write("				    <td colspan=3 height=10></td>" & vbCrLf)
				Response.Write("				  </tr>" & vbCrLf)
				Response.Write("				  <tr> " & vbCrLf)
				Response.Write("				    <td colspan=3 align=center> " & vbCrLf)
				Response.Write("							<H3>Error</H3>" & vbCrLf)
				Response.Write("				    </td>" & vbCrLf)
				Response.Write("				  </tr>" & vbCrLf)
				Response.Write("				  <tr> " & vbCrLf)
				Response.Write("				    <td width=20 height=10></td> " & vbCrLf)
				Response.Write("				    <td> " & vbCrLf)
				Response.Write("							<H4>Error saving report</H4>" & vbCrLf)
				Response.Write("				    </td>" & vbCrLf)
				Response.Write("				    <td width=20></td> " & vbCrLf)
				Response.Write("				  </tr>" & vbCrLf)
				Response.Write("				  <tr> " & vbCrLf)
				Response.Write("				    <td width=20 height=10></td> " & vbCrLf)
				Response.Write("				    <td> " & vbCrLf)
				Response.Write(ex.Message & vbCrLf)
				Response.Write("			    </td>" & vbCrLf)
				Response.Write("			    <td width=20></td> " & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr> " & vbCrLf)
				Response.Write("			    <td colspan=3 height=20></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr> " & vbCrLf)
				Response.Write("			    <td colspan=3 height=10 align=center>" & vbCrLf)
				Response.Write("						<INPUT TYPE=button VALUE=""Retry"" NAME=""GoBack"" OnClick=""window.history.back(1)"" style=""WIDTH: 80px"" width=80 id=cmdGoBack>" & vbCrLf)
				Response.Write("			    </td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr>" & vbCrLf)
				Response.Write("			    <td colspan=3 height=10></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			</table>" & vbCrLf)
				Response.Write("    </td>" & vbCrLf)
				Response.Write("  </tr>" & vbCrLf)
				Response.Write("</table>" & vbCrLf)
				Response.Write("	</BODY>" & vbCrLf)
				Response.Write("<HTML>" & vbCrLf)

			End Try

			Return RedirectToAction("ConfirmOK")

		End Function

		Public Function util_def_calendarreportdates() As ActionResult
			Return View()
		End Function

		Public Function util_def_calendarreportdates_main() As ActionResult
			Return View()
		End Function

		Public Function util_def_calendarreport() As ActionResult
			Return View()
		End Function

		Function util_def_customreports_submit()

			Try

				Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
				Dim prmID = New SqlParameter("piId", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = CleanNumeric(Request.Form("txtSend_ID"))}

				objDataAccess.ExecuteSP("sp_ASRIntSaveCustomReport", _
						New SqlParameter("psName", SqlDbType.VarChar, 255) With {.Value = Request.Form("txtSend_name")}, _
						New SqlParameter("psDescription", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_description")}, _
						New SqlParameter("piBaseTableID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_baseTable"))}, _
						New SqlParameter("pfAllRecords", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_allRecords"))}, _
						New SqlParameter("piPicklistID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_picklist"))}, _
						New SqlParameter("piFilterID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_filter"))}, _
						New SqlParameter("piParent1TableID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_parent1Table"))}, _
						New SqlParameter("piParent1FilterID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_parent1Filter"))}, _
						New SqlParameter("piParent2TableID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_parent2Table"))}, _
						New SqlParameter("piParent2FilterID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_parent2Filter"))}, _
						New SqlParameter("pfSummary", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_summary"))}, _
						New SqlParameter("pfPrintFilterHeader", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_printFilterHeader"))}, _
						New SqlParameter("psUserName", SqlDbType.VarChar, 255) With {.Value = Request.Form("txtSend_userName")}, _
						New SqlParameter("pfOutputPreview", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_OutputPreview"))}, _
						New SqlParameter("piOutputFormat", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_OutputFormat"))}, _
						New SqlParameter("pfOutputScreen", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_OutputScreen"))}, _
						New SqlParameter("pfOutputPrinter", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_OutputPrinter"))}, _
						New SqlParameter("psOutputPrinterName", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_OutputPrinterName")}, _
						New SqlParameter("pfOutputSave", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_OutputSave"))}, _
						New SqlParameter("piOutputSaveExisting", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_OutputSaveExisting"))}, _
						New SqlParameter("pfOutputEmail", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_OutputEmail"))}, _
						New SqlParameter("piOutputEmailAddr", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_OutputEmailAddr"))}, _
						New SqlParameter("psOutputEmailSubject", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_OutputEmailSubject")}, _
						New SqlParameter("psOutputEmailAttachAs", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_OutputEmailAttachAs")}, _
						New SqlParameter("psOutputFilename", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_OutputFilename")}, _
						New SqlParameter("pfParent1AllRecords", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_parent1AllRecords"))}, _
						New SqlParameter("piParent1Picklist", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_parent1Picklist"))}, _
						New SqlParameter("pfParent2AllRecords", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_parent2AllRecords"))}, _
						New SqlParameter("piParent2Picklist", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_parent2Picklist"))}, _
						New SqlParameter("psAccess", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_access")}, _
						New SqlParameter("psJobsToHide", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_jobsToHide")}, _
						New SqlParameter("psJobsToHideGroups", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_jobsToHideGroups")}, _
						New SqlParameter("psColumns", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_columns")}, _
						New SqlParameter("psColumns2", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_columns2")}, _
						New SqlParameter("psChildString", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_childTable")}, _
						prmID,
						New SqlParameter("pfIgnoreZeros", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_IgnoreZeros"))})

				Session("confirmtext") = "Report has been saved successfully"
				Session("confirmtitle") = "Custom Reports"
				Session("followpage") = "defsel"
				Session("reaction") = Request.Form("txtSend_reaction")
				Session("utilid") = prmID.Value

				Return RedirectToAction("confirmok")

			Catch ex As Exception

				' TO DO error reprting
				Response.Write("<HTML>" & vbCrLf)
				Response.Write("	<HEAD>" & vbCrLf)
				Response.Write("		<META NAME=""GENERATOR"" Content=""Microsoft Visual Studio 6.0"">" & vbCrLf)
				Response.Write("		<LINK href=""OpenHR.css"" rel=stylesheet type=text/css >" & vbCrLf)
				Response.Write("		<TITLE>" & vbCrLf)
				Response.Write("			OpenHR Intranet" & vbCrLf)
				Response.Write("		</TITLE>" & vbCrLf)
				Response.Write("  <!--#INCLUDE FILE=""include/ctl_SetStyles.txt"" -->")
				Response.Write("	</HEAD>" & vbCrLf)
				Response.Write("	<BODY id=bdyMainBody name=""bdyMainBody"" " & Session("BodyTag") & ">" & vbCrLf)

				Response.Write("	<table align=center class=""outline"" cellPadding=5 cellSpacing=0>" & vbCrLf)
				Response.Write("		<TR>" & vbCrLf)
				Response.Write("			<TD>" & vbCrLf)
				Response.Write("				<table class=""invisible"" cellspacing=0 cellpadding=0>" & vbCrLf)
				Response.Write("				  <tr> " & vbCrLf)
				Response.Write("				    <td colspan=3 height=10></td>" & vbCrLf)
				Response.Write("				  </tr>" & vbCrLf)
				Response.Write("				  <tr> " & vbCrLf)
				Response.Write("				    <td colspan=3 align=center> " & vbCrLf)
				Response.Write("							<H3>Error</H3>" & vbCrLf)
				Response.Write("				    </td>" & vbCrLf)
				Response.Write("				  </tr>" & vbCrLf)
				Response.Write("				  <tr> " & vbCrLf)
				Response.Write("				    <td width=20 height=10></td> " & vbCrLf)
				Response.Write("				    <td> " & vbCrLf)
				Response.Write("							<H4>Error saving report</H4>" & vbCrLf)
				Response.Write("				    </td>" & vbCrLf)
				Response.Write("				    <td width=20></td> " & vbCrLf)
				Response.Write("				  </tr>" & vbCrLf)
				Response.Write("				  <tr> " & vbCrLf)
				Response.Write("				    <td width=20 height=10></td> " & vbCrLf)
				Response.Write("				    <td> " & vbCrLf)
				Response.Write(ex.Message & vbCrLf)
				Response.Write("			    </td>" & vbCrLf)
				Response.Write("			    <td width=20></td> " & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr> " & vbCrLf)
				Response.Write("			    <td colspan=3 height=20></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr> " & vbCrLf)
				Response.Write("			    <td colspan=3 height=10 align=center>" & vbCrLf)
				Response.Write("						<INPUT TYPE=button VALUE=""Retry"" NAME=""GoBack"" class=""btn"" OnClick=""window.history.back(1)"" style=""WIDTH: 80px"" width=80 id=cmdGoBack>" & vbCrLf)
				Response.Write("                      onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
				Response.Write("                      onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
				Response.Write("		                  onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
				Response.Write("                      onblur=""try{button_onBlur(this);}catch(e){}"" />" & vbCrLf)
				Response.Write("			    </td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr>" & vbCrLf)
				Response.Write("			    <td colspan=3 height=10></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			</table>" & vbCrLf)
				Response.Write("    </td>" & vbCrLf)
				Response.Write("  </tr>" & vbCrLf)
				Response.Write("</table>" & vbCrLf)
				Response.Write("	</BODY>" & vbCrLf)
				Response.Write("<HTML>" & vbCrLf)
			End Try

			Return RedirectToAction("confirmok")

		End Function

		Public Function util_def_calendarreportdates_data() As ActionResult
			Return View()
		End Function

		Function util_validate_customreports() As ActionResult
			Return View()
		End Function

		Function util_validate_calendarreport() As ActionResult
			Return View()
		End Function

		Function util_validate_crosstabs() As ActionResult
			Return View()
		End Function

#End Region

#Region "Expression Builder"

		Function util_def_expression() As ActionResult
			Return PartialView()
		End Function

		<HttpPost(), ValidateInput(False)>
		Function util_def_expression_Submit()


			Dim objExpression As HR.Intranet.Server.Expression
			Dim iExprType As Integer
			Dim iReturnType As Integer
			Dim sUtilType As String
			Dim sUtilType2 As String
			Dim fok As Boolean
			Dim cmdMakeHidden
			Dim prmUtilType
			Dim prmUtilID

			On Error Resume Next

			' Get the server DLL to save the expression definition

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			Dim objContext = CType(Session("SessionContext"), SessionInfo)
			objExpression = New Expression(objContext.LoginInfo)

			If Request.Form("txtSend_type") = 11 Then
				iExprType = 11
				iReturnType = 3
				sUtilType = "Filter"
				sUtilType2 = "filter"
			Else
				iExprType = 10
				iReturnType = 0
				sUtilType = "Calculation"
				sUtilType2 = "calculation"
			End If

			fok = objExpression.Initialise(NullSafeInteger(Request.Form("txtSend_tableID")), _
				NullSafeInteger(Request.Form("txtSend_ID")), CInt(iExprType), CInt(iReturnType))

			If fok Then
				fok = objExpression.SetExpressionDefinition(CStr(Request.Form("txtSend_components1")), _
					"", "", "", "", CStr(Request.Form("txtSend_names")))
			End If

			If fok Then
				fok = objExpression.SaveExpression(CStr(Request.Form("txtSend_name")), _
					CStr(Request.Form("txtSend_userName")), _
					CStr(Request.Form("txtSend_access")), _
					CStr(Request.Form("txtSend_description")))

				If fok Then
					If (Request.Form("txtSend_access") = "HD") And _
						(Request.Form("txtSend_ID") > 0) Then
						' Hide any utilities that use this filter/calc.
						' NB. The check to see if we can do this has already been done as part of the filter/calc validation. */

						objDataAccess.ExecuteSP("sp_ASRIntMakeUtilitiesHidden" _
							, New SqlParameter("piUtilityType", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_type"))} _
							, New SqlParameter("piUtilityID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_ID"))})

					End If

					Session("confirmtext") = sUtilType & " has been saved successfully"
					Session("confirmtitle") = sUtilType & "s"
					Session("followpage") = "defsel"
					Session("reaction") = Request.Form("txtSend_reaction")
					Session("utilid") = objExpression.ExpressionID

				Else

					' TODO ERROR REPORTING
					Response.Write("<HTML>" & vbCrLf)
					Response.Write("	<HEAD>" & vbCrLf)
					Response.Write("		<META NAME=""GENERATOR"" Content=""Microsoft Visual Studio 6.0"">" & vbCrLf)
					Response.Write("		<LINK href=""OpenHR.css"" rel=stylesheet type=text/css >" & vbCrLf)
					Response.Write("		<TITLE>" & vbCrLf)
					Response.Write("			OpenHR Intranet" & vbCrLf)
					Response.Write("		</TITLE>" & vbCrLf)
					Response.Write("  <!--#INCLUDE FILE=""include/ctl_SetStyles.txt"" -->")
					Response.Write("	</HEAD>" & vbCrLf)
					Response.Write("	<BODY id=bdyMainBody name=""bdyMainBody"" " & Session("BodyTag") & ">" & vbCrLf)

					Response.Write("	<table align=center class=""outline"" cellPadding=5 cellSpacing=0>" & vbCrLf)
					Response.Write("		<TR>" & vbCrLf)
					Response.Write("			<TD>" & vbCrLf)
					Response.Write("				<table class=""invisible"" cellspacing=0 cellpadding=0>" & vbCrLf)
					Response.Write("				  <tr> " & vbCrLf)
					Response.Write("				    <td colspan=3 height=10></td>" & vbCrLf)
					Response.Write("				  </tr>" & vbCrLf)
					Response.Write("				  <tr> " & vbCrLf)
					Response.Write("				    <td colspan=3 align=center> " & vbCrLf)
					Response.Write("							<H3>Error</H3>" & vbCrLf)
					Response.Write("				    </td>" & vbCrLf)
					Response.Write("				  </tr>" & vbCrLf)
					Response.Write("				  <tr> " & vbCrLf)
					Response.Write("				    <td width=20 height=10></td> " & vbCrLf)
					Response.Write("				    <td> " & vbCrLf)
					Response.Write("							<H4>Error saving " & sUtilType2 & "</H4>" & vbCrLf)
					Response.Write("				    </td>" & vbCrLf)
					Response.Write("				    <td width=20></td> " & vbCrLf)
					Response.Write("				  </tr>" & vbCrLf)
					Response.Write("				  <tr> " & vbCrLf)
					Response.Write("				    <td width=20 height=10></td> " & vbCrLf)
					Response.Write("				    <td> " & vbCrLf)
					Response.Write("							Unknown error" & vbCrLf)
					Response.Write("			    </td>" & vbCrLf)
					Response.Write("			    <td width=20></td> " & vbCrLf)
					Response.Write("			  </tr>" & vbCrLf)
					Response.Write("			  <tr> " & vbCrLf)
					Response.Write("			    <td colspan=3 height=20></td>" & vbCrLf)
					Response.Write("			  </tr>" & vbCrLf)
					Response.Write("			  <tr> " & vbCrLf)
					Response.Write("			    <td colspan=3 height=10 align=center>" & vbCrLf)
					Response.Write("						<INPUT TYPE=button VALUE=""Retry"" NAME=""GoBack"" class=""btn"" OnClick=""window.history.back(1)"" style=""WIDTH: 80px"" width=80 id=cmdGoBack>" & vbCrLf)
					Response.Write("                      onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
					Response.Write("                      onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
					Response.Write("		                  onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
					Response.Write("                      onblur=""try{button_onBlur(this);}catch(e){}"" />" & vbCrLf)
					Response.Write("			    </td>" & vbCrLf)
					Response.Write("			  </tr>" & vbCrLf)
					Response.Write("			  <tr>" & vbCrLf)
					Response.Write("			    <td colspan=3 height=10></td>" & vbCrLf)
					Response.Write("			  </tr>" & vbCrLf)
					Response.Write("			</table>" & vbCrLf)
					Response.Write("    </td>" & vbCrLf)
					Response.Write("  </tr>" & vbCrLf)
					Response.Write("</table>" & vbCrLf)
					Response.Write("	</BODY>" & vbCrLf)
					Response.Write("<HTML>" & vbCrLf)
				End If

			End If

			objExpression = Nothing

			'If fok Then
			'Return RedirectToAction("DefSel")
			' Else
			'TODO - error message
			Return RedirectToAction("confirmok")
			' End If

		End Function

		<HttpPost()>
		Function quickfind_Submit(value As FormCollection)
			Dim sErrorMsg = ""

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			' Only process the form submission if the referring page was the default page.
			' If it wasn't then redirect to the login page.

			Dim sFilterSQL = Request.Form("txtGotoOptionFilterSQL")
			Dim sFilterDef = Request.Form("txtGotoOptionFilterDef")
			Dim sValue = Request.Form("txtGotoOptionValue")
			Dim sNextPage = Request.Form("txtGotoOptionPage")
			Dim sAction = Request.Form("txtGotoOptionAction")

			Dim lngRecordID = 0

			Session("optionScreenID") = Request.Form("txtGotoOptionScreenID")
			Session("optionTableID") = Request.Form("txtGotoOptionTableID")
			Session("optionViewID") = Request.Form("txtGotoOptionViewID")
			Session("optionOrderID") = Request.Form("txtGotoOptionOrderID")
			Session("optionRecordID") = Request.Form("txtGotoOptionRecordID")
			Session("optionFilterDef") = Request.Form("txtGotoOptionFilterDef")
			Session("optionFilterSQL") = sFilterSQL
			Session("optionFilterDef") = sFilterDef
			Session("optionValue") = sValue
			Session("optionLinkTableID") = Request.Form("txtGotoOptionLinkTableID")
			Session("optionLinkOrderID") = Request.Form("txtGotoOptionLinkOrderID")
			Session("optionLinkViewID") = Request.Form("txtGotoOptionLinkViewID")
			Session("optionLinkRecordID") = Request.Form("txtGotoOptionLinkRecordID")
			Session("optionColumnID") = Request.Form("txtGotoOptionColumnID")
			Session("optionLookupColumnID") = Request.Form("txtGotoOptionLookupColumnID")
			Session("optionLookupMandatory") = Request.Form("txtGotoOptionLookupMandatory")
			Session("optionLookupValue") = Request.Form("txtGotoOptionLookupValue")
			Session("optionFile") = Request.Form("txtGotoOptionFile")
			Session("optionExtension") = Request.Form("txtGotoOptionExtension")
			'Session("optionOLEOnServer") = Request.Form("txtGotoOptionOLEOnServer")
			Session("optionOLEType") = Request.Form("txtGotoOptionOLEType")
			Session("optionAction") = sAction
			Session("optionFirstRecPos") = Request.Form("txtGotoOptionFirstRecPos")
			Session("optionCurrentRecCount") = Request.Form("txtGotoOptionCurrentRecCount")
			Session("optionPageAction") = Request.Form("txtGotoOptionPageAction")
			Session("optionCourseTitle") = Request.Form("txtGotoOptionCourseTitle")
			Session("optionExprType") = Request.Form("txtGotoOptionExprType")
			Session("optionExprID") = Request.Form("txtGotoOptionExprID")
			Session("optionFunctionID") = Request.Form("txtGotoOptionFunctionID")
			Session("optionParameterIndex") = Request.Form("txtGotoOptionParameterIndex")

			If sAction = "" Then
				' Go to the requested page.
				Return RedirectToAction(sNextPage)
			End If

			If sAction = "CANCEL" Then
				' Go to the requested page.
				Session("errorMessage") = sErrorMsg
				Return RedirectToAction(sNextPage)
			End If

			If sAction = "QUICKFIND" Then

				Dim prmResult = New SqlParameter("@plngRecordID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

				Try

					objDataAccess.ExecuteSP("spASRIntGetQuickFindRecord" _
						, New SqlParameter("@plngTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionTableID"))} _
						, New SqlParameter("@plngViewID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionViewID"))} _
						, New SqlParameter("@plngColumnID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionColumnID"))} _
						, New SqlParameter("@psValue", SqlDbType.VarChar, -1) With {.Value = sValue} _
						, New SqlParameter("@psFilterDef", SqlDbType.VarChar, -1) With {.Value = sFilterDef} _
						, prmResult _
						, New SqlParameter("@psDecimalSeparator", SqlDbType.VarChar, 100) With {.Value = Session("LocaleDecimalSeparator")} _
						, New SqlParameter("@psLocaleDateFormat", SqlDbType.VarChar, 100) With {.Value = Session("LocaleDateFormat")})

					If (CInt(prmResult.Value) = 0) Then
						sErrorMsg = "No records match the criteria."

						If Len(sFilterDef) > 0 Then
							sErrorMsg = sErrorMsg & vbCrLf & "Try removing the filter."
						End If
					Else
						' A record has been found !
						lngRecordID = CInt(prmResult.Value)
					End If


				Catch ex As Exception
					sErrorMsg = "Error trying to run 'quick find'." & vbCrLf & FormatError(Err.Description)

				End Try


				Session("errorMessage") = sErrorMsg

				If Len(sErrorMsg) > 0 Then
					' Go to the requested page.
					Return RedirectToAction("Quickfind")
				End If

			End If

			' Go to the requested page.
			Session("optionRecordID") = lngRecordID
			Return RedirectToAction(sNextPage)

		End Function


		Function emptyoption() As ActionResult
			Return View()
		End Function


		<HttpPost()>
		Function util_def_exprcomponent_submit(value As FormCollection)

			Dim sErrorMsg As String = ""
			Dim sNextPage As String
			Dim sAction As String

			On Error Resume Next

			' Read the information from the calling form.
			sNextPage = Request.Form("txtGotoOptionPage")
			sAction = Request.Form("txtGotoOptionAction")

			Session("optionScreenID") = Request.Form("txtGotoOptionScreenID")
			Session("optionTableID") = Request.Form("txtGotoOptionTableID")
			Session("optionViewID") = Request.Form("txtGotoOptionViewID")
			Session("optionOrderID") = Request.Form("txtGotoOptionOrderID")
			Session("optionRecordID") = Request.Form("txtGotoOptionRecordID")
			Session("optionFilterDef") = Request.Form("txtGotoOptionFilterDef")
			Session("optionFilterSQL") = Request.Form("txtGotoOptionFilterSQL")
			Session("optionValue") = Request.Form("txtGotoOptionValue")
			Session("optionLinkTableID") = Request.Form("txtGotoOptionLinkTableID")
			Session("optionLinkOrderID") = Request.Form("txtGotoOptionLinkOrderID")
			Session("optionLinkViewID") = Request.Form("txtGotoOptionLinkViewID")
			Session("optionLinkRecordID") = Request.Form("txtGotoOptionLinkRecordID")
			Session("optionColumnID") = Request.Form("txtGotoOptionColumnID")
			Session("optionLookupColumnID") = Request.Form("txtGotoOptionLookupColumnID")
			Session("optionLookupMandatory") = Request.Form("txtGotoOptionLookupMandatory")
			Session("optionLookupValue") = Request.Form("txtGotoOptionLookupValue")
			Session("optionFile") = Request.Form("txtGotoOptionFile")
			Session("optionExtension") = Request.Form("txtGotoOptionExtension")
			'Session("optionOLEOnServer") = Request.Form("txtGotoOptionOLEOnServer")
			Session("optionOLEType") = Request.Form("txtGotoOptionOLEType")
			Session("optionAction") = sAction
			Session("optionPageAction") = Request.Form("txtGotoOptionPageAction")
			Session("optionCourseTitle") = Request.Form("txtGotoOptionCourseTitle")
			Session("optionFirstRecPos") = Request.Form("txtGotoOptionFirstRecPos")
			Session("optionCurrentRecCount") = Request.Form("txtGotoOptionCurrentRecCount")
			Session("optionExprType") = Request.Form("txtGotoOptionExprType")
			Session("optionExprID") = Request.Form("txtGotoOptionExprID")
			Session("optionFunctionID") = Request.Form("txtGotoOptionFunctionID")
			Session("optionParameterIndex") = Request.Form("txtGotoOptionParameterIndex")
			Session("optionDefSelType") = Request.Form("txtGotoOptionDefSelType")

			If sAction = "CANCEL" Then
				' Go to the requested page.

				Session("errorMessage") = sErrorMsg
			End If

			If sAction = "SELECTCOMPONENT" Then
				Session("errorMessage") = sErrorMsg
			End If

			' Go to the requested page.
			Return RedirectToAction(sNextPage)


		End Function

		Function util_def_exprcomponent() As ActionResult
			Return PartialView()
		End Function

		<ValidateInput(False)>
		Function util_test_expression() As ActionResult
			Return View()
		End Function

		<ValidateInput(False)>
		Function util_test_expression_pval() As ActionResult
			Return View()
		End Function

		<ValidateInput(False)>
		Function util_test_expression_submit(value As FormCollection)
			Return RedirectToAction("util_def_expression")
		End Function

		<ValidateInput(False)>
		Function util_validate_expression() As ActionResult
			Return View()
		End Function

		Function util_dialog_expression() As ActionResult
			Return View()
		End Function




		Function FieldRec() As ActionResult
			Return View()
		End Function


#End Region

		Function recordEdit(Optional sParameters As String = "") As ActionResult

			If Len(sParameters) > 0 Then
				' SSI Mode

				Dim lngTopLevelRecordID As Int32
				Dim lngRecID As Int32

				' Response.Write "#<FONT COLOR='Red'><B>session(linkID) = " & session("linkID") & "</B></FONT>#<BR>"
				' Response.Write "#<FONT COLOR='Red'><B>sParameters = " & sParameters & "</B></FONT>#"

				lngTopLevelRecordID = Session("TopLevelRecID")

				If NullSafeInteger(Session("tableID")) = NullSafeInteger(Session("SSILinkTableID")) Then
					' Top Level table.
					Session("recordID") = lngTopLevelRecordID
					Session("parentTableID") = 0
					Session("parentRecordID") = 0
				Else
					' Child table.
					Session("viewID") = 0
					Session("recordID") = lngRecID
					Session("parentTableID") = Session("SSILinkTableID")
					Session("parentRecordID") = lngTopLevelRecordID
				End If

				' Order not important.
				Session("orderID") = 0




			End If



			Return PartialView()
		End Function

		<HttpPost()>
		Function recordEditMain(psScreenInfo As String) As ActionResult

			Dim sErrorDescription As String = ""

			Session("action") = ""
			Session("parentTableID") = 0
			Session("parentRecordID") = 0
			Session("selectSQL") = ""
			Session("errorMessage") = ""
			Session("warningFlag") = ""
			Session("previousAction") = ""
			Session("orderID") = 0

			Dim objDatabaseAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			Dim sParameters As String = psScreenInfo

			Session("linkType") = Left(sParameters, InStr(sParameters, "_") - 1)

			sParameters = Mid(sParameters, InStr(sParameters, "_") + 1)

			Session("TopLevelRecID") = Left(sParameters, InStr(sParameters, "_") - 1)

			If Session("linkType") = "multifind" Then
				'NHRD26112013 Jira-3484 default to personnel table instead of 0 as it only seems to be a problem in that first Next on fresh SSI
				Session("screenID") = 1
				Session("title") = ""
				Session("startMode") = 0
				Session("tableID") = Mid(sParameters, InStr(sParameters, "_") + 1, ((InStr(sParameters, "!") - 1) - InStr(sParameters, "_")))
				Session("viewID") = Mid(sParameters, InStr(sParameters, "!") + 1)
				Session("tableType") = 1
			Else
				Session("linkID") = Mid(sParameters, InStr(sParameters, "_") + 1)

				Dim prmLinkID = New SqlParameter("piLinkID", SqlDbType.Int)
				prmLinkID.Value = NullSafeInteger(CleanNumeric(Session("linkID")))

				Dim prmScreenID = New SqlParameter("piScreenID", SqlDbType.Int)
				prmScreenID.Direction = ParameterDirection.Output

				Dim prmTableID = New SqlParameter("piTableID", SqlDbType.Int)
				prmTableID.Direction = ParameterDirection.Output

				Dim prmTitle = New SqlParameter("psTitle", SqlDbType.VarChar, 8000)
				prmTitle.Direction = ParameterDirection.Output

				Dim prmStartMode = New SqlParameter("piStartMode", SqlDbType.Int)
				prmStartMode.Direction = ParameterDirection.Output

				Dim prmTableType = New SqlParameter("piTableType", SqlDbType.Int)
				prmTableType.Direction = ParameterDirection.Output

				Try
					objDatabaseAccess.ExecuteSP("spASRIntGetLinkInfo", prmLinkID, prmScreenID, prmTableID, prmTitle, prmStartMode, prmTableType)

					Session("screenID") = CInt(prmScreenID.Value)
					Session("tableID") = CInt(prmTableID.Value)
					Session("title") = prmTitle.Value.ToString()
					Session("startMode") = CInt(prmStartMode.Value)
					Session("tableType") = CInt(prmTableType.Value)
					Session("viewID") = Session("SSILinkViewID")

				Catch ex As Exception
					sErrorDescription = "Unable to get the link definition." & vbCrLf & FormatError(Err.Description)

				End Try

			End If

			If Session("linkType") = "multifind" Then
				Return RedirectToAction("Find", New With {.sParameters = "LOAD_0_0_"})
			Else
				If (NullSafeInteger(Session("SSILinkTableID")) = NullSafeInteger(Session("SingleRecordTableID"))) _
						And (NullSafeInteger(Session("SSILinkViewID")) = NullSafeInteger(Session("SingleRecordViewID"))) _
						And (NullSafeInteger(Session("TopLevelRecID")) = 0) _
						And (NullSafeInteger(Session("tableID")) <> NullSafeInteger(Session("SingleRecordTableID"))) Then
					'TODO: error - no parent record in the current view.          
				End If
				If CleanNumeric(Session("startMode")) <> 3 Then
					Return RedirectToAction("recordEdit", New With {.sParameters = sParameters})
				Else
					Return RedirectToAction("Find", New With {.sParameters = "LOAD_0_0_"})
				End If
			End If

		End Function


		Function FormError() As JsonResult
			' replaces response.redirect("error") 
			If NullSafeString(Session("ErrorTitle")).Length = 0 Then Session("ErrorTitle") = "Unspecified Form"
			If NullSafeString(Session("ErrorText")).Length = 0 Then Session("ErrorText") = "Unspecified Error (" & Session("ErrorTitle") & ")"

			Dim errorResponse = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = ""}
			Return Json(errorResponse, JsonRequestBehavior.AllowGet)

		End Function


#Region "Picklists"

		Function util_def_picklist() As ActionResult
			Return PartialView()
		End Function

		<HttpPost()>
		Function util_def_picklist_submit()

			Try

				Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
				Dim prmID = New SqlParameter("piId", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = CleanNumeric(Request.Form("txtSend_ID"))}

				objDataAccess.ExecuteSP("sp_ASRIntSavePicklist", _
					New SqlParameter("psName", SqlDbType.VarChar, 255) With {.Value = Request.Form("txtSend_name")}, _
					New SqlParameter("psDescription", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_description")}, _
					New SqlParameter("psAccess", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_access")}, _
					New SqlParameter("psUserName", SqlDbType.VarChar, 255) With {.Value = Request.Form("txtSend_userName")}, _
					New SqlParameter("psColumns", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_columns")}, _
					New SqlParameter("psColumns2", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_columns2")}, _
					prmID, _
					New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_tableID"))})

				Session("confirmtext") = "Picklist has been saved successfully"
				Session("confirmtitle") = "Picklists"
				Session("followpage") = "defsel"
				Session("reaction") = Request.Form("txtSend_reaction")
				Session("utilid") = prmID.Value

			Catch ex As Exception

				Response.Write("<HTML>" & vbCrLf)
				Response.Write("	<HEAD>" & vbCrLf)
				Response.Write("		<META NAME=""GENERATOR"" Content=""Microsoft Visual Studio 6.0"">" & vbCrLf)
				Response.Write("		<LINK href=""OpenHR.css"" rel=stylesheet type=text/css >" & vbCrLf)
				Response.Write("		<TITLE>" & vbCrLf)
				Response.Write("			OpenHR Intranet" & vbCrLf)
				Response.Write("		</TITLE>" & vbCrLf)
				Response.Write("	</HEAD>" & vbCrLf)
				Response.Write("	<BODY id=bdyMainBody name=""bdyMainBody"" " & Session("BodyTag") & ">" & vbCrLf)

				Response.Write("	<table align=center class=""outline"" cellPadding=5 cellSpacing=0>" & vbCrLf)
				Response.Write("		<TR>" & vbCrLf)
				Response.Write("			<TD>" & vbCrLf)
				Response.Write("				<table class=""invisible"" cellspacing=0 cellpadding=0>" & vbCrLf)
				Response.Write("				  <tr> " & vbCrLf)
				Response.Write("				    <td colspan=3 height=10></td>" & vbCrLf)
				Response.Write("				  </tr>" & vbCrLf)
				Response.Write("				  <tr> " & vbCrLf)
				Response.Write("				    <td colspan=3 align=center> " & vbCrLf)
				Response.Write("							<H3>Error</H3>" & vbCrLf)
				Response.Write("				    </td>" & vbCrLf)
				Response.Write("				  </tr>" & vbCrLf)
				Response.Write("				  <tr> " & vbCrLf)
				Response.Write("				    <td width=20 height=10></td> " & vbCrLf)
				Response.Write("				    <td> " & vbCrLf)
				Response.Write("							<H4>Error saving picklist</H4>" & vbCrLf)
				Response.Write("				    </td>" & vbCrLf)
				Response.Write("				    <td width=20></td> " & vbCrLf)
				Response.Write("				  </tr>" & vbCrLf)
				Response.Write("				  <tr> " & vbCrLf)
				Response.Write("				    <td width=20 height=10></td> " & vbCrLf)
				Response.Write("				    <td> " & vbCrLf)
				Response.Write(ex.Message & vbCrLf)
				Response.Write("			    </td>" & vbCrLf)
				Response.Write("			    <td width=20></td> " & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr> " & vbCrLf)
				Response.Write("			    <td colspan=3 height=20></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr> " & vbCrLf)
				Response.Write("			    <td colspan=3 height=10 align=center>" & vbCrLf)
				Response.Write("						<INPUT TYPE=button VALUE=""Retry"" NAME=""GoBack"" OnClick=""window.history.back(1)"" class=""btn"" style=""WIDTH: 80px"" width=80 id=cmdGoBack>" & vbCrLf)
				Response.Write("                      onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
				Response.Write("                      onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
				Response.Write("		                  onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
				Response.Write("                      onblur=""try{button_onBlur(this);}catch(e){}"" />" & vbCrLf)
				Response.Write("			    </td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			  <tr>" & vbCrLf)
				Response.Write("			    <td colspan=3 height=10></td>" & vbCrLf)
				Response.Write("			  </tr>" & vbCrLf)
				Response.Write("			</table>" & vbCrLf)
				Response.Write("    </td>" & vbCrLf)
				Response.Write("  </tr>" & vbCrLf)
				Response.Write("</table>" & vbCrLf)
				Response.Write("	</BODY>" & vbCrLf)
				Response.Write("<HTML>" & vbCrLf)

			End Try

			Return RedirectToAction("ConfirmOK")

		End Function

		Function picklistSelectionMain() As ActionResult
			Return View()
		End Function

		Function picklistSelection() As ActionResult
			Return View()
		End Function

		Function picklistSelectionData() As ActionResult
			Return View()
		End Function

		Function picklistSelectionData_Submit(value As FormCollection)

			' Read the information from the calling form.
			Session("tableID") = Request.Form("txtTableID")
			Session("viewID") = Request.Form("txtViewID")
			Session("orderID") = Request.Form("txtOrderID")
			Session("pageAction") = Request.Form("txtPageAction")
			Session("firstRecPos") = Request.Form("txtFirstRecPos")
			Session("currentRecCount") = Request.Form("txtCurrentRecCount")
			Session("locateValue") = Request.Form("txtGotoLocateValue")

			Session("picklistSelectionDataLoading") = False

			' Go to the requested page.
			Return RedirectToAction("picklistSelectionData")

		End Function

		Function util_validate_picklist() As ActionResult
			Return View()
		End Function

#End Region

#Region "Utilities"
		Function util_def_mailmerge() As ActionResult
			'Throw New NotImplementedException()
			Return View()
		End Function

		<HttpPost()>
		Function util_def_mailmerge_submit()

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			Try

				Dim prmID = New SqlParameter("piID", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = CleanNumeric(Request.Form("txtSend_ID"))}

				objDataAccess.ExecuteSP("sp_ASRIntSaveMailMerge" _
					, New SqlParameter("@psName", SqlDbType.VarChar, 255) With {.Value = Request.Form("txtSend_name")} _
					, New SqlParameter("@psDescription", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_description")} _
					, New SqlParameter("@piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_baseTable"))} _
					, New SqlParameter("@piSelection", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_selection"))} _
					, New SqlParameter("@piPicklistID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_picklist"))} _
					, New SqlParameter("@piFilterID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_filter"))} _
					, New SqlParameter("@piOutputFormat", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_outputformat"))} _
					, New SqlParameter("@pfOutputSave", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_outputsave"))} _
					, New SqlParameter("@psOutputFilename", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_outputfilename")} _
					, New SqlParameter("@piEmailAddrID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_emailaddrid"))} _
					, New SqlParameter("@psEmailSubject", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_emailsubject")} _
					, New SqlParameter("@psTemplateFileName", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_templatefilename")} _
					, New SqlParameter("@pfOutputScreen", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_outputscreen"))} _
					, New SqlParameter("@psUserName", SqlDbType.VarChar, 255) With {.Value = Request.Form("txtSend_userName")} _
					, New SqlParameter("@pfEmailAsAttachment", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_emailasattachment"))} _
					, New SqlParameter("@psEmailAttachmentName", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_emailattachmentname")} _
					, New SqlParameter("@pfSuppressBlanks", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_suppressblanks"))} _
					, New SqlParameter("@pfPauseBeforeMerge", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_pausebeforemerge"))} _
					, New SqlParameter("@pfOutputPrinter", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_outputprinter"))} _
					, New SqlParameter("@psOutputPrinterName", SqlDbType.VarChar, 255) With {.Value = Request.Form("txtSend_outputprintername")} _
					, New SqlParameter("@piDocumentMapID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_documentmapid"))} _
					, New SqlParameter("@pfManualDocManHeader", SqlDbType.Bit) With {.Value = CleanBoolean(Request.Form("txtSend_manualdocmanheader"))} _
					, New SqlParameter("@psAccess", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_access")} _
					, New SqlParameter("@psJobsToHide", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_jobsToHide")} _
					, New SqlParameter("@psJobsToHideGroups", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_jobsToHideGroups")} _
					, New SqlParameter("@psColumns", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_columns")} _
					, New SqlParameter("@psColumns2", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_columns2")} _
				, prmID)

				Session("confirmtext") = "Mail Merge has been saved successfully"
				Session("confirmtitle") = "Mail Merge"
				Session("followpage") = "defsel"
				Session("reaction") = Request.Form("txtSend_reaction")
				Session("utilid") = CInt(prmID.Value)

				Return RedirectToAction("confirmok")

			Catch ex As Exception

				Response.Write("<HTML>" & vbCrLf)
				Response.Write("	<HEAD>" & vbCrLf)
				Response.Write("		<META NAME=""GENERATOR"" Content=""Microsoft Visual Studio 6.0"">" & vbCrLf)
				Response.Write("		<LINK href=""OpenHR.css"" rel=stylesheet type=text/css >" & vbCrLf)
				Response.Write("		<TITLE>" & vbCrLf)
				Response.Write("			OpenHR Intranet" & vbCrLf)
				Response.Write("		</TITLE>" & vbCrLf)
				Response.Write("		<meta http-equiv=""X-UA-Compatible"" content=""IE=5"">" & vbCrLf)
				Response.Write("  <!--#INCLUDE FILE=""include/ctl_SetStyles.txt"" -->")
				Response.Write("	</HEAD>" & vbCrLf)
				Response.Write("	<BODY>" & vbCrLf)
				Response.Write("Error saving definition : <BR>" & Err.Description & "<BR>" & vbCrLf)
				Response.Write("<INPUT TYPE=button VALUE=Retry NAME=GoBack OnClick=" & Chr(34) & "window.history.back(1)" & Chr(34) & " class=""btn"" style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & " width=100 id=cmdGoBack>")
				Response.Write("                      onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
				Response.Write("                      onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
				Response.Write("		                  onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
				Response.Write("                      onblur=""try{button_onBlur(this);}catch(e){}"" />" & vbCrLf)
				'Response.Write(vbCrLf & vbCrLf & sSQLString)
				Response.Write("	</BODY>" & vbCrLf)
				Response.Write("<HTML>" & vbCrLf)

			End Try


		End Function

		'ND my original call for reference later delete when approp
		'<ValidateInput(False)>
		Function util_validate_mailmerge() As ActionResult
			Return View()
		End Function

#End Region


		Function Quickfind() As ActionResult
			Return View()
		End Function

		Function Filterselect() As ActionResult
			Return View()
		End Function

		<HttpPost()>
		Function filterselect_Submit(value As FormCollection)
			Dim sErrorMsg = ""

			' Only process the form submission if the referring page was the default page.
			' If it wasn't then redirect to the login page.
			' Read the information from the calling form.
			Dim sNextPage = Request.Form("txtGotoOptionPage")
			Dim sAction = Request.Form("txtGotoOptionAction")

			Session("optionScreenID") = Request.Form("txtGotoOptionScreenID")
			Session("optionTableID") = Request.Form("txtGotoOptionTableID")
			Session("optionViewID") = Request.Form("txtGotoOptionViewID")
			Session("optionOrderID") = Request.Form("txtGotoOptionOrderID")
			Session("optionRecordID") = Request.Form("txtGotoOptionRecordID")
			Session("optionFilterDef") = Request.Form("txtGotoOptionFilterDef")
			Session("optionFilterSQL") = Request.Form("txtGotoOptionFilterSQL")
			Session("optionValue") = Request.Form("txtGotoOptionValue")
			Session("optionLinkTableID") = Request.Form("txtGotoOptionLinkTableID")
			Session("optionLinkOrderID") = Request.Form("txtGotoOptionLinkOrderID")
			Session("optionLinkViewID") = Request.Form("txtGotoOptionLinkViewID")
			Session("optionLinkRecordID") = Request.Form("txtGotoOptionLinkRecordID")
			Session("optionColumnID") = Request.Form("txtGotoOptionColumnID")
			Session("optionLookupColumnID") = Request.Form("txtGotoOptionLookupColumnID")
			Session("optionLookupMandatory") = Request.Form("txtGotoOptionLookupMandatory")
			Session("optionLookupValue") = Request.Form("txtGotoOptionLookupValue")
			Session("optionFile") = Request.Form("txtGotoOptionFile")
			Session("optionExtension") = Request.Form("txtGotoOptionExtension")
			'Session("optionOLEOnServer") = Request.Form("txtGotoOptionOLEOnServer")
			Session("optionOLEType") = Request.Form("txtGotoOptionOLEType")
			Session("optionAction") = sAction
			Session("optionPageAction") = Request.Form("txtGotoOptionPageAction")
			Session("optionCourseTitle") = Request.Form("txtGotoOptionCourseTitle")
			Session("optionFirstRecPos") = Request.Form("txtGotoOptionFirstRecPos")
			Session("optionCurrentRecCount") = Request.Form("txtGotoOptionCurrentRecCount")
			Session("optionExprType") = Request.Form("txtGotoOptionExprType")
			Session("optionExprID") = Request.Form("txtGotoOptionExprID")
			Session("optionFunctionID") = Request.Form("txtGotoOptionFunctionID")
			Session("optionParameterIndex") = Request.Form("txtGotoOptionParameterIndex")


			If sAction = "CANCEL" Then
				' Go to the requested page.
				Session("errorMessage") = sErrorMsg
				Return RedirectToAction(sNextPage)
			End If

			If sAction = "SELECTFILTER" Then
				Session("errorMessage") = sErrorMsg

				' Go to the requested page.
				Return RedirectToAction(sNextPage)
			End If

			Return RedirectToAction(sNextPage)

		End Function

		Function tbAddFromWaitingListFind() As ActionResult
			Return View()
		End Function

		<HttpPost()>
	 Function tbAddFromWaitingListFind_Submit(value As FormCollection)

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			Dim sErrorMsg = ""
			Dim iTBResultCode = 0
			Dim sPreReqFails = ""

			' Only process the form submission if the referring page was the default page.
			' If it wasn't then redirect to the login page.

			' Read the information from the calling form.
			Dim sNextPage = Request.Form("txtGotoOptionPage")

			Dim sAction = Request.Form("txtGotoOptionAction")

			Session("optionScreenID") = Request.Form("txtGotoOptionScreenID")
			Session("optionTableID") = Request.Form("txtGotoOptionTableID")
			Session("optionViewID") = Request.Form("txtGotoOptionViewID")
			Session("optionOrderID") = Request.Form("txtGotoOptionOrderID")
			Session("optionRecordID") = Request.Form("txtGotoOptionRecordID")
			Session("optionFilterDef") = Request.Form("txtGotoOptionFilterDef")
			Session("optionFilterSQL") = Request.Form("txtGotoOptionFilterSQL")
			Session("optionValue") = Request.Form("txtGotoOptionValue")
			Session("optionLinkTableID") = Request.Form("txtGotoOptionLinkTableID")
			Session("optionLinkOrderID") = Request.Form("txtGotoOptionLinkOrderID")
			Session("optionLinkViewID") = Request.Form("txtGotoOptionLinkViewID")
			Session("optionLinkRecordID") = Request.Form("txtGotoOptionLinkRecordID")
			Session("optionColumnID") = Request.Form("txtGotoOptionColumnID")
			Session("optionLookupColumnID") = Request.Form("txtGotoOptionLookupColumnID")
			Session("optionLookupMandatory") = Request.Form("txtGotoOptionLookupMandatory")
			Session("optionLookupValue") = Request.Form("txtGotoOptionLookupValue")
			Session("optionFile") = Request.Form("txtGotoOptionFile")
			Session("optionExtension") = Request.Form("txtGotoOptionExtension")
			'Session("optionOLEOnServer") = Request.Form("txtGotoOptionOLEOnServer")
			Session("optionOLEType") = Request.Form("txtGotoOptionOLEType")
			Session("optionAction") = sAction
			Session("optionPageAction") = Request.Form("txtGotoOptionPageAction")
			Session("optionCourseTitle") = Request.Form("txtGotoOptionCourseTitle")
			Session("optionFirstRecPos") = Request.Form("txtGotoOptionFirstRecPos")
			Session("optionCurrentRecCount") = Request.Form("txtGotoOptionCurrentRecCount")
			Session("optionExprType") = Request.Form("txtGotoOptionExprType")
			Session("optionExprID") = Request.Form("txtGotoOptionExprID")
			Session("optionFunctionID") = Request.Form("txtGotoOptionFunctionID")
			Session("optionParameterIndex") = Request.Form("txtGotoOptionParameterIndex")
			Session("optionDefSelType") = Request.Form("txtGotoOptionDefSelType")

			If (sAction = "SELECTADDFROMWAITINGLIST_1") Then
				If NullSafeInteger(Session("optionRecordID")) > 0 Then
					' First pass after selecting the employee to book.
					' Get the user to choose whether to make the booking 'provisional'
					' or confirmed.
					If Session("TB_TBStatusPExists") Then
						Return RedirectToAction("tbStatusPrompt")
					Else
						sAction = "SELECTADDFROMWAITINGLIST_2"
						Session("optionAction") = sAction
						Session("optionLookupValue") = "B"
					End If
				End If
			End If

			If (sAction = "SELECTADDFROMWAITINGLIST_2") Then
				If NullSafeInteger(Session("optionRecordID")) > 0 Then
					If Len(sErrorMsg) = 0 Then
						' Validate the booking.					
						iTBResultCode = 0

						Try

							Dim prmResult = New SqlParameter("@piResultCode", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

							objDataAccess.ExecuteSP("sp_ASRIntValidateTrainingBooking" _
								, prmResult _
								, New SqlParameter("piEmpRecID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionLinkRecordID"))} _
								, New SqlParameter("piCourseRecID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionRecordID"))} _
								, New SqlParameter("psBookingStatus", SqlDbType.VarChar, -1) With {.Value = Session("optionLookupValue")} _
								, New SqlParameter("piTBRecID", SqlDbType.Int) With {.Value = 0})

							iTBResultCode = prmResult.Value

						Catch ex As Exception
							sErrorMsg = "Error validating training booking." & vbCrLf & FormatError(ex.Message)

						End Try

					End If
				End If
			End If

			' Go to the requested page.
			Session("TBResultCode") = iTBResultCode
			Session("errorMessage") = sErrorMsg
			Session("PreReqFails") = sPreReqFails	' This will be a sp output in the future along the lines of Bulkbooking
			Return RedirectToAction(sNextPage)

		End Function

		Function tbStatusPrompt() As ActionResult
			Return View()
		End Function

		Function tbBookCourseFind() As ActionResult
			Return View()
		End Function

		<HttpPost()>
		Function tbBookCourseFind_Submit(value As FormCollection)

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			Dim sErrorMsg = ""
			Dim iTBResultCode = 0

			' Only process the form submission if the referring page was the default page.
			' If it wasn't then redirect to the login page.
			' Read the information from the calling form.
			Dim sNextPage = Request.Form("txtGotoOptionPage")
			Dim sAction = Request.Form("txtGotoOptionAction")

			Session("optionScreenID") = Request.Form("txtGotoOptionScreenID")
			Session("optionTableID") = Request.Form("txtGotoOptionTableID")
			Session("optionViewID") = Request.Form("txtGotoOptionViewID")
			Session("optionOrderID") = Request.Form("txtGotoOptionOrderID")
			Session("optionRecordID") = Request.Form("txtGotoOptionRecordID")
			Session("optionFilterDef") = Request.Form("txtGotoOptionFilterDef")
			Session("optionFilterSQL") = Request.Form("txtGotoOptionFilterSQL")
			Session("optionValue") = Request.Form("txtGotoOptionValue")
			Session("optionLinkTableID") = Request.Form("txtGotoOptionLinkTableID")
			Session("optionLinkOrderID") = Request.Form("txtGotoOptionLinkOrderID")
			Session("optionLinkViewID") = Request.Form("txtGotoOptionLinkViewID")
			Session("optionLinkRecordID") = Request.Form("txtGotoOptionLinkRecordID")
			Session("optionColumnID") = Request.Form("txtGotoOptionColumnID")
			Session("optionLookupColumnID") = Request.Form("txtGotoOptionLookupColumnID")
			Session("optionLookupMandatory") = Request.Form("txtGotoOptionLookupMandatory")
			Session("optionLookupValue") = Request.Form("txtGotoOptionLookupValue")
			Session("optionFile") = Request.Form("txtGotoOptionFile")
			Session("optionExtension") = Request.Form("txtGotoOptionExtension")
			'Session("optionOLEOnServer") = Request.Form("txtGotoOptionOLEOnServer")
			Session("optionOLEType") = Request.Form("txtGotoOptionOLEType")
			Session("optionAction") = sAction
			Session("optionPageAction") = Request.Form("txtGotoOptionPageAction")
			Session("optionCourseTitle") = Request.Form("txtGotoOptionCourseTitle")
			Session("optionFirstRecPos") = Request.Form("txtGotoOptionFirstRecPos")
			Session("optionCurrentRecCount") = Request.Form("txtGotoOptionCurrentRecCount")
			Session("optionExprType") = Request.Form("txtGotoOptionExprType")
			Session("optionExprID") = Request.Form("txtGotoOptionExprID")
			Session("optionFunctionID") = Request.Form("txtGotoOptionFunctionID")
			Session("optionParameterIndex") = Request.Form("txtGotoOptionParameterIndex")
			Session("optionDefSelType") = Request.Form("txtGotoOptionDefSelType")

			If (sAction = "SELECTBOOKCOURSE_1") Then
				If NullSafeInteger(Session("optionRecordID")) > 0 Then
					' First pass after selecting the course to book.
					' Get the user to choose whether to make the booking 'provisional'
					' or confirmed.
					If Session("TB_TBStatusPExists") Then
						Return RedirectToAction("tbStatusPrompt")
					Else
						sAction = "SELECTBOOKCOURSE_2"
						Session("optionAction") = sAction
						Session("optionLookupValue") = "B"
					End If
				End If
			End If

			If (sAction = "SELECTBOOKCOURSE_2") Then
				If NullSafeInteger(Session("optionRecordID")) > 0 Then
					' Get the employee record ID from the given Waiting List record.
					Dim iEmpRecID = 0

					Try

						Dim prmTBEmployeeRecordID = New SqlParameter("piEmpRecordID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
						objDataAccess.ExecuteSP("sp_ASRIntGetEmpIDFromWLID", _
								prmTBEmployeeRecordID, _
								New SqlParameter("@piWLRecordID", SqlDbType.Int) With {.Value = CleanNumeric(NullSafeInteger(Session("optionRecordID")))})

						iEmpRecID = CInt(prmTBEmployeeRecordID.Value)

						If iEmpRecID = 0 Then
							sErrorMsg = "Error getting employee ID."
						End If

					Catch ex As Exception
						sErrorMsg = "Error getting employee ID." & vbCrLf & FormatError(Err.Description)

					End Try

					If Len(sErrorMsg) = 0 Then
						' Validate the booking.
						iTBResultCode = 0

						Try

							Dim prmResult = New SqlParameter("@piResultCode", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

							objDataAccess.ExecuteSP("sp_ASRIntValidateTrainingBooking" _
								, prmResult _
								, New SqlParameter("piEmpRecID", SqlDbType.Int) With {.Value = CleanNumeric(iEmpRecID)} _
								, New SqlParameter("piCourseRecID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionLinkRecordID"))} _
								, New SqlParameter("psBookingStatus", SqlDbType.VarChar, -1) With {.Value = Session("optionLookupValue")} _
								, New SqlParameter("piTBRecID", SqlDbType.Int) With {.Value = 0})

							iTBResultCode = prmResult.Value

						Catch ex As Exception
							sErrorMsg = "Error validating training booking." & vbCrLf & FormatError(ex.Message)

						End Try

					End If
				End If
			End If

			' Go to the requested page.
			Session("TBResultCode") = iTBResultCode
			Session("errorMessage") = sErrorMsg
			Return RedirectToAction(sNextPage)

		End Function

		Function tbBulkBooking() As ActionResult
			Return View()
		End Function

		<HttpPost()>
		Function tbBulkBooking_Submit(value As FormCollection)

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			Dim sErrorMsg = ""
			Dim iTBResultCode = 0
			Dim sPreReqFails = ""
			Dim sUnAvailFails = ""
			Dim sOverlapFails = ""
			Dim sOverBookFails = ""

			' Read the information from the calling form.
			Dim sNextPage = Request.Form("txtGotoOptionPage")
			Dim sAction = Request.Form("txtGotoOptionAction")

			Session("optionScreenID") = Request.Form("txtGotoOptionScreenID")
			Session("optionTableID") = Request.Form("txtGotoOptionTableID")
			Session("optionViewID") = Request.Form("txtGotoOptionViewID")
			Session("optionOrderID") = Request.Form("txtGotoOptionOrderID")
			Session("optionRecordID") = Request.Form("txtGotoOptionRecordID")
			Session("optionFilterDef") = Request.Form("txtGotoOptionFilterDef")
			Session("optionFilterSQL") = Request.Form("txtGotoOptionFilterSQL")
			Session("optionValue") = Request.Form("txtGotoOptionValue")
			Session("optionLinkTableID") = Request.Form("txtGotoOptionLinkTableID")
			Session("optionLinkOrderID") = Request.Form("txtGotoOptionLinkOrderID")
			Session("optionLinkViewID") = Request.Form("txtGotoOptionLinkViewID")
			Session("optionLinkRecordID") = Request.Form("txtGotoOptionLinkRecordID")
			Session("optionColumnID") = Request.Form("txtGotoOptionColumnID")
			Session("optionLookupColumnID") = Request.Form("txtGotoOptionLookupColumnID")
			Session("optionLookupMandatory") = Request.Form("txtGotoOptionLookupMandatory")
			Session("optionLookupValue") = Request.Form("txtGotoOptionLookupValue")
			Session("optionFile") = Request.Form("txtGotoOptionFile")
			Session("optionExtension") = Request.Form("txtGotoOptionExtension")
			'Session("optionOLEOnServer") = Request.Form("txtGotoOptionOLEOnServer")
			Session("optionOLEType") = Request.Form("txtGotoOptionOLEType")
			Session("optionAction") = sAction
			Session("optionPageAction") = Request.Form("txtGotoOptionPageAction")
			Session("optionCourseTitle") = Request.Form("txtGotoOptionCourseTitle")
			Session("optionFirstRecPos") = Request.Form("txtGotoOptionFirstRecPos")
			Session("optionCurrentRecCount") = Request.Form("txtGotoOptionCurrentRecCount")
			Session("optionExprType") = Request.Form("txtGotoOptionExprType")
			Session("optionExprID") = Request.Form("txtGotoOptionExprID")
			Session("optionFunctionID") = Request.Form("txtGotoOptionFunctionID")
			Session("optionParameterIndex") = Request.Form("txtGotoOptionParameterIndex")
			Session("optionDefSelType") = Request.Form("txtGotoOptionDefSelType")

			If (sAction = "SELECTBULKBOOKINGS") Then
				If Len(Session("optionLinkRecordID")) > 0 Then

					Try

						Dim prmResult = New SqlParameter("piResultCode", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
						Dim prmErrorMsg = New SqlParameter("psErrorMessage", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
						Dim prmPreRequisites = New SqlParameter("psWhoFailedPreReqCheck", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
						Dim prmAvailability = New SqlParameter("psWhoFailedUnavailabilityCheck", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
						Dim prmOverLapping = New SqlParameter("psWhoFailedOverlapCheck", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
						Dim prmOverBooking = New SqlParameter("psWhoFailedOverbookingCheck", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

						objDataAccess.ExecuteSP("sp_ASRIntValidateBulkBookings" _
							, New SqlParameter("piCourseRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionRecordID"))} _
							, New SqlParameter("psEmployeeRecordIDs", SqlDbType.VarChar, -1) With {.Value = Session("optionLinkRecordID")} _
							, New SqlParameter("psBookingStatus", SqlDbType.VarChar, -1) With {.Value = Session("optionLookupValue")} _
							, prmResult _
							, prmErrorMsg _
							, prmPreRequisites _
							, prmAvailability _
							, prmOverLapping _
							, prmOverBooking)

						iTBResultCode = prmResult.Value
						sPreReqFails = prmPreRequisites.Value.ToString()
						sUnAvailFails = prmAvailability.Value.ToString()
						sOverlapFails = prmOverLapping.Value.ToString()
						sOverBookFails = prmOverBooking.Value.ToString()

					Catch ex As Exception
						sErrorMsg = "Error validating training booking transfers." & vbCrLf & FormatError(ex.Message)
					End Try

				End If
			End If

			' Go to the requested page.
			Session("TBResultCode") = iTBResultCode
			Session("errorMessage") = sErrorMsg
			Session("PreReqFails") = sPreReqFails
			Session("UnAvailFails") = sUnAvailFails
			Session("OverlapFails") = sOverlapFails
			Session("OverBookFails") = sOverBookFails

			Return RedirectToAction(sNextPage)

		End Function

		Public Function tbBulkBookingSelectionMain() As ActionResult
			Return View()
		End Function


		<HttpPost()>
		Function tbBulkBookingSelectionData_Submit(value As FormCollection)

			On Error Resume Next

			Response.Expires = -1

			' Read the information from the calling form.
			'		session("action") = Request.Form("txtAction")
			Session("tableID") = Request.Form("txtTableID")
			Session("viewID") = Request.Form("txtViewID")
			Session("orderID") = Request.Form("txtOrderID")
			'		Session("columnID") = Request.Form("txtColumnID")
			Session("pageAction") = Request.Form("txtPageAction")
			Session("firstRecPos") = Request.Form("txtFirstRecPos")
			Session("currentRecCount") = Request.Form("txtCurrentRecCount")
			Session("locateValue") = Request.Form("txtGotoLocateValue")
			'		session("recordID") = Request.Form("txtRecordID")
			'		session("linkRecordID") = Request.Form("txtLinkRecordID")
			'		session("value") = Request.Form("txtValue")
			'		session("SQL") = Request.Form("txtSQL")
			'		session("promptSQL") = Request.Form("txtPromptSQL")
			Session("fromMenu") = Request.Form("txtGotoFromMenu")

			Session("tbSelectionDataLoading") = False

			' Go to the requested page.
			Return RedirectToAction("tbBulkBookingSelectionData")

		End Function

		Public Function tbBulkBookingSelectionData() As ActionResult
			Return View()
		End Function

		Public Function util_run_mailmerge_completed() As FileStreamResult

			Dim objMergeDocument As Code.MailMergeRun = Session("MailMerge_CompletedDocument")

			Return File(objMergeDocument.MergeDocument, "application/vnd.openxmlformats-officedocument.wordprocessingml.document" _
				, Path.GetFileName(objMergeDocument.OutputFileName))

		End Function

		Function promptedValues() As ActionResult
			Return View()
		End Function


		<HttpPost()>
		Function promptedValues_Submit(value As FormCollection)
			On Error Resume Next

			Session("filterID") = Request.Form("filterID")
			'Response.Write("<input type=""hidden"" id=filterID name=filterID value=" & Request.Form("filterID") & ">" & vbCrLf)

			Dim sPrompts
			Dim aPrompts(1, 0)
			Dim j = 0
			sPrompts = ""
			' ReDim Preserve aPrompts(1, 0)
			For i = 0 To Request.Form.Count - 1
				Dim sKey = Request.Form.Keys(i)
				If ((UCase(Left(sKey, 7)) = "PROMPT_") And (Mid(sKey, 8, 1) <> "3")) Or _
					(UCase(Left(sKey, 10)) = "PROMPTCHK_") Then
					ReDim Preserve aPrompts(1, j)

					If (UCase(Left(sKey, 10)) = "PROMPTCHK_") Then
						aPrompts(0, j) = "prompt_3_" & Mid(sKey, 11)
						aPrompts(1, j) = UCase(Request.Form.Item(i))
					Else
						aPrompts(0, j) = sKey
						Select Case Mid(sKey, 8, 1)
							Case "2"
								' Numeric. Replace locale decimal point with '.'
								aPrompts(1, j) = Replace(Request.Form.Item(i), Session("LocaleDecimalSeparator"), ".")
							Case "4"
								' Date. Reformat to match SQL's mm/dd/yyyy format.
								aPrompts(1, j) = convertLocaleDateToSQL(Request.Form.Item(i))
							Case Else
								aPrompts(1, j) = Request.Form.Item(i)
						End Select
					End If

					sPrompts = sPrompts & aPrompts(0, j) & vbTab & aPrompts(1, j) & vbTab

					j += 1
				End If
			Next

			Session("filterIDvalue") = Request.Form("filterID")
			Session("promptsvalue") = sPrompts

			'Response.Write("<input type=""hidden"" id=prompts name=prompts value=""" & sPrompts & """>" & vbCrLf)

			Return RedirectToAction("promptedValues_completed")

		End Function


		Function promptedValues_completed() As ActionResult
			Return View()
		End Function

		Function tbTransferBookingFind() As ActionResult
			Return View()
		End Function



		<HttpPost()>
		Function tbTransferBookingFind_Submit(value As FormCollection)

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			Dim sErrorMsg = ""
			Dim iTBResultCode = 0

			' Read the information from the calling form.
			Dim sNextPage = Request.Form("txtGotoOptionPage")
			Dim sAction = Request.Form("txtGotoOptionAction")

			Session("optionScreenID") = Request.Form("txtGotoOptionScreenID")
			Session("optionTableID") = Request.Form("txtGotoOptionTableID")
			Session("optionViewID") = Request.Form("txtGotoOptionViewID")
			Session("optionOrderID") = Request.Form("txtGotoOptionOrderID")
			Session("optionRecordID") = Request.Form("txtGotoOptionRecordID")
			Session("optionFilterDef") = Request.Form("txtGotoOptionFilterDef")
			Session("optionFilterSQL") = Request.Form("txtGotoOptionFilterSQL")
			Session("optionValue") = Request.Form("txtGotoOptionValue")
			Session("optionLinkTableID") = Request.Form("txtGotoOptionLinkTableID")
			Session("optionLinkOrderID") = Request.Form("txtGotoOptionLinkOrderID")
			Session("optionLinkViewID") = Request.Form("txtGotoOptionLinkViewID")
			Session("optionLinkRecordID") = Request.Form("txtGotoOptionLinkRecordID")
			Session("optionColumnID") = Request.Form("txtGotoOptionColumnID")
			Session("optionLookupColumnID") = Request.Form("txtGotoOptionLookupColumnID")
			Session("optionLookupMandatory") = Request.Form("txtGotoOptionLookupMandatory")
			Session("optionLookupValue") = Request.Form("txtGotoOptionLookupValue")
			Session("optionFile") = Request.Form("txtGotoOptionFile")
			Session("optionExtension") = Request.Form("txtGotoOptionExtension")
			'Session("optionOLEOnServer") = Request.Form("txtGotoOptionOLEOnServer")
			Session("optionOLEType") = Request.Form("txtGotoOptionOLEType")
			Session("optionAction") = sAction
			Session("optionPageAction") = Request.Form("txtGotoOptionPageAction")
			Session("optionCourseTitle") = Request.Form("txtGotoOptionCourseTitle")
			Session("optionFirstRecPos") = Request.Form("txtGotoOptionFirstRecPos")
			Session("optionCurrentRecCount") = Request.Form("txtGotoOptionCurrentRecCount")
			Session("optionExprType") = Request.Form("txtGotoOptionExprType")
			Session("optionExprID") = Request.Form("txtGotoOptionExprID")
			Session("optionFunctionID") = Request.Form("txtGotoOptionFunctionID")
			Session("optionParameterIndex") = Request.Form("txtGotoOptionParameterIndex")
			Session("optionDefSelType") = Request.Form("txtGotoOptionDefSelType")

			If (sAction = "SELECTTRANSFERBOOKING_1") Then
				If NullSafeInteger(Session("optionRecordID")) > 0 Then
					' Get the employee record ID from the given Training Booking record.
					Dim iEmpRecID = 0

					Try

						Dim prmEmployeeRecordID = New SqlParameter("piEmpRecordID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

						objDataAccess.ExecuteSP("sp_ASRIntGetEmpIDFromTBID" _
							, prmEmployeeRecordID _
							, New SqlParameter("piTBRecordID", SqlDbType.Int) With {.Value = CleanNumeric(NullSafeInteger(Session("optionRecordID")))})

						iEmpRecID = prmEmployeeRecordID.Value

						If iEmpRecID = 0 Then
							sErrorMsg = "Error getting employee ID."
						End If

					Catch ex As Exception
						sErrorMsg = "Error getting employee ID." & vbCrLf & FormatError(ex.Message)

					End Try


					If Len(sErrorMsg) = 0 Then
						' Validate the booking.
						iTBResultCode = 0

						Try

							Dim prmResult = New SqlParameter("@piResultCode", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

							objDataAccess.ExecuteSP("sp_ASRIntValidateTrainingBooking" _
								, prmResult _
								, New SqlParameter("piEmpRecID", SqlDbType.Int) With {.Value = CleanNumeric(iEmpRecID)} _
								, New SqlParameter("piCourseRecID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionLinkRecordID"))} _
								, New SqlParameter("psBookingStatus", SqlDbType.VarChar, -1) With {.Value = Session("optionLookupValue")} _
								, New SqlParameter("piTBRecID", SqlDbType.Int) With {.Value = 0})

							iTBResultCode = prmResult.Value

						Catch ex As Exception
							sErrorMsg = "Error validating training booking." & vbCrLf & FormatError(ex.Message)

						End Try

					End If
				End If
			End If

			' Go to the requested page.
			Session("TBResultCode") = iTBResultCode
			Session("errorMessage") = sErrorMsg
			Return RedirectToAction(sNextPage)

		End Function

		<ValidateInput(False)>
		Function util_run_outputoptions() As ActionResult

			Session("CT_Mode") = Request("txtMode")
			Session("OutputOptions_Format") = Request("txtFormat")
			Session("OutputOptions_Screen") = Request("txtScreen")
			Session("OutputOptions_Printer") = Request("txtPrinter")
			Session("OutputOptions_PrinterName") = Request("txtPrinterName")
			Session("OutputOptions_Save") = Request("txtSave")
			Session("OutputOptions_SaveExisting") = Request("txtSaveExisting")
			Session("OutputOptions_Email") = Request("txtEmail")
			Session("OutputOptions_EmailGroupID") = Request("txtEmailGroupID")
			Session("OutputOptions_EmailGroup") = Request("txtEmailGroup")
			Session("OutputOptions_EmailSubject") = Request("txtEmailSubject")
			Session("OutputOptions_EmailAttachAs") = Request("txtEmailAttachAs")
			Session("OutputOptions_Filename") = Request("txtFilename")

			Session("utiltype") = Request.Form("txtUtilType")

			Return View()
		End Function

		Function tbTransferCourseFind() As ActionResult
			Return View()
		End Function

		<HttpPost()>
	 Function tbTransferCourseFind_Submit(value As FormCollection)

			Dim sErrorMsg = ""
			Dim iTBResultCode = 0

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			' Read the information from the calling form.
			Dim sNextPage = Request.Form("txtGotoOptionPage")
			Dim sAction = Request.Form("txtGotoOptionAction")

			Session("optionScreenID") = Request.Form("txtGotoOptionScreenID")
			Session("optionTableID") = Request.Form("txtGotoOptionTableID")
			Session("optionViewID") = Request.Form("txtGotoOptionViewID")
			Session("optionOrderID") = Request.Form("txtGotoOptionOrderID")
			Session("optionRecordID") = Request.Form("txtGotoOptionRecordID")
			Session("optionFilterDef") = Request.Form("txtGotoOptionFilterDef")
			Session("optionFilterSQL") = Request.Form("txtGotoOptionFilterSQL")
			Session("optionValue") = Request.Form("txtGotoOptionValue")
			Session("optionLinkTableID") = Request.Form("txtGotoOptionLinkTableID")
			Session("optionLinkOrderID") = Request.Form("txtGotoOptionLinkOrderID")
			Session("optionLinkViewID") = Request.Form("txtGotoOptionLinkViewID")
			Session("optionLinkRecordID") = Request.Form("txtGotoOptionLinkRecordID")
			Session("optionColumnID") = Request.Form("txtGotoOptionColumnID")
			Session("optionLookupColumnID") = Request.Form("txtGotoOptionLookupColumnID")
			Session("optionLookupMandatory") = Request.Form("txtGotoOptionLookupMandatory")
			Session("optionLookupValue") = Request.Form("txtGotoOptionLookupValue")
			Session("optionFile") = Request.Form("txtGotoOptionFile")
			Session("optionExtension") = Request.Form("txtGotoOptionExtension")
			'Session("optionOLEOnServer") = Request.Form("txtGotoOptionOLEOnServer")
			Session("optionOLEType") = Request.Form("txtGotoOptionOLEType")
			Session("optionAction") = sAction
			Session("optionFirstRecPos") = Request.Form("txtGotoOptionFirstRecPos")
			Session("optionCurrentRecCount") = Request.Form("txtGotoOptionCurrentRecCount")
			Session("optionPageAction") = Request.Form("txtGotoOptionPageAction")
			Session("optionCourseTitle") = Request.Form("txtGotoOptionCourseTitle")
			Session("optionExprType") = Request.Form("txtGotoOptionExprType")
			Session("optionExprID") = Request.Form("txtGotoOptionExprID")
			Session("optionFunctionID") = Request.Form("txtGotoOptionFunctionID")
			Session("optionParameterIndex") = Request.Form("txtGotoOptionParameterIndex")
			Session("optionDefSelType") = Request.Form("txtGotoOptionDefSelType")

			If sAction = "" Then
				' Go to the requested page.
				Return RedirectToAction(sNextPage)
			End If

			If sAction = "SELECTTRANSFERCOURSE" Then

				If Session("optionLinkRecordID") > 0 Then
					' Validate the booking transfers.

					Try

						Dim prmResult = New SqlParameter("@piResultCode", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
						Dim prmErrorMessage = New SqlParameter("@psErrorMessage", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

						objDataAccess.ExecuteSP("sp_ASRIntValidateTransfers" _
							, New SqlParameter("piEmployeeTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("TB_EmpTableID"))} _
							, New SqlParameter("piCourseTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("TB_CourseTableID"))} _
							, New SqlParameter("piCourseRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionRecordID"))} _
							, New SqlParameter("piTransferCourseRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionLinkRecordID"))} _
							, New SqlParameter("piTrainBookTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("TB_TBTableID"))} _
							, New SqlParameter("piTrainBookStatusColumnID", SqlDbType.Int) With {.Value = CleanNumeric(Session("TB_TBStatusColumnID"))} _
							, prmResult _
							, prmErrorMessage)

						If (Len(sErrorMsg) = 0) And Len(prmErrorMessage.Value.ToString()) > 0 Then
							sErrorMsg = "Error validating training booking transfers." & vbCrLf & prmErrorMessage.Value.ToString
						End If

						iTBResultCode = prmResult.Value

					Catch ex As Exception
						sErrorMsg = "Error validating training booking transfers." & vbCrLf & FormatError(ex.Message)

					End Try

				End If

				Session("TBResultCode") = iTBResultCode
				Session("errorMessage") = sErrorMsg
				Return RedirectToAction(sNextPage)
			End If

		End Function

		Function orderselect() As ActionResult
			Return View()
		End Function

		<HttpPost()>
	 Function orderselect_Submit(value As FormCollection)

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			Dim sErrorMsg = ""

			' Read the information from the calling form.
			Dim lngScreenID = CleanNumeric(Request.Form("txtGotoOptionScreenID"))
			Dim lngViewID = CleanNumeric(Request.Form("txtGotoOptionViewID"))
			Dim lngOrderID = CleanNumeric(Request.Form("txtGotoOptionOrderID"))
			Dim sNextPage = Request.Form("txtGotoOptionPage")
			Dim sAction = Request.Form("txtGotoOptionAction")

			Session("optionScreenID") = lngScreenID
			Session("optionTableID") = Request.Form("txtGotoOptionTableID")
			Session("optionViewID") = lngViewID
			Session("optionOrderID") = lngOrderID
			Session("optionRecordID") = Request.Form("txtGotoOptionRecordID")
			Session("optionFilterDef") = Request.Form("txtGotoOptionFilterDef")
			Session("optionFilterSQL") = Request.Form("txtGotoOptionFilterSQL")
			Session("optionValue") = Request.Form("txtGotoOptionValue")
			Session("optionLinkTableID") = Request.Form("txtGotoOptionLinkTableID")
			Session("optionLinkOrderID") = Request.Form("txtGotoOptionLinkOrderID")
			Session("optionLinkViewID") = Request.Form("txtGotoOptionLinkViewID")
			Session("optionLinkRecordID") = Request.Form("txtGotoOptionLinkRecordID")
			Session("optionColumnID") = Request.Form("txtGotoOptionColumnID")
			Session("optionLookupColumnID") = Request.Form("txtGotoOptionLookupColumnID")
			Session("optionLookupMandatory") = Request.Form("txtGotoOptionLookupMandatory")
			Session("optionLookupValue") = Request.Form("txtGotoOptionLookupValue")
			Session("optionFile") = Request.Form("txtGotoOptionFile")
			Session("optionExtension") = Request.Form("txtGotoOptionExtension")
			'Session("optionOLEOnServer") = Request.Form("txtGotoOptionOLEOnServer")
			Session("optionOLEType") = Request.Form("txtGotoOptionOLEType")
			Session("optionAction") = sAction
			Session("optionPageAction") = Request.Form("txtGotoOptionPageAction")
			Session("optionCourseTitle") = Request.Form("txtGotoOptionCourseTitle")
			Session("optionAction") = sAction
			Session("orderID") = lngOrderID
			Session("optionFirstRecPos") = Request.Form("txtGotoOptionFirstRecPos")
			Session("optionCurrentRecCount") = Request.Form("txtGotoOptionCurrentRecCount")
			Session("optionExprType") = Request.Form("txtGotoOptionExprType")
			Session("optionExprID") = Request.Form("txtGotoOptionExprID")
			Session("optionFunctionID") = Request.Form("txtGotoOptionFunctionID")
			Session("optionParameterIndex") = Request.Form("txtGotoOptionParameterIndex")


			If sAction = "CANCEL" Then
				' Go to the requested page.
				Session("errorMessage") = sErrorMsg
				Return RedirectToAction(sNextPage)
			End If

			If sAction = "SELECTORDER" Then

				Try
					Dim prmFromDef = New SqlParameter("psFromDef", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

					objDataAccess.ExecuteSP("sp_ASRIntGetOrderSQL" _
							, New SqlParameter("piScreenID", SqlDbType.Int) With {.Value = lngScreenID} _
							, New SqlParameter("piViewID", SqlDbType.Int) With {.Value = lngViewID} _
							, New SqlParameter("piOrderID", SqlDbType.Int) With {.Value = lngOrderID} _
							, prmFromDef)

					Session("fromDef") = prmFromDef.Value

				Catch ex As Exception
					sErrorMsg = "Error retrieving the new order definition." & vbCrLf & FormatError(Err.Description)
					Session("errorMessage") = sErrorMsg

				End Try

				' Go to the requested page.
				Return RedirectToAction(sNextPage)
			End If

			Return RedirectToAction(sNextPage)

		End Function


		Function lookupFind() As ActionResult
			Return View()
		End Function

		<HttpPost()>
	 Function lookupFind_Submit(value As FormCollection)

			On Error Resume Next

			Dim sErrorMsg = ""

			' Read the information from the calling form.
			Dim sNextPage = Request.Form("txtGotoOptionPage")
			Dim sAction = Request.Form("txtGotoOptionAction")

			Session("optionScreenID") = Request.Form("txtGotoOptionScreenID")
			Session("optionTableID") = Request.Form("txtGotoOptionTableID")
			Session("optionViewID") = Request.Form("txtGotoOptionViewID")
			Session("optionOrderID") = Request.Form("txtGotoOptionOrderID")
			Session("optionRecordID") = Request.Form("txtGotoOptionRecordID")
			Session("optionFilterDef") = Request.Form("txtGotoOptionFilterDef")
			Session("optionFilterSQL") = Request.Form("txtGotoOptionFilterSQL")
			Session("optionValue") = Request.Form("txtGotoOptionValue")
			Session("optionLinkTableID") = Request.Form("txtGotoOptionLinkTableID")
			Session("optionLinkOrderID") = Request.Form("txtGotoOptionLinkOrderID")
			Session("optionLinkViewID") = Request.Form("txtGotoOptionLinkViewID")
			Session("optionRecordID") = Request.Form("txtGotoOptionLinkRecordID")
			Session("optionColumnID") = Request.Form("txtGotoOptionColumnID")
			Session("optionLookupColumnID") = Request.Form("txtGotoOptionLookupColumnID")
			Session("optionLookupMandatory") = Request.Form("txtGotoOptionLookupMandatory")
			Session("optionLookupValue") = Request.Form("txtGotoOptionLookupValue")
			Session("optionFile") = Request.Form("txtGotoOptionFile")
			Session("optionExtension") = Request.Form("txtGotoOptionExtension")
			'Session("optionOLEOnServer") = Request.Form("txtGotoOptionOLEOnServer")
			Session("optionOLEType") = Request.Form("txtGotoOptionOLEType")
			Session("optionAction") = sAction
			Session("optionPageAction") = Request.Form("txtGotoOptionPageAction")
			Session("optionCourseTitle") = Request.Form("txtGotoOptionCourseTitle")
			Session("optionFirstRecPos") = Request.Form("txtGotoOptionFirstRecPos")
			Session("optionCurrentRecCount") = Request.Form("txtGotoOptionCurrentRecCount")
			Session("optionExprType") = Request.Form("txtGotoOptionExprType")
			Session("optionExprID") = Request.Form("txtGotoOptionExprID")
			Session("optionFunctionID") = Request.Form("txtGotoOptionFunctionID")
			Session("optionParameterIndex") = Request.Form("txtGotoOptionParameterIndex")

			If sAction = "" Then
				' Go to the requested page.
				'Return RedirectToAction(sNextPage)
			End If

			If sAction = "CANCEL" Then
				' Go to the requested page.
				Session("errorMessage") = sErrorMsg
				'Return RedirectToAction(sNextPage)
			End If

			If sAction = "SELECTLOOKUP" Then
				Session("errorMessage") = sErrorMsg

				' Go to the requested page.
				'Return RedirectToAction(sNextPage)
			End If

			' Go to the requested page.
			Return RedirectToAction(sNextPage)

		End Function

		Function themeEditor() As PartialViewResult
			Return PartialView()
		End Function

		Function linkFind() As ActionResult
			Return View()
		End Function

		<HttpPost()>
		Function linkFind_Submit(value As FormCollection)
			On Error Resume Next

			Dim sErrorMsg As String = ""
			Dim sNextPage As String, sAction As String

			' Read the information from the calling form.
			sNextPage = Request.Form("txtGotoOptionPage")
			sAction = Request.Form("txtGotoOptionAction")

			Session("optionScreenID") = Request.Form("txtGotoOptionScreenID")
			Session("optionTableID") = Request.Form("txtGotoOptionTableID")
			Session("optionViewID") = Request.Form("txtGotoOptionViewID")
			Session("optionOrderID") = Request.Form("txtGotoOptionOrderID")
			Session("optionRecordID") = Request.Form("txtGotoOptionRecordID")
			Session("optionFilterDef") = Request.Form("txtGotoOptionFilterDef")
			Session("optionFilterSQL") = Request.Form("txtGotoOptionFilterSQL")
			Session("optionValue") = Request.Form("txtGotoOptionValue")
			Session("optionLinkTableID") = Request.Form("txtGotoOptionLinkTableID")
			Session("optionLinkOrderID") = Request.Form("txtGotoOptionLinkOrderID")
			Session("optionLinkViewID") = Request.Form("txtGotoOptionLinkViewID")
			Session("optionRecordID") = Request.Form("txtGotoOptionLinkRecordID")
			Session("optionColumnID") = Request.Form("txtGotoOptionColumnID")
			Session("optionLookupColumnID") = Request.Form("txtGotoOptionLookupColumnID")
			Session("optionLookupMandatory") = Request.Form("txtGotoOptionLookupMandatory")
			Session("optionLookupValue") = Request.Form("txtGotoOptionLookupValue")
			Session("optionFile") = Request.Form("txtGotoOptionFile")
			Session("optionExtension") = Request.Form("txtGotoOptionExtension")
			'Session("optionOLEOnServer") = Request.Form("txtGotoOptionOLEOnServer")
			Session("optionOLEType") = Request.Form("txtGotoOptionOLEType")
			Session("optionAction") = sAction
			Session("optionPageAction") = Request.Form("txtGotoOptionPageAction")
			Session("optionCourseTitle") = Request.Form("txtGotoOptionCourseTitle")
			Session("optionFirstRecPos") = Request.Form("txtGotoOptionFirstRecPos")
			Session("optionCurrentRecCount") = Request.Form("txtGotoOptionCurrentRecCount")
			Session("optionExprType") = Request.Form("txtGotoOptionExprType")
			Session("optionExprID") = Request.Form("txtGotoOptionExprID")
			Session("optionFunctionID") = Request.Form("txtGotoOptionFunctionID")
			Session("optionParameterIndex") = Request.Form("txtGotoOptionParameterIndex")

			If sAction = "CANCEL" Or sAction = "SELECTLINK" Then
				' Go to the requested page.
				Session("errorMessage") = sErrorMsg
			End If

			Return RedirectToAction(sNextPage)

		End Function

		Function oleFind() As ActionResult

			If Session("optionOLEType") > 1 Then
				Dim objOLE As HR.Intranet.Server.Ole = Session("OLEObject")
				Dim sFile As String = ""

				If NullSafeString(Session("optionFile")) <> vbNullString Then
					sFile = Session("optionFile")
				End If

				objOLE.IsPhoto = False
				objOLE.OLEType = Session("optionOLEType")
				objOLE.DisplayFilename = Path.GetFileName(sFile)
				objOLE.FileName = sFile
				Session("OLEObject") = objOLE
				objOLE = Nothing
			End If

			Return View()
		End Function

		<HttpPost()>
		Public Function importTheme_Submit(themeFile As HttpPostedFileBase) As ActionResult

			If themeFile.ContentLength > 0 Then

				Dim validThemes As New Dictionary(Of String, String)
				validThemes.Add("Aliceblue", "#F0F8FF")
				validThemes.Add("Antiquewhite", "#FAEBD7")
				validThemes.Add("Aqua", "#00ffff")
				validThemes.Add("Azure", "#F0FFFF")
				validThemes.Add("Black", "#000000")
				validThemes.Add("Blanco", "#FFFFFF")
				validThemes.Add("Blue", "#6699CC")
				validThemes.Add("Burlywood", "#DEB887")
				validThemes.Add("Chocolate", "#D2691E")
				validThemes.Add("Damson", "#7D388A")
				validThemes.Add("Darkgray", "#A9A9A9")
				validThemes.Add("Darkkhaki", "#BDB76B")
				validThemes.Add("Darkorange", "#FF8C00")
				validThemes.Add("Darkseagreen", "#8FBC8B")
				validThemes.Add("Darkviolet", "#9400D3")
				validThemes.Add("DeepRed", "#C90016")
				validThemes.Add("DeepSkyBlue", "#00BFFF")
				validThemes.Add("DodgerBlue", "#1E90FF")
				validThemes.Add("Forestgreen", "#228B22")
				validThemes.Add("Gold", "#FFD700")
				validThemes.Add("GreySkyBlue", "#DEE7EF")
				validThemes.Add("Ivy", "#A6B540")
				validThemes.Add("LightSkyBlue", "#87CEFA")
				validThemes.Add("Limegreen", "#32CD32")
				validThemes.Add("Maroon", "#700017")
				validThemes.Add("MidnightNavy", "#330066")
				validThemes.Add("Navy", "#000080")
				validThemes.Add("Navy2", "#000099")
				validThemes.Add("Olive", "#CCCC99")
				validThemes.Add("PantoneBlue", "#003F6E")
				validThemes.Add("PantoneGold", "#F7C046")
				validThemes.Add("Raspberry", "#C71444")
				validThemes.Add("Red", "#CC3300")
				validThemes.Add("Red2", "#FF0000")
				validThemes.Add("RichGrey", "#807A6E")
				validThemes.Add("RipeTomato", "#DF0029")
				validThemes.Add("Teal", "#008080")
				validThemes.Add("Teal2", "#009999")
				validThemes.Add("TuscanOrange", "#F39900")
				validThemes.Add("VioletBlue", "#B0B2F5")
				validThemes.Add("VioletGrey", "#CFCCE5")
				validThemes.Add("VioletGreyer", "#C8C9E4")

				Dim buffer As Byte() = New Byte(themeFile.InputStream.Length - 1) {}
				Dim offset As Integer = 0
				Dim cnt As Integer = 0
				While (InlineAssignHelper(cnt, themeFile.InputStream.Read(buffer, offset, 10))) > 0
					offset += cnt
				End While

				Dim ms As MemoryStream = New MemoryStream(buffer)

				Dim configFile As New Dictionary(Of String, String)

				Using myReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(ms)
					myReader.TextFieldType = FileIO.FieldType.Delimited
					myReader.SetDelimiters("=")

					Dim currentRow As String()
					While Not myReader.EndOfData
						Try
							currentRow = myReader.ReadFields()
							If currentRow.Length > 1 Then
								configFile.Add(currentRow.GetValue(0), ConvertConfigValue(currentRow.GetValue(1), validThemes))
							End If
						Catch
						End Try
					End While
				End Using

				' we now have a dictionary of key/value pairs from the old config file.
				If configFile.Count > 0 Then

					Dim cssOutput As New StringBuilder()

					cssOutput.AppendLine(CssCheck(".ui-widget-header { background-color: " & configFile("generaltheme") & "}", configFile("generaltheme")))
					cssOutput.AppendLine(".ui-widget-header { background-image: none}")

					cssOutput.AppendLine(CssCheck(".hypertextlinktextseparator { background-color: " & configFile("generaltheme") & "!important;background-image: none!important;}", configFile("generaltheme")))
					cssOutput.AppendLine(CssCheck(".hypertextlinktext { background-color: " & configFile("generaltheme") & "!important;background-image: none!important;}", configFile("generaltheme")))

					cssOutput.AppendLine(CssCheck(".hypertextlinkseparator-font { font-family: " & configFile("hypertextlinkseparator-font") & "}", configFile("hypertextlinkseparator-font")))
					cssOutput.AppendLine(CssCheck(".hypertextlinkseparator-colour { color: " & configFile("hypertextlinkseparator-colour") & "}", configFile("hypertextlinkseparator-colour")))
					cssOutput.AppendLine(CssCheck(".hypertextlinkseparator-size { font-size: " & configFile("hypertextlinkseparator-size") & "pt}", configFile("hypertextlinkseparator-size")))
					cssOutput.AppendLine(CssCheck(".hypertextlinkseparator-bold { font-weight: " & configFile("hypertextlinkseparator-bold") & "}", configFile("hypertextlinkseparator-bold")))
					cssOutput.AppendLine(CssCheck(".hypertextlinkseparator-italics { font-style: " & configFile("hypertextlinkseparator-italics") & "}", configFile("hypertextlinkseparator-italics")))

					cssOutput.AppendLine(CssCheck(".hypertextlinktext-font { font-family: " & configFile("hypertextlinktext-font") & "}", configFile("hypertextlinktext-font")))
					cssOutput.AppendLine(CssCheck(".hypertextlinktext-colour { color: " & configFile("hypertextlinktext-colour") & "!important;}", configFile("hypertextlinktext-colour")))
					cssOutput.AppendLine(CssCheck(".hypertextlinktext-size { font-size: " & configFile("hypertextlinktext-size") & "pt}", configFile("hypertextlinktext-size")))
					cssOutput.AppendLine(CssCheck(".hypertextlinktext-bold { font-weight: " & configFile("hypertextlinktext-bold") & "}", configFile("hypertextlinktext-bold")))
					cssOutput.AppendLine(CssCheck(".hypertextlinktext-italics { font-style: " & configFile("hypertextlinktext-italics") & "}", configFile("hypertextlinktext-italics")))

					cssOutput.AppendLine(CssCheck(".hypertextlinktext-highlightcolour:hover { background-color: " & configFile("hypertextlinktext-highlightcolour") & "}", configFile("hypertextlinktext-highlightcolour")))

					cssOutput.AppendLine(CssCheck(".linkspageprompttext-font { font-family: " & configFile("linkspageprompttext-font") & "}", configFile("linkspageprompttext-font")))
					cssOutput.AppendLine(CssCheck(".linkspageprompttext-colour { color: " & configFile("linkspageprompttext-colour") & "!important;}", configFile("linkspageprompttext-colour")))
					cssOutput.AppendLine(CssCheck(".linkspageprompttext-size { font-size: " & configFile("linkspageprompttext-size") & "pt}", configFile("linkspageprompttext-size")))
					cssOutput.AppendLine(CssCheck(".linkspageprompttext-bold { font-weight: " & configFile("linkspageprompttext-bold") & "}", configFile("linkspageprompttext-bold")))
					cssOutput.AppendLine(CssCheck(".linkspageprompttext-italics { font-style: " & configFile("linkspageprompttext-italics") & "}", configFile("linkspageprompttext-italics")))

					If configFile("linkspagebutton-displaytype").ToLower() <> "rounded" Then
						cssOutput.AppendLine(".linkspagebutton-displaytype { border-radius: 0!important;}")
					End If
					cssOutput.AppendLine(CssCheck(".linkspagebuttontext-alignment { float: none; text-align: " & configFile("linkspagebuttontext-alignment") & "}", configFile("linkspagebuttontext-alignment")))
					cssOutput.AppendLine(CssCheck(".linkspagebutton-colourtheme { background-color: " & configFile("linkspagebutton-colourtheme") & "; padding-top: 0!important;padding-bottom: 0!important;margin-bottom: 2px!important;}", configFile("linkspagebutton-colourtheme")))

					cssOutput.AppendLine(CssCheck(".linkspagebuttonseparator-font { font-family: " & configFile("linkspagebuttonseparator-font") & "}", configFile("linkspagebuttonseparator-font")))
					cssOutput.AppendLine(CssCheck(".linkspagebuttonseparator-colour { color: " & configFile("linkspagebuttonseparator-colour") & "!important;}", configFile("linkspagebuttonseparator-colour")))
					cssOutput.AppendLine(CssCheck(".linkspagebuttonseparator-size { font-size: " & configFile("linkspagebuttonseparator-size") & "pt}", configFile("linkspagebuttonseparator-size")))
					cssOutput.AppendLine(CssCheck(".linkspagebuttonseparator-bold { font-weight: " & configFile("linkspagebuttonseparator-bold") & "}", configFile("linkspagebuttonseparator-bold")))
					cssOutput.AppendLine(CssCheck(".linkspagebuttonseparator-italics { font-style: " & configFile("linkspagebuttonseparator-italics") & "}", configFile("linkspagebuttonseparator-italics")))
					cssOutput.AppendLine(CssCheck(".linkspagebuttonseparator-bordercolour { background-color: " & configFile("linkspagebuttonseparator-bordercolour") & "!important; background-image: none!important;}", configFile("linkspagebuttonseparator-bordercolour")))
					cssOutput.AppendLine(CssCheck(".linkspagebuttonseparator-alignment { float: none; padding-left: 0!important; text-align: " & configFile("linkspagebuttonseparator-alignment") & "}", configFile("linkspagebuttonseparator-alignment")))
					cssOutput.AppendLine(".ui-accordion-header { border-radius: 0!important;}")


					cssOutput.AppendLine(CssCheck(".linkspagebuttontext-font { font-family: " & configFile("linkspagebuttontext-font") & "}", configFile("linkspagebuttontext-font")))
					cssOutput.AppendLine(CssCheck(".linkspagebuttontext-colour { color: " & configFile("linkspagebuttontext-colour") & "!important;}", configFile("linkspagebuttontext-colour")))
					cssOutput.AppendLine(CssCheck(".linkspagebuttontext-size { font-size: " & configFile("linkspagebuttontext-size") & "pt}", configFile("linkspagebuttontext-size")))
					cssOutput.AppendLine(CssCheck(".linkspagebuttontext-bold { font-weight: " & configFile("linkspagebuttontext-bold") & "}", configFile("linkspagebuttontext-bold")))
					cssOutput.AppendLine(CssCheck(".linkspagebuttontext-italics { font-style: " & configFile("linkspagebuttontext-italics") & "}", configFile("linkspagebuttontext-italics")))


					' output to css file.
					Using cssFile As New StreamWriter(Server.MapPath("~/Content/DashboardStyles/themes/upgraded.css"))
						cssFile.Write(cssOutput.ToString())
					End Using

				End If

			End If

			Return RedirectToAction("Main", "Home", New With {.SSIMode = True})

		End Function

		Function CssCheck(cssString As String, configValue As String) As String
			If NullSafeString(configValue).Length > 0 Then
				Return cssString
			End If

			Return vbNullString

		End Function

		Function ConvertConfigValue(configValue As String, ByRef validThemes As Dictionary(Of String, String)) As String

			' remove hash for hex colours
			' If configValue.StartsWith("#") Then Return configValue.Substring(1)

			' check for it being a theme colour
			If validThemes.ContainsKey(configValue) Then Return validThemes(configValue)

			' try to convert from known name
			If Color.FromName(configValue).IsKnownColor Then
				Try
					Dim c As Color = Color.FromName(configValue)
					'Return c.ToArgb().ToString()
					Return "#" & ColorTranslator.FromHtml(String.Format("#{0:X2}{1:X2}{2:X2}", c.R, c.G, c.B)).Name.Remove(0, 2)
				Catch ex As Exception
				End Try
			End If



			Return configValue

		End Function


		<HttpPost()> _
		Function oleFind_Submit(filSelectFile As HttpPostedFileBase) As PartialViewResult
			' On Error Resume Next

			'Dim objOLE
			Dim filesize As Integer = 0
			Dim buffer As Byte()

			Dim sErrorMsg = ""
			' Read the information from the calling form.
			Dim sNextPage = Request.Form("txtGotoOptionPage")
			Dim sAction = Request.Form("txtGotoOptionAction")

			If CInt(Request.Form("txtOLEType")) < 2 And sAction = "" Then
				' We're just copying a file from client to server.
				' Read custom attributes
				Dim fileName As String = Request.Form("txtOLEJustFileName")
				Dim serverPath As String = Request.Form("txtOLEServerPath")

				If serverPath.Substring(serverPath.Length - 1) <> "\" And serverPath.Length > 0 Then
					serverPath &= "\"
				End If

				Try
					' Read input stream from request
					buffer = New Byte(filSelectFile.InputStream.Length - 1) {}
					Dim offset As Integer = 0
					Dim cnt As Integer = 0
					While (InlineAssignHelper(cnt, filSelectFile.InputStream.Read(buffer, offset, 10))) > 0
						offset += cnt
					End While

					IO.File.WriteAllBytes(serverPath + fileName, buffer)

				Catch generatedExceptionName As Exception
					Session("ErrorTitle") = "File upload"
					Session("ErrorText") = "You could not upload the file because of the following error:<p>" & FormatError(Err.Description)
					Dim data1 = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = ""}
					'Return Json(data1, JsonRequestBehavior.AllowGet)
				End Try

			Else
				' Moved to embedfile:
				' Commit changes to the database		
				If sAction = "LINKOLE" Then

					If Not filSelectFile Is Nothing Then
						filesize = filSelectFile.InputStream.Length
						buffer = New Byte(filSelectFile.InputStream.Length - 1) {}
						Dim offset As Integer = 0
						Dim cnt As Integer = 0

						While (InlineAssignHelper(cnt, filSelectFile.InputStream.Read(buffer, offset, 10))) > 0
							offset += cnt
						End While
					End If

					' The file will (should) have already been copied from the client to the temp path
					Dim objOLE = Session("OLEObject")
					With objOLE
						.UseEncryption = Request.Form("txtOLEEncryption")
						.OLEType = Request.Form("txtOLEType")
						.FileName = Request.Form("txtOLEFile")
						.DisplayFilename = Request.Form("txtOLEJustFileName")
						.OLEFileSize = filesize	' Request.Form("txtOLEFileSize")
						.OLEModifiedDate = Request.Form("txtOLEModifiedDate")
						Dim oleErrorResponse As String = .SaveStream(Session("optionRecordID"), Session("optionColumnID"), Session("realSource"), False, buffer)

						If oleErrorResponse.Length > 0 Then
							oleErrorResponse = Server.HtmlEncode("Unable to embed file:" & vbCrLf & oleErrorResponse)
						End If
						Session("errorMessage") = oleErrorResponse

						If .OLEType = 2 Then
							Session("optionFileValue") = .ExtractPhotoToBase64(Session("optionRecordID"), Session("optionColumnID"), Session("realSource"))
						Else
							Session("optionFileValue") = .FileName
						End If

						' .DeleteTempFile()
					End With
					Session("OLEObject") = objOLE
					objOLE = Nothing

					Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
					Session("timestamp") = objDatabase.GetRecordTimestamp(CleanNumeric(Session("optionRecordID")), Session("realSource"))

					'Update the ID badge picture in Session
					Session("SelfServicePhotograph_Src") = "data:image/jpeg;base64," & Session("optionFileValue")

				End If

				Session("optionScreenID") = Request.Form("txtGotoOptionScreenID")
				Session("optionTableID") = Request.Form("txtGotoOptionTableID")
				Session("optionViewID") = Request.Form("txtGotoOptionViewID")
				Session("optionOrderID") = Request.Form("txtGotoOptionOrderID")
				Session("optionRecordID") = Request.Form("txtGotoOptionRecordID")
				Session("optionFilterDef") = Request.Form("txtGotoOptionFilterDef")
				Session("optionFilterSQL") = Request.Form("txtGotoOptionFilterSQL")
				Session("optionValue") = Request.Form("txtGotoOptionValue")
				Session("optionLinkTableID") = Request.Form("txtGotoOptionLinkTableID")
				Session("optionLinkOrderID") = Request.Form("txtGotoOptionLinkOrderID")
				Session("optionLinkViewID") = Request.Form("txtGotoOptionLinkViewID")
				Session("optionLinkRecordID") = Request.Form("txtGotoOptionLinkRecordID")
				Session("optionColumnID") = Request.Form("txtGotoOptionColumnID")
				Session("optionLookupColumnID") = Request.Form("txtGotoOptionLookupColumnID")
				Session("optionLookupMandatory") = Request.Form("txtGotoOptionLookupMandatory")
				Session("optionLookupValue") = Request.Form("txtGotoOptionLookupValue")
				Session("optionFile") = Request.Form("txtGotoOptionFile")
				Session("optionExtension") = Request.Form("txtGotoOptionExtension")
				'Session("optionOLEOnServer") = Request.Form("txtGotoOptionOLEOnServer")
				Session("optionAction") = sAction
				Session("optionPageAction") = Request.Form("txtGotoOptionPageAction")
				Session("optionCourseTitle") = Request.Form("txtGotoOptionCourseTitle")
				Session("optionFirstRecPos") = Request.Form("txtGotoOptionFirstRecPos")
				Session("optionCurrentRecCount") = Request.Form("txtGotoOptionCurrentRecCount")
				Session("optionExprType") = Request.Form("txtGotoOptionExprType")
				Session("optionExprID") = Request.Form("txtGotoOptionExprID")
				Session("optionFunctionID") = Request.Form("txtGotoOptionFunctionID")
				Session("optionParameterIndex") = Request.Form("txtGotoOptionParameterIndex")
				Session("optionOLEType") = Request.Form("txtGotoOptionOLEType")
				Session("optionOLEMaxEmbedSize") = Request.Form("txtGotoOptionOLEMaxEmbedSize")

				If sAction = "" Then
					' Go to the requested page.
					'Return PartialView(sNextPage)	' Moved to oleFind.ascx, after .submit()
				End If

				If sAction = "CANCEL" Then

					' Clear up any temp files
					If Request.Form("txtOLEType") > 1 Then
						' No temp files, so skip this bit.
						' objOLE = Session("OLEObject")
						' objOLE.DeleteTempFile()
						' Session("OLEObject") = objOLE
						' objOLE = Nothing
					End If

					' Go to the requested page.
					Session("errorMessage") = sErrorMsg
					' Return PartialView(sNextPage)		' Moved to oleFind.ascx, after .submit()
				End If

				If sAction = "SELECTOLE" Then
					' Go to the requested page.
					'Return PartialView(sNextPage)		' Moved to oleFind.ascx, after .submit()
				End If

				' Commit changes to the database		
				If sAction = "LINKOLE" Then
					' Go to the requested page.
					'Return PartialView(sNextPage)		' Moved to oleFind.ascx, after .submit()
				End If
			End If

		End Function


		Public Function FolderList(folderPath As String) As ActionResult

			Dim directory As New DirectoryInfo(folderPath)

			Dim filelist As List(Of String) = (From filedetail In directory.GetFiles() Select filedetail.Name).ToList()
			Dim files() As String = filelist.ToArray()

			'Dim s As JavaScriptSerializer = New JavaScriptSerializer()
			'Dim result As String = s.Serialize(files)

			Return Json(files, JsonRequestBehavior.AllowGet)

		End Function

		<HttpPost> _
	 Public Function Upload(filSelectFile As HttpPostedFileBase) As ActionResult
			Const path As String = "D:\Temp\"

			If filSelectFile IsNot Nothing Then
				filSelectFile.SaveAs(path & Convert.ToString(filSelectFile.FileName))
			End If

			'Return RedirectToAction("Index")
		End Function

		'<HttpPost()> _
		'Public Function EmbedFile(filSelectFile As HttpPostedFileBase) As JsonResult

		'	' Commit changes to the database		
		'	' The file will (should) have already been copied from the client to the temp path

		'	Try
		'		' Read input stream from request
		'		Dim buffer As Byte() = New Byte(Request.InputStream.Length - 1) {}
		'		Dim offset As Integer = 0
		'		Dim cnt As Integer = 0

		'		While (InlineAssignHelper(cnt, Request.InputStream.Read(buffer, offset, 10))) > 0
		'			offset += cnt
		'		End While

		'		Dim objOLE = Session("OLEObject")
		'		With objOLE
		'			.UseEncryption = Request("HTTP_X_USEENCRYPTION")
		'			.OLEType = Request("HTTP_X_OLETYPE")
		'			.FileName = Request("HTTP_X_FILE_NAME")
		'			.DisplayFilename = Request("HTTP_X_DISPLAYFILENAME")
		'			.OLEFileSize = Request("HTTP_X_FILE_SIZE")
		'			.OLEModifiedDate = Request("HTTP_X_OLEMODIFIEDDATE")

		'			.SaveStream(Session("optionRecordID"), Session("optionColumnID"), Session("realSource"), False, buffer)
		'			'.DeleteTempFile()
		'		End With
		'		Session("OLEObject") = objOLE
		'		objOLE = Nothing

		'		Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
		'		Session("timestamp") = objDatabase.GetRecordTimestamp(CleanNumeric(Session("optionRecordID")), Session("realSource"))

		'	Catch generatedExceptionName As Exception
		'		Session("ErrorTitle") = "File upload"
		'		Session("ErrorText") = "You could not upload the file because of the following error:<p>" & FormatError(Err.Description)
		'		Dim data1 = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = ""}
		'		Return Json(data1, JsonRequestBehavior.AllowGet)
		'	End Try

		'End Function

		'<HttpPost()>
		'Public Function UploadFile() As JsonResult

		'	' Read custom attributes
		'	Dim fileName As String = Request("HTTP_X_FILE_NAME")
		'	Dim fileSize As String = Request("HTTP_X_FILE_SIZE")
		'	Dim serverPath As String = Request("HTTP_X_OLE_PATH")
		'	If serverPath.Substring(serverPath.Length - 1) <> "\" And serverPath.Length > 0 Then
		'		serverPath &= "\"
		'	End If

		'	Try
		'		' Read input stream from request
		'		Dim buffer As Byte() = New Byte(Request.InputStream.Length - 1) {}
		'		Dim offset As Integer = 0
		'		Dim cnt As Integer = 0
		'		While (InlineAssignHelper(cnt, Request.InputStream.Read(buffer, offset, 10))) > 0
		'			offset += cnt
		'		End While

		'		IO.File.WriteAllBytes(serverPath + fileName, buffer)

		'	Catch generatedExceptionName As Exception
		'		Session("ErrorTitle") = "File upload"
		'		Session("ErrorText") = "You could not upload the file because of the following error:<p>" & FormatError(Err.Description)
		'		Dim data1 = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = ""}
		'		Return Json(data1, JsonRequestBehavior.AllowGet)
		'	End Try

		'End Function

		Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
			target = value
			Return value
		End Function


		Public Function DownloadFile(filename As String, serverpath As String) As ActionResult

			If filename.Length > 0 And serverpath.Length > 0 Then

				If serverpath.Substring(serverpath.Length - 1) <> "\" Then serverpath &= "\"

				' TODO: add the file path!
				Response.ContentType = "application/octet-stream"
				Response.AppendHeader("Content-Disposition", "attachment; filename=" & filename)
				Dim fullpath = serverpath & filename
				Response.TransmitFile(fullpath)
				Response.End()
			End If

		End Function

		Public Function EditFile(plngRecordID As Integer, plngColumnID As Integer, pstrRealSource As String)

			Dim objOLE As Ole = Session("OLEObject")
			Dim fileResponse As Byte() = objOLE.CreateOLEDocument(plngRecordID, plngColumnID, pstrRealSource)

			Response.ContentType = "application/octet-stream"
			Response.AppendHeader("Content-Disposition", "attachment; filename=" & objOLE.DisplayFilename)

			Response.BinaryWrite(fileResponse)
			Response.Flush()
			Response.End()

		End Function

#Region "Standard Reports"

		Public Function stdrpt_AbsenceCalendar() As ActionResult
			Return PartialView()
		End Function

		Public Function stdrpt_AbsenceCalendar_details() As ActionResult
			Return View()
		End Function

		<HttpPost()>
		Function stdrpt_AbsenceCalendar_submit(value As FormCollection)

			Session("stdrpt_AbsenceCalendar_StartMonth") = Request.Form("txtStartMonth")
			Session("stdrpt_AbsenceCalendar_StartYear") = Request.Form("txtStartYear")
			Session("stdrpt_AbsenceCalendar_IncludeBankHolidays") = Request.Form("txtIncludeBankHolidays")
			Session("stdrpt_AbsenceCalendar_IncludeWorkingDaysOnly") = Request.Form("txtIncludeWorkingDaysOnly")
			Session("stdrpt_AbsenceCalendar_ShowBankHolidays") = Request.Form("txtShowBankHolidays")
			Session("stdrpt_AbsenceCalendar_ShowCaptions") = Request.Form("txtShowCaptions")
			Session("stdrpt_AbsenceCalendar_ShowWeekends") = Request.Form("txtShowWeekends")
			Return RedirectToAction("stdrpt_AbsenceCalendar")

		End Function

		Public Function stdrpt_def_absence() As ActionResult
			Return View()
		End Function

		<HttpPost()>
		Public Function stdrpt_run_AbsenceBreakdown() As ActionResult
			Return View()
		End Function

#End Region

		Public Function OrgChart() As PartialViewResult

			Dim m As OrgChart = New OrgChart()
			Dim model = m.LoadModel()

			Return PartialView(model)
		End Function

		<HttpPost()>
		Public Sub ResetSessionVars()
			Session("recordID") = ""
			Session("linkType") = ""
			Session("ViewDescription") = ""
		End Sub

	End Class

	Public Class ErrMsgJsonAjaxResponse

		Public Property ErrorTitle As String
		Public Property ErrorMessage As String
		Public Property Redirect As String
	End Class

	Public Class ViewDataUploadFilesResult
		Public Property Name As String
		Public Property Length As Integer
	End Class

End Namespace




