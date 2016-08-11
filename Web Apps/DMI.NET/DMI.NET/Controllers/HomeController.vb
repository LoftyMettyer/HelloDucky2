Option Explicit On
Option Strict Off

Imports System.Web.Mvc
Imports System.Web.UI.DataVisualization.Charting
Imports System.IO
Imports System.Web
Imports System.Drawing
Imports DMI.NET.Classes
Imports DMI.NET.Code
Imports DMI.NET.ViewModels.Home
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server
Imports DMI.NET.Models
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Net
Imports DMI.NET.ViewModels
Imports System.Collections.ObjectModel
Imports DMI.NET.Code.Hubs
Imports System.Web.Script.Serialization
Imports Aspose.Cells
Imports Newtonsoft.Json
Imports HR.Intranet.Server.Expressions
Imports HR.Intranet.Server.Extensions
Imports HR.Intranet.Server.ReportOutput
Imports DMI.NET.Models.ObjectRequests
Imports DMI.NET.Models.Responses
Imports HR.Intranet.Server.Structures
Imports DMI.NET.Code.Attributes
Imports DMI.NET.Infrastructure
Imports HR.Intranet.Server.Metadata

Namespace Controllers
	Public Class HomeController
		Inherits Controller

		Private _controllerRecord As New RecordController
		Private _controllerTraining As New TrainingController

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
		<ValidateAntiForgeryToken>
		Function Configuration_Submit(value As FormCollection)

			Dim sTemp
			Dim sType = ""
			Dim sControlName
			Dim fOK As Boolean = True

			If (Request.Form("txtPrimaryStartMode") <> "") Then

				' Save the user configuration settings.
				Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)

				' Some XSS/Injection defiant checks.
				Dim primary As String = NullSafeString(Request.Form("txtPrimaryStartMode"))
				If primary <> "1" And primary <> "2" And primary <> "3" Then fOK = False
				
				Dim history As String = NullSafeString(Request.Form("txtHistoryStartMode"))
				If history <> "1" And history <> "2" And history <> "3" Then fOK = False
				Dim lookup As String = NullSafeString(Request.Form("txtLookupStartMode"))
				If lookup <> "1" And lookup <> "2" And lookup <> "3" Then fOK = False
				Dim quickaccess As String = NullSafeString(Request.Form("txtQuickAccessStartMode"))
				If quickaccess <> "1" And quickaccess <> "2" And quickaccess <> "3" Then fOK = False

				Dim viewcolours As String = "2"  'Always colour in web 1=Mono.
				'If viewcolours <> "1" And viewcolours <> "2" Then fOK = False
				Dim nodesize As String = NullSafeString(Request.Form("txtExprNodeMode"))
				If nodesize <> "1" And nodesize <> "2" And nodesize <> "4" Then fOK = False
				Dim blocksize As String = NullSafeString(Request.Form("txtFindSize"))
				If Not IsNumeric(blocksize) Then fOK = False

				If fOK = False Then Throw New Exception("The user configuration settings could not be saved.")

				objDatabase.SaveUserSetting("RecordEditing", "Primary", primary)
				objDatabase.SaveUserSetting("RecordEditing", "History", history)
				objDatabase.SaveUserSetting("RecordEditing", "LookUp", lookup)
				objDatabase.SaveUserSetting("RecordEditing", "QuickAccess", quickaccess)
				objDatabase.SaveUserSetting("ExpressionBuilder", "ViewColours", viewcolours)
				objDatabase.SaveUserSetting("ExpressionBuilder", "NodeSize", nodesize)
				objDatabase.SaveUserSetting("IntranetFindWindow", "BlockSize", blocksize)

				Session("PrimaryStartMode") = primary
				Session("HistoryStartMode") = history
				Session("LookupStartMode") = lookup
				Session("QuickAccessStartMode") = quickaccess
				Session("ExprColourMode") = "2"
				Session("ExprNodeMode") = nodesize
				Session("FindRecords") = blocksize

            '--------------------------------------------
            ' Save the DefSel 'only mine' settings.
            '--------------------------------------------
            For i = 0 To 23
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
                  Case 21
                     sType = "NineBoxGrid"
                  Case 22
                     sType = "TalentReports"
                  Case 23
                     sType = "orgreporting"

               End Select

               sControlName = "txtOwner_" & sType
               sTemp = "onlymine " & sType

               Dim controlValue As String = Request.Form(sControlName)
               If controlValue <> "0" And controlValue <> "1" Then Throw New Exception("The user configuration settings could not be saved.")

               objDatabase.SaveUserSetting("defsel", sTemp, controlValue)

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

					Dim controlValue As String = Request.Form(sControlName)
					If controlValue <> "0" And controlValue <> "1" Then Throw New Exception("The user configuration settings could not be saved.")

					objDatabase.SaveUserSetting("warningmsg", sTemp, controlValue)

				Next

			End If

			Return RedirectToAction("CONFIGURATION")

		End Function

#End Region

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function newUser_Submit(value As NewUserModel) As JsonResult

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
			Dim objResponse As PostResponse

			' Create an OpenHR user associated with the given SQL Server login.
			Try
				objDataAccess.ExecuteSP("sp_ASRIntNewUser", _
						New SqlParameter("@psUserName", SqlDbType.VarChar, 128) With {.Value = value.Login})
				objResponse = New PostResponse With {.Action = "success", .Message = "User added successfully."}

			Catch ex As Exception
				objResponse = New PostResponse With {.Action = "fail", .Message = "You could not add the user because of the following error:" & vbNewLine & ex.Message.RemoveSensitive()}

			End Try

			Return Json(objResponse, JsonRequestBehavior.AllowGet)

		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function passwordChange_Submit(value As FormCollection) As JsonResult

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			Dim fSubmitPasswordChange = ""
			Dim sErrorText = ""
			Dim fRedirectToSSI As Boolean
      Dim verifyCode as String = Request.Form("txtVerify")

      If verifyCode Is Nothing OrElse Not verifyCode.Equals(Session("CaptchaText").ToString(), StringComparison.OrdinalIgnoreCase)      
				Session("ErrorTitle") = "Change Password Page"
				Session("ErrorText") = "Your password cannot be changed. Incorrect validation code."
				Dim data = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = "Main"}
				Return Json(data, JsonRequestBehavior.AllowGet)
      Else

				fSubmitPasswordChange = (Len(Request.Form("txtGotoPage")) = 0)

				If fSubmitPasswordChange Then
					' Force password change only if there are no other users logged in with the same name.
					Dim iUserSessionCount As Integer = GetCurrentUsersCountOnServer(Session("Username"))

					' variables to help select which main screen we return to after change or cancel
					fRedirectToSSI = CleanBoolean(Request.Form("txtRedirectToSSI"))
					Session("SSIMode") = fRedirectToSSI

					If iUserSessionCount < 2 Then
						' Read the Password details from the Password form.
						Dim sCurrentPassword = Request.Form("txtCurrentPassword")
						Dim sNewPassword As String = Request.Form("txtPassword1")

						Try
							objDataAccess.ChangePassword(New LoginInfo With {.Username = Session("Username"), .Password = sCurrentPassword}, sNewPassword)
							objDataAccess.Login.Password = sNewPassword

							' Tell the user that the password was changed okay.
							Session("ErrorTitle") = "Change Password Page"
							Session("ErrorText") = "Password changed successfully.<p>You will now be redirected to the login page."

							Dim data = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = "Login", .Route = "Account"}
							Return Json(data, JsonRequestBehavior.AllowGet)

						Catch ex As Exception
							Session("ErrorTitle") = "Change Password Page"
							Session("ErrorText") = "Your password cannot be changed. Please contact your administrator for assistance."
							Dim data = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = "Main"}
							Return Json(data, JsonRequestBehavior.AllowGet)

						End Try

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
					Session("lineage") = Request.Form("txtGotoLineage")
					Session("utilID") = Request.Form("txtGotoUtilID")
					Session("locateValue") = Request.Form("txtGotoLocateValue")
					Session("firstRecPos") = Request.Form("txtGotoFirstRecPos")
					Session("currentRecCount") = Request.Form("txtGotoCurrentRecCount")

					' Go to the requested page.
					' Return RedirectToAction(Request.Form("txtGotoPage"))
					Session("txtGotoPage") = Request.Form("txtGotoPage")
				End If
			End If
		End Function

		Function MainSSI() As ActionResult
			Session("SSIMode") = True
			Return RedirectToAction("Main")
		End Function

		Function MainDMI() As ActionResult
			Session("SSIMode") = False
			Return RedirectToAction("Main")
		End Function

		' GET: /Home
		<HttpGet>
		Function Main() As ActionResult

			Dim objSessionInfo = CType(Session("SessionContext"), SessionInfo)
			Dim bOK As Boolean = True
			Dim targetWebArea As WebArea = WebArea.SSI

			If objSessionInfo Is Nothing Then
				Return RedirectToAction("login", "Account")
			End If

			ResetSessionVars()

			Session("utilid") = ""
			Session("selectSQL") = ""

			' Reset value of IsLoadedFromReportDefinition flag. 
			' Which used to loads the tools screen to the ToolsFrame from report definition
			Session("IsLoadedFromReportDefinition") = False

			If Session("SSIMode") <> True Then Session("SSIMode") = False ' set default value

			Session("ErrorText") = ""
			Session("WarningText") = ""

			If Session("SSIMode") = True AndAlso Not objSessionInfo.LoginInfo.IsSSIUser Then
				Session("ErrorText") = "Login Failed."
				bOK = False
			End If

			If Session("SSIMode") = False AndAlso Not objSessionInfo.LoginInfo.IsDMIUser Then
				Session("ErrorText") = "Login Failed."
				bOK = False
			End If

			If Not Session("SSIMode") Then
				targetWebArea = WebArea.DMI
			End If

			If bOK Then

                ' Check the overnight
                Dim objDatabase As Database = Session("DatabaseFunctions")

                If targetWebArea = WebArea.DMI AndAlso Not objDatabase.IsOvernightJobOk() Then
                    Session("WarningText") = "The overnight job has not completed within the last 24 hours.<br/>This may affect date dependant areas of OpenHR." _
                        & "<br/><br/>" & "Please inform your SQL administrator."
                End If

                ' Licence check
                Dim objCurrentLogin = CType(Session("sessionCurrentUser"), LoginViewModel)
				Dim licenceValidate = LicenceHub.NavigateWebArea(objCurrentLogin, targetWebArea)

				Select Case licenceValidate
					Case LicenceValidation.Failure
						Session("ErrorText") = LicenceHub.ErrorMessage(licenceValidate)
						bOK = False

					Case LicenceValidation.Expired, LicenceValidation.Insufficient
						Session("ErrorText") = LicenceHub.ErrorMessage(licenceValidate)
						bOK = False

					Case LicenceValidation.HeadcountWarning
						If LicenceHub.DisplayWarningToUser(Session("Username").ToString(), Enums.WarningType.Headcount95Percent, 7) Then
							Session("WarningText") = LicenceHub.ErrorMessage(licenceValidate)
						End If
						bOK = True

					Case LicenceValidation.ExpiryWarning, LicenceValidation.HeadcountAndExpiryWarning
						If LicenceHub.DisplayWarningToUser(Session("Username").ToString(), Enums.WarningType.Licence5DayExpiry, 1) Then
							Session("WarningText") = LicenceHub.ErrorMessage(licenceValidate)
						End If
						bOK = True

					Case LicenceValidation.HeadcountExceeded, LicenceValidation.HeadcountExceededAndExpiryWarning
						Session("WarningText") = LicenceHub.ErrorMessage(licenceValidate)
						bOK = True

				End Select

				ViewData("showOutOfOffice") = ShowOutOfOffice(NullSafeInteger(Session("SingleRecordTableID")), NullSafeInteger(Session("SingleRecordViewID")))
			End If

			Session("ViewDescription") = Session("welcomemessage")

			If bOK Then
				Return View()
			Else
				Session("MessageTitle") = "" 'You could put a title here in the future if it is decided.
				Session("MessageText") = Session("ErrorText")
				Return RedirectToAction("LoginMessage", "Account")
			End If

		End Function

		Function GetFindRecordByID(RecordID As Integer) As String
			Dim objSession As SessionInfo = CType(Session("SessionContext"), SessionInfo)	 'Set session info
			Dim objDataAccess As New clsDataAccess(objSession.LoginInfo) 'Instantiate DataAccess class
			Dim SPParameters() As SqlParameter
			Dim resultDataSet As DataSet

			Dim prmError As New SqlParameter("@pfError", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmSomeSelectable As New SqlParameter("@pfSomeSelectable", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmSomeNotSelectable As New SqlParameter("@pfSomeNotSelectable", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmRealSource As New SqlParameter("@psRealSource", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prmInsertGranted As New SqlParameter("@pfInsertGranted", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmDeleteGranted As New SqlParameter("@pfDeleteGranted", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmIsFirstPage As New SqlParameter("@pfFirstPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmIsLastPage As New SqlParameter("@pfLastPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmColumnType As New SqlParameter("@piColumnType", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmColumnSize As New SqlParameter("@piColumnSize", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmColumnDecimals As New SqlParameter("@piColumnDecimals", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmTotalRecCount As New SqlParameter("@piTotalRecCount", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmFirstRecPos As New SqlParameter("@piFirstRecPos", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = CleanNumeric(Session("firstRecPos"))}
			Dim prmIsValidFilterColumns As New SqlParameter("@bIsValidFilter", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}

			Dim filterDefForCurrentTable As String = IIf(IsNothing(Session("filterDef_" & Session("tableID"))), "", Session("filterDef_" & Session("tableID")))

			SPParameters = New SqlParameter() { _
					prmError, _
					prmSomeSelectable, _
					prmSomeNotSelectable, _
					prmRealSource, _
					prmInsertGranted, _
					prmDeleteGranted, _
					New SqlParameter("@piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("tableID"))}, _
					New SqlParameter("@piViewID", SqlDbType.Int) With {.Value = CleanNumeric(Session("viewID"))}, _
					New SqlParameter("@piOrderID ", SqlDbType.Int) With {.Value = CleanNumeric(Session("orderID"))}, _
					New SqlParameter("@piParentTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("parentTableID"))}, _
					New SqlParameter("@piParentRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("parentRecordID"))}, _
					New SqlParameter("@psFilterDef", SqlDbType.VarChar, -1) With {.Value = filterDefForCurrentTable}, _
					New SqlParameter("@piRecordsRequired", SqlDbType.Int) With {.Value = 10000000}, _
					prmIsFirstPage, _
					prmIsLastPage, _
					New SqlParameter("@psLocateValue", SqlDbType.VarChar, -1) With {.Value = Session("locateValue")}, _
					prmColumnType, _
					prmColumnSize, _
					prmColumnDecimals, _
					New SqlParameter("@psAction", SqlDbType.VarChar) With {.Value = Session("action"), .Size = 255}, _
					prmTotalRecCount, _
					prmFirstRecPos, _
					New SqlParameter("@piCurrentRecCount", SqlDbType.Int) With {.Value = CleanNumeric(Session("currentRecCount"))}, _
					New SqlParameter("@psDecimalSeparator", SqlDbType.VarChar, 255) With {.Value = Session("LocaleDecimalSeparator")}, _
					New SqlParameter("@psLocaleDateFormat", SqlDbType.VarChar, 255) With {.Value = Platform.LocaleDateFormatForSQL()}, _
					New SqlParameter("@RecordID", SqlDbType.Int) With {.Value = RecordID}, _
					prmIsValidFilterColumns
					}
			Try
				resultDataSet = objDataAccess.GetDataSet("spASRIntGetFindRecords", SPParameters)

				'If no data is returned then that means that the row is no longer part of the table/view
				If resultDataSet.Tables(1).Rows.Count = 0 Then
					Return JsonConvert.SerializeObject("")
				End If

				Return JsonConvert.SerializeObject(resultDataSet.Tables(1))
			Catch ex As Exception
				Throw New Exception("The find records could not be retrieved." & vbCrLf & FormatError(ex.Message))
			End Try
		End Function


		Function GetSummaryColumns(parentTableID As String, parentRecordID As String) As String

			Dim SPParameters() As SqlParameter
			Dim resultsDataTable As New DataTable
			parentTableID = CleanNumeric(parentTableID)
			parentRecordID = CleanNumeric(parentRecordID)

			Dim objSession As SessionInfo = CType(Session("SessionContext"), SessionInfo)	'Set session info
			Dim objDataAccess As New clsDataAccess(objSession.LoginInfo) 'Instantiate DataAccess class
			SPParameters = New SqlParameter() { _
				New SqlParameter("@piHistoryTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("tableID"))}, _
				New SqlParameter("@piParentTableID", SqlDbType.Int) With {.Value = parentTableID}, _
				New SqlParameter("@piParentRecordID", SqlDbType.Int) With {.Value = parentRecordID} _
			}

			Try
				resultsDataTable = objDataAccess.GetDataTable("spASRIntGetSummaryValues", CommandType.StoredProcedure, SPParameters)
			Catch
			End Try

			If resultsDataTable Is Nothing OrElse resultsDataTable.Rows.Count = 0 Then
				Return JsonConvert.SerializeObject("")
			End If

			'Convert the integers to strings otherwise we lose the precision when reading in JS.
			Dim resultsAsString As New Dictionary(Of String, String)
			For Each col As DataColumn In resultsDataTable.Columns
				resultsAsString.Add(col.ColumnName, resultsDataTable.Rows(0).Item(col.ColumnName).ToString())
			Next

			Return JsonConvert.SerializeObject(resultsAsString)

		End Function


		<AllowJwtAccess()>
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

			ViewData("showOutOfOffice") = ShowOutOfOffice(NullSafeInteger(Session("tableID")), NullSafeInteger(Session("viewID")))

			Return View()

		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function WorkAreaRefresh()
			Dim previousTableID As String = Session("tableID")
			If ValidateLineageValue(Request.Form("txtGotoLineage")) = "-1" Then
				' We're flipping between histories, reuse session variables where possible.
				Session("tableID") = Request.Form("txtGotoTableID")
				Session("viewID") = Request.Form("txtGotoViewID")
				Session("screenID") = Request.Form("txtGotoScreenID")
				Session("orderID") = Request.Form("txtGotoOrderID")
				Session("recordID") = Request.Form("txtGotoRecordID")
				Session("currentPageID") = Request.Form("txtGotoPageID")
				Session("CurrentPageSortColumn") = Request.Form("txtGotoPageSortColumn")
				Session("CurrentPageSortOrder") = Request.Form("txtGotoPageSortOrder")
				Session("firstRecPos") = Request.Form("txtGotoFirstRecPos")
				Session("currentRecCount") = Request.Form("txtGotoCurrentRecCount")
				Session("parentTableID") = Request.Form("txtGotoParentTableID")
				Session("parentRecordID") = Request.Form("txtGotoParentRecordID")
			Else
				' Save the required table/view and screen IDs in session variables.
				Session("action") = ValidateFromWhiteList(Request.Form("txtAction"), InputValidation.WhiteListCollections.Actions)
				Session("tableID") = ValidateIntegerValue(Request.Form("txtGotoTableID"))
				Session("viewID") = ValidateIntegerValue(Request.Form("txtGotoViewID"))
				Session("screenID") = ValidateIntegerValue(Request.Form("txtGotoScreenID"))
				Session("orderID") = ValidateIntegerValue(Request.Form("txtGotoOrderID"))
				Session("recordID") = ValidateIntegerValue(Request.Form("txtGotoRecordID"))
				Session("currentPageID") = ValidateIntegerValue(Request.Form("txtGotoPageID"))
				Session("CurrentPageSortColumn") = Request.Form("txtGotoPageSortColumn")
				Session("CurrentPageSortOrder") = Request.Form("txtGotoPageSortOrder")
				Session("parentTableID") = ValidateIntegerValue(Request.Form("txtGotoParentTableID"))
				Session("parentRecordID") = ValidateIntegerValue(Request.Form("txtGotoParentRecordID"))
				Session("realSource") = Request.Form("txtGotoRealSource")

				Session("lineage") = ValidateLineageValue(Request.Form("txtGotoLineage"))
				Session("utilID") = ValidateIntegerValue(Request.Form("txtGotoUtilID"))
				Session("locateValue") = ValidateIntegerValue(Request.Form("txtGotoLocateValue"))
				Session("firstRecPos") = ValidateIntegerValue(Request.Form("txtGotoFirstRecPos"))
				Session("currentRecCount") = ValidateIntegerValue(Request.Form("txtGotoCurrentRecCount"))
			End If

			Session("optionRecordID") = 0
			Session("optionAction") = OptionActionType.Empty

			Session("OptionSelectedRecordIds") = Request.Form("txtSelectedRecordsInFindGrid")

			Dim currentTableID As String = Session("tableID")

			If Not String.IsNullOrEmpty(previousTableID) AndAlso Not String.IsNullOrEmpty(currentTableID) AndAlso currentTableID <> "0" Then 'If we have a tableID in session
				Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
				Dim currentTable As Metadata.Table = objDatabase.GetTableByID(Integer.Parse(Session("tableID")))
				If currentTable.TableType = TableTypes.tabTopLevel And Request.Form("txtGotoPage") <> "recordEdit" Then
					For Each relation As Metadata.Relation In objDatabase.GetRelationsByParentTableID(Integer.Parse(Session("tableID")))
						Session("filterDef_" & relation.ChildID) = ""	'Clear the child table filter
						Session("filterSQL_" & relation.ChildID) = ""	'Clear the child table SQL filter
					Next
				End If
			End If

			' Go to the requested page.
			Return RedirectToAction(Request.Form("txtGotoPage"))

		End Function

#Region "Split Out EmptyOption"

		Private Sub emptyoption_Submit_BASE(form As GotoOptionDataModel)

			' Save the required information in session variables.
			Session("optionScreenID") = form.txtGotoOptionScreenID
			Session("optionTableID") = form.txtGotoOptionTableID
			Session("optionViewID") = form.txtGotoOptionViewID
			Session("optionOrderID") = form.txtGotoOptionOrderID
			Session("optionRecordID") = form.txtGotoOptionRecordID
			Session("optionFilterDef") = form.txtGotoOptionFilterDef
			Session("optionFilterSQL") = form.txtGotoOptionFilterSQL
			Session("optionValue") = form.txtGotoOptionValue
			Session("optionLinkTableID") = form.txtGotoOptionLinkTableID
			Session("optionLinkOrderID") = form.txtGotoOptionLinkOrderID
			Session("optionLinkViewID") = form.txtGotoOptionLinkViewID
			Session("optionLinkRecordID") = form.txtGotoOptionLinkRecordID
			Session("optionColumnID") = form.txtGotoOptionColumnID
			Session("optionLookupColumnID") = form.txtGotoOptionLookupColumnID
			Session("optionLookupMandatory") = form.txtGotoOptionLookupMandatory
			Session("optionLookupValue") = form.txtGotoOptionLookupValue
			Session("optionLookupFilterValue") = form.txtGotoOptionLookupFilterValue
			Session("optionFile") = form.txtGotoOptionFile
			Session("optionAction") = form.txtGotoOptionAction
			Session("optionPageAction") = form.txtGotoOptionPageAction
			Session("optionCourseTitle") = form.txtGotoOptionCourseTitle
			Session("optionFirstRecPos") = form.txtGotoOptionFirstRecPos
			Session("optionCurrentRecCount") = form.txtGotoOptionCurrentRecCount
			Session("OptionRealsource") = form.txtGotoOptionRealsource

			Session("OptionSelectedRecordIds") = form.txtGotoLocateSelectedRecordsInFindGrid

		End Sub

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function tbTransferCourseFind(form As GotoOptionDataModel) As RedirectToRouteResult
			emptyoption_Submit_BASE(form)
			Return RedirectToAction("OptionDataGrid", "Home", New With {.GotoOptionPage = "tbTransferCourseFind"})
		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function tbBookCourseFind(form As GotoOptionDataModel) As RedirectToRouteResult
			emptyoption_Submit_BASE(form)
			Return RedirectToAction("OptionDataGrid", "Home", New With {.GotoOptionPage = "tbBookCourseFind"})
		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function tbTransferBookingFind(form As GotoOptionDataModel) As RedirectToRouteResult
			emptyoption_Submit_BASE(form)
			Return RedirectToAction("OptionDataGrid", "Home", New With {.GotoOptionPage = "tbTransferBookingFind"})
		End Function

		Function BulkBooking(form As GotoOptionDataModel) As ActionResult
			emptyoption_Submit_BASE(form)
			Dim m As New BulkBookingViewModel()
			Return PartialView("BulkBooking", m)
		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function tbAddFromWaitingListFind(form As GotoOptionDataModel) As RedirectToRouteResult
			emptyoption_Submit_BASE(form)
			Return RedirectToAction("OptionDataGrid", "Home", New With {.GotoOptionPage = "tbAddFromWaitingListFind"})
		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function menu_loadLookupPage(form As GotoOptionDataModel) As RedirectToRouteResult
			emptyoption_Submit_BASE(form)
			Return RedirectToAction("lookupFind")
		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function menu_loadLinkPage(form As GotoOptionDataModel) As RedirectToRouteResult
			emptyoption_Submit_BASE(form)
			Return RedirectToAction("linkfind")
		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function menu_oleFind(form As GotoOptionDataModel) As RedirectToRouteResult
			emptyoption_Submit_BASE(form)

			Session("optionOLEType") = Request.Form("txtGotoOptionOLEType")
			Session("optionOLEMaxEmbedSize") = Request.Form("txtGotoOptionOLEMaxEmbedSize")
			Session("optionOLEReadOnly") = Request.Form("txtGotoOptionOLEReadOnly")
			Session("optionIsPhoto") = Request.Form("txtGotoOptionIsPhoto")

			Return RedirectToAction("olefind")
		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function menu_loadQuickFindNoSaveCheck(form As GotoOptionDataModel) As RedirectToRouteResult
			emptyoption_Submit_BASE(form)
			Return RedirectToAction("Quickfind")
		End Function

		<HttpPost()>
	<ValidateAntiForgeryToken>
		Function orderfilter_RecordEdit(form As GotoOptionDataModel) As RedirectToRouteResult
			emptyoption_Submit_BASE(form)
			Return RedirectToAction(Request.Form("txtGotoOptionPage"))
		End Function

		<HttpPost()>
	<ValidateAntiForgeryToken>
		Function menu_loadSelectOrderFilter(form As GotoOptionDataModel) As RedirectToRouteResult
			emptyoption_Submit_BASE(form)
			Return RedirectToAction(Request.Form("txtGotoOptionPage"))
		End Function

		<HttpPost()>
	<ValidateAntiForgeryToken>
		Function menu_LoadAbsenceCalendar(form As GotoOptionDataModel) As RedirectToRouteResult
			emptyoption_Submit_BASE(form)
			Return RedirectToAction("stdrpt_AbsenceCalendar")
		End Function

		<HttpPost()>
<ValidateAntiForgeryToken>
		Function menu_LoadAbsenceCalendarNoSaveCheck(form As GotoOptionDataModel) As RedirectToRouteResult
			emptyoption_Submit_BASE(form)
			Return RedirectToAction("stdrpt_AbsenceCalendar")
		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function expression_addClick(value As ExpressionComponentModel) As ActionResult
			Return View("util_def_exprComponent", value)
		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function expression_insertClick(value As ExpressionComponentModel) As ActionResult
			Return View("util_def_exprComponent", value)
		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function expression_editClick(value As ExpressionComponentModel) As ActionResult
			Return View("util_def_exprComponent", value)
		End Function

		<HttpPost()>
<ValidateAntiForgeryToken>
		Function data_window_onload(form As GotoOptionDataModel) As RedirectToRouteResult
			emptyoption_Submit_BASE(form)
			Return RedirectToAction("filterselect")
		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function LoadStandardReport(postData As StandardReportModel) As ActionResult

			Dim objSession = CType(Session("SessionContext"), SessionInfo)
			Dim bPermitted As Boolean
			Dim sUtilName As String

			Select Case postData.utiltype
				Case UtilityType.utlAbsenceBreakdown
					sUtilName = "Absence Breakdown"
					bPermitted = objSession.IsPermissionGranted("STANDARDREPORTS", "RUN_AB")
				Case Else
					sUtilName = "Bradford Factor"
					bPermitted = objSession.IsPermissionGranted("STANDARDREPORTS", "RUN_BF")
			End Select

			' Validate permission (should only be hit if user "hacked" the accordian)
			If Not bPermitted Then
				Return RedirectToAction("PermissionsError", "Error")
			Else

				Session("optionAction") = postData.Action
				Session("optionRecordID") = postData.EmployeeID
				Session("singleRecordID") = postData.EmployeeID
				Session("utiltype") = postData.utiltype
				Session("utilname") = sUtilName

				Return View("stdrpt_def_absence")
			End If

		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function linkFind_submit(postData As LinkFindModel) As RedirectToRouteResult

			' Save the required information in session variables.
			Session("optionAction") = postData.Action
			Session("optionScreenID") = postData.ScreenID
			Session("optionLinkTableID") = postData.LinkTableID
			Session("optionRecordID") = postData.LinkRecordID

			Return RedirectToAction("emptyoption")
		End Function

#End Region

		Function DefSel() As ActionResult
			Return View()
		End Function

		<ValidateAntiForgeryToken()>
		<HttpPost()>
		Function RefreshSession() As Boolean
			' This function simply resets the session timeout
			Return True
		End Function


		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function DefSel(value As DefSelModel) As ActionResult

			Dim objSession = CType(Session("SessionContext"), SessionInfo)

         ' Validate permission (should only be hit if user "hacked" the accordian)
         If Not objSession.IsCategoryGranted(value.utiltype) Then
            Return RedirectToAction("PermissionsError", "Error")
         End If

         Session("defseltype") = value.utiltype
         Session("utilTableID") = IIf(value.txtTableID = -1, SettingsConfig.Personnel_EmpTableID, value.txtTableID)
         Session("fromMenu") = IIf(value.txtGotoFromMenu, "1", "0") ' No idea what this is doing, just placed for backward compatability. Candidate for removal!
         Session("singleRecordID") = value.RecordID
         Session("multipleRecordIDs") = value.MultipleRecordIDs

         If (value.ResetCategoryAndOwner) Then
            Session("defsel_categoryId_" & value.utiltype) = -1
            Session("defsel_ownerName_" & value.utiltype) = ""
         End If

         If (value.txtGotoFromMenu) Or ((Not Session("multipleRecordIDs") Is Nothing AndAlso Session("multipleRecordIDs").ToString().Length > 0) Or (Not Session("singleRecordID") Is Nothing AndAlso Session("singleRecordID") > 0)) Then
            Session("OnlyMine") = CBool(objSession.GetUserSetting("defsel", "onlymine " + value.utiltype.ToSecurityPrefix, False))
         Else
            Session("OnlyMine") = value.OnlyMine
         End If

         Return View()

      End Function

      <HttpPost()>
      <ValidateAntiForgeryToken>
      Function DefSel_Submit(value As DefSelModel)

         Try

            Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
            Dim objSession = CType(Session("SessionContext"), SessionInfo)

            Dim sRequiredPermission = value.Action

            If value.Action = "copy" Then sRequiredPermission = "NEW"

            ' Validate permission (should only be hit if user "hacked" the button properties)
            If Not objSession.IsPermissionGranted(value.utiltype.ToSecurityPrefix, sRequiredPermission) Then
               Return RedirectToAction("PermissionsError", "Error")
            End If

            ' Set some session variables used by all the util pages
            Session("utiltype") = value.utiltype
            Session("utilid") = value.utilID
            Session("utilname") = value.utilName
            Session("action") = value.Action
            Session("utilTableID") = value.txtTableID

            ' Reset ownername and categoryid sessions
            Session("defsel_categoryId_" & value.utiltype) = value.CategoryId
            Session("defsel_ownerName_" & value.utiltype) = value.OwnerName

            ' Now examine what we are doing and redirect as appropriate
            If (Session("action") = "new") Or
               (Session("action") = "edit") Or
               (Session("action") = "view") Or
               (Session("action") = "copy") Then
               Select Case Session("utiltype")
                  Case UtilityType.utlCrossTab
                     Return RedirectToAction("util_def_crosstab", "reports")
                  Case UtilityType.utlCustomReport
                     Return RedirectToAction("util_def_customreport", "reports")
                  Case UtilityType.utlMailMerge
                     Return RedirectToAction("util_def_mailmerge", "reports")
                  Case UtilityType.utlPicklist
                     Return RedirectToAction("util_def_picklist")
                  Case UtilityType.utlFilter
                     Return PartialView("util_def_expression")
                  Case UtilityType.utlCalculation
                     Return PartialView("util_def_expression")
                  Case UtilityType.utlCalendarReport
                     Return RedirectToAction("util_def_calendarreport", "reports")
                  Case UtilityType.utlNineBoxGrid
                     Return RedirectToAction("util_def_9boxgrid", "reports")
                  Case UtilityType.TalentReport
                     Return RedirectToAction("util_def_talentreport", "reports")
                  Case UtilityType.OrgReporting
                     Return RedirectToAction("util_def_organisationreport", "reports")
               End Select

            ElseIf Session("action") = "delete" Then

               Dim rstUsage = objDatabase.GetUtilityUsage(value.utiltype, value.utilID)

					If rstUsage.Rows.Count = 0 Then
						Return RedirectToAction("util_delete", value)
					Else

						value.Usage = New Collection(Of DefinitionPropertiesViewModel)
						For Each objRow As DataRow In rstUsage.Rows
							Dim objUsage As New DefinitionPropertiesViewModel With {.Name = objRow("description").ToString}
							value.Usage.Add(objUsage)
						Next

						Return View("checkforusage", value)
					End If

				End If

			Catch ex As Exception
				Throw

			End Try

		End Function

		<ValidateAntiForgeryToken()>
		<HttpPost>
		Function DefinitionProperties(objModel As DefinitionPropertiesViewModel, type As UtilityType) As ActionResult
			Dim objDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			Dim dsDefProp = objDataAccess.GetDataSet("spASRIntDefProperties" _
				, New SqlParameter("intType", SqlDbType.Int) With {.Value = CInt(type)} _
				, New SqlParameter("intID", SqlDbType.Int) With {.Value = objModel.ID})

			objModel.Type = type

			If dsDefProp.Tables(0).Rows.Count > 0 Then
				Dim rowAccess = dsDefProp.Tables(0).Rows(0)
				objModel.CreatedDate = rowAccess("CreatedDate").ToString() & " by " & rowAccess("Createdby").ToString
				If objModel.CreatedDate = " by " Then objModel.CreatedDate = "<Unknown>"
				objModel.LastSaveDate = rowAccess("SavedDate").ToString() & " by " & rowAccess("Savedby").ToString()
				If objModel.LastSaveDate = " by " Then objModel.LastSaveDate = "<Unknown>"
				objModel.LastRunDate = rowAccess("RunDate").ToString() & " by " & rowAccess("Runby").ToString()
				If objModel.LastRunDate = " by " Then objModel.LastRunDate = "<Unknown>"
			End If

			objModel.Usage = New Collection(Of DefinitionPropertiesViewModel)
			For Each objRow As DataRow In dsDefProp.Tables(1).Rows
				Dim objUsage As New DefinitionPropertiesViewModel With {.Name = objRow("description").ToString()}
				objModel.Usage.Add(objUsage)
			Next

			Return PartialView(objModel)

		End Function

		Function util_delete(value As DefSelModel) As ActionResult

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			Dim sCheckStatus As String = ""

			Dim sUtilTypeName As String = ""

			Select Case value.utiltype

				Case UtilityType.utlCrossTab
					sUtilTypeName = "cross tab"

				Case UtilityType.utlCustomReport
					sUtilTypeName = "report"

				Case UtilityType.utlMailMerge
					sUtilTypeName = "mail merge"

				Case UtilityType.utlPicklist
					sUtilTypeName = "picklist"

				Case UtilityType.utlFilter
					sUtilTypeName = "filter"

				Case UtilityType.utlCalculation
					sUtilTypeName = "calculation"

				Case UtilityType.utlCalendarReport
					sUtilTypeName = "calendar report"

				Case UtilityType.utlNineBoxGrid
					sUtilTypeName = "9-box grid report"

 				Case UtilityType.TalentReport
               sUtilTypeName = "talent report"

            Case UtilityType.OrgReporting
               sUtilTypeName = "organisation report"

         End Select

			Try

				Dim prmDeleted = New SqlParameter("pfDeleted", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmAccess = New SqlParameter("psAccess", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

				objDataAccess.ExecuteSP("spASRIntDeleteCheck" _
							, New SqlParameter("piUtilityType", SqlDbType.Int) With {.Value = value.utiltype} _
							, New SqlParameter("plngID", SqlDbType.Int) With {.Value = value.utilID} _
							, prmDeleted _
							, prmAccess)

				If CBool(prmDeleted.Value) = True Then
					sCheckStatus = value.utilName & " " & sUtilTypeName & " has been deleted by another user."
				ElseIf prmAccess.Value.ToString() = "HD" Then
					sCheckStatus = value.utilName & " " & sUtilTypeName & " has been made hidden by another user."
				ElseIf prmAccess.Value.ToString() = "RO" Then
					sCheckStatus = value.utilName & " " & sUtilTypeName & " has been made read only by another user."
				End If

				If Len(sCheckStatus) > 0 Then
					value.Status = sCheckStatus

				Else

					objDataAccess.ExecuteSP("sp_ASRIntDeleteUtility" _
											, New SqlParameter("piUtilType", SqlDbType.Int) With {.Value = value.utiltype} _
											, New SqlParameter("piUtilID", SqlDbType.Int) With {.Value = value.utilID})

					value.Status = "'" & value.utilName & "' " & sUtilTypeName & " has been deleted."
					Return View("CheckForUsage", value)

				End If

			Catch ex As Exception
				Throw

			End Try

		End Function

		Function Data() As ActionResult
			Dim m As New DataViewModel
			Return View(m)
		End Function

		Function OptionData() As ActionResult
			Return View()
		End Function

		<HttpPost>
		<ValidateAntiForgeryToken>
		Function optionData_Submit(form As OptionDataModel) As ActionResult

			' Read the information from the calling form.
			Session("optionAction") = form.txtOptionAction
			Session("optionTableID") = form.txtoptionTableID
			Session("optionViewID") = form.txtOptionViewID
			Session("optionOrderID") = form.txtOptionOrderID
			Session("optionColumnID") = form.txtOptionColumnID
			Session("optionPageAction") = form.txtOptionPageAction
			Session("optionFirstRecPos") = form.txtOptionFirstRecPos
			Session("optionCurrentRecCount") = form.txtOptionCurrentRecCount
			Session("optionLocateValue") = form.txtGotoLocateValue
			Session("optionCourseTitle") = form.txtOptionCourseTitle
			Session("optionRecordID") = form.txtOptionRecordID
			Session("optionLinkRecordID") = form.txtOptionLinkRecordID
			Session("optionValue") = form.txtOptionValue
			Session("optionSQL") = form.txtOptionSQL
			Session("optionPromptSQL") = form.txtOptionPromptSQL
			Session("optionOnlyNumerics") = form.txtOptionOnlyNumerics
			Session("optionLookupColumnID") = form.txtOptionLookupColumnID
			Session("optionFilterValue") = form.txtOptionLookupFilterValue
			Session("IsLookupTable") = form.txtOptionIsLookupTable
			Session("optionParentTableID") = form.txtOptionParentTableID
			Session("optionParentRecordID") = form.txtOptionParentRecordID
			Session("option1000SepCols") = form.txtOption1000SepCols

			' Go to the requested page.
			Return RedirectToAction("OptionData")

		End Function

		<HttpPost>
		<ValidateAntiForgeryToken>
		Function Data_Submit(dataViewModel As DataViewModel) As ActionResult
			Dim sErrorMsg As String = ""
			Dim fWarning As Boolean = False
			Dim fOk As Boolean = False
			Dim fTBOverride As Boolean = False
			Dim sTBResultCode As String = "000"	'Validation OK
			Dim sCourseOverbooked As String = ""


			' Read the information from the calling form.
			Dim sRealSource = dataViewModel.txtRealSource
			Dim lngTableID = ValidateIntegerValue(dataViewModel.txtCurrentTableID)
			Dim lngScreenID = ValidateIntegerValue(dataViewModel.txtCurrentScreenID)
			Dim lngViewID = ValidateIntegerValue(dataViewModel.txtCurrentViewID)
			Dim lngRecordID = ValidateIntegerValue(dataViewModel.txtRecordID)
			Dim sAction = ValidateFromWhiteList(dataViewModel.txtAction, InputValidation.WhiteListCollections.Actions)
			Dim sReaction = ValidateFromWhiteList(dataViewModel.txtReaction, InputValidation.WhiteListCollections.Actions)
			Dim sInsertUpdateDef = dataViewModel.txtInsertUpdateDef
			Dim iTimestamp = ValidateIntegerValue(dataViewModel.txtTimestamp)
			Dim iTBEmployeeRecordID = ValidateIntegerValue(dataViewModel.txtTBEmployeeRecordID)
			Dim iTBCourseRecordID = ValidateIntegerValue(dataViewModel.txtTBCourseRecordID)
			Dim sTBBookingStatusValue = dataViewModel.txtTBBookingStatusValue
			Dim fUserChoice = dataViewModel.txtUserChoice

			Session("currentPageID") = dataViewModel.txtGotoCurrentPageID

			If dataViewModel.txtTBOverride = "" Then
				fTBOverride = False
			Else
				fTBOverride = ValidateBooleanValue(dataViewModel.txtTBOverride)
			End If

			If sAction = "SAVE" Then
				Dim lngOriginalRecordID = ValidateIntegerValue(dataViewModel.txtOriginalRecordID)
				Dim result = _controllerRecord.data_submit_SAVE(lngTableID, lngRecordID, sReaction, fTBOverride, iTBEmployeeRecordID, iTBCourseRecordID, sTBBookingStatusValue, sInsertUpdateDef, sRealSource, iTimestamp, lngOriginalRecordID)
				lngRecordID = result.RecordID
				sAction = result.Action
				sErrorMsg = result.Message
				sTBResultCode = result.TBResultCode
				sCourseOverbooked = result.CourseOverbooked
				fWarning = result.Warning
				fOk = result.OK

			ElseIf sAction = "DELETE" Then
				Dim result = _controllerRecord.data_submit_DELETE(lngTableID, sRealSource, lngRecordID, sReaction)
				lngRecordID = result.RecordID
				sAction = result.Action
				sErrorMsg = result.Message
				Session("currentPageID") = 0

			ElseIf sAction = "CANCELCOURSE" Then
				Dim result = _controllerTraining.data_submit_CancelCourse(lngRecordID, sRealSource)
				sAction = result.Action
				Session("numberOfBookings") = result.NumberOfBookings
				Session("tbErrorMessage") = result.Message
				Session("tbCourseTitle") = result.CourseTitle

			ElseIf sAction = "CANCELCOURSE_2" Then
				Dim txtTBCreateWLRecords = CleanBoolean(dataViewModel.txtTBCreateWLRecords)
				Dim result = _controllerTraining.data_submit_CancelCourse2(lngRecordID, sRealSource, iTBCourseRecordID, txtTBCreateWLRecords)
				sErrorMsg = result.Message
				sAction = result.Action

			ElseIf sAction = "CANCELBOOKING" Then
				Dim result = _controllerTraining.data_submit_CancelBooking(fUserChoice, lngRecordID)
				sErrorMsg = result.Message
				sAction = result.Action


			Else
				' randomly called for no obvious reason! (but it happens... a lot...)


			End If

			Session("selectSQL") = dataViewModel.txtSelectSQL
			Session("fromDef") = dataViewModel.txtFromDef

			Session("realSource") = sRealSource
			Session("tableID") = lngTableID
			Session("screenID") = lngScreenID
			Session("viewID") = lngViewID
			Session("recordID") = lngRecordID
			Session("action") = sAction
			Session("reaction") = ""
			Session("warningFlag") = fWarning
			Session("parentTableID") = ValidateIntegerValue(dataViewModel.txtParentTableID)
			Session("parentRecordID") = ValidateIntegerValue(dataViewModel.txtParentRecordID)
			Session("defaultCalcColumns") = dataViewModel.txtDefaultCalcCols
			Session("insertUpdateDef") = sInsertUpdateDef
			Session("errorMessage") = sErrorMsg
			Session("ReportBaseTableID") = ValidateIntegerValue(dataViewModel.txtReportBaseTableID)
			Session("ReportParent1TableID") = ValidateIntegerValue(dataViewModel.txtReportParent1TableID)
			Session("ReportParent2TableID") = ValidateIntegerValue(dataViewModel.txtReportParent2TableID)
			Session("ReportChildTableID") = ValidateIntegerValue(dataViewModel.txtReportChildTableID)
			Session("Param1") = dataViewModel.txtParam1

			'JDM - 24/07/02 - Fault 3917 - Reset year for absence calendar
			Session("stdrpt_AbsenceCalendar_StartYear") = Year(DateTime.Now())

			'JDM - 10/10/02 - Fault 4534 - Reset start month for absence calendar
			Session("stdrpt_AbsenceCalendar_StartMonth") = ""

			'TM - 05/09/02 - Store the event log parameters in session vaiables.
			Session("ELFilterUser") = dataViewModel.txtELFilterUser
			Session("ELFilterType") = dataViewModel.txtELFilterType
			Session("ELFilterStatus") = dataViewModel.txtELFilterStatus
			Session("ELFilterMode") = dataViewModel.txtELFilterMode
			Session("ELOrderColumn") = "datetime"	'dataViewModel.txtELOrderColumn
			Session("ELOrderOrder") = "DESC"	'dataViewModel.txtELOrderOrder

			Session("ELAction") = dataViewModel.txtELAction

			Session("ELCurrentRecCount") = ValidateIntegerValue(dataViewModel.txtELCurrRecCount)
			If Session("ELCurrentRecCount") < 1 Or Len(Session("ELCurrentRecCount")) < 1 Then
				Session("ELCurrentRecCount") = 0
			End If

			Session("ELFirstRecPos") = ValidateIntegerValue(dataViewModel.txtEL1stRecPos)
			If Session("ELFirstRecPos") < 1 Or Len(Session("ELFirstRecPos")) < 1 Then
				Session("ELFirstRecPos") = 1
			End If

			' Go to the requested page.
			Return RedirectToAction("Data", "Home")

		End Function

		<ValidateAntiForgeryToken()>
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
					If Session("LoggedInUserRecordID") >= 0 Then
						Session("TopLevelRecID") = Session("LoggedInUserRecordID")
					Else
						' More than one personnel record.
						sErrorDescription = "You have access to more than one record in the defined Single-record view."

						Session("ErrorTitle") = "Login Page"
						Session("ErrorText") =
						 "You could not login to the OpenHR database because of the following reason:" & sErrorDescription & "<p>" & vbCrLf

						Response.Redirect("FormError")
					End If

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
					If sViewName.Length > 0 Then sViewDescription = sViewName.Replace("_", " ") & sViewDescription.Replace("'", "\'")

					Session("ViewDescription") = sViewDescription

				Catch ex As Exception
					Throw

				End Try

			Else
				Session("ViewDescription") = Session("welcomemessage")
			End If

			Session("welcomemessage") = ""	' reset welcome message so that it shows just once

			Dim objNavigation = New HR.Intranet.Server.clsNavigationLinks
			objNavigation.SessionInfo = CType(Session("SessionContext"), SessionInfo)

			objNavigation.ClearLinks()
			objNavigation.SSITableID = Session("SSILinkTableID")
			objNavigation.SSIViewID = Session("SSILinkViewID")
			objNavigation.LoadLinks()
			objNavigation.LoadNavigationLinks()

			If Not Licence.IsModuleLicenced(SoftwareModule.Workflow) Then
				objNavigation.ColLinks.RemoveAll(Function(m) m.UtilityType = UtilityType.utlWorkflow)
			End If

			If Not Licence.IsModuleLicenced(SoftwareModule.NineBoxGrid) Then
				objNavigation.ColLinks.RemoveAll(Function(m) m.UtilityType = UtilityType.utlNineBoxGrid)
			End If

			Dim viewModel = New NavLinksViewModel With {.NavigationLinks = objNavigation.GetAllLinks, .NumberOfLinks = objNavigation.GetAllLinks.Count, .DocumentDisplayLinkCount = objNavigation.GetLinks(LinkType.DocumentDisplay).Count}
			Session("NavigationLinks") = objNavigation

			ViewData("showOutOfOffice") = ShowOutOfOffice(NullSafeInteger(Session("SSILinkTableID")), NullSafeInteger(Session("SSILinkViewID")))

			Return View(viewModel)
		End Function

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

			Dim mrstChartData As DataTable
			Dim sErrorDescription As String

			Dim objChart = New HR.Intranet.Server.clsChart
			objChart.SessionInfo = CType(Session("SessionContext"), SessionInfo)

			Try

				mrstChartData = objChart.GetChartData(tableID, columnID, filterID, aggregateType, elementType, 0, 0, 0, 0, sortOrderID, sortDirection, colourID)

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
								' kill the title
								title = ""
								If Not String.IsNullOrEmpty(title) Then
									chart1.Titles.Add("MainTitle")
									chart1.Titles(0).Text = title
									chart1.Titles(0).Font = New System.Drawing.Font(chart1.Titles(0).Font.Name, 20) 'Set the font size without changing the font family
								End If

								chart1.ChartAreas.Add("ChartArea1")

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


								If chartType = 0 Or chartType = 2 Or chartType = 4 Or chartType = 6 Then
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
											chart1.Series("Default").Points.Add(New DataPoint() With {.AxisLabel = objRow(0).ToString(), .YValues = New Double() {objRow(1)}, .Color = pointBackColor})
										Else
											chart1.Series("Default").Points.Add(New DataPoint() With {.Label = " ", .YValues = New Double() {objRow(1)}, .Color = pointBackColor})
										End If

										If showLegend = True Then
											chart1.Legends("Default").CustomItems.Add(New LegendItem(objRow(0).ToString(), pointBackColor, ""))
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

			Catch ex As Exception
				sErrorDescription = "The Chart field values could not be retrieved." & vbCrLf & ex.Message.RemoveSensitive()

			End Try

		End Function

		' This function creates and HTML table with the chart data, as well as jquery script that wll turn the table into a jqGrid
		Function GetChartDataAsHTML(
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
											MultiAxisChart As Boolean) As String

			Dim ChartData As DataTable

			Dim objChart = New HR.Intranet.Server.clsChart
			objChart.SessionInfo = CType(Session("SessionContext"), SessionInfo)
			If MultiAxisChart Then
				ChartData = objChart.GetChartData(tableID, columnID, filterID, aggregateType, elementType, tableID_2, columnID_2, tableID_3, columnID_3, sortOrderID, sortDirection, colourID)
			Else
				ChartData = objChart.GetChartData(tableID, columnID, filterID, aggregateType, elementType, 0, 0, 0, 0, sortOrderID, sortDirection, colourID)
			End If

			Dim HTMLTable As String = ""
			Dim colNames As String = ""
			Dim colModel As String = ""
			Dim Script As String = ""

			'Create the HTML table with the data
			HTMLTable = "<table id='chartData'>"
			HTMLTable = String.Concat(HTMLTable, "<tr>")
			For Each col As DataColumn In ChartData.Columns
				HTMLTable = String.Concat(HTMLTable, "<th>", col.ColumnName, "</th>")
				colNames = String.Concat(colNames, "'", col.ColumnName, "', ")
				colModel = String.Concat(colModel, "{ name: '", col.ColumnName, "', index: '", col.ColumnName, "', sortable: 'true'},")
			Next
			colNames = colNames.TrimEnd(",") 'Remove extra comma
			colModel = colModel.TrimEnd(",") 'Remove extra comma
			HTMLTable = String.Concat(HTMLTable, "</tr>")

			'Loop over the records
			For Each objRow As DataRow In ChartData.Rows
				HTMLTable = String.Concat(HTMLTable, "<tr>")
				For Each col As DataColumn In ChartData.Columns
					HTMLTable = String.Concat(HTMLTable, "<td>", objRow(col).ToString(), "</td>")
				Next
				HTMLTable = String.Concat(HTMLTable, "</tr>")
			Next
			HTMLTable = String.Concat(HTMLTable, "</table>")

			'Create the script that will turn the table above into a datagrid
			Script = String.Concat("<script type='text/javascript'>", _
														 "tableToGrid('#chartData', {")
			colNames = String.Concat("colNames:[", colNames, "]")
			colModel = String.Concat("colModel:[", colModel, "]")
			'Script = String.Concat(Script,
			'										 colNames, ",", _
			'										 colModel, ",", _
			'										 "rownum: 1000,", _
			'										 "scroll: true,", _
			'										 "autowidth: true", _
			'										 "});", _
			'										 "</script>")

			Script = String.Concat(Script,
													 colNames, ",", _
													 colModel, _
													 "});", _
													 "</script>")

			Return String.Concat(HTMLTable, Script)
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

			Dim RotateX As Integer = HttpContext.Request.QueryString("rotateX")
			If RotateX = 0 Then RotateX = 15
			Dim RotateY As Integer = HttpContext.Request.QueryString("rotateY")
			If RotateY = 0 Then RotateY = 10

			Dim mrstChartData As DataTable
			Dim sErrorDescription As String

			Try

				Dim objChart = New HR.Intranet.Server.clsMultiAxisChart
				objChart.SessionInfo = CType(Session("SessionContext"), SessionInfo)

				' Pass required info to the DLL
				mrstChartData = objChart.GetChartData(tableID, columnID, filterID, aggregateType, elementType, tableID_2, columnID_2, tableID_3, columnID_3, sortOrderID, sortDirection, colourID)

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
								' kill the title
								title = ""
								If Not String.IsNullOrEmpty(title) Then
									MultiAxisChart.Titles.Add("MainTitle")
									MultiAxisChart.Titles(0).Text = title
									MultiAxisChart.Titles(0).Font = New System.Drawing.Font(MultiAxisChart.Titles(0).Font.Name, 20) 'Set the font size without changing the font family
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
									MultiAxisChartHorizontals.Add(New MultiAxisChartHorizontal With {.Horizontal_ID = r("HORIZONTAL_ID"), .Horizontal = iif(isdbnull(r("HORIZONTAL")), "", r("HORIZONTAL")), .Colour = r("COLOUR")})
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

									Dim safeHorizontalString As string = iif(isdbnull(objRow("HORIZONTAL")), "", objRow("HORIZONTAL").tostring())
									Dim safeVerticalString As string = iif(isdbnull(objRow("VERTICAL")), "", objRow("VERTICAL").tostring())
									
									If TryCast(objRow("HORIZONTAL_ID"), String) <> "No Access" And TryCast(objRow("HORIZONTAL_ID"), String) <> "No Data" Then
										seriesName = safeVerticalString
										If seriesName = "" Then
											seriesName = "(No name)"
										End If
										Dim columnName As String = safeHorizontalString
										Dim yVal As Double = objRow("Aggregate")
										Dim pointBackColor As Color

										pointNum = CInt(objRow("HORIZONTAL_ID"))

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
													.YValues = New Double() {yVal},
													.AxisLabel = columnName,
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
												If legItem.Name = safeHorizontalString Then legendAdded = True
											Next

											If Not legendAdded Then
												MultiAxisChart.Legends("Default").CustomItems.Add(New LegendItem(safeHorizontalString, pointBackColor, ""))
											End If
										End If

									End If

								Next

								'For 2D pie charts with more than one series we need to add a chart area for each series
								Dim thisSeries As String
								For Each s As Series In MultiAxisChart.Series
									'Add a chart area for the series and set its properties
									thisSeries = s.Name
									If chartType = 14 And MultiAxisChart.Series.Count > 1 Then '2D pie
										MultiAxisChart.ChartAreas.Add(thisSeries)
									Else
										thisSeries = "Default"
										If MultiAxisChart.ChartAreas.Count = 0 Then
											MultiAxisChart.ChartAreas.Add(thisSeries)
										End If
									End If

									MultiAxisChart.ChartAreas(thisSeries).BackSecondaryColor = Color.Transparent
									MultiAxisChart.ChartAreas(thisSeries).ShadowColor = Color.Transparent
									MultiAxisChart.ChartAreas(thisSeries).AxisY.LineColor = Color.FromArgb(64, 64, 64, 64)
									MultiAxisChart.ChartAreas(thisSeries).AxisY.MajorGrid.LineColor = Color.FromArgb(64, 64, 64, 64)
									MultiAxisChart.ChartAreas(thisSeries).AxisX.LineColor = Color.FromArgb(64, 64, 64, 64)
									MultiAxisChart.ChartAreas(thisSeries).AxisX.MajorGrid.LineColor = Color.FromArgb(64, 64, 64, 64)

									MultiAxisChart.ChartAreas(thisSeries).AxisX.LabelStyle.Enabled = showLabels
									MultiAxisChart.ChartAreas(thisSeries).AxisY.LabelStyle.Enabled = showLabels

									' Gridlines
									If dottedGrid = True Then
										MultiAxisChart.ChartAreas(thisSeries).AxisX.LineDashStyle = ChartDashStyle.Dot
										MultiAxisChart.ChartAreas(thisSeries).AxisY.LineDashStyle = ChartDashStyle.Dot
										MultiAxisChart.ChartAreas(thisSeries).AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dot
										MultiAxisChart.ChartAreas(thisSeries).AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot
									Else
										MultiAxisChart.ChartAreas(thisSeries).AxisX.LineDashStyle = ChartDashStyle.NotSet
										MultiAxisChart.ChartAreas(thisSeries).AxisY.LineDashStyle = ChartDashStyle.NotSet
										MultiAxisChart.ChartAreas(thisSeries).AxisX.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
										MultiAxisChart.ChartAreas(thisSeries).AxisY.MajorGrid.LineDashStyle = ChartDashStyle.NotSet
									End If

									' 3D Settings
									If chartType = 0 Or chartType = 2 Or chartType = 4 Or chartType = 6 Then
										MultiAxisChart.ChartAreas(thisSeries).Area3DStyle.Enable3D = True
										MultiAxisChart.ChartAreas(thisSeries).Area3DStyle.Perspective = 10
										MultiAxisChart.ChartAreas(thisSeries).Area3DStyle.Inclination = RotateX
										MultiAxisChart.ChartAreas(thisSeries).Area3DStyle.Rotation = RotateY
										MultiAxisChart.ChartAreas(thisSeries).Area3DStyle.IsRightAngleAxes = False
										MultiAxisChart.ChartAreas(thisSeries).Area3DStyle.WallWidth = 0
										MultiAxisChart.ChartAreas(thisSeries).Area3DStyle.IsClustered = False
									End If

									If chartType = 14 And MultiAxisChart.Series.Count > 1 Then '2D pie
										'Add an annotation (legend) to the chart area
										Dim chartLegend As New RectangleAnnotation
										chartLegend.Text = thisSeries
										chartLegend.AxisX = MultiAxisChart.ChartAreas(thisSeries).AxisX
										chartLegend.AxisY = MultiAxisChart.ChartAreas(thisSeries).AxisY
										chartLegend.AnchorX = MultiAxisChart.ChartAreas(thisSeries).Position.X
										chartLegend.AnchorY = MultiAxisChart.ChartAreas(thisSeries).Position.Y
										MultiAxisChart.Annotations.Add(chartLegend)
									End If

									MultiAxisChart.Series(s.Name).ChartArea = thisSeries 'Assign the series to the chart area

									If showLabels Then
										MultiAxisChart.ChartAreas(thisSeries).AxisX.Interval = 1 'Show all X axis legends (labels?)
									End If
								Next

								If showLabels Then
									Try
										'The "AlignDataPointsByAxisLabel" method below, "Aligns data points along the X axis using their axis labels. Applicable when multiple series are indexed and their X-values are strings."
										'according to MSDN (http://msdn.microsoft.com/en-us/library/system.web.ui.datavisualization.charting.chart.aligndatapointsbyaxislabel(v=vs.100).aspx)
										'Instead of checking that all series are indexed and their X-values are strings, I decided to put a Try-Catch and be done with it
										MultiAxisChart.AlignDataPointsByAxisLabel()
									Catch ex As Exception

									End Try
								End If

								'Make all the datapoints semi-transparent							
								MultiAxisChart.ApplyPaletteColors()
								For Each series As Series In MultiAxisChart.Series
									For Each point As DataPoint In series.Points
										point.Color = Color.FromArgb(180, point.Color)
									Next
								Next

								'Helpful for debugging: save the chart data as XML
								'MultiAxisChart.Serializer.Save(Server.MapPath("/") & "Chart.xml")
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

			Catch ex As Exception
				sErrorDescription = "The Chart field values could not be retrieved." & vbCrLf & ex.Message.RemoveSensitive

			End Try

		End Function

		Function PasswordChange() As ActionResult
			Return View()
		End Function

		<PermissionAuthorizeAttribute(Category:="INTRANET", Item:="NEW USER")>
		Function NewUser() As ActionResult
			Return View()
		End Function

#Region "Event Log Forms"

		'<ValidateAntiForgeryToken>
		'<PermissionAuthorizeAttribute(Category:="EVENTLOG", Item:="EMAIL")>
		'Function EventLogEmail(value As EmailSelectionModel) As ActionResult
		'	Return View(value)
		'End Function

		<HttpPost>
		<ValidateAntiForgeryToken>
		Function EventLog() As ActionResult
			Return View()
		End Function

		<ValidateAntiForgeryToken>
		Function EventLogDetails(value As EventDetailModel) As ActionResult
			Return View(value)
		End Function

		<PermissionAuthorizeAttribute(Category:="EVENTLOG", Item:="PURGE")>
		Function EventLogPurge() As ActionResult
			Return View()
		End Function

		<PermissionAuthorizeAttribute(Category:="EVENTLOG", Item:="DELETE")>
		Function EventLogDelete() As ActionResult
			Return View()
		End Function

		<HttpPost>
		<ValidateAntiForgeryToken>
		<PermissionAuthorizeAttribute(Category:="EVENTLOG", Item:="EMAIL")>
		Function SendEmail(postData As SendEmailModel) As ActionResult

			Dim emailTo As String = postData.To
			Dim emailCC As String = postData.CC
			Dim emailBCC As String = postData.BCC
			Dim emailSubject As String = postData.Subject
			Dim emailBody As String = postData.Body

			Try
				Dim message As New MailMessage()
				message.Subject = emailSubject
				message.Body = emailBody.ToString.Replace("\n", vbCrLf)

				If Not emailTo = "" Then
					If emailTo.Contains(";") = True Then

						Dim aRecipientList = Split(emailTo, ";")

						For iLoop = 0 To UBound(aRecipientList) - 1
							message.To.Add(aRecipientList(iLoop))
						Next
					Else
						message.To.Add(emailTo)
					End If
				End If

				If Not emailCC = "" Then
					If emailCC.Contains(";") = True Then

						Dim aRecipientList = Split(emailCC, ";")

						For iLoop = 0 To UBound(aRecipientList) - 1
							message.CC.Add(aRecipientList(iLoop))
						Next
					Else
						message.CC.Add(emailCC)
					End If
				End If

				If Not emailBCC = "" Then
					If emailBCC.Contains(";") = True Then

						Dim aRecipientList = Split(emailBCC, ";")

						For iLoop = 0 To UBound(aRecipientList) - 1
							message.Bcc.Add(aRecipientList(iLoop))
						Next
					Else
						message.Bcc.Add(emailBCC)
					End If
				End If

				Dim client As New SmtpClient()

				client.Send(message)

				Return New HttpStatusCodeResult(HttpStatusCode.OK, "Email sent successfully")
			Catch ex As Exception
				' error generated - return error
				Dim errMessage As String
				If ex.InnerException Is Nothing Then
					errMessage = ""
				Else
					errMessage = ex.InnerException.Message
				End If

				Dim strErrors As String = ""

				If emailTo = "" And emailCC = "" And emailBCC = "" Then
					strErrors = "Please select recipient(s) to email"
				Else
					strErrors = String.Format("The following error occured when emailing your document:" _
						& "{0}{0}{1}{0}{0}{2}{0}{0}Please check with your administrator for further details.", "<br/>", _
						ex.Message, errMessage)
				End If

				'I used StatusCode BadRequest (400) below instead of StatusCode InternalServerError (500) because error 500 is not being properly caught by the ajax call in emailSelection.aspx
				Return New HttpStatusCodeResult(HttpStatusCode.BadRequest, strErrors)
			End Try
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
		<ValidateAntiForgeryToken>
		Function util_run_crosstabsDataSubmit(value As CrossTabDataSubmitModel)
			Session("CT_Mode") = value.txtMode
			Session("CT_EmailGroupID") = value.txtEmailGroupID
			Session("CT_UtilID") = value.txtUtilID

			If value.txtMode = "BREAKDOWN" Then
				Session("CT_Hor") = value.txtHor
				Session("CT_Ver") = value.txtVer
				Session("CT_Pgb") = value.txtPgb
				Session("CT_IntersectionType") = value.txtIntersectionType
				Session("CT_CellValue") = value.txtCellValue
				Session("CT_Use1000") = value.txtUse1000
			Else
				Session("CT_PageNumber") = value.txtPageNumber
				Session("CT_IntersectionType") = value.txtIntersectionType
				Session("CT_ShowPercentage") = value.txtShowPercentage
				Session("CT_PercentageOfPage") = value.txtPercentageOfPage
				Session("CT_SuppressZeros") = value.txtSuppressZeros
				Session("CT_Use1000") = value.txtUse1000
			End If

			' Go to the requested page.
			Return RedirectToAction("util_run_crosstabsData")

		End Function

		' This is passes through to the utili run prompted values, but is called from a standard report. 
		' This is because we set some settings which are later called in cross tabs and custom reports
		' This could be improved with proper modelling!
		<ValidateAntiForgeryToken>
		Function util_run_standardreport_promptedvalues(value As StandardReportRunModel) As ActionResult
			Try

				Session("stdReport_StartDate") = value.txtFromDate
				Session("stdReport_EndDate") = value.txtToDate
				Session("stdReport_AbsenceTypes") = value.txtAbsenceTypes
				Session("stdReport_FilterID") = value.txtBaseFilterID
				Session("stdReport_FilterName") = value.txtFilterName
				Session("stdReport_PicklistID") = value.txtBasePicklistID
				Session("stdReport_PicklistName") = value.txtPicklistName
				Session("stdReport_Bradford_SRV") = value.txtSRV
				Session("stdReport_Bradford_ShowDurations") = value.txtShowDurations
				Session("stdReport_Bradford_ShowFormula") = value.txtShowFormula
				Session("stdReport_Bradford_ShowInstances") = value.txtShowInstances
				Session("stdReport_Bradford_OmitBeforeStart") = value.txtOmitBeforeStart
				Session("stdReport_Bradford_OmitAfterEnd") = value.txtOmitAfterEnd
				Session("stdReport_Bradford_txtOrderBy1") = value.txtOrderBy1
				Session("stdReport_Bradford_txtOrderBy1ID") = value.txtOrderBy1ID
				Session("stdReport_Bradford_txtOrderBy1Asc") = value.txtOrderBy1Asc
				Session("stdReport_Bradford_txtOrderBy2") = value.txtOrderBy2
				Session("stdReport_Bradford_txtOrderBy2ID") = value.txtOrderBy2ID
				Session("stdReport_Bradford_txtOrderBy2Asc") = value.txtOrderBy2Asc
				Session("stdReport_PrintFilterPicklistHeader") = value.txtPrintFPinReportHeader
				Session("stdReport_MinimumBradfordFactor") = value.txtMinimumBradfordFactor
				Session("stdReport_MinimumBradfordFactorAmount") = value.txtMinimumBradfordFactorAmount
				Session("stdReport_DisplayBradfordDetail") = value.txtDisplayBradfordDetail

				Session("stdReport_OutputPreview") = value.txtSend_OutputPreview
				Session("stdReport_OutputFormat") = value.txtSend_OutputFormat
				Session("stdReport_OutputSave") = value.txtSend_OutputSave
				Session("stdReport_OutputSaveExisting") = 0
				Session("stdReport_OutputEmail") = value.txtSend_OutputEmail
				Session("stdReport_OutputEmailAddr") = value.txtSend_OutputEmailAddr
				Session("stdReport_OutputEmailSubject") = value.txtSend_OutputEmailSubject
				Session("stdReport_OutputEmailAttachAs") = value.txtSend_OutputEmailAttachAs
				Session("stdReport_OutputFilename") = value.txtSend_OutputFilename

				Session("utiltype") = CInt(value.UtilType)
				Session("utilid") = value.ID
				Session("utilname") = value.Name
				Session("action") = "run"

			Catch ex As Exception
				Throw

			End Try

			Return View("util_run_promptedvalues", value)
		End Function

		<ValidateAntiForgeryToken>
		Function util_run_promptedvalues(value As PromptedValuesModel) As ActionResult
			Try

				Session("utiltype") = CInt(value.UtilType)
				Session("utilid") = value.ID
				Session("utilname") = value.Name
				Session("action") = "run"

			Catch ex As Exception
				Throw

			End Try

			Return View(value)
		End Function

		<HttpPost()>
<ValidateAntiForgeryToken>
		Function util_run_promptedvalues_submit(value As PromptedValuesModel) As ActionResult

			Try

				Dim sKey As String
				Dim sPrompts As String = ""

				Dim aPrompts(1, 0) As String
				Dim j = 0
				ReDim Preserve aPrompts(1, 0)

				If value.PromptValues IsNot Nothing Then
					For Each objPrompt In value.PromptValues
						sKey = objPrompt.Key
						If ((UCase(Left(sKey, 7)) = "PROMPT_") And (Mid(sKey, 8, 1) <> "3")) Or _
								(UCase(Left(sKey, 10)) = "PROMPTCHK_") Then
							ReDim Preserve aPrompts(1, j)

							If (UCase(Left(sKey, 10)) = "PROMPTCHK_") Then
								aPrompts(0, j) = "prompt_3_" & Mid(sKey, 11)
								aPrompts(1, j) = UCase(objPrompt.Value)
							Else
								aPrompts(0, j) = sKey
								Select Case objPrompt.Type
									Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
										' Numeric. Replace locale decimal point with '.'
										aPrompts(1, j) = Replace(objPrompt.Value, CType(Session("LocaleDecimalSeparator"), String), ".")
									Case ExpressionValueTypes.giEXPRVALUE_DATE
										' Date. Reformat to match SQL's mm/dd/yyyy format.
										aPrompts(1, j) = ConvertLocaleDateToSQL(objPrompt.Value)
									Case Else
										aPrompts(1, j) = objPrompt.Value
								End Select
							End If

							sPrompts = sPrompts & aPrompts(0, j) & vbTab & aPrompts(1, j) & vbTab
							j += 1
						End If
					Next
				End If

				sKey = "Prompts_" & CInt(value.UtilType) & "_" & value.ID.ToString
				Session("promptsvalue") = sPrompts
				Session(sKey) = aPrompts

				' Act dependent on utility type
				Select Case value.UtilType
					Case UtilityType.utlDataTransfer

            Dim filteredIds as string
					  If CInt(Session("singleRecordID")) > 0 then
              filteredIds =  Session("singleRecordID") 
            Else
					    filteredIds = Session("multipleRecordIDs")
					  End If

						Dim message = RunDataTransfer(value.ID, filteredIds, aPrompts)
						Return View("util_run_message", message)

					Case Else
						Return View("util_run", value)

				End Select

			Catch ex As Exception
				Throw

			End Try

		End Function
    
		Function GetTalentReportData() As ActionResult

			Dim matchReport = New MatchReportRun
			matchReport.SessionInfo = CType(Session("SessionContext"), SessionInfo)

			Dim prompts = Session("Prompts_" & Session("utiltype") & "_" & Session("utilid"))

			matchReport.UtilityType = CType(Session("utiltype"), UtilityType)
			matchReport.SetPromptedValues(prompts)
			matchReport.MatchReportID = CInt(Session("utilid"))
			matchReport.RunMatchReport(Session("utiltableID"), Session("singleRecordID"))

      Session("MatchReport") = matchReport

		  If len(matchReport.ErrorString) > 0 Then
        Return Json(matchReport.ErrorString, JsonRequestBehavior.AllowGet)
      Else

			  Dim rows As New List(Of Dictionary(Of String, Object))()
			  Dim row As Dictionary(Of String, Object)

			  For Each dr As DataRow In matchReport.ReportDataTable.Rows
				  row = New Dictionary(Of String, Object)()

                    For Each col As DataColumn In matchReport.ReportDataTable.Columns
                        row.Add(col.ColumnName, dr(col))
                    Next
				  rows.Add(row)
			  Next

  		  Dim colModel As List(Of Object) = JqGridColModel.CreateColModel(matchReport.ReportDataTable, "", "", True)

			  Dim results = New With {.total = 1, .page = 1, .records = 0, .rows = rows, .colModel = colModel}
			  Return Json(results, JsonRequestBehavior.AllowGet)
        
 		  End If


		End Function

 		<HttpPost>
		<ValidateAntiForgeryToken>
		Public Function util_run_talentreport_downloadoutput() As FilePathResult

			Dim lngOutputFormat = CType(Request("txtFormat"), OutputFormats)
			Dim bPreview As Boolean = Request("txtPreview")
			Dim blnEmail As Boolean = Request("txtEmail")
			Dim lngEmailGroupId As Integer = Request("txtEmailGroupID")
			Dim strEmailSubject As String = Request("txtEmailSubject")
			Dim strEmailAttachAs As String = Request("txtEmailAttachAs")
			Dim downloadTokenValue As String = Request("download_token_value_id")
      Dim objReportData As MatchReportRun = CType(Session("MatchReport"), MatchReportRun)
      Dim outputFile = Path.GetTempFileName.Replace(".tmp", ".xlsx")

      If Len(objReportData.ErrorString) > 0 Then
        Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client		
				Response.AppendCookie(New HttpCookie("fileDownloadErrors", objReportData.ErrorString))	' marks the download as complete on the client	
        Exit Function
      End If

			Dim strDownloadFileName = objReportData.DownloadFileName

			Dim objDocument As New Workbook
			Dim objCellsLicense As New License
			objCellsLicense.SetLicense("Aspose.Cells.lic")

      objDocument.DefaultStyle.Font.Name = "Calibri"
			objDocument.DefaultStyle.Font.Size += 1

      Dim worksheet = objDocument.Worksheets(0)
			worksheet.Name = "Data"

			'Sort output in DESC mode by Match Score column
			Dim dataView As New DataView(objReportData.ReportDataTable)
			dataView.Sort = "Match Score % DESC"
			objReportData.ReportDataTable = dataView.ToTable

			' Rename last column with the second last column. (Second last column bullet graph would be removed)
			Dim colCount = objReportData.ReportDataTable.Columns.Count
			objReportData.ReportDataTable.Columns(colCount - 1).Caption = objReportData.ReportDataTable.Columns(colCount - 2).Caption

      worksheet.Cells.ImportDataTable(objReportData.ReportDataTable, true, "A1")

      worksheet.Cells.DeleteColumns(0,2, True)
      For columnCount = 0 To worksheet.Cells.MaxDataColumn
				worksheet.AutoFitColumn(columnCount)
      Next

			worksheet.Cells.DeleteColumns(worksheet.Cells.MaxDataColumn - 1, 1, True)

			' Build the talent chart column (.ApplyStyleColumn does not seem to work?!?!?)
			Dim style As New Style With {.Number = 49, .IsTextWrapped = True}
      For rowNumber = 0 To worksheet.Cells.MaxDataRow
        worksheet.Cells(rowNumber, worksheet.Cells.MaxDataColumn).SetStyle(style)
			Next

			objDocument.Save(outputFile, SaveFormat.Xlsx)

			' Send email
			If blnEmail And lngEmailGroupId > 0 Then
				Dim sEmailAddresses = GetEmailAddressesForGroup(lngEmailGroupId)

				Dim downloadedDoc As New FileStream(outputFile, FileMode.Open)
				Try
					If IO.File.Exists(outputFile) Then
						SendMailWithAttachment(strEmailSubject, downloadedDoc, sEmailAddresses, strEmailAttachAs)
						Response.AppendCookie(New HttpCookie("fileDownloadErrors", "Email sent successfully."))	' Send completion message	
					End If
				Catch ex As Exception
					' error generated - return error
					Dim errMessage As String
					If ex.InnerException Is Nothing Then
						errMessage = ""
					Else
						errMessage = ex.InnerException.Message
					End If

					Dim strErrors = String.Format("The following error occured when emailing your document:" _
						& "{0}{0}{1}{0}{0}{2}{0}{0}Please check with your administrator for further details.", "<br/>", _
						ex.Message, errMessage)

					Response.AppendCookie(New HttpCookie("fileDownloadErrors", strErrors))	' marks the download as complete on the client		
				Finally
					Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client		
				End Try
			End If

			' Download the file			
			If (lngOutputFormat = OutputFormats.ExcelWorksheet And Not blnEmail) Then
				If IO.File.Exists(outputFile) Then
					Try
						Dim fileInfo As FileInfo = New FileInfo(outputFile)
						Response.ContentType = "application/octet-stream"
						Response.Clear()
						Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client
						Response.AppendCookie(New HttpCookie("fileDownloadErrors", vbNullString))	' Clear error message response cookie
						Response.AddHeader("Content-Disposition", String.Format("attachment;filename=""{0}""", strDownloadFileName))
						Response.AddHeader("Content-Length", fileInfo.Length.ToString())
						Response.WriteFile(fileInfo.FullName)
						'Response.End()
						Response.Flush()
					Catch ex As Exception
						' error generated - return error
						Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client		
						Response.AppendCookie(New HttpCookie("fileDownloadErrors", ex.Message))	' marks the download as complete on the client		
					Finally
						IO.File.Delete(outputFile)
					End Try
				Else
					' No file generated - return error
					Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client		
					Response.AppendCookie(New HttpCookie("fileDownloadErrors", "No output file was generated. Check your data."))	' marks the download as complete on the client		
				End If
			End If

		End Function

		Private Function RunDataTransfer(id As Integer, multipleRecordIds As String, prompts(,) As String) As PostResponse

			Dim dataTransfer = New clsDataTransferRun
			dataTransfer.SessionInfo = CType(Session("SessionContext"), SessionInfo)

			dataTransfer.SetPromptedValues(prompts)
			dataTransfer.ExecuteDataTransfer(id, multipleRecordIds)

			Dim message As New PostResponse With {
					.Message = dataTransfer.StatusMessage
			}

			Return message

		End Function
		
		<HttpPost>
<ValidateAntiForgeryToken>
			Public Function util_run_crosstab_downloadoutput() As FilePathResult

			Dim lngOutputFormat As OutputFormats = Request("txtFormat")
			Dim bPreview As Boolean = Request("txtPreview")
			Dim sUtilID As String = Request("txtUtilID")
			Dim blnSavetoFile As Boolean = Request("txtSave")
			Dim lngSaveExisting As Long = Request("txtSaveExisting")
			Dim blnEmail As Boolean = Request("txtEmail")
			Dim lngEmailGroupID As Integer = Request("txtEmailGroupID")
			Dim strEmailSubject As String = Request("txtEmailSubject")
			Dim strEmailAttachAs As String = Request("txtEmailAttachAs")
			Dim strDownloadFileName As String = Request("txtFilename")
			Dim downloadTokenValue As String = Request("download_token_value_id")
			Dim strDownloadExtension As String
			Dim strInterSectionType As String
			Dim sEmailAddresses As String

			Dim lngLoopMin As Long
			Dim lngLoopMax As Long

			Dim objCrossTab As CrossTab = CType(Session("objCrossTab" & sUtilID), CrossTab)

			Dim ClientDLL As New clsOutputRun
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
				, 4)

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
			If Not bPreview And objCrossTab.CrossTabType = CrossTabType.cttNormal Then
				lngOutputFormat = objCrossTab.OutputFormat
				blnSavetoFile = objCrossTab.OutputSave
				lngSaveExisting = objCrossTab.OutputSaveExisting
				blnEmail = objCrossTab.OutputEmail
				lngEmailGroupID = objCrossTab.OutputEmailID
				strEmailSubject = objCrossTab.OutputEmailSubject
				strEmailAttachAs = GetEmailAttachFileNameWithExcelExtension(objCrossTab.OutputEmailAttachAs)
				strDownloadFileName = objCrossTab.DownloadFileName
			End If

			If strDownloadFileName.Length = 0 Then
				objCrossTab.OutputFormat = lngOutputFormat
				objCrossTab.OutputFilename = ""
				strDownloadFileName = objCrossTab.DownloadFileName
			End If

			strDownloadExtension = Path.GetExtension(strDownloadFileName)

			Dim fOK = ClientDLL.SetOptions(False, lngOutputFormat, False, False, "", True, lngSaveExisting _
				, blnEmail, lngEmailGroupID, strEmailSubject, strEmailAttachAs, strDownloadExtension)

			If fOK Then
				If ClientDLL.GetFile() Then
					If lngOutputFormat = OutputFormats.DataOnly Then

					ElseIf lngOutputFormat = OutputFormats.ExcelPivotTable Then

						ClientDLL.AddColumn(" ", ColumnDataType.sqlVarChar, 0, False)
						For intCount = 0 To objCrossTab.ColumnHeadingUbound(0)
							ClientDLL.AddColumn(objCrossTab.ColumnHeading(0, intCount), ColumnDataType.sqlNumeric, objCrossTab.IntersectionDecimals, objCrossTab.Use1000Separator)
						Next

						If objCrossTab.CrossTabType = CrossTabType.cttAbsenceBreakdown Then
							ClientDLL.IntersectionType = IntersectionType.Count
						Else
							ClientDLL.IntersectionType = CInt(Session("CT_IntersectionType"))
						End If

						ClientDLL.AddColumn(strInterSectionType, ColumnDataType.sqlInteger, objCrossTab.IntersectionDecimals, objCrossTab.Use1000Separator)

						Dim strOutput(,) As String
						Dim strPageValue As String = ""
						Dim lngGroupNum As Integer
						Dim lngCol As Integer
						Dim lngRow As Integer

						With objCrossTab.PivotData

							If Not objCrossTab.PageBreakColumn Then
								lngRow = 1
								ReDim strOutput(.Columns.Count - 1, 0)
								For lngCol = 0 To .Columns.Count - 1
									strOutput(lngCol, 0) = objCrossTab.PivotData.Columns(lngCol).ColumnName
								Next
							End If

							For Each objRow As DataRow In objCrossTab.PivotData.Rows

								If objCrossTab.PageBreakColumn Then

									Dim sPageBreak As String
									If IsDate(objRow("Page Break")) Then
										sPageBreak = CDate(objRow("Page Break")).ToString(objCrossTab.LocaleDateFormat)
									Else
										sPageBreak = objRow("Page Break").ToString()
									End If

									If strPageValue <> sPageBreak Then

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

										If IsDate(objRow("Page Break")) Then
											strPageValue = CDate(objRow("Page Break")).ToString(objCrossTab.LocaleDateFormat)
										Else
											strPageValue = objRow("Page Break").ToString()
										End If

										lngRow = 1
										ReDim strOutput(.Columns.Count - 1, 0)
										For lngCol = 0 To .Columns.Count - 1
											strOutput(lngCol, 0) = .Columns(lngCol).ColumnName
										Next

									End If
								Else
									strPageValue = objCrossTab.BaseTableName

								End If

								ReDim Preserve strOutput(.Columns.Count - 1, lngRow)
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
							For lngRow = 0 To UBound(strOutput, 2)
								ClientDLL.ArrayAddTo(lngCol, lngRow, strOutput(lngCol, lngRow))
							Next
						Next

						ClientDLL.DataArray()
						ClientDLL.Complete()

					Else

						ClientDLL.AddColumn(" ", 12, 0, False)
						For intCount = 0 To objCrossTab.ColumnHeadingUbound(0)
							ClientDLL.AddColumn(Left(objCrossTab.ColumnHeading(0, intCount), 255), ColumnDataType.sqlNumeric, objCrossTab.IntersectionDecimals _
							, LCase(objCrossTab.Use1000Separator))
						Next

						strInterSectionType = Session("CT_IntersectionType")
						ClientDLL.AddColumn(strInterSectionType, ColumnDataType.sqlNumeric, objCrossTab.IntersectionDecimals, objCrossTab.Use1000Separator)


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
								ClientDLL.AddPage(sOutputGridCaption, Regex.Replace(Left(objCrossTab.ColumnHeading(2, lngCount), 255), "[:\\\/?\*\[\]]", " ")) 'Replace invalid characters with space so Aspose doesn't throw a wobbly when creating the Excel tabs
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


			' Email the generated file
			If blnEmail And lngEmailGroupID > 0 Then
				sEmailAddresses = GetEmailAddressesForGroup(lngEmailGroupID)

				Dim objDocument As New FileStream(ClientDLL.GeneratedFile, FileMode.Open)
				Try
					SendMailWithAttachment(strEmailSubject, objDocument, sEmailAddresses, strEmailAttachAs)
					Response.AppendCookie(New HttpCookie("fileDownloadErrors", "Email sent successfully."))	' Send completion message	
				Catch ex As Exception
					' error generated - return error
					Dim errMessage As String
					If ex.InnerException Is Nothing Then
						errMessage = ""
					Else
						errMessage = ex.InnerException.Message
					End If

					Dim strErrors = String.Format("The following error occured when emailing your document:" _
						& "{0}{0}{1}{0}{0}{2}{0}{0}Please check with your administrator for further details.", "<br/>", _
						ex.Message, errMessage)

					Response.AppendCookie(New HttpCookie("fileDownloadErrors", strErrors))	' marks the download as complete on the client		
				Finally
					Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client		
				End Try

			End If

			' Return the generated file
			If blnSavetoFile Or (Not blnSavetoFile And Not blnEmail) Then
				If IO.File.Exists(ClientDLL.GeneratedFile) Then
					Try
						Dim fileInfo As FileInfo = New FileInfo(ClientDLL.GeneratedFile)
						Response.ContentType = "application/octet-stream"
						Response.Clear()
						Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client
						If Not blnEmail Then Response.AppendCookie(New HttpCookie("fileDownloadErrors", vbNullString)) ' Clear error message response cookie
						Response.AddHeader("Content-Disposition", String.Format("attachment;filename=""{0}""", strDownloadFileName))
						Response.AddHeader("Content-Length", fileInfo.Length.ToString())
						Response.WriteFile(fileInfo.FullName)
						'Response.End()
						Response.Flush()
					Catch ex As Exception
						' error generated - return error
						Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client		
						Response.AppendCookie(New HttpCookie("fileDownloadErrors", ex.Message))	' marks the download as complete on the client		
					Finally
						IO.File.Delete(ClientDLL.GeneratedFile)
					End Try
				Else
					' No file generated - return error
					Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client		
					Response.AppendCookie(New HttpCookie("fileDownloadErrors", "No output file was generated. Check your data."))	' marks the download as complete on the client
				End If
			End If

		End Function

		<HttpPost>
		<ValidateAntiForgeryToken>
		Public Function util_run_nineboxgrid_downloadoutput() As FilePathResult

			Dim lngOutputFormat As OutputFormats = Request("txtFormat")
			Dim bPreview As Boolean = Request("txtPreview")
			Dim sUtilID As String = Request("txtUtilID")
			Dim blnSavetoFile As Boolean = Request("txtSave")
			Dim lngSaveExisting As Long = -1
			Dim blnEmail As Boolean = Request("txtEmail")
			Dim lngEmailGroupID As Integer = Request("txtEmailGroupID")
			Dim strEmailSubject As String = Request("txtEmailSubject")
			Dim strEmailAttachAs As String = Request("txtEmailAttachAs")
			Dim strDownloadFileName As String = Request("txtFilename")
			Dim downloadTokenValue As String = Request("download_token_value_id")
			Dim strDownloadExtension As String
			Dim strInterSectionType As String
			Dim sEmailAddresses As String

			Dim lngLoopMin As Long
			Dim lngLoopMax As Long

			Dim objCrossTab As CrossTab = CType(Session("objCrossTab" & sUtilID), CrossTab)

			Dim ClientDLL As New clsOutputRun
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
				, 4)

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
			ClientDLL.HeaderCols = 0
			ClientDLL.HeaderRows = 1

			'Set Options
			If Not bPreview And objCrossTab.CrossTabType = CrossTabType.ctt9GridBox Then
				lngOutputFormat = objCrossTab.OutputFormat
				blnSavetoFile = objCrossTab.OutputSave
				lngSaveExisting = -1
				blnEmail = objCrossTab.OutputEmail
				lngEmailGroupID = objCrossTab.OutputEmailID
				strEmailSubject = objCrossTab.OutputEmailSubject
				strEmailAttachAs = GetEmailAttachFileNameWithExcelExtension(objCrossTab.OutputEmailAttachAs)
				strDownloadFileName = objCrossTab.DownloadFileName
			End If

			If strDownloadFileName.Length = 0 Then
				objCrossTab.OutputFormat = lngOutputFormat
				objCrossTab.OutputFilename = ""
				strDownloadFileName = objCrossTab.DownloadFileName
			End If

			strDownloadExtension = Path.GetExtension(strDownloadFileName)

			Dim fOK = ClientDLL.SetOptions(False, lngOutputFormat, False, False, "", True, lngSaveExisting, blnEmail, lngEmailGroupID, strEmailSubject, strEmailAttachAs, strDownloadExtension)

			If fOK Then
				If ClientDLL.GetFile() Then
					ClientDLL.AddColumn(" ", 12, 0, False)
					For intCount = 3 To 5
						ClientDLL.AddColumn(Left(objCrossTab.ColumnHeading(0, intCount), 255), ColumnDataType.sqlNumeric, objCrossTab.IntersectionDecimals, LCase(objCrossTab.Use1000Separator))
					Next

					strInterSectionType = Session("CT_IntersectionType")
					ClientDLL.AddColumn(strInterSectionType, ColumnDataType.sqlNumeric, objCrossTab.IntersectionDecimals, objCrossTab.Use1000Separator)

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
							ClientDLL.AddPage(sOutputGridCaption, HttpUtility.HtmlDecode(Left(objCrossTab.ColumnHeading(2, lngCount), 255)))
						Else
							If objCrossTab.CrossTabType = CrossTabType.cttAbsenceBreakdown Then
								ClientDLL.AddPage(sOutputGridCaption, "Absence Breakdown")
							Else
								ClientDLL.AddPage(sOutputGridCaption, objCrossTab.BaseTableName)
							End If
						End If

						objCrossTab.BuildOutputStrings(lngCount)
						ClientDLL.ArrayDim(2, 2)
						ClientDLL.AxisLabelsAsArray = objCrossTab.AxisLabelsAsArray
						For intCol = 3 To 5
							For intRow = 5 To 3 Step -1
								ClientDLL.ArrayAddToNineBoxGrid(
																								intCol - 3, _
																								intRow - 3, _
																								objCrossTab.ReturnDescriptionOrColourForNineBoxGridCell(intCol - 3, intRow - 3, CrossTab.enumNineBoxDescriptionOrColour.Description), _
																								HttpUtility.HtmlDecode(Left(objCrossTab.DataArray(intCol, intRow), 255)), _
																								objCrossTab.ReturnDescriptionOrColourForNineBoxGridCell(intCol - 3, intRow - 3, CrossTab.enumNineBoxDescriptionOrColour.Colour))
							Next
						Next

						ClientDLL.DataArrayNineBoxGrid()
					Next

					ClientDLL.Complete()
				End If
			End If

			' Email the generated file
			If blnEmail And lngEmailGroupID > 0 Then
				sEmailAddresses = GetEmailAddressesForGroup(lngEmailGroupID)

				Dim objDocument As New FileStream(ClientDLL.GeneratedFile, FileMode.Open)
				Try
					SendMailWithAttachment(strEmailSubject, objDocument, sEmailAddresses, strEmailAttachAs)
					Response.AppendCookie(New HttpCookie("fileDownloadErrors", "Email sent successfully."))	' Send completion message	
				Catch ex As Exception
					' error generated - return error
					Dim errMessage As String
					If ex.InnerException Is Nothing Then
						errMessage = ""
					Else
						errMessage = ex.InnerException.Message
					End If

					Dim strErrors = String.Format("The following error occured when emailing your document:" _
						& "{0}{0}{1}{0}{0}{2}{0}{0}Please check with your administrator for further details.", "<br/>", _
						ex.Message, errMessage)

					Response.AppendCookie(New HttpCookie("fileDownloadErrors", strErrors))	' marks the download as complete on the client		
				Finally
					Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client		
				End Try

			End If

			' Return the generated file
			If blnSavetoFile Or (Not blnSavetoFile And Not blnEmail) Then
				If IO.File.Exists(ClientDLL.GeneratedFile) Then
					Try
						Dim fileInfo As FileInfo = New FileInfo(ClientDLL.GeneratedFile)
						Response.ContentType = "application/octet-stream"
						Response.Clear()
						Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client
						If Not blnEmail Then Response.AppendCookie(New HttpCookie("fileDownloadErrors", vbNullString)) ' Clear error message response cookie
						Response.AddHeader("Content-Disposition", String.Format("attachment;filename=""{0}""", strDownloadFileName))
						Response.AddHeader("Content-Length", fileInfo.Length.ToString())
						Response.WriteFile(fileInfo.FullName)
						'Response.End()
						Response.Flush()
					Catch ex As Exception
						' error generated - return error
						Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client		
						Response.AppendCookie(New HttpCookie("fileDownloadErrors", ex.Message))	' marks the download as complete on the client		
					Finally
						IO.File.Delete(ClientDLL.GeneratedFile)
					End Try
				Else
					' No file generated - return error
					Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client		
					Response.AppendCookie(New HttpCookie("fileDownloadErrors", "No output file was generated. Check your data."))	' marks the download as complete on the client
				End If
			End If

		End Function

		<HttpPost>
		<ValidateAntiForgeryToken>
		Public Function util_run_customreport_downloadoutput() As FilePathResult

			Dim lngOutputFormat = CType(Request("txtFormat"), OutputFormats)
			Dim bPreview As Boolean = Request("txtPreview")
			Dim blnSavetoFile As Boolean = Request("txtSave")
			Dim blnEmail As Boolean = Request("txtEmail")
			Dim lngEmailGroupID As Integer = Request("txtEmailGroupID")
			Dim strEmailSubject As String = Request("txtEmailSubject")
			Dim strEmailAttachAs As String = Request("txtEmailAttachAs")
			Dim strDownloadFileName As String = Request("txtFilename")
			Dim downloadTokenValue As String = Request("download_token_value_id")

			Dim objReport As Report = Session("CustomReport")
			Dim ClientDLL As New clsOutputRun
			ClientDLL.SessionInfo = CType(Session("SessionContext"), SessionInfo)
			Dim objUser As New clsSettings
			objUser.SessionInfo = CType(Session("SessionContext"), SessionInfo)

			Dim fOK As Boolean
			Dim bBradfordFactor As Boolean

			ClientDLL.ResetColumns()
			ClientDLL.ResetStyles()
			ClientDLL.SaveAsValues = Session("OfficeSaveAsValues").ToString()

			ClientDLL.SettingLocations(CInt(objUser.GetUserSetting("Output", "TitleCol", 3)) _
				, CInt(objUser.GetUserSetting("Output", "TitleRow", 2)) _
				, CInt(objUser.GetUserSetting("Output", "DataCol", 2)) _
				, 4)

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
				, False)

			Dim arrayPageBreakValues
			Dim arrayVisibleColumns
			Dim sEmailAddresses As String = ""
			Dim strDownloadExtension As String = ""

			'Set Options
			If Not bPreview And Not objReport.IsBradfordReport Then
				blnSavetoFile = objReport.OutputSave
				lngOutputFormat = objReport.OutputFormat
				blnEmail = objReport.OutputEmail
				lngEmailGroupID = CLng(objReport.OutputEmailID)
				strEmailSubject = objReport.OutputEmailSubject
				strEmailAttachAs = GetEmailAttachFileNameWithExcelExtension(objReport.OutputEmailAttachAs)
				strDownloadFileName = objReport.DownloadFileName
			End If

			If strDownloadFileName.Length = 0 Then
				objReport.OutputFormat = lngOutputFormat
				objReport.OutputFilename = ""
				strDownloadFileName = objReport.DownloadFileName
			End If

			strDownloadExtension = Path.GetExtension(strDownloadFileName)

			fOK = ClientDLL.SetOptions(False, lngOutputFormat, False, False, "", True, False, False, 0, "", "", strDownloadExtension)

			arrayPageBreakValues = objReport.OutputArray_PageBreakValues
			arrayVisibleColumns = objReport.OutputArray_VisibleColumns

			ClientDLL.SizeColumnsIndependently = True

			If lngOutputFormat = OutputFormats.ExcelGraph Then ClientDLL.SummaryReport = objReport.CustomReportsSummaryReport Or objReport.IsBradfordReport

			Dim sColHeading As String
			Dim iColDataType As Integer
			Dim iColDecimals As Integer
			Dim sBreakValue As String
			Dim blnBreakCheck As Boolean
			Dim bIsCol1000 As Boolean
			Dim lngCol As Integer

			Dim lngDataPageRow As Integer
			Dim lngDataRow As Integer
			Dim iBreakCount As Integer

			ClientDLL.ArrayDim(UBound(arrayVisibleColumns, 2), 0)

			If Not lngOutputFormat = OutputFormats.DataOnly Then

				ClientDLL.HeaderRows = 1
				If ClientDLL.GetFile() = True Then

					If objReport.ReportHasPageBreak Then

						ClientDLL.ArrayDim(UBound(arrayVisibleColumns, 2), 0)
						lngDataRow = 0
						iBreakCount = 0

						For Each objRow As DataRow In objReport.ReportDataTable.Rows

							lngDataRow += 1
							lngDataPageRow += 1

							If CInt(objRow(0)) = RowType.PageBreak And Not blnBreakCheck Then
								sBreakValue = arrayPageBreakValues(iBreakCount)

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
								blnBreakCheck = True
								ClientDLL.ResetColumns()
								ClientDLL.ResetStyles()
								lngDataPageRow = 0

								iBreakCount += 1
								If iBreakCount = arrayPageBreakValues.length Then Exit For
							Else

								blnBreakCheck = False
								lngCol = 0

								ClientDLL.ArrayReDim()

								For lngCount = 0 To UBound(arrayVisibleColumns, 2)
									ClientDLL.ArrayAddTo(lngCol, lngDataPageRow, objRow.Item(lngCount + 1).ToString())
									lngCol += 1
								Next

							End If

						Next

						If objReport.ReportHasSummaryInfo Then
							sBreakValue = "Grand Totals"

							If bBradfordFactor = True Then
								ClientDLL.AddPage(objReport.ReportCaption, "Bradford Factor")
							Else
								ClientDLL.AddPage(objReport.ReportCaption, Replace(sBreakValue, "&&", "&"))
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
						End If


					Else ' no page break

						If (lngOutputFormat = OutputFormats.ExcelGraph And Not objReport.CustomReportsSummaryReport) Or lngOutputFormat = OutputFormats.ExcelPivotTable Then
							Dim trueRowCount As Integer = (From row In objReport.ReportDataTable.AsEnumerable() Where row(0).ToString() = "0" Where String.Join("", row.ItemArray) <> "0").Count()
							ClientDLL.ArrayDim(UBound(arrayVisibleColumns, 2), trueRowCount)
						Else
							' if all columns are hidden then dont generate output
							If (objReport.ReportDataTable.Columns.Count <= 1) Then
								ClientDLL.ArrayDim(UBound(arrayVisibleColumns, 2), 0)
							Else
								ClientDLL.ArrayDim(UBound(arrayVisibleColumns, 2), objReport.ReportDataTable.Rows.Count)
							End If
						End If

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


						lngDataRow = 1
						For Each objRow As DataRow In objReport.ReportDataTable.Rows

							If (lngOutputFormat = OutputFormats.ExcelGraph And Not objReport.CustomReportsSummaryReport) Or lngOutputFormat = OutputFormats.ExcelPivotTable Then
								' Ignore non-data rows.
								If objRow(0).ToString() <> "0" Then Continue For

								' Ignore empty data rows
								If String.Join("", objRow.ItemArray) = "0" Then Continue For
							End If

							For iCountColumns = 1 To objReport.ReportDataTable.Columns.Count - 1
								ClientDLL.ArrayAddTo(iCountColumns - 1, lngDataRow, objRow(iCountColumns).ToString())
							Next

							lngDataRow += 1

						Next

						ClientDLL.DataArray()

					End If

				End If

			End If

			ClientDLL.Complete()

			' Email the generated file
			If blnEmail And lngEmailGroupID > 0 Then
				sEmailAddresses = GetEmailAddressesForGroup(lngEmailGroupID)

				Dim objDocument As New FileStream(ClientDLL.GeneratedFile, FileMode.Open)
				Try
					SendMailWithAttachment(strEmailSubject, objDocument, sEmailAddresses, strEmailAttachAs)
					Response.AppendCookie(New HttpCookie("fileDownloadErrors", "Email sent successfully."))	' Send completion message	
				Catch ex As Exception
					' error generated - return error
					Dim errMessage As String
					If ex.InnerException Is Nothing Then
						errMessage = ""
					Else
						errMessage = ex.InnerException.Message
					End If

					Dim strErrors = String.Format("The following error occured when emailing your document:" _
						& "{0}{0}{1}{0}{0}{2}{0}{0}Please check with your administrator for further details.", "<br/>", _
						ex.Message, errMessage)

					Response.AppendCookie(New HttpCookie("fileDownloadErrors", strErrors))	' marks the download as complete on the client		
				Finally
					Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client		
				End Try
			End If

			' Download the file
			If blnSavetoFile Or (Not blnSavetoFile And Not blnEmail) Then
				If IO.File.Exists(ClientDLL.GeneratedFile) Then
					Try
						Dim fileInfo As FileInfo = New FileInfo(ClientDLL.GeneratedFile)
						Response.ContentType = "application/octet-stream"
						Response.Clear()
						Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client
						If Not blnEmail Then Response.AppendCookie(New HttpCookie("fileDownloadErrors", vbNullString)) ' Clear error message response cookie
						Response.AddHeader("Content-Disposition", String.Format("attachment;filename=""{0}""", strDownloadFileName))
						Response.AddHeader("Content-Length", fileInfo.Length.ToString())
						Response.WriteFile(fileInfo.FullName)
						'Response.End()
						Response.Flush()
					Catch ex As Exception
						' error generated - return error
						Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client		
						Response.AppendCookie(New HttpCookie("fileDownloadErrors", ex.Message))	' marks the download as complete on the client		
					Finally
						IO.File.Delete(ClientDLL.GeneratedFile)
					End Try
				Else
					' No file generated - return error
					Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client		
					Response.AppendCookie(New HttpCookie("fileDownloadErrors", "No output file was generated. Check your data."))	' marks the download as complete on the client		
				End If
			Else
				Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client		
			End If

		End Function

		Public Function util_run_calendarreport_data() As ActionResult
			Return View()
		End Function

		<HttpPost>
		<ValidateAntiForgeryToken>
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
			objCalendarEvent.Reason = objRow.Item("EventDescription1").ToString()
			objCalendarEvent.Duration = objRow.Item("Duration").ToString()
			objCalendarEvent.Description1 = objRow.Item("EventDescription1").ToString()
			objCalendarEvent.Description2 = objRow.Item("EventDescription2").ToString()
			objCalendarEvent.Description1Column = objRow.Item("EventDescription1Column").ToString() + " :"
			If objCalendarEvent.Description1Column.ToString().Substring(1, 1) = ":" Then objCalendarEvent.Description1Column = "<Undefined>"
			objCalendarEvent.Description2Column = objRow.Item("EventDescription2Column").ToString() + " :"
			If objCalendarEvent.Description2Column.ToString().Substring(1, 1) = ":" Then objCalendarEvent.Description2Column = "<Undefined>"
			objCalendarEvent.Region = objRow.Item("Region").ToString()
			'Get region, when region is selected as 'Default' in report
			If String.IsNullOrEmpty(objCalendarEvent.Region) Then
				objCalendarEvent.Region = objCalendar.GetDefaultRegion(objCalendarEvent.BaseID, objCalendarEvent.EndDate)
			End If

			objCalendarEvent.CalendarCode = objRow.Item("Legend").ToString()
			Dim datWorkingPatterns As DataTable = objCalendar.rsCareerChange
			If Not datWorkingPatterns Is Nothing Then
				sSQL = String.Format("BaseID = {0} AND [WP_Date] <= '{1}'", objCalendarEvent.BaseID, objCalendarEvent.StartDate)
				' Get the workingpattern
				objRow = datWorkingPatterns.Select(sSQL, "[WP_Date]").LastOrDefault()
				If Not objRow Is Nothing Then
					objCalendarEvent.WorkingPattern = Trim(objRow.Item("WP_Pattern").ToString())
				End If
			End If

			Return View(objCalendarEvent)

		End Function

		<HttpPost>
		<ValidateAntiForgeryToken>
		Function util_run_calendarreport_download() As FileStreamResult

			Dim objCalendar = CType(Session("objCalendar" & Session("UtilID")), CalendarReport)
			Dim sEmailAddresses As String

			Dim blnSavetoFile As Boolean = Request("txtSave")
			Dim bPreview As Boolean = Request("txtPreview")
			Dim blnEmail As Boolean = Request("txtEmail")
			Dim lngEmailGroupID As Integer = Request("txtEmailGroupID")
			Dim strEmailSubject As String = Request("txtEmailSubject")
			Dim strEmailAttachAs As String = Request("txtEmailAttachAs")
			Dim strDownloadFileName As String = Request("txtFilename")
			Dim downloadTokenValue As String = Request("download_token_value_id")

			Dim objOutput As New CalendarOutput
			objOutput.ReportData = objCalendar.Events
			objOutput.Calendar = objCalendar

			If Not bPreview Then
				blnEmail = objCalendar.OutputEmail
				lngEmailGroupID = objCalendar.OutputEmailID
				strEmailSubject = objCalendar.OutputEmailSubject
				strEmailAttachAs = GetEmailAttachFileNameWithExcelExtension(objCalendar.OutputEmailAttachAs)
				strDownloadFileName = objCalendar.DownloadFileName
			End If

			If strDownloadFileName.Length = 0 Then
				objCalendar.OutputFormat = OutputFormats.ExcelWorksheet
				objCalendar.OutputFilename = ""
				strDownloadFileName = objCalendar.DownloadFileName
			End If

			objOutput.DownloadFileName = strDownloadFileName
			objOutput.Generate(objCalendar.OutputFormat)

			If blnEmail And lngEmailGroupID > 0 Then
				sEmailAddresses = GetEmailAddressesForGroup(lngEmailGroupID)

				Dim objDocument As New FileStream(objOutput.GeneratedFile, FileMode.Open)
				Try
					SendMailWithAttachment(strEmailSubject, objDocument, sEmailAddresses, strEmailAttachAs)
					Response.AppendCookie(New HttpCookie("fileDownloadErrors", "Email sent successfully."))	' Send completion message	
				Catch ex As Exception
					' error generated - return error
					Dim errMessage As String
					If ex.InnerException Is Nothing Then
						errMessage = ""
					Else
						errMessage = ex.InnerException.Message
					End If

					Dim strErrors = String.Format("The following error occured when emailing your document:" _
						& "{0}{0}{1}{0}{0}{2}{0}{0}Please check with your administrator for further details.", "<br/>", _
						ex.Message, errMessage)

					Response.AppendCookie(New HttpCookie("fileDownloadErrors", strErrors))	' marks the download as complete on the client		
				Finally
					Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client		
				End Try

			End If

			objOutput.DownloadFileName = strDownloadFileName
			objOutput.Generate(objCalendar.OutputFormat)

			If blnSavetoFile Or (Not blnSavetoFile And Not blnEmail) Then
				If IO.File.Exists(objOutput.GeneratedFile) Then
					Try
						Dim fileInfo As FileInfo = New FileInfo(objOutput.GeneratedFile)
						Response.ContentType = "application/octet-stream"
						Response.Clear()
						Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client
						If Not blnEmail Then Response.AppendCookie(New HttpCookie("fileDownloadErrors", vbNullString)) ' Clear error message response cookie
						Response.AddHeader("Content-Disposition", String.Format("attachment;filename=""{0}""", strDownloadFileName))
						Response.AddHeader("Content-Length", fileInfo.Length.ToString())
						Response.WriteFile(fileInfo.FullName)
						' Response.End()
						Response.Flush()
					Catch ex As Exception
						' error generated - return error
						Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client		
						Response.AppendCookie(New HttpCookie("fileDownloadErrors", ex.Message))	' marks the download as complete on the client	
					Finally
						IO.File.Delete(objOutput.GeneratedFile)
					End Try
				Else
					' No file generated - return error
					Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client		
					Response.AppendCookie(New HttpCookie("fileDownloadErrors", "No output file was generated. Check your data."))	' marks the download as complete on the client		
				End If
			End If

		End Function

		<HttpPost>
		<ValidateAntiForgeryToken>
		Function util_run_calendarreport_data_submit() As ActionResult
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

			' Set the temp data to store the state of check/uncheck value of ShowWeekends checkbox 
			TempData("IsShowWeekendsChecked") = Request.Form("txtShowWeekends")

			' Go to the requested page.
			Return RedirectToAction("util_run_calendarreport_data")

		End Function

		<ValidateAntiForgeryToken>
		Function util_run_workflow(postData As WorkflowRunModel) As ActionResult
			Return PartialView(postData)
		End Function

		<OutputCache(NoStore:=True, Duration:=0, VaryByParam:="None")>
		Function WorkflowPendingSteps() As ActionResult
			Return PartialView()
		End Function

		Function Progress() As ActionResult
			Return PartialView()
		End Function

#End Region

#Region "Expression Builder"

		<ValidateAntiForgeryToken>
		Function util_def_expression() As ActionResult
			Return PartialView()
		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function util_def_expression_Submit(value As SubmitExpressionModel)

			Dim objExpression As Expression
			Dim iExprType As Integer
			Dim iReturnType As ExpressionValueTypes
			Dim sUtilType As String
			Dim fok As Boolean
			Session("errorMessage") = ""

			' Get the server DLL to save the expression definition
			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			Dim objContext = CType(Session("SessionContext"), SessionInfo)
			If value.txtSend_type = UtilityType.utlFilter Then
				iExprType = ExpressionTypes.giEXPR_RUNTIMEFILTER
				iReturnType = ExpressionValueTypes.giEXPRVALUE_LOGIC
				sUtilType = "Filter"
			Else
				iExprType = value.txtSend_ExpressionType
				iReturnType = value.txtSend_ReturnType
				sUtilType = "Calculation"
			End If

			Try

				objExpression = New Expression(objContext)

				fok = objExpression.Initialise(value.txtSend_tableID, value.txtSend_ID, iExprType, iReturnType)
				If Not fok Then Session("errorMessage") = "<h3>Error saving " & sUtilType.ToLower() & "</h3>Error initialising save definition."

				If fok Then
					fok = objExpression.SetExpressionDefinition(value.txtSend_components1, "", "", "", "", value.txtSend_names)
					If Not fok Then Session("errorMessage") = "<h3>Error saving " & sUtilType.ToLower() & "</h3>Error setting expression definition."
				End If

				If fok Then
					fok = objExpression.SaveExpression(value.txtSend_name, value.txtSend_userName, value.txtSend_access, value.txtSend_description)
					If Not fok Then Session("errorMessage") = "<h3>Error saving " & sUtilType.ToLower() & "</h3>Error saving expression definition."

					If fok Then
						If (value.txtSend_access = "HD") And _
							(value.txtSend_ID > 0) Then
							' Hide any utilities that use this filter/calc.
							' NB. The check to see if we can do this has already been done as part of the filter/calc validation. */

							objDataAccess.ExecuteSP("sp_ASRIntMakeUtilitiesHidden" _
								, New SqlParameter("piUtilityType", SqlDbType.Int) With {.Value = value.txtSend_type} _
								, New SqlParameter("piUtilityID", SqlDbType.Int) With {.Value = value.txtSend_ID})

						End If

						Session("confirmtext") = sUtilType & " has been saved successfully"
						Session("confirmtitle") = sUtilType & "s"
						Session("followpage") = "defsel"
						Session("reaction") = value.txtSend_reaction
						Session("utilid") = objExpression.ExpressionID

					Else

						Session("errorMessage") = "Error saving " & sUtilType.ToLower()

					End If

				End If

			Catch ex As Exception
				Session("errorMessage") = "<h3>Error saving " & sUtilType.ToLower() & "</h3>" & ex.Message
			End Try

			Return RedirectToAction("DefSel")

		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function quickfind_Submit(postData As QuickFindModel)

			Dim lngRecordID As Integer
			Dim sErrorMsg As String

			Session("optionAction") = postData.Action

			If postData.Action = OptionActionType.CANCEL Then
				Session("errorMessage") = ""
				Return RedirectToAction("emptyoption")
			End If

			If postData.Action = OptionActionType.QUICKFIND Then

				Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
				Dim prmResult = New SqlParameter("@plngRecordID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

				Try

					objDataAccess.ExecuteSP("spASRIntGetQuickFindRecord" _
							, New SqlParameter("@plngTableID", SqlDbType.Int) With {.Value = postData.TableID} _
							, New SqlParameter("@plngViewID", SqlDbType.Int) With {.Value = postData.ViewID} _
							, New SqlParameter("@plngColumnID", SqlDbType.Int) With {.Value = postData.ColumnID} _
							, New SqlParameter("@psValue", SqlDbType.VarChar, -1) With {.Value = postData.Value} _
							, New SqlParameter("@psFilterDef", SqlDbType.VarChar, -1) With {.Value = postData.FilterDef} _
							, prmResult _
							, New SqlParameter("@psDecimalSeparator", SqlDbType.VarChar, 100) With {.Value = Session("LocaleDecimalSeparator")} _
							, New SqlParameter("@psLocaleDateFormat", SqlDbType.VarChar, 100) With {.Value = Platform.LocaleDateFormatForSQL()})

					If (CInt(prmResult.Value) = 0) Then
						sErrorMsg = "No records match the criteria."

						If Len(postData.FilterDef) > 0 Then
							sErrorMsg = sErrorMsg & vbCrLf & "Try removing the filter."
						End If
					Else
						' A record has been found !
						lngRecordID = CInt(prmResult.Value)
					End If

				Catch ex As Exception
					sErrorMsg = "Error trying to run 'quick find'." & vbCrLf & ex.Message.RemoveSensitive

				End Try

				Session("errorMessage") = sErrorMsg

				If Len(sErrorMsg) > 0 Then
					' Go to the requested page.
					Return RedirectToAction("Quickfind")
				End If

			End If

			' Go to the requested page.
			Session("optionRecordID") = lngRecordID
			Return RedirectToAction("emptyoption")

		End Function

		Function emptyoption() As ActionResult

			If Len(Session("timestamp")) = 0 Then
				Session("timestamp") = 0
			End If

			Dim m As New EmptyOptionViewModel
			Return View(m)

		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function util_test_expression(value As TestPromptedValuesModel) As ActionResult

			Try

				Dim sKey As String

				Dim aPrompts(1, 0) As String
				Dim j = 0
				ReDim Preserve aPrompts(1, 0)

				If value.PromptValues IsNot Nothing Then
					For Each objPrompt In value.PromptValues
						sKey = objPrompt.Key
						If ((UCase(Left(sKey, 7)) = "PROMPT_") And (Mid(sKey, 8, 1) <> "3")) Or _
								(UCase(Left(sKey, 10)) = "PROMPTCHK_") Then
							ReDim Preserve aPrompts(1, j)

							If (UCase(Left(sKey, 10)) = "PROMPTCHK_") Then
								aPrompts(0, j) = "prompt_3_" & Mid(sKey, 11)
								aPrompts(1, j) = UCase(objPrompt.Value)
							Else
								aPrompts(0, j) = sKey
								Select Case objPrompt.Type
									Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
										' Numeric. Replace locale decimal point with '.'
										aPrompts(1, j) = Replace(objPrompt.Value, CType(Session("LocaleDecimalSeparator"), String), ".")
									Case ExpressionValueTypes.giEXPRVALUE_DATE
										' Date. Reformat to match SQL's mm/dd/yyyy format.
										aPrompts(1, j) = ConvertLocaleDateToSQL(objPrompt.Value)
									Case Else
										aPrompts(1, j) = objPrompt.Value
								End Select
							End If
							j = j + 1
						End If
					Next
				End If

				Session("TestPrompts") = aPrompts

			Catch ex As Exception
				Throw

			End Try

			Return View(value)
		End Function

		Function util_test_expression_pval(value As TestExpressionModel) As ActionResult
			Return View(value)
		End Function

		<HttpPost>
		<ValidateAntiForgeryToken>
		Function util_validate_expression(value As ValidateExpressionModel) As ActionResult
			Return View(value)
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

			If Session("isPortalLogin") = True Then
				Session("SSILinkTableID") = Session("SingleRecordTableID")
				Session("SSILinkViewID") = Session("SingleRecordViewID")
				Session("TopLevelRecID") = 0
			End If

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

					Select Case CInt(prmStartMode.Value)
						Case RecEditStartType.NewRecord
							Session("action") = "NEW"

						Case RecEditStartType.FirstRecord
							Session("action") = "LOAD"

						Case RecEditStartType.FindWindow
							Session("action") = ""

					End Select

				Catch ex As Exception
					sErrorDescription = "Unable to get the link definition." & vbCrLf & ex.Message

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
			If Session("sessionCurrentUser") Is Nothing Then
				Return RedirectToAction("Login", "Account")
			End If
			Return PartialView()
		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function util_def_picklist_submit()
			Try

				Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
				Dim prmID = New SqlParameter("piId", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = CleanNumeric(Request.Form("txtSend_ID"))}

				objDataAccess.ExecuteSP("sp_ASRIntSavePicklist", _
					New SqlParameter("psName", SqlDbType.VarChar, 255) With {.Value = Request.Form("txtSend_name")}, _
					New SqlParameter("psDescription", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_description")}, _
					New SqlParameter("psAccess", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_access")}, _
					New SqlParameter("psUserName", SqlDbType.VarChar, 255) With {.Value = Request.Form("txtSend_userName")}, _
					New SqlParameter("psColumns", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_columns") & ","}, _
					New SqlParameter("psColumns2", SqlDbType.VarChar, -1) With {.Value = Request.Form("txtSend_columns2")}, _
					prmID, _
					New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Request.Form("txtSend_tableID"))})

				Session("confirmtext") = "Picklist has been saved successfully"
				Session("confirmtitle") = "Picklists"
				Session("followpage") = "defsel"
				Session("utilid") = prmID.Value

			Catch ex As Exception

				Response.Write("<h3>Error saving picklist</h3>" & vbCrLf)
				Response.Write(ex.Message & vbCrLf)
				Response.Write("<br/><br/>" & vbCrLf)
				Response.Write("<input type='button' value='Retry' name='GoBack' OnClick='$("".popup"").dialog(""close"");' class='btn' style='width: 80px; float: right;' id='cmdGoBack' />" & vbCrLf)

			End Try

			Return RedirectToAction("DefSel")

		End Function

		<ValidateAntiForgeryToken>
		Function picklistSelectionMain(value As PicklistSelectionModel) As ActionResult
			Return View(value)
		End Function

		<HttpPost>
		<ValidateAntiForgeryToken>
		Function picklistSelectionData_Submit(value As FormCollection) As ActionResult

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
			Return View("picklistSelectionData")

		End Function

      <HttpPost>
      <ValidateAntiForgeryToken>
		Function util_validate_picklist(model As ValidatePicklistModel) As ActionResult
			Return View(model)
		End Function

#End Region

		Function Quickfind() As ActionResult
			Return View()
		End Function

		Function Filterselect() As ActionResult
			Return View()
		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function filterselect_Submit(postData As FilterSelectModel) As RedirectToRouteResult
			Session("optionScreenID") = postData.ScreenID
			Session("optionTableID") = postData.TableID
			Session("optionViewID") = postData.ViewID
			Session("optionFilterDef") = postData.FilterDef
			Session("optionFilterSQL") = postData.FilterSQL
			Session("optionAction") = postData.Action
			Session("filterDef_" & postData.TableID) = postData.FilterDef
			Session("filterSQL_" & postData.TableID) = postData.FilterSQL

			Session("OptionSelectedRecordIds") = postData.SelectedRecordsInFindGrid

			'If (postData.FilterDef = "" AndAlso postData.FilterSQL = "") Then
			'	Session("OptionSelectedRecordIds") = postData.SelectedRecordsInFindGrid
			'End If

			Return RedirectToAction("emptyoption")
		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Sub filter_clear(tableID As Integer)
			Session("filterDef_" & tableID) = ""
			Session("filterSQL_" & tableID) = ""
			Session("OptionSelectedRecordIds") = ""
		End Sub

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function tbAddFromWaitingListFind_Submit(postData As DelegateBookingModel)

			Session("optionAction") = postData.Action

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			Dim sErrorMsg As String = ""
			Dim sTBResultCode As String = "000"	'Validation OK
			Dim sCourseOverbooked As String = ""
			Dim sPreReqFails As String = ""

			Session("optionLinkRecordID") = postData.Key2

			If postData.Action = OptionActionType.SELECTADDFROMWAITINGLIST_1 Then
				If postData.Key1 > 0 Then
					' First pass after selecting the employee to book.
					' Get the user to choose whether to make the booking 'provisional'
					' or confirmed.
					If Session("TB_TBStatusPExists") Then
						Return RedirectToAction("tbStatusPrompt")
					Else
						Session("optionAction") = OptionActionType.SELECTADDFROMWAITINGLIST_2
						Session("optionLookupValue") = "B"
					End If
				End If
			End If

			If postData.Action = OptionActionType.SELECTADDFROMWAITINGLIST_2 Then
				If postData.Key1 > 0 Then
					If Len(sErrorMsg) = 0 Then
						' Validate the booking.					
						Try

							Dim prmResult = New SqlParameter("@piResultCode", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
							Dim prmCourseOverbooked = New SqlParameter("@psCourseOverbooked", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

							objDataAccess.ExecuteSP("sp_ASRIntValidateTrainingBooking" _
								, prmResult _
								, New SqlParameter("piEmpRecID", SqlDbType.Int) With {.Value = postData.Key2} _
								, New SqlParameter("piCourseRecID", SqlDbType.Int) With {.Value = postData.Key1} _
								, New SqlParameter("psBookingStatus", SqlDbType.VarChar, -1) With {.Value = postData.BookingStatus} _
								, New SqlParameter("piTBRecID", SqlDbType.Int) With {.Value = 0} _
								, prmCourseOverbooked)

							sTBResultCode = prmResult.Value
							sCourseOverbooked = prmCourseOverbooked.Value
						Catch ex As Exception
							sErrorMsg = "Error validating training booking." & vbCrLf & ex.Message.RemoveSensitive()
						End Try
					End If
				End If
			End If

			' Go to the requested page.
			Session("TBResultCode") = sTBResultCode
			Session("errorMessage") = sErrorMsg
			Session("PreReqFails") = sPreReqFails	' This will be a sp output in the future along the lines of Bulkbooking
			Session("Overbooked") = sCourseOverbooked
			Session("optionLookupValue") = postData.BookingStatus

			' Go to the requested page.
			Return RedirectToAction("emptyoption")

		End Function

		Function tbStatusPrompt() As ActionResult
			Return View()
		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function tbBookCourseFind_Submit(postData As DelegateBookingModel)

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
			Dim sErrorMsg As String = ""
			Dim sTBResultCode As String = "000"	'Validation OK
			Dim sCourseOverbooked As String = ""

			Session("optionAction") = postData.Action
			Session("optionLinkRecordID") = postData.Key1
			Session("optionRecordID") = postData.Key2

			If postData.Action = OptionActionType.SELECTBOOKCOURSE_1 Then
				If postData.Key2 > 0 Then
					' First pass after selecting the course to book.
					' Get the user to choose whether to make the booking 'provisional'
					' or confirmed.
					If Session("TB_TBStatusPExists") Then
						Return RedirectToAction("tbStatusPrompt")
					Else
						Session("optionAction") = OptionActionType.SELECTBOOKCOURSE_2
						Session("optionLookupValue") = "B"
					End If
				End If
			End If

			If postData.Action = OptionActionType.SELECTBOOKCOURSE_2 Then
				If postData.Key2 > 0 Then
					' Get the employee record ID from the given Waiting List record.
					Dim iEmpRecID = 0

					Try

						Dim prmTBEmployeeRecordID = New SqlParameter("piEmpRecordID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
						objDataAccess.ExecuteSP("sp_ASRIntGetEmpIDFromWLID", _
								prmTBEmployeeRecordID, _
								New SqlParameter("@piWLRecordID", SqlDbType.Int) With {.Value = postData.Key2})

						iEmpRecID = CInt(prmTBEmployeeRecordID.Value)

						If iEmpRecID = 0 Then
							sErrorMsg = "Error getting employee ID."
						End If

					Catch ex As Exception
						sErrorMsg = "Error getting employee ID." & vbCrLf & ex.Message.RemoveSensitive

					End Try

					If Len(sErrorMsg) = 0 Then
						' Validate the booking.
						Try
							Dim prmResult = New SqlParameter("@piResultCode", SqlDbType.VarChar) With {.Direction = ParameterDirection.Output, .Size = 100}
							Dim prmCourseOverbooked = New SqlParameter("@psCourseOverbooked", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

							objDataAccess.ExecuteSP("sp_ASRIntValidateTrainingBooking" _
								, prmResult _
								, New SqlParameter("piEmpRecID", SqlDbType.Int) With {.Value = iEmpRecID} _
								, New SqlParameter("piCourseRecID", SqlDbType.Int) With {.Value = postData.Key1} _
								, New SqlParameter("psBookingStatus", SqlDbType.VarChar, -1) With {.Value = postData.BookingStatus} _
								, New SqlParameter("piTBRecID", SqlDbType.Int) With {.Value = 0} _
								, prmCourseOverbooked)

							sTBResultCode = prmResult.Value
							sCourseOverbooked = prmCourseOverbooked.Value
							Session("optionLinkRecordID") = iEmpRecID
							Session("optionLookupValue") = postData.BookingStatus
							Session("TBCourseID") = postData.Key1

						Catch ex As Exception
							sErrorMsg = "Error validating training booking." & vbCrLf & ex.Message.RemoveSensitive
						End Try
					End If
				End If
			End If

			' Go to the requested page.
			Session("TBResultCode") = sTBResultCode
			Session("errorMessage") = sErrorMsg
			Session("Overbooked") = sCourseOverbooked

			Return RedirectToAction("emptyoption")

		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function BulkBooking_Submit(postData As DelegateBookingModel)

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			Dim sErrorMsg As String = ""
			Dim sPreReqFailsCount As String = ""
			Dim sUnAvailFailsCount As String = ""
			Dim sOverlapFailsCount As String = ""
			Dim sCourseOverbooked As String = ""
			Dim sTBResults As String = ""

			' Read the information from the calling form.
			Session("optionAction") = postData.Action
			Session("optionRecordID") = postData.Key1
			Session("optionLinkRecordID") = postData.Key2
			Session("optionLookupValue") = postData.BookingStatus

			If postData.Action = OptionActionType.SELECTBULKBOOKINGS Then
				If Len(postData.Key2) > 0 Then

					Try
						Dim prmErrorMsg = New SqlParameter("psErrorMessage", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
						Dim prmPreRequisitesFailsCount = New SqlParameter("psPreReqCheckFailsCount", SqlDbType.Int, -1) With {.Direction = ParameterDirection.Output}
						Dim prmAvailabilityFailsCount = New SqlParameter("psUnavailabilityCheckFailCount", SqlDbType.Int, -1) With {.Direction = ParameterDirection.Output}
						Dim prmOverLappingFailsCount = New SqlParameter("psOverlapCheckFailCount", SqlDbType.Int, -1) With {.Direction = ParameterDirection.Output}
						Dim prmCourseOverbooked = New SqlParameter("psCourseOverbooked", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

						Dim dt As DataTable = objDataAccess.GetDataTable("sp_ASRIntValidateBulkBookings", CommandType.StoredProcedure _
							, New SqlParameter("piCourseRecordID", SqlDbType.Int) With {.Value = postData.Key1} _
							, New SqlParameter("psEmployeeRecordIDs", SqlDbType.VarChar, -1) With {.Value = postData.Key2} _
							, New SqlParameter("psBookingStatus", SqlDbType.VarChar, -1) With {.Value = postData.BookingStatus} _
							, prmErrorMsg _
							, prmPreRequisitesFailsCount _
							, prmAvailabilityFailsCount _
							, prmOverLappingFailsCount _
							, prmCourseOverbooked)

						sPreReqFailsCount = prmPreRequisitesFailsCount.Value.ToString()
						sUnAvailFailsCount = prmAvailabilityFailsCount.Value.ToString()
						sOverlapFailsCount = prmOverLappingFailsCount.Value.ToString()
						sCourseOverbooked = prmCourseOverbooked.Value.ToString

						'Loop over the returned records
						For Each r As DataRow In dt.Rows
							If r("ResultCode").ToString <> "000" Then	'Ignore records that have passed all checks
								sTBResults = String.Concat(sTBResults, r("EmployeeName"), "\", r("ResultCode"), "|") 'The format is EmployeeName\ResultCode|EmployeeName\ResultCode...
							End If
						Next
						sTBResults = sTBResults.TrimEnd("|") 'Trim the last separator
					Catch ex As Exception
						sErrorMsg = "Error validating training booking transfers." & vbCrLf & ex.Message.RemoveSensitive()
					End Try

				End If
			End If

			' Go to the requested page.
			Session("TBResultCode") = sTBResults
			Session("errorMessage") = sErrorMsg
			Session("PreReqFails") = sPreReqFailsCount
			Session("UnAvailFails") = sUnAvailFailsCount
			Session("OverlapFails") = sOverlapFailsCount
			Session("Overbooked") = sCourseOverbooked

			Return RedirectToAction("emptyoption")

		End Function

		Public Function BulkBookingSelection() As ActionResult
			Dim m As New BulkBookingSelectionViewModel
			Return PartialView("BulkBookingSelection", m)
		End Function

		Public Function BulkBookingSelectionData(tableID As String, viewID As String, orderID As String, pageAction As String) As JsonResult

			Dim sErrorDescription = ""

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			Dim sThousandColumns As String = ""
			Dim sBlankIfZeroColumns As String = ""

			Try
				Get1000SeparatorBlankIfZeroFindColumns(CleanNumeric(tableID), CleanNumeric(viewID), CleanNumeric(orderID), sThousandColumns, sBlankIfZeroColumns)

				Dim prmError = New SqlParameter("pfError", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmIsFirstPage = New SqlParameter("pfFirstPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmIsLastPage = New SqlParameter("pfLastPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmColumnType = New SqlParameter("piColumnType", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmTotalRecCount = New SqlParameter("piTotalRecCount", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmFirstRecPos = New SqlParameter("piFirstRecPos", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = CleanNumeric(Session("firstRecPos"))}
				Dim prmColumnSize = New SqlParameter("piColumnSize", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmColumnDecimals = New SqlParameter("piColumnDecimals", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

				Dim dsFindData = objDataAccess.GetDataSet("sp_ASRIntGetLinkFindRecords" _
					, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(tableID)} _
					, New SqlParameter("piViewID", SqlDbType.Int) With {.Value = CleanNumeric(viewID)} _
					, New SqlParameter("piOrderID", SqlDbType.Int) With {.Value = CleanNumeric(orderID)} _
					, prmError _
					, New SqlParameter("piRecordsRequired", SqlDbType.Int) With {.Value = 20000} _
					, prmIsFirstPage _
					, prmIsLastPage _
					, New SqlParameter("psLocateValue", SqlDbType.VarChar, -1) With {.Value = Session("locateValue")} _
					, prmColumnType _
					, New SqlParameter("psAction", SqlDbType.VarChar, 100) With {.Value = pageAction} _
					, prmTotalRecCount _
					, prmFirstRecPos _
					, New SqlParameter("piCurrentRecCount", SqlDbType.Int) With {.Value = CleanNumeric(Session("currentRecCount"))} _
					, New SqlParameter("psExcludedIDs", SqlDbType.VarChar, -1) With {.Value = ""} _
					, prmColumnSize _
					, prmColumnDecimals)

            Dim rstFindRecords As New DataTable
            Dim rstColumnProperties As New DataTable

            If dsFindData.Tables.Count > 0 Then
               rstFindRecords = dsFindData.Tables(0)
               rstColumnProperties = dsFindData.Tables(1)
            End If

				If prmError.Value <> 0 Then
					Session("ErrorTitle") = "Bulk Booking Selection Find Page"
					Session("ErrorText") = "Error reading employee records definition."
					Response.Clear()
					Response.Redirect("error")
				End If

				Dim jqGridColDef = New Dictionary(Of String, String)
				Dim rows As New List(Of Dictionary(Of String, Object))()
				Dim row As Dictionary(Of String, Object)
				Dim iLoop As Integer = 0

				For Each dr As DataRow In rstFindRecords.Rows
					iLoop += 1
					row = New Dictionary(Of String, Object)()
               For Each col As DataColumn In rstFindRecords.Columns

                  Dim objRow = rstColumnProperties.Select("ColumnName='" & col.ColumnName & "'")

                  If Not jqGridColDef.ContainsKey(col.ColumnName) Then
                     Dim sColDef As String = col.DataType.Name

                     If sColDef = "Decimal" Then
                        Dim numberAsString As String = dr(col).ToString()
                        Dim indexOfDecimalPoint As Integer = numberAsString.IndexOf(LocaleDecimalSeparator(), System.StringComparison.Ordinal)
                        Dim numberOfDecimals As Integer = 0
                        If indexOfDecimalPoint > 0 Then numberOfDecimals = numberAsString.Substring(indexOfDecimalPoint + 1).Length

                        If Mid(sThousandColumns, iLoop + 1, 1) = "1" Then
                           sColDef &= vbTab & numberOfDecimals.ToString() & vbTab & "true"
                        Else
                           sColDef &= vbTab & numberOfDecimals.ToString() & vbTab & "false"
                        End If
                     End If

                     jqGridColDef.Add(col.ColumnName, sColDef)
                  End If

                  If col.DataType.Name = "DateTime" And dr(col).ToString().Length > 0 Then
                     row.Add(col.ColumnName, dr(col).ToShortDateString())
                  Else
                     If objRow.Length > 0 AndAlso objRow.FirstOrDefault.Item("BlankIfZero") AndAlso dr(col).ToString() = "0.00" Then
                        row.Add(col.ColumnName, "")
                     Else
                        row.Add(col.ColumnName, dr(col))
                     End If
                  End If

               Next

               rows.Add(row)

				Next

				Dim results = New With {.total = 1, .page = 1, .records = 0, .rows = rows, .colDef = jqGridColDef}
				Return Json(results, JsonRequestBehavior.AllowGet)

			Catch ex As Exception
				sErrorDescription = "The find records could not be retrieved." & vbCrLf & FormatError(ex.Message)
			End Try

		End Function

		<HttpPost>
		<ValidateAntiForgeryToken>
		Public Function util_run_mailmerge_completed() As FileStreamResult

			Dim downloadTokenValue As String = Request("download_token_value_id")

			Try

				Dim objMergeDocument As Code.MailMergeRun = Session("MailMerge_CompletedDocument")

				Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client		
				Response.AppendCookie(New HttpCookie("fileDownloadErrors", "Mail merge completed successfully."))	' Send completion message

        Dim filename As String = objMergeDocument.OutputFileName
        For Each c In IO.Path.GetInvalidFileNameChars
          filename = filename.Replace(c, "")
        Next

        Return File(objMergeDocument.MergeDocument, "application/vnd.openxmlformats-officedocument.wordprocessingml.document" _
          , Path.GetFileName(filename))

      Catch ex As Exception
        ' error generated - return error
				Response.AppendCookie(New HttpCookie("fileDownloadErrors", ex.Message))	' marks the download as complete on the client		
        Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client		

      Finally
				Response.AppendCookie(New HttpCookie("fileDownloadToken", downloadTokenValue)) ' marks the download as complete on the client	
			End Try

    End Function

    <HttpPost()>
    <ValidateAntiForgeryToken>
		Function tbTransferBookingFind_Submit(form As DelegateBookingModel)

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			Dim sErrorMsg As String = ""
			Dim sTBResultCode As String = "000"	'Validation OK
			Dim sCourseOverbooked As String = ""

			Session("optionAction") = form.Action
			Session("optionLinkRecordID") = form.Key2

			If form.Action = OptionActionType.SELECTTRANSFERBOOKING_1 Then
				If form.Key1 > 0 Then
					' Get the employee record ID from the given Training Booking record.
					Dim iEmpRecID As Integer = 0
					Try
						Dim prmEmployeeRecordID = New SqlParameter("piEmpRecordID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

						objDataAccess.ExecuteSP("sp_ASRIntGetEmpIDFromTBID" _
							, prmEmployeeRecordID _
							, New SqlParameter("piTBRecordID", SqlDbType.Int) With {.Value = form.Key1})

						iEmpRecID = prmEmployeeRecordID.Value

						If iEmpRecID = 0 Then
							sErrorMsg = "Error getting employee ID."
						End If

					Catch ex As Exception
						sErrorMsg = "Error getting employee ID." & vbCrLf & ex.Message.RemoveSensitive()

					End Try

					If Len(sErrorMsg) = 0 Then
						' Validate the booking.
						Try
							Dim prmResult = New SqlParameter("@piResultCode", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
							Dim prmCourseOverbooked = New SqlParameter("@psCourseOverbooked", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

							objDataAccess.ExecuteSP("sp_ASRIntValidateTrainingBooking" _
								, prmResult _
								, New SqlParameter("piEmpRecID", SqlDbType.Int) With {.Value = iEmpRecID} _
								, New SqlParameter("piCourseRecID", SqlDbType.Int) With {.Value = form.Key2} _
								, New SqlParameter("psBookingStatus", SqlDbType.VarChar, -1) With {.Value = form.BookingStatus} _
								, New SqlParameter("piTBRecID", SqlDbType.Int) With {.Value = 0} _
								, prmCourseOverbooked)

							sTBResultCode = prmResult.Value
							sCourseOverbooked = prmCourseOverbooked.Value
						Catch ex As Exception
							sErrorMsg = "Error validating training booking." & vbCrLf & ex.Message.RemoveSensitive()
						End Try
					End If
				End If
			End If

			' Go to the requested page.
			Session("TBResultCode") = sTBResultCode
			Session("errorMessage") = sErrorMsg
			Session("Overbooked") = sCourseOverbooked

			Return RedirectToAction("emptyoption")
		End Function

		Function OptionDataGrid(GotoOptionPage As String) As PartialViewResult
			Dim m As New OptionDataGridViewModel(GotoOptionPage)
			Return PartialView("OptionDataGrid", m)
		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
	 Function tbTransferCourseFind_Submit(form As GotoOptionDataModel)

			Dim sErrorMsg = ""
			Dim iTBResultCode = 0

			emptyoption_Submit_BASE(form)

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			If form.txtGotoOptionAction = OptionActionType.SELECTTRANSFERCOURSE Then

				If Session("optionLinkRecordID") > 0 Then

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
				Return RedirectToAction("emptyoption")
			End If

		End Function

		Function orderselect() As ActionResult
			Return View()
		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function orderselect_Submit(postData As SelectOrderModel)
			If postData.Action = OptionActionType.CANCEL Then
				Session("errorMessage") = ""

			ElseIf postData.Action = OptionActionType.SELECTORDER Then

				Session("optionScreenID") = postData.ScreenID
				Session("optionTableID") = postData.TableID
				Session("optionViewID") = postData.ViewID
				Session("optionAction") = postData.Action

				' Do we need both session variables set?
				Session("optionOrderID") = postData.OrderID
				Session("orderID") = postData.OrderID

				Try
					Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
					Dim prmFromDef = New SqlParameter("psFromDef", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

					objDataAccess.ExecuteSP("sp_ASRIntGetOrderSQL" _
							, New SqlParameter("piScreenID", SqlDbType.Int) With {.Value = postData.ScreenID} _
							, New SqlParameter("piViewID", SqlDbType.Int) With {.Value = postData.ViewID} _
							, New SqlParameter("piOrderID", SqlDbType.Int) With {.Value = postData.OrderID} _
							, prmFromDef)

					Session("fromDef") = prmFromDef.Value

				Catch ex As Exception
					Session("errorMessage") = "Error retrieving the new order definition." & vbCrLf & ex.Message.RemoveSensitive

				End Try

			End If

			Return RedirectToAction("emptyoption")

		End Function

		Function lookupFind() As ActionResult
			Return View()
		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function lookupFind_Submit(postData As LookupFindModel) As RedirectToRouteResult
			Session("optionAction") = postData.Action
			Session("optionColumnID") = postData.ColumnID
			Session("optionLookupColumnID") = postData.LookupColumnID
			Session("optionLookupValue") = postData.LookupValue
			Return RedirectToAction("emptyoption")
		End Function

		Function themeEditor() As PartialViewResult
			Return PartialView()
		End Function

		Function linkFind() As ActionResult
			Return View()
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
		<ValidateAntiForgeryToken>
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

			Return RedirectToAction("Main", "Home")

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

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Function oleFind_Submit(form As FormCollection) As ActionResult

			Dim filesize As Integer = 0
			Dim buffer As Byte()
			Dim iOLEType As OLEType

			Dim sErrorMsg = ""
			' Read the information from the calling form.
			Dim sAction = CType(Request.Form("txtGotoOptionAction"), OptionActionType)

			iOLEType = CType(Request.Form("txtOLEType"), OLEType)

			' Check that any uploaded files have a valid file extension
			If Not IsValidFileExtension(Request.Form("txtOLEJustFileName")) Then
				Return New HttpStatusCodeResult(400, "OpenHR Web cannot upload this type of file.")
			End If

			' Commit changes to the database		
			If sAction = OptionActionType.LINKOLE Then

				If Request.Files.Count > 0 Then
					Dim filSelectFile = Request.Files(0)

					If filSelectFile IsNot Nothing Then
						filesize = filSelectFile.InputStream.Length
						buffer = New Byte(filSelectFile.InputStream.Length - 1) {}
						Dim offset As Integer = 0
						Dim cnt As Integer = 0

						While (InlineAssignHelper(cnt, filSelectFile.InputStream.Read(buffer, offset, 10))) > 0
							offset += cnt
						End While

						' If this is an image type, check it's contents really are an image.
						If Path.GetExtension(filSelectFile.FileName).IsPictureExtension() Then
							If Not IsValidImageFromStream(filSelectFile.InputStream) Then
								Return New HttpStatusCodeResult(400, "Not a valid image file.")
							End If
						End If

					End If

				End If

				' The file will (should) have already been copied from the client to the temp path
				Dim objOLE As Ole = Session("OLEObject")

				Dim safeFileName As String = Request.Form("txtOLEFile")	' contains file path in IE

				If safeFileName.Length > 0 Then
					' Remove file path info.
					safeFileName = Path.GetFileName(safeFileName)

					' Strip out all unicode characters/special characters to meet OWASP requirement
					safeFileName = String.Join("_", safeFileName.Split(Path.GetInvalidFileNameChars()))

					' Ensure only one dot and alpha-numeric characters.
					' Force filename and extension to both be present.
					Dim regex As New Regex("[\w]{1,200}\.[\w]{1,10}", DirectCast(0, RegexOptions))	' taken from OWASP website.
					If Not regex.IsMatch(safeFileName) Then
						Return New HttpStatusCodeResult(400, "Your selected filename contains invalid characters and cannot be uploaded.")
					End If
				End If


				With objOLE
					.OLEType = CType(Request.Form("txtOLEType"), OLEType)
					.FileName = safeFileName
					.DisplayFilename = safeFileName
					.OLEFileSize = filesize
					Dim oleErrorResponse As String = .SaveStream(Session("optionRecordID"), Session("optionColumnID"), buffer)

					If oleErrorResponse.Length > 0 Then
						oleErrorResponse = Server.HtmlEncode("Unable to embed file:" & vbCrLf & oleErrorResponse)
					End If
					Session("errorMessage") = oleErrorResponse

					If .OLEType = OLEType.Embedded Then
						Session("optionFileValue") = .ExtractPhotoToBase64(Session("optionRecordID"), Session("optionColumnID"), Session("realSource"))
					Else
						Session("optionFileValue") = .FileName
					End If

				End With
				Session("OLEObject") = objOLE
				objOLE = Nothing

				Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
				Session("timestamp") = objDatabase.GetRecordTimestamp(CleanNumeric(Session("optionRecordID")), Session("realSource"))

				'Update the ID badge picture in Session only if the user that is being edited is the same as the logged in user and we are embeding a photo
				If CInt(Session("PreviousRecordID")) = CInt(Session("LoggedInUserRecordID")) And Session("optionIsPhoto") = True Then
					If Session("optionFileValue") = "" Then
						Session("SelfServicePhotograph_Src") = Url.Content("~/Content/images/anonymous.png")
					Else
						Session("SelfServicePhotograph_Src") = "data:image/jpeg;base64," & Session("optionFileValue")
					End If
				End If
			End If

			Session("optionScreenID") = Request.Form("txtGotoOptionScreenID")
			Session("optionRecordID") = Request.Form("txtGotoOptionRecordID")
			Session("optionFilterDef") = Request.Form("txtGotoOptionFilterDef")
			Session("optionFilterSQL") = Request.Form("txtGotoOptionFilterSQL")
			Session("optionValue") = Request.Form("txtGotoOptionValue")
			Session("optionLinkTableID") = Request.Form("txtGotoOptionLinkTableID")
			Session("optionLinkOrderID") = Request.Form("txtGotoOptionLinkOrderID")
			Session("optionLinkViewID") = Request.Form("txtGotoOptionLinkViewID")
			Session("optionLinkRecordID") = Request.Form("txtGotoOptionLinkRecordID")
			Session("optionLookupColumnID") = Request.Form("txtGotoOptionLookupColumnID")
			Session("optionLookupMandatory") = Request.Form("txtGotoOptionLookupMandatory")
			Session("optionLookupValue") = Request.Form("txtGotoOptionLookupValue")
			Session("optionFile") = Request.Form("txtGotoOptionFile")
			Session("optionAction") = sAction
			Session("optionPageAction") = Request.Form("txtGotoOptionPageAction")
			Session("optionCourseTitle") = Request.Form("txtGotoOptionCourseTitle")
			Session("optionFirstRecPos") = Request.Form("txtGotoOptionFirstRecPos")
			Session("optionCurrentRecCount") = Request.Form("txtGotoOptionCurrentRecCount")
			Session("optionOLEType") = Request.Form("txtGotoOptionOLEType")
			Session("optionOLEMaxEmbedSize") = Request.Form("txtGotoOptionOLEMaxEmbedSize")

			If sAction = OptionActionType.CANCEL Then
				Session("errorMessage") = sErrorMsg
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

		'<ValidateAntiForgeryToken>
		'<HttpPost>
		'Public Function Upload(filSelectFile As HttpPostedFileBase) As ActionResult
		'	Const path As String = "D:\Temp\"

		'	If filSelectFile IsNot Nothing Then
		'		filSelectFile.SaveAs(path & Convert.ToString(filSelectFile.FileName))
		'	End If

		'	'Return RedirectToAction("Index")
		'End Function

		<ValidateAntiForgeryToken>
		<HttpPost> _
	 Public Function AjaxFileUpload(form As FormCollection) As String

			If Request.Files.Count > 0 Then
				For Each file As String In Request.Files

					' The file will (should) have already been copied from the client to the temp path
					Dim objOLE As Ole = Session("OLEObject")
					Session("errorMessage") = ""
					Dim fileContent = Request.Files(file)
					Dim columnID = form.Item("columnID")
					Dim recordID = form.Item("recordID")
					Dim fOk As Boolean = True

					' OWASP checks on filename
					Dim safeFileName As String = fileContent.FileName	' contains file path in IE

					' Remove file path info.
					safeFileName = Path.GetFileName(safeFileName)

					' Strip out all unicode characters/special characters to meet OWASP requirement
					safeFileName = String.Join("_", safeFileName.Split(Path.GetInvalidFileNameChars()))

					' Ensure only one dot and alpha-numeric characters.
					' Force filename and extension to both be present.
					Dim regex As New Regex("[\w]{1,200}\.[\w]{1,10}", DirectCast(0, RegexOptions))	' taken from OWASP website.
					If Not regex.IsMatch(safeFileName) Then
						Session("errorMessage") = "Your selected filename contains invalid characters and cannot be uploaded."
						Return Session("errorMessage").ToString()
					Else
						If (fileContent IsNot Nothing) AndAlso fileContent.ContentLength > 0 Then

							' If this is an image type, check it's contents really are an image.
							If Path.GetExtension(fileContent.FileName).IsPictureExtension() Then
								If Not IsValidImageFromStream(fileContent.InputStream) Then
									Session("errorMessage") = "Not a valid image file."
									Return Session("errorMessage").ToString()
								End If
							End If

							Dim buffer As Byte()
							buffer = New Byte(fileContent.InputStream.Length - 1) {}
							Dim offset As Integer = 0
							Dim cnt As Integer = 0
							While (InlineAssignHelper(cnt, fileContent.InputStream.Read(buffer, offset, 10))) > 0
								offset += cnt
							End While

							With objOLE
								.OLEType = OLEType.Embedded
								.FileName = fileContent.FileName
								.DisplayFilename = Path.GetFileName(fileContent.FileName)
								.OLEFileSize = fileContent.ContentLength.ToString()

								Dim oleErrorResponse As String = .SaveStream(recordID, columnID, buffer)

								If oleErrorResponse.Length > 0 Then
									oleErrorResponse = Server.HtmlEncode("Unable to embed file:" & vbCrLf & oleErrorResponse)
								End If
								Session("errorMessage") = oleErrorResponse


							End With
							Session("OLEObject") = objOLE
							objOLE = Nothing
						End If


					End If

				Next
			Else
				' deleting
				' The file will (should) have already been copied from the client to the temp path
				Dim objOLE As Ole = Session("OLEObject")
				Session("errorMessage") = ""
				Dim fileContent = Nothing
				Dim columnID = form.Item("columnID")
				Dim recordID = form.Item("recordID")

				Dim buffer As Byte()

				With objOLE
					.OLEType = OLEType.Embedded
					.FileName = ""
					.DisplayFilename = ""
					.OLEFileSize = 0

					Dim oleErrorResponse As String = .SaveStream(recordID, columnID, buffer)

					If oleErrorResponse.Length > 0 Then
						oleErrorResponse = Server.HtmlEncode("Unable to embed file:" & vbCrLf & oleErrorResponse)
					End If
					Session("errorMessage") = oleErrorResponse


				End With
				Session("OLEObject") = objOLE
				objOLE = Nothing


			End If


			Return Session("errorMessage").ToString()

		End Function

		Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
			target = value
			Return value
		End Function

		Public Function DownloadFile(filename As String, serverpath As String) As ActionResult

			If filename.Length > 0 And serverpath.Length > 0 Then

				If serverpath.Substring(serverpath.Length - 1) <> "\" Then serverpath &= "\"

				' TODO: add the file path!
				Dim fullpath = serverpath & filename
				Dim fileInfo As FileInfo = New FileInfo(fullpath)
				Response.ContentType = "application/octet-stream"
				Response.AddHeader("Content-Disposition", String.Format("attachment;filename=""{0}""", filename))
				Response.AddHeader("Content-Length", fileInfo.Length.ToString())
				Response.WriteFile(fileInfo.FullName)
				Response.End()
				Response.Flush()
			End If

		End Function

		Public Function EditFile(plngRecordID As Integer, plngColumnID As Integer, pstrRealSource As String)

			Dim objOLE As Ole = Session("OLEObject")
			Dim fileResponse As Byte() = objOLE.CreateOLEDocument(plngRecordID, plngColumnID, pstrRealSource)
			Dim filename = Server.UrlDecode(objOLE.DisplayFilename)

			Response.ContentType = "application/octet-stream"
			Response.AppendHeader("Content-Disposition", String.Format("attachment;filename=""{0}""", filename))

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
		<ValidateAntiForgeryToken>
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

		Public Function stdrpt_run_AbsenceBreakdown() As ActionResult
			Return View()
		End Function

		''' <summary>
		''' Returns the absence breakdown report configuration view
		''' </summary>
		Public Function AbsenceBreakdownConfiguration() As ActionResult

			Dim strReportType = "AbsenceBreakdown"

			Dim objModel As StandardReportsConfigurationModel = New StandardReportsConfigurationModel()

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
			Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)

			' Settings objects
			Dim standardReportSettings As New HR.Intranet.Server.StandardReport
			standardReportSettings.SessionInfo = CType(Session("SessionContext"), SessionInfo)

			' Set parent table id
			objModel.TableId = SettingsConfig.Personnel_EmpTableID

			Dim absenceTypeTable = objDataAccess.GetDataTable("sp_ASRIntGetAbsenceTypes", CommandType.StoredProcedure)
			Dim absenceType As New AbsenceType()

			' Fill Absence types
			For Each objRow As DataRow In absenceTypeTable.Rows
				absenceType = New AbsenceType()
				absenceType.Type = objRow(0).ToString()
				absenceType.IsSelected = objDatabase.GetSystemSetting(strReportType, "Absence Type " & objRow(0).ToString(), "0")
				objModel.AbsenceTypes.Add(absenceType)
			Next

			' Is custom date
			objModel.IsCustomDate = objDatabase.GetSystemSetting(strReportType, "Custom Dates", "0")

			If objModel.IsCustomDate Then
				Dim strRecSelStatus As String
				Dim customDateId As Long

				customDateId = objDatabase.GetSystemSetting(strReportType, "Start Date", "0")
				strRecSelStatus = standardReportSettings.IsCalculationValid(customDateId)
				If (strRecSelStatus <> vbNullString) Then
					objModel.StartDate = "None"
					objModel.StartDateId = 0
				Else
					objModel.StartDate = standardReportSettings.GetFilterName(customDateId)
					objModel.StartDateId = customDateId
				End If

				customDateId = objDatabase.GetSystemSetting(strReportType, "End Date", "0")
				strRecSelStatus = standardReportSettings.IsCalculationValid(customDateId)
				If (strRecSelStatus <> vbNullString) Then
					objModel.EndDate = "None"
					objModel.EndDateId = 0
				Else
					objModel.EndDate = standardReportSettings.GetFilterName(customDateId)
					objModel.EndDateId = customDateId
				End If
			Else
				objModel.IsDefaultDate = True
			End If

			' Record type selection
			If Session("optionRecordID") = "0" Then

				Dim strType As String = objDatabase.GetSystemSetting(strReportType, "Type", "A").ToString()

				Select Case strType
					Case "A"
						objModel.SelectionType = RecordSelectionType.AllRecords
					Case "P"
						objModel.SelectionType = RecordSelectionType.Picklist
						objModel.PicklistId = objDatabase.GetSystemSetting(strReportType, "ID", "0")
						objModel.PicklistName = standardReportSettings.GetPicklistFilterName(strType, objModel.PicklistId)
						If (objModel.PicklistName Is Nothing) Then
							objModel.PicklistName = "None"
						End If
					Case "F"
						objModel.SelectionType = RecordSelectionType.Filter
						objModel.FilterId = objDatabase.GetSystemSetting(strReportType, "ID", "0")
						objModel.FilterName = standardReportSettings.GetPicklistFilterName(strType, objModel.FilterId)
						If (objModel.FilterName Is Nothing) Then
							objModel.FilterName = "None"
						End If
				End Select
			End If

			' Flag to identify that the display of the picklist and filter title in header allowed
			objModel.DisplayTitleInReportHeader = objDatabase.GetSystemSetting(strReportType, "PrintFilterHeader", "0")

			Return View(objModel)

		End Function

		''' <summary>
		''' Saves report configuration
		''' </summary>
		''' <param name="objModel">The StandardReportsConfigurationModel model</param>
		''' <returns></returns>
		''' <remarks></remarks>
		<HttpPost>
		<ValidateAntiForgeryToken>
	 Function Absence_Breakdown_Configuration(objModel As StandardReportsConfigurationModel) As ActionResult

			Dim deserializer = New JavaScriptSerializer()
			Const strReportType As String = "AbsenceBreakdown"

			Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)

			' Deserialize the absence list from string
			If objModel.AbsenceTypesAsString IsNot Nothing Then
				If objModel.AbsenceTypesAsString.Length > 0 Then
					objModel.AbsenceTypes = deserializer.Deserialize(Of List(Of AbsenceType))(objModel.AbsenceTypesAsString)
				End If
			End If

			' Save absence types
			For Each absenceType As AbsenceType In objModel.AbsenceTypes
				objDatabase.SaveSystemSetting(strReportType, ("Absence Type " & absenceType.Type), absenceType.IsSelected)
			Next

			objDatabase.SaveSystemSetting(strReportType, "Custom Dates", objModel.IsCustomDate)

			If objModel.StartDate IsNot Nothing Then
				objDatabase.SaveSystemSetting(strReportType, "Start Date", objModel.StartDateId)
			End If

			If objModel.EndDate IsNot Nothing Then
				objDatabase.SaveSystemSetting(strReportType, "End Date", objModel.EndDateId)
			End If

			If objModel.SelectionType = RecordSelectionType.AllRecords Then
				objDatabase.SaveSystemSetting(strReportType, "Type", "A")
				objDatabase.SaveSystemSetting(strReportType, "ID", 0)
			ElseIf objModel.SelectionType = RecordSelectionType.Picklist Then
				objDatabase.SaveSystemSetting(strReportType, "Type", "P")
				objDatabase.SaveSystemSetting(strReportType, "ID", objModel.PicklistId)

			Else
				objDatabase.SaveSystemSetting(strReportType, "Type", "F")
				objDatabase.SaveSystemSetting(strReportType, "ID", objModel.FilterId)
			End If

			objDatabase.SaveSystemSetting(strReportType, ("PrintFilterHeader"), objModel.DisplayTitleInReportHeader)

			Return RedirectToAction("AbsenceBreakdownConfiguration")

		End Function

#End Region

		Public Function OrgChart() As PartialViewResult

			Dim m As OrgChart = New OrgChart()
			Dim model = m.LoadModel()

			Return PartialView(model)
		End Function

		Sub ResetSessionVars()
			Session("recordID") = ""
			Session("linkType") = ""
			Session("ViewDescription") = ""
		End Sub

		<ValidateAntiForgeryToken()>
		<HttpPost()>
		Public Function ShowChart(model As PopoutChartModel) As PartialViewResult

			Return PartialView("_showChart", model)

		End Function

		<HttpGet()>
		Function GetDefinitionsForType(UtilType As Integer, TableID As Integer, OnlyMine As Boolean, CategoryID As Integer, OwnerName As String) As JsonResult

			Dim rstDefSelRecords As DataTable

			' Get the records.
			Try

				Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

				Dim prmType = New SqlParameter("intType", SqlDbType.Int)
				prmType.Direction = ParameterDirection.Input
				prmType.Value = UtilType

				Dim prmOnlyMine = New SqlParameter("blnOnlyMine", SqlDbType.Bit)
				prmOnlyMine.Direction = ParameterDirection.Input
				prmOnlyMine.Value = OnlyMine

				Dim prmTableId = New SqlParameter("intTableID", SqlDbType.Int)
				prmTableId.Direction = ParameterDirection.Input

				If CleanNumeric(Request.Form("SelectedTableID")) = 0 Then
					prmTableId.Value = TableID
				Else
					prmTableId.Value = CleanNumeric(Request.Form("SelectedTableID"))
				End If

				Dim prmCategoryID = New SqlParameter("@intCategoryID", SqlDbType.Int)
				prmCategoryID.Direction = ParameterDirection.Input
				prmCategoryID.Value = CategoryID

				Dim prmOwner = New SqlParameter("@strOwner", SqlDbType.VarChar, 255)
				prmOwner.Direction = ParameterDirection.Input
				If OwnerName.IndexOf("'") > 0 Then
					prmOwner.Value = Replace(OwnerName, "'", "''")
				Else
					prmOwner.Value = OwnerName
				End If

				Dim userName = Session("Username")
				Dim isLoggedinUser = Session("OnlyMine")
				If isLoggedinUser = True AndAlso userName = OwnerName Then
					If userName.IndexOf("'") > 0 Then
						prmOwner.Value = Replace(userName, "'", "''")
					Else
						prmOwner.Value = userName
					End If
					Session("defsel_ownerName_" & UtilType) = userName
				Else
					Session("defsel_ownerName_" & UtilType) = OwnerName
				End If


				If UtilType = UtilityType.utlCustomReport Or UtilType = UtilityType.utlCrossTab Or UtilType = UtilityType.utlCalendarReport Or UtilType = UtilityType.utlMailMerge Or UtilType = UtilityType.utlNineBoxGrid Then
					Session("CategoryID") = CategoryID
				End If

				rstDefSelRecords = objDataAccess.GetDataTable("sp_ASRIntPopulateDefsel", CommandType.StoredProcedure, prmType, prmOnlyMine, prmTableId, prmCategoryID, prmOwner)

				Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
				Dim rows As New List(Of Dictionary(Of String, Object))()
				Dim row As Dictionary(Of String, Object)
				For Each dr As DataRow In rstDefSelRecords.Rows
					row = New Dictionary(Of String, Object)()
					For Each col As DataColumn In rstDefSelRecords.Columns
						row.Add(col.ColumnName, dr(col))
					Next
					rows.Add(row)
				Next




				Dim results = New With {.total = 1, .page = 1, .records = 0, .rows = rows}
				Return Json(results, JsonRequestBehavior.AllowGet)

			Catch ex As Exception
				Throw
			End Try


		End Function

		<HttpGet()>
		Public Function GetDefaultCalcValueForColumn(defaultCalcColumns As String) As JsonResult

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
			Dim prmRecordCount As New SqlParameter("piRecordCount", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

			Try
				Dim sFromDef = Session("RealSource") & Chr(9)	' NPG: Really not sure if this is right. It's supposed to replicate the session('FromDef') in recedit.

				Dim filterDefForCurrentTable As String = IIf(IsNothing(Session("filterDef_" & Session("tableID"))), "", Session("filterDef_" & Session("tableID")))

				Dim rstRecord = objDataAccess.GetDataTable("sp_ASRIntCalcDefaults", CommandType.StoredProcedure _
						, prmRecordCount _
						, New SqlParameter("psFromDef", SqlDbType.VarChar, -1) With {.Value = sFromDef} _
						, New SqlParameter("psFilterDef", SqlDbType.VarChar, -1) With {.Value = filterDefForCurrentTable} _
						, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("tableID"))} _
						, New SqlParameter("piParentTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("parentTableID"))} _
						, New SqlParameter("piParentRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("parentRecordID"))} _
						, New SqlParameter("psDefaultCalcColumns", SqlDbType.VarChar, -1) With {.Value = defaultCalcColumns} _
						, New SqlParameter("psDecimalSeparator", SqlDbType.VarChar, 255) With {.Value = Session("LocaleDecimalSeparator")} _
						, New SqlParameter("psLocaleDateFormat", SqlDbType.VarChar, 255) With {.Value = Platform.LocaleDateFormatForSQL()})

				Dim rows As New List(Of Dictionary(Of String, Object))()
				Dim row As Dictionary(Of String, Object)

				If Not rstRecord Is Nothing Then
					If rstRecord.Rows.Count > 0 Then
						For iloop = 0 To (rstRecord.Columns.Count - 1)
							If Not IsDBNull(rstRecord(iloop)) Then
								row = New Dictionary(Of String, Object)()
								row.Add(rstRecord.Columns(iloop).ColumnName, Replace(rstRecord(0)(iloop).ToString(), """", "&quot;"))
								rows.Add(row)
							End If
						Next
					End If
				End If

				Return Json(rows, JsonRequestBehavior.AllowGet)

			Catch ex As Exception
				Throw
			End Try

		End Function

		<ValidateAntiForgeryToken()>
		<HttpPost>
		Function WorkflowOutOfOffice_Check() As JsonResult

			Dim bOutOfOffice As Boolean
			Dim iRecordCount As Integer = 0
			Dim sErrorMessage As String = ""

			Try

				Dim objSession As SessionInfo = CType(Session("SessionContext"), SessionInfo)
				Dim objDataAccess As New clsDataAccess(objSession.LoginInfo)

				Dim prmOutOfOffice As SqlParameter = New SqlParameter("pfOutOfOffice", SqlDbType.Bit)
				prmOutOfOffice.Direction = ParameterDirection.Output

				Dim prmRecordCount As SqlParameter = New SqlParameter("piRecordCount", SqlDbType.Int)
				prmRecordCount.Direction = ParameterDirection.Output

				objDataAccess.ExecuteSP("spASRWorkflowOutOfOfficeCheck", prmOutOfOffice, prmRecordCount)

				bOutOfOffice = CBool(prmOutOfOffice.Value)
				iRecordCount = CInt(prmRecordCount.Value)

				If iRecordCount = 0 Then
					sErrorMessage = "Unable to set Workflow Out of Office.<br/>You do not have an identifiable personnel record."
				End If
			Catch ex As Exception
				sErrorMessage = "Unable to set your out of office.<br/><br/>Your personnel record cannot be updated."

			End Try

			Dim result = New With {.outOfOfficeOn = bOutOfOffice, .recordCount = iRecordCount, .error = sErrorMessage}
			Return Json(result, JsonRequestBehavior.AllowGet)

		End Function

		<ValidateAntiForgeryToken()>
		<HttpPost>
		Function WorkflowOutOfOffice_Enable(enable As Boolean) As JsonResult

			Dim iRecordCount As Integer = 0
			Dim sErrorMessage As String = ""

			Try

				Dim objSession As SessionInfo = CType(Session("SessionContext"), SessionInfo)
				Dim objDataAccess As New clsDataAccess(objSession.LoginInfo)

				Dim prmSetOffice As SqlParameter = New SqlParameter("pfOutOfOffice", SqlDbType.Bit)
				prmSetOffice.Value = enable
				objDataAccess.ExecuteSP("spASRWorkflowOutOfOfficeSet", prmSetOffice)

			Catch ex As Exception
				sErrorMessage = "Unable to set your out of office.<br/>Your personnel record cannot be updated."

			End Try

			Dim result = New With {.outOfOfficeOn = enable, .error = sErrorMessage}
			Return Json(result, JsonRequestBehavior.AllowGet)

		End Function

		' Reset the page source and IsLoadedFromReportDefinition flag which used to identify that if the tools screen loaded inside the ToolsFrame from the report definition
		<HttpPost()>
		<ValidateAntiForgeryToken>
		Sub ResetPageSourceFlag(isLoadedFromReportDefinition As Boolean, Optional ByVal utiltype As UtilityType? = Nothing)
			If (utiltype IsNot Nothing) Then
				Session("defseltype") = utiltype
			End If
			Session("IsLoadedFromReportDefinition") = isLoadedFromReportDefinition
		End Sub

		' Gets report & utilities which matches the search criteria
		<ValidateAntiForgeryToken()>
		<HttpPost>
		Public Function GetDefinitionSearchResult(value As DefinitionSearchResultModel) As JsonResult

			Dim sErrorMessage As String = ""
			Dim outputList As New List(Of DefinitionSearchResultModel)

			Try

				Dim searchResultRow As DefinitionSearchResultModel
				Dim objSessionInfo = CType(Session("SessionContext"), SessionInfo)
				Dim objDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

				Dim outputResult = objDataAccess.GetDataSet("sp_ASRGetReportsAndUtilitiesSearchResult" _
						, New SqlParameter("@searchText", SqlDbType.VarChar, 255) With {.Value = value.SearchText}).Tables(0)

				' Gets RUN permission for the logged in user
				Dim isMailMergeRunPermitted = objSessionInfo.IsPermissionGranted(UtilityType.utlMailMerge.ToSecurityPrefix, "RUN").ToString.ToLower
				Dim isCustomReportRunPermitted = objSessionInfo.IsPermissionGranted(UtilityType.utlCustomReport.ToSecurityPrefix, "RUN").ToString.ToLower
				Dim isCalendarReportRunPermitted = objSessionInfo.IsPermissionGranted(UtilityType.utlCalendarReport.ToSecurityPrefix, "RUN").ToString.ToLower
				Dim isCrossTabReportRunPermitted = objSessionInfo.IsPermissionGranted(UtilityType.utlCrossTab.ToSecurityPrefix, "RUN").ToString.ToLower
				Dim isNineBoxGridRunPermitted = objSessionInfo.IsPermissionGranted(UtilityType.utlNineBoxGrid.ToSecurityPrefix, "RUN").ToString.ToLower AndAlso Licence.IsModuleLicenced(SoftwareModule.NineBoxGrid)
				Dim isTalentRunPermitted = objSessionInfo.IsPermissionGranted(UtilityType.TalentReport.ToSecurityPrefix, "RUN").ToString.ToLower

				Dim isRunAllowed As Boolean = False
				Dim reportType As UtilityType

				For Each datarow As DataRow In outputResult.Rows

					isRunAllowed = False
					reportType = datarow("objectType")

					Select Case reportType
						Case UtilityType.utlMailMerge
							isRunAllowed = isMailMergeRunPermitted
						Case UtilityType.utlCustomReport
							isRunAllowed = isCustomReportRunPermitted
						Case UtilityType.utlCalendarReport
							isRunAllowed = isCalendarReportRunPermitted
						Case UtilityType.utlCrossTab
							isRunAllowed = isCrossTabReportRunPermitted
						Case UtilityType.utlNineBoxGrid
							isRunAllowed = isNineBoxGridRunPermitted
						Case UtilityType.TalentReport
							isRunAllowed = isTalentRunPermitted
					End Select

					' If edit/view allowed for the MailMerge and RUN allowed for the customreport, calendarreport, crosstab and ninebox grid then 
					'	only show those definition in the accordian menu search output
					If (isRunAllowed) Then
						searchResultRow = New DefinitionSearchResultModel
						searchResultRow.ReportType = datarow("objectType")
						searchResultRow.Id = datarow("ID")
						searchResultRow.Name = datarow("Name")
						searchResultRow.TextToDisplay = datarow("TextToDisplay")
						outputList.Add(searchResultRow)
					End If

				Next

			Catch ex As Exception
				sErrorMessage = "Unable to get search result."
			End Try

			Dim results = New With {.total = 1, .page = 1, .records = 0, .rows = outputList, .error = sErrorMessage}
			Return Json(results, JsonRequestBehavior.AllowGet)

		End Function

		<HttpPost()>
		<ValidateAntiForgeryToken>
		Public Function SetCookie(cookieName As String, cookieValue As String, expiryDays As String, path as String) As ActionResult

      Dim cookie = New HttpCookie(cookieName)
      cookie.Expires = Now.AddDays(expiryDays)
			cookie.HttpOnly = True
			cookie.Value = cookieValue
      cookie.Path = path

			Response.Cookies.Add(cookie)

			Return New HttpStatusCodeResult(HttpStatusCode.OK)

		End Function

		Private Function GetEmailAttachFileNameWithExcelExtension(fileName As String) As String
			Return IIf(Path.GetFileNameWithoutExtension(fileName) = fileName, fileName + ".xlsx", fileName)
		End Function

	End Class

	Public Class ErrMsgJsonAjaxResponse

		Public Property ErrorTitle As String
		Public Property ErrorMessage As String
		Public Property Redirect As String
		Public Property Route As String = "home"
	End Class

	Public Class ViewDataUploadFilesResult
		Public Property Name As String
		Public Property Length As Integer
	End Class

End Namespace

