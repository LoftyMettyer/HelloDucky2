'Imports ADODB
Imports System.Web.Configuration
Imports ADODB
Imports System.IO
Imports System.Drawing
Imports HR.Intranet.Server
Imports System.Data.SqlClient
Imports HR.Intranet.Server.Structures

Namespace Controllers
	Public Class AccountController
		Inherits Controller

		Function GetWidgetData(values As FormCollection,
		 widgetName As String,
		 isWidgetLogin As Boolean,
		 widgetUser As String,
		 widgetPassword As String,
		 widgetDatabase As String,
		 widgetServer As String,
		 widgetID As String) As JsonResult

			' get the session("databaseConnection") object
			Dim dbConn As Object

			'If isWidgetLogin Then
			If (Session("databaseConnection")) Is Nothing Then
				Login(values, isWidgetLogin, widgetUser, widgetPassword, widgetDatabase, widgetServer)
			End If

			If Not (Session("databaseConnection")) Is Nothing Then
				dbConn = Session("databaseConnection")

				Select Case widgetName
					Case "DBValue"

						' get collection of all links
						Dim objButtonInfo As Collection = Session("objButtonInfo")
						Dim sDBValue As String = ""
						Dim sCaption As String = ""
						Dim sFormattingSuffix As String = ""
						Dim sFormattingPrefix As String = ""

						' will Linq this at some point.
						For Each collectionItem As Object In objButtonInfo
							If collectionItem.ID = widgetID Then

								' sDBValue = DBValue(dbConn, 1, 163, 16771, 1, 4, 0, 0, 0)
								sDBValue = DBValue(collectionItem.Chart_TableID, collectionItem.Chart_ColumnID,
								 collectionItem.Chart_FilterID,
								 collectionItem.Chart_AggregateType, collectionItem.Element_Type,
								 collectionItem.Chart_SortOrderID, collectionItem.Chart_SortDirection,
								 collectionItem.Chart_ColourID)

								sCaption = collectionItem.Text

								sFormattingSuffix = collectionItem.Formatting_Suffix

								sFormattingPrefix = collectionItem.Formatting_Prefix
								Exit For
							End If

						Next
						Dim iDBValue As Integer

						Try
							iDBValue = CInt(sDBValue)
						Catch ex As Exception
							iDBValue = 0
						End Try

						Dim data = New JsonData_DBValue() _
						 With {.Caption = sCaption, .DBValue = iDBValue, .Formatting_Suffix = sFormattingSuffix,
						 .Formatting_Prefix = sFormattingPrefix}
						Return Json(data, JsonRequestBehavior.AllowGet)

					Case "GetLinks"

						' return a list of all ID's and element types and short descriptions
						Return Json(GetLinks, JsonRequestBehavior.AllowGet)


				End Select

			End If
			Return Json("Undefined")
		End Function

		Function GetLinks() As List(Of navigationLinks)

			Dim objButtonInfo As Collection = Session("objButtonInfo")

			Return (From collectionItem As Object In objButtonInfo Select New navigationLinks(collectionItem.linkType, collectionItem.linkOrder, collectionItem.prompt, collectionItem.text, collectionItem.element_Type, collectionItem.ID)).ToList()
		End Function

		Function DBValue(iChartTableID As Long,
		 iChartColumnID As Long,
		 iChartFilterID As Long,
		 iChartAggregateType As Long,
		 iChartElementType As Long,
		 iChartSortOrderID As Long,
		 iChartSortDirection As Long,
		 iChartColourID As Long) As String

			Dim objChart = New HR.Intranet.Server.clsChart
			objChart.SessionInfo = CType(Session("SessionContext"), SessionInfo)

			' reset the globals
			objChart.resetGlobals()

			Dim mrstDBValueData As DataTable

			mrstDBValueData = objChart.GetChartData(iChartTableID, iChartColumnID, iChartFilterID, iChartAggregateType,
				iChartElementType, 0, 0, 0, 0, iChartSortOrderID, iChartSortDirection, iChartColourID)

			If (Err.Number <> 0) Then
				Session("ErrorTitle") = "The Database Values could not be retrieved." & vbCrLf & FormatError(Err.Description)
			Else
				Session("ErrorTitle") = ""
			End If
			Dim sText As String = ""

			If Len(Session("ErrorTitle")) = 0 Then
				Try

					If mrstDBValueData.Rows.Count > 0 Then
						For Each objRow As DataRow In mrstDBValueData.Rows
							sText = objRow(0).ToString()
						Next

					Else ' no results - return zero
						sText = "No Data"
					End If
				Catch ex As Exception
					sText = "No Data"
				End Try

			End If

			Return sText
		End Function

		Function Login() As ActionResult

			Dim objServerSession As HR.Intranet.Server.SessionInfo = Session("sessionContext")

			Session("ErrorText") = Nothing

			' Are we already logged in on the session?
			If Not objServerSession Is Nothing Then
				If objServerSession.ActiveConnections > 0 Then
					objServerSession.ActiveConnections += 1
					Return RedirectToAction("Main", "Home", New With {.SSIMode = ViewBag.SSIMode})
				End If
			End If

			Return View()
		End Function

		<HttpPost()>
		Function Login(values As FormCollection, Optional isWidgetLogin As Boolean = False,
					Optional widgetUser As String = "",
					Optional widgetPassword As String = "",
					Optional widgetDatabase As String = "",
					Optional widgetServer As String = "") As ActionResult

			'On Error Resume Next

			'Dim sReferringPage
			Dim sUserName As String
			Dim sPassword As String
			Dim sDatabaseName As String
			Dim sServerName As String
			Dim sLocaleDateFormat As String
			Dim sLocaleDateSeparator As String
			Dim sLocaleDecimalSeparator As String
			Dim sLocaleThousandSeparator As String
			Dim fForcePasswordChange As Boolean
			Dim sConnectString As String
			Dim bWindowsAuthentication As Boolean = False

			fForcePasswordChange = False
			Session("ConvertedDesktopColour") = "#f9f7fb"

			If Not isWidgetLogin Then
				' Read the User Name and Password from the Login form.
				sUserName = Request.Form("txtUserNameCopy")
				sPassword = Request.Form("txtPassword")
				sDatabaseName = Request.Form("txtDatabase")
				sServerName = Request.Form("txtServer")
				If Request.Form("chkWindowsAuthentication") = "on" Then
					bWindowsAuthentication = True
				End If
				sLocaleDateFormat = Request.Form("txtLocaleDateFormat")
				sLocaleDateSeparator = Request.Form("txtLocaleDateSeparator")

				sLocaleDecimalSeparator = Request.Form("txtLocaleDecimalSeparator")
				sLocaleThousandSeparator = Request.Form("txtLocaleThousandSeparator")

				Session("WordVer") = Request.Form("txtWordVer")
				Session("ExcelVer") = Request.Form("txtExcelVer")
			Else
				' Read the User Name and Password from the Login form.
				sUserName = widgetUser
				sPassword = widgetPassword
				sDatabaseName = widgetDatabase
				sServerName = ".\sql2012"
				' widgetServer
				bWindowsAuthentication = False
				sLocaleDateFormat = "ddmmYYYY"
				sLocaleDateSeparator = "/"

				sLocaleDecimalSeparator = "."
				sLocaleThousandSeparator = ","

				Session("WordVer") = "12"
				Session("ExcelVer") = "12"
			End If

			Session("LocaleDateFormat") = sLocaleDateFormat
			Session("LocaleDateSeparator") = sLocaleDateSeparator
			Session("LocaleDecimalSeparator") = sLocaleDecimalSeparator
			Session("LocaleThousandSeparator") = sLocaleThousandSeparator

			' Store the username, for use in forcedchangepassword.
			Session("Username") = LCase(sUserName)

			' Check if the server DLL is registered.
			Try

			Catch ex As Exception
				If Err.Number <> 0 Then
					Session("ErrorTitle") = "Login Page"
					Session("ErrorText") =
					 "You could not login to the OpenHR database because of the following reason:<p>COAInt_Server.DLL has not been registered on the IIS server.  Please contact support." & vbCrLf &
					 "error: " & Err.Number.ToString & ": " & Err.Description
					Return RedirectToAction("Loginerror")
				End If
			End Try

			Dim objSettings = New HR.Intranet.Server.clsSettings
			sConnectString = objSettings.GetSQLProviderString.ToString() & "Data Source=" & sServerName & ";Initial Catalog=" &
			 sDatabaseName & ";Application Name=OpenHR Intranet;DataTypeCompatibility=80;MARS Connection=True;"
			objSettings = Nothing

			' Different connection string depending if use are using Windows Authentication
			If bWindowsAuthentication Then
				sConnectString = sConnectString & ";Trusted_Connection=yes;"
				sConnectString = sConnectString & ";Integrated Security=SSPI;"
			Else
				sConnectString = sConnectString & ";User ID=" & sUserName & ";Password=" & sPassword
			End If

			sConnectString = sConnectString & ";Persist Security Info=True;"

			' Open a connection to the database.
			Dim conX As New Connection
			conX.ConnectionTimeout = 60

			Try
				conX.Open(sConnectString)
			Catch ex As Exception
				If InStr(1, Err.Description, "The password of the account must be changed") Or
				 InStr(1, Err.Description, "The password for this login has expired") Or
				 InStr(1, Err.Description, "The password of the account has expired") Then
					Session("SQL2005Force") = "Server=" & sServerName & ";UID=" & sUserName
					Return RedirectToAction("ForcedPasswordChange", "Account")
				End If

				If Err.Number <> 0 Then
					Session("ErrorTitle") = "Login Page"
					Session("ErrorText") = "The system could not log you on. Make sure your details are correct, then retype your password."
					Return RedirectToAction("Loginerror")
				End If
			End Try

			' Set no command timeout
			conX.CommandTimeout = 0

			' Enter the current session in the poll table. This will ensure that even if the login checks fail, the session will still be killed after 1 minute.
			Dim cmdHit = New Command
			cmdHit.CommandText = "sp_ASRIntPoll"
			cmdHit.CommandType = 4
			' Stored Procedure
			cmdHit.ActiveConnection = conX
			Err.Number = 0
			cmdHit.Execute()
			If (Err.Number <> 0) Then
				Session("ErrorTitle") = "Login Page"
				Dim sErrorText = "You could not login to the OpenHR database because of the following reason:<p>"

				If (Err.Number = -2147217900) _
				 And (UCase(Left(FormatError(Err.Description), 31)) = "COULD NOT FIND STORED PROCEDURE") Then
					sErrorText = sErrorText &
					 "The database has not been scripted to run the intranet.<P>" &
					 "Contact your system administrator."
				Else
					sErrorText = sErrorText & FormatError(Err.Description)
				End If

				Session("ErrorText") = sErrorText
				Return RedirectToAction("Loginerror")
			End If
			cmdHit = Nothing

			Session("databaseConnection") = conX

			Dim objLogin As New HR.Intranet.Server.Structures.LoginInfo
			Dim objServerSession As New HR.Intranet.Server.SessionInfo

			objLogin.Server = sServerName
			objLogin.Database = sDatabaseName
			objLogin.Username = sUserName
			objLogin.Password = sPassword
			objLogin.TrustedConnection = bWindowsAuthentication

			objServerSession.Username = sUserName
			objServerSession.LoginInfo = objLogin
			objServerSession.Initialise()

			Dim objDataAccess As New clsDataAccess(objServerSession.LoginInfo)
			Dim objDatabase As New Database

			objDatabase.SessionInfo = objServerSession
			Session("DatabaseFunctions") = objDatabase
			Session("DatabaseAccess") = objDataAccess


			Try
				objDatabase.CheckLogin(objServerSession.LoginInfo, Session("version").ToString())

				If objServerSession.LoginInfo.LoginFailReason.Length <> 0 Then
					Session("ErrorText") = "You could not login to the OpenHR database because of the following reason:<p>" &
					FormatError(objServerSession.LoginInfo.LoginFailReason)
					Return RedirectToAction("Loginerror")
				End If

			Catch ex As Exception
				' These error codes need updating
				Session("ErrorTitle") = "Login Page"
				If Err.Number = -2147217900 Then
					Session("ErrorText") = "Unable to login to the OpenHR database:<p>" &
					 FormatError(
						"Please ask the System Administrator to update the database in the System Manager.")
					Return RedirectToAction("Loginerror")
				Else
					Session("ErrorText") = "You could not login to the OpenHR database because of the following reason:<p>" &
					 FormatError(Err.Description)
				End If
				Return RedirectToAction("Loginerror")

			End Try

			Session("sessionContext") = objServerSession
			Session("Server") = sServerName
			Session("Database") = sDatabaseName
			Session("WinAuth") = bWindowsAuthentication
			Session("userType") = objServerSession.LoginInfo.UserType
			Session("SelfServiceUserType") = objServerSession.LoginInfo.SelfServiceUserType
			Session("UserGroup") = objServerSession.LoginInfo.UserGroup


			' If the users default database is not 'master' then make it so.
			Dim cmdDefaultDB = New Command
			cmdDefaultDB.CommandText =
			 "IF EXISTS(SELECT 1 FROM master..syslogins WHERE loginname = SUSER_NAME() AND dbname <> 'master')" & vbNewLine &
			 "	EXEC sp_defaultdb [" & sUserName & "], master"
			cmdDefaultDB.ActiveConnection = conX
			cmdDefaultDB.Execute()
			cmdDefaultDB = Nothing

			' Put the username in a session variable	


			' RH 18/04/01 - Put entry in the audit access log
			Dim cmdAudit = New Command
			cmdAudit.CommandText = "sp_ASRIntAuditAccess"
			cmdAudit.CommandType = 4
			' Stored Procedure
			cmdAudit.ActiveConnection = conX

			Dim prmLoggingIn = cmdAudit.CreateParameter("LoggingIn", 11, 1, , True)	' 11 = boolean, 3 = int, 200 = varchar, 2 = output, 8000 = size
			cmdAudit.Parameters.Append(prmLoggingIn)

			Dim prmUser = cmdAudit.CreateParameter("Username", 200, 1, 1000)
			cmdAudit.Parameters.Append(prmUser)
			prmUser.Value = sUserName

			Err.Number = 0
			cmdAudit.Execute()

			If (Err.Number <> 0) Then
				Session("ErrorTitle") = "Login Page - Audit Access"
				Session("ErrorText") = "You could not login to the OpenHR database because of the following reason:<p>" &
				 FormatError(Err.Description)
				Return RedirectToAction("Loginerror")
			End If

			cmdAudit = Nothing

			' Successful login.
			Dim dtSettings = objDataAccess.GetDataTable("spASRIntGetSessionSettings", CommandType.StoredProcedure)
			Dim rowSettings = dtSettings.Rows(0)

			Session("FindRecords") = CLng(rowSettings("BlockSize"))
			Session("PrimaryStartMode") = CInt(rowSettings("PrimaryStartMode"))
			Session("HistoryStartMode") = CInt(rowSettings("HistoryStartMode"))
			Session("LookupStartMode") = CInt(rowSettings("LookupStartMode"))
			Session("QuickAccessStartMode") = CInt(rowSettings("QuickAccessStartMode"))
			Session("ExprColourMode") = CLng(rowSettings("ExprColourMode"))
			Session("ExprNodeMode") = CLng(rowSettings("ExprNodeMode"))
			Session("SupportTelNo") = rowSettings("SupportTelNo")
			Session("SupportFax") = rowSettings("SupportFax")
			Session("SupportEmail") = rowSettings("SupportEmail")
			Session("SupportWebpage") = rowSettings("SupportWebpage")
			Session("DesktopColour") = rowSettings("DesktopColour")


			' Convert the read Desktop colour into a value that can be used by the BODY tag.
			Select Case CStr(Session("DesktopColour"))
				Case "-2147483648"
					Session("ConvertedDesktopColour") = "scrollbar"
				Case "-2147483647"
					Session("ConvertedDesktopColour") = "background"
				Case "-2147483646"
					Session("ConvertedDesktopColour") = "activecaption"
				Case "-2147483645"
					Session("ConvertedDesktopColour") = "inactivecaption"
				Case "-2147483644"
					Session("ConvertedDesktopColour") = "menu"
				Case "-2147483643"
					Session("ConvertedDesktopColour") = "window"
				Case "-2147483642"
					Session("ConvertedDesktopColour") = "windowframe"
				Case "-2147483641"
					Session("ConvertedDesktopColour") = "menutext"
				Case "-2147483640"
					Session("ConvertedDesktopColour") = "windowtext"
				Case "-2147483639"
					Session("ConvertedDesktopColour") = "captiontext"
				Case "-2147483638"
					Session("ConvertedDesktopColour") = "activeborder"
				Case "-2147483637"
					Session("ConvertedDesktopColour") = "inactiveborder"
				Case "-2147483636"
					Session("ConvertedDesktopColour") = "appworkspace"
				Case "-2147483635"
					Session("ConvertedDesktopColour") = "highlight"
				Case "-2147483634"
					Session("ConvertedDesktopColour") = "highlighttext"
				Case "-2147483633"
					Session("ConvertedDesktopColour") = "threedface"
				Case "-2147483632"
					Session("ConvertedDesktopColour") = "threedshadow"
				Case "-2147483631"
					Session("ConvertedDesktopColour") = "graytext"
				Case "-2147483630"
					Session("ConvertedDesktopColour") = "buttontext"
				Case "-2147483629"
					Session("ConvertedDesktopColour") = "inactivecaptiontext"
				Case "-2147483628"
					Session("ConvertedDesktopColour") = "threedhighlight"
				Case "-2147483627"
					Session("ConvertedDesktopColour") = "threeddarkshadow"
				Case "-2147483626"
					Session("ConvertedDesktopColour") = "threedlightshadow"
				Case "-2147483625"
					Session("ConvertedDesktopColour") = "infotext"
				Case "-2147483624"
					Session("ConvertedDesktopColour") = "infobackground"
				Case Else
					Dim sColour = Hex(CLng(Session("DesktopColour")))

					Do While (Len(sColour) < 6)
						sColour = "0" & sColour
					Loop

					Dim sConvertedColour = "#"
					sConvertedColour = sConvertedColour & Mid(sColour, 5, 2)
					sConvertedColour = sConvertedColour & Mid(sColour, 3, 2)
					sConvertedColour = sConvertedColour & Mid(sColour, 1, 2)

					Session("ConvertedDesktopColour") = sConvertedColour
			End Select


			' Get Personnel module parameters		
			Dim cmdPersonnel = New Command
			cmdPersonnel.CommandText = "sp_ASRIntGetPersonnelParameters"
			cmdPersonnel.CommandType = 4 ' Stored Procedure
			cmdPersonnel.ActiveConnection = conX

			Dim prmEmpTableID = cmdPersonnel.CreateParameter("empTableID", 3, 2) ' 3=integer, 2=output
			cmdPersonnel.Parameters.Append(prmEmpTableID)

			Err.Number = 0
			cmdPersonnel.Execute()

			If (Err.Number <> 0) Then
				Session("ErrorTitle") = "Login Page"
				Session("ErrorText") = "You could not login to the OpenHR database because of the following reason:<p>" &
				 FormatError(Err.Description)
				Return RedirectToAction("Loginerror")
			End If

			Session("Personnel_EmpTableID") = cmdPersonnel.Parameters("empTableID").Value

			cmdPersonnel = Nothing

			' Get Workflow module parameters		
			Dim cmdWorkflow = New Command
			cmdWorkflow.CommandText = "spASRIntGetWorkflowParameters"
			cmdWorkflow.CommandType = 4
			' Stored Procedure
			cmdWorkflow.ActiveConnection = conX

			Dim prmWFEnabled = cmdWorkflow.CreateParameter("wfEnabled", 11, 2)
			' 11=boolean, 2=output
			cmdWorkflow.Parameters.Append(prmWFEnabled)

			Err.Number = 0
			cmdWorkflow.Execute()

			If (Err.Number <> 0) Then
				Session("ErrorTitle") = "Login Page"
				Session("ErrorText") = "You could not login to the OpenHR database because of the following reason:<p>" &
				 FormatError(Err.Description)
				Return RedirectToAction("Loginerror")
			End If

			Session("WF_Enabled") = cmdWorkflow.Parameters("wfEnabled").Value

			cmdWorkflow = Nothing

			' Check if the OutOfOffice parameters have been configured.
			Dim fWorkflowOutOfOfficeConfigured = False
			Dim fWorkflowOutOfOffice = False
			Dim iWorkflowRecordCount = 0

			If Session("WF_Enabled") Then
				cmdWorkflow = New Command
				cmdWorkflow.CommandText = "spASRWorkflowOutOfOfficeConfigured"
				cmdWorkflow.CommandType = 4
				' Stored Procedure
				cmdWorkflow.ActiveConnection = conX

				Dim prmWFOutOfOfficeConfig = cmdWorkflow.CreateParameter("wfOutOfOfficeConfig", 11, 2)
				' 11=boolean, 2=output
				cmdWorkflow.Parameters.Append(prmWFOutOfOfficeConfig)

				Err.Number = 0
				cmdWorkflow.Execute()

				fWorkflowOutOfOfficeConfigured = cmdWorkflow.Parameters("wfOutOfOfficeConfig").Value

				cmdWorkflow = Nothing

				If fWorkflowOutOfOfficeConfigured Then
					' Check if the current user OutOfOffice
					Dim prmOutOfOffice As SqlParameter = New SqlParameter("pfOutOfOffice", SqlDbType.Bit)
					prmOutOfOffice.Direction = ParameterDirection.Output

					Dim prmRecordCount As SqlParameter = New SqlParameter("piRecordCount", SqlDbType.Int)
					prmRecordCount.Direction = ParameterDirection.Output

					objDataAccess.ExecuteSP("spASRWorkflowOutOfOfficeCheck", prmOutOfOffice, prmRecordCount)

					fWorkflowOutOfOffice = prmOutOfOffice.Value
					iWorkflowRecordCount = prmRecordCount.Value

				End If
			End If
			Session("WF_OutOfOfficeConfigured") = fWorkflowOutOfOfficeConfigured
			Session("WF_OutOfOffice") = fWorkflowOutOfOffice
			Session("WF_RecordCount") = iWorkflowRecordCount

			' Get Training Booking module parameters		
			Dim cmdTrainingBooking = New Command
			cmdTrainingBooking.CommandText = "sp_ASRIntGetTrainingBookingParameters"
			cmdTrainingBooking.CommandType = 4	' Stored Procedure
			cmdTrainingBooking.ActiveConnection = conX

			prmEmpTableID = cmdTrainingBooking.CreateParameter("empTableID", 3, 2)	' 3=integer, 2=output
			cmdTrainingBooking.Parameters.Append(prmEmpTableID)

			Dim prmCourseTableID = cmdTrainingBooking.CreateParameter("courseTableID", 3, 2) ' 3=integer, 2=output
			cmdTrainingBooking.Parameters.Append(prmCourseTableID)

			Dim prmCourseCancelDateColumnID = cmdTrainingBooking.CreateParameter("courseCancelDateColumnID", 3, 2) ' 3=integer, 2=output
			cmdTrainingBooking.Parameters.Append(prmCourseCancelDateColumnID)

			Dim prmTBTableID = cmdTrainingBooking.CreateParameter("tbTableID", 3, 2) ' 3=integer, 2=output
			cmdTrainingBooking.Parameters.Append(prmTBTableID)

			Dim prmTBTableSelect = cmdTrainingBooking.CreateParameter("tbTableSelect", 11, 2)	' 11=boolean, 2=output
			cmdTrainingBooking.Parameters.Append(prmTBTableSelect)

			Dim prmTBTableInsert = cmdTrainingBooking.CreateParameter("tbTableInsert", 11, 2)	' 11=boolean, 2=output
			cmdTrainingBooking.Parameters.Append(prmTBTableInsert)

			Dim prmTBTableUpdate = cmdTrainingBooking.CreateParameter("tbTableUpdate", 11, 2)	' 11=boolean, 2=output
			cmdTrainingBooking.Parameters.Append(prmTBTableUpdate)

			Dim prmTBStatusColumnID = cmdTrainingBooking.CreateParameter("tbStatusColumnID", 3, 2) ' 3=integer, 2=output
			cmdTrainingBooking.Parameters.Append(prmTBStatusColumnID)

			Dim prmTBStatusColumnUpdate = cmdTrainingBooking.CreateParameter("tbStatusColumnUpdate", 11, 2)	' 11=boolean, 2=output
			cmdTrainingBooking.Parameters.Append(prmTBStatusColumnUpdate)

			Dim prmTBCancelDateColumnID = cmdTrainingBooking.CreateParameter("tbCancelDateColumnID", 3, 2) ' 3=integer, 2=output
			cmdTrainingBooking.Parameters.Append(prmTBCancelDateColumnID)

			Dim prmTBCancelDateColumnUpdate = cmdTrainingBooking.CreateParameter("tbCancelDateColumnUpdate", 11, 2)	' 11=boolean, 2=output
			cmdTrainingBooking.Parameters.Append(prmTBCancelDateColumnUpdate)

			Dim prmTBStatusPExists = cmdTrainingBooking.CreateParameter("tbStatusPExists", 11, 2)	' 11=boolean, 2=output
			cmdTrainingBooking.Parameters.Append(prmTBStatusPExists)

			Dim prmWaitListTableID = cmdTrainingBooking.CreateParameter("waitListTableID", 3, 2) ' 3=integer, 2=output
			cmdTrainingBooking.Parameters.Append(prmWaitListTableID)

			Dim prmWaitListTableInsert = cmdTrainingBooking.CreateParameter("waitListTableInsert", 11, 2)	' 11=boolean, 2=output
			cmdTrainingBooking.Parameters.Append(prmWaitListTableInsert)

			Dim prmWaitListTableDelete = cmdTrainingBooking.CreateParameter("waitListTableDelete", 11, 2)	' 11=boolean, 2=output
			cmdTrainingBooking.Parameters.Append(prmWaitListTableDelete)

			Dim prmWaitListCourseTitleColumnID = cmdTrainingBooking.CreateParameter("waitListCourseTitleColumnID", 3, 2) ' 3=integer, 2=output
			cmdTrainingBooking.Parameters.Append(prmWaitListCourseTitleColumnID)

			Dim prmWaitListCourseTitleColumnUpdate = cmdTrainingBooking.CreateParameter("waitListCourseTitleColumnUpdate", 11, 2)	' 11=boolean, 2=output
			cmdTrainingBooking.Parameters.Append(prmWaitListCourseTitleColumnUpdate)

			Dim prmWaitListCourseTitleColumnSelect = cmdTrainingBooking.CreateParameter("waitListCourseTitleColumnSelect", 11, 2)	' 11=boolean, 2=output
			cmdTrainingBooking.Parameters.Append(prmWaitListCourseTitleColumnSelect)

			Dim prmBulkBookingDefaultViewID = cmdTrainingBooking.CreateParameter("bulkBookingDefaultViewID", 3, 2) ' 3=integer, 2=output
			cmdTrainingBooking.Parameters.Append(prmBulkBookingDefaultViewID)

			'		Set prmWaitListOverRideColumnID = cmdTrainingBooking.CreateParameter("WaitListOverRideColumnID", 3, 2) ' 3=integer, 2=output
			'		cmdTrainingBooking.Parameters.Append prmWaitListOverRideColumnID

			Err.Number = 0
			cmdTrainingBooking.Execute()

			If (Err.Number <> 0) Then
				Session("ErrorTitle") = "Login Page"
				Session("ErrorText") = "You could not login to the OpenHR database because of the following reason:<p>" &
				 FormatError(Err.Description)
				Return RedirectToAction("Loginerror")
			End If

			Session("TB_EmpTableID") = cmdTrainingBooking.Parameters("empTableID").Value

			Session("TB_CourseTableID") = cmdTrainingBooking.Parameters("courseTableID").Value
			Session("TB_CourseCancelDateColumnID") = cmdTrainingBooking.Parameters("courseCancelDateColumnID").Value

			Session("TB_TBTableID") = cmdTrainingBooking.Parameters("tbTableID").Value
			Session("TB_TBTableSelect") = cmdTrainingBooking.Parameters("tbTableSelect").Value
			Session("TB_TBTableInsert") = cmdTrainingBooking.Parameters("tbTableInsert").Value
			Session("TB_TBTableUpdate") = cmdTrainingBooking.Parameters("tbTableUpdate").Value
			Session("TB_TBStatusColumnID") = cmdTrainingBooking.Parameters("tbStatusColumnID").Value
			Session("TB_TBStatusColumnUpdate") = cmdTrainingBooking.Parameters("tbStatusColumnUpdate").Value
			Session("TB_TBCancelDateColumnID") = cmdTrainingBooking.Parameters("tbCancelDateColumnID").Value
			Session("TB_TBCancelDateColumnUpdate") = cmdTrainingBooking.Parameters("tbCancelDateColumnUpdate").Value
			Session("TB_TBStatusPExists") = cmdTrainingBooking.Parameters("tbStatusPExists").Value

			Session("TB_WaitListTableID") = cmdTrainingBooking.Parameters("waitListTableID").Value
			Session("TB_WaitListTableInsert") = cmdTrainingBooking.Parameters("waitListTableInsert").Value
			Session("TB_WaitListTableDelete") = cmdTrainingBooking.Parameters("waitListTableDelete").Value
			Session("TB_WaitListCourseTitleColumnID") = cmdTrainingBooking.Parameters("waitListCourseTitleColumnID").Value
			Session("TB_WaitListCourseTitleColumnUpdate") =
			 cmdTrainingBooking.Parameters("waitListCourseTitleColumnUpdate").Value
			Session("TB_WaitListCourseTitleColumnSelect") =
			 cmdTrainingBooking.Parameters("waitListCourseTitleColumnSelect").Value

			Session("TB_BulkBookingDefaultViewID") = cmdTrainingBooking.Parameters("bulkBookingDefaultViewID").Value
			'session("TB_WaitListOverRideColumnID") = cmdTrainingBooking.Parameters("WaitListOverRideColumnID").Value

			cmdTrainingBooking = Nothing

			If CStr(Session("TB_TBTableID")) = "" Then Session("TB_TBTableID") = 0

			'MH 07/07/2004: Moved from default.asp so background stuff only gets called on
			'login and not every time you go back to default.asp (as per request from JPD).

			Dim sTempPath As String, sBGImage As String = "", intBGPos As Short = 2, strRepeat As String, strBGPos As String

			Dim objUtilities = New Utilities
			objUtilities.SessionInfo = CType(Session("SessionContext"), SessionInfo)

			sTempPath = Server.MapPath("~/pictures")
			sBGImage = objUtilities.GetBackgroundPicture(CStr(sTempPath))
			intBGPos = objUtilities.GetBackgroundPosition

			objUtilities.OfficeInitialise(CInt(Session("WordVer")), CInt(Session("ExcelVer")))
			Session("WordFormats") = objUtilities.OfficeGetCommonDialogFormatsWord
			Session("ExcelFormats") = objUtilities.OfficeGetCommonDialogFormatsExcel
			Session("WordFormatDefaultIndex") = objUtilities.OfficeGetDefaultIndexWord
			Session("ExcelFormatDefaultIndex") = objUtilities.OfficeGetDefaultIndexExcel
			Session("OfficeSaveAsValues") = objUtilities.OfficeGetSaveAsValues

			objUtilities = Nothing

			Select Case intBGPos
				Case 0
					'Top Left
					strRepeat = "no-repeat"
					strBGPos = "top left"

				Case 1
					'Top Right
					strRepeat = "no-repeat"
					strBGPos = "top right"

				Case 2
					'Centre
					strRepeat = "no-repeat"
					strBGPos = "center"

				Case 3
					'Left Tile
					strRepeat = "repeat-y"
					strBGPos = "left"

				Case 4
					'Right Tile
					strRepeat = "repeat-y"
					strBGPos = "right"

				Case 5
					'Top Tile
					strRepeat = "repeat-x"
					strBGPos = "top"

				Case 6
					'Bottom Tile
					strRepeat = "repeat-x"
					strBGPos = "bottom"

				Case 7
					'Tile
					strRepeat = "repeat"
					strBGPos = "center"

				Case Else
					'Centre
					strRepeat = "no-repeat"
					strBGPos = "center"

			End Select



			' Get the configured Single record view ID. 	For the dashboard.
			Dim cmdModuleInfo = New Command
			cmdModuleInfo.CommandText = "spASRIntGetSingleRecordViewID"
			cmdModuleInfo.CommandType = 4	' Stored Procedure
			cmdModuleInfo.ActiveConnection = conX

			Dim prmTableID = cmdModuleInfo.CreateParameter("tableID", 3, 2)	' 3=integer, 2=output
			cmdModuleInfo.Parameters.Append(prmTableID)

			Dim prmViewID = cmdModuleInfo.CreateParameter("viewID", 3, 2)	' 3=integer, 2=output
			cmdModuleInfo.Parameters.Append(prmViewID)

			Err.Number = 0
			cmdModuleInfo.Execute()

			If (Err.Number <> 0) Then
				Session("ErrorTitle") = "Login Page - Module setup"
				Session("ErrorText") = "You could not login to the OpenHR database because of the following error :<p>" &
				 FormatError(Err.Description)
				Return RedirectToAction("Loginerror", "Account")
			End If

			Session("SingleRecordTableID") = cmdModuleInfo.Parameters("tableID").Value
			Session("SingleRecordViewID") = cmdModuleInfo.Parameters("viewID").Value

			' SSI Stuff:
			Session("SSILinkTableID") = 0
			Session("SSILinkViewID") = 0

			cmdModuleInfo = Nothing

			' Get the configured SSI Welcome info; name, last login date & time. 		
			cmdModuleInfo = New Command
			cmdModuleInfo.CommandText = "sp_ASRIntGetModuleParameter"
			cmdModuleInfo.CommandType = 4	' Stored Procedure
			cmdModuleInfo.ActiveConnection = conX

			Dim prmModuleKey = cmdModuleInfo.CreateParameter("moduleKey", 200, 1, 255)
			cmdModuleInfo.Parameters.Append(prmModuleKey)
			prmModuleKey.Value = "MODULE_PERSONNEL"

			Dim prmParameterKey = cmdModuleInfo.CreateParameter("parameterKey", 200, 1, 255)
			cmdModuleInfo.Parameters.Append(prmParameterKey)
			prmParameterKey.Value = "Param_FieldsSSIWelcome"

			Dim prmParameterValue = cmdModuleInfo.CreateParameter("parameterValue", 200, 2, 1000)
			cmdModuleInfo.Parameters.Append(prmParameterValue)

			Err.Clear()
			cmdModuleInfo.Execute()

			If (Err.Number <> 0) Then
				Session("ErrorTitle") = "Login Page - Module setup"
				Session("ErrorText") = "You could not login to the OpenHR database because of the following error :<p>" & FormatError(Err.Description)
				Return RedirectToAction("loginerror")
			End If

			Dim lngSSIWelcomeColumnID = CLng(cmdModuleInfo.Parameters("parameterValue").Value)

			If lngSSIWelcomeColumnID <= 0 Then lngSSIWelcomeColumnID = 0

			' photo
			cmdModuleInfo = New Command
			cmdModuleInfo.CommandText = "sp_ASRIntGetModuleParameter"
			cmdModuleInfo.CommandType = CommandType.StoredProcedure
			cmdModuleInfo.ActiveConnection = conX

			prmModuleKey = cmdModuleInfo.CreateParameter("moduleKey", 200, 1, 255)
			cmdModuleInfo.Parameters.Append(prmModuleKey)
			prmModuleKey.Value = "MODULE_PERSONNEL"

			prmParameterKey = cmdModuleInfo.CreateParameter("parameterKey", 200, 1, 255)
			cmdModuleInfo.Parameters.Append(prmParameterKey)
			prmParameterKey.Value = "Param_FieldsSSIPhotograph"

			prmParameterValue = cmdModuleInfo.CreateParameter("parameterValue", 200, 2, 1000)
			cmdModuleInfo.Parameters.Append(prmParameterValue)

			Err.Clear()
			cmdModuleInfo.Execute()

			If (Err.Number <> 0) Then
				Session("ErrorTitle") = "Login Page - Module setup"
				Session("ErrorText") = "You could not login to the OpenHR database because of the following error :<p>" & FormatError(Err.Description)
				Return RedirectToAction("loginerror")
			End If

			Dim lngSSIPhotographColumnID As Long = 0
			If Not IsDBNull(cmdModuleInfo.Parameters("parameterValue").Value) Then
				lngSSIPhotographColumnID = CLng(cmdModuleInfo.Parameters("parameterValue").Value)
			End If

			If lngSSIPhotographColumnID <= 0 Then lngSSIPhotographColumnID = 0

			Dim cmdSSIWelcomeDetails = Nothing

			cmdSSIWelcomeDetails = New Command
			cmdSSIWelcomeDetails.CommandText = "spASRIntGetSSIWelcomeDetails"
			cmdSSIWelcomeDetails.CommandType = CommandType.StoredProcedure
			cmdSSIWelcomeDetails.ActiveConnection = conX

			prmModuleKey = cmdSSIWelcomeDetails.CreateParameter("WelcomeColumnID", DataTypeEnum.adInteger, ParameterDirection.Input)
			cmdSSIWelcomeDetails.Parameters.Append(prmModuleKey)
			prmModuleKey.Value = lngSSIWelcomeColumnID

			prmModuleKey = cmdSSIWelcomeDetails.CreateParameter("SSIPhotographColumnID", DataTypeEnum.adInteger, ParameterDirection.Input)
			cmdSSIWelcomeDetails.Parameters.Append(prmModuleKey)
			prmModuleKey.Value = lngSSIPhotographColumnID

			prmParameterKey = cmdSSIWelcomeDetails.CreateParameter("SingleRecordViewID", DataTypeEnum.adInteger, ParameterDirection.Input)
			cmdSSIWelcomeDetails.Parameters.Append(prmParameterKey)
			prmParameterKey.Value = Session("SingleRecordViewID")

			prmParameterKey = cmdSSIWelcomeDetails.CreateParameter("UserName", DataTypeEnum.adVarChar, ParameterDirection.Input, 255)
			cmdSSIWelcomeDetails.Parameters.Append(prmParameterKey)
			prmParameterKey.Value = Session("username")

			prmParameterValue = cmdSSIWelcomeDetails.CreateParameter("WelcomeMessage", DataTypeEnum.adVarChar, ParameterDirection.Output, 255)
			cmdSSIWelcomeDetails.Parameters.Append(prmParameterValue)

			prmParameterValue = cmdSSIWelcomeDetails.CreateParameter("WelcomeName", DataTypeEnum.adVarChar, ParameterDirection.Output, 255)
			cmdSSIWelcomeDetails.Parameters.Append(prmParameterValue)

			prmParameterValue = cmdSSIWelcomeDetails.CreateParameter("SelfServicePhotograph", DataTypeEnum.adVarBinary, ParameterDirection.Output, 10000000)	'Big number, we couldn't determine the MAX value for a VarBinary
			cmdSSIWelcomeDetails.Parameters.Append(prmParameterValue)

			Err.Clear()
			cmdSSIWelcomeDetails.Execute()

			If (Err.Number <> 0) Then
				Session("welcomemessage") = "error: " & Err.Description & ". ID: " & CStr(lngSSIWelcomeColumnID)
			Else
				Session("welcomemessage") = cmdSSIWelcomeDetails.Parameters("WelcomeMessage").Value
				Session("welcomeName") = cmdSSIWelcomeDetails.Parameters("WelcomeName").Value
				If Not IsDBNull(cmdSSIWelcomeDetails.Parameters("SelfServicePhotograph").Value) Then
					Dim OLEType As Short = Val(Encoding.UTF8.GetString(cmdSSIWelcomeDetails.Parameters("SelfServicePhotograph").Value, 8, 2))
					If OLEType = 2 Then	'Embeded
						Dim abtImage = CType(cmdSSIWelcomeDetails.Parameters("SelfServicePhotograph").Value, Byte())
						Dim binaryData As Byte() = New Byte(abtImage.Length - 400) {}
						Try
							Buffer.BlockCopy(abtImage, 400, binaryData, 0, abtImage.Length - 400)
							'Create an image based on the embeded (Base64) image and resize it to 48x48
							Dim img As Image = Base64StringToImage(Convert.ToBase64String(binaryData, 0, binaryData.Length))
							img = img.GetThumbnailImage(48, 48, Nothing, IntPtr.Zero)
							Session("SelfServicePhotograph_Src") = "data:image/jpeg;base64," & ImageToBase64String(img)
						Catch exp As System.ArgumentNullException

						End Try
					ElseIf OLEType = 3 Then	'Link
						Dim UNC As String = Trim(Encoding.UTF8.GetString(cmdSSIWelcomeDetails.Parameters("SelfServicePhotograph").Value, 290, 60))
						Dim FileName As String = Trim(Path.GetFileName(Encoding.UTF8.GetString(cmdSSIWelcomeDetails.Parameters("SelfServicePhotograph").Value, 10, 70))).Replace("\", "/")
						Dim FullPath As String = Trim(Encoding.UTF8.GetString(cmdSSIWelcomeDetails.Parameters("SelfServicePhotograph").Value, 80, 210)).Replace("\", "/")
						Session("SelfServicePhotograph_src") = "file:///" & UNC & "/" & FullPath & "/" & FileName
					End If
				Else 'No picture is defined for user, use anonymous one
					Session("SelfServicePhotograph_Src") = Url.Content("~/Content/images/anonymous.png")
				End If
			End If

			'clear the welcome message if neither the name or last logon information are available.
			If Session("welcomemessage") = "Welcome " Then
				Session("welcomemessage") = ""
				Session("welcomeName") = ""
			End If

			cmdSSIWelcomeDetails = Nothing

			Session("EnableSQL2000Functions") = False

			If Session("WinAuth") Then
				' Do not force password change for Windows Authenticated users.
				fForcePasswordChange = False
			End If

			If fForcePasswordChange = True Then
				' Force password change only if there are no other users logged in with the same name.
				Dim cmdCheckUserSessions = New Command
				cmdCheckUserSessions.CommandText = "spASRGetCurrentUsersCountOnServer"
				cmdCheckUserSessions.CommandType = 4 ' Stored procedure.
				cmdCheckUserSessions.ActiveConnection = Session("databaseConnection")

				Dim prmCount = cmdCheckUserSessions.CreateParameter("count", 3, 2) ' 3=integer, 2=output
				cmdCheckUserSessions.Parameters.Append(prmCount)

				Dim prmUserName = cmdCheckUserSessions.CreateParameter("userName", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
				cmdCheckUserSessions.Parameters.Append(prmUserName)
				prmUserName.Value = Session("Username")

				Err.Number = 0
				cmdCheckUserSessions.Execute()

				Dim iUserSessionCount = CLng(cmdCheckUserSessions.Parameters("count").Value)
				cmdCheckUserSessions = Nothing

				If iUserSessionCount < 2 Then
					Return RedirectToAction("ForcedPasswordChange")
				Else
					Return RedirectToAction("Main", "Home")
				End If
			Else



				Dim cookie = New HttpCookie("Login")
				cookie.Expires = DateTime.Now.AddYears(1)
				cookie.HttpOnly = True
				cookie("User") = Request.Form("txtUserNameCopy")
				'dont save or retrieve these anymore HRPRO-3030 / 3031
				'cookie("Database") = Request.Form("txtDatabase")
				'cookie("Server") = Request.Form("txtServer")
				cookie("WindowsAuthentication") = Request.Form("chkWindowsAuthentication")
				Response.Cookies.Add(cookie)

				' HRPRO-3531
				Session("MSBrowser") = (Request.Form("txtMSBrowser") = "true")

				If Session("DMIRequiresIE") = "TRUE" And Session("MSBrowser") <> True Then
					' non-IE browsers don't get DMI access yet.
					ViewBag.SSIMode = True
				Else
					Select Case Session("SelfServiceUserType")
						Case 1		'IF DMI Multi
							' Return RedirectToAction("Main", "Home")
							ViewBag.SSIMode = False
						Case 2		'IF DMI Single
							' Return RedirectToAction("Main", "Home")
							ViewBag.SSIMode = False
						Case 3		'IF DMI Single And SSI
							' Return RedirectToAction("LinksMain", "Home")
							ViewBag.SSIMode = True
						Case 4		'IF SSI Only
							' Return RedirectToAction("LinksMain", "Home")
							ViewBag.SSIMode = True
						Case Else
							Return RedirectToAction("login", "account")
					End Select
				End If

				' always main.
				Return RedirectToAction("Main", "Home", New With {.SSIMode = ViewBag.SSIMode})

			End If

			Return RedirectToAction("login", "account")
		End Function

		Public Function LogOff()

			Dim objServerSession As HR.Intranet.Server.SessionInfo = Session("sessionContext")
			Dim objConnection As Connection

			Dim objDataAccess As New clsDataAccess(objServerSession.LoginInfo)

			Session("ErrorText") = Nothing

			Try

				Dim prmLogIn = New SqlParameter("blnLoggingIn", SqlDbType.Bit, 1, ParameterDirection.Input)
				prmLogIn.Value = False

				Dim prmUserName = New SqlParameter("strUsername", SqlDbType.VarChar, 1000, ParameterDirection.Input)
				prmUserName.Value = Replace(Session("Username"), "'", "''")

				objDataAccess.ExecuteSP("sp_ASRIntAuditAccess", prmLogIn, prmUserName)

				objConnection = Session("databaseConnection")

				If objConnection.State = 1 Then
					objConnection.Close()
				End If

				Session("databaseConnection") = Nothing
				Session("avPrimaryMenuInfo") = Nothing
				Session("avSubMenuInfo") = Nothing
				Session("avQuickEntryMenuInfo") = Nothing
				Session("avTableMenuInfo") = Nothing
				Session("avTableHistoryMenuInfo") = Nothing

				objServerSession.ActiveConnections -= 1

			Catch ex As Exception
				Throw

			End Try

			Return RedirectToAction("Login", "Account")

		End Function

		<HttpPost()>
		Function ForcedPasswordChange_Submit(value As FormCollection) As ActionResult

			On Error Resume Next

			Dim strErrorMessage As String = ""
			Dim sConnectString As String = ""

			Dim fSubmitPasswordChange = (Len(Request.Form("txtGotoPage")) = 0)

			If fSubmitPasswordChange Then
				' Force password change only if there are no other users logged in with the same name.
				Dim cmdCheckUserSessions = CreateObject("ADODB.Command")
				cmdCheckUserSessions.CommandText = "spASRGetCurrentUsersCountOnServer"
				cmdCheckUserSessions.CommandType = 4 ' Stored procedure.
				cmdCheckUserSessions.ActiveConnection = Session("databaseConnection")

				Dim prmCount = cmdCheckUserSessions.CreateParameter("count", 3, 2) ' 3=integer, 2=output
				cmdCheckUserSessions.Parameters.Append(prmCount)

				Dim prmUserName = cmdCheckUserSessions.CreateParameter("userName", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
				cmdCheckUserSessions.Parameters.Append(prmUserName)
				prmUserName.value = Session("Username")

				Err.Clear()
				cmdCheckUserSessions.Execute()

				Dim iUserSessionCount = CLng(cmdCheckUserSessions.Parameters("count").Value)
				cmdCheckUserSessions = Nothing

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

					' SQL Native Client Stuff
					If Err.Number = 3709 Then
						Err.Clear()

						Dim conX = CreateObject("ADODB.Connection")
						conX.ConnectionTimeout = 60

						Dim objSettings = New Global.HR.Intranet.Server.clsSettings

						Select Case objSettings.GetSQLNCLIVersion
							Case 9
								sConnectString = "Provider=SQLNCLI;"
							Case 10
								sConnectString = "Provider=SQLNCLI10;"
							Case 11
								sConnectString = "Provider=SQLNCLI11;"
						End Select
						objSettings = Nothing

						sConnectString = sConnectString & "DataTypeCompatibility=80;MARS Connection=False;" & Session("SQL2005Force") & _
							 ";Old Password='" & Replace(sCurrentPassword, "'", "''") & "';Password='" & Replace(sNewPassword, "'", "''") & "'"

						conX.open(sConnectString)

						If Err.Number <> 0 Then
							If Err.Number <> 3706 Then	 ' 3706 = Provider not found
								strErrorMessage = Err.Description
							End If
							Session("ErrorTitle") = "Change Password Page"
							Session("ErrorText") = strErrorMessage
							' Return RedirectToAction("Loginerror", "Account")
							Return RedirectToAction("Loginerror", "Account")
						Else
							conX.close()
							Session("MessageTitle") = "Change Password Page"
							Session("MessageText") = "Password changed successfully. You may now login."
							' Response.Redirect("loginmessage")
							Return RedirectToAction("Loginmessage", "Account")
						End If
					End If

					If Err.Number <> 0 Then
						Session("ErrorTitle") = "Change Password Page"
						Session("ErrorText") = "You could not change your password because of the following error:<p>" & Err.Description
						Return RedirectToAction("Loginerror", "Account")
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
							Session("ErrorText") = "You could not change your password because of the following error:<p>" & Err.Description
							Return RedirectToAction("Loginerror", "Account")
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
								Session("ErrorText") = "You could not change your password because of the following error:<p>" & Err.Description
								Return RedirectToAction("Loginerror", "Account")
							End If

							Session("databaseConnection") = conX
						End If

						Session("MessageTitle") = "Change Password Page"
						Session("MessageText") = "Password changed successfully."
						Response.Redirect("loginmessage")
						'Response.Redirect("confirmok")
					End If
				Else
					Session("ErrorTitle") = "Change Password Page"
					Dim sErrorText = "You could not change your password.<p>The account is currently being used by "
					If iUserSessionCount > 2 Then
						sErrorText = sErrorText & iUserSessionCount & " users"
					Else
						sErrorText = sErrorText & "another user"
					End If
					sErrorText = sErrorText & " in the system."
					Session("ErrorText") = sErrorText

					Return RedirectToAction("Loginerror", "Account")
				End If
			Else
				' Go to the main page.
				Response.Redirect("main")
			End If
		End Function

		Function ForcedPasswordChange() As ActionResult
			Return View()
		End Function

		Function Loginerror() As ActionResult
			Return View()
		End Function

		Function Loginmessage() As ActionResult
			Return View()
		End Function

		Function About() As ActionResult
			Return View()
		End Function

		Function ForgotPassword() As ActionResult
			Return View()
		End Function

		<HttpPost()>
		Function ForgotPassword_Submit(value As FormCollection) As ActionResult
			Dim protocol As String = "http"
			Dim domainName As String
			Dim websiteURL As String
			Dim sMessage As String

			' run the sp's through the object
			Dim objResetPwd As New HR.Intranet.Server.clsResetPassword

			objResetPwd.Database = WebConfigurationManager.AppSettings("LoginPage:Database")
			objResetPwd.ServerName = WebConfigurationManager.AppSettings("LoginPage:Server")
			objResetPwd.Username = Request.Form("txtUserName")

			' Force password change only if there are no other users logged in with the same name.
			If Request.ServerVariables("HTTPS").ToLower <> "off" Then protocol = "https"
			domainName = Request.ServerVariables("HTTP_HOST")

			websiteURL = protocol & "://" & domainName & Url.Action("ResetPassword", "Account")	'Even though VS complains that it "Cannot resolve action 'ResetPassword'", it DOES resolve it!
			sMessage = objResetPwd.GenerateLinkAndEmail(websiteURL, Now())

			ViewData("RedirectToURLMessage") = "Go back"
			ViewData("RedirectToURL") = Url.Action("ForgotPassword", "Account")

			If Err.Number = 0 Then
				objResetPwd = Nothing

				' handle response from server...
				If Trim(sMessage) = "" Then
					' if OK...
					ViewData("Message") = "An e-mail has been sent to you. When you receive it, follow the directions in the email to reset your password."
					ViewData("RedirectToURLMessage") = "Login page"
					ViewData("RedirectToURL") = Url.Action("Login", "Account")
				Else
					' failure message from dll...
					ViewData("Message") = "You can not reset your password at this time.<br/><br/>" & sMessage
				End If
			Else
				ViewData("Message") = "You cannot reset your password at this time. <br/><br/>Intranet specifics have not been configured. <br/><br/>Please contact your system administrator."
			End If

			Return View()
		End Function

		Function ResetPassword() As ActionResult
			Return View()
		End Function

		<HttpPost()>
		Function ResetPassword_Submit(value As FormCollection) As ActionResult
			Dim Password As String = Request.Form("txtPassword1")
			Dim QueryString As String = Request.Form("txtQueryString")
			Dim Message As String
			Dim objResetPwd As New HR.Intranet.Server.clsResetPassword

			objResetPwd.Database = WebConfigurationManager.AppSettings("LoginPage:Database")
			objResetPwd.ServerName = WebConfigurationManager.AppSettings("LoginPage:Server")

			' Force password change only if there are no other users logged in with the same name.
			Message = objResetPwd.ResetPassword(QueryString, Password)
			objResetPwd = Nothing

			If UCase(Message) = UCase("Password changed successfully") Then
				' if OK...
				ViewData("Message") = "Your password has been reset successfully."
			Else
				' failure message from dll...	    
				ViewData("Message") = "You could not change your password at this time.<br/><br/>" & Message
			End If

			Return View()
		End Function
	End Class

	Public Class JsonData_DBValue
		Public Property Caption() As String
			Get
				Return m_Caption
			End Get
			Set(value As String)
				m_Caption = value
			End Set
		End Property

		Private m_Caption As String

		Public Property DBValue() As String
			Get
				Return m_DBValue
			End Get
			Set(value As String)
				m_DBValue = value
			End Set
		End Property

		Private m_DBValue As String

		Public Property Formatting_Suffix() As String
			Get
				Return m_Formatting_Suffix
			End Get
			Set(value As String)
				m_Formatting_Suffix = value
			End Set
		End Property

		Private m_Formatting_Suffix As String

		Public Property Formatting_Prefix() As String
			Get
				Return m_Formatting_Prefix
			End Get
			Set(value As String)
				m_Formatting_Prefix = value
			End Set
		End Property

		Private m_Formatting_Prefix As String
	End Class

	Public Class navigationLinks
		Private m_linkType As Integer
		Private m_linkOrder As Integer
		Private m_prompt As String
		Private m_text As String
		Private m_element_Type As Integer
		Private m_ID As Long

		Sub New(p_linkType As Integer, p_linkOrder As Integer, p_prompt As String, p_text As String, p_element_Type As Integer,
			p_ID As Long)

			linkType = p_linkType
			linkOrder = p_linkOrder
			prompt = p_prompt
			text = p_text
			element_Type = p_element_Type
			ID = p_ID
		End Sub


		Public Property linkType() As Integer
			Get
				Return m_linkType
			End Get
			Set(value As Integer)
				m_linkType = value
			End Set
		End Property

		Public Property linkOrder() As Integer
			Get
				Return m_linkOrder
			End Get
			Set(value As Integer)
				m_linkOrder = value
			End Set
		End Property

		Public Property prompt() As String
			Get
				Return m_prompt
			End Get
			Set(value As String)
				m_prompt = value
			End Set
		End Property

		Public Property text As String
			Get
				Return m_text
			End Get
			Set(value As String)
				m_text = value
			End Set
		End Property

		Public Property element_Type() As Integer
			Get
				Return m_element_Type
			End Get
			Set(value As Integer)
				m_element_Type = value
			End Set
		End Property

		Public Property ID() As Long
			Get
				Return m_ID
			End Get
			Set(value As Long)
				m_ID = value
			End Set
		End Property
	End Class
End Namespace

