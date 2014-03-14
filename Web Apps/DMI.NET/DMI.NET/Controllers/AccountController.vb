﻿Imports System.IO
Imports System.Drawing
Imports DMI.NET.Code
Imports HR.Intranet.Server
Imports System.Data.SqlClient
Imports System.Reflection
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
			Dim bWindowsAuthentication As Boolean = False

			fForcePasswordChange = False

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

			' HRPRO-3531
			Session("MSBrowser") = (Request.Form("txtMSBrowser") = "true")

			Dim objLogin As LoginInfo
			Dim objServerSession As New SessionInfo

			Dim objDatabase As New Database
			Dim objDataAccess As clsDataAccess

			Try

				' Validate the login
				objLogin = objServerSession.SessionLogin(sUserName, sPassword, sDatabaseName, sServerName, bWindowsAuthentication)

				' Has the password expired? Cannot log in until they've successfully changed it.
				If objLogin.MustChangePassword Then
					Return RedirectToAction("ForcedPasswordChange", "Account")
				End If

				' Generic login fail.
				If objLogin.LoginFailReason.Length <> 0 Then
					Session("ErrorText") = FormatError(objServerSession.LoginInfo.LoginFailReason)
					Return RedirectToAction("Loginerror")
				End If

				' Database update in progress
				If objServerSession.DatabaseStatus.IsUpdateInProgress Then
					Session("ErrorText") = "A database update is in progress."
					Return RedirectToAction("Loginerror")
				End If

				' Users that are assigned certain server roles cannot log in (I think its dodgy because we rely too heavily on dbo)
				If objLogin.IsServerRole Then
					Session("ErrorText") = "Users assigned to fixed SQL Server roles cannot use OpenHR web."
					FormatError(objServerSession.LoginInfo.LoginFailReason)
					Return RedirectToAction("Loginerror")
				End If

				' Is the DB the correct version
				Dim objAppVersion As Version = Assembly.GetExecutingAssembly().GetName().Version

				If Not CompareVersion(objServerSession.DatabaseStatus.IntranetVersion, objAppVersion, False) _
					Or Not CompareVersion(objServerSession.DatabaseStatus.SysMgrVersion, objAppVersion, True) Then
					Session("ErrorText") = String.Format("The database is out of date.<BR>Please ask the System Administrator to update the database for use with version {0}.{1}.{2}" _
							, objAppVersion.Major, objAppVersion.Minor, objAppVersion.Build)
					Return RedirectToAction("Loginerror")
				End If

				' Valid login, but do we have any kind of access?
				If Not objLogin.IsSSIUser And Not objLogin.IsDMIUser And Not objLogin.IsDMISingle Then
					Session("ErrorText") = "You are not permitted to use OpenHR Web with this user name."
					Return RedirectToAction("Loginerror")
				End If

				'Bounce non-IE users who have no SSI access...
				If Session("MSBrowser") = False And Not objLogin.IsSSIUser Then
					Session("ErrorText") = "You are not permitted to use OpenHR Self-service with this user name."
					Return RedirectToAction("Loginerror")
				End If

				' Track and audit that we've logged in
				objServerSession.TrackUser(True)

				' User is allowed into OpenHR, now populate some metadata
				objServerSession.RegionalSettings = Platform.GetRegionalSettings
				objDatabase.SessionInfo = objServerSession
				objServerSession.Initialise()


				objDataAccess = New clsDataAccess(objServerSession.LoginInfo)
				Session("DatabaseFunctions") = objDatabase
				Session("DatabaseAccess") = objDataAccess
				Session("sessionContext") = objServerSession

				' Get module parameters
				PopulatePersonnelSessionVariables()
				PopulateWorkflowSessionVariables()
				PopulateTrainingBookingSessionVariables()

				' Get parameters for the single record
				Dim prmTableID = New SqlParameter("piTableID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmViewID = New SqlParameter("piViewID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

				objDataAccess.ExecuteSP("spASRIntGetSingleRecordViewID", prmTableID, prmViewID)
				Session("SingleRecordTableID") = CInt(prmTableID.Value)
				Session("SingleRecordViewID") = CInt(prmViewID.Value)
				Session("SSILinkTableID") = 0
				Session("SSILinkViewID") = 0

				'Store in Session the logged in user's RecordID
				Session("LoggedInUserRecordID") = ASRIntranetFunctions.GetLoggedInUserRecordID(Session("SingleRecordViewID"))
			Catch ex As Exception
				Session("ErrorText") = FormatError(ex.Message)
				Return RedirectToAction("Loginerror")

			End Try

			' Are we displaying the Workflow Out of Office Hyperlink for this view?
			Dim fShowOOOHyperlink As Boolean = False

			Dim prmTableID2 = New SqlParameter("piTableID", SqlDbType.Int)
			prmTableID2.Value = Convert.ToInt16(Session("SingleRecordTableID"))

			Dim prmViewID2 = New SqlParameter("piViewID", SqlDbType.Int)
			prmViewID2.Value = Convert.ToInt16(Session("SingleRecordViewID"))

			Dim prmDisplayHyperlink = New SqlParameter("pfDisplayHyperlink", SqlDbType.Bit)
			prmDisplayHyperlink.Direction = ParameterDirection.Output
			Try
				objDataAccess.ExecuteSP("spASRIntShowOutOfOfficeHyperlink", prmTableID2, prmViewID2, prmDisplayHyperlink)
				fShowOOOHyperlink = prmDisplayHyperlink.Value
			Catch ex As Exception

			End Try
			Session("WF_ShowOutOfOffice") = fShowOOOHyperlink

			'
			Session("Server") = sServerName
			Session("Database") = sDatabaseName
			Session("WinAuth") = bWindowsAuthentication
			Session("UserGroup") = objServerSession.LoginInfo.UserGroup

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


			Dim lngSSIWelcomeColumnID = CLng(objDatabase.GetModuleParameter("MODULE_PERSONNEL", "Param_FieldsSSIWelcome"))
			If lngSSIWelcomeColumnID <= 0 Then lngSSIWelcomeColumnID = 0

			Dim lngSSIPhotographColumnID = CLng(objDatabase.GetModuleParameter("MODULE_PERSONNEL", "Param_FieldsSSIPhotograph"))
			If lngSSIPhotographColumnID <= 0 Then lngSSIPhotographColumnID = 0


			Try

				Dim prmWelcomeMessage = New SqlParameter("psWelcomeMessage", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmSelfServiceWelcomeColumn = New SqlParameter("psSelfServiceWelcomeColumn", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmSelfServicePhotograph = New SqlParameter("psSelfServicePhotograph", SqlDbType.VarBinary, -1) With {.Direction = ParameterDirection.Output}

				objDataAccess.ExecuteSP("spASRIntGetSSIWelcomeDetails" _
						, New SqlParameter("piWelcomeColumnID", SqlDbType.Int) With {.Value = lngSSIWelcomeColumnID} _
						, New SqlParameter("piPhotographColumnID", SqlDbType.Int) With {.Value = lngSSIPhotographColumnID} _
						, New SqlParameter("piSingleRecordViewID", SqlDbType.Int) With {.Value = Session("SingleRecordViewID")} _
						, New SqlParameter("psUserName", SqlDbType.VarChar, 255) With {.Value = Session("username")} _
						, prmWelcomeMessage _
						, prmSelfServiceWelcomeColumn _
						, prmSelfServicePhotograph)


				Session("welcomemessage") = prmWelcomeMessage.Value.ToString()
				Session("welcomeName") = prmSelfServiceWelcomeColumn.Value.ToString()
				If Not IsDBNull(prmSelfServicePhotograph.Value) Then
					Dim OLEType As Short = Val(Encoding.UTF8.GetString(prmSelfServicePhotograph.Value, 8, 2))
					If OLEType = 2 Then	'Embeded
						Dim abtImage = CType(prmSelfServicePhotograph.Value, Byte())
						Dim binaryData As Byte() = New Byte(abtImage.Length - 400) {}
						Try
							Buffer.BlockCopy(abtImage, 400, binaryData, 0, abtImage.Length - 400)
							'Create an image based on the embeded (Base64) image and resize it to 48x48
							Dim img As Image = Base64StringToImage(Convert.ToBase64String(binaryData, 0, binaryData.Length))
							img = img.GetThumbnailImage(48, 48, Nothing, IntPtr.Zero)
							Session("SelfServicePhotograph_Src") = "data:image/jpeg;base64," & ImageToBase64String(img)
						Catch exp As ArgumentNullException

						End Try
					ElseIf OLEType = 3 Then	'Link
						Dim UNC As String = Trim(Encoding.UTF8.GetString(prmSelfServicePhotograph.Value, 290, 60))
						Dim FileName As String = Trim(Path.GetFileName(Encoding.UTF8.GetString(prmSelfServicePhotograph.Value, 10, 70))).Replace("\", "/")
						Dim FullPath As String = Trim(Encoding.UTF8.GetString(prmSelfServicePhotograph.Value, 80, 210)).Replace("\", "/")
						Session("SelfServicePhotograph_src") = "file:///" & UNC & "/" & FullPath & "/" & FileName
					End If
				Else 'No picture is defined for user, use anonymous one
					Session("SelfServicePhotograph_Src") = Url.Content("~/Content/images/anonymous.png")
				End If

			Catch ex As Exception
				Session("welcomemessage") = "error: " & ex.Message & ". ID: " & CStr(lngSSIWelcomeColumnID)

			End Try

			Dim cookie = New HttpCookie("Login")
			cookie.Expires = DateTime.Now.AddYears(1)
			cookie.HttpOnly = True
			cookie("User") = Request.Form("txtUserNameCopy")
			'dont save or retrieve these anymore HRPRO-3030 / 3031
			'cookie("Database") = Request.Form("txtDatabase")
			'cookie("Server") = Request.Form("txtServer")
			cookie("WindowsAuthentication") = Request.Form("chkWindowsAuthentication")
			Response.Cookies.Add(cookie)

			If Session("AdminRequiresIE") = "TRUE" And Session("MSBrowser") <> True Then
				' non-IE browsers don't get DMI access yet.
				ViewBag.SSIMode = True
			Else

				If objLogin.IsDMIUser Or objLogin.IsDMISingle Then
					ViewBag.SSIMode = False
				Else
					ViewBag.SSIMode = True
				End If

			End If

			' always main.
			Return RedirectToAction("Main", "Home", New With {.SSIMode = ViewBag.SSIMode})

		End Function

		Public Function LogOff()
			Session("ErrorText") = Nothing

			Try
				Dim objServerSession As SessionInfo = Session("sessionContext")
				objServerSession.TrackUser(False)

				Session("avPrimaryMenuInfo") = Nothing
				Session("avSubMenuInfo") = Nothing
				Session("avQuickEntryMenuInfo") = Nothing
				Session("avTableMenuInfo") = Nothing
				Session("avTableHistoryMenuInfo") = Nothing

				objServerSession.ActiveConnections -= 1

			Catch ex As Exception

			End Try

			Return RedirectToAction("Login", "Account")

		End Function


		<HttpPost()>
		Function ForcedPasswordChange_Submit(value As FormCollection) As ActionResult

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			Dim fSubmitPasswordChange = (Len(Request.Form("txtGotoPage")) = 0)

			If fSubmitPasswordChange Then

				' Read the Password details from the Password form.
				Dim sNewPassword = Request.Form("txtPassword1")

				Try
					objDataAccess.Login.Password = Request.Form("txtCurrentPassword")
					clsDataAccess.ChangePassword(objDataAccess.Login, sNewPassword)
					objDataAccess.Login.Password = sNewPassword

					Session("MessageTitle") = "Change Password Page"
					Session("MessageText") = "Password changed successfully. You may now login."
					Return RedirectToAction("Loginmessage", "Account")

				Catch ex As SqlException
					Session("ErrorTitle") = "Change Password Page"

					Dim iOccurs = ex.Message.IndexOf("Changed database context", 0)
					Dim sMessage As String = ex.Message
					If iOccurs > 0 Then
						sMessage = ex.Message.Substring(0, iOccurs)
					End If
					Session("ErrorText") = sMessage
					Return RedirectToAction("ForcedPasswordChange", "Account")

				Catch ex As Exception
					Session("ErrorTitle") = "Change Password Page"
					Session("ErrorText") = ex.Message
					Return RedirectToAction("Loginerror", "Account")

				End Try

			End If

			' Go to the main page.
			Return RedirectToAction("Main", "Home")

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

			objResetPwd.Database = ApplicationSettings.LoginPage_Database
			objResetPwd.ServerName = ApplicationSettings.LoginPage_Server
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

			objResetPwd.Database = ApplicationSettings.LoginPage_Database
			objResetPwd.ServerName = ApplicationSettings.LoginPage_Server

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

