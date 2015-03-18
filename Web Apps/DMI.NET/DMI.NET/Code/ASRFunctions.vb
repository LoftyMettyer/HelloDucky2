Option Strict On
Option Explicit On

Imports DMI.NET.Code
Imports DMI.NET.Classes
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server
Imports System.Data.SqlClient
Imports System.Web.WebPages

Public Module ASRFunctions

	Public Function GetCurrentUsersCountOnServer(LoginName As String) As Integer

		Dim objDataAccess = New clsDataAccess(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)
		Dim iLoginCount = 0

		Dim drUsers = objDataAccess.GetFromSP("spASRGetCurrentUsers")
		For Each objRow As DataRow In drUsers.Rows
			If objRow("loginame").ToString().Trim() = LoginName.Trim() Then
				iLoginCount += 1
			End If
		Next

		Return iLoginCount

	End Function

	Public Sub PopulateWorkflowSessionVariables()

		Dim objDataAccess As clsDataAccess = CType(HttpContext.Current.Session("DatabaseAccess"), clsDataAccess)
		Dim prmWFEnabled = New SqlParameter("pfWFEnabled", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
		Dim prmWFOutOfOfficeConfig = New SqlParameter("pfOutOfOfficeConfigured", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
		Dim bWorkflowEnabled As Boolean = False
		Dim fWorkflowOutOfOfficeConfigured = False
		Dim fWorkflowOutOfOffice = False
		Dim iWorkflowRecordCount = 0

		Try

			bWorkflowEnabled = Licence.IsModuleLicenced(SoftwareModule.Workflow)
			HttpContext.Current.Session("WF_Enabled") = bWorkflowEnabled

			' Check if the OutOfOffice parameters have been configured.
			If bWorkflowEnabled Then

				objDataAccess.ExecuteSP("spASRWorkflowOutOfOfficeConfigured", prmWFOutOfOfficeConfig)

				fWorkflowOutOfOfficeConfigured = CBool(prmWFOutOfOfficeConfig.Value)

				If fWorkflowOutOfOfficeConfigured Then
					' Check if the current user OutOfOffice
					Dim prmOutOfOffice As SqlParameter = New SqlParameter("pfOutOfOffice", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
					Dim prmRecordCount As SqlParameter = New SqlParameter("piRecordCount", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

					objDataAccess.ExecuteSP("spASRWorkflowOutOfOfficeCheck", prmOutOfOffice, prmRecordCount)

					fWorkflowOutOfOffice = CBool(prmOutOfOffice.Value)
					iWorkflowRecordCount = CInt(prmRecordCount.Value)

				End If
			End If

			HttpContext.Current.Session("WF_OutOfOfficeConfigured") = fWorkflowOutOfOfficeConfigured

		Catch ex As Exception
			Throw

		End Try



	End Sub

	Public Sub PopulateTrainingBookingSessionVariables()

		Try

			Dim objDataAccess As clsDataAccess = CType(HttpContext.Current.Session("DatabaseAccess"), clsDataAccess)

			Dim prmEmpTableID = New SqlParameter("piEmployeeTableID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmCourseTableID = New SqlParameter("piCourseTableID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmCourseCancelDateColumnID = New SqlParameter("piCourseCancelDateColumnID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmTBTableID = New SqlParameter("piTBTableID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmTBTableSelect = New SqlParameter("pfTBTableSelect", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmTBTableInsert = New SqlParameter("pfTBTableInsert", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmTBTableUpdate = New SqlParameter("pfTBTableUpdate", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmTBStatusColumnID = New SqlParameter("piTBStatusColumnID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmTBStatusColumnUpdate = New SqlParameter("pfTBStatusColumnUpdate", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmTBCancelDateColumnID = New SqlParameter("piTBCancelDateColumnID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmTBCancelDateColumnUpdate = New SqlParameter("pfTBCancelDateColumnUpdate", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmTBStatusPExists = New SqlParameter("pfTBProvisionalStatusExists", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmWaitListTableID = New SqlParameter("piWaitListTableID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmWaitListTableInsert = New SqlParameter("pfWaitListTableInsert", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmWaitListTableDelete = New SqlParameter("pfWaitListTableDelete", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmWaitListCourseTitleColumnID = New SqlParameter("piWaitListCourseTitleColumnID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmWaitListCourseTitleColumnUpdate = New SqlParameter("pfWaitListCourseTitleColumnUpdate", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmWaitListCourseTitleColumnSelect = New SqlParameter("pfWaitListCourseTitleColumnSelect", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmBulkBookingDefaultViewID = New SqlParameter("piBulkBookingDefaultViewID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmBulkBookingDefaultOrderID = New SqlParameter("piBulkBookingDefaultOrderID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

			objDataAccess.ExecuteSP("sp_ASRIntGetTrainingBookingParameters", _
					prmEmpTableID, prmCourseTableID, prmCourseCancelDateColumnID, prmTBTableID, prmTBTableSelect, prmTBTableInsert, prmTBTableUpdate, _
					prmTBStatusColumnID, prmTBStatusColumnUpdate, prmTBCancelDateColumnID, prmTBCancelDateColumnUpdate, prmTBStatusPExists, prmWaitListTableID, _
					prmWaitListTableInsert, prmWaitListTableDelete, prmWaitListCourseTitleColumnID, prmWaitListCourseTitleColumnUpdate, prmWaitListCourseTitleColumnSelect, _
					prmBulkBookingDefaultViewID, prmBulkBookingDefaultOrderID)

			HttpContext.Current.Session("TB_EmpTableID") = prmEmpTableID.Value
			HttpContext.Current.Session("TB_CourseTableID") = prmCourseTableID.Value
			HttpContext.Current.Session("TB_CourseCancelDateColumnID") = prmCourseCancelDateColumnID.Value
			HttpContext.Current.Session("TB_TBTableID") = prmTBTableID.Value
			HttpContext.Current.Session("TB_TBTableSelect") = prmTBTableSelect.Value
			HttpContext.Current.Session("TB_TBTableInsert") = prmTBTableInsert.Value
			HttpContext.Current.Session("TB_TBTableUpdate") = prmTBTableUpdate.Value
			HttpContext.Current.Session("TB_TBStatusColumnID") = prmTBStatusColumnID.Value
			HttpContext.Current.Session("TB_TBStatusColumnUpdate") = prmTBStatusColumnUpdate.Value
			HttpContext.Current.Session("TB_TBCancelDateColumnID") = prmTBCancelDateColumnID.Value
			HttpContext.Current.Session("TB_TBCancelDateColumnUpdate") = prmTBCancelDateColumnUpdate.Value
			HttpContext.Current.Session("TB_TBStatusPExists") = prmTBStatusPExists.Value
			HttpContext.Current.Session("TB_WaitListTableID") = prmWaitListTableID.Value
			HttpContext.Current.Session("TB_WaitListTableInsert") = prmWaitListTableInsert.Value
			HttpContext.Current.Session("TB_WaitListTableDelete") = prmWaitListTableDelete.Value
			HttpContext.Current.Session("TB_WaitListCourseTitleColumnID") = prmWaitListCourseTitleColumnID.Value
			HttpContext.Current.Session("TB_WaitListCourseTitleColumnUpdate") = prmWaitListCourseTitleColumnUpdate.Value
			HttpContext.Current.Session("TB_WaitListCourseTitleColumnSelect") = prmWaitListCourseTitleColumnSelect.Value

			HttpContext.Current.Session("TB_BulkBookingDefaultViewID") = prmBulkBookingDefaultViewID.Value
			HttpContext.Current.Session("TB_BulkBookingDefaultOrderID") = prmBulkBookingDefaultOrderID.Value

			If CStr(HttpContext.Current.Session("TB_TBTableID")) = "" Then HttpContext.Current.Session("TB_TBTableID") = 0

		Catch ex As Exception
			Throw

		End Try

	End Sub

	Public Function CalculatePromptedDate(objRow As DataRow) As Date

		Dim iPromptDateType As PromptedDateType
		Dim iDay As Integer
		Dim dtDate As Date
		Dim iMonth As Integer

		If (IsDBNull(objRow("promptDateType"))) Or (objRow("promptDateType").ToString() = vbNullString) Then
			iPromptDateType = 0
		Else
			iPromptDateType = CType(objRow("promptDateType"), PromptedDateType)
		End If

		Select Case iPromptDateType
			Case PromptedDateType.Explicit
				If Not IsDBNull(objRow("valuedate")) Then
					If (CStr(objRow("valuedate")) <> "00:00:00") And _
							(CStr(objRow("valuedate")) <> "12:00:00 AM") Then
						Return CDate(objRow("valuedate"))

					End If
				End If

			Case PromptedDateType.Current
				Return Date.Now

			Case PromptedDateType.MonthStart
				iDay = (Day(Date.Now) * -1) + 1
				dtDate = DateAdd("d", iDay, Date.Now)
				Return dtDate

			Case PromptedDateType.MonthEnd
				iDay = (Day(Date.Now) * -1) + 1
				dtDate = DateAdd("d", iDay, Date.Now)
				dtDate = DateAdd("m", 1, dtDate)
				dtDate = DateAdd("d", -1, dtDate)
				Return dtDate

			Case PromptedDateType.YearStart
				iDay = (Day(Date.Now) * -1) + 1
				iMonth = (Month(Date.Now) * -1) + 1
				dtDate = DateAdd("d", iDay, Date.Now)
				dtDate = DateAdd("m", iMonth, dtDate)
				Return dtDate

			Case PromptedDateType.YearEnd
				iDay = (Day(Date.Now) * -1) + 1
				iMonth = (Month(Date.Now) * -1) + 1
				dtDate = DateAdd("d", iDay, Date.Now)
				dtDate = DateAdd("m", iMonth, dtDate)
				dtDate = DateAdd("yyyy", 1, dtDate)
				dtDate = DateAdd("d", -1, dtDate)
				Return dtDate

		End Select

		Return Date.Now

	End Function

	Public Function FormatEventDuration(lngSeconds As Integer) As String

		Dim strHours As String
		Dim strMins As String
		Dim strSeconds As String
		Dim dblRemainder As Double

		Const TIME_SEPARATOR As String = ":"

		If Not (lngSeconds < 0) Then
			strHours = CStr(Fix(lngSeconds / 3600))
			strHours = New String(CType("0", Char), 2 - Len(strHours)) & strHours
			dblRemainder = CDbl(lngSeconds Mod 3600)

			strMins = CStr(Fix(dblRemainder / 60))
			strMins = New String(CType("0", Char), 2 - Len(strMins)) & strMins
			'UPGRADE_WARNING: Mod has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			dblRemainder = CDbl(dblRemainder Mod 60)

			strSeconds = CStr(Fix(dblRemainder))
			strSeconds = New String(CType("0", Char), 2 - Len(strSeconds)) & strSeconds

			Return strHours & TIME_SEPARATOR & strMins & TIME_SEPARATOR & strSeconds
		Else
			Return ""
		End If

	End Function

	Public Function Max(Of T As IComparable)(value1 As T, value2 As T) As T
		Return If(value1.CompareTo(value2) > 0, value1, value2)
	End Function

	Public Function Min(Of T As IComparable)(value1 As T, value2 As T) As T
		Return If(value1.CompareTo(value2) < 0, value1, value2)
	End Function

	Public Function ShowOutOfOffice(tableID As Integer, viewID As Integer) As Boolean

		' Are we displaying the Workflow Out of Office Hyperlink for this view?
		Dim fShowOooHyperlink As Boolean = False
		Dim objDataAccess As clsDataAccess = CType(HttpContext.Current.Session("DatabaseAccess"), clsDataAccess)

		Dim prmTableID2 = New SqlParameter("piTableID", SqlDbType.Int)
		prmTableID2.Value = tableID

		Dim prmViewID2 = New SqlParameter("piViewID", SqlDbType.Int)
		prmViewID2.Value = viewID

		Dim prmDisplayHyperlink = New SqlParameter("pfDisplayHyperlink", SqlDbType.Bit)
		prmDisplayHyperlink.Direction = ParameterDirection.Output
		Try
			objDataAccess.ExecuteSP("spASRIntShowOutOfOfficeHyperlink", prmTableID2, prmViewID2, prmDisplayHyperlink)
			fShowOooHyperlink = CBool(prmDisplayHyperlink.Value)
		Catch ex As Exception
			fShowOooHyperlink = False
		End Try

		Return fShowOooHyperlink

	End Function

	Public Function ValidateStringValue(inputValue As String, sanitiseLevel As InputValidation.StringSanitiseLevel) As String
		If inputValue Is Nothing OrElse inputValue.ToString() = vbNullString OrElse inputValue.ToString() = "undefined" Then
			Return inputValue
		End If

		Select Case sanitiseLevel
			Case InputValidation.StringSanitiseLevel.HTMLEncode
				inputValue = HttpUtility.HtmlAttributeEncode(inputValue)
			Case InputValidation.StringSanitiseLevel.None 'Don't do any sanitisation of the string
			Case InputValidation.StringSanitiseLevel.FullOWASP
				'inputValue = Globals.AntiSamyInstance.scan(inputValue, Globals.AntiSamyPolicy).getCleanHTML
		End Select

		Return inputValue
	End Function
	Public Function ValidateIntegerValue(inputValue As String) As Integer
		If inputValue Is Nothing OrElse inputValue.ToString() = vbNullString OrElse inputValue.ToString() = "undefined" Then
			Return 0 'By convention return 0; is this correct?
		End If

		Dim number As Integer
		Dim result As Boolean = Int32.TryParse(inputValue.ToString(), number)
		If Not result Then
			Throw New Exception()
		End If

		Return number
	End Function
	Public Function ValidateBooleanValue(inputValue As String) As Boolean
		If inputValue Is Nothing OrElse inputValue.ToString() = vbNullString OrElse inputValue.ToString() = "undefined" Then
			Return Nothing
		End If

		Dim bool As Boolean
		Dim result As Boolean = Boolean.TryParse(inputValue.ToString(), bool)
		If Not result Then
			Throw New Exception()
		End If

		Return bool
	End Function
	Public Function ValidateFromWhiteList(inputValue As String, whiteList As InputValidation.WhiteListCollections) As String
		If inputValue Is Nothing OrElse inputValue.ToString() = vbNullString OrElse inputValue.ToString() = "undefined" Then
			Return inputValue
		End If

		Select Case whiteList
			Case InputValidation.WhiteListCollections.Actions
				If Not InputValidation.ListOfActions.Contains(inputValue.ToString().ToUpper()) Then
					Throw New Exception()
				End If
			Case InputValidation.WhiteListCollections.UtilTypes
				Dim fOK As Boolean = False

				' Numeric parameter is OK
				Dim number As Integer
				fOK = Int32.TryParse(inputValue.ToString(), number)

				' and specific utiltypes are ok
				If InputValidation.ListOfUtilTypes.Contains(inputValue.ToString().ToUpper()) Then fOK = True

				If Not fOK Then Throw New Exception()
			Case InputValidation.WhiteListCollections.CT_Modes
				If Not InputValidation.ListOfCT_Modes.Contains(inputValue.ToString().ToUpper()) Then
					Throw New Exception()
				End If
		End Select

		Return inputValue
	End Function
	Public Function ValidateLineageValue(inputValue As String) As String
		If inputValue Is Nothing OrElse inputValue.ToString() = vbNullString OrElse inputValue.ToString() = "undefined" Then
			Return inputValue
		End If

		'Allow integers and underscores only
		Dim arrInput As String() = inputValue.ToString().Split(CChar("_"))
		For Each inputchar As String In arrInput
			If Not (inputchar.IsInt() Or inputchar.Contains(":")) Then
				Throw New Exception()
			End If
		Next

		Return inputValue
	End Function
End Module
