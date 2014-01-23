Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server
Imports System.Data.SqlClient

Public Module ASRFunctions
	Public Function GetCurrentUsersCountOnServer(LoginName As String) As Integer
		Dim objSession As SessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)	'Set session info
		Dim objDataAccess As New clsDataAccess(objSession.LoginInfo) 'Instantiate DataAccess class
		Dim iLoginCount As New SqlParameter("@iLoginCount", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

		objDataAccess.ExecuteSP(
					"spASRGetCurrentUsersCountOnServer", _
					iLoginCount, _
					New SqlParameter("@psLoginName", SqlDbType.VarChar, -1) With {.Value = LoginName} _
		)

		Return CInt(iLoginCount.Value)

	End Function

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


End Module
