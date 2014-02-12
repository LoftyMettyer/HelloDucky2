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

End Module
