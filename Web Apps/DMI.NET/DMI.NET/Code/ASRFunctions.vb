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
End Module
