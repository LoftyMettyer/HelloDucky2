Option Strict On
Option Explicit On

Imports System.Data.SqlClient
Imports HR.Intranet.Server

Namespace Code

	Public Class ResetPassword

		Private ReadOnly _db As clsDataAccess
		'	Private ReadOnly _objLogin As New LoginInfo

		Public Sub New()

			MyBase.New()
			_db = New clsDataAccess(ConfigurationManager.ConnectionStrings("OpenHR").ConnectionString)

		End Sub

		Public Function GenerateLinkAndEmail(WebsiteURL As String, Timestamp As Date) As String
			Dim objCrypt As New clsCrypt
			Dim sEncryptedString As String
			Dim skey As String
			Dim sSourceString As String

			Const ENCRYPTIONKEY As String = "jmltn"

			Try

				skey = ENCRYPTIONKEY

				sSourceString = Username & vbTab & Timestamp.ToOADate() & vbTab & "" & vbTab & ""

				sEncryptedString = objCrypt.EncryptString(sSourceString, skey, True)
				sEncryptedString = objCrypt.CompactString(sEncryptedString)

				' Got the string, now send it to SQL with the timestamp and WebsiteURL
				' timestamp is used to ensure old links aren't reused
				' WebsiteURL used to generate the e-mail link to the resetpassword page.

				Dim prmMessage As New SqlParameter("psMessage", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.InputOutput, .Value = ""}

				_db.ExecuteSP("spadmin_resetpassword", _
											New SqlParameter("psWebsiteURL", SqlDbType.VarChar, 255) With {.Value = WebsiteURL}, _
											New SqlParameter("psUserName", SqlDbType.VarChar, 255) With {.Value = Username}, _
											New SqlParameter("psEncryptedLink", SqlDbType.VarChar, -1) With {.Value = sEncryptedString}, _
											prmMessage)

				Return prmMessage.Value.ToString()
				' E-mail has been sent. When you receive your sign in information, follow the directions in the email to reset your password.

			Catch ex As Exception
				Throw

			End Try

		End Function

		Public Function GetUsernameFromQueryString(QueryString As String) As String

			Dim objCrypt As New clsCrypt
			Dim Value As String
			Dim values() As String

			Dim tmpUser As String
			Dim tmpDateTimeStamp As Double

			Try

				Value = objCrypt.DecompactString(QueryString)
				If Value = "" Then Return ""

				Value = objCrypt.DecryptString(Value, "", True)
				If Value = "" Then Return ""

				values = Split(Value, vbTab)

				tmpUser = values(0)
				tmpDateTimeStamp = CDbl(values(1))

				' Only links less than 4 hours old are valid.
				Dim timestamp = Date.FromOADate(tmpDateTimeStamp)
				If DateDiff(DateInterval.Minute, timestamp, Now) < 241 Then
					Return tmpUser
				End If

				Return ""

			Catch ex As Exception
				Return ""

			End Try

		End Function

		Public Function ResetPassword(encryptedKey As String, newPassword As String) As String
			
			Dim sSQL As String
			Dim sUsername As String
			Dim values As String()

			Dim crypt As New clsCrypt
			encryptedKey = crypt.DecompactString(encryptedKey)
			encryptedKey = crypt.DecryptString(encryptedKey, "", True)

			'Extract the required parameters from the decrypted queryString.
			values = encryptedKey.Split(vbTab(0))
			sUsername = values(0)

			Try

            ' Verify the heartbeat has sufficient priviledges
            sSQL = "SELECT CASE WHEN IS_SRVROLEMEMBER('sysadmin') = 1 OR IS_SRVROLEMEMBER('securityadmin') = 1 THEN 1 ELSE 0 END"
				Dim bHasPermission = (CInt(_db.GetDataTable(sSQL).Rows(0)(0)) > 0)

            If Not bHasPermission Then
               Return "Password not reset (Error Code: CE007)"

            Else

				   sSQL = String.Format("SELECT COUNT(*) FROM master..sysprocesses p" & _
						   " WHERE    p.program_name LIKE 'OpenHR%'" & _
										   " AND p.program_name NOT LIKE 'OpenHR Workflow%'" & _
										   " AND p.program_name NOT LIKE 'OpenHR Outlook%'" & _
										   " AND p.program_name NOT LIKE 'OpenHR Server.Net%'" & _
										   " AND p.program_name NOT LIKE 'OpenHR Intranet Embedding%'" & _
										   " AND p.loginame = '{0}'", sUsername.Replace("'", "''"))
				   Dim bLoggedIn = (CInt(_db.GetDataTable(sSQL).Rows(0)(0)) > 0)

				   If Not bLoggedIn Then
					   sSQL = String.Format("IF EXISTS (SELECT * FROM sys.server_principals WHERE name = N'{0}')" & _
					   "ALTER LOGIN [{0}] WITH PASSWORD = '{1}'", sUsername, newPassword.Replace("'", "''"))

					   _db.ExecuteSql(sSQL)
					   Return "Password changed successfully"
				   Else
					   Return "User is currently logged in"
				   End If

            End If

			Catch ex As Exception
				Return ex.Message

			End Try

		End Function

		Public Property Username() As String

		Public ReadOnly Property minPasswordLength() As Integer
			Get
				minPasswordLength = 1
			End Get
		End Property

	End Class
End Namespace