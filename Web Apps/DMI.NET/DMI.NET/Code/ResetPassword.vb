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

				sSourceString = Username & vbTab & CStr(Timestamp) & vbTab & _db.Server & vbTab & _db.Database

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
			Dim tmpDateTimeStamp As String

			Try

				Value = objCrypt.DecompactString(QueryString)
				If Value = "" Then Return ""

				Value = objCrypt.DecryptString(Value, "", True)
				If Value = "" Then Return ""

				values = Split(Value, vbTab)

				tmpUser = values(0)
				tmpDateTimeStamp = values(1)

				' Only links less than 4 hours old are valid.
				If DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(tmpDateTimeStamp), Now) < 241 Then
					Return tmpUser
				End If

				Return ""

			Catch ex As Exception
				Return ""

			End Try

		End Function

		Public Function ResetPassword(psEncryptedKey As String, psNewPassword As String) As String

			Try
				Dim prmMessage As New SqlParameter("ErrorMessage", SqlDbType.VarChar, 4000) With {.Direction = ParameterDirection.Output}

				_db.ExecuteSP("spadmin_commitresetpassword", _
											New SqlParameter("code", SqlDbType.NVarChar, 4000) With {.Value = psEncryptedKey}, _
											New SqlParameter("NewPassword", SqlDbType.NVarChar, 4000) With {.Value = psNewPassword}, _
											prmMessage)

				Return prmMessage.Value.ToString()

			Catch ex As Exception
				Return ""

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