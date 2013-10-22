Option Strict Off
Option Explicit On
Imports ADODB

Public Class clsResetPassword
	Public gsDatabase As String
	Public gsServerName As String
	Public gsUsername As String
	Public mADOCon As Connection


	' We don't use gADOCon in this module as we're obviously not logged in when resetting a password.
	' so, run the sp to count users on server using fixed credentials.

	Public Function GetCurrentUsersCountOnServer() As Integer
		mADOCon = New Connection
		Dim cmdCheckUserSessions As New Command

		mADOCon.ConnectionString = ConnectionString()
		mADOCon.Open()

		cmdCheckUserSessions.CommandText = "spASRGetCurrentUsersCountOnServer"
		cmdCheckUserSessions.CommandType = CommandTypeEnum.adCmdStoredProc

		Dim prmCount As Parameter = cmdCheckUserSessions.CreateParameter("count", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamOutput)
		cmdCheckUserSessions.Parameters.Append(prmCount)

		Dim prmUserName As Parameter = cmdCheckUserSessions.CreateParameter("userName", DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamInput, 255)
		cmdCheckUserSessions.Parameters.Append(prmUserName)
		prmUserName.Value = gsUsername

		Err.Clear()

		cmdCheckUserSessions.ActiveConnection = mADOCon
		cmdCheckUserSessions.Execute()

		Return CInt(cmdCheckUserSessions.Parameters("count").Value)
	End Function
	Public Function GenerateLinkAndEmail(WebsiteURL As String, Timestamp As Date) As String
		Dim objCrypt As New clsCrypt
		Dim sEncryptedString As String
		Dim cmdResetPassword As Command
		Dim skey As String
		Dim sSourceString As String

		Const ENCRYPTIONKEY As String = "jmltn"

		skey = ENCRYPTIONKEY

		sSourceString = Username & vbTab & CStr(Timestamp) & vbTab & ServerName & vbTab & Database

		sEncryptedString = objCrypt.EncryptString(sSourceString, skey, True)
		sEncryptedString = objCrypt.CompactString(sEncryptedString)

		' Got the string, now send it to SQL with the timestamp and WebsiteURL
		' timestamp is used to ensure old links aren't reused
		' WebsiteURL used to generate the e-mail link to the resetpassword page.

		mADOCon = New Connection
		cmdResetPassword = New Command

		mADOCon.ConnectionString = ConnectionString()
		mADOCon.Open()

		cmdResetPassword.CommandText = "spadmin_resetpassword"
		cmdResetPassword.CommandType = CommandTypeEnum.adCmdStoredProc

		Dim prmWebsiteURL As Parameter = cmdResetPassword.CreateParameter("WebsiteURL", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 255)
		cmdResetPassword.Parameters.Append(prmWebsiteURL)
		prmWebsiteURL.Value = WebsiteURL

		Dim prmUserName As Parameter = cmdResetPassword.CreateParameter("userName", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 255)
		cmdResetPassword.Parameters.Append(prmUserName)
		prmUserName.Value = gsUsername

		Dim prmEncryptedLink As Parameter = cmdResetPassword.CreateParameter("encryptedLink", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 2147483646)
		cmdResetPassword.Parameters.Append(prmEncryptedLink)
		prmEncryptedLink.Value = sEncryptedString

		Dim prmResult As Parameter = cmdResetPassword.CreateParameter("message", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 2147483646)
		cmdResetPassword.Parameters.Append(prmResult)

		Err.Clear()

		cmdResetPassword.ActiveConnection = mADOCon

		cmdResetPassword.Execute()

		Return CStr(cmdResetPassword.Parameters("message").Value)
		' E-mail has been sent. When you receive your sign in information, follow the directions in the email to reset your password.
	End Function

	Public Function GetUsernameFromQueryString(QueryString As String) As String

		Dim objCrypt As New clsCrypt
		Dim Value As String
		Dim values() As String

		Dim tmpInstanceID As Short
		Dim tmpElementID As Short
		Dim tmpUser As String
		Dim tmpDateTimeStamp As String
		Dim tmpServer As String
		Dim tmpDatabase As String
		Dim tmpUsername As String

		On Error GoTo GetUsernameFromQueryString_ERROR

		Value = objCrypt.DecompactString(QueryString)
		If Value = "" Then GoTo GetUsernameFromQueryString_ERROR

		Value = objCrypt.DecryptString(Value, "", True)
		If Value = "" Then GoTo GetUsernameFromQueryString_ERROR

		values = Split(Value, vbTab)

		tmpUser = values(0)
		tmpDateTimeStamp = values(1)
		tmpServer = values(2)
		tmpDatabase = values(3)

		' Only links less than 4 hours old are valid.
		If DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(tmpDateTimeStamp), Now) < 241 Then
			GetUsernameFromQueryString = tmpUser
			Exit Function
		End If

GetUsernameFromQueryString_ERROR:

		GetUsernameFromQueryString = ""

	End Function

	Public Function ResetPassword(psEncryptedKey As String, psNewPassword As String) As String

		Dim cmdResetPassword As Command
		' using sp_password cos everything else does. When this sp is removed,
		' we can then move everything up to the replacement together.

		mADOCon = New Connection
		cmdResetPassword = New Command

		mADOCon.ConnectionString = ConnectionString()
		mADOCon.Open()

		cmdResetPassword.CommandText = "spadmin_commitresetpassword"
		cmdResetPassword.CommandType = CommandTypeEnum.adCmdStoredProc

		Dim prmEncryptedKey As Parameter = cmdResetPassword.CreateParameter("EncryptedKey", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 4000)
		cmdResetPassword.Parameters.Append(prmEncryptedKey)
		prmEncryptedKey.Value = psEncryptedKey

		Dim prmNewPwd As Parameter = cmdResetPassword.CreateParameter("newPassword", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 4000)
		cmdResetPassword.Parameters.Append(prmNewPwd)
		prmNewPwd.Value = psNewPassword

		Dim prmMessage As Parameter = cmdResetPassword.CreateParameter("message", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 4000)
		cmdResetPassword.Parameters.Append(prmMessage)

		Err.Clear()

		cmdResetPassword.ActiveConnection = mADOCon

		cmdResetPassword.Execute()

		Return cmdResetPassword.Parameters("message").Value.ToString
	End Function

	Private Function ConnectionString() As String

		Dim objSettings As New clsSettings

		ConnectionString = objSettings.GetSQLProviderString & "Data Source=" & ServerName & ";Initial Catalog=" & Database & ";Application Name=OpenHR Self-service Intranet;DataTypeCompatibility=80;" & ";User ID=" & fixedUsername & ";Password=" & fixedPassword & ";Persist Security Info=True;"

	End Function

	' Paul says it's fine to hard code these in a vb6 dll!!
	Private ReadOnly Property fixedUsername() As String
		Get
			fixedUsername = "openhr2iis"
		End Get
	End Property

	Private ReadOnly Property fixedPassword() As String
		Get
			fixedPassword = "H@Rp3Nd3N"
		End Get
	End Property

	Public Property Username() As String
		Get
			Username = gsUsername
		End Get
		Set(ByVal Value As String)
			' Username passed in from the asp page
			gsUsername = Value
		End Set
	End Property

	Public Property Database() As String
		Get
			Database = gsDatabase
		End Get
		Set(ByVal Value As String)
			' Database passed in from the asp page
			gsDatabase = Value
		End Set
	End Property

	Public Property ServerName() As String
		Get
			ServerName = gsServerName
		End Get
		Set(ByVal Value As String)
			' ServerName passed in from the asp page
			gsServerName = Value
		End Set
	End Property

	Public ReadOnly Property minPasswordLength() As Integer
		Get

			minPasswordLength = 1

			' TODO:
			'exec dbo.sp_ASRIntGetSystemSetting 'password', 'minimum length', 'minimumPasswordLength',
			'    @sValue OUTPUT, @fNewSettingFound OUTPUT, @fOldSettingFound OUTPUT
			'return  convert(integer, @sValue)


		End Get
	End Property
End Class