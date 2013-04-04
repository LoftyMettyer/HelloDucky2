Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsResetPassword_NET.clsResetPassword")> Public Class clsResetPassword
	Public gsDatabase As String
	Public gsServerName As String
	Public gsUsername As String
	Public mADOCon As ADODB.Connection
	
	
	' We don't use gADOCon in this module as we're obviously not logged in when resetting a password.
	' so, run the sp to count users on server using fixed credentials.
	
	Public Function GetCurrentUsersCountOnServer() As Integer
		Dim cmdCheckUserSessions As Object
		
		mADOCon = CreateObject("ADODB.Connection")
		cmdCheckUserSessions = CreateObject("ADODB.Command")
		
		mADOCon.ConnectionString = ConnectionString()
		mADOCon.Open()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdCheckUserSessions.CommandText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cmdCheckUserSessions.CommandText = "spASRGetCurrentUsersCountOnServer"
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdCheckUserSessions.CommandType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cmdCheckUserSessions.CommandType = 4 ' adStoredProc
		
		Dim prmCount As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdCheckUserSessions.CreateParameter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		prmCount = cmdCheckUserSessions.CreateParameter("count", 3, 2) ' 3=integer, 2=output
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdCheckUserSessions.Parameters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cmdCheckUserSessions.Parameters.Append(prmCount)
		
		Dim prmUserName As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdCheckUserSessions.CreateParameter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		prmUserName = cmdCheckUserSessions.CreateParameter("userName", 200, 1, 255)
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdCheckUserSessions.Parameters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cmdCheckUserSessions.Parameters.Append(prmUserName)
		'UPGRADE_WARNING: Couldn't resolve default property of object prmUserName.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		prmUserName.Value = gsUsername
		
		Err.Clear()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdCheckUserSessions.ActiveConnection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cmdCheckUserSessions.ActiveConnection = mADOCon.ConnectionString
		
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdCheckUserSessions.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cmdCheckUserSessions.Execute()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdCheckUserSessions.Parameters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetCurrentUsersCountOnServer = CInt(cmdCheckUserSessions.Parameters("count").Value)
		
	End Function
	
	
	
	Public Function GenerateLinkAndEmail(ByRef WebsiteURL As Object, ByRef Timestamp As Object) As String
		Dim objCrypt As New clsCrypt
		Dim sEncryptedString As String
		Dim cmdResetPassword As Object
		Dim skey As String
		Dim sSourceString As String
		
		GenerateLinkAndEmail = CStr(False)
		
		Const ENCRYPTIONKEY As String = "jmltn"
		
		' sEncryptedString = objCrypt.EncryptQueryString(-1, -1, gsUsername, CStr(Timestamp), gsServerName, gsDatabase, mADOCon)
		skey = ENCRYPTIONKEY
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Timestamp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sSourceString = Username & vbTab & CStr(Timestamp) & vbTab & ServerName & vbTab & Database
		
		sEncryptedString = objCrypt.EncryptString(sSourceString, skey, True)
		sEncryptedString = objCrypt.CompactString(sEncryptedString)
		
		' Got the string, now send it to SQL with the timestamp and WebsiteURL
		' timestamp is used to ensure old links aren't reused
		' WebsiteURL used to generate the e-mail link to the resetpassword page.
		
		mADOCon = CreateObject("ADODB.Connection")
		cmdResetPassword = CreateObject("ADODB.Command")
		
		mADOCon.ConnectionString = ConnectionString()
		mADOCon.Open()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.CommandText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cmdResetPassword.CommandText = "spadmin_resetpassword"
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.CommandType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cmdResetPassword.CommandType = 4 ' adStoredProc
		
		Dim prmWebsiteURL As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.CreateParameter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		prmWebsiteURL = cmdResetPassword.CreateParameter("WebsiteURL", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 255)
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.Parameters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cmdResetPassword.Parameters.Append(prmWebsiteURL)
		'UPGRADE_WARNING: Couldn't resolve default property of object prmWebsiteURL.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object WebsiteURL. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		prmWebsiteURL.Value = WebsiteURL
		
		Dim prmUserName As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.CreateParameter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		prmUserName = cmdResetPassword.CreateParameter("userName", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 255)
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.Parameters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cmdResetPassword.Parameters.Append(prmUserName)
		'UPGRADE_WARNING: Couldn't resolve default property of object prmUserName.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		prmUserName.Value = gsUsername
		
		Dim prmEncryptedLink As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.CreateParameter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		prmEncryptedLink = cmdResetPassword.CreateParameter("encryptedLink", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 2147483646)
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.Parameters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cmdResetPassword.Parameters.Append(prmEncryptedLink)
		'UPGRADE_WARNING: Couldn't resolve default property of object prmEncryptedLink.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		prmEncryptedLink.Value = sEncryptedString
		
		Dim prmResult As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.CreateParameter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		prmResult = cmdResetPassword.CreateParameter("message", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 2147483646)
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.Parameters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cmdResetPassword.Parameters.Append(prmResult)
		
		Err.Clear()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.ActiveConnection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cmdResetPassword.ActiveConnection = mADOCon.ConnectionString
		
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cmdResetPassword.Execute()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.Parameters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GenerateLinkAndEmail = CStr(cmdResetPassword.Parameters("message").Value)
		
		
		' E-mail has been sent. When you receive your sign in information, follow the directions in the email to reset your password.
		
		
		
	End Function
	
	Public Function GetUsernameFromQueryString(ByRef QueryString As Object) As String
		
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
		
		'UPGRADE_WARNING: Couldn't resolve default property of object QueryString. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
		'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
		If DateDiff(Microsoft.VisualBasic.DateInterval.Minute, CDate(tmpDateTimeStamp), Now) < 241 Then
			GetUsernameFromQueryString = tmpUser
			Exit Function
		End If
		
GetUsernameFromQueryString_ERROR: 
		
		GetUsernameFromQueryString = ""
		
	End Function
	
	Public Function ResetPassword(ByRef psEncryptedKey As Object, ByRef psNewPassword As Object) As String
		
		Dim cmdResetPassword As Object
		' using sp_password cos everything else does. When this sp is removed,
		' we can then move everything up to the replacement together.
		
		mADOCon = CreateObject("ADODB.Connection")
		cmdResetPassword = CreateObject("ADODB.Command")
		
		mADOCon.ConnectionString = ConnectionString()
		mADOCon.Open()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.CommandText. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cmdResetPassword.CommandText = "spadmin_commitresetpassword"
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.CommandType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cmdResetPassword.CommandType = 4 ' adStoredProc
		
		Dim prmEncryptedKey As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.CreateParameter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		prmEncryptedKey = cmdResetPassword.CreateParameter("EncryptedKey", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 4000)
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.Parameters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cmdResetPassword.Parameters.Append(prmEncryptedKey)
		'UPGRADE_WARNING: Couldn't resolve default property of object prmEncryptedKey.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object psEncryptedKey. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		prmEncryptedKey.Value = psEncryptedKey
		
		Dim prmNewPwd As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.CreateParameter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		prmNewPwd = cmdResetPassword.CreateParameter("newPassword", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 4000)
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.Parameters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cmdResetPassword.Parameters.Append(prmNewPwd)
		'UPGRADE_WARNING: Couldn't resolve default property of object prmNewPwd.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object psNewPassword. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		prmNewPwd.Value = psNewPassword
		
		Dim prmMessage As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.CreateParameter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		prmMessage = cmdResetPassword.CreateParameter("message", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 4000)
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.Parameters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cmdResetPassword.Parameters.Append(prmMessage)
		
		Err.Clear()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.ActiveConnection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cmdResetPassword.ActiveConnection = mADOCon.ConnectionString
		
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.Execute. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cmdResetPassword.Execute()
		
		If Err.Number <> 0 Then
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object cmdResetPassword.Parameters. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ResetPassword = cmdResetPassword.Parameters("message").Value
		
		
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