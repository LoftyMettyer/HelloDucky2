<%@ Page Title="" Language="VB" MasterPageFile="~/Views/Shared/Site.Master" Inherits="System.Web.Mvc.ViewPage" %>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
ForcedPasswordChange_Submit
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

<%@ Language=VBScript %>
<!--#INCLUDE FILE="include/svrCleanup.asp" -->
<% 
	on error resume next

	dim strErrorMessage

	' Only process the form submission if the referring page was the newUser page.
	' If it wasn't then redirect to the login page.
	sReferringPage = Request.ServerVariables("HTTP_REFERER") 
	if inStrRev(sReferringPage, "/") > 0 then
		sReferringPage = mid(sReferringPage, inStrRev(sReferringPage, "/") + 1)
	end if

	if ucase(sReferringPage) <> ucase("forcedPasswordChange.asp") then
		Response.Redirect("login.asp")
	else

		fSubmitPasswordChange = (len(Request.Form("txtGotoPage")) = 0) 

		if fSubmitPasswordChange then
			' Force password change only if there are no other users logged in with the same name.
			Set cmdCheckUserSessions = Server.CreateObject("ADODB.Command")
			cmdCheckUserSessions.CommandText = "spASRGetCurrentUsersCountOnServer"
			cmdCheckUserSessions.CommandType = 4 ' Stored procedure.
			Set cmdCheckUserSessions.ActiveConnection = session("databaseConnection")

			Set prmCount = cmdCheckUserSessions.CreateParameter("count",3,2) ' 3=integer, 2=output
			cmdCheckUserSessions.Parameters.Append prmCount

			Set prmUserName = cmdCheckUserSessions.CreateParameter("userName",200,1,8000) ' 200=varchar, 1=input, 8000=size
			cmdCheckUserSessions.Parameters.Append prmUserName
			prmUserName.value = session("Username")

			err = 0
			cmdCheckUserSessions.Execute
			
			iUserSessionCount = clng(cmdCheckUserSessions.Parameters("count").Value)
			set cmdCheckUserSessions = nothing
			
			if iUserSessionCount < 2 then
				' Read the Password details from the Password form.
				sCurrentPassword = Request.Form("txtCurrentPassword")
				sNewPassword = Request.Form("txtPassword1")

				' Attempt to change the password on the SQL Server.
				Set cmdChangePassword = Server.CreateObject("ADODB.Command")
				cmdChangePassword.CommandText = "sp_password"
				cmdChangePassword.CommandType = 4 ' Stored Procedure
				Set cmdChangePassword.ActiveConnection = session("databaseConnection")

				Set prmCurrentPassword = cmdChangePassword.CreateParameter("currentPassword",200,1,255)
				cmdChangePassword.Parameters.Append prmCurrentPassword
				if len(sCurrentPassword) > 0 then
					prmCurrentPassword.value = sCurrentPassword
				else
					prmCurrentPassword.value = null				
				end if

				Set prmNewPassword = cmdChangePassword.CreateParameter("newPassword",200,1,255)
				cmdChangePassword.Parameters.Append prmNewPassword
				if len(sNewPassword) > 0 then
					prmNewPassword.value = sNewPassword
				else
					prmNewPassword.value = null				
				end if
	
				err = 0
				cmdChangePassword.Execute

				' Release the ADO command object.
				Set cmdChangePassword = nothing

				' SQL Native Client Stuff
				if Err = 3709  then
					err.number = 0
				  
					set conX = Server.CreateObject("ADODB.Connection")
					conX.ConnectionTimeout = 60
					set objSettings = server.CreateObject("COAIntServer.clsSettings")
					
					select case objSettings.GetSQLNCLIVersion
					case 9
					  sConnectString = "Provider=SQLNCLI;"
					case 10
                 sConnectString = "Provider=SQLNCLI10;"
					case 11
                 sConnectString = "Provider=SQLNCLI11;"
					end select
					set objSettings = nothing
					
					sConnectString = sConnectString & "DataTypeCompatibility=80;" & Session("SQL2005Force") & _
					  ";Old Password='" & Replace(sCurrentPassword, "'", "''") & "';Password='" & Replace(sNewPassword, "'", "''") & "'"
					
					conX.open sConnectString				

					'if err.number <> 0 then
				  '  strErrorMessage = Err.Description				
					'  err.number = 0
					'  set conX = Server.CreateObject("ADODB.Connection")
					'  conX.ConnectionTimeout = 60
					'  sConnectString = "Provider=SQLNCLI;DataTypeCompatibility=80;" & Session("SQL2005Force") & ";Old Password='" & sCurrentPassword & "';Password='" & sNewPassword & "'"
					'  conX.open sConnectString
          'end if
          
					if err.number <> 0 then
            if err.number <> 3706 then  ' 3706 = Provider not found
						  strErrorMessage = Err.Description
						end if
						Session("ErrorTitle") = "Change Password Page"
						Session("ErrorText") = strErrorMessage
						Response.Redirect("loginerror.asp")
						
					else
						conX.close				
						session("MessageTitle") = "Change Password Page"
						session("MessageText") = "Password changed successfully. You may now login."
						Response.Redirect("loginmessage.asp")
					end if
				end if

				if Err.number <> 0 then
					Session("ErrorTitle") = "Change Password Page"
					Session("ErrorText") = "You could not change your password because of the following error:<p>" & Err.Description
					Response.Redirect("loginerror.asp")
				else
					' Password changed okay. Update the appropriate record in the ASRSysPasswords table.
					Set cmdPasswordOK = Server.CreateObject("ADODB.Command")
					cmdPasswordOK.CommandText = "sp_ASRIntPasswordOK"
					cmdPasswordOK.CommandType = 4 ' Stored Procedure
					Set cmdPasswordOK.ActiveConnection = session("databaseConnection")

					err = 0
					cmdPasswordOK.Execute
					if err <> 0 then
						Session("ErrorTitle") = "Change Password Page"
						Session("ErrorText") = "You could not change your password because of the following error:<p>" & Err.Description
						Response.Redirect("loginerror.asp")
					end if
	
					' Release the ADO command object.
					Set cmdPasswordOK = nothing

					' Close and reopen the connection object.
					set conX = session("databaseConnection")
					sConnString = conX.connectionString

					iPos1 = instr(ucase(sConnString), ucase(";PWD=" & sCurrentPassword))
					if iPos1 > 0 then
						conX.close
						set conX = nothing
						session("databaseConnection") = ""
	
						sNewConnString = left(sConnString, iPos1 + 4) & sNewPassword & mid(sConnString, iPos1 + 5 + len(sCurrentPassword))
						' Open a connection to the database.
						set conX = Server.CreateObject("ADODB.Connection")
						conX.open sNewConnString
			
						if err <> 0 then
							Session("ErrorTitle") = "Change Password Page"
							Session("ErrorText") = "You could not change your password because of the following error:<p>" & Err.Description
							Response.Redirect("loginerror.asp")
						end if

						Set session("databaseConnection") = conX
					end if

					' Create the cached system tables on the server - Don;t do it in a stored procedure because the #temp will then only be visible to that stored procedure
					Set cmdCreateCache = Server.CreateObject("ADODB.Command")
					cmdCreateCache.CommandText = 		"DECLARE @iUserGroupID	integer, " & vbnewline & _
														"	@sUserGroupName		sysname, " & vbnewline & _
														"	@sActualLoginName	varchar(250) " & vbnewline & _
														"-- Get the current user's group ID. " & vbnewline & _
														"EXEC spASRIntGetActualUserDetails " & vbnewline & _
														"	@sActualLoginName OUTPUT, " & vbnewline & _
														"	@sUserGroupName OUTPUT, " & vbnewline & _
														"	@iUserGroupID OUTPUT " & vbnewline & _
														"-- Create the SysProtects cache table " & vbnewline & _
														"IF OBJECT_ID('tempdb..#SysProtects') IS NOT NULL " & vbnewline & _
														"	DROP TABLE #SysProtects " & vbnewline & _
														"CREATE TABLE #SysProtects(ID int, Action tinyint, Columns varbinary(8000), ProtectType int) " & vbnewline & _
														"	INSERT #SysProtects " & vbnewline & _
														"	SELECT ID, Action, Columns, ProtectType " & vbnewline & _
														"       FROM sysprotects " & vbnewline & _
														"       WHERE uid = @iUserGroupID"
					'cmdCreateCache.CommandType = 4 ' Stored Procedure
					Set cmdCreateCache.ActiveConnection = conX
					cmdCreateCache.execute
					Set cmdCreateCache = nothing

					Session("MessageTitle") = "Change Password Page"
					Session("MessageText") = "Password changed successfully."
					Response.Redirect("loginmessage.asp")
				end if
			else
				Session("ErrorTitle") = "Change Password Page"
				sErrorText = "You could not change your password.<p>The account is currently being used by "
				if iUserSessionCount > 2 then
					sErrorText = sErrorText & iUserSessionCount & " users"
				else
					sErrorText = sErrorText & "another user"
				end if
				sErrorText = sErrorText & " in the system."
				Session("ErrorText") = sErrorText
				
				Response.Redirect("loginerror.asp")
			end if
		else
			' Go to the main page.
			Response.Redirect("main.asp")
		end if
	end if
%>

</asp:Content>
