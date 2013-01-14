<%@ Page Title="" Language="VB" MasterPageFile="~/Views/Shared/Site.Master" Inherits="System.Web.Mvc.ViewPage" %>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
passwordChange_Submit
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

<%@ Language=VBScript %>
<%@Import namespace="DMI.NET" %>

<% 
	on error resume next
    
    Dim sReferringPage = ""
    Dim fSubmitPasswordChange = ""
    Dim sErrorText = ""
    
	' Only process the form submission if the referring page was the newUser page.
	' If it wasn't then redirect to the login page.
	sReferringPage = Request.ServerVariables("HTTP_REFERER") 
	if inStrRev(sReferringPage, "/") > 0 then
		sReferringPage = mid(sReferringPage, inStrRev(sReferringPage, "/") + 1)
	end if

	if ucase(sReferringPage) <> ucase("passwordChange.asp") then
		Response.Redirect("login.asp")
	else
		fSubmitPasswordChange = (len(Request.Form("txtGotoPage")) = 0) 

		if fSubmitPasswordChange then
			' Force password change only if there are no other users logged in with the same name.
            Dim cmdCheckUserSessions = Server.CreateObject("ADODB.Command")
			cmdCheckUserSessions.CommandText = "spASRGetCurrentUsersCountOnServer"
			cmdCheckUserSessions.CommandType = 4 ' Stored procedure.
            cmdCheckUserSessions.ActiveConnection = Session("databaseConnection")

            Dim prmCount = cmdCheckUserSessions.CreateParameter("count", 3, 2) ' 3=integer, 2=output
            cmdCheckUserSessions.Parameters.Append(prmCount)

            Dim prmUserName = cmdCheckUserSessions.CreateParameter("userName", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
            cmdCheckUserSessions.Parameters.Append(prmUserName)
            prmUserName.value = Session("Username")

            Err.Number = 0
            cmdCheckUserSessions.Execute()
			
            Dim iUserSessionCount = CLng(cmdCheckUserSessions.Parameters("count").Value)
            cmdCheckUserSessions = Nothing
			
            If iUserSessionCount < 2 Then
                ' Read the Password details from the Password form.
                Dim sCurrentPassword = Request.Form("txtCurrentPassword")
                Dim sNewPassword = Request.Form("txtPassword1")

                ' Attempt to change the password on the SQL Server.
                Dim cmdChangePassword = Server.CreateObject("ADODB.Command")
                cmdChangePassword.CommandText = "sp_password"
                cmdChangePassword.CommandType = 4 ' Stored Procedure
                cmdChangePassword.ActiveConnection = Session("databaseConnection")

                Dim prmCurrentPassword = cmdChangePassword.CreateParameter("currentPassword", 200, 1, 255)
                cmdChangePassword.Parameters.Append(prmCurrentPassword)
                If Len(sCurrentPassword) > 0 Then
                    prmCurrentPassword.value = sCurrentPassword
                Else
                    prmCurrentPassword.value = DBNull.Value
                End If

                Dim prmNewPassword = cmdChangePassword.CreateParameter("newPassword", 200, 1, 255)
                cmdChangePassword.Parameters.Append(prmNewPassword)
                If Len(sNewPassword) > 0 Then
                    prmNewPassword.value = sNewPassword
                Else
                    prmNewPassword.value = DBNull.Value
                End If

                Err.Number = 0
                cmdChangePassword.Execute()

                ' Release the ADO command object.
                cmdChangePassword = Nothing

                If Err.Number <> 0 Then
                    Session("ErrorTitle") = "Change Password Page"
                    Session("ErrorText") = "You could not change your password because of the following error:<p>" & formatError(Err.Description)
                    Response.Redirect("error.asp")
                Else
                    ' Password changed okay. Update the appropriate record in the ASRSysPasswords table.
                    Dim cmdPasswordOK = Server.CreateObject("ADODB.Command")
                    cmdPasswordOK.CommandText = "sp_ASRIntPasswordOK"
                    cmdPasswordOK.CommandType = 4 ' Stored Procedure
                    cmdPasswordOK.ActiveConnection = Session("databaseConnection")

                    Err.Number = 0
                    cmdPasswordOK.Execute()
                    If Err.Number <> 0 Then
                        Session("ErrorTitle") = "Change Password Page"
                        Session("ErrorText") = "You could not change your password because of the following error:<p>" & formatError(Err.Description)
                        Response.Redirect("error.asp")
                    End If
	
                    ' Release the ADO command object.
                    cmdPasswordOK = Nothing

                    ' Close and reopen the connection object.
                    Dim conX = Session("databaseConnection")
                    Dim sConnString = conX.connectionString

                    Dim iPos1 = InStr(UCase(sConnString), UCase(";PWD=" & sCurrentPassword))
                    If iPos1 > 0 Then
                        conX.close()
                        conX = Nothing
                        Session("databaseConnection") = ""
	
                        
                        Dim sNewConnString = Left(sConnString, iPos1 + 4) & sNewPassword & Mid(sConnString, iPos1 + 5 + Len(sCurrentPassword))
                        ' Open a connection to the database.
                        conX = Server.CreateObject("ADODB.Connection")
                        conX.open(sNewConnString)
			
                        If Err.Number <> 0 Then
                            Session("ErrorTitle") = "Change Password Page"
                            Session("ErrorText") = "You could not change your password because of the following error:<p>" & formatError(Err.Description)
                            Response.Redirect("error.asp")
                        End If

                        Session("databaseConnection") = conX
                        
                    End If

                    ' Create the cached system tables on the server - Don;t do it in a stored procedure because the #temp will then only be visible to that stored procedure
                    Dim cmdCreateCache = Server.CreateObject("ADODB.Command")
                    cmdCreateCache.CommandText = "DECLARE @iUserGroupID	integer, " & vbNewLine & _
                                                        "	@sUserGroupName		sysname, " & vbNewLine & _
                                                        "	@sActualLoginName	varchar(250) " & vbNewLine & _
                                                        "-- Get the current user's group ID. " & vbNewLine & _
                                                        "EXEC spASRIntGetActualUserDetails " & vbNewLine & _
                                                        "	@sActualLoginName OUTPUT, " & vbNewLine & _
                                                        "	@sUserGroupName OUTPUT, " & vbNewLine & _
                                                        "	@iUserGroupID OUTPUT " & vbNewLine & _
                                                        "-- Create the SysProtects cache table " & vbNewLine & _
                                                        "IF OBJECT_ID('tempdb..#SysProtects') IS NOT NULL " & vbNewLine & _
                                                        "	DROP TABLE #SysProtects " & vbNewLine & _
                                                        "CREATE TABLE #SysProtects(ID int, Action tinyint, Columns varbinary(8000), ProtectType int) " & vbNewLine & _
                                                        "	INSERT #SysProtects " & vbNewLine & _
                                                        "	SELECT ID, Action, Columns, ProtectType " & vbNewLine & _
                                                        "       FROM sysprotects " & vbNewLine & _
                                                        "       WHERE uid = @iUserGroupID"
                    'cmdCreateCache.CommandType = 4 ' Stored Procedure
                    cmdCreateCache.ActiveConnection = conX
                    cmdCreateCache.execute()
                    cmdCreateCache = Nothing

                    ' Tell the user that the password was changed okay.
                    Session("MessageTitle") = "Change Password Page"
                    Session("MessageText") = "Password changed successfully."
                    Response.Redirect("message.asp")
                End If
            Else
                Session("ErrorTitle") = "Change Password Page"
                sErrorText = "You could not change your password.<p>The account is currently being used by "
                If iUserSessionCount > 2 Then
                    sErrorText = sErrorText & iUserSessionCount & " users"
                Else
                    sErrorText = sErrorText & "another user"
                End If
                sErrorText = sErrorText & " in the system."
                Session("ErrorText") = sErrorText
				
                Response.Redirect("error.asp")
            End If
		else
			' Save the required table/view and screen IDs in session variables.
			Session("action") = Request.Form("txtAction")
			Session("tableID") = Request.Form("txtGotoTableID")
			Session("viewID") = Request.Form("txtGotoViewID")
			Session("screenID") = Request.Form("txtGotoScreenID")
			Session("orderID") = Request.Form("txtGotoOrderID")
			Session("recordID") = Request.Form("txtGotoRecordID")
			Session("parentTableID") = Request.Form("txtGotoParentTableID")
			Session("parentRecordID") = Request.Form("txtGotoParentRecordID")
			Session("realSource") = Request.Form("txtGotoRealSource")
			Session("filterDef") = Request.Form("txtGotoFilterDef")
			Session("filterSQL") = Request.Form("txtGotoFilterSQL")
			Session("lineage") = Request.Form("txtGotoLineage")
			Session("defseltype") = Request.Form("txtGotoDefSelType")
			Session("utilID") = Request.Form("txtGotoUtilID")
			Session("locateValue") = Request.Form("txtGotoLocateValue")
			Session("firstRecPos") = Request.Form("txtGotoFirstRecPos")
			Session("currentRecCount") = Request.Form("txtGotoCurrentRecCount")
			Session("fromMenu") = Request.Form("txtGotoFromMenu")

			' Go to the requested page.
			Response.Redirect(Request.Form("txtGotoPage"))
		end if
	end if
%>
</asp:Content>
