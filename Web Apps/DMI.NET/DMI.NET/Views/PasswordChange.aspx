<%@ Page Title="" Language="VB" MasterPageFile="~/Views/Shared/Site.Master" Inherits="System.Web.Mvc.ViewPage" %>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
PasswordChange
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

<%@Import namespace="DMI.NET" %>
<%
	Dim sReferringPage

	' Only open the form if there was a referring page.
	' If it wasn't then redirect to the login page.
	sReferringPage = Request.ServerVariables("HTTP_REFERER") 
	if inStrRev(sReferringPage, "/") > 0 then
		sReferringPage = mid(sReferringPage, inStrRev(sReferringPage, "/") + 1)
	end if

	if len(sReferringPage) = 0 then
		Response.Redirect("login.asp")
	end if
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="OpenHR.css">
<TITLE>OpenHR Intranet</TITLE>

<SCRIPT FOR=window EVENT=onload LANGUAGE=JavaScript>
    window.parent.document.all.item("workframeset").cols = "*, 0";	

    fOK = true;	

    var sErrMsg = txtErrorDescription.value;

    if (sErrMsg.length > 0) {
        fOK = false;
        window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sErrMsg);
        window.parent.location.replace("login.asp");
    }
	
    if (fOK == true) {
        // Get menu.asp to refresh the menu.
        window.parent.frames("menuframe").refreshMenu();
	
        if (frmPasswordChangeForm.txtUserSessionCount.value < 2) {
            frmPasswordChangeForm.txtCurrentPassword.focus();
        }
    }
</SCRIPT>

<!--Client script to handle the screen events.-->
<script LANGUAGE="JavaScript">
<!--
    /* Validate the password change, and change the user's password
    on the SQL database if everything is okay. */
    function SubmitPasswordDetails()
    {
        var sCurrentPassword;
        var sNewPassword1;
        var sNewPassword2;
        var fChangeOK;
        var sErrorMessage;

        fChangeOK = true;
        sCurrentPassword = frmPasswordChangeForm.txtCurrentPassword.value.toLowerCase();
        sNewPassword1 = frmPasswordChangeForm.txtPassword1.value.toLowerCase();
        sNewPassword2 = frmPasswordChangeForm.txtPassword2.value.toLowerCase();
  
        /* Check that the two new passwords are the same. */
        if (sNewPassword1 != sNewPassword2)
        {
            window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("The confirmation password is not correct.");
            fChangeOK = false;
            frmPasswordChangeForm.txtPassword2.value = "";
            frmPasswordChangeForm.txtPassword2.focus();
        }
  
        /* Check that the new password is different to the old one. */
        if (sNewPassword1 == sCurrentPassword)
        {
            window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("The new password cannot be the same as the old one.");
            fChangeOK = false;
            frmPasswordChangeForm.txtPassword1.value = "";
            frmPasswordChangeForm.txtPassword2.value = "";
            frmPasswordChangeForm.txtPassword1.focus();
        }

        /* Check if the new password is long enough. */
        if ((fChangeOK) && (txtMinPasswordLength.value > 0) && (txtMinPasswordLength.value > sNewPassword1.length)) {
            sErrorMessage = "The password must be at least ";
            sErrorMessage = sErrorMessage.concat(txtMinPasswordLength.value);
            sErrorMessage = sErrorMessage.concat(" characters long.");
            window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sErrorMessage);
            fChangeOK = false;
            frmPasswordChangeForm.txtPassword1.value = "";
            frmPasswordChangeForm.txtPassword2.value = "";
            frmPasswordChangeForm.txtPassword1.focus();
        }

        /* JDM - 30/06/2004 - Fault 6709 - Don't disallow blank passwords */
        /* Check if the new password is long enough. */
        /*
          if ((fChangeOK) && (sNewPassword1.length == 0)) {
                sErrorMessage = "The password cannot be blank.";
                window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sErrorMessage);
                fChangeOK = false;
                frmPasswordChangeForm.txtPassword1.value = "";
                frmPasswordChangeForm.txtPassword2.value = "";
              frmPasswordChangeForm.txtPassword1.focus();
          }
        */

        /* If everything is okay, submit the password change. */
        if (fChangeOK)
        {
            frmPasswordChangeForm.submit();
        }
    }

    /* Return to the default page. */
    function cancelClick()
    {  
        window.location.href="default.asp";
    }
    -->
</script>

<!--#INCLUDE FILE="include/ctl_SetStyles.txt" -->
</HEAD>
<BODY <%=session("BodyTag")%>>
<form action="passwordChange_Submit.asp" method="post" id="frmPasswordChangeForm" name="frmPasswordChangeForm">
<BR>
<table align=center class="outline" cellPadding=5 cellSpacing=0> 
    <tr>
        <td>
            <table align=center class="invisible" cellPadding=0 cellSpacing=0> 
			    <tr>
			        <td colSpan=5 height=10></td>
			    </tr>
			    <tr>
			        <td colSpan=5><H3 align=center>Change Password</H3></td>
			    </tr>
<%
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

	Response.Write "<INPUT type='hidden' id=txtUserSessionCount name=txtUserSessionCount value=" & iUserSessionCount & ">"
			
	if iUserSessionCount < 2 then
%>
			    <tr>
			        <td width=20></td>
			        <td align=left nowrap>Current Password :</td>
			        <td width=20></td>
			        <td align=left>
			            <input id="txtCurrentPassword" name="txtCurrentPassword" type="password" class="text" style="WIDTH: 200px">
			        </td>
			        <td width=20></td>
			    </tr>
			    <tr>
			        <td width=20></td>
			        <td align=left nowrap>New Password :</td>
			        <td width=20></td>
			        <td align=left>
			            <input id="txtPassword1" name="txtPassword1" type="password" class="text" style="WIDTH: 200px">
			        </td>
			        <td width=20></td>
			    </tr>
			    <tr>
			        <td width=20></td>
			        <td align=left nowrap>Confirm New Password :</td>
			        <td width=20></td>
			        <td align=left>
			            <input id="txtPassword2" name="txtPassword2" type="password" class="text" style="WIDTH: 200px">
			        </td>
			        <td width=20></td>
			    </tr>
			    <tr>
			        <td colSpan=5 height=20></td>
			    </tr>
			    <tr>
			        <td  colSpan=5>
						<table class="invisible" CELLSPACING="0" CELLPADDING="0" align="center">
							<td align=center>
							    <input id="submitPasswordDetails" name="submitPasswordDetails" type="button" class="btn" value="OK" style="WIDTH: 75px" width="75" 
							        onclick="SubmitPasswordDetails()"
			                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
			                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
			                        onfocus="try{button_onFocus(this);}catch(e){}"
			                        onblur="try{button_onBlur(this);}catch(e){}" />
							</td>
							<td width=20></td>
							<td align=center>
							    <input id="btnCancel" name="btnCancel" type="button" class="btn" value="Cancel" style="WIDTH: 75px" width="75" 
							        onclick="cancelClick()"
			                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
			                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
			                        onfocus="try{button_onFocus(this);}catch(e){}"
			                        onblur="try{button_onBlur(this);}catch(e){}" />
							</td>
						</table>
					</td>
			    </tr>
			    <tr>
			        <td colSpan=5 height=10></td>
                </tr>
            </table>
        </td>
    </tr>
</table>
<P></P>
<%
	else
%>
			    <tr>
			        <td width=20></td>
			        <td colspan=3>
<%
	sErrorText = "You cannot change your password.<p>The account is currently being used by "
	if iUserSessionCount > 2 then
		sErrorText = sErrorText & iUserSessionCount & " users"
	else
		sErrorText = sErrorText & "another user"
	end if
	sErrorText = sErrorText & " in the system."
	Response.write sErrorText
%>
			        </td>
			        <td width=20></td>
			    </tr>
			  
			    <tr> 
			        <td colspan=5 height=20></td>
			    </tr>

			    <tr> 
			        <td colspan=5 height=10 align=center> 
						<INPUT TYPE=button VALUE="Cancel" NAME="btnCancel" class="btn" style="WIDTH: 80px" width=80 id=Button1
						    OnClick="cancelClick()" 
	                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
	                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
	                        onfocus="try{button_onFocus(this);}catch(e){}"
	                        onblur="try{button_onBlur(this);}catch(e){}" />
                    </td>
			    </tr>

			    <tr> 
			        <td colspan=5 height=10></td>
			    </tr>
            </table>
        </td>
    </tr>
</table>
<%
	end if
%>
</form>

<FORM action="passwordChange_Submit.asp" method=post id=frmGoto name=frmGoto><INPUT type="hidden" id=txtAction name=txtAction>
<!--#include file="include\gotoWork.txt"-->
</FORM>

<%
	on error resume next
	
	Dim sErrorDescription
	sErrorDescription = ""

	' Get the minimum password length.
	Set cmdPwdLength = Server.CreateObject("ADODB.Command")
	cmdPwdLength.CommandText = "sp_ASRIntGetMinimumPasswordLength"
	cmdPwdLength.CommandType = 4 ' Stored Procedure
	Set cmdPwdLength.ActiveConnection = session("databaseConnection")
		
	Set prmMinPasswordLength = cmdPwdLength.CreateParameter("MinPasswordLength",3,2) ' 3 = integer, 2 = output
	cmdPwdLength.Parameters.Append prmMinPasswordLength

	err = 0
	cmdPwdLength.Execute
	if (err <> 0) then
		sErrorDescription = "The minimum password length could not be determined." & vbcrlf & formatError(Err.Description)
	end if

	Response.Write "<INPUT type='hidden' id=txtMinPasswordLength name=txtMinPasswordLength value=" & cmdPwdLength.Parameters("MinPasswordLength").Value & ">" & vbcrlf

	' Release the ADO command object.
	Set cmdPwdLength = nothing

	Response.Write "<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>"
%>

</BODY>
</HTML>

<% 
function formatError(psErrMsg)
  Dim iStart 
  dim iFound 
  
  iFound = 0
  Do
    iStart = iFound
    iFound = InStr(iStart + 1, psErrMsg, "]")
  Loop While iFound > 0
  
  If (iStart > 0) And (iStart < Len(Trim(psErrMsg))) Then
    formatError = Trim(Mid(psErrMsg, iStart + 1))
  Else
    formatError = psErrMsg
  End If
end function
%>


</asp:Content>
