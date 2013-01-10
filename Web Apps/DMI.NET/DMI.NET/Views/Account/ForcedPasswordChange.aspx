<%@ Page Title="Title" Language="VB" Inherits="System.Web.Mvc.ViewPage" MasterPageFile="~/Views/Shared/Site.Master" %>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
<%
	' Only open the form if there was a referring page.
	' If it wasn't then redirect to the login page.
	Dim sReferringPage = Request.ServerVariables("HTTP_REFERER")
	if inStrRev(sReferringPage, "/") > 0 then
		sReferringPage = mid(sReferringPage, inStrRev(sReferringPage, "/") + 1)
	end if

	If Len(sReferringPage) = 0 Then
		Response.Redirect("Login")
	End If
%>

<!--Client script to handle the screen events.-->
<script type="text/javascript">
<!--
	/* Validate the password change, and change the user's password
	on the SQL database if everything is okay. */
	function SubmitPasswordDetails() {
		var sCurrentPassword;
		var sNewPassword1;
		var sNewPassword2;
		var fChangeOK;
		var sErrorMessage;
		var frmPasswordChangeForm = document.getElementById('frmPasswordChangeForm');
		var txtMinPasswordLength = document.getElementById('txtMinPasswordLength');
		
		fChangeOK = true;
		sCurrentPassword = frmPasswordChangeForm.txtCurrentPassword.value.toLowerCase();
		sNewPassword1 = frmPasswordChangeForm.txtPassword1.value.toLowerCase();
		sNewPassword2 = frmPasswordChangeForm.txtPassword2.value.toLowerCase();

		/* Check that the two new passwords are the same. */
		if (sNewPassword1 != sNewPassword2) {
			alert("The confirmation password is not correct.");
			fChangeOK = false;
			frmPasswordChangeForm.txtPassword2.value = "";
			frmPasswordChangeForm.txtPassword2.focus();
		}

		/* Check that the new password is different to the old one. */
		if (sNewPassword1 == sCurrentPassword) {
			alert("The new password cannot be the same as the old one.");
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
			alert(sErrorMessage);
			fChangeOK = false;
			frmPasswordChangeForm.txtPassword1.value = "";
			frmPasswordChangeForm.txtPassword2.value = "";
			frmPasswordChangeForm.txtPassword1.focus();
		}

		/* If everything is okay, submit the password change. */
		if (fChangeOK) {
			frmPasswordChangeForm.submit();
		}
	}
-->
</script>

<div <%=session("BodyTag")%>>

<% Html.BeginForm("ForcedPasswordChange", "Account", FormMethod.Post, New With {.id = "frmPasswordChangeForm", .name = "frmPasswordChangeForm"})%>

<BR>
<table class="outline" align=center cellPadding=5 cellSpacing=0> 
    <tr>
	    <td>
		    <table align=center class="invisible" cellPadding=0 cellSpacing=0 width=100% height=100%>
			    <TR>
					<TD>
						<table align=center class="invisible" cellPadding=0 cellSpacing=0> 
						    <tr>
						        <td colSpan=5 height=40><H3 align=center>You must change your password</H3></td>
						    </tr>
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
							    <td  colSpan=5 align=center>
									<input id="submitPasswordDetails" name="submitPasswordDetails" type="button" class="btn" value="OK" style="WIDTH: 100px" width="100" 
									    onclick="SubmitPasswordDetails()"
					                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
			                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
			                            onfocus="try{button_onFocus(this);}catch(e){}"
			                            onblur="try{button_onBlur(this);}catch(e){}" />
								</td>
							</tr>
							<tr>
							  <td colSpan=5 height=10></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<% Html.EndForm()%>

<%
	On Error Resume Next
	Response.Write("<INPUT type='hidden' id=txtMinPasswordLength name=txtMinPasswordLength value=" & Session("minPasswordLength") & ">" & vbCrLf)
%>
</div>
</asp:Content>
