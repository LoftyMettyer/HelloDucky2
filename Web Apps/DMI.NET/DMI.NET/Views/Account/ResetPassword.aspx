<%@ Page Title="" Language="VB" MasterPageFile="~/Views/Shared/Site.Master" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET.Code" %>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
	
	<%
		Dim sUserName As String = ""
		Dim sQueryString As String = ""
			
		sQueryString = Request.ServerVariables("QUERY_STRING")
	
		If Len(sQueryString) = 0 Then
			Response.Redirect(Url.Action("Login", "Account"))
		End If
 
		Dim objResetPwd As New ResetPassword
		
		' Retrieve the username from the encrypted string 
		' NB this returns null if validation fails, i.e. expired link.
		sUserName = objResetPwd.GetUsernameFromQueryString(sQueryString)
	
		'We are using SQL Server's response to find out the minimum apassword length
		'Session("minPasswordLength") = objResetPwd.minPasswordLength 
		
		objResetPwd = Nothing
 %>
	
	<script type="text/javascript">
		$(document).ready(function ()
		{
			$("#MessageBox").dialog({
				autoOpen: false,
				modal: true,
				resizable: false,
				height: 'auto',
				width: 'auto'
			});
		});
</script>

<div id="MessageBox" style="display: none;">
	<p id="Message"></p>
	<input type="button" class="btn" value="OK" style="float: right; margin-left: auto; margin-right: auto;" onclick="$('#MessageBox').dialog('close'); return false;" />
</div>

<script type="text/javascript">

	function window_Onload()
	{
		frmResetPasswordForm.txtPassword1.focus();
	}

	function HelpAbout() {
			$("#About").dialog( "open" );
	}
	
	function CheckKeyPressed(e) {
		var keynum;

		if (window.event) // IE8 and earlier
		{
			keynum = e.keyCode;
		}
		else if (e.which) // IE9/Firefox/Chrome/Opera/Safari
		{
			keynum = e.which;
		}

		if (keynum == 13) { // 13 = enter key
			SubmitPasswordDetails();			
		}
	}
	
	/* Validate the password change, and change the user's password
	on the SQL database if everything is okay. */
	function SubmitPasswordDetails()
	{
		var sNewPassword1;
		var sNewPassword2;
		var fChangeOK = true;
		
		sNewPassword1 = frmResetPasswordForm.txtPassword1.value.toLowerCase();
		sNewPassword2 = frmResetPasswordForm.txtPassword2.value.toLowerCase();
		
		// Check that the user has entered a new password
		if (sNewPassword1 == "") {
			$("#Message").html("Please enter a new password.");
			$("#MessageBox").dialog("open");
			fChangeOK = false;
		}

		/* Check that the two new passwords are the same. */
		if (sNewPassword1 != sNewPassword2) {
			$("#Message").html("The confirmation password is not correct.");
			$("#MessageBox").dialog("open");
			fChangeOK = false;
			frmResetPasswordForm.txtPassword2.value = "";
		} 
		
		if (fChangeOK == true) { /* Everything is okay, submit the password change. */
			frmResetPasswordForm.submit();
		}
	}
</script>

	<div class="divLogin">
		<%Html.BeginForm("ResetPassword_Submit", "Account", FormMethod.Post, New With {.id = "frmResetPasswordForm"})%>
		<div class="ui-dialog-titlebar ui-widget-header loginTitleBar">
		</div>
		<div class="verticalpadding200"></div>
						
	<div class="ui-widget-content ui-corner-tl ui-corner-br loginframe">
		<table style="margin: 0 auto; width: 1px;">
			<tr> 
					<td> 
							<img height="188" src="<%=Url.Content("~/Content/images/OpenHRWeb_Splash.png")%>" style="width: 410px;" alt="">
					</td>
			</tr>
			<tr>
				<td style="text-align: center" > 
						<h3 style="text-align: center;">Reset your password</h3>
				</td>
			</tr>
				
			<tr> 
				<td style="text-align:center">
					<label for="txtCurrentPassword" style="float: left">Current Password : </label>
					<input type="password" name="txtCurrentPassword" disabled="disabled" id="txtCurrentPassword" style="width: 180px; float: right;" value="*****" class="text" />
				</td>
			</tr>

			<tr> 
				<td style="text-align:center">
					<label for="txtPassword1" style="float: left">New Password : </label>
					<input autocomplete="off" type="password" name="txtPassword1" id="txtPassword1" style="width: 180px; float: right;" value="" class="text" onkeypress="CheckKeyPressed(event);" />
				</td>
			</tr>
			
			<tr> 
				<td style="text-align:center">
					<label for="txtPassword2" style="float: left">Confirm New Password : </label>
					<input autocomplete="off" type="password" name="txtPassword2" id="txtPassword2" style="width: 180px; float: right;" value="" class="text" onkeypress="CheckKeyPressed(event);" />
				</td>
			</tr>
						
			<tr>
				<td><br/></td>
			</tr>

			<tr>
				<td style="text-align:center">
				  <img src="<%=Url.Action("GetCaptcha","Generic")%>" />
				</td>
			</tr>

			<tr style="text-align:center">
				<td>Please type the above characters to ensure that you a person, not an automated program submitting this request.</td>
			</tr>
 			<tr style="text-align:center"> 
				<td style="text-align:center">
					Verify : &nbsp;&nbsp;&nbsp;
					<input type="text" name="txtVerify" id="txtVerify" value="" class="text" />
				</td>
			</tr>

			<tr> 
				<td style="text-align: center;">
					<input type="button" value="OK" onclick="SubmitPasswordDetails();" style="width: 100px;" />
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<input type="button" value="Cancel" onclick="window.location='<%=Url.Action("Login", "Account")%>';" style="width: 100px;" />
				</td>
			</tr>
		</table>
		<br/>
	</div>
		<input type="hidden" id="txtQueryString" name="txtQueryString" value="<%=sQueryString%>"/>
		<input type="hidden" id="txtUser" name="txtUser" value="<%=sUserName%>"/>      

		<%=Html.AntiForgeryToken()%>

	<%Html.EndForm()%>
</div>

<style>
	header { height: 48px; width: 99.9%; z-index: -1; }
</style>

<%If String.IsNullOrEmpty(sUserName) Then%>
		<script type="text/javascript">
			alert("Unfortunately the link you've clicked is no longer valid. Please click OK to return to the main page and start again.");
			window.location = '<%=Url.Action("Login", "Account")%>';
		</script>
<%End If%>
		
<script type="text/javascript">
	window_Onload();
	
	//Prevent the form from being submitted (without being checked) when the user presses Enter
	$('#frmforgotPasswordForm').bind("keypress", function(e) {
		var code = e.keyCode || e.which; 
		if (code == 13) {               
			e.preventDefault();
			return false;
		}
	});
</script>
</asp:Content>
