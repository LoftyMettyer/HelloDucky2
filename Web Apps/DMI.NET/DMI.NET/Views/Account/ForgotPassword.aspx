<%@ Page Title="" Language="VB" MasterPageFile="~/Views/Shared/Site.Master" Inherits="System.Web.Mvc.ViewPage" %>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
<script type="text/javascript">

	function window_Onload() {

		var usernamevalue = getQuerystring('username');

		if (usernamevalue.length > 0) {
			frmforgotPasswordForm.txtUserName.value = usernamevalue;      
		}
		frmforgotPasswordForm.txtUserName.focus();
	}
		
</script>

<script type="text/javascript">
	/* Validate the form, and continue if everything is okay. */
	function validateForm() {
		var sUserName = frmforgotPasswordForm.txtUserName.value.toLowerCase();
		var fChangeOK = true;
		var sErrorMessage;

		if (sUserName.length == 0) {
			$("#MessageBox").dialog("open");
			fChangeOK = false;
			return false;
		}
		
		/* If everything is okay, submit the password change. */
		frmforgotPasswordForm.submit();
	}


	function getQuerystring(key, default_) {
		if (default_ == null) default_ = "";
		key = key.replace(/[\[]/, "\\\[").replace(/[\]]/, "\\\]");
		var regex = new RegExp("[\\?&]" + key + "=([^&#]*)");
		var qs = regex.exec(window.location.href);
		if (qs == null)
			return default_;
		else
			return qs[1];
	}  
	
</script>

<div <%=Session("BodyTag")%> style="width: 98%; position: absolute; top: 170px;">
	<%Html.BeginForm("ForgotPassword_Submit", "Account", FormMethod.Post, New With {.id = "frmforgotPasswordForm"})%>
		<table style="margin: 0 auto; width: 1px;">
			<tr> 
					<td> 
							<img height="188" src="<%=Url.Content("~/Content/images/COAInt_Splash.png")%>" style="width: 410px;" alt="">
					</td>
			</tr>
			<tr>
				<td style="text-align: center" > 
						<h2 style="text-align: center;">Forgot password</h2>
						<p>You can change or reset the password for your account<br/>by providing some information.</p>
				</td>
			</tr>

			<tr> 
				<td style="text-align:center">
					User name : &nbsp;&nbsp;&nbsp;
					<input type="text" name="txtUserName" id="txtUserName" value="" class="text" />
				</td>
			</tr>
			
			<tr>
				<td></td>
			</tr>

			<tr>
				<td></td>
			</tr>

			<tr> 
				<td style="text-align: center;">
					<input type="button" value="OK" onclick="validateForm();" style="width: 100px;" />
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					<input type="button" value="Cancel" onclick="window.location='<%=Url.Action("Login", "Account")%>'" style="width: 100px;" />
				</td>
			</tr>
		</table>
	<%Html.EndForm()%>
</div>
<script type="text/javascript">
	$(document).ready(function() {
		$("#MessageBox").dialog({
			autoOpen: false,
			modal: true,
			resizable: false,
			height: 'auto',
			width: 'auto'
		});
	});
</script>

<div id="MessageBox" title="OpenHR" style="display: none;">
	<p>Please enter your user name.</p>
	<input type="button" class="btn" value="OK" style="float: right; margin-left: auto; margin-right: auto;" onclick="$('#MessageBox').dialog('close'); return false;" />
</div>
<script type="text/javascript">
	window_Onload();
</script>
</asp:Content>
