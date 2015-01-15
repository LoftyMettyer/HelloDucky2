<%@ Page Title="" Language="VB" Inherits="System.Web.Mvc.ViewPage(Of DMI.NET.Models.LoginViewModel)" MasterPageFile="~/Views/Shared/Site.Master" %>

<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="DMI.NET.Code" %>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server"><%= GetPageTitle("Login") %></asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
		
	<%Html.EnableClientValidation()%>
<div class="divLogin">
	<%Html.BeginForm("Login", "Account", FormMethod.Post, New With {.id = "frmLoginForm", .defaultbutton = "submitLoginDetails"})%>
	<div class="ui-dialog-titlebar ui-widget-header loginTitleBar">
		<img alt="about OpenHR" title="About OpenHR Web" src="<%= Url.Content("~/Content/images/help32.png")%>" />
	</div>
	
	<div class="verticalpadding200"></div>
						
	<div class="ui-widget-content ui-corner-tl ui-corner-br loginframe">

		<img alt="loginimage" class="loginframeImage" src="<%= Url.Content("~/Content/images/OpenHRWeb_Splash.png")%>">

		<p class="centered">Version <%=session("Version")%></p><br />

		<p id="ancientBrowser" class="centered hidden">OpenHR Web can only be accessed using Microsoft Internet Explorer 10 or later.</p>
		<p id="systemLocked" class="centered hidden">A system administrator has locked the database.</p>


		<%If Len(Session("version")) = 0 Then%>
		<p class="centered">
			Unable to determine the OpenHR version.<br />
			Ensure that a virtual directory has been configured on your web server.
		</p>
		<%Else%>
		<div id="divLoginDetails">

			<div class="loginframeField">
				<%=Html.LabelFor(Function(loginviewmodel) loginviewmodel.UserName)%>
				<%=Html.TextBoxFor(Function(loginviewmodel) loginviewmodel.UserName, New With {.id = "txtUserName", .onkeypress = "CheckKeyPressed(event)"})%>
				<%=Html.ValidationMessageFor(Function(loginviewmodel) loginviewmodel.UserName)%>
			</div>

			<div class="loginframeField">
				<%=Html.LabelFor(Function(loginviewmodel) loginviewmodel.Password)%>
				<%=Html.PasswordFor(Function(loginviewmodel) loginviewmodel.Password, New With {.id = "txtPassword", .onkeypress = "CheckKeyPressed(event)"})%>
				<%=Html.ValidationMessageFor(Function(loginviewmodel) loginviewmodel.Password)%>
			</div>
		
			<%If Platform.IsWindowsAuthenicatedEnabled() Then%>
			<div class="loginframeFieldWA">
				<%=Html.CheckBoxFor(Function(loginviewmodel) loginviewmodel.WindowsAuthentication, New With {.id = "chkWindowsAuthentication", .onclick = "ToggleWindowsAuthentication()"})%>
				<%=Html.LabelFor(Function(loginviewmodel) loginviewmodel.WindowsAuthentication)%>
				<%=Html.ValidationMessageFor(Function(loginviewmodel) loginviewmodel.WindowsAuthentication)%>
			</div>
			<%End If%>
		
			<div id="divDetails" style="display: none;">
				<div class="loginframeField">
					<%=Html.LabelFor(Function(loginviewmodel) loginviewmodel.Database)%>
					<%=Html.TextBoxFor(Function(loginviewmodel) loginviewmodel.Database, New With {.id = "txtDatabase", .onkeypress = "CheckKeyPressed(event)"})%>
					<%=Html.ValidationMessageFor(Function(loginviewmodel) loginviewmodel.Database)%>
				</div>
				<div class="loginframeField">
					<%=Html.LabelFor(Function(loginviewmodel) loginviewmodel.Server)%>
					<%=Html.TextBoxFor(Function(loginviewmodel) loginviewmodel.Server, New With {.id = "txtServer", .onkeypress = "CheckKeyPressed(event)"})%>
					<%=Html.ValidationMessageFor(Function(loginviewmodel) loginviewmodel.Server)%>
				</div>
			</div>
	
			<div class="centered">
				<input type="button" id="submitLoginDetails" name="submitLoginDetails" onclick="SubmitLoginDetails()" value="Login" />
				<input type="button" id="btnToggleDetailsDiv" name="details" class="ui-button <%=IIf(Model.SetDetails, "", "hidden")%>" value="Details >>" />		

				<br />
				<p id="ForgotPasswordLink" style="display: none;"><%=Html.ActionLink("Forgot password", "ForgotPassword", "Account")%></p>
			</div>
				
		</div>

		<br />
		<%End If%>
	
	</div>
	
	<input type="hidden" id="txtSetDetails" name="txtSetDetails" value="<%=Session("showLoginDetails")%>">
	
	<input type="hidden" id="txtLocaleCulture" name="txtLocaleCulture" value="">

	<input type="hidden" id="txtLocaleDecimalSeparator" name="txtLocaleDecimalSeparator" value="<%=LocaleDecimalSeparator()%>">
	<input type="hidden" id="txtLocaleThousandSeparator" name="txtLocaleThousandSeparator" value="<%: Html.Raw(LocaleThousandSeparator())%>">
	<input type="hidden" id="txtSystemUser" name="txtSystemUser" value="<%=replace(Request.ServerVariables("LOGON_USER"),"/","\")%>">
	<input type="hidden" id="txtWordVer" name="txtWordVer" value="12">
	<input type="hidden" id="txtExcelVer" name="txtExcelVer" value="12">
	<input type="hidden" id="txtMSBrowser" name="txtMSBrowser" value="false" />

	<%Html.EndForm()%>
	</div>
		
	<script type="text/javascript">
		
		$(document).ready(function () {

			$('#submitLoginDetails').button();
			$('#submitLoginDetails').button('disable');

			var licence = $.connection['LicenceHub'];
			licence['client'].ActivateLogin = function () {
				$('#submitLoginDetails').button('enable');
			};

			var hubProxy = $.connection.NotificationHub;
			hubProxy.client.ToggleLoginButton = function (disabled, message) {
				$("#submitLoginDetails").button({ disabled: disabled });

				if (disabled) {
					$('#systemLocked').removeClass('hidden');
					$('#systemLocked')[0].innerHTML = message;
					$('#divLoginDetails').hide();
				}
				else {
					$('#systemLocked').addClass('hidden');
					$('#divLoginDetails').show();
				}
			};

			if (!window.isMobileBrowser) {
				if ('<%=Model.UserName%>'.length > 0) $('#txtPassword').focus();
				if ('<%=Model.UserName%>'.length == 0) $('#txtUser').focus();
			}
			if ('<%=Model.WindowsAuthentication.ToString().ToLower()%>' == 'true') {
				$('#chkWindowsAuthentication').prop('checked', true);
				ToggleWindowsAuthentication();
			}

			//Set up button click events
			$('.loginTitleBar img').click(function () {
				OpenHR.showAboutPopup();
			});

			$('#btnToggleDetailsDiv').click(function () {
				setDetailsDisplay(!($('#divDetails').is(':visible')));
				if ($('#divDetails').is(':visible')) $('#txtDatabase').focus();
			});

		});

		function CheckKeyPressed(e) {
			if (e.which === 13) SubmitLoginDetails();
		}

		function SubmitLoginDetails() {

			if ($('#submitLoginDetails').prop('disabled')) {
				return false;
			}

			/* Try to login to the OpenHR database. */
			var frmLoginForm = document.getElementById('frmLoginForm');

			frmLoginForm.txtLocaleCulture.value = window.UserLocale;
	
			frmLoginForm.txtLocaleDecimalSeparator.value = OpenHR.LocaleDecimalSeparator();
			frmLoginForm.txtLocaleThousandSeparator.value = OpenHR.LocaleThousandSeparator();

			frmLoginForm.submit();			
		}

		function DisableUsernamePassword(pfDisable) {
			$('#txtUserName').css('color', pfDisable ? 'lightgray' : '').prop('readonly', pfDisable);
			$('#txtPassword').css('color', pfDisable ? 'lightgray' : '').prop('readonly', pfDisable);
		}


		function ToggleWindowsAuthentication() {
			if ($('#chkWindowsAuthentication').prop('checked') == true) {
				DisableUsernamePassword(true);
				$('#txtUserName').val($('#txtSystemUser').val());
				$('#txtPassword').val('*****');
				$("#ForgotPasswordLink").hide();
			}
			else {
				DisableUsernamePassword(false);
				$('#txtPassword').val('');
				$("#ForgotPasswordLink").show();
			}
		}

		function setDetailsDisplay(pfShow) {
			if (pfShow == true) {
				$('#btnToggleDetailsDiv').prop("value", "Details <<");
				$('#divDetails').show();
				$('#btnToggleDetailsDiv').removeClass('hidden');
			}
			else {
				$('#btnToggleDetailsDiv').prop("value", "Details >>");
				$('#divDetails').hide();
			}
		}


		//Set MS browser flag
		if ("ActiveXObject" in window) document.getElementById("txtMSBrowser").value = 'true';

		//Is this a browser that supports file API; which is OK for all modern browsers (IE10+ etc)
		if (!(window.File && window.FileReader && window.FileList && window.Blob)) {
			//Show 'browser not supported' message...
			$('#ancientBrowser').removeClass('hidden');
			$('#loginFrame').addClass('hidden');
		}
		else {
			//This browser meets requirements.
			$('#ForgotPasswordLink').css('display', 'block');
		}

</script>

</asp:Content>

