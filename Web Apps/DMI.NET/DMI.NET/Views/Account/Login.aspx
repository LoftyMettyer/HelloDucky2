<%@ Page Title="" Language="VB" Inherits="System.Web.Mvc.ViewPage" MasterPageFile="~/Views/Shared/Site.Master" %>
<%@ Import Namespace="DMI.NET" %>
<%@ import Namespace="System.Web.Configuration" %>
<%@ Import Namespace="DMI.NET.Code" %>

<script runat="server">
		Private _txtDatabaseValue As String = "" 'To set the value of the txtDatabase input tag
		Private _txtServerValue As String = "" 'To set the value of the txtServer input tag
		
		Private Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
				'If no query string is present, hide the "Details" button and the Database and Server labels and input box controls
				If Request.QueryString.Count = 0 Then
						btnToggleDetailsDiv.Attributes.Add("style", "display: none")
						DatabaseTextLabelDiv.Attributes.Add("style", "display: none")
						DatabaseTextValueDiv.Attributes.Add("style", "display: none")
						ServerTextLabelDiv.Attributes.Add("style", "display: none")
						ServerTextValueDiv.Attributes.Add("style", "display: none")
				Else 'Override database or server if a value is provided in the querystring
						If Not String.IsNullOrEmpty(Request.QueryString("database")) Then
								_txtDatabaseValue = Server.HtmlDecode(Request.QueryString("database"))
						End If
						If Not String.IsNullOrEmpty(Request.QueryString("server")) Then
								_txtServerValue = Server.HtmlDecode(Request.QueryString("server"))
						End If
				End If

				'If no override values were provided in the querystring, use default values from web.config
				If String.IsNullOrEmpty(_txtDatabaseValue) Then
						_txtDatabaseValue = ApplicationSettings.LoginPage_Database
				End If
				If String.IsNullOrEmpty(_txtServerValue) Then
						_txtServerValue = ApplicationSettings.LoginPage_Server
				End If
		End Sub
</script>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
		<%= GetPageTitle("Login") %>    
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

<%	
	On Error Resume Next            
		
	Dim sBrowserInfo As String
		
		' Ensure the database connection object is closed.
		Dim conX = Session("databaseConnection")
		If Not conX Is Nothing Then
				conX.Close()
		End If
		conX = Nothing
	
	Session("databaseConnection") = Nothing
	Session("action") = ""
	Session("selectSQL") = ""
	Session("filterSQL") = ""
	Session("filterDef") = ""
	Session("optionAction") = ""
	Session("server") = ""
	
	Session("showLoginDetails") = Request.QueryString("Details")

	'TODO
	' Clear out any session objects.
	'For Each sessitem in Session.Contents
	'	If TypeOf Session.Contents(sessitem) Is Object Then
	'		Session.Contents(sessitem) = Nothing
	'		Session.Contents(sessitem) = ""
	'		Session.Contents.Remove(sessitem)
	'	End If
	'Next 
	
	Session("dfltTempMenuFilePath") = "<NONE>"

	If (Len(Session("Version")) > 0) Then

		Response.Write("<script type='text/javascript'>" & vbCrLf)
		Response.Write("	function window_onload() {" & vbCrLf)
		Response.Write("    var sUserName;" & vbCrLf)
		Response.Write("    var sDatabase;" & vbCrLf)
		Response.Write("	  var sServer;" & vbCrLf)
		Response.Write("    var sWindowsAuthentication;" & vbCrLf)
		Response.Write("		var frmLoginForm = document.getElementById('frmLoginForm');" & vbCrLf)
		Response.Write("if(!(frmLoginForm.txtUserName)) {" & vbCrLf)
		Response.Write("	//An error message has been displayed and no tags configured, so quit." & vbCrLf)
		Response.Write("	return false;" & vbCrLf)
		Response.Write("	}" & vbCrLf)
		
		
		If Request.QueryString("user") <> "" Then
			Response.Write("    frmLoginForm.txtUserName.value = """ & CleanStringForJavaScript(Request.QueryString("user")) & """;" & vbCrLf)
			Response.Write("    if(!window.isMobileBrowser) frmLoginForm.txtPassword.focus();" & vbCrLf)
		ElseIf Request.QueryString("username") <> "" Then
			Response.Write("    frmLoginForm.txtUserName.value = """ & CleanStringForJavaScript(Request.QueryString("username")) & """;" & vbCrLf)
			Response.Write("    if(!window.isMobileBrowser) frmLoginForm.txtPassword.focus();" & vbCrLf)
		ElseIf Session("username") <> "" Then
			Response.Write("    frmLoginForm.txtUserName.value = """ & CleanStringForJavaScript(Session("username")) & """;" & vbCrLf)
			Response.Write("    if(!window.isMobileBrowser) frmLoginForm.txtPassword.focus();" & vbCrLf)
		Else
			If Not Request.Cookies("Login") Is Nothing Then
				Response.Write("    sUserName = '" & Server.HtmlEncode(Request.Cookies("Login")("User")) & "' ;" & vbCrLf)
				'Response.Write("    sDatabase = '" & Server.HtmlEncode(Request.Cookies("Login")("Database")) & "' ;" & vbCrLf)
				'Response.Write("    sServer = '" & Server.HtmlEncode(Request.Cookies("Login")("Server")) & "' ;" & vbCrLf)
				Response.Write("    sWindowsAuthentication = '" & Server.HtmlEncode(Request.Cookies("Login")("WindowsAuthentication")) & "' ;" & vbCrLf)
			End If
			
			Response.Write("    if (sUserName != """" && sUserName != null && sUserName != ""undefined"") {" & vbCrLf)
			Response.Write("      frmLoginForm.txtUserName.value = sUserName;" & vbCrLf)
			Response.Write("      if(!window.isMobileBrowser) frmLoginForm.txtPassword.focus();" & vbCrLf)
			Response.Write("    }" & vbCrLf)
			Response.Write("    else {" & vbCrLf)
			Response.Write("      if(!window.isMobileBrowser) frmLoginForm.txtUserName.focus();" & vbCrLf)
			Response.Write("    }" & vbCrLf)
		End If

		If Request.QueryString("database") <> "" Then
			Response.Write("    frmLoginForm.txtDatabase.value = """ & CleanStringForJavaScript(Request.QueryString("database")) & """;" & vbCrLf)
		ElseIf Session("database") <> "" Then
			Response.Write("    frmLoginForm.txtDatabase.value = """ & CleanStringForJavaScript(Session("database")) & """;" & vbCrLf)
			'Else
			'Response.Write("    sDatabase = getCookie('Intranet_Database');" & vbCrLf)
			'Response.Write("    if (sDatabase != """" && sDatabase != null && sDatabase != ""undefined"") {" & vbCrLf)
			'Response.Write("      frmLoginForm.txtDatabase.value = sDatabase;" & vbCrLf)
			'Response.Write("    }" & vbCrLf)
		End If

		If Request.QueryString("server") <> "" Then
			Response.Write("    frmLoginForm.txtServer.value = """ & CleanStringForJavaScript(Request.QueryString("server")) & """;" & vbCrLf)
		ElseIf Session("server") <> "" Then
			Response.Write("    frmLoginForm.txtServer.value = """ & CleanStringForJavaScript(Session("server")) & """;" & vbCrLf)
			'Else
			' Response.Write("    sServer = getCookie('Intranet_Server');" & vbCrLf)
			'	Response.Write("    if (sServer != """" && sServer != null && sServer != ""undefined"") {" & vbCrLf)
			'	Response.Write("      frmLoginForm.txtServer.value = sServer;" & vbCrLf)
			'	Response.Write("    }" & vbCrLf)
		End If

		If Request.ServerVariables("LOGON_USER") <> "" Then
			If Request.QueryString("WindowsAuthentication") <> "" Then
				Response.Write("    frmLoginForm.chkWindowsAuthentication.value = """ & CleanStringForJavaScript(Request.QueryString("WindowsAuthentication")) & """;" & vbCrLf)
			ElseIf Session("WindowsAuthentication") <> "" Then
				Response.Write("    frmLoginForm.chkWindowsAuthentication.value = """ & CleanStringForJavaScript(Session("WindowsAuthentication")) & """;" & vbCrLf)
			Else
				Response.Write("    sWindowsAuthentication = getCookie('Intranet_WindowsAuthentication');" & vbCrLf)
				Response.Write("    if (sWindowsAuthentication == ""True"" && sWindowsAuthentication != null && sWindowsAuthentication != ""undefined"") {" & vbCrLf)
				Response.Write("      frmLoginForm.chkWindowsAuthentication.checked = ""1"";" & vbCrLf)
				Response.Write("      ToggleWindowsAuthentication();" & vbCrLf)
				Response.Write("    }" & vbCrLf)
			End If
		End If

		Response.Write("    if ((frmLoginForm.txtDatabase.value.length == 0) ||" & vbCrLf)
		Response.Write("      (frmLoginForm.txtServer.value.length == 0) || " & vbCrLf)
		Response.Write("			(frmLoginForm.txtSetDetails.value == 1)) {" & vbCrLf)
		Response.Write("      setDetailsDisplay(true);" & vbCrLf)
		Response.Write("    }" & vbCrLf)
		Response.Write("    else {" & vbCrLf)
		Response.Write("	    setDetailsDisplay(false);" & vbCrLf)
		Response.Write("    }" & vbCrLf)

		Response.Write("}")
		Response.Write("</script>" & vbCrLf)
	End If
%>

<script type="text/javascript">	
	function HelpAbout() {
		$("#About").dialog( "open" );
	}

	function SubmitLoginDetails() {
		/* Try to login to the OpenHR database. */
		var sUserName;
		var sPassword;
		var sDatabase;
		var sServer;
		var fLoginOK;
		var sWindowsAuthentication;
		var frmLoginForm = document.getElementById('frmLoginForm');
	
		fLoginOK = true;
		frmLoginForm.txtUserNameCopy.value = frmLoginForm.txtUserName.value;
		sUserName = frmLoginForm.txtUserName.value;
		sUserName = sUserName.toUpperCase();
		sPassword = frmLoginForm.txtPassword.value;
		sDatabase = frmLoginForm.txtDatabase.value;
		sServer = frmLoginForm.txtServer.value;
		sWindowsAuthentication = frmLoginForm.chkWindowsAuthentication.checked;

		if (fLoginOK) {
			if (sUserName == "") {
				alert("The user name is not valid.");
				fLoginOK = false;
			}
		}

		if (fLoginOK) {
			if (sUserName == "SA") {
				alert("The System Administrator cannot use the OpenHR Web module.");
				fLoginOK = false;
			}
		}

		if (fLoginOK) {
			if (sDatabase == "") {
				alert("The database is not valid.");
				fLoginOK = false;
			}
		}

		if (fLoginOK) {
			if (sDatabase.indexOf("'") > 0) {
				alert("The database name contains an apostrophe.");
				fLoginOK = false;
			}
		}

		if (fLoginOK) {
			if (sServer == "") {
				alert("The server is not valid.");
				fLoginOK = false;
			}
		}

		if (fLoginOK) {
			// Save the values used for user name, database and server to the registry.
			//TODO
			setCookie('Intranet_UserName', frmLoginForm.txtUserName.value, 365);
			//setCookie('Intranet_Database', sDatabase, 365);
			//setCookie('Intranet_Server', sServer, 365);
			setCookie('Intranet_WindowsAuthentication', frmLoginForm.chkWindowsAuthentication.checked, 365);

			frmLoginForm.txtLocaleDateFormat.value = OpenHR.LocaleDateFormat();
			frmLoginForm.txtLocaleDecimalSeparator.value = OpenHR.LocaleDecimalSeparator();
			frmLoginForm.txtLocaleThousandSeparator.value = OpenHR.LocaleThousandSeparator();
			frmLoginForm.txtLocaleDateSeparator.value = OpenHR.LocaleDateSeparator();			

			//Splash
			$(".splashDiv").show();
						
			frmLoginForm.submit();			

		}

	}

	function CancelLogin() {
		/* Quit the browser. */
		window.close();
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
			SubmitLoginDetails();			
		}
	}

	function toggleDetails() {
		if (trDetails1.style.visibility == "visible") {
			setDetailsDisplay(false);
		}
		else {
			setDetailsDisplay(true);
			frmLoginForm.txtDatabase.select();
		}
	}

	function DisableUsernamePassword(pfDisable) {
		text_disable(frmLoginForm.txtUserName, pfDisable);
		text_disable(frmLoginForm.txtPassword, pfDisable);
	}

	function ToggleWindowsAuthentication() {
		if (frmLoginForm.chkWindowsAuthentication.checked == true) {
			DisableUsernamePassword(true);
			frmLoginForm.txtUserName.value = frmLoginForm.txtSystemUser.value;
			frmLoginForm.txtPassword.value = "*****";
			$("#ForgotPasswordLink").css("display", "none");
		}
		else {
			DisableUsernamePassword(false);
			frmLoginForm.txtPassword.value = "";
			$("#ForgotPasswordLink").css("display", "block");
		}
	}

	function setDetailsDisplay(pfShow) {
		var sVisibility;
		var sDisplay;

		if (pfShow == true) {
			frmLoginForm.details.value = "Details <<";
			sVisibility = "visible";
			sDisplay = "block";
		}
		else {
			frmLoginForm.details.value = "Details >>";
			sVisibility = "hidden";
			sDisplay = "none";
		}

		var trDetails1 = document.getElementById("trDetails1");

		trDetails1.style.visibility = sVisibility;
		trDetails1.style.display = sDisplay;
		trDetails2.style.visibility = sVisibility;
		trDetails2.style.display = sDisplay;
	}
	
	function updateViews(isKeyboardVisible) {		
		$('.header-banner').toggle();
		$('.ui-widget-header').toggle();
		$('.loginframetheme img').toggle();
		$('.loginframetheme img').toggle();
		$('.verticalpadding200').toggle();
		$('.loginframetheme img').toggle();		
	}	

	function toggleChromeIfAndroid() {
		var is_keyboard = false;
		var is_landscape = false;
		var initial_screen_size = window.innerHeight;
		/* Android */
		var ua = navigator.userAgent.toLowerCase();		
		var isAndroid = ua.indexOf("android") > -1; //&& ua.indexOf("mobile");
		if (isAndroid) {
			//remove some padding
			$('.verticalpadding200').hide();
			$('.android-padding').toggle();
			window.addEventListener("resize", function () {
				is_keyboard = (window.innerHeight < initial_screen_size);
				is_landscape = (screen.height < screen.width);
				updateViews(is_keyboard);
			}, false);
		}
	}

</script>

	<div class="COAwallpapered ui-widget-content ui-widget">

		<%Html.BeginForm("Login", "Account", FormMethod.Post, New With {.id = "frmLoginForm"})%>
		<table class="ui-dialog-titlebar ui-widget-header" style="margin: 0 auto; width: 100%">
			<tr>
				<td>
					<table border="0" cellspacing="0" cellpadding="0" height="100%" width="100%">
						<tr style="height: 40px">
							<td align="right">
								<img src="<%= Url.Content("~/Content/images/help32.png")%>" width="32" height="32" align="absbottom" onclick="HelpAbout();" />
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<div class="verticalpadding200"></div>
		<div class="" style="">
			<table cellspacing="0" cellpadding="0" align="center" class="invisible loginframetheme ui-widget-content">
				<tr>
					<td width="15"></td>
					<td colspan="3">
						<p align="center">
							<img height="188" src="<%= Url.Content("~/Content/images/OpenHRWeb_Splash.png")%>" width="410"></p>
					</td>
					<td width="15"></td>
				</tr>
				<tr class="android-padding" height="10">
					<td width="15"></td>
					<td colspan="3" style="font-weight: bold;" align="center">Version <%=session("Version")%></td>
					<td width="15"></td>
				</tr>
				<tr height="10" class="android-padding">
					<td colspan="5"></td>
				</tr>
			</table>

			<table id="ancientBrowser" cellspacing="0" cellpadding="0" align="center" class="hidden invisible loginframetheme ui-widget-content">
				<tr height="10">
					<td colspan="5"></td>
				</tr>
				<tr height="10">
					<td width="15"></td>
					<td colspan="3">OpenHR Web can only be accessed using Microsoft Internet Explorer 10 or later.</td>
					<td width="15"></td>
				</tr>
			</table>
			<%

				If Len(Session("version")) = 0 Then
			%>
			<table id="noAppVersion" cellspacing="0" cellpadding="0" align="center" class="invisible loginframetheme ui-widget-content">
				<tr height="10">
					<td colspan="5"></td>
				</tr>
				<tr class="" height="10">
					<td width="15"></td>
					<td style="font-weight: bold;" colspan="3">Unable to determine the OpenHR version.</td>
					<td width="15"></td>
				</tr>
				<tr class="" height="10">
					<td width="15"></td>
					<td style="font-weight: bold;" colspan="3">Ensure that a virtual directory has been configured on your web server.</td>
					<td width="15"></td>
				</tr>
			</table>
			<%
			Else
			%>
			<table id="loginFrame" cellspacing="0" cellpadding="0" align="center" class="invisible loginframetheme ui-widget-content">
				<tr style="height: 10px">
					<td style="height: 15px"></td>
					<td colspan="3" align="center">
						<table style="border: 0px; border-spacing: 0px; border-collapse: collapse;">
							<tr class="" style="display: block;">
								<td style="font-weight: bold; width: 120px; text-align: left;">User name :</td>
								<td style="width: 10px"></td>
								<td style="width: 200px;">
									<input id="txtUserName" autocomplete="off" autocorrect="off" name="txtUserName" class="text" style="height: 22px; width: 100%;" onkeypress="CheckKeyPressed(event)" />
									<input type="hidden" id="txtUserNameCopy" name="txtUserNameCopy" />
								</td>

							</tr>
							<tr class="" style="display: block;">
								<td style="font-weight: bold; width: 120px; text-align: left;">Password :</td>
								<td style="width: 10px"></td>
								<td style="width: 200px;">

									<input id="txtPassword" name="txtPassword" type="password" class="text" style="height: 22px; width: 100%;" onkeypress="CheckKeyPressed(event);" />
								</td>
							</tr>

							<tr class="">
								<%
									If Request.ServerVariables("LOGON_USER") <> "" Then
								%>
								<td style="font-weight: bold; text-align: left;" colspan="3">
									<input id="chkWindowsAuthentication" name="chkWindowsAuthentication" type="checkbox" tabindex="0"
										onclick="ToggleWindowsAuthentication()" />
									<label
										for="chkWindowsAuthentication"
										class="checkbox"
										tabindex="0"
										onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
										onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
										onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
										onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
										onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
										Use Windows Authentication
									</label>
								</td>
								<td></td>
								<td></td>
								<%
								Else
								%>
								<td class="" colspan="3">
									<input type="hidden" id="chkWindowsAuthentication" name="chkWindowsAuthentication" type="checkbox" />
								</td>
								<%
								End If
								%>
							</tr>

							<tr class="" style="visibility: hidden; display: none" id="trDetails1">
								<td style="width: 120px; font-weight: bold; text-align: left;">
									<div id="DatabaseTextLabelDiv" runat="server">Database :</div>
								</td>
								<td style="width: 10px"></td>
								<td style="width: 200px;">
									<div id="DatabaseTextValueDiv" runat="server">
										<input id="txtDatabase" autocomplete="off" autocorrect="off" name="txtDatabase" style="height: 22px; width: 100%;" class="text" onkeypress="CheckKeyPressed(event)" value="<%=_txtDatabaseValue%>" />
									</div>
								</td>
							</tr>

							<tr class="" style="visibility: hidden; display: none" id="trDetails2">
								<td style="width: 120px; font-weight: bold; text-align: left;">
									<div id="ServerTextLabelDiv" runat="server">Server :</div>
								</td>
								<td style="width: 10px"></td>
								<td style="width: 200px;">
									<div id="ServerTextValueDiv" runat="server">
										<input id="txtServer" autocomplete="off" autocorrect="off" name="txtServer" style="height: 22px; width: 100%;" class="text" onkeypress="CheckKeyPressed(event)" value="<%=_txtServerValue%>" />
									</div>
								</td>
							</tr>
						</table>
					</td>
					<td style="width: 15px"></td>
				</tr>

				<tr height="10">
					<td colspan="5"></td>
				</tr>

				<tr height="10">
					<td width="15"></td>
					<td colspan="3">
						<table border="0" cellspacing="0" cellpadding="0" align="center">
							<tr>
								<td align="center">
									<input type="button" id="submitLoginDetails" name="submitLoginDetails" class="ui-button" style="width: 90px;"
										onclick="SubmitLoginDetails()" value="Login" />
								</td>
								<td width="10"></td>
								<td align="center">
									<%--<input type="button" id="cancel" name="cancel" class="ui-button" style="width: 90px;" onclick="CancelLogin()" value="Cancel"/>--%>
								</td>
								<td width="10"></td>
								<td align="center">
									<div id="btnToggleDetailsDiv" runat="server">
										<input type="button" id="details" name="details" class="ui-button" style="" onclick="toggleDetails()" value="Details" />
									</div>
								</td>
							</tr>
						</table>
					</td>
					<td width="15"></td>
				</tr>
			</table>
			<%
			End If

			%>
			<table id="forgotPassword" cellspacing="0" cellpadding="0" align="center" class="invisible loginframetheme ui-widget-content">
				<tr height="10">
					<td colspan="5"></td>
				</tr>
				<tr height="5">
					<td colspan="5"></td>
				</tr>
				<tr height="10">
					<td width="15"></td>
					<td colspan="2">
						<p id="ForgotPasswordLink" style="display: none; text-align: center"><%=Html.ActionLink("Forgot password", "ForgotPassword", "Account")%></p>
					</td>
					<td width="15"></td>
				</tr>
			</table>
		</div>
	<input type="hidden" id="txtSetDetails" name="txtSetDetails" value="<%=Session("showLoginDetails")%>">
	<input type="hidden" id="txtLocaleDateFormat" name="txtLocaleDateFormat">
	<input type="hidden" id="txtLocaleDateSeparator" name="txtLocaleDateSeparator">
	<input type="hidden" id="txtLocaleDecimalSeparator" name="txtLocaleDecimalSeparator">
	<input type="hidden" id="txtLocaleThousandSeparator" name="txtLocaleThousandSeparator">
	<input type="hidden" id="txtSystemUser" name="txtSystemUser" value="<%=replace(Request.ServerVariables("LOGON_USER"),"/","\")%>">
	<input type="hidden" id="txtWordVer" name="txtWordVer" value="12">
	<input type="hidden" id="txtExcelVer" name="txtExcelVer" value="12">
	<input type="hidden" id="txtMSBrowser" name="txtMSBrowser" value="false" />


	<script type="text/javascript">		
		//Set MS browser flag
		if ("ActiveXObject" in window) {			
			document.getElementById("txtMSBrowser").value = 'true';
		}
		
		//Is this a browser that supports file API; which is OK for all modern browsers (IE10+ etc)
		if (!(window.File && window.FileReader && window.FileList && window.Blob)) { 
			//Show 'browser not supported' message...
			$('#ancientBrowser').removeClass('hidden');
			$('#loginFrame').addClass('hidden');
		}
		else {
			//This browser meets requirements. Do what we need to do...
			var fgpl = document.getElementById('ForgotPasswordLink');
			fgpl.style.display = 'block';
			
			window_onload();
			window.onunload = function () { };

		}

		toggleChromeIfAndroid();


	</script>
	<%	Html.EndForm()%>
</div>

<div class="splashDiv hidden"></div>


</asp:Content>

