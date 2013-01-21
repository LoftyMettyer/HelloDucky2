<%@ Page Title="" Language="VB" Inherits="System.Web.Mvc.ViewPage" MasterPageFile="~/Views/Shared/Site.Master" %>
<%@ Import Namespace="DMI.NET" %>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
    <%= GetPageTitle("Login") %>    
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

<%	
	On Error Resume Next            
    
	Dim sBrowserInfo As String
	  
	' Ensure the database connection object is closed.
	Dim conX = Session("databaseConnection")
	conX.Close()
	conX = Nothing

	Session("databaseConnection") = Nothing
	session("action") = ""
	session("selectSQL") = ""
	session("filterSQL") = ""
	session("filterDef") = ""
	session("optionAction") = ""

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
	
	' Check what browser is being used.
	sBrowserInfo = Request.ServerVariables("HTTP_USER_AGENT")
	if InStr(sBrowserInfo, "MSIE") then
		' Microsoft browser.
		sBrowserInfo = mid(sBrowserInfo, InStr(sBrowserInfo, "MSIE") + 5)

		if inStr(sBrowserInfo, ".") > 0 then
			sBrowserInfo = left(sBrowserInfo, instr(sBrowserInfo, ".") + 1)
		end if

		session("MSBrowser") = true
		Session("IEVersion") = CDbl(sBrowserInfo)
	else		
		' Non Microsoft browser.
		session("MSBrowser") = false
		session("IEVersion") = 0
	end if	

	session("dfltTempMenuFilePath") = "<NONE>"

	If (Len(Session("Version")) > 0) Then

		Response.Write("<script type='text/javascript'>" & vbCrLf)
		Response.Write("	function window_onload() {" & vbCrLf)
		Response.Write("    var sUserName;" & vbCrLf)
		Response.Write("    var sDatabase;" & vbCrLf)
		Response.Write("	  var sServer;" & vbCrLf)
		Response.Write("    var sWindowsAuthentication;" & vbCrLf)
		'if Request.QueryString("msoffice") = "" then
		'    Response.Write "    frmLoginForm.txtWordVer.value = ASRIntranetFunctions.GetOfficeWordVersion();" & vbcrlf
		'    Response.Write "    frmLoginForm.txtExcelVer.value = ASRIntranetFunctions.GetOfficeExcelVersion();" & vbcrlf
		'end if
		'Response.Write("alert(window.isMobileBrowser);" & vbCrLf)
		If Request.QueryString("user") <> "" Then
			Response.Write("    frmLoginForm.txtUserName.value = """ & cleanStringForJavaScript(Request.QueryString("user")) & """;" & vbCrLf)
			Response.Write("    if(!window.isMobileBrowser) frmLoginForm.txtPassword.focus();" & vbCrLf)
		ElseIf Request.QueryString("username") <> "" Then
			Response.Write("    frmLoginForm.txtUserName.value = """ & cleanStringForJavaScript(Request.QueryString("username")) & """;" & vbCrLf)
			Response.Write("    if(!window.isMobileBrowser) frmLoginForm.txtPassword.focus();" & vbCrLf)
		ElseIf Session("username") <> "" Then
			Response.Write("    frmLoginForm.txtUserName.value = """ & cleanStringForJavaScript(Session("username")) & """;" & vbCrLf)
			Response.Write("    if(!window.isMobileBrowser) frmLoginForm.txtPassword.focus();" & vbCrLf)
		Else
			If Not Request.Cookies("Login") Is Nothing Then
				Response.Write("    sUserName = '" & Server.HtmlEncode(Request.Cookies("Login")("User")) & "' ;" & vbCrLf)
				Response.Write("    sDatabase = '" & Server.HtmlEncode(Request.Cookies("Login")("Database")) & "' ;" & vbCrLf)
				Response.Write("    sServer = '" & Server.HtmlEncode(Request.Cookies("Login")("Server")) & "' ;" & vbCrLf)
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
			Response.Write("    frmLoginForm.txtDatabase.value = """ & cleanStringForJavaScript(Request.QueryString("database")) & """;" & vbCrLf)
		ElseIf Session("database") <> "" Then
			Response.Write("    frmLoginForm.txtDatabase.value = """ & cleanStringForJavaScript(Session("database")) & """;" & vbCrLf)
		Else
			Response.Write("    sDatabase = getCookie('Intranet_Database');" & vbCrLf)
			Response.Write("    if (sDatabase != """" && sDatabase != null && sDatabase != ""undefined"") {" & vbCrLf)
			Response.Write("      frmLoginForm.txtDatabase.value = sDatabase;" & vbCrLf)
			Response.Write("    }" & vbCrLf)
		End If

		If Request.QueryString("server") <> "" Then
			Response.Write("    frmLoginForm.txtServer.value = """ & cleanStringForJavaScript(Request.QueryString("server")) & """;" & vbCrLf)
		ElseIf Session("server") <> "" Then
			Response.Write("    frmLoginForm.txtServer.value = """ & cleanStringForJavaScript(Session("server")) & """;" & vbCrLf)
		Else
			Response.Write("    sServer = getCookie('Intranet_Server');" & vbCrLf)
			Response.Write("    if (sServer != """" && sServer != null && sServer != ""undefined"") {" & vbCrLf)
			Response.Write("      frmLoginForm.txtServer.value = sServer;" & vbCrLf)
			Response.Write("    }" & vbCrLf)
		End If

		If Request.ServerVariables("LOGON_USER") <> "" Then
			If Request.QueryString("WindowsAuthentication") <> "" Then
				Response.Write("    frmLoginForm.chkWindowsAuthentication.value = """ & cleanStringForJavaScript(Request.QueryString("WindowsAuthentication")) & """;" & vbCrLf)
			ElseIf Session("WindowsAuthentication") <> "" Then
				Response.Write("    frmLoginForm.chkWindowsAuthentication.value = """ & cleanStringForJavaScript(Session("WindowsAuthentication")) & """;" & vbCrLf)
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
		'Response.Write("-->" & vbCrLf)
		Response.Write("</SCRIPT>" & vbCrLf)
	End If
%>

<script type="text/javascript">
<!--

	function HelpAbout() {
		window.location.href = "<%: Url.Action("AboutHRPro", "Account") %>";
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
				alert("The System Administrator cannot use the OpenHR Intranet module.");
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
			setCookie('Intranet_Database', sDatabase, 365);
			setCookie('Intranet_Server', sServer, 365);
			setCookie('Intranet_WindowsAuthentication', frmLoginForm.chkWindowsAuthentication.checked, 365);

			frmLoginForm.txtLocaleDateFormat.value = '<% =LocaleDateFormat %>';
			frmLoginForm.txtLocaleDecimalSeparator.value = '<% =LocaleDecimalSeparator %>';
			frmLoginForm.txtLocaleThousandSeparator.value = '<% =LocaleThousandSeparator %>';
			frmLoginForm.txtLocaleDateSeparator.value = '<% =LocaleDateSeparator %>';			

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

		if (keynum == 13) {
			// 13 = enter key
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
		}
		else {
			DisableUsernamePassword(false);
			frmLoginForm.txtPassword.value = "";
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


-->
</script>

<div style="scroll:auto" class="COAwallpapered" leftmargin="0" topmargin="0" rightmargin="0" marginwidth="0" marginheight="0">
		
<%	Html.BeginForm("Login", "Account", FormMethod.Post, New With {.id = "frmLoginForm"})%>

<table width="100%" border="0" cellspacing="0" cellpadding="0" height="44px" style="border-bottom:solid 1px;border-color:#992b99">
  <tr> 
    <td align=left> 
		 <IMG SRC="<%= Url.Content("~/Content/images/COAInt_TopBar.jpg")%>" align=top width="138" height="44">
    </td>
    <td height="44px" width="100%" align=right> 
			<IMG style="display: none;" src="<%= Url.Content("~/Content/images/coaint_banner.jpg")%>" align=top width="600" height="44">
    </td>
  </tr>
  <tr height="10"> 
	  <td>
	  </td>
  </tr>  
</table>
<table class="headerbar" width="100%" border="0" cellspacing="0" cellpadding="0" height="40" style="border-bottom:solid 1px;border-color:#992b99">
	<tr> 
		<td>
			<table border="0" cellspacing="0" cellpadding="0" height="100%" width="100%">
				<tr style="height:40"> 
			    <td align=right>
            <img src="<%= Url.Content("~/Content/images/help32.png")%>" width="32" height="32" align=absbottom onclick="HelpAbout();"
            onmouseover="this.style.cursor='pointer'; this.style.color='<%=session("Config-hypertextlinktext-highlightcolour")%>';" 
													 onmouseout="this.style.color=''" 
            onclick="HelpAbout();">			      
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<div class="verticalpadding200"></div>
    <table CELLSPACING="0" CELLPADDING="0" align="center" class="invisible loginframetheme" >
        <tr> 
            <TD width=15 height="15"></TD>
            <td colSpan=3 height="15">&nbsp;</td>
            <TD width=15 height="15"></TD>
        </tr>
        <tr>
          <td height=40></td>
        </tr>
        <tr> 
            <TD width=15></TD>
            <td colSpan=3> 
                <p align="center"><IMG height=188 src="<%= Url.Content("~/Content/images/COAInt_Splash.jpg")%>" width=410></p>
            </td>
            <TD width=15></TD>
        </tr>
	    <tr height=10>
		    <td colSpan=5></td>
    	</tr>
	    <tr  class="logintext" height=10>
		    <TD width=15></TD>
		    <td colSpan=3 style="font-weight: bold;" align=center>Version <%=session("Version")%></td>
		    <TD width=15></TD>
	    </tr>
	    <tr height=10>
		    <td colSpan=5></td>
	    </tr>
<%
	If (Session("MSBrowser") = True) And (Session("IEVersion") < 8.0) Then
%>
        <tr height=10>
	        <td colSpan=5></td>
	    </tr>
	    <tr height=10>
		    <td width=15></td>
		    <td colSpan=3>OpenHR Intranet can only be accessed using Microsoft Internet Explorer 8 or later.</td>
    		<td width=15></td>
	    </tr>
<%
Else
	If Len(Session("version")) = 0 Then
%>
	    <tr height=10>
		    <td colSpan=5></td>
	    </tr>
	    <tr class="logintext" height=10>
		    <td width=15></td>
		    <td style="font-weight: bold;"  colSpan=3 >Unable to determine the intranet version.</td>
		    <td width=15></td>
	    </tr>
	    <tr  class="logintext" height=10>
		    <td width=15></td>
		    <td style="font-weight: bold;" colSpan=3 >Ensure that a virtual directory has been configured on your intranet web server.</td>
		    <td width=15></td>" & vbcrlf
	    </tr>
<%
		else
%>
        <tr style="height: 10px">
		    <td style="height: 15px"></td>
		    <td colspan="3">
			    <table style="border:0px; border-spacing: 0px; border-collapse: collapse;">
			        <tr class="logintext" style="display: block">
			            <td style="font-weight:bold ;width: 100px;">User name :</td>
			            <td style="width: 10px"></td>
			            <td style="width: 200px;">
								<input id="txtUserName" autocomplete="off" autocorrect="off" name="txtUserName" class="text" style="height: 22px;width: 100%; padding: 0px;" onkeypress="CheckKeyPressed(event)"/>
				        <input type="hidden" id="txtUserNameCopy" name="txtUserNameCopy" />    
			            </td>
			            
        		    </tr>
    			    <tr class="logintext" style="display: block">
	    		        <td style="font-weight:bold ;width: 100px;">Password :</td>
			            <td style="width: 10px">
							</td>
			            <td style="width: 200px;">

								<input id="txtPassword" name="txtPassword" type="password" class="text" style="height: 22px; width: 100%; padding: 0px;" onkeypress="CheckKeyPressed(event);"/>
			            </td>
    			    </tr>

	    		    <tr class="logintext" >
<%
			if Request.ServerVariables("LOGON_USER") <> "" then
%>			
		    		    <tdstyle="font-weight: bold;" colspan="3" >
		    		        <input id="chkWindowsAuthentication" name="chkWindowsAuthentication" type="checkbox" tabindex="-1"
		    		            onclick="ToggleWindowsAuthentication()"/> 
				            <label 
				                for="chkWindowsAuthentication"
				                class="checkbox"
				                tabindex=0 
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
			else
%>
				        <td class="logintext" colspan="3" >
				            <input type="hidden" id="chkWindowsAuthentication" name="chkWindowsAuthentication" type="checkbox" />
				        </td>
<%
			end if
%>
                 </tr>

			        <tr class="logintext" style="display: block;visibility:hidden;display:none" id="trDetails1">
			            <td style="width: 100px;font-weight: bold;">Database :</td>
			            <td style="width: 10px">
							</td>
			            <td style="width: 200px;">
			                <input id="txtDatabase"  autocomplete="off" autocorrect="off"  name="txtDatabase" style="height: 22px; width: 100%; padding: 0px;" class="text" onkeypress="CheckKeyPressed(event)" />
			            </td>
			        </tr>
			        <tr class="logintext" style="display: block;visibility:hidden;display:none" id="trDetails2">
			            <td style="width: 100px;font-weight: bold;">Server :</td>
			            <td style="width: 10px">
							</td>
			            <td style="width: 200px;">
			                <input id="txtServer"  autocomplete="off" autocorrect="off"  name="txtServer" style="height: 22px; width: 100%; padding: 0px;" class="text" onkeypress="CheckKeyPressed(event)" />
			            </td>
			        </tr>
			    </table>
	        </td>
		    <td style="width:15px"></td>
	    </tr>
	
	    <tr height=10>
	        <td colSpan=5></td>
	    </tr>

    <tr height="10">
        <td width="15"></td>
        <td colspan="3">
            <table border="0" cellspacing="0" cellpadding="0" align="center">
                <tr>
                    <td align="center">
                        <input id="submitLoginDetails" name="submitLoginDetails" type="button" class="btn loginbuttontext" value="OK" style="WIDTH: 90px"
                            onclick="SubmitLoginDetails()"
                            onmouseover="try{button_onMouseOver(this);}catch(e){}"
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
                    </td>
                    <td width="10"></td>
                    <td align="center">
                        <input id="cancel" name="cancel" type="button" value="Cancel" class="btn loginbuttontext" style="WIDTH: 90px"
                            onclick="CancelLogin()"
                            onmouseover="try{button_onMouseOver(this);}catch(e){}"
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
                    </td>
                    <td width="10"></td>
                    <td align="center">
                        <input id="details" name="details" type="button" value="Details" class="btn loginbuttontext" style="WIDTH: 90px"
                            onclick="toggleDetails()"
                            onmouseover="try{button_onMouseOver(this);}catch(e){}"
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
                    </td>
                </tr>
            </table>
        </td>
        <td width="15"></td>
    </tr>
<%
End If
End If
%>

	    <tr height=10>
		    <td colSpan=5></td>
	    </tr>
	    <tr height=5>
		    <td colSpan=5></td>
	    </tr>   	
	    <tr height=10>
		    <td colSpan=5></td>
	    </tr>
    </table>

	<INPUT type="hidden" id=txtSetDetails name=txtSetDetails value="<%=Session("showLoginDetails")%>">
   <INPUT type="hidden" id=txtLocaleDateFormat name=txtLocaleDateFormat>
   <INPUT type="hidden" id=txtLocaleDateSeparator name=txtLocaleDateSeparator>
   <INPUT type="hidden" id=txtLocaleDecimalSeparator name=txtLocaleDecimalSeparator>
   <INPUT type="hidden" id=txtLocaleThousandSeparator name=txtLocaleThousandSeparator>
   <INPUT type="hidden" id=txtSystemUser name=txtSystemUser value="<%=replace(Request.ServerVariables("LOGON_USER"),"/","\")%>">
	<INPUT type="hidden" id=txtWordVer name=txtWordVer value="12">
	<INPUT type="hidden" id=txtExcelVer name=txtExcelVer value="12">

<%if (session("MSBrowser") = False) or (session("MSBrowser") = true) and (session("IEVersion") > 7.0) then%>

	<script type="text/javascript">
		

		window_onload();
		

		window.onunload = function () { };
	</script>	

<%end if %>

<%	Html.EndForm()%>
</div>

<div class="splashDiv hidden"></div>


</asp:Content>
