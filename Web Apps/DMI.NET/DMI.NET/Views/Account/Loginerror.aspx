<%@ Page Title="" Language="VB" Inherits="System.Web.Mvc.ViewPage" MasterPageFile="~/Views/Shared/Site.Master" %>
<%@ Import Namespace="DMI.NET" %>


<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
    <%= GetPageTitle("Login") %>    
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

<%      
	' Only open the form if there was a referring page.
	' If it wasn't then redirect to the login page.
	'Dim sReferringPage = Request.ServerVariables("HTTP_REFERER")
	'if inStrRev(sReferringPage, "/") > 0 then
	'	sReferringPage = mid(sReferringPage, inStrRev(sReferringPage, "/") + 1)
	'end if

	'If (InStr(1, UCase(sReferringPage), UCase("Login")) < 1) And UCase(sReferringPage) <> UCase("ForcedPasswordChange") Then
	'	Response.Redirect("Login")
	'End If
%>

<script type="text/javascript">
	function Loginerror_window_onload() {
		window.cmdGoBack.focus();
	}
</script>

<script type="text/javascript">
    /* Go back to the previous page. */
	function GoBack() {
        
		//Loginerror is always from login. Return to login.
		window.location = "Login";
		return;
        
		if (InStrRev(document.referrer, "/") > 0) {
			var sReferringPage = (Mid(document.referrer, (InStrRev(document.referrer, "/") + 1), 255));
			if (sReferringPage.length > 0 && sReferringPage.toLowerCase() != "login" && sReferringPage.toLowerCase() != "forcedpasswordchange") {
				//Not referred from login page, so default behaviour
				window.history.back(2);
			}
			else {
				//referred from login page, so return to default.asp
				window.location = "Login";
			}
		}
		else {
			window.history.back(2);
		}
	}

	function InStrRev(strSearch, charSearchFor) {
		var j = -1;
		for (var i = 0; i < strSearch.length; i++) {
			if (charSearchFor == Mid(strSearch, i, 1)) {
				j = i;
			}
		}
		if (j > 0) {
			return j;
		}
		else {
			return -1;
		}
	}

	function Mid(str, start, len) {
		if (start < 0 || len < 0) return "";
		var iEnd, iLen = String(str).length;
		if (start + len > iLen)
			iEnd = iLen;
		else
			iEnd = start + len;
		return String(str).substring(start, iEnd);
	}

</script>
	<div class="divLogin">

		<div class="ui-dialog-titlebar ui-widget-header loginTitleBar">
		<img alt="about OpenHR" title="About OpenHR Web" src="<%= Url.Content("~/Content/images/help32.png")%>" />
	</div>

<div class="COAwallpapered" <%=session("BodyTag")%> style="top: 190px; position: absolute;">

    <table class="outline" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td>
                <table class="invisible" cellspacing="0" cellpadding="0">
                    <tr>
                        <td colspan="3" height="10"></td>
                    </tr>

                    <tr>
                        <td colspan="3" align="center">
                            <h3>You could not login to the OpenHR database because of the following reason: </h3>
                        </td>
                    </tr>

                    <tr>
                        <td width="20" height="10"></td>
                        <td>
                            <%=Replace(Session("ErrorText"), vbCr, "<br />")%>
                        </td>
                        <td width="20"></td>
                    </tr>

                    <tr>
                        <td colspan="3" height="10"></td>
                    </tr>

                    <tr>
                        <td colspan="3" height="10" align="center">
                            <input type="button" value="Retry" name="GoBack" class="btn" style="WIDTH: 80px" width="80" id="cmdGoBack" onclick="GoBack()" />
                        </td>
                    </tr>

                    <tr>
                        <td colspan="3" height="10"></td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>

    <input type="hidden" id="txtDesktopColour" name="txtDesktopColour" value="<%=session("DesktopColour")%>">
</div>
		</div>
<script type="text/javascript">
	Loginerror_window_onload();
	//Set up button click events
	$('.loginTitleBar img').click(function () {
		OpenHR.showAboutPopup();
	});
</script>

</asp:Content>