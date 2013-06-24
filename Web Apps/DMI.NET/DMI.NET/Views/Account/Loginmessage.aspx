<%@ Page Title="" Language="VB" Inherits="System.Web.Mvc.ViewPage" MasterPageFile="~/Views/Shared/Site.Master" %>
<%@ Import Namespace="DMI.NET" %>


<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
    <%= GetPageTitle("Login") %>    
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
<script type="text/javascript">
	/* Go back to the previous page. */
	function GoBack() {
        
	    //loginmessage is always called so back to login.
	    window.location.href("login");

	    return false;
        if (InStrRev(document.referrer, "/") > 0) {
            var sReferringPage = (Mid(document.location, (InStrRev(document.location, "/") + 1), 255));
            if (sReferringPage.length > 0 && sReferringPage.toLowerCase() != "login" && sReferringPage.toLowerCase() != "forcedpasswordchange") {
                //Not referred from login page, so default behaviour
                //window.history.back(2);
                window.location.href("Main");
            }
            else {
                //referred from login page, so return to default.asp
                window.location = "Login";
            }
        }
        else {
            //window.history.back(2);
            window.location.href("Main");
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

<div <%=session("BodyTag")%>>

    <table align="center" class="outline" cellpadding="5" cellspacing="0">
        <tr>
            <td>
                <table align="center" class="invisible" cellpadding="0" cellspacing="0" width="100%" height="100%">
                    <tr>
                        <td width="20"></td>
                        <td align="center">
						    <h3><%=Session("MessageTitle")%></h3>
                        </td>
                        <td width="20"></td>
                    </tr>
                    <tr>
                        <td width="20"></td>
                        <td align="center">
						    <%=Session("MessageText")%>
                        </td>
                        <td width="20"></td>
                    </tr>
                    <tr>
                        <td height="20" colspan="3"></td>
                    </tr>
                    <tr>
                        <td align="center" colspan="3">
                            <input type="button" class="btn" value="OK" name="GoBack" style="HEIGHT: 24px; WIDTH: 100px" width="100" id="cmdGoBack" onclick="GoBack()" />
                        </td>
                    </tr>
                    <tr>
                        <td height="10" colspan="3"></td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>

    <form action="main" method="post" id="frmGotoMain" name="frmGotoMain">
    </form>

    <input type="hidden" id="txtDesktopColour" name="txtDesktopColour" value="<%=session("DesktopColour")%>">

</div>

</asp:Content>