<%@ Page Title="" Language="VB" Inherits="System.Web.Mvc.ViewPage" MasterPageFile="~/Views/Shared/Site.Master" %>
<%@ Import Namespace="DMI.NET" %>


<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
    <%= GetPageTitle("Login") %>    
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
<script type="text/javascript">
<!--
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
-->
</script>

<div <%=session("BodyTag")%>>

<table align=center class="outline" cellPadding=5 cellSpacing=0> 
    <tr>
	    <td>
			<table align=center class="invisible" cellPadding=0 cellSpacing=0 width=100% height=100%>
				<TR>
					<td width=20></td>
					<TD align=center>
						<H3><%=session("MessageTitle") %></H3>
					</td>
					<td width=20></td>
				</tr>
				<TR>
					<td width=20></td>
					<TD align=center>
						<%=session("MessageText") %>
					</td>
					<td width=20></td>
				</tr>
				<tr>
					<TD height=20 colspan=3></td>
				</tr>
				<tr>
					<TD align=center colspan=3>
						<INPUT TYPE=button class="btn" VALUE="OK" NAME="GoBack" style="HEIGHT: 24px; WIDTH: 100px" width=100 id=cmdGoBack
						    OnClick="GoBack()" 
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
					</td>
				</tr>
				<tr>
					<TD height=10 colspan=3></td>
				</tr>
			</table>
		</td>
	</tr>
</table>

<%--TODO--%>
<FORM action="main.asp" method=post id=frmGotoMain name=frmGotoMain>
</FORM>

<INPUT type="hidden" id=txtDesktopColour name=txtDesktopColour value=<%=session("DesktopColour")%>>

</div>

</asp:Content>