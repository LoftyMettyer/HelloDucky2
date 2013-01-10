<%@ Page Title="" Language="VB" Inherits="System.Web.Mvc.ViewPage" MasterPageFile="~/Views/Shared/Site.Master" %>

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

	if len(sReferringPage) = 0 then
		Response.Redirect("Login")
	end if
%>

<%--TODO <SCRIPT FOR=window EVENT=onload LANGUAGE=JavaScript>
	cmdGoBack.focus();
</SCRIPT>--%>

<script type="text/javascript">
<!--
	/* Go back to the previous page. */
	function GoBack() {
		var frmGotoMain = document.getElementById('frmGotoMain');
		frmGotoMain.submit();
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
