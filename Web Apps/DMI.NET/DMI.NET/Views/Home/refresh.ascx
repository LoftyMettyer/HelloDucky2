<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<%
	Dim sReferringPage

	' Only open the form if the referring page was the main or data page.
	' If it wasn't then redirect to the login page.
	sReferringPage = Request.ServerVariables("HTTP_REFERER") 
	if inStrRev(sReferringPage, "/") > 0 then
		sReferringPage = mid(sReferringPage, inStrRev(sReferringPage, "/") + 1)
	end if

	if (ucase(sReferringPage) <> ucase("main.asp")) and (ucase(sReferringPage) <> ucase("refresh.asp")) then
		Response.Redirect("login.asp")
	end if
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="refresh" content="<%=session("TimeoutSecs")%>;URL=timeout.asp">
<LINK href="OpenHR.css" rel=stylesheet type=text/css >
</HEAD>

<BODY>
<FORM action="refresh.asp" method=post id=frmRefresh name=frmRefresh>
</FORM>
</BODY>
</HTML>
