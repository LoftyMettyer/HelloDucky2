<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<%
    'Dim sReferringPage

    '' Only open the form if the referring page was the main or data page.
    '' If it wasn't then redirect to the login page.
    'sReferringPage = Request.ServerVariables("HTTP_REFERER") 
    'if inStrRev(sReferringPage, "/") > 0 then
    '	sReferringPage = mid(sReferringPage, inStrRev(sReferringPage, "/") + 1)
    'end if

    'if (ucase(sReferringPage) <> ucase("main.asp")) and (ucase(sReferringPage) <> ucase("refresh.asp")) then
    '	Response.Redirect("login.asp")
    'end if
%>
<html>
<head>
    <title></title>
    <meta http-equiv="refresh" content="1;URL=<%=Url.Action("Login", "Account")%>">
    <link href="OpenHR.css" rel="stylesheet" type="text/css">
</head>

<body>
    <form action="refresh" method="post" id="frmRefresh" name="frmRefresh">
    </form>
</body>
</html>
