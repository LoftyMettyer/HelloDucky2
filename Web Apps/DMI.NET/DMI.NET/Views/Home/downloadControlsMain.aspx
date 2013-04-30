<%@ Page Title="" Language="VB" MasterPageFile="~/Views/Shared/Site.Master" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
	<%= GetPageTitle("downloadControlsMain")%>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
	
	<head>
		<meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
		<link rel="stylesheet" type="text/css" href="OpenHR.css">
		<title>OpenHR Intranet</title>
		<meta http-equiv="X-UA-Compatible" content="IE=5">
	</head>

	<%--<div rows="*,0" frameborder="0" framespacing="0">--%>

	<div id="downloadControlsstatusframe">
		<frame name="downloadControlsstatusframe" src="downloadControlsStatus.asp" scrolling="no" noresize hidefocus>
	<%If Request.Form("txtFromMenu") = "true" Then%>
	<div id="downloadControlsframe" style="display: none;">
		<%Html.RenderPartial("~/views/home/downloadControls.aspx")%>
	</div>
	<%Else%>
	<div id="downloadControlsframe" style="display: none;">
		
		<%Html.RenderPartial("~/views/home/downloadControls.aspx")%>
	</div>
  <%End If%>
	</div>
</asp:Content>

<asp:Content ID="Content3" ContentPlaceHolderID="FixedLinksContent" runat="server">
</asp:Content>


