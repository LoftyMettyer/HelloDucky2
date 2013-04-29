<%@ Page Title="" Language="VB" MasterPageFile="~/Views/Shared/Site.Master" Inherits="System.Web.Mvc.ViewPage" %>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
downloadControlsMain
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="OpenHR.css">
<TITLE>OpenHR Intranet</TITLE>
<meta http-equiv="X-UA-Compatible" content="IE=5">
</HEAD>
<FRAMESET ROWS="*,0" frameborder="0" framespacing="0">
	<FRAME NAME=downloadControlsstatusframe SRC="downloadControlsStatus.asp" scrolling="no" noresize hidefocus>
	<%if request.Form("txtFromMenu") = "true" then %>
	  <FRAME NAME=downloadControlsframe SRC="downloadControls.asp?fromMenu=true" scrolling="no" noresize hidefocus>
	<%else %>
	  <FRAME NAME=downloadControlsframe SRC="downloadControls.asp?fromMenu=false" scrolling="no" noresize hidefocus>
  <%end if %>
</FRAMESET>

</HTML>


</asp:Content>

<asp:Content ID="Content3" ContentPlaceHolderID="FixedLinksContent" runat="server">
</asp:Content>
