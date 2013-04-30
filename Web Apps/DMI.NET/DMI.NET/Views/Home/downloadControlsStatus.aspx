<%@ Page Title="" Language="VB" MasterPageFile="~/Views/Shared/Site.Master" Inherits="System.Web.Mvc.ViewPage" %>

<%@ Import Namespace="DMI.NET" %>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
	<%= GetPageTitle("downloadControlsStatus")%>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

	<script src="<%: Url.Content("~/Scripts/ctl_SetStyles.js") %>" type="text/javascript"></script>


	<html>
	<head>
		<meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
		<link rel="stylesheet" type="text/css" href="OpenHR.css">
		<title>OpenHR Intranet</title>
		<meta http-equiv="X-UA-Compatible" content="IE=5">
		<script language="JavaScript">
<!--

	function btnClick() 
	{
		if (frmDownloadStatus.txtDownloadStatus.value == 0)
		{
			parent.window.close();
			return;
		}
		else if (frmDownloadStatus.txtDownloadStatus.value == 1)
		{
			parent.window.close();
			return;
		}
		else if (frmDownloadStatus.txtDownloadStatus.value == 2)
		{			
			/* Reload the page. */
			window.parent.frames("downloadControlsFrame").setStatus(0,'Please wait...');
			window.parent.frames("downloadControlsFrame").frmRefresh.submit();
		}
		
		return;
	}
	-->
		</script>

	</head>
	<%
		Dim sBGColour As String
		If Len(Session("ConvertedDesktopColour")) = 0 Then
			sBGColour = "#f9f7fb"
		Else
			sBGColour = Session("ConvertedDesktopColour")
		End If
	%>
	<body topmargin="8" bottommargin="0" style="background-color: <%=sBGColour%>">
		<form id="frmDownloadStatus" name="frmDownloadStatus">

			<table class="outline" align="center" cellpadding="0" cellspacing="0">
				<tr>
					<td>
						<table class="invisible" align="center" cellpadding="0" cellspacing="0">
							<tr height="10">
								<td colspan="3"></td>
							</tr>
							<tr>
								<td width="20"></td>
								<td align="center">Downloading OpenHR Intranet Controls.
								</td>
								<td width="20"></td>
							</tr>
							<tr height="10">
								<td colspan="3" height="10"></td>
							</tr>
							<tr>
								<td width="20"></td>
								<td align="center" id="txtMessage" name="txtMessage">Please wait...
								</td>
								<td width="20"></td>
							</tr>
							<tr height="20">
								<td colspan="3">&nbsp;</td>
							</tr>
							<tr>
								<td width="20"></td>
								<td align="center">
									<input style="WIDTH: 75px" width="75" id="tdButton" name="tdButton" type="button" class="btn" value="Cancel"
										onclick="btnClick()"
										onmouseover="try{button_onMouseOver(this);}catch(e){}"
										onmouseout="try{button_onMouseOut(this);}catch(e){}"
										onfocus="try{button_onFocus(this);}catch(e){}"
										onblur="try{button_onBlur(this);}catch(e){}" />
								</td>
								<td width="20"></td>
							</tr>
							<tr height="5">
								<td colspan="5"></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<input id="txtDownloadStatus" name="txtDownloadStstus" type="hidden" value="0">
		</form>
	</body>
	</html>


</asp:Content>

<asp:Content ID="Content3" ContentPlaceHolderID="FixedLinksContent" runat="server">
</asp:Content>
