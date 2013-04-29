<%@ Page Title="" Language="VB" MasterPageFile="~/Views/Shared/Site.Master" Inherits="System.Web.Mvc.ViewPage" %>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
downloadControlsStatus
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

<%@ Language=VBScript %>
<!--#INCLUDE FILE="include/ctl_SetStyles.txt" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="OpenHR.css">
<TITLE>OpenHR Intranet</TITLE>
<meta http-equiv="X-UA-Compatible" content="IE=5">
<SCRIPT LANGUAGE=JavaScript>
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
</SCRIPT>

</HEAD>
<%
if len(session("ConvertedDesktopColour")) = 0 then 
    sBGColour = "#f9f7fb" 
else 
    sBGColour = session("ConvertedDesktopColour") 
end if
%>
<BODY TOPMARGIN=8 BOTTOMMARGIN=0 style="background-color:<%=sBGColour%>" >
<FORM ID=frmDownloadStatus NAME=frmDownloadStatus>

<table class="outline" align=center cellPadding=0 cellSpacing=0> 
    <tr>
	    <td>
		    <table class="invisible" align=center cellPadding=0 cellSpacing=0> 
			    <tr height=10>
			        <td colSpan=3 ></td>
			    </tr>
			    <tr>
					<td width=20></td>
			        <td align=center>
				        Downloading OpenHR Intranet Controls.
				    </td>
				    <td width=20></td>
			    </tr>
			    <tr height=10>
			        <td colSpan=3 height=10></td>
			    </tr>
			    <TR>
					<td width=20></td>
					<TD ALIGN=center ID=txtMessage NAME=txtMessage>
				        Please wait...
					</TD>
					<td width=20></td>
				</TR>
			    <tr height=20>
			        <td colSpan=3>&nbsp;</td>
			    </tr>
			    <tr>
					<td width=20></td>
			        <td align=center>
				        <INPUT STYLE="WIDTH: 75px" WIDTH=75 ID=tdButton NAME=tdButton TYPE=button class="btn" VALUE=Cancel 
					        onclick="btnClick()" 
					        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
			                onmouseout="try{button_onMouseOut(this);}catch(e){}"
			                onfocus="try{button_onFocus(this);}catch(e){}"
			                onblur="try{button_onBlur(this);}catch(e){}" />
				    </td>
					<td width=20></td>
                </tr>
			    <tr height=5>
			        <td colSpan=5></td>
			    </tr>
			</table>
		</td>
    </tr>
</table>
<INPUT ID=txtDownloadStatus NAME=txtDownloadStstus TYPE=hidden VALUE=0>
</FORM>
</BODY>
</HTML>


</asp:Content>

<asp:Content ID="Content3" ContentPlaceHolderID="FixedLinksContent" runat="server">
</asp:Content>
