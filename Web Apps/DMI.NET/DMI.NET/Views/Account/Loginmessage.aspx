<%@ Page Title="" Language="VB" Inherits="System.Web.Mvc.ViewPage" MasterPageFile="~/Views/Shared/Site.Master" %>
<%@ Import Namespace="DMI.NET" %>


<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
    <%= GetPageTitle("Login") %>    
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
	
	<script type="text/javascript">
		function GoBack() {
			window.location = "Login";
			return;
		}
	</script>

	<div class="COAwallpapered" <%=session("BodyTag")%> style="top: 190px; position: absolute;">
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
														<input type="button" class="btn" value="OK" name="GoBack" style="height: 33px; width: 100px" id="cmdGoBack" onclick="GoBack();" />
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
</div>

</asp:Content>