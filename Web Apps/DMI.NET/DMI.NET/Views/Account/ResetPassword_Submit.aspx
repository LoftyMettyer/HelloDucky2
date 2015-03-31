<%@ Page Title="" Language="VB" MasterPageFile="~/Views/Shared/Site.Master" Inherits="System.Web.Mvc.ViewPage" %>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

	<script type="text/javascript">
		function HelpAbout() {
			$("#About").dialog( "open" );
		}
	</script>

	<div class="divLogin">
		<%Html.BeginForm("ResetPassword_Submit", "Account", FormMethod.Post, New With {.id = "frmResetPasswordForm"})%>
		<div class="ui-dialog-titlebar ui-widget-header loginTitleBar">
		</div>
		<div class="verticalpadding200"></div>

		<div class="ui-widget-content ui-corner-tl ui-corner-br loginframe">
			<table style="margin: 0 auto; width: 1px;">
				<tr>
					<td>
						<img height="188" src="<%=Url.Content("~/Content/images/OpenHRWeb_Splash.png")%>" style="width: 410px;" alt="">
					</td>
				</tr>
				<tr>
					<td style="text-align: center">
						<h3 style="text-align: center;">Reset password</h3>
						<p style="text-align: center;">
							<%=ViewData("Message")%>
						</p>
						<p style="text-align: center;">
							<input type="button" value="Login page" onclick="window.location='<%=Url.Action("Login", "Account")%>	'" style="width: auto;" />
						</p>
					</td>
				</tr>
			</table>
		</div>
	</div>

	<style>
		header {
			height: 48px;
			width: 99.9%;
			z-index: -1;
		}
	</style>

</asp:Content>
