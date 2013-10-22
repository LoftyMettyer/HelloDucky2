<%@ Page Title="" Language="VB" MasterPageFile="~/Views/Shared/Site.Master" Inherits="System.Web.Mvc.ViewPage" %>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">

</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

<div <%=Session("BodyTag")%> style="width: 98%; position: absolute; top: 170px;">
		<table style="margin: 0 auto; width: 1px;">
			<tr> 
				<td> 
						<img height="188" src="<%=Url.Content("~/Content/images/COAInt_Splash.png")%>" style="width: 410px;" alt="">
				</td>
			</tr>
			<tr>
				<td style="text-align: center" > 
					<h2 style="text-align: center;">Forgot passsword</h2>
					<p style="text-align: center;">
						<%=ViewData("Message")%>
					</p>
					<p style="text-align: center;">
							<input type="button" value="<%=ViewData("RedirectToURLMessage")%>" onclick="window.location='<%=Url.Action("Login", "Account")%>';" style="width: auto;" />
					</p>
				</td>
			</tr>
		</table>
	</div>

</asp:Content>
