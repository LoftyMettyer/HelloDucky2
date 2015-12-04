<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<%-- For other devs: Do not remove below line. --%>
<%:"" %>
<%-- For other devs: Do not remove above line. --%>

<%If Session("Config-banner-justification") = "left" Then%>
<div style="float: left;">
	<img src="<%:session("TopBarFile")%>" width="<%:session("Config-banner-graphic-left-width")%>" height="44px" alt=""></div>
<div style="float: left;">
	<img src="<%:Session("LogoFile")%>" width="<%:Session("Config-banner-graphic-right-width")%>" height="44px" alt=""></div>
<%ElseIf Session("Config-banner-justification") = "right" Then%>
<div style="float: right;">
	<img src="<%:session("TopBarFile")%>" width="<%:session("Config-banner-graphic-left-width")%>" height="44px" alt=""></div>
<div style="float: right;">
	<img src="<%:Session("LogoFile")%>" width="<%:Session("Config-banner-graphic-right-width")%>" height="44px" alt=""></div>
<%ElseIf Session("Config-banner-justification") = "justify" Then%>
<div style="float: left;">
	<img src="<%:session("TopBarFile")%>" width="<%:session("Config-banner-graphic-left-width")%>" height="44px" alt=""></div>
<div style="float: right;">
	<img src="<%:Session("LogoFile")%>" width="<%:Session("Config-banner-graphic-right-width")%>" height="44px" alt=""></div>
<%Else
		Dim styleWidth = CInt(Session("Config-banner-graphic-left-width")) + CInt(Session("Config-banner-graphic-right-width")) & "px"%>
<div style="width: <%:styleWidth%>; margin: 0 auto;">
	<div style="float: left;">
		<img src="<%:session("TopBarFile")%>" width="<%:session("Config-banner-graphic-left-width")%>" height="44px" alt=""></div>
	<div style="float: left;">
		<img src="<%:Session("LogoFile")%>" width="<%:Session("Config-banner-graphic-right-width")%>" height="44px" alt=""></div>
</div>
<%End If%>

  <div id="signalRMessaging" class="container">
     <input id="signalRMessage" type="hidden" />
  </div>

	<script type="text/javascript">

		$(function () {

			$.connection.hub.start()
					.done(function() {console.log('Now connected, connection ID=' + $.connection.hub.id);})
					.fail(function (error) { console.log('Could not Connect! ' + error); });

			// Activity Hub
			var licenceHub = $.connection.LicenceHub;
			licenceHub.client.SessionTimeOut = function () {
				OpenHR.SessionTimeout();
			};

			// System/Security Messages
			var notificationHub = $.connection.NotificationHub;

			$.connection.hub.start().done(function () {
				notificationHub.server.joinGroup("<%:Session("Usergroup")%>");
			});

			notificationHub.client.notifyGroup = function (messageFrom, message, forceLogout) {
				OpenHR.displayServerMessage(messageFrom, message, forceLogout, true);
			};

			notificationHub.client.SystemAdminMessage = function (messageFrom, message, forceLogout, loggedInUsersOnly) {
				OpenHR.displayServerMessage(messageFrom, message, forceLogout, loggedInUsersOnly);
			};


		});
	</script>
	