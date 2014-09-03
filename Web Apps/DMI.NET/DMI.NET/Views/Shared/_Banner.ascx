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
    <input type="hidden" id="signalRUsersCount" />
    <input type="hidden" id="signalRMessage" />
  </div>

	<script type="text/javascript">

		$(function () {

			// Activity Hub
			var licence = $.connection.LicenceHub;

			// Create a function that the hub can call back to display messages.
			licence.client.updateUsersOnlineCount = function (count) {
				$('#signalRUsersCount').val(count);
			};

			// System Admin Message
			var hubProxy = $.connection.NotificationHub;
			hubProxy.client.SystemAdminMessage = function (messageFrom, message, forceLogout) {
				$("#SignalRDialogClick").val("Close");
				$("#SignalRDialogTitle").html(messageFrom);
				$("#SignalRDialogContentText").html(message);
				$("#divSignalRMessage").dialog('open');

				if (forceLogout == true) {
					$("#SignalRDialogClick").val("Log Out");
				}

				$("#SignalRDialogClick").off('click').on('click', function () {
					$("#divSignalRMessage").dialog("close");

					if (forceLogout == true) {
						menu_logoffIntranet();
					}

				});

			};

			// Session Timeout Message
			hubProxy.client.SessionTimeOut = function () {

				$("#SignalRDialogTitle").html("Message from Administrator");
				$("#SignalRDialogContentText").html("Your session has timed out. You will need to login again");
				$("#divSignalRMessage").dialog('open');

				$("#SignalRDialogClick").off('click').on('click', function () {
					window.onbeforeunload = null;
					try {
						window.location.href = "Main";
					} catch (e) {
					}
					return false;
				});

			};

			$.connection.hub.start()
					.done(function () { console.log('Now connected, connection ID=' + $.connection.hub.id); })
					.fail(function () { console.log('Could not Connect!'); });

		});
	</script>
	