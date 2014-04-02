<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage (Of DMI.NET.Models.PollMessageModel)" %>

<div class="left">
	<form id="frmPollMessage" action="pollMessage">

		<div style="height:80%">
			<br/>		
		<%: Html.DisplayFor(Function(m) m.Body)%>
		</div>
		
		<div class="ui-dialog-buttonpane ui-widget-content ui-helper-clearfix">
			<div class="ui-dialog-buttonset">
				<%If Model.IsTimedOut Then%>
					<input type="button" value="OK" onclick="pollmessage_logout();"/>
				<%Else%>
					<input type="button" value="OK" onclick="pollmessage_ok();"/>
				<%End If%>
			</div>
		</div>			
		<input type="hidden" id="txtIsSessionTimeout"/>

	</form>
</div>

<script type="text/javascript">
	
	function pollmessage_logout() {
		window.onbeforeunload = null;
		try { window.location.href = "Main"; } catch (e) { }
		return false;
	}

	function pollmessage_ok() {
		$("#divPollMessage").dialog("close");
		return false;
	}

	<%	If Model.Body.Length > 0 Then%>
		$("#divPollMessage").dialog("open");
	<% end if %>

	$("#divPollMessage").dialog({ dialogClass: 'no-close' }, 'option', 'title', '<%=Model.Caption%>' );

</script>

