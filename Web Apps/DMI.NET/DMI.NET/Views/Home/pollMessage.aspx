<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage (Of DMI.NET.Models.PollMessageModel)" %>

<div class="centered">
	<form id="frmPollMessage" action="pollMessage">

		<div style="height:80%">		
		<%: Html.DisplayFor(Function(m) m.Body)%>
		</div>
		
		<div class="centered" style="position: absolute; bottom:0; width:90%">
		<%

			If Model.IsTimedOut Then%>
			<input type="button" value="OK" onclick="pollmessage_logout();"/>
			<%		
			Else%>
				<input type="button" value="OK" onclick="pollmessage_ok();"/>
			<%End If%>
	
		</div>

		<input type="hidden" id="txtIsSessionTiemout"/>

	</form>
</div>

<script type="text/javascript">
	
	function pollMessage_logout() {
		menu_logoffIntranet();
		return false;
	}

	function pollmessage_ok() {
		$("#divPollMessage").dialog("close");
		return false;
	}

	<%	If Model.Body.Length > 0 Then%>
		$("#divPollMessage").dialog("open");
	<% end if %>

	$("#divPollMessage").dialog('option', 'title', '<%=Model.Caption%>' );

</script>

