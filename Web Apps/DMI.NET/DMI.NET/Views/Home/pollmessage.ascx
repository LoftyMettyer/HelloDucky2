﻿<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<script type="text/javascript">

	function pollmessage_window_onload() {
		var sMessage;
		var frmGetMessage = document.getElementById("frmGetMessage");

		sMessage = new String(frmGetMessage.txtMessage.value);
		if (sMessage.length > 0) {
			window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sMessage);
			frmGetMessage.txtMessage.value = "";
		}
	}

</script>

<script type="text/javascript">
	function pollmessage_refreshMessage() {
		//frmSetMessage.submit();  
		//var frmSetMessage = OpenHR.getForm("pollmessageframe", "frmSetMessage");
		//OpenHR.submitForm(frmSetMessage);
	}
</script>

<div bgcolor='<%=session("ConvertedDesktopColour")%>'>
	<form action="pollmessage_submit" method="post" id="frmSetMessage" name="frmSetMessage">
		<input type="hidden" id="txtMessage" name="txtMessage">
	</form>

	<form id="frmGetMessage" name="frmGetMessage">
		<%
			Response.Write("<INPUT type='hidden' id=txtMessage name=txtMessage value=""" & Replace(Session("pollMessage"), """", "&quot;") & """>" & vbCrLf)
		%>
	</form>
</div>

<script type="text/javascript"> pollmessage_window_onload();</script>

