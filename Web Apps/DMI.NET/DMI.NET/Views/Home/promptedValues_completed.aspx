<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

<!DOCTYPE html>

<html>
<head runat="server">
		<title></title>
	

	<script type="text/javascript">
		function promptedValues_completed_onload() {
			
			if (OpenHR.parentExists()) {

				try {
					window.parent.window.opener.window.makeSelection('FILTER', '<%=Session("filterIDvalue")%>', '<%=Session("promptsvalue")%>');
				}
				catch (e) {
					try {
						window.parent.opener.window.makeSelection('FILTER', '<%=Session("filterIDvalue")%>', '<%=Session("promptsvalue")%>');
					}
					catch (e) {
					}
				}
				window.parent.close();
			}
			else {
				//jquery div option				
				if ($('#tmpDialog').dialog('isOpen') == true) {
					//prompted Values for OpenHR.modalExpressionSelect screen.
					makeSelection('FILTER', '<%:Session("filterIDvalue")%>', '<%=Session("promptsvalue")%>');
					//$('#tmpDialog').dialog('close').dialog('destroy');
					$('#tmpDialog').remove();

				} else {
					picklistdef_makeSelection('FILTER', '<%=Session("filterIDvalue")%>', '<%=Session("promptsvalue")%>');
				}
			}
		}
		


	</script>

</head>
<body>
		<input type="hidden" id="txtDummyForJQuery" name="txtDummyForJQuery" value="0">
</body>
	
<script type="text/javascript"> promptedValues_completed_onload();</script>
</html>
