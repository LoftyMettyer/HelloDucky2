<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

<!DOCTYPE html>

<html>
<head runat="server">
		<title></title>
	

	<script type="text/javascript">
		function promptedValues_completed_onload() {

			if (OpenHR.parentExists()) {

				try {
					window.parent.window.dialogArguments.window.makeSelection('FILTER', '<%=Session("filterIDvalue")%>', '<%=Session("promptsvalue")%>');
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
				try {
					picklistdef_makeSelection('FILTER', '<%=Session("filterIDvalue")%>', '<%=Session("promptsvalue")%>');
				}
				catch (e) {				
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
