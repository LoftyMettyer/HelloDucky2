<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

<!DOCTYPE html>

<html>
<head runat="server">
		<title></title>
	
	<script src="<%: Url.Content("~/bundles/jQuery")%>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/bundles/jQueryUI7")%>" type="text/javascript"></script>

	<script type="text/javascript">
		function promptedValues_completed_onload() {
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

	</script>

</head>
<body>
		<input type="hidden" id="txtDummyForJQuery" name="txtDummyForJQuery" value="0">
</body>
	
<script type="text/javascript"> promptedValues_completed_onload();</script>
</html>
