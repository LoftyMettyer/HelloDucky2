<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

<!DOCTYPE html>

<html>
<head runat="server">
    <title></title>
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
</body>
	
<script type="text/javascript"> promptedValues_completed_onload();</script>
</html>
