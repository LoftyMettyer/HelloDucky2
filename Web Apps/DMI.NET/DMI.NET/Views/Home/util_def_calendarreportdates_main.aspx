<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

<html>
	
	<head>
		<script src="<%:Url.Content("~/Scripts/FormScripts/calendarreportdef.js?v=1")%>" type="text/javascript"> </script>
		<script src="<%:Url.Content("~/Scripts/FormScripts/calendarreportdef.js")%>" type="text/javascript"> </script>
		<script src="<%:Url.Content("~/Scripts/ctl_SetFont.js")%>" type="text/javascript"> </script>
		<script src="<%:Url.Content("~/Scripts/ctl_SetStyles.js")%>" type="text/javascript"> </script>
		<link href="<%:Url.Content("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />
		<script src="<%:Url.Content("~/Scripts/jquery/jquery-1.8.3.js")%>" type="text/javascript"> </script>
		<script src="<%:Url.Content("~/Scripts/openhr.js")%>" type="text/javascript"> </script>
		
		<title>OpenHR Intranet</title>

	</head>
	<body>
		<div id="calendarframeset" name="calendarframeset">
			<div data-framesource="util_def_calendarreportdates_data" name="calendardataframe" id="calendardataframe">
				<% Html.RenderPartial("~/views/home/util_def_calendarreportdates_data.ascx")%>
			</div>
			<div data-framesource="util_def_calendarreportdates" name="calendarworkframe" id="calendarworkframe">
				<% Html.RenderPartial("~/views/home/util_def_calendarreportdates.ascx")%>
			</div>
		</div>

	</body>
</html>