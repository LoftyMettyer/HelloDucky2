<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

<%@ Import Namespace="DMI.NET" %>

<html>
<head>
	<script src="<%:Url.Content("~/Include/ctl_SetFont.txt")%>" type="text/javascript"> </script>
	<link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css" />
	<script src="<%: Url.Content("~/bundles/jQuery")%>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/bundles/jQueryUI7")%>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/bundles/OpenHR_General")%>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/bundles/utilities_calendarreports")%>" type="text/javascript"></script>
	<link href="<%: Url.LatestContent("~/Content/Site.css")%>" rel="stylesheet" type="text/css" />
	<title>OpenHR Intranet</title>
</head>
<body>
	<div id="calendarframeset" name="calendarframeset">
		<div data-framesource="util_def_calendarreportdates_data" name="calendardataframe" id="calendardataframe" style="display: none;">
			<%	Html.RenderPartial("~/views/home/util_def_calendarreportdates_data.ascx")%>
		</div>
		<div data-framesource="util_def_calendarreportdates" name="calendarworkframe" id="calendarworkframe">
			<%	Html.RenderPartial("~/views/home/util_def_calendarreportdates.ascx")%>
		</div>
	</div>
</body>
</html>
