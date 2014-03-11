<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

<%@ Import Namespace="DMI.NET" %>

<!DOCTYPE html>

<html>
<head>
	<script src="<%:Url.LatestContent("~/Include/ctl_SetFont.txt")%>" type="text/javascript"> </script>
	<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />
	<script src="<%: Url.LatestContent("~/bundles/jQuery")%>" type="text/javascript"></script>
	<script src="<%: Url.LatestContent("~/bundles/jQueryUI7")%>" type="text/javascript"></script>
	<script src="<%: Url.LatestContent("~/bundles/OpenHR_General")%>" type="text/javascript"></script>
	<script src="<%: Url.LatestContent("~/bundles/utilities_calendarreports")%>" type="text/javascript"></script>
	<script src="<%: Url.LatestContent("~/bundles/OpenHR_ActiveX")%>" type="text/javascript"></script>
	<link href="<%: Url.LatestContent("~/Content/Site.css")%>" rel="stylesheet" type="text/css" />
	
	<script id="officebarscript" src="<%: Url.LatestContent("~/Scripts/officebar/jquery.officebar.js")%>" type="text/javascript"></script>
	<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />
	<link id="DMIthemeLink" href="<%: Url.LatestContent("~/Content/themes/" & Session("ui-admin-theme").ToString() & "/jquery-ui.min.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/general_enclosed_foundicons.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/font-awesome.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/fonts/SSI80v194934/style.css")%>" rel="stylesheet" />
	

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
