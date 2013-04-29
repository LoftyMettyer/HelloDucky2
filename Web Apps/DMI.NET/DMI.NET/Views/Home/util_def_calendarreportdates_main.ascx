<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>


<!DOCTYPE html>
<html>
<head>
	<link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css" />
	<title>OpenHR Intranet</title>
</head>

<div id="calendardataframe" data-framesource="util_def_calendarreportdates_data" style="display: none;">
	<%Html.RenderPartial("~/views/home/util_def_calendarreportdates_data.aspx")%>
</div>

<div id="calendarworkframe" data-framesource="util_def_calendarreportdates" style="display: none;">
	<%Html.RenderPartial("~/views/home/util_def_calendarreportdates.ascx")%>
</div>

<%--<frameset name="calendarframeset" rows="0, *" frameborder="0" framespacing="0">
	<frame src="util_def_calendarreportdates_data" name="calendardataframe"> 
	<frame src="util_def_calendarreportdates" name="calendarworkframe"> 
</frameset>--%>
</html>
