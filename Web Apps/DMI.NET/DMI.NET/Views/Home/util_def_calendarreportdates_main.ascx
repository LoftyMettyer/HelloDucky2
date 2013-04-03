<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<frameset name="calendarframeset" rows="0, *" frameborder="0" framespacing="0">
	<frame src="util_def_calendarreportdates_data" name="calendardataframe"> 
	<frame src="util_def_calendarreportdates" name="calendarworkframe"> 
</frameset>

<noframes>
<p>This web site depends on frames and it appears your browser does not support them.
</noframes>