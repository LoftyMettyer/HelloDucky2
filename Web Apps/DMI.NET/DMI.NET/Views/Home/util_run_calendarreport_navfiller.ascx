<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<%
	Dim objCalendar As CalendarReport
	objCalendar = CType(Session("objCalendar" & Session("CalRepUtilID")), CalendarReport)

	if Request.Form("txtChangeOptions") <> "" then
		objCalendar.IncludeBankHolidays = CBool(Request.Form("txtIncludeBankHolidays"))
		objCalendar.IncludeWorkingDaysOnly = CBool(Request.Form("txtIncludeWorkingDaysOnly"))
		objCalendar.ShowBankHolidays = CBool(Request.Form("txtShowBankHolidays"))
		objCalendar.ShowCaptions = CBool(Request.Form("txtShowCaptions"))
		objCalendar.ShowWeekends = CBool(Request.Form("txtShowWeekends"))
	end if
%>

<form name="frmNavFillerOptions" id="frmNavFillerOptions" action="util_run_calendarreport_navfiller?CalRepUtilID=<%=Session("CalRepUtilID").ToString()%>" style="visibility: hidden; display: none" method="post">
		<input type="hidden" name="txtIncludeBankHolidays" id="txtIncludeBankHolidays">
		<input type="hidden" name="txtIncludeWorkingDaysOnly" id="txtIncludeWorkingDaysOnly">
		<input type="hidden" name="txtShowBankHolidays" id="txtShowBankHolidays">
		<input type="hidden" name="txtShowCaptions" id="txtShowCaptions">
		<input type="hidden" name="txtShowWeekends" id="txtShowWeekends">
		<input type="hidden" name="txtChangeOptions" id="txtChangeOptions" value="1">
		<input type="hidden" id="txtCalRep_UtilID" name="txtCalRep_UtilID" value='<%=Session("CalRepUtilID").ToString()%>'>
</form>
