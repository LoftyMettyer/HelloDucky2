<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<%
    Dim objCalendar As Object
	
    objCalendar = Session("objCalendar" & Session("CalRepUtilID"))

	if Request.Form("txtChangeOptions") <> "" then
		objCalendar.IncludeBankHolidays = Request.Form("txtIncludeBankHolidays")
		objCalendar.IncludeWorkingDaysOnly = Request.Form("txtIncludeWorkingDaysOnly")
		objCalendar.ShowBankHolidays = Request.Form("txtShowBankHolidays")
		objCalendar.ShowCaptions = Request.Form("txtShowCaptions")
		objCalendar.ShowWeekends = Request.Form("txtShowWeekends")
	end if
%>

<form name="frmOptions" id="frmOptions" action="util_run_calendarreport_navfiller.asp?CalRepUtilID=<%=Session("CalRepUtilID").ToString()%>" style="visibility: hidden; display: none" method="post">
    <input type="hidden" name="txtIncludeBankHolidays" id="txtIncludeBankHolidays">
    <input type="hidden" name="txtIncludeWorkingDaysOnly" id="txtIncludeWorkingDaysOnly">
    <input type="hidden" name="txtShowBankHolidays" id="txtShowBankHolidays">
    <input type="hidden" name="txtShowCaptions" id="txtShowCaptions">
    <input type="hidden" name="txtShowWeekends" id="txtShowWeekends">
    <input type="hidden" name="txtChangeOptions" id="txtChangeOptions" value="1">
    <input type="hidden" id="txtCalRep_UtilID" name="txtCalRep_UtilID" value='<%=Session("CalRepUtilID").ToString()%>'>
</form>

<%	
    objCalendar = Nothing
%>