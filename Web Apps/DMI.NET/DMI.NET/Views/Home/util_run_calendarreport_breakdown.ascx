<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl(Of DMI.NET.Models.CalendarEvent)" %>

<script type="text/javascript">
	function closeCalendarEvent() {
		$("#CalendarEvent").dialog("close");
	}
</script>

<%--<fieldset>
	<legend></legend>

	<div class="display-label">Event Name :</div>
	<div class="display-field">
		<%: Html.DisplayFor(Function(m) m.EventName)%>
	</div>

	<div class="display-label">Description :</div>
	<div class="display-field">
		<%: Html.DisplayFor(Function(m) m.Description)%>
	</div>

	<div class="display-label">Start Date :</div>
	<div class="display-field">
		<%: Html.Label("StartDate", Model.StartDate.ToLongDateString())%>
		<%: Html.DisplayFor(Function(m) m.StartSession)%>
	</div>

	<div class="display-label">End Date :</div>
	<div class="display-field">
		<%: Html.Label("EndDate", Model.EndDate.ToLongDateString())%>
		<%: Html.DisplayFor(Function(m) m.EndSession)%>
	</div>

	<div class="display-label">Duration :</div>
	<div class="display-field">
		<%: Html.DisplayFor(Function(m) m.Duration)%>
	</div>

	<div class="display-label">Reason :</div>
	<div class="display-field">
		<%: Html.DisplayFor(Function(m) m.Reason)%>
	</div>

	<div class="display-label">Calendar Code :</div>
	<div class="display-field">
		<%: Html.DisplayFor(Function(m) m.CalendarCode)%>
	</div>

	<div class="display-label">Region :</div>
	<div class="display-field">
		<%: Html.DisplayFor(Function(m) m.Region)%>
	</div>

	<div class="display-label">Working Pattern :</div>
	<div class="display-field">
		<%: Html.DisplayFor(Function(m) m.WorkingPattern)%>
	</div>
	
</fieldset>--%>

<table>
	<tr>
		<td style="width:150px" >Event Name</td>
		<td class="border"><%: Html.DisplayFor(Function(m) m.EventName)%></td>
	</tr>
	<tr>
		<td>Description</td>
		<td class="border"><%: Html.DisplayFor(Function(m) m.Description)%></td>
	</tr>

	<tr>
		<td>Start Date</td>
		<td class="border"><%: Html.Label("StartDate", Model.StartDate.ToLongDateString())%> <%: Html.DisplayFor(Function(m) m.StartSession)%></td>
	</tr>
	<tr>
		<td>End Date</td>
		<td class="border"><%: Html.Label("EndDate", Model.EndDate.ToLongDateString())%><%: Html.DisplayFor(Function(m) m.EndSession)%></td>
	</tr>
	<tr>
		<td>Duration</td>
		<td class="border"><%: Html.DisplayFor(Function(m) m.Duration)%></td>
	</tr>
	<tr>
		<td>Reason</td>
		<td class="border"><%: Html.DisplayFor(Function(m) m.Reason)%></td>
	</tr>
	<tr>
		<td>Calendar Code</td>
		<td class="border"><%: Html.DisplayFor(Function(m) m.CalendarCode)%></td>
	</tr>
	<tr>
		<td>Region</td>
		<td class="border"><%: Html.DisplayFor(Function(m) m.Region)%></td>
	</tr>
	<tr>
		<td>Working Pattern</td>
		<td class="border"><%: Html.DisplayFor(Function(m) m.WorkingPattern)%></td>
	</tr>

</table>

<div style="text-align: center;padding-top: 20px">
<input id="cmdOK" type="button" value="OK" class="btn" name="cmdOK" onclick="closeCalendarEvent();" />
	</div>

