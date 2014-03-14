<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl(Of DMI.NET.Models.CalendarEvent)" %>

<script type="text/javascript">
	
	$(document).ready(function () {
		$("#CalendarEvent").dialog('option', 'title', "Calendar Breakdown");
	});
		
		
	function closeCalendarEvent() {
		$("#CalendarEvent").dialog("close");
	}
</script>

<table>
	<tr>
		<td style="width:150px" >Event Name :</td>
		<td><%: Html.DisplayFor(Function(m) m.EventName)%></td>
	</tr>
	<tr>
		<td>Description :</td>
		<td><%: Html.DisplayFor(Function(m) m.Description)%></td>
	</tr>

	<tr>
		<td>Start Date :</td>
		<td><%: Html.Label("StartDate", Model.StartDate.ToLongDateString())%> <%: Html.DisplayFor(Function(m) m.StartSession)%></td>
	</tr>
	<tr>
		<td>End Date :</td>
		<td><%: Html.Label("EndDate", Model.EndDate.ToLongDateString())%> <%: Html.DisplayFor(Function(m) m.EndSession)%></td>
	</tr>

	<tr>
		<td>Duration :</td>
		<td><%: Html.DisplayFor(Function(m) m.Duration)%></td>
	</tr>
	<tr>
		<td>Reason :</td>
		<td><%: Html.DisplayFor(Function(m) m.Reason)%></td>

	</tr>
	<tr>
		<td>Calendar Code :</td>
		<td><%: Html.DisplayFor(Function(m) m.CalendarCode)%></td>
	</tr>
	<tr>
		<td>Region :</td>
		<td><%: Html.DisplayFor(Function(m) m.Region)%></td>
	</tr>
	<tr>
		<td>Working Pattern :</td>
		<td><%: Html.DisplayFor(Function(m) m.WorkingPattern)%></td>
	</tr>

</table>

<div style="text-align: center;padding-top: 20px">
<input id="cmdCloseEvent" type="button" value="OK" class="btn" name="cmdCloseEvent" onclick="closeCalendarEvent();" />
	</div>

