<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl(Of DMI.NET.Models.CalendarEvent)" %>

<div class="pageTitleDiv" style="margin-bottom: 15px">
	<span class="pageTitle">Calendar Breakdown</span>
</div>

<div class="clearboth">
	<%: Html.LabelFor(Function(m) m.EventName)%>
	<%: Html.TextBoxFor(Function(m) m.EventName, New With {.disabled = True, .class = "inputProperty"})%>
</div>
<div class="clearboth">
	<%: Html.LabelFor(Function(m) m.Description)%>
	<%: Html.TextBoxFor(Function(m) m.Description, New With {.disabled = True, .class = "inputProperty"})%>
</div>
<div class="clearboth">
	<%: Html.LabelFor(Function(m) m.StartDate)%>
	<%: Html.TextBox("EndDate", Model.StartDate.ToLongDateString() + " " + Model.StartSession.ToString(), New With {.disabled = True, .class = "inputProperty"})%>
</div>
<div class="clearboth">
	<%: Html.LabelFor(Function(m) m.EndDate)%>
	<%: Html.TextBox("EndDate", Model.EndDate.ToLongDateString() + " " + Model.EndSession.ToString(), New With {.disabled = True, .class = "inputProperty"})%>
</div>
<div class="clearboth">
	<%: Html.LabelFor(Function(m) m.Duration)%>
	<%: Html.TextBoxFor(Function(m) m.Duration, New With {.disabled = True, .class = "inputProperty"})%>
</div>
<div class="clearboth">
	<%: Html.DisplayFor(Function(m) m.Description1Column)%>
	<%: Html.TextBoxFor(Function(m) m.Description1, New With {.disabled = True, .class = "inputProperty"})%>
</div>
<div class="clearboth">
	<%: Html.DisplayFor(Function(m) m.Description2Column)%>
	<%: Html.TextBoxFor(Function(m) m.Description2, New With {.disabled = True, .class = "inputProperty"})%>
</div>
<div class="clearboth">
	<%: Html.LabelFor(Function(m) m.Region)%>
	<%: Html.TextBoxFor(Function(m) m.Region, New With {.disabled = True, .class = "inputProperty"})%>
</div>
<div class="clearboth">
	<%: Html.LabelFor(Function(m) m.WorkingPattern) %>
	<%: Html.TextBoxFor(Function(m) m.WorkingPattern, New With {.disabled = True, .class = "inputProperty"})%>
</div>

<div id="divCalendarBreakDownButtons" class="clearboth">
	<input type="button" value="OK" class="button" name="cmdCloseEvent" onclick="closeCalendarEvent();" />
</div>

<script type="text/javascript">
	$(document).ready(function () {
		$("#CalendarEvent").dialog("option", "height", "auto");
		$("#CalendarEvent").dialog("option", "width", "500px");
		$('.inputProperty').css("float", "right");
		$('.inputProperty').css("width", "60%");
		$('.inputProperty').css("margin-bottom", "5px");
	});

	function closeCalendarEvent() {
		$("#CalendarEvent").html("");
		$("#CalendarEvent").dialog("close");
	}
</script>
