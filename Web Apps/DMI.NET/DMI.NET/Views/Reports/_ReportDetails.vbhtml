@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports HR.Intranet.Server.Enums
@Inherits System.Web.Mvc.WebViewPage(Of Models.CalendarReportModel)

<div class="left">
	Start Date:
	<br />
	@Html.RadioButton("StartType", CalendarDataType.Fixed, Model.StartType = CalendarDataType.Fixed)
	Fixed
	@Html.TextBox("StartFixedDate", Model.StartFixedDate, New With {.type = "datetime", .class = "datetimepicker"})
	<br />
	@Html.RadioButton("StartType", CalendarDataType.CurrentDate, Model.StartType = CalendarDataType.CurrentDate)
	Current Date
	<br />
	@Html.RadioButton("StartType", CalendarDataType.Offset, Model.StartType = CalendarDataType.Offset)
	Offset
	@Html.TextBox("StartOffset", Model.StartOffset)
	@Html.TextBox("StartOffsetPeriod", Model.StartOffsetPeriod)
	<br />
	@Html.RadioButton("StartType", CalendarDataType.Custom, Model.StartType = CalendarDataType.Custom)
	Custom
	@Html.HiddenFor(Function(m) m.StartCustomId)
	<input type="text" id="txtCustomStart" value="@Model.StartCustomName" disabled />
	<input type="button" id="cmdCustomStart"  value="..." onclick="selectCalc('startDate', true)" />

	<br />
	
</div>

<div class="right">
	End Date:
	<br />
	@Html.RadioButton("EndType", CalendarDataType.Fixed, Model.EndType = CalendarDataType.Fixed)
	Fixed
	@Html.TextBox("EndFixedDate", Model.EndFixedDate, New With {.type = "datetime", .class = "datetimepicker"})
	<br />
	@Html.RadioButton("EndType", CalendarDataType.CurrentDate, Model.EndType = CalendarDataType.CurrentDate)
	Current Date
	<br />
	@Html.RadioButton("EndType", CalendarDataType.Offset, Model.EndType = CalendarDataType.Offset)
	Offset
	@Html.TextBox("EndOffset", Model.EndOffset)
	@Html.TextBox("EndOffsetPeriod", Model.EndOffsetPeriod)
	<br />
	@Html.RadioButton("EndType", CalendarDataType.Custom, Model.EndType = CalendarDataType.Custom)
	Custom
	@Html.HiddenFor(Function(m) m.EndCustomId)
	<input type="text" id="txtCustomEnd" value="@Model.EndCustomName" disabled />
	<input type="button" id="cmdCustomEnd" value="..." onclick="selectCalc('endDate', true)" />

	<br />

</div>

<div>
	Default Display Options:
	<br/>

	@Html.CheckBoxFor(Function(m) m.IncludeBankHolidays)
	@Html.LabelFor(Function(m) m.IncludeBankHolidays)
	<br />
	@Html.CheckBoxFor(Function(m) m.WorkingDaysOnly)
	@Html.LabelFor(Function(m) m.WorkingDaysOnly)
	<br />
	@Html.CheckBoxFor(Function(m) m.ShowBankHolidays)
	@Html.LabelFor(Function(m) m.ShowBankHolidays)
	<br />
	@Html.CheckBoxFor(Function(m) m.ShowCaptions)
	@Html.LabelFor(Function(m) m.ShowCaptions)
	<br />
	@Html.CheckBoxFor(Function(m) m.ShowWeekends)
	@Html.LabelFor(Function(m) m.ShowWeekends)
	<br />
	@Html.CheckBoxFor(Function(m) m.StartOnCurrentMonth)
	@Html.LabelFor(Function(m) m.StartOnCurrentMonth)
	<br />

</div>
