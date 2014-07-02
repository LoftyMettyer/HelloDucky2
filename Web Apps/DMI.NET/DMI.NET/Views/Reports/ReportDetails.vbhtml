@Imports DMI.NET
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of Models.CalendarReportModel)

@Code
	Layout = Nothing
End Code

<div class="left">
	Start Date:
	<br />
	<input type="radio" name="StartType" value="0">Fixed
	<input type="datetime" name="StartDate">
	<br />
	<input type="radio" name="StartType" value="1">Current Date
	<br />
	<input type="radio" name="StartType" value="2">Offset
	<input type="text" name="StartOffset" />
	<input type="text" name="StartOffsetPeriod" />
	<br />
	<input type="radio" name="StartType" value="3">Custom
	<input type="hidden" name="StartCustomId" />
	<input type="text" id="txtCustomStart" />
	<input type="button" id="cmdCustomStart"  value="..." onclick="selectCalc('startDate', true)" />

	<br />
	
</div>

<div class="right">
	End Date:
	<br />
	<input type="radio" name="EndType" value="0">Fixed
	<input type="datetime" name="EndDate">
	<br />
	<input type="radio" name="EndType" value="1">Current Date
	<br />
	<input type="radio" name="EndType" value="2">Offset
	<input type="text" name="EndOffset" />
	<input type="text" name="EndOffsetPeriod" />
	<br />
	<input type="radio" name="EndType" value="3">Custom
	<input type="hidden" name="EndCustomId" />
	<input type="text" id="txtCustomEnd" />
	<input type="button" id="cmdCustomEnd" value="..." onclick="selectCalc('endDate', true)" />

	<br />

</div>

<div>

	<input type="checkbox" name="IncludeBankHolidays" />
	@Html.LabelFor(Function(m) m.IncludeBankHolidays)
	<br />
	<input type="checkbox" name="WorkingDaysOnly" />
	@Html.LabelFor(Function(m) m.WorkingDaysOnly)
	<br />
	<input type="checkbox" name="ShowBankHolidays" />
	@Html.LabelFor(Function(m) m.ShowBankHolidays)
	<br />
	<input type="checkbox" name="ShowCalendarOptions" />
	@Html.LabelFor(Function(m) m.ShowCalendarOptions)
	<br />
	<input type="checkbox" name="ShowWeekends" />
	@Html.LabelFor(Function(m) m.ShowWeekends)
	<br />
	<input type="checkbox" name="StartOnCurrentMonth" />
	@Html.LabelFor(Function(m) m.StartOnCurrentMonth)
	<br />

</div>
