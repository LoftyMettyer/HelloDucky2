@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports HR.Intranet.Server.Enums
@Inherits System.Web.Mvc.WebViewPage(Of Models.CalendarReportModel)

<div class="left">
	Start Date:
	<br />
	@Html.RadioButton("StartType", CalendarDataType.Fixed, Model.StartType = CalendarDataType.Fixed, New With {.onclick = "changeEventStartType('fixed')"})
	Fixed
	@Html.TextBox("StartFixedDate", Model.StartFixedDate, New With {.type = "datetime", .class = "datetimepicker", .readonly = Not (Model.StartType = CalendarDataType.Fixed)})
	<br />
	@Html.RadioButton("StartType", CalendarDataType.CurrentDate, Model.StartType = CalendarDataType.CurrentDate, New With {.onclick = "changeEventStartType('current')"})
	Current Date
	<br />
	@Html.RadioButton("StartType", CalendarDataType.Offset, Model.StartType = CalendarDataType.Offset, New With {.onclick = "changeEventStartType('offset')"})
	Offset
	@Html.TextBox("StartOffset", Model.StartOffset, New With {.readonly = Not (Model.StartType = CalendarDataType.Offset)})
	@Html.TextBox("StartOffsetPeriod", Model.StartOffsetPeriod, New With {.readonly = Not (Model.StartType = CalendarDataType.Offset)})
		
	<br />
	@Html.RadioButton("StartType", CalendarDataType.Custom, Model.StartType = CalendarDataType.Custom, New With {.onclick = "changeEventStartType('custom')"})
	Custom
	@Html.HiddenFor(Function(m) m.StartCustomId)
	<input type="text" id="txtCustomStart" value="@Model.StartCustomName" disabled />
	<input type="button" id="cmdCustomStart" value="..." onclick="selectCustomStartDate()" />
	<br />
	
</div>

<div class="right">
	End Date:
	<br />
	@Html.RadioButton("EndType", CalendarDataType.Fixed, Model.EndType = CalendarDataType.Fixed, New With {.onclick = "changeEventEndType('fixed')"})
	Fixed
	@Html.TextBox("EndFixedDate", Model.EndFixedDate, New With {.type = "datetime", .class = "datetimepicker", .readonly = Not (Model.EndType = CalendarDataType.Fixed)})
	<br />
	@Html.RadioButton("EndType", CalendarDataType.CurrentDate, Model.EndType = CalendarDataType.CurrentDate, New With {.onclick = "changeEventEndType('current')"})
	Current Date
	<br />
	@Html.RadioButton("EndType", CalendarDataType.Offset, Model.EndType = CalendarDataType.Offset, New With {.onclick = "changeEventEndType('offset')"})
	Offset
	@Html.TextBox("EndOffset", Model.EndOffset, New With {.readonly = Not (Model.EndType = CalendarDataType.Offset)})
	@Html.TextBox("EndOffsetPeriod", Model.EndOffsetPeriod, New With {.readonly = Not (Model.EndType = CalendarDataType.Offset)})
	<br />
	@Html.RadioButton("EndType", CalendarDataType.Custom, Model.EndType = CalendarDataType.Custom, New With {.onclick = "changeEventEndType('custom')"})
	Custom
	@Html.HiddenFor(Function(m) m.EndCustomId)
	<input type="text" id="txtCustomEnd" value="@Model.EndCustomName" disabled />
	<input type="button" id="cmdCustomEnd" value="..." onclick="selectCustomEndDate()" />

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

<script type="text/javascript">

	function changeEventStartType(type) {

		$("#StartFixedDate").attr("readonly", "true");
		$("#StartOffset").attr("readonly", "true");
		$("#StartOffsetPeriod").attr("readonly", "true");
		$("#cmdCustomStart").attr("disabled", "true");

		switch (type) {
			case "fixed":
				$("#StartFixedDate").removeAttr("readonly");
				$("#StartCustomId").val(0);
				$("#StartOffset").val(0);
				$("#StartOffsetPeriod").val(0);
				break;

			case "current":
				$("#StartFixedDate").val('');
				$("#StartCustomId").val(0);
				$("#StartOffset").val(0);
				$("#StartOffsetPeriod").val(0);
				break;

			case "offset":
				$("#StartFixedDate").val('');
				$("#StartOffset").removeAttr("readonly");
				$("#StartOffsetPeriod").removeAttr("readonly");
				break;

			default:
				$("#StartFixedDate").val('');
				$("#StartOffset").val(0);
				$("#StartOffsetPeriod").val(0);
				$("#cmdCustomStart").removeAttr("disabled", false);
				break;

		}

	}

	function changeEventEndType(type) {

		$("#EndFixedDate").attr("readonly", "true");
		$("#EndOffset").attr("readonly", "true");
		$("#EndOffsetPeriod").attr("readonly", "true");
		$("#cmdCustomEnd").attr("disabled", "true");

		switch (type) {
			case "fixed":
				$("#EndFixedDate").removeAttr("readonly");
				$("#EndCustomId").val(0);
				$("#EndOffset").val(0);
				$("#EndOffsetPeriod").val(0);
				break;

			case "current":
				$("#EndFixedDate").val('');
				$("#EndCustomId").val(0);
				$("#EndOffset").val(0);
				$("#EndOffsetPeriod").val(0);
				break;

			case "offset":
				$("#EndFixedDate").val('');
				$("#EndOffset").removeAttr("readonly");
				$("#EndOffsetPeriod").removeAttr("readonly");
				break;

			default:
				$("#EndFixedDate").val('');
				$("#EndOffset").val(0);
				$("#EndOffsetPeriod").val(0);
				$("#cmdCustomEnd").removeAttr("disabled", false);
				break;

		}

	}

	function selectCustomStartDate() {

		var tableID = $("#BaseTableID option:selected").val();
		var currentID = $("#StartCustomId").val();

		OpenHR.modalExpressionSelect("CALC", tableID, currentID, function (id, name) {
			$("#StartCustomId").val(id);
			$("#txtCustomStart").val(name);
		});

	}

	function selectCustomEndDate() {

		var tableID = $("#BaseTableID option:selected").val();
		var currentID = $("#EndCustomId").val();

		OpenHR.modalExpressionSelect("CALC", tableID, currentID, function (id, name) {
			$("#EndCustomId").val(id);
			$("#txtCustomEnd").val(name);
		});

	}



</script>
