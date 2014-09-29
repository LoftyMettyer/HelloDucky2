@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports System.Linq.Expressions
@Inherits System.Web.Mvc.WebViewPage(Of Models.CalendarReportModel)

<fieldset class="width100 floatleft">
	<legend class="fontsmalltitle">Start Date :</legend>
	<fieldset>
		@Html.HiddenFor(Function(m) m.StartCustomViewAccess)
		@Html.HiddenFor(Function(m) m.EndCustomViewAccess)

		<div class="width100 " style="">
			@Html.RadioButton("StartType", CalendarDataType.CurrentDate, Model.StartType = CalendarDataType.CurrentDate, New With {.onclick = "changeEventStartType('CurrentDate')"})
			<span>Today</span>
		</div>

		<div class="width100 " style="">
			<div class="width20 floatleft">
				@Html.RadioButton("StartType", CalendarDataType.Fixed, Model.StartType = CalendarDataType.Fixed, New With {.onclick = "changeEventStartType('Fixed')"})
				<span>Fixed</span>
			</div>
			<div class="formField">
				@Html.TextBoxFor(Function(m) m.StartFixedDate, "{0:dd/MM/yyyy}", New With {.class = "datepicker"})
			</div>
		</div>

		<div class="width100">
			<div class="width20 floatleft">
				@Html.RadioButton("StartType", CalendarDataType.Offset, Model.StartType = CalendarDataType.Offset, New With {.onclick = "changeEventStartType('Offset')"})
				<span>Offset</span>
			</div>
			@Html.TextBoxFor(Function(m) m.StartOffset, New With {.id = "StartOffset", .class = "spinner"})
			@Html.EnumDropDownListFor(Function(m) m.StartOffsetPeriod, New With {.id = "StartOffsetPeriod"})
		</div>

		<div class="width100 ">
			<div class="width20 floatleft">
				@Html.RadioButton("StartType", CalendarDataType.Custom, Model.StartType = CalendarDataType.Custom, New With {.onclick = "changeEventStartType('Custom')"})
				<span>Custom</span>
			</div>

			@Html.HiddenFor(Function(m) m.StartCustomId, New With {.id = "StartCustomId"})
			<div class="formField">
				<input class="floatleft" type="text" id="txtCustomStart" value="@Model.StartCustomName" disabled />
				<input class="floatleft" type="button" id="cmdCustomStart" value="..." onclick="selectCustomStartDate()" />
			</div>
		</div>
	</fieldset>
</fieldset>

<fieldset class="width100 floatleft">
	<legend class="fontsmalltitle">End Dates :</legend>
	<fieldset>
		<div class="width100 " style="">
			@Html.RadioButton("EndType", CalendarDataType.CurrentDate, Model.EndType = CalendarDataType.CurrentDate, New With {.onclick = "changeEventEndType('CurrentDate')"})
			<span>Today</span>
		</div>

		<div class="width100 " style="">
			<div class="width20 floatleft">
				@Html.RadioButton("EndType", CalendarDataType.Fixed, Model.EndType = CalendarDataType.Fixed, New With {.onclick = "changeEventEndType('Fixed')"})
				<span>Fixed</span>
			</div>
			<div class="formField">
				@Html.TextBoxFor(Function(m) m.EndFixedDate, "{0:dd/MM/yyyy}", New With {.class = "datepicker"})
			</div>
		</div>

		<div class="width100">
			<div class="width20 floatleft">
				@Html.RadioButton("EndType", CalendarDataType.Offset, Model.EndType = CalendarDataType.Offset, New With {.onclick = "changeEventEndType('Offset')"})
				<span>Offset</span>
			</div>
			@Html.TextBoxFor(Function(m) m.EndOffset, New With {.id = "EndOffset", .class = "spinner"})
			@Html.EnumDropDownListFor(Function(m) m.EndOffsetPeriod, New With {.id = "EndOffsetPeriod"})
		</div>

		<div class="width100 ">
			<div class="width20 floatleft">
				@Html.RadioButton("EndType", CalendarDataType.Custom, Model.EndType = CalendarDataType.Custom, New With {.onclick = "changeEventEndType('Custom')"})
				<span>Custom</span>
			</div>

			@Html.HiddenFor(Function(m) m.EndCustomId, New With {.id = "EndCustomId"})
			<div class="formField">
				<input class="floatleft" type="text" id="txtCustomEnd" value="@Model.EndCustomName" disabled />
				<input class="floatleft" type="button" id="cmdCustomEnd" value="..." onclick="selectCustomEndDate()" />
			</div>
		</div>
	</fieldset>
</fieldset>

<fieldset class="width100 floatleft">
	<legend class="fontsmalltitle">Default Display Options :</legend>
	<fieldset class="floatleft width25">
		<div class="padbot5">
			@Html.CheckBoxFor(Function(m) m.IncludeBankHolidays)
			@Html.LabelFor(Function(m) m.IncludeBankHolidays)
		</div>
		<div class="padbot5">
			@Html.CheckBoxFor(Function(m) m.WorkingDaysOnly)
			@Html.LabelFor(Function(m) m.WorkingDaysOnly)
		</div>
		<div class="padbot5">
			@Html.CheckBoxFor(Function(m) m.StartOnCurrentMonth)
			@Html.LabelFor(Function(m) m.StartOnCurrentMonth)
		</div>
		<div class="padbot5">
			@Html.CheckBoxFor(Function(m) m.ShowBankHolidays)
			@Html.LabelFor(Function(m) m.ShowBankHolidays)
		</div>
		<div class="padbot5">
			@Html.CheckBoxFor(Function(m) m.ShowCaptions)
			@Html.LabelFor(Function(m) m.ShowCaptions)
		</div>
		<div class="padbot5">
			@Html.CheckBoxFor(Function(m) m.ShowWeekends)
			@Html.LabelFor(Function(m) m.ShowWeekends)
		</div>
	</fieldset>
</fieldset>

<script>
	$(function () {

		$(".spinner").spinner({
			min: -99,
			max: 99,
			showOn: 'both'
		}).css("width", "20px");

		$(".datepicker").datepicker();
		changeEventStartType('@Model.StartType');
		changeEventEndType('@Model.EndType');
	});


	function changeEventStartType(type) {

		$("#StartFixedDate").attr("disabled", "true");
		$("#StartOffset").spinner("option", "disabled", true);
		$("#StartOffsetPeriod").attr("disabled", "true");
		button_disable($("#cmdCustomStart")[0], (type != "Custom"));

		switch (type) {
			case "Fixed":
				$("#StartFixedDate").removeAttr("disabled");
				$("#StartCustomId").val(0);
				$("#StartOffset").val(0);
				$("#StartOffsetPeriod").val(0);
				$("#StartCustomId").val(0);
				$("#txtCustomStart").val("");
				break;

			case "Current":
				$("#StartFixedDate").val('');
				$("#StartCustomId").val(0);
				$("#StartOffset").val(0);
				$("#StartOffsetPeriod").val(0);
				$("#StartCustomId").val(0);
				$("#txtCustomStart").val("");
				break;

			case "Offset":
				$("#StartFixedDate").val('');
				$("#StartOffset").spinner("option", "disabled", false);
				$("#StartOffsetPeriod").removeAttr("disabled");
				$("#StartCustomId").val(0);
				$("#txtCustomStart").val("");
				break;

			default:
				$("#StartFixedDate").val('');
				$("#StartOffset").val(0);
				$("#StartOffsetPeriod").val(0);
				break;

		}

		setViewAccess('CALC', $("#StartCustomViewAccess"), 'RW', '');

	}

	function changeEventEndType(type) {

		$("#EndFixedDate").attr("disabled", "true");
		$("#EndOffset").spinner("option", "disabled", true);
		$("#EndOffsetPeriod").attr("disabled", "true");
		button_disable($("#cmdCustomEnd")[0], (type != "Custom"));

		switch (type) {
			case "Fixed":
				$("#EndFixedDate").removeAttr("disabled");
				$("#EndCustomId").val(0);
				$("#EndOffset").val(0);
				$("#EndOffsetPeriod").val(0);
				$("#EndCustomId").val(0);
				$("#txtCustomEnd").val("");
				break;

			case "Current":
				$("#EndFixedDate").val('');
				$("#EndCustomId").val(0);
				$("#EndOffset").val(0);
				$("#EndOffsetPeriod").val(0);
				$("#EndCustomId").val(0);
				$("#txtCustomEnd").val("");
				break;

			case "Offset":
				$("#EndFixedDate").val('');
				$("#EndOffset").spinner("option", "disabled", false);
				$("#EndOffsetPeriod").removeAttr("disabled");
				$("#EndCustomId").val(0);
				$("#txtCustomEnd").val("");
				break;

			default:
				$("#EndFixedDate").val('');
				$("#EndOffset").val(0);
				$("#EndOffsetPeriod").val(0);
				break;

		}

		setViewAccess('CALC', $("#EndCustomViewAccess"), 'RW', '');

	}

	function selectCustomStartDate() {

		var tableID = $("#BaseTableID option:selected").val();
		var currentID = $("#StartCustomId").val();

		OpenHR.modalExpressionSelect("CALC", 0, currentID, function (id, name, access) {
			if (access == "HD" && $("#Owner").val() != '@Session("Username")') {
				$("#StartCustomId").val(0);
				$("#txtCustomStart").val('None');
				OpenHR.modalMessage("The report start date calculation will be removed from this definition as it is hidden and you do not have permission to make this definition hidden.");
			}
			else {
				$("#StartCustomId").val(id);
				$("#txtCustomStart").val(name);
				setViewAccess('CALC', $("#StartCustomViewAccess"), access, "report start date");
			}
		}, 400, 400);
	}

	function selectCustomEndDate() {

		var tableID = $("#BaseTableID option:selected").val();
		var currentID = $("#EndCustomId").val();

		OpenHR.modalExpressionSelect("CALC", 0, currentID, function (id, name, access) {
			if (access == "HD" && $("#Owner").val() != '@Session("Username")') {
				$("#EndCustomId").val(0);
				$("#txtCustomEnd").val('None');
				OpenHR.modalMessage("The report end date calculation will be removed from this definition as it is hidden and you do not have permission to make this definition hidden.");
			}
			else {
				$("#EndCustomId").val(id);
				$("#txtCustomEnd").val(name);
				setViewAccess('CALC', $("#EndCustomViewAccess"), access, "report end date");
			}
		}, 400, 400);

	}

</script>
