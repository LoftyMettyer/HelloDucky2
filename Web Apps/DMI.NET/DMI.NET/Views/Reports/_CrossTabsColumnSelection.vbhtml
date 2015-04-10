@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Classes
@Inherits System.Web.Mvc.WebViewPage(Of Models.CrossTabModel)

<fieldset id="CrossTabsColumnTab" class="width100">
	<legend class="fontsmalltitle">
		Headings &amp; Column Breaks
	</legend>
	<fieldset>
		<table class="width100">
			<thead class="fontsmalltitle">
				<tr>
					<td style="width:10%"></td>
					<td style="width:55%;text-align:center">Column</td>
					<td style="width:10%;text-align:center">Start</td>
					<td style="width:10%;text-align:center">Stop</td>
					<td style="width:10%;text-align:center">Increment</td>
				</tr>
			</thead>
			<tr>
				<td style="padding-right: 20px;">Horizontal :</td>
				<td>
					@Html.ColumnDropdownFor(Function(m) m.HorizontalID, New ColumnFilter() With {.TableID = Model.BaseTableID, .ExcludeOleAndPhoto = True}, New With {.onchange = "crossTabHorizontalChange()", .class = "enableSaveButtonOnComboChange"})
					@Html.ValidationMessageFor(Function(m) m.HorizontalID)
					@Html.Hidden("HorizontalDataType", CInt(Model.HorizontalDataType))
				</td>
				<td class="startstopincrementcol">@Html.TextBoxFor(Function(m) m.HorizontalStart, New With {.class = "selectFullText"})</td>
				<td class="startstopincrementcol">@Html.TextBoxFor(Function(m) m.HorizontalStop, New With {.class = "selectFullText"})</td>
				<td class="startstopincrementcol">@Html.TextBoxFor(Function(m) m.HorizontalIncrement, New With {.class = "selectFullText"})</td>
			</tr>
			<tr style="height: 10px;"></tr>
			<tr>
				<td>Vertical :</td>
				<td>
					@Html.ColumnDropdownFor(Function(m) m.VerticalID, New ColumnFilter() With {.TableID = Model.BaseTableID, .ExcludeOleAndPhoto = True}, New With {.onchange = "crossTabVerticalChange()", .class = "enableSaveButtonOnComboChange"})
					@Html.ValidationMessageFor(Function(m) m.VerticalID)
					@Html.Hidden("VerticalDataType", CInt(Model.VerticalDataType))
				</td>
				<td class="startstopincrementcol">@Html.TextBoxFor(Function(m) m.VerticalStart, New With {.class = "selectFullText"})</td>
				<td class="startstopincrementcol">@Html.TextBoxFor(Function(m) m.VerticalStop, New With {.class = "selectFullText"})</td>
				<td class="startstopincrementcol">@Html.TextBoxFor(Function(m) m.VerticalIncrement, New With {.class = "selectFullText"})</td>
			</tr>
			<tr style="height: 10px;"></tr>
			<tr>
				<td>Page Break :</td>
				<td>
					@Html.ColumnDropdownFor(Function(m) m.PageBreakID, New ColumnFilter() With {.TableID = Model.BaseTableID, .AddNone = True, .ExcludeOleAndPhoto = True}, New With {.onchange = "crossTabPageBreakChange()", .class = "enableSaveButtonOnComboChange"})
					@Html.Hidden("PageBreakDataType", CInt(Model.PageBreakDataType))
				</td>
				<td class="startstopincrementcol">@Html.TextBoxFor(Function(m) m.PageBreakStart, New With {.class = "selectFullText"})</td>
				<td class="startstopincrementcol">@Html.TextBoxFor(Function(m) m.PageBreakStop, New With {.class = "selectFullText"})</td>
				<td class="startstopincrementcol">@Html.TextBoxFor(Function(m) m.PageBreakIncrement, New With {.class = "selectFullText"})</td>
			</tr>
		
			<tr style="height:60px;">
			<td class="fontsmalltitle" style="position:absolute;margin-left: -13px; margin-top: 20px;">Intersection</td>
			</tr>

			<tr>
				<td>Column :</td>
				<td colspan="1">
					@Html.ColumnDropdownFor(Function(m) m.IntersectionID, New ColumnFilter() With {.TableID = Model.BaseTableID, .AddNone = True, .IsNumeric = True}, New With {.onchange = "crossTabIntersectionType();", .class = "enableSaveButtonOnComboChange"})
				</td>
				<td colspan="3" rowspan="4" style="padding-left:20px;vertical-align:top ">
					@Html.CheckBox("SuppressZeros", Model.SuppressZeros)
					@Html.LabelFor(Function(m) m.SuppressZeros)
					<br />
					@Html.CheckBox("UseThousandSeparators", Model.UseThousandSeparators)
					@Html.LabelFor(Function(m) m.UseThousandSeparators)
					<br />
					@Html.CheckBox("PercentageOfType", Model.PercentageOfType)
					@Html.LabelFor(Function(m) m.PercentageOfType)
					<br />
					@Html.CheckBox("PercentageOfPage", Model.PercentageOfPage)
					@Html.LabelFor(Function(m) m.PercentageOfPage)
				</td>
			</tr>
			<tr style="height: 10px;"></tr>
			<tr>
				<td>@Html.LabelFor(Function(m) m.IntersectionType)</td>
				<td colspan="1">
					@Html.EnumDropDownListFor(Function(m) m.IntersectionType, New With {.class = "enableSaveButtonOnComboChange"})
				</td>
			</tr>
			<tr style="height:60px;"></tr>
			
		</table>
	</fieldset>
</fieldset>
<br />
@Html.ValidationMessageFor(Function(m) m.HorizontalStart)
@Html.ValidationMessageFor(Function(m) m.HorizontalStop)
@Html.ValidationMessageFor(Function(m) m.HorizontalIncrement)
@Html.ValidationMessageFor(Function(m) m.VerticalStart)
@Html.ValidationMessageFor(Function(m) m.VerticalStop)
@Html.ValidationMessageFor(Function(m) m.VerticalIncrement)
@Html.ValidationMessageFor(Function(m) m.PageBreakStart)
@Html.ValidationMessageFor(Function(m) m.PageBreakStop)
@Html.ValidationMessageFor(Function(m) m.PageBreakIncrement)


<script type="text/javascript">

	function refreshCrossTabColumnsAvailable() {

		$.ajax({
			url: 'Reports/GetAvailableColumnsForTable?TableID=' + $("#BaseTableID").val(),
			datatype: 'json',
			mtype: 'GET',
			cache: false,
			success: function (json) {

				var OptionNone = '<option value=0 data-datatype=0 data-decimals=0 selected>None</option>';
				var optionHorizontal = "";
				var optionVertical = "";
				var optionPageBreak = "";
				var optionIntersection = "";

				var options = '';
				for (var i = 0; i < json.length; i++) {

					optionHorizontal += "<option value='" + json[i].ID + "' data-datatype='" + json[i].DataType + "' data-size='" + json[i].Size + "' data-decimals='" + json[i].Decimals + "'>" + json[i].Name + "</option>";
					optionVertical += "<option value='" + json[i].ID + "' data-datatype='" + json[i].DataType + "' data-size='" + json[i].Size + "' data-decimals='" + json[i].Decimals + "'>" + json[i].Name + "</option>";
					optionPageBreak += "<option value='" + json[i].ID + "' data-datatype='" + json[i].DataType + "' data-size='" + json[i].Size + "' data-decimals='" + json[i].Decimals + "'>" + json[i].Name + "</option>";

					if (json[i].IsNumeric) {
						optionIntersection += "<option value='" + json[i].ID + "' data-datatype='" + json[i].DataType + "' data-size='" + json[i].Size + "' data-decimals='" + json[i].Decimals + "'>" + json[i].Name + "</option>";
					}

				}

				$("select#HorizontalID").html(optionHorizontal);
				$("select#VerticalID").html(optionVertical);
				$("select#PageBreakID").html(OptionNone + optionPageBreak);
				$("select#IntersectionID").html(OptionNone + optionIntersection);

				crossTabHorizontalClick();
				crossTabVerticalClick();
				crossTabPageBreakClick();
				crossTabIntersectionType();

			}
		});
	}

	function crossTabIntersectionType() {

		var bReadOnly = isDefinitionReadOnly();
		var dropDown = $("#IntersectionID")[0];
		var iDataType = dropDown.options[dropDown.selectedIndex].attributes["data-datatype"].value;
		combo_disable($("#IntersectionType"), (iDataType == "0") || bReadOnly);
		$("#IntersectionType")[0].selectedIndex = 0;
		refreshTab2Controls();

	}

	function refreshCrossTabColumn(target, type) {

		var bReadOnly = isDefinitionReadOnly();
		var horizontalValue = $("#HorizontalID").val();
		var verticalValue = $("#VerticalID").val();
		var pageBreakValue = $("#PageBreakID").val();

		var iDataType = target.options[target.selectedIndex].attributes["data-datatype"].value;
		var iDecimals = target.options[target.selectedIndex].attributes["data-decimals"].value;
		
		$("#" + type + "DataType").val(iDataType);
		switch (iDataType) {
			case "2":
				$("#" + type + "Start").attr("disabled", bReadOnly);
				$("#" + type + "Stop").attr("disabled", bReadOnly);
				$("#" + type + "Increment").attr("disabled", bReadOnly);
				break;

			case "4":
				$("#" + type + "Start").attr("disabled", bReadOnly);
				$("#" + type + "Stop").attr("disabled", bReadOnly);
				$("#" + type + "Increment").attr("disabled", bReadOnly);
				break;

			default:
				$("#" + type + "Start").attr("disabled", "disabled");
				$("#" + type + "Start").val("");
				$("#" + type + "Stop").attr("disabled", "disabled");
				$("#" + type + "Stop").val("");
				$("#" + type + "Increment").attr("disabled", "disabled");
				$("#" + type + "Increment").val("");
		}

		$("#" + type + "Start").autoNumeric('destroy');
		$("#" + type + "Stop").autoNumeric('destroy');
		$("#" + type + "Increment").autoNumeric('destroy');
        
	    // Use autoNumeric formatting only for interger and float value
		if (iDataType == 2 || iDataType == 4)
		{
		    $("#" + type + "Start").autoNumeric({ aSep: '', aNeg: '', mDec: iDecimals, mRound: 'S', mNum: 10});
		    $("#" + type + "Stop").autoNumeric({ aSep: '', aNeg: '', mDec: iDecimals, mRound: 'S', mNum: 10});
		    $("#" + type + "Increment").autoNumeric({ aSep: '', aNeg: '', mDec: iDecimals, mRound: 'S', mNum: 10});
		}

	}

	function crossTabHorizontalChange() {
		$("#HorizontalStart").val(0);
		$("#HorizontalStop").val(0);
		$("#HorizontalIncrement").val(0);
		crossTabHorizontalClick();
		refreshTab2Controls();
	}

	function crossTabVerticalChange() {
		$("#VerticalStart").val(0);
		$("#VerticalStop").val(0);
		$("#VerticalIncrement").val(0);
		crossTabVerticalClick();
		refreshTab2Controls();
	}

	function crossTabPageBreakChange() {
		$("#PageBreakStart").val(0);
		$("#PageBreakStop").val(0);
		$("#PageBreakIncrement").val(0);
		crossTabPageBreakClick();
		refreshTab2Controls();
	}

	function crossTabHorizontalClick() {

		var horval = $("#HorizontalID").val();

		//reset ver and pb so none are disabled/hidden
		$('#VerticalID option').removeAttr('disabled');
		$('#PageBreakID option').removeAttr('disabled');

		//now hide/disable matching items in ver and pb
		$('#VerticalID option, #PageBreakID option').filter(function () {
			return $(this).val() == horval;
		}).attr('disabled', 'disabled');

		//reset ver if it is selected by hor
		if ($("#VerticalID option:selected").val() == horval) {
			//reset the value to top item
			$('#VerticalID').val($("#VerticalID option:not([disabled]):first").val());
		}

		//reset pb if it is selected by hor
		if ($("#PageBreakID option:selected").val() == horval) {
			//reset the value to top item
			$('#PageBreakID').val($("#PageBreakID option:not([disabled]):first").val());
		}

		refreshCrossTabColumn($("#HorizontalID")[0], 'Horizontal');

	}

	function crossTabVerticalClick() {

		var horval = $("#HorizontalID").val();
		var vertval = $("#VerticalID").val();

		//reset ver and pb so none are disabled/hidden
		$('#PageBreakID option').removeAttr('disabled');

		//now hide/disable matching items in ver and pb
		$('#PageBreakID option').filter(function () {
			return $(this).val() == vertval || $(this).val() == horval;
		}).attr('disabled', 'disabled');

		//reset pb if it is selected by hor or ver
		if ($("#PageBreakID option:selected").val() == vertval || $("#PageBreakID option:selected").val() == horval) {
			//reset the value to top item
			$('#PageBreakID').val($("#PageBreakID option:first").val());
		}

		refreshCrossTabColumn($("#VerticalID")[0], 'Vertical');

	}

	function crossTabPageBreakClick() {
		refreshCrossTabColumn($("#PageBreakID")[0], 'PageBreak');
	}

	$(function () {

		crossTabIntersectionType();

		$("#CrossTabsColumnTab select").css("width", "100%");
		$('table').attr('border', '0');

		crossTabHorizontalClick();
		crossTabVerticalClick();
		crossTabPageBreakClick();

		$('#PercentageOfType').click(function () { refreshTab2Controls(); });

		refreshTab2Controls();

	    //Note:-
		//This solution working in Firefox, Chrome and IE, both with keyboard focus and mouse focus. 
		//It also handles correctly clicks following the focus (it moves the caret and doesn't reselect the text):
	    //With keyboard focus, only onfocus triggers which selects the text because this.clicked is not set. With mouse focus, onmousedown triggers, then onfocus and then onclick which selects the text in onclick but not in onfocus (Chrome requires this).
		//Mouse clicks when the field is already focused don't trigger onfocus which results in not selecting anything.
		$(".selectFullText").bind({
		    click: function () {
		        if (this.clicked == 2) this.select(); this.clicked = 0;
		    },
		    mousedown: function () {
		        this.clicked = 1;
		    },
		    focus: function () {
		        if (!this.clicked) this.select(); else this.clicked = 2;
		    }
		});

	});

	function refreshTab2Controls() {
		
		if (($('#PageBreakID :selected').text() == 'None') ||
				($('#PercentageOfType').prop('checked') == false)) {
			$('input[name="PercentageOfPage"]').attr('disabled', true);
			$('label[for="PercentageOfPage"]').addClass('ui-state-disabled');
			$('#PercentageOfPage').prop('checked', false);
		}
		else {
			$('input[name="PercentageOfPage"]').attr('disabled', false);
			$('label[for="PercentageOfPage"]').removeClass('ui-state-disabled');
		}

	}
</script>

