@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Classes
@Inherits System.Web.Mvc.WebViewPage(Of Models.NineBoxGridModel)

<link href="@Url.LatestContent("~/Content/spectrum.css")" rel="stylesheet" type="text/css" />

<fieldset id="CrossTabsColumnTab" class="width100">
	<legend class="fontsmalltitle">
		Headings &amp; Column Breaks
	</legend>
	<table class="width100">
		<thead class="fontsmalltitle">
			<tr>
				<td style="width:10%"></td>
				<td style="width:55%;text-align:center">Column</td>
				<td style="width:10%;text-align:center">Minimum Value</td>
				<td style="width:10%;text-align:center">Maximum Value</td>
				<td style="width:10%;text-align:center"></td>
			</tr>
		</thead>
		<tr>
			<td style="padding-right: 20px;">Horizontal :</td>
			<td>
				@Html.ColumnDropdownFor(Function(m) m.HorizontalID, New ColumnFilter() With {.TableID = Model.BaseTableID, .IsNumeric = True, .ExcludeOleAndPhoto = True}, New With {.onchange = "crossTabHorizontalChange()", .class = "enableSaveButtonOnComboChange"})
				@Html.ValidationMessageFor(Function(m) m.HorizontalID)
				@Html.Hidden("HorizontalDataType", CInt(Model.HorizontalDataType))
			</td>
			<td class="startstopincrementcol">@Html.TextBoxFor(Function(m) m.HorizontalStart, New With {.class = "selectFullText"})</td>
			<td class="startstopincrementcol">@Html.TextBoxFor(Function(m) m.HorizontalStop, New With {.class = "selectFullText"})</td>
		</tr>
		<tr style="height: 10px;"></tr>
		<tr>
			<td>Vertical :</td>
			<td>
				@Html.ColumnDropdownFor(Function(m) m.VerticalID, New ColumnFilter() With {.TableID = Model.BaseTableID, .IsNumeric = True, .ExcludeOleAndPhoto = True}, New With {.onchange = "crossTabVerticalChange()", .class = "enableSaveButtonOnComboChange"})
				@Html.ValidationMessageFor(Function(m) m.VerticalID)
				@Html.Hidden("VerticalDataType", CInt(Model.VerticalDataType))
			</td>
			<td class="startstopincrementcol">@Html.TextBoxFor(Function(m) m.VerticalStart, New With {.class = "selectFullText"})</td>
			<td class="startstopincrementcol">@Html.TextBoxFor(Function(m) m.VerticalStop, New With {.class = "selectFullText"})</td>
		</tr>
		<tr style="height: 10px;"></tr>
		<tr>
			<td>Page Break :</td>
			<td>
				@Html.ColumnDropdownFor(Function(m) m.PageBreakID, New ColumnFilter() With {.TableID = Model.BaseTableID, .AddNone = True, .ExcludeOleAndPhoto = True}, New With {.onchange = "crossTabPageBreakChange()", .class = "enableSaveButtonOnComboChange"})
				@Html.Hidden("PageBreakDataType", CInt(Model.PageBreakDataType))
			</td>
			<td></td>
			<td></td>
		</tr>
		<tr style="height:20px;"></tr>

		<tr>
			<td style="vertical-align:top;">
				<label style="font-weight: bold;margin-left: -5px;">Label Settings</label>
				<br />
				<label style="font-size:small;">(Click the report labels to edit them)</label>
			</td>
			<td>
				<table id="tblNineBox_def" style="">
					<tr>
						<td class="yaxismajor" rowspan="3">
							<p>@Html.TextBoxFor(Function(m) m.YAxisLabel)</p>
						</td>
						<td class="yaxisminor">
							<p>@Html.TextBoxFor(Function(m) m.YAxisSubLabel1)</p>
						</td>
						<td id="nineBoxR1C1" class="nineBoxGridCell">
							<p>@Html.TextAreaFor(Function(m) m.Description1)</p>
							<p class="pcolpicker">@Html.TextBoxFor(Function(m) m.ColorDesc1)</p>
						</td>
						<td id="nineBoxR1C2" class="nineBoxGridCell">
							<p>@Html.TextAreaFor(Function(m) m.Description2)</p>
							<p class="pcolpicker">@Html.TextBoxFor(Function(m) m.ColorDesc2)</p>
						</td>
						<td id="nineBoxR1C3" class="nineBoxGridCell">
							<p>@Html.TextAreaFor(Function(m) m.Description3)</p>
							<p class="pcolpicker">@Html.TextBoxFor(Function(m) m.ColorDesc3)</p>
						</td>
					</tr>
					<tr>
						<td class="yaxisminor">
							<p>@Html.TextBoxFor(Function(m) m.YAxisSubLabel2)</p>
						</td>
						<td id="nineBoxR2C1" class="nineBoxGridCell">
							<p>@Html.TextAreaFor(Function(m) m.Description4)</p>
							<p class="pcolpicker">@Html.TextBoxFor(Function(m) m.ColorDesc4)</p>
						</td>
						<td id="nineBoxR2C2" class="nineBoxGridCell">
							<p>@Html.TextAreaFor(Function(m) m.Description5)</p>
							<p class="pcolpicker">@Html.TextBoxFor(Function(m) m.ColorDesc5)</p>
						</td>
						<td id="nineBoxR2C3" class="nineBoxGridCell">
							<p>@Html.TextAreaFor(Function(m) m.Description6)</p>
							<p class="pcolpicker">@Html.TextBoxFor(Function(m) m.ColorDesc6)</p>
						</td>
					</tr>
					<tr>
						<td class="yaxisminor">
							<p>@Html.TextBoxFor(Function(m) m.YAxisSubLabel3)</p>
						</td>
						<td id="nineBoxR3C1" class="nineBoxGridCell">
							<p>@Html.TextAreaFor(Function(m) m.Description7)</p>
							<p class="pcolpicker">@Html.TextBoxFor(Function(m) m.ColorDesc7)</p>
						</td>
						<td id="nineBoxR3C2" class="nineBoxGridCell">
							<p>@Html.TextAreaFor(Function(m) m.Description8)</p>
							<p class="pcolpicker">@Html.TextBoxFor(Function(m) m.ColorDesc8)</p>
						</td>
						<td id="nineBoxR3C3" class="nineBoxGridCell">
							<p>@Html.TextAreaFor(Function(m) m.Description9)</p>
							<p class="pcolpicker">@Html.TextBoxFor(Function(m) m.ColorDesc9)</p>
						</td>
					</tr>
					<tr>
						<td colspan="2" rowspan="2" class="xaxis"></td>
						<td class="xaxisminor">@Html.TextBoxFor(Function(m) m.XAxisSubLabel1)</td>
						<td class="xaxisminor">@Html.TextBoxFor(Function(m) m.XAxisSubLabel2)</td>
						<td class="xaxisminor">@Html.TextBoxFor(Function(m) m.XAxisSubLabel3)</td>
					</tr>
					<tr>
						<td colspan="3" class="xaxisminor">@Html.TextBoxFor(Function(m) m.XAxisLabel)</td>
					</tr>
				</table>
			</td>
			<td colspan="2" id="td9boxOptions">
				<label style="font-weight: bold;">Display Options</label>
				<br />
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
	</table>
</fieldset>



@Html.Hidden("IntersectionID", CInt(Model.IntersectionID))
@Html.Hidden("IntersectionType", CInt(Model.IntersectionType))
@Html.Hidden("PageBreakStart", CDbl(Model.PageBreakStart))
@Html.Hidden("PageBreakStop", CDbl(Model.PageBreakStop))

<br />
@Html.ValidationMessageFor(Function(m) m.HorizontalStart)
@Html.ValidationMessageFor(Function(m) m.HorizontalStop)
@Html.ValidationMessageFor(Function(m) m.VerticalStart)
@Html.ValidationMessageFor(Function(m) m.VerticalStop)


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

					// Add all available columns to PageBreak dropdown
					optionPageBreak += "<option value='" + json[i].ID + "' data-datatype='" + json[i].DataType + "' data-size='" + json[i].ColumnSize + "' data-decimals='" + json[i].Decimals + "'>" + json[i].Name + "</option>";

					if (!json[i].IsNumeric) { //Only add numeric columns to Horizontal & Vertical dropdown
						continue;
					}

					optionHorizontal += "<option value='" + json[i].ID + "' data-datatype='" + json[i].DataType + "' data-size='" + json[i].ColumnSize + "' data-decimals='" + json[i].Decimals + "'>" + json[i].Name + "</option>";
					optionVertical += "<option value='" + json[i].ID + "' data-datatype='" + json[i].DataType + "' data-size='" + json[i].ColumnSize + "' data-decimals='" + json[i].Decimals + "'>" + json[i].Name + "</option>";

					if (json[i].IsNumeric) {
						optionIntersection += "<option value='" + json[i].ID + "' data-datatype='" + json[i].DataType + "' data-size='" + json[i].ColumnSize + "' data-decimals='" + json[i].Decimals + "'>" + json[i].Name + "</option>";
					}
				}

				$("select#HorizontalID").html(optionHorizontal);
				$("select#VerticalID").html(optionVertical);
				$("select#PageBreakID").html(OptionNone + optionPageBreak);
				$("select#IntersectionID").html(OptionNone + optionIntersection);

				// Reset data on base table change
				crossTabHorizontalChange();
				crossTabVerticalChange();
				crossTabPageBreakChange();
			}
		});
	}

	function refreshCrossTabColumn(target, type) {
		if (target.options.length == 0) //Return if the selected base table doesn't have any numeric columns
			return;

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
				break;

			case "4":
				$("#" + type + "Start").attr("disabled", bReadOnly);
				$("#" + type + "Stop").attr("disabled", bReadOnly);
				break;

			default:
				$("#" + type + "Start").attr("disabled", "disabled");
				$("#" + type + "Start").val(0);
				$("#" + type + "Stop").attr("disabled", "disabled");
				$("#" + type + "Stop").val(0);
		}

		$("#" + type + "Start").autoNumeric('destroy');
		$("#" + type + "Stop").autoNumeric('destroy');

		// Use autoNumeric formatting only for Interger (datatype 4) and Numeric (datatype 2) value
		if (iDataType == 2 || iDataType == 4) {

			// For the integer datatype, 10 digits are allowed and max value is 9999999999
			var maximumValue = 9999999999;

			// If the numeric type, set the maximum value which can be entered before decimal
			if (iDataType == 2) {

				var columnSize = target.options[target.selectedIndex].attributes["data-size"].value;

				var maxValueAllowedBeforeDecimal = '9';
				for (var i = 1; i < (columnSize - iDecimals) ; i++) {
					maxValueAllowedBeforeDecimal = maxValueAllowedBeforeDecimal + '9';
				}

				var maxValueAllowedAfterDecimal = '9'
				for (var i = 1; i < iDecimals ; i++) {
					maxValueAllowedAfterDecimal = maxValueAllowedAfterDecimal + '9';
				}

				// Set the maximum value
				maximumValue = parseFloat(maxValueAllowedBeforeDecimal + "." + maxValueAllowedAfterDecimal);

				// Reset data to 0 when exceeds to maximum value. This check is required for the existing data.
				// Otherwise this will throw an exeption and application goes to unrecoverable state when binding autoNumeric (as below code).
				ResetWhenExceedsMaximumAllowedValue(type, maximumValue);
			}

			$("#" + type + "Start").autoNumeric({ aSep: '', aNeg: '', mDec: iDecimals, mRound: 'S', vMin: -999999999.99, vMax: maximumValue });
			$("#" + type + "Stop").autoNumeric({ aSep: '', aNeg: '', mDec: iDecimals, mRound: 'S', vMin: -999999999.99, vMax: maximumValue });
		}
	}

	// Reset the Horizontal,Vertical and Incremental values to 0 when they exceeds maximum allowed limit
	function ResetWhenExceedsMaximumAllowedValue(control, maxValueAllowed) {
		// Check Start value
		if ($("#" + control + "Start").val() > maxValueAllowed) {
			$("#" + control + "Start").val(0);
		}

		// Check Stop value
		if ($("#" + control + "Stop").val() > maxValueAllowed) {
			$("#" + control + "Stop").val(0);
		}
	}

	function crossTabHorizontalChange() {
		$("#HorizontalStart").val(0);
		$("#HorizontalStop").val(0);
		crossTabHorizontalClick();
	}

	function crossTabVerticalChange() {
		$("#VerticalStart").val(0);
		$("#VerticalStop").val(0);
		crossTabVerticalClick();
	}

	function crossTabPageBreakChange() {
		$("#PageBreakStart").val(0);
		$("#PageBreakStop").val(0);
		crossTabPageBreakClick();
	}

	function crossTabHorizontalClick() {

		var horval = $("#HorizontalID").val();
		var vertval = $("#VerticalID").val();

		//reset ver and pb so none are disabled/hidden
		$('#VerticalID option').removeAttr('disabled');
		$('#PageBreakID option').removeAttr('disabled');

		//Hide/disable matching items in vertical
		$('#VerticalID option').filter(function () {
			return $(this).val() == horval;
		}).attr('disabled', 'disabled');

		//Hide/disable matching items in pagebreak
		$('#PageBreakID option').filter(function () {
			return $(this).val() == vertval || $(this).val() == horval;
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

		//Reset hor and pb so none are disabled/hidden
		$('#HorizontalID option').removeAttr('disabled');
		$('#PageBreakID option').removeAttr('disabled');

		//Hide/disable matching items in horizontal col
		$('#HorizontalID option').filter(function () {
			return $(this).val() == vertval;
		}).attr('disabled', 'disabled');

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
		$('#PercentageOfType, label[for=PercentageOfType], #PercentageOfPage, label[for=PercentageOfPage]').css('display', 'none');

		$("#CrossTabsColumnTab select").css("width", "100%");
		$('table').attr('border', '0');

		crossTabHorizontalClick();
		crossTabVerticalClick();
		crossTabPageBreakClick();

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

	function initializeColorPicker(colorPickerId) {
		var ID = "ColorDesc" + colorPickerId;

		if ($("#" + ID).val() == "") { //Set default color to black if empty
			$("#" + ID).val("000000");
		}

		$("#" + ID).spectrum("destroy");
		$("#" + ID).spectrum({
			color: "#" + $("#" + ID).val(), //Set the initial color
			className: "nineboxgridColorpicker",
			showInput: true, //Show a textbox with the selected color in hex
			cancelText: "", //Hide the Cancel button
			change: function (color) { //On selecting a color...
				$("#" + ID).val(color.toHex()).change(); //Set the new color and trigger the change event so the Save button is enabled
				$(this).closest('td').css('background-color', color.toHexString());
			}
		});

		$("#" + ID).closest('td').css('background-color', '#' + $("#" + ID).val());

		$("#" + colorPickerId).next().css("top", $("#" + ID).attr("data-style-top") - 2 + "px");
		$("#" + ID).next().css("left", $("#" + ID).attr("data-style-left") - 1 + "px");
		$("#" + ID).next().css("height", $("#" + ID).attr("data-style-height") + "px");
		$("#" + ID).next().css("width", $("#" + ID).attr("data-style-width") + "px");
		$("#" + ID).next().css("position", "absolute");
		$("#" + ID).next().css("background", "none");
		$("#" + ID).next().css("border", "none");
		//First inner div of the div above
		$("#" + ID).next().children().first("sp-preview").css("height", $("#" + ID).attr("data-style-height") + "px");
		$("#" + ID).next().children().first("sp-preview").css("width", $("#" + ID).attr("data-style-width") - 16 + "px");
	}

	//Initialize the 9 color pickers
	for (i = 1 ; i <= 9; i++) {
		initializeColorPicker(i);
	}

	//On leaving the page, clear up any remaining colour picker debris
	$("#CrossTabsColumnTab").off('remove').on('remove', function () {
		$(".nineboxgridColorpicker").remove();
	});


</script>

