@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Classes
@Inherits System.Web.Mvc.WebViewPage(Of Models.CrossTabModel)

<fieldset id="CrossTabsColumnTab" class="width100">
		<legend class="fontsmalltitle">
			Headings &amp; Column Breaks
		</legend>
		<fieldset>
			<table class="width80">
				<thead class="fontsmalltitle">
					<tr>
						<td style="width:15%"></td>
						<td style="width:55%;text-align:center">Columns</td>
						<td style="width:10%;text-align:center">Start</td>
						<td style="width:10%;text-align:center">Stop</td>
						<td style="width:10%;text-align:center">Increment</td>
					</tr>
				</thead>
				<tr>
					<td>Horizontal :</td>
					<td>
						@Html.ColumnDropdownFor(Function(m) m.HorizontalID, New ColumnFilter() With {.TableID = Model.BaseTableID}, New With {.class = "crosstabDropdown", .onchange = "refreshCrossTabColumn(event.target, 'Horizontal');"})
						@Html.ValidationMessageFor(Function(m) m.HorizontalID)
						@Html.Hidden("HorizontalDataType", CInt(Model.HorizontalDataType))
					</td>
					<td class="startstopincrementcol">@Html.EditorFor(Function(m) m.HorizontalStart)</td>
					<td class="startstopincrementcol">@Html.EditorFor(Function(m) m.HorizontalStop)</td>
					<td class="startstopincrementcol">@Html.EditorFor(Function(m) m.HorizontalIncrement)</td>
				</tr>
				<tr>
					<td>Vertical :</td>
					<td>
						@Html.ColumnDropdownFor(Function(m) m.VerticalID, New ColumnFilter() With {.TableID = Model.BaseTableID}, New With {.class = "crosstabDropdown", .onchange = "refreshCrossTabColumn(event.target, 'Vertical');"})
						@Html.ValidationMessageFor(Function(m) m.VerticalID)
						@Html.Hidden("VerticalDataType", CInt(Model.VerticalDataType))
					</td>
					<td class="startstopincrementcol">@Html.EditorFor(Function(m) m.VerticalStart)</td>
					<td class="startstopincrementcol">@Html.EditorFor(Function(m) m.VerticalStop)</td>
					<td class="startstopincrementcol">@Html.EditorFor(Function(m) m.VerticalIncrement)</td>
				</tr>
				<tr>
					<td>Page Break :</td>
					<td>
						@Html.ColumnDropdownFor(Function(m) m.PageBreakID, New ColumnFilter() With {.TableID = Model.BaseTableID, .AddNone = True}, New With {.class = "crosstabDropdown", .onchange = "refreshCrossTabColumn(event.target, 'PageBreak');"})
						@Html.Hidden("PageBreakDataType", CInt(Model.PageBreakDataType))
					</td>
					<td class="startstopincrementcol">@Html.EditorFor(Function(m) m.PageBreakStart)</td>
					<td class="startstopincrementcol">@Html.EditorFor(Function(m) m.PageBreakStop)</td>
					<td class="startstopincrementcol">@Html.EditorFor(Function(m) m.PageBreakIncrement)</td>
				</tr>
			</table>
			<br />
			<table class="width80">
				<thead class="fontsmalltitle">
					<tr>
						<td style="width:15%">Intersection :</td>
						<td style="width:55%"></td>
						<td style="width:10%;text-align:center"></td>
						<td style="width:10%;text-align:center"></td>
						<td style="width:10%;text-align:center"></td>
					</tr>
				</thead>
				<tr>
					<td>Column :</td>
					<td>
						@Html.ColumnDropdownFor(Function(m) m.IntersectionID, New ColumnFilter() With {.TableID = Model.BaseTableID, .AddNone = True}, New With {.onchange = "crossTabIntersectionType();"})
					</td>
					<td colspan="3" rowspan="3" style="padding-left:10px">
						@Html.CheckBox("PercentageOfType", Model.PercentageOfType)
						@Html.LabelFor(Function(m) m.PercentageOfType)
						<br />
						@Html.CheckBox("PercentageOfPage", Model.PercentageOfPage)
						@Html.LabelFor(Function(m) m.PercentageOfPage)
						<br />
						@Html.CheckBox("SuppressZeros", Model.SuppressZeros)
						@Html.LabelFor(Function(m) m.SuppressZeros)
						<br />
						@Html.CheckBox("UseThousandSeparators", Model.UseThousandSeparators)
						@Html.LabelFor(Function(m) m.UseThousandSeparators)
					</td>
				</tr>
				<tr>
					<td>@Html.LabelFor(Function(m) m.IntersectionType)</td>
					<td>
						@Html.EnumDropDownListFor(Function(m) m.IntersectionType)
					</td>
				</tr>
				<tr style="height:22px"></tr>
			</table>

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
		</fieldset>
	</fieldset>

<script type="text/javascript">

	function refreshCrossTabColumnsAvailable() {

		$.ajax({
			url: 'Reports/GetAvailableColumnsForCrossTab?TableID=' + $("#BaseTableID").val() + '&&ReportID=' + '@Model.ID',
			datatype: 'json',
			mtype: 'GET',
			success: function (json) {

				var OptionNone = '<option value=0 data-datatype=0 selected>None</option>';
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

				refreshCrossTabColumn($("#HorizontalID")[0],"Horizontal");

			}
		});

	}

	function crossTabIntersectionType() {
		var dropDown = $("#IntersectionID")[0];
		var iDataType = dropDown.options[dropDown.selectedIndex].attributes["data-datatype"].value;
		combo_disable($("#IntersectionType"), (iDataType == "0"));
	}

	function refreshCrossTabColumn(target, type) {

		var horizontalValue = $("#HorizontalID").val();
		var verticalValue = $("#VerticalID").val();	
		var pageBreakValue = $("#PageBreakID").val();

		var iDataType = target.options[target.selectedIndex].attributes["data-datatype"].value;
		$("#" + type + "DataType").val(iDataType);
		switch (iDataType) {
			case "2":
				$("#" + type + "Start").removeAttr("disabled");
				$("#" + type + "Stop").removeAttr("disabled");
				$("#" + type + "Increment").removeAttr("disabled");
				break;

			case "4":
				$("#" + type + "Start").removeAttr("disabled");
				$("#" + type + "Stop").removeAttr("disabled");
				$("#" + type + "Increment").removeAttr("disabled");
				break;

			default:
				$("#" + type + "Start").attr("disabled", "disabled");
				$("#" + type + "Start").val(0);
				$("#" + type + "Stop").attr("disabled", "disabled");
				$("#" + type + "Stop").val(0);
				$("#" + type + "Increment").attr("disabled", "disabled");
				$("#" + type + "Increment").val(0);
		}
	}

	$(function () {

		$('#VerticalID option').clone().appendTo('#hiddenVertical');
		$('#PageBreakID option').clone().appendTo('#hiddenPageBreak');

		refreshCrossTabColumn($("#HorizontalID")[0], 'Horizontal');
		refreshCrossTabColumn($("#VerticalID")[0], 'Vertical');
		refreshCrossTabColumn($("#PageBreakID")[0], 'PageBreak');
		crossTabIntersectionType();

		$("#CrossTabsColumnTab select").css("width", "100%");
		$('table').attr('border', '0');

	});

</script>

