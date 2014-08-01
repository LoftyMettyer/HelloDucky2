@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Classes
@Inherits System.Web.Mvc.WebViewPage(Of Models.CrossTabModel)

<div>
	<fieldset>
		<legend class="fontsmalltitle">
Headings &amp; Breaks
		</legend>

		<fieldset class="CrosstabColumnWidth">
			<legend>Column</legend>

			<div class="display-label_crosstabs">
				Horizontal :
			</div>
			@Html.ColumnDropdownFor(Function(m) m.HorizontalID, New ColumnFilter() _
													 With {.TableID = Model.BaseTableID}, New With {.onchange = "refreshCrossTabColumn(event.target, 'Horizontal');"})
			@Html.ValidationMessageFor(Function(m) m.HorizontalID)
			@Html.Hidden("HorizontalDataType", CInt(Model.HorizontalDataType))
	<br />
			<div class="display-label_crosstabs">
				Vertical :
			</div>
			@Html.ColumnDropdownFor(Function(m) m.VerticalID, New ColumnFilter() _
													 With {.TableID = Model.BaseTableID}, New With {.onchange = "refreshCrossTabColumn(event.target, 'Vertical');"})
			@Html.ValidationMessageFor(Function(m) m.VerticalID)
			@Html.Hidden("VerticalDataType", CInt(Model.VerticalDataType))
			<br />
			<div class="display-label_crosstabs">
				Page Break :
			</div>
			@Html.ColumnDropdownFor(Function(m) m.PageBreakID, New ColumnFilter() _
													 With {.TableID = Model.BaseTableID, .AddNone = True}, New With {.onchange = "refreshCrossTabColumn(event.target, 'PageBreak');"})
			@Html.Hidden("PageBreakDataType", CInt(Model.PageBreakDataType))

		</fieldset>
		<fieldset class="CrosstabColumnWidthStartStopIncrement aligncenter">
			<legend class="aligncenter">Start</legend>
			@Html.EditorFor(Function(m) m.HorizontalStart)
			@Html.EditorFor(Function(m) m.VerticalStart)
			@Html.EditorFor(Function(m) m.PageBreakStart)
		</fieldset>

		<fieldset class="CrosstabColumnWidthStartStopIncrement aligncenter">
			<legend class="aligncenter">Stop</legend>
			@Html.EditorFor(Function(m) m.HorizontalStop)
			@Html.EditorFor(Function(m) m.VerticalStop)
			@Html.EditorFor(Function(m) m.PageBreakStop)
		</fieldset>

		<fieldset class="CrosstabColumnWidthStartStopIncrement aligncenter">
			<legend class="aligncenter">Increment</legend>
			@Html.EditorFor(Function(m) m.HorizontalIncrement)
	@Html.EditorFor(Function(m) m.VerticalIncrement)
	@Html.EditorFor(Function(m) m.PageBreakIncrement)
		</fieldset>
	</fieldset>

	<br/>
	@Html.ValidationMessageFor(Function(m) m.HorizontalStart)
	@Html.ValidationMessageFor(Function(m) m.HorizontalStop)
	@Html.ValidationMessageFor(Function(m) m.HorizontalIncrement)
	@Html.ValidationMessageFor(Function(m) m.VerticalStart)
	@Html.ValidationMessageFor(Function(m) m.VerticalStop)
	@Html.ValidationMessageFor(Function(m) m.VerticalIncrement)
	@Html.ValidationMessageFor(Function(m) m.PageBreakStart)
	@Html.ValidationMessageFor(Function(m) m.PageBreakStop)
	@Html.ValidationMessageFor(Function(m) m.PageBreakIncrement)
</div>

<br/>

<fieldset>
	<legend>Intersection</legend>
	<fieldset class="CrosstabColumnWidth">
		<div class="display-label_crosstabs">
Column:
		</div>
		@Html.ColumnDropdownFor(Function(m) m.IntersectionID, New ColumnFilter() _
													 With {.TableID = Model.BaseTableID, .AddNone = True}, New With {.onchange = "crossTabIntersectionType();"})
		<br />
		<div class="display-label_crosstabs">
	@Html.LabelFor(Function(m) m.IntersectionType)
		</div>
	@Html.EnumDropDownListFor(Function(m) m.IntersectionType)
	</fieldset>

	<fieldset class="display-label_crosstabs">
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
		<br />
	</fieldset>
</fieldset>

<script type="text/javascript">

	function crossTabIntersectionType() {
		var dropDown = $("#IntersectionID")[0];
		var iDataType = dropDown.options[dropDown.selectedIndex].attributes["data-datatype"].value;
		combo_disable($("#IntersectionType"),  (iDataType == "0") )
	}


	function refreshCrossTabColumn(target, type) {

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

		refreshCrossTabColumn($("#HorizontalID")[0], 'Horizontal');
		refreshCrossTabColumn($("#VerticalID")[0], 'Vertical');
		refreshCrossTabColumn($("#PageBreakID")[0], 'PageBreak');
		crossTabIntersectionType();

	});

</script>

