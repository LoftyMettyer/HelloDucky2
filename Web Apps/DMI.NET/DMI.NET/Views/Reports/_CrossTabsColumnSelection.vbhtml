@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports HR.Intranet.Server.Enums
@Inherits System.Web.Mvc.WebViewPage(Of Models.CrossTabModel)

<div>
Headings &amp; Breaks
	<br/>

	Horizontal :  @Html.ColumnDropdown("HorizontalID", Model.HorizontalID, Model.AvailableColumns, "refreshCrossTabColumn(event.target, 'Horizontal');")

	@Html.Hidden("HorizontalDataType", CInt(Model.HorizontalDataType))
	@Html.EditorFor(Function(m) m.HorizontalStart)
	@Html.EditorFor(Function(m) m.HorizontalStop)
	@Html.EditorFor(Function(m) m.HorizontalIncrement)

	<br />
	Vertical :@Html.ColumnDropdown("VerticalID", Model.VerticalID, Model.AvailableColumns, "refreshCrossTabColumn(event.target, 'Vertical');")
	@Html.Hidden("VerticalDataType", CInt(Model.VerticalDataType))
	@Html.EditorFor(Function(m) m.VerticalStart)
	@Html.EditorFor(Function(m) m.VerticalStop)
	@Html.EditorFor(Function(m) m.VerticalIncrement)

	<br/>
	Page Break :@Html.ColumnDropdown("PageBreakID", Model.PageBreakID, Model.AvailableColumns, "refreshCrossTabColumn(event.target, 'PageBreak');")
	@Html.Hidden("PageBreakDataType", CInt(Model.PageBreakDataType))
	@Html.EditorFor(Function(m) m.PageBreakStart)
	@Html.EditorFor(Function(m) m.PageBreakStop)
	@Html.EditorFor(Function(m) m.PageBreakIncrement)

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

<div>

	Intersection:
	<br/>
	Column :
	@Html.ColumnDropdown("IntersectionID", Model.IntersectionID, Model.AvailableColumns, "")
	<br/>
	@Html.LabelFor(Function(m) m.IntersectionType)
	@Html.EnumDropDownListFor(Function(m) m.IntersectionType)
	<br/>

Type :
	<br/>
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

</div>

<script type="text/javascript">

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
				$("#" + type + "Start").val("");
				$("#" + type + "Stop").attr("disabled", "disabled");
				$("#" + type + "Stop").val("");
				$("#" + type + "Increment").attr("disabled", "disabled");
				$("#" + type + "Increment").val("");

		}

	}

	$(function () {

		refreshCrossTabColumn($("#HorizontalID")[0], 'Horizontal');
		refreshCrossTabColumn($("#VerticalID")[0], 'Vertical');
		refreshCrossTabColumn($("#PageBreakID")[0], 'PageBreak');

	});

</script>

