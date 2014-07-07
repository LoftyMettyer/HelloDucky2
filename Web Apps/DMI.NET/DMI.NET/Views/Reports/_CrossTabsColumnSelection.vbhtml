@Imports DMI.NET
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of Models.CrossTabModel)

<div>
Headings &amp; Breaks
	<br/>

	Horizontal :  @Html.ColumnDropdown("HorizontalID", Model.HorizontalID, Model.AvailableColumns, "selectCrossTabColumn(event, 'Horizontal');")
	@Html.TextBox("HorizontalStart", Model.HorizontalStart)
	@Html.TextBox("HorizontalStop", Model.HorizontalStop)
	@Html.TextBox("HorizontalIncrement", Model.HorizontalIncrement)

	@Html.ValidationMessageFor(Function(m) m.HorizontalStart)
	@Html.ValidationMessageFor(Function(m) m.HorizontalStop)
	@Html.ValidationMessageFor(Function(m) m.HorizontalIncrement)


	<br />
	Vertical :@Html.ColumnDropdown("VerticalID", Model.VerticalID, Model.AvailableColumns, "selectCrossTabColumn(event, 'Vertical');")
	@Html.TextBox("VerticalStart", Model.VerticalStart)
	@Html.TextBox("VerticalStop", Model.VerticalStop)
	@Html.TextBox("VerticalIncrement", Model.VerticalIncrement)

	<br/>
	Page Break :@Html.ColumnDropdown("PageBreakID", Model.PageBreakID, Model.AvailableColumns, "selectCrossTabColumn(event, 'PageBreak');")
	@Html.TextBox("PageBreakStart", Model.PageBreakStart)
	@Html.TextBox("PageBreakStop", Model.PageBreakStop)
	@Html.TextBox("PageBreakIncrement", Model.PageBreakIncrement)

</div>

<br/>

<div>

	Intersection:@Html.ColumnDropdown("IntersectionID", Model.IntersectionID, Model.AvailableColumns, "")
	<br/>
	Column :If existing file : @Html.EnumDropDownListFor(Function(m) m.IntersectionType)
	<br/>

Type :

	@Html.CheckBox("PercentageOfType", Model.PercentageOfType)
	@Html.LabelFor(Function(m) m.PercentageOfType)
	<br/>
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

	function selectCrossTabColumn(event, type) {

		var iDataType = event.target.options[event.target.selectedIndex].attributes["data-datatype"].value;

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




</script>

