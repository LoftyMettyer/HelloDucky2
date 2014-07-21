﻿@Imports DMI.NET.ViewModels.Reports
@Imports DMI.NET.Helpers
@Imports HR.Intranet.Server.Enums
@Inherits System.Web.Mvc.WebViewPage(Of SortOrderViewModel)

@Code
	Html.BeginForm("PostSortOrder", "Reports", FormMethod.Post, New With {.id = "frmPostSortOrder"})
End Code

<fieldset>

	<legend>Sort Order</legend>

	@Html.HiddenFor(Function(m) m.ID, New With {.id = "SortOrderID"})
	@Html.HiddenFor(Function(m) m.ReportID)
	@Html.HiddenFor(Function(m) m.ReportType)
	@Html.HiddenFor(Function(m) m.TableID, New With {.id = "SortOrderTableID"})
	@Html.HiddenFor(Function(m) m.Sequence, New With {.id = "SortOrderSequence"})

	
	@Html.ColumnDropdown("ColumnID", "SortOrderColumnID", Model.ColumnID, Model.AvailableColumns, "")
	<br/>

	@Html.RadioButton("Order", CInt(OrderType.Ascending), Model.Order = OrderType.Ascending, New With {.id = "SortOrderOrder"})
	Ascending
	<br/>
	@Html.RadioButton("Order", CInt(OrderType.Descending), Model.Order = OrderType.Descending, New With {.id = "SortOrderOrder"})
	Descending
	<br />

	<div class="customReportsOnly">

		@Html.CheckBoxFor(Function(m) m.BreakOnChange)
		@Html.LabelFor(Function(m) m.BreakOnChange)
		<br />

		@Html.CheckBoxFor(Function(m) m.PageOnChange)
		@Html.LabelFor(Function(m) m.PageOnChange)
		<br />

		@Html.CheckBoxFor(Function(m) m.ValueOnChange)
		@Html.LabelFor(Function(m) m.ValueOnChange)
		<br />

		@Html.CheckBoxFor(Function(m) m.SuppressRepeated)
		@Html.LabelFor(Function(m) m.SuppressRepeated)
	</div>

</fieldset>

<input type="button" value="OK" onclick="postThisSortOrder();" />
<input type="button" value="Cancel" onclick="closeThisSortOrder();" />

@Code
	Html.EndForm()
End Code

<script type="text/javascript">

	// Initialise
	$(function () {

		if ('@Model.ReportType' != '@UtilityType.utlCustomReport') {
			$(".customReportsOnly").hide();
		}
	});

	function postThisSortOrder() {

		var datarow = {
			ID: $("#SortOrderID").val(),
			ReportID: '@Model.ReportID',
			ReportType: '@CInt(Model.ReportType)',
			TableID: $("#SortOrderTableID").val(),
			ColumnID: $("#SortOrderColumnID").val(),
			Name: $("#SortOrderColumnID option:selected").text(),
			Order: $("#SortOrderOrder:checked").val(),
			Sequence: $("#SortOrderSequence").val(),
			BreakOnChange: $("#BreakOnChange").is(':checked'),
			PageOnChange: $("#PageOnChange").is(':checked'),
			ValueOnChange: $("#ValueOnChange").is(':checked'),
			SuppressRepeated: $("#SuppressRepeated").is(':checked')
		};

		// Post to server
		OpenHR.postData("Reports/PostSortOrder", datarow)

		// Update client
		$('#SortOrders').jqGrid('delRowData', $("#SortOrderID").val())
		var su = $("#SortOrders").jqGrid('addRowData', $("#SortOrderID").val(), datarow);
		$('#SortOrders').setGridParam({ sortname: 'Sequence' }).trigger('reloadGrid');

		$("#divPopupReportDefinition").dialog("close");
		$("#divPopupReportDefinition").empty();
	}

	function closeThisSortOrder() {
		$("#divPopupReportDefinition").dialog("close");
		$("#divPopupReportDefinition").empty();
	}


</script>