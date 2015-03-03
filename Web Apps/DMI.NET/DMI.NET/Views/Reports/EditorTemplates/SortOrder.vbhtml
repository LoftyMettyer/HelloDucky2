@Imports DMI.NET.ViewModels.Reports
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of SortOrderViewModel)

@Code	
	Html.BeginForm("PostSortOrder", "Reports", FormMethod.Post, New With {.id = "frmPostSortOrder"})
End Code

<div class="pageTitleDiv padbot10">
	<span class="pageTitle" id="PopupReportDefinition_PageTitle">Sort Order</span>
</div>

<div class="padleft20">
	<div class="padbot10">
		@Html.HiddenFor(Function(m) m.ID, New With {.id = "SortOrderID"})
		@Html.HiddenFor(Function(m) m.ReportID)
		@Html.HiddenFor(Function(m) m.ReportType)
		@Html.HiddenFor(Function(m) m.TableID, New With {.id = "SortOrderTableID"})
		@Html.HiddenFor(Function(m) m.Sequence, New With {.id = "SortOrderSequence"})
		@Html.HiddenFor(Function(m) m.IsNew)
		@Html.LabelFor(Function(m) m.ColumnID)
		@Html.ColumnDropdown("ColumnID", "SortOrderColumnID", Model.ColumnID, Model.AvailableColumns, "updateCheckBoxes();")
	</div>

	<div class="padbot10">
		@Html.RadioButton("Order", CInt(OrderType.Ascending), Model.Order = OrderType.Ascending, New With {.id = "SortOrderOrder"})
		Ascending
		<br />
		@Html.RadioButton("Order", CInt(OrderType.Descending), Model.Order = OrderType.Descending, New With {.id = "SortOrderOrder"})
		Descending
	</div>

	<div class="customReportsOnly padbot10">
		@Html.CheckBoxFor(Function(m) m.BreakOnChange, New With {.onchange = "updatePageAndBreakOnChangeCheckBox(this)"})
		@Html.LabelFor(Function(m) m.BreakOnChange)
		<br />

		@Html.CheckBoxFor(Function(m) m.PageOnChange, New With {.onchange = "updatePageAndBreakOnChangeCheckBox(this)"})
		@Html.LabelFor(Function(m) m.PageOnChange)
		<br />

		@Html.CheckBoxFor(Function(m) m.ValueOnChange)
		@Html.LabelFor(Function(m) m.ValueOnChange)
		<br />

		@Html.CheckBoxFor(Function(m) m.SuppressRepeated)
		@Html.LabelFor(Function(m) m.SuppressRepeated)
	</div>
</div>


<div id="divSortOrderButtons">
	<input type="button" value="OK" onclick="postThisSortOrder();" />
	<input type="button" value="Cancel" id="butSortOrderEditCancel" onclick="closeThisSortOrder(false);" />
</div>

@Code
	Html.EndForm()
End Code

<script type="text/javascript">
	// Initialise
	$(function () {

		if ('@Model.ReportType' != '@UtilityType.utlCustomReport') {
			$(".customReportsOnly").hide();
		}

		if ('@Model.ReportType' == '@UtilityType.utlCustomReport') {
			updateCheckBoxes();
		}

		updatePageAndBreakOnChangeCheckBox($("#BreakOnChange")[0]);
		updatePageAndBreakOnChangeCheckBox($("#PageOnChange")[0]);

		// Add width to column dropdown
		$("#SortOrderColumnID").addClass('stretchyfixed600');

		if (isDefinitionReadOnly()) {
			$("#frmPostSortOrder input").prop('disabled', "disabled");
			$("#frmPostSortOrder select").prop('disabled', "disabled");
			$("#frmPostSortOrder :button").prop('disabled', "disabled");
		}

		button_disable($("#butSortOrderEditCancel")[0], false);
	});

	function updateCheckBoxes() {

		if ($('#SortOrderColumnID').find('option:selected').attr('data-ishidden').toUpperCase() == "TRUE") {
			$("#frmPostSortOrder  #SuppressRepeated, #frmPostSortOrder #ValueOnChange").prop({ checked: false, disabled: true });
			$("#frmPostSortOrder label[for='ValueOnChange'], #frmPostSortOrder label[for='SuppressRepeated']").css('opacity', '0.5');
		} else {
			$("#frmPostSortOrder #SuppressRepeated, #frmPostSortOrder #ValueOnChange").prop({ disabled: false });
			$("#frmPostSortOrder label[for='ValueOnChange'], #frmPostSortOrder label[for='SuppressRepeated']").css('opacity', '1');
		}
	}

	function updatePageAndBreakOnChangeCheckBox(control) {
		if (control.checked) {
			if (control.id == "BreakOnChange") {
				$("#frmPostSortOrder  #PageOnChange").prop({ checked: false, disabled: true });
				$("#frmPostSortOrder  #PageOnChange, #frmPostSortOrder label[for='PageOnChange']").css('opacity', '0.5');
			}
			if (control.id == "PageOnChange") {
				$("#frmPostSortOrder  #BreakOnChange").prop({ checked: false, disabled: true });
				$("#frmPostSortOrder  #BreakOnChange, #frmPostSortOrder label[for='BreakOnChange']").css('opacity', '0.5');
			}
		}
		else {
			if (control.id == "BreakOnChange") {
				$("#frmPostSortOrder  #PageOnChange").prop({ disabled: false });
				$("#frmPostSortOrder  #PageOnChange, #frmPostSortOrder label[for='PageOnChange']").css('opacity', '1');
			}
			if (control.id == "PageOnChange") {
				$("#frmPostSortOrder  #BreakOnChange").prop({ disabled: false });
				$("#frmPostSortOrder  #BreakOnChange, #frmPostSortOrder label[for='BreakOnChange']").css('opacity', '1');
			}
		}
	}

	function postThisSortOrder() {

		var datarow = {
			ID: '@Model.ID',
			ReportID: '@Model.ReportID',
			ReportType: '@CInt(Model.ReportType)',
			TableID: $("#frmPostSortOrder #SortOrderTableID").val(),
			ColumnID: $("#frmPostSortOrder #SortOrderColumnID").val(),
			Name: $("#frmPostSortOrder #SortOrderColumnID option:selected").text(),
			Order: $("#frmPostSortOrder #SortOrderOrder:checked").val(),
			Sequence: $("#frmPostSortOrder #SortOrderSequence").val(),
			BreakOnChange: $("#frmPostSortOrder #BreakOnChange").is(':checked'),
			PageOnChange: $("#frmPostSortOrder #PageOnChange").is(':checked'),
			ValueOnChange: $("#frmPostSortOrder #ValueOnChange").is(':checked'),
			SuppressRepeated: $("#frmPostSortOrder #SuppressRepeated").is(':checked')
		};

		// Post to server
		OpenHR.postData("Reports/PostSortOrder", datarow)

		// Update client
		$('#SortOrders').jqGrid('delRowData', '@Model.ID')
		var su = $("#SortOrders").jqGrid('addRowData', '@Model.ID', datarow);
		
		$('#SortOrders').setGridParam({ sortname: 'Sequence' }).trigger('reloadGrid');		
		$('#SortOrders').jqGrid("setSelection", '@Model.ID');	

		// Set rRowid back to what it was
		if ($("#IsNew").val().toUpperCase() == "TRUE") {
			$("#SortOrdersAvailable").val(parseInt($("#SortOrdersAvailable").val()) - 1);
		}

		button_disable($("#btnSortOrderAdd")[0], ($("#SortOrdersAvailable").val() == 0));
		closeThisSortOrder(true);
	}

	function closeThisSortOrder(isEnableSaveButton) {
		if (isEnableSaveButton) {
			enableSaveButton();
		}
		$("#divPopupReportDefinition").dialog("close");
		$("#divPopupReportDefinition").empty();
	}
</script>