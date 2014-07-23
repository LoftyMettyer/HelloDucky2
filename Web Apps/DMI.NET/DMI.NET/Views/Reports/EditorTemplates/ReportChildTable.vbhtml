@Imports DMI.NET.ViewModels.Reports
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of ChildTableViewModel)

<fieldset>
	@Using (Html.BeginForm("PostChildTable", "Reports", FormMethod.Post, New With {.id = "frmPostChildTable"}))

 		@Html.HiddenFor(Function(m) m.ReportID)

	 	@Html.LabelFor(Function(m) m.TableID) 
	 	@Html.TableDropdown("TableID", "ChildTableID", Model.TableID, Model.AvailableTables, "changeChildTable();")
	 
		@<br/>
		@Html.HiddenFor(Function(m) m.FilterID, New With {.id = "txtChildFilterID"})
		@Html.LabelFor(Function(m) m.FilterName)
		@Html.TextBoxFor(Function(m) m.FilterName, New With {.id = "txtChildFilter", .readonly = "true"})
		@Html.EllipseButton("cmdBaseFilter", "selectChildTableFilter()", True)
		@<br />
		@Html.LabelFor(Function(m) m.OrderName)
		@Html.HiddenFor(Function(m) m.OrderID, New With {.id = "txtChildFieldOrderID"})
		@Html.TextBoxFor(Function(m) m.OrderName, New With {.id = "txtFieldRecOrder", .readonly = "true"})
		@Html.EllipseButton("cmdBaseFilter", "selectRecordOrder()", True)
		@<br />
		@Html.LabelFor(Function(m) m.Records)
		@Html.TextBoxFor(Function(m) m.Records, New With {.id = "txtChildRecords"})
		@<br />
		@<input type="button" value="OK" onclick="postThisChildTable();" />
		@<input type="button" value="Cancel" onclick="closeThisChildTable();" />

		
	End Using 
</fieldset>

<script>

	function changeChildTable() {
		$("#txtChildFilterID").val(0);
		$("#txtChildFilter").val('');
		$("#txtChildFieldOrderID").val(0);
		$("#txtFieldRecOrder").val('');
		$("#txtChildRecords").val(0);
	}

	function selectChildTableFilter() {

		var tableID = $("#ChildTableID option:selected").val();
		var currentID = $("#txtChildFilterID").val();

		OpenHR.modalExpressionSelect("FILTER", tableID, currentID, function (id, name) {
			$("#txtChildFilterID").val(id);
			$("#txtChildFilter").val(name);
		});

	}

	function selectRecordOrder() {

		var tableID = $("#ChildTableID option:selected").val();
		var currentID = $("#txtChildFieldOrderID").val();

		OpenHR.modalExpressionSelect("ORDER", tableID, currentID, function (id, name) {
			$("#txtChildFieldOrderID").val(id);
			$("#txtFieldRecOrder").val(name);
		});

	}

	function closeThisChildTable() {
		$("#divPopupReportDefinition").dialog("close");
		$("#divPopupReportDefinition").empty();
	}

		function postThisChildTable() {

			var datarow = {
				ReportID: '@Model.ReportID',
				ReportType: '@Model.ReportType',
				ID: '@Model.ID',
				TableID: $("#ChildTableID").val(),
				FilterID: $("#txtChildFilterID").val(),
				OrderID: $("#txtChildFieldOrderID").val(),
				TableName: $("#ChildTableID option:selected").text(),
				FilterName: $("#txtChildFilter").val(),
				OrderName: $("#txtFieldRecOrder").val(),
				Records: $("#txtChildRecords").val()
			};

			// Update client
			var grid = $('#ChildTables');
			grid.jqGrid('delRowData', '@Model.ID');
			grid.jqGrid('addRowData', '@Model.ID', datarow);
			grid.setGridParam({ sortname: 'ID' }).trigger('reloadGrid');
			grid.jqGrid("setSelection", '@Model.ID');

			// Post to server
			OpenHR.postData("Reports/PostChildTable", datarow, loadAvailableTablesForReport)

			$("#divPopupReportDefinition").dialog("close");
			$("#divPopupReportDefinition").empty();

		}

</script>