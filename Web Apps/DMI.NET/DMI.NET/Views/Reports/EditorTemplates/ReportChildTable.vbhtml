@Imports DMI.NET.ViewModels.Reports
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of ChildTableViewModel)

@code
	Html.BeginForm("PostChildTable", "Reports", FormMethod.Post, New With {.id = "frmPostChildTable"})
End Code

<div class="pageTitleDiv" style="margin-bottom: 15px">
	<span class="pageTitle" id="PopupReportDefinition_PageTitle">Child Tables</span>
</div>

<div class="width100">
	<div id="ReportChildTableMainDiv">
		<div id="ReportChildTableDropdownDiv" class="clearboth">
			<div class="floatleft width20">
				@Html.HiddenFor(Function(m) m.ReportID)
				@Html.HiddenFor(Function(m) m.FilterViewAccess)	 
				@Html.LabelFor(Function(m) m.TableID, New With {.class = ""})
			</div>
			<div class="width80 floatleft">
	 	@Html.TableDropdown("TableID", "ChildTableID", Model.TableID, Model.AvailableTables, "changeChildTable();")
			</div>
		</div>
	 
		<div id="ReportChildTableFilterDiv" class="clearboth" style="">
			<div class="width20 floatleft">
		@Html.HiddenFor(Function(m) m.FilterID, New With {.id = "txtChildFilterID"})
		@Html.LabelFor(Function(m) m.FilterName)
			</div>
			<div class="floatleft width80">
		@Html.TextBoxFor(Function(m) m.FilterName, New With {.id = "txtChildFilter", .readonly = "true"})
		@Html.EllipseButton("cmdBaseFilter", "selectChildTableFilter()", True)
			</div>
		</div>

		<div id="ReportChildTableOrderDiv" class="clearboth">
			<div class="width20 floatleft">
		@Html.LabelFor(Function(m) m.OrderName)
		@Html.HiddenFor(Function(m) m.OrderID, New With {.id = "txtChildFieldOrderID"})
			</div>
			<div class="floatleft width80">
		@Html.TextBoxFor(Function(m) m.OrderName, New With {.id = "txtFieldRecOrder", .readonly = "true"})
				@Html.EllipseButton("cmdBasePicklist", "selectRecordOrder()", True)
			</div>
		</div>

		<div id="ReportChildTableRecordsDiv" class="clearboth">
			<div class="width20 floatleft">
		@Html.LabelFor(Function(m) m.Records)
			</div>
			<div class="floatleft">
		@Html.TextBoxFor(Function(m) m.Records, New With {.id = "txtChildRecords"})
			</div>
		</div>
	</div>

	<div id="divChildTablesButtons" class="clearboth">
		<input type="button" value="OK" onclick="postThisChildTable();" />
		<input type="button" value="Cancel" onclick="closeThisChildTable();" />
	</div>
</div>
		

@Code
	Html.EndForm()
End Code
<script>


	$(function () {

		//some styling
		//$("#ChildTableID").width("100%");
		$('div').css("padding-right", "3");
		//$('div').css("border", "0");

	})

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

		OpenHR.modalExpressionSelect("FILTER", tableID, currentID, function (id, name, access) {
			$("#txtChildFilterID").val(id);
			$("#txtChildFilter").val(name);
			$("#FilterViewAccess").val(access);
				}, 400, 200);

	}

	function selectRecordOrder() {

		var tableID = $("#ChildTableID option:selected").val();
		var currentID = $("#txtChildFieldOrderID").val();

		OpenHR.modalExpressionSelect("ORDER", tableID, currentID, function (id, name, access) {
			$("#txtChildFieldOrderID").val(id);
			$("#txtFieldRecOrder").val(name);
		}, 400, 200);
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
				FilterViewAccess: $("#FilterViewAccess").val(),
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

			setViewAccess('FILTER', $("#ChildTablesViewAccess"), $("#FilterViewAccess").val(), $("#ChildTableID option:selected").text());

			// Post to server
			OpenHR.postData("Reports/PostChildTable", datarow, loadAvailableTablesForReport)

			$("#divPopupReportDefinition").dialog("close");
			$("#divPopupReportDefinition").empty();
		}
</script>