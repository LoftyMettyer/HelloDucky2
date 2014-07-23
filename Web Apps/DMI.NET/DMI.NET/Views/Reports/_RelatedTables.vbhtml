@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Enums
@Imports DMI.NET.Classes
@Imports DMI.NET.Code.Extensions
@Inherits System.Web.Mvc.WebViewPage(Of Models.CustomReportModel)


<div id="divReportParents">

	@Html.CheckBoxFor(Function(m) m.p1Hidden, New With {.Type = "hidden"})
	@Html.CheckBoxFor(Function(m) m.p2Hidden, New With {.Type = "hidden"})
	@Html.CheckBoxFor(Function(m) m.childHidden, New With {.Type = "hidden"})

	<fieldset @Model.Parent1.Visibility>
		<legend>Parent 1 :</legend>

		@Html.HiddenFor(Function(m) m.ChildTablesString, New With {.id = "txtCTAAS"})

		<input type="hidden" id="txtParent1ID" name="Parent1.ID" value="@Model.Parent1.ID" />
		Table:
		@Html.TextBoxFor(Function(m) m.Parent1.Name, New With {.readonly = "true"})
		<br />
		@Html.RadioButton("Parent1.Selectiontype", RecordSelectionType.AllRecords, Model.Parent1.SelectionType = RecordSelectionType.AllRecords, New With {.onclick = "changeRecordOption('Parent1','all')"})
		All Records
		<br />

		@Html.RadioButton("Parent1.SelectionType", RecordSelectionType.Picklist, Model.Parent1.SelectionType = RecordSelectionType.Picklist, New With {.onclick = "changeRecordOption('Parent1','picklist')"})
		Picklist
		<input type="hidden" id="txtParent1PicklistID" name="Parent1.PicklistID" value="@Model.Parent1.PicklistID" />
		@Html.TextBoxFor(Function(m) m.Parent1.PicklistName, New With {.id = "txtParent1Picklist", .readonly = "true"})
		@Html.EllipseButton("cmdParent1Picklist", "selectParent1Picklist()", Model.Parent1.SelectionType = RecordSelectionType.Picklist)
		@Html.ValidationMessageFor(Function(m) m.Parent1.PicklistID)
		<br />

		@Html.RadioButton("Parent1.SelectionType", RecordSelectionType.Filter, Model.Parent1.SelectionType = RecordSelectionType.Filter, New With {.onclick = "changeRecordOption('Parent1','filter')"})
		Filter
		<input type="hidden" id="txtParent1FilterID" name="Parent1.FilterID" value="@Model.Parent1.FilterID" />
		@Html.TextBoxFor(Function(m) m.Parent1.FilterName, New With {.id = "txtParent1Filter", .readonly = "true"})
		@Html.EllipseButton("cmdParent1Filter", "selectParent1Filter()", Model.Parent1.SelectionType = RecordSelectionType.Filter)
		@Html.ValidationMessageFor(Function(m) m.Parent1.FilterID)

	</fieldset>

	<fieldset @Model.Parent2.Visibility>
		<legend>Parent 2 :</legend>

		<input type="hidden" id="txtParent2ID" name="Parent2.ID" value="@Model.Parent2.ID" />
		@Html.TextBoxFor(Function(m) m.Parent2.Name, New With {.readonly = "true"})
		<br />
		@Html.RadioButton("Parent2.Selectiontype", RecordSelectionType.AllRecords, Model.Parent2.SelectionType = RecordSelectionType.AllRecords, New With {.onclick = "changeRecordOption('Parent2','all')"})
		All Records
		<br />
		@Html.RadioButton("Parent2.SelectionType", RecordSelectionType.Picklist, Model.Parent2.SelectionType = RecordSelectionType.Picklist, New With {.onclick = "changeRecordOption('Parent2','picklist')"})
		Picklist
		<input type="hidden" id="txtParent2PicklistID" name="Parent2.PicklistID" value="@Model.Parent2.PicklistID" />
		@Html.TextBoxFor(Function(m) m.Parent2.PicklistName, New With {.id = "txtParent2Picklist", .readonly = "true"})
		@Html.EllipseButton("cmdParent2Picklist", "selectParent2Picklist()", Model.Parent2.SelectionType = RecordSelectionType.Picklist)
		@Html.ValidationMessageFor(Function(m) m.Parent2.PicklistID)
		<br />

		@Html.RadioButton("Parent2.SelectionType", RecordSelectionType.Filter, Model.Parent2.SelectionType = RecordSelectionType.Filter, New With {.onclick = "changeRecordOption('Parent2','filter')"})
		Filter
		<input type="hidden" id="txtParent2FilterID" name="Parent2.FilterID" value="@Model.Parent2.FilterID" />
		@Html.TextBoxFor(Function(m) m.Parent2.FilterName, New With {.id = "txtParent2Filter", .readonly = "true"})
		@Html.EllipseButton("cmdParent2Filter", "selectParent2Filter()", Model.Parent2.SelectionType = RecordSelectionType.Filter)
		@Html.ValidationMessageFor(Function(m) m.Parent2.FilterID)

	</fieldset>

</div>

<br/>

<fieldset class="relatedtables">
	<legend>Child Tables :</legend>

	<div class="stretchyfill">
		<table id ="ChildTables"></table>
	</div>

	<div class="stretchyfixed">
		<input type="button" id="btnChildAdd" value="Add..." onclick="addChildTable();" />
		<br/>
		<input type="button" id="btnChildEdit" value="Edit..." disabled onclick="editChildTable(0);" />
		<br />
		<input type="button" id="btnChildRemove" value="Remove" disabled onclick="removeChildTable();" />
		<br />
		<input type="button" id="btnChildRemoveAll" value="Remove All" disabled onclick="removeAllChildTables();" />				
	</div>

</fieldset>



<script type="text/javascript">

	function selectParent1Picklist() {

		var tableID = $("#txtParent1ID").val();
		var currentID = $("#txtParent1PicklistID").val();

		OpenHR.modalExpressionSelect("PICKLIST", tableID, currentID, function (id, name) {
			$("#txtParent1PicklistID").val(id);
			$("#txtParent1Picklist").val(name);
		});

	}

	function selectParent2Picklist() {

		var tableID = $("#txtParent2ID").val();
		var currentID = $("#txtParent2PicklistID").val();

		OpenHR.modalExpressionSelect("PICKLIST", tableID, currentID, function (id, name) {
			$("#txtParent2PicklistID").val(id);
			$("#txtParent2Picklist").val(name);
		});

	}

	function selectParent1Filter() {

		var tableID = $("#txtParent1ID").val();
		var currentID = $("#txtParent1FilterID").val();

		OpenHR.modalExpressionSelect("FILTER", tableID, currentID, function (id, name) {
			$("#txtParent1FilterID").val(id);
			$("#txtParent1Filter").val(name);
		});

	}

	function selectParent2Filter() {

		var tableID = $("#txtParent2ID").val();
		var currentID = $("#txtParent2FilterID").val();

		OpenHR.modalExpressionSelect("FILTER", tableID, currentID, function (id, name) {
			$("#txtParent2FilterID").val(id);
			$("#txtParent2Filter").val(name);
		});

	}


	function addChildTable() {

		OpenHR.OpenDialog("Reports/AddChildTable", "divPopupReportDefinition", { ReportID: "@Model.ID" });

	}

	function editChildTable(rowID) {

		if (rowID == 0) {
			rowID = $('#ChildTables').jqGrid('getGridParam', 'selrow');
		}

		var gridData = $("#ChildTables").getRowData(rowID);
		OpenHR.OpenDialog("Reports/EditChildTable", "divPopupReportDefinition", gridData);

	}

	function removeChildTable() {
		rowID = $('#ChildTables').jqGrid('getGridParam', 'selrow');
		$('#ChildTables').jqGrid('delRowData', rowID)
		loadAvailableTablesForReport();
	}
	
	$(function () {

		jQuery("#ChildTables").jqGrid({
			datatype: "jsonstring",
			datastr: '@Model.ChildTables.ToJsonResult',
			mtype: 'GET',
			jsonReader: {
				root: "rows", //array containing actual data
				page: "page", //current page
				total: "total", //total pages for the query
				records: "records", //total number of records
				repeatitems: false,
				id: "ID" //index of the column with the PK in it
			},
			colNames: ['ID', 'ReportID', 'ReportType', 'TableID', 'FilterID', 'OrderID', 'Table', 'Filter', 'Order', 'Records'],
			colModel: [
				{ name: 'ID', index: 'ID', sorttype: 'int', hidden: true },
				{ name: 'ReportID', index: 'ReportID', sorttype: 'int', hidden: true },
				{ name: 'ReportType', index: 'ReportType', sorttype: 'int', hidden: true },
				{ name: 'TableID', index: 'TableID', width: 100, hidden: true },
				{ name: 'FilterID', index: 'FilterID', width: 100, hidden: true },
				{ name: 'OrderID', index: 'OrderID', width: 100, hidden: true },
				{ name: 'TableName', index: 'TableName', width: 100 },
				{ name: 'FilterName', index: 'FilterName', width: 100 },
				{ name: 'OrderName', index: 'OrderName', width: 100 },
			{ name: 'Records', index: 'Records', width: 100 }
			],
			rowNum: 10,
			autowidth: true,
			rowTotal: 50,
			rowList: [10, 20, 30],
			shrinkToFit: true,
			pager: '#pcrud',
			sortname: 'TableName',
			loadonce: true,
			viewrecords: true,
			sortorder: "asc",
			ondblClickRow: function (rowID) {
				editChildTable(rowID);

			},
			onSelectRow: function (id) {
				button_disable($("#btnChildEdit")[0], false);
				button_disable($("#btnChildRemove")[0], false);

			},
			gridComplete: function() {
				var tablesSelected = $(this).getGridParam("reccount");
				button_disable($("#btnChildAdd")[0], tablesSelected > 4);
				button_disable($("#btnChildEdit")[0], true);
				button_disable($("#btnChildRemove")[0], true);
				button_disable($("#btnChildRemoveAll")[0], tablesSelected == 0);

			},
			loadComplete: function(json) {
				// Highlight top row
				var ids = $(this).jqGrid("getDataIDs");
				if (ids && ids.length > 0)
					$(this).jqGrid("setSelection", ids[0]);

			}

		});
		$("#ChildTables").jqGrid('navGrid', '#pcrud', {});

	});

	
</script>