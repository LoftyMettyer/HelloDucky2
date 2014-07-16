@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Enums
@Imports DMI.NET.Classes
@Imports DMI.NET.Code.Extensions
@Inherits System.Web.Mvc.WebViewPage(Of Models.CustomReportModel)


<div id="divReportParents">

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
		@Html.EllipseButton("cmdParent1Picklist", "selectRecordOption('p1', 'picklist')", Model.Parent1.SelectionType = RecordSelectionType.Picklist)
		@Html.ValidationMessageFor(Function(m) m.Parent1.PicklistID)
		<br />

		@Html.RadioButton("Parent1.SelectionType", RecordSelectionType.Filter, Model.Parent1.SelectionType = RecordSelectionType.Filter, New With {.onclick = "changeRecordOption('Parent1','filter')"})
		Filter
		<input type="hidden" id="txtParent1FilterID" name="Parent1.FilterID" value="@Model.Parent1.FilterID" />
		@Html.TextBoxFor(Function(m) m.Parent1.FilterName, New With {.id = "txtParent1Filter", .readonly = "true"})
		@Html.EllipseButton("cmdParent1Filter", "selectRecordOption('p1', 'filter')", Model.Parent1.SelectionType = RecordSelectionType.Filter)
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
		@Html.EllipseButton("cmdParent2Picklist", "selectRecordOption('p2', 'picklist')", Model.Parent2.SelectionType = RecordSelectionType.Picklist)
		@Html.ValidationMessageFor(Function(m) m.Parent2.PicklistID)
		<br />

		@Html.RadioButton("Parent2.SelectionType", RecordSelectionType.Filter, Model.Parent2.SelectionType = RecordSelectionType.Filter, New With {.onclick = "changeRecordOption('Parent2','filter')"})
		Filter
		<input type="hidden" id="txtParent2FilterID" name="Parent2.FilterID" value="@Model.Parent2.FilterID" />
		@Html.TextBoxFor(Function(m) m.Parent2.FilterName, New With {.id = "txtParent2Filter", .readonly = "true"})
		@Html.EllipseButton("cmdParent2Filter", "selectRecordOption('p2', 'filter')", Model.Parent2.SelectionType = RecordSelectionType.Filter)
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
		<input type="button" id="btnChildEdit" value="Edit..." onclick="editChildTable(0);" />
		<br />
		<input type="button" id="btnChildRemove" value="Remove" onclick="removeChildTable();" />
		<br />
		<input type="button" id="btnChildRemoveAll" value="Remove All" onclick="removeAllChildTables();" />				
	</div>

</fieldset>



<script type="text/javascript">


	function toggle_visibility(id) {

		var e = $("#Event_Detail_" + id)[0]
		if (e.style.display == 'block')
			e.style.display = 'none';
		else
			e.style.display = 'block';
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
		$('#ChildTables').jqGrid('delRowData', rowid)



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
				id: "TableID" //index of the column with the PK in it
			},
			colNames: ['ReportID', 'TableID', 'FilterID', 'OrderID', 'Table', 'Filter', 'Order', 'Records'],
			colModel: [
				{ name: 'ReportID', index: 'reportID', sorttype: 'int', hidden: true },
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
			sortorder: "desc",
			editurl: 'server.php', // this is dummy existing url
			ondblClickRow: function (rowID) {
				editChildTable(rowID);
			},
		});
		jQuery("#ChildTables").jqGrid('navGrid', '#pcrud', {});

	});

	
</script>