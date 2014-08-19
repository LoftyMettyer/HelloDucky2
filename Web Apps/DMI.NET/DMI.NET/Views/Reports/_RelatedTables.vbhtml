@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Classes
@Imports DMI.NET.Code.Extensions
@Inherits System.Web.Mvc.WebViewPage(Of Models.CustomReportModel)


<fieldset id="divReportParents">

	@Html.HiddenFor(Function(m) m.Parent1ViewAccess)
	@Html.HiddenFor(Function(m) m.Parent2ViewAccess)

	<fieldset id="RelatedTableParent1" class="width45 floatleft" @Model.Parent1.Visibility>
		<legend class="fontsmalltitle">Parent 1 :</legend>

		<fieldset>
			@Html.HiddenFor(Function(m) m.ChildTablesString, New With {.id = "txtCTAAS"})
			<input type="hidden" id="txtParent1ID" name="Parent1.ID" value="@Model.Parent1.ID" />
			<div class="width30 floatleft">
				Table:
			</div>
			<div class="width70 floatleft">
				@Html.TextBoxFor(Function(m) m.Parent1.Name, New With {.readonly = "true", .style = "width:100%"})
			</div>
		</fieldset>

		<fieldset>
			<div id="RelatedTableParent1AllRecordsDiv">
				@Html.RadioButton("Parent1.Selectiontype", RecordSelectionType.AllRecords, Model.Parent1.SelectionType = RecordSelectionType.AllRecords,
													New With {.id = "Parent1_SelectionTypeAll", .onclick = "changeRecordOption('Parent1','ALL')"})
				All Records
			</div>

			<div id="RelatedTablesParent1PicklistDiv">
				<div class="width30 floatleft">
					@Html.RadioButton("Parent1.SelectionType", RecordSelectionType.Picklist, Model.Parent1.SelectionType = RecordSelectionType.Picklist _
														, New With {.id = "Parent1_SelectionTypePicklist", .onclick = "changeRecordOption('Parent1','PICKLIST')"})
					Picklist
				</div>
				<div class="floatleft">
					<input type="hidden" id="txtParent1PicklistID" name="Parent1.PicklistID" value="@Model.Parent1.PicklistID" />
					@Html.TextBoxFor(Function(m) m.Parent1.PicklistName, New With {.id = "txtParent1Picklist", .readonly = "true"})
					@Html.ValidationMessageFor(Function(m) m.Parent1.PicklistID)
				</div>
				<div class="floatleft">
					@Html.EllipseButton("cmdParent1Picklist", "selectParent1Picklist()", Model.Parent1.SelectionType = RecordSelectionType.Picklist)
				</div>
			</div>

			<div class="width100 clearboth">
				<div class="width30 floatleft">
					@Html.RadioButton("Parent1.SelectionType", RecordSelectionType.Filter, Model.Parent1.SelectionType = RecordSelectionType.Filter _
														, New With {.id = "Parent1_SelectionTypeFilter", .onclick = "changeRecordOption('Parent1','FILTER')"})
					Filter
				</div>
				<div class="floatleft">
					<input type="hidden" id="txtParent1FilterID" name="Parent1.FilterID" value="@Model.Parent1.FilterID" />
					@Html.TextBoxFor(Function(m) m.Parent1.FilterName, New With {.id = "txtParent1Filter", .readonly = "true"})
					@Html.ValidationMessageFor(Function(m) m.Parent1.FilterID)
				</div>
				<div class="floatleft">
					@Html.EllipseButton("cmdParent1Filter", "selectParent1Filter()", Model.Parent1.SelectionType = RecordSelectionType.Filter)
				</div>
			</div>
		</fieldset>
	</fieldset>

	<fieldset id="RelatedTableParent2" class="width45 floatleft" @Model.Parent2.Visibility>
		<legend class="fontsmalltitle">Parent 2 :</legend>

		<fieldset>
			<input type="hidden" id="txtParent2ID" name="Parent2.ID" value="@Model.Parent2.ID" />
			<div class="width30 floatleft">
				Table:
			</div>
			<div class="width70 floatleft">
				@Html.TextBoxFor(Function(m) m.Parent2.Name, New With {.readonly = "true", .style = "width:100%"})
			</div>
		</fieldset>

		<fieldset>
			<div id="RelatedTableParent2AllRecordsDiv">
				@Html.RadioButton("Parent2.Selectiontype", RecordSelectionType.AllRecords, Model.Parent2.SelectionType = RecordSelectionType.AllRecords,
													New With {.id = "Parent2_SelectionTypeAll", .onclick = "changeRecordOption('Parent2','ALL')"})
				All Records
			</div>

			<div id="RelatedTablesParent2PicklistDiv">
				<div class="width30 floatleft">
					@Html.RadioButton("Parent2.SelectionType", RecordSelectionType.Picklist, Model.Parent2.SelectionType = RecordSelectionType.Picklist, New With {.id = "Parent2_SelectionTypePicklist", .onclick = "changeRecordOption('Parent2','PICKLIST')"})
					Picklist
				</div>
				<div class="floatleft">
					<input type="hidden" id="txtParent2PicklistID" name="Parent2.PicklistID" value="@Model.Parent2.PicklistID" />
					@Html.TextBoxFor(Function(m) m.Parent2.PicklistName, New With {.id = "txtParent2Picklist", .readonly = "true"})
					@Html.ValidationMessageFor(Function(m) m.Parent2.PicklistID)
				</div>
				<div class="floatleft">
					@Html.EllipseButton("cmdParent2Picklist", "selectParent2Picklist()", Model.Parent2.SelectionType = RecordSelectionType.Picklist)
				</div>
			</div>

			<div class="clearboth">
				<div class="width30 floatleft">
					@Html.RadioButton("Parent2.SelectionType", RecordSelectionType.Filter, Model.Parent2.SelectionType = RecordSelectionType.Filter,
											 New With {.id = "Parent2_SelectionTypeFilter", .onclick = "changeRecordOption('Parent2','FILTER')"})
					Filter
				</div>
				<div class="floatleft">
					<input type="hidden" id="txtParent2FilterID" name="Parent2.FilterID" value="@Model.Parent2.FilterID" />

					@Html.TextBoxFor(Function(m) m.Parent2.FilterName, New With {.id = "txtParent2Filter", .readonly = "true"})
					@Html.ValidationMessageFor(Function(m) m.Parent2.FilterID)
				</div>
				<div class="floatleft">
					@Html.EllipseButton("cmdParent2Filter", "selectParent2Filter()", Model.Parent2.SelectionType = RecordSelectionType.Filter)
				</div>
			</div>
		</fieldset>
	</fieldset>
	</fieldset>

	<br />

<div class="absolutefullchildtables">
<fieldset class="relatedtables width100">
	<legend class="fontsmalltitle">Child Tables :</legend>

	<div id="ChildTablesViewAccessdiv" class="width80 floatleft" style="">
		<input type="hidden" id="ChildTablesViewAccess" />
		<table id="ChildTables"></table>
	</div>

	<div class="stretchyfixed" style="padding-left:15px">
		<input type="button" id="btnChildAdd" value="Add..." onclick="addChildTable();" />
			<br />
		<input type="button" id="btnChildEdit" value="Edit..." disabled onclick="editChildTable(0);" />
		<br />
		<input type="button" id="btnChildRemove" value="Remove" disabled onclick="removeChildTable();" />
		<br />
		<input type="button" id="btnChildRemoveAll" value="Remove All" disabled onclick="removeAllChildTables();" />				
	</div>
</fieldset>
</div>


<script type="text/javascript">

	function selectParent1Picklist() {

		var tableID = $("#txtParent1ID").val();
		var currentID = $("#txtParent1PicklistID").val();

		OpenHR.modalExpressionSelect("PICKLIST", tableID, currentID, function (id, name, access) {
			$("#txtParent1PicklistID").val(id);
			$("#txtParent1Picklist").val(name);
			setViewAccess('PICKLIST', $("#Parent1ViewAccess"), access);
		}, 400, 400);

	}

	function selectParent2Picklist() {

		var tableID = $("#txtParent2ID").val();
		var currentID = $("#txtParent2PicklistID").val();

		OpenHR.modalExpressionSelect("PICKLIST", tableID, currentID, function (id, name, access) {
			$("#txtParent2PicklistID").val(id);
			$("#txtParent2Picklist").val(name);
			setViewAccess('PICKLIST', $("#Parent1ViewAccess"), access);
		}, 400, 400);

	}

	function selectParent1Filter() {

		var tableID = $("#txtParent1ID").val();
		var currentID = $("#txtParent1FilterID").val();

		OpenHR.modalExpressionSelect("FILTER", tableID, currentID, function (id, name, access) {
			$("#txtParent1FilterID").val(id);
			$("#txtParent1Filter").val(name);
			setViewAccess('FILTER', $("#Parent1ViewAccess"), access);
		}, 400, 400);

	}

	function selectParent2Filter() {

		var tableID = $("#txtParent2ID").val();
		var currentID = $("#txtParent2FilterID").val();

		OpenHR.modalExpressionSelect("FILTER", tableID, currentID, function (id, name, access) {
			$("#txtParent2FilterID").val(id);
			$("#txtParent2Filter").val(name);
			setViewAccess('FILTER', $("#Parent2ViewAccess"), access);
		}, 400, 400);

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
		loadAvailableTablesForReport(false);
	}

	$(function () {

		$("#ChildTables").jqGrid('setGridWidth', $("#ChildTablesViewAccessdiv").width() - 50);

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
			colNames: ['ID', 'ReportID', 'ReportType', 'TableID', 'FilterID', 'FilterViewAccess', 'OrderID', 'Table', 'Filter', 'Order', 'Records'],
			colModel: [
				{ name: 'ID', index: 'ID', sorttype: 'int', hidden: true },
				{ name: 'ReportID', index: 'ReportID', sorttype: 'int', hidden: true },
				{ name: 'ReportType', index: 'ReportType', sorttype: 'int', hidden: true },
				{ name: 'TableID', index: 'TableID', width: 100, hidden: true },
				{ name: 'FilterID', index: 'FilterID', width: 100, hidden: true },
				{ name: 'FilterViewAccess', index: 'Records', hidden: true },
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
			gridComplete: function () {
				var tablesSelected = $(this).getGridParam("reccount");
				button_disable($("#btnChildAdd")[0], tablesSelected > 4 || '@Model.ChildTablesAvailable' == 'False');
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