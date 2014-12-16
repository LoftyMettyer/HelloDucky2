@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Classes
@Imports DMI.NET.Code.Extensions
@Inherits System.Web.Mvc.WebViewPage(Of Models.CustomReportModel)

@Html.HiddenFor(Function(m) m.Parent1ViewAccess, New With {.class = "ViewAccess"})
@Html.HiddenFor(Function(m) m.Parent2ViewAccess, New With {.class = "ViewAccess"})
@Html.HiddenFor(Function(m) m.ChildTablesAvailable)
@Html.HiddenFor(Function(m) m.ChildTablesString, New With {.id = "txtCTAAS"})

<fieldset id="RelatedTableParent1" class="width45 floatleft" @Model.Parent1.Visibility>
	<legend class="fontsmalltitle">Parent 1 :</legend>

	<fieldset>
		<input type="hidden" id="txtParent1ID" name="Parent1.ID" value="@Model.Parent1.ID" />
		<div class="stretchyfixed">
			Table :
		</div>
		<div class="stretchyfill">
			@Html.TextBoxFor(Function(m) m.Parent1.Name, New With {.readonly = "true", .class = "width100"})
		</div>
	</fieldset>

	<fieldset>
		<div id="RelatedTableParent1AllRecordsDiv">
			@Html.RadioButton("Parent1.Selectiontype", RecordSelectionType.AllRecords, Model.Parent1.SelectionType = RecordSelectionType.AllRecords,
													New With {.id = "Parent1_SelectionTypeAll", .onclick = "changeRecordOption('Parent1','ALL')"})
			All Records
		</div>

		<div id="" class="tablerow">
			<div class="stretchyfixed">
				@Html.RadioButton("Parent1.SelectionType", RecordSelectionType.Picklist, Model.Parent1.SelectionType = RecordSelectionType.Picklist _
														, New With {.id = "Parent1_SelectionTypePicklist", .onclick = "changeRecordOption('Parent1','PICKLIST')"})
				Picklist
			</div>
			<div class="tablecell width100">
				<input type="hidden" id="txtParent1PicklistID" name="Parent1.PicklistID" value="@Model.Parent1.PicklistID" />
				@Html.TextBoxFor(Function(m) m.Parent1.PicklistName, New With {.id = "txtParent1Picklist", .readonly = "true", .class = "floatright width99"})
				@Html.ValidationMessageFor(Function(m) m.Parent1.PicklistID)
			</div>
			<div class="tablecell">
				@Html.EllipseButton("cmdParent1Picklist", "selectParent1Picklist()", Model.Parent1.SelectionType = RecordSelectionType.Picklist)
			</div>
		</div>

		<div id="" class="tablerow">
			<div class="stretchyfixed">
				@Html.RadioButton("Parent1.SelectionType", RecordSelectionType.Filter, Model.Parent1.SelectionType = RecordSelectionType.Filter _
														, New With {.id = "Parent1_SelectionTypeFilter", .onclick = "changeRecordOption('Parent1','FILTER')"})
				Filter
			</div>
			<div class="tablecell width100">
				<input type="hidden" id="txtParent1FilterID" name="Parent1.FilterID" value="@Model.Parent1.FilterID" />
				@Html.TextBoxFor(Function(m) m.Parent1.FilterName, New With {.id = "txtParent1Filter", .readonly = "true", .class = "floatright width99"})
				@Html.ValidationMessageFor(Function(m) m.Parent1.FilterID)

			</div>
			<div class="tablecell">
				@Html.EllipseButton("cmdParent1Filter", "selectParent1Filter()", Model.Parent1.SelectionType = RecordSelectionType.Filter)
			</div>
		</div>
	</fieldset>
</fieldset>

<fieldset id="RelatedTableParent2" class="width45 floatleft" @Model.Parent2.Visibility>
	<legend class="fontsmalltitle">Parent 2 :</legend>

	<fieldset>
		<input type="hidden" id="txtParent2ID" name="Parent2.ID" value="@Model.Parent2.ID" />
		<div class="stretchyfixed">
			Table :
		</div>
		<div class="stretchyfill">
			@Html.TextBoxFor(Function(m) m.Parent2.Name, New With {.readonly = "true", .class = "width100"})
		</div>
	</fieldset>

	<fieldset>
		<div id="RelatedTableParent2AllRecordsDiv">
			@Html.RadioButton("Parent2.Selectiontype", RecordSelectionType.AllRecords, Model.Parent2.SelectionType = RecordSelectionType.AllRecords,
													New With {.id = "Parent2_SelectionTypeAll", .onclick = "changeRecordOption('Parent2','ALL')"})
			All Records
		</div>

		<div id="" class="tablerow">
			<div class="stretchyfixed">
				@Html.RadioButton("Parent2.SelectionType", RecordSelectionType.Picklist, Model.Parent2.SelectionType = RecordSelectionType.Picklist,
														New With {.id = "Parent2_SelectionTypePicklist", .onclick = "changeRecordOption('Parent2','PICKLIST')"})
				Picklist
			</div>
			<div class="tablecell width100">
				<input type="hidden" id="txtParent2PicklistID" name="Parent2.PicklistID" value="@Model.Parent2.PicklistID" />
				@Html.TextBoxFor(Function(m) m.Parent2.PicklistName, New With {.id = "txtParent2Picklist", .readonly = "true", .class = "floatright width99"})
				@Html.ValidationMessageFor(Function(m) m.Parent2.PicklistID)
			</div>
			<div class="tablecell">
				@Html.EllipseButton("cmdParent2Picklist", "selectParent2Picklist()", Model.Parent2.SelectionType = RecordSelectionType.Picklist)
			</div>
		</div>

		<div id="" class="tablerow">
			<div class="stretchyfixed">
				@Html.RadioButton("Parent2.SelectionType", RecordSelectionType.Filter, Model.Parent2.SelectionType = RecordSelectionType.Filter,
											 New With {.id = "Parent2_SelectionTypeFilter", .onclick = "changeRecordOption('Parent2','FILTER')"})
				Filter
			</div>
			<div class="tablecell width100">
				<input type="hidden" id="txtParent2FilterID" name="Parent2.FilterID" value="@Model.Parent2.FilterID" />

				@Html.TextBoxFor(Function(m) m.Parent2.FilterName, New With {.id = "txtParent2Filter", .readonly = "true", .class = "floatright width99"})
				@Html.ValidationMessageFor(Function(m) m.Parent2.FilterID)
			</div>
			<div class="tablecell">
				@Html.EllipseButton("cmdParent2Filter", "selectParent2Filter()", Model.Parent2.SelectionType = RecordSelectionType.Filter)
			</div>
		</div>
	</fieldset>
</fieldset>

<br style="clear: left;" />

<div>
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
			<input type="button" id="btnChildRemove" value="Remove" disabled onclick="requestRemoveChildTable();" />
			<br />
			<input type="button" id="btnChildRemoveAll" value="Remove All" disabled onclick="requestRemoveAllChildTables();" />
		</div>
	</fieldset>
</div>


<script type="text/javascript">

	function selectParent1Picklist() {

		var tableID = $("#txtParent1ID").val();
		var currentID = $("#txtParent1PicklistID").val();
		var tableName = $("#Parent1_Name").val();

		OpenHR.modalExpressionSelect("PICKLIST", tableID, currentID, function (id, name, access) {
			if (access == "HD" && $("#Owner").val().toLowerCase() != '@Session("Username").ToString.ToLower') {
				$("#txtParent1PicklistID").val(0);
				$("#txtParent1Picklist").val('None');
				OpenHR.modalMessage("The " + tableName + " table picklist will be removed from this definition as it is hidden and you do not have permission to make this definition hidden.");
			}
			else {
				$("#txtParent1PicklistID").val(id);
				$("#txtParent1Picklist").val(name);
				setViewAccess('PICKLIST', $("#Parent1ViewAccess"), access, tableName);
				enableSaveButton();
			}
		}, 400, 400);

	}

	function selectParent2Picklist() {

		var tableID = $("#txtParent2ID").val();
		var currentID = $("#txtParent2PicklistID").val();
		var tableName = $("#Parent2_Name").val();

		OpenHR.modalExpressionSelect("PICKLIST", tableID, currentID, function (id, name, access) {
			if (access == "HD" && $("#Owner").val().toLowerCase() != '@Session("Username").ToString.ToLower') {
				$("#txtParent2PicklistID").val(0);
				$("#txtParent2Picklist").val('None');
				OpenHR.modalMessage("The " + tableName + " table picklist will be removed from this definition as it is hidden and you do not have permission to make this definition hidden.");
			}
			else {
				$("#txtParent2PicklistID").val(id);
				$("#txtParent2Picklist").val(name);
				setViewAccess('PICKLIST', $("#Parent2ViewAccess"), access, tableName);
				enableSaveButton();
			}
		}, 400, 400);

	}

	function selectParent1Filter() {

		var tableID = $("#txtParent1ID").val();
		var currentID = $("#txtParent1FilterID").val();
		var tableName = $("#Parent1_Name").val();

		OpenHR.modalExpressionSelect("FILTER", tableID, currentID, function (id, name, access) {
			if (access == "HD" && $("#Owner").val().toLowerCase() != '@Session("Username").ToString.ToLower') {
				$("#txtParent1FilterID").val(0);
				$("#txtParent1Filter").val('None');
				OpenHR.modalMessage("The " + tableName + " table filter will be removed from this definition as it is hidden and you do not have permission to make this definition hidden.");
			}
			else {
				$("#txtParent1FilterID").val(id);
				$("#txtParent1Filter").val(name);
				setViewAccess('FILTER', $("#Parent1ViewAccess"), access, tableName);
				enableSaveButton();
			}
		}, 400, 400);

	}

	function selectParent2Filter() {

		var tableID = $("#txtParent2ID").val();
		var currentID = $("#txtParent2FilterID").val();
		var tableName = $("#Parent2_Name").val();

		OpenHR.modalExpressionSelect("FILTER", tableID, currentID, function (id, name, access) {
			if (access == "HD" && $("#Owner").val().toLowerCase() != '@Session("Username").ToString.ToLower') {
				$("#txtParent2FilterID").val(0);
				$("#txtParent2Filter").val('None');
				OpenHR.modalMessage("The " + tableName + " table filter will be removed from this definition as it is hidden and you do not have permission to make this definition hidden.");
			}
			else {
				$("#txtParent2FilterID").val(id);
				$("#txtParent2Filter").val(name);
				setViewAccess('FILTER', $("#Parent2ViewAccess"), access, tableName);
				enableSaveButton();
			}
		}, 400, 400);

	}

	function addChildTable() {

		OpenHR.OpenDialog("Reports/AddChildTable", "divPopupReportDefinition", { ReportID: "@Model.ID" }, '800');

	}

	function editChildTable(rowID) {

		if (rowID == 0) {
			rowID = $('#ChildTables').jqGrid('getGridParam', 'selrow');
		}

		var gridData = $("#ChildTables").getRowData(rowID);
		OpenHR.OpenDialog("Reports/EditChildTable", "divPopupReportDefinition", gridData, '800');

	}

	function removeChildTableCompleted() {

		rowID = $('#ChildTables').jqGrid('getGridParam', 'selrow');
		var gridData = $("#ChildTables").getRowData(rowID);
		var columnList = $("#SelectedColumns").getDataIDs();
		var sortColumnList = $("#SortOrders").getDataIDs();

		$('#ChildTables').jqGrid('delRowData', rowID);
		loadAvailableTablesForReport(false);

		// Reset row selection for SelectedColumns and SortOrder grid
		$("#SelectedColumns").jqGrid('resetSelection');
		$("#SortOrders").jqGrid('resetSelection');

		for (i = 0; i < columnList.length; i++) {
			rowData = $("#SelectedColumns").getRowData(columnList[i]);

			// Remove all columns from selected columns grid whose table id is same as deleting table id
			if (rowData.TableID == gridData.TableID) {

				// Remove the matched sort columns from sort order
				for (j = 0; j < sortColumnList.length; j++) {
					var sortColumnRowId = sortColumnList[j];
					var dataRowOfSortColumn = $("#SortOrders").getRowData(sortColumnRowId);
					if (dataRowOfSortColumn.ColumnID == rowData.ID) {
						$("#SortOrders").delRowData(sortColumnRowId);
						break;
					}
				}

				// Remove the column from the selected column list
				$("#SelectedColumns").delRowData(columnList[i]);
			}
		}

		// Highlight top row of childTable grid, selectedColumns grid and sortedOrders grid
		selectGridTopRow($('#ChildTables'));
		selectGridTopRow($('#SelectedColumns'));
		selectGridTopRow($('#SortOrders'));

		checkIfDefinitionNeedsToBeHidden(0);
		enableSaveButton();
	}



	function requestRemoveAllChildTables() {

		OpenHR.modalPrompt("Removing all the child tables will remove all child table columns included in the report definition." +
		"<br/><br/>Are you sure you wish to continue ?", 4, "").then(function (answer) {
			if (answer == 6) { // Yes
				removeAllChildTables();
				//loadAvailableTablesForReport(true);
			}
		});

	}

	function requestRemoveChildTable() {
		rowID = $('#ChildTables').jqGrid('getGridParam', 'selrow');
		var gridData = $("#ChildTables").getRowData(rowID);
		var columnList = $("#SelectedColumns").getDataIDs();
		var iColumnCount = 0;

		for (i = 0; i < columnList.length; i++) {
			rowData = $("#SelectedColumns").getRowData(columnList[i]);
			if (rowData.TableID == gridData.TableID) {
				iColumnCount = iColumnCount + 1;
				break;
			}
		}

		if (iColumnCount > 0) {
			OpenHR.modalPrompt("One or more columns from '" + gridData.TableName + "' table have been included in the report definition." +
					"<br/><br/>Changing the child table will remove these columns from the report definition." +
					"<br/><br/>Are you sure you wish to continue ?", 4, "").then(function (answer) {
						if (answer == 6) { // Yes
							OpenHR.postData("Reports/RemoveChildTable", gridData, removeChildTableCompleted);
						}
					});
		}
		else {
			OpenHR.postData("Reports/RemoveChildTable", gridData, removeChildTableCompleted);
		}
	}

	function enableParent1RadioButtons()
	{
		$("#Parent1_SelectionTypeAll").removeAttr("disabled");
		$("#Parent1_SelectionTypePicklist").removeAttr("disabled");
		$("#Parent1_SelectionTypeFilter").removeAttr("disabled");
	}

	function disableParent1RadioButtons() {
		$("#Parent1_SelectionTypeAll").prop('disabled', "disabled");
		$("#Parent1_SelectionTypePicklist").prop('disabled', "disabled");
		$("#Parent1_SelectionTypeFilter").prop('disabled', "disabled");
	}

	function disableParent2RadioButtons() {
		$("#Parent2_SelectionTypeAll").prop('disabled', "disabled");
		$("#Parent2_SelectionTypePicklist").prop('disabled', "disabled");
		$("#Parent2_SelectionTypeFilter").prop('disabled', "disabled");
	}

	function enableParent2RadioButtons() {
		$("#Parent2_SelectionTypeAll").removeAttr("disabled");
		$("#Parent2_SelectionTypePicklist").removeAttr("disabled");
		$("#Parent2_SelectionTypeFilter").removeAttr("disabled");
	}

	function getColumnIndexByName(grid, columnName) {
		var cm = grid.jqGrid('getGridParam', 'colModel'), i, l;		
			for (i = 1, l = cm.length; i < l; i += 1) {
				if (cm[i].name === columnName) {
					return i; // return the index
				}
			}
			return -1;
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
				{ name: 'FilterViewAccess', index: 'Records', hidden: true, classes: 'ViewAccess' },
				{ name: 'OrderID', index: 'OrderID', width: 100, hidden: true },
				{ name: 'TableName', index: 'TableName', width: 100 },
				{ name: 'FilterName', index: 'FilterName', width: 100 },
				{ name: 'OrderName', index: 'OrderName', width: 100 },
			{ name: 'Records', index: 'Records', width: 100 }
			],
			autowidth: true,
			shrinkToFit: true,
			sortname: 'TableName',
			loadonce: true,
			viewrecords: true,
			sortorder: "asc",
			ondblClickRow: function (rowID) {
				editChildTable(rowID);
			},
			onSelectRow: function (id) {
				button_disable($("#btnChildEdit")[0], isDefinitionReadOnly());
				button_disable($("#btnChildRemove")[0], isDefinitionReadOnly());
			},
			gridComplete: function () {
				
				var tablesSelected = $(this).getGridParam("reccount");
				var tablesAvailable = $("#ChildTablesAvailable").val() - tablesSelected;

				button_disable($("#btnChildAdd")[0], tablesSelected > 4 || tablesAvailable == 0 || isDefinitionReadOnly());
				button_disable($("#btnChildEdit")[0], true);
				button_disable($("#btnChildRemove")[0], true);
				button_disable($("#btnChildRemoveAll")[0], tablesSelected == 0 || isDefinitionReadOnly());

				refreshViewAccess();
			},
			loadComplete: function(json) {
				
				// Highlight top row
				var ids = $(this).jqGrid("getDataIDs");
				if (ids && ids.length > 0)
					$(this).jqGrid("setSelection", ids[0]);

				var iCol = getColumnIndexByName($(this), 'Records'), rows = this.rows, i,	c = rows.length;
				for (i = 1; i < c; i += 1) {					
					if ($(rows[i].cells[iCol])[0].innerText == 0 ) {
						$(rows[i].cells[iCol])[0].innerText = "All Records"
					}
				}

			}
		});

		$("#ChildTables").jqGrid('navGrid', '#pcrud', {});

		//If Parent1 table does not exit then disabled the Parent1 radio buttons, because in IE browser the control shows as disabled but fires the click event while doing double click (i.e the fieldset control has disabled attribute in this scenario)
		if ('@Model.Parent1.Visibility' == 'disabled') {
			disableParent1RadioButtons();
		}

		//If Parent2 table does not exit then disabled the Parent2 radio buttons, because in IE browser the control shows as disabled but fires the click event while doing double click (i.e the fieldset control has disabled attribute in this scenario)
		if ('@Model.Parent2.Visibility' == 'disabled') {
			disableParent2RadioButtons();
		}

	});


</script>