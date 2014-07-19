@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.Enums
@imports HR.Intranet.Server.Enums
@Imports DMI.NET.ViewModels.Reports
@Imports DMI.NET.Code.Extensions
@Imports DMI.NET.Models
@Inherits System.Web.Mvc.WebViewPage(Of ReportBaseModel)

<style>

	#columnsAvailable {
		height: 500px;
		overflow: auto;
	}

</style>

@Html.HiddenFor(Function(m) m.ColumnsAsString, New With {.id = "txtCSAAS"})


<div id="columnsAvailable" style="float:left">

	Columns / Calculations Available :
	<br/>

	<select name="SelectedTableID" id="SelectedTableID" onchange="getAvailableTableColumnsCalcs();"></select>

	<br />
	@Html.RadioButton("columnSelectiontype", ColumnSelectionType.Columns, True, New With {.onclick = "toggleColumnsCalculations('column')"})
	Columns
	@Html.RadioButton("columnSelectiontype", ColumnSelectionType.Calculations, True, New With {.onclick = "toggleColumnsCalculations('calc')"})
	Calculations
	<br/>

	<table id="AvailableColumns" class="scroll" cellpadding="0" cellspacing="0"></table>
</div>

<div style="float:left">
	<input type="button" id="btnColumnAdd" value="Add" onclick="addColumnToSelected(0);" />
	<br />
	<input type="button" id="btnColumnAddAll" value="Add All" onclick="addAllColumnsToSelected();" />
	<br />
	<input type="button" id="btnColumnRemove" value="Remove" onclick="removeSelectedColumn();" />
	<br />
	<input type="button" id="btnColumnRemoveAll" value="Remove All" onclick="removeAllSelectedColumns();" />
	<br />
	<input type="button" id="btnColumnMoveUp" value="Move Up" onclick="moveUpSelectedColumn();" />
	<br />
	<input type="button" id="btnColumnMoveDown" value="Move Down" onclick="moveDownSelectedColumn();" />
</div>

<div id="columnsSelected" style="float:left">
	<table id="SelectedColumns" class="scroll" cellpadding="0" cellspacing="0"></table>
	<br/>

	<div id="columnsSelectedDrilldown">
		Heading:
		<input type='text' id="SelectedColumnHeading" onblur="updateColumnsSelectedGrid();" />
		<br />

		Size:
		<input type='text' id="SelectedColumnSize" onblur="updateColumnsSelectedGrid();" />
		Decimals:
		<input type='text' id="SelectedColumnDecimals" onblur="updateColumnsSelectedGrid();" />
		<br />

		Average:
		<input type='checkbox' id="SelectedColumnIsAverage" onblur="updateColumnsSelectedGrid();" />
		Count:
		<input type='checkbox' id="SelectedColumnIsCount" onblur="updateColumnsSelectedGrid();" />
		Total:
		<input type='checkbox' id="SelectedColumnIsTotal" onblur="updateColumnsSelectedGrid();" />
		<br />

		Hidden:
		<input type='checkbox' id="SelectedColumnIsHidden" onblur="updateColumnsSelectedGrid();" />
		Group With Next:
		<input type='checkbox' id="SelectedColumnIsGroupWithNext" onblur="updateColumnsSelectedGrid();" />
	</div>

</div>

<input type="hidden" name="Columns.BaseTableID" value="@Model.BaseTableID" />


  <script type="text/javascript">

		function moveUpSelectedColumn() {

		}

		function moveDownSelectedColumn() {

		}

		function toggleColumnsCalculations(type) {
		}

		function addColumnToSelected(rowID) {

			if (rowID == 0) {
				rowID = $("#AvailableColumns").getGridParam('selrow');
			}

			var datarow = $("#AvailableColumns").getRowData(rowID);
	
			datarow.IsExpression = false;
			datarow.Heading = datarow.Name;
			datarow.Sequence = $("#SelectedColumns").jqGrid('getGridParam', 'records') + 1;
			datarow.IsAverage = false;
			datarow.IsCount = false;
			datarow.IsTotal = false;
			datarow.IsHidden = false;
			datarow.IsGroupWithNext = false;
	
			$("#SelectedColumns").jqGrid('addRowData', datarow.ID, datarow);
			$("#AvailableColumns").jqGrid('delRowData', rowID);

		}

		function addAllColumnsToSelected() {

			var rows = $("#AvailableColumns").jqGrid('getDataIDs');

			for (var i = 0; i < rows.length; i++) {
				debugger;
				var datarow = $("#AvailableColumns").getRowData(rows[i]);
		//		$("#SelectedColumns").jqGrid('addRowData', datarow.ID, datarow);
				addColumnToSelected(datarow.ID);
			}

//			$("#SelectedColumns").jqGrid('clearGridData')

		}

		function removeSelectedColumn() {

			rowID = $("#SelectedColumns").getGridParam('selrow');
			var datarow = $("#SelectedColumns").getRowData(rowID);

			$("#AvailableColumns").jqGrid('addRowData', datarow.ID, datarow);
			$("#AvailableColumns").jqGrid("sortGrid", "Name", true)
			$("#SelectedColumns").jqGrid('delRowData', rowID);

		}

    function columndefinition_rowcolchange() {

    	iRowID = $("#ColumnsSelected").getGridParam('selrow') - 1;
    	$("[id^=columnproperty], [id$=Breakdown]").hide();
    	$("#columnproperty" + iRowID + "Breakdown").show();

    }

    function getAvailableTableColumnsCalcs() {

    	$("#AvailableColumns").jqGrid('GridUnload');

    	$("#AvailableColumns").jqGrid({
    		url: 'Reports/GetColumnsForTable?TableID=' + $("#SelectedTableID").val(),
    		datatype: 'json',
    		mtype: 'GET',
    		jsonReader: {
    			root: "rows", //array containing actual data
    			page: "page", //current page
    			total: "total", //total pages for the query
    			records: "records", //total number of records
    			repeatitems: false,
    			id: "ID" //index of the column with the PK in it
    		},
    		colNames: ['ID', 'Name', 'DataType', 'Size', 'Decimals'],
    		colModel: [
					{ name: 'ID', index: 'ID', hidden: true },
					{ name: 'Name', index: 'Name', width: 40, sortable: false },
    			{ name: 'DataType', index: 'DataType', hidden: true },
					{ name: 'Size', index: 'Size',  hidden: true },
					{ name: 'Decimals', index: 'Decimals', hidden: true }],
    		viewrecords: true,
    		width: 400,
    		sortname: 'Name',
    		sortorder: "desc",
    		rowNum: '',
    		ondblClickRow: function (rowid) {
    			addColumnToSelected(rowid);
    		}
    	});
    	
    	// TODO Loop through available removing any currently selected

    }

    function updateColumnsSelectedGrid() {

    	var rowId = $("#SelectedColumns").jqGrid('getGridParam', 'selrow');
    	var dataRow = $('#SelectedColumns').jqGrid('getRowData', rowId);

    	dataRow.Heading = $("#SelectedColumnHeading").val();
    	dataRow.Size = $("#SelectedColumnSize").val();
    	dataRow.Decimals = $("#SelectedColumnDecimals").val();
    	dataRow.IsAverage = $('#SelectedColumnIsAverage').is(':checked');
    	dataRow.IsCount = $("#SelectedColumnIsCount").is(':checked');
    	dataRow.IsTotal = $("#SelectedColumnIsTotal").is(':checked');
    	dataRow.IsHidden = $("#SelectedColumnIsHidden").is(':checked');
    	dataRow.IsGroupWithNext = $("#SelectedColumnIsGroupWithNext").is(':checked');
	
    	$('#SelectedColumns').jqGrid('setRowData', rowId, dataRow);

    }

    function attachGridToSelectedColumns() {

    	$("#SelectedColumns").jqGrid({
    		datatype: "jsonstring",
    		datastr: '@Model.Columns.ToJsonResult',
    		mtype: 'GET',
    		jsonReader: {
    			root: "rows", //array containing actual data
    			page: "page", //current page
    			total: "total", //total pages for the query
    			records: "records", //total number of records
    			repeatitems: false,
    			id: "ID" //index of the column with the PK in it
    		},
    		colNames: ['ID', 'IsExpression',	'Name',	'Sequence',	'Heading',	'DataType',
    							'Size',	'Decimals',	'IsAverage',	'IsCount',	'IsTotal',	'IsHidden',	'IsGroupWithNext'],
    		colModel: [
					{ name: 'ID', index: 'ID', hidden: true },
					{ name: 'IsExpression', index: 'IsExpression', hidden: true },
					{ name: 'Name', index: 'Name' },
					{ name: 'Sequence', index: 'Sequence', hidden: false },
					{ name: 'Heading', index: 'Heading', hidden: true },
					{ name: 'DataType', index: 'DataType', hidden: true },
					{ name: 'Size', index: 'Size', hidden: true },
					{ name: 'Decimals', index: 'Decimals', hidden: true },
					{ name: 'IsAverage', index: 'IsAverage', hidden: true },
					{ name: 'IsCount', index: 'IsCount', hidden: true },
					{ name: 'IsTotal', index: 'IsTotal', hidden: true },
					{ name: 'IsHidden', index: 'IsHidden', hidden: true },
					{ name: 'IsGroupWithNext', index: 'IsGroupWithNext', hidden: true }],
    		viewrecords: true,
    		width: 400,
    		sortname: 'Sequence',
    		sortorder: "asc",
    		rowNum: '',
    		onSelectRow: function (id) {

    			var rowId = $("#SelectedColumns").jqGrid('getGridParam', 'selrow');
    			var dataRow = $("#SelectedColumns").getRowData(rowId)

    			$("#SelectedColumnHeading").val(dataRow.Heading);
    			$("#SelectedColumnSize").val(dataRow.Size);
    			$("#SelectedColumnDecimals").val(dataRow.Decimals);
    			$('#SelectedColumnIsAverage').prop('checked', JSON.parse(dataRow.IsAverage))
    			$('#SelectedColumnIsCount').prop('checked', JSON.parse(dataRow.IsCount))
    			$('#SelectedColumnIsTotal').prop('checked', JSON.parse(dataRow.IsTotal))
    			$('#SelectedColumnIsHidden').prop('checked', JSON.parse(dataRow.IsHidden))
    			$('#SelectedColumnIsGroupWithNext').prop('checked', JSON.parse(dataRow.IsGroupWithNext))

    		},
    		gridComplete: function () {
					// Highlight top row
    			var ids = $(this).jqGrid("getDataIDs");
    			if (ids && ids.length > 0)
    				$(this).jqGrid("setSelection", ids[0]);
    		}
    	});



    }



		// Initialise
    $(function () {

    	attachGridToSelectedColumns();

    });

  </script>
