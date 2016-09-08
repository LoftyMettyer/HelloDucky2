@Imports DMI.NET
@Imports DMI.NET.Helpers
@Imports DMI.NET.ViewModels.Reports
@Imports DMI.NET.Code.Extensions
@Imports DMI.NET.Models
@Inherits System.Web.Mvc.WebViewPage(Of OrganisationReportModel)

@Html.HiddenFor(Function(m) m.ColumnsAsString, New With {.id = "txtCSAAS"})
@Html.HiddenFor(Function(m) m.DefinitionAccessBasedOnSelectedCalculationColumns, New With {.class = "ViewAccess"})

<div class="nowrap">
   <div class="tablerow">
      <fieldset id="selectedTable">
         <legend class="fontsmalltitle width100">View/Table(s) :</legend>
         @Html.DropDownListFor(Function(m) m.BaseViewID, New SelectList(Model.AllAvailableViewList, "Id", "Name"), New With {.class = "width70 floatright", .id = "SelectedTableID", .name = "SelectedTableID", .onchange = "ChangeColumnTableView(event.target);"})
      </fieldset>
   </div>

   <div class="tablerow coldefinition">
      <div class="tablecell">
         <fieldset id="columnsAvailable">
            <legend class="fontsmalltitle">Columns Available :</legend>
            <table id="AvailableColumns"></table>
         </fieldset>
      </div>

      <div class="tablecell">
         <fieldset class="" id="columnbuttons">
            <div id="colbtngrp1">
               <input type="button" id="btnColumnAdd" class="enableSaveButtonOnClick" value="Add" onclick="addColumnToSelected();" />
               <input type="button" id="btnColumnAddAll" class="enableSaveButtonOnClick" value="Add All" onclick="addAllColumnsToSelected();" />
            </div>
            <div id="colbtngrp2">
               <input type="button" id="btnColumnRemove" value="Remove" onclick="requestRemoveSelectedColumns();" />
               <input type="button" id="btnColumnRemoveAll" class="enableSaveButtonOnClick" value="Remove All" onclick="requestRemoveAllSelectedColumns();" />
            </div>
            <div id="colbtngrp3" class="customReportsOnly">
               <input type="button" id="btnColumnMoveUp" class="enableSaveButtonOnClick" value="Move Up" disabled onclick="moveSelectedColumn('up');" />
               <input type="button" id="btnColumnMoveDown" class="enableSaveButtonOnClick" value="Move Down" disabled onclick="moveSelectedColumn('down');" />
            </div>
         </fieldset>
      </div>

      <div class="tablecell">
         <fieldset class="left" id="columnsSelected">
            <legend class="fontsmalltitle">Columns Selected :</legend>
            <table id="SelectedColumns" class="scroll" cellpadding="0" cellspacing="0"></table>
         </fieldset>
      </div>
   </div>
   <div class="tablerow coldefinition">
      <div class="tablecell">
         <fieldset>
            <div id="OrganisationPreview" class="OrganisationReportsOnly">
               <br />
               <input type="button" id="btnOrgPreview" class="enableSaveButtonOnClick" value="Preview" onclick="ShowPreview();" />
            </div>
         </fieldset>
      </div>
      <div class="tablecell"></div>
      <div class="tablecell">
         <fieldset>
            <div id="definitionColumnProperties">
               <div class="formfieldfill OrgReportsOnly">
                  <label for="SelectedColumnPrefix">Prefix :</label>
                  <span><input type='text' id="SelectedColumnPrefix" maxlength="20" onchange="updateColumnsSelectedGrid();" /></span>
               </div>
               <div class="formfieldfill OrgReportsOnly">
                  <label for="SelectedColumnSuffix">Suffix :</label>
                  <span><input type='text' id="SelectedColumnSuffix" maxlength="20" onchange="updateColumnsSelectedGrid();" /></span>
               </div>
               <div class="formfieldfill">
                  <label class="fontsizeOnly" for="SelectedColumnFontSize">Font Size :</label>
                  <span><input class="fontsizeOnly" type="tel" id="SelectedColumnFontSize" maxlength="2" onchange="updateColumnsSelectedGrid();" /></span>
                  <div class="formfieldfill HeightOnly">
                     <label for="SelectedColumnHeight"> &nbsp; Height (Rows) :</label>
                     <span><input type="tel" class="selectHeight" id="SelectedColumnHeight" style="width:94%" maxlength="1" onchange="updateColumnsSelectedGrid();" /></span>
                  </div>
               </div>
               <div class="formfieldfill decimalsOnly">
                  <label for="SelectedColumnDecimals">Decimals :</label>
                  <span><input class="selectFullText" id="SelectedColumnDecimals" onchange="updateColumnsSelectedGrid();" /></span>
               </div>

               <div class="tablelayout customReportsOnly colAggregates">
                  <div class="tablerow">
                     <div class="tablecell canGroupWithNext" style="color: rgb(0, 0, 0);">
                        <input class="ui-widget ui-corner-all" id="SelectedColumnIsConcatenateWithNext" onchange="changeColumnIsConcatenateWithNext();" type="checkbox">
                        <label id="labelSelectedColumnIsConcatenateWithNext" for="SelectedColumnIsConcatenateWithNext">Concatenate with next</label>
                     </div>
                  </div>
               </div>
            </div>
         </fieldset>
      </div>
   </div>
</div>

<input type='hidden' id="SelectedColumnHeading" maxlength="50" />
<input type="hidden" name="Columns.BaseTableID" value="@Model.BaseTableID" />
<input type="hidden" name="Columns.BaseViewID" value="@Model.BaseViewID" />
<input type="hidden" id="PostBasedTableId" value="@Model.PostBasedTableId" />
<input type="hidden" id="PostBasedTableName" value="@Model.PostBasedTableName" />

<script type="text/javascript">

   $('#divPopupPreview').dialog({
      overflow: true,
      autoOpen: false,
      width: 'auto',
      height: 'auto',
      resizable: false,
      modal: true,
      title: 'Organisation Report Preview'
   });

   function ShowPreview() {
      var gridData = $('#SelectedColumns').jqGrid('getRowData');
      var postData = {
         ReportID: "@Model.ID",
         GridData: JSON.stringify(gridData),
         __RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
      };

      $.ajax({
         type: "POST",
         url: "Reports/ValidateSelectedColumn",
         data: postData,
         success: function (result) {
            OpenHR.OpenDialog("Reports/showPreviewPopup", "divPopupPreview", postData, '400', '400');
         },
         error: function (xhr, status, error) {
            var err = "Error " + " " + status + " " + error;
            if (xhr.responseText && xhr.responseText[0] == "{")
               err = JSON.parse(xhr.responseText).ErrorMessage;

            OpenHR.modalMessage(err);
         }
      });
   }

   function removeSelectedColumnsFromAvailable() {
      //Find row in Sort Order columns to see if Value On Change or Suppress Repeated Values is ticked.
      var SelectedRows = $("#SelectedColumns").getRowData();
      var AvailableRows = $("#AvailableColumns").getRowData();

      for (i = 0; i < AvailableRows.length; i++) {
         for (x = 0; x < SelectedRows.length; x++) {
            if (AvailableRows[i].ID == SelectedRows[x].ID) {
               $("#AvailableColumns").delRowData(AvailableRows[i].ID);
            }
         }
      }
   }

   function moveSelectedColumn(direction) {
      OpenHR.MoveItemInGrid($("#SelectedColumns"), direction);
      var rowId = $("#SelectedColumns").jqGrid('getGridParam', 'selrow');

      var allRows = $('#SelectedColumns').jqGrid('getDataIDs');
      var isBottomRow = (rowId == allRows[allRows.length - 1]);
      if (isBottomRow) {
         $('#SelectedColumnIsConcatenateWithNext').prop('checked', false);
         updateColumnsSelectedGrid();
      }
      ClearConcatenateCheckbox();

   }

   function ChangeColumnTableView(target) {
      getAvailableTableViewColumns();
   }

   function addColumnToSelected() {

      var rowID;

      $('#SelectedColumns').jqGrid('resetSelection');

      var selectedRows = $('#AvailableColumns').jqGrid('getGridParam', 'selarrrow');

      for (var i = 0; i <= selectedRows.length - 1; i++) {
         rowID = selectedRows[i];
         var datarow = getDatarowFromAvailable(selectedRows[i]);

         datarow["__RequestVerificationToken"] = $('[name="__RequestVerificationToken"]').val();
         OpenHR.postData("Reports/AddOrganisationReportColumn", datarow);

         $("#SelectedColumns").jqGrid('addRowData', datarow.ID, datarow);
         $('#SelectedColumns').jqGrid("setSelection", rowID);

      }

      var ids = $("#AvailableColumns").getDataIDs();
      var nextIndex = $("#AvailableColumns").getInd(rowID);

      // Position next selected column
      var recordCount = $("#AvailableColumns").jqGrid('getGridParam', 'records')
      if (nextIndex >= recordCount) { nextIndex = 0; }

      // Remove selected columns from available
      for (var i = selectedRows.length - 1; i >= 0; i--) {
         $("#AvailableColumns").delRowData(selectedRows[i]);
      }

      $("#AvailableColumns").jqGrid("setSelection", ids[nextIndex], true);
      refreshcolumnPropertiesPanel();
   }

   function getDatarowFromAvailable(index) {

      var datarow = $("#AvailableColumns").getRowData(index);
      datarow.Heading = datarow.Name;
      datarow.Name = $("#SelectedTableID option:selected").text() + '.' + datarow.Name;
      datarow.ReportType = '@Model.ReportType';
      datarow.ReportID = '@Model.ID';

      datarow.ColumnID = datarow.ID;
      datarow.FontSize = 11;

      if (datarow.DataType == -3) {
         datarow.Height = 3
      }
      else {
         datarow.Height = 1;
      }

      datarow.IsGroupWithNext = false;
      var bIsTable = false;
      var ViewID = $("#SelectedTableID").val();

      //For storing TableID and ViewID
      if (($("#PostBasedTableId").val() != undefined &&
          $("#PostBasedTableId").val() == ViewID) &&
          ($("#PostBasedTableName").val() != undefined &&
          $("#PostBasedTableName").val() == $("#SelectedTableID option:selected")[0].text)) {
         bIsTable = true;
      }

      if (bIsTable) {
         datarow.TableID = $("#SelectedTableID option:selected").val();
         datarow.ViewID = 0;
      }
      else {
         datarow.ViewID = $("#SelectedTableID option:selected").val();
         datarow.TableID = 0;
      }

      return datarow;
   }

   function addAllColumnsToSelected() {

      var allRows = $('#AvailableColumns').jqGrid('getDataIDs');

      var bIsTable = 'false';
      var ViewID = $("#SelectedTableID").val();

      if (($("#PostBasedTableId").val() != undefined &&
     $("#PostBasedTableId").val() == ViewID) &&
     ($("#PostBasedTableName").val() != undefined &&
     $("#PostBasedTableName").val() == $("#SelectedTableID option:selected")[0].text)) {
         bIsTable = true;
      }

      var postData = {
         ReportID: '@Model.ID',
         ReportType: '@Model.ReportType',
         SelectionType: 'C',
         ColumnsTableID: $("#SelectedTableID").val(),
         TableName: $("#SelectedTableID option:selected").text(),
         Columns: allRows,
         viewId: $("#SelectedTableID option:selected").val(),
         IsTable: bIsTable,
         __RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
      };
      OpenHR.postData("Reports/AddAllOrganisationReportColumn", postData);

      for (var i = 0; i <= allRows.length - 1; i++) {
         rowID = allRows[i];
         var datarow = getDatarowFromAvailable(allRows[i]);

         $("#SelectedColumns").jqGrid('addRowData', datarow.ID, datarow);

      }

      $('#SelectedColumns').jqGrid("setSelection", rowID);
      $('#AvailableColumns').jqGrid('clearGridData');

      refreshcolumnPropertiesPanel();

   }

   function requestRemoveAllSelectedColumns() {
      removeAllSelectedColumns(true);
      ClearConcatenateCheckbox();
   }

   function requestRemoveSelectedColumns() {

      removeSelectedColumns();
      enableSaveButton();
      ClearConcatenateCheckbox();
   }

   function removeSelectedColumns() {

      var thisIndex = 0;
      var selectedRows = $('#SelectedColumns').jqGrid('getGridParam', 'selarrrow');

      var postData = {
         ReportID: '@Model.ID',
         ReportType: '@Model.ReportType',
         ColumnsTableID: $("#SelectedTableID").val(),
         Columns: selectedRows,
         __RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
      };

      OpenHR.postData("Reports/RemoveOrganisationReportColumn", postData);
      getAvailableTableViewColumns();

      // Position next selected column
      var recordCount = $("#SelectedColumns").jqGrid('getGridParam', 'records')
      var ids = $("#SelectedColumns").getDataIDs();
      if (thisIndex >= recordCount) { thisIndex = 0; }

      // Remove removed columns
      for (var i = selectedRows.length - 1; i >= 0; i--) {
         $("#SelectedColumns").delRowData(selectedRows[i]);
      }

      $("#SelectedColumns").jqGrid("setSelection", ids[thisIndex], true);

      // If records available and no row selected then select the first row
      if (($("#SelectedColumns").getGridParam("records") > 0) && ($("#SelectedColumns").jqGrid('getGridParam', 'selrow') == null)) {
         selectGridTopRow($('#SelectedColumns'));
      }

      refreshcolumnPropertiesPanel();

   }

   function removeAllSelectedColumns(reloadColumns) {

      var dataSend = {
         ReportID: '@Model.ID',
         ReportType: '@Model.ReportType',
         __RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
      };

      OpenHR.postData("Reports/RemoveAllOrganisationReportColumns", dataSend);
      $('#SelectedColumns').jqGrid('clearGridData');

      if (reloadColumns == true) {
         getAvailableTableViewColumns();
      }

      refreshcolumnPropertiesPanel();
   }


   function getAvailableTableViewColumns() {

      var sType;
      var bIsTable = 'false';
      var ViewID = $("#SelectedTableID").val();
      var url;

      if (($("#PostBasedTableId").val() != undefined &&
      $("#PostBasedTableId").val() == ViewID) &&
      ($("#PostBasedTableName").val() != undefined &&
      $("#PostBasedTableName").val() == $("#SelectedTableID option:selected")[0].text)) {
         bIsTable = true;
      }

      $("#AvailableColumns").jqGrid('GridUnload');

      $("#AvailableColumns").jqGrid({
         url: 'Reports/GetAvailableItemsForview?ReportID=' + '@Model.ID' + '&&viewOrTableId=' + $("#SelectedTableID").val() + '&&IsTable=' + bIsTable,
         datatype: 'json',
         mtype: 'GET',
         jsonReader: {
            root: "rows", //array containing actual data
            page: "page", //current page
            total: "total", //total pages for the query
            records: "records", //total number of records
            repeatitems: false,
            id: "ID"
         },
         colNames: ['ID', 'Name', 'DataType', 'Size', 'Decimals', 'Access', 'ViewID', 'TableID'],
         colModel: [
            { name: 'ID', index: 'ID', hidden: true },
            { name: 'Name', index: 'Name', width: 40, sortable: false },
            { name: 'DataType', index: 'DataType', hidden: true },
            { name: 'Size', index: 'Size', hidden: true },
            { name: 'Decimals', index: 'Decimals', hidden: true },
            { name: 'Access', index: 'Access', hidden: true },
            { name: 'ViewID', index: 'ViewID', hidden: true },
            { name: 'TableID', index: 'TableID', hidden: true }],
         viewrecords: true,
         autowidth: false,
         sortname: 'Name',
         sortorder: "desc",
         rowNum: 10000,
         scrollrows: true,
         multiselect: true,
         beforeSelectRow: function (rowid, e) {

            // If defination is readonly then skip this opertion and it will result in return false
            // which will stop calling onSelectRow
            if (!isDefinitionReadOnly()) {
               var $this = $(this), rows = this.rows,
                            // get id of the previous selected row
                            startId = $this.jqGrid('getGridParam', 'selrow'),
                            startRow, endRow, iStart, iEnd, i, rowidIndex;

               if (!e.ctrlKey && !e.shiftKey) {
                  $this.jqGrid('resetSelection');
               } else if (startId && e.shiftKey) {
                  $this.jqGrid('resetSelection');

                  // get DOM elements of the previous selected and the currect selected rows
                  startRow = rows.namedItem(startId);
                  endRow = rows.namedItem(rowid);
                  if (startRow && endRow) {
                     // get min and max from the indexes of the previous selected
                     // and the currect selected rows
                     iStart = Math.min(startRow.rowIndex, endRow.rowIndex);
                     rowidIndex = endRow.rowIndex;
                     iEnd = Math.max(startRow.rowIndex, rowidIndex);
                     for (i = iStart; i <= iEnd; i++) {
                        // the row with rowid will be selected by jqGrid, so:
                        if (i != rowidIndex) {
                           $this.jqGrid('setSelection', rows[i].id, false);
                        }
                     }
                  }

                  // clear text selection
                  if (document.selection && document.selection.empty) {
                     document.selection.empty();
                  } else if (window.getSelection) {
                     window.getSelection().removeAllRanges();
                  }
               }
               return true;
            }
         },
         ondblClickRow: function (rowid) {
            doubleClickAvailableColumn();
         },
         loadComplete: function (data) {
            removeSelectedColumnsFromAvailable();
            refreshcolumnPropertiesPanel();
            var topID = $("#AvailableColumns").getDataIDs()[0]
            $("#AvailableColumns").jqGrid("setSelection", topID);
         }
      });

      $("#AvailableColumns").jqGrid('hideCol', 'cb');

      $('#AvailableColumns').keydown(function (event) {
         event.preventDefault(); //prevent grid scrolling.
         var keyPressed = event.which;
         var grid = $('#AvailableColumns');
         //Enter key
         if (keyPressed == 13) {
            //handle this locally
            doubleClickAvailableColumn();
         }
         else {
            OpenHR.gridKeyboardEvent(keyPressed, grid);
         }
      });

      resizeColumnGrids();
   }

   function doubleClickAvailableColumn() {
      if (!isDefinitionReadOnly()) {
         var grid = $('#AvailableColumns');
         var currentScrollPos = grid.parent().parent().scrollTop();
         var rowid = grid.jqGrid('getGridParam', 'selrow');
         addColumnToSelected(rowid);
         enableSaveButton();

         if ((grid.getGridParam("records") > 0) && (grid.jqGrid('getGridParam', 'selrow') == null)) {
            OpenHR.gridSelectLastRow(grid);	// assume last row has been removed from grid
         }
         grid.focus();
         grid.parent().parent().scrollTop(currentScrollPos);

         return false;
      }
   }

   // Removes a selected column from the selectedColumn grid on double click of column
   function doubleClickSelectedColumn() {
      if (!isDefinitionReadOnly()) {
         var grid = $('#SelectedColumns');
         var currentScrollPos = grid.parent().parent().scrollTop();
         var rowid = grid.jqGrid('getGridParam', 'selrow');
         requestRemoveSelectedColumns();

         grid.focus();
         grid.parent().parent().scrollTop(currentScrollPos);

         return false;
      }
   }

   function refreshcolumnPropertiesPanel() {
      
      var rowCount = $('#SelectedColumns').jqGrid('getGridParam', 'selarrrow').length;
      var rowId = $("#SelectedColumns").jqGrid('getGridParam', 'selrow');
      var dataRow = $("#SelectedColumns").getRowData(rowId)
      var allRows = $('#SelectedColumns').jqGrid('getDataIDs');
      var bDisableAdd = ($("#AvailableColumns").getGridParam("reccount") == 0);
      var isTopRow = true;
      var isBottomRow = true;
      var isReadOnly = isDefinitionReadOnly();
      var bRowSelected = false;

      if (allRows.length > 0) {
         bRowSelected = true;
         isTopRow = (rowId == allRows[0]);
         isBottomRow = (rowId == allRows[allRows.length - 1]);
      }

      if (rowCount > 1 || allRows.length == 0) {
         $("#definitionColumnProperties :input").attr("disabled", true);
         $("#SelectedColumnPrefix").val("");
         $("#SelectedColumnSuffix").val("");
         $("#SelectedColumnHeight").val("");
         $("#SelectedColumnFontSize").val("");
         $("#SelectedColumnDecimals").val("");
         $('#SelectedColumnIsConcatenateWithNext').prop('checked', false);

         $(".canGroupWithNext").css("color", "#A59393");
      }
      else {

         if (!isReadOnly) {
            $("#definitionColumnProperties :input").removeAttr("disabled");
         }

         var isNumeric = (dataRow.DataType == '2' || dataRow.DataType == '4');
         var isDecimals = (isNumeric == true || dataRow.IsExpression == "true");
         var isGroupWithNext = $("#SelectedColumnIsConcatenateWithNext").is(':checked');
         var isSize = (dataRow.DataType == '4');
         var isPhotograph = (dataRow.DataType == -3)

         $(".decimalsOnly *").prop("disabled", !isDecimals || isReadOnly || isSize || isPhotograph);
         $(".canGroupWithNext *").prop("disabled", isBottomRow || isReadOnly || isPhotograph);
         $(".OrgReportsOnly *").prop("disabled", isReadOnly || isPhotograph);
         $(".fontsizeOnly *").prop("disabled", isReadOnly || isPhotograph);
         //$("#SelectedColumnFontSize").prop("disabled", isReadOnly || isPhotograph);
         $(".HeightOnly *").prop("disabled", isReadOnly);
       

         if (isBottomRow || isReadOnly) {
            $(".canGroupWithNext").css("color", "#A59393");
         } else {
            $(".canGroupWithNext").css("color", "#000000");
         }
      }

      // Enable / Disable relevant buttons
      button_disable($("#btnColumnAdd")[0], bDisableAdd || isReadOnly);
      button_disable($("#btnColumnAddAll")[0], bDisableAdd || isReadOnly);
      button_disable($("#btnColumnRemove")[0], !bRowSelected || isReadOnly);
      button_disable($("#btnColumnRemoveAll")[0], !bRowSelected || isReadOnly);
      button_disable($("#btnColumnMoveUp")[0], isTopRow || isReadOnly || (rowCount > 1));
      button_disable($("#btnColumnMoveDown")[0], isBottomRow || isReadOnly || (rowCount > 1));
      button_disable($("#btnOrgPreview")[0], !bRowSelected || isReadOnly);
   }

   function updateColumnsSelectedGrid() {

      var rowId = $("#SelectedColumns").jqGrid('getGridParam', 'selrow');
      var dataRow = $('#SelectedColumns').jqGrid('getRowData', rowId);

      dataRow.ColumnID = rowId;
      dataRow.Decimals = $("#SelectedColumnDecimals").val();
      if (dataRow.Decimals == "") { dataRow.Decimals = 0 }; //If Decimals is empty then set to 0

      dataRow.Height = $("#SelectedColumnHeight").val();
      if (dataRow.Height == "" && dataRow.DataType != -3) { dataRow.Height = 1 }; //If Height is empty then set to 1
      if (dataRow.Height == "" && dataRow.DataType == -3) { dataRow.Height = 3 };

      dataRow.FontSize = $("#SelectedColumnFontSize").val();
      if (dataRow.FontSize == "") { dataRow.FontSize = 11 }; //If fontSize is empty then set to 1

      dataRow.Suffix = $("#SelectedColumnSuffix").val();
      dataRow.Prefix = $("#SelectedColumnPrefix").val();

      dataRow.IsGroupWithNext = $("#SelectedColumnIsConcatenateWithNext").is(':checked');
      dataRow.Heading = $("#SelectedColumnHeading").val();

      $('#SelectedColumns').jqGrid('setRowData', rowId, dataRow);

   }

   function ClearConcatenateCheckbox() {

      var allRows = $('#SelectedColumns').jqGrid('getRowData');
      var rowid;
      for (var i = 0; i <= allRows.length - 1; i++) {
         if (allRows[i].IsGroupWithNext == "true") {
            allRows[i].IsGroupWithNext = "false";
            rowid = allRows[i].ID
            $('#SelectedColumns').jqGrid('setRowData', rowid, allRows[i]);
         }
      }

      $("#SelectedColumnIsConcatenateWithNext").prop('checked', false);
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
         colNames: ['ID', 'TableID', 'Name', 'Prefix', 'Suffix', 'FontSize', 'Height', 'DataType', 'Size', 'Decimals', 'IsGroupWithNext', 'ReportID', 'ReportType', 'Access', 'ViewID', 'ColumnID', 'Heading'],
         colModel: [
         { name: 'ID', index: 'ID', hidden: true },
         { name: 'TableID', index: 'TableID', hidden: true },
         { name: 'Name', index: 'Name', sortable: false },
         { name: 'Prefix', index: 'Prefix', hidden: true },
         { name: 'Suffix', index: 'Suffix', hidden: true },
         { name: 'FontSize', index: 'FontSize', hidden: true },
         { name: 'Height', index: 'Height', hidden: true },
         { name: 'DataType', index: 'DataType', hidden: true },
         { name: 'Size', index: 'Size', hidden: true },
         { name: 'Decimals', index: 'Decimals', hidden: true },
         { name: 'IsGroupWithNext', index: 'IsGroupWithNext', hidden: true },
         { name: 'ReportID', index: 'ReportID', hidden: true },
         { name: 'ReportType', index: 'ReportType', hidden: true },
         { name: 'Access', index: 'Access', hidden: true },
         { name: 'ViewID', index: 'ViewID', hidden: true },
         { name: 'ColumnID', index: 'ColumnID', hidden: true },
         { name: 'Heading', index: 'Heading', hidden: true }],
         viewrecords: true,
         autowidth: false,
         sortname: 'Name',
         sortorder: "asc",
         rowNum: 10000,
         scrollrows: true,
         multiselect: true,
         beforeSelectRow: function (rowid, e) {

            // If defination is readonly then skip this opertion and it will result in return false
            // which will stop calling onSelectRow
            if (!isDefinitionReadOnly()) {
               if ($('#SelectedColumns').jqGrid('getGridParam', 'selarrrow').length == 1) {
                  updateColumnsSelectedGrid();
               }

               var $this = $(this), rows = this.rows,
                // get id of the previous selected row
                startId = $this.jqGrid('getGridParam', 'selrow'),
                startRow, endRow, iStart, iEnd, i, rowidIndex;

               if (!e.ctrlKey && !e.shiftKey) {
                  $this.jqGrid('resetSelection');
               } else if (startId && e.shiftKey) {
                  $this.jqGrid('resetSelection');

                  // get DOM elements of the previous selected and the currect selected rows
                  startRow = rows.namedItem(startId);
                  endRow = rows.namedItem(rowid);
                  if (startRow && endRow) {
                     // get min and max from the indexes of the previous selected
                     // and the currect selected rows
                     iStart = Math.min(startRow.rowIndex, endRow.rowIndex);
                     rowidIndex = endRow.rowIndex;
                     iEnd = Math.max(startRow.rowIndex, rowidIndex);
                     for (i = iStart; i <= iEnd; i++) {
                        // the row with rowid will be selected by jqGrid, so:
                        if (i != rowidIndex) {
                           $this.jqGrid('setSelection', rows[i].id, false);
                        }
                     }
                  }

                  // clear text selection
                  if (document.selection && document.selection.empty) {
                     document.selection.empty();
                  } else if (window.getSelection) {
                     window.getSelection().removeAllRanges();
                  }
               }
               return true;
            }
         },
         onSelectRow: function (id) {

            var rowId = $("#SelectedColumns").jqGrid('getGridParam', 'selrow');
            var dataRow = $("#SelectedColumns").getRowData(rowId)

            $("#SelectedColumnHeading").val(decodeURI(dataRow.Heading));
            $("#SelectedColumnPrefix").val(decodeURI(dataRow.Prefix));
            $("#SelectedColumnSuffix").val(decodeURI(dataRow.Suffix));
            $("#SelectedColumnDecimals").val(dataRow.Decimals);
            $("#SelectedColumnFontSize").val(dataRow.FontSize);
            $("#SelectedColumnHeight").val(dataRow.Height);

            if ($("#SelectedColumnFontSize").val() == "") {
               $("#SelectedColumnFontSize").val("11");
            }
            if ($("#SelectedColumnHeight").val() == "" && dataRow.DataType != -3) {
               $("#SelectedColumnHeight").val(1);
            }
            if ($("#SelectedColumnHeight").val() == "" && dataRow.DataType == -3) {
               $("#SelectedColumnHeight").val(3);
            }

            if (dataRow.IsGroupWithNext == "true") {
               $("#SelectedColumnIsConcatenateWithNext").prop('checked', true);
            }
            else {
               $("#SelectedColumnIsConcatenateWithNext").prop('checked', false);
            }

            refreshcolumnPropertiesPanel();
            disableColumnOptionsWhenConcatenateWithNextChecked();
         },
         ondblClickRow: function () {
            doubleClickSelectedColumn();
         },
         loadComplete: function (data) {
            var topID = $("#SelectedColumns").getDataIDs()[0]
            $("#SelectedColumns").jqGrid("setSelection", topID);

         }
      });

      $("#SelectedColumns").jqGrid('hideCol', 'cb');

      $('#SelectedColumns').keydown(function (event) {
         event.preventDefault(); //prevent grid scrolling.
         var keyPressed = event.which;
         var grid = $('#SelectedColumns');
         //Enter key
         if (keyPressed == 13) {
            //handle this locally
            requestRemoveSelectedColumns();
         }
         else {
            OpenHR.gridKeyboardEvent(keyPressed, grid);
         }
      });

   }
   function changeColumnIsConcatenateWithNext() {

      var IsGroupWithNext = $("#SelectedColumnIsConcatenateWithNext").is(':checked');

      refreshcolumnPropertiesPanel();
      updateColumnsSelectedGrid();
      disableColumnOptionsWhenConcatenateWithNextChecked();

   }

   // Disabled the column options for the current row and uncheck all the column options for the next row
   // when ConcatenateWithNext is checked for the current row.
   function disableColumnOptionsWhenConcatenateWithNextChecked() {

      var rowId = $("#SelectedColumns").jqGrid('getGridParam', 'selrow');

      // Gets all row ID'S of selected columns. Here the index begin with zero.
      var allRows = $('#SelectedColumns').jqGrid('getDataIDs');
      var dataRow = $("#SelectedColumns").getRowData(rowId)
      var isBottomRow = (rowId == allRows[allRows.length - 1]);
      var currentRowIndex = $("#SelectedColumns").getInd(rowId);
      var CanConcatenate = true;
      var PreviousRowConcatenate = false; //Checks if previous row Concatenate checked
      var isPhotograph = (dataRow.DataType == -3)

      var NextRowIndex = $('#SelectedColumns').jqGrid('getRowData', allRows[currentRowIndex]);     //Get Next Column's data
      var prevDataRow = $('#SelectedColumns').jqGrid('getRowData', allRows[currentRowIndex - 2]);  //Get Previous Column's data

      if (prevDataRow.IsGroupWithNext == "true") {

         PreviousRowConcatenate = true;
         $("#SelectedColumnFontSize").val(prevDataRow.FontSize); //Set FontSize As Previous Column
         $("#SelectedColumnHeight").val(prevDataRow.Height);  //  Set Height As Previous Column
         $("#SelectedColumnPrefix").val(""); // Clear Prefix value
      }

      // Check ViewID of next column
      // Check TableID of Next Column
      // Check If Selected Column or Next Column is not PhotoType
      if (dataRow.ViewID == NextRowIndex.ViewID && dataRow.TableID == NextRowIndex.TableID && dataRow.DataType != -3 && NextRowIndex.DataType != -3) {
         CanConcatenate = true;
      }
      else {
         CanConcatenate = false;
      }

      $(".canGroupWithNext *").prop("disabled", !CanConcatenate);
      $("#SelectedColumnPrefix").prop("disabled", PreviousRowConcatenate || isPhotograph);
      $("#SelectedColumnSuffix").prop("disabled", isPhotograph);
      $("#SelectedColumnHeight").prop("disabled", PreviousRowConcatenate);
      $("#SelectedColumnFontSize").prop("disabled", PreviousRowConcatenate || isPhotograph);

      if (!CanConcatenate) {
         $(".canGroupWithNext").css("color", "#A59393");
      } else {
         $(".canGroupWithNext").css("color", "#000000");
      }
   }

   function getUnauthorisedColumns()
   {
      var count = 0;
      var columnsString = "";
      var invalidColumns = @Html.Raw(Json.Encode(Model.InvalidColumnList));
      for (var i = 0; i < invalidColumns.length; i++) {
         count++;
         columnsString += invalidColumns[i].Heading + ", ";
      }
      if (count > 0 )
      {
         columnsString = columnsString.substring(0, columnsString.length - 2);
         OpenHR.modalMessage("You do not have access to following columns : " + columnsString + ".");
         enableSaveButton();
      }
   }

   // Initialise
   $(function () {

      // Sets Decimals textbox to allow numeric only
      $("#SelectedColumnDecimals").autoNumeric({ aSep: '', aNeg: '', mDec: "0", vMax: 999, vMin: 0 });

      //Note:-
      //This solution working in Firefox, Chrome and IE, both with keyboard focus and mouse focus.
      //It also handles correctly clicks following the focus (it moves the caret and doesn't reselect the text):
      //With keyboard focus, only onfocus triggers which selects the text because this.clicked is not set. With mouse focus, onmousedown triggers, then onfocus and then onclick which selects the text in onclick but not in onfocus (Chrome requires this).
      //Mouse clicks when the field is already focused don't trigger onfocus which results in not selecting anything.
      $(".selectFullText").bind({
         click: function () {
            if (this.clicked == 2) this.select(); this.clicked = 0;
         },
         mousedown: function () {
            this.clicked = 1;
         },
         focus: function () {
            if (!this.clicked) this.select(); else this.clicked = 2;
         }
      }).blur(function () {
         if (this.value == "") this.value = 0;
      });


      $("#SelectedColumnHeight").bind({
         click: function () {
            if (this.clicked == 2) this.select(); this.clicked = 0;
         },
         mousedown: function () {
            this.clicked = 1;
         },
         focus: function () {
            if (!this.clicked) this.select(); else this.clicked = 2;
         }
      }).blur(function (sender) {

         var rowId = $("#SelectedColumns").jqGrid('getGridParam', 'selrow');
         var dataRow = $("#SelectedColumns").getRowData(rowId)
         var ColumnName = dataRow.Name;

         if ((dataRow.DataType == -3)) { var Min = 3 }
         else { var Min = 1; }

         var Max = 6;

         if ((sender.target.value == "") || (sender.target.value < Min) || (sender.target.value > Max)) {
            OpenHR.modalMessage("Enter height (rows) between " + Min + " and " + Max); sender.target.value = Min;
            $(sender.target.id).focus();
         }
      });


      $("#SelectedColumnFontSize").bind({
         click: function () {
            if (this.clicked == 2) this.select(); this.clicked = 0;
         },
         mousedown: function () {
            this.clicked = 1;
         },
         focus: function () {
            if (!this.clicked) this.select(); else this.clicked = 2;
         }
      }).blur(function (sender) {
         var MinSize = 6;
         var MaxSize = 30;
         if ((sender.target.value == "") || (sender.target.value < MinSize) || (sender.target.value > MaxSize)) {
            OpenHR.modalMessage("Enter font size between 6 and 30");
            sender.target.value = "11";
            $(sender.target.id).focus();
         }
      });

      $("#SelectedColumnFontSize").bind('keypress', function (event) {
         var regex = new RegExp("^[0-9]+$");
         var key = String.fromCharCode(!event.charCode ? event.which : event.charCode);
         // Keycode = 8(backspace), 9(tab), 46(delete) // Will work for firefox
         if (!regex.test(key) && event.keyCode != 8 && event.keyCode != 9 && event.keyCode != 46 && event.keyCode != 39) {
            event.preventDefault();
            return false;
         }
      });

      $("#SelectedColumnHeight").bind('keypress', function (event) {
         var regex = new RegExp("^[0-9]+$");
         var key = String.fromCharCode(!event.charCode ? event.which : event.charCode);
         if (!regex.test(key) && event.keyCode != 8 && event.keyCode != 9 && event.keyCode != 46 && event.keyCode != 39) {
            event.preventDefault();
            return false;
         }
      });

   });
</script>