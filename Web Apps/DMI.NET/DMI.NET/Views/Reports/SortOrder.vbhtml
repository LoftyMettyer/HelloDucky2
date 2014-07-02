@Imports DMI.NET
@Imports DMI.NET.Helpers
@Inherits System.Web.Mvc.WebViewPage(Of Models.ReportBaseModel)

@Code
  Layout = Nothing
End Code


@Html.SortOrderGrid("SortOrderColumns", Model.SortOrderColumns, Nothing)

<input type="button" id="btnSortOrderAdd" value="Add" onclick="sortAdd();" />
<input type="button" id="btnSortOrderEdit" value="Edit" />
<input type="button" id="btnSortOrderRemove" value="Remove" />
<input type="button" id="btnSortOrderRemoveAll" value="Remove All" onclick="removeAllSortOrders();" />
<input type="button" id="btnSortOrderMoveUp" value="Move Up" />
<input type="button" id="btnSortOrderMoveDown" value="Move Down" />



<form id="frmSortOrder" name="frmSortOrder" action="util_sortorderselection" target="sortorderselection" method="post" style="visibility: hidden; display: none">
  <input type="hidden" id="txtSortInclude" name="txtSortInclude">
  <input type="hidden" id="txtSortExclude" name="txtSortExclude">
  <input type="hidden" id="txtSortEditing" name="txtSortEditing">
  <input type="hidden" id="txtSortColumnID" name="txtSortColumnID">
  <input type="hidden" id="txtSortColumnName" name="txtSortColumnName">
  <input type="hidden" id="txtSortOrder" name="txtSortOrder">
  <input type="hidden" id="txtSortBOC" name="txtSortBOC">
  <input type="hidden" id="txtSortPOC" name="txtSortPOC">
  <input type="hidden" id="txtSortVOC" name="txtSortVOC">
  <input type="hidden" id="txtSortSRV" name="txtSortSRV">
</form>

<div id="sortOrderPopup">
  <input type="checkbox" id="txtSortOrderBOC" />
  <input type="checkbox" id="txtSortOrderPOC" />
  <input type="checkbox" id="txtSortOrderVOC" />
  <input type="text" id="ColExprID" />
  <button name="butSaveOrderOrderColumn" value="Save" onclick="saveSortOrderItem();">Save</button>
</div>


<script type="text/javascript">

  $("#sortOrderPopup").dialog({
    autoOpen: false,
    modal: true,
    width: 500
  });

  function saveSortOrderItem() {

    var itemIndex = $("#SortOrderColumns tr").length - 1;
    var newItem = "<tr>"
      + "<td><input type='text' name='SortOrderColumns[" + itemIndex + "].ColumnID' value='" + $("#ColExprID").val() + "'/></td>"
      + "<td><input type='text' name='SortOrderColumns[" + itemIndex + "].BOC' value='" + $("#txtSortOrderBOC").val() + "'/></td>"
      + "</tr>";

    $("#SortOrderColumns").append(newItem);
    $("#sortOrderPopup").dialog("close");

  }


  function sortAdd() {
    var i;
    var iCalcsCount = 0;
    var iColumnsCount = 0;
    var sURL;


    var frmSortOrder = $("#frmSortOrder")[0];


    var sortOrderPopup = $("#sortOrderPopup")[0]

    $("#sortOrderPopup").dialog('open');
    return true

    // Loop through the columns added and populate the 
    // sort order text boxes to pass to util_sortorderselection.asp
    frmSortOrder.txtSortInclude.value = '';
    frmSortOrder.txtSortExclude.value = '';
    frmSortOrder.txtSortEditing.value = 'false';
    //frmSortOrder.txtSortColumnID.value = frmDefinition.ssOleDBGridSortOrder.Columns(0).text;
    //frmSortOrder.txtSortColumnName.value = frmDefinition.ssOleDBGridSortOrder.Columns(1).text;
    //frmSortOrder.txtSortOrder.value = frmDefinition.ssOleDBGridSortOrder.Columns(2).text;
    //frmSortOrder.txtSortBOC.value = frmDefinition.ssOleDBGridSortOrder.Columns(3).text;
    //frmSortOrder.txtSortPOC.value = frmDefinition.ssOleDBGridSortOrder.Columns(4).text;
    //frmSortOrder.txtSortVOC.value = frmDefinition.ssOleDBGridSortOrder.Columns(5).text;
    //frmSortOrder.txtSortSRV.value = frmDefinition.ssOleDBGridSortOrder.Columns(6).text;

    //if (frmUseful.txtSelectedColumnsLoaded.value == 1) {
    //  frmDefinition.ssOleDBGridSelectedColumns.Redraw = false;
    //  frmDefinition.ssOleDBGridSelectedColumns.movefirst();

    //  for (i = 0; i < frmDefinition.ssOleDBGridSelectedColumns.rows; i++) {
    //    if (frmDefinition.ssOleDBGridSelectedColumns.Columns(0).Text == 'C') {
    //      iColumnsCount++;
    //      if (frmSortOrder.txtSortInclude.value != '') {
    //        frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + ',';
    //      }
    //      frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + frmDefinition.ssOleDBGridSelectedColumns.columns(2).text;
    //    }
    //    else {
    //      iCalcsCount++;
    //    }
    //    frmDefinition.ssOleDBGridSelectedColumns.movenext();
    //  }

    //  frmDefinition.ssOleDBGridSelectedColumns.Redraw = true;
    //}
    //else {
    //  var dataCollection = frmOriginalDefinition.elements;
    //  if (dataCollection != null) {
    //    for (var iIndex = 0; iIndex < dataCollection.length; iIndex++) {
    //      var sControlName = dataCollection.item(iIndex).name;
    //      sControlName = sControlName.substr(0, 20);
    //      if (sControlName == "txtReportDefnColumn_") {
    //        var sDefnString = new String(dataCollection.item(iIndex).value);

    //        if (sDefnString.length > 0) {
    //          var sType = selectedColumnParameter(sDefnString, "TYPE");
    //          var sColumnID = selectedColumnParameter(sDefnString, "COLUMNID");

    //          if (sType == 'C') {
    //            iColumnsCount++;

    //            if (frmSortOrder.txtSortInclude.value != '') {
    //              frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + ',';
    //            }
    //            frmSortOrder.txtSortInclude.value = frmSortOrder.txtSortInclude.value + sColumnID;
    //          }
    //          else {
    //            iCalcsCount++;
    //          }
    //        }
    //      }
    //    }
    //  }
    //}

    //if (frmDefinition.ssOleDBGridSortOrder.Rows > 0) {
    //  frmDefinition.ssOleDBGridSortOrder.Redraw = false;
    //  frmDefinition.ssOleDBGridSortOrder.movefirst();

    //  for (i = 0; i < frmDefinition.ssOleDBGridSortOrder.rows; i++) {
    //    if (frmSortOrder.txtSortExclude.value != '') {
    //      frmSortOrder.txtSortExclude.value = frmSortOrder.txtSortExclude.value + ',';
    //    }

    //    frmSortOrder.txtSortExclude.value = frmSortOrder.txtSortExclude.value + frmDefinition.ssOleDBGridSortOrder.columns(0).text;

    //    frmDefinition.ssOleDBGridSortOrder.movenext();
    //  }

    //  frmDefinition.ssOleDBGridSortOrder.Redraw = true;
    //}

    //if (frmSortOrder.txtSortInclude.value == frmSortOrder.txtSortExclude.value) {
    //  OpenHR.messageBox("You must add more columns to the report before you can add to the sort order.", 48, "Custom Reports");
    //}
    //else if ((frmDefinition.ssOleDBGridSortOrder.Rows - iColumnsCount) == 0) {
    //  OpenHR.messageBox("You must add more columns to the report before you can add to the sort order.", 48, "Custom Reports");
    //}
    //else {
//      if (frmSortOrder.txtSortInclude.value != '') {
        sURL = "util_sortorderselection" +
					"?txtSortInclude=" + escape(frmSortOrder.txtSortInclude.value) +
						"&txtSortExclude=" + escape(frmSortOrder.txtSortExclude.value) +
							"&txtSortEditing=" + escape(frmSortOrder.txtSortEditing.value) +
								"&txtSortColumnID=" + escape(frmSortOrder.txtSortColumnID.value) +
									"&txtSortColumnName=" + escape(frmSortOrder.txtSortColumnName.value) +
										"&txtSortOrder=" + escape(frmSortOrder.txtSortOrder.value) +
											"&txtSortBOC=" + escape(frmSortOrder.txtSortBOC.value) +
												"&txtSortPOC=" + escape(frmSortOrder.txtSortPOC.value) +
													"&txtSortVOC=" + escape(frmSortOrder.txtSortVOC.value) +
														"&txtSortSRV=" + escape(frmSortOrder.txtSortSRV.value);
        openDialog(sURL, 600, 300, "no", "no");
        //openDialog(sURL, (screen.width) / 3 + 40, (screen.height) / 2, "no", "no");

        frmUseful.txtChanged.value = 1;
   //   }
 //   }

    //frmDefinition.ssOleDBGridSortOrder.Refresh();
    //frmDefinition.ssOleDBGridRepetition.Refresh();

    //refreshTab4Controls();
  }

  function removeAllSortOrders() {
    $("#SortOrderColumns").empty();
  }



</script>
