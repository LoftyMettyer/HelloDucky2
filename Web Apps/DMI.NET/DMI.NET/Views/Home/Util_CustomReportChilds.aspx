<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

<!DOCTYPE html>
<html>
<head>
    <title>OpenHR Intranet</title>
    <script src="<%: Url.Content("~/bundles/jQuery")%>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/bundles/OpenHR_General")%>" type="text/javascript"></script>           
</head>

<script type="text/javascript">
    window.onload = function () {
		
        self.focus();

        var iResizeBy, iNewWidth, iNewHeight, iNewLeft, iNewTop, frmPopup = document.getElementById("frmPopup");

        // Resize the grid to show all prompted values.
        iResizeBy = frmPopup.offsetParent.scrollWidth	- frmPopup.offsetParent.clientWidth;
        if (frmPopup.offsetParent.offsetWidth + iResizeBy > screen.width) {
            window.dialogWidth = new String(screen.width) + "px";
        }
        else {
            iNewWidth = new Number(window.dialogWidth.substr(0, window.dialogWidth.length-2));
            iNewWidth = iNewWidth + iResizeBy;
            window.dialogWidth = new String(iNewWidth) + "px";
        }

        iResizeBy = frmPopup.offsetParent.scrollHeight	- frmPopup.offsetParent.clientHeight;
        if (frmPopup.offsetParent.offsetHeight + iResizeBy > screen.height) {
            window.dialogHeight = new String(screen.height) + "px";
        }
        else {
            iNewHeight = new Number(window.dialogHeight.substr(0, window.dialogHeight.length-2));
            iNewHeight = iNewHeight + iResizeBy;
            window.dialogHeight = new String(iNewHeight) + "px";
        }

        iNewLeft = (screen.width - frmPopup.offsetParent.offsetWidth) / 2;
        iNewTop = (screen.height - frmPopup.offsetParent.offsetHeight) / 2;
        window.dialogLeft = new String(iNewLeft) + "px";
        window.dialogTop = new String(iNewTop) + "px";

        var frmChild = window.dialogArguments.OpenHR.getForm("workframe", "frmCustomReportChilds");
        var frmDef = window.dialogArguments.OpenHR.getForm("workframe", "frmDefinition");

        populateChildCombo();
	
        if(frmChild.childAction.value.toUpperCase() == "NEW")
        {
            frmPopup.cboChildTable.selectedIndex = 0;
            frmPopup.txtChildFilter.value = frmChild.childFilter.value;
            frmPopup.txtChildFilterID.value = frmChild.childFilterID.value;
            frmPopup.txtFieldRecOrder.value = frmChild.childOrder.value;
            frmPopup.txtChildFieldOrderID.value = frmChild.childOrderID.value;
            frmPopup.txtChildRecords.value = frmChild.childRecords.value;
            showAllRecords();
        }
        else
        {
            frmPopup.rowID.value = frmDef.ssOleDBGridChildren.AddItemRowIndex(frmDef.ssOleDBGridChildren.Bookmark);
            frmPopup.originalChildID.value = frmChild.childTableID.value;
            setChildTable(frmChild.childTableID.value);
            frmPopup.txtChildFilter.value = frmChild.childFilter.value;
            frmPopup.txtChildFilterID.value = frmChild.childFilterID.value;
            frmPopup.txtFieldRecOrder.value = frmChild.childOrder.value;
            frmPopup.txtChildFieldOrderID.value = frmChild.childOrderID.value;
            frmPopup.txtChildRecords.value = frmChild.childRecords.value;
            showAllRecords();
        }
    }
</script>
    
<script type="text/javascript">
    
    function populateChildCombo() {

        var frmChild = window.dialogArguments.OpenHR.getForm("workframe", "frmCustomReportChilds");
        var frmPopup = document.getElementById("frmPopup");
		
        //var sChildren = frmChild.childrenString.value;
        var sChildren = frmChild.childrenNames.value;
        var bAdded = false;
        var sChildID;
        var sChildName;
		
        var oOption;

        var iIndex = sChildren.indexOf("	");
        while (iIndex > 0) {
            sChildID = sChildren.substr(0, iIndex);

            if (alreadyUsedInReport(sChildID) == false
                || sChildID == frmChild.childTableID.value) {
                bAdded = true;

                oOption = document.createElement("OPTION");
                frmPopup.cboChildTable.options.add(oOption);

                //calling the getTableName() function for each of the child tables
                //as it loops throught the Table elements collection each time.
                //oOption.innerText = getTableName(sChildID);
                //oOption.innerText = sChildID;
                oOption.value = sChildID;

                sChildren = sChildren.substr(iIndex + 1);
                iIndex = sChildren.indexOf("	");

                sChildName = sChildren.substr(0, iIndex);
                oOption.innerText = sChildName;
            }

            if (bAdded) {
                sChildren = sChildren.substr(iIndex + 1);
                iIndex = sChildren.indexOf("	");

                bAdded = false;
            }
            else {
                sChildren = sChildren.substr(iIndex + 1);
                iIndex = sChildren.indexOf("	");

                sChildren = sChildren.substr(iIndex + 1);
                iIndex = sChildren.indexOf("	");

                bAdded = false;
            }
        }

        if (frmPopup.cboChildTable.options.length < 2) {
            combo_disable(frmPopup.cboChildTable, true);
        }
    }

    function alreadyUsedInReport(piChildID) {
        var frmDef = window.dialogArguments.OpenHR.getForm("workframe", "frmDefinition");
        
        var pvarbookmark;

        for (var i = 0; i < frmDef.ssOleDBGridChildren.rows; i++) {
            pvarbookmark = frmDef.ssOleDBGridChildren.AddItemBookmark(i);
            if (frmDef.ssOleDBGridChildren.Columns('TableID').CellText(pvarbookmark) == piChildID) {
                return true;
            }
        }
        return false;
    }

    function setChildTable(piTableID) {
        var i;
        var frmPopup = document.getElementById("frmPopup");
		
        for (i = 0; i < frmPopup.cboChildTable.options.length; i++) {
            if (frmPopup.cboChildTable.options(i).value == piTableID) {
                frmPopup.cboChildTable.selectedIndex = i;
                return;
            }
        }
        frmPopup.cboChildTable.selectedIndex = 0;
    }

    function changeChildTable() {
        var frmPopup = document.getElementById("frmPopup");
        frmPopup.txtChildFilterID.value = 0;
        frmPopup.txtChildFilter.value = "";
        frmPopup.txtChildFieldOrderID.value = 0;
        frmPopup.txtFieldRecOrder.value = "";
        frmPopup.txtChildRecords.value = 0;
    }

    function getTableName(piTableID) {

        var i;
        var sTableName;
        var frmTab = window.dialogArguments.OpenHR.getForm("workframe", "frmTables");

        var sReqdControlName = new String("txtTableName_");
        sReqdControlName = sReqdControlName.concat(piTableID);

        var dataCollection = frmTab.elements;
        if (dataCollection != null) {
            for (i = 0; i < dataCollection.length; i++) {
                var sControlName = dataCollection.item(i).name;

                if (sControlName == sReqdControlName) {
                    sTableName = dataCollection.item(i).value;
                    return sTableName;
                }
            }
        }
        return null;
    }

    function setRecordsNumeric() {
        var sConvertedValue;
        var sDecimalSeparator;
        var sThousandSeparator;
        var sPoint;
        var frmPopup = document.getElementById("frmPopup");
		
        sDecimalSeparator = "\\";
        sDecimalSeparator = sDecimalSeparator.concat(OpenHR.LocaleDecimalSeparator);
        var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");

        sThousandSeparator = "\\";
        sThousandSeparator = sThousandSeparator.concat(OpenHR.LocaleThousandSeparator);
        var reThousandSeparator = new RegExp(sThousandSeparator, "gi");

        sPoint = "\\.";
        var rePoint = new RegExp(sPoint, "gi");

        if (frmPopup.txtChildRecords.value == '') {
            frmPopup.txtChildRecords.value = 0;
        }

        // Convert the value from locale to UK settings for use with the isNaN funtion.
        sConvertedValue = new String(frmPopup.txtChildRecords.value);

        // Remove any thousand separators.
        sConvertedValue = sConvertedValue.replace(reThousandSeparator, "");
        frmPopup.txtChildRecords.value = sConvertedValue;

        // Convert any decimal separators to '.'.
        if (OpenHR.LocaleDecimalSeparator != ".") {
            // Remove decimal points.
            sConvertedValue = sConvertedValue.replace(rePoint, "A");
            // replace the locale decimal marker with the decimal point.
            sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
        }

        if (isNaN(sConvertedValue) == true) {
            OpenHR.messageBox("Invalid numeric value.", 48, "Custom Reports");
            frmPopup.txtChildRecords.value = 0;
        }
        else {
            if (sConvertedValue.indexOf(".") >= 0) {
                OpenHR.messageBox("Invalid integer value.", 48, "Custom Reports");
                frmPopup.txtChildRecords.value = 0;
            }
            else {
                if (frmPopup.txtChildRecords.value < 0) {
                    OpenHR.messageBox("The value cannot be negative.", 48, "Custom Reports");
                    frmPopup.txtChildRecords.value = 0;
                }
                else {
                    if (frmPopup.txtChildRecords.value > 999) {
                        OpenHR.messageBox("The value cannot be greater than 999.", 48, "Custom Reports");
                        frmPopup.txtChildRecords.value = 999;
                    }
                }
            }
        }
    }

    function showAllRecords() {
        var frmPopup = document.getElementById("frmPopup");
        if (frmPopup.txtChildRecords.value == 0) {
            frmPopup.txtAllRecords.value = "(All Records)";
        }
        else {
            frmPopup.txtAllRecords.value = "";
        }
    }

    function spinRecords(pfUp) {
        var frmPopup = document.getElementById("frmPopup");
        var iRecords = frmPopup.txtChildRecords.value;
        if (pfUp == true) {
            iRecords = ++iRecords;
        }
        else {
            if (iRecords > 0) {
                iRecords = iRecords - 1;
            }
        }
        frmPopup.txtChildRecords.value = iRecords;
    }

    function selectRecordOrder() {
        var sURL;
        var frmPopup = document.getElementById("frmPopup");
        var frmRecOrder = document.getElementById("frmRecOrder");
		
        frmRecOrder.selectionType.value = "ORDER";
        frmRecOrder.txtTableID.value = frmPopup.cboChildTable.options[frmPopup.cboChildTable.selectedIndex].value;
        frmRecOrder.selectedID.value = frmPopup.txtChildFieldOrderID.value;

        sURL = "fieldRec" +
            "?selectionType=" + escape(frmRecOrder.selectionType.value) +
            "&txtTableID=" + escape(frmRecOrder.txtTableID.value) +
            "&selectedID=" + escape(frmRecOrder.selectedID.value);
        openDialog(sURL, (screen.width) / 3, (screen.height) / 2, "yes", "yes");
    }

    function selectRecordOption(psTable, psType) {
        var frmUse = window.dialogArguments.OpenHR.getForm("workframe", "frmUseful");
        var frmDef = window.dialogArguments.OpenHR.getForm("workframe", "frmDefinition");

        var sURL;
        var frmPopup = document.getElementById("frmPopup");
        var frmRecordSelection = document.getElementById("frmRecordSelection");
		
        if (psTable == 'child') {
            var iTableID = frmPopup.cboChildTable.options[frmPopup.cboChildTable.selectedIndex].value;
            var iCurrentID = frmPopup.txtChildFilterID.value;
        }
        frmRecordSelection.recSelTable.value = psTable;
        frmRecordSelection.recSelType.value = psType;
        frmRecordSelection.recSelTableID.value = iTableID;
        frmRecordSelection.recSelCurrentID.value = iCurrentID;

        var strDefOwner = new String(frmDef.txtOwner.value);
        var strCurrentUser = new String(frmUse.txtUserName.value);

        strDefOwner = strDefOwner.toLowerCase();
        strCurrentUser = strCurrentUser.toLowerCase();

        if (strDefOwner == strCurrentUser) {
            frmRecordSelection.recSelDefOwner.value = '1';
        }
        else {
            frmRecordSelection.recSelDefOwner.value = '0';
        }
        frmRecordSelection.recSelDefType.value = "Custom Reports";

        sURL = "util_recordSelection" +
            "?recSelType=" + escape(frmRecordSelection.recSelType.value) +
            "&recSelTableID=" + escape(frmRecordSelection.recSelTableID.value) +
            "&recSelCurrentID=" + escape(frmRecordSelection.recSelCurrentID.value) +
            "&recSelTable=" + escape(frmRecordSelection.recSelTable.value) +
            "&recSelDefOwner=" + escape(frmRecordSelection.recSelDefOwner.value) +
            "&recSelDefType=" + escape(frmRecordSelection.recSelDefType.value);
        openDialog(sURL, (screen.width) / 3, (screen.height) / 2, "yes", "yes");
    }

    function openDialog(pDestination, pWidth, pHeight, psResizable, psScroll) {
        var dlgwinprops = "center:yes;" +
            "dialogHeight:" + pHeight + "px;" +
            "dialogWidth:" + pWidth + "px;" +
            "help:no;" +
            "resizable:" + psResizable + ";" +
            "scroll:" + psScroll + ";" +
            "status:no;";
        window.showModalDialog(pDestination, self, dlgwinprops);
    }

    function removeChildTable(piChildTableID) {
        var i;
        var iCount;
        var iTableID;
        var fChildColumnsSelected;
        var iIndex;
        var sControlName;
        var dataCollection;
		
        var frmUseful = window.dialogArguments.OpenHR.getForm("workframe", "frmUseful");
        var frmDefinition = window.dialogArguments.OpenHR.getForm("workframe", "frmDefinition");
        var frmOriginalDefinition = window.dialogArguments.OpenHR.getForm("workframe", "frmOriginalDefinition");

        frmUseful.txtCurrentChildTableID.value = piChildTableID;

        if (frmUseful.txtLoading.value == 'N') {
            if ((frmDefinition.ssOleDBGridSelectedColumns.Rows > 0) ||
                ((frmUseful.txtAction.value.toUpperCase() != "NEW") &&
                    (frmUseful.txtSelectedColumnsLoaded.value == 0))) {
                if (frmUseful.txtCurrentChildTableID.value != 0)
                    //if (frmDefinition.ssOleDBGridChildren.Rows > 0)
                {

                    // Check if there are any child columns in the selected columns list.
                    fChildColumnsSelected = false;
                    if (frmUseful.txtSelectedColumnsLoaded.value == 1) {
                        if (frmDefinition.ssOleDBGridSelectedColumns.Rows > 0) {
                            frmDefinition.ssOleDBGridSelectedColumns.Redraw = false;
                            frmDefinition.ssOleDBGridSelectedColumns.movefirst();

                            for (i = 0; i < frmDefinition.ssOleDBGridSelectedColumns.rows; i++) {
                                iTableID = frmDefinition.ssOleDBGridSelectedColumns.Columns("tableID").Text;

                                if (window.dialogArguments.isSelectedChildTable(iTableID)) {
                                    fChildColumnsSelected = true;
                                    break;
                                }

                                if (iTableID == frmUseful.txtCurrentChildTableID.value) {
                                    fChildColumnsSelected = true;
                                    break;
                                }

                                frmDefinition.ssOleDBGridSelectedColumns.movenext();
                            }

                            frmDefinition.ssOleDBGridSelectedColumns.Redraw = true;
                            frmDefinition.ssOleDBGridSelectedColumns.selbookmarks.removeall();
                            frmDefinition.ssOleDBGridSelectedColumns.selbookmarks.add(frmDefinition.ssOleDBGridSelectedColumns.bookmark);
                        }
                    }
                    else {
                        dataCollection = frmOriginalDefinition.elements;
                        if (dataCollection != null) {
                            for (iIndex = 0; iIndex < dataCollection.length; iIndex++) {
                                sControlName = dataCollection.item(iIndex).name;
                                sControlName = sControlName.substr(0, 20);
                                if (sControlName == "txtReportDefnColumn_") {
                                    iTableID = window.dialogArguments.window.selectedColumnParameter(dataCollection.item(iIndex).value, "TABLEID");

                                    if (window.dialogArguments.isSelectedChildTable(iTableID)) {
                                        fChildColumnsSelected = true;
                                        break;
                                    }

                                    if (iTableID == frmUseful.txtCurrentChildTableID.value) {
                                        fChildColumnsSelected = true;
                                        break;
                                    }

                                }
                            }
                        }
                    }

                    if (fChildColumnsSelected == true) {
                        var iAnswer = OpenHR.messageBox("One or more columns from the child table have been included in the report definition. Changing the child table will remove these columns from the report definition. Do you wish to continue ?", 36, "Custom Reports");

                        if (iAnswer == 7) {
                            // cancel and change back !
                            return false;
                        }
                        else {
                            // Remove the child table's columns from the selected columns collection.
                            if (frmUseful.txtSelectedColumnsLoaded.value == 1) {
                                if (frmDefinition.ssOleDBGridSelectedColumns.Rows > 0) {
                                    frmDefinition.ssOleDBGridSelectedColumns.Redraw = false;
                                    frmDefinition.ssOleDBGridSelectedColumns.MoveFirst();

                                    iCount = frmDefinition.ssOleDBGridSelectedColumns.rows;
                                    for (i = 0; i < iCount; i++) {
                                        iTableID = frmDefinition.ssOleDBGridSelectedColumns.Columns("tableID").Text;
                                        if (iTableID == frmUseful.txtCurrentChildTableID.value) {
                                            if (frmDefinition.ssOleDBGridSelectedColumns.rows == 1) {
                                                frmDefinition.ssOleDBGridSelectedColumns.RemoveAll();
                                            }
                                            else {
                                                frmDefinition.ssOleDBGridSelectedColumns.RemoveItem(frmDefinition.ssOleDBGridSelectedColumns.AddItemRowIndex(frmDefinition.ssOleDBGridSelectedColumns.Bookmark));
                                            }
                                        }
                                        frmDefinition.ssOleDBGridSelectedColumns.MoveNext();
                                    }

                                    frmDefinition.ssOleDBGridSelectedColumns.Redraw = true;
                                    frmDefinition.ssOleDBGridSelectedColumns.selbookmarks.removeall();
                                    frmDefinition.ssOleDBGridSelectedColumns.selbookmarks.add(frmDefinition.ssOleDBGridSelectedColumns.bookmark);
                                }
                            }
                            else {
                                dataCollection = frmOriginalDefinition.elements;
                                if (dataCollection != null) {
                                    for (iIndex = 0; iIndex < dataCollection.length; iIndex++) {
                                        sControlName = dataCollection.item(iIndex).name;
                                        sControlName = sControlName.substr(0, 20);
                                        if (sControlName == "txtReportDefnColumn_") {
                                            iTableID = window.dialogArguments.window.selectedColumnParameter(dataCollection.item(iIndex).value, "TABLEID");
                                            if (iTableID == frmUseful.txtCurrentChildTableID.value) {
                                                dataCollection.item(iIndex).value = "";
                                            }
                                        }
                                    }
                                }
                            }

                            // Remove the child table's columns from the sort order collection.
                            window.dialogArguments.window.removeSortColumn(0, frmUseful.txtCurrentChildTableID.value);
                        }
                    }
                }
            }
            frmUseful.txtChanged.value = 1;
        }

        window.dialogArguments.window.refreshTab2Controls();
        frmUseful.txtTablesChanged.value = 1;
        //TM 24/07/02 Fault 4215
        frmDefinition.ssOleDBGridSelectedColumns.selbookmarks.removeall();
        return true;
    }

    function setForm() {
        var frmPopup = document.getElementById("frmPopup");
        var sALL_RECORDS = "All Records";
        var frmChild = window.dialogArguments.OpenHR.getForm("workframe", "frmCustomReportChilds");
        var frmDef = window.dialogArguments.OpenHR.getForm("workframe", "frmDefinition");
        var plngRow = frmDef.ssOleDBGridChildren.AddItemRowIndex(frmDef.ssOleDBGridChildren.Bookmark);
        var frmSelectionAccess = document.getElementById("frmSelectionAccess");
		
        //Add the child tableid and child tablename to the srting.
        var sAdd = frmPopup.cboChildTable.options(frmPopup.cboChildTable.options.selectedIndex).value
            + '	' + frmPopup.cboChildTable.options(frmPopup.cboChildTable.options.selectedIndex).text;

        //Add the filterid and the filter name to the string.
        sAdd = sAdd + '	' + frmPopup.txtChildFilterID.value + '	' + frmPopup.txtChildFilter.value;

        //Add the orderid and the order name to the string.
        sAdd = sAdd + '	' + frmPopup.txtChildFieldOrderID.value + '	' + frmPopup.txtFieldRecOrder.value;

        //Add the max records values to the string.
        if (frmPopup.txtChildRecords.value == 0) {
            sAdd = sAdd + '	' + sALL_RECORDS;
        }
        else {
            sAdd = sAdd + '	' + frmPopup.txtChildRecords.value;
        }

        sAdd = sAdd + '	' + frmSelectionAccess.childHidden.value;

        if (frmChild.childAction.value.toUpperCase() == "NEW") {
            frmDef.ssOleDBGridChildren.additem(sAdd);
            frmDef.ssOleDBGridChildren.selbookmarks.RemoveAll();
            frmDef.ssOleDBGridChildren.MoveLast();
            frmDef.ssOleDBGridChildren.selbookmarks.Add(frmDef.ssOleDBGridChildren.Bookmark);
        }
        else {
            if (frmPopup.originalChildID.value != frmPopup.cboChildTable.options(frmPopup.cboChildTable.options.selectedIndex).value) {
                //' Check if any columns in the report definition are from the table that was
                //' previously selected in the child combo box. If so, prompt user for action.
                var bContinueRemoval;

                bContinueRemoval = removeChildTable(frmPopup.originalChildID.value);

                if (bContinueRemoval) {
                    frmDef.ssOleDBGridChildren.removeitem(plngRow);
                    frmDef.ssOleDBGridChildren.additem(sAdd, plngRow);
                    frmDef.ssOleDBGridChildren.Bookmark = frmDef.ssOleDBGridChildren.AddItemBookmark(plngRow);
                    frmDef.ssOleDBGridChildren.SelBookmarks.RemoveAll();
                    frmDef.ssOleDBGridChildren.SelBookmarks.Add(frmDef.ssOleDBGridChildren.AddItemBookmark(plngRow));
                }
            }
            else {
                frmDef.ssOleDBGridChildren.removeitem(plngRow);
                frmDef.ssOleDBGridChildren.additem(sAdd, plngRow);
                frmDef.ssOleDBGridChildren.Bookmark = frmDef.ssOleDBGridChildren.AddItemBookmark(plngRow);
                frmDef.ssOleDBGridChildren.SelBookmarks.RemoveAll();
                frmDef.ssOleDBGridChildren.SelBookmarks.Add(frmDef.ssOleDBGridChildren.AddItemBookmark(plngRow));
            }
        }

        self.close();
        return false;
    }
</script>

<body <%=session("BodyColour")%> leftmargin="20" topmargin="20" bottommargin="20" rightmargin="5">
<form id="frmPopup" name="frmPopup" onsubmit="return setForm();">
<table align=center class="outline" cellpadding=5 cellspacing=0>
	<tr>
		<td>
			<table class="invisible" cellspacing="0" cellpadding="0">
				<tr height=10> 
					<td height=10 colspan=5 align=center> 
	                    <H3>Select Child Table</H3>
					</td>
				</tr>
				<tr height=10>
					<td width=20>&nbsp;</td>
					<td nowrap>Table :</td>
					<td width=20>&nbsp;</td>
					<td>
						<select id=cboChildTable name=cboChildTable style="WIDTH: 100%" class="combo" onchange=changeChildTable();>
						</select>			
					</td>
					<td width=20>&nbsp;</td>
				</tr>
				<tr height=5> 
					<td colspan="5"></td>
				</tr>
				<tr height=10>
					<td width=20>&nbsp;</td>
					<td nowrap>Filter :</td>
					<td width=20>&nbsp;</td>
					<td>
						<table class="invisible" CELLSPACING=0 CELLPADDING=0 width="100%">
							<TR>
								<TD>
									<input id="txtChildFilter" name="txtChildFilter" class="text textdisabled" disabled="disabled" style="width: 100%" />
									<input type="hidden" id="txtChildFilterID" name="txtChildFilterID" readonly style="width: 100%" />
								</TD>
								<TD width=30>
									<input id="cmdChildFilter" name="cmdChildFilter" style="width: 100%" type="button" value="..." class="btn" 
									   onclick="selectRecordOption('child', 'filter')" onmouseover="try{button_onMouseOver(this);}catch(e){}"
										onmouseout="try{button_onMouseOut(this);}catch(e){}" onfocus="try{button_onFocus(this);}catch(e){}"
										onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
							</TR>
						</table>
					</td>
					<td width=20>&nbsp;</td>
				</tr>
				<tr height=5> 
					<td colspan="5"></td>
				</tr>
				<tr height=10>
					<td width=20>&nbsp;</td>
					<td nowrap>Order :</td>
					<td width=20>&nbsp;</td>
					<td>
						<table class="invisible" CELLSPACING=0 CELLPADDING=0 width="100%">
							<TR>
								<TD>
									<input id=txtFieldRecOrder name=txtFieldRecOrder class="text textdisabled" disabled="disabled" style="WIDTH: 100%">
									<input type=hidden id=txtChildFieldOrderID name=txtChildFieldOrderID readonly style="WIDTH: 100%">  
								</TD>
								<TD width=30>
									<input id=cmdChildOrder name=cmdChildOrder style="WIDTH: 100%" type=button value="..." class="btn"
									    onclick="selectRecordOrder();"  
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
							</TR>
						</table>
					</td>
					<td width=20>&nbsp;</td>
				</tr>
				<tr height=5> 
					<td colspan="5"></td>
				</tr>
				<tr>
					<td width=20>&nbsp;</td>
					<td nowrap>Records :</td>
					<td width=20>&nbsp;</td>
					<td>
						<table WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD>
									<input id=txtChildRecords name=txtChildRecords style="WIDTH: 60px" width="40" value="0" class="text"
									    onkeyup="setRecordsNumeric();showAllRecords();" 
									    onchange="setRecordsNumeric();showAllRecords();">
								</TD>
								<TD width=15>
									<input style="WIDTH: 100%" type="button" value="+" id="cmdChildRecordsUp" name="cmdChildRecordsUp" class="btn"
									    onclick="spinRecords(true);setRecordsNumeric();showAllRecords();"
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
								<TD width=15>
									<input style="WIDTH: 100%" type="button" value="-" id="cmdChildRecordsDown" name="cmdChildRecordsDown"  class="btn"
									    onclick="spinRecords(false);setRecordsNumeric();showAllRecords();"
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
								<TD width=5>&nbsp;</TD>
								<TD width=100>
									<input id=txtAllRecords name=txtAllRecords value="(All Records)" class="textinformation" style ="TEXT-ALIGN: center; WIDTH: 100%" disabled="disabled"/>	
								</TD>
								<TD width=5>&nbsp;</TD>
							</TR>
						</table>					
					</td>
					<td width=20>&nbsp;</td>
				</tr>

				<tr height=20> 
					<td colspan="5"></td>
				</tr>  
				<tr> 
					<td colspan="4"> 
						<table WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD>&nbsp;</TD>
								<TD width=10>
									<input id="cmdOK" type="button" value="OK" name="cmdOK" style="width: 80px" width="80"
										class="btn" onclick="setForm()" onmouseover="try{button_onMouseOver(this);}catch(e){}"
										onmouseout="try{button_onMouseOut(this);}catch(e){}" onfocus="try{button_onFocus(this);}catch(e){}"
										onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
								<TD width=10>&nbsp;</TD>
								<TD width=10>
									<input id="cmdCancel" type="button" value="Cancel" name="cmdCancel" style="width: 80px"
										width="80" class="btn" onclick="self.close();" onmouseover="try{button_onMouseOver(this);}catch(e){}"
										onmouseout="try{button_onMouseOut(this);}catch(e){}" onfocus="try{button_onFocus(this);}catch(e){}"
										onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
							</tr>
						</table>
					</td>
				</tr>
				<tr height="10"> 
					<td colspan="5"></td>
				</tr> 
			</table>
        </td>
	</tr>
</table>
<input type="hidden" id="rowID" name="rowID" />
<input type="hidden" id="originalChildID" name="originalChildID" />
</form>

	<form id="frmRecOrder" name="frmRecOrder" target="RecOrder" action="fieldRec" method="post" style="visibility: hidden; display: none">
		<input type="hidden" id="selectionType" name="selectionType"/>
		<input type="hidden" id="txtTableID" name="txtTableID"/>
		<input type="hidden" id="selectedID" name="selectedID"/>
	</form>

	<form id="frmRecordSelection" name="frmRecordSelection" target="recordSelection" action="util_recordSelection" method="post" style="visibility: hidden; display: none">
		<input type="hidden" id="recSelType" name="recSelType"/>
		<input type="hidden" id="recSelTableID" name="recSelTableID"/>
		<input type="hidden" id="recSelCurrentID" name="recSelCurrentID"/>
		<input type="hidden" id="recSelTable" name="recSelTable"/>
		<input type="hidden" id="recSelDefOwner" name="recSelDefOwner"/>
		<input type="hidden" id="recSelDefType" name="recSelDefType"/>
	</form>

	<form id="frmSelectionAccess" name="frmSelectionAccess" style="visibility: hidden;display: none">
		<input type="hidden" id="baseHidden" name="baseHidden" value="N"/>
		<input type="hidden" id="p1Hidden" name="p1Hidden" value="N"/>
		<input type="hidden" id="p2Hidden" name="p2Hidden" value="N"/>
		<input type="hidden" id="childHidden" name="childHidden" value=""/>
		<input type="hidden" id="calcsHiddenCount" name="calcsHiddenCount" value="0"/>
	</form>

</body>
</html>
