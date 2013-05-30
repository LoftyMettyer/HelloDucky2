
//todo remove this function!
//New functionality - get the selected row's record ID from the hidden tag		
function getRecordID(rowID) {
	return $("#findGridTable").find("#" + rowID + " input[type=hidden]").val();
}

function rowCount() {
    return $("#findGridTable tr").length - 1;
}

function bookmarksCount() {
    var selRowIds = $('#findGridTable').jqGrid('getGridParam', 'selarrrow');
    return selRowIds.length;
}

function moveFirst() {
    $("#findGridTable").jqGrid('setSelection', 1);
}


function find_window_onload() {
    var fOk;

    fOk = true;
    $("#workframe").attr("data-framesource", "FIND");
    $("#optionframe").hide();
    $("#workframe").show();

    $(function () {
        tableToGrid("#findGridTable", {
            onSelectRow: function (rowID) {
                //menu_refreshMenu();
            },
            ondblClickRow: function (rowID) {
                menu_editRecord();
            },
            rowNum: 1000    //TODO set this to blocksize...
        });
    });

    $("#findGridTable").jqGrid('bindKeys', {
        "onEnter": function (rowid) {
            menu_editRecord();
        }
    });

    //resize the grid to the height of its container.
    $("#findGridTable").jqGrid('setGridHeight', $("#findGridRow").height());
    var y = $("#gbox_findGridTable").height();
    var z = $('#gbox_findGridTable .ui-jqgrid-bdiv').height();

    var frmFindForm = document.getElementById("frmFindForm");

    var sErrMsg = frmFindForm.txtErrorDescription.value;
    if (sErrMsg.length > 0) {
        fOk = false;

        //				window.parent.frames("menuframe").ASRIntranetFunctions.Closepopup();
        OpenHR.messageBox(sErrMsg);
        menu_loadPage("default");
    }

    if (fOk == true) {

        //TODO: check the font settings.
        //setGridFont(frmFindForm.ssOleDBGridFindRecords);

        var frmMenuInfo = document.getElementById("frmMenuInfo");

        if ($("#workframe").length = 0) { //only check if not in SSI mode.
            if ((frmMenuInfo.txtUserType.value == 1) &&
                (frmMenuInfo.txtPersonnel_EmpTableID.value == frmFindForm.txtCurrentTableID.value) &&
                (frmFindForm.txtRecordCount.value > 1)) {

                $("#findGridTable").focus();
                $("#findGridTable").html = ""; //empty the grid

                // Get menu.asp to refresh the menu.
                menu_refreshMenu();

                /* The user does NOT have permission to create new records. */
                OpenHR.messageBox("Unable to load personnel records.\n\nYou are logged on as a self-service user and can access only single record personnel record sets.");

                /* Go to the default page. */
                menu_loadPage("default");
                return;
            }
        }
    }

    if (fOk == true) {
        var sControlName;
        var sControlPrefix;
        var sColumnId;
        var ctlSummaryControl;
        var sSummaryControlName;
        var sDataType;

        // Expand the work frame and hide the option frame.
        //window.parent.document.all.item("workframeset").cols = "*, 0";

        //moved this higher up - grid wasn't resized properly after filtering. (HRPRO-2797)
        //$("#workframe").attr("data-framesource", "FIND");
        //$("#optionframe").hide();
        //$("#workframe").show();

        // JPD20020903 Fault 2316 - Need to dim focus on the grid before adding the items.
        $("#findGridTable").focus();

        var controlCollection = frmFindForm.elements;
        if (controlCollection != null) {
            for (var i = 0; i < controlCollection.length; i++) {

                sControlName = controlCollection.item(i).name;
                sControlPrefix = sControlName.substr(0, 13);

                if (sControlPrefix == "txtAddString_") {
                    //frmFindForm.ssOleDBGridFindRecords.AddItem(controlCollection.item(i).value);
                }

                sControlName = controlCollection.item(i).name;
                sControlPrefix = sControlName.substr(0, 15);

                if (sControlPrefix == "txtSummaryData_") {
                    sColumnId = sControlName.substr(15);
                    sSummaryControlName = "ctlSummary_";
                    sSummaryControlName = sSummaryControlName.concat(sColumnId);
                    sSummaryControlName = sSummaryControlName.concat("_");

                    for (var j = 0; j < controlCollection.length; j++) {
                        sControlName = controlCollection.item(j).name;
                        sControlPrefix = sControlName.substr(0, sSummaryControlName.length);

                        if (sControlPrefix == sSummaryControlName) {
                            ctlSummaryControl = controlCollection.item(j);

                            if (ctlSummaryControl.type == "checkbox") {
                                ctlSummaryControl.checked = (controlCollection.item(i).value.toUpperCase() == "TRUE");
                            } else {
                                // Check if the control is for a datevalue.
                                sDataType = sControlName.substr(sSummaryControlName.length);

                                if (sDataType == "11") {
                                    // Format dates for the locale setting.							
                                    if (controlCollection.item(i).value == '') {
                                        ctlSummaryControl.value = '';
                                    } else {
                                        //TODO:ctlSummaryControl.value = window.parent.frames("menuframe").ASRIntranetFunctions.ConvertSQLDateToLocale(controlCollection.item(i).value);
                                        ctlSummaryControl.value = controlCollection.item(i).value;
                                    }
                                } else {
                                    ctlSummaryControl.value = controlCollection.item(i).value;
                                }
                            }

                            break;
                        }
                    }
                }
            }
        }

        //// dim focus onto one of the form controls. 
        //// NB. This needs to be done before making any reference to the grid
        //frmFindForm.ssOleDBGridFindRecords.focus();


        // Select the current record in the grid if its there, else select the top record if there is one.
        if (rowCount() > 0) {
            if ((frmFindForm.txtCurrentRecordID.value > 0) && (frmFindForm.txtGotoAction.value != 'LOCATE')) {
                // Try to select the current record.
                locateRecord(frmFindForm.txtCurrentRecordID.value, true);
            } else {
                // Select the top row.
                //frmFindForm.ssOleDBGridFindRecords.MoveFirst();					    
                //frmFindForm.ssOleDBGridFindRecords.SelBookmarks.Add(frmFindForm.ssOleDBGridFindRecords.Bookmark);
                moveFirst();
            }
        }

        // Get menu.asp to refresh the menu.	    		
        menu_refreshMenu();

        if ((rowCount() == 0) && (frmFindForm.txtFilterSQL.value.length > 0)) {
            OpenHR.messageBox("No records match the current filter.\nNo filter is applied.");
            menu_clearFilter();
        }
    }
}


/* Return the ID of the record selected in the find form. */
function selectedRecordID() {
	        
    var iRecordId;
			
    iRecordId = $("#findGridTable").getGridParam('selrow');
    iRecordId = getRecordID(iRecordId);
	        
    return (iRecordId);
}
	    

/* Sequential search the grid for the required ID. */
function locateRecord(psSearchFor, pfIdMatch) {
    var fFound;
    var iIndex;
    var iIdColumnIndex;
    var sColumnName;

    var frmFindForm = document.getElementById("frmFindForm");

		//select the grid row that contains the record with the passed in ID.
    var rowNumber = $("#findGridTable input[value='" + psSearchFor + "']").parent().parent().attr("id");
    if (rowNumber >= 0) {
    	$("#findGridTable").jqGrid('setSelection', rowNumber);
    } else {
    	$("#findGridTable").jqGrid('setSelection', 1);
    }
	


    //fFound = false;
    //frmFindForm.ssOleDBGridFindRecords.redraw = false;

    //if (pfIdMatch == true) {
    //    // Locate the ID column in the grid.
    //    iIdColumnIndex = -1;
    //    for (iIndex = 0; iIndex < frmFindForm.ssOleDBGridFindRecords.Cols; iIndex++) {
    //        sColumnName = frmFindForm.ssOleDBGridFindRecords.Columns(iIndex).Name;
    //        if (sColumnName.toUpperCase() == "ID") {
    //            iIdColumnIndex = iIndex;
    //            break;
    //        }
    //    }

    //    if (iIdColumnIndex >= 0) {
    //        frmFindForm.ssOleDBGridFindRecords.MoveLast();
    //        frmFindForm.ssOleDBGridFindRecords.MoveFirst();

    //        for (iIndex = 1; iIndex <= frmFindForm.ssOleDBGridFindRecords.rows; iIndex++) {
    //            if (frmFindForm.ssOleDBGridFindRecords.Columns(iIdColumnIndex).value == psSearchFor) {
    //                frmFindForm.ssOleDBGridFindRecords.FirstRow = frmFindForm.ssOleDBGridFindRecords.Bookmark;
    //                if ((frmFindForm.ssOleDBGridFindRecords.Rows - frmFindForm.ssOleDBGridFindRecords.AddItemRowIndex(frmFindForm.ssOleDBGridFindRecords.FirstRow) + 1) < frmFindForm.ssOleDBGridFindRecords.VisibleRows) {
    //                    if (frmFindForm.ssOleDBGridFindRecords.Rows - frmFindForm.ssOleDBGridFindRecords.VisibleRows + 1 >= 1) {
    //                        frmFindForm.ssOleDBGridFindRecords.FirstRow = frmFindForm.ssOleDBGridFindRecords.AddItemBookmark(frmFindForm.ssOleDBGridFindRecords.Rows - frmFindForm.ssOleDBGridFindRecords.VisibleRows + 1);
    //                    }
    //                    else {
    //                        frmFindForm.ssOleDBGridFindRecords.FirstRow = frmFindForm.ssOleDBGridFindRecords.AddItemBookmark(0);
    //                    }
    //                }

    //                frmFindForm.ssOleDBGridFindRecords.SelBookmarks.Add(frmFindForm.ssOleDBGridFindRecords.Bookmark);
    //                fFound = true;
    //                break;
    //            }

    //            if (iIndex < frmFindForm.ssOleDBGridFindRecords.rows) {
    //                frmFindForm.ssOleDBGridFindRecords.MoveNext();
    //            }
    //            else {
    //                break;
    //            }
    //        }
    //    }
    //}
    //else {
    //    for (iIndex = 1; iIndex <= frmFindForm.ssOleDBGridFindRecords.rows; iIndex++) {
    //        var sGridValue = new String(frmFindForm.ssOleDBGridFindRecords.Columns(0).value);
    //        sGridValue = sGridValue.substr(0, psSearchFor.length).toUpperCase();
    //        if (sGridValue == psSearchFor.toUpperCase()) {
    //            frmFindForm.ssOleDBGridFindRecords.SelBookmarks.Add(frmFindForm.ssOleDBGridFindRecords.Bookmark);
    //            fFound = true;
    //            break;
    //        }

    //        if (iIndex < frmFindForm.ssOleDBGridFindRecords.rows) {
    //            frmFindForm.ssOleDBGridFindRecords.MoveNext();
    //        }
    //        else {
    //            break;
    //        }
    //    }
    //}

    //if ((fFound == false) && (frmFindForm.ssOleDBGridFindRecords.rows > 0)) {
    //    // Select the top row.
    //    frmFindForm.ssOleDBGridFindRecords.MoveFirst();
    //    frmFindForm.ssOleDBGridFindRecords.SelBookmarks.Add(frmFindForm.ssOleDBGridFindRecords.Bookmark);
    //}

    //frmFindForm.ssOleDBGridFindRecords.redraw = true;
}
	
