
    function optiondata_onload() {

    		var frmOptionData = document.getElementById("frmOptionData");

        var sFatalErrorMsg = frmOptionData.txtErrorDescription.value;
        if (sFatalErrorMsg.length > 0) {
            //window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sFatalErrorMsg);
            //window.parent.location.replace("login.asp");
        } else {
            // Do nothing if the menu controls are not yet instantiated.
            var sCurrentWorkPage = OpenHR.currentWorkPage();
	        var sErrorMsg;
	        var sAction;
	        var dataCollection;
	        var sControlName;
	        var sColumnName;
	        var iColumnType;
	        var iCount;
	        var sAddString;
	        var fRecordAdded;
	        if (sCurrentWorkPage == "LINKFIND") {
		        sErrorMsg = frmOptionData.txtErrorMessage.value;
		        if (sErrorMsg.length > 0) {
                    // We've got an error so don't update the record edit form.

                    // Get menu.asp to refresh the menu.
                    menu_refreshMenu();                        
                    OpenHR.messageBox(sErrorMsg);
                }
		        sAction = frmOptionData.txtOptionAction.value; // Refresh the link find grid with the data if required.
                var grdLinkFind = OpenHR.getForm("optionframe","frmLinkFindForm").ssOleDBGridLinkRecords;

                grdLinkFind.redraw = false;
                grdLinkFind.removeAll();
                grdLinkFind.columns.removeAll();
		        dataCollection = frmOptionData.elements; // Configure the grid columns.
                if (dataCollection != null) {
                    for (i = 0; i < dataCollection.length; i++) {
                        sControlName = dataCollection.item(i).name;
                        sControlName = sControlName.substr(0, 16);
                        if (sControlName == "txtOptionColDef_") {
                            // Get the column name and type from the control.
                            sColDef = dataCollection.item(i).value;

                            iIndex = sColDef.indexOf("	");
                            if (iIndex >= 0) {
                                sColumnName = sColDef.substr(0, iIndex);
                                sColumnType = sColDef.substr(iIndex + 1);

                                grdLinkFind.columns.add(grdLinkFind.columns.count);
                                grdLinkFind.columns.item(grdLinkFind.columns.count - 1).name = sColumnName;
                                grdLinkFind.columns.item(grdLinkFind.columns.count - 1).caption = sColumnName;

                                if (sColumnName == "ID") {
                                    grdLinkFind.columns.item(grdLinkFind.columns.count - 1).Visible = false;
                                }

                                if ((sColumnType == "131") || (sColumnType == "3")) {
                                    grdLinkFind.columns.item(grdLinkFind.columns.count - 1).Alignment = 1;
                                } else {
                                    grdLinkFind.columns.item(grdLinkFind.columns.count - 1).Alignment = 0;
                                }
                                if (sColumnType == 11) {
                                    grdLinkFind.columns.item(grdLinkFind.columns.count - 1).Style = 2;
                                } else {
                                    grdLinkFind.columns.item(grdLinkFind.columns.count - 1).Style = 0;
                                }
                            }
                        }
                    }
                }

                // Add the grid records.
		        fRecordAdded = false;
                iCount = 0;
                if (dataCollection != null) {
                    for (i = 0; i < dataCollection.length; i++) {
                        sControlName = dataCollection.item(i).name;
                        sControlName = sControlName.substr(0, 14);
                        if (sControlName == "txtOptionData_") {
                            grdLinkFind.addItem(dataCollection.item(i).value);
                            fRecordAdded = true;
	                        iCount = iCount + 1;
                        }
                    }
                }
                grdLinkFind.redraw = true;

                frmOptionData.txtRecordCount.value = iCount;

                if (fRecordAdded == true) {
                    locateRecord(OpenHR.getForm("optionframe","frmLinkFindForm").txtOptionLinkRecordID.value, true); //should be in scope!
                }

                refreshControls();  ///should be in scope - from lookupFind.ascx
                    

                // Get menu.asp to refresh the menu.
                menu_refreshMenu();
            }
	        var grdFind;
	        if (sCurrentWorkPage == "LOOKUPFIND") {
	            sErrorMsg = frmOptionData.txtErrorMessage.value;
	            if (sErrorMsg.length > 0) {
                    // We've got an error so don't update the record edit form.

                    // Get menu.asp to refresh the menu.
                    menu_refreshMenu();
                    OpenHR.messageBox(sErrorMsg);
                }

                if (frmOptionData.txtFilterOverride.value == "True")
                    // No access to the lookup filter column?
                {
                    OpenHR.messageBox("You do not have 'read' permission on the lookup filter value column. No filter will be applied.");
                }
	            sAction = frmOptionData.txtOptionAction.value;
	            OpenHR.getForm("optionframe", "frmLookupFindForm").txtLookupColumnGridPosition.value = frmOptionData.txtLookupColumnGridPosition.value;

                // Refresh the link find grid with the data if required.
	            grdFind = OpenHR.getForm("optionframe", "frmLookupFindForm").ssOleDBGrid; // Clear the grid.
		        
                grdFind.redraw = false;
                grdFind.removeAll();
                grdFind.columns.removeAll();
                dataCollection = frmOptionData.elements; // Configure the grid columns.
                var sColumnType;
		        
                if (dataCollection != null) {
                    for (var i = 0; i < dataCollection.length; i++) {
                        sControlName = dataCollection.item(i).name;
                        sControlName = sControlName.substr(0, 16);
                        if (sControlName == "txtOptionColDef_") {
                            // Get the column name and type from the control.
                            var sColDef = dataCollection.item(i).value;

                            var iIndex = sColDef.indexOf("	");
                            if (iIndex >= 0) {
                                sColumnName = sColDef.substr(0, iIndex);
                                sColumnType = sColDef.substr(iIndex + 1);

                                grdFind.columns.add(grdFind.columns.count);
                                grdFind.columns.item(grdFind.columns.count - 1).name = sColumnName;
                                grdFind.columns.item(grdFind.columns.count - 1).caption = sColumnName;

                                if (sColumnName == "ID") {
                                    grdFind.columns.item(grdFind.columns.count - 1).Visible = false;
                                }

                                if ((sColumnType == "131") || (sColumnType == "3")) {
                                    grdFind.columns.item(grdFind.columns.count - 1).Alignment = 1;
                                } else {
                                    grdFind.columns.item(grdFind.columns.count - 1).Alignment = 0;
                                }
                                if (sColumnType == 11) {
                                    grdFind.columns.item(grdFind.columns.count - 1).Style = 2;
                                } else {
                                    grdFind.columns.item(grdFind.columns.count - 1).Style = 0;
                                }
                            }
                        }
                    }
                }

                // Add the grid records.
								fRecordAdded = false;
                iCount = 0;
                if (dataCollection != null) {
                    for (i = 0; i < dataCollection.length; i++) {
                        sControlName = dataCollection.item(i).name;
                        sControlName = sControlName.substr(0, 14);
                        if (sControlName == "txtOptionData_") {
                            grdFind.addItem(dataCollection.item(i).value);
                            fRecordAdded = true;
	                        iCount = iCount + 1;
                        }
                    }
                }

                grdFind.redraw = true;

                frmOptionData.txtRecordCount.value = iCount;

                if (fRecordAdded == true) {
                    locateRecord(OpenHR.getForm("optionframe","frmLookupFindForm").txtOptionLookupValue.value, true);
                }

                refreshControls();

                // Get menu.asp to refresh the menu.
                menu_refreshMenu();
            }

            if ((sCurrentWorkPage == "TBTRANSFERCOURSEFIND") ||
                (sCurrentWorkPage == "TBBOOKCOURSEFIND") ||
                (sCurrentWorkPage == "TBADDFROMWAITINGLISTFIND") ||
                (sCurrentWorkPage == "TBTRANSFERBOOKINGFIND")) {
	            sErrorMsg = frmOptionData.txtErrorMessage.value;
	            if (sErrorMsg.length > 0) {
                    // We've got an error.
                    // Get menu.asp to refresh the menu.
                    menu_refreshMenu();
                    OpenHR.messageBox(sErrorMsg);
                }

                if ((sCurrentWorkPage == "TBTRANSFERBOOKINGFIND") ||
                    (sCurrentWorkPage == "TBADDFROMWAITINGLISTFIND")) {
	                sErrorMsg = frmOptionData.txtErrorMessage2.value;
	                if (sErrorMsg.length > 0) {
                        // We've got an error.
                        Cancel(); //should be in scope!
                    	//window.parent.frames("menuframe").ASRIntranetFunctions.ClosePopup();
                        OpenHR.messageBox(sErrorMsg);
                        return;
                    }
                }
	            sAction = frmOptionData.txtOptionAction.value; // Refresh the link find grid with the data if required.
	            grdFind = OpenHR.getForm("optionframe", "frmtbFindForm").ssOleDBGridRecords;
	            grdFind.redraw = false;
                grdFind.removeAll();
                grdFind.columns.removeAll();
	            dataCollection = frmOptionData.elements; // Configure the grid columns.
                if (dataCollection != null) {
                    for (i = 0; i < dataCollection.length; i++) {
                        sControlName = dataCollection.item(i).name;
                        sControlName = sControlName.substr(0, 16);
                        if (sControlName == "txtOptionColDef_") {
                            // Get the column name and type from the control.
                            sColDef = dataCollection.item(i).value;

                            iIndex = sColDef.indexOf("	");
                            if (iIndex >= 0) {
                                sColumnName = sColDef.substr(0, iIndex);
                                sColumnType = sColDef.substr(iIndex + 1);

                                grdFind.columns.add(grdFind.columns.count);
                                grdFind.columns.item(grdFind.columns.count - 1).name = sColumnName;
                                grdFind.columns.item(grdFind.columns.count - 1).caption = sColumnName;

                                if (sColumnName == "ID") {
                                    grdFind.columns.item(grdFind.columns.count - 1).Visible = false;
                                }

                                if ((sColumnType == "131") || (sColumnType == "3")) {
                                    grdFind.columns.item(grdFind.columns.count - 1).Alignment = 1;
                                } else {
                                    grdFind.columns.item(grdFind.columns.count - 1).Alignment = 0;
                                }
                                if (sColumnType == 11) {
                                    grdFind.columns.item(grdFind.columns.count - 1).Style = 2;
                                } else {
                                    grdFind.columns.item(grdFind.columns.count - 1).Style = 0;
                                }
                            }
                        }
                    }
                }

                // Add the grid records.
	            fRecordAdded = false;
                iCount = 0;
                if (dataCollection != null) {
                    for (i = 0; i < dataCollection.length; i++) {
                        sControlName = dataCollection.item(i).name;
                        sControlName = sControlName.substr(0, 14);
                        if (sControlName == "txtOptionData_") {
                            grdFind.addItem(dataCollection.item(i).value);
                            fRecordAdded = true;
	                        iCount = iCount + 1;
                        }
                    }
                }

                grdFind.redraw = true;

                frmOptionData.txtRecordCount.value = iCount;

                // Select the top record.
                if (fRecordAdded == true) {
                    grdFind.MoveFirst();
                    grdFind.SelBookmarks.Add(grdFind.Bookmark);
                }

                tbrefreshControls();

                // Get menu.asp to refresh the menu.
                menu_refreshMenu();
            }

            if (sCurrentWorkPage == "TBBULKBOOKING") {
	            
            	 frmOptionData = document.getElementById("frmOptionData");
	            
                sAction = frmOptionData.txtOptionAction.value;

                // Refresh the link find grid with the data if required.
                grdFind = OpenHR.getForm("optionframe","frmBulkBooking").ssOleDBGridFindRecords;
                grdFind.redraw = false;
                grdFind.removeAll();

                dataCollection = frmOptionData.elements;

                // Add the grid records.
                fRecordAdded = false;
                iCount = 0;

                if (dataCollection != null) {
                    for (i = 0; i < dataCollection.length; i++) {
                        sControlName = dataCollection.item(i).name;
                        sControlName = sControlName.substr(0, 14);

                        if (sControlName == "txtOptionData_") {
                            grdFind.addItem(dataCollection.item(i).value);
                            fRecordAdded = true;
	                        iCount = iCount + 1;
                        }
                    }
                }

                grdFind.redraw = true;

                // Select the top record.
                if (fRecordAdded == true) {
                    grdFind.MoveFirst();
                    grdFind.SelBookmarks.Add(grdFind.Bookmark);
                }

                tbrefreshControls();

                // Get menu.asp to refresh the menu.
                menu_refreshMenu();
            }

            if (sCurrentWorkPage == "UTIL_DEF_PICKLIST") {
	            sAction = frmOptionData.txtOptionAction.value; // Refresh the link find grid with the data if required.
	            grdFind = OpenHR.getForm("workframe","frmDefinition").ssOleDBGrid;
	            grdFind.redraw = false;
                grdFind.removeAll();
	            dataCollection = frmOptionData.elements; // Add the grid records.
	            fRecordAdded = false;
                iCount = 0;

                if (dataCollection != null) {
                    for (i = 0; i < dataCollection.length; i++) {
                        sControlName = dataCollection.item(i).name;
                        sControlName = sControlName.substr(0, 14);

                        if (sControlName == "txtOptionData_") {
                            grdFind.addItem(dataCollection.item(i).value);
                            fRecordAdded = true;
                            iCount = iCount + 1
                        }
                    }
                }

                grdFind.redraw = true;

                if (frmOptionData.txtExpectedCount.value > iCount) {
                    if (iCount == 0) {
                        OpenHR.messageBox("You do not have 'read' permission on any of the records in the selected picklist.\nUnable to display records.");
                        OpenHR.getForm("workframe","frmUseful").txtAction.value = "VIEW";
                        OpenHR.getFrame("workframe").cancelClick();
                    } else {
                        if (OpenHR.getForm("workframe","frmUseful").txtAction.value.toUpperCase() == "COPY") {
                            OpenHR.messageBox("You do not have 'read' permission on all of the records in the selected picklist.\nOnly permitted records will be shown.");
                        } else {
                            OpenHR.messageBox("You do not have 'read' permission on all of the records in the selected picklist.\nOnly permitted records will be shown and the definition will be read only.");
                            OpenHR.getForm("workframe","frmUseful").txtAction.value = "VIEW";
                            OpenHR.getFrame("workframe").disableAll();
                        }
                    }
                }

                // Select the top record.
                if (fRecordAdded == true) {
                    grdFind.MoveFirst();
                    grdFind.SelBookmarks.Add(grdFind.Bookmark);
                }

                refreshControls();

                // Get menu.asp to refresh the menu.
                menu_refreshMenu();
            }

            if (sCurrentWorkPage == "UTIL_DEF_EXPRCOMPONENT") {
	            sAction = frmOptionData.txtOptionAction.value;
	            if ((sAction == "LOADEXPRFIELDCOLUMNS") ||
                    (sAction == "LOADEXPRLOOKUPCOLUMNS")) {
	                dataCollection = frmOptionData.elements;
	                if (dataCollection != null) {
                        for (i = 0; i < dataCollection.length; i++) {
                            sControlName = dataCollection.item(i).name;
                            sControlName = sControlName.substr(0, 10);
                            if (sControlName == "txtColumn_") {
                                component_addColumn(dataCollection.item(i).value);
                            }
                        }
                    }

                    component_setColumn(frmOptionData.txtOptionColumnID.value);
                }

                if (sAction == "LOADEXPRLOOKUPVALUES") {
	                dataCollection = frmOptionData.elements;
	                if (dataCollection != null) {
                        for (i = 0; i < dataCollection.length; i++) {
                            sControlName = dataCollection.item(i).name;
                            sControlName = sControlName.substr(0, 9);
                            if (sControlName == "txtValue_") {
                                component_addValue(dataCollection.item(i).value);
                            }
                        }
                    }

                    component_setValue(frmOptionData.txtOptionLocateValue.value);
                }

                // Get menu.asp to refresh the menu.
                menu_refreshMenu();
            }

            if (sCurrentWorkPage == "FIND") {
	            sAction = frmOptionData.txtOptionAction.value;
	            if ((sAction == "BOOKCOURSEERROR") ||
                    (sAction == "TRANSFERBOOKINGERROR") ||
                    (sAction == "ADDFROMWAITINGLISTERROR") ||
                    (sAction == "BULKBOOKINGERROR")) {
                    OpenHR.messageBox(frmOptionData.txtNonFatalErrorDescription.value);
                }
                if ((sAction == "BOOKCOURSESUCCESS") ||
                    (sAction == "TRANSFERBOOKINGSUCCESS") ||
                    (sAction == "ADDFROMWAITINGLISTSUCCESS") ||
                    (sAction == "BULKBOOKINGSUCCESS")) {
                    // Reload the find records.
                    OpenHR.messageBox("Booking(s) made successfully.");
                    menu_reloadFindPage("MOVEFIRST", "");
                }
            }
            
        }
    }   

function refreshOptionData() {
	    var frmGetOptionData = document.getElementById("frmGetOptionData");
	    OpenHR.submitForm(frmGetOptionData);		
	}
