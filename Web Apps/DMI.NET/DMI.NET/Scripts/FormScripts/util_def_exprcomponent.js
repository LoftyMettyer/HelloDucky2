
    
    function util_def_exprcomponent_onload() {
        
        var fOK;
        var sUDFFunction_Visibility;
        var sUDFFunction_Display;

        fOK = true;	

        setGridFont(frmMainForm.ssOleDBGridCalculations);
        setGridFont(frmMainForm.ssOleDBGridFilters);
        setTreeFont(frmMainForm.SSFunctionTree);
        setTreeFont(frmMainForm.SSOperatorTree);
	
        if (fOK == true) {
            // Expand the option frame and hide the work frame.
            $("#optionframe").attr("data-framesource", "UTIL_DEF_EXPRCOMPONENT");
        
            $("#workframe").hide();        
            $("#optionframe").show();
        
            formatComponentTypeFrame();
		
            if (util_def_exprcomponent_frmUseful.txtAction.value == "EDITEXPRCOMPONENT") {
                // Load the component definition.
                loadDefinition();
            }
            else {
                // Inserting/adding a new component.
                frmMainForm.optType_Field.checked = true;
                changeType(1);
            }	


            sUDFFunction_Visibility = "visible";
            sUDFFunction_Display = "block";
		
            if (frmMainForm.txtPassByType.value == 1) {
                frmMainForm.optFieldRecSel_Specific.style.visibility = sUDFFunction_Visibility;
                frmMainForm.optFieldRecSel_Specific.style.display = sUDFFunction_Display;
                frmMainForm.txtFieldRecSel_Specific.style.visibility = sUDFFunction_Visibility;
                frmMainForm.txtFieldRecSel_Specific.style.display = sUDFFunction_Display;
                divFieldRecSel_Specific.style.visibility = sUDFFunction_Visibility;
                divFieldRecSel_Specific.style.display = sUDFFunction_Display;
            }
		
            // Set focus onto one of the form controls. 
            frmMainForm.cmdCancel.focus();

            // Hide the workframe ActiveX treeview. IE6 still displays it.
            OpenHR.getForm("workframe","frmDefinition").SSTree1.style.visibility = "hidden";

        }
    }

function formatComponentTypeFrame() {
    var sType_PVal_Visibility;
    var sType_PVal_Display;
    var sType_Calc_Visibility;
    var sType_Calc_Display;
    var sType_Filter_Visibility;
    var sType_Filter_Display;

    sType_PVal_Visibility = "visible";
    sType_PVal_Display = "block";
    sType_Calc_Visibility = "visible";
    sType_Calc_Display = "block";
    sType_Filter_Visibility = "visible";
    sType_Filter_Display = "block";

    switch (util_def_exprcomponent_frmUseful.txtExprType.value) {
        case "10":
            // Runtime Calculation
            break;
        case "11":
            // Runtime Filter
            break;
        case "14":
            // Utility Runtime Calculation - no calcs, no filters, no prompted values
            sType_PVal_Visibility = "hidden";
            sType_PVal_Display = "none";

            sType_Calc_Visibility = "hidden";
            sType_Calc_Display = "none";

            sType_Filter_Visibility = "hidden";
            sType_Filter_Display = "none";
            break;
    }

    trType_PVal.style.visibility = sType_PVal_Visibility;
    trType_PVal.style.display = sType_PVal_Display;
    trType_PVal2.style.visibility = sType_PVal_Visibility;
    trType_PVal2.style.display = sType_PVal_Display;
	
    trType_Calc.style.visibility = sType_Calc_Visibility;
    trType_Calc.style.display = sType_Calc_Display;
    trType_Calc2.style.visibility = sType_Calc_Visibility;
    trType_Calc2.style.display = sType_Calc_Display;
	
    trType_Filter.style.visibility = sType_Filter_Visibility;
    trType_Filter.style.display = sType_Filter_Display;
}

function loadDefinition() {
    var iType;
    var i;
    var iIndex;

    iType = new Number(util_def_exprcomponent_frmOriginalDefinition.txtType.value);
	
    if (iType == 1) {
        // Field
        frmMainForm.optType_Field.checked = true;
        changeType(iType);

        util_def_exprcomponent_frmUseful.txtInitialising.value = 0;
		
        if (frmMainForm.txtPassByType.value == 1) {
            if (util_def_exprcomponent_frmOriginalDefinition.txtFieldSelectionRecord.value == 5) {
                frmMainForm.optField_Count.checked = true;	
            }
            else {
                if (util_def_exprcomponent_frmOriginalDefinition.txtFieldSelectionRecord.value == 4) {
                    frmMainForm.optField_Total.checked = true;	
                }
                else {
                    frmMainForm.optField_Field.checked = true;	
					
                    if (util_def_exprcomponent_frmOriginalDefinition.txtFieldSelectionRecord.value == 3) {
                        frmMainForm.optFieldRecSel_Specific.checked = true;	
                        frmMainForm.txtFieldRecSel_Specific.value = util_def_exprcomponent_frmOriginalDefinition.txtFieldSelectionLine.value;
                    }
                    else {
                        if (util_def_exprcomponent_frmOriginalDefinition.txtFieldSelectionRecord.value == 2) {
                            frmMainForm.optFieldRecSel_Last.checked = true;	
                        }
                        else {
                            frmMainForm.optFieldRecSel_First.checked = true;	
                        }
                    }
                }
            }
        }

        field_refreshTable();
		
        if (frmMainForm.txtPassByType.value == 1) {
            util_def_exprcomponent_frmUseful.txtChildFieldOrderID.value = util_def_exprcomponent_frmOriginalDefinition.txtFieldSelectionOrderID.value;
            frmMainForm.txtFieldRecOrder.value = util_def_exprcomponent_frmOriginalDefinition.txtFieldOrderName.value;
            util_def_exprcomponent_frmUseful.txtChildFieldFilterID.value = util_def_exprcomponent_frmOriginalDefinition.txtFieldSelectionFilter.value;
            frmMainForm.txtFieldRecFilter.value = util_def_exprcomponent_frmOriginalDefinition.txtFieldFilterName.value;
        }
        else {
            util_def_exprcomponent_frmUseful.txtChildFieldOrderID.value = 0;
            util_def_exprcomponent_frmUseful.txtChildFieldFilterID.value = 0;
        }
    }
	
    if (iType == 2) {
        // Function
        frmMainForm.optType_Function.checked = true;
        changeType(iType);

        util_def_exprcomponent_frmUseful.txtInitialising.value = 0;
        functionAndOperator_refresh();
        frmMainForm.SSFunctionTree.SelectedItem = frmMainForm.SSFunctionTree.Nodes(util_def_exprcomponent_frmOriginalDefinition.txtFunctionID.value);
        button_disable(frmMainForm.cmdOK, false);
    }
	
    if (iType == 3) {
        // Calculation.
        frmMainForm.optType_Calculation.checked = true;
        changeType(iType);
		
        util_def_exprcomponent_frmUseful.txtInitialising.value = 0;
        calculationsAndFilters_load();
		
        // Locate the current calc.
        locateGridRecord(util_def_exprcomponent_frmOriginalDefinition.txtCalculationID.value);
    }
	
    if (iType == 4) {
        // Value
        frmMainForm.optType_Value.checked = true;
        changeType(iType);

        if (util_def_exprcomponent_frmOriginalDefinition.txtValueType.value == 1) {
            // Character value
            frmMainForm.cboValueType.selectedIndex = 0;
            value_changeType();
            frmMainForm.txtValue.value = util_def_exprcomponent_frmOriginalDefinition.txtValueCharacter.value;
        }
        if (util_def_exprcomponent_frmOriginalDefinition.txtValueType.value == 2) {
            // Numeric value
            frmMainForm.cboValueType.selectedIndex = 1;
            value_changeType();
            frmMainForm.txtValue.value = util_def_exprcomponent_frmOriginalDefinition.txtValueNumeric.value;
        }
        if (util_def_exprcomponent_frmOriginalDefinition.txtValueType.value == 3) {
            // Logic value
            frmMainForm.cboValueType.selectedIndex = 2;
            value_changeType();
            if (util_def_exprcomponent_frmOriginalDefinition.txtValueLogic.value == 1) {
                frmMainForm.selectValue.selectedIndex = 0;
            }
            else {
                frmMainForm.selectValue.selectedIndex = 1;
            }
        }
        if (util_def_exprcomponent_frmOriginalDefinition.txtValueType.value == 4) {
            // Date value
            frmMainForm.cboValueType.selectedIndex = 3;
            value_changeType();
            frmMainForm.txtValue.value = menu_ConvertSQLDateToLocale(util_def_exprcomponent_frmOriginalDefinition.txtValueDate.value);
        }
    }

    if (iType == 5) {
        // Operator
        frmMainForm.optType_Operator.checked = true;
        changeType(iType);

        util_def_exprcomponent_frmUseful.txtInitialising.value = 0;
        functionAndOperator_refresh();
        frmMainForm.SSOperatorTree.SelectedItem = frmMainForm.SSOperatorTree.Nodes(util_def_exprcomponent_frmOriginalDefinition.txtOperatorID.value);
        button_disable(frmMainForm.cmdOK, false);
    }

    if (iType == 6) {
        // Lookup Table Value
        frmMainForm.optType_LookupTableValue.checked = true;
        changeType(iType);

        util_def_exprcomponent_frmUseful.txtInitialising.value = 0;
        lookupValue_refreshTable();
    }

    if (iType == 7) {
        // Prompted Value.
        frmMainForm.optType_PromptedValue.checked = true;
        changeType(iType);
		
        frmMainForm.txtPrompt.value = util_def_exprcomponent_frmOriginalDefinition.txtPromptDescription.value;
        frmMainForm.cboPValType.selectedIndex = util_def_exprcomponent_frmOriginalDefinition.txtValueType.value - 1;
        frmMainForm.txtPValSize.value = util_def_exprcomponent_frmOriginalDefinition.txtPromptSize.value;
        frmMainForm.txtPValDecimals.value = util_def_exprcomponent_frmOriginalDefinition.txtPromptDecimals.value;
        frmMainForm.txtPValFormat.value = util_def_exprcomponent_frmOriginalDefinition.txtPromptMask.value;

        pVal_changeType();
		
        if (util_def_exprcomponent_frmOriginalDefinition.txtValueType.value == 1) {
            // Character
            frmMainForm.txtPValDefault.value = util_def_exprcomponent_frmOriginalDefinition.txtValueCharacter.value;
        }		
        if (util_def_exprcomponent_frmOriginalDefinition.txtValueType.value == 2) {
            // Numeric
            frmMainForm.txtPValDefault.value = util_def_exprcomponent_frmOriginalDefinition.txtValueNumeric.value;
        }		
        if (util_def_exprcomponent_frmOriginalDefinition.txtValueType.value == 3) {
            // Logic
            iIndex = 0;
            for (i=0; i<frmMainForm.cboPValDefault.options.length; i++)  {
                if (frmMainForm.cboPValDefault.options(i).Value == util_def_exprcomponent_frmOriginalDefinition.txtValueLogic.value) {
                    iIndex = i;
                    break;
                }		
            }
	
            frmMainForm.cboPValDefault.selectedIndex = iIndex;
        }		
		
        if (util_def_exprcomponent_frmOriginalDefinition.txtValueType.value == 4) {
            // Date
            if (util_def_exprcomponent_frmOriginalDefinition.txtPromptDateType.value == 0) {
                if ((util_def_exprcomponent_frmOriginalDefinition.txtValueDate.value != "12/30/1899") &&
                    (util_def_exprcomponent_frmOriginalDefinition.txtValueDate.value != "")) {
                    frmMainForm.txtPValDefault.value = menu_ConvertSQLDateToLocale(util_def_exprcomponent_frmOriginalDefinition.txtValueDate.value);
                }
                frmMainForm.optPValDate_Explicit.checked = true;
            }
            if (util_def_exprcomponent_frmOriginalDefinition.txtPromptDateType.value == 2) {
                frmMainForm.optPValDate_MonthStart.checked = true;
            }
            if (util_def_exprcomponent_frmOriginalDefinition.txtPromptDateType.value == 1) {
                frmMainForm.optPValDate_Current.checked = true;
            }
            if (util_def_exprcomponent_frmOriginalDefinition.txtPromptDateType.value == 4) {
                frmMainForm.optPValDate_YearStart.checked = true;
            }
            if (util_def_exprcomponent_frmOriginalDefinition.txtPromptDateType.value == 3) {
                frmMainForm.optPValDate_MonthEnd.checked = true;
            }
            if (util_def_exprcomponent_frmOriginalDefinition.txtPromptDateType.value == 5) {
                frmMainForm.optPValDate_YearEnd.checked = true;
            }

            pVal_changeDateOption(util_def_exprcomponent_frmOriginalDefinition.txtPromptDateType.value)
        }		
    }

    if (iType == 10) {
        // Filter.
        frmMainForm.optType_Filter.checked = true;
        changeType(iType);
		
        util_def_exprcomponent_frmUseful.txtInitialising.value = 0;
        calculationsAndFilters_load();
		
        // Locate the current filter.
        locateGridRecord(util_def_exprcomponent_frmOriginalDefinition.txtFilterID.value);
    }
}

function changeType(piType) {
    var sField_Visibility;
    var sField_Display;
    var sFunction_Visibility;
    var sFunction_Display;
    var sOperator_Visibility;
    var sOperator_Display;
    var sValue_Visibility;
    var sValue_Display;
    var sLookupValue_Visibility;
    var sLookupValue_Display;
    var sCalculation_Visibility;
    var sCalculation_Display;
    var sFilter_Visibility;
    var sFilter_Display;
    var sPromptedValue_Visibility;
    var sPromptedValue_Display;
	
    sField_Visibility = "hidden";
    sField_Display = "none";
    sFunction_Visibility = "hidden";
    sFunction_Display = "none";
    sOperator_Visibility = "hidden";
    sOperator_Display = "none";
    sValue_Visibility = "hidden";
    sValue_Display = "none";
    sLookupValue_Visibility = "hidden";
    sLookupValue_Display = "none";
    sCalculation_Visibility = "hidden";
    sCalculation_Display = "none";
    sFilter_Visibility = "hidden";
    sFilter_Display = "none";
    sPromptedValue_Visibility = "hidden";
    sPromptedValue_Display = "none";

    if (piType == 1) {
        // Field
        sField_Visibility = "visible";
        sField_Display = "block";
    }
    if (piType == 2) {
        // Function
        sFunction_Visibility = "visible";
        sFunction_Display = "block";
    }
    if (piType == 3) {
        // Calculation
        sCalculation_Visibility = "visible";
        sCalculation_Display = "block";
    }
    if (piType == 4) {
        // Value
        sValue_Visibility = "visible";
        sValue_Display = "block";
    }
    if (piType == 5) {
        // Operator
        sOperator_Visibility = "visible";
        sOperator_Display = "block";
    }
    if (piType == 6) {
        // Table Value
        sLookupValue_Visibility = "visible";
        sLookupValue_Display = "block";
    }
    if (piType == 7) {
        // Prompted Value
        sPromptedValue_Visibility = "visible";
        sPromptedValue_Display = "block";
    }
    if (piType == 10) {
        // Filter
        sFilter_Visibility = "visible";
        sFilter_Display = "block";
    }

    divField.style.visibility = sField_Visibility;
    divField.style.display = sField_Display;
    divFunction.style.visibility = sFunction_Visibility;
    divFunction.style.display = sFunction_Display;
    divOperator.style.visibility = sOperator_Visibility;
    divOperator.style.display = sOperator_Display;
    divValue.style.visibility = sValue_Visibility;
    divValue.style.display = sValue_Display;
    divLookupValue.style.visibility = sLookupValue_Visibility;
    divLookupValue.style.display = sLookupValue_Display;
    divCalculation.style.visibility = sCalculation_Visibility;
    divCalculation.style.display = sCalculation_Display;
    divFilter.style.visibility = sFilter_Visibility;
    divFilter.style.display = sFilter_Display;
    divPromptedValue.style.visibility = sPromptedValue_Visibility;
    divPromptedValue.style.display = sPromptedValue_Display;

    initializeComponentControls(piType);
}

function initializeComponentControls(piType) {
    button_disable(frmMainForm.cmdOK, false);

    switch (piType) {
        case 1:
            // Field
            if (frmMainForm.txtPassByType.value == 1) {
                frmMainForm.optField_Field.checked = true;
            }
            util_def_exprcomponent_frmUseful.txtInitialising.value = 1;
            field_refreshTable();			
            break;
        case 2:
            // Function
            util_def_exprcomponent_frmUseful.txtInitialising.value = 1;
            functionAndOperator_refresh();
            break;
        case 3:
            // Calculation
            util_def_exprcomponent_frmUseful.txtInitialising.value = 1;
            calculationAndFilter_refresh();
            break;
        case 4:
            // Value
            frmMainForm.cboValueType.selectedIndex = 0;
            value_changeType();
            break;
        case 5:
            // Operator
            util_def_exprcomponent_frmUseful.txtInitialising.value = 1;
            functionAndOperator_refresh();
            break;
        case 6:
            // Table Value
            util_def_exprcomponent_frmUseful.txtInitialising.value = 1;
            lookupValue_refreshTable();			
            break;
        case 7:
            // Prompted Value
            frmMainForm.cboPValType.selectedIndex = 0;
            frmMainForm.txtPrompt.value = "";
            pVal_changeType();
            break;
        case 10:
            // Filter
            util_def_exprcomponent_frmUseful.txtInitialising.value = 1;
            calculationAndFilter_refresh();
            break;
    }			
}

function field_refreshTable() {
    var oOption;
    var sTableName;
    var sTableID;
    var sRelated;
    var sIsChild;
    var tableCollection = frmTables.elements;
    var sDefaultTableID;
    var i;
    var iIndex;
    var fInitialise = util_def_exprcomponent_frmUseful.txtInitialising.value;
    var fTableOK;
	
    sDefaultTableID = util_def_exprcomponent_frmOriginalDefinition.txtFieldTableID.value;

    if ((fInitialise == 0) &&
        (frmMainForm.cboFieldTable.selectedIndex >= 0)) {
        sDefaultTableID = frmMainForm.cboFieldTable.options[frmMainForm.cboFieldTable.selectedIndex].Value;
    }

    // Clear the current contents of the dropdown list.
    while (frmMainForm.cboFieldTable.options.length > 0) {
        frmMainForm.cboFieldTable.options.remove(0);
    }

    if (tableCollection != null) {
        for (i=0; i<tableCollection.length; i++)  {
            fTableOK = (frmMainForm.txtPassByType.value == 2);
            if (fTableOK == false) {
                sRelated = tableParameter(tableCollection.item(i).value, "RELATED");
                sIsChild = tableParameter(tableCollection.item(i).value, "ISCHILD");

                fTableOK =  (((sRelated == "1") && (frmMainForm.optField_Field.checked)) || 
                ((sIsChild == "1") && ((frmMainForm.optField_Count.checked) || (frmMainForm.optField_Total.checked))));
            }
			
            if (fTableOK == true) {
                sTableName = tableParameter(tableCollection.item(i).value, "NAME");
                sTableID = tableParameter(tableCollection.item(i).value, "TABLEID");
                oOption = document.createElement("OPTION");
                frmMainForm.cboFieldTable.options.add(oOption);
                oOption.innerText = sTableName;
                oOption.Value = sTableID;			
            }
        }
    }	

    if (frmMainForm.cboFieldTable.options.length > 0) {
        iIndex = 0;
        for (i=0; i<frmMainForm.cboFieldTable.options.length; i++)  {
            if (frmMainForm.cboFieldTable.options(i).Value == sDefaultTableID) {
                iIndex = i;
                break;
            }		
			
            if (frmMainForm.cboFieldTable.options(i).Value == util_def_exprcomponent_frmUseful.txtTableID.value) {
                iIndex = i;
            }		
        }

        frmMainForm.cboFieldTable.selectedIndex = iIndex;
	
        combo_disable(frmMainForm.cboFieldTable, false);
    }
    else {
        combo_disable(frmMainForm.cboFieldTable, true);
    }

    field_refreshColumn();
}

function field_refreshColumn() {

    var sDefaultColumnID;
    var fInitialise = util_def_exprcomponent_frmUseful.txtInitialising.value;
	
    sDefaultColumnID = util_def_exprcomponent_frmOriginalDefinition.txtFieldColumnID.value;

    if ((fInitialise == 0) &&
        (frmMainForm.cboFieldColumn.selectedIndex >= 0)) {
        sDefaultColumnID = frmMainForm.cboFieldColumn.options[frmMainForm.cboFieldColumn.selectedIndex].Value;
    }

    if (frmMainForm.txtPassByType.value == 2) {
        frmMainForm.cboFieldColumn.style.visibility = "visible";
        frmMainForm.cboFieldColumn.style.display= "block";
        frmMainForm.cboFieldDummyColumn.style.visibility = "hidden"
        frmMainForm.cboFieldDummyColumn.style.display = "none";
    }
    else {	
        if (frmMainForm.optField_Count.checked == true) {
            frmMainForm.cboFieldColumn.style.visibility = "hidden";
            frmMainForm.cboFieldColumn.style.display= "none";
            frmMainForm.cboFieldDummyColumn.style.visibility = "visible"
            frmMainForm.cboFieldDummyColumn.style.display = "block";
        }
        else {
            frmMainForm.cboFieldColumn.style.visibility = "visible";
            frmMainForm.cboFieldColumn.style.display= "block";
            frmMainForm.cboFieldDummyColumn.style.visibility = "hidden"
            frmMainForm.cboFieldDummyColumn.style.display = "none";
        }
    }
	
    // Clear the current contents of the dropdown list.
    while (frmMainForm.cboFieldColumn.options.length > 0) {
        frmMainForm.cboFieldColumn.options.remove(0);
    }

    if (frmMainForm.cboFieldTable.selectedIndex >= 0) {
        // Get the optionData page to get the columns for the current table.
        var optionDataForm = OpenHR.getForm("optiondataframe","frmGetOptionData");
        optionDataForm.txtOptionAction.value = "LOADEXPRFIELDCOLUMNS";
        optionDataForm.txtOptionTableID.value = frmMainForm.cboFieldTable.options[frmMainForm.cboFieldTable.selectedIndex].Value;
        optionDataForm.txtOptionColumnID.value = sDefaultColumnID;

        if (frmMainForm.txtPassByType.value == 2) {
            optionDataForm.txtOptionOnlyNumerics.value = 0;
        }
        else {
            if (frmMainForm.optField_Total.checked == true) {
                optionDataForm.txtOptionOnlyNumerics.value = 1;
            }
            else {
                optionDataForm.txtOptionOnlyNumerics.value = 0;
            }
        }
           
        //OpenHR.getFrame("optiondataframe").refreshOptionData();
        refreshOptionData();

    }
    else {
        // No table selected. Clear the column combo.
        field_refreshChildFrame();
    }
}

function field_refreshChildFrame() {
    var tableCollection = frmTables.elements;
    var i;
    var fIsChild;
	
    if (frmMainForm.txtPassByType.value == 1) {
        fIsChild = false;

        if (frmMainForm.cboFieldTable.selectedIndex >= 0) {
            if (tableCollection != null) {
                for (i=0; i<tableCollection.length; i++) {
                    if (tableParameter(tableCollection.item(i).value, "TABLEID") == frmMainForm.cboFieldTable.options[frmMainForm.cboFieldTable.selectedIndex].Value) {
                        if (tableParameter(tableCollection.item(i).value, "ISCHILD") == 1) {
                            fIsChild = true;
                        }
                    }
                }	
            }
        }

        if (fIsChild == true) {
            // Enable the child record selection controls.
            if (frmMainForm.optField_Field.checked == true) {
                if ((frmMainForm.optFieldRecSel_Last.checked == false) &&
                    (frmMainForm.optFieldRecSel_Specific.checked == false)) {
                    frmMainForm.optFieldRecSel_First.checked = true;
                }
				
                radio_disable(frmMainForm.optFieldRecSel_First, false);
                radio_disable(frmMainForm.optFieldRecSel_Last, false);
                radio_disable(frmMainForm.optFieldRecSel_Specific, false);
				
                if (frmMainForm.optFieldRecSel_Specific.checked == false) {
                    frmMainForm.txtFieldRecSel_Specific.value = "";
                    text_disable(frmMainForm.txtFieldRecSel_Specific, true);
                }
                else {
                    if (frmMainForm.txtFieldRecSel_Specific.value == "") {
                        frmMainForm.txtFieldRecSel_Specific.value = "1";
                    }

                    text_disable(frmMainForm.txtFieldRecSel_Specific, false);
                }

                button_disable(frmMainForm.btnFieldRecOrder, false);
            }
            else {
                frmMainForm.optFieldRecSel_First.checked = true;

                radio_disable(frmMainForm.optFieldRecSel_First, true);
                radio_disable(frmMainForm.optFieldRecSel_Last, true);
                radio_disable(frmMainForm.optFieldRecSel_Specific, true);
                frmMainForm.txtFieldRecSel_Specific.value = "";
                text_disable(frmMainForm.txtFieldRecSel_Specific, true);

                button_disable(frmMainForm.btnFieldRecOrder, true);
                frmMainForm.txtFieldRecOrder.value = "";
                util_def_exprcomponent_frmUseful.txtChildFieldOrderID.value = 0;
            }
			
            button_disable(frmMainForm.btnFieldRecFilter, false);
        }
        else {
            // Disable the child record selection controls.
            frmMainForm.optFieldRecSel_First.checked = true;
			
            radio_disable(frmMainForm.optFieldRecSel_First, true);
            radio_disable(frmMainForm.optFieldRecSel_Last, true);
            radio_disable(frmMainForm.optFieldRecSel_Specific, true);
			
            frmMainForm.txtFieldRecSel_Specific.value = "";
            text_disable(frmMainForm.txtFieldRecSel_Specific, true);
            button_disable(frmMainForm.btnFieldRecOrder, true);
            frmMainForm.txtFieldRecOrder.value = "";
            util_def_exprcomponent_frmUseful.txtChildFieldOrderID.value = 0;
            button_disable(frmMainForm.btnFieldRecFilter, true);
            frmMainForm.txtFieldRecFilter.value = "";
            util_def_exprcomponent_frmUseful.txtChildFieldFilterID.value = 0;
        }
    }
		
    if (frmMainForm.cboFieldColumn.selectedIndex < 0) {
        combo_disable(frmMainForm.cboFieldColumn, true);
    }
    else {
        combo_disable(frmMainForm.cboFieldColumn, false);
    }

    if ((frmMainForm.cboFieldTable.selectedIndex < 0) ||
        (frmMainForm.cboFieldColumn.selectedIndex < 0)) {
        button_disable(frmMainForm.cmdOK, true);
    }
    else {
        button_disable(frmMainForm.cmdOK, false);
    }
	
}

function field_changeTable() {
    field_refreshColumn();

    if (frmMainForm.txtPassByType.value == 1) {
        frmMainForm.txtFieldRecFilter.value = "";
        frmMainForm.txtFieldRecOrder.value = "";
    }
		
    util_def_exprcomponent_frmUseful.txtChildFieldFilterID.value = 0;
    util_def_exprcomponent_frmUseful.txtChildFieldFilterHidden.value = "N";
    util_def_exprcomponent_frmUseful.txtChildFieldOrderID.value = 0;
}

function field_selectRecOrder() {
    var sURL;

    frmFieldRec.selectionType.value = "ORDER";
    frmFieldRec.txtTableID.value = frmMainForm.cboFieldTable.options[frmMainForm.cboFieldTable.selectedIndex].Value;
    frmFieldRec.selectedID.value = util_def_exprcomponent_frmUseful.txtChildFieldOrderID.value;
	
    sURL = "fieldRec" +
        "?selectionType=" + escape(frmFieldRec.selectionType.value) +
        "&txtTableID=" + escape(frmFieldRec.txtTableID.value) +
        "&selectedID=" + escape(frmFieldRec.selectedID.value);
    openDialog(sURL, (screen.width)/3,(screen.height)/2, "yes", "yes");
}

function field_selectRecFilter() {
    var sURL;

    frmFieldRec.selectionType.value = "FILTER";
    frmFieldRec.txtTableID.value = frmMainForm.cboFieldTable.options[frmMainForm.cboFieldTable.selectedIndex].Value;
    frmFieldRec.selectedID.value = util_def_exprcomponent_frmUseful.txtChildFieldFilterID.value;
	
    sURL = "fieldRec" +
        "?selectionType=" + escape(frmFieldRec.selectionType.value) +
        "&txtTableID=" + escape(frmFieldRec.txtTableID.value) +
        "&selectedID=" + escape(frmFieldRec.selectedID.value);
    openDialog(sURL, (screen.width)/3,(screen.height)/2, "yes", "yes");
}

function functionAndOperator_refresh() {
    // Load the function treeview with the functions.
    var i;
    var objNode;
    var colCollection;
    var sName;
    var sID;
    var sCategory;
    var fCategoryDone;
    var iLoop;
    var trvTreeView;
    var sRootKey;
    var sRootText;
    var ctlLoadedFlag;
    var fNoOperatorSelected;

    if (frmMainForm.optType_Function.checked == true) {
        trvTreeView = frmMainForm.SSFunctionTree;
        colCollection = frmFunctions.elements;
        sRootKey = "FUNCTION_ROOT";
        sRootText = "Functions";
        ctlLoadedFlag = util_def_exprcomponent_frmUseful.txtFunctionsLoaded;
    }
    else {
        trvTreeView = frmMainForm.SSOperatorTree;
        colCollection = frmOperators.elements;
        sRootKey = "OPERATOR_ROOT";
        sRootText = "Operators";
        ctlLoadedFlag = util_def_exprcomponent_frmUseful.txtOperatorsLoaded;
    }
	
    if (ctlLoadedFlag.value == 0) {
        // Clear the treeview.
        trvTreeView.Nodes.Clear();

        // Create the root node.
        objNode = trvTreeView.Nodes.Add();
        objNode.key = sRootKey;
        objNode.text = sRootText;
        objNode.font.Bold = true;
        objNode.expanded = true;
        objNode.sorted = 1;

        if (colCollection != null) {
            for (i=0; i<colCollection.length; i++)  {
                sName = functionAndOperatorParameter(colCollection.item(i).value, "NAME");
                sID = functionAndOperatorParameter(colCollection.item(i).value, "ID");
                sCategory = functionAndOperatorParameter(colCollection.item(i).value, "CATEGORY");

                // Add a category node if required.
                fCategoryDone = false;
                for (iLoop=1; iLoop<=trvTreeView.Nodes.Count; iLoop++)  {
                    if (trvTreeView.Nodes(iLoop).Key == sCategory) {
                        fCategoryDone = true;
                        break;
                    }
                }

                if (fCategoryDone == false) {
                    objNode = trvTreeView.Nodes.Add(sRootKey, 4, sCategory, sCategory)
                    objNode.Font.Bold = true;
                    objNode.Sorted = 1;
                }

                // Add the function node.
                objNode = trvTreeView.Nodes.Add(sCategory, 4, sID, sName)
            }
        }	
	
        ctlLoadedFlag.value = 1;
    }
	
    // Enable the treeview only if there are items.
    if (trvTreeView.Nodes.Count > 0) {
        trvTreeView.Enabled = true;
        trvTreeView.focus();
		
        fNoOperatorSelected = true;
        if (trvTreeView.selectedNodes.count > 0) {
            if ((trvTreeView.selectedItem.key !=	"OPERATOR_ROOT") &&
                (trvTreeView.selectedItem.key !=	"FUNCTION_ROOT")) {
                if ((trvTreeView.selectedItem.Parent.Key != "OPERATOR_ROOT") &&
                    (trvTreeView.selectedItem.Parent.Key != "FUNCTION_ROOT")) {
                    fNoOperatorSelected = false;
                }
            }
        }

        button_disable(frmMainForm.cmdOK, fNoOperatorSelected);
    }
    else {
        trvTreeView.Enabled = false;
        button_disable(frmMainForm.cmdOK, true);
    }

    util_def_exprcomponent_frmUseful.txtInitialising.value = 0;

}

function calculationAndFilter_refresh() {
    var iCurrentID;
    var grdGrid;
	
    iCurrentID = 0;
	
    if (frmMainForm.optType_Filter.checked == true) {
        grdGrid = frmMainForm.ssOleDBGridFilters;
    }
    else {
        grdGrid = frmMainForm.ssOleDBGridCalculations;
    }
	
    if (grdGrid.SelBookmarks.Count > 0) {
        iCurrentID = grdGrid.Columns("id").Value;
    }
	
    calculationsAndFilters_load();

    if (grdGrid.Rows > 0) {
        if (util_def_exprcomponent_frmUseful.txtInitialising.value == 1) {
            // Goto top record.
            grdGrid.MoveFirst();
            grdGrid.SelBookmarks.Add(grdGrid.Bookmark);
        }
        else {
            // Locate the current calc/filter if required.
            locateGridRecord(iCurrentID);
        }

        button_disable(frmMainForm.cmdOK, false);
    }
    else {
        button_disable(frmMainForm.cmdOK, true);
    }
	
    frmMainForm.ssOleDBGridCalculations.RowHeight = 19;
    frmMainForm.ssOleDBGridFilters.RowHeight = 19;
	
    util_def_exprcomponent_frmUseful.txtInitialising.value = 0;
}

function calculationsAndFilters_load() {
    // Load the calculations/filters grid with the calcs.
    var i;
    var colCollection;
    var sName;
    var sID;
    var sOwner;
    var iLoop;
    var sAddString;
    var sCurrentOwner = new String(util_def_exprcomponent_frmUseful.txtUserName.value);
    var grdGrid;
    var fOwners;
	
    sCurrentOwner = sCurrentOwner.toUpperCase();
	
    if (frmMainForm.optType_Filter.checked == true) {
        grdGrid = frmMainForm.ssOleDBGridFilters;
        colCollection = frmFilters.elements;
        fOwners = frmMainForm.chkOwnersFilters.checked;
    }
    else {
        grdGrid = frmMainForm.ssOleDBGridCalculations;
        colCollection = frmCalcs.elements;
        fOwners = frmMainForm.chkOwnersCalcs.checked;
    }

    grdGrid.focus();
    grdGrid.Redraw = false;

    // Clear the grid.
    if(grdGrid.Rows > 0) {
        grdGrid.RemoveAll();
    }

    if (colCollection != null) {
        for (i=0; i<colCollection.length; i++)  {
            if (colCollection.item(i).name.indexOf("Desc_") < 0) {
                sName = calculationAndFilterParameter(colCollection.item(i).value, "NAME");
                sID = calculationAndFilterParameter(colCollection.item(i).value, "EXPRID");
                sOwner = calculationAndFilterParameter(colCollection.item(i).value, "OWNER");
			
                sOwner = sOwner.toUpperCase();

                if ((fOwners == false) ||
                    (sCurrentOwner == sOwner)) {
								
                    // Add the grid records.
                    sAddString = sName + "	" + sID;
                    grdGrid.addItem(sAddString);
                }
            }
        }
    }	

    grdGrid.Redraw = true;

}

function value_changeType() {
    if (frmMainForm.cboValueType.options[frmMainForm.cboValueType.selectedIndex].value == 3) {
        frmMainForm.txtValue.style.width = 0;
        frmMainForm.txtValue.style.visibility = "hidden";
        frmMainForm.txtValue.style.position = "absolute";
        frmMainForm.txtValue.style.top = 0;
        frmMainForm.txtValue.style.left = 0;

        frmMainForm.selectValue.style.width = "100%";
        frmMainForm.selectValue.style.visibility = "";
        frmMainForm.selectValue.style.position = "";
        frmMainForm.selectValue.style.top = "";
        frmMainForm.selectValue.style.left = "";
    }
    else {
        frmMainForm.selectValue.style.width = 0;
        frmMainForm.selectValue.style.visibility = "hidden";
        frmMainForm.selectValue.style.position = "absolute";
        frmMainForm.selectValue.style.top = 0;
        frmMainForm.selectValue.style.left = 0;

        frmMainForm.txtValue.style.width = "100%";
        frmMainForm.txtValue.style.visibility = "";
        frmMainForm.txtValue.style.position = "";
        frmMainForm.txtValue.style.top = "";
        frmMainForm.txtValue.style.left = "";
    }

    frmMainForm.txtValue.value = "";
    frmMainForm.selectValue.selectedIndex = 0;
}

function lookupValue_refreshTable() {
    var oOption;
    var sTableName;
    var sTableID;
    var sType;
    var tableCollection = frmTables.elements;
    var sDefaultTableID;
    var i;
    var iIndex;
    var fInitialise = util_def_exprcomponent_frmUseful.txtInitialising.value;
	
    sDefaultTableID = util_def_exprcomponent_frmOriginalDefinition.txtLookupTableID.value;

    if ((fInitialise == 0) &&
        (frmMainForm.cboLookupValueTable.selectedIndex >= 0)) {
        sDefaultTableID = frmMainForm.cboLookupValueTable.options[frmMainForm.cboLookupValueTable.selectedIndex].Value;
    }

    if (util_def_exprcomponent_frmUseful.txtLookupTablesLoaded.value == 0) {
        // Clear the current contents of the dropdown list.
        while (frmMainForm.cboLookupValueTable.options.length > 0) {
            frmMainForm.cboLookupValueTable.options.remove(0);
        }

        if (tableCollection != null) {
            for (i=0; i<tableCollection.length; i++)  {
                sType = tableParameter(tableCollection.item(i).value, "TYPE");
					
                if (sType == "3") {
                    sTableName = tableParameter(tableCollection.item(i).value, "NAME");
                    sTableID = tableParameter(tableCollection.item(i).value, "TABLEID");
                    oOption = document.createElement("OPTION");
                    frmMainForm.cboLookupValueTable.options.add(oOption);
                    oOption.innerText = sTableName;
                    oOption.Value = sTableID;			
                }
            }
        }	
		
        util_def_exprcomponent_frmUseful.txtLookupTablesLoaded.value = 1;
    }
	
    if (frmMainForm.cboLookupValueTable.options.length > 0) {
        iIndex = 0;
        for (i=0; i<frmMainForm.cboLookupValueTable.options.length; i++)  {
            if (frmMainForm.cboLookupValueTable.options(i).Value == sDefaultTableID) {
                iIndex = i;
                break;
            }		
        }

        frmMainForm.cboLookupValueTable.selectedIndex = iIndex;
		
        combo_disable(frmMainForm.cboLookupValueTable, false);
    }
    else {
        combo_disable(frmMainForm.cboLookupValueTable, true);
    }
		
    lookupValue_refreshColumn()
}

function lookupValue_refreshColumn() {
    var sDefaultColumnID;
    var fInitialise = util_def_exprcomponent_frmUseful.txtInitialising.value;
	
    sDefaultColumnID = util_def_exprcomponent_frmOriginalDefinition.txtLookupColumnID.value;

    if ((fInitialise == 0) &&
        (frmMainForm.cboLookupValueColumn.selectedIndex >= 0)) {
        sDefaultColumnID = frmMainForm.cboLookupValueColumn.options[frmMainForm.cboLookupValueColumn.selectedIndex].Value;
    }

    // Clear the current contents of the dropdown list.
    while (frmMainForm.cboLookupValueColumn.options.length > 0) {
        frmMainForm.cboLookupValueColumn.options.remove(0);
    }

    if (frmMainForm.cboLookupValueTable.selectedIndex >= 0) {
        // Get the optionData page to get the columns for the current table.
        var optionDataForm = OpenHR.getForm("optiondataframe","frmGetOptionData");
        optionDataForm.txtOptionAction.value = "LOADEXPRLOOKUPCOLUMNS";
        optionDataForm.txtOptionTableID.value = frmMainForm.cboLookupValueTable.options[frmMainForm.cboLookupValueTable.selectedIndex].Value;
        optionDataForm.txtOptionColumnID.value = sDefaultColumnID;

        //OpenHR.getFrame("optiondataframe").refreshOptionData();
        refreshOptionData();
    }
    else {
        combo_disable(frmMainForm.cboLookupValueColumn, true);

        while (frmMainForm.cboLookupValueValue.options.length > 0) {
            frmMainForm.cboLookupValueValue.options.remove(0);
        }

        combo_disable(frmMainForm.cboLookupValueValue, true);

        button_disable(frmMainForm.cmdOK, true);
    }
}

function lookupValue_refreshValues() {
    var sDefaultValue = "";
    var fInitialise = util_def_exprcomponent_frmUseful.txtInitialising.value;
    var iDataType;

    if (frmMainForm.cboLookupValueColumn.selectedIndex >= 0) {
        iDataType = columnParameter(frmMainForm.cboLookupValueColumn.options[frmMainForm.cboLookupValueColumn.selectedIndex].Value, "DATATYPE");

        if ((frmMainForm.cboLookupValueTable.options[frmMainForm.cboLookupValueTable.selectedIndex].Value == util_def_exprcomponent_frmOriginalDefinition.txtLookupTableID.value)  &&
            (columnParameter(frmMainForm.cboLookupValueColumn.options[frmMainForm.cboLookupValueColumn.selectedIndex].Value, "COLUMNID") == util_def_exprcomponent_frmOriginalDefinition.txtLookupColumnID.value)) {

            if (iDataType == 11) {
                // Date type lookup column.
                sDefaultValue = util_def_exprcomponent_frmOriginalDefinition.txtValueDate.value;
                sDefaultValue = menu_ConvertSQLDateToLocale(sDefaultValue);
            }
            if (iDataType == 12) {
                // Character type lookup column.
                sDefaultValue = util_def_exprcomponent_frmOriginalDefinition.txtValueCharacter.value;
            }
            if ((iDataType == 2) || (iDataType == 4)) {
                // Numeric/integer type lookup column.
                sDefaultValue = util_def_exprcomponent_frmOriginalDefinition.txtValueNumeric.value;
            }
			
            if ((fInitialise == 0) &&
                (frmMainForm.cboLookupValueValue.selectedIndex >= 0)) {
                sDefaultValue = frmMainForm.cboLookupValueValue.options[frmMainForm.cboLookupValueValue.selectedIndex].Value;
            }
        }
    }
			
    // Clear the current contents of the dropdown list.
    while (frmMainForm.cboLookupValueValue.options.length > 0) {
        frmMainForm.cboLookupValueValue.options.remove(0);
    }

    if (frmMainForm.cboLookupValueColumn.selectedIndex >= 0) {
        // Get the optionData page to get the columns for the current table.
        var optionDataForm = OpenHR.getForm("optiondataframe","frmGetOptionData");
        optionDataForm.txtOptionAction.value = "LOADEXPRLOOKUPVALUES";
        optionDataForm.txtOptionColumnID.value = columnParameter(frmMainForm.cboLookupValueColumn.options[frmMainForm.cboLookupValueColumn.selectedIndex].Value, "COLUMNID");
        optionDataForm.txtGotoLocateValue.value = sDefaultValue;
            
        //OpenHR.getFrame("optiondataframe").refreshOptionData();
        refreshOptionData();
    }
    else {
        combo_disable(frmMainForm.cboLookupValueValue, true);

        button_disable(frmMainForm.cmdOK, true);
    }
}

function lookupValue_changeTable() {
    lookupValue_refreshColumn();
}

function lookupValue_changeColumn() {
    util_def_exprcomponent_frmUseful.txtInitialising.value = 1;
    lookupValue_refreshValues();
}

function pVal_changeType() {
    var sSizeVisibility;
    var sDecimalsVisibility;
    var sFormatVisibility;
    var sFormatDisplay;
    var sLookupVisibility;
    var sLookupDisplay;
    var sTextDefaultVisibility;
    var sTextDefaultDisplay;
    var sComboDefaultVisibility;
    var sComboDefaultDisplay;
    var sDateOptionsVisibility;
    var sDateOptionsDisplay;
    var oOption;
    var iPValType = frmMainForm.cboPValType.options[frmMainForm.cboPValType.selectedIndex].value;
	
    sSizeVisibility = "hidden";
    sDecimalsVisibility = "hidden";
    sFormatVisibility = "hidden";
    sFormatDisplay = "none";
    sLookupVisibility = "hidden";
    sLookupDisplay = "none";
    sTextDefaultVisibility = "hidden";
    sTextDefaultDisplay = "none";
    sComboDefaultVisibility = "hidden";
    sComboDefaultDisplay = "none";
    sDateOptionsVisibility = "hidden";
    sDateOptionsDisplay = "none";

    if (iPValType == 1) {
        // Character
        sSizeVisibility = "visible";
        sFormatVisibility = "visible";
        sFormatDisplay = "block";
        sTextDefaultVisibility = "visible";
        sTextDefaultDisplay = "block";
        text_disable(frmMainForm.txtPValDefault, false);		
    }

    if (iPValType == 2) {
        // Numeric
        sSizeVisibility = "visible";
        sDecimalsVisibility = "visible";
        sTextDefaultVisibility = "visible";
        sTextDefaultDisplay = "block";
        text_disable(frmMainForm.txtPValDefault, false);
    }

    if (iPValType == 3) {
        // Logic
        sComboDefaultVisibility = "visible";
        sComboDefaultDisplay = "block";
		
        // Clear the current contents of the dropdown list.
        while (frmMainForm.cboPValDefault.options.length > 0) {
            frmMainForm.cboPValDefault.options.remove(0);
        }

        oOption = document.createElement("OPTION");
        frmMainForm.cboPValDefault.options.add(oOption);
        oOption.innerText = "Yes";
        oOption.Value = 1;			

        oOption = document.createElement("OPTION");
        frmMainForm.cboPValDefault.options.add(oOption);
        oOption.innerText = "No";
        oOption.Value = 0;	
		
        frmMainForm.cboPValDefault.selectedIndex = 0;		
    }

    if (iPValType == 4) {
        // Date
        sTextDefaultVisibility = "visible";
        sTextDefaultDisplay = "block";
        sDateOptionsVisibility = "visible";
        sDateOptionsDisplay = "block";
		
        frmMainForm.optPValDate_Explicit.checked = true;
    }

    if (iPValType == 5) {
        // Lookup Table Value
        sLookupVisibility = "visible";
        sLookupDisplay = "block";
        sComboDefaultVisibility = "visible";
        sComboDefaultDisplay = "block";

        // Clear the current contents of the dropdown list.
        while (frmMainForm.cboPValDefault.options.length > 0) {
            frmMainForm.cboPValDefault.options.remove(0);
        }

        pVal_refreshTable();			
    }

    frmMainForm.txtPValSize.style.visibility = sSizeVisibility;
    tdPValSizePrompt.style.visibility = sSizeVisibility;
    frmMainForm.txtPValDecimals.style.visibility = sDecimalsVisibility;
    tdPValDecimalsPrompt.style.visibility = sDecimalsVisibility;
    trPValFormat.style.visibility = sFormatVisibility;
    trPValFormat.style.display = sFormatDisplay;
    trPValFormat2.style.visibility = sFormatVisibility;
    trPValFormat2.style.display = sFormatDisplay;
    trPValLookup.style.visibility = sLookupVisibility;
    trPValLookup.style.display = sLookupDisplay;
    trPValLookup2.style.visibility = sLookupVisibility;
    trPValLookup2.style.display = sLookupDisplay;
	
    trPValTextDefault.style.visibility = sTextDefaultVisibility;
    trPValTextDefault.style.display = sTextDefaultDisplay;
    trPValComboDefault.style.visibility = sComboDefaultVisibility;
    trPValComboDefault.style.display = sComboDefaultDisplay;
    trPValDateOptions.style.visibility = sDateOptionsVisibility;
    trPValDateOptions.style.display = sDateOptionsDisplay;
    trPValDateOptions2.style.visibility = sDateOptionsVisibility;
    trPValDateOptions2.style.display = sDateOptionsDisplay;

    frmMainForm.txtPValDefault.value = "";
	
    pVal_changePrompt();
}

function pVal_refreshTable() {
    var oOption;
    var sTableName;
    var sTableID;
    var sType;
    var tableCollection = frmTables.elements;
    var sDefaultTableID;
    var i;
    var iIndex;
    var fInitialise = util_def_exprcomponent_frmUseful.txtInitialising.value;
    var sDefaultTableID;
	
    sDefaultTableID = util_def_exprcomponent_frmOriginalDefinition.txtFieldTableID.value;

    if ((fInitialise == 0) &&
        (frmMainForm.cboPValTable.selectedIndex >= 0)) {
        sDefaultTableID = frmMainForm.cboPValTable.options[frmMainForm.cboPValTable.selectedIndex].Value;
    }

    if (util_def_exprcomponent_frmUseful.txtPValLookupTablesLoaded.value == 0) {
        // Clear the current contents of the dropdown list.
        while (frmMainForm.cboPValTable.options.length > 0) {
            frmMainForm.cboPValTable.options.remove(0);
        }

        if (tableCollection != null) {
            for (i=0; i<tableCollection.length; i++)  {
                sType = tableParameter(tableCollection.item(i).value, "TYPE");
					
                if (sType == "3") {
                    sTableName = tableParameter(tableCollection.item(i).value, "NAME");
                    sTableID = tableParameter(tableCollection.item(i).value, "TABLEID");
                    oOption = document.createElement("OPTION");
                    frmMainForm.cboPValTable.options.add(oOption);
                    oOption.innerText = sTableName;
                    oOption.Value = sTableID;			
                }
            }
        }	
		
        util_def_exprcomponent_frmUseful.txtPValLookupTablesLoaded.value = 1;
    }
	
    if (frmMainForm.cboPValTable.options.length > 0) {
        iIndex = 0;
        for (i=0; i<frmMainForm.cboPValTable.options.length; i++)  {
            if (frmMainForm.cboPValTable.options(i).Value == sDefaultTableID) {
                iIndex = i;
                break;
            }		
        }

        frmMainForm.cboPValTable.selectedIndex = iIndex;
        combo_disable(frmMainForm.cboPValTable, false);
    }
    else {
        combo_disable(frmMainForm.cboPValTable, true);
    }
	
    pVal_refreshColumn()
}

function pVal_refreshColumn() {
    var sDefaultColumnID;
    var fInitialise = util_def_exprcomponent_frmUseful.txtInitialising.value;
	
    sDefaultColumnID = util_def_exprcomponent_frmOriginalDefinition.txtFieldColumnID.value;

    if ((fInitialise == 0) &&
        (frmMainForm.cboPValColumn.selectedIndex >= 0)) {
        sDefaultColumnID = frmMainForm.cboPValColumn.options[frmMainForm.cboPValColumn.selectedIndex].Value;
    }

    // Clear the current contents of the dropdown list.
    while (frmMainForm.cboPValColumn.options.length > 0) {
        frmMainForm.cboPValColumn.options.remove(0);
    }

    if (frmMainForm.cboPValTable.selectedIndex >= 0) {
        // Get the optionData page to get the columns for the current table.
        var optionDataForm = OpenHR.getForm("optiondataframe","frmGetOptionData");
        optionDataForm.txtOptionAction.value = "LOADEXPRLOOKUPCOLUMNS";
        optionDataForm.txtOptionTableID.value = frmMainForm.cboPValTable.options[frmMainForm.cboPValTable.selectedIndex].Value;
        optionDataForm.txtOptionColumnID.value = sDefaultColumnID;

        //OpenHR.getFrame("optiondataframe").refreshOptionData();
        refreshOptionData();
    }
    else {
        combo_disable(frmMainForm.cboPValColumn, true);

        while (frmMainForm.cboPValDefault.options.length > 0) {
            frmMainForm.cboPValDefault.options.remove(0);
        }

        combo_disable(frmMainForm.cboPValDefault, true);
        button_disable(frmMainForm.cmdOK, true);
    }
}

function pVal_refreshValues() {
    var sDefaultValue = "";
    var fInitialise = util_def_exprcomponent_frmUseful.txtInitialising.value;
    var iDataType;
	
    if (frmMainForm.cboPValColumn.selectedIndex >= 0) {
        iDataType = columnParameter(frmMainForm.cboPValColumn.options[frmMainForm.cboPValColumn.selectedIndex].Value, "DATATYPE");

        if ((frmMainForm.cboPValTable.options[frmMainForm.cboPValTable.selectedIndex].Value == util_def_exprcomponent_frmOriginalDefinition.txtFieldTableID.value)  &&
            (columnParameter(frmMainForm.cboPValColumn.options[frmMainForm.cboPValColumn.selectedIndex].Value, "COLUMNID") == util_def_exprcomponent_frmOriginalDefinition.txtFieldColumnID.value)) {

            if (iDataType == 11) {
                // Date type lookup column.
                sDefaultValue = util_def_exprcomponent_frmOriginalDefinition.txtValueCharacter.value;
                sDefaultValue = menu_ConvertSQLDateToLocale(sDefaultValue);
            }
            if (iDataType == 12) {
                // Character type lookup column.
                sDefaultValue = util_def_exprcomponent_frmOriginalDefinition.txtValueCharacter.value;
            }
            if ((iDataType == 2) || (iDataType == 4)) {
                // Numeric/integer type lookup column.
                sDefaultValue = util_def_exprcomponent_frmOriginalDefinition.txtValueCharacter.value;
            }
			
            if ((fInitialise == 0) &&
                (frmMainForm.cboPValDefault.selectedIndex >= 0)) {
                sDefaultValue = frmMainForm.cboPValDefault.options[frmMainForm.cboPValDefault.selectedIndex].Value;
            }
        }
    }
	
    // Clear the current contents of the dropdown list.
    while (frmMainForm.cboPValDefault.options.length > 0) {
        frmMainForm.cboPValDefault.options.remove(0);
    }

    if (frmMainForm.cboPValColumn.selectedIndex >= 0) {
        // Get the optionData page to get the columns for the current table.
        var optionDataForm = OpenHR.getForm("optiondataframe","frmGetOptionData");
        optionDataForm.txtOptionAction.value = "LOADEXPRLOOKUPVALUES";
        optionDataForm.txtOptionColumnID.value = columnParameter(frmMainForm.cboPValColumn.options[frmMainForm.cboPValColumn.selectedIndex].Value, "COLUMNID");
        optionDataForm.txtGotoLocateValue.value = sDefaultValue;

        //OpenHR.getFrame("optiondataframe").refreshOptionData();
        refreshOptionData();
    }
    else {
        combo_disable(frmMainForm.cboPValDefault, true);
        button_disable(frmMainForm.cmdOK, true);
    }
}

function pVal_changePrompt() {
    if (frmMainForm.optType_PromptedValue.checked == true) {
        if (frmMainForm.txtPrompt.value.length == 0) {
            button_disable(frmMainForm.cmdOK, true);
        }
        else {
            button_disable(frmMainForm.cmdOK, false);
        }
    }
}

function pVal_changeDateOption(piDateOption) {
    if (piDateOption == 0) {
        text_disable(frmMainForm.txtPValDefault, false);
    }
    else {
        frmMainForm.txtPValDefault.value = "";
        text_disable(frmMainForm.txtPValDefault, true);
    }
}

function pVal_changeTable() {
    pVal_refreshColumn();
}

function pVal_changeColumn() {
    util_def_exprcomponent_frmUseful.txtInitialising.value = 1;
    pVal_refreshValues();
}

function component_addColumn(psDefn) {
    var sColumnName;
    var oOption;
    var cboCombo;
	
    if (frmMainForm.optType_Field.checked == true) {
        cboCombo = frmMainForm.cboFieldColumn;
    }
    else {
        if (frmMainForm.optType_PromptedValue.checked == true) {
            cboCombo = frmMainForm.cboPValColumn;
        }
        else {
            cboCombo = frmMainForm.cboLookupValueColumn;
        }
    }
				
    sColumnName = columnParameter(psDefn, "NAME");
    oOption = document.createElement("OPTION");
    cboCombo.options.add(oOption);
    oOption.innerText = sColumnName;
    oOption.Value = psDefn;			
}

function component_setColumn(piColumnID) {
    var iIndex;
    var i;
    var cboCombo;
	
    if (frmMainForm.optType_Field.checked == true) {
        cboCombo = frmMainForm.cboFieldColumn;
    }
    else {
        if (frmMainForm.optType_PromptedValue.checked == true) {
            cboCombo = frmMainForm.cboPValColumn;
        }
        else {
            cboCombo = frmMainForm.cboLookupValueColumn;
        }
    }

    if (cboCombo.options.length > 0) {
        iIndex = 0;
        for (i=0; i<cboCombo.options.length; i++)  {
            if (columnParameter(cboCombo.options(i).Value, "COLUMNID") == piColumnID) {
                iIndex = i;
                break;
            }				
        }

        cboCombo.selectedIndex = iIndex;

        combo_disable(cboCombo, false);
    }
    else {
        combo_disable(cboCombo, true);
        button_disable(frmMainForm.cmdOK, true);
    }
	
    if (frmMainForm.optType_Field.checked == true) {
        field_refreshChildFrame();
    }
    else {
        if (frmMainForm.optType_PromptedValue.checked == true) {
            pVal_refreshValues();
        }
        else {
            lookupValue_refreshValues();
        }
    }

    util_def_exprcomponent_frmUseful.txtInitialising.value = 0;
}

function component_addValue(psValue) {
    var fOK;
    var cboColumnCombo;
    var cboValueCombo;
	
    if (frmMainForm.optType_PromptedValue.checked == true) {
        cboValueCombo = frmMainForm.cboPValDefault;
        cboColumnCombo = frmMainForm.cboPValColumn;
    }
    else {
        cboValueCombo = frmMainForm.cboLookupValueValue;
        cboColumnCombo = frmMainForm.cboLookupValueColumn;
    }

    fOK = true;
    if(columnParameter(cboColumnCombo.options[cboColumnCombo.selectedIndex].Value, "DATATYPE") == 11) {
        // Date type lookup column.
        psValue = menu_ConvertSQLDateToLocale(psValue);
    }
    else {
        if (psValue.length == 0) {
            fOK = false;
        }
    }

    if (fOK == true) {
        // Remove trailing spaces.
        while (psValue.substr(psValue.length - 1, 1) == " ") {
            psValue = psValue.substr(0, psValue.length - 1);
        }
	
        oOption = document.createElement("OPTION");
        cboValueCombo.options.add(oOption);
        oOption.innerText = psValue;
        oOption.Value = psValue;			
    }
}

function component_setValue(psValue) {
    var iIndex;
    var i;
    var fFound = false;
    var sVisibility = "hidden";
    var sDisplay = "none";
    var cboCombo;

    if (frmMainForm.optType_PromptedValue.checked == true) {
        cboCombo = frmMainForm.cboPValDefault;
    }
    else {
        cboCombo = frmMainForm.cboLookupValueValue;
    }

    for (i=0; i<cboCombo.options.length; i++)  {
        if (cboCombo.options(i).Value == psValue) {
            cboCombo.selectedIndex = i;
            fFound = true;
            break;
        }				
    }
	
    combo_disable(cboCombo, false);

    if (frmMainForm.optType_LookupTableValue.checked == true) {
        button_disable(frmMainForm.cmdOK, false);
    }
	
    if (fFound == false) {
        if ((frmMainForm.optType_LookupTableValue.checked == true) &&
            (psValue.length > 0)) {
					
            oOption = document.createElement("OPTION");
            cboCombo.options.add(oOption);
            oOption.innerText = psValue;
            oOption.Value = psValue;			

            cboCombo.selectedIndex = cboCombo.options.length - 1;
					
            frmMainForm.txtValueNotInLookup.value = psValue + 
                " does not appear in " +
                frmMainForm.cboLookupValueTable.options[frmMainForm.cboLookupValueTable.selectedIndex].text + 
                "." + 
                frmMainForm.cboLookupValueColumn.options[frmMainForm.cboLookupValueColumn.selectedIndex].text;
            var sVisibility = "visible";
            var sDisplay = "block";
        }
        else {
            if (cboCombo.options.length > 0) {
                cboCombo.selectedIndex = 0;
            }
            else {
                combo_disable(cboCombo, true);
                if (frmMainForm.optType_LookupTableValue.checked == true) {
                    button_disable(frmMainForm.cmdOK, true);
                }
            }
        }
    }
		
    frmMainForm.txtValueNotInLookup.style.visibility = sVisibility;
    frmMainForm.txtValueNotInLookup.style.display = sDisplay;

}

function columnParameter(psDefnString, psParameter)
{
    var iCharIndex;
    var sDefn;
	
    sDefn = new String(psDefnString);
	
    iCharIndex = sDefn.indexOf("	");
    if (iCharIndex >= 0) {
        if (psParameter == "COLUMNID") return sDefn.substr(0, iCharIndex);
        sDefn = sDefn.substr(iCharIndex + 1);
        iCharIndex = sDefn.indexOf("	");
        if (iCharIndex >= 0) {
            if (psParameter == "NAME") return sDefn.substr(0, iCharIndex);
            sDefn = sDefn.substr(iCharIndex + 1);
            if (psParameter == "DATATYPE") return sDefn;
        }
    }
	
    return "";
}

function tableParameter(psDefnString, psParameter)
{
    var iCharIndex;
    var sDefn;
	
    sDefn = new String(psDefnString);
	
    iCharIndex = sDefn.indexOf("	");
    if (iCharIndex >= 0) {
        if (psParameter == "TABLEID") return sDefn.substr(0, iCharIndex);
        sDefn = sDefn.substr(iCharIndex + 1);
        iCharIndex = sDefn.indexOf("	");
        if (iCharIndex >= 0) {
            if (psParameter == "NAME") return sDefn.substr(0, iCharIndex);
            sDefn = sDefn.substr(iCharIndex + 1);
            iCharIndex = sDefn.indexOf("	");
            if (iCharIndex >= 0) {
                if (psParameter == "TYPE") return sDefn.substr(0, iCharIndex);
                sDefn = sDefn.substr(iCharIndex + 1);
                iCharIndex = sDefn.indexOf("	");
                if (iCharIndex >= 0) {
                    if (psParameter == "RELATED") return sDefn.substr(0, iCharIndex);
                    sDefn = sDefn.substr(iCharIndex + 1);
                    if (psParameter == "ISCHILD") return sDefn;
                }
            }
        }
    }
	
    return "";
}

function functionAndOperatorParameter(psDefnString, psParameter)
{
    var iCharIndex;
    var sDefn;
	
    sDefn = new String(psDefnString);
	
    iCharIndex = sDefn.indexOf("	");
    if (iCharIndex >= 0) {
        if (psParameter == "ID") return sDefn.substr(0, iCharIndex);
        sDefn = sDefn.substr(iCharIndex + 1);
        iCharIndex = sDefn.indexOf("	");
        if (iCharIndex >= 0) {
            if (psParameter == "NAME") return sDefn.substr(0, iCharIndex);
            sDefn = sDefn.substr(iCharIndex + 1);
            if (psParameter == "CATEGORY") return sDefn;
        }
    }
	
    return "";
}

function calculationAndFilterParameter(psDefnString, psParameter)
{
    var iCharIndex;
    var sDefn;
	
    sDefn = new String(psDefnString);
	
    iCharIndex = sDefn.indexOf("	");
    if (iCharIndex >= 0) {
        if (psParameter == "NAME") return sDefn.substr(0, iCharIndex);
        sDefn = sDefn.substr(iCharIndex + 1);
        iCharIndex = sDefn.indexOf("	");
        if (iCharIndex >= 0) {
            if (psParameter == "EXPRID") return sDefn.substr(0, iCharIndex);
            sDefn = sDefn.substr(iCharIndex + 1);
            if (psParameter == "OWNER") return sDefn;
        }
    }
	
    return "";
}

function locateGridRecord(piID)
{  
    var fFound;
    var iIndex;
    var grdGrid;
	
    fFound = false;

    if (frmMainForm.optType_Filter.checked == true) {
        grdGrid = frmMainForm.ssOleDBGridFilters;
    }
    else {
        grdGrid = frmMainForm.ssOleDBGridCalculations;
    }
		
    grdGrid.redraw = false;
	
    grdGrid.MoveLast();
    grdGrid.MoveFirst();
			
    for (iIndex = 1; iIndex <= grdGrid.rows; iIndex++) {		
        if (grdGrid.Columns("id").value == piID) {
            grdGrid.SelBookmarks.Add(grdGrid.Bookmark);
            fFound = true;
            break;
        }

        if (iIndex < grdGrid.rows) {
            grdGrid.MoveNext();
        }
        else {
            break;
        }
    }
	
    if ((fFound == false) && (grdGrid.rows > 0)) {
        // Select the top row.
        grdGrid.MoveFirst();
        grdGrid.SelBookmarks.Add(grdGrid.Bookmark);
    }

    grdGrid.redraw = true;
}

/* Sequential search the grid for the required OLE. */
function locateGridRecordString(psString)
{  
    var fFound
    var grdGrid;
	
    if (frmMainForm.optType_Filter.checked == true) {
        grdGrid = frmMainForm.ssOleDBGridFilters;
    }
    else {
        grdGrid = frmMainForm.ssOleDBGridCalculations;
    }

    fFound = false;
	
    grdGrid.redraw = false;
    grdGrid.MoveLast();
    grdGrid.MoveFirst();

    for (iIndex = 1; iIndex <= grdGrid.rows; iIndex++) {		
        var sGridValue = new String(grdGrid.Columns(0).value);
        sGridValue = sGridValue.substr(0, psString.length).toUpperCase();
        if (sGridValue == psString.toUpperCase()) {
            grdGrid.SelBookmarks.Add(grdGrid.Bookmark);
            fFound = true;
            break;
        }
		
        if (iIndex < grdGrid.rows) {
            grdGrid.MoveNext();
        }
        else {
            break;
        }
    }

    if ((fFound == false) && (grdGrid.rows > 0)) {
        // Select the top row.
        grdGrid.MoveFirst();
        grdGrid.SelBookmarks.Add(grdGrid.Bookmark);
    }

    grdGrid.redraw = true;
}

function gridKeyPress(iKeyAscii) {
    if ((iKeyAscii >= 32) && (iKeyAscii <= 255)) {	
        var dtTicker = new Date();
        var iThisTick = new Number(dtTicker.getTime());
        if (txtLastKeyFind.value.length > 0) {
            var iLastTick = new Number(txtTicker.value);
        }
        else {
            var iLastTick = new Number("0");
        }
		
        if (iThisTick > (iLastTick + 1500)) {
            var sFind = String.fromCharCode(iKeyAscii);
        }
        else {
            var sFind = txtLastKeyFind.value + String.fromCharCode(iKeyAscii);
        }
		
        txtTicker.value = iThisTick;
        txtLastKeyFind.value = sFind;

        locateGridRecordString(sFind);
    }
}

function openDialog(pDestination, pWidth, pHeight, psResizable, psScroll)
{
    dlgwinprops = "center:yes;" +
        "dialogHeight:" + pHeight + "px;" +
        "dialogWidth:" + pWidth + "px;" +
        "help:no;" +
        "resizable:" + psResizable + ";" +
        "scroll:" + psScroll + ";" +
        "status:no;";
    window.showModalDialog(pDestination, self, dlgwinprops);
}

function component_OKClick()
{  
    var sDefn;
    var i;
    var iIndex;
    var iDataType;
    var fIsChild = false;
    var tableCollection = frmTables.elements;
    var sFunctionParameters;
    var colFunctionParameters = frmFunctionParameters.elements;

    if (frmMainForm.cmdOK.disabled == true) {
        return;
    }
	
    if (validateComponent() == true) {		
        // Component definition is valid. Pass it back to the 
        // expression page.
	
        sDefn = util_def_exprcomponent_frmOriginalDefinition.txtComponentID.value + "	0	";
	
        // Component type.
        if (frmMainForm.optType_Field.checked == true) {
            sDefn = sDefn + "1	";
        }
        if (frmMainForm.optType_Function.checked == true) {
            sDefn = sDefn + "2	";
        }
        if (frmMainForm.optType_Calculation.checked == true) {
            sDefn = sDefn + "3	";
        }
        if (frmMainForm.optType_Value.checked == true) {
            sDefn = sDefn + "4	";
        }
        if (frmMainForm.optType_Operator.checked == true) {
            sDefn = sDefn + "5	";
        }
        if (frmMainForm.optType_LookupTableValue.checked == true) {
            sDefn = sDefn + "6	";
        }
        if (frmMainForm.optType_PromptedValue.checked == true) {
            sDefn = sDefn + "7	";
        }
        if (frmMainForm.optType_Filter.checked == true) {
            sDefn = sDefn + "10	";
        }
		
        if (frmMainForm.optType_Field.checked == true) {
            // Field column ID.
            sDefn = sDefn + columnParameter(frmMainForm.cboFieldColumn.options[frmMainForm.cboFieldColumn.selectedIndex].Value, "COLUMNID") + "	";
        }
        else {
            if ((frmMainForm.optType_PromptedValue.checked == true) &&
                (frmMainForm.cboPValType.selectedIndex == 4)) {
                // Lookup prompted value.
                sDefn = sDefn + columnParameter(frmMainForm.cboPValColumn.options[frmMainForm.cboPValColumn.selectedIndex].Value, "COLUMNID") + "	";
            }
            else {
                sDefn = sDefn + "	";
            }
        }
		
        if (frmMainForm.optType_Field.checked == true) {
            // Field pass by.
            sDefn = sDefn + frmMainForm.txtPassByType.value + "	";
			
            // Field Selection Table ID (not used)
            sDefn = sDefn + "	";
				
            // Field Selection Record
            iFieldSelection = 1;
            if (frmMainForm.txtPassByType.value == 1) {
                if (frmMainForm.optField_Count.checked == true) {
                    iFieldSelection = 5;
                }
                else {
                    if (frmMainForm.optField_Total.checked == true) {
                        iFieldSelection = 4;
                    }
                    else {
                        if (frmMainForm.optFieldRecSel_Specific.checked == true) {
                            iFieldSelection = 3;
                        }
                        else {
                            if (frmMainForm.optFieldRecSel_Last.checked == true) {
                                iFieldSelection = 2;
                            }
                        }
                    }
                }
            }
            sDefn = sDefn + iFieldSelection + "	";
			
            // Field Selection Line
            if (frmMainForm.txtPassByType.value == 1) {
                if ((frmMainForm.optField_Field.checked == true) &&
                    (frmMainForm.optFieldRecSel_Specific.checked == true)) {
                    sDefn = sDefn + frmMainForm.txtFieldRecSel_Specific.value + "	";
                }
                else {
                    sDefn = sDefn + "1	";
                }
            }
            else {
                sDefn = sDefn + "1	";
            }

            // Field Selection Order ID
            sDefn = sDefn + util_def_exprcomponent_frmUseful.txtChildFieldOrderID.value + "	";

            // Field Selection Filter ID
            sDefn = sDefn + util_def_exprcomponent_frmUseful.txtChildFieldFilterID.value + "	";
        }
        else {
            sDefn = sDefn + "						";
        }		
		
        if (frmMainForm.optType_Function.checked == true) {
            // Function ID.
            sDefn = sDefn + frmMainForm.SSFunctionTree.SelectedItem.Key + "	";
        }
        else {
            sDefn = sDefn + "	";
        }
		
        if (frmMainForm.optType_Calculation.checked == true) {
            // Calculation ID.
            sDefn = sDefn + frmMainForm.ssOleDBGridCalculations.Columns("id").Value + "	";
        }
        else {
            sDefn = sDefn + "	";
        }
		
        if (frmMainForm.optType_Operator.checked == true) {
            // Operator ID.
            sDefn = sDefn + frmMainForm.SSOperatorTree.SelectedItem.Key + "	";
        }
        else {
            sDefn = sDefn + "	";
        }
		
        if (frmMainForm.optType_Value.checked == true) {
            // Value type.		
            sDefn = sDefn + frmMainForm.cboValueType.options[frmMainForm.cboValueType.selectedIndex].value + "	";

            if (frmMainForm.cboValueType.selectedIndex == 0) {
                // Character value.
                sDefn = sDefn + frmMainForm.txtValue.value + "	";
            }
            else {
                sDefn = sDefn + "	";
            }

            if (frmMainForm.cboValueType.selectedIndex == 1) {
                // Numeric value.
                sDefn = sDefn + frmMainForm.txtValue.value + "	";
            }
            else {
                sDefn = sDefn + "	";
            }

            if (frmMainForm.cboValueType.selectedIndex == 2) {
                // Logic value.
                if (frmMainForm.selectValue.selectedIndex == 0) {
                    sDefn = sDefn + "1	";
                }
                else {
                    sDefn = sDefn + "0	";
                }
            }
            else {
                sDefn = sDefn + "	";
            }
			
            if (frmMainForm.cboValueType.selectedIndex == 3) {
                // Date value.
                sDefn = sDefn + menu_convertLocaleDateToSQL(frmMainForm.txtValue.value) + "	";
            }
            else {
                sDefn = sDefn + "	";
            }
        }
        else {
            if (frmMainForm.optType_PromptedValue.checked == true) {
                // Value type.		
                sDefn = sDefn + frmMainForm.cboPValType.options[frmMainForm.cboPValType.selectedIndex].value + "	";

                if (frmMainForm.cboPValType.selectedIndex == 0) {
                    // Character value.
                    sDefn = sDefn + frmMainForm.txtPValDefault.value + "	";
                }
                else {
                    if (frmMainForm.cboPValType.selectedIndex == 4) {
                        // Lookup table value.
                        iDataType = columnParameter(frmMainForm.cboPValColumn.options[frmMainForm.cboPValColumn.selectedIndex].Value, "DATATYPE");
                        if (iDataType == 11) {
                            // Date type lookup column.
                            sDefn = sDefn + menu_convertLocaleDateToSQL(frmMainForm.cboPValDefault.options[frmMainForm.cboPValDefault.selectedIndex].Value) + "	";
                        }
                        else  {
                            // Character /Numeric/integer type lookup column.
                            if (frmMainForm.cboPValDefault.selectedIndex >= 0)
                            {
                                sDefn = sDefn + frmMainForm.cboPValDefault.options[frmMainForm.cboPValDefault.selectedIndex].Value + "	";
                            }
                            else
                            {
                                sDefn = sDefn + "" + "	";
                            }								
                        }
                    }
                    else {
                        sDefn = sDefn + "	";
                    }
                }

                if (frmMainForm.cboPValType.selectedIndex == 1) {
                    // Numeric value.
                    sDefn = sDefn + frmMainForm.txtPValDefault.value + "	";
                }
                else {
                    sDefn = sDefn + "	";
                }

                if (frmMainForm.cboPValType.selectedIndex == 2) {
                    // Logic value.
                    if (frmMainForm.cboPValDefault.selectedIndex == 0) {
                        sDefn = sDefn + "1	";
                    }
                    else {
                        sDefn = sDefn + "0	";
                    }
                }
                else {
                    sDefn = sDefn + "	";
                }
			
                if (frmMainForm.cboPValType.selectedIndex == 3) {
                    // Date value.
                    sDefn = sDefn + menu_convertLocaleDateToSQL(frmMainForm.txtPValDefault.value) + "	";
                }
                else {
                    sDefn = sDefn + "	";
                }
            }
            else {
                if (frmMainForm.optType_LookupTableValue.checked == true) {
                    iDataType = columnParameter(frmMainForm.cboLookupValueColumn.options[frmMainForm.cboLookupValueColumn.selectedIndex].Value, "DATATYPE");
					
                    // Value type
                    if ((iDataType == 2) || (iDataType == 4)) {
                        sDefn = sDefn + "2	";
                    }
                    else {
                        if (iDataType == 11) {
                            sDefn = sDefn + "4	";
                        }
                        else {
                            sDefn = sDefn + "1	";
                        }
                    }
					
                    if ((iDataType != 2) && (iDataType != 4) && (iDataType != 11)) {
                        // Character value.				
                        sDefn = sDefn + frmMainForm.cboLookupValueValue.options[frmMainForm.cboLookupValueValue.selectedIndex].Value + "	";
                    }
                    else {
                        sDefn = sDefn + "	";
                    }

                    if ((iDataType == 2) || (iDataType == 4)) {
                        // Numeric/integer value.
                        sDefn = sDefn + frmMainForm.cboLookupValueValue.options[frmMainForm.cboLookupValueValue.selectedIndex].Value + "	";
                    }
                    else {
                        sDefn = sDefn + "	";
                    }

                    // Logic value.
                    sDefn = sDefn + "	";
			
                    if (iDataType == 11) {
                        // Date value.
                        sDefn = sDefn + menu_convertLocaleDateToSQL(frmMainForm.cboLookupValueValue.options[frmMainForm.cboLookupValueValue.selectedIndex].Value) + "	";
                    }
                    else {
                        sDefn = sDefn + "	";
                    }
                }
                else {
                    sDefn = sDefn + "					";
                }
            }
        }

        if (frmMainForm.optType_PromptedValue.checked == true) {
            // Prompt.		
            sDefn = sDefn + frmMainForm.txtPrompt.value + "	";

            // Mask.		
            if (frmMainForm.cboPValType.selectedIndex == 0) {
                // Character
                sDefn = sDefn + frmMainForm.txtPValFormat.value + "	";
            }
            else {
                sDefn = sDefn + "	";
            }

            // Size.		
            if ((frmMainForm.cboPValType.selectedIndex == 0) ||
                (frmMainForm.cboPValType.selectedIndex == 1)) {
                // Character or numeric
                sDefn = sDefn + frmMainForm.txtPValSize.value + "	";
            }
            else {
                sDefn = sDefn + "	";
            }
			
            // Decimals.		
            if (frmMainForm.cboPValType.selectedIndex == 1) {
                // Numeric
                sDefn = sDefn + frmMainForm.txtPValDecimals.value + "	";
            }
            else {
                sDefn = sDefn + "	";
            }
        }
        else {
            sDefn = sDefn + "				";
        }

        // Function Return Type (not used)
        sDefn = sDefn + "	";

        if (frmMainForm.optType_LookupTableValue.checked == true) {
            // Lookup Table ID
            sDefn = sDefn + frmMainForm.cboLookupValueTable.options[frmMainForm.cboLookupValueTable.selectedIndex].Value + "	";

            // Lookup Column ID
            sDefn = sDefn + columnParameter(frmMainForm.cboLookupValueColumn.options[frmMainForm.cboLookupValueColumn.selectedIndex].Value, "COLUMNID") + "	";
        }
        else {
            sDefn = sDefn + "		";
        }
		
        if (frmMainForm.optType_Filter.checked == true) {
            // Filter ID.
            sDefn = sDefn + frmMainForm.ssOleDBGridFilters.Columns("id").Value + "	";
        }
        else {
            sDefn = sDefn + "	";
        }
		
        // Expanded Node (not used)
        sDefn = sDefn + "	";

        if (frmMainForm.optType_PromptedValue.checked == true) {
            // Prompted Value Date type.
            if (frmMainForm.cboPValType.selectedIndex == 3) {
                // Date type
                if (frmMainForm.optPValDate_Explicit.checked == true) {
                    sDefn = sDefn + "0	";
                }
                if (frmMainForm.optPValDate_MonthStart.checked == true) {
                    sDefn = sDefn + "2	";
                }
                if (frmMainForm.optPValDate_Current.checked == true) {
                    sDefn = sDefn + "1	";
                }
                if (frmMainForm.optPValDate_YearStart.checked == true) {
                    sDefn = sDefn + "4	";
                }
                if (frmMainForm.optPValDate_MonthEnd.checked == true) {
                    sDefn = sDefn + "3	";
                }
                if (frmMainForm.optPValDate_YearEnd.checked == true) {
                    sDefn = sDefn + "5	";
                }
            }
            else {
                sDefn = sDefn + "	";
            }
        }
        else {
            sDefn = sDefn + "	";
        }		

        // Description		
        if (frmMainForm.optType_Field.checked == true) {
            sDefn = sDefn + frmMainForm.cboFieldTable.options[frmMainForm.cboFieldTable.selectedIndex].text;
			
            if (frmMainForm.txtPassByType.value == 2) {
                sDefn = sDefn + 
                    " : " + frmMainForm.cboFieldColumn.options[frmMainForm.cboFieldColumn.selectedIndex].text;
            }
            else {
                if (frmMainForm.optField_Count.checked==false) {
                    sDefn = sDefn + 
                        " : " + frmMainForm.cboFieldColumn.options[frmMainForm.cboFieldColumn.selectedIndex].text;
                }
            }
				
            if (frmMainForm.txtPassByType.value == 1) {
                if (tableCollection != null) {
                    for (i=0; i<tableCollection.length; i++)  {
                        if (tableParameter(tableCollection.item(i).value, "TABLEID") == frmMainForm.cboFieldTable.options[frmMainForm.cboFieldTable.selectedIndex].Value) {
                            if (tableParameter(tableCollection.item(i).value, "ISCHILD") == 1) {
                                fIsChild = true;
                            }
                            break;
                        }
                    }
                }	

                if (fIsChild == true) {
                    if (iFieldSelection == 1) {
                        sDefn = sDefn + " (first record";
                    }
                    if (iFieldSelection == 2) {
                        sDefn = sDefn + " (last record";
                    }
                    if (iFieldSelection == 3) {
                        sDefn = sDefn + " (line " + frmMainForm.txtFieldRecSel_Specific.value;
                    }
                    if (iFieldSelection == 4) {
                        sDefn = sDefn + " (total";
                    }
                    if (iFieldSelection == 5) {
                        sDefn = sDefn + " (record count";
                    }

                    if (util_def_exprcomponent_frmUseful.txtChildFieldOrderID.value > 0) {
                        sDefn = sDefn + ", order by '" + frmMainForm.txtFieldRecOrder.value + "'";
                    }
                    if (util_def_exprcomponent_frmUseful.txtChildFieldFilterID.value > 0) {
                        sDefn = sDefn + ", filter by '" + frmMainForm.txtFieldRecFilter.value + "'";
                    }

                    sDefn = sDefn + ")";
                }
            }				
			
            sDefn = sDefn + "	";
        }
        else {
            if (frmMainForm.optType_Function.checked == true) {
                sTemp = frmMainForm.SSFunctionTree.SelectedItem.Text;
                iIndex = sTemp.indexOf("(");
                if (iIndex >= 0) {
                    sTemp = sTemp.substr(0, iIndex-1);
                }

                sDefn = sDefn + sTemp + "	";
            }
            else {
                if (frmMainForm.optType_Calculation.checked == true) {
                    sDefn = sDefn + frmMainForm.ssOleDBGridCalculations.Columns("name").Value + "	";
                }
                else {
                    if (frmMainForm.optType_Operator.checked == true) {
                        sTemp = frmMainForm.SSOperatorTree.SelectedItem.Text;
                        iIndex = sTemp.indexOf("(");
                        if (iIndex >= 0) {
                            sTemp = sTemp.substr(0, iIndex-1);
                        }
                        sDefn = sDefn + sTemp + "	";
                    }
                    else {
                        if (frmMainForm.optType_Filter.checked == true) {
                            sDefn = sDefn + frmMainForm.ssOleDBGridFilters.Columns("name").Value + "	";
                        }
                        else {
                            sDefn = sDefn + "	";
                        }
                    }
                }
            }
        }
		
        if (frmMainForm.optType_Field.checked == true) {
            // Field table ID.
            sDefn = sDefn + frmMainForm.cboFieldTable.options[frmMainForm.cboFieldTable.selectedIndex].Value + "	";

            // Field Selection Order Name
            if (frmMainForm.txtPassByType.value == 1) {
                sDefn = sDefn + frmMainForm.txtFieldRecOrder.value + "	";

                // Field Selection Filter Name
                sDefn = sDefn + frmMainForm.txtFieldRecFilter.value;
            }
            else {
                sDefn = sDefn + "	";
            }
        }
        else {
            if ((frmMainForm.optType_PromptedValue.checked == true) &&
                (frmMainForm.cboPValType.selectedIndex == 4)) {
                // Lookup prompted value.
                sDefn = sDefn + frmMainForm.cboPValTable.options[frmMainForm.cboPValTable.selectedIndex].Value + "		";
            }
            else {
                sDefn = sDefn + "		";
            }
        }

        // Determine the function parameters (if required)
        sFunctionParameters = "";
        if (frmMainForm.optType_Function.checked == true) {
            if (colFunctionParameters != null) {
                for (i=0; i<colFunctionParameters.length; i++)  {
                    sKey = colFunctionParameters.item(i).name;
                    sKey = sKey.substr(22);
                    iIndex = sKey.indexOf("_");
                    if (iIndex >= 0) {
                        sKey = sKey.substr(0, iIndex);
                    }
					
                    if (sKey == frmMainForm.SSFunctionTree.SelectedItem.Key) {
                        if (colFunctionParameters.item(i).value.length > 0) {
                            if (sFunctionParameters.length > 0) {
                                sFunctionParameters = sFunctionParameters + "	";
                            }
							
                            sFunctionParameters = sFunctionParameters + colFunctionParameters.item(i).value;
                        }
                    }
                }
            }	
        }		
		
        // Pass the component definition back to the expression page.
        //OpenHR.getFrame("workframe").setComponent(sDefn, util_def_exprcomponent_frmUseful.txtAction.value, util_def_exprcomponent_frmUseful.txtLinkRecordID.value, sFunctionParameters);
        setComponent(sDefn, util_def_exprcomponent_frmUseful.txtAction.value, util_def_exprcomponent_frmUseful.txtLinkRecordID.value, sFunctionParameters);
            

    }
}

function component_CancelClick()
{  
    //OpenHR.getFrame("workframe").cancelComponent();
    cancelComponent();
}

function validateComponent() {
    var sErrorMsg = "";
    var sValue;
    var sTemp;
    var iValue;
    var i;
    var iRealFormatLength;
    var fLastLiteralChar;
    var fUserEnterableChar;
    var cChar;
    var sDecimalSeparator;
    var sThousandSeparator;
    var sPoint;
    var sConvertedValue;

    sDecimalSeparator = "\\";
    sDecimalSeparator = sDecimalSeparator.concat(OpenHR.LocaleDecimalSeparator);
    var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");
  
    sThousandSeparator = "\\";
    sThousandSeparator = sThousandSeparator.concat(OpenHR.LocaleThousandSeparator);
    var reThousandSeparator = new RegExp(sThousandSeparator, "gi");

    sPoint = "\\.";
    var rePoint = new RegExp(sPoint, "gi");
	
    if (frmMainForm.optType_Field.checked == true) {
        // If the selected table is a child of the expression base table,
        // then an order must be selected.
        if ((frmMainForm.txtPassByType.value == 1) &&
            (frmMainForm.btnFieldRecOrder.disabled == false) &&
            (util_def_exprcomponent_frmUseful.txtChildFieldOrderID.value <= 0)) {
            sErrorMsg = "An order must be specified when referring to child fields.";		
        }
        else {
            if ((frmMainForm.txtPassByType.value == 1) &&
                (frmMainForm.optFieldRecSel_Specific.checked == true)) {

                sValue = frmMainForm.txtFieldRecSel_Specific.value;

                // Numeric
                if (sValue.length == 0) {
                    frmMainForm.txtFieldRecSel_Specific.value = 1;
                    sValue = "1";
                }

                // Convert the value from locale to UK settings for use with the isNaN funtion.
                sConvertedValue = new String(sValue);
                // Remove any thousand separators.
                sConvertedValue = sConvertedValue.replace(reThousandSeparator, "");
                frmMainForm.txtFieldRecSel_Specific.value = sConvertedValue;
                   
                // Convert any decimal separators to '.'.
                if (OpenHR.LocaleDecimalSeparator != ".") {
                    // Remove decimal points.
                    sConvertedValue = sConvertedValue.replace(rePoint, "A");
                    // replace the locale decimal marker with the decimal point.
                    sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
                }

                if (isNaN(sConvertedValue) == true) {
                    sErrorMsg = "Invalid specific line value entered.";
                }
                else {
                    if (frmMainForm.txtFieldRecSel_Specific.value < 1) {
                        sErrorMsg = "Specific line value must be greater than 0.";
                    }
                    else {
                        if (sConvertedValue.indexOf(".") >= 0 ) {
                            sErrorMsg = "Specific line value must be an integer.";
                        }
                    }
                }
            }
        }
    }

    if (frmMainForm.optType_Value.checked == true) {
        // Check that numeric and date values are valid.
        sValue = frmMainForm.txtValue.value;

        if (frmMainForm.cboValueType.value == 2) {
            // Numeric
            if (sValue.length == 0) {
                frmMainForm.txtValue.value = 0;
                sValue = "0";
            }

            // Convert the value from locale to UK settings for use with the isNaN funtion.
            sConvertedValue = new String(sValue);
            // Remove any thousand separators.
            sConvertedValue = sConvertedValue.replace(reThousandSeparator, "");
            frmMainForm.txtValue.value = sConvertedValue;

            // Convert any decimal separators to '.'.
            if (OpenHR.LocaleDecimalSeparator != ".") {
                // Remove decimal points.
                sConvertedValue = sConvertedValue.replace(rePoint, "A");
                // replace the locale decimal marker with the decimal point.
                sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
            }

            if (isNaN(sConvertedValue) == true) {
                sErrorMsg = "Invalid numeric value entered.";
            }
        }

        if (frmMainForm.cboValueType.value == 4) {
            // Date
            // Convert the date to SQL format (use this as a validation check).
            // An empty string is returned if the date is invalid.
            sValue = menu_convertLocaleDateToSQL(sValue);
            if (sValue.length == 0) {
                sErrorMsg = "Invalid date value entered.";
            }
        }
    }

    if (frmMainForm.optType_PromptedValue.checked == true) {
        // Check that size value is valid.
        if ((frmMainForm.cboPValType.value == 1) ||
            (frmMainForm.cboPValType.value == 2)) {
            sValue = frmMainForm.txtPValSize.value;
			
            if (sValue.length == 0) {
                frmMainForm.txtPValSize.value = 0;
                sValue = "0";
            }

            // Convert the value from locale to UK settings for use with the isNaN funtion.
            sConvertedValue = new String(sValue);
            // Remove any thousand separators.
            sConvertedValue = sConvertedValue.replace(reThousandSeparator, "");
            frmMainForm.txtPValSize.value = sConvertedValue;

            // Convert any decimal separators to '.'.
            if (OpenHR.LocaleDecimalSeparator != ".") {
                // Remove decimal points.
                sConvertedValue = sConvertedValue.replace(rePoint, "A");
                // replace the locale decimal marker with the decimal point.
                sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
            }

            if (isNaN(sConvertedValue) == true) {
                sErrorMsg = "Invalid size value entered.";
            }
            else {
                // Size must be integer.		
                if (sConvertedValue.indexOf(".") >= 0 ) {
                    sErrorMsg = "Size value must be an integer value.";
                }
                else {
                    iValue = new Number(sValue);
                    if ((iValue < 1) 	
                        || (iValue > 250)) {
                        sErrorMsg = "The size value must be between 1 and 250.";
                    }
                }
            }
        }
		
        // Check that decimals value is valid.
        if (sErrorMsg.length == 0) {
            if (frmMainForm.cboPValType.value == 2) {
                sValue = frmMainForm.txtPValDecimals.value;
				
                if (sValue.length == 0) {
                    frmMainForm.txtPValDecimals.value = 0;
                    sValue = "0";
                }

                // Convert the value from locale to UK settings for use with the isNaN funtion.
                sConvertedValue = new String(sValue);
                // Remove any thousand separators.
                sConvertedValue = sConvertedValue.replace(reThousandSeparator, "");
                frmMainForm.txtPValDecimals.value = sConvertedValue;

                // Convert any decimal separators to '.'.
                if (OpenHR.LocaleDecimalSeparator != ".") {
                    // Remove decimal points.
                    sConvertedValue = sConvertedValue.replace(rePoint, "A");
                    // replace the locale decimal marker with the decimal point.
                    sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
                }

                if (isNaN(sConvertedValue) == true) {
                    sErrorMsg = "Invalid decimals value entered.";
                }
                else {
                    // Size must be integer.		
                    if (sConvertedValue.indexOf(".") >= 0 ) {
                        sErrorMsg = "Decimals value must be an integer value.";
                    }
                    else {
                        iValue = new Number(sValue);
                        if ((iValue < 0) 	
                            || (iValue > 4)) {
                            sErrorMsg = "The decimals value must be between 0 and 4.";
                        }
                    }
                }
            }
        }
		
        // Check that the format is valid.		
        if (sErrorMsg.length == 0) {
            iRealFormatLength = 0;
            fLastLiteralChar = false;
            fUserEnterableChar = false;

            if (frmMainForm.cboPValType.value == 1) {
                // Format must match the defined size, and
                // include at least on user enterable character.
                sValue = frmMainForm.txtPValFormat.value;
                sTemp = "";
				
                for(i=0;i<sValue.length;i++) {
                    cChar = sValue.substr(i, 1);			

                    if (fLastLiteralChar == false) {
                        if (cChar == "\\") {
                            // Literal marker.
                            fLastLiteralChar = true;
                        }
                        else {
                            fLastLiteralChar = false;
                            iRealFormatLength = iRealFormatLength + 1;
							
                            if ((cChar == "A") ||
                                (cChar == "a") ||
                                (cChar == "9") ||
                                (cChar == "#") ||
                                (cChar == "B")) {
                                fUserEnterableChar = true;
                            }
                        }
                    }
                    else {
                        fLastLiteralChar = false;
                        iRealFormatLength = iRealFormatLength + 1;
                    }
                }	
            }

            if (iRealFormatLength > 0) {
                if (iRealFormatLength != frmMainForm.txtPValSize.value) {
                    sErrorMsg = "The mask must correspond with the 'size' setting.";
                }
                else {
                    if (fUserEnterableChar == false) {
                        sErrorMsg = "You must have at least one user enterable character in the mask field.";
                    }
                }
            }
        }

        // Check that the default numeric value is valid.
        if (sErrorMsg.length == 0) {
            if (frmMainForm.cboPValType.value == 2) {
                sValue = frmMainForm.txtPValDefault.value;
				
                if (sValue.length > 0) {
                    // Convert the value from locale to UK settings for use with the isNaN funtion.
                    sConvertedValue = new String(sValue);
                    // Remove any thousand separators.
                    sConvertedValue = sConvertedValue.replace(reThousandSeparator, "");
                    frmMainForm.txtPValDefault.value = sConvertedValue;

                    // Convert any decimal separators to '.'.
                    if (OpenHR.LocaleDecimalSeparator != ".") {
                        // Remove decimal points.
                        sConvertedValue = sConvertedValue.replace(rePoint, "A");
                        // replace the locale decimal marker with the decimal point.
                        sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
                    }

                    if (isNaN(sConvertedValue) == true) {
                        sErrorMsg = "Invalid default numeric value entered.";
                    }
                }
            }
        }

        // Check that the default date value is valid.
        if (sErrorMsg.length == 0) {
            if (frmMainForm.cboPValType.value == 4) {
                sValue = frmMainForm.txtPValDefault.value;
				
                if (sValue.length > 0) {
                    // Convert the date to SQL format (use this as a validation check).
                    // An empty string is returned if the date is invalid.
                    sValue = menu_convertLocaleDateToSQL(sValue);
                    if (sValue.length == 0) {
                        sErrorMsg = "Invalid default date value entered.";
                    }
                }
            }
        }
    }
	
    if (sErrorMsg.length > 0) {
        OpenHR.messageBox(sErrorMsg);
        return false;
    }
	
    return true;
}

function component_saveChanges(psAction, pfPrompt, pfTBOverride)
{
    // Expand the work frame and hide the option frame.
    $("#optionframe").attr("data-framesource", "UTIL_DEF_EXPRCOMPONENT");
    $("#optionframe").hide();
    $("#workframe").show();
        
    // Pass the component definition back to the expression page.        
    //OpenHR.getFrame("workframe").saveChanges(psAction, pfPrompt, pfTBOverride);
    saveChanges(psAction, pfPrompt, pfTBOverride);
}

</script>

<script type="text/javascript">    
    function util_def_exprcomponent_addhandlers() {
        
        OpenHR.addActiveXHandler("ssOleDBGridFilters", "KeyPress", ssOleDBGridFilters_KeyPress);
        OpenHR.addActiveXHandler("ssOleDBGridFilters", "rowcolchange", ssOleDBGridFilters_rowcolchange);
        OpenHR.addActiveXHandler("ssOleDBGridFilters", "dblClick", ssOleDBGridFilters_dblClick);
        OpenHR.addActiveXHandler("ssOleDBGridCalculations", "KeyPress", ssOleDBGridCalculations_KeyPress);
        OpenHR.addActiveXHandler("ssOleDBGridCalculations", "rowcolchange", ssOleDBGridCalculations_rowcolchange);
        OpenHR.addActiveXHandler("ssOleDBGridCalculations", "dblClick", ssOleDBGridCalculations_dblClick);
        OpenHR.addActiveXHandler("SSOperatorTree", "nodeClick", SSOperatorTree_nodeClick);
        OpenHR.addActiveXHandler("SSOperatorTree", "dblClick", SSOperatorTree_dblClick);
        OpenHR.addActiveXHandler("SSFunctionTree", "nodeClick", SSFunctionTree_nodeClick);
        OpenHR.addActiveXHandler("SSFunctionTree", "dblClick", SSFunctionTree_dblClick);

    }
    
function ssOleDBGridFilters_KeyPress(iKeyAscii) {
    gridKeyPress(iKeyAscii);        
}
    
function ssOleDBGridFilters_rowcolchange() {
    var sDesc;

    var filterCollection = frmFilters.elements;
    sDesc = "txtFilterDesc_" + (frmMainForm.ssOleDBGridFilters.AddItemRowIndex(frmMainForm.ssOleDBGridFilters.Bookmark) + 1);
    frmMainForm.txtFilterDescription.value = frmFilters.all.item(sDesc).value;

    button_disable(frmMainForm.cmdOK, (frmMainForm.ssOleDBGridFilters.SelBookmarks.Count == 0));
}

function ssOleDBGridFilters_dblClick() {
    component_OKClick();
}

function ssOleDBGridCalculations_KeyPress(iKeyAscii) {
    gridKeyPress(iKeyAscii);        
}

function ssOleDBGridCalculations_rowcolchange() {

    var sDesc;

    var calcCollection = frmCalcs.elements;
    sDesc = "txtCalcDesc_" + (frmMainForm.ssOleDBGridCalculations.AddItemRowIndex(frmMainForm.ssOleDBGridCalculations.Bookmark) + 1);
    frmMainForm.txtCalcDescription.value = frmCalcs.all.item(sDesc).value;

    button_disable(frmMainForm.cmdOK, (frmMainForm.ssOleDBGridCalculations.SelBookmarks.Count == 0));
}

function ssOleDBGridCalculations_dblClick() {
    component_OKClick();        
}

function SSOperatorTree_nodeClick(node) {

    var fNoNodeSelected;

    fNoNodeSelected = true;

    if (node.Key != "OPERATOR_ROOT") {
        if (node.Parent.Key != "OPERATOR_ROOT") {
            fNoNodeSelected = false;
        }
    }
  
    button_disable(frmMainForm.cmdOK, fNoNodeSelected);  
        
}

function SSOperatorTree_dblClick() {
    component_OKClick();        
}

function SSFunctionTree_nodeClick(node) {
        
    var fNoNodeSelected;

    fNoNodeSelected = true;

    if (node.Key != "FUNCTION_ROOT") {
        if (node.Parent.Key != "FUNCTION_ROOT") {
            fNoNodeSelected = false;
        }
    }
  
    button_disable(frmMainForm.cmdOK, fNoNodeSelected);

}
    
function SSFunctionTree_dblClick() {
    component_OKClick();        
}

