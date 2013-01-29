<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
   
    <OBJECT 
        classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" 
        id="Microsoft_Licensed_Class_Manager_1_0" 
        VIEWASTEXT>
        <PARAM NAME="LPKPath" VALUE="lpks/main.lpk">
    </OBJECT>

        
<script type="text/javascript">
<!--
    
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

        // Hide SQL2000 specifics if required
        if (util_def_exprcomponent_frmUseful.txtEnableSQL2000Functions.value == "False") {
            sUDFFunction_Visibility = "hidden";
            sUDFFunction_Display = "none";
        }
        else {
            sUDFFunction_Visibility = "visible";
            sUDFFunction_Display = "block";
        }
		
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

        // Little dodge to get around a browser bug that
        // does not refresh the display on all controls.
        try
        {
            window.resizeBy(0,-1);
            window.resizeBy(0,1);	
        }
        catch(e){}
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

        iType = new Number(frmOriginalDefinition.txtType.value);
	
        if (iType == 1) {
            // Field
            frmMainForm.optType_Field.checked = true;
            changeType(iType);

            util_def_exprcomponent_frmUseful.txtInitialising.value = 0;
		
            if (frmMainForm.txtPassByType.value == 1) {
                if (frmOriginalDefinition.txtFieldSelectionRecord.value == 5) {
                    frmMainForm.optField_Count.checked = true;	
                }
                else {
                    if (frmOriginalDefinition.txtFieldSelectionRecord.value == 4) {
                        frmMainForm.optField_Total.checked = true;	
                    }
                    else {
                        frmMainForm.optField_Field.checked = true;	
					
                        if (frmOriginalDefinition.txtFieldSelectionRecord.value == 3) {
                            frmMainForm.optFieldRecSel_Specific.checked = true;	
                            frmMainForm.txtFieldRecSel_Specific.value = frmOriginalDefinition.txtFieldSelectionLine.value;
                        }
                        else {
                            if (frmOriginalDefinition.txtFieldSelectionRecord.value == 2) {
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
                util_def_exprcomponent_frmUseful.txtChildFieldOrderID.value = frmOriginalDefinition.txtFieldSelectionOrderID.value;
                frmMainForm.txtFieldRecOrder.value = frmOriginalDefinition.txtFieldOrderName.value;
                util_def_exprcomponent_frmUseful.txtChildFieldFilterID.value = frmOriginalDefinition.txtFieldSelectionFilter.value;
                frmMainForm.txtFieldRecFilter.value = frmOriginalDefinition.txtFieldFilterName.value;
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
            frmMainForm.SSFunctionTree.SelectedItem = frmMainForm.SSFunctionTree.Nodes(frmOriginalDefinition.txtFunctionID.value);
            button_disable(frmMainForm.cmdOK, false);
        }
	
        if (iType == 3) {
            // Calculation.
            frmMainForm.optType_Calculation.checked = true;
            changeType(iType);
		
            util_def_exprcomponent_frmUseful.txtInitialising.value = 0;
            calculationsAndFilters_load();
		
            // Locate the current calc.
            locateGridRecord(frmOriginalDefinition.txtCalculationID.value);
        }
	
        if (iType == 4) {
            // Value
            frmMainForm.optType_Value.checked = true;
            changeType(iType);

            if (frmOriginalDefinition.txtValueType.value == 1) {
                // Character value
                frmMainForm.cboValueType.selectedIndex = 0;
                value_changeType();
                frmMainForm.txtValue.value = frmOriginalDefinition.txtValueCharacter.value;
            }
            if (frmOriginalDefinition.txtValueType.value == 2) {
                // Numeric value
                frmMainForm.cboValueType.selectedIndex = 1;
                value_changeType();
                frmMainForm.txtValue.value = frmOriginalDefinition.txtValueNumeric.value;
            }
            if (frmOriginalDefinition.txtValueType.value == 3) {
                // Logic value
                frmMainForm.cboValueType.selectedIndex = 2;
                value_changeType();
                if (frmOriginalDefinition.txtValueLogic.value == 1) {
                    frmMainForm.selectValue.selectedIndex = 0;
                }
                else {
                    frmMainForm.selectValue.selectedIndex = 1;
                }
            }
            if (frmOriginalDefinition.txtValueType.value == 4) {
                // Date value
                frmMainForm.cboValueType.selectedIndex = 3;
                value_changeType();
                frmMainForm.txtValue.value = menu_ConvertSQLDateToLocale(frmOriginalDefinition.txtValueDate.value);
            }
        }

        if (iType == 5) {
            // Operator
            frmMainForm.optType_Operator.checked = true;
            changeType(iType);

            util_def_exprcomponent_frmUseful.txtInitialising.value = 0;
            functionAndOperator_refresh();
            frmMainForm.SSOperatorTree.SelectedItem = frmMainForm.SSOperatorTree.Nodes(frmOriginalDefinition.txtOperatorID.value);
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
		
            frmMainForm.txtPrompt.value = frmOriginalDefinition.txtPromptDescription.value;
            frmMainForm.cboPValType.selectedIndex = frmOriginalDefinition.txtValueType.value - 1;
            frmMainForm.txtPValSize.value = frmOriginalDefinition.txtPromptSize.value;
            frmMainForm.txtPValDecimals.value = frmOriginalDefinition.txtPromptDecimals.value;
            frmMainForm.txtPValFormat.value = frmOriginalDefinition.txtPromptMask.value;

            pVal_changeType();
		
            if (frmOriginalDefinition.txtValueType.value == 1) {
                // Character
                frmMainForm.txtPValDefault.value = frmOriginalDefinition.txtValueCharacter.value;
            }		
            if (frmOriginalDefinition.txtValueType.value == 2) {
                // Numeric
                frmMainForm.txtPValDefault.value = frmOriginalDefinition.txtValueNumeric.value;
            }		
            if (frmOriginalDefinition.txtValueType.value == 3) {
                // Logic
                iIndex = 0;
                for (i=0; i<frmMainForm.cboPValDefault.options.length; i++)  {
                    if (frmMainForm.cboPValDefault.options(i).Value == frmOriginalDefinition.txtValueLogic.value) {
                        iIndex = i;
                        break;
                    }		
                }
	
                frmMainForm.cboPValDefault.selectedIndex = iIndex;
            }		
		
            if (frmOriginalDefinition.txtValueType.value == 4) {
                // Date
                if (frmOriginalDefinition.txtPromptDateType.value == 0) {
                    if ((frmOriginalDefinition.txtValueDate.value != "12/30/1899") &&
                        (frmOriginalDefinition.txtValueDate.value != "")) {
                        frmMainForm.txtPValDefault.value = menu_ConvertSQLDateToLocale(frmOriginalDefinition.txtValueDate.value);
                    }
                    frmMainForm.optPValDate_Explicit.checked = true;
                }
                if (frmOriginalDefinition.txtPromptDateType.value == 2) {
                    frmMainForm.optPValDate_MonthStart.checked = true;
                }
                if (frmOriginalDefinition.txtPromptDateType.value == 1) {
                    frmMainForm.optPValDate_Current.checked = true;
                }
                if (frmOriginalDefinition.txtPromptDateType.value == 4) {
                    frmMainForm.optPValDate_YearStart.checked = true;
                }
                if (frmOriginalDefinition.txtPromptDateType.value == 3) {
                    frmMainForm.optPValDate_MonthEnd.checked = true;
                }
                if (frmOriginalDefinition.txtPromptDateType.value == 5) {
                    frmMainForm.optPValDate_YearEnd.checked = true;
                }

                pVal_changeDateOption(frmOriginalDefinition.txtPromptDateType.value)
            }		
        }

        if (iType == 10) {
            // Filter.
            frmMainForm.optType_Filter.checked = true;
            changeType(iType);
		
            util_def_exprcomponent_frmUseful.txtInitialising.value = 0;
            calculationsAndFilters_load();
		
            // Locate the current filter.
            locateGridRecord(frmOriginalDefinition.txtFilterID.value);
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

        // Little dodge to get around a browser bug that
        // does not refresh the display on all controls.
        try
        {
            window.resizeBy(0,-1);
            window.resizeBy(0,1);
        }
        catch(e) {}
	
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
	
        sDefaultTableID = frmOriginalDefinition.txtFieldTableID.value;

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
	
        field_refreshColumn()
    }

    function field_refreshColumn() {
        var sDefaultColumnID;
        var fInitialise = util_def_exprcomponent_frmUseful.txtInitialising.value;
	
        sDefaultColumnID = frmOriginalDefinition.txtFieldColumnID.value;

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
            var optionDataForm = OpenHR.getFrame("optiondataframe","frmGetOptionData");
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

            OpenHR.getFrame("optiondataframe").refreshOptionData();

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
                    radio_disable(frmMainForm.optFieldRecSel_Specific, (util_def_exprcomponent_frmUseful.txtEnableSQL2000Functions.value == "False"));
				
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
	
        // Little dodge to get around a browser bug that
        // does not refresh the display on all controls.
        try
        {
            window.resizeBy(0,-1);
            window.resizeBy(0,1);	
        }
        catch(e) {}
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

        // Little dodge to get around a browser bug that
        // does not refresh the display on all controls.
        try
        {
            window.resizeBy(0,-1);
            window.resizeBy(0,1);	
        }
        catch(e) {}
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

        // Little dodge to get around a browser bug that
        // does not refresh the display on all controls.
        try
        {
            window.resizeBy(0,-1);
            window.resizeBy(0,1);	
        }
        catch(e) {}
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
	
        sDefaultTableID = frmOriginalDefinition.txtLookupTableID.value;

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
	
        sDefaultColumnID = frmOriginalDefinition.txtLookupColumnID.value;

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
            var optionDataForm = window.parent.frames("optiondataframe").document.forms("frmGetOptionData");
            optionDataForm.txtOptionAction.value = "LOADEXPRLOOKUPCOLUMNS";
            optionDataForm.txtOptionTableID.value = frmMainForm.cboLookupValueTable.options[frmMainForm.cboLookupValueTable.selectedIndex].Value;
            optionDataForm.txtOptionColumnID.value = sDefaultColumnID;

            OpenHR.getFrame("optiondataframe").refreshOptionData();
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

            if ((frmMainForm.cboLookupValueTable.options[frmMainForm.cboLookupValueTable.selectedIndex].Value == frmOriginalDefinition.txtLookupTableID.value)  &&
                (columnParameter(frmMainForm.cboLookupValueColumn.options[frmMainForm.cboLookupValueColumn.selectedIndex].Value, "COLUMNID") == frmOriginalDefinition.txtLookupColumnID.value)) {

                if (iDataType == 11) {
                    // Date type lookup column.
                    sDefaultValue = frmOriginalDefinition.txtValueDate.value;
                    sDefaultValue = menu_ConvertSQLDateToLocale(sDefaultValue);
                }
                if (iDataType == 12) {
                    // Character type lookup column.
                    sDefaultValue = frmOriginalDefinition.txtValueCharacter.value;
                }
                if ((iDataType == 2) || (iDataType == 4)) {
                    // Numeric/integer type lookup column.
                    sDefaultValue = frmOriginalDefinition.txtValueNumeric.value;
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
            var optionDataForm = OpenHR.getFrame("optiondataframe","frmGetOptionData");
            optionDataForm.txtOptionAction.value = "LOADEXPRLOOKUPVALUES";
            optionDataForm.txtOptionColumnID.value = columnParameter(frmMainForm.cboLookupValueColumn.options[frmMainForm.cboLookupValueColumn.selectedIndex].Value, "COLUMNID");
            optionDataForm.txtGotoLocateValue.value = sDefaultValue;
            
            OpenHR.getFrame("optiondataframe").refreshOptionData();
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
	
        sDefaultTableID = frmOriginalDefinition.txtFieldTableID.value;

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
	
        sDefaultColumnID = frmOriginalDefinition.txtFieldColumnID.value;

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
            var optionDataForm = OpenHR.getFrame("optiondataframe","frmGetOptionData");
            optionDataForm.txtOptionAction.value = "LOADEXPRLOOKUPCOLUMNS";
            optionDataForm.txtOptionTableID.value = frmMainForm.cboPValTable.options[frmMainForm.cboPValTable.selectedIndex].Value;
            optionDataForm.txtOptionColumnID.value = sDefaultColumnID;

            OpenHR.getFrame("optiondataframe").refreshOptionData();
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

            if ((frmMainForm.cboPValTable.options[frmMainForm.cboPValTable.selectedIndex].Value == frmOriginalDefinition.txtFieldTableID.value)  &&
                (columnParameter(frmMainForm.cboPValColumn.options[frmMainForm.cboPValColumn.selectedIndex].Value, "COLUMNID") == frmOriginalDefinition.txtFieldColumnID.value)) {

                if (iDataType == 11) {
                    // Date type lookup column.
                    sDefaultValue = frmOriginalDefinition.txtValueCharacter.value;
                    sDefaultValue = menu_ConvertSQLDateToLocale(sDefaultValue);
                }
                if (iDataType == 12) {
                    // Character type lookup column.
                    sDefaultValue = frmOriginalDefinition.txtValueCharacter.value;
                }
                if ((iDataType == 2) || (iDataType == 4)) {
                    // Numeric/integer type lookup column.
                    sDefaultValue = frmOriginalDefinition.txtValueCharacter.value;
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
            var optionDataForm = OpenHR.getFrame("optiondataframe","frmGetOptionData");
            optionDataForm.txtOptionAction.value = "LOADEXPRLOOKUPVALUES";
            optionDataForm.txtOptionColumnID.value = columnParameter(frmMainForm.cboPValColumn.options[frmMainForm.cboPValColumn.selectedIndex].Value, "COLUMNID");
            optionDataForm.txtGotoLocateValue.value = sDefaultValue;

            OpenHR.getFrame("optiondataframe").refreshOptionData();
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

    function addColumn(psDefn) {
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

    function setColumn(piColumnID) {
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

    function addValue(psValue) {
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

    function setValue(psValue) {
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

        // Little dodge to get around a browser bug that
        // does not refresh the display on all controls.
        try
        {
            window.resizeBy(0,-1);
            window.resizeBy(0,1);	
        }
        catch(e) {}
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

    function OKClick()
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
	
            sDefn = frmOriginalDefinition.txtComponentID.value + "	0	";
	
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
            OpenHR.getFrame("workframe").setComponent(sDefn, util_def_exprcomponent_frmUseful.txtAction.value, util_def_exprcomponent_frmUseful.txtLinkRecordID.value, sFunctionParameters);
            

        }
    }

    function CancelClick()
    {  
        OpenHR.getFrame("workframe").cancelComponent();
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
        sDecimalSeparator = sDecimalSeparator.concat(window.parent.frames("menuframe").ASRIntranetFunctions.LocaleDecimalSeparator);
        var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");
  
        sThousandSeparator = "\\";
        sThousandSeparator = sThousandSeparator.concat(window.parent.frames("menuframe").ASRIntranetFunctions.LocaleThousandSeparator);
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
                    if (window.parent.frames("menuframe").ASRIntranetFunctions.LocaleDecimalSeparator != ".") {
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
                if (window.parent.frames("menuframe").ASRIntranetFunctions.LocaleDecimalSeparator != ".") {
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
                if (window.parent.frames("menuframe").ASRIntranetFunctions.LocaleDecimalSeparator != ".") {
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
                    if (window.parent.frames("menuframe").ASRIntranetFunctions.LocaleDecimalSeparator != ".") {
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
                        if (window.parent.frames("menuframe").ASRIntranetFunctions.LocaleDecimalSeparator != ".") {
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

    function saveChanges(psAction, pfPrompt, pfTBOverride)
    {
        // Expand the work frame and hide the option frame.
        $("#optionframe").attr("data-framesource", "UTIL_DEF_EXPRCOMPONENT");
        $("#optionframe").hide();
        $("#workframe").show();
        
        // Pass the component definition back to the expression page.        
        OpenHR.getFrame("workframe").saveChanges(psAction, pfPrompt, pfTBOverride);
    }
    -->
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
        OKClick();
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
        OKClick();        
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
        OKClick();        
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
        OKClick();        
    }

</script>


<FORM action="" method=POST id=frmMainForm name=frmMainForm>
<%
    Dim cmdParameter
    Dim prmFunctionID
    Dim prmParameterIndex
    Dim prmPassByType
    
    Dim iPassBy As Integer
    Dim sErrMsg As String
    
    
iPassBy = 1	
if (len(sErrMsg) = 0) and (Session("optionFunctionID") > 0) then
        cmdParameter = Server.CreateObject("ADODB.Command")
	cmdParameter.CommandText = "spASRIntGetParameterPassByType"
	cmdParameter.CommandType = 4 ' Stored Procedure
        cmdParameter.ActiveConnection = Session("databaseConnection")

        prmFunctionID = cmdParameter.CreateParameter("functionID", 3, 1) ' 3=integer, 1=input
        cmdParameter.Parameters.Append(prmFunctionID)
	prmFunctionID.value = cleanNumeric(clng(Session("optionFunctionID")))

        prmParameterIndex = cmdParameter.CreateParameter("parameterIndex", 3, 1) ' 3=integer, 1=input
        cmdParameter.Parameters.Append(prmParameterIndex)
	prmParameterIndex.value = cleanNumeric(clng(Session("optionParameterIndex")))

        prmPassByType = cmdParameter.CreateParameter("passByType", 3, 2) ' 3=integer, 2=output
        cmdParameter.Parameters.Append(prmPassByType)

        Err.Clear()
	cmdParameter.Execute
        If (Err.Number <> 0) Then
            sErrMsg = "Error checking parameter pass-by type." & vbCrLf & FormatError(Err.Description)
        Else
            iPassBy = cmdParameter.Parameters("passByType").Value
        End If

	' Release the ADO command object.
        cmdParameter = Nothing
end if
    Response.Write("<INPUT type='hidden' id=txtPassByType name=txtPassByType value=" & iPassBy & ">" & vbCrLf)
%>	

<table align=center class="outline" cellPadding=5 cellSpacing=0 width=100% height=100%>
	<TR>
		<TD>
			<TABLE WIDTH="100%" height="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
				<TR height=5>
					<TD height=5 colspan=5></td>
				</tr>
				
				<tr>
					<td width=10>&nbsp;&nbsp;</td>
					
					<TD width=10%>
						<TABLE height=100% width=100% class="outline" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD valign=top>
									<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0>
										<TR height=5>
											<TD colspan=5>&nbsp;&nbsp;</TD>
										</TR>
										
										<TR height=10>
											<TD width=5>&nbsp;</TD>
											<TD width=5><STRONG>Type</STRONG></TD>
											<TD colspan=3></TD>
										</TR>
										
										<TR height=10>
											<TD colspan=5></TD>
										</TR>
										
										<TR height=10>
											<TD width=5>&nbsp;</TD>
											<TD width=5>
												<input id=optType_Field name=optType type=radio selected
												    onclick="changeType(1)" 
                                                    onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
											</TD>
											<TD width=5>&nbsp;</TD>
											<TD nowrap>
                                                <label 
                                                    tabindex=-1
                                                    for="optType_Field"
                                                    class="radio"
                                                    onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                />
    												Field
                           	    		        </label>
											</TD>
											<TD width=5>&nbsp;&nbsp;</TD>
										</TR>
										
										<TR height=5>
											<TD colspan=5></TD>
										</TR>
										
										<TR height=10>
											<TD width=5>&nbsp;</TD>
											<TD width=5>
												<input id=optType_Operator name=optType type=radio
                                                    <%
												    If iPassBy = 2 Then
												        Response.Write("disabled")												        
												    End If
												    %>
												    onclick="changeType(5)" 
                                                    onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
											</TD>
											<TD width=5>&nbsp;</TD>
											<TD nowrap>
                                                <label 
                                                    tabindex=-1
                                                    for="optType_Operator"
                                                    class="radio"
                                                    onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                />
    												Operator
                           	    		        </label>
											</TD>
											<TD width=5>&nbsp;</TD>
										</TR>

										<TR height=5>
											<TD colspan=5></TD>
										</TR>

										<TR height=10>
											<TD width=5>&nbsp;</TD>
											<TD width=5>
												<input id=optType_Function name=optType type=radio 
                                                    <%
												    If iPassBy = 2 Then											        
												        Response.Write("disabled")
												    End If
												    %>
												    onclick="changeType(2)" 
                                                    onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
											</TD>
											<TD width=5>&nbsp;</TD>
											<TD nowrap>
                                                <label 
                                                    tabindex=-1
                                                    for="optType_Function"
                                                    class="radio"
                                                    onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                />
    												Function
                           	    		        </label>
											</TD>
											<TD width=5>&nbsp;</TD>
										</TR>

										<TR height=5>
											<TD colspan=5></TD>
										</TR>

										<TR height=10>
											<TD width=5>&nbsp;</TD>
											<TD width=5>
												<input id=optType_Value name=optType type=radio <%
												    If iPassBy = 2 Then
												        Response.Write("disabled")
												    End If
												    %>
												    onclick="changeType(4)" 
                                                    onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
											</TD>
											<TD width=5>&nbsp;</TD>
											<TD nowrap>
                                                <label 
                                                    tabindex=-1
                                                    for="optType_Value"
                                                    class="radio"
                                                    onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                />
    												Value
                           	    		        </label>
											</TD>
											<TD width=5>&nbsp;</TD>
										</TR>

										<TR height=5>
											<TD colspan=5></TD>
										</TR>

										<TR height=10>
											<TD width=5>&nbsp;</TD>
											<TD width=5>
                                                <input id="optType_LookupTableValue" name="optType" type="radio" <% 
                                                    If iPassBy = 2 Then
                                                        Response.Write("disabled")
                                                    End If
                                                    %>
                                                    onclick="changeType(6)"
                                                    onmouseover="try{radio_onMouseOver(this);}catch(e){}"
                                                    onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                    onblur="try{radio_onBlur(this);}catch(e){}" />
											</TD>
											<TD width=5>&nbsp;</TD>
											<TD nowrap>
                                                <label 
                                                    tabindex=-1
                                                    for="optType_LookupTableValue"
                                                    class="radio"
                                                    onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                />
    												Lookup Table Value
                           	    		        </label>
											</TD>
											<TD width=5>&nbsp;</TD>
										</TR>

										<TR height=5>
											<TD colspan=5></TD>
										</TR>

										<TR height=10 style="visibility:hidden;display:none" id=trType_PVal>
											<TD width=5>&nbsp;</TD>
											<TD width=5>
												<input id=optType_PromptedValue name=optType type=radio <%if iPassBy = 2 then
    Response.write("disabled")
												    End If%>
												    onclick="changeType(7)" 
                                                    onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
											</TD>
											<TD width=5>&nbsp;</TD>
											<TD nowrap>
                                                <label 
                                                    tabindex=-1
                                                    for="optType_PromptedValue"
                                                    class="radio"
                                                    onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                />
    												Prompted Value
                           	    		        </label>
											</TD>
											<TD width=5>&nbsp;</TD>
										</TR>

										<TR height=5 style="visibility:hidden;display:none" id=trType_PVal2>
											<TD colspan=5></TD>
										</TR>

										<TR height=10 style="visibility:hidden;display:none" id=trType_Calc>
											<TD width=5>&nbsp;</TD>
											<TD width=5>
												<input id=optType_Calculation name=optType type=radio <%  If iPassBy = 2 Then
												        Response.Write("disabled")
												    End If%>
												    onclick="changeType(3)" 
                                                    onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
											</TD>
											<TD width=5>&nbsp;</TD>
											<TD nowrap>
                                                <label 
                                                    tabindex=-1
                                                    for="optType_Calculation"
                                                    class="radio"
                                                    onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                />
    												Calculation
                           	    		        </label>
											</TD>
											<TD width=5>&nbsp;</TD>
										</TR>

										<TR height=5 style="visibility:hidden;display:none" id=trType_Calc2>
											<TD colspan=5></TD>
										</TR>

										<TR height=10 style="visibility:hidden;display:none" id=trType_Filter>
											<TD width=5>&nbsp;</TD>
											<TD width=5>
												<input id=optType_Filter name=optType type=radio <%
												    If iPassBy = 2 Then
												        Response.Write("disabled")
												    End If

												    %>
												    onclick="changeType(10)" 
                                                    onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
											</TD>
											<TD width=5>&nbsp;</TD>
											<TD nowrap>
                                                <label 
                                                    tabindex=-1
                                                    for="optType_Filter"
                                                    class="radio"
                                                    onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                />
    												Filter
                           	    		        </label>
											</TD>
											<TD width=5>&nbsp;</TD>
										</TR>

										<TR>
											<TD colspan=5></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
								
					<td width=10>&nbsp;&nbsp;</td>

					<TD>
						<TABLE height=100% width=100% class="outline" CELLSPACING=0 CELLPADDING=0>
							<TR height=100%>
								<TD valign=top>
									<DIV id=divField>
										<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0>
											<TR height=10>
												<TD colspan=6>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD colspan=4><STRONG>Field</STRONG></TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=6></TD>
											</TR>
<%if iPassBy = 1 then%>
											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD colspan=4>
													<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
														<TR>
															<TD>
																<input id=optField_Field name=optField type=radio selected
																    onclick="field_refreshTable()" 
                                                                    onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
															</TD>
															<TD width=5>&nbsp;</TD>
															<TD nowrap>
                                                                <label 
                                                                    tabindex=-1
                                                                    for="optField_Field"
                                                                    class="radio"
                                                                    onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                />
	    														    Field
                                           	    		        </label>
    														</TD>
															<TD width=20>&nbsp;&nbsp;&nbsp;</TD>
															<TD>
																<input id=optField_Count name=optField type=radio selected
																    onclick="field_refreshTable()" 
                                                                    onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
															</TD>
															<TD width=5>&nbsp;</TD>
															<TD nowrap>
                                                                <label 
                                                                    tabindex=-1
                                                                    for="optField_Count"
                                                                    class="radio"
                                                                    onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                />
															        Count
                                           	    		        </label>
															</TD>
															<TD width=20>&nbsp;&nbsp;&nbsp;</TD>
															<TD>
																<input id=optField_Total name=optField type=radio selected
																    onclick="field_refreshTable()"
                                                                    onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
															</TD>
															<TD width=5>&nbsp;</TD>
															<TD nowrap>
                                                                <label 
                                                                    tabindex=-1
                                                                    for="optField_Total"
                                                                    class="radio"
                                                                    onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                />
															        Total
                                           	    		        </label>
															</TD>
															<TD width=100%></TD>
														</TR>
													</TABLE>
												</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=6>&nbsp;&nbsp;</TD>
											</TR>
<%end if%>

											<TR height=10>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD width=10 nowrap>Table :</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD width=50%>
													<select id="cboFieldTable" name="cboFieldTable" class="combo" style="WIDTH: 100%" 
													    onchange="field_changeTable()"> 
													</select>
												</TD>
												<TD width=50%>&nbsp;</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=5>
												<TD colspan=6></TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD width=10 nowrap>Column :</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD width=50%>
													<select id=cboFieldColumn name=cboFieldColumn class="combo" style="WIDTH: 100%"> 
													</select>
													<select id=cboFieldDummyColumn name=cboFieldDummyColumn class="combo combodisabled" style="WIDTH: 100%;visibility:hidden;display:none" disabled="disabled"> 
													</select>
												</TD>
												<TD width=50%>&nbsp;</TD>
												<TD width=10>&nbsp;&nbsp;</TD>
											</TR>

<%if iPassBy = 1 then%>
											<TR height=5>
												<TD colspan=6>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD colspan=4>
													<TABLE WIDTH=100% height=100% class="outline" CELLSPACING=0 CELLPADDING=0>
														<TR>
															<TD>
																<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
																	<TR height=10>
																		<TD colspan=6></TD>
																	</TR>

																	<TR height=10>
																		<TD width=10>&nbsp;</TD>
																		<TD colspan=4><STRONG>Child Field Options</STRONG></TD>
																		<TD width=10>&nbsp;</TD>
																	</TR>

																	<TR height=10>
																		<TD colspan=6></TD>
																	</TR>

																	<TR height=10>
																		<TD width=10>&nbsp;</TD>
																		<TD colspan=4>
																			<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
																				<TR>
																					<TD>
																						<input id=optFieldRecSel_First name=optFieldRecSel type=radio 
																						    onclick="field_refreshChildFrame()"
                                                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                            onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                            onblur="try{radio_onBlur(this);}catch(e){}"/>
																					</TD>
																					<TD width=5>&nbsp;</TD>
																					<TD nowrap>
                                                                                        <label 
                                                                                            tabindex=-1
                                                                                            for="optFieldRecSel_First"
                                                                                            class="radio"
                                                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                                        />
																					        First
                                                                   	    		        </label>
																					</TD>
																					<TD width=20>&nbsp;&nbsp;&nbsp;</TD>
																					<TD>
																						<input id=optFieldRecSel_Last name=optFieldRecSel type=radio 
																						    onclick="field_refreshChildFrame()"
                                                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                            onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                            onblur="try{radio_onBlur(this);}catch(e){}"/>
																					</TD>
																					<TD width=5>&nbsp;</TD>
																					<TD nowrap>
                                                                                        <label 
                                                                                            tabindex=-1
                                                                                            for="optFieldRecSel_Last"
                                                                                            class="radio"
                                                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                                        />
																					        Last
                                                                   	    		        </label>
																					</TD>
																					<TD width=20>&nbsp;&nbsp;&nbsp;</TD>
																					<TD>
																						<input id=optFieldRecSel_Specific name=optFieldRecSel type=radio 
																						    onclick="field_refreshChildFrame()"
                                                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                            onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                            onblur="try{radio_onBlur(this);}catch(e){}"/>
																					</TD>
																					<TD width=5>&nbsp;</TD>
																					<TD nowrap>
																						<DIV id=divFieldRecSel_Specific style="visibility:hidden;display:none">
                                                                                            <label 
                                                                                                tabindex=-1
                                                                                                for="optFieldRecSel_Specific"
                                                                                                class="radio"
                                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                                            />
		    																					Specific
                                                                       	    		        </label>
        																				</DIV>
																					</TD>
																					<TD width=5>&nbsp;</TD>
																					<TD width=100%>
																						<INPUT id=txtFieldRecSel_Specific name=txtFieldRecSel_Specific class="text">	
																					</TD>
																				</TR>
																			</TABLE>
																		</TD>
																		<TD width=10>&nbsp;</TD>
																	</TR>
																	
																	<TR height=10>
																		<TD colspan=6></TD>
																	</TR>

																	<TR height=10>
																		<TD width=20>&nbsp;&nbsp;</TD>
																		<TD width=110 nowrap>Order :</TD>
																		<TD width=20>&nbsp;&nbsp;</TD>
																		<TD width=50%>
																			<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
																				<TR>
																					<TD>
																						<INPUT type="text" id=txtFieldRecOrder name=txtFieldRecOrder class="text textdisabled" style="WIDTH: 100%" disabled="disabled">
																					</TD>
																					<TD style="width:30px;">
																						<INPUT id=btnFieldRecOrder name=btnFieldRecOrder style="WIDTH: 100%" class="btn" type=button value="..."
																						    onclick="field_selectRecOrder()" 
		                                                                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                                                    onfocus="try{button_onFocus(this);}catch(e){}"
		                                                                                    onblur="try{button_onBlur(this);}catch(e){}" />
																					</TD>
																				</TR>
																			</TABLE>
																		</TD>
																		<TD width=30%>&nbsp;</TD>
																		<TD width=10>&nbsp;&nbsp;</TD>
																	</TR>
																	
																	<TR height=5>
																		<TD colspan=6></TD>
																	</TR>

																	<TR height=10>
																		<TD width=20>&nbsp;&nbsp;</TD>
																		<TD width=110 nowrap>Filter :</TD>
																		<TD width=20>&nbsp;&nbsp;</TD>
																		<TD width=50%>
																			<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
																				<TR>
																					<TD>
																						<INPUT type="text" id=txtFieldRecFilter name=txtFieldRecFilter class="text textdisabled" style="WIDTH: 100%" disabled="disabled">
																					</TD>
																					<TD width=30>
																						<INPUT id=btnFieldRecFilter name=btnFieldRecFilter class="btn" style="WIDTH: 100%" type=button value="..."
																						    onclick="field_selectRecFilter()" 
			                                                                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
			                                                                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
			                                                                                onfocus="try{button_onFocus(this);}catch(e){}"
			                                                                                onblur="try{button_onBlur(this);}catch(e){}" />
																					</TD>
																				</TR>
																			</TABLE>
																		</TD>
																		<TD width=30%>&nbsp;</TD>
																		<TD width=10>&nbsp;&nbsp;</TD>
																	</TR>

																	<TR height=10>
																		<TD colspan=6></TD>
																	</TR>
																</TABLE>
															</TD>
														</TR>
													</TABLE>
												</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=100%>
												<TD colspan=6>&nbsp;</TD>
											</TR>
<%end if%>
										</TABLE>
									</DIV>

									<DIV id=divFunction style="visibility:hidden;display:none">
										<TABLE height=100% width=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
											<TR height=10>
												<TD colspan=3>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD><STRONG>Function</STRONG></TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=3></TD>
											</TR>

											<TR>
												<TD width=10>&nbsp;</TD>
												<TD>
													<OBJECT classid="clsid:1C203F13-95AD-11D0-A84B-00A0247B735B" id=SSFunctionTree codebase="cabs/SStree.cab#version=1,0,2,24" style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:100%" VIEWASTEXT>
														<PARAM NAME="_ExtentX" VALUE="2646">
														<PARAM NAME="_ExtentY" VALUE="1323">
														<PARAM NAME="_Version" VALUE="65538">
														<PARAM NAME="BackColor" VALUE="-2147483643">
														<PARAM NAME="ForeColor" VALUE="-2147483640">
														<PARAM NAME="ImagesMaskColor" VALUE="12632256">
														<PARAM NAME="PictureBackgroundMaskColor" VALUE="12632256">
														<PARAM NAME="Appearance" VALUE="1">
														<PARAM NAME="BorderStyle" VALUE="0">
														<PARAM NAME="LabelEdit" VALUE="1">
														<PARAM NAME="LineStyle" VALUE="0">
														<PARAM NAME="LineType" VALUE="1">
														<PARAM NAME="MousePointer" VALUE="0">
														<PARAM NAME="NodeSelectionStyle" VALUE="2">
														<PARAM NAME="PictureAlignment" VALUE="0">
														<PARAM NAME="ScrollStyle" VALUE="0">
														<PARAM NAME="Style" VALUE="6">
														<PARAM NAME="IndentationStyle" VALUE="0">
														<PARAM NAME="TreeTips" VALUE="3">
														<PARAM NAME="PictureBackgroundStyle" VALUE="0">
														<PARAM NAME="Indentation" VALUE="38">
														<PARAM NAME="MaxLines" VALUE="1">
														<PARAM NAME="TreeTipDelay" VALUE="500">
														<PARAM NAME="ImageCount" VALUE="0">
														<PARAM NAME="ImageListIndex" VALUE="-1">
														<PARAM NAME="OLEDragMode" VALUE="0">
														<PARAM NAME="OLEDropMode" VALUE="0">
														<PARAM NAME="AllowDelete" VALUE="0">
														<PARAM NAME="AutoSearch" VALUE="0">
														<PARAM NAME="Enabled" VALUE="-1">
														<PARAM NAME="HideSelection" VALUE="0">
														<PARAM NAME="ImagesUseMask" VALUE="0">
														<PARAM NAME="Redraw" VALUE="-1">
														<PARAM NAME="UseImageList" VALUE="-1">
														<PARAM NAME="PictureBackgroundUseMask" VALUE="0">
														<PARAM NAME="HasFont" VALUE="0">
														<PARAM NAME="HasMouseIcon" VALUE="0">
														<PARAM NAME="HasPictureBackground" VALUE="0">
														<PARAM NAME="PathSeparator" VALUE="\">
														<PARAM NAME="TabStops" VALUE="32">
														<PARAM NAME="ImageList" VALUE="<None>">
														<PARAM NAME="LoadStyleRoot" VALUE="1">
														<PARAM NAME="Sorted" VALUE="0">
														<PARAM NAME="OnDemandDiscardBuffer" VALUE="10">
													</OBJECT>
												</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=3>&nbsp;</TD>
											</TR>
										</TABLE>
									</DIV>
									
									<DIV id=divOperator style="visibility:hidden;display:none">
										<TABLE height=100% width=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
											<TR height=10>
												<TD colspan=3>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD><STRONG>Operator</STRONG></TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=3></TD>
											</TR>

											<TR>
												<TD width=10>&nbsp;</TD>
												<TD>
													<OBJECT classid="clsid:1C203F13-95AD-11D0-A84B-00A0247B735B" id=SSOperatorTree codebase="cabs/SStree.cab#version=1,0,2,24" style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:100%" VIEWASTEXT>
														<PARAM NAME="_ExtentX" VALUE="2646">
														<PARAM NAME="_ExtentY" VALUE="1323">
														<PARAM NAME="_Version" VALUE="65538">
														<PARAM NAME="BackColor" VALUE="-2147483643">
														<PARAM NAME="ForeColor" VALUE="-2147483640">
														<PARAM NAME="ImagesMaskColor" VALUE="12632256">
														<PARAM NAME="PictureBackgroundMaskColor" VALUE="12632256">
														<PARAM NAME="Appearance" VALUE="1">
														<PARAM NAME="BorderStyle" VALUE="0">
														<PARAM NAME="LabelEdit" VALUE="1">
														<PARAM NAME="LineStyle" VALUE="0">
														<PARAM NAME="LineType" VALUE="1">
														<PARAM NAME="MousePointer" VALUE="0">
														<PARAM NAME="NodeSelectionStyle" VALUE="2">
														<PARAM NAME="PictureAlignment" VALUE="0">
														<PARAM NAME="ScrollStyle" VALUE="0">
														<PARAM NAME="Style" VALUE="6">
														<PARAM NAME="IndentationStyle" VALUE="0">
														<PARAM NAME="TreeTips" VALUE="3">
														<PARAM NAME="PictureBackgroundStyle" VALUE="0">
														<PARAM NAME="Indentation" VALUE="38">
														<PARAM NAME="MaxLines" VALUE="1">
														<PARAM NAME="TreeTipDelay" VALUE="500">
														<PARAM NAME="ImageCount" VALUE="0">
														<PARAM NAME="ImageListIndex" VALUE="-1">
														<PARAM NAME="OLEDragMode" VALUE="0">
														<PARAM NAME="OLEDropMode" VALUE="0">
														<PARAM NAME="AllowDelete" VALUE="0">
														<PARAM NAME="AutoSearch" VALUE="0">
														<PARAM NAME="Enabled" VALUE="-1">
														<PARAM NAME="HideSelection" VALUE="0">
														<PARAM NAME="ImagesUseMask" VALUE="0">
														<PARAM NAME="Redraw" VALUE="-1">
														<PARAM NAME="UseImageList" VALUE="-1">
														<PARAM NAME="PictureBackgroundUseMask" VALUE="0">
														<PARAM NAME="HasFont" VALUE="0">
														<PARAM NAME="HasMouseIcon" VALUE="0">
														<PARAM NAME="HasPictureBackground" VALUE="0">
														<PARAM NAME="PathSeparator" VALUE="\">
														<PARAM NAME="TabStops" VALUE="32">
														<PARAM NAME="ImageList" VALUE="<None>">
														<PARAM NAME="LoadStyleRoot" VALUE="1">
														<PARAM NAME="Sorted" VALUE="0">
														<PARAM NAME="OnDemandDiscardBuffer" VALUE="10">
													</OBJECT>
												</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=3>&nbsp;</TD>
											</TR>
										</TABLE>
									</DIV>

									<DIV id=divValue style="visibility:hidden;display:none">
										<TABLE height=100% width=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
											<TR height=10>
												<TD colspan=6>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD colspan=4><STRONG>Value</STRONG></TD>
												<TD width=20>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=6></TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD width=10 nowrap>Type :</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD width=50%>
													<select id=cboValueType name=cboValueType class="combo" style="WIDTH: 100%" onchange="value_changeType()"> 
														<OPTION value=1>Character
														<OPTION value=2>Numeric
														<OPTION value=3>Logic
														<OPTION value=4>Date
													</select>
												</TD>
												<TD width=50%>&nbsp;</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=5>
												<TD colspan=6></TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD width=10 nowrap>Value :</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD width=50%>
													<SELECT id=selectValue name='selectValue"' class="combo" style="WIDTH: 100%">
														<OPTION value=1>True</OPTION>
														<OPTION value=0>False</OPTION>
													</SELECT>
													<INPUT id=txtValue name=txtValue class="text" style="LEFT: 0px; POSITION: absolute; TOP: 0px; VISIBILITY: hidden; WIDTH: 0px">	
												</TD>
												<TD width=50%>&nbsp;</TD>
												<TD width=10>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=100%>
												<TD colspan=6>&nbsp;</TD>
											</TR>
										</TABLE>
									</DIV>

									<DIV id=divLookupValue style="visibility:hidden;display:none">
										<TABLE height=100% width=100% class="invisible"  CELLSPACING=0 CELLPADDING=0>
											<TR height=10>
												<TD colspan=6>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD colspan=4><STRONG>Lookup Table Value</STRONG></TD>
												<TD width=20>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=6></TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD width=10 nowrap>Table :</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD width=50%>
													<select id=cboLookupValueTable name=cboLookupValueTable class="combo" style="WIDTH: 100%" onchange="lookupValue_changeTable()"> 
													</select>
												</TD>
												<TD width=50%>&nbsp;</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=5>
												<TD colspan=6></TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD width=10 nowrap>Column :</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD width=50%>
													<select id=cboLookupValueColumn name=cboLookupValueColumn class="combo" style="WIDTH: 100%" onchange="lookupValue_changeColumn()"> 
													</select>
												</TD>
												<TD width=50%>&nbsp;</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=5>
												<TD colspan=6></TD>
											</TR>
											
											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD width=10 nowrap>Value :</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD width=50%>
													<select id=cboLookupValueValue name=cboLookupValueValue class="combo" style="WIDTH: 100%"> 
													</select>
												</TD>
												<TD width=50%>&nbsp;</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=6></TD>
											</TR>
											
											<TR height=10>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD colspan=4>
													<input type=text class="textwarning" id=txtValueNotInLookup name=txtValueNotInLookup value="<value> does not appear in <table>.<column>" style ="TEXT-ALIGN: left; WIDTH: 100%; visibility:hidden; display:none" readonly>
												</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=100%>
												<TD colspan=6>&nbsp;</TD>
											</TR>
										</TABLE>
									</DIV>

									<DIV id=divCalculation style="visibility:hidden;display:none">
										<TABLE height=100% width=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
											<TR height=10>
												<TD colspan=3>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD><STRONG>Calculation</STRONG></TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=3></TD>
											</TR>

											<TR>
												<TD width=10>&nbsp;</TD>
												<TD>
													<OBJECT classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" id=ssOleDBGridCalculations name=ssOleDBGridCalculations codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:100%">
														<PARAM NAME="ScrollBars" VALUE="4">
														<PARAM NAME="_Version" VALUE="196616">
														<PARAM NAME="DataMode" VALUE="2">
														<PARAM NAME="Cols" VALUE="0">
														<PARAM NAME="Rows" VALUE="0">
														<PARAM NAME="BorderStyle" VALUE="1">
														<PARAM NAME="RecordSelectors" VALUE="0">
														<PARAM NAME="GroupHeaders" VALUE="0">
														<PARAM NAME="ColumnHeaders" VALUE="0">
														<PARAM NAME="GroupHeadLines" VALUE="0">
														<PARAM NAME="HeadLines" VALUE="0">
														<PARAM NAME="FieldDelimiter" VALUE="(None)">
														<PARAM NAME="FieldSeparator" VALUE="(Tab)">
														<PARAM NAME="Col.Count" VALUE="2">
														<PARAM NAME="stylesets.count" VALUE="0">
														<PARAM NAME="TagVariant" VALUE="EMPTY">
														<PARAM NAME="UseGroups" VALUE="0">
														<PARAM NAME="HeadFont3D" VALUE="0">
														<PARAM NAME="Font3D" VALUE="0">
														<PARAM NAME="DividerType" VALUE="3">
														<PARAM NAME="DividerStyle" VALUE="1">
														<PARAM NAME="DefColWidth" VALUE="0">
														<PARAM NAME="BeveColorScheme" VALUE="2">
														<PARAM NAME="BevelColorFrame" VALUE="-2147483642">
														<PARAM NAME="BevelColorHighlight" VALUE="-2147483628">
														<PARAM NAME="BevelColorShadow" VALUE="-2147483632">
														<PARAM NAME="BevelColorFace" VALUE="-2147483633">
														<PARAM NAME="CheckBox3D" VALUE="-1">
														<PARAM NAME="AllowAddNew" VALUE="0">
														<PARAM NAME="AllowDelete" VALUE="0">
														<PARAM NAME="AllowUpdate" VALUE="0">
														<PARAM NAME="MultiLine" VALUE="0">
														<PARAM NAME="ActiveCellStyleSet" VALUE="">
														<PARAM NAME="RowSelectionStyle" VALUE="0">
														<PARAM NAME="AllowRowSizing" VALUE="0">
														<PARAM NAME="AllowGroupSizing" VALUE="0">
														<PARAM NAME="AllowColumnSizing" VALUE="0">
														<PARAM NAME="AllowGroupMoving" VALUE="0">
														<PARAM NAME="AllowColumnMoving" VALUE="0">
														<PARAM NAME="AllowGroupSwapping" VALUE="0">
														<PARAM NAME="AllowColumnSwapping" VALUE="0">
														<PARAM NAME="AllowGroupShrinking" VALUE="0">
														<PARAM NAME="AllowColumnShrinking" VALUE="0">
														<PARAM NAME="AllowDragDrop" VALUE="0">
														<PARAM NAME="UseExactRowCount" VALUE="-1">
														<PARAM NAME="SelectTypeCol" VALUE="0">
														<PARAM NAME="SelectTypeRow" VALUE="1">
														<PARAM NAME="SelectByCell" VALUE="-1">
														<PARAM NAME="BalloonHelp" VALUE="0">
														<PARAM NAME="RowNavigation" VALUE="1">
														<PARAM NAME="CellNavigation" VALUE="0">
														<PARAM NAME="MaxSelectedRows" VALUE="1">
														<PARAM NAME="HeadStyleSet" VALUE="">
														<PARAM NAME="StyleSet" VALUE="">
														<PARAM NAME="ForeColorEven" VALUE="0">
														<PARAM NAME="ForeColorOdd" VALUE="0">
														<PARAM NAME="BackColorEven" VALUE="16777215">
														<PARAM NAME="BackColorOdd" VALUE="16777215">
														<PARAM NAME="Levels" VALUE="1">
														<PARAM NAME="RowHeight" VALUE="503">
														<PARAM NAME="ExtraHeight" VALUE="0">
														<PARAM NAME="ActiveRowStyleSet" VALUE="">
														<PARAM NAME="CaptionAlignment" VALUE="2">
														<PARAM NAME="SplitterPos" VALUE="0">
														<PARAM NAME="SplitterVisible" VALUE="0">
														<PARAM NAME="Columns.Count" VALUE="2">
														<PARAM NAME="Columns(0).Width" VALUE="100000">
														<PARAM NAME="Columns(0).Visible" VALUE="-1">
														<PARAM NAME="Columns(0).Columns.Count" VALUE="1">
														<PARAM NAME="Columns(0).Caption" VALUE="Name">
														<PARAM NAME="Columns(0).Name" VALUE="Name">
														<PARAM NAME="Columns(0).Alignment" VALUE="0">
														<PARAM NAME="Columns(0).CaptionAlignment" VALUE="3">
														<PARAM NAME="Columns(0).Bound" VALUE="0">
														<PARAM NAME="Columns(0).AllowSizing" VALUE="1">
														<PARAM NAME="Columns(0).DataField" VALUE="Column 0">
														<PARAM NAME="Columns(0).DataType" VALUE="8">
														<PARAM NAME="Columns(0).Level" VALUE="0">
														<PARAM NAME="Columns(0).NumberFormat" VALUE="">
														<PARAM NAME="Columns(0).Case" VALUE="0">
														<PARAM NAME="Columns(0).FieldLen" VALUE="256">
														<PARAM NAME="Columns(0).VertScrollBar" VALUE="0">
														<PARAM NAME="Columns(0).Locked" VALUE="0">
														<PARAM NAME="Columns(0).Style" VALUE="0">
														<PARAM NAME="Columns(0).ButtonsAlways" VALUE="0">
														<PARAM NAME="Columns(0).RowCount" VALUE="0">
														<PARAM NAME="Columns(0).ColCount" VALUE="1">
														<PARAM NAME="Columns(0).HasHeadForeColor" VALUE="0">
														<PARAM NAME="Columns(0).HasHeadBackColor" VALUE="0">
														<PARAM NAME="Columns(0).HasForeColor" VALUE="0">
														<PARAM NAME="Columns(0).HasBackColor" VALUE="0">
														<PARAM NAME="Columns(0).HeadForeColor" VALUE="0">
														<PARAM NAME="Columns(0).HeadBackColor" VALUE="0">
														<PARAM NAME="Columns(0).ForeColor" VALUE="0">
														<PARAM NAME="Columns(0).BackColor" VALUE="0">
														<PARAM NAME="Columns(0).HeadStyleSet" VALUE="">
														<PARAM NAME="Columns(0).StyleSet" VALUE="">
														<PARAM NAME="Columns(0).Nullable" VALUE="1">
														<PARAM NAME="Columns(0).Mask" VALUE="">
														<PARAM NAME="Columns(0).PromptInclude" VALUE="0">
														<PARAM NAME="Columns(0).ClipMode" VALUE="0">
														<PARAM NAME="Columns(0).PromptChar" VALUE="95">
														<PARAM NAME="Columns(1).Width" VALUE="0">
														<PARAM NAME="Columns(1).Visible" VALUE="0">
														<PARAM NAME="Columns(1).Columns.Count" VALUE="1">
														<PARAM NAME="Columns(1).Caption" VALUE="id">
														<PARAM NAME="Columns(1).Name" VALUE="id">
														<PARAM NAME="Columns(1).Alignment" VALUE="0">
														<PARAM NAME="Columns(1).CaptionAlignment" VALUE="3">
														<PARAM NAME="Columns(1).Bound" VALUE="0">
														<PARAM NAME="Columns(1).AllowSizing" VALUE="1">
														<PARAM NAME="Columns(1).DataField" VALUE="Column 1">
														<PARAM NAME="Columns(1).DataType" VALUE="8">
														<PARAM NAME="Columns(1).Level" VALUE="0">
														<PARAM NAME="Columns(1).NumberFormat" VALUE="">
														<PARAM NAME="Columns(1).Case" VALUE="0">
														<PARAM NAME="Columns(1).FieldLen" VALUE="256">
														<PARAM NAME="Columns(1).VertScrollBar" VALUE="0">
														<PARAM NAME="Columns(1).Locked" VALUE="0">
														<PARAM NAME="Columns(1).Style" VALUE="0">
														<PARAM NAME="Columns(1).ButtonsAlways" VALUE="0">
														<PARAM NAME="Columns(1).RowCount" VALUE="0">
														<PARAM NAME="Columns(1).ColCount" VALUE="1">
														<PARAM NAME="Columns(1).HasHeadForeColor" VALUE="0">
														<PARAM NAME="Columns(1).HasHeadBackColor" VALUE="0">
														<PARAM NAME="Columns(1).HasForeColor" VALUE="0">
														<PARAM NAME="Columns(1).HasBackColor" VALUE="0">
														<PARAM NAME="Columns(1).HeadForeColor" VALUE="0">
														<PARAM NAME="Columns(1).HeadBackColor" VALUE="0">
														<PARAM NAME="Columns(1).ForeColor" VALUE="0">
														<PARAM NAME="Columns(1).BackColor" VALUE="0">
														<PARAM NAME="Columns(1).HeadStyleSet" VALUE="">
														<PARAM NAME="Columns(1).StyleSet" VALUE="">
														<PARAM NAME="Columns(1).Nullable" VALUE="1">
														<PARAM NAME="Columns(1).Mask" VALUE="">
														<PARAM NAME="Columns(1).PromptInclude" VALUE="0">
														<PARAM NAME="Columns(1).ClipMode" VALUE="0">
														<PARAM NAME="Columns(1).PromptChar" VALUE="95">
														<PARAM NAME="UseDefaults" VALUE="-1">
														<PARAM NAME="TabNavigation" VALUE="1">
														<PARAM NAME="_ExtentX" VALUE="17330">
														<PARAM NAME="_ExtentY" VALUE="1323">
														<PARAM NAME="_StockProps" VALUE="79">
														<PARAM NAME="Caption" VALUE="">
														<PARAM NAME="ForeColor" VALUE="0">
														<PARAM NAME="BackColor" VALUE="16777215">
														<PARAM NAME="Enabled" VALUE="-1">
														<PARAM NAME="DataMember" VALUE="">
														<PARAM NAME="Row.Count" VALUE="0">
													</OBJECT>
												</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<tr> 
											  <td colspan=3 height=10></td>
											</tr>

											<tr> 
												<TD width=10>&nbsp;</TD>
											    <td height="60"> 
													<TEXTAREA id=txtCalcDescription name=txtCalcDescription class="textarea disabled" tabindex="-1"
													    style="HEIGHT: 99%; WIDTH: 100%; " wrap=VIRTUAL disabled="disabled">
													</TEXTAREA>
												</td>
												<TD width=10>&nbsp;</TD>
											</tr>

											<tr> 
											  <td colspan=3 height=10></td>
											</tr>

											<tr> 
												<TD width=10>&nbsp;</TD>
											    <td height="10"> 
											    <input <% If Session("OnlyMine") Then Response.Write("checked")%> type="checkbox" name="chkOwnersCalcs" id="chkOwnersCalcs" value="chkOwnersCalcs" tabindex="-1"
                                                    onclick="calculationAndFilter_refresh();"
                                                    onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
                                                    onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                <label 
			                                        for="chkOwnersCalcs"
			                                        class="checkbox"
			                                        tabindex=0 
			                                        onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
				                                    onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
				                                    onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
	                                                onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
	                                                onblur="try{checkboxLabel_onBlur(this);}catch(e){}">

                                                    Only show calculations where owner is '<% =session("Username") %>'
                    		    		        </label>
												</td>
												<TD width=10>&nbsp;</TD>
											</tr>

											<TR height=10>
												<TD colspan=3>&nbsp;</TD>
											</TR>
										</TABLE>
									</DIV>
									
									<DIV id=divFilter style="visibility:hidden;display:none">
										<TABLE height=100% width=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
											<TR height=10>
												<TD colspan=3>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD><STRONG>Filter</STRONG></TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=3></TD>
											</TR>

											<TR>
												<TD width=10>&nbsp;</TD>
												<TD>
													<OBJECT classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" id=ssOleDBGridFilters name=ssOleDBGridFilters   codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:100%">
														<PARAM NAME="ScrollBars" VALUE="4">
														<PARAM NAME="_Version" VALUE="196616">
														<PARAM NAME="DataMode" VALUE="2">
														<PARAM NAME="Cols" VALUE="0">
														<PARAM NAME="Rows" VALUE="0">
														<PARAM NAME="BorderStyle" VALUE="1">
														<PARAM NAME="RecordSelectors" VALUE="0">
														<PARAM NAME="GroupHeaders" VALUE="0">
														<PARAM NAME="ColumnHeaders" VALUE="0">
														<PARAM NAME="GroupHeadLines" VALUE="0">
														<PARAM NAME="HeadLines" VALUE="0">
														<PARAM NAME="FieldDelimiter" VALUE="(None)">
														<PARAM NAME="FieldSeparator" VALUE="(Tab)">
														<PARAM NAME="Col.Count" VALUE="2">
														<PARAM NAME="stylesets.count" VALUE="0">
														<PARAM NAME="TagVariant" VALUE="EMPTY">
														<PARAM NAME="UseGroups" VALUE="0">
														<PARAM NAME="HeadFont3D" VALUE="0">
														<PARAM NAME="Font3D" VALUE="0">
														<PARAM NAME="DividerType" VALUE="3">
														<PARAM NAME="DividerStyle" VALUE="1">
														<PARAM NAME="DefColWidth" VALUE="0">
														<PARAM NAME="BeveColorScheme" VALUE="2">
														<PARAM NAME="BevelColorFrame" VALUE="-2147483642">
														<PARAM NAME="BevelColorHighlight" VALUE="-2147483628">
														<PARAM NAME="BevelColorShadow" VALUE="-2147483632">
														<PARAM NAME="BevelColorFace" VALUE="-2147483633">
														<PARAM NAME="CheckBox3D" VALUE="-1">
														<PARAM NAME="AllowAddNew" VALUE="0">
														<PARAM NAME="AllowDelete" VALUE="0">
														<PARAM NAME="AllowUpdate" VALUE="0">
														<PARAM NAME="MultiLine" VALUE="0">
														<PARAM NAME="ActiveCellStyleSet" VALUE="">
														<PARAM NAME="RowSelectionStyle" VALUE="0">
														<PARAM NAME="AllowRowSizing" VALUE="0">
														<PARAM NAME="AllowGroupSizing" VALUE="0">
														<PARAM NAME="AllowColumnSizing" VALUE="0">
														<PARAM NAME="AllowGroupMoving" VALUE="0">
														<PARAM NAME="AllowColumnMoving" VALUE="0">
														<PARAM NAME="AllowGroupSwapping" VALUE="0">
														<PARAM NAME="AllowColumnSwapping" VALUE="0">
														<PARAM NAME="AllowGroupShrinking" VALUE="0">
														<PARAM NAME="AllowColumnShrinking" VALUE="0">
														<PARAM NAME="AllowDragDrop" VALUE="0">
														<PARAM NAME="UseExactRowCount" VALUE="-1">
														<PARAM NAME="SelectTypeCol" VALUE="0">
														<PARAM NAME="SelectTypeRow" VALUE="1">
														<PARAM NAME="SelectByCell" VALUE="-1">
														<PARAM NAME="BalloonHelp" VALUE="0">
														<PARAM NAME="RowNavigation" VALUE="1">
														<PARAM NAME="CellNavigation" VALUE="0">
														<PARAM NAME="MaxSelectedRows" VALUE="1">
														<PARAM NAME="HeadStyleSet" VALUE="">
														<PARAM NAME="StyleSet" VALUE="">
														<PARAM NAME="ForeColorEven" VALUE="0">
														<PARAM NAME="ForeColorOdd" VALUE="0">
														<PARAM NAME="BackColorEven" VALUE="16777215">
														<PARAM NAME="BackColorOdd" VALUE="16777215">
														<PARAM NAME="Levels" VALUE="1">
														<PARAM NAME="RowHeight" VALUE="503">
														<PARAM NAME="ExtraHeight" VALUE="0">
														<PARAM NAME="ActiveRowStyleSet" VALUE="">
														<PARAM NAME="CaptionAlignment" VALUE="2">
														<PARAM NAME="SplitterPos" VALUE="0">
														<PARAM NAME="SplitterVisible" VALUE="0">
														<PARAM NAME="Columns.Count" VALUE="2">
														<PARAM NAME="Columns(0).Width" VALUE="100000">
														<PARAM NAME="Columns(0).Visible" VALUE="-1">
														<PARAM NAME="Columns(0).Columns.Count" VALUE="1">
														<PARAM NAME="Columns(0).Caption" VALUE="Name">
														<PARAM NAME="Columns(0).Name" VALUE="Name">
														<PARAM NAME="Columns(0).Alignment" VALUE="0">
														<PARAM NAME="Columns(0).CaptionAlignment" VALUE="3">
														<PARAM NAME="Columns(0).Bound" VALUE="0">
														<PARAM NAME="Columns(0).AllowSizing" VALUE="1">
														<PARAM NAME="Columns(0).DataField" VALUE="Column 0">
														<PARAM NAME="Columns(0).DataType" VALUE="8">
														<PARAM NAME="Columns(0).Level" VALUE="0">
														<PARAM NAME="Columns(0).NumberFormat" VALUE="">
														<PARAM NAME="Columns(0).Case" VALUE="0">
														<PARAM NAME="Columns(0).FieldLen" VALUE="256">
														<PARAM NAME="Columns(0).VertScrollBar" VALUE="0">
														<PARAM NAME="Columns(0).Locked" VALUE="0">
														<PARAM NAME="Columns(0).Style" VALUE="0">
														<PARAM NAME="Columns(0).ButtonsAlways" VALUE="0">
														<PARAM NAME="Columns(0).RowCount" VALUE="0">
														<PARAM NAME="Columns(0).ColCount" VALUE="1">
														<PARAM NAME="Columns(0).HasHeadForeColor" VALUE="0">
														<PARAM NAME="Columns(0).HasHeadBackColor" VALUE="0">
														<PARAM NAME="Columns(0).HasForeColor" VALUE="0">
														<PARAM NAME="Columns(0).HasBackColor" VALUE="0">
														<PARAM NAME="Columns(0).HeadForeColor" VALUE="0">
														<PARAM NAME="Columns(0).HeadBackColor" VALUE="0">
														<PARAM NAME="Columns(0).ForeColor" VALUE="0">
														<PARAM NAME="Columns(0).BackColor" VALUE="0">
														<PARAM NAME="Columns(0).HeadStyleSet" VALUE="">
														<PARAM NAME="Columns(0).StyleSet" VALUE="">
														<PARAM NAME="Columns(0).Nullable" VALUE="1">
														<PARAM NAME="Columns(0).Mask" VALUE="">
														<PARAM NAME="Columns(0).PromptInclude" VALUE="0">
														<PARAM NAME="Columns(0).ClipMode" VALUE="0">
														<PARAM NAME="Columns(0).PromptChar" VALUE="95">
														<PARAM NAME="Columns(1).Width" VALUE="0">
														<PARAM NAME="Columns(1).Visible" VALUE="0">
														<PARAM NAME="Columns(1).Columns.Count" VALUE="1">
														<PARAM NAME="Columns(1).Caption" VALUE="id">
														<PARAM NAME="Columns(1).Name" VALUE="id">
														<PARAM NAME="Columns(1).Alignment" VALUE="0">
														<PARAM NAME="Columns(1).CaptionAlignment" VALUE="3">
														<PARAM NAME="Columns(1).Bound" VALUE="0">
														<PARAM NAME="Columns(1).AllowSizing" VALUE="1">
														<PARAM NAME="Columns(1).DataField" VALUE="Column 1">
														<PARAM NAME="Columns(1).DataType" VALUE="8">
														<PARAM NAME="Columns(1).Level" VALUE="0">
														<PARAM NAME="Columns(1).NumberFormat" VALUE="">
														<PARAM NAME="Columns(1).Case" VALUE="0">
														<PARAM NAME="Columns(1).FieldLen" VALUE="256">
														<PARAM NAME="Columns(1).VertScrollBar" VALUE="0">
														<PARAM NAME="Columns(1).Locked" VALUE="0">
														<PARAM NAME="Columns(1).Style" VALUE="0">
														<PARAM NAME="Columns(1).ButtonsAlways" VALUE="0">
														<PARAM NAME="Columns(1).RowCount" VALUE="0">
														<PARAM NAME="Columns(1).ColCount" VALUE="1">
														<PARAM NAME="Columns(1).HasHeadForeColor" VALUE="0">
														<PARAM NAME="Columns(1).HasHeadBackColor" VALUE="0">
														<PARAM NAME="Columns(1).HasForeColor" VALUE="0">
														<PARAM NAME="Columns(1).HasBackColor" VALUE="0">
														<PARAM NAME="Columns(1).HeadForeColor" VALUE="0">
														<PARAM NAME="Columns(1).HeadBackColor" VALUE="0">
														<PARAM NAME="Columns(1).ForeColor" VALUE="0">
														<PARAM NAME="Columns(1).BackColor" VALUE="0">
														<PARAM NAME="Columns(1).HeadStyleSet" VALUE="">
														<PARAM NAME="Columns(1).StyleSet" VALUE="">
														<PARAM NAME="Columns(1).Nullable" VALUE="1">
														<PARAM NAME="Columns(1).Mask" VALUE="">
														<PARAM NAME="Columns(1).PromptInclude" VALUE="0">
														<PARAM NAME="Columns(1).ClipMode" VALUE="0">
														<PARAM NAME="Columns(1).PromptChar" VALUE="95">
														<PARAM NAME="UseDefaults" VALUE="-1">
														<PARAM NAME="TabNavigation" VALUE="1">
														<PARAM NAME="_ExtentX" VALUE="17330">
														<PARAM NAME="_ExtentY" VALUE="1323">
														<PARAM NAME="_StockProps" VALUE="79">
														<PARAM NAME="Caption" VALUE="">
														<PARAM NAME="ForeColor" VALUE="0">
														<PARAM NAME="BackColor" VALUE="16777215">
														<PARAM NAME="Enabled" VALUE="-1">
														<PARAM NAME="DataMember" VALUE="">
														<PARAM NAME="Row.Count" VALUE="0">
													</OBJECT>
												</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<tr> 
											  <td colspan=3 height=10></td>
											</tr>

											<tr> 
												<TD width=10>&nbsp;</TD>
											    <td height="60"> 
													<TEXTAREA id=txtFilterDescription name=txtFilterDescription class="textarea disabled" tabindex="-1"
													style="HEIGHT: 99%; WIDTH: 100%; " wrap=VIRTUAL disabled="disabled">
													</TEXTAREA>
												</td>
												<TD width=10>&nbsp;</TD>
											</tr>

											<tr> 
											  <td colspan=3 height=10></td>
											</tr>

											<tr>
                                                <td width="10">&nbsp;</td>
                                                <td height="10">
                                                    <input <% If Session("OnlyMine") Then Response.Write("checked")%> type="checkbox" name="chkOwnersFilters" id="chkOwnersFilters" value="chkOwnersFilters" tabindex="-1"
                                                        onclick="calculationAndFilter_refresh();"
                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                    <label
                                                        for="chkOwnersFilters"
                                                        class="checkbox"
                                                        tabindex="0"
                                                        onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                                        onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
                                                        onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
                                                        Only show filters where owner is '<% =session("Username") %>'
                                                    </label>
												</td>
												<TD width=10>&nbsp;</TD>
											</tr>

											<TR height=10>
												<TD colspan=3>&nbsp;</TD>
											</TR>
										</TABLE>
									</DIV>

									<DIV id=divPromptedValue style="visibility:hidden;display:none">
										<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0>
											<TR height=10>
												<TD colspan=6>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD colspan=4><STRONG>Prompted Value</STRONG></TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=6></TD>
											</TR>

											<TR height=10>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD width=10 nowrap>Prompt :</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD width=50%>
													<INPUT id=txtPrompt name=txtPrompt class="text" onkeyup="pVal_changePrompt()" style="WIDTH: 100%" maxlength=40>	
												</TD>
												<TD width=50%>&nbsp;</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=5>
												<TD colspan=6>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD colspan=4>
												<TABLE WIDTH=100% height=100% class="outline" CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD>
															<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
																<TR height=5>
																	<TD colspan=11></TD>
																</TR>

																<TR height=10>
																	<TD width=10>&nbsp;</TD>
																	<TD colspan=9><STRONG>Type</STRONG></TD>
																	<TD width=10>&nbsp;</TD>
																</TR>

																<TR height=10>
																	<TD colspan=11></TD>
																</TR>

																<TR height=10>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=40%>
																		<select id=cboPValType name=cboPValType class="combo" style="WIDTH: 100%" onchange="pVal_changeType()"> 
																			<OPTION value=1>Character
																			<OPTION value=2>Numeric
																			<OPTION value=3>Logic
																			<OPTION value=4>Date
																			<OPTION value=5>Lookup Value
																		</select>
																	</TD>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=10 nowrap id=tdPValSizePrompt name=tdPValSizePrompt>
																		Size :
																	</TD>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=30%>
																		<INPUT class="text" id=txtPValSize name=txtPValSize style="WIDTH: 100%">	
																	</TD>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=10 nowrap id=tdPValDecimalsPrompt name=tdPValDecimalsPrompt>
																		Decimals :
																	</TD>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=30%>
																		<INPUT class="text" id=txtPValDecimals name=txtPValDecimals style="WIDTH: 100%">	
																	</TD>
																	<TD width=10>&nbsp;&nbsp;</TD>
																</TR>
																
																<TR height=10>
																	<TD colspan=11></TD>
																</TR>
															</TABLE>
														</TD>
													</TR>
												</TABLE>

												</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=5>
												<TD colspan=6>&nbsp;</TD>
											</TR>

											<TR height=10 id=trPValFormat>
												<TD width=10>&nbsp;</TD>
												<TD colspan=4>
												<TABLE WIDTH=100% height=100% class="outline" CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD>
															<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
																<TR height=5>
																	<TD colspan=8></TD>
																</TR>

																<TR height=10>
																	<TD width=10>&nbsp;</TD>
																	<TD colspan=6><STRONG>Mask</STRONG></TD>
																	<TD width=10>&nbsp;</TD>
																</TR>

																<TR height=10>
																	<TD colspan=8></TD>
																</TR>

																<TR height=10>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=100% colspan=6>
																		<INPUT id=txtPValFormat name=txtPValFormat class="text" style="WIDTH: 100%">	
																	</TD>
																	<TD width=20>&nbsp;&nbsp;</TD>
																</TR>

																<TR height=5>
																	<TD colspan=8></TD>
																</TR>

																<TR height=10>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD nowrap width=5%>A - Uppercase</TD>
																	<TD width=20%>&nbsp;&nbsp;</TD>
																	<TD nowrap width=5%>9 - Numbers (0-9)</TD>
																	<TD width=20%>&nbsp;&nbsp;</TD>
																	<TD nowrap width=5%>B - Binary (0 or 1)</TD>
																	<TD></TD>
																</TR>

																<TR height=5>
																	<TD colspan=8></TD>
																</TR>

																<TR height=10>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD nowrap width=5%>a - Lowercase</TD>
																	<TD width=20%>&nbsp;&nbsp;</TD>
																	<TD nowrap width=5%># - Numbers, Symbols</TD>
																	<TD width=20%>&nbsp;&nbsp;</TD>
																	<TD nowrap width=5%>\ - Follow by any literal</TD>
																	<TD></TD>
																</TR>

																<TR height=10>
																	<TD colspan=11></TD>
																</TR>
															</TABLE>
														</TD>
													</TR>
												</TABLE>

												</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=5 id=trPValFormat2>
												<TD colspan=6>&nbsp;</TD>
											</TR>

											<TR height=10 id=trPValLookup style="visibility:hidden;display:none">
												<TD width=10>&nbsp;</TD>
												<TD colspan=4>
												<TABLE WIDTH=100% height=100% class="outline" CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD>
															<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
																<TR height=5>
																	<TD colspan=6></TD>
																</TR>

																<TR height=10>
																	<TD width=10>&nbsp;</TD>
																	<TD colspan=4><STRONG>Lookup Table Value</STRONG></TD>
																	<TD width=10>&nbsp;</TD>
																</TR>

																<TR height=10>
																	<TD colspan=6></TD>
																</TR>

																<TR height=10>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=10 nowrap>Table :</TD>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=50%>
																		<select id=cboPValTable name=cboPValTable class="combo" style="WIDTH: 100%" onchange="pVal_changeTable()"> 
																		</select>
																	</TD>
																	<TD width=50%>&nbsp;</TD>
																	<TD width=10>&nbsp;</TD>
																</TR>

																<TR height=5>
																	<TD colspan=6></TD>
																</TR>

																<TR height=10>
																	<TD width=10>&nbsp;</TD>
																	<TD width=10 nowrap>Column :</TD>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=50%>
																		<select id=cboPValColumn name=cboPValColumn class="combo" style="WIDTH: 100%" onchange="pVal_changeColumn()"> 
																		</select>
																	</TD>
																	<TD width=50%>&nbsp;</TD>
																	<TD width=10>&nbsp;&nbsp;</TD>
																</TR>

																<TR height=10>
																	<TD colspan=6></TD>
																</TR>
															</TABLE>
														</TD>
													</TR>
												</TABLE>

												</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=5 id=trPValLookup2>
												<TD colspan=6>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD colspan=4>
												<TABLE WIDTH=100% height=100% class="outline" CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD>
															<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
																<TR height=5>
																	<TD colspan=8></TD>
																</TR>

																<TR height=10>
																	<TD width=10>&nbsp;</TD>
																	<TD colspan=6><STRONG>Default Value</STRONG></TD>
																	<TD width=10>&nbsp;</TD>
																</TR>

																<TR height=10>
																	<TD colspan=8></TD>
																</TR>

																<TR height=10 id=trPValDateOptions>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=100% colspan=6>
																		<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
																			<TR height=10>
																				<TD width=5>
																					<input id=optPValDate_Explicit name=optPValDate type=radio selected
																					    onclick="pVal_changeDateOption(0)" 
		                                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                        onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                        onblur="try{radio_onBlur(this);}catch(e){}"/>
																				</TD>
																				<TD width=5>&nbsp;</TD>
																				<TD nowrap>
                                                                                    <label 
                                                                                        tabindex=-1
	                                                                                    for="optPValDate_Explicit"
	                                                                                    class="radio"
		                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                                            />
	    																				Explicit
                                                            	    		        </label>
    																			</TD>
																				<TD width=20>&nbsp;&nbsp;</TD>
																				<TD width=5>
																					<input id=optPValDate_MonthStart name=optPValDate type=radio 
																					    onclick="pVal_changeDateOption(2)"
		                                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                        onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                        onblur="try{radio_onBlur(this);}catch(e){}"/>
																				</TD>
																				<TD width=5>&nbsp;</TD>
																				<TD nowrap>
                                                                                    <label 
                                                                                        tabindex=-1
	                                                                                    for="optPValDate_MonthStart"
	                                                                                    class="radio"
		                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                                            />
    																					Month Start
                                                            	    		        </label>
																				</TD>
																				<TD width=20>&nbsp;&nbsp;</TD>
																				<TD width=5>
																					<input id=optPValDate_YearStart name=optPValDate type=radio 
																					    onclick="pVal_changeDateOption(4)"
		                                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                        onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                        onblur="try{radio_onBlur(this);}catch(e){}"/>
																				</TD>
																				<TD width=5>&nbsp;</TD>
																				<TD nowrap>
                                                                                    <label 
                                                                                        tabindex=-1
	                                                                                    for="optPValDate_YearStart"
	                                                                                    class="radio"
		                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                                            />
    																					Year Start
                                                            	    		        </label>
																				</TD>
																				<TD width=100%>&nbsp;</TD>
																			</TR>
																			
																			<TR height=10>
																				<TD width=5>
																					<input id=optPValDate_Current name=optPValDate type=radio 
																					    onclick="pVal_changeDateOption(1)"
		                                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                        onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                        onblur="try{radio_onBlur(this);}catch(e){}"/>
																				</TD>
																				<TD width=5>&nbsp;</TD>
																				<TD nowrap>
                                                                                    <label 
                                                                                        tabindex=-1
	                                                                                    for="optPValDate_Current"
	                                                                                    class="radio"
		                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                                            />
    																					Current
                                                            	    		        </label>
																				</TD>
																				<TD width=20>&nbsp;&nbsp;</TD>
																				<TD width=5>
																					<input id=optPValDate_MonthEnd name=optPValDate type=radio 
																					    onclick="pVal_changeDateOption(3)"
		                                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                        onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                        onblur="try{radio_onBlur(this);}catch(e){}"/>
																				</TD>
																				<TD width=5>&nbsp;</TD>
																				<TD nowrap>
                                                                                    <label 
                                                                                        tabindex=-1
	                                                                                    for="optPValDate_MonthEnd"
	                                                                                    class="radio"
		                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                                            />
    																					Month End
                                                            	    		        </label>
																				</TD>
																				<TD width=20>&nbsp;&nbsp;</TD>
																				<TD width=5>
																					<input id=optPValDate_YearEnd name=optPValDate type=radio 
																					    onclick="pVal_changeDateOption(5)"
		                                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                        onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                        onblur="try{radio_onBlur(this);}catch(e){}"/>
																				</TD>
																				<TD width=5>&nbsp;</TD>
																				<TD nowrap>
                                                                                    <label 
                                                                                        tabindex=-1
	                                                                                    for="optPValDate_YearEnd"
	                                                                                    class="radio"
		                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                                            />
    																					Year End
                                                            	    		        </label>
																				</TD>
																				<TD width=100%>&nbsp;</TD>
																			</TR>
																		</TABLE>
																	</TD>
																	<TD width=20>&nbsp;&nbsp;</TD>
																</TR>

																<TR height=10 id=trPValDateOptions2>
																	<TD colspan=8></TD>
																</TR>

																<TR height=10 id=trPValTextDefault>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=100% colspan=6>
																		<INPUT id=txtPValDefault name=txtPValDefault class="text" style="WIDTH: 100%">	
																	</TD>
																	<TD width=20>&nbsp;&nbsp;</TD>
																</TR>

																<TR height=10 id=trPValComboDefault style="visibility:hidden;display:none">
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=100% colspan=6>
																		<select id=cboPValDefault name=cboPValDefault style="WIDTH: 100%"> 
																		</select>
																	</TD>
																	<TD width=20>&nbsp;&nbsp;</TD>
																</TR>

																<TR height=10>
																	<TD colspan=11></TD>
																</TR>
															</TABLE>
														</TD>
													</TR>
												</TABLE>

												</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=100%>
												<TD colspan=6>&nbsp;</TD>
											</TR>
										</TABLE>
									</DIV>
								</TD>
							</TR>
						</TABLE>
					</TD>
					<td width=10>&nbsp;&nbsp;</td>
				</tr>
								
				<TR>
					<TD height=10 colspan=5></td>
				</tr>
				
				<tr height=10>
					<td width=10></td>
					<td colspan=3>
						<table WIDTH=100% class="invisible" CELLSPACING="0" CELLPADDING="0">
							<TR>
								<TD colspan=4>
								</TD>
							</TR>
							<tr>	
								<td>
								</td>
								<td width=10>
									<input id=cmdOK name=cmdOK type="button" class="btn" value="OK" style="WIDTH: 75px" width="75" 
									    onclick="OKClick()"
		                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                onfocus="try{button_onFocus(this);}catch(e){}"
		                                onblur="try{button_onBlur(this);}catch(e){}" />
								</td>
								<td width=40>
								</td>
								<td width=10>
									<input id="cmdCancel" name="cmdCancel" type="button" class="btn" value="Cancel" style="WIDTH: 75px" width="75" 
									    onclick="CancelClick()"
		                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                onfocus="try{button_onFocus(this);}catch(e){}"
		                                onblur="try{button_onBlur(this);}catch(e){}" />
								</td>
							</tr>			
						</table>
					</td>
					<td width=10></td>
				</tr>
				<TR>
					<TD height=10 colspan=7></td>
				</tr>
			</TABLE>
		</td>
	</tr>
</TABLE>
</FORM>

<FORM id=util_def_exprcomponent_frmUseful name=util_def_exprcomponent_frmUseful style="visibility:hidden;display:none">
	<INPUT type="hidden" id=txtUserName name=txtUserName value="<%=session("username")%>">
	<INPUT type="hidden" id=txtExprType name=txtExprType value=<%=session("optionExprType")%>>
	<INPUT type="hidden" id=txtExprID name=txtExprID value=<%=session("optionExprID")%>>
	<INPUT type="hidden" id=txtAction name=txtAction value=<%=session("optionAction")%>>
	<INPUT type="hidden" id=txtLinkRecordID name=txtLinkRecordID value=<%=session("optionLinkRecordID")%>>
	<INPUT type="hidden" id=txtTableID name=txtTableID value=<%=session("optionTableID")%>>

	<INPUT type="hidden" id=txtInitialising name=txtInitialising value=0>

	<INPUT type="hidden" id=txtChildFieldOrderID name=txtChildFieldOrderID value=0>
	<INPUT type="hidden" id=txtChildFieldFilterID name=txtChildFieldFilterID value=0>
	<INPUT type="hidden" id=txtChildFieldFilterHidden name=txtChildFieldFilterHidden value=0>

	<INPUT type="hidden" id=txtFunctionsLoaded name=txtFunctionsLoaded value=0>
	<INPUT type="hidden" id=txtOperatorsLoaded name=txtOperatorsLoaded value=0>
	<INPUT type="hidden" id=txtLookupTablesLoaded name=txtLookupTablesLoaded value=0>
	<INPUT type="hidden" id=txtPValLookupTablesLoaded name=txtPValLookupTablesLoaded value=0>
	
	<INPUT type="hidden" id=txtEnableSQL2000Functions name=txtEnableSQL2000Functions value=<%=session("EnableSQL2000Functions")%>>
</FORM>

<FORM action="util_def_exprComponent_Submit" method=post id=frmGotoOption name=frmGotoOption>
  	<%Html.RenderPartial("~/Views/Shared/gotoOption.ascx")%>
</FORM>

<FORM id=frmOriginalDefinition name=frmOriginalDefinition>
<%
    Dim sDefnString As String
    Dim sFieldTableID As String
    Dim sFieldColumnID As String
    Dim sLookupTableID As String
    Dim sLookupColumnID As String
    
	on error resume next
	sErrMsg = ""

	if session("optionAction") = "EDITEXPRCOMPONENT"	then
		sDefnString = Session("optionExtension")

        Response.Write("<INPUT type='hidden' id=txtComponentID name=txtComponentID value=" & componentParameter(sDefnString, "COMPONENTID") & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtType name=txtType value=" & componentParameter(sDefnString, "TYPE") & ">" & vbCrLf)
        sFieldTableID = componentParameter(sDefnString, "FIELDTABLEID")
        sFieldColumnID = componentParameter(sDefnString, "FIELDCOLUMNID")
        Response.Write("<INPUT type='hidden' id=txtFieldPassBy name=txtFieldPassBy value=" & componentParameter(sDefnString, "FIELDPASSBY") & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtFieldSelectionTableID name=txtFieldSelectionTableID value=" & componentParameter(sDefnString, "FIELDSELECTIONTABLEID") & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtFieldSelectionRecord name=txtFieldSelectionRecord value=" & componentParameter(sDefnString, "FIELDSELECTIONRECORD") & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtFieldSelectionLine name=txtFieldSelectionLine value=" & componentParameter(sDefnString, "FIELDSELECTIONLINE") & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtFieldSelectionOrderID name=txtFieldSelectionOrderID value=" & componentParameter(sDefnString, "FIELDSELECTIONORDERID") & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtFieldSelectionFilter name=txtFieldSelectionFilter value=" & componentParameter(sDefnString, "FIELDSELECTIONFILTER") & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtFunctionID name=txtFunctionID value=" & componentParameter(sDefnString, "FUNCTIONID") & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtCalculationID name=txtCalculationID value=" & componentParameter(sDefnString, "CALCULATIONID") & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtOperatorID name=txtOperatorID value=" & componentParameter(sDefnString, "OPERATORID") & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtValueType name=txtValueType value=" & componentParameter(sDefnString, "VALUETYPE") & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtValueCharacter name=txtValueCharacter value=""" & Replace(componentParameter(sDefnString, "VALUECHARACTER"), """", "&quot;") & """>" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtValueNumeric name=txtValueNumeric value=" & componentParameter(sDefnString, "VALUENUMERIC") & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtValueLogic name=txtValueLogic value=" & componentParameter(sDefnString, "VALUELOGIC") & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtValueDate name=txtValueDate value=" & componentParameter(sDefnString, "VALUEDATE") & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtPromptDescription name=txtPromptDescription value=""" & Replace(componentParameter(sDefnString, "PROMPTDESCRIPTION"), """", "&quot;") & """>" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtPromptMask name=txtPromptMask value=""" & Replace(componentParameter(sDefnString, "PROMPTMASK"), """", "&quot;") & """>" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtPromptSize name=txtPromptSize value=" & componentParameter(sDefnString, "PROMPTSIZE") & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtPromptDecimals name=txtPromptDecimals value=" & componentParameter(sDefnString, "PROMPTDECIMALS") & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtFunctionReturnType name=txtFunctionReturnType value=" & componentParameter(sDefnString, "FUNCTIONRETURNTYPE") & ">" & vbCrLf)
        sLookupTableID = componentParameter(sDefnString, "LOOKUPTABLEID")
        sLookupColumnID = componentParameter(sDefnString, "LOOKUPCOLUMNID")
        Response.Write("<INPUT type='hidden' id=txtFilterID name=txtFilterID value=" & componentParameter(sDefnString, "FILTERID") & ">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtFieldOrderName name=txtFieldOrderName value=""" & componentParameter(sDefnString, "FIELDSELECTIONORDERNAME") & """>" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtFieldFilterName name=txtFieldFilterName value=""" & componentParameter(sDefnString, "FIELDSELECTIONFILTERNAME") & """>" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtPromptDateType name=txtPromptDateType value=" & componentParameter(sDefnString, "PROMPTDATETYPE") & ">" & vbCrLf)
    Else
        Response.Write("<INPUT type='hidden' id=txtComponentID name=txtComponentID value=0>" & vbCrLf)
        sFieldTableID = Session("optionTableID")
		sFieldColumnID = 0
		sLookupTableID = 0
        sLookupColumnID = 0
        Response.Write("<INPUT type='hidden' id=txtFieldSelectionRecord name=txtFieldSelectionRecord value=1>" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtValueCharacter name=txtValueCharacter value="""">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtValueNumeric name=txtValueNumeric value=0>" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtValueLogic name=txtValueLogic value=""False"">" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtValueDate name=txtValueDate value="""">" & vbCrLf)
    End If

    Response.Write("<INPUT type='hidden' id=txtFieldTableID name=txtFieldTableID value=" & sFieldTableID & ">" & vbCrLf)
    Response.Write("<INPUT type='hidden' id=txtFieldColumnID name=txtFieldColumnID value=" & sFieldColumnID & ">" & vbCrLf)
    Response.Write("<INPUT type='hidden' id=txtLookupTableID name=txtLookupTableID value=" & sLookupTableID & ">" & vbCrLf)
    Response.Write("<INPUT type='hidden' id=txtLookupColumnID name=txtLookupColumnID value=" & sLookupColumnID & ">" & vbCrLf)
%>
</FORM>

<FORM id=frmTables name=frmTables>
<%
    Dim cmdTables
    Dim prmTableID
    Dim rstTables
    Dim iCount As Integer
    
	if len(sErrMsg) = 0 then
        cmdTables = Server.CreateObject("ADODB.Command")
		cmdTables.CommandText = "sp_ASRIntGetExprTables"
		cmdTables.CommandType = 4 ' Stored Procedure
        cmdTables.ActiveConnection = Session("databaseConnection")

        prmTableID = cmdTables.CreateParameter("tableID", 3, 1) ' 3=integer, 1=input
        cmdTables.Parameters.Append(prmTableID)
		prmTableID.value = cleanNumeric(session("optionTableID"))

        Err.Clear()
        rstTables = cmdTables.Execute
        If (Err.Number <> 0) Then
            sErrMsg = "Error reading component tables." & vbCrLf & FormatError(Err.Description)
        Else
            If rstTables.state <> 0 Then
                ' Read recordset values.
                iCount = 0
                Do While Not rstTables.EOF
                    iCount = iCount + 1
                    Response.Write("<INPUT type='hidden' id=txtTable_" & iCount & " name=txtTable_" & iCount & " value=""" & rstTables.fields("definitionString").value & """>" & vbCrLf)
                    rstTables.MoveNext()
                Loop

                ' Release the ADO recordset object.
                rstTables.close()
            End If
            rstTables = Nothing
        End If

		' Release the ADO command object.
        cmdTables = Nothing
	end if
%>
</FORM>

<FORM id=frmFunctions name=frmFunctions>
<%
    Dim cmdFunctions
    Dim rstFunctions
    
	if len(sErrMsg) = 0 then
        cmdFunctions = Server.CreateObject("ADODB.Command")
        cmdFunctions.CommandText = "sp_ASRIntGetExprFunctions"
		cmdFunctions.CommandType = 4 ' Stored Procedure
        cmdFunctions.ActiveConnection = Session("databaseConnection")

        prmTableID = cmdFunctions.CreateParameter("tableID", 3, 1) ' 3=integer, 1=input
        cmdFunctions.Parameters.Append(prmTableID)
		prmTableID.value = cleanNumeric(session("optionTableID"))

        Err.Clear()
        rstFunctions = cmdFunctions.Execute
        If (Err.Number <> 0) Then
            sErrMsg = "Error reading component functions." & vbCrLf & FormatError(Err.Description)
        Else
            If rstFunctions.state <> 0 Then
                ' Read recordset values.
                iCount = 0
                Do While Not rstFunctions.EOF
                    iCount = iCount + 1
                    Response.Write("<INPUT type='hidden' id=txtFunction_" & iCount & " name=txtFunction_" & iCount & " value=""" & rstFunctions.fields("definitionString").value & """>" & vbCrLf)
                    rstFunctions.MoveNext()
                Loop

                ' Release the ADO recordset object.
                rstFunctions.close()
            End If
            rstFunctions = Nothing
        End If

		' Release the ADO command object.
        cmdFunctions = Nothing
	end if
%>
</FORM>

<FORM id=frmFunctionParameters name=frmFunctionParameters>
<%
    Dim cmdFunctionParameters
    Dim rstFunctionParameters
    
	if len(sErrMsg) = 0 then
        cmdFunctionParameters = Server.CreateObject("ADODB.Command")
		cmdFunctionParameters.CommandText = "sp_ASRIntGetExprFunctionParameters"
		cmdFunctionParameters.CommandType = 4 ' Stored Procedure
        cmdFunctionParameters.ActiveConnection = Session("databaseConnection")

        Err.Clear()
        rstFunctionParameters = cmdFunctionParameters.Execute
        If (Err.Number <> 0) Then
            sErrMsg = "Error reading component functions." & vbCrLf & FormatError(Err.Description)
        Else
            If rstFunctionParameters.state <> 0 Then
                ' Read recordset values.
                iCount = 1
                Do While Not rstFunctionParameters.EOF
                    Response.Write("<INPUT type='hidden' id=txtFunctionParameters_" & rstFunctionParameters.fields("functionID").value & "_" & iCount & " name=txtFunctionParameters_" & rstFunctionParameters.fields("functionID").value & "_" & iCount & " value=""" & rstFunctionParameters.fields("parameterName").value & """>" & vbCrLf)
                    iCount = iCount + 1
                    rstFunctionParameters.MoveNext()
                Loop

                ' Release the ADO recordset object.
                rstFunctionParameters.close()
            End If
            rstFunctionParameters = Nothing
        End If

		' Release the ADO command object.
        cmdFunctionParameters = Nothing
	end if
%>
</FORM>

<FORM id=frmOperators name=frmOperators>
<%
    Dim cmdOperators
    Dim rstOperators
    
	if len(sErrMsg) = 0 then
        cmdOperators = Server.CreateObject("ADODB.Command")
		cmdOperators.CommandText = "sp_ASRIntGetExprOperators"
		cmdOperators.CommandType = 4 ' Stored Procedure
        cmdOperators.ActiveConnection = Session("databaseConnection")

        Err.Clear()
        rstOperators = cmdOperators.Execute
        If (Err.Number <> 0) Then
            sErrMsg = "Error reading component operators." & vbCrLf & FormatError(Err.Description)
        Else
            If rstOperators.state <> 0 Then
                ' Read recordset values.
                iCount = 0
                Do While Not rstOperators.EOF
                    iCount = iCount + 1
                    Response.Write("<INPUT type='hidden' id=txtOperator_" & iCount & " name=txtOperator_" & iCount & " value=""" & rstOperators.fields("definitionString").value & """>" & vbCrLf)
                    rstOperators.MoveNext()
                Loop

                ' Release the ADO recordset object.
                rstOperators.close()
            End If
            rstOperators = Nothing
        End If

		' Release the ADO command object.
        cmdOperators = Nothing
	end if
%>
</FORM>

<FORM id=frmCalcs name=frmCalcs>
<%
    Dim cmdCalcs
    Dim rstCalcs
    Dim prmExprID
    Dim prmBaseTableID
    
    
    
	if len(sErrMsg) = 0 then
        cmdCalcs = Server.CreateObject("ADODB.Command")
		cmdCalcs.CommandText = "sp_ASRIntGetExprCalcs"
		cmdCalcs.CommandType = 4 ' Stored Procedure
        cmdCalcs.ActiveConnection = Session("databaseConnection")

        prmExprID = cmdCalcs.CreateParameter("exprID", 3, 1) ' 3=integer, 1=input
        cmdCalcs.Parameters.Append(prmExprID)
		prmExprID.value = cleanNumeric(clng(session("optionExprID")))

        prmBaseTableID = cmdCalcs.CreateParameter("baseTableID", 3, 1) ' 3=integer, 1=input
        cmdCalcs.Parameters.Append(prmBaseTableID)
		prmBaseTableID.value = cleanNumeric(clng(session("optionTableID")))

        Err.Clear()
        rstCalcs = cmdCalcs.Execute
        If (Err.Number <> 0) Then
            sErrMsg = "Error reading component calculations." & vbCrLf & FormatError(Err.Description)
        Else
            If rstCalcs.state <> 0 Then
                ' Read recordset values.
                iCount = 0
                Do While Not rstCalcs.EOF
                    iCount = iCount + 1
                    Response.Write("<INPUT type='hidden' id=txtCalc_" & iCount & " name=txtCalc_" & iCount & " value=""" & Replace(rstCalcs.fields("definitionString").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtCalcDesc_" & iCount & " name=txtCalcDesc_" & iCount & " value=""" & Replace(rstCalcs.fields("description").value, """", "&quot;") & """>" & vbCrLf)
                    rstCalcs.MoveNext()
                Loop

                ' Release the ADO recordset object.
                rstCalcs.close()
            End If
            rstCalcs = Nothing
        End If

		' Release the ADO command object.
        cmdCalcs = Nothing
	end if
%>
</FORM>
	
<FORM id=frmFilters name=frmFilters>
<%
    Dim cmdFilters
    Dim rstFilters
    
	if len(sErrMsg) = 0 then
        cmdFilters = Server.CreateObject("ADODB.Command")
		cmdFilters.CommandText = "sp_ASRIntGetExprFilters"
		cmdFilters.CommandType = 4 ' Stored Procedure
        cmdFilters.ActiveConnection = Session("databaseConnection")

        prmExprID = cmdFilters.CreateParameter("exprID", 3, 1) ' 3=integer, 1=input
        cmdFilters.Parameters.Append(prmExprID)
		prmExprID.value = cleanNumeric(clng(session("optionExprID")))

        prmBaseTableID = cmdFilters.CreateParameter("baseTableID", 3, 1) ' 3=integer, 1=input
        cmdFilters.Parameters.Append(prmBaseTableID)
		prmBaseTableID.value = cleanNumeric(clng(session("optionTableID")))

        Err.Clear()
        rstFilters = cmdFilters.Execute
        If (Err.Number <> 0) Then
            sErrMsg = "Error reading component filters." & vbCrLf & FormatError(Err.Description)
        Else
            If rstFilters.state <> 0 Then
                ' Read recordset values.
                iCount = 0
                Do While Not rstFilters.EOF
                    iCount = iCount + 1
                    Response.Write("<INPUT type='hidden' id=txtFilter_" & iCount & " name=txtFilter_" & iCount & " value=""" & Replace(rstFilters.fields("definitionString").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtFilterDesc_" & iCount & " name=txtFilterDesc_" & iCount & " value=""" & Replace(rstFilters.fields("description").value, """", "&quot;") & """>" & vbCrLf)
                    rstFilters.MoveNext()
                Loop

                ' Release the ADO recordset object.
                rstFilters.close()
            End If
            rstFilters = Nothing
        End If

		' Release the ADO command object.
        cmdFilters = Nothing
	end if
%>
</FORM>
	
<FORM id=frmFieldRec name=frmFieldRec target="fieldRec" action="fieldRec" method=post style="visibility:hidden;display:none">
	<INPUT type="hidden" id=selectionType name=selectionType>
	<INPUT type="hidden" id=Hidden1 name=txtTableID>
	<INPUT type="hidden" id=selectedID name=selectedID>
</FORM>

<INPUT type='hidden' id=txtTicker name=txtTicker value=0>
<INPUT type='hidden' id=txtLastKeyFind name=txtLastKeyFind value="">

<%
    Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrMsg & """>" & vbCrLf)
%>


<script runat="server" language="vb">

function componentParameter(psDefnString, psParameter)
	dim iCharIndex
	dim sDefn
	
	sDefn = psDefnString
	
	iCharIndex = instr(sDefn, "	")
	if iCharIndex >= 0 then
		if psParameter = "COMPONENTID" then
			componentParameter = left(sDefn, iCharIndex - 1)
			exit function
		end if
		
		sDefn = mid(sDefn, iCharIndex + 1)
		iCharIndex = instr(sDefn, "	")
		if iCharIndex >= 0 then
			if psParameter = "EXPRID" then
				componentParameter = left(sDefn, iCharIndex - 1)
				exit function
			end if
			
			sDefn = mid(sDefn, iCharIndex + 1)
			iCharIndex = instr(sDefn, "	")
			if iCharIndex >= 0 then
				if psParameter = "TYPE" then
					componentParameter = left(sDefn, iCharIndex - 1)
					exit function
				end if
				
				sDefn = mid(sDefn, iCharIndex + 1)
				iCharIndex = instr(sDefn, "	")
				if iCharIndex >= 0 then
					if psParameter = "FIELDCOLUMNID" then
						componentParameter = left(sDefn, iCharIndex - 1)
						exit function
					end if
					
					sDefn = mid(sDefn, iCharIndex + 1)
					iCharIndex = instr(sDefn, "	")
					if iCharIndex >= 0 then
						if psParameter = "FIELDPASSBY" then
							componentParameter = left(sDefn, iCharIndex - 1)
							exit function
						end if
						
						sDefn = mid(sDefn, iCharIndex + 1)
						iCharIndex = instr(sDefn, "	")
						if iCharIndex >= 0 then
							if psParameter = "FIELDSELECTIONTABLEID" then
								componentParameter = left(sDefn, iCharIndex - 1)
								exit function
							end if
							
							sDefn = mid(sDefn, iCharIndex + 1)
							iCharIndex = instr(sDefn, "	")
							if iCharIndex >= 0 then
								if psParameter = "FIELDSELECTIONRECORD" then
									componentParameter = left(sDefn, iCharIndex - 1)
									exit function
								end if
								
								sDefn = mid(sDefn, iCharIndex + 1)
								iCharIndex = instr(sDefn, "	")
								if iCharIndex >= 0 then
									if psParameter = "FIELDSELECTIONLINE" then
										componentParameter = left(sDefn, iCharIndex - 1)
										exit function
									end if
									
									sDefn = mid(sDefn, iCharIndex + 1)
									iCharIndex = instr(sDefn, "	")
									if iCharIndex >= 0 then
										if psParameter = "FIELDSELECTIONORDERID" then
											componentParameter = left(sDefn, iCharIndex - 1)
											exit function
										end if
										
										sDefn = mid(sDefn, iCharIndex + 1)
										iCharIndex = instr(sDefn, "	")
										if iCharIndex >= 0 then
											if psParameter = "FIELDSELECTIONFILTER" then
												componentParameter = left(sDefn, iCharIndex - 1)
												exit function
											end if
											
											sDefn = mid(sDefn, iCharIndex + 1)
											iCharIndex = instr(sDefn, "	")
											if iCharIndex >= 0 then
												if psParameter = "FUNCTIONID" then
													componentParameter = left(sDefn, iCharIndex - 1)
													exit function
												end if
												
												sDefn = mid(sDefn, iCharIndex + 1)
												iCharIndex = instr(sDefn, "	")
												if iCharIndex >= 0 then
													if psParameter = "CALCULATIONID" then
														componentParameter = left(sDefn, iCharIndex - 1)
														exit function
													end if
													
													sDefn = mid(sDefn, iCharIndex + 1)
													iCharIndex = instr(sDefn, "	")
													if iCharIndex >= 0 then
														if psParameter = "OPERATORID" then
															componentParameter = left(sDefn, iCharIndex - 1)
															exit function
														end if
														
														sDefn = mid(sDefn, iCharIndex + 1)
														iCharIndex = instr(sDefn, "	")
														if iCharIndex >= 0 then
															if psParameter = "VALUETYPE" then
																componentParameter = left(sDefn, iCharIndex - 1)
																exit function
															end if
															
															sDefn = mid(sDefn, iCharIndex + 1)
															iCharIndex = instr(sDefn, "	")
															if iCharIndex >= 0 then
																if psParameter = "VALUECHARACTER" then
																	componentParameter = left(sDefn, iCharIndex - 1)
																	exit function
																end if
																
																sDefn = mid(sDefn, iCharIndex + 1)
																iCharIndex = instr(sDefn, "	")
																if iCharIndex >= 0 then
																	if psParameter = "VALUENUMERIC" then
																		componentParameter = left(sDefn, iCharIndex - 1)
																		exit function
																	end if
																	
																	sDefn = mid(sDefn, iCharIndex + 1)
																	iCharIndex = instr(sDefn, "	")
																	if iCharIndex >= 0 then
																		if psParameter = "VALUELOGIC" then
																			componentParameter = left(sDefn, iCharIndex - 1)
																			exit function
																		end if
																		
																		sDefn = mid(sDefn, iCharIndex + 1)
																		iCharIndex = instr(sDefn, "	")
																		if iCharIndex >= 0 then
																			if psParameter = "VALUEDATE" then
																				componentParameter = left(sDefn, iCharIndex - 1)
																				exit function
																			end if
																			
																			sDefn = mid(sDefn, iCharIndex + 1)
																			iCharIndex = instr(sDefn, "	")
																			if iCharIndex >= 0 then
																				if psParameter = "PROMPTDESCRIPTION" then
																					componentParameter = left(sDefn, iCharIndex - 1)
																					exit function
																				end if
																				
																				sDefn = mid(sDefn, iCharIndex + 1)
																				iCharIndex = instr(sDefn, "	")
																				if iCharIndex >= 0 then
																					if psParameter = "PROMPTMASK" then
																						componentParameter = left(sDefn, iCharIndex - 1)
																						exit function
																					end if
																					
																					sDefn = mid(sDefn, iCharIndex + 1)
																					iCharIndex = instr(sDefn, "	")
																					if iCharIndex >= 0 then
																						if psParameter = "PROMPTSIZE" then
																							componentParameter = left(sDefn, iCharIndex - 1)
																							exit function
																						end if
																						
																						sDefn = mid(sDefn, iCharIndex + 1)
																						iCharIndex = instr(sDefn, "	")
																						if iCharIndex >= 0 then
																							if psParameter = "PROMPTDECIMALS" then
																								componentParameter = left(sDefn, iCharIndex - 1)
																								exit function
																							end if
																							
																							sDefn = mid(sDefn, iCharIndex + 1)
																							iCharIndex = instr(sDefn, "	")
																							if iCharIndex >= 0 then
																								if psParameter = "FUNCTIONRETURNTYPE" then
																									componentParameter = left(sDefn, iCharIndex - 1)
																									exit function
																								end if
																								
																								sDefn = mid(sDefn, iCharIndex + 1)
																								iCharIndex = instr(sDefn, "	")
																								if iCharIndex >= 0 then
																									if psParameter = "LOOKUPTABLEID" then
																										componentParameter = left(sDefn, iCharIndex - 1)
																										exit function
																									end if
																									
																									sDefn = mid(sDefn, iCharIndex + 1)
																									iCharIndex = instr(sDefn, "	")
																									if iCharIndex >= 0 then
																										if psParameter = "LOOKUPCOLUMNID" then
																											componentParameter = left(sDefn, iCharIndex - 1)
																											exit function
																										end if
																										
																										sDefn = mid(sDefn, iCharIndex + 1)
																										iCharIndex = instr(sDefn, "	")
																										if iCharIndex >= 0 then
																											if psParameter = "FILTERID" then
																												componentParameter = left(sDefn, iCharIndex - 1)
																												exit function
																											end if
																											
																											sDefn = mid(sDefn, iCharIndex + 1)
																											iCharIndex = instr(sDefn, "	")
																											if iCharIndex >= 0 then
																												if psParameter = "EXPANDEDNODE" then
																													componentParameter = left(sDefn, iCharIndex - 1)
																													exit function
																												end if
																												
																												sDefn = mid(sDefn, iCharIndex + 1)
																												iCharIndex = instr(sDefn, "	")
																												if iCharIndex >= 0 then
																													if psParameter = "PROMPTDATETYPE" then
																														componentParameter = left(sDefn, iCharIndex - 1)
																														exit function
																													end if
																													
																													sDefn = mid(sDefn, iCharIndex + 1)
																													iCharIndex = instr(sDefn, "	")
																													if iCharIndex >= 0 then
																														if psParameter = "DESCRIPTION" then
																															componentParameter = left(sDefn, iCharIndex - 1)
																															exit function
																														end if
																														
																														sDefn = mid(sDefn, iCharIndex + 1)
																														iCharIndex = instr(sDefn, "	")
																														if iCharIndex >= 0 then
																															if psParameter = "FIELDTABLEID" then
																																componentParameter = left(sDefn, iCharIndex - 1)
																																exit function
																															end if
																															
																															sDefn = mid(sDefn, iCharIndex + 1)
																															iCharIndex = instr(sDefn, "	")
																															if iCharIndex >= 0 then
																																if psParameter = "FIELDSELECTIONORDERNAME" then
																																	componentParameter = left(sDefn, iCharIndex - 1)
																																	exit function
																																end if
																																
																																sDefn = mid(sDefn, iCharIndex + 1)
																																if psParameter = "FIELDSELECTIONFILTERNAME" then
																																	componentParameter = sDefn
																																	exit function
																																end if
																															end if
																														end if	
																													end if	
																												end if	
																											end if	
																										end if	
																									end if	
																								end if	
																							end if	
																						end if	
																					end if	
																				end if	
																			end if	
																		end if	
																	end if	
																end if	
															end if	
														end if	
													end if	
												end if	
											end if	
										end if	
									end if	
								end if	
							end if	
						end if	
					end if	
				end if	
			end if	
		end if	
	end if
	
	componentParameter = ""
end function

    </script>


<script type="text/javascript">
    util_def_exprcomponent_addhandlers();
    util_def_exprcomponent_onload();
</script>


