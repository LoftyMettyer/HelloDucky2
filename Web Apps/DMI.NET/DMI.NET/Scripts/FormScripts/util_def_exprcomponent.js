
var frmMainForm = OpenHR.getForm("optionframe", "frmMainForm");
var util_def_exprcomponent_frmOriginalDefinition = OpenHR.getForm("optionframe", "util_def_exprcomponent_frmOriginalDefinition");
var frmFieldRec = OpenHR.getForm("workframe", "frmFieldRec");
var _FunctionTreeLoaded = false;
var _OperatorTreeLoaded = false;

function util_def_exprcomponent_onload() {

	var options = {};
	options["themes"] = {
		"dots": true, "icons": false
		, "theme": "adv_themeroller",
		"url": window.ROOT + "Scripts/jquery/jstree/themes/adv_themeroller/style.css"
	};
	options["plugins"] = ["html_data", "ui", "contextmenu", "crrm", "hotkeys", "themes", "themeroller", "sort"];

	options["themeroller"] = {
		"item_leaf": false,
		"item_clsd": false,
		"item_open": false,
		"item": "ui-menu-item"
	};


		$("#SSOperatorTree").bind("loaded.jstree", function () {
		_OperatorTreeLoaded = true;
		checkTreesLoaded();
	})
		.bind("select_node.jstree", function () { componentRefreshControls(); });

	$("#SSFunctionTree").bind("loaded.jstree", function () {
			_FunctionTreeLoaded = true;
			checkTreesLoaded();
		})
		.bind("select_node.jstree", function () { componentRefreshControls(); });

	options["core"] = { 'check_callback': true };	// Must have - this enables inline renaming etc...

	$('#SSOperatorTree').jstree(options);
	$('#SSFunctionTree').jstree(options);


	$('#txtFieldRecSel_Specific').spinner({
		min: 1,
		max: 250,
		step: 1
	});

	$('#txtPValSize, #txtPValDecimals').spinner({
		min: 1,
		max: 250,
		step: 1
	});

	$("#SSOperatorTree").delegate("a", "dblclick", function () {
		component_OKClick();
	});

	$("#SSFunctionTree").delegate("a", "dblclick", function () {
		component_OKClick();
	});


	setTimeout('resizeForm()', 300);
}

function checkTreesLoaded() {
	if (_OperatorTreeLoaded && _FunctionTreeLoaded) {
		onload2();
	}
}

function resizeForm() {
	var formHeight = $('#optionframe').height();
	var titleHeight = $('#divOperator>h3').outerHeight(true);
	$('#SSOperatorTree, #SSFunctionTree').height(formHeight - titleHeight - 25);
}

function onload2() {

	var sUdfFunctionVisibility;
	var sUdfFunctionDisplay;

	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;
	
	$("#optionframe").attr("data-framesource", "UTIL_DEF_EXPRCOMPONENT");

	var newHeight = (screen.height) / 2;
	var newWidth = (screen.width) / 2;
	$('#optionframe').dialog({
		height: newHeight,
		width: newWidth,
		modal: true,
		resizable: false,
		buttons: [
					{
						text: "OK",
						click: function () {
							component_OKClick();
						},
						"class": "width6em",
						"id": "cmdOK"
					},
					{
						text: "Cancel",
						click: function () {
							component_CancelClick();
						},
						"class": "width6em",
						"id": "cmdCancel"
					}
		],
	});

	formatComponentTypeFrame();

	if (util_def_exprcomponent_frmUseful.txtAction.value == "EDITEXPRCOMPONENT") {
		// Load the component definition.
		loadComponentDefinition();
	}
	else {
		// Inserting/adding a new component.
		frmMainForm.optType_Field.checked = true;
		changeType(1);
	}

	sUdfFunctionVisibility = "visible";
	sUdfFunctionDisplay = "inline-block";

	if (frmMainForm.txtPassByType.value == 1) {
		// ReSharper disable once InconsistentNaming
		var divFieldRecSel_Specific = document.getElementById("divFieldRecSel_Specific");
		frmMainForm.optFieldRecSel_Specific.style.visibility = sUdfFunctionVisibility;
		frmMainForm.optFieldRecSel_Specific.style.display = sUdfFunctionDisplay;
		frmMainForm.txtFieldRecSel_Specific.style.visibility = sUdfFunctionVisibility;
		frmMainForm.txtFieldRecSel_Specific.style.display = sUdfFunctionDisplay;
		divFieldRecSel_Specific.style.visibility = sUdfFunctionVisibility;
		divFieldRecSel_Specific.style.display = sUdfFunctionDisplay;
	}

	// Set focus onto one of the form controls. 
	$('#optionframe #cmdCancel').focus();

}

function formatComponentTypeFrame() {
	var sTypePValVisibility;
	var sTypePValDisplay;
	var sTypeCalcVisibility;
	var sTypeCalcDisplay;
	var sTypeFilterVisibility;
	var sTypeFilterDisplay;

	sTypePValVisibility = "visible";
	sTypePValDisplay = "inline-block";
	sTypeCalcVisibility = "visible";
	sTypeCalcDisplay = "inline-block";
	sTypeFilterVisibility = "visible";
	sTypeFilterDisplay = "inline-block";

	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;

	switch (util_def_exprcomponent_frmUseful.txtExprType.value) {
		case "10":
			// Runtime Calculation
			break;
		case "11":
			// Runtime Filter
			break;
		case "14":
			// Utility Runtime Calculation - no calcs, no filters, no prompted values
			sTypePValVisibility = "hidden";
			sTypePValDisplay = "none";

			sTypeCalcVisibility = "hidden";
			sTypeCalcDisplay = "none";

			sTypeFilterVisibility = "hidden";
			sTypeFilterDisplay = "none";
			break;
	}

	document.getElementById('trType_PVal').style.visibility = sTypePValVisibility;
	document.getElementById('trType_PVal').style.display = sTypePValDisplay;

	document.getElementById('trType_Calc').style.visibility = sTypeCalcVisibility;
	document.getElementById('trType_Calc').style.display = sTypeCalcDisplay;

	document.getElementById('trType_Filter').style.visibility = sTypeFilterVisibility;
	document.getElementById('trType_Filter').style.display = sTypeFilterDisplay;
}

function loadComponentDefinition() {
	var iType;
	var i;
	var iIndex;

	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;

	iType = new Number(util_def_exprcomponent_frmOriginalDefinition.txtType.value);

	if (iType == 1) {
		// Field
		frmMainForm.optType_Field.checked = true;
		frmMainForm.optType_Field.focus();

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
		frmMainForm.optType_Function.focus();

		changeType(iType);

		util_def_exprcomponent_frmUseful.txtInitialising.value = 0;
		functionAndOperator_refresh();
		$('#SSFunctionTree').jstree('select_node', '#' + util_def_exprcomponent_frmOriginalDefinition.txtFunctionID.value);
		$('#cmdOK').button('enable');
	}

	if (iType == 3) {
		// Calculation.
		frmMainForm.optType_Calculation.checked = true;
		frmMainForm.optType_Calculation.focus();

		changeType(iType);

		util_def_exprcomponent_frmUseful.txtInitialising.value = 0;
		calculationsAndFilters_load();

		// Locate the current calc.
		locateGridRecord(util_def_exprcomponent_frmOriginalDefinition.txtCalculationID.value);
	}

	if (iType == 4) {
		// Value
		frmMainForm.optType_Value.checked = true;
		frmMainForm.optType_Value.focus();

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
			frmMainForm.txtValue.value = util_def_exprcomponent_frmOriginalDefinition.txtValueNumeric.value; // parseFloat(util_def_exprcomponent_frmOriginalDefinition.txtValueNumeric.value).toString();
			$('#frmMainForm #txtValue').autoNumeric('update');
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
			frmMainForm.txtValue.value = OpenHR.ConvertSQLDateToLocale(util_def_exprcomponent_frmOriginalDefinition.txtValueDate.value);
		}
	}

	if (iType == 5) {
		// Operator
		frmMainForm.optType_Operator.checked = true;
		frmMainForm.optType_Operator.focus();

		changeType(iType);

		util_def_exprcomponent_frmUseful.txtInitialising.value = 0;
		functionAndOperator_refresh();
		$('#SSOperatorTree').jstree('select_node', '#' + util_def_exprcomponent_frmOriginalDefinition.txtOperatorID.value);
		$('#cmdOK').button('enable');
	}

	if (iType == 6) {
		// Lookup Table Value
		frmMainForm.optType_LookupTableValue.checked = true;
		frmMainForm.optType_LookupTableValue.focus();

		changeType(iType);

		util_def_exprcomponent_frmUseful.txtInitialising.value = 0;
		lookupValue_refreshTable();
	}

	if (iType == 7) {
		// Prompted Value.
		frmMainForm.optType_PromptedValue.checked = true;
		frmMainForm.optType_PromptedValue.focus();

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
			frmMainForm.txtPValDefault.value = util_def_exprcomponent_frmOriginalDefinition.txtValueNumeric.value; //parseFloat(util_def_exprcomponent_frmOriginalDefinition.txtValueNumeric.value);
			$('#frmMainForm #txtPValDefault').autoNumeric('update');
		}
		if (util_def_exprcomponent_frmOriginalDefinition.txtValueType.value == 3) {
			// Logic
			iIndex = 0;
			for (i = 0; i < frmMainForm.cboPValDefault.options.length; i++) {
				if (frmMainForm.cboPValDefault.options[i].Value == util_def_exprcomponent_frmOriginalDefinition.txtValueLogic.value) {
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
					frmMainForm.txtPValDefault.value = OpenHR.ConvertSQLDateToLocale(util_def_exprcomponent_frmOriginalDefinition.txtValueDate.value);
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

			pVal_changeDateOption(util_def_exprcomponent_frmOriginalDefinition.txtPromptDateType.value);
		}
	}

	if (iType == 10) {
		// Filter.
		frmMainForm.optType_Filter.checked = true;
		frmMainForm.optType_Filter.focus();

		changeType(iType);

		util_def_exprcomponent_frmUseful.txtInitialising.value = 0;
		calculationsAndFilters_load();

		// Locate the current filter.
		locateGridRecord(util_def_exprcomponent_frmOriginalDefinition.txtFilterID.value);
	}
}

function changeType(piType) {
	var sFieldVisibility;
	var sFieldDisplay;
	var sFunctionVisibility;
	var sFunctionDisplay;
	var sOperatorVisibility;
	var sOperatorDisplay;
	var sValueVisibility;
	var sValueDisplay;
	var sLookupValueVisibility;
	var sLookupValueDisplay;
	var sCalculationVisibility;
	var sCalculationDisplay;
	var sFilterVisibility;
	var sFilterDisplay;
	var sPromptedValueVisibility;
	var sPromptedValueDisplay;

	sFieldVisibility = "hidden";
	sFieldDisplay = "none";
	sFunctionVisibility = "hidden";
	sFunctionDisplay = "none";
	sOperatorVisibility = "hidden";
	sOperatorDisplay = "none";
	sValueVisibility = "hidden";
	sValueDisplay = "none";
	sLookupValueVisibility = "hidden";
	sLookupValueDisplay = "none";
	sCalculationVisibility = "hidden";
	sCalculationDisplay = "none";
	sFilterVisibility = "hidden";
	sFilterDisplay = "none";
	sPromptedValueVisibility = "hidden";
	sPromptedValueDisplay = "none";

	if (piType == 1) {
		// Field
		sFieldVisibility = "visible";
		sFieldDisplay = "inline-block";
	}
	if (piType == 2) {
		// Function
		sFunctionVisibility = "visible";
		sFunctionDisplay = "inline-block";
	}
	if (piType == 3) {
		// Calculation
		sCalculationVisibility = "visible";
		sCalculationDisplay = "inline-block";
	}
	if (piType == 4) {
		// Value
		sValueVisibility = "visible";
		sValueDisplay = "inline-block";
	}
	if (piType == 5) {
		// Operator
		sOperatorVisibility = "visible";
		sOperatorDisplay = "inline-block";
	}
	if (piType == 6) {
		// Table Value
		sLookupValueVisibility = "visible";
		sLookupValueDisplay = "inline-block";
	}
	if (piType == 7) {
		// Prompted Value
		sPromptedValueVisibility = "visible";
		sPromptedValueDisplay = "inline-block";
	}
	if (piType == 10) {
		// Filter
		sFilterVisibility = "visible";
		sFilterDisplay = "inline-block";
	}

	document.getElementById('divField').style.visibility = sFieldVisibility;
	document.getElementById('divField').style.display = sFieldDisplay;
	document.getElementById('divFunction').style.visibility = sFunctionVisibility;
	document.getElementById('divFunction').style.display = sFunctionDisplay;
	document.getElementById('divOperator').style.visibility = sOperatorVisibility;
	document.getElementById('divOperator').style.display = sOperatorDisplay;
	document.getElementById('divValue').style.visibility = sValueVisibility;
	document.getElementById('divValue').style.display = sValueDisplay;
	document.getElementById('divLookupValue').style.visibility = sLookupValueVisibility;
	document.getElementById('divLookupValue').style.display = sLookupValueDisplay;
	document.getElementById('divCalculation').style.visibility = sCalculationVisibility;
	document.getElementById('divCalculation').style.display = sCalculationDisplay;
	document.getElementById('divFilter').style.visibility = sFilterVisibility;
	document.getElementById('divFilter').style.display = sFilterDisplay;
	document.getElementById('divPromptedValue').style.visibility = sPromptedValueVisibility;
	document.getElementById('divPromptedValue').style.display = sPromptedValueDisplay;

	initializeComponentControls(piType);
}

function initializeComponentControls(piType) {
	//button_disable(frmMainForm.cmdOK, false);
	$('#cmdOK').button('enable');

	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;

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

	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;
	var oOption;
	var sTableName;
	var sTableID;
	var sRelated;
	var sIsChild;
	var tableCollection = $("#frmExprTables")[0].elements;
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
		for (i = 0; i < tableCollection.length; i++) {
			fTableOK = (frmMainForm.txtPassByType.value == 2);
			if (fTableOK == false) {
				sRelated = tableParameter(tableCollection.item(i).value, "RELATED");
				sIsChild = tableParameter(tableCollection.item(i).value, "ISCHILD");

				fTableOK = (((sRelated == "1") && (frmMainForm.optField_Field.checked)) ||
				((sIsChild == "1") && ((frmMainForm.optField_Count.checked) || (frmMainForm.optField_Total.checked))));
			}

			if (fTableOK == true) {
				sTableName = tableParameter(tableCollection.item(i).value, "NAME");
				sTableID = tableParameter(tableCollection.item(i).value, "TABLEID");
				oOption = document.createElement("OPTION");
				frmMainForm.cboFieldTable.options.add(oOption);
				oOption.innerHTML = sTableName;
				oOption.Value = sTableID;
			}
		}
	}

	if (frmMainForm.cboFieldTable.options.length > 0) {
		iIndex = 0;
		for (i = 0; i < frmMainForm.cboFieldTable.options.length; i++) {
			if (frmMainForm.cboFieldTable.options[i].Value == sDefaultTableID) {
				iIndex = i;
				break;
			}

			if (frmMainForm.cboFieldTable.options[i].Value == util_def_exprcomponent_frmUseful.txtTableID.value) {
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
	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;
	var fInitialise = util_def_exprcomponent_frmUseful.txtInitialising.value;

	sDefaultColumnID = util_def_exprcomponent_frmOriginalDefinition.txtFieldColumnID.value;

	if ((fInitialise == 0) &&
			(frmMainForm.cboFieldColumn.selectedIndex >= 0)) {
		sDefaultColumnID = frmMainForm.cboFieldColumn.options[frmMainForm.cboFieldColumn.selectedIndex].Value;
	}

	if (frmMainForm.txtPassByType.value == 2) {
		frmMainForm.cboFieldColumn.style.display = "inline-block";
		frmMainForm.cboFieldDummyColumn.style.visibility = "hidden";
		frmMainForm.cboFieldDummyColumn.style.display = "none";
	}
	else {
		if (frmMainForm.optField_Count.checked == true) {
			frmMainForm.cboFieldColumn.style.display = "none";
			frmMainForm.cboFieldDummyColumn.style.visibility = "visible";
			frmMainForm.cboFieldDummyColumn.style.display = "inline-block";
		}
		else {
			frmMainForm.cboFieldColumn.style.display = "inline-block";
			frmMainForm.cboFieldDummyColumn.style.visibility = "hidden";
			frmMainForm.cboFieldDummyColumn.style.display = "none";
		}
	}

	// Clear the current contents of the dropdown list.
	while (frmMainForm.cboFieldColumn.options.length > 0) {
		frmMainForm.cboFieldColumn.options.remove(0);
	}

	if (frmMainForm.cboFieldTable.selectedIndex >= 0) {
		// Get the optionData page to get the columns for the current table.
		var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
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

		refreshOptionData();

	}
	else {
		// No table selected. Clear the column combo.
		field_refreshChildFrame();
	}
}

function field_refreshChildFrame() {

	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;
	var tableCollection = $("#frmExprTables")[0].elements;
	var i;
	var fIsChild;

	if (frmMainForm.txtPassByType.value == 1) {
		fIsChild = false;

		if (frmMainForm.cboFieldTable.selectedIndex >= 0) {
			if (tableCollection != null) {
				for (i = 0; i < tableCollection.length; i++) {
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
					frmMainForm.txtFieldRecSel_Specific.value = "1";
					$('#txtFieldRecSel_Specific').spinner('disable');
				}
				else {
					$('#txtFieldRecSel_Specific').spinner('enable');
				}

				button_disable(frmMainForm.btnFieldRecOrder, false);
			}
			else {
				frmMainForm.optFieldRecSel_First.checked = true;

				radio_disable(frmMainForm.optFieldRecSel_First, true);
				radio_disable(frmMainForm.optFieldRecSel_Last, true);
				radio_disable(frmMainForm.optFieldRecSel_Specific, true);
				frmMainForm.txtFieldRecSel_Specific.value = "1";
				//text_disable(frmMainForm.txtFieldRecSel_Specific, true);
				$('#txtFieldRecSel_Specific').spinner('disable');

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

			frmMainForm.txtFieldRecSel_Specific.value = "1";
			$('#txtFieldRecSel_Specific').spinner('disable');
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
		$('#cmdOK').button('disable');
	}
	else {
		$('#cmdOK').button('enable');
	}

}

function field_changeTable() {
	field_refreshColumn();

	if (frmMainForm.txtPassByType.value == 1) {
		frmMainForm.txtFieldRecFilter.value = "";
		frmMainForm.txtFieldRecOrder.value = "";
	}

	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;
	util_def_exprcomponent_frmUseful.txtChildFieldFilterID.value = 0;
	util_def_exprcomponent_frmUseful.txtChildFieldFilterHidden.value = "N";
	util_def_exprcomponent_frmUseful.txtChildFieldOrderID.value = 0;
}

function field_selectRecOrder() {

	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;
	var tableID = frmMainForm.cboFieldTable.options[frmMainForm.cboFieldTable.selectedIndex].Value;
	var currentID = util_def_exprcomponent_frmUseful.txtChildFieldOrderID.value;
	var newHeight = (screen.height) / 2;
	var newWidth = (screen.width) / 2;
	OpenHR.modalExpressionSelect("ORDER", tableID, currentID, function (id, name, access) {
		makeSelection("ORDER", id, name, access);
	}, newWidth - 40, newHeight - 160);
}

function field_selectRecFilter() {

	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;
	var tableID = frmMainForm.cboFieldTable.options[frmMainForm.cboFieldTable.selectedIndex].Value;
	var currentID = util_def_exprcomponent_frmUseful.txtChildFieldFilterID.value;
	var newHeight = (screen.height) / 2;
	var newWidth = (screen.width) / 2;
	OpenHR.modalExpressionSelect("FILTER", tableID, currentID, function (id, name, access) {
		makeSelection("FILTER", id, name, access);
	}, newWidth - 40, newHeight - 160);
}

function makeSelection(psType, psSelectedID, psSelectedName, psSelectedAccess) {
	//we are doing this for the order
	if (psType == 'ORDER') {
		$('#txtFieldRecOrder').val(psSelectedName);
		$('#txtChildFieldOrderID').val(psSelectedID);

		try {
			$('#btnFieldRecOrder').focus();
		}
		catch (e) {
		}
	}
	else {
		//we are doing this for the filter
		$('#txtFieldRecFilter').val(psSelectedName);
		$('#txtChildFieldFilterID').val(psSelectedID);

		//if its hidden, set the relevant textbox value
		if (psSelectedAccess == "HD") {
			$('#txtChildFieldFilterHidden').val('Y');
		}
		else {
			$('#txtChildFieldFilterHidden').val('');
		}

		try {
			$('#btnFieldRecFilter').focus();
		}
		catch (e) {
		}
	}

	return false;
}

function functionAndOperator_refresh() {
	// Load the function treeview with the functions.
	var i;
	var colCollection;
	var sName;
	var sID;
	var sCategory;
	var fCategoryDone;
	var trvTreeView;
	var sRootKey;
	var ctlLoadedFlag;
	var treeID;
	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;

	if (frmMainForm.optType_Function.checked == true) {
		trvTreeView = document.getElementById('SSFunctionTree');
		var frmFunctions = document.getElementById('frmFunctions');
		treeID = 'SSFunctionTree';
		colCollection = frmFunctions.elements;
		sRootKey = "FUNCTION_ROOT";
		ctlLoadedFlag = util_def_exprcomponent_frmUseful.txtFunctionsLoaded;
	}
	else {
		trvTreeView = document.getElementById('SSOperatorTree');
		treeID = 'SSOperatorTree';
		var frmOperators = document.getElementById('frmOperators');
		colCollection = frmOperators.elements;
		sRootKey = "OPERATOR_ROOT";
		ctlLoadedFlag = util_def_exprcomponent_frmUseful.txtOperatorsLoaded;
	}

	if (ctlLoadedFlag.value == 0) {
		// Clear the treeview.
		$(trvTreeView).html('');

		if (colCollection != null) {
			for (i = 0; i < colCollection.length; i++) {
				sName = functionAndOperatorParameter(colCollection.item(i).value, "NAME");
				sID = functionAndOperatorParameter(colCollection.item(i).value, "ID");
				sCategory = functionAndOperatorParameter(colCollection.item(i).value, "CATEGORY");

				fCategoryDone = ($(trvTreeView).find('#' + safeID(sCategory)).length > 0);

				if (fCategoryDone == false) {
					addNode(treeID, '#' + sRootKey, 'last', sCategory, safeID(sCategory), true);
				}

				// Add the function node.							
				addNode(treeID, '#' + safeID(sCategory), 'inside', sName, sID, false);
			}
		}

		ctlLoadedFlag.value = 1;
	}
	$('#SSFunctionTree').jstree("close_all");
	$('#SSOperatorTree').jstree("close_all");

	$('#cmdOK').button('disable');

	util_def_exprcomponent_frmUseful.txtInitialising.value = 0;

}

function componentRefreshControls() {
	// Enable the treeview only if there are items.	
	if ($.jstree._focused()._get_parent() != -1) {
		$('#cmdOK').button('enable');
	} else {
		$('#cmdOK').button('disable');
	}
}

function calculationAndFilter_refresh() {
	var iCurrentID;
	var grdGrid;
	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;

	iCurrentID = 0;

	if (frmMainForm.optType_Filter.checked == true) {
		grdGrid = $('#ssOleDBGridFilters');
	}
	else {
		grdGrid = $('#ssOleDBGridCalculations');
	}

	$.when(calculationsAndFilters_load()).then(function () {

		if (grdGrid.getGridParam("reccount") > 0) {
			iCurrentID = grdGrid.getGridParam('selrow');

			if (util_def_exprcomponent_frmUseful.txtInitialising.value == 1) {
				// Goto top record.
				var topRowID = grdGrid.getDataIDs()[0];
				grdGrid.jqGrid('setSelection', topRowID);
			}
			else {
				// Locate the last selected record. if not exist then select the top record.
				if (iCurrentID == null) {
					var gotoTopRow = grdGrid.getDataIDs()[0];
					grdGrid.jqGrid('setSelection', gotoTopRow);

					// Move scrollbar to top position at first row
					grdGrid.closest(".ui-jqgrid-bdiv").scrollTop(0);
				}
			}

			//button_disable(frmMainForm.cmdOK, false);
			$('#cmdOK').button('enable');
		}
		else {
			//button_disable(frmMainForm.cmdOK, true);
			$('#cmdOK').button('disable');
		}

		util_def_exprcomponent_frmUseful.txtInitialising.value = 0;

	});
}

function calculationsAndFilters_load() {

	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;

	// Load the calculations/filters grid with the calcs.
	var i;
	var colCollection;
	var sName;
	var sID;
	var sOwner;
	var sCurrentOwner = new String(util_def_exprcomponent_frmUseful.txtUserName.value);
	var grdGrid;
	var fOwners;
	var idBeforeClearingGrid;

	var dfd = new $.Deferred();

	sCurrentOwner = sCurrentOwner.toUpperCase();
	var formHeight;
	var titleHeight;
	var textareaHeight;
	var onlyMineHeight;
	var marginHeight;
	var gridHeight;
	if (frmMainForm.optType_Filter.checked == true) {
		grdGrid = $('#ssOleDBGridFilters');
		var frmFilters = document.getElementById('frmFilters');
		colCollection = frmFilters.elements;
		fOwners = frmMainForm.chkOwnersFilters.checked;
		formHeight = $('#optionframe').height();
		titleHeight = $('#divFilter>h3').outerHeight(true);
		textareaHeight = $('#txtFilterDescription').outerHeight();
		onlyMineHeight = $('#divFilter>label').outerHeight(true);
		marginHeight = 60;
		gridHeight = formHeight - titleHeight - textareaHeight - onlyMineHeight - marginHeight;
	}
	else {
		grdGrid = $('#ssOleDBGridCalculations');
		var frmCalcs = document.getElementById('frmCalcs');
		colCollection = frmCalcs.elements;
		fOwners = frmMainForm.chkOwnersCalcs.checked;
		formHeight = $('#optionframe').height();
		titleHeight = $('#divCalculation>h3').outerHeight(true);
		textareaHeight = $('#txtCalcDescription').outerHeight();
		onlyMineHeight = $('#divCalculation>label').outerHeight(true);
		marginHeight = 60;
		gridHeight = formHeight - titleHeight - textareaHeight - onlyMineHeight - marginHeight;
	}

	if (grdGrid.getGridParam("reccount") > 0) {
		idBeforeClearingGrid = grdGrid.getGridParam('selrow');
		grdGrid.jqGrid('clearGridData');
	}

	grdGrid.jqGrid({
		datatype: "local",
		colNames: ['id', 'Name'],
		colModel: [
			{ name: 'id', index: 'id', sorttype: "int", hidden: true },
			{ name: 'name', index: 'name' }
		],
		multiselect: false,
		autowidth: true,
		height: gridHeight,
		caption: '',
		onSelectRow: function () {
			ssOleDBGridCalculations_rowcolchange();
		},
		ondblClickRow: function () {
			ssOleDBGridCalculations_dblClick();
		},
		rowNum: colCollection.length, // Set the number of records to display
		ignoreCase: true, // This make the local-search and sorting of values be case insensitive...
		loadComplete: function () {
			// Highlight top row. this will be called when doing soring too.
			var ids = $(this).jqGrid("getDataIDs");
			if (ids && ids.length > 0)
				$(this).jqGrid("setSelection", ids[0]);
		}
	});

	if (colCollection != null) {
		for (i = 0; i < colCollection.length; i++) {
			if (colCollection.item(i).name.indexOf("Desc_") < 0) {
				sName = calculationAndFilterParameter(colCollection.item(i).value, "NAME");
				sID = calculationAndFilterParameter(colCollection.item(i).value, "EXPRID");
				sOwner = calculationAndFilterParameter(colCollection.item(i).value, "OWNER");

				sOwner = sOwner.toUpperCase();

				if ((fOwners == false) ||
						(sCurrentOwner == sOwner)) {

					// Add the grid records.
					grdGrid.jqGrid('addRowData', sID, { id: sID, name: sName });

					// Set the selected row
					if (idBeforeClearingGrid == sID) {
						grdGrid.jqGrid('setSelection', idBeforeClearingGrid);
					}

				}
			}
		}
	}

	dfd.resolve();
	return dfd.promise();

}

function value_changeType() {

	if ($('#frmMainForm #cboValueType').val() == "3") {
		$('#frmMainForm #txtValue').hide();
		$('#frmMainForm #selectValue').show();
	}
	else {
		$('#frmMainForm #txtValue').show();
		$('#frmMainForm #selectValue').hide();

		if($('#frmMainForm #cboValueType').val() == "4") {
			//Date value
			$('#txtValue').datepicker();
			$('#txtValue').on('change', function(sender) {
				if (OpenHR.IsValidDate(sender.target.value) == false && sender.target.value != "") {
					OpenHR.modalMessage("Invalid date value entered.");
				}
			});
		} else {
			$('#txtValue').datepicker('destroy');
			$('#txtValue').off('change');
		}
	}

	if ($('#frmMainForm #cboValueType').val() == "2") {
		$('#frmMainForm #txtValue').autoNumeric('init', {vMin: -99999999.9999999,vMax: 99999999.9999999, mDec: 7, aPad: false});
		$('#frmMainForm #txtValue').val(0);
	} else {
		frmMainForm.txtValue.value = "";
			try {
				$('#frmMainForm #txtValue').autoNumeric('destroy');
			}
			catch (e) { }
		}
	frmMainForm.selectValue.selectedIndex = 0;
}

function lookupValue_refreshTable() {

	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;
	var oOption;
	var sTableName;
	var sTableID;
	var sType;
	var tableCollection = $("#frmExprTables")[0].elements;
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
			for (i = 0; i < tableCollection.length; i++) {
				sType = tableParameter(tableCollection.item(i).value, "TYPE");

				if (sType == "3") {
					sTableName = tableParameter(tableCollection.item(i).value, "NAME");
					sTableID = tableParameter(tableCollection.item(i).value, "TABLEID");
					oOption = document.createElement("OPTION");
					frmMainForm.cboLookupValueTable.options.add(oOption);
					oOption.innerHTML = sTableName;
					oOption.Value = sTableID;
				}
			}
		}

		util_def_exprcomponent_frmUseful.txtLookupTablesLoaded.value = 1;
	}

	if (frmMainForm.cboLookupValueTable.options.length > 0) {
		iIndex = 0;
		for (i = 0; i < frmMainForm.cboLookupValueTable.options.length; i++) {
			if (frmMainForm.cboLookupValueTable.options[i].Value == sDefaultTableID) {
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

	lookupValue_refreshColumn();
}

function lookupValue_refreshColumn() {

	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;
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
		var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
		optionDataForm.txtOptionAction.value = "LOADEXPRLOOKUPCOLUMNS";
		optionDataForm.txtOptionTableID.value = frmMainForm.cboLookupValueTable.options[frmMainForm.cboLookupValueTable.selectedIndex].Value;
		optionDataForm.txtOptionColumnID.value = sDefaultColumnID;

		refreshOptionData();
	}
	else {
		combo_disable(frmMainForm.cboLookupValueColumn, true);

		while (frmMainForm.cboLookupValueValue.options.length > 0) {
			frmMainForm.cboLookupValueValue.options.remove(0);
		}

		combo_disable(frmMainForm.cboLookupValueValue, true);

		$('#cmdOK').button('disable');
	}
}

function lookupValue_refreshValues() {

	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;
	var sDefaultValue = "";
	var fInitialise = util_def_exprcomponent_frmUseful.txtInitialising.value;
	var iDataType;

	if (frmMainForm.cboLookupValueColumn.selectedIndex >= 0) {
		iDataType = columnParameter(frmMainForm.cboLookupValueColumn.options[frmMainForm.cboLookupValueColumn.selectedIndex].Value, "DATATYPE");

		if ((frmMainForm.cboLookupValueTable.options[frmMainForm.cboLookupValueTable.selectedIndex].Value == util_def_exprcomponent_frmOriginalDefinition.txtLookupTableID.value) &&
				(columnParameter(frmMainForm.cboLookupValueColumn.options[frmMainForm.cboLookupValueColumn.selectedIndex].Value, "COLUMNID") == util_def_exprcomponent_frmOriginalDefinition.txtLookupColumnID.value)) {

			if (iDataType == 11) {
				// Date type lookup column.
				sDefaultValue = util_def_exprcomponent_frmOriginalDefinition.txtValueDate.value;
				sDefaultValue = OpenHR.ConvertSQLDateToLocale(sDefaultValue);

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
		var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
		optionDataForm.txtOptionAction.value = "LOADEXPRLOOKUPVALUES";
		optionDataForm.txtOptionColumnID.value = columnParameter(frmMainForm.cboLookupValueColumn.options[frmMainForm.cboLookupValueColumn.selectedIndex].Value, "COLUMNID");
		optionDataForm.txtGotoLocateValue.value = sDefaultValue;

		refreshOptionData();
	}
	else {
		combo_disable(frmMainForm.cboLookupValueValue, true);

		$('#cmdOK').button('disable');
	}
}

function lookupValue_changeTable() {
	lookupValue_refreshColumn();
}

function lookupValue_changeColumn() {
	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;
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

	$('#frmMainForm #txtPValDefault').autoNumeric('destroy');
	$('#frmMainForm #txtPValDefault').removeClass('number');

	if (iPValType == 1) {
		// Character
		sSizeVisibility = "visible";
		sFormatVisibility = "visible";
		sFormatDisplay = "inline";
		sTextDefaultVisibility = "visible";
		sTextDefaultDisplay = "inline-block";
		text_disable(frmMainForm.txtPValDefault, false);
	}

	if (iPValType == 2) {
		// Numeric
		sSizeVisibility = "visible";
		sDecimalsVisibility = "visible";
		sTextDefaultVisibility = "visible";
		sTextDefaultDisplay = "inline-block";
		text_disable(frmMainForm.txtPValDefault, false);
		$('#frmMainForm #txtPValDefault').addClass('number');
		$('#frmMainForm #txtPValDefault').autoNumeric('init', {
			vMin: -99999999.9999999, vMax: 99999999.9999, mDec: 4, aPad: false
		});

		//Change prompted value sizes to match default value
		$('.number').on('keyup', function () {
			//txtPValSize
			//txtPValDecimals

			var newSize = $('#frmMainForm #txtPValDefault').val().length;
			if ($('#frmMainForm #txtPValSize').val() < newSize) $('#frmMainForm #txtPValSize').val(newSize);

			var decimalSeparator = OpenHR.LocaleDecimalSeparator();
			if ($('#frmMainForm #txtPValDefault').val().indexOf(decimalSeparator) > 0) {
				var newDecimals = $('#frmMainForm #txtPValDefault').val().split(decimalSeparator)[1].length;
				if ($('#frmMainForm #txtPValDefault').val().indexOf(decimalSeparator) > 0) {
					if ($('#frmMainForm #txtPValDecimals').val() < newDecimals) $('#frmMainForm #txtPValDecimals').val(newDecimals);
				}
			}
		});
	}

	if (iPValType == 3) {
		// Logic
		sComboDefaultVisibility = "visible";
		sComboDefaultDisplay = "inline-block";

		// Clear the current contents of the dropdown list.
		while (frmMainForm.cboPValDefault.options.length > 0) {
			frmMainForm.cboPValDefault.options.remove(0);
		}

		oOption = document.createElement("OPTION");
		frmMainForm.cboPValDefault.options.add(oOption);
		oOption.innerHTML = "True";
		oOption.Value = 1;

		oOption = document.createElement("OPTION");
		frmMainForm.cboPValDefault.options.add(oOption);
		oOption.innerHTML = "False";
		oOption.Value = 0;

		frmMainForm.cboPValDefault.selectedIndex = 0;
	}

	if (iPValType == 4) {
		// Date
		sTextDefaultVisibility = "visible";
		sTextDefaultDisplay = "inline-block";
		sDateOptionsVisibility = "visible";
		sDateOptionsDisplay = "inline";

		frmMainForm.optPValDate_Explicit.checked = true;
	}

	if (iPValType == 5) {
		// Lookup Table Value
		sLookupVisibility = "visible";
		sLookupDisplay = "inline";
		sComboDefaultVisibility = "visible";
		sComboDefaultDisplay = "inline-block";

		// Clear the current contents of the dropdown list.
		while (frmMainForm.cboPValDefault.options.length > 0) {
			frmMainForm.cboPValDefault.options.remove(0);
		}

		pVal_refreshTable();
	}

	$('#txtPValSize').parent().toggle(sSizeVisibility == 'visible');
	$('#tdPValSizePrompt').toggle(sSizeVisibility == 'visible');
	$('#txtPValDecimals').parent().toggle(sDecimalsVisibility == 'visible');
	$('#tdPValDecimalsPrompt').toggle(sDecimalsVisibility == 'visible');
	$('#trPValFormat').toggle(sFormatVisibility == 'visible');
	$('#trPValFormat2').toggle(sFormatVisibility == 'visible');
	$('#trPValLookup').toggle(sLookupVisibility == 'visible');
	$('#trPValLookup2').toggle(sLookupVisibility == 'visible');

	$('#trPValTextDefault').toggle(sTextDefaultVisibility == 'visible');
	$('#trPValComboDefault').toggle(sComboDefaultVisibility == 'visible');
	$('#trPValDateOptions').toggle(sDateOptionsVisibility == 'visible');
	$('#trPValDateOptions2').toggle(sDateOptionsVisibility == 'visible');

	if (iPValType == 4) {		
		$('#txtPValDefault').datepicker();
		$('#txtPValDefault').on('change', function (sender) {
			if (OpenHR.IsValidDate(sender.target.value) == false && sender.target.value != "") {
				OpenHR.modalMessage("Invalid date value entered.");
			}
		});
	} else {
		$('#txtPValDefault').datepicker('destroy');
		$('#txtPValDefault').off('change');
	}

	frmMainForm.txtPValDefault.value = "";

	pVal_changePrompt();
}

function validateDate(sender) {
	if (OpenHR.IsValidDate(sender.target.value) == false && sender.target.value != "") {
		var exprid = sender.target.id;
		OpenHR.modalPrompt("Invalid date value entered.", 0, "Error").then(function () {
			setTimeout("$('#" + exprid + "').focus()", 100);
			return false;
		});
	}
	return true;
}

function pVal_refreshTable() {
	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;
	var oOption;
	var sTableName;
	var sTableID;
	var sType;
	var tableCollection = $("#frmExprTables")[0].elements;
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
			for (i = 0; i < tableCollection.length; i++) {
				sType = tableParameter(tableCollection.item(i).value, "TYPE");

				if (sType == "3") {
					sTableName = tableParameter(tableCollection.item(i).value, "NAME");
					sTableID = tableParameter(tableCollection.item(i).value, "TABLEID");
					oOption = document.createElement("OPTION");
					frmMainForm.cboPValTable.options.add(oOption);
					oOption.innerHTML = sTableName;
					oOption.Value = sTableID;
				}
			}
		}

		util_def_exprcomponent_frmUseful.txtPValLookupTablesLoaded.value = 1;
	}

	if (frmMainForm.cboPValTable.options.length > 0) {
		iIndex = 0;
		for (i = 0; i < frmMainForm.cboPValTable.options.length; i++) {
			if (frmMainForm.cboPValTable.options[i].Value == sDefaultTableID) {
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

	pVal_refreshColumn();
}

function pVal_refreshColumn() {
	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;
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
		var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
		optionDataForm.txtOptionAction.value = "LOADEXPRLOOKUPCOLUMNS";
		optionDataForm.txtOptionTableID.value = frmMainForm.cboPValTable.options[frmMainForm.cboPValTable.selectedIndex].Value;
		optionDataForm.txtOptionColumnID.value = sDefaultColumnID;

		refreshOptionData();
	}
	else {
		combo_disable(frmMainForm.cboPValColumn, true);

		while (frmMainForm.cboPValDefault.options.length > 0) {
			frmMainForm.cboPValDefault.options.remove(0);
		}

		combo_disable(frmMainForm.cboPValDefault, true);
		$('#cmdOK').button('disable');
	}
}

function pVal_refreshValues() {
	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;
	var sDefaultValue = "";
	var fInitialise = util_def_exprcomponent_frmUseful.txtInitialising.value;
	var iDataType;

	if (frmMainForm.cboPValColumn.selectedIndex >= 0) {
		iDataType = columnParameter(frmMainForm.cboPValColumn.options[frmMainForm.cboPValColumn.selectedIndex].Value, "DATATYPE");

		if ((frmMainForm.cboPValTable.options[frmMainForm.cboPValTable.selectedIndex].Value == util_def_exprcomponent_frmOriginalDefinition.txtFieldTableID.value) &&
				(columnParameter(frmMainForm.cboPValColumn.options[frmMainForm.cboPValColumn.selectedIndex].Value, "COLUMNID") == util_def_exprcomponent_frmOriginalDefinition.txtFieldColumnID.value)) {

			if (iDataType == 11) {
				// Date type lookup column.
				sDefaultValue = util_def_exprcomponent_frmOriginalDefinition.txtValueCharacter.value;
				sDefaultValue = OpenHR.ConvertSQLDateToLocale(sDefaultValue);
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
		var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
		optionDataForm.txtOptionAction.value = "LOADEXPRLOOKUPVALUES";
		optionDataForm.txtOptionColumnID.value = columnParameter(frmMainForm.cboPValColumn.options[frmMainForm.cboPValColumn.selectedIndex].Value, "COLUMNID");
		optionDataForm.txtGotoLocateValue.value = sDefaultValue;

		refreshOptionData();
	}
	else {
		combo_disable(frmMainForm.cboPValDefault, true);
		$('#cmdOK').button('disable');
	}
}

function pVal_changePrompt() {
	if (frmMainForm.optType_PromptedValue.checked == true) {
		if (frmMainForm.txtPrompt.value.length == 0) {
			$('#cmdOK').button('disable');
		}
		else {
			$('#cmdOK').button('enable');
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
	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;
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
	oOption = document.createElement("option");
	cboCombo.options.add(oOption);
	oOption.innerHTML = sColumnName;
	oOption.Value = psDefn;


}

function component_setColumn(piColumnID) {

	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;
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
		for (i = 0; i < cboCombo.options.length; i++) {
			if (columnParameter(cboCombo.options[i].Value, "COLUMNID") == piColumnID) {
				iIndex = i;
				break;
			}
		}

		cboCombo.selectedIndex = iIndex;

		combo_disable(cboCombo, false);
	}
	else {
		combo_disable(cboCombo, true);
		$('#cmdOK').button('disable');
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
	if (columnParameter(cboColumnCombo.options[cboColumnCombo.selectedIndex].Value, "DATATYPE") == 11) {
		// Date type lookup column.
		psValue = OpenHR.ConvertSQLDateToLocale(psValue);
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

		var oOption = document.createElement("OPTION");
		cboValueCombo.options.add(oOption);
		oOption.innerHTML = psValue;
		oOption.Value = psValue;
	}
}

function component_setValue(psValue) {
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

	for (i = 0; i < cboCombo.options.length; i++) {
		if (cboCombo.options[i].Value == psValue) {
			cboCombo.selectedIndex = i;
			fFound = true;
			break;
		}
	}

	combo_disable(cboCombo, false);

	if (frmMainForm.optType_LookupTableValue.checked == true) {
		$('#cmdOK').button('enable');
	}

	if (fFound == false) {
		if ((frmMainForm.optType_LookupTableValue.checked == true) &&
				(psValue.length > 0)) {

			var oOption = document.createElement("OPTION");
			cboCombo.options.add(oOption);
			oOption.innerHTML = psValue;
			oOption.Value = psValue;

			cboCombo.selectedIndex = cboCombo.options.length - 1;

			frmMainForm.txtValueNotInLookup.value = psValue +
					" does not appear in " +
					frmMainForm.cboLookupValueTable.options[frmMainForm.cboLookupValueTable.selectedIndex].text +
					"." +
					frmMainForm.cboLookupValueColumn.options[frmMainForm.cboLookupValueColumn.selectedIndex].text;
			sVisibility = "visible";
			sDisplay = "inline-block";
		}
		else {
			if (cboCombo.options.length > 0) {
				cboCombo.selectedIndex = 0;
			}
			else {
				combo_disable(cboCombo, true);
				if (frmMainForm.optType_LookupTableValue.checked == true) {
					$('#cmdOK').button('disable');
				}
			}
		}
	}

	frmMainForm.txtValueNotInLookup.style.visibility = sVisibility;
	frmMainForm.txtValueNotInLookup.style.display = sDisplay;

}

function columnParameter(psDefnString, psParameter) {
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

function tableParameter(psDefnString, psParameter) {
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

function functionAndOperatorParameter(psDefnString, psParameter) {
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

function calculationAndFilterParameter(psDefnString, psParameter) {
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

function locateGridRecord(piID) {
	var fFound;
	var grdGrid;

	if (frmMainForm.optType_Filter.checked == true) {
		grdGrid = $('#ssOleDBGridFilters');
	}
	else {
		grdGrid = $('#ssOleDBGridCalculations');
	}

	try {
		grdGrid.jqGrid('setSelection', piID);
		fFound = true;
	} catch (e) {
		fFound = false;
	}


	if ((fFound == false) && (grdGrid.getGridParam("reccount") > 0)) {
		// Goto top record.
		var topRowID = grdGrid.getDataIDs()[0];
		grdGrid.jqGrid('setSelection', topRowID);
	}
}

function component_OKClick() {

	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;
	var sDefn;
	var sTemp;
	var sKey;
	var i;
	var iIndex;
	var iDataType;
	var fIsChild = false;
	var tableCollection = $("#frmExprTables")[0].elements;
	var sFunctionParameters;
	var frmFunctionParameters = document.getElementById('frmFunctionParameters');
	var colFunctionParameters = frmFunctionParameters.elements;

	if ($('#optionframe #cmdCancel').hasClass('ui-state-disabled')) {
		return;
	}
	
	if ((frmMainForm.optType_Function.checked == true) || (frmMainForm.optType_Operator.checked == true)) {
		//Has user selected a valid item
		if ($.jstree._focused().get_selected().hasClass('toplevelnode')) {
			//expand selected node and get out of here!
			var selectedNode = $.jstree._focused().get_selected()[0].id;
			$.jstree._focused().open_node('#' + selectedNode);
			return false;
		}

		//Has user selected more than one item
		if ($.jstree._focused().get_selected().length != 1) {
			OpenHR.modalMessage('Please select one item');
			return false;
		}
	}


	if (validateComponent() == true) {
		// Component definition is valid. Pass it back to the expression page.

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
			var iFieldSelection = 1;
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
			sDefn = sDefn + $.jstree._focused().get_selected().attr('id') + "\t";
		}
		else {
			sDefn = sDefn + "	";
		}

		if (frmMainForm.optType_Calculation.checked == true) {
			// Calculation ID.
			sDefn = sDefn + $('#ssOleDBGridCalculations').getGridParam('selrow') + "\t";
		}
		else {
			sDefn = sDefn + "	";
		}

		if (frmMainForm.optType_Operator.checked == true) {
			// Operator ID.
			sDefn = sDefn + $.jstree._focused().get_selected().attr('id') + "\t"; //frmMainForm.SSOperatorTree.SelectedItem.Key + "	";
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
				sDefn = sDefn + OpenHR.convertLocaleDateToSQL(frmMainForm.txtValue.value) + "	";
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
							sDefn = sDefn + OpenHR.convertLocaleDateToSQL(frmMainForm.cboPValDefault.options[frmMainForm.cboPValDefault.selectedIndex].Value) + "	";
						}
						else {
							// Character /Numeric/integer type lookup column.
							if (frmMainForm.cboPValDefault.selectedIndex >= 0) {
								sDefn = sDefn + frmMainForm.cboPValDefault.options[frmMainForm.cboPValDefault.selectedIndex].Value + "	";
							}
							else {
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
					sDefn = sDefn + OpenHR.convertLocaleDateToSQL(frmMainForm.txtPValDefault.value) + "	";
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
						sDefn = sDefn + OpenHR.convertLocaleDateToSQL(frmMainForm.cboLookupValueValue.options[frmMainForm.cboLookupValueValue.selectedIndex].Value) + "	";
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
			sDefn = sDefn + $('#ssOleDBGridFilters').getGridParam('selrow') + "\t";	// frmMainForm.ssOleDBGridFilters.Columns("id").Value + "	";
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
				if (frmMainForm.optField_Count.checked == false) {
					sDefn = sDefn +
							" : " + frmMainForm.cboFieldColumn.options[frmMainForm.cboFieldColumn.selectedIndex].text;
				}
			}

			if (frmMainForm.txtPassByType.value == 1) {
				if (tableCollection != null) {
					for (i = 0; i < tableCollection.length; i++) {
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

				sTemp = tree_Nodetext($.jstree._focused().get_selected());
				iIndex = sTemp.indexOf("(");
				if (iIndex >= 0) {
					sTemp = sTemp.substr(0, iIndex - 1);
				}

				sDefn = sDefn + sTemp + "	";
			}
			else {
				var selRowId;
				if (frmMainForm.optType_Calculation.checked == true) {
					selRowId = $('#ssOleDBGridCalculations').getGridParam('selrow');
					sDefn = sDefn + $('#ssOleDBGridCalculations').getCell(selRowId, 'name') + "\t";
				}
				else {
					if (frmMainForm.optType_Operator.checked == true) {
						sTemp = tree_Nodetext($.jstree._focused().get_selected());
						iIndex = sTemp.indexOf("(");
						if (iIndex >= 0) {
							sTemp = sTemp.substr(0, iIndex - 1);
						}
						sDefn = sDefn + sTemp + "	";
					}
					else {
						if (frmMainForm.optType_Filter.checked == true) {
							selRowId = $('#ssOleDBGridFilters').getGridParam('selrow');
							sDefn = sDefn + $('#ssOleDBGridFilters').getCell(selRowId, 'name') + "\t";
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
				for (i = 0; i < colFunctionParameters.length; i++) {
					sKey = colFunctionParameters.item(i).name;
					sKey = sKey.substr(22);
					iIndex = sKey.indexOf("_");
					if (iIndex >= 0) {
						sKey = sKey.substr(0, iIndex);
					}

					if (sKey == $.jstree._focused().get_selected().attr('id')) {
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
		setComponent(sDefn, util_def_exprcomponent_frmUseful.txtAction.value, util_def_exprcomponent_frmUseful.txtLinkRecordID.value, sFunctionParameters);

		$('#optionframe').dialog('close');
	}
}

function component_CancelClick() {
	cancelComponent();
}

function validateComponent() {

	var util_def_exprcomponent_frmUseful = $("#util_def_exprcomponent_frmUseful")[0].children;
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
	sDecimalSeparator = sDecimalSeparator.concat(OpenHR.LocaleDecimalSeparator());
	var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");

	sThousandSeparator = "\\";
	sThousandSeparator = sThousandSeparator.concat(OpenHR.LocaleThousandSeparator());
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
				if (OpenHR.LocaleDecimalSeparator() != ".") {
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
						if (sConvertedValue.indexOf(".") >= 0) {
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
			if (OpenHR.LocaleDecimalSeparator() != ".") {
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
			sValue = OpenHR.convertLocaleDateToSQL(sValue);
			if ((sValue.length == 0) || (sValue == 'null')) {
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
			if (OpenHR.LocaleDecimalSeparator() != ".") {
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
				if (sConvertedValue.indexOf(".") >= 0) {
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
				if (OpenHR.LocaleDecimalSeparator() != ".") {
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
					if (sConvertedValue.indexOf(".") >= 0) {
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

				for (i = 0; i < sValue.length; i++) {
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
					if (OpenHR.LocaleDecimalSeparator() != ".") {
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
					if (!OpenHR.IsValidDate(sValue)) {
						sErrorMsg = "Invalid default date value entered.";
					}
					// Convert the date to SQL format (use this as a validation check).
					// An empty string is returned if the date is invalid.
					sValue = OpenHR.convertLocaleDateToSQL(sValue);
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

function component_saveChanges(psAction, pfPrompt, pfTbOverride) {
	// Expand the work frame and hide the option frame.
	$("#optionframe").attr("data-framesource", "UTIL_DEF_EXPRCOMPONENT");
	$('#optionframe').dialog('close');

	// Pass the component definition back to the expression page.        
	saveChanges(psAction, pfPrompt, pfTbOverride);
}

function ssOleDBGridCalculations_rowcolchange() {

	var sDesc;
	var rowId;
	var rowIndex;
	if (frmMainForm.optType_Filter.checked == true) {
		rowId = $('#ssOleDBGridFilters').getGridParam('selrow');
		rowIndex = $('#ssOleDBGridFilters').jqGrid('getInd', rowId); // counting from 1
		sDesc = "txtFilterDesc_" + rowIndex;
		frmMainForm.txtFilterDescription.value = $('#' + sDesc).val();
		$('#cmdOK').button('enable');
	} else {
		rowId = $('#ssOleDBGridCalculations').getGridParam('selrow');
		rowIndex = $('#ssOleDBGridCalculations').jqGrid('getInd', rowId); // counting from 1
		sDesc = "txtCalcDesc_" + rowIndex;
		frmMainForm.txtCalcDescription.value = $('#' + sDesc).val();
		$('#cmdOK').button('enable');
	}

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

	if (fNoNodeSelected) {
		$('#cmdOK').button('disable');
	} else {
		$('#cmdOK').button('enable');
	}

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

	if (fNoNodeSelected) {
		$('#cmdOK').button('disable');
	} else {
		$('#cmdOK').button('enable');
	}
}

function SSFunctionTree_dblClick() {
	component_OKClick();
}

function addNode(treeID, parentID, position, text, newID, isParent) {
	$('#' + treeID).jstree('create', parentID, position, text, function (data) {
		data[0].id = newID;
	}, true);

	if(isParent) $('#' + newID).addClass('toplevelnode');
}

function safeID(oldID) {
	var returnID = OpenHR.replaceAll(oldID, '/', '');
	returnID = OpenHR.replaceAll(returnID, '+', '');

	return returnID;

}