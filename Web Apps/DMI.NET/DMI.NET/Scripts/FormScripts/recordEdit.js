
//functions that replicate COAIntRecordDMI.ocx

//array to hold changed photos/oles.
window.malngChangedOLEPhotos = [];


function addTabControl(tabNumber) {
	
	var tabID = "FI_21_" + tabNumber;

	if ($("#" + tabID).length <= 0) {
			//The control to be added has a tab number, but the tab doesn't yet exist - create it...
			var tabFontName = $("#txtRecEditFontName").val();
			var tabFontSize = $("#txtRecEditFontSize ").val();
			//var tabCss = "style='font-family: " + tabFontName + " ; font-size: " + tabFontSize + "pt'";
			var tabCss = "";
			var tabs = $("#ctlRecordEdit").tabs();
			var label = getTabCaption(tabNumber),
				li = "<li><a " + tabCss + " href='#" + tabID + "'>" + label + "</a></li>";
			tabs.find(".ui-tabs-nav").append(li);
			tabs.append("<div style='position: relative;' id='" + tabID + "'></div>");
			tabs.tabs("refresh");
			if (tabNumber == 1) tabs.tabs("option", "active", 0);
	}
}


function addControl(tabNumber, controlDef) {

	var tabID = "FI_21_" + tabNumber;

	if (($("#" + tabID).length <= 0) && (tabNumber <= 0)) {
		$("#ctlRecordEdit").append("<div style='position: relative;' id='" + tabID + "'></div>");
		//$("#ctlRecordEdit").css("background-color", "white");
		$("#ctlRecordEdit").css("border", "1px solid gray");
	}


	//add control to tab.
	try {
		$("#" + tabID).append(controlDef);
	}
	catch (e) { alert("unable to add control!"); }

}


function applyLocation(formItem, controlItemArray, bordered) {
	//reduce all sizes by 2 for border.
	var borderwidth;
	borderwidth = (bordered ? 2 : -2);
	formItem.style.top = (Number(controlItemArray[4]) / 15) + "px";
	formItem.style.left = (Number(controlItemArray[5]) / 15) + "px";
	formItem.style.height = (Number((controlItemArray[6]) / 15) - borderwidth) + "px";
	formItem.style.width = (Number((controlItemArray[7]) / 15) - borderwidth) + "px";
	formItem.style.position = "absolute";
}


function applyDefaultValues() {
	//From activeX control recordDMI
	//Populate the screen contrls with the defined column default values.
	//NB calculated default values are handled elsewhere
	//this method also clears all other controls
	
	$('input[id^="txtRecEditControl_"]').each(function (index) {
		var objScreenControl = getScreenControl_Collection($(this).val());

		//NB we don#t use .Tag any more. Removed as we can grab controls from the txtRecordEditControl_ list instead.

		if ((objScreenControl.ColumnID > 0) &&
			(objScreenControl.SelectGranted)) {

			var sDefaultValue = objScreenControl.DefaultValue;
			var lngColumnID = objScreenControl.ColumnID;

			//use updatecontrol???
			if ((sDefaultValue != null) && (sDefaultValue != undefined))
				updateControl(lngColumnID, sDefaultValue);
		}
	});
}

function ClearUniqueColumnControls() {

	$('input[id^="txtRecEditControl_"]').each(function (index) {
		var objScreenControl = getScreenControl_Collection($(this).val());

		//NB we don#t use .Tag any more. Removed as we can grab controls from the txtRecordEditControl_ list instead.

		if ((objScreenControl.ColumnID > 0) &&
			(objScreenControl.TableID = $("#txtRecEditTableID").val()) &&
			(objScreenControl.SelectGranted)) {

			if ((objScreenControl.UniqueCheck) || (!objScreenControl.UpdateGranted)) {

				var sDefaultValue = objScreenControl.DefaultValue;
				var lngColumnID = objScreenControl.ColumnID;

				//use updatecontrol???
				if ((sDefaultValue != null) && (sDefaultValue != undefined))
					updateControl(lngColumnID, sDefaultValue);

			}
		}
	});
}

function CalculatedDefaultColumns() {
	var sColsAndCalcs = "";

	$('input[id^="txtRecEditControl_"]').each(function (index) {
		var objScreenControl = getScreenControl_Collection($(this).val());

		if ((objScreenControl.ColumnID > 0) &&
			(objScreenControl.TableID == $("#txtRecEditTableID").val()) &&
			(objScreenControl.SelectGranted) &&
			(objScreenControl.DfltValueExprID > 0)) {

			sColsAndCalcs += (sColsAndCalcs.length > 0 ? "," : "") + objScreenControl.ColumnID;
		}
	});

	return sColsAndCalcs;
}


function insertUpdateDef() {

	// Adapted from recordDMI.ocx.
	//	Return the SQL string for inserting/updating the current record.
	var fFound = false;
	var fColumnDone = false;
	var sTag = "";
	var iLoop = 0;
	var iLoop2 = 0;
	var iNextIndex = 0;
	var sColumnName = "";
	var sColumnID = "";
	var sInsertUpdateDef = "";
	//	Dim objControl As Control
	var asColumns = new Array();
	var asColumnsToAdd = [0, 0, 0, 0];
	var sTemp = "";
	var fDoControl = false;
	var bCopyImageDataType = false;

	//TODO: Check if there's unsaved changes in the controls.. Can this happen?
	//	If Not UserControl.ActiveControl Is Nothing Then
	//	If Not LostFocusCheck(UserControl.ActiveControl) Then
	//	Exit Function
	//	End If
	//	End If

	//	Dimension the array of columns and values to be updated.
	//	NB. Column 1 = column name in uppercase.
	//	    Column 2 = column value as it needs to appear in the SQL update/insert string.
	//	    Column 3 = column ID (unless it is a relationship column in which case this is the column name eg. ID_1).
	//	    Column 4 = column value formatted for SQL (dates, numbers) but without enclosing within single quotes,
	//	               or doubling up single quotes
	//	ReDim asColumns(4, 0)

	//	' Loop through the screen controls, creating an array of columns and values with
	//	' which we'll construct an insert or update SQL string.

	var uniqueIdentifier = 0;

	$('input[id^="txtRecEditControl_"]').each(function (index) {


		uniqueIdentifier += 1;
		
		var objScreenControl = getScreenControl_Collection($(this).val()); //the properties for the screen element

		//because we can display a column on the screen multiple times, with different properties (e.g. readonly),
		//we need to ensure we're dealing with the right one. So, unique id is used.
		var uniqueID = "FI_" + objScreenControl.ColumnID + "_" + objScreenControl.ControlType + "_" + uniqueIdentifier;
		var objControl = $("#" + uniqueID);	//the actual screen object.

		if ($(objControl).length == 0) {
			fDoControl = false;
		}

		//	sTag = objControl.Tag
		//	' Check if it is a user editable control.
		//	If Len(sTag) > 0 Then
		//	' Check that the control is associated with a column in the current table/view,
		//	' and is updatable.
		//	If (mobjScreenControls.Item(sTag).ColumnID > 0) And _
		//		(Not TypeOf objControl Is COA_Label) And (Not TypeOf objControl Is COA_Navigation) Then

		//only process column controls. (not labels, frames etc...)
		if ((objScreenControl.ColumnID > 0) && ((objScreenControl.ControlType !== 256) && (objScreenControl.ControlType !== Math.pow(2, 14)))) {

			fDoControl = objScreenControl.UpdateGranted;

			if (fDoControl) {
				if ((objScreenControl.ControlType == 64) && (objScreenControl.Multiline)) {	//tdbtextctl.tdbtext
					// NPG20120801 Fault HRPRO-2276
					// Use the original control value for ReadOnly, not the inherited control value.
					// fDoControl = Not objControl.ReadOnly
					// fDoControl = Not mobjScreenControls.Item(sTag).ReadOnly
					// Fault HRPRO-2860 - try again. Neither screen readonly nor db readonly should be included.
					if ((objScreenControl.ReadOnly) || (objScreenControl.ScreenReadOnly)) {
						fDoControl = false;
					}
				}
			}

			if (fDoControl) {
				//	Get the name of the column associated with the current control.
				if (objScreenControl.ControlType == 2048) { //command button.
					sColumnName = "ID_" + objScreenControl.LinkTableID;
					sColumnID = sColumnName;
				}
				else {
					sColumnName = objScreenControl.ColumnName;
					sColumnID = objScreenControl.ColumnID;
				}

				//	Check if the column's update string has already been constructed.
				fColumnDone = false;
				//TODO: test this:
				if (asColumns.length > 0) {
					ubound = asColumns.length - 1;
					for (iNextIndex = 0; iNextIndex <= ubound; iNextIndex++) {
						if (asColumns[iNextIndex][0] == sColumnName) {
							fColumnDone = true;
							break;
						}
					}
				}

				if (!fColumnDone) {
					if (objScreenControl.ControlType == 1024) fDoControl = false;
					if ((objScreenControl.ControlType == 4) || (objScreenControl.ControlType == 8)) { //coa_image or coaint_OLE
						fFound = false;
						ubound = Math.max(0, window.malngChangedOLEPhotos.length - 1);
						for (iLoop2 = 0; iLoop2 <= ubound; iLoop2++) {
							if (window.malngChangedOLEPhotos[iLoop2] == objScreenControl.ColumnID) {
								fFound = true;
								break;
							}
						}
						if (!fFound) {
							fColumnDone = true;
						}
					}
				}

				if (!fColumnDone) {
					//	Add the column name to the array of columns that have already been entered in the
					//	SQL update/insert string.
					//  iNextIndex = UBound(asColumns, 2) + 1
					//	ReDim Preserve asColumns(4, iNextIndex)
					//	asColumns(1, iNextIndex) = sColumnName
					//	asColumns(2, iNextIndex) = ""
					//	asColumns(3, iNextIndex) = sColumnID
					//	asColumns(4, iNextIndex) = ""
					asColumnsToAdd = [sColumnName, "", sColumnID, ""];

					//NB: javascript zero based arrays

					//	Construct the SQL update/insert string for the column.
					if ((objScreenControl.ControlType == 64) && (objScreenControl.Multiline)) {
						// Multi-line character field from a masked textbox (CHAR type column). Save the text from the control.
						$(objControl).val(CaseConversion($(objControl).val(), objScreenControl.ConvertCase));

						if (ConvertData($(objControl).val(), objScreenControl.DataType) == null) {
							asColumnsToAdd[1] = "''";
							asColumnsToAdd[3] = "";
						} else {							
							//asColumnsToAdd[1] = "'" + ConvertData($(objControl).val(), objScreenControl.DataType).replace("'", "''") + "'";
							asColumnsToAdd[1] = "'" + ConvertData($(objControl).val(), objScreenControl.DataType).split("'").join("''") + "'";
							//JPD 20051121 Fault 10583
							//asColumns(4, iNextIndex) = ConvertData(objControl.Text, mobjScreenControls.Item(sTag).DataType)
							//asColumnsToAdd[3] = ConvertData($(objControl).val(), objScreenControl.DataType).replace("\t", " ");
							asColumnsToAdd[3] = ConvertData($(objControl).val(), objScreenControl.DataType).split("\t").join(" ");
						}
					}

					else if (objScreenControl.ControlType == 4) {
						//COA_Image
						//TODO: somehow.....
						if (objScreenControl.OLEType < 2) {
							//	sTemp = objControl.ASRDataField
							//	asColumns(2, iNextIndex) = "'" & Replace(sTemp, "'", "''") & "'"
							//	asColumns(4, iNextIndex) = CStr(objControl.OLEType) & sTemp
						} else {
							//	asColumns(2, iNextIndex) = asColumns(1, iNextIndex)
							//	asColumns(4, iNextIndex) = CStr(objControl.OLEType) & asColumns(1, iNextIndex)
							//	bCopyImageDataType = True
						}
					}

					else if ((objScreenControl.ControlType == 64) && (objScreenControl.Multiline == false) && (objScreenControl.Mask.length > 0)) {



						//TDBMask6Ctl.TDBMask Then
						//	Character field from a masked textbox (CHAR type column). Save the text from the control.
						if ($(objControl).val() == 0) {
							asColumnsToAdd[1] = "null";
							asColumnsToAdd[3] = "null";
						} else {
							$(objControl).val(CaseConversion($(objControl).val(), objScreenControl.ConvertCase));
							asColumnsToAdd[1] = "'" + $(objControl).val().split("'").join("''") + "'";
							//	JPD 20051121 Fault 10583
							//	'asColumns(4, iNextIndex) = objControl.Text
							asColumnsToAdd[3] = $(objControl).val().split("\t").join(" ");
						}
					}

					else if (((objScreenControl.ControlType == 64) &&
						(!objScreenControl.Multiline)) &&
						((objScreenControl.DataType !== 2) &&
						(objScreenControl.DataType !== 4) &&
						(objScreenControl.DataType !== 11))
					) {
						//TextBox Then
						//	Character field from an unmasked textbox (CHAR type column). Save the text from the control.
						$(objControl).val(CaseConversion($(objControl).val(), objScreenControl.ConvertCase));
						asColumnsToAdd[1] = "'" + $(objControl).val().split("'").join("''") + "'";
						//	JPD 20051121 Fault 10583
						//	asColumns(4, iNextIndex) = objControl.Text
						asColumnsToAdd[3] = $(objControl).val().split("\t").join(" ");
					}

					else if ((objScreenControl.ControlType == 64) &&
						((objScreenControl.DataType == 2) || (objScreenControl.DataType == 4))) {						
						//TDBNumber6Ctl.TDBNumber Then
						//	Integer or Numeric field from a numeric textbox (INT or NUM type column). Save the value from the control.
						if (ConvertData($(objControl).val(), objScreenControl.DataType) == null) {
							asColumnsToAdd[1] = "null";
							asColumnsToAdd[3] = "null";
						} else {
							asColumnsToAdd[1] = ConvertData($(objControl).val(), objScreenControl.DataType);
							asColumnsToAdd[1] = ConvertNumberForSQL(asColumnsToAdd[1]);
							asColumnsToAdd[3] = asColumnsToAdd[1];
						}
					}

					else if (objScreenControl.ControlType == 1) {
						//CheckBox Then
						//	 Logic field (BIT type column). Save 1 for true, 0 for False.
						asColumnsToAdd[1] = $(objControl).is(":checked") ? "1" : "0";
						asColumnsToAdd[3] = $(objControl).is(":checked") ? "1" : "0";
					}

					else if (objScreenControl.ControlType == 2) { //&& (objScreenControl.ColumnType == 1))
						// objControl Is COAInt_Lookup Then
						//	Lookup field from a combo (unknown type column). Get the column type and save the appropraite value from the combo.
						switch (objScreenControl.DataType) {


							case 12:
							case -1:
								asColumnsToAdd[1] = "'" + $(objControl).val().split("'").join("''") + "'";
								//	JPD 20051121 Fault 10583
								//	asColumns(4, iNextIndex) = objControl.Text
								asColumnsToAdd[3] = $(objControl).val().split("\t").join(" ");
								break;
							case 4:
							case 2:
								if ($(objControl).val().length > 0) {								
									asColumnsToAdd[1] = ConvertData($(objControl).val(), objScreenControl.DataType);
									asColumnsToAdd[1] = ConvertNumberForSQL(asColumnsToAdd[1]);
									asColumnsToAdd[3] = asColumnsToAdd[1];

								} else {
									asColumnsToAdd[1] = "null";
									asColumnsToAdd[3] = "null";
								}
								break;
							case 11:
								if ($(objControl).val().length > 0) {
									asColumnsToAdd[1] = "'" + OpenHR.convertLocaleDateToSQL($(objControl).val()) + "'";
									asColumnsToAdd[4] = OpenHR.convertLocaleDateToSQL($(objControl).val()).split("\t").join(" ");
								} else {
									asColumnsToAdd[1] = "null";
									asColumnsToAdd[3] = "null";
								}
								break;
							default:
								asColumnsToAdd[0] = "";
						}
					}

					else if (objScreenControl.ControlType == 16) {
						//TypeOf objControl Is COAInt_OptionGroup Then
						//	 Character field from an option group (CHAR type column). Save the text from the option group.
						var optionSelected = $("input:radio[name='" + $(objControl).attr("id") + "']:checked").val();

						if (optionSelected == undefined) {
							asColumnsToAdd[1] = "";
							asColumnsToAdd[3] = "";
						} else {
							asColumnsToAdd[1] = "'" + optionSelected.split("'").join("''") + "'";
							//asColumnsToAdd[1] = "'" + $(objControl).val().replace("'", "''") + "'";
							//	'JPD 20051121 Fault 10583
							//	'asColumns(4, iNextIndex) = objControl.Text
							asColumnsToAdd[3] = optionSelected.split("\t").join(" ");
						}
					}

					else if (objScreenControl.ControlType == 8) {
						//TypeOf objControl Is COAInt_OLE Then
						// OLE field (CHAR type column). Save the name of the OLE file.
						
						if (($(objControl).attr('data-filename').length > 0) && (objScreenControl.OLEType < 2)) {
							asColumnsToAdd[1] = "'" + objControl.attr('data-filename') + "'";
							asColumnsToAdd[3] = objScreenControl.OLEType + objControl.attr('data-filename');
						}
						else if (objScreenControl.OLEType < 2) {
							asColumnsToAdd[1] = "''";
							asColumnsToAdd[3] = objScreenControl.OLEType;
						} else {
							asColumnsToAdd[1] = asColumnsToAdd[0];
							asColumnsToAdd[3] = objScreenControl.OLEType + asColumnsToAdd[0];
							bCopyImageDataType = true;
						}					
					}

					else if (objScreenControl.ControlType == 32) {
						//TypeOf objControl Is COA_Spinner Then
						//	' Integer field from an spinner (INT type column). Save the value from the spinner.
						asColumnsToAdd[1] = $.trim($(objControl).val());
						asColumnsToAdd[3] = $.trim($(objControl).val());
					}
						//	'JPD 20040714 Fault 8333
						//	'ElseIf TypeOf objControl Is GTMaskDate.GTMaskDate Then
					else if ((objScreenControl.ControlType == 64) && (objScreenControl.DataType == 11)) {
						////TypeOf objControl Is TDBDate6Ctl.TDBDate Then
						////	' Date field from a date control (DATETIME type column). Save the value from the control formatted as 'mm/dd/yyyy' for SQL.
						//if (ConvertData($(objControl).val(), objScreenControl.DataType) == null) {
						//	asColumnsToAdd[1] = "null";
						//	asColumnsToAdd[3] = "null";
						//} else {

							asColumnsToAdd[1] = "'" + OpenHR.convertLocaleDateToSQL($(objControl).val()) + "'";
							asColumnsToAdd[3] = OpenHR.convertLocaleDateToSQL($(objControl).val()).split("\t").join(" ");
			//			}
					}

					else if (objScreenControl.ControlType == 4096) { //Working Pattern Field (CHAR type column, len 14).
						fDoControl = true;
						
						//At this point objControl contains the fieldset that contains the actual working pattern checkboxes; so we need
						//to iterate over them to get the value that needs to be passed to SQL
						var workingPatternTemplate = "SSMMTTWWTTFFSS";
						var workingPatternString = "";

						$(objControl).children("input").each(function (itemIndex) {
							if ($(this).is(":checked") == true) {
								workingPatternString += workingPatternTemplate[itemIndex];
							} else {
								workingPatternString += " ";
							}
						});
						
						asColumnsToAdd[1] = workingPatternString;
						asColumnsToAdd[3] = workingPatternString;
					}

					else if (objScreenControl.ControlType == Math.pow(2, 15)) { // 32768 - ctlColourPicker
						fDoControl = true;
						asColumnsToAdd[1] = objControl.val();
						asColumnsToAdd[3] = objControl.val();
					}

					else if (objScreenControl.ControlType == 2048) {						
						//TypeOf objControl Is CommandButton (link button)
						if (objScreenControl.LinkTableID != $("#txtCurrentParentTableID").val()) {
							ubound = (window.mavIDColumns.length);

							for (iLoop = 0; iLoop < (ubound); iLoop++) {

								if (window.mavIDColumns[iLoop][1] == "ID_" + objScreenControl.LinkTableID) {
									asColumnsToAdd[1] = window.mavIDColumns[iLoop][2];
									asColumnsToAdd[3] = window.mavIDColumns[iLoop][2];
									break;
								}

							}
						} else {
							asColumnsToAdd[1] = $("#txtCurrentParentTableID").val();
							asColumnsToAdd[3] = $("#txtCurrentParentTableID").val();
						}		
					}
				}
			}
			//	End If

			if ((fDoControl) && (!fColumnDone)) asColumns.push(asColumnsToAdd);

		}


	});

	//	Next objControl
	//	Set objControl = Nothing

	//	See if we are a history screen and if we are save away the id of the parent also
	if ($("#txtCurrentParentTableID").val() > 0) {
		//	Check if the column's update string has already been constructed.
		fColumnDone = false;
		var ubound = asColumns.length - 1;
		for (iNextIndex = 0; iNextIndex <= ubound; iNextIndex++) {
			if (asColumns[iNextIndex][0] == "ID_" + $.trim($("#txtCurrentParentTableID").val())) {
				fColumnDone = true;
				break;
			}
		}

		if (!fColumnDone) {
			
			var asIDToAdd = [0, 0, 0, 0];
			//	Add the column name to the array of columns that have already been entered in the
			//	SQL update/insert string.
			//iNextIndex = asColumns.length + 1; //TODO: check this...			
			asIDToAdd[0] = "ID_" + $.trim($("#txtCurrentParentTableID").val());
			asIDToAdd[1] = $.trim($("#txtCurrentParentRecordID").val());
			asIDToAdd[2] = "ID_" + $.trim($("#txtCurrentParentTableID").val());
			asIDToAdd[3] = $.trim($("#txtCurrentParentRecordID").val());

			asColumns.push(asIDToAdd);

		}
	}

	if (asColumns.length > 0) {
		//	Create a SQL string to update the record with.
		if ($("#txtCurrentRecordID").val() == 0) {
			sInsertUpdateDef = $("#txtRecEditRealSource").val() + "\t";

			if (bCopyImageDataType) {
				sInsertUpdateDef += "1" + "\t";
			} else {
				sInsertUpdateDef += "0" + "\t";
			}

			sInsertUpdateDef += $("#txtCopiedRecordID").val() + "\t";
			ubound = asColumns.length - 1;
			for (iLoop = 0; iLoop <= ubound; iLoop++) {
				if (asColumns[iLoop][0].length > 0) {
					sInsertUpdateDef += asColumns[iLoop][2] + "\t" + asColumns[iLoop][3] + "\t";
				}
			}
		} else {
			//	Construct the SQL update string from the array of columns and values.
			ubound = asColumns.length - 1;
			for (iLoop = 0; iLoop <= ubound; iLoop++) {
				if (asColumns[iLoop][0].length > 0) {
					sInsertUpdateDef += asColumns[iLoop][2] + "\t" + asColumns[iLoop][3] + "\t";
				}
			}
		}
	}

	return sInsertUpdateDef;

}

function ConvertSQLNumberToLocale(strInput) {
	return OpenHR.replaceAll(String(strInput), ".", window.LocaleDecimalSeparator);
}

function ConvertNumberForSQL(strInput) {
	// Get a number in the correct format for a SQL string
	// (e.g. on french systems replace decimal comma for a decimal point)
	// TODO: return strInput.replace(msLocaleDecimalSeparator, ".");

	return OpenHR.replaceAll(String(strInput),window.LocaleDecimalSeparator, ".");
}


function ConvertData(pvData, pDataType) {
	//TODO: Test this function!!
	// Convert the given variant value into the given data type.
	var vReturnData;

	if (pvData == null) {
		vReturnData = null;
	} else {
		switch (pDataType) {
			case -7:
				//sqlBoolean
				vReturnData = Boolean(pvData);
				break;
			case -1:
			case 12:
				//sqlVarChar, sqlLongVarChar
				vReturnData = (pvData).split(/~+$/).join(''); //rtrim function
				if (vReturnData.length == 0) {
					vReturnData = null;
				}
				break;
			case 11:
				//sqlDate				
				if (isValidDate(new Date(pvData))) {
					vReturnData = Date.parse(pvData);
				} else {
					vReturnData = null;
				}
				break;
			case 2:
				//sqlNumeric
				if ($.trim(pvData).length == 0) {
					vReturnData = null;
				} else {
					vReturnData = Number(pvData.toString().split(window.LocaleThousandSeparator).join("").split(window.LocaleDecimalSeparator).join("."));
				}
				break;
			case 4:
				// sqlInteger
				if ($.trim(pvData).length == 0) {

					vReturnData = null;
				} else {
					vReturnData = Number(pvData.toString().split(window.LocaleThousandSeparator).join("").split(window.LocaleDecimalSeparator).join("."));
				}
				break;

			default:
				vReturnData = pvData;
		}
	}

	return vReturnData;

}

function isValidDate(d) {
	if (Object.prototype.toString.call(d) !== "[object Date]")
		return false;
	return !isNaN(d.getTime());
}

function CaseConversion(psText, piCaseConversion) {
	// Perform the required case conversion on the given text.
	var lngPos;
	var sLastCharacter = "";

	// Do nothing if the given text is empty.
	if ($.trim(psText).length > 0) {
		// Do nothing if the given text is numeric.
		if (isNaN(psText)) {
			switch (piCaseConversion) {
				case 0:
					//No conversion
					break;
				case 1:
					//Upper case
					psText = psText.toUpperCase();
					break;
				case 2:
					//Lower case
					psText = psText.toLowerCase();
					break;
				case 3:
					//Proper conversion
					//First LCase everything !
					psText = $.trim(psText).toLowerCase();

					// Then Ucase first letter
					psText = psText.charAt(0).toUpperCase() + psText.slice(1); //UCase(Left(psText, 1)) & Right(psText, Len(psText) - 1)
					for (lngPos = 1; lngPos <= psText.length; lngPos++) {
						sLastCharacter = psText.substr(lngPos - 1, 1);
						if (((sLastCharacter < "A") || (sLastCharacter > "Z")) &&
							((sLastCharacter < "a") || (sLastCharacter > "z")) &&
							((sLastCharacter < "À") || (sLastCharacter > "Ö")) &&
							((sLastCharacter < "Ù") || (sLastCharacter > "Ý")) &&
							((sLastCharacter < "ß") || (sLastCharacter > "ö")) &&
							((sLastCharacter < "ù") || (sLastCharacter > "ÿ"))) {
							psText = Left(psText, lngPos - 1) + psText.substr(lngPos - 1, 1).toUpperCase() + Right(psText, psText.length - lngPos);
							//psText = psText.substr(0, lngPos - 1) + psText.substr(lngPos, 1).toUpperCase() + Right(psText, psText.length - lngPos);
						} else if (lngPos > 2) {
							// Catch the McName.
							if ((psText.substr(lngPos - 2, 1) == "M") && (sLastCharacter = "c")) {
								psText = Left(psText, lngPos - 1) & psText.substr(lngPos - 1, 1).toUpperCase() & Right(psText, psText.length - lngPos);
								//psText = psText.substr(0, lngPos - 1) + psText.substr(lngPos, 1).toUpperCase() + Right(psText, psText.length - lngPos);
							}
						}
					}
			}
		}
	}

	return psText;

}

function Left(str, n) {
	if (n <= 0)
		return "";
	else if (n > String(str).length)
		return str;
	else
		return String(str).substring(0, n);
}

function Right(str, n) {
	if (n <= 0)
		return "";
	else if (n > String(str).length)
		return str;
	else {
		var iLen = String(str).length;
		return String(str).substring(iLen, iLen - n);
	}
}

function allDefaults() {
	//TODO:
	//this function checks if all the user updatable controls have defaults, this is so the save button can be enabled if the all controls have defaults.
	return false;
}

function validateSave() {
	// Validate the recordd before saving.
	var fValid = true;
	var fLinked = false;
	var fHasParents = false;
	var iLoop;

	//converted from activeX function validateSave()
	// Check that at least one parent is linked to.
	if (Number($("#txtCurrentParentRecordID").val()) == 0) {
		fHasParents = false;
		fLinked = false;

		var ubound = (this.mavIDColumns.length);

		for (iLoop = 0; iLoop < (ubound) ; iLoop++) {

			if (this.mavIDColumns[iLoop][2].length > 2) {
				// Must be parent id column (ID_) rather that own record id column (id)
				fHasParents = true;

				if (Number(this.mavIDColumns[iLoop][3]) > 0) {
					fLinked = true;
					break;
				}
			}
		}

		if ((!fLinked) && (fHasParents)) {
			fValid = false;
			OpenHR.messageBox("Unable to save record, a link must be made with the parent table.", vbExclamation + vbOKOnly, "OpenHR");
		}
	}

	return fValid;
}


function getScreenControl_Collection(screenControlValue) {

	var controlItemArray = screenControlValue.split("\t");

	var defaultValue = function () {
		//if this is a decimal number, convert it to locale
		if (Number(controlItemArray[26]) > 0) {
			return ConvertSQLNumberToLocale(controlItemArray[24]);
		} else {
			return controlItemArray[24];
		}
	};

	var screenControls = {
		PageNo: Number(controlItemArray[0]),
		TableID: Number(controlItemArray[1]),
		ColumnID: Number(controlItemArray[2]),
		ControlType: Number(controlItemArray[3]),
		TopCoord: Number(controlItemArray[4]),
		LeftCoord: Number(controlItemArray[5]),
		Height: Number(controlItemArray[6]),
		Width: Number(controlItemArray[7]),
		Caption: controlItemArray[8],
		BackColor: Number(controlItemArray[9]),
		ForeColor: Number(controlItemArray[10]),
		FontName: controlItemArray[11],
		FontSize: Number(controlItemArray[12]),
		FontBold: (Number(controlItemArray[13]) != 0),
		FontItalic: (Number(controlItemArray[14]) != 0),
		FontStrikethru: (Number(controlItemArray[15]) != 0),
		FontUnderline: (Number(controlItemArray[16]) != 0),
		DisplayType: Number(controlItemArray[17]),
		TabIndex: Number(controlItemArray[18]),
		BorderStyle: Number(controlItemArray[19]),
		Alignment: Number(controlItemArray[20]),
		ColumnName: controlItemArray[21],
		ColumnType: Number(controlItemArray[22]),
		DataType: Number(controlItemArray[23]),
		DefaultValue: defaultValue(),
		Size: Number(controlItemArray[25]),
		Decimals: Number(controlItemArray[26]),
		LookupTableID: Number(controlItemArray[27]),
		LookupColumnID: Number(controlItemArray[28]),
		SpinnerMinimum: Number(controlItemArray[29]),
		SpinnerMaximum: Number(controlItemArray[30]),
		SpinnerIncrement: Number(controlItemArray[31]),
		Mandatory: (Number(controlItemArray[32]) != 0),
		UniqueCheck: (Number(controlItemArray[33]) != 0),
		ConvertCase: Number(controlItemArray[34]),
		Mask: controlItemArray[35],
		BlankIfZero: (Number(controlItemArray[36]) != 0),
		Multiline: (Number(controlItemArray[37]) != 0),
		ColumnAlignment: Number(controlItemArray[38]),
		DfltValueExprID: Number(controlItemArray[39]),
		ReadOnly: (Number(controlItemArray[40]) != 0),
		StatusBarMessage: controlItemArray[41],
		LinkTableID: Number(controlItemArray[42]),
		LinkOrderID: Number(controlItemArray[43]),
		LinkViewID: Number(controlItemArray[44]),
		AFDEnabled: (Number(controlItemArray[45]) != 0),
		TableName: controlItemArray[46],
		SelectGranted: (Number(controlItemArray[47]) != 0),
		UpdateGranted: (Number(controlItemArray[48]) != 0),
		OLEOnServer: (Number(controlItemArray[49]) != 0),
		PictureID: Number(controlItemArray[50]),
		TrimmingType: Number(controlItemArray[51]),
		Use1000Separator: (Number(controlItemArray[52]) != 0),
		LookupFilterColumnID: Number(controlItemArray[53]),
		LookupFilterValueID: Number(controlItemArray[54]),
		OLEType: Number(controlItemArray[55]),
		EmbeddedEnabled: (Number(controlItemArray[56]) != 0),
		MaxOleSize: Number(controlItemArray[57]),
		NavigateTo: formatAddress(controlItemArray[58]),
		NavigateIn: controlItemArray[59],
		NavigateOnSave: controlItemArray[60],
		ScreenReadOnly: (Number(controlItemArray[61]) != 0)
	};


	return screenControls;

}



// -------------------------------------------------- Add the record Edit controls ------------------------------------------------
function AddHtmlControl(controlItem, txtcontrolID, key) {
	try {
		var controlItemArray = controlItem.split("\t");
	} catch (e) {
		return false;
	}

	var iPageNo = 0;
	var controlID = "";
	var tmpNum = txtcontrolID.indexOf("_");
	txtcontrolID = txtcontrolID.substring(tmpNum);

	if (controlItemArray[0] < 0) {
		//The definition is for an id column.            
		this.mavIDColumns.push([Number(controlItemArray[1]), controlItemArray[2], 0]);
	}

	//-------------------------------------------------Get permissions for this control first -----------------------------------------------------------------
	var controlType = Number(controlItemArray[3]);
	var fSelectOK = false;
	var fParentTableControl = false;
	var fControlEnabled = true;
	var fReadOnly = false;
	
	//Permissions. From activeX recordDMI.formatscreen function.
	if ($("#txtRecEditTableID").val() == controlItemArray[1]) {
		if (controlItemArray[2] > 0) { }		
		fSelectOK = (Number(controlItemArray[47]) != 0);
		fParentTableControl = false;

		// Disable control if no permission is granted.
		fControlEnabled = (Number(controlItemArray[40]) == 0);  //database readonly value
		if ((controlType == 8) || ((controlType == 64) && (Number(controlItemArray[37]) != 0))) fControlEnabled = true;    //enable all multiline text, or OLEs


		if (fControlEnabled) {
			if ((controlType == 64) && (Number(controlItemArray[23]) == 11)) {
				//Date Control
				fControlEnabled = (Number(controlItemArray[48]) !== 0);  // UpdateGranted property
			}
			else if (controlType == 2048) {
				//CommandButton
				fControlEnabled = false;
			}
			else {
				fControlEnabled = (Number(controlItemArray[48]) !== 0);  // UpdateGranted property

				if ((controlType == 64) && (Number(controlItemArray[37]) != 0) && ((Number(controlItemArray[23]) == 12) || (Number(controlItemArray[23]) == -1))) {
					//if multiline text and (sqlVarchar or sqllongvarchar)
					if ((!fControlEnabled) || (Number(controlItemArray[61]) != 0)) {
						//if screen.readonly or disabled
						fControlEnabled = true;
						fReadOnly = true;
					}
				}

			}
		}
	}
	else {
		//Parent table control.
		fParentTableControl = true;
		if ((controlType == 256) || (controlType == 512) || (controlType == 4) || (controlType == 2 ^ 13) || (controlType == 2 ^ 14) || (controlType == 2 ^ 15)) {
			//label, frame, image, line, navigation or colourpicker
			fControlEnabled = false;
		}

		if ((controlType == 64) && (Number(controlItemArray[37]) != 0) && ((Number(controlItemArray[23]) == 12) || (Number(controlItemArray[23]) == -1))) {
			//if multiline text and (sqlVarchar or sqllongvarchar)
			if ((!fControlEnabled) || (Number(controlItemArray[61]) != 0)) {
				//if screen.readonly or disabled
				fControlEnabled = true;
				fReadOnly = true;
			}
		}
	}

	if (Number(controlItemArray[61]) != 0) {
		//Screen.Readonly
		fControlEnabled = false;
	}


	//----------------------------------------------------------------------- Now add the control to the form... ---------------------------------------------------------------

	iPageNo = Number(controlItemArray[0]);
	controlID = "FI_" + controlItemArray[2] + "_" + controlItemArray[3] + "" + txtcontrolID;
	var columnID = controlItemArray[2];
	var tabIndex = Number(controlItemArray[18]);

	//TODO: move styling to classes?    
	//TODO: move duplicated property setting blocks to separate functions
	
	var span;
	var top;
	var left;
	var height;
	var width;
	var borderCss;
	var radioTop;
	var button;
	var image;
	switch (Number(controlItemArray[3])) {
		case 1: //checkbox
			span = document.createElement('span');

			applyLocation(span, controlItemArray, false);
			//span.style.margin = "0px";
			//span.style.textAlign = "left";
			span.style.display = "block";
			span.style.overflow = 'hidden';
			var checkbox = span.appendChild(document.createElement('input'));
			checkbox.type = "checkbox";
			checkbox.id = controlID;
			checkbox.style.fontFamily = controlItemArray[11];
			checkbox.style.fontSize = controlItemArray[12] + 'pt';
			//checkbox.style.position = "absolute";
			//checkbox.style.top = "50%";

			//checkbox.style.padding = "0px";
			//checkbox.style.margin = "-7px 0px 0px 0px";
			checkbox.style.textAlign = "left";
			checkbox.style.verticalAlign = 'middle';
			checkbox.style.borderStyle = 'none';
			var label = span.appendChild(document.createElement('label'));
			label.htmlFor = checkbox.id;
			label.appendChild(document.createTextNode(controlItemArray[8]));

			label.style.fontFamily = controlItemArray[11];
			label.style.fontSize = controlItemArray[12] + 'pt';

			//Check if control should be disabled (read only or screen read only)
			if (controlItemArray[40] != "0" || controlItemArray[61] != "0") {
				//checkbox.setAttribute("disabled", "disabled");
				//$(checkbox).addClass("ui-state-disabled");
				$(checkbox).prop('disabled', true);
			}

			//align left or right...
			if (controlItemArray[20] != "0") {
				//right align				
				checkbox.style.float = 'right';
				checkbox.style.paddingTop = '5px';
				//span.className = "checkbox right";
				//checkbox.style.right = "0px";
			} else {
				//left align
				label.style.paddingLeft = '8px';
				//span.className = "checkbox left";
				//checkbox.style.left = "0px";
				//label.style.marginLeft = "18px";
			}

			label.style.verticalAlign = 'middle';
			
			if (tabIndex > 0) checkbox.tabindex = tabIndex;

			checkbox.setAttribute("data-columnID", columnID);
			checkbox.setAttribute('data-controlType', controlItemArray[3]);
			checkbox.setAttribute("data-control-tag", key);
			checkbox.setAttribute("data-Mandatory", controlItemArray[32]);

			$(checkbox).attr("title", controlItemArray[41]);
			$(label).attr("title", controlItemArray[41]);

			if (!fControlEnabled) {
				$(span).prop('disabled', true);
				$(checkbox).prop('disabled', true);
				$(label).prop('disabled', true);
			}

			//Add control to relevant tab, create if required.                
			addControl(iPageNo, span);

			break;
		case 2: //ctlCombo


			var selector = document.createElement('select');
			selector.id = controlID;
			applyLocation(selector, controlItemArray, true);
			//override default height calc:
			selector.style.height = (Number(controlItemArray[6]) / 15) + "px";
			//selector.style.backgroundColor = "White";
			//selector.style.color = "Black";
			selector.style.fontFamily = controlItemArray[11];
			selector.style.fontSize = controlItemArray[12] + 'pt';
			selector.style.borderWidth = "1px";
			selector.setAttribute("data-columnID", columnID);
			selector.setAttribute('data-controlType', controlItemArray[3]);
			selector.setAttribute("data-control-key", key);
			selector.setAttribute("data-Mandatory", controlItemArray[32]);

			if (controlItemArray[22] == 1) {
				//column type = ---- LOOKUPS ----
				selector.setAttribute("data-columntype", "lookup");
				//plngColumnID, plngLookupColumnID, psLookupValue, pfMandatory, pstrFilterValue
				selector.setAttribute("data-LookupTableID", controlItemArray[27]);
				selector.setAttribute("data-LookupColumnID", controlItemArray[28]);
				selector.setAttribute("data-LookupFilterColumnID", controlItemArray[53]);
				selector.setAttribute("data-LookupFilterValueID", controlItemArray[54]);				
			}

			if (!fControlEnabled) $(selector).prop('disabled', true);

			if (tabIndex > 0) selector.tabindex = tabIndex;

			$(selector).attr("title", controlItemArray[41]);

			addControl(iPageNo, selector);

			if (controlItemArray[22] == 0) {
				//Add empty option for dropdown lists
				var option = document.createElement('option');
				option.value = '';
				option.appendChild(document.createTextNode(''));
				selector.appendChild(option);
			}

			break;

		case 4: //Image - NOT PHOTO!!!
			image = document.createElement('img');
			image.id = controlID;
			applyLocation(image, controlItemArray, true);
			image.style.border = "1px solid gray";
			image.style.padding = "0px";
			image.setAttribute("data-columnID", columnID);
			image.setAttribute('data-controlType', controlItemArray[3]);
			image.setAttribute("data-control-key", key);
			image.setAttribute("data-Mandatory", controlItemArray[32]);

			if (!fControlEnabled) $(image).prop('disabled', true);

			var path = window.ROOT + 'Home/ShowImageFromDb?imageID=' + controlItemArray[50];

			image.setAttribute('src', path);

			//Add control to relevant tab, create if required.                
			addControl(iPageNo, image);

			break;
		case 8: //ctlOle

			button = document.createElement('input');
			button.type = "button";
			button.id = controlID;			
			applyLocation(button, controlItemArray, true);
			button.style.padding = "0px";
			button.setAttribute("data-columnID", columnID);
			button.setAttribute('data-controlType', controlItemArray[3]);
			button.setAttribute("data-control-key", key);
			button.setAttribute('data-OleType', controlItemArray[55]); // == 2 ? 3 : controlItemArray[55]);			
			button.setAttribute('data-maxEmbedSize', controlItemArray[57]);
			button.setAttribute('data-readOnly', (!(fControlEnabled)));
			button.style.overflow = 'hidden';
			button.style.fontWeight = 'normal';
			button.style.fontSize = '10px';
			
			if (tabIndex > 0) button.tabindex = tabIndex;

			//Check if control should be disabled (read only or screen read only)
			if (controlItemArray[40] != "0" || controlItemArray[61] != "0") {
				button.setAttribute("disabled", "disabled");
				$(button).addClass("ui-state-disabled");
			}

			$(button).attr("title", controlItemArray[41]);

			addControl(iPageNo, button);

			break;
		case 16: //ctlRadio
			//TODO: set 'maxlength=.size' if fselectOK is true and not fparentcontrol
			//TODO: .disabled = (!fControlEnabled);
			top = (Number(controlItemArray[4]) / 15);
			left = (Number(controlItemArray[5]) / 15);
			height = (Number((controlItemArray[6]) / 15) - 2);
			width = (Number((controlItemArray[7]) / 15) - 2);

			var cssBorderStyle = new Object();

			if (controlItemArray[19] == "0") {
				//pictureborder?
				cssBorderStyle.width = "0px";
				cssBorderStyle.style = "none";
				cssBorderStyle.color = "transparent";
				radioTop = 2;
				width += 2;
			} else {
				cssBorderStyle.width = "1px";
				cssBorderStyle.style = "solid";
				cssBorderStyle.color = "#999";
				width -= 2;
				height -= 2;

				//TODO ??  fontadjustment?

				radioTop = 19 + Number((controlItemArray[12] - 8) * 1.375);

				//TODO - android browser/tablet adjustment
			}

			fieldset = document.createElement("fieldset");
			fieldset.style.position = "absolute";
			fieldset.style.top = top + "px";
			fieldset.style.left = left + "px";
			fieldset.style.width = width + "px";
			fieldset.style.height = height + "px";
			fieldset.style.padding = "0px";
			fieldset.style.borderWidth = cssBorderStyle.width;
			fieldset.style.borderStyle = cssBorderStyle.style;
			fieldset.style.borderColor = cssBorderStyle.color;
			//appply font at fieldset level; it cascades.
			fieldset.style.fontFamily = controlItemArray[11];
			fieldset.style.fontSize = controlItemArray[12] + 'pt';
			fieldset.style.fontWeight = (Number(controlItemArray[13]) != 0) ? "bold" : "normal";
			fieldset.style.whiteSpace = 'nowrap';
			
			fieldset.id = controlID;
			fieldset.setAttribute("data-datatype", "Option Group");
			fieldset.setAttribute("data-columnID", columnID);
			fieldset.setAttribute('data-controlType', controlItemArray[3]);
			fieldset.setAttribute("data-alignment", controlItemArray[20]);
			fieldset.setAttribute("data-control-key", key);
			fieldset.setAttribute("data-Mandatory", controlItemArray[32]);

			if ((controlItemArray[19] != "0") && (controlItemArray[8].length > 0)) {
				//has a border and a caption
				legend = fieldset.appendChild(document.createElement('legend'));
				legend.style.fontFamily = controlItemArray[11];
				legend.style.fontSize = controlItemArray[12] + 'pt';
				legend.style.fontWeight = (Number(controlItemArray[13]) != 0) ? "bold" : "normal";
				legend.appendChild(document.createTextNode(controlItemArray[8].split('&&').join('&')));
			}

			$(fieldset).attr("title", controlItemArray[41]);
			
			if (!fControlEnabled) $(fieldset).prop('disabled', true);
			//No Option Group buttons - these are added as values next.

			addControl(iPageNo, fieldset);


			break;
		case 32: //ctlSpinner		
			if (window.isMobileDevice == "True") {				
				spinner = document.createElement('input');			
				spinner.className = "number";
				
				applyLocation(spinner, controlItemArray, true);
				spinner.style.padding = "0 0 0 0";
				spinner.id = controlID;

				spinner.style.fontFamily = controlItemArray[11];
				spinner.style.fontSize = controlItemArray[12] + 'pt';
				spinner.style.width = (Number((controlItemArray[7]) / 15)) + "px";
				spinner.style.margin = "0px";
				spinner.setAttribute("data-columnID", columnID);
				spinner.setAttribute('data-controlType', controlItemArray[3]);
				spinner.setAttribute("data-control-key", key);
				spinner.setAttribute("data-Mandatory", controlItemArray[32]);

				//Add some attributes used by the autoNumeric plugin we are using to validate numeric text boxes
				var x; //For the loop below
				var value = "";

				spinner.setAttribute("data-a-sep", ''); //No thousand separator
				spinner.setAttribute('data-m-dec', '0'); //Decimal places

				//Size of field includes decimals but not the decimal point; For example if Size=6 and Decimals=2 the maximum value to be allowed is 9999.99
				spinner.setAttribute('data-v-min', '-2147483647'); //This is -Int32.MaxValue
				spinner.setAttribute('data-v-max', '2147483647'); //This is Int32.MaxValue
				
				//Alignment; this is not used by the plugin so we'll add it as a CSS style
				if (controlItemArray[38] == "0") {
					$(textbox).css('text-align', 'left');
				} else if (controlItemArray[38] == "1") {
					$(textbox).css('text-align', 'right');
				} else {
					$(textbox).css('text-align', 'center');
				}

				//Blank if zero; set the value of the attribute data-blankIfZeroValue for this textbox depending on its blankIfZero setting and its decimal places
				if (controlItemArray[36] == "1") { //Blank if zero
					spinner.setAttribute("data-blankIfZeroValue", "");
				} else {
					var decimalPlaces = Number(controlItemArray[26]);
					if (decimalPlaces == 0) {
						spinner.setAttribute("data-blankIfZeroValue", "0");
					} else {
						value = "0.";
						for (x = decimalPlaces; x--;) value += "0";
						spinner.setAttribute("data-blankIfZeroValue", value);
					}
				}

				if (tabIndex > 0) spinner.tabindex = tabIndex;
				//if (!fControlEnabled) spinnerContainer.disabled = true;
				if (!fControlEnabled) spinner.setAttribute('data-disabled', 'true');

				$(spinner).attr("title", controlItemArray[41]);

				//Add control to relevant tab, create if required.                
				addControl(iPageNo, spinner);
				

			} else {
				var spinnerContainer = document.createElement('div');
				applyLocation(spinnerContainer, controlItemArray, true);
				spinnerContainer.style.padding = "0 0 0 0";

				var spinner = spinnerContainer.appendChild(document.createElement("input"));
				spinner.className = "spinner";
				spinner.id = controlID;
				spinner.style.fontFamily = controlItemArray[11];
				spinner.style.fontSize = controlItemArray[12] + 'pt';
				spinner.style.width = (Number((controlItemArray[7]) / 15)) + "px";
				spinner.style.margin = "0px";
				spinner.setAttribute("data-columnID", columnID);
				spinner.setAttribute('data-controlType', controlItemArray[3]);
				spinner.setAttribute("data-control-key", key);
				spinner.setAttribute('data-minval', controlItemArray[29]);
				spinner.setAttribute('data-maxval', controlItemArray[30]);
				spinner.setAttribute('data-increment', controlItemArray[31]);
				spinner.setAttribute("data-Mandatory", controlItemArray[32]);

				if (tabIndex > 0) spinner.tabindex = tabIndex;
				//if (!fControlEnabled) spinnerContainer.disabled = true;
				if (!fControlEnabled) spinner.setAttribute('data-disabled', 'true');

				$(spinner).attr("title", controlItemArray[41]);

				//Add control to relevant tab, create if required.                
				addControl(iPageNo, spinnerContainer);
			}
			break;
		case 64: //ctlText
			var textbox;
			if (Number(controlItemArray[37]) !== 0) {				
				//Multi-line textbox
				textbox = document.createElement('textarea'); //textbox.disabled = false;  //always enabled.
			} else {
				textbox = document.createElement('input');
				var controlDataType = Number(controlItemArray[23]);
				if (controlDataType == 11) { //sqlDate
					textbox.type = "text";
					textbox.className = "datepicker";
				}
				else if (controlDataType == 2 || controlDataType == 4) { //sqlNumeric, sqlInteger
					textbox.className = "number";


					//Add some attributes used by the autoNumeric plugin we are using to validate numeric text boxes
					//var x; //For the loop below
					value = "";
					
					textbox.setAttribute("data-a-dec", $('#txtRecEditControlNumberDecimalSeparator').val()); //Decimal separator
					if (Number(controlItemArray[52]) != 0) { //Use 1000 separator?
						textbox.setAttribute("data-a-sep", $('#txtRecEditControlNumberGroupSeparator').val()); //Thousand separator
					} else {
						textbox.setAttribute("data-a-sep", ''); //No thousand separator
					}
					textbox.setAttribute('data-m-dec', controlItemArray[26]); //Decimal places
					
					//Size of field includes decimals but not the decimal point; For example if Size=6 and Decimals=2 the maximum value to be allowed is 9999.99
					if (controlItemArray[25] == "0") { //No size specified, set a very long limit
						textbox.setAttribute('data-v-min', '-2147483647'); //This is -Int32.MaxValue
						textbox.setAttribute('data-v-max', '2147483647'); //This is Int32.MaxValue
					} else {
						//Determine the length we need and "translate" that to use it in the plugin
						var n = Number(controlItemArray[25]) - Number(controlItemArray[26]); //Size minus decimal places
						for (x = n; x--;) value += "9"; //Create a string of the form "999"
						
						if (controlItemArray[26] != "0") { //If decimal places are specified, add a period and an appropriate number of "9"s
							value += ".";
							for (x = Number(controlItemArray[26]); x--;) value += "9";
						}
						
						textbox.setAttribute('data-v-min', '-' + value);
						textbox.setAttribute('data-v-max', value);

						$(textbox).addClass("autoNumeric"); //Add this class so the Limit plugin won't be attached to this textbox, which was causing clashes between the Limit and AutoNumeric plugins
					}
					
					//Alignment; this is not used by the plugin so we'll add it as a CSS style
					if (controlItemArray[38] == "0") {
						$(textbox).css('text-align', 'left');
					} else if (controlItemArray[38] == "1") {
						$(textbox).css('text-align', 'right');
					} else {
						$(textbox).css('text-align', 'center');
					}
					
					//Blank if zero; set the value of the attribute data-blankIfZeroValue for this textbox depending on its blankIfZero setting and its decimal places
					if (controlItemArray[36] == "1") { //Blank if zero
						textbox.setAttribute("data-blankIfZeroValue", "");
					} else {
						decimalPlaces = Number(controlItemArray[26]);
						if (decimalPlaces == 0) {
							textbox.setAttribute("data-blankIfZeroValue", "0");
						} else {
							value = "0.";
							for (x = decimalPlaces; x--;) value += "0";
							textbox.setAttribute("data-blankIfZeroValue", value);
						}
					}
				}
				else {
					textbox.type = "text";
					textbox.isMultiLine = false;

					if (controlItemArray[35].length > 0) {
						$(textbox).mask(controlItemArray[35]); //One less TODO to do!
					}
								
					//Alignment; this is not used by the plugin so we'll add it as a CSS style
					if (controlItemArray[38] == "0") {
						$(textbox).css('text-align', 'left');
					} else if (controlItemArray[38] == "1") {
						$(textbox).css('text-align', 'right');
					} else {
						$(textbox).css('text-align', 'center');
					}

				}

				if (!fControlEnabled) $(textbox).prop('disabled', true);
				

			}

			$(textbox).attr("title", controlItemArray[41]);

			textbox.id = controlID;
			applyLocation(textbox, controlItemArray, true);
			textbox.style.fontFamily = controlItemArray[11];
			textbox.style.fontSize = controlItemArray[12] + 'pt';
			if (Number(controlItemArray[37]) == 0)  textbox.style.padding = "0 2px 0 2px";
			textbox.setAttribute("data-columnID", columnID);
			textbox.setAttribute('data-controlType', controlItemArray[3]);
			textbox.setAttribute("data-control-key", key);
			textbox.setAttribute("data-Mandatory", controlItemArray[32]);

			if (controlItemArray[25] > 0) {
				//set maximum input length for this control
				$(textbox).attr('maxlength', controlItemArray[25]);
			}

			if (tabIndex > 0) textbox.tabindex = tabIndex;

			if (Number(controlItemArray[37]) != 0) { //Multi-line textbox (i.e. textarea); for this we need a slight adjustment to the height
				textbox.style.height = (Number((controlItemArray[6]) / 15 - 1)) + "px";

				//Alignment				
				if (controlItemArray[38] == "0") {
					$(textbox).css('text-align', 'left');
				} else if (controlItemArray[38] == "1") {
					$(textbox).css('text-align', 'right');
				} else {
					$(textbox).css('text-align', 'center');
				}
				
				//add readonly property for multiline textboxes if necessary
				if (fReadOnly) $(textbox).attr('readonly', 'readonly').css('color', 'gray');

			}

			//Check if control should be disabled (read only or screen read only)
			if (controlItemArray[40] != "0" || controlItemArray[61] != "0") {
				//textbox.setAttribute("disabled", "disabled");
				//$(textbox).addClass("ui-state-disabled");
				$(textbox).prop('disabled', true);
			}		

			//Add control to relevant tab, create if required.                
			addControl(iPageNo, textbox);
			break;
		case 128: //ctlTab
			break;
		case 256: //Label
			span = document.createElement('span');
			applyLocation(span, controlItemArray, false);
			//span.style.backgroundColor = "transparent";
			span.style.backgroundColor = (Number(controlItemArray[9]) !== -2147483633) ? decimalColorToHTMLcolor(controlItemArray[9]) : 'transparent';
			span.style.fontFamily = controlItemArray[11];
			span.style.fontSize = controlItemArray[12] + 'pt';
			label = document.createElement('label');
			span.style.textDecoration = (Number(controlItemArray[15]) != 0) ? "line-through" : "none";
			if ((Number(controlItemArray[10]) !== 0) || (Number(controlItemArray[9]) !== -2147483633)) span.style.color = decimalColorToHTMLcolor(controlItemArray[10]);			//only colour label if not black (0)
			label.style.textDecoration = (Number(controlItemArray[16]) != 0) ? "underline" : "none";
			label.textContent = controlItemArray[8];
			//span.textContent = controlItemArray[8];
			span.appendChild(label);
			span.setAttribute("data-control-key", key);

			//replaces the SetControlLevel function in recordDMI.ocx.
			span.style.zIndex = 0;

			//if (!fControlEnabled) span.disabled = true;

			addControl(iPageNo, span);

			break;
		case 512: //Frame
			
			var fieldset = document.createElement('fieldset');
			applyLocation(fieldset, controlItemArray, true);
			fieldset.style.backgroundColor = ((controlItemArray[9] !== '-2147483633') ? decimalColorToHTMLcolor(controlItemArray[9]) : "transparent");
			//fieldset.style.color = "Black";			
			fieldset.style.padding = "0px";
			if (controlItemArray[8].length > 0) {
				var legend = fieldset.appendChild(document.createElement('legend'));
				if (Number(controlItemArray[10]) !== 0) legend.style.color = decimalColorToHTMLcolor(controlItemArray[10]);
				legend.style.backgroundColor = (Number(controlItemArray[9]) !== -2147483633) ? decimalColorToHTMLcolor(controlItemArray[9]) : 'transparent';
				legend.style.fontFamily = controlItemArray[11];
				legend.style.fontSize = controlItemArray[12] + 'pt';
				legend.style.fontWeight = (Number(controlItemArray[13]) != 0) ? "bold" : "normal";
				legend.style.textDecoration = (Number(controlItemArray[16]) != 0) ? "underline" : "none";
				if (Number(controlItemArray[15]) != 0) {
					span = document.createElement('span');
					span.style.textDecoration = "line-through";
					span.appendChild(document.createTextNode(controlItemArray[8].split('&&').join('&')));
					legend.appendChild(span);
				} else {
					legend.appendChild(document.createTextNode(controlItemArray[8].split('&&').join('&')));
				}				
				legend.className = 'ui-helper-reset';

			}
			fieldset.setAttribute("data-control-key", key);

			addControl(iPageNo, fieldset);

			break;
		case 1024: //ctlPhoto
			image = document.createElement('img');
			image.id = controlID;
			applyLocation(image, controlItemArray, true);
			image.style.border = "1px solid gray";
			image.style.padding = "0px";
			image.setAttribute("data-columnID", columnID);
			image.setAttribute('data-controlType', controlItemArray[3]);
			image.setAttribute("data-control-key", key);
			image.setAttribute('data-OleType', controlItemArray[55]); // == 2 ? 3 : controlItemArray[55]);
			image.setAttribute('data-maxEmbedSize', controlItemArray[57]);

			if (!fControlEnabled) image.disabled = true;

			//Check if control should be disabled (read only or screen read only)
			if (controlItemArray[40] != "0" || controlItemArray[61] != "0") {
				image.setAttribute("disabled", "disabled");
				$(image).addClass("ui-state-disabled");
			}
			
			//var path = window.ROOT + 'Home/ShowImageFromDb?imageID=' + controlItemArray[50];

			//image.setAttribute('src', path);
			$(image).attr("title", controlItemArray[41]);

			//Add control to relevant tab, create if required.                
			addControl(iPageNo, image);


			break;


		case 2048: //ctlCommand
			button = document.createElement('input');
			button.type = "button";
		    button.id = controlID;
			button.value = controlItemArray[8];
			applyLocation(button, controlItemArray, true);
			button.style.whiteSpace = "normal";
		    button.style.padding = "0px";
			button.setAttribute("data-columnID", columnID);
			button.setAttribute('data-controlType', controlItemArray[3]);
			button.setAttribute("data-control-key", key);
			button.setAttribute("data-columnName", "ID_" + controlItemArray[42]);
			button.setAttribute("data-linkTableID", controlItemArray[42]);
			button.setAttribute("data-linkOrderID", controlItemArray[43]);
			button.setAttribute("data-linkViewID", controlItemArray[44]);

			if (tabIndex > 0) button.tabindex = tabIndex;
			//button.disabled = false;    //always enabled
			$(button).attr("title", controlItemArray[41]);

			addControl(iPageNo, button);

			break;
			

		case 4096: //ctlWorking Pattern
			//TODO - android browser/tablet adjustment
			var fontSize = Number(controlItemArray[12]);
			top = (Number(controlItemArray[4]) / 15);
			left = (Number(controlItemArray[5]) / 15);
			height = (Number((controlItemArray[6]) / 15) - 2);
			width = (Number((controlItemArray[7]) / 15) - 2);
			if (controlItemArray[19] == "0") {
				borderCss = "border-style: none;";
			} else {
				borderCss = "border: 1px solid #999;";
				width -= 2;
				height -= 2;		
			}

			fieldset = document.createElement("fieldset");
			fieldset.id = controlID;
			fieldset.setAttribute("data-columnID", columnID);
			fieldset.setAttribute('data-controlType', controlItemArray[3]);
			fieldset.setAttribute("data-datatype", "Working Pattern");
			fieldset.setAttribute("data-control-key", key);
			fieldset.setAttribute("data-Mandatory", controlItemArray[32]);

			fieldset.style.position = "absolute";
			fieldset.style.top = top + "px";
			fieldset.style.left = left + "px";
			fieldset.style.width = width + "px";
			fieldset.style.height = height + "px";
			fieldset.style.padding = "0px";
			fieldset.style.border = borderCss;

			$(fieldset).attr("title", controlItemArray[41]);

			var offsetLeft;

			var table = fieldset.appendChild(document.createElement("table"));
			var tr = table.appendChild(document.createElement("tr"));
			//Get local weekday names...
			var weekDays = [];
			
			//steal the weekday names from the datepicker library.
			var userLocale;
			if ($.datepicker.regional[window.UserLocale]) {
				userLocale = window.UserLocale;
			}
			else if ($.datepicker.regional[window.UserLocale.substr(0, 2)]) {
				userLocale = window.UserLocale.substr(0, 2);
			}

			if (userLocale !== undefined) {
				for (var iDayName = 0; iDayName < 7; iDayName++) {
					weekDays.push($.datepicker.regional[userLocale].dayNames[iDayName].substr(0, 1));
				}
			} else {
				weekDays = new Array("S", "M", "T", "W", "T", "F", "S");
			}

			for (var i = 0; i <= 7; i++) {
				offsetLeft = (width / 10) * (i + 1);
				var td = tr.appendChild(document.createElement("td"));
				td.style.textAlign = "center";
				var dayLink = td.appendChild(document.createElement("a"));
				
				switch (i) {
				case 0:
					dayLink.textContent = " ";
					break;
				default :
					dayLink.textContent = weekDays[i - 1];
					break;
				}

				//Day labels (they are in reality anchors ['a'])
				dayLink.style.fontFamily = controlItemArray[11];
				dayLink.style.fontSize = fontSize + "pt";
				dayLink.style.fontWeight = (Number(controlItemArray[13]) != 0) ? "bold" : "normal";
				if (Number(controlItemArray[10]) !== 0) dayLink.style.color = decimalColorToHTMLcolor(controlItemArray[10]);
				dayLink.style.position = "absolute";
				dayLink.style.top = "0px";
				dayLink.style.left = offsetLeft + "px";
				dayLink.style.textDecoration = "none";
				dayLink.style.cursor = "default";
				dayLink.style.textAlign = "center";
				dayLink.style.width = "1em";
				dayLink.style.minWidth = "13px";
				$(dayLink).attr('href', '#');
				$(dayLink).attr('data-checkboxes', controlID + "_" + ((i * 2) + 1) + "," + controlID + "_" + ((i * 2) + 2)); //Data attribute to hold associated checkboxes

				//If the control is enabled, add an event on clicking the link to toggle its associated checkboxes
				if (fControlEnabled) {
					$(dayLink).click(function(ev1) {
						ev1.preventDefault();
						toggleCheckboxes(this);
					});
				}
			}
			
			for (i = 0; i <= 7; i++) {
				if (i > 0) {
					offsetLeft = (width / 10) * (i + 1);
					//AM Boxes
					var amCheckbox = fieldset.appendChild(document.createElement("input"));
					amCheckbox.type = "checkbox";
					amCheckbox.id = controlID + "_" + ((i * 2) + 1);
					amCheckbox.style.padding = "0px";
					amCheckbox.style.position = "absolute";
					amCheckbox.style.top = fontSize * 12 / 8 + "px";
					amCheckbox.style.left = offsetLeft + "px";
					if (!fControlEnabled) amCheckbox.disabled = true;

					//PM Boxes
					var pmCheckbox = fieldset.appendChild(document.createElement("input"));
					pmCheckbox.type = "checkbox";
					pmCheckbox.id = controlID + "_" + ((i * 2) + 2);
					pmCheckbox.style.padding = "0px";
					pmCheckbox.style.position = "absolute";
					pmCheckbox.style.top = fontSize * 26 / 8 - 2 + "px";
					pmCheckbox.style.left = offsetLeft + "px";
					if (!fControlEnabled) pmCheckbox.disabled = true;
				}
			}

			var checkboxesToAssociate, j;
			
			//AM label (it's actually an anchor ['a'])
			var link = document.createElement("a");
			link.textContent = "AM";
			link.style.fontFamily = controlItemArray[11];
			link.style.fontSize = fontSize + 'pt';
			link.style.fontWeight = (Number(controlItemArray[13]) != 0) ? "bold" : "normal";
			link.style.position = "absolute";
			link.style.top = fontSize + 5 + "px";
			link.style.left = "4px";
			link.style.textDecoration = "none";
			link.style.cursor = "default";
			$(link).attr('href', '#');
			checkboxesToAssociate = "";
			for (j = 1; j <= 15; j += 2) {
				checkboxesToAssociate += controlID + "_" + j.toString() + ",";
			}
			$(link).attr('data-checkboxes', checkboxesToAssociate.substring(0, checkboxesToAssociate.length - 1)); //Data attribute to hold associated checkboxes
		
			//If the control is enabled, add an event on clicking the link to toggle its associated checkboxes
			if (fControlEnabled) {
				$(link).click(function (ev2) {
					ev2.preventDefault();
					toggleCheckboxes(this);
				});
			}

			fieldset.appendChild(link);

			//PM label (it's actually an anchor ['a'])
			link = document.createElement("a");
			link.textContent = "PM";
			link.style.fontFamily = controlItemArray[11];
			link.style.fontSize = fontSize + 'pt';
			link.style.fontWeight = (Number(controlItemArray[13]) != 0) ? "bold" : "normal";
			link.style.position = "absolute";
			link.style.top = fontSize * 21 / 8 + 5 + "px";
			link.style.left = "4px";
			link.style.textDecoration = "none";
			link.style.cursor = "default";
			$(link).attr('href', '#');
			checkboxesToAssociate = "";
			for (j = 2; j <= 16; j += 2)
			{
				checkboxesToAssociate += controlID + "_" + j.toString() + ",";
			}
			$(link).attr('data-checkboxes', checkboxesToAssociate.substring(0, checkboxesToAssociate.length - 1)); //Data attribute to hold associated checkboxes
			
			//If the control is enabled, add an event on clicking the link to toggle its associated checkboxes
			if (fControlEnabled) {
				$(link).click(function(ev3) {
					ev3.preventDefault();
					toggleCheckboxes(this);
				});
			}

			fieldset.appendChild(link);

			//ADD FIELDSET AND ITS CONTENTS.
			addControl(iPageNo, fieldset);
			break;
		case 8192: //2 ^ 13: //ctlLine
			var line = document.createElement('div');
			applyLocation(line, controlItemArray, true);
			if (controlItemArray[20] != 0) {
				//Vertical line
				line.style.height = "1px";
			} else {
				line.style.width = "1px";
			}

			line.style.backgroundColor = "gray";
			line.style.padding = "0px";
			line.setAttribute("data-control-key", key);
			//.visible = true
			//.container = tabnumber
			//.alignment
			//.border
			//.top
			//.left
			//.height
			//.width
			//.caption
			//tabIndex
			//.backColor
			//.oletype
			//font, fontsize, fontbold, fontitalic, fontstrikethrough, fontunderline
			//forecolor
			//enabled
			//columnid
			//displaytype
			//navto
			//navin
			//navonsave
			//radio options and border
			//spinner min max increment spinnerposition
			//numeric separator, alignment
			//date/number max size
			//format
			//mask
			//showliterals
			//allow space
			//screenreadonly

			addControl(iPageNo, line);

			break;
		case Math.pow(2, 14): //ctlNavigation

			var displayType = controlItemArray[17];
			//var navigateIn = controlItemArray[59];
			var displayText = (controlItemArray[8].length <= 0 ? "Navigate..." : controlItemArray[8]);
			//var navigateTo = controlItemArray[58];
			switch (Number(displayType)) {

				case 0: //Hyperlink
					var hyperlink = document.createElement("a");
					applyLocation(hyperlink, controlItemArray, false);
					hyperlink.style.fontFamily = controlItemArray[11];
					hyperlink.style.fontSize = controlItemArray[12] + 'pt';
					hyperlink.style.fontWeight = (Number(controlItemArray[13]) != 0) ? "bold" : "normal";
					hyperlink.style.textDecoration = "underline"; //(Number(controlItemArray[16]) != 0) ? "underline" : "none";
					hyperlink.appendChild(document.createTextNode(displayText));
					if(Number(controlItemArray[10]) !== 0) hyperlink.style.color = controlItemArray[10];
					hyperlink.style.backgroundColor = controlItemArray[9];
					hyperlink.setAttribute("href", formatAddress(controlItemArray[58]));
					hyperlink.setAttribute("target", "_blank");

					hyperlink.id = controlID;
					hyperlink.style.padding = "0px";
					hyperlink.setAttribute("data-columnID", columnID);
					hyperlink.setAttribute('data-controlType', controlItemArray[3]);
					hyperlink.setAttribute("data-control-key", key);

					if (tabIndex > 0) hyperlink.tabindex = tabIndex;

					addControl(iPageNo, hyperlink);

					break;
				case 1: //Button
					button = document.createElement("input");
					button.type = "button";
					applyLocation(button, controlItemArray, false);
					button.style.fontFamily = controlItemArray[11];
					button.style.fontSize = controlItemArray[12] + 'pt';
					button.style.fontWeight = (Number(controlItemArray[13]) != 0) ? "bold" : "normal";
					button.style.textDecoration = (Number(controlItemArray[16]) != 0) ? "underline" : "none";
					button.value = displayText;
					button.style.color = controlItemArray[10];
					button.style.backgroundColor = controlItemArray[9];

					button.setAttribute("onclick", "window.open('" + formatAddress(controlItemArray[58]) + "')");

					button.id = controlID;
					button.style.padding = "0px";
					button.setAttribute("data-columnID", columnID);
					button.setAttribute('data-controlType', controlItemArray[3]);
					button.setAttribute("data-control-key", key);

					if (tabIndex > 0) button.tabindex = tabIndex;

					addControl(iPageNo, button);

					break;
				case 2: //Browser
					var el = document.createElement("iframe");
					applyLocation(el, controlItemArray, true);
					el.id = controlID;
					el.setAttribute("data-columnID", columnID);
					el.setAttribute('data-controlType', controlItemArray[3]);
					el.setAttribute("data-control-key", key);
					if (tabIndex > 0) el.tabindex = tabIndex;

					addControl(iPageNo, el);
					el.setAttribute('src', formatAddress(controlItemArray[58]));

					break;
				case 3: //Hidden
					break;

			}

			//TODO: Nav control always .disabled = false.
			//if (tabIndex > 0) checkbox.tabindex = tabIndex;
			break;
		case Math.pow(2, 15): // 32768 - ctlColourPicker
			//The color picker plugin takes an input box, hides it and creates some divs to show the color picker;
			//so we need two input boxes: one that will contain the color itself and another one to be used by the plugin
			var textboxColorPickerPlugin; //To be used by the plugin
			textboxColorPickerPlugin = document.createElement('input');
			textboxColorPickerPlugin.type = "text";
			textboxColorPickerPlugin.className = "colorPicker";
			textboxColorPickerPlugin.id = "colorPicker_" + controlID.substr(3);
			textboxColorPickerPlugin.setAttribute("data-style-top", Number(controlItemArray[4]) / 15);
			textboxColorPickerPlugin.setAttribute("data-style-left", Number(controlItemArray[5]) / 15);
			textboxColorPickerPlugin.setAttribute("data-style-height", Number(controlItemArray[6]) / 15);
			textboxColorPickerPlugin.setAttribute("data-style-width", Number(controlItemArray[7]) / 15);
			textboxColorPickerPlugin.setAttribute("data-readonly", !fControlEnabled);
			
			var textboxColorPicker; //To contain the value that will be saved
			textboxColorPicker = document.createElement('input');
			textboxColorPicker.type = "text";
			textboxColorPicker.id = controlID;
			textboxColorPicker.value = controlItemArray[24];
			textboxColorPicker.style.display = "none";
			textboxColorPicker.className = "colorPicker";
			textboxColorPicker.setAttribute("data-columnID", columnID);
			textboxColorPicker.setAttribute('data-controlType', controlItemArray[3]);
			textboxColorPicker.setAttribute("data-control-tag", key);
			textboxColorPicker.setAttribute("data-Mandatory", controlItemArray[32]);

			//Set attributes that link both controls
			textboxColorPicker.setAttribute("data-associated-control-id", textboxColorPickerPlugin.id);
			textboxColorPickerPlugin.setAttribute("data-associated-control-id", textboxColorPicker.id);
			
			// Add both controls to the page
			addControl(iPageNo, textboxColorPickerPlugin);
			addControl(iPageNo, textboxColorPicker);
			
			//Note: the plugin is hooked up to the control in the updateControl function
			break;
		default:
			break;
	}
}

//Function below nicked from http://bytes.com/topic/javascript/insights/636088-function-convert-decimal-color-number-into-html-hex-color-string
function decimalColorToHTMLcolor(number) {
	//converts to a integer
	var intnumber = number - 0;

	// isolate the colors - really not necessary
	var red, green, blue;

	// needed since toString does not zero fill on left
	var template = "#000000";

	// in the MS Windows world RGB colors are 0xBBGGRR because of the way Intel chips store bytes
	red = (intnumber & 0x0000ff) << 16;
	green = intnumber & 0x00ff00;
	blue = (intnumber & 0xff0000) >>> 16;

	// mask out each color and reverse the order
	intnumber = red | green | blue;

	// toString converts a number to a hexstring
	var HTMLcolor = intnumber.toString(16);

	//template adds # for standard HTML #RRGGBB
	HTMLcolor = template.substring(0, 7 - HTMLcolor.length) + HTMLcolor;

	return HTMLcolor;
}

function toggleCheckboxes(control) {
	var checkboxes = $(control).attr('data-checkboxes').split(',');
	//Loop over the checkboxes
	$.each(checkboxes, function (k, v)
	{
		$('#' + v).attr('checked', !$('#' + v).attr('checked')); //Toggle the checkbox
	});
	
	//Indicate that a control has changed and enable the Save button
	$("#ctlRecordEdit #changed").val("true");
	menu_toolbarEnableItem("mnutoolSaveRecord", true);
}

function addHTMLControlValues(controlValues) {

	try {
		var controlValuesArray = controlValues.split("\t");
	} catch (e) { return false; }

	var lngColumnID = 0;
	var sValue;

	//NB This function is only valid for radio buttons (option groups) and dropdown lists (not lookups).

	for (var i = 0; i < controlValuesArray.length; i++) {

		sValue = controlValuesArray[i];

		if (lngColumnID > 0) {
			if (sValue.length > 0) {

				//get the column type, then add this value to it/them.
				//we use a .each function, as there may be more than one column with this ID on the screen.
				$("#ctlRecordEdit").find("[data-columnID='" + lngColumnID + "']").each(function () {


					//Option Groups
					if (($(this).is("fieldset")) && ($(this).attr("data-datatype") === "Option Group")) {
						//unique groupname for the radio buttons.
						var uniqueID = $(this).attr("id");
						var alignment = $(this).attr("data-alignment");

						var fieldset = document.getElementById($(this).attr("id"));

						var radio = fieldset.appendChild(document.createElement("input"));
						var label = fieldset.appendChild(document.createElement("label"));

						radio.type = "radio";
						radio.className = "radio";
						radio.name = uniqueID;  //used to tie separate radio buttons together.
						radio.value = sValue;
						radio.id = uniqueID + "_" + i;

						if (alignment == 0) {
							//Vertical alignment
							radio.style.position = "absolute";
							radio.style.top = (i * 16) + "px";
							radio.style.left = "12px";
							radio.style.padding = "0px";
							radio.style.margin = "0px";

							//add text to radio button
							label.style.position = "absolute";
							label.style.top = (i * 16) + "px";
							label.style.left = "28px";
							label.style.padding = "0px";
							label.htmlFor = uniqueID + "_" + i;
							label.appendChild(document.createTextNode(sValue));
						}
						if (alignment == 1) {
							$(this).css("padding-left", "4px");
							$(this).css("padding-top", "2px");
							//Horizontal alignment
							//radio.style.padding = "0px";
							radio.style.borderStyle = 'none';
							radio.style.verticalAlign = 'middle';

							//add text to radio button
							label.style.marginLeft = "3px";
							label.style.marginRight = "32px";
							label.htmlFor = uniqueID + "_" + i;
							label.style.verticalAlign = 'middle';
							
							label.appendChild(document.createTextNode(sValue));
						}
					}

					//Dropdown Lists
					if ($(this).is("select")) {
						var option = document.createElement('option');
						option.value = sValue;
						option.appendChild(document.createTextNode(sValue));
						$(this).append(option);
					}
				});
			}
		}
		else {
			if (sValue.length > 0) {
				//set the column ID to apply list to.
				lngColumnID = Number(sValue);
			}
			else { return false; }
		}
	}
	return false;
}

function recEdit_setData(columnID, value) {
	
	//Set the given column's value
	//copied from recordDMI.ocx        	
	if (columnID.toUpperCase() == "TIMESTAMP") {
		// The column is the timestamp column.
		$("#txtRecEditTimeStamp").val(value);
	}
	else {
		var fIsIDColumn = false;

		var ubound = (window.mavIDColumns.length);
		for (var i = 0; i < (ubound); i++) {
				if (window.mavIDColumns[i][0] == Number(columnID)) {
					this.mavIDColumns[i][2] = Number(value);
					fIsIDColumn = true;
				}
		}
		
		if (!fIsIDColumn) {
			if ((value != null) && (value != undefined))
				updateControl(Number(columnID), value);
		}
	}
}

function recEdit_setTimeStamp() {
	//TODO:
	
}

function recEdit_setRecordID(plngRecordID) {
	var frmRecordEditForm = document.getElementById("frmRecordEditForm");
	frmRecordEditForm.txtCurrentRecordID.value = plngRecordID;
	//frmRecordEditForm.ctlRecordEdit.recordID = plngRecordID;
}

function recEdit_setCopiedRecordID(plngRecordID) {
	var frmRecordEditForm = document.getElementById("frmRecordEditForm");
	frmRecordEditForm.txtCopiedRecordID.value = plngRecordID;
	//frmRecordEditForm.ctlRecordEdit.CopiedRecordID = plngRecordID;
}

function recEdit_setParentTableID(plngParentTableID) {
	var frmRecordEditForm = document.getElementById("frmRecordEditForm");
	frmRecordEditForm.txtCurrentParentTableID.value = plngParentTableID;
	//frmRecordEditForm.ctlRecordEdit.ParentTableID = plngParentTableID;
}

function recEdit_setParentRecordID(plngParentRecordID) {
	var frmRecordEditForm = document.getElementById("frmRecordEditForm");
	frmRecordEditForm.txtCurrentParentRecordID.value = plngParentRecordID;
	//frmRecordEditForm.ctlRecordEdit.ParentRecordID = plngParentRecordID;
}

function updateControl(lngColumnID, value) {

	//get the column type, then add this value to it/them.
	$("#ctlRecordEdit").find("[data-columnID='" + lngColumnID + "']").each(function () {

		//TODO: is this control tagged to a column?

		//TODO: is this controls columnID = lngColumnID?

		//Get the controlType from the ID
		try {
			var arrIDProps = $(this).attr("id").split("_");
		} catch (e) {
			return false;
		}

		var controlType = Number(arrIDProps[2]);
		if ($(this).is("textarea")) {
			//was TDBText6Ctl.TDBText
			$(this).val(value);
		}

		//TODO if mask
		// .Text = RTrim(CStr(pvValue) & vbNullString)

		//Input type controls...
		var oleType;
		var filename;
		if ($(this).is("input")) {
			switch ($(this).attr("type")) {
				case "text":
					if ($(this).hasClass("datepicker")) {
						//date control
						if (value.length == 0) {
							$(this).val("");
						} else {
							$(this).val(OpenHR.ConvertSQLDateToLocale(value));
						}
					} else if ($(this).hasClass("colorPicker")) { //Color picker
						var textboxId = $(this).attr("id"); //This is the ID of the textbox that the color picker plugin will use to store the value of the selected color
						var colorPickerId = $("#" + textboxId).attr('data-associated-control-id'); //This is the ID of the color picker associated with the textbox above
						$("#" + textboxId).val(value); //Set the value that came from the database
						$("#" + colorPickerId).spectrum("destroy");
						$("#" + colorPickerId + "_div").remove(); //Remove the previous div that existed for the plugin (if any)
						//Hook up the plugin to the control
						var initialColor = (parseInt(value, 10)).toString(16);
						initialColor = Array(7 - initialColor.length).join("0") + initialColor;
						initialColor = initialColor.substr(4, 2) + initialColor.substr(2, 2) + initialColor.substr(0, 2);
						$("#" + colorPickerId).spectrum({
							color: "#" + initialColor, //Set the initial color
							className: "colorPicker",
							cancelText: "", //Hide the Cancel button
							disabled: ($("#" + colorPickerId).attr("data-readonly") == "true"),
							change: function(color) { //On selecting a color...
								var newColor = color.toHex();
								newColor = newColor.substr(4, 2) + newColor.substr(2, 2) + newColor.substr(0, 2);
								$("#" + textboxId).val(parseInt(newColor, 16)).change(); //We need to trigger the change event here so the Save button is enabled
							}
						});
						//After the plugin has been applied, there will be an added div (containing other divs); these need to be repositioned and styled
						//DIV
						$("#" + colorPickerId).next().first().attr("id", colorPickerId + "_div"); //Assign an ID to the new DIV
						$("#" + colorPickerId).next().css("top", $("#" + colorPickerId).attr("data-style-top") - 2 + "px");
						$("#" + colorPickerId).next().css("left", $("#" + colorPickerId).attr("data-style-left") - 1 + "px");
						$("#" + colorPickerId).next().css("height", $("#" + colorPickerId).attr("data-style-height") + "px");
						$("#" + colorPickerId).next().css("width", $("#" + colorPickerId).attr("data-style-width") + "px");
						$("#" + colorPickerId).next().css("position", "absolute");
						$("#" + colorPickerId).next().css("background", "none");
						$("#" + colorPickerId).next().css("border", "none");
						//First inner div of the div above
						$("#" + colorPickerId).next().children().first("sp-preview").css("height", $("#" + colorPickerId).attr("data-style-height") + "px");
						$("#" + colorPickerId).next().children().first("sp-preview").css("width", $("#" + colorPickerId).attr("data-style-width") - 16 + "px");
					} else {
						$(this).val(value);
					}
					break;
				case "number":
					$(this).val(Number(value));
					break;
				case "checkbox":
					$(this).prop("checked", value.toLowerCase() == "true" ? true : false);
					break;							
				case "button":
					if (controlType == 8) {						
						//OLE																		
						if (value.indexOf('::LINKED_OLE_DOCUMENT') > 0) {
							oleType = 3;
						}
						else if (value.indexOf('::EMBEDDED_OLE_DOCUMENT') > 0) {
							oleType = 2;
						} else {
							oleType = $(this).attr('data-OleType');
						}
						filename = value.replace('::LINKED_OLE_DOCUMENT::', '').replace('::EMBEDDED_OLE_DOCUMENT::', '');
						var filesize = $('#txtData_' + lngColumnID).attr('data-filesize');
						var createdate = $('#txtData_' + lngColumnID).attr('data-createdate');
						var modifydate = $('#txtData_' + lngColumnID).attr('data-filemodifydate');
						
						//OLE_LOCAL = 0
						//OLE_SERVER = 1
						//OLE_EMBEDDED = 2
						//OLE_UNC = 3
						var strOLEType = 'OLE';
						
						switch (Number(oleType)) {
							case 0:
								strOLEType = 'Local';
								break;
							case 1:
								strOLEType = 'Server';
								break;
							case 2:
								strOLEType = (filename.length > 0 ? 'Embedded' : 'Embed');
								break;
							case 3:
								strOLEType = (filename.length > 0 ? 'Linked' : 'Link');								
								break;
							default:
								strOLEType = 'failed to load caption';
								break;
						}
						
						var tooltipText = (filename.length > 0 ? filename + ' ' + strOLEType : 'empty');

						$(this).val(strOLEType);
						$(this).attr('title', tooltipText);
						$(this).attr('data-fileName', filename);						
						$(this).removeClass("Embed Embedded Link Linked");
						$(this).addClass(strOLEType);


					}
					if (controlType == Math.pow(2, 14)) {
						//Navigation Control						
						if (value.length <= 0) {
							$(this).attr("onclick", "javascript:window.open('about:blank');");
						} else {
							$(this).attr("onclick", "javascript:window.open('" + formatAddress(value) + "');");
						}
					}
					else if (controlType == 2048) {
						//Link Button
						var lngLinkTableID = $(this).attr("data-linkTableID");
						var lngLinkOrderID = $(this).attr("data-linkOrderID");
						var lngLinkViewID = $(this).attr("data-linkViewID");
						
						$(this).attr('onclick', 'javascript:linkButtonClick(' + lngLinkTableID + ',' + lngLinkOrderID + ',' + lngLinkViewID + ');');												
					}
					break;
				default:
					$(this).val(value);
			}
			
			//refresh 'autoNumeric' Columns
			if ($(this).hasClass('number')) {
				$(this).autoNumeric('update');
			}
		}

		if ($(this).is("img")) {

			filename = value.replace('::LINKED_OLE_DOCUMENT::', '').replace('::EMBEDDED_OLE_DOCUMENT::', '');
			var msPhotoPath = $('#frmRecordEditForm #txtPicturePath').val();
			
			if (value.indexOf('::LINKED_OLE_DOCUMENT::') >= 0) {
				//no linked photos in the web. Only IE11 currently supports local images. 
				$(this).attr('src', '../Content/Images/anonymous.png');
				oleType = 3;
			}
			else if (value.indexOf('::EMBEDDED_OLE_DOCUMENT::') >= 0) {
				//point source at hidden tag value.
				$(this).attr('src', 'data:image/jpeg;base64, ' + $('#txtData_' + lngColumnID).attr('data-Img'));
				oleType = 2;
			} else {
				if (value != "") {
					$(this).attr('src', msPhotoPath + "\\" + value);
					oleType = $(this).attr('data-OleType');
				} else {
					$(this).attr('src', '../Content/Images/anonymous.png');
				}
			}
			
			var filesize = $('#txtData_' + lngColumnID).attr('data-filesize');
			var createdate = $('#txtData_' + lngColumnID).attr('data-createdate');
			var modifydate = $('#txtData_' + lngColumnID).attr('data-filemodifydate');

			//OLE_LOCAL = 0
			//OLE_SERVER = 1
			//OLE_EMBEDDED = 2
			//OLE_UNC = 3
			var strOLEType = 'OLE';

			switch (Number(oleType)) {
				case 0:
					strOLEType = '(Local)';
					break;
				case 1:
					strOLEType = '(Server)';
					break;
				case 2:
					strOLEType = '(Embedded)';
					break;
				case 3:
					strOLEType = (filename.length > 0 ? '(Linked)' : '(Link)');
					break;
				default:
					strOLEType = 'failed to load caption';
					break;
			}

			var tooltipText = (filename.length > 0 ? filename + ' ' + strOLEType : 'empty');

			$(this).val(strOLEType);
			$(this).attr('title', tooltipText);
			$(this).attr('data-fileName', filename);
			

		}

		//Working pattern & Option group
		if ($(this).is("fieldset")) {
			if ($(this).attr("data-datatype") === "Working Pattern") {
				//ensure the value is 14 characters long.
				if (value.length < 14) value = value.concat("              ").substring(0, 14);
				var tthisId = "#" + $(this).attr("id");
				//tick relevant boxes.
				for (var i = 1; i <= 14; i++) {
					$(tthisId + "_" + (i + 2)).prop("checked", value.substring(i - 1, i) != " " ? true : false);
				}
			}

			if ($(this).attr("data-datatype") === "Option Group") {
				//TODO
				$("input[name='" + $(this).attr("id") + "'][value='" + value + "']").prop('checked', true);
			}
		}


		if ($(this).is("select")) {
			//does value exist in the dropdown?
			if ($(this).find('option[value="' + value + '"]').length) {
				$(this).val(value);
			} else {

				if ($(this).attr("data-columntype") == "lookup") $(this).empty();	//For lookups, clear out all values, so the newly selected value is all there is.							

				var option = document.createElement('option');
				option.value = value;
				option.appendChild(document.createTextNode(value));
				$(this).append(option);
				$(this).val(value);
			}
		}


		////ComboBox
		//if (($(this).is("select")) && (this.length > 0)) {
		//	$(this).val(value);
		//}

		////Lookup
		//if (($(this).is("select")) && (this.length == 0)) {
		//	var option = document.createElement('option');
		//	option.value = value;
		//	option.appendChild(document.createTextNode(value));
		//	$(this).append(option);
		//	$(this).val(value);
		//}

		//Option Group

		//OLE

		//Spinner - done nwith number above..

		//Nav controls
		if ($(this).is("a")) {		
			if ((value == null) || (value == undefined)) {
				$(this).attr("href", "about:blank");
			} else {
				if (value.length <= 0) {
					$(this).attr("href", "about:blank");
				} else {
					$(this).attr("href", formatAddress(value));
				}
			}
		}

		if ($(this).is('iframe')) {

			if ((value == null) || (value == undefined)) {
				$(this).attr('src', 'about:blank');
			} else {
				if (value.length <= 0) {
					$(this).attr('src', 'about:blank');
				} else {
					$(this).attr('src', formatAddress(value));
				}
			}
		}


	});


}

function getTabCaption(tabNumber) {
	
	var psNewValues = $("#txtRecEditTabCaptions").val();
	try {
		var arr = psNewValues.split("\t");
	} catch (e) {
		return false;
	}

	var tabCaption = arr[tabNumber - 1].split('&&').join('&');

	return tabCaption;

}



function TBCourseRecordID() {
	// Training Booking specific.
	// Return the Course Record ID.
	// Used when editing a Training Booking record.
	var iLoop;

	var TBCourseRecordID = 0;
	var mlngCourseTableID = $("#txtTB_CourseTableID").val();
	var mlngParentTableID = $("#txtCurrentParentTableID").val();
	var mlngParentRecordID = $("#txtCurrentParentRecordID").val();

	if (mlngCourseTableID > 0) {
		if (mlngCourseTableID == mlngParentTableID) {
			TBCourseRecordID = mlngParentRecordID;
		} else {
			//TODO: Linked record not saved?
			//For iLoop = 1 To UBound(mavIDColumns, 2)
			//If UCase(mavIDColumns(2, iLoop)) = "ID_" & Trim(Str(mlngCourseTableID)) Then
			//TBCourseRecordID = mavIDColumns(3, iLoop)
			//Exit For
			//End If
			//Next iLoop
		}
	}
}

function TBEmployeeRecordID() {
	// Training Booking specific.
	// Return the Employee Record ID.
	// Used when editing a Training Booking record.
	var iLoop;

	var TBEmployeeRecordID = 0;
	var mlngEmployeeTableID = $("#txtTB_EmpTableID").val();
	var mlngParentTableID = $("#txtCurrentParentTableID").val();
	var mlngParentRecordID = $("#txtCurrentParentRecordID").val();

	if (mlngEmployeeTableID > 0) {
		if (mlngEmployeeTableID == mlngParentTableID) {
			TBEmployeeRecordID = mlngParentRecordID;
		} else {
			//TODO: linked records?
			//For iLoop = 1 To UBound(mavIDColumns, 2)
			//If UCase(mavIDColumns(2, iLoop)) = "ID_" & Trim(Str(mlngEmployeeTableID)) Then
			//TBEmployeeRecordID = mavIDColumns(3, iLoop)
			//Exit For
			//End If
			//Next iLoop
		}
	}
}

function TBBookingStatusValue() {
	//TODO: 
	//' Training Booking specific.
	//' Return the Booking Status Value.
	//' Used when editing a Training Booking record.
	//Dim sTag As String
	//Dim objControl As Control
	//Dim objScreenControl As clsScreenControl

	//TBBookingStatusValue = ""

	//If mlngTBStatusColumnID > 0 Then
	//For Each objControl In UserControl.Controls
	//sTag = objControl.Tag

	//' Check if it is a user editable control.
	//If Len(sTag) > 0 Then
	//' Check that the control is associated with a column in the current table/view,
	//' and is updatable.
	//If mobjScreenControls.Item(sTag).ColumnID = mlngTBStatusColumnID Then
	//If TypeOf objControl Is TDBText6Ctl.TDBText Then

	//' Multi-line character field from a masked textbox (CHAR type column). Save the text from the control.
	//If IsNull(ConvertData(objControl.Text, mobjScreenControls.Item(sTag).DataType)) Then
	//TBBookingStatusValue = ""
	//Else
	//TBBookingStatusValue = ConvertData(objControl.Text, mobjScreenControls.Item(sTag).DataType)
	//End If

	//ElseIf TypeOf objControl Is TDBMask6Ctl.TDBMask Then
	//' Character field from a masked textbox (CHAR type column). Save the text from the control.
	//If Len(objControl.Value) = 0 Then
	//TBBookingStatusValue = ""
	//Else
	//TBBookingStatusValue = objControl.Text
	//End If

	//ElseIf TypeOf objControl Is TextBox Then
	//' Character field from an unmasked textbox (CHAR type column). Save the text from the control.
	//TBBookingStatusValue = objControl.Text

	//ElseIf TypeOf objControl Is ComboBox Then
	//' Character field from a combo (CHAR type column). Save the text from the combo.
	//TBBookingStatusValue = objControl.Text

	//ElseIf TypeOf objControl Is COAInt_Lookup Then
	//' Lookup field from a combo (unknown type column). Get the column type and save the appropraite value from the combo.
	//Select Case mobjScreenControls.Item(sTag).DataType
	//Case sqlVarChar, sqlLongVarChar
	//TBBookingStatusValue = objControl.Text
	//Case Else
	//TBBookingStatusValue = ""
	//End Select)

	//ElseIf TypeOf objControl Is COAInt_OptionGroup Then
	//' Character field from an option group (CHAR type column). Save the text from the option group.
	//TBBookingStatusValue = objControl.Text
	//End If

	//Exit For
	//End If
	//End If
	//Next objControl
	//Set objControl = Nothing
	//End If


}

function ExecutePostSaveCode() {
	//TODO:

}

function linkButtonClick(lngLinkTableID, lngLinkOrderID, lngLinkViewID) {
	//Get the ID of the linked table.
	var lngLinkRecordID = 0;
	var ubound = (window.mavIDColumns.length);

	for (var iLoop = 0; iLoop < (ubound) ; iLoop++) {

		if (window.mavIDColumns[iLoop][2] == "ID_" + lngLinkTableID) {
			// The given column is an ID column so put the value into the ID column array.
			lngLinkRecordID = window.mavIDColumns[iLoop][3];
			break;
		}
	}
	
	menu_loadLinkPage(lngLinkTableID, lngLinkOrderID, lngLinkViewID, lngLinkRecordID);

}

function formatAddress(addressUrl) {
	if (addressUrl == undefined) return false;

	if ((addressUrl.substr(0, 7).toLowerCase() == 'http://') || (addressUrl.substr(0, 8).toLowerCase() == 'https://')) return addressUrl;

	return 'http://' + addressUrl;

}


function recEdit_ChangedOLEPhoto(plngColumnID, psWhat) {
	//get info about the uploaded item

	switch (psWhat) {
		case "ALL":
			//TODO:
			break;
		case "NONE":
			window.malngChangedOLEPhotos = [];
			break;
		default:
			var fFound = false;
			var ubound = Math.max(0, window.malngChangedOLEPhotos.length - 1);
			for (var i = 0; i <= ubound; i++) {
				if (window.malngChangedOLEPhotos[i] == plngColumnID) {
					fFound = true;
					break;
				}
			}
			
			if (!fFound) {
				window.malngChangedOLEPhotos.push(plngColumnID);
			}
			break;
	}
}
