
//functions that replicate COAIntRecordDMI.ocx

function addControl(tabNumber, controlDef) {

    var tabID = "FI_21_" + tabNumber;
		
    if ($("#" + tabID).length <= 0) {
	    if (tabNumber > 0) {
		    //tab doesn't exist - create it...
		    var tabFontName = $("#txtRecEditFontName").val();
		    var tabFontSize = $("#txtRecEditFontSize ").val();

	    	//var tabCss = "style='font-family: " + tabFontName + " ; font-size: " + tabFontSize + "pt'";
		    var tabCss = "";

		    var tabs = $("#ctlRecordEdit").tabs(),
			    tabTemplate = "<li><a " + tabCss + " href='#{href}'>#{label}</a></li>";

		    var label = getTabCaption(tabNumber),
			    li = $(tabTemplate.replace(/#\{href\}/g, "#" + tabID).replace(/#\{label\}/g, label));

		    tabs.find(".ui-tabs-nav").append(li);
		    tabs.append("<div style='position: relative;' id='" + tabID + "'></div>");
		    tabs.tabs("refresh");
		    if (tabNumber == 1) tabs.tabs("option", "active", 0);
	    } else {
	    	$("#ctlRecordEdit").append("<div style='position: relative;' id='" + tabID + "'></div>");
	    	$("#ctlRecordEdit").css("background-color", "white");
		    $("#ctlRecordEdit").css("border", "1px solid gray");
	    }
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

	$('input[id^="txtRecEditControl_"]').each(function(index) {
		var objScreenControl = getScreenControl_Collection($(this).val());

		//NB we don#t use .Tag any more. Removed as we can grab controls from the txtRecordEditControl_ list instead.
		
		if ((objScreenControl.ColumnID > 0) &&
			(objScreenControl.SelectGranted)) {

			var sDefaultValue = objScreenControl.DefaultValue;
			var lngColumnID = objScreenControl.ColumnID;
			
			//use updatecontrol???
			if((sDefaultValue != null) && (sDefaultValue != undefined)) 
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
			(objScreenControl.TableID = $("#txtRecEditTableID").val()) &&
			(objScreenControl.SelectGranted) &&
			(objScreenControl.DfltValueExprID > 0)) {
			
			sColsAndCalcs += (sColsAndCalcs.length > 0?",":"") + objScreenControl.ColumnID;
		}
	});

	return sColsAndCalcs;
}


function insertUpdateDef() {
	

// Adapted from recordDMI.ocx.
//	Return the SQL string for inserting/updating the current record.
	var  fFound = false;
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

			//		'JPD 20040706 Fault 8884
			fDoControl = (!$(objControl).is(":disabled")); //objScreenControl.enabled;

			if (fDoControl) {
				if((objScreenControl.ControlType == 64) && (objScreenControl.Multiline)) {	//tdbtextctl.tdbtext
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
				if (objScreenControl.ControlType == 2048) { //command button. TODO: should this include nav control?
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
					//TODO: ready to go when we have a malngChangedOLEPhotos array....
					if (objScreenControl.ControlType == 1024) fDoControl = false;
					//if ((objScreenControl.ControlType == 4) || (objScreenControl.ControlType == 8)) { //coa_image or coaint_OLE
					//	fFound = false;
					//	for (iLoop2 = 1; iLoop2 <= malngChangedOLEPhotos.length; iLoop2++) {
					//		if (malngChangedOLEPhotos(iLoop2) == objScreenControl.ColumnID) {
					//			fFound = true;
					//			break;
					//		}
					//	}
					//	if (!fFound) {
					//		fColumnDone = true;
					//	}
					//}
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
							asColumnsToAdd[1] = "'" + ConvertData($(objControl).val(), objScreenControl.DataType).replace("'", "''") + "'";
							//JPD 20051121 Fault 10583
							//asColumns(4, iNextIndex) = ConvertData(objControl.Text, mobjScreenControls.Item(sTag).DataType)
							asColumnsToAdd[3] = ConvertData($(objControl).val(), objScreenControl.DataType).replace("\t", " ");
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
							asColumnsToAdd[1] = "'" + $(objControl).val().replace("'", "''") + "'";
							//	JPD 20051121 Fault 10583
							//	'asColumns(4, iNextIndex) = objControl.Text
							asColumnsToAdd[3] = $(objControl).val().replace("\t", " ");
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
						asColumnsToAdd[1] = "'" + $(objControl).val().replace("'", "''") + "'";
						//	JPD 20051121 Fault 10583
						//	asColumns(4, iNextIndex) = objControl.Text
						asColumnsToAdd[3] = $(objControl).val().replace("\t", " ");
					}
					
					else if ((objScreenControl.ControlType == 64) && 
						((objScreenControl.DataType == 2) || (objScreenControl.DataType==4))) {
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
					
					else if (objScreenControl.ControlType == 2) {
												
						//	Character field from a combo (CHAR type column). Save the text from the combo.
						asColumnsToAdd[1] = "'" + $(objControl).val().replace("'", "''") + "'";
						//	'JPD 20051121 Fault 10583
						//	'asColumns(4, iNextIndex) = objControl.Text
						asColumnsToAdd[3] = $(objControl).val().replace("\t", " ");
					}
					
					else if (objScreenControl.ControlType == 2)  { //&& (objScreenControl.ColumnType == 1))
						// objControl Is COAInt_Lookup Then
						//	Lookup field from a combo (unknown type column). Get the column type and save the appropraite value from the combo.
						switch (objScreenControl.DataType) {
						case 12, -1:
							asColumnsToAdd[1] = "'" + $(objControl).val().replace("'", "''") + "'";
						//	JPD 20051121 Fault 10583
						//	asColumns(4, iNextIndex) = objControl.Text
							asColumnsToAdd[3] = $(objControl).val().replace("\t", " ");
							break;
						case 2, 4:
							if ($(objControl).val(), length > 0) {
								asColumnsToAdd[1] = $(objControl).val();
								//	'TM20070328 - Fault 12053
								asColumnsToAdd[1] = ConvertData($(objControl).val(), objScreenControl.DataType);
								//TODO: remove the next line and fetch value from 'somewhere'...
								var msLocaleThousandSeparator = ",";
								asColumnsToAdd[1] = ConvertNumberForSQL(asColumns(2, iNextIndex)).replace(msLocaleThousandSeparator, "");
								asColumnsToAdd[3] = ConvertNumberForSQL(asColumns(2, iNextIndex)).replace(msLocaleThousandSeparator, "");
							} else {
								asColumnsToAdd[1] = "null";
								asColumnsToAdd[3] = "null";
							}
							break;
						case 11:
							if ($(objControl).val().length > 0) {
								//asColumnsToAdd[1] = "'" + Replace(Format(ConvertData(objControl.Text, mobjScreenControls.Item(sTag).DataType), "mm/dd/yyyy"), msLocaleDateSeparator, "/") & "'"
								asColumnsToAdd[1] = "'" + OpenHR.convertLocaleDateToSQL($(objControl).val()) + "'"; 
								//	'JPD 20051121 Fault 10583
								//	'asColumns(4, iNextIndex) = Replace(Format(ConvertData(objControl.Text, mobjScreenControls.Item(sTag).DataType), "mm/dd/yyyy"), msLocaleDateSeparator, "/")
								// asColumns(4, iNextIndex) = Replace(Replace(Format(ConvertData(objControl.Text, mobjScreenControls.Item(sTag).DataType), "mm/dd/yyyy"), msLocaleDateSeparator, "/"), vbTab, " ")
								asColumnsToAdd[4] = OpenHR.convertLocaleDateToSQL($(objControl).val()).replace("\t", " ");
							} else {
								asColumnsToAdd[1] = "null";
								asColumnsToAdd[3] = "null";
							}
							break;
						default :
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
							asColumnsToAdd[1] = "'" + optionSelected.replace("'", "''") + "'";
							//asColumnsToAdd[1] = "'" + $(objControl).val().replace("'", "''") + "'";
							//	'JPD 20051121 Fault 10583
							//	'asColumns(4, iNextIndex) = objControl.Text
							asColumnsToAdd[3] = optionSelected.replace("\t", " ");
						}
					}
					
					else if (objScreenControl.ControlType == 8) {
						fDoControl = false;
						//TODO: OLE stuff....
						//TypeOf objControl Is COAInt_OLE Then
						//	' OLE field (CHAR type column). Save the name of the OLE file.
						//if Len(objControl.FileName) > 0 And objControl.OLEType < 2 Then
						//	asColumns(2, iNextIndex) = "'" & Replace(Mid(objControl.FileName, InStrRev(objControl.FileName, "\") + 1), "'", "''") & "'"
						//	asColumns(4, iNextIndex) = CStr(objControl.OLEType) & Mid(objControl.FileName, InStrRev(objControl.FileName, "\") + 1)
						//ElseIf objControl.OLEType < 2 Then
						//	asColumns(2, iNextIndex) = "''"
						//	asColumns(4, iNextIndex) = CStr(objControl.OLEType)
						//	Else
						//	asColumns(2, iNextIndex) = asColumns(1, iNextIndex)
						//	asColumns(4, iNextIndex) = CStr(objControl.OLEType) & asColumns(1, iNextIndex)
						//	bCopyImageDataType = True

						//	End If
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
						//TypeOf objControl Is TDBDate6Ctl.TDBDate Then
						//	' Date field from a date control (DATETIME type column). Save the value from the control formatted as 'mm/dd/yyyy' for SQL.
						if (ConvertData($(objControl).val(), objScreenControl.DataType) == null) {
							asColumnsToAdd[1] = "null";
							asColumnsToAdd[3] = "null";
						} else {
							//	asColumns(2, iNextIndex) = "'" & Replace(Format(ConvertData(objControl.Text, mobjScreenControls.Item(sTag).DataType), "mm/dd/yyyy"), msLocaleDateSeparator, "/") & "'"
							asColumnsToAdd[1] = "'" + OpenHR.convertLocaleDateToSQL($(objControl).val()) + "'";
							//	'JPD 20051121 Fault 10583
							//	'asColumns(4, iNextIndex) = Replace(Format(ConvertData(objControl.Text, mobjScreenControls.Item(sTag).DataType), "mm/dd/yyyy"), msLocaleDateSeparator, "/")
							//	asColumns(4, iNextIndex) = Replace(Replace(Format(ConvertData(objControl.Text, mobjScreenControls.Item(sTag).DataType), "mm/dd/yyyy"), msLocaleDateSeparator, "/"), vbTab, " ")
							asColumnsToAdd[3] = OpenHR.convertLocaleDateToSQL($(objControl).val()).replace("\t", " ");
						}
					}

					else if (objScreenControl.ControlType == 4096) {
						fDoControl = false;
						//TODO: Working patterns...
						//TypeOf objControl Is COA_WorkingPattern Then
						//	' Working Pattern Field (CHAR type column, len 14).
						//	asColumns(2, iNextIndex) = "'" & Replace(objControl.Value, "'", "''") & "'"
						//	'JPD 20051121 Fault 10583
						//	'asColumns(4, iNextIndex) = objControl.Value
						//	asColumns(4, iNextIndex) = Replace(objControl.Value, vbTab, " ")
					}
					
					else if (objScreenControl.ControlType == Math.pow(2, 15)) {
						fDoControl = false;
						//TODO: Colour picker...
						//TypeOf objControl Is COA_ColourSelector Then
						//	asColumns(2, iNextIndex) = Val(objControl.BackColor)
						//	asColumns(4, iNextIndex) = Val(objControl.BackColor)
					}
					
					//TODO: else if(objScreenControl.ControlType == ??????) TypeOf objControl Is CommandButton Then
					//	If mobjScreenControls.Item(sTag).LinkTableID <> mlngParentTableID Then
					//	For iLoop = 1 To UBound(mavIDColumns, 2)
					//	If UCase(mavIDColumns(2, iLoop)) = "ID_" & Trim(Str(mobjScreenControls.Item(sTag).LinkTableID)) Then
					//	asColumns(2, iNextIndex) = Trim(Str(mavIDColumns(3, iLoop)))
					//	asColumns(4, iNextIndex) = Trim(Str(mavIDColumns(3, iLoop)))
					//	Exit For
					//	End If
					//	Next iLoop
					//	Else
					//	asColumns(2, iNextIndex) = Trim(Str(mlngParentRecordID))
					//	asColumns(4, iNextIndex) = Trim(Str(mlngParentRecordID))
					//	End If
					//	End If
					//	End If
				}
			}
			//	End If
			
			if((fDoControl) && (!fColumnDone)) asColumns.push(asColumnsToAdd);

		}

		
	});
	
//	Next objControl
//	Set objControl = Nothing

//	See if we are a history screen and if we are save away the id of the parent also
	if($("txtCurrentParentTableID").val() > 0)
	{
		//	Check if the column's update string has already been constructed.
		fColumnDone = false;
		var ubound = asColumns.length -1;
		for (iNextIndex = 0; iNextIndex <= ubound; iNextIndex++) {
			if (asColumns[iNextIndex][0] == "ID_" + $("#txtCurrentParentTableID").val()) {
				fColumnDone = true;
				break;
			}
		}

		if (!fColumnDone) {
			//	Add the column name to the array of columns that have already been entered in the
			//	SQL update/insert string.
			iNextIndex = asColumns.length + 1; //TODO: check this...			
			asColumnsToAdd[iNextIndex][0] = "ID_" + $.trim($("#txtCurrentParentTableID").val());
			asColumnsToAdd[iNextIndex][1] = $.trim($("#txtCurrentParentRecordID").val());
			asColumnsToAdd[iNextIndex][2] = "ID_" + $.trim($("#txtCurrentParentTableID").val());
			asColumnsToAdd[iNextIndex][3] = $.trim($("#txtCurrentParentRecordID").val());
			
			asColumns.push(asColumnsToAdd);
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

	if (console) {
		console.log(sInsertUpdateDef);
	}
	
	return sInsertUpdateDef;

}


function ConvertNumberForSQL(strInput) {
	// Get a number in the correct format for a SQL string
	// (e.g. on french systems replace decimal comma for a decimal point)
	// TODO: return strInput.replace(msLocaleDecimalSeparator, ".");
	return strInput;

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
			case 12, -1:
				//sqlVarChar, sqlLongVarChar
				vReturnData = (pvData).replace(/~+$/, ''); //rtrim function
				if (vReturnData.length == 0) {
					vReturnData = null;
				}
				break;
			case 11:
				//sqlDate				
				if (isValidDate(new Date(pvData))) {
					vReturnData =Date.parse(pvData);
				} else {
					vReturnData = null;
				}
				break;
			case 2:
				//sqlNumeric
				if ($.trim(pvData).length == 0) {
					vReturnData = null;
				} else {
					vReturnData = Number(pvData);
				}
				break;
			case 4:
				// sqlInteger
				if ($.trim(pvData).length == 0) {

					vReturnData = null;
				} else {
					vReturnData = Number(pvData);
				}
				break;

			default:
				vReturnData = pvData;
		}
	}

return vReturnData;

}

	function isValidDate(d) {
		if ( Object.prototype.toString.call(d) !== "[object Date]" )
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
						psText = Left(psText, lngPos - 1) + psText.substr(lngPos -1, 1).toUpperCase() + Right(psText, psText.length - lngPos);
						//psText = psText.substr(0, lngPos - 1) + psText.substr(lngPos, 1).toUpperCase() + Right(psText, psText.length - lngPos);
					} else if (lngPos > 2) {
						// Catch the McName.
						if ((psText.substr(lngPos - 2, 1) == "M") && (sLastCharacter = "c")) {
							psText = Left(psText, lngPos - 1) & psText.substr(lngPos -1, 1).toUpperCase() & Right(psText, psText.length - lngPos);
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

function Right(str, n){
	if (n <= 0)
		return "";
	else if (n > String(str).length)
		return str;
	else {
		var iLen = String(str).length;
		return String(str).substring(iLen, iLen - n);
	}
}

function allDefaults()
{
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

		var ubound = (this.mavIDColumns.length / 3);
		
		for (iLoop = 1; iLoop <= (ubound); iLoop++) {

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
			OpenHR.messageBox("Unable to save record, a link must be made with the parent table.", vbExclamation + vbOKOnly, "OpenHR Intranet");
		}
	}

	return fValid;		
}


function getScreenControl_Collection(screenControlValue) {

	var controlItemArray = screenControlValue.split("\t");

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
			DefaultValue: controlItemArray[24],
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
			NavigateTo: controlItemArray[58],
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
        var nextAvail;
        if (this.mavIDColumns.length <= 0) {
            nextAvail = 0;
        }
        else {
            nextAvail = this.mavIDColumns.length / 3;
        }

        this.mavIDColumns[nextAvail] = new Array(3);

        this.mavIDColumns[nextAvail][1] = Number(controlItemArray[1]);   //ColumnID
        this.mavIDColumns[nextAvail][2] = controlItemArray[2];   //Column Name
        this.mavIDColumns[nextAvail][3] = 0;   //Value

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
                fControlEnabled = (Number(controlItemArray[48] != 0));  // UpdateGranted property
            }
            else if (controlType == 2048) {
                //CommandButton
                fControlEnabled = false;
            }
            else {
                fControlEnabled = (Number(controlItemArray[48] != 0));  // UpdateGranted property

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
    
    switch (Number(controlItemArray[3])) {
        case 1: //checkbox
            span = document.createElement('span');
            
            applyLocation(span, controlItemArray, true);
            span.style.margin = "0px";
            span.style.textAlign = "left";
            span.style.display = "inline-block";

            var checkbox = span.appendChild(document.createElement('input'));
            checkbox.type = "checkbox";
            checkbox.id = controlID;
            checkbox.style.fontFamily = controlItemArray[11];
            checkbox.style.fontSize = controlItemArray[12] + 'pt';
            checkbox.style.position = "absolute";
            checkbox.style.top = "50%";
            
            checkbox.style.padding = "0px";
            checkbox.style.margin = "-7px 0px 0px 0px";
            checkbox.style.textAlign = "left";

            var label = span.appendChild(document.createElement('label'));
            label.htmlFor = checkbox.id;
            label.appendChild(document.createTextNode(controlItemArray[8]));
            
            label.style.fontFamily = controlItemArray[11];
            label.style.fontSize = controlItemArray[12] + 'pt';

            //align left or right...
            if (controlItemArray[20] != "0") {            
                //right align
                span.className = "checkbox right";
                checkbox.style.right = "0px";
            } else {
                //left align
                span.className = "checkbox left";
                checkbox.style.left = "0px";
                label.style.marginLeft = "18px";
            }

            if(tabIndex > 0) checkbox.tabindex = tabIndex;

            checkbox.setAttribute("data-columnID", columnID);
            checkbox.setAttribute("data-control-tag", key);

            if (!fControlEnabled) span.disabled = true;

            //Add control to relevant tab, create if required.                
            addControl(iPageNo, span);

            break;
    	case 2: //ctlCombo


            var selector = document.createElement('select');
            selector.id = controlID;
            applyLocation(selector, controlItemArray, true);
            selector.style.backgroundColor = "White";
            selector.style.color = "Black";
            selector.style.fontFamily = controlItemArray[11];
            selector.style.fontSize = controlItemArray[12] + 'pt';
            selector.style.borderWidth = "1px";
            selector.setAttribute("data-columnID", columnID);
            selector.setAttribute("data-control-key", key);
            if (controlItemArray[22] == 1) {
            	//column type = ---- LOOKUPS ----
            	selector.setAttribute("data-columntype", "lookup");
            	//plngColumnID, plngLookupColumnID, psLookupValue, pfMandatory, pstrFilterValue
            	selector.setAttribute("data-LookupTableID", controlItemArray[27]);
            	selector.setAttribute("data-LookupColumnID", controlItemArray[28]);
            	selector.setAttribute("data-LookupFilterColumnID", controlItemArray[53]);
            	selector.setAttribute("data-LookupFilterValueID", controlItemArray[54]);
            	selector.setAttribute("data-Mandatory", controlItemArray[32]);
            }

            if (!fControlEnabled) selector.disabled = true;

            if (tabIndex > 0) selector.tabindex = tabIndex;

            addControl(iPageNo, selector);

            //var option = document.createElement('option');
            //option.value = '0';
            //option.appendChild(document.createTextNode(''));
            //selector.appendChild(option);

            break;

        case 4: //Image
            var image = document.createElement('img');
            image.id = controlID;
            applyLocation(image, controlItemArray, true);
            image.style.border = "1px solid gray";
            image.style.padding = "0px";
            image.setAttribute("data-columnID", columnID);
            image.setAttribute("data-control-key", key);
            
            if (!fControlEnabled) image.disabled = true;

            var path = window.ROOT + 'Home/ShowImageFromDb?imageID=' + controlItemArray[50];
            
            image.setAttribute('src', path);            

            //Add control to relevant tab, create if required.                
            addControl(iPageNo, image);

            break;
        case 8: //ctlOle
            var button = document.createElement('input');
            button.type = "button";
            button.id = controlID;
            button.value = "OLE";
            applyLocation(button, controlItemArray, true);
            button.style.padding = "0px";
            button.setAttribute("data-columnID", columnID);
            button.setAttribute("data-control-key", key);
            
            if (tabIndex > 0) button.tabindex = tabIndex;

            //button.disabled = false;    //always enabled
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
            fieldset.id = controlID;
            fieldset.setAttribute("data-datatype", "Option Group");
            fieldset.setAttribute("data-columnID", columnID);
            fieldset.setAttribute("data-alignment", controlItemArray[20]);
            fieldset.setAttribute("data-control-key", key);
            
            if ((controlItemArray[19] != "0") && (controlItemArray[8].length > 0)) {
                //has a border and a caption
                legend = fieldset.appendChild(document.createElement('legend'));
                legend.style.fontFamily = controlItemArray[11];
                legend.style.fontSize = controlItemArray[12] + 'pt';
                legend.style.fontWeight = (Number(controlItemArray[13]) != 0) ? "bold" : "normal";
                legend.appendChild(document.createTextNode(controlItemArray[8]));
            }

            
            //No Option Group buttons - these are added as values next.

            addControl(iPageNo, fieldset);           


            break;
        case 32: //ctlSpinner
            var spinnerContainer = document.createElement('div');
            applyLocation(spinnerContainer, controlItemArray, true);
            spinnerContainer.style.padding = "0px";

            var spinner = spinnerContainer.appendChild(document.createElement("input"));
            spinner.className = "spinner";
            spinner.id = controlID;
            spinner.style.fontFamily = controlItemArray[11];
            spinner.style.fontSize = controlItemArray[12] + 'pt';
            spinner.style.width = (Number((controlItemArray[7]) / 15)) + "px";
            spinner.style.margin = "0px";
            spinner.setAttribute("data-columnID", columnID);
            spinner.setAttribute("data-control-key", key);
            
            if (tabIndex > 0) spinner.tabindex = tabIndex;
            if (!fControlEnabled) spinnerContainer.disabled = true;

            //Add control to relevant tab, create if required.                
            addControl(iPageNo, spinnerContainer);
            break;
        case 64: //ctlText
            var textbox;
            if (Number(controlItemArray[37]) !== 0) {
                //Multi-line textbox
                textbox = document.createElement('textarea'); //textbox.disabled = false;  //always enabled.
            } else {
                textbox = document.createElement('input');
                switch (Number(controlItemArray[23])) {
                    case 11: //sqlDate
                        textbox.type = "text";
                        textbox.className = "datepicker";
                        break;
                    case 2, 4: //sqlNumeric, sqlInteger
                        textbox.className = "number";
                        break;
                    default:
                        textbox.type = "text";
                        textbox.isMultiLine = false;

                        if (controlItemArray[35].length > 0) {
                            //TODO: apply mask to control
                        }
                }

                if (!fControlEnabled) textbox.disabled = true;

            }

            textbox.id = controlID;
            applyLocation(textbox, controlItemArray, true);
            textbox.style.fontFamily = controlItemArray[11];
            textbox.style.fontSize = controlItemArray[12] + 'pt';
            textbox.style.padding = "0px";
            textbox.setAttribute("data-columnID", columnID);
            textbox.setAttribute("data-control-key", key);
            
            if (tabIndex > 0) textbox.tabindex = tabIndex;
            
            //Add control to relevant tab, create if required.                
            addControl(iPageNo, textbox);
            break;
        case 128: //ctlTab
            break;
        case 256: //Label
            span = document.createElement('span');
            applyLocation(span, controlItemArray, false);
            span.style.backgroundColor = "transparent";
            //span.style.color = "Black";
            span.style.fontFamily = controlItemArray[11];
            span.style.fontSize = controlItemArray[12] + 'pt';
            span.textContent = controlItemArray[8];

            span.setAttribute("data-control-key", key);
            
            //replaces the SetControlLevel function in recordDMI.ocx.
            span.style.zIndex = 0;

            if (!fControlEnabled) span.disabled = true;

            addControl(iPageNo, span);

            break;
        case 512: //Frame
            var fieldset = document.createElement('fieldset');
            applyLocation(fieldset, controlItemArray, true);
            fieldset.style.backgroundColor = "transparent";
            //fieldset.style.color = "Black";
            fieldset.style.padding = "0px";
            
            var legend = fieldset.appendChild(document.createElement('legend'));
            legend.style.fontFamily = controlItemArray[11];
            legend.style.fontSize = controlItemArray[12] + 'pt';
            legend.style.fontWeight = (Number(controlItemArray[13]) != 0) ? "bold" : "normal";
            legend.style.textDecoration = (Number(controlItemArray[16]) != 0) ? "underline" : "none";
            legend.appendChild(document.createTextNode(controlItemArray[8]));

            fieldset.setAttribute("data-control-key", key);
            
            addControl(iPageNo, fieldset);

            break;
        case 1024: //ctlPhoto
            


        case 2048: //ctlCommand
            break;
        case 4096: //ctlWorking Pattern
            //TODO: Font size change - this control is fixed in size.
            top = (Number(controlItemArray[4]) / 15);
            left = (Number(controlItemArray[5]) / 15);
            height = 58; //(Number((controlItemArray[6]) / 15) - 2);
            width = 125; //(Number((controlItemArray[7]) / 15) - 2);
            if (controlItemArray[19] == "0") {
                //pictureborder?
                borderCss = "border-style: none;";
            } else {
                borderCss = "border: 1px solid #999;";
                width -= 2;
                height -= 2;

                //TODO ??  fontadjustment?

                //TODO - android browser/tablet adjustment
            }

            fieldset = document.createElement("fieldset");
            fieldset.id = controlID;
            fieldset.setAttribute("data-columnID", columnID);
            fieldset.setAttribute("data-datatype", "Working Pattern");
            fieldset.setAttribute("data-control-key", key);
            
            fieldset.style.position = "absolute";
            fieldset.style.top = top + "px";
            fieldset.style.left = left + "px";
            fieldset.style.width = width + "px";
            fieldset.style.height = height + "px";
            fieldset.style.padding = "0px";
            fieldset.style.border = borderCss;

            for (var i = 0; i < 7; i++) {
                var offsetLeft = 26 + (i * 13);
                var dayLabel = fieldset.appendChild(document.createElement("span"));
                switch (i) {
                    case 0:
                        dayLabel.textContent  = "S";
                        break;
                    case 1:
                        dayLabel.textContent  = "M";
                        break;
                    case 2:
                        dayLabel.textContent  = "T";
                        break;
                    case 3:
                        dayLabel.textContent  = "W";
                        break;
                    case 4:
                        dayLabel.textContent  = "T";
                        break;
                    case 5:
                        dayLabel.textContent  = "F";
                        break;
                    case 6:
                        dayLabel.textContent  = "S";
                        break;
                }
                
                //Day labels
                dayLabel.style.fontFamily = controlItemArray[11];
                dayLabel.style.fontSize = controlItemArray[12] + 'pt';
                dayLabel.style.fontWeight = (Number(controlItemArray[13]) != 0) ? "bold" : "normal";
                dayLabel.style.position = "absolute";
                dayLabel.style.top = "6px";
                dayLabel.style.left = offsetLeft + 3 + "px";

                //AM Boxes
                var amCheckbox = fieldset.appendChild(document.createElement("input"));
                amCheckbox.type = "checkbox";
                amCheckbox.id = controlID + "_" + ((i * 2) + 1);
                amCheckbox.style.padding = "0px";
                amCheckbox.style.position = "absolute";
                amCheckbox.style.top = "22px";
                amCheckbox.style.left = offsetLeft + "px";
                if (!fControlEnabled) amCheckbox.disabled = true;
                
                //PM Boxes
                var pmCheckbox = fieldset.appendChild(document.createElement("input"));
                pmCheckbox.type = "checkbox";
                pmCheckbox.id = controlID + "_" + ((i * 2) + 2);
                pmCheckbox.style.padding = "0px";
                pmCheckbox.style.position = "absolute";
                pmCheckbox.style.top ="36px";
                pmCheckbox.style.left = offsetLeft + "px";
                if (!fControlEnabled) pmCheckbox.disabled = true;
            }

            //AM/PM Labels
            label = document.createElement("span");
            label.textContent  = "AM";
            label.style.fontFamily = controlItemArray[11];
            label.style.fontSize = controlItemArray[12] + 'pt';
            label.style.fontWeight = (Number(controlItemArray[13]) != 0) ? "bold" : "normal";
            label.style.position = "absolute";
            label.style.top = top + 22 + "px";
            label.style.left = left + 4 + "px";
            addControl(iPageNo, label);

            label = document.createElement("span");
            label.textContent  = "PM";
            label.style.fontFamily = controlItemArray[11];
            label.style.fontSize = controlItemArray[12] + 'pt';
            label.style.fontWeight = (Number(controlItemArray[13]) != 0) ? "bold" : "normal";
            label.style.position = "absolute";
            label.style.top = top + 36 + "px";
            label.style.left = left + 4 + "px";
            addControl(iPageNo, label);

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
                    hyperlink.style.color = controlItemArray[10];
                    hyperlink.style.backgroundColor = controlItemArray[9];                    
                    hyperlink.setAttribute("href", controlItemArray[58]);
                    hyperlink.setAttribute("target", "_blank");

                    hyperlink.id = controlID;
                    hyperlink.style.padding = "0px";
                    hyperlink.setAttribute("data-columnID", columnID);
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

                    button.setAttribute("onclick", "window.open('" + controlItemArray[58] + "')");                                        
                    
                    button.id = controlID;
                    button.style.padding = "0px";
                    button.setAttribute("data-columnID", columnID);
                    button.setAttribute("data-control-key", key);

                    if (tabIndex > 0) button.tabindex = tabIndex;

                    addControl(iPageNo, button);

                    break;
                case 2: //Browser
                    var el = document.createElement("iframe");
                    applyLocation(el, controlItemArray, true);
                    el.id = controlID;
                    el.setAttribute("data-columnID", columnID);
                    el.setAttribute("data-control-key", key);
                    if (tabIndex > 0) el.tabindex = tabIndex;

                    addControl(iPageNo, el);
                    el.setAttribute('src', controlItemArray[58]);

                    break;
                case 3: //Hidden
                    break;
                    
            }

            //TODO: Nav control always .disabled = false.
            //if (tabIndex > 0) checkbox.tabindex = tabIndex;
            break;
        case 2 ^ 15: //ctlColourPicker
            //TODO: .disabled = (!fControlEnabled);
            //if(tabIndex > 0) checkbox.tabindex = tabIndex;
            break;
        default:
            break;
    }
}

function addHTMLControlValues(controlValues) {
    
    try {
        var controlValuesArray = controlValues.split("\t");
    } catch(e) {return false;}

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

                            //add text to radio button
                            label.style.position = "absolute";
                            label.style.top = (i * 16) + "px";
                            label.style.left = "28px";
                            label.style.padding = "0px";
                            label.htmlFor = uniqueID + "_" + i;
                            label.appendChild(document.createTextNode(sValue));
                        }
                        if (alignment == 1) {
                            $(this).css("padding-left", "17px");
                            //Horizontal alignment
                            radio.style.padding = "0px";

                            //add text to radio button
                            label.style.marginLeft = "3px";
                            label.style.marginRight = "32px";
                            label.htmlFor = uniqueID + "_" + i;
                            
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
    var fIsIDColumn = false;

    if (columnID.toUpperCase() == "TIMESTAMP") {
        // The column is the timestamp column.
	    $("#txtRecEditTimeStamp").val(value);
    }
    else {
        var tmp = this.mavIDColumns.indexOf(Number(columnID));
        if (tmp > 0) {
            this.mavIDColumns[tmp][3] = Number(value);
            fIsIDColumn = true;
        }

        if (!fIsIDColumn) {
        	if ((value != null) && (value != undefined))
            updateControl(Number(columnID), value);
        }
    }
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

		//TODO if coa_image.....
		//If InStr(1, CStr(pvValue), "::LINKED_OLE_DOCUMENT::", vbTextCompare) Then
		//.SetPicturePath Replace(CStr(pvValue), "::LINKED_OLE_DOCUMENT::", "")
		//.ASRDataField = GetFileNameOnly(Replace(CStr(pvValue), "::LINKED_OLE_DOCUMENT::", ""))
		//.OLEType = 3
	 // ElseIf InStr(1, CStr(pvValue), "::EMBEDDED_OLE_DOCUMENT::", vbTextCompare) Then
	 //.SetPicturePath Replace(CStr(pvValue), "::EMBEDDED_OLE_DOCUMENT::", "")
	 //.ASRDataField = GetFileNameOnly(Replace(CStr(pvValue), "::EMBEDDED_OLE_DOCUMENT::", ""))
	 //.OLEType = 2
	 // Else
	 // 	.SetPicturePath msPhotoPath & "\" & CStr(pvValue)
	 // 	.ASRDataField = CStr(pvValue)
	 // 	.OLEType = mobjScreenControls.Item(sTag).OLEType
		// End If
		
		//TODO if mask
		// .Text = RTrim(CStr(pvValue) & vbNullString)
		
		//Input type controls...
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
					} else {
						$(this).val(value);
					}
					break;
				case "number":
					$(this).val(Number(value));
					break;
				case "checkbox":
					$(this).prop("checked", value == "True" ? true : false);
					break;
				case "button":
					//link button or nav control
					if (controlType == Math.pow(2, 14)) {
						//Navigation Control
						if (value.length <= 0) {
							$(this).attr("href", "about:blank");
						} else {
							$(this).attr("href", value);
						}
					}
					//TODO: link button.
				default:
					$(this).val(value);

			}
		}

		//Working pattern & Option group
		if ($(this).is("fieldset")) {
			if ($(this).attr("data-datatype") === "Working Pattern") {
				//ensure the value is 14 characters long.
				if (value.length < 14) value = value.concat("              ").substring(0, 14);
				var tthisId = "#" + $(this).attr("id");
				//tick relevant boxes.
				for (var i = 1; i <= 14; i++) {
					$(tthisId + "_" + i).prop("checked", value.substring(i - 1, i) != " " ? true : false);
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
					$(this).attr("href", value);
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

    var tabCaption = arr[tabNumber - 1];

    return tabCaption;

}



function TBCourseRecordID() {
	// Training Booking specific.
	// Return the Course Record ID.
	// Used when editing a Training Booking record.
	var iLoop;

	var TBCourseRecordID = 0;
	var mlngCourseTableID = $("#txtRecEditCourseTableID").val();
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
	var mlngEmployeeTableID = $("#txtRecEditEmpTableID").val();
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
	//End Select

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