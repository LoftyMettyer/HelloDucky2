<%@ control language="VB" inherits="System.Web.Mvc.ViewUserControl" %>
<%@ import namespace="DMI.NET" %>

<object
	classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
	id="Microsoft_Licensed_Class_Manager_1_0"
	viewastext>
	<param name="LPKPath" value="lpks/main.lpk">
</object>

<script type="text/javascript">
	function filterselect_window_onload() {
		var fOK;
		var fFilterOK;
		var sAddString;
		var sColumnName;
		var iColumnID;
		var sOperatorName;
		var iOperatorID;
		var sValue;
		var iIndex;
		var fFound;
		var iLoop;
		var iColumnType;
		var sReqdControlName;
		var sControlName;
		
		var frmFilterForm = document.getElementById("frmFilterForm");

		var controlCollection = frmFilterForm.elements;

		fOK = true;

		var sErrMsg = frmFilterForm.txtErrorDescription.value;
		if (sErrMsg.length > 0) {
			fOK = false;
			OpenHR.messageBox(sErrMsg);
			window.parent.location.replace("login");
		}

		if (fOK == true) {
			setGridFont(frmFilterForm.ssOleDBGridFilterRecords);

			// Expand the option frame and hide the work frame.
			//window.parent.document.all.item("workframeset").cols = "0, *";
			$("#optionframe").attr("data-framesource", "SELECTFILTER");
			$("#workframe").hide();
			$("#optionframe").show();

			frmFilterForm.selectColumn.focus();

			// Add the existing filter definition to the grid. Select the first row.
			fFilterOK = true;
			var sFilterDef = frmFilterForm.txtOptionFilterDef.value;

			while (sFilterDef.length > 0) {
				iIndex = sFilterDef.indexOf("	");
				if (iIndex < 0) {
					fFilterOK = false;
				}
				if (fFilterOK == true) {
					// Get the column ID from the filter definition string.
					iColumnID = sFilterDef.substr(0, iIndex);
					sFilterDef = sFilterDef.substr(iIndex + 1);

					iIndex = sFilterDef.indexOf("	");
					if (iIndex < 0) {
						fFilterOK = false;
					}
				}
				if (fFilterOK == true) {
					// Get the operator ID from the filter definition string.
					iOperatorID = sFilterDef.substr(0, iIndex);
					sFilterDef = sFilterDef.substr(iIndex + 1);

					iIndex = sFilterDef.indexOf("	");
					if (iIndex < 0) {
						fFilterOK = false;
					}
				}
				if (fFilterOK == true) {
					// Get the value from the filter definition string.
					sValue = sFilterDef.substr(0, iIndex);

					sFilterDef = sFilterDef.substr(iIndex + 1);
				}
				if (fFilterOK == true) {
					// Get the column name.
					fFound = false;
					for (iLoop = 0; iLoop < frmFilterForm.selectColumn.options.length; iLoop++) {
						if (iColumnID == frmFilterForm.selectColumn.options(iLoop).value) {
							fFound = true;
							sColumnName = frmFilterForm.selectColumn.options(iLoop).text;
						}
					}
					if (fFound == false) {
						fFilterOK = false;
					} else {
						// Get the data type of the column.
						fFilterOK = false;
						sReqdControlName = "txtFilterColumn_";
						sReqdControlName = sReqdControlName.concat(iColumnID);
						if (controlCollection != null) {
							for (var i = 0; i < controlCollection.length; i++) {
								sControlName = controlCollection.item(i).name;
								if (sControlName == sReqdControlName) {
									iColumnType = controlCollection.item(i).value;
									fFilterOK = true;
									break;
								}
							}
						}
					}
				}
				if (fFilterOK == true) {
					// Get the operator name.
					fFound = false;
					if (iOperatorID == 1) {
						fFound = true;
						sOperatorName = "is equal to";
					}
					if (iOperatorID == 2) {
						fFound = true;
						sOperatorName = "is NOT equal to";
					}
					if (iOperatorID == 3) {
						fFound = true;
						if (iColumnType == 11) {
							sOperatorName = "is equal to or before";
						} else {
							sOperatorName = "is less than or equal to";
						}
					}
					if (iOperatorID == 4) {
						fFound = true;
						if (iColumnType == 11) {
							sOperatorName = "is equal to or after";
						} else {
							sOperatorName = "is greater than or equal to";
						}
					}
					if (iOperatorID == 5) {
						fFound = true;
						if (iColumnType == 11) {
							sOperatorName = "after";
						} else {
							sOperatorName = "is greater than";
						}
					}
					if (iOperatorID == 6) {
						fFound = true;
						if (iColumnType == 11) {
							sOperatorName = "before";
						} else {
							sOperatorName = "is less than";
						}
					}
					if (iOperatorID == 7) {
						fFound = true;
						sOperatorName = "contains";
					}
					if (iOperatorID == 8) {
						fFound = true;
						sOperatorName = "does not contain";
					}
					if (fFound == false) {
						fFilterOK = false;
					}
				}
				if (fFilterOK == true) {
					// Add the filter definition to the grid.
					sAddString = sColumnName;
					sAddString = sAddString.concat("	");
					sAddString = sAddString.concat(sOperatorName);
					sAddString = sAddString.concat("	");
					sAddString = sAddString.concat(sValue);
					sAddString = sAddString.concat("	");
					sAddString = sAddString.concat(iColumnID);
					sAddString = sAddString.concat("	");
					sAddString = sAddString.concat(iOperatorID);

					frmFilterForm.ssOleDBGridFilterRecords.AddItem(sAddString);
				}
			}

			// Select the top filter record (if one exists).	
			if (frmFilterForm.ssOleDBGridFilterRecords.rows > 0) {
				frmFilterForm.ssOleDBGridFilterRecords.MoveFirst();
				frmFilterForm.ssOleDBGridFilterRecords.SelBookmarks.Add(frmFilterForm.ssOleDBGridFilterRecords.Bookmark);
			}

			// Get menu.asp to refresh the menu.
			// NPG20100824 Fault HRPRO1065 - leave menus disabled in these modal screens		
			//window.parent.frames("menuframe").refreshMenu();

			// Hide the workframe recedit control. IE6 still displays it.
			var sWorkPage = currentWorkFramePage();
			if (sWorkPage == "RECORDEDIT") {
				//window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "hidden";
			} else {
				if (sWorkPage == "FIND") {
					//window.parent.frames("workframe").document.forms("frmFindForm").ssOleDBGridFindRecords.style.visibility = "hidden";
				}
			}

			frmFilterForm.selectColumn.focus();
			refreshOperatorCombo();
			refreshControls();
		}

	}
</script>

<script language="JavaScript">
	function SelectFilter() {
		var frmFilterForm = document.getElementById("frmFilterForm");
		var sRealSource;
		var sColumnName;
		var sValue;
		var sFilterValue;
		var sFilterDef;
		var sFilterSQL;
		var sSubFilterSQL;
		var iIndex;
		var reSpace;
		var iColumnID;
		var iOperatorID;
		var fOK;
		var iDataType;
		var sReqdControlName;
		var controlCollection = frmFilterForm.elements;
		var sDecimalSeparator;
		var sModifiedFilterValue;
		var sControlName;

		// Create some regular expressions to be used when replacing characters 
		// in the filter string later on.
		sDecimalSeparator = "\\";
		sDecimalSeparator = OpenHR.LocaleDecimalSeparator;
		var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");

		reSpace = / /gi;
		var reApostrophe = new RegExp("\\'", "gi");
		var reStar = new RegExp("\\*", "gi");
		var reQuestion = new RegExp("\\?", "gi");

		sFilterSQL = "";
		sFilterDef = "";
		sRealSource = frmFilterForm.txtRealSource.value;
		sRealSource = sRealSource.concat(".");

		if (frmFilterForm.ssOleDBGridFilterRecords.rows > 0) {
			frmFilterForm.ssOleDBGridFilterRecords.MoveFirst();
		}

		// Loop through the grid records, building the filter code for each record.
		for (iIndex = 1; iIndex <= frmFilterForm.ssOleDBGridFilterRecords.rows; iIndex++) {
			// Get the column name & id and the value used in the filter operation.
			sColumnName = frmFilterForm.ssOleDBGridFilterRecords.Columns(0).value;
			sColumnName = sColumnName.replace(reSpace, "_");
			sColumnName = sRealSource.concat(sColumnName);

			sValue = frmFilterForm.ssOleDBGridFilterRecords.Columns(2).value;
			iColumnID = frmFilterForm.ssOleDBGridFilterRecords.Columns(3).value;
			iOperatorID = frmFilterForm.ssOleDBGridFilterRecords.Columns(4).value;

			fOK = false;
			iDataType = 12;
			sSubFilterSQL = "";

			// Get the data type of the column.
			sReqdControlName = "txtFilterColumn_";
			sReqdControlName = sReqdControlName.concat(iColumnID);
			if (controlCollection != null) {
				for (var i = 0; i < controlCollection.length; i++) {
					sControlName = controlCollection.item(i).name;
					if (sControlName == sReqdControlName) {
						iDataType = controlCollection.item(i).value;
						fOK = true;
						break;
					}
				}
			}

			if (fOK == true) {
				// If we've found the column's data type go ahead and build the filter code.
				if (iDataType == -7) {
					// Logic column (must be the equals operator).	
					sSubFilterSQL = sColumnName.concat(" = ");

					if (sValue.toUpperCase() == "TRUE") {
						sSubFilterSQL = sSubFilterSQL.concat("1");
					}
					else {
						sSubFilterSQL = sSubFilterSQL.concat("0");
					}
				}

				if ((iDataType == 2) || (iDataType == 4)) {
					// Numeric/Integer column.
					// Replace the locale decimal separator with '.' for SQL's benefit.
					sFilterValue = sValue.replace(reDecimalSeparator, ".");

					if (iOperatorID == 1) {
						// Equals.
						sSubFilterSQL = sColumnName.concat(" = ");
						sSubFilterSQL = sSubFilterSQL.concat(sFilterValue);

						if (parseFloat(sValue) == 0) {
							sSubFilterSQL = sSubFilterSQL.concat(" OR ");
							sSubFilterSQL = sSubFilterSQL.concat(sColumnName);
							sSubFilterSQL = sSubFilterSQL.concat(" IS NULL");
						}
					}

					if (iOperatorID == 2) {
						// Not Equal To.
						sSubFilterSQL = sColumnName.concat(" <> ");
						sSubFilterSQL = sSubFilterSQL.concat(sFilterValue);

						if (parseFloat(sValue) == 0) {
							sSubFilterSQL = sSubFilterSQL.concat(" AND ");
							sSubFilterSQL = sSubFilterSQL.concat(sColumnName);
							sSubFilterSQL = sSubFilterSQL.concat(" IS NOT NULL");
						}
					}

					if (iOperatorID == 3) {
						// Less than or Equal To.
						sSubFilterSQL = sColumnName.concat(" <= ");
						sSubFilterSQL = sSubFilterSQL.concat(sFilterValue);

						if (parseFloat(sValue) >= 0) {
							sSubFilterSQL = sSubFilterSQL.concat(" OR ");
							sSubFilterSQL = sSubFilterSQL.concat(sColumnName);
							sSubFilterSQL = sSubFilterSQL.concat(" IS NULL");
						}
					}

					if (iOperatorID == 4) {
						// Greater than or Equal To.
						sSubFilterSQL = sColumnName.concat(" >= ");
						sSubFilterSQL = sSubFilterSQL.concat(sFilterValue);

						if (parseFloat(sValue) <= 0) {
							sSubFilterSQL = sSubFilterSQL.concat(" OR ");
							sSubFilterSQL = sSubFilterSQL.concat(sColumnName);
							sSubFilterSQL = sSubFilterSQL.concat(" IS NULL");
						}
					}

					if (iOperatorID == 5) {
						// Greater than.
						sSubFilterSQL = sColumnName.concat(" > ");
						sSubFilterSQL = sSubFilterSQL.concat(sFilterValue);

						if (parseFloat(sValue) < 0) {
							sSubFilterSQL = sSubFilterSQL.concat(" OR ");
							sSubFilterSQL = sSubFilterSQL.concat(sColumnName);
							sSubFilterSQL = sSubFilterSQL.concat(" IS NULL");
						}
					}

					if (iOperatorID == 6) {
						// Less than.
						sSubFilterSQL = sColumnName.concat(" < ");
						sSubFilterSQL = sSubFilterSQL.concat(sFilterValue);

						if (parseFloat(sValue) > 0) {
							sSubFilterSQL = sSubFilterSQL.concat(" OR ");
							sSubFilterSQL = sSubFilterSQL.concat(sColumnName);
							sSubFilterSQL = sSubFilterSQL.concat(" IS NULL");
						}
					}
				}

				if (iDataType == 11) {
					// Date column.
					if (sValue.length > 0) {
						// Convert the locale date into the SQL format.
						sFilterValue = menu_convertLocaleDateToSQL(sValue);
					}

					if ((sValue.length == 0) || (sFilterValue.length > 0)) {
						// The data is only valid id it is completely empty, or if the
						// convertLocaleDateToSQL function has returned a non-empty string.
						if (iOperatorID == 1) {
							// Equal To.
							if (sValue.length > 0) {
								sSubFilterSQL = sColumnName.concat(" = '");
								sSubFilterSQL = sSubFilterSQL.concat(sFilterValue);
								sSubFilterSQL = sSubFilterSQL.concat("'");
							}
							else {
								sSubFilterSQL = sColumnName.concat(" IS NULL");
							}
						}

						if (iOperatorID == 2) {
							// Not Equal To.
							if (sValue.length > 0) {
								sSubFilterSQL = sColumnName.concat(" <> '");
								sSubFilterSQL = sSubFilterSQL.concat(sFilterValue);
								sSubFilterSQL = sSubFilterSQL.concat("'");
							}
							else {
								sSubFilterSQL = sColumnName.concat(" IS NOT NULL");
							}
						}

						if (iOperatorID == 3) {
							// Less than or Equal To.
							if (sValue.length > 0) {
								sSubFilterSQL = sColumnName.concat(" <= '");
								sSubFilterSQL = sSubFilterSQL.concat(sFilterValue);
								sSubFilterSQL = sSubFilterSQL.concat("' OR ");
								sSubFilterSQL = sSubFilterSQL.concat(sColumnName);
								sSubFilterSQL = sSubFilterSQL.concat(" IS NULL");
							}
							else {
								sSubFilterSQL = sColumnName.concat(" IS NULL");
							}
						}

						if (iOperatorID == 4) {
							// Greater than or Equal To.
							if (sValue.length > 0) {
								sSubFilterSQL = sColumnName.concat(" >= '");
								sSubFilterSQL = sSubFilterSQL.concat(sFilterValue);
								sSubFilterSQL = sSubFilterSQL.concat("'");
							}
							else {
								sSubFilterSQL = sColumnName.concat(" IS NULL OR ");
								sSubFilterSQL = sSubFilterSQL.concat(sColumnName);
								sSubFilterSQL = sSubFilterSQL.concat(" IS NOT NULL");
							}
						}

						if (iOperatorID == 5) {
							// Greater than.
							if (sValue.length > 0) {
								sSubFilterSQL = sColumnName.concat(" > '");
								sSubFilterSQL = sSubFilterSQL.concat(sFilterValue);
								sSubFilterSQL = sSubFilterSQL.concat("'");
							}
							else {
								sSubFilterSQL = sColumnName.concat(" IS NOT NULL");
							}
						}

						if (iOperatorID == 6) {
							// Less than.
							if (sValue.length > 0) {
								sSubFilterSQL = sColumnName.concat(" < '");
								sSubFilterSQL = sSubFilterSQL.concat(sFilterValue);
								sSubFilterSQL = sSubFilterSQL.concat("' OR ");
								sSubFilterSQL = sSubFilterSQL.concat(sColumnName);
								sSubFilterSQL = sSubFilterSQL.concat(" IS NULL");
							}
							else {
								sSubFilterSQL = sColumnName.concat(" IS NULL AND ");
								sSubFilterSQL = sSubFilterSQL.concat(sColumnName);
								sSubFilterSQL = sSubFilterSQL.concat(" IS NOT NULL");
							}
						}
					}
				}

				if ((iDataType != -7) && (iDataType != 2) && (iDataType != 4) && (iDataType != 11)) {
					// Character/Working Pattern column.
					if (iOperatorID == 1) {
						// Equal To.
						if (sValue.length == 0) {
							sSubFilterSQL = sColumnName.concat(" = '' OR ");
							sSubFilterSQL = sSubFilterSQL.concat(sColumnName);
							sSubFilterSQL = sSubFilterSQL.concat(" IS NULL");
						}
						else {
							// Replace the standard * and ? characters with the SQL % and _ characters.
							sModifiedFilterValue = sValue.replace(reApostrophe, "''");
							sModifiedFilterValue = sModifiedFilterValue.replace(reStar, "%");
							sModifiedFilterValue = sModifiedFilterValue.replace(reQuestion, "_");

							sSubFilterSQL = sColumnName.concat(" LIKE '");
							sSubFilterSQL = sSubFilterSQL.concat(sModifiedFilterValue);
							sSubFilterSQL = sSubFilterSQL.concat("'");
						}
					}

					if (iOperatorID == 2) {
						// Not Equal To.
						if (sValue.length == 0) {
							sSubFilterSQL = sColumnName.concat(" <> '' AND ");
							sSubFilterSQL = sSubFilterSQL.concat(sColumnName);
							sSubFilterSQL = sSubFilterSQL.concat(" IS NOT NULL");
						}
						else {
							// Replace the standard * and ? characters with the SQL % and _ characters.
							sModifiedFilterValue = sValue.replace(reApostrophe, "''");
							sModifiedFilterValue = sModifiedFilterValue.replace(reStar, "%");
							sModifiedFilterValue = sModifiedFilterValue.replace(reQuestion, "_");

							sSubFilterSQL = sColumnName.concat(" NOT LIKE '");
							sSubFilterSQL = sSubFilterSQL.concat(sModifiedFilterValue);
							sSubFilterSQL = sSubFilterSQL.concat("'");
						}
					}

					if (iOperatorID == 7) {
						// Contains.
						if (sValue.length == 0) {
							sSubFilterSQL = sColumnName.concat(" IS NULL OR ");
							sSubFilterSQL = sSubFilterSQL.concat(sColumnName);
							sSubFilterSQL = sSubFilterSQL.concat(" IS NOT NULL");
						}
						else {
							// Replace the standard * and ? characters with the SQL % and _ characters.
							sModifiedFilterValue = sValue.replace(reApostrophe, "''");

							sSubFilterSQL = sColumnName.concat(" LIKE '%");
							sSubFilterSQL = sSubFilterSQL.concat(sModifiedFilterValue);
							sSubFilterSQL = sSubFilterSQL.concat("%'");
						}
					}

					if (iOperatorID == 8) {
						// Does Not Contain.
						if (sValue.length == 0) {
							sSubFilterSQL = sColumnName.concat(" IS NULL AND ");
							sSubFilterSQL = sSubFilterSQL.concat(sColumnName);
							sSubFilterSQL = sSubFilterSQL.concat(" IS NOT NULL");
						}
						else {
							// Replace the standard * and ? characters with the SQL % and _ characters.
							sModifiedFilterValue = sValue.replace(reApostrophe, "''");

							sSubFilterSQL = sColumnName.concat(" NOT LIKE '%");
							sSubFilterSQL = sSubFilterSQL.concat(sModifiedFilterValue);
							sSubFilterSQL = sSubFilterSQL.concat("%'");
						}
					}
				}

				if (sSubFilterSQL.length > 0) {
					// Add the filter code for this grid record into the complete filter code.
					if (sFilterSQL.length > 0) {
						sFilterSQL = sFilterSQL.concat(" AND (");
					}
					else {
						sFilterSQL = sFilterSQL.concat("(");
					}

					sFilterSQL = sFilterSQL.concat(sSubFilterSQL);
					sFilterSQL = sFilterSQL.concat(")");

					// Add the definition to the definition string (colID<tab>opID<tab>value<tab>).
					sFilterDef = sFilterDef.concat(iColumnID);
					sFilterDef = sFilterDef.concat("	");
					sFilterDef = sFilterDef.concat(iOperatorID);
					sFilterDef = sFilterDef.concat("	");
					sFilterDef = sFilterDef.concat(sValue);
					sFilterDef = sFilterDef.concat("	");
				}
			}

			if (iIndex < frmFilterForm.ssOleDBGridFilterRecords.rows) {
				frmFilterForm.ssOleDBGridFilterRecords.MoveNext();
			}
			else {
				break;
			}
		}

		// Redisplay the workframe recedit control. 
		var sWorkPage = currentWorkFramePage();
		if (sWorkPage == "RECORDEDIT") {
			//window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "visible";
		}
		else {
			if (sWorkPage == "FIND") {
				//window.parent.frames("workframe").document.forms("frmFindForm").ssOleDBGridFindRecords.style.visibility = "visible";
			}
		}
		var frmGotoOption = document.getElementById("frmGotoOption");

		frmGotoOption.txtGotoOptionScreenID.value = frmFilterForm.txtOptionScreenID.value;
		frmGotoOption.txtGotoOptionTableID.value = frmFilterForm.txtOptionTableID.value;
		frmGotoOption.txtGotoOptionViewID.value = frmFilterForm.txtOptionViewID.value;
		frmGotoOption.txtGotoOptionFilterSQL.value = sFilterSQL;
		frmGotoOption.txtGotoOptionFilterDef.value = sFilterDef;
		frmGotoOption.txtGotoOptionPage.value = "emptyoption";
		frmGotoOption.txtGotoOptionAction.value = "SELECTFILTER";
		
		OpenHR.submitForm(frmGotoOption);

	}

	function CancelFilter() {
		// Redisplay the workframe recedit control. 
		var sWorkPage = currentWorkFramePage();
		if (sWorkPage == "RECORDEDIT") {
			//window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "visible";
			//window.parent.document.all.item("workframeset").cols = "*, 0";	
			$("#workframe").attr("data-framesource", "RECORDEDIT");
			$("#optionframe").hide();
			$("#workframe").show();
			refreshData(); //recedit
		}
		else {
			if (sWorkPage == "FIND") {
				//window.parent.frames("workframe").document.forms("frmFindForm").ssOleDBGridFindRecords.style.visibility = "visible";
				$("#workframe").attr("data-framesource", "FIND");
				$("#optionframe").hide();
				$("#workframe").show();
			}
		}



		var frmGotoOption = document.getElementById("frmGotoOption");

		frmGotoOption.txtGotoOptionAction.value = "CANCEL";
		frmGotoOption.txtGotoOptionPage.value = "emptyoption";
		OpenHR.submitForm(frmGotoOption);
		

	}

	function AddToList() {
		var frmFilterForm = document.getElementById("frmFilterForm");
		var fOK;
		var iDataType;
		var iSize;
		var iDecimals;
		var iTempSize;
		var iTempDecimals;
		var iIndex;
		var sValue;
		var sAddString;
		var sControlName;
		var sReqdControlName;
		var sReqdControlSizeName;
		var sReqdControlDecimalsName;
		var controlCollection = frmFilterForm.elements;
		var sConvertedValue;
		var sDecimalSeparator;
		var sThousandSeparator;
		var sPoint;
		var fDataTypeFound;
		var fSizeFound;
		var fDecimalsFound;

		sDecimalSeparator = "\\";
		sDecimalSeparator = sDecimalSeparator.concat(OpenHR.LocaleDecimalSeparator);
		var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");

		sThousandSeparator = "\\";
		sThousandSeparator = sThousandSeparator.concat(OpenHR.LocaleThousandSeparator);
		var reThousandSeparator = new RegExp(sThousandSeparator, "gi");

		sPoint = "\\.";
		var rePoint = new RegExp(sPoint, "gi");

		fOK = false;

		// Determine the data type of the filter column.
		iDataType = 12;
		iSize = 0;
		iDecimals = 0;
		sReqdControlName = "txtFilterColumn_";
		sReqdControlName = sReqdControlName.concat(frmFilterForm.selectColumn.options[frmFilterForm.selectColumn.selectedIndex].value);

		sReqdControlSizeName = "txtFilterColumnSize_";
		sReqdControlSizeName = sReqdControlSizeName.concat(frmFilterForm.selectColumn.options[frmFilterForm.selectColumn.selectedIndex].value);

		sReqdControlDecimalsName = "txtFilterColumnDecimals_";
		sReqdControlDecimalsName = sReqdControlDecimalsName.concat(frmFilterForm.selectColumn.options[frmFilterForm.selectColumn.selectedIndex].value);

		fDataTypeFound = false;
		fSizeFound = false;
		fDecimalsFound = false;
		if (controlCollection != null) {
			for (var i = 0; i < controlCollection.length; i++) {
				sControlName = controlCollection.item(i).name;

				if (sControlName == sReqdControlName) {
					iDataType = controlCollection.item(i).value;
					fDataTypeFound = true;
				}

				if (sControlName == sReqdControlSizeName) {
					iSize = controlCollection.item(i).value;
					fSizeFound = true;
				}

				if (sControlName == sReqdControlDecimalsName) {
					iDecimals = controlCollection.item(i).value;
					fDecimalsFound = true;
				}

				if ((fDataTypeFound == true) &&
					(fSizeFound == true) &&
					(fDecimalsFound == true)) {
					fOK = true;
					break;
				}
			}
		}

		if (fOK == true) {
			sAddString = frmFilterForm.selectColumn.options[frmFilterForm.selectColumn.selectedIndex].text;
			sAddString = sAddString.concat("	");

			if (iDataType == -7) {
				// Logic column (must be the equals operator).	
				sAddString = sAddString.concat("equals");
				sAddString = sAddString.concat("	");
				sAddString = sAddString.concat(frmFilterForm.selectValue.options[frmFilterForm.selectValue.selectedIndex].text);
				sAddString = sAddString.concat("	");
				sAddString = sAddString.concat(frmFilterForm.selectColumn.options[frmFilterForm.selectColumn.selectedIndex].value);
				sAddString = sAddString.concat("	");
				sAddString = sAddString.concat("1");
			}
		}

		if (fOK == true) {
			if ((iDataType == 2) || (iDataType == 4)) {
				// Numeric/Integer column.
				// Ensure that the value entered is numeric.
				sValue = frmFilterForm.txtValue.value;
				if (sValue.length == 0) {
					sValue = "0";
				}

				// Convert the value from locale to UK settings for use with the isNaN funtion.
				sConvertedValue = new String(sValue);
				// Remove any thousand separators.
				sConvertedValue = sConvertedValue.replace(reThousandSeparator, "");
				sValue = sConvertedValue;

				// Convert any decimal separators to '.'.
				if (OpenHR.LocaleDecimalSeparator != ".") {
					// Existing decimal points are invalid characters.
					sConvertedValue = sConvertedValue.replace(rePoint, "A");
					// Replace the locale decimal marker with the decimal point.
					sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
				}

				if (isNaN(sConvertedValue) == true) {
					fOK = false;
					OpenHR.messageBox("Invalid numeric value entered.");
					frmFilterForm.txtValue.focus();
				}
				else {
					iIndex = sConvertedValue.indexOf(".");
					if (iDataType == 4) {
						// Ensure that integer columns are compared with integer values.
						if (iIndex >= 0) {
							fOK = false;
							OpenHR.messageBox("Invalid integer value entered.");
							frmFilterForm.txtValue.focus();
						}
					}
					else {
						// Ensure numeric columns are compared with numeric values that do not exceed
						// their defined size and decimals settings.
						if (iIndex >= 0) {
							iTempSize = iIndex;
							iTempDecimals = sConvertedValue.length - iIndex - 1;
						}
						else {
							iTempSize = sConvertedValue.length;
							iTempDecimals = 0;
						}

						if ((sConvertedValue.substr(0, 1) == "+") ||
							(sConvertedValue.substr(0, 1) == "-")) {
							iTempSize = iTempSize - 1;
						}

						if (iTempSize > (iSize - iDecimals)) {
							fOK = false;
							OpenHR.messageBox("The column can only be compared to values with " + (iSize - iDecimals) + " digit(s) to the left of the decimal separator.");
							frmFilterForm.txtValue.focus();
						}
						else {
							if (iTempDecimals > iDecimals) {
								fOK = false;
								OpenHR.messageBox("The column can only be compared to values with " + iDecimals + " decimal place(s).");
								frmFilterForm.txtValue.focus();
							}
						}
					}

					if (fOK == true) {
						sAddString = sAddString.concat(frmFilterForm.selectConditionNum.options[frmFilterForm.selectConditionNum.selectedIndex].text);
						sAddString = sAddString.concat("	");
						sAddString = sAddString.concat(sValue);
						sAddString = sAddString.concat("	");
						sAddString = sAddString.concat(frmFilterForm.selectColumn.options[frmFilterForm.selectColumn.selectedIndex].value);
						sAddString = sAddString.concat("	");
						sAddString = sAddString.concat(frmFilterForm.selectConditionNum.options[frmFilterForm.selectConditionNum.selectedIndex].value);
					}
				}
			}
		}

		if (fOK == true) {
			if (iDataType == 11) {
				// Date column.
				// Ensure that the value entered is a date.
				sValue = frmFilterForm.txtValue.value;

				if (sValue.length == 0) {
					sAddString = sAddString.concat(frmFilterForm.selectConditionDate.options[frmFilterForm.selectConditionDate.selectedIndex].text);
					sAddString = sAddString.concat("	");
					sAddString = sAddString.concat("");
					sAddString = sAddString.concat("	");
					sAddString = sAddString.concat(frmFilterForm.selectColumn.options[frmFilterForm.selectColumn.selectedIndex].value);
					sAddString = sAddString.concat("	");
					sAddString = sAddString.concat(frmFilterForm.selectConditionDate.options[frmFilterForm.selectConditionDate.selectedIndex].value);
				}
				else {
					// Convert the date to SQL format (use this as a validation check).
					// An empty string is returned if the date is invalid.
					sValue = menu_convertLocaleDateToSQL(sValue);
					if (sValue.length == 0) {
						fOK = false;
						OpenHR.messageBox("Invalid date value entered.");
						frmFilterForm.txtValue.focus();
					}
					else {
						sValue = menu_convertLocaleDateToSQL(sValue);

						sAddString = sAddString.concat(frmFilterForm.selectConditionDate.options[frmFilterForm.selectConditionDate.selectedIndex].text);
						sAddString = sAddString.concat("	");
						sAddString = sAddString.concat(sValue);
						sAddString = sAddString.concat("	");
						sAddString = sAddString.concat(frmFilterForm.selectColumn.options[frmFilterForm.selectColumn.selectedIndex].value);
						sAddString = sAddString.concat("	");
						sAddString = sAddString.concat(frmFilterForm.selectConditionDate.options[frmFilterForm.selectConditionDate.selectedIndex].value);
					}
				}
			}
		}

		if (fOK == true) {
			if ((iDataType != -7) && (iDataType != 2) && (iDataType != 4) && (iDataType != 11)) {
				// Character/Working Pattern column.
				sValue = frmFilterForm.txtValue.value;

				sAddString = sAddString.concat(frmFilterForm.selectConditionChar.options[frmFilterForm.selectConditionChar.selectedIndex].text);
				sAddString = sAddString.concat("	");
				sAddString = sAddString.concat(sValue);
				sAddString = sAddString.concat("	");
				sAddString = sAddString.concat(frmFilterForm.selectColumn.options[frmFilterForm.selectColumn.selectedIndex].value);
				sAddString = sAddString.concat("	");
				sAddString = sAddString.concat(frmFilterForm.selectConditionChar.options[frmFilterForm.selectConditionChar.selectedIndex].value);
			}
		}
		if (fOK == true) {
			frmFilterForm.ssOleDBGridFilterRecords.AddItem(sAddString);
			frmFilterForm.ssOleDBGridFilterRecords.MoveLast();
			frmFilterForm.ssOleDBGridFilterRecords.SelBookmarks.Add(frmFilterForm.ssOleDBGridFilterRecords.Bookmark);

			refreshControls();
			frmFilterForm.selectColumn.focus();
		}
	}

	function displayOperatorSelector(piDataType) {		
		var frmFilterForm = document.getElementById("frmFilterForm");

		if (piDataType == -7) {
			// Display the logic operator control.
			frmFilterForm.txtConditionLogic.style.width = "175px";
			frmFilterForm.txtConditionLogic.style.visibility = "";
			frmFilterForm.txtConditionLogic.style.position = "";
			frmFilterForm.txtConditionLogic.style.top = "";
			frmFilterForm.txtConditionLogic.style.left = "";

			frmFilterForm.selectValue.style.width = "175px";
			frmFilterForm.selectValue.style.visibility = "";
			frmFilterForm.selectValue.style.position = "";
			frmFilterForm.selectValue.style.top = "";
			frmFilterForm.selectValue.style.left = "";

			frmFilterForm.txtValue.style.width = "0px";
			frmFilterForm.txtValue.style.visibility = "hidden";
			frmFilterForm.txtValue.style.position = "absolute";
			frmFilterForm.txtValue.style.top = 0;
			frmFilterForm.txtValue.style.left = 0;
		}
		else {
			// Hide the logic operator control.
			frmFilterForm.txtConditionLogic.style.width = "0px";
			frmFilterForm.txtConditionLogic.style.visibility = "hidden";
			frmFilterForm.txtConditionLogic.style.position = "absolute";
			frmFilterForm.txtConditionLogic.style.top = "0px";
			frmFilterForm.txtConditionLogic.style.left = "0px";

			frmFilterForm.selectValue.style.width = "0px";
			frmFilterForm.selectValue.style.visibility = "hidden";
			frmFilterForm.selectValue.style.position = "absolute";
			frmFilterForm.selectValue.style.top = "0px";
			frmFilterForm.selectValue.style.left = "0px";

			frmFilterForm.txtValue.style.width = "175px";
			frmFilterForm.txtValue.style.visibility = "";
			frmFilterForm.txtValue.style.position = "";
			frmFilterForm.txtValue.style.top = "";
			frmFilterForm.txtValue.style.left = "";
		}

		if ((piDataType == 2) || (piDataType == 4)) {
			// Display the Numeric/Integer operator control.
			frmFilterForm.selectConditionNum.style.width = "175px";
			frmFilterForm.selectConditionNum.style.visibility = "";
			frmFilterForm.selectConditionNum.style.position = "";
			frmFilterForm.selectConditionNum.style.top = "";
			frmFilterForm.selectConditionNum.style.left = "";
		}
		else {
			// Hide the Numeric/Integer operator control.
			frmFilterForm.selectConditionNum.style.width = "0px";
			frmFilterForm.selectConditionNum.style.visibility = "hidden";
			frmFilterForm.selectConditionNum.style.position = "absolute";
			frmFilterForm.selectConditionNum.style.top = "0px";
			frmFilterForm.selectConditionNum.style.left = "0px";
		}

		if (piDataType == 11) {
			// Display the Date operator control.
			frmFilterForm.selectConditionDate.style.width = "175px";
			frmFilterForm.selectConditionDate.style.visibility = "";
			frmFilterForm.selectConditionDate.style.position = "";
			frmFilterForm.selectConditionDate.style.top = "";
			frmFilterForm.selectConditionDate.style.left = "";
		}
		else {
			// Hide the Date operator control.
			frmFilterForm.selectConditionDate.style.width = "0px";
			frmFilterForm.selectConditionDate.style.visibility = "hidden";
			frmFilterForm.selectConditionDate.style.position = "absolute";
			frmFilterForm.selectConditionDate.style.top = "0px";
			frmFilterForm.selectConditionDate.style.left = "0px";
		}

		if ((piDataType != -7) && (piDataType != 2) && (piDataType != 4) && (piDataType != 11)) {
			// Display the Character/Working Pattern operator control.
			frmFilterForm.selectConditionChar.style.width = "175px";
			frmFilterForm.selectConditionChar.style.visibility = "";
			frmFilterForm.selectConditionChar.style.position = "";
			frmFilterForm.selectConditionChar.style.top = "";
			frmFilterForm.selectConditionChar.style.left = "";
		}
		else {
			// Hide the Character/Working Pattern operator control.
			frmFilterForm.selectConditionChar.style.width = "0px";
			frmFilterForm.selectConditionChar.style.visibility = "hidden";
			frmFilterForm.selectConditionChar.style.position = "absolute";
			frmFilterForm.selectConditionChar.style.top = "0px";
			frmFilterForm.selectConditionChar.style.left = "0px";
		}
	}

	function refreshOperatorCombo() {		
		var fFound;
		var sControlName;
		var sReqdControlName;
		var frmFilterForm = document.getElementById("frmFilterForm");
		var controlCollection = frmFilterForm.elements;

		fFound = false;
		sReqdControlName = "txtFilterColumn_";
		sReqdControlName = sReqdControlName.concat(frmFilterForm.selectColumn.options[frmFilterForm.selectColumn.selectedIndex].value);

		if (controlCollection != null) {
			for (var i = 0; i < controlCollection.length; i++) {
				sControlName = controlCollection.item(i).name;
				if (sControlName == sReqdControlName) {
					fFound = true;
					displayOperatorSelector(controlCollection.item(i).value);
					break;
				}
			}
		}

		if (fFound == false) {
			displayOperatorSelector(12);
		}
	}


	function removeAll() {
		var frmFilterForm = document.getElementById("frmFilterForm");
		frmFilterForm.ssOleDBGridFilterRecords.RemoveAll();
		refreshControls();
	}

	function remove() {
		var frmFilterForm = document.getElementById("frmFilterForm");

		if (frmFilterForm.ssOleDBGridFilterRecords.Rows > 0) {
			var iRowIndex = frmFilterForm.ssOleDBGridFilterRecords.AddItemRowIndex(frmFilterForm.ssOleDBGridFilterRecords.Bookmark);

			if ((frmFilterForm.ssOleDBGridFilterRecords.Rows == 1) && (iRowIndex == 0)) {
				frmFilterForm.ssOleDBGridFilterRecords.RemoveAll();
			}
			else {
				frmFilterForm.ssOleDBGridFilterRecords.RemoveItem(iRowIndex);
			}

			if (frmFilterForm.ssOleDBGridFilterRecords.Rows > 0) {
				if (iRowIndex == 0) {
					frmFilterForm.ssOleDBGridFilterRecords.MoveFirst();
				}
				else {
					if (iRowIndex >= frmFilterForm.ssOleDBGridFilterRecords.Rows) {
						frmFilterForm.ssOleDBGridFilterRecords.MoveLast();
					}
				}

				frmFilterForm.ssOleDBGridFilterRecords.SelBookmarks.Add(frmFilterForm.ssOleDBGridFilterRecords.Bookmark);
			}
		}

		refreshControls();
	}

	function refreshControls() {
		var frmFilterForm = document.getElementById("frmFilterForm");

		if (frmFilterForm.ssOleDBGridFilterRecords.Rows > 0) {
			button_disable(frmFilterForm.cmdRemoveAll, false);

			if (frmFilterForm.ssOleDBGridFilterRecords.SelBookmarks.Count > 0) {
				button_disable(frmFilterForm.cmdRemove, false);
			}
			else {
				button_disable(frmFilterForm.cmdRemove, true);
			}
		}
		else {
			button_disable(frmFilterForm.cmdRemoveAll, true);
			button_disable(frmFilterForm.cmdRemove, true);
		}
	}

	function currentWorkFramePage() {
		// Return the current page in the workframeset.

		var sCurrentPage = $("#workframe").attr("data-framesource").replace(".asp", "");

		//var sCols = window.parent.document.all.item("workframeset").cols;

		//re = / /gi;
		//sCols = sCols.replace(re, "");
		//sCols = sCols.substr(0, 1);

		//// Work frame is in view.
		//sCurrentPage = window.parent.frames("workframe").document.location;
		//sCurrentPage = sCurrentPage.toString();

		//if (sCurrentPage.lastIndexOf("/") > 0) {
		//	sCurrentPage = sCurrentPage.substr(sCurrentPage.lastIndexOf("/") + 1);
		//}

		//if (sCurrentPage.indexOf(".") > 0) {
		//	sCurrentPage = sCurrentPage.substr(0, sCurrentPage.indexOf("."));
		//}

		//re = / /gi;
		//sCurrentPage = sCurrentPage.replace(re, "");
		//sCurrentPage = sCurrentPage.toUpperCase();

		return (sCurrentPage);
	}

</script>

<div <%=session("BodyTag")%>>
	<form action="" method="POST" id="frmFilterForm" name="frmFilterForm">
		<table style="text-align: center; margin: 0px auto; border-spacing: 5px; border-collapse: collapse; width: 100%; height: 100%;" class="outline">
			<tr>
				<td>
					<table id="filterTable" style="width: 100%; height: 100%; border-spacing: 0px; border-collapse: collapse;" class="invisible">
						<tr>
							<td style="height: 10%;" colspan="3">
								<h3 class="pageTitle">Define Filter</h3>
							</td>
						</tr>

						<tr height="160">
							<td width="10"></td>
							<td>
								<object classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"
									codebase="cabs/COAInt_Grid.cab#version=3,1,3,6"
									id="ssOleDBGridFilterRecords"
									name="ssOleDBGridFilterRecords"
									style="HEIGHT: 100%; LEFT: 0; TOP: 0; WIDTH: 100%"
									viewastext>
									<param name="ScrollBars" value="4">
									<param name="_Version" value="196617">
									<param name="DataMode" value="2">
									<param name="Cols" value="0">
									<param name="Rows" value="0">
									<param name="BorderStyle" value="1">
									<param name="RecordSelectors" value="0">
									<param name="GroupHeaders" value="0">
									<param name="ColumnHeaders" value="1">
									<param name="GroupHeadLines" value="1">
									<param name="HeadLines" value="1">
									<param name="FieldDelimiter" value="(None)">
									<param name="FieldSeparator" value="(Tab)">
									<param name="Row.Count" value="0">
									<param name="Col.Count" value="1">
									<param name="stylesets.count" value="0">
									<param name="TagVariant" value="EMPTY">
									<param name="UseGroups" value="0">
									<param name="HeadFont3D" value="0">
									<param name="Font3D" value="0">
									<param name="DividerType" value="3">
									<param name="DividerStyle" value="1">
									<param name="DefColWidth" value="0">
									<param name="BeveColorScheme" value="2">
									<param name="BevelColorFrame" value="-2147483642">
									<param name="BevelColorHighlight" value="-2147483628">
									<param name="BevelColorShadow" value="-2147483632">
									<param name="BevelColorFace" value="-2147483633">
									<param name="CheckBox3D" value="-1">
									<param name="AllowAddNew" value="0">
									<param name="AllowDelete" value="0">
									<param name="AllowUpdate" value="0">
									<param name="MultiLine" value="0">
									<param name="ActiveCellStyleSet" value="">
									<param name="RowSelectionStyle" value="0">
									<param name="AllowRowSizing" value="0">
									<param name="AllowGroupSizing" value="0">
									<param name="AllowColumnSizing" value="-1">
									<param name="AllowGroupMoving" value="0">
									<param name="AllowColumnMoving" value="0">
									<param name="AllowGroupSwapping" value="0">
									<param name="AllowColumnSwapping" value="0">
									<param name="AllowGroupShrinking" value="0">
									<param name="AllowColumnShrinking" value="0">
									<param name="AllowDragDrop" value="0">
									<param name="UseExactRowCount" value="-1">
									<param name="SelectTypeCol" value="0">
									<param name="SelectTypeRow" value="1">
									<param name="SelectByCell" value="-1">
									<param name="BalloonHelp" value="0">
									<param name="RowNavigation" value="1">
									<param name="CellNavigation" value="0">
									<param name="MaxSelectedRows" value="1">
									<param name="HeadStyleSet" value="">
									<param name="StyleSet" value="">
									<param name="ForeColorEven" value="0">
									<param name="ForeColorOdd" value="0">
									<param name="BackColorEven" value="16777215">
									<param name="BackColorOdd" value="16777215">
									<param name="Levels" value="1">
									<param name="RowHeight" value="503">
									<param name="ExtraHeight" value="0">
									<param name="ActiveRowStyleSet" value="">
									<param name="CaptionAlignment" value="2">
									<param name="SplitterPos" value="0">
									<param name="SplitterVisible" value="0">
									<param name="Columns.Count" value="5">
									<param name="Columns(0).Width" value="6500">
									<param name="Columns(0).Visible" value="-1">
									<param name="Columns(0).Columns.Count" value="1">
									<param name="Columns(0).Caption" value="Field">
									<param name="Columns(0).Name" value="FilterColumn">
									<param name="Columns(0).Alignment" value="0">
									<param name="Columns(0).CaptionAlignment" value="3">
									<param name="Columns(0).Bound" value="0">
									<param name="Columns(0).AllowSizing" value="1">
									<param name="Columns(0).DataField" value="Column 0">
									<param name="Columns(0).DataType" value="8">
									<param name="Columns(0).Level" value="0">
									<param name="Columns(0).NumberFormat" value="">
									<param name="Columns(0).Case" value="0">
									<param name="Columns(0).FieldLen" value="256">
									<param name="Columns(0).VertScrollBar" value="0">
									<param name="Columns(0).Locked" value="0">
									<param name="Columns(0).Style" value="0">
									<param name="Columns(0).ButtonsAlways" value="0">
									<param name="Columns(0).RowCount" value="0">
									<param name="Columns(0).ColCount" value="1">
									<param name="Columns(0).HasHeadForeColor" value="0">
									<param name="Columns(0).HasHeadBackColor" value="0">
									<param name="Columns(0).HasForeColor" value="0">
									<param name="Columns(0).HasBackColor" value="0">
									<param name="Columns(0).HeadForeColor" value="0">
									<param name="Columns(0).HeadBackColor" value="0">
									<param name="Columns(0).ForeColor" value="0">
									<param name="Columns(0).BackColor" value="0">
									<param name="Columns(0).HeadStyleSet" value="">
									<param name="Columns(0).StyleSet" value="">
									<param name="Columns(0).Nullable" value="1">
									<param name="Columns(0).Mask" value="">
									<param name="Columns(0).PromptInclude" value="0">
									<param name="Columns(0).ClipMode" value="0">
									<param name="Columns(0).PromptChar" value="95">
									<param name="Columns(1).Width" value="6500">
									<param name="Columns(1).Visible" value="-1">
									<param name="Columns(1).Columns.Count" value="1">
									<param name="Columns(1).Caption" value="Operator">
									<param name="Columns(1).Name" value="FilterOperator">
									<param name="Columns(1).Alignment" value="0">
									<param name="Columns(1).CaptionAlignment" value="3">
									<param name="Columns(1).Bound" value="0">
									<param name="Columns(1).AllowSizing" value="1">
									<param name="Columns(1).DataField" value="Column 1">
									<param name="Columns(1).DataType" value="8">
									<param name="Columns(1).Level" value="0">
									<param name="Columns(1).NumberFormat" value="">
									<param name="Columns(1).Case" value="0">
									<param name="Columns(1).FieldLen" value="256">
									<param name="Columns(1).VertScrollBar" value="0">
									<param name="Columns(1).Locked" value="0">
									<param name="Columns(1).Style" value="0">
									<param name="Columns(1).ButtonsAlways" value="0">
									<param name="Columns(1).RowCount" value="0">
									<param name="Columns(1).ColCount" value="1">
									<param name="Columns(1).HasHeadForeColor" value="0">
									<param name="Columns(1).HasHeadBackColor" value="0">
									<param name="Columns(1).HasForeColor" value="0">
									<param name="Columns(1).HasBackColor" value="0">
									<param name="Columns(1).HeadForeColor" value="0">
									<param name="Columns(1).HeadBackColor" value="0">
									<param name="Columns(1).ForeColor" value="0">
									<param name="Columns(1).BackColor" value="0">
									<param name="Columns(1).HeadStyleSet" value="">
									<param name="Columns(1).StyleSet" value="">
									<param name="Columns(1).Nullable" value="1">
									<param name="Columns(1).Mask" value="">
									<param name="Columns(1).PromptInclude" value="0">
									<param name="Columns(1).ClipMode" value="0">
									<param name="Columns(1).PromptChar" value="95">
									<param name="Columns(2).Width" value="6500">
									<param name="Columns(2).Visible" value="-1">
									<param name="Columns(2).Columns.Count" value="1">
									<param name="Columns(2).Caption" value="Value">
									<param name="Columns(2).Name" value="FilterText">
									<param name="Columns(2).Alignment" value="0">
									<param name="Columns(2).CaptionAlignment" value="3">
									<param name="Columns(2).Bound" value="0">
									<param name="Columns(2).AllowSizing" value="1">
									<param name="Columns(2).DataField" value="Column 2">
									<param name="Columns(2).DataType" value="8">
									<param name="Columns(2).Level" value="0">
									<param name="Columns(2).NumberFormat" value="">
									<param name="Columns(2).Case" value="0">
									<param name="Columns(2).FieldLen" value="256">
									<param name="Columns(2).VertScrollBar" value="0">
									<param name="Columns(2).Locked" value="0">
									<param name="Columns(2).Style" value="0">
									<param name="Columns(2).ButtonsAlways" value="0">
									<param name="Columns(2).RowCount" value="0">
									<param name="Columns(2).ColCount" value="1">
									<param name="Columns(2).HasHeadForeColor" value="0">
									<param name="Columns(2).HasHeadBackColor" value="0">
									<param name="Columns(2).HasForeColor" value="0">
									<param name="Columns(2).HasBackColor" value="0">
									<param name="Columns(2).HeadForeColor" value="0">
									<param name="Columns(2).HeadBackColor" value="0">
									<param name="Columns(2).ForeColor" value="0">
									<param name="Columns(2).BackColor" value="0">
									<param name="Columns(2).HeadStyleSet" value="">
									<param name="Columns(2).StyleSet" value="">
									<param name="Columns(2).Nullable" value="1">
									<param name="Columns(2).Mask" value="">
									<param name="Columns(2).PromptInclude" value="0">
									<param name="Columns(2).ClipMode" value="0">
									<param name="Columns(2).PromptChar" value="95">
									<param name="Columns(3).Width" value="3200">
									<param name="Columns(3).Visible" value="0">
									<param name="Columns(3).Columns.Count" value="1">
									<param name="Columns(3).Caption" value="FilterColumnID">
									<param name="Columns(3).Name" value="FilterColumnID">
									<param name="Columns(3).Alignment" value="0">
									<param name="Columns(3).CaptionAlignment" value="3">
									<param name="Columns(3).Bound" value="0">
									<param name="Columns(3).AllowSizing" value="0">
									<param name="Columns(3).DataField" value="Column 3">
									<param name="Columns(3).DataType" value="8">
									<param name="Columns(3).Level" value="0">
									<param name="Columns(3).NumberFormat" value="">
									<param name="Columns(3).Case" value="0">
									<param name="Columns(3).FieldLen" value="256">
									<param name="Columns(3).VertScrollBar" value="0">
									<param name="Columns(3).Locked" value="0">
									<param name="Columns(3).Style" value="0">
									<param name="Columns(3).ButtonsAlways" value="0">
									<param name="Columns(3).RowCount" value="0">
									<param name="Columns(3).ColCount" value="1">
									<param name="Columns(3).HasHeadForeColor" value="0">
									<param name="Columns(3).HasHeadBackColor" value="0">
									<param name="Columns(3).HasForeColor" value="0">
									<param name="Columns(3).HasBackColor" value="0">
									<param name="Columns(3).HeadForeColor" value="0">
									<param name="Columns(3).HeadBackColor" value="0">
									<param name="Columns(3).ForeColor" value="0">
									<param name="Columns(3).BackColor" value="0">
									<param name="Columns(3).HeadStyleSet" value="">
									<param name="Columns(3).StyleSet" value="">
									<param name="Columns(3).Nullable" value="1">
									<param name="Columns(3).Mask" value="">
									<param name="Columns(3).PromptInclude" value="0">
									<param name="Columns(3).ClipMode" value="0">
									<param name="Columns(3).PromptChar" value="95">
									<param name="Columns(4).Width" value="0">
									<param name="Columns(4).Visible" value="0">
									<param name="Columns(4).Columns.Count" value="1">
									<param name="Columns(4).Caption" value="FilterOperatorID">
									<param name="Columns(4).Name" value="FilterOperatorID">
									<param name="Columns(4).Alignment" value="0">
									<param name="Columns(4).CaptionAlignment" value="3">
									<param name="Columns(4).Bound" value="0">
									<param name="Columns(4).AllowSizing" value="0">
									<param name="Columns(4).DataField" value="Column 4">
									<param name="Columns(4).DataType" value="8">
									<param name="Columns(4).Level" value="0">
									<param name="Columns(4).NumberFormat" value="">
									<param name="Columns(4).Case" value="0">
									<param name="Columns(4).FieldLen" value="256">
									<param name="Columns(4).VertScrollBar" value="0">
									<param name="Columns(4).Locked" value="0">
									<param name="Columns(4).Style" value="0">
									<param name="Columns(4).ButtonsAlways" value="0">
									<param name="Columns(4).RowCount" value="0">
									<param name="Columns(4).ColCount" value="1">
									<param name="Columns(4).HasHeadForeColor" value="0">
									<param name="Columns(4).HasHeadBackColor" value="0">
									<param name="Columns(4).HasForeColor" value="0">
									<param name="Columns(4).HasBackColor" value="0">
									<param name="Columns(4).HeadForeColor" value="0">
									<param name="Columns(4).HeadBackColor" value="0">
									<param name="Columns(4).ForeColor" value="0">
									<param name="Columns(4).BackColor" value="0">
									<param name="Columns(4).HeadStyleSet" value="">
									<param name="Columns(4).StyleSet" value="">
									<param name="Columns(4).Nullable" value="1">
									<param name="Columns(4).Mask" value="">
									<param name="Columns(4).PromptInclude" value="0">
									<param name="Columns(4).ClipMode" value="0">
									<param name="Columns(4).PromptChar" value="95">
									<param name="UseDefaults" value="-1">
									<param name="TabNavigation" value="1">
									<param name="BatchUpdate" value="0">
									<param name="_ExtentX" value="16087">
									<param name="_ExtentY" value="4630">
									<param name="_StockProps" value="79">
									<param name="Caption" value="">
									<param name="ForeColor" value="0">
									<param name="BackColor" value="16777215">
									<param name="Enabled" value="-1">
									<param name="DataMember" value="">
								</object>
							</td>
							<td width="10"></td>
						</tr>

						<tr>
							<td width="10"></td>
							<td height="10">
								<table class="invisible" style="width: 100%; border-spacing: 5px; border-collapse: collapse;">
									<tr height="10">
										<td height="10">&nbsp;
										</td>
										<td width="10" height="10">
											<input id="cmdRemove" name="cmdRemove" type="button" value="Remove" style="WIDTH: 100px" width="100" class="btn"
												onclick="remove()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
										<td width="10" height="10"></td>
										<td width="10" height="10">
											<input id="cmdRemoveAll" name="cmdRemoveAll" type="button" value="Remove All" style="WIDTH: 100px" width="100" class="btn"
												onclick="removeAll()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
									</tr>
								</table>
							</td>
							<td width="10"></td>
						</tr>

						<tr height="10">
							<td width="10"></td>
							<td height="10">
								<strong>Define more criteria</strong>
							</td>
							<td width="10"></td>
						</tr>

						<tr>
							<td width="10"></td>
							<td height="10">

								<table class="invisible" style="width: 100%; border-spacing: 5px; border-collapse: collapse;">
									<tr>
										<td width="10" height="10"></td>
										<td width="175" height="10">Field :
										</td>
										<td width="10" height="10"></td>
										<td width="175" height="10">Operator :
										</td>
										<td width="10" height="10"></td>
										<td width="175">Value :
										</td>
										<td colspan="2" height="10"></td>
									</tr>

									<tr>
										<td height="10" colspan="8"></td>
									</tr>

									<tr>
										<td width="10" height="10"></td>
										<td width="175" height="10">
											<select id="selectColumn" name="selectColumn" class="combo" style="HEIGHT: 22px; WIDTH: 200px"
												onchange="refreshOperatorCombo()">
											<%
	' Populate the columns combo.
	dim iCount
	dim sErrorDescription = ""
	if Len(sErrorDescription) = 0 then
		' Get the column records.
		dim  cmdFilterColumns = CreateObject("ADODB.Command")
		cmdFilterColumns.CommandText = "sp_ASRIntGetFilterColumns"
		cmdFilterColumns.CommandType = 4 ' Stored Procedure
		cmdFilterColumns.ActiveConnection = session("databaseConnection")

		dim  prmTableID = cmdFilterColumns.CreateParameter("tableID",3,1)
		cmdFilterColumns.Parameters.Append(prmTableID)
		prmTableID.value = cleanNumeric(session("optionTableID"))

		dim prmViewID = cmdFilterColumns.CreateParameter("viewID",3,1)
		cmdFilterColumns.Parameters.Append(prmViewID)
		prmViewID.value = cleanNumeric(session("optionViewID"))

		Dim prmRealSource = cmdFilterColumns.CreateParameter("realSource",200,2,8000) '200=varchar, 2=output, 8000=size
		cmdFilterColumns.Parameters.Append(prmRealSource)

		err.Clear()
		dim rstFilterColumns = cmdFilterColumns.Execute

		if (err.Number <> 0) then
			sErrorDescription = "The filter columns could not be retrieved." & vbcrlf & formatError(Err.Description)
		end if

		if len(sErrorDescription) = 0 then
			iCount = 0
			do while not rstFilterColumns.EOF
				Response.Write("						<OPTION value=" & rstFilterColumns.Fields(0).Value)
				if iCount = 0 then
					Response.Write(" SELECTED")
				end if
				
				Response.write(">" & replace(rstFilterColumns.Fields(1).Value, "_", " ") & "</OPTION>" & vbcrlf)
				iCount = iCount + 1
				rstFilterColumns.MoveNext
			loop

			Response.Write("					</SELECT>" & vbcrlf)

			' Release the ADO recordset object.
			rstFilterColumns.close
			rstFilterColumns = nothing

			' NB. IMPORTANT ADO NOTE.
			' When calling a stored procedure which returns a recordset AND has output parameters
			' you need to close the recordset and set it to nothing before using the output parameters. 
			Response.Write( "<INPUT type='hidden' id=txtRealSource name=txtRealSource value=""" & replace(replace(cmdFilterColumns.Parameters("realSource").Value, "'", "'''"), """", "&quot;") & """>" & vbcrlf)
		end if
	
		' Release the ADO command object.
		cmdFilterColumns = nothing
	end if

	' Populate the columns combo.
	if len(sErrorDescription) = 0 then
		' Get the column records.
		dim cmdFilterColumns = CreateObject("ADODB.Command")
		cmdFilterColumns.CommandText = "sp_ASRIntGetFilterColumns"
		cmdFilterColumns.CommandType = 4 ' Stored Procedure
		cmdFilterColumns.ActiveConnection = session("databaseConnection")

		dim prmTableID = cmdFilterColumns.CreateParameter("tableID",3,1) '3=integer, 1=input
		cmdFilterColumns.Parameters.Append(prmTableID)
		prmTableID.value = cleanNumeric(session("optionTableID"))

		dim prmViewID = cmdFilterColumns.CreateParameter("viewID",3,1) '3=integer, 1=input
		cmdFilterColumns.Parameters.Append(prmViewID)
		prmViewID.value = cleanNumeric(session("optionViewID"))

		dim prmRealSource = cmdFilterColumns.CreateParameter("realSource",200,2,8000) '200=varchar, 2=output, 8000=size
		cmdFilterColumns.Parameters.Append(prmRealSource)

		err.Clear()
		dim rstFilterColumns = cmdFilterColumns.Execute

		if (err.Number <> 0) then
			sErrorDescription = "The filter columns could not be retrieved." & vbcrlf & formatError(Err.Description)
		end if

		if len(sErrorDescription) = 0 then
			do while not rstFilterColumns.EOF
				Response.Write("					<INPUT type='hidden' id=txtFilterColumn_" & rstFilterColumns.Fields(0).Value & " name=txtFilterColumn_" & rstFilterColumns.Fields(0).Value & " value=" & rstFilterColumns.Fields(2).Value & ">")
				Response.Write("					<INPUT type='hidden' id=txtFilterColumnSize_" & rstFilterColumns.Fields(0).Value & " name=txtFilterColumnSize_" & rstFilterColumns.Fields(0).Value & " value=" & rstFilterColumns.Fields("size").Value & ">")
				Response.Write("					<INPUT type='hidden' id=txtFilterColumnDecimals_" & rstFilterColumns.Fields(0).Value & " name=txtFilterColumnDecimals_" & rstFilterColumns.Fields(0).Value & " value=" & rstFilterColumns.Fields("decimals").Value & ">")
				rstFilterColumns.MoveNext
			loop

			' Release the ADO recordset object.
			rstFilterColumns.close
			rstFilterColumns = nothing
		end if
	
		' Release the ADO command object.
		cmdFilterColumns = nothing
	end if
											%>

										</td>
										<td width="10" height="10"></td>
										<td width="175" height="10">
											<input type="text" id="txtConditionLogic" class="text textdisabled" name="selectConditionLogic" disabled="disabled" value="is equal to">

											<select id="selectConditionDate" name="selectConditionDate" class="combo" style="HEIGHT: 22px; LEFT: 0; POSITION: absolute; TOP: 0; VISIBILITY: hidden; WIDTH: 200px">
												<option value="1">is equal to</option>
												<option value="2">is NOT equal to</option>
												<option value="5">after</option>
												<option value="6">before</option>
												<option value="4">is equal to or after</option>
												<option value="3">is equal to or before</option>
											</select>

											<select id="selectConditionNum" name="selectConditionNum" class="combo" style="HEIGHT: 22px; LEFT: 0; POSITION: absolute; TOP: 0; VISIBILITY: hidden; WIDTH: 200px">
												<option value="1">is equal to</option>
												<option value="2">is NOT equal to</option>
												<option value="5">is greater than</option>
												<option value="4">is greater than or equal to</option>
												<option value="6">is less than</option>
												<option value="3">is less than or equal to</option>
											</select>

											<select id="selectConditionChar" name="selectConditionChar" class="combo" style="HEIGHT: 22px; LEFT: 0; POSITION: absolute; TOP: 0; VISIBILITY: hidden; WIDTH: 200px">
												<option value="1">is equal to</option>
												<option value="2">is NOT equal to</option>
												<option value="7">contains</option>
												<option value="8">does not contain</option>
											</select>
										</td>
										<td width="10" height="10"></td>
										<td width="175" height="10">
											<select id="selectValue" name='selectValue"' class="combo" style="HEIGHT: 22px; WIDTH: 200px">
												<option value="1">True</option>
												<option value="0">False</option>
											</select>
											<input id="txtValue" name="txtValue" class="text" style="HEIGHT: 22px; LEFT: 0; POSITION: absolute; TOP: 0; VISIBILITY: hidden; WIDTH: 175px">
										</td>
										<td height="10"></td>
										<td width="10" height="10">
											<input id="cmdAddToList" name="cmdAddToList" class="btn" type="button" value="Add To List" style="WIDTH: 100px" width="100"
												onclick="AddToList()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
									</tr>

								</table>
							</td>
							<td width="10"></td>
						</tr>

						<tr height="10">
							<td width="10"></td>
							<td height="10">
								<table width="100%" class="invisible" style="border-spacing: 0; border-collapse: collapse;">
									<tr height="10">
										<td height="10">&nbsp;
										</td>
										<td width="10" height="10">
											<input id="cmdSelect" name="cmdSelect" class="btn" type="button" value="OK" style="WIDTH: 100px" width="100"
												onclick="SelectFilter()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
										<td width="10" height="10"></td>
										<td width="10" height="10">
											<input id="cmdCancel" name="cmdCancel" type="button" value="Cancel" style="WIDTH: 100px" width="100" class="btn"
												onclick="CancelFilter()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
									</tr>
								</table>
							</td>
							<td width="10"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>

		<%
	Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbcrlf)
	Response.Write("<INPUT type='hidden' id=txtOptionScreenID name=txtOptionScreenID value=" & session("optionScreenID") & ">" & vbcrlf)
	Response.Write("<INPUT type='hidden' id=txtOptionTableID name=txtOptionTableID value=" & session("optionTableID") & ">" & vbcrlf)
	Response.Write("<INPUT type='hidden' id=txtOptionViewID name=txtOptionViewID value=" & session("optionViewID") & ">" & vbcrlf)
	Response.Write("<INPUT type='hidden' id=txtOptionFilterDef name=txtOptionFilterDef value=""" & replace(session("optionFilterDef"), """", "&quot;") & """>" & vbcrlf)
		%>
	</form>
	<form action="filterselect_Submit" method="post" id="frmGotoOption" name="frmGotoOption">
		<%Html.RenderPartial("~/Views/Shared/gotoOption.ascx")%>
	</form>
	
	<script type="text/javascript">filterselect_window_onload()</script>
	

</div>
