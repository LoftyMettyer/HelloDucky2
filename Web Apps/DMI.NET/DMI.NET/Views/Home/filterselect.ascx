<%@ control language="VB" inherits="System.Web.Mvc.ViewUserControl" %>
<%@ import namespace="DMI.NET" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="DMI.NET.Helpers" %>

<script type="text/javascript">
	var colMode = [];
	var colData = [];
	var colNames = [];

	//Create the column model
	colMode.push({ name: 'ID', hidden: true });
	colMode.push({ name: 'Field', width: 100 });
	colMode.push({ name: 'Operator', width: 100 });
	colMode.push({ name: 'Value', width: 100 });
	colMode.push({ name: 'ColumnID', hidden: true });
	colMode.push({ name: 'ConditionID', hidden: true });

	colNames.push('ID');
	colNames.push('Field');
	colNames.push('Operator');
	colNames.push('Value');
	colNames.push('ColumnID');
	colNames.push('ConditionID');

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

		//usefull jquery snippets for formatting.
		$('table').attr('border', '0'); //Change 0 to 1 to show borders.

		$(".datepicker").datepicker();
		$(document).on('keydown', '.datepicker', function (event) {

			switch (event.keyCode) {
				case 113:
					$(this).datepicker("setDate", new Date());
					$(this).datepicker('widget').hide('true');
					break;
			}
		});

		
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
						if (iColumnID == frmFilterForm.selectColumn.options[iLoop].value) {
							fFound = true;
							sColumnName = frmFilterForm.selectColumn.options[iLoop].text;
							break;
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

					//Determine if the grid already exists...
					if ($("#DBGridFilterRecords").getGridParam("reccount") == undefined) { //It doesn't exist, create it
						$("#DBGridFilterRecords").jqGrid({
							multiselect: false,
							data: colData,
							datatype: 'local',
							colNames: colNames,
							colModel: colMode,
							rowNum: 1000,
							autowidth: true,
							shrinkToFit: true,
							onSelectRow: function () {
								button_disable(frmFilterForm.cmdRemoveAll, false);
								button_disable(frmFilterForm.cmdRemove, false);
							},
							editurl: 'clientArray'
						}).jqGrid('hideCol', 'cb');
					}

					var items = sAddString.split("\t");
					$("#DBGridFilterRecords").addRowData(
							$("#DBGridFilterRecords").getGridParam("reccount") + 1, //ID
							{ //Data
								'Field': items[0],
								'Operator': items[1],
								'Value': items[2],
								'ColumnID': items[3],
								'ConditionID': items[4]
							},
							'last'); //Add the record at the end

					//Select the newly added record
					$("#DBGridFilterRecords").jqGrid('setSelection', $("#DBGridFilterRecords").getGridParam("reccount"));

					FilterSelect_refreshControls();
				}
			}

			//Select the top filter record
			if ($("#DBGridFilterRecords").getGridParam("reccount") > 0) {
				$("#DBGridFilterRecords").jqGrid('setSelection', 1);
			}

			//Determine if the grid already exists...
			if ($("#DBGridFilterRecords").getGridParam("reccount") == undefined) { //It doesn't exist, create it
				$("#DBGridFilterRecords").jqGrid({
					multiselect: false,
					data: colData,
					datatype: 'local',
					colNames: colNames,
					colModel: colMode,
					rowNum: 1000,
					autowidth: true,
					shrinkToFit: true,
					onSelectRow: function() {
						button_disable(frmFilterForm.cmdRemoveAll, false);
						button_disable(frmFilterForm.cmdRemove, false);
					},
					editurl: 'clientArray'
				}).jqGrid('hideCol', 'cb');
			}

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
			FilterSelect_refreshControls();
		}

	}
</script>

<script type="text/javascript">
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
		sDecimalSeparator = "";
		sDecimalSeparator = '<%:LocaleDecimalSeparator()%>';

		var sApostrophe = "'";
		var sStar = "*";
		var sQuestion = "?";

		sFilterSQL = "";
		sFilterDef = "";
		sRealSource = frmFilterForm.txtRealSource.value;
		sRealSource = sRealSource.concat(".");

		// Loop through the grid records, building the filter code for each record.
		var allIDs = $('#DBGridFilterRecords').getDataIDs();
		var rowData;
		for (iIndex = 0; iIndex < allIDs.length; iIndex++) {
			rowData = $('#DBGridFilterRecords').getRowData(allIDs[iIndex]);
			// Get the column name & id and the value used in the filter operation.
			sColumnName = rowData.Field;
			sColumnName = OpenHR.replaceAll(sColumnName, " ", "_");
			sColumnName = sRealSource.concat(sColumnName);

			sValue = rowData.Value;
			iColumnID = rowData.ColumnID;
			iOperatorID = rowData.ConditionID;

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
					sFilterValue = OpenHR.replaceAll(sValue, sDecimalSeparator, ".");

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
						sFilterValue = OpenHR.convertLocaleDateToSQL(sValue);
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
							sModifiedFilterValue = OpenHR.replaceAll(sValue, sApostrophe, "''");
							sModifiedFilterValue = OpenHR.replaceAll(sModifiedFilterValue, sStar, "%");
							sModifiedFilterValue = OpenHR.replaceAll(sModifiedFilterValue, sQuestion, "_");

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
							sModifiedFilterValue = OpenHR.replaceAll(sValue, sApostrophe, "''");
							sModifiedFilterValue = OpenHR.replaceAll(sModifiedFilterValue, sStar, "%");
							sModifiedFilterValue = OpenHR.replaceAll(sModifiedFilterValue, sQuestion, "_");

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
							sModifiedFilterValue = OpenHR.replaceAll(sValue, sApostrophe, "''");

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
							sModifiedFilterValue = OpenHR.replaceAll(sValue, sApostrophe, "''");

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
					sFilterDef = sFilterDef.concat(((iDataType == 2) || (iDataType == 4)) ? sFilterValue : sValue);
					sFilterDef = sFilterDef.concat("	");
				}
			}
		}

		var postData = {
			Action: optionActionType.SELECTFILTER,
			ScreenID: <%:Session("optionScreenID")%>,
			TableID: <%:Session("optionTableID")%>,
			ViewID: <%:Session("optionViewID")%>,
			FilterSQL: sFilterSQL,
			FilterDef: sFilterDef,
			<%:Html.AntiForgeryTokenForAjaxPost() %> };
		OpenHR.submitForm(null, "optionframe", null, postData, "filterselect_Submit");

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

		var postData = {
			Action: optionActionType.CANCEL,
			<%:Html.AntiForgeryTokenForAjaxPost() %> };
		OpenHR.submitForm(null, "optionframe", null, postData, "filterselect_Submit");
		
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

		sDecimalSeparator = '<%:LocaleDecimalSeparator()%>';
		sThousandSeparator = '<%:LocaleThousandSeparator()%>';		
		sPoint = ".";

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
				sConvertedValue = OpenHR.replaceAll(sConvertedValue, sThousandSeparator, "");
				sValue = sConvertedValue;

				// Convert any decimal separators to '.'.
				if ('<%:LocaleDecimalSeparator()%>' != ".") {
					// Existing decimal points are invalid characters.
					sConvertedValue = OpenHR.replaceAll(sConvertedValue, sPoint, "A");
					// Replace the locale decimal marker with the decimal point.										
					sConvertedValue = OpenHR.replaceAll(sConvertedValue, sDecimalSeparator, ".");
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
				sValue = frmFilterForm.selectDate.value;

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
					if (OpenHR.convertLocaleDateToSQL(sValue) == "null") {
						fOK = false;
						OpenHR.messageBox("Invalid date value entered.");
						frmFilterForm.txtValue.focus();
					}
					else {
//						sValue = OpenHR.convertLocaleDateToSQL(sValue);

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
			var items = sAddString.split("\t");
			$("#DBGridFilterRecords").addRowData(
					$("#DBGridFilterRecords").getGridParam("reccount") + 1, //ID
					{ //Data
						'Field': items[0],
						'Operator': items[1],
						'Value': items[2],
						'ColumnID': items[3],
						'ConditionID': items[4]
					},
					'last'); //Add the record at the end

			//Select the newly added record
			$("#DBGridFilterRecords").jqGrid('setSelection', $("#DBGridFilterRecords").getGridParam("reccount"));

			FilterSelect_refreshControls();
			frmFilterForm.selectColumn.focus();
		}
	}

	function displayOperatorSelector(piDataType) {		
		var frmFilterForm = document.getElementById("frmFilterForm");

		if (piDataType == -7) {
			// Display the logic operator control.
			frmFilterForm.txtConditionLogic.style.width = "100%";
			frmFilterForm.txtConditionLogic.style.visibility = "";
			frmFilterForm.txtConditionLogic.style.position = "";
			frmFilterForm.txtConditionLogic.style.top = "";
			frmFilterForm.txtConditionLogic.style.left = "";

			frmFilterForm.selectValue.style.width = "25%";
			frmFilterForm.selectValue.style.visibility = "";
			frmFilterForm.selectValue.style.position = "";
			frmFilterForm.selectValue.style.top = "";
			frmFilterForm.selectValue.style.left = "";

			frmFilterForm.selectDate.style.width = "0px";
			frmFilterForm.selectDate.style.visibility = "hidden";
			frmFilterForm.selectDate.style.position = "absolute";
			frmFilterForm.selectDate.style.top = 0;
			frmFilterForm.selectDate.style.left = 0;

			frmFilterForm.txtValue.style.width = "0px";
			frmFilterForm.txtValue.style.visibility = "hidden";
			frmFilterForm.txtValue.style.position = "absolute";
			frmFilterForm.txtValue.style.top = 0;
			frmFilterForm.txtValue.style.left = 0;

		} else if (piDataType == 11) {

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

			frmFilterForm.selectDate.style.width = "100%";
			frmFilterForm.selectDate.style.visibility = "";
			frmFilterForm.selectDate.style.position = "";
			frmFilterForm.selectDate.style.top = "";
			frmFilterForm.selectDate.style.left = "";
			
			frmFilterForm.txtValue.style.width = "0px";
			frmFilterForm.txtValue.style.visibility = "hidden";
			frmFilterForm.txtValue.style.position = "absolute";
			frmFilterForm.txtValue.style.top = 0;
			frmFilterForm.txtValue.style.left = 0;

		} else {
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

			frmFilterForm.selectDate.style.width = "0px";
			frmFilterForm.selectDate.style.visibility = "hidden";
			frmFilterForm.selectDate.style.position = "absolute";
			frmFilterForm.selectDate.style.top = 0;
			frmFilterForm.selectDate.style.left = 0;

			frmFilterForm.txtValue.style.width = "100%";
			frmFilterForm.txtValue.style.visibility = "";
			frmFilterForm.txtValue.style.position = "";
			frmFilterForm.txtValue.style.top = "";
			frmFilterForm.txtValue.style.left = "";
		}

		if ((piDataType == 2) || (piDataType == 4)) {
			// Display the Numeric/Integer operator control.
			frmFilterForm.selectConditionNum.style.width = "100%";
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
			frmFilterForm.selectConditionDate.style.width = "100%";
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
			frmFilterForm.selectConditionChar.style.width = "100%";
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


	function FilterSelect_removeAll() {
		$("#DBGridFilterRecords").jqGrid('clearGridData');
		FilterSelect_refreshControls();
	}

	function FilterSelect_remove() {
		var frmFilterForm = document.getElementById("frmFilterForm");

		if ($("#DBGridFilterRecords").getGridParam("reccount") > 0) {
			var iRowIndex = $("#DBGridFilterRecords").jqGrid('getCell', $('#DBGridFilterRecords').jqGrid('getGridParam', 'selrow'), 5); //5 -> ID

			if (($("#DBGridFilterRecords").getGridParam("reccount") == 1) && (iRowIndex == 0)) {
				FilterSelect_removeAll();
			}
			else {
				var grid = $("#DBGridFilterRecords");
				var myDelOptions = {
					// because I use "local" data I don't want to send the changes
					// to the server so I use "processing:true" setting and delete
					// the row manually in onclickSubmit
					onclickSubmit: function (options) {
						var grid_id = $.jgrid.jqID(grid[0].id),
								grid_p = grid[0].p,
								newPage = grid_p.page,
								rowids = grid_p.multiselect ? grid_p.selarrrow : [grid_p.selrow];

						// reset the value of processing option which could be modified
						options.processing = true;

						// delete the row
						$.each(rowids, function () {
							grid.delRowData(this);
						});
						$.jgrid.hideModal("#delmod" + grid_id,
															{
																gb: "#gbox_" + grid_id,
																jqm: options.jqModal, onClose: options.onClose
															});

						if (grid_p.lastpage > 1) {// on the multipage grid reload the grid
							if (grid_p.reccount === 0 && newPage === grid_p.lastpage) {
								// if after deliting there are no rows on the current page
								// which is the last page of the grid
								newPage--; // go to the previous page
							}
							// reload grid to make the row from the next page visable.
							grid.trigger("reloadGrid", [{ page: newPage }]);
						}
						return true;
					},
					processing: true
				};

				grid.jqGrid('delGridRow', grid.jqGrid('getGridParam', 'selarrrow'), myDelOptions);

				$("#dData").click(); //To remove the "delete confirmation" dialog
			}
		}

		FilterSelect_refreshControls();
	}

	function FilterSelect_refreshControls() {
		var frmFilterForm = document.getElementById("frmFilterForm");

		if ($("#DBGridFilterRecords").getGridParam("reccount") > 0) {
			button_disable(frmFilterForm.cmdRemoveAll, false);

			if ($('#DBGridFilterRecords').jqGrid('getGridParam', 'selrow') != null && $('#DBGridFilterRecords').jqGrid('getGridParam', 'selrow').length > 0) {
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

		var sCurrentPage = $("#workframe").attr("data-framesource");

		return (sCurrentPage);
	}

</script>

<div <%=session("BodyTag")%> style="padding: 10px 0px 0px 20px;">
	<form id="frmFilterForm" name="frmFilterForm">

		<div class="pageTitleDiv" style="margin-bottom: 15px">
			<span class="pageTitle" id="PopupReportDefinition_PageTitle">Define Filter</span>
		</div>

		<table style="text-align: center; margin: 0px auto; border-spacing: 5px; border-collapse: collapse; width: 100%; height: 100%;" class="outline">
			<tr>
				<td>
					<table id="filterTable" style="width: 100%; height: 100%; border-spacing: 0px; border-collapse: collapse;" class="invisible">
						<tr style="height: 160px">
							<td style="width: 10px"></td>
							<td>
								<div id="FilterRecordsGrid">
									<table id="DBGridFilterRecords"></table>
								</div>
							</td>
							<td style="width: 10px"></td>
						</tr>

						<tr>
							<td style="width: 10px"></td>
							<td style="height: 10px">
								<table class="invisible" style="width: 100%; border-spacing: 5px; border-collapse: collapse;">
									<tr style="height: 10px">
										<td style="height: 10px">&nbsp;
										</td>
										<td style="width: 10px; height: 10px">
											<input id="cmdRemove" name="cmdRemove" type="button" value="Remove" style="width: 100px; margin-top: 10px;" class="btn"
												onclick="FilterSelect_remove()" />
										</td>
										<td style="width: 10px; height: 10px"></td>
										<td style="width: 10px; height: 10px">
											<input id="cmdRemoveAll" name="cmdRemoveAll" type="button" value="Remove All" style="width: 100px; margin-top: 10px;" class="btn"
												onclick="FilterSelect_removeAll()" />
										</td>
									</tr>
								</table>
							</td>
							<td style="width: 10px"></td>
						</tr>

						<tr style="height: 10px">
							<td style="width: 10px"></td>
							<td style="height: 10px; text-align: left">
								<strong>Define more criteria :</strong>
							</td>
							<td style="width: 10px"></td>
						</tr>

						<tr>
							<td style="width: 10px"></td>
							<td>
								<table class="invisible" style="width: 100%; border-spacing: 5px; border-collapse: collapse;">
									<tr style="">
										<td style="width: 1%;"></td>
										<td style="width: 30%; text-align: left">Field :</td>
										<td style="width: 1%;"></td>
										<td style="width: 20%; text-align: left">Operator :</td>
										<td style="width: 1%;"></td>
										<td style="width: 35%; text-align: left">Value :</td>
										<td style="width: 1%;"></td>
										<td style="width: 100%;"></td>
									</tr>

									<tr style="">
										<td></td>
										<td>
											<select id="selectColumn" name="selectColumn" class="combo" style="width: 100%" onchange="refreshOperatorCombo()">
												<%
													' Populate the columns combo.
													Dim iCount
													Dim sErrorDescription = ""
													Dim dtFilterColumns As DataTable
													If Len(sErrorDescription) = 0 Then
														' Get the column records.
														Try
															Dim objSession As SessionInfo = CType(HttpContext.Current.Session("SessionContext"), SessionInfo)	'Set session info
															Dim objDataAccess As New clsDataAccess(objSession.LoginInfo) 'Instantiate DataAccess class
															Dim psRealSource As New SqlParameter("@psRealSource", SqlDbType.VarChar) With {.Direction = ParameterDirection.Output, .Size = 8000}
															dtFilterColumns = objDataAccess.GetDataTable("sp_ASRIntGetFilterColumns", _
																											CommandType.StoredProcedure, _
																											New SqlParameter("@plngTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionTableID"))}, _
																											New SqlParameter("@plngViewID ", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionViewID"))}, _
																											psRealSource _
																											)
	
															iCount = 0
															For Each dr As DataRow In dtFilterColumns.Rows
																Response.Write("						<OPTION value=" & dr(0).ToString)
																If iCount = 0 Then
																	Response.Write(" SELECTED")
																End If
																Response.Write(">" & Replace(dr(1).ToString, "_", " ") & "</OPTION>" & vbCrLf)
																iCount = iCount + 1
															Next
												%>
											</select>
											<%
												Response.Write("<INPUT type='hidden' id=txtRealSource name=txtRealSource value=""" & Replace(Replace(psRealSource.Value, "'", "'''"), """", "&quot;") & """>" & vbCrLf)
			
												For Each dr As DataRow In dtFilterColumns.Rows
													Response.Write("					<INPUT type='hidden' id=txtFilterColumn_" & dr(0).ToString & " name=txtFilterColumn_" & dr(0).ToString & " value=" & dr(2).ToString & ">")
													Response.Write("					<INPUT type='hidden' id=txtFilterColumnSize_" & dr(0).ToString & " name=txtFilterColumnSize_" & dr(0).ToString & " value=" & dr("size").ToString & ">")
													Response.Write("					<INPUT type='hidden' id=txtFilterColumnDecimals_" & dr(0).ToString & " name=txtFilterColumnDecimals_" & dr(0).ToString & " value=" & dr("decimals").ToString & ">")
												Next
			
											Catch ex As Exception
												sErrorDescription = "The filter columns could not be retrieved." & vbCrLf & FormatError(ex.Message)
											End Try
										End If
											%>
										</td>
										<td></td>
										<td>
											<input type="text" id="txtConditionLogic" class="text textdisabled" name="selectConditionLogic" disabled="disabled" value="is equal to">

											<select id="selectConditionDate" name="selectConditionDate" class="combo" style="width: 100%; left: 0; position: absolute; top: 0; visibility: hidden;">
												<option value="1">is equal to</option>
												<option value="2">is NOT equal to</option>
												<option value="5">after</option>
												<option value="6">before</option>
												<option value="4">is equal to or after</option>
												<option value="3">is equal to or before</option>
											</select>

											<select id="selectConditionNum" name="selectConditionNum" class="combo" style="width: 100%; left: 0; position: absolute; top: 0; visibility: hidden;">
												<option value="1">is equal to</option>
												<option value="2">is NOT equal to</option>
												<option value="5">is greater than</option>
												<option value="4">is greater than or equal to</option>
												<option value="6">is less than</option>
												<option value="3">is less than or equal to</option>
											</select>

											<select id="selectConditionChar" name="selectConditionChar" class="combo" style="width: 100%; left: 0; position: absolute; top: 0; visibility: hidden;">
												<option value="1">is equal to</option>
												<option value="2">is NOT equal to</option>
												<option value="7">contains</option>
												<option value="8">does not contain</option>
											</select>
										</td>
										<td></td>
										<td style="text-align: left">
											<select id="selectValue" name='selectValue"' class="combo" style="width: 30%;">
												<option value="1">True</option>
												<option value="0">False</option>
											</select>
											<input id="txtValue" name="txtValue" class="text" style="width: 100%; left: 0; position: absolute; top: 0; visibility: hidden;">
											<input id="selectDate" name="selectDate" type="text" class="datepicker" style="width: 100%; left: 0; position: absolute; top: 0; visibility: hidden;" />
										</td>
										<td></td>
										<td style="width: 175px; text-align: right">
											<input id="cmdAddToList" name="cmdAddToList" class="btn" type="button" value="Add To List" style="width: 100px"
												onclick="AddToList()" />
										</td>
									</tr>

									<tr>
										<td colspan="8" style="height: 50px"></td>
									</tr>
								</table>
							</td>
							<td style="width: 10px"></td>
						</tr>

						<tr style="height: 10px">
							<td style="width: 10px"></td>
							<td style="height: 10px">
								<table class="invisible" style="border-spacing: 0; border-collapse: collapse; width: 100%">
									<tr>
										<td colspan = 8>
										<div class="floatright">
											<input id="cmdSelect" name="cmdSelect" class="btn" type="button" value="OK"
												onclick="SelectFilter()"
												style="width: 100px; margin-right:10px"  />
											<input id="cmdCancel" name="cmdCancel" class="btn" type="button" value="Cancel" 
												onclick="CancelFilter()"
												style="width: 100px"  />
										</div>
											</td>
									</tr>
								</table>
							</td>
							<td style="width: 10px"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>



		<%
			Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtOptionFilterDef name=txtOptionFilterDef value=""" & Replace(Session("optionFilterDef"), """", "&quot;") & """>" & vbCrLf)
		%>
	</form>

	<script type="text/javascript">
		filterselect_window_onload()		
	</script>
</div>
