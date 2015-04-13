<%@ control language="VB" inherits="System.Web.Mvc.ViewUserControl" %>
<%@ import namespace="DMI.NET" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="DMI.NET.Helpers" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<%
	dim sErrorDescription as String = ""
%>

<script type="text/javascript">
	function quickfind_window_onload() {
		var fOK;
		fOK = true;
		var frmQuickFindForm = document.getElementById("frmQuickFindForm");

		var sErrMsg = frmQuickFindForm.txtErrorDescription.value;
		if (sErrMsg.length > 0) {
			fOK = false;
			OpenHR.messageBox(sErrMsg);
			window.parent.location.replace("login");
		}

		if (fOK == true) {
			var sMsg = frmQuickFindForm.txtOptionMessage.value;
			if (sMsg.length > 0) {
				OpenHR.messageBox(sMsg);
			}

			// Expand the option frame and hide the work frame.
			//window.parent.document.all.item("workframeset").cols = "0, *";
			$("#optionframe").attr("data-framesource", "QUICKFIND");
			$("#workframe").hide();
			$("#optionframe").show();

			// Set focus onto one of the form controls. 
			// NB. This needs to be done before making any reference to the grid
			if (frmQuickFindForm.selectField == null) {
				frmQuickFindForm.cmdCancel.focus();
			} else {
				frmQuickFindForm.selectField.focus();
			}
		}
	}

	function selectQuickFind() {
		var frmQuickFindForm = document.getElementById("frmQuickFindForm");

		var fOK;
		var fSizeFound;
		var fDecimalsFound;
		var fDataTypeFound;
		var iDataType;
		var iSize;
		var sSize;
		var iDecimals;
		var sDecimals;
		var iIndex;
		var sValue;
		var sControlName;
		var sModifiedValue;
		var sReqdSizeControlName;
		var sReqdDecimalsControlName;
		var sReqdDataTypeControlName;
		var controlCollection = frmQuickFindForm.elements;
		var sDecimalSeparator;
		var sPoint;
		var sConvertedValue;
		var iTempSize;
		var iTempDecimals;

		// Create some regular expressions to be used when replacing characters 
		// in the filter string later on.
		sDecimalSeparator = "\\";
		sDecimalSeparator = OpenHR.LocaleDecimalSeparator();
		var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");

		var sThousandSeparator = "\\";
		sThousandSeparator = sThousandSeparator.concat(OpenHR.LocaleThousandSeparator());
		var reThousandSeparator = new RegExp(sThousandSeparator, "gi");

		sPoint = "\\.";
		var rePoint = new RegExp(sPoint, "gi");

		fOK = false;

		// Start to construct the SQL code to get the id of
		// the record matching the entered quick find criteria.	

		fSizeFound = false;
		fDecimalsFound = false;
		fDataTypeFound = false;
		iSize = 0;
		iDecimals = 0;
		iDataType = 12;
		sReqdDataTypeControlName = "txtColumnDataType_";
		sReqdDataTypeControlName = sReqdDataTypeControlName.concat(frmQuickFindForm.selectField.options[frmQuickFindForm.selectField.selectedIndex].value);
		sReqdSizeControlName = "txtColumnSize_";
		sReqdSizeControlName = sReqdSizeControlName.concat(frmQuickFindForm.selectField.options[frmQuickFindForm.selectField.selectedIndex].value);
		sReqdDecimalsControlName = "txtColumnDecimals_";
		sReqdDecimalsControlName = sReqdDecimalsControlName.concat(frmQuickFindForm.selectField.options[frmQuickFindForm.selectField.selectedIndex].value);

		// Determine the data type, size and decimals of the quick find column.
		if (controlCollection != null) {
			for (var i = 0; i < controlCollection.length; i++) {
				sControlName = controlCollection.item(i).name;

				if (fSizeFound == false) {
					if (sControlName == sReqdSizeControlName) {
						// Get the string version of the column size.
						sSize = controlCollection.item(i).value;
						// Get the numeric version of the column size.
						// This has to be done as we'll be adding 1 to it later, 
						// and adding 1 to the string version just concatenates '1' onto it.
						iSize = new Number(sSize);
						fSizeFound = true;
					}
				}

				if (fDecimalsFound == false) {
					if (sControlName == sReqdDecimalsControlName) {
						// Get the string version of the column decimals.
						sDecimals = controlCollection.item(i).value;
						// Get the numeric version of the column decimals.
						iDecimals = new Number(sDecimals);
						fDecimalsFound = true;
					}
				}

				if (fDataTypeFound == false) {
					if (sControlName == sReqdDataTypeControlName) {
						iDataType = controlCollection.item(i).value;
						fDataTypeFound = true;
					}
				}

				if ((fSizeFound == true) && (fDataTypeFound == true) && (fDecimalsFound == true)) {
					fOK = true;
					break;
				}
			}
		}

		if ((fOK == true) && (iDataType == 2)) {
			// Numeric column.
			// Ensure that the value entered is numeric.
			sValue = frmQuickFindForm.txtValue.value;
			if (sValue.length == 0) {
				sValue = "0";
			}

			// Convert the value from locale to UK settings for use with the isNaN funtion.
			sConvertedValue = new String(sValue);
			// Remove any thousand separators.
			sConvertedValue = sConvertedValue.replace(reThousandSeparator, "");

			// Convert any decimal separators to '.'.
			if (OpenHR.LocaleDecimalSeparator() != ".") {
				// Remove decimal points.
				sConvertedValue = sConvertedValue.replace(rePoint, "A");
				// replace the locale decimal marker with the decimal point.
				sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
			}

			if (isNaN(sConvertedValue) == true) {
				fOK = false;
				OpenHR.messageBox("Invalid numeric value entered.");
				frmQuickFindForm.txtValue.focus();
			}
			else {
				// Ensure the value is not too big for the selected column.
				iIndex = sConvertedValue.indexOf(".");
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
					frmQuickFindForm.txtValue.focus();
				}
				else {
					if (iTempDecimals > iDecimals) {
						fOK = false;
						OpenHR.messageBox("The column can only be compared to values with " + iDecimals + " decimal place(s).");
						frmQuickFindForm.txtValue.focus();
					}
				}

				if (fOK == true) {
					// Construct the SQL code for getting the record with the entered unique value.
					sModifiedValue = sConvertedValue;
				}
			}
		}

		if ((fOK == true) && (iDataType == 4)) {
			// Integer column.
			// Ensure that the value entered is numeric.
			sValue = frmQuickFindForm.txtValue.value;
			if (sValue.length == 0) {
				sValue = "0";
			}

			// Convert the value from locale to UK settings for use with the isNaN funtion.
			sConvertedValue = new String(sValue);
			// Remove any thousand separators.
			sConvertedValue = sConvertedValue.replace(reThousandSeparator, "");
			sValue = sConvertedValue;

			// Convert any decimal separators to '.'.
			if (OpenHR.LocaleDecimalSeparator() != ".") {
				// Remove decimal points.
				sConvertedValue = sConvertedValue.replace(rePoint, "A");
				// replace the locale decimal marker with the decimal point.
				sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
			}

			if (isNaN(sConvertedValue) == true) {
				fOK = false;
				OpenHR.messageBox("Invalid integer value entered.");
				frmQuickFindForm.txtValue.focus();
			}
			else {
				// Ensure the value is not too big for the selected column.
				iIndex = sConvertedValue.indexOf(".");
				if (iIndex >= 0) {
					fOK = false;
					OpenHR.messageBox("Invalid integer value entered.");
					frmQuickFindForm.txtValue.focus();
				}
				else {
					// Construct the SQL code for getting the record with the entered unique value.
					sModifiedValue = sConvertedValue;
				}
			}
		}

		if ((fOK == true) && (iDataType == 11)) {
			// Date column.
			// Ensure that the value entered is a date.
			sValue = frmQuickFindForm.txtValue.value;

			if (sValue.length == 0) {
				sModifiedValue = "";
			}
			else {
				// Convert the date to SQL format (use this as a validation check).
				// An empty string is returned if the date is invalid.
				sValue = OpenHR.convertLocaleDateToSQL(sValue);		//TODO: empty function.
				if (sValue.length == 0) {
					fOK = false;
					OpenHR.messageBox("Invalid date value entered.");
					frmQuickFindForm.txtValue.focus();
				}
				else {
					sModifiedValue = sValue;
				}
			}
		}

		if ((fOK == true) && (iDataType == 12)) {
			// Character column.
			sValue = frmQuickFindForm.txtValue.value;

			// Construct the SQL code for getting the record with the entered unique value.
			sModifiedValue = sValue;
		}

		if (fOK == true) {

			var postData = {
				Action: optionActionType.QUICKFIND,
				TableID: frmQuickFindForm.txtOptionTableID.value,
				ViewID: frmQuickFindForm.txtOptionViewID.value,
				FilterSQL: frmQuickFindForm.txtOptionFilterSQL.value,
				FilterDef: frmQuickFindForm.txtOptionFilterDef.value,
				Value: sModifiedValue,
				ColumnID: frmQuickFindForm.selectField.options[frmQuickFindForm.selectField.selectedIndex].value,
				<%:Html.AntiForgeryTokenForAjaxPost() %>
			};
			OpenHR.submitForm(null, "optionframe", null, postData, "quickfind_Submit");

		}
	}

	function CancelQuickFind() {
		// Redisplay the workframe recedit control. 
		$("#workframe").attr("data-framesource", "RECORDEDIT");
		$("#optionframe").hide();
		$("#workframe").show();

		refreshData();	//recedit

		var postData = {
			Action: optionActionType.CANCEL,
				<%:Html.AntiForgeryTokenForAjaxPost() %>
			};
		OpenHR.submitForm(null, "optionframe", null, postData, "quickfind_Submit");

	}

</script>

<form action="" method="POST" id="frmQuickFindForm" name="frmQuickFindForm" onsubmit="return false;">

		<table style="margin: 0 auto; text-align: center; border-spacing: 5px; border-collapse: collapse;" class="outline">
			<tr>
				<td>
					<table style="text-align: center; border-spacing: 0; border-collapse: collapse;" class="invisible">
						<tr>
							<td colspan="3" height="10"></td>
						</tr>
						<tr>
							<td colspan="3">
								<h3 align="center">Quick Find</h3>
							</td>
						</tr>

	<%

		' Create the table row with the Field selection combo.
		' If no valid fields exist then display a message telling the user why
		' no columns are valid.		

		Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
									
		Dim prmRealSource = New SqlParameter("psRealSource", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
		Dim rstColumns As DataTable
		
		Try
			
			rstColumns = objDataAccess.GetFromSP("sp_ASRIntGetUniqueColumns" _
				, New SqlParameter("plngTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionTableID"))} _
				, New SqlParameter("plngViewID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionViewID"))} _
			, prmRealSource)
							
		If rstColumns.Rows.Count = 0 Then
						%>
						<tr>
							<td width="20"></td>
							<td style="text-align: center;">Quick Find can only be used on tables with columns defined as unique. 
							</td>
							<td width="20"></td>
						</tr>
						<tr>
							<td width="20"></td>
							<td style="text-align: center;">The current table has no unique columns. 
							</td>
							<td width="20"></td>
						</tr>
						<tr>
							<td height="20" colspan="3"></td>
						</tr>
						<tr>
							<td width="20"></td>
							<td style="text-align: center;">
								<input id="cmdCancel" name="cmdCancel" class="btn" type="button" value="Cancel" style="WIDTH: 75px" width="75"
									onclick="CancelQuickFind()" />
							</td>
							<td width="20"></td>
						</tr>
						<tr>
							<td height="10" colspan="3"></td>
						</tr>
						<%
			else 
						%>
						<tr>
							<td width="20"></td>
							<td>

								<table style="width: 100%; border-spacing: 0; border-collapse: collapse;" class="invisible">
									<tr height="10">
										<td width="50" height="10">Field :
										</td>
										<td width="10" height="10"></td>
										<td width="175" height="10">
											<select id="selectField" name="selectField" class="combo" style="HEIGHT: 22px; WIDTH: 200px">
												<%
													Dim iCount = 0
													For Each objRow As DataRow In rstColumns.Rows
														Response.Write("						<option value=" & objRow(0).ToString())
														If iCount = 0 Then
															Response.Write(" selected")
														End If
				
														Response.Write(">" & Replace(objRow(1).ToString(), "_", " ") & "</option>" & vbCrLf)
														iCount += 1
													Next
												%>
											</select>
										</td>
									</tr>
									<tr height="10">
										<td height="10" colspan="3"></td>
									</tr>

									<tr height="10">
										<td width="50" height="10">Value :
										</td>
										<td width="10" height="10"></td>
										<td width="175" height="10">
											<input id="txtValue" name="txtValue" class="text" style="HEIGHT: 22px; WIDTH: 200px">
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr height="20">
							<td colspan="3" height="20"></td>
						</tr>
						<tr height="10">
							<td width="20"></td>
							<td height="10" style="text-align: center;">
								<table style="width: 100%; border: 0; border-spacing: 0; border-collapse: collapse;">
									<tr height="10">
										<td>&nbsp;</td>
										<td width="10" height="10">
											<input id="cmdSelect" name="cmdSelect" class="btn" type="button" value="Find" style="WIDTH: 75px" width="75"
												onclick="selectQuickFind()" />
										</td>
										<td width="10" height="10"></td>
										<td width="10" height="10">
											<input id="cmdCancel" name="cmdCancel" class="btn" type="button" value="Cancel" style="WIDTH: 75px" width="75"
												onclick="CancelQuickFind()" />
										</td>
										<td>&nbsp;</td>
									</tr>
								</table>
							</td>
							<td width="20"></td>
						</tr>
						<tr height="20">
							<td colspan="3" height="10"></td>
						</tr>
			<%
			
			End If

		
			For Each objRow As DataRow In rstColumns.Rows
				Response.Write("					<input type='hidden' id=txtColumnDataType_" & objRow(0).ToString() & " name=txtColumnDataType_" & objRow(0).ToString() & " value=" & objRow(2).ToString() & ">")
				Response.Write("					<input type='hidden' id=txtColumnSize_" & objRow(0).ToString() & " name=txtColumnSize_" & objRow(0).ToString() & " value=" & objRow(3).ToString() & ">")
				Response.Write("					<input type='hidden' id=txtColumnDecimals_" & objRow(0).ToString() & " name=txtColumnDecimals_" & objRow(0).ToString() & " value=" & objRow(4).ToString() & ">")
			Next
				
			Response.Write("<input type='hidden' id=txtRealSource name=txtRealSource value=""" & Replace(Replace(prmRealSource.Value.ToString(), "'", "'''"), """", "&quot;") & """>" & vbCrLf)

		Catch ex As Exception
			sErrorDescription = "The unique fields could not be retrieved." & vbCrLf & FormatError(ex.Message)

		End Try


						%>
					</table>
				</td>
			</tr>
		</table>

		<input type='hidden' id="txtErrorDescription" name="txtErrorDescription" value="<%=sErrorDescription%>">
		<input type='hidden' id="txtOptionScreenID" name="txtOptionScreenID" value='<%=session("optionScreenID")%>'>
		<input type='hidden' id="txtOptionTableID" name="txtOptionTableID" value='<%=session("optionTableID")%>'>
		<input type='hidden' id="txtOptionViewID" name="txtOptionViewID" value='<%=session("optionViewID")%>'>
		<input type='hidden' id="txtOptionFilterSQL" name="txtOptionFilterSQL" value="<%=replace(session("optionFilterSQL"), """", "&quot;")%>">
		<input type='hidden' id="txtOptionFilterDef" name="txtOptionFilterDef" value="<%=replace(session("optionFilterDef"), """", "&quot;")%>">
		<input type='hidden' id="txtOptionMessage" name="txtOptionMessage" value="<%=replace(session("errorMessage"), """", "&quot;")%>">
	</form>

<script type="text/javascript"> quickfind_window_onload();</script>
