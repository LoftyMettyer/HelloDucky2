<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@Import namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<%
	Response.Expires = -1
	Dim sErrorDescription = ""
	Dim sFailureDescription = ""
%>

<script src="<%: Url.LatestContent("~/Scripts/ctl_SetFont.js")%>" type="text/javascript"></script>

<script type="text/javascript">

	function tbBulkBookingSelection_onload() {		
		var fOK = true;
		var frmUseful = document.getElementById("frmUseful");

		$('input[type=button]').button();	//jquery style the buttons.

		if ((frmUseful.txtSelectionType.value.toUpperCase() != "FILTER") &&
			(frmUseful.txtSelectionType.value.toUpperCase() != "PICKLIST")) {

			var sErrMsg = document.getElementById("txtErrorDescription").value;
			if (sErrMsg.length > 0) {
				fOK = false;
				OpenHR.messageBox(sErrMsg);
				window.parent.close();
			}

			if (fOK == true) {
				if (selectView.length == 0) {
					fOK = false;
					OpenHR.messageBox("You do not have permission to read the employee table.");
					window.parent.close();
				}
			}

			if (fOK == true) {
				if (selectOrder.length == 0) {
					fOK = false;
					OpenHR.messageBox("You do not have permission to use any of the employee table orders.");
					window.parent.close();					
				}
			}

		}

		cmdCancel.focus();

		if ((frmUseful.txtSelectionType.value.toUpperCase() == "FILTER") ||
			(frmUseful.txtSelectionType.value.toUpperCase() == "PICKLIST")) {

			var gridTop = $('#findGridRow').offset().top;
			var gridBottom = $('#trButtons').offset().top;
			var gridWidth = $('body').outerWidth() - 50;
			var gridHeight = gridBottom - gridTop - 30; //30px for the new table header.

			var SelectionType = frmUseful.txtSelectionType.value;
			tableToGrid("#ssOleDBGridSelRecords", {
				width: gridWidth,
				height: gridHeight,
				onSelectRow: function (id) {
					tbrefreshControls();
				},
				ondblClickRow: function () {
					ssOleDBGridSelRecords_dblClick();
				},
				colNames: [SelectionType],
				colModel: [
						{ name: 'name', label: SelectionType, index: 'name', sortable: false }
				]
			});

			tbrefreshControls();
		} else {

			var ssOleDBGridSelRecords = document.getElementById("ssOleDBGridSelRecords");

			window.parent.dialogLeft = new String((screen.width - (9 * screen.width / 10)) / 2) + "px";
			window.parent.dialogTop = new String((screen.height - (3 * screen.height / 4)) / 2) + "px";
			window.parent.dialogWidth = new String((9 * screen.width / 10)) + "px";
			window.parent.dialogHeight = new String((3 * screen.height / 4)) + "px";

			window.parent.txtTableID.value = frmUseful.txtTableID.value;
			window.parent.txtViewID.value = selectView.options[selectView.selectedIndex].value;
			window.parent.txtOrderID.value = selectOrder.options[selectOrder.selectedIndex].value;
			window.parent.loadAddRecords();
		}
	}
</script>

<script type="text/javascript">

	function tbrefreshControls()
	{
		var fNoneSelected;

		var selRowId = $("#ssOleDBGridSelRecords").jqGrid('getGridParam', 'selrow');
		fNoneSelected = (selRowId == null || selRowId == 'undefined');

		button_disable(cmdOK, fNoneSelected);

		var frmUseful = document.getElementById("frmUseful");

		if ((frmUseful.txtSelectionType.value.toUpperCase() != "FILTER") &&
			(frmUseful.txtSelectionType.value.toUpperCase() != "PICKLIST")) {

			if (selectOrder.length <= 1) {
				combo_disable(selectOrder, true);
				button_disable(btnGoOrder, true);
			}

			if (selectView.length <= 1) {
				combo_disable(selectView, true);
				button_disable(btnGoView, true);
			}
		}
	}

	function Selection_makeSelection() {

		var frmUseful = document.getElementById("frmUseful");
		var frmPrompt = document.getElementById("frmPrompt");
		var ssOleDBGridSelRecords = document.getElementById("ssOleDBGridSelRecords");
				
		if (frmUseful.txtSelectionType.value.toUpperCase() == "FILTER") {
			// Go to the prompted values form to get any required prompts. 
			frmPrompt.filterID.value = selectedRecordID();			
			OpenHR.submitForm(frmPrompt);
		}
		else {
			if (frmUseful.txtSelectionType.value.toUpperCase() == "PICKLIST") {
				try {
					window.dialogArguments.makeSelection(frmUseful.txtSelectionType.value, selectedRecordID(), "");
				}
				catch (e) {
				}
			}
			else {

				var sSelectedRows = $('#ssOleDBGridSelRecords').jqGrid('getGridParam', 'selarrrow');

				var sSelectedIDs = "";
				for (var iIndex = 0; iIndex < sSelectedRows.length ; iIndex++) {
					var sRecordID = $("#ssOleDBGridSelRecords").jqGrid('getCell', sSelectedRows[iIndex], 'ID');
					if (sSelectedIDs.length > 0) {
						sSelectedIDs = sSelectedIDs + ",";
					}
					sSelectedIDs = sSelectedIDs + sRecordID;
				}

				try {
					window.dialogArguments.makeSelection(frmUseful.txtSelectionType.value, 0, sSelectedIDs);
				}
				catch (e) {
					
				}
			}
			window.parent.close();
		}
	}

	/* Return the ID of the record selected in the find form. */
	function selectedRecordID() {
		var iRecordID = $("#ssOleDBGridSelRecords").jqGrid('getGridParam', 'selrow');

		return(iRecordID);
	}

	function locateRecord(psSearchFor) {
		var fFound;

		fFound = false;
	
		ssOleDBGridSelRecords.redraw = false;

		ssOleDBGridSelRecords.MoveLast();
		ssOleDBGridSelRecords.MoveFirst();

		ssOleDBGridSelRecords.SelBookmarks.removeall();
	
		for (iIndex = 1; iIndex <= ssOleDBGridSelRecords.rows; iIndex++) 
		{	
			var sGridValue = new String(ssOleDBGridSelRecords.Columns(0).value);
			sGridValue = sGridValue.substr(0, psSearchFor.length).toUpperCase();
			if (sGridValue == psSearchFor.toUpperCase()) 
			{
				ssOleDBGridSelRecords.SelBookmarks.Add(ssOleDBGridSelRecords.Bookmark);
				fFound = true;
				break;
			}

			if (iIndex < ssOleDBGridSelRecords.rows) 
			{
				ssOleDBGridSelRecords.MoveNext();
			}
			else 
			{
				break;
			}
		}

		if ((fFound == false) && (ssOleDBGridSelRecords.rows > 0)) {
			// Select the top row.
			ssOleDBGridSelRecords.MoveFirst();
			ssOleDBGridSelRecords.SelBookmarks.Add(ssOleDBGridSelRecords.Bookmark);
		}

		ssOleDBGridSelRecords.redraw = true;
	}

	function goView() {
		// Get the tbBulkBookingSelectionData.asp to get the find records.
		var dataForm = OpenHR.getForm("dataframe", "frmGetData");
		var frmUseful = document.getElementById("frmUseful");
		dataForm.txtTableID.value = frmUseful.txtTableID.value;
		dataForm.txtViewID.value = selectView.options[selectView.selectedIndex].value;
		dataForm.txtOrderID.value = selectOrder.options[selectOrder.selectedIndex].value;
		dataForm.txtFirstRecPos.value = 1;
		dataForm.txtCurrentRecCount.value = 0;
		dataForm.txtPageAction.value = "LOAD";

		refreshData();
	}

	function goOrder() {
		// Get the tbBulkBookingSelectionData.asp to get the find records.
		var dataForm = OpenHR.getForm("dataframe", "frmGetData");
		var frmUseful = document.getElementById("frmUseful");
		dataForm.txtTableID.value = frmUseful.txtTableID.value;
		dataForm.txtViewID.value = selectView.options[selectView.selectedIndex].value;
		dataForm.txtOrderID.value = selectOrder.options[selectOrder.selectedIndex].value;
		dataForm.txtFirstRecPos.value = 1;
		dataForm.txtCurrentRecCount.value = 0;
		dataForm.txtPageAction.value = "LOAD";

		refreshData();
	}

	function selectedOrderID() {
		return selectOrder.options[selectOrder.selectedIndex].value;
	}

	function selectedViewID() {
		return selectView.options[selectView.selectedIndex].value;
	}

	function reloadPage(psAction, psLocateValue) {
		var sConvertedValue;
		var sDecimalSeparator;
		var sThousandSeparator;
		var sPoint;
		var iIndex ;
		var iTempSize;
		var iTempDecimals;
	
		sDecimalSeparator = "\\";
		sDecimalSeparator = sDecimalSeparator.concat(OpenHR.LocaleDecimalSeparator);
		var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");

		sThousandSeparator = "\\";
		sThousandSeparator = sThousandSeparator.concat(OpenHR.LocaleThousandSeparator);
		var reThousandSeparator = new RegExp(sThousandSeparator, "gi");

		sPoint = "\\.";
		var rePoint = new RegExp(sPoint, "gi");

		var fValidLocateValue = true;

		var dataForm = OpenHR.getForm("dataframe", "frmData");
		var getDataForm = OpenHR.getForm("dataframe", "frmGetData");

		if (psAction == "LOCATE") {
			// Check that the entered value is valid for the first order column type.
			var iDataType = dataForm.txtFirstColumnType.value;

			if ((iDataType == 2) || (iDataType == 4)) {
				// Numeric/Integer column.
				// Ensure that the value entered is numeric.
				if (psLocateValue.length == 0) {
					psLocateValue = "0";
				}

				// Convert the value from locale to UK settings for use with the isNaN funtion.
				sConvertedValue = new String(psLocateValue);
				// Remove any thousand separators.
				sConvertedValue = sConvertedValue.replace(reThousandSeparator, "");
				psLocateValue = sConvertedValue;

				// Convert any decimal separators to '.'.
				if (OpenHR.LocaleDecimalSeparator != ".") {
					// Remove decimal points.
					sConvertedValue = sConvertedValue.replace(rePoint, "A");
					// replace the locale decimal marker with the decimal point.
					sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
				}

				if (isNaN(sConvertedValue) == true) {
					fValidLocateValue = false;
					OpenHR.messageBox("Invalid numeric value entered.");
				}
				else {
					psLocateValue = sConvertedValue;
					iIndex = sConvertedValue.indexOf(".");
					if (iDataType == 4) {
						// Ensure that integer columns are compared with integer values.
						if (iIndex >= 0 ) {
							fValidLocateValue = false;
							OpenHR.messageBox("Invalid integer value entered.");
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
					
						if ((sConvertedValue.substr(0,1) == "+") ||
							(sConvertedValue.substr(0,1) == "-")) {
							iTempSize = iTempSize - 1;
						}

						if(iTempSize > (dataForm.txtFirstColumnSize.value - dataForm.txtFirstColumnDecimals.value)) {
							fValidLocateValue = false;
							OpenHR.messageBox("The value cannot have more than " + (dataForm.txtFirstColumnSize.value - dataForm.txtFirstColumnDecimals.value) + " digit(s) to the left of the decimal separator.");
						}
						else {
							if(iTempDecimals > dataForm.txtFirstColumnDecimals.value) {
								fValidLocateValue = false;
								OpenHR.messageBox("The value cannot have more than " + dataForm.txtFirstColumnDecimals.value + " decimal place(s).");
							}
						}
					}
				}
			}
			else {
				if (iDataType == 11) {
					// Date column.
					// Ensure that the value entered is a date.
					if (psLocateValue.length > 0) {
						// Convert the date to SQL format (use this as a validation check).
						// An empty string is returned if the date is invalid.
						psLocateValue = convertLocaleDateToSQL(psLocateValue);
						if (psLocateValue.length = 0) {
							fValidLocateValue = false;
							OpenHR.messageBox("Invalid date value entered.");
						}
					}
				}
			}
		}
	
		if (fValidLocateValue == true) {
			disableMenu();
			var frmUseful = document.getElementById("frmUseful");
			
			// Get the optionData.asp to get the link find records.
			getDataForm.txtTableID.value = frmUseful.txtTableID.value;
			getDataForm.txtViewID.value = selectView.options[selectView.selectedIndex].value;
			getDataForm.txtOrderID.value = selectOrder.options[selectOrder.selectedIndex].value;
			getDataForm.txtFirstRecPos.value = dataForm.txtFirstRecPos.value;
			getDataForm.txtCurrentRecCount.value = dataForm.txtRecordCount.value;
			getDataForm.txtGotoLocateValue.value = psLocateValue;
			getDataForm.txtPageAction.value = psAction;

			refreshData(); // should be in scope (tbBulkBookingSelectionData)
		}

	}

	function convertLocaleDateToSQL(psDateString)
	{ 
		/* Convert the given date string (in locale format) into 
		SQL format (mm/dd/yyyy). */
		var sDateFormat;
		var iDays;
		var iMonths;
		var iYears;
		var sDays;
		var sMonths;
		var sYears;
		var iValuePos;
		var sTempValue;
		var sValue;
		var iLoop;
		
		sDateFormat = OpenHR.LocaleDateFormat;

		sDays="";
		sMonths="";
		sYears="";
		iValuePos = 0;

		// Trim leading spaces.
		sTempValue = psDateString.substr(iValuePos,1);
		while (sTempValue.charAt(0) == " ") 
		{
			iValuePos = iValuePos + 1;		
			sTempValue = psDateString.substr(iValuePos,1);
		}

		for (iLoop=0; iLoop<sDateFormat.length; iLoop++)  {
			if ((sDateFormat.substr(iLoop,1).toUpperCase() == 'D') && (sDays.length==0)){
				sDays = psDateString.substr(iValuePos,1);
				iValuePos = iValuePos + 1;
				sTempValue = psDateString.substr(iValuePos,1);

				if (isNaN(sTempValue) == false) {
					sDays = sDays.concat(sTempValue);			
				}
				iValuePos = iValuePos + 1;		
			}

			if ((sDateFormat.substr(iLoop,1).toUpperCase() == 'M') && (sMonths.length==0)){
				sMonths = psDateString.substr(iValuePos,1);
				iValuePos = iValuePos + 1;
				sTempValue = psDateString.substr(iValuePos,1);

				if (isNaN(sTempValue) == false) {
					sMonths = sMonths.concat(sTempValue);			
				}
				iValuePos = iValuePos + 1;
			}

			if ((sDateFormat.substr(iLoop,1).toUpperCase() == 'Y') && (sYears.length==0)){
				sYears = psDateString.substr(iValuePos,1);
				iValuePos = iValuePos + 1;
				sTempValue = psDateString.substr(iValuePos,1);

				if (isNaN(sTempValue) == false) {
					sYears = sYears.concat(sTempValue);			
				}
				iValuePos = iValuePos + 1;
				sTempValue = psDateString.substr(iValuePos,1);

				if (isNaN(sTempValue) == false) {
					sYears = sYears.concat(sTempValue);			
				}
				iValuePos = iValuePos + 1;
				sTempValue = psDateString.substr(iValuePos,1);

				if (isNaN(sTempValue) == false) {
					sYears = sYears.concat(sTempValue);			
				}
				iValuePos = iValuePos + 1;
			}

			// Skip non-numerics
			sTempValue = psDateString.substr(iValuePos,1);
			while (isNaN(sTempValue) == true) {
				iValuePos = iValuePos + 1;		
				sTempValue = psDateString.substr(iValuePos,1);
			}
		}

		while (sDays.length < 2) {
			sTempValue = "0";
			sDays = sTempValue.concat(sDays);
		}

		while (sMonths.length < 2) {
			sTempValue = "0";
			sMonths = sTempValue.concat(sMonths);
		}

		while (sYears.length < 2) {
			sTempValue = "0";
			sYears = sTempValue.concat(sYears);
		}

		if (sYears.length == 2) {
			iValue = parseInt(sYears);
			if (iValue < 30) {
				sTempValue = "20";
			}
			else {
				sTempValue = "19";
			}
		
			sYears = sTempValue.concat(sYears);
		}

		while (sYears.length < 4) {
			sTempValue = "0";
			sYears = sTempValue.concat(sYears);
		}

		sTempValue = sMonths.concat("/");
		sTempValue = sTempValue.concat(sDays);
		sTempValue = sTempValue.concat("/");
		sTempValue = sTempValue.concat(sYears);
	
		sValue = OpenHR.ConvertSQLDateToLocale(sTempValue);

		iYears = parseInt(sYears);
	
		while (sMonths.substr(0, 1) == "0") {
			sMonths = sMonths.substr(1);
		}
		iMonths = parseInt(sMonths);
	
		while (sDays.substr(0, 1) == "0") {
			sDays = sDays.substr(1);
		}
		iDays = parseInt(sDays);

		var newDateObj = new Date(iYears, iMonths - 1, iDays);
		if ((newDateObj.getDate() != iDays) || 
			(newDateObj.getMonth() + 1 != iMonths) || 
			(newDateObj.getFullYear() != iYears)) {
			return "";
		}
		else {
			return sTempValue;
		}
	}
</script>

<script type="text/javascript">

	function ssOleDBGridSelRecords_rowcolchange() {
		tbrefreshControls();
	}

	function ssOleDBGridSelRecords_dblClick() {
		Selection_makeSelection();
	}

</script>

<script src="<%: Url.LatestContent("~/Scripts/ctl_SetStyles.js")%>" type="text/javascript"></script>

<div leftmargin=20 topmargin=20 bottommargin=20 rightmargin=5>

<table align=center class="outline" cellPadding=5 cellSpacing=0 width=100% height="100%">
	<tr>
		<td>
			<table align="center" class="invisible" cellspacing="0" cellpadding="0" width="100%" height="100%">
				<tr height=10>
					<td colspan="3" align="center" height="10">
						<H3 class="pageTitle" align="left">
<% 
	if ucase(session("selectionType")) = ucase("picklist") then 
		Response.Write("Select Picklist")
	else
		if ucase(session("selectionType")) = ucase("filter") then 
			Response.Write("Select Filter")
		else
			Response.Write("Select Records")
		end if
	End If
	
%>
						</H3>
					</td>
				</tr>
				<tr>
					<td width=20></td>
					<td>
<%
	
	Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
	Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

	Dim rstSelRecords As DataTable
	
	session("optionLinkViewID") = session("TB_BulkBookingDefaultViewID")
	session("optionLinkOrderID") = 0
	
	sErrorDescription = ""

	if (ucase(session("selectionType")) = ucase("picklist")) or _
		(ucase(session("selectionType")) = ucase("filter")) then 

		If UCase(Session("selectionType")) = UCase("picklist") Then
			
			rstSelRecords = objDataAccess.GetFromSP("spASRIntGetAvailablePicklists" _
				, New SqlParameter("plngTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("TB_EmpTableID"))} _
				, New SqlParameter("psUserName", SqlDbType.VarChar, 255) With {.Value = Session("username")})

		Else

			rstSelRecords = objDataAccess.GetFromSP("spASRIntGetAvailableFilters" _
				, New SqlParameter("plngTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("TB_EmpTableID"))} _
				, New SqlParameter("psUserName", SqlDbType.VarChar, 255) With {.Value = Session("username")})

		End If

		' Create a table for jqGrid to convert. 		
		Response.Write("<div id='findGridRow' style='height: 400px;'>" & vbCrLf)
		Response.Write("<table id='ssOleDBGridSelRecords'>" & vbCrLf)
		Response.Write("<thead>" & vbCrLf)
		Response.Write("<tr>" & vbCrLf)			
		For iLoop = 0 To (rstSelRecords.Columns.Count - 1)
			If rstSelRecords.Columns(iLoop).ColumnName <> "name" Then
				Response.Write("<th></th>" & vbCrLf)	' id column - don't populate, but leave as a column
			Else
				Response.Write("<th id='" & rstSelRecords.Columns(iLoop).ColumnName & "'>" & Replace(rstSelRecords.Columns(iLoop).ColumnName, "_", " ") & "</th>" & vbCrLf)
			End If
		Next				
		Response.Write("</tr>" & vbCrLf)
		Response.Write("</thead>" & vbCrLf)
									
		Dim lngRowCount = 0
		
		Response.Write("<tbody>" & vbCrLf)
		For Each objRow As DataRow In rstSelRecords.Rows
			Dim sID As String = ""
			Dim sRowString As String = ""
			
			For iLoop = 0 To (rstSelRecords.Columns.Count - 1)
				
				Dim sColumnValue As String = Replace(Replace(objRow(iLoop).ToString(), "_", " "), """", "&quot;")
				
				If rstSelRecords.Columns(iLoop).ColumnName <> "name" Then
					' ID column; store value
					sID = sColumnValue
					sRowString &= String.Format("<td><input type='radio' name='sel' value='{0}'/></td>", sColumnValue)
				Else
					sRowString &= String.Format("<td>{0}</td>", sColumnValue)
				End If
				
			Next
			
			Response.Write(String.Format("<tr id={0}>{1}<tr>", sID, sRowString) & vbCrLf)
			
			lngRowCount += 1
		Next
		Response.Write("</tbody>" & vbCrLf)
		Response.Write("</table>" & vbCrLf)
		Response.Write("</div>" & vbCrLf)
				
	Else
		' Select individual employee records.
%>
						<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
							<tr height="10">
								<td height="10">
									<TABLE WIDTH="100%" height="10" class="invisible" CELLSPACING="0" CELLPADDING="0">
										<TR>
											<TD width="40">
												View :
											</TD>
											<TD width="10">
												&nbsp;
											</TD>
											<TD width="175" >
												<SELECT id="selectView" name="selectView" class="combo" style="WIDTH: 200px">
<%
	If Len(sErrorDescription) = 0 Then
			
		Dim prmDfltOrderID As New SqlParameter("plngDfltOrderID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
		Dim rstViewRecords = objDatabase.DB.GetDataTable("sp_ASRIntGetLinkViews", CommandType.StoredProcedure _
				, New SqlParameter("plngTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("TB_EmpTableID"))} _
				, prmDfltOrderID)

		If (Len(sErrorDescription) = 0) And (Len(sFailureDescription) = 0) Then
			For Each objRow As DataRow In rstViewRecords.Rows
				Response.Write("						<option value=" & objRow(0))
				If CInt(objRow(0)) = CInt(Session("optionLinkViewID")) Then
					Response.Write(" SELECTED")
				End If

				If objRow(0) = 0 Then
					Response.Write(">" & Replace(objRow(1).ToString(), "_", " ") & "</option>" & vbCrLf)
				Else
					Response.Write(">'" & Replace(objRow(1).ToString, "_", " ") & "' view</option>" & vbCrLf)
				End If
																
			Next
	
			If Session("optionLinkOrderID") <= 0 Then
				Session("optionLinkOrderID") = prmDfltOrderID.Value
			End If

		End If

	End If
%>
												</SELECT>
											</TD>
											<TD width="10" >
												<INPUT type="button" value="Go" id="btnGoView" name="btnGoView" class="btn"
														onclick="goView()" />
											</TD>
											<TD >
												&nbsp;
											</TD>
											<TD width=40>
												Order :
											</TD>
											<TD width=10>
												&nbsp;
											</TD>
											<TD width=175 >
												<SELECT id=selectOrder name=selectOrder class="combo" style="WIDTH: 200px">
<%
	If Len(sErrorDescription) = 0 Then
		
		Dim rstOrderRecords = objDatabase.GetTableOrders(CInt(CleanNumeric(Session("TB_EmpTableID"))), 0)
		For Each objRow As DataRow In rstOrderRecords.Rows
			Response.Write("						<option value=" & objRow(1))
			If objRow(1) = CInt(Session("optionLinkOrderID")) Then
				Response.Write(" SELECTED")
			End If
			Response.Write(">" & Replace(objRow(0).ToString(), "_", " ") & "</option>" & vbCrLf)
		Next

	End If
%>
												</SELECT>
											</TD>
											<TD width=10 height=10>
												<INPUT type="button" value="Go" id=btnGoOrder name=btnGoOrder class="btn"
														onclick="goOrder()" />
											</TD>
										</TR>
									</table>
								</td>
							</tr>
							<tr height=10>
								<td height=10>&nbsp;</td>
							</tr>
							<TR>
								<td>
									<div id="FindGridRow" style="height: 400px; margin-bottom: 50px;">
										<table id="ssOleDBGridSelRecords" name="ssOleDBGridSelRecords" style="width: 100%"></table>
										<div id="ssOLEDBPager" style=""></div>
									</div>
								</TD>
							</TR>
						</TABLE>
<%	
	end if
%>
											<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value="<%=sErrorDescription%>">

					</td>
					<td width=20></td>
				</tr>
				<tr id="trButtons" height=10>
					<td height=10 colspan=3>&nbsp;</td>
				</tr>
				<tr height=10>
					<td width=20></td>
					<td height=10>
						<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD>&nbsp;</TD>
								<TD width=10>
									<INPUT id=cmdOK type=button value=OK name=cmdOK style="WIDTH: 80px" width="80" class="btn"
											onclick="Selection_makeSelection()" />
								</TD>
								<TD width=10>&nbsp;</TD>
								<TD width=10>
									<INPUT id=cmdCancel type=button value=Cancel name=cmdCancel style="WIDTH: 80px" width="80" class="btn"
											onclick="window.parent.close();" />
								</TD>
							</TR>
						</TABLE>
					</td>
					<td width=20></td>
				</tr>
			</TABLE>
		</td>
	</tr>
</table>

<INPUT type='hidden' id="txtTicker" name="txtTicker" value="0">
<INPUT type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">

<FORM id="frmUseful" name="frmUseful" style="visibility:hidden;display:none">
	<INPUT type='hidden' id="txtSelectionType" name="txtSelectionType" value="<%=session("selectionType")%>">
	<INPUT type='hidden' id="txtTableID" name="txtTableID" value="<%=session("TB_EmpTableID")%>">
	<INPUT type="hidden" id="txtMenuSaved" name="txtMenuSaved" value=0>
</FORM>

<form name="frmPrompt" method="post" action="promptedValues" id="frmPrompt" style="visibility:hidden;display:none">
	<input type="hidden" id="filterID" name="filterID">
</form>

</div>

<script type="text/javascript">
	tbBulkBookingSelection_onload();	
</script>
