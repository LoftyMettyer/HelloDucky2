<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<form id="frmpicklistSelectionUseful" name="frmpicklistSelectionUseful" style="visibility: hidden; display: none">
	<input type='hidden' id="txtSelectionType" name="txtSelectionType" value='<%=session("selectionType")%>'>
	<input type='hidden' id="txtTableID" name="txtTableID" value='<%=session("selectionTableID")%>'>
</form>

<script type="text/javascript">
	function picklistSelection_window_onload() {
		$("#picklistworkframe").attr("data-framesource", "PICKLISTSELECTION");

		var fOK = true;

		var frmUseful = document.getElementById("frmpicklistSelectionUseful");
		var selectView = document.getElementById("selectView");
		var selectOrder = document.getElementById("selectOrder");

		if ((frmUseful.txtSelectionType.value.toUpperCase() != "FILTER") &&
				(frmUseful.txtSelectionType.value.toUpperCase() != "PICKLIST")) {

			var sErrMsg = $('#txtpicklistSelectionErrorDescription').val();
			if (sErrMsg.length > 0) {
				fOK = false;
				OpenHR.messageBox(sErrMsg);
			}

			if (fOK == true) {
				if (selectView.length == 0) {
					fOK = false;
					OpenHR.messageBox("You do not have permission to read the table.");
				}
			}

			if (fOK == true) {
				if (selectOrder.length == 0) {
					fOK = false;
					OpenHR.messageBox("You do not have permission to use any of the table orders.");
				}
			}
		}

		if ((frmUseful.txtSelectionType.value.toUpperCase() == "FILTER") ||
				(frmUseful.txtSelectionType.value.toUpperCase() == "PICKLIST")) {

			//resize the grid to the height of its container.		
			var workPageHeight = $('.optiondatagridpage').outerHeight(true);
			var pageTitleHeight = $('.optiondatagridpage .pageTitle').outerHeight(true);
			var dropDownHeight = $('.optiondatagridpage .formField:first').outerHeight(true);
			var footerheight = $('.optiondatagridpage footer').outerHeight(true);

			var newGridHeight = workPageHeight - pageTitleHeight - dropDownHeight - footerheight;

			var selectionType = frmUseful.txtSelectionType.value;
			tableToGrid("#ssOleDBGridSelRecords", {
				height: newGridHeight,
				autowidth: true,
				onSelectRow: function (id) {
					picklistSelection_refreshControls();
					$('#cmdSelectFilter').button('enable');
				},
				ondblClickRow: function () {
					makeSelection();
				},
				colNames: [selectionType],
				colModel: [
						{ name: 'name', label: selectionType, index: 'name', sortable: false }
				]
			});

			picklistSelection_refreshControls();
		}
		else {			
			//txtTableID.value = frmUseful.txtTableID.value;
			//txtViewID.value = selectView.options[selectView.selectedIndex].value;
			//txtOrderID.value = selectOrder.options[selectOrder.selectedIndex].value;
			loadAddRecords();
		}
	}

	function picklistSelection_refreshControls() {		
		if ((frmpicklistSelectionUseful.txtSelectionType.value.toUpperCase() != "FILTER") &&
				(frmpicklistSelectionUseful.txtSelectionType.value.toUpperCase() != "PICKLIST")) {

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

	function makeSelection() {

	
		var frmUseful = document.getElementById("frmpicklistSelectionUseful");
		var frmParentUseful = OpenHR.getForm("workframe", "frmUseful") || OpenHR.getForm("ToolsFrame", "frmUseful");

		if (frmUseful.txtSelectionType.value.toUpperCase() == "FILTER") {
			try {
				frmParentUseful.txtChanged.value = 1;
			}
			catch (e) {
			}

			// Go to the prompted values form to get any required prompts. 
			var postData = {
				UtilType: utilityType.Filter,
				ID: selectedRecordID(),
				FilteredAdd: true,
			__RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
			}
			OpenHR.submitForm(null, "picklistworkframe", null, postData, "util_run_promptedValues");

		}
		else {
			if (frmUseful.txtSelectionType.value.toUpperCase() == "PICKLIST") {
				try {
					picklistdef_makeSelection(frmUseful.txtSelectionType.value, selectedRecordID(), "");
				}
				catch (e) {
				}
			}
			else {
				var sSelectedIDs = "";
				var sSelectedRows = $('#ssOleDBGridSelRecords').jqGrid('getGridParam', 'selarrrow');
				for (var iIndex = 0; iIndex < sSelectedRows.length; iIndex++) {
					var sRecordID = $("#ssOleDBGridSelRecords").jqGrid('getCell', sSelectedRows[iIndex], 'ID');
					if (sSelectedIDs.length > 0) {
						sSelectedIDs = sSelectedIDs + ",";
					}
					sSelectedIDs = sSelectedIDs + sRecordID;
				}

				if (sSelectedIDs == "")
					return;

				try {
					frmParentUseful = OpenHR.getForm("workframe", "frmUseful") || OpenHR.getForm("ToolsFrame", "frmUseful");
					frmParentUseful.txtChanged.value = 1;

					picklistdef_makeSelection(frmUseful.txtSelectionType.value, 0, sSelectedIDs);
				}
				catch (e) {
				}
			}
		}
	}

	/* Return the ID of the record selected in the find form. */
	function selectedRecordID() {
		var iRecordID = $("#ssOleDBGridSelRecords").jqGrid('getGridParam', 'selrow');

		return (iRecordID);
	}

	function goView() {
		// Get the picklistSelectionData.asp to get the find records.
		var dataForm = OpenHR.getForm("picklistdataframe", "frmPicklistGetData");
		dataForm.txtTableID.value = $('#frmpicklistSelectionUseful #txtTableID').val();
		dataForm.txtViewID.value = $('#selectView').val();
		dataForm.txtOrderID.value = $('#selectOrder').val();
		dataForm.txtFirstRecPos.value = 1;
		dataForm.txtCurrentRecCount.value = 0;
		dataForm.txtPageAction.value = "LOAD";

		picklist_refreshData();
	}

	function goOrder() {
		// Get the picklistSelectionData.asp to get the find records.
		var dataForm = OpenHR.getForm("picklistdataframe", "frmPicklistGetData");
		dataForm.txtTableID.value = $('#frmpicklistSelectionUseful #txtTableID').val();
		dataForm.txtViewID.value = $('#selectView').val();
		dataForm.txtOrderID.value = $('#selectOrder').val();
		dataForm.txtFirstRecPos.value = 1;
		dataForm.txtCurrentRecCount.value = 0;
		dataForm.txtPageAction.value = "LOAD";

		picklist_refreshData();
	}

	//function selectedOrderID() {
	//	return selectOrder.options[selectOrder.selectedIndex].value;
	//}

	//function selectedViewID() {
	//	return selectView.options[selectView.selectedIndex].value;
	//}

	function convertLocaleDateToSQL(psDateString) {
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
		var iValue;

		sDateFormat = OpenHR.LocaleDateFormat.LocaleDateFormat;

		sDays = "";
		sMonths = "";
		sYears = "";
		iValuePos = 0;

		// Trim leading spaces.
		sTempValue = psDateString.substr(iValuePos, 1);
		while (sTempValue.charAt(0) == " ") {
			iValuePos = iValuePos + 1;
			sTempValue = psDateString.substr(iValuePos, 1);
		}

		for (iLoop = 0; iLoop < sDateFormat.length; iLoop++) {
			if ((sDateFormat.substr(iLoop, 1).toUpperCase() == 'D') && (sDays.length == 0)) {
				sDays = psDateString.substr(iValuePos, 1);
				iValuePos = iValuePos + 1;
				sTempValue = psDateString.substr(iValuePos, 1);

				if (isNaN(sTempValue) == false) {
					sDays = sDays.concat(sTempValue);
				}
				iValuePos = iValuePos + 1;
			}

			if ((sDateFormat.substr(iLoop, 1).toUpperCase() == 'M') && (sMonths.length == 0)) {
				sMonths = psDateString.substr(iValuePos, 1);
				iValuePos = iValuePos + 1;
				sTempValue = psDateString.substr(iValuePos, 1);

				if (isNaN(sTempValue) == false) {
					sMonths = sMonths.concat(sTempValue);
				}
				iValuePos = iValuePos + 1;
			}

			if ((sDateFormat.substr(iLoop, 1).toUpperCase() == 'Y') && (sYears.length == 0)) {
				sYears = psDateString.substr(iValuePos, 1);
				iValuePos = iValuePos + 1;
				sTempValue = psDateString.substr(iValuePos, 1);

				if (isNaN(sTempValue) == false) {
					sYears = sYears.concat(sTempValue);
				}
				iValuePos = iValuePos + 1;
				sTempValue = psDateString.substr(iValuePos, 1);

				if (isNaN(sTempValue) == false) {
					sYears = sYears.concat(sTempValue);
				}
				iValuePos = iValuePos + 1;
				sTempValue = psDateString.substr(iValuePos, 1);

				if (isNaN(sTempValue) == false) {
					sYears = sYears.concat(sTempValue);
				}
				iValuePos = iValuePos + 1;
			}

			// Skip non-numerics
			sTempValue = psDateString.substr(iValuePos, 1);
			while (isNaN(sTempValue) == true) {
				iValuePos = iValuePos + 1;
				sTempValue = psDateString.substr(iValuePos, 1);
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

	function cancelSelection()
	{	
			if ($("#ssOleDBGrid").getGridParam('reccount') == 0) { 
					//Hide Remove and RemoveAll button
					button_disable(frmDefinition.cmdRemove, true); 
					button_disable(frmDefinition.cmdRemoveAll, true); 
					button_disable(frmDefinition.cmdFilteredAdd, false); 
				}			
			closeclick();
	}

</script>

<div class="absolutefull optiondatagridpage">
	<div>
		<span class="pageTitle">
			<%
				Select Case UCase(Session("selectionType"))
					Case UCase("picklist")
						Response.Write("Select Picklist")
					Case UCase("filter")
						Response.Write("Select Filter")
					Case Else
						Response.Write("Select Records")
				End Select
			%></span>


		<%
			Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
			Dim rstSelRecords As DataTable
			Dim sErrorDescription As String
			Dim sFailureDescription As String
																		
			Session("optionLinkViewID") = 0
			Session("optionLinkOrderID") = 0
	
			sErrorDescription = ""

			If (UCase(Session("selectionType")) = UCase("picklist")) Or _
					(UCase(Session("selectionType")) = UCase("filter")) Then

				If UCase(Session("selectionType")) = UCase("picklist") Then
					rstSelRecords = objDataAccess.GetFromSP("spASRIntGetAvailablePicklists" _
						, New SqlParameter("plngTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("selectionTableID"))} _
						, New SqlParameter("psUserName", SqlDbType.VarChar, 255) With {.Value = Session("username")})

				Else

					rstSelRecords = objDataAccess.GetFromSP("spASRIntGetAvailableFilters" _
						, New SqlParameter("plngTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("selectionTableID"))} _
						, New SqlParameter("psUserName", SqlDbType.VarChar, 255) With {.Value = Session("username")})

				End If

				' Instantiate and initialise the grid. 
				
				' Create a table for jqGrid to convert. 	
				Response.Write("<main>" & vbCrLf)

				Response.Write("<div id='findGridRow'>" & vbCrLf)
				Response.Write("<div class='clearboth'>" & vbCrLf)
				Response.Write("<div id='ssOleDBGridSelRecordsDiv'>" & vbCrLf)
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
				
						Dim sColumnValue As String = HttpUtility.HtmlEncode((objRow(iLoop).ToString()))
				
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
				Response.Write("</div>" & vbCrLf)
				Response.Write("</div>" & vbCrLf)
				Response.Write("</main>" & vbCrLf)
				
		Else
			' Select individual employee records.
		%>
		
		
		<div class="nowrap">
			<div class="tablerow">
				<label>View :</label>
				<select id="selectView" name="selectView" class="combo" onchange="goView()" style="width:90%;">
					<%
						If Len(sErrorDescription) = 0 Then
																
							Dim prmDfltOrderID As New SqlParameter("@plngDfltOrderID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
							Dim rstViewRecords = objDatabase.DB.GetDataTable("sp_ASRIntGetLinkViews", CommandType.StoredProcedure _
									, New SqlParameter("plngTableID", SqlDbType.Int) With {.Value = CInt(CleanNumeric(Session("selectionTableID")))} _
									, prmDfltOrderID)

							If (rstViewRecords.Rows.Count = 0) Then
								sFailureDescription = "You do not have permission to read the Employee table."
							End If
																
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
					%>
				</select>

				<label>Order :</label>
				<select id="selectOrder" name="selectOrder" class="combo" onchange="goOrder()" style="margin-left: 4px; width: 90%;" name="txtOwner">
					<%
						If Len(sErrorDescription) = 0 Then
																
							Dim rstTableOrderRecords = objDatabase.GetTableOrders(CInt(Session("selectionTableID")), 0)
							For Each objRow As DataRow In rstTableOrderRecords.Rows
								Response.Write("<option value=" & objRow(1))
								If objRow(1) = CInt(Session("optionLinkOrderID")) Then
									Response.Write(" selected")
								End If
								Response.Write(">" & Replace(objRow(0).ToString(), "_", " ") & "</option>" & vbCrLf)
							Next
															
						End If
					%>
				</select>
			</div>
			<br/>
		</div>
	</div>
	
	<main>
	<div class="clearboth">
		<div id='ssOleDBGridSelRecordsDiv'>
			<table id='ssOleDBGridSelRecords'></table>
		</div>
	</div>
	</main>
	<%	
	End If
	%>
	<input type='hidden' id="txtpicklistSelectionErrorDescription" name="txtpicklistSelectionErrorDescription" value="<%=sErrorDescription%>">

	<footer>
		<button id="cmdSelectFilter" name="cmdSelectFilter" disabled="disabled" onclick="makeSelection()" style="width: 80px;">OK</button>
		<button id="cmdCancelFilter" name="cmdCancelFilter" onclick="cancelSelection()" style="width: 80px;">Cancel</button>
	</footer>

</div>

<input type='hidden' id="txtTicker" name="txtTicker" value="0">
<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">

