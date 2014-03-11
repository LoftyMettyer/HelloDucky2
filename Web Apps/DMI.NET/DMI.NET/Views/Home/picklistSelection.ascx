<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>


<form id="frmpicklistSelectionUseful" name="frmpicklistSelectionUseful" style="visibility: hidden; display: none">
	<input type='hidden' id="txtSelectionType" name="txtSelectionType" value='<%=session("selectionType")%>'>
	<input type='hidden' id="txtTableID" name="txtTableID" value='<%=session("selectionTableID")%>'>
	<input type="hidden" id="txtMenuSaved" name="txtMenuSaved" value="0">
</form>

<script type="text/javascript">
	function picklistSelection_window_onload() {

		$("#picklistworkframe").attr("data-framesource", "PICKLISTSELECTION");

		var fOK = true;

		var frmUseful = document.getElementById("frmpicklistSelectionUseful");

		if ((frmUseful.txtSelectionType.value.toUpperCase() != "FILTER") &&
				(frmUseful.txtSelectionType.value.toUpperCase() != "PICKLIST")) {

			var sErrMsg = txtpicklistSelectionErrorDescription.value;
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

		cmdCancel.focus();

		if ((frmUseful.txtSelectionType.value.toUpperCase() == "FILTER") ||
				(frmUseful.txtSelectionType.value.toUpperCase() == "PICKLIST")) {
			setGridFont(ssOleDBGridSelRecords);

			// Set focus onto one of the form controls. 
			// NB. This needs to be done before making any reference to the grid
			ssOleDBGridSelRecords.focus();

			if (ssOleDBGridSelRecords.rows > 0) {
				// Select the top row.
				ssOleDBGridSelRecords.MoveFirst();
				ssOleDBGridSelRecords.SelBookmarks.Add(ssOleDBGridSelRecords.Bookmark);
			}

			picklistSelection_refreshControls();
		}
		else {
			setGridFont(ssOleDBGridSelRecords);

			txtTableID.value = frmUseful.txtTableID.value;
			txtViewID.value = selectView.options[selectView.selectedIndex].value;
			txtOrderID.value = selectOrder.options[selectOrder.selectedIndex].value;

			loadAddRecords();
		}
	}

	function picklistSelection_refreshControls() {
		var fNoneSelected;

		fNoneSelected = (ssOleDBGridSelRecords.SelBookmarks.Count == 0);

		button_disable(cmdOK, fNoneSelected);

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

		if (frmUseful.txtSelectionType.value.toUpperCase() == "FILTER") {
			try {
				var frmParentUseful = OpenHR.getForm("workframe", "frmUseful");
				frmParentUseful.txtChanged.value = 1;
			}
			catch (e) {
			}

			// Go to the prompted values form to get any required prompts. 
			frmPrompt.filterID.value = selectedRecordID();
			OpenHR.showInReportFrame(frmPrompt);

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
				sSelectedIDs = "";

				ssOleDBGridSelRecords.Redraw = false;
				for (iIndex = 0; iIndex < ssOleDBGridSelRecords.selbookmarks.Count() ; iIndex++) {
					ssOleDBGridSelRecords.bookmark = ssOleDBGridSelRecords.selbookmarks(iIndex);

					sRecordID = ssOleDBGridSelRecords.Columns("id").Value;

					if (sSelectedIDs.length > 0) {
						sSelectedIDs = sSelectedIDs + ",";
					}
					sSelectedIDs = sSelectedIDs + sRecordID;
				}
				ssOleDBGridSelRecords.Redraw = true;

				try {
					var frmParentUseful = OpenHR.getForm("workframe", "frmUseful");
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
		var iRecordID;

		iRecordID = 0;

		if (ssOleDBGridSelRecords.SelBookmarks.Count > 0) {
			iRecordID = ssOleDBGridSelRecords.Columns(0).Value;
		}

		return (iRecordID);
	}

	function locateRecord(psSearchFor) {
		var fFound;

		fFound = false;

		ssOleDBGridSelRecords.Redraw = false;

		ssOleDBGridSelRecords.MoveLast();
		ssOleDBGridSelRecords.MoveFirst();

		ssOleDBGridSelRecords.SelBookmarks.removeall();

		for (iIndex = 1; iIndex <= ssOleDBGridSelRecords.rows; iIndex++) {
			var sGridValue = new String(ssOleDBGridSelRecords.Columns(0).value);
			sGridValue = sGridValue.substr(0, psSearchFor.length).toUpperCase();
			if (sGridValue == psSearchFor.toUpperCase()) {
				ssOleDBGridSelRecords.SelBookmarks.Add(ssOleDBGridSelRecords.Bookmark);
				fFound = true;
				break;
			}

			if (iIndex < ssOleDBGridSelRecords.rows) {
				ssOleDBGridSelRecords.MoveNext();
			}
			else {
				break;
			}
		}

		if ((fFound == false) && (ssOleDBGridSelRecords.rows > 0)) {
			// Select the top row.
			ssOleDBGridSelRecords.MoveFirst();
			ssOleDBGridSelRecords.SelBookmarks.Add(ssOleDBGridSelRecords.Bookmark);
		}

		ssOleDBGridSelRecords.Redraw = true;
	}

	function goView() {

		// Get the picklistSelectionData.asp to get the find records.
		var dataForm = OpenHR.getForm("picklistdataframe", "frmPicklistGetData");
		dataForm.txtTableID.value = frmpicklistSelectionUseful.txtTableID.value;
		dataForm.txtViewID.value = selectView.options[selectView.selectedIndex].value;
		dataForm.txtOrderID.value = selectOrder.options[selectOrder.selectedIndex].value;
		dataForm.txtFirstRecPos.value = 1;
		dataForm.txtCurrentRecCount.value = 0;
		dataForm.txtPageAction.value = "LOAD";

		picklist_refreshData();
	}

	function goOrder() {
		// Get the picklistSelectionData.asp to get the find records.
		var dataForm = OpenHR.getForm("picklistdataframe", "frmPicklistGetData");
		dataForm.txtTableID.value = frmpicklistSelectionUseful.txtTableID.value;
		dataForm.txtViewID.value = selectView.options[selectView.selectedIndex].value;
		dataForm.txtOrderID.value = selectOrder.options[selectOrder.selectedIndex].value;
		dataForm.txtFirstRecPos.value = 1;
		dataForm.txtCurrentRecCount.value = 0;
		dataForm.txtPageAction.value = "LOAD";

		picklist_refreshData();
	}

	function selectedOrderID() {
		return selectOrder.options[selectOrder.selectedIndex].value;
	}

	function selectedViewID() {
		return selectView.options[selectView.selectedIndex].value;
	}

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

	function picklistSelection_addhandlers() {
		OpenHR.addActiveXHandler("ssOleDBGridSelRecords", "RowColChange", "ssOleDBGridSelRecords_RowColChange()");
		OpenHR.addActiveXHandler("ssOleDBGridSelRecords", "DblClick", "ssOleDBGridSelRecords_DblClick()");
		OpenHR.addActiveXHandler("ssOleDBGridSelRecords", "KeyPress", "ssOleDBGridSelRecords_KeyPress()");
	}

	function ssOleDBGridSelRecords_RowColChange() {
		picklistSelection_refreshControls();
	}

	function ssOleDBGridSelRecords_DblClick() {

		if (frmpicklistSelectionUseful.txtSelectionType.value != "ALL") {
			makeSelection();
		}
	}

	function ssOleDBGridSelRecords_KeyPress(iKeyAscii) {

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

			locateRecord(sFind);
		}
	}

</script>


<%
	If (UCase(Session("selectionType")) <> UCase("picklist")) And (UCase(Session("selectionType")) <> UCase("filter")) Then
%>

<table align="center" class="outline" cellpadding="5" cellspacing="0" width="100%" height="95%">
	<%
	Else
	%>
	<table align="center" class="outline" cellpadding="5" cellspacing="0" width="100%" height="100%">
		<%
		End If
		%>
		<tr>
			<td>
				<table class="invisible" style="text-align: center; border-spacing: 0; border: thick; padding: 0; width:100%; height:100%">
					<tr height="10">
						<td colspan="3" align="center" height="10">
							<h3 align="center">
								<% 
									If UCase(Session("selectionType")) = UCase("picklist") Then
								%>	
									Select Picklist
								<%
								Else
									If UCase(Session("selectionType")) = UCase("filter") Then
										%>
											Select Filter
										<%
									Else
										%>		
											Select Records
										<%
									End If
							End If%>
							</h3>
						</td>
					</tr>
					<tr>
						<td width="20"></td>
						<td>
							<%
								Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
								Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
								
								Dim rstSelRecords As DataTable
								Dim sErrorDescription As String
								Dim lngRowCount As Long
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
							%>
							<object classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" id="ssOleDBGridSelRecords" name="ssOleDBGridSelRecords" codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" style="LEFT: 0px; TOP: 0px; WIDTH: 100%; HEIGHT: 400px">
								<param name="ScrollBars" value="4">
								<param name="_Version" value="196616">
								<param name="DataMode" value="2">
								<param name="Cols" value="0">
								<param name="Rows" value="0">
								<param name="BorderStyle" value="1">
								<param name="RecordSelectors" value="0">
								<param name="GroupHeaders" value="0">
								<param name="ColumnHeaders" value="0">
								<param name="GroupHeadLines" value="0">
								<param name="HeadLines" value="0">
								<param name="FieldDelimiter" value="(None)">
								<param name="FieldSeparator" value="(Tab)">
								<param name="Col.Count" value="<%=rstSelRecords.Columns.Count%>">
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
								<param name="AllowColumnSizing" value="0">
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
								<param name="Columns.Count" value="<%=rstSelRecords.Columns.Count%>">
								<%
									For iLoop = 0 To (rstSelRecords.Columns.Count - 1)
										If rstSelRecords.Columns(iLoop).ColumnName <> "name" Then
								%>
								<param name="Columns(<%=iLoop%>).Width" value="0">
								<param name="Columns(<%=iLoop%>).Visible" value="0">
								<% 
								Else
								%>
								<param name="Columns(<%=iLoop%>).Width" value="100000">
								<param name="Columns(<%=iLoop%>).Visible" value="-1">
								<%
								End If
								%>
								<param name="Columns(<%=iLoop%>).Columns.Count" value="1">
								<param name="Columns(<%=iLoop%>).Caption" value="<%=Replace(rstSelRecords.Columns(iLoop).ColumnName, "_", "> ")%>">
								<param name="Columns(<%=iLoop%>).Name" value="<%=rstSelRecords.Columns(iLoop).ColumnName%>">
								<param name="Columns(<%=iLoop%>).Alignment" value="0">
								<param name="Columns(<%=iLoop%>).CaptionAlignment" value="3">
								<param name="Columns(<%=iLoop%>).Bound" value="0">
								<param name="Columns(<%=iLoop%>).AllowSizing" value="1">
								<param name="Columns(<%=iLoop%>).DataField" value="Column <%=iLoop%>">
								<param name="Columns(<%=iLoop%>).DataType" value="8">
								<param name="Columns(<%=iLoop%>).Level" value="0">
								<param name="Columns(<%=iLoop%>).NumberFormat" value="">
								<param name="Columns(<%=iLoop%>).Case" value="0">
								<param name="Columns(<%=iLoop%>).FieldLen" value="4096">
								<param name="Columns(<%=iLoop%>).VertScrollBar" value="0">
								<param name="Columns(<%=iLoop%>).Locked" value="0">
								<param name="Columns(<%=iLoop%>).Style" value="0">
								<param name="Columns(<%=iLoop%>).ButtonsAlways" value="0">
								<param name="Columns(<%=iLoop%>).RowCount" value="0">
								<param name="Columns(<%=iLoop%>).ColCount" value="1">
								<param name="Columns(<%=iLoop%>).HasHeadForeColor" value="0">
								<param name="Columns(<%=iLoop%>).HasHeadBackColor" value="0">
								<param name="Columns(<%=iLoop%>).HasForeColor" value="0">
								<param name="Columns(<%=iLoop%>).HasBackColor" value="0">
								<param name="Columns(<%=iLoop%>).HeadForeColor" value="0">
								<param name="Columns(<%=iLoop%>).HeadBackColor" value="0">
								<param name="Columns(<%=iLoop%>).ForeColor" value="0">
								<param name="Columns(<%=iLoop%>).BackColor" value="0">
								<param name="Columns(<%=iLoop%>).HeadStyleSet" value="">
								<param name="Columns(<%=iLoop%>).StyleSet" value="">
								<param name="Columns(<%=iLoop%>).Nullable" value="1">
								<param name="Columns(<%=iLoop%>).Mask" value="">
								<param name="Columns(<%=iLoop%>).PromptInclude" value="0">
								<param name="Columns(<%=iLoop%>).ClipMode" value="0">
								<param name="Columns(<%=iLoop%>).PromptChar" value="95">
								<%
								Next
								%>
								<param name="UseDefaults" value="-1">
								<param name="TabNavigation" value="1">
								<param name="_ExtentX" value="17330">
								<param name="_ExtentY" value="1323">
								<param name="_StockProps" value="79">
								<param name="Caption" value="">
								<param name="ForeColor" value="0">
								<param name="BackColor" value="16777215">
								<param name="Enabled" value="-1">
								<param name="DataMember" value="">
								<% 
									lngRowCount = 0
									For Each objRow As DataRow In rstSelRecords.Rows


										For iLoop = 0 To (rstSelRecords.Columns.Count - 1)
								%>
								<param name="Row(<%=lngRowCount%>).Col(<%=iLoop%>)" value="<%=Replace(objRow(iLoop).ToString(), "_", " ")%>">
								<%
								Next
								lngRowCount += 1
								Next%>
								<param name="Row.Count" value="<%=lngRowCount%>">
							</object>
							<%	

		
							Else
								' Select individual employee records.
							%>
							<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
								<tr height="10">
									<td height="10">
										<table width="100%" height="10" class="invisible" cellspacing="0" cellpadding="0">
											<tr height="10">
												<td width="40" height="10">View :
												</td>
												<td width="10" height="10">&nbsp;
												</td>
												<td width="175" height="10">
													<select id="selectView" name="selectView" class="combo" style="width: 200px">
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
												</td>
												<td width="10" height="10">
													<input type="button" value="Go" id="btnGoView" name="btnGoView" class="btn"
														onclick="goView()" />
												</td>
												<td height="10">&nbsp;
												</td>
												<td width="40" height="10">Order :
												</td>
												<td width="10" height="10">&nbsp;
												</td>
												<td width="175" height="10">
													<select id="selectOrder" name="selectOrder" class="combo" style="width: 200px">
														<%
															If Len(sErrorDescription) = 0 Then
																
																Dim rstTableOrderRecords = objDatabase.GetTableOrders(CInt(Session("selectionTableID")), 0)
																For Each objRow As DataRow In rstTableOrderRecords.Rows
																	Response.Write("						<option value=" & objRow(1))
																	If objRow(1) = CInt(Session("optionLinkOrderID")) Then
																		Response.Write(" SELECTED")
																	End If
																	Response.Write(">" & Replace(objRow(0).ToString(), "_", " ") & "</option>" & vbCrLf)
																Next
															
												End If
														%>
													</select>
												</td>
												<td width="10" height="10">
													<input type="button" value="Go" id="btnGoOrder" name="btnGoOrder" class="btn"
														onclick="goOrder()" />
												</td>
											</tr>
										</table>
									</td>
								</tr>
								<tr height="10">
									<td height="10">&nbsp;</td>
								</tr>
								<tr>
									<td style="height: 260px;vertical-align: top">
										<object classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" id="ssOleDBGridSelRecords" name="ssOleDBGridSelRecords" codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" style="LEFT: 0px; TOP: 0px; WIDTH: 100%; HEIGHT: 100%">
											<param name="ScrollBars" value="4">
											<param name="_Version" value="196617">
											<param name="DataMode" value="2">
											<param name="Cols" value="0">
											<param name="Rows" value="0">
											<param name="BorderStyle" value="1">
											<param name="RecordSelectors" value="0">
											<param name="GroupHeaders" value="0">
											<param name="ColumnHeaders" value="-1">
											<param name="GroupHeadLines" value="1">
											<param name="HeadLines" value="1">
											<param name="FieldDelimiter" value="(None)">
											<param name="FieldSeparator" value="(Tab)">
											<param name="Col.Count" value="0">
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
											<param name="SelectTypeRow" value="3">
											<param name="SelectByCell" value="-1">
											<param name="BalloonHelp" value="0">
											<param name="RowNavigation" value="1">
											<param name="CellNavigation" value="0">
											<param name="MaxSelectedRows" value="0">
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
											<param name="Columns.Count" value="0">
											<param name="UseDefaults" value="-1">
											<param name="TabNavigation" value="1">
											<param name="_ExtentX" value="17330">
											<param name="_ExtentY" value="1323">
											<param name="_StockProps" value="79">
											<param name="Caption" value="">
											<param name="ForeColor" value="0">
											<param name="BackColor" value="16777215">
											<param name="Enabled" value="-1">
											<param name="DataMember" value="">
											<param name="Row.Count" value="0">
										</object>
									</td>
								</tr>
							</table>
							<%	
							End If
							%>
							<input type='hidden' id="txtpicklistSelectionErrorDescription" name="txtpicklistSelectionErrorDescription" value="<%=sErrorDescription%>">
						</td>
						<td width="20"></td>
					</tr>
					<tr height="10">
						<td height="10" colspan="3">&nbsp;</td>
					</tr>
					<tr height="10">
						<td width="20"></td>
						<td height="10">
							<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
								<tr>
									<td>&nbsp;</td>
									<td width="10">
										<input id="cmdOK" type="button" value="OK" name="cmdOK" class="btn" style="WIDTH: 80px" width="80"
											onclick="makeSelection()" />
									</td>
									<td width="10">&nbsp;</td>
									<td width="10">
										<input id="cmdCancel" type="button" value="Cancel" class="btn" name="cmdCancel" style="WIDTH: 80px" width="80"
											onclick="closeclick();" />
									</td>
								</tr>
							</table>
						</td>
						<td width="20"></td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</table>


<input type='hidden' id="txtTicker" name="txtTicker" value="0">
<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">

<form name="frmPrompt" method="post" action="promptedValues" id="frmPrompt" style="visibility: hidden; display: none">
	<input type="hidden" id="filterID" name="filterID">
</form>

