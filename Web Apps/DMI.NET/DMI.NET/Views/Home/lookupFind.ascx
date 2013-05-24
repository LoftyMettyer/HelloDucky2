<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script src="<%: Url.Content("~/Scripts/ctl_SetFont.js") %>" type="text/javascript"></script>

<script type="text/javascript">

	function lookupFind_window_onload() {
		var fOK;
		fOK = true;
		var frmLookupFindForm = document.getElementById("frmLookupFindForm");
		var sErrMsg = frmLookupFindForm.txtErrorDescription.value;
		if (sErrMsg.length > 0) {
			fOK = false;
			OpenHR.messageBox(sErrMsg);
			window.parent.location.replace("login");
		}

		if (fOK == true) {
			sErrMsg = frmLookupFindForm.txtFailureDescription.value;
			if (sErrMsg.length > 0) {
				fOK = false;
				OpenHR.messageBox(sErrMsg);
				window.parent.location.replace("login");
			}
		}

		if (fOK == true) {
			setGridFont(frmLookupFindForm.ssOleDBGrid);

			// Expand the option frame and hide the work frame.
			//window.parent.document.all.item("workframeset").cols = "0, *";	
			$("#optionframe").attr("data-framesource", "LOOKUPFIND");
			$("#workframe").hide();
			$("#optionframe").show();

			// Set focus onto one of the form controls. 
			// NB. This needs to be done before making any reference to the grid
			frmLookupFindForm.cmdCancel.focus();

			// Fault 3503
			//todo: window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "hidden";

			// Get the optionData.asp to get the link find records.
			var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
			optionDataForm.txtOptionAction.value = "LOADLOOKUPFIND";
			optionDataForm.txtOptionColumnID.value = frmLookupFindForm.txtOptionColumnID.value;
			optionDataForm.txtOptionLookupColumnID.value = frmLookupFindForm.txtOptionLookupColumnID.value;
			optionDataForm.txtOptionLookupFilterValue.value = frmLookupFindForm.txtOptionLookupFilterValue.value;
			optionDataForm.txtOptionPageAction.value = "LOAD";
			optionDataForm.txtOptionFirstRecPos.value = 1;
			optionDataForm.txtOptionCurrentRecCount.value = 0;
			optionDataForm.txtOptionIsLookupTable.value = frmLookupFindForm.txtIsLookupTable.value;
			optionDataForm.txtOptionRecordID.value = OpenHR.getForm("workframe", "frmRecordEditForm").txtCurrentRecordID.value;

			optionDataForm.txtOptionParentTableID.value = OpenHR.getForm("workframe", "frmRecordEditForm").txtCurrentParentTableID.value;
			optionDataForm.txtOptionParentRecordID.value = OpenHR.getForm("workframe", "frmRecordEditForm").txtCurrentParentRecordID.value;

			if (frmLookupFindForm.txtIsLookupTable.value == "False") {
				optionDataForm.txtOptionTableID.value = frmLookupFindForm.txtOptionLinkTableID.value;
				optionDataForm.txtOptionViewID.value = frmLookupFindForm.selectView.options[frmLookupFindForm.selectView.selectedIndex].value;
				optionDataForm.txtOptionOrderID.value = frmLookupFindForm.selectOrder.options[frmLookupFindForm.selectOrder.selectedIndex].value;
			}
			else {
				optionDataForm.txtOptionTableID.value = 0;
				optionDataForm.txtOptionViewID.value = 0;
				optionDataForm.txtOptionOrderID.value = 0;
			}

			refreshOptionData();	//should be in scope now...
		}
	}
</script>

<script type="text/javascript">
	function SelectLookup() {
		var frmLookupFindForm = document.getElementById("frmLookupFindForm");
		var frmGotoOption = document.getElementById("frmGotoOption");

		if (frmLookupFindForm.ssOleDBGrid.SelBookmarks.Count > 0) {
			// Fault 3503
			//TODO: ?window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "visible";
			$("#optionframe").hide();
			$("#workframe").show();

			frmGotoOption.txtGotoOptionColumnID.value = frmLookupFindForm.txtOptionColumnID.value;
			frmGotoOption.txtGotoOptionLookupColumnID.value = frmLookupFindForm.txtOptionLookupColumnID.value;
			frmGotoOption.txtGotoOptionLookupValue.value = selectedValue();
			frmGotoOption.txtGotoOptionAction.value = "SELECTLOOKUP";
			frmGotoOption.txtGotoOptionPage.value = "emptyoption";
			OpenHR.submitForm(frmGotoOption);
		}
	}

	function ClearLookup() {
		var frmLookupFindForm = document.getElementById("frmLookupFindForm");
		var frmGotoOption = document.getElementById("frmGotoOption");
		// Fault 3503
		//TODO: ?window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "visible";
		$("#optionframe").hide();
		$("#workframe").show();

		frmGotoOption.txtGotoOptionColumnID.value = frmLookupFindForm.txtOptionColumnID.value;
		frmGotoOption.txtGotoOptionLookupColumnID.value = frmLookupFindForm.txtOptionLookupColumnID.value;
		frmGotoOption.txtGotoOptionLookupValue.value = "";
		frmGotoOption.txtGotoOptionAction.value = "SELECTLOOKUP";
		frmGotoOption.txtGotoOptionPage.value = "emptyoption";
		OpenHR.submitForm(frmGotoOption);
	}

	function CancelLookup() {
		// Fault 3503
		//TODO: ?window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "visible";
		$("#optionframe").hide();
		$("#workframe").show();

		var frmGotoOption = document.getElementById("frmGotoOption");

		frmGotoOption.txtGotoOptionAction.value = "CANCEL";
		frmGotoOption.txtGotoOptionPage.value = "emptyoption";
		OpenHR.submitForm(frmGotoOption);
	}

	/* Return the value of the record selected in the find form. */
	function selectedValue() {
		var frmLookupFindForm = document.getElementById("frmLookupFindForm");
		var sValue;

		sValue = "";
		if (frmLookupFindForm.ssOleDBGrid.SelBookmarks.Count > 0) {
			if (frmLookupFindForm.txtIsLookupTable.value == "False") {
				sValue = frmLookupFindForm.ssOleDBGrid.Columns(parseInt(frmLookupFindForm.txtLookupColumnGridPosition.value)).value;
			}
			else {
				sValue = frmLookupFindForm.ssOleDBGrid.Columns(0).Value;
			}
		}

		return (sValue);
	}

	/* Sequential search the grid for the required OLE. */
	function locateRecord(psFileName, pfExactMatch) {
		var frmLookupFindForm = document.getElementById("frmLookupFindForm");
		var fFound;
		var iPos;

		iPos = 0;
		if (frmLookupFindForm.txtIsLookupTable.value == "False") {
			iPos = parseInt(frmLookupFindForm.txtLookupColumnGridPosition.value);
		}

		fFound = false;

		frmLookupFindForm.ssOleDBGrid.redraw = false;

		frmLookupFindForm.ssOleDBGrid.MoveLast();
		frmLookupFindForm.ssOleDBGrid.MoveFirst();

		for (var iIndex = 1; iIndex <= frmLookupFindForm.ssOleDBGrid.rows; iIndex++) {
			if (pfExactMatch == true) {
				if (frmLookupFindForm.ssOleDBGrid.Columns(iPos).value == psFileName) {
					frmLookupFindForm.ssOleDBGrid.SelBookmarks.Add(frmLookupFindForm.ssOleDBGrid.Bookmark);
					fFound = true;
					break;
				}
			}
			else {
				var sGridValue = new String(frmLookupFindForm.ssOleDBGrid.Columns(iPos).value);
				sGridValue = sGridValue.substr(0, psFileName.length).toUpperCase();
				if (sGridValue == psFileName.toUpperCase()) {
					frmLookupFindForm.ssOleDBGrid.SelBookmarks.Add(frmLookupFindForm.ssOleDBGrid.Bookmark);
					fFound = true;
					break;
				}
			}

			if (iIndex < frmLookupFindForm.ssOleDBGrid.rows) {
				frmLookupFindForm.ssOleDBGrid.MoveNext();
			}
			else {
				break;
			}
		}

		if ((fFound == false) && (frmLookupFindForm.ssOleDBGrid.rows > 0)) {
			// Select the top row.
			frmLookupFindForm.ssOleDBGrid.MoveFirst();
			frmLookupFindForm.ssOleDBGrid.SelBookmarks.Add(frmLookupFindForm.ssOleDBGrid.Bookmark);
		}

		frmLookupFindForm.ssOleDBGrid.redraw = true;
	}

	function refreshControls() {
		//lookupFind...
		var frmLookupFindForm = document.getElementById("frmLookupFindForm");

		if (frmLookupFindForm.ssOleDBGrid.rows > 0) {
			if (frmLookupFindForm.ssOleDBGrid.SelBookmarks.Count > 0) {
				button_disable(frmLookupFindForm.cmdSelectLookup, false);
			}
			else {
				button_disable(frmLookupFindForm.cmdSelectLookup, true);
			}
		}
		else {
			button_disable(frmLookupFindForm.cmdSelectLookup, true);
		}

		if (frmLookupFindForm.txtOptionLookupMandatory.value == "true") {
			button_disable(frmLookupFindForm.cmdClearLookup, true);
		}
		else {
			button_disable(frmLookupFindForm.cmdClearLookup, false);
		}
	}

	function goView() {
		var frmLookupFindForm = document.getElementById("frmLookupFindForm");

		var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
		optionDataForm.txtOptionAction.value = "LOADLOOKUPFIND";
		optionDataForm.txtOptionColumnID.value = frmLookupFindForm.txtOptionColumnID.value;
		optionDataForm.txtOptionLookupColumnID.value = frmLookupFindForm.txtOptionLookupColumnID.value;
		optionDataForm.txtOptionLookupFilterValue.value = frmLookupFindForm.txtOptionLookupFilterValue.value;
		optionDataForm.txtOptionPageAction.value = "LOAD";
		optionDataForm.txtOptionFirstRecPos.value = 1;
		optionDataForm.txtOptionCurrentRecCount.value = 0;
		optionDataForm.txtOptionIsLookupTable.value = frmLookupFindForm.txtIsLookupTable.value;
		optionDataForm.txtOptionRecordID.value = OpenHR.getForm("workframe", "frmRecordEditForm").txtCurrentRecordID.value;

		optionDataForm.txtOptionParentTableID.value = OpenHR.getForm("workframe", "frmRecordEditForm").txtCurrentParentTableID.value;
		optionDataForm.txtOptionParentRecordID.value = OpenHR.getForm("workframe", "frmRecordEditForm").txtCurrentParentRecordID.value;

		if (frmLookupFindForm.txtIsLookupTable.value == "False") {
			optionDataForm.txtOptionTableID.value = frmLookupFindForm.txtOptionLinkTableID.value;
			optionDataForm.txtOptionViewID.value = frmLookupFindForm.selectView.options[frmLookupFindForm.selectView.selectedIndex].value;
			optionDataForm.txtOptionOrderID.value = frmLookupFindForm.selectOrder.options[frmLookupFindForm.selectOrder.selectedIndex].value;
		}
		else {
			optionDataForm.txtOptionTableID.value = 0;
			optionDataForm.txtOptionViewID.value = 0;
			optionDataForm.txtOptionOrderID.value = 0;
		}

		refreshOptionData();	//should be in scope...
	}

	function goOrder() {
		var frmLookupFindForm = document.getElementById("frmLookupFindForm");

		var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
		optionDataForm.txtOptionAction.value = "LOADLOOKUPFIND";
		optionDataForm.txtOptionColumnID.value = frmLookupFindForm.txtOptionColumnID.value;
		optionDataForm.txtOptionLookupColumnID.value = frmLookupFindForm.txtOptionLookupColumnID.value;
		optionDataForm.txtOptionLookupFilterValue.value = frmLookupFindForm.txtOptionLookupFilterValue.value;
		optionDataForm.txtOptionPageAction.value = "LOAD";
		optionDataForm.txtOptionFirstRecPos.value = 1;
		optionDataForm.txtOptionCurrentRecCount.value = 0;
		optionDataForm.txtOptionIsLookupTable.value = frmLookupFindForm.txtIsLookupTable.value;
		optionDataForm.txtOptionRecordID.value = OpenHR.getForm("workframe", "frmRecordEditForm").txtCurrentRecordID.value;

		optionDataForm.txtOptionParentTableID.value = OpenHR.getForm("workframe", "frmRecordEditForm").txtCurrentParentTableID.value;
		optionDataForm.txtOptionParentRecordID.value = OpenHR.getForm("workframe", "frmRecordEditForm").txtCurrentParentRecordID.value;

		if (frmLookupFindForm.txtIsLookupTable.value == "False") {
			optionDataForm.txtOptionTableID.value = frmLookupFindForm.txtOptionLinkTableID.value;
			optionDataForm.txtOptionViewID.value = frmLookupFindForm.selectView.options[frmLookupFindForm.selectView.selectedIndex].value;
			optionDataForm.txtOptionOrderID.value = frmLookupFindForm.selectOrder.options[frmLookupFindForm.selectOrder.selectedIndex].value;
		}
		else {
			optionDataForm.txtOptionTableID.value = 0;
			optionDataForm.txtOptionViewID.value = 0;
			optionDataForm.txtOptionOrderID.value = 0;
		}

		refreshOptionData();	//should be in scope...
	}

</script>

<script type="text/javascript">
	function lookupFind_addhandlers() {
		OpenHR.addActiveXHandler("ssOleDBGrid", "dblClick", ssOleDBGrid_dblClick);
		OpenHR.addActiveXHandler("ssOleDBGrid", "click", ssOleDBGrid_click);
		OpenHR.addActiveXHandler("ssOleDBGrid", "KeyPress", ssOleDBGrid_KeyPress);
	}

	function ssOleDBGrid_dblClick() {
		SelectLookup();
	}

	function ssOleDBGrid_click() {
		refreshControls();
	}

	function ssOleDBGrid_KeyPress(iKeyAscii) {
		var iLastTick;
		var sFind;

		if ((iKeyAscii >= 32) && (iKeyAscii <= 255)) {
			var dtTicker = new Date();
			var iThisTick = new Number(dtTicker.getTime());
			if ($("#txtLastKeyFind").val().length > 0) {
				iLastTick = new Number($("#txtTicker").val());
			}
			else {
				iLastTick = new Number("0");
			}

			if (iThisTick > (iLastTick + 1500)) {
				sFind = String.fromCharCode(iKeyAscii);
			}
			else {
				sFind = $("#txtLastKeyFind").val() + String.fromCharCode(iKeyAscii);
			}

			$("#txtTicker").val(iThisTick);
			$("#txtLastKeyFind").val(sFind);

			locateRecord(sFind, false);
		}

	}
</script>

<script src="<%: Url.Content("~/Scripts/ctl_SetStyles.js") %>" type="text/javascript"></script>

<div <%=session("BodyTag")%>>
	<form action="" method="POST" id="frmLookupFindForm" name="frmLookupFindForm">
		<%

			Dim fIsLookupTable = False
			Dim lngLookupTableID = 0
			Dim sErrorDescription = ""
			Dim sFailureDescription = ""
	
			Dim cmdGetTable = CreateObject("ADODB.Command")
			cmdGetTable.CommandText = "spASRIntGetColumnTableID"
			cmdGetTable.CommandType = 4	' Stored procedure
			cmdGetTable.ActiveConnection = Session("databaseConnection")

			Dim prmColumnID = cmdGetTable.CreateParameter("LookupColumnID", 3, 1)
			cmdGetTable.Parameters.Append(prmColumnID)
			prmColumnID.value = CleanNumeric(Session("optionLookupColumnID"))

			Dim prmTableID = cmdGetTable.CreateParameter("tableID", 3, 2)	' 3=integer, 2=output
			cmdGetTable.Parameters.Append(prmTableID)

			Err.Clear()
			cmdGetTable.Execute()
	
			If (Err.Number <> 0) Then
				sErrorDescription = "Error getting the lookup column table ID." & vbCrLf & FormatError(Err.Description)
			Else
				lngLookupTableID = CType(cmdGetTable.Parameters("tableID").Value, Integer)
			End If

			cmdGetTable = Nothing

			If Len(sErrorDescription) = 0 Then
				Dim cmdIsLookupTable = CreateObject("ADODB.Command")
				cmdIsLookupTable.CommandText = "spASRIntIsLookupTable"
				cmdIsLookupTable.CommandType = 4 ' Stored procedure
				cmdIsLookupTable.ActiveConnection = Session("databaseConnection")

				prmTableID = cmdIsLookupTable.CreateParameter("tableID", 3, 1)
				cmdIsLookupTable.Parameters.Append(prmTableID)
				prmTableID.value = CleanNumeric(lngLookupTableID)

				Dim prmIsLookup = cmdIsLookupTable.CreateParameter("isLookup", 11, 2)	' 11=bit, 2=output
				cmdIsLookupTable.Parameters.Append(prmIsLookup)

				Err.Clear()
				cmdIsLookupTable.Execute()
	
				If (Err.Number <> 0) Then
					sErrorDescription = "Error checking the lookup column table type." & vbCrLf & FormatError(Err.Description)
				Else
					fIsLookupTable = CBool(cmdIsLookupTable.Parameters("isLookup").Value)
				End If

				cmdIsLookupTable = Nothing
			End If
		%>

		<table align="center" class="outline" cellpadding="5" cellspacing="0" width="100%" height="100%">
			<tr>
				<td>
					<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
						<tr>
							<td height="10" colspan="3"></td>
						</tr>
						<tr>
							<td align="center" height="10" colspan="3">
								<h3 align="center">Find Lookup Record</h3>
							</td>
						</tr>

						<%If Not fIsLookupTable Then%>
						<tr>
							<td height="10">&nbsp;&nbsp;</td>
							<td height="10">
								<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
									<tr>
										<td width="40">View :
										</td>
										<td width="10">&nbsp;
										</td>
										<td width="175">
											<select id="selectView" name="selectView" class="combo" style="HEIGHT: 22px; WIDTH: 200px">
												<%
													On Error Resume Next

													If (Len(sErrorDescription) = 0) And (Len(sFailureDescription) = 0) Then
														' Get the view records.
														Dim cmdViewRecords = CreateObject("ADODB.Command")
														cmdViewRecords.CommandText = "spASRIntGetLookupViews"
														cmdViewRecords.CommandType = 4 ' Stored Procedure
														cmdViewRecords.ActiveConnection = Session("databaseConnection")

														prmTableID = cmdViewRecords.CreateParameter("tableID", 3, 1)
														cmdViewRecords.Parameters.Append(prmTableID)
														prmTableID.value = CleanNumeric(lngLookupTableID)

														Dim prmDfltOrderID = cmdViewRecords.CreateParameter("dfltOrderID", 3, 2) ' 11=integer, 2=output
														cmdViewRecords.Parameters.Append(prmDfltOrderID)

														prmColumnID = cmdViewRecords.CreateParameter("columnID", 3, 1)
														cmdViewRecords.Parameters.Append(prmColumnID)
														prmColumnID.value = CleanNumeric(Session("optionColumnID"))

														Err.Clear()
														Dim rstViewRecords = cmdViewRecords.Execute

														If (Err.Number <> 0) Then
															sErrorDescription = "The lookup view records could not be retrieved." & vbCrLf & FormatError(Err.Description)
														End If

														If (Len(sErrorDescription) = 0) And (Len(sFailureDescription) = 0) Then
															Do While Not rstViewRecords.EOF
																Response.Write("						<OPTION value=" & rstViewRecords.Fields(0).Value)
																If rstViewRecords.Fields(0).Value = CLng(Session("optionLinkViewID")) Then
																	Response.Write(" SELECTED")
																End If

																If rstViewRecords.Fields(0).Value = 0 Then
																	Response.Write(">" & Replace(rstViewRecords.Fields(1).Value, "_", " ") & "</OPTION>" & vbCrLf)
																Else
																	Response.Write(">'" & Replace(rstViewRecords.Fields(1).Value, "_", " ") & "' view</OPTION>" & vbCrLf)
																End If

																rstViewRecords.MoveNext()
															Loop
			
															If (rstViewRecords.EOF And rstViewRecords.BOF) Then
																sFailureDescription = "You do not have permission to read the lookup table."
															End If
		
															' Release the ADO recordset object.
															rstViewRecords.close()
															rstViewRecords = Nothing
	
															' NB. IMPORTANT ADO NOTE.
															' When calling a stored procedure which returns a recordset AND has output parameters
															' you need to close the recordset and set it to nothing before using the output parameters. 
															If Session("optionLookupOrderID") <= 0 Then
																Session("optionLookupOrderID") = cmdViewRecords.Parameters("dfltOrderID").Value
															End If
														End If

														' Release the ADO command object.
														cmdViewRecords = Nothing
													End If
												%>
											</select>
										</td>
										<td width="10">
											<input type="button" value="Go" class="btn" id="btnGoView" name="btnGoView"
												onclick="goView()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
										<td>&nbsp;
										</td>
										<td width="40">Order :
										</td>
										<td width="10">&nbsp;
										</td>
										<td width="175">
											<select id="selectOrder" name="selectOrder" class="combo" style="HEIGHT: 22px; WIDTH: 200px">
												<%
													If (Len(sErrorDescription) = 0) And (Len(sFailureDescription) = 0) Then
														' Get the order records.
														Dim cmdOrderRecords = CreateObject("ADODB.Command")
														cmdOrderRecords.CommandText = "sp_ASRIntGetTableOrders"
														cmdOrderRecords.CommandType = 4	' Stored Procedure
														cmdOrderRecords.ActiveConnection = Session("databaseConnection")

														prmTableID = cmdOrderRecords.CreateParameter("tableID", 3, 1)
														cmdOrderRecords.Parameters.Append(prmTableID)
														prmTableID.value = CleanNumeric(lngLookupTableID)

														Dim prmViewID = cmdOrderRecords.CreateParameter("viewID", 3, 1)
														cmdOrderRecords.Parameters.Append(prmViewID)
														prmViewID.value = 0

														Err.Clear()
														Dim rstOrderRecords = cmdOrderRecords.Execute

														If (Err.Number <> 0) Then
															sErrorDescription = "The order records could not be retrieved." & vbCrLf & FormatError(Err.Description)
														End If

														If (Len(sErrorDescription) = 0) And (Len(sFailureDescription) = 0) Then
															Do While Not rstOrderRecords.EOF
																Response.Write("						<OPTION value=" & rstOrderRecords.Fields(1).Value)
																If rstOrderRecords.Fields(1).Value = CInt(Session("optionOrderID")) Then
																	Response.Write(" SELECTED")
																End If

																Response.Write(">" & Replace(rstOrderRecords.Fields(0).Value, "_", " ") & "</OPTION>" & vbCrLf)

																rstOrderRecords.MoveNext()
															Loop

															' Release the ADO recordset object.
															rstOrderRecords.close()
															rstOrderRecords = Nothing
														End If
	
														' Release the ADO command object.
														cmdOrderRecords = Nothing
													End If
												%>
											</select>
										</td>
										<td width="10">
											<input type="button" value="Go" class="btn" id="btnGoOrder" name="btnGoOrder"
												onclick="goOrder()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
									</tr>
								</table>
							</td>
							<td height="10">&nbsp;&nbsp;</td>
						</tr>
						<tr>
							<td height="10" colspan="3"></td>
						</tr>
						<%End If 'if fIsLookupTable then%>

						<tr>
							<td width="10"></td>
							<td>
								<object classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" id="ssOleDBGrid" name="ssOleDBGrid" codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" style="LEFT: 0px; TOP: 0px; WIDTH: 100%; HEIGHT: 400px">
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
							<td width="10"></td>
						</tr>
						<tr>
							<td height="10" colspan="3"></td>
						</tr>
						<tr>
							<td width="20"></td>
							<td height="10">
								<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
									<tr>
										<td colspan="6"></td>
									</tr>
									<tr>
										<td></td>
										<td width="10">
											<input id="cmdSelectLookup" name="cmdSelectLookup" type="button" value="Select" style="WIDTH: 75px" width="75" class="btn"
												onclick="SelectLookup()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
										<td width="40"></td>
										<td width="10">
											<input id="cmdClearLookup" name="cmdClearLookup" type="button" value="Clear" style="WIDTH: 75px" width="75" class="btn"
												onclick="ClearLookup()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
										<td width="40"></td>
										<td width="10">
											<input id="cmdCancel" name="cmdCancel" type="button" value="Cancel" style="WIDTH: 75px" width="75" class="btn"
												onclick="CancelLookup()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
									</tr>
								</table>
							</td>
							<td width="20"></td>
						</tr>
						<tr>
							<td height="10" colspan="3"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>

		<input type='hidden' id="txtErrorDescription" name="txtErrorDescription" value="<%=sErrorDescription%>">
		<input type='hidden' id="txtFailureDescription" name="txtFailureDescription" value="<%=sFailureDescription%>">
		<input type='hidden' id="txtOptionColumnID" name="txtOptionColumnID" value='<%=session("optionColumnID")%>'>
		<input type='hidden' id="txtOptionLookupColumnID" name="txtOptionLookupColumnID" value='<%=session("optionLookupColumnID")%>'>
		<input type='hidden' id="txtOptionLookupMandatory" name="txtOptionLookupMandatory" value='<%=session("optionLookupMandatory")%>'>
		<input type='hidden' id="txtOptionLookupValue" name="txtOptionLookupValue" value='<%=session("optionLookupValue")%>'>
		<input type='hidden' id="txtOptionLookupFilterValue" name="txtOptionLookupFilterValue" value="<%=session("optionLookupFilterValue")%>">
		<input type='hidden' id="txtIsLookupTable" name="txtIsLookupTable" value="<%=fIsLookupTable%>">
		<input type='hidden' id="txtOptionLinkTableID" name="txtOptionLinkTableID" value="<%=lngLookupTableID%>">
		<input type='hidden' id="txtLookupColumnGridPosition" name="txtLookupColumnGridPosition" value="0">
	</form>

	<input type='hidden' id="txtTicker" name="txtTicker" value="0">
	<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">

	<form action="lookupFind_Submit" method="post" id="frmGotoOption" name="frmGotoOption">
		<%Html.RenderPartial("~/Views/Shared/gotoOption.ascx")%>
	</form>

</div>

<script type="text/javascript">
	lookupFind_addhandlers();
	lookupFind_window_onload();
</script>