<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script src="<%: Url.Content("~/Scripts/ctl_SetFont.js") %>" type="text/javascript"></script>

<script type="text/javascript">

	function lookupFind_removeAll(jqGridID) {
		//remove all rows from the jqGrid.
		$('#' + jqGridID).jqGrid('clearGridData');
	}

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
			//$("#workframe").hide();
			//$("#optionframe").show();
			$("#optionframe").dialog({
				autoOpen: true,
				modal: true,
				width: 750,
				height: 600
			});

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

		var selRowId = $("#ssOleDBGrid").jqGrid('getGridParam', 'selrow');
		if (selRowId > 0) {
			//$("#optionframe").hide();
			//$("#workframe").show();
			$("#optionframe").dialog("destroy");

			frmGotoOption.txtGotoOptionColumnID.value = frmLookupFindForm.txtOptionColumnID.value;
			frmGotoOption.txtGotoOptionLookupColumnID.value = frmLookupFindForm.txtOptionLookupColumnID.value;
			frmGotoOption.txtGotoOptionLookupValue.value = selectedValue(selRowId);
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
		//$("#optionframe").hide();
		//$("#workframe").show();
		$("#optionframe").dialog("destroy");

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
		//$("#optionframe").hide();
		//$("#workframe").show();
		$("#optionframe").dialog("destroy");

		var frmGotoOption = document.getElementById("frmGotoOption");

		frmGotoOption.txtGotoOptionAction.value = "CANCEL";
		frmGotoOption.txtGotoOptionPage.value = "emptyoption";
		OpenHR.submitForm(frmGotoOption);
	}

	/* Return the value of the record selected in the find form. */
	function selectedValue(selRowId) {
		var frmLookupFindForm = document.getElementById("frmLookupFindForm");
		var sValue;

		if (frmLookupFindForm.txtIsLookupTable.value == "False") {
			//sValue = frmLookupFindForm.ssOleDBGrid.Columns(parseInt(frmLookupFindForm.txtLookupColumnGridPosition.value)).value;
			var cellNumber = parseInt(frmLookupFindForm.txtLookupColumnGridPosition.value);
			sValue = $("#ssOleDBGrid").jqGrid('getCell', selRowId, cellNumber);
		}
		else {
			//sValue = frmLookupFindForm.ssOleDBGrid.Columns(0).Value;
			sValue = $("#ssOleDBGrid").jqGrid('getCell', selRowId, 0);
		}

		return (sValue);
	}

	/* Sequential search the grid for the required OLE. */
	function locateRecord(psFileName, pfExactMatch) {
		//select the grid row that contains the record with the passed in ID.
		var rowNumber = $("#ssOleDBGrid input[value='" + psFileName + "']").parent().parent().attr("id");
		if (rowNumber >= 0) {
			$("#ssOleDBGrid").jqGrid('setSelection', rowNumber);
		} else {
			$("#ssOleDBGrid").jqGrid('setSelection', 1);
		}
	}

	function lookupFind_refreshControls() {
		//lookupFind...
		var frmLookupFindForm = document.getElementById("frmLookupFindForm");

		var selRowId = $("#ssOleDBGrid").jqGrid('getGridParam', 'selrow');
		if (selRowId > 0) {
			button_disable(frmLookupFindForm.cmdSelectLookup, false);
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

		//need this as this grid won't accept live changes :/		
		$("#ssOleDBGrid").jqGrid('GridUnload');

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
		//need this as this grid won't accept live changes :/		
		$("#ssOleDBGrid").jqGrid('GridUnload');

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
	}

	function ssOleDBGrid_dblClick() {
		SelectLookup();
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
		<div id="divFindForm" <%=session("BodyTag")%>>
			<div class="absolutefull">
				<div id="row1">
					<h3 class="pageTitle" align="left">Find Lookup Record</h3>
				</div>

				<%If Not fIsLookupTable Then%>
				<div id="row1a">
					<table align="center" class="outline" cellpadding="5" cellspacing="0" width="100%" height="100%">
						<tr>
							<td>
								<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
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
								</table>
							</td>
						</tr>
					</table>
				</div>
				<%End If 'if fIsLookupTable then%>
				<div id="lookupFindGridRow" style="height: <%If fIsLookupTable Then%>75%<%Else%>60%<%End If%>; margin-bottom: 50px;">
					<table class="outline" style="width: 100%;" id="ssOleDBGrid"></table>
				</div>
				<div id="row3">
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
				</div>
			</div>
		</div>



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
