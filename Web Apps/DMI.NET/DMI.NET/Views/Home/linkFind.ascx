<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script src="<%: Url.Content("~/Scripts/ctl_SetFont.js") %>" type="text/javascript"></script>

<script type="text/javascript">
	
	function linkFind_removeAll(jqGridID) {
		//remove all rows from the jqGrid.
		$('#' + jqGridID).jqGrid('clearGridData');
	}

	function linkFind_window_onload() {
		var fOK;
		var frmLinkFindForm = document.getElementById('frmLinkFindForm');
		fOK = true;

		var sErrMsg = frmLinkFindForm.txtErrorDescription.value;
		if (sErrMsg.length > 0) {
			fOK = false;
			OpenHR.messageBox(sErrMsg);
			window.parent.location.replace("login");
		}

		if (fOK == true) {
			sErrMsg = frmLinkFindForm.txtFailureDescription.value;
			if (sErrMsg.length > 0) {
				fOK = false;
				OpenHR.messageBox(sErrMsg);
				CancelLink();
			}
		}

		if (fOK == true) {
			if (frmLinkFindForm.selectView.length == 0) {
				fOK = false;
				OpenHR.messageBox("You do not have permission to read the link table.");
				CancelLink();
			}
		}

		if (fOK == true) {
			if (frmLinkFindForm.selectOrder.length == 0) {
				fOK = false;
				OpenHR.messageBox("You do not have permission to use any of the link table orders.");
				CancelLink();
			}
		}

		if (fOK == true) {
			setGridFont(frmLinkFindForm.ssOleDBGridLinkRecords);

			// Expand the option frame and hide the work frame.
			//window.parent.document.all.item("workframeset").cols = "0, *";
			$("#optionframe").attr("data-framesource", "LINKFIND");
			
	

			// Fault 3300
			frmLinkFindForm.txtOptionLinkViewID.value = frmLinkFindForm.selectView.options[frmLinkFindForm.selectView.selectedIndex].value;

			// Fault 3503
			//TODO: window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "hidden";

			// Get the optionData.asp to get the link find records.
			var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
			optionDataForm.txtOptionAction.value = "LOADFIND";
			optionDataForm.txtOptionTableID.value = frmLinkFindForm.txtOptionLinkTableID.value;
			optionDataForm.txtOptionViewID.value = frmLinkFindForm.selectView.options[frmLinkFindForm.selectView.selectedIndex].value;
			optionDataForm.txtOptionOrderID.value = frmLinkFindForm.selectOrder.options[frmLinkFindForm.selectOrder.selectedIndex].value;
			optionDataForm.txtOptionFirstRecPos.value = 1;
			optionDataForm.txtOptionCurrentRecCount.value = 0;
			optionDataForm.txtOptionPageAction.value = "LOAD";

			frmLinkFindForm.txtOptionLinkOrderID.value = optionDataForm.txtOptionOrderID.value;

			refreshOptionData();	//should be in scope now...


			//get width

			$("#optionframe").dialog({
				autoOpen: true,
				modal: true,
				width: 'auto',
				height: 'auto'
			});
			var width = document.getElementById('tbllinkFind').offsetWidth;
			


			// Set focus onto one of the form controls. 
			// NB. This needs to be done before making any reference to the grid
			frmLinkFindForm.cmdCancel.focus();

		}
	}
</script>

<script type="text/javascript">
	function SelectLink() {
		
		// Fault 3503
		//window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "visible";
		var selRowId = $("#ssOleDBGridLinkRecords").jqGrid('getGridParam', 'selrow');
		if (selRowId > 0) {
			$("#optionframe").dialog("destroy");
			var frmGotoOption = document.getElementById('frmGotoOption');
			var frmLinkFindForm = document.getElementById('frmLinkFindForm');

			var recordID = $("#ssOleDBGridLinkRecords").jqGrid('getCell', selRowId, 'ID');


			frmGotoOption.txtGotoOptionLinkRecordID.value = recordID;	//selectedRecordID();
			frmGotoOption.txtGotoOptionScreenID.value = frmLinkFindForm.txtOptionScreenID.value;
			frmGotoOption.txtGotoOptionLinkTableID.value = frmLinkFindForm.txtOptionLinkTableID.value;
			frmGotoOption.txtGotoOptionAction.value = "SELECTLINK";
			frmGotoOption.txtGotoOptionPage.value = "emptyoption";
			OpenHR.submitForm(frmGotoOption);
		}
	}

	function CancelLink() {
		// Fault 3503
		//window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "visible";
		$("#optionframe").dialog("destroy");
		var frmGotoOption = document.getElementById('frmGotoOption');

		frmGotoOption.txtGotoOptionAction.value = "CANCEL";
		frmGotoOption.txtGotoOptionPage.value = "emptyoption";
		OpenHR.submitForm(frmGotoOption);
	}

	/* Return the ID of the record selected in the find form. */
	//function selectedRecordID() {
	//	var iRecordID;
	//	var iIndex;
	//	var iIDColumnIndex;
	//	var sColumnName;
	//	var frmLinkFindForm = document.getElementById('frmLinkFindForm');

	//	iRecordID = 0;
	//	iIDColumnIndex = 0;

	//	//TODO: ActiveX!!
	//	if (frmLinkFindForm.ssOleDBGridLinkRecords.SelBookmarks.Count > 0) {
	//		for (iIndex = 0; iIndex < frmLinkFindForm.ssOleDBGridLinkRecords.Cols; iIndex++) {
	//			sColumnName = frmLinkFindForm.ssOleDBGridLinkRecords.Columns(iIndex).Name;
	//			if (sColumnName.toUpperCase() == "ID") {
	//				iIDColumnIndex = iIndex;
	//				break;
	//			}
	//		}

	//		iRecordID = frmLinkFindForm.ssOleDBGridLinkRecords.Columns(iIDColumnIndex).Value;
	//	}

	//	return (iRecordID);
	//}

	/* Sequential search the grid for the required ID. */
	function locateRecord(psSearchFor, pfIDMatch) {
		var fFound;
		var iIndex;
		var iIDColumnIndex;
		var sColumnName;
		var frmLinkFindForm = document.getElementById('frmLinkFindForm');

		fFound = false;

		frmLinkFindForm.ssOleDBGridLinkRecords.redraw = false;

		if (pfIDMatch == true) {
			// Locate the ID column in the grid.
			iIDColumnIndex = -1;
			for (iIndex = 0; iIndex < frmLinkFindForm.ssOleDBGridLinkRecords.Cols; iIndex++) {
				sColumnName = frmLinkFindForm.ssOleDBGridLinkRecords.Columns(iIndex).Name;
				if (sColumnName.toUpperCase() == "ID") {
					iIDColumnIndex = iIndex;
					break;
				}
			}

			if (iIDColumnIndex >= 0) {
				frmLinkFindForm.ssOleDBGridLinkRecords.MoveLast();
				frmLinkFindForm.ssOleDBGridLinkRecords.MoveFirst();

				for (iIndex = 1; iIndex <= frmLinkFindForm.ssOleDBGridLinkRecords.rows; iIndex++) {
					if (frmLinkFindForm.ssOleDBGridLinkRecords.Columns(iIDColumnIndex).value == psSearchFor) {
						frmLinkFindForm.ssOleDBGridLinkRecords.SelBookmarks.Add(frmLinkFindForm.ssOleDBGridLinkRecords.Bookmark);
						fFound = true;
						break;
					}

					if (iIndex < frmLinkFindForm.ssOleDBGridLinkRecords.rows) {
						frmLinkFindForm.ssOleDBGridLinkRecords.MoveNext();
					}
					else {
						break;
					}
				}
			}
		}
		else {
			for (iIndex = 1; iIndex <= frmLinkFindForm.ssOleDBGridLinkRecords.rows; iIndex++) {
				var sGridValue = new String(frmLinkFindForm.ssOleDBGridLinkRecords.Columns(0).value);
				sGridValue = sGridValue.substr(0, psSearchFor.length).toUpperCase();
				if (sGridValue == psSearchFor.toUpperCase()) {
					frmLinkFindForm.ssOleDBGridLinkRecords.SelBookmarks.Add(frmLinkFindForm.ssOleDBGridLinkRecords.Bookmark);
					fFound = true;
					break;
				}

				if (iIndex < frmLinkFindForm.ssOleDBGridLinkRecords.rows) {
					frmLinkFindForm.ssOleDBGridLinkRecords.MoveNext();
				}
				else {
					break;
				}
			}
		}

		if ((fFound == false) && (frmLinkFindForm.ssOleDBGridLinkRecords.rows > 0)) {
			// Select the top row.
			frmLinkFindForm.ssOleDBGridLinkRecords.MoveFirst();
			frmLinkFindForm.ssOleDBGridLinkRecords.SelBookmarks.Add(frmLinkFindForm.ssOleDBGridLinkRecords.Bookmark);
		}

		frmLinkFindForm.ssOleDBGridLinkRecords.redraw = true;
	}


	function linkFind_refreshControls() {
		//linkFind...
		var frmLinkFindForm = document.getElementById("frmLinkFindForm");

		var selRowId = $("#ssOleDBGridLinkRecords").jqGrid('getGridParam', 'selrow');
		if (selRowId > 0) {
			button_disable(frmLinkFindForm.cmdSelectLink, false);
		}
		else {
			button_disable(frmLinkFindForm.cmdSelectLink, true);
		}

		if (frmLinkFindForm.selectOrder.length <= 1) {
			combo_disable(frmLinkFindForm.selectOrder, true);
			button_disable(frmLinkFindForm.btnGoLinkOrder, true);
		}

		if (frmLinkFindForm.selectView.length <= 1) {
			combo_disable(frmLinkFindForm.selectView, true);
			button_disable(frmLinkFindForm.btnGoLinkView, true);
		}
	}
	

	function goLinkView() {
		//need this as this grid won't accept live changes :/		
		$("#ssOleDBGridLinkRecords").jqGrid('GridUnload');
		// Get the optionData.asp to get the link find records.
		var frmLinkFindForm = document.getElementById('frmLinkFindForm');
		var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
		optionDataForm.txtOptionAction.value = "LOADFIND";
		optionDataForm.txtOptionTableID.value = frmLinkFindForm.txtOptionLinkTableID.value;
		optionDataForm.txtOptionViewID.value = frmLinkFindForm.selectView.options[frmLinkFindForm.selectView.selectedIndex].value;
		optionDataForm.txtOptionOrderID.value = frmLinkFindForm.selectOrder.options[frmLinkFindForm.selectOrder.selectedIndex].value;
		optionDataForm.txtOptionFirstRecPos.value = 1;
		optionDataForm.txtOptionCurrentRecCount.value = 0;

		frmLinkFindForm.txtOptionLinkViewID.value = optionDataForm.txtOptionViewID.value;
		frmLinkFindForm.txtOptionLinkOrderID.value = optionDataForm.txtOptionOrderID.value;

		refreshOptionData();
	}

	function goLinkOrder() {
		// Get the optionData.asp to get the link find records.

		//need this as this grid won't accept live changes :/		
		$("#ssOleDBGridLinkRecords").jqGrid('GridUnload');

		var frmLinkFindForm = document.getElementById('frmLinkFindForm');
		var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");

		optionDataForm.txtOptionAction.value = "LOADFIND";
		optionDataForm.txtOptionTableID.value = frmLinkFindForm.txtOptionLinkTableID.value;
		optionDataForm.txtOptionViewID.value = frmLinkFindForm.selectView.options[frmLinkFindForm.selectView.selectedIndex].value;
		optionDataForm.txtOptionOrderID.value = frmLinkFindForm.selectOrder.options[frmLinkFindForm.selectOrder.selectedIndex].value;
		optionDataForm.txtOptionFirstRecPos.value = 1;
		optionDataForm.txtOptionCurrentRecCount.value = 0;

		frmLinkFindForm.txtOptionLinkViewID.value = optionDataForm.txtOptionViewID.value;
		frmLinkFindForm.txtOptionLinkOrderID.value = optionDataForm.txtOptionOrderID.value;

		refreshOptionData();
	}

	function selectedOrderID() {
		var frmLinkFindForm = document.getElementById('frmLinkFindForm');
		return frmLinkFindForm.selectOrder.options[frmLinkFindForm.selectOrder.selectedIndex].value;
	}

	function selectedViewID() {
		var frmLinkFindForm = document.getElementById('frmLinkFindForm');
		return frmLinkFindForm.selectView.options[frmLinkFindForm.selectView.selectedIndex].value;
	}

	function doViewHelp() {
		var helpText = "The 'View' defines the subset of data from the table that is displayed in the grid." +
			"The name of the view should give an indication of which data is included in the subset.";
		OpenHR.messageBox(helpText, 48, "Information");
	}

	function doOrderHelp() {
		var helpText = "The 'Order' defines which columns are displayed in the grid, and the order in which the data is listed.";
		OpenHR.messageBox(helpText, 48, "Information");
	}

</script>

<script for="ssOleDBGridLinkRecords" event="dblClick" language="JavaScript">
//SelectLink();
</script>

<script for="ssOleDBGridLinkRecords" event="KeyPress(iKeyAscii)" language="JavaScript">
//if ((iKeyAscii >= 32) && (iKeyAscii <= 255)) {	
	//	var dtTicker = new Date();
	//	var iThisTick = new Number(dtTicker.getTime());
	//	if (txtLastKeyFind.value.length > 0) {
	//		var iLastTick = new Number(txtTicker.value);
	//	}
	//	else {
	//		var iLastTick = new Number("0");
	//	}

	//	if (iThisTick > (iLastTick + 1500)) {
	//		var sFind = String.fromCharCode(iKeyAscii);
	//	}
	//	else {
	//		var sFind = txtLastKeyFind.value + String.fromCharCode(iKeyAscii);
	//	}

	//	txtTicker.value = iThisTick;
	//	txtLastKeyFind.value = sFind;

	//	locateRecord(sFind, false);
	//}
</script>

<script for="ssOleDBGridLinkRecords" event="click" language="JavaScript">
//refreshControls();
</script>
<script src="<%: Url.Content("~/Scripts/ctl_SetStyles.js") %>" type="text/javascript"></script>

<div id="divLinkFindForm" <%=session("BodyTag")%>>
	<form action="" method="POST" id="frmLinkFindForm" name="frmLinkFindForm">
		<table id="tbllinkFind" align="center" class="outline" cellpadding="5" cellspacing="0" width="100%" height="100%">
			<tr>
				<td>
					<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
						<tr>
							<td height="10" colspan="3"></td>
						</tr>
						<tr>
							<td align="center" height="10" colspan="3">
								<h3 class="pageTitle" align="left">Find Link Record</h3>
							</td>
						</tr>
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

													Dim sErrorDescription = ""
													Dim sFailureDescription = ""
	
													If (Len(sErrorDescription) = 0) And (Len(sFailureDescription) = 0) Then
														' Get the view records.
														Dim cmdViewRecords = CreateObject("ADODB.Command")
														cmdViewRecords.CommandText = "sp_ASRIntGetLinkViews"
														cmdViewRecords.CommandType = 4 ' Stored Procedure
														cmdViewRecords.ActiveConnection = Session("databaseConnection")

														Dim prmTableID = cmdViewRecords.CreateParameter("tableID", 3, 1)
														cmdViewRecords.Parameters.Append(prmTableID)
														prmTableID.value = CleanNumeric(Session("optionLinkTableID"))

														Dim prmDfltOrderID = cmdViewRecords.CreateParameter("dfltOrderID", 3, 2) ' 11=integer, 2=output
														cmdViewRecords.Parameters.Append(prmDfltOrderID)
														Err.Clear()
														Dim rstViewRecords = cmdViewRecords.Execute

														If (Err.Number <> 0) Then
															sErrorDescription = "The link view records could not be retrieved." & vbCrLf & FormatError(Err.Description)
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
																sFailureDescription = "You do not have permission to read the link table."
															End If
		
															' Release the ADO recordset object.
															rstViewRecords.close()
															rstViewRecords = Nothing
	
															' NB. IMPORTANT ADO NOTE.
															' When calling a stored procedure which returns a recordset AND has output parameters
															' you need to close the recordset and set it to nothing before using the output parameters. 
															If Session("optionLinkOrderID") <= 0 Then
																Session("optionLinkOrderID") = cmdViewRecords.Parameters("dfltOrderID").Value
															End If
														End If

														' Release the ADO command object.
														cmdViewRecords = Nothing
													End If
												%>
											</select>
										</td>
										<td width="17" id="tdTViewHelp" name="tdTViewHelp" onclick="doViewHelp()" style="white-space: nowrap; " disabled>
											<img id="imgTViewHelp" name="imgTViewHelp" alt="help"
												src="<%=Url.Content("~/Content/images/Help32.png")%>"
												title="What happens if I change the view?" style="width: 17px; height: 17px; border: 0; cursor: pointer" />
										</td>
										<td width="10">
											<input type="button" value="Go" id="btnGoLinkView" name="btnGoLinkView" class="btn"
												onclick="goLinkView()"
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

														Dim prmTableID = cmdOrderRecords.CreateParameter("tableID", 3, 1)
														cmdOrderRecords.Parameters.Append(prmTableID)
														prmTableID.value = CleanNumeric(Session("optionLinkTableID"))

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

																If rstOrderRecords.Fields(1).Value = CLng(Session("optionLinkOrderID")) Then
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
										<td width="17" id="tdTOrderHelp" name="tdTOrderHelp" onclick="doOrderHelp()" style="white-space: nowrap; " disabled>
											<img id="imgTOrderHelp" name="imgTOrderHelp" alt="help"
												src="<%=Url.Content("~/Content/images/Help32.png")%>"
												title="What happens if I change the order?" style="width: 17px; height: 17px; border: 0; cursor: pointer" />
										</td>
										<td width="10">
											<input type="button" value="Go" id="btnGoLinkOrder" name="btnGoLinkOrder" class="btn"
												onclick="goLinkOrder()"
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
						<tr>
							<td></td>
							<td>
								<div id="linkFindGridRow" style="height: 75%; margin-bottom: 50px;">
									<table id="ssOleDBGridLinkRecords" name="ssOleDBGridLinkRecords" style="LEFT: 0px; TOP: 0px; WIDTH: 100%; HEIGHT: 400px"></table>
								</div>
							</td>
							<td></td>
						</tr>
						<tr>
							<td height="10" colspan="3"></td>
						</tr>
						<tr>
							<td height="10"></td>
							<td height="10">
								<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
									<tr>
										<td colspan="4"></td>
									</tr>
									<tr>
										<td></td>
										<td width="10">
											<input id="cmdSelectLink" name="cmdSelectLink" type="button" value="Select" style="WIDTH: 75px" width="75" class="btn"
												onclick="SelectLink()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
										<td width="40"></td>
										<td width="10">
											<input id="cmdCancel" name="cmdCancel" type="button" value="Cancel" style="WIDTH: 75px" width="75" class="btn"
												onclick="CancelLink()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
									</tr>
								</table>
							</td>
							<td height="10"></td>
						</tr>
						<tr>
							<td height="10" colspan="3"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<%
			Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtFailureDescription name=txtFailureDescription value=""" & sFailureDescription & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtOptionScreenID name=txtOptionScreenID value=" & Session("optionScreenID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtOptionLinkTableID name=txtOptionLinkTableID value=" & Session("optionLinkTableID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtOptionLinkViewID name=txtOptionLinkViewID value=" & Session("optionLinkViewID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtOptionLinkOrderID name=txtOptionLinkOrderID value=" & Session("optionLinkOrderID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtOptionLinkRecordID name=txtOptionLinkRecordID value=" & Session("optionLinkRecordID") & ">" & vbCrLf)
		%>
	</form>
	<input type="hidden" id="txtTicker" name="txtTicker" value="0">
	<input type="hidden" id="txtLastKeyFind" name="txtLastKeyFind" value="">

	<form action="linkFind_Submit" method="post" id="frmGotoOption" name="frmGotoOption">
		<%Html.RenderPartial("~/Views/Shared/gotoOption.ascx")%>
	</form>

</div>

<script type="text/javascript">

	linkFind_window_onload();
</script>
