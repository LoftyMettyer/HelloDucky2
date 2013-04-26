<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<%Dim sErrorDescription = ""
	Dim sFailureDescription = ""%>

<script type="text/javascript">
	
	function tbAddFromWaitingListFind_onload() {		
		var fOK;
		fOK = true;
		
		var frmtbFindForm = document.getElementById("frmtbFindForm");

		var sErrMsg = frmtbFindForm.txtErrorDescription.value;
		if (sErrMsg.length > 0) {
			fOK = false;
			OpenHR.messageBox(sErrMsg);
			window.parent.location.replace("login.asp");
		}

		if (fOK == true) {
			sErrMsg = frmtbFindForm.txtFailureDescription.value;
			if (sErrMsg.length > 0) {
				fOK = false;
				OpenHR.messageBox(sErrMsg);
				Cancel();
			}
		}

		if (fOK == true) {
			if (frmtbFindForm.selectView.length == 0) {
				fOK = false;
				OpenHR.messageBox("You do not have permission to read the course table.");
				Cancel();
			}
		}

		if (fOK == true) {
			if (frmtbFindForm.selectOrder.length == 0) {
				fOK = false;
				OpenHR.messageBox("You do not have permission to use any of the course table orders.");
				Cancel();
			}
		}

		if (fOK == true) {
			setGridFont(frmtbFindForm.ssOleDBGridRecords);

			// Expand the option frame and hide the work frame.
			//window.parent.document.all.item("workframeset").cols = "0, *";	
			$("#optionframe").attr("data-framesource", "TBADDFROMWAITINGLISTFIND");
			$("#workframe").hide();
			$("#optionframe").show();

			// Set focus onto one of the form controls. 
			// NB. This needs to be done before making any reference to the grid
			frmtbFindForm.cmdCancel.focus();

			//TODO: window.parent.frames("workframe").document.forms("frmtbFindForm").ssOleDBGridFindRecords.style.visibility = "hidden";

			// Get the optionData.asp to get the link find records.
			var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
			optionDataForm.txtOptionAction.value = "LOADADDFROMWAITINGLIST";
			optionDataForm.txtOptionTableID.value = frmtbFindForm.txtOptionLinkTableID.value;
			optionDataForm.txtOptionViewID.value = frmtbFindForm.selectView.options[frmtbFindForm.selectView.selectedIndex].value;
			optionDataForm.txtOptionOrderID.value = frmtbFindForm.selectOrder.options[frmtbFindForm.selectOrder.selectedIndex].value;
			optionDataForm.txtOptionRecordID.value = frmtbFindForm.txtOptionRecordID.value;
			optionDataForm.txtOptionFirstRecPos.value = 1;
			optionDataForm.txtOptionCurrentRecCount.value = 0;
			optionDataForm.txtOptionPageAction.value = "LOAD";

			frmtbFindForm.txtOptionLinkViewID.value = optionDataForm.txtOptionViewID.value;
			frmtbFindForm.txtOptionLinkOrderID.value = optionDataForm.txtOptionOrderID.value;

			refreshOptionData();	//should be in scope.
		}
	}
</script>

<script type="text/javascript">

	function Select() {		
		var frmGotoOption = document.getElementById("frmGotoOption");
		var frmtbFindForm = document.getElementById("frmtbFindForm");

		if ($("#txtStatusPExists").val() != "True") {
			//TODO: window.parent.frames("workframe").document.forms("frmtbFindForm").ssOleDBGridFindRecords.style.visibility = "visible";
		}

		frmGotoOption.txtGotoOptionAction.value = "SELECTADDFROMWAITINGLIST_1";
		frmGotoOption.txtGotoOptionRecordID.value = frmtbFindForm.txtOptionRecordID.value;
		frmGotoOption.txtGotoOptionLinkRecordID.value = ssselectedRecordID();
		frmGotoOption.txtGotoOptionPage.value = "emptyoption";
		OpenHR.submitForm(frmGotoOption);
	}

	function Cancel() {
		var frmGotoOption = document.getElementById("frmGotoOption");

		//TODO: window.parent.frames("workframe").document.forms("frmtbFindForm").ssOleDBGridFindRecords.style.visibility = "visible";

		frmGotoOption.txtGotoOptionAction.value = "CANCEL";
		frmGotoOption.txtGotoOptionLinkRecordID.value = 0;
		frmGotoOption.txtGotoOptionPage.value = "emptyoption";
		OpenHR.submitForm(frmGotoOption);
	}

	/* Return the ID of the record selected in the find form. */
	function ssselectedRecordID() {
		var iRecordID;
		var iIndex;
		var iIDColumnIndex;
		var sColumnName;

		iRecordID = 0;
		iIDColumnIndex = 0;
		var frmtbFindForm = document.getElementById("frmtbFindForm");
		if (frmtbFindForm.ssOleDBGridRecords.SelBookmarks.Count > 0) {
			for (iIndex = 0; iIndex < frmtbFindForm.ssOleDBGridRecords.Cols; iIndex++) {
				sColumnName = frmtbFindForm.ssOleDBGridRecords.Columns(iIndex).Name;
				if (sColumnName.toUpperCase() == "ID") {
					iIDColumnIndex = iIndex;
					break;
				}
			}

			iRecordID = frmtbFindForm.ssOleDBGridRecords.Columns(iIDColumnIndex).Value;
		}

		return (iRecordID);
	}

	function tbrefreshControls() {
		var frmtbFindForm = document.getElementById("frmtbFindForm");
		if (frmtbFindForm.ssOleDBGridRecords.rows > 0) {
			if (frmtbFindForm.ssOleDBGridRecords.SelBookmarks.Count > 0) {
				button_disable(frmtbFindForm.cmdSelect, false);
			}
			else {
				button_disable(frmtbFindForm.cmdSelect, true);
			}
		}
		else {
			button_disable(frmtbFindForm.cmdSelect, true);
		}

		if (frmtbFindForm.selectOrder.length <= 1) {
			combo_disable(frmtbFindForm.selectOrder, true);
			button_disable(frmtbFindForm.btnGoOrder, true);
		}

		if (frmtbFindForm.selectView.length <= 1) {
			combo_disable(frmtbFindForm.selectView, true);
			button_disable(frmtbFindForm.btnGoView, true);
		}
	}

	function goView() {
		// Get the optionData.asp to get the link find records.
		var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
		var frmtbFindForm = document.getElementById("frmtbFindForm");
		optionDataForm.txtOptionAction.value = "LOADADDFROMWAITINGLIST";
		optionDataForm.txtOptionTableID.value = frmtbFindForm.txtOptionLinkTableID.value;
		optionDataForm.txtOptionViewID.value = frmtbFindForm.selectView.options[frmtbFindForm.selectView.selectedIndex].value;
		optionDataForm.txtOptionOrderID.value = frmtbFindForm.selectOrder.options[frmtbFindForm.selectOrder.selectedIndex].value;
		optionDataForm.txtOptionRecordID.value = frmtbFindForm.txtOptionRecordID.value;
		optionDataForm.txtOptionFirstRecPos.value = 1;
		optionDataForm.txtOptionCurrentRecCount.value = 0;

		frmtbFindForm.txtOptionLinkViewID.value = optionDataForm.txtOptionViewID.value;
		frmtbFindForm.txtOptionLinkOrderID.value = optionDataForm.txtOptionOrderID.value;

		refreshOptionData();	//should be in scope...
	}

	function goOrder() {
		// Get the optionData.asp to get the link find records.
		var frmtbFindForm = document.getElementById("frmtbFindForm");
		var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
		optionDataForm.txtOptionAction.value = "LOADADDFROMWAITINGLIST";
		optionDataForm.txtOptionTableID.value = frmtbFindForm.txtOptionLinkTableID.value;
		optionDataForm.txtOptionViewID.value = frmtbFindForm.selectView.options[frmtbFindForm.selectView.selectedIndex].value;
		optionDataForm.txtOptionOrderID.value = frmtbFindForm.selectOrder.options[frmtbFindForm.selectOrder.selectedIndex].value;
		optionDataForm.txtOptionRecordID.value = frmtbFindForm.txtOptionRecordID.value;
		optionDataForm.txtOptionFirstRecPos.value = 1;
		optionDataForm.txtOptionCurrentRecCount.value = 0;

		frmtbFindForm.txtOptionLinkViewID.value = optionDataForm.txtOptionViewID.value;
		frmtbFindForm.txtOptionLinkOrderID.value = optionDataForm.txtOptionOrderID.value;

		refreshOptionData();	//should be in scope.
	}

	function selectedOrderID() {
		var frmtbFindForm = document.getElementById("frmtbFindForm");
		return frmtbFindForm.selectOrder.options[frmtbFindForm.selectOrder.selectedIndex].value;
	}

	function selectedViewID() {
		var frmtbFindForm = document.getElementById("frmtbFindForm");
		return frmtbFindForm.selectView.options[frmtbFindForm.selectView.selectedIndex].value;
	}

	function locateRecord(psFileName) {
		var fFound;

		fFound = false;
		var frmtbFindForm = document.getElementById("frmtbFindForm");
		frmtbFindForm.ssOleDBGridRecords.redraw = false;

		frmtbFindForm.ssOleDBGridRecords.MoveLast();
		frmtbFindForm.ssOleDBGridRecords.MoveFirst();

		for (var iIndex = 1; iIndex <= frmtbFindForm.ssOleDBGridRecords.rows; iIndex++) {
			var sGridValue = new String(frmtbFindForm.ssOleDBGridRecords.Columns(0).value);
			sGridValue = sGridValue.substr(0, psFileName.length).toUpperCase();
			if (sGridValue == psFileName.toUpperCase()) {
				frmtbFindForm.ssOleDBGridRecords.SelBookmarks.Add(frmtbFindForm.ssOleDBGridRecords.Bookmark);
				fFound = true;
				break;
			}

			if (iIndex < frmtbFindForm.ssOleDBGridRecords.rows) {
				frmtbFindForm.ssOleDBGridRecords.MoveNext();
			}
			else {
				break;
			}
		}

		if ((fFound == false) && (frmtbFindForm.ssOleDBGridRecords.rows > 0)) {
			// Select the top row.
			frmtbFindForm.ssOleDBGridRecords.MoveFirst();
			frmtbFindForm.ssOleDBGridRecords.SelBookmarks.Add(frmtbFindForm.ssOleDBGridRecords.Bookmark);
		}

		frmtbFindForm.ssOleDBGridRecords.redraw = true;
	}


	function util_def_addhandlers() {
		OpenHR.addActiveXHandler("ssOleDBGridRecords", "dblClick", ssOleDBGridRecords_dblClick);
		OpenHR.addActiveXHandler("ssOleDBGridRecords", "click", ssOleDBGridRecords_click);
		OpenHR.addActiveXHandler("ssOleDBGridRecords", "KeyPress", ssOleDBGridRecords_KeyPress);
	}

	function ssOleDBGridRecords_dblClick() {
		Select();
	}

	function ssOleDBGridRecords_click() {
		tbrefreshControls();
	}

	function ssOleDBGridRecords_KeyPress(iKeyAscii) {
		var iLastTick;
		var sFind;

		if ((iKeyAscii >= 32) && (iKeyAscii <= 255)) {
			var dtTicker = new Date();
			var iThisTick = new Number(dtTicker.getTime());
			if ($("#txtLastKeyFind").val().length > 0) {
				iLastTick = new Number($("#txtTicker").val());
			} else {
				iLastTick = new Number("0");
			}

			if (iThisTick > (iLastTick + 1500)) {
				sFind = String.fromCharCode(iKeyAscii);
			} else {
				sFind = txtLastKeyFind.value + String.fromCharCode(iKeyAscii);
			}
			//TODO: Confirm this works
			$("#txtTicker").val(iThisTick);
			$("#txtLastKeyFind").val(sFind);

			locateRecord(sFind);
		}
	}


</script>

<script src="<%: Url.Content("~/Scripts/ctl_SetStyles.js") %>" type="text/javascript"></script>


<div <%=session("BodyTag")%>>
	<form action="" method="POST" id="frmtbFindForm" name="frmtbFindForm">

		<table align="center" class="outline" cellpadding="5" cellspacing="0" width="100%" height="100%">
			<tr>
				<td>
					<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
						<tr>
							<td height="10" colspan="3"></td>
						</tr>
						<tr>
							<td align="center" height="10" colspan="3">
								<h3 align="center">Add From Waiting List</h3>
							</td>
						</tr>
						<tr height="10">
							<td width="20"></td>
							<td>
								<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
									<tr>
										<td width="40">View :
										</td>
										<td width="10">&nbsp;
										</td>
										<td width="175">
											<select id="selectView" name="selectView" style="HEIGHT: 22px; WIDTH: 200px" class="combo">
												<%
													On Error Resume Next

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
															sErrorDescription = "The Course views could not be retrieved." & vbCrLf & FormatError(Err.Description)
														End If

														If (Len(sErrorDescription) = 0) And (Len(sFailureDescription) = 0) Then
															Do While Not rstViewRecords.EOF
																Response.Write("						<OPTION value=" & rstViewRecords.Fields(0).Value)
																If rstViewRecords.Fields(0).Value = Session("optionLinkViewID") Then
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
																sFailureDescription = "You do not have permission to read the Course table."
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
										<td width="10">
											<input type="button" value="Go" id="btnGoView" class="btn" name="btnGoView"
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

																If rstOrderRecords.Fields(1).Value = Session("optionLinkOrderID") Then
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
											<input type="button" value="Go" id="btnGoOrder" name="btnGoOrder" class="btn"
												onclick="goOrder()"
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
						<tr>
							<td></td>
							<td>
								<object classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" id="ssOleDBGridRecords" name="ssOleDBGridRecords" codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" style="LEFT: 0px; TOP: 0px; WIDTH: 100%; HEIGHT: 400px">
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
									<param name="MultiLine" value="-1">
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
							<td></td>
						</tr>
						<tr>
							<td height="10" colspan="3"></td>
						</tr>
						<tr>
							<td></td>
							<td height="10">
								<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
									<tr>
										<td colspan="4"></td>
									</tr>
									<tr>
										<td></td>
										<td width="10">
											<input id="cmdSelect" name="cmdSelect" type="button" value="Select" style="WIDTH: 75px" width="75" class="btn"
												onclick="Select()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
										<td width="40"></td>
										<td width="10">
											<input id="cmdCancel" name="cmdCancel" type="button" value="Cancel" style="WIDTH: 75px" width="75" class="btn"
												onclick="Cancel()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
									</tr>
								</table>
							</td>
							<td></td>
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
			Response.Write("<INPUT type='hidden' id=txtOptionLinkTableID name=txtOptionLinkTableID value=" & Session("optionLinkTableID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtOptionLinkViewID name=txtOptionLinkViewID value=" & Session("optionLinkViewID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtOptionLinkOrderID name=txtOptionLinkOrderID value=" & Session("optionLinkOrderID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtOptionCourseTitle name=txtOptionCourseTitle value=" & Session("optionCourseTitle") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtOptionRecordID name=txtOptionRecordID value=" & Session("optionRecordID") & ">" & vbCrLf)
		%>
	</form>
	<input type="hidden" id="txtTicker" name="txtTicker" value="0">
	<input type="hidden" id="txtLastKeyFind" name="txtLastKeyFind" value="">
	<input type="hidden" id="txtStatusPExists" name="txtStatusPExists" value="<%=session("TB_TBStatusPExists")%>">

	<form action="tbAddFromWaitingListFind_Submit" method="post" id="frmGotoOption" name="frmGotoOption">
		<%Html.RenderPartial("~/Views/Shared/gotoOption.ascx")%>
	</form>

</div>

<script type="text/javascript">
	tbAddFromWaitingListFind_onload();
</script>

