<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script src="<%: Url.Content("~/Scripts/FormScripts/calendarreportdef.js")%>" type="text/javascript"></script>
<script src="<%: Url.Content("~/Scripts/ctl_SetFont.js") %>" type="text/javascript"></script>
<script src="<%: Url.Content("~/Scripts/ctl_SetStyles.js") %>" type="text/javascript"></script>
<link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css" />
<script src="<%: Url.Content("~/Scripts/jquery-1.8.2.js") %>" type="text/javascript"></script>
<script src="<%: Url.Content("~/Scripts/openhr.js") %>" type="text/javascript"></script>

<script type="text/javascript">
	function util_def_calendarreportdates_window_onload() {
		//debugger;
		var frmPopup = document.getElementById("frmPopup");
		var frmSelectionAccess = document.getElementById("frmSelectionAccess");
		frmPopup.txtLoading.value = 1;
		frmPopup.txtFirstLoad_Event.value = 1;
		frmPopup.txtFirstLoad_Lookup.value = 1;
		frmPopup.txtHaveSetLookupValues.value = 0;
		
		button_disable(frmPopup.cmdCancel, true);
		button_disable(frmPopup.cmdOK, true);

		populateEventTableCombo();

		frmPopup.cboLegendTable.selectedIndex = -1;

		var frmEvent = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmEventDetails");
		var frmDef = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmDefinition");

		if (frmEvent.eventAction.value.toUpperCase() == "NEW") {
			frmPopup.optNoEnd.checked = true;
			frmPopup.optCharacter.checked = true;
		} else {
			frmPopup.rowID.value = frmDef.grdEvents.AddItemRowIndex(frmDef.grdEvents.Bookmark);

			frmPopup.txtEventName.value = frmEvent.eventName.value;
			setEventTable(frmEvent.eventTableID.value);
			frmPopup.txtEventFilter.value = frmEvent.eventFilter.value;
			frmPopup.txtEventFilterID.value = frmEvent.eventFilterID.value;
			frmSelectionAccess.baseHidden.value = frmEvent.eventFilterHidden.value;
		}

		disabledAll();

		frmPopup.txtEventColumnsLoaded.value = 0;
		frmPopup.txtLookupColumnsLoaded.value = 0;

		populateEventColumns();

		frmPopup.txtLoading.value = 0;
	}
	
	function populateEventTableCombo() {
		var frmDef = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmDefinition");
		var frmRefresh = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmRefresh");
		var frmPopup = document.getElementById("frmPopup");
		var frmEvent = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmEventDetails");

		var sRelationString = frmEvent.relationNames.value;
		var sTableName;
		var bAdded = false;
		//var dataCollection = frmTab.elements;
		var oOption;
		var iIndex = sRelationString.indexOf("	");

		while (iIndex > 0) {
			var iRelationID = sRelationString.substr(0, iIndex);

			frmRefresh.submit();
			bAdded = true;
			oOption = document.createElement("OPTION");
			frmPopup.cboEventTable.options.add(oOption);
			oOption.value = iRelationID;

			if (iRelationID == frmDef.cboBaseTable.options[frmDef.cboBaseTable.selectedIndex].value) {
				oOption.selected = true;
			}

			sRelationString = sRelationString.substr(iIndex + 1);
			iIndex = sRelationString.indexOf("	");

			sTableName = sRelationString.substr(0, iIndex);
			oOption.innerText = sTableName;

			if (bAdded) {
				sRelationString = sRelationString.substr(iIndex + 1);
				iIndex = sRelationString.indexOf("	");

				bAdded = false;
			}
			else {
				sRelationString = sRelationString.substr(iIndex + 1);
				iIndex = sRelationString.indexOf("	");

				sRelationString = sRelationString.substr(iIndex + 1);
				iIndex = sRelationString.indexOf("	");

				bAdded = false;
			}
		}
	}
	
	function populateEventColumns() {
		// Get the columns/calcs for the current table selection.
		var frmGetDataForm = OpenHR.getForm("calendardataframe", "frmGetCalendarData");
		var frmDef = document.parentWindow.parent.window.dialogArguments.OpenHR.getForm("workframe", "frmDefinition");
		var frmPopup = document.getElementById("frmPopup");

		frmGetDataForm.txtCalendarAction.value = "LOADCALENDAREVENTDETAILSCOLUMNS";
		frmGetDataForm.txtCalendarBaseTableID.value = frmDef.cboBaseTable.options[frmDef.cboBaseTable.selectedIndex].value;
		frmGetDataForm.txtCalendarEventTableID.value = frmPopup.cboEventTable.options[frmPopup.cboEventTable.selectedIndex].value;

		//window.parent.frames("calendardataframe").refreshData();
		//data_refreshData();
		window.parent.refreshData();
	}
	
	function disabledAll() {
		var frmPopup = document.getElementById("frmPopup");
		/*Event Frame*/
		text_disable(frmPopup.txtEventName, true);
		combo_disable(frmPopup.cboEventTable, true);
		text_disable(frmPopup.txtEventFilter, true);
		button_disable(frmPopup.cmdEventFilter, true);

		/*Event Start Frame*/
		combo_disable(frmPopup.cboStartDate, true);
		combo_disable(frmPopup.cboStartSession, true);

		/*Event End Frame*/
		radio_disable(frmPopup.optNoEnd, true);
		radio_disable(frmPopup.optEndDate, true);
		combo_disable(frmPopup.cboEndDate, true);
		combo_disable(frmPopup.cboEndSession, true);
		radio_disable(frmPopup.optDuration, true);
		combo_disable(frmPopup.cboDuration, true);

		/*Key Frame*/
		text_disable(frmPopup.txtCharacter, true);
		combo_disable(frmPopup.cboEventType, true);
		combo_disable(frmPopup.cboLegendTable, true);
		combo_disable(frmPopup.cboLegendColumn, true);
		combo_disable(frmPopup.cboLegendCode, true);

		/*Event Description Frame*/
		combo_disable(frmPopup.cboEventDesc1, true);
		combo_disable(frmPopup.cboEventDesc2, true);
	}
	
	
</script>

<div id="bdyMain" name="bdyMain" <%=session("BodyColour")%> leftmargin="20" topmargin="20" bottommargin="20" rightmargin="5">
	<form id="frmPopup" name="frmPopup" onsubmit="return setForm();">
		<table align="center" width="100%" height="100%" class="outline" cellpadding="0" cellspacing="0">
			<tr>
				<td>
					<table align="center" width="100%" height="100%" class="invisible" cellpadding="4" cellspacing="0">
						<tr height="5">
							<td align="center" colspan="2" height="10">
								Select Event Information
							</td>
						</tr>
						<tr>
							<td valign="top" width="50%">
								<table class="outline" cellspacing="0" cellpadding="4" width="100%" height="100%">
									<tr height="10">
										<td height="10" colspan="5" align="left" valign="top">Event :
											<br>
											<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
												<tr height="5">
													<td width="5"></td>
													<td nowrap width="100">Name :</td>
													<td width="5"></td>
													<td>
														<input id="txtEventName" name="txtEventName" class="text textdisabled" style="WIDTH: 100%" disabled="disabled"
															onkeypress="eventChanged();"
															onkeydown="eventChanged();"
															onchange="eventChanged();">
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="5"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td nowrap width="100">Event Table :</td>
													<td width="5"></td>
													<td>
														<select id="cboEventTable" name="cboEventTable" class="combo combodisabled" disabled="disabled" style="WIDTH: 100%"
															onchange="changeEventTable();">
														</select>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="5"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td nowrap width="100">Filter :</td>
													<td width="5"></td>
													<td>
														<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
															<tr>
																<td>
																	<input id="txtEventFilter" name="txtEventFilter" class="text textdisabled" disabled="disabled" style="WIDTH: 100%">
																	<input type="hidden" id="txtEventFilterID" name="txtEventFilterID" class="text textdisabled" disabled="disabled" style="WIDTH: 100%"
																		onchange="eventChanged();">
																</td>
																<td width="25">
																	<input id="cmdEventFilter" name="cmdEventFilter" disabled="disabled" class="btn btndisabled" style="WIDTH: 100%" type="button" value="..."
																		onclick="selectRecordOption('event', 'filter')"
																		onmouseover="try{button_onMouseOver(this);}catch(e){}"
																		onmouseout="try{button_onMouseOut(this);}catch(e){}"
																		onfocus="try{button_onFocus(this);}catch(e){}"
																		onblur="try{button_onBlur(this);}catch(e){}" />
																</td>
															</tr>
														</table>
													</td>
													<td width="5"></td>
												</tr>
											</table>
										</td>
									</tr>
								</table>
							</td>
							<td valign="top" rowspan="2" width="50%">
								<table class="outline" cellspacing="0" cellpadding="4" width="100%" height="100%">
									<tr height="10">
										<td height="10" colspan="5" align="left" valign="top">Key :
											<br>
											<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
												<tr height="5">
													<td width="5"></td>
													<td colspan="1">
														<input type="radio" name="optKey" id="optCharacter" disabled="disabled"
															onclick="refreshLegendControls(); eventChanged();"
															onmouseover="try{radio_onMouseOver(this);}catch(e){}"
															onmouseout="try{radio_onMouseOut(this);}catch(e){}"
															onfocus="try{radio_onFocus(this);}catch(e){}"
															onblur="try{radio_onBlur(this);}catch(e){}" />&nbsp;
													</td>
													<td nowrap colspan="1">
														<label
															tabindex="-1"
															for="optCharacter"
															class="radio radiodisabled"
															onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
															onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
															Character</label>
													</td>
													<td width="5"></td>
													<td nowrap width="100%">
														<input id="txtCharacter" maxlength="2" name="txtCharacter" class="text textdisabled" disabled="disabled" style="WIDTH: 60px"
															onkeypress="eventChanged();"
															onkeydown="eventChanged();"
															onchange="eventChanged();">
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="6"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td colspan="1">
														<input type="radio" name="optKey" id="optLegendLookup" disabled="disabled"
															onclick="refreshLegendControls(); eventChanged();"
															onmouseover="try{radio_onMouseOver(this);}catch(e){}"
															onmouseout="try{radio_onMouseOut(this);}catch(e){}"
															onfocus="try{radio_onFocus(this);}catch(e){}"
															onblur="try{radio_onBlur(this);}catch(e){}" />
													</td>
													<td nowrap width="100" colspan="3">
														<label
															tabindex="-1"
															for="optLegendLookup"
															class="radio radiodisabled"
															onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
															onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
															Lookup Table</label>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="6"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td>&nbsp</td>
													<td width="100" nowrap>Event Type : 
													</td>
													<td width="5"></td>
													<td>
														<select id="cboEventType" name="cboEventType" disabled="disabled" width="100%" style="WIDTH: 100%" class="combo combodisabled"
															onchange="eventChanged();">
														</select>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td colspan="4">
														<hr>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td nowrap></td>
													<td width="100" nowrap>Table :
													</td>
													<td width="5"></td>
													<td>
														<select id="cboLegendTable" name="cboLegendTable" disabled="disabled" class="combo combodisabled" style="WIDTH: 100%"
															onchange="changeLegendTable();">
															<%
																Dim sErrorDescription = ""

																' Get the lookup table records.
																Dim cmdLookupTables = CreateObject("ADODB.Command")
																cmdLookupTables.CommandText = "spASRIntGetLookupTables"
																cmdLookupTables.CommandType = 4	' Stored Procedure
																cmdLookupTables.ActiveConnection = Session("databaseConnection")
	
																Err.Clear()
																Dim rstLookupTablesInfo = cmdLookupTables.Execute
	
																If (Err.Number <> 0) Then
																	sErrorDescription = "The lookup tables information could not be retrieved." & vbCrLf & FormatError(Err.Description)
																End If

																If Len(sErrorDescription) = 0 Then
																	Dim iCount = 0
																	Do While Not rstLookupTablesInfo.EOF
																		Response.Write("<OPTION value='" & rstLookupTablesInfo.fields("tableID").value & "'>" & rstLookupTablesInfo.fields("tableName").value & vbCrLf)
																		rstLookupTablesInfo.MoveNext()
																	Loop

																	' Release the ADO recordset object.
																	rstLookupTablesInfo.close()
																	rstLookupTablesInfo = Nothing
																End If
	
																' Release the ADO command object.
																cmdLookupTables = Nothing
															%>
														</select>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="6"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td nowrap></td>
													<td width="100" nowrap>Column :								
													</td>
													<td width="5"></td>
													<td>
														<select id="cboLegendColumn" name="cboLegendColumn" width="100%" style="WIDTH: 100%" class="combo"
															onchange="eventChanged();">
														</select>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="6"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td></td>
													<td width="100" nowrap>Code : 
													</td>
													<td width="5"></td>
													<td>
														<select id="cboLegendCode" name="cboLegendCode" width="100%" style="WIDTH: 100%" class="combo"
															onchange="eventChanged();">
														</select>
													</td>
													<td width="5"></td>
												</tr>
											</table>
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td valign="top" width="50%">
								<table class="outline" cellspacing="0" cellpadding="4" width="100%" height="100%">
									<tr height="5">
										<td height="5" colspan="5" align="left" valign="top">Event Start :
											<br>
											<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
												<tr height="5">
													<td width="5"></td>
													<td nowrap width="100">Start Date :</td>
													<td width="5"></td>
													<td>
														<select disabled="disabled" id="cboStartDate" name="cboStartDate" class="combo combodisabled" style="WIDTH: 100%"
															onchange="eventChanged();">
														</select>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="5"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td nowrap width="100">Start Session :</td>
													<td width="5"></td>
													<td>
														<select disabled="disabled" id="cboStartSession" name="cboStartSession" class="combo combodisabled" style="WIDTH: 100%"
															onchange="eventChanged();">
														</select>
													</td>
													<td width="5"></td>
												</tr>
											</table>
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td valign="top" width="50%" rowspan="2">
								<table class="outline" cellspacing="0" cellpadding="4" width="100%" height="100%">
									<tr height="10">
										<td height="10" colspan="5" rowspan="2" align="left" valign="top">Event End :
											<br>
											<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
												<tr height="5">
													<td width="5"></td>
													<td colspan="1">
														<input type="radio" name="optEnd" id="optNoEnd"
															onclick="refreshEventControls(); eventChanged();"
															onmouseover="try{radio_onMouseOver(this);}catch(e){}"
															onmouseout="try{radio_onMouseOut(this);}catch(e){}"
															onfocus="try{radio_onFocus(this);}catch(e){}"
															onblur="try{radio_onBlur(this);}catch(e){}" />&nbsp;
													</td>
													<td nowrap colspan="3">
														<label
															tabindex="-1"
															for="optNoEnd"
															class="radio"
															onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
															onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
															None</label>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="6"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td colspan="1">
														<input type="radio" name="optEnd" id="optEndDate"
															onclick="refreshEventControls(); eventChanged();"
															onmouseover="try{radio_onMouseOver(this);}catch(e){}"
															onmouseout="try{radio_onMouseOut(this);}catch(e){}"
															onfocus="try{radio_onFocus(this);}catch(e){}"
															onblur="try{radio_onBlur(this);}catch(e){}" />
													</td>
													<td nowrap colspan="3">
														<label tabindex="-1"
															for="optEndDate"
															class="radio"
															onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
															onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
															End</label>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="6"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td width="5"></td>
													<td nowrap width="65">Date : </td>
													<td width="5"></td>
													<td>
														<select disabled="disabled" id="cboEndDate" name="cboEndDate" style="WIDTH: 100%" class="combo combodisabled"
															onchange="eventChanged();">
														</select>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="6"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td width="5"></td>
													<td nowrap width="65">Session : </td>
													<td width="5"></td>
													<td>
														<select disabled="disabled" id="cboEndSession" name="cboEndSession" style="WIDTH: 100%" class="combo combodisabled"
															onchange="eventChanged();">
														</select>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="6"></td>
												</tr>
												<tr height="5">
													<td width="5"></td>
													<td colspan="1">
														<input type="radio" name="optEnd" id="optDuration"
															onclick="refreshEventControls();"
															onchange="eventChanged();"
															onmouseover="try{radio_onMouseOver(this);}catch(e){}"
															onmouseout="try{radio_onMouseOut(this);}catch(e){}"
															onfocus="try{radio_onFocus(this);}catch(e){}"
															onblur="try{radio_onBlur(this);}catch(e){}"></td>
													<td nowrap width="65">
														<label
															tabindex="-1"
															for="optDuration"
															class="radio"
															onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
															onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
															Duration</label>
													</td>
													<td width="5"></td>
													<td>
														<select disabled="disabled" id="cboDuration" name="cboDuration" style="WIDTH: 100%" class="combo combodisabled"
															onchange="eventChanged();">
														</select>
													</td>
													<td width="5"></td>
												</tr>
											</table>
										</td>
									</tr>
								</table>
							</td>
							<td valign="top">
								<table class="outline" cellspacing="0" cellpadding="4" width="100%">
									<tr height="10">
										<td height="10" colspan="5" align="left">Event Description :
											<br>
											<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
												<tr height="10">
													<td width="5"></td>
													<td nowrap width="100">Description 1 : </td>
													<td width="5">&nbsp</td>
													<td>
														<select disabled="disabled" id="cboEventDesc1" name="cboEventDesc1" width="100%" class="combo combodisabled" style="WIDTH: 100%"
															onchange="eventChanged();">
														</select>
													</td>
													<td width="5"></td>
												</tr>
												<tr height="5">
													<td colspan="5"></td>
												</tr>
												<tr height="10">
													<td width="5"></td>
													<td nowrap width="100">Description 2 : </td>
													<td width="5"></td>
													<td>
														<select disabled="disabled" id="cboEventDesc2" name="cboEventDesc2" width="100%" class="combo combodisabled" style="WIDTH: 100%"
															onchange="eventChanged();">
														</select>
													</td>
													<td width="5"></td>
												</tr>
											</table>
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td valign="bottom" align="right">
								<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
									<tr>
										<td>&nbsp;</td>
										<td width="10">
											<input id="cmdOK" type="button" value="OK" name="cmdOK" class="btn" style="WIDTH: 80px" width="80"
												onclick="setForm()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
										<td width="10">&nbsp;</td>
										<td width="10">
											<input id="cmdCancel" type="button" value="Cancel" name="cmdCancel" class="btn" style="WIDTH: 80px" width="80"
												onclick="cancelClick();"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr height="5">
							<td colspan="5"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>

		<input type="hidden" id="txtLookupColumnsLoaded" name="txtLookupColumnsLoaded">
		<input type="hidden" id="txtEventColumnsLoaded" name="txtEventColumnsLoaded">
		<input type="hidden" id="txtFirstLoad_Lookup" name="txtFirstLoad_Lookup">
		<input type="hidden" id="txtFirstLoad_Event" name="txtFirstLoad_Event">
		<input type="hidden" id="txtHaveSetLookupValues" name="txtHaveSetLookupValues">
		<input type="hidden" id="txtLoading" name="txtLoading">
		<input type="hidden" id="rowID" name="rowID" value="-1">
		<input type="hidden" id="txtNoDateColumns" name="txtNoDateColumns" value="0">
	</form>

	<form id="frmRecordSelection" name="frmRecordSelection" target="recordSelection" action="util_recordSelection" method="post" style="visibility: hidden; display: none">
		<input type="hidden" id="recSelType" name="recSelType">
		<input type="hidden" id="recSelTableID" name="recSelTableID">
		<input type="hidden" id="recSelCurrentID" name="recSelCurrentID">
		<input type="hidden" id="recSelTable" name="recSelTable">
		<input type="hidden" id="recSelDefOwner" name="recSelDefOwner">
	</form>

	<form id="frmSelectionAccess" name="frmSelectionAccess" style="visibility: hidden; display: none">
		<input type="hidden" id="baseHidden" name="baseHidden" value='N'>
	</form>

</div>

<script type="text/javascript">
	util_def_calendarreportdates_window_onload();
</script>
