<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<script src="<%: Url.LatestContent("~/bundles/utilities_calendarreports")%>" type="text/javascript"></script>  

<%Html.RenderPartial("Util_Def_CustomReports/dialog")%>

<script type="text/javascript">
function selectRecordOptionCalDef(psTable, psType) {
	var sURL;
	var iCurrentID;
	var iTableID;
	if (psTable == 'base') {
		iTableID = frmDefinition.cboBaseTable.options[frmDefinition.cboBaseTable.selectedIndex].value;

		if (psType == 'picklist') {
			iCurrentID = frmDefinition.txtBasePicklistID.value;
		}
		else {
			iCurrentID = frmDefinition.txtBaseFilterID.value;
		}
	}
	if (psTable == 'p1') {
		iTableID = frmDefinition.txtParent1ID.value;

		if (psType == 'picklist') {
			iCurrentID = frmDefinition.txtParent1PicklistID.value;
		}
		else {
			iCurrentID = frmDefinition.txtParent1FilterID.value;
		}
	}
	if (psTable == 'p2') {
		iTableID = frmDefinition.txtParent2ID.value;

		if (psType == 'picklist') {
			iCurrentID = frmDefinition.txtParent2PicklistID.value;
		}
		else {
			iCurrentID = frmDefinition.txtParent2FilterID.value;
		}
	}

	frmRecordSelection.recSelTable.value = psTable;
	frmRecordSelection.recSelType.value = psType;
	frmRecordSelection.recSelTableID.value = iTableID;
	frmRecordSelection.recSelCurrentID.value = iCurrentID;

	var strDefOwner = new String(frmDefinition.txtOwner.value);
	var strCurrentUser = new String(frmUseful.txtUserName.value);

	strDefOwner = strDefOwner.toLowerCase();
	strCurrentUser = strCurrentUser.toLowerCase();

	if (strDefOwner == strCurrentUser) {
		frmRecordSelection.recSelDefOwner.value = '1';
	}
	else {
		frmRecordSelection.recSelDefOwner.value = '0';
	}

	sURL = "util_recordSelection" +
			"?recSelType=" + escape(frmRecordSelection.recSelType.value) +
			"&recSelTableID=" + escape(frmRecordSelection.recSelTableID.value) +
			"&recSelCurrentID=" + escape(frmRecordSelection.recSelCurrentID.value) +
			"&recSelTable=" + escape(frmRecordSelection.recSelTable.value) +
			"&recSelDefOwner=" + escape(frmRecordSelection.recSelDefOwner.value);
	//openDialog(sURL, (screen.width) / 3, (screen.height) / 2, "yes", "yes");
	openDialog(sURL, (screen.width) / 3 + 40, (screen.height) / 2 - 30, "no", "no");

	frmUseful.txtChanged.value = 1;
	refreshTab1Controls();
}

function eventAdd() {
	var sURL;
	var frmEvent = OpenHR.getForm("workframe", "frmEventDetails");
	frmEvent.eventAction.value = "NEW";
	frmEvent.eventID.value = getEventKey();
	frmEvent.eventFilterHidden.value = "";

	if (frmDefinition.grdEvents.Rows < 999) {
		sURL = "util_def_calendarreportdates_main" +
			"?eventAction=" + escape(frmEvent.eventAction.value) +
			"&eventName=" + escape(frmEvent.eventName.value) +
			"&eventID=" + escape(frmEvent.eventID.value) +
			"&eventTableID=" + escape(frmEvent.eventTableID.value) +
			"&eventTable=" + escape(frmEvent.eventTable.value) +
			"&eventFilterID=" + escape(frmEvent.eventFilterID.value) +
			"&eventFilter=" + escape(frmEvent.eventFilter.value) +
			"&eventFilterHidden=" + escape(frmEvent.eventFilterHidden.value) +
			"&eventStartDateID=" + escape(frmEvent.eventStartDateID.value) +
			"&eventStartDate=" + escape(frmEvent.eventStartDate.value) +
			"&eventStartSessionID=" + escape(frmEvent.eventStartSessionID.value) +
			"&eventStartSession=" + escape(frmEvent.eventStartSession.value) +
			"&eventEndDateID=" + escape(frmEvent.eventEndDateID.value) +
			"&eventEndDate=" + escape(frmEvent.eventEndDate.value) +
			"&eventEndSessionID=" + escape(frmEvent.eventEndSessionID.value) +
			"&eventEndSession=" + escape(frmEvent.eventEndSession.value) +
			"&eventDurationID=" + escape(frmEvent.eventDurationID.value) +
			"&eventDuration=" + escape(frmEvent.eventDuration.value) +
			"&eventLookupType=" + escape(frmEvent.eventLookupType.value) +
			"&eventKeyCharacter=" + escape(frmEvent.eventKeyCharacter.value) +
			"&eventLookupTableID=" + escape(frmEvent.eventLookupTableID.value) +
			"&eventLookupColumnID=" + escape(frmEvent.eventLookupColumnID.value) +
			"&eventLookupCodeID=" + escape(frmEvent.eventLookupCodeID.value) +
			"&eventTypeColumnID=" + escape(frmEvent.eventTypeColumnID.value) +
			"&eventDesc1ID=" + escape(frmEvent.eventDesc1ID.value) +
			"&eventDesc1=" + escape(frmEvent.eventDesc1.value) +
			"&eventDesc2ID=" + escape(frmEvent.eventDesc2ID.value) +
			"&eventDesc2=" + escape(frmEvent.eventDesc2.value) +
			"&relationNames=" + escape(frmEvent.relationNames.value);
		//openDialogCalEvent(sURL, 650, 500, "no", "no");
		openDialog(sURL, (screen.width) / 3.4, (screen.height) / 1.4, "no", "no");
		
		frmUseful.txtChanged.value = 1;
	} else {
		var sMessage = "";
		sMessage = "The maximum of 999 events has been selected.";
		OpenHR.messageBox(sMessage, 64, "Calendar Reports");
	}

	refreshTab2Controls();

}

function openDialogCalEvent(pDestination, pWidth, pHeight, psResizable, psScroll) {
	var dlgwinprops = "center:yes;" +
		"dialogHeight:" + pHeight + "px;" +
		"dialogWidth:" + pWidth + "px;" +
		"help:no;" +
		"resizable:" + psResizable + ";" +
		"scroll:" + psScroll + ";" +
		"status:no;";
	window.showModalDialog(pDestination, self, dlgwinprops);
}
</script>

<div id="divCalendarReportDefinition">

	<form id="frmDefinition" name="frmDefinition">
		<table  >
			<tr>
				<td>
					<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0" >
						<tr height="5">
							<td colspan="3"></td>
						</tr>

						<tr height="10">
							<td width="10"></td>
							<td>
								<input type="button" value="Definition" id="btnTab1" name="btnTab1" disabled="disabled" class="btn btndisabled"
									onclick="display_CalendarReport_Page(1)"/>
								<input type="button" value="Event Details" id="btnTab2" name="btnTab2" class="btn"
									onclick="display_CalendarReport_Page(2)"/>
								<input type="button" value="Report Details" id="btnTab3" name="btnTab3" class="btn"
									onclick="display_CalendarReport_Page(3)"/>
								<input type="button" value="Sort Order" id="btnTab4" name="btnTab4" class="btn"
									onclick="display_CalendarReport_Page(4)"/>
								<input type="button" value="Output" id="btnTab5" name="btnTab5" class="btn"
									onclick="display_CalendarReport_Page(5)"/>
							</td>
							<td width="10"></td>
						</tr>

						<tr height="10">
							<td colspan="3"></td>
						</tr>

						<tr>
							<td></td>
							<td>
								<!-- First tab -->
								<div id="div1">
									<table width="1000px" height="80%"  cellspacing="0" cellpadding="5" >
										<tr valign="top">
											<td>
												<table width="100%" class="invisible" cellspacing="0" cellpadding="0" >
													<tr>
														<td colspan="10" height="5"></td>
													</tr>

													<tr height="10">
														<td ></td>
														<td >Name :</td>
														<td></td>
														<td style="padding-right:10px">
															<input id="txtName" name="txtName" maxlength="50" style="width: 100%" class="text"
																onkeyup="changeTab1Control()">
														</td>
														<td></td>
														<td>Owner :</td>
														<td><input id="txtOwner" name="txtOwner" style="width:100%" 
																placeholder="Enter Owner"
																class="text textdisabled"
																disabled="disabled"></td>
														<td>
															
														</td>
														<td></td>
													</tr>

													<tr>
														<td colspan="9" height="5"></td>
													</tr>

													<tr style="height:150px">
														<td ></td>
														<td nowrap valign="top">Description :</td>
														<td ></td>
														<td style="vertical-align: top; width:40%;padding-right: 10px" rowspan="1" colspan="1">
															<textarea id="txtDescription" name="txtDescription" class="textarea" style="HEIGHT: 99%; WIDTH: 100%" wrap="VIRTUAL" height="0" maxlength="255"
																onkeyup="changeTab1Control()">
													</textarea>
														</td>
														<td ></td>
														<td width="10" valign="top" nowrap>Access :</td>
														
														<td rowspan="1" style="vertical-align: top; width: 40%;">
															<%Html.RenderPartial("Util_Def_CustomReports/grdaccess")%> 
															
														</td>
                                                        
														<td></td>
													</tr>

													<tr height="10">
														<td ></td>
														<td style="vertical-align: top">Base Table :</td>
														<td ></td>
														<td style="vertical-align: top; width:40%;padding-right: 10px" rowspan="1" colspan="1">
															<select id="cboBaseTable" name="cboBaseTable" style="WIDTH: 100%" class="combo combodisabled"
																onchange="changeBaseTable()" disabled="disabled">
															</select>
														</td>
														<td></td>
														<td width="10" valign="top">Records :</td>
														<td width="40%" colspan="2">
															<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																<tr>
																	<td width="5" nowrap>
																		<input checked id="optRecordSelection1" name="optRecordSelection" type="radio"
																			onclick="changeBaseTableRecordOptions()" />
																	</td>
																	<td></td>
																	<td colspan="4">
																		<label tabindex="-1"
																			for="optRecordSelection1"
																			class="radio">
																			All</label>
																	</td>
																	<td colspan="3">&nbsp;</td>
																</tr>
																<tr>
																	<td colspan="6" height="5"></td>
																</tr>
																<tr>
																	<td width="5" nowrap>
																		<input id="optRecordSelection2" name="optRecordSelection" type="radio"
																			onclick="changeBaseTableRecordOptions()" />
																	</td>
																	<td></td>
																	<td width="5">
																		<label
																			tabindex="-1"
																			for="optRecordSelection2"
																			class="radio" >
																			Picklist</label>
																	</td>
																	<td width="5">&nbsp;</td>
																	<td width="100%">
																		<input id="txtBasePicklist" name="txtBasePicklist" disabled="disabled" class="text textdisabled" style="WIDTH: 98%">
																	</td>
																	<td>
																		<input id="cmdBasePicklist" name="cmdBasePicklist"  type="button" disabled="disabled" class="btn btndisabled" value="..."
																			onclick="selectRecordOptionCalDef('base', 'picklist')" style="WIDTH: 100%"/>
																	</td>
																</tr>
																<tr>
																	<td colspan="6" height="5"></td>
																</tr>
																<tr>
																	<td width="5" >
																		<input id="optRecordSelection3" name="optRecordSelection" type="radio"
																			onclick="changeBaseTableRecordOptions()" />
																	</td>
																	<td></td>
																	<td width="5">
																		<label
																			tabindex="-1"
																			for="optRecordSelection3"
																			class="radio" >
																			Filter</label>
																	</td>
																	<td width="5"></td>
																	<td width="100%">
																		<input id="txtBaseFilter" name="txtBaseFilter" class="text textdisabled" disabled="disabled" style="WIDTH: 98%">
																	</td>
																	<td>
																		<input id="cmdBaseFilter" name="cmdBaseFilter"  type="button" disabled="disabled" value="..." class="btn btndisabled"
																			onclick="selectRecordOptionCalDef('base', 'filter')"/>
																	</td>
																</tr>
															</table>
														</td>
														<td></td>
													</tr>
													<tr>
														<td colspan="9" height="5"></td>
													</tr>
                                                    <tr>
														<td></td>
														<td >Description 1 :</td>
														<td></td>
														<td style="vertical-align: top; width:40%;padding-right: 10px" rowspan="1" colspan="1">
															<select id="cboDescription1" name="cboDescription1" style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
																onchange="changeTab1Control();">
															</select>
														</td>
														<td></td>
														<td></td>
														<td style="white-space: nowrap" colspan="2">
															<input name="chkPrintFilterHeader" id="chkPrintFilterHeader" type="checkbox" disabled="disabled" tabindex="0"
																onclick="changeTab1Control();">
															<label
																id="lblPrintFilterHeader"
																name="lblPrintFilterHeader"
																for="chkPrintFilterHeader"
																class="checkbox checkboxdisabled"
																tabindex="-1"

																style="white-space: nowrap">
																Display filter or picklist title in the report header 
															</label>
														</td>
														<td></td>
													</tr>

                                                    <tr>
														<td></td>
														<td>Description 2 :</td>
														<td></td>
														<td style="vertical-align: top; width:40%;padding-right: 10px" rowspan="1" colspan="1" >
															<select id="cboDescription2" name="cboDescription2" style="WIDTH: 100%" disabled="disabled" class="combo combodisabled"
																onchange="changeTab1Control();">
															</select>
														</td>
														<td></td>
														<td width="10" valign="top">Region :</td>

														<td colspan ="2">
															<select id="cboRegion" name="cboRegion" style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
																onchange="changeTab1Control();refreshTab3Controls();">
															</select>
														</td>
														
                                                        <td></td>
													</tr>

													<tr>
														<td></td>
														<td style="width:11%">Description 3 :</td>
														<td></td>
														<td style="vertical-align: top; width:40%;padding-right: 10px" rowspan="1" colspan="1">
                                                            <table style="width:100%">
                                                                <tr>
                                                                    <td style="width:90%">
															            <input id="txtDescExpr" name="txtDescExpr" disabled="disabled" class="text textdisabled" style="width:100%">
															        </td>
                                                                    <td style="width:10%">
                                                                        <input id="cmdDescExpr" name="cmdDescExpr" style="width:100%" type="button" disabled="disabled" class="btn btndisabled" value="..."
																        onclick="selectCalc('baseDesc', false)" />
                                                                     </td>
                                                                    </tr>
                                                            </table>
														</td>
														<td>
															
														</td>
														<td></td>
														<td style="text-align: left; width: 15px; vertical-align: central" id="qq" colspan ="2">
															<input valign="center" name="chkGroupByDesc" id="chkGroupByDesc" type="checkbox" disabled="disabled" tabindex="0"
																onclick="changeTab1Control(); refreshTab3Controls();" />
                                                            <label
																for="chkGroupByDesc"
																class="checkbox checkboxdisabled"
																tabindex="-1">
																Group By Description</label>
														</td>
														<td>															
														</td>
                                                       													
																													
													
														
													</tr>

													<tr>
														<td></td>
														<td></td>
														<td></td>
														<td></td>
														<td></td>
														<td>Separator :</td>
														<td colspan="2"> 
															<select name="cboDescriptionSeparator" id="cboDescriptionSeparator" style="WIDTH: 100%" disabled="disabled" class="combo combodisabled"
																onchange="changeTab1Control();">
																<option value="">
																&lt;None&gt;
																	<option value=" ">
																&lt;Space&gt;
																	<option value=", ">
																, 
																	<option value=".  ">
																.  
																	<option value=" - ">
																- 
																	<option value=" : ">
																: 
																	<option value=" ; ">
																; 
																	<option value=" / ">
																/ 
																	<option value=" \ ">
																\ 
																	<option value=" # ">
																# 
																	<option value=" ~ ">
																~ 
																	<option value=" ^ ">
																^       
															</select>
														</td>
														<td></td>
													</tr>
												</table>
											</td>
										</tr>
									</table>

									
										
												
											
									
								</div>

								<!-- Second tab -->
								<div id="div2" style="visibility: hidden; display: none">
									<table width="100%" height="100%"  cellspacing="0" cellpadding="5">
										<tr valign="top">
											<td>
												<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
													<tr>
														<td colspan="3" height="5"></td>
													</tr>

													<tr height="10">
														<td width="5">&nbsp;</td>
														<td width="90" nowrap colspan="7">Events :</td>
														<td width="5">&nbsp;</td>
													</tr>
													<tr>
														<td width="5">&nbsp;</td>
														<td colspan="7">
															<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
																<tr>
																	<td colspan="3" height="5"></td>
																</tr>
																<tr height="5">
																	<td rowspan="8" style="width: 700px">
																		<%Dim avColumnDef(26, 4)
																			
																			avColumnDef(0, 0) = "Name"			 'name
																			avColumnDef(0, 1) = "Name"			 'caption
																			avColumnDef(0, 2) = "2000"			 'width
																			avColumnDef(0, 3) = "-1"				 'visible
	
																			avColumnDef(1, 0) = "TableID"			 'name
																			avColumnDef(1, 1) = "TableID"			 'caption
																			avColumnDef(1, 2) = "1814"			 'width
																			avColumnDef(1, 3) = "0"					 'visible

																			avColumnDef(2, 0) = "Table"				 'name
																			avColumnDef(2, 1) = "Table"				 'caption
																			avColumnDef(2, 2) = "2000"			 'width
																			avColumnDef(2, 3) = "-1"				 'visible

																			avColumnDef(3, 0) = "FilterID"		 'name
																			avColumnDef(3, 1) = "FilterID"		 'caption
																			avColumnDef(3, 2) = "1814"			 'width
																			avColumnDef(3, 3) = "0"					 'visible

																			avColumnDef(4, 0) = "Filter"			 'name
																			avColumnDef(4, 1) = "Filter"			 'caption
																			avColumnDef(4, 2) = "2000"			 'width
																			avColumnDef(4, 3) = "-1"				 'visible

																			avColumnDef(5, 0) = "StartDateID"		 'name
																			avColumnDef(5, 1) = "StartDateID"		 'caption
																			avColumnDef(5, 2) = "1814"			 'width
																			avColumnDef(5, 3) = "0"					 'visible

																			avColumnDef(6, 0) = "Start Date"		 'name
																			avColumnDef(6, 1) = "Start Date"		 'caption
																			avColumnDef(6, 2) = "2100"			 'width
																			avColumnDef(6, 3) = "-1"				 'visible

																			avColumnDef(7, 0) = "StartSessionID"	 'name
																			avColumnDef(7, 1) = "StartSessionID"	 'caption
																			avColumnDef(7, 2) = "1814"			 'width
																			avColumnDef(7, 3) = "0"					 'visible
	
																			avColumnDef(8, 0) = "Start Session"		 'name
																			avColumnDef(8, 1) = "Start Session"		 'caption
																			avColumnDef(8, 2) = "2600"			 'width
																			avColumnDef(8, 3) = "-1"				 'visible
	
																			avColumnDef(9, 0) = "EndDateID"			 'name
																			avColumnDef(9, 1) = "EndDateID"			 'caption
																			avColumnDef(9, 2) = "1814"			 'width
																			avColumnDef(9, 3) = "0"					 'visible
	
																			avColumnDef(10, 0) = "End Date"			 'name
																			avColumnDef(10, 1) = "End Date"			 'caption
																			avColumnDef(10, 2) = "2100"				 'width
																			avColumnDef(10, 3) = "-1"				 'visible
	
																			avColumnDef(11, 0) = "EndSessionID"		 'name
																			avColumnDef(11, 1) = "EndSessionID"		 'caption
																			avColumnDef(11, 2) = "1814"				 'width
																			avColumnDef(11, 3) = "0"				 'visible
	
																			avColumnDef(12, 0) = "End Session"	 'name
																			avColumnDef(12, 1) = "End Session"	 'caption
																			avColumnDef(12, 2) = "2600"				 'width
																			avColumnDef(12, 3) = "-1"				 'visible
	
																			avColumnDef(13, 0) = "DurationID"		 'name
																			avColumnDef(13, 1) = "DurationID"		 'caption
																			avColumnDef(13, 2) = "1814"				 'width
																			avColumnDef(13, 3) = "0"				 'visible
	
																			avColumnDef(14, 0) = "Duration"			 'name
																			avColumnDef(14, 1) = "Duration"			 'caption
																			avColumnDef(14, 2) = "2000"				 'width
																			avColumnDef(14, 3) = "-1"				 'visible
	
																			avColumnDef(15, 0) = "LegendType"		 'name
																			avColumnDef(15, 1) = "LegendType"		 'caption
																			avColumnDef(15, 2) = "1814"				 'width
																			avColumnDef(15, 3) = "0"				 'visible
	
																			avColumnDef(16, 0) = "Legend"			 'name
																			avColumnDef(16, 1) = "Key"			 'caption
																			avColumnDef(16, 2) = "2000"				 'width
																			avColumnDef(16, 3) = "-1"				 'visible
	
																			avColumnDef(17, 0) = "LegendTableID"	 'name
																			avColumnDef(17, 1) = "LegendTableID"	 'caption
																			avColumnDef(17, 2) = "1814"				 'width
																			avColumnDef(17, 3) = "0"				 'visible
	
																			avColumnDef(18, 0) = "LegendColumnID"	 'name
																			avColumnDef(18, 1) = "LegendColumnID"	 'caption
																			avColumnDef(18, 2) = "1814"				 'width
																			avColumnDef(18, 3) = "0"				 'visible
	
																			avColumnDef(19, 0) = "LegendCodeID"		 'name
																			avColumnDef(19, 1) = "LegendCodeID"		 'caption
																			avColumnDef(19, 2) = "1814"				 'width
																			avColumnDef(19, 3) = "0"				 'visible
	
																			avColumnDef(20, 0) = "LegendEventTypeID" 'name
																			avColumnDef(20, 1) = "LegendEventTypeID" 'caption
																			avColumnDef(20, 2) = "1814"				 'width
																			avColumnDef(20, 3) = "0"				 'visible
	
																			avColumnDef(21, 0) = "Desc1ID"		 'name
																			avColumnDef(21, 1) = "Desc1ID"		 'caption
																			avColumnDef(21, 2) = "1814"				 'width
																			avColumnDef(21, 3) = "0"				 'visible
	
																			avColumnDef(22, 0) = "Description 1"	 'name
																			avColumnDef(22, 1) = "Description 1"	 'caption
																			avColumnDef(22, 2) = "2600"				 'width
																			avColumnDef(22, 3) = "-1"				 'visible
	
																			avColumnDef(23, 0) = "Desc2ID"		 'name
																			avColumnDef(23, 1) = "Desc2ID"		 'caption
																			avColumnDef(23, 2) = "1814"				 'width
																			avColumnDef(23, 3) = "0"				 'visible
	
																			avColumnDef(24, 0) = "Description 2"	 'name
																			avColumnDef(24, 1) = "Description 2"	 'caption
																			avColumnDef(24, 2) = "2600"				 'width
																			avColumnDef(24, 3) = "-1"				 'visible
	
																			avColumnDef(25, 0) = "EventKey"			 'name
																			avColumnDef(25, 1) = "EventKey"			 'caption
																			avColumnDef(25, 2) = "1395"				 'width
																			avColumnDef(25, 3) = "0"				 'visible
	
																			avColumnDef(26, 0) = "FilterHidden"			 'name
																			avColumnDef(26, 1) = "FilterHidden"			 'caption
																			avColumnDef(26, 2) = "1395"				 'width
																			avColumnDef(26, 3) = "0"				 'visible
			
																			Response.Write("											<OBJECT classid=clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" & vbCrLf)
																			Response.Write("													 codebase=""cabs/COAInt_Grid.cab#version=3,1,3,6""" & vbCrLf)
																			Response.Write("													height=""300px""" & vbCrLf)
																			Response.Write("													id=grdEvents" & vbCrLf)
																			Response.Write("													name=grdEvents" & vbCrLf)
																			Response.Write("													style=""HEIGHT: 300px; VISIBILITY: visible; WIDTH: 100%"">" & vbCrLf)
																			'Response.Write("													width=""300px"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""ScrollBars"" VALUE=""4"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""_Version"" VALUE=""196617"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""DataMode"" VALUE=""2"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""Cols"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""Rows"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""BorderStyle"" VALUE=""1"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""RecordSelectors"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""GroupHeaders"" VALUE=""-1"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""ColumnHeaders"" VALUE=""-1"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""GroupHeadLines"" VALUE=""1"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""HeadLines"" VALUE=""1"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""FieldDelimiter"" VALUE=""(None)"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""FieldSeparator"" VALUE=""(Tab)"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""Row.Count"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""Col.Count"" VALUE=""1"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""TagVariant"" VALUE=""EMPTY"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""UseGroups"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""HeadFont3D"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""Font3D"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""DividerType"" VALUE=""3"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""DividerStyle"" VALUE=""1"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""DefColWidth"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""BeveColorScheme"" VALUE=""2"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""BevelColorFrame"" VALUE=""-2147483642"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""BevelColorHighlight"" VALUE=""-2147483628"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""BevelColorShadow"" VALUE=""-2147483632"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""BevelColorFace"" VALUE=""-2147483633"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""CheckBox3D"" VALUE=""-1"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""AllowAddNew"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""AllowDelete"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""AllowUpdate"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""MultiLine"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""ActiveCellStyleSet"" VALUE="""">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""RowSelectionStyle"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""AllowRowSizing"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""AllowGroupSizing"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""AllowColumnSizing"" VALUE=""-1"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""AllowGroupMoving"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""AllowColumnMoving"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""AllowGroupSwapping"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""AllowColumnSwapping"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""AllowGroupShrinking"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""AllowColumnShrinking"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""AllowDragDrop"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""UseExactRowCount"" VALUE=""-1"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""SelectTypeCol"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""SelectTypeRow"" VALUE=""3"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""SelectByCell"" VALUE=""-1"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""BalloonHelp"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""RowNavigation"" VALUE=""2"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""CellNavigation"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""MaxSelectedRows"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""HeadStyleSet"" VALUE="""">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""StyleSet"" VALUE="""">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""ForeColorEven"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""ForeColorOdd"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""BackColorEven"" VALUE=""16777215"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""BackColorOdd"" VALUE=""16777215"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""Levels"" VALUE=""1"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""RowHeight"" VALUE=""503"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""ExtraHeight"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""ActiveRowStyleSet"" VALUE="""">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""CaptionAlignment"" VALUE=""2"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""SplitterPos"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""SplitterVisible"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""Columns.Count"" VALUE=""" & (UBound(avColumnDef) + 1) & """>" & vbCrLf)
	
																			For i = 0 To UBound(avColumnDef) Step 1
																				Response.Write("												<!--" & avColumnDef(i, 0) & "-->  " & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").Width"" VALUE=""" & avColumnDef(i, 2) & """>" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").Visible"" VALUE=""" & avColumnDef(i, 3) & """>" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").Columns.Count"" VALUE=""1"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").Caption"" VALUE=""" & avColumnDef(i, 1) & """>" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").Name"" VALUE=""" & avColumnDef(i, 0) & """>" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").Alignment"" VALUE=""0"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").CaptionAlignment"" VALUE=""3"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").Bound"" VALUE=""0"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").AllowSizing"" VALUE=""1"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").DataField"" VALUE=""Column " & i & """>" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").DataType"" VALUE=""8"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").Level"" VALUE=""0"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").NumberFormat"" VALUE="""">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").Case"" VALUE=""0"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").FieldLen"" VALUE=""256"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").VertScrollBar"" VALUE=""0"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").Locked"" VALUE=""0"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").Style"" VALUE=""0"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").ButtonsAlways"" VALUE=""0"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").RowCount"" VALUE=""0"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").ColCount"" VALUE=""1"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").HasHeadForeColor"" VALUE=""0"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").HasHeadBackColor"" VALUE=""0"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").HasForeColor"" VALUE=""0"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").HasBackColor"" VALUE=""0"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").HeadForeColor"" VALUE=""0"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").HeadBackColor"" VALUE=""0"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").ForeColor"" VALUE=""0"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").BackColor"" VALUE=""0"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").HeadStyleSet"" VALUE="""">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").StyleSet"" VALUE="""">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").Nullable"" VALUE=""1"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").Mask"" VALUE="""">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").PromptInclude"" VALUE=""0"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").ClipMode"" VALUE=""0"">" & vbCrLf)
																				Response.Write("												<PARAM NAME=""Columns(" & i & ").PromptChar"" VALUE=""95"">" & vbCrLf)
																			Next
		
																			Response.Write("												<PARAM NAME=""UseDefaults"" VALUE=""-1"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""TabNavigation"" VALUE=""1"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""BatchUpdate"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""_ExtentX"" VALUE=""11298"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""_ExtentY"" VALUE=""3969"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""_StockProps"" VALUE=""79"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""Caption"" VALUE="""">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""ForeColor"" VALUE=""0"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""BackColor"" VALUE=""-2147483633"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""Enabled"" VALUE=""-1"">" & vbCrLf)
																			Response.Write("												<PARAM NAME=""DataMember"" VALUE="""">" & vbCrLf)

																			Response.Write("												<PARAM NAME=""Row.Count"" VALUE=""0"">" & vbCrLf)
																			Response.Write("											</OBJECT>" & vbCrLf)%>											
																	</td>

																	<td width="10">&nbsp;</td>
																	<td width="80">
																		<input type="button" id="cmdAddEvent" name="cmdAddEvent" value="Add..." style="WIDTH: 100%" class="btn"
																			onclick="eventAdd()" />
																	</td>
																	<td width="5">&nbsp;</td>
																</tr>

																<tr height="5">
																	<td colspan="3"></td>
																</tr>

																<tr height="5">

																	<td width="5">&nbsp;</td>
																	<td width="80">
																		<input type="button" id="cmdEditEvent" name="cmdChildEvent" value="Edit..." style="WIDTH: 100%" class="btn"
																			onclick="eventEdit()"/>
																	</td>
																	<td width="5">&nbsp;</td>
																</tr>

																<tr height="5">
																	<td colspan="3"></td>
																</tr>

																<tr height="5">

																	<td width="5">&nbsp;</td>
																	<td width="80">
																		<input type="button" id="cmdRemoveEvent" name="cmdRemoveEvent" value="Remove" style="WIDTH: 100%" class="btn"
																			onclick="eventRemove()" />
																	</td>
																	<td width="5">&nbsp;</td>
																</tr>

																<tr height="5">
																	<td colspan="3"></td>
																</tr>

																<tr height="5">

																	<td width="5">&nbsp;</td>
																	<td width="80">
																		<input type="button" id="cmdRemoveAllEvents" name="cmdRemoveAllEvents" value="Remove All" style="WIDTH: 100%" class="btn"
																			onclick="eventRemoveAll()" />
																	</td>
																	<td width="5">&nbsp;</td>
																</tr>

																<tr>
																	<td colspan="3"></td>
																</tr>
															</table>
														</td>
														<td width="5">&nbsp;</td>
													</tr>
													<tr>
														<td colspan="9" height="5"></td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
								</div>

								<!-- Third tab -->
								<div id="div3" style="visibility: hidden; display: none">
									<table width="100%" height="100%"  cellspacing="0" cellpadding="5">
										<tr valign="top">
											<td>
												<table width="100%" class="invisible" cellspacing="10" cellpadding="0">
													<tr>
														<td valign="top" rowspan="1" width="50%" height="100%">
															<table cellspacing="0" cellpadding="4" width="100%" height="100%">
																<tr height="10">
																	<td style="height:10px; text-align:  left; vertical-align:top"><strong>Start Date :</strong>
																		<br>
																		<br>
																		<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																			<tr height="20">
																				<td width="5">&nbsp;</td>
																				<td align="left" width="15">
																					<input type="radio" name="optStart" id="optFixedStart"
																						onclick="changeTab3Control();" />
																				</td>
																				<td width="10">&nbsp;</td>
																				<td align="left" width="15" nowrap>
																					<label
																						tabindex="-1"
																						for="optFixedStart"
																						class="radio">
																						Fixed</label>
																				</td>
																				<td width="30" nowrap>&nbsp;</td>
																				<td align="left" width="100%">
																					<input type="text" id="txtFixedStart" name="txtFixedStart" value="" style="WIDTH: 100%" class="text"
																						onkeyup="changeTab3Control()">
																				</td>
																				<td width="5">&nbsp;</td>
																			</tr>
																			<tr height="5">
																				<td colspan="7"></td>
																			</tr>
																			<tr height="20">
																				<td width="5">&nbsp;</td>
																				<td align="left" width="15">
																					<input type="radio" name="optStart" id="optCurrentStart"
																						onclick="changeTab3Control();"/>
																				</td>
																				<td width="10">&nbsp;</td>
																				<td colspan="3" align="left" nowrap>
																					<label
																						tabindex="-1"
																						for="optCurrentStart"
																						class="radio">
																						Current Date</label>
																				</td>
																				<td width="5">&nbsp;</td>
																			</tr>
																			<tr height="5">
																				<td colspan="7"></td>
																			</tr>
																			<tr height="20">
																				<td width="5">&nbsp;</td>
																				<td align="left" width="15" nowrap>
																					<input type="radio" name="optStart" id="optOffsetStart"
																						onclick="changeTab3Control();"/>
																				</td>
																				<td width="10">&nbsp;</td>
																				<td align="left" width="15" nowrap>
																					<label
																						tabindex="-1"
																						for="optOffsetStart"
																						class="radio"
																						onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
																						Offset</label>
																				</td>
																				<td width="30" nowrap>&nbsp;</td>
																				<td align="left" width="100%">
																					<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																						<tr>
																							<td width="40">
																								<input id="txtFreqStart" maxlength="4" name="txtFreqStart" style="WIDTH: 40px" width="40" value="0" class="text"
																									onkeyup="setRecordsNumeric(frmDefinition.txtFreqStart);changeTab3Control();validateOffsets();"
																									onchange="setRecordsNumeric(frmDefinition.txtFreqStart);changeTab3Control();validateOffsets();">
																							</td>
																							<td>
																								<input style="WIDTH: 15px" type="button" value="+" id="cmdPeriodStartUp" name="cmdPeriodStartUp" class="btn"
																									onclick="spinRecords(true, frmDefinition.txtFreqStart); setRecordsNumeric(frmDefinition.txtFreqStart); changeTab3Control(); validateOffsets();" />
																							</td>
																							<td>
																								<input style="WIDTH: 15px" type="button" value="-" id="cmdPeriodStartDown" name="cmdPeriodStartDown" class="btn"
																									onclick="spinRecords(false, frmDefinition.txtFreqStart); setRecordsNumeric(frmDefinition.txtFreqStart); changeTab3Control(); validateOffsets();" />
																							</td>
																							<td width="10">&nbsp;</td>
																							<td width="100%">
																								<select name="cboPeriodStart" id="cboPeriodStart" style="WIDTH: 100%" width="100%" class="combo"
																									onchange="changeTab3Control();validateOffsets();">
																									<option name="Day" value="0" selected>
																									Day(s)
																							<option name="Week" value="1">
																									Week(s)
																							<option name="Month" value="2">
																									Month(s)
																							<option name="Year" value="3">
																									Year(s)
																								</select>
																							</td>
																						</tr>
																					</table>
																				</td>
																				<td width="5">&nbsp;</td>
																			</tr>
																			<tr height="5">
																				<td colspan="7"></td>
																			</tr>
																			<tr height="20">
																				<td width="5">&nbsp;</td>
																				<td align="left" width="15" nowrap>
																					<input type="radio" name="optStart" id="optCustomStart"
																						onclick="changeTab3Control();" />
																				</td>
																				<td width="10">&nbsp;</td>
																				<td align="left" width="15" nowrap>
																					<label
																						tabindex="-1"
																						for="optCustomStart"
																						class="radio">
																						Custom</label>
																				</td>
																				<td width="30" nowrap>&nbsp;</td>
																				<td align="left" width="100%">
																					<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																						<tr>
																							<td>
																								<input id="txtCustomStart" name="txtCustomStart" disabled="disabled" style="WIDTH: 100%" class="text textdisabled">
																							</td>
																							<td width="30">
																								<input id="cmdCustomStart" name="cmdCustomStart" style="WIDTH: 100%" type="button" disabled="disabled" value="..." class="btn btndisabled"
																									onclick="selectCalc('startDate', true)" />
																							</td>
																						</tr>
																					</table>
																				</td>
																				<td width="5">&nbsp;</td>
																			</tr>
																			<tr height="5">
																				<td colspan="7"></td>
																			</tr>
																		</table>
																	</td>
																</tr>
															</table>
														</td>
														<td style="vertical-align: top; width:50%">
															<table style="padding: 2px; width:100%; height:100%">
																<tr height="10">
																	<td style="height:10px; text-align:  left; vertical-align:top"><strong>End Date :</strong>
																		<br>
																		<br>
																		<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																			<tr height="20">
																				<td width="5">&nbsp;</td>
																				<td align="left" width="15">
																					<input type="radio" name="optEnd" id="optFixedEnd"
																						onclick="changeTab3Control();" />
																				</td>
																				<td width="10">&nbsp;</td>
																				<td align="left" width="15" nowrap>
																					<label
																						tabindex="-1"
																						for="optFixedEnd"
																						class="radio">
																						Fixed</label>
																				</td>
																				<td width="30" nowrap>&nbsp;</td>
																				<td align="left" width="100%">
																					<input type="text" id="txtFixedEnd" name="txtFixedEnd" value="" style="WIDTH: 100%" class="text"
																						onkeyup="changeTab3Control()">
																				</td>
																				<td width="5">&nbsp;</td>
																			</tr>
																			<tr height="5">
																				<td colspan="7"></td>
																			</tr>
																			<tr height="20">
																				<td width="5">&nbsp;</td>
																				<td align="left" width="15">
																					<input type="radio" name="optEnd" id="optCurrentEnd"
																						onclick="changeTab3Control();" />
																				</td>
																				<td width="10">&nbsp;</td>
																				<td colspan="3" align="left" nowrap>
																					<label
																						tabindex="-1"
																						for="optCurrentEnd"
																						class="radio">
																						Current Date</label>
																				</td>
																				<td width="5">&nbsp;</td>
																			</tr>
																			<tr height="5">
																				<td colspan="7"></td>
																			</tr>
																			<tr height="20">
																				<td width="5">&nbsp;</td>
																				<td align="left" width="15" nowrap>
																					<input type="radio" name="optEnd" id="optOffsetEnd"
																						onclick="changeTab3Control();"/>
																				</td>
																				<td width="10">&nbsp;</td>
																				<td align="left" width="15" nowrap>
																					<label
																						tabindex="-1"
																						for="optOffsetEnd"
																						class="radio">
																						Offset</label>
																				</td>
																				<td width="30" nowrap>&nbsp;</td>
																				<td align="left" width="100%">
																					<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																						<tr>
																							<td width="40">
																								<input id="txtFreqEnd" maxlength="4" name="txtFreqEnd" style="WIDTH: 40px" width="40" value="0" class="text"
																									onkeyup="setRecordsNumeric(frmDefinition.txtFreqEnd);changeTab3Control();validateOffsets();"
																									onchange="setRecordsNumeric(frmDefinition.txtFreqEnd);changeTab3Control();validateOffsets();">
																							</td>
																							<td>
																								<input style="WIDTH: 15px" type="button" value="+" id="cmdPeriodEndUp" name="cmdPeriodEndUp" class="btn"
																									onclick="spinRecords(true, frmDefinition.txtFreqEnd); setRecordsNumeric(frmDefinition.txtFreqEnd); changeTab3Control(); validateOffsets();" />
																							</td>
																							<td>
																								<input style="WIDTH: 15px" type="button" value="-" id="cmdPeriodEndDown" name="cmdPeriodEndDown" class="btn"
																									onclick="spinRecords(false, frmDefinition.txtFreqEnd); setRecordsNumeric(frmDefinition.txtFreqEnd); changeTab3Control(); validateOffsets();" />
																							</td>
																							<td width="10">&nbsp;</td>
																							<td width="100%">
																								<select name="cboPeriodEnd" id="cboPeriodEnd" style="WIDTH: 100%" width="100%" class="combo"
																									onchange="changeTab3Control();validateOffsets();">
																									<option name="Day" value="0" selected>
																									Day(s)
																							<option name="Week" value="1">
																									Week(s)
																							<option name="Month" value="2">
																									Month(s)
																							<option name="Year" value="3">
																									Year(s)
																								</select>
																							</td>
																						</tr>
																					</table>
																				</td>
																				<td width="5">&nbsp;</td>
																			</tr>
																			<tr height="5">
																				<td colspan="7"></td>
																			</tr>
																			<tr height="20">
																				<td width="5">&nbsp;</td>
																				<td align="left" width="15" nowrap>
																					<input type="radio" name="optEnd" id="optCustomEnd"
																						onclick="changeTab3Control();" />
																				</td>
																				<td width="10">&nbsp;</td>
																				<td align="left" width="15" nowrap>
																					<label
																						tabindex="-1"
																						for="optCustomEnd"
																						class="radio">
																						Custom
																					</label>
																				</td>
																				<td width="30" nowrap>&nbsp;</td>
																				<td align="left" width="100%">
																					<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																						<tr>
																							<td>
																								<input id="txtCustomEnd" name="txtCustomEnd" disabled="disabled" class="text textdisabled" style="WIDTH: 100%">
																							</td>
																							<td width="30">
																								<input id="cmdCustomEnd" name="cmdCustomEnd" style="WIDTH: 100%" type="button" disabled="disabled" value='...' class="btn btndisabled"
																									onclick="selectCalc('endDate', true)" />
																							</td>
																						</tr>
																					</table>
																				</td>
																				<td width="5">&nbsp;</td>
																			</tr>
																			<tr height="5">
																				<td colspan="7"></td>
																			</tr>
																		</table>
																	</td>
																</tr>
															</table>
														</td>
													</tr>
													<tr>
														<td valign="top" width="100%" colspan="2">
															<table cellspacing="0" cellpadding="2" width="100%" height="100%">
																<tr height="10">
																	<td style="height:10px; text-align:  left; vertical-align:top"><strong>Default Display Options :</strong>
																		<br>
																		<br>
																		<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																			<tr>
																				<td width="5">&nbsp;</td>
																				<td>
																					<input name="chkIncludeBHols" id="chkIncludeBHols" type="checkbox" disabled="disabled" tabindex="0"
																						onclick="changeTab3Control();" />
																					<label
																						for="chkIncludeBHols"
																						class="checkbox"
																						tabindex="-1">
																					</label>
																					<label
																						for="chkIncludeBHols"
																						class="ui-state-error-text"
																						tabindex="-1">
																						Include Bank Holidays *
																					</label>
																				</td>
																				<td width="5">&nbsp;</td>
																			</tr>
																			<tr height="2">
																				<td colspan="3"></td>
																			</tr>
																			<tr>
																				<td width="5">&nbsp;</td>
																				<td>
																					<input name="chkIncludeWorkingDaysOnly" id="chkIncludeWorkingDaysOnly" type="checkbox" disabled="disabled" tabindex="0"
																						onclick="changeTab3Control();" />
																					<label
																						for="chkIncludeWorkingDaysOnly"
																						class="checkbox"
																						tabindex="-1">
																					</label>
																					<label
																						for="chkIncludeWorkingDaysOnly"
																						class="ui-state-error-text"
																						tabindex="-1">
																						Working Days Only *
																					</label>
																				</td>
																				<td width="5">&nbsp;</td>
																			</tr>
																			<tr height="2">
																				<td colspan="7"></td>
																			</tr>
																			<tr>
																				<td width="5">&nbsp;</td>
																				<td>
																					<input name="chkShadeBHols" id="chkShadeBHols" type="checkbox" disabled="disabled" tabindex="0"
																						onclick="changeTab3Control();" />
																					<label
																						for="chkShadeBHols"
																						class="checkbox"
																						tabindex="-1">
																					</label>
																					<label
																						for="chkShadeBHols"
																						class="ui-state-error-text"
																						tabindex="-1">
																						Show Bank Holidays *
																					</label>
																				</td>
																				<td width="5">&nbsp;</td>
																			</tr>
																			<tr height="2">
																				<td colspan="7"></td>
																			</tr>
																			<tr>
																				<td width="5">&nbsp;</td>
																				<td>
																					<input name="chkCaptions" id="chkCaptions" type="checkbox" disabled="disabled" tabindex="0"
																						onclick="changeTab3Control();" />
																					<label
																						for="chkCaptions"
																						class="checkbox"
																						tabindex="-1">
																					</label>
																					<label
																						for="chkCaptions"
																						class="ui-state-error-text"
																						tabindex="-1">
																						Show Calendar Captions *
																					</label>
																			</td>
																				<td width="5">&nbsp;</td>
																			</tr>
																			<tr height="2">
																				<td colspan="7"></td>
																			</tr>
																			<tr>
																				<td width="5">&nbsp;</td>
																				<td>
																					<input name="chkShadeWeekends" id="chkShadeWeekends" type="checkbox" disabled="disabled" tabindex="0"
																						onclick="changeTab3Control();"/>
																					<label
																						for="chkShadeWeekends"
																						class="checkbox checkboxdisabled"
																						tabindex="-1">
																						Show Weekends 
																					</label>
																				</td>
																				<td width="5">&nbsp;</td>
																			</tr>
																			<tr height="2">
																				<td colspan="7"></td>
																			</tr>
																			<tr>
																				<td width="5">&nbsp;</td>
																				<td>
																					<input name="chkStartOnCurrentMonth" id="chkStartOnCurrentMonth" type="checkbox" disabled="disabled" tabindex="0"
																						onclick="changeTab3Control();"/>
																					<label
																						for="chkStartOnCurrentMonth"
																						class="checkbox checkboxdisabled"
																						tabindex="-1">
																						Start on Current Month 
																					</label>
																				</td>
																				<td width="5">&nbsp;</td>
																			</tr>
																			<tr>
																				<td colspan="7">&nbsp;</td>
																			</tr>
																			<tr>
																				<td colspan="7">&nbsp;</td>
																			</tr>
																			<tr height="2">
																				<td colspan="7">
																					<label class="ui-state-error-text">* Not supported in OpenHR 8.0 Web</label>
																				</td>
																			</tr>
																		</table>
																	</td>
																</tr>
															</table>
														</td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
								</div>

								<!-- Fourth tab -->
								<div id="div4" style="visibility: hidden; display: none">
									<table width="100%" height="100%"  cellspacing="0" cellpadding="5">
										<tr valign="top">
											<td>
												<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
													<tr>
														<td colspan="5" height="5"></td>
													</tr>

													<tr height="20">
														<td width="5">&nbsp;</td>
														<td width="90" colspan="3" nowrap>Sort Order :</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td width="5">&nbsp;</td>
														<td rowspan="12">
															<%Html.RenderPartial("Util_Def_CustomReports/ssCalOleDBGridSortOrder")%>
														</td>

														<td width="10">&nbsp;</td>
														<td width="100">
															<input type="button" id="cmdSortAdd" name="cmdSortAdd" value="Add..." style="WIDTH: 100%" class="btn"
																onclick="sortAdd()" />
														</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td colspan="4">&nbsp;</td>
													</tr>

													<tr height="5">
														<td width="5">&nbsp;</td>
														<td width="5">&nbsp;</td>
														<td width="100">
															<input type="button" id="cmdSortEdit" name="cmdSortEdit" value="Edit..." style="WIDTH: 100%" class="btn"
																onclick="sortEdit()" />
														</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td colspan="4">&nbsp;</td>
													</tr>

													<tr height="5">
														<td width="5">&nbsp;</td>

														<td width="5">&nbsp;</td>
														<td width="100">
															<input type="button" id="cmdSortRemove" name="cmdSortRemove" value="Remove" style="WIDTH: 100%" class="btn"
																onclick="sortRemove()" />
														</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td colspan="4">&nbsp;</td>
													</tr>

													<tr height="5">
														<td width="5">&nbsp;</td>

														<td width="5">&nbsp;</td>
														<td width="100">
															<input type="button" id="cmdSortRemoveAll" name="cmdSortRemoveAll" value="Remove All" style="WIDTH: 100%" class="btn"
																onclick="sortRemoveAll()" />
														</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td colspan="4">&nbsp;</td>
													</tr>

													<tr height="5">
														<td width="5">&nbsp;</td>
														<td width="5">&nbsp;</td>
														<td width="100">
															<input type="button" id="cmdSortMoveUp" name="cmdSortMoveUp" value="Move Up" style="WIDTH: 100%" class="btn"
																onclick="sortMove(true)" />
														</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td colspan="4">&nbsp;</td>
													</tr>

													<tr height="5">
														<td width="5">&nbsp;</td>
														<td width="5">&nbsp;</td>
														<td width="100">
															<input type="button" id="cmdSortMoveDown" name="cmdSortMoveDown" value="Move Down" style="WIDTH: 100%" class="btn"
																onclick="sortMove(false)" />
														</td>

														<td width="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td colspan="5"></td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
								</div>

								<!-- Fifth tab -->
								<!-- OUTPUT OPTIONS -->
								<div id="div5" style="visibility: hidden; display: none">
									<table width="100%" height="100%" cellspacing="0" cellpadding="5">
										<tr valign="top">
											<td>
												<table width="100%" class="invisible" cellspacing="10" cellpadding="0">
													<tr>
														<td style="padding-right: 30px; vertical-align: top; width: 25%; height: 100%" rowspan="2">
															<table cellspacing="0" cellpadding="4" width="100%" height="100%">
																<tr height="10">
																	<td height="10" align="left" valign="top"><strong>Output Format :</strong>
																		<br>
																		<br>
																		<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																			<tr height="20">
																				<td width="5">&nbsp;</td>
																				<td align="left" width="15">
																					<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat0" value="0"
																						onclick="formatClick(0);"/>
																				</td>
																				<td style="text-align: left; white-space: nowrap; padding-left: 5px">
																					<label
																						tabindex="-1"
																						for="optOutputFormat0"
																						class="radio">
																						Data Only
																					</label>
																				</td>
																				<td width="5">&nbsp;</td>
																			</tr>
																			<tr height="10">
																				<td colspan="4"></td>
																			</tr>
																			<tr height="20">
																				<td width="5">&nbsp;</td>
																				<td align="left" width="15">
																					<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat2" value="2"
																						onclick="formatClick(2);" />
																				</td>
																				<td style="text-align: left; white-space: nowrap; padding-left: 5px">
																					<label
																						tabindex="-1"
																						for="optOutputFormat2"
																						class="radio ui-state-error-text">
																						HTML Document
																					</label>
																				</td>
																				<td width="5">&nbsp;</td>
																			</tr>
																			<tr height="10">
																				<td colspan="4"></td>
																			</tr>
																			<tr height="20">
																				<td width="5">&nbsp;</td>
																				<td align="left" width="15">
																					<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat3" value="3"
																						onclick="formatClick(3);"/>
																				</td>
																				<td style="text-align: left; white-space: nowrap; padding-left: 5px">
																					<label
																						tabindex="-1"
																						for="optOutputFormat3"
																						class="radio ui-state-error-text">
																						Word Document
																					</label>
																				</td>
																				<td width="5">&nbsp;</td>
																			</tr>
																			<tr height="10">
																				<td colspan="4"></td>
																			</tr>
																			<tr height="20">
																				<td width="5">&nbsp;</td>
																				<td align="left" width="15">
																					<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat4" value="4"
																						onclick="formatClick(4);"/>
																				</td>
																				<td style="text-align: left; white-space: nowrap; padding-left: 5px">
																					<label
																						tabindex="-1"
																						for="optOutputFormat4"
																						class="radio">
																						Excel Worksheet</label>
																				</td>
																				<td width="5">&nbsp;</td>
																			</tr>
																			<tr height="10">
																				<td colspan="4"></td>
																			</tr>
																			<tr height="20">
																				<td width="5">&nbsp;</td>
																				<td align="left" width="15">
																					<input type="radio" width="20" name="optOutputFormat" id="optOutputFormat1" value="1" style="visibility: hidden"
																						onclick="formatClick(1);" />
																				</td>
																				<td align="left" nowrap>
																				</td>
																				<td width="5">&nbsp;</td>
																			</tr>
																			<tr height="10">
																				<td colspan="4"></td>
																			</tr>
																			<tr height="5">
																				<td width="5">&nbsp;</td>
																				<td align="left" width="15">
																					<input type="radio" width="20" name="optOutputFormat" id="optOutputFormat5" value="5" style="visibility: hidden"
																						onclick="formatClick(5);" />
																				</td>
																				<td>
																				</td>
																				<td width="5">&nbsp;</td>
																			</tr>
																			<tr height="10">
																				<td colspan="4"></td>
																			</tr>
																			<tr height="5">
																				<td width="5">&nbsp;</td>
																				<td align="left" width="15">
																					<input type="radio" width="20" name="optOutputFormat" id="optOutputFormat6" value="6" style="visibility: hidden"
																						onclick="formatClick(6);"/>
																				</td>
																				<td>

																				</td>
																				<td width="5">&nbsp;</td>
																			</tr>
																			<tr height="5">
																				<td colspan="4"></td>
																			</tr>
																		</table>
																	</td>
																</tr>
															</table>
														</td>
														<td style="width: 75%">
															<table cellspacing="0" cellpadding="4" width="100%" height="100%">
																<tr height="10">
																	<td height="10" align="left" valign="top"><strong>Output Destination(s) :</strong>
																		<br>
																		<br>
																		<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																			<tr height="20">
																				<td width="5">&nbsp;</td>
																				<td align="left" colspan="6" nowrap>
																					<input name="chkPreview" id="chkPreview" type="checkbox" disabled="disabled" tabindex="0"
																						onclick="changeTab5Control();" />
																					<label
																						for="chkPreview"
																						class="checkbox checkboxdisabled"
																						tabindex="-1">
																						Preview on screen
																					</label>
																				</td>
																				<td width="5">&nbsp;</td>
																			</tr>

																			<tr height="10">
																				<td colspan="8"></td>
																			</tr>

																			<tr height="20">
																				<td></td>
																				<td align="left" colspan="6" nowrap>
																					<input name="chkDestination0" id="chkDestination0" type="checkbox" disabled="disabled" tabindex="0"
																						onclick="changeTab5Control();" />
																					<label
																						for="chkDestination0"
																						class="checkbox checkboxdisabled"
																						tabindex="-1">
																						Display output on screen
																					</label>
																				</td>
																				<td></td>
																			</tr>

																			<tr height="10">
																				<td colspan="8"></td>
																			</tr>

																			<tr height="40px">
																				<td></td>
																				<td align="left" nowrap>
																					<input name="chkDestination1" id="chkDestination1" type="checkbox" disabled="disabled" tabindex="0"
																						onclick="changeTab5Control();" />
																					<label
																						for="chkDestination1"
																						class="checkbox checkboxdisabled "
																						tabindex="-1">
																					</label>
																					<label
																						for="chkDestination1"
																						class="ui-state-error-text"
																						tabindex="-1">
																						Send to printer 
																					</label>
																				</td>
																				<td width="30" nowrap>&nbsp;</td>
																				<td style="width: 15%">&nbsp;</td>
																				<td style="width: 25%; text-align: left; white-space: nowrap" class="ui-state-error-text">Printer location : </td>
																				<td style="width: 100%">
																					<select id="cboPrinterName" name="cboPrinterName" class="combo"
																						style="width: 100%"
																						onchange="changeTab5Control()">
																					</select>
																				</td>
																				<td width="30" nowrap>&nbsp;</td>
																				<td width="30" nowrap>&nbsp;</td>
																			</tr>

																			<tr height="10">
																				<td colspan="8"></td>
																			</tr>

																			<tr style="height: 40px">
																				<td></td>
																				<td align="left" nowrap>
																					<input name="chkDestination2" id="chkDestination2" type="checkbox" disabled="disabled" tabindex="0"
																						onclick="changeTab5Control();" />
																					<label
																						for="chkDestination2"
																						class="checkbox checkboxdisabled"
																						tabindex="-1">
																						Save to file 
																					</label>
																				</td>
																				<td></td>
																				<td></td>
																				<td align="left" nowrap>File name : </td>

																				<td align="right">
																					<input id="txtFilename" name="txtFilename"
																						style="width: 84%;vertical-align:central;"
																						class="text textdisabled" disabled="disabled" tabindex="-1">
																				
																					<input id="cmdFilename" name="cmdFilename" class="btn btndisabled" type="button" value='...' disabled="disabled"
																						onclick="saveFile(); changeTab5Control();"style="width: 12%;"/>
																				</td>
                                                                                <td></td>
																				<td></td>
																			</tr>

																			<tr height="40px">
																				<td></td>
																				<td></td>
																				<td></td>
																				<td></td>
																				<td align="left" nowrap class="ui-state-error-text">If existing file :</td>

																				<td style="white-space: nowrap">
																					<select id="cboSaveExisting" name="cboSaveExisting"
																						style="width: 100%"
																						class="combo"
																						onchange="changeTab5Control()">
																					</select>
																				</td>
																				<td></td>
																				<td></td>
																			</tr>

																			<tr height="40px">
																				<td></td>
																				<td align="left" nowrap>
																					<input name="chkDestination3" id="chkDestination3" type="checkbox" disabled="disabled" tabindex="0"
																						onclick="changeTab5Control();"/>
																					<label for="chkDestination3"
																						class="checkbox checkboxdisabled"
																						tabindex="-1">
																						Send as email
																					</label>
																				</td>
																				<td></td>
																				<td></td>
																				<td align="left" nowrap>Email group :   </td>
																				<td <%--align="right"--%>>
																					<input id="txtEmailGroup" name="txtEmailGroup"
																						style="width: 84%; vertical-align:central;"
																						class="text textdisabled" disabled="disabled" tabindex="-1" >
																					<input id="txtEmailGroupID" name="txtEmailGroupID" type="hidden" class="text textdisabled" disabled="disabled" tabindex="-1">
																				
																					<input id="cmdEmailGroup" name="cmdEmailGroup" type="button" value='...' disabled="disabled" class="btn btndisabled"
																						onclick="selectEmailGroup(); changeTab5Control();" style="width: 12%;" />
																				</td>
																				<td></td>
                                                                                <td></td>
																			</tr>
																			<tr height="40px">
																				<td></td>
																				<td></td>
																				<td></td>
																				<td></td>
																				<td style="white-space: nowrap;">
																					<label for="txtEmailSubject"
																						tabindex="-1">
																						Email subject :</label>
																				</td>
																				<td style="padding-top: 5px">
																					<input id="txtEmailSubject"
																						style="width: 100%;"
																						class="text textdisabled" disabled="disabled" maxlength="255" name="txtEmailSubject"
																						onchange="frmUseful.txtChanged.value = 1;"
																						onkeydown="frmUseful.txtChanged.value = 1;">
																				</td>
																				<td></td>
																				<td></td>
																			</tr>
																			<tr height="40px">
																				<td style="width: 130px" colspan="1"></td>
																				<td></td>
																				<td></td>
																				<td></td>
																				<td style="white-space: nowrap;">
																					<label for="txtEmailAttachAs"
																						tabindex="-1">
																						Attach as : 
																					</label>
																				</td>
																				<td style="padding-top: 5px">
																					<input id="txtEmailAttachAs" maxlength="255"
																						style="width: 100%;"
																						class="text textdisabled" disabled="disabled" name="txtEmailAttachAs"
																						onchange="frmUseful.txtChanged.value = 1;"
																						onkeydown="frmUseful.txtChanged.value = 1;"></td>
																				<td></td>
																				<td></td>
																			</tr>
																		</table>
																	</td>
																</tr>
															</table>
														</td>
													</tr>
												</table>
											</td>
										<tr height="20">
											<td colspan="5" class="ui-state-error-text">Note: In OpenHR Web Output Format is restricted to Excel. Existing files will be overwritten. Send to Printer is not supported.</td>
										</tr>
										</tr>
									</table>
								</div>
							</td>
						</tr>
					</table>

		</table>

		<input type='hidden' id="txtBasePicklistID" name="txtBasePicklistID">
		<input type='hidden' id="txtBaseFilterID" name="txtBaseFilterID">
		<input type='hidden' id="txtDescExprID" name="txtDescExprID">
		<input type='hidden' id="txtCustomStartID" name="txtCustomStartID">
		<input type='hidden' id="txtCustomEndID" name="txtCustomEndID">
		<input type='hidden' id="txtDatabase" name="txtDatabase" value="<%=session("Database")%>">

		<input type='hidden' id="txtWordVer" name="txtWordVer" value="<%=Session("WordVer")%>">
		<input type='hidden' id="txtExcelVer" name="txtExcelVer" value="<%=Session("ExcelVer")%>">
		<input type='hidden' id="txtWordFormats" name="txtWordFormats" value="<%=Session("WordFormats")%>">
		<input type='hidden' id="txtExcelFormats" name="txtExcelFormats" value="<%=Session("ExcelFormats")%>">
		<input type='hidden' id="txtWordFormatDefaultIndex" name="txtWordFormatDefaultIndex" value="<%=Session("WordFormatDefaultIndex")%>">
		<input type='hidden' id="txtExcelFormatDefaultIndex" name="txtExcelFormatDefaultIndex" value="<%=Session("ExcelFormatDefaultIndex")%>">
		
		<div>
			<table>
				
				<tr height="10">
					<td width="10"></td>
					<td>
						<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
							<tr>
								<td>&nbsp;</td>
								<td width="80">
									<input type="button" id="cmdOK" name="cmdOK" value="OK" style="WIDTH: 100%" class="btn"
										onclick="okCalClick()" />
								</td>
								<td width="10"></td>
								<td width="80">
									<input type="button" id="cmdCancel" name="cmdCancel" value="Cancel" style="WIDTH: 100%" class="btn"
										onclick="cancelClick()"/>
								</td>
							</tr>
						</table>
					</td>
					<td width="10"></td>
				</tr>

				<tr height="5">
					<td colspan="3"></td>
				</tr>
			</table>
		</div>
</form>

<form id="frmTables" style="visibility: hidden; display: none">
	<%
		
		Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
		Dim sErrorDescription = ""
			
		Try
			Dim rstTablesInfo = objDataAccess.GetDataTable("sp_ASRIntGetTablesInfo", CommandType.StoredProcedure)

			For Each objRow As DataRow In rstTablesInfo.Rows
				Response.Write("<input type='hidden' id=txtTableName_" & objRow("tableID") & " name=txtTableName_" & objRow("tableID") & " value=""" & objRow("tableName") & """>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtTableType_" & objRow("tableID") & " name=txtTableType_" & objRow("tableID") & " value=" & objRow("tableType") & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtTableChildren_" & objRow("tableID") & " name=txtTableChildren_" & objRow("tableID") & " value=""" & objRow("childrenString") & """>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtTableChildrenNames_" & objRow("tableID") & " name=txtTableChildrenNames_" & objRow("tableID") & " value=""" & objRow("childrenNames") & """>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtTableParents_" & objRow("tableID") & " name=txtTableParents_" & objRow("tableID") & " value=""" & objRow("parentsString") & """>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtTableRelations_" & objRow("tableID") & " name=txtTableRelations_" & objRow("tableID") & " value=""" & objRow("relatedString") & """>" & vbCrLf)
			Next

			
		Catch ex As Exception
			sErrorDescription = "The tables information could not be retrieved." & vbCrLf & ex.Message

		End Try		
					
	%>
</form>

	<form action="default_Submit" method=post id=frmGoto name=frmGoto style="visibility:hidden;display:none">
		<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
	</form>


<form id="frmOriginalDefinition" name="frmOriginalDefinition" style="visibility: hidden; display: none">
	<%
		Dim sErrMsg = ""
	
		If LCase(Session("action")) <> "new" Then
			
			Dim prmErrMsg = New SqlParameter("psErrorMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prmName = New SqlParameter("psCalendarReportName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
			Dim prmOwner = New SqlParameter("psCalendarReportOwner", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
			Dim prmDescription = New SqlParameter("psCalendarReportDesc", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prmBaseTableID = New SqlParameter("piBaseTableID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmAllRecords = New SqlParameter("pfAllRecords", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmPicklistID = New SqlParameter("piPicklistID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmPicklistName = New SqlParameter("psPicklistName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
			Dim prmPicklistHidden = New SqlParameter("pfPicklistHidden", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmFilterID = New SqlParameter("piFilterID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmFilterName = New SqlParameter("psFilterName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
			Dim prmFilterHidden = New SqlParameter("pfFilterHidden", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmPrintFilterHeader = New SqlParameter("pfPrintFilterHeader", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmDesc1ID = New SqlParameter("piDesc1ID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmDesc2ID = New SqlParameter("piDesc2ID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmDescExprID = New SqlParameter("piDescExprID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmDescExprName = New SqlParameter("psDescExprName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
			Dim prmDescCalcHidden = New SqlParameter("pfDescCalcHidden", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmRegionID = New SqlParameter("piRegionID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmGroupByDesc = New SqlParameter("pfGroupByDesc", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmDescSeparator = New SqlParameter("pfDescSeparator", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
			Dim prmStartType = New SqlParameter("piStartType", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmFixedStart = New SqlParameter("pdFixedStart", SqlDbType.Date) With {.Direction = ParameterDirection.Output}
			Dim prmStartFrequency = New SqlParameter("piStartFrequency", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmStartPeriod = New SqlParameter("piStartPeriod", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmCustomStartID = New SqlParameter("piCustomStartID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmCustomStartName = New SqlParameter("psCustomStartName", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prmStartDateCalcHidden = New SqlParameter("pfStartDateCalcHidden", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmEndType = New SqlParameter("piEndType", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmFixedEnd = New SqlParameter("pdFixedEnd", SqlDbType.Date) With {.Direction = ParameterDirection.Output}
			Dim prmEndFrequency = New SqlParameter("piEndFrequency", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmEndPeriod = New SqlParameter("piEndPeriod", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmCustomEndID = New SqlParameter("piCustomEndID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmCustomEndName = New SqlParameter("psCustomEndName", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prmEndDateCalcHidden = New SqlParameter("pfEndDateCalcHidden", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmShadeBHols = New SqlParameter("pfShadeBHols", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmShowCaptions = New SqlParameter("pfShowCaptions", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmShadeWeekends = New SqlParameter("pfShadeWeekends", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmStartOnCurrentMonth = New SqlParameter("pfStartOnCurrentMonth", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmIncludeWorkingDaysOnly = New SqlParameter("pfIncludeWorkingDaysOnly", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmIncludeBHols = New SqlParameter("pfIncludeBHols", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmOutputPreview = New SqlParameter("pfOutputPreview", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmOutputFormat = New SqlParameter("piOutputFormat", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmOutputScreen = New SqlParameter("pfOutputScreen", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmOutputPrinter = New SqlParameter("pfOutputPrinter", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmOutputPrinterName = New SqlParameter("psOutputPrinterName", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prmOutputSave = New SqlParameter("pfOutputSave", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmOutputSaveExisting = New SqlParameter("piOutputSaveExisting", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmOutputEmail = New SqlParameter("pfOutputEmail", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmOutputEmailAddr = New SqlParameter("piOutputEmailAddr", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmOutputEmailAddrName = New SqlParameter("psOutputEmailName", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prmOutputEmailSubject = New SqlParameter("psOutputEmailSubject", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prmOutputEmailAttachAs = New SqlParameter("psOutputEmailAttachAs", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prmOutputFilename = New SqlParameter("psOutputFilename", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
			Dim prmTimestamp = New SqlParameter("piTimestamp", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

			Try				

				Dim rstDefinition = objDataAccess.GetFromSP("spASRIntGetCalendarReportDefinition", _
						New SqlParameter("@piCalendarReportID", SqlDbType.Int) With {.Value = CleanNumeric(Session("utilid"))}, _
						New SqlParameter("psCurrentUser", SqlDbType.VarChar, 255) With {.Value = Session("username")}, _
						New SqlParameter("psAction", SqlDbType.VarChar, 255) With {.Value = Session("action")}, _
						prmErrMsg, prmName, prmOwner, prmDescription, prmBaseTableID, _
						prmAllRecords, prmPicklistID, prmPicklistName, prmPicklistHidden, _
						prmFilterID, prmFilterName, prmFilterHidden, prmPrintFilterHeader, _
						prmDesc1ID, prmDesc2ID, prmDescExprID, prmDescExprName, prmDescCalcHidden, _
						prmRegionID, prmGroupByDesc, prmDescSeparator, prmStartType, prmFixedStart, _
						prmStartFrequency, prmStartPeriod, prmCustomStartID, prmCustomStartName, prmStartDateCalcHidden, _
						prmEndType, prmFixedEnd, prmEndFrequency, prmEndPeriod, prmCustomEndID, prmCustomEndName, prmEndDateCalcHidden, _
						prmShadeBHols, prmShowCaptions, prmShadeWeekends, prmStartOnCurrentMonth, prmIncludeWorkingDaysOnly, prmIncludeBHols, _
						prmOutputPreview, prmOutputFormat, prmOutputScreen, prmOutputPrinter, prmOutputPrinterName, prmOutputSave, prmOutputSaveExisting, _
						prmOutputEmail, prmOutputEmailAddr, prmOutputEmailAddrName, prmOutputEmailSubject, prmOutputEmailAttachAs, prmOutputFilename, prmTimestamp)

				
				Dim iHiddenEventFilterCount = 0
				Dim iCount = 0
				

				For Each objRow As DataRow In rstDefinition.Rows
					iCount += 1
					
					Response.Write("<input type='hidden' id=txtReportDefnEvent_" & iCount & " name=txtReportDefnEvent_" & iCount & " value=""" & Replace(objRow("definitionString").ToString(), """", "&quot;") & """>" & vbCrLf)
					
					If objRow("FilterHidden").ToString() = "Y" Then
						iHiddenEventFilterCount += 1
					End If
					
				Next
			
				Session("hiddenfiltercount") = iHiddenEventFilterCount
				Session("CalendarEventCount") = iCount
				' NB. IMPORTANT ADO NOTE.
				' When calling a stored procedure which returns a recordset AND has output parameters
				' you need to close the recordset and set it to nothing before using the output parameters. 
				If Len(prmErrMsg.Value) > 0 Then
					sErrMsg = CType(("'" & Session("utilname") & "' " & prmErrMsg.Value), String)
				End If

				Response.Write("<input type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Replace(prmName.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_Owner name=txtDefn_Owner value=""" & Replace(prmOwner.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_Description name=txtDefn_Description value=""" & Replace(prmDescription.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_BaseTableID name=txtDefn_BaseTableID value=" & prmBaseTableID.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_AllRecords name=txtDefn_AllRecords value=" & prmAllRecords.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_PicklistID name=txtDefn_PicklistID value=" & prmPicklistID.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_PicklistName name=txtDefn_PicklistName value=""" & Replace(prmPicklistName.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_PicklistHidden name=txtDefn_PicklistHidden value=" & prmPicklistHidden.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_FilterID name=txtDefn_FilterID value=" & prmFilterID.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_FilterName name=txtDefn_FilterName value=""" & Replace(prmFilterName.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_FilterHidden name=txtDefn_FilterHidden value=" & prmFilterHidden.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_PrintFilterHeader name=txtDefn_PrintFilterHeader value=" & prmPrintFilterHeader.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_Desc1ID name=txtDefn_Desc1ID value=" & prmDesc1ID.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_Desc2ID name=txtDefn_Desc2ID value=" & prmDesc2ID.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_DescExprID name=txtDefn_DescExprID value=" & prmDescExprID.Value & ">" & vbCrLf)
				If IsDBNull(prmDescExprName.Value) Then
					Response.Write("<input type='hidden' id=txtDefn_DescExprName name=txtDefn_DescExprName value="""">" & vbCrLf)
				Else
					Response.Write("<input type='hidden' id=txtDefn_DescExprName name=txtDefn_DescExprName value=""" & Replace(prmDescExprName.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
				End If
				Response.Write("<input type='hidden' id=txtDefn_DescExprHidden name=txtDefn_DescExprHidden value=" & prmDescCalcHidden.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_RegionID name=txtDefn_RegionID value=" & prmRegionID.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_GroupByDesc name=txtDefn_GroupByDesc value=" & prmGroupByDesc.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_DescSeparator name=txtDefn_DescSeparator value=""" & prmDescSeparator.Value & """>" & vbCrLf)

				Response.Write("<input type='hidden' id=txtDefn_StartType name=txtDefn_StartType value=" & prmStartType.Value & ">" & vbCrLf)
			
				If IsDBNull(prmFixedStart.Value) Then
					Response.Write("<input type='hidden' id=txtDefn_FixedStart name=txtDefn_FixedStart value="""">" & vbCrLf)
				Else
					Response.Write("<input type='hidden' id=txtDefn_FixedStart name=txtDefn_FixedStart value=" & ConvertSQLDateToLocale(prmFixedStart.Value) & ">" & vbCrLf)
				End If
				Response.Write("<input type='hidden' id=txtDefn_StartFrequency name=txtDefn_StartFrequency value=" & prmStartFrequency.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_StartPeriod name=txtDefn_StartPeriod value=" & prmStartPeriod.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_CustomStartID name=txtDefn_CustomStartID value=" & prmCustomStartID.Value & ">" & vbCrLf)
				If IsDBNull(prmCustomStartName.Value) Then
					Response.Write("<input type='hidden' id=txtDefn_CustomStartName name=txtDefn_CustomStartName value="""">" & vbCrLf)
				Else
					Response.Write("<input type='hidden' id=txtDefn_CustomStartName name=txtDefn_CustomStartName value=""" & Replace(prmCustomStartName.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
				End If
				Response.Write("<input type='hidden' id=txtDefn_CustomStartCalcHidden name=txtDefn_CustomStartCalcHidden value=" & prmStartDateCalcHidden.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_EndType name=txtDefn_EndType value=" & prmEndType.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_FixedEnd name=txtDefn_FixedEnd value=" & ConvertSQLDateToLocale(prmFixedEnd.Value) & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_EndFrequency name=txtDefn_EndFrequency value=" & prmEndFrequency.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_EndPeriod name=txtDefn_EndPeriod value=" & prmEndPeriod.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_CustomEndID name=txtDefn_CustomEndID value=" & prmCustomEndID.Value & ">" & vbCrLf)
				If IsDBNull(prmCustomEndName.Value) Then
					Response.Write("<input type='hidden' id=txtDefn_CustomEndName name=txtDefn_CustomEndName value="""">" & vbCrLf)
				Else
					Response.Write("<input type='hidden' id=txtDefn_CustomEndName name=txtDefn_CustomEndName value=""" & Replace(prmCustomEndName.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
				End If
				Response.Write("<input type='hidden' id=txtDefn_CustomEndCalcHidden name=txtDefn_CustomEndCalcHidden value=" & prmEndDateCalcHidden.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_ShadeBHols name=txtDefn_ShadeBHols value=" & prmShadeBHols.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_ShowCaptions name=txtDefn_ShowCaptions value=" & prmShowCaptions.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_ShadeWeekends name=txtDefn_ShadeWeekends value=" & prmShadeWeekends.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_StartOnCurrentMonth name=txtDefn_StartOnCurrentMonth value=" & prmStartOnCurrentMonth.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_IncludeWorkingDaysOnly name=txtDefn_IncludeWorkingDaysOnly value=" & prmIncludeWorkingDaysOnly.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_IncludeBHols name=txtDefn_IncludeBHols value=" & prmIncludeBHols.Value & ">" & vbCrLf)
			
				Response.Write("<input type='hidden' id=txtDefn_OutputPreview name=txtDefn_OutputPreview value=" & prmOutputPreview.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_OutputFormat name=txtDefn_OutputFormat value=" & prmOutputFormat.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_OutputScreen name=txtDefn_OutputScreen value=" & prmOutputScreen.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_OutputPrinter name=txtDefn_OutputPrinter value=" & prmOutputPrinter.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_OutputPrinterName name=txtDefn_OutputPrinterName value=""" & prmOutputPrinterName.Value & """>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_OutputSave name=txtDefn_OutputSave value=" & prmOutputSave.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_OutputSaveExisting name=txtDefn_OutputSaveExisting value=" & prmOutputSaveExisting.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_OutputEmail name=txtDefn_OutputEmail value=" & prmOutputEmail.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_OutputEmailAddr name=txtDefn_OutputEmailAddr value=" & prmOutputEmailAddr.Value & ">" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_OutputEmailAddrName name=txtDefn_OutputEmailName value=""" & Replace(prmOutputEmailAddrName.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_OutputEmailSubject name=txtDefn_OutputEmailSubject value=""" & Replace(prmOutputEmailSubject.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_OutputEmailAttachAs name=txtDefn_OutputEmailAttachAs value=""" & Replace(prmOutputEmailAttachAs.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
				Response.Write("<input type='hidden' id=txtDefn_OutputFilename name=txtDefn_OutputFilename value=""" & prmOutputFilename.Value & """>" & vbCrLf)
			
				Response.Write("<input type='hidden' id=txtDefn_Timestamp name=txtDefn_Timestamp value=" & prmTimestamp.Value & ">" & vbCrLf)

				'********************************************************************************

				Dim prmErrMsg2 = New SqlParameter("psErrorMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

				Dim rstOrder = objDataAccess.GetFromSP("spASRIntGetCalendarReportOrder", _
						New SqlParameter("@piCalendarReportID", SqlDbType.Int) With {.Value = CleanNumeric(Session("utilid"))}, _
						prmErrMsg2)
		
				iCount = 0
				For Each objRow As DataRow In rstOrder.Rows
					iCount += 1
					Response.Write("<input type='hidden' id=txtReportDefnOrder_" & iCount & " name=txtReportDefnOrder_" & iCount & " value=""" & objRow("orderString").ToString() & """>" & vbCrLf)
				Next

				Session("CalendarOrderCount") = iCount

				If Len(prmErrMsg2.Value) > 0 Then
					sErrMsg = "'" & Session("utilname") & "' " & prmErrMsg2.Value
				End If

	
			Catch ex As Exception
				sErrMsg = CType(("'" & Session("utilname") & "' report definition could not be read." & vbCrLf) & FormatError(Err.Description), String)
				Session("confirmtext") = sErrMsg
				Session("confirmtitle") = "OpenHR"
				Session("followpage") = "defsel"
				Session("reaction") = "CALENDARREPORTS"
				Response.Clear()
				Response.Redirect("confirmok")

			End Try

	
		Else
			Session("CalendarEventCount") = 0
			Session("CalendarOrderCount") = 0
			Session("hiddenfiltercount") = 0
		End If
	
	%>
</form>

<form id="frmAccess">
	<%
		
		Try
			
			Dim prmAccessUtilID = New SqlParameter("piID", SqlDbType.Int)
			Dim prmFromCopy = New SqlParameter("piFromCopy", SqlDbType.Int)
		
			sErrorDescription = ""

			If UCase(Session("action")) = "NEW" Then
				prmAccessUtilID.Value = 0
			Else
				prmAccessUtilID.Value = CleanNumeric(Session("utilid"))
			End If

			If UCase(Session("action")) = "COPY" Then
				prmFromCopy.Value = 1
			Else
				prmFromCopy.Value = 0
			End If

			Dim rstAccessInfo = objDataAccess.GetDataTable("spASRIntGetUtilityAccessRecords", CommandType.StoredProcedure _
				, New SqlParameter("piUtilityType", SqlDbType.Int) With {.Value = UtilityType.utlCalendarReport} _
				, prmAccessUtilID, prmFromCopy)

			Dim iCount = 0		
			For Each objRow As DataRow In rstAccessInfo.Rows
				Response.Write("<input type='hidden' id=txtAccess_" & iCount & " name=txtAccess_" & iCount & " value=""" & objRow("accessDefinition").ToString() & """>" & vbCrLf)
				iCount += 1				
			Next
			
		Catch ex As Exception
			sErrorDescription = "The access information could not be retrieved." & vbCrLf & FormatError(ex.Message)

		End Try
		
	%>
</form>

<form id="frmUseful" name="frmUseful" style="visibility: hidden; display: none">
	<input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
	<input type="hidden" id="txtLoading" name="txtLoading" value="Y">
	<input type="hidden" id="txtFirstLoad" name="txtFirstLoad" value="Y">
	<input type="hidden" id="txtCurrentBaseTableID" name="txtCurrentBaseTableID">
	<input type="hidden" id="txtAvailableColumnsLoaded" name="txtAvailableColumnsLoaded" value="0">
	<input type="hidden" id="txtEventsLoaded" name="txtEventsLoaded" value="0">
	<input type="hidden" id="txtSortLoaded" name="txtSortLoaded" value="0">
	<input type="hidden" id="txtChanged" name="txtChanged" value="0">
	<input type="hidden" id="txtUtilID" name="txtUtilID" value='<%=session("utilid")%>'>
	<input type="hidden" id="txtEventCount" name="txtEventCount" value='<%=session("CalendarEventCount")%>'>
	<input type="hidden" id="txtOrderCount" name="txtOrderCount" value='<%=session("CalendarOrderCount")%>'>
	<input type="hidden" id="txtHiddenEventFilterCount" name="txtHiddenEventFilterCount" value='<%=session("hiddenfiltercount")%>'>
	<input type="hidden" id="txtLockGridEvents" name="txtLockGridEvents" value="0">
	<%
		Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)

		Dim sParameterValue As String = objDatabase.GetModuleParameter("MODULE_PERSONNEL", "Param_TablePersonnel")
		Response.Write("<input type='hidden' id='txtPersonnelTableID' name='txtPersonnelTableID' value=" & sParameterValue & ">" & vbCrLf)
		
		Response.Write("<input type='hidden' id='txtErrorDescription' name='txtErrorDescription' value="""">" & vbCrLf)
		Response.Write("<input type='hidden' id='txtAction' name='txtAction' value=" & Session("action") & ">" & vbCrLf)
		
	%>
</form>

<form id="frmValidate" name="frmValidate" target="validate" method="post" action="util_validate_calendarreport" style="visibility: hidden; display: none">
	<input type="hidden" id="validateBaseFilter" name="validateBaseFilter" value="0">
	<input type="hidden" id="validateBasePicklist" name="validateBasePicklist" value="0">
	<input type="hidden" id="validateEmailGroup" name="validateEmailGroup" value="0">
	<input type="hidden" id="validateEventFilter" name="validateEventFilter" value="0">
	<input type="hidden" id="validateDescExpr" name="validateDescExpr" value="0">
	<input type="hidden" id="validateCustomStart" name="validateCustomStart" value="0">
	<input type="hidden" id="validateCustomEnd" name="validateCustomEnd" value="0">
	<input type="hidden" id="validateHiddenGroups" name="validateHiddenGroups" value=''>
	<input type="hidden" id="validateName" name="validateName" value=''>

	<input type="hidden" id="validateTimestamp" name="validateTimestamp" value=''>
	<input type="hidden" id="validateUtilID" name="validateUtilID" value=''>
</form>

<form id="frmSend" name="frmSend" method="post" action="util_def_calendarreport_Submit" style="visibility: hidden; display: none">

	<input type="hidden" id="txtSend_ID" name="txtSend_ID">
	<input type="hidden" id="txtSend_name" name="txtSend_name">
	<input type="hidden" id="txtSend_description" name="txtSend_description">
	<input type="hidden" id="txtSend_access" name="txtSend_access">
	<input type="hidden" id="txtSend_userName" name="txtSend_userName">
	<input type="hidden" id="txtSend_baseTable" name="txtSend_baseTable">
	<input type="hidden" id="txtSend_allRecords" name="txtSend_allRecords">
	<input type="hidden" id="txtSend_picklist" name="txtSend_picklist">
	<input type="hidden" id="txtSend_filter" name="txtSend_filter">
	<input type="hidden" id="txtSend_printFilterHeader" name="txtSend_printFilterHeader">
	<input type="hidden" id="txtSend_desc1" name="txtSend_desc1">
	<input type="hidden" id="txtSend_desc2" name="txtSend_desc2">
	<input type="hidden" id="txtSend_descExpr" name="txtSend_descExpr">
	<input type="hidden" id="txtSend_region" name="txtSend_region">
	<input type="hidden" id="txtSend_groupbydesc" name="txtSend_groupbydesc">
	<input type="hidden" id="txtSend_descseparator" name="txtSend_descseparator">

	<input type="hidden" id="txtSend_StartType" name="txtSend_StartType">
	<input type="hidden" id="txtSend_FixedStart" name="txtSend_FixedStart">
	<input type="hidden" id="txtSend_StartFrequency" name="txtSend_StartFrequency">
	<input type="hidden" id="txtSend_StartPeriod" name="txtSend_StartPeriod">
	<input type="hidden" id="txtSend_CustomStart" name="txtSend_CustomStart">
	<input type="hidden" id="txtSend_EndType" name="txtSend_EndType">
	<input type="hidden" id="txtSend_FixedEnd" name="txtSend_FixedEnd">
	<input type="hidden" id="txtSend_EndFrequency" name="txtSend_EndFrequency">
	<input type="hidden" id="txtSend_EndPeriod" name="txtSend_EndPeriod">
	<input type="hidden" id="txtSend_CustomEnd" name="txtSend_CustomEnd">

	<input type="hidden" id="txtSend_IncludeBHols" name="txtSend_IncludeBHols">
	<input type="hidden" id="txtSend_IncludeWorkingDaysOnly" name="txtSend_IncludeWorkingDaysOnly">
	<input type="hidden" id="txtSend_ShadeBHols" name="txtSend_ShadeBHols">
	<input type="hidden" id="txtSend_Captions" name="txtSend_Captions">
	<input type="hidden" id="txtSend_ShadeWeekends" name="txtSend_ShadeWeekends">
	<input type="hidden" id="txtSend_StartOnCurrentMonth" name="txtSend_StartOnCurrentMonth">

	<input type="hidden" id="txtSend_OutputPreview" name="txtSend_OutputPreview">
	<input type="hidden" id="txtSend_OutputFormat" name="txtSend_OutputFormat">
	<input type="hidden" id="txtSend_OutputScreen" name="txtSend_OutputScreen">
	<input type="hidden" id="txtSend_OutputPrinter" name="txtSend_OutputPrinter">
	<input type="hidden" id="txtSend_OutputPrinterName" name="txtSend_OutputPrinterName">
	<input type="hidden" id="txtSend_OutputSave" name="txtSend_OutputSave">
	<input type="hidden" id="txtSend_OutputSaveExisting" name="txtSend_OutputSaveExisting">
	<input type="hidden" id="txtSend_OutputEmail" name="txtSend_OutputEmail">
	<input type="hidden" id="txtSend_OutputEmailAddr" name="txtSend_OutputEmailAddr">
	<input type="hidden" id="txtSend_OutputEmailSubject" name="txtSend_OutputEmailSubject">
	<input type="hidden" id="txtSend_OutputEmailAttachAs" name="txtSend_OutputEmailAttachAs">
	<input type="hidden" id="txtSend_OutputFilename" name="txtSend_OutputFilename">

	<input type="hidden" id="txtSend_columns" name="txtSend_Events">
	<input type="hidden" id="txtSend_columns2" name="txtSend_Events2">

	<input type="hidden" id="txtSend_OrderString" name="txtSend_OrderString">

	<input type="hidden" id="txtSend_reaction" name="txtSend_reaction">

	<input type="hidden" id="txtSend_jobsToHide" name="txtSend_jobsToHide">
	<input type="hidden" id="txtSend_jobsToHideGroups" name="txtSend_jobsToHideGroups">
</form>

<form id="frmEventDetails" name="frmEventDetails" target="eventselection" action="util_def_calendarreportdates_main" method="post" style="visibility: hidden; display: none">
	<input type="hidden" id="eventAction" name="eventAction">
	<input type="hidden" id="eventName" name="eventName">
	<input type="hidden" id="eventID" name="eventID">
	<input type="hidden" id="eventTableID" name="eventTableID">
	<input type="hidden" id="eventTable" name="eventTable">
	<input type="hidden" id="eventFilterID" name="eventFilterID">
	<input type="hidden" id="eventFilter" name="eventFilter">
	<input type="hidden" id="eventFilterHidden" name="eventFilterHidden">

	<input type="hidden" id="eventStartDateID" name="eventStartDateID">
	<input type="hidden" id="eventStartDate" name="eventStartDate">
	<input type="hidden" id="eventStartSessionID" name="eventStartSessionID">
	<input type="hidden" id="eventStartSession" name="eventStartSession">

	<input type="hidden" id="eventEndDateID" name="eventEndDateID">
	<input type="hidden" id="eventEndDate" name="eventEndDate">
	<input type="hidden" id="eventEndSessionID" name="eventEndSessionID">
	<input type="hidden" id="eventEndSession" name="eventEndSession">

	<input type="hidden" id="eventDurationID" name="eventDurationID">
	<input type="hidden" id="eventDuration" name="eventDuration">

	<input type="hidden" id="eventLookupType" name="eventLookupType">
	<input type="hidden" id="eventKeyCharacter" name="eventKeyCharacter">
	<input type="hidden" id="eventLookupTableID" name="eventLookupTableID">
	<input type="hidden" id="eventLookupColumnID" name="eventLookupColumnID">
	<input type="hidden" id="eventLookupCodeID" name="eventLookupCodeID">
	<input type="hidden" id="eventTypeColumnID" name="eventTypeColumnID">

	<input type="hidden" id="eventDesc1ID" name="eventDesc1ID">
	<input type="hidden" id="eventDesc1" name="eventDesc1">
	<input type="hidden" id="eventDesc2ID" name="eventDesc2ID">
	<input type="hidden" id="eventDesc2" name="eventDesc2">

	<input type="hidden" id="relationNames" name="relationNames">
</form>

<form id="frmRecordSelection" name="frmRecordSelection" target="recordSelection" action="util_recordSelection" method="post" style="visibility: hidden; display: none">
	<input type="hidden" id="recSelType" name="recSelType">
	<input type="hidden" id="recSelTableID" name="recSelTableID">
	<input type="hidden" id="recSelCurrentID" name="recSelCurrentID">
	<input type="hidden" id="recSelTable" name="recSelTable">
	<input type="hidden" id="recSelDefOwner" name="recSelDefOwner">
</form>

<form id="frmCalcSelection" name="frmCalcSelection" target="calcSelection" action="util_calcSelection" method="post" style="visibility: hidden; display: none">
	<input type="hidden" id="calcSelRecInd" name="calcSelRecInd">
	<input type="hidden" id="calcSelType" name="calcSelType">
	<input type="hidden" id="calcSelTableID" name="calcSelTableID">
	<input type="hidden" id="calcSelCurrentID" name="calcSelCurrentID">
	<input type="hidden" id="Hidden1" name="recSelDefOwner">
</form>

<form id="frmEmailSelection" name="frmEmailSelection" target="emailSelection" action="util_emailSelection" method="post" style="visibility: hidden; display: none">
	<input type="hidden" id="EmailSelCurrentID" name="EmailSelCurrentID">
</form>

<form id="frmSortOrder" name="frmSortOrder" target="sortorderselection" action="util_sortorderselection" method="post" style="visibility: hidden; display: none">
	<input type="hidden" id="txtSortInclude" name="txtSortInclude">
	<input type="hidden" id="txtSortExclude" name="txtSortExclude">
	<input type="hidden" id="txtSortEditing" name="txtSortEditing">
	<input type="hidden" id="txtSortColumnID" name="txtSortColumnID">
	<input type="hidden" id="txtSortColumnName" name="txtSortColumnName">
	<input type="hidden" id="txtSortOrder" name="txtSortOrder">
</form>

<form id="frmSelectionAccess" name="frmSelectionAccess" style="visibility: hidden; display: none">
	<input type="hidden" id="forcedHidden" name="forcedHidden" value="N">
	<input type="hidden" id="baseHidden" name="baseHidden" value="N">
	<input type="hidden" id="eventHidden" name="eventHidden" value="0">
	<input type="hidden" id="descHidden" name="descHidden" value="N">
	<input type="hidden" id="calcStartDateHidden" name="calcStartDateHidden" value="N">
	<input type="hidden" id="calcEndDateHidden" name="calcEndDateHidden" value="N">
</form>

	<input type='hidden' id="txtTicker" name="txtTicker" value="0">
	<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">
</div>

<script type="text/javascript">
	util_def_calendarreport_window_onload();
	util_def_calendarreport_addhandlers();
</script>


