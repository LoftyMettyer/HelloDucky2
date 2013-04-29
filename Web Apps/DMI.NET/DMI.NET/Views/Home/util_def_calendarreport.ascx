<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script src="<%: Url.Content("~/Scripts/FormScripts/calendarreportdef.js")%>" type="text/javascript"></script>
<%Html.RenderPartial("Util_Def_CustomReports/dialog")%>

<div <%=session("BodyTag")%>>
	<form id="frmDefinition" name="frmDefinition">
		<table class="outline">
			<tr>
				<td>
					<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
						<tr height="5">
							<td colspan="3"></td>
						</tr>

						<tr height="10">
							<td width="10"></td>
							<td>
								<input type="button" value="Definition" id="btnTab1" name="btnTab1" disabled="disabled" class="btn btndisabled"
									onclick="displayPage(1)"
									onmouseover="try{button_onMouseOver(this);}catch(e){}"
									onmouseout="try{button_onMouseOut(this);}catch(e){}"
									onfocus="try{button_onFocus(this);}catch(e){}"
									onblur="try{button_onBlur(this);}catch(e){}" />
								<input type="button" value="Event Details" id="btnTab2" name="btnTab2" class="btn"
									onclick="displayPage(2)"
									onmouseover="try{button_onMouseOver(this);}catch(e){}"
									onmouseout="try{button_onMouseOut(this);}catch(e){}"
									onfocus="try{button_onFocus(this);}catch(e){}"
									onblur="try{button_onBlur(this);}catch(e){}" />
								<input type="button" value="Report Details" id="btnTab3" name="btnTab3" class="btn"
									onclick="displayPage(3)"
									onmouseover="try{button_onMouseOver(this);}catch(e){}"
									onmouseout="try{button_onMouseOut(this);}catch(e){}"
									onfocus="try{button_onFocus(this);}catch(e){}"
									onblur="try{button_onBlur(this);}catch(e){}" />
								<input type="button" value="Sort Order" id="btnTab4" name="btnTab4" class="btn"
									onclick="displayPage(4)"
									onmouseover="try{button_onMouseOver(this);}catch(e){}"
									onmouseout="try{button_onMouseOut(this);}catch(e){}"
									onfocus="try{button_onFocus(this);}catch(e){}"
									onblur="try{button_onBlur(this);}catch(e){}" />
								<input type="button" value="Output" id="btnTab5" name="btnTab5" class="btn"
									onclick="displayPage(5)"
									onmouseover="try{button_onMouseOver(this);}catch(e){}"
									onmouseout="try{button_onMouseOut(this);}catch(e){}"
									onfocus="try{button_onFocus(this);}catch(e){}"
									onblur="try{button_onBlur(this);}catch(e){}" />
							</td>
							<td width="10"></td>
						</tr>

						<tr height="10">
							<td colspan="3"></td>
						</tr>

						<tr>
							<td width="10"></td>
							<td>
								<!-- First tab -->
								<div id="div1">
									<table width="100%" height="80%" class="outline" cellspacing="0" cellpadding="5">
										<tr valign="top">
											<td>
												<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
													<tr>
														<td colspan="10" height="5"></td>
													</tr>

													<tr height="10">
														<td width="5"></td>
														<td width="10">Name :</td>
														<td width="5">&nbsp;</td>
														<td>
															<input id="txtName" name="txtName" maxlength="50" style="WIDTH: 100%" class="text"
																onkeyup="changeTab1Control()">
														</td>
														<td width="20"></td>
														<td width="10">Owner :</td>
														<td width="5">&nbsp;</td>
														<td>
															<input id="txtOwner" name="txtOwner" style="WIDTH: 100%" class="text textdisabled"
																disabled="disabled">
														</td>
														<td width="5"></td>
													</tr>

													<tr>
														<td colspan="9" height="5"></td>
													</tr>

													<tr height="60">
														<td width="5"></td>
														<td width="10" nowrap valign="top">Description :</td>
														<td width="5"></td>
														<td width="40%" rowspan="1" colspan="1">
															<textarea id="txtDescription" name="txtDescription" class="textarea" style="HEIGHT: 99%; WIDTH: 100%" wrap="VIRTUAL" height="0" maxlength="255"
																onkeyup="changeTab1Control()"
																onpaste="var selectedLength = document.selection.createRange().text.length;var pasteData = window.clipboardData.getData('Text');if ((this.value.length + pasteData.length - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}"
																onkeypress="var selectedLength = document.selection.createRange().text.length;if ((this.value.length + 1 - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}">
													</textarea>
														</td>
														<td width="20" nowrap></td>
														<td width="10" valign="top" nowrap>Access :</td>
														<td width="5"></td>
														<td width="40%" rowspan="1" valign="top" height="78">
															<%Html.RenderPartial("Util_Def_CustomReports/grdaccess")%> 
														</td>
														<td width="5"></td>
													</tr>

													<tr height="20">
														<td width="5">&nbsp;</td>
														<td colspan="7">
															<hr>
														</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr height="10">
														<td width="5">&nbsp;</td>
														<td width="85" nowrap valign="top">Base Table :</td>
														<td width="5"></td>
														<td width="40%" valign="top" colspan="1">
															<select id="cboBaseTable" name="cboBaseTable" style="WIDTH: 100%" class="combo combodisabled"
																onchange="changeBaseTable()" disabled="disabled">
															</select>
														</td>
														<td width="20" nowrap></td>
														<td width="10" valign="top">Records :</td>
														<td width="40%" colspan="2">
															<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																<tr>
																	<td width="5" nowrap>
																		<input checked id="optRecordSelection1" name="optRecordSelection" type="radio"
																			onclick="changeBaseTableRecordOptions()"
																			onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																			onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																			onfocus="try{radio_onFocus(this);}catch(e){}"
																			onblur="try{radio_onBlur(this);}catch(e){}" />
																	</td>
																	<td></td>
																	<td colspan="4">
																		<label tabindex="-1"
																			for="optRecordSelection1"
																			class="radio"
																			onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																			onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
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
																			onclick="changeBaseTableRecordOptions()"
																			onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																			onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																			onfocus="try{radio_onFocus(this);}catch(e){}"
																			onblur="try{radio_onBlur(this);}catch(e){}" />
																	</td>
																	<td></td>
																	<td width="5">
																		<label
																			tabindex="-1"
																			for="optRecordSelection2"
																			class="radio"
																			onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																			onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
																			Picklist</label>
																	</td>
																	<td width="5">&nbsp;</td>
																	<td width="100%">
																		<input id="txtBasePicklist" name="txtBasePicklist" disabled="disabled" class="text textdisabled" style="WIDTH: 100%">
																	</td>
																	<td>
																		<input id="cmdBasePicklist" name="cmdBasePicklist" style="WIDTH: 30px" type="button" disabled="disabled" class="btn btndisabled" value="..."
																			onclick="selectRecordOption('base', 'picklist')"
																			onmouseover="try{button_onMouseOver(this);}catch(e){}"
																			onmouseout="try{button_onMouseOut(this);}catch(e){}"
																			onfocus="try{button_onFocus(this);}catch(e){}"
																			onblur="try{button_onBlur(this);}catch(e){}" />
																	</td>
																</tr>
																<tr>
																	<td colspan="6" height="5"></td>
																</tr>
																<tr>
																	<td width="5" nowrap>
																		<input id="optRecordSelection3" name="optRecordSelection" type="radio"
																			onclick="changeBaseTableRecordOptions()"
																			onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																			onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																			onfocus="try{radio_onFocus(this);}catch(e){}"
																			onblur="try{radio_onBlur(this);}catch(e){}" />
																	</td>
																	<td></td>
																	<td width="5">
																		<label
																			tabindex="-1"
																			for="optRecordSelection3"
																			class="radio"
																			onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																			onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
																			Filter</label>
																	</td>
																	<td width="5"></td>
																	<td width="100%">
																		<input id="txtBaseFilter" name="txtBaseFilter" class="text textdisabled" disabled="disabled" style="WIDTH: 100%">
																	</td>
																	<td>
																		<input id="cmdBaseFilter" name="cmdBaseFilter" style="WIDTH: 30px" type="button" disabled="disabled" value="..." class="btn btndisabled"
																			onclick="selectRecordOption('base', 'filter')"
																			onmouseover="try{button_onMouseOver(this);}catch(e){}"
																			onmouseout="try{button_onMouseOut(this);}catch(e){}"
																			onfocus="try{button_onFocus(this);}catch(e){}"
																			onblur="try{button_onBlur(this);}catch(e){}" />
																	</td>
																</tr>
																<tr>
																	<td colspan="6">
																		<input name="chkPrintFilterHeader" id="chkPrintFilterHeader" type="checkbox" disabled="disabled" tabindex="-1"
																			onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																			onmouseout="try{checkbox_onMouseOut(this);}catch(e){}"
																			onclick="changeTab1Control();">
																		<label
																			id="lblPrintFilterHeader"
																			name="lblPrintFilterHeader"
																			for="chkPrintFilterHeader"
																			class="checkbox checkboxdisabled"
																			tabindex="0"
																			onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																			onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																			onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																			onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																			onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																			Display filter or picklist title in the report header 
																		</label>
																	</td>
																</tr>
															</table>
														</td>
														<td width="5">&nbsp;</td>
													</tr>
													<tr>
														<td colspan="9" height="5"></td>
													</tr>

													<tr height="5">
														<td width="5"></td>
														<td nowrap valign="top">Description 1 :</td>
														<td width="5"></td>
														<td valign="top" colspan="1">
															<select id="cboDescription1" name="cboDescription1" style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
																onchange="changeTab1Control();">
															</select>
														</td>
														<td width="20" nowrap></td>
														<td width="10" valign="top">Region :</td>
														<td width="5"></td>
														<td>
															<select id="cboRegion" name="cboRegion" style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
																onchange="changeTab1Control();refreshTab3Controls();">
															</select>
														</td>
														<td width="5"></td>
													</tr>
													<tr>
														<td colspan="9" height="3"></td>
													</tr>

													<tr height="10">
														<td width="5"></td>
														<td style="white-space: nowrap" valign="top">Description 2 :</td>
														<td width="5">&nbsp;</td>
														<td valign="top" colspan="1">
															<select id="cboDescription2" name="cboDescription2" width="100%" disabled="disabled" class="combo combodisabled"
																onchange="changeTab1Control();">
															</select>
														</td>
														<td width="20" nowrap></td>
														<td width="10" valign="top" colspan="3"></td>
														<td width="5"></td>
													</tr>
													<tr>
														<td colspan="9" height="3"></td>
													</tr>
													<tr height="10">
														<td width="5"></td>
														<td nowrap valign="top">Description 3 :</td>
														<td width="5" nowrap>&nbsp;</td>
														<td>
															<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																<tr>
																	<td width="100%">
																		<input id="txtDescExpr" name="txtDescExpr" disabled="disabled" class="text textdisabled" style="WIDTH: 100%">
																	</td>
																	<td width="30" nowrap>
																		<input id="cmdDescExpr" name="cmdDescExpr" style="WIDTH: 30px" type="button" disabled="disabled" class="btn btndisabled" value="..."
																			onclick="selectCalc('baseDesc', false)"
																			onmouseover="try{button_onMouseOver(this);}catch(e){}"
																			onmouseout="try{button_onMouseOut(this);}catch(e){}"
																			onfocus="try{button_onFocus(this);}catch(e){}"
																			onblur="try{button_onBlur(this);}catch(e){}" />
																	</td>
																</tr>
															</table>
														</td>
														<td width="5"></td>
														<td width="10" valign="top" colspan="3"></td>
														<td width="5"></td>
													</tr>
													<tr>
														<td colspan="9" height="3"></td>
													</tr>

													<tr height="10">
														<td style="column-span: all">
															<table style="width: 100%" class="invisible">
																<td width="5">&nbsp;</td>
																<td style="text-align: left; width: 15px; vertical-align: central" id="qq">
																	<input valign="center" name="chkGroupByDesc" id="chkGroupByDesc" type="checkbox" disabled="disabled" tabindex="-1"
																		onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																		onmouseout="try{checkbox_onMouseOut(this);}catch(e){}"
																		onclick="changeTab1Control(); refreshTab3Controls();" />
																</td>
																<td style="text-align: left; width: 150px; vertical-align: central">
																	<label
																		for="chkGroupByDesc"
																		class="checkbox checkboxdisabled"
																		tabindex="0"
																		onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																		onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																		onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																		onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																		onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																		Group By Description</label>
																</td>
																<td>&nbsp;</td>
																<td>Separator : </td>
																<td width="5">&nbsp;</td>
																<td style="vertical-align: central; width: 85%">
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
																</>
															</table>
														</td>
													</tr>
													<tr>
														<td colspan="9" height="5"></td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
								</div>

								<!-- Second tab -->
								<div id="div2" style="visibility: hidden; display: none">
									<table width="100%" height="100%" class="outline" cellspacing="0" cellpadding="5">
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
																	<td rowspan="8">
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
																			Response.Write("													height=""100%""" & vbCrLf)
																			Response.Write("													id=grdEvents" & vbCrLf)
																			Response.Write("													name=grdEvents" & vbCrLf)
																			Response.Write("													style=""HEIGHT: 100%; VISIBILITY: visible; WIDTH: 100%""" & vbCrLf)
																			Response.Write("													width=""100%"">" & vbCrLf)
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
																			onclick="eventAdd()"
																			onmouseover="try{button_onMouseOver(this);}catch(e){}"
																			onmouseout="try{button_onMouseOut(this);}catch(e){}"
																			onfocus="try{button_onFocus(this);}catch(e){}"
																			onblur="try{button_onBlur(this);}catch(e){}" />
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
																			onclick="eventEdit()"
																			onmouseover="try{button_onMouseOver(this);}catch(e){}"
																			onmouseout="try{button_onMouseOut(this);}catch(e){}"
																			onfocus="try{button_onFocus(this);}catch(e){}"
																			onblur="try{button_onBlur(this);}catch(e){}" />
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
																			onclick="eventRemove()"
																			onmouseover="try{button_onMouseOver(this);}catch(e){}"
																			onmouseout="try{button_onMouseOut(this);}catch(e){}"
																			onfocus="try{button_onFocus(this);}catch(e){}"
																			onblur="try{button_onBlur(this);}catch(e){}" />
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
																			onclick="eventRemoveAll()"
																			onmouseover="try{button_onMouseOver(this);}catch(e){}"
																			onmouseout="try{button_onMouseOut(this);}catch(e){}"
																			onfocus="try{button_onFocus(this);}catch(e){}"
																			onblur="try{button_onBlur(this);}catch(e){}" />
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
									<table width="100%" height="100%" class="outline" cellspacing="0" cellpadding="5">
										<tr valign="top">
											<td>
												<table width="100%" class="invisible" cellspacing="10" cellpadding="0">
													<tr>
														<td valign="top" rowspan="1" width="50%" height="100%">
															<table class="outline" cellspacing="0" cellpadding="4" width="100%" height="100%">
																<tr height="10">
																	<td height="10" align="left" valign="top">Start Date :
																		<br>
																		<br>
																		<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																			<tr height="20">
																				<td width="5">&nbsp;</td>
																				<td align="left" width="15">
																					<input type="radio" name="optStart" id="optFixedStart"
																						onclick="changeTab3Control();"
																						onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																						onfocus="try{radio_onFocus(this);}catch(e){}"
																						onblur="try{radio_onBlur(this);}catch(e){}" />
																				</td>
																				<td width="10">&nbsp;</td>
																				<td align="left" width="15" nowrap>
																					<label
																						tabindex="-1"
																						for="optFixedStart"
																						class="radio"
																						onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
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
																						onclick="changeTab3Control();"
																						onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																						onfocus="try{radio_onFocus(this);}catch(e){}"
																						onblur="try{radio_onBlur(this);}catch(e){}" />
																				</td>
																				<td width="10">&nbsp;</td>
																				<td colspan="3" align="left" nowrap>
																					<label
																						tabindex="-1"
																						for="optCurrentStart"
																						class="radio"
																						onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
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
																						onclick="changeTab3Control();"
																						onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																						onfocus="try{radio_onFocus(this);}catch(e){}"
																						onblur="try{radio_onBlur(this);}catch(e){}" />
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
																									onclick="spinRecords(true, frmDefinition.txtFreqStart); setRecordsNumeric(frmDefinition.txtFreqStart); changeTab3Control(); validateOffsets();"
																									onmouseover="try{button_onMouseOver(this);}catch(e){}"
																									onmouseout="try{button_onMouseOut(this);}catch(e){}"
																									onfocus="try{button_onFocus(this);}catch(e){}"
																									onblur="try{button_onBlur(this);}catch(e){}" />
																							</td>
																							<td>
																								<input style="WIDTH: 15px" type="button" value="-" id="cmdPeriodStartDown" name="cmdPeriodStartDown" class="btn"
																									onclick="spinRecords(false, frmDefinition.txtFreqStart); setRecordsNumeric(frmDefinition.txtFreqStart); changeTab3Control(); validateOffsets();"
																									onmouseover="try{button_onMouseOver(this);}catch(e){}"
																									onmouseout="try{button_onMouseOut(this);}catch(e){}"
																									onfocus="try{button_onFocus(this);}catch(e){}"
																									onblur="try{button_onBlur(this);}catch(e){}" />
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
																						onclick="changeTab3Control();"
																						onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																						onfocus="try{radio_onFocus(this);}catch(e){}"
																						onblur="try{radio_onBlur(this);}catch(e){}" />
																				</td>
																				<td width="10">&nbsp;</td>
																				<td align="left" width="15" nowrap>
																					<label
																						tabindex="-1"
																						for="optCustomStart"
																						class="radio"
																						onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
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
																									onclick="selectCalc('startDate', true)"
																									onmouseover="try{button_onMouseOver(this);}catch(e){}"
																									onmouseout="try{button_onMouseOut(this);}catch(e){}"
																									onfocus="try{button_onFocus(this);}catch(e){}"
																									onblur="try{button_onBlur(this);}catch(e){}" />
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
														<td valign="top" width="50%">
															<table class="outline" cellspacing="0" cellpadding="2" width="100%" height="100%">
																<tr height="10">
																	<td height="10" align="left" valign="top">End Date :
																		<br>
																		<br>
																		<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																			<tr height="20">
																				<td width="5">&nbsp;</td>
																				<td align="left" width="15">
																					<input type="radio" name="optEnd" id="optFixedEnd"
																						onclick="changeTab3Control();"
																						onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																						onfocus="try{radio_onFocus(this);}catch(e){}"
																						onblur="try{radio_onBlur(this);}catch(e){}" />
																				</td>
																				<td width="10">&nbsp;</td>
																				<td align="left" width="15" nowrap>
																					<label
																						tabindex="-1"
																						for="optFixedEnd"
																						class="radio"
																						onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
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
																						onclick="changeTab3Control();"
																						onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																						onfocus="try{radio_onFocus(this);}catch(e){}"
																						onblur="try{radio_onBlur(this);}catch(e){}" />
																				</td>
																				<td width="10">&nbsp;</td>
																				<td colspan="3" align="left" nowrap>
																					<label
																						tabindex="-1"
																						for="optCurrentEnd"
																						class="radio"
																						onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
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
																						onclick="changeTab3Control();"
																						onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																						onfocus="try{radio_onFocus(this);}catch(e){}"
																						onblur="try{radio_onBlur(this);}catch(e){}" />
																				</td>
																				<td width="10">&nbsp;</td>
																				<td align="left" width="15" nowrap>
																					<label
																						tabindex="-1"
																						for="optOffsetEnd"
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
																								<input id="txtFreqEnd" maxlength="4" name="txtFreqEnd" style="WIDTH: 40px" width="40" value="0" class="text"
																									onkeyup="setRecordsNumeric(frmDefinition.txtFreqEnd);changeTab3Control();validateOffsets();"
																									onchange="setRecordsNumeric(frmDefinition.txtFreqEnd);changeTab3Control();validateOffsets();">
																							</td>
																							<td>
																								<input style="WIDTH: 15px" type="button" value="+" id="cmdPeriodEndUp" name="cmdPeriodEndUp" class="btn"
																									onclick="spinRecords(true, frmDefinition.txtFreqEnd); setRecordsNumeric(frmDefinition.txtFreqEnd); changeTab3Control(); validateOffsets();"
																									onmouseover="try{button_onMouseOver(this);}catch(e){}"
																									onmouseout="try{button_onMouseOut(this);}catch(e){}"
																									onfocus="try{button_onFocus(this);}catch(e){}"
																									onblur="try{button_onBlur(this);}catch(e){}" />
																							</td>
																							<td>
																								<input style="WIDTH: 15px" type="button" value="-" id="cmdPeriodEndDown" name="cmdPeriodEndDown" class="btn"
																									onclick="spinRecords(false, frmDefinition.txtFreqEnd); setRecordsNumeric(frmDefinition.txtFreqEnd); changeTab3Control(); validateOffsets();"
																									onmouseover="try{button_onMouseOver(this);}catch(e){}"
																									onmouseout="try{button_onMouseOut(this);}catch(e){}"
																									onfocus="try{button_onFocus(this);}catch(e){}"
																									onblur="try{button_onBlur(this);}catch(e){}" />
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
																						onclick="changeTab3Control();"
																						onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																						onfocus="try{radio_onFocus(this);}catch(e){}"
																						onblur="try{radio_onBlur(this);}catch(e){}" />
																				</td>
																				<td width="10">&nbsp;</td>
																				<td align="left" width="15" nowrap>
																					<label
																						tabindex="-1"
																						for="optCustomEnd"
																						class="radio"
																						onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
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
																									onclick="selectCalc('endDate', true)"
																									onmouseover="try{button_onMouseOver(this);}catch(e){}"
																									onmouseout="try{button_onMouseOut(this);}catch(e){}"
																									onfocus="try{button_onFocus(this);}catch(e){}"
																									onblur="try{button_onBlur(this);}catch(e){}" />
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
															<table class="outline" cellspacing="0" cellpadding="2" width="100%" height="100%">
																<tr height="10">
																	<td height="10" align="left" valign="top" width="100%">Default Display Options :
																		<br>
																		<br>
																		<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																			<tr>
																				<td width="5">&nbsp;</td>
																				<td>
																					<input name="chkIncludeBHols" id="chkIncludeBHols" type="checkbox" disabled="disabled" tabindex="-1"
																						onclick="changeTab3Control();"
																						onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																					<label
																						for="chkIncludeBHols"
																						class="checkbox checkboxdisabled"
																						tabindex="0"
																						onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																						onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																						onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																						onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																						Include Bank Holidays 
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
																					<input name="chkIncludeWorkingDaysOnly" id="chkIncludeWorkingDaysOnly" type="checkbox" disabled="disabled" tabindex="-1"
																						onclick="changeTab3Control();"
																						onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																					<label
																						for="chkIncludeWorkingDaysOnly"
																						class="checkbox checkboxdisabled"
																						tabindex="0"
																						onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																						onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																						onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																						onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																						Working Days Only 
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
																					<input name="chkShadeBHols" id="chkShadeBHols" type="checkbox" disabled="disabled" tabindex="-1"
																						onclick="changeTab3Control();"
																						onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																					<label
																						for="chkShadeBHols"
																						class="checkbox checkboxdisabled"
																						tabindex="0"
																						onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																						onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																						onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																						onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																						Show Bank Holidays 
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
																					<input name="chkCaptions" id="chkCaptions" type="checkbox" disabled="disabled" tabindex="-1"
																						onclick="changeTab3Control();"
																						onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																					<label
																						for="chkCaptions"
																						class="checkbox checkboxdisabled"
																						tabindex="0"
																						onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																						onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																						onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																						onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																						Show Calendar Captions
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
																					<input name="chkShadeWeekends" id="chkShadeWeekends" type="checkbox" disabled="disabled" tabindex="-1"
																						onclick="changeTab3Control();"
																						onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																					<label
																						for="chkShadeWeekends"
																						class="checkbox checkboxdisabled"
																						tabindex="0"
																						onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																						onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																						onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																						onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
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
																					<input name="chkStartOnCurrentMonth" id="chkStartOnCurrentMonth" type="checkbox" disabled="disabled" tabindex="-1"
																						onclick="changeTab3Control();"
																						onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																					<label
																						for="chkStartOnCurrentMonth"
																						class="checkbox checkboxdisabled"
																						tabindex="0"
																						onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																						onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																						onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																						onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																						Start on Current Month 
																					</label>
																				</td>
																				<td width="5">&nbsp;</td>
																			</tr>
																			<tr height="2">
																				<td colspan="7"></td>
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
									<table width="100%" height="100%" class="outline" cellspacing="0" cellpadding="5">
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
															<%Html.RenderPartial("Util_Def_CustomReports/ssOleDBGridSortOrder")%>
														</td>

														<td width="10">&nbsp;</td>
														<td width="100">
															<input type="button" id="cmdSortAdd" name="cmdSortAdd" value="Add..." style="WIDTH: 100%" class="btn"
																onclick="sortAdd()"
																onmouseover="try{button_onMouseOver(this);}catch(e){}"
																onmouseout="try{button_onMouseOut(this);}catch(e){}"
																onfocus="try{button_onFocus(this);}catch(e){}"
																onblur="try{button_onBlur(this);}catch(e){}" />
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
																onclick="sortEdit()"
																onmouseover="try{button_onMouseOver(this);}catch(e){}"
																onmouseout="try{button_onMouseOut(this);}catch(e){}"
																onfocus="try{button_onFocus(this);}catch(e){}"
																onblur="try{button_onBlur(this);}catch(e){}" />
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
																onclick="sortRemove()"
																onmouseover="try{button_onMouseOver(this);}catch(e){}"
																onmouseout="try{button_onMouseOut(this);}catch(e){}"
																onfocus="try{button_onFocus(this);}catch(e){}"
																onblur="try{button_onBlur(this);}catch(e){}" />
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
																onclick="sortRemoveAll()"
																onmouseover="try{button_onMouseOver(this);}catch(e){}"
																onmouseout="try{button_onMouseOut(this);}catch(e){}"
																onfocus="try{button_onFocus(this);}catch(e){}"
																onblur="try{button_onBlur(this);}catch(e){}" />
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
																onclick="sortMove(true)"
																onmouseover="try{button_onMouseOver(this);}catch(e){}"
																onmouseout="try{button_onMouseOut(this);}catch(e){}"
																onfocus="try{button_onFocus(this);}catch(e){}"
																onblur="try{button_onBlur(this);}catch(e){}" />
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
																onclick="sortMove(false)"
																onmouseover="try{button_onMouseOver(this);}catch(e){}"
																onmouseout="try{button_onMouseOut(this);}catch(e){}"
																onfocus="try{button_onFocus(this);}catch(e){}"
																onblur="try{button_onBlur(this);}catch(e){}" />
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
									<table width="100%" height="100%" class="outline" cellspacing="0" cellpadding="5">
										<tr valign="top">
											<td>
												<table width="100%" class="invisible" cellspacing="10" cellpadding="0">
													<tr>
														<td valign="top" rowspan="2" width="25%" height="100%">
															<table class="outline" cellspacing="0" cellpadding="4" width="100%" height="100%">
																<tr height="10">
																	<td height="10" align="left" valign="top">Output Format :
																		<br>
																		<br>
																		<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																			<tr height="20">
																				<td width="5">&nbsp;</td>
																				<td align="left" width="15">
																					<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat0" value="0"
																						onclick="formatClick(0);"
																						onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																						onfocus="try{radio_onFocus(this);}catch(e){}"
																						onblur="try{radio_onBlur(this);}catch(e){}" />
																				</td>
																				<td align="left" nowrap>
																					<label
																						tabindex="-1"
																						for="optOutputFormat0"
																						class="radio"
																						onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
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
																						onclick="formatClick(2);"
																						onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																						onfocus="try{radio_onFocus(this);}catch(e){}"
																						onblur="try{radio_onBlur(this);}catch(e){}" />
																				</td>
																				<td align="left" nowrap>
																					<label
																						tabindex="-1"
																						for="optOutputFormat2"
																						class="radio"
																						onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
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
																						onclick="formatClick(3);"
																						onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																						onfocus="try{radio_onFocus(this);}catch(e){}"
																						onblur="try{radio_onBlur(this);}catch(e){}" />
																				</td>
																				<td align="left" nowrap>
																					<label
																						tabindex="-1"
																						for="optOutputFormat3"
																						class="radio"
																						onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
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
																						onclick="formatClick(4);"
																						onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																						onfocus="try{radio_onFocus(this);}catch(e){}"
																						onblur="try{radio_onBlur(this);}catch(e){}" />
																				</td>
																				<td align="left" nowrap>
																					<label
																						tabindex="-1"
																						for="optOutputFormat4"
																						class="radio"
																						onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
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
																						onclick="formatClick(1);"
																						onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																						onfocus="try{radio_onFocus(this);}catch(e){}"
																						onblur="try{radio_onBlur(this);}catch(e){}" />
																				</td>
																				<td align="left" nowrap>
																					<!--
                                                                            <label 
                                                                                tabindex=-1
                                                                                for="optOutputFormat1"
                                                                                class="radio"
                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                            />
	        																	CSV File
                                                       	    		        </label>
																			-->
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
																						onclick="formatClick(5);"
																						onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																						onfocus="try{radio_onFocus(this);}catch(e){}"
																						onblur="try{radio_onBlur(this);}catch(e){}" />
																				</td>
																				<td>
																					<!--
                                                                            <label 
                                                                                tabindex=-1
                                                                                for="optOutputFormat5"
                                                                                class="radio"
                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                            />
																			    Excel Chart
                                                       	    		        </label>
																			-->
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
																						onclick="formatClick(6);"
																						onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																						onfocus="try{radio_onFocus(this);}catch(e){}"
																						onblur="try{radio_onBlur(this);}catch(e){}" />
																				</td>
																				<td>
																					<!--
                                                                            <label 
                                                                                tabindex=-1
                                                                                for="optOutputFormat6"
                                                                                class="radio"
                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                            />
																			    Excel Pivot Table
                                                       	    		        </label>
																			-->
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
														<td valign="top" width="75%">
															<table class="outline" cellspacing="0" cellpadding="4" width="100%" height="100%">
																<tr height="10">
																	<td height="10" align="left" valign="top">Output Destination(s) :
																		<br>
																		<br>
																		<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																			<tr height="20">
																				<td width="5">&nbsp;</td>
																				<td align="left" colspan="6" nowrap>
																					<input name="chkPreview" id="chkPreview" type="checkbox" disabled="disabled" tabindex="-1"
																						onclick="changeTab5Control();"
																						onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																					<label
																						for="chkPreview"
																						class="checkbox checkboxdisabled"
																						tabindex="0"
																						onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																						onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																						onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																						onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
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
																					<input name="chkDestination0" id="chkDestination0" type="checkbox" disabled="disabled" tabindex="-1"
																						onclick="changeTab5Control();"
																						onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																					<label
																						for="chkDestination0"
																						class="checkbox checkboxdisabled"
																						tabindex="0"
																						onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																						onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																						onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																						onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																						Display output on screen
																					</label>
																				</td>
																				<td></td>
																			</tr>

																			<tr height="10">
																				<td colspan="8"></td>
																			</tr>

																			<tr height="20">
																				<td></td>
																				<td align="left" nowrap>
																					<input name="chkDestination1" id="chkDestination1" type="checkbox" disabled="disabled" tabindex="-1"
																						onclick="changeTab5Control();"
																						onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																					<label
																						for="chkDestination1"
																						class="checkbox checkboxdisabled"
																						tabindex="0"
																						onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																						onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																						onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																						onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																						Send to printer 
																					</label>
																				</td>
																				<td width="30" nowrap>&nbsp;</td>
																				<td align="left" nowrap>Printer location : 
																				</td>
																				<td width="15">&nbsp;</td>
																				<td colspan="2">
																					<select id="cboPrinterName" name="cboPrinterName" class="combo" width="100%" style="WIDTH: 400px"
																						onchange="changeTab5Control()">
																					</select>
																				</td>
																				<td></td>
																			</tr>

																			<tr height="10">
																				<td colspan="8"></td>
																			</tr>

																			<tr height="20">
																				<td></td>
																				<td align="left" nowrap>
																					<input name="chkDestination2" id="chkDestination2" type="checkbox" disabled="disabled" tabindex="-1"
																						onclick="changeTab5Control();"
																						onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																					<label
																						for="chkDestination2"
																						class="checkbox checkboxdisabled"
																						tabindex="0"
																						onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																						onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																						onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																						onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																						Save to file 
																					</label>
																				</td>
																				<td></td>
																				<td align="left" nowrap>File name : 
																				</td>
																				<td></td>
																				<td colspan="2">
																					<table class="invisible" cellspacing="0" cellpadding="0" style="WIDTH: 400px">
																						<tr>
																							<td>
																								<input id="txtFilename" width="100%" name="txtFilename" class="text textdisabled" disabled="disabled" tabindex="-1" style="WIDTH: 375px">
																							</td>
																							<td width="25">
																								<input id="cmdFilename" width="100%" name="cmdFilename" class="btn btndisabled" style="WIDTH: 100%" type="button" value='...' disabled="disabled"
																									onclick="saveFile(); changeTab5Control();"
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

																			<tr height="10">
																				<td colspan="8"></td>
																			</tr>

																			<tr height="20">
																				<td colspan="3"></td>
																				<td align="left" nowrap>If existing file :</td>
																				<td></td>
																				<td colspan="2" width="100%" nowrap>
																					<select id="cboSaveExisting" name="cboSaveExisting" class="combo" style="WIDTH: 400px"
																						onchange="changeTab5Control()">
																					</select>
																				</td>
																				<td></td>
																			</tr>

																			<tr height="10">
																				<td colspan="8"></td>
																			</tr>

																			<tr height="20">
																				<td></td>
																				<td align="left" nowrap>
																					<input name="chkDestination3" id="chkDestination3" type="checkbox" disabled="disabled" tabindex="-1"
																						onclick="changeTab5Control();"
																						onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />

																					<label for="chkDestination3"
																						class="checkbox checkboxdisabled"
																						tabindex="0"
																						onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																						onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																						onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																						onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																						Send as email
																					</label>
																				</td>
																				<td></td>
																				<td align="left" nowrap>Email group :   </td>
																				<td></td>
																				<td colspan="2">
																					<table class="invisible" cellspacing="0" cellpadding="0" style="WIDTH: 400px">
																						<tr>
																							<td>
																								<input id="txtEmailGroup" name="txtEmailGroup" class="text textdisabled" disabled="disabled" tabindex="-1" style="WIDTH: 100%">
																								<input id="txtEmailGroupID" name="txtEmailGroupID" type="hidden" class="text textdisabled" disabled="disabled" tabindex="-1">
																							</td>
																							<td width="25">
																								<input id="cmdEmailGroup" name="cmdEmailGroup" style="WIDTH: 100%" type="button" value='...' disabled="disabled" class="btn btndisabled"
																									onclick="selectEmailGroup(); changeTab5Control();"
																									onmouseover="try{button_onMouseOver(this);}catch(e){}"
																									onmouseout="try{button_onMouseOut(this);}catch(e){}"
																									onfocus="try{button_onFocus(this);}catch(e){}"
																									onblur="try{button_onBlur(this);}catch(e){}" />
																							</td>
																					</table>
																				</td>
																		</table>
																	</td>
																	<td></td>
																</tr>

																<tr height="10">
																	<td colspan="8"></td>
																</tr>

																<tr height="20">
																	<td colspan="3"></td>
																	<td align="left" nowrap>Email subject :</td>
																	<td></td>
																	<td colspan="2" width="100%" nowrap>
																		<input id="txtEmailSubject" class="text textdisabled" disabled="disabled" maxlength="255" name="txtEmailSubject" style="WIDTH: 400px"
																			onchange="frmUseful.txtChanged.value = 1;"
																			onkeydown="frmUseful.txtChanged.value = 1;">
																	</td>
																	<td></td>
																</tr>

																<tr height="10">
																	<td colspan="8"></td>
																</tr>

																<tr height="20">
																	<td colspan="3"></td>
																	<td align="left" nowrap>Attach as : 
																	</td>
																	<td></td>
																	<td colspan="2" width="100%" nowrap>
																		<input id="txtEmailAttachAs" maxlength="255" class="text textdisabled" disabled="disabled" name="txtEmailAttachAs" style="WIDTH: 400px"
																			onchange="frmUseful.txtChanged.value = 1;"
																			onkeydown="frmUseful.txtChanged.value = 1;">
																	</td>
																	<td></td>
																</tr>

																<tr height="10">
																	<td colspan="8"></td>
																</tr>
															</table>
														</td>
													</tr>
												</table>
											</td>
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
				<td width="10"></td>
				<tr height="10">
					<td colspan="3"></td>
				</tr>

				<tr height="10">
					<td width="10"></td>
					<td>
						<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
							<tr>
								<td>&nbsp;</td>
								<td width="80">
									<input type="button" id="cmdOK" name="cmdOK" value="OK" style="WIDTH: 100%" class="btn"
										onclick="okClick()"
										onmouseover="try{button_onMouseOver(this);}catch(e){}"
										onmouseout="try{button_onMouseOut(this);}catch(e){}"
										onfocus="try{button_onFocus(this);}catch(e){}"
										onblur="try{button_onBlur(this);}catch(e){}" />
								</td>
								<td width="10"></td>
								<td width="80">
									<input type="button" id="cmdCancel" name="cmdCancel" value="Cancel" style="WIDTH: 100%" class="btn"
										onclick="cancelClick()"
										onmouseover="try{button_onMouseOver(this);}catch(e){}"
										onmouseout="try{button_onMouseOut(this);}catch(e){}"
										onfocus="try{button_onFocus(this);}catch(e){}"
										onblur="try{button_onBlur(this);}catch(e){}" />
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
</div>

<form id="frmTables" style="visibility: hidden; display: none">
	<%
		Dim sErrorDescription = ""
	
		' Get the table records.
		Dim cmdTables = CreateObject("ADODB.Command")
		cmdTables.CommandText = "sp_ASRIntGetTablesInfo"
		cmdTables.CommandType = 4	' Stored Procedure
		cmdTables.ActiveConnection = Session("databaseConnection")
	
		Response.Write("<B>Set Connection</B>")
	
		Err.Number = 0
		Dim rstTablesInfo = cmdTables.Execute
	
		Response.Write("<B>Executed SP</B>")
	
		If (Err.Number <> 0) Then
			sErrorDescription = "The tables information could not be retrieved." & vbCrLf & FormatError(Err.Description)
		End If

		If Len(sErrorDescription) = 0 Then
			Dim iCount = 0
			Do While Not rstTablesInfo.EOF
				Response.Write("<INPUT type='hidden' id=txtTableName_" & rstTablesInfo.fields("tableID").value & " name=txtTableName_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("tableName").value & """>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtTableType_" & rstTablesInfo.fields("tableID").value & " name=txtTableType_" & rstTablesInfo.fields("tableID").value & " value=" & rstTablesInfo.fields("tableType").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtTableChildren_" & rstTablesInfo.fields("tableID").value & " name=txtTableChildren_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("childrenString").value & """>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtTableChildrenNames_" & rstTablesInfo.fields("tableID").value & " name=txtTableChildrenNames_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("childrenNames").value & """>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtTableParents_" & rstTablesInfo.fields("tableID").value & " name=txtTableParents_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("parentsString").value & """>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtTableRelations_" & rstTablesInfo.fields("tableID").value & " name=txtTableRelations_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("relatedString").value & """>" & vbCrLf)
				rstTablesInfo.MoveNext()
			Loop

			' Release the ADO recordset object.
			rstTablesInfo.close()
			rstTablesInfo = Nothing
		End If
	
		' Release the ADO command object.
		cmdTables = Nothing
	%>
</form>

<%--<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
</form>--%>
	<form action="default_Submit" method=post id=Form1 name=frmGoto style="visibility:hidden;display:none">
		<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
	</form>


<form id="frmOriginalDefinition" name="frmOriginalDefinition" style="visibility: hidden; display: none">
	<%
		Dim sErrMsg = ""
		Dim prmUtilID
	
		If LCase(Session("action")) <> "new" Then
			Dim cmdDefn = CreateObject("ADODB.Command")
			cmdDefn.CommandText = "spASRIntGetCalendarReportDefinition"
			cmdDefn.CommandType = 4	' Stored Procedure
			cmdDefn.ActiveConnection = Session("databaseConnection")
		
			prmUtilID = cmdDefn.CreateParameter("utilID", 3, 1)	' 3=integer, 1=input
			cmdDefn.Parameters.Append(prmUtilID)
			prmUtilID.value = CleanNumeric(Session("utilid"))

			Dim prmUser = cmdDefn.CreateParameter("user", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
			cmdDefn.Parameters.Append(prmUser)
			prmUser.value = Session("username")

			Dim prmAction = cmdDefn.CreateParameter("action", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
			cmdDefn.Parameters.Append(prmAction)
			prmAction.value = Session("action")

			Dim prmErrMsg = cmdDefn.CreateParameter("errMsg", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
			cmdDefn.Parameters.Append(prmErrMsg)

			Dim prmName = cmdDefn.CreateParameter("name", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
			cmdDefn.Parameters.Append(prmName)

			Dim prmOwner = cmdDefn.CreateParameter("owner", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
			cmdDefn.Parameters.Append(prmOwner)

			Dim prmDescription = cmdDefn.CreateParameter("description", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
			cmdDefn.Parameters.Append(prmDescription)

			Dim prmBaseTableID = cmdDefn.CreateParameter("baseTableID", 3, 2)	'3=integer, 2=output
			cmdDefn.Parameters.Append(prmBaseTableID)

			Dim prmAllRecords = cmdDefn.CreateParameter("allRecords", 11, 2) '11=bit, 2=output
			cmdDefn.Parameters.Append(prmAllRecords)

			Dim prmPicklistID = cmdDefn.CreateParameter("picklistID", 3, 2)	'3=integer, 2=output
			cmdDefn.Parameters.Append(prmPicklistID)

			Dim prmPicklistName = cmdDefn.CreateParameter("picklistName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
			cmdDefn.Parameters.Append(prmPicklistName)

			Dim prmPicklistHidden = cmdDefn.CreateParameter("picklistHidden", 11, 2) '11=bit, 2=output
			cmdDefn.Parameters.Append(prmPicklistHidden)

			Dim prmFilterID = cmdDefn.CreateParameter("filterID", 3, 2)	'3=integer, 2=output
			cmdDefn.Parameters.Append(prmFilterID)

			Dim prmFilterName = cmdDefn.CreateParameter("filterName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
			cmdDefn.Parameters.Append(prmFilterName)

			Dim prmFilterHidden = cmdDefn.CreateParameter("filterHidden", 11, 2) '11=bit, 2=output
			cmdDefn.Parameters.Append(prmFilterHidden)
		
			Dim prmPrintFilterHeader = cmdDefn.CreateParameter("printFilterHeader", 11, 2) '11=bit, 2=output
			cmdDefn.Parameters.Append(prmPrintFilterHeader)
		
			Dim prmDesc1ID = cmdDefn.CreateParameter("desc1ID", 3, 2)	'3=integer, 2=output
			cmdDefn.Parameters.Append(prmDesc1ID)

			Dim prmDesc2ID = cmdDefn.CreateParameter("desc2ID", 3, 2)	'3=integer, 2=output
			cmdDefn.Parameters.Append(prmDesc2ID)
		
			Dim prmDescExprID = cmdDefn.CreateParameter("descExprID", 3, 2)	'3=integer, 2=output
			cmdDefn.Parameters.Append(prmDescExprID)
		
			Dim prmDescExprName = cmdDefn.CreateParameter("descExprName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
			cmdDefn.Parameters.Append(prmDescExprName)
		
			Dim prmDescCalcHidden = cmdDefn.CreateParameter("descCalcHidden", 11, 2) '11=bit, 2=output
			cmdDefn.Parameters.Append(prmDescCalcHidden)
		
			Dim prmRegionID = cmdDefn.CreateParameter("regionID", 3, 2)	'3=integer, 2=output
			cmdDefn.Parameters.Append(prmRegionID)

			Dim prmGroupByDesc = cmdDefn.CreateParameter("groupByDesc", 11, 2) '11=bit, 2=output
			cmdDefn.Parameters.Append(prmGroupByDesc)
		
			Dim prmDescSeparator = cmdDefn.CreateParameter("descSeparator", 200, 2, 8000)	'11=bit, 2=output
			cmdDefn.Parameters.Append(prmDescSeparator)
		
			Dim prmStartType = cmdDefn.CreateParameter("startType", 3, 2)	'3=integer, 2=output
			cmdDefn.Parameters.Append(prmStartType)

			Dim prmFixedStart = cmdDefn.CreateParameter("fixedStart", 135, 2)	'135=datetime, 2=output
			cmdDefn.Parameters.Append(prmFixedStart)
		
			Dim prmStartFrequency = cmdDefn.CreateParameter("startFrequency", 3, 2)	'3=integer, 2=output
			cmdDefn.Parameters.Append(prmStartFrequency)
		
			Dim prmStartPeriod = cmdDefn.CreateParameter("startPeriod", 3, 2)	'3=integer, 2=output
			cmdDefn.Parameters.Append(prmStartPeriod)
		
			Dim prmCustomStartID = cmdDefn.CreateParameter("customStartID", 3, 2)	'3=integer, 2=output
			cmdDefn.Parameters.Append(prmCustomStartID)
		
			Dim prmCustomStartName = cmdDefn.CreateParameter("customStartName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
			cmdDefn.Parameters.Append(prmCustomStartName)
		
			Dim prmStartDateCalcHidden = cmdDefn.CreateParameter("startDateCalcHidden", 11, 2) '11=bit, 2=output
			cmdDefn.Parameters.Append(prmStartDateCalcHidden)
		
			Dim prmEndType = cmdDefn.CreateParameter("endType", 3, 2)	'3=integer, 2=output
			cmdDefn.Parameters.Append(prmEndType)

			Dim prmFixedEnd = cmdDefn.CreateParameter("fixedEnd", 135, 2)	'135=datetime, 2=output
			cmdDefn.Parameters.Append(prmFixedEnd)
		
			Dim prmEndFrequency = cmdDefn.CreateParameter("endFrequency", 3, 2)	'3=integer, 2=output
			cmdDefn.Parameters.Append(prmEndFrequency)
		
			Dim prmEndPeriod = cmdDefn.CreateParameter("endPeriod", 3, 2)	'3=integer, 2=output
			cmdDefn.Parameters.Append(prmEndPeriod)
		
			Dim prmCustomEndID = cmdDefn.CreateParameter("customEndID", 3, 2)	'3=integer, 2=output
			cmdDefn.Parameters.Append(prmCustomEndID)

			Dim prmCustomEndName = cmdDefn.CreateParameter("customEndName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
			cmdDefn.Parameters.Append(prmCustomEndName)
		
			Dim prmEndDateCalcHidden = cmdDefn.CreateParameter("endDateCalcHidden", 11, 2) '11=bit, 2=output
			cmdDefn.Parameters.Append(prmEndDateCalcHidden)
		
			Dim prmShadeBHols = cmdDefn.CreateParameter("shadeBHols", 11, 2) '11=bit, 2=output
			cmdDefn.Parameters.Append(prmShadeBHols)
		
			Dim prmShowCaptions = cmdDefn.CreateParameter("showCaptions", 11, 2) '11=bit, 2=output
			cmdDefn.Parameters.Append(prmShowCaptions)
		
			Dim prmShadeWeekends = cmdDefn.CreateParameter("shadeWeekends", 11, 2) '11=bit, 2=output
			cmdDefn.Parameters.Append(prmShadeWeekends)
		
			Dim prmStartOnCurrentMonth = cmdDefn.CreateParameter("startOnCurrentMonth", 11, 2) '11=bit, 2=output
			cmdDefn.Parameters.Append(prmStartOnCurrentMonth)
		
			Dim prmIncludeWorkingDaysOnly = cmdDefn.CreateParameter("includeWorkingDaysOnly", 11, 2) '11=bit, 2=output
			cmdDefn.Parameters.Append(prmIncludeWorkingDaysOnly)
		
			Dim prmIncludeBHols = cmdDefn.CreateParameter("includeBHols", 11, 2) '11=bit, 2=output
			cmdDefn.Parameters.Append(prmIncludeBHols)
			'-----------------------------------------
			Dim prmOutputPreview = cmdDefn.CreateParameter("outputPreview", 11, 2) '11=bit, 2=output
			cmdDefn.Parameters.Append(prmOutputPreview)
		
			Dim prmOutputFormat = cmdDefn.CreateParameter("outputFormat", 3, 2)	'3=integer, 2=output
			cmdDefn.Parameters.Append(prmOutputFormat)
		
			Dim prmOutputScreen = cmdDefn.CreateParameter("outputScreen", 11, 2) '11=bit, 2=output
			cmdDefn.Parameters.Append(prmOutputScreen)
		
			Dim prmOutputPrinter = cmdDefn.CreateParameter("outputPrinter", 11, 2) '11=bit, 2=output
			cmdDefn.Parameters.Append(prmOutputPrinter)
		
			Dim prmOutputPrinterName = cmdDefn.CreateParameter("outputPrinterName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
			cmdDefn.Parameters.Append(prmOutputPrinterName)
		
			Dim prmOutputSave = cmdDefn.CreateParameter("outputSave", 11, 2) '11=bit, 2=output
			cmdDefn.Parameters.Append(prmOutputSave)
		
			Dim prmOutputSaveExisting = cmdDefn.CreateParameter("outputSaveExisting", 3, 2)	'3=integer, 2=output
			cmdDefn.Parameters.Append(prmOutputSaveExisting)
		
			Dim prmOutputEmail = cmdDefn.CreateParameter("outputEmail", 11, 2) '11=bit, 2=output
			cmdDefn.Parameters.Append(prmOutputEmail)
		
			Dim prmOutputEmailAddr = cmdDefn.CreateParameter("outputEmailAddr", 3, 2)	'3=integer, 2=output
			cmdDefn.Parameters.Append(prmOutputEmailAddr)
		
			Dim prmOutputEmailAddrName = cmdDefn.CreateParameter("outputEmailAddrName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
			cmdDefn.Parameters.Append(prmOutputEmailAddrName)
		
			Dim prmOutputEmailSubject = cmdDefn.CreateParameter("outputEmailSubject", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
			cmdDefn.Parameters.Append(prmOutputEmailSubject)

			Dim prmOutputEmailAttachAs = cmdDefn.CreateParameter("outputEmailAttachAs", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
			cmdDefn.Parameters.Append(prmOutputEmailAttachAs)

			Dim prmOutputFilename = cmdDefn.CreateParameter("outputFilename", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
			cmdDefn.Parameters.Append(prmOutputFilename)
			'-----------------------------------------
	
			Dim prmTimestamp = cmdDefn.CreateParameter("timestamp", 3, 2)	' 3=integer, 2=output
			cmdDefn.Parameters.Append(prmTimestamp)

			Err.Number = 0
			Dim rstDefinition = cmdDefn.Execute
		
			Dim iHiddenEventFilterCount = 0
			Dim iCount = 0
			If (Err.Number <> 0) Then
				sErrMsg = CType(("'" & Session("utilname") & "' report definition could not be read." & vbCrLf) & FormatError(Err.Description), String)
			Else
				If rstDefinition.state <> 0 Then
					' Read recordset values.
					iCount = 0
					Do While Not rstDefinition.EOF
						iCount = iCount + 1
					
						Response.Write("<INPUT type='hidden' id=txtReportDefnEvent_" & iCount & " name=txtReportDefnEvent_" & iCount & " value=""" & Replace(rstDefinition.fields("definitionString").value, """", "&quot;") & """>" & vbCrLf)
					
						If rstDefinition.fields("FilterHidden").value = "Y" Then
							iHiddenEventFilterCount = iHiddenEventFilterCount + 1
						End If
					
						rstDefinition.MoveNext()
					Loop

					' Release the ADO recordset object.
					rstDefinition.close()
				End If
				rstDefinition = Nothing
			
				Session("hiddenfiltercount") = iHiddenEventFilterCount
				Session("CalendarEventCount") = iCount
				' NB. IMPORTANT ADO NOTE.
				' When calling a stored procedure which returns a recordset AND has output parameters
				' you need to close the recordset and set it to nothing before using the output parameters. 
				If Len(cmdDefn.Parameters("errMsg").value) > 0 Then
					sErrMsg = CType(("'" & Session("utilname") & "' " & cmdDefn.Parameters("errMsg").value), String)
				End If

				Response.Write("<INPUT type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Replace(cmdDefn.Parameters("name").value, """", "&quot;") & """>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_Owner name=txtDefn_Owner value=""" & Replace(cmdDefn.Parameters("owner").value, """", "&quot;") & """>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_Description name=txtDefn_Description value=""" & Replace(cmdDefn.Parameters("description").value, """", "&quot;") & """>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_BaseTableID name=txtDefn_BaseTableID value=" & cmdDefn.Parameters("baseTableID").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_AllRecords name=txtDefn_AllRecords value=" & cmdDefn.Parameters("allRecords").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_PicklistID name=txtDefn_PicklistID value=" & cmdDefn.Parameters("picklistID").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_PicklistName name=txtDefn_PicklistName value=""" & Replace(cmdDefn.Parameters("picklistName").value, """", "&quot;") & """>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_PicklistHidden name=txtDefn_PicklistHidden value=" & cmdDefn.Parameters("picklistHidden").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_FilterID name=txtDefn_FilterID value=" & cmdDefn.Parameters("filterID").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_FilterName name=txtDefn_FilterName value=""" & Replace(cmdDefn.Parameters("filterName").value, """", "&quot;") & """>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_FilterHidden name=txtDefn_FilterHidden value=" & cmdDefn.Parameters("filterHidden").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_PrintFilterHeader name=txtDefn_PrintFilterHeader value=" & cmdDefn.Parameters("PrintFilterHeader").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_Desc1ID name=txtDefn_Desc1ID value=" & cmdDefn.Parameters("Desc1ID").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_Desc2ID name=txtDefn_Desc2ID value=" & cmdDefn.Parameters("Desc2ID").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_DescExprID name=txtDefn_DescExprID value=" & cmdDefn.Parameters("descExprID").value & ">" & vbCrLf)
				If IsDBNull(cmdDefn.Parameters("descExprName").value) Then
					Response.Write("<INPUT type='hidden' id=txtDefn_DescExprName name=txtDefn_DescExprName value="""">" & vbCrLf)
				Else
					Response.Write("<INPUT type='hidden' id=txtDefn_DescExprName name=txtDefn_DescExprName value=""" & Replace(cmdDefn.Parameters("descExprName").value, """", "&quot;") & """>" & vbCrLf)
				End If
				Response.Write("<INPUT type='hidden' id=txtDefn_DescExprHidden name=txtDefn_DescExprHidden value=" & cmdDefn.Parameters("descCalcHidden").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_RegionID name=txtDefn_RegionID value=" & cmdDefn.Parameters("RegionID").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_GroupByDesc name=txtDefn_GroupByDesc value=" & cmdDefn.Parameters("GroupByDesc").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_DescSeparator name=txtDefn_DescSeparator value=""" & cmdDefn.Parameters("DescSeparator").value & """>" & vbCrLf)

				Response.Write("<INPUT type='hidden' id=txtDefn_StartType name=txtDefn_StartType value=" & cmdDefn.Parameters("StartType").value & ">" & vbCrLf)
			
				If IsDBNull(cmdDefn.Parameters("FixedStart").value) Then
					Response.Write("<INPUT type='hidden' id=txtDefn_FixedStart name=txtDefn_FixedStart value="""">" & vbCrLf)
				Else
					Response.Write("<INPUT type='hidden' id=txtDefn_FixedStart name=txtDefn_FixedStart value=" & ConvertSqlDateToLocale(cmdDefn.Parameters("FixedStart").value) & ">" & vbCrLf)
				End If
				Response.Write("<INPUT type='hidden' id=txtDefn_StartFrequency name=txtDefn_StartFrequency value=" & cmdDefn.Parameters("StartFrequency").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_StartPeriod name=txtDefn_StartPeriod value=" & cmdDefn.Parameters("StartPeriod").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_CustomStartID name=txtDefn_CustomStartID value=" & cmdDefn.Parameters("customStartID").value & ">" & vbCrLf)
				If IsDBNull(cmdDefn.Parameters("customStartName").value) Then
					Response.Write("<INPUT type='hidden' id=txtDefn_CustomStartName name=txtDefn_CustomStartName value="""">" & vbCrLf)
				Else
					Response.Write("<INPUT type='hidden' id=txtDefn_CustomStartName name=txtDefn_CustomStartName value=""" & Replace(cmdDefn.Parameters("customStartName").value, """", "&quot;") & """>" & vbCrLf)
				End If
				Response.Write("<INPUT type='hidden' id=txtDefn_CustomStartCalcHidden name=txtDefn_CustomStartCalcHidden value=" & cmdDefn.Parameters("startDateCalcHidden").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_EndType name=txtDefn_EndType value=" & cmdDefn.Parameters("EndType").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_FixedEnd name=txtDefn_FixedEnd value=" & ConvertSqlDateToLocale(cmdDefn.Parameters("FixedEnd").value) & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_EndFrequency name=txtDefn_EndFrequency value=" & cmdDefn.Parameters("EndFrequency").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_EndPeriod name=txtDefn_EndPeriod value=" & cmdDefn.Parameters("EndPeriod").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_CustomEndID name=txtDefn_CustomEndID value=" & cmdDefn.Parameters("customEndID").value & ">" & vbCrLf)
				If IsDBNull(cmdDefn.Parameters("customEndName").value) Then
					Response.Write("<INPUT type='hidden' id=txtDefn_CustomEndName name=txtDefn_CustomEndName value="""">" & vbCrLf)
				Else
					Response.Write("<INPUT type='hidden' id=txtDefn_CustomEndName name=txtDefn_CustomEndName value=""" & Replace(cmdDefn.Parameters("customEndName").value, """", "&quot;") & """>" & vbCrLf)
				End If
				Response.Write("<INPUT type='hidden' id=txtDefn_CustomEndCalcHidden name=txtDefn_CustomEndCalcHidden value=" & cmdDefn.Parameters("endDateCalcHidden").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_ShadeBHols name=txtDefn_ShadeBHols value=" & cmdDefn.Parameters("ShadeBHols").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_ShowCaptions name=txtDefn_ShowCaptions value=" & cmdDefn.Parameters("ShowCaptions").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_ShadeWeekends name=txtDefn_ShadeWeekends value=" & cmdDefn.Parameters("ShadeWeekends").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_StartOnCurrentMonth name=txtDefn_StartOnCurrentMonth value=" & cmdDefn.Parameters("StartOnCurrentMonth").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_IncludeWorkingDaysOnly name=txtDefn_IncludeWorkingDaysOnly value=" & cmdDefn.Parameters("IncludeWorkingDaysOnly").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_IncludeBHols name=txtDefn_IncludeBHols value=" & cmdDefn.Parameters("IncludeBHols").value & ">" & vbCrLf)
			
				Response.Write("<INPUT type='hidden' id=txtDefn_OutputPreview name=txtDefn_OutputPreview value=" & cmdDefn.Parameters("OutputPreview").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_OutputFormat name=txtDefn_OutputFormat value=" & cmdDefn.Parameters("OutputFormat").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_OutputScreen name=txtDefn_OutputScreen value=" & cmdDefn.Parameters("OutputScreen").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_OutputPrinter name=txtDefn_OutputPrinter value=" & cmdDefn.Parameters("OutputPrinter").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_OutputPrinterName name=txtDefn_OutputPrinterName value=""" & cmdDefn.Parameters("OutputPrinterName").value & """>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_OutputSave name=txtDefn_OutputSave value=" & cmdDefn.Parameters("OutputSave").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_OutputSaveExisting name=txtDefn_OutputSaveExisting value=" & cmdDefn.Parameters("OutputSaveExisting").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmail name=txtDefn_OutputEmail value=" & cmdDefn.Parameters("OutputEmail").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmailAddr name=txtDefn_OutputEmailAddr value=" & cmdDefn.Parameters("OutputEmailAddr").value & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmailAddrName name=txtDefn_OutputEmailName value=""" & Replace(cmdDefn.Parameters("OutputEmailAddrName").value, """", "&quot;") & """>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmailSubject name=txtDefn_OutputEmailSubject value=""" & Replace(cmdDefn.Parameters("OutputEmailSubject").value, """", "&quot;") & """>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmailAttachAs name=txtDefn_OutputEmailAttachAs value=""" & Replace(cmdDefn.Parameters("OutputEmailAttachAs").value, """", "&quot;") & """>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtDefn_OutputFilename name=txtDefn_OutputFilename value=""" & cmdDefn.Parameters("OutputFilename").value & """>" & vbCrLf)
			
				Response.Write("<INPUT type='hidden' id=txtDefn_Timestamp name=txtDefn_Timestamp value=" & cmdDefn.Parameters("timestamp").value & ">" & vbCrLf)

				'********************************************************************************

				Dim cmdReportOrder = CreateObject("ADODB.Command")
				cmdReportOrder.CommandText = "spASRIntGetCalendarReportOrder"
				cmdReportOrder.CommandType = 4	'Stored Procedure
				cmdReportOrder.ActiveConnection = Session("databaseConnection")
		
				Dim prmUtilID2 = cmdReportOrder.CreateParameter("utilID2", 3, 1) ' 3=integer, 1=input
				cmdReportOrder.Parameters.Append(prmUtilID2)
				prmUtilID2.value = CleanNumeric(Session("utilid"))
		
				Dim prmErrMsg2 = cmdReportOrder.CreateParameter("errMsg2", 200, 2, 8000) '200=varchar, 2=output, 8000=size
				cmdReportOrder.Parameters.Append(prmErrMsg2)

				Err.Clear()
				Dim rstOrder = cmdReportOrder.Execute
		
				iCount = 0
				If (Err.Number <> 0) Then
					sErrMsg = "'" & Session("utilname") & "' report order definition could not be read." & vbCrLf & FormatError(Err.Description)
				Else
					If rstOrder.state <> 0 Then
						' Read recordset values.
			
						Do While Not rstOrder.EOF
							iCount = iCount + 1
							Response.Write("<INPUT type='hidden' id=txtReportDefnOrder_" & iCount & " name=txtReportDefnOrder_" & iCount & " value=""" & rstOrder.fields("orderString").value & """>" & vbCrLf)

							rstOrder.MoveNext()
						Loop
						' Release the ADO recordset object.
						rstOrder.close()
					End If
					rstOrder = Nothing
				End If

				Session("CalendarOrderCount") = iCount

				' NB. IMPORTANT ADO NOTE.
				' When calling a stored procedure which returns a recordset AND has output parameters
				' you need to close the recordset and set it to nothing before using the output parameters. 
				If Len(cmdReportOrder.Parameters("errMsg2").value) > 0 Then
					sErrMsg = "'" & Session("utilname") & "' " & cmdDefn.Parameters("errMsg2").value
				End If

				cmdReportOrder = Nothing

				'********************************************************************************

			End If

			' Release the ADO command object.
			cmdDefn = Nothing

			If Len(sErrMsg) > 0 Then
				Session("confirmtext") = sErrMsg
				Session("confirmtitle") = "OpenHR Intranet"
	                                    	    
				Session("followpage") = "defsel"
	                                    	    
				Session("reaction") = "CALENDARREPORTS"
	                                    	    
				Response.Clear()
	                                    	    
				Response.Redirect("confirmok")
			End If
	
		Else
			Session("CalendarEventCount") = 0
			Session("CalendarOrderCount") = 0
			Session("hiddenfiltercount") = 0
		End If
	
	%>
</form>

<form id="frmAccess">
	<%
		sErrorDescription = ""
	
		' Get the table records.
		Dim cmdAccess = CreateObject("ADODB.Command")
		cmdAccess.CommandText = "spASRIntGetUtilityAccessRecords"
		cmdAccess.CommandType = 4	' Stored Procedure
		cmdAccess.ActiveConnection = Session("databaseConnection")

		Dim prmUtilType = cmdAccess.CreateParameter("utilType", 3, 1)	' 3=integer, 1=input
		cmdAccess.Parameters.Append(prmUtilType)
		prmUtilType.value = 17 ' 17 = calendar report

		prmUtilID = cmdAccess.CreateParameter("utilID", 3, 1)	' 3=integer, 1=input
		cmdAccess.Parameters.Append(prmUtilID)
		If UCase(Session("action")) = "NEW" Then
			prmUtilID.value = 0
		Else
			prmUtilID.value = CleanNumeric(Session("utilid"))
		End If

		Dim prmFromCopy = cmdAccess.CreateParameter("fromCopy", 3, 1)	' 3=integer, 1=input
		cmdAccess.Parameters.Append(prmFromCopy)
		If UCase(Session("action")) = "COPY" Then
			prmFromCopy.value = 1
		Else
			prmFromCopy.value = 0
		End If

		Err.Number = 0
		Dim rstAccessInfo = cmdAccess.Execute
		If (Err.Number <> 0) Then
			sErrorDescription = "The access information could not be retrieved." & vbCrLf & FormatError(Err.Description)
		End If

		If Len(sErrorDescription) = 0 Then
			Dim iCount = 0
			Do While Not rstAccessInfo.EOF
				Response.Write("<INPUT type='hidden' id=txtAccess_" & iCount & " name=txtAccess_" & iCount & " value=""" & rstAccessInfo.fields("accessDefinition").value & """>" & vbCrLf)

				iCount = iCount + 1
				rstAccessInfo.MoveNext()
			Loop

			' Release the ADO recordset object.
			rstAccessInfo.close()
			rstAccessInfo = Nothing
		End If
	
		' Release the ADO command object.
		cmdAccess = Nothing
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
		Dim cmdDefinition = CreateObject("ADODB.Command")
		cmdDefinition.CommandText = "sp_ASRIntGetModuleParameter"
		cmdDefinition.CommandType = 4	' Stored procedure.
		cmdDefinition.ActiveConnection = Session("databaseConnection")

		Dim prmModuleKey = cmdDefinition.CreateParameter("moduleKey", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
		cmdDefinition.Parameters.Append(prmModuleKey)
		prmModuleKey.value = "MODULE_PERSONNEL"

		Dim prmParameterKey = cmdDefinition.CreateParameter("paramKey", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
		cmdDefinition.Parameters.Append(prmParameterKey)
		prmParameterKey.value = "Param_TablePersonnel"

		Dim prmParameterValue = cmdDefinition.CreateParameter("paramValue", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
		cmdDefinition.Parameters.Append(prmParameterValue)

		Err.Number = 0
		cmdDefinition.Execute()

		Response.Write("<INPUT type='hidden' id=txtPersonnelTableID name=txtPersonnelTableID value=" & cmdDefinition.Parameters("paramValue").value & ">" & vbCrLf)
	
		cmdDefinition = Nothing

		Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtAction name=txtAction value=" & Session("action") & ">" & vbCrLf)
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


<div>
	<input type='hidden' id="txtTicker" name="txtTicker" value="0">
	<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">
</div>

<script type="text/javascript">
	util_def_calendarreport_window_onload();
	util_def_calendarreport_addhandlers();
</script>


