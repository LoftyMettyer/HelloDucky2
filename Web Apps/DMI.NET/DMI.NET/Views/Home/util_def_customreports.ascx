<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="ADODB" %>

<script src="<%: Url.Content("~/bundles/utilities_customreports")%>" type="text/javascript"></script>

<%Html.RenderPartial("Util_Def_CustomReports/dialog")%>

<div <%=session("BodyTag")%>>
	<form id="frmDefinition" name="frmDefinition">

		<table align="center" class="outline" cellpadding="5" cellspacing="0" width="100%" height="100%">
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
									onclick="displayPage(1)" />
								<input type="button" value="Related Tables" id="btnTab2" name="btnTab2" class="btn"
									onclick="displayPage(2)" />
								<input type="button" value="Columns" id="btnTab3" name="btnTab3" class="btn"
									onclick="displayPage(3)" />
								<input type="button" value="Sort Order" id="btnTab4" name="btnTab4" class="btn"
									onclick="displayPage(4)" />
								<input type="button" value="Output" id="btnTab5" name="btnTab5" class="btn"
									onclick="displayPage(5)" />
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
									<table width="100%" height="100%" class="outline" cellspacing="0" cellpadding="5">
										<tr valign="top">
											<td>
												<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
													<tr>
														<td colspan="10" height="5"></td>
													</tr>

													<tr height="10">
														<td width="5">&nbsp;</td>
														<td width="10">Name :</td>
														<td width="5">&nbsp;</td>
														<td colspan="2">
															<input id="txtName" name="txtName" class="text" maxlength="50" style="WIDTH: 100%" onkeyup="changeTab1Control()">
														</td>
														<td width="20">&nbsp;</td>
														<td width="10">Owner :</td>
														<td width="5">&nbsp;</td>
														<td width="40%">
															<input id="txtOwner" name="txtOwner" class="text textdisabled" style="WIDTH: 100%" disabled="disabled">
														</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr>
														<td colspan="10" height="5"></td>
													</tr>

													<tr height="60">
														<td width="5">&nbsp;</td>
														<td width="10" nowrap valign="top">Description :</td>
														<td width="5">&nbsp;</td>
														<td width="40%" rowspan="3" colspan="2">
															<textarea id="txtDescription" name="txtDescription" class="textarea" style="HEIGHT: 99%; WIDTH: 100%" wrap="VIRTUAL" height="0" maxlength="255"
																onkeyup="changeTab1Control()"
																onpaste="var selectedLength = document.selection.createRange().text.length;var pasteData = window.clipboardData.getData('Text');if ((this.value.length + pasteData.length - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}"
																onkeypress="var selectedLength = document.selection.createRange().text.length;if ((this.value.length + 1 - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}">
													</textarea>
														</td>
														<td width="20" nowrap>&nbsp;</td>
														<td width="10" valign="top">Access :</td>
														<td width="5">&nbsp;</td>
														<td width="40%" rowspan="3" valign="top">
															<%Html.RenderPartial("Util_Def_CustomReports/grdaccess")%>         
														</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr height="10">
														<td colspan="8">&nbsp;</td>
													</tr>

													<tr height="10">
														<td colspan="8">&nbsp;</td>
													</tr>

													<tr height="40">
														<td width="5">&nbsp;</td>
														<td colspan="8">
															<hr>
														</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr height="10">
														<td width="5">&nbsp;</td>
														<td width="85" nowrap valign="top">Base Table :</td>
														<td width="5">&nbsp;</td>
														<td valign="top" colspan="2">
															<select id="cboBaseTable" name="cboBaseTable" style="WIDTH: 100%" class="combo combodisabled"
																onchange="changeBaseTable()" disabled="disabled">
															</select>
														</td>
														<td width="20" nowrap>&nbsp;</td>
														<td width="10" valign="top">Records :</td>
														<td width="5">&nbsp;</td>
														<td width="40%">
															<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																<tr>
																	<td width="5">
																		<input checked id="optRecordSelection1" name="optRecordSelection" type="radio"
																			onclick="changeBaseTableRecordOptions()" align="" />
																	</td>
																	<td width="5">&nbsp;</td>
																	<td>
																		<label tabindex="-1"
																			for="optRecordSelection1"
																			class="radio">
																			All</label>
																	</td>
																	<td colspan="3">&nbsp;</td>
																</tr>
															</table>
														</td>
													</tr>
													<tr height="10">
														<td width="5">&nbsp;</td>
														<td width="85" nowrap valign="top"></td>
														<td width="5">&nbsp;</td>
														<td valign="top" colspan="2"></td>
														<td width="20" nowrap>&nbsp;</td>
														<td width="10" valign="top"></td>
														<td width="5">&nbsp;</td>
														<td width="40%">
															<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																<tr>
																	<td width="5">
																		<input id="optRecordSelection2" name="optRecordSelection" type="radio"
																			onclick="changeBaseTableRecordOptions()" />
																	</td>
																	<td width="5">&nbsp;</td>
																	<td width="50" nowrap>
																		<label
																			tabindex="-1"
																			for="optRecordSelection2"
																			class="radio">
																			Picklist</label>
																	</td>
																	<td width="5">&nbsp;</td>
																	<td>
																		<input id="txtBasePicklist" name="txtBasePicklist" disabled="disabled" class="text textdisabled" style="WIDTH: 100%">
																	</td>
																	<td width="30" nowrap>
																		<input id="cmdBasePicklist" name="cmdBasePicklist" style="WIDTH: 100%" type="button" disabled="disabled" class="btn btndisabled" value="..."
																			onclick="selectRecordOption('base', 'picklist')" />
																	</td>
																</tr>
															</table>
														</td>
													</tr>
													<tr height="10">
														<td width="5">&nbsp;</td>
														<td width="85" nowrap valign="top"></td>
														<td width="5">&nbsp;</td>
														<td valign="top" colspan="2"></td>
														<td width="20" nowrap>&nbsp;</td>
														<td width="10" valign="top"></td>
														<td width="5">&nbsp;</td>
														<td width="40%">
															<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																<tr>
																	<td width="5">
																		<input id="optRecordSelection3" name="optRecordSelection" type="radio"
																			onclick="changeBaseTableRecordOptions()" />
																	</td>
																	<td width="5">&nbsp;</td>
																	<td width="50" nowrap>
																		<label
																			tabindex="-1"
																			for="optRecordSelection3"
																			class="radio">
																			Filter</label>
																	</td>
																	<td width="5">&nbsp;</td>
																	<td>
																		<input id="txtBaseFilter" name="txtBaseFilter" class="text textdisabled" disabled="disabled" style="WIDTH: 100%">
																	</td>
																	<td width="30" nowrap>
																		<input id="cmdBaseFilter" name="cmdBaseFilter" style="WIDTH: 100%" type="button" disabled="disabled" value="..." class="btn btndisabled"
																			onclick="selectRecordOption('base', 'filter')" />
																	</td>
																</tr>
															</table>
														</td>
													</tr>
													<tr height="10">
														<td width="5">&nbsp;</td>
														<td width="85" nowrap valign="top"></td>
														<td width="5">&nbsp;</td>
														<td valign="top" colspan="2"></td>
														<td width="20" nowrap>&nbsp;</td>
														<td width="10" valign="top"></td>
														<td width="5">&nbsp;</td>
														<td width="40%">
															<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																<tr>
																	<td colspan="6" height="5"></td>
																</tr>
																<tr>
																	<td colspan="6" nowrap>
																		<input name="chkPrintFilter" id="chkPrintFilter" type="checkbox" disabled="disabled" tabindex="-1"
																			onclick="changeTab1Control();" />
																		<label
																			id="lblPrintFilter"
																			name="lblPrintFilter"
																			for="chkPrintFilter"
																			class="checkbox checkboxdisabled"
																			tabindex="0">
																			Display filter or picklist title in the report header</label>
																	</td>
																</tr>
																<tr>
																	<td colspan="6" height="5"></td>
																</tr>
															</table>
														</td>
														<td width="5">&nbsp;</td>
													</tr>
													<tr>
														<td colspan="10" height="30"></td>
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
												<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
													<tr>
														<td colspan="9" height="5"></td>
													</tr>
													<tr height="10">
														<td width="5">&nbsp;</td>
														<td width="110" nowrap valign="top">Parent Table 1 :</td>
														<td width="5">&nbsp;</td>
														<td width="40%" valign="top">
															<input id="txtParent1" name="txtParent1" class="text textdisabled" style="WIDTH: 100%" disabled="disabled">
														</td>
														<td width="20" nowrap>&nbsp;</td>
														<td width="10" valign="top">Records :</td>
														<td width="5">&nbsp;</td>
														<td width="40%">
															<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																<tr>
																	<td width="5">
																		<input checked id="optParent1RecordSelection1" name="optParent1RecordSelection" type="radio"
																			onclick="changeParent1TableRecordOptions()" />
																	</td>
																	<td width="5">&nbsp;</td>
																	<td width="30">
																		<label
																			tabindex="-1"
																			for="optParent1RecordSelection1"
																			class="radio">
																			All</label>
																	</td>
																	<td>&nbsp;</td>
																</tr>
																<tr>
																	<td colspan="3" height="5"></td>
																</tr>
																<tr>
																	<td width="5">
																		<input id="optParent1RecordSelection2" name="optParent1RecordSelection" type="radio"
																			onclick="changeParent1TableRecordOptions()" />
																	</td>
																	<td width="5">&nbsp;</td>
																	<td width="20">
																		<label
																			tabindex="-1"
																			for="optParent1RecordSelection2"
																			class="radio">
																			Picklist</label>
																	</td>
																	<td width="5">&nbsp;</td>
																	<td>
																		<input id="txtParent1Picklist" name="txtParent1Picklist" disabled="disabled" style="WIDTH: 100%" class="text textdisabled">
																	</td>
																	<td width="30">
																		<input id="cmdParent1Picklist" name="cmdParent1Picklist" style="WIDTH: 100%" type="button" class="btn btndisabled" value="..." disabled="disabled"
																			onclick="selectRecordOption('p1', 'picklist')" />
																	</td>
																</tr>
																<tr>
																	<td colspan="3" height="5"></td>
																</tr>
																<tr>
																	<td width="5">
																		<input id="optParent1RecordSelection3" name="optParent1RecordSelection" type="radio"
																			onclick="changeParent1TableRecordOptions()" />
																	</td>
																	<td width="5">&nbsp;</td>
																	<td width="20">
																		<label
																			tabindex="-1"
																			for="optParent1RecordSelection3"
																			class="radio">
																			Filter</label>
																	</td>
																	<td width="5">&nbsp;</td>
																	<td>
																		<input id="txtParent1Filter" name="txtParent1Filter" class="text textdisabled" disabled="disabled" style="WIDTH: 100%">
																	</td>
																	<td width="30">
																		<input id="cmdParent1Filter" name="cmdParent1Filter" style="WIDTH: 100%" type="button" value="..." disabled="disabled" class="btn btndisabled"
																			onclick="selectRecordOption('p1', 'filter')" />
																	</td>
																</tr>
															</table>
														</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr height="15">
														<td colspan="9">
															<hr>
														</td>
													</tr>

													<tr height="10">
														<td width="5">&nbsp;</td>
														<td width="110" nowrap valign="top">Parent Table 2 :</td>
														<td width="5">&nbsp;</td>
														<td width="40%" valign="top">
															<input id="txtParent2" name="txtParent2" class="text textdisabled" style="WIDTH: 100%" disabled="disabled">
														</td>
														<td width="20" nowrap>&nbsp;</td>
														<td width="10" valign="top">Records :</td>
														<td width="5">&nbsp;</td>
														<td width="40%">
															<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																<tr>
																	<td width="5">
																		<input checked id="optParent2RecordSelection1" name="optParent2RecordSelection" type="radio"
																			onclick="changeParent2TableRecordOptions()" />
																	</td>
																	<td width="5">&nbsp;</td>
																	<td width="30">
																		<label
																			tabindex="-1"
																			for="optParent2RecordSelection1"
																			class="radio">
																			All</label>
																	</td>
																	<td>&nbsp;</td>
																</tr>
																<tr>
																	<td colspan="3" height="5"></td>
																</tr>
																<tr>
																	<td width="5">
																		<input id="optParent2RecordSelection2" name="optParent2RecordSelection" type="radio"
																			onclick="changeParent2TableRecordOptions()" />
																	</td>
																	<td width="5">&nbsp;</td>
																	<td width="20">
																		<label
																			tabindex="-1"
																			for="optParent2RecordSelection2"
																			class="radio">
																			Picklist</label>
																	</td>
																	<td width="5">&nbsp;</td>
																	<td>
																		<input id="txtParent2Picklist" name="txtParent2Picklist" disabled="disabled" style="WIDTH: 100%" class="text textdisabled">
																	</td>
																	<td width="30">
																		<input id="cmdParent2Picklist" name="cmdParent2Picklist" style="WIDTH: 100%" type="button" class="btn btndisabled" value="..." disabled="disabled"
																			onclick="selectRecordOption('p2', 'picklist')" />
																	</td>
																</tr>
																<tr>
																	<td colspan="3" height="5"></td>
																</tr>
																<tr>
																	<td width="5">
																		<input id="optParent2RecordSelection3" name="optParent2RecordSelection" type="radio"
																			onclick="changeParent2TableRecordOptions()" />
																	</td>
																	<td width="5">&nbsp;</td>
																	<td width="20">
																		<label
																			tabindex="-1"
																			for="optParent2RecordSelection3"
																			class="radio">
																			Filter</label>
																	</td>
																	<td width="5">&nbsp;</td>
																	<td>
																		<input id="txtParent2Filter" name="txtParent2Filter" class="text textdisabled" disabled="disabled" style="WIDTH: 100%">
																	</td>
																	<td width="30">
																		<input id="cmdParent2Filter" name="cmdParent2Filter" style="WIDTH: 100%" type="button" value="..." disabled="disabled" class="btn btndisabled"
																			onclick="selectRecordOption('p2', 'filter')" />
																	</td>
																</tr>
															</table>
														</td>
														<td width="5">&nbsp;</td>
													</tr>


													<tr height="15">
														<td colspan="9">
															<hr>
														</td>
													</tr>

													<tr height="5">
														<td width="5"></td>
														<td width="90" nowrap colspan="7">Child Tables :</td>
														<td width="5"></td>
													</tr>
													<tr>
														<td width="5">&nbsp;</td>
														<td colspan="7">
															<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
																<tr>
																	<td colspan="3" height="5"></td>
																</tr>
																<tr height="5">
																	<td rowspan="7">
																		<%Html.RenderPartial("Util_Def_CustomReports/ssOleDBGridChildren")%>
																	</td>

																	<td width="10">&nbsp;</td>
																	<td width="90">
																		<input type="button" id="cmdAddChild" name="cmdAddChild" value="Add..." style="WIDTH: 100%" class="btn"
																			onclick="childAdd()" />
																	</td>
																	<td width="5">&nbsp;</td>
																</tr>

																<tr height="5">
																	<td colspan="3"></td>
																</tr>

																<tr height="5">

																	<td width="5">&nbsp;</td>
																	<td width="90">
																		<input type="button" id="cmdEditChild" name="cmdChildEdit" value="Edit..." style="WIDTH: 100%" class="btn"
																			onclick="childEdit()" />
																	</td>
																	<td width="5">&nbsp;</td>
																</tr>

																<tr height="11">
																	<td colspan="3"></td>
																</tr>

																<tr height="5">

																	<td width="5">&nbsp;</td>
																	<td width="90">
																		<input type="button" id="cmdRemoveChild" name="cmdRemoveChild" value="Remove" style="WIDTH: 100%" class="btn"
																			onclick="childRemove()" />
																	</td>
																	<td width="5">&nbsp;</td>
																</tr>

																<tr height="5">
																	<td colspan="3"></td>
																</tr>

																<tr height="5">

																	<td width="5">&nbsp;</td>
																	<td width="90">
																		<input type="button" id="cmdRemoveAllChilds" name="cmdRemoveAllChilds" value="Remove All" style="WIDTH: 100%" class="btn"
																			onclick="childRemoveAll()" />
																	</td>
																	<td width="5">&nbsp;</td>
																</tr>
															</table>
														</td>
														<td width="5">&nbsp;</td>
													</tr>
													<!---------------------------------------------------------------------------------------------->

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
												<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
													<tr height="5">
														<td colspan="7" height="5"></td>
													</tr>
													<tr height="5">
														<td width="5" height="5"></td>
														<td valign="top" height="5">
															<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
																<tr height="5">
																	<td height="5" colspan="7" width="100%">
																		<select id="cboTblAvailable" name="cboTblAvailable" style="WIDTH: 100%;" disabled="disabled" class="combo combodisabled"
																			onchange="refreshAvailableColumns();">
																		</select>
																	</td>
																</tr>
																<tr height="10">
																	<td height="10" colspan="7" width="100%"></td>
																</tr>
																<tr height="5">
																	<td height="5"></td>
																	<td height="5">
																		<input id="optColumns" name="optAvailType" type="radio" checked disabled="disabled"
																			onclick="refreshAvailableColumns();" />
																	</td>
																	<td height="5" width="5">
																		<label
																			tabindex="-1"
																			for="optColumns"
																			class="radio radiodisabled">
																			Columns</label>
																	</td>
																	<td width="5" height="5"></td>
																	<td height="5">
																		<input id="optCalc" name="optAvailType" type="radio" disabled="disabled"
																			onclick="refreshAvailableColumns();" />
																	</td>
																	<td width="5" height="5">
																		<label
																			tabindex="-1"
																			for="optCalc"
																			class="radio radiodisabled">
																			Calculations</label>
																	</td>
																	<td height="5"></td>
																	<tr height="10">
																		<td height="10" colspan="7" width="100%"></td>
																	</tr>
																</tr>
															</table>
														</td>
														<td width="10"></td>
														<td width="5" nowrap></td>
														<td width="10"></td>
														<td rowspan="3" width="40%" height="100%">
															<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
																<tr>
																	<td width="100%" height="100%">
																		<%Html.RenderPartial("Util_Def_CustomReports/ssOleDBGridSelectedColumns")%>
																	</td>
																</tr>
															</table>
														</td>
														<td width="5"></td>
													</tr>

													<tr height="5">
														<td height="5" colspan="7"></td>
													</tr>

													<tr>
														<td width="5"></td>
														<td rowspan="5" width="40%" height="100%">
															<%Html.RenderPartial("Util_Def_CustomReports/ssOleDBGridAvailableColumns")%>
														</td>
														<td width="10" nowrap></td>
														<td height="5" valign="top" align="center">
															<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																<tr height="25">
																	<td>&nbsp</td>
																	<td width="90" nowrap align="center">
																		<input type="button" name="cmdColumnAdd" id="cmdColumnAdd" value="Add..." style="WIDTH: 100%; HEIGHT: 100%" class="btn"
																			onclick="columnSwap(true)" />
																	</td>
																	<td>&nbsp</td>
																</tr>
																<tr height="5">
																	<td></td>
																</tr>
																<tr height="25">
																	<td></td>
																	<td width="90" nowrap align="center">
																		<input type="button" name="cmdColumnAddAll" id="cmdColumnAddAll" value="Add All" style="WIDTH: 100%; HEIGHT: 100%" class="btn"
																			onclick="columnSwapAll(true)" />
																	</td>
																	<td></td>
																</tr>
																<tr height="15">
																	<td></td>
																</tr>
																<tr height="25">
																	<td></td>
																	<td width="90" nowrap align="center">
																		<input type="button" name="cmdColumnRemove" id="cmdColumnRemove" value="Remove" style="WIDTH: 100%; HEIGHT: 100%" class="btn"
																			onclick="columnSwap(false)" />
																	</td>
																	<td></td>
																</tr>
																<tr height="5">
																	<td></td>
																</tr>
																<tr height="25">
																	<td></td>
																	<td width="90" nowrap align="center">
																		<input type="button" name="cmdColumnRemoveAll" id="cmdColumnRemoveAll" value="Remove All" style="WIDTH: 100%; HEIGHT: 100%" class="btn"
																			onclick="columnSwapAll(false)" />
																	</td>
																	<td></td>
																</tr>
																<tr height="15">
																	<td></td>
																</tr>
																<tr height="25">
																	<td></td>
																	<td width="90" nowrap align="center">
																		<input type="button" name="cmdColumnMoveUp" id="cmdColumnMoveUp" value="Up" style="WIDTH: 100%; HEIGHT: 100%" class="btn"
																			onclick="columnMove(true)" />
																	</td>
																	<td></td>
																</tr>
																<tr height="5">
																	<td></td>
																</tr>
																<tr height="25">
																	<td></td>
																	<td width="90" nowrap align="center">
																		<input type="button" name="cmdColumnMoveDown" id="cmdColumnMoveDown" value="Down" style="WIDTH: 100%; HEIGHT: 100%" class="btn"
																			onclick="columnMove(false)" />
																	</td>
																	<td></td>
																</tr>
															</table>
														</td>
														<td width="10" nowrap></td>
														<td width="5"></td>
													</tr>

													<tr height="5">
														<td colspan="7" height="5"></td>
													</tr>

													<tr height="5">
														<td width="5"></td>
														<td width="10"></td>
														<td width="80"></td>
														<td width="10"></td>
														<td valign="top">
															<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																<tr>
																	<td width="125">Heading :</td>
																	<td width="5"></td>
																	<td>
																		<input id="txtColHeading" name="txtColHeading" maxlength="50" class="text" style="WIDTH: 100%"
																			onchange="validateColHeading()"
																			onkeyup="validateColHeading();"
																			onblur="trimColHeading();">
																	</td>
																</tr>
																<tr>
																	<td width="125">Size :</td>
																	<td width="5"></td>
																	<td>
																		<input id="txtSize" name="txtSize" maxlength="50" class="text" style="WIDTH: 100%"
																			onchange="validateColSize();"
																			onkeyup="validateColSize();">
																	</td>
																</tr>
																<tr>
																	<td width="125">Decimals :</td>
																	<td width="5"></td>
																	<td>
																		<input id="txtDecPlaces" name="txtDecPlaces" maxlength="50" class="text" style="WIDTH: 100%"
																			onchange="validateColDecimals();"
																			onkeyup="validateColDecimals();">
																	</td>
																</tr>
															</table>
														</td>
														<td width="5"></td>
													</tr>

													<tr height="5">
														<td colspan="7" height="5"></td>
													</tr>

													<tr height="5">
														<td width="5"></td>

														<td width="10"></td>
														<td width="80"></td>
														<td width="10"></td>
														<td valign="top">
															<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																<tr>
																	<td width="33%" align="left" nowrap>
																		<input type="checkbox" name="chkColAverage" id="chkColAverage" tabindex="-1"
																			onclick="setAggregate(0)" />
																		<label
																			for="chkColAverage"
																			class="checkbox"
																			tabindex="0">
																			Average
																		</label>
																	</td>
																	<td width="33%" align="left" nowrap>
																		<input type="checkbox" name="chkColCount" id="chkColCount" tabindex="-1"
																			onclick="setAggregate(1)" />
																		<label
																			for="chkColCount"
																			class="checkbox"
																			tabindex="0">
																			Count 
																		</label>
																	</td>
																	<td width="33%" align="left" nowrap>
																		<input type="checkbox" name="chkColTotal" id="chkColTotal" tabindex="-1"
																			onclick="setAggregate(2)" />
																		<label
																			for="chkColTotal"
																			class="checkbox"
																			tabindex="0">
																			Total 
																		</label>
																	</td>
																</tr>
																<tr height="5">
																	<td colspan="3" height="5"></td>
																</tr>
																<tr>
																	<td width="33%" align="left" nowrap>
																		<input type="checkbox" name="chkColHidden" id="chkColHidden" tabindex="-1"
																			onclick="setAggregate(3);" />
																		<label
																			for="chkColHidden"
																			class="checkbox"
																			tabindex="0">
																			Hidden
																		</label>
																	</td>
																	<td colspan="2" align="left" nowrap>
																		<input type="checkbox" name="chkColGroup" id="chkColGroup" tabindex="-1"
																			onclick="setAggregate(4)" />
																		<label
																			for="chkColGroup"
																			class="checkbox"
																			tabindex="0">
																			Group With Next 
																		</label>
																	</td>
																</tr>
															</table>
														</td>
														<td width="5"></td>
													</tr>

													<tr height="5">
														<td colspan="7" height="5"></td>
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
														<td colspan="3">Sort Order :</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td width="5">&nbsp;</td>
														<td rowspan="11">
															<%Html.RenderPartial("Util_Def_CustomReports/ssOleDBGridSortOrder")%>
														</td>

														<td width="10">&nbsp;</td>
														<td width="100">
															<input type="button" id="cmdSortAdd" name="cmdSortAdd" class="btn" value="Add..." style="WIDTH: 100%"
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
															<input type="button" id="cmdSortEdit" name="cmdSortEdit" class="btn" value="Edit..." style="WIDTH: 100%"
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
															<input type="button" id="cmdSortRemove" name="cmdSortRemove" class="btn" value="Remove" style="WIDTH: 100%"
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
															<input type="button" id="cmdSortRemoveAll" name="cmdSortRemoveAll" class="btn" value="Remove All" style="WIDTH: 100%"
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
															<input type="button" id="cmdSortMoveUp" name="cmdSortMoveUp" class="btn" value="Move Up" style="WIDTH: 100%"
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
															<input type="button" id="cmdSortMoveDown" name="cmdSortMoveDown" class="btn" value="Move Down" style="WIDTH: 100%"
																onclick="sortMove(false)" />
														</td>

														<td width="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td colspan="5"></td>
													</tr>

													<tr height="20">
														<td colspan="5">
															<hr>
														</td>
													</tr>

													<tr height="20">
														<td width="5">&nbsp;</td>
														<td colspan="3">Repetition :</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td width="5">&nbsp;</td>
														<td colspan="3" rowspan="9">
															<%Html.RenderPartial("Util_Def_CustomReports/ssOleDBGridRepetition")%>
														</td>

														<td width="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td colspan="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td colspan="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td colspan="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td colspan="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td colspan="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td colspan="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td colspan="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td colspan="5">&nbsp;</td>
													</tr>

												</table>
											</td>
										</tr>

									</table>
								</div>

								<!-- Fifth tab -->
								<div id="div5" style="visibility: hidden; display: none">
									<table width="100%" height="100%" class="outline" cellspacing="0" cellpadding="5">
										<tr valign="top">
											<td>
												<table width="100%" class="invisible" cellspacing="10" cellpadding="0">
													<tr>
														<td valign="top" colspan="2" width="100%" height="65">
															<table class="outline" cellspacing="0" cellpadding="4" width="100%" height="100%">
																<tr height="10">
																	<td height="10" align="left" valign="top">Report Options :
																		<br>
																		<br>
																		<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																			<tr height="20">
																				<td width="5">&nbsp</td>
																				<td align="left">
																					<input type="checkbox" id="chkSummary" name="chkSummary" tabindex="-1"
																						onclick="changeTab5Control()" />
																					<label
																						for="chkSummary"
																						class="checkbox"
																						tabindex="0">
																						Summary report</label>
																				</td>
																				<td width="5">&nbsp</td>
																			</tr>
																			<tr height="10">
																				<td colspan="3"></td>
																			</tr>
																			<tr height="20">
																				<td width="5">&nbsp</td>
																				<td align="left">
																					<input type="checkbox" id="chkIgnoreZeros" name="chkIgnoreZeros" tabindex="-1"
																						onclick="changeTab5Control()" />
																					<label
																						for="chkIgnoreZeros"
																						class="checkbox"
																						tabindex="0">
																						Ignore zeros when calculating aggregates</label>
																				</td>
																				<td width="5">&nbsp</td>
																			</tr>
																			<tr height="10">
																				<td colspan="3"></td>
																			</tr>
																		</table>
																	</td>
																</tr>
															</table>
														</td>
													</tr>
													<tr>
														<td valign="top" rowspan="2" width="25%" height="100%">
															<table class="outline" cellspacing="0" cellpadding="4" width="100%" height="100%">
																<tr height="10">
																	<td height="10" align="left" valign="top">Output Format :
																		<br>
																		<br>
																		<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																			<tr height="20">
																				<td width="5">&nbsp</td>
																				<td align="left" width="15">
																					<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat0" value="0"
																						onclick="formatClick(0);" />
																				</td>
																				<td align="left" nowrap>
																					<label
																						tabindex="-1"
																						for="optOutputFormat0"
																						class="radio">
																						Data Only</label>
																				</td>
																				<td width="5">&nbsp</td>
																			</tr>
																			<tr height="10">
																				<td colspan="4"></td>
																			</tr>
																			<tr height="20">
																				<td width="5">&nbsp</td>
																				<td align="left" width="15">
																					<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat1" value="1"
																						onclick="formatClick(1);" />
																				</td>
																				<td align="left" nowrap>
																					<label
																						tabindex="-1"
																						for="optOutputFormat1"
																						class="radio">
																						CSV File</label>
																				</td>
																				<td width="5">&nbsp</td>
																			</tr>
																			<tr height="10">
																				<td colspan="4"></td>
																			</tr>
																			<tr height="20">
																				<td width="5">&nbsp</td>
																				<td align="left" width="15">
																					<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat2" value="2"
																						onclick="formatClick(2);" />
																				</td>
																				<td align="left" nowrap>
																					<label
																						tabindex="-1"
																						for="optOutputFormat2"
																						class="radio">
																						HTML Document</label>
																				</td>
																				<td width="5">&nbsp</td>
																			</tr>
																			<tr height="10">
																				<td colspan="4"></td>
																			</tr>
																			<tr height="20">
																				<td width="5">&nbsp</td>
																				<td align="left" width="15">
																					<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat3" value="3"
																						onclick="formatClick(3);" />
																				</td>
																				<td align="left" nowrap>
																					<label
																						tabindex="-1"
																						for="optOutputFormat3"
																						class="radio">
																						Word Document</label>
																				</td>
																				<td width="5">&nbsp</td>
																			</tr>
																			<tr height="10">
																				<td colspan="4"></td>
																			</tr>
																			<tr height="20">
																				<td width="5">&nbsp</td>
																				<td align="left" width="15">
																					<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat4" value="4"
																						onclick="formatClick(4);" />
																				</td>
																				<td align="left" nowrap>
																					<label
																						tabindex="-1"
																						for="optOutputFormat4"
																						class="radio">
																						Excel Worksheet</label>
																				</td>
																				<td width="5">&nbsp</td>
																			</tr>
																			<tr height="10">
																				<td colspan="4"></td>
																			</tr>
																			<tr height="5">
																				<td width="5">&nbsp</td>
																				<td align="left" width="15">
																					<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat5" value="5"
																						onclick="formatClick(5);" />
																				</td>
																				<td>
																					<label
																						tabindex="-1"
																						for="optOutputFormat5"
																						class="radio">
																						Excel Chart</label>
																				</td>
																				<td width="5">&nbsp</td>
																			</tr>
																			<tr height="10">
																				<td colspan="4"></td>
																			</tr>
																			<tr height="5">
																				<td width="5">&nbsp</td>
																				<td align="left" width="15">
																					<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat6" value="6"
																						onclick="formatClick(6);" />
																				</td>
																				<td nowrap>
																					<label
																						tabindex="-1"
																						for="optOutputFormat6"
																						class="radio">
																						Excel Pivot Table</label>
																				</td>
																				<td width="5">&nbsp</td>
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
																				<td width="5">&nbsp</td>
																				<td align="left" colspan="6" nowrap>
																					<input name="chkPreview" id="chkPreview" type="checkbox" disabled="disabled" tabindex="-1"
																						onclick="changeTab5Control();" />
																					<label
																						for="chkPreview"
																						class="checkbox checkboxdisabled"
																						tabindex="0">
																						Preview on screen</label>
																				</td>
																				<td width="5">&nbsp</td>
																			</tr>

																			<tr height="10">
																				<td colspan="8"></td>
																			</tr>

																			<tr height="20">
																				<td></td>
																				<td align="left" colspan="6" nowrap>
																					<input name="chkDestination0" id="chkDestination0" type="checkbox" disabled="disabled" tabindex="-1"
																						onclick="changeTab5Control();" />
																					<label
																						for="chkDestination0"
																						class="checkbox checkboxdisabled"
																						tabindex="0">
																						Display output on screen</label>
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
																						onclick="changeTab5Control();" />
																					<label
																						for="chkDestination1"
																						class="checkbox checkboxdisabled"
																						tabindex="0">
																						Send to printer</label>
																				</td>
																				<td width="30" nowrap>&nbsp</td>
																				<td align="left" nowrap>Printer location : 
																				</td>
																				<td width="15">&nbsp</td>
																				<td colspan="2">
																					<select id="cboPrinterName" name="cboPrinterName" class="combo" style="WIDTH: 400px"
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
																						onclick="changeTab5Control();" />
																					<label
																						for="chkDestination2"
																						class="checkbox checkboxdisabled"
																						tabindex="0">
																						Save to file
																					</label>
																				</td>
																				<td nowrap></td>
																				<td align="left" nowrap>File name :   
																				</td>
																				<td nowrap></td>
																				<td colspan="2">
																					<table class="invisible" cellspacing="0" cellpadding="0" style="WIDTH: 400px">
																						<tr>
																							<td>
																								<input id="txtFilename" name="txtFilename" class="text textdisabled" disabled="disabled" tabindex="-1" style="WIDTH: 375px">
																							</td>
																							<td width="25">
																								<input id="cmdFilename" name="cmdFilename" class="btn" style="WIDTH: 100%" type="button" value="..."
																									onclick="saveFile(); changeTab5Control();" />
																							</td>
																						</tr>
																					</table>
																				</td>
																				<td width="5">&nbsp</td>
																			</tr>

																			<tr height="10">
																				<td colspan="8"></td>
																			</tr>

																			<tr height="20">
																				<td colspan="3"></td>
																				<td align="left" nowrap>If existing file :
																				</td>
																				<td></td>
																				<td colspan="2" width="100%" nowrap>
																					<select id="cboSaveExisting" name="cboSaveExisting" style="WIDTH: 400px" class="combo" onchange="changeTab5Control()"></select>
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
																						onclick="changeTab5Control();" />
																					<label
																						for="chkDestination3"
																						class="checkbox checkboxdisabled"
																						tabindex="0">
																						Send as email</label>
																				</td>
																				<td></td>
																				<td align="left" nowrap>Email group :   
																				</td>
																				<td></td>
																				<td colspan="2">
																					<table class="invisible" cellspacing="0" cellpadding="0" style="WIDTH: 400px">
																						<tr>
																							<td>
																								<input id="txtEmailGroup" name="txtEmailGroup" class="text textdisabled" disabled="disabled" tabindex="-1" style="WIDTH: 100%">
																								<input id="txtEmailGroupID" name="txtEmailGroupID" type="hidden">
																							</td>
																							<td width="25">
																								<input id="cmdEmailGroup" name="cmdEmailGroup" style="WIDTH: 100%" type="button" value="..." class="btn"
																									onclick="selectEmailGroup(); changeTab5Control();" />
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
																				<td align="left" nowrap>Email subject :   
																				</td>
																				<td></td>
																				<td colspan="2" width="100%" nowrap>
																					<input id="txtEmailSubject" class="text textdisabled" disabled="disabled" maxlength="255" name="txtEmailSubject" style="WIDTH: 400px"
																						onchange="frmUseful.txtChanged.value = 1;"
																						onkeydown="frmUseful.txtChanged.value = 1;">
																				</td>
																				<td width="5">&nbsp</td>
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
																					<input id="txtEmailAttachAs" class="text textdisabled" disabled="disabled" maxlength="255" name="txtEmailAttachAs" style="WIDTH: 400px"
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
											</td>
										</tr>
									</table>
								</div>

							</td>
							<td width="10"></td>
						</tr>

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
												onclick="okClick()" />
										</td>
										<td width="10"></td>
										<td width="80">
											<input type="button" id="cmdCancel" name="cmdCancel" value="Cancel" style="WIDTH: 100%" class="btn"
												onclick="cancelClick()" />
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
				</td>
			</tr>
		</table>

		<input type='hidden' id="txtBasePicklistID" name="txtBasePicklistID">
		<input type='hidden' id="txtBaseFilterID" name="txtBaseFilterID">

		<input type='hidden' id="txtParent1ID" name="txtParent1ID">
		<input type='hidden' id="txtParent2ID" name="txtParent2ID">
		<input type='hidden' id="txtParent1FilterID" name="txtParent1FilterID">
		<input type='hidden' id="txtParent1PicklistID" name="txtParent1PicklistID">
		<input type='hidden' id="txtParent2FilterID" name="txtParent2FilterID">
		<input type='hidden' id="txtParent2PicklistID" name="txtParent2PicklistID">

		<input type='hidden' id="txtBaseTableChildCount" name="txtBaseTableChildCount">
		<input type='hidden' id="txtDatabase" name="txtDatabase" value="<%=session("Database")%>">

		<input type='hidden' id="txtWordVer" name="txtWordVer" value="<%=Session("WordVer")%>">
		<input type='hidden' id="txtExcelVer" name="txtExcelVer" value="<%=Session("ExcelVer")%>">
		<input type='hidden' id="txtWordFormats" name="txtWordFormats" value="<%=Session("WordFormats")%>">
		<input type='hidden' id="txtExcelFormats" name="txtExcelFormats" value="<%=Session("ExcelFormats")%>">
		<input type='hidden' id="txtWordFormatDefaultIndex" name="txtWordFormatDefaultIndex" value="<%=Session("WordFormatDefaultIndex")%>">
		<input type='hidden' id="txtExcelFormatDefaultIndex" name="txtExcelFormatDefaultIndex" value="<%=Session("ExcelFormatDefaultIndex")%>">

		<input type='hidden' id="txtChangeCancelled" name="txtChangeCancelled" value="0">
		<input type='hidden' id="txtCheckingSuppressOptions" name="txtCheckingSuppressOptions" value="0">
	</form>

	<form id="frmTables" style="visibility: hidden; display: none">
		<%
			Dim sErrorDescription = ""
	
			' Get the table records.
			Dim cmdTables As New Command
			cmdTables.CommandText = "sp_ASRIntGetTablesInfo"
			cmdTables.CommandType = CommandTypeEnum.adCmdStoredProc
			cmdTables.ActiveConnection = Session("databaseConnection")
	
			Response.Write("<B>Set Connection</B>")
	
			Err.Clear()
			Dim rstTablesInfo = cmdTables.Execute
	
			Response.Write("<B>Executed SP</B>")
	
			If (Err.Number <> 0) Then
				sErrorDescription = "The tables information could not be retrieved." & vbCrLf & FormatError(Err.Description)
			End If

			If Len(sErrorDescription) = 0 Then
				Dim iCount = 0
				Do While Not rstTablesInfo.EOF
					Response.Write("<INPUT type='hidden' id=txtTableName_" & rstTablesInfo.Fields("tableID").Value & " name=txtTableName_" & rstTablesInfo.Fields("tableID").Value & " value=""" & rstTablesInfo.Fields("tableName").Value & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtTableType_" & rstTablesInfo.Fields("tableID").Value & " name=txtTableType_" & rstTablesInfo.Fields("tableID").Value & " value=" & rstTablesInfo.Fields("tableType").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtTableChildren_" & rstTablesInfo.Fields("tableID").Value & " name=txtTableChildren_" & rstTablesInfo.Fields("tableID").Value & " value=""" & rstTablesInfo.Fields("childrenString").Value & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtTableChildrenNames_" & rstTablesInfo.Fields("tableID").Value & " name=txtTableChildrenNames_" & rstTablesInfo.Fields("tableID").Value & " value=""" & rstTablesInfo.Fields("childrenNames").Value & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtTableParents_" & rstTablesInfo.Fields("tableID").Value & " name=txtTableParents_" & rstTablesInfo.Fields("tableID").Value & " value=""" & rstTablesInfo.Fields("parentsString").Value & """>" & vbCrLf)

					rstTablesInfo.MoveNext()
				Loop

				' Release the ADO recordset object.
				rstTablesInfo.Close()
				rstTablesInfo = Nothing
			End If
	
			' Release the ADO command object.
			cmdTables = Nothing
		%>
	</form>

	<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
		<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
	</form>

	<form id="frmOriginalDefinition" name="frmOriginalDefinition" style="visibility: hidden; display: none">
		<%
			Dim sErrMsg = ""

			If Session("action") <> "new" Then
				Dim cmdDefn As New Command
				cmdDefn.CommandText = "sp_ASRIntGetReportDefinition"
				cmdDefn.CommandType = CommandTypeEnum.adCmdStoredProc
				cmdDefn.ActiveConnection = Session("databaseConnection")
		
				Dim prmUtilID = cmdDefn.CreateParameter("utilID", 3, 1)	' 3=integer, 1=input
				cmdDefn.Parameters.Append(prmUtilID)
				prmUtilID.Value = CleanNumeric(Session("utilid"))

				Dim prmUser = cmdDefn.CreateParameter("user", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
				cmdDefn.Parameters.Append(prmUser)
				prmUser.Value = Session("username")

				Dim prmAction = cmdDefn.CreateParameter("action", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
				cmdDefn.Parameters.Append(prmAction)
				prmAction.Value = Session("action")

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

				Dim prmPicklistHidden = cmdDefn.CreateParameter("picklistHidden", 11, 2)	'11=bit, 2=output
				cmdDefn.Parameters.Append(prmPicklistHidden)

				Dim prmFilterID = cmdDefn.CreateParameter("filterID", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmFilterID)

				Dim prmFilterName = cmdDefn.CreateParameter("filterName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmFilterName)

				Dim prmFilterHidden = cmdDefn.CreateParameter("filterHidden", 11, 2)	'11=bit, 2=output
				cmdDefn.Parameters.Append(prmFilterHidden)

				Dim prmParent1TableID = cmdDefn.CreateParameter("parent1TableID", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmParent1TableID)

				Dim prmParent1TableName = cmdDefn.CreateParameter("parent1TableName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmParent1TableName)

				Dim prmParent1FilterID = cmdDefn.CreateParameter("parent1FilterID", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmParent1FilterID)

				Dim prmParent1FilterName = cmdDefn.CreateParameter("parent1FilterName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmParent1FilterName)

				Dim prmParent1FilterHidden = cmdDefn.CreateParameter("parent1FilterHidden", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmParent1FilterHidden)

				Dim prmParent2TableID = cmdDefn.CreateParameter("parent2TableID", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmParent2TableID)

				Dim prmParent2TableName = cmdDefn.CreateParameter("parent2TableName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmParent2TableName)

				Dim prmParent2FilterID = cmdDefn.CreateParameter("parent2FilterID", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmParent2FilterID)

				Dim prmParent2FilterName = cmdDefn.CreateParameter("parent2FilterName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmParent2FilterName)
		
				Dim prmParent2FilterHidden = cmdDefn.CreateParameter("parent2FilterHidden", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmParent2FilterHidden)

				'********************************************************************************

				Dim cmdReportChilds As New Command
				cmdReportChilds.CommandText = "sp_ASRIntGetReportChilds"
				cmdReportChilds.CommandType = CommandTypeEnum.adCmdStoredProc
				cmdReportChilds.ActiveConnection = Session("databaseConnection")
		
				Dim prmUtilID2 = cmdReportChilds.CreateParameter("utilID2", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput)
				cmdReportChilds.Parameters.Append(prmUtilID2)
				prmUtilID2.Value = CleanNumeric(Session("utilid"))
		
				Err.Clear()
				Dim rstChilds = cmdReportChilds.Execute
				Dim iHiddenChildFilterCount = 0
				Dim iCount = 0
				Dim sChildInfo = ""
		
				If (Err.Number <> 0) Then
					sErrMsg = "'" & Session("utilname") & "' report definition could not be read." & vbCrLf & FormatError(Err.Description)
				Else
					If rstChilds.State <> 0 Then
						' Read recordset values.
				
						Do While Not rstChilds.EOF
							iCount = iCount + 1
							Response.Write("<INPUT type='hidden' id=txtReportDefnChildTableID_" & iCount & " name=txtReportDefnChildTableID_" & iCount & " value=""" & rstChilds.Fields("TableID").Value & """>" & vbCrLf)
							Response.Write("<INPUT type='hidden' id=txtReportDefnChildTable_" & iCount & " name=txtReportDefnChildTable_" & iCount & " value=""" & rstChilds.Fields("Table").Value & """>" & vbCrLf)
							Response.Write("<INPUT type='hidden' id=txtReportDefnChildFilterID_" & iCount & " name=txtReportDefnChildFilterID_" & iCount & " value=""" & rstChilds.Fields("FilterID").Value & """>" & vbCrLf)
					
							Dim sTemp As String
							If IsDBNull(rstChilds.Fields("Filter").Value) Then
								sTemp = ""
							Else
								sTemp = Replace(rstChilds.Fields("Filter").Value, """", "&quot;")
							End If
							Response.Write("<INPUT type='hidden' id=txtReportDefnChildFilter_" & iCount & " name=txtReportDefnChildFilter_" & iCount & " value=""" & sTemp & """>" & vbCrLf)
					
							Response.Write("<INPUT type='hidden' id=txtReportDefnChildOrderID_" & iCount & " name=txtReportDefnChildOrderID_" & iCount & " value=""" & rstChilds.Fields("OrderID").Value & """>" & vbCrLf)
							Response.Write("<INPUT type='hidden' id=txtReportDefnChildOrder_" & iCount & " name=txtReportDefnChildOrder_" & iCount & " value=""" & rstChilds.Fields("Order").Value & """>" & vbCrLf)
							Response.Write("<INPUT type='hidden' id=txtReportDefnChildRecords_" & iCount & " name=txtReportDefnChildRecords_" & iCount & " value=""" & rstChilds.Fields("Records").Value & """>" & vbCrLf)
							Response.Write("<INPUT type='hidden' id=txtReportDefnChildGridString_" & iCount & " name=txtReportDefnChildGridString_" & iCount & " value=""" & Replace(rstChilds.Fields("gridstring").Value, """", "&quot;") & vbTab & rstChilds.Fields("FilterHidden").Value & """>" & vbCrLf)
							Response.Write("<INPUT type='hidden' id=txtReportDefnChildFilterHidden_" & iCount & " name=txtReportDefnChildFilterHidden_" & iCount & " value=""" & rstChilds.Fields("FilterHidden").Value & """>" & vbCrLf)

							' Check if the child table filter is a hidden calc.
							If rstChilds.Fields("FilterHidden").Value = "Y" Then
								iHiddenChildFilterCount = iHiddenChildFilterCount + 1
							End If

							If rstChilds.Fields("OrderDeleted").Value = "Y" Then
								If Len(sChildInfo) > 0 Then
									sChildInfo = sChildInfo & vbCrLf
								End If
								sChildInfo = sChildInfo & "The '" & rstChilds.Fields("Table").Value & "' table order will be removed from this definition as it has been deleted by another user."
							End If

							If rstChilds.Fields("FilterDeleted").Value = "Y" Then
								If Len(sChildInfo) > 0 Then
									sChildInfo = sChildInfo & vbCrLf
								End If
								sChildInfo = sChildInfo & "The '" & rstChilds.Fields("Table").Value & "' table filter will be removed from this definition as it has been deleted by another user."
							End If

							If rstChilds.Fields("FilterHiddenByOther").Value = "Y" Then
								If Len(sChildInfo) > 0 Then
									sChildInfo = sChildInfo & vbCrLf
								End If
								sChildInfo = sChildInfo & "The '" & rstChilds.Fields("Table").Value & "' table filter will be removed from this definition as it has been made hidden by another user."
							End If

							rstChilds.MoveNext()
						Loop
						' Release the ADO recordset object.
						rstChilds.Close()
					End If
					rstChilds = Nothing
				End If
				cmdReportChilds = Nothing

				Session("childcount") = iCount
				Session("hiddenfiltercount") = iHiddenChildFilterCount
		
				'********************************************************************************
		
				Dim prmSummary = cmdDefn.CreateParameter("summary", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmSummary)

				Dim prmPrintFilterHeader = cmdDefn.CreateParameter("printFilterHeader", 11, 2)	'11=bit, 2=output
				cmdDefn.Parameters.Append(prmPrintFilterHeader)

				'-----------------------------------------
				Dim prmOutputPreview = cmdDefn.CreateParameter("outputPreview", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmOutputPreview)
		
				Dim prmOutputFormat = cmdDefn.CreateParameter("outputFormat", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmOutputFormat)
		
				Dim prmOutputScreen = cmdDefn.CreateParameter("outputScreen", 11, 2)	'11=bit, 2=output
				cmdDefn.Parameters.Append(prmOutputScreen)
		
				Dim prmOutputPrinter = cmdDefn.CreateParameter("outputPrinter", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmOutputPrinter)
		
				Dim prmOutputPrinterName = cmdDefn.CreateParameter("outputPrinterName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmOutputPrinterName)
		
				Dim prmOutputSave = cmdDefn.CreateParameter("outputSave", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmOutputSave)
		
				Dim prmOutputSaveExisting = cmdDefn.CreateParameter("outputSaveExisting", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmOutputSaveExisting)
		
				Dim prmOutputEmail = cmdDefn.CreateParameter("outputEmail", 11, 2)	'11=bit, 2=output
				cmdDefn.Parameters.Append(prmOutputEmail)
		
				Dim prmOutputEmailAddr = cmdDefn.CreateParameter("outputEmailAddr", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmOutputEmailAddr)
		
				Dim prmOutputEmailName = cmdDefn.CreateParameter("outputEmailName", DataTypeEnum.adVarChar, ParameterDirectionEnum.adParamOutput, 8000)
				cmdDefn.Parameters.Append(prmOutputEmailName)
		
				Dim prmOutputEmailSubject = cmdDefn.CreateParameter("outputEmailSubject", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmOutputEmailSubject)

				Dim prmOutputEmailAttachAs = cmdDefn.CreateParameter("outputEmailAttachAs", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmOutputEmailAttachAs)

				Dim prmOutputFilename = cmdDefn.CreateParameter("outputFilename", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmOutputFilename)
		
				Dim prmTimestamp = cmdDefn.CreateParameter("timestamp", 3, 2)	' 3=integer, 2=output
				cmdDefn.Parameters.Append(prmTimestamp)

				Dim prmParent1AllRecords = cmdDefn.CreateParameter("parent1AllRecords", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmParent1AllRecords)

				Dim prmParent1PicklistID = cmdDefn.CreateParameter("parent1PicklistID", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmParent1PicklistID)

				Dim prmParent1PicklistName = cmdDefn.CreateParameter("parent1PicklistName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmParent1PicklistName)

				Dim prmParent1PicklistHidden = cmdDefn.CreateParameter("parent1PicklistHidden", 11, 2)	'11=bit, 2=output
				cmdDefn.Parameters.Append(prmParent1PicklistHidden)

				Dim prmParent2AllRecords = cmdDefn.CreateParameter("parent2AllRecords", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmParent2AllRecords)

				Dim prmParent2PicklistID = cmdDefn.CreateParameter("parent2PicklistID", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmParent2PicklistID)

				Dim prmParent2PicklistName = cmdDefn.CreateParameter("parent2PicklistName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmParent2PicklistName)

				Dim prmParent2PicklistHidden = cmdDefn.CreateParameter("parent2PicklistHidden", 11, 2)	'11=bit, 2=output
				cmdDefn.Parameters.Append(prmParent2PicklistHidden)

				Dim prmInfo = cmdDefn.CreateParameter("info", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmInfo)

				Dim prmIgnoreZeros = cmdDefn.CreateParameter("ignoreZeros", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmIgnoreZeros)
		
				Err.Clear()
				Dim rstDefinition = cmdDefn.Execute
		
				Dim iHiddenCalcCount = 0
				If (Err.Number <> 0) Then
					sErrMsg = "'" & Session("utilname") & "' report definition could not be read." & vbCrLf & FormatError(Err.Description)
				Else
					If rstDefinition.State <> 0 Then
						' Read recordset values.
						iCount = 0
						Do While Not rstDefinition.EOF
							iCount = iCount + 1
							If rstDefinition.Fields("definitionType").Value = "ORDER" Then
								Response.Write("<INPUT type='hidden' id=txtReportDefnOrder_" & iCount & " name=txtReportDefnOrder_" & iCount & " value=""" & rstDefinition.Fields("definitionString").Value & """>" & vbCrLf)
							ElseIf rstDefinition.Fields("definitionType").Value = "REPETITION" Then
								Response.Write("<INPUT type='hidden' id=txtReportDefnRepetition_" & iCount & " name=txtReportDefnRepetition_" & iCount & " value=""" & rstDefinition.Fields("definitionString").Value & """>" & vbCrLf)
							Else
								Response.Write("<INPUT type='hidden' id=txtReportDefnColumn_" & iCount & " name=txtReportDefnColumn_" & iCount & " value=""" & Replace(rstDefinition.Fields("definitionString").Value, """", "&quot;") & """>" & vbCrLf)
	
								' Check if the report column is a hidden calc.
								If rstDefinition.Fields("hidden").Value = "Y" Then
									iHiddenCalcCount = iHiddenCalcCount + 1
								End If
							End If
							rstDefinition.MoveNext()
						Loop

						' Release the ADO recordset object.
						rstDefinition.Close()
					End If
					rstDefinition = Nothing
			
					' NB. IMPORTANT ADO NOTE.
					' When calling a stored procedure which returns a recordset AND has output parameters
					' you need to close the recordset and set it to nothing before using the output parameters. 
					If Len(cmdDefn.Parameters("errMsg").Value) > 0 Then
						sErrMsg = "'" & Session("utilname") & "' " & cmdDefn.Parameters("errMsg").Value
					End If

					Response.Write("<INPUT type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Replace(cmdDefn.Parameters("name").Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Owner name=txtDefn_Owner value=""" & Replace(cmdDefn.Parameters("owner").Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Description name=txtDefn_Description value=""" & Replace(cmdDefn.Parameters("description").Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_BaseTableID name=txtDefn_BaseTableID value=" & cmdDefn.Parameters("baseTableID").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_AllRecords name=txtDefn_AllRecords value=" & cmdDefn.Parameters("allRecords").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_PicklistID name=txtDefn_PicklistID value=" & cmdDefn.Parameters("picklistID").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_PicklistName name=txtDefn_PicklistName value=""" & Replace(cmdDefn.Parameters("picklistName").Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_PicklistHidden name=txtDefn_PicklistHidden value=" & cmdDefn.Parameters("picklistHidden").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_FilterID name=txtDefn_FilterID value=" & cmdDefn.Parameters("filterID").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_FilterName name=txtDefn_FilterName value=""" & Replace(cmdDefn.Parameters("filterName").Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_FilterHidden name=txtDefn_FilterHidden value=" & cmdDefn.Parameters("filterHidden").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Parent1TableID name=txtDefn_Parent1TableID value=" & cmdDefn.Parameters("parent1TableID").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Parent1TableName name=txtDefn_Parent1TableName value=""" & cmdDefn.Parameters("parent1TableName").Value & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Parent1FilterID name=txtDefn_Parent1FilterID value=" & cmdDefn.Parameters("parent1FilterID").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Parent1FilterName name=txtDefn_Parent1FilterName value=""" & Replace(cmdDefn.Parameters("parent1FilterName").Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Parent1FilterHidden name=txtDefn_Parent1FilterHidden value=" & cmdDefn.Parameters("parent1FilterHidden").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Parent2TableID name=txtDefn_Parent2TableID value=" & cmdDefn.Parameters("parent2TableID").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Parent2TableName name=txtDefn_Parent2TableName value=""" & cmdDefn.Parameters("parent2TableName").Value & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Parent2FilterID name=txtDefn_Parent2FilterID value=" & cmdDefn.Parameters("parent2FilterID").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Parent2FilterName name=txtDefn_Parent2FilterName value=""" & Replace(cmdDefn.Parameters("parent2FilterName").Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Parent2FilterHidden name=txtDefn_Parent2FilterHidden value=" & cmdDefn.Parameters("parent2FilterHidden").Value & ">" & vbCrLf)

					Response.Write("<INPUT type='hidden' id=txtDefn_Summary name=txtDefn_Summary value=" & cmdDefn.Parameters("summary").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_PrintFilterHeader name=txtDefn_PrintFilterHeader value=" & cmdDefn.Parameters("printFilterHeader").Value & ">" & vbCrLf)

					Response.Write("<INPUT type='hidden' id=txtDefn_OutputPreview name=txtDefn_OutputPreview value=" & cmdDefn.Parameters("OutputPreview").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputFormat name=txtDefn_OutputFormat value=" & cmdDefn.Parameters("OutputFormat").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputScreen name=txtDefn_OutputScreen value=" & cmdDefn.Parameters("OutputScreen").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputPrinter name=txtDefn_OutputPrinter value=" & cmdDefn.Parameters("OutputPrinter").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputPrinterName name=txtDefn_OutputPrinterName value=""" & cmdDefn.Parameters("OutputPrinterName").Value & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputSave name=txtDefn_OutputSave value=" & cmdDefn.Parameters("OutputSave").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputSaveExisting name=txtDefn_OutputSaveExisting value=" & cmdDefn.Parameters("OutputSaveExisting").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmail name=txtDefn_OutputEmail value=" & cmdDefn.Parameters("OutputEmail").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmailAddr name=txtDefn_OutputEmailAddr value=" & cmdDefn.Parameters("OutputEmailAddr").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmailAddrName name=txtDefn_OutputEmailAddrName value=""" & Replace(cmdDefn.Parameters("outputEmailName").Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmailSubject name=txtDefn_OutputEmailSubject value=""" & Replace(cmdDefn.Parameters("OutputEmailSubject").Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputEmailAttachAs name=txtDefn_OutputEmailAttachAs value=""" & Replace(cmdDefn.Parameters("OutputEmailAttachAs").Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputFilename name=txtDefn_OutputFilename value=""" & cmdDefn.Parameters("OutputFilename").Value & """>" & vbCrLf)

					Response.Write("<INPUT type='hidden' id=txtDefn_Timestamp name=txtDefn_Timestamp value=" & cmdDefn.Parameters("timestamp").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_HiddenCalcCount name=txtDefn_HiddenCalcCount value=" & iHiddenCalcCount & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Parent1AllRecords name=txtDefn_Parent1AllRecords value=" & cmdDefn.Parameters("parent1AllRecords").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Parent1PicklistID name=txtDefn_Parent1PicklistID value=" & cmdDefn.Parameters("parent1PicklistID").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Parent1PicklistName name=txtDefn_Parent1PicklistName value=""" & Replace(cmdDefn.Parameters("parent1PicklistName").Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Parent1PicklistHidden name=txtDefn_Parent1PicklistHidden value=" & cmdDefn.Parameters("parent1PicklistHidden").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Parent2AllRecords name=txtDefn_Parent2AllRecords value=" & cmdDefn.Parameters("parent2AllRecords").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Parent2PicklistID name=txtDefn_Parent2PicklistID value=" & cmdDefn.Parameters("parent2PicklistID").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Parent2PicklistName name=txtDefn_Parent2PicklistName value=""" & Replace(cmdDefn.Parameters("parent2PicklistName").Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Parent2PicklistHidden name=txtDefn_Parent2PicklistHidden value=" & cmdDefn.Parameters("parent2PicklistHidden").Value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_IgnoreZeros name=txtDefn_IgnoreZeros value=" & cmdDefn.Parameters("ignoreZeros").Value & ">" & vbCrLf)
			
					Dim sInfo = cmdDefn.Parameters("info").Value
					If Len(sChildInfo) > 0 Then
						If Len(sInfo) > 0 Then
							sInfo = sInfo & vbCrLf
						End If
				
						sInfo = sInfo & sChildInfo
					End If
					Response.Write("<INPUT type='hidden' id=txtDefn_Info name=txtDefn_Info value=""" & Replace(sInfo, """", "&quot;") & """>" & vbCrLf)
				End If

				' Release the ADO command object.
				cmdDefn = Nothing

			Else
				Session("childcount") = 0
				Session("hiddenfiltercount") = 0
			End If
		%>
	</form>

	<form id="frmAccess">
		<%
			sErrorDescription = ""
	
			' Get the table records.
			Dim cmdAccess As New Command
			cmdAccess.CommandText = "spASRIntGetUtilityAccessRecords"
			cmdAccess.CommandType = CommandTypeEnum.adCmdStoredProc
			cmdAccess.ActiveConnection = Session("databaseConnection")

			Dim prmUtilType = cmdAccess.CreateParameter("utilType", 3, 1)	' 3=integer, 1=input
			cmdAccess.Parameters.Append(prmUtilType)
			prmUtilType.Value = 2	' 2 = custom report

			Dim prmUtilID3 = cmdAccess.CreateParameter("utilID", 3, 1) ' 3=integer, 1=input
			cmdAccess.Parameters.Append(prmUtilID3)
			If UCase(Session("action")) = "NEW" Then
				prmUtilID3.Value = 0
			Else
				prmUtilID3.Value = CleanNumeric(Session("utilid"))
			End If

			Dim prmFromCopy = cmdAccess.CreateParameter("fromCopy", 3, 1)	' 3=integer, 1=input
			cmdAccess.Parameters.Append(prmFromCopy)
			If UCase(Session("action")) = "COPY" Then
				prmFromCopy.Value = 1
			Else
				prmFromCopy.Value = 0
			End If

			Err.Clear()
			Dim rstAccessInfo = cmdAccess.Execute
			If (Err.Number <> 0) Then
				sErrorDescription = "The access information could not be retrieved." & vbCrLf & FormatError(Err.Description)
			End If

			If Len(sErrorDescription) = 0 Then
				Dim iCount = 0
				Do While Not rstAccessInfo.EOF
					Response.Write("<INPUT type='hidden' id=txtAccess_" & iCount & " name=txtAccess_" & iCount & " value=""" & rstAccessInfo.Fields("accessDefinition").Value & """>" & vbCrLf)

					iCount = iCount + 1
					rstAccessInfo.MoveNext()
				Loop

				' Release the ADO recordset object.
				rstAccessInfo.Close()
				rstAccessInfo = Nothing
			End If
	
			' Release the ADO command object.
			cmdAccess = Nothing
		%>
	</form>

	<form id="frmUseful" name="frmUseful" style="visibility: hidden; display: none">
		<input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
		<input type="hidden" id="txtLoading" name="txtLoading" value="Y">
		<input type="hidden" id="txtCurrentBaseTableID" name="txtCurrentBaseTableID">
		<input type="hidden" id="txtCurrentChildTableID" name="txtCurrentChildTableID" value="0">
		<input type="hidden" id="txtTablesChanged" name="txtTablesChanged">
		<input type="hidden" id="txtSelectedColumnsLoaded" name="txtSelectedColumnsLoaded" value="0">
		<input type="hidden" id="txtSortLoaded" name="txtSortLoaded" value="0">
		<input type="hidden" id="txtRepetitionLoaded" name="txtRepetitionLoaded" value="0">
		<input type="hidden" id="txtChildsLoaded" name="txtChildsLoaded" value="0">
		<input type="hidden" id="txtChanged" name="txtChanged" value="0">
		<input type="hidden" id="txtUtilID" name="txtUtilID" value='<%=session("utilid")%>'>
		<input type="hidden" id="txtChildCount" name="txtChildCount" value='<%=session("childcount")%>'>
		<input type="hidden" id="txtHiddenChildFilterCount" name="txtHiddenChildFilterCount" value='<%=session("hiddenfiltercount")%>'>
		<input type="hidden" id="txtLockGridEvents" name="txtLockGridEvents" value="0">
		<input type="hidden" id="txtChildColumnSelected" name="txtChildColumnSelected" value="0">
		<input type="hidden" id="txtGridActionCancelled" name="txtGridActionCancelled" value="0">
		<input type="hidden" id="txtGridChangeRecursive" name="txtGridChangeRecursive" value="0">

		<%
			Dim cmdDefinition As New Command
			cmdDefinition.CommandText = "sp_ASRIntGetModuleParameter"
			cmdDefinition.CommandType = CommandTypeEnum.adCmdStoredProc
			cmdDefinition.ActiveConnection = Session("databaseConnection")

			Dim prmModuleKey = cmdDefinition.CreateParameter("moduleKey", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
			cmdDefinition.Parameters.Append(prmModuleKey)
			prmModuleKey.Value = "MODULE_PERSONNEL"

			Dim prmParameterKey = cmdDefinition.CreateParameter("paramKey", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
			cmdDefinition.Parameters.Append(prmParameterKey)
			prmParameterKey.Value = "Param_TablePersonnel"

			Dim prmParameterValue = cmdDefinition.CreateParameter("paramValue", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
			cmdDefinition.Parameters.Append(prmParameterValue)

			Err.Clear()
			cmdDefinition.Execute()

			Response.Write("<INPUT type='hidden' id=txtPersonnelTableID name=txtPersonnelTableID value=" & cmdDefinition.Parameters("paramValue").Value & ">" & vbCrLf)
	
			cmdDefinition = Nothing

			Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtAction name=txtAction value=" & Session("action") & ">" & vbCrLf)
		%>
	</form>

	<form id="frmValidate" name="frmValidate" target="validate" method="post" action="util_validate_customreports" style="visibility: hidden; display: none">
		<input type="hidden" id="validateBaseFilter" name="validateBaseFilter" value="0">
		<input type="hidden" id="validateBasePicklist" name="validateBasePicklist" value="0">
		<input type="hidden" id="validateEmailGroup" name="validateEmailGroup" value="0">
		<input type="hidden" id="validateP1Filter" name="validateP1Filter" value="0">
		<input type="hidden" id="validateP1Picklist" name="validateP1Picklist" value="0">
		<input type="hidden" id="validateP2Filter" name="validateP2Filter" value="0">
		<input type="hidden" id="validateP2Picklist" name="validateP2Picklist" value="0">
		<input type="hidden" id="validateChildFilter" name="validateChildFilter" value="0">
		<input type="hidden" id="validateChildOrders" name="validateChildOrders" value="0">
		<input type="hidden" id="validateCalcs" name="validateCalcs" value=''>
		<input type="hidden" id="validateHiddenGroups" name="validateHiddenGroups" value=''>
		<input type="hidden" id="validateName" name="validateName" value=''>
		<input type="hidden" id="validateTimestamp" name="validateTimestamp" value=''>
		<input type="hidden" id="validateUtilID" name="validateUtilID" value=''>
	</form>

	<form id="frmSend" name="frmSend" method="post" action="util_def_customreports_Submit" style="visibility: hidden; display: none">
		<input type="hidden" id="txtSend_ID" name="txtSend_ID">
		<input type="hidden" id="txtSend_name" name="txtSend_name">
		<input type="hidden" id="txtSend_description" name="txtSend_description">
		<input type="hidden" id="txtSend_baseTable" name="txtSend_baseTable">
		<input type="hidden" id="txtSend_allRecords" name="txtSend_allRecords">
		<input type="hidden" id="txtSend_picklist" name="txtSend_picklist">
		<input type="hidden" id="txtSend_filter" name="txtSend_filter">
		<input type="hidden" id="txtSend_parent1Table" name="txtSend_parent1Table">
		<input type="hidden" id="txtSend_parent1AllRecords" name="txtSend_parent1AllRecords">
		<input type="hidden" id="txtSend_parent1Filter" name="txtSend_parent1Filter">
		<input type="hidden" id="txtSend_parent1Picklist" name="txtSend_parent1Picklist">
		<input type="hidden" id="txtSend_parent2Table" name="txtSend_parent2Table">
		<input type="hidden" id="txtSend_parent2AllRecords" name="txtSend_parent2AllRecords">
		<input type="hidden" id="txtSend_parent2Filter" name="txtSend_parent2Filter">
		<input type="hidden" id="txtSend_parent2Picklist" name="txtSend_parent2Picklist">
		<input type="hidden" id="txtSend_childTable" name="txtSend_childTable">
		<input type="hidden" id="txtSend_summary" name="txtSend_summary">
		<input type="hidden" id="txtSend_IgnoreZeros" name="txtSend_IgnoreZeros">
		<input type="hidden" id="txtSend_printFilterHeader" name="txtSend_printFilterHeader">
		<input type="hidden" id="txtSend_access" name="txtSend_access">
		<input type="hidden" id="txtSend_userName" name="txtSend_userName">
		<input type="hidden" id="txtSend_OutputPreview" name="txtSend_OutputPreview">admin
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
		<input type="hidden" id="txtSend_columns" name="txtSend_columns">
		<input type="hidden" id="txtSend_columns2" name="txtSend_columns2">
		<input type="hidden" id="txtSend_reaction" name="txtSend_reaction">
		<input type="hidden" id="txtSend_jobsToHide" name="txtSend_jobsToHide">
		<input type="hidden" id="txtSend_jobsToHideGroups" name="txtSend_jobsToHideGroups">
	</form>

	<form id="frmCustomReportChilds" name="frmCustomReportChilds" target="childselection" action="util_customreportchilds" method="post" style="visibility: hidden; display: none">
		<input type="hidden" id="childTableID" name="childTableID">
		<input type="hidden" id="childTable" name="childTable">
		<input type="hidden" id="childFilterID" name="childFilterID">
		<input type="hidden" id="childFilter" name="childFilter">
		<input type="hidden" id="childOrderID" name="childOrderID">
		<input type="hidden" id="childOrder" name="childOrder">
		<input type="hidden" id="childRecords" name="childRecords">
		<input type="hidden" id="childrenString" name="childrenString">
		<input type="hidden" id="childrenNames" name="childrenNames">
		<input type="hidden" id="selectedChildString" name="selectedChildString">
		<input type="hidden" id="childAction" name="childAction" value="NEW">
		<input type="hidden" id="childMax" name="childMax" value="5">
	</form>

	<form id="frmRecordSelection" name="frmRecordSelection" target="recordSelection" action="util_recordSelection" method="post" style="visibility: hidden; display: none">
		<input type="hidden" id="recSelType" name="recSelType">
		<input type="hidden" id="recSelTableID" name="recSelTableID">
		<input type="hidden" id="recSelCurrentID" name="recSelCurrentID">
		<input type="hidden" id="recSelTable" name="recSelTable">
		<input type="hidden" id="recSelDefOwner" name="recSelDefOwner">
		<input type="hidden" id="recSelDefType" name="recSelDefType">
	</form>

	<form id="frmEmailSelection" name="frmEmailSelection" target="emailSelection" action="util_emailSelection" method="post" style="visibility: hidden; display: none">
		<input type="hidden" id="EmailSelCurrentID" name="EmailSelCurrentID">
	</form>

	<form id="frmSortOrder" name="frmSortOrder" action="util_sortorderselection" target="sortorderselection" method="post" style="visibility: hidden; display: none">
		<input type="hidden" id="txtSortInclude" name="txtSortInclude">
		<input type="hidden" id="txtSortExclude" name="txtSortExclude">
		<input type="hidden" id="txtSortEditing" name="txtSortEditing">
		<input type="hidden" id="txtSortColumnID" name="txtSortColumnID">
		<input type="hidden" id="txtSortColumnName" name="txtSortColumnName">
		<input type="hidden" id="txtSortOrder" name="txtSortOrder">
		<input type="hidden" id="txtSortBOC" name="txtSortBOC">
		<input type="hidden" id="txtSortPOC" name="txtSortPOC">
		<input type="hidden" id="txtSortVOC" name="txtSortVOC">
		<input type="hidden" id="txtSortSRV" name="txtSortSRV">
	</form>

	<form id="frmSelectionAccess" name="frmSelectionAccess" style="visibility: hidden; display: none">
		<input type="hidden" id="forcedHidden" name="forcedHidden" value="N">
		<input type="hidden" id="baseHidden" name="baseHidden" value="N">
		<input type="hidden" id="p1Hidden" name="p1Hidden" value="N">
		<input type="hidden" id="p2Hidden" name="p2Hidden" value="N">
		<input type="hidden" id="childHidden" name="childHidden" value="0">
		<input type="hidden" id="calcsHiddenCount" name="calcsHiddenCount" value="0">
	</form>

	<input type='hidden' id="txtTicker" name="txtTicker" value="0">
	<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">

	<%Html.RenderPartial("Util_Def_CustomReports/grdColProps")%>
</div>

<script type="text/javascript">

	function updateCurrentColProp(psProp, pbValue) {
		with (grdColProps) {
			Columns(psProp).Value = pbValue;
		}
		return;
	}

	function getCurrentColProp(psProp) {
		with (grdColProps) {
			if (Columns(psProp).Value == "-1") {
				return true;
			}
			else {
				return false;
			}
		}
	}

</script>


<script type="text/javascript">
	util_def_customreports_onload();
	util_def_customreports_addhandlers();
</script>



