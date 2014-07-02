<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<script src="<%: Url.LatestContent("~/bundles/utilities_customreports")%>" type="text/javascript"></script>

<%--licence manager reference for activeX--%>
<object classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
	id="Microsoft_Licensed_Class_Manager_1_0"
	viewastext>
	<param name="LPKPath" value="<%: Url.Content("~/lpks/ssmain.lpk")%>">
</object>

<div <%=session("BodyTag")%>>
	<form id="frmDefinition" name="frmDefinition">

		<table align="center"  cellpadding="5" cellspacing="0" width="100%" height="100%">
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
									onclick="display_CustomReport_Page(1)" />
								<input type="button" value="Related Tables" id="btnTab2" name="btnTab2" class="btn"
									onclick="display_CustomReport_Page(2)" />
								<input type="button" value="Columns" id="btnTab3" name="btnTab3" class="btn"
									onclick="display_CustomReport_Page(3)" />
								<input type="button" value="Sort Order" id="btnTab4" name="btnTab4" class="btn"
									onclick="display_CustomReport_Page(4)" />
								<input type="button" value="Output" id="btnTab5" name="btnTab5" class="btn"
									onclick="display_CustomReport_Page(5)" />
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
									<table width="100%" height="100%"  cellspacing="0" cellpadding="5">
										<tr valign="top">
											<td>
												<table width="100%" class="invisible" cellspacing="0" cellpadding="0" style="border: 1px">
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

													<tr height="140px">
														<td width="5">&nbsp;</td>
														<td width="10" nowrap valign="top">Description :</td>
														<td width="5">&nbsp;</td>
														<td style="width: 40%;vertical-align: top" rowspan="2" colspan="2">
															<textarea id="txtDescription" name="txtDescription" class="textarea" style="HEIGHT: 99%; WIDTH: 100%" wrap="VIRTUAL" height="0" maxlength="255"
																onkeyup="changeTab1Control()">
													</textarea>
														</td>
														<td width="20" nowrap>&nbsp;</td>
														<td width="10" valign="top">Access :</td>
														<td width="5">&nbsp;</td>
														<td style="width: 40%; vertical-align: top">
															<%Html.RenderPartial("Util_Def_CustomReports/grdaccess")%>         
														</td>
														<td width="5">&nbsp;</td>
													</tr>

												<%--	<tr height="10">
														<td colspan="8">&nbsp;</td>
													</tr>

													<tr height="10">
														<td colspan="8">&nbsp;</td>
													</tr>--%>

													<tr style="height: 10px">
														<td width="5">&nbsp;</td>
														<td colspan="8">
															
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
																			onclick="changeBaseTableRecordOptions()" />
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
																		<input name="chkPrintFilter" id="chkPrintFilter" type="checkbox" disabled="disabled" tabindex="0"
																			onclick="changeTab1Control();" />
																		<label
																			id="lblPrintFilter"
																			name="lblPrintFilter"
																			for="chkPrintFilter"
																			class="checkbox checkboxdisabled"
																			tabindex="-1">
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
									<table width="100%" height="100%"  cellspacing="0" cellpadding="5">
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
																		<input id="txtParent1Picklist" name="txtParent1Picklist" disabled="disabled" style="WIDTH: 98%" class="text textdisabled">
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
																		<input id="txtParent1Filter" name="txtParent1Filter" class="text textdisabled" disabled="disabled" style="WIDTH: 98%">
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
																		<input id="txtParent2Picklist" name="txtParent2Picklist" disabled="disabled" style="WIDTH: 98%" class="text textdisabled">
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
																		<input id="txtParent2Filter" name="txtParent2Filter" class="text textdisabled" disabled="disabled" style="WIDTH: 98%">
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
															
														</td>
													</tr>

													<tr height="5">
														<td width="5"></td>
														<td width="90" nowrap colspan="7"><strong>Child Tables :</strong></td>
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
									<table width="100%" height="100%"  cellspacing="0" cellpadding="5">
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
																	<%--<tr height="10">
																		<td height="10" colspan="7" width="100%"></td>
																	</tr>--%>
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
																		<input type="checkbox" name="chkColAverage" id="chkColAverage" tabindex="0"
																			onclick="setAggregate(0)" />
																		<label
																			for="chkColAverage"
																			class="checkbox"
																			tabindex="-1">
																			Average
																		</label>
																	</td>
																	<td width="33%" align="left" nowrap>
																		<input type="checkbox" name="chkColCount" id="chkColCount" tabindex="0"
																			onclick="setAggregate(1)" />
																		<label
																			for="chkColCount"
																			class="checkbox"
																			tabindex="-1">
																			Count 
																		</label>
																	</td>
																	<td width="33%" align="left" nowrap>
																		<input type="checkbox" name="chkColTotal" id="chkColTotal" tabindex="0"
																			onclick="setAggregate(2)" />
																		<label
																			for="chkColTotal"
																			class="checkbox"
																			tabindex="-1">
																			Total 
																		</label>
																	</td>
																</tr>
																<tr height="5">
																	<td colspan="3" height="5"></td>
																</tr>
																<tr>
																	<td width="33%" align="left" nowrap>
																		<input type="checkbox" name="chkColHidden" id="chkColHidden" tabindex="1"
																			onclick="setAggregate(3);" />
																		<label
																			for="chkColHidden"
																			class="checkbox"
																			tabindex="-1">
																			Hidden
																		</label>
																	</td>
																	<td colspan="2" align="left" nowrap>
																		<input type="checkbox" name="chkColGroup" id="chkColGroup" tabindex="0"
																			onclick="setAggregate(4)" />
																		<label
																			for="chkColGroup"
																			class="checkbox"
																			tabindex="-1">
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
									<table width="100%" height="100%"  cellspacing="0" cellpadding="5">
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

													<tr>
														<td colspan="5" height="5"></td>
													</tr>

													<tr height="5">
														<td width="5">&nbsp;</td>
														<td rowspan="11" style="vertical-align: top">
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
															
														</td>
													</tr>

													<tr height="20">
														<td width="5">&nbsp;</td>
														<td colspan="3">Repetition :</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td width="5">&nbsp;</td>
														<td colspan="1" rowspan="9">
															<%Html.RenderPartial("Util_Def_CustomReports/ssOleDBGridRepetition")%>
														</td>

														<td width="5">&nbsp;</td>
													</tr>

<%--													<tr height="5">
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
													</tr>--%>

												</table>
											</td>
										</tr>

									</table>
								</div>

								<!-- Fifth tab -->
								<div id="div5" style="visibility: hidden; display: none">
									<table width="100%" height="100%"  cellspacing="0" cellpadding="5">
										<tr valign="top">
											<td>
												<table width="100%" class="invisible" cellspacing="10" cellpadding="0">
													<tr>
														<td valign="top" colspan="2" width="100%" height="65">
															<table  cellspacing="0" cellpadding="4" width="100%" height="100%">
																<tr height="10">
																	<td height="10" align="left" valign="top"><strong>Report Options :</strong>
																		<br>
																		<br>
																		<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																			<tr height="20">
																				<td width="5">&nbsp</td>
																				<td align="left">
																					<input type="checkbox" id="chkSummary" name="chkSummary" tabindex="0"
																						onclick="changeTab5Control()" />
																					<label
																						for="chkSummary"
																						class="checkbox"
																						tabindex="-1">
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
																					<input type="checkbox" id="chkIgnoreZeros" name="chkIgnoreZeros" tabindex="0"
																						onclick="changeTab5Control()" />
																					<label
																						for="chkIgnoreZeros"
																						class="checkbox"
																						tabindex="-1">
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
															<table  cellspacing="0" cellpadding="4" width="100%" height="100%">
																<tr height="10">
																	<td height="10" align="left" valign="top"><strong>Output Format :</strong>
																		<br>
																		<br>
																		<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																			<tr height="20">
																				<td width="5">&nbsp</td>
																				<td align="left" width="15">
																					<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat0" value="0"
																						onclick="formatClick(0);" />
																				</td>
																				<td style="text-align: left; white-space: nowrap;padding-left: 5px">
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
																				<td style="text-align: left; white-space: nowrap;padding-left: 5px">
																					<label
																						tabindex="-1"
																						for="optOutputFormat1"
																						class="radio ui-state-error-text">
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
																				<td style="text-align: left; white-space: nowrap;padding-left: 5px">
																					<label
																						tabindex="-1"
																						for="optOutputFormat2"
																						class="radio ui-state-error-text">
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
																				<td style="text-align: left; white-space: nowrap;padding-left: 5px">
																					<label
																						tabindex="-1"
																						for="optOutputFormat3"
																						class="radio ui-state-error-text">
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
																				<td style="text-align: left; white-space: nowrap;padding-left: 5px">
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
																				<td style="text-align: left; white-space: nowrap;padding-left: 5px">
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
																				<td style="text-align: left; white-space: nowrap;padding-left: 5px">
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
															<table cellspacing="0" cellpadding="4" width="100%" height="100%">
																<tr height="10">
																	<td height="10" align="left" valign="top"><strong>Output Destination(s) :</strong>
																		<br>
																		<br>
																		<table class="invisible" cellspacing="0" cellpadding="0" style="width: 100%; border:1px">
																			<tr height="20">
																				<td width="5">&nbsp</td>
																				<td align="left" colspan="6" nowrap>
																					<input name="chkPreview" id="chkPreview" type="checkbox" disabled="disabled" tabindex="0"
																						onclick="changeTab5Control();" />
																					<label
																						for="chkPreview"
																						class="checkbox"
																						tabindex="-1">
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
																					<input name="chkDestination0" id="chkDestination0" type="checkbox" disabled="disabled" tabindex="0"
																						onclick="changeTab5Control();" />
																					<label
																						for="chkDestination0"
																						class="checkbox"
																						tabindex="-1">
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
																					<input name="chkDestination1" id="chkDestination1" type="checkbox" disabled="disabled" tabindex="0"
																						onclick="changeTab5Control();"/>
																					<label
																						for="chkDestination1"
																						class="checkbox checkboxdisabled"
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

																			<tr style="height: 20px">
																				<td></td>
																				<td align="left" nowrap>
																					<input name="chkDestination2" id="chkDestination2" type="checkbox" disabled="disabled" tabindex="0"
																						onclick="changeTab5Control();" />
																					<label
																						for="chkDestination2"
																						class="checkbox"
																						tabindex="-1">
																						Save to file 
																					</label>
																				</td>
																				<td></td>
																				<td></td>
																				<td align="left" nowrap>File name : </td>
																				<td>
																					<input id="txtFilename" name="txtFilename"
																						style="width: 100%;"
																						class="text textdisabled" disabled="disabled" tabindex="-1">
																				</td>
																				<td width="25">
																					<input id="cmdFilename" name="cmdFilename" class="btn" type="button" value='...' disabled="disabled"
																						onclick="populateFileName(frmDefinition); changeTab5Control();" />
																				</td>
																				<td></td>
																			</tr>

																			<tr style="height: 20px">
																				<td></td>
																				<td></td>
																				<td></td>
																				<td></td>
																				<td style="white-space: nowrap;text-align: left" class="ui-state-error-text">If existing file :</td>
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

																			<tr style="height: 20px">
																				<td></td>
																				<td style="white-space: nowrap;text-align: left">
																					<input name="chkDestination3" id="chkDestination3" type="checkbox" disabled="disabled" tabindex="0"
																						onclick="changeTab5Control();"/>
																					<label for="chkDestination3"
																						class="checkbox"
																						tabindex="-1">
																						Send as email
																					</label>
																				</td>
																				<td></td>
																				<td></td>
																				<td style="white-space: nowrap;text-align: left">Email group :   </td>
																				<td>
																					<input id="txtEmailGroup" name="txtEmailGroup"
																						style="width: 100%;"
																						class="text textdisabled" disabled="disabled" tabindex="-1">
																					<input id="txtEmailGroupID" name="txtEmailGroupID" type="hidden" 
																						class="text textdisabled" disabled="disabled" tabindex="-1">
																				</td>
																				<td style="width: 25px;">
																					<input id="cmdEmailGroup" name="cmdEmailGroup" 
																						type="button" 
																						value='...' 
																						disabled="disabled" 
																						class="btn"
																						onclick="selectEmailGroup(); changeTab5Control();"/>
																				</td>
																				<td></td>
																			</tr>
																			<tr>
																				<td></td>
																				<td></td>
																				<td></td>
																				<td></td>
																				<td style="white-space: nowrap;">
																					<label for="txtEmailSubject"
																						tabindex="-1">
																						Email subject :</label>
																				</td>
																				<td>
																					<input id="txtEmailSubject"
																						style="width: 100%;"
																						class="text textdisabled" disabled="disabled" maxlength="255" name="txtEmailSubject"
																						onchange="frmUseful.txtChanged.value = 1;"
																						onkeydown="frmUseful.txtChanged.value = 1;">
																				</td>
																				<td></td>
																				<td></td>
																			</tr>
																			<tr>
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
											<td colspan="5" class="ui-state-error-text">Note: Options marked in red are unavailable in OpenHR Web.</td>
										</tr>
										</tr>
									</table>
								</div>
							</td>
							<td width="10"></td>
						</tr>

						<tr height="10">
							<td width="10"></td>
							<td>
								<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
									<tr>
										<td width="80">
											<input type="button" id="cmdOK" name="cmdOK" value="OK" style="WIDTH: 100%" class="btn"
												onclick="okClick()" />
										</td>
										<td width="10"></td>
										<td width="80">
											<input type="button" id="cmdCancel" name="cmdCancel" value="Cancel" style="WIDTH: 100%" class="btn"
												onclick="cancelClick()" />
										</td>
										<td>&nbsp;</td>
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
</div>

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
				Next
											
			Catch ex As Exception
				sErrorDescription = "The tables information could not be retrieved." & vbCrLf & ex.Message

			End Try
			
		%>
	</form>

	<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
		<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
	</form>

	<form id="frmOriginalDefinition" name="frmOriginalDefinition" style="visibility: hidden; display: none">
		<%
			Dim sErrMsg = ""
			
			If Session("action") <> "new" Then
								
				'********************************************************************************
				Dim iHiddenChildFilterCount = 0
				Dim iCount = 0
				Dim sChildInfo = ""
				
				Try

					Dim rstChilds = objDataAccess.GetFromSP("sp_ASRIntGetReportChilds" _
						, New SqlParameter("piReportID", SqlDbType.Int) With {.Value = CleanNumeric(Session("utilid"))})
								
					For Each objRow As DataRow In rstChilds.Rows

						iCount += 1
						Response.Write("<input type='hidden' id=txtReportDefnChildTableID_" & iCount & " name=txtReportDefnChildTableID_" & iCount & " value=""" & objRow("TableID") & """>" & vbCrLf)
						Response.Write("<input type='hidden' id=txtReportDefnChildTable_" & iCount & " name=txtReportDefnChildTable_" & iCount & " value=""" & objRow("Table") & """>" & vbCrLf)
						Response.Write("<input type='hidden' id=txtReportDefnChildFilterID_" & iCount & " name=txtReportDefnChildFilterID_" & iCount & " value=""" & objRow("FilterID") & """>" & vbCrLf)
					
						Dim sTemp As String
						If IsDBNull(objRow("Filter")) Then
							sTemp = ""
						Else
							sTemp = Replace(objRow("Filter"), """", "&quot;")
						End If
						Response.Write("<input type='hidden' id=txtReportDefnChildFilter_" & iCount & " name=txtReportDefnChildFilter_" & iCount & " value=""" & sTemp & """>" & vbCrLf)
					
						Response.Write("<input type='hidden' id=txtReportDefnChildOrderID_" & iCount & " name=txtReportDefnChildOrderID_" & iCount & " value=""" & objRow("OrderID") & """>" & vbCrLf)
						Response.Write("<input type='hidden' id=txtReportDefnChildOrder_" & iCount & " name=txtReportDefnChildOrder_" & iCount & " value=""" & objRow("Order") & """>" & vbCrLf)
						Response.Write("<input type='hidden' id=txtReportDefnChildRecords_" & iCount & " name=txtReportDefnChildRecords_" & iCount & " value=""" & objRow("Records") & """>" & vbCrLf)
						Response.Write("<input type='hidden' id=txtReportDefnChildGridString_" & iCount & " name=txtReportDefnChildGridString_" & iCount & " value=""" & Replace(objRow("gridstring"), """", "&quot;") & vbTab & objRow("FilterHidden") & """>" & vbCrLf)
						Response.Write("<input type='hidden' id=txtReportDefnChildFilterHidden_" & iCount & " name=txtReportDefnChildFilterHidden_" & iCount & " value=""" & objRow("FilterHidden") & """>" & vbCrLf)

						' Check if the child table filter is a hidden calc.
						If objRow("FilterHidden") = "Y" Then
							iHiddenChildFilterCount = iHiddenChildFilterCount + 1
						End If

						If objRow("OrderDeleted") = "Y" Then
							If Len(sChildInfo) > 0 Then
								sChildInfo = sChildInfo & vbCrLf
							End If
							sChildInfo = sChildInfo & "The '" & objRow("Table") & "' table order will be removed from this definition as it has been deleted by another user."
						End If

						If objRow("FilterDeleted") = "Y" Then
							If Len(sChildInfo) > 0 Then
								sChildInfo = sChildInfo & vbCrLf
							End If
							sChildInfo = sChildInfo & "The '" & objRow("Table") & "' table filter will be removed from this definition as it has been deleted by another user."
						End If

						If objRow("FilterHiddenByOther") = "Y" Then
							If Len(sChildInfo) > 0 Then
								sChildInfo = sChildInfo & vbCrLf
							End If
							sChildInfo = sChildInfo & "The '" & objRow("Table") & "' table filter will be removed from this definition as it has been made hidden by another user."
						End If

					Next
					
				Catch ex As Exception
					sErrMsg = "'" & Session("utilname") & "' report definition could not be read." & vbCrLf & FormatError(ex.Message)

				End Try
		
				Session("childcount") = iCount
				Session("hiddenfiltercount") = iHiddenChildFilterCount
		
				'********************************************************************************
				Dim prmErrMsg = New SqlParameter("psErrorMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmName = New SqlParameter("psReportName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmOwner = New SqlParameter("psReportOwner", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmDescription = New SqlParameter("psReportDesc", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmBaseTableID = New SqlParameter("piBaseTableID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmAllRecords = New SqlParameter("pfAllRecords", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmPicklistID = New SqlParameter("piPicklistID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmPicklistName = New SqlParameter("psPicklistName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmPicklistHidden = New SqlParameter("pfPicklistHidden", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmFilterID = New SqlParameter("piFilterID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmFilterName = New SqlParameter("psFilterName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmFilterHidden = New SqlParameter("pfFilterHidden", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmParent1TableID = New SqlParameter("piParent1TableID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmParent1TableName = New SqlParameter("psParent1Name", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmParent1FilterID = New SqlParameter("piParent1FilterID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmParent1FilterName = New SqlParameter("psParent1FilterName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmParent1FilterHidden = New SqlParameter("pfParent1FilterHidden", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmParent2TableID = New SqlParameter("piParent2TableID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmParent2TableName = New SqlParameter("psParent2Name", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmParent2FilterID = New SqlParameter("piParent2FilterID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmParent2FilterName = New SqlParameter("psParent2FilterName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmParent2FilterHidden = New SqlParameter("pfParent2FilterHidden", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}

				Dim prmSummary = New SqlParameter("pfSummary", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmPrintFilterHeader = New SqlParameter("pfPrintFilterHeader", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmOutputPreview = New SqlParameter("pfOutputPreview", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmOutputFormat = New SqlParameter("piOutputFormat", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmOutputScreen = New SqlParameter("pfOutputScreen", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmOutputPrinter = New SqlParameter("pfOutputPrinter", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmOutputPrinterName = New SqlParameter("psOutputPrinterName", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmOutputSave = New SqlParameter("pfOutputSave", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmOutputSaveExisting = New SqlParameter("piOutputSaveExisting", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmOutputEmail = New SqlParameter("pfOutputEmail", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmOutputEmailAddr = New SqlParameter("piOutputEmailAddr", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmOutputEmailName = New SqlParameter("psOutputEmailName", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmOutputEmailSubject = New SqlParameter("psOutputEmailSubject", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmOutputEmailAttachAs = New SqlParameter("psOutputEmailAttachAs", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmOutputFilename = New SqlParameter("psOutputFilename", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmTimestamp = New SqlParameter("piTimestamp", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmParent1AllRecords = New SqlParameter("pfParent1AllRecords", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmParent1PicklistID = New SqlParameter("piParent1PicklistID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmParent1PicklistName = New SqlParameter("psParent1PicklistName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmParent1PicklistHidden = New SqlParameter("pfParent1PicklistHidden", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmParent2AllRecords = New SqlParameter("pfParent2AllRecords", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmParent2PicklistID = New SqlParameter("piParent2PicklistID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
				Dim prmParent2PicklistName = New SqlParameter("psParent2PicklistName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmParent2PicklistHidden = New SqlParameter("pfParent2PicklistHidden", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
				Dim prmInfo = New SqlParameter("psInfoMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmIgnoreZeros = New SqlParameter("pfIgnoreZeros", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}

				Try
					
					Dim rstDefinition = objDataAccess.GetFromSP("sp_ASRIntGetReportDefinition" _
						, New SqlParameter("piReportID", SqlDbType.Int) With {.Value = CleanNumeric(Session("utilid"))} _
						, New SqlParameter("psCurrentUser", SqlDbType.VarChar, 255) With {.Value = Session("username")} _
						, New SqlParameter("psAction", SqlDbType.VarChar, 255) With {.Value = Session("action")} _
						, prmErrMsg, prmName, prmOwner, prmDescription, prmBaseTableID, prmAllRecords, prmPicklistID, prmPicklistName, prmPicklistHidden _
						, prmFilterID, prmFilterName, prmFilterHidden _
						, prmParent1TableID, prmParent1TableName, prmParent1FilterID, prmParent1FilterName, prmParent1FilterHidden _
						, prmParent2TableID, prmParent2TableName, prmParent2FilterID, prmParent2FilterName, prmParent2FilterHidden _
						, prmSummary, prmPrintFilterHeader, prmOutputPreview, prmOutputFormat, prmOutputScreen, prmOutputPrinter, prmOutputPrinterName _
						, prmOutputSave, prmOutputSaveExisting _
						, prmOutputEmail, prmOutputEmailAddr, prmOutputEmailName, prmOutputEmailSubject, prmOutputEmailAttachAs _
						, prmOutputFilename, prmTimestamp, prmParent1AllRecords, prmParent1PicklistID, prmParent1PicklistName, prmParent1PicklistHidden _
						, prmParent2AllRecords, prmParent2PicklistID, prmParent2PicklistName, prmParent2PicklistHidden _
						, prmInfo, prmIgnoreZeros)
		
					Dim iHiddenCalcCount = 0

					' Read recordset values.
					iCount = 0
					For Each objRow As DataRow In rstDefinition.Rows
							
						iCount += 1
						If objRow("definitionType").ToString() = "ORDER" Then
							Response.Write("<input type='hidden' id=txtReportDefnOrder_" & iCount & " name=txtReportDefnOrder_" & iCount & " value=""" & objRow("definitionString").ToString() & """>" & vbCrLf)
						ElseIf objRow("definitionType").ToString() = "REPETITION" Then
							Response.Write("<input type='hidden' id=txtReportDefnRepetition_" & iCount & " name=txtReportDefnRepetition_" & iCount & " value=""" & objRow("definitionString").ToString & """>" & vbCrLf)
						Else
							Response.Write("<input type='hidden' id=txtReportDefnColumn_" & iCount & " name=txtReportDefnColumn_" & iCount & " value=""" & Replace(objRow("definitionString").ToString(), """", "&quot;") & """>" & vbCrLf)
	
							' Check if the report column is a hidden calc.
							If objRow("hidden").ToString = "Y" Then
								iHiddenCalcCount += 1
							End If
						End If
					Next

					If Len(prmErrMsg.Value.ToString()) > 0 Then
						sErrMsg = "'" & Session("utilname").ToString() & "' " & prmErrMsg.Value.ToString
					End If
						
					Response.Write("<input type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Replace(prmName.Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Owner name=txtDefn_Owner value=""" & Replace(prmOwner.Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Description name=txtDefn_Description value=""" & Replace(prmDescription.Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_BaseTableID name=txtDefn_BaseTableID value=" & prmBaseTableID.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_AllRecords name=txtDefn_AllRecords value=" & prmAllRecords.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_PicklistID name=txtDefn_PicklistID value=" & prmPicklistID.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_PicklistName name=txtDefn_PicklistName value=""" & Replace(prmPicklistName.Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_PicklistHidden name=txtDefn_PicklistHidden value=" & prmPicklistHidden.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_FilterID name=txtDefn_FilterID value=" & prmFilterID.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_FilterName name=txtDefn_FilterName value=""" & Replace(prmFilterName.Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_FilterHidden name=txtDefn_FilterHidden value=" & prmFilterHidden.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Parent1TableID name=txtDefn_Parent1TableID value=" & prmParent1TableID.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Parent1TableName name=txtDefn_Parent1TableName value=""" & prmParent1TableName.Value & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Parent1FilterID name=txtDefn_Parent1FilterID value=" & prmParent1FilterID.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Parent1FilterName name=txtDefn_Parent1FilterName value=""" & Replace(prmParent1FilterName.Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Parent1FilterHidden name=txtDefn_Parent1FilterHidden value=" & prmParent1FilterHidden.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Parent2TableID name=txtDefn_Parent2TableID value=" & prmParent2TableID.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Parent2TableName name=txtDefn_Parent2TableName value=""" & prmParent2TableName.Value & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Parent2FilterID name=txtDefn_Parent2FilterID value=" & prmParent2FilterID.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Parent2FilterName name=txtDefn_Parent2FilterName value=""" & Replace(prmParent2FilterName.Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Parent2FilterHidden name=txtDefn_Parent2FilterHidden value=" & prmParent2FilterHidden.Value & ">" & vbCrLf)
						
					Response.Write("<input type='hidden' id=txtDefn_Summary name=txtDefn_Summary value=" & prmSummary.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_PrintFilterHeader name=txtDefn_PrintFilterHeader value=" & prmPrintFilterHeader.Value & ">" & vbCrLf)

					Response.Write("<input type='hidden' id=txtDefn_OutputPreview name=txtDefn_OutputPreview value=" & prmOutputPreview.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputFormat name=txtDefn_OutputFormat value=" & prmOutputFormat.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputScreen name=txtDefn_OutputScreen value=" & prmOutputScreen.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputPrinter name=txtDefn_OutputPrinter value=" & prmOutputPrinter.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputPrinterName name=txtDefn_OutputPrinterName value=""" & prmOutputPrinterName.Value & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputSave name=txtDefn_OutputSave value=" & prmOutputSave.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputSaveExisting name=txtDefn_OutputSaveExisting value=" & prmOutputSaveExisting.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputEmail name=txtDefn_OutputEmail value=" & prmOutputEmail.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputEmailAddr name=txtDefn_OutputEmailAddr value=" & prmOutputEmailAddr.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputEmailAddrName name=txtDefn_OutputEmailAddrName value=""" & Replace(prmOutputEmailName.Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputEmailSubject name=txtDefn_OutputEmailSubject value=""" & Replace(prmOutputEmailSubject.Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputEmailAttachAs name=txtDefn_OutputEmailAttachAs value=""" & Replace(prmOutputEmailAttachAs.Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputFilename name=txtDefn_OutputFilename value=""" & prmOutputFilename.Value & """>" & vbCrLf)
						
					Response.Write("<input type='hidden' id=txtDefn_Timestamp name=txtDefn_Timestamp value=" & prmTimestamp.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_HiddenCalcCount name=txtDefn_HiddenCalcCount value=" & iHiddenCalcCount & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Parent1AllRecords name=txtDefn_Parent1AllRecords value=" & prmParent1AllRecords.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Parent1PicklistID name=txtDefn_Parent1PicklistID value=" & prmParent1PicklistID.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Parent1PicklistName name=txtDefn_Parent1PicklistName value=""" & Replace(prmParent1PicklistName.Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Parent1PicklistHidden name=txtDefn_Parent1PicklistHidden value=" & prmParent1PicklistHidden.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Parent2AllRecords name=txtDefn_Parent2AllRecords value=" & prmParent2AllRecords.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Parent2PicklistID name=txtDefn_Parent2PicklistID value=" & prmParent2PicklistID.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Parent2PicklistName name=txtDefn_Parent2PicklistName value=""" & Replace(prmParent2PicklistName.Value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Parent2PicklistHidden name=txtDefn_Parent2PicklistHidden value=" & prmParent2PicklistHidden.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_IgnoreZeros name=txtDefn_IgnoreZeros value=" & prmIgnoreZeros.Value & ">" & vbCrLf)
			
					Dim sInfo = prmInfo.Value.ToString
					If Len(sChildInfo) > 0 Then
						If Len(sInfo) > 0 Then
							sInfo = sInfo & vbCrLf
						End If
				
						sInfo = sInfo & sChildInfo
					End If
					Response.Write("<input type='hidden' id=txtDefn_Info name=txtDefn_Info value=""" & Replace(sInfo, """", "&quot;") & """>" & vbCrLf)


				Catch ex As Exception
					sErrMsg = "'" & Session("utilname") & "' report definition could not be read." & vbCrLf & FormatError(ex.Message)

				End Try

			Else
				Session("childcount") = 0
				Session("hiddenfiltercount") = 0
			End If
		%>
	</form>

	<form id="frmAccess">
		<%
			Dim prmAccessUtilID = New SqlParameter("piID", SqlDbType.Int)
			Dim prmFromCopy = New SqlParameter("piFromCopy", SqlDbType.Int)
		
			sErrorDescription = ""

			Try
			
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
					, New SqlParameter("piUtilityType", SqlDbType.Int) With {.Value = UtilityType.utlCustomReport} _
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
			Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)

			Dim sParameterValue As String = objDatabase.GetModuleParameter("MODULE_PERSONNEL", "Param_TablePersonnel")
			Response.Write("<input type='hidden' id='txtPersonnelTableID' name='txtPersonnelTableID' value=" & sParameterValue & ">" & vbCrLf)
		
			Response.Write("<input type='hidden' id='txtErrorDescription' name='txtErrorDescription' value="""">" & vbCrLf)
			Response.Write("<input type='hidden' id='txtAction' name='txtAction' value=" & Session("action") & ">" & vbCrLf)

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

	<div style='height: 0;width:0; overflow:hidden;'>
		<input id="cmdGetFilename" name="cmdGetFilename" type="file" />
	</div>

	<%Html.RenderPartial("Util_Def_CustomReports/grdColProps")%>
<%--</div>--%>

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



