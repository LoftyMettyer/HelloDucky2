<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<script src="<%: Url.LatestContent("~/bundles/utilities_mailmerge")%>" type="text/javascript"></script>

<%--licence manager reference for activeX--%>
<object classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
	id="Microsoft_Licensed_Class_Manager_1_0"
	viewastext>
	<param name="LPKPath" value="<%: Url.Content("~/lpks/ssmain.lpk")%>">
</object>


<%  
	
	Dim objSession As SessionInfo = CType(Session("sessionContext"), SessionInfo)
	Dim bVersionOneEnabled = objSession.IsModuleEnabled("VERSIONONE")


%>

<div <%=session("BodyTag")%>>
	<form id="frmDefinition" name="frmDefinition">

		<table valign="top" align="center" cellpadding="5" cellspacing="0" width="100%" height="100%">
			<tr>
				<td>
					<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
						<tr height="5">
							<td colspan="3"></td>
						</tr>

						<tr height="10">
							<td width="10"></td>
							<td>
								<input type="button" value="Definition" id="btnTab1" name="btnTab1" class="btn btndisabled" disabled="disabled"
									onclick="display_MailMerge_Page(1)" />
								<input type="button" value="Columns" id="btnTab2" name="btnTab2" class="btn"
									onclick="display_MailMerge_Page(2)" />
								<input type="button" value="Sort Order" id="btnTab3" name="btnTab3" class="btn"
									onclick="display_MailMerge_Page(3)" />
								<input type="button" value="Output" id="btnTab4" name="btnTab4" class="btn"
									onclick="display_MailMerge_Page(4)" />
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
									<table width="100%" height="100%" cellspacing="0" cellpadding="5">
										<tr valign="top">
											<td>
												<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
													<td>
														<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
															<tr height="10">
																<td width="5">&nbsp;</td>
																<td width="10">Name :</td>
																<td width="5">&nbsp;</td>
																<td>
																	<input id="txtName" name="txtName" maxlength="50" style="width: 100%" class="text"
																		onkeyup="changeName()">
																</td>
																<td width="20">&nbsp;</td>
																<td style="padding-left: 10px; width: 10px">Owner :</td>
																<td width="5">&nbsp;</td>
																<td width="40%">
																	<input id="txtOwner" name="txtOwner" style="width: 100%" disabled="disabled" class="text textdisabled">
																</td>
																<td style="width: 40%">&nbsp;</td>
															</tr>

															<tr>
																<td colspan="9" height="5"></td>
															</tr>

															<tr>
																<td style="width: 5px">&nbsp;</td>
																<td style="width: 10px; white-space: nowrap; vertical-align: top">Description :</td>
																<td style="width: 5px">&nbsp;</td>
																<td style="width: 40%;vertical-align: top" rowspan="3">
																	<textarea id="txtDescription"
																		name="txtDescription"
																		class="textarea"
																		style="height: 99%; width: 100%" maxlength="255"
																		onkeyup="changeDescription()">
																	</textarea>
																</td>
																<td style="width: 20px; white-space: nowrap">&nbsp;</td>
																<td style="padding-left: 10px; width: 10px">Access :</td>
																<td style="width: 5px">&nbsp;</td>
																<td style="width: 40%; vertical-align: top" rowspan="3">
																	<%Html.RenderPartial("Util_Def_CustomReports/grdaccess")%>
																</td>
																<td style="width: 5px">&nbsp;</td>
															</tr>

															<tr height="10">
																<td colspan="7">&nbsp;</td>
															</tr>
															<tr height="10">
																<td colspan="7">&nbsp;</td>
															</tr>

															<tr>
																<td colspan="9">
																	<%--<hr>--%>
																</td>
															</tr>

															<tr height="10">
																<td width="5">&nbsp;</td>
																<td width="100" nowrap valign="top">Base Table :</td>
																<td width="5">&nbsp;</td>
																<td width="40%" valign="top">
																	<select id="cboBaseTable" name="cboBaseTable" class="combo" style="width: 100%"
																		onchange="changeBaseTable()">
																	</select>
																</td>
																<td width="20" nowrap>&nbsp;</td>
																<td style="padding-left: 10px; width: 10px">Records :</td>
																<td width="5">&nbsp;</td>
																<td width="40%">
																	<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																		<tr style="height: 35px">
																			<td width="5">
																				<input checked id="optRecordSelection1" name="optRecordSelection" type="radio"
																					onclick="changeBaseTableRecordOptions()"/>
																			</td>
																			<td width="5">&nbsp;</td>
																			<td width="30">
																				<label
																					tabindex="-1"
																					for="optRecordSelection1"
																					class="radio">
																					All
																				</label>
																			</td>
																			<td>&nbsp;</td>
																		</tr>
																		<tr>
																			<td width="5">
																				<input id="optRecordSelection2" name="optRecordSelection" type="radio"
																					onclick="changeBaseTableRecordOptions()" />
																			</td>
																			<td width="5">&nbsp;</td>
																			<td width="20">
																				<label
																					tabindex="-1"
																					for="optRecordSelection2"
																					class="radio">
																					Picklist</label>
																			</td>
																			<td width="5">&nbsp;</td>
																			<td>
																				<input id="txtBasePicklist" name="txtBasePicklist" disabled="disabled" style="width: 100%" class="text textdisabled">
																			</td>
																			<td width="30">
																				<input id="cmdBasePicklist" name="cmdBasePicklist" style="width: 100%" type="button" value="..." class="btn"
																					onclick="selectRecordOption('base', 'picklist')" />
																			</td>
																		</tr>
																		<tr>
																			<td width="5">
																				<input id="optRecordSelection3" name="optRecordSelection" type="radio"
																					onclick="changeBaseTableRecordOptions()" />
																			</td>
																			<td width="5">&nbsp;</td>
																			<td width="20">
																				<label
																					tabindex="-1"
																					for="optRecordSelection3"
																					class="radio">
																					Filter</label>
																			</td>
																			<td width="5">&nbsp;</td>
																			<td>
																				<input id="txtBaseFilter" name="txtBaseFilter" disabled="disabled" class="text textdisabled" style="width: 100%">
																			</td>
																			<td width="30">
																				<input id="cmdBaseFilter" name="cmdBaseFilter" style="width: 100%" type="button" value="..." class="btn"
																					onclick="selectRecordOption('base', 'filter')" />
																			</td>
																		</tr>
																	</table>
																</td>
																<td width="5">&nbsp;</td>
															</tr>

															<tr height="10">
																<td width="5">&nbsp;</td>
																<td width="90" nowrap>&nbsp;</td>
																<td width="5">&nbsp;</td>
																<td width="40%">
																	<input id="txtParent1" name="txtParent1" style="width: 100%" disabled="disabled" class="text textdisabled" type="hidden">
																</td>
																<td width="20" nowrap>&nbsp;</td>
																<td width="10">&nbsp;</td>
																<td width="5">&nbsp;</td>
																<td width="40%">
																	<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																		<tr>
																			<td>&nbsp;</td>
																			<td width="30">&nbsp;</td>
																		</tr>
																	</table>
																</td>
																<td width="5">&nbsp;</td>
															</tr>

															<tr height="10">
																<td width="5">&nbsp;</td>
																<td width="90">&nbsp;</td>
																<td width="5">&nbsp;</td>
																<td width="40%">
																	<input id="txtParent2" name="txtParent2" style="width: 100%" disabled="disabled" class="text textdisabled" type="hidden">
																</td>
																<td width="20" nowrap>&nbsp;</td>
																<td width="10">&nbsp;</td>
																<td width="5">&nbsp;</td>
																<td width="40%">
																	<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																		<tr>
																			<td>&nbsp;
																			</td>
																			<td width="30">&nbsp;</td>
																		</tr>
																	</table>
																</td>
																<td width="5">&nbsp;</td>
															</tr>
														</table>
													</td>
												</table>
									</table>
								</div>

								<!-- Second tab -->
								<div id="div2" style="visibility: hidden; display: none">
									<table width="100%" height="100%" cellspacing="0" cellpadding="5">
										<tr valign="top">
											<td>
												<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
													<tr height="5">
														<td colspan="7" height="5"></td>
													</tr>

													<tr height="5">
														<td width="5" height="5"></td>
														<td valign="top" height="5">
															<table style="width: 100%;padding-bottom: 10px" class="invisible" cellspacing="0" cellpadding="0">
																<tr height="25px">
																	<td height="25px" colspan="7" width="100%">
																		<select id="cboTblAvailable" name="cboTblAvailable" disabled="disabled" class="combo combodisabled" style="width: 100%; HEIGHT: 100%"
																			onchange="refreshAvailableColumns();">
																		</select>
																	</td>
																</tr>
																<tr height="10">
																	<td height="10" colspan="7" width="100%"></td>
																</tr>
																<tr height="5">
																	<td height="5"></td>
																	<td style="height: 5px; width: 5%">
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
																	<td style="height: 5px; width: 5%">
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
																</tr>
															</table>
														</td>
														<td width="10"></td>
														<td width="5" nowrap></td>
														<td width="10"></td>
														<td rowspan="4" width="40%" height="100%">
															<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
																<tr>
																	<td style="width: 5px"></td>
																</tr>
																<tr>
																	<td width="100%" height="100%">
																		<%Html.RenderPartial("Util_Def_CustomReports/ssOleDBGridSelectedColumns")%>
																	</td>
																</tr>
															</table>
														</td>
														<td  style="width: 100%"></td>
													</tr>

													<tr>
														<td width="5"></td>
														<td rowspan="4" width="40%" height="100%">
															<%Html.RenderPartial("Util_Def_CustomReports/ssOleDBGridAvailableColumns")%>
														</td>
														<td width="10" nowrap></td>
														<td height="5" valign="top" align="center">
															<table   style="width:100%;padding: 5px" class="invisible" cellspacing="0">
																<tr height="25">
																	<td>&nbsp</td>
																	<td width="100" nowrap align="center">
																		<input type="button" name="cmdColumnAdd" id="cmdColumnAdd" value="Add..." style="width: 100%; HEIGHT: 100%" class="btn"
																			onclick="columnSwap(true)"/>
																	</td>
																	<td>&nbsp;</td>
																</tr>
																<tr height="5">
																	<td></td>
																</tr>
																<tr height="25">
																	<td></td>
																	<td width="100" nowrap align="center">
																		<input type="button" name="cmdColumnAddAll" id="cmdColumnAddAll" value="Add All" style="width: 100%; HEIGHT: 100%" class="btn"
																			onclick="columnSwapAll(true)" />
																	</td>
																	<td></td>
																</tr>
																<tr style="height: 100px">
																	<td></td>
																</tr>
																<tr height="25">
																	<td></td>
																	<td width="100" nowrap align="center">
																		<input type="button" name="cmdColumnRemove" id="cmdColumnRemove" value="Remove" style="width: 100%; HEIGHT: 100%" class="btn"
																			onclick="columnSwap(false)" />
																	</td>
																	<td></td>
																</tr>
																<tr height="5">
																	<td></td>
																</tr>
																<tr height="25">
																	<td></td>
																	<td width="100" nowrap align="center">
																		<input type="button" name="cmdColumnRemoveAll" id="cmdColumnRemoveAll" value="Remove All" style="width: 100%; HEIGHT: 100%" class="btn"
																			onclick="columnSwapAll(false)" />
																	</td>
																	<td></td>
																</tr>
															</table>
														</td>
														<td width="10" nowrap></td>
														<td width="5"></td>
													</tr>

													<tr>
														<td colspan="5"></td>
													</tr>

													<tr height="5">
														<td colspan="6" height="5"></td>
													</tr>

													<tr height="5">
														<td width="5"></td>
														<td width="10"></td>
														<td width="80"></td>
														<td width="10"></td>
														<td valign="top">
															<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																<tr>
																	<td width="125">Size :</td>
																	<td width="5"></td>
																	<td>
																		<input id="txtSize" name="txtSize" maxlength="50" style="width: 100%" class="text"
																			onchange="validateColSize();"
																			onkeyup="validateColSize();">
																	</td>
																</tr>
																<tr>
																	<td width="125">Decimals :</td>
																	<td width="5"></td>
																	<td>
																		<input id="txtDecPlaces" name="txtDecPlaces" maxlength="50" style="width: 100%" class="text"
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
												</table>
											</td>
										</tr>
									</table>
								</div>

								<!-- Third tab -->
								<div id="div3" style="visibility: hidden; display: none">
									<table width="50%" height="80%" cellspacing="0" cellpadding="5">
										<tr valign="top">
											<td>
												<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
													<tr>
														<td colspan="5" height="5"></td>
													</tr>

													<tr height="20">
														<td width="5">&nbsp;</td>
														<%--<td colspan="3">Sort Order :</td>--%>
														<td colspan="3"></td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr height="5">
														<td width="5">&nbsp;</td>
														<td style="vertical-align: top" rowspan="12">
															<%Html.RenderPartial("Util_Def_CustomReports/ssMailMergeOleDBGridSortOrder")%>
																
														</td>

														<td width="10">&nbsp;</td>
														<td width="100">
															<input type="button" id="cmdSortAdd" name="cmdSortAdd" value="Add..." style="width: 100%" class="btn"
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
															<input type="button" id="cmdSortEdit" name="cmdSortEdit" value="Edit..." style="width: 100%" class="btn"
																onclick="sortEdit()"/>
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
															<input type="button" id="cmdSortRemove" name="cmdSortRemove" value="Remove" style="width: 100%" class="btn"
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
															<input type="button" id="cmdSortRemoveAll" name="cmdSortRemoveAll" value="Remove All" style="width: 100%" class="btn"
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
															<input type="button" id="cmdSortMoveUp" name="cmdSortMoveUp" value="Move Up" style="width: 100%" class="btn"
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
															<input type="button" id="cmdSortMoveDown" name="cmdSortMoveDown" value="Move Down" style="width: 100%" class="btn"
																onclick="sortMove(false)" />
														</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr>
														<td colspan="4">&nbsp;</td>
													</tr>

													<tr>
														<td colspan="5" height="5"></td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
								</div>

								<!-- Fourth tab -->
								<div id="div4" style="visibility: hidden; display: none">
									<table width="100%" height="80%" cellspacing="0" cellpadding="0">
										<tr valign="top">
											<td>
												<table style="width:100%;padding-bottom: 20px" class="invisible" cellspacing="0" cellpadding="4">
													<tr height="5">
														<td colspan="9"></td>
													</tr>
													<tr>
														<td width="20">&nbsp;&nbsp;&nbsp;</td>
														<td style="width: 100px;white-space: nowrap; font-weight: bold;">Template :</td>
														<td width="20">&nbsp;&nbsp;&nbsp;</td>
														<td style="width: 45%">
															<input id="txtTemplate" name="txtTemplate" style="width: 100%" class="text textdisabled" disabled="disabled">
														</td>
														<td width="30">
															<table style="width: 100%; padding: 0" class="invisible" cellspacing="0">
																<tr>
																	<td>
																		<input type="button" value="..." id="cmdTemplateSelect" name="cmdTemplateSelect" class="btn"
																			onclick="TemplateSelect()" />
																	</td>
																	<td>
																		<input type="button" value="Clear" id="cmdTemplateClear" name="cmdTemplateClear" class="btn"
																			onclick="TemplateClear()" />
																	</td>
																</tr>
															</table>
														</td>
														<td style="width: 100%" > </td>
														<td> </td>
														<td> </td>
														<td> </td>
													</tr>

													<tr>
														<td></td>
														<td></td>
														<td></td>
														<td nowrap>
															<input type="checkbox" id="chkPause" name="chkPause" tabindex="0"
																onclick="changeTab4Control()" />
															<label
																for="chkPause"
																class="checkbox"
																tabindex="-1">
																Pause before mail merge</label>
														</td>
														
														
														<td> </td>
														<td> </td>
														
														<td> </td>
														<td> </td>
														<td> </td>

													</tr>
													<tr>
														<td></td>
														<td></td>
														<td></td>
														<td nowrap>
															<input type="checkbox" id="chkSuppressBlanks" name="chkSuppressBlanks" tabindex="0"
																onclick="changeTab4Control()" />
																<label
																	for="chkSuppressBlanks"
																	class="checkbox"
																	tabindex="-1">Suppress blank lines</label>
														</td>
														<td> </td>
														<td> </td>
														
														<td> </td>
														<td> </td>
														<td> </td>
													</tr>
												</table>

												<table width="100%" class="invisible" cellspacing="0" cellpadding="0" height="100%">
													<tr style="height: 100%">
														<td></td>
														<td colspan="6">

															<table style="width: 100%; height: 100%">
																<tr>
																	<td style="width: 10px">&nbsp;</td>
																	<td width="220px" valign="top">
																		<table style="vertical-align: text-top" cellspacing="0" cellpadding="4" width="100%" height="200px">
																			<tr style="height: 20px; font-weight: bold;">
																				<td colspan="4" align="left" style="vertical-align: text-top">Output Format :
																					<br>
																				</td>
																			</tr>

																			<tr style="height: 20px">
																				<td width="5" style="vertical-align: text-top">
																					<input checked id="optDestination0" name="optDestination" type="radio"
																						onclick="refreshDestination(); changeTab4Control(); " />
																				</td>
																				<%--<td width="5">&nbsp;</td>--%>
																				<td width="130px" style="vertical-align: text-top; white-space: nowrap">
																					<label tabindex="-1"
																						for="optDestination0"
																						class="radio">
																						Word Document</label>
																				</td>
																				<%--<td>&nbsp;</td>--%>
																			</tr>
																			<tr style="height: 20px">
																				<td width="5">
																					<input id="optDestination1" name="optDestination" type="radio"
																						onclick="refreshDestination(); changeTab4Control(); " />
																				</td>
																				<%--<td width="5">&nbsp;</td>--%>
																				<td style="width: 130px; white-space: nowrap">
																					<label tabindex="-1"
																						for="optDestination1"
																						class="radio">
																						Individual Emails</label>
																				</td>
																				<td width="5px">&nbsp;</td>

																			</tr>
																			<%If Not bVersionOneEnabled Then%>
																			<tr style="height: 20px; visibility: hidden; display: none">
																				<%Else%>
																			<tr style="height: 20px;">
																				<%End If%>
																				<td width="5">
																					<input id="optDestination2" name="optDestination" type="radio"
																						onclick="refreshDestination(); changeTab4Control(); "/>
																				</td>
																				
																				<td style="white-space: nowrap">
																					<label tabindex="-1"
																						for="optDestination2"
																						class="radio ui-state-error-text">
																						Document Management</label>
																				</td>
																				<td width="5">&nbsp;</td>
																			</tr>
																			<tr></tr>
																		</table>
																		<td style="width: 10px"></td>
																	<td valign="top">
																		<table  style="width: 100%; height: 200px; vertical-align: top;">
																			<tr style="height: 20px; font-weight: bold;">
																				<td colspan="8" style="text-align: left">Output Destinations :
																					<br>
																				</td>
																			</tr>
																			<tr style="height: 20px; padding: 5px" name="row1" id="row1">
																				<td style="width: 60px;white-space: nowrap">Engine :</td>
																				<td style="width: 5px"></td>
																				<td colspan="2">
																					<select id="cboDMEngine" name="cboDMEngine" 
																						style="width: 400px" class="combo"
																						onchange="changeTab4Control()">
																					</select>
																				</td>
																			</tr>
																			<tr style="width: 5px"></tr>
																			<tr id="row4" name="row4" style="height: 20px; padding: 5px">
																				<td style="white-space: nowrap" colspan="2">
																					<input type="checkbox" id="chkOutputScreen" name="chkOutputScreen" tabindex="0"
																						onclick="changeTab4Control()" />
																				</td>
																				<td style="width: 200px; white-space: nowrap">
																					<label
																						for="chkOutputScreen"
																						class="checkbox"
																						tabindex="-1">
																						Display output on screen
																					</label>
																				</td>
																			</tr>
																			<tr style="height: 20px; padding: 5px" name="row2" id="row2">
																				<td nowrap></td>
																				<td></td>
																				<td style="width: 30px" colspan="3">
																					<table class="invisible" cellspacing="0" cellpadding="0">
																						<tr>
																							<td width="20"></td>
																							<td style="padding-right: 0; vertical-align: middle"></td>
																							<td></td>
																						</tr>
																					</table>
																				</td>
																			</tr>

																			<tr style="height: 20px" name="row3" id="row3">
																				<td nowrap colspan="8"></td>
																			</tr>
																					
																				<tr name="row5" id="row5" style="height: 20px; padding: 5px">
																					<td style="white-space: nowrap;width:200px" colspan="3" >
																						<input type="checkbox" id="chkOutputPrinter" name="chkOutputPrinter" tabindex="0"
																							onclick="chkOutputPrinter_Click(); changeTab4Control(); "/>
																						<label
																							for="chkOutputPrinter"
																							class="checkbox"
																							tabindex="-1">
																							Send to printer</label>
																					</td>
																					<%--<td>&nbsp;</td>--%>
																					<td class="ui-state-error-text" style="white-space: nowrap; text-align: left;width: 120px;">Printer location : </td>
																					<td style="width: 350px">
																						<select
																							style="width: 100%"
																							id="cboPrinterName" name="cboPrinterName" class="combo"
																							onchange="changeTab4Control()">
																						</select>
																					</td>
																					<td style="width: 5px;">&nbsp;</td>
																					<td style="width: 5px;">&nbsp;</td>
																					<td style="width: 5px;">&nbsp;</td>
																				</tr>

																				<tr name="row6" id="row6" style="height: 20px; padding: 5px">
																					<td style="white-space: nowrap;width:200px" colspan="3">
																						<input type="checkbox" id="chkSave" name="chkSave" tabindex="0"
																							onclick="chkSave_Click(); changeTab4Control(); " />
																						<label
																							for="chkSave"
																							class="checkbox"
																							tabindex="-1">
																							Save to file</label>
																					</td>
																					<td class="text" style="white-space: nowrap;width: 120px;text-align: left">File name :</td>
																					<td style="width: 350px">
																						<input style="width: 100%" id="  " name="txtSaveFile" disabled="disabled" class="text textdisabled">
																					</td>
																					<td>
																						<input type="button" value="..." id="cmdSaveFile" name="cmdSaveFile" class="btn"
																							onclick="populateMailMergeFileName(); changeTab4Control();"/>
																					</td>
																					<td>
																						<input type="button" value="Clear" id="cmdClearFile" name="cmdClearFile" class="btn"
																							onclick="fileClear()"/>
																					</td>
																					<td>&nbsp;</td>
																				</tr>

																			<tr name="row7" id="row7" style="height: 20px; padding: 5px">
																				<td style="width: 150px; white-space: nowrap">Email Address :</td>
																				<td width="5px"></td>
																				<td>
																					<select id="cboEmail" name="cboEmail" style="width: 400px" class="combo"
																						onchange="changeTab4Control()">
																					</select>
																				</td>
																			</tr>
																			<tr name="row8" id="row8" style="height: 20px; padding: 5px">
																				<td style="width: 150px; white-space: nowrap">Subject :</td>
																				<td width="5px"></td>
																				<td colspan="2">
																					<input id="txtSubject" name="txtSubject" style="width: 400px" maxlength="255" class="text"
																						onkeyup="changeTab4Control()">
																				</td>
																			</tr>
																			<tr name="row9" id="row9" style="height: 20px; padding: 5px">
																				<td nowrap colspan="3">
																					<input type="checkbox" id="chkAttachment" name="chkAttachment" tabindex="0"
																						onclick="chkAttachment_Click(); changeTab4Control(); " />
																					<label
																						for="chkAttachment"
																						class="checkbox"
																						tabindex="-1">
																						Send as attachment
																					</label>
																				</td>
																			</tr>
																			<tr style="height: 20px; padding: 5px" name="row10" id="row10">
																				<td style="width: 150px; white-space: nowrap; text-align: left">Attach as :</td>
																				<td style="width: 5px"></td>
																				<td colspan="2">
																					<input id="txtAttachmentName" name="txtAttachmentName" maxlength="255" style="width: 400px" class="text"
																						onkeyup="changeTab4Control()" />
																				</td>
																			</tr>
																		</table>
																	</td>
																</tr>
																<tr height="100%"></tr>
															</table>
														</td>
													</tr>
												</table>
											</td>
											<td width="20">&nbsp;&nbsp;&nbsp;</td>
										</tr>
										<tr height="20">
											<td colspan="5" class="ui-state-error-text">Note: Options marked in red are unavailable in OpenHR Web.</td>
										</tr>
									</table>
								</div>
							</td>
						</tr>
					</table>
		</table>
		<div id="MailMergeOldScreenButtons">
			<%--these buttons are superceded by ribbon but are 
			here because the ribbon calls their click event--%>
			<input type="button" id="cmdOK" name="cmdOK" value="OK" class="btn"
				onclick="okClick()"/>
			<input type="button" id="cmdCancel" name="cmdCancel" value="Cancel" class="btn"
				onclick="cancelClick()" />
		</div>

		<input type='hidden' id="txtBasePicklistID" name="txtBasePicklistID">
		<input type='hidden' id="txtBaseFilterID" name="txtBaseFilterID">

		<input type='hidden' id="txtParent1ID" name="txtParent1ID">
		<input type='hidden' id="txtParent2ID" name="txtParent2ID">
		<input type='hidden' id="txtParent1FilterID" name="txtParent1FilterID">
		<input type='hidden' id="txtParent1PicklistID" name="txtParent1PicklistID">
		<input type='hidden' id="txtParent2FilterID" name="txtParent2FilterID">
		<input type='hidden' id="txtParent2PicklistID" name="txtParent2PicklistID">

		<input type='hidden' id="txtChildFilterID" name="txtChildFilterID">

		<input type="hidden" id="txtWordFormats" name="txtWordFormats" value="<%=Session("WordFormats")%>">
		<input type="hidden" id="txtWordFormatDefaultIndex" name="txtWordFormatDefaultIndex" value="<%=Session("WordFormatDefaultIndex")%>">
	</form>

	<form id="frmTables">
		<%					
			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
			Dim sErrorDescription = ""
			
			Try
				Dim rstTablesInfo = objDataAccess.GetDataTable("sp_ASRIntGetTablesInfo", CommandType.StoredProcedure)
				
				For Each objRow As DataRow In rstTablesInfo.Rows
					Response.Write("<input type='hidden' id=txtTableName_" & objRow("tableID") & " name=txtTableName_" & objRow("tableID") & " value=""" & objRow("tableName") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtTableType_" & objRow("tableID") & " name=txtTableType_" & objRow("tableID") & " value=" & objRow("tableType") & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtTableChildren_" & objRow("tableID") & " name=txtTableChildren_" & objRow("tableID") & " value=""" & objRow("childrenString") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtTableParents_" & objRow("tableID") & " name=txtTableParents_" & objRow("tableID") & " value=""" & objRow("parentsString") & """>" & vbCrLf)
				Next
											
			Catch ex As Exception
				sErrorDescription = "The tables information could not be retrieved." & vbCrLf & ex.Message

			End Try
			
		%>
	</form>


<form id="frmOriginalDefinition">
	<%
		Dim sErrMsg = ""

		Dim prmErrMsg = New SqlParameter("psErrorMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
		Dim prmName = New SqlParameter("psReportName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
		Dim prmOwner = New SqlParameter("psReportOwner", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
		Dim prmDescription = New SqlParameter("psReportDesc", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
		Dim prmBaseTableID = New SqlParameter("piBaseTableID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
		Dim prmSelection = New SqlParameter("piSelection", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
		Dim prmPicklistID = New SqlParameter("piPicklistID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
		Dim prmPicklistName = New SqlParameter("psPicklistName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
		Dim prmPicklistHidden = New SqlParameter("pfPicklistHidden", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
		Dim prmFilterID = New SqlParameter("piFilterID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
		Dim prmFilterName = New SqlParameter("psFilterName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
		Dim prmFilterHidden = New SqlParameter("pfFilterHidden", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
		Dim prmOutputFormat = New SqlParameter("piOutputFormat", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
		Dim prmOutputSave = New SqlParameter("pfOutputSave", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
		Dim prmOutputFileName = New SqlParameter("psOutputFileName", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
		Dim prmEmailAddrID = New SqlParameter("piEmailAddrID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
		Dim prmEmailSubject = New SqlParameter("psEmailSubject", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
		Dim prmTemplateFileName = New SqlParameter("psTemplateFileName", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
		Dim prmOutputScreen = New SqlParameter("pfOutputScreen", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
		Dim prmEmailAsAttachment = New SqlParameter("pfEmailAsAttachment", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
		Dim prmEmailAttachmentName = New SqlParameter("psEmailAttachmentName", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
		Dim prmSuppressBlanks = New SqlParameter("pfSuppressBlanks", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
		Dim prmPauseBeforeMerge = New SqlParameter("pfPauseBeforeMerge", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
		Dim prmOutputPrinter = New SqlParameter("pfOutputPrinter", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
		Dim prmOutputPrinterName = New SqlParameter("psOutputPrinterName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
		Dim prmDocumentMapID = New SqlParameter("piDocumentMapID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
		Dim prmManualDocManHeader = New SqlParameter("pfManualDocManHeader", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
		Dim prmTimestamp = New SqlParameter("piTimestamp", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
		Dim prmWarningMsg = New SqlParameter("psWarningMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

		If Session("action") <> "new" Then

			Dim rstDefinition = objDataAccess.GetFromSP("sp_ASRIntGetMailMergeDefinition" _
				, New SqlParameter("@piReportID", SqlDbType.Int) With {.Value = CleanNumeric(Session("utilid"))} _
				, New SqlParameter("@psCurrentUser", SqlDbType.VarChar, 255) With {.Value = Session("username")} _
				, New SqlParameter("@psAction", SqlDbType.VarChar, 255) With {.Value = Session("action")} _
				, prmErrMsg, prmName, prmOwner, prmDescription, prmBaseTableID _
				, prmSelection, prmPicklistID, prmPicklistName, prmPicklistHidden _
				, prmFilterID, prmFilterName, prmFilterHidden _
				, prmOutputFormat, prmOutputSave, prmOutputFileName, prmEmailAddrID, prmEmailSubject _
				, prmTemplateFileName, prmOutputScreen, prmEmailAsAttachment, prmEmailAttachmentName, prmSuppressBlanks _
				, prmPauseBeforeMerge, prmOutputPrinter, prmOutputPrinterName _
				, prmDocumentMapID, prmManualDocManHeader, prmTimestamp, prmWarningMsg)

			Dim iHiddenCalcCount = 0
			Dim iCount = 0

			For Each objRow As DataRow In rstDefinition.Rows
				
				iCount = iCount + 1
				If objRow("definitionType").ToString() = "ORDER" Then
					Response.Write("<input type=""hidden"" id=txtReportDefnOrder_" & iCount & " name=txtReportDefnOrder_" & iCount & " value=""" & objRow("definitionString").ToString() & """>" & vbCrLf)
				Else
					Response.Write("<input type=""hidden"" id=txtReportDefnColumn_" & iCount & " name=txtReportDefnColumn_" & iCount & " value=""" & Replace(objRow("definitionString").ToString(), """", "&quot;") & """>" & vbCrLf)
	
					' Check if the report column is a hidden calc.
					If objRow("hidden").ToString() = "Y" Then
						iHiddenCalcCount += 1
					End If
				End If

			Next

			If Len(prmErrMsg.Value) > 0 Then
				sErrMsg = CType(("'" & Session("utilname") & "' " & prmErrMsg.Value), String)
			End If

			Response.Write("<input type=""hidden"" id=txtDefn_Name name=txtDefn_Name value=""" & Replace(prmName.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_Owner name=txtDefn_Owner value=""" & Replace(prmOwner.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_Description name=txtDefn_Description value=""" & Replace(prmDescription.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_BaseTableID name=txtDefn_BaseTableID value=" & prmBaseTableID.Value & ">" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_Selection name=txtDefn_Selection value=" & prmSelection.Value & ">" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_PicklistID name=txtDefn_PicklistID value=" & prmPicklistID.Value & ">" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_PicklistName name=txtDefn_PicklistName value=""" & Replace(prmPicklistName.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_PicklistHidden name=txtDefn_PicklistHidden value=" & LCase(prmPicklistHidden.Value) & ">" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_FilterID name=txtDefn_FilterID value=" & prmFilterID.Value & ">" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_FilterName name=txtDefn_FilterName value=""" & Replace(prmFilterName.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_FilterHidden name=txtDefn_FilterHidden value=" & LCase(prmFilterHidden.Value) & ">" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_HiddenCalcCount name=txtDefn_HiddenCalcCount value=" & iHiddenCalcCount & ">" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_OutputFormat name=txtDefn_OutputFormat value=" & prmOutputFormat.Value & ">" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_OutputSave name=txtDefn_OutputSave value=" & LCase(prmOutputSave.Value) & ">" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_OutputFileName name=txtDefn_OutputFileName value=""" & prmOutputFileName.Value & """>" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_EmailAddrID name=txtDefn_EmailAddrID value=" & prmEmailAddrID.Value & ">" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_EmailSubject name=txtDefn_EmailSubject value=""" & Replace(prmEmailSubject.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_TemplateFileName name=txtDefn_TemplateFileName value=""" & prmTemplateFileName.Value & """>" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_OutputScreen name=txtDefn_OutputScreen value=" & LCase(prmOutputScreen.Value) & ">" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_EmailAsAttachment name=txtDefn_EmailAsAttachment value=" & Replace(LCase(prmEmailAsAttachment.Value.ToString()), """", "&quot;") & ">" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_EmailAttachmentName name=txtDefn_EmailAttachmentName value=""" & prmEmailAttachmentName.Value & """>" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_SuppressBlanks name=txtDefn_SuppressBlanks value=" & LCase(prmSuppressBlanks.Value) & ">" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_PauseBeforeMerge name=txtDefn_PauseBeforeMerge value=" & LCase(prmPauseBeforeMerge.Value) & ">" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_OutputPrinter name=txtDefn_OutputPrinter value=" & LCase(prmOutputPrinter.Value) & ">" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_OutputPrinterName name=txtDefn_OutputPrinterName value=""" & prmOutputPrinterName.Value & """>" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_DocumentMapID name=txtDefn_DocumentMapID value=" & prmDocumentMapID.Value & ">" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_ManualDocManHeader name=txtDefn_ManualDocManHeader value=" & LCase(prmManualDocManHeader.Value) & ">" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_Timestamp name=txtDefn_Timestamp value=" & prmTimestamp.Value & ">" & vbCrLf)
			Response.Write("<input type=""hidden"" id=txtDefn_Warning name=txtDefn_Warning value=""" & Replace(prmWarningMsg.Value.ToString(), """", "&quot;") & """>" & vbCrLf)

		End If
		
		Dim fDocManagement = False
		Dim lngDocumentMapID = 0
		If prmDocumentMapID.value > 0 Then
			fDocManagement = True
			lngDocumentMapID = CInt(prmDocumentMapID.value)
		End If
		
		If fDocManagement = True Then

			Dim rstDocManRecords = objDataAccess.GetFromSP("spASRIntGetDocumentManagementTypes")

			For Each objRow As DataRow In rstDocManRecords.Rows
				If CInt(objRow(0)) = lngDocumentMapID Then
					Response.Write("<input type=""hidden"" id=txtDefn_DocumentMapName name=txtDefn_DocumentMapName value=""" & Replace(objRow(1).ToString(), """", "&quot;") & """>" & vbCrLf)
				End If
			Next
				

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
				, New SqlParameter("piUtilityType", SqlDbType.Int) With {.Value = UtilityType.utlMailMerge} _
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

<form id="frmUseful" name="frmUseful">
	<input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
	<input type="hidden" id="txtLoading" name="txtLoading" value="Y">
	<input type="hidden" id="txtCurrentBaseTableID" name="txtCurrentBaseTableID">
	<input type="hidden" id="txtCurrentChildTableID" name="txtCurrentChildTableID" value="0">
	<input type="hidden" id="txtTablesChanged" name="txtTablesChanged">
	<input type="hidden" id="txtSelectedColumnsLoaded" name="txtSelectedColumnsLoaded" value="0">
	<input type="hidden" id="txtSortLoaded" name="txtSortLoaded" value="0">
	<input type="hidden" id="txtChanged" name="txtChanged" value="0">
	<input type="hidden" id="txtLockGridEvents" name="txtLockGridEvents" value="0">
	<input type="hidden" id="txtUtilID" name="txtUtilID" value='<% =session("utilid")%>'>
	<input type="hidden" id="txtEmailPermission" name="txtEmailPermission">
	<%
				
		Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)

		Dim sParameterValue As String = objDatabase.GetModuleParameter("MODULE_PERSONNEL", "Param_TablePersonnel")
		Response.Write("<input type='hidden' id='txtPersonnelTableID' name='txtPersonnelTableID' value=" & sParameterValue & ">" & vbCrLf)
		
		Response.Write("<input type='hidden' id='txtErrorDescription' name='txtErrorDescription' value="""">" & vbCrLf)
		Response.Write("<input type='hidden' id='txtAction' name='txtAction' value=" & Session("action") & ">" & vbCrLf)

	%>
</form>

<form id="frmValidate" name="frmValidate" target="validate" method="post" action="util_validate_mailmerge">
	<input type="hidden" id="validateBaseFilter" name="validateBaseFilter" value="0">
	<input type="hidden" id="validateBasePicklist" name="validateBasePicklist" value="0">
	<input type="hidden" id="validateCalcs" name="validateCalcs" value=''>
	<input type="hidden" id="validateHiddenGroups" name="validateHiddenGroups" value=''>
	<input type="hidden" id="validateName" name="validateName" value=''>
	<input type="hidden" id="validateTimestamp" name="validateTimestamp" value=''>
	<input type="hidden" id="validateUtilID" name="validateUtilID" value=''>
</form>

<form id="frmSend" name="frmSend" method="post" action="util_def_mailmerge_Submit">
	<input type="hidden" id="txtSend_ID" name="txtSend_ID">
	<input type="hidden" id="txtSend_name" name="txtSend_name">
	<input type="hidden" id="txtSend_description" name="txtSend_description">
	<input type="hidden" id="txtSend_baseTable" name="txtSend_baseTable">
	<input type="hidden" id="txtSend_selection" name="txtSend_selection">
	<input type="hidden" id="txtSend_picklist" name="txtSend_picklist">
	<input type="hidden" id="txtSend_filter" name="txtSend_filter">
	<input type="hidden" id="txtSend_outputformat" name="txtSend_outputformat">
	<input type="hidden" id="txtSend_outputsave" name="txtSend_outputsave">
	<input type="hidden" id="txtSend_outputfilename" name="txtSend_outputfilename">
	<input type="hidden" id="txtSend_emailaddrid" name="txtSend_emailaddrid">
	<input type="hidden" id="txtSend_emailsubject" name="txtSend_emailsubject">
	<input type="hidden" id="txtSend_templatefilename" name="txtSend_templatefilename">
	<input type="hidden" id="txtSend_outputscreen" name="txtSend_outputscreen">
	<input type="hidden" id="txtSend_access" name="txtSend_access">
	<input type="hidden" id="txtSend_userName" name="txtSend_userName">
	<input type="hidden" id="txtSend_emailasattachment" name="txtSend_emailasattachment">
	<input type="hidden" id="txtSend_emailattachmentname" name="txtSend_emailattachmentname">
	<input type="hidden" id="txtSend_suppressblanks" name="txtSend_suppressblanks" value="0">
	<input type="hidden" id="txtSend_pausebeforemerge" name="txtSend_pausebeforemerge" value="0">
	<input type="hidden" id="txtSend_outputprinter" name="txtSend_outputprinter">
	<input type="hidden" id="txtSend_outputprintername" name="txtSend_outputprintername">
	<input type="hidden" id="txtSend_documentmapid" name="txtSend_documentmapid">
	<input type="hidden" id="txtSend_manualdocmanheader" name="txtSend_manualdocmanheader">

	<input type="hidden" id="txtSend_columns" name="txtSend_columns">
	<input type="hidden" id="txtSend_columns2" name="txtSend_columns2">

	<input type="hidden" id="txtSend_reaction" name="txtSend_reaction">

	<input type="hidden" id="txtSend_jobsToHide" name="txtSend_jobsToHide">
	<input type="hidden" id="txtSend_jobsToHideGroups" name="txtSend_jobsToHideGroups">
</form>

<form id="frmRecordSelection" name="frmRecordSelection" target="recordSelection" action="util_recordSelection" method="post">
	<input type="hidden" id="recSelType" name="recSelType">
	<input type="hidden" id="recSelTableID" name="recSelTableID">
	<input type="hidden" id="recSelCurrentID" name="recSelCurrentID">
	<input type="hidden" id="recSelTable" name="recSelTable">
	<input type="hidden" id="recSelDefOwner" name="recSelDefOwner">
	<input type="hidden" id="recSelDefType" name="recSelDefType">
</form>

<form id="frmDocTypeSelection" name="frmDocTypeSelection" target="doctypeSelection" action="util_doctypeSelection" method="post" style="visibility: hidden; display: none">
	<input type="hidden" id="DocTypeSelCurrentID" name="DocTypeSelCurrentID">
</form>

<form id="frmSortOrder" name="frmSortOrder" action="util_sortorderselection" target="sortorderselection" method="post">
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

<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
	<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</form>

<form id="frmSelectionAccess" name="frmSelectionAccess">
	<input type="hidden" id="forcedHidden" name="forcedHidden" value="N">
	<input type="hidden" id="baseHidden" name="baseHidden" value="N">
	<input type="hidden" id="p1Hidden" name="p1Hidden" value="N">
	<input type="hidden" id="p2Hidden" name="p2Hidden" value="N">
	<input type="hidden" id="childHidden" name="childHidden" value="N">
	<input type="hidden" id="calcsHiddenCount" name="calcsHiddenCount" value="0">
</form>

<input type="hidden" id="txtTicker" name="txtTicker" value="0">
<input type="hidden" id="txtLastKeyFind" name="txtLastKeyFind" value="">
</div>

<div style='height: 0;width:0; overflow:hidden;'>
	<input id="cmdGetFilename" name="cmdGetFilename" type="file" />
</div>

<script type="text/javascript">
	util_def_mailmerge_window_onload();
	utilDefMailmergeAddActiveXHandlers();
</script>
