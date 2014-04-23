<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>

<script src="<%: Url.LatestContent("~/Scripts/FormScripts/crosstabdef.js")%>" type="text/javascript"></script>

<div <%=session("BodyTag")%>>

<form id="frmTables" style="visibility: hidden; display: none">
		<%

			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			Dim sErrorDescription = ""

			' Get the table records.
			Try
				Dim rstTablesInfo = objDataAccess.GetFromSP("sp_ASRIntGetCrossTabTablesInfo")
				
				For Each objRow As DataRow In rstTablesInfo.Rows
					Response.Write("<input type='hidden' id=txtTableName_" & objRow("tableID") & " name=txtTableName_" & objRow("tableID") & " value=""" & objRow("tableName") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtTableType_" & objRow("tableID") & " name=txtTableType_" & objRow("tableID") & " value=" & objRow("tableType") & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtTableChildren_" & objRow("tableID") & " name=txtTableChildren_" & objRow("tableID") & " value=""" & objRow("childrenString") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtTableChildrenNames_" & objRow("tableID") & " name=txtTableChildrenNames_" & objRow("tableID") & " value=""" & objRow("childrenNames") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtTableParents_" & objRow("tableID") & " name=txtTableParents_" & objRow("tableID") & " value=""" & objRow("parentsString") & """>" & vbCrLf)
				Next

			Catch ex As Exception
				sErrorDescription = "The tables information could not be retrieved." & vbCrLf & FormatError(ex.Message)

			End Try
				
		%>
	</form>

<form id="frmOriginalDefinition" name="frmOriginalDefinition" style="visibility: hidden; display: none">
		<%
			Dim sErrMsg = ""
			Dim lngHStart = 0
			Dim lngHStop = 0
			Dim lngHStep = 0
			Dim lngVStart = 0
			Dim lngVStop = 0
			Dim lngVStep = 0
			Dim lngPStart = 0
			Dim lngPStop = 0
			Dim lngPStep = 0
			
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
			Dim prmPrintFilter = New SqlParameter("pfPrintFilterHeader", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmHColID = New SqlParameter("HColID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmHStart = New SqlParameter("HStart", SqlDbType.VarChar, 20) With {.Direction = ParameterDirection.Output}
			Dim prmHStop = New SqlParameter("HStop", SqlDbType.VarChar, 20) With {.Direction = ParameterDirection.Output}
			Dim prmHStep = New SqlParameter("HStep", SqlDbType.VarChar, 20) With {.Direction = ParameterDirection.Output}
			Dim prmVColID = New SqlParameter("VColID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmVStart = New SqlParameter("VStart", SqlDbType.VarChar, 20) With {.Direction = ParameterDirection.Output}
			Dim prmVStop = New SqlParameter("VStop", SqlDbType.VarChar, 20) With {.Direction = ParameterDirection.Output}
			Dim prmVStep = New SqlParameter("VStep", SqlDbType.VarChar, 20) With {.Direction = ParameterDirection.Output}
			Dim prmPColID = New SqlParameter("PColID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmPStart = New SqlParameter("PStart", SqlDbType.VarChar, 20) With {.Direction = ParameterDirection.Output}
			Dim prmPStop = New SqlParameter("PStop", SqlDbType.VarChar, 20) With {.Direction = ParameterDirection.Output}
			Dim prmPStep = New SqlParameter("PStep", SqlDbType.VarChar, 20) With {.Direction = ParameterDirection.Output}
			Dim prmIType = New SqlParameter("IType", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmIColID = New SqlParameter("IColID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
			Dim prmPercentage = New SqlParameter("Percentage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmPerPage = New SqlParameter("PerPage", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmSuppress = New SqlParameter("Suppress", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
			Dim prmThousand = New SqlParameter("Thousand", SqlDbType.Bit) With {.Direction = ParameterDirection.Output}
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
																	
			If Session("action") <> "new" Then

				Try
					
					objDataAccess.GetFromSP("sp_ASRIntGetCrossTabDefinition", _
							New SqlParameter("piReportID", SqlDbType.Int) With {.Value = CleanNumeric(Session("utilid"))}, _
							New SqlParameter("psCurrentUser", SqlDbType.VarChar, 255) With {.Value = Session("username")}, _
							New SqlParameter("psAction", SqlDbType.VarChar, 255) With {.Value = Session("action")}, _
							prmErrMsg, prmName, prmOwner, prmDescription, prmBaseTableID, _
							prmAllRecords, prmPicklistID, prmPicklistName, prmPicklistHidden, prmFilterID, prmFilterName, prmFilterHidden, _
							prmPrintFilter, prmHColID, prmHStart, prmHStop, prmHStep, prmVColID, prmVStart, prmVStop, prmVStep, prmPColID, _
							prmPStart, prmPStop, prmPStep, prmIType, prmIColID, prmPercentage, prmPerPage, prmSuppress, prmThousand, _
							prmOutputPreview, prmOutputFormat, prmOutputScreen, prmOutputPrinter, prmOutputPrinterName, prmOutputSave, prmOutputSaveExisting, _
							prmOutputEmail, prmOutputEmailAddr, prmOutputEmailAddrName, prmOutputEmailSubject, prmOutputEmailAttachAs, prmOutputFilename, prmTimestamp)

					Dim iHiddenCalcCount As Integer = 0

					If Len(prmErrMsg.Value) > 0 Then
						sErrMsg = "'" & Session("utilname") & "' " & prmErrMsg.Value
					End If

					lngHStart = CInt(prmHStart.Value)
					lngHStop = CInt(prmHStop.Value)
					lngHStep = CInt(prmHStep.Value)
					lngVStart = CInt(prmVStart.Value)
					lngVStop = CInt(prmVStop.Value)
					lngVStep = CInt(prmVStep.Value)
					lngPStart = CInt(prmPStart.Value)
					lngPStop = CInt(prmPStop.Value)
					lngPStep = CInt(prmPStep.Value)

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
					Response.Write("<input type='hidden' id=txtDefn_FilterHeader name=txtDefn_FilterHeader value=" & prmPrintFilter.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_PrintFilter name=txtDefn_PrintFilter value=" & prmPrintFilter.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_HColID name=txtDefn_HColID value=" & prmHColID.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_HStart name=txtDefn_HStart value=" & prmHStart.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_HStop name=txtDefn_HStop value=" & prmHStop.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_HStep name=txtDefn_HStep value=" & prmHStep.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_VColID name=txtDefn_VColID value=" & prmVColID.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_VStart name=txtDefn_VStart value=" & prmVStart.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_VStop name=txtDefn_VStop value=" & prmVStop.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_VStep name=txtDefn_VStep value=" & prmVStep.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_PColID name=txtDefn_PColID value=" & prmPColID.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_PStart name=txtDefn_PStart value=" & prmPStart.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_PStop name=txtDefn_PStop value=" & prmPStop.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_PStep name=txtDefn_PStep value=" & prmPStep.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_IType name=txtDefn_IType value=" & prmIType.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_IColID name=txtDefn_IColID value=" & prmIColID.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Percentage name=txtDefn_Percentage value=" & prmPercentage.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_PerPage name=txtDefn_PerPage value=" & prmPerPage.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Suppress name=txtDefn_Suppress value=" & prmSuppress.Value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Use1000 name=txtDefn_Use1000 value=" & prmThousand.Value & ">" & vbCrLf)

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
					Response.Write("<input type='hidden' id=txtDefn_HiddenCalcCount name=txtDefn_HiddenCalcCount value=" & iHiddenCalcCount & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=session_action name=session_action value=" & Session("action") & ">" & vbCrLf)
					Response.Write("</form>" & vbCrLf)

				Catch ex As Exception
					sErrMsg = "'" & Session("utilname") & "' cross tab definition could not be read." & vbCrLf & FormatError(ex.Message)
					Session("confirmtext") = sErrMsg
					Session("confirmtitle") = "OpenHR"
					Session("followpage") = "defsel"
					Session("reaction") = "CROSSTABS"
					Response.Clear()
					Response.Redirect("confirmok")

				End Try
					
					
			Else
				Session("childcount") = 0
				Session("hiddenfiltercount") = 0
			End If
		%>
	</form>

<form id="frmDefinition" name="frmDefinition">
		<table valign="top" align="center" cellpadding="5" width="100%" height="100%" cellspacing="0">
			<tr>
				<td colspan="2">
					<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
						<tr height="5">
							<td colspan="3"></td>
						</tr>
						<tr height="10">
							<td width="10"></td>
							<td>
								<input type="button" value="Definition" id="btnTab1" name="btnTab1" class="btn btndisabled" disabled="disabled"
									onclick="display_CrossTab_Page(1)" />
								<input type="button" value="Columns" id="btnTab2" name="btnTab2" class="btn btndisabled" disabled="disabled"
									onclick="display_CrossTab_Page(2)" />
								<input type="button" value="Output" id="btnTab3" name="btnTab3" class="btn btndisabled" disabled="disabled"
									onclick="display_CrossTab_Page(3)" />
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
												<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
													<tr height="10">
														<td width="5">&nbsp;</td>
														<td width="10">Name :</td>
														<td width="5">&nbsp;</td>
														<td>
															<input id="txtName" name="txtName" maxlength="50" style="WIDTH: 100%" class="text"
																onkeyup="changeTab1Control()">
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
														<td colspan="9" height="5"></td>
													</tr>

													<tr height="60">
														<td width="5">&nbsp;</td>
														<td width="10" nowrap valign="top">Description :</td>
														<td width="5">&nbsp;</td>
														<td style="width: 40%; vertical-align: top" rowspan="3">
															<textarea id="txtDescription" name="txtDescription" class="textarea" wrap="VIRTUAL" height="0" maxlength="255"
																style="width: 100%"
																onkeyup="changeTab1Control()">
															</textarea>
														</td>
														<td width="20" nowrap>&nbsp;</td>
														<td width="10" valign="top">Access :</td>
														<td width="5">&nbsp;</td>
														<td style="width: 40%; vertical-align: top" rowspan="2" valign="top">
															<%Html.RenderPartial("Util_Def_CustomReports/grdaccess")%>    
														</td>
														<td width="5">&nbsp;</td>
													</tr>
													<tr height="10">
														<td colspan="7">&nbsp;</td>
													</tr>
													<tr height="10">
														<td colspan="7">&nbsp;</td>
													</tr>
													<tr>
														<td colspan="9"></td>
													</tr>

													<tr height="10">
														<td width="5">&nbsp;</td>
														<td width="100" nowrap valign="top">Base Table :</td>
														<td width="5">&nbsp;</td>
														<td width="40%" valign="top">
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
																	<td colspan="6" height="5"></td>
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
																			Picklist
																		</label>
																	</td>
																	<td width="5">&nbsp;</td>
																	<td>
																		<input id="txtBasePicklist" name="txtBasePicklist" class="text textdisabled" disabled="disabled" style="WIDTH: 85%">
																	
																		<input id="cmdBasePicklist" name="cmdBasePicklist" style="WIDTH: 10%" type="button" disabled="disabled" value="..." class="btn btndisabled"
																			onclick="selectRecordOption('base', 'picklist')" />
																	</td>
                                                                    <td></td>
																</tr>
																<tr>
																	<td colspan="6" height="5"></td>
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
																			Filter
																		</label>
																	</td>
																	<td width="5">&nbsp;</td>
																	<td>
																		<input id="txtBaseFilter" name="txtBaseFilter" disabled="disabled" class="text textdisabled" style="WIDTH: 85%">
																	
																		<input id="cmdBaseFilter" name="cmdBaseFilter" style="WIDTH: 10%" type="button" class="btn btndisabled" disabled="disabled" value="..."
																			onclick="selectRecordOption('base', 'filter')" />
																	</td>
                                                                    <td></td>
																</tr>
															</table>
														<%--</td>--%>
														<td width="5">&nbsp;</td>
													</tr>

													<tr>
														<td colspan="9" height="5">&nbsp;</td>
													</tr>

													<tr>
														<td colspan="5">&nbsp;</td>
														<td colspan="3">
															<input name="chkPrintFilter" id="chkPrintFilter" type="checkbox" disabled="disabled" tabindex="0"
																onclick="changeTab1Control()" />
															<label
																for="chkPrintFilter"
																class="checkbox checkboxdisabled"
																tabindex="-1">
																Display filter or picklist title in the report header
															</label>
														</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr>
														<td colspan="9" height="5">&nbsp;</td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
								</div>
								<!-- Second tab -->
								<div id="div2" style="visibility: hidden; display: none">
									<table width="100%" height="100%" cellspacing="0" cellpadding="5">

										<tr valign="top">
											<td>
												<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
													<tr>
														<td colspan="9" height="5"></td>
													</tr>
													<tr height="10">
														<td width="5">&nbsp;</td>
														<td colspan="4" valign="top"><u>Headings & Breaks</u></td>
														<td width="15%" align="Center">Start</td>
														<td width="5">&nbsp;</td>
														<td width="15%" align="Center">Stop</td>
														<td width="5">&nbsp;</td>
														<td width="15%" align="Center">Increment</td>
														<td>&nbsp;</td>
													</tr>
													<tr>
														<td colspan="9" height="5"></td>
													</tr>
													<tr height="23">
														<td width="5">&nbsp;</td>
														<td width="80" nowrap valign="top">Horizontal :</td>
														<td width="5">&nbsp;</td>
														<td width="20%" valign="top">
															<select id="cboHor" name="cboHor" style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
																onchange="cboHor_Change();changeTab2Control(); ">
															</select>
														</td>
														<td width="15">&nbsp;</td>
														<td>
															<object classid="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"
																codebase="cabs/tinumb6.cab#version=6,0,1,1"
																id="txtHorStart" name="txtHorStart"
																style="height: 24px; WIDTH: 195px"
																onkeyup="changeTab2Control()">
																<param name="DisplayFormat" value="##########0.0000">
																<param name="Format" value="##########0.0000">
																<param name="MaxValue" value="99999999999.9999">
																<param name="Value" value="<%=lngHStart%>">
																<param name="Appearance" value="0">
																<param name="BackColor" value="15988214">
																<param name="ForeColor" value="6697779">
															</object>
														</td>
														<td width="5">&nbsp;</td>
														<td width="15%">
															<object classid="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"
																codebase="cabs/tinumb6.cab#version=6,0,1,1"
																id="txtHorStop"
																name="txtHorStop"
																style="height: 24px; WIDTH: 195px"
																onkeyup="changeTab2Control()">
																<param name="DisplayFormat" value="##########0.0000">
																<param name="Format" value="##########0.0000">
																<param name="MaxValue" value="99999999999.9999">
																<param name="Value" value="<%=lngHStop%>">
																<param name="Appearance" value="0">
																<param name="BackColor" value="15988214">
																<param name="ForeColor" value="6697779">
															</object>
														</td>
														<td width="5">&nbsp;</td>
														<td width="15%">
															<object classid="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"
																codebase="cabs/tinumb6.cab#version=6,0,1,1"
																id="txtHorStep"
																name="txtHorStep"
																style="height: 24px; WIDTH: 195px"
																onkeyup="changeTab2Control()">
																<param name="DisplayFormat" value="##########0.0000">
																<param name="Format" value="##########0.0000">
																<param name="MaxValue" value="99999999999.9999">
																<param name="Value" value="<%=lngHStep%>">
																<param name="Appearance" value="0">
																<param name="BackColor" value="15988214">
																<param name="ForeColor" value="6697779">
															</object>
														</td>

														<td>&nbsp;</td>
													</tr>
													<tr>
														<td colspan="9" height="5"></td>
													</tr>
													<tr height="23">
														<td width="5">&nbsp;</td>
														<td width="80" nowrap valign="top">Vertical :</td>
														<td width="5">&nbsp;</td>
														<td width="20%" valign="top">
															<select id="cboVer" name="cboVer" style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
																onchange="cboVer_Change();changeTab2Control(); ">
															</select>
														</td>
														<td width="15">&nbsp;</td>
														<td width="15%">
															<object classid="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"
																codebase="cabs/tinumb6.cab#version=6,0,1,1"
																id="txtVerStart"
																name="txtVerStart"
																style="height: 24px; WIDTH: 195px"
																onkeyup="changeTab2Control()">
																<param name="DisplayFormat" value="##########0.0000">
																<param name="Format" value="##########0.0000">
																<param name="MaxValue" value="99999999999.9999">
																<param name="Value" value="<%=lngVStart%>">
																<param name="Appearance" value="0">
																<param name="BackColor" value="15988214">
																<param name="ForeColor" value="6697779">
															</object>
														</td>
														<td width="5">&nbsp;</td>
														<td width="15%">
															<object classid="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"
																codebase="cabs/tinumb6.cab#version=6,0,1,1"
																id="txtVerStop"
																name="txtVerStop"
																style="height: 24px; WIDTH: 195px"
																onkeyup="changeTab2Control()">
																<param name="DisplayFormat" value="##########0.0000">
																<param name="Format" value="##########0.0000">
																<param name="MaxValue" value="99999999999.9999">
																<param name="Value" value="<%=lngVStop%>">
																<param name="Appearance" value="0">
																<param name="BackColor" value="15988214">
																<param name="ForeColor" value="6697779">
															</object>
														</td>
														<td width="5">&nbsp;</td>
														<td width="15%">
															<object classid="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"
																codebase="cabs/tinumb6.cab#version=6,0,1,1"
																id="txtVerStep"
																name="txtVerStep"
																style="height: 24px; WIDTH: 195px"
																onkeyup="changeTab2Control()">
																<param name="DisplayFormat" value="##########0.0000">
																<param name="Format" value="##########0.0000">
																<param name="MaxValue" value="99999999999.9999">
																<param name="Value" value="<%=lngVStep%>">
																<param name="Appearance" value="0">
																<param name="BackColor" value="15988214">
																<param name="ForeColor" value="6697779">
															</object>
														</td>

														<td>&nbsp;</td>
													</tr>
													<tr>
														<td colspan="9" height="5"></td>
													</tr>
													<tr height="23">
														<td width="5">&nbsp;</td>
														<td width="100" nowrap valign="top">Page Break :</td>
														<td width="5">&nbsp;</td>
														<td width="20%" valign="top">
															<select id="cboPgb" name="cboPgb" style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
																onchange="cboPgb_Change();changeTab2Control(); ">
															</select>
														</td>
														<td width="15">&nbsp;</td>
														<td width="15%">
															<object classid="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"
																codebase="cabs/tinumb6.cab#version=6,0,1,1" id="txtPgbStart" name="txtPgbStart"
																style="LEFT: 0px; TOP: 0px; height: 24px; WIDTH: 195px"
																onkeyup="changeTab2Control()">
																<param name="DisplayFormat" value="##########0.0000">
																<param name="Format" value="##########0.0000">
																<param name="MaxValue" value="99999999999.9999">
																<param name="Value" value="<%=lngPStart%>">
																<param name="Appearance" value="0">
																<param name="BackColor" value="15988214">
																<param name="ForeColor" value="6697779">
															</object>
														</td>
														<td width="5">&nbsp;</td>
														<td width="15%">
															<object classid="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"
																codebase="cabs/tinumb6.cab#version=6,0,1,1" id="txtPgbStop"
																name="txtPgbStop" style="LEFT: 0px; TOP: 0px; height: 24px; WIDTH: 195px"
																onkeyup="changeTab2Control()">
																<param name="DisplayFormat" value="##########0.0000">
																<param name="Format" value="##########0.0000">
																<param name="MaxValue" value="99999999999.9999">
																<param name="Value" value="<%=lngPStop%>">
																<param name="Appearance" value="0">
																<param name="BackColor" value="15988214">
																<param name="ForeColor" value="6697779">
															</object>
														</td>
														<td width="5">&nbsp;</td>
														<td width="15%">
															<object classid="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"
																codebase="cabs/tinumb6.cab#version=6,0,1,1"
																id="txtPgbStep" name="txtPgbStep"
																style="LEFT: 0px; TOP: 0px; height: 24px; WIDTH: 195px"
																onkeyup="changeTab2Control()">
																<param name="DisplayFormat" value="##########0.0000">
																<param name="Format" value="##########0.0000">
																<param name="MaxValue" value="99999999999.9999">
																<param name="Value" value="<%=lngPStep%>">
																<param name="Appearance" value="0">
																<param name="BackColor" value="15988214">
																<param name="ForeColor" value="6697779">
															</object>
														</td>
														<td>&nbsp;</td>
													</tr>
													<tr height="20">
														<td colspan="11"></td>
													</tr>
													<tr height="10">
														<td width="5">&nbsp;</td>
														<td style="width: 80px; white-space: nowrap; vertical-align: top; text-decoration: underline" colspan="4">Intersection</td>
													</tr>
													<tr>
														<td colspan="12" height="5"></td>
													</tr>
												</table>
												<table style="width: 60%;" class="invisible" cellspacing="0" cellpadding="0">
													<tr>
														<td colspan="2">
															<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																<tr height="10">
																	<td width="5">&nbsp;</td>
																	<td style="width:80px;white-space: nowrap; vertical-align: top">Column :</td>
																	<td width="5">&nbsp;</td>
																	<td style="vertical-align: top" >
																		<select id="cboInt" name="cboInt" class="combo combodisabled" disabled="disabled"
																			style="width: 100%"
																			onchange="cboInt_Change();changeTab2Control(); ">
																		</select>
																	</td>
																</tr>
																<tr>
																	<td colspan="9" height="5"></td>
																</tr>
																<tr height="10">
																	<td width="5">&nbsp;</td>
																	<td width="80" nowrap valign="top">Type :</td>
																	<td width="5">&nbsp;</td>
																	<td>
																		<select id="cboIntType" name="cboIntType"
																			class="combo combodisabled" disabled="disabled"
																			style="width: 100%"
																			onchange="changeTab2Control()">
																			<option value="1">Average</option>
																			<option value="0" selected>Count</option>
																			<option value="2">Maximum</option>
																			<option value="3">Minimum</option>
																			<option value="4">Total</option>
																		</select>
																	</td>
																</tr>
															</table>
														</td>

														<td style="width: 15px;">&nbsp;</td>

														<td>
															<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																<tr>
																	<td>
																		<input type="checkbox" id="chkPercentage" name="chkPercentage" tabindex="0"
																			onclick="changeTab2Control()" />
																		<label
																			for="chkPercentage"
																			class="checkbox"
																			tabindex="-1">
																			Percentage of Type
																		</label>
																	</td>
																</tr>
																<tr>
																	<td height="5"></td>
																</tr>
																<tr>
																	<td>
																		<input type="checkbox" id="chkPerPage" name="chkPerPage" tabindex="0"
																			onclick="changeTab2Control()" />
																		<label
																			for="chkPerPage"
																			class="checkbox"
																			tabindex="-1">
																			Percentage of Page
																		</label>
																	</td>
																</tr>
																<tr>
																	<td height="5"></td>
																</tr>
																<tr>
																	<td>
																		<input type="checkbox" id="chkSuppress" name="chkSuppress" tabindex="0"
																			onclick="changeTab2Control()" />
																		<label
																			for="chkSuppress"
																			class="checkbox"
																			tabindex="-1">
																			Suppress Zeros
																		</label>
																	</td>
																</tr>
																<tr>
																	<td height="5"></td>
																</tr>
																<tr>
																	<td>
																		<input type="checkbox" id="chkUse1000" name="chkUse1000" tabindex="0"
																			onclick="changeTab2Control()" />
																		<label
																			for="chkUse1000"
																			class="checkbox"
																			tabindex="-1">
																			Use 1000 Separators (,)
																		</label>
																	</td>
																</tr>
															</table>
														</td>

													</tr>
													<tr>
														<td colspan="9" height="5"></td>
													</tr>
												</table>
									</table>
								</div>
								<!-- OUTPUT OPTIONS -->
								<div id="div3" style="visibility: hidden; display: none">
									<table width="100%" height="100%" cellspacing="0" cellpadding="5">
										<tr valign="top">
											<td>
												<table width="100%" class="invisible" cellspacing="10" cellpadding="0">
													<tr>
														<td valign="top" rowspan="2" width="15%" height="100%">
															<table cellspacing="0" cellpadding="4" width="100%" height="100%">
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
																				<td align="left" nowrap>
																					<label
																						tabindex="-1"
																						for="optOutputFormat0"
																						class="radio">
																						Data Only
																					</label>
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
																						class="radio ui-state-error-text">
																						CSV File
																					</label>
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
																						class="radio ui-state-error-text">
																						HTML Document
																					</label>
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
																						class="radio ui-state-error-text">
																						Word Document
																					</label>
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
																						Excel Worksheet
																					</label>
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
																						Excel Chart
																					</label>
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
																						Excel Pivot Table
																					</label>
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
																						onclick="changeTab3Control();" />
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
																						onclick="changeTab3Control();" />
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
																						onclick="changeTab3Control();"
																						onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																					<label
																						for="chkDestination1"
																						class="checkbox checkboxdisabled"
																						tabindex="-1"
																						onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																						onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																						onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																						onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																						Send to printer 
																					</label>
																				</td>
																				<td width="30" nowrap>&nbsp;</td>
																				<td style="width: 15%">&nbsp;</td>
																				<td style="width: 25%; text-align: left; white-space: nowrap" class="ui-state-error-text">Printer location : </td>
																				<td style="width: 100%">
																					<select id="cboPrinterName" name="cboPrinterName" class="combo"
																						style="width: 100%"
																						onchange="changeTab3Control()">
																					</select>
																				</td>
																				<td width="30" nowrap>&nbsp;</td>
																				<td width="30" nowrap>&nbsp;</td>
																			</tr>

																			<tr style="height: 20px">
																				<td></td>
																				<td align="left" nowrap>
																					<input name="chkDestination2" id="chkDestination2" type="checkbox" disabled="disabled" tabindex="0"
																						onclick="changeTab3Control();" />
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
																						style="width: 85%;"
																						class="text textdisabled" disabled="disabled" tabindex="-1">

																					<input id="cmdFilename" name="cmdFilename" class="btn" type="button" value='...' disabled="disabled"
																						onclick="populateFileName(frmDefinition); changeTab3Control();" style="width: 12%;" />
																				</td>
                                                                                <td></td>
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
																						onchange="changeTab3Control()">
																					</select>
																				</td>
																				<td></td>
																				<td></td>
																			</tr>

																			<tr style="height: 20px">
																				<td></td>
																				<td style="white-space: nowrap;text-align: left">
																					<input name="chkDestination3" id="chkDestination3" type="checkbox" disabled="disabled" tabindex="0"
																						onclick="changeTab3Control();"/>
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
																						style="width: 85%;"
																						class="text textdisabled" disabled="disabled" tabindex="-1">
																					
																				
																					<input id="cmdEmailGroup" name="cmdEmailGroup" 
																						type="button" 
																						value='...' 
																						disabled="disabled" 
																						class="btn"
																						onclick="selectEmailGroup(); changeTab3Control();" style="width:12%;"/>
                                                                                    <input id="txtEmailGroupID" name="txtEmailGroupID" type="hidden" 
																						class="text textdisabled" disabled="disabled" tabindex="-1" style="width:0px;"/>
																				</td>
                                                                                <td></td>
																				<td></td>
																			</tr>
																			<tr>
																				<td></td>
																				<td></td>
																				<td></td>
																				<td></td>
																				<td style="white-space: nowrap;">
																					<label for="txtEmailSubject"
																						tabindex="-1"
																						onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																						onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																						onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																						onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
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
																						tabindex="-1"
																						onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}" 
																						onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
																						onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}" 
																						onfocus="try{checkboxLabel_onFocus(this);}catch(e){}" 
																						onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
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
										<tr height="20">
											<td colspan="5" class="ui-state-error-text">Note: In OpenHR Web Output Format is restricted to Excel. Existing files will be overwritten.</td>
										</tr>
									</table>
								</div>
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr height="5">
				<td colspan="3"></td>
			</tr>
			<tr height="10">
				<td width="10"></td>
				<td>
					<table style="width: 100%" class="invisible" cellspacing="0" cellpadding="0">
						<tr>
							<td width="80">
								<input type="button" id="cmdOK" name="cmdOK" value="OK" style="width: 100%" class="btn"
									onclick="okClick()" />
							</td>
							<td width="10"></td>
							<td width="80">
								<input type="button" id="cmdCancel" name="cmdCancel" value="Cancel" style="width: 100%" class="btn"
									onclick="cancelClick()" />
							</td>
							<td>&nbsp;</td>
						</tr>
					</table>
				</td>
				<td width="10"></td>
			</tr>
		</table>

		<input type='hidden' id="txtBasePicklistID" name="txtBasePicklistID">
		<input type='hidden' id="txtBaseFilterID" name="txtBaseFilterID">
		<input type='hidden' id="txtDatabase" name="txtDatabase" value="<%=session("Database")%>">
		<input type='hidden' id="txtWordVer" name="txtWordVer" value="<%=Session("WordVer")%>">
		<input type='hidden' id="txtExcelVer" name="txtExcelVer" value="<%=Session("ExcelVer")%>">
		<input type='hidden' id="txtWordFormats" name="txtWordFormats" value="<%=Session("WordFormats")%>">
		<input type='hidden' id="txtExcelFormats" name="txtExcelFormats" value="<%=Session("ExcelFormats")%>">
		<input type='hidden' id="txtWordFormatDefaultIndex" name="txtWordFormatDefaultIndex" value="<%=Session("WordFormatDefaultIndex")%>">
		<input type='hidden' id="txtExcelFormatDefaultIndex" name="txtExcelFormatDefaultIndex" value="<%=Session("ExcelFormatDefaultIndex")%>">
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
				, New SqlParameter("piUtilityType", SqlDbType.Int) With {.Value = UtilityType.utlCrossTab} _
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
	<input type="hidden" id="txtCurrentHColID" name="txtCurrentHColID">
	<input type="hidden" id="txtCurrentVColID" name="txtCurrentVColID">
	<input type="hidden" id="txtCurrentPColID" name="txtCurrentPColID">
	<input type="hidden" id="txtCurrentIColID" name="txtCurrentIColID">
	<input type="hidden" id="txtTablesChanged" name="txtTablesChanged">
	<input type="hidden" id="txtSelectedColumnsLoaded" name="txtSelectedColumnsLoaded" value="0">
	<input type="hidden" id="txtSortLoaded" name="txtSortLoaded" value="0">
	<input type="hidden" id="txtSecondTabShown" name="txtSecondTabShown" value="0">
	<input type="hidden" id="txtRepetitionLoaded" name="txtRepetitionLoaded" value="0">
	<input type="hidden" id="txtChanged" name="txtChanged" value="0">
	<input type="hidden" id="txtUtilID" name="txtUtilID" value='<%=session("utilid")%>'>
	<%		
		Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)

		Dim sParameterValue As String = objDatabase.GetModuleParameter("MODULE_PERSONNEL", "Param_TablePersonnel")
		Response.Write("<input type='hidden' id='txtPersonnelTableID' name='txtPersonnelTableID' value=" & sParameterValue & ">" & vbCrLf)
		
		Response.Write("<input type='hidden' id='txtErrorDescription' name='txtErrorDescription' value="""">" & vbCrLf)
		Response.Write("<input type='hidden' id='txtAction' name='txtAction' value=" & Session("action") & ">" & vbCrLf)

	%>
</form>

<form id="frmValidate" name="frmValidate" target="validate" method="post" action="util_validate_crosstab" style="visibility: hidden; display: none">
	<input type="hidden" id="validateBaseFilter" name="validateBaseFilter" value="0">
	<input type="hidden" id="validateBasePicklist" name="validateBasePicklist" value="0">
	<input type="hidden" id="validateEmailGroup" name="validateEmailGroup" value="0">
	<input type="hidden" id="validateCalcs" name="validateCalcs" value=''>
	<input type="hidden" id="validateHiddenGroups" name="validateHiddenGroups" value=''>
	<input type="hidden" id="validateName" name="validateName" value=''>
	<input type="hidden" id="validateTimestamp" name="validateTimestamp" value=''>
	<input type="hidden" id="validateUtilID" name="validateUtilID" value=''>
</form>

<form id="frmSend" name="frmSend" method="post" action="util_def_crosstabs_submit" style="visibility: hidden; display: none">
	<input type="hidden" id="txtSend_ID" name="txtSend_ID" value="0">
	<input type="hidden" id="txtSend_name" name="txtSend_name" value=''>
	<input type="hidden" id="txtSend_description" name="txtSend_description" value=''>
	<input type="hidden" id="txtSend_baseTable" name="txtSend_baseTable" value="0">
	<input type="hidden" id="txtSend_allRecords" name="txtSend_allRecords" value="0">
	<input type="hidden" id="txtSend_picklist" name="txtSend_picklist" value="0">
	<input type="hidden" id="txtSend_filter" name="txtSend_filter" value="0">
	<input type="hidden" id="txtSend_PrintFilter" name="txtSend_PrintFilter" value="0">
	<input type="hidden" id="txtSend_access" name="txtSend_access" value=''>
	<input type="hidden" id="txtSend_userName" name="txtSend_userName" value=''>

	<input type="hidden" id="txtSend_HColID" name="txtSend_HColID" value="0">
	<input type="hidden" id="txtSend_HStart" name="txtSend_HStart" value=''>
	<input type="hidden" id="txtSend_HStop" name="txtSend_HStop" value=''>
	<input type="hidden" id="txtSend_HStep" name="txtSend_HStep" value=''>
	<input type="hidden" id="txtSend_VColID" name="txtSend_VColID" value="0">
	<input type="hidden" id="txtSend_VStart" name="txtSend_VStart" value=''>
	<input type="hidden" id="txtSend_VStop" name="txtSend_VStop" value=''>
	<input type="hidden" id="txtSend_VStep" name="txtSend_VStep" value=''>
	<input type="hidden" id="txtSend_PColID" name="txtSend_PColID" value="0">
	<input type="hidden" id="txtSend_PStart" name="txtSend_PStart" value=''>
	<input type="hidden" id="txtSend_PStop" name="txtSend_PStop" value=''>
	<input type="hidden" id="txtSend_PStep" name="txtSend_PStep" value=''>
	<input type="hidden" id="txtSend_IType" name="txtSend_IType" value="0">
	<input type="hidden" id="txtSend_IColID" name="txtSend_IColID" value="0">
	<input type="hidden" id="txtSend_Percentage" name="txtSend_Percentage" value="0">
	<input type="hidden" id="txtSend_PerPage" name="txtSend_PerPage" value="0">
	<input type="hidden" id="txtSend_Suppress" name="txtSend_Suppress" value="0">
	<input type="hidden" id="txtSend_Use1000Separator" name="txtSend_Use1000Separator" value="0">

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

	<input type="hidden" id="txtSend_reaction" name="txtSend_reaction">

	<input type="hidden" id="txtSend_jobsToHide" name="txtSend_jobsToHide">
	<input type="hidden" id="txtSend_jobsToHideGroups" name="txtSend_jobsToHideGroups">
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

<form id="frmSelectionAccess" name="frmSelectionAccess" style="visibility: hidden; display: none">
	<input type="hidden" id="forcedHidden" name="forcedHidden" value="N">
	<input type="hidden" id="baseHidden" name="baseHidden" value="N">

	<!-- need the count of hidden child filter access info -->
	<input type="hidden" id="childHidden" name="childHidden" value="N">
	<input type="hidden" id="calcsHiddenCount" name="calcsHiddenCount" value="0">
</form>

<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
	<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</form>

	<input type='hidden' id="txtTicker" name="txtTicker" value="0">
	<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">
</div>

<div style='height: 0;width:0; overflow:hidden;'>
	<input id="cmdGetFilename" name="cmdGetFilename" type="file" />
</div>

<script type="text/javascript">
	util_def_crosstabs_window_onload();
	util_def_crosstabs_addhandlers();
</script>
