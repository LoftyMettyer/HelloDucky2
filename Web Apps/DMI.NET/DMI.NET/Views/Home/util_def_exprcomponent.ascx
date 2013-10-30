<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="ADODB" %>

<script src="<%: Url.Content("~/bundles/utilities_expressions")%>" type="text/javascript"></script>


<form action="" method="POST" id="frmMainForm" name="frmMainForm">
	<%
		Dim cmdParameter As Command
		Dim prmFunctionID As ADODB.Parameter
		Dim prmParameterIndex As ADODB.Parameter
		Dim prmPassByType As ADODB.Parameter
		
		Dim iPassBy As Integer
		Dim sErrMsg As String
		
		
		iPassBy = 1
		If (Len(sErrMsg) = 0) And (Session("optionFunctionID") > 0) Then
			cmdParameter = New Command
			cmdParameter.CommandText = "spASRIntGetParameterPassByType"
			cmdParameter.CommandType = CommandTypeEnum.adCmdStoredProc
			cmdParameter.ActiveConnection = Session("databaseConnection")

			prmFunctionID = cmdParameter.CreateParameter("functionID", 3, 1) ' 3=integer, 1=input
			cmdParameter.Parameters.Append(prmFunctionID)
			prmFunctionID.value = CleanNumeric(CLng(Session("optionFunctionID")))

			prmParameterIndex = cmdParameter.CreateParameter("parameterIndex", 3, 1) ' 3=integer, 1=input
			cmdParameter.Parameters.Append(prmParameterIndex)
			prmParameterIndex.value = CleanNumeric(CLng(Session("optionParameterIndex")))

			prmPassByType = cmdParameter.CreateParameter("passByType", 3, 2) ' 3=integer, 2=output
			cmdParameter.Parameters.Append(prmPassByType)

			Err.Clear()
			cmdParameter.Execute()
			If (Err.Number <> 0) Then
				sErrMsg = "Error checking parameter pass-by type." & vbCrLf & FormatError(Err.Description)
			Else
				iPassBy = cmdParameter.Parameters("passByType").Value
			End If

			' Release the ADO command object.
			cmdParameter = Nothing
		End If
		Response.Write("<INPUT type='hidden' id=txtPassByType name=txtPassByType value=" & iPassBy & ">" & vbCrLf)
	%>

	<table align="center" cellpadding="5" cellspacing="0" width="100%" height="100%">
		<tr>
			<td>
				<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
					<%--					<tr height="5">
						<td height="5" colspan="5"></td>
					</tr>--%>

					<tr>
						<%--<td style="width: 10px; vertical-align: top">&nbsp;&nbsp;</td>--%>

						<td style="width: 10px; vertical-align: top">
							<table height="100%" width="100%" cellspacing="0" cellpadding="0">
								<tr>
									<td valign="top">
										<table border="0" cellspacing="0" cellpadding="0">
											<tr height="5">
												<td colspan="5">&nbsp;</td>
											</tr>

											<tr height="10" style="height: 10px; display: block; padding-bottom: 10px">
												<td width="5">&nbsp;</td>
												<td width="5"><strong>Type</strong></td>
												<td colspan="3"></td>
											</tr>

											<tr height="10">
												<td colspan="5"></td>
											</tr>

											<tr height="10" style="height: 10px; display: block; padding-bottom: 10px">
												<td width="5">&nbsp;</td>
												<td width="5">
													<input id="optType_Field" name="optType" type="radio" selected
														onclick="changeType(1)"/>
												</td>
												<td width="5">&nbsp;</td>
												<td nowrap>
													<label
														tabindex="-1"
														for="optType_Field"
														class="radio">
														Field</label>
												</td>
												<td width="5">&nbsp;&nbsp;</td>
											</tr>

											<tr height="5">
												<td colspan="5"></td>
											</tr>

											<tr height="10" style="height: 10px; display: block; padding-bottom: 10px">
												<td width="5">&nbsp;</td>
												<td width="5">
													<input id="optType_Operator" name="optType" type="radio"
														<%
														If iPassBy = 2 Then
															Response.Write("disabled")
														End If
														%>
														onclick="changeType(5)" />
												</td>
												<td width="5">&nbsp;</td>
												<td nowrap>
													<label
														tabindex="-1"
														for="optType_Operator"
														class="radio">
														Operator</label>
												</td>
												<td width="5">&nbsp;</td>
											</tr>

											<tr height="5">
												<td colspan="5"></td>
											</tr>

											<tr height="10" style="height: 10px; display: block; padding-bottom: 10px">
												<td width="5">&nbsp;</td>
												<td width="5">
													<input id="optType_Function" name="optType" type="radio"
														<%
														If iPassBy = 2 Then
															Response.Write("disabled")
														End If
														%>
														onclick="changeType(2)" />
												</td>
												<td width="5">&nbsp;</td>
												<td nowrap>
													<label
														tabindex="-1"
														for="optType_Function"
														class="radio">
														Function</label>
												</td>
												<td width="5">&nbsp;</td>
											</tr>

											<tr height="5">
												<td colspan="5"></td>
											</tr>

											<tr height="10" style="height: 10px; display: block; padding-bottom: 10px">
												<td width="5">&nbsp;</td>
												<td width="5">
													<input id="optType_Value" name="optType" type="radio" <%
														If iPassBy = 2 Then
															Response.Write("disabled")
														End If
														%>
														onclick="changeType(4)" />
												</td>
												<td width="5">&nbsp;</td>
												<td nowrap>
													<label
														tabindex="-1"
														for="optType_Value"
														class="radio">
														Value</label>
												</td>
												<td width="5">&nbsp;</td>
											</tr>

											<tr height="5">
												<td colspan="5"></td>
											</tr>

											<tr height="10" style="height: 10px; display: block; padding-bottom: 10px">
												<td width="5">&nbsp;</td>
												<td width="5">
													<input id="optType_LookupTableValue" name="optType" type="radio" <% 
														If iPassBy = 2 Then
															Response.Write("disabled")
														End If
																										%>
														onclick="changeType(6)" />
												</td>
												<td width="5">&nbsp;</td>
												<td nowrap>
													<label
														tabindex="-1"
														for="optType_LookupTableValue"
														class="radio">
														Lookup Table Value</label>
												</td>
												<td width="5">&nbsp;</td>
											</tr>

											<tr height="5">
												<td colspan="5"></td>
											</tr>

											<tr height="10" style="visibility: hidden; display: none; padding-bottom: 10px" id="trType_PVal">
												<td width="5">&nbsp;</td>
												<td width="5">
													<input id="optType_PromptedValue" name="optType" type="radio" <%If iPassBy = 2 Then
															Response.Write("disabled")
																	End If%>
														onclick="changeType(7)" />
												</td>
												<td width="5">&nbsp;</td>
												<td nowrap>
													<label
														tabindex="-1"
														for="optType_PromptedValue"
														class="radio">
														Prompted Value</label>
												</td>
												<td width="5">&nbsp;</td>
											</tr>

											<tr height="5" style="visibility: hidden; display: none" id="trType_PVal2">
												<td colspan="5"></td>
											</tr>

											<tr height="10" style="visibility: hidden; display: none; padding-bottom: 10px" id="trType_Calc">
												<td width="5">&nbsp;</td>
												<td width="5">
													<input id="optType_Calculation" name="optType" type="radio" <%	 If iPassBy = 2 Then
															Response.Write("disabled")
																	End If%>
														onclick="changeType(3)" />
												</td>
												<td width="5">&nbsp;</td>
												<td nowrap>
													<label
														tabindex="-1"
														for="optType_Calculation"
														class="radio"
														onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
														onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}" >Calculation</label>
												</td>
												<td width="5">&nbsp;</td>
											</tr>

											<tr height="5" style="visibility: hidden; display: none" id="trType_Calc2">
												<td colspan="5"></td>
											</tr>

											<tr height="10" style="visibility: hidden; display: none; padding-bottom: 10px" id="trType_Filter">
												<td width="5">&nbsp;</td>
												<td width="5">
													<input id="optType_Filter" name="optType" type="radio" <%
														If iPassBy = 2 Then
															Response.Write("disabled")
														End If

														%>
														onclick="changeType(10)" />
												</td>
												<td width="5">&nbsp;</td>
												<td nowrap>
													<label
														tabindex="-1"
														for="optType_Filter"
														class="radio">
														Filter</label>
												</td>
												<td width="5">&nbsp;</td>
											</tr>

											<tr>
												<td colspan="5"></td>
											</tr>
										</table>
									</td>
								</tr>
							</table>
						</td>

						<td style="vertical-align: top; width: 10px">&nbsp;&nbsp;</td>

						<td style="vertical-align: top">
							<table height="100%" width="100%" cellspacing="0" cellpadding="0">
								<tr height="100%">
									<td valign="top">
										<div id="divField">
											<table class="invisible" cellspacing="0" cellpadding="0">
												<tr height="10">
													<td colspan="6">&nbsp;&nbsp;</td>
												</tr>

												<tr height="10">
													<td width="10">&nbsp;</td>
													<td colspan="4"><strong>Field</strong></td>
													<td width="10">&nbsp;</td>
												</tr>

												<tr height="10">
													<td colspan="6"></td>
												</tr>
												<%If iPassBy = 1 Then%>
												<tr height="10">
													<td width="10">&nbsp;</td>
													<td colspan="4">
														<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
															<tr>
																<td>
																	<input id="optField_Field" name="optField" type="radio" selected
																		onclick="field_refreshTable()" />
																</td>
																<td width="5">&nbsp;</td>
																<td nowrap>
																	<label
																		tabindex="-1"
																		for="optField_Field"
																		class="radio">
																		Field</label>
																</td>
																<td width="20">&nbsp;&nbsp;&nbsp;</td>
																<td>
																	<input id="optField_Count" name="optField" type="radio" selected
																		onclick="field_refreshTable()" />
																</td>
																<td width="5">&nbsp;</td>
																<td nowrap>
																	<label
																		tabindex="-1"
																		for="optField_Count"
																		class="radio">
																		Count</label>
																</td>
																<td width="20">&nbsp;&nbsp;&nbsp;</td>
																<td>
																	<input id="optField_Total" name="optField" type="radio" selected
																		onclick="field_refreshTable()" />
																</td>
																<td width="5">&nbsp;</td>
																<td nowrap>
																	<label
																		tabindex="-1"
																		for="optField_Total"
																		class="radio">
																		Total</label>
																</td>
																<td width="100%"></td>
															</tr>
														</table>
													</td>
													<td width="10">&nbsp;</td>
												</tr>

												<tr height="10">
													<td colspan="6">&nbsp;&nbsp;</td>
												</tr>
												<%End If%>

												<tr height="10">
													<td width="20">&nbsp;&nbsp;</td>
													<td width="10" nowrap>Table :</td>
													<td width="20">&nbsp;&nbsp;</td>
													<td width="50%">
														<select id="cboFieldTable" name="cboFieldTable" class="combo" style="width: 100%"
															onchange="field_changeTable()">
														</select>
													</td>
													<td width="50%">&nbsp;</td>
													<td width="10">&nbsp;</td>
												</tr>

												<tr height="5">
													<td colspan="6"></td>
												</tr>

												<tr height="10">
													<td width="10">&nbsp;</td>
													<td width="10" nowrap>Column :</td>
													<td width="20">&nbsp;&nbsp;</td>
													<td width="50%">
														<select id="cboFieldColumn" name="cboFieldColumn" class="combo" style="width: 100%">
														</select>
														<select id="cboFieldDummyColumn" name="cboFieldDummyColumn" class="combo combodisabled" style="width: 100%; visibility: hidden; display: none" disabled="disabled">
														</select>
													</td>
													<td width="50%">&nbsp;</td>
													<td width="10">&nbsp;&nbsp;</td>
												</tr>

												<%If iPassBy = 1 Then%>
												<tr height="5">
													<td colspan="6">&nbsp;</td>
												</tr>

												<tr height="10">
													<td width="10">&nbsp;</td>
													<td colspan="4">
														<table width="100%" height="100%">
															<tr>
																<td>
																	<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																		<tr height="10">
																			<td colspan="6"></td>
																		</tr>

																		<tr height="10">
																			<%--<td width="10">&nbsp;</td>--%>
																			<td colspan="4"><strong>Child Field Options</strong></td>
																			<td width="10">&nbsp;</td>
																		</tr>

																		<tr height="10">
																			<td colspan="6"></td>
																		</tr>

																		<tr height="10">
																			<%--<td width="10">&nbsp;</td>--%>
																			<td colspan="4">
																				<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																					<tr>
																						<td>
																							<input id="optFieldRecSel_First" name="optFieldRecSel" type="radio"
																								onclick="field_refreshChildFrame()" />
																						</td>
																						<td width="5">&nbsp;</td>
																						<td nowrap>
																							<label
																								tabindex="-1"
																								for="optFieldRecSel_First"
																								class="radio">
																								First</label>
																						</td>
																						<td width="20">&nbsp;&nbsp;&nbsp;</td>
																						<td>
																							<input id="optFieldRecSel_Last" name="optFieldRecSel" type="radio"
																								onclick="field_refreshChildFrame()" />
																						</td>
																						<td width="5">&nbsp;</td>
																						<td nowrap>
																							<label
																								tabindex="-1"
																								for="optFieldRecSel_Last"
																								class="radio">
																								Last</label>
																						</td>
																						<td width="20">&nbsp;&nbsp;&nbsp;</td>
																						<td>
																							<input id="optFieldRecSel_Specific" name="optFieldRecSel" type="radio"
																								onclick="field_refreshChildFrame()" />
																						</td>
																						<td width="5">&nbsp;</td>
																						<td nowrap>
																							<div id="divFieldRecSel_Specific" style="visibility: hidden; display: none">
																								<label
																									tabindex="-1"
																									for="optFieldRecSel_Specific"
																									class="radio">
																									Specific</label>
																							</div>
																						</td>
																						<td width="5">&nbsp;</td>
																						<td width="100%">
																							<input id="txtFieldRecSel_Specific" name="txtFieldRecSel_Specific" class="text">
																						</td>
																					</tr>
																				</table>
																			</td>
																			<td width="10">&nbsp;</td>
																		</tr>

																		<tr height="10">
																			<td colspan="6"></td>
																		</tr>

																		<tr height="10">
																			<%--<td width="20">&nbsp;&nbsp;</td>--%>
																			<td width="65" nowrap>Order :</td>
																			<%--<td width="20">&nbsp;&nbsp;</td>--%>
																			<td width="50%">
																				<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																					<tr>
																						<td width="95%">
																							<input type="text" id="txtFieldRecOrder" name="txtFieldRecOrder" class="text textdisabled" style="width: 99%" disabled="disabled">
																						</td>
																						<td >
																							<input id="btnFieldRecOrder" name="btnFieldRecOrder" style="width: 100%; " class="btn" type="button" value="..."
																								onclick="field_selectRecOrder()"/>
																						</td>
																					</tr>
																				</table>
																			</td>
																			<td width="30%">&nbsp;</td>
																			<td width="10">&nbsp;&nbsp;</td>
																		</tr>

																		<tr height="5">
																			<td colspan="6"></td>
																		</tr>

																		<tr height="10">
																			<%--<td width="20">&nbsp;&nbsp;</td>--%>
																			<td width="65" nowrap>Filter :</td>
																			<%--<td width="20">&nbsp;&nbsp;</td>--%>
																			<td width="50%">
																				<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																					<tr>
																						<td  width="95%">
																							<input type="text" id="txtFieldRecFilter" name="txtFieldRecFilter" class="text textdisabled" style="width: 99%" disabled="disabled">
																						</td>
																						<td >
																							<input id="btnFieldRecFilter" name="btnFieldRecFilter" class="btn" style="width: 100%; " type="button" value="..."
																								onclick="field_selectRecFilter()"/>
																						</td>
																					</tr>
																				</table>
																			</td>
																			<td width="30%">&nbsp;</td>
																			<td width="10">&nbsp;&nbsp;</td>
																		</tr>

																		<tr height="10">
																			<td colspan="6"></td>
																		</tr>
																	</table>
																</td>
															</tr>
														</table>
													</td>
													<td width="10">&nbsp;</td>
												</tr>

												<tr height="100%">
													<td colspan="6">&nbsp;</td>
												</tr>
												<%End If%>
											</table>
										</div>

										<div id="divFunction" style="visibility: hidden; display: none">
											<table height="100%" width="100%" class="invisible" cellspacing="0" cellpadding="0">
												<tr height="10">
													<td colspan="3">&nbsp;&nbsp;</td>
												</tr>

												<tr height="10">
													<td width="10">&nbsp;</td>
													<td><strong>Function</strong></td>
													<td width="10">&nbsp;</td>
												</tr>

												<tr height="10">
													<td colspan="3"></td>
												</tr>

												<tr>
													<td width="10">&nbsp;</td>
													<td>
														<object classid="clsid:1C203F13-95AD-11D0-A84B-00A0247B735B" id="SSFunctionTree" codebase="cabs/SStree.cab#version=1,0,2,24" style="LEFT: 0px; TOP: 0px; WIDTH: 100%; HEIGHT: 400px" viewastext>
															<param name="_ExtentX" value="2646">
															<param name="_ExtentY" value="1323">
															<param name="_Version" value="65538">
															<param name="BackColor" value="-2147483643">
															<param name="ForeColor" value="-2147483640">
															<param name="ImagesMaskColor" value="12632256">
															<param name="PictureBackgroundMaskColor" value="12632256">
															<param name="Appearance" value="1">
															<param name="BorderStyle" value="0">
															<param name="LabelEdit" value="1">
															<param name="LineStyle" value="0">
															<param name="LineType" value="1">
															<param name="MousePointer" value="0">
															<param name="NodeSelectionStyle" value="2">
															<param name="PictureAlignment" value="0">
															<param name="ScrollStyle" value="0">
															<param name="Style" value="6">
															<param name="IndentationStyle" value="0">
															<param name="TreeTips" value="3">
															<param name="PictureBackgroundStyle" value="0">
															<param name="Indentation" value="38">
															<param name="MaxLines" value="1">
															<param name="TreeTipDelay" value="500">
															<param name="ImageCount" value="0">
															<param name="ImageListIndex" value="-1">
															<param name="OLEDragMode" value="0">
															<param name="OLEDropMode" value="0">
															<param name="AllowDelete" value="0">
															<param name="AutoSearch" value="0">
															<param name="Enabled" value="-1">
															<param name="HideSelection" value="0">
															<param name="ImagesUseMask" value="0">
															<param name="Redraw" value="-1">
															<param name="UseImageList" value="-1">
															<param name="PictureBackgroundUseMask" value="0">
															<param name="HasFont" value="0">
															<param name="HasMouseIcon" value="0">
															<param name="HasPictureBackground" value="0">
															<param name="PathSeparator" value="\">
															<param name="TabStops" value="32">
															<param name="ImageList" value="<None>">
															<param name="LoadStyleRoot" value="1">
															<param name="Sorted" value="0">
															<param name="OnDemandDiscardBuffer" value="10">
														</object>
													</td>
													<td width="10">&nbsp;</td>
												</tr>

												<tr height="10">
													<td colspan="3">&nbsp;</td>
												</tr>
											</table>
										</div>

										<div id="divOperator" style="visibility: hidden; display: none">
											<table height="100%" width="100%" class="invisible" cellspacing="0" cellpadding="0">
												<tr height="10">
													<td colspan="3">&nbsp;&nbsp;</td>
												</tr>

												<tr height="10">
													<td width="10">&nbsp;</td>
													<td><strong>Operator</strong></td>
													<td width="10">&nbsp;</td>
												</tr>

												<tr height="10">
													<td colspan="3"></td>
												</tr>

												<tr>
													<td width="10">&nbsp;</td>
													<td>
														<object classid="clsid:1C203F13-95AD-11D0-A84B-00A0247B735B" id="SSOperatorTree" codebase="cabs/SStree.cab#version=1,0,2,24" style="LEFT: 0px; TOP: 0px; WIDTH: 100%; HEIGHT: 400px" viewastext>
															<param name="_ExtentX" value="2646">
															<param name="_ExtentY" value="1323">
															<param name="_Version" value="65538">
															<param name="BackColor" value="-2147483643">
															<param name="ForeColor" value="-2147483640">
															<param name="ImagesMaskColor" value="12632256">
															<param name="PictureBackgroundMaskColor" value="12632256">
															<param name="Appearance" value="1">
															<param name="BorderStyle" value="0">
															<param name="LabelEdit" value="1">
															<param name="LineStyle" value="0">
															<param name="LineType" value="1">
															<param name="MousePointer" value="0">
															<param name="NodeSelectionStyle" value="2">
															<param name="PictureAlignment" value="0">
															<param name="ScrollStyle" value="0">
															<param name="Style" value="6">
															<param name="IndentationStyle" value="0">
															<param name="TreeTips" value="3">
															<param name="PictureBackgroundStyle" value="0">
															<param name="Indentation" value="38">
															<param name="MaxLines" value="1">
															<param name="TreeTipDelay" value="500">
															<param name="ImageCount" value="0">
															<param name="ImageListIndex" value="-1">
															<param name="OLEDragMode" value="0">
															<param name="OLEDropMode" value="0">
															<param name="AllowDelete" value="0">
															<param name="AutoSearch" value="0">
															<param name="Enabled" value="-1">
															<param name="HideSelection" value="0">
															<param name="ImagesUseMask" value="0">
															<param name="Redraw" value="-1">
															<param name="UseImageList" value="-1">
															<param name="PictureBackgroundUseMask" value="0">
															<param name="HasFont" value="0">
															<param name="HasMouseIcon" value="0">
															<param name="HasPictureBackground" value="0">
															<param name="PathSeparator" value="\">
															<param name="TabStops" value="32">
															<param name="ImageList" value="<None>">
															<param name="LoadStyleRoot" value="1">
															<param name="Sorted" value="0">
															<param name="OnDemandDiscardBuffer" value="10">
														</object>
													</td>
													<td width="10">&nbsp;</td>
												</tr>

												<tr height="10">
													<td colspan="3">&nbsp;</td>
												</tr>
											</table>
										</div>

										<div id="divValue" style="visibility: hidden; display: none">
											<table height="100%" width="100%" class="invisible" cellspacing="0" cellpadding="0">
												<tr height="10">
													<td colspan="6">&nbsp;&nbsp;</td>
												</tr>

												<tr height="10">
													<td width="20">&nbsp;&nbsp;</td>
													<td colspan="4"><strong>Value</strong></td>
													<td width="20">&nbsp;&nbsp;</td>
												</tr>

												<tr height="10">
													<td colspan="6"></td>
												</tr>

												<tr height="10">
													<td width="10">&nbsp;</td>
													<td width="10" nowrap>Type :</td>
													<td width="20">&nbsp;&nbsp;</td>
													<td width="50%">
														<select id="cboValueType" name="cboValueType" class="combo" style="width: 100%" onchange="value_changeType()">
															<option value="1">
															Character
														<option value="2">
															Numeric
														<option value="3">
															Logic
														<option value="4">
															Date
														</select>
													</td>
													<td width="50%">&nbsp;</td>
													<td width="10">&nbsp;</td>
												</tr>

												<tr height="5">
													<td colspan="6"></td>
												</tr>

												<tr height="10">
													<td width="10">&nbsp;</td>
													<td width="10" nowrap>Value :</td>
													<td width="20">&nbsp;&nbsp;</td>
													<td width="50%">
														<select id="selectValue" name='selectValue"' class="combo" style="width: 100%">
															<option value="1">True</option>
															<option value="0">False</option>
														</select>
														<input id="txtValue" name="txtValue" class="text" style="LEFT: 0px; POSITION: absolute; TOP: 0px; VISIBILITY: hidden; WIDTH: 0px">
													</td>
													<td width="50%">&nbsp;</td>
													<td width="10">&nbsp;&nbsp;</td>
												</tr>

												<tr height="100%">
													<td colspan="6">&nbsp;</td>
												</tr>
											</table>
										</div>

										<div id="divLookupValue" style="visibility: hidden; display: none">
											<table height="100%" width="100%" class="invisible" cellspacing="0" cellpadding="0">
												<tr height="10">
													<td colspan="6">&nbsp;&nbsp;</td>
												</tr>

												<tr height="10">
													<td width="20">&nbsp;&nbsp;</td>
													<td colspan="4"><strong>Lookup Table Value</strong></td>
													<td width="20">&nbsp;&nbsp;</td>
												</tr>

												<tr height="10">
													<td colspan="6"></td>
												</tr>

												<tr height="10">
													<td width="10">&nbsp;</td>
													<td width="10" nowrap>Table :</td>
													<td width="20">&nbsp;&nbsp;</td>
													<td width="50%">
														<select id="cboLookupValueTable" name="cboLookupValueTable" class="combo" style="width: 100%" onchange="lookupValue_changeTable()">
														</select>
													</td>
													<td width="50%">&nbsp;</td>
													<td width="10">&nbsp;</td>
												</tr>

												<tr height="5">
													<td colspan="6"></td>
												</tr>

												<tr height="10">
													<td width="10">&nbsp;</td>
													<td width="10" nowrap>Column :</td>
													<td width="20">&nbsp;&nbsp;</td>
													<td width="50%">
														<select id="cboLookupValueColumn" name="cboLookupValueColumn" class="combo" style="width: 100%" onchange="lookupValue_changeColumn()">
														</select>
													</td>
													<td width="50%">&nbsp;</td>
													<td width="10">&nbsp;</td>
												</tr>

												<tr height="5">
													<td colspan="6"></td>
												</tr>

												<tr height="10">
													<td width="10">&nbsp;</td>
													<td width="10" nowrap>Value :</td>
													<td width="20">&nbsp;&nbsp;</td>
													<td width="50%">
														<select id="cboLookupValueValue" name="cboLookupValueValue" class="combo" style="width: 100%">
														</select>
													</td>
													<td width="50%">&nbsp;</td>
													<td width="10">&nbsp;</td>
												</tr>

												<tr height="10">
													<td colspan="6"></td>
												</tr>

												<tr height="10">
													<td width="20">&nbsp;&nbsp;</td>
													<td colspan="4">
														<input type="text" class="textwarning" id="txtValueNotInLookup" name="txtValueNotInLookup" value="<value> does not appear in <table>.<column>" style="TEXT-ALIGN: left; WIDTH: 100%; visibility: hidden; display: none" readonly>
													</td>
													<td width="20">&nbsp;&nbsp;</td>
												</tr>

												<tr height="100%">
													<td colspan="6">&nbsp;</td>
												</tr>
											</table>
										</div>

										<div id="divCalculation" style="visibility: hidden; display: none">
											<table height="100%" width="100%" class="invisible" cellspacing="0" cellpadding="0">
												<tr height="10">
													<td colspan="3">&nbsp;&nbsp;</td>
												</tr>

												<tr height="10">
													<td width="10">&nbsp;</td>
													<td><strong>Calculation</strong></td>
													<td width="10">&nbsp;</td>
												</tr>

												<tr height="10">
													<td colspan="3"></td>
												</tr>

												<tr>
													<td width="10">&nbsp;</td>
													<td>
														<%Html.RenderPartial("~/Views/Shared/Util_Def_CustomReports/ssOleDBGridCalculations.ascx")%>
													</td>
													<td width="10">&nbsp;</td>
												</tr>

												<tr>
													<td colspan="3" height="10"></td>
												</tr>

												<tr>
													<td width="10">&nbsp;</td>
													<td height="60">
														<textarea id="txtCalcDescription" name="txtCalcDescription" class="textarea disabled" tabindex="-1"
															style="HEIGHT: 99%; WIDTH: 100%;" wrap="VIRTUAL" disabled="disabled">
													</textarea>
													</td>
													<td width="10">&nbsp;</td>
												</tr>

												<tr>
													<td colspan="3" height="10"></td>
												</tr>

												<tr>
													<td width="10">&nbsp;</td>
													<td height="10">
														<input <%	If Session("OnlyMine") Then Response.Write("checked")%> type="checkbox" name="chkOwnersCalcs" id="chkOwnersCalcs" value="chkOwnersCalcs" tabindex="0"
															onclick="calculationAndFilter_refresh();"/>
														<label
															for="chkOwnersCalcs"
															class="checkbox"
															tabindex="-1"
															onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}">
															Only show calculations where owner is '<% =session("Username") %>'
														</label>
													</td>
													<td width="10">&nbsp;</td>
												</tr>

												<tr height="10">
													<td colspan="3">&nbsp;</td>
												</tr>
											</table>
										</div>

										<div id="divFilter" style="visibility: hidden; display: none">
											<table height="100%" width="100%" class="invisible" cellspacing="0" cellpadding="0">
												<tr height="10">
													<td colspan="3">&nbsp;&nbsp;</td>
												</tr>

												<tr height="10">
													<td width="10">&nbsp;</td>
													<td><strong>Filter</strong></td>
													<td width="10">&nbsp;</td>
												</tr>

												<tr height="10">
													<td colspan="3"></td>
												</tr>

												<tr>
													<td width="10">&nbsp;</td>
													<td>
														<object classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" id="ssOleDBGridFilters" name="ssOleDBGridFilters" codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" style="LEFT: 0px; TOP: 0px; WIDTH: 100%; HEIGHT: 400px">
															<param name="ScrollBars" value="4">
															<param name="_Version" value="196616">
															<param name="DataMode" value="2">
															<param name="Cols" value="0">
															<param name="Rows" value="0">
															<param name="BorderStyle" value="1">
															<param name="RecordSelectors" value="0">
															<param name="GroupHeaders" value="0">
															<param name="ColumnHeaders" value="0">
															<param name="GroupHeadLines" value="0">
															<param name="HeadLines" value="0">
															<param name="FieldDelimiter" value="(None)">
															<param name="FieldSeparator" value="(Tab)">
															<param name="Col.Count" value="2">
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
															<param name="MultiLine" value="0">
															<param name="ActiveCellStyleSet" value="">
															<param name="RowSelectionStyle" value="0">
															<param name="AllowRowSizing" value="0">
															<param name="AllowGroupSizing" value="0">
															<param name="AllowColumnSizing" value="0">
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
															<param name="Columns.Count" value="2">
															<param name="Columns(0).Width" value="100000">
															<param name="Columns(0).Visible" value="-1">
															<param name="Columns(0).Columns.Count" value="1">
															<param name="Columns(0).Caption" value="Name">
															<param name="Columns(0).Name" value="Name">
															<param name="Columns(0).Alignment" value="0">
															<param name="Columns(0).CaptionAlignment" value="3">
															<param name="Columns(0).Bound" value="0">
															<param name="Columns(0).AllowSizing" value="1">
															<param name="Columns(0).DataField" value="Column 0">
															<param name="Columns(0).DataType" value="8">
															<param name="Columns(0).Level" value="0">
															<param name="Columns(0).NumberFormat" value="">
															<param name="Columns(0).Case" value="0">
															<param name="Columns(0).FieldLen" value="256">
															<param name="Columns(0).VertScrollBar" value="0">
															<param name="Columns(0).Locked" value="0">
															<param name="Columns(0).Style" value="0">
															<param name="Columns(0).ButtonsAlways" value="0">
															<param name="Columns(0).RowCount" value="0">
															<param name="Columns(0).ColCount" value="1">
															<param name="Columns(0).HasHeadForeColor" value="0">
															<param name="Columns(0).HasHeadBackColor" value="0">
															<param name="Columns(0).HasForeColor" value="0">
															<param name="Columns(0).HasBackColor" value="0">
															<param name="Columns(0).HeadForeColor" value="0">
															<param name="Columns(0).HeadBackColor" value="0">
															<param name="Columns(0).ForeColor" value="0">
															<param name="Columns(0).BackColor" value="0">
															<param name="Columns(0).HeadStyleSet" value="">
															<param name="Columns(0).StyleSet" value="">
															<param name="Columns(0).Nullable" value="1">
															<param name="Columns(0).Mask" value="">
															<param name="Columns(0).PromptInclude" value="0">
															<param name="Columns(0).ClipMode" value="0">
															<param name="Columns(0).PromptChar" value="95">
															<param name="Columns(1).Width" value="0">
															<param name="Columns(1).Visible" value="0">
															<param name="Columns(1).Columns.Count" value="1">
															<param name="Columns(1).Caption" value="id">
															<param name="Columns(1).Name" value="id">
															<param name="Columns(1).Alignment" value="0">
															<param name="Columns(1).CaptionAlignment" value="3">
															<param name="Columns(1).Bound" value="0">
															<param name="Columns(1).AllowSizing" value="1">
															<param name="Columns(1).DataField" value="Column 1">
															<param name="Columns(1).DataType" value="8">
															<param name="Columns(1).Level" value="0">
															<param name="Columns(1).NumberFormat" value="">
															<param name="Columns(1).Case" value="0">
															<param name="Columns(1).FieldLen" value="256">
															<param name="Columns(1).VertScrollBar" value="0">
															<param name="Columns(1).Locked" value="0">
															<param name="Columns(1).Style" value="0">
															<param name="Columns(1).ButtonsAlways" value="0">
															<param name="Columns(1).RowCount" value="0">
															<param name="Columns(1).ColCount" value="1">
															<param name="Columns(1).HasHeadForeColor" value="0">
															<param name="Columns(1).HasHeadBackColor" value="0">
															<param name="Columns(1).HasForeColor" value="0">
															<param name="Columns(1).HasBackColor" value="0">
															<param name="Columns(1).HeadForeColor" value="0">
															<param name="Columns(1).HeadBackColor" value="0">
															<param name="Columns(1).ForeColor" value="0">
															<param name="Columns(1).BackColor" value="0">
															<param name="Columns(1).HeadStyleSet" value="">
															<param name="Columns(1).StyleSet" value="">
															<param name="Columns(1).Nullable" value="1">
															<param name="Columns(1).Mask" value="">
															<param name="Columns(1).PromptInclude" value="0">
															<param name="Columns(1).ClipMode" value="0">
															<param name="Columns(1).PromptChar" value="95">
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
													<td width="10">&nbsp;</td>
												</tr>

												<tr>
													<td colspan="3" height="10"></td>
												</tr>

												<tr>
													<td width="10">&nbsp;</td>
													<td height="60">
														<textarea id="txtFilterDescription" name="txtFilterDescription" class="textarea disabled" tabindex="-1"
															style="HEIGHT: 99%; WIDTH: 100%;" wrap="VIRTUAL" disabled="disabled">
													</textarea>
													</td>
													<td width="10">&nbsp;</td>
												</tr>

												<tr>
													<td colspan="3" height="10"></td>
												</tr>

												<tr>
													<td width="10">&nbsp;</td>
													<td height="10">
														<input <%	If Session("OnlyMine") Then Response.Write("checked")%> type="checkbox" name="chkOwnersFilters" id="chkOwnersFilters" value="chkOwnersFilters" tabindex="0"
															onclick="calculationAndFilter_refresh();"/>
														<label
															for="chkOwnersFilters"
															class="checkbox"
															tabindex="-1">
															Only show filters where owner is '<% =session("Username") %>'
														</label>
													</td>
													<td width="10">&nbsp;</td>
												</tr>

												<tr height="10">
													<td colspan="3">&nbsp;</td>
												</tr>
											</table>
										</div>

										<div id="divPromptedValue" style="visibility: hidden; display: none">
											<div style="padding-bottom: 10px">
													<br/>
													<strong>Prompted Value</strong>
											</div>
											<div>Prompt : <input id="Text1" name="txtPrompt" class="text" onkeyup="pVal_changePrompt()" style="width: 500px"></div>

											<table class="invisible">

												<tr height="5">
													<td colspan="6">&nbsp;</td>
												</tr>

												<tr height="10">
													<td colspan="4">
														<table width="100%" height="100%" cellspacing="0" cellpadding="0">
															<tr>
																<td>
																	<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																		<tr height="5">
																			<td colspan="11"></td>
																		</tr>

																		<tr height="10">
																			<td width="10">&nbsp;</td>
																			<td colspan="9"><strong>Type</strong></td>
																			<td width="10">&nbsp;</td>
																		</tr>

																		<tr height="10">
																			<td colspan="11"></td>
																		</tr>

																		<tr height="10">
																			<td width="20">&nbsp;&nbsp;</td>
																			<td width="40%">
																				<select id="cboPValType" name="cboPValType" class="combo" style="width: 100%" onchange="pVal_changeType()">
																					<option value="1">
																					Character
																			<option value="2">
																					Numeric
																			<option value="3">
																					Logic
																			<option value="4">
																					Date
																			<option value="5">
																					Lookup Value
																				</select>
																			</td>
																			<td width="20">&nbsp;&nbsp;</td>
																			<td width="10" nowrap id="tdPValSizePrompt" name="tdPValSizePrompt">Size :
																			</td>
																			<td width="20">&nbsp;&nbsp;</td>
																			<td width="30%">
																				<input class="text" id="txtPValSize" name="txtPValSize" style="width: 100%">
																			</td>
																			<td width="20">&nbsp;&nbsp;</td>
																			<td width="10" nowrap id="tdPValDecimalsPrompt" name="tdPValDecimalsPrompt">Decimals :
																			</td>
																			<td width="20">&nbsp;&nbsp;</td>
																			<td width="30%">
																				<input class="text" id="txtPValDecimals" name="txtPValDecimals" style="width: 100%">
																			</td>
																			<td width="10">&nbsp;&nbsp;</td>
																		</tr>

																		<tr height="10">
																			<td colspan="11"></td>
																		</tr>
																	</table>
																</td>
															</tr>
														</table>

													</td>
													<td width="20">&nbsp;&nbsp;</td>
												</tr>

<%--												<tr height="5">
													<td colspan="6">&nbsp;</td>
												</tr>--%>

												<tr height="10" id="trPValFormat" style="height: 10px">
													<%--<td width="10">&nbsp;</td>--%>
													<td colspan="8">
														<table style="width:100%; height:100%">
															<tr>
																<%--<td style="height: 125px">--%>
																<td>
																	<table width="100%" class="invisible">
<%--																		<tr height="5">
																			<td colspan="8"></td>
																		</tr>--%>

																		<tr height="10">
																			<td colspan="6"><strong>Mask</strong></td>
																			<td width="10">&nbsp;</td>
																		</tr>

																		<tr height="10">
																		</tr>

																		<tr height="10">
																			<%--<td width="20">&nbsp;&nbsp;</td>--%>
																			<td width="100%" colspan="7">
																				<input id="txtPValFormat" name="txtPValFormat" class="text" style="width: 100%">
																			</td>
																			<td style="width:20px">&nbsp;&nbsp;</td>
																		</tr>

																		<tr height="5">
																			<td colspan="8"></td>
																		</tr>

																		<tr height="10">
																			<%--<td width="20">&nbsp;&nbsp;</td>--%>
																			<td nowrap width="5%">A - Uppercase</td>
																			<td width="20%">&nbsp;&nbsp;</td>
																			<td nowrap width="5%">9 - Numbers (0-9)</td>
																			<td width="20%">&nbsp;&nbsp;</td>
																			<td nowrap width="5%">B - Binary (0 or 1)</td>
																			<td></td>
																		</tr>

																		<tr height="5">
																			<td colspan="8"></td>
																		</tr>

																		<tr height="10">
																			<%--<td width="20">&nbsp;&nbsp;</td>--%>
																			<td nowrap width="5%">a - Lowercase</td>
																			<td width="20%">&nbsp;&nbsp;</td>
																			<td nowrap width="5%"># - Numbers, Symbols</td>
																			<td width="20%">&nbsp;&nbsp;</td>
																			<td nowrap width="5%">\ - Follow by any literal</td>
																			<td></td>
																		</tr>

																		<tr height="10">
																			<td colspan="11"></td>
																		</tr>
																	</table>
																</td>
															</tr>
														</table>

													</td>
													<%--<td width="20">&nbsp;&nbsp;</td>--%>
												</tr>

												<tr height="5" id="trPValFormat2">
													<td colspan="6">&nbsp;</td>
												</tr>

												<tr height="10" id="trPValLookup" style="visibility: hidden; display: none">
													<%--<td width="10">&nbsp;</td>--%>
													<td colspan="4">
														<table width="100%" height="100%"  cellspacing="0" cellpadding="0">
															<tr>
																<td>
																	<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																		<tr height="5">
																			<td colspan="6"></td>
																		</tr>

																		<tr height="10">
																			<td width="10">&nbsp;</td>
																			<td colspan="4"><strong>Lookup Table Value</strong></td>
																			<td width="10">&nbsp;</td>
																		</tr>

																		<tr height="10">
																			<td colspan="6"></td>
																		</tr>

																		<tr height="10">
																			<td width="20">&nbsp;&nbsp;</td>
																			<td width="10" nowrap>Table :</td>
																			<td width="20">&nbsp;&nbsp;</td>
																			<td width="50%">
																				<select id="cboPValTable" name="cboPValTable" class="combo" style="width: 400px" onchange="pVal_changeTable()">
																				</select>
																			</td>
																			<td width="50%">&nbsp;</td>
																			<td width="10">&nbsp;</td>
																		</tr>

																		<tr height="5">
																			<td colspan="6"></td>
																		</tr>

																		<tr height="10">
																			<td width="10">&nbsp;</td>
																			<td width="10" nowrap>Column :</td>
																			<td width="20">&nbsp;&nbsp;</td>
																			<td width="50%">
																				<select id="cboPValColumn" name="cboPValColumn" class="combo" style="width: 100%" onchange="pVal_changeColumn()">
																				</select>
																			</td>
																			<td width="50%">&nbsp;</td>
																			<td width="10">&nbsp;&nbsp;</td>
																		</tr>

																		<tr height="10">
																			<td colspan="6"></td>
																		</tr>
																	</table>
																</td>
															</tr>
														</table>

													</td>
													<td width="20">&nbsp;&nbsp;</td>
												</tr>

												<tr height="5" id="trPValLookup2">
													<td colspan="6">&nbsp;</td>
												</tr>
												<%--<tr height="100%">
													<td colspan="6">&nbsp;</td>
												</tr>--%>

												<tr height="10">
													<%--<td width="10">&nbsp;</td>--%>
													<td colspan="4">
														<%--<table width="100%" height="100%">--%>
															<table width="100%" height="100%" >
															<tr>
																<td>
																	<table width="100%" class="invisible" >
																		<tr height="5">
																			<td colspan="8"></td>
																		</tr>

																		<tr height="10">
																			<%--<td width="10">&nbsp;</td>--%>
																			<td colspan="6"><strong>Default Value</strong></td>
																			<td width="10">&nbsp;</td>
																		</tr>

																		<%--<tr height="10">
																			<td></td>
																		</tr>--%>

																		<tr height="10" id="trPValDateOptions">
																			<td width="20">&nbsp;&nbsp;</td>
																			<td width="100%" colspan="6">
																				<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																					<tr height="10">
																						<td width="5">
																							<input id="optPValDate_Explicit" name="optPValDate" type="radio" selected
																								onclick="pVal_changeDateOption(0)" />
																						</td>
																						<td width="5">&nbsp;</td>
																						<td nowrap>
																							<label
																								tabindex="-1"
																								for="optPValDate_Explicit"
																								class="radio">Explicit</label>
																						</td>
																						<td width="20">&nbsp;&nbsp;</td>
																						<td width="5">
																							<input id="optPValDate_MonthStart" name="optPValDate" type="radio"
																								onclick="pVal_changeDateOption(2)"/>
																						</td>
																						<td width="5">&nbsp;</td>
																						<td nowrap>
																							<label
																								tabindex="-1"
																								for="optPValDate_MonthStart"
																								class="radio" >Month Start</label>
																						</td>
																						<td width="20">&nbsp;&nbsp;</td>
																						<td width="5">
																							<input id="optPValDate_YearStart" name="optPValDate" type="radio"
																								onclick="pVal_changeDateOption(4)"/>
																						</td>
																						<td width="5">&nbsp;</td>
																						<td nowrap>
																							<label
																								tabindex="-1"
																								for="optPValDate_YearStart"
																								class="radio">Year Start</label>
																						</td>
																						<td width="100%">&nbsp;</td>
																					</tr>

																					<tr height="10">
																						<td width="5">
																							<input id="optPValDate_Current" name="optPValDate" type="radio"
																								onclick="pVal_changeDateOption(1)" />
																						</td>
																						<td width="5">&nbsp;</td>
																						<td nowrap>
																							<label
																								tabindex="-1"
																								for="optPValDate_Current"
																								class="radio" >Current</label>
																						</td>
																						<td width="20">&nbsp;&nbsp;</td>
																						<td width="5">
																							<input id="optPValDate_MonthEnd" name="optPValDate" type="radio"
																								onclick="pVal_changeDateOption(3)" />
																						</td>
																						<td width="5">&nbsp;</td>
																						<td nowrap>
																							<label
																								tabindex="-1"
																								for="optPValDate_MonthEnd"
																								class="radio">Month End</label>
																						</td>
																						<td width="20">&nbsp;&nbsp;</td>
																						<td width="5">
																							<input id="optPValDate_YearEnd" name="optPValDate" type="radio"
																								onclick="pVal_changeDateOption(5)" />
																						</td>
																						<td width="5">&nbsp;</td>
																						<td nowrap>
																							<label
																								tabindex="-1"
																								for="optPValDate_YearEnd"
																								class="radio" >Year End</label>
																						</td>
																						<td width="100%">&nbsp;</td>
																					</tr>
																				</table>
																			</td>
																			<td width="20">&nbsp;&nbsp;</td>
																		</tr>

																		<tr height="10" id="trPValDateOptions2">
																			<td colspan="8"></td>
																		</tr>

																		<tr height="10" id="trPValTextDefault">
																			<%--<td width="20">&nbsp;&nbsp;</td>--%>
																			<td width="100%" colspan="6">
																				<input id="txtPValDefault" name="txtPValDefault" class="text" style="width: 100%">
																			</td>
																			<td width="20">&nbsp;&nbsp;</td>
																		</tr>

																		<tr height="10" id="trPValComboDefault" style="visibility: hidden; display: none">
																			<%--<td width="20">&nbsp;&nbsp;</td>--%>
																			<td width="100%" colspan="6">
																				<select id="cboPValDefault" name="cboPValDefault" style="width: 100%">
																				</select>
																			</td>
																			<td width="20">&nbsp;&nbsp;</td>
																		</tr>

																		<tr height="10">
																			<td colspan="11"></td>
																		</tr>
																	</table>
																</td>
															</tr>
														</table>

													</td>
													<td width="20">&nbsp;&nbsp;</td>
												</tr>

												<tr height="100%">
													<td colspan="6">&nbsp;</td>
												</tr>
											</table>
										</div>
									</td>
								</tr>
							</table>
						</td>
						<td width="10">&nbsp;&nbsp;</td>
					</tr>

					<tr>
						<td height="10" colspan="5"></td>
					</tr>

					<tr height="10">
						
						<td colspan="3" alinn="left">
							<table   width="100%" class="invisible" cellspacing="0" cellpadding="0">
								<tr>
									<td colspan="4"></td>
								</tr>
								<tr>
									
									<td width="10">
										<input id="cmdOK" name="cmdOK" type="button" class="btn" value="OK" style="width: 75px" width="75"
											onclick="component_OKClick()" />
									</td>
									<td width="20"></td>
									<td width="10">
										<input id="cmdCancel" name="cmdCancel" type="button" class="btn" value="Cancel" style="width: 75px" width="75"
											onclick="component_CancelClick()"/>
									</td>
                                    <td></td>
								</tr>
							</table>
						</td>
						<td width="10"></td>
                        <td width="10"></td>
					</tr>
					<tr>
						<td height="10" colspan="7"></td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</form>

<form id="util_def_exprcomponent_frmUseful" name="util_def_exprcomponent_frmUseful" style="visibility: hidden; display: none">
	<input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
	<input type="hidden" id="txtExprType" name="txtExprType" value='<%=session("optionExprType")%>'>
	<input type="hidden" id="txtExprID" name="txtExprID" value='<%=session("optionExprID")%>'>
	<input type="hidden" id="txtAction" name="txtAction" value='<%=session("optionAction")%>'>
	<input type="hidden" id="txtLinkRecordID" name="txtLinkRecordID" value='<%=session("optionLinkRecordID")%>'>
	<input type="hidden" id="txtTableID" name="txtTableID" value='<%=session("optionTableID")%>'>
	<input type="hidden" id="txtInitialising" name="txtInitialising" value="0">
	<input type="hidden" id="txtChildFieldOrderID" name="txtChildFieldOrderID" value="0">
	<input type="hidden" id="txtChildFieldFilterID" name="txtChildFieldFilterID" value="0">
	<input type="hidden" id="txtChildFieldFilterHidden" name="txtChildFieldFilterHidden" value="0">
	<input type="hidden" id="txtFunctionsLoaded" name="txtFunctionsLoaded" value="0">
	<input type="hidden" id="txtOperatorsLoaded" name="txtOperatorsLoaded" value="0">
	<input type="hidden" id="txtLookupTablesLoaded" name="txtLookupTablesLoaded" value="0">
	<input type="hidden" id="txtPValLookupTablesLoaded" name="txtPValLookupTablesLoaded" value="0">
</form>

<form action="util_def_exprComponent_Submit" method="post" id="frmGotoOption" name="frmGotoOption">
	<%Html.RenderPartial("~/Views/Shared/gotoOption.ascx")%>
</form>

<form id="util_def_exprcomponent_frmOriginalDefinition" name="util_def_exprcomponent_frmOriginalDefinition">
	<%
		Dim sDefnString As String
		Dim sFieldTableID As String
		Dim sFieldColumnID As String
		Dim sLookupTableID As String
		Dim sLookupColumnID As String
		
		On Error Resume Next
		sErrMsg = ""

		If Session("optionAction") = "EDITEXPRCOMPONENT" Then
			sDefnString = Session("optionExtension")

			Response.Write("<INPUT type='hidden' id=txtComponentID name=txtComponentID value=" & componentParameter(sDefnString, "COMPONENTID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtType name=txtType value=" & componentParameter(sDefnString, "TYPE") & ">" & vbCrLf)
			sFieldTableID = componentParameter(sDefnString, "FIELDTABLEID")
			sFieldColumnID = componentParameter(sDefnString, "FIELDCOLUMNID")
			Response.Write("<INPUT type='hidden' id=txtFieldPassBy name=txtFieldPassBy value=" & componentParameter(sDefnString, "FIELDPASSBY") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtFieldSelectionTableID name=txtFieldSelectionTableID value=" & componentParameter(sDefnString, "FIELDSELECTIONTABLEID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtFieldSelectionRecord name=txtFieldSelectionRecord value=" & componentParameter(sDefnString, "FIELDSELECTIONRECORD") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtFieldSelectionLine name=txtFieldSelectionLine value=" & componentParameter(sDefnString, "FIELDSELECTIONLINE") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtFieldSelectionOrderID name=txtFieldSelectionOrderID value=" & componentParameter(sDefnString, "FIELDSELECTIONORDERID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtFieldSelectionFilter name=txtFieldSelectionFilter value=" & componentParameter(sDefnString, "FIELDSELECTIONFILTER") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtFunctionID name=txtFunctionID value=" & componentParameter(sDefnString, "FUNCTIONID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtCalculationID name=txtCalculationID value=" & componentParameter(sDefnString, "CALCULATIONID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtOperatorID name=txtOperatorID value=" & componentParameter(sDefnString, "OPERATORID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtValueType name=txtValueType value=" & componentParameter(sDefnString, "VALUETYPE") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtValueCharacter name=txtValueCharacter value=""" & Replace(componentParameter(sDefnString, "VALUECHARACTER"), """", "&quot;") & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtValueNumeric name=txtValueNumeric value=" & componentParameter(sDefnString, "VALUENUMERIC") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtValueLogic name=txtValueLogic value=" & componentParameter(sDefnString, "VALUELOGIC") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtValueDate name=txtValueDate value=" & componentParameter(sDefnString, "VALUEDATE") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtPromptDescription name=txtPromptDescription value=""" & Replace(componentParameter(sDefnString, "PROMPTDESCRIPTION"), """", "&quot;") & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtPromptMask name=txtPromptMask value=""" & Replace(componentParameter(sDefnString, "PROMPTMASK"), """", "&quot;") & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtPromptSize name=txtPromptSize value=" & componentParameter(sDefnString, "PROMPTSIZE") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtPromptDecimals name=txtPromptDecimals value=" & componentParameter(sDefnString, "PROMPTDECIMALS") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtFunctionReturnType name=txtFunctionReturnType value=" & componentParameter(sDefnString, "FUNCTIONRETURNTYPE") & ">" & vbCrLf)
			sLookupTableID = componentParameter(sDefnString, "LOOKUPTABLEID")
			sLookupColumnID = componentParameter(sDefnString, "LOOKUPCOLUMNID")
			Response.Write("<INPUT type='hidden' id=txtFilterID name=txtFilterID value=" & componentParameter(sDefnString, "FILTERID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtFieldOrderName name=txtFieldOrderName value=""" & componentParameter(sDefnString, "FIELDSELECTIONORDERNAME") & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtFieldFilterName name=txtFieldFilterName value=""" & componentParameter(sDefnString, "FIELDSELECTIONFILTERNAME") & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtPromptDateType name=txtPromptDateType value=" & componentParameter(sDefnString, "PROMPTDATETYPE") & ">" & vbCrLf)
		Else
			Response.Write("<INPUT type='hidden' id=txtComponentID name=txtComponentID value=0>" & vbCrLf)
			sFieldTableID = Session("optionTableID")
			sFieldColumnID = 0
			sLookupTableID = 0
			sLookupColumnID = 0
			Response.Write("<INPUT type='hidden' id=txtFieldSelectionRecord name=txtFieldSelectionRecord value=1>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtValueCharacter name=txtValueCharacter value="""">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtValueNumeric name=txtValueNumeric value=0>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtValueLogic name=txtValueLogic value=""False"">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtValueDate name=txtValueDate value="""">" & vbCrLf)
		End If

		Response.Write("<INPUT type='hidden' id=txtFieldTableID name=txtFieldTableID value=" & sFieldTableID & ">" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtFieldColumnID name=txtFieldColumnID value=" & sFieldColumnID & ">" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtLookupTableID name=txtLookupTableID value=" & sLookupTableID & ">" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtLookupColumnID name=txtLookupColumnID value=" & sLookupColumnID & ">" & vbCrLf)
	%>
</form>

<form id="frmTables" name="frmTables">
	<%
		Dim cmdTables As Command
		Dim prmTableID As ADODB.Parameter
		Dim rstTables As Recordset
		Dim iCount As Integer
		
		If Len(sErrMsg) = 0 Then
			cmdTables = New Command
			cmdTables.CommandText = "sp_ASRIntGetExprTables"
			cmdTables.CommandType = CommandTypeEnum.adCmdStoredProc
			cmdTables.ActiveConnection = Session("databaseConnection")

			prmTableID = cmdTables.CreateParameter("tableID", 3, 1)	' 3=integer, 1=input
			cmdTables.Parameters.Append(prmTableID)
			prmTableID.value = CleanNumeric(Session("optionTableID"))

			Err.Clear()
			rstTables = cmdTables.Execute
			If (Err.Number <> 0) Then
				sErrMsg = "Error reading component tables." & vbCrLf & FormatError(Err.Description)
			Else
				If rstTables.state <> 0 Then
					' Read recordset values.
					iCount = 0
					Do While Not rstTables.EOF
						iCount = iCount + 1
						Response.Write("<input type='hidden' id=txtTable_" & iCount & " name=txtTable_" & iCount & " value=""" & rstTables.Fields("definitionString").Value & """>" & vbCrLf)
						rstTables.MoveNext()
					Loop

					' Release the ADO recordset object.
					rstTables.close()
				End If
				rstTables = Nothing
			End If

			' Release the ADO command object.
			cmdTables = Nothing
		End If
	%>
</form>

<form id="frmFunctions" name="frmFunctions">
	<%
		Dim cmdFunctions As Command
		Dim rstFunctions As Recordset
		
		If Len(sErrMsg) = 0 Then
			cmdFunctions = New Command
			cmdFunctions.CommandText = "sp_ASRIntGetExprFunctions"
			cmdFunctions.CommandType = CommandTypeEnum.adCmdStoredProc
			cmdFunctions.ActiveConnection = Session("databaseConnection")

			prmTableID = cmdFunctions.CreateParameter("tableID", 3, 1) ' 3=integer, 1=input
			cmdFunctions.Parameters.Append(prmTableID)
			prmTableID.value = CleanNumeric(Session("optionTableID"))

			Err.Clear()
			rstFunctions = cmdFunctions.Execute
			If (Err.Number <> 0) Then
				sErrMsg = "Error reading component functions." & vbCrLf & FormatError(Err.Description)
			Else
				If rstFunctions.state <> 0 Then
					' Read recordset values.
					iCount = 0
					Do While Not rstFunctions.EOF
						iCount = iCount + 1
						Response.Write("<input type='hidden' id=txtFunction_" & iCount & " name=txtFunction_" & iCount & " value=""" & rstFunctions.Fields("definitionString").Value & """>" & vbCrLf)
						rstFunctions.MoveNext()
					Loop

					' Release the ADO recordset object.
					rstFunctions.close()
				End If
				rstFunctions = Nothing
			End If

			' Release the ADO command object.
			cmdFunctions = Nothing
		End If
	%>
</form>

<form id="frmFunctionParameters" name="frmFunctionParameters">
	<%
		Dim cmdFunctionParameters As Command
		Dim rstFunctionParameters As Recordset
		
		If Len(sErrMsg) = 0 Then
			cmdFunctionParameters = New Command()
			cmdFunctionParameters.CommandText = "sp_ASRIntGetExprFunctionParameters"
			cmdFunctionParameters.CommandType = CommandTypeEnum.adCmdStoredProc
			cmdFunctionParameters.ActiveConnection = Session("databaseConnection")

			Err.Clear()
			rstFunctionParameters = cmdFunctionParameters.Execute
			If (Err.Number <> 0) Then
				sErrMsg = "Error reading component functions." & vbCrLf & FormatError(Err.Description)
			Else
				If rstFunctionParameters.state <> 0 Then
					' Read recordset values.
					iCount = 1
					Do While Not rstFunctionParameters.EOF
						Response.Write("<input type='hidden' id=txtFunctionParameters_" & rstFunctionParameters.Fields("functionID").Value & "_" & iCount & " name=txtFunctionParameters_" & rstFunctionParameters.Fields("functionID").Value & "_" & iCount & " value=""" & rstFunctionParameters.Fields("parameterName").Value & """>" & vbCrLf)
						iCount = iCount + 1
						rstFunctionParameters.MoveNext()
					Loop

					' Release the ADO recordset object.
					rstFunctionParameters.close()
				End If
				rstFunctionParameters = Nothing
			End If

			' Release the ADO command object.
			cmdFunctionParameters = Nothing
		End If
	%>
</form>

<form id="frmOperators" name="frmOperators">
	<%
		Dim cmdOperators As Command
		Dim rstOperators As Recordset
		
		If Len(sErrMsg) = 0 Then
			cmdOperators = New Command()
			cmdOperators.CommandText = "sp_ASRIntGetExprOperators"
			cmdOperators.CommandType = CommandTypeEnum.adCmdStoredProc
			cmdOperators.ActiveConnection = Session("databaseConnection")

			Err.Clear()
			rstOperators = cmdOperators.Execute
			If (Err.Number <> 0) Then
				sErrMsg = "Error reading component operators." & vbCrLf & FormatError(Err.Description)
			Else
				If rstOperators.state <> 0 Then
					' Read recordset values.
					iCount = 0
					Do While Not rstOperators.EOF
						iCount = iCount + 1
						Response.Write("<input type='hidden' id=txtOperator_" & iCount & " name=txtOperator_" & iCount & " value=""" & rstOperators.Fields("definitionString").Value & """>" & vbCrLf)
						rstOperators.MoveNext()
					Loop

					' Release the ADO recordset object.
					rstOperators.close()
				End If
				rstOperators = Nothing
			End If

			' Release the ADO command object.
			cmdOperators = Nothing
		End If
	%>
</form>

<form id="frmCalcs" name="frmCalcs">
	<%
		Dim cmdCalcs As Command
		Dim rstCalcs As Recordset
		Dim prmExprID As ADODB.Parameter
		Dim prmBaseTableID As ADODB.Parameter	
		
		If Len(sErrMsg) = 0 Then
			cmdCalcs = New Command
			cmdCalcs.CommandText = "sp_ASRIntGetExprCalcs"
			cmdCalcs.CommandType = CommandTypeEnum.adCmdStoredProc
			cmdCalcs.ActiveConnection = Session("databaseConnection")

			prmExprID = cmdCalcs.CreateParameter("exprID", 3, 1) ' 3=integer, 1=input
			cmdCalcs.Parameters.Append(prmExprID)
			prmExprID.value = CleanNumeric(CLng(Session("optionExprID")))

			prmBaseTableID = cmdCalcs.CreateParameter("baseTableID", 3, 1) ' 3=integer, 1=input
			cmdCalcs.Parameters.Append(prmBaseTableID)
			prmBaseTableID.value = CleanNumeric(CLng(Session("optionTableID")))

			Err.Clear()
			rstCalcs = cmdCalcs.Execute
			If (Err.Number <> 0) Then
				sErrMsg = "Error reading component calculations." & vbCrLf & FormatError(Err.Description)
			Else
				If rstCalcs.state <> 0 Then
					' Read recordset values.
					iCount = 0
					Do While Not rstCalcs.EOF
						iCount = iCount + 1
						Response.Write("<input type='hidden' id=txtCalc_" & iCount & " name=txtCalc_" & iCount & " value=""" & Replace(rstCalcs.Fields("definitionString").Value, """", "&quot;") & """>" & vbCrLf)
						Response.Write("<input type='hidden' id=txtCalcDesc_" & iCount & " name=txtCalcDesc_" & iCount & " value=""" & Replace(rstCalcs.Fields("description").Value, """", "&quot;") & """>" & vbCrLf)
						rstCalcs.MoveNext()
					Loop

					' Release the ADO recordset object.
					rstCalcs.close()
				End If
				rstCalcs = Nothing
			End If

			' Release the ADO command object.
			cmdCalcs = Nothing
		End If
	%>
</form>

<form id="frmFilters" name="frmFilters">
	<%
		Dim cmdFilters As Command
		Dim rstFilters As Recordset
		
		If Len(sErrMsg) = 0 Then
			cmdFilters = New Command
			cmdFilters.CommandText = "sp_ASRIntGetExprFilters"
			cmdFilters.CommandType = CommandTypeEnum.adCmdStoredProc
			cmdFilters.ActiveConnection = Session("databaseConnection")

			prmExprID = cmdFilters.CreateParameter("exprID", 3, 1) ' 3=integer, 1=input
			cmdFilters.Parameters.Append(prmExprID)
			prmExprID.value = CleanNumeric(CLng(Session("optionExprID")))

			prmBaseTableID = cmdFilters.CreateParameter("baseTableID", 3, 1) ' 3=integer, 1=input
			cmdFilters.Parameters.Append(prmBaseTableID)
			prmBaseTableID.value = CleanNumeric(CLng(Session("optionTableID")))

			Err.Clear()
			rstFilters = cmdFilters.Execute
			If (Err.Number <> 0) Then
				sErrMsg = "Error reading component filters." & vbCrLf & FormatError(Err.Description)
			Else
				If rstFilters.state <> 0 Then
					' Read recordset values.
					iCount = 0
					Do While Not rstFilters.EOF
						iCount = iCount + 1
						Response.Write("<input type='hidden' id=txtFilter_" & iCount & " name=txtFilter_" & iCount & " value=""" & Replace(rstFilters.Fields("definitionString").Value, """", "&quot;") & """>" & vbCrLf)
						Response.Write("<input type='hidden' id=txtFilterDesc_" & iCount & " name=txtFilterDesc_" & iCount & " value=""" & Replace(rstFilters.Fields("description").Value, """", "&quot;") & """>" & vbCrLf)
						rstFilters.MoveNext()
					Loop

					' Release the ADO recordset object.
					rstFilters.close()
				End If
				rstFilters = Nothing
			End If

			' Release the ADO command object.
			cmdFilters = Nothing
		End If
	%>
</form>

<form id="frmFieldRec" name="frmFieldRec" target="fieldRec" action="fieldRec" method="post" style="visibility: hidden; display: none">
	<input type="hidden" id="selectionType" name="selectionType">
	<input type="hidden" id="Hidden1" name="txtTableID">
	<input type="hidden" id="selectedID" name="selectedID">
</form>

<input type='hidden' id="txtTicker" name="txtTicker" value="0">
<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">

<%
	Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrMsg & """>" & vbCrLf)
%>


<script runat="server" language="vb">

	Function componentParameter(psDefnString, psParameter)
		Dim iCharIndex As Integer
		Dim sDefn As String
	
		sDefn = psDefnString
	
		iCharIndex = InStr(sDefn, "	")
		If iCharIndex >= 0 Then
			If psParameter = "COMPONENTID" Then
				componentParameter = Left(sDefn, iCharIndex - 1)
				Exit Function
			End If
		
			sDefn = Mid(sDefn, iCharIndex + 1)
			iCharIndex = InStr(sDefn, "	")
			If iCharIndex >= 0 Then
				If psParameter = "EXPRID" Then
					componentParameter = Left(sDefn, iCharIndex - 1)
					Exit Function
				End If
			
				sDefn = Mid(sDefn, iCharIndex + 1)
				iCharIndex = InStr(sDefn, "	")
				If iCharIndex >= 0 Then
					If psParameter = "TYPE" Then
						componentParameter = Left(sDefn, iCharIndex - 1)
						Exit Function
					End If
				
					sDefn = Mid(sDefn, iCharIndex + 1)
					iCharIndex = InStr(sDefn, "	")
					If iCharIndex >= 0 Then
						If psParameter = "FIELDCOLUMNID" Then
							componentParameter = Left(sDefn, iCharIndex - 1)
							Exit Function
						End If
					
						sDefn = Mid(sDefn, iCharIndex + 1)
						iCharIndex = InStr(sDefn, "	")
						If iCharIndex >= 0 Then
							If psParameter = "FIELDPASSBY" Then
								componentParameter = Left(sDefn, iCharIndex - 1)
								Exit Function
							End If
						
							sDefn = Mid(sDefn, iCharIndex + 1)
							iCharIndex = InStr(sDefn, "	")
							If iCharIndex >= 0 Then
								If psParameter = "FIELDSELECTIONTABLEID" Then
									componentParameter = Left(sDefn, iCharIndex - 1)
									Exit Function
								End If
							
								sDefn = Mid(sDefn, iCharIndex + 1)
								iCharIndex = InStr(sDefn, "	")
								If iCharIndex >= 0 Then
									If psParameter = "FIELDSELECTIONRECORD" Then
										componentParameter = Left(sDefn, iCharIndex - 1)
										Exit Function
									End If
								
									sDefn = Mid(sDefn, iCharIndex + 1)
									iCharIndex = InStr(sDefn, "	")
									If iCharIndex >= 0 Then
										If psParameter = "FIELDSELECTIONLINE" Then
											componentParameter = Left(sDefn, iCharIndex - 1)
											Exit Function
										End If
									
										sDefn = Mid(sDefn, iCharIndex + 1)
										iCharIndex = InStr(sDefn, "	")
										If iCharIndex >= 0 Then
											If psParameter = "FIELDSELECTIONORDERID" Then
												componentParameter = Left(sDefn, iCharIndex - 1)
												Exit Function
											End If
										
											sDefn = Mid(sDefn, iCharIndex + 1)
											iCharIndex = InStr(sDefn, "	")
											If iCharIndex >= 0 Then
												If psParameter = "FIELDSELECTIONFILTER" Then
													componentParameter = Left(sDefn, iCharIndex - 1)
													Exit Function
												End If
											
												sDefn = Mid(sDefn, iCharIndex + 1)
												iCharIndex = InStr(sDefn, "	")
												If iCharIndex >= 0 Then
													If psParameter = "FUNCTIONID" Then
														componentParameter = Left(sDefn, iCharIndex - 1)
														Exit Function
													End If
												
													sDefn = Mid(sDefn, iCharIndex + 1)
													iCharIndex = InStr(sDefn, "	")
													If iCharIndex >= 0 Then
														If psParameter = "CALCULATIONID" Then
															componentParameter = Left(sDefn, iCharIndex - 1)
															Exit Function
														End If
													
														sDefn = Mid(sDefn, iCharIndex + 1)
														iCharIndex = InStr(sDefn, "	")
														If iCharIndex >= 0 Then
															If psParameter = "OPERATORID" Then
																componentParameter = Left(sDefn, iCharIndex - 1)
																Exit Function
															End If
														
															sDefn = Mid(sDefn, iCharIndex + 1)
															iCharIndex = InStr(sDefn, "	")
															If iCharIndex >= 0 Then
																If psParameter = "VALUETYPE" Then
																	componentParameter = Left(sDefn, iCharIndex - 1)
																	Exit Function
																End If
															
																sDefn = Mid(sDefn, iCharIndex + 1)
																iCharIndex = InStr(sDefn, "	")
																If iCharIndex >= 0 Then
																	If psParameter = "VALUECHARACTER" Then
																		componentParameter = Left(sDefn, iCharIndex - 1)
																		Exit Function
																	End If
																
																	sDefn = Mid(sDefn, iCharIndex + 1)
																	iCharIndex = InStr(sDefn, "	")
																	If iCharIndex >= 0 Then
																		If psParameter = "VALUENUMERIC" Then
																			componentParameter = Left(sDefn, iCharIndex - 1)
																			Exit Function
																		End If
																	
																		sDefn = Mid(sDefn, iCharIndex + 1)
																		iCharIndex = InStr(sDefn, "	")
																		If iCharIndex >= 0 Then
																			If psParameter = "VALUELOGIC" Then
																				componentParameter = Left(sDefn, iCharIndex - 1)
																				Exit Function
																			End If
																		
																			sDefn = Mid(sDefn, iCharIndex + 1)
																			iCharIndex = InStr(sDefn, "	")
																			If iCharIndex >= 0 Then
																				If psParameter = "VALUEDATE" Then
																					componentParameter = Left(sDefn, iCharIndex - 1)
																					Exit Function
																				End If
																			
																				sDefn = Mid(sDefn, iCharIndex + 1)
																				iCharIndex = InStr(sDefn, "	")
																				If iCharIndex >= 0 Then
																					If psParameter = "PROMPTDESCRIPTION" Then
																						componentParameter = Left(sDefn, iCharIndex - 1)
																						Exit Function
																					End If
																				
																					sDefn = Mid(sDefn, iCharIndex + 1)
																					iCharIndex = InStr(sDefn, "	")
																					If iCharIndex >= 0 Then
																						If psParameter = "PROMPTMASK" Then
																							componentParameter = Left(sDefn, iCharIndex - 1)
																							Exit Function
																						End If
																					
																						sDefn = Mid(sDefn, iCharIndex + 1)
																						iCharIndex = InStr(sDefn, "	")
																						If iCharIndex >= 0 Then
																							If psParameter = "PROMPTSIZE" Then
																								componentParameter = Left(sDefn, iCharIndex - 1)
																								Exit Function
																							End If
																						
																							sDefn = Mid(sDefn, iCharIndex + 1)
																							iCharIndex = InStr(sDefn, "	")
																							If iCharIndex >= 0 Then
																								If psParameter = "PROMPTDECIMALS" Then
																									componentParameter = Left(sDefn, iCharIndex - 1)
																									Exit Function
																								End If
																							
																								sDefn = Mid(sDefn, iCharIndex + 1)
																								iCharIndex = InStr(sDefn, "	")
																								If iCharIndex >= 0 Then
																									If psParameter = "FUNCTIONRETURNTYPE" Then
																										componentParameter = Left(sDefn, iCharIndex - 1)
																										Exit Function
																									End If
																								
																									sDefn = Mid(sDefn, iCharIndex + 1)
																									iCharIndex = InStr(sDefn, "	")
																									If iCharIndex >= 0 Then
																										If psParameter = "LOOKUPTABLEID" Then
																											componentParameter = Left(sDefn, iCharIndex - 1)
																											Exit Function
																										End If
																									
																										sDefn = Mid(sDefn, iCharIndex + 1)
																										iCharIndex = InStr(sDefn, "	")
																										If iCharIndex >= 0 Then
																											If psParameter = "LOOKUPCOLUMNID" Then
																												componentParameter = Left(sDefn, iCharIndex - 1)
																												Exit Function
																											End If
																										
																											sDefn = Mid(sDefn, iCharIndex + 1)
																											iCharIndex = InStr(sDefn, "	")
																											If iCharIndex >= 0 Then
																												If psParameter = "FILTERID" Then
																													componentParameter = Left(sDefn, iCharIndex - 1)
																													Exit Function
																												End If
																											
																												sDefn = Mid(sDefn, iCharIndex + 1)
																												iCharIndex = InStr(sDefn, "	")
																												If iCharIndex >= 0 Then
																													If psParameter = "EXPANDEDNODE" Then
																														componentParameter = Left(sDefn, iCharIndex - 1)
																														Exit Function
																													End If
																												
																													sDefn = Mid(sDefn, iCharIndex + 1)
																													iCharIndex = InStr(sDefn, "	")
																													If iCharIndex >= 0 Then
																														If psParameter = "PROMPTDATETYPE" Then
																															componentParameter = Left(sDefn, iCharIndex - 1)
																															Exit Function
																														End If
																													
																														sDefn = Mid(sDefn, iCharIndex + 1)
																														iCharIndex = InStr(sDefn, "	")
																														If iCharIndex >= 0 Then
																															If psParameter = "DESCRIPTION" Then
																																componentParameter = Left(sDefn, iCharIndex - 1)
																																Exit Function
																															End If
																														
																															sDefn = Mid(sDefn, iCharIndex + 1)
																															iCharIndex = InStr(sDefn, "	")
																															If iCharIndex >= 0 Then
																																If psParameter = "FIELDTABLEID" Then
																																	componentParameter = Left(sDefn, iCharIndex - 1)
																																	Exit Function
																																End If
																															
																																sDefn = Mid(sDefn, iCharIndex + 1)
																																iCharIndex = InStr(sDefn, "	")
																																If iCharIndex >= 0 Then
																																	If psParameter = "FIELDSELECTIONORDERNAME" Then
																																		componentParameter = Left(sDefn, iCharIndex - 1)
																																		Exit Function
																																	End If
																																
																																	sDefn = Mid(sDefn, iCharIndex + 1)
																																	If psParameter = "FIELDSELECTIONFILTERNAME" Then
																																		componentParameter = sDefn
																																		Exit Function
																																	End If
																																End If
																															End If
																														End If
																													End If
																												End If
																											End If
																										End If
																									End If
																								End If
																							End If
																						End If
																					End If
																				End If
																			End If
																		End If
																	End If
																End If
															End If
														End If
													End If
												End If
											End If
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	
		componentParameter = ""
	End Function

</script>


<script type="text/javascript">
	util_def_exprcomponent_addhandlers();
	util_def_exprcomponent_onload();
</script>


