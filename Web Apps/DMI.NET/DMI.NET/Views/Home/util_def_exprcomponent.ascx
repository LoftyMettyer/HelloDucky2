<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script src="<%: Url.Content("~/bundles/utilities_expressions")%>" type="text/javascript"></script>

<object
    classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
    id="Microsoft_Licensed_Class_Manager_1_0"
    viewastext>
    <param name="LPKPath" value="lpks/main.lpk">
</object>


<FORM action="" method=POST id=frmMainForm name=frmMainForm>
<%
    Dim cmdParameter
    Dim prmFunctionID
    Dim prmParameterIndex
    Dim prmPassByType
    
    Dim iPassBy As Integer
    Dim sErrMsg As String
    
    
iPassBy = 1	
if (len(sErrMsg) = 0) and (Session("optionFunctionID") > 0) then
        cmdParameter = CreateObject("ADODB.Command")
	cmdParameter.CommandText = "spASRIntGetParameterPassByType"
	cmdParameter.CommandType = 4 ' Stored Procedure
        cmdParameter.ActiveConnection = Session("databaseConnection")

        prmFunctionID = cmdParameter.CreateParameter("functionID", 3, 1) ' 3=integer, 1=input
        cmdParameter.Parameters.Append(prmFunctionID)
	prmFunctionID.value = cleanNumeric(clng(Session("optionFunctionID")))

        prmParameterIndex = cmdParameter.CreateParameter("parameterIndex", 3, 1) ' 3=integer, 1=input
        cmdParameter.Parameters.Append(prmParameterIndex)
	prmParameterIndex.value = cleanNumeric(clng(Session("optionParameterIndex")))

        prmPassByType = cmdParameter.CreateParameter("passByType", 3, 2) ' 3=integer, 2=output
        cmdParameter.Parameters.Append(prmPassByType)

        Err.Clear()
	cmdParameter.Execute
        If (Err.Number <> 0) Then
            sErrMsg = "Error checking parameter pass-by type." & vbCrLf & FormatError(Err.Description)
        Else
            iPassBy = cmdParameter.Parameters("passByType").Value
        End If

	' Release the ADO command object.
        cmdParameter = Nothing
end if
    Response.Write("<INPUT type='hidden' id=txtPassByType name=txtPassByType value=" & iPassBy & ">" & vbCrLf)
%>	

<table align=center class="outline" cellPadding=5 cellSpacing=0 width=100% height=100%>
	<TR>
		<TD>
			<TABLE WIDTH="100%" height="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
				<TR height=5>
					<TD height=5 colspan=5></td>
				</tr>
				
				<tr>
					<td width=10>&nbsp;&nbsp;</td>
					
					<TD width=10%>
						<TABLE height=100% width=100% class="outline" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD valign=top>
									<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0>
										<TR height=5>
											<TD colspan=5>&nbsp;&nbsp;</TD>
										</TR>
										
										<TR height=10>
											<TD width=5>&nbsp;</TD>
											<TD width=5><STRONG>Type</STRONG></TD>
											<TD colspan=3></TD>
										</TR>
										
										<TR height=10>
											<TD colspan=5></TD>
										</TR>
										
										<TR height=10>
											<TD width=5>&nbsp;</TD>
											<TD width=5>
												<input id=optType_Field name=optType type=radio selected
												    onclick="changeType(1)" 
                                                    onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
											</TD>
											<TD width=5>&nbsp;</TD>
											<TD nowrap>
                                                <label 
                                                    tabindex=-1
                                                    for="optType_Field"
                                                    class="radio"
                                                    onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                />
    												Field
                           	    		        </label>
											</TD>
											<TD width=5>&nbsp;&nbsp;</TD>
										</TR>
										
										<TR height=5>
											<TD colspan=5></TD>
										</TR>
										
										<TR height=10>
											<TD width=5>&nbsp;</TD>
											<TD width=5>
												<input id=optType_Operator name=optType type=radio
                                                    <%
												    If iPassBy = 2 Then
												        Response.Write("disabled")												        
												    End If
												    %>
												    onclick="changeType(5)" 
                                                    onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
											</TD>
											<TD width=5>&nbsp;</TD>
											<TD nowrap>
                                                <label 
                                                    tabindex=-1
                                                    for="optType_Operator"
                                                    class="radio"
                                                    onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                />
    												Operator
                           	    		        </label>
											</TD>
											<TD width=5>&nbsp;</TD>
										</TR>

										<TR height=5>
											<TD colspan=5></TD>
										</TR>

										<TR height=10>
											<TD width=5>&nbsp;</TD>
											<TD width=5>
												<input id=optType_Function name=optType type=radio 
                                                    <%
												    If iPassBy = 2 Then											        
												        Response.Write("disabled")
												    End If
												    %>
												    onclick="changeType(2)" 
                                                    onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
											</TD>
											<TD width=5>&nbsp;</TD>
											<TD nowrap>
                                                <label 
                                                    tabindex=-1
                                                    for="optType_Function"
                                                    class="radio"
                                                    onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                />
    												Function
                           	    		        </label>
											</TD>
											<TD width=5>&nbsp;</TD>
										</TR>

										<TR height=5>
											<TD colspan=5></TD>
										</TR>

										<TR height=10>
											<TD width=5>&nbsp;</TD>
											<TD width=5>
												<input id=optType_Value name=optType type=radio <%
												    If iPassBy = 2 Then
												        Response.Write("disabled")
												    End If
												    %>
												    onclick="changeType(4)" 
                                                    onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
											</TD>
											<TD width=5>&nbsp;</TD>
											<TD nowrap>
                                                <label 
                                                    tabindex=-1
                                                    for="optType_Value"
                                                    class="radio"
                                                    onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                />
    												Value
                           	    		        </label>
											</TD>
											<TD width=5>&nbsp;</TD>
										</TR>

										<TR height=5>
											<TD colspan=5></TD>
										</TR>

										<TR height=10>
											<TD width=5>&nbsp;</TD>
											<TD width=5>
                                                <input id="optType_LookupTableValue" name="optType" type="radio" <% 
                                                    If iPassBy = 2 Then
                                                        Response.Write("disabled")
                                                    End If
                                                    %>
                                                    onclick="changeType(6)"
                                                    onmouseover="try{radio_onMouseOver(this);}catch(e){}"
                                                    onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                    onblur="try{radio_onBlur(this);}catch(e){}" />
											</TD>
											<TD width=5>&nbsp;</TD>
											<TD nowrap>
                                                <label 
                                                    tabindex=-1
                                                    for="optType_LookupTableValue"
                                                    class="radio"
                                                    onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                />
    												Lookup Table Value
                           	    		        </label>
											</TD>
											<TD width=5>&nbsp;</TD>
										</TR>

										<TR height=5>
											<TD colspan=5></TD>
										</TR>

										<TR height=10 style="visibility:hidden;display:none" id=trType_PVal>
											<TD width=5>&nbsp;</TD>
											<TD width=5>
												<input id=optType_PromptedValue name=optType type=radio <%if iPassBy = 2 then
    Response.write("disabled")
												    End If%>
												    onclick="changeType(7)" 
                                                    onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
											</TD>
											<TD width=5>&nbsp;</TD>
											<TD nowrap>
                                                <label 
                                                    tabindex=-1
                                                    for="optType_PromptedValue"
                                                    class="radio"
                                                    onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                />
    												Prompted Value
                           	    		        </label>
											</TD>
											<TD width=5>&nbsp;</TD>
										</TR>

										<TR height=5 style="visibility:hidden;display:none" id=trType_PVal2>
											<TD colspan=5></TD>
										</TR>

										<TR height=10 style="visibility:hidden;display:none" id=trType_Calc>
											<TD width=5>&nbsp;</TD>
											<TD width=5>
												<input id=optType_Calculation name=optType type=radio <%  If iPassBy = 2 Then
												        Response.Write("disabled")
												    End If%>
												    onclick="changeType(3)" 
                                                    onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
											</TD>
											<TD width=5>&nbsp;</TD>
											<TD nowrap>
                                                <label 
                                                    tabindex=-1
                                                    for="optType_Calculation"
                                                    class="radio"
                                                    onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                />
    												Calculation
                           	    		        </label>
											</TD>
											<TD width=5>&nbsp;</TD>
										</TR>

										<TR height=5 style="visibility:hidden;display:none" id=trType_Calc2>
											<TD colspan=5></TD>
										</TR>

										<TR height=10 style="visibility:hidden;display:none" id=trType_Filter>
											<TD width=5>&nbsp;</TD>
											<TD width=5>
												<input id=optType_Filter name=optType type=radio <%
												    If iPassBy = 2 Then
												        Response.Write("disabled")
												    End If

												    %>
												    onclick="changeType(10)" 
                                                    onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
											</TD>
											<TD width=5>&nbsp;</TD>
											<TD nowrap>
                                                <label 
                                                    tabindex=-1
                                                    for="optType_Filter"
                                                    class="radio"
                                                    onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                />
    												Filter
                           	    		        </label>
											</TD>
											<TD width=5>&nbsp;</TD>
										</TR>

										<TR>
											<TD colspan=5></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
								
					<td width=10>&nbsp;&nbsp;</td>

					<TD>
						<TABLE height=100% width=100% class="outline" CELLSPACING=0 CELLPADDING=0>
							<TR height=100%>
								<TD valign=top>
									<DIV id=divField>
										<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0>
											<TR height=10>
												<TD colspan=6>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD colspan=4><STRONG>Field</STRONG></TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=6></TD>
											</TR>
<%if iPassBy = 1 then%>
											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD colspan=4>
													<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
														<TR>
															<TD>
																<input id=optField_Field name=optField type=radio selected
																    onclick="field_refreshTable()" 
                                                                    onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
															</TD>
															<TD width=5>&nbsp;</TD>
															<TD nowrap>
                                                                <label 
                                                                    tabindex=-1
                                                                    for="optField_Field"
                                                                    class="radio"
                                                                    onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                />
	    														    Field
                                           	    		        </label>
    														</TD>
															<TD width=20>&nbsp;&nbsp;&nbsp;</TD>
															<TD>
																<input id=optField_Count name=optField type=radio selected
																    onclick="field_refreshTable()" 
                                                                    onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
															</TD>
															<TD width=5>&nbsp;</TD>
															<TD nowrap>
                                                                <label 
                                                                    tabindex=-1
                                                                    for="optField_Count"
                                                                    class="radio"
                                                                    onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                />
															        Count
                                           	    		        </label>
															</TD>
															<TD width=20>&nbsp;&nbsp;&nbsp;</TD>
															<TD>
																<input id=optField_Total name=optField type=radio selected
																    onclick="field_refreshTable()"
                                                                    onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
															</TD>
															<TD width=5>&nbsp;</TD>
															<TD nowrap>
                                                                <label 
                                                                    tabindex=-1
                                                                    for="optField_Total"
                                                                    class="radio"
                                                                    onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                    onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                />
															        Total
                                           	    		        </label>
															</TD>
															<TD width=100%></TD>
														</TR>
													</TABLE>
												</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=6>&nbsp;&nbsp;</TD>
											</TR>
<%end if%>

											<TR height=10>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD width=10 nowrap>Table :</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD width=50%>
													<select id="cboFieldTable" name="cboFieldTable" class="combo" style="WIDTH: 100%" 
													    onchange="field_changeTable()"> 
													</select>
												</TD>
												<TD width=50%>&nbsp;</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=5>
												<TD colspan=6></TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD width=10 nowrap>Column :</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD width=50%>
													<select id=cboFieldColumn name=cboFieldColumn class="combo" style="WIDTH: 100%"> 
													</select>
													<select id=cboFieldDummyColumn name=cboFieldDummyColumn class="combo combodisabled" style="WIDTH: 100%;visibility:hidden;display:none" disabled="disabled"> 
													</select>
												</TD>
												<TD width=50%>&nbsp;</TD>
												<TD width=10>&nbsp;&nbsp;</TD>
											</TR>

<%if iPassBy = 1 then%>
											<TR height=5>
												<TD colspan=6>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD colspan=4>
													<TABLE WIDTH=100% height=100% class="outline" CELLSPACING=0 CELLPADDING=0>
														<TR>
															<TD>
																<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
																	<TR height=10>
																		<TD colspan=6></TD>
																	</TR>

																	<TR height=10>
																		<TD width=10>&nbsp;</TD>
																		<TD colspan=4><STRONG>Child Field Options</STRONG></TD>
																		<TD width=10>&nbsp;</TD>
																	</TR>

																	<TR height=10>
																		<TD colspan=6></TD>
																	</TR>

																	<TR height=10>
																		<TD width=10>&nbsp;</TD>
																		<TD colspan=4>
																			<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
																				<TR>
																					<TD>
																						<input id=optFieldRecSel_First name=optFieldRecSel type=radio 
																						    onclick="field_refreshChildFrame()"
                                                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                            onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                            onblur="try{radio_onBlur(this);}catch(e){}"/>
																					</TD>
																					<TD width=5>&nbsp;</TD>
																					<TD nowrap>
                                                                                        <label 
                                                                                            tabindex=-1
                                                                                            for="optFieldRecSel_First"
                                                                                            class="radio"
                                                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                                        />
																					        First
                                                                   	    		        </label>
																					</TD>
																					<TD width=20>&nbsp;&nbsp;&nbsp;</TD>
																					<TD>
																						<input id=optFieldRecSel_Last name=optFieldRecSel type=radio 
																						    onclick="field_refreshChildFrame()"
                                                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                            onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                            onblur="try{radio_onBlur(this);}catch(e){}"/>
																					</TD>
																					<TD width=5>&nbsp;</TD>
																					<TD nowrap>
                                                                                        <label 
                                                                                            tabindex=-1
                                                                                            for="optFieldRecSel_Last"
                                                                                            class="radio"
                                                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                                        />
																					        Last
                                                                   	    		        </label>
																					</TD>
																					<TD width=20>&nbsp;&nbsp;&nbsp;</TD>
																					<TD>
																						<input id=optFieldRecSel_Specific name=optFieldRecSel type=radio 
																						    onclick="field_refreshChildFrame()"
                                                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                            onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                            onblur="try{radio_onBlur(this);}catch(e){}"/>
																					</TD>
																					<TD width=5>&nbsp;</TD>
																					<TD nowrap>
																						<DIV id=divFieldRecSel_Specific style="visibility:hidden;display:none">
                                                                                            <label 
                                                                                                tabindex=-1
                                                                                                for="optFieldRecSel_Specific"
                                                                                                class="radio"
                                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                                            />
		    																					Specific
                                                                       	    		        </label>
        																				</DIV>
																					</TD>
																					<TD width=5>&nbsp;</TD>
																					<TD width=100%>
																						<INPUT id=txtFieldRecSel_Specific name=txtFieldRecSel_Specific class="text">	
																					</TD>
																				</TR>
																			</TABLE>
																		</TD>
																		<TD width=10>&nbsp;</TD>
																	</TR>
																	
																	<TR height=10>
																		<TD colspan=6></TD>
																	</TR>

																	<TR height=10>
																		<TD width=20>&nbsp;&nbsp;</TD>
																		<TD width=110 nowrap>Order :</TD>
																		<TD width=20>&nbsp;&nbsp;</TD>
																		<TD width=50%>
																			<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
																				<TR>
																					<TD>
																						<INPUT type="text" id=txtFieldRecOrder name=txtFieldRecOrder class="text textdisabled" style="WIDTH: 100%" disabled="disabled">
																					</TD>
																					<TD style="width:30px;">
																						<INPUT id=btnFieldRecOrder name=btnFieldRecOrder style="WIDTH: 100%" class="btn" type=button value="..."
																						    onclick="field_selectRecOrder()" 
		                                                                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                                                    onfocus="try{button_onFocus(this);}catch(e){}"
		                                                                                    onblur="try{button_onBlur(this);}catch(e){}" />
																					</TD>
																				</TR>
																			</TABLE>
																		</TD>
																		<TD width=30%>&nbsp;</TD>
																		<TD width=10>&nbsp;&nbsp;</TD>
																	</TR>
																	
																	<TR height=5>
																		<TD colspan=6></TD>
																	</TR>

																	<TR height=10>
																		<TD width=20>&nbsp;&nbsp;</TD>
																		<TD width=110 nowrap>Filter :</TD>
																		<TD width=20>&nbsp;&nbsp;</TD>
																		<TD width=50%>
																			<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
																				<TR>
																					<TD>
																						<INPUT type="text" id=txtFieldRecFilter name=txtFieldRecFilter class="text textdisabled" style="WIDTH: 100%" disabled="disabled">
																					</TD>
																					<TD width=30>
																						<INPUT id=btnFieldRecFilter name=btnFieldRecFilter class="btn" style="WIDTH: 100%" type=button value="..."
																						    onclick="field_selectRecFilter()" 
			                                                                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
			                                                                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
			                                                                                onfocus="try{button_onFocus(this);}catch(e){}"
			                                                                                onblur="try{button_onBlur(this);}catch(e){}" />
																					</TD>
																				</TR>
																			</TABLE>
																		</TD>
																		<TD width=30%>&nbsp;</TD>
																		<TD width=10>&nbsp;&nbsp;</TD>
																	</TR>

																	<TR height=10>
																		<TD colspan=6></TD>
																	</TR>
																</TABLE>
															</TD>
														</TR>
													</TABLE>
												</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=100%>
												<TD colspan=6>&nbsp;</TD>
											</TR>
<%end if%>
										</TABLE>
									</DIV>

									<DIV id=divFunction style="visibility:hidden;display:none">
										<TABLE height=100% width=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
											<TR height=10>
												<TD colspan=3>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD><STRONG>Function</STRONG></TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=3></TD>
											</TR>

											<TR>
												<TD width=10>&nbsp;</TD>
												<TD>
													<OBJECT classid="clsid:1C203F13-95AD-11D0-A84B-00A0247B735B" id=SSFunctionTree codebase="cabs/SStree.cab#version=1,0,2,24" style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:400px" VIEWASTEXT>
														<PARAM NAME="_ExtentX" VALUE="2646">
														<PARAM NAME="_ExtentY" VALUE="1323">
														<PARAM NAME="_Version" VALUE="65538">
														<PARAM NAME="BackColor" VALUE="-2147483643">
														<PARAM NAME="ForeColor" VALUE="-2147483640">
														<PARAM NAME="ImagesMaskColor" VALUE="12632256">
														<PARAM NAME="PictureBackgroundMaskColor" VALUE="12632256">
														<PARAM NAME="Appearance" VALUE="1">
														<PARAM NAME="BorderStyle" VALUE="0">
														<PARAM NAME="LabelEdit" VALUE="1">
														<PARAM NAME="LineStyle" VALUE="0">
														<PARAM NAME="LineType" VALUE="1">
														<PARAM NAME="MousePointer" VALUE="0">
														<PARAM NAME="NodeSelectionStyle" VALUE="2">
														<PARAM NAME="PictureAlignment" VALUE="0">
														<PARAM NAME="ScrollStyle" VALUE="0">
														<PARAM NAME="Style" VALUE="6">
														<PARAM NAME="IndentationStyle" VALUE="0">
														<PARAM NAME="TreeTips" VALUE="3">
														<PARAM NAME="PictureBackgroundStyle" VALUE="0">
														<PARAM NAME="Indentation" VALUE="38">
														<PARAM NAME="MaxLines" VALUE="1">
														<PARAM NAME="TreeTipDelay" VALUE="500">
														<PARAM NAME="ImageCount" VALUE="0">
														<PARAM NAME="ImageListIndex" VALUE="-1">
														<PARAM NAME="OLEDragMode" VALUE="0">
														<PARAM NAME="OLEDropMode" VALUE="0">
														<PARAM NAME="AllowDelete" VALUE="0">
														<PARAM NAME="AutoSearch" VALUE="0">
														<PARAM NAME="Enabled" VALUE="-1">
														<PARAM NAME="HideSelection" VALUE="0">
														<PARAM NAME="ImagesUseMask" VALUE="0">
														<PARAM NAME="Redraw" VALUE="-1">
														<PARAM NAME="UseImageList" VALUE="-1">
														<PARAM NAME="PictureBackgroundUseMask" VALUE="0">
														<PARAM NAME="HasFont" VALUE="0">
														<PARAM NAME="HasMouseIcon" VALUE="0">
														<PARAM NAME="HasPictureBackground" VALUE="0">
														<PARAM NAME="PathSeparator" VALUE="\">
														<PARAM NAME="TabStops" VALUE="32">
														<PARAM NAME="ImageList" VALUE="<None>">
														<PARAM NAME="LoadStyleRoot" VALUE="1">
														<PARAM NAME="Sorted" VALUE="0">
														<PARAM NAME="OnDemandDiscardBuffer" VALUE="10">
													</OBJECT>
												</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=3>&nbsp;</TD>
											</TR>
										</TABLE>
									</DIV>
									
									<DIV id=divOperator style="visibility:hidden;display:none">
										<TABLE height=100% width=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
											<TR height=10>
												<TD colspan=3>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD><STRONG>Operator</STRONG></TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=3></TD>
											</TR>

											<TR>
												<TD width=10>&nbsp;</TD>
												<TD>
													<OBJECT classid="clsid:1C203F13-95AD-11D0-A84B-00A0247B735B" id=SSOperatorTree codebase="cabs/SStree.cab#version=1,0,2,24" style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:400px" VIEWASTEXT>
														<PARAM NAME="_ExtentX" VALUE="2646">
														<PARAM NAME="_ExtentY" VALUE="1323">
														<PARAM NAME="_Version" VALUE="65538">
														<PARAM NAME="BackColor" VALUE="-2147483643">
														<PARAM NAME="ForeColor" VALUE="-2147483640">
														<PARAM NAME="ImagesMaskColor" VALUE="12632256">
														<PARAM NAME="PictureBackgroundMaskColor" VALUE="12632256">
														<PARAM NAME="Appearance" VALUE="1">
														<PARAM NAME="BorderStyle" VALUE="0">
														<PARAM NAME="LabelEdit" VALUE="1">
														<PARAM NAME="LineStyle" VALUE="0">
														<PARAM NAME="LineType" VALUE="1">
														<PARAM NAME="MousePointer" VALUE="0">
														<PARAM NAME="NodeSelectionStyle" VALUE="2">
														<PARAM NAME="PictureAlignment" VALUE="0">
														<PARAM NAME="ScrollStyle" VALUE="0">
														<PARAM NAME="Style" VALUE="6">
														<PARAM NAME="IndentationStyle" VALUE="0">
														<PARAM NAME="TreeTips" VALUE="3">
														<PARAM NAME="PictureBackgroundStyle" VALUE="0">
														<PARAM NAME="Indentation" VALUE="38">
														<PARAM NAME="MaxLines" VALUE="1">
														<PARAM NAME="TreeTipDelay" VALUE="500">
														<PARAM NAME="ImageCount" VALUE="0">
														<PARAM NAME="ImageListIndex" VALUE="-1">
														<PARAM NAME="OLEDragMode" VALUE="0">
														<PARAM NAME="OLEDropMode" VALUE="0">
														<PARAM NAME="AllowDelete" VALUE="0">
														<PARAM NAME="AutoSearch" VALUE="0">
														<PARAM NAME="Enabled" VALUE="-1">
														<PARAM NAME="HideSelection" VALUE="0">
														<PARAM NAME="ImagesUseMask" VALUE="0">
														<PARAM NAME="Redraw" VALUE="-1">
														<PARAM NAME="UseImageList" VALUE="-1">
														<PARAM NAME="PictureBackgroundUseMask" VALUE="0">
														<PARAM NAME="HasFont" VALUE="0">
														<PARAM NAME="HasMouseIcon" VALUE="0">
														<PARAM NAME="HasPictureBackground" VALUE="0">
														<PARAM NAME="PathSeparator" VALUE="\">
														<PARAM NAME="TabStops" VALUE="32">
														<PARAM NAME="ImageList" VALUE="<None>">
														<PARAM NAME="LoadStyleRoot" VALUE="1">
														<PARAM NAME="Sorted" VALUE="0">
														<PARAM NAME="OnDemandDiscardBuffer" VALUE="10">
													</OBJECT>
												</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=3>&nbsp;</TD>
											</TR>
										</TABLE>
									</DIV>

									<DIV id=divValue style="visibility:hidden;display:none">
										<TABLE height=100% width=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
											<TR height=10>
												<TD colspan=6>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD colspan=4><STRONG>Value</STRONG></TD>
												<TD width=20>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=6></TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD width=10 nowrap>Type :</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD width=50%>
													<select id=cboValueType name=cboValueType class="combo" style="WIDTH: 100%" onchange="value_changeType()"> 
														<OPTION value=1>Character
														<OPTION value=2>Numeric
														<OPTION value=3>Logic
														<OPTION value=4>Date
													</select>
												</TD>
												<TD width=50%>&nbsp;</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=5>
												<TD colspan=6></TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD width=10 nowrap>Value :</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD width=50%>
													<SELECT id=selectValue name='selectValue"' class="combo" style="WIDTH: 100%">
														<OPTION value=1>True</OPTION>
														<OPTION value=0>False</OPTION>
													</SELECT>
													<INPUT id=txtValue name=txtValue class="text" style="LEFT: 0px; POSITION: absolute; TOP: 0px; VISIBILITY: hidden; WIDTH: 0px">	
												</TD>
												<TD width=50%>&nbsp;</TD>
												<TD width=10>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=100%>
												<TD colspan=6>&nbsp;</TD>
											</TR>
										</TABLE>
									</DIV>

									<DIV id=divLookupValue style="visibility:hidden;display:none">
										<TABLE height=100% width=100% class="invisible"  CELLSPACING=0 CELLPADDING=0>
											<TR height=10>
												<TD colspan=6>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD colspan=4><STRONG>Lookup Table Value</STRONG></TD>
												<TD width=20>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=6></TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD width=10 nowrap>Table :</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD width=50%>
													<select id=cboLookupValueTable name=cboLookupValueTable class="combo" style="WIDTH: 100%" onchange="lookupValue_changeTable()"> 
													</select>
												</TD>
												<TD width=50%>&nbsp;</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=5>
												<TD colspan=6></TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD width=10 nowrap>Column :</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD width=50%>
													<select id=cboLookupValueColumn name=cboLookupValueColumn class="combo" style="WIDTH: 100%" onchange="lookupValue_changeColumn()"> 
													</select>
												</TD>
												<TD width=50%>&nbsp;</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=5>
												<TD colspan=6></TD>
											</TR>
											
											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD width=10 nowrap>Value :</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD width=50%>
													<select id=cboLookupValueValue name=cboLookupValueValue class="combo" style="WIDTH: 100%"> 
													</select>
												</TD>
												<TD width=50%>&nbsp;</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=6></TD>
											</TR>
											
											<TR height=10>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD colspan=4>
													<input type=text class="textwarning" id=txtValueNotInLookup name=txtValueNotInLookup value="<value> does not appear in <table>.<column>" style ="TEXT-ALIGN: left; WIDTH: 100%; visibility:hidden; display:none" readonly>
												</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=100%>
												<TD colspan=6>&nbsp;</TD>
											</TR>
										</TABLE>
									</DIV>

									<DIV id=divCalculation style="visibility:hidden;display:none">
										<TABLE height=100% width=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
											<TR height=10>
												<TD colspan=3>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD><STRONG>Calculation</STRONG></TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=3></TD>
											</TR>

											<TR>
												<TD width=10>&nbsp;</TD>
												<TD>
													<OBJECT classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" id=ssOleDBGridCalculations name=ssOleDBGridCalculations codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:400px">
														<PARAM NAME="ScrollBars" VALUE="4">
														<PARAM NAME="_Version" VALUE="196616">
														<PARAM NAME="DataMode" VALUE="2">
														<PARAM NAME="Cols" VALUE="0">
														<PARAM NAME="Rows" VALUE="0">
														<PARAM NAME="BorderStyle" VALUE="1">
														<PARAM NAME="RecordSelectors" VALUE="0">
														<PARAM NAME="GroupHeaders" VALUE="0">
														<PARAM NAME="ColumnHeaders" VALUE="0">
														<PARAM NAME="GroupHeadLines" VALUE="0">
														<PARAM NAME="HeadLines" VALUE="0">
														<PARAM NAME="FieldDelimiter" VALUE="(None)">
														<PARAM NAME="FieldSeparator" VALUE="(Tab)">
														<PARAM NAME="Col.Count" VALUE="2">
														<PARAM NAME="stylesets.count" VALUE="0">
														<PARAM NAME="TagVariant" VALUE="EMPTY">
														<PARAM NAME="UseGroups" VALUE="0">
														<PARAM NAME="HeadFont3D" VALUE="0">
														<PARAM NAME="Font3D" VALUE="0">
														<PARAM NAME="DividerType" VALUE="3">
														<PARAM NAME="DividerStyle" VALUE="1">
														<PARAM NAME="DefColWidth" VALUE="0">
														<PARAM NAME="BeveColorScheme" VALUE="2">
														<PARAM NAME="BevelColorFrame" VALUE="-2147483642">
														<PARAM NAME="BevelColorHighlight" VALUE="-2147483628">
														<PARAM NAME="BevelColorShadow" VALUE="-2147483632">
														<PARAM NAME="BevelColorFace" VALUE="-2147483633">
														<PARAM NAME="CheckBox3D" VALUE="-1">
														<PARAM NAME="AllowAddNew" VALUE="0">
														<PARAM NAME="AllowDelete" VALUE="0">
														<PARAM NAME="AllowUpdate" VALUE="0">
														<PARAM NAME="MultiLine" VALUE="0">
														<PARAM NAME="ActiveCellStyleSet" VALUE="">
														<PARAM NAME="RowSelectionStyle" VALUE="0">
														<PARAM NAME="AllowRowSizing" VALUE="0">
														<PARAM NAME="AllowGroupSizing" VALUE="0">
														<PARAM NAME="AllowColumnSizing" VALUE="0">
														<PARAM NAME="AllowGroupMoving" VALUE="0">
														<PARAM NAME="AllowColumnMoving" VALUE="0">
														<PARAM NAME="AllowGroupSwapping" VALUE="0">
														<PARAM NAME="AllowColumnSwapping" VALUE="0">
														<PARAM NAME="AllowGroupShrinking" VALUE="0">
														<PARAM NAME="AllowColumnShrinking" VALUE="0">
														<PARAM NAME="AllowDragDrop" VALUE="0">
														<PARAM NAME="UseExactRowCount" VALUE="-1">
														<PARAM NAME="SelectTypeCol" VALUE="0">
														<PARAM NAME="SelectTypeRow" VALUE="1">
														<PARAM NAME="SelectByCell" VALUE="-1">
														<PARAM NAME="BalloonHelp" VALUE="0">
														<PARAM NAME="RowNavigation" VALUE="1">
														<PARAM NAME="CellNavigation" VALUE="0">
														<PARAM NAME="MaxSelectedRows" VALUE="1">
														<PARAM NAME="HeadStyleSet" VALUE="">
														<PARAM NAME="StyleSet" VALUE="">
														<PARAM NAME="ForeColorEven" VALUE="0">
														<PARAM NAME="ForeColorOdd" VALUE="0">
														<PARAM NAME="BackColorEven" VALUE="16777215">
														<PARAM NAME="BackColorOdd" VALUE="16777215">
														<PARAM NAME="Levels" VALUE="1">
														<PARAM NAME="RowHeight" VALUE="503">
														<PARAM NAME="ExtraHeight" VALUE="0">
														<PARAM NAME="ActiveRowStyleSet" VALUE="">
														<PARAM NAME="CaptionAlignment" VALUE="2">
														<PARAM NAME="SplitterPos" VALUE="0">
														<PARAM NAME="SplitterVisible" VALUE="0">
														<PARAM NAME="Columns.Count" VALUE="2">
														<PARAM NAME="Columns(0).Width" VALUE="100000">
														<PARAM NAME="Columns(0).Visible" VALUE="-1">
														<PARAM NAME="Columns(0).Columns.Count" VALUE="1">
														<PARAM NAME="Columns(0).Caption" VALUE="Name">
														<PARAM NAME="Columns(0).Name" VALUE="Name">
														<PARAM NAME="Columns(0).Alignment" VALUE="0">
														<PARAM NAME="Columns(0).CaptionAlignment" VALUE="3">
														<PARAM NAME="Columns(0).Bound" VALUE="0">
														<PARAM NAME="Columns(0).AllowSizing" VALUE="1">
														<PARAM NAME="Columns(0).DataField" VALUE="Column 0">
														<PARAM NAME="Columns(0).DataType" VALUE="8">
														<PARAM NAME="Columns(0).Level" VALUE="0">
														<PARAM NAME="Columns(0).NumberFormat" VALUE="">
														<PARAM NAME="Columns(0).Case" VALUE="0">
														<PARAM NAME="Columns(0).FieldLen" VALUE="256">
														<PARAM NAME="Columns(0).VertScrollBar" VALUE="0">
														<PARAM NAME="Columns(0).Locked" VALUE="0">
														<PARAM NAME="Columns(0).Style" VALUE="0">
														<PARAM NAME="Columns(0).ButtonsAlways" VALUE="0">
														<PARAM NAME="Columns(0).RowCount" VALUE="0">
														<PARAM NAME="Columns(0).ColCount" VALUE="1">
														<PARAM NAME="Columns(0).HasHeadForeColor" VALUE="0">
														<PARAM NAME="Columns(0).HasHeadBackColor" VALUE="0">
														<PARAM NAME="Columns(0).HasForeColor" VALUE="0">
														<PARAM NAME="Columns(0).HasBackColor" VALUE="0">
														<PARAM NAME="Columns(0).HeadForeColor" VALUE="0">
														<PARAM NAME="Columns(0).HeadBackColor" VALUE="0">
														<PARAM NAME="Columns(0).ForeColor" VALUE="0">
														<PARAM NAME="Columns(0).BackColor" VALUE="0">
														<PARAM NAME="Columns(0).HeadStyleSet" VALUE="">
														<PARAM NAME="Columns(0).StyleSet" VALUE="">
														<PARAM NAME="Columns(0).Nullable" VALUE="1">
														<PARAM NAME="Columns(0).Mask" VALUE="">
														<PARAM NAME="Columns(0).PromptInclude" VALUE="0">
														<PARAM NAME="Columns(0).ClipMode" VALUE="0">
														<PARAM NAME="Columns(0).PromptChar" VALUE="95">
														<PARAM NAME="Columns(1).Width" VALUE="0">
														<PARAM NAME="Columns(1).Visible" VALUE="0">
														<PARAM NAME="Columns(1).Columns.Count" VALUE="1">
														<PARAM NAME="Columns(1).Caption" VALUE="id">
														<PARAM NAME="Columns(1).Name" VALUE="id">
														<PARAM NAME="Columns(1).Alignment" VALUE="0">
														<PARAM NAME="Columns(1).CaptionAlignment" VALUE="3">
														<PARAM NAME="Columns(1).Bound" VALUE="0">
														<PARAM NAME="Columns(1).AllowSizing" VALUE="1">
														<PARAM NAME="Columns(1).DataField" VALUE="Column 1">
														<PARAM NAME="Columns(1).DataType" VALUE="8">
														<PARAM NAME="Columns(1).Level" VALUE="0">
														<PARAM NAME="Columns(1).NumberFormat" VALUE="">
														<PARAM NAME="Columns(1).Case" VALUE="0">
														<PARAM NAME="Columns(1).FieldLen" VALUE="256">
														<PARAM NAME="Columns(1).VertScrollBar" VALUE="0">
														<PARAM NAME="Columns(1).Locked" VALUE="0">
														<PARAM NAME="Columns(1).Style" VALUE="0">
														<PARAM NAME="Columns(1).ButtonsAlways" VALUE="0">
														<PARAM NAME="Columns(1).RowCount" VALUE="0">
														<PARAM NAME="Columns(1).ColCount" VALUE="1">
														<PARAM NAME="Columns(1).HasHeadForeColor" VALUE="0">
														<PARAM NAME="Columns(1).HasHeadBackColor" VALUE="0">
														<PARAM NAME="Columns(1).HasForeColor" VALUE="0">
														<PARAM NAME="Columns(1).HasBackColor" VALUE="0">
														<PARAM NAME="Columns(1).HeadForeColor" VALUE="0">
														<PARAM NAME="Columns(1).HeadBackColor" VALUE="0">
														<PARAM NAME="Columns(1).ForeColor" VALUE="0">
														<PARAM NAME="Columns(1).BackColor" VALUE="0">
														<PARAM NAME="Columns(1).HeadStyleSet" VALUE="">
														<PARAM NAME="Columns(1).StyleSet" VALUE="">
														<PARAM NAME="Columns(1).Nullable" VALUE="1">
														<PARAM NAME="Columns(1).Mask" VALUE="">
														<PARAM NAME="Columns(1).PromptInclude" VALUE="0">
														<PARAM NAME="Columns(1).ClipMode" VALUE="0">
														<PARAM NAME="Columns(1).PromptChar" VALUE="95">
														<PARAM NAME="UseDefaults" VALUE="-1">
														<PARAM NAME="TabNavigation" VALUE="1">
														<PARAM NAME="_ExtentX" VALUE="17330">
														<PARAM NAME="_ExtentY" VALUE="1323">
														<PARAM NAME="_StockProps" VALUE="79">
														<PARAM NAME="Caption" VALUE="">
														<PARAM NAME="ForeColor" VALUE="0">
														<PARAM NAME="BackColor" VALUE="16777215">
														<PARAM NAME="Enabled" VALUE="-1">
														<PARAM NAME="DataMember" VALUE="">
														<PARAM NAME="Row.Count" VALUE="0">
													</OBJECT>
												</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<tr> 
											  <td colspan=3 height=10></td>
											</tr>

											<tr> 
												<TD width=10>&nbsp;</TD>
											    <td height="60"> 
													<TEXTAREA id=txtCalcDescription name=txtCalcDescription class="textarea disabled" tabindex="-1"
													    style="HEIGHT: 99%; WIDTH: 100%; " wrap=VIRTUAL disabled="disabled">
													</TEXTAREA>
												</td>
												<TD width=10>&nbsp;</TD>
											</tr>

											<tr> 
											  <td colspan=3 height=10></td>
											</tr>

											<tr> 
												<TD width=10>&nbsp;</TD>
											    <td height="10"> 
											    <input <% If Session("OnlyMine") Then Response.Write("checked")%> type="checkbox" name="chkOwnersCalcs" id="chkOwnersCalcs" value="chkOwnersCalcs" tabindex="-1"
                                                    onclick="calculationAndFilter_refresh();"
                                                    onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
                                                    onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                <label 
			                                        for="chkOwnersCalcs"
			                                        class="checkbox"
			                                        tabindex=0 
			                                        onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
				                                    onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
				                                    onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
	                                                onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
	                                                onblur="try{checkboxLabel_onBlur(this);}catch(e){}">

                                                    Only show calculations where owner is '<% =session("Username") %>'
                    		    		        </label>
												</td>
												<TD width=10>&nbsp;</TD>
											</tr>

											<TR height=10>
												<TD colspan=3>&nbsp;</TD>
											</TR>
										</TABLE>
									</DIV>
									
									<DIV id=divFilter style="visibility:hidden;display:none">
										<TABLE height=100% width=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
											<TR height=10>
												<TD colspan=3>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD><STRONG>Filter</STRONG></TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=3></TD>
											</TR>

											<TR>
												<TD width=10>&nbsp;</TD>
												<TD>
													<OBJECT classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" id=ssOleDBGridFilters name=ssOleDBGridFilters   codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:400px">
														<PARAM NAME="ScrollBars" VALUE="4">
														<PARAM NAME="_Version" VALUE="196616">
														<PARAM NAME="DataMode" VALUE="2">
														<PARAM NAME="Cols" VALUE="0">
														<PARAM NAME="Rows" VALUE="0">
														<PARAM NAME="BorderStyle" VALUE="1">
														<PARAM NAME="RecordSelectors" VALUE="0">
														<PARAM NAME="GroupHeaders" VALUE="0">
														<PARAM NAME="ColumnHeaders" VALUE="0">
														<PARAM NAME="GroupHeadLines" VALUE="0">
														<PARAM NAME="HeadLines" VALUE="0">
														<PARAM NAME="FieldDelimiter" VALUE="(None)">
														<PARAM NAME="FieldSeparator" VALUE="(Tab)">
														<PARAM NAME="Col.Count" VALUE="2">
														<PARAM NAME="stylesets.count" VALUE="0">
														<PARAM NAME="TagVariant" VALUE="EMPTY">
														<PARAM NAME="UseGroups" VALUE="0">
														<PARAM NAME="HeadFont3D" VALUE="0">
														<PARAM NAME="Font3D" VALUE="0">
														<PARAM NAME="DividerType" VALUE="3">
														<PARAM NAME="DividerStyle" VALUE="1">
														<PARAM NAME="DefColWidth" VALUE="0">
														<PARAM NAME="BeveColorScheme" VALUE="2">
														<PARAM NAME="BevelColorFrame" VALUE="-2147483642">
														<PARAM NAME="BevelColorHighlight" VALUE="-2147483628">
														<PARAM NAME="BevelColorShadow" VALUE="-2147483632">
														<PARAM NAME="BevelColorFace" VALUE="-2147483633">
														<PARAM NAME="CheckBox3D" VALUE="-1">
														<PARAM NAME="AllowAddNew" VALUE="0">
														<PARAM NAME="AllowDelete" VALUE="0">
														<PARAM NAME="AllowUpdate" VALUE="0">
														<PARAM NAME="MultiLine" VALUE="0">
														<PARAM NAME="ActiveCellStyleSet" VALUE="">
														<PARAM NAME="RowSelectionStyle" VALUE="0">
														<PARAM NAME="AllowRowSizing" VALUE="0">
														<PARAM NAME="AllowGroupSizing" VALUE="0">
														<PARAM NAME="AllowColumnSizing" VALUE="0">
														<PARAM NAME="AllowGroupMoving" VALUE="0">
														<PARAM NAME="AllowColumnMoving" VALUE="0">
														<PARAM NAME="AllowGroupSwapping" VALUE="0">
														<PARAM NAME="AllowColumnSwapping" VALUE="0">
														<PARAM NAME="AllowGroupShrinking" VALUE="0">
														<PARAM NAME="AllowColumnShrinking" VALUE="0">
														<PARAM NAME="AllowDragDrop" VALUE="0">
														<PARAM NAME="UseExactRowCount" VALUE="-1">
														<PARAM NAME="SelectTypeCol" VALUE="0">
														<PARAM NAME="SelectTypeRow" VALUE="1">
														<PARAM NAME="SelectByCell" VALUE="-1">
														<PARAM NAME="BalloonHelp" VALUE="0">
														<PARAM NAME="RowNavigation" VALUE="1">
														<PARAM NAME="CellNavigation" VALUE="0">
														<PARAM NAME="MaxSelectedRows" VALUE="1">
														<PARAM NAME="HeadStyleSet" VALUE="">
														<PARAM NAME="StyleSet" VALUE="">
														<PARAM NAME="ForeColorEven" VALUE="0">
														<PARAM NAME="ForeColorOdd" VALUE="0">
														<PARAM NAME="BackColorEven" VALUE="16777215">
														<PARAM NAME="BackColorOdd" VALUE="16777215">
														<PARAM NAME="Levels" VALUE="1">
														<PARAM NAME="RowHeight" VALUE="503">
														<PARAM NAME="ExtraHeight" VALUE="0">
														<PARAM NAME="ActiveRowStyleSet" VALUE="">
														<PARAM NAME="CaptionAlignment" VALUE="2">
														<PARAM NAME="SplitterPos" VALUE="0">
														<PARAM NAME="SplitterVisible" VALUE="0">
														<PARAM NAME="Columns.Count" VALUE="2">
														<PARAM NAME="Columns(0).Width" VALUE="100000">
														<PARAM NAME="Columns(0).Visible" VALUE="-1">
														<PARAM NAME="Columns(0).Columns.Count" VALUE="1">
														<PARAM NAME="Columns(0).Caption" VALUE="Name">
														<PARAM NAME="Columns(0).Name" VALUE="Name">
														<PARAM NAME="Columns(0).Alignment" VALUE="0">
														<PARAM NAME="Columns(0).CaptionAlignment" VALUE="3">
														<PARAM NAME="Columns(0).Bound" VALUE="0">
														<PARAM NAME="Columns(0).AllowSizing" VALUE="1">
														<PARAM NAME="Columns(0).DataField" VALUE="Column 0">
														<PARAM NAME="Columns(0).DataType" VALUE="8">
														<PARAM NAME="Columns(0).Level" VALUE="0">
														<PARAM NAME="Columns(0).NumberFormat" VALUE="">
														<PARAM NAME="Columns(0).Case" VALUE="0">
														<PARAM NAME="Columns(0).FieldLen" VALUE="256">
														<PARAM NAME="Columns(0).VertScrollBar" VALUE="0">
														<PARAM NAME="Columns(0).Locked" VALUE="0">
														<PARAM NAME="Columns(0).Style" VALUE="0">
														<PARAM NAME="Columns(0).ButtonsAlways" VALUE="0">
														<PARAM NAME="Columns(0).RowCount" VALUE="0">
														<PARAM NAME="Columns(0).ColCount" VALUE="1">
														<PARAM NAME="Columns(0).HasHeadForeColor" VALUE="0">
														<PARAM NAME="Columns(0).HasHeadBackColor" VALUE="0">
														<PARAM NAME="Columns(0).HasForeColor" VALUE="0">
														<PARAM NAME="Columns(0).HasBackColor" VALUE="0">
														<PARAM NAME="Columns(0).HeadForeColor" VALUE="0">
														<PARAM NAME="Columns(0).HeadBackColor" VALUE="0">
														<PARAM NAME="Columns(0).ForeColor" VALUE="0">
														<PARAM NAME="Columns(0).BackColor" VALUE="0">
														<PARAM NAME="Columns(0).HeadStyleSet" VALUE="">
														<PARAM NAME="Columns(0).StyleSet" VALUE="">
														<PARAM NAME="Columns(0).Nullable" VALUE="1">
														<PARAM NAME="Columns(0).Mask" VALUE="">
														<PARAM NAME="Columns(0).PromptInclude" VALUE="0">
														<PARAM NAME="Columns(0).ClipMode" VALUE="0">
														<PARAM NAME="Columns(0).PromptChar" VALUE="95">
														<PARAM NAME="Columns(1).Width" VALUE="0">
														<PARAM NAME="Columns(1).Visible" VALUE="0">
														<PARAM NAME="Columns(1).Columns.Count" VALUE="1">
														<PARAM NAME="Columns(1).Caption" VALUE="id">
														<PARAM NAME="Columns(1).Name" VALUE="id">
														<PARAM NAME="Columns(1).Alignment" VALUE="0">
														<PARAM NAME="Columns(1).CaptionAlignment" VALUE="3">
														<PARAM NAME="Columns(1).Bound" VALUE="0">
														<PARAM NAME="Columns(1).AllowSizing" VALUE="1">
														<PARAM NAME="Columns(1).DataField" VALUE="Column 1">
														<PARAM NAME="Columns(1).DataType" VALUE="8">
														<PARAM NAME="Columns(1).Level" VALUE="0">
														<PARAM NAME="Columns(1).NumberFormat" VALUE="">
														<PARAM NAME="Columns(1).Case" VALUE="0">
														<PARAM NAME="Columns(1).FieldLen" VALUE="256">
														<PARAM NAME="Columns(1).VertScrollBar" VALUE="0">
														<PARAM NAME="Columns(1).Locked" VALUE="0">
														<PARAM NAME="Columns(1).Style" VALUE="0">
														<PARAM NAME="Columns(1).ButtonsAlways" VALUE="0">
														<PARAM NAME="Columns(1).RowCount" VALUE="0">
														<PARAM NAME="Columns(1).ColCount" VALUE="1">
														<PARAM NAME="Columns(1).HasHeadForeColor" VALUE="0">
														<PARAM NAME="Columns(1).HasHeadBackColor" VALUE="0">
														<PARAM NAME="Columns(1).HasForeColor" VALUE="0">
														<PARAM NAME="Columns(1).HasBackColor" VALUE="0">
														<PARAM NAME="Columns(1).HeadForeColor" VALUE="0">
														<PARAM NAME="Columns(1).HeadBackColor" VALUE="0">
														<PARAM NAME="Columns(1).ForeColor" VALUE="0">
														<PARAM NAME="Columns(1).BackColor" VALUE="0">
														<PARAM NAME="Columns(1).HeadStyleSet" VALUE="">
														<PARAM NAME="Columns(1).StyleSet" VALUE="">
														<PARAM NAME="Columns(1).Nullable" VALUE="1">
														<PARAM NAME="Columns(1).Mask" VALUE="">
														<PARAM NAME="Columns(1).PromptInclude" VALUE="0">
														<PARAM NAME="Columns(1).ClipMode" VALUE="0">
														<PARAM NAME="Columns(1).PromptChar" VALUE="95">
														<PARAM NAME="UseDefaults" VALUE="-1">
														<PARAM NAME="TabNavigation" VALUE="1">
														<PARAM NAME="_ExtentX" VALUE="17330">
														<PARAM NAME="_ExtentY" VALUE="1323">
														<PARAM NAME="_StockProps" VALUE="79">
														<PARAM NAME="Caption" VALUE="">
														<PARAM NAME="ForeColor" VALUE="0">
														<PARAM NAME="BackColor" VALUE="16777215">
														<PARAM NAME="Enabled" VALUE="-1">
														<PARAM NAME="DataMember" VALUE="">
														<PARAM NAME="Row.Count" VALUE="0">
													</OBJECT>
												</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<tr> 
											  <td colspan=3 height=10></td>
											</tr>

											<tr> 
												<TD width=10>&nbsp;</TD>
											    <td height="60"> 
													<TEXTAREA id=txtFilterDescription name=txtFilterDescription class="textarea disabled" tabindex="-1"
													style="HEIGHT: 99%; WIDTH: 100%; " wrap=VIRTUAL disabled="disabled">
													</TEXTAREA>
												</td>
												<TD width=10>&nbsp;</TD>
											</tr>

											<tr> 
											  <td colspan=3 height=10></td>
											</tr>

											<tr>
                                                <td width="10">&nbsp;</td>
                                                <td height="10">
                                                    <input <% If Session("OnlyMine") Then Response.Write("checked")%> type="checkbox" name="chkOwnersFilters" id="chkOwnersFilters" value="chkOwnersFilters" tabindex="-1"
                                                        onclick="calculationAndFilter_refresh();"
                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                    <label
                                                        for="chkOwnersFilters"
                                                        class="checkbox"
                                                        tabindex="0"
                                                        onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                                        onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
                                                        onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
                                                        Only show filters where owner is '<% =session("Username") %>'
                                                    </label>
												</td>
												<TD width=10>&nbsp;</TD>
											</tr>

											<TR height=10>
												<TD colspan=3>&nbsp;</TD>
											</TR>
										</TABLE>
									</DIV>

									<DIV id=divPromptedValue style="visibility:hidden;display:none">
										<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0>
											<TR height=10>
												<TD colspan=6>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD colspan=4><STRONG>Prompted Value</STRONG></TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD colspan=6></TD>
											</TR>

											<TR height=10>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD width=10 nowrap>Prompt :</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
												<TD width=50%>
													<INPUT id=txtPrompt name=txtPrompt class="text" onkeyup="pVal_changePrompt()" style="WIDTH: 100%" maxlength=40>	
												</TD>
												<TD width=50%>&nbsp;</TD>
												<TD width=10>&nbsp;</TD>
											</TR>

											<TR height=5>
												<TD colspan=6>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD colspan=4>
												<TABLE WIDTH=100% height=100% class="outline" CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD>
															<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
																<TR height=5>
																	<TD colspan=11></TD>
																</TR>

																<TR height=10>
																	<TD width=10>&nbsp;</TD>
																	<TD colspan=9><STRONG>Type</STRONG></TD>
																	<TD width=10>&nbsp;</TD>
																</TR>

																<TR height=10>
																	<TD colspan=11></TD>
																</TR>

																<TR height=10>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=40%>
																		<select id=cboPValType name=cboPValType class="combo" style="WIDTH: 100%" onchange="pVal_changeType()"> 
																			<OPTION value=1>Character
																			<OPTION value=2>Numeric
																			<OPTION value=3>Logic
																			<OPTION value=4>Date
																			<OPTION value=5>Lookup Value
																		</select>
																	</TD>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=10 nowrap id=tdPValSizePrompt name=tdPValSizePrompt>
																		Size :
																	</TD>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=30%>
																		<INPUT class="text" id=txtPValSize name=txtPValSize style="WIDTH: 100%">	
																	</TD>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=10 nowrap id=tdPValDecimalsPrompt name=tdPValDecimalsPrompt>
																		Decimals :
																	</TD>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=30%>
																		<INPUT class="text" id=txtPValDecimals name=txtPValDecimals style="WIDTH: 100%">	
																	</TD>
																	<TD width=10>&nbsp;&nbsp;</TD>
																</TR>
																
																<TR height=10>
																	<TD colspan=11></TD>
																</TR>
															</TABLE>
														</TD>
													</TR>
												</TABLE>

												</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=5>
												<TD colspan=6>&nbsp;</TD>
											</TR>

											<TR height=10 id=trPValFormat>
												<TD width=10>&nbsp;</TD>
												<TD colspan=4>
												<TABLE WIDTH=100% height=100% class="outline" CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD>
															<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
																<TR height=5>
																	<TD colspan=8></TD>
																</TR>

																<TR height=10>
																	<TD width=10>&nbsp;</TD>
																	<TD colspan=6><STRONG>Mask</STRONG></TD>
																	<TD width=10>&nbsp;</TD>
																</TR>

																<TR height=10>
																	<TD colspan=8></TD>
																</TR>

																<TR height=10>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=100% colspan=6>
																		<INPUT id=txtPValFormat name=txtPValFormat class="text" style="WIDTH: 100%">	
																	</TD>
																	<TD width=20>&nbsp;&nbsp;</TD>
																</TR>

																<TR height=5>
																	<TD colspan=8></TD>
																</TR>

																<TR height=10>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD nowrap width=5%>A - Uppercase</TD>
																	<TD width=20%>&nbsp;&nbsp;</TD>
																	<TD nowrap width=5%>9 - Numbers (0-9)</TD>
																	<TD width=20%>&nbsp;&nbsp;</TD>
																	<TD nowrap width=5%>B - Binary (0 or 1)</TD>
																	<TD></TD>
																</TR>

																<TR height=5>
																	<TD colspan=8></TD>
																</TR>

																<TR height=10>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD nowrap width=5%>a - Lowercase</TD>
																	<TD width=20%>&nbsp;&nbsp;</TD>
																	<TD nowrap width=5%># - Numbers, Symbols</TD>
																	<TD width=20%>&nbsp;&nbsp;</TD>
																	<TD nowrap width=5%>\ - Follow by any literal</TD>
																	<TD></TD>
																</TR>

																<TR height=10>
																	<TD colspan=11></TD>
																</TR>
															</TABLE>
														</TD>
													</TR>
												</TABLE>

												</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=5 id=trPValFormat2>
												<TD colspan=6>&nbsp;</TD>
											</TR>

											<TR height=10 id=trPValLookup style="visibility:hidden;display:none">
												<TD width=10>&nbsp;</TD>
												<TD colspan=4>
												<TABLE WIDTH=100% height=100% class="outline" CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD>
															<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
																<TR height=5>
																	<TD colspan=6></TD>
																</TR>

																<TR height=10>
																	<TD width=10>&nbsp;</TD>
																	<TD colspan=4><STRONG>Lookup Table Value</STRONG></TD>
																	<TD width=10>&nbsp;</TD>
																</TR>

																<TR height=10>
																	<TD colspan=6></TD>
																</TR>

																<TR height=10>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=10 nowrap>Table :</TD>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=50%>
																		<select id=cboPValTable name=cboPValTable class="combo" style="WIDTH: 100%" onchange="pVal_changeTable()"> 
																		</select>
																	</TD>
																	<TD width=50%>&nbsp;</TD>
																	<TD width=10>&nbsp;</TD>
																</TR>

																<TR height=5>
																	<TD colspan=6></TD>
																</TR>

																<TR height=10>
																	<TD width=10>&nbsp;</TD>
																	<TD width=10 nowrap>Column :</TD>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=50%>
																		<select id=cboPValColumn name=cboPValColumn class="combo" style="WIDTH: 100%" onchange="pVal_changeColumn()"> 
																		</select>
																	</TD>
																	<TD width=50%>&nbsp;</TD>
																	<TD width=10>&nbsp;&nbsp;</TD>
																</TR>

																<TR height=10>
																	<TD colspan=6></TD>
																</TR>
															</TABLE>
														</TD>
													</TR>
												</TABLE>

												</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=5 id=trPValLookup2>
												<TD colspan=6>&nbsp;</TD>
											</TR>

											<TR height=10>
												<TD width=10>&nbsp;</TD>
												<TD colspan=4>
												<TABLE WIDTH=100% height=100% class="outline" CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD>
															<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
																<TR height=5>
																	<TD colspan=8></TD>
																</TR>

																<TR height=10>
																	<TD width=10>&nbsp;</TD>
																	<TD colspan=6><STRONG>Default Value</STRONG></TD>
																	<TD width=10>&nbsp;</TD>
																</TR>

																<TR height=10>
																	<TD colspan=8></TD>
																</TR>

																<TR height=10 id=trPValDateOptions>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=100% colspan=6>
																		<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
																			<TR height=10>
																				<TD width=5>
																					<input id=optPValDate_Explicit name=optPValDate type=radio selected
																					    onclick="pVal_changeDateOption(0)" 
		                                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                        onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                        onblur="try{radio_onBlur(this);}catch(e){}"/>
																				</TD>
																				<TD width=5>&nbsp;</TD>
																				<TD nowrap>
                                                                                    <label 
                                                                                        tabindex=-1
	                                                                                    for="optPValDate_Explicit"
	                                                                                    class="radio"
		                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                                            />
	    																				Explicit
                                                            	    		        </label>
    																			</TD>
																				<TD width=20>&nbsp;&nbsp;</TD>
																				<TD width=5>
																					<input id=optPValDate_MonthStart name=optPValDate type=radio 
																					    onclick="pVal_changeDateOption(2)"
		                                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                        onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                        onblur="try{radio_onBlur(this);}catch(e){}"/>
																				</TD>
																				<TD width=5>&nbsp;</TD>
																				<TD nowrap>
                                                                                    <label 
                                                                                        tabindex=-1
	                                                                                    for="optPValDate_MonthStart"
	                                                                                    class="radio"
		                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                                            />
    																					Month Start
                                                            	    		        </label>
																				</TD>
																				<TD width=20>&nbsp;&nbsp;</TD>
																				<TD width=5>
																					<input id=optPValDate_YearStart name=optPValDate type=radio 
																					    onclick="pVal_changeDateOption(4)"
		                                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                        onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                        onblur="try{radio_onBlur(this);}catch(e){}"/>
																				</TD>
																				<TD width=5>&nbsp;</TD>
																				<TD nowrap>
                                                                                    <label 
                                                                                        tabindex=-1
	                                                                                    for="optPValDate_YearStart"
	                                                                                    class="radio"
		                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                                            />
    																					Year Start
                                                            	    		        </label>
																				</TD>
																				<TD width=100%>&nbsp;</TD>
																			</TR>
																			
																			<TR height=10>
																				<TD width=5>
																					<input id=optPValDate_Current name=optPValDate type=radio 
																					    onclick="pVal_changeDateOption(1)"
		                                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                        onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                        onblur="try{radio_onBlur(this);}catch(e){}"/>
																				</TD>
																				<TD width=5>&nbsp;</TD>
																				<TD nowrap>
                                                                                    <label 
                                                                                        tabindex=-1
	                                                                                    for="optPValDate_Current"
	                                                                                    class="radio"
		                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                                            />
    																					Current
                                                            	    		        </label>
																				</TD>
																				<TD width=20>&nbsp;&nbsp;</TD>
																				<TD width=5>
																					<input id=optPValDate_MonthEnd name=optPValDate type=radio 
																					    onclick="pVal_changeDateOption(3)"
		                                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                        onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                        onblur="try{radio_onBlur(this);}catch(e){}"/>
																				</TD>
																				<TD width=5>&nbsp;</TD>
																				<TD nowrap>
                                                                                    <label 
                                                                                        tabindex=-1
	                                                                                    for="optPValDate_MonthEnd"
	                                                                                    class="radio"
		                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                                            />
    																					Month End
                                                            	    		        </label>
																				</TD>
																				<TD width=20>&nbsp;&nbsp;</TD>
																				<TD width=5>
																					<input id=optPValDate_YearEnd name=optPValDate type=radio 
																					    onclick="pVal_changeDateOption(5)"
		                                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                        onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                        onblur="try{radio_onBlur(this);}catch(e){}"/>
																				</TD>
																				<TD width=5>&nbsp;</TD>
																				<TD nowrap>
                                                                                    <label 
                                                                                        tabindex=-1
	                                                                                    for="optPValDate_YearEnd"
	                                                                                    class="radio"
		                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                                            />
    																					Year End
                                                            	    		        </label>
																				</TD>
																				<TD width=100%>&nbsp;</TD>
																			</TR>
																		</TABLE>
																	</TD>
																	<TD width=20>&nbsp;&nbsp;</TD>
																</TR>

																<TR height=10 id=trPValDateOptions2>
																	<TD colspan=8></TD>
																</TR>

																<TR height=10 id=trPValTextDefault>
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=100% colspan=6>
																		<INPUT id=txtPValDefault name=txtPValDefault class="text" style="WIDTH: 100%">	
																	</TD>
																	<TD width=20>&nbsp;&nbsp;</TD>
																</TR>

																<TR height=10 id=trPValComboDefault style="visibility:hidden;display:none">
																	<TD width=20>&nbsp;&nbsp;</TD>
																	<TD width=100% colspan=6>
																		<select id=cboPValDefault name=cboPValDefault style="WIDTH: 100%"> 
																		</select>
																	</TD>
																	<TD width=20>&nbsp;&nbsp;</TD>
																</TR>

																<TR height=10>
																	<TD colspan=11></TD>
																</TR>
															</TABLE>
														</TD>
													</TR>
												</TABLE>

												</TD>
												<TD width=20>&nbsp;&nbsp;</TD>
											</TR>

											<TR height=100%>
												<TD colspan=6>&nbsp;</TD>
											</TR>
										</TABLE>
									</DIV>
								</TD>
							</TR>
						</TABLE>
					</TD>
					<td width=10>&nbsp;&nbsp;</td>
				</tr>
								
				<TR>
					<TD height=10 colspan=5></td>
				</tr>
				
				<tr height=10>
					<td width=10></td>
					<td colspan=3>
						<table WIDTH=100% class="invisible" CELLSPACING="0" CELLPADDING="0">
							<TR>
								<TD colspan=4>
								</TD>
							</TR>
							<tr>	
								<td>
								</td>
								<td width=10>
									<input id=cmdOK name=cmdOK type="button" class="btn" value="OK" style="WIDTH: 75px" width="75" 
									    onclick="component_OKClick()"
		                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                onfocus="try{button_onFocus(this);}catch(e){}"
		                                onblur="try{button_onBlur(this);}catch(e){}" />
								</td>
								<td width=40>
								</td>
								<td width=10>
									<input id="cmdCancel" name="cmdCancel" type="button" class="btn" value="Cancel" style="WIDTH: 75px" width="75" 
									    onclick="component_CancelClick()"
		                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                onfocus="try{button_onFocus(this);}catch(e){}"
		                                onblur="try{button_onBlur(this);}catch(e){}" />
								</td>
							</tr>			
						</table>
					</td>
					<td width=10></td>
				</tr>
				<TR>
					<TD height=10 colspan=7></td>
				</tr>
			</TABLE>
		</td>
	</tr>
</TABLE>
</FORM>

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

<FORM id=util_def_exprcomponent_frmOriginalDefinition name=util_def_exprcomponent_frmOriginalDefinition>
<%
    Dim sDefnString As String
    Dim sFieldTableID As String
    Dim sFieldColumnID As String
    Dim sLookupTableID As String
    Dim sLookupColumnID As String
    
	on error resume next
	sErrMsg = ""

	if session("optionAction") = "EDITEXPRCOMPONENT"	then
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
</FORM>

<FORM id=frmTables name=frmTables>
<%
    Dim cmdTables
    Dim prmTableID
    Dim rstTables
    Dim iCount As Integer
    
	if len(sErrMsg) = 0 then
        cmdTables = CreateObject("ADODB.Command")
		cmdTables.CommandText = "sp_ASRIntGetExprTables"
		cmdTables.CommandType = 4 ' Stored Procedure
        cmdTables.ActiveConnection = Session("databaseConnection")

        prmTableID = cmdTables.CreateParameter("tableID", 3, 1) ' 3=integer, 1=input
        cmdTables.Parameters.Append(prmTableID)
		prmTableID.value = cleanNumeric(session("optionTableID"))

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
                    Response.Write("<INPUT type='hidden' id=txtTable_" & iCount & " name=txtTable_" & iCount & " value=""" & rstTables.fields("definitionString").value & """>" & vbCrLf)
                    rstTables.MoveNext()
                Loop

                ' Release the ADO recordset object.
                rstTables.close()
            End If
            rstTables = Nothing
        End If

		' Release the ADO command object.
        cmdTables = Nothing
	end if
%>
</FORM>

<FORM id=frmFunctions name=frmFunctions>
<%
    Dim cmdFunctions
    Dim rstFunctions
    
	if len(sErrMsg) = 0 then
        cmdFunctions = CreateObject("ADODB.Command")
        cmdFunctions.CommandText = "sp_ASRIntGetExprFunctions"
		cmdFunctions.CommandType = 4 ' Stored Procedure
        cmdFunctions.ActiveConnection = Session("databaseConnection")

        prmTableID = cmdFunctions.CreateParameter("tableID", 3, 1) ' 3=integer, 1=input
        cmdFunctions.Parameters.Append(prmTableID)
		prmTableID.value = cleanNumeric(session("optionTableID"))

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
                    Response.Write("<INPUT type='hidden' id=txtFunction_" & iCount & " name=txtFunction_" & iCount & " value=""" & rstFunctions.fields("definitionString").value & """>" & vbCrLf)
                    rstFunctions.MoveNext()
                Loop

                ' Release the ADO recordset object.
                rstFunctions.close()
            End If
            rstFunctions = Nothing
        End If

		' Release the ADO command object.
        cmdFunctions = Nothing
	end if
%>
</FORM>

<FORM id=frmFunctionParameters name=frmFunctionParameters>
<%
    Dim cmdFunctionParameters
    Dim rstFunctionParameters
    
	if len(sErrMsg) = 0 then
        cmdFunctionParameters = CreateObject("ADODB.Command")
		cmdFunctionParameters.CommandText = "sp_ASRIntGetExprFunctionParameters"
		cmdFunctionParameters.CommandType = 4 ' Stored Procedure
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
                    Response.Write("<INPUT type='hidden' id=txtFunctionParameters_" & rstFunctionParameters.fields("functionID").value & "_" & iCount & " name=txtFunctionParameters_" & rstFunctionParameters.fields("functionID").value & "_" & iCount & " value=""" & rstFunctionParameters.fields("parameterName").value & """>" & vbCrLf)
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
	end if
%>
</FORM>

<FORM id=frmOperators name=frmOperators>
<%
    Dim cmdOperators
    Dim rstOperators
    
	if len(sErrMsg) = 0 then
        cmdOperators = CreateObject("ADODB.Command")
		cmdOperators.CommandText = "sp_ASRIntGetExprOperators"
		cmdOperators.CommandType = 4 ' Stored Procedure
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
                    Response.Write("<INPUT type='hidden' id=txtOperator_" & iCount & " name=txtOperator_" & iCount & " value=""" & rstOperators.fields("definitionString").value & """>" & vbCrLf)
                    rstOperators.MoveNext()
                Loop

                ' Release the ADO recordset object.
                rstOperators.close()
            End If
            rstOperators = Nothing
        End If

		' Release the ADO command object.
        cmdOperators = Nothing
	end if
%>
</FORM>

<FORM id=frmCalcs name=frmCalcs>
<%
    Dim cmdCalcs
    Dim rstCalcs
    Dim prmExprID
    Dim prmBaseTableID
    
    
    
	if len(sErrMsg) = 0 then
        cmdCalcs = CreateObject("ADODB.Command")
		cmdCalcs.CommandText = "sp_ASRIntGetExprCalcs"
		cmdCalcs.CommandType = 4 ' Stored Procedure
        cmdCalcs.ActiveConnection = Session("databaseConnection")

        prmExprID = cmdCalcs.CreateParameter("exprID", 3, 1) ' 3=integer, 1=input
        cmdCalcs.Parameters.Append(prmExprID)
		prmExprID.value = cleanNumeric(clng(session("optionExprID")))

        prmBaseTableID = cmdCalcs.CreateParameter("baseTableID", 3, 1) ' 3=integer, 1=input
        cmdCalcs.Parameters.Append(prmBaseTableID)
		prmBaseTableID.value = cleanNumeric(clng(session("optionTableID")))

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
                    Response.Write("<INPUT type='hidden' id=txtCalc_" & iCount & " name=txtCalc_" & iCount & " value=""" & Replace(rstCalcs.fields("definitionString").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtCalcDesc_" & iCount & " name=txtCalcDesc_" & iCount & " value=""" & Replace(rstCalcs.fields("description").value, """", "&quot;") & """>" & vbCrLf)
                    rstCalcs.MoveNext()
                Loop

                ' Release the ADO recordset object.
                rstCalcs.close()
            End If
            rstCalcs = Nothing
        End If

		' Release the ADO command object.
        cmdCalcs = Nothing
	end if
%>
</FORM>
	
<FORM id=frmFilters name=frmFilters>
<%
    Dim cmdFilters
    Dim rstFilters
    
	if len(sErrMsg) = 0 then
        cmdFilters = CreateObject("ADODB.Command")
		cmdFilters.CommandText = "sp_ASRIntGetExprFilters"
		cmdFilters.CommandType = 4 ' Stored Procedure
        cmdFilters.ActiveConnection = Session("databaseConnection")

        prmExprID = cmdFilters.CreateParameter("exprID", 3, 1) ' 3=integer, 1=input
        cmdFilters.Parameters.Append(prmExprID)
		prmExprID.value = cleanNumeric(clng(session("optionExprID")))

        prmBaseTableID = cmdFilters.CreateParameter("baseTableID", 3, 1) ' 3=integer, 1=input
        cmdFilters.Parameters.Append(prmBaseTableID)
		prmBaseTableID.value = cleanNumeric(clng(session("optionTableID")))

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
                    Response.Write("<INPUT type='hidden' id=txtFilter_" & iCount & " name=txtFilter_" & iCount & " value=""" & Replace(rstFilters.fields("definitionString").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtFilterDesc_" & iCount & " name=txtFilterDesc_" & iCount & " value=""" & Replace(rstFilters.fields("description").value, """", "&quot;") & """>" & vbCrLf)
                    rstFilters.MoveNext()
                Loop

                ' Release the ADO recordset object.
                rstFilters.close()
            End If
            rstFilters = Nothing
        End If

		' Release the ADO command object.
        cmdFilters = Nothing
	end if
%>
</FORM>
	
<form id="frmFieldRec" name="frmFieldRec" target="fieldRec" action="fieldRec" method="post" style="visibility: hidden; display: none">
    <input type="hidden" id="selectionType" name="selectionType">
    <input type="hidden" id="Hidden1" name="txtTableID">
    <input type="hidden" id="selectedID" name="selectedID">
</form>

<INPUT type='hidden' id=txtTicker name=txtTicker value=0>
<INPUT type='hidden' id=txtLastKeyFind name=txtLastKeyFind value="">

<%
    Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrMsg & """>" & vbCrLf)
%>


<script runat="server" language="vb">

function componentParameter(psDefnString, psParameter)
	dim iCharIndex
	dim sDefn
	
	sDefn = psDefnString
	
	iCharIndex = instr(sDefn, "	")
	if iCharIndex >= 0 then
		if psParameter = "COMPONENTID" then
			componentParameter = left(sDefn, iCharIndex - 1)
			exit function
		end if
		
		sDefn = mid(sDefn, iCharIndex + 1)
		iCharIndex = instr(sDefn, "	")
		if iCharIndex >= 0 then
			if psParameter = "EXPRID" then
				componentParameter = left(sDefn, iCharIndex - 1)
				exit function
			end if
			
			sDefn = mid(sDefn, iCharIndex + 1)
			iCharIndex = instr(sDefn, "	")
			if iCharIndex >= 0 then
				if psParameter = "TYPE" then
					componentParameter = left(sDefn, iCharIndex - 1)
					exit function
				end if
				
				sDefn = mid(sDefn, iCharIndex + 1)
				iCharIndex = instr(sDefn, "	")
				if iCharIndex >= 0 then
					if psParameter = "FIELDCOLUMNID" then
						componentParameter = left(sDefn, iCharIndex - 1)
						exit function
					end if
					
					sDefn = mid(sDefn, iCharIndex + 1)
					iCharIndex = instr(sDefn, "	")
					if iCharIndex >= 0 then
						if psParameter = "FIELDPASSBY" then
							componentParameter = left(sDefn, iCharIndex - 1)
							exit function
						end if
						
						sDefn = mid(sDefn, iCharIndex + 1)
						iCharIndex = instr(sDefn, "	")
						if iCharIndex >= 0 then
							if psParameter = "FIELDSELECTIONTABLEID" then
								componentParameter = left(sDefn, iCharIndex - 1)
								exit function
							end if
							
							sDefn = mid(sDefn, iCharIndex + 1)
							iCharIndex = instr(sDefn, "	")
							if iCharIndex >= 0 then
								if psParameter = "FIELDSELECTIONRECORD" then
									componentParameter = left(sDefn, iCharIndex - 1)
									exit function
								end if
								
								sDefn = mid(sDefn, iCharIndex + 1)
								iCharIndex = instr(sDefn, "	")
								if iCharIndex >= 0 then
									if psParameter = "FIELDSELECTIONLINE" then
										componentParameter = left(sDefn, iCharIndex - 1)
										exit function
									end if
									
									sDefn = mid(sDefn, iCharIndex + 1)
									iCharIndex = instr(sDefn, "	")
									if iCharIndex >= 0 then
										if psParameter = "FIELDSELECTIONORDERID" then
											componentParameter = left(sDefn, iCharIndex - 1)
											exit function
										end if
										
										sDefn = mid(sDefn, iCharIndex + 1)
										iCharIndex = instr(sDefn, "	")
										if iCharIndex >= 0 then
											if psParameter = "FIELDSELECTIONFILTER" then
												componentParameter = left(sDefn, iCharIndex - 1)
												exit function
											end if
											
											sDefn = mid(sDefn, iCharIndex + 1)
											iCharIndex = instr(sDefn, "	")
											if iCharIndex >= 0 then
												if psParameter = "FUNCTIONID" then
													componentParameter = left(sDefn, iCharIndex - 1)
													exit function
												end if
												
												sDefn = mid(sDefn, iCharIndex + 1)
												iCharIndex = instr(sDefn, "	")
												if iCharIndex >= 0 then
													if psParameter = "CALCULATIONID" then
														componentParameter = left(sDefn, iCharIndex - 1)
														exit function
													end if
													
													sDefn = mid(sDefn, iCharIndex + 1)
													iCharIndex = instr(sDefn, "	")
													if iCharIndex >= 0 then
														if psParameter = "OPERATORID" then
															componentParameter = left(sDefn, iCharIndex - 1)
															exit function
														end if
														
														sDefn = mid(sDefn, iCharIndex + 1)
														iCharIndex = instr(sDefn, "	")
														if iCharIndex >= 0 then
															if psParameter = "VALUETYPE" then
																componentParameter = left(sDefn, iCharIndex - 1)
																exit function
															end if
															
															sDefn = mid(sDefn, iCharIndex + 1)
															iCharIndex = instr(sDefn, "	")
															if iCharIndex >= 0 then
																if psParameter = "VALUECHARACTER" then
																	componentParameter = left(sDefn, iCharIndex - 1)
																	exit function
																end if
																
																sDefn = mid(sDefn, iCharIndex + 1)
																iCharIndex = instr(sDefn, "	")
																if iCharIndex >= 0 then
																	if psParameter = "VALUENUMERIC" then
																		componentParameter = left(sDefn, iCharIndex - 1)
																		exit function
																	end if
																	
																	sDefn = mid(sDefn, iCharIndex + 1)
																	iCharIndex = instr(sDefn, "	")
																	if iCharIndex >= 0 then
																		if psParameter = "VALUELOGIC" then
																			componentParameter = left(sDefn, iCharIndex - 1)
																			exit function
																		end if
																		
																		sDefn = mid(sDefn, iCharIndex + 1)
																		iCharIndex = instr(sDefn, "	")
																		if iCharIndex >= 0 then
																			if psParameter = "VALUEDATE" then
																				componentParameter = left(sDefn, iCharIndex - 1)
																				exit function
																			end if
																			
																			sDefn = mid(sDefn, iCharIndex + 1)
																			iCharIndex = instr(sDefn, "	")
																			if iCharIndex >= 0 then
																				if psParameter = "PROMPTDESCRIPTION" then
																					componentParameter = left(sDefn, iCharIndex - 1)
																					exit function
																				end if
																				
																				sDefn = mid(sDefn, iCharIndex + 1)
																				iCharIndex = instr(sDefn, "	")
																				if iCharIndex >= 0 then
																					if psParameter = "PROMPTMASK" then
																						componentParameter = left(sDefn, iCharIndex - 1)
																						exit function
																					end if
																					
																					sDefn = mid(sDefn, iCharIndex + 1)
																					iCharIndex = instr(sDefn, "	")
																					if iCharIndex >= 0 then
																						if psParameter = "PROMPTSIZE" then
																							componentParameter = left(sDefn, iCharIndex - 1)
																							exit function
																						end if
																						
																						sDefn = mid(sDefn, iCharIndex + 1)
																						iCharIndex = instr(sDefn, "	")
																						if iCharIndex >= 0 then
																							if psParameter = "PROMPTDECIMALS" then
																								componentParameter = left(sDefn, iCharIndex - 1)
																								exit function
																							end if
																							
																							sDefn = mid(sDefn, iCharIndex + 1)
																							iCharIndex = instr(sDefn, "	")
																							if iCharIndex >= 0 then
																								if psParameter = "FUNCTIONRETURNTYPE" then
																									componentParameter = left(sDefn, iCharIndex - 1)
																									exit function
																								end if
																								
																								sDefn = mid(sDefn, iCharIndex + 1)
																								iCharIndex = instr(sDefn, "	")
																								if iCharIndex >= 0 then
																									if psParameter = "LOOKUPTABLEID" then
																										componentParameter = left(sDefn, iCharIndex - 1)
																										exit function
																									end if
																									
																									sDefn = mid(sDefn, iCharIndex + 1)
																									iCharIndex = instr(sDefn, "	")
																									if iCharIndex >= 0 then
																										if psParameter = "LOOKUPCOLUMNID" then
																											componentParameter = left(sDefn, iCharIndex - 1)
																											exit function
																										end if
																										
																										sDefn = mid(sDefn, iCharIndex + 1)
																										iCharIndex = instr(sDefn, "	")
																										if iCharIndex >= 0 then
																											if psParameter = "FILTERID" then
																												componentParameter = left(sDefn, iCharIndex - 1)
																												exit function
																											end if
																											
																											sDefn = mid(sDefn, iCharIndex + 1)
																											iCharIndex = instr(sDefn, "	")
																											if iCharIndex >= 0 then
																												if psParameter = "EXPANDEDNODE" then
																													componentParameter = left(sDefn, iCharIndex - 1)
																													exit function
																												end if
																												
																												sDefn = mid(sDefn, iCharIndex + 1)
																												iCharIndex = instr(sDefn, "	")
																												if iCharIndex >= 0 then
																													if psParameter = "PROMPTDATETYPE" then
																														componentParameter = left(sDefn, iCharIndex - 1)
																														exit function
																													end if
																													
																													sDefn = mid(sDefn, iCharIndex + 1)
																													iCharIndex = instr(sDefn, "	")
																													if iCharIndex >= 0 then
																														if psParameter = "DESCRIPTION" then
																															componentParameter = left(sDefn, iCharIndex - 1)
																															exit function
																														end if
																														
																														sDefn = mid(sDefn, iCharIndex + 1)
																														iCharIndex = instr(sDefn, "	")
																														if iCharIndex >= 0 then
																															if psParameter = "FIELDTABLEID" then
																																componentParameter = left(sDefn, iCharIndex - 1)
																																exit function
																															end if
																															
																															sDefn = mid(sDefn, iCharIndex + 1)
																															iCharIndex = instr(sDefn, "	")
																															if iCharIndex >= 0 then
																																if psParameter = "FIELDSELECTIONORDERNAME" then
																																	componentParameter = left(sDefn, iCharIndex - 1)
																																	exit function
																																end if
																																
																																sDefn = mid(sDefn, iCharIndex + 1)
																																if psParameter = "FIELDSELECTIONFILTERNAME" then
																																	componentParameter = sDefn
																																	exit function
																																end if
																															end if
																														end if	
																													end if	
																												end if	
																											end if	
																										end if	
																									end if	
																								end if	
																							end if	
																						end if	
																					end if	
																				end if	
																			end if	
																		end if	
																	end if	
																end if	
															end if	
														end if	
													end if	
												end if	
											end if	
										end if	
									end if	
								end if	
							end if	
						end if	
					end if	
				end if	
			end if	
		end if	
	end if
	
	componentParameter = ""
end function

    </script>


<script type="text/javascript">
    util_def_exprcomponent_addhandlers();
    util_def_exprcomponent_onload();
</script>


