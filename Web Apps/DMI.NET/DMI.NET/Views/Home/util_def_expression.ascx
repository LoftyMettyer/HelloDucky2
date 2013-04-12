<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script src="<%: Url.Content("~/Scripts/FormScripts/util_def_expression.js")%>" type="text/javascript"></script>

<object
    classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
    id="Microsoft_Licensed_Class_Manager_1_0"
    viewastext>
    <param name="LPKPath" value="lpks/main.lpk">
</object>


<OBJECT classid="clsid:6976CB54-C39B-4181-B1DC-1A829068E2E7" codebase="cabs/COAInt_Client.cab#Version=1,0,0,5" 
	id=abExprMenu name=abExprMenu style="left:0px;top:0px;position:absolute; height: 10px;" VIEWASTEXT>
	<PARAM NAME="_ExtentX" VALUE="0">
	<PARAM NAME="_ExtentY" VALUE="0">
</OBJECT>

<OBJECT classid="clsid:1C203F13-95AD-11D0-A84B-00A0247B735B" id=SSTreeClipboard   codebase="cabs/SStree.cab#version=1,0,2,24" style="LEFT: 0px; TOP: 0px; WIDTH:0px; HEIGHT:0px" VIEWASTEXT>
	<PARAM NAME="_ExtentX" VALUE="370">
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

<OBJECT classid="clsid:1C203F13-95AD-11D0-A84B-00A0247B735B" id=SSTreeUndo codebase="cabs/SStree.cab#version=1,0,2,24" style="LEFT: 0px; TOP: 0px; WIDTH:0px; HEIGHT:0px" VIEWASTEXT>
	<PARAM NAME="_ExtentX" VALUE="370">
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

<form id=frmDefinition>
<table align=center class="outline" cellPadding=5 cellSpacing=0 width=100% height=100%>
	<TR>
		<TD>
			<TABLE WIDTH="100%" height="100%" class="invisible" cellspacing=0 cellpadding=0>
				<tr> 
					<TD width=10></td>
					<td>
						<TABLE WIDTH="100%" height="100%" class="invisible" cellspacing=0 cellpadding=5>
							<tr valign=top> 
								<td>
									<TABLE WIDTH="100%" height="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD colspan=9 height=5></TD>
										</TR>

										<TR height=10>
											<TD width=5>&nbsp;</TD>
											<TD width=10>Name :</TD>
											<TD width=5>&nbsp;</TD>
											<TD>
												<INPUT id=txtName name=txtName class="text" maxlength="50" style="WIDTH: 100%" onkeyup="changeName()">
											</TD>
											<TD width=20>&nbsp;</TD>
											<TD width=10>Owner :</TD>
											<TD width=5>&nbsp;</TD>
											<TD width="40%">
												<INPUT id=txtOwner name=txtOwner class="text textdisabled" style="WIDTH: 100%" disabled="disabled" tabindex="-1">
											</TD>
											<TD width=5>&nbsp;</TD>
										</TR>
											
										<TR>
											<TD colspan=9 height=5></TD>
										</TR>
											
										<TR height=10>
											<TD width=5>&nbsp;</TD>
											<TD width=10 nowrap>Description :</TD>
											<TD width=5>&nbsp;</TD>
											<TD width="40%" rowspan="5">
												<TEXTAREA id=txtDescription name=txtDescription class="textarea" style="HEIGHT: 99%; WIDTH: 100%" wrap=VIRTUAL height="0" maxlength="255" onkeyup="changeDescription()" 
												    onpaste="var selectedLength = document.selection.createRange().text.length;var pasteData = window.clipboardData.getData('Text');if ((this.value.length + pasteData.length - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}" 
												    onkeypress="var selectedLength = document.selection.createRange().text.length;if ((this.value.length + 1 - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}">
												</TEXTAREA>
											</TD>
											<TD width=20 nowrap>&nbsp;</TD>
											<TD width=10>Access :</TD>
											<TD width=5>&nbsp;</TD>
											<TD width="40%">
												<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 width="100%">
													<TR>
														<TD width=5>
															<INPUT CHECKED id=optAccessRW name=optAccess type=radio 
															    onclick="changeAccess()"
		                                                        onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
														</TD>
														<TD width=5>&nbsp;</TD>
														<TD width=30>
                                                            <label 
                                                                tabindex="-1"
	                                                            for="optAccessRW"
	                                                            class="radio"
		                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                    />
															    Read/Write
                                    	    		        </label>
														</TD>
														<TD>&nbsp;</TD>
													</TR>
												</TABLE>
											</TD>
											<TD width=5>&nbsp;</TD>
										</TR>
											
										<TR>
											<TD colspan=8 height=5></TD>
										</TR>					

										<TR height=10>
											<TD width=5>&nbsp;</TD>

											<TD width=10>&nbsp;</TD>
											<TD width=5>&nbsp;</TD>

											<TD width=20 nowrap>&nbsp;</TD>

											<TD width=10>&nbsp;</TD>
											<TD width=5>&nbsp;</TD>
											<TD width="40%">
												<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 width="100%">
													<TR>
														<TD width=5>
															<input id=optAccessRO name=optAccess type=radio 
															    onclick="changeAccess()"
		                                                        onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
														</TD>
														<TD width=5>&nbsp;</TD>
														<TD width=80 nowrap>
                                                            <label 
                                                                tabindex="-1"
	                                                            for="optAccessRO"
	                                                            class="radio"
		                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                    />
    															Read Only
                                    	    		        </label>
														</TD>
														<TD>&nbsp;</TD>
													</TR>
												</TABLE>
											</TD>
											<TD width=5>&nbsp;</TD>
										</TR>
											
										<TR>
											<TD colspan=8 height=5></TD>
										</TR>					

										<TR height=10>
											<TD width=5>&nbsp;</TD>
											<TD width=10>&nbsp;</TD>
											<TD width=5>&nbsp;</TD>
											<TD width=20 nowrap>&nbsp;</TD>
											<TD width=10>&nbsp;</TD>
											<TD width=5>&nbsp;</TD>
											<TD width="40%">
												<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 width="100%">
													<TR>
														<TD width=5>
															<input id=optAccessHD name=optAccess type=radio 
															    onclick="changeAccess()"
		                                                        onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
														</TD>
														<TD width=5>&nbsp;</TD>
														<TD width=60 nowrap>
                                                            <label 
                                                                tabindex="-1"
	                                                            for="optAccessHD"
	                                                            class="radio"
		                                                        onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                    />
    															Hidden
                                    	    		        </label>
														</TD>
														<TD>&nbsp;</TD>
													</TR>
												</TABLE>
											</TD>
											<TD width=5>&nbsp;</TD>
										</TR>
											
										<TR>
											<TD colspan=9>
												<TABLE WIDTH=100% HEIGHT=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD colspan=3 height=30><hr></TD>
													</TR>
													<TR height=10>
														<TD rowspan=16>
															<OBJECT classid="clsid:1C203F13-95AD-11D0-A84B-00A0247B735B" id=SSTree1 
                                                                codebase="cabs/SStree.cab#version=1,0,2,24" style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:400px; VISIBILITY: visible;" VIEWASTEXT>
																<PARAM NAME="_ExtentX" VALUE="31882">
																<PARAM NAME="_ExtentY" VALUE="16404">
																<PARAM NAME="_Version" VALUE="65538">
																<PARAM NAME="BackColor" VALUE="-2147483643">
																<PARAM NAME="ForeColor" VALUE="-2147483640">
																<PARAM NAME="ImagesMaskColor" VALUE="12632256">
																<PARAM NAME="PictureBackgroundMaskColor" VALUE="12632256">
																<PARAM NAME="Appearance" VALUE="0">
																<PARAM NAME="BorderStyle" VALUE="1">
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
														<TD rowspan=16 width=10>&nbsp;</TD>
														<TD width=80>
															<input type=button id=cmdAdd name=cmdAdd class="btn" value=Add style="WIDTH: 100%"  
															    onclick="addClick()"
		                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{button_onFocus(this);}catch(e){}"
		                                                        onblur="try{button_onBlur(this);}catch(e){}" />
														</TD>
													</TR>
													<TR height=10>
														<TD>&nbsp;</TD>
													</TR>
													<TR height=10>
														<TD width=80>
															<input type=button id=cmdInsert name=cmdInsert class="btn" value="Insert" style="WIDTH: 100%"  
															    onclick="insertClick()"
		                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{button_onFocus(this);}catch(e){}"
		                                                        onblur="try{button_onBlur(this);}catch(e){}" />
														</TD>
													</TR>
													<TR height=10>
														<TD>&nbsp;</TD>
													</TR>
													<TR height=10>
														<TD width=80>
															<input type=button id=cmdEdit name=cmdEdit class="btn" value="Edit" 

style="WIDTH: 100%"  
															    onclick="editClick()"
		                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{button_onFocus(this);}catch(e){}"
		                                                        onblur="try{button_onBlur(this);}catch(e){}" />
														</TD>
													</TR>
													<TR height=10>
														<TD>&nbsp;</TD>
													</TR>
													<TR height=10>
														<TD width=80>
															<input type=button id=cmdDelete name=cmdDelete class="btn" value="Delete" 

style="WIDTH: 100%"  
															    onclick="deleteClick()"
		                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{button_onFocus(this);}catch(e){}"
		                                                        onblur="try{button_onBlur(this);}catch(e){}" />
														</TD>
													</TR>
													<TR height=10>
														<TD>&nbsp;</TD>
													</TR>
													<TR height=10>
														<TD width=80>
															<input type=button id=cmdPrint name=cmdPrint class="btn" value="Print" 

style="WIDTH: 100%"  
															    onclick="printClick(true)"
		                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{button_onFocus(this);}catch(e){}"
		                                                        onblur="try{button_onBlur(this);}catch(e){}" />
														</TD>
													</TR>

<%
	if session("utiltype") = 11 then
%>
                                                    <TR height=10>
                                                        <TD>&nbsp;</TD>
                                                    </TR>
                                                    <TR height=10>
                                                        <TD width=80>
                                                            <input type=button id=cmdTest name=cmdTest class="btn" value="Test" style="WIDTH: 100%" 
                                                                onclick="testClick()"
		                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{button_onFocus(this);}catch(e){}"
		                                                        onblur="try{button_onBlur(this);}catch(e){}" />
                                                        </TD>
                                                    </TR>
<%	
    end if
%>													
													<TR>
														<TD></TD>
													</TR>
													<TR height=10>
														<TD>&nbsp;</TD>
													</TR>
													<TR height=10>
														<TD width=80>
															<input type=button id=cmdOK name=cmdOK class="btn" value=OK style="WIDTH: 100%"
															    onclick="okClick()"
		                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{button_onFocus(this);}catch(e){}"
		                                                        onblur="try{button_onBlur(this);}catch(e){}" />
														</TD>
													</TR>
													<TR height=10>
														<TD>&nbsp;</TD>
													</TR>
													<TR height=10>
														<TD width=80>
															<input type=button id=cmdCancel name=cmdCancel class="btn" value=Cancel style="WIDTH: 100%"  
															    onclick="cancelClick()"
		                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{button_onFocus(this);}catch(e){}"
		                                                        onblur="try{button_onBlur(this);}catch(e){}" />
														</TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
											
										<TR height=5>
											<TD colspan=9 height=5></TD>
										</TR>
									</TABLE>
								</td>
							</tr>
						</TABLE>
					</td>
					<TD width=10></td>
				</tr> 

				<tr height=5> 
					<td colspan=3></td>
				</tr> 
			</TABLE>
		</td>
	</tr> 
</TABLE>

</form>

<FORM action="default_Submit" method=post id=frmGoto name=frmGoto style="visibility:hidden;display:none">
	<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</FORM>
 
<form id=frmOriginalDefinition style="visibility:hidden;display:none">
<%
    Dim sReaction As String
    Dim sUtilTypeName As String
    Dim sErrMsg As String
    Dim iCount As Integer
    
    sUtilTypeName = "expression"
	if session("utiltype") = 11 then
		sUtilTypeName = "filter"
		sReaction = "FILTERS"
	else
		if session("utiltype") = 12 then
			sUtilTypeName = "calculation"
			sReaction = "CALCULATIONS"
		end if
	end if

	if session("action") <> "new"	then
        Dim cmdDefn = CreateObject("ADODB.Command")
		cmdDefn.CommandText = "sp_ASRIntGetExpressionDefinition"
		cmdDefn.CommandType = 4 ' Stored Procedure
        cmdDefn.ActiveConnection = Session("databaseConnection")

        Dim prmUtilID = cmdDefn.CreateParameter("utilID", 3, 1) ' 3=integer, 1=input
        cmdDefn.Parameters.Append(prmUtilID)
		prmUtilID.value = cleanNumeric(session("utilid"))

        Dim prmAction = cmdDefn.CreateParameter("action", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
        cmdDefn.Parameters.Append(prmAction)
		prmAction.value = session("action")

        Dim prmErrMsg = cmdDefn.CreateParameter("errMsg", 200, 2, 8000) '200=varchar, 2=output, 8000=size
        cmdDefn.Parameters.Append(prmErrMsg)

        Dim prmTimestamp = cmdDefn.CreateParameter("timestamp", 3, 2) '3=integer, 2=output
        cmdDefn.Parameters.Append(prmTimestamp)

        Err.Clear()
        Dim rstDefinition = cmdDefn.Execute
        If (Err.Number <> 0) Then
            sErrMsg = "'" & Session("utilname") & "' " & sUtilTypeName & " definition could not be read." & vbCrLf & FormatError(Err.Description)
        Else
            If rstDefinition.state <> 0 Then
                ' Read recordset values.
                iCount = 0
                Do While Not rstDefinition.EOF
                    Response.Write("<INPUT type='hidden' id=txtDefn_" & rstDefinition.fields("type").value & "_" & iCount & " name=txtDefn_" & rstDefinition.fields("type").value & "_" & iCount & " value=""" & Replace(rstDefinition.fields("definition").value, """", "&quot;") & """>" & vbCrLf)

                    iCount = iCount + 1
                    rstDefinition.MoveNext()
                Loop
	
                ' Release the ADO recordset object.
                rstDefinition.close()
            End If
            rstDefinition = Nothing
			
            ' NB. IMPORTANT ADO NOTE.
            ' When calling a stored procedure which returns a recordset AND has output parameters
            ' you need to close the recordset and set it to nothing before using the output parameters. 
            If Len(cmdDefn.Parameters("errMsg").Value) > 0 Then
                sErrMsg = "'" & Session("utilname") & "' " & cmdDefn.Parameters("errMsg").Value
            End If

            Response.Write("<INPUT type='hidden' id=txtDefn_Timestamp name=txtDefn_Timestamp value=" & cmdDefn.Parameters("timestamp").Value & ">" & vbCrLf)
        End If

		' Release the ADO command object.
        cmdDefn = Nothing

		if len(sErrMsg) > 0 then
			session("confirmtext") = sErrMsg
			session("confirmtitle") = "OpenHR Intranet"
            Session("followpage") = "defsel"
			Session("reaction") = sReaction
			Response.Clear
            Response.Redirect("confirmok")
		end if
	end if
%>
	<INPUT type="hidden" id=txtOriginalAccess name=txtOriginalAccess value="RW">
</form>

<FORM id=frmUseful name=frmUseful style="visibility:hidden;display:none">
	<INPUT type="hidden" id=txtUserName name=txtUserName value="<%=session("username")%>">
	<INPUT type="hidden" id=txtLoading name=txtLoading value="Y">
	<INPUT type="hidden" id=txtChanged name=txtChanged value=0>
	<INPUT type="hidden" id=txtUtilID name=txtUtilID value=<% =session("utilid")%>>
	<INPUT type="hidden" id=txtTableID name=txtTableID value=<% =session("utiltableid")%>>
	<INPUT type="hidden" id=txtAction name=txtAction value=<% =session("action")%>>
	<INPUT type="hidden" id=txtUtilType name=txtUtilType value=<% =session("utiltype")%>>
	<INPUT type="hidden" id=txtLocaleDecimal name=txtLocaleDecimal value=<% =session("LocaleDecimalSeparator")%>>
	<INPUT type="hidden" id=txtExprColourMode name=txtExprColourMode value=<% =session("ExprColourMode")%>>
	<INPUT type="hidden" id=txtExprNodeMode name=txtExprNodeMode value=<% =session("ExprNodeMode")%>>
	<INPUT type="hidden" id=txtLastNode name=txtLastNode>
	<INPUT type="hidden" id=txtMenuSaved name=txtMenuSaved value=0>

    <%
        Dim sErrorDescription As String
        
        Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
	
        Dim cmdBaseTable = CreateObject("ADODB.Command")
	cmdBaseTable.CommandText = "sp_ASRIntGetTableName"
	cmdBaseTable.CommandType = 4 ' Stored Procedure
        cmdBaseTable.ActiveConnection = Session("databaseConnection")

        Dim prmTableID = cmdBaseTable.CreateParameter("tableID", 3, 1) ' 3=integer, 1=input
        cmdBaseTable.Parameters.Append(prmTableID)
	prmTableID.value = cleanNumeric(session("utiltableid"))

        Dim prmTableName = cmdBaseTable.CreateParameter("tableName", 200, 2, 255)
        cmdBaseTable.Parameters.Append(prmTableName)

        Err.Clear()
	cmdBaseTable.Execute
			
        Response.Write("<INPUT type='hidden' id=txtTableName name=txtTableName value=""" & cmdBaseTable.Parameters("tableName").Value & """>" & vbCrLf)

	' Release the ADO command object.
        cmdBaseTable = Nothing
	
    %>
	<INPUT type="hidden" id=txtCanDelete name=txtCanDelete value=0>
	<INPUT type="hidden" id=txtCanInsert name=txtCanInsert value=0>
	<INPUT type="hidden" id=txtCanCut name=txtCanCut value=0>
	<INPUT type="hidden" id=txtCanCopy name=txtCanCopy value=0>
	<INPUT type="hidden" id=txtCanPaste name=txtCanPaste value=0>
	<INPUT type="hidden" id=txtCanMoveUp name=txtCanMoveUp value=0>
	<INPUT type="hidden" id=txtCanMoveDown name=txtCanMoveDown value=0>
	<INPUT type="hidden" id=txtUndoType name=txtUndoType value="">
	<INPUT type="hidden" id=txtOldText name=txtOldText value="">	
</FORM>

<FORM id=frmValidate name=frmValidate target=validate method=post action=util_validate_expression style="visibility:hidden;display:none">
	<INPUT type=hidden id=validatePass name=validatePass value=0>
	<INPUT type=hidden id=validateName name=validateName value=''>
	<INPUT type=hidden id=validateOwner name=validateOwner value=''>
	<INPUT type=hidden id=validateTimestamp name=validateTimestamp value=''>
	<INPUT type=hidden id=validateUtilID name=validateUtilID value=''>
	<INPUT type=hidden id=validateUtilType name=validateUtilType value=''>
	<INPUT type=hidden id=validateAccess name=validateAccess value=''>
	<INPUT type=hidden id=components1 name=components1 value="">
	<INPUT type=hidden id=validateBaseTableID name=validateBaseTableID value=<%=session("utiltableid")%>>
	<INPUT type=hidden id=validateOriginalAccess name=validateOriginalAccess value="RW">
</FORM>

    <form id="frmSend" name="frmSend" method="post" action="util_def_expression_Submit" style="visibility: hidden; display: none">
        <input type="hidden" id="txtSend_ID" name="txtSend_ID">
        <input type="hidden" id="txtSend_type" name="txtSend_type">
        <input type="hidden" id="txtSend_name" name="txtSend_name">
        <input type="hidden" id="txtSend_description" name="txtSend_description">
        <input type="hidden" id="txtSend_access" name="txtSend_access">
        <input type="hidden" id="txtSend_userName" name="txtSend_userName">
        <input type="hidden" id="txtSend_components1" name="txtSend_components1">
        <input type="hidden" id="txtSend_reaction" name="txtSend_reaction">
        <input type="hidden" id="txtSend_tableID" name="txtSend_tableID" value='<% =session("utiltableid")%>'>
        <input type="hidden" id="txtSend_names" name="txtSend_names" value="">
    </form>

<FORM id=frmTest name=frmTest target=test method=post action=util_test_expression_pval style="visibility:hidden;display:none">
	<INPUT type="hidden" id=type name=type>	
	<INPUT type="hidden" id=Hidden1 name=components1>
	<INPUT type="hidden" id=tableID name=tableID value=<% =session("utiltableid")%>>
	<INPUT type="hidden" id=prompts name=prompts>
	<INPUT type="hidden" id=filtersAndCalcs name=filtersAndCalcs>
</FORM>

<INPUT type='hidden' id=txtTicker name=txtTicker value=0>
<INPUT type='hidden' id=txtLastKeyFind name=txtLastKeyFind value="">

<form id="frmShortcutKeys" name="frmShortcutKeys" style="visibility: hidden; display: none">
    <%
        Dim sShortcutKeys As String
    
        sShortcutKeys = ""
	
        Dim cmdShortcutKeys = CreateObject("ADODB.Command")
        cmdShortcutKeys.CommandText = "spASRIntGetOpFuncShortcuts"
        cmdShortcutKeys.CommandType = 4 ' Stored Procedure
        cmdShortcutKeys.ActiveConnection = Session("databaseConnection")

        Err.Clear()
        Dim rstShortcutKeys = cmdShortcutKeys.Execute
        If (Err.Number <> 0) Then
            sErrMsg = "'" & Session("utilname") & "' " & sUtilTypeName & " definition could not be read." & vbCrLf & FormatError(Err.Description)
        Else
            If rstShortcutKeys.state <> 0 Then
                ' Read recordset values.
                iCount = 0
                Do While Not rstShortcutKeys.EOF
                    sShortcutKeys = sShortcutKeys & rstShortcutKeys.fields("shortcutKeys").value

                    Response.Write("<INPUT type='hidden' id=txtShortcutKeys_" & iCount & " name=txtShortcutKeys_" & iCount & " value=""" & Replace(rstShortcutKeys.fields("shortcutKeys").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtShortcutType_" & iCount & " name=txtShortcutType_" & iCount & " value=""" & Replace(rstShortcutKeys.fields("componentType").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtShortcutID_" & iCount & " name=txtShortcutID_" & iCount & " value=""" & Replace(rstShortcutKeys.fields("ID").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtShortcutParams_" & iCount & " name=txtShortcutParams_" & iCount & " value=""" & Replace(rstShortcutKeys.fields("params").value, """", "&quot;") & """>" & vbCrLf)
                    Response.Write("<INPUT type='hidden' id=txtShortcutName_" & iCount & " name=txtShortcutName_" & iCount & " value=""" & Replace(rstShortcutKeys.fields("name").value, """", "&quot;") & """>" & vbCrLf)

                    iCount = iCount + 1
                    rstShortcutKeys.MoveNext()
                Loop
	
                ' Release the ADO recordset object.
                rstShortcutKeys.close()
            End If
            rstShortcutKeys = Nothing
        End If

        Response.Write("<INPUT type='hidden' id=txtShortcutKeys name=txtShortcutKeys value=""" & Replace(sShortcutKeys, """", "&quot;") & """>" & vbCrLf)

        ' Release the ADO command object.
        cmdShortcutKeys = Nothing
	
    %>
</form>


<script type="text/javascript">
    util_def_expression_addhandlers();
    util_def_expression_onload();
</script>
