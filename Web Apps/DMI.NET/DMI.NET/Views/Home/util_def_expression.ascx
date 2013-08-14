<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script src="<%: Url.Content("~/bundles/utilities_expressions")%>" type="text/javascript"></script>


<object classid="clsid:6976CB54-C39B-4181-B1DC-1A829068E2E7" codebase="cabs/COAInt_Client.cab#Version=1,0,0,5"
	id="abExprMenu" name="abExprMenu" style="left: 0px; top: 0px; position: absolute; height: 10px;" viewastext>
	<param name="_ExtentX" value="0">
	<param name="_ExtentY" value="0">
</object>

<object classid="clsid:1C203F13-95AD-11D0-A84B-00A0247B735B" id="SSTreeClipboard" codebase="cabs/SStree.cab#version=1,0,2,24" style="LEFT: 0px; TOP: 0px; WIDTH: 0px; HEIGHT: 0px" viewastext>
	<param name="_ExtentX" value="370">
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

<object classid="clsid:1C203F13-95AD-11D0-A84B-00A0247B735B" id="SSTreeUndo" codebase="cabs/SStree.cab#version=1,0,2,24" style="LEFT: 0px; TOP: 0px; WIDTH: 0px; HEIGHT: 0px" viewastext>
	<param name="_ExtentX" value="370">
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

<form id="frmDefinition">
	<table align="center" cellpadding="5" cellspacing="0" width="100%" height="100%">
		<tr>
			<td>
				<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
					<tr>
						<td width="10"></td>
						<td>
							<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="5">
								<tr valign="top">
									<td>
										<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
											<tr>
												<td colspan="9" height="5"></td>
											</tr>

											<tr height="10">
												<td width="5">&nbsp;</td>
												<td width="10">Name :</td>
												<td width="5">&nbsp;</td>
												<td>
													<input id="txtName" name="txtName" class="text" maxlength="50" style="WIDTH: 100%" onkeyup="changeName()">
												</td>
												<td width="20">&nbsp;</td>
												<td width="10">Owner :</td>
												<td width="5">&nbsp;</td>
												<td width="40%">
													<input id="txtOwner" name="txtOwner" class="text textdisabled" style="WIDTH: 70%" disabled="disabled" tabindex="-1">
												</td>
												<td style="width:50px">&nbsp;</td>
											</tr>

											<tr>
												<td colspan="9" height="5"></td>
											</tr>

											<tr height="10">
												<td width="5">&nbsp;</td>
												<td width="10" nowrap>Description :</td>
												<td width="5">&nbsp;</td>
												<td width="40%" rowspan="5">
													<textarea id="txtDescription" name="txtDescription" class="textarea" style="HEIGHT: 99%; WIDTH: 100%" wrap="VIRTUAL" height="0" maxlength="255" onkeyup="changeDescription()"
														onpaste="var selectedLength = document.selection.createRange().text.length;var pasteData = window.clipboardData.getData('Text');if ((this.value.length + pasteData.length - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}"
														onkeypress="var selectedLength = document.selection.createRange().text.length;if ((this.value.length + 1 - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}">
												</textarea>
												</td>
												<td width="20" nowrap>&nbsp;</td>
												<td width="10">Access :</td>
												<td width="5">&nbsp;</td>
												<td width="40%">
													<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
														<tr>
															<td width="5">
																<input checked id="optAccessRW" name="optAccess" type="radio"
																	onclick="changeAccess()"
																	onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																	onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																	onfocus="try{radio_onFocus(this);}catch(e){}"
																	onblur="try{radio_onBlur(this);}catch(e){}" />
															</td>
															<td width="5">&nbsp;</td>
															<td width="30">
																<label
																	tabindex="-1"
																	for="optAccessRW"
																	class="radio"
																	onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																	onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}" >Read/Write</label>
															</td>
															<td>&nbsp;</td>
														</tr>
													</table>
												</td>
												<td width="5">&nbsp;</td>
											</tr>

											<tr>
												<td colspan="8" height="5"></td>
											</tr>

											<tr height="10">
												<td width="5">&nbsp;</td>

												<td width="10">&nbsp;</td>
												<td width="5">&nbsp;</td>

												<td width="20" nowrap>&nbsp;</td>

												<td width="10">&nbsp;</td>
												<td width="5">&nbsp;</td>
												<td width="40%">
													<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
														<tr>
															<td width="5">
																<input id="optAccessRO" name="optAccess" type="radio"
																	onclick="changeAccess()"
																	onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																	onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																	onfocus="try{radio_onFocus(this);}catch(e){}"
																	onblur="try{radio_onBlur(this);}catch(e){}" />
															</td>
															<td width="5">&nbsp;</td>
															<td width="80" nowrap>
																<label
																	tabindex="-1"
																	for="optAccessRO"
																	class="radio"
																	onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																	onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}" >Read Only</label>
															</td>
															<td>&nbsp;</td>
														</tr>
													</table>
												</td>
												<td width="5">&nbsp;</td>
											</tr>

											<tr>
												<td colspan="8" height="5"></td>
											</tr>

											<tr height="10">
												<td width="5">&nbsp;</td>
												<td width="10">&nbsp;</td>
												<td width="5">&nbsp;</td>
												<td width="20" nowrap>&nbsp;</td>
												<td width="10">&nbsp;</td>
												<td width="5">&nbsp;</td>
												<td width="40%">
													<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
														<tr>
															<td width="5">
																<input id="optAccessHD" name="optAccess" type="radio"
																	onclick="changeAccess()"
																	onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																	onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																	onfocus="try{radio_onFocus(this);}catch(e){}"
																	onblur="try{radio_onBlur(this);}catch(e){}" />
															</td>
															<td width="5">&nbsp;</td>
															<td width="60" nowrap>
																<label
																	tabindex="-1"
																	for="optAccessHD"
																	class="radio"
																	onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																	onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}" >Hidden</label>
															</td>
															<td>&nbsp;</td>
														</tr>
													</table>
												</td>
												<td width="5">&nbsp;</td>
											</tr>

											<tr>
												<td colspan="9">
													<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
														<tr>
															<%--<TD colspan=3 height=30><hr></TD>--%>
															<TD colspan=3 height=10></TD>
														</tr>
														<tr height="10">
															<td rowspan="16">
																<object classid="clsid:1C203F13-95AD-11D0-A84B-00A0247B735B" id="SSTree1"
																	codebase="cabs/SStree.cab#version=1,0,2,24" style="LEFT: 0px; TOP: 0px; WIDTH: 100%; HEIGHT: 400px; VISIBILITY: visible;" viewastext>
																	<param name="_ExtentX" value="30163">
																	<param name="_ExtentY" value="10583">
																	<param name="_Version" value="65538">
																	<param name="BackColor" value="-2147483643">
																	<param name="ForeColor" value="-2147483640">
																	<param name="ImagesMaskColor" value="12632256">
																	<param name="PictureBackgroundMaskColor" value="12632256">
																	<param name="Appearance" value="0">
																	<param name="BorderStyle" value="1">
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
															<td rowspan="16" width="10">&nbsp;</td>
															<td width="80">
																<input type="button" id="cmdAdd" name="cmdAdd" class="btn" value="Add" style="WIDTH: 100%"
																	onclick="addClick()" />
															</td>
														</tr>
														<tr height="10">
															<td>&nbsp;</td>
														</tr>
														<tr height="10">
															<td width="80">
																<input type="button" id="cmdInsert" name="cmdInsert" class="btn" value="Insert" style="WIDTH: 100%"
																	onclick="insertClick()" />
															</td>
														</tr>
														<tr height="10">
															<td>&nbsp;</td>
														</tr>
														<tr height="10">
															<td width="80">
																<input type="button" id="cmdEdit" name="cmdEdit" class="btn" value="Edit"
																	style="WIDTH: 100%"
																	onclick="editClick()" />
															</td>
														</tr>
														<tr height="10">
															<td>&nbsp;</td>
														</tr>
														<tr height="10">
															<td width="80">
																<input type="button" id="cmdDelete" name="cmdDelete" class="btn" value="Delete"
																	style="WIDTH: 100%"
																	onclick="deleteClick()" />
															</td>
														</tr>
														<tr height="10">
															<td>&nbsp;</td>
														</tr>
														<tr height="10">
															<td width="80">
																<input type="button" id="cmdPrint" name="cmdPrint" class="btn" value="Print"
																	style="WIDTH: 100%"
																	onclick="printClick(true)" />
															</td>
														</tr>

														<%
															If Session("utiltype") = 11 Then
														%>
														<tr height="10">
															<td>&nbsp;</td>
														</tr>
														<tr height="10">
															<td width="80">
																<input type="button" id="cmdTest" name="cmdTest" class="btn" value="Test" style="WIDTH: 100%"
																	onclick="testClick()" />
															</td>
														</tr>
														<%	
														End If
														%>
														<tr>
															<td></td>
														</tr>
														<tr height="10">
															<td>&nbsp;</td>
														</tr>
														<tr height="10">
															<td width="80">
																<input type="button" id="cmdOK" name="cmdOK" class="btn" value="OK" style="WIDTH: 100%"
																	onclick="okClick()" />
															</td>
														</tr>
														<tr height="10">
															<td>&nbsp;</td>
														</tr>
														<tr height="10">
															<td width="80">
																<input type="button" id="cmdCancel" name="cmdCancel" class="btn" value="Cancel" style="WIDTH: 100%"
																	onclick="cancelClick()" />
															</td>
														</tr>
													</table>
												</td>
											</tr>

											<tr height="5">
												<td colspan="9" height="5"></td>
											</tr>
										</table>
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

</form>

<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
	<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</form>

<form id="frmOriginalDefinition" style="visibility: hidden; display: none">
	<%
		Dim sReaction As String
		Dim sUtilTypeName As String
		Dim sErrMsg As String
		Dim iCount As Integer
		
		sUtilTypeName = "expression"
		If Session("utiltype") = 11 Then
			sUtilTypeName = "filter"
			sReaction = "FILTERS"
		Else
			If Session("utiltype") = 12 Then
				sUtilTypeName = "calculation"
				sReaction = "CALCULATIONS"
			End If
		End If

		If Session("action") <> "new" Then
			Dim cmdDefn = CreateObject("ADODB.Command")
			cmdDefn.CommandText = "sp_ASRIntGetExpressionDefinition"
			cmdDefn.CommandType = 4	' Stored Procedure
			cmdDefn.ActiveConnection = Session("databaseConnection")

			Dim prmUtilID = cmdDefn.CreateParameter("utilID", 3, 1)	' 3=integer, 1=input
			cmdDefn.Parameters.Append(prmUtilID)
			prmUtilID.value = CleanNumeric(Session("utilid"))

			Dim prmAction = cmdDefn.CreateParameter("action", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
			cmdDefn.Parameters.Append(prmAction)
			prmAction.value = Session("action")

			Dim prmErrMsg = cmdDefn.CreateParameter("errMsg", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
			cmdDefn.Parameters.Append(prmErrMsg)

			Dim prmTimestamp = cmdDefn.CreateParameter("timestamp", 3, 2)	'3=integer, 2=output
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

			If Len(sErrMsg) > 0 Then
				Session("confirmtext") = sErrMsg
				Session("confirmtitle") = "OpenHR Intranet"
				Session("followpage") = "defsel"
				Session("reaction") = sReaction
				Response.Clear()
				Response.Redirect("confirmok")
			End If
		End If
	%>
	<input type="hidden" id="txtOriginalAccess" name="txtOriginalAccess" value="RW">
</form>

<form id="frmUseful" name="frmUseful" style="visibility: hidden; display: none">
	<input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
	<input type="hidden" id="txtLoading" name="txtLoading" value="Y">
	<input type="hidden" id="txtChanged" name="txtChanged" value="0">
	<input type="hidden" id="txtUtilID" name="txtUtilID" value='<% =session("utilid")%>'>
	<input type="hidden" id="txtTableID" name="txtTableID" value='<% =session("utiltableid")%>'>
	<input type="hidden" id="txtAction" name="txtAction" value='<% =session("action")%>'>
	<input type="hidden" id="txtUtilType" name="txtUtilType" value='<% =session("utiltype")%>'>
	<input type="hidden" id="txtLocaleDecimal" name="txtLocaleDecimal" value='<% =session("LocaleDecimalSeparator")%>'>
	<input type="hidden" id="txtExprColourMode" name="txtExprColourMode" value='<% =session("ExprColourMode")%>'>
	<input type="hidden" id="txtExprNodeMode" name="txtExprNodeMode" value='<% =session("ExprNodeMode")%>'>
	<input type="hidden" id="txtLastNode" name="txtLastNode">
	<input type="hidden" id="txtMenuSaved" name="txtMenuSaved" value="0">

	<%
		Dim sErrorDescription As String
				
		Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
	
		Dim cmdBaseTable = CreateObject("ADODB.Command")
		cmdBaseTable.CommandText = "sp_ASRIntGetTableName"
		cmdBaseTable.CommandType = 4 ' Stored Procedure
		cmdBaseTable.ActiveConnection = Session("databaseConnection")

		Dim prmTableID = cmdBaseTable.CreateParameter("tableID", 3, 1) ' 3=integer, 1=input
		cmdBaseTable.Parameters.Append(prmTableID)
		prmTableID.value = CleanNumeric(Session("utiltableid"))

		Dim prmTableName = cmdBaseTable.CreateParameter("tableName", 200, 2, 255)
		cmdBaseTable.Parameters.Append(prmTableName)

		Err.Clear()
		cmdBaseTable.Execute()
			
		Response.Write("<INPUT type='hidden' id=txtTableName name=txtTableName value=""" & cmdBaseTable.Parameters("tableName").Value & """>" & vbCrLf)

		' Release the ADO command object.
		cmdBaseTable = Nothing
	
	%>
	<input type="hidden" id="txtCanDelete" name="txtCanDelete" value="0">
	<input type="hidden" id="txtCanInsert" name="txtCanInsert" value="0">
	<input type="hidden" id="txtCanCut" name="txtCanCut" value="0">
	<input type="hidden" id="txtCanCopy" name="txtCanCopy" value="0">
	<input type="hidden" id="txtCanPaste" name="txtCanPaste" value="0">
	<input type="hidden" id="txtCanMoveUp" name="txtCanMoveUp" value="0">
	<input type="hidden" id="txtCanMoveDown" name="txtCanMoveDown" value="0">
	<input type="hidden" id="txtUndoType" name="txtUndoType" value="">
	<input type="hidden" id="txtOldText" name="txtOldText" value="">
</form>

<form id="frmValidate" name="frmValidate" target="validate" method="post" action="util_validate_expression" style="visibility: hidden; display: none">
	<input type="hidden" id="validatePass" name="validatePass" value="0">
	<input type="hidden" id="validateName" name="validateName" value=''>
	<input type="hidden" id="validateOwner" name="validateOwner" value=''>
	<input type="hidden" id="validateTimestamp" name="validateTimestamp" value=''>
	<input type="hidden" id="validateUtilID" name="validateUtilID" value=''>
	<input type="hidden" id="validateUtilType" name="validateUtilType" value=''>
	<input type="hidden" id="validateAccess" name="validateAccess" value=''>
	<input type="hidden" id="components1" name="components1" value="">
	<input type="hidden" id="validateBaseTableID" name="validateBaseTableID" value='<%=session("utiltableid")%>'>
	<input type="hidden" id="validateOriginalAccess" name="validateOriginalAccess" value="RW">
</form>

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

<form id="frmTest" name="frmTest" target="test" method="post" action="util_test_expression_pval" style="visibility: hidden; display: none">
	<input type="hidden" id="type" name="type">
	<input type="hidden" id="Hidden1" name="components1">
	<input type="hidden" id="tableID" name="tableID" value='<% =session("utiltableid")%>'>
	<input type="hidden" id="prompts" name="prompts">
	<input type="hidden" id="filtersAndCalcs" name="filtersAndCalcs">
</form>

<input type='hidden' id="txtTicker" name="txtTicker" value="0">
<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">

<form id="frmShortcutKeys" name="frmShortcutKeys" style="visibility: hidden; display: none">
	<%
		Dim sShortcutKeys As String
		
		sShortcutKeys = ""
	
		Dim cmdShortcutKeys = CreateObject("ADODB.Command")
		cmdShortcutKeys.CommandText = "spASRIntGetOpFuncShortcuts"
		cmdShortcutKeys.CommandType = 4	' Stored Procedure
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
