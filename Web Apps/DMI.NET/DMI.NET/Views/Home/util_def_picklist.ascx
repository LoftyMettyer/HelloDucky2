<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<script src="<%: Url.LatestContent("~/bundles/utilities_picklists")%>" type="text/javascript"></script>
<script src="<%: Url.LatestContent("~/bundles/jQueryUI7")%>" type="text/javascript"></script>

<%--licence manager reference for activeX--%>

<form id="frmDefinition">
	<table align="center" class="outline" cellpadding="5" cellspacing="0" width="100%" height="100%">
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
													<input id="txtOwner" name="txtOwner" class="text textdisabled" style="WIDTH: 100%" disabled="disabled" tabindex="-1">
												</td>
												<td width="5">&nbsp;</td>
											</tr>

											<tr>
												<td colspan="9" height="5"></td>
											</tr>

											<tr height="10">
												<td width="5">&nbsp;</td>
												<td width="10" nowrap>Description :</td>
												<td width="5">&nbsp;</td>
												<td width="40%" rowspan="5">
													<textarea id="txtDescription" name="txtDescription" class="textarea" style="HEIGHT: 99%; WIDTH: 100%" wrap="VIRTUAL" height="0" maxlength="255"
														onkeyup="changeDescription()">
												</textarea>
												</td>
												<td width="20" nowrap>&nbsp;</td>
												<td width="10">Access :</td>
												<td width="5">&nbsp;</td>
												<td width="40%">
													<table border="0" cellspacing="0" cellpadding="0" width="100%">
														<tr>
															<td width="5">
																<input checked id="optAccessRW" name="optAccess" type="radio"
																	onclick="changeAccess()" />
															</td>
															<td width="5">&nbsp;</td>
															<td width="30">
																<label tabindex="-1" for="optAccessRW" class="radio">
																	Read/Write
																</label>
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
																<input id="optAccessRO" name="optAccess" type="radio" onclick="changeAccess()" />
															</td>
															<td width="5">&nbsp;</td>
															<td width="80" nowrap>
																<label tabindex="-1" for="optAccessRO" class="radio">
																	Read Only
																</label>
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
																<input id="optAccessHD" name="optAccess" type="radio" onclick="changeAccess()" />
															</td>
															<td width="5">&nbsp;</td>
															<td width="60" nowrap>
																<label tabindex="-1" for="optAccessHD" class="radio">
																	Hidden
																</label>
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
															<td colspan="3" height="30">
																<hr>
															</td>
														</tr>
														<tr height="10">
															<td rowspan="14">
																<%
																																	
																	' Get the employee find columns.
																	Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

																	Dim sErrorDescription As String
																	Dim lngColCount As Long															

																	Try
																		Dim prmErrMsg = New SqlParameter("psErrorMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
																		Dim prm1000SepCols = New SqlParameter("ps1000SeparatorCols", SqlDbType.VarChar, 8000) With {.Direction = ParameterDirection.Output}

																		Dim rstFindRecords = objDataAccess.GetFromSP("sp_ASRIntGetDefaultOrderColumns" _
																				, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("utiltableid"))} _
																				, prmErrMsg, prm1000SepCols)
																		

																		
																		If Len(prmErrMsg.Value) > 0 Then
																			Session("ErrorTitle") = "Picklist Definition Page"
																			Session("ErrorText") = prmErrMsg.Value
																			Response.Clear()
			
																			'Response.Redirect("error.asp")
																			Response.Redirect("FormError")
			
																		Else
																			Response.Write("<INPUT type='hidden' id=txt1000SepCols name=txt1000SepCols value=""" & prm1000SepCols.Value & """>" & vbCrLf)
																		End If

																	Catch ex As Exception
																		sErrorDescription = "The find columns could not be retrieved." & vbCrLf & FormatError(ex.Message)

																	End Try

																%>
																<div id="PickListGrid" style="height: 400px; margin-bottom: 50px; width: 75%;">
																	<table id="ssOleDBGrid" style="width: 100%"></table>
																</div>
															</td>
															<td rowspan="14" width="10">&nbsp;</td>
															<td width="100">
																<input type="button" id="cmdAdd" name="cmdAdd" class="btn" value="Add" style="WIDTH: 100%" onclick="addClick()" />
															</td>
														</tr>
														<tr height="10">
															<td></td>
														</tr>
														<tr height="10">
															<td width="100">
																<input type="button" id="cmdAddAll" name="cmdAddAll" class="btn" value="Add All" style="WIDTH: 100%" onclick="addAllClick()" />
															</td>
														</tr>
														<tr height="10">
															<td></td>
														</tr>
														<tr height="10">
															<td width="100">
																<input type="button" id="cmdFilteredAdd" disabled="disabled" name="cmdFilteredAdd" class="btn" value="Filtered Add" style="WIDTH: 100%" onclick="filteredAddClick()" />
															</td>
														</tr>
														<tr height="10">
															<td></td>
														</tr>
														<tr height="10">
															<td width="100">
																<input type="button" id="cmdRemove" name="cmdRemove" class="btn" value="Remove" style="WIDTH: 100%" onclick="removeClick()" />
															</td>
														</tr>
														<tr height="10">
															<td></td>
														</tr>
														<tr height="10">
															<td width="100">
																<input type="button" id="cmdRemoveAll" name="cmdRemoveAll" class="btn" value="Remove All" style="WIDTH: 100%" onclick="removeAllClick()" />
															</td>
														</tr>
														<tr height="10">
															<td></td>
														</tr>
														<tr>
															<td></td>
														</tr>
														<tr height="10">
															<td width="100">
																<input type="button" id="cmdOK" name="cmdOK" class="btn" value="OK" style="WIDTH: 100%" onclick="okClick()" />
															</td>
														</tr>
														<tr height="10">
															<td></td>
														</tr>
														<tr height="10">
															<td width="100">
																<input type="button" id="cmdCancel" name="cmdCancel" class="btn" value="Cancel" style="WIDTH: 100%" onclick="cancelClick()" />
															</td>
														</tr>
													</table>
													<div id="RecordCountDIV"></div>
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
		Dim sErrMsg As String
		Dim sSelectedRecords As String
		
		sErrMsg = ""

		If Session("action") <> "new" Then
			
			Try

				Dim prmErrMsg = New SqlParameter("psErrorMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
				Dim prmName = New SqlParameter("psPicklistName", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmOwner = New SqlParameter("psPicklistOwner", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmDescription = New SqlParameter("psPicklistDesc", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmAccess = New SqlParameter("psAccess", SqlDbType.VarChar, 255) With {.Direction = ParameterDirection.Output}
				Dim prmTimestamp = New SqlParameter("piTimestamp", SqlDbType.Int) With {.Direction = ParameterDirection.Output}

				Dim rstDefinition = objDataAccess.GetFromSP("sp_ASRIntGetPicklistDefinition" _
					, New SqlParameter("piPicklistID", SqlDbType.Int) With {.Value = CleanNumeric(Session("utilid"))} _
					, New SqlParameter("psAction", SqlDbType.VarChar, 255) With {.Value = Session("action")} _
					, prmErrMsg, prmName, prmOwner, prmDescription, prmAccess, prmTimestamp)
		
			
				sSelectedRecords = "0"
				Response.Write("<input type='hidden' id='txtSelectedRecords' name='txtSelectedRecords' value='" & sSelectedRecords & "'>" & vbCrLf)
				
				If Len(prmErrMsg.Value) > 0 Then
					sErrMsg = "'" & Session("utilname") & "' " & prmErrMsg.Value
				Else
				
					'Response.Write("<input type='hidden' id='txtDefn_Name' name='txtDefn_Name' value='" & Replace(prmName.Value.ToString(), """", "&quot;") & "'>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Replace(prmName.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id='txtDefn_Owner' name='txtDefn_Owner' value=""" & Replace(prmOwner.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id='txtDefn_Description' name='txtDefn_Description' value=""" & Replace(prmDescription.Value.ToString(), """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id='txtDefn_Access' name='txtDefn_Access' value='" & prmAccess.Value & "'>" & vbCrLf)
					Response.Write("<input type='hidden' id='txtDefn_Timestamp' name='txtDefn_Timestamp' value='" & prmTimestamp.Value & "'>" & vbCrLf)
				End If

			Catch ex As Exception
				sErrMsg = "'" & Session("utilname") & "' picklist definition could not be read." & vbCrLf & FormatError(ex.Message)

				
			End Try

		End If
	%>
</form>

<form id="frmUseful" name="frmUseful" style="visibility: hidden; display: none">
	<input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
	<input type="hidden" id="txtLoading" name="txtLoading" value="Y">
	<input type="hidden" id="txtChanged" name="txtChanged" value="0">
	<input type="hidden" id="txtUtilID" name="txtUtilID" value='<% =session("utilid")%>'>
	<input type="hidden" id="txtTableID" name="txtTableID" value='<% =session("utiltableid")%>'>
	<input type="hidden" id="txtAction" name="txtAction" value='<% =session("action")%>'>
	<%
		Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
	%>
</form>

<form id="frmValidate" name="frmValidate" method="post" action="util_validate_picklist" style="visibility: hidden; display: none">
	<input type="hidden" id="validatePass" name="validatePass" value="0">
	<input type="hidden" id="validateName" name="validateName" value=''>
	<input type="hidden" id="validateTimestamp" name="validateTimestamp" value=''>
	<input type="hidden" id="validateUtilID" name="validateUtilID" value=''>
	<input type="hidden" id="validateAccess" name="validateAccess" value=''>
	<input type="hidden" id="validateBaseTableID" name="validateBaseTableID" value='<%=session("utiltableid")%>'>
</form>

<form id="frmSend" name="frmSend" method="post" action="util_def_picklist_Submit" style="visibility: hidden; display: none">
	<input type="hidden" id="txtSend_ID" name="txtSend_ID">
	<input type="hidden" id="txtSend_name" name="txtSend_name">
	<input type="hidden" id="txtSend_description" name="txtSend_description">
	<input type="hidden" id="txtSend_access" name="txtSend_access">
	<input type="hidden" id="txtSend_userName" name="txtSend_userName">
	<input type="hidden" id="txtSend_columns" name="txtSend_columns">
	<input type="hidden" id="txtSend_columns2" name="txtSend_columns2">
	<input type="hidden" id="txtSend_reaction" name="txtSend_reaction">
	<input type="hidden" id="txtSend_tableID" name="txtSend_tableID" value='<% =session("utiltableid")%>'>
</form>

<input type='hidden' id="txtTicker" name="txtTicker" value="0">
<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">

<form id="frmPicklistSelection" name="frmPicklistSelection" action="picklistSelectionMain" method="post" style="visibility: hidden; display: none">
	<input type="hidden" id="selectionType" name="selectionType">
	<input type="hidden" id="txtTableID" name="txtTableID" value='<% =session("utiltableid")%>'>
	<input type="hidden" id="selectedIDs1" name="selectedIDs1">
</form>

<script type="text/javascript">
	util_def_picklist_onload();
</script>


