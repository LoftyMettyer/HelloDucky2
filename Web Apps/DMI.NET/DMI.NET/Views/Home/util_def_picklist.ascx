<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<script src="<%: Url.LatestContent("~/bundles/utilities_picklists")%>" type="text/javascript"></script>


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
																		

																		' Instantiate and initialise the grid. 
																		Response.Write("<OBJECT classid=""clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"" id=ssOleDBGrid name=ssOleDBGrid  codebase=""cabs/COAInt_Grid.cab#version=3,1,3,6"" style=""LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:400px"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""ScrollBars"" VALUE=""4"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""_Version"" VALUE=""196617"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""DataMode"" VALUE=""2"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""Cols"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""Rows"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""BorderStyle"" VALUE=""1"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""RecordSelectors"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""GroupHeaders"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""ColumnHeaders"" VALUE=""-1"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""GroupHeadLines"" VALUE=""1"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""HeadLines"" VALUE=""1"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""FieldDelimiter"" VALUE=""(None)"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""FieldSeparator"" VALUE=""(Tab)"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""Row.Count"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""stylesets.count"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""TagVariant"" VALUE=""EMPTY"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""UseGroups"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""HeadFont3D"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""Font3D"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""DividerType"" VALUE=""3"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""DividerStyle"" VALUE=""1"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""DefColWidth"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""BeveColorScheme"" VALUE=""2"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""BevelColorFrame"" VALUE=""-2147483642"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""BevelColorHighlight"" VALUE=""-2147483628"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""BevelColorShadow"" VALUE=""-2147483632"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""BevelColorFace"" VALUE=""-2147483633"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""CheckBox3D"" VALUE=""-1"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""AllowAddNew"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""AllowDelete"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""AllowUpdate"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""MultiLine"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""ActiveCellStyleSet"" VALUE="""">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""RowSelectionStyle"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""AllowRowSizing"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""AllowGroupSizing"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""AllowColumnSizing"" VALUE=""-1"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""AllowGroupMoving"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""AllowColumnMoving"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""AllowGroupSwapping"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""AllowColumnSwapping"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""AllowGroupShrinking"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""AllowColumnShrinking"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""AllowDragDrop"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""UseExactRowCount"" VALUE=""-1"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""SelectTypeCol"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""SelectTypeRow"" VALUE=""3"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""SelectByCell"" VALUE=""-1"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""BalloonHelp"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""RowNavigation"" VALUE=""1"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""CellNavigation"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""MaxSelectedRows"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""HeadStyleSet"" VALUE="""">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""StyleSet"" VALUE="""">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""ForeColorEven"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""ForeColorOdd"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""BackColorEven"" VALUE=""16777215"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""BackColorOdd"" VALUE=""16777215"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""Levels"" VALUE=""1"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""RowHeight"" VALUE=""503"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""ExtraHeight"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""ActiveRowStyleSet"" VALUE="""">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""CaptionAlignment"" VALUE=""2"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""SplitterPos"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""SplitterVisible"" VALUE=""0"">" & vbCrLf)

																		lngColCount = 0
																		For Each objRow As DataRow In rstFindRecords.Rows
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Width"" VALUE=""3200"">" & vbCrLf)
	
																			If objRow("columnName").ToString() = "ID" Then
																				Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Visible"" VALUE=""0"">" & vbCrLf)
																			Else
																				Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Visible"" VALUE=""-1"">" & vbCrLf)
																			End If
	
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Columns.Count"" VALUE=""1"">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Caption"" VALUE=""" & Replace(objRow("columnName").ToString(), "_", " ") & """>" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Name"" VALUE=""" & objRow("columnName").ToString() & """>" & vbCrLf)
				
																			If (objRow("dataType") = 131) Or (objRow("dataType") = 3) Then
																				Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Alignment"" VALUE=""1"">" & vbCrLf)
																			Else
																				Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Alignment"" VALUE=""0"">" & vbCrLf)
																			End If
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").CaptionAlignment"" VALUE=""3"">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Bound"" VALUE=""0"">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").AllowSizing"" VALUE=""1"">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").DataField"" VALUE=""Column " & lngColCount & """>" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").DataType"" VALUE=""8"">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Level"" VALUE=""0"">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").NumberFormat"" VALUE="""">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Case"" VALUE=""0"">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").FieldLen"" VALUE=""4096"">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").VertScrollBar"" VALUE=""0"">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Locked"" VALUE=""0"">" & vbCrLf)
				
																			If objRow("dataType") = -7 Then
																				' Find column is a logic column.
																				Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Style"" VALUE=""2"">" & vbCrLf)
																			Else
																				Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Style"" VALUE=""0"">" & vbCrLf)
																			End If

																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").ButtonsAlways"" VALUE=""0"">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").RowCount"" VALUE=""0"">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").ColCount"" VALUE=""1"">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").HasHeadForeColor"" VALUE=""0"">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").HasHeadBackColor"" VALUE=""0"">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").HasForeColor"" VALUE=""0"">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").HasBackColor"" VALUE=""0"">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").HeadForeColor"" VALUE=""0"">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").HeadBackColor"" VALUE=""0"">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").ForeColor"" VALUE=""0"">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").BackColor"" VALUE=""0"">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").HeadStyleSet"" VALUE="""">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").StyleSet"" VALUE="""">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Nullable"" VALUE=""1"">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Mask"" VALUE="""">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").PromptInclude"" VALUE=""0"">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").ClipMode"" VALUE=""0"">" & vbCrLf)
																			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").PromptChar"" VALUE=""95"">" & vbCrLf)

																			lngColCount += 1
																		Next
																		
																		Response.Write("	<PARAM NAME=""Columns.Count"" VALUE=""" & lngColCount & """>" & vbCrLf)
																		Response.Write("	<PARAM NAME=""Col.Count"" VALUE=""" & lngColCount & """>" & vbCrLf)

																		Response.Write("	<PARAM NAME=""UseDefaults"" VALUE=""-1"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""TabNavigation"" VALUE=""1"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""_ExtentX"" VALUE=""17330"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""_ExtentY"" VALUE=""1323"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""_StockProps"" VALUE=""79"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""Caption"" VALUE="""">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""ForeColor"" VALUE=""0"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""BackColor"" VALUE=""16777215"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""Enabled"" VALUE=""-1"">" & vbCrLf)
																		Response.Write("	<PARAM NAME=""DataMember"" VALUE="""">" & vbCrLf)

																		Response.Write("</OBJECT>" & vbCrLf)

																		' NB. IMPORTANT ADO NOTE.
																		' When calling a stored procedure which returns a recordset AND has output parameters
																		' you need to close the recordset and set it to nothing before using the output parameters. 
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
																<input type="button" id="cmdFilteredAdd" name="cmdFilteredAdd" class="btn" value="Filtered Add" style="WIDTH: 100%" onclick="filteredAddClick()" />
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
				
					Response.Write("<input type='hidden' id='txtDefn_Name' name='txtDefn_Name' value='" & Replace(prmName.Value.ToString(), """", "&quot;") & "'>" & vbCrLf)
					Response.Write("<input type='hidden' id='txtDefn_Owner' name='txtDefn_Owner' value='" & Replace(prmOwner.Value.ToString(), """", "&quot;") & "'>" & vbCrLf)
					Response.Write("<input type='hidden' id='txtDefn_Description' name='txtDefn_Description' value='" & Replace(prmDescription.Value.ToString(), """", "&quot;") & "'>" & vbCrLf)
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
	util_def_addhandlers();
	util_def_picklist_onload();
</script>


