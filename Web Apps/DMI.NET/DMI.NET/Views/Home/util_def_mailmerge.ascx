<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script src="<%: Url.Content("~/Scripts/FormScripts/mailmergedef.js")%>" type="text/javascript"></script>

<%  
	'Dim iVersionOneEnabled = 0
	Dim cmdVersionOneModule = CreateObject("ADODB.Command")
	cmdVersionOneModule.CommandText = "spASRIntActivateModule"
	cmdVersionOneModule.CommandType = 4	' Stored Procedure
	cmdVersionOneModule.ActiveConnection = Session("databaseConnection")
	cmdVersionOneModule.CommandTimeout = 300

	Dim prmModuleKey = cmdVersionOneModule.CreateParameter("moduleKey", 200, 1, 50)	'200=varchar, 1=input, 50=size
	cmdVersionOneModule.Parameters.Append(prmModuleKey)
	prmModuleKey.value = "VERSIONONE"

	Dim prmEnabled = cmdVersionOneModule.CreateParameter("enabled", 11, 2) ' 11=bit, 2=output
	cmdVersionOneModule.Parameters.Append(prmEnabled)

	Err.Number = 0
	cmdVersionOneModule.Execute()

	Dim iVersionOneEnabled = CInt(cmdVersionOneModule.Parameters("enabled").Value)
	If iVersionOneEnabled < 0 Then
		iVersionOneEnabled = 1
	End If
	cmdVersionOneModule = Nothing
%>

<object classid="clsid:F9043C85-F6F2-101A-A3C9-08002B2F49FB"id="dialog" codebase="cabs/comdlg32.cab#Version=1,0,0,0" style="LEFT: 0px; TOP: 0px" viewastext>
	<param name="_ExtentX" value="847">
	<param name="_ExtentY" value="847">
	<param name="_Version" value="393216">
	<param name="CancelError" value="0">
	<param name="Color" value="0">
	<param name="Copies" value="1">
	<param name="DefaultExt" value="">
	<param name="DialogTitle" value="">
	<param name="FileName" value="">
	<param name="Filter" value="">
	<param name="FilterIndex" value="0">
	<param name="Flags" value="0">
	<param name="FontBold" value="0">
	<param name="FontItalic" value="0">
	<param name="FontName" value="">
	<param name="FontSize" value="8">
	<param name="FontStrikeThru" value="0">
	<param name="FontUnderLine" value="0">
	<param name="FromPage" value="0">
	<param name="HelpCommand" value="0">
	<param name="HelpContext" value="0">
	<param name="HelpFile" value="">
	<param name="HelpKey" value="">
	<param name="InitDir" value="">
	<param name="Max" value="0">
	<param name="Min" value="0">
	<param name="MaxFileSize" value="260">
	<param name="PrinterDefault" value="1">
	<param name="ToPage" value="0">
	<param name="Orientation" value="1">
</object>

<div <%=session("BodyTag")%>>
	<form id="frmDefinition" name="frmDefinition">

		<table valign="top" align="center" class="outline" cellpadding="5" cellspacing="0" width="100%" height="100%">
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
									onclick="displayPage(1)"
									onmouseover="try{button_onMouseOver(this);}catch(e){}"
									onmouseout="try{button_onMouseOut(this);}catch(e){}"
									onfocus="try{button_onFocus(this);}catch(e){}"
									onblur="try{button_onBlur(this);}catch(e){}" />
								<input type="button" value="Columns" id="btnTab2" name="btnTab2" class="btn"
									onclick="displayPage(2)"
									onmouseover="try{button_onMouseOver(this);}catch(e){}"
									onmouseout="try{button_onMouseOut(this);}catch(e){}"
									onfocus="try{button_onFocus(this);}catch(e){}"
									onblur="try{button_onBlur(this);}catch(e){}" />
								<input type="button" value="Sort Order" id="btnTab3" name="btnTab3" class="btn"
									onclick="displayPage(3)"
									onmouseover="try{button_onMouseOver(this);}catch(e){}"
									onmouseout="try{button_onMouseOut(this);}catch(e){}"
									onfocus="try{button_onFocus(this);}catch(e){}"
									onblur="try{button_onBlur(this);}catch(e){}" />
								<input type="button" value="Output" id="btnTab4" name="btnTab4" class="btn"
									onclick="displayPage(4)"
									onmouseover="try{button_onMouseOver(this);}catch(e){}"
									onmouseout="try{button_onMouseOut(this);}catch(e){}"
									onfocus="try{button_onFocus(this);}catch(e){}"
									onblur="try{button_onBlur(this);}catch(e){}" />
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
												<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
													<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
														<tr height="10">
															<td width="5">&nbsp;</td>
															<td width="10">Name :</td>
															<td width="5">&nbsp;</td>
															<td>
																<input id="txtName" name="txtName" maxlength="50" style="WIDTH: 100%" class="text"
																	onkeyup="changeName()">
															</td>
															<td width="20">&nbsp;</td>
															<td width="10">Owner :</td>
															<td width="5">&nbsp;</td>
															<td width="40%">
																<input id="txtOwner" name="txtOwner" style="WIDTH: 100%" disabled="disabled" class="text textdisabled">
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
															<td width="40%" rowspan="3">
																<textarea id="txtDescription" name="txtDescription" class="textarea" style="HEIGHT: 99%; WIDTH: 100%" height="0" maxlength="255"
																	onkeyup="changeDescription()"
																	onpaste="var selectedLength = document.selection.createRange().text.length;var pasteData = window.clipboardData.getData('Text');if ((this.value.length + pasteData.length - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}"
																	onkeypress="var selectedLength = document.selection.createRange().text.length;if ((this.value.length + 1 - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}">
																</textarea>
															</td>
															<td width="20" nowrap>&nbsp;</td>
															<td width="10" valign="top">Access :</td>
															<td width="5">&nbsp;</td>
															<td width="40%" rowspan="3" valign="top"></td>
												<TD width="40%" rowspan="3" valign=top>
													<%Html.RenderPartial("Util_Def_CustomReports/grdaccess")%>         
												</TD>
															<td width="5">&nbsp;</td>
														</tr>

														<tr height="10">
															<td colspan="7">&nbsp;</td>
														</tr>

														<tr height="10">
															<td colspan="7">&nbsp;</td>
														</tr>

														<tr>
															<td colspan="9">
																<hr>
															</td>
														</tr>

														<tr height="10">
															<td width="5">&nbsp;</td>
															<td width="100" nowrap valign="top">Base Table :</td>
															<td width="5">&nbsp;</td>
															<td width="40%" valign="top">
																<select id="cboBaseTable" name="cboBaseTable" class="combo" style="WIDTH: 100%"
																	onchange="changeBaseTable()">
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
																				onclick="changeBaseTableRecordOptions()"
																				onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																				onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																				onfocus="try{radio_onFocus(this);}catch(e){}"
																				onblur="try{radio_onBlur(this);}catch(e){}" />
																		</td>
																		<td width="5">&nbsp;</td>
																		<td width="30">
																			<label
																				tabindex="-1"
																				for="optRecordSelection1"
																				class="radio"
																				onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																				onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
																				All
																			</label>
																		</td>
																		<td>&nbsp;</td>
																	</tr>
																	<tr>
																		<td width="5">
																			<input id="optRecordSelection2" name="optRecordSelection" type="radio"
																				onclick="changeBaseTableRecordOptions()"
																				onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																				onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																				onfocus="try{radio_onFocus(this);}catch(e){}"
																				onblur="try{radio_onBlur(this);}catch(e){}" />
																		</td>
																		<td width="5">&nbsp;</td>
																		<td width="20">
																			<label
																				tabindex="-1"
																				for="optRecordSelection2"
																				class="radio"
																				onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																				onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
																				Picklist</label>
																		</td>
																		<td width="5">&nbsp;</td>
																		<td>
																			<input id="txtBasePicklist" name="txtBasePicklist" disabled="disabled" style="WIDTH: 100%" class="text textdisabled">
																		</td>
																		<td width="30">
																			<input id="cmdBasePicklist" name="cmdBasePicklist" style="WIDTH: 100%" type="button" value="..." class="btn"
																				onclick="selectRecordOption('base', 'picklist')"
																				onmouseover="try{button_onMouseOver(this);}catch(e){}"
																				onmouseout="try{button_onMouseOut(this);}catch(e){}"
																				onfocus="try{button_onFocus(this);}catch(e){}"
																				onblur="try{button_onBlur(this);}catch(e){}" />
																		</td>
																	</tr>
																	<tr>
																		<td width="5">
																			<input id="optRecordSelection3" name="optRecordSelection" type="radio"
																				onclick="changeBaseTableRecordOptions()"
																				onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																				onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																				onfocus="try{radio_onFocus(this);}catch(e){}"
																				onblur="try{radio_onBlur(this);}catch(e){}" />
																		</td>
																		<td width="5">&nbsp;</td>
																		<td width="20">
																			<label
																				tabindex="-1"
																				for="optRecordSelection3"
																				class="radio"
																				onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																				onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
																				Filter</label>
																		</td>
																		<td width="5">&nbsp;</td>
																		<td>
																			<input id="txtBaseFilter" name="txtBaseFilter" disabled="disabled" class="text textdisabled" style="WIDTH: 100%">
																		</td>
																		<td width="30">
																			<input id="cmdBaseFilter" name="cmdBaseFilter" style="WIDTH: 100%" type="button" value="..." class="btn"
																				onclick="selectRecordOption('base', 'filter')"
																				onmouseover="try{button_onMouseOver(this);}catch(e){}"
																				onmouseout="try{button_onMouseOut(this);}catch(e){}"
																				onfocus="try{button_onFocus(this);}catch(e){}"
																				onblur="try{button_onBlur(this);}catch(e){}" />
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
																<input id="txtParent1" name="txtParent1" style="WIDTH: 100%" disabled="disabled" class="text textdisabled" type="hidden">
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
																<input id="txtParent2" name="txtParent2" style="WIDTH: 100%" disabled="disabled" class="text textdisabled" type="hidden">
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
												</table>
									</table>
								</div>

								<!-- Second tab -->
								<div id="div2" style="visibility: hidden; display: none">
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
															<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																<tr height="5">
																	<td height="5" colspan="7" width="100%">
																		<select id="cboTblAvailable" name="cboTblAvailable" disabled="disabled" class="combo combodisabled" style="WIDTH: 100%; HEIGHT: 100%"
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
																			onclick="refreshAvailableColumns();"
																			onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																			onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																			onfocus="try{radio_onFocus(this);}catch(e){}"
																			onblur="try{radio_onBlur(this);}catch(e){}" />
																	</td>
																	<td height="5" width="5">
																		<label
																			tabindex="-1"
																			for="optColumns"
																			class="radio radiodisabled"
																			onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																			onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}" >Columns</label>
																	</td>
																	<td width="5" height="5"></td>
																	<td height="5">
																		<input id="optCalc" name="optAvailType" type="radio" disabled="disabled"
																			onclick="refreshAvailableColumns();"
																			onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																			onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																			onfocus="try{radio_onFocus(this);}catch(e){}"
																			onblur="try{radio_onBlur(this);}catch(e){}" />
																	</td>
																	<td width="5" height="5">
																		<label
																			tabindex="-1"
																			for="optCalc"
																			class="radio radiodisabled"
																			onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																			onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}" >Calculations</label>
																	</td>
																	<td height="5"></td>
<%--																	<tr height="10">
																		<td height="10" colspan="7" width="100%"></td>
																	</tr>--%>
																</tr>
															</table>
														</td>
														<td width="10"></td>
														<td width="5" nowrap></td>
														<td width="10"></td>
														<td rowspan="4" width="40%" height="100%">
															<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
																<tr>
																	<td width="100%" height="100%">
																		<object classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"
																			codebase="cabs/COAInt_Grid.cab#version=3,1,3,6"
																			id="ssOleDBGridSelectedColumns"
																			name="ssOleDBGridSelectedColumns"
																			style="HEIGHT: 100%; LEFT: 0; TOP: 0; WIDTH: 100%; height:100%">
																			<param name="ScrollBars" value="2">
																			<param name="_Version" value="196617">
																			<param name="DataMode" value="2">
																			<param name="Cols" value="0">
																			<param name="Rows" value="0">
																			<param name="BorderStyle" value="1">
																			<param name="RecordSelectors" value="0">
																			<param name="GroupHeaders" value="0">
																			<param name="ColumnHeaders" value="-1">
																			<param name="GroupHeadLines" value="1">
																			<param name="HeadLines" value="1">
																			<param name="FieldDelimiter" value="(None)">
																			<param name="FieldSeparator" value="(Tab)">
																			<param name="Row.Count" value="0">
																			<param name="Col.Count" value="12">
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
																			<param name="AllowColumnSizing" value="-1">
																			<param name="AllowGroupMoving" value="0">
																			<param name="AllowColumnMoving" value="0">
																			<param name="AllowGroupSwapping" value="0">
																			<param name="AllowColumnSwapping" value="0">
																			<param name="AllowGroupShrinking" value="0">
																			<param name="AllowColumnShrinking" value="0">
																			<param name="AllowDragDrop" value="0">
																			<param name="UseExactRowCount" value="-1">
																			<param name="SelectTypeCol" value="0">
																			<param name="SelectTypeRow" value="3">
																			<param name="SelectByCell" value="-1">
																			<param name="BalloonHelp" value="0">
																			<param name="RowNavigation" value="1">
																			<param name="CellNavigation" value="0">
																			<param name="MaxSelectedRows" value="0">
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
																			<param name="Columns.Count" value="14">

																			<param name="Columns(0).Width" value="0">
																			<param name="Columns(0).Visible" value="0">
																			<param name="Columns(0).Columns.Count" value="1">
																			<param name="Columns(0).Caption" value="type">
																			<param name="Columns(0).Name" value="type">
																			<param name="Columns(0).Alignment" value="0">
																			<param name="Columns(0).CaptionAlignment" value="3">
																			<param name="Columns(0).Bound" value="0">
																			<param name="Columns(0).AllowSizing" value="1">
																			<param name="Columns(0).DataField" value="Column 0">
																			<param name="Columns(0).DataType" value="8">
																			<param name="Columns(0).Level" value="0">
																			<param name="Columns(0).NumberFormat" value="">
																			<param name="Columns(0).Case" value="0">
																			<param name="Columns(0).FieldLen" value="4096">
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
																			<param name="Columns(1).Caption" value="tableID">
																			<param name="Columns(1).Name" value="tableID">
																			<param name="Columns(1).Alignment" value="0">
																			<param name="Columns(1).CaptionAlignment" value="3">
																			<param name="Columns(1).Bound" value="0">
																			<param name="Columns(1).AllowSizing" value="1">
																			<param name="Columns(1).DataField" value="Column 1">
																			<param name="Columns(1).DataType" value="8">
																			<param name="Columns(1).Level" value="0">
																			<param name="Columns(1).NumberFormat" value="">
																			<param name="Columns(1).Case" value="0">
																			<param name="Columns(1).FieldLen" value="4096">
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

																			<param name="Columns(2).Width" value="0">
																			<param name="Columns(2).Visible" value="0">
																			<param name="Columns(2).Columns.Count" value="1">
																			<param name="Columns(2).Caption" value="columnID">
																			<param name="Columns(2).Name" value="columnID">
																			<param name="Columns(2).Alignment" value="0">
																			<param name="Columns(2).CaptionAlignment" value="3">
																			<param name="Columns(2).Bound" value="0">
																			<param name="Columns(2).AllowSizing" value="1">
																			<param name="Columns(2).DataField" value="Column 2">
																			<param name="Columns(2).DataType" value="8">
																			<param name="Columns(2).Level" value="0">
																			<param name="Columns(2).NumberFormat" value="">
																			<param name="Columns(2).Case" value="0">
																			<param name="Columns(2).FieldLen" value="4096">
																			<param name="Columns(2).VertScrollBar" value="0">
																			<param name="Columns(2).Locked" value="0">
																			<param name="Columns(2).Style" value="0">
																			<param name="Columns(2).ButtonsAlways" value="0">
																			<param name="Columns(2).RowCount" value="0">
																			<param name="Columns(2).ColCount" value="1">
																			<param name="Columns(2).HasHeadForeColor" value="0">
																			<param name="Columns(2).HasHeadBackColor" value="0">
																			<param name="Columns(2).HasForeColor" value="0">
																			<param name="Columns(2).HasBackColor" value="0">
																			<param name="Columns(2).HeadForeColor" value="0">
																			<param name="Columns(2).HeadBackColor" value="0">
																			<param name="Columns(2).ForeColor" value="0">
																			<param name="Columns(2).BackColor" value="0">
																			<param name="Columns(2).HeadStyleSet" value="">
																			<param name="Columns(2).StyleSet" value="">
																			<param name="Columns(2).Nullable" value="1">
																			<param name="Columns(2).Mask" value="">
																			<param name="Columns(2).PromptInclude" value="0">
																			<param name="Columns(2).ClipMode" value="0">
																			<param name="Columns(2).PromptChar" value="95">

																			<param name="Columns(3).Width" value="100000">
																			<param name="Columns(3).Visible" value="-1">
																			<param name="Columns(3).Columns.Count" value="1">
																			<param name="Columns(3).Caption" value="Columns / Calculations Selected">
																			<param name="Columns(3).Name" value="display">
																			<param name="Columns(3).Alignment" value="0">
																			<param name="Columns(3).CaptionAlignment" value="3">
																			<param name="Columns(3).Bound" value="0">
																			<param name="Columns(3).AllowSizing" value="1">
																			<param name="Columns(3).DataField" value="Column 3">
																			<param name="Columns(3).DataType" value="8">
																			<param name="Columns(3).Level" value="0">
																			<param name="Columns(3).NumberFormat" value="">
																			<param name="Columns(3).Case" value="0">
																			<param name="Columns(3).FieldLen" value="4096">
																			<param name="Columns(3).VertScrollBar" value="0">
																			<param name="Columns(3).Locked" value="0">
																			<param name="Columns(3).Style" value="0">
																			<param name="Columns(3).ButtonsAlways" value="0">
																			<param name="Columns(3).RowCount" value="0">
																			<param name="Columns(3).ColCount" value="1">
																			<param name="Columns(3).HasHeadForeColor" value="0">
																			<param name="Columns(3).HasHeadBackColor" value="0">
																			<param name="Columns(3).HasForeColor" value="0">
																			<param name="Columns(3).HasBackColor" value="0">
																			<param name="Columns(3).HeadForeColor" value="0">
																			<param name="Columns(3).HeadBackColor" value="0">
																			<param name="Columns(3).ForeColor" value="0">
																			<param name="Columns(3).BackColor" value="0">
																			<param name="Columns(3).HeadStyleSet" value="">
																			<param name="Columns(3).StyleSet" value="">
																			<param name="Columns(3).Nullable" value="1">
																			<param name="Columns(3).Mask" value="">
																			<param name="Columns(3).PromptInclude" value="0">
																			<param name="Columns(3).ClipMode" value="0">
																			<param name="Columns(3).PromptChar" value="95">

																			<param name="Columns(4).Width" value="0">
																			<param name="Columns(4).Visible" value="0">
																			<param name="Columns(4).Columns.Count" value="1">
																			<param name="Columns(4).Caption" value="size">
																			<param name="Columns(4).Name" value="size">
																			<param name="Columns(4).Alignment" value="0">
																			<param name="Columns(4).CaptionAlignment" value="3">
																			<param name="Columns(4).Bound" value="0">
																			<param name="Columns(4).AllowSizing" value="1">
																			<param name="Columns(4).DataField" value="Column 4">
																			<param name="Columns(4).DataType" value="8">
																			<param name="Columns(4).Level" value="0">
																			<param name="Columns(4).NumberFormat" value="">
																			<param name="Columns(4).Case" value="0">
																			<param name="Columns(4).FieldLen" value="4096">
																			<param name="Columns(4).VertScrollBar" value="0">
																			<param name="Columns(4).Locked" value="0">
																			<param name="Columns(4).Style" value="0">
																			<param name="Columns(4).ButtonsAlways" value="0">
																			<param name="Columns(4).RowCount" value="0">
																			<param name="Columns(4).ColCount" value="1">
																			<param name="Columns(4).HasHeadForeColor" value="0">
																			<param name="Columns(4).HasHeadBackColor" value="0">
																			<param name="Columns(4).HasForeColor" value="0">
																			<param name="Columns(4).HasBackColor" value="0">
																			<param name="Columns(4).HeadForeColor" value="0">
																			<param name="Columns(4).HeadBackColor" value="0">
																			<param name="Columns(4).ForeColor" value="0">
																			<param name="Columns(4).BackColor" value="0">
																			<param name="Columns(4).HeadStyleSet" value="">
																			<param name="Columns(4).StyleSet" value="">
																			<param name="Columns(4).Nullable" value="1">
																			<param name="Columns(4).Mask" value="">
																			<param name="Columns(4).PromptInclude" value="0">
																			<param name="Columns(4).ClipMode" value="0">
																			<param name="Columns(4).PromptChar" value="95">

																			<param name="Columns(5).Width" value="0">
																			<param name="Columns(5).Visible" value="0">
																			<param name="Columns(5).Columns.Count" value="1">
																			<param name="Columns(5).Caption" value="decimals">
																			<param name="Columns(5).Name" value="decimals">
																			<param name="Columns(5).Alignment" value="0">
																			<param name="Columns(5).CaptionAlignment" value="3">
																			<param name="Columns(5).Bound" value="0">
																			<param name="Columns(5).AllowSizing" value="1">
																			<param name="Columns(5).DataField" value="Column 5">
																			<param name="Columns(5).DataType" value="8">
																			<param name="Columns(5).Level" value="0">
																			<param name="Columns(5).NumberFormat" value="">
																			<param name="Columns(5).Case" value="0">
																			<param name="Columns(5).FieldLen" value="4096">
																			<param name="Columns(5).VertScrollBar" value="0">
																			<param name="Columns(5).Locked" value="0">
																			<param name="Columns(5).Style" value="0">
																			<param name="Columns(5).ButtonsAlways" value="0">
																			<param name="Columns(5).RowCount" value="0">
																			<param name="Columns(5).ColCount" value="1">
																			<param name="Columns(5).HasHeadForeColor" value="0">
																			<param name="Columns(5).HasHeadBackColor" value="0">
																			<param name="Columns(5).HasForeColor" value="0">
																			<param name="Columns(5).HasBackColor" value="0">
																			<param name="Columns(5).HeadForeColor" value="0">
																			<param name="Columns(5).HeadBackColor" value="0">
																			<param name="Columns(5).ForeColor" value="0">
																			<param name="Columns(5).BackColor" value="0">
																			<param name="Columns(5).HeadStyleSet" value="">
																			<param name="Columns(5).StyleSet" value="">
																			<param name="Columns(5).Nullable" value="1">
																			<param name="Columns(5).Mask" value="">
																			<param name="Columns(5).PromptInclude" value="0">
																			<param name="Columns(5).ClipMode" value="0">
																			<param name="Columns(5).PromptChar" value="95">

																			<param name="Columns(6).Width" value="0">
																			<param name="Columns(6).Visible" value="0">
																			<param name="Columns(6).Columns.Count" value="1">
																			<param name="Columns(6).Caption" value="hidden">
																			<param name="Columns(6).Name" value="hidden">
																			<param name="Columns(6).Alignment" value="0">
																			<param name="Columns(6).CaptionAlignment" value="3">
																			<param name="Columns(6).Bound" value="0">
																			<param name="Columns(6).AllowSizing" value="1">
																			<param name="Columns(6).DataField" value="Column 6">
																			<param name="Columns(6).DataType" value="8">
																			<param name="Columns(6).Level" value="0">
																			<param name="Columns(6).NumberFormat" value="">
																			<param name="Columns(6).Case" value="0">
																			<param name="Columns(6).FieldLen" value="4096">
																			<param name="Columns(6).VertScrollBar" value="0">
																			<param name="Columns(6).Locked" value="0">
																			<param name="Columns(6).Style" value="0">
																			<param name="Columns(6).ButtonsAlways" value="0">
																			<param name="Columns(6).RowCount" value="0">
																			<param name="Columns(6).ColCount" value="1">
																			<param name="Columns(6).HasHeadForeColor" value="0">
																			<param name="Columns(6).HasHeadBackColor" value="0">
																			<param name="Columns(6).HasForeColor" value="0">
																			<param name="Columns(6).HasBackColor" value="0">
																			<param name="Columns(6).HeadForeColor" value="0">
																			<param name="Columns(6).HeadBackColor" value="0">
																			<param name="Columns(6).ForeColor" value="0">
																			<param name="Columns(6).BackColor" value="0">
																			<param name="Columns(6).HeadStyleSet" value="">
																			<param name="Columns(6).StyleSet" value="">
																			<param name="Columns(6).Nullable" value="1">
																			<param name="Columns(6).Mask" value="">
																			<param name="Columns(6).PromptInclude" value="0">
																			<param name="Columns(6).ClipMode" value="0">
																			<param name="Columns(6).PromptChar" value="95">

																			<param name="Columns(7).Width" value="0">
																			<param name="Columns(7).Visible" value="0">
																			<param name="Columns(7).Columns.Count" value="1">
																			<param name="Columns(7).Caption" value="numeric">
																			<param name="Columns(7).Name" value="numeric">
																			<param name="Columns(7).Alignment" value="0">
																			<param name="Columns(7).CaptionAlignment" value="3">
																			<param name="Columns(7).Bound" value="0">
																			<param name="Columns(7).AllowSizing" value="1">
																			<param name="Columns(7).DataField" value="Column 7">
																			<param name="Columns(7).DataType" value="8">
																			<param name="Columns(7).Level" value="0">
																			<param name="Columns(7).NumberFormat" value="">
																			<param name="Columns(7).Case" value="0">
																			<param name="Columns(7).FieldLen" value="4096">
																			<param name="Columns(7).VertScrollBar" value="0">
																			<param name="Columns(7).Locked" value="0">
																			<param name="Columns(7).Style" value="0">
																			<param name="Columns(7).ButtonsAlways" value="0">
																			<param name="Columns(7).RowCount" value="0">
																			<param name="Columns(7).ColCount" value="1">
																			<param name="Columns(7).HasHeadForeColor" value="0">
																			<param name="Columns(7).HasHeadBackColor" value="0">
																			<param name="Columns(7).HasForeColor" value="0">
																			<param name="Columns(7).HasBackColor" value="0">
																			<param name="Columns(7).HeadForeColor" value="0">
																			<param name="Columns(7).HeadBackColor" value="0">
																			<param name="Columns(7).ForeColor" value="0">
																			<param name="Columns(7).BackColor" value="0">
																			<param name="Columns(7).HeadStyleSet" value="">
																			<param name="Columns(7).StyleSet" value="">
																			<param name="Columns(7).Nullable" value="1">
																			<param name="Columns(7).Mask" value="">
																			<param name="Columns(7).PromptInclude" value="0">
																			<param name="Columns(7).ClipMode" value="0">
																			<param name="Columns(7).PromptChar" value="95">

																			<param name="Columns(8).Width" value="0">
																			<param name="Columns(8).Visible" value="0">
																			<param name="Columns(8).Columns.Count" value="1">
																			<param name="Columns(8).Caption" value="heading">
																			<param name="Columns(8).Name" value="heading">
																			<param name="Columns(8).Alignment" value="0">
																			<param name="Columns(8).CaptionAlignment" value="3">
																			<param name="Columns(8).Bound" value="0">
																			<param name="Columns(8).AllowSizing" value="1">
																			<param name="Columns(8).DataField" value="Column 8">
																			<param name="Columns(8).DataType" value="8">
																			<param name="Columns(8).Level" value="0">
																			<param name="Columns(8).NumberFormat" value="">
																			<param name="Columns(8).Case" value="0">
																			<param name="Columns(8).FieldLen" value="4096">
																			<param name="Columns(8).VertScrollBar" value="0">
																			<param name="Columns(8).Locked" value="0">
																			<param name="Columns(8).Style" value="0">
																			<param name="Columns(8).ButtonsAlways" value="0">
																			<param name="Columns(8).RowCount" value="0">
																			<param name="Columns(8).ColCount" value="1">
																			<param name="Columns(8).HasHeadForeColor" value="0">
																			<param name="Columns(8).HasHeadBackColor" value="0">
																			<param name="Columns(8).HasForeColor" value="0">
																			<param name="Columns(8).HasBackColor" value="0">
																			<param name="Columns(8).HeadForeColor" value="0">
																			<param name="Columns(8).HeadBackColor" value="0">
																			<param name="Columns(8).ForeColor" value="0">
																			<param name="Columns(8).BackColor" value="0">
																			<param name="Columns(8).HeadStyleSet" value="">
																			<param name="Columns(8).StyleSet" value="">
																			<param name="Columns(8).Nullable" value="1">
																			<param name="Columns(8).Mask" value="">
																			<param name="Columns(8).PromptInclude" value="0">
																			<param name="Columns(8).ClipMode" value="0">
																			<param name="Columns(8).PromptChar" value="95">

																			<param name="Columns(9).Width" value="0">
																			<param name="Columns(9).Visible" value="0">
																			<param name="Columns(9).Columns.Count" value="1">
																			<param name="Columns(9).Caption" value="average">
																			<param name="Columns(9).Name" value="average">
																			<param name="Columns(9).Alignment" value="0">
																			<param name="Columns(9).CaptionAlignment" value="3">
																			<param name="Columns(9).Bound" value="0">
																			<param name="Columns(9).AllowSizing" value="1">
																			<param name="Columns(9).DataField" value="Column 9">
																			<param name="Columns(9).DataType" value="8">
																			<param name="Columns(9).Level" value="0">
																			<param name="Columns(9).NumberFormat" value="">
																			<param name="Columns(9).Case" value="0">
																			<param name="Columns(9).FieldLen" value="4096">
																			<param name="Columns(9).VertScrollBar" value="0">
																			<param name="Columns(9).Locked" value="0">
																			<param name="Columns(9).Style" value="0">
																			<param name="Columns(9).ButtonsAlways" value="0">
																			<param name="Columns(9).RowCount" value="0">
																			<param name="Columns(9).ColCount" value="1">
																			<param name="Columns(9).HasHeadForeColor" value="0">
																			<param name="Columns(9).HasHeadBackColor" value="0">
																			<param name="Columns(9).HasForeColor" value="0">
																			<param name="Columns(9).HasBackColor" value="0">
																			<param name="Columns(9).HeadForeColor" value="0">
																			<param name="Columns(9).HeadBackColor" value="0">
																			<param name="Columns(9).ForeColor" value="0">
																			<param name="Columns(9).BackColor" value="0">
																			<param name="Columns(9).HeadStyleSet" value="">
																			<param name="Columns(9).StyleSet" value="">
																			<param name="Columns(9).Nullable" value="1">
																			<param name="Columns(9).Mask" value="">
																			<param name="Columns(9).PromptInclude" value="0">
																			<param name="Columns(9).ClipMode" value="0">
																			<param name="Columns(9).PromptChar" value="95">

																			<param name="Columns(10).Width" value="0">
																			<param name="Columns(10).Visible" value="0">
																			<param name="Columns(10).Columns.Count" value="1">
																			<param name="Columns(10).Caption" value="count">
																			<param name="Columns(10).Name" value="count">
																			<param name="Columns(10).Alignment" value="0">
																			<param name="Columns(10).CaptionAlignment" value="3">
																			<param name="Columns(10).Bound" value="0">
																			<param name="Columns(10).AllowSizing" value="1">
																			<param name="Columns(10).DataField" value="Column 10">
																			<param name="Columns(10).DataType" value="8">
																			<param name="Columns(10).Level" value="0">
																			<param name="Columns(10).NumberFormat" value="">
																			<param name="Columns(10).Case" value="0">
																			<param name="Columns(10).FieldLen" value="4096">
																			<param name="Columns(10).VertScrollBar" value="0">
																			<param name="Columns(10).Locked" value="0">
																			<param name="Columns(10).Style" value="0">
																			<param name="Columns(10).ButtonsAlways" value="0">
																			<param name="Columns(10).RowCount" value="0">
																			<param name="Columns(10).ColCount" value="1">
																			<param name="Columns(10).HasHeadForeColor" value="0">
																			<param name="Columns(10).HasHeadBackColor" value="0">
																			<param name="Columns(10).HasForeColor" value="0">
																			<param name="Columns(10).HasBackColor" value="0">
																			<param name="Columns(10).HeadForeColor" value="0">
																			<param name="Columns(10).HeadBackColor" value="0">
																			<param name="Columns(10).ForeColor" value="0">
																			<param name="Columns(10).BackColor" value="0">
																			<param name="Columns(10).HeadStyleSet" value="">
																			<param name="Columns(10).StyleSet" value="">
																			<param name="Columns(10).Nullable" value="1">
																			<param name="Columns(10).Mask" value="">
																			<param name="Columns(10).PromptInclude" value="0">
																			<param name="Columns(10).ClipMode" value="0">
																			<param name="Columns(10).PromptChar" value="95">

																			<param name="Columns(11).Width" value="0">
																			<param name="Columns(11).Visible" value="0">
																			<param name="Columns(11).Columns.Count" value="1">
																			<param name="Columns(11).Caption" value="total">
																			<param name="Columns(11).Name" value="total">
																			<param name="Columns(11).Alignment" value="0">
																			<param name="Columns(11).CaptionAlignment" value="3">
																			<param name="Columns(11).Bound" value="0">
																			<param name="Columns(11).AllowSizing" value="1">
																			<param name="Columns(11).DataField" value="Column 11">
																			<param name="Columns(11).DataType" value="8">
																			<param name="Columns(11).Level" value="0">
																			<param name="Columns(11).NumberFormat" value="">
																			<param name="Columns(11).Case" value="0">
																			<param name="Columns(11).FieldLen" value="4096">
																			<param name="Columns(11).VertScrollBar" value="0">
																			<param name="Columns(11).Locked" value="0">
																			<param name="Columns(11).Style" value="0">
																			<param name="Columns(11).ButtonsAlways" value="0">
																			<param name="Columns(11).RowCount" value="0">
																			<param name="Columns(11).ColCount" value="1">
																			<param name="Columns(11).HasHeadForeColor" value="0">
																			<param name="Columns(11).HasHeadBackColor" value="0">
																			<param name="Columns(11).HasForeColor" value="0">
																			<param name="Columns(11).HasBackColor" value="0">
																			<param name="Columns(11).HeadForeColor" value="0">
																			<param name="Columns(11).HeadBackColor" value="0">
																			<param name="Columns(11).ForeColor" value="0">
																			<param name="Columns(11).BackColor" value="0">
																			<param name="Columns(11).HeadStyleSet" value="">
																			<param name="Columns(11).StyleSet" value="">
																			<param name="Columns(11).Nullable" value="1">
																			<param name="Columns(11).Mask" value="">
																			<param name="Columns(11).PromptInclude" value="0">
																			<param name="Columns(11).ClipMode" value="0">
																			<param name="Columns(11).PromptChar" value="95">

																			<param name="Columns(12).Width" value="0">
																			<param name="Columns(12).Visible" value="0">
																			<param name="Columns(12).Columns.Count" value="1">
																			<param name="Columns(12).Caption" value="hidden">
																			<param name="Columns(12).Name" value="hidden">
																			<param name="Columns(12).Alignment" value="0">
																			<param name="Columns(12).CaptionAlignment" value="3">
																			<param name="Columns(12).Bound" value="0">
																			<param name="Columns(12).AllowSizing" value="1">
																			<param name="Columns(12).DataField" value="Column 12">
																			<param name="Columns(12).DataType" value="8">
																			<param name="Columns(12).Level" value="0">
																			<param name="Columns(12).NumberFormat" value="">
																			<param name="Columns(12).Case" value="0">
																			<param name="Columns(12).FieldLen" value="4096">
																			<param name="Columns(12).VertScrollBar" value="0">
																			<param name="Columns(12).Locked" value="0">
																			<param name="Columns(12).Style" value="0">
																			<param name="Columns(12).ButtonsAlways" value="0">
																			<param name="Columns(12).RowCount" value="0">
																			<param name="Columns(12).ColCount" value="1">
																			<param name="Columns(12).HasHeadForeColor" value="0">
																			<param name="Columns(12).HasHeadBackColor" value="0">
																			<param name="Columns(12).HasForeColor" value="0">
																			<param name="Columns(12).HasBackColor" value="0">
																			<param name="Columns(12).HeadForeColor" value="0">
																			<param name="Columns(12).HeadBackColor" value="0">
																			<param name="Columns(12).ForeColor" value="0">
																			<param name="Columns(12).BackColor" value="0">
																			<param name="Columns(12).HeadStyleSet" value="">
																			<param name="Columns(12).StyleSet" value="">
																			<param name="Columns(12).Nullable" value="1">
																			<param name="Columns(12).Mask" value="">
																			<param name="Columns(12).PromptInclude" value="0">
																			<param name="Columns(12).ClipMode" value="0">
																			<param name="Columns(12).PromptChar" value="95">

																			<param name="Columns(13).Width" value="0">
																			<param name="Columns(13).Visible" value="0">
																			<param name="Columns(13).Columns.Count" value="1">
																			<param name="Columns(13).Caption" value="GroupWithNext">
																			<param name="Columns(13).Name" value="GroupWithNext">
																			<param name="Columns(13).Alignment" value="0">
																			<param name="Columns(13).CaptionAlignment" value="3">
																			<param name="Columns(13).Bound" value="0">
																			<param name="Columns(13).AllowSizing" value="1">
																			<param name="Columns(13).DataField" value="Column 13">
																			<param name="Columns(13).DataType" value="8">
																			<param name="Columns(13).Level" value="0">
																			<param name="Columns(13).NumberFormat" value="">
																			<param name="Columns(13).Case" value="0">
																			<param name="Columns(13).FieldLen" value="4096">
																			<param name="Columns(13).VertScrollBar" value="0">
																			<param name="Columns(13).Locked" value="0">
																			<param name="Columns(13).Style" value="0">
																			<param name="Columns(13).ButtonsAlways" value="0">
																			<param name="Columns(13).RowCount" value="0">
																			<param name="Columns(13).ColCount" value="1">
																			<param name="Columns(13).HasHeadForeColor" value="0">
																			<param name="Columns(13).HasHeadBackColor" value="0">
																			<param name="Columns(13).HasForeColor" value="0">
																			<param name="Columns(13).HasBackColor" value="0">
																			<param name="Columns(13).HeadForeColor" value="0">
																			<param name="Columns(13).HeadBackColor" value="0">
																			<param name="Columns(13).ForeColor" value="0">
																			<param name="Columns(13).BackColor" value="0">
																			<param name="Columns(13).HeadStyleSet" value="">
																			<param name="Columns(13).StyleSet" value="">
																			<param name="Columns(13).Nullable" value="1">
																			<param name="Columns(13).Mask" value="">
																			<param name="Columns(13).PromptInclude" value="0">
																			<param name="Columns(13).ClipMode" value="0">
																			<param name="Columns(13).PromptChar" value="95">

																			<param name="UseDefaults" value="-1">
																			<param name="TabNavigation" value="1">
																			<param name="BatchUpdate" value="0">
																			<param name="_ExtentX" value="4577">
																			<param name="_ExtentY" value="8202">
																			<param name="_StockProps" value="79">
																			<param name="Caption" value="">
																			<param name="ForeColor" value="0">
																			<param name="BackColor" value="16777215">
																			<param name="Enabled" value="-1">
																			<param name="DataMember" value="">
																		</object>
																	</td>
																</tr>
															</table>
														</td>
														<td width="5"></td>
													</tr>

													<tr height="5">
														<td height="5" colspan="6"></td>
													</tr>

													<tr>
														<td width="5"></td>
														<td rowspan="4" width="40%" height="100%">
															<object classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"
																codebase="cabs/COAInt_Grid.cab#version=3,1,3,6"
																id="ssOleDBGridAvailableColumns"
																name="ssOleDBGridAvailableColumns"
																style="HEIGHT: 100%; LEFT: 0px; TOP: 0px; WIDTH: 100%">
																<param name="ScrollBars" value="2">
																<param name="_Version" value="196617">
																<param name="DataMode" value="2">
																<param name="Cols" value="0">
																<param name="Rows" value="0">
																<param name="BorderStyle" value="1">
																<param name="RecordSelectors" value="0">
																<param name="GroupHeaders" value="0">
																<param name="ColumnHeaders" value="-1">
																<param name="GroupHeadLines" value="1">
																<param name="HeadLines" value="1">
																<param name="FieldDelimiter" value="(None)">
																<param name="FieldSeparator" value="(Tab)">
																<param name="Row.Count" value="0">
																<param name="Col.Count" value="8">
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
																<param name="AllowColumnSizing" value="-1">
																<param name="AllowGroupMoving" value="0">
																<param name="AllowColumnMoving" value="0">
																<param name="AllowGroupSwapping" value="0">
																<param name="AllowColumnSwapping" value="0">
																<param name="AllowGroupShrinking" value="0">
																<param name="AllowColumnShrinking" value="0">
																<param name="AllowDragDrop" value="0">
																<param name="UseExactRowCount" value="-1">
																<param name="SelectTypeCol" value="0">
																<param name="SelectTypeRow" value="3">
																<param name="SelectByCell" value="-1">
																<param name="BalloonHelp" value="0">
																<param name="RowNavigation" value="1">
																<param name="CellNavigation" value="0">
																<param name="MaxSelectedRows" value="0">
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
																<param name="Columns.Count" value="8">

																<param name="Columns(0).Width" value="0">
																<param name="Columns(0).Visible" value="0">
																<param name="Columns(0).Columns.Count" value="1">
																<param name="Columns(0).Caption" value="type">
																<param name="Columns(0).Name" value="type">
																<param name="Columns(0).Alignment" value="0">
																<param name="Columns(0).CaptionAlignment" value="3">
																<param name="Columns(0).Bound" value="0">
																<param name="Columns(0).AllowSizing" value="1">
																<param name="Columns(0).DataField" value="Column 0">
																<param name="Columns(0).DataType" value="8">
																<param name="Columns(0).Level" value="0">
																<param name="Columns(0).NumberFormat" value="">
																<param name="Columns(0).Case" value="0">
																<param name="Columns(0).FieldLen" value="4096">
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
																<param name="Columns(1).Caption" value="tableID">
																<param name="Columns(1).Name" value="tableID">
																<param name="Columns(1).Alignment" value="0">
																<param name="Columns(1).CaptionAlignment" value="3">
																<param name="Columns(1).Bound" value="0">
																<param name="Columns(1).AllowSizing" value="1">
																<param name="Columns(1).DataField" value="Column 1">
																<param name="Columns(1).DataType" value="8">
																<param name="Columns(1).Level" value="0">
																<param name="Columns(1).NumberFormat" value="">
																<param name="Columns(1).Case" value="0">
																<param name="Columns(1).FieldLen" value="4096">
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

																<param name="Columns(2).Width" value="0">
																<param name="Columns(2).Visible" value="0">
																<param name="Columns(2).Columns.Count" value="1">
																<param name="Columns(2).Caption" value="columnID">
																<param name="Columns(2).Name" value="columnID">
																<param name="Columns(2).Alignment" value="0">
																<param name="Columns(2).CaptionAlignment" value="3">
																<param name="Columns(2).Bound" value="0">
																<param name="Columns(2).AllowSizing" value="1">
																<param name="Columns(2).DataField" value="Column 2">
																<param name="Columns(2).DataType" value="8">
																<param name="Columns(2).Level" value="0">
																<param name="Columns(2).NumberFormat" value="">
																<param name="Columns(2).Case" value="0">
																<param name="Columns(2).FieldLen" value="4096">
																<param name="Columns(2).VertScrollBar" value="0">
																<param name="Columns(2).Locked" value="0">
																<param name="Columns(2).Style" value="0">
																<param name="Columns(2).ButtonsAlways" value="0">
																<param name="Columns(2).RowCount" value="0">
																<param name="Columns(2).ColCount" value="1">
																<param name="Columns(2).HasHeadForeColor" value="0">
																<param name="Columns(2).HasHeadBackColor" value="0">
																<param name="Columns(2).HasForeColor" value="0">
																<param name="Columns(2).HasBackColor" value="0">
																<param name="Columns(2).HeadForeColor" value="0">
																<param name="Columns(2).HeadBackColor" value="0">
																<param name="Columns(2).ForeColor" value="0">
																<param name="Columns(2).BackColor" value="0">
																<param name="Columns(2).HeadStyleSet" value="">
																<param name="Columns(2).StyleSet" value="">
																<param name="Columns(2).Nullable" value="1">
																<param name="Columns(2).Mask" value="">
																<param name="Columns(2).PromptInclude" value="0">
																<param name="Columns(2).ClipMode" value="0">
																<param name="Columns(2).PromptChar" value="95">

																<param name="Columns(3).Width" value="100000">
																<param name="Columns(3).Visible" value="-1">
																<param name="Columns(3).Columns.Count" value="1">
																<param name="Columns(3).Caption" value="Columns / Calculations Available">
																<param name="Columns(3).Name" value="display">
																<param name="Columns(3).Alignment" value="0">
																<param name="Columns(3).CaptionAlignment" value="3">
																<param name="Columns(3).Bound" value="0">
																<param name="Columns(3).AllowSizing" value="1">
																<param name="Columns(3).DataField" value="Column 3">
																<param name="Columns(3).DataType" value="8">
																<param name="Columns(3).Level" value="0">
																<param name="Columns(3).NumberFormat" value="">
																<param name="Columns(3).Case" value="0">
																<param name="Columns(3).FieldLen" value="4096">
																<param name="Columns(3).VertScrollBar" value="0">
																<param name="Columns(3).Locked" value="0">
																<param name="Columns(3).Style" value="0">
																<param name="Columns(3).ButtonsAlways" value="0">
																<param name="Columns(3).RowCount" value="0">
																<param name="Columns(3).ColCount" value="1">
																<param name="Columns(3).HasHeadForeColor" value="0">
																<param name="Columns(3).HasHeadBackColor" value="0">
																<param name="Columns(3).HasForeColor" value="0">
																<param name="Columns(3).HasBackColor" value="0">
																<param name="Columns(3).HeadForeColor" value="0">
																<param name="Columns(3).HeadBackColor" value="0">
																<param name="Columns(3).ForeColor" value="0">
																<param name="Columns(3).BackColor" value="0">
																<param name="Columns(3).HeadStyleSet" value="">
																<param name="Columns(3).StyleSet" value="">
																<param name="Columns(3).Nullable" value="1">
																<param name="Columns(3).Mask" value="">
																<param name="Columns(3).PromptInclude" value="0">
																<param name="Columns(3).ClipMode" value="0">
																<param name="Columns(3).PromptChar" value="95">

																<param name="Columns(4).Width" value="0">
																<param name="Columns(4).Visible" value="0">
																<param name="Columns(4).Columns.Count" value="1">
																<param name="Columns(4).Caption" value="size">
																<param name="Columns(4).Name" value="size">
																<param name="Columns(4).Alignment" value="0">
																<param name="Columns(4).CaptionAlignment" value="3">
																<param name="Columns(4).Bound" value="0">
																<param name="Columns(4).AllowSizing" value="1">
																<param name="Columns(4).DataField" value="Column 4">
																<param name="Columns(4).DataType" value="8">
																<param name="Columns(4).Level" value="0">
																<param name="Columns(4).NumberFormat" value="">
																<param name="Columns(4).Case" value="0">
																<param name="Columns(4).FieldLen" value="4096">
																<param name="Columns(4).VertScrollBar" value="0">
																<param name="Columns(4).Locked" value="0">
																<param name="Columns(4).Style" value="0">
																<param name="Columns(4).ButtonsAlways" value="0">
																<param name="Columns(4).RowCount" value="0">
																<param name="Columns(4).ColCount" value="1">
																<param name="Columns(4).HasHeadForeColor" value="0">
																<param name="Columns(4).HasHeadBackColor" value="0">
																<param name="Columns(4).HasForeColor" value="0">
																<param name="Columns(4).HasBackColor" value="0">
																<param name="Columns(4).HeadForeColor" value="0">
																<param name="Columns(4).HeadBackColor" value="0">
																<param name="Columns(4).ForeColor" value="0">
																<param name="Columns(4).BackColor" value="0">
																<param name="Columns(4).HeadStyleSet" value="">
																<param name="Columns(4).StyleSet" value="">
																<param name="Columns(4).Nullable" value="1">
																<param name="Columns(4).Mask" value="">
																<param name="Columns(4).PromptInclude" value="0">
																<param name="Columns(4).ClipMode" value="0">
																<param name="Columns(4).PromptChar" value="95">

																<param name="Columns(5).Width" value="0">
																<param name="Columns(5).Visible" value="0">
																<param name="Columns(5).Columns.Count" value="1">
																<param name="Columns(5).Caption" value="decimals">
																<param name="Columns(5).Name" value="decimals">
																<param name="Columns(5).Alignment" value="0">
																<param name="Columns(5).CaptionAlignment" value="3">
																<param name="Columns(5).Bound" value="0">
																<param name="Columns(5).AllowSizing" value="1">
																<param name="Columns(5).DataField" value="Column 5">
																<param name="Columns(5).DataType" value="8">
																<param name="Columns(5).Level" value="0">
																<param name="Columns(5).NumberFormat" value="">
																<param name="Columns(5).Case" value="0">
																<param name="Columns(5).FieldLen" value="4096">
																<param name="Columns(5).VertScrollBar" value="0">
																<param name="Columns(5).Locked" value="0">
																<param name="Columns(5).Style" value="0">
																<param name="Columns(5).ButtonsAlways" value="0">
																<param name="Columns(5).RowCount" value="0">
																<param name="Columns(5).ColCount" value="1">
																<param name="Columns(5).HasHeadForeColor" value="0">
																<param name="Columns(5).HasHeadBackColor" value="0">
																<param name="Columns(5).HasForeColor" value="0">
																<param name="Columns(5).HasBackColor" value="0">
																<param name="Columns(5).HeadForeColor" value="0">
																<param name="Columns(5).HeadBackColor" value="0">
																<param name="Columns(5).ForeColor" value="0">
																<param name="Columns(5).BackColor" value="0">
																<param name="Columns(5).HeadStyleSet" value="">
																<param name="Columns(5).StyleSet" value="">
																<param name="Columns(5).Nullable" value="1">
																<param name="Columns(5).Mask" value="">
																<param name="Columns(5).PromptInclude" value="0">
																<param name="Columns(5).ClipMode" value="0">
																<param name="Columns(5).PromptChar" value="95">

																<param name="Columns(6).Width" value="0">
																<param name="Columns(6).Visible" value="0">
																<param name="Columns(6).Columns.Count" value="1">
																<param name="Columns(6).Caption" value="hidden">
																<param name="Columns(6).Name" value="hidden">
																<param name="Columns(6).Alignment" value="0">
																<param name="Columns(6).CaptionAlignment" value="3">
																<param name="Columns(6).Bound" value="0">
																<param name="Columns(6).AllowSizing" value="1">
																<param name="Columns(6).DataField" value="Column 6">
																<param name="Columns(6).DataType" value="8">
																<param name="Columns(6).Level" value="0">
																<param name="Columns(6).NumberFormat" value="">
																<param name="Columns(6).Case" value="0">
																<param name="Columns(6).FieldLen" value="4096">
																<param name="Columns(6).VertScrollBar" value="0">
																<param name="Columns(6).Locked" value="0">
																<param name="Columns(6).Style" value="0">
																<param name="Columns(6).ButtonsAlways" value="0">
																<param name="Columns(6).RowCount" value="0">
																<param name="Columns(6).ColCount" value="1">
																<param name="Columns(6).HasHeadForeColor" value="0">
																<param name="Columns(6).HasHeadBackColor" value="0">
																<param name="Columns(6).HasForeColor" value="0">
																<param name="Columns(6).HasBackColor" value="0">
																<param name="Columns(6).HeadForeColor" value="0">
																<param name="Columns(6).HeadBackColor" value="0">
																<param name="Columns(6).ForeColor" value="0">
																<param name="Columns(6).BackColor" value="0">
																<param name="Columns(6).HeadStyleSet" value="">
																<param name="Columns(6).StyleSet" value="">
																<param name="Columns(6).Nullable" value="1">
																<param name="Columns(6).Mask" value="">
																<param name="Columns(6).PromptInclude" value="0">
																<param name="Columns(6).ClipMode" value="0">
																<param name="Columns(6).PromptChar" value="95">

																<param name="Columns(7).Width" value="0">
																<param name="Columns(7).Visible" value="0">
																<param name="Columns(7).Columns.Count" value="1">
																<param name="Columns(7).Caption" value="numeric">
																<param name="Columns(7).Name" value="numeric">
																<param name="Columns(7).Alignment" value="0">
																<param name="Columns(7).CaptionAlignment" value="3">
																<param name="Columns(7).Bound" value="0">
																<param name="Columns(7).AllowSizing" value="1">
																<param name="Columns(7).DataField" value="Column 7">
																<param name="Columns(7).DataType" value="8">
																<param name="Columns(7).Level" value="0">
																<param name="Columns(7).NumberFormat" value="">
																<param name="Columns(7).Case" value="0">
																<param name="Columns(7).FieldLen" value="4096">
																<param name="Columns(7).VertScrollBar" value="0">
																<param name="Columns(7).Locked" value="0">
																<param name="Columns(7).Style" value="0">
																<param name="Columns(7).ButtonsAlways" value="0">
																<param name="Columns(7).RowCount" value="0">
																<param name="Columns(7).ColCount" value="1">
																<param name="Columns(7).HasHeadForeColor" value="0">
																<param name="Columns(7).HasHeadBackColor" value="0">
																<param name="Columns(7).HasForeColor" value="0">
																<param name="Columns(7).HasBackColor" value="0">
																<param name="Columns(7).HeadForeColor" value="0">
																<param name="Columns(7).HeadBackColor" value="0">
																<param name="Columns(7).ForeColor" value="0">
																<param name="Columns(7).BackColor" value="0">
																<param name="Columns(7).HeadStyleSet" value="">
																<param name="Columns(7).StyleSet" value="">
																<param name="Columns(7).Nullable" value="1">
																<param name="Columns(7).Mask" value="">
																<param name="Columns(7).PromptInclude" value="0">
																<param name="Columns(7).ClipMode" value="0">
																<param name="Columns(7).PromptChar" value="95">

																<param name="UseDefaults" value="-1">
																<param name="TabNavigation" value="1">
																<param name="BatchUpdate" value="0">
																<param name="_ExtentX" value="4577">
																<param name="_ExtentY" value="8202">
																<param name="_StockProps" value="79">
																<param name="Caption" value="">
																<param name="ForeColor" value="0">
																<param name="BackColor" value="16777215">
																<param name="Enabled" value="-1">
																<param name="DataMember" value="">
															</object>
														</td>
														<td width="10" nowrap></td>
														<td height="5" valign="top" align="center">
															<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																<tr height="25">
																	<td>&nbsp</td>
																	<td width="100" nowrap align="center">
																		<input type="button" name="cmdColumnAdd" id="cmdColumnAdd" value="Add..." style="WIDTH: 100%; HEIGHT: 100%" class="btn"
																			onclick="columnSwap(true)"
																			onmouseover="try{button_onMouseOver(this);}catch(e){}"
																			onmouseout="try{button_onMouseOut(this);}catch(e){}"
																			onfocus="try{button_onFocus(this);}catch(e){}"
																			onblur="try{button_onBlur(this);}catch(e){}" />
																	</td>
																	<td>&nbsp;</td>
																</tr>
																<tr height="5">
																	<td></td>
																</tr>
																<tr height="25">
																	<td></td>
																	<td width="100" nowrap align="center">
																		<input type="button" name="cmdColumnAddAll" id="cmdColumnAddAll" value="Add All" style="WIDTH: 100%; HEIGHT: 100%" class="btn"
																			onclick="columnSwapAll(true)"
																			onmouseover="try{button_onMouseOver(this);}catch(e){}"
																			onmouseout="try{button_onMouseOut(this);}catch(e){}"
																			onfocus="try{button_onFocus(this);}catch(e){}"
																			onblur="try{button_onBlur(this);}catch(e){}" />
																	</td>
																	<td></td>
																</tr>
																<tr height="15">
																	<td></td>
																</tr>
																<tr height="25">
																	<td></td>
																	<td width="100" nowrap align="center">
																		<input type="button" name="cmdColumnRemove" id="cmdColumnRemove" value="Remove" style="WIDTH: 100%; HEIGHT: 100%" class="btn"
																			onclick="columnSwap(false)"
																			onmouseover="try{button_onMouseOver(this);}catch(e){}"
																			onmouseout="try{button_onMouseOut(this);}catch(e){}"
																			onfocus="try{button_onFocus(this);}catch(e){}"
																			onblur="try{button_onBlur(this);}catch(e){}" />
																	</td>
																	<td></td>
																</tr>
																<tr height="5">
																	<td></td>
																</tr>
																<tr height="25">
																	<td></td>
																	<td width="100" nowrap align="center">
																		<input type="button" name="cmdColumnRemoveAll" id="cmdColumnRemoveAll" value="Remove All" style="WIDTH: 100%; HEIGHT: 100%" class="btn"
																			onclick="columnSwapAll(false)"
																			onmouseover="try{button_onMouseOver(this);}catch(e){}"
																			onmouseout="try{button_onMouseOut(this);}catch(e){}"
																			onfocus="try{button_onFocus(this);}catch(e){}"
																			onblur="try{button_onBlur(this);}catch(e){}" />
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
																		<input id="txtSize" name="txtSize" maxlength="50" style="WIDTH: 100%" class="text"
																			onchange="validateColSize();"
																			onkeyup="validateColSize();">
																	</td>
																</tr>
																<tr>
																	<td width="125">Decimals :</td>
																	<td width="5"></td>
																	<td>
																		<input id="txtDecPlaces" name="txtDecPlaces" maxlength="50" style="WIDTH: 100%" class="text"
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
									<table width="100%" height="80%" class="outline" cellspacing="0" cellpadding="5">
										<tr valign="top">
											<td>
												<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
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
														<td rowspan="12">
															<object classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"
																codebase="cabs/COAInt_Grid.cab#version=3,1,3,6"
																id="ssOleDBGridSortOrder"
																name="ssOleDBGridSortOrder"
																style="BACKGROUND-COLOR: threedface; HEIGHT: 100%; VISIBILITY: visible; WIDTH: 100%">
																<param name="ScrollBars" value="2">
																<param name="_Version" value="196617">
																<param name="DataMode" value="2">
																<param name="Cols" value="0">
																<param name="Rows" value="0">
																<param name="BorderStyle" value="1">
																<param name="RecordSelectors" value="0">
																<param name="GroupHeaders" value="-1">
																<param name="ColumnHeaders" value="-1">
																<param name="GroupHeadLines" value="1">
																<param name="HeadLines" value="1">
																<param name="FieldDelimiter" value="(None)">
																<param name="FieldSeparator" value="(Tab)">
																<param name="Row.Count" value="0">
																<param name="Col.Count" value="1">
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
																<param name="AllowUpdate" value="-1">
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
																<param name="RowNavigation" value="2">
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
																<param name="Columns.Count" value="4">

																<param name="Columns(0).Width" value="3200">
																<param name="Columns(0).Visible" value="0">
																<param name="Columns(0).Columns.Count" value="1">
																<param name="Columns(0).Caption" value="id">
																<param name="Columns(0).Name" value="columnID">
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

																<param name="Columns(1).Width" value="5292">
																<param name="Columns(1).Visible" value="-1">
																<param name="Columns(1).Columns.Count" value="1">
																<param name="Columns(1).Caption" value="Column">
																<param name="Columns(1).Name" value="column">
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
																<param name="Columns(1).Locked" value="1">
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

																<param name="Columns(2).Width" value="3000">
																<param name="Columns(2).Visible" value="-1">
																<param name="Columns(2).Columns.Count" value="1">
																<param name="Columns(2).Caption" value="Sort Order">
																<param name="Columns(2).Name" value="order">
																<param name="Columns(2).Alignment" value="0">
																<param name="Columns(2).CaptionAlignment" value="3">
																<param name="Columns(2).Bound" value="0">
																<param name="Columns(2).AllowSizing" value="1">
																<param name="Columns(2).DataField" value="Column 2">
																<param name="Columns(2).DataType" value="8">
																<param name="Columns(2).Level" value="0">
																<param name="Columns(2).NumberFormat" value="">
																<param name="Columns(2).Case" value="0">
																<param name="Columns(2).FieldLen" value="256">
																<param name="Columns(2).VertScrollBar" value="0">
																<param name="Columns(2).Locked" value="-1">
																<param name="Columns(2).Style" value="3">
																<param name="Columns(2).ButtonsAlways" value="0">
																<param name="Columns(2).Row.Count" value="2">
																<param name="Columns(2).Col.Count" value="2">
																<param name="Columns(2).Row(0).Col(0)" value="Asc">
																<param name="Columns(2).Row(0).Col(1)" value="">
																<param name="Columns(2).Row(1).Col(0)" value="Desc">
																<param name="Columns(2).Row(1).Col(1)" value="">
																<param name="Columns(2).HasHeadForeColor" value="0">
																<param name="Columns(2).HasHeadBackColor" value="0">
																<param name="Columns(2).HasForeColor" value="0">
																<param name="Columns(2).HasBackColor" value="0">
																<param name="Columns(2).HeadForeColor" value="0">
																<param name="Columns(2).HeadBackColor" value="0">
																<param name="Columns(2).ForeColor" value="0">
																<param name="Columns(2).BackColor" value="0">
																<param name="Columns(2).HeadStyleSet" value="">
																<param name="Columns(2).StyleSet" value="">
																<param name="Columns(2).Nullable" value="1">
																<param name="Columns(2).Mask" value="">
																<param name="Columns(2).PromptInclude" value="0">
																<param name="Columns(2).ClipMode" value="0">
																<param name="Columns(2).PromptChar" value="95">

																<param name="Columns(3).Width" value="3200">
																<param name="Columns(3).Visible" value="0">
																<param name="Columns(3).Columns.Count" value="1">
																<param name="Columns(3).Caption" value="tableID">
																<param name="Columns(3).Name" value="tableID">
																<param name="Columns(3).Alignment" value="0">
																<param name="Columns(3).CaptionAlignment" value="3">
																<param name="Columns(3).Bound" value="0">
																<param name="Columns(3).AllowSizing" value="1">
																<param name="Columns(3).DataField" value="Column 7">
																<param name="Columns(3).DataType" value="8">
																<param name="Columns(3).Level" value="0">
																<param name="Columns(3).NumberFormat" value="">
																<param name="Columns(3).Case" value="0">
																<param name="Columns(3).FieldLen" value="256">
																<param name="Columns(3).VertScrollBar" value="0">
																<param name="Columns(3).Locked" value="0">
																<param name="Columns(3).Style" value="0">
																<param name="Columns(3).ButtonsAlways" value="0">
																<param name="Columns(3).RowCount" value="0">
																<param name="Columns(3).ColCount" value="1">
																<param name="Columns(3).HasHeadForeColor" value="0">
																<param name="Columns(3).HasHeadBackColor" value="0">
																<param name="Columns(3).HasForeColor" value="0">
																<param name="Columns(3).HasBackColor" value="0">
																<param name="Columns(3).HeadForeColor" value="0">
																<param name="Columns(3).HeadBackColor" value="0">
																<param name="Columns(3).ForeColor" value="0">
																<param name="Columns(3).BackColor" value="0">
																<param name="Columns(3).HeadStyleSet" value="">
																<param name="Columns(3).StyleSet" value="">
																<param name="Columns(3).Nullable" value="1">
																<param name="Columns(3).Mask" value="">
																<param name="Columns(3).PromptInclude" value="0">
																<param name="Columns(3).ClipMode" value="0">
																<param name="Columns(3).PromptChar" value="95">

																<param name="UseDefaults" value="-1">
																<param name="TabNavigation" value="1">
																<param name="BatchUpdate" value="0">
																<param name="_ExtentX" value="11298">
																<param name="_ExtentY" value="3969">
																<param name="_StockProps" value="79">
																<param name="Caption" value="">
																<param name="ForeColor" value="0">
																<param name="BackColor" value="-2147483633">
																<param name="Enabled" value="-1">
																<param name="DataMember" value="">
															</object>
														</td>

														<td width="10">&nbsp;</td>
														<td width="100">
															<input type="button" id="cmdSortAdd" name="cmdSortAdd" value="Add..." style="WIDTH: 100%" class="btn"
																onclick="sortAdd()"
																onmouseover="try{button_onMouseOver(this);}catch(e){}"
																onmouseout="try{button_onMouseOut(this);}catch(e){}"
																onfocus="try{button_onFocus(this);}catch(e){}"
																onblur="try{button_onBlur(this);}catch(e){}" />
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
															<input type="button" id="cmdSortEdit" name="cmdSortEdit" value="Edit..." style="WIDTH: 100%" class="btn"
																onclick="sortEdit()"
																onmouseover="try{button_onMouseOver(this);}catch(e){}"
																onmouseout="try{button_onMouseOut(this);}catch(e){}"
																onfocus="try{button_onFocus(this);}catch(e){}"
																onblur="try{button_onBlur(this);}catch(e){}" />
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
															<input type="button" id="cmdSortRemove" name="cmdSortRemove" value="Remove" style="WIDTH: 100%" class="btn"
																onclick="sortRemove()"
																onmouseover="try{button_onMouseOver(this);}catch(e){}"
																onmouseout="try{button_onMouseOut(this);}catch(e){}"
																onfocus="try{button_onFocus(this);}catch(e){}"
																onblur="try{button_onBlur(this);}catch(e){}" />
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
															<input type="button" id="cmdSortRemoveAll" name="cmdSortRemoveAll" value="Remove All" style="WIDTH: 100%" class="btn"
																onclick="sortRemoveAll()"
																onmouseover="try{button_onMouseOver(this);}catch(e){}"
																onmouseout="try{button_onMouseOut(this);}catch(e){}"
																onfocus="try{button_onFocus(this);}catch(e){}"
																onblur="try{button_onBlur(this);}catch(e){}" />
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
															<input type="button" id="cmdSortMoveUp" name="cmdSortMoveUp" value="Move Up" style="WIDTH: 100%" class="btn"
																onclick="sortMove(true)"
																onmouseover="try{button_onMouseOver(this);}catch(e){}"
																onmouseout="try{button_onMouseOut(this);}catch(e){}"
																onfocus="try{button_onFocus(this);}catch(e){}"
																onblur="try{button_onBlur(this);}catch(e){}" />
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
															<input type="button" id="cmdSortMoveDown" name="cmdSortMoveDown" value="Move Down" style="WIDTH: 100%" class="btn"
																onclick="sortMove(false)"
																onmouseover="try{button_onMouseOver(this);}catch(e){}"
																onmouseout="try{button_onMouseOut(this);}catch(e){}"
																onfocus="try{button_onFocus(this);}catch(e){}"
																onblur="try{button_onBlur(this);}catch(e){}" />
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
									<table width="100%" height="80%" class="outline" cellspacing="0" cellpadding="0">
										<tr valign="top">
											<td>
												<table width="100%" class="invisible" cellspacing="0" cellpadding="4">

													<tr height="5">
														<td colspan="9"></td>
													</tr>

													<tr>
														<td width="20">&nbsp;&nbsp;&nbsp;</td>
														<td nowrap width="100">Template :</td>
														<td width="20">&nbsp;&nbsp;&nbsp;</td>
														<td width="20">
															<input id="txtTemplate" name="txtTemplate" style="width: 400px" class="text textdisabled" disabled="disabled">
														</td>
														<td width="30">
															<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																<tr>
																	<td>
																		<input type="button" value="..." id="cmdTemplateSelect" name="cmdTemplateSelect" style="WIDTH: 25px" class="btn"
																			onclick="TemplateSelect()"
																			onmouseover="try{button_onMouseOver(this);}catch(e){}"
																			onmouseout="try{button_onMouseOut(this);}catch(e){}"
																			onfocus="try{button_onFocus(this);}catch(e){}"
																			onblur="try{button_onBlur(this);}catch(e){}" />
																	</td>
																	<td>
																		<input type="button" value="Clear" id="cmdTemplateClear" name="cmdTemplateClear" style="WIDTH: 50px" class="btn"
																			onclick="TemplateClear()"
																			onmouseover="try{button_onMouseOver(this);}catch(e){}"
																			onmouseout="try{button_onMouseOut(this);}catch(e){}"
																			onfocus="try{button_onFocus(this);}catch(e){}"
																			onblur="try{button_onBlur(this);}catch(e){}" />
																	</td>
																</tr>
															</table>
														</td>
														<td width="80">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
														<td nowrap>
															<input type="checkbox" id="chkPause" name="chkPause" tabindex="-1"
																onclick="changeTab4Control()"
																onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" /><label
																for="chkPause"
																class="checkbox"
																tabindex="0"
																onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																onblur="try{checkboxLabel_onBlur(this);}catch(e){}">Pause before mail merge</label>
														</td>
														<td width="100%">&nbsp;&nbsp;&nbsp;</td>
														<td width="20">&nbsp;&nbsp;&nbsp;</td>
													</tr>

													<tr>
														<td width="20">&nbsp;&nbsp;&nbsp;</td>
														<td nowrap width="100"></td>
														<td width="20">&nbsp;&nbsp;&nbsp;</td>
														<td width="420"></td>
														<td width="30"></td>
														<td width="80">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>

														<td nowrap>
															<input type="checkbox" id="chkSuppressBlanks" name="chkSuppressBlanks" tabindex="-1"
																onclick="changeTab4Control()"
																onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" /><label
																for="chkSuppressBlanks"
																class="checkbox"
																tabindex="0"
																onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																onblur="try{checkboxLabel_onBlur(this);}catch(e){}">Suppress blank lines</label>
														</td>

														<td colspan="2"></td>
													</tr>

													<tr height="5">
														<td></td>
														<td colspan="7">
															<hr>
														</td>
														<td></td>
													</tr>

												</table>

												<table width="100%" class="invisible" cellspacing="0" cellpadding="0" height="100%">

													<tr style="height: 100%">
														<td></td>
														<td colspan="6">

															<table style="width: 100%; height: 100%">
																<tr>
																	<td width="20">&nbsp;&nbsp;&nbsp;</td>
																	<td width="220px" valign="top">
																		<table style="vertical-align: text-top" class="outline" cellspacing="0" cellpadding="4" width="100%" height="200px">
																			<tr style="height: 20px">
																				<td colspan="4" align="left" style="vertical-align: text-top">Output Format :
																					<br>
																				</td>
																			</tr>

																			<tr style="height: 20px">
																				<td width="5" style="vertical-align: text-top">
																					<input checked id="optDestination0" name="optDestination" type="radio"
																						onclick="refreshDestination(); changeTab4Control(); "
																						onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																						onfocus="try{radio_onFocus(this);}catch(e){}"
																						onblur="try{radio_onBlur(this);}catch(e){}" />
																				</td>
																				<td width="5">&nbsp;</td>
																				<td width="130px" style="vertical-align: text-top">
																					<label tabindex="-1"
																						for="optDestination0"
																						class="radio"
																						onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
																						Word Document</label>
																				</td>
																				<td>&nbsp;</td>
																			</tr>
																			<tr style="height: 20px">
																				<td width="5">
																					<input id="optDestination1" name="optDestination" type="radio"
																						onclick="refreshDestination(); changeTab4Control(); "
																						onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																						onfocus="try{radio_onFocus(this);}catch(e){}"
																						onblur="try{radio_onBlur(this);}catch(e){}" />
																				</td>
																				<td width="5">&nbsp;</td>
																				<td width="130">
																					<label tabindex="-1"
																						for="optDestination1"
																						class="radio"
																						onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
																						Individual Emails</label>
																				</td>
																				<td width="5">&nbsp;</td>

																			</tr>
																			<%If iVersionOneEnabled = 0 Then%>
																			<tr style="height: 20px; visibility: hidden; display: none">
																				<%Else%>
																			<tr style="height: 20px; visibility: visible; display: block">
																				<%End If%>
																				<td width="5">
																					<input id="optDestination2" name="optDestination" type="radio"
																						onclick="refreshDestination(); changeTab4Control(); "
																						onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																						onfocus="try{radio_onFocus(this);}catch(e){}"
																						onblur="try{radio_onBlur(this);}catch(e){}" />
																				</td>
																				<td width="5">&nbsp;</td>
																				<td nowrap>
																					<label tabindex="-1"
																						for="optDestination2"
																						class="radio"
																						onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
																						Document Management</label>
																				</td>
																			</tr>
																			<tr></tr>
																		</table>


																		<td style="width: 5px"></td>

																	<td valign="top">
																		<table class="outline" cellspacing="0" cellpadding="4" style="width: 100%; height: 200px; vertical-align: top">
																			<tr style="height: 20px">
																				<td colspan="4" align="left">Output Destinations :
																					<br>
																				</td>
																			</tr>

																			<tr style="height: 20px" name="row1" id="row1">
																				<td width="150px" nowrap>Engine :</td>
																				<td width="5px"></td>
																				<td colspan="2">
																					<select id="cboDMEngine" name="cboDMEngine" style="WIDTH: 400px" class="combo"
																						onchange="changeTab4Control()">
																					</select>
																				</td>
																			</tr>


																			<tr style="height: 20px" name="row4" id="row4">
																				<td nowrap colspan="2">
																					<input type="checkbox" id="chkOutputScreen" name="chkOutputScreen" tabindex="-1"
																						onclick="changeTab4Control()"
																						onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																					<label
																						for="chkOutputScreen"
																						class="checkbox"
																						tabindex="0"
																						onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																						onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																						onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																						onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																						Display output on screen
																					</label>
																				</td>
																			</tr>

																			<tr style="height: 20px" name="row2" id="row2">
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
																				<td nowrap colspan="6"></td>
																			</tr>

																			<tr style="height: 20px" name="row5" id="row5">
																				<td nowrap>
																					<input type="checkbox" id="chkOutputPrinter" name="chkOutputPrinter" tabindex="-1"
																						onclick="chkOutputPrinter_Click(); changeTab4Control(); "
																						onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																					<label
																						for="chkOutputPrinter"
																						class="checkbox"
																						tabindex="0"
																						onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																						onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																						onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																						onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																						Send to printer
																					</label>
																				</td>
																				<td class="text">Printer location :</td>
																				<td colspan="2">
																					<select id="cboPrinterName" name="cboPrinterName" style="WIDTH: 400px" class="combo"
																						onchange="changeTab4Control()">
																					</select>
																				</td>
																			</tr>

																			<tr style="height: 20px" name="row6" id="row6">
																				<td nowrap>
																					<input type="checkbox" id="chkSave" name="chkSave" tabindex="-1"
																						onclick="chkSave_Click(); changeTab4Control(); "
																						onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																					<label
																						for="chkSave"
																						class="checkbox"
																						tabindex="0"
																						onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																						onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																						onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																						onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																						Save to file
																					</label>
																				</td>
																				<td class="text">File name :</td>
																				<td colspan="2">
																					<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
																						<tr>
																							<td width="20">
																								<input id="  " name="txtSaveFile" style="WIDTH: 325px" disabled="disabled" class="text textdisabled">
																							</td>
																							<td width="20">
																								<input type="button" value="..." id="cmdSaveFile" name="cmdSaveFile" style="WIDTH: 25px" class="btn"
																									onclick="saveFile()"
																									onmouseover="try{button_onMouseOver(this);}catch(e){}"
																									onmouseout="try{button_onMouseOut(this);}catch(e){}"
																									onfocus="try{button_onFocus(this);}catch(e){}"
																									onblur="try{button_onBlur(this);}catch(e){}" />
																							</td>
																							<td>
																								<input type="button" value="Clear" id="cmdClearFile" name="cmdClearFile" style="WIDTH: 50px" class="btn"
																									onclick="fileClear()"
																									onmouseover="try{button_onMouseOver(this);}catch(e){}"
																									onmouseout="try{button_onMouseOut(this);}catch(e){}"
																									onfocus="try{button_onFocus(this);}catch(e){}"
																									onblur="try{button_onBlur(this);}catch(e){}" />
																							</td>
																						</tr>
																					</table>
																				</td>
																			</tr>

																			<tr style="height: 20px" name="row7" id="row7">
																				<td width="150px" nowrap>Email Address :</td>
																				<td width="5px"></td>
																				<td>
																					<select id="cboEmail" name="cboEmail" style="WIDTH: 400px" class="combo"
																						onchange="changeTab4Control()">
																					</select>
																				</td>
																			</tr>

																			<tr style="height: 20px" name="row8" id="row8">
																				<td nowrap>Subject :</td>
																				<td width="5px"></td>
																				<td colspan="2">
																					<input id="txtSubject" name="txtSubject" style="WIDTH: 400px" maxlength="255" class="text"
																						onkeyup="changeTab4Control()">
																				</td>
																			</tr>

																			<tr style="height: 20px" name="row9" id="row9">
																				<td nowrap colspan="3">
																					<input type="checkbox" id="chkAttachment" name="chkAttachment" tabindex="-1"
																						onclick="chkAttachment_Click(); changeTab4Control(); "
																						onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																					<label
																						for="chkAttachment"
																						class="checkbox"
																						tabindex="0"
																						onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																						onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																						onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																						onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																						onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																						Send as attachment
																					</label>
																				</td>
																			</tr>

																			<tr style="height: 20px" name="row10" id="row10">
																				<td nowrap>Attach as :</td>
																				<td></td>
																				<td colspan="2">
																					<input id="txtAttachmentName" name="txtAttachmentName" maxlength="255" style="WIDTH: 400px" class="text"
																						onkeyup="changeTab4Control()" />
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
												onclick="okClick()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
										<td width="10"></td>
										<td width="80">
											<input type="button" id="cmdCancel" name="cmdCancel" value="Cancel" style="WIDTH: 100%" class="btn"
												onclick="cancelClick()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
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

		<input type='hidden' id="txtChildFilterID" name="txtChildFilterID">

		<input type="hidden" id="txtWordFormats" name="txtWordFormats" value="<%=Session("WordFormats")%>">
		<input type="hidden" id="txtWordFormatDefaultIndex" name="txtWordFormatDefaultIndex" value="<%=Session("WordFormatDefaultIndex")%>">
	</form>

	<form id="frmTables">
		<%
			Dim sErrorDescription = ""
	
			' Get the table records.
			Dim cmdTables = CreateObject("ADODB.Command")
			cmdTables.CommandText = "sp_ASRIntGetTablesInfo"
			cmdTables.CommandType = 4	' Stored Procedure
			cmdTables.ActiveConnection = Session("databaseConnection")

			Err.Number = 0
			Dim rstTablesInfo = cmdTables.Execute
			If (Err.Number <> 0) Then
				sErrorDescription = "The tables information could not be retrieved." & vbCrLf & FormatError(Err.Description)
			End If

			If Len(sErrorDescription) = 0 Then
				Do While Not rstTablesInfo.EOF
					Response.Write("<INPUT type='hidden' id=txtTableName_" & rstTablesInfo.fields("tableID").value & " name=txtTableName_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("tableName").value & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtTableType_" & rstTablesInfo.fields("tableID").value & " name=txtTableType_" & rstTablesInfo.fields("tableID").value & " value=" & rstTablesInfo.fields("tableType").value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtTableChildren_" & rstTablesInfo.fields("tableID").value & " name=txtTableChildren_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("childrenString").value & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtTableParents_" & rstTablesInfo.fields("tableID").value & " name=txtTableParents_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("parentsString").value & """>" & vbCrLf)

					rstTablesInfo.MoveNext()
				Loop

				' Release the ADO recordset object.
				rstTablesInfo.close()
				'rstTablesInfo = Nothing
			End If
	
			' Release the ADO command object.
			'cmdTables = Nothing
		%>
	</form>

	
	<form id="frmOriginalDefinition">
		<%
			Dim sErrMsg = ""
			Dim prmUtilID As Object
	
			If Session("action") <> "new" Then
				Dim cmdDefn = CreateObject("ADODB.Command")
				cmdDefn.CommandText = "sp_ASRIntGetMailMergeDefinition"
				cmdDefn.CommandType = 4	' Stored Procedure
				cmdDefn.ActiveConnection = Session("databaseConnection")
		
				prmUtilID = cmdDefn.CreateParameter("utilID", 3, 1)
				' 3=integer, 1=input
				cmdDefn.Parameters.Append(prmUtilID)
				prmUtilID.value = CleanNumeric(Session("utilid"))

				Dim prmUser = cmdDefn.CreateParameter("user", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
				cmdDefn.Parameters.Append(prmUser)
				prmUser.value = Session("username")

				Dim prmAction = cmdDefn.CreateParameter("action", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
				cmdDefn.Parameters.Append(prmAction)
				prmAction.value = Session("action")

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

				Dim prmSelection = cmdDefn.CreateParameter("selection", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmSelection)

				Dim prmPicklistID = cmdDefn.CreateParameter("picklistID", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmPicklistID)

				Dim prmPicklistName = cmdDefn.CreateParameter("picklistName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmPicklistName)

				Dim prmPicklistHidden = cmdDefn.CreateParameter("picklistHidden", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmPicklistHidden)

				Dim prmFilterID = cmdDefn.CreateParameter("filterID", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmFilterID)

				Dim prmFilterName = cmdDefn.CreateParameter("filterName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmFilterName)

				Dim prmFilterHidden = cmdDefn.CreateParameter("filterHidden", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmFilterHidden)

				Dim prmOutputFormat = cmdDefn.CreateParameter("outputFormat", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmOutputFormat)

				Dim prmOutputSave = cmdDefn.CreateParameter("outputSave", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmOutputSave)

				Dim prmOutputFileName = cmdDefn.CreateParameter("outputFileName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmOutputFileName)

				Dim prmEmailAddrID = cmdDefn.CreateParameter("EmailAddrID", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmEmailAddrID)

				Dim prmEmailSubject = cmdDefn.CreateParameter("EmailSubject", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmEmailSubject)

				Dim prmTemplateFileName = cmdDefn.CreateParameter("TemplateFileName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmTemplateFileName)

				Dim prmOutputScreen = cmdDefn.CreateParameter("outputScreen", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmOutputScreen)

				Dim prmEmailAsAttachment = cmdDefn.CreateParameter("EmailAsAttachment", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmEmailAsAttachment)

				Dim prmEmailAttachmentName = cmdDefn.CreateParameter("EmailAttachmentName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmEmailAttachmentName)

				Dim prmSuppressBlanks = cmdDefn.CreateParameter("SuppressBlanks", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmSuppressBlanks)

				Dim prmPauseBeforeMerge = cmdDefn.CreateParameter("PauseBeforeMerge", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmPauseBeforeMerge)

				Dim prmOutputPrinter = cmdDefn.CreateParameter("outputPrinter", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmOutputPrinter)
		
				Dim prmOutputPrinterName = cmdDefn.CreateParameter("outputPrinterName", 200, 2, 255) '200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmOutputPrinterName)

				Dim prmDocumentMapID = cmdDefn.CreateParameter("documentMapID", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmDocumentMapID)

				Dim prmManualDocManHeader = cmdDefn.CreateParameter("manualDocManHeader", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmManualDocManHeader)

				Dim prmTimestamp = cmdDefn.CreateParameter("timestamp", 3, 2)	' 3=integer, 2=output
				cmdDefn.Parameters.Append(prmTimestamp)

				Dim prmWarningMsg = cmdDefn.CreateParameter("warningMsg", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmWarningMsg)

				Err.Number = 0
				Dim rstDefinition = cmdDefn.Execute
				Dim iHiddenCalcCount = 0
		
				If (Err.Number <> 0) Then
					sErrMsg = CType(("'" & Session("utilname") & "' definition could not be read." & vbCrLf & FormatError(Err.Description)), String)
				Else
					If rstDefinition.state <> 0 Then
						' Read recordset values.
						Dim iCount = 0
						Do While Not rstDefinition.EOF
							iCount = iCount + 1
							If rstDefinition.fields("definitionType").value = "ORDER" Then
								Response.Write("<INPUT type='hidden' id=txtReportDefnOrder_" & iCount & " name=txtReportDefnOrder_" & iCount & " value=""" & rstDefinition.fields("definitionString").value & """>" & vbCrLf)
							Else
								Response.Write("<INPUT type='hidden' id=txtReportDefnColumn_" & iCount & " name=txtReportDefnColumn_" & iCount & " value=""" & Replace(rstDefinition.fields("definitionString").value, """", "&quot;") & """>" & vbCrLf)
	
								' Check if the report column is a hidden calc.
								If rstDefinition.fields("hidden").value = "Y" Then
									iHiddenCalcCount = iHiddenCalcCount + 1
								End If
							End If
							rstDefinition.MoveNext()
						Loop

						' Release the ADO recordset object.
						rstDefinition.close()
					End If
					'rstDefinition = Nothing
			
					' NB. IMPORTANT ADO NOTE.
					' When calling a stored procedure which returns a recordset AND has output parameters
					' you need to close the recordset and set it to nothing before using the output parameters. 
					If Len(cmdDefn.Parameters("errMsg").value) > 0 Then
						sErrMsg = CType(("'" & Session("utilname") & "' " & cmdDefn.Parameters("errMsg").value), String)
					End If

					Response.Write("<INPUT type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Replace(cmdDefn.Parameters("name").value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Owner name=txtDefn_Owner value=""" & Replace(cmdDefn.Parameters("owner").value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Description name=txtDefn_Description value=""" & Replace(cmdDefn.Parameters("description").value, """", "&quot;") & """>" & vbCrLf)
					'			Response.Write "<INPUT type='hidden' id=txtDefn_Access name=txtDefn_Access value=""" & cmdDefn.Parameters("access").value & """>" & vbcrlf
					Response.Write("<INPUT type='hidden' id=txtDefn_BaseTableID name=txtDefn_BaseTableID value=" & cmdDefn.Parameters("baseTableID").value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Selection name=txtDefn_Selection value=" & cmdDefn.Parameters("Selection").value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_PicklistID name=txtDefn_PicklistID value=" & cmdDefn.Parameters("picklistID").value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_PicklistName name=txtDefn_PicklistName value=""" & Replace(cmdDefn.Parameters("picklistName").value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_PicklistHidden name=txtDefn_PicklistHidden value=" & LCase(cmdDefn.Parameters("picklistHidden").value) & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_FilterID name=txtDefn_FilterID value=" & cmdDefn.Parameters("filterID").value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_FilterName name=txtDefn_FilterName value=""" & Replace(cmdDefn.Parameters("filterName").value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_FilterHidden name=txtDefn_FilterHidden value=" & LCase(cmdDefn.Parameters("filterHidden").value) & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_HiddenCalcCount name=txtDefn_HiddenCalcCount value=" & iHiddenCalcCount & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputFormat name=txtDefn_OutputFormat value=" & cmdDefn.Parameters("OutputFormat").value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputSave name=txtDefn_OutputSave value=" & LCase(cmdDefn.Parameters("OutputSave").value) & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputFileName name=txtDefn_OutputFileName value=""" & cmdDefn.Parameters("OutputFileName").value & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_EmailAddrID name=txtDefn_EmailAddrID value=" & cmdDefn.Parameters("EmailAddrID").value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_EmailSubject name=txtDefn_EmailSubject value=""" & Replace(cmdDefn.Parameters("EmailSubject").value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_TemplateFileName name=txtDefn_TemplateFileName value=""" & cmdDefn.Parameters("TemplateFileName").value & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputScreen name=txtDefn_OutputScreen value=" & LCase(cmdDefn.Parameters("OutputScreen").value) & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_EmailAsAttachment name=txtDefn_EmailAsAttachment value=" & Replace(LCase(cmdDefn.Parameters("EmailAsAttachment").value), """", "&quot;") & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_EmailAttachmentName name=txtDefn_EmailAttachmentName value=""" & cmdDefn.Parameters("EmailAttachmentName").value & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_SuppressBlanks name=txtDefn_SuppressBlanks value=" & LCase(cmdDefn.Parameters("SuppressBlanks").value) & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_PauseBeforeMerge name=txtDefn_PauseBeforeMerge value=" & LCase(cmdDefn.Parameters("PauseBeforeMerge").value) & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputPrinter name=txtDefn_OutputPrinter value=" & LCase(cmdDefn.Parameters("OutputPrinter").value) & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_OutputPrinterName name=txtDefn_OutputPrinterName value=""" & cmdDefn.Parameters("OutputPrinterName").value & """>" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_DocumentMapID name=txtDefn_DocumentMapID value=" & cmdDefn.Parameters("DocumentMapID").value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_ManualDocManHeader name=txtDefn_ManualDocManHeader value=" & LCase(cmdDefn.Parameters("ManualDocManHeader").value) & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Timestamp name=txtDefn_Timestamp value=" & cmdDefn.Parameters("timestamp").value & ">" & vbCrLf)
					Response.Write("<INPUT type='hidden' id=txtDefn_Warning name=txtDefn_Warning value=""" & Replace(cmdDefn.Parameters("warningMsg").value, """", "&quot;") & """>" & vbCrLf)
				End If
    
				Dim fDocManagement = False
				Dim lngDocumentMapID = 0
				If cmdDefn.Parameters("DocumentMapID").value > 0 Then
					fDocManagement = True
					lngDocumentMapID = CInt(cmdDefn.Parameters("DocumentMapID").value)
				End If

				' Release the ADO command object.
				cmdDefn = Nothing

				If Len(sErrMsg) > 0 Then
					Session("confirmtext") = sErrMsg
					Session("confirmtitle") = "OpenHR Intranet"
					Session("followpage") = "defsel"
					Session("reaction") = "MAILMERGE"
					Response.Clear()
					Response.Redirect("confirmok")
				End If
    
				If fDocManagement = True Then
					' Get the Document Type 'Name' (only the ID is stored in the table)
					Dim cmdDocManRecords = CreateObject("ADODB.Command")
					cmdDocManRecords.CommandText = "spASRIntGetDocumentManagementTypes"
					cmdDocManRecords.CommandType = 4 ' Stored Procedure
					cmdDocManRecords.ActiveConnection = Session("databaseConnection")
					Err.Number = 0
					Dim rstDocManRecords = cmdDocManRecords.Execute
	    
					Dim lngCount = 1
					Do While Not rstDocManRecords.EOF
						If CInt(rstDocManRecords.Fields(0).Value) = lngDocumentMapID Then
							Response.Write("<INPUT type='hidden' id=txtDefn_DocumentMapName name=txtDefn_DocumentMapName value=""" & Replace(CType(rstDocManRecords.Fields(1).Value, String), """", "&quot;") & """>" & vbCrLf)
						End If

						rstDocManRecords.MoveNext()
						lngCount = lngCount + 1
					Loop
        
					cmdDocManRecords = Nothing
				End If

			End If
		%>
	</form>

	<form id="frmAccess">
		<%
			sErrorDescription = ""
	
			' Get the table records.
			Dim cmdAccess = CreateObject("ADODB.Command")
			cmdAccess.CommandText = "spASRIntGetUtilityAccessRecords"
			cmdAccess.CommandType = 4	' Stored Procedure
			cmdAccess.ActiveConnection = Session("databaseConnection")

			Dim prmUtilType = cmdAccess.CreateParameter("utilType", 3, 1)	' 3=integer, 1=input
			cmdAccess.Parameters.Append(prmUtilType)
			prmUtilType.value = 9	' 9 = mail merge

			prmUtilID = cmdAccess.CreateParameter("utilID", 3, 1)	' 3=integer, 1=input
			cmdAccess.Parameters.Append(prmUtilID)
			If UCase(Session("action")) = "NEW" Then
				prmUtilID.value = 0
			Else
				prmUtilID.value = CleanNumeric(Session("utilid"))
			End If

			Dim prmFromCopy = cmdAccess.CreateParameter("fromCopy", 3, 1)	' 3=integer, 1=input
			cmdAccess.Parameters.Append(prmFromCopy)
			If UCase(Session("action")) = "COPY" Then
				prmFromCopy.value = 1
			Else
				prmFromCopy.value = 0
			End If

			Err.Number = 0
			Dim rstAccessInfo = cmdAccess.Execute
			If (Err.Number <> 0) Then
				sErrorDescription = "The access information could not be retrieved." & vbCrLf & FormatError(Err.Description)
			End If

			If Len(sErrorDescription) = 0 Then
				Dim iCount = 0
				Do While Not rstAccessInfo.EOF
					Response.Write("<INPUT type='hidden' id=txtAccess_" & iCount & " name=txtAccess_" & iCount & " value=""" & rstAccessInfo.fields("accessDefinition").value & """>" & vbCrLf)

					iCount = iCount + 1
					rstAccessInfo.MoveNext()
				Loop

				' Release the ADO recordset object.
				rstAccessInfo.close()
				'rstAccessInfo = Nothing
			End If
	
			' Release the ADO command object.
			'cmdAccess = Nothing
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
			Dim cmdDefinition = CreateObject("ADODB.Command")
			cmdDefinition.CommandText = "sp_ASRIntGetModuleParameter"
			cmdDefinition.CommandType = 4	' Stored procedure.
			cmdDefinition.ActiveConnection = Session("databaseConnection")

			prmModuleKey = cmdDefinition.CreateParameter("moduleKey", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
			cmdDefinition.Parameters.Append(prmModuleKey)
			prmModuleKey.value = "MODULE_PERSONNEL"

			Dim prmParameterKey = cmdDefinition.CreateParameter("paramKey", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
			cmdDefinition.Parameters.Append(prmParameterKey)
			prmParameterKey.value = "Param_TablePersonnel"

			Dim prmParameterValue = cmdDefinition.CreateParameter("paramValue", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
			cmdDefinition.Parameters.Append(prmParameterValue)

			Err.Number = 0
			cmdDefinition.Execute()

			Response.Write("<INPUT type='hidden' id=txtPersonnelTableID name=txtPersonnelTableID value=" & cmdDefinition.Parameters("paramValue").value & ">" & vbCrLf)
	
			'cmdDefinition = Nothing

			Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtAction name=txtAction value=" & Session("action") & ">" & vbCrLf)
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


	<form id="frmSelectionAccess" name="frmSelectionAccess">
		<input type="hidden" id="forcedHidden" name="forcedHidden" value="N">
		<input type="hidden" id="baseHidden" name="baseHidden" value="N">
		<input type="hidden" id="p1Hidden" name="p1Hidden" value="N">
		<input type="hidden" id="p2Hidden" name="p2Hidden" value="N">
		<input type="hidden" id="childHidden" name="childHidden" value="N">
		<input type="hidden" id="calcsHiddenCount" name="calcsHiddenCount" value="0">
	</form>

	<input type='hidden' id="txtTicker" name="txtTicker" value="0">
	<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">
</div>

<script type="text/javascript">
	//debugger;
	util_def_mailmerge_window_onload();
	utilDefMailmergeAddActiveXHandlers();
</script>