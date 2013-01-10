<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>

<!DOCTYPE html>
<html>
<head>
<link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css"/>
<script src="<%: Url.Content("~/Scripts/jquery-1.8.2.js") %>" type="text/javascript"></script>
<script src="<%: Url.Content("~/Scripts/openhr.js") %>" type="text/javascript"></script>
<script src="<%: Url.Content("~/Scripts/ctl_SetFont.js") %>" type="text/javascript"></script>
<title>OpenHR Intranet</title>
</head>

<body id=bdyMain bgcolor='<%=session("ConvertedDesktopColour")%>' leftmargin=20 topmargin=20 bottommargin=20 rightmargin=5>
	
<form id="frmPopup" name="frmPopup" onsubmit="return setForm();" style="visibility: hidden;display: none">
	<INPUT type=hidden id=txtSelectedID name=txtSelectedID>
	<INPUT type=hidden id=txtSelectedName name=txtSelectedName>
	<INPUT type=hidden id=txtSelectedAccess name=txtSelectedAccess>
	<INPUT type=hidden id=txtSelectedUserName name=txtSelectedUserName>
</form>

<table align=center class="outline" cellPadding=5 cellSpacing=0 width=100% height=100%>
	<tr>
		<td>
			<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0" >
				<tr height=10>
					<td colspan=3 align=center height=10>
						<H3>
<% 
	if Request("recseltype") = "picklist" then 
		Response.Write("Picklists")
	else
		Response.Write("Filters")
	end if
%>
						</H3>
					</td>
				</tr>
				<tr>
					<td width=20></td>
					<td>
<%
	' Get the order records.
	Dim cmdDefSelRecords = CreateObject("ADODB.Command")
	cmdDefSelRecords.CommandText = "spASRIntGetRecordSelection"
	cmdDefSelRecords.CommandType = 4 ' Stored Procedure
	cmdDefSelRecords.ActiveConnection = Session("databaseConnection")

	Dim prmType = cmdDefSelRecords.CreateParameter("type", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
	cmdDefSelRecords.Parameters.Append(prmType)
	prmType.value = Request("recseltype")

	Dim prmTableID = cmdDefSelRecords.CreateParameter("tableID", 3, 1) ' 3=integer, 1=input
	cmdDefSelRecords.Parameters.Append(prmTableID)
	prmTableID.value = CleanNumeric(Request("recseltableID"))

	Err.Clear()
	Dim rstDefSelRecords = cmdDefSelRecords.Execute

	' Instantiate and initialise the grid. 
	Response.Write("			<OBJECT classid=""clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"" id=ssOleDBGridDefSelRecords name=ssOleDBGridDefselRecords codebase=""cabs/COAInt_Grid.cab#version=3,1,3,6"" style=""LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:100%"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""ScrollBars"" VALUE=""4"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""_Version"" VALUE=""196616"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""DataMode"" VALUE=""2"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""Cols"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""Rows"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""BorderStyle"" VALUE=""1"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""RecordSelectors"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""GroupHeaders"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""ColumnHeaders"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""GroupHeadLines"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""HeadLines"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""FieldDelimiter"" VALUE=""(None)"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""FieldSeparator"" VALUE=""(Tab)"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""Col.Count"" VALUE=""" & rstDefselRecords.fields.count & """>" & vbCrLf)
	Response.Write("				<PARAM NAME=""stylesets.count"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""TagVariant"" VALUE=""EMPTY"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""UseGroups"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""HeadFont3D"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""Font3D"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""DividerType"" VALUE=""3"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""DividerStyle"" VALUE=""1"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""DefColWidth"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""BeveColorScheme"" VALUE=""2"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""BevelColorFrame"" VALUE=""-2147483642"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""BevelColorHighlight"" VALUE=""-2147483628"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""BevelColorShadow"" VALUE=""-2147483632"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""BevelColorFace"" VALUE=""-2147483633"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""CheckBox3D"" VALUE=""-1"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""AllowAddNew"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""AllowDelete"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""AllowUpdate"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""MultiLine"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""ActiveCellStyleSet"" VALUE="""">" & vbCrLf)
	Response.Write("				<PARAM NAME=""RowSelectionStyle"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""AllowRowSizing"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""AllowGroupSizing"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""AllowColumnSizing"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""AllowGroupMoving"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""AllowColumnMoving"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""AllowGroupSwapping"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""AllowColumnSwapping"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""AllowGroupShrinking"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""AllowColumnShrinking"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""AllowDragDrop"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""UseExactRowCount"" VALUE=""-1"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""SelectTypeCol"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""SelectTypeRow"" VALUE=""1"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""SelectByCell"" VALUE=""-1"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""BalloonHelp"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""RowNavigation"" VALUE=""1"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""CellNavigation"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""MaxSelectedRows"" VALUE=""1"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""HeadStyleSet"" VALUE="""">" & vbCrLf)
	Response.Write("				<PARAM NAME=""StyleSet"" VALUE="""">" & vbCrLf)
	Response.Write("				<PARAM NAME=""ForeColorEven"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""ForeColorOdd"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""BackColorEven"" VALUE=""16777215"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""BackColorOdd"" VALUE=""16777215"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""Levels"" VALUE=""1"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""RowHeight"" VALUE=""503"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""ExtraHeight"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""ActiveRowStyleSet"" VALUE="""">" & vbCrLf)
	Response.Write("				<PARAM NAME=""CaptionAlignment"" VALUE=""2"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""SplitterPos"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""SplitterVisible"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""Columns.Count"" VALUE=""" & rstDefSelRecords.fields.count & """>" & vbCrLf)

	for iLoop = 0 to (rstDefSelRecords.fields.count - 1)

		if rstDefSelRecords.fields(iLoop).name <> "name" then
			Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Width"" VALUE=""0"">" & vbCrLf)
			Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Visible"" VALUE=""0"">" & vbCrLf)
		else
			Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Width"" VALUE=""100000"">" & vbCrLf)
			Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Visible"" VALUE=""-1"">" & vbCrLf)
		end if
							
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Columns.Count"" VALUE=""1"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Caption"" VALUE=""" & Replace(rstDefSelRecords.fields(iLoop).name, "_", " ") & """>" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Name"" VALUE=""" & rstDefSelRecords.fields(iLoop).name & """>" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Alignment"" VALUE=""0"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").CaptionAlignment"" VALUE=""3"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Bound"" VALUE=""0"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").AllowSizing"" VALUE=""1"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").DataField"" VALUE=""Column " & iLoop & """>" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").DataType"" VALUE=""8"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Level"" VALUE=""0"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").NumberFormat"" VALUE="""">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Case"" VALUE=""0"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").FieldLen"" VALUE=""4096"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").VertScrollBar"" VALUE=""0"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Locked"" VALUE=""0"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Style"" VALUE=""0"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").ButtonsAlways"" VALUE=""0"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").RowCount"" VALUE=""0"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").ColCount"" VALUE=""1"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").HasHeadForeColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").HasHeadBackColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").HasForeColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").HasBackColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").HeadForeColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").HeadBackColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").ForeColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").BackColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").HeadStyleSet"" VALUE="""">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").StyleSet"" VALUE="""">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Nullable"" VALUE=""1"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Mask"" VALUE="""">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").PromptInclude"" VALUE=""0"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").ClipMode"" VALUE=""0"">" & vbCrLf)
		Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").PromptChar"" VALUE=""95"">" & vbCrLf)
	next 

	Response.Write("				<PARAM NAME=""UseDefaults"" VALUE=""-1"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""TabNavigation"" VALUE=""1"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""_ExtentX"" VALUE=""17330"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""_ExtentY"" VALUE=""1323"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""_StockProps"" VALUE=""79"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""Caption"" VALUE="""">" & vbCrLf)
	Response.Write("				<PARAM NAME=""ForeColor"" VALUE=""0"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""BackColor"" VALUE=""16777215"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""Enabled"" VALUE=""-1"">" & vbCrLf)
	Response.Write("				<PARAM NAME=""DataMember"" VALUE="""">" & vbCrLf)
							
	Dim lngRowCount = 0
	do while not rstDefSelRecords.EOF
		for iLoop = 0 to (rstDefSelRecords.fields.count - 1)							
			Response.Write("				<PARAM NAME=""Row(" & lngRowCount & ").Col(" & iLoop & ")"" VALUE=""" & Replace(Replace(rstDefSelRecords.Fields(iLoop).Value, "_", " "), """", "&quot;") & """>" & vbCrLf)
		next 				
		lngRowCount = lngRowCount + 1
		rstDefSelRecords.MoveNext
	loop
	Response.Write("				<PARAM NAME=""Row.Count"" VALUE=""" & lngRowCount & """>" & vbCrLf)
	Response.Write("			</OBJECT>" & vbCrLf)
%>
					</td>
					<td width=20></td>
				</tr>
				<tr height=10>
					<td height=10 colspan=3>&nbsp;</td>
				</tr>
				<tr height=10>
					<td width=20></td>
					<td height=10>
						<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD>&nbsp;</TD>
								<TD width=10>
									<INPUT id=cmdok type=button value=OK name=cmdok style="WIDTH: 80px" width="80" class="btn"
									    onclick="setForm();"
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
								<TD width=10>&nbsp;</TD>
								<TD width=10>
									<INPUT id=cmdnone type=button value=None name=cmdnone style="WIDTH: 80px" width="80" class="btn"
									    onclick="frmPopup.txtSelectedID.value=0;frmPopup.txtSelectedName.value='';frmPopup.txtSelectedAccess.value='';frmPopup.txtSelectedUserName.value='';setForm();"
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
									
								</TD>
								<TD width=10>&nbsp;</TD>
								<TD width=10>
									<INPUT id=cmdcancel type=button value=Cancel name=cmdcancel style="WIDTH: 80px" width="80" class="btn"
									    onclick="self.close();" 
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
							</tr>
						</table>
					</td>
				</tr>
			</table>	
		</td>
	</tr>
</TABLE>

<form id="frmFromOpener" name="frmFromOpener" style="visibility: hidden; display: none">
	<INPUT type=hidden id=txtSelType name=txtSelType value=<% =Request("recSelType") %>>
	<INPUT type=hidden id=txtSelTableid name=txtSelTableid value=<% =Request("recSelTableID") %>>
	<INPUT type=hidden id=txtSelCurrentID name=txtSelCurrentID value=<% =Request("recSelCurrentID") %>>
	<INPUT type=hidden id=txtSelTable name=txtSelTable value=<% =Request("recSelTable") %>>
	<INPUT type=hidden id=txtSelDefOwner name=txtSelDefOwner value=<% =Request("recSelDefOwner") %>>
	<INPUT type=hidden id=txtSelDefType name=txtSelDefType value="<% =Request("recSelDefType") %>">
</form>

<INPUT type='hidden' id=txtTicker name=txtTicker value=0>
<INPUT type='hidden' id=txtLastKeyFind name=txtLastKeyFind value="">

<script type="text/javascript">
	var frmPopup = document.getElementById("frmPopup");
	var frmFromOpener = document.getElementById("frmFromOpener");
	
	window.onload = util_recordselection_window_onload;
	
	OpenHR.addActiveXHandler("ssOleDBGridDefSelRecords", "rowcolchange", ssOleDBGridDefSelRecords_rowcolchange);
	OpenHR.addActiveXHandler("ssOleDBGridDefSelRecords", "dblClick", ssOleDBGridDefSelRecords_dblClick);
	OpenHR.addActiveXHandler("ssOleDBGridDefSelRecords", "KeyPress", ssOleDBGridDefSelRecords_KeyPress);

	function ssOleDBGridDefSelRecords_rowcolchange() {
		// Populate the textboxs with the selected rows details
		frmPopup.txtSelectedID.value = document.getElementById('ssOleDBGridDefSelRecords').Columns(0).Value;
		frmPopup.txtSelectedUserName.value = document.getElementById('ssOleDBGridDefSelRecords').Columns("username").Value;
		frmPopup.txtSelectedAccess.value = document.getElementById('ssOleDBGridDefSelRecords').Columns("access").Value;
		frmPopup.txtSelectedName.value = document.getElementById('ssOleDBGridDefSelRecords').Columns("name").Value;
	}

	function ssOleDBGridDefSelRecords_dblClick() {
		setForm();
	}

	function ssOleDBGridDefSelRecords_KeyPress(iKeyAscii) {

		var sFind,
		    iLastTick,
		    txtLastKeyFind = document.getElementById("txtLastKeyFind"),
		    txtTicker = document.getElementById("txtTicker");

		if ((iKeyAscii >= 32) && (iKeyAscii <= 255)) {
			var dtTicker = new Date();
			var iThisTick = new Number(dtTicker.getTime());
			if (txtLastKeyFind.value.length > 0) {
				iLastTick = new Number(txtTicker.value);
			}
			else {
				iLastTick = new Number("0");
			}

			if (iThisTick > (iLastTick + 1500)) {
				sFind = String.fromCharCode(iKeyAscii);
			}
			else {
				sFind = txtLastKeyFind.value + String.fromCharCode(iKeyAscii);
			}

			txtTicker.value = iThisTick;
			txtLastKeyFind.value = sFind;

			locateRecord(sFind);
		}
	}

	function locateRecord(psSearchFor) {

		var fFound = false;
		var ssOleDBGridDefSelRecords = document.getElementById("ssOleDBGridDefSelRecords");
		ssOleDBGridDefSelRecords.redraw = false;
		ssOleDBGridDefSelRecords.MoveLast();
		ssOleDBGridDefSelRecords.MoveFirst();

		for (var iIndex = 1; iIndex <= ssOleDBGridDefSelRecords.rows; iIndex++) {
			var sGridValue = new String(ssOleDBGridDefSelRecords.Columns("name").value);
			sGridValue = sGridValue.substr(0, psSearchFor.length).toUpperCase();
			if (sGridValue == psSearchFor.toUpperCase()) {
				ssOleDBGridDefSelRecords.SelBookmarks.Add(ssOleDBGridDefSelRecords.Bookmark);
				fFound = true;
				break;
			}

			if (iIndex < ssOleDBGridDefSelRecords.rows) {
				ssOleDBGridDefSelRecords.MoveNext();
			}
			else {
				break;
			}
		}

		if ((fFound == false) && (ssOleDBGridDefSelRecords.rows > 0)) {
			// Select the top row.
			ssOleDBGridDefSelRecords.MoveFirst();
			ssOleDBGridDefSelRecords.SelBookmarks.Add(ssOleDBGridDefSelRecords.Bookmark);
		}

		ssOleDBGridDefSelRecords.redraw = true;
	}

	function setForm() {

		if (frmFromOpener.txtSelDefOwner.value == "0" && frmPopup.txtSelectedAccess.value == "HD") {

			var sMessage = "Unable to select this " + frmFromOpener.txtSelType.value +
				" as it is a hidden " + frmFromOpener.txtSelType.value +
					" and you are not the owner of this definition";

			OpenHR.messageBox(sMessage, 48, frmFromOpener.txtSelDefType.value);

			self.close();
			return false;
		}

		//we are doing this for the base table
		if (frmFromOpener.txtSelTable.value == 'base') {
			//we are doing this for the picklist
			if (frmFromOpener.txtSelType.value == 'picklist') {
				window.dialogArguments.document.getElementById('txtBasePicklist').value = frmPopup.txtSelectedName.value;
				window.dialogArguments.document.getElementById('txtBasePicklistID').value = frmPopup.txtSelectedID.value;

				//if its hidden, set the relevant textbox value
				if (frmPopup.txtSelectedAccess.value == "HD") {
					window.dialogArguments.document.getElementById('baseHidden').value = 'Y';
				}
				else {
					window.dialogArguments.document.getElementById('baseHidden').value = '';
				}

				try {
					window.dialogArguments.document.getElementById('cmdBasePicklist').focus();
				}
				catch (e) {
				}
			}

			//we are doing this for the filter
			if (frmFromOpener.txtSelType.value == 'filter') {
				//we are doing this for the picklist
				window.dialogArguments.document.getElementById('txtBaseFilter').value = frmPopup.txtSelectedName.value;
				window.dialogArguments.document.getElementById('txtBaseFilterID').value = frmPopup.txtSelectedID.value;

				//if its hidden, set the relevant textbox value
				if (frmPopup.txtSelectedAccess.value == "HD") {
					window.dialogArguments.document.getElementById('baseHidden').value = 'Y';
				}
				else {
					window.dialogArguments.document.getElementById('baseHidden').value = '';
				}

				try {
					window.dialogArguments.document.getElementById('cmdBaseFilter').focus();
				}
				catch (e) {
				}
			}
		}

		//we are doing this for the parent 1 table
		if (frmFromOpener.txtSelTable.value == 'p1') {
			//we are doing this for the picklist
			if (frmFromOpener.txtSelType.value == 'picklist') {
				window.dialogArguments.document.getElementById('txtParent1Picklist').value = frmPopup.txtSelectedName.value;
				window.dialogArguments.document.getElementById('txtParent1PicklistID').value = frmPopup.txtSelectedID.value;

				//if its hidden, set the relevant textbox value
				if (frmPopup.txtSelectedAccess.value == "HD") {
					window.dialogArguments.document.getElementById('p1Hidden').value = 'Y';
				}
				else {
					window.dialogArguments.document.getElementById('p1Hidden').value = '';
				}

				try {
					window.dialogArguments.document.getElementById('cmdParent1Picklist').focus();
				}
				catch (e) {
				}
			}

			//we are doing this for the filter
			if (frmFromOpener.txtSelType.value == 'filter') {
				//we are doing this for the picklist
				window.dialogArguments.document.getElementById('txtParent1Filter').value = frmPopup.txtSelectedName.value;
				window.dialogArguments.document.getElementById('txtParent1FilterID').value = frmPopup.txtSelectedID.value;

				//if its hidden, set the relevant textbox value
				if (frmPopup.txtSelectedAccess.value == "HD") {
					window.dialogArguments.document.getElementById('p1Hidden').value = 'Y';
				}
				else {
					window.dialogArguments.document.getElementById('p1Hidden').value = '';
				}

				try {
					window.dialogArguments.document.getElementById('cmdParent1Filter').focus();
				}
				catch (e) {
				}
			}
		}

		//we are doing this for the parent 2 table
		if (frmFromOpener.txtSelTable.value == 'p2') {
			//we are doing this for the picklist
			if (frmFromOpener.txtSelType.value == 'picklist') {
				window.dialogArguments.document.getElementById('txtParent2Picklist').value = frmPopup.txtSelectedName.value;
				window.dialogArguments.document.getElementById('txtParent2PicklistID').value = frmPopup.txtSelectedID.value;

				//if its hidden, set the relevant textbox value
				if (frmPopup.txtSelectedAccess.value == "HD") {
					window.dialogArguments.document.getElementById('p2Hidden').value = 'Y';
				}
				else {
					window.dialogArguments.document.getElementById('p2Hidden').value = '';
				}

				try {
					window.dialogArguments.document.getElementById('cmdParent2Picklist').focus();
				}
				catch (e) {
				}
			}

			//we are doing this for the filter
			if (frmFromOpener.txtSelType.value == 'filter') {
				//we are doing this for the picklist
				window.dialogArguments.document.getElementById('txtParent2Filter').value = frmPopup.txtSelectedName.value;
				window.dialogArguments.document.getElementById('txtParent2FilterID').value = frmPopup.txtSelectedID.value;

				//if its hidden, set the relevant textbox value
				if (frmPopup.txtSelectedAccess.value == "HD") {
					window.dialogArguments.document.getElementById('p2Hidden').value = 'Y';
				}
				else {
					window.dialogArguments.document.getElementById('p2Hidden').value = '';
				}

				try {
					window.dialogArguments.document.getElementById('cmdParent2Filter').focus();
				}
				catch (e) {
				}
			}
		}

		//we are doing this for the child table
		if (frmFromOpener.txtSelTable.value == 'child') {
			//we are doing this for the filter
			if (frmFromOpener.txtSelType.value == 'filter') {
				//we are doing this for the filter
				window.dialogArguments.document.getElementById('txtChildFilter').value = frmPopup.txtSelectedName.value;
				window.dialogArguments.document.getElementById('txtChildFilterID').value = frmPopup.txtSelectedID.value;

				//if its hidden, set the relevant textbox value
				if (frmPopup.txtSelectedAccess.value == "HD") {
					window.dialogArguments.document.getElementById('childHidden').value = 'Y';
				}
				else {
					window.dialogArguments.document.getElementById('childHidden').value = '';
				}

				try {
					window.dialogArguments.document.getElementById('cmdChildFilter').focus();
				}
				catch (e) {
				}
			}
		}

		// Are we are doing this for a standard report
		if (frmFromOpener.txtSelTable.value == 'standardreport') {
			//we are doing this for the filter
			if (frmFromOpener.txtSelType.value == 'filter') {
				//we are doing this for the filter
				window.dialogArguments.document.getElementById('txtFilter').value = frmPopup.txtSelectedName.value;
				window.dialogArguments.document.getElementById('txtFilterID').value = frmPopup.txtSelectedID.value;

				try {
					window.dialogArguments.document.getElementById('cmdFilter').focus();
				}
				catch (e) {
				}
			}

			if (frmFromOpener.txtSelType.value == 'picklist') {
				window.dialogArguments.document.getElementById('txtPicklist').value = frmPopup.txtSelectedName.value;
				window.dialogArguments.document.getElementById('txtPicklistID').value = frmPopup.txtSelectedID.value;

				try {
					window.dialogArguments.document.getElementById('cmdPicklist').focus();
				}
				catch (e) {
				}
			}
		}

		// Are we are doing this for a calendar report
		if (frmFromOpener.txtSelTable.value == 'event') {
			//we are doing this for the filter
			if (frmFromOpener.txtSelType.value == 'filter') {
				//we are doing this for the filter
				window.dialogArguments.document.getElementById('txtEventFilter').value = frmPopup.txtSelectedName.value;
				window.dialogArguments.document.getElementById('txtEventFilterID').value = frmPopup.txtSelectedID.value;

				//if its hidden, set the relevant textbox value
				if (frmPopup.txtSelectedAccess.value == "HD") {
					window.dialogArguments.document.getElementById('baseHidden').value = 'Y';
				}
				else {
					window.dialogArguments.document.getElementById('baseHidden').value = '';
				}

				try {
					window.dialogArguments.document.getElementById('cmdEventFilter').focus();
				}
				catch (e) {
				}
			}
		}

		self.close();
		return false;
	}

	function util_recordselection_window_onload() {

		var ssOleDBGridDefSelRecords = document.getElementById("ssOleDBGridDefSelRecords"),
		    iResizeBy,
		    iNewWidth,
		    iNewHeight,
		    bdyMain = document.getElementById("bdyMain");

		setGridFont(ssOleDBGridDefSelRecords);

		// Resize the popup.
		iResizeBy = bdyMain.scrollWidth - bdyMain.clientWidth;
		if (bdyMain.offsetWidth + iResizeBy > screen.width) {
			window.dialogWidth = new String(screen.width) + "px";
		}
		else {
			iNewWidth = new Number(window.dialogWidth.substr(0, window.dialogWidth.length - 2));
			iNewWidth = iNewWidth + iResizeBy;
			window.dialogWidth = new String(iNewWidth) + "px";
		}

		iResizeBy = bdyMain.scrollHeight - bdyMain.clientHeight;
		if (bdyMain.offsetHeight + iResizeBy > screen.height) {
			window.dialogHeight = new String(screen.height) + "px";
		}
		else {
			iNewHeight = new Number(window.dialogHeight.substr(0, window.dialogHeight.length - 2));
			iNewHeight = iNewHeight + iResizeBy;
			window.dialogHeight = new String(iNewHeight) + "px";
		}
	}
</script>
</body>
</html>
