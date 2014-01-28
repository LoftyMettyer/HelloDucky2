<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<%
		Session("selectionType") = Request("selectionType")
	session("selectionTableID") = Request("txtTableID")
	session("selectedID") = Request("selectedID")
%>

<!DOCTYPE html>

<html>
<head runat="server">
		<title>OpenHR Intranet</title>

	<script src="<%: Url.Content("~/bundles/jQuery")%>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/bundles/jQueryUI7")%>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/bundles/OpenHR_General")%>" type="text/javascript"></script>
	<script id="officebarscript" src="<%: Url.Content("~/Scripts/officebar/jquery.officebar.js") %>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/Scripts/ctl_SetFont.js") %>" type="text/javascript"></script>
	<link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.LatestContent("~/Content/Site.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />
	<link id="DMIthemeLink" href="<%: Url.LatestContent("~/Content/themes/" & Session("ui-theme").ToString() & "/jquery-ui.min.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/general_enclosed_foundicons.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/font-awesome.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/fonts/SSI80v194934/style.css")%>" rel="stylesheet" />

	<object classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" id="Microsoft_Licensed_Class_Manager_1_0">
		<param name="LPKPath" value="lpks/main.lpk">
	</object>

	<script type="text/javascript">

				function fieldRec_window_onload() {
						
						var fOK = true;

						$("input[type=submit], input[type=button], button")
							.button();
						$("input").addClass("ui-widget ui-widget-content ui-corner-all");
						$("input").removeClass("text");

						$("select").addClass("ui-widget ui-widget-content ui-corner-tl ui-corner-bl");
						$("select").removeClass("text");


						cmdCancel.focus();
	
						// Set focus onto one of the form controls. 
						// NB. This needs to be done before making any reference to the grid
						ssOleDBGridSelRecords.focus();
						locateRecordID(frmUseful.txtSelectedID.value);

						setGridFont(ssOleDBGridSelRecords);
	
						refreshControls();

						// Resize the popup.
						iResizeBy = bdyMain.scrollWidth	- bdyMain.clientWidth;
						if (bdyMain.offsetWidth + iResizeBy > screen.width) {
								window.dialogWidth = new String(screen.width) + "px";
						}
						else {
								iNewWidth = new Number(window.dialogWidth.substr(0, window.dialogWidth.length-2));
								iNewWidth = iNewWidth + iResizeBy;
								window.dialogWidth = new String(iNewWidth) + "px";
						}

						iResizeBy = bdyMain.scrollHeight	- bdyMain.clientHeight;
						if (bdyMain.offsetHeight + iResizeBy > screen.height) {
								window.dialogHeight = new String(screen.height) + "px";
						}
						else {
								iNewHeight = new Number(window.dialogHeight.substr(0, window.dialogHeight.length-2));
								iNewHeight = iNewHeight + iResizeBy;
								window.dialogHeight = new String(iNewHeight) + "px";
						}
				}

				function refreshControls() {
					button_disable(cmdOK, (ssOleDBGridSelRecords.SelBookmarks.Count == 0));
				}

				function setForm()
				{
						//we are doing this for the order
						if (frmUseful.txtSelectionType.value == 'ORDER') {
								window.dialogArguments.document.getElementById('txtFieldRecOrder').value = frmPopup.txtSelectedName.value;
								window.dialogArguments.document.getElementById('txtChildFieldOrderID').value = frmPopup.txtSelectedID.value;

								try {
										window.dialogArguments.document.getElementById('btnFieldRecOrder').focus();
								}
								catch(e) {
								}
						}
						else {
								//we are doing this for the filter
								window.dialogArguments.document.getElementById('txtFieldRecFilter').value = frmPopup.txtSelectedName.value;
								window.dialogArguments.document.getElementById('txtChildFieldFilterID').value =  frmPopup.txtSelectedID.value;
			
								//if its hidden, set the relevant textbox value
								if (frmPopup.txtSelectedAccess.value == "HD") {
										window.dialogArguments.document.getElementById('txtChildFieldFilterHidden').value = 'Y';
								}
								else {
										window.dialogArguments.document.getElementById('txtChildFieldFilterHidden').value = '';
								}
			
								try {
										window.dialogArguments.document.getElementById('btnFieldRecFilter').focus();
								}
								catch(e) {
								}
						}

						self.close();
						return false;
				}

				function makeSelection()
				{
						frmPopup.txtSelectedID.value = ssOleDBGridSelRecords.Columns("id").Value; 	
						frmPopup.txtSelectedUserName.value = ssOleDBGridSelRecords.Columns("username").Value;
						frmPopup.txtSelectedAccess.value = ssOleDBGridSelRecords.Columns("access").Value;
						frmPopup.txtSelectedName.value = ssOleDBGridSelRecords.Columns("name").Value;
						setForm();
				}

				function clearSelection()
				{
						frmPopup.txtSelectedID.value=0;
						frmPopup.txtSelectedName.value='';
						frmPopup.txtSelectedAccess.value='';
						frmPopup.txtSelectedUserName.value='';
						setForm();
				}

				function locateRecord(psSearchFor) {
					var fFound;

						fFound = false;
	
						ssOleDBGridSelRecords.redraw = false;

						ssOleDBGridSelRecords.MoveLast();
						ssOleDBGridSelRecords.MoveFirst();

						for (iIndex = 1; iIndex <= ssOleDBGridSelRecords.rows; iIndex++) {	
								var sGridValue = new String(ssOleDBGridSelRecords.Columns("name").value);
								sGridValue = sGridValue.substr(0, psSearchFor.length).toUpperCase();
								if (sGridValue == psSearchFor.toUpperCase()) {
										ssOleDBGridSelRecords.SelBookmarks.Add(ssOleDBGridSelRecords.Bookmark);
										fFound = true;
										break;
								}

								if (iIndex < ssOleDBGridSelRecords.rows) {
										ssOleDBGridSelRecords.MoveNext();
								}
								else {
										break;
								}
						}

						if ((fFound == false) && (ssOleDBGridSelRecords.rows > 0)) {
								// Select the top row.
								ssOleDBGridSelRecords.MoveFirst();
								ssOleDBGridSelRecords.SelBookmarks.Add(ssOleDBGridSelRecords.Bookmark);
						}

						ssOleDBGridSelRecords.redraw = true;
				}

				function locateRecordID(piRecordID) {
					var fFound;

						fFound = false;
	
						ssOleDBGridSelRecords.redraw = false;

						ssOleDBGridSelRecords.MoveLast();
						ssOleDBGridSelRecords.MoveFirst();

						if (frmUseful.txtSelectedID.value > 0) {
								for (iIndex = 1; iIndex <= ssOleDBGridSelRecords.rows; iIndex++) {	
										var sGridValue = new String(ssOleDBGridSelRecords.Columns("id").value);
										if (sGridValue == piRecordID) {
												ssOleDBGridSelRecords.SelBookmarks.Add(ssOleDBGridSelRecords.Bookmark);
												fFound = true;
												break;
										}

										if (iIndex < ssOleDBGridSelRecords.rows) {
												ssOleDBGridSelRecords.MoveNext();
										}
										else {
												break;
										}
								}
						}
	
						if ((fFound == false) && (ssOleDBGridSelRecords.rows > 0)) {
								// Select the top row.
								ssOleDBGridSelRecords.MoveFirst();
								ssOleDBGridSelRecords.SelBookmarks.Add(ssOleDBGridSelRecords.Bookmark);
						}

						ssOleDBGridSelRecords.redraw = true;
				}

				function fieldrec_addhandlers() {        
					OpenHR.addActiveXHandler("ssOleDBGridSelRecords", "rowcolchange", "ssOleDBGridSelRecords_rowcolchange()");
					OpenHR.addActiveXHandler("ssOleDBGridSelRecords", "dblClick", "ssOleDBGridSelRecords_dblClick()");
					OpenHR.addActiveXHandler("ssOleDBGridSelRecords", "KeyPress", "ssOleDBGridSelRecords_KeyPress()");
				}

				function ssOleDBGridSelRecords_rowcolchange() {
						// Populate the textboxs with the selected rows details
						refreshControls();
				}

				function ssOleDBGridSelRecords_dblClick() {
						makeSelection();        
				}

				function ssOleDBGridSelRecords_KeyPress(iKeyAscii) {

						if ((iKeyAscii >= 32) && (iKeyAscii <= 255)) {	
								var dtTicker = new Date();
								var iThisTick = new Number(dtTicker.getTime());
								if (txtLastKeyFind.value.length > 0) {
										var iLastTick = new Number(txtTicker.value);
								}
								else {
										var iLastTick = new Number("0");
								}
		
								if (iThisTick > (iLastTick + 1500)) {
										var sFind = String.fromCharCode(iKeyAscii);
								}
								else {
										var sFind = txtLastKeyFind.value + String.fromCharCode(iKeyAscii);
								}
		
								txtTicker.value = iThisTick;
								txtLastKeyFind.value = sFind;

								locateRecord(sFind);
						}
				
				}
	</script>
	 
</head>

<body id=bdyMain >
		
		<table align=center class="outline" cellPadding=5 cellSpacing=0 width=100% height=100%>
	<tr>
		<td>
			<table align=center class="invisible" cellspacing=0 cellpadding=0 width=100% height=100%>
				<tr height=10>
					<td colspan=3 align=center height=10>
						<H3 align=center>
<% 
	if ucase(session("selectionType")) = ucase("order") then 
				Response.Write("Select Order")
		Else
				Response.Write("Select Filter")
		End If
%>
						</H3>
					</td>
				</tr>
				<tr>
					<td style="width: 20px;"></td>
					<td style="height: 350px">
						<%
							Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

							Dim lngRowCount As Long
							Dim rstSelRecords As DataTable
							
							If UCase(Session("selectionType")) = UCase("order") Then
			
								rstSelRecords = objDataAccess.GetFromSP("spASRIntGetAvailableOrdersInfo" _
										, New SqlParameter("plngTableID", SqlDbType.Int) With {.Value = CInt(CleanNumeric(Session("selectionTableID")))})
								
							Else

								rstSelRecords = objDataAccess.GetFromSP("spASRIntGetAvailableFiltersInfo" _
										, New SqlParameter("plngTableID", SqlDbType.Int) With {.Value = CInt(CleanNumeric(Session("selectionTableID")))} _
										, New SqlParameter("psUserName", SqlDbType.VarChar, 255) With {.Value = CStr(Session("username"))})
																
							End If

							' Instantiate and initialise the grid. 
						%>
						<object classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"
							id="ssOleDBGridSelRecords"
							name="ssOleDBGridSelRecords"
							codebase="cabs/COAInt_Grid.cab#version=3,1,3,6"
							style="LEFT: 0px; TOP: 0px; WIDTH: 100%; HEIGHT: 100%">
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
							<param name="Col.Count" value="<%=rstSelRecords.Columns.Count%>">
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
							<param name="Columns.Count" value="<%=rstSelRecords.columns.count%>">
							<%
								For iLoop = 0 To (rstSelRecords.columns.count - 1)
									If rstSelRecords.Columns(iLoop).ColumnName <> "name" Then
							%>
							<param name="Columns(<%=iLoop%>).Width" value="0">
							<param name="Columns(<%=iLoop%>).Visible" value="0">
							<%
							Else
							%>
							<param name="Columns(<%=iLoop%>).Width" value="100000">
							<param name="Columns(<%=iLoop%>).Visible" value="-1">
							<%
							End If
							%>
							<param name="Columns(<%=iLoop%>).Columns.Count" value="1">
							<param name="Columns(<%=iLoop%>).Caption" value="<%=Replace(rstSelRecords.Columns(iLoop).ColumnName, "_", " ")%>">
							<param name="Columns(<%=iLoop%>).Name" value="<%=rstSelRecords.Columns(iLoop).ColumnName%>">
							<param name="Columns(<%=iLoop%>).Alignment" value="0">
							<param name="Columns(<%=iLoop%>).CaptionAlignment" value="3">
							<param name="Columns(<%=iLoop%>).Bound" value="0">
							<param name="Columns(<%=iLoop%>).AllowSizing" value="1">
							<param name="Columns(<%=iLoop%>).DataField" value="Column <%=iLoop%>">
							<param name="Columns(<%=iLoop%>).DataType" value="8">
							<param name="Columns(<%=iLoop%>).Level" value="0">
							<param name="Columns(<%=iLoop%>).NumberFormat" value="">
							<param name="Columns(<%=iLoop%>).Case" value="0">
							<param name="Columns(<%=iLoop%>).FieldLen" value="4096">
							<param name="Columns(<%=iLoop%>).VertScrollBar" value="0">
							<param name="Columns(<%=iLoop%>).Locked" value="0">
							<param name="Columns(<%=iLoop%>).Style" value="0">
							<param name="Columns(<%=iLoop%>).ButtonsAlways" value="0">
							<param name="Columns(<%=iLoop%>).RowCount" value="0">
							<param name="Columns(<%=iLoop%>).ColCount" value="1">
							<param name="Columns(<%=iLoop%>).HasHeadForeColor" value="0">
							<param name="Columns(<%=iLoop%>).HasHeadBackColor" value="0">
							<param name="Columns(<%=iLoop%>).HasForeColor" value="0">
							<param name="Columns(<%=iLoop%>).HasBackColor" value="0">
							<param name="Columns(<%=iLoop%>).HeadForeColor" value="0">
							<param name="Columns(<%=iLoop%>).HeadBackColor" value="0">
							<param name="Columns(<%=iLoop%>).ForeColor" value="0">
							<param name="Columns(<%=iLoop%>).BackColor" value="0">
							<param name="Columns(<%=iLoop%>).HeadStyleSet" value="">
							<param name="Columns(<%=iLoop%>).StyleSet" value="">
							<param name="Columns(<%=iLoop%>).Nullable" value="1">
							<param name="Columns(<%=iLoop%>).Mask" value="">
							<param name="Columns(<%=iLoop%>).PromptInclude" value="0">
							<param name="Columns(<%=iLoop%>).ClipMode" value="0">
							<param name="Columns(<%=iLoop%>).PromptChar" value="95">
							<%
							Next
							%>
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
							<%								
								lngRowCount = 0
								
								For Each objRow As DataRow In rstSelRecords.Rows

									For iLoop = 0 To (rstSelRecords.Columns.Count - 1)
							%>
							<param name="Row(<%=lngRowCount%>).Col(<%=iLoop%>)" value="<%=Replace(Replace(objRow(iLoop).ToString(), "_", " "), "", "&quot;")%>">
							<%
							Next
							lngRowCount += 1
						Next
							%>
							<param name="Row.Count" value="<%=lngRowCount%>">
						</object>
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
									<INPUT id=cmdOK type=button value=OK name=cmdOK class="btn" style="WIDTH: 80px" width="80"
											onclick="makeSelection()" />
								</TD>
								<TD width=10>&nbsp;</TD>
								<TD width=10>
									<INPUT id=cmdnone type=button value=None name=cmdnone class="btn" style="WIDTH: 80px" width="80"
											onclick="clearSelection()" />
								</TD>
								<TD width=10>&nbsp;</TD>
								<TD width=10>
									<INPUT id=cmdCancel type=button value=Cancel name=cmdCancel class="btn" style="WIDTH: 80px" width="80"
											onclick="self.close()" />
								</TD>
							</TR>
						</TABLE>
					</td>
					<td width=20></td>
				</tr>
			</TABLE>
		</td>
	</tr>
</table>

<INPUT type='hidden' id=txtTicker name=txtTicker value=0>
<INPUT type='hidden' id=txtLastKeyFind name=txtLastKeyFind value="">

	<form id="frmUseful" name="frmUseful" style="visibility: hidden; display: none">
		<input type='hidden' id="txtSelectionType" name="txtSelectionType" value='<%=Request("selectionType")%>'>
		<input type='hidden' id="txtTableID" name="txtTableID" value='<%=Request("txtTableID")%>'>
		<input type='hidden' id="txtSelectedID" name="txtSelectedID" value='<%=Request("selectedID")%>'>
	</form>

	<form id="frmPopup" name="frmPopup" style="visibility: hidden; display: none">
		<input type="hidden" id="Hidden1" name="txtSelectedID">
		<input type="hidden" id="txtSelectedName" name="txtSelectedName">
		<input type="hidden" id="txtSelectedAccess" name="txtSelectedAccess">
		<input type="hidden" id="txtSelectedUserName" name="txtSelectedUserName">
	</form>

</body>
</html>

		<script type="text/javascript">
				fieldrec_addhandlers();
				fieldRec_window_onload();
		</script>
