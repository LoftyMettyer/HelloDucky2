<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@Import namespace="DMI.NET" %>

<script src="<%: Url.Content("~/Scripts/ctl_SetFont.js") %>" type="text/javascript"></script>

<script type="text/javascript">
	function tbBulkBooking_onload() {		
		var frmBulkBooking = document.getElementById("frmBulkBooking");

		setGridFont(frmBulkBooking.ssOleDBGridFindRecords);
		
		$("#optionframe").attr("data-framesource", "TBBULKBOOKING");
		$("#workframe").hide();
		$("#optionframe").show();
		
		frmBulkBooking.cmdCancel.focus();

		//TODO: window.parent.frames("workframe").document.forms("frmFindForm").ssOleDBGridFindRecords.style.visibility = "hidden";

		tbrefreshControls();
		menu_refreshMenu();
	}
</script>

<script type="text/javascript">
	function ok()
	{  
		var sSelectedIDs = "";
		var frmBulkBooking = document.getElementById("frmBulkBooking");
		
		frmBulkBooking.ssOleDBGridFindRecords.redraw = false;
		frmBulkBooking.ssOleDBGridFindRecords.MoveFirst();
		for (var iIndex = 1; iIndex <= frmBulkBooking.ssOleDBGridFindRecords.rows; iIndex++) {	
			var sGridValue = new String(frmBulkBooking.ssOleDBGridFindRecords.Columns("ID").value);

			if (sSelectedIDs.length > 0) {
				sSelectedIDs = sSelectedIDs + ",";
			}

			sSelectedIDs = sSelectedIDs + sGridValue;

			if (iIndex < frmBulkBooking.ssOleDBGridFindRecords.rows) {
				frmBulkBooking.ssOleDBGridFindRecords.MoveNext();
			}
			else {
				break;
			}
		}
		frmBulkBooking.ssOleDBGridFindRecords.redraw = true;
		//TODO: window.parent.frames("workframe").document.forms("frmFindForm").ssOleDBGridFindRecords.style.visibility = "visible";
		
		var frmGotoOption = document.getElementById("frmGotoOption");

		frmGotoOption.txtGotoOptionAction.value = "SELECTBULKBOOKINGS";
		frmGotoOption.txtGotoOptionRecordID.value = $("#txtOptionRecordID").val();
		frmGotoOption.txtGotoOptionLinkRecordID.value = sSelectedIDs;
		<%
    if session("TB_TBStatusPExists") then
%>
		frmGotoOption.txtGotoOptionLookupValue.value = frmBulkBooking.selStatus.options[frmBulkBooking.selStatus.selectedIndex].value;
		<%
    else
%>
		frmGotoOption.txtGotoOptionLookupValue.value = "B";
		<%
    end if 
%>

		frmGotoOption.txtGotoOptionPage.value = "emptyoption";
		OpenHR.submitForm(frmGotoOption);
	}

	function cancel()
	{  		
		//$("#optionframe").attr("data-framesource", "TBBULKBOOKING");
		$("#optionframe").hide();
		$("#workframe").show();
		

		var frmGotoOption = document.getElementById("frmGotoOption");
	
		frmGotoOption.txtGotoOptionAction.value = "CANCEL";
		frmGotoOption.txtGotoOptionLinkRecordID.value = 0;
		frmGotoOption.txtGotoOptionPage.value = "emptyoption";
		OpenHR.submitForm(frmGotoOption);
	}

	function locateRecord(psFileName) {
		var fFound;
		var frmBulkBooking = document.getElementById("frmBulkBooking");
	
		fFound = false;
	
		frmBulkBooking.ssOleDBGridFindRecords.redraw = false;

		frmBulkBooking.ssOleDBGridFindRecords.MoveLast();
		frmBulkBooking.ssOleDBGridFindRecords.MoveFirst();

		for (var iIndex = 1; iIndex <= frmBulkBooking.ssOleDBGridFindRecords.rows; iIndex++) {	
			var sGridValue = new String(frmBulkBooking.ssOleDBGridFindRecords.Columns(0).value);
			sGridValue = sGridValue.substr(0, psFileName.length).toUpperCase();
			if (sGridValue == psFileName.toUpperCase()) {
				frmBulkBooking.ssOleDBGridFindRecords.SelBookmarks.Add(frmBulkBooking.ssOleDBGridFindRecords.Bookmark);
				fFound = true;
				break;
			}

			if (iIndex < frmBulkBooking.ssOleDBGridFindRecords.rows) {
				frmBulkBooking.ssOleDBGridFindRecords.MoveNext();
			}
			else {
				break;
			}
		}

		if ((fFound == false) && (frmBulkBooking.ssOleDBGridFindRecords.rows > 0)) {
			// Select the top row.
			frmBulkBooking.ssOleDBGridFindRecords.MoveFirst();
			frmBulkBooking.ssOleDBGridFindRecords.SelBookmarks.Add(frmBulkBooking.ssOleDBGridFindRecords.Bookmark);
		}

		frmBulkBooking.ssOleDBGridFindRecords.redraw = true;
	}

	function tbrefreshControls()
	{
		var fNoneSelected;
		var frmBulkBooking = document.getElementById("frmBulkBooking");
	
		fNoneSelected = (frmBulkBooking.ssOleDBGridFindRecords.SelBookmarks.Count == 0);

		button_disable(frmBulkBooking.cmdRemove, fNoneSelected);
		button_disable(frmBulkBooking.cmdRemoveAll, fNoneSelected);
		button_disable(frmBulkBooking.cmdOK, fNoneSelected);
	}

	function add()
	{	
		var sURL;
		var frmBookingSelection = document.getElementById("frmBookingSelection");
	
		frmBookingSelection.selectionType.value = "ALL";

		sURL = "tbBulkBookingSelectionMain" +
			"?selectionType=" + frmBookingSelection.selectionType.value;
		openDialog(sURL, (screen.width)/3,(screen.height)/2);
	}

	function filteredAdd()
	{	
		var sURL;

		frmBookingSelection.selectionType.value = "FILTER";
	
		sURL = "tbBulkBookingSelectionMain" +
			"?selectionType=" + frmBookingSelection.selectionType.value;
		openDialog(sURL, (screen.width)/3,(screen.height)/2);
	}

	function addPicklist()
	{	
		var sURL;

		frmBookingSelection.selectionType.value = "PICKLIST";

		sURL = "tbBulkBookingSelectionMain" +
			"?selectionType=" + frmBookingSelection.selectionType.value;
		openDialog(sURL, (screen.width)/3,(screen.height)/2);
	}

	function remove()
	{	
		var frmBulkBooking = document.getElementById("frmBulkBooking");
		var iRowIndex;
	
		var iCount = frmBulkBooking.ssOleDBGridFindRecords.selbookmarks.Count();		
		for (var i=iCount-1; i >= 0; i--) {
			frmBulkBooking.ssOleDBGridFindRecords.bookmark = frmBulkBooking.ssOleDBGridFindRecords.selbookmarks(i);
			iRowIndex = frmBulkBooking.ssOleDBGridFindRecords.AddItemRowIndex(frmBulkBooking.ssOleDBGridFindRecords.Bookmark);
				
			if ((frmBulkBooking.ssOleDBGridFindRecords.Rows == 1) && (iRowIndex == 0)) {
				frmBulkBooking.ssOleDBGridFindRecords.RemoveAll();
			}
			else {
				frmBulkBooking.ssOleDBGridFindRecords.RemoveItem(iRowIndex);
			}
		}

		// Select the top row if there is one.
		if (frmBulkBooking.ssOleDBGridFindRecords.rows > 0) {
			frmBulkBooking.ssOleDBGridFindRecords.MoveFirst();
			frmBulkBooking.ssOleDBGridFindRecords.SelBookmarks.Add(frmBulkBooking.ssOleDBGridFindRecords.Bookmark);
		}
	
		tbrefreshControls();
	}

	function removeAll()
	{	
		var frmBulkBooking = document.getElementById("frmBulkBooking");
		frmBulkBooking.ssOleDBGridFindRecords.removeAll();
		tbrefreshControls();
	}

	function makeSelection(psType, piID, psPrompts) {

		/* Get the current selected delegate IDs. */
		var frmBulkBooking = document.getElementById("frmBulkBooking");
		var sSelectedIDs = "";
		frmBulkBooking.ssOleDBGridFindRecords.redraw = false;
		if (frmBulkBooking.ssOleDBGridFindRecords.rows > 0) {
			frmBulkBooking.ssOleDBGridFindRecords.MoveFirst();
		}
		var sRecordID;
		for (var iIndex = 1; iIndex <= frmBulkBooking.ssOleDBGridFindRecords.rows; iIndex++) {	
			sRecordID = new String(frmBulkBooking.ssOleDBGridFindRecords.Columns("ID").Value);

			if (sSelectedIDs.length > 0) {
				sSelectedIDs = sSelectedIDs + ",";
			}
			sSelectedIDs = sSelectedIDs + sRecordID;
			
			if (iIndex < frmBulkBooking.ssOleDBGridFindRecords.rows) {
				frmBulkBooking.ssOleDBGridFindRecords.MoveNext();
			}
			else {
				break;
			}
		}
		frmBulkBooking.ssOleDBGridFindRecords.redraw = true;

		if ((psType == "ALL") && (psPrompts.length > 0)) {
			if (sSelectedIDs.length > 0) {
				sSelectedIDs = sSelectedIDs + ",";
			}
			sSelectedIDs = sSelectedIDs + psPrompts;
		}


		// Get the optionData.asp to get the required records.
		var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
		optionDataForm.txtOptionAction.value = "GETBULKBOOKINGSELECTION";
		optionDataForm.txtOptionPageAction.value = psType;
		optionDataForm.txtOptionRecordID.value = piID;
		optionDataForm.txtOptionValue.value = sSelectedIDs;
		optionDataForm.txtOptionPromptSQL.value = psPrompts;
		optionDataForm.txtOption1000SepCols.value = frmBulkBooking.txt1000SepCols.value;
		refreshOptionData(); //should be in scope.
	}

	function openDialog(pDestination, pWidth, pHeight)
	{
		var dlgwinprops = "center:yes;" +
			"dialogHeight:" + pHeight + "px;" +
			"dialogWidth:" + pWidth + "px;" +
			"help:no;" +
			"resizable:yes;" +
			"scroll:yes;" +
			"status:no;";
		window.showModalDialog(pDestination, self, dlgwinprops);
	}
</script>



<script type="text/javascript">
	
	function tbBulkBooking_addhandlers() {
		OpenHR.addActiveXHandler("ssOleDBGridRecords", "KeyPress", ssOleDBGridRecords_KeyPress);
		OpenHR.addActiveXHandler("ssOleDBGridRecords", "RowColChange", ssOleDBGridRecords_RowColChange);
	}

	function ssOleDBGridRecords_KeyPress(iKeyAscii) {
		if ((iKeyAscii >= 32) && (iKeyAscii <= 255)) {
			var iLastTick;
			var sFind;
			var dtTicker = new Date();
			var iThisTick = new Number(dtTicker.getTime());
			if ($("#txtLastKeyFind").val().length > 0) {
				iLastTick = new Number($("#txtTicker").val());
			} else {
				iLastTick = new Number("0");
			}

			if (iThisTick > (iLastTick + 1500)) {
				sFind = String.fromCharCode(iKeyAscii);
			} else {
				sFind = $("#txtLastKeyFind").val() + String.fromCharCode(iKeyAscii);
			}

			$("#txtTicker").val(iThisTick);
			$("#txtLastKeyFind").val(sFind);

			var frmBulkBooking = document.getElementById("frmBulkBooking");
			frmBulkBooking.ssOleDBGridFindRecords.SelBookmarks.RemoveAll();

			locateRecord(sFind, false);
		}
	}

	function ssOleDBGridRecords_RowColChange() {
		tbrefreshControls();
	}
	
</script>

<script src="<%: Url.Content("~/Scripts/ctl_SetStyles.js") %>" type="text/javascript"></script>

<div <%=session("BodyTag")%>>
<form name="frmBulkBooking" action="tbBulkBooking_Submit" method="post" id="frmBulkBooking" style="TEXT-ALIGN: center">

<table align=center class="outline" cellPadding=5 cellSpacing=0 width=100% height=100%>
	<TR>
		<TD>
			<TABLE WIDTH="100%" height="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
				<TR>
					<TD align=center height=10 colspan=5>
						<h3 align="left" class="pageTitle">Bulk Booking</h3>
					</td>
				</tr>

<%
	Dim sErrorDescription = ""
	
	if session("TB_TBStatusPExists") then
		Response.Write("				<TR height=10>" & vbCrLf)
		Response.Write("					<TD width=20></TD>" & vbCrLf)
		Response.Write("					<TD colspan=3>" & vbCrLf)
		Response.Write("						<TABLE WIDTH=""100%"" class=""invisible"" CELLSPACING=0 CELLPADDING=0>" & vbCrLf)
		Response.Write("							<TR height=10>" & vbCrLf)
		Response.Write("								<TD  nowrap>Booking Status :</TD>" & vbCrLf)
		Response.Write("								<TD width=20>&nbsp;</TD>" & vbCrLf)
		Response.Write("								<TD>" & vbCrLf)
		Response.Write("									<SELECT id=selStatus name=selStatus class=""combo"">" & vbCrLf)
		Response.Write("										<OPTION value=B selected>Booked</OPTION>" & vbCrLf)
		Response.Write("										<OPTION value=P>Provisional</OPTION></SELECT>" & vbCrLf)
		Response.Write("								</TD>" & vbCrLf)
		Response.Write("								<TD style='width: 100%;'></TD>" & vbCrLf)
		Response.Write("								<TD ></TD>" & vbCrLf)
		Response.Write("							</TR>" & vbCrLf)
		Response.Write("						</TABLE>" & vbCrLf)
		Response.Write("					</TD>" & vbCrLf)
		Response.Write("					<TD width=20></TD>" & vbCrLf)
		Response.Write("				</TR>" & vbCrLf)
		Response.Write("				<TR>" & vbCrLf)
		Response.Write("				  <td height=10 colspan=5></td>" & vbCrLf)
		Response.Write("				</TR>" & vbCrLf)
	end if
%>		
  
				<tr> 
					<td rowspan=13 width="20">&nbsp;&nbsp;&nbsp;&nbsp;</td>    
					<td rowspan=13 width=100%>
<%
	' Get the employee find columns.
	Dim cmdFindRecords = CreateObject("ADODB.Command")
	cmdFindRecords.CommandText = "sp_ASRIntGetTBEmployeeColumns"
	cmdFindRecords.CommandType = 4 ' Stored Procedure
	cmdFindRecords.ActiveConnection = Session("databaseConnection")
	cmdFindRecords.CommandTimeout = 180

	Dim prmErrorMsg = cmdFindRecords.CreateParameter("errorMsg", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
	cmdFindRecords.Parameters.Append(prmErrorMsg)

	Dim prm1000SepCols = cmdFindRecords.CreateParameter("1000SepCols", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
	cmdFindRecords.Parameters.Append(prm1000SepCols)

	Err.Clear()
	Dim rstFindRecords = cmdFindRecords.Execute

	If (Err.Number <> 0) Then
		sErrorDescription = "The Employee table find columns could not be retrieved." & vbCrLf & formatError(Err.Description)
	End If

	if len(sErrorDescription) = 0 then
		' Instantiate and initialise the grid. 
		Response.Write("<OBJECT classid=""clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"" id=ssOleDBGridFindRecords name=ssOleDBGridFindRecords  codebase=""cabs/COAInt_Grid.cab#version=3,1,3,6"" style=""LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:400px;"">" & vbCrLf)
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

		Dim lngColCount = 0
		do while not rstFindRecords.EOF
			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Width"" VALUE=""3200"">" & vbCrLf)
	
			if rstFindRecords.fields("columnName").value = "ID" then
				Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Visible"" VALUE=""0"">" & vbCrLf)
			else
				Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Visible"" VALUE=""-1"">" & vbCrLf)
			end if
	
			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Columns.Count"" VALUE=""1"">" & vbCrLf)
			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Caption"" VALUE=""" & Replace(rstFindRecords.fields("columnName").value, "_", " ") & """>" & vbCrLf)
			Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Name"" VALUE=""" & rstFindRecords.fields("columnName").value & """>" & vbCrLf)
				
			if (rstFindRecords.fields("dataType").value = 131) or (rstFindRecords.fields("dataType").value = 3) then
				Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Alignment"" VALUE=""1"">" & vbCrLf)
			else
				Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Alignment"" VALUE=""0"">" & vbCrLf)
			end if
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
				
			if rstFindRecords.fields("dataType").value = -7 then
				' Find column is a logic column.
				Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Style"" VALUE=""2"">" & vbCrLf)
			else	
				Response.Write("	<PARAM NAME=""Columns(" & lngColCount & ").Style"" VALUE=""0"">" & vbCrLf)
			end if

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

			lngColCount = lngColCount + 1
			rstFindRecords.MoveNext
		loop
		
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

		' Release the ADO recordset object.
		rstFindRecords.close
		rstFindRecords = Nothing

		' NB. IMPORTANT ADO NOTE.
		' When calling a stored procedure which returns a recordset AND has output parameters
		' you need to close the recordset and set it to nothing before using the output parameters. 
		if len(cmdFindRecords.Parameters("errorMsg").Value) > 0 then
	    Session("ErrorTitle") = "Bulk Booking Page"
		  Session("ErrorText") = cmdFindRecords.Parameters("errorMsg").Value
			Response.Clear	  
			Response.Redirect("error.asp")
		else
			Response.Write("<INPUT type='hidden' id=txt1000SepCols name=txt1000SepCols value=""" & cmdFindRecords.Parameters("1000SepCols").Value & """>" & vbCrLf)
		end if
	end if
	
	' Release the ADO command object.
	cmdFindRecords = Nothing
%>
					</td>
					<td rowspan=13 width="20">&nbsp;&nbsp;&nbsp;&nbsp;</td>    
					<TD width=80 height=10>
						<input type="button" id=cmdAdd name=cmdAdd value="Add" style="WIDTH: 100px" width="100" class="btn"  
						    onclick="add()" 
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
					</TD>
					<td rowspan=13 width="20">&nbsp;&nbsp;&nbsp;&nbsp;</td>    
				</tr>
		
				<TR>
					<TD height=10></TD>
				</TR>
		
				<TR>
					<TD height=10>
						<input type="button" name=cmdAddFilter id=cmdAddFilter value="Filtered Add" style="WIDTH: 100px" width="100" class="btn"
						    onclick="filteredAdd()" 
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
					</TD>
				</TR>
		
				<TR>
					<TD height=10></TD>
				</TR>
		
				<TR>
					<TD height=10>
						<input type="button" name=cmdAddPicklist id=cmdAddPicklist value="Picklist Add" style="WIDTH: 100px" width="100" class="btn"
						    onclick="addPicklist()" 
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
					</TD>
				</TR>
		
				<TR>
					<TD height=10></TD>
				</TR>

				<TR>
					<TD height=10>
						<input type="button" name=cmdRemove value="Remove" style="WIDTH: 100px" width="100" class="btn"
						    onclick="remove()" 
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
					</TD>
				</TR>

				<TR>
					<TD height=10></TD>
				</TR>

				<TR>
					<TD height=10>
						<input type="button" name=cmdRemoveAll value="Remove All" style="WIDTH: 100px" width="100" class="btn"
						    onclick="removeAll()" 
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
					</TD>
				</TR>

				<TR>
					<TD></TD>
				</TR>

				<TR>
					<TD height=10>
						<input type="button" name=cmdOK value="OK" style="WIDTH: 100px" width="100" id=cmdOK class="btn"
						    onclick="ok()" 
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
					</TD>
				</TR>

				<TR>
					<TD height=10></TD>
				</TR>

				<TR>
					<TD height=10>
						<input type="button" name="cmdCancel" value="Cancel" style="WIDTH: 100px" width="100" class="btn"
						    onclick="cancel()" 
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
					</TD>
				</TR>     

				<TR>
					<TD height=10 colspan=5></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</form>

<INPUT type='hidden' id=txtTicker name=txtTicker value=0>
<INPUT type='hidden' id=txtLastKeyFind name=txtLastKeyFind >

<INPUT type='hidden' id=txtSelectionID name=txtSelectionID value=0>
<INPUT type='hidden' id=txtOptionRecordID name=txtOptionRecordID value=<%=session("optionRecordID")%>>


<FORM action="tbBulkBooking_Submit" method="post" id="frmGotoOption" name="frmGotoOption" style="visibility:hidden;display:none">
<%Html.RenderPartial("~/Views/Shared/gotoOption.ascx")%>
</FORM>

<FORM id="frmBookingSelection" name="frmBookingSelection" target="tbBulkBookingSelection" action="tbBulkBookingSelectionMain" method="post" style="visibility:hidden;display:none">
	<INPUT type="hidden" id="selectionType" name="selectionType">
</FORM>

</div>

<script type="text/javascript">
	tbBulkBooking_addhandlers();
	tbBulkBooking_onload();
</script>
