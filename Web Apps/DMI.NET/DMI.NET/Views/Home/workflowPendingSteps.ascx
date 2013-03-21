<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<%-- For other devs: Do not remove below line. --%>
<%="" %>
<%-- For other devs: Do not remove above line. --%>

<form id="frmSteps" name="frmSteps" style="visibility: hidden; display: none">
	<%On Error Resume Next

		Response.Expires = -1
	
		If (Session("fromMenu") = 0) And (Session("reset") = 1) Then
			' Reset the Workflow OutOfOffice flag.
			Dim cmdOutOfOffice = CreateObject("ADODB.Command")
			cmdOutOfOffice.CommandText = "spASRWorkflowOutOfOfficeSet"
			cmdOutOfOffice.CommandType = 4 ' Stored Procedure
			cmdOutOfOffice.ActiveConnection = Session("databaseConnection")

			Dim prmValue = cmdOutOfOffice.CreateParameter("value", 11, 1)	' 11=bit, 1=input
			cmdOutOfOffice.Parameters.Append(prmValue)
			prmValue.value = 0

			Err.Clear()
			cmdOutOfOffice.Execute()

			cmdOutOfOffice = Nothing

			Session("reset") = 0
		End If
	
		Dim fWorkflowGood = True
		Dim iStepCount As Integer = 0
		Dim sAddString As String
	
		Dim cmdDefSelRecords = CreateObject("ADODB.Command")
		cmdDefSelRecords.CommandText = "spASRIntCheckPendingWorkflowSteps"
		cmdDefSelRecords.CommandType = 4 ' Stored Procedure
		cmdDefSelRecords.ActiveConnection = Session("databaseConnection")

		Err.Clear()
		Dim rstDefSelRecords = cmdDefSelRecords.Execute

		If Err.Number <> 0 Then
			' Workflow not licensed or configured. Go to default page.
			fWorkflowGood = False
		Else
			Do Until rstDefSelRecords.eof
				If iStepCount = 0 Then
					' Add the <All> row.
					sAddString = CType(("0" & vbTab & "<All>" & vbTab), String)
	%>
	<input type='hidden' id="txtAddString_<%=iStepCount%>" name="txtAddString_<%=iStepCount%>" value="<%=sAddString%>">
	<%			
				
	End If

	iStepCount = iStepCount + 1
	sAddString = "0" & vbTab & _
		Replace(CType(rstDefSelRecords.Fields("description").Value, String), Chr(34), "&quot;") & vbTab & _
		Replace(CType(rstDefSelRecords.Fields("url").Value, String), Chr(34), "&quot;")
	%>
	<input type='hidden' id="Hidden1" name="txtAddString_<%=iStepCount%>" value="<%=sAddString%>">
	<%			
		rstDefSelRecords.movenext()
	Loop

	rstDefSelRecords.close()
	rstDefSelRecords = Nothing
End If
							
' Release the ADO command object.
cmdDefSelRecords = Nothing
	%>
	<input type='hidden' id="txtFromMenu" name="txtFromMenu" value="<%=Session("fromMenu")%>">
</form>

<script type="text/javascript">
	function workflowPendingSteps_window_onload() {
		var sControlName;
		var sControlPrefix;

		var frmDefSel = document.getElementById('frmDefSel');
		//$("#workframe").attr("data-framesource", "DEFSEL");

		var frmSteps = document.getElementById('frmSteps');

		<%If iStepCount > 0 Then%>

		setGridFont(frmDefSel.ssOleDBGridDefSelRecords);

		frmDefSel.ssOleDBGridDefSelRecords.focus();
		frmDefSel.cmdCancel.focus();

		var controlCollection = frmSteps.elements;
		if (controlCollection != null) {
			for (i = 0; i < controlCollection.length; i++) {

				sControlName = controlCollection.item(i).name;
				sControlPrefix = sControlName.substr(0, 13);

				if (sControlPrefix == "txtAddString_") {
					frmDefSel.ssOleDBGridDefSelRecords.AddItem(controlCollection.item(i).value);
				}
			}
		}

		if (frmDefSel.ssOleDBGridDefSelRecords.rows > 0) {
			// Need to refresh the grid before we movefirst.
			frmDefSel.ssOleDBGridDefSelRecords.refresh();
			// Select the top row.
			frmDefSel.ssOleDBGridDefSelRecords.MoveFirst();
			frmDefSel.ssOleDBGridDefSelRecords.SelBookmarks.Add(frmDefSel.ssOleDBGridDefSelRecords.Bookmark);
		}

		refreshControls();

		sizeColumnsToFitGrid(frmDefSel.ssOleDBGridDefSelRecords);
		<%Else
		If Session("fromMenu") = 0 Then%>
		//TODO
		//window.parent.frames("menuframe").openPersonnelRecEdit();
		<%End If
End If
%>
		//TODO
		//window.parent.frames("menuframe").refreshMenu();
		//TODO		
		//window.parent.document.all.item("workframeset").cols = "*, 0";	

		// Little dodge to get around a browser bug that
		// does not refresh the display on all controls.
		try {
			window.resizeBy(0, -1);
			window.resizeBy(0, 1);
			window.resizeBy(0, -1);
			window.resizeBy(0, 1);
		} catch (e) {
		}
	}
</script>

<script type="text/javascript">
	OpenHR.addActiveXHandler("ssOleDBGridDefSelRecords", "Change", ssOleDBGridDefSelRecords_Change);
	function ssOleDBGridDefSelRecords_Change() {
		RefreshGrid();
	}
</script>

<script type="text/javascript">
	OpenHR.addActiveXHandler("ssOleDBGridDefSelRecords", "KeyPress", ssOleDBGridDefSelRecords_KeyPress);
	function ssOleDBGridDefSelRecords_KeyPress(iKeyAscii) {

		//if ((iKeyAscii >= 32) && (iKeyAscii <= 255)) {	
		//	var dtTicker = new Date();
		//	var iThisTick = new Number(dtTicker.getTime());
		//	if (txtLastKeyFind.value.length > 0) {
		//		var iLastTick = new Number(txtTicker.value);
		//	}
		//	else {
		//		var iLastTick = new Number("0");
		//	}

		//	if (iThisTick > (iLastTick + 1500)) {
		//		var sFind = String.fromCharCode(iKeyAscii);
		//	}
		//	else {
		//		var sFind = txtLastKeyFind.value + String.fromCharCode(iKeyAscii);
		//	}

		//	txtTicker.value = iThisTick;
		//	txtLastKeyFind.value = sFind;

		//	locateRecord(sFind);
		//ToggleCurrentRow();
		//}
	}
</script>

<script type="text/javascript">
	function RefreshGrid() {
		var iLoop;
		var iRowIndex = frmDefSel.ssOleDBGridDefSelRecords.AddItemRowIndex(frmDefSel.ssOleDBGridDefSelRecords.Bookmark);
		var sRowTickValue;
		var fAllTicked = true;

		frmDefSel.ssOleDBGridDefSelRecords.Update();

		if (iRowIndex == 0) {
			// <All> row. Ensure all other rows match.
			var varBookmark = frmDefSel.ssOleDBGridDefSelRecords.AddItemBookmark(0);
			sRowTickValue = frmDefSel.ssOleDBGridDefSelRecords.Columns("TickBox").CellText(varBookmark);

			frmDefSel.ssOleDBGridDefSelRecords.MoveFirst();
			frmDefSel.ssOleDBGridDefSelRecords.MoveNext();

			for (iLoop = 1; iLoop < frmDefSel.ssOleDBGridDefSelRecords.Rows; iLoop++) {
				frmDefSel.ssOleDBGridDefSelRecords.Columns("TickBox").Text = sRowTickValue;
				frmDefSel.ssOleDBGridDefSelRecords.MoveNext();
			}
			frmDefSel.ssOleDBGridDefSelRecords.MoveFirst();
		}
		else {
			// Step row. Check if all step rows now have the same value.
			// If so, ensure the <All> row matches.

			for (iLoop = 1; iLoop < frmDefSel.ssOleDBGridDefSelRecords.Rows; iLoop++) {
				varBookmark = frmDefSel.ssOleDBGridDefSelRecords.AddItemBookmark(iLoop);
				sRowTickValue = frmDefSel.ssOleDBGridDefSelRecords.Columns("TickBox").CellText(varBookmark);

				if (sRowTickValue == "0") {
					fAllTicked = false;
				}
			}

			varBookmark = frmDefSel.ssOleDBGridDefSelRecords.Bookmark;

			if (fAllTicked == true) {

				frmDefSel.ssOleDBGridDefSelRecords.Bookmark = frmDefSel.ssOleDBGridDefSelRecords.AddItemBookmark(0);
				frmDefSel.ssOleDBGridDefSelRecords.Columns("TickBox").Text = "-1";
			}
			else {
				frmDefSel.ssOleDBGridDefSelRecords.Bookmark = frmDefSel.ssOleDBGridDefSelRecords.AddItemBookmark(0);
				frmDefSel.ssOleDBGridDefSelRecords.Columns("TickBox").Text = "0";
			}

			frmDefSel.ssOleDBGridDefSelRecords.Bookmark = varBookmark;
		}

		refreshControls();
	}

	function ToggleCurrentRow() {
		if (frmDefSel.ssOleDBGridDefSelRecords.Columns("TickBox").Text == "-1") {
			frmDefSel.ssOleDBGridDefSelRecords.Columns("TickBox").Text = "0";
		}
		else {
			frmDefSel.ssOleDBGridDefSelRecords.Columns("TickBox").Text = "-1";
		}
		RefreshGrid();
	}

</script>

<script type="text/javascript">
	function refreshControls() {
		var fSomeSelected;
		fSomeSelected = SomeSelected();
		button_disable(frmDefSel.cmdRun, (fSomeSelected == false));
	}

	function SomeSelected() {
		var varBookmark;
		var iLoop;
		frmDefSel.ssOleDBGridDefSelRecords.Update()

		for (iLoop = 1; iLoop < frmDefSel.ssOleDBGridDefSelRecords.Rows; iLoop++) {
			varBookmark = frmDefSel.ssOleDBGridDefSelRecords.AddItemBookmark(iLoop);
			if (frmDefSel.ssOleDBGridDefSelRecords.Columns("TickBox").CellText(varBookmark) == "-1") {
				return (true);
			}
		}

		return (false);
	}

	function pausecomp(millis) {
		var date = new Date();
		var curDate = null;
		do {
			curDate = new Date();
		}
		while (curDate - date < millis);
	}

	function spawnWindow(mypage, myname, w, h, scroll) {
		var newWin;
		var winl = (screen.availWidth - w) / 2;
		var wint = (screen.availHeight - h) / 2;

		var winprops = 'height=' + h + ',width=' + w + ',top=' + wint + ',left=' + winl + ',scrollbars=' + scroll + ',resizable';

		try {
			newWin = window.open(mypage, myname, winprops);

			if (parseInt(navigator.appVersion) >= 4) {
				try {
					pausecomp(300);
					newWin.focus();
				}
				catch (e) { }
			}
		}
		catch (e) {
			try {
				newWin.close();
			}
			catch (e) { }

			spawnWindow(mypage, myname, w, h, scroll)
		}
	}

	function setrun() {
		var varBookmark;
		var sForm;
		var iSelectedCount = 0;
		var sMessage;
		var iLoop;
		var frmRefresh = OpenHR.getForm("refreshframe", "frmRefresh");
		var frmDefSel = document.getElementById('frmDefSel');

		//window.parent.frames("refreshframe").document.forms("frmRefresh").submit();
		OpenHR.submitForm(frmRefresh);
		try {
			for (iLoop = 1; iLoop < frmDefSel.ssOleDBGridDefSelRecords.Rows; iLoop++) {
				varBookmark = frmDefSel.ssOleDBGridDefSelRecords.AddItemBookmark(iLoop);

				if (frmDefSel.ssOleDBGridDefSelRecords.Columns("TickBox").CellText(varBookmark) == "-1") {
					sForm = frmDefSel.ssOleDBGridDefSelRecords.Columns("URL").CellText(varBookmark);
					spawnWindow(sForm, "_blank", screen.availWidth, screen.availHeight, 'yes');

					iSelectedCount = iSelectedCount + 1;
				}
			}

			if (iSelectedCount == 0) {
				sMessage = "You must select a workflow step to run";
				//window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sMessage,48,"OpenHR Intranet");
				OpenHR.messageBox(sMessage, 48, "OpenHR Intranet");

			}
			else {
				<%
	If Session("fromMenu") = 0 Then
					%>
				menu_autoLoadPage("workflowPendingSteps", true);

				<%
Else
					%>
				menu_autoLoadPage("workflowPendingSteps", false);
				<%
End If
				%>
			}
		}
		catch (e) {
			sMessage = "Error opening workflow forms : " + e.description;
			OpenHR.messageBox(sMessage, 48, "OpenHR Intranet");
		}
	}

	function setcancel() {
		// Goto self-service recedit page (if Self-service user at login)
		// Otherwise load the default page.

		<%If Session("fromMenu") = 0 Then%>
		window.parent.frames("menuframe").openPersonnelRecEdit();
		<%Else%>
		//window.location = "default";
		window.location = "main";
		<%End If%>
	}

	function setrefresh() {

		OpenHR.submitForm("frmRefresh");
			<%If Session("fromMenu") = 0 Then%>
		menu_autoLoadPage("workflowPendingSteps", true);
			<%Else%>
		menu_autoLoadPage("workflowPendingSteps", false);
			<%End If%>
	}

	//TODO - parts of this next function are still todo

	function currentWorkFramePage() {
		//// Return the current page in the workframeset.
		//var sCols = window.parent.document.all.item("workframeset").cols;

		//var re = / /gi;
		//sCols = sCols.replace(re, "");
		//sCols = sCols.substr(0, 1);

		//// Work frame is in view.
		////var sCurrentPage = window.parent.frames("workframe").document.location;
		//var sCurrentPage = OpenHR.getForm("workframe");
		//sCurrentPage = sCurrentPage.toString();

		//if (sCurrentPage.lastIndexOf("/") > 0) {
		//	sCurrentPage = sCurrentPage.substr(sCurrentPage.lastIndexOf("/") + 1);
		//}

		//if (sCurrentPage.indexOf(".") > 0) {
		//	sCurrentPage = sCurrentPage.substr(0, sCurrentPage.indexOf("."));
		//}

		//re = / /gi;
		//sCurrentPage = sCurrentPage.replace(re, "");
		//sCurrentPage = sCurrentPage.toUpperCase();

		//return(sCurrentPage);	
	}

	function sizeColumnsToFitGrid(pctlGrid) {
		var iLoop;
		var iVisibleColumnCount;
		var iVisibleCheckboxCount;
		var iNewColWidth;
		var iLastVisibleColumn;
		var iUsedWidth;
		var iUsableWidth;
		var iMinWidth = 100;
		var fScrollBarVisible;
		var iCheckboxWidth = 100;

		iVisibleCheckboxCount = 0;
		iVisibleColumnCount = 0;
		iLastVisibleColumn = 0;
		iUsedWidth = 0;
		for (iLoop = 0; iLoop < pctlGrid.Columns.Count; iLoop++) {
			if (pctlGrid.Columns.Item(iLoop).Visible == true) {
				if (pctlGrid.Columns.Item(iLoop).Style == 2) {
					iVisibleCheckboxCount = iVisibleCheckboxCount + 1;
				}

				iVisibleColumnCount = iVisibleColumnCount + 1;
				iLastVisibleColumn = iLoop;
			}
		}

		if (iVisibleColumnCount > 0) {
			fScrollBarVisible = (pctlGrid.Rows > pctlGrid.VisibleRows);
			if (fScrollBarVisible == true) {
				//NPG20090403 Fault 13516
				//iUsableWidth = pctlGrid.style.pixelWidth - 20;
				iUsableWidth = findTable.clientWidth - 20;
			} else {
				//NPG20090403 Fault 13516
				//iUsableWidth = pctlGrid.style.pixelWidth;
				iUsableWidth = findTable.clientWidth;
			}

			iNewColWidth = (iUsableWidth - (iVisibleCheckboxCount * iCheckboxWidth)) / (iVisibleColumnCount - iVisibleCheckboxCount);
			if (iNewColWidth < iMinWidth) {
				iNewColWidth = iMinWidth;
			}

			for (iLoop = 0; iLoop < iLastVisibleColumn; iLoop++) {
				if (pctlGrid.Columns.Item(iLoop).Visible == true) {
					if (pctlGrid.Columns.Item(iLoop).Style == 2) {
						pctlGrid.Columns(iLoop).Width = iCheckboxWidth;
					} else {
						pctlGrid.Columns.Item(iLoop).Width = iNewColWidth;
					}
					iUsedWidth = iUsedWidth + pctlGrid.Columns.Item(iLoop).Width;
				}
			}

			iNewColWidth = iUsableWidth - iUsedWidth - 2;
			if (iNewColWidth < iMinWidth) {
				iNewColWidth = iMinWidth;
			}
			pctlGrid.Columns.Item(iLastVisibleColumn).Width = iNewColWidth;
		}
	}
</script>

<div <%=session("BodyTag")%>>

	<form name="frmDefSel" method="post" id="frmDefSel">

		<%If (fWorkflowGood = True) Or (Session("fromMenu") = 1) Then%>
		<%	If iStepCount > 0 Then%>
		<table align="center" class="outline" cellpadding="5" cellspacing="0" height="100%" width="100%">
			<tr>
				<td>
					<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
						<tr>
							<td colspan="5" align="center" height="10">
								<h3>Pending Workflow Steps
								</h3>
							</td>
						</tr>

						<tr>
							<td width="20">&nbsp;&nbsp;&nbsp;&nbsp;</td>
							<td width="100%">
								<table height="100%" width="100%" class="invisible" cellspacing="0" cellpadding="0" id="findTable">
									<tr>
										<td width="100%">
											<object classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"
												id="ssOleDBGridDefSelRecords"
												name="ssOleDBGridDefselRecords"
												codebase="cabs/COAInt_Grid.cab#version=3,1,3,6"
												style="LEFT: 0px; TOP: 0px; WIDTH: 100%; HEIGHT: 400px">
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
												<param name="Col.Count" value="3">
												<param name="stylesets.count" value="0">

												<param name="TagVariant" value="EMPTY">
												<param name="UseGroups" value="0">
												<param name="HeadFont3D" value="0">
												<param name="Font3D" value="0">
												<param name="DividerType" value="3">
												<param name="DividerStyle" value="1">
												<param name="DefColWidth" value="0">
												<param name="BevelColorScheme" value="2">
												<param name="BevelColorFrame" value="0">
												<param name="BevelColorHighlight" value="0">
												<param name="BevelColorShadow" value="0">
												<param name="BevelColorFace" value="0">
												<param name="CheckBox3D" value="0">
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
												<param name="RowNavigation" value="1">
												<param name="CellNavigation" value="0">
												<param name="MaxSelectedRows" value="1">
												<param name="HeadStyleSet" value="">
												<param name="StyleSet" value="">
												<param name="ForeColorEven" value="0">
												<param name="ForeColorOdd" value="0">
												<param name="BackColorEven" value="0">
												<param name="BackColorOdd" value="0">
												<param name="Levels" value="1">
												<param name="RowHeight" value="503">
												<param name="ExtraHeight" value="0">
												<param name="ActiveRowStyleSet" value="">
												<param name="CaptionAlignment" value="2">
												<param name="SplitterPos" value="0">
												<param name="SplitterVisible" value="0">
												<param name="Columns.Count" value="3">

												<param name="Columns(0).Width" value="1000">
												<param name="Columns(0).Visible" value="-1">
												<param name="Columns(0).Columns.Count" value="1">
												<param name="Columns(0).Caption" value="">
												<param name="Columns(0).Name" value="TickBox">
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
												<param name="Columns(0).Style" value="2">
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

												<param name="Columns(1).Width" value="1000">
												<param name="Columns(1).Visible" value="-1">
												<param name="Columns(1).Columns.Count" value="1">
												<param name="Columns(1).Caption" value="">
												<param name="Columns(1).Name" value="Description">
												<param name="Columns(1).Alignment" value="0">
												<param name="Columns(1).CaptionAlignment" value="3">
												<param name="Columns(1).Bound" value="0">
												<param name="Columns(1).AllowSizing" value="1">
												<param name="Columns(1).DataField" value="Column 0">
												<param name="Columns(1).DataType" value="8">
												<param name="Columns(1).Level" value="0">
												<param name="Columns(1).NumberFormat" value="">
												<param name="Columns(1).Case" value="0">
												<param name="Columns(1).FieldLen" value="4096">
												<param name="Columns(1).VertScrollBar" value="0">
												<param name="Columns(1).Locked" value="-1">
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
												<param name="Columns(2).Caption" value="">
												<param name="Columns(2).Name" value="URL">
												<param name="Columns(2).Alignment" value="0">
												<param name="Columns(2).CaptionAlignment" value="3">
												<param name="Columns(2).Bound" value="0">
												<param name="Columns(2).AllowSizing" value="1">
												<param name="Columns(2).DataField" value="Column 0">
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

												<param name="UseDefaults" value="-1">
												<param name="TabNavigation" value="1">
												<param name="_ExtentX" value="17330">
												<param name="_ExtentY" value="1323">
												<param name="_StockProps" value="79">
												<param name="Caption" value="">
												<param name="ForeColor" value="0">
												<param name="BackColor" value="0">
												<param name="Enabled" value="-1">
												<param name="DataMember" value="">
												<param name="Row.Count" value="0">
											</object>
										</td>
									</tr>
								</table>
							</td>

							<td width="20">&nbsp;&nbsp;&nbsp;&nbsp;</td>

							<td width="80">
								<table height="100%" class="invisible" cellspacing="0" cellpadding="0">
									<tr>
										<td>
											<input type="button" name="cmdRefresh" value="Refresh" style="WIDTH: 80px" width="80" id="cmdRefresh" class="btn"
												onclick="setrefresh();"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
									</tr>
									<tr height="100%">
										<td></td>
									</tr>
									<tr>
										<td>
											<input type="button" name="cmdRun" value="Run" style="WIDTH: 80px" width="80" id="cmdRun" class="btn"
												onclick="setrun();"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
									</tr>
									<tr height="10">
										<td></td>
									</tr>
									<tr>
										<td>
											<input type="button" name="cmdCancel" value="Cancel" style="WIDTH: 80px" width="80" class="btn"
												onclick="setcancel()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
									</tr>
								</table>
							</td>
							<td width="20">&nbsp;&nbsp;&nbsp;&nbsp;</td>
						</tr>
						<tr>
							<td colspan="5" align="center" height="10"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<%							
		Else
			If (Session("fromMenu") = 1) Then
				Dim sMessage As String
				If fWorkflowGood = True Then
					' Display message saying no pending steps.
					sMessage = "No pending workflow steps"
				Else
					' Display error message.
					sMessage = "Error getting the pending workflow steps"
				End If
		%>
		<table align="center" class="outline" cellpadding="5" cellspacing="0">
			<tr>
				<td width="20"></td>
				<td>
					<table class="invisible" cellspacing="0" cellpadding="0">
						<tr>
							<td height="10"></td>
						</tr>

						<tr>
							<td align="center">
								<h3>Pending Workflow Steps</h3>
							</td>
						</tr>

						<tr>
							<td align="center">
								<%=sMessage%>
							</td>
						</tr>

						<tr>
							<td height="20"></td>
						</tr>

						<tr>
							<td height="10" align="center">
								<input id="cmdOK" name="cmdOK" type="button" class="btn" value="OK" style="WIDTH: 75px" width="75"
									onclick="setcancel()"
									onmouseover="try{button_onMouseOver(this);}catch(e){}"
									onmouseout="try{button_onMouseOut(this);}catch(e){}"
									onfocus="try{button_onFocus(this);}catch(e){}"
									onblur="try{button_onBlur(this);}catch(e){}" />
							</td>
						</tr>

						<tr>
							<td height="10"></td>
						</tr>
					</table>
				</td>
				<td width="20"></td>
			</tr>
		</table>
		<%			
		End If
	End If
End If
		%>
	</form>

	<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
		<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
	</form>
</div>

<script type="text/javascript">
	workflowPendingSteps_window_onload();
</script>


<%--		
		OpenHR.addActiveXHandler("<SSGRIDCONTROL>", "RowColChange", <SSGRIDCONTROL>_RowColChange);
		OpenHR.addActiveXHandler("<SSGRIDCONTROL>", "DblClick", <SSGRIDCONTROL>_DblClick);
		OpenHR.addActiveXHandler("<SSGRIDCONTROL>", "KeyPress", <SSGRIDCONTROL>_KeyPress);
		OpenHR.addActiveXHandler("<SSGRIDCONTROL>", "SelChange", <SSGRIDCONTROL>_SelChange);
		OpenHR.addActiveXHandler("<SSGRIDCONTROL>", "BeforeUpdate", <SSGRIDCONTROL>_BeforeUpdate);
		OpenHR.addActiveXHandler("<SSGRIDCONTROL>", "AfterInsert", <SSGRIDCONTROL>_AfterInsert);
		OpenHR.addActiveXHandler("<SSGRIDCONTROL>", "RowLoaded", <SSGRIDCONTROL>_RowLoaded);
		OpenHR.addActiveXHandler("<SSGRIDCONTROL>", "Change", <SSGRIDCONTROL>_Change);
		OpenHR.addActiveXHandler("<SSGRIDCONTROL>", "KeyUp", <SSGRIDCONTROL>_KeyUp);
		OpenHR.addActiveXHandler("<SSGRIDCONTROL>", "Click", <SSGRIDCONTROL>_Click);
		OpenHR.addActiveXHandler("<SSGRIDCONTROL>", "ComboCloseUp", <SSGRIDCONTROL>_ComboCloseUp);
		OpenHR.addActiveXHandler("<SSGRIDCONTROL>", "GotFocus", <SSGRIDCONTROL>_GotFocus);
--%>