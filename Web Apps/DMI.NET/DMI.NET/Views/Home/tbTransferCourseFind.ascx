﻿<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data" %>

<script src="<%: Url.Content("~/Scripts/ctl_SetFont.js") %>" type="text/javascript"></script>

<script type="text/javascript">
	function tbTransferCourseFind_onload() {
		var fOK;
		fOK = true;
		var frmtbFindForm = document.getElementById("frmtbFindForm");

		var sErrMsg = frmtbFindForm.txtErrorDescription.value;
		if (sErrMsg.length > 0) {
			fOK = false;
			OpenHR.messageBox(sErrMsg);
			window.parent.location.replace("login");
		}

		if (fOK == true) {
			sErrMsg = frmtbFindForm.txtFailureDescription.value;
			if (sErrMsg.length > 0) {
				fOK = false;
				OpenHR.messageBox(sErrMsg);
				tbCancel();
			}
		}

		if (fOK == true) {
			if (frmtbFindForm.selectView.length == 0) {
				fOK = false;
				OpenHR.messageBox("You do not have permission to read the link table.");
				tbCancel();
			}
		}

		if (fOK == true) {
			if (frmtbFindForm.selectOrder.length == 0) {
				fOK = false;
				OpenHR.messageBox("You do not have permission to use any of the link table orders.");
				tbCancel();
			}
		}

		if (fOK == true) {
			setGridFont(frmtbFindForm.ssOleDBGridRecords);

			// Expand the option frame and hide the work frame.
			//window.parent.document.all.item("workframeset").cols = "0, *";
			$("#optionframe").attr("data-framesource", "TBTRANSFERCOURSEFIND");
			$("#workframe").hide();
			$("#optionframe").show();
			
			// Set focus onto one of the form controls. 
			// NB. This needs to be done before making any reference to the grid
			frmtbFindForm.cmdCancel.focus();

			// Fault 3503
			//TODO: window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "hidden";

			// Get the optionData.asp to get the link find records.
			var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
			optionDataForm.txtOptionAction.value = "LOADTRANSFERCOURSE";
			optionDataForm.txtOptionTableID.value = frmtbFindForm.txtOptionLinkTableID.value;
			optionDataForm.txtOptionViewID.value = frmtbFindForm.selectView.options[frmtbFindForm.selectView.selectedIndex].value;
			optionDataForm.txtOptionOrderID.value = frmtbFindForm.selectOrder.options[frmtbFindForm.selectOrder.selectedIndex].value;
			optionDataForm.txtOptionCourseTitle.value = frmtbFindForm.txtOptionCourseTitle.value;
			optionDataForm.txtOptionRecordID.value = frmtbFindForm.txtOptionRecordID.value;
			optionDataForm.txtOptionFirstRecPos.value = 1;
			optionDataForm.txtOptionCurrentRecCount.value = 0;
			optionDataForm.txtOptionPageAction.value = "LOAD";

			refreshOptionData();	//should be in scope
		}
	}
</script>

<script type="text/javascript">

	function tbSelect() {
		
		// Fault 3503
		//TODO: window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "visible";
		var frmtbFindForm = document.getElementById("frmtbFindForm");
		var frmGotoOption = document.getElementById("frmGotoOption");
		
		frmGotoOption.txtGotoOptionAction.value = "SELECTTRANSFERCOURSE";
		frmGotoOption.txtGotoOptionRecordID.value = frmtbFindForm.txtOptionRecordID.value;
		frmGotoOption.txtGotoOptionLinkRecordID.value = tbselectedRecordID();
		frmGotoOption.txtGotoOptionPage.value = "emptyoption";
		OpenHR.submitForm(frmGotoOption);
	}

	function tbCancel()
	{  
		// Fault 3503
		//TODO: window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "visible";
		$("#optionframe").hide();
		$("#workframe").show();
		
		var frmGotoOption = document.getElementById("frmGotoOption");
		
		frmGotoOption.txtGotoOptionAction.value = "SELECTTRANSFERCOURSE";
		frmGotoOption.txtGotoOptionLinkRecordID.value = 0;
		frmGotoOption.txtGotoOptionPage.value = "emptyoption";
		OpenHR.submitForm(frmGotoOption);
	}

	/* Return the ID of the record selected in the find form. */
	function tbselectedRecordID() {
		var iRecordID;
		var iIndex;
		var iIDColumnIndex;
		var sColumnName;
		var frmtbFindForm = document.getElementById("frmtbFindForm");
		
		iRecordID = 0;
		iIDColumnIndex = 0;
	
		if (frmtbFindForm.ssOleDBGridRecords.SelBookmarks.Count > 0) {
			for (iIndex = 0; iIndex < frmtbFindForm.ssOleDBGridRecords.Cols; iIndex++) {
				sColumnName = frmtbFindForm.ssOleDBGridRecords.Columns(iIndex).Name;
				if (sColumnName.toUpperCase() == "ID") {
					iIDColumnIndex = iIndex;
					break;
				}
			}
		
			iRecordID = frmtbFindForm.ssOleDBGridRecords.Columns(iIDColumnIndex).Value;
		}

		return(iRecordID);
	}

	function tbrefreshControls() {
		var frmtbFindForm = document.getElementById("frmtbFindForm");
		
		if (frmtbFindForm.ssOleDBGridRecords.rows > 0) {
			if (frmtbFindForm.ssOleDBGridRecords.SelBookmarks.Count > 0) {
				button_disable(frmtbFindForm.cmdSelect, false);
			}
			else {
				button_disable(frmtbFindForm.cmdSelect, true);
			}
		}
		else {
			button_disable(frmtbFindForm.cmdSelect, true);
		}
	
		if (frmtbFindForm.selectOrder.length <= 1) {
			combo_disable(frmtbFindForm.selectOrder, true);
			button_disable(frmtbFindForm.btnGoOrder, true);
		}

		if (frmtbFindForm.selectView.length <= 1) {
			combo_disable(frmtbFindForm.selectView, true);
			button_disable(frmtbFindForm.btnGoView, true);
		}
	}

	function tbgoView() {
		// Get the optionData.asp to get the link find records.
		var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
		var frmtbFindForm = document.getElementById("frmtbFindForm");
		
		optionDataForm.txtOptionAction.value = "LOADTRANSFERCOURSE";
		optionDataForm.txtOptionTableID.value = frmtbFindForm.txtOptionLinkTableID.value;
		optionDataForm.txtOptionViewID.value = frmtbFindForm.selectView.options[frmtbFindForm.selectView.selectedIndex].value;
		optionDataForm.txtOptionOrderID.value = frmtbFindForm.selectOrder.options[frmtbFindForm.selectOrder.selectedIndex].value;
		optionDataForm.txtOptionCourseTitle.value = frmtbFindForm.txtOptionCourseTitle.value;
		optionDataForm.txtOptionRecordID.value = frmtbFindForm.txtOptionRecordID.value;
		optionDataForm.txtOptionFirstRecPos.value = 1;
		optionDataForm.txtOptionCurrentRecCount.value = 0;

		frmtbFindForm.txtOptionLinkViewID.value = optionDataForm.txtOptionViewID.value;
		frmtbFindForm.txtOptionLinkOrderID.value = optionDataForm.txtOptionOrderID.value;

		refreshOptionData();
	}

	function tbgoOrder() {
		// Get the optionData.asp to get the link find records.
		var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
		var frmtbFindForm = document.getElementById("frmtbFindForm");
		
		optionDataForm.txtOptionAction.value = "LOADTRANSFERCOURSE";
		optionDataForm.txtOptionTableID.value = frmtbFindForm.txtOptionLinkTableID.value;
		optionDataForm.txtOptionViewID.value = frmtbFindForm.selectView.options[frmtbFindForm.selectView.selectedIndex].value;
		optionDataForm.txtOptionOrderID.value = frmtbFindForm.selectOrder.options[frmtbFindForm.selectOrder.selectedIndex].value;
		optionDataForm.txtOptionCourseTitle.value = frmtbFindForm.txtOptionCourseTitle.value;
		optionDataForm.txtOptionRecordID.value = frmtbFindForm.txtOptionRecordID.value;
		optionDataForm.txtOptionFirstRecPos.value = 1;
		optionDataForm.txtOptionCurrentRecCount.value = 0;

		frmtbFindForm.txtOptionLinkViewID.value = optionDataForm.txtOptionViewID.value;
		frmtbFindForm.txtOptionLinkOrderID.value = optionDataForm.txtOptionOrderID.value;

		refreshOptionData();
	}

	function tbselectedOrderID() {
		var frmtbFindForm = document.getElementById("frmtbFindForm");
		return frmtbFindForm.selectOrder.options[frmtbFindForm.selectOrder.selectedIndex].value;
	}

	function tbselectedViewID() {
		var frmtbFindForm = document.getElementById("frmtbFindForm");
		return frmtbFindForm.selectView.options[frmtbFindForm.selectView.selectedIndex].value;
	}

	function tblocateRecord(psFileName) {
		var fFound;
		var frmtbFindForm = document.getElementById("frmtbFindForm");
		
		fFound = false;
	
		frmtbFindForm.ssOleDBGridRecords.redraw = false;

		frmtbFindForm.ssOleDBGridRecords.MoveLast();
		frmtbFindForm.ssOleDBGridRecords.MoveFirst();

		for (var iIndex = 1; iIndex <= frmtbFindForm.ssOleDBGridRecords.rows; iIndex++) {	
			var sGridValue = new String(frmtbFindForm.ssOleDBGridRecords.Columns(0).value);
			sGridValue = sGridValue.substr(0, psFileName.length).toUpperCase();
			if (sGridValue == psFileName.toUpperCase()) {
				frmtbFindForm.ssOleDBGridRecords.SelBookmarks.Add(frmtbFindForm.ssOleDBGridRecords.Bookmark);
				fFound = true;
				break;
			}

			if (iIndex < frmtbFindForm.ssOleDBGridRecords.rows) {
				frmtbFindForm.ssOleDBGridRecords.MoveNext();
			}
			else {
				break;
			}
		}

		if ((fFound == false) && (frmtbFindForm.ssOleDBGridRecords.rows > 0)) {
			// Select the top row.
			frmtbFindForm.ssOleDBGridRecords.MoveFirst();
			frmtbFindForm.ssOleDBGridRecords.SelBookmarks.Add(frmtbFindForm.ssOleDBGridRecords.Bookmark);
		}

		frmtbFindForm.ssOleDBGridRecords.redraw = true;
	}
</script>


<script type="text/javascript">
	function tbAddFromWaitingListFind_addhandlers() {
		OpenHR.addActiveXHandler("ssOleDBGridRecords", "dblClick", "tbssOleDBGridRecords_dblClick()");
		OpenHR.addActiveXHandler("ssOleDBGridRecords", "click", "tbssOleDBGridRecords_click()");
		OpenHR.addActiveXHandler("ssOleDBGridRecords", "KeyPress", "tbssOleDBGridRecords_KeyPress()");
	}

	function tbssOleDBGridRecords_dblClick() {
		tbSelect();
	}

	function tbssOleDBGridRecords_click() {
		refreshControls();
	}

	function tbssOleDBGridRecords_KeyPress(iKeyAscii) {
		var iLastTick;
		var sFind;
		if ((iKeyAscii >= 32) && (iKeyAscii <= 255)) {
			var dtTicker = new Date();
			var iThisTick = new Number(dtTicker.getTime());
			if ($("#txtLastKeyFind").val().length > 0) {
				iLastTick = new Number($("#txtTicker").val());
			}
			else {
				iLastTick = new Number("0");
			}

			if (iThisTick > (iLastTick + 1500)) {
				sFind = String.fromCharCode(iKeyAscii);
			}
			else {
				sFind = $("#txtLastKeyFind").val() + String.fromCharCode(iKeyAscii);
			}

			$("#txtTicker").val(iThisTick);
			$("#txtLastKeyFind").val(sFind);

			locateRecord(sFind);
		}
	}
</script>

<script src="<%: Url.Content("~/Scripts/ctl_SetStyles.js") %>" type="text/javascript"></script>

<div <%=session("BodyTag")%>>

<FORM action="" method="POST" id="frmtbFindForm" name="frmtbFindForm">

<table align=center class="outline" cellPadding=5 cellSpacing=0 width=100% height=100%>
	<TR>
		<TD>
			<TABLE WIDTH="100%" height="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
				<TR>
					<TD height=10 colspan=3></td>
				</tr>
				<TR>
					<TD align=center height=10 colspan=3>
						<h3 align=center>Find Course Record</h3>
					</td>
				</tr>
				<tr height=10>
					<td width=20></td>
					<td>
						<table WIDTH=100% class="invisible" CELLSPACING="0" CELLPADDING="0">
							<TR>
								<TD width=40>
									View :
								</TD>
								<TD width=10>
									&nbsp;
								</TD>
								<TD width=175>
									<SELECT id=selectView name=selectView class="combo" style="HEIGHT: 22px; WIDTH: 200px">
<%
	on error resume next
	Dim sErrorDescription = "", sFailureDescription = ""

	Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
	
	if (len(sErrorDescription) = 0) and (len(sFailureDescription) = 0) then
		' Get the view records.
		Dim cmdViewRecords = CreateObject("ADODB.Command")
		cmdViewRecords.CommandText = "sp_ASRIntGetLinkViews"
		cmdViewRecords.CommandType = 4 ' Stored Procedure
		cmdViewRecords.ActiveConnection = Session("databaseConnection")

		Dim prmTableID = cmdViewRecords.CreateParameter("tableID", 3, 1)
		cmdViewRecords.Parameters.Append(prmTableID)
		prmTableID.value = cleanNumeric(session("optionLinkTableID"))

		Dim prmDfltOrderID = cmdViewRecords.CreateParameter("dfltOrderID", 3, 2) ' 11=integer, 2=output
		cmdViewRecords.Parameters.Append(prmDfltOrderID)

		Err.Clear()
		Dim rstViewRecords = cmdViewRecords.Execute

		If (Err.Number <> 0) Then
			sErrorDescription = "The Course views could not be retrieved." & vbCrLf & FormatError(Err.Description)
		End If

		if (len(sErrorDescription) = 0) and (len(sFailureDescription) = 0) then
			do while not rstViewRecords.EOF
				Response.Write("						<OPTION value=" & rstViewRecords.Fields(0).Value)
				if rstViewRecords.Fields(0).Value = session("optionLinkViewID") then
					Response.Write(" SELECTED")
				end if

				if rstViewRecords.Fields(0).Value = 0 then
					Response.Write(">" & Replace(rstViewRecords.Fields(1).Value, "_", " ") & "</OPTION>" & vbCrLf)
				Else
					Response.Write(">'" & Replace(rstViewRecords.Fields(1).Value, "_", " ") & "' view</OPTION>" & vbCrLf)
				end if

				rstViewRecords.MoveNext
			loop
			
			if (rstViewRecords.EOF and rstViewRecords.BOF) then
				sFailureDescription = "You do not have permission to read the Course table."
			end if
		
			' Release the ADO recordset object.
			rstViewRecords.close
			rstViewRecords = Nothing
	
			' NB. IMPORTANT ADO NOTE.
			' When calling a stored procedure which returns a recordset AND has output parameters
			' you need to close the recordset and set it to nothing before using the output parameters. 
			if session("optionLinkOrderID") <= 0 then
				session("optionLinkOrderID") = cmdViewRecords.Parameters("dfltOrderID").Value
			end if
		end if

		' Release the ADO command object.
		cmdViewRecords = Nothing
	end if
%>
									</SELECT>						
								</TD>
								<TD width=10>
									<INPUT type="button" value="Go" id=btnGoView name=btnGoView class="btn"
											onclick="tbgoView()" />
								</TD>
								<TD>
									&nbsp;
								</TD>
								<TD width=40>
									Order :
								</TD>
								<TD width=10>
									&nbsp;
								</TD>
								<TD width=175>
									<SELECT id=selectOrder name=selectOrder class="combo" style="HEIGHT: 22px; WIDTH: 200px">
<%
	If (Len(sErrorDescription) = 0) And (Len(sFailureDescription) = 0) Then
		
		Dim rstOrderRecords = objDatabase.GetTableOrders(CInt(CleanNumeric(Session("optionLinkTableID"))), 0)
		For Each objRow As DataRow In rstOrderRecords.Rows
			Response.Write("						<option value=" & objRow(1))
			If objRow(1) = CInt(Session("optionLinkOrderID")) Then
				Response.Write(" SELECTED")
			End If
			Response.Write(">" & Replace(objRow(0).ToString(), "_", " ") & "</option>" & vbCrLf)
		Next
		
	End If
%>
									</SELECT>
								</TD>
								<TD width=10>
									<INPUT type="button" value="Go" id=btnGoOrder name=btnGoOrder class="btn"
											onclick="tbgoOrder()" />
								</TD>
							</TR>
						</table>
					</td>
					<td width=20></td>
				</tr>
				<TR>
					<TD height=10 colspan=3></td>
				</tr>
				<TR>
					<td></td>
					<TD>
						<OBJECT classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" id=ssOleDBGridRecords name=ssOleDBGridRecords  codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:100%">
							<PARAM NAME="ScrollBars" VALUE="4">
							<PARAM NAME="_Version" VALUE="196617">
							<PARAM NAME="DataMode" VALUE="2">							
							<PARAM NAME="Cols" VALUE="0">
							<PARAM NAME="Rows" VALUE="0">
							<PARAM NAME="BorderStyle" VALUE="1">
							<PARAM NAME="RecordSelectors" VALUE="0">
							<PARAM NAME="GroupHeaders" VALUE="0">
							<PARAM NAME="ColumnHeaders" VALUE="-1">
							<PARAM NAME="GroupHeadLines" VALUE="1">
							<PARAM NAME="HeadLines" VALUE="1">
							<PARAM NAME="FieldDelimiter" VALUE="(None)">
							<PARAM NAME="FieldSeparator" VALUE="(Tab)">
							<PARAM NAME="Col.Count" VALUE="0">
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
							<PARAM NAME="AllowColumnSizing" VALUE="-1">
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
							<PARAM NAME="Columns.Count" VALUE="0">
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
					<td></td>
				</TR>
				<TR>
					<TD height=10 colspan=3></td>
				</tr>
				<tr>
					<td></td>
					<td height="10">
						<table WIDTH=100% BORDER=0 CELLSPACING="0" CELLPADDING="0">
							<TR>
								<TD colspan=4>
								</TD>
							</TR>
							<tr>	
								<td>
								</td>
								<td width=10>
									<input id="cmdSelect" name="cmdSelect" type="button" class="btn" value="Select" style="WIDTH: 75px" width="75" 
											onclick="tbSelect()" />
								</td>
								<td width=40>
								</td>
								<td width=10>
									<input id="cmdCancel" name="cmdCancel" type="button" class="btn" value="Cancel" style="WIDTH: 75px" width="75" 
											onclick="tbCancel()" />
								</td>
							</tr>			
						</table>
					</td>
					<td></td>
				</tr>
				<TR>
					<TD height=10 colspan=3></td>
				</tr>
			</TABLE>
		</td>
	</tr>
</TABLE>
<%
	Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtFailureDescription name=txtFailureDescription value=""" & sFailureDescription & """>" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtOptionLinkTableID name=txtOptionLinkTableID value=" & Session("optionLinkTableID") & ">" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtOptionLinkViewID name=txtOptionLinkViewID value=" & Session("optionLinkViewID") & ">" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtOptionLinkOrderID name=txtOptionLinkOrderID value=" & Session("optionLinkOrderID") & ">" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtOptionCourseTitle name=txtOptionCourseTitle value=""" & Session("optionCourseTitle") & """>" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtOptionRecordID name=txtOptionRecordID value=" & Session("optionRecordID") & ">" & vbCrLf)
%>
</FORM>
<INPUT type='hidden' id=txtTicker name=txtTicker value=0>
<INPUT type='hidden' id=txtLastKeyFind name=txtLastKeyFind value="">

<FORM action="tbTransferCourseFind_Submit" method="post" id="frmGotoOption" name="frmGotoOption" style="visibility:hidden;display:none">
		<%Html.RenderPartial("~/Views/Shared/gotoOption.ascx")%>
</FORM>

<script type="text/javascript"> tbTransferCourseFind_onload();</script>	
	

</div>