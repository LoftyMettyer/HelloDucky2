<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<%@ Language=VBScript %>
<!--#INCLUDE FILE="include/svrCleanup.asp" -->
<%
	Dim sReferringPage

	' Only open the form if there was a referring page.
	' If it wasn't then redirect to the login page.
	sReferringPage = Request.ServerVariables("HTTP_REFERER") 
	if inStrRev(sReferringPage, "/") > 0 then
		sReferringPage = mid(sReferringPage, inStrRev(sReferringPage, "/") + 1)
	end if

	if len(sReferringPage) = 0 then
		Response.Redirect("login.asp")
	end if
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK href="OpenHR.css" rel=stylesheet type=text/css >
<TITLE>OpenHR Intranet</TITLE>
<meta http-equiv="X-UA-Compatible" content="IE=5">
<!--#include file="include\ctl_SetFont.txt"-->

<OBJECT 
	classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" 
	id="Microsoft_Licensed_Class_Manager_1_0" 
	VIEWASTEXT>
	<PARAM NAME="LPKPath" VALUE="lpks/main.lpk">
</OBJECT>

<SCRIPT FOR=window EVENT=onload LANGUAGE=JavaScript>
<!--
	var fOK
	fOK = true;	

	var sErrMsg = frmFindForm.txtErrorDescription.value;
	if (sErrMsg.length > 0) {
		fOK = false;
		window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sErrMsg);
		window.parent.location.replace("login.asp");
	}
	
	if (fOK == true) {
		sErrMsg = frmFindForm.txtFailureDescription.value;
		if (sErrMsg.length > 0) {
			fOK = false;
			window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sErrMsg);
			Cancel();
		}
	}
	
	if (fOK == true) {
		if (frmFindForm.selectView.length == 0) {
			fOK = false;
			window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("You do not have permission to read the course table.");
			Cancel();
		}
	}
	
	if (fOK == true) {
		if (frmFindForm.selectOrder.length == 0) {
			fOK = false;
			window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("You do not have permission to use any of the course table orders.");
			Cancel();
		}
	}
	
	if (fOK == true) {
		setGridFont(frmFindForm.ssOleDBGridRecords);

		// Expand the option frame and hide the work frame.
		window.parent.document.all.item("workframeset").cols = "0, *";	
				
		// Set focus onto one of the form controls. 
		// NB. This needs to be done before making any reference to the grid
		frmFindForm.cmdCancel.focus();

		window.parent.frames("workframe").document.forms("frmFindForm").ssOleDBGridFindRecords.style.visibility = "hidden";

		// Get the optionData.asp to get the link find records.
		var optionDataForm = window.parent.frames("optiondataframe").document.forms("frmGetOptionData");
		optionDataForm.txtOptionAction.value = "LOADADDFROMWAITINGLIST";
		optionDataForm.txtOptionTableID.value = frmFindForm.txtOptionLinkTableID.value;
		optionDataForm.txtOptionViewID.value = frmFindForm.selectView.options[frmFindForm.selectView.selectedIndex].value;
		optionDataForm.txtOptionOrderID.value = frmFindForm.selectOrder.options[frmFindForm.selectOrder.selectedIndex].value;
		optionDataForm.txtOptionRecordID.value = frmFindForm.txtOptionRecordID.value;
		optionDataForm.txtOptionFirstRecPos.value = 1;
		optionDataForm.txtOptionCurrentRecCount.value = 0;
		optionDataForm.txtOptionPageAction.value = "LOAD"

		frmFindForm.txtOptionLinkViewID.value = optionDataForm.txtOptionViewID.value;
		frmFindForm.txtOptionLinkOrderID.value = optionDataForm.txtOptionOrderID.value;

		window.parent.frames("optiondataframe").refreshOptionData();
	}
	-->
</SCRIPT>

<script LANGUAGE="JavaScript">
<!--
	function Select()
	{  
		if (txtStatusPExists.value != "True") {
			window.parent.frames("workframe").document.forms("frmFindForm").ssOleDBGridFindRecords.style.visibility = "visible";
		}
	
		frmGotoOption.txtGotoOptionAction.value = "SELECTADDFROMWAITINGLIST_1";
		frmGotoOption.txtGotoOptionRecordID.value = frmFindForm.txtOptionRecordID.value;
		frmGotoOption.txtGotoOptionLinkRecordID.value = selectedRecordID();
		frmGotoOption.txtGotoOptionPage.value = "emptyoption.asp";
		frmGotoOption.submit();
	}

	function Cancel()
	{  
		window.parent.frames("workframe").document.forms("frmFindForm").ssOleDBGridFindRecords.style.visibility = "visible";

		frmGotoOption.txtGotoOptionAction.value = "CANCEL";
		frmGotoOption.txtGotoOptionLinkRecordID.value = 0;
		frmGotoOption.txtGotoOptionPage.value = "emptyoption.asp";
		frmGotoOption.submit();
	}

	/* Return the ID of the record selected in the find form. */
	function selectedRecordID()
	{  
		var iRecordID
		var iIndex
		var iIDColumnIndex
		var sColumnName

		iRecordID = 0
		iIDColumnIndex = 0
	
		if (frmFindForm.ssOleDBGridRecords.SelBookmarks.Count > 0) {
			for (iIndex = 0; iIndex < frmFindForm.ssOleDBGridRecords.Cols; iIndex++) {
				sColumnName = frmFindForm.ssOleDBGridRecords.Columns(iIndex).Name;
				if (sColumnName.toUpperCase() == "ID") {
					iIDColumnIndex = iIndex;
					break;
				}
			}
    
			iRecordID = frmFindForm.ssOleDBGridRecords.Columns(iIDColumnIndex).Value;
		}

		return(iRecordID);
	}

	function refreshControls() {
		if (frmFindForm.ssOleDBGridRecords.rows > 0) {
			if (frmFindForm.ssOleDBGridRecords.SelBookmarks.Count > 0) {
				button_disable(frmFindForm.cmdSelect, false);
			}
			else {
				button_disable(frmFindForm.cmdSelect, true);
			}
		}
		else {
			button_disable(frmFindForm.cmdSelect, true);
		}
	
		if (frmFindForm.selectOrder.length <= 1) {
			combo_disable(frmFindForm.selectOrder, true);
			button_disable(frmFindForm.btnGoOrder, true);
		}

		if (frmFindForm.selectView.length <= 1) {
			combo_disable(frmFindForm.selectView, true);
			button_disable(frmFindForm.btnGoView, true);
		}
	}

	function goView() {
		// Get the optionData.asp to get the link find records.
		var optionDataForm = window.parent.frames("optiondataframe").document.forms("frmGetOptionData");
		optionDataForm.txtOptionAction.value = "LOADADDFROMWAITINGLIST";
		optionDataForm.txtOptionTableID.value = frmFindForm.txtOptionLinkTableID.value;
		optionDataForm.txtOptionViewID.value = frmFindForm.selectView.options[frmFindForm.selectView.selectedIndex].value;
		optionDataForm.txtOptionOrderID.value = frmFindForm.selectOrder.options[frmFindForm.selectOrder.selectedIndex].value;
		optionDataForm.txtOptionRecordID.value = frmFindForm.txtOptionRecordID.value;
		optionDataForm.txtOptionFirstRecPos.value = 1;
		optionDataForm.txtOptionCurrentRecCount.value = 0;

		frmFindForm.txtOptionLinkViewID.value = optionDataForm.txtOptionViewID.value;
		frmFindForm.txtOptionLinkOrderID.value = optionDataForm.txtOptionOrderID.value;

		window.parent.frames("optiondataframe").refreshOptionData();
	}

	function goOrder() {
		// Get the optionData.asp to get the link find records.
		var optionDataForm = window.parent.frames("optiondataframe").document.forms("frmGetOptionData");
		optionDataForm.txtOptionAction.value = "LOADADDFROMWAITINGLIST";
		optionDataForm.txtOptionTableID.value = frmFindForm.txtOptionLinkTableID.value;
		optionDataForm.txtOptionViewID.value = frmFindForm.selectView.options[frmFindForm.selectView.selectedIndex].value;
		optionDataForm.txtOptionOrderID.value = frmFindForm.selectOrder.options[frmFindForm.selectOrder.selectedIndex].value;
		optionDataForm.txtOptionRecordID.value = frmFindForm.txtOptionRecordID.value;
		optionDataForm.txtOptionFirstRecPos.value = 1;
		optionDataForm.txtOptionCurrentRecCount.value = 0;

		frmFindForm.txtOptionLinkViewID.value = optionDataForm.txtOptionViewID.value;
		frmFindForm.txtOptionLinkOrderID.value = optionDataForm.txtOptionOrderID.value;

		window.parent.frames("optiondataframe").refreshOptionData();
	}

	function selectedOrderID() {
		return frmFindForm.selectOrder.options[frmFindForm.selectOrder.selectedIndex].value;
	}

	function selectedViewID() {
		return frmFindForm.selectView.options[frmFindForm.selectView.selectedIndex].value;
	}

	function locateRecord(psFileName)
	{  
		var fFound

		fFound = false;
	
		frmFindForm.ssOleDBGridRecords.redraw = false;

		frmFindForm.ssOleDBGridRecords.MoveLast();
		frmFindForm.ssOleDBGridRecords.MoveFirst();

		for (iIndex = 1; iIndex <= frmFindForm.ssOleDBGridRecords.rows; iIndex++) {	
			var sGridValue = new String(frmFindForm.ssOleDBGridRecords.Columns(0).value);
			sGridValue = sGridValue.substr(0, psFileName.length).toUpperCase();
			if (sGridValue == psFileName.toUpperCase()) {
				frmFindForm.ssOleDBGridRecords.SelBookmarks.Add(frmFindForm.ssOleDBGridRecords.Bookmark);
				fFound = true;
				break;
			}

			if (iIndex < frmFindForm.ssOleDBGridRecords.rows) {
				frmFindForm.ssOleDBGridRecords.MoveNext();
			}
			else {
				break;
			}
		}

		if ((fFound == false) && (frmFindForm.ssOleDBGridRecords.rows > 0)) {
			// Select the top row.
			frmFindForm.ssOleDBGridRecords.MoveFirst();
			frmFindForm.ssOleDBGridRecords.SelBookmarks.Add(frmFindForm.ssOleDBGridRecords.Bookmark);
		}

		frmFindForm.ssOleDBGridRecords.redraw = true;
	}
	-->
</script>

<SCRIPT FOR=ssOleDBGridRecords EVENT=dblClick LANGUAGE=JavaScript>
<!--
	Select();
	-->
</script>

<SCRIPT FOR=ssOleDBGridRecords EVENT=click LANGUAGE=JavaScript>
<!--
	refreshControls();
	-->
</script>

<SCRIPT FOR=ssOleDBGridRecords EVENT=KeyPress(iKeyAscii) LANGUAGE=JavaScript>
<!--
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
	-->
</script>
<!--#INCLUDE FILE="include/ctl_SetStyles.txt" -->
</HEAD>

<BODY <%=session("BodyTag")%>>
<FORM action="" method=POST id=frmFindForm name=frmFindForm>

<table align=center class="outline" cellPadding=5 cellSpacing=0 width=100% height=100%>
	<TR>
		<TD>
			<TABLE WIDTH="100%" height="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
				<TR>
					<TD height=10 colspan=3></td>
				</tr>
				<TR>
					<TD align=center height=10 colspan=3>
						<h3 align=center>Add From Waiting List</h3>
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
									<SELECT id=selectView name=selectView style="HEIGHT: 22px; WIDTH: 200px" class="combo">
<%
	on error resume next

	if (len(sErrorDescription) = 0) and (len(sFailureDescription) = 0) then
		' Get the view records.
		Set cmdViewRecords = Server.CreateObject("ADODB.Command")
		cmdViewRecords.CommandText = "sp_ASRIntGetLinkViews"
		cmdViewRecords.CommandType = 4 ' Stored Procedure
		Set cmdViewRecords.ActiveConnection = session("databaseConnection")

		Set prmTableID = cmdViewRecords.CreateParameter("tableID",3,1)
		cmdViewRecords.Parameters.Append prmTableID
		prmTableID.value = cleanNumeric(session("optionLinkTableID"))

		Set prmDfltOrderID = cmdViewRecords.CreateParameter("dfltOrderID",3,2) ' 11=integer, 2=output
		cmdViewRecords.Parameters.Append prmDfltOrderID

		err = 0
		Set rstViewRecords = cmdViewRecords.Execute

		if (err <> 0) then
			sErrorDescription = "The Course views could not be retrieved." & vbcrlf & formatError(Err.Description)
		end if

		if (len(sErrorDescription) = 0) and (len(sFailureDescription) = 0) then
			do while not rstViewRecords.EOF
				Response.Write "						<OPTION value=" & rstViewRecords.Fields(0).Value 
				if rstViewRecords.Fields(0).Value = session("optionLinkViewID") then
					Response.Write " SELECTED"
				end if

				if rstViewRecords.Fields(0).Value = 0 then
					Response.Write ">" & replace(rstViewRecords.Fields(1).Value, "_", " ") & "</OPTION>" & vbcrlf
				else
					Response.Write ">'" & replace(rstViewRecords.Fields(1).Value, "_", " ") & "' view</OPTION>" & vbcrlf
				end if

				rstViewRecords.MoveNext
			loop
			
			if (rstViewRecords.EOF and rstViewRecords.BOF) then
				sFailureDescription = "You do not have permission to read the Course table."
			end if
		
			' Release the ADO recordset object.
			rstViewRecords.close
			Set rstViewRecords = nothing
	
			' NB. IMPORTANT ADO NOTE.
			' When calling a stored procedure which returns a recordset AND has output parameters
			' you need to close the recordset and set it to nothing before using the output parameters. 
			if session("optionLinkOrderID") <= 0 then
				session("optionLinkOrderID") = cmdViewRecords.Parameters("dfltOrderID").Value
			end if
		end if

		' Release the ADO command object.
		Set cmdViewRecords = nothing
	end if
%>
						</SELECT>						
					</TD>
					<TD width=10>
						<INPUT type="button" value="Go" id=btnGoView class="btn" name=btnGoView 
						    onclick="goView()"
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
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
	if (len(sErrorDescription) = 0) and (len(sFailureDescription) = 0) then
		' Get the order records.
		Set cmdOrderRecords = Server.CreateObject("ADODB.Command")
		cmdOrderRecords.CommandText = "sp_ASRIntGetTableOrders"
		cmdOrderRecords.CommandType = 4 ' Stored Procedure
		Set cmdOrderRecords.ActiveConnection = session("databaseConnection")

		Set prmTableID = cmdOrderRecords.CreateParameter("tableID",3,1)
		cmdOrderRecords.Parameters.Append prmTableID
		prmTableID.value = cleanNumeric(session("optionLinkTableID"))

		Set prmViewID = cmdOrderRecords.CreateParameter("viewID",3,1)
		cmdOrderRecords.Parameters.Append prmViewID
		prmViewID.value = 0

		err = 0
		Set rstOrderRecords = cmdOrderRecords.Execute

		if (err <> 0) then
			sErrorDescription = "The order records could not be retrieved." & vbcrlf & formatError(Err.Description)
		end if

		if (len(sErrorDescription) = 0) and (len(sFailureDescription) = 0) then
			do while not rstOrderRecords.EOF
				Response.Write "						<OPTION value=" & rstOrderRecords.Fields(1).Value 

				if rstOrderRecords.Fields(1).Value = session("optionLinkOrderID") then
					Response.Write " SELECTED"
				end if

				Response.Write ">" & replace(rstOrderRecords.Fields(0).Value, "_", " ") & "</OPTION>" & vbcrlf

				rstOrderRecords.MoveNext
			loop

			' Release the ADO recordset object.
			rstOrderRecords.close
			Set rstOrderRecords = nothing
		end if
	
		' Release the ADO command object.
		Set cmdOrderRecords = nothing
	end if
%>
									</SELECT>
								</TD>
								<TD width=10>
									<INPUT type="button" value="Go" id=btnGoOrder name=btnGoOrder class="btn"
									    onclick="goOrder()"
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
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
						<OBJECT classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" id=ssOleDBGridRecords name=ssOleDBGridRecords     codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:100%">
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
							<PARAM NAME="MultiLine" VALUE="-1">
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
						<table WIDTH=100% class="invisible" CELLSPACING="0" CELLPADDING="0">
							<TR>
								<TD colspan=4>
								</TD>
							</TR>
							<tr>	
								<td>
								</td>
								<td width=10>
									<input id="cmdSelect" name="cmdSelect" type="button" value="Select" style="WIDTH: 75px" width="75" class="btn"
									    onclick="Select()"
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
								</td>
								<td width=40>
								</td>
								<td width=10>
									<input id="cmdCancel" name="cmdCancel" type="button" value="Cancel" style="WIDTH: 75px" width="75" class="btn"
									    onclick="Cancel()"
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
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
	Response.Write "<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtFailureDescription name=txtFailureDescription value=""" & sFailureDescription & """>" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtOptionLinkTableID name=txtOptionLinkTableID value=" & session("optionLinkTableID") & ">" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtOptionLinkViewID name=txtOptionLinkViewID value=" & session("optionLinkViewID") & ">" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtOptionLinkOrderID name=txtOptionLinkOrderID value=" & session("optionLinkOrderID") & ">" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtOptionCourseTitle name=txtOptionCourseTitle value=" & session("optionCourseTitle") & ">" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtOptionRecordID name=txtOptionRecordID value=" & session("optionRecordID") & ">" & vbcrlf
%>
</FORM>
<INPUT type='hidden' id=txtTicker name=txtTicker value=0>
<INPUT type='hidden' id=txtLastKeyFind name=txtLastKeyFind value="">
<INPUT type='hidden' id=txtStatusPExists name=txtStatusPExists value=<%=session("TB_TBStatusPExists")%>>

<FORM action="tbAddFromWaitingListFind_Submit.asp" method=post id=frmGotoOption name=frmGotoOption>
<!--#include file="include\gotoOption.txt"-->
</FORM>

</BODY>
</HTML>

<!-- Embeds createActiveX.js script reference -->
<!--#include file="include\ctl_CreateControl.txt"-->

