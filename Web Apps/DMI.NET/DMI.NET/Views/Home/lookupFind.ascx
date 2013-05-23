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
	
	var sErrMsg = frmLookupFindForm.txtErrorDescription.value;
	if (sErrMsg.length > 0) {
		fOK = false;
		window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sErrMsg);
		window.parent.location.replace("login.asp");
	}

	if (fOK == true) {
		sErrMsg = frmLookupFindForm.txtFailureDescription.value;
		if (sErrMsg.length > 0) {
			fOK = false;
			window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sErrMsg);
			window.parent.location.replace("login.asp");
		}
	}

	if (fOK == true) {
		setGridFont(frmLookupFindForm.ssOleDBGrid);
	
		// Expand the option frame and hide the work frame.
		window.parent.document.all.item("workframeset").cols = "0, *";	
				
		// Set focus onto one of the form controls. 
		// NB. This needs to be done before making any reference to the grid
		frmLookupFindForm.cmdCancel.focus();

		// Fault 3503
		window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "hidden";

		// Get the optionData.asp to get the link find records.
		var optionDataForm = window.parent.frames("optiondataframe").document.forms("frmGetOptionData");
		optionDataForm.txtOptionAction.value = "LOADLOOKUPFIND";
		optionDataForm.txtOptionColumnID.value = frmLookupFindForm.txtOptionColumnID.value;
		optionDataForm.txtOptionLookupColumnID.value = frmLookupFindForm.txtOptionLookupColumnID.value;
		optionDataForm.txtOptionLookupFilterValue.value = frmLookupFindForm.txtOptionLookupFilterValue.value;
		optionDataForm.txtOptionPageAction.value = "LOAD"
		optionDataForm.txtOptionFirstRecPos.value = 1;
		optionDataForm.txtOptionCurrentRecCount.value = 0;
		optionDataForm.txtOptionIsLookupTable.value = frmLookupFindForm.txtIsLookupTable.value;
		optionDataForm.txtOptionRecordID.value = window.parent.frames("workframe").document.forms("frmRecordEditForm").txtCurrentRecordID.value;

		optionDataForm.txtOptionParentTableID.value = window.parent.frames("workframe").document.forms("frmRecordEditForm").txtCurrentParentTableID.value;
		optionDataForm.txtOptionParentRecordID.value = window.parent.frames("workframe").document.forms("frmRecordEditForm").txtCurrentParentRecordID.value;

		if (frmLookupFindForm.txtIsLookupTable.value == "False") {
			optionDataForm.txtOptionTableID.value = frmLookupFindForm.txtOptionLinkTableID.value;
			optionDataForm.txtOptionViewID.value = frmLookupFindForm.selectView.options[frmLookupFindForm.selectView.selectedIndex].value;
			optionDataForm.txtOptionOrderID.value = frmLookupFindForm.selectOrder.options[frmLookupFindForm.selectOrder.selectedIndex].value;
		}
		else {
			optionDataForm.txtOptionTableID.value = 0;
			optionDataForm.txtOptionViewID.value = 0;
			optionDataForm.txtOptionOrderID.value = 0;
		}

		window.parent.frames("optiondataframe").refreshOptionData();
	}
	-->
</SCRIPT>

<script LANGUAGE="JavaScript">
<!--
	function SelectLookup()
	{  
		if (frmLookupFindForm.ssOleDBGrid.SelBookmarks.Count > 0) {
			// Fault 3503
			window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "visible";

			frmGotoOption.txtGotoOptionColumnID.value = frmLookupFindForm.txtOptionColumnID.value;
			frmGotoOption.txtGotoOptionLookupColumnID.value = frmLookupFindForm.txtOptionLookupColumnID.value;
			frmGotoOption.txtGotoOptionLookupValue.value = selectedValue();
			frmGotoOption.txtGotoOptionAction.value = "SELECTLOOKUP";
			frmGotoOption.txtGotoOptionPage.value = "emptyoption.asp";
			frmGotoOption.submit();
		}
	}

	function ClearLookup()
	{  
		// Fault 3503
		window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "visible";

		frmGotoOption.txtGotoOptionColumnID.value = frmLookupFindForm.txtOptionColumnID.value;
		frmGotoOption.txtGotoOptionLookupColumnID.value = frmLookupFindForm.txtOptionLookupColumnID.value;
		frmGotoOption.txtGotoOptionLookupValue.value = "";
		frmGotoOption.txtGotoOptionAction.value = "SELECTLOOKUP";
		frmGotoOption.txtGotoOptionPage.value = "emptyoption.asp";
		frmGotoOption.submit();
	}

	function CancelLookup()
	{  
		// Fault 3503
		window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "visible";

		frmGotoOption.txtGotoOptionAction.value = "CANCEL";
		frmGotoOption.txtGotoOptionPage.value = "emptyoption.asp";
		frmGotoOption.submit();
	}

	/* Return the value of the record selected in the find form. */
	function selectedValue()
	{  
		var sValue

		sValue = "";
		if (frmLookupFindForm.ssOleDBGrid.SelBookmarks.Count > 0) {
			if (frmLookupFindForm.txtIsLookupTable.value == "False") {
				sValue = frmLookupFindForm.ssOleDBGrid.Columns(parseInt(frmLookupFindForm.txtLookupColumnGridPosition.value)).value;
			}
			else {
				sValue = frmLookupFindForm.ssOleDBGrid.Columns(0).Value;
			}
		}

		return(sValue);
	}

	/* Sequential search the grid for the required OLE. */
	function locateRecord(psFileName, pfExactMatch)
	{  
		var fFound;
		var iPos;
	
		iPos = 0;
		if (frmLookupFindForm.txtIsLookupTable.value == "False") {
			iPos = parseInt(frmLookupFindForm.txtLookupColumnGridPosition.value);
		}

		fFound = false;
	
		frmLookupFindForm.ssOleDBGrid.redraw = false;

		frmLookupFindForm.ssOleDBGrid.MoveLast();
		frmLookupFindForm.ssOleDBGrid.MoveFirst();

		for (iIndex = 1; iIndex <= frmLookupFindForm.ssOleDBGrid.rows; iIndex++) {		
			if (pfExactMatch == true) {	
				if (frmLookupFindForm.ssOleDBGrid.Columns(iPos).value == psFileName) {
					frmLookupFindForm.ssOleDBGrid.SelBookmarks.Add(frmLookupFindForm.ssOleDBGrid.Bookmark);
					fFound = true;
					break;
				}
			}
			else {
				var sGridValue = new String(frmLookupFindForm.ssOleDBGrid.Columns(iPos).value);
				sGridValue = sGridValue.substr(0, psFileName.length).toUpperCase();
				if (sGridValue == psFileName.toUpperCase()) {
					frmLookupFindForm.ssOleDBGrid.SelBookmarks.Add(frmLookupFindForm.ssOleDBGrid.Bookmark);
					fFound = true;
					break;
				}
			}
		
			if (iIndex < frmLookupFindForm.ssOleDBGrid.rows) {
				frmLookupFindForm.ssOleDBGrid.MoveNext();
			}
			else {
				break;
			}
		}

		if ((fFound == false) && (frmLookupFindForm.ssOleDBGrid.rows > 0)) {
			// Select the top row.
			frmLookupFindForm.ssOleDBGrid.MoveFirst();
			frmLookupFindForm.ssOleDBGrid.SelBookmarks.Add(frmLookupFindForm.ssOleDBGrid.Bookmark);
		}

		frmLookupFindForm.ssOleDBGrid.redraw = true;
	}

	function refreshControls() {
		if (frmLookupFindForm.ssOleDBGrid.rows > 0) {
			if (frmLookupFindForm.ssOleDBGrid.SelBookmarks.Count > 0) {
				button_disable(frmLookupFindForm.cmdSelectLookup, false);
			}
			else {
				button_disable(frmLookupFindForm.cmdSelectLookup, true);
			}
		}
		else {
			button_disable(frmLookupFindForm.cmdSelectLookup, true);
		}

		if (frmLookupFindForm.txtOptionLookupMandatory.value == "true") {
			button_disable(frmLookupFindForm.cmdClearLookup, true);
		}
		else {
			button_disable(frmLookupFindForm.cmdClearLookup, false);
		}	
	}

	function goView() {
		var optionDataForm = window.parent.frames("optiondataframe").document.forms("frmGetOptionData");
		optionDataForm.txtOptionAction.value = "LOADLOOKUPFIND";
		optionDataForm.txtOptionColumnID.value = frmLookupFindForm.txtOptionColumnID.value;
		optionDataForm.txtOptionLookupColumnID.value = frmLookupFindForm.txtOptionLookupColumnID.value;
		optionDataForm.txtOptionLookupFilterValue.value = frmLookupFindForm.txtOptionLookupFilterValue.value;
		optionDataForm.txtOptionPageAction.value = "LOAD"
		optionDataForm.txtOptionFirstRecPos.value = 1;
		optionDataForm.txtOptionCurrentRecCount.value = 0;
		optionDataForm.txtOptionIsLookupTable.value = frmLookupFindForm.txtIsLookupTable.value;
		optionDataForm.txtOptionRecordID.value = window.parent.frames("workframe").document.forms("frmRecordEditForm").txtCurrentRecordID.value;

		optionDataForm.txtOptionParentTableID.value = window.parent.frames("workframe").document.forms("frmRecordEditForm").txtCurrentParentTableID.value;
		optionDataForm.txtOptionParentRecordID.value = window.parent.frames("workframe").document.forms("frmRecordEditForm").txtCurrentParentRecordID.value;

		if (frmLookupFindForm.txtIsLookupTable.value == "False") {
			optionDataForm.txtOptionTableID.value = frmLookupFindForm.txtOptionLinkTableID.value;
			optionDataForm.txtOptionViewID.value = frmLookupFindForm.selectView.options[frmLookupFindForm.selectView.selectedIndex].value;
			optionDataForm.txtOptionOrderID.value = frmLookupFindForm.selectOrder.options[frmLookupFindForm.selectOrder.selectedIndex].value;
		}
		else {
			optionDataForm.txtOptionTableID.value = 0;
			optionDataForm.txtOptionViewID.value = 0;
			optionDataForm.txtOptionOrderID.value = 0;
		}
		
		window.parent.frames("optiondataframe").refreshOptionData();
	}

	function goOrder() {
		var optionDataForm = window.parent.frames("optiondataframe").document.forms("frmGetOptionData");
		optionDataForm.txtOptionAction.value = "LOADLOOKUPFIND";
		optionDataForm.txtOptionColumnID.value = frmLookupFindForm.txtOptionColumnID.value;
		optionDataForm.txtOptionLookupColumnID.value = frmLookupFindForm.txtOptionLookupColumnID.value;
		optionDataForm.txtOptionLookupFilterValue.value = frmLookupFindForm.txtOptionLookupFilterValue.value;
		optionDataForm.txtOptionPageAction.value = "LOAD"
		optionDataForm.txtOptionFirstRecPos.value = 1;
		optionDataForm.txtOptionCurrentRecCount.value = 0;
		optionDataForm.txtOptionIsLookupTable.value = frmLookupFindForm.txtIsLookupTable.value;
		optionDataForm.txtOptionRecordID.value = window.parent.frames("workframe").document.forms("frmRecordEditForm").txtCurrentRecordID.value;

		optionDataForm.txtOptionParentTableID.value = window.parent.frames("workframe").document.forms("frmRecordEditForm").txtCurrentParentTableID.value;
		optionDataForm.txtOptionParentRecordID.value = window.parent.frames("workframe").document.forms("frmRecordEditForm").txtCurrentParentRecordID.value;

		if (frmLookupFindForm.txtIsLookupTable.value == "False") {
			optionDataForm.txtOptionTableID.value = frmLookupFindForm.txtOptionLinkTableID.value;
			optionDataForm.txtOptionViewID.value = frmLookupFindForm.selectView.options[frmLookupFindForm.selectView.selectedIndex].value;
			optionDataForm.txtOptionOrderID.value = frmLookupFindForm.selectOrder.options[frmLookupFindForm.selectOrder.selectedIndex].value;
		}
		else {
			optionDataForm.txtOptionTableID.value = 0;
			optionDataForm.txtOptionViewID.value = 0;
			optionDataForm.txtOptionOrderID.value = 0;
		}
		
		window.parent.frames("optiondataframe").refreshOptionData();
	}
	-->
</script>

<SCRIPT FOR=ssOleDBGrid EVENT=dblClick LANGUAGE=JavaScript>
<!--
	SelectLookup();
	-->
</script>

<SCRIPT FOR=ssOleDBGrid EVENT=KeyPress(iKeyAscii) LANGUAGE=JavaScript>
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

		locateRecord(sFind, false);
	}
	-->
</script>

<SCRIPT FOR=ssOleDBGrid EVENT=click LANGUAGE=JavaScript>
<!--
	refreshControls();
	-->
</script>
<!--#INCLUDE FILE="include/ctl_SetStyles.txt" -->
</HEAD>

<BODY <%=session("BodyTag")%>>
<FORM action="" method=POST id=frmLookupFindForm name=frmLookupFindForm>
<%

	fIsLookupTable = false
	lngLookupTableID = 0
	
	Set cmdGetTable = Server.CreateObject("ADODB.Command")
	cmdGetTable.CommandText = "spASRIntGetColumnTableID"
	cmdGetTable.CommandType = 4 ' Stored procedure
	Set cmdGetTable.ActiveConnection = session("databaseConnection")

	Set prmColumnID = cmdGetTable.CreateParameter("LookupColumnID",3,1)
	cmdGetTable.Parameters.Append prmColumnID
	prmColumnID.value = cleanNumeric(session("optionLookupColumnID"))

	Set prmTableID = cmdGetTable.CreateParameter("tableID",3,2) ' 3=integer, 2=output
	cmdGetTable.Parameters.Append prmTableID

	err = 0
	cmdGetTable.Execute
	
	if (err <> 0) then
		sErrorDescription = "Error getting the lookup column table ID." & vbcrlf & formatError(Err.Description)
	else
		lngLookupTableID = clng(cmdGetTable.Parameters("tableID").Value)
	end if

	set cmdGetTable = nothing

	if len(sErrorDescription) = 0 then
		Set cmdIsLookupTable = Server.CreateObject("ADODB.Command")
		cmdIsLookupTable.CommandText = "spASRIntIsLookupTable"
		cmdIsLookupTable.CommandType = 4 ' Stored procedure
		Set cmdIsLookupTable.ActiveConnection = session("databaseConnection")

		Set prmTableID = cmdIsLookupTable.CreateParameter("tableID",3,1)
		cmdIsLookupTable.Parameters.Append prmTableID
		prmTableID.value = cleanNumeric(lngLookupTableID)

		Set prmIsLookup = cmdIsLookupTable.CreateParameter("isLookup",11,2) ' 11=bit, 2=output
		cmdIsLookupTable.Parameters.Append prmIsLookup

		err = 0
		cmdIsLookupTable.Execute
	
		if (err <> 0) then
			sErrorDescription = "Error checking the lookup column table type." & vbcrlf & formatError(Err.Description)
		else
			fIsLookupTable = cbool(cmdIsLookupTable.Parameters("isLookup").Value)
		end if

		set cmdIsLookupTable = nothing
	end if
%>

<table align=center class="outline" cellPadding=5 cellSpacing=0 width=100% height=100%>
	<TR>
		<TD>
			<TABLE WIDTH="100%" height="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
				<TR>
					<TD height=10 colspan=3></td>
				</tr>
				<TR>
					<TD align=center height=10 colspan=3>
						<h3 align=center>Find Lookup Record</h3>
					</td>
				</tr>
				
<%if not fIsLookupTable then%>
				<tr>
					<td height=10>&nbsp;&nbsp;</td>
					<td height="10">
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

	if (len(sErrorDescription) = 0) and (len(sFailureDescription) = 0) then
		' Get the view records.
		Set cmdViewRecords = Server.CreateObject("ADODB.Command")
		cmdViewRecords.CommandText = "spASRIntGetLookupViews"
		cmdViewRecords.CommandType = 4 ' Stored Procedure
		Set cmdViewRecords.ActiveConnection = session("databaseConnection")

		Set prmTableID = cmdViewRecords.CreateParameter("tableID",3,1)
		cmdViewRecords.Parameters.Append prmTableID
		prmTableID.value = cleanNumeric(lngLookupTableID)

		Set prmDfltOrderID = cmdViewRecords.CreateParameter("dfltOrderID",3,2) ' 11=integer, 2=output
		cmdViewRecords.Parameters.Append prmDfltOrderID

		Set prmColumnID = cmdViewRecords.CreateParameter("columnID",3,1)
		cmdViewRecords.Parameters.Append prmColumnID
		prmColumnID.value = cleanNumeric(session("optionColumnID"))

		err = 0
		Set rstViewRecords = cmdViewRecords.Execute

		if (err <> 0) then
			sErrorDescription = "The lookup view records could not be retrieved." & vbcrlf & formatError(Err.Description)
		end if

		if (len(sErrorDescription) = 0) and (len(sFailureDescription) = 0) then
			do while not rstViewRecords.EOF
				Response.Write "						<OPTION value=" & rstViewRecords.Fields(0).Value 
				if rstViewRecords.Fields(0).Value = clng(session("optionLinkViewID")) then
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
				sFailureDescription = "You do not have permission to read the lookup table."
			end if
		
			' Release the ADO recordset object.
			rstViewRecords.close
			Set rstViewRecords = nothing
	
			' NB. IMPORTANT ADO NOTE.
			' When calling a stored procedure which returns a recordset AND has output parameters
			' you need to close the recordset and set it to nothing before using the output parameters. 
			if session("optionLookupOrderID") <= 0 then
				session("optionLookupOrderID") = cmdViewRecords.Parameters("dfltOrderID").Value
			end if
		end if

		' Release the ADO command object.
		Set cmdViewRecords = nothing
	end if
%>
									</SELECT>						
								</TD>
								<TD width=10>
									<INPUT type="button" value="Go" class="btn" id=btnGoView name=btnGoView 
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
		prmTableID.value = cleanNumeric(lngLookupTableID)

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
				if rstOrderRecords.Fields(1).Value = cint(session("optionOrderID")) then
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
									<INPUT type="button" value="Go" class="btn" id=btnGoOrder name=btnGoOrder 
									    onclick="goOrder()"
	                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
	                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
	                                    onfocus="try{button_onFocus(this);}catch(e){}"
	                                    onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
							</TR>
						</table>
					</td>
					<td height=10>&nbsp;&nbsp;</td>
				</tr>
				<TR>
					<TD height=10 colspan=3></td>
				</tr>
<%end if 'if fIsLookupTable then%>
				
				<TR>
					<td width=10></td>
					<TD>
						<OBJECT classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" id=ssOleDBGrid name=ssOleDBGrid  codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:100%">
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
					<td width=10></td>
				</TR>
				<TR>
					<td height=10 colspan=3>
					</td>
				</TR>
				<tr>
					<td width=20></td>
					<td height="10">
						<table WIDTH=100% class="invisible" CELLSPACING="0" CELLPADDING="0">
							<TR>
								<TD colspan=6>
								</TD>
							</TR>
							<tr>	
								<td>
								</td>
								<td width=10>
									<input id="cmdSelectLookup" name="cmdSelectLookup" type="button" value="Select" style="WIDTH: 75px" width="75" class="btn"
									    onclick="SelectLookup()"
	                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
	                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
	                                    onfocus="try{button_onFocus(this);}catch(e){}"
	                                    onblur="try{button_onBlur(this);}catch(e){}" />
								</td>
								<td width=40>
								</td>
								<td width=10>
									<input id="cmdClearLookup" name="cmdClearLookup" type="button" value="Clear" style="WIDTH: 75px" width="75" class="btn"
									    onclick="ClearLookup()"
	                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
	                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
	                                    onfocus="try{button_onFocus(this);}catch(e){}"
	                                    onblur="try{button_onBlur(this);}catch(e){}" />
								</td>
								<td width=40>
								</td>
								<td width=10>
									<input id="cmdCancel" name="cmdCancel" type="button" value="Cancel" style="WIDTH: 75px" width="75" class="btn" 
									    onclick="CancelLookup()"
	                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
	                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
	                                    onfocus="try{button_onFocus(this);}catch(e){}"
	                                    onblur="try{button_onBlur(this);}catch(e){}" />
								</td>
							</tr>			
						</table>
					</td>
					<td width=20></td>
				</tr>
				<TR>
					<TD height=10 colspan=3></td>
				</tr>
			</TABLE>
		</td>
	</tr>
</TABLE>

	<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value="<%=sErrorDescription%>">
	<INPUT type='hidden' id=txtFailureDescription name=txtFailureDescription value="<%=sFailureDescription%>">
	<INPUT type='hidden' id=txtOptionColumnID name=txtOptionColumnID value=<%=session("optionColumnID")%>>
	<INPUT type='hidden' id=txtOptionLookupColumnID name=txtOptionLookupColumnID value=<%=session("optionLookupColumnID")%>>
	<INPUT type='hidden' id=txtOptionLookupMandatory name=txtOptionLookupMandatory value=<%=session("optionLookupMandatory")%>>
	<INPUT type='hidden' id=txtOptionLookupValue name=txtOptionLookupValue value=<%=session("optionLookupValue")%>>
	<INPUT type='hidden' id=txtOptionLookupFilterValue name=txtOptionLookupFilterValue value="<%=session("optionLookupFilterValue")%>">	
	<INPUT type='hidden' id=txtIsLookupTable name=txtIsLookupTable value="<%=fIsLookupTable%>">	
	<INPUT type='hidden' id=txtOptionLinkTableID name=txtOptionLinkTableID value=<%=lngLookupTableID%>>	
	<INPUT type='hidden' id=txtLookupColumnGridPosition name=txtLookupColumnGridPosition value=0>	
</FORM>

<INPUT type='hidden' id=txtTicker name=txtTicker value=0>
<INPUT type='hidden' id=txtLastKeyFind name=txtLastKeyFind value="">

<FORM action="lookupFind_Submit.asp" method=post id=frmGotoOption name=frmGotoOption>
<!--#include file="include\gotoOption.txt"-->
</FORM>

</BODY>
</HTML>

<!-- Embeds createActiveX.js script reference -->
<!--#include file="include\ctl_CreateControl.txt"-->
