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

	var sErrMsg = frmOrderForm.txtErrorDescription.value;
	if (sErrMsg.length > 0) {
		fOK = false;
		window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sErrMsg);
		window.parent.location.replace("login.asp");
	}
	
	if (fOK == true) {
		setGridFont(frmOrderForm.ssOleDBGridOrderRecords);
	
		// Expand the option frame and hide the work frame.
		window.parent.document.all.item("workframeset").cols = "0, *";	
				
		// Set focus onto one of the form controls. 
		// NB. This needs to be done before making any reference to the grid
		frmOrderForm.cmdCancel.focus();

		// Select the current record in the grid if its there, else select the top record if there is one.
		if (frmOrderForm.ssOleDBGridOrderRecords.rows > 0) {
			if (frmOrderForm.txtCurrentOrderID.value > 0) {
				// Try to select the current record.
				locateRecord(frmOrderForm.txtCurrentOrderID.value, true);
			}
			else {
				// Select the top row.
				frmOrderForm.ssOleDBGridOrderRecords.MoveFirst();
				frmOrderForm.ssOleDBGridOrderRecords.SelBookmarks.Add(frmOrderForm.ssOleDBGridOrderRecords.Bookmark);
			}
		}

		// Get menu.asp to refresh the menu.
		// NPG20100824 Fault HRPRO1065 - leave menus disabled in these modal screens
		//window.parent.frames("menuframe").refreshMenu();

		// Hide the workframe recedit control. IE6 still displays it.
		sWorkPage = currentWorkFramePage();
		if (sWorkPage == "RECORDEDIT") {
			window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "hidden";
		}
		else {
			if (sWorkPage == "FIND") {
				window.parent.frames("workframe").document.forms("frmFindForm").ssOleDBGridFindRecords.style.visibility = "hidden";
			}
		}

		refreshControls();
	}
	-->
</SCRIPT>

<script LANGUAGE="JavaScript">
<!--
	function SelectOrder()
	{  
		// Redisplay the workframe recedit control. 
		sWorkPage = currentWorkFramePage();
		if (sWorkPage == "RECORDEDIT") {
			window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "visible";
		}
		else {
			if (sWorkPage == "FIND") {
				window.parent.frames("workframe").document.forms("frmFindForm").ssOleDBGridFindRecords.style.visibility = "visible";
			}
		}

		frmGotoOption.txtGotoOptionScreenID.value = frmOrderForm.txtOptionScreenID.value;
		frmGotoOption.txtGotoOptionTableID.value = frmOrderForm.txtOptionTableID.value;
		frmGotoOption.txtGotoOptionViewID.value = frmOrderForm.txtOptionViewID.value;
		frmGotoOption.txtGotoOptionOrderID.value = selectedRecordID();
		frmGotoOption.txtGotoOptionPage.value = "emptyoption.asp";
		frmGotoOption.txtGotoOptionAction.value = "SELECTORDER";
		frmGotoOption.submit();
	}

	function CancelOrder()
	{  
		// Redisplay the workframe recedit control. 
		sWorkPage = currentWorkFramePage();
		if (sWorkPage == "RECORDEDIT") {
			window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "visible";
			window.parent.document.all.item("workframeset").cols = "*, 0";	
			window.parent.frames("workframe").refreshData();
		}
		else {
			if (sWorkPage == "FIND") {
				window.parent.frames("workframe").document.forms("frmFindForm").ssOleDBGridFindRecords.style.visibility = "visible";
			}
		}

		frmGotoOption.txtGotoOptionAction.value = "CANCEL";
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
	
		if (frmOrderForm.ssOleDBGridOrderRecords.SelBookmarks.Count > 0) {
			for (iIndex = 0; iIndex < frmOrderForm.ssOleDBGridOrderRecords.Cols; iIndex++) {
				sColumnName = frmOrderForm.ssOleDBGridOrderRecords.Columns(iIndex).Name;
				if (sColumnName.toUpperCase() == "ORDERID") {
					iIDColumnIndex = iIndex;
					break;
				}
			}
    
			iRecordID = frmOrderForm.ssOleDBGridOrderRecords.Columns(iIDColumnIndex).Value;
		}

		return(iRecordID);
	}

	/* Sequential search the grid for the required ID. */
	function locateRecord(psSearchFor, pfIDMatch)
	{  
		var fFound
		var iIndex
		var iIDColumnIndex
		var sColumnName

		fFound = false;
	
		frmOrderForm.ssOleDBGridOrderRecords.redraw = false;

		if (pfIDMatch == true) {	
			// Locate the ID column in the grid.
			iIDColumnIndex = -1    
			for (iIndex = 0; iIndex < frmOrderForm.ssOleDBGridOrderRecords.Cols; iIndex++) {
				sColumnName = frmOrderForm.ssOleDBGridOrderRecords.Columns(iIndex).Name;
				if (sColumnName.toUpperCase() == "ORDERID") {
					iIDColumnIndex = iIndex;
					break;
				}
			}

			if (iIDColumnIndex >= 0) {	
				frmOrderForm.ssOleDBGridOrderRecords.MoveLast();
				frmOrderForm.ssOleDBGridOrderRecords.MoveFirst();

				for (iIndex = 1; iIndex <= frmOrderForm.ssOleDBGridOrderRecords.rows; iIndex++) {		
					if (frmOrderForm.ssOleDBGridOrderRecords.Columns(iIDColumnIndex).value == psSearchFor) {
						frmOrderForm.ssOleDBGridOrderRecords.SelBookmarks.Add(frmOrderForm.ssOleDBGridOrderRecords.Bookmark);
						fFound = true;
						break;
					}

					if (iIndex < frmOrderForm.ssOleDBGridOrderRecords.rows) {
						frmOrderForm.ssOleDBGridOrderRecords.MoveNext();
					}
					else {
						break;
					}
				}
			}
		}
		else {
			for (iIndex = 1; iIndex <= frmOrderForm.ssOleDBGridOrderRecords.rows; iIndex++) {		
				var sGridValue = new String(frmOrderForm.ssOleDBGridOrderRecords.Columns(0).value);
				sGridValue = sGridValue.substr(0, psSearchFor.length).toUpperCase();
				if (sGridValue == psSearchFor.toUpperCase()) {
					frmOrderForm.ssOleDBGridOrderRecords.SelBookmarks.Add(frmOrderForm.ssOleDBGridOrderRecords.Bookmark);
					fFound = true;
					break;
				}
			
				if (iIndex < frmOrderForm.ssOleDBGridOrderRecords.rows) {
					frmOrderForm.ssOleDBGridOrderRecords.MoveNext();
				}
				else {
					break;
				}
			}
		}
	
		if ((fFound == false) && (frmOrderForm.ssOleDBGridOrderRecords.rows > 0)) {
			// Select the top row.
			frmOrderForm.ssOleDBGridOrderRecords.MoveFirst();
			frmOrderForm.ssOleDBGridOrderRecords.SelBookmarks.Add(frmOrderForm.ssOleDBGridOrderRecords.Bookmark);
		}

		frmOrderForm.ssOleDBGridOrderRecords.redraw = true;
	}

	function refreshControls() {
		if (frmOrderForm.ssOleDBGridOrderRecords.rows > 0) {
			if (frmOrderForm.ssOleDBGridOrderRecords.SelBookmarks.Count > 0) {
				button_disable(frmOrderForm.cmdSelectOrder, false);
			}
			else {
				button_disable(frmOrderForm.cmdSelectOrder, true);
			}
		}
		else {
			button_disable(frmOrderForm.cmdSelectOrder, true);
		}
	}

	function currentWorkFramePage()
	{
		// Return the current page in the workframeset.
		sCols = window.parent.document.all.item("workframeset").cols;

		re = / /gi;
		sCols = sCols.replace(re, "");
		sCols = sCols.substr(0, 1);

		// Work frame is in view.
		sCurrentPage = window.parent.frames("workframe").document.location;
		sCurrentPage = sCurrentPage.toString();
	
		if (sCurrentPage.lastIndexOf("/") > 0) {
			sCurrentPage = sCurrentPage.substr(sCurrentPage.lastIndexOf("/") + 1);
		}

		if (sCurrentPage.indexOf(".") > 0) {
			sCurrentPage = sCurrentPage.substr(0, sCurrentPage.indexOf("."));
		}

		re = / /gi;
		sCurrentPage = sCurrentPage.replace(re, "");
		sCurrentPage = sCurrentPage.toUpperCase();
	
		return(sCurrentPage);	
	}
	-->
</script>

<SCRIPT FOR=ssOleDBGridOrderRecords EVENT=dblClick LANGUAGE=JavaScript>
<!--
	SelectOrder();
	-->
</script>

<SCRIPT FOR=ssOleDBGridOrderRecords EVENT=KeyPress(iKeyAscii) LANGUAGE=JavaScript>
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
<!--#INCLUDE FILE="include/ctl_SetStyles.txt" -->
</HEAD>

<BODY <%=session("BodyTag")%>>
<FORM action="" method=POST id=frmOrderForm name=frmOrderForm>
<table align=center class="outline" cellPadding=5 cellSpacing=0 width=100% height=100%>
	<TR>
		<TD>
			<TABLE ID="orderTable" WIDTH="100%" height="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
				<TR>
					<TD height=10 colspan=3>
						<h3 align=center>Select Order</h3>
					</TD>
				</TR>
				<TR>
					<TD width=20></TD>
					<TD>
<%
	if len(sErrorDescription) = 0 then
		' Get the order records.
		Set cmdOrderRecords = Server.CreateObject("ADODB.Command")
		cmdOrderRecords.CommandText = "sp_ASRIntGetTableOrders"
		cmdOrderRecords.CommandType = 4 ' Stored Procedure
		Set cmdOrderRecords.ActiveConnection = session("databaseConnection")

		Set prmTableID = cmdOrderRecords.CreateParameter("tableID",3,1)
		cmdOrderRecords.Parameters.Append prmTableID
		prmTableID.value = cleanNumeric(session("optionTableID"))

		Set prmViewID = cmdOrderRecords.CreateParameter("viewID",3,1)
		cmdOrderRecords.Parameters.Append prmViewID
		prmViewID.value = cleanNumeric(session("optionViewID"))

		err = 0
		Set rstOrderRecords = cmdOrderRecords.Execute

		if (err <> 0) then
			sErrorDescription = "The order records could not be retrieved." & vbcrlf & formatError(Err.Description)
		end if

		if len(sErrorDescription) = 0 then
			' Instantiate and initialise the grid. 
			Response.Write "			<OBJECT classid=""clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"" id=ssOleDBGridOrderRecords name=ssOleDBGridOrderRecords codebase=""cabs/COAInt_Grid.cab#version=3,1,3,6"" style=""LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:100%"">" & vbcrlf
			Response.Write "				<PARAM NAME=""ScrollBars"" VALUE=""4"">" & vbcrlf
			Response.Write "				<PARAM NAME=""_Version"" VALUE=""196617"">" & vbcrlf
			Response.Write "				<PARAM NAME=""DataMode"" VALUE=""2"">" & vbcrlf
			Response.Write "				<PARAM NAME=""Cols"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""Rows"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""BorderStyle"" VALUE=""1"">" & vbcrlf
			Response.Write "				<PARAM NAME=""RecordSelectors"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""GroupHeaders"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""ColumnHeaders"" VALUE=""1"">" & vbcrlf
			Response.Write "				<PARAM NAME=""GroupHeadLines"" VALUE=""1"">" & vbcrlf
			Response.Write "				<PARAM NAME=""HeadLines"" VALUE=""1"">" & vbcrlf
			Response.Write "				<PARAM NAME=""FieldDelimiter"" VALUE=""(None)"">" & vbcrlf
			Response.Write "				<PARAM NAME=""FieldSeparator"" VALUE=""(Tab)"">" & vbcrlf
			Response.Write "				<PARAM NAME=""Col.Count"" VALUE=""" & rstOrderRecords.fields.count & """>" & vbcrlf
			Response.Write "				<PARAM NAME=""stylesets.count"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""TagVariant"" VALUE=""EMPTY"">" & vbcrlf
			Response.Write "				<PARAM NAME=""UseGroups"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""HeadFont3D"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""Font3D"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""DividerType"" VALUE=""3"">" & vbcrlf
			Response.Write "				<PARAM NAME=""DividerStyle"" VALUE=""1"">" & vbcrlf
			Response.Write "				<PARAM NAME=""DefColWidth"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""BeveColorScheme"" VALUE=""2"">" & vbcrlf
			Response.Write "				<PARAM NAME=""BevelColorFrame"" VALUE=""-2147483642"">" & vbcrlf
			Response.Write "				<PARAM NAME=""BevelColorHighlight"" VALUE=""-2147483628"">" & vbcrlf
			Response.Write "				<PARAM NAME=""BevelColorShadow"" VALUE=""-2147483632"">" & vbcrlf
			Response.Write "				<PARAM NAME=""BevelColorFace"" VALUE=""-2147483633"">" & vbcrlf
			Response.Write "				<PARAM NAME=""CheckBox3D"" VALUE=""-1"">" & vbcrlf
			Response.Write "				<PARAM NAME=""AllowAddNew"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""AllowDelete"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""AllowUpdate"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""MultiLine"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""ActiveCellStyleSet"" VALUE="""">" & vbcrlf
			Response.Write "				<PARAM NAME=""RowSelectionStyle"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""AllowRowSizing"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""AllowGroupSizing"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""AllowColumnSizing"" VALUE=""-1"">" & vbcrlf
			Response.Write "				<PARAM NAME=""AllowGroupMoving"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""AllowColumnMoving"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""AllowGroupSwapping"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""AllowColumnSwapping"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""AllowGroupShrinking"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""AllowColumnShrinking"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""AllowDragDrop"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""UseExactRowCount"" VALUE=""-1"">" & vbcrlf
			Response.Write "				<PARAM NAME=""SelectTypeCol"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""SelectTypeRow"" VALUE=""1"">" & vbcrlf
			Response.Write "				<PARAM NAME=""SelectByCell"" VALUE=""-1"">" & vbcrlf
			Response.Write "				<PARAM NAME=""BalloonHelp"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""RowNavigation"" VALUE=""1"">" & vbcrlf
			Response.Write "				<PARAM NAME=""CellNavigation"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""MaxSelectedRows"" VALUE=""1"">" & vbcrlf
			Response.Write "				<PARAM NAME=""HeadStyleSet"" VALUE="""">" & vbcrlf
			Response.Write "				<PARAM NAME=""StyleSet"" VALUE="""">" & vbcrlf
			Response.Write "				<PARAM NAME=""ForeColorEven"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""ForeColorOdd"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""BackColorEven"" VALUE=""16777215"">" & vbcrlf
			Response.Write "				<PARAM NAME=""BackColorOdd"" VALUE=""16777215"">" & vbcrlf
			Response.Write "				<PARAM NAME=""Levels"" VALUE=""1"">" & vbcrlf
			Response.Write "				<PARAM NAME=""RowHeight"" VALUE=""503"">" & vbcrlf
			Response.Write "				<PARAM NAME=""ExtraHeight"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""ActiveRowStyleSet"" VALUE="""">" & vbcrlf
			Response.Write "				<PARAM NAME=""CaptionAlignment"" VALUE=""2"">" & vbcrlf
			Response.Write "				<PARAM NAME=""SplitterPos"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""SplitterVisible"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""Columns.Count"" VALUE=""" & rstOrderRecords.fields.count & """>" & vbcrlf

			for iLoop = 0 to (rstOrderRecords.fields.count - 1)

				if rstOrderRecords.fields(iLoop).name = "orderID" then
					Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").Width"" VALUE=""0"">" & vbcrlf
					Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").Visible"" VALUE=""0"">" & vbcrlf
				else
					Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").Width"" VALUE=""100000"">" & vbcrlf
					Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").Visible"" VALUE=""-1"">" & vbcrlf
				end if
	
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").Columns.Count"" VALUE=""1"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").Caption"" VALUE=""" & replace(rstOrderRecords.fields(iLoop).name, "_", " ") & """>" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").Name"" VALUE=""" & rstOrderRecords.fields(iLoop).name & """>" & vbcrlf			
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").Alignment"" VALUE=""0"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").CaptionAlignment"" VALUE=""3"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").Bound"" VALUE=""0"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").AllowSizing"" VALUE=""1"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").DataField"" VALUE=""Column " & iLoop & """>" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").DataType"" VALUE=""8"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").Level"" VALUE=""0"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").NumberFormat"" VALUE="""">" & vbcrlf			
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").Case"" VALUE=""0"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").FieldLen"" VALUE=""4096"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").VertScrollBar"" VALUE=""0"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").Locked"" VALUE=""0"">" & vbcrlf			
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").Style"" VALUE=""0"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").ButtonsAlways"" VALUE=""0"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").RowCount"" VALUE=""0"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").ColCount"" VALUE=""1"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").HasHeadForeColor"" VALUE=""0"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").HasHeadBackColor"" VALUE=""0"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").HasForeColor"" VALUE=""0"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").HasBackColor"" VALUE=""0"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").HeadForeColor"" VALUE=""0"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").HeadBackColor"" VALUE=""0"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").ForeColor"" VALUE=""0"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").BackColor"" VALUE=""0"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").HeadStyleSet"" VALUE="""">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").StyleSet"" VALUE="""">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").Nullable"" VALUE=""1"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").Mask"" VALUE="""">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").PromptInclude"" VALUE=""0"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").ClipMode"" VALUE=""0"">" & vbcrlf
				Response.Write "				<PARAM NAME=""Columns(" & iLoop & ").PromptChar"" VALUE=""95"">" & vbcrlf
			next 

			Response.Write "				<PARAM NAME=""UseDefaults"" VALUE=""-1"">" & vbcrlf
			Response.Write "				<PARAM NAME=""TabNavigation"" VALUE=""1"">" & vbcrlf
			Response.Write "				<PARAM NAME=""_ExtentX"" VALUE=""17330"">" & vbcrlf
			Response.Write "				<PARAM NAME=""_ExtentY"" VALUE=""1323"">" & vbcrlf
			Response.Write "				<PARAM NAME=""_StockProps"" VALUE=""79"">" & vbcrlf
			Response.Write "				<PARAM NAME=""Caption"" VALUE="""">" & vbcrlf
			Response.Write "				<PARAM NAME=""ForeColor"" VALUE=""0"">" & vbcrlf
			Response.Write "				<PARAM NAME=""BackColor"" VALUE=""16777215"">" & vbcrlf
			Response.Write "				<PARAM NAME=""Enabled"" VALUE=""-1"">" & vbcrlf
			Response.Write "				<PARAM NAME=""DataMember"" VALUE="""">" & vbcrlf

			lngRowCount = 0
			do while not rstOrderRecords.EOF
				for iLoop = 0 to (rstOrderRecords.fields.count - 1)							
				Response.Write "				<PARAM NAME=""Row(" & lngRowCount & ").Col(" & iLoop & ")"" VALUE=""" & replace(rstOrderRecords.Fields(iLoop).Value, "_", " ") & """>" & vbcrlf
				next 				
				lngRowCount = lngRowCount + 1
				rstOrderRecords.MoveNext
			loop
			Response.Write "				<PARAM NAME=""Row.Count"" VALUE=""" & lngRowCount & """>" & vbcrlf
			Response.Write "			</OBJECT>" & vbcrlf
			Response.Write "			<INPUT type='hidden' id=txtCurrentOrderID name=txtCurrentOrderID value=" & session("optionOrderID") & ">" & vbcrlf

			' Release the ADO recordset object.
			rstOrderRecords.close
			Set rstOrderRecords = nothing
		end if
	
		' Release the ADO command object.
		Set cmdOrderRecords = nothing
	end if
%>
					</TD>
					<TD width=20></TD>
				</TR>
				<TR>
					<td height=10 colspan=3>
					</td>
				</TR>
				<tr>
					<TD width=20></TD>
					<td height=10>
						<table WIDTH=100% class="invisible" CELLSPACING="0" CELLPADDING="0">
							<tr>	
								<td>&nbsp;</td>
								<td width=10>
									<input id="cmdSelectOrder" name="cmdSelectOrder" type="button" value="Select" style="WIDTH: 75px" width="75" class="btn"
									    onclick="SelectOrder()"
		                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                onfocus="try{button_onFocus(this);}catch(e){}"
		                                onblur="try{button_onBlur(this);}catch(e){}" />
									</TD>
								</td>
								<td width=40></td>
								<td width=10>
									<input id="cmdCancel" name="cmdCancel" type="button" value="Cancel" style="WIDTH: 75px" width="75" class="btn" 
									    onclick="CancelOrder()"
		                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                onfocus="try{button_onFocus(this);}catch(e){}"
		                                onblur="try{button_onBlur(this);}catch(e){}" />
								</td>
							</tr>			
						</table>
					</td>
					<TD width=20></TD>
				</tr>
				<TR>
					<td height=10 colspan=3>
					</td>
				</TR>
			</TABLE>
		</td>
	</tr>
</TABLE>
<%
	Response.Write "<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtOptionScreenID name=txtOptionScreenID value=" & session("optionScreenID") & ">" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtOptionTableID name=txtOptionTableID value=" & session("optionTableID") & ">" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtOptionViewID name=txtOptionViewID value=" & session("optionViewID") & ">" & vbcrlf
%>
</FORM>
<INPUT type='hidden' id=txtTicker name=txtTicker value=0>
<INPUT type='hidden' id=txtLastKeyFind name=txtLastKeyFind value="">

<FORM action="orderselect_Submit.asp" method=post id=frmGotoOption name=frmGotoOption>
<!--#include file="include\gotoOption.txt"-->
</FORM>

</BODY>
</HTML>

<!-- Embeds createActiveX.js script reference -->
<!--#include file="include\ctl_CreateControl.txt"-->
