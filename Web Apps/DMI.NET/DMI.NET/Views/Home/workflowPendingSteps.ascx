<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<%@ Language=VBScript %>
<!--#INCLUDE FILE="include/svrCleanup.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK href="OpenHR.css" rel=stylesheet type=text/css >
<TITLE>OpenHR Intranet</TITLE>
<meta http-equiv="X-UA-Compatible" content="IE=5">
<!--#include file="include\ctl_SetFont.txt"-->

<FORM id=frmSteps name=frmSteps style="visibility:hidden;display:none">
<%
	on error resume next

	Response.Expires = -1
	
	Dim sReferringPage
	Dim fError

	if (session("fromMenu") = 0) and (session("reset") = 1) then
		' Reset the Workflow OutOfOffice flag.
		Set cmdOutOfOffice = Server.CreateObject("ADODB.Command")
		cmdOutOfOffice.CommandText = "spASRWorkflowOutOfOfficeSet"
		cmdOutOfOffice.CommandType = 4 ' Stored Procedure
		Set cmdOutOfOffice.ActiveConnection = session("databaseConnection")

		Set prmValue = cmdOutOfOffice.CreateParameter("value",11,1) ' 11=bit, 1=input
		cmdOutOfOffice.Parameters.Append prmValue
		prmValue.value = 0

		err = 0
		cmdOutOfOffice.Execute
		set cmdOutOfOffice = nothing

		session("reset") = 0
	end if

	fWorkflowGood = true
	iStepCount = 0
	
	Set cmdDefSelRecords = Server.CreateObject("ADODB.Command")
	cmdDefSelRecords.CommandText = "spASRIntCheckPendingWorkflowSteps"
	cmdDefSelRecords.CommandType = 4 ' Stored Procedure
	Set cmdDefSelRecords.ActiveConnection = session("databaseConnection")

	err = 0
	Set rstDefSelRecords = cmdDefSelRecords.Execute

	if (err <> 0) then
		' Workflow not licensed or configured. Go to default page.
		fWorkflowGood = false
	else
		do until rstDefSelRecords.eof
			if iStepCount = 0 then
				' Add the <All> row.
				sAddString = "0" & vbtab & "<All>" & vbtab
%>
	<INPUT type='hidden' id="txtAddString_<%=iStepCount%>" name=txtAddString_<%=iStepCount%> value="<%=sAddString%>">
<%			
				
			end if

			iStepCount = iStepCount + 1
			
			sAddString = "0" & vbtab & _
				replace(rstDefSelRecords.Fields("description").Value, chr(34),"&quot;") & vbtab & _
				replace(rstDefSelRecords.Fields("url").Value, chr(34),"&quot;")
%>
	<INPUT type='hidden' id=Hidden1 name=txtAddString_<%=iStepCount%> value="<%=sAddString%>">
<%			
			rstDefSelRecords.movenext
		loop

		rstDefSelRecords.close
		Set rstDefSelRecords = nothing
	end if
							
	' Release the ADO command object.
	Set cmdDefSelRecords = nothing
%>
	<INPUT type='hidden' id=txtFromMenu name=txtFromMenu value="<%=session("fromMenu")%>">
</FORM>

<OBJECT 
	classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" 
	id="Microsoft_Licensed_Class_Manager_1_0" 
	VIEWASTEXT>
	<PARAM NAME="LPKPath" VALUE="lpks/main.lpk">
</OBJECT>

<SCRIPT FOR=window EVENT=onload LANGUAGE=JavaScript>
<!--
	var iPollCounter;
	var iPollPeriod;
	var frmRefresh;
	var sControlName;
	var sControlPrefix;

<%
	if iStepCount > 0 then
%>	

	setGridFont(frmDefSel.ssOleDBGridDefSelRecords);
	
	iPollPeriod = 100;
	iPollCounter = iPollPeriod;
	frmRefresh = window.parent.frames("pollframe").document.forms("frmHit");	

	frmDefSel.ssOleDBGridDefSelRecords.focus();
	frmDefSel.cmdCancel.focus;

	var controlCollection = frmSteps.elements;
	if (controlCollection!=null) 
	{
		for (i=0; i<controlCollection.length; i++)  
		{
			if (i==iPollCounter) 
			{			
				frmRefresh.submit();
				iPollCounter = iPollCounter + iPollPeriod;
			}

			sControlName = controlCollection.item(i).name;
			sControlPrefix = sControlName.substr(0, 13);
					
			if (sControlPrefix=="txtAddString_") 
			{
				frmDefSel.ssOleDBGridDefSelRecords.AddItem(controlCollection.item(i).value);
			}
		}
	}	

	frmRefresh.submit();

	if (frmDefSel.ssOleDBGridDefSelRecords.rows > 0) 
	{
		// Need to refresh the grid before we movefirst.
		frmDefSel.ssOleDBGridDefSelRecords.refresh();			

		// Select the top row.
		frmDefSel.ssOleDBGridDefSelRecords.MoveFirst();
		frmDefSel.ssOleDBGridDefSelRecords.SelBookmarks.Add(frmDefSel.ssOleDBGridDefSelRecords.Bookmark);
	}

	refreshControls();

	sizeColumnsToFitGrid(frmDefSel.ssOleDBGridDefSelRecords);
<%
	else
		if session("fromMenu") = 0 then
%>	
	window.parent.frames("menuframe").openPersonnelRecEdit();
<%
		end if
	end if
%>	

	window.parent.frames("menuframe").refreshMenu();
	window.parent.document.all.item("workframeset").cols = "*, 0";	

	// Little dodge to get around a browser bug that
	// does not refresh the display on all controls.
	try	
	{
		window.resizeBy(0,-1);
		window.resizeBy(0,1);
		window.resizeBy(0,-1);
		window.resizeBy(0,1);
	}
	catch(e) {}
	-->	
</SCRIPT>

<SCRIPT FOR=window EVENT=onfocus LANGUAGE=JavaScript>
<!--
	// Little dodge to get around a browser bug that
	// does not refresh the display on all controls.
	try 
	{
		window.resizeBy(0,-1);
		window.resizeBy(0,1);
		window.resizeBy(0,-1);
		window.resizeBy(0,1);
	}
	catch(e) {}
	-->
</SCRIPT>

<SCRIPT FOR=ssOleDBGridDefSelRecords EVENT=change LANGUAGE=JavaScript>
<!--
	RefreshGrid();
	-->
</script>

<SCRIPT FOR=ssOleDBGridDefSelRecords EVENT=KeyPress(iKeyAscii) LANGUAGE=JavaScript>
<!--
	if(iKeyAscii == 32)
	{
		// Space pressed. Toggle the current row value.
		ToggleCurrentRow();
	}
	-->
</script>

<SCRIPT LANGUAGE="JavaScript">
<!--
	function RefreshGrid()
	{
		var iLoop;
		var iRowIndex = frmDefSel.ssOleDBGridDefSelRecords.AddItemRowIndex(frmDefSel.ssOleDBGridDefSelRecords.Bookmark);
		var sRowTickValue;
		var fAllTicked = true;
	
		frmDefSel.ssOleDBGridDefSelRecords.Update();

		if(iRowIndex == 0)
		{
			// <All> row. Ensure all other rows match.
			varBookmark = frmDefSel.ssOleDBGridDefSelRecords.AddItemBookmark(0);
			sRowTickValue = frmDefSel.ssOleDBGridDefSelRecords.Columns("TickBox").CellText(varBookmark); 

			frmDefSel.ssOleDBGridDefSelRecords.MoveFirst();
			frmDefSel.ssOleDBGridDefSelRecords.MoveNext();
		
			for (iLoop=1; iLoop<frmDefSel.ssOleDBGridDefSelRecords.Rows; iLoop++)  
			{
				frmDefSel.ssOleDBGridDefSelRecords.Columns("TickBox").Text = sRowTickValue;
				frmDefSel.ssOleDBGridDefSelRecords.MoveNext();
			}
			frmDefSel.ssOleDBGridDefSelRecords.MoveFirst();
		}
		else
		{
			// Step row. Check if all step rows now have the same value.
			// If so, ensure the <All> row matches.
		
			for (iLoop=1; iLoop<frmDefSel.ssOleDBGridDefSelRecords.Rows; iLoop++)  
			{
				varBookmark = frmDefSel.ssOleDBGridDefSelRecords.AddItemBookmark(iLoop);
				sRowTickValue = frmDefSel.ssOleDBGridDefSelRecords.Columns("TickBox").CellText(varBookmark); 
			
				if (sRowTickValue == "0")
				{
					fAllTicked = false;
				}
			}
		
			varBookmark = frmDefSel.ssOleDBGridDefSelRecords.Bookmark;

			if (fAllTicked == true)
			{

				frmDefSel.ssOleDBGridDefSelRecords.Bookmark = frmDefSel.ssOleDBGridDefSelRecords.AddItemBookmark(0);
				frmDefSel.ssOleDBGridDefSelRecords.Columns("TickBox").Text = "-1";
			}
			else
			{
				frmDefSel.ssOleDBGridDefSelRecords.Bookmark = frmDefSel.ssOleDBGridDefSelRecords.AddItemBookmark(0);
				frmDefSel.ssOleDBGridDefSelRecords.Columns("TickBox").Text = "0";
			}

			frmDefSel.ssOleDBGridDefSelRecords.Bookmark = varBookmark;
		}
	
		refreshControls();
	}

	function ToggleCurrentRow()
	{
		if (frmDefSel.ssOleDBGridDefSelRecords.Columns("TickBox").Text == "-1")
		{
			frmDefSel.ssOleDBGridDefSelRecords.Columns("TickBox").Text = "0";
		}
		else
		{
			frmDefSel.ssOleDBGridDefSelRecords.Columns("TickBox").Text = "-1";
		}
	
		RefreshGrid();	
	}
	-->
</script>

<SCRIPT LANGUAGE="JavaScript">
<!--
	function refreshControls()
	{
		var fSomeSelected;

		fSomeSelected = SomeSelected();
		button_disable(frmDefSel.cmdRun, (fSomeSelected == false));
	}

	function SomeSelected()
	{
		var varBookmark;
		var iLoop;
	
		frmDefSel.ssOleDBGridDefSelRecords.Update()
	
		for (iLoop=1; iLoop<frmDefSel.ssOleDBGridDefSelRecords.Rows; iLoop++)  
		{
			varBookmark = frmDefSel.ssOleDBGridDefSelRecords.AddItemBookmark(iLoop);
			if (frmDefSel.ssOleDBGridDefSelRecords.Columns("TickBox").CellText(varBookmark) == "-1") {
				return(true);
			}
		}
	
		return(false);
	}

	function pausecomp(millis) 
	{
		var date = new Date();
		var curDate = null;

		do 
		{ 
			curDate = new Date(); 
		} 
		while(curDate-date < millis);
	} 

	function spawnWindow(mypage, myname, w, h, scroll) 
	{
		var newWin;
		var winl = (screen.availWidth - w) / 2;
		var wint = (screen.availHeight - h) / 2;
	
		winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',scrollbars='+scroll+',resizable';

		try
		{
			newWin = window.open(mypage, myname, winprops);

			if (parseInt(navigator.appVersion) >= 4) 
			{ 
				try 
				{
					pausecomp(300);
					newWin.focus(); 
				}
				catch(e) {}
			}
		}
		catch(e)
		{	
			try
			{
				newWin.close();
			}
			catch(e){}

			spawnWindow(mypage, myname, w, h, scroll)
		}
	}

	function setrun()
	{
		var varBookmark;
		var sForm;
		var iSelectedCount = 0;
		var sMessage;
		
		window.parent.frames("refreshframe").document.forms("frmRefresh").submit();

		try
		{
			for (iLoop=1; iLoop<frmDefSel.ssOleDBGridDefSelRecords.Rows; iLoop++)  
			{
				varBookmark = frmDefSel.ssOleDBGridDefSelRecords.AddItemBookmark(iLoop);
		      
				if (frmDefSel.ssOleDBGridDefSelRecords.Columns("TickBox").CellText(varBookmark) == "-1") {
					sForm = frmDefSel.ssOleDBGridDefSelRecords.Columns("URL").CellText(varBookmark);
					spawnWindow(sForm, "_blank", screen.availWidth, screen.availHeight,'yes');
				
					iSelectedCount = iSelectedCount + 1;
				}
			}

			if (iSelectedCount == 0) 
			{
				sMessage = "You must select a workflow step to run";
				window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sMessage,48,"OpenHR Intranet");
			}
			else
			{
				//NPG20090403 Fault 13512
				//sMessage = "Workflow forms opened successfully";
				//window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sMessage,64,"OpenHR Intranet");
<%
	if session("fromMenu") = 0 then
%>				
			window.parent.frames("menuframe").autoLoadPage("workflowPendingSteps", true);
<%
	else
%>			
			window.parent.frames("menuframe").autoLoadPage("workflowPendingSteps", false);
<%
	end if
%>			
		}
	}
	catch(e)
	{
		sMessage = "Error opening workflow forms : " + e.description;
		window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sMessage,48,"OpenHR Intranet");
	}
}

function setcancel()
{
	// Goto self-service recedit page (if Self-service user at login)
	// Otherwise load the default page.
	if (<%=session("fromMenu")%> == 0)
	{
		window.parent.frames("menuframe").openPersonnelRecEdit();
	}
else
{
		window.location.href = "default.asp";
}
}

function setrefresh()
{
	window.parent.frames("refreshframe").document.forms("frmRefresh").submit();

<%
	if session("fromMenu") = 0 then
%>				
	window.parent.frames("menuframe").autoLoadPage("workflowPendingSteps", true);
<%
	else
%>			
	window.parent.frames("menuframe").autoLoadPage("workflowPendingSteps", false);
<%
	end if
%>			
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

	function sizeColumnsToFitGrid(pctlGrid)
	{
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
		for (iLoop=0; iLoop<pctlGrid.Columns.Count; iLoop++)
		{
			if (pctlGrid.Columns.Item(iLoop).Visible == true)
			{
				if (pctlGrid.Columns.Item(iLoop).Style == 2)
				{
					iVisibleCheckboxCount = iVisibleCheckboxCount + 1;
				}
			
				iVisibleColumnCount = iVisibleColumnCount + 1;
				iLastVisibleColumn = iLoop;
			}
		}

		if (iVisibleColumnCount > 0) 
		{
			fScrollBarVisible = (pctlGrid.Rows > pctlGrid.VisibleRows);
			if (fScrollBarVisible == true)
			{
				//NPG20090403 Fault 13516
				//iUsableWidth = pctlGrid.style.pixelWidth - 20;
				iUsableWidth = findTable.clientWidth - 20;
			}
			else
			{
				//NPG20090403 Fault 13516
				//iUsableWidth = pctlGrid.style.pixelWidth;
				iUsableWidth = findTable.clientWidth;
			}
		
			iNewColWidth = (iUsableWidth - (iVisibleCheckboxCount * iCheckboxWidth)) / (iVisibleColumnCount - iVisibleCheckboxCount);
			if (iNewColWidth < iMinWidth) 
			{
				iNewColWidth = iMinWidth;
			}
		
			for (iLoop=0; iLoop<iLastVisibleColumn; iLoop++)
			{
				if (pctlGrid.Columns.Item(iLoop).Visible == true)
				{
					if (pctlGrid.Columns.Item(iLoop).Style == 2)
					{
						pctlGrid.Columns(iLoop).Width = iCheckboxWidth;
					}
					else
					{
						pctlGrid.Columns.Item(iLoop).Width = iNewColWidth;
					}
					iUsedWidth = iUsedWidth + pctlGrid.Columns.Item(iLoop).Width;
				}
			}

			iNewColWidth = iUsableWidth - iUsedWidth - 2;
			if (iNewColWidth < iMinWidth) 
			{
				iNewColWidth = iMinWidth;
			}
			pctlGrid.Columns.Item(iLastVisibleColumn).Width = iNewColWidth;
		}
	}
	-->
</script>
<!--#INCLUDE FILE="include/ctl_SetStyles.txt" -->
</HEAD>

<BODY <%=session("BodyTag")%>>

<form name="frmDefSel" method="post" id="frmDefSel">

<%if (fWorkflowGood = true) or (session("fromMenu") = 1) then%>
<%	if iStepCount > 0 then%>
<table align=center class="outline" cellPadding=5 cellSpacing=0 height="100%" width=100%>
	<TR>
		<TD>
			<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0" >
				<tr> 
					<td colspan=5 align=center height=10>
						<H3>
							Pending Workflow Steps
						</H3>
					</td>
				</tr>
				
				<tr> 
					<td width=20>&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td width=100%>
						<table height=100% width=100% class="invisible" cellspacing=0 cellpadding=0 id="findTable">
							<tr>
								<td width=100%>
									<OBJECT classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" 
										id=ssOleDBGridDefSelRecords 
										name=ssOleDBGridDefselRecords 
										codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" 
										style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:100%">
										<PARAM NAME="ScrollBars" VALUE="4">
										<PARAM NAME="_Version" VALUE="196616">
										<PARAM NAME="DataMode" VALUE="2">
										<PARAM NAME="Cols" VALUE="0">
										<PARAM NAME="Rows" VALUE="0">
										<PARAM NAME="BorderStyle" VALUE="1">
										<PARAM NAME="RecordSelectors" VALUE="0">
										<PARAM NAME="GroupHeaders" VALUE="0">
										<PARAM NAME="ColumnHeaders" VALUE="0">
										<PARAM NAME="GroupHeadLines" VALUE="0">
										<PARAM NAME="HeadLines" VALUE="0">
										<PARAM NAME="FieldDelimiter" VALUE="(None)">
										<PARAM NAME="FieldSeparator" VALUE="(Tab)">
										<PARAM NAME="Col.Count" VALUE="3">
										<PARAM NAME="stylesets.count" VALUE="0">

										<PARAM NAME="TagVariant" VALUE="EMPTY">
										<PARAM NAME="UseGroups" VALUE="0">
										<PARAM NAME="HeadFont3D" VALUE="0">
										<PARAM NAME="Font3D" VALUE="0">
										<PARAM NAME="DividerType" VALUE="3">
										<PARAM NAME="DividerStyle" VALUE="1">
										<PARAM NAME="DefColWidth" VALUE="0">
										<PARAM NAME="BevelColorScheme" VALUE="2">
										<PARAM NAME="BevelColorFrame" VALUE="0">
										<PARAM NAME="BevelColorHighlight" VALUE="0">
										<PARAM NAME="BevelColorShadow" VALUE="0">
										<PARAM NAME="BevelColorFace" VALUE="0">
										<PARAM NAME="CheckBox3D" VALUE="0">
										<PARAM NAME="AllowAddNew" VALUE="0">
										<PARAM NAME="AllowDelete" VALUE="0">
										<PARAM NAME="AllowUpdate" VALUE="-1">
										<PARAM NAME="MultiLine" VALUE="0">
										<PARAM NAME="ActiveCellStyleSet" VALUE="">
										<PARAM NAME="RowSelectionStyle" VALUE="0">
										<PARAM NAME="AllowRowSizing" VALUE="0">
										<PARAM NAME="AllowGroupSizing" VALUE="0">
										<PARAM NAME="AllowColumnSizing" VALUE="0">
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
										<PARAM NAME="BackColorEven" VALUE="0">
										<PARAM NAME="BackColorOdd" VALUE="0">
										<PARAM NAME="Levels" VALUE="1">
										<PARAM NAME="RowHeight" VALUE="503">
										<PARAM NAME="ExtraHeight" VALUE="0">
										<PARAM NAME="ActiveRowStyleSet" VALUE="">
										<PARAM NAME="CaptionAlignment" VALUE="2">
										<PARAM NAME="SplitterPos" VALUE="0">
										<PARAM NAME="SplitterVisible" VALUE="0">
										<PARAM NAME="Columns.Count" VALUE="3">

										<PARAM NAME="Columns(0).Width" VALUE="1000">
										<PARAM NAME="Columns(0).Visible" VALUE="-1">
										<PARAM NAME="Columns(0).Columns.Count" VALUE="1">
										<PARAM NAME="Columns(0).Caption" VALUE="">
										<PARAM NAME="Columns(0).Name" VALUE="TickBox">			
										<PARAM NAME="Columns(0).Alignment" VALUE="0">
										<PARAM NAME="Columns(0).CaptionAlignment" VALUE="3">
										<PARAM NAME="Columns(0).Bound" VALUE="0">
										<PARAM NAME="Columns(0).AllowSizing" VALUE="1">
										<PARAM NAME="Columns(0).DataField" VALUE="Column 0">
										<PARAM NAME="Columns(0).DataType" VALUE="8">
										<PARAM NAME="Columns(0).Level" VALUE="0">
										<PARAM NAME="Columns(0).NumberFormat" VALUE="">			
										<PARAM NAME="Columns(0).Case" VALUE="0">
										<PARAM NAME="Columns(0).FieldLen" VALUE="4096">
										<PARAM NAME="Columns(0).VertScrollBar" VALUE="0">
										<PARAM NAME="Columns(0).Locked" VALUE="0">			
										<PARAM NAME="Columns(0).Style" VALUE="2">
										<PARAM NAME="Columns(0).ButtonsAlways" VALUE="0">
										<PARAM NAME="Columns(0).RowCount" VALUE="0">
										<PARAM NAME="Columns(0).ColCount" VALUE="1">
										<PARAM NAME="Columns(0).HasHeadForeColor" VALUE="0">
										<PARAM NAME="Columns(0).HasHeadBackColor" VALUE="0">
										<PARAM NAME="Columns(0).HasForeColor" VALUE="0">
										<PARAM NAME="Columns(0).HasBackColor" VALUE="0">
										<PARAM NAME="Columns(0).HeadForeColor" VALUE="0">
										<PARAM NAME="Columns(0).HeadBackColor" VALUE="0">
										<PARAM NAME="Columns(0).ForeColor" VALUE="0">
										<PARAM NAME="Columns(0).BackColor" VALUE="0">
										<PARAM NAME="Columns(0).HeadStyleSet" VALUE="">
										<PARAM NAME="Columns(0).StyleSet" VALUE="">
										<PARAM NAME="Columns(0).Nullable" VALUE="1">
										<PARAM NAME="Columns(0).Mask" VALUE="">
										<PARAM NAME="Columns(0).PromptInclude" VALUE="0">
										<PARAM NAME="Columns(0).ClipMode" VALUE="0">
										<PARAM NAME="Columns(0).PromptChar" VALUE="95">
										
										<PARAM NAME="Columns(1).Width" VALUE="1000">
										<PARAM NAME="Columns(1).Visible" VALUE="-1">
										<PARAM NAME="Columns(1).Columns.Count" VALUE="1">
										<PARAM NAME="Columns(1).Caption" VALUE="">
										<PARAM NAME="Columns(1).Name" VALUE="Description">			
										<PARAM NAME="Columns(1).Alignment" VALUE="0">
										<PARAM NAME="Columns(1).CaptionAlignment" VALUE="3">
										<PARAM NAME="Columns(1).Bound" VALUE="0">
										<PARAM NAME="Columns(1).AllowSizing" VALUE="1">
										<PARAM NAME="Columns(1).DataField" VALUE="Column 0">
										<PARAM NAME="Columns(1).DataType" VALUE="8">
										<PARAM NAME="Columns(1).Level" VALUE="0">
										<PARAM NAME="Columns(1).NumberFormat" VALUE="">			
										<PARAM NAME="Columns(1).Case" VALUE="0">
										<PARAM NAME="Columns(1).FieldLen" VALUE="4096">
										<PARAM NAME="Columns(1).VertScrollBar" VALUE="0">
										<PARAM NAME="Columns(1).Locked" VALUE="-1">			
										<PARAM NAME="Columns(1).Style" VALUE="0">
										<PARAM NAME="Columns(1).ButtonsAlways" VALUE="0">
										<PARAM NAME="Columns(1).RowCount" VALUE="0">
										<PARAM NAME="Columns(1).ColCount" VALUE="1">
										<PARAM NAME="Columns(1).HasHeadForeColor" VALUE="0">
										<PARAM NAME="Columns(1).HasHeadBackColor" VALUE="0">
										<PARAM NAME="Columns(1).HasForeColor" VALUE="0">
										<PARAM NAME="Columns(1).HasBackColor" VALUE="0">
										<PARAM NAME="Columns(1).HeadForeColor" VALUE="0">
										<PARAM NAME="Columns(1).HeadBackColor" VALUE="0">
										<PARAM NAME="Columns(1).ForeColor" VALUE="0">
										<PARAM NAME="Columns(1).BackColor" VALUE="0">
										<PARAM NAME="Columns(1).HeadStyleSet" VALUE="">
										<PARAM NAME="Columns(1).StyleSet" VALUE="">
										<PARAM NAME="Columns(1).Nullable" VALUE="1">
										<PARAM NAME="Columns(1).Mask" VALUE="">
										<PARAM NAME="Columns(1).PromptInclude" VALUE="0">
										<PARAM NAME="Columns(1).ClipMode" VALUE="0">
										<PARAM NAME="Columns(1).PromptChar" VALUE="95">
										
										<PARAM NAME="Columns(2).Width" VALUE="0">
										<PARAM NAME="Columns(2).Visible" VALUE="0">
										<PARAM NAME="Columns(2).Columns.Count" VALUE="1">
										<PARAM NAME="Columns(2).Caption" VALUE="">
										<PARAM NAME="Columns(2).Name" VALUE="URL">			
										<PARAM NAME="Columns(2).Alignment" VALUE="0">
										<PARAM NAME="Columns(2).CaptionAlignment" VALUE="3">
										<PARAM NAME="Columns(2).Bound" VALUE="0">
										<PARAM NAME="Columns(2).AllowSizing" VALUE="1">
										<PARAM NAME="Columns(2).DataField" VALUE="Column 0">
										<PARAM NAME="Columns(2).DataType" VALUE="8">
										<PARAM NAME="Columns(2).Level" VALUE="0">
										<PARAM NAME="Columns(2).NumberFormat" VALUE="">			
										<PARAM NAME="Columns(2).Case" VALUE="0">
										<PARAM NAME="Columns(2).FieldLen" VALUE="4096">
										<PARAM NAME="Columns(2).VertScrollBar" VALUE="0">
										<PARAM NAME="Columns(2).Locked" VALUE="0">			
										<PARAM NAME="Columns(2).Style" VALUE="0">
										<PARAM NAME="Columns(2).ButtonsAlways" VALUE="0">
										<PARAM NAME="Columns(2).RowCount" VALUE="0">
										<PARAM NAME="Columns(2).ColCount" VALUE="1">
										<PARAM NAME="Columns(2).HasHeadForeColor" VALUE="0">
										<PARAM NAME="Columns(2).HasHeadBackColor" VALUE="0">
										<PARAM NAME="Columns(2).HasForeColor" VALUE="0">
										<PARAM NAME="Columns(2).HasBackColor" VALUE="0">
										<PARAM NAME="Columns(2).HeadForeColor" VALUE="0">
										<PARAM NAME="Columns(2).HeadBackColor" VALUE="0">
										<PARAM NAME="Columns(2).ForeColor" VALUE="0">
										<PARAM NAME="Columns(2).BackColor" VALUE="0">
										<PARAM NAME="Columns(2).HeadStyleSet" VALUE="">
										<PARAM NAME="Columns(2).StyleSet" VALUE="">
										<PARAM NAME="Columns(2).Nullable" VALUE="1">
										<PARAM NAME="Columns(2).Mask" VALUE="">
										<PARAM NAME="Columns(2).PromptInclude" VALUE="0">
										<PARAM NAME="Columns(2).ClipMode" VALUE="0">
										<PARAM NAME="Columns(2).PromptChar" VALUE="95">

										<PARAM NAME="UseDefaults" VALUE="-1">
										<PARAM NAME="TabNavigation" VALUE="1">
										<PARAM NAME="_ExtentX" VALUE="17330">
										<PARAM NAME="_ExtentY" VALUE="1323">
										<PARAM NAME="_StockProps" VALUE="79">
										<PARAM NAME="Caption" VALUE="">
										<PARAM NAME="ForeColor" VALUE="0">
										<PARAM NAME="BackColor" VALUE="0">
										<PARAM NAME="Enabled" VALUE="-1">
										<PARAM NAME="DataMember" VALUE="">
										<PARAM NAME="Row.Count" VALUE="0">
									</OBJECT>
								</td>
							</tr>
						</table>							
					</td>
					
					<td width=20>&nbsp;&nbsp;&nbsp;&nbsp;</td>
	        
			        <td width=80> 
						<table height=100% class="invisible" cellspacing=0 cellpadding=0>
							<tr>
								<td>
									<input type="button" name=cmdRefresh value="Refresh" style="WIDTH: 80px" width="80" id=cmdRefresh class="btn"
									    onclick="setrefresh();"
					                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
			                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
			                            onfocus="try{button_onFocus(this);}catch(e){}"
			                            onblur="try{button_onBlur(this);}catch(e){}" />
								</td>
							</tr>
							<tr height=100%>
								<td></td>
							</tr>
							<tr>
								<td>
									<input type="button" name=cmdRun value="Run" style="WIDTH: 80px" width="80" id=cmdRun class="btn"
									    onclick="setrun();"
					                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
			                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
			                            onfocus="try{button_onFocus(this);}catch(e){}"
			                            onblur="try{button_onBlur(this);}catch(e){}" />
								</td>
							</tr>
							<tr height=10>
								<td></td>
							</tr>
							<tr>
								<td>
									<input type="button" name="cmdCancel" value=Cancel style="WIDTH: 80px" width="80" class="btn"
										onclick="setcancel()" 
					                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
			                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
			                            onfocus="try{button_onFocus(this);}catch(e){}"
			                            onblur="try{button_onBlur(this);}catch(e){}" />
							  </td>
							</tr>
						</table>	
					</td>
					<td width=20>&nbsp;&nbsp;&nbsp;&nbsp;</td>
				</tr>
				<tr> 
					<td colspan=5 align=center height=10>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<%							
	else
		if (session("fromMenu") = 1) then
			if fWorkflowGood = true then
				' Display message saying no pending steps.
				sMessage = "No pending workflow steps"
			else
				' Display error message.
				sMessage = "Error getting the pending workflow steps"
			end if
%>
<table align=center class="outline" cellPadding=5 cellSpacing=0>
	<TR>
        <td width=20></td> 
		<TD>
			<table class="invisible" cellspacing="0" cellpadding="0">
                <tr> 
			        <td height=10></td>
			    </tr>

			    <tr> 
			        <td align=center> 
			            <H3>Pending Workflow Steps</H3>
			        </td>
			    </tr>
			  
			    <tr> 
			        <td align=center> 
                        <%=sMessage%>
			        </td>
			    </tr>
			  
			    <tr> 
			        <td height=20></td>
			    </tr>

			    <tr> 
			        <td height=10 align=center> 
		                <input id="cmdOK" name="cmdOK" type=button class="btn" value="OK" style="WIDTH: 75px" width="75" 
        		            onclick="setcancel()"
		                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
                    </td>
			    </tr>

                <tr> 
			        <td height=10></td>
			    </tr>
			</table>
        </td>
        <td width=20></td> 
    </tr>
</table>
<%			
		end if
	end if
end if
%>							
</form>

<FORM action="default_Submit.asp" method=post id=frmGoto name=frmGoto style="visibility:hidden;display:none">
<!--#include file="include\gotoWork.txt"-->
</FORM>

</BODY>
</html>

<!-- Embeds createActiveX.js script reference -->
<!--#include file="include\ctl_CreateControl.txt"-->