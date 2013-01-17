<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<%
	'This section of script is used for saving the new purge criteria.
	dim bDoesPurge
	dim sPeriod
	dim iFrequency
	dim sSQL
    Dim cmdPurge
    Dim prmPeriod
    Dim prmFrequency
    
	bDoesPurge = trim(Request.Form("txtDoesPurge"))
	sPeriod = Request.Form("txtPurgePeriod")
	iFrequency = Request.Form("txtPurgeFrequency")
	
	if bDoesPurge <> vbNullString then
		
		' Delete old purge information to the database
        cmdPurge = Server.CreateObject("ADODB.Command")
		cmdPurge.CommandText = "spASRIntClearEventLogPurge"
		cmdPurge.CommandType = 4 ' Stored procedure.
        cmdPurge.ActiveConnection = Session("databaseConnection")
		err.clear()
		cmdPurge.Execute 
        cmdPurge = Nothing
 
		if bDoesPurge = 1 then
			' Insert the new purge criteria
            cmdPurge = Server.CreateObject("ADODB.Command")
			cmdPurge.CommandText = "spASRIntSetEventLogPurge"
			cmdPurge.CommandType = 4 ' Stored procedure.
            cmdPurge.ActiveConnection = Session("databaseConnection")
		
			prmPeriod = cmdPurge.CreateParameter("period",200,1,8000) ' 200=varchar, 1=input, 8000=size
            cmdPurge.Parameters.Append(prmPeriod)
			prmPeriod.value = cstr(sPeriod)

			prmFrequency = cmdPurge.CreateParameter("frequency",3,1) ' 3=integer, 1=input
            cmdPurge.Parameters.Append(prmFrequency)
			prmFrequency.value = cleanNumeric(clng(iFrequency))
			
			err.clear()
			cmdPurge.Execute 
            cmdPurge = Nothing

			Session("showPurgeMessage") = 1
		else
			Session("showPurgeMessage") = 0 
		end if
	end if
	
    bDoesPurge = Nothing
    sPeriod = Nothing
    iFrequency = Nothing
    sSQL = Nothing
%>

<%
	'This section of script is used for deleting Event Log records according to the selection on the Delete screen. 
	dim iDeleteSelection
	dim sSelectedEventIDs 
	dim cmdDelete
    Dim bHasViewAllPermission
    Dim prmEventIDs
    Dim prmType
    Dim prmCanViewAll
	
	iDeleteSelection = Request.Form("txtDeleteSel") 
	sSelectedEventIDs = Request.Form("txtSelectedIDs")
	bHasViewAllPermission = Request.Form("txtViewAllPerm")

	if iDeleteSelection <> vbNullString then
		iDeleteSelection = CInt(iDeleteSelection)
		
        cmdDelete = Server.CreateObject("ADODB.Command")
		cmdDelete.CommandText = "spASRIntDeleteEventLogRecords"
		cmdDelete.CommandType = 4 ' Stored procedure.
        cmdDelete.ActiveConnection = Session("databaseConnection")
		
		prmType = cmdDelete.CreateParameter("type",3,1) ' 3=integer, 1=input
        cmdDelete.Parameters.Append(prmType)
		prmType.value = cleanNumeric(clng(iDeleteSelection))

		prmEventIDs = cmdDelete.CreateParameter("eventIDs",200,1,8000) ' 200=varchar, 1=input, 8000=size
        cmdDelete.Parameters.Append(prmEventIDs)
		prmEventIDs.value = cstr(sSelectedEventIDs)

		prmCanViewAll = cmdDelete.CreateParameter("canViewAll",11,1) ' 11=bit, 1=input
        cmdDelete.Parameters.Append(prmCanViewAll)
		prmCanViewAll.value = cleanBoolean(cbool(bHasViewAllPermission))

		err.clear()
		cmdDelete.Execute 
        cmdDelete = Nothing
	end if

    iDeleteSelection = Nothing
    sSelectedEventIDs = Nothing
	cmdDelete = nothing
    bHasViewAllPermission = Nothing
%>


<OBJECT 
	classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" 
	id="Microsoft_Licensed_Class_Manager_1_0" 
	VIEWASTEXT>
	<PARAM NAME="LPKPath" VALUE="lpks/main.lpk">
</OBJECT>

<script type="text/javascript">
    function EventLog_window_onload() {

        //window.parent.document.all.item("workframeset").cols = "*, 0";
        $("#workframe").attr("data-framesource", "EVENTLOG");
        
        frmLog.cboUsername.style.color = 'white';
        frmLog.cboType.style.color = 'white';
        frmLog.cboMode.style.color = 'white';
        frmLog.cboStatus.style.color = 'white';
	
        setGridFont(frmLog.ssOleDBGridEventLog);
	
        var fOK
        fOK = true;	

        var sErrMsg = frmUseful.txtErrorDescription.value;
        if (sErrMsg.length > 0) {
            fOK = false;
            OpenHR.messageBox(sErrMsg);
            window.parent.location.replace("login");
        }
	
        if (fOK == true) {
            // Get menu to refresh the menu.
            menu_refreshMenu();		  
        }
	
        frmLog.txtELDeletePermission.value =  menu_GetItemValue("txtSysPerm_EVENTLOG_DELETE");
        frmLog.txtELViewAllPermission.value =menu_GetItemValue("txtSysPerm_EVENTLOG_VIEWALL");
        frmLog.txtELPurgePermission.value = menu_GetItemValue("txtSysPerm_EVENTLOG_PURGE");
        frmLog.txtELEmailPermission.value = menu_GetItemValue("txtSysPerm_EVENTLOG_EMAIL");

        refreshUsers();

        // Little dodge to get around a browser bug that
        // does not refresh the display on all controls.
        try
        {
            window.resizeBy(0,-1);
            window.resizeBy(0,1);
        }
        catch(e) {}
    }

</script>


    

<script type="text/javascript" id=scptGeneralFunctions>
<!--
	
    function moveRecord(psMovement)
    {
        var frmGetData = OpenHR.getForm("dataframe", "frmGetData");
        var frmData = OpenHR.getForm("dataframe", "frmData");

        frmGetData.txtELAction.value = psMovement;
        frmGetData.txtELCurrRecCount.value = frmData.txtELCurrentRecCount.value;
        frmGetData.txtEL1stRecPos.value = frmData.txtELFirstRecPos.value;
	
        refreshGrid();
	
        return;
    }
	
    function refreshStatusBar()
    {
        var sText;
        var sOrderColumn;
        var sOrderOrder;
        var sRecords;
        var frmData = OpenHR.getForm("dataframe", "frmData");
	
        sOrderColumn = frmLog.ssOleDBGridEventLog.Columns(parseInt(frmLog.txtELSortColumnIndex.value)).caption;
        sOrderOrder = frmLog.txtELOrderOrder.value;
        sRecords = frmData.txtELTotalRecordCount.value;

        if (sRecords == 0)
        {
            sText = '0 Records';
        }
        else if (sRecords == 1)
        {
            sText = '1 Record'
        }
        else
        {
            sText =  sRecords + ' Records Sorted by ' + sOrderColumn; 
	
            if (sOrderOrder == 'ASC')
            {
                sText = sText + ' in Ascending order';
            }
            else
            {
                sText = sText + ' in Descending order';
            }
        }

        document.getElementById('sbEventLog').innerText = sText;
	
        if (sRecords > 0) 
        {
            iStartPosition = parseInt(frmData.txtELFirstRecPos.value);
            iEndPosition = iStartPosition - 1 + parseInt(frmData.txtELCurrentRecCount.value);
						
            sCaption = "Records " +
                        iStartPosition +
                        " to " +
                        iEndPosition +
                        " of " +
                        sRecords;
        }
        else 
        {
            sCaption = "No Records";
        }
	
        if (frmLog.txtELViewAllPermission.value == 0)
        {
            sCaption = sCaption + "     [Viewing own entries only]";
        }
	
        menu_setVisibleMenuItem("mnutoolRecordPosition", true);
        menu_SetmnutoolRecordPositionCaption(sCaption);
	
        return true;
    }
	
    function loadEventLog()
    {
        var i;
        var sAddLine;

        var iPollCounter;
        var iPollPeriod;

        iPollPeriod = 100;
        iPollCounter = iPollPeriod;

        var frmUtilDefForm = OpenHR.getForm("dataframe", "frmData");
        var dataCollection = frmUtilDefForm.elements;
	
        var frmRefresh;
        frmRefresh = OpenHR.getForm("pollframe","frmHit");
	
        if (dataCollection!=null) 
        {            
            frmLog.ssOleDBGridEventLog.focus();
            frmLog.ssOleDBGridEventLog.Redraw = false;
            if(frmLog.ssOleDBGridEventLog.Rows > 0)
            {
                frmLog.ssOleDBGridEventLog.RemoveAll();
            }
		
            for (i=0; i<dataCollection.length; i++)  
            {
			
                if (i==iPollCounter) {			
                    //TODO
                    //frmRefresh.submit();
                    iPollCounter = iPollCounter + iPollPeriod;
                }
				
                sControlName = dataCollection.item(i).name;
                sControlPrefix = sControlName.substr(0, 13);
			
                if (sControlPrefix=="txtAddString_") 
                {
                    frmLog.ssOleDBGridEventLog.AddItem(dataCollection.item(i).value);
                }
            }
			
            frmLog.ssOleDBGridEventLog.Redraw = true;
            //TODO
            //frmRefresh.submit();

            if (frmLog.ssOleDBGridEventLog.Rows > 0)
            {
                frmLog.ssOleDBGridEventLog.SelBookmarks.RemoveAll();
                frmLog.ssOleDBGridEventLog.MoveFirst();
                frmLog.ssOleDBGridEventLog.SelBookmarks.Add(frmLog.ssOleDBGridEventLog.Bookmark);
            }
        }	

        frmLog.cboUsername.style.color = 'black';
        frmLog.cboType.style.color = 'black';
        frmLog.cboMode.style.color = 'black';
        frmLog.cboStatus.style.color = 'black';

        refreshButtons();

        //Set the event log loaded flag, used in the menu
        frmLog.txtELLoaded.value = 1;
	
        // Get menu to refresh the menu.
        menu_refreshMenu();
	
        refreshStatusBar()
	
        if (frmPurge.txtShowPurgeMSG.value == 1)
        {
            OpenHR.messageBox("Purge completed.",64,"Event Log");
            frmPurge.txtShowPurgeMSG.value = 0;
        }
    }

    function filterSQL()
    {
        var sSQL = new String(""); 
	
        if (frmLog.cboUsername.options[frmLog.cboUsername.selectedIndex].value != -1)
        {
            var sUsername = new String(frmLog.cboUsername.options[frmLog.cboUsername.selectedIndex].value);
            sSQL = sSQL + " LOWER(Username) = '" + sUsername.toLowerCase() + "' ";   
        }
		
        if (frmLog.cboType.options[frmLog.cboType.selectedIndex].value != -1)
        {
            if (sSQL.length > 0)
            {
                sSQL = sSQL + " AND ";
            }
            sSQL = sSQL + " Type = " + frmLog.cboType.options[frmLog.cboType.selectedIndex].value + " ";
        }
	
        if (frmLog.cboStatus.options[frmLog.cboStatus.selectedIndex].value != -1)
        {
            if (sSQL.length > 0)
            {
                sSQL = sSQL + " AND ";
            }
            sSQL = sSQL + "Status = " + frmLog.cboStatus.options[frmLog.cboStatus.selectedIndex].value + " ";
        }

        if (frmLog.cboMode.options[frmLog.cboMode.selectedIndex].value != -1)
        {
            if (sSQL.length > 0)
            {
                sSQL = sSQL + " AND ";
            }
            sSQL = sSQL + " Mode = " + frmLog.cboMode.options[frmLog.cboMode.selectedIndex].value + " ";
        }
	
        return sSQL;
    }
	
    function refreshGrid()
    {
        var frmGetDataForm = OpenHR.getForm("dataframe", "frmGetData");
        frmGetDataForm.txtAction.value = "LOADEVENTLOG";
	
        frmGetDataForm.txtELFilterUser.value = frmLog.cboUsername.options[frmLog.cboUsername.selectedIndex].value; 
        frmGetDataForm.txtELFilterType.value = frmLog.cboType.options[frmLog.cboType.selectedIndex].value; 
        frmGetDataForm.txtELFilterStatus.value = frmLog.cboStatus.options[frmLog.cboStatus.selectedIndex].value; 
        frmGetDataForm.txtELFilterMode.value = frmLog.cboMode.options[frmLog.cboMode.selectedIndex].value; 
        frmGetDataForm.txtELOrderColumn.value = frmLog.txtELOrderColumn.value; 
        frmGetDataForm.txtELOrderOrder.value = frmLog.txtELOrderOrder.value;

        refreshButtons();	
        OpenHR.submitForm(frmGetDataForm);

    }

    function viewEvent()
    {
        var sURL;
	
        if (frmLog.ssOleDBGridEventLog.Rows > 0 && frmLog.ssOleDBGridEventLog.SelBookmarks.Count == 1)
        {
            frmDetails.txtEventID.value = frmLog.ssOleDBGridEventLog.Columns(0).text;
		
            frmDetails.txtEventName.value = frmLog.ssOleDBGridEventLog.Columns(5).text;
            frmDetails.txtEventMode.value = frmLog.ssOleDBGridEventLog.Columns(7).text;
		
            frmDetails.txtEventStartTime.value = frmLog.ssOleDBGridEventLog.Columns(1).text;
            frmDetails.txtEventEndTime.value = frmLog.ssOleDBGridEventLog.Columns(2).text;
            frmDetails.txtEventDuration.value = frmLog.ssOleDBGridEventLog.Columns(3).text;
		
            frmDetails.txtEventType.value = frmLog.ssOleDBGridEventLog.Columns(4).text;
            frmDetails.txtEventStatus.value = frmLog.ssOleDBGridEventLog.Columns(6).text;
            frmDetails.txtEventUser.value = frmLog.ssOleDBGridEventLog.Columns(8).text;

            frmDetails.txtEventSuccessCount.value = frmLog.ssOleDBGridEventLog.Columns(12).text;
            frmDetails.txtEventFailCount.value = frmLog.ssOleDBGridEventLog.Columns(13).text;
		
            frmDetails.txtEventBatchName.value = frmLog.ssOleDBGridEventLog.Columns("BatchName").text;
            frmDetails.txtEventBatchJobID.value = frmLog.ssOleDBGridEventLog.Columns("BatchJobID").text;
            frmDetails.txtEventBatchRunID.value = frmLog.ssOleDBGridEventLog.Columns("BatchRunID").text;
		
            frmDetails.txtEmailPermission.value = frmLog.txtELEmailPermission.value;
	
            sURL = "eventLogDetails" +
                "?txtEventID=" + frmDetails.txtEventID.value +
                "&txtEventName=" + escape(frmDetails.txtEventName.value) + 
                "&txtEventMode=" + escape(frmDetails.txtEventMode.value) +
                "&txtEventStartTime=" + frmDetails.txtEventStartTime.value +
                "&txtEventEndTime=" + frmDetails.txtEventEndTime.value +
                "&txtEventDuration=" + frmDetails.txtEventDuration.value +
                "&txtEventType=" + escape(frmDetails.txtEventType.value) +
                "&txtEventStatus=" + escape(frmDetails.txtEventStatus.value) +
                "&txtEventUser=" + escape(frmDetails.txtEventUser.value) +
                "&txtEventSuccessCount=" + frmDetails.txtEventSuccessCount.value +
                "&txtEventFailCount=" + frmDetails.txtEventFailCount.value +
                "&txtEventBatchName=" + escape(frmDetails.txtEventBatchName.value) +
                "&txtEventBatchJobID=" + frmDetails.txtEventBatchJobID.value +
                "&txtEventBatchRunID=" + frmDetails.txtEventBatchRunID.value +
                "&txtEmailPermission=" + escape(frmDetails.txtEmailPermission.value);

            openDialog(sURL, 750, 450);
        }
	
        refreshButtons();
    }

    function deleteEvent()
    {
        var sURL;
		
        sURL = "eventLogSelection" +
			"?txtEventID=" + frmDetails.txtEventID.value +
			"&txtEventName=" + escape(frmDetails.txtEventName.value) + 
			"&txtEventMode=" + escape(frmDetails.txtEventMode.value) +
			"&txtEventStartTime=" + frmDetails.txtEventStartTime.value +
			"&txtEventEndTime=" + frmDetails.txtEventEndTime.value +
			"&txtEventDuration=" + frmDetails.txtEventDuration.value +
			"&txtEventType=" + escape(frmDetails.txtEventType.value) +
			"&txtEventStatus=" + escape(frmDetails.txtEventStatus.value) +
			"&txtEventUser=" + escape(frmDetails.txtEventUser.value) +
			"&txtEventSuccessCount=" + frmDetails.txtEventSuccessCount.value +
			"&txtEventFailCount=" + frmDetails.txtEventFailCount.value +
			"&txtEventBatchName=" + escape(frmDetails.txtEventBatchName.value) +
			"&txtEventBatchJobID=" + frmDetails.txtEventBatchJobID.value +
			"&txtEventBatchRunID=" + frmDetails.txtEventBatchRunID.value +
			"&txtEmailPermission=" + escape(frmDetails.txtEmailPermission.value);

        openDialog(sURL, 500,225);
    }

    function purgeEvent()
    {
        var sURL;
		
        sURL = "EventLogPurge" +
			"?txtEventID=" + frmDetails.txtEventID.value +
			"&txtEventName=" + escape(frmDetails.txtEventName.value) + 
			"&txtEventMode=" + escape(frmDetails.txtEventMode.value) +
			"&txtEventStartTime=" + frmDetails.txtEventStartTime.value +
			"&txtEventEndTime=" + frmDetails.txtEventEndTime.value +
			"&txtEventDuration=" + frmDetails.txtEventDuration.value +
			"&txtEventType=" + escape(frmDetails.txtEventType.value) +
			"&txtEventStatus=" + escape(frmDetails.txtEventStatus.value) +
			"&txtEventUser=" + escape(frmDetails.txtEventUser.value) +
			"&txtEventSuccessCount=" + frmDetails.txtEventSuccessCount.value +
			"&txtEventFailCount=" + frmDetails.txtEventFailCount.value +
			"&txtEventBatchName=" + escape(frmDetails.txtEventBatchName.value) +
			"&txtEventBatchJobID=" + frmDetails.txtEventBatchJobID.value +
			"&txtEventBatchRunID=" + frmDetails.txtEventBatchRunID.value +
			"&txtEmailPermission=" + escape(frmDetails.txtEmailPermission.value);

        openDialog(sURL, 500, 180);

    }

    function emailEvent()
    {
        var eventID;
        var sEventList = new String("");
        var sURL;
	
        //populate the txtSelectedIDs list
        for (var i=0; i<frmLog.ssOleDBGridEventLog.SelBookmarks.Count; i++)
        {
            eventID = frmLog.ssOleDBGridEventLog.Columns("ID").CellText(frmLog.ssOleDBGridEventLog.SelBookmarks(i));
		
            sEventList = sEventList + eventID + ",";
        }
	
        frmEmail.txtSelectedEventIDs.value = sEventList.substr(0,sEventList.length-1);
		
        sURL = "emailSelection" +
            "?txtSelectedEventIDs=" + frmEmail.txtSelectedEventIDs.value +
            "&txtFromMain=" + frmEmail.txtFromMain.value + 
            "&txtEmailOrderColumn=" + frmLog.txtELOrderColumn.value + 
            "&txtEmailOrderOrder=" + frmLog.txtELOrderOrder.value;

        openDialog(sURL, 435, 350);
    }

    function refreshButtons()
    {
        with (frmLog.ssOleDBGridEventLog)
        {
            if (Rows > 0)
            {
                button_disable(frmLog.cmdView, false);
            }
            else
            {
                button_disable(frmLog.cmdView, true);
            }
		
					
            if ((frmLog.txtELPurgePermission.value == 1))
            {
                button_disable(frmLog.cmdPurge, false);
            }
            else
            {
                button_disable(frmLog.cmdPurge, true);
            }

            if ((Rows > 0) && (frmLog.txtELDeletePermission.value == 1))
            {
                button_disable(frmLog.cmdDelete, false);
            }
            else
            {
                button_disable(frmLog.cmdDelete, true);
            }
			
            if ((SelBookmarks.Count > 0) && (frmLog.txtELEmailPermission.value == 1))
            {
                button_disable(frmLog.cmdEmail, false);
            }
            else
            {
                button_disable(frmLog.cmdEmail, true);
            }
        }
    }

    function refreshUsers()
    {        
        // Get the columns/calcs for the current table selection.
        var frmGetDataForm = OpenHR.getForm("dataframe", "frmGetData");
        frmGetDataForm.txtAction.value = "LOADEVENTLOGUSERS";
    //    data_refreshData();
        OpenHR.submitForm(frmGetDataForm);

    }
	
    function loadEventLogUsers(pbViewAll, psCurrentFilterUser)
    {
        var i;
        var bFoundUser = false;
	
        if (pbViewAll == 1)
        {
            var oOptionALL = document.createElement("OPTION");
            frmLog.cboUsername.options.add(oOptionALL);
            oOptionALL.innerText = '<All>';
            oOptionALL.value = -1;	
		
            var frmUtilDefForm = OpenHR.getForm("dataframe", "frmData");
            var dataCollection = frmUtilDefForm.elements;

            if (dataCollection!=null) 
            {
                for (i=0; i<dataCollection.length; i++)  
                {
                    sControlName = dataCollection.item(i).name;
                    sControlName = sControlName.substr(0, 16);
                    if (sControlName=="txtEventLogUser_") 
                    {
                        var oOption = document.createElement("OPTION");
                        frmLog.cboUsername.options.add(oOption);
                        oOption.innerText = dataCollection.item(i).value;
                        oOption.value = dataCollection.item(i).value;	
                        combo_disable(frmLog.cboUsername, false);
				
                        if (psCurrentFilterUser == dataCollection.item(i).value)
                        {
                            bFoundUser = true;
                            oOption.selected = true;
                        }
                    }	
                }
            }
			
            if (psCurrentFilterUser == '-1' || psCurrentFilterUser == '' || !bFoundUser)
            {
                oOptionALL.selected = true;
            }
						
            // Get menu to refresh the menu.
            menu_refreshMenu();


        }
        else
        {
            combo_disable(frmLog.cboUsername, true);
            var oOption = document.createElement("OPTION");
            frmLog.cboUsername.options.add(oOption);
            oOption.innerText = frmUseful.txtUserName.value;
            oOption.value = oOption.innerText;	
            oOption.selected = true;
        }

        refreshButtons();
		
        refreshGrid();
    }
	
    function okClick()
    {
        window.location.href="default";
    }

    function openDialog(pDestination, pWidth, pHeight) {

        dlgwinprops = "center:yes;" +
            "dialogHeight:" + pHeight + "px;" +
            "dialogWidth:" + pWidth + "px;" +
            "help:no;" +
            "resizable:yes;" +
            "scroll:yes;" +
            "status:no;";
        window.showModalDialog(pDestination, self, dlgwinprops);
    }

    -->
</script>

<OBJECT classid="clsid:F9043C85-F6F2-101A-A3C9-08002B2F49FB" 
	id=dialog 
  codebase="cabs/comdlg32.cab#Version=1,0,0,0"
	style="LEFT: 0px; TOP: 0px" 
	VIEWASTEXT>
	<PARAM NAME="_ExtentX" VALUE="847">
	<PARAM NAME="_ExtentY" VALUE="847">
	<PARAM NAME="_Version" VALUE="393216">
	<PARAM NAME="CancelError" VALUE="0">
	<PARAM NAME="Color" VALUE="0">
	<PARAM NAME="Copies" VALUE="1">
	<PARAM NAME="DefaultExt" VALUE="">
	<PARAM NAME="DialogTitle" VALUE="">
	<PARAM NAME="FileName" VALUE="">
	<PARAM NAME="Filter" VALUE="">
	<PARAM NAME="FilterIndex" VALUE="0">
	<PARAM NAME="Flags" VALUE="0">
	<PARAM NAME="FontBold" VALUE="0">
	<PARAM NAME="FontItalic" VALUE="0">
	<PARAM NAME="FontName" VALUE="">
	<PARAM NAME="FontSize" VALUE="8">
	<PARAM NAME="FontStrikeThru" VALUE="0">
	<PARAM NAME="FontUnderLine" VALUE="0">
	<PARAM NAME="FromPage" VALUE="0">
	<PARAM NAME="HelpCommand" VALUE="0">
	<PARAM NAME="HelpContext" VALUE="0">
	<PARAM NAME="HelpFile" VALUE="">
	<PARAM NAME="HelpKey" VALUE="">
	<PARAM NAME="InitDir" VALUE="">
	<PARAM NAME="Max" VALUE="0">
	<PARAM NAME="Min" VALUE="0">
	<PARAM NAME="MaxFileSize" VALUE="260">
	<PARAM NAME="PrinterDefault" VALUE="1">
	<PARAM NAME="ToPage" VALUE="0">
	<PARAM NAME="Orientation" VALUE="1"></OBJECT>


<form id=frmLog>
<table align=center class="outline" cellPadding=5 cellSpacing=0 width=100% height=100%>
	<TR>
		<TD>
			<TABLE WIDTH="100%" height="100%" class="invisible" cellspacing=0 cellpadding=0>
				<tr height=5> 
					<td colspan=3></td>
				</tr> 
								
				<tr> 
					<TD width=5></td>
					<td>
						<TABLE WIDTH="100%" height="100%" class="outline" cellspacing=0 cellpadding=5>
							<tr valign=top> 
								<td>
									<TABLE HEIGHT="100%" WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=4>
										<TR height=10>
											<TD colspan=8>
												Filters : 
											</TD>
										</TR>
										<TR height=10>
											<TD width=82 nowrap>
												User name : 
											</TD>
											<TD>
												<select id=cboUsername name=cboUsername class="combo" style="WIDTH: 100%" onchange="refreshGrid();">
												</select>		
											</TD>
											<TD width=25>
												Type : 
											</TD>
											<TD>
												<select id=cboType name=cboType class="combo" style="WIDTH: 100%" onchange="refreshGrid();">
												
<%
	if Session("CurrentType") = "-1" then
        Response.Write("											<option value=-1 selected>&lt;All&gt;" & vbCrLf)
	else
        Response.Write("											<option value=-1>&lt;All&gt;" & vbCrLf)
    End If
	
    If Session("CurrentType") = "17" Then
        Response.Write("											<option value=17 selected>Calendar Report" & vbCrLf)
    Else
        Response.Write("											<option value=17>Calendar Report" & vbCrLf)
    End If

    If Session("CurrentType") = "22" Then
        Response.Write("											<option value=22 selected>Career Progression" & vbCrLf)
    Else
        Response.Write("											<option value=22>Career Progression" & vbCrLf)
    End If

    If Session("CurrentType") = "1" Then
        Response.Write("											<option value=1 selected>Cross Tab" & vbCrLf)
    Else
        Response.Write("											<option value=1>Cross Tab" & vbCrLf)
    End If
	
    If Session("CurrentType") = "2" Then
        Response.Write("											<option value=2 selected>Custom Report" & vbCrLf)
    Else
        Response.Write("											<option value=2>Custom Report" & vbCrLf)
    End If
	
    If Session("CurrentType") = "3" Then
        Response.Write("											<option value=3 selected>Data Transfer" & vbCrLf)
    Else
        Response.Write("											<option value=3>Data Transfer" & vbCrLf)
    End If
	
    If Session("CurrentType") = "11" Then
        Response.Write("											<option value=11 selected>Diary Rebuild" & vbCrLf)
    Else
        Response.Write("											<option value=11>Diary Rebuild" & vbCrLf)
    End If
	
    If Session("CurrentType") = "12" Then
        Response.Write("											<option value=12 selected>Email Rebuild" & vbCrLf)
    Else
        Response.Write("											<option value=12>Email Rebuild" & vbCrLf)
    End If

    If Session("CurrentType") = "18" Then
        Response.Write("											<option value=18 selected>Envelopes & Labels" & vbCrLf)
    Else
        Response.Write("											<option value=18>Envelopes & Labels" & vbCrLf)
    End If
	
    If Session("CurrentType") = "4" Then
        Response.Write("											<option value=4 selected>Export" & vbCrLf)
    Else
        Response.Write("											<option value=4>Export" & vbCrLf)
    End If
	
    If Session("CurrentType") = "5" Then
        Response.Write("											<option value=5 selected>Global Add" & vbCrLf)
    Else
        Response.Write("											<option value=5>Global Add" & vbCrLf)
    End If
	
    If Session("CurrentType") = "6" Then
        Response.Write("											<option value=6 selected>Global Delete" & vbCrLf)
    Else
        Response.Write("											<option value=6>Global Delete" & vbCrLf)
    End If
	
    If Session("CurrentType") = "7" Then
        Response.Write("											<option value=7 selected>Global Update" & vbCrLf)
    Else
        Response.Write("											<option value=7>Global Update" & vbCrLf)
    End If
	
    If Session("CurrentType") = "8" Then
        Response.Write("											<option value=8 selected>Import" & vbCrLf)
    Else
        Response.Write("											<option value=8>Import" & vbCrLf)
    End If

    'if Session("CurrentType") = "19" then
    '	Response.Write "											<option value=19 selected>Label Definition" & vbCrLf
    'else
    '	Response.Write "											<option value=19>Label Definition" & vbCrLf
    'end if
	
    If Session("CurrentType") = "9" Then
        Response.Write("											<option value=9 selected>Mail Merge" & vbCrLf)
    Else
        Response.Write("											<option value=9>Mail Merge" & vbCrLf)
    End If

    If Session("CurrentType") = "16" Then
        Response.Write("											<option value=16 selected>Match Report" & vbCrLf)
    Else
        Response.Write("											<option value=16>Match Report" & vbCrLf)
    End If
	
	'if Session("CurrentType") = "14" then
	'	Response.Write "											<option value=14 selected>Record Editing" & vbCrLf
	'else
	'	Response.Write "											<option value=14>Record Editing" & vbCrLf
	'end if

    If Session("CurrentType") = "20" Then
        Response.Write("											<option value=20 selected>Record Profile" & vbCrLf)
    Else
        Response.Write("											<option value=20>Record Profile" & vbCrLf)
    End If

    If Session("CurrentType") = "13" Then
        Response.Write("											<option value=13 selected>Standard Report" & vbCrLf)
    Else
        Response.Write("											<option value=13>Standard Report" & vbCrLf)
    End If

    If Session("CurrentType") = "21" Then
        Response.Write("											<option value=21 selected>Succession Planning" & vbCrLf)
    Else
        Response.Write("											<option value=21>Succession Planning" & vbCrLf)
    End If

    If Session("CurrentType") = "15" Then
        Response.Write("											<option value=15 selected>System Error" & vbCrLf)
    Else
        Response.Write("											<option value=15>System Error" & vbCrLf)
    End If

    If Session("WF_Enabled") Then
        If Session("CurrentType") = "25" Then
            Response.Write("											<option value=25 selected>Workflow Rebuild" & vbCrLf)
        Else
            Response.Write("											<option value=25>Workflow Rebuild" & vbCrLf)
        End If
    End If
		
%>												
												</select>		
											</TD>
											<TD width=25>
												Mode : 
											</TD>
											<TD>
												<select id=cboMode name=cboMode class="combo" style="WIDTH: 100%" onchange="refreshGrid();">
<%
    If Session("CurrentMode") = "-1" Then
        Response.Write("											<option value=-1 selected>&lt;All&gt;" & vbCrLf)
    Else
        Response.Write("											<option value=-1>&lt;All&gt;" & vbCrLf)
    End If
	
    If Session("CurrentMode") = "1" Then
        Response.Write("											<option value=1 selected>Batch" & vbCrLf)
    Else
        Response.Write("											<option value=1>Batch" & vbCrLf)
    End If
	
    If Session("CurrentMode") = "0" Then
        Response.Write("											<option value=0 selected>Manual" & vbCrLf)
    Else
        Response.Write("											<option value=0>Manual" & vbCrLf)
    End If
%>
                                                </select>
                                            </td>
                                            <td width="25">Status : 
                                            </td>
                                            <td>
                                                <select id="cboStatus" name="cboStatus" class="combo" style="width: 100%" onchange="refreshGrid();">
                                                    <%	
                                                        If Session("CurrentStatus") = "-1" Then
                                                            Response.Write("											<option value=-1 selected>&lt;All&gt;" & vbCrLf)
                                                        Else
                                                            Response.Write("											<option value=-1>&lt;All&gt;" & vbCrLf)
                                                        End If
		
                                                        If Session("CurrentStatus") = "1" Then
                                                            Response.Write("											<option value=1 selected>Cancelled" & vbCrLf)
                                                        Else
                                                            Response.Write("											<option value=1>Cancelled" & vbCrLf)
                                                        End If
	
                                                        If Session("CurrentStatus") = "5" Then
                                                            Response.Write("											<option value=5 selected>Error" & vbCrLf)
                                                        Else
                                                            Response.Write("											<option value=5>Error" & vbCrLf)
                                                        End If
	
                                                        If Session("CurrentStatus") = "2" Then
                                                            Response.Write("											<option value=2 selected>Failed" & vbCrLf)
                                                        Else
                                                            Response.Write("											<option value=2>Failed" & vbCrLf)
                                                        End If
	
                                                        If Session("CurrentStatus") = "0" Then
                                                            Response.Write("											<option value=0 selected>Pending" & vbCrLf)
                                                        Else
                                                            Response.Write("											<option value=0>Pending" & vbCrLf)
                                                        End If
	
                                                        If Session("CurrentStatus") = "4" Then
                                                            Response.Write("											<option value=4 selected>Skipped" & vbCrLf)
                                                        Else
                                                            Response.Write("											<option value=4>Skipped" & vbCrLf)
                                                        End If
	
                                                        If Session("CurrentStatus") = "3" Then
                                                            Response.Write("											<option value=3 selected>Successful" & vbCrLf)
                                                        Else
                                                            Response.Write("											<option value=3>Successful" & vbCrLf)
                                                        End If
                                                    %>
												</select>		
											</TD>
										</TR>
										<TR height=5>
											<TD colspan=8></TD>
										</TR>
										<TR>
											<TD colspan=8>
<%

	dim avColumnDef(13,4)
	dim cmdEventLogRecords
	dim rsEventLogRecords
	dim lngRowCount
	dim iLoop
	dim sAddString
	
	avColumnDef(0,0) = "ID"				'name
	avColumnDef(0,1) = "ID"				'caption
	avColumnDef(0,2) = "1600"			'width
	avColumnDef(0,3) = "0"				'visible
	
	avColumnDef(1,0) = "DateTime"	'name
	avColumnDef(1,1) = "Start Time"		'caption
	avColumnDef(1,2) = "3300"					'width
	avColumnDef(1,3) = "-1"						'visible

	avColumnDef(2,0) = "EndTime"	'name
	avColumnDef(2,1) = "End Time"	'caption
	avColumnDef(2,2) = "3300"			'width
	avColumnDef(2,3) = "-1"				'visible
	
	avColumnDef(3,0) = "Duration"	'name
	avColumnDef(3,1) = "Duration"	'caption
	avColumnDef(3,2) = "1750"			'width
	avColumnDef(3,3) = "-1"				'visible

	avColumnDef(4,0) = "Type"		'name
	avColumnDef(4,1) = "Type"		'caption
	avColumnDef(4,2) = "3250"		'width
	avColumnDef(4,3) = "-1"			'visible

	avColumnDef(5,0) = "Name"		'name
	avColumnDef(5,1) = "Name"		'caption
	avColumnDef(5,2) = "5500"		'width
	avColumnDef(5,3) = "-1"			'visible

	avColumnDef(6,0) = "Status"		'name
	avColumnDef(6,1) = "Status"		'caption
	avColumnDef(6,2) = "2100"			'width
	avColumnDef(6,3) = "-1"				'visible

	avColumnDef(7,0) = "Mode"			'name
	avColumnDef(7,1) = "Mode"			'caption
	avColumnDef(7,2) = "1500"			'width
	avColumnDef(7,3) = "-1"				'visible

	avColumnDef(8,0) = "Username"		'name
	avColumnDef(8,1) = "User name"	'caption
	avColumnDef(8,2) = "2500"				'width
	avColumnDef(8,3) = "-1"					'visible

	avColumnDef(9,0) = "BatchJobID"	'name
	avColumnDef(9,1) = "BatchJobID"	'caption
	avColumnDef(9,2) = "1800"				'width
	avColumnDef(9,3) = "0"					'visible
	
	avColumnDef(10,0) = "BatchRunID"	'name
	avColumnDef(10,1) = "BatchRunID"	'caption
	avColumnDef(10,2) = "1800"				'width
	avColumnDef(10,3) = "0"						'visible

	avColumnDef(11,0) = "BatchName"		'name
	avColumnDef(11,1) = "Batch Name"	'caption
	avColumnDef(11,2) = "1800"				'width
	avColumnDef(11,3) = "0"						'visible
	
	avColumnDef(12,0) = "SuccessCount"	'name
	avColumnDef(12,1) = "SuccessCount"	'caption
	avColumnDef(12,2) = "1800"					'width
	avColumnDef(12,3) = "0"							'visible
	
	avColumnDef(13,0) = "FailCount"		'name
	avColumnDef(13,1) = "FailCount"		'caption
	avColumnDef(13,2) = "1800"				'width
	avColumnDef(13,3) = "0"						'visible
		
    Response.Write("											<OBJECT classid=clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" & vbCrLf)
    Response.Write("													 codebase=""cabs/COAInt_Grid.cab#version=3,1,3,6""" & vbCrLf)
    Response.Write("													height=""100%""" & vbCrLf)
    Response.Write("													id=ssOleDBGridEventLog" & vbCrLf)
    Response.Write("													name=ssOleDBGridEventLog" & vbCrLf)
    Response.Write("													style=""HEIGHT: 100%; VISIBILITY: visible; WIDTH: 100%""" & vbCrLf)
    Response.Write("													width=""100%"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""ScrollBars"" VALUE=""3"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""_Version"" VALUE=""196617"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""DataMode"" VALUE=""2"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""Cols"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""Rows"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""BorderStyle"" VALUE=""1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""RecordSelectors"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""GroupHeaders"" VALUE=""-1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""ColumnHeaders"" VALUE=""-1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""GroupHeadLines"" VALUE=""1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""HeadLines"" VALUE=""1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""FieldDelimiter"" VALUE=""(None)"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""FieldSeparator"" VALUE=""(Tab)"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""Row.Count"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""Col.Count"" VALUE=""1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""stylesets.count"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""TagVariant"" VALUE=""EMPTY"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""UseGroups"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""HeadFont3D"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""Font3D"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""DividerType"" VALUE=""3"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""DividerStyle"" VALUE=""1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""DefColWidth"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""BeveColorScheme"" VALUE=""2"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""BevelColorFrame"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""BevelColorHighlight"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""BevelColorShadow"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""BevelColorFace"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""CheckBox3D"" VALUE=""-1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowAddNew"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowDelete"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowUpdate"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""MultiLine"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""ActiveCellStyleSet"" VALUE="""">" & vbCrLf)
    Response.Write("												<PARAM NAME=""RowSelectionStyle"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowRowSizing"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowGroupSizing"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowColumnSizing"" VALUE=""-1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowGroupMoving"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowColumnMoving"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowGroupSwapping"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowColumnSwapping"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowGroupShrinking"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowColumnShrinking"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""AllowDragDrop"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""UseExactRowCount"" VALUE=""-1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""SelectTypeCol"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""SelectTypeRow"" VALUE=""3"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""SelectByCell"" VALUE=""-1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""BalloonHelp"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""RowNavigation"" VALUE=""2"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""CellNavigation"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""MaxSelectedRows"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""HeadStyleSet"" VALUE="""">" & vbCrLf)
    Response.Write("												<PARAM NAME=""StyleSet"" VALUE="""">" & vbCrLf)
    Response.Write("												<PARAM NAME=""ForeColorEven"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""ForeColorOdd"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""BackColorEven"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""BackColorOdd"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""Levels"" VALUE=""1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""RowHeight"" VALUE=""503"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""ExtraHeight"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""ActiveRowStyleSet"" VALUE="""">" & vbCrLf)
    Response.Write("												<PARAM NAME=""CaptionAlignment"" VALUE=""2"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""SplitterPos"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""SplitterVisible"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""Columns.Count"" VALUE=""" & (UBound(avColumnDef) + 1) & """>" & vbCrLf)
	
    For i = 0 To UBound(avColumnDef) Step 1
        Response.Write("												<!--" & avColumnDef(i, 0) & "-->  " & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Width"" VALUE=""" & avColumnDef(i, 2) & """>" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Visible"" VALUE=""" & avColumnDef(i, 3) & """>" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Columns.Count"" VALUE=""1"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Caption"" VALUE=""" & avColumnDef(i, 1) & """>" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Name"" VALUE=""" & avColumnDef(i, 0) & """>" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Alignment"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").CaptionAlignment"" VALUE=""3"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Bound"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").AllowSizing"" VALUE=""1"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").DataField"" VALUE=""Column " & i & """>" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").DataType"" VALUE=""8"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Level"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").NumberFormat"" VALUE="""">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Case"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").FieldLen"" VALUE=""256"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").VertScrollBar"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Locked"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Style"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").ButtonsAlways"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").RowCount"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").ColCount"" VALUE=""1"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").HasHeadForeColor"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").HasHeadBackColor"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").HasForeColor"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").HasBackColor"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").HeadForeColor"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").HeadBackColor"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").ForeColor"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").BackColor"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").HeadStyleSet"" VALUE="""">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").StyleSet"" VALUE="""">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Nullable"" VALUE=""1"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").Mask"" VALUE="""">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").PromptInclude"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").ClipMode"" VALUE=""0"">" & vbCrLf)
        Response.Write("												<PARAM NAME=""Columns(" & i & ").PromptChar"" VALUE=""95"">" & vbCrLf)
	next
		
    Response.Write("												<PARAM NAME=""UseDefaults"" VALUE=""-1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""TabNavigation"" VALUE=""1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""BatchUpdate"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""_ExtentX"" VALUE=""11298"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""_ExtentY"" VALUE=""3969"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""_StockProps"" VALUE=""79"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""Caption"" VALUE="""">" & vbCrLf)
    Response.Write("												<PARAM NAME=""ForeColor"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""BackColor"" VALUE=""0"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""Enabled"" VALUE=""-1"">" & vbCrLf)
    Response.Write("												<PARAM NAME=""DataMember"" VALUE="""">" & vbCrLf)

    Response.Write("												<PARAM NAME=""Row.Count"" VALUE=""0"">" & vbCrLf)
    Response.Write("											</OBJECT>" & vbCrLf)
%>											
											</TD>
										</TR>									
									</TABLE>
								</td>
								<td width=80>
									<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=2>
										<TR>
											<TD width=10>
												<INPUT id=cmdOK type=button class="btn" value=OK name=cmdOK style="WIDTH: 80px" width="80" 
											        onclick="okClick();"
		                                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                            onfocus="try{button_onFocus(this);}catch(e){}"
		                                            onblur="try{button_onBlur(this);}catch(e){}" />
											</TD>
										</TR>
										<tr height=10>
										<td></td>
										</tr>
										<TR>
											<TD width=10>
												<INPUT id=cmdView class="btn" type=button value="View..." name=cmdView style="WIDTH: 80px" width="80"
                                                    onclick="viewEvent();"
		                                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                            onfocus="try{button_onFocus(this);}catch(e){}"
		                                            onblur="try{button_onBlur(this);}catch(e){}" />
											</TD>
										</tr>
										<tr height=10>
										<td></td>
										</tr>
										<TR>
											<TD width=10>
												<INPUT id=cmdDelete class="btn" type=button value="Delete..." name=cmdDelete style="WIDTH: 80px" width="80"
												    onclick="deleteEvent();" 
		                                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                            onfocus="try{button_onFocus(this);}catch(e){}"
		                                            onblur="try{button_onBlur(this);}catch(e){}" />
											</TD>
										</tr>
										<tr height=10>
										<td></td>
										</tr>
										<TR>
											<TD width=10>
												<INPUT id=cmdPurge class="btn" type=button value="Purge..." name=cmdPurge style="WIDTH: 80px" width="80"
												    onclick="purgeEvent();" 
		                                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                            onfocus="try{button_onFocus(this);}catch(e){}"
		                                            onblur="try{button_onBlur(this);}catch(e){}" />
											</TD>
										</tr>
										<tr height=10>
										<td></td>
										</tr>
										<TR>
											<TD width=10>
												<INPUT id=cmdEmail class="btn" type=button value="Email..." name=cmdEmail style="WIDTH: 80px" width="80"
												    onclick="emailEvent();" 
		                                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                            onfocus="try{button_onFocus(this);}catch(e){}"
		                                            onblur="try{button_onBlur(this);}catch(e){}" />
											</TD>
										</tr>
									</table>								
								</td>
							</tr>
							
						</TABLE>
					</td>
					<TD width=5></td>
				</tr> 
				<tr height=8> 
					<TD width=5></td>
					<td colspan=1>
						<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=1>
							<TR >
								<TD name=sbEventLog id=sbEventLog>
								&nbsp
								</TD>
							</TD>
						</TABLE>
					</td>
					<TD width=5></td>
				</tr> 
			</TABLE>
		</td>
	</tr> 
</TABLE>

<INPUT type='hidden' id=txtELDeletePermission name=txtELDeletePermission>
<INPUT type='hidden' id=txtELViewAllPermission name=txtELViewAllPermission>
<INPUT type='hidden' id=txtELPurgePermission name=txtELPurgePermission>
<INPUT type='hidden' id=txtELEmailPermission name=txtELEmailPermission>
<INPUT type='hidden' id=txtELOrderColumn name=txtELOrderColumn value='DateTime'>
<INPUT type='hidden' id=txtELOrderOrder name=txtELOrderOrder value='DESC'>
<INPUT type='hidden' id=txtELSortColumnIndex name=txtELSortColumnIndex value=1>
<INPUT type='hidden' id=txtELLoaded name=txtELLoaded value=0>
<INPUT type="hidden" id=txtCurrUserFilter name=txtCurrUserFilter value='<%=Session("CurrentUsername")%>'>

</form>

<FORM action="default_Submit" method=post id=frmGoto name=frmGoto>
    <%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</FORM>

<FORM id=frmDetails name=frmDetails method=post style="visibility:hidden;display:none">
	<INPUT type="hidden" id=txtEventID name=txtEventID>

	<INPUT type="hidden" id=txtEventName name=txtEventName>
	<INPUT type="hidden" id=txtEventMode name=txtEventMode>

	<INPUT type="hidden" id=txtEventStartTime name=txtEventStartTime>
	<INPUT type="hidden" id=txtEventEndTime name=txtEventEndTime>
	<INPUT type="hidden" id=txtEventDuration name=txtEventDuration>

	<INPUT type="hidden" id=txtEventType name=txtEventType>
	<INPUT type="hidden" id=txtEventStatus name=txtEventStatus>
	<INPUT type="hidden" id=txtEventUser name=txtEventUser>
	
	<INPUT type="hidden" id=txtEventSuccessCount name=txtEventSuccessCount>
	<INPUT type="hidden" id=txtEventFailCount name=txtEventFailCount>

	<INPUT type="hidden" id=txtEventBatchName name=txtEventBatchName>
	<INPUT type="hidden" id=txtEventBatchJobID name=txtEventBatchJobID>
	<INPUT type="hidden" id=txtEventBatchRunID name=txtEventBatchRunID>
	
	<INPUT type="hidden" id=txtEmailPermission name=txtEmailPermission>
</FORM>

<FORM id=frmPurge name=frmPurge method=post style="visibility:hidden;display:none" action="eventLog">
	<INPUT type="hidden" id=txtDoesPurge name=txtDoesPurge>
	<INPUT type="hidden" id=txtPurgePeriod name=txtPurgePeriod>
	<INPUT type="hidden" id=txtPurgeFrequency name=txtPurgeFrequency>
	<INPUT type="hidden" id=txtShowPurgeMSG name=txtShowPurgeMSG value=<%=Session("showPurgeMessage")%>>
	<INPUT type="hidden" id=txtCurrentUsername name=txtCurrentUsername>
	<INPUT type="hidden" id=txtCurrentType name=txtCurrentType>
	<INPUT type="hidden" id=txtCurrentMode name=txtCurrentMode>
	<INPUT type="hidden" id=txtCurrentStatus name=txtCurrentStatus>
</FORM>

<FORM id=frmDelete name=frmDelete method=post style="visibility:hidden;display:none" action="eventLog">
	<INPUT type="hidden" id=txtDeleteSel name=txtDeleteSel>
	<INPUT type="hidden" id=txtSelectedIDs name=txtSelectedIDs>
	<INPUT type="hidden" id=txtViewAllPerm name=txtViewAllPerm>
	<INPUT type="hidden" id=Hidden1 name=txtCurrentUsername>
	<INPUT type="hidden" id=Hidden2 name=txtCurrentType>
	<INPUT type="hidden" id=Hidden3 name=txtCurrentMode>
	<INPUT type="hidden" id=Hidden4 name=txtCurrentStatus>
</FORM>

<FORM id=frmEmail name=frmEmail method=post style="visibility:hidden;display:none" action="emailSelection">
	<INPUT type="hidden" id=txtSelectedEventIDs name=txtSelectedEventIDs>
	<INPUT type="hidden" id=txtFromMain name=txtFromMain value=1>
	<INPUT type="hidden" id=txtEmailOrderColumn name=txtEmailOrderColumn>
	<INPUT type="hidden" id=txtEmailOrderOrder name=txtEmailOrderOrder>
</FORM>

<FORM id=frmRefresh name=frmRefresh method=post style="visibility:hidden;display:none" action="eventLog">
	<INPUT type="hidden" id=txtEventExisted name=txtEventExisted>
	<INPUT type="hidden" id=Hidden5 name=txtCurrentUsername>
	<INPUT type="hidden" id=Hidden6 name=txtCurrentType>
	<INPUT type="hidden" id=Hidden7 name=txtCurrentMode>
	<INPUT type="hidden" id=Hidden8 name=txtCurrentStatus>
</FORM>

<%
	Session("showPurgeMessage") = 0
%>

<FORM id=frmUseful name=frmUseful style="visibility:hidden;display:none">
	<INPUT type="hidden" id=txtUserName name=txtUserName value="<%=session("username")%>">
<%
    Dim cmdDefinition
    Dim prmModuleKey
    Dim prmParameterKey
    Dim prmParameterValue
    Dim sErrorDescription As String
    
	cmdDefinition = Server.CreateObject("ADODB.Command")
	cmdDefinition.CommandText = "sp_ASRIntGetModuleParameter"
	cmdDefinition.CommandType = 4 ' Stored procedure.
    cmdDefinition.ActiveConnection = Session("databaseConnection")

    prmModuleKey = cmdDefinition.CreateParameter("moduleKey", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDefinition.Parameters.Append(prmModuleKey)
    prmModuleKey.value = "MODULE_PERSONNEL"

    prmParameterKey = cmdDefinition.CreateParameter("paramKey", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdDefinition.Parameters.Append(prmParameterKey)
    prmParameterKey.value = "Param_TablePersonnel"

    prmParameterValue = cmdDefinition.CreateParameter("paramValue", 200, 2, 8000) '200=varchar, 2=output, 8000=size
    cmdDefinition.Parameters.Append(prmParameterValue)

    err.clear()
    cmdDefinition.Execute()

    Response.Write("<INPUT type='hidden' id=txtPersonnelTableID name=txtPersonnelTableID value=" & cmdDefinition.Parameters("paramValue").Value & ">" & vbCrLf)
	
	cmdDefinition = nothing

    Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
    Response.Write("<INPUT type='hidden' id=txtAction name=txtAction value=" & Session("action") & ">" & vbCrLf)
%>
</FORM>

<INPUT type='hidden' id=txtTicker name=txtTicker value=0>
<INPUT type='hidden' id=txtLastKeyFind name=txtLastKeyFind value="">

<%
	Session("CurrentUsername") = ""
	Session("CurrentType") = ""
	Session("CurrentMode") = ""
	Session("CurrentStatus") = ""
%>


<script type="text/javascript">

    function addActiveXHandlers() {
        debugger;

        OpenHR.addActiveXHandler("ssOleDBGridEventLog", "DblClick", ssOleDBGridEventLog_dblclick);
        OpenHR.addActiveXHandler("ssOleDBGridEventLog", "rowcolchange", ssOleDBGridEventLog_rowcolchange);
        OpenHR.addActiveXHandler("ssOleDBGridEventLog", "Click", ssOleDBGridEventLog_click);
        OpenHR.addActiveXHandler("ssOleDBGridEventLog", "HeadClick", ssOleDBGridEventLog_headclick);
    }

    function ssOleDBGridEventLog_dblclick()
    {
        if ((frmLog.ssOleDBGridEventLog.Rows > 0) && (frmLog.ssOleDBGridEventLog.SelBookmarks.Count == 1)) 
        {
            viewEvent();
        }
    }

    function ssOleDBGridEventLog_rowcolchange()
    {
        if (frmLog.ssOleDBGridEventLog.SelBookmarks.Count > 1) 
        {
            button_disable(frmLog.cmdView, true);
        }
        else
        {
            button_disable(frmLog.cmdView, false);
        }
    }

    function ssOleDBGridEventLog_click()
    {
        if ((frmLog.ssOleDBGridEventLog.SelBookmarks.Count > 1) || (frmLog.ssOleDBGridEventLog.Rows == 0)) 
        {
            button_disable(frmLog.cmdView, true);
        }
        else
        {
            button_disable(frmLog.cmdView, false);
        }
    }

    function ssOleDBGridEventLog_headclick()
    {

        var ColIndex = arguments[0];
 
        //Set the sort criteria depending on the column header clicked and refresh the grid
        if (ColIndex == 1)
        { 
            frmLog.txtELOrderColumn.value = 'DateTime';
        }
        else if (ColIndex == 2)
        { 
            frmLog.txtELOrderColumn.value = 'EndTime';
        }
        else if (ColIndex == 3)
        {  
            frmLog.txtELOrderColumn.value = 'Duration';
        }
        else if (ColIndex == 4)
        { 
            frmLog.txtELOrderColumn.value = 'Type';
        }
        else if (ColIndex == 5)
        { 	
            frmLog.txtELOrderColumn.value = 'Name';
        }
        else if (ColIndex == 6)
        { 	
            frmLog.txtELOrderColumn.value = 'Status';
        }
        else if (ColIndex == 7)
        { 	
            frmLog.txtELOrderColumn.value = 'Mode';
        }
        else if (ColIndex == 8)
        { 	
            frmLog.txtELOrderColumn.value = 'Username';
        }
        else
        { 
            frmLog.txtELOrderColumn.value = 'DateTime';
        }
		
        if (ColIndex == frmLog.txtELSortColumnIndex.value)
        {
            if (frmLog.txtELOrderOrder.value == 'ASC') 
            {
                frmLog.txtELOrderOrder.value = 'DESC'; 
            }
            else
            {
                frmLog.txtELOrderOrder.value = 'ASC';
            }
        }
        else
        {
            frmLog.txtELOrderOrder.value = 'ASC';
        }
  
        frmLog.txtELSortColumnIndex.value = ColIndex;
 
        refreshGrid();
    }






</script>


<script type="text/javascript">
  //  addActiveXHandlers();
    EventLog_window_onload();   
</script>
