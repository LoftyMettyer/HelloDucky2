<%@ Page Title="" Language="VB" MasterPageFile="~/Views/Shared/Site.Master" Inherits="System.Web.Mvc.ViewPage" %>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
EventLog
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

<%@ Language=VBScript %>
<!--#INCLUDE FILE="include/svrCleanup.asp" -->
<%
	Response.Expires = -1

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
	
	Response.Buffer = True
	
	Session("CurrentUsername") = CStr(Request.Form("txtCurrentUsername"))
	Session("CurrentType") = CStr(Request.Form("txtCurrentType"))
	Session("CurrentMode") = CStr(Request.Form("txtCurrentMode"))
	Session("CurrentStatus") = CStr(Request.Form("txtCurrentStatus"))	
	
	if Session("CurrentUsername") = "" then
		Session("CurrentUsername") = "-1"
	end if
	
	if Session("CurrentType") = "" then
		Session("CurrentType") = "-1"
	end if

	if Session("CurrentMode") = "" then
		Session("CurrentMode") = "-1"
	end if

	if Session("CurrentStatus") = "" then
		Session("CurrentStatus") = "-1"
	end if
%>

<%
	'This section of script is used for saving the new purge criteria.
	dim bDoesPurge
	dim sPeriod
	dim iFrequency
	dim sSQL
	dim cmdPurge
	
	bDoesPurge = trim(Request.Form("txtDoesPurge"))
	sPeriod = Request.Form("txtPurgePeriod")
	iFrequency = Request.Form("txtPurgeFrequency")
	
	if bDoesPurge <> vbNullString then
		
		' Delete old purge information to the database
		Set cmdPurge = Server.CreateObject("ADODB.Command")
		cmdPurge.CommandText = "spASRIntClearEventLogPurge"
		cmdPurge.CommandType = 4 ' Stored procedure.
		Set cmdPurge.ActiveConnection = session("databaseConnection")
		err = 0
		cmdPurge.Execute 
		set cmdPurge = nothing
 
		if bDoesPurge = 1 then
			' Insert the new purge criteria
			Set cmdPurge = Server.CreateObject("ADODB.Command")
			cmdPurge.CommandText = "spASRIntSetEventLogPurge"
			cmdPurge.CommandType = 4 ' Stored procedure.
			Set cmdPurge.ActiveConnection = session("databaseConnection")
		
			Set prmPeriod = cmdPurge.CreateParameter("period",200,1,8000) ' 200=varchar, 1=input, 8000=size
			cmdPurge.Parameters.Append prmPeriod
			prmPeriod.value = cstr(sPeriod)

			Set prmFrequency = cmdPurge.CreateParameter("frequency",3,1) ' 3=integer, 1=input
			cmdPurge.Parameters.Append prmFrequency
			prmFrequency.value = cleanNumeric(clng(iFrequency))
			
			err = 0
			cmdPurge.Execute 
			set cmdPurge = nothing

			Session("showPurgeMessage") = 1
		else
			Session("showPurgeMessage") = 0 
		end if
	end if
	
	set bDoesPurge = nothing
	set sPeriod = nothing
	set iFrequency = nothing
	set sSQL = nothing
%>

<%
	'This section of script is used for deleting Event Log records according to the selection on the Delete screen. 
	dim iDeleteSelection
	dim sSelectedEventIDs 
	dim cmdDelete
	dim bHasViewAllPermission
	
	iDeleteSelection = Request.Form("txtDeleteSel") 
	sSelectedEventIDs = Request.Form("txtSelectedIDs")
	bHasViewAllPermission = Request.Form("txtViewAllPerm")

	if iDeleteSelection <> vbNullString then
		iDeleteSelection = CInt(iDeleteSelection)
		
		set	cmdDelete = Server.CreateObject("ADODB.Command")
		cmdDelete.CommandText = "spASRIntDeleteEventLogRecords"
		cmdDelete.CommandType = 4 ' Stored procedure.
		Set cmdDelete.ActiveConnection = session("databaseConnection")
		
		Set prmType = cmdDelete.CreateParameter("type",3,1) ' 3=integer, 1=input
		cmdDelete.Parameters.Append prmType
		prmType.value = cleanNumeric(clng(iDeleteSelection))

		Set prmEventIDs = cmdDelete.CreateParameter("eventIDs",200,1,8000) ' 200=varchar, 1=input, 8000=size
		cmdDelete.Parameters.Append prmEventIDs
		prmEventIDs.value = cstr(sSelectedEventIDs)

		Set prmCanViewAll = cmdDelete.CreateParameter("canViewAll",11,1) ' 11=bit, 1=input
		cmdDelete.Parameters.Append prmCanViewAll
		prmCanViewAll.value = cleanBoolean(cbool(bHasViewAllPermission))

		err = 0
		cmdDelete.Execute 
		set cmdDelete = nothing
	end if

	set iDeleteSelection = nothing
	set sSelectedEventIDs = nothing
	set cmdDelete = nothing
	set bHasViewAllPermission = nothing
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK href="OpenHR.css" rel=stylesheet type=text/css >

<!--#include file="include\ctl_SetFont.txt"-->

<OBJECT 
	classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" 
	id="Microsoft_Licensed_Class_Manager_1_0" 
	VIEWASTEXT>
	<PARAM NAME="LPKPath" VALUE="lpks/main.lpk">
</OBJECT>

<SCRIPT FOR=window EVENT=onload LANGUAGE=JavaScript>
<!--
    window.parent.document.all.item("workframeset").cols = "*, 0";
	
    window.parent.frames("menuframe").abMainMenu.Bands("mnubandMainToolBar").tools("mnutoolRecordPosition").caption = '';
	
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
        window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sErrMsg);
        window.parent.location.replace("login.asp");
    }
	
    if (fOK == true) {
        // Get menu.asp to refresh the menu.
        window.parent.frames("menuframe").refreshMenu();		  
    }
	
    frmLog.txtELDeletePermission.value = window.parent.frames("menuframe").document.all.item("txtSysPerm_EVENTLOG_DELETE").value;
    frmLog.txtELViewAllPermission.value = window.parent.frames("menuframe").document.all.item("txtSysPerm_EVENTLOG_VIEWALL").value;
    frmLog.txtELPurgePermission.value = window.parent.frames("menuframe").document.all.item("txtSysPerm_EVENTLOG_PURGE").value;
    frmLog.txtELEmailPermission.value = window.parent.frames("menuframe").document.all.item("txtSysPerm_EVENTLOG_EMAIL").value;

    refreshUsers();

    // Little dodge to get around a browser bug that
    // does not refresh the display on all controls.
    try
    {
        window.resizeBy(0,-1);
        window.resizeBy(0,1);
    }
    catch(e) {}

    -->
</SCRIPT>

<SCRIPT LANGUAGE=JavaScript id=scptGeneralFunctions>
<!--
	
    function moveRecord(psMovement)
    {
        var frmGetData = window.parent.frames("dataframe").document.forms("frmGetData");
        var frmData = window.parent.frames("dataframe").document.forms("frmData");

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
        var frmData = window.parent.frames("dataframe").document.forms("frmData");
	
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
	
        window.parent.frames("menuframe").abMainMenu.Tools("mnutoolRecordPosition").visible = true;
        window.parent.frames("menuframe").abMainMenu.Bands("mnubandMainToolBar").tools("mnutoolRecordPosition").caption = sCaption;
        window.parent.frames("menuframe").abMainMenu.RecalcLayout();
	
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

        var frmUtilDefForm = window.parent.frames("dataframe").document.forms("frmData");
        var dataCollection = frmUtilDefForm.elements;
	
        var frmRefresh;
        frmRefresh = window.parent.frames("pollframe").document.forms("frmHit");
	
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
                    frmRefresh.submit();
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
		
            frmRefresh.submit();

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

        //Set the event log loaded flag, used in the menu.asp
        frmLog.txtELLoaded.value = 1;
	
        // Get menu.asp to refresh the menu.
        window.parent.frames("menuframe").refreshMenu();
	
        refreshStatusBar()
	
        if (frmPurge.txtShowPurgeMSG.value == 1)
        {
            window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("Purge completed.",64,"Event Log");
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
        var frmGetDataForm = window.parent.frames("dataframe").document.forms("frmGetData");
	
        frmGetDataForm.txtAction.value = "LOADEVENTLOG";
	
        frmGetDataForm.txtELFilterUser.value = frmLog.cboUsername.options[frmLog.cboUsername.selectedIndex].value; 
        frmGetDataForm.txtELFilterType.value = frmLog.cboType.options[frmLog.cboType.selectedIndex].value; 
        frmGetDataForm.txtELFilterStatus.value = frmLog.cboStatus.options[frmLog.cboStatus.selectedIndex].value; 
        frmGetDataForm.txtELFilterMode.value = frmLog.cboMode.options[frmLog.cboMode.selectedIndex].value; 
        frmGetDataForm.txtELOrderColumn.value = frmLog.txtELOrderColumn.value; 
        frmGetDataForm.txtELOrderOrder.value = frmLog.txtELOrderOrder.value; 
	
        refreshButtons();
	
        window.parent.frames("dataframe").refreshData();
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
	
            sURL = "eventLogDetails.asp" +
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
		
        sURL = "eventLogSelection.asp" +
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
		
        sURL = "eventLogPurge.asp" +
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
		
        sURL = "emailSelection.asp" +
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
        var frmGetDataForm = window.parent.frames("dataframe").document.forms("frmGetData");
	
        frmGetDataForm.txtAction.value = "LOADEVENTLOGUSERS";

        window.parent.frames("dataframe").refreshData();
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
		
            var frmUtilDefForm = window.parent.frames("dataframe").document.forms("frmData");
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
						
            // Get menu.asp to refresh the menu.
            window.parent.frames("menuframe").refreshMenu();		
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
        window.location.href="default.asp";
    }

    function openDialog(pDestination, pWidth, pHeight)
    {
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
</SCRIPT>

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

<SCRIPT FOR=ssOleDBGridEventLog EVENT=Click LANGUAGE=JavaScript>
<!--
  
    if ((frmLog.ssOleDBGridEventLog.SelBookmarks.Count > 1) || (frmLog.ssOleDBGridEventLog.Rows == 0)) 
    {
        button_disable(frmLog.cmdView, true);
    }
    else
    {
        button_disable(frmLog.cmdView, false);
    }

    -->
</SCRIPT>

<SCRIPT FOR=ssOleDBGridEventLog EVENT=HeadClick LANGUAGE=JavaScript>
<!--
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

    -->
</SCRIPT>

<SCRIPT FOR=ssOleDBGridEventLog EVENT=RowColChange LANGUAGE=JavaScript>
<!--
    if (frmLog.ssOleDBGridEventLog.SelBookmarks.Count > 1) 
    {
        button_disable(frmLog.cmdView, true);
    }
    else
    {
        button_disable(frmLog.cmdView, false);
    }
    -->
</SCRIPT>

<SCRIPT FOR=ssOleDBGridEventLog EVENT=DblClick LANGUAGE=JavaScript>
<!--
  
    if ((frmLog.ssOleDBGridEventLog.Rows > 0) && (frmLog.ssOleDBGridEventLog.SelBookmarks.Count == 1)) 
    {
        viewEvent();
    }

    -->
</SCRIPT>
<!--#INCLUDE FILE="include/ctl_SetStyles.txt" -->
</HEAD>

<BODY <%=session("BodyTag")%>>
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
		Response.Write "											<option value=-1 selected>&lt;All&gt;" & vbCrLf 
	else
		Response.Write "											<option value=-1>&lt;All&gt;" & vbCrLf 
	end if
	
	if Session("CurrentType") = "17" then
		Response.Write "											<option value=17 selected>Calendar Report" & vbCrLf
	else
		Response.Write "											<option value=17>Calendar Report" & vbCrLf
	end if

	if Session("CurrentType") = "22" then
		Response.Write "											<option value=22 selected>Career Progression" & vbCrLf
	else
		Response.Write "											<option value=22>Career Progression" & vbCrLf
	end if

	if Session("CurrentType") = "1" then
		Response.Write "											<option value=1 selected>Cross Tab" & vbCrLf
	else
		Response.Write "											<option value=1>Cross Tab" & vbCrLf
	end if
	
	if Session("CurrentType") = "2" then
		Response.Write "											<option value=2 selected>Custom Report" & vbCrLf
	else
		Response.Write "											<option value=2>Custom Report" & vbCrLf
	end if
	
	if Session("CurrentType") = "3" then
		Response.Write "											<option value=3 selected>Data Transfer" & vbCrLf
	else
		Response.Write "											<option value=3>Data Transfer" & vbCrLf
	end if
	
	if Session("CurrentType") = "11" then
		Response.Write "											<option value=11 selected>Diary Rebuild" & vbCrLf
	else
		Response.Write "											<option value=11>Diary Rebuild" & vbCrLf
	end if
	
	if Session("CurrentType") = "12" then
		Response.Write "											<option value=12 selected>Email Rebuild" & vbCrLf
	else
		Response.Write "											<option value=12>Email Rebuild" & vbCrLf
	end if

	if Session("CurrentType") = "18" then
		Response.Write "											<option value=18 selected>Envelopes & Labels" & vbCrLf
	else
		Response.Write "											<option value=18>Envelopes & Labels" & vbCrLf
	end if
	
	if Session("CurrentType") = "4" then
		Response.Write "											<option value=4 selected>Export" & vbCrLf
	else
		Response.Write "											<option value=4>Export" & vbCrLf
	end if
	
	if Session("CurrentType") = "5" then
		Response.Write "											<option value=5 selected>Global Add" & vbCrLf
	else
		Response.Write "											<option value=5>Global Add" & vbCrLf
	end if
	
	if Session("CurrentType") = "6" then
		Response.Write "											<option value=6 selected>Global Delete" & vbCrLf
	else
		Response.Write "											<option value=6>Global Delete" & vbCrLf
	end if
	
	if Session("CurrentType") = "7" then
		Response.Write "											<option value=7 selected>Global Update" & vbCrLf
	else
		Response.Write "											<option value=7>Global Update" & vbCrLf
	end if
	
	if Session("CurrentType") = "8" then
		Response.Write "											<option value=8 selected>Import" & vbCrLf
	else
		Response.Write "											<option value=8>Import" & vbCrLf
	end if

	'if Session("CurrentType") = "19" then
	'	Response.Write "											<option value=19 selected>Label Definition" & vbCrLf
	'else
	'	Response.Write "											<option value=19>Label Definition" & vbCrLf
	'end if
	
	if Session("CurrentType") = "9" then
		Response.Write "											<option value=9 selected>Mail Merge" & vbCrLf
	else
		Response.Write "											<option value=9>Mail Merge" & vbCrLf
	end if

	if Session("CurrentType") = "16" then
		Response.Write "											<option value=16 selected>Match Report" & vbCrLf
	else
		Response.Write "											<option value=16>Match Report" & vbCrLf
	end if
	
	'if Session("CurrentType") = "14" then
	'	Response.Write "											<option value=14 selected>Record Editing" & vbCrLf
	'else
	'	Response.Write "											<option value=14>Record Editing" & vbCrLf
	'end if

	if Session("CurrentType") = "20" then
		Response.Write "											<option value=20 selected>Record Profile" & vbCrLf
	else
		Response.Write "											<option value=20>Record Profile" & vbCrLf
	end if

	if Session("CurrentType") = "13" then
		Response.Write "											<option value=13 selected>Standard Report" & vbCrLf
	else
		Response.Write "											<option value=13>Standard Report" & vbCrLf
	end if

	if Session("CurrentType") = "21" then
		Response.Write "											<option value=21 selected>Succession Planning" & vbCrLf
	else
		Response.Write "											<option value=21>Succession Planning" & vbCrLf
	end if

	if Session("CurrentType") = "15" then
		Response.Write "											<option value=15 selected>System Error" & vbCrLf
	else
		Response.Write "											<option value=15>System Error" & vbCrLf
	end if

	if session("WF_Enabled") then
		if Session("CurrentType") = "25" then
			Response.Write "											<option value=25 selected>Workflow Rebuild" & vbCrLf
		else
			Response.Write "											<option value=25>Workflow Rebuild" & vbCrLf
		end if
	end if
		
%>												
												</select>		
											</TD>
											<TD width=25>
												Mode : 
											</TD>
											<TD>
												<select id=cboMode name=cboMode class="combo" style="WIDTH: 100%" onchange="refreshGrid();">
<%
	if Session("CurrentMode") = "-1" then
		Response.Write "											<option value=-1 selected>&lt;All&gt;" & vbCrLf
	else
		Response.Write "											<option value=-1>&lt;All&gt;" & vbCrLf 
	end if
	
	if Session("CurrentMode") = "1" then
		Response.Write "											<option value=1 selected>Batch" & vbCrLf
	else
		Response.Write "											<option value=1>Batch" & vbCrLf
	end if
	
	if Session("CurrentMode") = "0" then
		Response.Write "											<option value=0 selected>Manual" & vbCrLf
	else
		Response.Write "											<option value=0>Manual" & vbCrLf
	end if
%>									
												</select>		
											</TD>
											<TD width=25>
												Status : 
											</TD>
											<TD>
												<select id=cboStatus name=cboStatus class="combo" style="WIDTH: 100%" onchange="refreshGrid();">
<%	
	if Session("CurrentStatus") = "-1" then
		Response.Write "											<option value=-1 selected>&lt;All&gt;" & vbCrLf
	else
		Response.Write "											<option value=-1>&lt;All&gt;" & vbCrLf
	end if
		
	if Session("CurrentStatus") = "1" then
		Response.Write "											<option value=1 selected>Cancelled" & vbCrLf
	else
		Response.Write "											<option value=1>Cancelled" & vbCrLf
	end if
	
	if Session("CurrentStatus") = "5" then
		Response.Write "											<option value=5 selected>Error" & vbCrLf
	else
		Response.Write "											<option value=5>Error" & vbCrLf
	end if
	
	if Session("CurrentStatus") = "2" then
		Response.Write "											<option value=2 selected>Failed" & vbCrLf
	else
		Response.Write "											<option value=2>Failed" & vbCrLf
	end if
	
	if Session("CurrentStatus") = "0" then
		Response.Write "											<option value=0 selected>Pending" & vbCrLf
	else
		Response.Write "											<option value=0>Pending" & vbCrLf
	end if
	
	if Session("CurrentStatus") = "4" then
		Response.Write "											<option value=4 selected>Skipped" & vbCrLf
	else
		Response.Write "											<option value=4>Skipped" & vbCrLf
	end if
	
	if Session("CurrentStatus") = "3" then
		Response.Write "											<option value=3 selected>Successful" & vbCrLf
	else
		Response.Write "											<option value=3>Successful" & vbCrLf
	end if
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
		
	Response.Write "											<OBJECT classid=clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" & vbCrLf
	Response.Write "													 codebase=""cabs/COAInt_Grid.cab#version=3,1,3,6""" & vbCrLf
	Response.Write "													height=""100%""" & vbCrLf 
	Response.Write "													id=ssOleDBGridEventLog" & vbCrLf
	Response.Write "													name=ssOleDBGridEventLog" & vbCrLf
	Response.Write "													style=""HEIGHT: 100%; VISIBILITY: visible; WIDTH: 100%""" & vbCrLf 
	Response.Write "													width=""100%"">" & vbCrLf
	Response.Write "												<PARAM NAME=""ScrollBars"" VALUE=""3"">" & vbCrLf
	Response.Write "												<PARAM NAME=""_Version"" VALUE=""196617"">" & vbCrLf
	Response.Write "												<PARAM NAME=""DataMode"" VALUE=""2"">" & vbCrLf
	Response.Write "												<PARAM NAME=""Cols"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""Rows"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""BorderStyle"" VALUE=""1"">" & vbCrLf
	Response.Write "												<PARAM NAME=""RecordSelectors"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""GroupHeaders"" VALUE=""-1"">" & vbCrLf
	Response.Write "												<PARAM NAME=""ColumnHeaders"" VALUE=""-1"">" & vbCrLf
	Response.Write "												<PARAM NAME=""GroupHeadLines"" VALUE=""1"">" & vbCrLf
	Response.Write "												<PARAM NAME=""HeadLines"" VALUE=""1"">" & vbCrLf
	Response.Write "												<PARAM NAME=""FieldDelimiter"" VALUE=""(None)"">" & vbCrLf
	Response.Write "												<PARAM NAME=""FieldSeparator"" VALUE=""(Tab)"">" & vbCrLf
	Response.Write "												<PARAM NAME=""Row.Count"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""Col.Count"" VALUE=""1"">" & vbCrLf
	Response.Write "												<PARAM NAME=""stylesets.count"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""TagVariant"" VALUE=""EMPTY"">" & vbCrLf
	Response.Write "												<PARAM NAME=""UseGroups"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""HeadFont3D"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""Font3D"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""DividerType"" VALUE=""3"">" & vbCrLf
	Response.Write "												<PARAM NAME=""DividerStyle"" VALUE=""1"">" & vbCrLf
	Response.Write "												<PARAM NAME=""DefColWidth"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""BeveColorScheme"" VALUE=""2"">" & vbCrLf
	Response.Write "												<PARAM NAME=""BevelColorFrame"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""BevelColorHighlight"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""BevelColorShadow"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""BevelColorFace"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""CheckBox3D"" VALUE=""-1"">" & vbCrLf
	Response.Write "												<PARAM NAME=""AllowAddNew"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""AllowDelete"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""AllowUpdate"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""MultiLine"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""ActiveCellStyleSet"" VALUE="""">" & vbCrLf
	Response.Write "												<PARAM NAME=""RowSelectionStyle"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""AllowRowSizing"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""AllowGroupSizing"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""AllowColumnSizing"" VALUE=""-1"">" & vbCrLf
	Response.Write "												<PARAM NAME=""AllowGroupMoving"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""AllowColumnMoving"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""AllowGroupSwapping"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""AllowColumnSwapping"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""AllowGroupShrinking"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""AllowColumnShrinking"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""AllowDragDrop"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""UseExactRowCount"" VALUE=""-1"">" & vbCrLf
	Response.Write "												<PARAM NAME=""SelectTypeCol"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""SelectTypeRow"" VALUE=""3"">" & vbCrLf
	Response.Write "												<PARAM NAME=""SelectByCell"" VALUE=""-1"">" & vbCrLf
	Response.Write "												<PARAM NAME=""BalloonHelp"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""RowNavigation"" VALUE=""2"">" & vbCrLf
	Response.Write "												<PARAM NAME=""CellNavigation"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""MaxSelectedRows"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""HeadStyleSet"" VALUE="""">" & vbCrLf
	Response.Write "												<PARAM NAME=""StyleSet"" VALUE="""">" & vbCrLf
	Response.Write "												<PARAM NAME=""ForeColorEven"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""ForeColorOdd"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""BackColorEven"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""BackColorOdd"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""Levels"" VALUE=""1"">" & vbCrLf
	Response.Write "												<PARAM NAME=""RowHeight"" VALUE=""503"">" & vbCrLf
	Response.Write "												<PARAM NAME=""ExtraHeight"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""ActiveRowStyleSet"" VALUE="""">" & vbCrLf
	Response.Write "												<PARAM NAME=""CaptionAlignment"" VALUE=""2"">" & vbCrLf
	Response.Write "												<PARAM NAME=""SplitterPos"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""SplitterVisible"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""Columns.Count"" VALUE=""" & (UBound(avColumnDef)	+ 1) & """>" & vbCrLf
	
	for i = 0 to ubound(avColumnDef) step 1
		Response.Write "												<!--" & avColumnDef(i,0) & "-->  " & vbCrLf      
		Response.Write "												<PARAM NAME=""Columns(" & i & ").Width"" VALUE=""" & avColumnDef(i,2) & """>" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").Visible"" VALUE=""" & avColumnDef(i,3) & """>" & vbCrLf 
		Response.Write "												<PARAM NAME=""Columns(" & i & ").Columns.Count"" VALUE=""1"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").Caption"" VALUE=""" & avColumnDef(i,1) & """>" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").Name"" VALUE=""" & avColumnDef(i,0) & """>" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").Alignment"" VALUE=""0"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").CaptionAlignment"" VALUE=""3"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").Bound"" VALUE=""0"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").AllowSizing"" VALUE=""1"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").DataField"" VALUE=""Column " & i & """>" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").DataType"" VALUE=""8"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").Level"" VALUE=""0"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").NumberFormat"" VALUE="""">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").Case"" VALUE=""0"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").FieldLen"" VALUE=""256"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").VertScrollBar"" VALUE=""0"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").Locked"" VALUE=""0"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").Style"" VALUE=""0"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").ButtonsAlways"" VALUE=""0"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").RowCount"" VALUE=""0"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").ColCount"" VALUE=""1"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").HasHeadForeColor"" VALUE=""0"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").HasHeadBackColor"" VALUE=""0"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").HasForeColor"" VALUE=""0"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").HasBackColor"" VALUE=""0"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").HeadForeColor"" VALUE=""0"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").HeadBackColor"" VALUE=""0"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").ForeColor"" VALUE=""0"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").BackColor"" VALUE=""0"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").HeadStyleSet"" VALUE="""">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").StyleSet"" VALUE="""">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").Nullable"" VALUE=""1"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").Mask"" VALUE="""">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").PromptInclude"" VALUE=""0"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").ClipMode"" VALUE=""0"">" & vbCrLf
		Response.Write "												<PARAM NAME=""Columns(" & i & ").PromptChar"" VALUE=""95"">" & vbCrLf
	next
		
	Response.Write "												<PARAM NAME=""UseDefaults"" VALUE=""-1"">" & vbCrLf
	Response.Write "												<PARAM NAME=""TabNavigation"" VALUE=""1"">" & vbCrLf
	Response.Write "												<PARAM NAME=""BatchUpdate"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""_ExtentX"" VALUE=""11298"">" & vbCrLf
	Response.Write "												<PARAM NAME=""_ExtentY"" VALUE=""3969"">" & vbCrLf
	Response.Write "												<PARAM NAME=""_StockProps"" VALUE=""79"">" & vbCrLf
	Response.Write "												<PARAM NAME=""Caption"" VALUE="""">" & vbCrLf
	Response.Write "												<PARAM NAME=""ForeColor"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""BackColor"" VALUE=""0"">" & vbCrLf
	Response.Write "												<PARAM NAME=""Enabled"" VALUE=""-1"">" & vbCrLf
	Response.Write "												<PARAM NAME=""DataMember"" VALUE="""">" & vbCrLf

	Response.Write "												<PARAM NAME=""Row.Count"" VALUE=""0"">" & vbcrlf
	Response.Write "											</OBJECT>" & vbCrLf
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

<FORM action="default_Submit.asp" method=post id=frmGoto name=frmGoto style="visibility:hidden;display:none">
<!--#include file="include\gotoWork.txt"-->
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

<FORM id=frmPurge name=frmPurge method=post style="visibility:hidden;display:none" action="eventLog.asp">
	<INPUT type="hidden" id=txtDoesPurge name=txtDoesPurge>
	<INPUT type="hidden" id=txtPurgePeriod name=txtPurgePeriod>
	<INPUT type="hidden" id=txtPurgeFrequency name=txtPurgeFrequency>
	<INPUT type="hidden" id=txtShowPurgeMSG name=txtShowPurgeMSG value=<%=Session("showPurgeMessage")%>>
	<INPUT type="hidden" id=txtCurrentUsername name=txtCurrentUsername>
	<INPUT type="hidden" id=txtCurrentType name=txtCurrentType>
	<INPUT type="hidden" id=txtCurrentMode name=txtCurrentMode>
	<INPUT type="hidden" id=txtCurrentStatus name=txtCurrentStatus>
</FORM>

<FORM id=frmDelete name=frmDelete method=post style="visibility:hidden;display:none" action="eventLog.asp">
	<INPUT type="hidden" id=txtDeleteSel name=txtDeleteSel>
	<INPUT type="hidden" id=txtSelectedIDs name=txtSelectedIDs>
	<INPUT type="hidden" id=txtViewAllPerm name=txtViewAllPerm>
	<INPUT type="hidden" id=Hidden1 name=txtCurrentUsername>
	<INPUT type="hidden" id=Hidden2 name=txtCurrentType>
	<INPUT type="hidden" id=Hidden3 name=txtCurrentMode>
	<INPUT type="hidden" id=Hidden4 name=txtCurrentStatus>
</FORM>

<FORM id=frmEmail name=frmEmail method=post style="visibility:hidden;display:none" action="emailSelection.asp">
	<INPUT type="hidden" id=txtSelectedEventIDs name=txtSelectedEventIDs>
	<INPUT type="hidden" id=txtFromMain name=txtFromMain value=1>
	<INPUT type="hidden" id=txtEmailOrderColumn name=txtEmailOrderColumn>
	<INPUT type="hidden" id=txtEmailOrderOrder name=txtEmailOrderOrder>
</FORM>

<FORM id=frmRefresh name=frmRefresh method=post style="visibility:hidden;display:none" action="eventLog.asp">
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
	Set cmdDefinition = Server.CreateObject("ADODB.Command")
	cmdDefinition.CommandText = "sp_ASRIntGetModuleParameter"
	cmdDefinition.CommandType = 4 ' Stored procedure.
	Set cmdDefinition.ActiveConnection = session("databaseConnection")

	Set prmModuleKey = cmdDefinition.CreateParameter("moduleKey",200,1,8000) ' 200=varchar, 1=input, 8000=size
	cmdDefinition.Parameters.Append prmModuleKey
	prmModuleKey.value = "MODULE_PERSONNEL"

	Set prmParameterKey = cmdDefinition.CreateParameter("paramKey",200,1,8000) ' 200=varchar, 1=input, 8000=size
	cmdDefinition.Parameters.Append prmParameterKey
	prmParameterKey.value = "Param_TablePersonnel"

	Set prmParameterValue = cmdDefinition.CreateParameter("paramValue",200,2,8000) '200=varchar, 2=output, 8000=size
	cmdDefinition.Parameters.Append prmParameterValue

	err = 0
	cmdDefinition.Execute

	Response.Write "<INPUT type='hidden' id=txtPersonnelTableID name=txtPersonnelTableID value=" & cmdDefinition.Parameters("paramValue").Value & ">" & vbcrlf
	
	set cmdDefinition = nothing

	Response.Write "<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbcrlf
	Response.Write "<INPUT type='hidden' id=txtAction name=txtAction value=" & session("action") & ">" & vbcrlf
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

</BODY>
</HTML>

<!-- Embeds createActiveX.js script reference -->
<!--#include file="include\ctl_CreateControl.txt"-->

<% 

function formatError(psErrMsg)
  Dim iStart 
  dim iFound 
  
  iFound = 0
  Do
    iStart = iFound
    iFound = InStr(iStart + 1, psErrMsg, "]")
  Loop While iFound > 0
  
  If (iStart > 0) And (iStart < Len(Trim(psErrMsg))) Then
    formatError = Trim(Mid(psErrMsg, iStart + 1))
  Else
    formatError = psErrMsg
  End If
end function
%>


</asp:Content>
