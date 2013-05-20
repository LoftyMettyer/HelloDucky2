<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script type="text/javascript">

    function stdrpt_AbsenceCalendar_window_onload() {

        var fOK;
        fOK = true;

        // Permission denied on absence table
        if (frmChangeDetails.txtReportFailed.value == 'True') {
            OpenHR.messageBox(frmChangeDetails.txtErrorMSG.value, 48, "Absence Calendar");
            absence_calendar_OKClick();
            return;
        }

        if (fOK == true) {

            // Set focus onto one of the form controls. 
            cmdOK.focus();

            // Get menu.asp to refresh the menu.
            menu_refreshMenu();

            refreshDateSpecifics();

            // Expand the option frame and hide the work frame.
            $("#workframe").hide();
            $("#optionframe").show();

        }

        // Disable the menu
        menu_disableMenu();

        // Force this combo to be displayed.
        //cboStartMonth.style.visibility = "visible";
    }

    function cboStartMonth_onchange() 
    {
        frmChangeDetails.txtStartMonth.value = cboStartMonth.value;
        OpenHR.submitForm(frmChangeDetails);
    }

    function cmdPreviousYear_onclick() {
        frmChangeDetails.txtStartYear.value = Number(frmChangeDetails.txtStartYear.value) - 1;
        OpenHR.submitForm(frmChangeDetails);
    }

    function cmdNextYear_onclick() {
        frmChangeDetails.txtStartYear.value = Number(frmChangeDetails.txtStartYear.value) + 1;
        OpenHR.submitForm(frmChangeDetails);
    }

    function refreshToggleValues() {
        // Show Captions setting
        if (chkShowCaptions.checked == false) {
            frmChangeDetails.txtShowCaptions.value = "hide";
        }
        else {
            frmChangeDetails.txtShowCaptions.value = "show";
        }

        // Show Weekends setting
        if (chkShowWeekends.checked == false) {
            frmChangeDetails.txtShowWeekends.value = "unhighlighted";
        }
        else {
            frmChangeDetails.txtShowWeekends.value = "highlighted";
        }

        // Include Bank Holidays setting
        if (chkIncludeBankHolidays.checked == false) {
            frmChangeDetails.txtIncludeBankHolidays.value = "unincluded";
        }
        else {
            frmChangeDetails.txtIncludeBankHolidays.value = "included";
        }

        // Show Bank Holidays setting
        if (chkShowBankHolidays.checked == false) {
            frmChangeDetails.txtShowBankHolidays.value = "unhighlighted";
        }
        else {
            frmChangeDetails.txtShowBankHolidays.value = "highlighted";
        }

        // Working Days Only setting
        if (chkIncludeWorkingDaysOnly.checked == false) {
            frmChangeDetails.txtIncludeWorkingDaysOnly.value = "unincluded";
        }
        else {
            frmChangeDetails.txtIncludeWorkingDaysOnly.value = "included";
        }
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

    function ShowDetails(pdStartDate, pstrStartSession, pdEndDate, pstrEndSession, intDuration, strType, strTypeCode, strCalCode, strReason, strRegion, strWorkingPattern) 
    {
        var sURL;

        // Populate the form with the day's details
        frmAbsenceDetails.txtStartDate.value = pdStartDate;
        frmAbsenceDetails.txtStartSession.value = pstrStartSession;
        frmAbsenceDetails.txtEndDate.value = pdEndDate;
        frmAbsenceDetails.txtEndSession.value = pstrEndSession;
        frmAbsenceDetails.txtDuration.value = intDuration;
        frmAbsenceDetails.txtType.value = strType;
        frmAbsenceDetails.txtTypeCode.value = strTypeCode;
        frmAbsenceDetails.txtCalCode.value = strCalCode;
        frmAbsenceDetails.txtReason.value = strReason;
        frmAbsenceDetails.txtRegion.value = strRegion;
        frmAbsenceDetails.txtWorkingPattern.value = strWorkingPattern;
	
        sURL = "stdrpt_AbsenceCalendar_Details" +
            "?txtStartDate=" + frmAbsenceDetails.txtStartDate.value +
            "&txtStartSession=" + escape(frmAbsenceDetails.txtStartSession.value) +
            "&txtEndDate=" + frmAbsenceDetails.txtEndDate.value +
            "&txtEndSession=" + escape(frmAbsenceDetails.txtEndSession.value) +
            "&txtDuration=" + frmAbsenceDetails.txtDuration.value +
            "&txtType=" + escape(frmAbsenceDetails.txtType.value) +
            "&txtTypeCode=" + escape(frmAbsenceDetails.txtTypeCode.value) +
            "&txtCalCode=" + escape(frmAbsenceDetails.txtCalCode.value) +
            "&txtReason=" + escape(frmAbsenceDetails.txtReason.value) +
            "&txtDisableRegions=" + escape(frmChangeDetails.txtDisableRegions.value) +
            "&txtRegion=" + escape(frmAbsenceDetails.txtRegion.value) +
            "&txtDisableWPs=" + escape(frmChangeDetails.txtDisableWPs.value) +
            "&txtWorkingPattern=" + escape(frmAbsenceDetails.txtWorkingPattern.value);
        openDialog(sURL, 350,300);
    }

    // Returns to the recordedit screen
    function absence_calendar_OKClick() {

        refreshData();

        $("#optionframe").hide();
        $("#workframe").show();
        OpenHR.submitForm(frmRecordEdit);
    }

    // Prints the screen
    function PrintGrid() {
        window.print();
    }

</script>

<%
	if Session("stdrpt_AbsenceCalendar_StartMonth") = "" then
	
        Dim cmdDefinition As Object
        Dim prmModuleKey As Object
        Dim prmParameterKey As Object
        Dim prmParameterValue As Object
        
        cmdDefinition = Server.CreateObject("ADODB.Command")
		cmdDefinition.CommandText = "sp_ASRIntGetModuleParameter"
		cmdDefinition.CommandType = 4 ' Stored procedure.
        cmdDefinition.ActiveConnection = Session("databaseConnection")

        prmModuleKey = cmdDefinition.CreateParameter("moduleKey", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
        cmdDefinition.Parameters.Append(prmModuleKey)
		prmModuleKey.value = "MODULE_ABSENCE"

        prmParameterKey = cmdDefinition.CreateParameter("paramKey", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
        cmdDefinition.Parameters.Append(prmParameterKey)
		prmParameterKey.value = "Param_FieldStartMonth"

        prmParameterValue = cmdDefinition.CreateParameter("paramValue", 200, 2, 8000) '200=varchar, 2=output, 8000=size
        cmdDefinition.Parameters.Append(prmParameterValue)

        Err.Clear()
		cmdDefinition.Execute

		Session("stdrpt_AbsenceCalendar_StartMonth") = cmdDefinition.Parameters("paramValue").Value
        If Month(Now) < CInt(Session("stdrpt_AbsenceCalendar_StartMonth")) Then
            Session("stdrpt_AbsenceCalendar_StartYear") = Year(Now) - 1
        Else
            Session("stdrpt_AbsenceCalendar_StartYear") = Year(Now)
        End If

        cmdDefinition = Nothing

	end if

	' Create absence calendar object
    Dim objAbsenceCalendar As HR.Intranet.Server.AbsenceCalendar
    objAbsenceCalendar = New HR.Intranet.Server.AbsenceCalendar()

	' Pass required info to the DLL
    objAbsenceCalendar.Username = Session("username").ToString()
	objAbsenceCalendar.Connection = session("databaseConnection")

	' Pass in the recordID for the current record
	objAbsenceCalendar.RealSource = Session("optionRealsource")
	objAbsenceCalendar.RecordID = Session("optionRecordID")
    objAbsenceCalendar.ClientDateFormat = Session("LocaleDateFormat").ToString()
    objAbsenceCalendar.StartMonth = Session("stdrpt_AbsenceCalendar_StartMonth")
	objAbsenceCalendar.StartYear = Session("stdrpt_AbsenceCalendar_StartYear")

	objAbsenceCalendar.Initialise()

	objAbsenceCalendar.IncludeBankHolidays = session("stdrpt_AbsenceCalendar_IncludeBankHolidays")
	objAbsenceCalendar.IncludeWorkingDaysOnly = session("stdrpt_AbsenceCalendar_IncludeWorkingDaysOnly")
	objAbsenceCalendar.ShowBankHolidays = session("stdrpt_AbsenceCalendar_ShowBankHolidays") 		
	objAbsenceCalendar.ShowWeekends = session("stdrpt_AbsenceCalendar_ShowWeekends")
	objAbsenceCalendar.ShowCaptions = session("stdrpt_AbsenceCalendar_ShowCaptions")
	
	objAbsenceCalendar.StartMonth = Session("stdrpt_AbsenceCalendar_StartMonth")
	objAbsenceCalendar.StartYear = Session("stdrpt_AbsenceCalendar_StartYear")

    
if objAbsenceCalendar.ReportFailed = false then 
%>
    <table valign=top align="center" class="outline" id="Background" cellSpacing="2" cellPadding="0">
        <tr>
	        <td valign=top><!-- Display the month details -->
<%
    Response.Write(objAbsenceCalendar.HTML_Calendar)
%>
            </td>
	        <td valign=top>
		        <Table class="invisible"> 
		            <TR height=8>
		                <TD colspan=2></TD>
		            </TR>
		            <TR>
		                <TD colspan=2>
			                <!-- Draw the Employee information box -->
			                <table valign=top width="250" class="outline" id="tblEmpoyeeInformation" cellSpacing="2" cellPadding="0">
<%
        ' Write a row in this table for the forward/back year controls
    Response.Write(objAbsenceCalendar.HTML_ForwardBackYear)
    ' Stuff the employee information
    Response.Write(objAbsenceCalendar.HTML_EmployeeInformation)
%>		  
                            </table>
		                </TD>
		            </TR>
    		
		            <TR height=3>
		                <TD colspan=2></TD>
		            </TR>
		            <TR>
		                <TD colspan=2>
			                <!-- Draw the option checkboxes -->
			                <table width="250" class="outline" id="tblOptions" cellPadding="0" cellSpacing="2">
			                    <tr>		  
			                        <td>&nbsp;Start Month</td>
			                        <td>
<%
        ' Load the start month combo
    Response.Write(objAbsenceCalendar.HTML_SelectedStartMonthCombo(objAbsenceCalendar.StartMonth))
%>
                                    </td>
			                    </tr>
			 	                <!-- Show the display options -->
<%
    Response.Write(objAbsenceCalendar.HTML_DisplayOptions)
%>
                            </table>
		                </TD>
		            </TR>
    				
		            <TR height=3>
		                <TD colspan=2></TD>
		            </TR>
                    <TR>
		                <TD colspan=2>
<% 
        ' Generate HTML for the absence key types
    Response.Write(objAbsenceCalendar.HTML_LoadColourKey)
%>
                        </TD>
		            </TR>
		            <TR height=3>
		                <TD colspan=2></TD>
		            </TR>
                    <TR>
		                <!-- OK/Print Buttons -->
		                <td colspan=2 align=right>
<%--				            <input id="cmdPrint" name="cmdPrint" type="button" value="Print" style="HEIGHT: 25px; WIDTH: 80px" class="btn"
				                onclick="PrintGrid()" 
                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
                            &nbsp; --%>
				            <input id="cmdOK" name="cmdOK" type="button" value="Back" style="HEIGHT: 25px; WIDTH: 80px" class="btn"
				                onclick="absence_calendar_OKClick()" 
                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                onfocus="try{button_onFocus(this);}catch(e){}"
                                onblur="try{button_onBlur(this);}catch(e){}" />
                        </td>
		            </TR>
	            </TABLE>
            </TD>
	    </tr>
    </table>

<%
	'Populate the grid with data
	objAbsenceCalendar.StartYear = session("stdrpt_AbsenceCalendar_StartYear")
	objAbsenceCalendar.StartMonth = session("stdrpt_AbsenceCalendar_StartMonth")
	
	' Write navigation/option functions
    Response.Write(objAbsenceCalendar.HTML_ToggleDisplay)
end if 
%>

<!-- Data for the absence calendar -->
<form action="stdrpt_AbsenceCalendar_submit" method="post" id="frmChangeDetails" name="frmChangeDetails">
    <input type="hidden" id="txtStartMonth" name="txtStartMonth" value="<%Response.Write(objAbsenceCalendar.StartMonth)%>">
    <input type="hidden" id="txtStartYear" name="txtStartYear" value="<%Response.Write(objAbsenceCalendar.StartYear)%>">
    <input type="hidden" id="txtIncludeBankHolidays" name="txtIncludeBankHolidays" value="<%Response.Write(Session("stdrpt_AbsenceCalendar_IncludeBankHolidays"))%>">
    <input type="hidden" id="txtIncludeWorkingDaysOnly" name="txtIncludeWorkingDaysOnly" value="<%Response.Write(Session("stdrpt_AbsenceCalendar_IncludeWorkingDaysOnly"))%>">
    <input type="hidden" id="txtShowBankHolidays" name="txtShowBankHolidays" value="<%Response.Write(Session("stdrpt_AbsenceCalendar_ShowBankHolidays"))%>">
    <input type="hidden" id="txtShowCaptions" name="txtShowCaptions" value="<%Response.Write(Session("stdrpt_AbsenceCalendar_ShowCaptions"))%>">
    <input type="hidden" id="txtShowWeekends" name="txtShowWeekends" value="<%Response.Write(Session("stdrpt_AbsenceCalendar_ShowWeekends"))%>">
    <input type="hidden" id="txtAbsenceRecordsFound" name="txtAbsenceRecordsFound" value="<%Response.Write(objAbsenceCalendar.AbsenceRecordCount)%>">
    <input type="hidden" id="txtReportFailed" name="txtReportFailed" value="<%Response.Write(objAbsenceCalendar.ReportFailed)%>">
    <input type="hidden" id="txtErrorMSG" name="txtErrorMSG" value="<%Response.Write(objAbsenceCalendar.ErrorMSG)%>">
    <input type="hidden" id="txtDisableRegions" name="txtDisableRegions" value="<%Response.Write(objAbsenceCalendar.DisableRegions)%>">
    <input type="hidden" id="txtDisableWPs" name="txtDisableWPs" value="<%Response.Write(objAbsenceCalendar.DisableWPs)%>">
</form>

<form action="stdrpt_AbsenceCalendar_submit" method="post" id="frmGotoOption" name="frmGotoOption" style="visibility: hidden; display: none">
    <%Html.RenderPartial("~/Views/Shared/gotoOption.ascx")%>
</form>

<!-- Form to return to record edit screen -->
<form action="emptyoption" method="post" id="frmRecordEdit" name="frmRecordEdit">
</form>

<form action="stdrpt_AbsenceCalendar_Details" target="ShowDetails" method="post" id="frmAbsenceDetails" name="frmAbsenceDetails">
    <input type="hidden" id="txtStartDate" name="txtStartDate">
    <input type="hidden" id="txtStartSession" name="txtStartSession">
    <input type="hidden" id="txtEndDate" name="txtEndDate">
    <input type="hidden" id="txtEndSession" name="txtEndSession">
    <input type="hidden" id="txtDuration" name="txtDuration">
    <input type="hidden" id="txtType" name="txtType">
    <input type="hidden" id="txtTypeCode" name="txtTypeCode">
    <input type="hidden" id="txtCalCode" name="txtcalCode">
    <input type="hidden" id="txtReason" name="txtReason">
    <input type="hidden" id="txtRegion" name="txtRegion">
    <input type="hidden" id="txtWorkingPattern" name="txtWorkingPattern">
</form>


<% 
	' Cleanup code
    objAbsenceCalendar = Nothing
%>

<script type="text/javascript">
    stdrpt_AbsenceCalendar_window_onload();
</script>