<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<%
    Dim objCalendar As Object
    objCalendar = Session("objCalendar" & Session("CalRepUtilID"))
%>


<object classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" 
	id="Microsoft_Licensed_Class_Manager_1_0">
	<param NAME="LPKPath" VALUE="lpks/main.lpk">
</object>


<script type="text/javascript">

    function util_run_calendarreport_window_onload() {
        loadAddRecords('nav');      
    }

    function ExportDataPrompt() 
    {

        var sURL = "util_run_outputoptions" +
            "?txtUtilType=" + escape(frmExportData.txtUtilType.value) +
            "&txtPreview=" + escape(frmExportData.txtPreview.value) +
            "&txtFormat=" + escape(frmExportData.txtFormat.value) +
            "&txtScreen=" + escape(frmExportData.txtScreen.value) +
            "&txtPrinter=" + escape(frmExportData.txtPrinter.value) +
            "&txtPrinterName=" + escape(frmExportData.txtPrinterName.value) +
            "&txtSave=" + escape(frmExportData.txtSave.value) +
            "&txtSaveExisting=" + escape(frmExportData.txtSaveExisting.value) +
            "&txtEmail=" + escape(frmExportData.txtEmail.value) +
            "&txtEmailAddr=" + escape(frmExportData.txtEmailAddr.value) +
            "&txtEmailAddrName=" + escape(frmExportData.txtEmailAddrName.value) +
            "&txtEmailSubject=" + escape(frmExportData.txtEmailSubject.value) +
            "&txtEmailAttachAs=" + escape(frmExportData.txtEmailAttachAs.value) +
            "&txtFileName=" + escape(frmExportData.txtFileName.value);
        ShowOutputOptionsFrame(sURL);
    }

    function openDialog(pDestination, pWidth, pHeight, psResizable, psScroll)
    {
        dlgwinprops = "center:yes;" +
            "dialogHeight:" + pHeight + "px;" +
            "dialogWidth:" + pWidth + "px;" +
            "help:no;" +
            "resizable:" + psResizable + ";" +
            "scroll:" + psScroll + ";" +
            "status:no;";
        window.showModalDialog(pDestination, self, dlgwinprops);
    }

    function dataOnlyPrint()
    {
        var lngPageColumnCount = 3;
        var lngActualRow = new Number(0);
        var blnSettingsDone = false;
        var sColHeading = new String('');
        var iColDataType = new Number(12);
        var iColDecimals = new Number(0);
        var blnNewPage = true;
        var lngPageCount = new Number(0);
        var frmOutput = OpenHR.getForm("dataframe","frmCalendarData");
	
        frmOutput.grdCalendarOutput.focus();
	
        if (frmOutput.ssHiddenGrid.Columns.Count > 0) 
        {
            frmOutput.ssHiddenGrid.Columns.RemoveAll();
        }
		
        if (frmOutput.ssHiddenGrid.Rows > 0)
        {	
            frmOutput.ssHiddenGrid.RemoveAll();
        }
		
        frmOutput.ssHiddenGrid.Columns.RemoveAll();
        frmOutput.ssHiddenGrid.Font.Name = "Verdana";
        frmOutput.ssHiddenGrid.Font.Size = 8;
        frmOutput.ssHiddenGrid.Font.Bold = false;
        frmOutput.ssHiddenGrid.Font.Underline = false;
        frmOutput.ssHiddenGrid.focus();
	
        frmOriginalDefinition.txtOptionsDone.value = 0;

        // Need to loop through the grid, selecting rows until we find a '*' in
        // the first column ('PageBreak').  
        frmOriginalDefinition.txtCancelPrint.value = 0;
        frmOutput.grdCalendarOutput.MoveFirst();
        for (var lngRow=0; lngRow<frmOutput.grdCalendarOutput.Rows; lngRow++)
        {
            bm = frmOutput.grdCalendarOutput.AddItemBookmark(lngRow);
		
            if (lngRow == (frmOutput.grdCalendarOutput.Rows - 1))
            {
                sBreakValue = frmOutput.grdCalendarOutput.Columns(1).CellText(bm);
                frmOutput.ssHiddenGrid.Caption = txtTitle.value + ' - ' + sBreakValue;
                // PRINT DATA
                if (frmOriginalDefinition.txtOptionsDone.value == 0) 
                {
                    frmOutput.ssHiddenGrid.PrintData(23,false,true);	
                    frmOriginalDefinition.txtOptionsDone.value = 1;
                    if (frmOriginalDefinition.txtCancelPrint.value == 1) 
                    {
                        frmOutput.grdCalendarOutput.redraw = true;
                        return;
                    }
                }
                else 
                {
                    if (frmOriginalDefinition.txtCancelPrint.value == 1) 
                    {
                        frmOutput.grdCalendarOutput.redraw = true;
                        return;
                    }

                    frmOutput.ssHiddenGrid.PrintData(23,false,false);	
                }
                frmOutput.ssHiddenGrid.RemoveAll();			
                blnBreakCheck = true;
                sBreakValue = '';
                lngActualRow = 0;
            }
            else if ((frmOutput.grdCalendarOutput.Columns(0).CellText(bm) == '*') 
                  && (!blnBreakCheck))
            {
                sBreakValue = frmOutput.grdCalendarOutput.Columns(1).CellText(bm);
                frmOutput.ssHiddenGrid.Caption = txtTitle.value + ' - ' + sBreakValue;
                // PRINT DATA
                if (frmOriginalDefinition.txtOptionsDone.value == 0) 
                {
                    frmOutput.ssHiddenGrid.PrintData(23,false,true);	
                    frmOriginalDefinition.txtOptionsDone.value = 1;
                    if (frmOriginalDefinition.txtCancelPrint.value == 1) 
                    {
                        frmOutput.grdCalendarOutput.redraw = true;
                        return;
                    }
                }
                else 
                {
                    if (frmOriginalDefinition.txtCancelPrint.value == 1) 
                    {
                        frmOutput.grdCalendarOutput.redraw = true;
                        return;
                    }

                    frmOutput.ssHiddenGrid.PrintData(23,false,false);	
                }
                frmOutput.ssHiddenGrid.RemoveAll();
                frmOutput.ssHiddenGrid.Columns.RemoveAll();
			
                lngPageColumnCount = 38;
                lngPageCount++;
                blnBreakCheck = true;
                sBreakValue = '';
                lngActualRow = 0;
                blnNewPage = true;
            } 
            else if (frmOutput.grdCalendarOutput.Columns(0).CellText(bm) != '*')
            {
                if (blnNewPage)
                {
                    for (var lngCol=0; lngCol<lngPageColumnCount; lngCol++)
                    {
                        // ADD COLUMN TO GRID
                        frmOutput.ssHiddenGrid.Columns.Add(lngCol);
                        frmOutput.ssHiddenGrid.Columns(lngCol).width = 15;
                    }
                }
                blnBreakCheck = false;
                blnNewPage = false;

                // CREATE TAB DELIMITED STRING AND ADD TO GRID
                sAddItem = new String("");
                for (var lngCol=0; lngCol<lngPageColumnCount; lngCol++)
                {
                    if (frmOutput.grdCalendarOutput.Columns(lngCol).visible)
                    {
                        if (lngCol > 0) 
                        {
                            sAddItem = sAddItem + "	";
                        }

                        var sValue = new String(trim(frmOutput.grdCalendarOutput.Columns(lngCol).CellText(bm)));
                        sAddItem = sAddItem + sValue;
                        if ((sValue.length * 8) > frmOutput.ssHiddenGrid.Columns(lngCol).width)
                        {
                            frmOutput.ssHiddenGrid.Columns(lngCol).width = (sValue.length * 8);
                        }
                    }
                }
                frmOutput.ssHiddenGrid.AddItem(sAddItem);
            }
		
            if (!blnNewPage) 
            {
                lngActualRow = lngActualRow + 1; 
            }
        }
    }

    function trim(strInput)
    {
        if (strInput.length < 1)
        {
            return "";
        }
		
        while (strInput.substr(strInput.length-1, 1) == " ") 
        {
            strInput = strInput.substr(0, strInput.length - 1);
        }
	
        while (strInput.substr(0, 1) == " ") 
        {
            strInput = strInput.substr(1, strInput.length);
        }
	
        return strInput;
    }

    function startCalendar()
    {
        var dtReportStart = createDate(frmDate.txtReportStartDate.value);	
        var dtReportEnd = createDate(frmDate.txtReportEndDate.value);
        var dtSystem = new Date();
	
        //TM20070514 - Fault 12236 fixed.
        var fStartMonthBeforeSysMonth = new Boolean();
        var fStartMonthIsSysMonth = new Boolean();
        var fEndMonthAfterSysMonth = new Boolean();
        var fEndMonthIsSysMonth = new Boolean();
        var fDefaultToSystemDate = new Boolean();
  
        fStartMonthBeforeSysMonth = (((dtReportStart.getFullYear() == dtSystem.getFullYear()) 
                  && (dtReportStart.getMonth() < dtSystem.getMonth())) 
              || (dtReportStart.getFullYear() < dtSystem.getFullYear()));
															
        fStartMonthIsSysMonth = ((dtReportStart.getFullYear() == dtSystem.getFullYear()) 
              && (dtReportStart.getMonth() == dtSystem.getMonth()));

        fEndMonthAfterSysMonth = (((dtReportEnd.getFullYear() == dtSystem.getFullYear()) 
                  && (dtReportEnd.getMonth() > dtSystem.getMonth())) 
          || (dtReportEnd.getFullYear() > dtSystem.getFullYear()));
                    
        fEndMonthIsSysMonth = ((dtReportEnd.getFullYear() == dtSystem.getFullYear()) 
              && (dtReportEnd.getMonth() == dtSystem.getMonth()));
													
        fDefaultToSystemDate = (((fStartMonthBeforeSysMonth && fEndMonthAfterSysMonth) 
             || (fStartMonthBeforeSysMonth && fEndMonthIsSysMonth) 
             || (fEndMonthAfterSysMonth && fStartMonthIsSysMonth)) 
          && (frmDate.txtStartOnCurrentMonth.value == 1));

        frmUseful.txtLoading.value = 1;
        frmUseful.txtCTLsPopulated.value = 0;

        populateCTL_Collections();
	
        if (fDefaultToSystemDate)
        {		
            thisMonth();
        }
        else
        {	
            firstMonth();
        }
	
        frmUseful.txtLoading.value = 0;
    }

    function addString(pintNumber,pstrChar)
    {
        var sRetString = new String('');
	
        for (var i=1; i<=pintNumber; i++)
        {
            sRetString = sRetString + pstrChar;
        }
	
        return sRetString;
    }
	
    function trim(strInput)
    {
        if (strInput.length < 1)
        {
            return "";
        }
		
        while (strInput.substr(strInput.length-1, 1) == " ") 
        {
            strInput = strInput.substr(0, strInput.length - 1);
        }
	
        while (strInput.substr(0, 1) == " ") 
        {
            strInput = strInput.substr(1, strInput.length);
        }
	
        return strInput;
    }

    function enableDisableNavigation()
    {
        var dtShownStart = createDate('01/'+frmNav.cboMonth.options[frmNav.cboMonth.selectedIndex].value+'/'+frmNav.txtYear.value);
        var dtShownEnd = createDate(frmDate.txtDaysInMonth.value+'/'+frmNav.cboMonth.options[frmNav.cboMonth.selectedIndex].value+'/'+frmNav.txtYear.value);
        var dtReportStart = createDate(frmDate.txtReportStartDate.value);	
        var dtReportEnd = createDate(frmDate.txtReportEndDate.value);
        var dtSystemMonth = new Date();
        var dtReportStartMonth = new Date(dtReportStart);
        var dtReportEndMonth = new Date(dtReportEnd);
        var intDaysInMonth;
        var bNextEnabled;
        var bPrevEnabled;
	
        if (dtShownStart <= dtReportStart)
        {
            bNextEnabled = false;
            document.getElementById('imgPrevMonth').src = "images/previous_disabled.gif";
            document.getElementById('imgFirstMonth').src = "images/first_disabled.gif";
            image_disable(document.getElementById('imgPrevMonth'), true);
            image_disable(document.getElementById('imgFirstMonth'), true);
            document.getElementById('imgPrevMonth').style.cursor = 'default';
            document.getElementById('imgFirstMonth').style.cursor = 'default';
        }
        else
        {
            bNextEnabled = true;
            document.getElementById('imgPrevMonth').src = "images/previous_enabled.gif";
            document.getElementById('imgFirstMonth').src = "images/first_enabled.gif";
            image_disable(document.getElementById('imgPrevMonth'), false);
            image_disable(document.getElementById('imgFirstMonth'), false);
            document.getElementById('imgPrevMonth').style.cursor = 'hand';
            document.getElementById('imgFirstMonth').style.cursor = 'hand';
        }
	
        if (dtShownEnd >= dtReportEnd)
        {
            bPrevEnabled = false;
            document.getElementById('imgNextMonth').src = "images/next_disabled.gif";
            document.getElementById('imgLastMonth').src = "images/last_disabled.gif";
            image_disable(document.getElementById('imgNextMonth'), true);
            image_disable(document.getElementById('imgLastMonth'), true);
            document.getElementById('imgNextMonth').style.cursor = 'default';
            document.getElementById('imgLastMonth').style.cursor = 'default';
        }
        else
        {
            bPrevEnabled = true;
            document.getElementById('imgNextMonth').src = "images/next_enabled.gif";
            document.getElementById('imgLastMonth').src = "images/last_enabled.gif";
            image_disable(document.getElementById('imgNextMonth'), false);
            image_disable(document.getElementById('imgLastMonth'), false);
            document.getElementById('imgNextMonth').style.cursor = 'hand';
            document.getElementById('imgLastMonth').style.cursor = 'hand';
        }
		
        combo_disable(frmNav.cboMonth, ((!bNextEnabled) && (!bPrevEnabled)));
        text_disable(frmNav.txtYear, ((!bNextEnabled) && (!bPrevEnabled)));
        button_disable(frmNav.cmdYearUp, frmNav.txtYear.disabled);
        button_disable(frmNav.cmdYearDown, frmNav.txtYear.disabled);
    
        dtReportStartMonth.setDate('01');
        intDaysInMonth = daysInMonth(dtReportEnd.getFullYear(), dtReportEnd.getMonth());
        dtReportEndMonth.setDate(intDaysInMonth.toString());
        if ((dtShownStart.getMonth() == dtSystemMonth.getMonth()) 
            && (dtShownStart.getFullYear() == dtSystemMonth.getFullYear()))
        {
            document.getElementById('imgToday').src = "images/today_disabled.gif";
            image_disable(document.getElementById('imgToday'), true);
            document.getElementById('imgToday').style.cursor = 'default';
        }
        else if ((dtSystemMonth < dtReportEndMonth) && (dtSystemMonth > dtReportStartMonth))
        {
            document.getElementById('imgToday').src = "images/today_enabled.gif";
            image_disable(document.getElementById('imgToday'), false);
            document.getElementById('imgToday').style.cursor = 'hand';
        }
        else
        {
            document.getElementById('imgToday').src = "images/today_disabled.gif";
            image_disable(document.getElementById('imgToday'), true);
            document.getElementById('imgToday').style.cursor = 'default';
        }
		
        frmDate.txtCurrentMonth.value = dtShownStart.getMonth();
        frmDate.txtCurrentYear.value = dtShownStart.getFullYear();
    }

    function daysInMonth(year, month) 
    {
        var dt = new Date(year, month - 1, 32);
        return 32 - dt.getDate();
    }

    function monthChange()
    {
        if (frmUseful.txtLoading.value == 0)
        {
            var blnShowMSG = false;
            var strMessage = "The selected date is outside of the report date boundaries.";
            var dtShownStart = createDate('01/'+frmNav.cboMonth.options[frmNav.cboMonth.selectedIndex].value+'/'+frmNav.txtYear.value);
            var dtShownEnd = createDate(frmDate.txtDaysInMonth.value+'/'+frmNav.cboMonth.options[frmNav.cboMonth.selectedIndex].value+'/'+frmNav.txtYear.value);
            var dtReportStart = createDate(frmDate.txtReportStartDate.value);	
            var dtReportEnd = createDate(frmDate.txtReportEndDate.value);
		
            if (dtShownStart > dtReportEnd)
            {
                blnShowMSG = true;
                OpenHR.messageBox(strMessage,48,"Calendar Reports");
                window.focus();
                //lastMonth();			
                frmNav.cboMonth.selectedIndex = Number(frmDate.txtCurrentMonthIndex.value);
                frmNav.txtYear.value = Number(frmDate.txtCurrentYearValue.value);
                return;
            }
            else if (dtShownEnd < dtReportStart)
            {
                blnShowMSG = true;
                OpenHR.messageBox(strMessage,48,"Calendar Reports");
                window.focus();
                //firstMonth();			
                frmNav.cboMonth.selectedIndex = Number(frmDate.txtCurrentMonthIndex.value);
                frmNav.txtYear.value = Number(frmDate.txtCurrentYearValue.value);
                return;
            }
            else
            {
                frmDate.txtCurrentMonthIndex.value = frmNav.cboMonth.selectedIndex;
                frmDate.txtCurrentYearValue.value = frmNav.txtYear.value;
                dateChange();
            }
        }
    }

    function thisMonth()
    {
        var dtSystemDate = new Date();

        frmUseful.txtLoading.value = 1;
        frmNav.cboMonth.selectedIndex = (dtSystemDate.getMonth());
        frmUseful.txtLoading.value = 0;
        frmNav.txtYear.value = String(dtSystemDate.getFullYear());
        dateChange();
    }
	
    function prevMonth()
    {
        var intListIndex;
        var lngYearValue;
	
        intListIndex = frmNav.cboMonth.selectedIndex;
	
        if (intListIndex == 0)
        {
            lngYearValue = Number(frmNav.txtYear.value);
            frmUseful.txtLoading.value = 1;
            frmNav.cboMonth.selectedIndex = 11;
            frmUseful.txtLoading.value = 0;
            frmNav.txtYear.value = String(lngYearValue - 1);
            dateChange();
        }
        else
        {
            frmNav.cboMonth.selectedIndex = intListIndex - 1;
            monthChange();
        }
    }

    function firstMonth()
    {
        var lngYearValue;
        var lngMonthValue;
        var dtReportStart = new createDate(frmDate.txtReportStartDate.value);

        lngMonthValue = (dtReportStart.getMonth() + 1);
        lngYearValue = dtReportStart.getFullYear();

        frmNav.cboMonth.selectedIndex = lngMonthValue - 1;
        frmNav.txtYear.value = lngYearValue;

        dateChange();
    }

    function lastMonth()
    {
        var lngYearValue;
        var lngMonthValue;
        var dtReportEnd = new createDate(frmDate.txtReportEndDate.value);
	
        lngMonthValue = (dtReportEnd.getMonth() + 1);
        lngYearValue = dtReportEnd.getFullYear();
        frmNav.cboMonth.selectedIndex = lngMonthValue - 1;
        frmNav.txtYear.value = lngYearValue;
	
        dateChange();
    }

    function nextMonth()
    {
        var intListIndex;
        var lngYearValue;
	
        intListIndex = frmNav.cboMonth.selectedIndex;
	
        if (intListIndex == 11)
        {
            lngYearValue = Number(frmNav.txtYear.value);
            frmUseful.txtLoading.value = 1;
            frmNav.cboMonth.selectedIndex = 0;
            frmUseful.txtLoading.value = 0;
            frmNav.txtYear.value = String(lngYearValue + 1);
            dateChange();
        }
        else
        {
            frmNav.cboMonth.selectedIndex = intListIndex + 1;
            monthChange();
        }
    }
	
    function dateChange()
    {
        frmUseful.txtChangingDate.value = 1;
	
        refreshCalendar();
			
        enableDisableNavigation();
	
        frmDate.txtCurrentMonthIndex.value = frmNav.cboMonth.selectedIndex;
        frmDate.txtCurrentYearValue.value = frmNav.txtYear.value;

        frmUseful.txtChangingDate.value = 0;
    }

    function populateCTL_Collections() 
    {
        var frmCalendar = OpenHR.getForm("calendarframe_calendar","frmCalendar");

        if (frmCalendar.txtGroupByDesc.value == 1)
        {
            frmUseful.txtCTLsPopulated.value = 1;
            return;
        }

        //var docCalendar = window.parent.frames("calendarframe_calendar").document;

        var objBaseCTL;
        var vControlName;
        var strSession;
        var dtLabelsDate = new Date();
        var lblTemp;
        var INPUT_VALUE = new String("");
        var intWPCOUNT = new Number(0);
        var intBHolCOUNT = new Number(0);
        var intRegionCOUNT = new Number(0);
        var intBaseID = new Number(0);
	
        for (var i=1; i<=Number(frmCalendar.txtBaseCtlCount.value); i++) 
        {
            vControlName = 'ctlCalRec_' + i;
            objBaseCTL = document.getElementById(vControlName);
            intBaseID = objBaseCTL.BaseDescTag;
		
<%
	if objCalendar.StaticWP = true then 
%>
	    //add the Static WP.
            INPUT_VALUE = new String("");
            vControlName = "";
	    objBaseCTL.StaticWP_Populated = true;
	    objBaseCTL.HistoricWP_Populated = false;
	    vControlName = 'txtWPCOUNT_' + intBaseID;
	    try
	    {
	        intWPCOUNT = Number(document.getElementById(vControlName).value);
	    }
	    catch(e)
	    {
	        intWPCOUNT = 0;
	    }
	    for (var iElement=1; iElement<=intWPCOUNT; iElement++)
	    {
	        vControlName = 'txtWP_' + intBaseID;
	        INPUT_VALUE = document.getElementById(vControlName).value;
	        objBaseCTL.AddWorkingPattern(INPUT_VALUE, true);
	    }
<% 
	else 
%>

	    //add all the historic WPs.
            INPUT_VALUE = new String("");
            vControlName = "";
	    objBaseCTL.StaticWP_Populated = false
	    objBaseCTL.HistoricWP_Populated = true;
	    vControlName = 'txtWPCOUNT_' + intBaseID;
	    try
	    {
	        intWPCOUNT = Number(document.getElementById(vControlName).value);
	    }
	    catch(e)
	    {
	        intWPCOUNT = 0;
	    }
	    for (var iElement=1; iElement<=intWPCOUNT; iElement++)
	    {
	        vControlName = 'txtWP_' + intBaseID + '_' + iElement;
	        INPUT_VALUE = document.getElementById(vControlName).value;
	        objBaseCTL.AddWorkingPattern(INPUT_VALUE, false);
	    }
<% 
	end if 

 if objCalendar.StaticReg = true then 
%>
	    //add the Static BHol. 
            INPUT_VALUE = new String("");
            vControlName = "";
	    objBaseCTL.StaticRegion_Populated = true;
	    objBaseCTL.HistoricRegion_Populated = false;
	    vControlName = 'txtBHolCOUNT_' + intBaseID;
	    try 
	    {
	        intBHolCOUNT = Number(document.getElementById(vControlName).value);
	    }
	    catch(e)	
	    {
	        intBHolCOUNT = 0;
	    }

	    for (var iElement=1; iElement<=intBHolCOUNT; iElement++)
	    {
	        vControlName = 'txtBHol_' + intBaseID + '_' + iElement;
	        INPUT_VALUE = document.getElementById(vControlName).value;
	        objBaseCTL.AddBankHoliday(INPUT_VALUE, true);
	    }
<% 
	else 
%>
	    //add all the historic BHols. 
            INPUT_VALUE = new String("");
            vControlName = "";
	    objBaseCTL.StaticRegion_Populated = false;
	    objBaseCTL.HistoricRegion_Populated = true;
	    vControlName = 'txtBHolCOUNT_' + intBaseID;
	    try 
	    {
	        intBHolCOUNT = Number(document.getElementById(vControlName).value);
	    }
	    catch(e)	
	    {
	        intBHolCOUNT = 0;
	    }

	    for (var iElement=1; iElement<=intBHolCOUNT; iElement++)
	    {
	        vControlName = 'txtBHol_' + intBaseID + '_' + iElement;
	        INPUT_VALUE = document.getElementById(vControlName).value;
	        objBaseCTL.AddBankHoliday(INPUT_VALUE, false);
	    }
			
<% 
	end if 
%>

	    //add all the historic Career Changes. 
	    vControlName = 'txtRegionCOUNT_' + intBaseID;
	    try
	    {
	        intRegionCOUNT = Number(document.getElementById(vControlName).value);
	    }
	    catch(e)
	    {
	        intRegionCOUNT = 0;
	    }

	    for (var iElement=1; iElement<=intRegionCOUNT; iElement++)
	    {
	        vControlName = 'txtRegion_' + intBaseID + '_' + iElement;
	        INPUT_VALUE = document.getElementById(vControlName).value;

	        objBaseCTL.AddCareerChange(INPUT_VALUE);
	    }
	} //for (var i=1; i<=frmCalendar.txtBaseCtlCount.value; i++) 

    frmUseful.txtCTLsPopulated.value = 1;
    return true;	
}
	
function refreshCalendar() {

    if (frmUseful.txtCTLsPopulated.value != 1)
    {
        populateCTL_Collections();
    }

    var frmCalendar = OpenHR.getForm("calendarframe_calendar","frmCalendar");
    //var docCalendar = window.parent.frames("calendarframe_calendar").document;
    var tempMonth = frmNav.cboMonth.options[frmNav.cboMonth.selectedIndex].value;
    var tempYear = frmNav.txtYear.value;
    var intSessionCount = new Number(0);
    var intControlCount = new Number(0);
    var intDateCount = new Number(0);
    var objBaseCTL;
    var vControlName;
    var strSession;
    var dtLabelsDate = new Date();
    var lblTemp;
    var strReportStart = new String(frmDate.txtReportStartDate.value);
    var strReportEnd = new String(frmDate.txtReportEndDate.value);
    var strClientDateFormat = new String(frmDate.txtClientDateFormat.value);
    var strClientDateSeparator = new String(frmDate.txtClientDateSeparator.value);
	
    frmNav.ctlDates.ClientDateFormat = strClientDateFormat;
    frmNav.ctlDates.ReportStartDate = strReportStart;
    frmNav.ctlDates.ReportEndDate = strReportEnd;
	
    frmNav.ctlDates.SetDate(tempMonth,tempYear);
    frmDate.txtDaysInMonth.value = frmNav.ctlDates.CurrentDaysInMonth;
	
    for (var i=1; i<=Number(frmCalendar.txtBaseCtlCount.value); i++) 
    {
        vControlName = 'ctlCalRec_' + i;
        objBaseCTL = document.getElementById(vControlName);
		
        objBaseCTL.ClientDateFormat = strClientDateFormat;
        objBaseCTL.ClientDateSeparator = strClientDateSeparator;
		
        objBaseCTL.HideSeparators();
		
        objBaseCTL.ReportStartDate = strReportStart;
        objBaseCTL.ReportEndDate = strReportEnd;
		
        objBaseCTL.RefreshCalendar(tempMonth,tempYear); 
    } //for (var i=1; i<=frmCalendar.txtBaseCtlCount.value; i++) 
	
    var frmGetDataForm = OpenHR.getForm("dataframe", "frmCalendarGetData");
			
    frmGetDataForm.txtMode.value = "LOADCALENDARREPORTDATA";
    frmGetDataForm.txtDaysInMonth.value = frmDate.txtDaysInMonth.value;
    frmGetDataForm.txtMonth.value = (frmNav.cboMonth.options[frmNav.cboMonth.selectedIndex].value);
    frmGetDataForm.txtYear.value = frmNav.txtYear.value;

    refreshData();
		
    return true;
}

function fillCalBoxes()
{
   
    var frmKey =  OpenHR.getForm("calendarframe_key","frmKey");

    //var docCalendar = window.parent.frames("calendarframe_calendar").document;
    var frmCalendarData =OpenHR.getForm("dataframe","frmCalendarData");
    //var docData = window.parent.frames("dataframe").document;
    var eventCollection = frmCalendarData.elements;
    var objBaseCTL;
    var vControlName;
    var intBaseRecordIndex;
    var INPUT_STRING = new String("");
    var EVENT_DETAIL = new String("");
    var lngColour;
    var strEventID;
    var elementName = new String("");
    var sDetailControl;
    var iEventNumber;
	
    for (var i=0; i<frmCalendarData.elements.length; i++)
    {
        elementName = frmCalendarData.item(i).name;

        if (elementName.substring(0,6) == "Event_")
        {
            INPUT_STRING = frmCalendarData.item(i).value;
            strEventID = INPUT_STRING.substring(INPUT_STRING.indexOf("***")+3,INPUT_STRING.indexOf("***",INPUT_STRING.indexOf("***")+3))
            lngColour = Number(frmKey.ctlKey.GetKeyColour(strEventID));
            intBaseRecordIndex = Number(INPUT_STRING.substring(0,INPUT_STRING.indexOf("***")));

            sDetailControl = "EventDetail_" + elementName.substring(elementName.lastIndexOf("_")+1,elementName.length);
            EVENT_DETAIL = document.getElementById(sDetailControl).value;
			
            vControlName = 'ctlCalRec_' + String(intBaseRecordIndex);
            objBaseCTL = document.getElementById(vControlName);
            objBaseCTL.FillEventCalBoxes(INPUT_STRING,lngColour,EVENT_DETAIL);
        }
    }
	
    refreshDateSpecifics();

    refreshDisplay();	

    return true;
}
	
function refreshDateSpecifics()
{
    var frmCalendar = OpenHR.getForm("calendarframe_calendar","frmCalendar");
    //var docCalendar = window.parent.frames("calendarframe_calendar").document;
    var frmOptions = OpenHR.getForm("calendarframe_options","frmOptions");
    var objBaseCTL;
    var bIncBHols, bIncWorkDaysOnly, bShowBHols, bShowCaptions, bShowWeekends;
    var frmKey = OpenHR.getForm("calendarframe_key","frmKey");
	
    bIncBHols = (frmOptions.chkIncludeBHols.checked == true);
    bIncWorkDaysOnly = (frmOptions.chkIncludeWorkingDaysOnly.checked == true);
    bShowBHols = (frmOptions.chkShadeBHols.checked == true);
    bShowCaptions = (frmOptions.chkCaptions.checked == true);
    bShowWeekends = (frmOptions.chkShadeWeekends.checked == true);

    frmKey.ctlKey.CaptionsVisible = (bShowCaptions);
	
    for (var i=1; i<=Number(frmCalendar.txtBaseCtlCount.value); i++) 
    {
        vControlName = 'ctlCalRec_' + i;
        objBaseCTL = document.getElementById(vControlName);
		
        //require for the event separator
        objBaseCTL.DateChanged = true;
        objBaseCTL.IncludeBankHolidays = bIncBHols;
        objBaseCTL.IncludeWorkingDaysOnly = bIncWorkDaysOnly;
        objBaseCTL.ShowBankHolidays = bShowBHols;
        objBaseCTL.ShowCaptions = bShowCaptions;
        objBaseCTL.ShowWeekends = bShowWeekends;
		
        objBaseCTL.RefreshDateSpecifics();
    }
    return true;
}

function createDate(psDateString)
{
    var dtDate = new Date();
    var strDate = new String(psDateString);
    var lngDate, lngMonth, lngYear;
    var charIndex = new Number(0);
	
    //Eg. 23/08/2003
    //		0123456789   (substr index)
	
    charIndex = strDate.indexOf("/",charIndex);
    lngDate = Number(strDate.substring(0,charIndex));

    lngMonth = Number(strDate.substring(charIndex+1,strDate.indexOf("/",charIndex+1)));
    lngMonth = lngMonth - 1;
	
    charIndex = strDate.indexOf("/",charIndex+1);
    lngYear = Number(strDate.substring(charIndex+1,(strDate.length)));
	
    dtDate.setFullYear(lngYear,lngMonth,lngDate);
	
    return dtDate;
}

function setRecordsNumeric()
{
    var sConvertedValue;
    var sDecimalSeparator;
    var sThousandSeparator;
    var sPoint;
		
    sDecimalSeparator = "\\";

    sDecimalSeparator = sDecimalSeparator.concat(OpenHR.LocaleDecimalSeparator);
    var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");

    sThousandSeparator = "\\";
    sThousandSeparator = sThousandSeparator.concat(OpenHR.LocaleThousandSeparator);
    var reThousandSeparator = new RegExp(sThousandSeparator, "gi");
		
    sPoint = "\\.";
    var rePoint = new RegExp(sPoint, "gi");
	
    if (frmNav.txtYear.value == '') 
    {
        frmNav.txtYear.value = 0;
    }
		
    // Convert the value from locale to UK settings for use with the isNaN funtion.
    sConvertedValue = new String(frmNav.txtYear.value);
	
    // Remove any thousand separators.
    sConvertedValue = sConvertedValue.replace(reThousandSeparator, "");
    frmNav.txtYear.value = sConvertedValue;

    // Convert any decimal separators to '.'.
    if (OpenHR.LocaleDecimalSeparator != ".") 
    {
        // Remove decimal points.
        sConvertedValue = sConvertedValue.replace(rePoint, "A");
        // replace the locale decimal marker with the decimal point.
        sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
    }
	
    if(isNaN(sConvertedValue) == true) 
    {
        OpenHR.messageBox("Invalid numeric value.",48,"Calendar Reports");
        window.focus();
        frmNav.txtYear.value = 1900;
    }
    else 
    {
        if (sConvertedValue.indexOf(".") >= 0 ) 
        {
            OpenHR.messageBox("Invalid integer value.",48,"Calendar Reports");
            window.focus();
            frmNav.txtYear.value = 1900;
        }
        else 
        {
            if (frmNav.txtYear.value < 0) 
            {
                OpenHR.messageBox("The year value cannot be negative.",48,"Calendar Reports");
                window.focus();
                frmNav.txtYear.value = 1900;
            }
            else 
            { 
                if (frmNav.txtYear.value < 1900) 
                {
                    OpenHR.messageBox("The year value must be between 1900 and 3000.",48,"Calendar Reports");
                    window.focus();
                    frmNav.txtYear.value = 1900;
                }
                else
                {
                    if (frmNav.txtYear.value > 3000) 
                    {
                        OpenHR.messageBox("The year value must be between 1900 and 3000.",48,"Calendar Reports");
                        window.focus();
                        frmNav.txtYear.value = 3000;
                    }
                }
            }
        }
    }
		
    monthChange();
}
		
function spinRecords(pfUp)
{
    var iRecords = frmNav.txtYear.value; 
    if (pfUp == true) 
    {
        iRecords = ++iRecords;
    }
    else 
    {
        if (iRecords > 0) 
        {
            iRecords = iRecords - 1;
        }
    }
		
    frmNav.txtYear.value = iRecords;
}

function styleArgument(psDefnString, psParameter)
{
    var iCharIndex;
    var sDefn;
	
    sDefn = new String(psDefnString);
    psParameter = psParameter.toUpperCase(); 
	
    iCharIndex = sDefn.indexOf("	");
    if (iCharIndex >= 0) 
    {
        if (psParameter == "TYPE") return sDefn.substr(0, iCharIndex);
        sDefn = sDefn.substr(iCharIndex + 1);
        iCharIndex = sDefn.indexOf("	");
        if (iCharIndex >= 0) 
        {
            if (psParameter == "STARTCOL") return sDefn.substr(0, iCharIndex);
            sDefn = sDefn.substr(iCharIndex + 1);
            iCharIndex = sDefn.indexOf("	");
            if (iCharIndex >= 0) 
            {
                if (psParameter == "STARTROW") return sDefn.substr(0, iCharIndex);
                sDefn = sDefn.substr(iCharIndex + 1);
                iCharIndex = sDefn.indexOf("	");
                if (iCharIndex >= 0) 
                {
                    if (psParameter == "ENDCOL") return sDefn.substr(0, iCharIndex);
                    sDefn = sDefn.substr(iCharIndex + 1);
                    iCharIndex = sDefn.indexOf("	");
                    if (iCharIndex >= 0) 
                    {
                        if (psParameter == "ENDROW") return sDefn.substr(0, iCharIndex);
                        sDefn = sDefn.substr(iCharIndex + 1);
                        iCharIndex = sDefn.indexOf("	");
                        if (iCharIndex >= 0) 
                        {
                            if (psParameter == "BACKCOLOR") return sDefn.substr(0, iCharIndex);
                            sDefn = sDefn.substr(iCharIndex + 1);
                            iCharIndex = sDefn.indexOf("	");
                            if (iCharIndex >= 0) 
                            {
                                if (psParameter == "FORECOLOR") return sDefn.substr(0, iCharIndex);
                                sDefn = sDefn.substr(iCharIndex + 1);
                                iCharIndex = sDefn.indexOf("	");
                                if (iCharIndex >= 0) 
                                {
                                    if (psParameter == "BOLD") return sDefn.substr(0, iCharIndex);
                                    sDefn = sDefn.substr(iCharIndex + 1);
                                    iCharIndex = sDefn.indexOf("	");
                                    if (iCharIndex >= 0) 
                                    {
                                        if (psParameter == "UNDERLINE") return sDefn.substr(0, iCharIndex);
                                        sDefn = sDefn.substr(iCharIndex + 1);

                                        if (psParameter == "GRIDLINES") return sDefn;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
	
    return "";
}

function mergeArgument(psDefnString, psParameter)
{
    var iCharIndex;
    var sDefn;
	
    sDefn = new String(psDefnString);
    psParameter = psParameter.toUpperCase(); 
	
    iCharIndex = sDefn.indexOf("	");
    if (iCharIndex >= 0) 
    {
        if (psParameter == "STARTCOL") return sDefn.substr(0, iCharIndex);
        sDefn = sDefn.substr(iCharIndex + 1);
        iCharIndex = sDefn.indexOf("	");
        if (iCharIndex >= 0) 
        {
            if (psParameter == "STARTROW") return sDefn.substr(0, iCharIndex);
            sDefn = sDefn.substr(iCharIndex + 1);
            iCharIndex = sDefn.indexOf("	");
            if (iCharIndex >= 0) 
            {
                if (psParameter == "ENDCOL") return sDefn.substr(0, iCharIndex);
                sDefn = sDefn.substr(iCharIndex + 1);

                if (psParameter == "ENDROW") return sDefn;
            }
        }
    }
	
    return "";	
}
	
function replace(sExpression, sFind, sReplace)
{
    //gi (global search, ignore case)
    var re = new RegExp(sFind,"gi");
    sExpression = sExpression.replace(re, sReplace);
    return(sExpression);
}

function writeText(sText)
{
    document.write(sText);
}

</script>

		<form name="frmNav" id="frmNav">
			<table align="center" class="invisible" cellPadding="0" cellSpacing="0" width="100%" height="100%">
				<tr>
					<td>
						<table align="center" class="invisible" cellPadding="0" cellSpacing="0" width="100%" height="100%">
							<tr height="47">
								<td width="*"  rowspan="1">
									&nbsp;
								</td>
								<td align="right">
									<table class="outline" cellspacing="0" cellpadding="2" height="5">
										<tr height="5" valign="middle">
											<td width="5"></td> 
											<td height="5" align="left" valign="middle" width="5">
												<img ALT="First Month" align="center" valign="middle" border="0" src="images/first_disabled.gif" name="imgFirstMonth" id="imgFirstMonth" WIDTH="16" HEIGHT="16"
												    onClick="firstMonth();" >
											</td>
											<td width="5"></td> 
											<td height="5" align="left" valign="middle" width="5">
												<img ALT="Previous Month" align="center" valign="middle" border="0" src="images/previous_disabled.gif" name="imgPrevMonth" id="imgPrevMonth" WIDTH="16" HEIGHT="16"
												    onClick="prevMonth();" >
											</td>
											<td width="5"></td> 
											<td height="5" align="left" valign="top" width="100">
<%
    Response.Write(objCalendar.HTML_MonthCombo(0))
%>	
												<!--						<select name="cboMonth" id="cboMonth" style="WIDTH: 100px" onChange="monthChange();">							<option value="1" selected>January							<option value="2">February							<option value="3">March							<option value="4">April							<option value="5">May							<option value="6">June							<option value="7">July							<option value="8">August							<option value="9">September							<option value="10">October							<option value="11">November							<option value="12">December						</select>						-->
											</td>
											<td width="5"></td> 
											<td height="5" align="left" valign="top">
												<table WIDTH="100%" class="invisible" CELLSPACING="0" CELLPADDING="0">
													<tr>
														<td width="40">
															<input maxlength="4" value="2003" id="txtYear" name="txtYear" class="text" style="WIDTH: 40px" width="40" value="0" 
															    onkeypress="if(window.event.keyCode==13) {frmNav.txtYear.blur(); return false;}" 
															    onblur="setRecordsNumeric();" 
															    onchange="setRecordsNumeric();">
														</td>
														<td width="15" align="center">
															<input style="WIDTH: 15px; Font-Bold: true" type="button" value="+" id="cmdYearUp" name="cmdYearUp" class="btn"
															    onclick="spinRecords(true);setRecordsNumeric();"
			                                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
			                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
			                                                    onfocus="try{button_onFocus(this);}catch(e){}"
			                                                    onblur="try{button_onBlur(this);}catch(e){}" />
														</td>
														<td width="15" align="center">
															<input style="WIDTH: 15px; Font-Bold: true" type="button" value="-" id="cmdYearDown" name="cmdYearDown" class="btn"
															    onclick="spinRecords(false);setRecordsNumeric();"
			                                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
			                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
			                                                    onfocus="try{button_onFocus(this);}catch(e){}"
			                                                    onblur="try{button_onBlur(this);}catch(e){}" />
														</td>
													</tr>
												</table>
											</td>
											<td width="5"></td> 
											<td height="5" align="left" valign="middle" width="5">
												<img ALT="Current Month" align="center" valign="middle" style="margin-top:1px;" src="images/today_disabled.gif" name="imgToday" id="imgToday" WIDTH="16" HEIGHT="16"
												    onClick="thisMonth();" >
											</td>
											<td width="5"></td> 
											<td height="5" align="left" valign="middle" width="5">
												<img ALT="Next Month" align="center" valign="middle" border="0" src="images/next_disabled.gif" name="imgNextMonth" id="imgNextMonth" WIDTH="16" HEIGHT="16"
												    onClick="nextMonth();" >
											</td>
											<td width="5"></td> 
											<td height="5" align="left" valign="middle" width="5">
												<img ALT="Last Month" align="center" valign="middle" border="0" src="images/last_disabled.gif" name="imgLastMonth" id="imgLastMonth" WIDTH="16" HEIGHT="16"
												    onClick="lastMonth();" >
											</td>
											<td width="5"></td> 
										</tr>
									</table>
								</td>
								<td width="62" rowspan="1">
									&nbsp;
								</td>
							</tr>
						</table>
					</td>
				</tr>
				
				<tr>
					<td>
						<table align="center" class="invisible" cellPadding="0" cellSpacing="0" width="100%">
							<tr height="30">
								<td align="right" nowrap width="100%" colspan="2">
									<table class="invisible" cellspacing="0" cellpadding="0" width="100%" height="100%">
										<tr>
										<td width="100%">
											<object CLASSID="CLSID:41021C13-8D42-4364-8388-9506F0755AE3" 
															CODEBASE="cabs/COAInt_CalRepDates.cab#version=1,0,0,2" 
															id="ctlDates" name="ctlDates" style="VISIBILITY: visible; WIDTH: 100%" 
															width="100%" VIEWASTEXT>
    											<param NAME="BackColor" VALUE="16513017">
											</object>
										</td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</form>

		<form id="frmUseful" name="frmUseful" style="visibility:hidden;display:none">
			<input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
			<input type="hidden" id="txtLoading" name="txtLoading" value="1">
			<input type="hidden" id="txtChangingDate" name="txtChangingDate" value="0">
			<input type="hidden" id="txtCurrentBaseTableID" name="txtCurrentBaseTableID">
			<input type="hidden" id="txtAvailableColumnsLoaded" name="txtAvailableColumnsLoaded" value="0">
			<input type="hidden" id="txtEventsLoaded" name="txtEventsLoaded" value="0">
			<input type="hidden" id="txtSortLoaded" name="txtSortLoaded" value="0">
			<input type="hidden" id="txtChanged" name="txtChanged" value="0">
			<input type="hidden" id="txtUtilID" name="txtUtilID" value="<%=session("utilid")%>">
			<input type="hidden" id="txtUtilType" name="txtUtilType" value="<%=session("utiltype")%>">
			<input type="hidden" id="txtEventCount" name="txtEventCount" value="<%=session("eventcount")%>">
			<input type="hidden" id="txtHiddenEventFilterCount" name="txtHiddenEventFilterCount" value="<%=session("hiddenfiltercount")%>">
			<input type="hidden" id="txtLockGridEvents" name="txtLockGridEvents" value="0">
			<input type="hidden" id="txtCTLsPopulated" name="txtCTLsPopulated" value="0">
<%
    Dim cmdDefinition As Object
    Dim prmModuleKey As Object
    Dim prmParameterKey As Object
    Dim prmParameterValue As Object
        
    cmdDefinition = CreateObject("ADODB.Command")
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

    Err.Clear()
	cmdDefinition.Execute

    Response.Write("<INPUT type='hidden' id=txtPersonnelTableID name=txtPersonnelTableID value=" & cmdDefinition.Parameters("paramValue").value & ">" & vbCrLf)
	
    cmdDefinition = Nothing

    Dim sErrorDescription As String
    
    Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
    Response.Write("<INPUT type='hidden' id=txtMode name=txtMode value=" & Session("action") & ">" & vbCrLf)
%>
		</form>

		<form id="frmDate" name="frmDate" style="visibility:hidden;display:none">
			<input type="hidden" id="txtFirstDayOfMonth" name="txtFirstDayOfMonth">
			<input type="hidden" id="txtDaysInMonth" name="txtDaysInMonth">
			<input type="hidden" id="txtDayControlCount" name="txtDayControlCount" value="37">
<%
    Response.Write("	<INPUT type=""hidden"" id=txtReportStartDate name=txtReportStartDate value=""" & objCalendar.ReportStartDate_CalendarString & """>" & vbCrLf)
    Response.Write("	<INPUT type=""hidden"" id=txtReportEndDate name=txtReportEndDate value=""" & objCalendar.ReportEndDate_CalendarString & """>" & vbCrLf)
%>
			<INPUT type="hidden" id=txtStartOnCurrentMonth name=txtStartOnCurrentMonth value="
<% 
	if objCalendar.StartOnCurrentMonth then 
%> 
				1
<% 
	else 
%>
				0
<% 
	end if 
%>
			">
			
			<input type="hidden" id="txtClientDateFormat" name="txtClientDateFormat" value="<%=Session("LocaleDateFormat")%>">
			<input type="hidden" id="txtClientDateSeparator" name="txtClientDateSeparator" value="<%=session("LocaleDateSeparator")%>">
			<input type="hidden" id="txtCurrentMonth" name="txtCurrentMonth">
			<input type="hidden" id="txtCurrentYear" name="txtCurrentYear">
			<input type="hidden" id="txtCurrentMonthIndex" name="txtCurrentMonthIndex">
			<input type="hidden" id="txtCurrentYearValue" name="txtCurrentYearValue">
		</form>

		<form id="frmBankHolidays" name="frmBankHolidays" style="visibility:hidden;display:none">
			<input type="hidden" id="Hidden1" name="txtFirstDayOfMonth">
		</form>
<%
	Dim mblnOutputPreview 
	Dim mlngOutputFormat 
	Dim mblnOutputScreen 
	Dim mblnOutputPrinter 
	Dim mstrOutputPrinterName 
	Dim mblnOutputSave 
	Dim mlngOutputSaveExisting
	Dim mblnOutputEmail
	Dim mlngOutputEmailID 
	Dim mstrOutputEmailName 
	Dim mstrOutputEmailSubject 
	Dim mstrOutputEmailAttachAs 
	Dim mstrOutputFilename

	mblnOutputPreview = objCalendar.OutputPreview
	mlngOutputFormat = objCalendar.OutputFormat 
	mblnOutputScreen = objCalendar.OutputScreen
	mblnOutputPrinter = objCalendar.OutputPrinter
	mstrOutputPrinterName = objCalendar.OutputPrinterName
	mblnOutputSave = objCalendar.OutputSave
	mlngOutputSaveExisting = objCalendar.OutputSaveExisting
	mblnOutputEmail = objCalendar.OutputEmail
	mlngOutputEmailID = objCalendar.OutputEmailID
	mstrOutputEmailName = objCalendar.OutputEmailGroupName
	mstrOutputEmailSubject = objCalendar.OutputEmailSubject
	mstrOutputEmailAttachAs = objCalendar.OutputEmailAttachAs
	mstrOutputFilename = objCalendar.OutputFileName
%>

		<form target="Output" action="util_run_outputoptions" method="post" id="frmExportData" name="frmExportData">
			<input type="hidden" id="txtPreview" name="txtPreview" value="<%=mblnOutputPreview%>">
			<input type="hidden" id="txtFormat" name="txtFormat" value="<%=mlngOutputFormat%>">
			<input type="hidden" id="txtScreen" name="txtScreen" value="<%=mblnOutputScreen%>">
			<input type="hidden" id="txtPrinter" name="txtPrinter" value="<%=mblnOutputPrinter%>">
			<input type="hidden" id="txtPrinterName" name="txtPrinterName" value="<%=mstrOutputPrinterName%>">
			<input type="hidden" id="txtSave" name="txtSave" value="<%=mblnOutputSave%>">
			<input type="hidden" id="txtSaveExisting" name="txtSaveExisting" value="<%=mlngOutputSaveExisting%>">
			<input type="hidden" id="txtEmail" name="txtEmail" value="<%=mblnOutputEmail%>">
			<input type="hidden" id="txtEmailAddr" name="txtEmailAddr" value="<%=mlngOutputEmailID%>">
			<input type="hidden" id="txtEmailAddrName" name="txtEmailAddrName" value="<%=replace(mstrOutputEmailName, """", "&quot;")%>">
			<input type="hidden" id="txtEmailSubject" name="txtEmailSubject" value="<%=replace(mstrOutputEmailSubject, """", "&quot;")%>">
			<input type="hidden" id="txtEmailAttachAs" name="txtEmailAttachAs" value="<%=replace(mstrOutputEmailAttachAs, """", "&quot;")%>">
			<input type="hidden" id="txtFileName" name="txtFileName" value="<%=mstrOutputFilename%>">
			<input type="hidden" id="Hidden2" name="txtUtilType" value="<%=session("utilType")%>">
		</form>
<%
    Dim objUser As HR.Intranet.Server.clsSettings
    
	'**************************************************************
	'Output forms and the respective elements for Region/BHol/WPs
	
    Response.Write(objCalendar.Write_Static_Historic_Forms)
	
    '**************************************************************

    'Write the function that Outputs the report to the Output Classes in the Client DLL.
		
    Response.Write("<script type=""text/javascript"">" & vbCrLf)  
    Response.Write("function outputReport() " & vbCrLf)
    Response.Write("	{" & vbCrLf & vbCrLf)
  
    Response.Write("	var frmOutput = openHR.getForm(""dataframe"",""frmCalendarData"");" & vbCrLf)

    Response.Write("	var lngPageColumnCount = 3;" & vbCrLf)
    Response.Write("    var lngActualRow = new Number(0);" & vbCrLf)
    Response.Write("    var blnSettingsDone = false;" & vbCrLf)
    Response.Write("	var sColHeading = new String(''); " & vbCrLf)
    Response.Write("	var iColDataType = new Number(12); " & vbCrLf)
    Response.Write("	var iColDecimals = new Number(0); " & vbCrLf)
    Response.Write("    var blnNewPage = false;" & vbCrLf)
    Response.Write("    var lngPageCount = new Number(0);" & vbCrLf)

    Response.Write("  var strType = new String('');" & vbCrLf)
    Response.Write("  var lngStartCol = new Number(0);" & vbCrLf)
    Response.Write("  var lngStartRow = new Number(0);" & vbCrLf)
    Response.Write("  var lngEndCol = new Number(0);" & vbCrLf)
    Response.Write("  var lngEndRow = new Number(0);" & vbCrLf)
    Response.Write("  var lngBackCol = new Number(0);" & vbCrLf)
    Response.Write("  var lngForeCol = new Number(0);" & vbCrLf)
    Response.Write("  var blnBold = false;" & vbCrLf)
    Response.Write("  var blnUnderline = false;" & vbCrLf)
    Response.Write("  var blnGridlines = false;" & vbCrLf)
	
    objUser = New HR.Intranet.Server.clsSettings
    
    Response.Write("  window.parent.parent.ASRIntranetOutput.UserName = """ & cleanStringForJavaScript(Session("Username")) & """;" & vbCrLf)
    Response.Write("  window.parent.parent.ASRIntranetOutput.SaveAsValues = """ & cleanStringForJavaScript(Session("OfficeSaveAsValues")) & """;" & vbCrLf)
	
    Response.Write("  frmMenuFrame = window.parent.parent.opener.window.parent.frames(""menuframe"");" & vbCrLf)

    Response.Write("	window.parent.parent.ASRIntranetOutput.SettingOptions(")
    Response.Write("""" & cleanStringForJavaScript(objUser.GetUserSetting("Output", "WordTemplate", "")) & """, ")
    Response.Write("""" & cleanStringForJavaScript(objUser.GetUserSetting("Output", "ExcelTemplate", "")) & """, ")

    If (objUser.GetUserSetting("Output", "ExcelGridlines", "0") = "1") Then
        Response.Write("true, ")
    Else
        Response.Write("false, ")
    End If

    If (objUser.GetUserSetting("Output", "ExcelHeaders", "0") = "1") Then
        Response.Write("true, ")
    Else
        Response.Write("false, ")
    End If

    If (objUser.GetUserSetting("Output", "AutoFitCols", "1") = "1") Then
        Response.Write("true, ")
    Else
        Response.Write("false, ")
    End If

    If (objUser.GetUserSetting("Output", "Landscape", "1") = "1") Then
        Response.Write("true, " & vbCrLf)
    Else
        Response.Write("false, " & vbCrLf)
    End If
			   
    Response.Write("frmMenuFrame.document.all.item(""txtSysPerm_EMAILGROUPS_VIEW"").value);" & vbCrLf)

    Response.Write("  window.parent.parent.ASRIntranetOutput.SettingLocations(")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleCol", "3")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleRow", "2")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataCol", "2")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataRow", "4")) & ");" & vbCrLf)

    Response.Write("  window.parent.parent.ASRIntranetOutput.SettingTitle(")
    If (objUser.GetUserSetting("Output", "TitleGridLines", "0") = "1") Then
        Response.Write("true, ")
    Else
        Response.Write("false, ")
    End If

    If (objUser.GetUserSetting("Output", "TitleBold", "1") = "1") Then
        Response.Write("true, ")
    Else
        Response.Write("false, ")
    End If

    If (objUser.GetUserSetting("Output", "TitleUnderline", "0") = "1") Then
        Response.Write("true, ")
    Else
        Response.Write("false, ")
    End If

    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleBackcolour", "16777215")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleForecolour", "6697779")) & ", ")
    '    Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "TitleBackcolour", "16777215"))) & ", ")
    '   Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "TitleForecolour", "6697779"))) & ");" & vbCrLf)
    Response.Write("1,1);" & vbCrLf)
    
    Response.Write("window.parent.parent.ASRIntranetOutput.SettingHeading(")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingGridLines", "1")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingBold", "1")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingUnderline", "0")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingBackcolour", "16248553")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingForecolour", "6697779")) & ", ")
    '    Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "HeadingBackcolour", "16248553"))) & ", ")
    '   Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "HeadingForecolour", "6697779"))) & ");" & vbCrLf)
    Response.Write("1,1);" & vbCrLf)

    Response.Write("window.parent.parent.ASRIntranetOutput.SettingData(")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataGridLines", "1")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataBold", "0")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataUnderline", "0")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataBackcolour", "15988214")) & ", ")
    Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataForecolour", "6697779")) & ", ")
    '    Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "DataBackcolour", "15988214"))) & ", ")
    '   Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "DataForecolour", "6697779"))) & ");" & vbCrLf)
    Response.Write("1,1);" & vbCrLf)

	if Session("EmailGroupID") = "" then
		Session("EmailGroupID") = 0
	end if
	
    Response.Write("window.parent.parent.ASRIntranetOutput.SetOptions(false, " & _
                                            "frmExportData.txtFormat.value, frmExportData.txtScreen.value, " & _
                                            "frmExportData.txtPrinter.value, frmExportData.txtPrinterName.value, " & _
                                            "frmExportData.txtSave.value, frmExportData.txtSaveExisting.value, " & _
                                            "frmExportData.txtEmail.value, frmOutput.txtEmailGroupAddr.value, " & _
                                            "frmExportData.txtEmailSubject.value, frmExportData.txtEmailAttachAs.value, frmExportData.txtFileName.value);" & vbCrLf)

    Response.Write("  if (frmExportData.txtFormat.value == ""0"") {" & vbCrLf)
    Response.Write("    if (frmExportData.txtPrinter.value == ""true"") {" & vbCrLf)
    Response.Write("			window.parent.parent.ASRIntranetOutput.SetPrinter();" & vbCrLf)
    Response.Write("      dataOnlyPrint();" & vbCrLf)
    Response.Write("			window.parent.parent.ASRIntranetOutput.ResetDefaultPrinter();" & vbCrLf)
    Response.Write("    }" & vbCrLf)
    Response.Write("  }" & vbCrLf)
    Response.Write("  else {" & vbCrLf)

    Response.Write("if (window.parent.parent.ASRIntranetOutput.GetFile() == true) " & vbCrLf)
    Response.Write("	{" & vbCrLf)
    Response.Write("	window.parent.parent.ASRIntranetOutput.InitialiseStyles();" & vbCrLf)
    Response.Write("	window.parent.parent.ASRIntranetOutput.ResetStyles();" & vbCrLf)
    Response.Write("	window.parent.parent.ASRIntranetOutput.ResetColumns();" & vbCrLf)
    Response.Write("	window.parent.parent.ASRIntranetOutput.ResetMerges();" & vbCrLf)

    Response.Write("	window.parent.parent.ASRIntranetOutput.HeaderRows = 1;" & vbCrLf)
    Response.Write("	window.parent.parent.ASRIntranetOutput.HeaderCols = 0;" & vbCrLf)
    Response.Write("	window.parent.parent.ASRIntranetOutput.SizeColumnsIndependently = true;" & vbCrLf)
	
    Response.Write("	window.parent.parent.ASRIntranetOutput.ArrayDim((lngPageColumnCount-1), 0);" & vbCrLf & vbCrLf)
    Response.Write("  frmOutput.grdCalendarOutput.focus();")
	
    Response.Write("  frmOutput.grdCalendarOutput.MoveFirst();" & vbCrLf)
    Response.Write("  for (var lngRow=0; lngRow<frmOutput.grdCalendarOutput.Rows; lngRow++)" & vbCrLf)
    Response.Write("		{" & vbCrLf)
    Response.Write("		bm = frmOutput.grdCalendarOutput.AddItemBookmark(lngRow);" & vbCrLf)
	
    Response.Write("		if (lngRow == (frmOutput.grdCalendarOutput.Rows - 1))" & vbCrLf)
    Response.Write("			{" & vbCrLf)
    Response.Write("			sBreakValue = frmOutput.grdCalendarOutput.Columns(1).CellText(bm);" & vbCrLf)
    Response.Write("			window.parent.parent.ASRIntranetOutput.AddPage(replace(frmOutput.grdCalendarOutput.Caption,'&&','&') + ' - ' + sBreakValue,sBreakValue);" & vbCrLf)

    Response.Write("			var frmMerge = window.parent.frames('dataframe').document.forms('frmCalendarMerge_'+lngPageCount);" & vbCrLf)
	
    Response.Write("			var dataCollection = frmMerge.elements;" & vbCrLf)
    Response.Write("			if (dataCollection!=null) " & vbCrLf)
    Response.Write("				{" & vbCrLf)
    Response.Write("				for (i=0; i<dataCollection.length; i++)  " & vbCrLf)
    Response.Write("					{" & vbCrLf)
    Response.Write("					strMergeString = dataCollection.item(i).value;" & vbCrLf)
    Response.Write("					if (strMergeString != '')" & vbCrLf)
    Response.Write("						{" & vbCrLf)
    Response.Write("						lngStartCol = Number(mergeArgument(strMergeString,'STARTCOL'));" & vbCrLf)
    Response.Write("						lngStartRow = Number(mergeArgument(strMergeString,'STARTROW'));" & vbCrLf)
    Response.Write("						lngEndCol = Number(mergeArgument(strMergeString,'ENDCOL'));" & vbCrLf)
    Response.Write("						lngEndRow = Number(mergeArgument(strMergeString,'ENDROW'));" & vbCrLf)
    Response.Write("						window.parent.parent.ASRIntranetOutput.AddMerge(lngStartCol,lngStartRow,lngEndCol,lngEndRow);" & vbCrLf)
    Response.Write("						}" & vbCrLf)
    Response.Write("					}" & vbCrLf)
    Response.Write("				}" & vbCrLf)

    Response.Write("			var frmStyle = window.parent.frames('dataframe').document.forms('frmCalendarStyle_'+lngPageCount);" & vbCrLf)
    Response.Write("			var dataCollection = frmStyle.elements;" & vbCrLf)
    Response.Write("			if (dataCollection!=null) " & vbCrLf)
    Response.Write("				{" & vbCrLf)
    Response.Write("				for (i=0; i<dataCollection.length; i++)  " & vbCrLf)
    Response.Write("					{" & vbCrLf)
    Response.Write("					strStyleString = dataCollection.item(i).value;" & vbCrLf)
    Response.Write("					if (strStyleString != '')" & vbCrLf)
    Response.Write("						{" & vbCrLf)
    Response.Write("						strType = styleArgument(strStyleString,'TYPE');" & vbCrLf)
    Response.Write("						lngStartCol = Number(styleArgument(strStyleString,'STARTCOL'));" & vbCrLf)
    Response.Write("						lngStartRow = Number(styleArgument(strStyleString,'STARTROW'));" & vbCrLf)
    Response.Write("						lngEndCol = Number(styleArgument(strStyleString,'ENDCOL'));" & vbCrLf)
    Response.Write("						lngEndRow = Number(styleArgument(strStyleString,'ENDROW'));" & vbCrLf)
    Response.Write("						lngBackCol = Number(styleArgument(strStyleString,'BACKCOLOR'));" & vbCrLf)
    Response.Write("						lngForeCol = Number(styleArgument(strStyleString,'FORECOLOR'));" & vbCrLf)
    Response.Write("						blnBold = styleArgument(strStyleString,'BOLD');" & vbCrLf)
    Response.Write("						blnUnderline = styleArgument(strStyleString,'UNDERLINE');" & vbCrLf)
    Response.Write("						blnGridlines = styleArgument(strStyleString,'GRIDLINES');" & vbCrLf)
    Response.Write("						window.parent.parent.ASRIntranetOutput.AddStyle(strType,lngStartCol,lngStartRow,lngEndCol,lngEndRow,lngBackCol,lngForeCol,blnBold,blnUnderline,blnGridlines);" & vbCrLf)
    Response.Write("						}" & vbCrLf)
    Response.Write("					}" & vbCrLf)
    Response.Write("				}" & vbCrLf)
	
    Response.Write("			for (var lngCol=0; lngCol<lngPageColumnCount; lngCol++)" & vbCrLf)
    Response.Write("				{" & vbCrLf)
    Response.Write("				window.parent.parent.ASRIntranetOutput.AddColumn(sColHeading, iColDataType, iColDecimals, false);" & vbCrLf)
    Response.Write("				}" & vbCrLf)
    Response.Write("			window.parent.parent.ASRIntranetOutput.DataArray();" & vbCrLf)
    Response.Write("			blnBreakCheck = true;" & vbCrLf)
    Response.Write("			sBreakValue = '';" & vbCrLf)
    Response.Write("			lngActualRow = 0;" & vbCrLf)
    Response.Write("			}" & vbCrLf)
	
    Response.Write("    else if ((frmOutput.grdCalendarOutput.Columns(0).CellText(bm) == '*')" & vbCrLf)
    Response.Write("					&& (!blnBreakCheck))" & vbCrLf)
    Response.Write("			{" & vbCrLf)
    Response.Write("			sBreakValue = frmOutput.grdCalendarOutput.Columns(1).CellText(bm);" & vbCrLf)
    Response.Write("			if ((sBreakValue == 'Key') && (frmExportData.txtFormat.value != '4')) " & vbCrLf)
    Response.Write("				{ " & vbCrLf)
    Response.Write("				window.parent.parent.ASRIntranetOutput.AddPage(replace(frmOutput.grdCalendarOutput.Caption,'&&','&') ,sBreakValue);" & vbCrLf)
    Response.Write("				} " & vbCrLf)
    Response.Write("			else " & vbCrLf)
    Response.Write("				{ " & vbCrLf)
    Response.Write("				window.parent.parent.ASRIntranetOutput.AddPage(replace(frmOutput.grdCalendarOutput.Caption,'&&','&') + ' - ' + sBreakValue,sBreakValue);" & vbCrLf)
    Response.Write("				} " & vbCrLf)
	
    Response.Write("			var frmMerge = window.parent.frames('dataframe').document.forms('frmCalendarMerge_'+lngPageCount);" & vbCrLf)
    Response.Write("			var dataCollection = frmMerge.elements;" & vbCrLf)
    Response.Write("			if (dataCollection!=null) " & vbCrLf)
    Response.Write("				{" & vbCrLf)
    Response.Write("				for (i=0; i<dataCollection.length; i++)  " & vbCrLf)
    Response.Write("					{" & vbCrLf)
    Response.Write("					strMergeString = dataCollection.item(i).value;" & vbCrLf)
    Response.Write("					if (strMergeString != '')" & vbCrLf)
    Response.Write("						{" & vbCrLf)
    Response.Write("						lngStartCol = Number(mergeArgument(strMergeString,'STARTCOL'));" & vbCrLf)
    Response.Write("						lngStartRow = Number(mergeArgument(strMergeString,'STARTROW'));" & vbCrLf)
    Response.Write("						lngEndCol = Number(mergeArgument(strMergeString,'ENDCOL'));" & vbCrLf)
    Response.Write("						lngEndRow = Number(mergeArgument(strMergeString,'ENDROW'));" & vbCrLf)
    Response.Write("						window.parent.parent.ASRIntranetOutput.AddMerge(lngStartCol,lngStartRow,lngEndCol,lngEndRow);" & vbCrLf)
    Response.Write("						}" & vbCrLf)
    Response.Write("					}" & vbCrLf)
    Response.Write("				}" & vbCrLf)

    Response.Write("			var frmStyle = window.parent.frames('dataframe').document.forms('frmCalendarStyle_'+lngPageCount);" & vbCrLf)
    Response.Write("			var dataCollection = frmStyle.elements;" & vbCrLf)
    Response.Write("			if (dataCollection!=null) " & vbCrLf)
    Response.Write("				{" & vbCrLf)
    Response.Write("				for (i=0; i<dataCollection.length; i++)  " & vbCrLf)
    Response.Write("					{" & vbCrLf)
    Response.Write("					strStyleString = dataCollection.item(i).value;" & vbCrLf)
    Response.Write("					if (strStyleString != '')" & vbCrLf)
    Response.Write("						{" & vbCrLf)
    Response.Write("						strType = styleArgument(strStyleString,'TYPE');" & vbCrLf)
    Response.Write("						lngStartCol = Number(styleArgument(strStyleString,'STARTCOL'));" & vbCrLf)
    Response.Write("						lngStartRow = Number(styleArgument(strStyleString,'STARTROW'));" & vbCrLf)
    Response.Write("						lngEndCol = Number(styleArgument(strStyleString,'ENDCOL'));" & vbCrLf)
    Response.Write("						lngEndRow = Number(styleArgument(strStyleString,'ENDROW'));" & vbCrLf)
    Response.Write("						lngBackCol = Number(styleArgument(strStyleString,'BACKCOLOR'));" & vbCrLf)
    Response.Write("						lngForeCol = Number(styleArgument(strStyleString,'FORECOLOR'));" & vbCrLf)
    Response.Write("						blnBold = styleArgument(strStyleString,'BOLD');" & vbCrLf)
    Response.Write("						blnUnderline = styleArgument(strStyleString,'UNDERLINE');" & vbCrLf)
    Response.Write("						blnGridlines = styleArgument(strStyleString,'GRIDLINES');" & vbCrLf)
    Response.Write("						window.parent.parent.ASRIntranetOutput.AddStyle(strType,lngStartCol,lngStartRow,lngEndCol,lngEndRow,lngBackCol,lngForeCol,blnBold,blnUnderline,blnGridlines);" & vbCrLf)
    Response.Write("						}" & vbCrLf)
    Response.Write("					}" & vbCrLf)
    Response.Write("				}" & vbCrLf)
	
    Response.Write("			for (var lngCol=0; lngCol<lngPageColumnCount; lngCol++)" & vbCrLf)
    Response.Write("				{" & vbCrLf)
    Response.Write("				window.parent.parent.ASRIntranetOutput.AddColumn(sColHeading, iColDataType, iColDecimals, false);" & vbCrLf)
    Response.Write("				}" & vbCrLf)
    Response.Write("      window.parent.parent.ASRIntranetOutput.DataArray();" & vbCrLf)
    Response.Write("			lngPageColumnCount = frmOutput.grdCalendarOutput.Columns.Count;" & vbCrLf)
    Response.Write("			if (!blnSettingsDone)" & vbCrLf)
    Response.Write("				{" & vbCrLf)
    Response.Write("				window.parent.parent.ASRIntranetOutput.HeaderRows = 2;" & vbCrLf)
    Response.Write("				window.parent.parent.ASRIntranetOutput.HeaderCols = 1;" & vbCrLf)
    Response.Write("				window.parent.parent.ASRIntranetOutput.SizeColumnsIndependently = true;" & vbCrLf)
    Response.Write("				blnSettingsDone = true;" & vbCrLf)
    Response.Write("				}" & vbCrLf)
    Response.Write("			window.parent.parent.ASRIntranetOutput.InitialiseStyles();" & vbCrLf)
    Response.Write("			window.parent.parent.ASRIntranetOutput.ResetStyles();" & vbCrLf)
    Response.Write("			window.parent.parent.ASRIntranetOutput.ResetColumns();" & vbCrLf)
    Response.Write("			window.parent.parent.ASRIntranetOutput.ResetMerges();" & vbCrLf)
    Response.Write("			lngPageCount++;" & vbCrLf)
    Response.Write("			window.parent.parent.ASRIntranetOutput.ArrayDim((lngPageColumnCount-1), 0);" & vbCrLf)
    Response.Write("			blnBreakCheck = true;" & vbCrLf)
    Response.Write("			sBreakValue = '';" & vbCrLf)
    Response.Write("			lngActualRow = 0;" & vbCrLf)
    Response.Write("			blnNewPage = true;" & vbCrLf)
    Response.Write("			}" & vbCrLf & vbCrLf)

    Response.Write("		else if (frmOutput.grdCalendarOutput.Columns(0).CellText(bm) != '*')" & vbCrLf)
    Response.Write("			{" & vbCrLf)
    Response.Write("			blnBreakCheck = false;" & vbCrLf)
    Response.Write("			blnNewPage = false;" & vbCrLf)
    Response.Write("			if (lngActualRow > 0)" & vbCrLf)
    Response.Write("				{" & vbCrLf)
    Response.Write("				window.parent.parent.ASRIntranetOutput.ArrayReDim();" & vbCrLf)
    Response.Write("				}" & vbCrLf)
    Response.Write("			for (var lngCol=0; lngCol<lngPageColumnCount; lngCol++)" & vbCrLf)
    Response.Write("				{" & vbCrLf)
    Response.Write("				window.parent.parent.ASRIntranetOutput.ArrayAddTo(lngCol, (lngActualRow), frmOutput.grdCalendarOutput.Columns(lngCol).CellText(bm));" & vbCrLf)
    Response.Write("				}" & vbCrLf)
    Response.Write("			}" & vbCrLf)
	
	
    Response.Write("		if (!blnNewPage) " & vbCrLf)
    Response.Write("			{" & vbCrLf)
    Response.Write("			lngActualRow = lngActualRow + 1; " & vbCrLf)
    Response.Write("			}" & vbCrLf)
    Response.Write("		}" & vbCrLf)
    Response.Write("		window.parent.parent.ASRIntranetOutput.Complete();" & vbCrLf)
    Response.Write("		ShowDataFrame();" & vbCrLf)
    Response.Write("		}" & vbCrLf)
    Response.Write("	}" & vbCrLf)
  
    Response.Write("	ShowDataFrame();" & vbCrLf)
	
    Response.Write("  try {" & vbCrLf)
    Response.Write("  if (frmOriginalDefinition.txtCancelPrint.value == 1) {" & vbCrLf)
    Response.Write("    OpenHR.messageBox(""Calendar Report '""+frmOriginalDefinition.txtDefn_Name.value+""' output failed.\n\nCancelled by user."",64,""Calendar Report"");" & vbCrLf)
    Response.Write("		window.focus();" & vbCrLf)
    Response.Write("  }" & vbCrLf)
    Response.Write("  else if (window.parent.parent.ASRIntranetOutput.ErrorMessage != """") {" & vbCrLf)
    Response.Write("    OpenHR.messageBox(""Calendar Report '""+frmOriginalDefinition.txtDefn_Name.value+""' output failed.\n\n""+window.parent.parent.ASRIntranetOutput.ErrorMessage,48,""Calendar Report"");" & vbCrLf)
    Response.Write("		window.focus();" & vbCrLf)
    Response.Write("		}" & vbCrLf)
    Response.Write("  else {" & vbCrLf)
    Response.Write("    OpenHR.messageBox(""Calendar Report '""+frmOriginalDefinition.txtDefn_Name.value+""' output complete."",64,""Calendar Report"");" & vbCrLf)
    Response.Write("		window.focus();" & vbCrLf)
    Response.Write("		}" & vbCrLf)
    Response.Write("	}" & vbCrLf)
    Response.Write("	catch (e) {}" & vbCrLf)
    Response.Write("	}" & vbCrLf)
  
    Response.Write("</script>" & vbCrLf & vbCrLf)

    Response.Write("<input type=hidden id=txtTitle name=txtTitle value=""" & objCalendar.CalendarReportName & """>" & vbCrLf)
	
    objCalendar = Nothing
	
%>
	
	<form id="frmOriginalDefinition" style="visibility:hidden;display:none">
        <%
            Dim sErrMsg As String
            Response.Write("	<INPUT type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Replace(Session("utilname"), """", "&quot;") & """>" & vbCrLf)
            Response.Write("	<INPUT type='hidden' id=txtDefn_ErrMsg name=txtDefn_ErrMsg value=""" & sErrMsg & """>" & vbCrLf)
        %>
        <input type="hidden" id="Hidden3" name="txtUserName" value="<%=session("username")%>">
		<input type="hidden" id="txtDateFormat" name="txtDateFormat" value="<%=session("LocaleDateFormat")%>">
		<input type="hidden" id="txtCancelPrint" name="txtCancelPrint">
		<input type="hidden" id="txtOptionsDone" name="txtOptionsDone">
		<input type="hidden" id="txtOptionsPortrait" name="txtOptionsPortrait">
		<input type="hidden" id="txtOptionsMarginLeft" name="txtOptionsMarginLeft">
		<input type="hidden" id="txtOptionsMarginRight" name="txtOptionsMarginRight">
		<input type="hidden" id="txtOptionsMarginTop" name="txtOptionsMarginTop">
		<input type="hidden" id="txtOptionsMarginBottom" name="txtOptionsMarginBottom">
		<input type="hidden" id="txtOptionsCopies" name="txtOptionsCopies">
		<input type="hidden" id="txtCalRep_UtilID" name="txtCalRep_UtilID" value="<%=Request("CalRepUtilID")%>">
	</form>


<script type="text/javascript">
    util_run_calendarreport_window_onload();
</script>
