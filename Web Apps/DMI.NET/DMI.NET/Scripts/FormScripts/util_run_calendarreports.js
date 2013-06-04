
//
//Taken from util_run_calendarrepoart_nav.ascx
//
var frmUseful = OpenHR.getForm("calendarworkframe", "frmUseful");

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

//TODO - This function name is duplicated.
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

function trim(strInput) {

    if (strInput.length < 1 ){
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

    var frmUseful = OpenHR.getForm("calendarworkframe", "frmUseful");
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

function enableDisableNavigation() {

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
        document.getElementById('imgPrevMonth').src = window.ROOT + "Content/images/CalendarReports/previous_disabled.gif";
        document.getElementById('imgFirstMonth').src = window.ROOT + "Content/images/CalendarReports/first_disabled.gif";
        image_disable(document.getElementById('imgPrevMonth'), true);
        image_disable(document.getElementById('imgFirstMonth'), true);
        document.getElementById('imgPrevMonth').style.cursor = 'default';
        document.getElementById('imgFirstMonth').style.cursor = 'default';
    }
    else
    {
        bNextEnabled = true;
        document.getElementById('imgPrevMonth').src = window.ROOT + "Content/images/CalendarReports/previous_enabled.gif";
        document.getElementById('imgFirstMonth').src = window.ROOT + "Content/images/CalendarReports/first_enabled.gif";
        image_disable(document.getElementById('imgPrevMonth'), false);
        image_disable(document.getElementById('imgFirstMonth'), false);
        document.getElementById('imgPrevMonth').style.cursor = 'hand';
        document.getElementById('imgFirstMonth').style.cursor = 'hand';
    }
	
    if (dtShownEnd >= dtReportEnd)
    {
        bPrevEnabled = false;
        document.getElementById('imgNextMonth').src = window.ROOT + "Content/images/CalendarReports/next_disabled.gif";
        document.getElementById('imgLastMonth').src = window.ROOT + "Content/images/CalendarReports/last_disabled.gif";
        image_disable(document.getElementById('imgNextMonth'), true);
        image_disable(document.getElementById('imgLastMonth'), true);
        document.getElementById('imgNextMonth').style.cursor = 'default';
        document.getElementById('imgLastMonth').style.cursor = 'default';
    }
    else
    {
        bPrevEnabled = true;
        document.getElementById('imgNextMonth').src = window.ROOT + "Content/images/CalendarReports/Next_Enabled.gif";
        document.getElementById('imgLastMonth').src = window.ROOT + "Content/images/CalendarReports/Last_Enabled.gif";
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
        document.getElementById('imgToday').src = window.ROOT + "content/images/CalendarReports/today_disabled.gif";
        image_disable(document.getElementById('imgToday'), true);
        document.getElementById('imgToday').style.cursor = 'default';
    }
    else if ((dtSystemMonth < dtReportEndMonth) && (dtSystemMonth > dtReportStartMonth))
    {
        document.getElementById('imgToday').src = window.ROOT + "content/images/CalendarReports/today_enabled.gif";
        image_disable(document.getElementById('imgToday'), false);
        document.getElementById('imgToday').style.cursor = 'hand';
    }
    else
    {
        document.getElementById('imgToday').src = window.ROOT + "content/images/CalendarReports/today_disabled.gif";
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
        var strMessage = "The selected date is outside of the report date boundaries.";
        var dtShownStart = createDate('01/'+frmNav.cboMonth.options[frmNav.cboMonth.selectedIndex].value+'/'+frmNav.txtYear.value);
        var dtShownEnd = createDate(frmDate.txtDaysInMonth.value+'/'+frmNav.cboMonth.options[frmNav.cboMonth.selectedIndex].value+'/'+frmNav.txtYear.value);
        var dtReportStart = createDate(frmDate.txtReportStartDate.value);	
        var dtReportEnd = createDate(frmDate.txtReportEndDate.value);
		
        if (dtShownStart > dtReportEnd)
        {
            OpenHR.messageBox(strMessage,48,"Calendar Reports");
            window.focus();
            //lastMonth();			
            frmNav.cboMonth.selectedIndex = Number(frmDate.txtCurrentMonthIndex.value);
            frmNav.txtYear.value = Number(frmDate.txtCurrentYearValue.value);
            return;
        }
        else if (dtShownEnd < dtReportStart)
        {
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

    refreshCalendarData();
		
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
		
function spinRecords(pfUp) {

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

function styleArgument(psDefnString, psParameter) {
    
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

function mergeArgument(psDefnString, psParameter) {
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

// REMOVED - function writeText()

//
//Taken from util_run_calendarrepoart_main.ascx
//

// REMOVED - function refreshDefSelAndClose()

// REMOVED - function replace()

//TODO - This function name is duplicated.
function dataOnlyPrint() {
    var lngPageColumnCount = 3;
    var lngActualRow = new Number(0);
    var blnSettingsDone = false;
    var sColHeading = new String('');
    var iColDataType = new Number(12);
    var iColDecimals = new Number(0);
    var blnNewPage = true;
    var lngPageCount = new Number(0);

    frmOutput.grdCalendarOutput.focus();
    frmOutput.grdCalendarOutput.caption = replace(frmOutput.grdCalendarOutput.caption, '&', '&&');

    if (frmOutput.ssHiddenGrid.Columns.Count > 0) {
        frmOutput.ssHiddenGrid.Columns.RemoveAll();
    }

    if (frmOutput.ssHiddenGrid.Rows > 0) {
        frmOutput.ssHiddenGrid.RemoveAll();
    }

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
    for (var lngRow = 0; lngRow < frmOutput.grdCalendarOutput.Rows; lngRow++) {
        bm = frmOutput.grdCalendarOutput.AddItemBookmark(lngRow);

        if (lngRow == (frmOutput.grdCalendarOutput.Rows - 1)) {
            sBreakValue = frmOutput.grdCalendarOutput.Columns(1).CellText(bm);
            frmOutput.ssHiddenGrid.Caption = txtTitle.value + ' - ' + sBreakValue;

            // PRINT DATA
            if (frmOriginalDefinition.txtOptionsDone.value == 0) {
                frmOutput.ssHiddenGrid.PrintData(23, false, true);
                frmOriginalDefinition.txtOptionsDone.value = 1;
                if (frmOriginalDefinition.txtCancelPrint.value == 1) {
                    frmOutput.grdCalendarOutput.redraw = true;
                    return;
                }
            }
            else {
                frmOutput.ssHiddenGrid.PrintData(23, false, false);
            }
            frmOutput.ssHiddenGrid.RemoveAll();
            blnBreakCheck = true;
            sBreakValue = '';
            lngActualRow = 0;
        }
        else if ((frmOutput.grdCalendarOutput.Columns(0).CellText(bm) == '*')
              && (!blnBreakCheck)) {
            sBreakValue = frmOutput.grdCalendarOutput.Columns(1).CellText(bm);
            frmOutput.ssHiddenGrid.Caption = txtTitle.value + ' - ' + sBreakValue;

            // PRINT DATA
            if (frmOriginalDefinition.txtOptionsDone.value == 0) {
                frmOutput.ssHiddenGrid.PrintData(23, false, true);
                frmOriginalDefinition.txtOptionsDone.value = 1;
                if (frmOriginalDefinition.txtCancelPrint.value == 1) {
                    frmOutput.grdCalendarOutput.redraw = true;
                    return;
                }
            }
            else {
                frmOutput.ssHiddenGrid.PrintData(23, false, false);
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
        else if (frmOutput.grdCalendarOutput.Columns(0).CellText(bm) != '*') {
            if (blnNewPage) {
                for (var lngCol = 0; lngCol < lngPageColumnCount; lngCol++) {
                    // ADD COLUMN TO GRID
                    frmOutput.ssHiddenGrid.Columns.Add(lngCol);
                    frmOutput.ssHiddenGrid.Columns(lngCol).width = 15;
                }
            }
            blnBreakCheck = false;
            blnNewPage = false;

            // CREATE TAB DELIMITED STRING AND ADD TO GRID
            sAddItem = new String("");
            for (var lngCol = 0; lngCol < lngPageColumnCount; lngCol++) {
                if (frmOutput.grdCalendarOutput.Columns(lngCol).visible) {
                    if (lngCol > 0) {
                        sAddItem = sAddItem + "	";
                    }
                    var sValue = new String(trim(frmOutput.grdCalendarOutput.Columns(lngCol).CellText(bm)));
                    sAddItem = sAddItem + sValue;
                    if ((sValue.length * 8) > frmOutput.ssHiddenGrid.Columns(lngCol).width) {
                        frmOutput.ssHiddenGrid.Columns(lngCol).width = (sValue.length * 8);
                    }
                }
            }
            frmOutput.ssHiddenGrid.AddItem(sAddItem);
        }

        if (!blnNewPage) {
            lngActualRow = lngActualRow + 1;
        }
    }
}

// REMOVED - function trim()

// REMOVED - function styleArguement()

// REMOVED - function mergeArguement()

// REMOVED - function getDBName()


function loadAddRecords(sFrom) {
    var iCount;

    iCount = new Number(txtLoadCount.value);

    txtLoadCount.value = iCount + 1;

    if (iCount > 1) {
        startCalendar();
    }

}


//
//Taken from util_run_calendarrepoart_data.ascx
//

function util_run_calendarreport_data_window_onload() {

    if (txtFirstLoad.value == 1) {
        loadAddRecords('data');
        return;
    }

    if (frmCalendarData.txtCalendarMode.value == "LOADCALENDARREPORTDATA") {
        fillCalBoxes();
    }
    else if (frmCalendarData.txtCalendarMode.value == "OUTPUTREPORT") {
        setGridFont(frmCalendarData.grdCalendarOutput);
        setGridFont(frmCalendarData.ssHiddenGrid);

        outputReport();
    }
}

function ExportData(strMode) {  
    var frmGetDataForm = OpenHR.getForm("dataframe", "frmCalendarGetData");
    frmGetDataForm.txtMode.value = "OUTPUTREPORT";
    refreshCalendarData();
}
	
function refreshCalendarData() {
    var frmGetData = OpenHR.getForm("dataframe", "frmCalendarGetData");
    OpenHR.submitForm(frmGetData);
}

//
//Taken from util_run_calendarrepoart_options.ascx
//

function refreshInfo() {
    var frmUseful = OpenHR.getForm("calendarworkframe", "frmUseful");

    if (frmUseful.txtLoading.value == 0) {
        refreshDateSpecifics();
    }

    setOptions();
    return true;
}

function setOptions() {
    var frmNavFillerOptions = OpenHR.getForm("workframefiller", "frmNavFillerOptions");

    with (frmNavFillerOptions) {
        txtIncludeBankHolidays.value = (frmOptions.chkIncludeBHols.checked);
        txtIncludeWorkingDaysOnly.value = (frmOptions.chkIncludeWorkingDaysOnly.checked);
        txtShowBankHolidays.value = (frmOptions.chkShadeBHols.checked);
        txtShowCaptions.value = (frmOptions.chkCaptions.checked);
        txtShowWeekends.value = (frmOptions.chkShadeWeekends.checked);
    }
    
    OpenHR.submitForm(frmNavFillerOptions);
    return; 
}


//
//Taken from util_run_calendarreport_calendar.ascx
//

function util_run_calendarreport_calendar_window_onload() {
    loadAddRecords('calendar');
}

// REMOVED - function openDialog()


//
//Taken from util_run_calendarreport_key.ascx
//

function populateKey() {
    var strKey, strDescription, strCode;
    var lngColour;
    var strControlName = '';

    frmKey.ctlKey.Clear_Key();

    for (var i = 1; i <= frmKeyInfo.key_Count.value; i++) {
        strControlName = 'key_ID' + i;
        strKey = document.getElementById(strControlName).getAttribute('value');
        strControlName = 'key_Name' + i;
        strDescription = document.getElementById(strControlName).getAttribute('value');
        strControlName = 'key_Code' + i;
        strCode = document.getElementById(strControlName).getAttribute('value');
        strControlName = 'key_Colour' + i;
        lngColour = Number(document.getElementById(strControlName).getAttribute('value'));

        frmKey.ctlKey.Add_Key(strKey, strDescription, strCode, lngColour);
    }

    frmKey.ctlKey.Sort();

    if (frmKeyInfo.txtHasMultiple.value == '1') {
        strKey = "EVENT_MULTIPLE";
        strDescription = "Multiple Events";
        strCode = ".";
        lngColour = 16777215;

        frmKey.ctlKey.Add_Key(strKey, strDescription, strCode, lngColour);
    }

    return true;
}



//
//Taken from util_run_calendarreport_main.ascx
//

function util_run_calendarreport_main_window_onload() {
    
    if ((txtPreview.value == 0) && (txtOK.value == "True")) {
        setGridFont(frmOutput.grdCalendarOutput);
        setGridFont(frmOutput.ssHiddenGrid);

        outputCalendarReport();
        document.getElementById('tdDisplay').innerText = 'Calendar Report Output Complete.';
        document.getElementById('Cancel').value = 'OK';
    }
}
