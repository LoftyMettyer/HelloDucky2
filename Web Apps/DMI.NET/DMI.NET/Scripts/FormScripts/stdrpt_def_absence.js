
function stdrpt_def_absence_window_onload() {

    $("#workframe").attr("data-framesource", "STDRPT_DEF_ABSENCE");
        
    // Set the frameset as 1 because 0 doesn't clear any combo boxes /dropdown controls
    if (frmPostDefinition.txtRecSelCurrentID.value > 0) {
        $("#workframe").show();
    }

    menu_refreshMenu();

    populatePrinters();
    SetReportDefaults();
    displayPage(1);
    refreshTab3Controls();
	
    // Disable the menu
    menu_disableMenu();

}

function changeTab1Control() {
    frmAbsenceUseful.txtChanged.value = 1;
    refreshTab1Controls();
}

function changeTab2Control() {
    frmAbsenceUseful.txtChanged.value = 1;
    refreshTab2Controls();
}

function changeTab3Control() {
    frmAbsenceUseful.txtChanged.value = 1;
    refreshTab3Controls();
}

function populatePrinters()
{
    //with (frmAbsenceDefinition.cboPrinterName)
    //{

    //    strCurrentPrinter = '';
    //    if (selectedIndex > 0) 
    //    {
    //        strCurrentPrinter = options[selectedIndex].innerText;
    //    }

    //    length = 0;
    //    var oOption = document.createElement("OPTION");
    //    options.add(oOption);
    //    oOption.innerText = "<Default Printer>";
    //    oOption.value = 0;

    //    for (iLoop=0; iLoop<window.parent.frames("menuframe").ASRIntranetFunctions.PrinterCount(); iLoop++)  
    //    {
    //        var oOption = document.createElement("OPTION");
    //        options.add(oOption);
    //        oOption.innerText = window.parent.frames("menuframe").ASRIntranetFunctions.PrinterName(iLoop);
    //        oOption.value = iLoop+1;

    //        if (oOption.innerText == strCurrentPrinter) 
    //        {
    //            selectedIndex = iLoop+1
    //        }
    //    }

    //    if (strCurrentPrinter != '') 
    //    {
    //        if (frmAbsenceDefinition.cboPrinterName.options(frmAbsenceDefinition.cboPrinterName.selectedIndex).innerText != strCurrentPrinter) 
    //        {
    //            var oOption = document.createElement("OPTION");
    //            frmAbsenceDefinition.cboPrinterName.options.add(oOption);
    //            oOption.innerText = strCurrentPrinter;
    //            oOption.value = frmAbsenceDefinition.cboPrinterName.options.length-1;
    //            selectedIndex = oOption.value;
    //        }
    //    }
    //}	
}

function formatClick(index)
{
    var fViewing = (frmAbsenceUseful.txtAction.value.toUpperCase() == "VIEW");

    checkbox_disable(frmAbsenceDefinition.chkPreview, ((index == 0) || (fViewing == true)))
    frmAbsenceDefinition.chkPreview.checked = (index != 0);

    frmAbsenceDefinition.chkDestination0.checked = false;
    frmAbsenceDefinition.chkDestination1.checked = false;
    frmAbsenceDefinition.chkDestination2.checked = false;
    frmAbsenceDefinition.chkDestination3.checked = false;

    if (index == 1) {
        frmAbsenceDefinition.chkDestination2.checked = true;
        frmAbsenceDefinition.cboSaveExisting.length = 0;
        frmAbsenceDefinition.txtFilename.value = '';		
    }
    else {
        frmAbsenceDefinition.chkDestination0.checked = true;
    }

    frmAbsenceUseful.txtChanged.value = 1;
    refreshTab3Controls();
}

function selectEmailGroup()
{
    var sURL;
	
    frmEmailSelection.EmailSelCurrentID.value = frmAbsenceDefinition.txtEmailGroupID.value; 

    sURL = "util_emailSelection" +
        "?EmailSelCurrentID=" + frmEmailSelection.EmailSelCurrentID.value;
    openDialog(sURL, (screen.width)/3,(screen.height)/2, "yes", "yes");
}

function validateNumeric(pobjNumericControl)
{
    var sValue = pobjNumericControl.value;

    if (sValue.length == 0) 
    {            
        OpenHR.messageBox("Invalid numeric value entered.");
        pobjNumericControl.focus();
        return false;
    }
    else 
    {
        if (isNaN(sValue) == true)
        {
            OpenHR.messageBox("Invalid numeric value entered.");
            pobjNumericControl.focus();
            return false;
        }
        else 
        {
            return true;
        }
    }	
}

function validateDate(pobjDateControl)
{
    // Date column.
    // Ensure that the value entered is a date.

    var sValue = pobjDateControl.value;
	
    if (sValue.length == 0) 
    {
        //		OpenHR.messageBox("Invalid date value entered.");
        //		pobjDateControl.focus()
        return false;
    }
    else 
    {
        // Convert the date to SQL format (use this as a validation check).
        // An empty string is returned if the date is invalid.
        sValue = absencedef_convertLocaleDateToSQL(sValue);
        if (sValue.length == 0) 
        {
            OpenHR.messageBox("Invalid date value entered.");
            pobjDateControl.value = "";
            pobjDateControl.focus();
            return false;
        }
        else 
        {
            return true;
        }
    }
}

function validateAbsenceTab3()
{
    var sErrMsg;
	
    sErrMsg = "";
	
    if (!frmAbsenceDefinition.chkDestination0.checked 
        && !frmAbsenceDefinition.chkDestination1.checked 
        && !frmAbsenceDefinition.chkDestination2.checked 
        && !frmAbsenceDefinition.chkDestination3.checked)
    {
        sErrMsg = "You must select a destination";
    }

    if ((frmAbsenceDefinition.txtFilename.value == "") &&
        (frmAbsenceDefinition.cmdFilename.disabled == false)) {
        sErrMsg = "You must enter a file name";
    }

    if ((frmAbsenceDefinition.txtEmailGroup.value == "") &&
        (frmAbsenceDefinition.cmdEmailGroup.disabled == false)) {
        sErrMsg = "You must select an email group";
    }
	
    if ((frmAbsenceDefinition.chkDestination3 .checked) 
        && (frmAbsenceDefinition.txtEmailAttachAs.value == ''))
    {
        sErrMsg = "You must enter an email attachment file name.";
    }

    if (frmAbsenceDefinition.chkDestination3.checked &&
        (frmAbsenceDefinition.optOutputFormat3.checked || frmAbsenceDefinition.optOutputFormat4.checked || frmAbsenceDefinition.optOutputFormat5.checked || frmAbsenceDefinition.optOutputFormat6.checked) &&
        frmAbsenceDefinition.txtEmailAttachAs.value.match(/.html$/)) {
        sErrMsg = "You cannot email html output from word or excel.";
    }

    if (sErrMsg.length > 0) 
    {    
        OpenHR.messageBox(sErrMsg,48);
        displayPage(3);
        return (false);
    }
	
    try 
    {
        var testDataCollection = frmRefresh.elements;
        OpenHR.submitForm(frmRefresh);
    }
    catch(e) 
    {
    }
		
    return (true);
}

function absence_okClick(){

    var fOK = true;
    var dataCollection = frmAbsenceDefinition.elements;

    var frmRefresh = OpenHR.getForm("refreshframe", "frmRefresh");
    OpenHR.submitForm(frmRefresh);

    var sValue = frmAbsenceDefinition.txtDateFrom.value;
    if (sValue.length == 0) 
    {
        fOK = false;
    }
    else
    {
        sValue = absencedef_convertLocaleDateToSQL(sValue);
        if (sValue.length == 0)
        {
            fOK = false;
        }
        else
        {
            frmAbsenceDefinition.txtDateFrom.value = OpenHR.ConvertSQLDateToLocale(sValue);
        }
    }
			
    if (fOK == false)
    {
        OpenHR.messageBox("Invalid start date value entered.");
        displayPage(1);		
        frmAbsenceDefinition.txtDateFrom.focus();
        return;
    }

    sValue = frmAbsenceDefinition.txtDateTo.value;
    if (sValue.length == 0)
    {
        fOK = false;
    }
    else
    {
        sValue = absencedef_convertLocaleDateToSQL(sValue);
        if (sValue.length == 0) 
        {
            fOK = false;
        }
        else 
        {
            frmAbsenceDefinition.txtDateTo.value = OpenHR.ConvertSQLDateToLocale(sValue);               
        }
    }

    if (fOK == false)
    {
        OpenHR.messageBox("Invalid end date value entered.");
        displayPage(1);
        frmAbsenceDefinition.txtDateTo.focus();
        return;
    }

    //Check if report end date is before the report start date
    with (frmAbsenceDefinition.txtDateFrom.value.toString()) {
        lngStart = substr(6,4) + substr(3,2) + substr(0,2);	//yyyymmdd
    }
    with (frmAbsenceDefinition.txtDateTo.value.toString()) {
        lngEnd = substr(6,4) + substr(3,2) + substr(0,2);	//yyyymmdd
    }
    if (lngEnd < lngStart) {
        OpenHR.messageBox("The report end date is before the report start date.");
        displayPage(1);
        frmAbsenceDefinition.txtDateFrom.focus();
        return;
    }

    frmPostDefinition.txtFromDate.value = frmAbsenceDefinition.txtDateFrom.value;
    frmPostDefinition.txtToDate.value = frmAbsenceDefinition.txtDateTo.value;
    frmPostDefinition.txtAbsenceTypes.value = "";

    if (dataCollection!=null) 
    {
        for (iIndex=0; iIndex<dataCollection.length; iIndex++)  
        {
            sControlName = dataCollection.item(iIndex).name;

            if (sControlName.substr(0, 15) == "chkAbsenceType_") 
            {
                if (dataCollection.item(iIndex).checked == true)
                {
                    frmPostDefinition.txtAbsenceTypes.value = frmPostDefinition.txtAbsenceTypes.value + dataCollection.item(iIndex).attributes[7].nodeValue + ",";
                }
            }
        }
    }


    if (frmPostDefinition.txtAbsenceTypes.value == "")
    {
        OpenHR.messageBox("You must have at least 1 absence type selected.");
        displayPage(1);		
        fOK = false;
    }

    frmPostDefinition.utilid.value = "0";	
    if (frmAbsenceDefinition.optPickList.checked == true) frmPostDefinition.utilid.value = "0";
    if (frmAbsenceDefinition.optPickList.checked == true) frmPostDefinition.utilid.value = frmPostDefinition.txtBasePicklistID.value;
    if (frmAbsenceDefinition.optFilter.checked == true) frmPostDefinition.utilid.value = frmPostDefinition.txtBaseFilterID.value;
    if ((frmAbsenceDefinition.optPickList.checked == true) && (frmPostDefinition.txtBasePicklistID.value == "0")) 
    {
        OpenHR.messageBox("You must have a picklist selected.");
        displayPage(1);
        fOK = false;
    }
		
    if ((frmAbsenceDefinition.optFilter.checked == true) && (frmPostDefinition.txtBaseFilterID.value == "0"))
    {
        OpenHR.messageBox("You must have a filter selected.");
        displayPage(1);		
        fOK = false;
    }

    frmPostDefinition.txtPrintFPinReportHeader.value = frmAbsenceDefinition.chkPrintInReportHeader.checked;

    // Bradford Specific data
    frmPostDefinition.txtSRV.value = frmAbsenceDefinition.chkSRV.checked;
    frmPostDefinition.txtShowDurations.value = frmAbsenceDefinition.chkShowDurations.checked;
    frmPostDefinition.txtShowInstances.value = frmAbsenceDefinition.chkShowInstances.checked;
    frmPostDefinition.txtShowFormula.value = frmAbsenceDefinition.chkShowFormula.checked;
    frmPostDefinition.txtOmitBeforeStart.value = frmAbsenceDefinition.chkOmitBeforeStart.checked;
    frmPostDefinition.txtOmitAfterEnd.value = frmAbsenceDefinition.chkOmitAfterEnd.checked;
    frmPostDefinition.txtOrderBy1.value = frmAbsenceDefinition.cboOrderBy1.options[frmAbsenceDefinition.cboOrderBy1.selectedIndex].text;
    frmPostDefinition.txtOrderBy1ID.value = frmAbsenceDefinition.cboOrderBy1.options[frmAbsenceDefinition.cboOrderBy1.selectedIndex].value;
    frmPostDefinition.txtOrderBy1Asc.value = frmAbsenceDefinition.chkOrderBy1Asc.checked;
    frmPostDefinition.txtOrderBy2.value = frmAbsenceDefinition.cboOrderBy2.options[frmAbsenceDefinition.cboOrderBy2.selectedIndex].text;
    frmPostDefinition.txtOrderBy2ID.value = frmAbsenceDefinition.cboOrderBy2.options[frmAbsenceDefinition.cboOrderBy2.selectedIndex].value;
    frmPostDefinition.txtOrderBy2Asc.value = frmAbsenceDefinition.chkOrderBy2Asc.checked;
    frmPostDefinition.txtMinimumBradfordFactor.value = frmAbsenceDefinition.chkMinimumBradfordFactor.checked;
    frmPostDefinition.txtMinimumBradfordFactorAmount.value = frmAbsenceDefinition.txtMinimumBradfordFactor.value;
    frmPostDefinition.txtDisplayBradfordDetail.value = frmAbsenceDefinition.chkShowAbsenceDetails.checked;
	
    // Validate the output options
    if (fOK == true)
    {
        if (validateAbsenceTab3() == false) 
        {
            return;
        }
    }

    if (frmAbsenceDefinition.chkPreview.checked == true)
    {
        frmPostDefinition.txtSend_OutputPreview.value = 1;
    }
    else
    {
        frmPostDefinition.txtSend_OutputPreview.value = 0;
    }
	
    frmPostDefinition.txtSend_OutputFormat.value = 0;
    if (frmAbsenceDefinition.optOutputFormat1.checked)	frmPostDefinition.txtSend_OutputFormat.value = 1;
    if (frmAbsenceDefinition.optOutputFormat2.checked)	frmPostDefinition.txtSend_OutputFormat.value = 2;
    if (frmAbsenceDefinition.optOutputFormat3.checked)	frmPostDefinition.txtSend_OutputFormat.value = 3;
    if (frmAbsenceDefinition.optOutputFormat4.checked)	frmPostDefinition.txtSend_OutputFormat.value = 4;
    if (frmAbsenceDefinition.optOutputFormat5.checked)	frmPostDefinition.txtSend_OutputFormat.value = 5;
    if (frmAbsenceDefinition.optOutputFormat6.checked)	frmPostDefinition.txtSend_OutputFormat.value = 6;
	
    if (frmAbsenceDefinition.chkDestination0.checked == true)
    {
        frmPostDefinition.txtSend_OutputScreen.value = 1;
    }
    else
    {
        frmPostDefinition.txtSend_OutputScreen.value = 0;
    }
	
    if (frmAbsenceDefinition.chkDestination1.checked == true)
    {
        frmPostDefinition.txtSend_OutputPrinter.value = 1;
        frmPostDefinition.txtSend_OutputPrinterName.value = frmAbsenceDefinition.cboPrinterName.options[frmAbsenceDefinition.cboPrinterName.selectedIndex].innerText;
    }
    else
    {
        frmPostDefinition.txtSend_OutputPrinter.value = 0;
        frmPostDefinition.txtSend_OutputPrinterName.value = '';
    }
	
    if (frmAbsenceDefinition.chkDestination2.checked == true)
    {
        frmPostDefinition.txtSend_OutputSave.value = 1;
        frmPostDefinition.txtSend_OutputSaveExisting.value = frmAbsenceDefinition.cboSaveExisting.options[frmAbsenceDefinition.cboSaveExisting.selectedIndex].value;
    }
    else
    {
        frmPostDefinition.txtSend_OutputSave.value = 0;
        frmPostDefinition.txtSend_OutputSaveExisting.value = 0;
    }
	
    if (frmAbsenceDefinition.chkDestination3.checked == true)
    {
        frmPostDefinition.txtSend_OutputEmail.value = 1;
        frmPostDefinition.txtSend_OutputEmailAddr.value = frmAbsenceDefinition.txtEmailGroupID.value;
        frmPostDefinition.txtSend_OutputEmailSubject.value = frmAbsenceDefinition.txtEmailSubject.value;
        frmPostDefinition.txtSend_OutputEmailAttachAs.value = frmAbsenceDefinition.txtEmailAttachAs.value;
    }
    else
    {
        frmPostDefinition.txtSend_OutputEmail.value = 0;
        frmPostDefinition.txtSend_OutputEmailAddr.value = 0;
        frmPostDefinition.txtSend_OutputEmailSubject.value = '';
        frmPostDefinition.txtSend_OutputEmailAttachAs.value = '';
    }

    frmPostDefinition.txtSend_OutputFilename.value = frmAbsenceDefinition.txtFilename.value;

    if (fOK == true) {
        var sUtilID = new String(16);
        frmPostDefinition.target = sUtilID;
        OpenHR.showInReportFrame(frmPostDefinition);
    }

    return;
}

function selectRecordOption(psType) {	
    var sURL;
		
    if (psType == 'picklist') {
        iCurrentID = frmPostDefinition.txtBasePicklistID.value;
    }
    else {
        iCurrentID = frmPostDefinition.txtBaseFilterID.value;
    }

    frmRecordSelection.recSelType.value = psType;
    frmRecordSelection.recSelTableID.value = frmSessionInformation.txtPersonnelTableID.value;
    frmRecordSelection.recSelCurrentID.value = iCurrentID; 
	
    sURL = "util_recordSelection" +
        "?recSelType=" + escape(frmRecordSelection.recSelType.value) +
        "&recSelTableID=" + escape(frmRecordSelection.recSelTableID.value) + 
        "&recSelCurrentID=" + escape(frmRecordSelection.recSelCurrentID.value) +
        "&recSelTable=" + escape(frmRecordSelection.recSelTable.value);
    openDialog(sURL, (screen.width)/3,(screen.height)/2, "yes", "yes");
}

function changeRecordOptions(psType)
{

    if (psType == 'picklist') 
    {
        button_disable(frmAbsenceDefinition.cmdBasePicklist, false);
        button_disable(frmAbsenceDefinition.cmdBaseFilter, true);
		
        frmAbsenceDefinition.optAllRecords.checked = false;
        frmAbsenceDefinition.optFilter.checked = false;
        frmAbsenceDefinition.txtBaseFilter.value = "";
        frmPostDefinition.txtBaseFilter.value = "";
        frmPostDefinition.txtBaseFilterID.value = 0;	
    }

    if (psType == 'filter') 
    {
        button_disable(frmAbsenceDefinition.cmdBasePicklist, true);
        button_disable(frmAbsenceDefinition.cmdBaseFilter, false);

        frmAbsenceDefinition.optAllRecords.checked = false;	
        frmAbsenceDefinition.optPickList.checked = false;
        frmAbsenceDefinition.txtBasePicklist.value = "";
        frmPostDefinition.txtBasePicklist.value = "";
        frmPostDefinition.txtBasePicklistID.value = 0;
    }

    if (psType == 'all') 
    {
        button_disable(frmAbsenceDefinition.cmdBasePicklist, true);
        button_disable(frmAbsenceDefinition.cmdBaseFilter, true);
	
        frmAbsenceDefinition.optPickList.checked = false;
        frmAbsenceDefinition.optFilter.checked = false;

        frmAbsenceDefinition.txtBasePicklist.value = "";
        frmPostDefinition.txtBasePicklist.value = "";
        frmPostDefinition.txtBasePicklistID.value = 0;

        frmAbsenceDefinition.txtBaseFilter.value = "";
        frmPostDefinition.txtBaseFilter.value = "";
        frmPostDefinition.txtBaseFilterID.value = 0;	
    }

    refreshTab1Controls();

}

function refreshControls()
{
}

function openWindow(mypage, myname, w, h, scroll)
{
    var winl = (screen.width - w) / 2;
    var wint = (screen.height - h) / 2;

    if (scroll == 'no')	{
        winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',scrollbars='+scroll+',resize=no';
    }
    else {
        winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',scrollbars='+scroll+',resizable';
    }

    win = window.open(mypage, myname, winprops);
    if (win.opener == null) win.opener = self;
    if (parseInt(navigator.appVersion) >= 4) win.window.focus();
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

function displayPage(piPageNumber) {
    var iLoop;
    var iCurrentChildCount;
    //window.parent.frames("refreshframe").document.forms("frmRefresh").submit();
			
    if (piPageNumber == 1) 
    {
        div1.style.visibility="visible";
        div1.style.display="block";
        div2.style.visibility="hidden";
        div2.style.display="none";
        div3.style.visibility="hidden";
        div3.style.display="none";
        button_disable(frmAbsenceDefinition.btnTab1, true);
        if (frmSessionInformation.txtUtilID.value = 16) 
        {
            button_disable(frmAbsenceDefinition.btnTab2, false);
        }
        button_disable(frmAbsenceDefinition.btnTab3, false);
        refreshTab1Controls();
    }

    if (piPageNumber == 2) 
    {
        div1.style.visibility="hidden";
        div1.style.display="none";
        div2.style.visibility="visible";
        div2.style.display="block";		
        div3.style.visibility="hidden";
        div3.style.display="none";
        button_disable(frmAbsenceDefinition.btnTab1, false);
        button_disable(frmAbsenceDefinition.btnTab2, true);
        button_disable(frmAbsenceDefinition.btnTab3, false);
        refreshTab2Controls();	
    }

    if (piPageNumber == 3) {
        div1.style.visibility="hidden";
        div1.style.display="none";
        div2.style.visibility="hidden";
        div2.style.display="none";
        div3.style.visibility="visible";
        div3.style.display="block";
        button_disable(frmAbsenceDefinition.btnTab1, false);
        if (frmSessionInformation.txtUtilID.value = 16) 
        {
            button_disable(frmAbsenceDefinition.btnTab2, false);
        }
        button_disable(frmAbsenceDefinition.btnTab3, true);
        refreshTab3Controls();
    }
}

function refreshTab1Controls()
{		 
    if (frmAbsenceDefinition.optAllRecords.checked == true) 
    {
        checkbox_disable(frmAbsenceDefinition.chkPrintInReportHeader, true);
        frmAbsenceDefinition.chkPrintInReportHeader.checked = false;
    }
    else 
    {
        checkbox_disable(frmAbsenceDefinition.chkPrintInReportHeader, false);
    }
}

function refreshTab2Controls()
{
    if (frmPostDefinition.txtRecSelCurrentID.value > 0)
    {
        combo_disable(frmAbsenceDefinition.cboOrderBy1, true);
        combo_disable(frmAbsenceDefinition.cboOrderBy2, true);
		
        checkbox_disable(frmAbsenceDefinition.chkOrderBy1Asc, true);
        checkbox_disable(frmAbsenceDefinition.chkOrderBy2Asc, true);
		
        checkbox_disable(frmAbsenceDefinition.chkMinimumBradfordFactor, true);
        frmAbsenceDefinition.chkMinimumBradfordFactor.checked = false;
        text_disable(frmAbsenceDefinition.txtMinimumBradfordFactor, true);
        frmAbsenceDefinition.txtMinimumBradfordFactor.value = 0;
    }
    else
    {
        combo_disable(frmAbsenceDefinition.cboOrderBy1, false);
        combo_disable(frmAbsenceDefinition.cboOrderBy2, false);

        if (!frmAbsenceDefinition.chkMinimumBradfordFactor.checked)
        {
            frmAbsenceDefinition.txtMinimumBradfordFactor.value = 0;
            text_disable(frmAbsenceDefinition.txtMinimumBradfordFactor, true);
        }
        else
        {
            text_disable(frmAbsenceDefinition.txtMinimumBradfordFactor, false);
        }
			
        if (frmAbsenceDefinition.cboOrderBy1.options[frmAbsenceDefinition.cboOrderBy1.selectedIndex].value > 0)
        {
            checkbox_disable(frmAbsenceDefinition.chkOrderBy1Asc, false);
        }
        else
        {
            frmAbsenceDefinition.chkOrderBy1Asc.checked = false;
            checkbox_disable(frmAbsenceDefinition.chkOrderBy1Asc, true);
        }

        if (frmAbsenceDefinition.cboOrderBy2.options[frmAbsenceDefinition.cboOrderBy2.selectedIndex].value > 0)
        {
            checkbox_disable(frmAbsenceDefinition.chkOrderBy2Asc, false);
        }
        else
        {
            frmAbsenceDefinition.chkOrderBy2Asc.checked = false;
            checkbox_disable(frmAbsenceDefinition.chkOrderBy2Asc, true);
        }
			
    }
	
    if (!frmAbsenceDefinition.chkShowAbsenceDetails.checked) 
    {
        frmAbsenceDefinition.chkSRV.checked = false;
        checkbox_disable(frmAbsenceDefinition.chkSRV, true);
    }
    else
    {
        checkbox_disable(frmAbsenceDefinition.chkSRV, false);
    }
}

function refreshTab3Controls()
{
    
    var fViewing = (frmAbsenceUseful.txtAction.value.toUpperCase() == "VIEW");

    with (frmAbsenceDefinition)
    {
        if (optOutputFormat0.checked == true)		//Data Only
        {
            //disable preview opitons
            chkPreview.checked = false;
            checkbox_disable(chkPreview, true);
			
            //enable display on screen options
            checkbox_disable(chkDestination0, (fViewing == true));
			
            //enable-disable printer options
            checkbox_disable(chkDestination1, (fViewing == true));
            if (chkDestination1.checked == true)
            {
                populatePrinters();
                combo_disable(cboPrinterName, (fViewing == true));
            }
            else
            {
                cboPrinterName.length = 0;
                combo_disable(cboPrinterName, true);
            }
			
            //disable save options
            chkDestination2.checked = false
            checkbox_disable(chkDestination2, true);
            combo_disable(cboSaveExisting, true);
            cboSaveExisting.length = 0;
            txtFilename.value = '';
            text_disable(txtFilename, true);
            button_disable(cmdFilename, true);
			
            //disable email options
            chkDestination3.checked = false;
            checkbox_disable(chkDestination3, true);
            text_disable(txtEmailGroup, true);
            txtEmailGroup.value = '';
            txtEmailGroupID.value = 0;
            button_disable(cmdEmailGroup, true);
            text_disable(txtEmailSubject, true);
            text_disable(txtEmailAttachAs, true);

        }
        else if (optOutputFormat1.checked == true)   //CSV File
        {
            //enable preview opitons
            checkbox_disable(chkPreview, (fViewing == true));
			
            //disable display on screen options
            chkDestination0.checked = false;
            checkbox_disable(chkDestination0, (fViewing == true));
			
            //disable printer options
            chkDestination1.checked = false;
            checkbox_disable(chkDestination1, true);
            cboPrinterName.length = 0;
            combo_disable(cboPrinterName, true);
						
            //enable-disable save options
            checkbox_disable(chkDestination2, false);
            if (chkDestination2.checked == true)
            {
                populateSaveExisting();
                combo_disable(cboSaveExisting, false);
                text_disable(txtFilename, false);
                button_disable(cmdFilename, false);
            }	
            else
            {
                cboSaveExisting.length = 0;
                combo_disable(cboSaveExisting, true);
                text_disable(txtFilename, true);
                txtFilename.value = '';
                button_disable(cmdFilename, true);
            }
			
            //enable-disable email options
            checkbox_disable(chkDestination3, false);
            if (chkDestination3.checked == true)
            {
                text_disable(txtEmailGroup, false);
                text_disable(txtEmailSubject, false);
                button_disable(cmdEmailGroup, false);
                text_disable(txtEmailAttachAs, false);
            }
            else
            {
                text_disable(txtEmailGroup, true);
                txtEmailGroup.value = '';
                txtEmailGroupID.value = 0;
                button_disable(cmdEmailGroup, true);
                text_disable(txtEmailSubject, true);
                text_disable(txtEmailAttachAs, true);
            }
        }
        else if (optOutputFormat2.checked == true)		//HTML Document
        {
            //enable preview opitons
            checkbox_disable(chkPreview, (fViewing == true));
			
            //disable display on screen options
            checkbox_disable(chkDestination0, (fViewing == true));
			
            //disable printer options
            chkDestination1.checked = false;
            checkbox_disable(chkDestination1, true);
            cboPrinterName.length = 0;
            combo_disable(cboPrinterName, true);
						
            //enable-disable save options
            checkbox_disable(chkDestination2, false);
            if (chkDestination2.checked == true)
            {
                populateSaveExisting();
                combo_disable(cboSaveExisting, false);
                text_disable(txtFilename, false);
                button_disable(cmdFilename, false);
            }	
            else
            {
                cboSaveExisting.length = 0;
                combo_disable(cboSaveExisting, true);
                text_disable(txtFilename, true);
                txtFilename.value = '';
                button_disable(cmdFilename, true);
            }

            //enable-disable email options
            checkbox_disable(chkDestination3, false);
            if (chkDestination3.checked == true)
            {
                text_disable(txtEmailGroup, false);
                text_disable(txtEmailSubject, false);
                button_disable(cmdEmailGroup, false);
                text_disable(txtEmailAttachAs, false);
            }
            else
            {
                text_disable(txtEmailGroup, true);
                txtEmailGroup.value = '';
                txtEmailGroupID.value = 0;
                button_disable(cmdEmailGroup, true);
                text_disable(txtEmailSubject, true);
                text_disable(txtEmailAttachAs, true);
            }
        }
        else if (optOutputFormat3.checked == true)		//Word Document
        {
            //enable preview opitons
            checkbox_disable(chkPreview, (fViewing == true));
			
            //enable display on screen options
            checkbox_disable(chkDestination0, (fViewing == true));
			
            //enable-disable printer options
            checkbox_disable(chkDestination1, (fViewing == true));	
            if (chkDestination1.checked == true)
            {
                populatePrinters();
                combo_disable(cboPrinterName,  (fViewing == true));
            }
            else
            {
                cboPrinterName.length = 0;
                combo_disable(cboPrinterName,  true);
            }
										
            //enable-disable save options
            checkbox_disable(chkDestination2, false);
            if (chkDestination2.checked == true)
            {
                populateSaveExisting();
                combo_disable(cboSaveExisting,  false);
                text_disable(txtFilename, false);
                button_disable(cmdFilename, false);
            }	
            else
            {
                cboSaveExisting.length = 0;
                combo_disable(cboSaveExisting,  true);
                text_disable(txtFilename, true);
                txtFilename.value = '';
                button_disable(cmdFilename, true);
            }

            //enable-disable email options
            checkbox_disable(chkDestination3, false);
            if (chkDestination3.checked == true)
            {
                text_disable(txtEmailGroup, false);
                text_disable(txtEmailSubject, false);
                button_disable(cmdEmailGroup, false);
                text_disable(txtEmailAttachAs, false);
            }
            else
            {
                text_disable(txtEmailGroup, true);
                txtEmailGroup.value = '';
                txtEmailGroupID.value = 0;
                button_disable(cmdEmailGroup, true);
                text_disable(txtEmailSubject, true);
                text_disable(txtEmailAttachAs, true);
            }
        }
        else if ((optOutputFormat4.checked == true) ||		//Excel Worksheet
                 (optOutputFormat5.checked == true) ||
                 (optOutputFormat6.checked == true))
        {
            //enable preview opitons
            checkbox_disable(chkPreview, (fViewing == true));
			
            //enable display on screen options
            checkbox_disable(chkDestination0, (fViewing == true));
			
            //enable-disable printer options
            checkbox_disable(chkDestination1, (fViewing == true));
            if (chkDestination1.checked == true)
            {
                populatePrinters();
                combo_disable(cboPrinterName, (fViewing == true));
            }
            else
            {
                cboPrinterName.length = 0;
                combo_disable(cboPrinterName, true);
            }
										
            //enable-disable save options
            checkbox_disable(chkDestination2, false);
            if (chkDestination2.checked == true)
            {
                populateSaveExisting();
                combo_disable(cboSaveExisting, false);
                text_disable(txtFilename, false);
                button_disable(cmdFilename, false);
            }	
            else
            {
                cboSaveExisting.length = 0;
                combo_disable(cboSaveExisting, true);
                text_disable(txtFilename, true);
                txtFilename.value = '';
                button_disable(cmdFilename, true);
            }

            //enable-disable email options
            checkbox_disable(chkDestination3, false);
            if (chkDestination3.checked == true)
            {
                text_disable(txtEmailGroup, false);
                text_disable(txtEmailSubject, false);
                button_disable(cmdEmailGroup, false);
                text_disable(txtEmailAttachAs, false);
            }
            else
            {
                text_disable(txtEmailGroup, true);
                txtEmailGroup.value = '';
                txtEmailGroupID.value = 0;
                button_disable(cmdEmailGroup, true);
                text_disable(txtEmailSubject, true);
                text_disable(txtEmailAttachAs, true);
            }
        }
            /*else if (optOutputFormat5.checked == true)		//Excel Chart
                {
                }
            else if (optOutputFormat6.checked == true)		//Excel Pivot Table
                {
                }*/
        else
        {
            optOutputFormat0.checked = true;
            chkDestination0.checked=true;
            refreshTab3Controls();
        }
		
        if (txtEmailSubject.disabled)
        {
            txtEmailSubject.value = '';
        }

        if (txtEmailAttachAs.disabled)
        {
            txtEmailAttachAs.value = '';
        }
        else
        {
            if (txtEmailAttachAs.value == '') {
                if (txtFilename.value != '') {
                    sAttachmentName = new String(txtFilename.value);
                    txtEmailAttachAs.value = sAttachmentName.substr(sAttachmentName.lastIndexOf("\\")+1);
                }
            }
        }

        if (cmdFilename.disabled == true) {
            txtFilename.value = "";
        }

    }

    // Little dodge to get around a browser bug that
    // does not refresh the display on all controls.
    try
    {
        window.resizeBy(0,-1);
        window.resizeBy(0,1);
    }
    catch(e) {}
}

function saveFile()
{

    dialog.CancelError = true;
    dialog.DialogTitle = "Output Document";
    dialog.Flags = 2621444;

    if (frmAbsenceDefinition.optOutputFormat1.checked == true) {
        //CSV
        dialog.Filter = "Comma Separated Values (*.csv)|*.csv";
    }

    else if (frmAbsenceDefinition.optOutputFormat2.checked == true) {
        //HTML
        dialog.Filter = "HTML Document (*.htm)|*.htm";
    }

    else if (frmAbsenceDefinition.optOutputFormat3.checked == true) {
        //WORD
        //dialog.Filter = "Word Document (*.doc)|*.doc";
        dialog.Filter = frmAbsenceDefinition.txtWordFormats.value;
        dialog.FilterIndex = frmAbsenceDefinition.txtWordFormatDefaultIndex.value;
    }

    else {
        //EXCEL
        //dialog.Filter = "Excel Workbook (*.xls)|*.xls";
        dialog.Filter = frmAbsenceDefinition.txtExcelFormats.value;
        dialog.FilterIndex = frmAbsenceDefinition.txtExcelFormatDefaultIndex.value;
    }



    if (frmAbsenceDefinition.txtFilename.value.length == 0) {
        sKey = new String("documentspath_");
        sKey = sKey.concat(frmAbsenceDefinition.txtDatabase.value);
        sPath = window.parent.frames("menuframe").ASRIntranetFunctions.GetRegistrySetting("HR Pro", "DataPaths", sKey);
        dialog.InitDir = sPath;
    }
    else {
        dialog.FileName = frmAbsenceDefinition.txtFilename.value;
    }


    try {
        dialog.ShowSave();

        if (dialog.FileName.length > 256) {
            OpenHR.messageBox("Path and file name must not exceed 256 characters in length");
            return;
        }

        frmAbsenceDefinition.txtFilename.value = dialog.FileName;

    }
    catch(e) {
    }

}

function populateSaveExisting()
{
    with (frmAbsenceDefinition.cboSaveExisting)
    {
        lngCurrentOption = 0;
        if (selectedIndex > 0) {
            lngCurrentOption = options[selectedIndex].value;
        }
        length = 0;

        var oOption = document.createElement("OPTION");
        options.add(oOption);
        oOption.innerText = "Overwrite";
        oOption.value = 0;
		
        var oOption = document.createElement("OPTION");
        options.add(oOption);
        oOption.innerText = "Do not overwrite";
        oOption.value = 1;
		
        var oOption = document.createElement("OPTION");
        options.add(oOption);
        oOption.innerText = "Add sequential number to name";
        oOption.value = 2;
		
        var oOption = document.createElement("OPTION");
        options.add(oOption);
        oOption.innerText = "Append to file";
        oOption.value = 3;
		
        if ((frmAbsenceDefinition.optOutputFormat4.checked) ||
            (frmAbsenceDefinition.optOutputFormat5.checked) ||
            (frmAbsenceDefinition.optOutputFormat6.checked))
        {
            var oOption = document.createElement("OPTION");
            options.add(oOption);
            oOption.innerText = "Create new sheet in workbook";
            oOption.value = 4;
        }

        for (iLoop=0; iLoop<options.length; iLoop++)  {
            if (options(iLoop).value == lngCurrentOption) {
                selectedIndex = iLoop
                break;
            }
        }

    }
}

function absencedef_convertLocaleDateToSQL(psDateString)
{ 
    /* Convert the given date string (in locale format) into 
    SQL format (mm/dd/yyyy). */
    var sDateFormat;
    var iDays;
    var iMonths;
    var iYears;
    var sDays;
    var sMonths;
    var sYears;
    var iValuePos;
    var sTempValue;
    var sValue;
    var iLoop;

    sDateFormat = OpenHR.LocaleDateFormat();

    sDays="";
    sMonths="";
    sYears="";
    iValuePos = 0;

    // Trim leading spaces.
    sTempValue = psDateString.substr(iValuePos,1);
    while (sTempValue.charAt(0) == " ") 
    {
        iValuePos = iValuePos + 1;		
        sTempValue = psDateString.substr(iValuePos,1);
    }

    for (iLoop=0; iLoop<sDateFormat.length; iLoop++)  {
        if ((sDateFormat.substr(iLoop,1).toUpperCase() == 'D') && (sDays.length==0)){
            sDays = psDateString.substr(iValuePos,1);
            iValuePos = iValuePos + 1;
            sTempValue = psDateString.substr(iValuePos,1);

            if (isNaN(sTempValue) == false) {
                sDays = sDays.concat(sTempValue);			
            }
            iValuePos = iValuePos + 1;		
        }

        if ((sDateFormat.substr(iLoop,1).toUpperCase() == 'M') && (sMonths.length==0)){
            sMonths = psDateString.substr(iValuePos,1);
            iValuePos = iValuePos + 1;
            sTempValue = psDateString.substr(iValuePos,1);

            if (isNaN(sTempValue) == false) {
                sMonths = sMonths.concat(sTempValue);			
            }
            iValuePos = iValuePos + 1;
        }

        if ((sDateFormat.substr(iLoop,1).toUpperCase() == 'Y') && (sYears.length==0)){
            sYears = psDateString.substr(iValuePos,1);
            iValuePos = iValuePos + 1;
            sTempValue = psDateString.substr(iValuePos,1);

            if (isNaN(sTempValue) == false) {
                sYears = sYears.concat(sTempValue);			
            }
            iValuePos = iValuePos + 1;
            sTempValue = psDateString.substr(iValuePos,1);

            if (isNaN(sTempValue) == false) {
                sYears = sYears.concat(sTempValue);			
            }
            iValuePos = iValuePos + 1;
            sTempValue = psDateString.substr(iValuePos,1);

            if (isNaN(sTempValue) == false) {
                sYears = sYears.concat(sTempValue);			
            }
            iValuePos = iValuePos + 1;
        }

        // Skip non-numerics
        sTempValue = psDateString.substr(iValuePos,1);
        while (isNaN(sTempValue) == true) {
            iValuePos = iValuePos + 1;		
            sTempValue = psDateString.substr(iValuePos,1);
        }
    }

    while (sDays.length < 2) {
        sTempValue = "0";
        sDays = sTempValue.concat(sDays);
    }

    while (sMonths.length < 2) {
        sTempValue = "0";
        sMonths = sTempValue.concat(sMonths);
    }

    while (sYears.length < 2) {
        sTempValue = "0";
        sYears = sTempValue.concat(sYears);
    }

    if (sYears.length == 2) {
        iValue = parseInt(sYears);
        if (iValue < 30) {
            sTempValue = "20";
        }
        else {
            sTempValue = "19";
        }
		
        sYears = sTempValue.concat(sYears);
    }

    while (sYears.length < 4) {
        sTempValue = "0";
        sYears = sTempValue.concat(sYears);
    }

    sTempValue = sMonths.concat("/");
    sTempValue = sTempValue.concat(sDays);
    sTempValue = sTempValue.concat("/");
    sTempValue = sTempValue.concat(sYears);
	
    sValue = OpenHR.ConvertSQLDateToLocale(sTempValue);

    iYears = parseInt(sYears);
	
    while (sMonths.substr(0, 1) == "0") {
        sMonths = sMonths.substr(1);
    }
    iMonths = parseInt(sMonths);
	
    while (sDays.substr(0, 1) == "0") {
        sDays = sDays.substr(1);
    }
    iDays = parseInt(sDays);

    var newDateObj = new Date(iYears, iMonths - 1, iDays);
    if ((newDateObj.getDate() != iDays) || 
        (newDateObj.getMonth() + 1 != iMonths) || 
        (newDateObj.getFullYear() != iYears)) {
        return "";
    }
    else {
        return sTempValue;
    }
}

