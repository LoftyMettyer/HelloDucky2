<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@Import namespace="DMI.NET" %>

<%="" %>

<script src="<%: Url.Content("~/scripts/date.js")%>" type="text/javascript"></script>           

<%
    Dim sKey As String
    
    ' Clear the session action which is used to botch the prompted values screen in
	session("action") = ""
	session("optionaction") = ""

	' Read the prompted start/end dates if there were any
	dim aPrompts(1)

    For i = 0 To (Request.Form.Count) - 1
        sKey = Request.Form.Keys(i)
        If ((UCase(Left(sKey, 7)) = "PROMPT_") And (Mid(sKey, 8, 1) <> "3")) Or _
            (UCase(Left(sKey, 10)) = "PROMPTCHK_") Then
			
            If Mid(sKey, 8, 5) = "start" Then
                aPrompts(0) = Request.Form.Item(i)
            Else
                aPrompts(1) = Request.Form.Item(i)
            End If

        End If
    Next
%>

<object classid="clsid:F9043C85-F6F2-101A-A3C9-08002B2F49FB"
    id="dialog"
    codebase="cabs/comdlg32.cab#Version=1,0,0,0"
    style="LEFT: 0px; TOP: 0px">
    <param name="_ExtentX" value="847">
    <param name="_ExtentY" value="847">
    <param name="_Version" value="393216">
    <param name="CancelError" value="0">
    <param name="Color" value="0">
    <param name="Copies" value="1">
    <param name="DefaultExt" value="">
    <param name="DialogTitle" value="">
    <param name="FileName" value="">
    <param name="Filter" value="">
    <param name="FilterIndex" value="0">
    <param name="Flags" value="0">
    <param name="FontBold" value="0">
    <param name="FontItalic" value="0">
    <param name="FontName" value="">
    <param name="FontSize" value="8">
    <param name="FontStrikeThru" value="0">
    <param name="FontUnderLine" value="0">
    <param name="FromPage" value="0">
    <param name="HelpCommand" value="0">
    <param name="HelpContext" value="0">
    <param name="HelpFile" value="">
    <param name="HelpKey" value="">
    <param name="InitDir" value="">
    <param name="Max" value="0">
    <param name="Min" value="0">
    <param name="MaxFileSize" value="260">
    <param name="PrinterDefault" value="1">
    <param name="ToPage" value="0">
    <param name="Orientation" value="1">
</object>

    <script type="text/javascript">

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

    </script>

<%
    ' Settings objects
    Dim objSettings As New HR.Intranet.Server.clsSettings    
    objSettings.Connection = Session("databaseConnection")

    Dim aColumnNames
    Dim aAbsenceTypes
    Dim cmdReportsCols
    Dim prmBaseTableID
    Dim rstReportColumns
    Dim sErrorDescription As String
    Dim iCount As Integer
    
	' Load Absence Types and Personnel columns into array
	redim aColumnNames(1,0)
	redim aAbsenceTypes(0)

	' Get the table records.
    cmdReportsCols = Server.CreateObject("ADODB.Command")
	cmdReportsCols.CommandText = "sp_ASRIntGetColumns"
	cmdReportsCols.CommandType = 4 ' Stored procedure
    cmdReportsCols.ActiveConnection = Session("databaseConnection")
																				
    prmBaseTableID = cmdReportsCols.CreateParameter("piTableID", 3, 1) ' 3=integer, 1=input
    cmdReportsCols.Parameters.Append(prmBaseTableID)
	prmBaseTableID.value = cleanNumeric(session("Personnel_EmpTableID"))

    Err.Clear()
    rstReportColumns = cmdReportsCols.Execute
																																			
    If (Err.Number <> 0) Then
        sErrorDescription = "The personnel column information could not be retrieved." & vbCrLf & FormatError(Err.Description)
    End If

	if len(sErrorDescription) = 0 then
		iCount = 0
		do while not rstReportColumns.EOF
		
			if rstReportColumns.fields("OLEType").value <> 2 then
		
				aColumnNames(0,iCount) = rstReportColumns.fields("ColumnID").value
				aColumnNames(1,iCount) = rstReportColumns.fields("ColumnName").value
			
				redim preserve aColumnNames(1, ubound(aColumnNames,2) + 1)
				iCount = iCount + 1
			end if
			
			rstReportColumns.MoveNext
			
		loop

		rstReportColumns.close

	end if

    ' Load absence types
    Dim cmdTables
    Dim rstTablesInfo
    
    cmdTables = Server.CreateObject("ADODB.Command")
	cmdTables.CommandText = "sp_ASRIntGetAbsenceTypes"
	cmdTables.CommandType = 4 ' Stored Procedure
    cmdTables.ActiveConnection = Session("databaseConnection")

    rstTablesInfo = cmdTables.Execute
																			
    If (Err.Number <> 0) Then
        sErrorDescription = "The absence type information could not be retrieved." & vbCrLf & FormatError(Err.Description)
    End If

	if len(sErrorDescription) = 0 then
		iCount = 0
		do while not rstTablesInfo.EOF
			aAbsenceTypes(iCount) = rstTablesInfo.fields("Type").value
			redim preserve aAbsenceTypes(ubound(aAbsenceTypes) + 1)
			iCount = iCount + 1
			rstTablesInfo.MoveNext
		loop

		rstTablesInfo.close

	end if
												
	' Release the ADO objects.
    cmdTables = Nothing
    cmdReportsCols = Nothing


	' Set the default settings
    Dim strReportType As String = "AbsenceBreakdown"
    Dim strDate
    Dim strType
    Dim lngDefaultColumnID As Long
    Dim lngConfigColumnID As Long
    Dim strSaveExisting As String
	
    Response.Write("<script type=""text/javascript"">" & vbCrLf)

    Response.Write("function SetReportDefaults(){" & vbCrLf)
    Response.Write("   var frmAbsenceDefinition = OpenHR.getForm(""workframe"",""frmAbsenceDefinition"");" & vbCrLf)
    
    ' Type of standard report being run
    If Session("StandardReport_Type") = 16 Then
        strReportType = "BradfordFactor"
    End If

    If Session("StandardReport_Type") = 15 Then
        strReportType = "AbsenceBreakdown"
        Response.Write("frmAbsenceDefinition.btnTab2.style.visibility = ""hidden"";" & vbCrLf)
    End If

    ' Absence types
    For iCount = 0 To UBound(aAbsenceTypes) - 1
        If objSettings.GetSystemSetting(strReportType, "Absence Type " & aAbsenceTypes(iCount), "0") = "1" Then
            Response.Write("frmAbsenceDefinition.chkAbsenceType_" & iCount & ".checked = 1;" & vbCrLf)
        End If
    Next

                            ' Date range
                            If Len(aPrompts(0)) = 0 Then
                                strDate = ConvertSQLDateToLocale(objSettings.GetStandardReportDate(strReportType, "Start Date"))
                            Else
                                strDate = aPrompts(0)
                            End If
                            Response.Write("frmAbsenceDefinition.txtDateFrom.value = " & """" & CleanStringForJavaScript(strDate) & """" & ";" & vbCrLf)

                            ' Date range
                            If Len(aPrompts(1)) = 0 Then
                                strDate = ConvertSQLDateToLocale(objSettings.GetStandardReportDate(strReportType, "End Date"))
                            Else
                                strDate = aPrompts(1)
                            End If
                            Response.Write("frmAbsenceDefinition.txtDateTo.value = " & """" & CleanStringForJavaScript(strDate) & """" & ";" & vbCrLf)

                            ' Record Selection
                            If Session("optionRecordID") = 0 Then

                                strType = objSettings.GetSystemSetting(strReportType, "Type", "A")
		
                                Select Case strType
                                    Case "A"
                                        Response.Write("frmAbsenceDefinition.optAllRecords.checked = 1;" & vbCrLf)
                                        Response.Write("frmAbsenceDefinition.optPickList.checked = 0;" & vbCrLf)
                                        Response.Write("frmAbsenceDefinition.optFilter.checked = 0;" & vbCrLf)
                                    Case "P"
                                        Response.Write("frmAbsenceDefinition.optAllRecords.checked = 0;" & vbCrLf)
                                        Response.Write("frmAbsenceDefinition.optPickList.checked = 1;" & vbCrLf)
                                        Response.Write("frmAbsenceDefinition.optFilter.checked = 0;" & vbCrLf)
                                        Response.Write("frmAbsenceDefinition.txtBasePicklist.value = " & """" & CleanStringForJavaScript(objSettings.GetPicklistFilterName(strReportType, strType)) & """" & ";" & vbCrLf)
                                        Response.Write("button_disable(frmAbsenceDefinition.cmdBasePicklist, false);" & vbCrLf)
                                        Response.Write("frmPostDefinition.txtBasePicklistID.value = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "ID", "0")) & ";" & vbCrLf)
                                    Case "F"
                                        Response.Write("frmAbsenceDefinition.optAllRecords.checked = 0;" & vbCrLf)
                                        Response.Write("frmAbsenceDefinition.optPickList.checked = 0;" & vbCrLf)
                                        Response.Write("frmAbsenceDefinition.optFilter.checked = 1;" & vbCrLf)
                                        Response.Write("frmAbsenceDefinition.txtBaseFilter.value = " & """" & CleanStringForJavaScript(objSettings.GetPicklistFilterName(strReportType, strType)) & """" & ";" & vbCrLf)
                                        Response.Write("button_disable(frmAbsenceDefinition.cmdBaseFilter, false);" & vbCrLf)
                                        Response.Write("frmPostDefinition.txtBaseFilterID.value = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "ID", "0")) & ";" & vbCrLf)
                                End Select
                            Else
                                Response.Write("RecordSelection.style.visibility = ""hidden"";" & vbCrLf)
                            End If

                            ' Display picklist in header
                            Response.Write("frmAbsenceDefinition.chkPrintInReportHeader.checked =  " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "PrintFilterHeader", "0")) & vbCrLf)

                            ' Bradford Factor specific stuff
                            If Session("StandardReport_Type") = 16 Then
                                Response.Write("frmAbsenceDefinition.chkSRV.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "SRV", "0")) & ";" & vbCrLf)
                                Response.Write("frmAbsenceDefinition.chkShowDurations.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Show Totals", "1")) & ";" & vbCrLf)
                                Response.Write("frmAbsenceDefinition.chkShowInstances.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Show Count", "0")) & ";" & vbCrLf)
                                Response.Write("frmAbsenceDefinition.chkShowFormula.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Show Workings", "0")) & ";" & vbCrLf)
                                Response.Write("frmAbsenceDefinition.chkShowAbsenceDetails.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Display Absence Details", "1")) & ";" & vbCrLf)
                                Response.Write("frmAbsenceDefinition.chkOmitBeforeStart.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Omit Before", "0")) & ";" & vbCrLf)
                                Response.Write("frmAbsenceDefinition.chkOmitAfterEnd.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Omit After", "0")) & ";" & vbCrLf)
                                Response.Write("frmAbsenceDefinition.chkMinimumBradfordFactor.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Minimum Bradford Factor", "0")) & ";" & vbCrLf)
                                Response.Write("frmAbsenceDefinition.txtMinimumBradfordFactor.value = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Minimum Bradford Factor Amount", "0")) & ";" & vbCrLf)
                                Response.Write("frmAbsenceDefinition.chkOrderBy1Asc.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Order By Asc", "1")) & ";" & vbCrLf)
                                Response.Write("frmAbsenceDefinition.chkOrderBy2Asc.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Group By Asc", "1")) & ";" & vbCrLf)
       
                                lngDefaultColumnID = objSettings.GetModuleParameter("MODULE_PERSONNEL", "Param_FieldsSurname")
                                lngConfigColumnID = objSettings.GetSystemSetting(strReportType, "Order By", lngDefaultColumnID)
                                'Response.Write "frmAbsenceDefinition.cboOrderBy1.value = " & """" & sFieldName & """" & ";" & vbcrlf
                                Response.Write("for (var i=0; i<frmAbsenceDefinition.cboOrderBy1.options.length; i++)" & vbCrLf)
                                Response.Write("	{" & vbCrLf)
                                Response.Write("	if (frmAbsenceDefinition.cboOrderBy1.options[i].value == " & lngConfigColumnID & ")" & vbCrLf)
                                Response.Write("		{" & vbCrLf)
                                Response.Write("		frmAbsenceDefinition.cboOrderBy1.selectedIndex = i; " & vbCrLf)
                                Response.Write("		}" & vbCrLf)
                                Response.Write("	}" & vbCrLf)
		
                                lngDefaultColumnID = objSettings.GetModuleParameter("MODULE_PERSONNEL", "Param_FieldsForename")
                                lngConfigColumnID = objSettings.GetSystemSetting(strReportType, "Group By", lngDefaultColumnID)
                                'Response.Write "frmAbsenceDefinition.cboOrderBy2.value = " & """" & sFieldName & """" & ";" & vbcrlf
                                Response.Write("for (var i=0; i<frmAbsenceDefinition.cboOrderBy2.options.length; i++)" & vbCrLf)
                                Response.Write("	{" & vbCrLf)
                                Response.Write("	if (frmAbsenceDefinition.cboOrderBy2.options[i].value == " & lngConfigColumnID & ")" & vbCrLf)
                                Response.Write("		{" & vbCrLf)
                                Response.Write("		frmAbsenceDefinition.cboOrderBy2.selectedIndex = i; " & vbCrLf)
                                Response.Write("		}" & vbCrLf)
                                Response.Write("	}" & vbCrLf)
                            End If

                            ' Output Options
                            Select Case objSettings.GetSystemSetting(strReportType, "Format", 0)
                                Case "0"
                                    Response.Write("frmAbsenceDefinition.optOutputFormat0.checked = 1;" & vbCrLf)
                                Case "1"
                                    Response.Write("frmAbsenceDefinition.optOutputFormat1.checked = 1;" & vbCrLf)
                                Case "2"
                                    Response.Write("frmAbsenceDefinition.optOutputFormat2.checked = 1;" & vbCrLf)
                                Case "3"
                                    Response.Write("frmAbsenceDefinition.optOutputFormat3.checked = 1;" & vbCrLf)
                                Case "4"
                                    Response.Write("frmAbsenceDefinition.optOutputFormat4.checked = 1;" & vbCrLf)
                                Case "5"
                                    Response.Write("frmAbsenceDefinition.optOutputFormat5.checked = 1;" & vbCrLf)
                                Case "6"
                                    'MH20031211 Fault 7787
                                    'If Bradford then disallow Pivot (make it worksheet instead)
                                    If Session("StandardReport_Type") = 16 Then
                                        Response.Write("frmAbsenceDefinition.optOutputFormat4.checked = 1;" & vbCrLf)
                                    Else
                                        Response.Write("frmAbsenceDefinition.optOutputFormat6.checked = 1;" & vbCrLf)
                                    End If
                                Case Else
                                    ' Charts and pivot not in Intranet yet
                                    Response.Write("frmAbsenceDefinition.optOutputFormat0.checked = 1" & vbCrLf)
                            End Select
	
                            Response.Write("frmAbsenceDefinition.chkPreview.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Preview", 0)) & ";" & vbCrLf)
                            Response.Write("frmAbsenceDefinition.chkDestination0.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Screen", 1)) & ";" & vbCrLf)

                            Response.Write("frmAbsenceDefinition.chkDestination1.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Printer", 0)) & ";" & vbCrLf)

                            Response.Write("strPrinterName = '" & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "PrinterName", "")) & "';" & vbCrLf)

                            'Set the printer as defined in Report Configuration in DAT.
                            'Response.Write "frmAbsenceDefinition.cboPrinterName.value = " & """" & objSettings.GetSystemSetting(strReportType, "PrinterName", "") & """" & ";" & vbcrlf
                            Response.Write("for (var i=0; i<frmAbsenceDefinition.cboPrinterName.options.length; i++)" & vbCrLf)
                            Response.Write("	{" & vbCrLf)
                            Response.Write("	if (frmAbsenceDefinition.cboPrinterName.options[i].innerText.toLowerCase() == strPrinterName.toLowerCase())" & vbCrLf)
                            Response.Write("		{" & vbCrLf)
                            Response.Write("		frmAbsenceDefinition.cboPrinterName.selectedIndex = i; " & vbCrLf)
                            Response.Write("		}" & vbCrLf)
                            Response.Write("	}" & vbCrLf)

                            'MH20040311
                            Response.Write("if (frmAbsenceDefinition.chkDestination1.checked == true) " & vbCrLf)
                            Response.Write("	{" & vbCrLf)
                            Response.Write("	if (strPrinterName != """") " & vbCrLf)
                            Response.Write("		{" & vbCrLf) '
                            Response.Write("		if (frmAbsenceDefinition.cboPrinterName.options[frmAbsenceDefinition.cboPrinterName.selectedIndex].innerText != strPrinterName) " & vbCrLf)
                            Response.Write("			{" & vbCrLf)
                            Response.Write("			window.parent.frames(""menuframe"").ASRIntranetFunctions.MessageBox(""This definition is set to output to printer ""+strPrinterName+"" which is not set up on your PC."");" & vbCrLf)
                            Response.Write("			var oOption = document.createElement(""OPTION"");" & vbCrLf)
                            Response.Write("			frmAbsenceDefinition.cboPrinterName.options.add(oOption);" & vbCrLf)
                            Response.Write("			oOption.innerText = strPrinterName;" & vbCrLf)
                            Response.Write("			oOption.value = frmAbsenceDefinition.cboPrinterName.options.length-1;" & vbCrLf)
                            Response.Write("			frmAbsenceDefinition.cboPrinterName.selectedIndex = oOption.value;" & vbCrLf)
                            Response.Write("			}" & vbCrLf)
                            Response.Write("		}" & vbCrLf)
                            Response.Write("	}" & vbCrLf)
	
                            Response.Write("frmAbsenceDefinition.chkDestination2.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Save", 0)) & ";" & vbCrLf)
                            Response.Write("frmAbsenceDefinition.txtFilename.value = " & """" & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "FileName", "")) & """" & ";" & vbCrLf)

                            Response.Write("populateSaveExisting();" & vbCrLf)
                            Select Case objSettings.GetSystemSetting(strReportType, "SaveExisting", 0)
                                Case 0
                                    strSaveExisting = "Overwrite"
                                    Response.Write("frmAbsenceDefinition.cboSaveExisting.selectedIndex = 0;" & vbCrLf)
                                Case 1
                                    strSaveExisting = "Do not overwrite"
                                    Response.Write("frmAbsenceDefinition.cboSaveExisting.selectedIndex = 1;" & vbCrLf)
                                Case 2
                                    strSaveExisting = "Add sequential number to name"
                                    Response.Write("frmAbsenceDefinition.cboSaveExisting.selectedIndex = 2;" & vbCrLf)
                                Case 3
                                    strSaveExisting = "Append to file"
                                    Response.Write("frmAbsenceDefinition.cboSaveExisting.selectedIndex = 3;" & vbCrLf)
                                Case 4
                                    strSaveExisting = "Create new sheet in workbook"
                                    Response.Write("frmAbsenceDefinition.cboSaveExisting.selectedIndex = 4;" & vbCrLf)
                            End Select

                            'Response.Write "frmAbsenceDefinition.cboSaveExisting.value = " & """" & strSaveExisting & """"  & ";" & vbcrlf
                            Response.Write("frmAbsenceDefinition.chkDestination3.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Email", 0)) & ";" & vbCrLf)
                            Response.Write("frmAbsenceDefinition.txtEmailGroup.value = " & """" & CleanStringForJavaScript(objSettings.GetEmailGroupName(objSettings.GetSystemSetting(strReportType, "EmailAddr", "0"))) & """" & ";" & vbCrLf)
                            Response.Write("frmAbsenceDefinition.txtEmailGroupID.value = " & """" & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "EmailAddr", "")) & """" & ";" & vbCrLf)
                            Response.Write("frmAbsenceDefinition.txtEmailAttachAs.value = " & """" & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "EmailAttachAs", "")) & """" & ";" & vbCrLf)
                            Response.Write("frmAbsenceDefinition.txtEmailSubject.value = " & """" & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "EmailSubject", "")) & """" & ";" & vbCrLf)

    Response.Write(vbCrLf & "}")
    Response.Write("</script>" & vbCrLf)

    objSettings = Nothing

%>
												
<form id=frmAbsenceDefinition name=frmAbsenceDefinition>
<table align=center class="outline" cellPadding=5 cellSpacing=0 width="700" height="60%">
	<TR>
		<TD>
			<TABLE WIDTH="100%" height="100%" class="invisible" cellspacing=0 cellpadding=0>
				<tr height=5> 
					<td colspan=3></td>
				</tr> 

				<tr height=10>
					<TD width=10></TD>
					<td>
						<INPUT type="button" class="btn btndisabled" value="Definition" id=btnTab1 name=btnTab1 disabled="disabled"
						    onclick="displayPage(1)" 
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
<%
    if session("StandardReport_Type") = 16 then
%>    
	                    <INPUT type="button" class="btn" value="Options" id=btnTab2 name=btnTab2 
	                        onclick="displayPage(2)"
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
<%
    end if
%>
						<INPUT type="button" class="btn" value="Output" id=btnTab3 name=btnTab3 
						    onclick="displayPage(3)"
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
<%
	' Causes problems if button isn't there
	if session("StandardReport_Type") <> 16 then
%>
		                <INPUT type="button" class="btn" value="Options" id=btnTab2 name=btnTab2 
		                    onclick="displayPage(2)"
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
<%
	end if
%>

					</td>
					<TD width=10></TD>
				</tr> 
				
				<tr height=10> 
					<td colspan=3></td>
				</tr> 

				<tr> 
					<TD width=10></TD>
					<td>
						<!-- First tab -->
						<DIV id=div1>
							<TABLE WIDTH="100%" height="100%" class="outline" cellspacing=0 cellpadding=5>
								<tr valign=top> 
									<td valign=top rowspan=2 width="25%" height="100%">
										<table class="invisible" cellspacing="0" cellpadding="4" width="100%" height="100%">
											<tr height=10> 
												<td height=10 align=left valign=top>
													Absence Types : <BR><BR>
													<SPAN id=AbsenceTypes style="width:300px;height:200px; overflow:auto;" class="outline">
														<TABLE class="invisible" cellspacing="0" cellpadding="0" width="100%" bgColor=white>
															<TR>
																<TD>
<%
    for iCount = 0 to ubound(aAbsenceTypes) - 1
%>																		
																			
																    <TR>
																        <TD>
																            <INPUT id="chkAbsenceType_<%=iCount%>" name=chkAbsenceType_<%=iCount%> type=checkbox tagname=<%=aAbsenceTypes(iCount)%> tabindex=-1 
		                                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                            <label 
				                                                                for="chkAbsenceType_<%=iCount%>"
				                                                                class="checkbox"
				                                                                tabindex=0 
				                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
		                                                                        
											    					            <%=aAbsenceTypes(iCount)%>
                                                		    		        </label>
																        </TD>
																    </TR>
<%
    next
%>
																</TD>
															</TR>																	
														</TABLE>
													</SPAN>
												</td>

												<td height=10 align=left valign=top>
													<TABLE cellSpacing=1 cellPadding=1 width="100%" class="outline">
														<TR>
															<TD>
																<TABLE border=0>
																	<TR>
																		<TD colspan=2>
																			Date Range :
																		</TD>
																	</TR>
																	<TR>
																		<TD width=100>
																			Start Date :
																		</TD>
																		<TD>
																			<INPUT id=txtDateFrom class="text" name=txtDateFrom onblur="validateDate(this);">
																		</TD>
																	</TR>
																	<TR>
																		<TD width=100>
																			End Date :
																		</TD>
																		<TD>
																			<INPUT id=txtDateTo class="text" name=txtDateTo onblur="validateDate(this);">
																		</TD>
																	</TR>
																</TABLE>
															</TD>
														</TR>
													</TABLE>

													&nbsp

													<SPAN id=RecordSelection>
														<TABLE cellSpacing=1 cellPadding=1 width="300" class="outline">
															<tr height=10>
																<td height=10 align=left valign=top>
																	Record Selection :
																	<TABLE class="invisible" cellspacing=0 cellpadding=3>
																		<TR>															
																			<TABLE WIDTH="325" height="80%" border=0 cellspacing=0 cellpadding=5>
																				<TD>
																					<TABLE WIDTH="360" class="invisible" CELLSPACING=0 CELLPADDING=0>
																						<TR>
																							<TD width=95 colspan=3>
																								<INPUT CHECKED id=optAllRecords name=optAllRecords type=radio 
																								    onclick="changeRecordOptions('all')"
		                                                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
                                                                                                <label 
                                                                                                    tabindex="-1"
	                                                                                                for="optAllRecords"
	                                                                                                class="radio"
		                                                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                                                        />
																								    All
																								</label>
																							</TD>
																						</TR>
																						<TR>
																							<TD width=95>
																								<INPUT id=optPickList name=optPickList type=radio 
																								    onclick="changeRecordOptions('picklist')"
		                                                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
                                                                                                <label 
                                                                                                    tabindex="-1"
	                                                                                                for="optPickList"
	                                                                                                class="radio"
		                                                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                                                        />
    																								Picklist
    																							</label>
																							</TD>
																							<TD>
																								<INPUT id=txtBasePicklist name=txtBasePicklist class="text textdisabled" disabled="disabled" style="width=250">
																							</TD>
																							<TD width=15>
																								<INPUT id=cmdBasePicklist name=cmdBasePicklist class="btn btndisabled" disabled="disabled" type=button value="..." 
																								    onclick="selectRecordOption('picklist')" 
			                                                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
			                                                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
			                                                                                        onfocus="try{button_onFocus(this);refreshControls();}catch(e){}"
			                                                                                        onblur="try{button_onBlur(this);}catch(e){}" />
																							</TD>
																						</TR>
																						<TR>
																							<TD>
																								<INPUT id=optFilter name=optFilter type=radio 
																								    onclick="changeRecordOptions('filter')"
		                                                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                                    onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                                    onblur="try{radio_onBlur(this);}catch(e){}"/>
                                                                                                <label 
                                                                                                    tabindex="-1"
	                                                                                                for="optFilter"
	                                                                                                class="radio"
		                                                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
		                                                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
		                                                                                        />
																								    Filter
																								</label>
																							</TD>
																							<TD>
																								<INPUT id=txtBaseFilter name=txtBaseFilter class="text textdisabled" disabled="disabled" style="width=250">
																							</TD>
																							<TD>
																								<INPUT id=cmdBaseFilter name=cmdBaseFilter class="btn btndisabled" disabled="disabled" type=button value="..."
																								    onclick="selectRecordOption('filter')" 
			                                                                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
			                                                                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
			                                                                                        onfocus="try{button_onFocus(this);refreshControls();}catch(e){}"
			                                                                                        onblur="try{button_onBlur(this);}catch(e){}" />
																							</TD>
																						</TR>
																					</TABLE>
																					
																				</TD>
																			</TR>
																		</TR>
																		<TR>
																			<TD>
																				<TABLE WIDTH="100%" class="invisible" cellspacing=0 cellpadding=5>
																					<TR>
																						<TD>
																							<INPUT id=chkPrintInReportHeader name=chkPrintInReportHeader type=checkbox tabindex=-1 
		                                                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                                            <label 
			                                                                                    for="chkPrintInReportHeader"
			                                                                                    class="checkbox"
			                                                                                    tabindex=0 
			                                                                                    onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
				                                                                                onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
				                                                                                onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
	                                                                                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
	                                                                                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
	    																						Display filter or picklist title in the report header
	    																					</label>
    																					</TD>
																					</TR> 
																				</TABLE> 
																			</TD>
																		</TR>																
																	</TABLE>
																</TD>
															</TR>
														</TABLE>											
													</SPAN>
						                       </td>
											</tr>
										</table>
									</td>
								</tr>
							</TABLE>
	
						</div>
						<!-- Second Tab (Options) -->
						<DIV id=div2 style="display:none">
							<TABLE width=100% class="outline" cellspacing=0 cellpadding=5>
								<TR>
									<TD>
										<TABLE width=100% class="invisible" cellspacing=0 cellpadding=5>
											<TD>
												<TABLE class="invisible" cellspacing=0 cellpadding=0>
													<TR>
														<TD>
														    <INPUT type="checkbox" id=chkSRV name=chkSRV tabindex=-1 
		                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                            <label 
				                                                for="chkSRV"
				                                                class="checkbox"
				                                                tabindex=0 
				                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
														        Suppress Repeated Personnel Details
														    </label>
														</TD>
													</TR>
													<TR>
														<TD>
															<INPUT type="checkbox" id=chkShowDurations name=chkShowDurations tabindex=-1 
		                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                            <label 
				                                                for="chkShowDurations"
				                                                class="checkbox"
				                                                tabindex=0 
				                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
															    Show Duration Totals
														    </label>
														</TD>
													</TR>
													<TR>
														<TD>
															<INPUT type="checkbox" id=chkShowInstances name=chkShowInstances tabindex=-1 
		                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                            <label 
				                                                for="chkShowInstances"
				                                                class="checkbox"
				                                                tabindex=0 
				                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
															    Show Instances Count
															</label>
														</TD>
													</TR>
													<TR>
														<TD>
															<INPUT type="checkbox" id=chkShowFormula name=chkShowFormula tabindex=-1 
		                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                            <label 
				                                                for="chkShowFormula"
				                                                class="checkbox"
				                                                tabindex=0 
				                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
															    Show Bradford Factor Formula
															</label>
														</TD>
													</TR>
													<TR>
														<TD>
															<INPUT type="checkbox" id=chkShowAbsenceDetails name=chkAbsenceDetails tabindex=-1
															    onClick="refreshTab2Controls();" 
															    onChange="refreshTab2Controls();"  
		                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                            <label 
				                                                for="chkShowAbsenceDetails"
				                                                class="checkbox"
				                                                tabindex=0 
				                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
															    Show Absence Details
															</label>
														</TD>
													</TR>							
												</TABLE>
											</TD>
										</TABLE>
									    <hr />
										<TABLE width=100% class="invisible" cellspacing=0 cellpadding=5>
											<TD>
												<TABLE class="invisible" cellspacing=0 cellpadding=0>
													<TR>
														<TD>
															<INPUT type="checkbox" id=chkOmitBeforeStart name=chkOmitBeforeStart tabindex=-1
		                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                            <label 
				                                                for="chkOmitBeforeStart"
				                                                class="checkbox"
				                                                tabindex=0 
				                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
															    Omit absences starting before the report start date
															</label>
														</TD>
													</TR>
													<TR>
														<TD>
															<INPUT type="checkbox" id=chkOmitAfterEnd name=chkOmitAfterEnd tabindex=-1
		                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                            <label 
				                                                for="chkOmitAfterEnd"
				                                                class="checkbox"
				                                                tabindex=0 
				                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
															    Omit absences ending after the report end date
															</label>
														</TD>
													</TR>
													<TR>
														<TD>
															<TABLE class="invisible" cellspacing=0 cellpadding=0>
																<TR>
																	<TD>
																		<INPUT type="checkbox" id=chkMinimumBradfordFactor name=chkMinimumBradfordFactor tabindex=-1
																		    onClick="refreshTab2Controls();" 
																		    onChange="refreshTab2Controls();" 
		                                                                    onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                                    onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                        <label 
				                                                            for="chkMinimumBradfordFactor"
				                                                            class="checkbox"
				                                                            tabindex=0 
				                                                            onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                                                        onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                                                        onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                                                                    onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                                                                    onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																		    Minimum Bradford Factor
																		</label>
																		&nbsp
																		&nbsp
																		<INPUT id=txtMinimumBradfordFactor name=txtMinimumBradfordFactor class="text"
																		    onblur="validateNumeric(this);">
																	</TD>
																</TR>										
															</TABLE>
														</TD>
													</TR>
												</TABLE>
											</TD>
										</TABLE>
										<hr />
										<TABLE width=100% class="invisible" cellspacing=0 cellpadding=5>
											<TD>
												<TABLE width=100% class="invisible" cellspacing=0 cellpadding=0>
													<TR>
														<TD>
															Order By : 
														</TD>
														<TD width=50%>
															<SELECT id=cboOrderBy1 name=cboOrderBy1 style="WIDTH: 50%" class="combo"
															    onChange="refreshTab2Controls();">
																<OPTION VALUE="0">&lt;None&gt;</OPTION>															
																<%
																	for iCount = 0 to ubound(aColumnNames,2) - 1
																        Response.Write("<OPTION VALUE = " & """" & aColumnNames(0, iCount) & """" & ">" & aColumnNames(1, iCount) & "</OPTION>")
																    Next
																%>
															</SELECT>
														</TD>
														<TD>
															<INPUT type="checkbox" id=chkOrderBy1Asc name=chkOrderBy1Asc tabindex=-1
                                                                onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
                                                                onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                            <label 
	                                                            for="chkOrderBy1Asc"
	                                                            class="checkbox"
	                                                            tabindex=0 
	                                                            onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
		                                                        onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
															    Ascending
															</label>
														</TD>
													</TR>
													<TR>
														<TD>
															Then : 
														</TD>
														<TD width=50%>
															<SELECT id=cboOrderBy2 name=cboOrderBy2 style="WIDTH: 50%" class="combo"
															    onChange="refreshTab2Controls();">
																<OPTION VALUE="0">&lt;None&gt;</OPTION>															
																<%
																	for iCount = 0 to ubound(aColumnNames,2) - 1
																        Response.Write("<OPTION VALUE = " & """" & aColumnNames(0, iCount) & """" & ">" & aColumnNames(1, iCount) & "</OPTION>")
																    Next
																%>
															</SELECT>
														</TD>
														<TD>
															<INPUT type="checkbox" id=chkOrderBy2Asc name=chkOrderBy2Asc tabindex=-1
                                                                onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
                                                                onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                            <label 
	                                                            for="chkOrderBy2Asc"
	                                                            class="checkbox"
	                                                            tabindex=0 
	                                                            onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
		                                                        onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
		                                                        onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                                                onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                                                onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
															    Ascending
															</label>
														</TD>
													</TR>
												</TABLE>
											</TD>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
						</DIV> 

						<!-- Third tab -->
						<DIV id=div3 style="visibility:hidden;display:none">
							<TABLE WIDTH="100%" height="100%" class="outline" cellspacing=0 cellpadding=5>
								<tr valign=top> 
									<td>
										<TABLE WIDTH="100%" class="invisible" CELLSPACING=10 CELLPADDING=0>
											<tr>						
												<td valign=top rowspan=2 width=25% height="100%">
													<table class="outline" cellspacing="0" cellpadding="4" width=160 height=100%>
														<tr height=10> 
															<td height=10 align=left valign=top>
																Output Format : <BR><BR>
																<TABLE class="invisible" cellspacing="0" cellpadding="0" width=100%>
																	<tr height=20>
																		<td width=5>&nbsp</td>
																		<td align=left width=15>
																			<INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat0 value=0
																			    onClick="formatClick(0);" 
		                                                                        onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																		</td>
																		<td align=left nowrap>
                                                                            <label 
                                                                                tabindex="-1"
                                                                                for="optOutputFormat0"
                                                                                class="radio"
	                                                                            onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
	                                                                            onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
	                                                                        />
    																			Data Only
    																		</label>
																		</td>
																		<td width=5>&nbsp</td>
																	</tr>
<%
'MH20040705
'Don't allow CSV for Bradford
if session("StandardReport_Type") = 16 then
%>
	                                                                <INPUT type=hidden width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat1 value=1
	                                                                    onClick="formatClick(1);" 
		                                                                onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
		                                                                onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                        onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                        onblur="try{radio_onBlur(this);}catch(e){}"/>
<%
else
%>
        														    <tr height=10>
																	    <td colspan=4></td>
																    </tr>
																    <tr height=20>
																	    <td width=5>&nbsp</td>
																	    <td align=left width=15>
																		    <INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat1 value=1
																		        onClick="formatClick(1);" 
	                                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
	                                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																	    </td>
																	    <td align=left nowrap>
                                                                            <label 
                                                                                tabindex="-1"
                                                                                for="optOutputFormat1"
                                                                                class="radio"
                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                            />
																			    CSV File
																		    </label>
																	    </td>
																	    <td width=5>&nbsp</td>
																	</tr>
<%
end if
%>
																    <tr height=10> 
																	    <td colspan=4></td>
																    </tr>
																    <tr height=20>
																	    <td width=5>&nbsp</td>
																	    <td align=left width=15>
																		    <INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat2 value=2
																		        onClick="formatClick(2);" 
	                                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
	                                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																	    </td>
																	    <td align=left nowrap>
                                                                            <label 
                                                                                tabindex="-1"
                                                                                for="optOutputFormat2"
                                                                                class="radio"
                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                            />
																			    HTML Document
																		    </label>
																	    </td>
																	    <td width=5>&nbsp</td>
																    </tr>
																    <tr height=10> 
																	    <td colspan=4></td>
																    </tr>
																    <tr height=20>
																	    <td width=5>&nbsp</td>
																	    <td align=left width=15>
																		    <INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat3 value=3
																		        onClick="formatClick(3);" 
	                                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
	                                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																	    </td>
																	    <td align=left nowrap>
                                                                            <label 
                                                                                tabindex="-1"
                                                                                for="optOutputFormat3"
                                                                                class="radio"
                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                            />
																		        Word Document
																		    </label>
																	    </td>
																	    <td width=5>&nbsp</td>
																    </tr>
																    <tr height=10> 
																	    <td colspan=4></td>
																    </tr>
																    <tr height=20>
																	    <td width=5>&nbsp</td>
																	    <td align=left width=15>
																		    <INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat4 value=4
																		        onClick="formatClick(4);" 
	                                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
	                                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																	    </td>
																	    <td align=left nowrap>
                                                                            <label 
                                                                                tabindex="-1"
                                                                                for="optOutputFormat4"
                                                                                class="radio"
                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                            />
																		        Excel Worksheet
																		    </label>
																	    </td>
																	    <td width=5>&nbsp</td>
																    </tr>
																    <tr height=10> 
																	    <td colspan=4></td>
																    </tr>

																    <tr height=5>
																	    <td width=5>&nbsp</td>
																	    <td align=left width=15>
																		    <INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat5 value=5
																		        onClick="formatClick(5);" 
	                                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
	                                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																	    </td>
																	    <td>
                                                                            <label 
                                                                                tabindex="-1"
                                                                                for="optOutputFormat5"
                                                                                class="radio"
                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                            />
																		        Excel Chart
																		    </label>
																	    </td>
																	    <td width=5>&nbsp</td>
																    </tr>
<%
'MH20031211 Fault 7787
'Don't allow Pivot for Bradford
if session("StandardReport_Type") = 16 then
%>
																	<INPUT type=hidden width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat6 value=6
																	    onClick="formatClick(6);" 
                                                                        onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
                                                                        onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                        onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                        onblur="try{radio_onBlur(this);}catch(e){}"/>
<%
else
%>
																	<tr height=10>
																		<td colspan=4></td>
																	</tr>
																	<tr height=5>
																		<td width=5>&nbsp</td>
																		<td align=left width=15>
																			<INPUT type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat6 value=6
																			    onClick="formatClick(6);" 
	                                                                            onmouseover="try{radio_onMouseOver(this);}catch(e){}" 
	                                                                            onmouseout="try{radio_onMouseOut(this);}catch(e){}"
                                                                                onfocus="try{radio_onFocus(this);}catch(e){}"
                                                                                onblur="try{radio_onBlur(this);}catch(e){}"/>
																		</td>
																		<td>
                                                                            <label 
                                                                                tabindex="-1"
                                                                                for="optOutputFormat6"
                                                                                class="radio"
                                                                                onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}" 
                                                                                onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}"
                                                                            />
																			    Excel Pivot Table
																			</label>
																		</td>
																		<td width=5>&nbsp</td>
																	</tr>
<%
end if
%>
																	<tr height=5> 
																		<td colspan=4></td>
																	</tr>
																</TABLE>
															</td>
														</tr>
													</table>
												</td>
												<td  valign=top width=75%>
													<table class="outline" cellspacing="0" cellpadding="4" width=100%  height=100%>
														<tr height=10> 
															<td height=10 align=left valign=top>
																Output Destination(s) : <BR><BR>
																<TABLE class="invisible" cellspacing="0" cellpadding="0" width=100%>
																	<tr height=20>
																		<td width=5>&nbsp</td>
																		<td align=left colspan=6 nowrap>
																			<input name=chkPreview id=chkPreview type=checkbox disabled="disabled" tabindex=-1 
																			    onClick="refreshTab3Controls();"
		                                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                            <label 
				                                                                for="chkPreview"
				                                                                class="checkbox checkboxdisabled"
				                                                                tabindex=0 
				                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																			    Preview on screen
																			</label>
																		</td>
																		<td width=5>&nbsp</td>
																	</tr>
																	<tr height=10> 
																		<td colspan=8></td>
																	</tr>
																	<tr height=20>
																		<td width=5>&nbsp</td>
																		<td align=left colspan=6 nowrap>
																			<input name=chkDestination0 id=chkDestination0 type=checkbox disabled="disabled" tabindex=-1 
																			    onClick="refreshTab3Controls();"
		                                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                            <label 
				                                                                for="chkDestination0"
				                                                                class="checkbox checkboxdisabled"
				                                                                tabindex=0 
				                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
    																			Display output on screen
    																		</label>
																		</td>
																		<td width=5>&nbsp</td>
																	</tr>
																	<tr height=10> 
																		<td colspan=8></td>
																	</tr>
																	<tr height=20>
																		<td width=5>&nbsp</td>
																		<td align=left nowrap>
																			<input name=chkDestination1 id=chkDestination1 type=checkbox disabled="disabled" tabindex=-1  
																			    onClick="refreshTab3Controls();"
		                                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                            <label 
				                                                                for="chkDestination1"
				                                                                class="checkbox checkboxdisabled"
				                                                                tabindex=0 
				                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																			    Send to printer 
																			</label>
																		</td>
																		<td width=30 nowrap>&nbsp</td>
																		<td align=left nowrap>
																			Printer location : 
																		</td>
																		<td width=15>&nbsp</td>
																		<td colspan=2>
																			<select id=cboPrinterName name=cboPrinterName class="combo" width=100% style="WIDTH: 220">
																			</select>								
																		</td>
																		<td width=5>&nbsp</td>
																	</tr>
																	<tr height=10> 
																		<td colspan=8></td>
																	</tr>
																	<tr height=20>
																		<td width=5>&nbsp</td>
																		<td align=left nowrap>
																			<input name=chkDestination2 id=chkDestination2 type=checkbox disabled="disabled" tabindex=-1  
																			    onClick="refreshTab3Controls();"
		                                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                            <label 
				                                                                for="chkDestination2"
				                                                                class="checkbox checkboxdisabled"
				                                                                tabindex=0 
				                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																			    Save to file
																			</label>
																		</td>
																		<td width=30 nowrap>&nbsp</td>
																		<td align=left nowrap>
																			File name :   
																		</td>
																		<td width=15 nowrap>&nbsp</td>
																		<td colspan=2>
																			<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 width="100%">
																				<TR>
																					<TD>
																						<INPUT id=txtFilename name=txtFilename class="text textdisabled" disabled="disabled" style="WIDTH: 200">
																					</TD>
																					<TD width=25>
																						<INPUT id=cmdFilename name=cmdFilename class="btn" style="WIDTH: 100%" type=button value="..."
																						    onClick="saveFile();"  
			                                                                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
			                                                                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
			                                                                                onfocus="try{button_onFocus(this);}catch(e){}"
			                                                                                onblur="try{button_onBlur(this);}catch(e){}" />
																					</TD>
																				</TD>
																			</TABLE>
																		</TD>
																		<td width=5>&nbsp</td>
																	</tr>

																	<tr height=10> 
																		<td colspan=8></td>
																	</tr>
																	<tr height=20>
																		<td width=5>&nbsp</td>
																		<td align=left nowrap>
																		</td>
																		<td width=30 nowrap>&nbsp</td>
																		<td align=left nowrap>
																			If existing file :
																		</td>
																		<td width=15 nowrap>&nbsp</td>
																		<td colspan=2 width=100% nowrap>
																			<select id=cboSaveExisting name=cboSaveExisting class="combo" width=100% style="WIDTH: 100%">	
																			</select>								
																		</td>
																		<td width=5>&nbsp</td>
																	</tr>

																	<tr height=10> 
																		<td colspan=8></td>
																	</tr>
																	<tr height=20>
																		<td width=5>&nbsp</td>
																		<td align=left nowrap>
																			<input name=chkDestination3 id=chkDestination3 type=checkbox disabled="disabled" tabindex=-1
																			    onClick="refreshTab3Controls();"
		                                                                        onmouseover="try{checkbox_onMouseOver(this);}catch(e){}" 
		                                                                        onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                                                            <label 
				                                                                for="chkDestination3"
				                                                                class="checkbox checkboxdisabled"
				                                                                tabindex=0 
				                                                                onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
					                                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}" 
					                                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
		                                                                        onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
		                                                                        onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																			    Send as email 
																			</label>
																		</td>
																		<td width=30 nowrap>&nbsp</td>
																		<td align=left nowrap>
																			Email group :   
																		</td>
																		<td width=15 nowrap>&nbsp</td>
																		<td colspan=2>
																			<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 width="100%">
																				<TR>
																					<TD>
																						<INPUT id=txtEmailGroup name=txtEmailGroup class="text textdisabled" disabled="disabled" style="WIDTH: 200">
																						<INPUT id=txtEmailGroupID name=txtEmailGroupID type=hidden>																						
																					</TD>
																					<TD width=25>
																						<INPUT id=cmdEmailGroup name=cmdEmailGroup class="btn" style="WIDTH: 100%" type=button value="..."
																						    onClick="selectEmailGroup();"  
			                                                                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
			                                                                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
			                                                                                onfocus="try{button_onFocus(this);}catch(e){}"
			                                                                                onblur="try{button_onBlur(this);}catch(e){}" />
																					</TD>
																				</TD>
																			</TABLE>
																		</TD>
																		<td width=5>&nbsp</td>
																	</tr>
																	<tr height=10> 
																		<td colspan=8></td>
																	</tr>
																	<tr height=20>
																		<td width=5>&nbsp</td>
																		<td align=left>&nbsp</td>
																		<td width=30 nowrap>&nbsp</td>
																		<td align=left nowrap>
																			Email subject :   
																		</td>
																		<td width=15>&nbsp</td>
																		<TD colspan=2 width=100% nowrap>
																			<INPUT id=txtEmailSubject class="text textdisabled" disabled="disabled" maxlength=255 name=txtEmailSubject style=" WIDTH: 220">
																		</TD>
																		<td width=5>&nbsp</td>
																	</tr>
																	<tr height=10> 
																		<td colspan=8></td>
																	</tr>
																	<tr height=20>
																		<td width=5>&nbsp</td>
																		<td align=left>&nbsp</td>
																		<td width=30 nowrap>&nbsp</td>
																		<td align=left nowrap>
																			Attach as :   
																		</td>
																		<td width=15>&nbsp</td>
																		<TD colspan=2 width=100% nowrap>
																			<INPUT id=txtEmailAttachAs class="text textdisabled" disabled="disabled" maxlength=255 name=txtEmailAttachAs style=" WIDTH: 220">
																		</TD>
																		<td width=5>&nbsp</td>
																	</tr>
																	<tr height=10> 
																		<td colspan=8></td>
																	</tr>
																</TABLE>
															</td>
														</tr>
													</table>
												</td>
											</tr>
										</TABLE>
									</td>
								</tr>
							</TABLE>
						</DIV>
													
					</TD>
					<TD width=10></TD>
				</TR> 

				<tr height=10> 
					<td colspan=3></td>
				</tr> 

				<TR height=10>
					<TD width=10></TD>
					<TD>
						<TABLE WIDTH="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD>&nbsp;</TD>
								<TD width=80>
									<input type=button id=cmdOK name=cmdOK class="btn" value=Run style="WIDTH: 100%" 
									    onclick="absence_okClick()"
		                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
		                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
		                                onfocus="try{button_onFocus(this);}catch(e){}"
		                                onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
								<TD width=10></TD>
							</TR>
						</TABLE>
					</TD>
					<TD width=10></TD>
				</TR> 

				<tr height=5> 
					<td colspan=3></td>
				</tr>
			</TABLE>
		</TD>
	</TR>
</TABLE>

    <input type='hidden' id="txtDatabase" name="txtDatabase" value="<%=session("Database")%>">
    <input type="hidden" id="txtWordVer" name="txtWordVer" value="<%=Session("WordVer")%>">
    <input type="hidden" id="txtExcelVer" name="txtExcelVer" value="<%=Session("ExcelVer")%>">
    <input type="hidden" id="txtWordFormats" name="txtWordFormats" value="<%=Session("WordFormats")%>">
    <input type="hidden" id="txtExcelFormats" name="txtExcelFormats" value="<%=Session("ExcelFormats")%>">
    <input type="hidden" id="txtWordFormatDefaultIndex" name="txtWordFormatDefaultIndex" value="<%=Session("WordFormatDefaultIndex")%>">
    <input type="hidden" id="txtExcelFormatDefaultIndex" name="txtExcelFormatDefaultIndex" value="<%=Session("ExcelFormatDefaultIndex")%>">
</FORM>

<form id="frmAbsenceUseful" name="frmAbsenceUseful" style="visibility: hidden; display: none">
    <input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
    <input type="hidden" id="txtLoading" name="txtLoading" value="Y">
    <input type="hidden" id="txtCurrentBaseTableID" name="txtCurrentBaseTableID">
    <input type="hidden" id="txtCurrentChildTableID" name="txtCurrentChildTableID" value="0">
    <input type="hidden" id="txtTablesChanged" name="txtTablesChanged">
    <input type="hidden" id="txtSelectedColumnsLoaded" name="txtSelectedColumnsLoaded" value="0">
    <input type="hidden" id="txtSortLoaded" name="txtSortLoaded" value="0">
    <input type="hidden" id="txtRepetitionLoaded" name="txtRepetitionLoaded" value="0">
    <input type="hidden" id="txtChildsLoaded" name="txtChildsLoaded" value="0">
    <input type="hidden" id="txtChanged" name="txtChanged" value="0">
    <input type="hidden" id="txtUtilID" name="txtUtilID" value='<%=session("utilid")%>'>
    <input type="hidden" id="txtChildCount" name="txtChildCount" value='<%=session("childcount")%>'>
    <input type="hidden" id="txtHiddenChildFilterCount" name="txtHiddenChildFilterCount" value='<%=session("hiddenfiltercount")%>'>
    <input type="hidden" id="txtLockGridEvents" name="txtLockGridEvents" value="0">
    <%
        Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
        Response.Write("<INPUT type='hidden' id=txtAction name=txtAction value=" & Session("action") & ">" & vbCrLf)
    %>
</form>

<form action="util_run_promptedvalues" target="string(15)" method="post" id="frmPostDefinition" name="frmPostDefinition">
    <input type="hidden" id="txtRecordSelectionType" name="txtRecordSelectionType">
    <input type="hidden" id="txtFromDate" name="txtFromDate">
    <input type="hidden" id="txtToDate" name="txtToDate">
    <input type="hidden" id="txtBasePicklistID" name="txtBasePicklistID" value="0">
    <input type="hidden" id="txtBasePicklist" name="txtBasePicklist">
    <input type="hidden" id="txtBaseFilterID" name="txtBaseFilterID" value="0">
    <input type="hidden" id="txtBaseFilter" name="txtBaseFilter">
    <input type="hidden" id="txtAbsenceTypes" name="txtAbsenceTypes">
    <input type="hidden" id="txtSRV" name="txtSRV">
    <input type="hidden" id="txtShowDurations" name="txtShowDurations">
    <input type="hidden" id="txtShowInstances" name="txtShowInstances">
    <input type="hidden" id="txtShowFormula" name="txtShowFormula">
    <input type="hidden" id="txtOmitBeforeStart" name="txtOmitBeforeStart">
    <input type="hidden" id="txtOmitAfterEnd" name="txtOmitAfterEnd">
    <input type="hidden" id="txtOrderBy1" name="txtOrderBy1">
    <input type="hidden" id="txtOrderBy1ID" name="txtOrderBy1ID">
    <input type="hidden" id="txtOrderBy1Asc" name="txtOrderBy1Asc">
    <input type="hidden" id="txtOrderBy2" name="txtOrderBy2">
    <input type="hidden" id="txtOrderBy2ID" name="txtOrderBy2ID">
    <input type="hidden" id="txtOrderBy2Asc" name="txtOrderBy2Asc">
    <input type="hidden" id="txtMinimumBradfordFactor" name="txtMinimumBradfordFactor">
    <input type="hidden" id="txtMinimumBradfordFactorAmount" name="txtMinimumBradfordFactorAmount">
    <input type="hidden" id="txtDisplayBradfordDetail" name="txtDisplayBradfordDetail">
    <input type="hidden" id="txtPrintFPinReportHeader" name="txtPrintFPinReportHeader">
    <input type="hidden" id="txtRecSelCurrentID" name="txtRecSelCurrentID" value='<%=Session("optionRecordID")%>'>
    <input type="hidden" id="utiltype" name="utiltype" value='<%=Session("StandardReport_Type")%>'>
    <input type="hidden" id="utilid" name="utilid" value="0">
    <input type="hidden" id="utilname" name="utilname" value="Standard Report">
    <input type="hidden" id="action" name="action" value="run">
    <input type="hidden" id="txtSend_OutputPreview" name="txtSend_OutputPreview">
    <input type="hidden" id="txtSend_OutputFormat" name="txtSend_OutputFormat">
    <input type="hidden" id="txtSend_OutputScreen" name="txtSend_OutputScreen">
    <input type="hidden" id="txtSend_OutputPrinter" name="txtSend_OutputPrinter">
    <input type="hidden" id="txtSend_OutputPrinterName" name="txtSend_OutputPrinterName">
    <input type="hidden" id="txtSend_OutputSave" name="txtSend_OutputSave">
    <input type="hidden" id="txtSend_OutputSaveExisting" name="txtSend_OutputSaveExisting">
    <input type="hidden" id="txtSend_OutputEmail" name="txtSend_OutputEmail">
    <input type="hidden" id="txtSend_OutputEmailAddr" name="txtSend_OutputEmailAddr">
    <input type="hidden" id="txtSend_OutputEmailSubject" name="txtSend_OutputEmailSubject">
    <input type="hidden" id="txtSend_OutputEmailAttachAs" name="txtSend_OutputEmailAttachAs">
    <input type="hidden" id="txtSend_OutputFilename" name="txtSend_OutputFilename">
</form>

	<!-- Stuff required to make record selection stuff work -->
<form id="frmCustomReportStuff" name="frmCustomReportStuff">
    <input type="hidden" id="baseHidden" name="baseHidden">
</form>

<form id="frmEmailSelection" name="frmEmailSelection" target="emailSelection" action="util_emailSelection.asp" method="post" style="visibility: hidden; display: none">
    <input type="hidden" id="EmailSelCurrentID" name="EmailSelCurrentID">
</form>

<form id="frmRecordSelection" name="frmRecordSelection" target="recordSelection" action="util_recordSelection.asp" method="post" style="visibility: hidden; display: none">
    <input type="hidden" id="recSelTable" name="recSelTable" value='base'>
    <input type="hidden" id="recSelType" name="recSelType">
    <input type="hidden" id="recSelTableID" name="recSelTableID">
    <input type="hidden" id="recSelCurrentID" name="recSelCurrentID" value='<%=Session("optionRecordID")%>'>
</form>

<form action="default_submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
    <%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</form>

<!-- Form to return to record edit screen -->
<form action="emptyoption" method="post" id="frmRecordEdit" name="frmRecordEdit">
</form>

<form id="frmSessionInformation" name="frmSessionInformation" style="visibility: hidden; display: none">
    <input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
    <input type="hidden" id="txtLoading" name="txtLoading" value="Y">
    <input type="hidden" id="txtChanged" name="txtChanged" value="0">
    <input type="hidden" id="txtUtilID" name="txtUtilID" value='<% =session("StandardReport_Type")%>'>
    <%
        Dim cmdDefinition
        Dim prmModuleKey
        Dim prmParameterKey
        Dim prmParameterValue
        
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
        cmdDefinition.Execute()

        Response.Write("<INPUT type='hidden' id=txtPersonnelTableID name=txtPersonnelTableID value=" & cmdDefinition.Parameters("paramValue").Value & ">" & vbCrLf)
			
        cmdDefinition = Nothing
    %>
</form>


<script type="text/javascript">
    stdrpt_def_absence_window_onload();
</script>
