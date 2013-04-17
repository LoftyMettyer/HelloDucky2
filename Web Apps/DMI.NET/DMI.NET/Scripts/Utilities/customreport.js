

function util_run_customreports_addActiveXHandlers() {

    OpenHR.addActiveXHandler("tblGrid", "onresize", tblGrid_onresize);
    OpenHR.addActiveXHandler("ssOleDBGridDefSelRecords", "printinitialize", ssOleDBGridDefSelRecords_PrintInitialize);
    //    OpenHR.addActiveXHandler("ssOleDBGridDefselRecords", "printerror", ssOleDBGridDefselRecords_PrintError);

    OpenHR.addActiveXHandler("ssHiddenGrid", "printinitialize", ssHiddenGrid_PrintInitialize);
    OpenHR.addActiveXHandler("ssHiddenGrid", "printbegin", ssHiddenGrid_PrintBegin);
    ///  OpenHR.addActiveXHandler("ssHiddenGrid", "printerror", ssHiddenGrid_PrintError);
}

function tblGrid_onresize() {

    try {

        if (txtNoRecs.value == 0) {
            frmOutput.ssOleDBGridDefSelRecords.Refresh();
            if ((frmOutput.ssOleDBGridDefSelRecords.visiblerows() + 1) >= frmOutput.ssOleDBGridDefSelRecords.rows()) {
                frmOutput.ssOleDBGridDefSelRecords.FirstRow = frmOutput.ssOleDBGridDefSelRecords.AddItemBookmark(0);
            }
        }
    } catch (e) {
        return;
    }
}

function ssOleDBGridDefSelRecords_PrintInitialize(ssPrintInfo) {
    //Underline headings if not printing gridlines
    ssPrintInfo.PrintGridlines = 3; // 0 = none, 3 = all

    ssPrintInfo.PrintHeaders = 0; // 0 = top of every page, 1 = top of report
    ssPrintInfo.Portrait = false;
    ssPrintInfo.Copies = 1;
    ssPrintInfo.Collate = true;
    ssPrintInfo.PrintColors = true;

    ssPrintInfo.RowAutoSize = true;
    ssPrintInfo.PrintColumnHeaders = 1;
    ssPrintInfo.MaxLinesPerRow = 2;

    ssPrintInfo.PageHeader = "	" + frmOutput.ssOleDBGridDefSelRecords.Caption + "	";
    ssPrintInfo.PageFooter = "Printed on <date> at <time> by " + frmOriginalDefinition.txtUserName.value + "	" + "	" + "Page <page number>";
}

function ssOleDBGridDefselRecords_PrintError(lngPrintError, iResponse) {
    if (lngPrintError == 30457) {
        frmOriginalDefinition.txtCancelPrint.value = 1;
    }
}

function ssHiddenGrid_PrintInitialize(ssPrintInfo) {

    //Underline headings if not printing gridlines
    ssPrintInfo.PrintGridlines = 3; // 0 = none, 3 = all

    ssPrintInfo.PrintHeaders = 0; // 0 = top of every page, 1 = top of report
    ssPrintInfo.Portrait = false;
    ssPrintInfo.Copies = 1;
    ssPrintInfo.Collate = true;
    ssPrintInfo.PrintColors = true;

    ssPrintInfo.RowAutoSize = true;
    ssPrintInfo.PrintColumnHeaders = 1;
    ssPrintInfo.MaxLinesPerRow = 2;

    ssPrintInfo.PageHeader = "	" + frmOutput.ssOleDBGridDefSelRecords.Caption + "	";
    ssPrintInfo.PageFooter = "Printed on <date> at <time> by " + frmOriginalDefinition.txtUserName.value + "	" + "	" + "Page <page number>";
}

function ssHiddenGrid_PrintBegin(ssPrintInfo) {

    if (frmOriginalDefinition.txtOptionsDone.value == 0) {
        frmOriginalDefinition.txtOptionsPortrait.value = ssPrintInfo.Portrait;
        frmOriginalDefinition.txtOptionsMarginLeft.value = ssPrintInfo.MarginLeft;
        frmOriginalDefinition.txtOptionsMarginRight.value = ssPrintInfo.MarginRight;
        frmOriginalDefinition.txtOptionsMarginTop.value = ssPrintInfo.MarginTop;
        frmOriginalDefinition.txtOptionsMarginBottom.value = ssPrintInfo.MarginBottom;
        frmOriginalDefinition.txtOptionsCopies.value = ssPrintInfo.Copies;
    }
    else {
        ssPrintInfo.Portrait = frmOriginalDefinition.txtOptionsPortrait.value;
        ssPrintInfo.MarginLeft = frmOriginalDefinition.txtOptionsMarginLeft.value;
        ssPrintInfo.MarginRight = frmOriginalDefinition.txtOptionsMarginRight.value;
        ssPrintInfo.MarginTop = frmOriginalDefinition.txtOptionsMarginTop.value;
        ssPrintInfo.MarginBottom = frmOriginalDefinition.txtOptionsMarginBottom.value;
        ssPrintInfo.Copies = frmOriginalDefinition.txtOptionsCopies.value;
    }
}

function ssHiddenGrid_PrintError(lngPrintError, iResponse) {
    if (lngPrintError == 30457) {
        frmOriginalDefinition.txtCancelPrint.value = 1;
    }
}


function ShowReport() {
    var iPollPeriod;
    var iPollCounter;
    var iDummy;

    iPollPeriod = 100;
    iPollCounter = iPollPeriod;

    var frmPopup = document.getElementById("frmPopup");
    var i;

    if (txtSuccessFlag.value == 2) {

        var frmOutput = document.getElementById("frmOutput");

        setGridFont(frmOutput.ssHiddenGrid);
        setGridFont(frmOutput.ssOleDBGridDefSelRecords);

        frmOutput.ssOleDBGridDefSelRecords.style.visibility = 'hidden';
        frmOutput.ssOleDBGridDefSelRecords.Redraw = false;
        frmOutput.ssOleDBGridDefSelRecords.style.visibility = 'visible';
        frmOutput.ssOleDBGridDefSelRecords.focus();

        var dataCollection = frmGridItems.elements;
        if (dataCollection != null) {
            for (i = 0; i < dataCollection.length; i++) {
                //                   if (i==iPollCounter) 
                //                    {			
                //                       try {
                //                            var frmRefresh = OpenHR.getForm("pollframe","frmHit");	
                //                            var testDataCollection = frmRefresh.elements;
                //                            iDummy = testDataCollection.txtDummy.value;
                //                            OpenHR.submitForm(frmRefresh);
                //                            iPollCounter = iPollCounter + iPollPeriod;
                //                        }
                //                        catch(e) {}
                //                    }                        

                sControlName = dataCollection.item(i).name;
                sControlName = sControlName.substr(0, 12);
                if (sControlName == "txtGridItem_") {
                    frmOutput.ssOleDBGridDefSelRecords.AddItem(dataCollection.item(i).value);
                }
            }
        }

        // JPD 19/03/02 Fault 3665
        for (i = 0; i < frmOutput.ssOleDBGridDefSelRecords.Columns.Count; i++) {
            if (frmOutput.ssOleDBGridDefSelRecords.Columns(i).Width > 32000) {
                frmOutput.ssOleDBGridDefSelRecords.Columns(i).Width = 32000;
            }
        }

        //debugger;
        //if (frmExportData.txtPreview.value == 'False') {
        //    frmOutput.ssOleDBGridDefSelRecords.style.visibility = 'hidden';
        //    frmOutput.ssOleDBGridDefSelRecords.Redraw = true;
        //    ExportData("OUTPUTRUN");
        //    document.getElementById('output').style.visibility = 'hidden';
        //    document.getElementById('close').value = 'OK';
        //    document.getElementById('tdOutputMSG').innerText = "Custom Report : '" + "' Completed Successfully.";
        //    return;
        //}
        //else {
            frmOutput.ssOleDBGridDefSelRecords.Redraw = true;
            frmOutput.ssOleDBGridDefSelRecords.style.visibility = 'visible';
        //}
    }

    $("#top").hide();
    $("#reportworkframe").show();

}

function ExportDataPrompt() {
    sURL = "util_run_outputoptions" +
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

function openDialog(pDestination, pWidth, pHeight, psResizable, psScroll) {
    dlgwinprops = "center:yes;" +
        "dialogHeight:" + pHeight + "px;" +
        "dialogWidth:" + pWidth + "px;" +
        "help:no;" +
        "resizable:" + psResizable + ";" +
        "scroll:" + psScroll + ";" +
        "status:no;";
    window.showModalDialog(pDestination, self, dlgwinprops);
}

function replace(sExpression, sFind, sReplace) {
    //gi (global search, ignore case)
    var re = new RegExp(sFind, "gi");
    sExpression = sExpression.replace(re, sReplace);
    return (sExpression);
}

function getData() {
    loadAddRecords();
}

function dataOnlyPrint() {
    // PageHeaderFont and PageFooterFont don't function in ASPs.
    //  frmOutput.ssOleDBGridDefSelRecords.PageHeaderFont.Name = "Verdana";
    //  frmOutput.ssOleDBGridDefSelRecords.PageHeaderFont.Size = 12;
    //  frmOutput.ssOleDBGridDefSelRecords.PageHeaderFont.Bold = true;
    //  frmOutput.ssOleDBGridDefSelRecords.PageHeaderFont.Underline = true;

    //  frmOutput.ssOleDBGridDefSelRecords.PageFooterFont.Name = "Verdana";
    //  frmOutput.ssOleDBGridDefSelRecords.PageFooterFont.Size = 8;
    //  frmOutput.ssOleDBGridDefSelRecords.PageFooterFont.Bold = false;
    //  frmOutput.ssOleDBGridDefSelRecords.PageFooterFont.Underline = false;

    frmOriginalDefinition.txtOptionsDone.value = 0;

    if (frmExportData.pagebreak.value == "True") {
        // Need to loop through the grid, selecting rows until we find a '*' in
        // the first column ('PageBreak').  
        frmOriginalDefinition.txtCancelPrint.value = 0;

        frmOutput.ssHiddenGrid.Caption = frmOutput.ssOleDBGridDefSelRecords.caption;
        frmOutput.ssHiddenGrid.RemoveAll();
        frmOutput.ssHiddenGrid.Columns.RemoveAll();

        for (iColIndex = 0; iColIndex < frmOutput.ssOleDBGridDefSelRecords.Cols; iColIndex++) {
            frmOutput.ssHiddenGrid.Columns.Add(iColIndex);
            frmOutput.ssHiddenGrid.Columns(iColIndex).Width = frmOutput.ssOleDBGridDefSelRecords.Columns(iColIndex).Width;
            frmOutput.ssHiddenGrid.Columns(iColIndex).Visible = frmOutput.ssOleDBGridDefSelRecords.Columns(iColIndex).Visible;
            frmOutput.ssHiddenGrid.Columns(iColIndex).Caption = frmOutput.ssOleDBGridDefSelRecords.Columns(iColIndex).Caption;
            frmOutput.ssHiddenGrid.Columns(iColIndex).Name = frmOutput.ssOleDBGridDefSelRecords.Columns(iColIndex).Name;
            frmOutput.ssHiddenGrid.Columns(iColIndex).Alignment = frmOutput.ssOleDBGridDefSelRecords.Columns(iColIndex).Alignment;
            frmOutput.ssHiddenGrid.Columns(iColIndex).CaptionAlignment = frmOutput.ssOleDBGridDefSelRecords.Columns(iColIndex).CaptionAlignment;
        }

        frmOutput.ssOleDBGridDefSelRecords.redraw = false;
        frmOutput.ssOleDBGridDefSelRecords.moveFirst();

        for (iIndex = 1; iIndex <= frmOutput.ssOleDBGridDefSelRecords.rows; iIndex++) {
            if (frmOutput.ssOleDBGridDefSelRecords.Columns(0).value == "*") {
                // NB. In DatMgr we just printSelectedRows. This doesn't work in an ASP
                // so I copy the required rows to a hidden grid and do printAll on that.
                if (frmOriginalDefinition.txtOptionsDone.value == 0) {
                    button_disable(window.parent.parent.parent.frames("top").frmPopup.Cancel, true);
                    frmOutput.ssHiddenGrid.PrintData(23, false, true);
                    try {
                        button_disable(window.parent.parent.frames("top").frmPopup.Cancel, false);
                    }
                    catch (e) { }

                    frmOriginalDefinition.txtOptionsDone.value = 1;
                    if (frmOriginalDefinition.txtCancelPrint.value == 1) {
                        frmOutput.ssOleDBGridDefSelRecords.redraw = true;
                        return;
                    }
                }
                else {
                    frmOutput.ssHiddenGrid.PrintData(23, false, false);
                }
                frmOutput.ssHiddenGrid.RemoveAll();
            }
            else {
                sAddItem = new String("");
                for (iColIndex = 0; iColIndex < frmOutput.ssOleDBGridDefSelRecords.Cols; iColIndex++) {
                    if (iColIndex > 0) {
                        sAddItem = sAddItem + "	";
                    }
                    sAddItem = sAddItem + frmOutput.ssOleDBGridDefSelRecords.Columns(iColIndex).value;
                }



                frmOutput.ssHiddenGrid.AddItem(sAddItem);
            }

            if (iIndex < frmOutput.ssOleDBGridDefSelRecords.rows) {
                frmOutput.ssOleDBGridDefSelRecords.MoveNext();
            }
            else {
                if (frmOriginalDefinition.txtOptionsDone.value == 0) {
                    button_disable(window.parent.parent.parent.frames("top").frmPopup.Cancel, true);
                    frmOutput.ssHiddenGrid.PrintData(23, false, true);
                    try {
                        button_disable(window.parent.parent.frames("top").frmPopup.Cancel, false);
                    }
                    catch (e) { }

                    if (frmOriginalDefinition.txtCancelPrint.value == 1) {
                        frmOutput.ssOleDBGridDefSelRecords.redraw = true;
                        return;
                    }
                }
                else {
                    frmOutput.ssHiddenGrid.PrintData(23, false, false);
                }
                break;
            }
        }
        frmOutput.ssOleDBGridDefSelRecords.redraw = true;
    }
    else {
        button_disable(window.parent.parent.parent.frames("top").frmPopup.Cancel, true);
        frmOutput.ssOleDBGridDefSelRecords.PrintData(23, false, true);
        try {
            button_disable(window.parent.parent.frames("top").frmPopup.Cancel, false);
        }
        catch (e) { }
    }
}

