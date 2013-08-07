
function ShowCustomReport() {

	var i;
	var sControlName;

	if (txtSuccessFlag.value == 2) {

		var frmGridItems = OpenHR.getForm("reportworkframe", "frmGridItems");

		var dataCollection = frmGridItems.elements;
		if (dataCollection != null) {

			for (i = 0; i < dataCollection.length; i++) {
				sControlName = dataCollection.item(i).name;
				sControlName = sControlName.substr(0, 12);
				if (sControlName == "txtGridItem_") {
					$('#ssOleDBGridDefSelRecords > tbody:last').append(dataCollection.item(i).value);
					}
				}
			}

		if ($("#txtHasSummaryColumns")[0].value == "True") {
			$(".summarytablecolumn").css("visibility", "visible");
			$(".summarytablecolumn").css("display", "block");

		} else {
			$(".summarytablecolumn").css("visibility", "hidden");
			$(".summarytablecolumn").css("display", "none");
		}

	}

	$("#top").hide();
	$("#reportworkframe").show();

}

function ExportDataPrompt() {

	var frmExportData = OpenHR.getForm("reportworkframe", "frmExportData");
	OpenHR.submitForm(frmExportData, "outputoptions");

	$("#reportworkframe").hide();
	$("#reportbreakdownframe").hide();
	$("#outputoptions").show();
	
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
		customreport_loadAddRecords();
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

