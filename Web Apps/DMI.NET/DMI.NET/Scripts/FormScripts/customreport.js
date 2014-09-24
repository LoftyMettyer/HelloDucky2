function replace(sExpression, sFind, sReplace) {
		//gi (global search, ignore case)
		var re = new RegExp(sFind, "gi");
		sExpression = sExpression.replace(re, sReplace);
		return (sExpression);
}

function dataOnlyPrint() {
		// PageHeaderFont and PageFooterFont don't function in ASPs.
		//  frmOutput.ssOleDBGrid.PageHeaderFont.Name = "Verdana";
		//  frmOutput.ssOleDBGrid.PageHeaderFont.Size = 12;
		//  frmOutput.ssOleDBGrid.PageHeaderFont.Bold = true;
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

