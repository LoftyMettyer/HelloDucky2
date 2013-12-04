<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<script type="text/javascript">
	function output_setOptions() {

		var frmOutputDef = OpenHR.getForm("outputoptions", "frmOutputDef");
		var frmExport;

		$("#outputoptions").attr("data-framesource", "OUTPUTOPTIONS");

		if (menu_isSSIMode() == true) {
			frmExport = OpenHR.getForm("reportworkframe", "frmExportData");
		} else {
			frmExport = OpenHR.getForm("reportframe", "frmExportData");
		}

		if (frmExport == null) {
			return;
		}

		var outType = "#optOutputFormat" + frmExport.txtFormat.value;
		var i;

		$(outType)[0].checked = true;
		frmOutputDef.chkDestination0.checked = frmExport.txtScreen;

		if (frmExport.txtPrinter.value.toLowerCase() == "false" && frmExport.txtFormat.value != 0) {
			frmOutputDef.chkDestination1.checked = false;
		} else {
			frmOutputDef.chkDestination1.checked = true;
			populatePrinters();
			for (i = 0; i < frmOutputDef.cboPrinterName.options.length; i++) {
				if (frmOutputDef.cboPrinterName.options[i].innerText == frmExport.txtPrinterName.value) {
					frmOutputDef.cboPrinterName.selectedIndex = i;
					break;
				}
			}
		}

		if (frmExport.txtSave.value.toLowerCase() == "false") {
			frmOutputDef.chkDestination2.checked = false;
		} else {
			frmOutputDef.chkDestination2.checked = true;
			populateSaveExisting();
			frmOutputDef.cboSaveExisting.selectedIndex = frmExport.txtSaveExisting.value;
		}

		if (frmExport.txtEmail.value.toLowerCase() == "false") {
			frmOutputDef.chkDestination3.checked = false;
		} else {
			frmOutputDef.chkDestination3.checked = true;
			frmOutputDef.txtEmailGroupID.value = frmExport.txtEmailAddr.value;
			frmOutputDef.txtEmailGroup.value = frmExport.txtEmailAddrName.value;
			frmOutputDef.txtEmailSubject.value = frmExport.txtEmailSubject.value;
			frmOutputDef.txtEmailAttachAs.value = frmExport.txtEmailAttachAs.value;
		}

		frmOutputDef.txtFilename.value = frmExport.txtFileName.value;
		outputOptionsRefreshControls();

	}

	function outputOptionsFormatClick(index) {
		
		var frmOutputDef = OpenHR.getForm("outputoptions", "frmOutputDef");
		
		frmOutputDef.chkDestination0.checked = false;
		frmOutputDef.chkDestination1.checked = false;
		frmOutputDef.chkDestination2.checked = false;
		frmOutputDef.chkDestination3.checked = false;

		if (index == 1) {
			frmOutputDef.chkDestination2.checked = true;
			frmOutputDef.cboSaveExisting.length = 0;
			frmOutputDef.txtFilename.value = '';
		}
		else if (index == 0) {
			frmOutputDef.chkDestination1.checked = true;
		}
		else {
			frmOutputDef.chkDestination0.checked = true;
		}

		outputOptionsRefreshControls();
	}

	function outputOptionsRefreshControls() {
		
		var frmOutputDef = OpenHR.getForm("outputoptions", "frmOutputDef");
		
		with (frmOutputDef) {
			if (optOutputFormat0.checked == true)		//Data Only
			{
				//disable display on screen options FOR OUTPUT SCREEN ONLY
				chkDestination0.checked = false;
				checkbox_disable(chkDestination0, true);


				//enable-disable printer options
				checkbox_disable(chkDestination1, false);
				if (chkDestination1.checked == true) {
					populatePrinters();
					combo_disable(cboPrinterName, false);
				}
				else {
					cboPrinterName.length = 0;
					combo_disable(cboPrinterName, true);
				}

				//disable save options
				chkDestination2.checked = false;
				checkbox_disable(chkDestination2, true);
				combo_disable(cboSaveExisting, true);
				cboSaveExisting.length = 0;
				txtFilename.value = '';
				text_disable(txtFilename, true);
				button_disable(cmdFilename, true);

				//disable email options
				chkDestination3.checked = false;
				checkbox_disable(chkDestination3, true);
				//text_disable(txtEmailGroup, true);
				txtEmailGroup.value = '';
				txtEmailGroupID.value = 0;
				button_disable(cmdEmailGroup, true);
				text_disable(txtEmailSubject, true);
				text_disable(txtEmailAttachAs, true);
			}
			else if (optOutputFormat1.checked == true)   //CSV File
			{
				//disable display on screen options
				chkDestination0.checked = false;
				checkbox_disable(chkDestination0, true);

				//disable printer options
				chkDestination1.checked = false;
				checkbox_disable(chkDestination1, true);
				cboPrinterName.length = 0;
				combo_disable(cboPrinterName, true);

				//enable-disable save options
				checkbox_disable(chkDestination2, false);
				if (chkDestination2.checked == true) {
					populateSaveExisting();
					combo_disable(cboSaveExisting, false);
					//text_disable(txtFilename, false);
					button_disable(cmdFilename, false);
				}
				else {
					cboSaveExisting.length = 0;
					combo_disable(cboSaveExisting, true);
					//text_disable(txtFilename, true);
					txtFilename.value = '';
					button_disable(cmdFilename, true);
				}

				//enable-disable email options
				checkbox_disable(chkDestination3, false);
				if (chkDestination3.checked == true) {
					//text_disable(txtEmailGroup, false);
					text_disable(txtEmailSubject, false);
					button_disable(cmdEmailGroup, false);
					text_disable(txtEmailAttachAs, false);
				}
				else {
					//text_disable(txtEmailGroup, true);
					txtEmailGroup.value = '';
					txtEmailGroupID.value = 0;
					button_disable(cmdEmailGroup, true);
					text_disable(txtEmailSubject, true);
					text_disable(txtEmailAttachAs, true);
				}
			}
			else if (optOutputFormat2.checked == true)		//HTML Document
			{
				//disable display on screen options
				checkbox_disable(chkDestination0, false);

				//disable printer options
				chkDestination1.checked = false;
				checkbox_disable(chkDestination1, true);
				cboPrinterName.length = 0;
				combo_disable(cboPrinterName, true);

				//enable-disable save options
				checkbox_disable(chkDestination2, false);
				if (chkDestination2.checked == true) {
					populateSaveExisting();
					combo_disable(cboSaveExisting, false);
					//text_disable(txtFilename, false);
					button_disable(cmdFilename, false);
				}
				else {
					cboSaveExisting.length = 0;
					combo_disable(cboSaveExisting, true);
					//text_disable(txtFilename, true);
					txtFilename.value = '';
					button_disable(cmdFilename, true);
				}

				//enable-disable email options
				checkbox_disable(chkDestination3, false);
				if (chkDestination3.checked == true) {
					//text_disable(txtEmailGroup, false);
					text_disable(txtEmailSubject, false);
					button_disable(cmdEmailGroup, false);
					text_disable(txtEmailAttachAs, false);
				}
				else {
					//text_disable(txtEmailGroup, true);
					txtEmailGroup.value = '';
					txtEmailGroupID.value = 0;
					button_disable(cmdEmailGroup, true);
					text_disable(txtEmailSubject, true);
					text_disable(txtEmailAttachAs, true);
				}
			}
			else if (optOutputFormat3.checked == true)		//Word Document
			{
				//enable display on screen options
				checkbox_disable(chkDestination0, false);

				//enable-disable printer options
				checkbox_disable(chkDestination1, false);
				if (chkDestination1.checked == true) {
					populatePrinters();
					combo_disable(cboPrinterName, false);
				}
				else {
					cboPrinterName.length = 0;
					combo_disable(cboPrinterName, true);
				}

				//enable-disable save options
				checkbox_disable(chkDestination2, false);
				if (chkDestination2.checked == true) {
					populateSaveExisting();
					combo_disable(cboSaveExisting, false);
					//text_disable(txtFilename, false);
					button_disable(cmdFilename, false);
				}
				else {
					cboSaveExisting.length = 0;
					combo_disable(cboSaveExisting, true);
					//text_disable(txtFilename, true);
					txtFilename.value = '';
					button_disable(cmdFilename, true);
				}

				//enable-disable email options
				checkbox_disable(chkDestination3, false);
				if (chkDestination3.checked == true) {
					//text_disable(txtEmailGroup, false);
					text_disable(txtEmailSubject, false);
					button_disable(cmdEmailGroup, false);
					text_disable(txtEmailAttachAs, false);
				}
				else {
					//text_disable(txtEmailGroup, true);
					txtEmailGroup.value = '';
					txtEmailGroupID.value = 0;
					button_disable(cmdEmailGroup, true);
					text_disable(txtEmailSubject, true);
					text_disable(txtEmailAttachAs, true);
				}
			}
			else if ((optOutputFormat4.checked == true) ||
					(optOutputFormat5.checked == true) ||
					(optOutputFormat6.checked == true)) {
				//enable display on screen options
				checkbox_disable(chkDestination0, false);

				//enable-disable printer options
				checkbox_disable(chkDestination1, false);
				if (chkDestination1.checked == true) {
					populatePrinters();
					combo_disable(cboPrinterName, false);
				}
				else {
					cboPrinterName.length = 0;
					combo_disable(cboPrinterName, true);
				}

				//enable-disable save options
				checkbox_disable(chkDestination2, false);
				if (chkDestination2.checked == true) {
					populateSaveExisting();
					combo_disable(cboSaveExisting, false);
					//text_disable(txtFilename, false);
					button_disable(cmdFilename, false);
				}
				else {
					cboSaveExisting.length = 0;
					combo_disable(cboSaveExisting, true);
					//text_disable(txtFilename, true);
					txtFilename.value = '';
					button_disable(cmdFilename, true);
				}

				//enable-disable email options
				checkbox_disable(chkDestination3, false);
				if (chkDestination3.checked == true) {
					//text_disable(txtEmailGroup, false);
					text_disable(txtEmailSubject, false);
					button_disable(cmdEmailGroup, false);
					text_disable(txtEmailAttachAs, false);
				}
				else {
					//text_disable(txtEmailGroup, true);
					txtEmailGroup.value = '';
					txtEmailGroupID.value = 0;
					button_disable(cmdEmailGroup, true);
					text_disable(txtEmailSubject, true);
					text_disable(txtEmailAttachAs, true);
				}
			}
			else {
				optOutputFormat0.checked = true;
				outputOptionsRefreshControls();
			}

			if (txtEmailSubject.disabled) {
				txtEmailSubject.value = '';
			}

			if (txtEmailAttachAs.disabled) {
				txtEmailAttachAs.value = '';
			}
			else {

				if (txtEmailAttachAs.value == '') {
					if (txtFilename.value != '') {
						sAttachmentName = new String(txtFilename.value);
						txtEmailAttachAs.value = sAttachmentName.substr(sAttachmentName.lastIndexOf("\\") + 1);
					}
				}
			}

			if (cmdFilename.disabled == true) {
				txtFilename.value = "";
			}
		}

	}

	function populatePrinters() {

		var frmOutputDef = OpenHR.getForm("outputoptions", "frmOutputDef");
		var oOption;

		var strCurrentPrinter = '';
		if (frmOutputDef.cboPrinterName.selectedIndex > 0) {
			strCurrentPrinter = options[frmOutputDef.cboPrinterName.selectedIndex].innerText;
		}

		oOption = document.createElement("OPTION");
		frmOutputDef.cboPrinterName.options.add(oOption);
		oOption.innerHTML = "<Default Printer>";
		oOption.value = 0;

		for (var iLoop = 0; iLoop < OpenHR.PrinterCount() ; iLoop++) {
			oOption = document.createElement("OPTION");
			frmOutputDef.cboPrinterName.options.add(oOption);
			oOption.innerHTML = OpenHR.PrinterName(iLoop);
			oOption.value = iLoop + 1;

			if (oOption.innerText == strCurrentPrinter) {
				frmOutputDef.cboPrinterName.selectedIndex = iLoop + 1;
			}
		}
	}

	function populateSaveExisting() {

		var frmOutputDef = OpenHR.getForm("outputoptions", "frmOutputDef");
		var oOption;

		var lngCurrentOption = 0;
		var selectedIndex = frmOutputDef.cboSaveExisting.selectedIndex;

		frmOutputDef.cboSaveExisting.length = 0;
		
		if (selectedIndex > 0) {
			lngCurrentOption = options[selectedIndex].value;
		}

		oOption = document.createElement("OPTION");
		frmOutputDef.cboSaveExisting.options.add(oOption);
		oOption.innerHTML = "Overwrite";
		oOption.value = 0;

		oOption = document.createElement("OPTION");
		frmOutputDef.cboSaveExisting.options.add(oOption);
		oOption.innerHTML = "Do not overwrite";
		oOption.value = 1;

		oOption = document.createElement("OPTION");
		frmOutputDef.cboSaveExisting.options.add(oOption);
		oOption.innerHTML = "Add sequential number to name";
		oOption.value = 2;

		oOption = document.createElement("OPTION");
		frmOutputDef.cboSaveExisting.options.add(oOption);
		oOption.innerHTML = "Append to file";
		oOption.value = 3;

		if ((frmOutputDef.optOutputFormat4.checked) ||
				(frmOutputDef.optOutputFormat5.checked) ||
				(frmOutputDef.optOutputFormat6.checked)) {
			oOption = document.createElement("OPTION");
			frmOutputDef.cboSaveExisting.options.add(oOption);
			oOption.innerHTML = "Create new sheet in workbook";
			oOption.value = 4;
		}

		for (var iLoop = 0; iLoop < frmOutputDef.cboSaveExisting.options.length; iLoop++) {
			if (frmOutputDef.cboSaveExisting.options[iLoop].value == lngCurrentOption) {
				frmOutputDef.cboSaveExisting.selectedIndex = iLoop;
				break;
			}
		}
	}

	function selectEmailGroup() {
		var sURL;

		var frmOutputDef = OpenHR.getForm("outputoptions", "frmOutputDef");
		frmEmailSelection.EmailSelCurrentID.value = frmOutputDef.txtEmailGroupID.value;

		sURL = "util_emailSelection" +
				"?EmailSelCurrentID=" + frmEmailSelection.EmailSelCurrentID.value;
		openDialog(sURL, (screen.width) / 3, (screen.height) / 2, "yes", "yes");
	}

	function outputOptionsOKClick() {

		var frmOutputDef = OpenHR.getForm("outputoptions", "frmOutputDef");

		if ((frmOutputDef.chkDestination0.checked == false) &&
				(frmOutputDef.chkDestination1.checked == false) &&
				(frmOutputDef.chkDestination2.checked == false) &&
				(frmOutputDef.chkDestination3.checked == false)) {
			OpenHR.messageBox("You must select a destination", 48, "Output Options");
			window.focus();
			return;
		}

		var sAttachmentName = new String(frmOutputDef.txtEmailAttachAs.value);
		if ((sAttachmentName.indexOf("/") != -1) ||
				(sAttachmentName.indexOf(":") != -1) ||
				(sAttachmentName.indexOf("?") != -1) ||
				(sAttachmentName.indexOf(String.fromCharCode(34)) != -1) ||
				(sAttachmentName.indexOf("<") != -1) ||
				(sAttachmentName.indexOf(">") != -1) ||
				(sAttachmentName.indexOf("|") != -1) ||
				(sAttachmentName.indexOf("\\") != -1) ||
				(sAttachmentName.indexOf("*") != -1)) {
			OpenHR.messageBox("The attachment file name can not contain any of the following characters:\n/ : ? " + String.fromCharCode(34) + " < > | \\ *", 48, "Output Options");
			window.focus();
			return;
		}

		if ((frmOutputDef.txtFilename.value == "")
				&& (frmOutputDef.cmdFilename.disabled == false)) {
			OpenHR.messageBox("You must enter a file name", 48, "Output Options");
			window.focus();
			return;
		}

		if ((frmOutputDef.txtEmailGroup.value == "")
				&& (frmOutputDef.cmdEmailGroup.disabled == false)) {
			OpenHR.messageBox("You must select an email group", 48, "Output Options");
			window.focus();
			return;
		}

		if ((frmOutputDef.chkDestination3.checked)
				&& (frmOutputDef.txtEmailAttachAs.value == '')) {
			OpenHR.messageBox("You must enter an email attachment file name.", 48, "Output Options");
			window.focus();
			return;
		}

		window.ShowWaitFrame("Outputting...");

		//  The doExport function is where it all continues
		window.setTimeout('doExport()', 1000);
	}

	function doExport() {
		
		//Send the values back to the calling form...
		var frmOutputDef = OpenHR.getForm("outputoptions", "frmOutputDef");
		var frmExportData = OpenHR.getForm("main", "frmExportData");

		frmExportData.txtFormat.value = 0;
		if (frmOutputDef.optOutputFormat1.checked == true) { frmExportData.txtFormat.value = 1; }

		//CSV
		if (frmOutputDef.optOutputFormat2.checked == true) { frmExportData.txtFormat.value = 2; }

		//HTML
		if (frmOutputDef.optOutputFormat3.checked == true) { frmExportData.txtFormat.value = 3; }

		//WORD
		if (frmOutputDef.optOutputFormat4.checked == true) { frmExportData.txtFormat.value = 4; }

		//EXCEL
		if (frmOutputDef.optOutputFormat5.checked == true) { frmExportData.txtFormat.value = 5; }

		//GRAPH
		if (frmOutputDef.optOutputFormat6.checked == true) { frmExportData.txtFormat.value = 6; }

		//PIVOT

		frmExportData.txtScreen.value = frmOutputDef.chkDestination0.checked;

		frmExportData.txtPrinter.value = frmOutputDef.chkDestination1.checked;
		frmExportData.txtPrinterName.value = '';
		if (frmOutputDef.cboPrinterName.selectedIndex != -1) {
			frmExportData.txtPrinterName.value = frmOutputDef.cboPrinterName.options[frmOutputDef.cboPrinterName.selectedIndex].innerText;
		}

		frmExportData.txtSave.value = frmOutputDef.chkDestination2.checked;
		frmExportData.txtSaveExisting.value = frmOutputDef.cboSaveExisting.selectedIndex;
		frmExportData.txtEmail.value = frmOutputDef.chkDestination3.checked;
		frmExportData.txtEmailAddr.value = frmOutputDef.txtEmailGroupID.value;
		frmExportData.txtEmailAddrName.value = frmOutputDef.txtEmailGroup.value;
		frmExportData.txtEmailSubject.value = frmOutputDef.txtEmailSubject.value;
		frmExportData.txtEmailAttachAs.value = frmOutputDef.txtEmailAttachAs.value;
		frmExportData.txtFileName.value = frmOutputDef.txtFilename.value;

		var frmGetDataForm = OpenHR.getForm("reportworkframe", "frmExportData");

		if (frmOutputDef.txtEmailGroupID.value > 0) {
			if (frmOutputDef.txtUtilType.value == 17) {
				frmGetDataForm.txtEmailGroupID.value = frmOutputDef.txtEmailGroupID.value;
				ExportData("OUTPUTRUN");
			}
			else {
				frmGetDataForm.txtMode.value = "EMAILGROUP";
				frmGetDataForm.txtEmailGroupID.value = frmOutputDef.txtEmailGroupID.value;
				frmGetDataForm.submit();
			}
		}
		else {
			frmGetDataForm.txtEmailGroupID.value = 0;
			frmGetDataForm.submit();
		}

		if (frmOutputDef.txtUtilType.value == 2) {
			ShowDataFrame();
		}
	}

	function saveFile() {
		return false;
	}

</script>



<form id="frmOutputDef" name="frmOutputDef">

	<table WIDTH="100%" class="invisible" CELLSPACING=10 CELLPADDING=0>
		<tr>						
			<td valign=top rowspan=2 width=25% height="100%">
				<table  cellspacing="0" cellpadding="4" width="100%" height="100%">
					<tr height=10> 
						<td height=10 align=left valign=top>
							<strong>Output Format : </strong><BR><BR>
							<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
								<tr height=20>
									<td width=5>&nbsp</td>
									<td align=left width=15>
									<input type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat0 value=0 onClick="outputOptionsFormatClick(0);" />
									</td>
									<td align=left nowrap>
												<label 
														tabindex=-1
														for="optOutputFormat0"
														class="radio" />
										Data Only
																					
									<td width=5>&nbsp</td>
								</tr>
								<tr height=10> 
									<td colspan=4></td>
								</tr>
<% if Session("utilType") <> 17 and Session("utilType") <> 16 then %>																	
								<tr height=20>
									<td width=5>&nbsp</td>
									<td align=left width=15>
									<input type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat1 value=1 onClick="outputOptionsFormatClick(1);" />
									</td>
									<td align=left nowrap>
												<label 
														tabindex=-1
														for="optOutputFormat1"
														class="radio"/>
										CSV File
																					
									<td width=5>&nbsp</td>
								</tr>
								<tr height=10> 
									<td colspan=4></td>
								</tr>
<% end if %>
								<tr height=20>
									<td width=5>&nbsp</td>
									<td align=left width=15>																		
									<input type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat2 value=2 onClick="outputOptionsFormatClick(2);" />
									</td>
									<td align=left nowrap>
												<label 
														tabindex=-1
														for="optOutputFormat2"
														class="radio" />
										HTML Document
																					
									<td width=5>&nbsp</td>
								</tr>
								<tr height=10> 
									<td colspan=4></td>
								</tr>
								<tr height=20>
									<td width=5>&nbsp</td>
									<td align=left width=15>
									<input type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat3 value=3 onClick="outputOptionsFormatClick(3);" />
									</td>
									<td align=left nowrap>
												<label 
														tabindex=-1
														for="optOutputFormat3"
														class="radio" />
										Word Document
									<td width=5>&nbsp</td>
								</tr>
								<tr height=10> 
									<td colspan=4></td>
								</tr>
								<tr height=20>
									<td width=5>&nbsp</td>
									<td align=left width=15>
									<input type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat4 value=4 onClick="outputOptionsFormatClick(4);" />
									</td>
									<td align=left nowrap>
												<label 
														tabindex=-1
														for="optOutputFormat4"
														class="radio"/>
										Excel Worksheet
																					
									<td width=5>&nbsp</td>
								</tr>
								<tr height=10> 
									<td colspan=4></td>
								</tr>
																	
<% if Session("utilType") = 17 then %>																	
								<tr height=5>
									<td width=5>&nbsp</td>
									<td align=left width=15>
										<input DISABLED type=radio width=20 style="WIDTH: 20px; visibility: hidden" name=optOutputFormat id=optOutputFormat5 value=5>
									</td>
									<td>
																			
									</td>
									<td width=5>&nbsp</td>
								</tr>
								<tr height=10> 
									<td colspan=4></td>
								</tr>
								<tr height=5>
									<td width=5>&nbsp</td>
									<td align=left width=15>
										<input DISABLED type=radio width=20 style="WIDTH: 20px; visibility: hidden" name=optOutputFormat id=optOutputFormat6 value=6>
									</td>
									<td>
																			
									</td>
									<td width=5>&nbsp</td>
								</tr>
								<tr height=5> 
									<td colspan=4></td>
								</tr>		
								<tr height=20>
									<td width=5>&nbsp</td>
									<td align=left width=15>
										<input DISABLED type=radio width=20 style="WIDTH: 20px; visibility: hidden" name=optOutputFormat id=optOutputFormat1 value=1>
									</td>
									<td align=left nowrap>
									</td>
									<td width=5>&nbsp</td>
								</tr>
								<tr height=10> 
									<td colspan=4></td>
								</tr>
<% elseif Session("utilType") = 16 then %>																	
								<tr height=5>
									<td width=5>&nbsp</td>
									<td align=left width=15>
									<input type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat5 value=5 onClick="outputOptionsFormatClick(5);" />
									</td>
									<td align=left nowrap>
												<label 
														tabindex=-1
														for="optOutputFormat5"
														class="radio"/>
										Excel Chart
									<td width=5>&nbsp</td>
								</tr>
								<tr height=10> 
									<td colspan=4></td>
								</tr>
								<tr height=5>
									<td width=5>&nbsp</td>
									<td align=left width=15>
										<input DISABLED type=radio width=20 style="WIDTH: 20px; visibility: hidden" name=optOutputFormat id=optOutputFormat6 value=6>
									</td>
									<td>
																			
									</td>
									<td width=5>&nbsp</td>
								</tr>
								<tr height=5> 
									<td colspan=4></td>
								</tr>		
								<tr height=20>
									<td width=5>&nbsp</td>
									<td align=left width=15>
										<input DISABLED type=radio width=20 style="WIDTH: 20px; visibility: hidden" name=optOutputFormat id=optOutputFormat1 value=1>
									</td>
									<td align=left nowrap>
									</td>
									<td width=5>&nbsp</td>
								</tr>
								<tr height=10> 
									<td colspan=4></td>
								</tr>
																									
<% else %>
								<tr height=5>
									<td width=5>&nbsp</td>
									<td align=left width=15>
									<input type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat5 value=5 onClick="outputOptionsFormatClick(5);" />
									</td>
									<td align=left nowrap>
												<label 
														tabindex=-1
														for="optOutputFormat5"
														class="radio" />
										Excel Chart
									<td width=5>&nbsp</td>
								</tr>
								<tr height=10> 
									<td colspan=4></td>
								</tr>
								<tr height=5>
									<td width=5>&nbsp</td>
									<td align=left width=15>																		
									<input type=radio width=20 style="WIDTH: 20px" name=optOutputFormat id=optOutputFormat6 value=6 onClick="outputOptionsFormatClick(6);" />
									</td>
									<td align=left nowrap>
												<label 
														tabindex=-1
														for="optOutputFormat6"
														class="radio" />
										Excel Pivot Table
															
									<td width=5>&nbsp</td>
								</tr>
								<tr height=5> 
									<td colspan=4></td>
								</tr>										
<% end if %>														
							</table>
						</td>
					</tr>
				</table>
			</td>
			<td valign=top width="75%">
				<table cellspacing="0" cellpadding="4" width="100%" height="100%">
					<tr height=10> 
						<td height=10 align=left valign=top>
							<strong>Output Destination(s) : </strong><BR><BR>
							<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
								<tr height=20>
									<td width=5>&nbsp</td>
									<td align=left colspan=6 nowrap>
									<input name=chkDestination0 id=chkDestination0 type=checkbox disabled="disabled" tabindex="0" onClick="outputOptionsRefreshControls();"/>
										<label 
											for="chkDestination0"
											class="checkbox"
											tabindex="-1" />																	
									Display output on screen 
																		
									</td>
									<td width=5>&nbsp</td>
								</tr>
								<tr height=10> 
									<td colspan=8></td>
								</tr>
								<tr height=20>
									<td width=5>&nbsp</td>
									<td align=left nowrap>																	
									<input name=chkDestination1 id=chkDestination1 type=checkbox disabled="disabled" tabindex="0" onClick="outputOptionsRefreshControls();"/>
										<label 
											for="chkDestination1"
											class="checkbox"
											tabindex="-1" />																		
									Send to printer 
																																				
									</td>
									<td width=30 nowrap>&nbsp</td>
									<td align=left nowrap>
										Printer location : 
									</td>
									<td width=15>&nbsp</td>
									<td colspan=2>
										<select id="cboPrinterName" name="cboPrinterName" class="combo" width="100%" style="WIDTH: 100%" />
									</td>
									<td width=5>&nbsp</td>
								</tr>
								<tr height=10> 
									<td colspan=8></td>
								</tr>
								<tr height=20>
									<td width=5>&nbsp</td>
									<td align=left nowrap>
									<input name=chkDestination2 id=chkDestination2 type=checkbox disabled="disabled" tabindex="0" onClick="outputOptionsRefreshControls();"/>
										<label 
											for="chkDestination2"
											class="checkbox"
											tabindex="-1" />																			
									Save to file 
																																				
									</td>
									<td width=30 nowrap>&nbsp</td>
									<td align=left nowrap>
										File name :
									</td>
									<td width=15 nowrap>&nbsp</td>
									<td colspan=2>
										<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
											<tr>
												<td>
													<input id="txtFilename" name="txtFilename" class="text textdisabled" disabled="disabled" tabindex="-1" style="WIDTH: 100%">
												</td>
												<td width="25">
													<input type="button" id="cmdFilename" name="cmdFilename" value="..." style="WIDTH: 100%" class="btn" onclick="saveFile()" />
												</td>
											</tr>
										</table>
									</td>
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
										<select id=cboSaveExisting name=cboSaveExisting class="combo" width=100% style="WIDTH: 100%" />																											
									</td>
									<td width=5>&nbsp</td>
								</tr>

								<tr height=10> 
									<td colspan=8></td>
								</tr>
								<tr height=20>
									<td width=5>&nbsp</td>
									<td align=left nowrap>
									<input name=chkDestination3 id=chkDestination3 type=checkbox disabled="disabled" tabindex="0" onClick="outputOptionsRefreshControls();"/>
									<label 
												for="chkDestination3"
												class="checkbox"
												tabindex="-1">
																			
									Send as email
									</label>																		
									</td>
									<td width=30 nowrap>&nbsp</td>
									<td align=left nowrap>
										Email group :   
									</td>
									<td width=15 nowrap>&nbsp</td>
									<td colspan="2">
										<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
											<tr>
												<td>
													<input id="txtEmailGroup" name="txtEmailGroup" class="text textdisabled" disabled="disabled" style="WIDTH: 100%">
													<input id="txtEmailGroupID" name="txtEmailGroupID" type="hidden">
												</td>
												<td width="25">
													<input type="button" id="cmdEmailGroup" name="cmdEmailGroup" value="..." style="WIDTH: 100%" class="btn" onclick="selectEmailGroup()" />
												</td>
											</tr>
										</table>
									</td>
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
									<td colspan="2" width="100%" nowrap>
										<input id="txtEmailSubject" class="text textdisabled" disabled="disabled" maxlength="255" name="txtEmailSubject" style="WIDTH: 100%">
									</td>
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
									<td colspan="2" width="100%" nowrap>
										<input id="txtEmailAttachAs" class="text textdisabled" disabled="disabled" maxlength="255" name="txtEmailAttachAs" style="WIDTH: 100%">
									</td>
									<td width=5>&nbsp</td>
								</tr>
								<tr height=10> 
									<td colspan=8></td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>

	<input type="hidden" id="txtDatabase" name="txtDatabase" value="<%=Session("Database")%>">
	<input type="hidden" id="txtUtilType" name="txtUtilType" value="<%=Session("utilType")%>">
	<input type="hidden" id="txtWordVer" name="txtWordVer" value="<%=Session("WordVer")%>">
	<input type="hidden" id="txtExcelVer" name="txtExcelVer" value="<%=Session("ExcelVer")%>">
	<input type="hidden" id="txtWordFormats" name="txtWordFormats" value="<%=Session("WordFormats")%>">
	<input type="hidden" id="txtExcelFormats" name="txtExcelFormats" value="<%=Session("ExcelFormats")%>">
	<input type="hidden" id="txtWordFormatDefaultIndex" name="txtWordFormatDefaultIndex" value="<%=Session("WordFormatDefaultIndex")%>">
	<input type="hidden" id="txtExcelFormatDefaultIndex" name="txtExcelFormatDefaultIndex" value="<%=Session("ExcelFormatDefaultIndex")%>">
	<input type="hidden" id="txtOfficeSaveAsFormats" name="txtOfficeSaveAsFormats" value="<%=Session("OfficeSaveAsValues")%>">
</form>

<form id="frmEmailSelection" name="frmEmailSelection" target="emailSelection" action="util_emailSelection" method="post" style="visibility: hidden; display: none">
	<input type="hidden" id="EmailSelCurrentID" name="EmailSelCurrentID">
</form>


<script type="text/javascript">
	output_setOptions();
</script>