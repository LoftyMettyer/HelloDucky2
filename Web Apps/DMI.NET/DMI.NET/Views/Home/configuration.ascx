<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<%
	Dim sTemp As String
	
	Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
		
	' Get the DefSel 'only mine' settings.
	For i = 0 To 21
		sTemp = "onlymine "

		Select Case i
			Case 0
				sTemp = sTemp & "BatchJobs"
			Case 1
				sTemp = sTemp & "Calculations"
			Case 2
				sTemp = sTemp & "CrossTabs"
			Case 3
				sTemp = sTemp & "CustomReports"
			Case 4
				sTemp = sTemp & "DataTransfer"
			Case 5
				sTemp = sTemp & "Export"
			Case 6
				sTemp = sTemp & "Filters"
			Case 7
				sTemp = sTemp & "GlobalAdd"
			Case 8
				sTemp = sTemp & "GlobalUpdate"
			Case 9
				sTemp = sTemp & "GlobalDelete"
			Case 10
				sTemp = sTemp & "Import"
			Case 11
				sTemp = sTemp & "MailMerge"
			Case 12
				sTemp = sTemp & "Picklists"
			Case 13
				sTemp = sTemp & "CalendarReports"
			Case 14
				sTemp = sTemp & "Labels"
			Case 15
				sTemp = sTemp & "LabelDefinition"
			Case 16
				sTemp = sTemp & "MatchReports"
			Case 17
				sTemp = sTemp & "CareerProgression"
			Case 18
				sTemp = sTemp & "EmailGroups"
			Case 19
				sTemp = sTemp & "RecordProfile"
			Case 20
				sTemp = sTemp & "SuccessionPlanning"
			Case 21
				sTemp = sTemp & "NineBoxGrid"
		End Select

		Session(sTemp) = CLng(objDatabase.GetUserSetting("defsel", sTemp, 0))

	Next

	' Get the Utility Warning settings.
	For i = 0 To 4
		sTemp = "warning "

		Select Case i
			Case 0
				sTemp = sTemp & "DataTransfer"
			Case 1
				sTemp = sTemp & "GlobalAdd"
			Case 2
				sTemp = sTemp & "GlobalUpdate"
			Case 3
				sTemp = sTemp & "GlobalDelete"
			Case 4
				sTemp = sTemp & "Import"
		End Select
			
		Session(sTemp) = CLng(objDatabase.GetUserSetting("warningmsg", sTemp, 1))
		
	Next
%>

<script type="text/javascript">
	function configuration_window_onload() {
		////        var frmOriginalConfiguration = OpenHR.getForm("workframe", "frmOriginalConfiguration");
		var frmMenu = $("#frmMenuInfo")[0].children;
		$("#workframe").attr("data-framesource", "CONFIGURATION");
		showDefaultRibbon();

		//// Get menu to refresh the menu.
		menu_refreshMenu();

		// Load the original values into tab 1.
		setComboValue("PARENT", frmOriginalConfiguration.txtPrimaryStartMode.value);
		setComboValue("HISTORY", frmOriginalConfiguration.txtHistoryStartMode.value);
		setComboValue("LOOKUP", frmOriginalConfiguration.txtLookupStartMode.value);
		setComboValue("QUICKACCESS", frmOriginalConfiguration.txtQuickAccessStartMode.value);
		setComboValue("EXPRCOLOURMODE", frmOriginalConfiguration.txtExprColourMode.value);
		setComboValue("EXPRNODEMODE", frmOriginalConfiguration.txtExprNodeMode.value);
		frmConfiguration.txtFindSize.value = frmOriginalConfiguration.txtFindSize.value;

		// Load the original values into tab 2. 
		frmConfiguration.chkOwner_Calculations.checked = (frmOriginalConfiguration.txtOnlyMineCalculations.value == 1);
		frmConfiguration.chkOwner_CrossTabs.checked = (frmOriginalConfiguration.txtOnlyMineCrossTabs.value == 1);
		frmConfiguration.chkOwner_NineBoxGrid.checked = (frmOriginalConfiguration.txtOnlyMineNineBoxGrid.value == 1);
		frmConfiguration.chkOwner_CustomReports.checked = (frmOriginalConfiguration.txtOnlyMineCustomReports.value == 1);
		frmConfiguration.chkOwner_Filters.checked = (frmOriginalConfiguration.txtOnlyMineFilters.value == 1);
		frmConfiguration.chkOwner_MailMerge.checked = (frmOriginalConfiguration.txtOnlyMineMailMerge.value == 1);
		frmConfiguration.chkOwner_Picklists.checked = (frmOriginalConfiguration.txtOnlyMinePicklists.value == 1);
		frmConfiguration.chkOwner_CalendarReports.checked = (frmOriginalConfiguration.txtOnlyMineCalendarReports.value == 1);

		display_Configuration_Page(1);

		menu_setVisibleMenuItem("mnutoolSaveAdminConfig", true);
		menu_toolbarEnableItem('mnutoolSaveAdminConfig', (!definitionChanged() == false))
		// $('#mnutoolSaveAdminConfig').click('okClick()');

		$("#toolbarAdminConfig").parent().show();
		$("#toolbarAdminConfig").click();
		//$('input[name^="txt"]').on("blur", function () { enableSaveButton(); });
		$('input[name^="txt"]').on("input", function () { enableSaveButton(this); });
		$('select[name^="cbo"]').on("change", function () { enableSaveButton(); });
		$('input[name^="chk"]').on("change", function () { enableSaveButton(); });
		$("#optionframe").hide();
		$("#workframe").show();
	}

	function enableSaveButton() {
		if (definitionChanged()) menu_toolbarEnableItem('mnutoolSaveAdminConfig', true);
	}
</script>

<script type="text/javascript">
	function display_Configuration_Page(piPageNumber) {
		if (piPageNumber == 1) {
			$("#div1").css("visibility", "visible").css("display", "block");
			$("#div2").css("visibility", "hidden").css("display", "none");
			$('#btnDiv1OK').hide();
			$('#btnDiv1Cancel').hide();

			frmConfiguration.cboPrimaryTableDisplay.focus();
		}

		if (piPageNumber == 2) {
			$("#div1").css("visibility", "hidden").css("display", "none");
			$("#div2").css("visibility", "visible").css("display", "block");
			$('#btnDiv2OK').hide();
			$('#btnDiv2Cancel').hide();
		}

		var styles = { 'border-spacing': 0, 'border-collapse': 'separate' };
		$('table').css(styles);
		$('td').css('padding', '0');
		$('select, #txtFindSize').css('float', 'right');
		$('select').css('width', '200px');
		$('#txtFindSize').css('width', '100px');

		//$('table').attr('border', '1');
	}

	function setComboValue(psCombo, piValue) {
		var i;
		var cboCombo;

		if (psCombo == "PARENT") {
			cboCombo = frmConfiguration.cboPrimaryTableDisplay;
		}
		if (psCombo == "HISTORY") {
			cboCombo = frmConfiguration.cboHistoryTableDisplay;
		}
		if (psCombo == "LOOKUP") {
			cboCombo = frmConfiguration.cboLookupTableDisplay;
		}
		if (psCombo == "QUICKACCESS") {
			cboCombo = frmConfiguration.cboQuickAccessDisplay;
		}
		if (psCombo == "EXPRCOLOURMODE") {
			cboCombo = frmConfiguration.cboViewInColour;
		}
		if (psCombo == "EXPRNODEMODE") {
			cboCombo = frmConfiguration.cboExpandNodes;
		}

		for (i = 0; i < cboCombo.options.length; i++) {
			if (cboCombo.options[i].value == piValue) {
				cboCombo.selectedIndex = i;
				return;
			}
		}

		cboCombo.selectedIndex = 0;
	}

	function saveConfiguration() {
		var chkControl;
		var txtControl;
		var sType;
		var frmConfiguration = OpenHR.getForm("workframe", "frmConfiguration");
		// Validate the find window block size.
		if (validateFindBlockSize() == false) {
			return false;
		}

		frmConfiguration.txtPrimaryStartMode.value = frmConfiguration.cboPrimaryTableDisplay.options[frmConfiguration.cboPrimaryTableDisplay.options.selectedIndex].value;
		frmConfiguration.txtHistoryStartMode.value = frmConfiguration.cboHistoryTableDisplay.options[frmConfiguration.cboHistoryTableDisplay.options.selectedIndex].value;
		frmConfiguration.txtLookupStartMode.value = frmConfiguration.cboLookupTableDisplay.options[frmConfiguration.cboLookupTableDisplay.options.selectedIndex].value;
		frmConfiguration.txtQuickAccessStartMode.value = frmConfiguration.cboQuickAccessDisplay.options[frmConfiguration.cboQuickAccessDisplay.options.selectedIndex].value;
		frmConfiguration.txtExprColourMode.value = frmConfiguration.cboViewInColour.options[frmConfiguration.cboViewInColour.options.selectedIndex].value;
		frmConfiguration.txtExprNodeMode.value = frmConfiguration.cboExpandNodes.options[frmConfiguration.cboExpandNodes.options.selectedIndex].value;

		menu_refreshMenu();
		var menuForm = $("#frmMenuInfo")[0].children;
		menuForm.txtPrimaryStartMode.value = frmConfiguration.txtPrimaryStartMode.value;
		menuForm.txtHistoryStartMode.value = frmConfiguration.txtHistoryStartMode.value;
		menuForm.txtLookupStartMode.value = frmConfiguration.txtLookupStartMode.value;
		menuForm.txtQuickAccessStartMode.value = frmConfiguration.txtQuickAccessStartMode.value;

		if (frmConfiguration.chkOwner_Calculations.checked == true) frmConfiguration.txtOwner_Calculations.value = 1;
		if (frmConfiguration.chkOwner_CrossTabs.checked == true) frmConfiguration.txtOwner_CrossTabs.value = 1;
		if (frmConfiguration.chkOwner_NineBoxGrid.checked == true) frmConfiguration.txtOwner_NineBoxGrid.value = 1;
		if (frmConfiguration.chkOwner_CustomReports.checked == true) frmConfiguration.txtOwner_CustomReports.value = 1;
		if (frmConfiguration.chkOwner_Filters.checked == true) frmConfiguration.txtOwner_Filters.value = 1;
		if (frmConfiguration.chkOwner_MailMerge.checked == true) frmConfiguration.txtOwner_MailMerge.value = 1;
		if (frmConfiguration.chkOwner_Picklists.checked == true) frmConfiguration.txtOwner_Picklists.value = 1;
		if (frmConfiguration.chkOwner_CalendarReports.checked == true) frmConfiguration.txtOwner_CalendarReports.value = 1;

		OpenHR.submitForm(frmConfiguration);

	}

	function validateFindBlockSize() {
		var sConvertedFindSize;
		var sDecimalSeparator;
		var sThousandSeparator;
		var sPoint;
		var iValue;

		sDecimalSeparator = "\\";
		sDecimalSeparator = sDecimalSeparator.concat(OpenHR.LocaleDecimalSeparator());
		var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");

		sThousandSeparator = "\\";
		sThousandSeparator = sThousandSeparator.concat(OpenHR.LocaleThousandSeparator());
		var reThousandSeparator = new RegExp(sThousandSeparator, "gi");

		sPoint = "\\.";
		var rePoint = new RegExp(sPoint, "gi");

		if (frmConfiguration.txtFindSize.value == '') {
			frmConfiguration.txtFindSize.value = 0;
		}

		// Convert the find size value from locale to UK settings for use with the isNaN funtion.
		sConvertedFindSize = new String(frmConfiguration.txtFindSize.value);
		// Remove any thousand separators.
		sConvertedFindSize = sConvertedFindSize.replace(reThousandSeparator, "");
		frmConfiguration.txtFindSize.value = sConvertedFindSize;

		// Convert any decimal separators to '.'.
		if (OpenHR.LocaleDecimalSeparator() != ".") {
			// Remove decimal points.
			sConvertedFindSize = sConvertedFindSize.replace(rePoint, "A");
			// replace the locale decimal marker with the decimal point.
			sConvertedFindSize = sConvertedFindSize.replace(reDecimalSeparator, ".");
		}

		if (isNaN(sConvertedFindSize) == true) {
			OpenHR.messageBox("Find window block size must be numeric.");
			frmConfiguration.txtFindSize.value = frmOriginalConfiguration.txtLastFindSize.value;
			display_Configuration_Page(1);
			frmConfiguration.txtFindSize.focus();
			return false;
		}

		if (frmConfiguration.txtFindSize.value <= 0) {
			OpenHR.messageBox("Find window block size must be greater than 0.");
			frmConfiguration.txtFindSize.value = frmOriginalConfiguration.txtLastFindSize.value;
			display_Configuration_Page(1);
			frmConfiguration.txtFindSize.focus();
			return false;
		}

		// Find size must be integer.		
		if (sConvertedFindSize.indexOf(".") >= 0) {
			OpenHR.messageBox("Find window block size must be an integer value.");
			frmConfiguration.txtFindSize.value = frmOriginalConfiguration.txtLastFindSize.value;
			display_Configuration_Page(1);
			frmConfiguration.txtFindSize.focus();
			return false;
		}

		iValue = new Number(frmConfiguration.txtFindSize.value);
		if (iValue > 100000) {
			OpenHR.messageBox("Find window block size cannot be greater than 100000.");
			frmConfiguration.txtFindSize.value = "100000";
			display_Configuration_Page(1);
			frmConfiguration.txtFindSize.focus();
			return false;
		}

		frmOriginalConfiguration.txtLastFindSize.value = frmConfiguration.txtFindSize.value;

		return true;
	}

	function Configuration_okClick() {
		frmConfiguration.txtReaction.value = "DEFAULT";
		saveConfiguration();
	}

	/* Return to the default page. */
	function cancelClick() {

	}

	function saveChanges(psAction, pfPrompt, pfTBOverride) {
		if (definitionChanged() == false) {
			return 6; //No to saving the changes, as none have been made.
		} else
			return 0;
	}

	function definitionChanged() {
		//In certain circumstances frmConfiguration is not defined, so check first if it's defined before attempting to use it
		try {
			if (frmConfiguration == undefined) {
			}
		} catch (e) {
			return false; //i.e. nothing has changed
		}

		// Compare the tab 1 controls with the original values.
		if (frmConfiguration.cboPrimaryTableDisplay.options[frmConfiguration.cboPrimaryTableDisplay.selectedIndex].value != frmOriginalConfiguration.txtPrimaryStartMode.value) {
			return true;
		}

		if (frmConfiguration.cboHistoryTableDisplay.options[frmConfiguration.cboHistoryTableDisplay.selectedIndex].value != frmOriginalConfiguration.txtHistoryStartMode.value) {
			return true;
		}

		if (frmConfiguration.cboLookupTableDisplay.options[frmConfiguration.cboLookupTableDisplay.selectedIndex].value != frmOriginalConfiguration.txtLookupStartMode.value) {
			return true;
		}
		if (frmConfiguration.cboQuickAccessDisplay.options[frmConfiguration.cboQuickAccessDisplay.selectedIndex].value != frmOriginalConfiguration.txtQuickAccessStartMode.value) {
			return true;
		}
		if (frmConfiguration.cboViewInColour.options[frmConfiguration.cboViewInColour.selectedIndex].value != frmOriginalConfiguration.txtExprColourMode.value) {
			return true;
		}
		if (frmConfiguration.cboExpandNodes.options[frmConfiguration.cboExpandNodes.selectedIndex].value != frmOriginalConfiguration.txtExprNodeMode.value) {
			return true;
		}

		if (frmConfiguration.txtFindSize.value != frmOriginalConfiguration.txtFindSize.value) {
			return true;
		}

		if ((frmConfiguration.chkOwner_Calculations.checked != (frmOriginalConfiguration.txtOnlyMineCalculations.value == 1)) ||
		(frmConfiguration.chkOwner_CrossTabs.checked != (frmOriginalConfiguration.txtOnlyMineCrossTabs.value == 1)) ||
		(frmConfiguration.chkOwner_NineBoxGrid.checked != (frmOriginalConfiguration.txtOnlyMineNineBoxGrid.value == 1)) ||
		(frmConfiguration.chkOwner_CustomReports.checked != (frmOriginalConfiguration.txtOnlyMineCustomReports.value == 1)) ||
		(frmConfiguration.chkOwner_Filters.checked != (frmOriginalConfiguration.txtOnlyMineFilters.value == 1)) ||
		(frmConfiguration.chkOwner_MailMerge.checked != (frmOriginalConfiguration.txtOnlyMineMailMerge.value == 1)) ||
		(frmConfiguration.chkOwner_Picklists.checked != (frmOriginalConfiguration.txtOnlyMinePicklists.value == 1)) ||
		(frmConfiguration.chkOwner_CalendarReports.checked != (frmOriginalConfiguration.txtOnlyMineCalendarReports.value == 1))) {
			return true;
		}

		// If you reach here then nothing has changed.
		return false;
	}

	function restoreDefaults() {

		var answer;

		answer = OpenHR.messageBox("Are you sure you want to restore all default settings?", 36);
		if (answer == 6) {

			setComboValue("PARENT", 3);
			setComboValue("HISTORY", 3);
			setComboValue("LOOKUP", 3);
			setComboValue("QUICKACCESS", 1);

			setComboValue("EXPRCOLOURMODE", 1);
			setComboValue("EXPRNODEMODE", 1);

			frmConfiguration.txtFindSize.value = 1000;

			frmConfiguration.chkOwner_Calculations.checked = false;
			frmConfiguration.chkOwner_CrossTabs.checked = false;
			frmConfiguration.chkOwner_NineBoxGrid.checked = false;
			frmConfiguration.chkOwner_CustomReports.checked = false;
			frmConfiguration.chkOwner_Filters.checked = false;
			frmConfiguration.chkOwner_MailMerge.checked = false;
			frmConfiguration.chkOwner_Picklists.checked = false;
			frmConfiguration.chkOwner_CalendarReports.checked = false;

			enableSaveButton();
		}
	}
</script>

<form action="Configuration_Submit" onsubmit="return false;" method="post" id="frmConfiguration" name="frmConfiguration">
	<div class="pageTitleDiv" style="margin: 10px">
		<span class="pageTitle" id="PopupReportDefinition_PageTitle">Configuration</span>
	</div>

	<div id="div1">
		<table class="outline padleft20">
			<tr>
				<td>
					<table class="invisible" style="width: 500px;">
						<tr>
							<td height="10" colspan="5"></td>
						</tr>
						<tr>
							<td height="10" colspan="5">
								<table class="invisible">
									<tr>
										<td width="10">
											<input type="button" value="Display Defaults" id="btnDummyTab1" name="btnDummyTab1" class="btn btndisabled" disabled="true">
										</td>
										<td width="10"></td>
										<td width="10">
											<input type="button" value="Reports/Utilities & Tools" id="btnTab2" name="btnTab2" class="btn"
												onclick=" display_Configuration_Page(2)" />
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td height="20" colspan="5"></td>
						</tr>
						<tr class="fontsmalltitle">
							<td colspan="5">Record Editing Start Mode</td>
						</tr>
						<tr>
							<td height="10" colspan="5"></td>
						</tr>
						<tr>
							<td width="20"></td>
							<td align="left" nowrap>Parent Tables :
							</td>
							<td width="20"></td>
							<td align="left">
								<select id="cboPrimaryTableDisplay" name="cboPrimaryTableDisplay" class="combo" style="height: 22px; width: 200px;">
									<option value="3" selected>Find Window</option>
									<option value="2">First Record</option>
									<option value="1">New Record</option>
								</select>
							</td>
							<td width="20"></td>
						</tr>
						<tr>
							<td height="5" colspan="5"></td>
						</tr>
						<tr>
							<td width="20"></td>
							<td align="left" nowrap>Child Tables :
							</td>
							<td width="20"></td>
							<td align="left">
								<select id="cboHistoryTableDisplay" name="cboHistoryTableDisplay" class="combo" style="height: 22px; width: 200px">
									<option value="3" selected>Find Window</option>
									<option value="2">First Record</option>
									<option value="1">New Record</option>
								</select>
							</td>
							<td width="20"></td>
						</tr>
						<tr>
							<td height="5" colspan="5"></td>
						</tr>
						<tr>
							<td width="20"></td>
							<td align="left" nowrap>Lookup Tables :
							</td>
							<td width="20"></td>
							<td align="left">
								<select id="cboLookupTableDisplay" name="cboLookupTableDisplay" class="combo" style="height: 22px; width: 200px">
									<option value="3" selected>Find Window</option>
									<option value="2">First Record</option>
									<option value="1">New Record</option>
								</select>
							</td>
							<td width="20"></td>
						</tr>
						<tr>
							<td height="5" colspan="5"></td>
						</tr>
						<tr>
							<td width="20"></td>
							<td align="left" nowrap>Quick Access :
							</td>
							<td width="20"></td>
							<td>
								<select id="cboQuickAccessDisplay" name="cboQuickAccessDisplay" class="combo" style="height: 22px; width: 200px">
									<option value="3" selected>Find Window</option>
									<option value="2">First Record</option>
									<option value="1">New Record</option>
								</select>
							</td>
							<td width="20"></td>
						</tr>
						<tr>
							<td height="20" colspan="5"></td>
						</tr>
						<tr class="fontsmalltitle">
							<td colspan="5">Filters / Calculations</td>
						</tr>
						<tr>
							<td height="10" colspan="5"></td>
						</tr>
						<tr>
							<td width="20"></td>
							<td align="left" nowrap>View in Colour :
							</td>
							<td width="20"></td>
							<td align="left">
								<select id="cboViewInColour" name="cboViewInColour" class="combo" style="height: 22px; width: 200px">
									<option value="1" selected>Monochrome</option>
									<option value="2">Colour Levels</option>
								</select>
							</td>
							<td width="20"></td>
						</tr>
						<tr>
							<td height="5" colspan="5"></td>
						</tr>
						<tr>
							<td width="20"></td>
							<td align="left" nowrap>Expand Nodes :
							</td>
							<td width="20"></td>
							<td align="left">
								<select id="cboExpandNodes" name="cboExpandNodes" class="combo" style="height: 22px; width: 200px">
									<option value="1" selected>Minimized</option>
									<option value="2">Expand All</option>
									<option value="4">Expand Top Level</option>
								</select>
							</td>
							<td width="20"></td>
						</tr>
						<tr>
							<td height="20" colspan="5"></td>
						</tr>
						<tr class="fontsmalltitle">
							<td colspan="5">Find Window / Event Log</td>
						</tr>
						<tr>
							<td height="10" colspan="5"></td>
						</tr>
						<tr>
							<td width="20"></td>
							<td align="left" nowrap>Block Size :
							</td>
							<td width="20"></td>
							<td align="left">
								<input id="txtFindSize" name="txtFindSize" class="text" style="height: 22px; width: 200px" width="200"
									onkeyup="validateFindBlockSize()"
									onchange="validateFindBlockSize()" />
							</td>
							<td width="20"></td>
						</tr>
						<tr>
							<td height="20" colspan="5"></td>
						</tr>

						<tr>
							<td colspan="4">
								<input id="btnDiv1Restore" name="btnDiv1Restore" type="button" value="Restore Defaults" class="btn floatright" onclick="restoreDefaults()" />
								<input id="btnDiv1OK" name="btnDiv1OK" type="button" value="OK" class="btn" />
								<input id="btnDiv1Cancel" name="btnDiv1Cancel" type="button" value="Cancel" class="btn" />
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</div>

	<div id="div2" style="visibility: hidden; display: none">
		<table class="outline padleft20">
			<tr>
				<td>
					<table class="invisible">
						<tr>
							<td style="height: 10px" colspan="7"></td>
						</tr>
						<tr>
							<td height="10" colspan="7">
								<table class="invisible">
									<tr>
										<td style="width: 10px">
											<input type="button" value="Display Defaults" id="btnTab1" name="btnTab1" class="btn" onclick=" display_Configuration_Page(1)" />
										</td>
										<td style="width: 10px"></td>
										<td style="width: 10px">
											<input type="button" value="Reports/Utilities & Tools" id="btnDummyTab2" name="btnDummyTab2" class="btn btndisabled" disabled="true" />
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td style="height: 20px" colspan="7"></td>
						</tr>
						<tr>
							<td></td>
							<td align="center" colspan="5">Only show definitions where owner is '<%=session("username")%>' for the following :
							</td>
							<td width="20"></td>
						</tr>
						<tr>
							<td style="height: 10px" colspan="7"></td>
						</tr>

						<tr>
							<td></td>
							<td class="fontsmalltitle" colspan="6">Reports</td>
						</tr>

						<tr>
							<td></td>
							<td style="width: 20px"></td>
							<td align="left" nowrap>
								<input type="checkbox" id="chkOwner_CustomReports" name="chkOwner_CustomReports"/>
								<label for="chkOwner_CustomReports" class="checkbox">Custom Reports</label>
							</td>
							<td colspan="4"></td>
						</tr>

						<tr>
							<td></td>
							<td style="width: 20px"></td>
							<td align="left" nowrap>
								<input type="checkbox" id="chkOwner_CalendarReports" name="chkOwner_CalendarReports" />
								<label for="chkOwner_CalendarReports" class="checkbox">Calendar Reports</label>
							</td>
							<td colspan="4"></td>
						</tr>

						<tr>
							<td></td>
							<td style="width: 20px"></td>
							<td align="left" nowrap>
								<input type="checkbox" id="chkOwner_MailMerge" name="chkOwner_MailMerge" />
								<label for="chkOwner_MailMerge" class="checkbox">Mail Merge</label>
							</td>
							<td colspan="4"></td>
						</tr>

						<tr>
							<td></td>
							<td style="width: 20px"></td>
							<td align="left" nowrap>
								<input type="checkbox" id="chkOwner_CrossTabs" name="chkOwner_CrossTabs" />
								<label for="chkOwner_CrossTabs" class="checkbox">Cross Tabs</label>
							</td>
							<td colspan="4"></td>
						</tr>

						<tr>
							<td></td>
							<td style="width: 20px"></td>
							<td align="left" nowrap>
								<input type="checkbox" id="chkOwner_NineBoxGrid" name="chkOwner_NineBoxGrid" />
								<label for="chkOwner_NineBoxGrid" class="checkbox">9-Box Grid Reports</label>
							</td>
							<td colspan="4"></td>
						</tr>

						<tr>
							<td colspan="7" style="height: 10px"></td>
						</tr>
						<tr>
							<td></td>
							<td class="fontsmalltitle" colspan="6">Utilities / Tools</td>
						</tr>

						<tr>
							<td></td>
							<td style="width: 20px"></td>
							<td align="left" nowrap>
								<input type="checkbox" id="chkOwner_Calculations" name="chkOwner_Calculations" />
								<label for="chkOwner_Calculations" class="checkbox">Calculations</label>
							</td>
							<td colspan="4"></td>
						</tr>

						<tr>
							<td></td>
							<td style="width: 20px"></td>
							<td align="left" nowrap>
								<input type="checkbox" id="chkOwner_Filters" name="chkOwner_Filters" />
								<label for="chkOwner_Filters" class="checkbox">Filters</label>
							</td>
							<td colspan="4"></td>
						</tr>

						<tr>
							<td></td>
							<td style="width: 20px"></td>
							<td align="left" nowrap>
								<input type="checkbox" id="chkOwner_Picklists" name="chkOwner_Picklists" />
								<label for="chkOwner_Picklists" class="checkbox">Picklists</label>
							</td>
							<td colspan="4"></td>
						</tr>



						<tr>
							<td colspan="6">
								<input id="btnDiv2Restore" name="btnDiv2Restore" type="button" class="btn floatright" value="Restore Defaults" style="width: 150px"  onclick="restoreDefaults()" />
								<input id="btnDiv2OK" name="btnDiv2OK" type="button" class="btn" value="OK" style="width: 75px" />
								<input id="btnDiv2Cancel" name="btnDiv2Cancel" type="button" class="btn" value="Cancel" style="width: 75px" />
							</td>
							<td style="width: 20px"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</div>

	<input type="hidden" id="txtReaction" name="txtReaction">

	<input type="hidden" id="txtPrimaryStartMode" name="txtPrimaryStartMode">
	<input type="hidden" id="txtHistoryStartMode" name="txtHistoryStartMode">
	<input type="hidden" id="txtLookupStartMode" name="txtLookupStartMode">
	<input type="hidden" id="txtQuickAccessStartMode" name="txtQuickAccessStartMode">
	<input type="hidden" id="txtExprColourMode" name="txtExprColourMode">
	<input type="hidden" id="txtExprNodeMode" name="txtExprNodeMode">

	<input type="hidden" id="txtOwner_BatchJobs" name="txtOwner_BatchJobs" value="0">
	<input type="hidden" id="txtOwner_Calculations" name="txtOwner_Calculations" value="0">
	<input type="hidden" id="txtOwner_CrossTabs" name="txtOwner_CrossTabs" value="0">
	<input type="hidden" id="txtOwner_NineBoxGrid" name="txtOwner_NineBoxGrid" value="0">
	<input type="hidden" id="txtOwner_CustomReports" name="txtOwner_CustomReports" value="0">
	<input type="hidden" id="txtOwner_DataTransfer" name="txtOwner_DataTransfer" value="0">
	<input type="hidden" id="txtOwner_Export" name="txtOwner_Export" value="0">
	<input type="hidden" id="txtOwner_Filters" name="txtOwner_Filters" value="0">
	<input type="hidden" id="txtOwner_GlobalAdd" name="txtOwner_GlobalAdd" value="0">
	<input type="hidden" id="txtOwner_GlobalUpdate" name="txtOwner_GlobalUpdate" value="0">
	<input type="hidden" id="txtOwner_GlobalDelete" name="txtOwner_GlobalDelete" value="0">
	<input type="hidden" id="txtOwner_Import" name="txtOwner_Import" value="0">
	<input type="hidden" id="txtOwner_MailMerge" name="txtOwner_MailMerge" value="0">
	<input type="hidden" id="txtOwner_Picklists" name="txtOwner_Picklists" value="0">
	<input type="hidden" id="txtOwner_CalendarReports" name="txtOwner_CalendarReports" value="0">
	<input type="hidden" id="txtOwner_CareerProgression" name="txtOwner_CareerProgression" value="0">
	<input type="hidden" id="txtOwner_EmailGroups" name="txtOwner_EmailGroups" value="0">
	<input type="hidden" id="txtOwner_Labels" name="txtOwner_Labels" value="0">
	<input type="hidden" id="txtOwner_LabelDefinition" name="txtOwner_LabelDefinition" value="0">
	<input type="hidden" id="txtOwner_MatchReports" name="txtOwner_MatchReports" value="0">
	<input type="hidden" id="txtOwner_RecordProfile" name="txtOwner_RecordProfile" value="0">
	<input type="hidden" id="txtOwner_SuccessionPlanning" name="txtOwner_SuccessionPlanning" value="0">

	<input type="hidden" id="txtWarn_DataTransfer" name="txtWarn_DataTransfer" value="0">
	<input type="hidden" id="txtWarn_GlobalAdd" name="txtWarn_GlobalAdd" value="0">
	<input type="hidden" id="txtWarn_GlobalUpdate" name="txtWarn_GlobalUpdate" value="0">
	<input type="hidden" id="txtWarn_GlobalDelete" name="txtWarn_GlobalDelete" value="0">
	<input type="hidden" id="txtWarn_Import" name="txtWarn_Import" value="0">

	<%=Html.AntiForgeryToken()%>
</form>

<form id="frmOriginalConfiguration" name="frmOriginalConfiguration">
	<input type="hidden" id="Hidden1" name="txtPrimaryStartMode" value='<%=session("PrimaryStartMode")%>'>
	<input type="hidden" id="Hidden2" name="txtHistoryStartMode" value='<%=session("HistoryStartMode")%>'>
	<input type="hidden" id="Hidden3" name="txtLookupStartMode" value='<%=session("LookupStartMode")%>'>
	<input type="hidden" id="Hidden4" name="txtQuickAccessStartMode" value='<%=session("QuickAccessStartMode")%>'>
	<input type="hidden" id="Hidden5" name="txtExprColourMode" value='<%=session("ExprColourMode")%>'>
	<input type="hidden" id="Hidden6" name="txtExprNodeMode" value='<%=session("ExprNodeMode")%>'>
	<input type="hidden" id="Hidden7" name="txtFindSize" value='<%=Session("FindRecords")%>'>
	<input type="hidden" id="txtLastFindSize" name="txtLastFindSize" value='<%=session("FindRecords")%>'>

	<input type="hidden" id="txtOnlyMineBatchJobs" name="txtOnlyMineBatchJobs" value='<%=session("onlyMine BatchJobs")%>'>
	<input type="hidden" id="txtOnlyMineCalculations" name="txtOnlyMineCalculations" value='<%=session("onlyMine Calculations")%>'>
	<input type="hidden" id="txtOnlyMineCrossTabs" name="txtOnlyMineCrossTabs" value='<%=session("onlyMine CrossTabs")%>'>
	<input type="hidden" id="txtOnlyMineNineBoxGrid" name="txtOnlyMineNineBoxGrid" value='<%=Session("onlyMine NineBoxGrid")%>'>
	<input type="hidden" id="txtOnlyMineCustomReports" name="txtOnlyMineCustomReports" value='<%=session("onlyMine CustomReports")%>'>
	<input type="hidden" id="txtOnlyMineDataTransfer" name="txtOnlyMineDataTransfer" value='<%=session("onlyMine DataTransfer")%>'>
	<input type="hidden" id="txtOnlyMineExport" name="txtOnlyMineExport" value='<%=session("onlyMine Export")%>'>
	<input type="hidden" id="txtOnlyMineFilters" name="txtOnlyMineFilters" value='<%=session("onlyMine Filters")%>'>
	<input type="hidden" id="txtOnlyMineGlobalAdd" name="txtOnlyMineGlobalAdd" value='<%=session("onlyMine GlobalAdd")%>'>
	<input type="hidden" id="txtOnlyMineGlobalUpdate" name="txtOnlyMineGlobalUpdate" value='<%=session("onlyMine GlobalUpdate")%>'>
	<input type="hidden" id="txtOnlyMineGlobalDelete" name="txtOnlyMineGlobalDelete" value='<%=session("onlyMine GlobalDelete")%>'>
	<input type="hidden" id="txtOnlyMineImport" name="txtOnlyMineImport" value='<%=session("onlyMine Import")%>'>
	<input type="hidden" id="txtOnlyMineMailMerge" name="txtOnlyMineMailMerge" value='<%=session("onlyMine MailMerge")%>'>
	<input type="hidden" id="txtOnlyMinePicklists" name="txtOnlyMinePicklists" value='<%=session("onlyMine Picklists")%>'>
	<input type="hidden" id="txtOnlyMineCalendarReports" name="txtOnlyMineCalendarReports" value='<%=session("onlyMine CalendarReports")%>'>
	<input type="hidden" id="txtOnlyMineCareerProgression" name="txtOnlyMineCareerProgression" value='<%=session("onlyMine CareerProgression")%>'>
	<input type="hidden" id="txtOnlyMineEmailGroups" name="txtOnlyMineEmailGroups" value='<%=session("onlyMine EmailGroups")%>'>
	<input type="hidden" id="txtOnlyMineLabels" name="txtOnlyMineLabels" value='<%=session("onlyMine Labels")%>'>
	<input type="hidden" id="txtOnlyMineLabelDefinition" name="txtOnlyMineLabelDefinition" value='<%=session("onlyMine LabelDefinition")%>'>
	<input type="hidden" id="txtOnlyMineMatchReports" name="txtOnlyMineMatchReports" value='<%=session("onlyMine MatchReports")%>'>
	<input type="hidden" id="txtOnlyMineRecordProfile" name="txtOnlyMineRecordProfile" value='<%=session("onlyMine RecordProfile")%>'>
	<input type="hidden" id="txtOnlyMineSuccessionPlanning" name="txtOnlyMineSuccessionPlanning" value='<%=session("onlyMine SuccessionPlanning")%>'>

	<input type="hidden" id="txtUtilWarnDataTransfer" name="txtUtilWarnDataTransfer" value='<%=session("warning DataTransfer")%>'>
	<input type="hidden" id="txtUtilWarnGlobalAdd" name="txtUtilWarnGlobalAdd" value='<%=session("warning GlobalAdd")%>'>
	<input type="hidden" id="txtUtilWarnGlobalUpdate" name="txtUtilWarnGlobalUpdate" value='<%=session("warning GlobalUpdate")%>'>
	<input type="hidden" id="txtUtilWarnGlobalDelete" name="txtUtilWarnGlobalDelete" value='<%=session("warning GlobalDelete")%>'>
	<input type="hidden" id="txtUtilWarnImport" name="txtUtilWarnImport" value='<%=session("warning Import")%>'>
</form>

<script type="text/javascript">
	configuration_window_onload();

</script>
