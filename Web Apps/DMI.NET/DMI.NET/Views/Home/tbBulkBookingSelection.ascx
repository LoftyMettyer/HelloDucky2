<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@Import namespace="DMI.NET" %>

<%
	Response.Expires = -1
	Dim sErrorDescription = ""
	Dim sFailureDescription = ""
%>
<script src="<%: Url.Content("~/Scripts/ctl_SetFont.js") %>" type="text/javascript"></script>

<script type="text/javascript">

	function tbBulkBookingSelection_onload() {		
		var fOK = true;

		if ((frmUseful.txtSelectionType.value.toUpperCase() != "FILTER") &&
			(frmUseful.txtSelectionType.value.toUpperCase() != "PICKLIST")) {

			var sErrMsg = document.getElementById("txtErrorDescription").value;
			if (sErrMsg.length > 0) {
				fOK = false;
				OpenHR.messageBox(sErrMsg);
				//TODO: //window.parent.close();
			}

			if (fOK == true) {
				if (selectView.length == 0) {
					fOK = false;
					OpenHR.messageBox("You do not have permission to read the employee table.");
					//TODO: //window.parent.close();
				}
			}

			if (fOK == true) {
				if (selectOrder.length == 0) {
					fOK = false;
					OpenHR.messageBox("You do not have permission to use any of the employee table orders.");
					//TODO: //window.parent.close();					
				}
			}

		}

		cmdCancel.focus();

		if ((frmUseful.txtSelectionType.value.toUpperCase() == "FILTER") ||
			(frmUseful.txtSelectionType.value.toUpperCase() == "PICKLIST")) {

			setGridFont(ssOleDBGridSelRecords);

			// Set focus onto one of the form controls. 
			// NB. This needs to be done before making any reference to the grid
			ssOleDBGridSelRecords.focus();

			if (ssOleDBGridSelRecords.rows > 0) {
				// Select the top row.
				ssOleDBGridSelRecords.MoveFirst();
				ssOleDBGridSelRecords.SelBookmarks.Add(ssOleDBGridSelRecords.Bookmark);
			}

			tbrefreshControls();
		} else {
			setGridFont(ssOleDBGridSelRecords);

			setMenuFont(abMainMenu);

			abMainMenu.Attach();
			abMainMenu.DataPath = "misc\\mainmenu.htm";
			abMainMenu.RecalcLayout();

			window.parent.dialogLeft = new String((screen.width - (9 * screen.width / 10)) / 2) + "px";
			window.parent.dialogTop = new String((screen.height - (3 * screen.height / 4)) / 2) + "px";
			window.parent.dialogWidth = new String((9 * screen.width / 10)) + "px";
			window.parent.dialogHeight = new String((3 * screen.height / 4)) + "px";

			window.parent.txtTableID.value = frmUseful.txtTableID.value;
			window.parent.txtViewID.value = selectView.options[selectView.selectedIndex].value;
			window.parent.txtOrderID.value = selectOrder.options[selectOrder.selectedIndex].value;
			window.parent.loadAddRecords();
		}
	}
</script>

<script type="text/javascript">

	function tbrefreshControls()
	{
		var fNoneSelected;

		fNoneSelected = (ssOleDBGridSelRecords.SelBookmarks.Count == 0);

		button_disable(cmdOK, fNoneSelected);

		var frmUseful = document.getElementById("frmUseful");

		if ((frmUseful.txtSelectionType.value.toUpperCase() != "FILTER") &&
			(frmUseful.txtSelectionType.value.toUpperCase() != "PICKLIST")) {

			if (selectOrder.length <= 1) {
				combo_disable(selectOrder, true);
				button_disable(btnGoOrder, true);
			}

			if (selectView.length <= 1) {
				combo_disable(selectView, true);
				button_disable(btnGoView, true);
			}
		}
	}

	function makeSelection() {
		var frmUseful = document.getElementById("frmUseful");
		var frmPrompt = document.getElementById("frmPrompt");
		
		if (frmUseful.txtSelectionType.value.toUpperCase() == "FILTER") {
			// Go to the prompted values form to get any required prompts. 
			frmPrompt.filterID.value = selectedRecordID();
			frmPrompt.submit();		
		}
		else {
			if (frmUseful.txtSelectionType.value.toUpperCase() == "PICKLIST") {
				try {
					//TODO: window.parent.window.dialogArguments.window.makeSelection(frmUseful.txtSelectionType.value, selectedRecordID(), "");
					alert("not done");
				}
				catch(e) {
				}
			}
			else {
				var sSelectedIDs = "";
			
				ssOleDBGridSelRecords.redraw = false;
				for (var iIndex = 0; iIndex < ssOleDBGridSelRecords.selbookmarks.Count(); iIndex++) {	
					ssOleDBGridSelRecords.bookmark = ssOleDBGridSelRecords.selbookmarks(iIndex);

					sRecordID = ssOleDBGridSelRecords.Columns("ID").Value;

					if (sSelectedIDs.length > 0) {
						sSelectedIDs = sSelectedIDs + ",";
					}
					sSelectedIDs = sSelectedIDs + sRecordID;				
				}
				ssOleDBGridSelRecords.redraw = true;

				try {
					//TODO: window.parent.window.dialogArguments.window.makeSelection(frmUseful.txtSelectionType.value, 0, sSelectedIDs);
					makeSelection(frmUseful.txtSelectionType.value, 0, sSelectedIDs);					
				}
				catch(e) {
				}
			}
			//window.parent.close();
		}
	}

	/* Return the ID of the record selected in the find form. */
	function selectedRecordID() {
		var iRecordID;

		iRecordID = 0;
	
		if (ssOleDBGridSelRecords.SelBookmarks.Count > 0) {   
			iRecordID = ssOleDBGridSelRecords.Columns("ID").Value;
		}

		return(iRecordID);
	}

	function locateRecord(psSearchFor) {
		var fFound;

		fFound = false;
	
		ssOleDBGridSelRecords.redraw = false;

		ssOleDBGridSelRecords.MoveLast();
		ssOleDBGridSelRecords.MoveFirst();

		ssOleDBGridSelRecords.SelBookmarks.removeall();
	
		for (iIndex = 1; iIndex <= ssOleDBGridSelRecords.rows; iIndex++) 
		{	
			var sGridValue = new String(ssOleDBGridSelRecords.Columns(0).value);
			sGridValue = sGridValue.substr(0, psSearchFor.length).toUpperCase();
			if (sGridValue == psSearchFor.toUpperCase()) 
			{
				ssOleDBGridSelRecords.SelBookmarks.Add(ssOleDBGridSelRecords.Bookmark);
				fFound = true;
				break;
			}

			if (iIndex < ssOleDBGridSelRecords.rows) 
			{
				ssOleDBGridSelRecords.MoveNext();
			}
			else 
			{
				break;
			}
		}

		if ((fFound == false) && (ssOleDBGridSelRecords.rows > 0)) {
			// Select the top row.
			ssOleDBGridSelRecords.MoveFirst();
			ssOleDBGridSelRecords.SelBookmarks.Add(ssOleDBGridSelRecords.Bookmark);
		}

		ssOleDBGridSelRecords.redraw = true;
	}

	function goView() {
		// Get the tbBulkBookingSelectionData.asp to get the find records.
		var dataForm = OpenHR.getForm("dataframe", "frmGetData");
		dataForm.txtTableID.value = frmUseful.txtTableID.value;
		dataForm.txtViewID.value = selectView.options[selectView.selectedIndex].value;
		dataForm.txtOrderID.value = selectOrder.options[selectOrder.selectedIndex].value;
		dataForm.txtFirstRecPos.value = 1;
		dataForm.txtCurrentRecCount.value = 0;
		dataForm.txtPageAction.value = "LOAD";

		refreshData();
	}

	function goOrder() {
		// Get the tbBulkBookingSelectionData.asp to get the find records.
		var dataForm = OpenHR.getForm("dataframe", "frmGetData");
		dataForm.txtTableID.value = frmUseful.txtTableID.value;
		dataForm.txtViewID.value = selectView.options[selectView.selectedIndex].value;
		dataForm.txtOrderID.value = selectOrder.options[selectOrder.selectedIndex].value;
		dataForm.txtFirstRecPos.value = 1;
		dataForm.txtCurrentRecCount.value = 0;
		dataForm.txtPageAction.value = "LOAD";

		refreshData();
	}

	function selectedOrderID() {
		return selectOrder.options[selectOrder.selectedIndex].value;
	}

	function selectedViewID() {
		return selectView.options[selectView.selectedIndex].value;
	}

	function tbrefreshMenu() 
	{
		if (abMainMenu.Bands.Count() > 0) 
		{
			window.setTimeout("enableMenu()", 250);

			var frmData = window.parent.frames("dataframe").document.forms("frmData");
				
			for (i=0; i< abMainMenu.tools.count(); i++) 
			{
				abMainMenu.tools(i).visible = false;
			}				
		
			abMainMenu.Bands("mnuMainMenu").visible = false;
			abMainMenu.Bands("mnubandMainToolBar").visible = true;

			// Enable the record editing options as necessary.
			abMainMenu.tools("mnutoolFirstRecord").visible = true;
			abMainMenu.tools("mnutoolFirstRecord").enabled = (frmData.txtIsFirstPage.value != "True");
			abMainMenu.tools("mnutoolPreviousRecord").visible = true;
			abMainMenu.tools("mnutoolPreviousRecord").enabled = (frmData.txtIsFirstPage.value != "True");
			abMainMenu.tools("mnutoolNextRecord").visible = true;
			abMainMenu.tools("mnutoolNextRecord").enabled = (frmData.txtIsLastPage.value != "True");
			abMainMenu.tools("mnutoolLastRecord").visible = true;
			abMainMenu.tools("mnutoolLastRecord").enabled = (frmData.txtIsLastPage.value != "True");

			abMainMenu.tools("mnutoolLocateRecordsCaption").visible = true;
			abMainMenu.tools("mnutoolLocateRecords").visible = (frmData.txtFirstColumnType.value != "-7");
			abMainMenu.Tools("mnutoolLocateRecordsLogic").CBList.Clear();
			abMainMenu.Tools("mnutoolLocateRecordsLogic").CBList.AddItem("True");
			abMainMenu.Tools("mnutoolLocateRecordsLogic").CBList.AddItem("False");
			abMainMenu.tools("mnutoolLocateRecordsLogic").visible = (frmData.txtFirstColumnType.value == "-7");

			sCaption = "";
			if (frmData.txtRecordCount.value > 0) 
			{
				iStartPosition = new Number(frmData.txtFirstRecPos.value);
				iEndPosition = new Number(frmData.txtRecordCount.value);
				iEndPosition = iStartPosition - 1 + iEndPosition;
				sCaption = "Records " +
					iStartPosition + 
					" to " +
					iEndPosition +
					" of " +
					frmData.txtTotalRecordCount.value;
			}
			else 
			{
				sCaption = "No Records";
			}

			abMainMenu.tools("mnutoolRecordPosition").visible = true;
			abMainMenu.Bands("mnubandMainToolBar").tools("mnutoolRecordPosition").caption = sCaption;
			
			try
			{
				window.resizeBy(1,1);	
				window.resizeBy(-1,-1);	
				window.resizeBy(1,1);	
				window.resizeBy(-1,-1);	
			}
			catch(e) {}

			try
			{
				abMainMenu.Attach();
				abMainMenu.RecalcLayout();
				abMainMenu.ResetHooks();
				abMainMenu.Refresh();
			}
			catch(e) {}			
				
			// Adjust the framset dimensions to suit the size of the menu.
			/*lngMenuHeight = abMainMenu.Bands("mnubandMainToolBar").height;
			sTemp = new String(lngMenuHeight);
			if(frmUseful.txtIEVersion.value >= 5.5) 
			{
				window.parent.document.all.item("mainframeset").rows = "*, " + sTemp;
			}
			else 
			{
				window.parent.document.all.item("mainframeset").rows = "*, 0";
			}*/
		}
	}

	function reloadPage(psAction, psLocateValue) {
		var sConvertedValue;
		var sDecimalSeparator;
		var sThousandSeparator;
		var sPoint;
		var iIndex ;
		var iTempSize;
		var iTempDecimals;
	
		sDecimalSeparator = "\\";
		sDecimalSeparator = sDecimalSeparator.concat(ASRIntranetFunctions.LocaleDecimalSeparator);
		var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");

		sThousandSeparator = "\\";
		sThousandSeparator = sThousandSeparator.concat(ASRIntranetFunctions.LocaleThousandSeparator);
		var reThousandSeparator = new RegExp(sThousandSeparator, "gi");

		sPoint = "\\.";
		var rePoint = new RegExp(sPoint, "gi");

		fValidLocateValue = true;

		var dataForm = window.parent.frames("dataframe").document.forms("frmData");
		var getDataForm = window.parent.frames("dataframe").document.forms("frmGetData");

		if (psAction == "LOCATE") {
			// Check that the entered value is valid for the first order column type.
			iDataType = dataForm.txtFirstColumnType.value;

			if ((iDataType == 2) || (iDataType == 4)) {
				// Numeric/Integer column.
				// Ensure that the value entered is numeric.
				if (psLocateValue.length == 0) {
					psLocateValue = "0";
				}

				// Convert the value from locale to UK settings for use with the isNaN funtion.
				sConvertedValue = new String(psLocateValue);
				// Remove any thousand separators.
				sConvertedValue = sConvertedValue.replace(reThousandSeparator, "");
				psLocateValue = sConvertedValue;

				// Convert any decimal separators to '.'.
				if (ASRIntranetFunctions.LocaleDecimalSeparator != ".") {
					// Remove decimal points.
					sConvertedValue = sConvertedValue.replace(rePoint, "A");
					// replace the locale decimal marker with the decimal point.
					sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
				}

				if (isNaN(sConvertedValue) == true) {
					fValidLocateValue = false;
					ASRIntranetFunctions.MessageBox("Invalid numeric value entered.");
				}
				else {
					psLocateValue = sConvertedValue;
					iIndex = sConvertedValue.indexOf(".");
					if (iDataType == 4) {
						// Ensure that integer columns are compared with integer values.
						if (iIndex >= 0 ) {
							fValidLocateValue = false;
							ASRIntranetFunctions.MessageBox("Invalid integer value entered.");
						}
					} 
					else {
						// Ensure numeric columns are compared with numeric values that do not exceed
						// their defined size and decimals settings.
						if (iIndex >= 0) {
							iTempSize = iIndex;
							iTempDecimals = sConvertedValue.length - iIndex - 1;
						}
						else {
							iTempSize = sConvertedValue.length;
							iTempDecimals = 0;
						}
					
						if ((sConvertedValue.substr(0,1) == "+") ||
							(sConvertedValue.substr(0,1) == "-")) {
							iTempSize = iTempSize - 1;
						}

						if(iTempSize > (dataForm.txtFirstColumnSize.value - dataForm.txtFirstColumnDecimals.value)) {
							fValidLocateValue = false;
							ASRIntranetFunctions.MessageBox("The value cannot have more than " + (dataForm.txtFirstColumnSize.value - dataForm.txtFirstColumnDecimals.value) + " digit(s) to the left of the decimal separator.");
						}
						else {
							if(iTempDecimals > dataForm.txtFirstColumnDecimals.value) {
								fValidLocateValue = false;
								ASRIntranetFunctions.MessageBox("The value cannot have more than " + dataForm.txtFirstColumnDecimals.value + " decimal place(s).");
							}
						}
					}
				}
			}
			else {
				if (iDataType == 11) {
					// Date column.
					// Ensure that the value entered is a date.
					if (psLocateValue.length > 0) {
						// Convert the date to SQL format (use this as a validation check).
						// An empty string is returned if the date is invalid.
						psLocateValue = convertLocaleDateToSQL(psLocateValue)
						if (psLocateValue.length = 0) {
							fValidLocateValue = false;
							ASRIntranetFunctions.MessageBox("Invalid date value entered.");
						}
					}
				}
			}
		}
	
		if (fValidLocateValue == true) {
			disableMenu();

			// Get the optionData.asp to get the link find records.
			getDataForm.txtTableID.value = frmUseful.txtTableID.value;
			getDataForm.txtViewID.value = selectView.options[selectView.selectedIndex].value;
			getDataForm.txtOrderID.value = selectOrder.options[selectOrder.selectedIndex].value;
			getDataForm.txtFirstRecPos.value = dataForm.txtFirstRecPos.value;
			getDataForm.txtCurrentRecCount.value = dataForm.txtRecordCount.value;
			getDataForm.txtGotoLocateValue.value = psLocateValue;
			getDataForm.txtPageAction.value = psAction;

			window.parent.frames("dataframe").refreshData();
		}

		// Clear the locate value from the menu.
		abMainMenu.Tools("mnutoolLocateRecords").Text = "";
	}

	function disableMenu() {
		for(iLoop = 0; iLoop < abMainMenu.Bands.Item("mnubandMainToolBar").Tools.Count(); iLoop ++) {
			abMainMenu.Bands.Item("mnubandMainToolBar").tools.Item(iLoop).Enabled = false;
		}

		abMainMenu.RecalcLayout();
		abMainMenu.ResetHooks();
		abMainMenu.Refresh();
	}

	function enableMenu() {
		for(iLoop = 0; iLoop < abMainMenu.Bands.Item("mnubandMainToolBar").Tools.Count(); iLoop ++) {
			abMainMenu.Bands.Item("mnubandMainToolBar").tools.Item(iLoop).Enabled = true;
		}
	}

	function convertLocaleDateToSQL(psDateString)
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
		
		sDateFormat = ASRIntranetFunctions.LocaleDateFormat;

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
	
		sValue = ASRIntranetFunctions.ConvertSQLDateToLocale(sTempValue);

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

<SCRIPT FOR=abMainMenu EVENT=DataReady LANGUAGE=JavaScript>

	var sKey;
	sKey = new String("tempmenufilepath_");
	sKey = sKey.concat(window.parent.window.dialogArguments.window.parent.frames("menuframe").document.forms("frmMenuInfo").txtDatabase.value);	
	sPath = ASRIntranetFunctions.GetRegistrySetting("HR Pro", "DataPaths", sKey);
	if(sPath == "") {
		sPath = "c:\\";
	}

	if(sPath == "<NONE>") {
		frmUseful.txtMenuSaved.value = 1;
		abMainMenu.RecalcLayout();
	}
	else {
		if (sPath.substr(sPath.length - 1, 1) != "\\") {
			sPath = sPath.concat("\\");
		}
		
		sPath = sPath.concat("tempmenu.asp");
		if ((abMainMenu.Bands.Count() > 0) && (frmUseful.txtMenuSaved.value == 0)) {
			try {
				abMainMenu.save(sPath, "");
			}
			catch(e) {
				ASRIntranetFunctions.MessageBox("The specified temporary menu file path cannot be written to. The temporary menu file path will be cleared."); 
				sKey = new String("tempMenuFilePath_");
				sKey = sKey.concat(window.parent.window.dialogArguments.window.parent.frames("menuframe").document.forms("frmMenuInfo").txtDatabase.value);	
				ASRIntranetFunctions.SaveRegistrySetting("HR Pro", "DataPaths", sKey, "<NONE>");
			}

			frmUseful.txtMenuSaved.value = 1;
		}
		else {
			if ((abMainMenu.Bands.Count() == 0) && (frmUseful.txtMenuSaved.value == 1)) {
				abMainMenu.DataPath = sPath;
				abMainMenu.RecalcLayout();
				return;
			}
		}
	}
</script>

<SCRIPT FOR=abMainMenu EVENT=PreCustomizeMenu(pfCancel) LANGUAGE=JavaScript>
	pfCancel = true;
	ASRIntranetFunctions.MessageBox("The menu cannot be customized. Errors will occur if you attempt to customize it. Click anywhere in your browser to remove the dummy customisation menu.");
</script>

<SCRIPT FOR=abMainMenu EVENT=Click(pTool) LANGUAGE=JavaScript>
	switch (pTool.name) {
	case "mnutoolFirstRecord" :
		reloadPage("MOVEFIRST", "");
		break;
	case "mnutoolPreviousRecord" :
		reloadPage("MOVEPREVIOUS", "");
		break;
	case "mnutoolNextRecord" :
		reloadPage("MOVENEXT", "");
		break;
	case "mnutoolLastRecord" :
		reloadPage("MOVELAST", "");
		break;
	}
</script>

<SCRIPT FOR=abMainMenu EVENT="KeyDown(piKeyCode, piShift)" LANGUAGE=JavaScript>
	iIndex = abMainMenu.ActiveBand.CurrentTool;
	
	if (abMainMenu.ActiveBand.Tools(iIndex).Name == "mnutoolLocateRecords") {
		if (piKeyCode == 13) {
			sLocateValue = abMainMenu.ActiveBand.Tools(iIndex).Text;

			reloadPage("LOCATE", sLocateValue);
		}
	}
</script>

<SCRIPT FOR=abMainMenu EVENT=ComboSelChange(pTool) LANGUAGE=JavaScript>
	if (pTool.Name == "mnutoolLocateRecordsLogic") {
		sLocateValue = pTool.Text;

		reloadPage("LOCATE", sLocateValue);
	}
</script>

<SCRIPT FOR=abMainMenu EVENT=PreSysMenu(pBand) LANGUAGE=JavaScript>
	if(pBand.Name == "SysCustomize") {
		pBand.Tools.RemoveAll();
	}
</script>

<SCRIPT FOR=ssOleDBGridSelRecords EVENT=rowcolchange LANGUAGE=JavaScript>
	tbrefreshControls();
</script>

<SCRIPT FOR=ssOleDBGridSelRecords EVENT=dblClick LANGUAGE=JavaScript>
	// JPD20021031 Fault 4631
	makeSelection();
</script>

<SCRIPT FOR=ssOleDBGridSelRecords EVENT=KeyPress(iKeyAscii) LANGUAGE=JavaScript>
	if ((iKeyAscii >= 32) && (iKeyAscii <= 255)) {	
		var dtTicker = new Date();
		var iThisTick = new Number(dtTicker.getTime());
		if (txtLastKeyFind.value.length > 0) {
			var iLastTick = new Number(txtTicker.value);
		}
		else {
			var iLastTick = new Number("0");
		}
		
		if (iThisTick > (iLastTick + 1500)) {
			var sFind = String.fromCharCode(iKeyAscii);
		}
		else {
			var sFind = txtLastKeyFind.value + String.fromCharCode(iKeyAscii);
		}
		
		txtTicker.value = iThisTick;
		txtLastKeyFind.value = sFind;

		locateRecord(sFind);
	}
</script>

<script src="<%: Url.Content("~/Scripts/ctl_SetStyles.js") %>" type="text/javascript"></script>

<div bgcolor='<%=session("ConvertedDesktopColour")%>' leftmargin=20 topmargin=20 bottommargin=20 rightmargin=5>

<%
	if (ucase(session("selectionType")) <> ucase("picklist")) and _
		(ucase(session("selectionType")) <> ucase("filter")) then 
		Response.Write("<OBJECT classid=""clsid:6976CB54-C39B-4181-B1DC-1A829068E2E7"" codebase=""cabs/COAInt_Client.cab#Version=1,0,0,5""" & vbCrLf)
		Response.Write("	height=32 id=abMainMenu name=abMainMenu style=""LEFT: 0px; TOP: 0px"" width=100% VIEWASTEXT>" & vbCrLf)
		Response.Write("	<PARAM NAME=""_ExtentX"" VALUE=""847"">" & vbCrLf)
		Response.Write("	<PARAM NAME=""_ExtentY"" VALUE=""847"">" & vbCrLf)
		Response.Write("</OBJECT>" & vbCrLf)
		Response.Write("<table align=center class=""outline"" cellPadding=5 cellSpacing=0 width=100% height=""95%"">" & vbCrLf)
	else
		Response.Write("<table align=center class=""outline"" cellPadding=5 cellSpacing=0 width=100% height=""100%"">" & vbCrLf)
	end if
%>


	<tr>
		<td>
			<table align="center" class="invisible" cellspacing="0" cellpadding="0" width="100%" height="100%">
				<tr height=10>
					<td colspan="3" align="center" height="10">
						<H3 align="center">
<% 
	if ucase(session("selectionType")) = ucase("picklist") then 
		Response.Write("Select Picklist")
	else
		if ucase(session("selectionType")) = ucase("filter") then 
			Response.Write("Select Filter")
		else
			Response.Write("Select Records")
		end if
	End If
	
%>
						</H3>
					</td>
				</tr>
				<tr>
					<td width=20></td>
					<td>
<%
	session("optionLinkViewID") = session("TB_BulkBookingDefaultViewID")
	session("optionLinkOrderID") = 0
	
	sErrorDescription = ""

	if (ucase(session("selectionType")) = ucase("picklist")) or _
		(ucase(session("selectionType")) = ucase("filter")) then 

		Dim cmdSelRecords = CreateObject("ADODB.Command")
		cmdSelRecords.CommandType = 4
		cmdSelRecords.ActiveConnection = Session("databaseConnection")

		if ucase(session("selectionType")) = ucase("picklist") then
			cmdSelRecords.CommandText = "spASRIntGetAvailablePicklists"

			Dim prmTableID = cmdSelRecords.CreateParameter("tableID", 3, 1)	' 3 = integer, 1 = input
			cmdSelRecords.Parameters.Append(prmTableID)
			prmTableID.value = cleanNumeric(clng(session("TB_EmpTableID")))
			
			Dim prmUser = cmdSelRecords.CreateParameter("user", 200, 1, 255)
			cmdSelRecords.Parameters.Append(prmUser)
			prmUser.value = session("username")
		else
			cmdSelRecords.CommandText = "spASRIntGetAvailableFilters"

			Dim prmTableID = cmdSelRecords.CreateParameter("tableID", 3, 1)	' 3 = integer, 1 = input
			cmdSelRecords.Parameters.Append(prmTableID)
			prmTableID.value = cleanNumeric(clng(session("TB_EmpTableID")))
			
			Dim prmUser = cmdSelRecords.CreateParameter("user", 200, 1, 255)
			cmdSelRecords.Parameters.Append(prmUser)
			prmUser.value = session("username")
		end if

		Err.Clear()
		Dim rstSelRecords = cmdSelRecords.Execute

		' Instantiate and initialise the grid. 
		Response.Write("					<OBJECT classid=""clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"" id=ssOleDBGridSelRecords name=ssOleDBGridSelRecords codebase=""cabs/COAInt_Grid.cab#version=3,1,3,6"" style=""LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:400px;"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""ScrollBars"" VALUE=""4"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""_Version"" VALUE=""196616"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""DataMode"" VALUE=""2"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""Cols"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""Rows"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""BorderStyle"" VALUE=""1"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""RecordSelectors"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""GroupHeaders"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""ColumnHeaders"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""GroupHeadLines"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""HeadLines"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""FieldDelimiter"" VALUE=""(None)"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""FieldSeparator"" VALUE=""(Tab)"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""Col.Count"" VALUE=""" & rstSelRecords.fields.count & """>" & vbCrLf)
		Response.Write("						<PARAM NAME=""stylesets.count"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""TagVariant"" VALUE=""EMPTY"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""UseGroups"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""HeadFont3D"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""Font3D"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""DividerType"" VALUE=""3"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""DividerStyle"" VALUE=""1"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""DefColWidth"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""BeveColorScheme"" VALUE=""2"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""BevelColorFrame"" VALUE=""-2147483642"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""BevelColorHighlight"" VALUE=""-2147483628"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""BevelColorShadow"" VALUE=""-2147483632"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""BevelColorFace"" VALUE=""-2147483633"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""CheckBox3D"" VALUE=""-1"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""AllowAddNew"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""AllowDelete"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""AllowUpdate"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""MultiLine"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""ActiveCellStyleSet"" VALUE="""">" & vbCrLf)
		Response.Write("						<PARAM NAME=""RowSelectionStyle"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""AllowRowSizing"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""AllowGroupSizing"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""AllowColumnSizing"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""AllowGroupMoving"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""AllowColumnMoving"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""AllowGroupSwapping"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""AllowColumnSwapping"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""AllowGroupShrinking"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""AllowColumnShrinking"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""AllowDragDrop"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""UseExactRowCount"" VALUE=""-1"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""SelectTypeCol"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""SelectTypeRow"" VALUE=""1"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""SelectByCell"" VALUE=""-1"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""BalloonHelp"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""RowNavigation"" VALUE=""1"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""CellNavigation"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""MaxSelectedRows"" VALUE=""1"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""HeadStyleSet"" VALUE="""">" & vbCrLf)
		Response.Write("						<PARAM NAME=""StyleSet"" VALUE="""">" & vbCrLf)
		Response.Write("						<PARAM NAME=""ForeColorEven"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""ForeColorOdd"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""BackColorEven"" VALUE=""16777215"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""BackColorOdd"" VALUE=""16777215"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""Levels"" VALUE=""1"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""RowHeight"" VALUE=""503"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""ExtraHeight"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""ActiveRowStyleSet"" VALUE="""">" & vbCrLf)
		Response.Write("						<PARAM NAME=""CaptionAlignment"" VALUE=""2"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""SplitterPos"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""SplitterVisible"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""Columns.Count"" VALUE=""" & rstSelRecords.fields.count & """>" & vbCrLf)

		for iLoop = 0 to (rstSelRecords.fields.count - 1)
			if rstSelRecords.fields(iLoop).name <> "name" then
				Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").Width"" VALUE=""0"">" & vbCrLf)
				Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").Visible"" VALUE=""0"">" & vbCrLf)
			else
				Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").Width"" VALUE=""100000"">" & vbCrLf)
				Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").Visible"" VALUE=""-1"">" & vbCrLf)
			End If
								
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").Columns.Count"" VALUE=""1"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").Caption"" VALUE=""" & Replace(rstSelRecords.fields(iLoop).name, "_", " ") & """>" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").Name"" VALUE=""" & rstSelRecords.fields(iLoop).name & """>" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").Alignment"" VALUE=""0"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").CaptionAlignment"" VALUE=""3"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").Bound"" VALUE=""0"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").AllowSizing"" VALUE=""1"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").DataField"" VALUE=""Column " & iLoop & """>" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").DataType"" VALUE=""8"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").Level"" VALUE=""0"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").NumberFormat"" VALUE="""">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").Case"" VALUE=""0"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").FieldLen"" VALUE=""4096"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").VertScrollBar"" VALUE=""0"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").Locked"" VALUE=""0"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").Style"" VALUE=""0"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").ButtonsAlways"" VALUE=""0"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").RowCount"" VALUE=""0"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").ColCount"" VALUE=""1"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").HasHeadForeColor"" VALUE=""0"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").HasHeadBackColor"" VALUE=""0"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").HasForeColor"" VALUE=""0"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").HasBackColor"" VALUE=""0"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").HeadForeColor"" VALUE=""0"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").HeadBackColor"" VALUE=""0"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").ForeColor"" VALUE=""0"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").BackColor"" VALUE=""0"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").HeadStyleSet"" VALUE="""">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").StyleSet"" VALUE="""">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").Nullable"" VALUE=""1"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").Mask"" VALUE="""">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").PromptInclude"" VALUE=""0"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").ClipMode"" VALUE=""0"">" & vbCrLf)
			Response.Write("						<PARAM NAME=""Columns(" & iLoop & ").PromptChar"" VALUE=""95"">" & vbCrLf)
		next 

		Response.Write("						<PARAM NAME=""UseDefaults"" VALUE=""-1"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""TabNavigation"" VALUE=""1"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""_ExtentX"" VALUE=""17330"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""_ExtentY"" VALUE=""1323"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""_StockProps"" VALUE=""79"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""Caption"" VALUE="""">" & vbCrLf)
		Response.Write("						<PARAM NAME=""ForeColor"" VALUE=""0"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""BackColor"" VALUE=""16777215"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""Enabled"" VALUE=""-1"">" & vbCrLf)
		Response.Write("						<PARAM NAME=""DataMember"" VALUE="""">" & vbCrLf)
								
		Dim lngRowCount = 0
		do while not rstSelRecords.EOF
			for iLoop = 0 to (rstSelRecords.fields.count - 1)							
				Response.Write("						<PARAM NAME=""Row(" & lngRowCount & ").Col(" & iLoop & ")"" VALUE=""" & Replace(Replace(rstSelRecords.Fields(iLoop).Value, "_", " "), """", "&quot;") & """>" & vbCrLf)
			next 				
			lngRowCount = lngRowCount + 1
			rstSelRecords.MoveNext
		loop
		Response.Write("						<PARAM NAME=""Row.Count"" VALUE=""" & lngRowCount & """>" & vbCrLf)
		Response.Write("					</OBJECT>" & vbCrLf)
	
		rstSelRecords.close
		rstSelRecords = Nothing

		' Release the ADO command object.
		cmdSelRecords = Nothing
		
	else 
		' Select individual employee records.
%>
						<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
							<tr height="10">
								<td height="10">
									<TABLE WIDTH="100%" height="10" class="invisible" CELLSPACING="0" CELLPADDING="0">
										<TR height="10">
											<TD width="40" height="10">
												View :
											</TD>
											<TD width="10" height="10">
												&nbsp;
											</TD>
											<TD width="175" >
												<SELECT id="selectView" name="selectView" class="combo" style="HEIGHT: 10px; WIDTH: 200px">
<%
	If Len(sErrorDescription) = 0 Then
		' Get the view records.
		Dim cmdViewRecords = CreateObject("ADODB.Command")
		cmdViewRecords.CommandText = "sp_ASRIntGetLinkViews"
		cmdViewRecords.CommandType = 4 ' Stored Procedure
		cmdViewRecords.ActiveConnection = Session("databaseConnection")

		Dim prmTableID = cmdViewRecords.CreateParameter("tableID", 3, 1)
		cmdViewRecords.Parameters.Append(prmTableID)
		prmTableID.value = CleanNumeric(Session("TB_EmpTableID"))

		Dim prmDfltOrderID = cmdViewRecords.CreateParameter("dfltOrderID", 3, 2) ' 11=integer, 2=output
		cmdViewRecords.Parameters.Append(prmDfltOrderID)

		Err.Clear()
		Dim rstViewRecords = cmdViewRecords.Execute

		If (Err.Number <> 0) Then
			sErrorDescription = "The Employee views could not be retrieved." & vbCrLf & FormatError(Err.Description)
		End If

		If Len(sErrorDescription) = 0 Then
			Do While Not rstViewRecords.EOF
				Response.Write("													<OPTION value=" & rstViewRecords.Fields(0).Value)
				If rstViewRecords.Fields(0).Value = Session("optionLinkViewID") Then
					Response.Write(" SELECTED")
				End If

				If rstViewRecords.Fields(0).Value = 0 Then
					Response.Write(">" & Replace(rstViewRecords.Fields(1).Value, "_", " ") & "</OPTION>" & vbCrLf)
				Else
					Response.Write(">'" & Replace(rstViewRecords.Fields(1).Value, "_", " ") & "' view</OPTION>" & vbCrLf)
				End If

				rstViewRecords.MoveNext()
			Loop
			
			If (rstViewRecords.EOF And rstViewRecords.BOF) Then
				sFailureDescription = "You do not have permission to read the Employee table."
			End If
		
			' Release the ADO recordset object.
			rstViewRecords.close()
			rstViewRecords = Nothing
	
			' NB. IMPORTANT ADO NOTE.
			' When calling a stored procedure which returns a recordset AND has output parameters
			' you need to close the recordset and set it to nothing before using the output parameters. 
			If Session("optionLinkOrderID") <= 0 Then
				Session("optionLinkOrderID") = cmdViewRecords.Parameters("dfltOrderID").Value
			End If
		End If

		' Release the ADO command object.
		cmdViewRecords = Nothing
	End If
%>
												</SELECT>
											</TD>
											<TD width="10" >
												<INPUT type="button" value="Go" id="btnGoView" name="btnGoView" class="btn"
												    onclick="goView()"
                                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                                    onblur="try{button_onBlur(this);}catch(e){}" />
											</TD>
											<TD height=10>
												&nbsp;
											</TD>
											<TD width=40 height=10>
												Order :
											</TD>
											<TD width=10 height=10>
												&nbsp;
											</TD>
											<TD width=175 >
												<SELECT id=selectOrder name=selectOrder class="combo" style="HEIGHT: 10px; WIDTH: 200px">
<%
		if len(sErrorDescription) = 0 then
			' Get the order records.
		Dim cmdOrderRecords = CreateObject("ADODB.Command")
			cmdOrderRecords.CommandText = "sp_ASRIntGetTableOrders"
			cmdOrderRecords.CommandType = 4 ' Stored Procedure
		cmdOrderRecords.ActiveConnection = Session("databaseConnection")

		Dim prmTableID = cmdOrderRecords.CreateParameter("tableID", 3, 1)
		cmdOrderRecords.Parameters.Append(prmTableID)
			prmTableID.value = cleanNumeric(session("TB_EmpTableID"))

		Dim prmViewID = cmdOrderRecords.CreateParameter("viewID", 3, 1)
		cmdOrderRecords.Parameters.Append(prmViewID)
			prmViewID.value = 0

		Err.Clear()
		Dim rstOrderRecords = cmdOrderRecords.Execute

		If (Err.Number <> 0) Then
			sErrorDescription = "The order records could not be retrieved." & vbCrLf & FormatError(Err.Description)
		End If

			if len(sErrorDescription) = 0 then
				do while not rstOrderRecords.EOF
				Response.Write("													<OPTION value=" & rstOrderRecords.Fields(1).Value)

					if rstOrderRecords.Fields(1).Value = session("optionLinkOrderID") then
					Response.Write(" SELECTED")
					end if
	
				Response.Write(">" & Replace(rstOrderRecords.Fields(0).Value, "_", " ") & "</OPTION>" & vbCrLf)

					rstOrderRecords.MoveNext
				loop

				' Release the ADO recordset object.
				rstOrderRecords.close
			rstOrderRecords = Nothing
			end if
	
			' Release the ADO command object.
		cmdOrderRecords = Nothing
		end if
%>
												</SELECT>
											</TD>
											<TD width=10 height=10>
												<INPUT type="button" value="Go" id=btnGoOrder name=btnGoOrder class="btn"
												    onclick="goOrder()"
                                                    onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                                    onfocus="try{button_onFocus(this);}catch(e){}"
                                                    onblur="try{button_onBlur(this);}catch(e){}" />
											</TD>
										</TR>
									</table>
								</td>
							</tr>
							<tr height=10>
								<td height=10>&nbsp;</td>
							</tr>
							<TR>
								<TD>
									<OBJECT classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1" id=ssOleDBGridSelRecords name=ssOleDBGridSelRecords codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" style="LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:400px;">
										<PARAM NAME="ScrollBars" VALUE="4">
										<PARAM NAME="_Version" VALUE="196617">
										<PARAM NAME="DataMode" VALUE="2">
										<PARAM NAME="Cols" VALUE="0">
										<PARAM NAME="Rows" VALUE="0">
										<PARAM NAME="BorderStyle" VALUE="1">
										<PARAM NAME="RecordSelectors" VALUE="0">
										<PARAM NAME="GroupHeaders" VALUE="0">
										<PARAM NAME="ColumnHeaders" VALUE="-1">
										<PARAM NAME="GroupHeadLines" VALUE="1">
										<PARAM NAME="HeadLines" VALUE="1">
										<PARAM NAME="FieldDelimiter" VALUE="(None)">
										<PARAM NAME="FieldSeparator" VALUE="(Tab)">
										<PARAM NAME="Col.Count" VALUE="0">
										<PARAM NAME="stylesets.count" VALUE="0">
										<PARAM NAME="TagVariant" VALUE="EMPTY">
										<PARAM NAME="UseGroups" VALUE="0">
										<PARAM NAME="HeadFont3D" VALUE="0">
										<PARAM NAME="Font3D" VALUE="0">
										<PARAM NAME="DividerType" VALUE="3">
										<PARAM NAME="DividerStyle" VALUE="1">
										<PARAM NAME="DefColWidth" VALUE="0">
										<PARAM NAME="BeveColorScheme" VALUE="2">
										<PARAM NAME="BevelColorFrame" VALUE="-2147483642">
										<PARAM NAME="BevelColorHighlight" VALUE="-2147483628">
										<PARAM NAME="BevelColorShadow" VALUE="-2147483632">
										<PARAM NAME="BevelColorFace" VALUE="-2147483633">
										<PARAM NAME="CheckBox3D" VALUE="-1">
										<PARAM NAME="AllowAddNew" VALUE="0">
										<PARAM NAME="AllowDelete" VALUE="0">
										<PARAM NAME="AllowUpdate" VALUE="0">
										<PARAM NAME="MultiLine" VALUE="0">
										<PARAM NAME="ActiveCellStyleSet" VALUE="">
										<PARAM NAME="RowSelectionStyle" VALUE="0">
										<PARAM NAME="AllowRowSizing" VALUE="0">
										<PARAM NAME="AllowGroupSizing" VALUE="0">
										<PARAM NAME="AllowColumnSizing" VALUE="-1">
										<PARAM NAME="AllowGroupMoving" VALUE="0">
										<PARAM NAME="AllowColumnMoving" VALUE="0">
										<PARAM NAME="AllowGroupSwapping" VALUE="0">
										<PARAM NAME="AllowColumnSwapping" VALUE="0">
										<PARAM NAME="AllowGroupShrinking" VALUE="0">
										<PARAM NAME="AllowColumnShrinking" VALUE="0">
										<PARAM NAME="AllowDragDrop" VALUE="0">
										<PARAM NAME="UseExactRowCount" VALUE="-1">
										<PARAM NAME="SelectTypeCol" VALUE="0">
										<PARAM NAME="SelectTypeRow" VALUE="3">
										<PARAM NAME="SelectByCell" VALUE="-1">
										<PARAM NAME="BalloonHelp" VALUE="0">
										<PARAM NAME="RowNavigation" VALUE="1">
										<PARAM NAME="CellNavigation" VALUE="0">
										<PARAM NAME="MaxSelectedRows" VALUE="0">
										<PARAM NAME="HeadStyleSet" VALUE="">
										<PARAM NAME="StyleSet" VALUE="">
										<PARAM NAME="ForeColorEven" VALUE="0">
										<PARAM NAME="ForeColorOdd" VALUE="0">
										<PARAM NAME="BackColorEven" VALUE="16777215">
										<PARAM NAME="BackColorOdd" VALUE="16777215">
										<PARAM NAME="Levels" VALUE="1">
										<PARAM NAME="RowHeight" VALUE="503">
										<PARAM NAME="ExtraHeight" VALUE="0">
										<PARAM NAME="ActiveRowStyleSet" VALUE="">
										<PARAM NAME="CaptionAlignment" VALUE="2">
										<PARAM NAME="SplitterPos" VALUE="0">
										<PARAM NAME="SplitterVisible" VALUE="0">
										<PARAM NAME="Columns.Count" VALUE="0">
										<PARAM NAME="UseDefaults" VALUE="-1">
										<PARAM NAME="TabNavigation" VALUE="1">
										<PARAM NAME="_ExtentX" VALUE="17330">
										<PARAM NAME="_ExtentY" VALUE="1323">
										<PARAM NAME="_StockProps" VALUE="79">
										<PARAM NAME="Caption" VALUE="">
										<PARAM NAME="ForeColor" VALUE="0">
										<PARAM NAME="BackColor" VALUE="16777215">
										<PARAM NAME="Enabled" VALUE="-1">
										<PARAM NAME="DataMember" VALUE="">
										<PARAM NAME="Row.Count" VALUE="0">
									</OBJECT>
								</TD>
							</TR>
						</TABLE>
<%	
	end if
%>
	                    <INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value="<%=sErrorDescription%>">

					</td>
					<td width=20></td>
				</tr>
				<tr height=10>
					<td height=10 colspan=3>&nbsp;</td>
				</tr>
				<tr height=10>
					<td width=20></td>
					<td height=10>
						<TABLE WIDTH=100% class="invisible" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD>&nbsp;</TD>
								<TD width=10>
									<INPUT id=cmdOK type=button value=OK name=cmdOK style="WIDTH: 80px" width="80" class="btn"
									    onclick="makeSelection()" 
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
								<TD width=10>&nbsp;</TD>
								<TD width=10>
									<INPUT id=cmdCancel type=button value=Cancel name=cmdCancel style="WIDTH: 80px" width="80" class="btn"
									    onclick="window.parent.close();" 
                                        onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                        onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                        onfocus="try{button_onFocus(this);}catch(e){}"
                                        onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
							</TR>
						</TABLE>
					</td>
					<td width=20></td>
				</tr>
			</TABLE>
		</td>
	</tr>
</table>

<INPUT type='hidden' id="txtTicker" name="txtTicker" value="0">
<INPUT type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">

<FORM id="frmUseful" name="frmUseful" style="visibility:hidden;display:none">
	<INPUT type="hidden" id="txtIEVersion" name="txtIEVersion" value="<%=session("IEVersion")%>">
	<INPUT type='hidden' id="txtSelectionType" name="txtSelectionType" value="<%=session("selectionType")%>">
	<INPUT type='hidden' id="txtTableID" name="txtTableID" value="<%=session("TB_EmpTableID")%>">
	<INPUT type="hidden" id="txtMenuSaved" name="txtMenuSaved" value=0>
</FORM>

<form name="frmPrompt" method="post" action="promptedValues" id="frmPrompt" style="visibility:hidden;display:none">
	<input type="hidden" id="filterID" name="filterID">
</form>

</div>

<script type="text/javascript"> tbBulkBookingSelection_onload();</script>
