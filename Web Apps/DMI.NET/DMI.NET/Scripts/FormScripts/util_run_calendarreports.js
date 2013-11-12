
var objCalendarKey = [];

//
//Taken from util_run_calendarrepoart_nav.ascx
//
var frmUseful = OpenHR.getForm("calendarworkframe", "frmUseful");

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

function addString(pintNumber,pstrChar)
{
		var sRetString = new String('');
	
		for (var i=1; i<=pintNumber; i++)
		{
				sRetString = sRetString + pstrChar;
		}
	
		return sRetString;
}	

function refreshCalendar() {

	var frmGetDataForm = OpenHR.getForm("dataframe", "frmCalendarGetData");

	frmGetDataForm.txtMode.value = "LOADCALENDARREPORTDATA";
	frmGetDataForm.txtMonth.value = (frmNav.cboMonth.options[frmNav.cboMonth.selectedIndex].value);
	frmGetDataForm.txtYear.value = frmNav.txtYear.value;

	refreshCalendarData();

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

function util_run_calendarreport_data_window_onload() {

	if (txtFirstLoad.value == 1) {
		$("#divReportButtons").css("visibility", "visible");
		return;
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

	var frmGetData = OpenHR.getForm("dataframe", "frmCalendarGetData");

	frmGetData.txtIncludeBankHolidays.value = (frmOptions.chkIncludeBHols.checked);
	frmGetData.txtIncludeWorkingDaysOnly.value = (frmOptions.chkIncludeWorkingDaysOnly.checked);
	frmGetData.txtShowBankHolidays.value = (frmOptions.chkShadeBHols.checked);
	frmGetData.txtShowCaptions.value = (frmOptions.chkCaptions.checked);
	frmGetData.txtShowWeekends.value = (frmOptions.chkShadeWeekends.checked);
	OpenHR.submitForm(frmGetData);

	return true;
}

//
//Taken from util_run_calendarreport_main.ascx
//

function util_run_calendarreport_main_window_onload() {
		if ((txtPreview.value == 0) && (txtOK.value == "True")) {

				outputCalendarReport();
				document.getElementById('tdDisplay').innerText = 'Calendar Report Output Complete.';
				document.getElementById('Cancel').value = 'OK';
		}
		//Replace the calendar link that is shown by 
		//default before we can get the true report name
		$("#PageDivTitle").html($("#txtTitle").val());
}
