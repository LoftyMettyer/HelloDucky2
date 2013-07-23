﻿<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script src="<%: Url.Content("~/bundles/utilities_calendarreport_run")%>" type="text/javascript"></script>

<%
	Dim objCalendar As HR.Intranet.Server.CalendarReport
	objCalendar = Session("objCalendar" & Session("CalRepUtilID"))
%>

<script type="text/javascript">

	function populateCTL_Collections() {

		var frmCalendar = OpenHR.getForm("calendarframe_calendar", "frmCalendar");
		var frmUseful = OpenHR.getForm("calendarworkframe", "frmUseful");

		if (frmCalendar.txtGroupByDesc.value == 1) {
			frmUseful.txtCTLsPopulated.value = 1;
			return;
		}

		//var docCalendar = window.parent.frames("calendarframe_calendar").document;

		var objBaseCTL;
		var vControlName;
		var strSession;
		var dtLabelsDate = new Date();
		var lblTemp;
		var INPUT_VALUE = new String("");
		var intWPCOUNT = new Number(0);
		var intBHolCOUNT = new Number(0);
		var intRegionCOUNT = new Number(0);
		var intBaseID = new Number(0);

		for (var i = 1; i <= Number(frmCalendar.txtBaseCtlCount.value) ; i++) {
			vControlName = 'ctlCalRec_' + i;
			objBaseCTL = document.getElementById(vControlName);
			intBaseID = objBaseCTL.BaseDescTag;

				<%
	If objCalendar.StaticWP = True Then
				%>
			//add the Static WP.
			INPUT_VALUE = new String("");
			vControlName = "";
			objBaseCTL.StaticWP_Populated = true;
			objBaseCTL.HistoricWP_Populated = false;
			vControlName = 'txtWPCOUNT_' + intBaseID;
			try {
				intWPCOUNT = Number(document.getElementById(vControlName).value);
			}
			catch (e) {
				intWPCOUNT = 0;
			}
			for (var iElement = 1; iElement <= intWPCOUNT; iElement++) {
				vControlName = 'txtWP_' + intBaseID;
				INPUT_VALUE = document.getElementById(vControlName).value;
				objBaseCTL.AddWorkingPattern(INPUT_VALUE, true);
			}
				<% 
Else
				%>

			//add all the historic WPs.
			INPUT_VALUE = new String("");
			vControlName = "";
			objBaseCTL.StaticWP_Populated = false
			objBaseCTL.HistoricWP_Populated = true;
			vControlName = 'txtWPCOUNT_' + intBaseID;
			try {
				intWPCOUNT = Number(document.getElementById(vControlName).value);
			}
			catch (e) {
				intWPCOUNT = 0;
			}
			for (var iElement = 1; iElement <= intWPCOUNT; iElement++) {
				vControlName = 'txtWP_' + intBaseID + '_' + iElement;
				INPUT_VALUE = document.getElementById(vControlName).value;
				objBaseCTL.AddWorkingPattern(INPUT_VALUE, false);
			}
				<% 
End If

If objCalendar.StaticReg = True Then
				%>
			//add the Static BHol. 
			INPUT_VALUE = new String("");
			vControlName = "";
			objBaseCTL.StaticRegion_Populated = true;
			objBaseCTL.HistoricRegion_Populated = false;
			vControlName = 'txtBHolCOUNT_' + intBaseID;
			try {
				intBHolCOUNT = Number(document.getElementById(vControlName).value);
			}
			catch (e) {
				intBHolCOUNT = 0;
			}

			for (var iElement = 1; iElement <= intBHolCOUNT; iElement++) {
				vControlName = 'txtBHol_' + intBaseID + '_' + iElement;
				INPUT_VALUE = document.getElementById(vControlName).value;
				objBaseCTL.AddBankHoliday(INPUT_VALUE, true);
			}
				<% 
Else
				%>
			//add all the historic BHols. 
			INPUT_VALUE = new String("");
			vControlName = "";
			objBaseCTL.StaticRegion_Populated = false;
			objBaseCTL.HistoricRegion_Populated = true;
			vControlName = 'txtBHolCOUNT_' + intBaseID;
			try {
				intBHolCOUNT = Number(document.getElementById(vControlName).value);
			}
			catch (e) {
				intBHolCOUNT = 0;
			}

			for (var iElement = 1; iElement <= intBHolCOUNT; iElement++) {
				vControlName = 'txtBHol_' + intBaseID + '_' + iElement;
				INPUT_VALUE = document.getElementById(vControlName).value;
				objBaseCTL.AddBankHoliday(INPUT_VALUE, false);
			}

				<% 
End If
				%>

			//add all the historic Career Changes. 
			vControlName = 'txtRegionCOUNT_' + intBaseID;
			try {
				intRegionCOUNT = Number(document.getElementById(vControlName).value);
			}
			catch (e) {
				intRegionCOUNT = 0;
			}

			for (var iElement = 1; iElement <= intRegionCOUNT; iElement++) {
				vControlName = 'txtRegion_' + intBaseID + '_' + iElement;
				INPUT_VALUE = document.getElementById(vControlName).value;

				objBaseCTL.AddCareerChange(INPUT_VALUE);
			}
		} //for (var i=1; i<=frmCalendar.txtBaseCtlCount.value; i++) 

		frmUseful.txtCTLsPopulated.value = 1;
		return true;
	}
</script>

<form name="frmNav" id="frmNav">
	<table align="center" class="invisible" cellpadding="0" cellspacing="0" width="100%" height="100%">
		<tr>
			<td>
				<table align="center" class="invisible" cellpadding="0" cellspacing="0" width="100%" height="100%">
					<tr height="47">
						<td width="*" rowspan="1">&nbsp;
						</td>
						<td align="right">
							<table class="outline" cellspacing="0" cellpadding="2" height="5">
								<tr height="5" valign="middle">
									<td width="5"></td>
									<td height="5" align="left" valign="middle" width="5">
										<img alt="First Month" align="center" valign="middle" border="0" src="images/first_disabled.gif" 
											name="imgFirstMonth" id="imgFirstMonth" width="16" height="16"
											title="First Month"
											onclick="firstMonth();">
									</td>
									<td width="5"></td>
									<td height="5" align="left" valign="middle" width="5">
										<img alt="Previous Month" align="center" valign="middle" border="0"
											src="images/previous_disabled.gif"
											name="imgPrevMonth" id="imgPrevMonth" width="16" height="16"
											title="Previous Month"
											onclick="prevMonth();">
									</td>
									<td width="5"></td>
									<td height="5" align="left" valign="top" width="100">
										<%
											Response.Write(objCalendar.HTML_MonthCombo(0))
										%>	
									</td>
									<td width="5"></td>
									<td height="5" align="left" valign="top">
										<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
											<tr>
												<td width="40">
													<input maxlength="4" value="2003" id="txtYear" 
														name="txtYear" class="text" style="WIDTH: 40px" width="40" value="0"
														title="Year"
														onkeypress="if(window.event.keyCode==13) {frmNav.txtYear.blur(); return false;}"
														onblur="setRecordsNumeric();"
														onchange="setRecordsNumeric();">
												</td>
												<td width="15" align="center">
													<input style="WIDTH: 15px;" type="button" value="+" id="cmdYearUp"
														name="cmdYearUp" class="btn"
														title="Add Year"
														onclick="spinRecords(true); setRecordsNumeric();" />
												</td>
												<td width="15" align="center">
													<input style="WIDTH: 15px;" type="button" value="-" id="cmdYearDown" 
														name="cmdYearDown" class="btn"
														title="Subtract Year"
														onclick="spinRecords(false); setRecordsNumeric();" />
												</td>
											</tr>
										</table>
									</td>
									<td width="5"></td>
									<td height="5" align="left" valign="middle" width="5">
										<img alt="Current Month" align="center" valign="middle" style="margin-top: 1px;"
											src="images/today_disabled.gif"
											title="Current Month"
											name="imgToday" id="imgToday" width="16" height="16"
											onclick="thisMonth();">
									</td>
									<td width="5"></td>
									<td height="5" align="left" valign="middle" width="5">
										<img alt="Next Month" align="center" valign="middle" border="0"
											src="images/next_disabled.gif"
											title="Next Month"
											name="imgNextMonth" id="imgNextMonth" width="16" height="16"
											onclick="nextMonth();">
									</td>
									<td width="5"></td>
									<td height="5" align="left" valign="middle" width="5">
										<img alt="Last Month" align="center" valign="middle" border="0"
											src="images/last_disabled.gif"
											title="Last Month"
											name="imgLastMonth" id="imgLastMonth" width="16" height="16"
											onclick="lastMonth();">
									</td>
									<td width="5"></td>
								</tr>
							</table>
						</td>
						<td width="62" rowspan="1">&nbsp;
						</td>
					</tr>
				</table>
			</td>
		</tr>

		<tr>
			<td>
				<table align="center" class="invisible" cellpadding="0" cellspacing="0" width="100%">
					<tr height="40">
						<td align="right" nowrap width="100%" colspan="2">
							<table class="invisible" cellspacing="0" cellpadding="0" width="100%" height="100%">
								<tr>
									<td width="100%">
										<object classid="CLSID:41021C13-8D42-4364-8388-9506F0755AE3"
											codebase="cabs/COAInt_CalRepDates.cab#version=1,0,0,2"
											id="ctlDates" name="ctlDates" style="VISIBILITY: visible; WIDTH: 100%"
											width="100%" 
											height="50px"
											viewastext>
											<param name="BackColor" value="16513017">
										</object>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr height="5"></tr>
				</table>
			</td>
		</tr>
	</table>
</form>

<form id="frmUseful" name="frmUseful" style="visibility: hidden; display: none">
	<input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
	<input type="hidden" id="txtLoading" name="txtLoading" value="1">
	<input type="hidden" id="txtChangingDate" name="txtChangingDate" value="0">
	<input type="hidden" id="txtCurrentBaseTableID" name="txtCurrentBaseTableID">
	<input type="hidden" id="txtAvailableColumnsLoaded" name="txtAvailableColumnsLoaded" value="0">
	<input type="hidden" id="txtEventsLoaded" name="txtEventsLoaded" value="0">
	<input type="hidden" id="txtSortLoaded" name="txtSortLoaded" value="0">
	<input type="hidden" id="txtChanged" name="txtChanged" value="0">
	<input type="hidden" id="txtUtilID" name="txtUtilID" value="<%=session("utilid")%>">
	<input type="hidden" id="txtUtilType" name="txtUtilType" value="<%=session("utiltype")%>">
	<input type="hidden" id="txtEventCount" name="txtEventCount" value="<%=session("eventcount")%>">
	<input type="hidden" id="txtHiddenEventFilterCount" name="txtHiddenEventFilterCount" value="<%=session("hiddenfiltercount")%>">
	<input type="hidden" id="txtLockGridEvents" name="txtLockGridEvents" value="0">
	<input type="hidden" id="txtCTLsPopulated" name="txtCTLsPopulated" value="0">
	<%
		Dim cmdDefinition As Object
		Dim prmModuleKey As Object
		Dim prmParameterKey As Object
		Dim prmParameterValue As Object
				
		cmdDefinition = CreateObject("ADODB.Command")
		cmdDefinition.CommandText = "sp_ASRIntGetModuleParameter"
		cmdDefinition.CommandType = 4	' Stored procedure.
		cmdDefinition.ActiveConnection = Session("databaseConnection")

		prmModuleKey = cmdDefinition.CreateParameter("moduleKey", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
		cmdDefinition.Parameters.Append(prmModuleKey)
		prmModuleKey.value = "MODULE_PERSONNEL"

		prmParameterKey = cmdDefinition.CreateParameter("paramKey", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
		cmdDefinition.Parameters.Append(prmParameterKey)
		prmParameterKey.value = "Param_TablePersonnel"

		prmParameterValue = cmdDefinition.CreateParameter("paramValue", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
		cmdDefinition.Parameters.Append(prmParameterValue)

		Err.Clear()
		cmdDefinition.Execute()

		Response.Write("<INPUT type='hidden' id=txtPersonnelTableID name=txtPersonnelTableID value=" & cmdDefinition.Parameters("paramValue").value & ">" & vbCrLf)
	
		cmdDefinition = Nothing

		Dim sErrorDescription As String
		
		Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtMode name=txtMode value=" & Session("action") & ">" & vbCrLf)
	%>
</form>

<form id="frmDate" name="frmDate" style="visibility: hidden; display: none">
	<input type="hidden" id="txtFirstDayOfMonth" name="txtFirstDayOfMonth">
	<input type="hidden" id="txtDaysInMonth" name="txtDaysInMonth">
	<input type="hidden" id="txtDayControlCount" name="txtDayControlCount" value="37">
	<%
		Response.Write("	<INPUT type=""hidden"" id=txtReportStartDate name=txtReportStartDate value=""" & objCalendar.ReportStartDate_CalendarString & """>" & vbCrLf)
		Response.Write("	<INPUT type=""hidden"" id=txtReportEndDate name=txtReportEndDate value=""" & objCalendar.ReportEndDate_CalendarString & """>" & vbCrLf)
	%>
	<input type="hidden" id="txtStartOnCurrentMonth" name="txtStartOnCurrentMonth" value="
<% 
		If objCalendar.StartOnCurrentMonth Then
%> 
				1
<% 
	Else
%>
				0
<% 
	End If
%>
			">

	<input type="hidden" id="txtClientDateFormat" name="txtClientDateFormat" value="<%=Session("LocaleDateFormat")%>">
	<input type="hidden" id="txtClientDateSeparator" name="txtClientDateSeparator" value="<%=session("LocaleDateSeparator")%>">
	<input type="hidden" id="txtCurrentMonth" name="txtCurrentMonth">
	<input type="hidden" id="txtCurrentYear" name="txtCurrentYear">
	<input type="hidden" id="txtCurrentMonthIndex" name="txtCurrentMonthIndex">
	<input type="hidden" id="txtCurrentYearValue" name="txtCurrentYearValue">
</form>

<form id="frmBankHolidays" name="frmBankHolidays" style="visibility: hidden; display: none">
	<input type="hidden" id="Hidden1" name="txtFirstDayOfMonth">
</form>
<%
	Dim mblnOutputPreview As Boolean
	Dim mlngOutputFormat As Long
	Dim mblnOutputScreen As Boolean
	Dim mblnOutputPrinter As Boolean
	Dim mstrOutputPrinterName As String
	Dim mblnOutputSave As Boolean
	Dim mlngOutputSaveExisting As Long
	Dim mblnOutputEmail As Boolean
	Dim mlngOutputEmailID As Long
	Dim mstrOutputEmailName As String
	Dim mstrOutputEmailSubject As String
	Dim mstrOutputEmailAttachAs As String
	Dim mstrOutputFilename As String

	mblnOutputPreview = objCalendar.OutputPreview
	mlngOutputFormat = objCalendar.OutputFormat
	mblnOutputScreen = objCalendar.OutputScreen
	mblnOutputPrinter = objCalendar.OutputPrinter
	mstrOutputPrinterName = objCalendar.OutputPrinterName
	mblnOutputSave = objCalendar.OutputSave
	mlngOutputSaveExisting = objCalendar.OutputSaveExisting
	mblnOutputEmail = objCalendar.OutputEmail
	mlngOutputEmailID = objCalendar.OutputEmailID
	mstrOutputEmailName = objCalendar.OutputEmailGroupName
	mstrOutputEmailSubject = objCalendar.OutputEmailSubject
	mstrOutputEmailAttachAs = objCalendar.OutputEmailAttachAs
	mstrOutputFilename = objCalendar.OutputFilename
%>

<form target="Output" action="util_run_outputoptions" method="post" id="frmExportData" name="frmExportData">
	<input type="hidden" id="txtPreview" name="txtPreview" value="<%=mblnOutputPreview%>">
	<input type="hidden" id="txtFormat" name="txtFormat" value="<%=mlngOutputFormat%>">
	<input type="hidden" id="txtScreen" name="txtScreen" value="<%=mblnOutputScreen%>">
	<input type="hidden" id="txtPrinter" name="txtPrinter" value="<%=mblnOutputPrinter%>">
	<input type="hidden" id="txtPrinterName" name="txtPrinterName" value="<%=mstrOutputPrinterName%>">
	<input type="hidden" id="txtSave" name="txtSave" value="<%=mblnOutputSave%>">
	<input type="hidden" id="txtSaveExisting" name="txtSaveExisting" value="<%=mlngOutputSaveExisting%>">
	<input type="hidden" id="txtEmail" name="txtEmail" value="<%=mblnOutputEmail%>">
	<input type="hidden" id="txtEmailAddr" name="txtEmailAddr" value="<%=mlngOutputEmailID%>">
	<input type="hidden" id="txtEmailAddrName" name="txtEmailAddrName" value="<%=replace(mstrOutputEmailName, """", "&quot;")%>">
	<input type="hidden" id="txtEmailSubject" name="txtEmailSubject" value="<%=replace(mstrOutputEmailSubject, """", "&quot;")%>">
	<input type="hidden" id="txtEmailAttachAs" name="txtEmailAttachAs" value="<%=replace(mstrOutputEmailAttachAs, """", "&quot;")%>">
	<input type="hidden" id="txtFileName" name="txtFileName" value="<%=mstrOutputFilename%>">
	<input type="hidden" id="Hidden2" name="txtUtilType" value="<%=session("utilType")%>">
</form>

<%
	Dim objUser As HR.Intranet.Server.clsSettings
		
	'**************************************************************
	'Output forms and the respective elements for Region/BHol/WPs
	
	Response.Write(objCalendar.Write_Static_Historic_Forms)
	
	'**************************************************************

	'Write the function that Outputs the report to the Output Classes in the Client DLL.
		
	Response.Write("<script type=""text/javascript"">" & vbCrLf)
	Response.Write("function outputReport() " & vbCrLf)
	Response.Write("	{" & vbCrLf & vbCrLf)
	
	Response.Write("	var frmOutput = openHR.getForm(""dataframe"",""frmCalendarData"");" & vbCrLf)

	Response.Write("	var lngPageColumnCount = 3;" & vbCrLf)
	Response.Write("    var lngActualRow = new Number(0);" & vbCrLf)
	Response.Write("    var blnSettingsDone = false;" & vbCrLf)
	Response.Write("	var sColHeading = new String(''); " & vbCrLf)
	Response.Write("	var iColDataType = new Number(12); " & vbCrLf)
	Response.Write("	var iColDecimals = new Number(0); " & vbCrLf)
	Response.Write("    var blnNewPage = false;" & vbCrLf)
	Response.Write("    var lngPageCount = new Number(0);" & vbCrLf)

	Response.Write("  var strType = new String('');" & vbCrLf)
	Response.Write("  var lngStartCol = new Number(0);" & vbCrLf)
	Response.Write("  var lngStartRow = new Number(0);" & vbCrLf)
	Response.Write("  var lngEndCol = new Number(0);" & vbCrLf)
	Response.Write("  var lngEndRow = new Number(0);" & vbCrLf)
	Response.Write("  var lngBackCol = new Number(0);" & vbCrLf)
	Response.Write("  var lngForeCol = new Number(0);" & vbCrLf)
	Response.Write("  var blnBold = false;" & vbCrLf)
	Response.Write("  var blnUnderline = false;" & vbCrLf)
	Response.Write("  var blnGridlines = false;" & vbCrLf)
	
	objUser = New HR.Intranet.Server.clsSettings
		
	Response.Write("  window.parent.parent.ASRIntranetOutput.UserName = """ & CleanStringForJavaScript(Session("Username")) & """;" & vbCrLf)
	Response.Write("  window.parent.parent.ASRIntranetOutput.SaveAsValues = """ & CleanStringForJavaScript(Session("OfficeSaveAsValues")) & """;" & vbCrLf)
	
	Response.Write("  frmMenuFrame = window.parent.parent.opener.window.parent.frames(""menuframe"");" & vbCrLf)

	Response.Write("	window.parent.parent.ASRIntranetOutput.SettingOptions(")
	Response.Write("""" & CleanStringForJavaScript(objUser.GetUserSetting("Output", "WordTemplate", "")) & """, ")
	Response.Write("""" & CleanStringForJavaScript(objUser.GetUserSetting("Output", "ExcelTemplate", "")) & """, ")

	If (objUser.GetUserSetting("Output", "ExcelGridlines", "0") = "1") Then
		Response.Write("true, ")
	Else
		Response.Write("false, ")
	End If

	If (objUser.GetUserSetting("Output", "ExcelHeaders", "0") = "1") Then
		Response.Write("true, ")
	Else
		Response.Write("false, ")
	End If

	If (objUser.GetUserSetting("Output", "AutoFitCols", "1") = "1") Then
		Response.Write("true, ")
	Else
		Response.Write("false, ")
	End If

	If (objUser.GetUserSetting("Output", "Landscape", "1") = "1") Then
		Response.Write("true, " & vbCrLf)
	Else
		Response.Write("false, " & vbCrLf)
	End If
				 
	Response.Write("frmMenuFrame.document.all.item(""txtSysPerm_EMAILGROUPS_VIEW"").value);" & vbCrLf)

	Response.Write("  window.parent.parent.ASRIntranetOutput.SettingLocations(")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleCol", "3")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleRow", "2")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataCol", "2")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataRow", "4")) & ");" & vbCrLf)

	Response.Write("  window.parent.parent.ASRIntranetOutput.SettingTitle(")
	If (objUser.GetUserSetting("Output", "TitleGridLines", "0") = "1") Then
		Response.Write("true, ")
	Else
		Response.Write("false, ")
	End If

	If (objUser.GetUserSetting("Output", "TitleBold", "1") = "1") Then
		Response.Write("true, ")
	Else
		Response.Write("false, ")
	End If

	If (objUser.GetUserSetting("Output", "TitleUnderline", "0") = "1") Then
		Response.Write("true, ")
	Else
		Response.Write("false, ")
	End If

	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleBackcolour", "16777215")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleForecolour", "6697779")) & ", ")
	'    Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "TitleBackcolour", "16777215"))) & ", ")
	'   Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "TitleForecolour", "6697779"))) & ");" & vbCrLf)
	Response.Write("1,1);" & vbCrLf)
		
	Response.Write("window.parent.parent.ASRIntranetOutput.SettingHeading(")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingGridLines", "1")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingBold", "1")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingUnderline", "0")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingBackcolour", "16248553")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingForecolour", "6697779")) & ", ")
	'    Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "HeadingBackcolour", "16248553"))) & ", ")
	'   Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "HeadingForecolour", "6697779"))) & ");" & vbCrLf)
	Response.Write("1,1);" & vbCrLf)

	Response.Write("window.parent.parent.ASRIntranetOutput.SettingData(")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataGridLines", "1")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataBold", "0")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataUnderline", "0")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataBackcolour", "15988214")) & ", ")
	Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataForecolour", "6697779")) & ", ")
	'    Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "DataBackcolour", "15988214"))) & ", ")
	'   Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(objUser.GetUserSetting("Output", "DataForecolour", "6697779"))) & ");" & vbCrLf)
	Response.Write("1,1);" & vbCrLf)

	Response.Write("window.parent.parent.ASRIntranetOutput.SetOptions(false, " & _
																					"frmExportData.txtFormat.value, frmExportData.txtScreen.value, " & _
																					"frmExportData.txtPrinter.value, frmExportData.txtPrinterName.value, " & _
																					"frmExportData.txtSave.value, frmExportData.txtSaveExisting.value, " & _
																					"frmExportData.txtEmail.value, frmOutput.txtEmailGroupAddr.value, " & _
																					"frmExportData.txtEmailSubject.value, frmExportData.txtEmailAttachAs.value, frmExportData.txtFileName.value);" & vbCrLf)

	Response.Write("  if (frmExportData.txtFormat.value == ""0"") {" & vbCrLf)
	Response.Write("    if (frmExportData.txtPrinter.value == ""true"") {" & vbCrLf)
	Response.Write("			window.parent.parent.ASRIntranetOutput.SetPrinter();" & vbCrLf)
	Response.Write("      dataOnlyPrint();" & vbCrLf)
	Response.Write("			window.parent.parent.ASRIntranetOutput.ResetDefaultPrinter();" & vbCrLf)
	Response.Write("    }" & vbCrLf)
	Response.Write("  }" & vbCrLf)
	Response.Write("  else {" & vbCrLf)

	Response.Write("if (window.parent.parent.ASRIntranetOutput.GetFile() == true) " & vbCrLf)
	Response.Write("	{" & vbCrLf)
	Response.Write("	window.parent.parent.ASRIntranetOutput.InitialiseStyles();" & vbCrLf)
	Response.Write("	window.parent.parent.ASRIntranetOutput.ResetStyles();" & vbCrLf)
	Response.Write("	window.parent.parent.ASRIntranetOutput.ResetColumns();" & vbCrLf)
	Response.Write("	window.parent.parent.ASRIntranetOutput.ResetMerges();" & vbCrLf)

	Response.Write("	window.parent.parent.ASRIntranetOutput.HeaderRows = 1;" & vbCrLf)
	Response.Write("	window.parent.parent.ASRIntranetOutput.HeaderCols = 0;" & vbCrLf)
	Response.Write("	window.parent.parent.ASRIntranetOutput.SizeColumnsIndependently = true;" & vbCrLf)
	
	Response.Write("	window.parent.parent.ASRIntranetOutput.ArrayDim((lngPageColumnCount-1), 0);" & vbCrLf & vbCrLf)
	Response.Write("  frmOutput.grdCalendarOutput.focus();")
	
	Response.Write("  frmOutput.grdCalendarOutput.MoveFirst();" & vbCrLf)
	Response.Write("  for (var lngRow=0; lngRow<frmOutput.grdCalendarOutput.Rows; lngRow++)" & vbCrLf)
	Response.Write("		{" & vbCrLf)
	Response.Write("		bm = frmOutput.grdCalendarOutput.AddItemBookmark(lngRow);" & vbCrLf)
	
	Response.Write("		if (lngRow == (frmOutput.grdCalendarOutput.Rows - 1))" & vbCrLf)
	Response.Write("			{" & vbCrLf)
	Response.Write("			sBreakValue = frmOutput.grdCalendarOutput.Columns(1).CellText(bm);" & vbCrLf)
	Response.Write("			window.parent.parent.ASRIntranetOutput.AddPage(replace(frmOutput.grdCalendarOutput.Caption,'&&','&') + ' - ' + sBreakValue,sBreakValue);" & vbCrLf)

	Response.Write("			var frmMerge = window.parent.frames('dataframe').document.forms('frmCalendarMerge_'+lngPageCount);" & vbCrLf)
	
	Response.Write("			var dataCollection = frmMerge.elements;" & vbCrLf)
	Response.Write("			if (dataCollection!=null) " & vbCrLf)
	Response.Write("				{" & vbCrLf)
	Response.Write("				for (i=0; i<dataCollection.length; i++)  " & vbCrLf)
	Response.Write("					{" & vbCrLf)
	Response.Write("					strMergeString = dataCollection.item(i).value;" & vbCrLf)
	Response.Write("					if (strMergeString != '')" & vbCrLf)
	Response.Write("						{" & vbCrLf)
	Response.Write("						lngStartCol = Number(mergeArgument(strMergeString,'STARTCOL'));" & vbCrLf)
	Response.Write("						lngStartRow = Number(mergeArgument(strMergeString,'STARTROW'));" & vbCrLf)
	Response.Write("						lngEndCol = Number(mergeArgument(strMergeString,'ENDCOL'));" & vbCrLf)
	Response.Write("						lngEndRow = Number(mergeArgument(strMergeString,'ENDROW'));" & vbCrLf)
	Response.Write("						window.parent.parent.ASRIntranetOutput.AddMerge(lngStartCol,lngStartRow,lngEndCol,lngEndRow);" & vbCrLf)
	Response.Write("						}" & vbCrLf)
	Response.Write("					}" & vbCrLf)
	Response.Write("				}" & vbCrLf)

	Response.Write("			var frmStyle = window.parent.frames('dataframe').document.forms('frmCalendarStyle_'+lngPageCount);" & vbCrLf)
	Response.Write("			var dataCollection = frmStyle.elements;" & vbCrLf)
	Response.Write("			if (dataCollection!=null) " & vbCrLf)
	Response.Write("				{" & vbCrLf)
	Response.Write("				for (i=0; i<dataCollection.length; i++)  " & vbCrLf)
	Response.Write("					{" & vbCrLf)
	Response.Write("					strStyleString = dataCollection.item(i).value;" & vbCrLf)
	Response.Write("					if (strStyleString != '')" & vbCrLf)
	Response.Write("						{" & vbCrLf)
	Response.Write("						strType = styleArgument(strStyleString,'TYPE');" & vbCrLf)
	Response.Write("						lngStartCol = Number(styleArgument(strStyleString,'STARTCOL'));" & vbCrLf)
	Response.Write("						lngStartRow = Number(styleArgument(strStyleString,'STARTROW'));" & vbCrLf)
	Response.Write("						lngEndCol = Number(styleArgument(strStyleString,'ENDCOL'));" & vbCrLf)
	Response.Write("						lngEndRow = Number(styleArgument(strStyleString,'ENDROW'));" & vbCrLf)
	Response.Write("						lngBackCol = Number(styleArgument(strStyleString,'BACKCOLOR'));" & vbCrLf)
	Response.Write("						lngForeCol = Number(styleArgument(strStyleString,'FORECOLOR'));" & vbCrLf)
	Response.Write("						blnBold = styleArgument(strStyleString,'BOLD');" & vbCrLf)
	Response.Write("						blnUnderline = styleArgument(strStyleString,'UNDERLINE');" & vbCrLf)
	Response.Write("						blnGridlines = styleArgument(strStyleString,'GRIDLINES');" & vbCrLf)
	Response.Write("						window.parent.parent.ASRIntranetOutput.AddStyle(strType,lngStartCol,lngStartRow,lngEndCol,lngEndRow,lngBackCol,lngForeCol,blnBold,blnUnderline,blnGridlines);" & vbCrLf)
	Response.Write("						}" & vbCrLf)
	Response.Write("					}" & vbCrLf)
	Response.Write("				}" & vbCrLf)
	
	Response.Write("			for (var lngCol=0; lngCol<lngPageColumnCount; lngCol++)" & vbCrLf)
	Response.Write("				{" & vbCrLf)
	Response.Write("				window.parent.parent.ASRIntranetOutput.AddColumn(sColHeading, iColDataType, iColDecimals, false);" & vbCrLf)
	Response.Write("				}" & vbCrLf)
	Response.Write("			window.parent.parent.ASRIntranetOutput.DataArray();" & vbCrLf)
	Response.Write("			blnBreakCheck = true;" & vbCrLf)
	Response.Write("			sBreakValue = '';" & vbCrLf)
	Response.Write("			lngActualRow = 0;" & vbCrLf)
	Response.Write("			}" & vbCrLf)
	
	Response.Write("    else if ((frmOutput.grdCalendarOutput.Columns(0).CellText(bm) == '*')" & vbCrLf)
	Response.Write("					&& (!blnBreakCheck))" & vbCrLf)
	Response.Write("			{" & vbCrLf)
	Response.Write("			sBreakValue = frmOutput.grdCalendarOutput.Columns(1).CellText(bm);" & vbCrLf)
	Response.Write("			if ((sBreakValue == 'Key') && (frmExportData.txtFormat.value != '4')) " & vbCrLf)
	Response.Write("				{ " & vbCrLf)
	Response.Write("				window.parent.parent.ASRIntranetOutput.AddPage(replace(frmOutput.grdCalendarOutput.Caption,'&&','&') ,sBreakValue);" & vbCrLf)
	Response.Write("				} " & vbCrLf)
	Response.Write("			else " & vbCrLf)
	Response.Write("				{ " & vbCrLf)
	Response.Write("				window.parent.parent.ASRIntranetOutput.AddPage(replace(frmOutput.grdCalendarOutput.Caption,'&&','&') + ' - ' + sBreakValue,sBreakValue);" & vbCrLf)
	Response.Write("				} " & vbCrLf)
	
	Response.Write("			var frmMerge = window.parent.frames('dataframe').document.forms('frmCalendarMerge_'+lngPageCount);" & vbCrLf)
	Response.Write("			var dataCollection = frmMerge.elements;" & vbCrLf)
	Response.Write("			if (dataCollection!=null) " & vbCrLf)
	Response.Write("				{" & vbCrLf)
	Response.Write("				for (i=0; i<dataCollection.length; i++)  " & vbCrLf)
	Response.Write("					{" & vbCrLf)
	Response.Write("					strMergeString = dataCollection.item(i).value;" & vbCrLf)
	Response.Write("					if (strMergeString != '')" & vbCrLf)
	Response.Write("						{" & vbCrLf)
	Response.Write("						lngStartCol = Number(mergeArgument(strMergeString,'STARTCOL'));" & vbCrLf)
	Response.Write("						lngStartRow = Number(mergeArgument(strMergeString,'STARTROW'));" & vbCrLf)
	Response.Write("						lngEndCol = Number(mergeArgument(strMergeString,'ENDCOL'));" & vbCrLf)
	Response.Write("						lngEndRow = Number(mergeArgument(strMergeString,'ENDROW'));" & vbCrLf)
	Response.Write("						window.parent.parent.ASRIntranetOutput.AddMerge(lngStartCol,lngStartRow,lngEndCol,lngEndRow);" & vbCrLf)
	Response.Write("						}" & vbCrLf)
	Response.Write("					}" & vbCrLf)
	Response.Write("				}" & vbCrLf)

	Response.Write("			var frmStyle = window.parent.frames('dataframe').document.forms('frmCalendarStyle_'+lngPageCount);" & vbCrLf)
	Response.Write("			var dataCollection = frmStyle.elements;" & vbCrLf)
	Response.Write("			if (dataCollection!=null) " & vbCrLf)
	Response.Write("				{" & vbCrLf)
	Response.Write("				for (i=0; i<dataCollection.length; i++)  " & vbCrLf)
	Response.Write("					{" & vbCrLf)
	Response.Write("					strStyleString = dataCollection.item(i).value;" & vbCrLf)
	Response.Write("					if (strStyleString != '')" & vbCrLf)
	Response.Write("						{" & vbCrLf)
	Response.Write("						strType = styleArgument(strStyleString,'TYPE');" & vbCrLf)
	Response.Write("						lngStartCol = Number(styleArgument(strStyleString,'STARTCOL'));" & vbCrLf)
	Response.Write("						lngStartRow = Number(styleArgument(strStyleString,'STARTROW'));" & vbCrLf)
	Response.Write("						lngEndCol = Number(styleArgument(strStyleString,'ENDCOL'));" & vbCrLf)
	Response.Write("						lngEndRow = Number(styleArgument(strStyleString,'ENDROW'));" & vbCrLf)
	Response.Write("						lngBackCol = Number(styleArgument(strStyleString,'BACKCOLOR'));" & vbCrLf)
	Response.Write("						lngForeCol = Number(styleArgument(strStyleString,'FORECOLOR'));" & vbCrLf)
	Response.Write("						blnBold = styleArgument(strStyleString,'BOLD');" & vbCrLf)
	Response.Write("						blnUnderline = styleArgument(strStyleString,'UNDERLINE');" & vbCrLf)
	Response.Write("						blnGridlines = styleArgument(strStyleString,'GRIDLINES');" & vbCrLf)
	Response.Write("						window.parent.parent.ASRIntranetOutput.AddStyle(strType,lngStartCol,lngStartRow,lngEndCol,lngEndRow,lngBackCol,lngForeCol,blnBold,blnUnderline,blnGridlines);" & vbCrLf)
	Response.Write("						}" & vbCrLf)
	Response.Write("					}" & vbCrLf)
	Response.Write("				}" & vbCrLf)
	
	Response.Write("			for (var lngCol=0; lngCol<lngPageColumnCount; lngCol++)" & vbCrLf)
	Response.Write("				{" & vbCrLf)
	Response.Write("				window.parent.parent.ASRIntranetOutput.AddColumn(sColHeading, iColDataType, iColDecimals, false);" & vbCrLf)
	Response.Write("				}" & vbCrLf)
	Response.Write("      window.parent.parent.ASRIntranetOutput.DataArray();" & vbCrLf)
	Response.Write("			lngPageColumnCount = frmOutput.grdCalendarOutput.Columns.Count;" & vbCrLf)
	Response.Write("			if (!blnSettingsDone)" & vbCrLf)
	Response.Write("				{" & vbCrLf)
	Response.Write("				window.parent.parent.ASRIntranetOutput.HeaderRows = 2;" & vbCrLf)
	Response.Write("				window.parent.parent.ASRIntranetOutput.HeaderCols = 1;" & vbCrLf)
	Response.Write("				window.parent.parent.ASRIntranetOutput.SizeColumnsIndependently = true;" & vbCrLf)
	Response.Write("				blnSettingsDone = true;" & vbCrLf)
	Response.Write("				}" & vbCrLf)
	Response.Write("			window.parent.parent.ASRIntranetOutput.InitialiseStyles();" & vbCrLf)
	Response.Write("			window.parent.parent.ASRIntranetOutput.ResetStyles();" & vbCrLf)
	Response.Write("			window.parent.parent.ASRIntranetOutput.ResetColumns();" & vbCrLf)
	Response.Write("			window.parent.parent.ASRIntranetOutput.ResetMerges();" & vbCrLf)
	Response.Write("			lngPageCount++;" & vbCrLf)
	Response.Write("			window.parent.parent.ASRIntranetOutput.ArrayDim((lngPageColumnCount-1), 0);" & vbCrLf)
	Response.Write("			blnBreakCheck = true;" & vbCrLf)
	Response.Write("			sBreakValue = '';" & vbCrLf)
	Response.Write("			lngActualRow = 0;" & vbCrLf)
	Response.Write("			blnNewPage = true;" & vbCrLf)
	Response.Write("			}" & vbCrLf & vbCrLf)

	Response.Write("		else if (frmOutput.grdCalendarOutput.Columns(0).CellText(bm) != '*')" & vbCrLf)
	Response.Write("			{" & vbCrLf)
	Response.Write("			blnBreakCheck = false;" & vbCrLf)
	Response.Write("			blnNewPage = false;" & vbCrLf)
	Response.Write("			if (lngActualRow > 0)" & vbCrLf)
	Response.Write("				{" & vbCrLf)
	Response.Write("				window.parent.parent.ASRIntranetOutput.ArrayReDim();" & vbCrLf)
	Response.Write("				}" & vbCrLf)
	Response.Write("			for (var lngCol=0; lngCol<lngPageColumnCount; lngCol++)" & vbCrLf)
	Response.Write("				{" & vbCrLf)
	Response.Write("				window.parent.parent.ASRIntranetOutput.ArrayAddTo(lngCol, (lngActualRow), frmOutput.grdCalendarOutput.Columns(lngCol).CellText(bm));" & vbCrLf)
	Response.Write("				}" & vbCrLf)
	Response.Write("			}" & vbCrLf)
	
	
	Response.Write("		if (!blnNewPage) " & vbCrLf)
	Response.Write("			{" & vbCrLf)
	Response.Write("			lngActualRow = lngActualRow + 1; " & vbCrLf)
	Response.Write("			}" & vbCrLf)
	Response.Write("		}" & vbCrLf)
	Response.Write("		window.parent.parent.ASRIntranetOutput.Complete();" & vbCrLf)
	Response.Write("		ShowDataFrame();" & vbCrLf)
	Response.Write("		}" & vbCrLf)
	Response.Write("	}" & vbCrLf)
	
	Response.Write("	ShowDataFrame();" & vbCrLf)
	
	Response.Write("  try {" & vbCrLf)
	Response.Write("  if (frmOriginalDefinition.txtCancelPrint.value == 1) {" & vbCrLf)
	Response.Write("    OpenHR.messageBox(""Calendar Report '""+frmOriginalDefinition.txtDefn_Name.value+""' output failed.\n\nCancelled by user."",64,""Calendar Report"");" & vbCrLf)
	Response.Write("		window.focus();" & vbCrLf)
	Response.Write("  }" & vbCrLf)
	Response.Write("  else if (window.parent.parent.ASRIntranetOutput.ErrorMessage != """") {" & vbCrLf)
	Response.Write("    OpenHR.messageBox(""Calendar Report '""+frmOriginalDefinition.txtDefn_Name.value+""' output failed.\n\n""+window.parent.parent.ASRIntranetOutput.ErrorMessage,48,""Calendar Report"");" & vbCrLf)
	Response.Write("		window.focus();" & vbCrLf)
	Response.Write("		}" & vbCrLf)
	Response.Write("  else {" & vbCrLf)
	Response.Write("    OpenHR.messageBox(""Calendar Report '""+frmOriginalDefinition.txtDefn_Name.value+""' output complete."",64,""Calendar Report"");" & vbCrLf)
	Response.Write("		window.focus();" & vbCrLf)
	Response.Write("		}" & vbCrLf)
	Response.Write("	}" & vbCrLf)
	Response.Write("	catch (e) {}" & vbCrLf)
	Response.Write("	}" & vbCrLf)
	
	Response.Write("</script>" & vbCrLf & vbCrLf)

	Response.Write("<input type=hidden id=txtTitle name=txtTitle value=""" & objCalendar.CalendarReportName & """>" & vbCrLf)
	
	objCalendar = Nothing
	
%>

<form id="frmOriginalDefinition" style="visibility: hidden; display: none">
	<%
		Dim sErrMsg As String
		Response.Write("	<INPUT type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Replace(Session("utilname"), """", "&quot;") & """>" & vbCrLf)
		Response.Write("	<INPUT type='hidden' id=txtDefn_ErrMsg name=txtDefn_ErrMsg value=""" & sErrMsg & """>" & vbCrLf)
	%>
	<input type="hidden" id="Hidden3" name="txtUserName" value="<%=session("username")%>">
	<input type="hidden" id="txtDateFormat" name="txtDateFormat" value="<%=session("LocaleDateFormat")%>">
	<input type="hidden" id="txtCancelPrint" name="txtCancelPrint">
	<input type="hidden" id="txtOptionsDone" name="txtOptionsDone">
	<input type="hidden" id="txtOptionsPortrait" name="txtOptionsPortrait">
	<input type="hidden" id="txtOptionsMarginLeft" name="txtOptionsMarginLeft">
	<input type="hidden" id="txtOptionsMarginRight" name="txtOptionsMarginRight">
	<input type="hidden" id="txtOptionsMarginTop" name="txtOptionsMarginTop">
	<input type="hidden" id="txtOptionsMarginBottom" name="txtOptionsMarginBottom">
	<input type="hidden" id="txtOptionsCopies" name="txtOptionsCopies">
	<input type="hidden" id="txtCalRep_UtilID" name="txtCalRep_UtilID" value="<%=Request("CalRepUtilID")%>">
</form>


<script type="text/javascript">
	util_run_calendarreport_window_onload();
</script>
