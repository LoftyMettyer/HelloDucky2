<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<script type="text/javascript">

		function stdrpt_AbsenceCalendar_window_onload() {

			$("#optionframe").attr("data-framesource", "STDRPT_ABSENCECALENDAR");

			$("#DisplayAbsenceCalendarEventDetail").dialog({
				autoOpen: false,
				modal: true,
				width: 450,
				height: 323,
				resizable: false
			});

				var fOK;
				fOK = true;

				// Permission denied on absence table
				if (frmChangeDetails.txtReportFailed.value == 'True') {
						OpenHR.messageBox(frmChangeDetails.txtErrorMSG.value, 48, "Absence Calendar");
						absence_calendar_OKClick();
						return;
				}

				if (fOK == true) {
					// Set focus onto one of the form controls. 
					// cmdOK.focus(); NHRD This line was erroring
					$("#cmdOK").focus(); //This is the back button on the Absence Calendar
					
						// Get menu.asp to refresh the menu.
						menu_refreshMenu();
						refreshDateSpecifics();
						// Expand the option frame and hide the work frame.
						$("#workframe").hide();
						$("#optionframe").show();
				}

				showDefaultRibbon();
				$("#toolbarHome").click();

		}

		function cboStartMonth_onchange() 
		{
				frmChangeDetails.txtStartMonth.value = cboStartMonth.value;
				OpenHR.submitForm(frmChangeDetails);
		}

		function cmdPreviousYear_onclick() {
				frmChangeDetails.txtStartYear.value = Number(frmChangeDetails.txtStartYear.value) - 1;
				OpenHR.submitForm(frmChangeDetails);
		}

		function cmdNextYear_onclick() {
				frmChangeDetails.txtStartYear.value = Number(frmChangeDetails.txtStartYear.value) + 1;
				OpenHR.submitForm(frmChangeDetails);
		}

		function refreshToggleValues() {

			// Show Captions setting
			if (chkShowCaptions.checked == false) {
				$("#txtShowCaptions").val("hide");
			}
			else {
				$("#txtShowCaptions").val("show");
			}

			// Show Weekends setting
			if (chkShowWeekends.checked == false) {
				$("#txtShowWeekends").val("unhighlighted");
			}
			else {
				$("#txtShowWeekends").val("highlighted");
			}

			// Include Bank Holidays setting
			if (chkIncludeBankHolidays.checked == false) {
				$("#txtIncludeBankHolidays").val("unincluded");
			}
			else {
				$("#txtIncludeBankHolidays").val("included");
			}

			// Show Bank Holidays setting
			if (chkShowBankHolidays.checked == false) {
				$("#txtShowBankHolidays").val("unhighlighted");
			}
			else {
				$("#txtShowBankHolidays").val("highlighted");
			}

			// Working Days Only setting
			if (chkIncludeWorkingDaysOnly.checked == false) {
				$("#txtIncludeWorkingDaysOnly").val("unincluded");
			}
			else {
				$("#txtIncludeWorkingDaysOnly").val("included");
			}
		}

		function ShowDetails(pdStartDate, pstrStartSession, pdEndDate, pstrEndSession, intDuration, strType, strTypeCode, strCalCode, strReason, strRegion, strWorkingPattern, bIsWorkingDay) {

			if (bIsWorkingDay || !$("#chkIncludeWorkingDaysOnly")[0].checked) {

				var frmAbsenceDetails = OpenHR.getForm("divAbsenceCalendarEventDetail", "frmAbsenceDetails");
				frmAbsenceDetails.txtStartDate.value = pdStartDate;
				frmAbsenceDetails.txtStartSession.value = pstrStartSession;
				frmAbsenceDetails.txtEndDate.value = pdEndDate;
				frmAbsenceDetails.txtEndSession.value = pstrEndSession;
				frmAbsenceDetails.txtDuration.value = intDuration;
				frmAbsenceDetails.txtType.value = strType;
				frmAbsenceDetails.txtTypeCode.value = strTypeCode;
				frmAbsenceDetails.txtCalCode.value = strCalCode;
				frmAbsenceDetails.txtReason.value = strReason;
				frmAbsenceDetails.txtRegion.value = strRegion;
				frmAbsenceDetails.txtWorkingPattern.value = strWorkingPattern;

				OpenHR.submitForm(frmAbsenceDetails, "DisplayAbsenceCalendarEventDetail");
				$("#DisplayAbsenceCalendarEventDetail").dialog("open");
				$("#DisplayAbsenceCalendarEventDetail").dialog("option", "position", ['center', 'center']);
				$("#DisplayAbsenceCalendarEventDetail").dialog("option", "height", 340);
			}
		}

		// Returns to the recordedit screen
		function absence_calendar_OKClick() {
			
			refreshData();

			menu_disableMenu();

			$("#optionframe").hide();
			$("#workframe").show();
			OpenHR.submitForm(frmRecordEdit);
			menu_refreshMenu();
			$("#toolbarRecord").click();

			}

		function PrintGrid() {
			var divToPrint = $("#optionframe").html();
			var newWin = window.open("", "_blank", 'toolbar=no,status=no,menubar=no,scrollbars=yes,resizable=yes, width=1, height=1, visible=none', "");
						
			newWin.document.write("<link href=\"" + window.ROOT + "Content/OpenHR.css" + "\" rel=\"stylesheet\" />");
			newWin.document.write("<link href=\"" + window.ROOT + "Content/table.css" + "\" rel=\"stylesheet\" />");

			var headstr = "<html><head><title></title></head><body>";
			var footstr = "</body>";

			newWin.document.write(headstr + divToPrint + footstr);

			//Hide buttons so they don't show in the printout
			var elementToHide = newWin.document.getElementById("cmdPrint");
			elementToHide.style.display = "none";
			elementToHide = newWin.document.getElementById("cmdOK");
			elementToHide.style.display = "none";
			elementToHide = newWin.document.getElementById("cmdPreviousYear");
			elementToHide.style.display = "none";
			elementToHide = newWin.document.getElementById("cmdNextYear");
			elementToHide.style.display = "none";
			newWin.document.close();
			newWin.focus();
			newWin.print();
			newWin.close();
		}

</script>

<%
	Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)

	If Session("stdrpt_AbsenceCalendar_StartMonth") = "" Then

		Session("stdrpt_AbsenceCalendar_StartMonth") = objDatabase.GetModuleParameter("MODULE_ABSENCE", "Param_FieldStartMonth")

		If Month(Now) < CInt(Session("stdrpt_AbsenceCalendar_StartMonth")) Then
			Session("stdrpt_AbsenceCalendar_StartYear") = Year(Now) - 1
		Else
			Session("stdrpt_AbsenceCalendar_StartYear") = Year(Now)
		End If

	End If

	' Create absence calendar object
	Dim objAbsenceCalendar As New AbsenceCalendar
	objAbsenceCalendar.SessionInfo = CType(Session("SessionContext"), SessionInfo)

	' Pass in the recordID for the current record
	objAbsenceCalendar.RealSource = Session("optionRealsource")
	objAbsenceCalendar.RecordID = Session("optionRecordID")
		objAbsenceCalendar.ClientDateFormat = Session("LocaleDateFormat").ToString()
		objAbsenceCalendar.StartMonth = Session("stdrpt_AbsenceCalendar_StartMonth")
	objAbsenceCalendar.StartYear = Session("stdrpt_AbsenceCalendar_StartYear")

	objAbsenceCalendar.Initialise()

	objAbsenceCalendar.IncludeBankHolidays = session("stdrpt_AbsenceCalendar_IncludeBankHolidays")
	objAbsenceCalendar.IncludeWorkingDaysOnly = session("stdrpt_AbsenceCalendar_IncludeWorkingDaysOnly")
	objAbsenceCalendar.ShowBankHolidays = session("stdrpt_AbsenceCalendar_ShowBankHolidays") 		
	objAbsenceCalendar.ShowWeekends = session("stdrpt_AbsenceCalendar_ShowWeekends")
	objAbsenceCalendar.ShowCaptions = session("stdrpt_AbsenceCalendar_ShowCaptions")
	
	objAbsenceCalendar.StartMonth = Session("stdrpt_AbsenceCalendar_StartMonth")
	objAbsenceCalendar.StartYear = Session("stdrpt_AbsenceCalendar_StartYear")

		
if objAbsenceCalendar.ReportFailed = false then 
%>
		<table valign=top align="center" class="outline" id="Background" cellSpacing="2" cellPadding="0">
				<tr>
					<td valign=top><!-- Display the month details -->
<%
		Response.Write(objAbsenceCalendar.HTML_Calendar)
%>
						</td>
					<td valign=top>
						<Table class="invisible"> 
								<TR height=8>
										<TD colspan=2></TD>
								</TR>
								<TR>
										<TD colspan=2>
											<!-- Draw the Employee information box -->
											<table valign=top width="250" class="outline" id="tblEmpoyeeInformation" cellSpacing="2" cellPadding="0">
<%
				' Write a row in this table for the forward/back year controls
		Response.Write(objAbsenceCalendar.HTML_ForwardBackYear)
		' Stuff the employee information
		Response.Write(objAbsenceCalendar.HTML_EmployeeInformation)
%>		  
														</table>
										</TD>
								</TR>
				
								<TR height=3>
										<TD colspan=2></TD>
								</TR>
								<TR>
										<TD colspan=2>
											<!-- Draw the option checkboxes -->
											<table width="250" class="outline" id="tblOptions" cellPadding="0" cellSpacing="2">
													<tr>		  
															<td>&nbsp;Start Month</td>
															<td>
<%
				' Load the start month combo
		Response.Write(objAbsenceCalendar.HTML_SelectedStartMonthCombo(objAbsenceCalendar.StartMonth))
%>
																		</td>
													</tr>
												<!-- Show the display options -->
<%
		Response.Write(objAbsenceCalendar.HTML_DisplayOptions)
%>
														</table>
										</TD>
								</TR>
						
								<TR height=3>
										<TD colspan=2></TD>
								</TR>
										<TR>
										<TD colspan=2>
<% 
				' Generate HTML for the absence key types
		Response.Write(objAbsenceCalendar.HTML_LoadColourKey)
%>
												</TD>
								</TR>
								<TR height=3>
										<TD colspan=2></TD>
								</TR>
										<TR>
										<!-- OK/Print Buttons -->
										<td colspan=2 align=right>
				            <input id="cmdPrint" name="cmdPrint" type="button" value="Print" style="WIDTH: 80px" class="btn"
												onclick="PrintGrid()" />
														&nbsp; 
										<input id="cmdOK" name="cmdOK" type="button" value="Back" style="WIDTH: 80px" class="btn"
												onclick="absence_calendar_OKClick()" />
												</td>
								</TR>
							</TABLE>
						</TD>
			</tr>
		</table>

<%
	'Populate the grid with data
	objAbsenceCalendar.StartYear = session("stdrpt_AbsenceCalendar_StartYear")
	objAbsenceCalendar.StartMonth = session("stdrpt_AbsenceCalendar_StartMonth")
	
	' Write navigation/option functions
		Response.Write(objAbsenceCalendar.HTML_ToggleDisplay)
end if 
%>

<!-- Data for the absence calendar -->
<form action="stdrpt_AbsenceCalendar_submit" method="post" id="frmChangeDetails" name="frmChangeDetails">
		<input type="hidden" id="txtStartMonth" name="txtStartMonth" value="<%Response.Write(objAbsenceCalendar.StartMonth)%>">
		<input type="hidden" id="txtStartYear" name="txtStartYear" value="<%Response.Write(objAbsenceCalendar.StartYear)%>">
		<input type="hidden" id="txtIncludeBankHolidays" name="txtIncludeBankHolidays" value="<%Response.Write(Session("stdrpt_AbsenceCalendar_IncludeBankHolidays"))%>">
		<input type="hidden" id="txtIncludeWorkingDaysOnly" name="txtIncludeWorkingDaysOnly" value="<%Response.Write(Session("stdrpt_AbsenceCalendar_IncludeWorkingDaysOnly"))%>">
		<input type="hidden" id="txtShowBankHolidays" name="txtShowBankHolidays" value="<%Response.Write(Session("stdrpt_AbsenceCalendar_ShowBankHolidays"))%>">
		<input type="hidden" id="txtShowCaptions" name="txtShowCaptions" value="<%Response.Write(Session("stdrpt_AbsenceCalendar_ShowCaptions"))%>">
		<input type="hidden" id="txtShowWeekends" name="txtShowWeekends" value="<%Response.Write(Session("stdrpt_AbsenceCalendar_ShowWeekends"))%>">
		<input type="hidden" id="txtAbsenceRecordsFound" name="txtAbsenceRecordsFound" value="<%Response.Write(objAbsenceCalendar.AbsenceRecordCount)%>">
		<input type="hidden" id="txtReportFailed" name="txtReportFailed" value="<%Response.Write(objAbsenceCalendar.ReportFailed)%>">
		<input type="hidden" id="txtErrorMSG" name="txtErrorMSG" value="<%Response.Write(objAbsenceCalendar.ErrorMSG)%>">
		<input type="hidden" id="txtDisableRegions" name="txtDisableRegions" value="<%Response.Write(objAbsenceCalendar.DisableRegions)%>">
		<input type="hidden" id="txtDisableWPs" name="txtDisableWPs" value="<%Response.Write(objAbsenceCalendar.DisableWPs)%>">
		<%=Html.AntiForgeryToken()%>
</form>

<form action="emptyoption_submit" method="post" id="frmGotoOption" name="frmGotoOption" style="visibility: hidden; display: none">
		<%Html.RenderPartial("~/Views/Shared/gotoOption.ascx")%>
		<%=Html.AntiForgeryToken()%>
</form>

<!-- Form to return to record edit screen -->
<form action="emptyoption" method="post" id="frmRecordEdit" name="frmRecordEdit">
</form>

<div id="divAbsenceCalendarEventDetail">
	<form id="frmAbsenceDetails" name="frmAbsenceDetails" action="stdrpt_AbsenceCalendar_Details" method="post" style="visibility: hidden; display: none">
			<input type="hidden" id="txtStartDate" name="txtStartDate">
			<input type="hidden" id="txtStartSession" name="txtStartSession">
			<input type="hidden" id="txtEndDate" name="txtEndDate">
			<input type="hidden" id="txtEndSession" name="txtEndSession">
			<input type="hidden" id="txtDuration" name="txtDuration">
			<input type="hidden" id="txtType" name="txtType">
			<input type="hidden" id="txtTypeCode" name="txtTypeCode">
			<input type="hidden" id="txtCalCode" name="txtcalCode">
			<input type="hidden" id="txtReason" name="txtReason">
			<input type="hidden" id="txtRegion" name="txtRegion">
			<input type="hidden" id="txtWorkingPattern" name="txtWorkingPattern">
			<input type="hidden" id="txtDisableRegions" name="txtDisableRegions" value="<%Response.Write(objAbsenceCalendar.DisableRegions)%>">
			<input type="hidden" id="txtDisableWPs" name="txtDisableWPs" value="<%Response.Write(objAbsenceCalendar.DisableWPs)%>">
	</form>
</div>

<div id="DisplayAbsenceCalendarEventDetail"></div>

<% 
	' Cleanup code
		objAbsenceCalendar = Nothing
%>

<script type="text/javascript">

	$(document).ready(function () {
		stdrpt_AbsenceCalendar_window_onload();
	});

</script>