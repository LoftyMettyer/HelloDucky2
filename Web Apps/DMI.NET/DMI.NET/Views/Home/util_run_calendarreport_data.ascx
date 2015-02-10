<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Register TagPrefix="DayPilot" Namespace="DayPilot.Web.Ui" Assembly="DayPilot" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="System.Data" %>

<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<link href="<%: Url.LatestContent("~/Themes/scheduler_white.css")%>" rel="stylesheet" type="text/css" />
<link href="<%: Url.LatestContent("~/Themes/calendar_white.css")%>" rel="stylesheet" type="text/css" />
<link href="<%: Url.LatestContent("~/Themes/layout.css")%>" rel="stylesheet" type="text/css" />
<script src="<%: Url.LatestContent("~/scripts/jquery/jquery.maskedinput.js")%>" type="text/javascript"></script>
<script src="<%: Url.LatestContent("~/scripts/jquery/MonthPicker.js")%>" type="text/javascript"></script>

<%	
	Dim objCalendar As CalendarReport = CType(Session("objCalendar" & Session("CalRepUtilID")), CalendarReport)
%>

<script type="text/javascript">

	if ($("#chkShowWeekends")[0].checked == true) {
		toggleWeekends();
	}

	<%If objCalendar.ReportStartDate > Now Or objCalendar.ReportEndDate < Now Then%>
		$('#cmdToday')[0].disabled = true;	
	<%End If%>

	$('#StartYearDemo').monthpicker({
		selectedYear: $("#txtYear").val(),
		startYear: <% =objCalendar.ReportStartDate.Year%> - 0,
		startMonth: <% =objCalendar.ReportStartDate.Month%> - 0,
		endMonth: <% =objCalendar.ReportEndDate.Month%> - 0,
		endYear: <% =objCalendar.ReportEndDate.Year%> - 0,
		pattern: 'mm/yyyy',
		openOnFocus: false
	});

	// Bind click event to the textbox to open the month picker
	$('#StartYearDemo').bind('click', function () {
		$(this).monthpicker('show');
		// When a new value is set we need to disable invalid months.
		$(this).monthpicker('disableMonths');
	});

	$('#StartYearDemo').monthpicker().bind('monthpicker-click-month', function (e, month) {
		var sMonthYear = $('#StartYearDemo').val();
		var frmGetDataForm = OpenHR.getForm("reportworkframe", "frmCalendarGetData");
		frmGetDataForm.txtMonth.value = sMonthYear.substring(0, 2);
		frmGetDataForm.txtYear.value = sMonthYear.substring(3, 7);
		OpenHR.submitForm(frmGetDataForm);
	});

	function eventCalendarClick(eventID, eventType) {

		if (eventType != "bank") {
			var frmEvent = OpenHR.getForm("divEventDetail", "frmEventDetails");
			frmEvent.txtBaseIndex.value = eventID;
			OpenHR.submitForm(frmEvent, "CalendarEvent");
			$("#CalendarEvent").dialog("open");
			$("#CalendarEvent").dialog("option", "position", ['center', 'center']); //Center popup in screen
		}
	}

	function todayClick() {

		var frmGetDataForm = OpenHR.getForm("reportworkframe", "frmCalendarGetData");
		var d = new Date();
		frmGetDataForm.txtMonth.value = d.getMonth() + 1;
		frmGetDataForm.txtYear.value = d.getFullYear();
		OpenHR.submitForm(frmGetDataForm);
		return true;
	}

	function toggleWeekends() {
		$(".scheduler_white_weekend").toggleClass("scheduler_white_weekendcell");

	}

	function toggleBankHolidays() {
		$(".scheduler_white_weekend").toggleClass("scheduler_white_weekendcell");
	}

	

</script>

<div style="float: left; width: 80%; height: 500px;overflow: auto">
	<DayPilot:DayPilotScheduler ID="DayPilotScheduler1" runat="server"
		HeaderFontSize="8pt" HeaderHeight="20"
		DataStartField="startdate"
		DataEndField="enddate"
		DataTextField="description"
		DataValueField="id"
		DataTypeField="eventtype"
		DataResourceField="resource"
		EventFontSize="11px"
		CellDuration="1440"
		NonBusinessBackColor="#FF0000"
		OnBeforeEventRender="DayPilotScheduler1_BeforeEventRender"
		EventClickHandling="JavaScript"
		EventClickJavaScript="eventCalendarClick({0},'{1}');"
		TimeFormat="Clock24Hours" 
		CssOnly="True"
		CssClassPrefix="scheduler_white"
		EventHeight="25" RowHeaderColumnWidths="200">
		<Resources>
		</Resources>
	</DayPilot:DayPilotScheduler>
</div>



	<div id="CalendarLegend" style="float:right;width:18%">
		
		<strong>Select Month :</strong>

		<input id="StartYearDemo" class="monthpicker" 
					
			<% 
			If objCalendar.StartOnCurrentMonth And Now < objCalendar.ReportEndDate Then
				Session("CALREP_Year") = Date.Now.Year.ToString.PadLeft(4, "0"c)
				Session("CALREP_Month") = Date.Now.Month.ToString.PadLeft(2, "0"c)
				objCalendar.StartOnCurrentMonth = False
			ElseIf Session("CALREP_Year") Is Nothing Then
				Session("CALREP_Year") = objCalendar.ReportStartDate.Year.ToString.PadLeft(4, "0"c)
				Session("CALREP_Month") = objCalendar.ReportStartDate.Month.ToString.PadLeft(2, "0"c)
			End If
				
			Dim dStartDate = DateTime.Parse(String.Format("{0}-{1}-01", Session("CALREP_Year"), Session("CALREP_Month")))

			Response.Write(String.Format("data-selected-year={0} ", dStartDate.Year))
			Response.Write(String.Format("data-start-year={0} ", objCalendar.ReportStartDate.Year))
			Response.Write(String.Format("data-final-year={0} ", objCalendar.ReportEndDate.Year))
			Response.Write(String.Format("value={0}/{1}", Session("CALREP_Month"), Session("CALREP_Year")))
			%>
			/>
		

		<input class="btn" type="button" id="cmdToday" name="cmdToday" value="Today" onclick="todayClick()" />

		<p></p>

		<strong>Legend :</strong>

		<%
			For Each objLegend In objCalendar.Legend
				If objLegend.Count > 0 Then

				%>
		<div class="scheduler_white_event_inner" style="position: relative; background: <% =objLegend.HexColor %>; width: 150px; height: 20px">
			<% =objLegend.LegendDescription%>
		</div>

				<%			
			
				End If
			Next
			
			objCalendar.IncludeBankHolidays = CBool(Session("CALREP_IncludeBankHolidays"))
			objCalendar.IncludeWorkingDaysOnly = CBool(Session("CALREP_IncludeWorkingDaysOnly"))
			objCalendar.ShowBankHolidays = CBool(Session("CALREP_ShowBankHolidays"))
			objCalendar.ShowCaptions = CBool(Session("CALREP_ShowCaptions"))
				
				%>
		
	<strong>Options :</strong>
		<%--<div class="scheduler_white_event_inner" style="position: relative;">--%>
			<div  style="position: relative;">
		<% 
			If objCalendar.ShowWeekends Then
				Response.Write("<input type='checkbox' id='chkShowWeekends' name='chkShowWeekends' onclick=""toggleWeekends();"" checked=""checked""/>Show Weekends" & vbNewLine)
			Else
				Response.Write("<input type='checkbox' id='chkShowWeekends' name='chkShowWeekends' onclick=""toggleWeekends();""/>Show Weekends" & vbNewLine)
			End If
						
		%>

		</div>

	</div>

<script runat="server">
	
	Private Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
			
		Dim objCalendar As HR.Intranet.Server.CalendarReport = CType(Session("objCalendar" & Session("CalRepUtilID")), HR.Intranet.Server.CalendarReport)
		Dim dStartDate As DateTime = New DateTime(objCalendar.ReportStartDate.Year, objCalendar.ReportStartDate.Month, 1)
		
		If Session("CALREP_Year") Is Nothing Then
			If objCalendar.StartOnCurrentMonth And Now < objCalendar.ReportEndDate Then
				dStartDate = New DateTime(Now.Year, Now.Month, 1)
			End If
		Else
			dStartDate = DateTime.Parse(String.Format("{0}-{1}-01", Session("CALREP_Year"), Session("CALREP_Month")))
		End If
		
		DayPilotScheduler1.StartDate = dStartDate
		DayPilotScheduler1.Days = DateTime.DaysInMonth(dStartDate.Year, dStartDate.Month)
		Calendar_BindDataset(DayPilotScheduler1)
		DataBind()
	
	End Sub
	
	Protected Sub DayPilotScheduler1_BeforeEventRender(sender As Object, e As Events.Scheduler.BeforeEventRenderEventArgs)
		Dim color As String = TryCast(e.DataItem("color"), String)
		If Not [String].IsNullOrEmpty(color) Then
			e.DurationBarColor = color
		End If
	End Sub

	Protected Sub DayPilotCalendar1_BeforeEventRender(sender As Object, e As Events.Calendar.BeforeEventRenderEventArgs)
		Dim color As String = TryCast(e.DataItem("color"), String)
		If Not [String].IsNullOrEmpty(color) Then
			e.DurationBarColor = color
		End If
	End Sub
	

	
</script>

<input type="hidden" name="txtFirstLoad" id="txtFirstLoad" value="<%=Session("CALREP_firstLoad").ToString()%>">

<form action="util_run_calendarreport_data_submit?CalRepUtilID=<%=Session("CalRepUtilID").ToString()%>" method="post" id="frmCalendarGetData" name="frmCalendarGetData">
		<input type="hidden" id="txtMonth" name="txtMonth" value="<%=Session("CALREP_Month").ToString()%>">
		<input type="hidden" id="txtYear" name="txtYear" value="<%=Session("CALREP_Year").ToString()%>">
		<input type="hidden" id="txtVisibleStartDate" name="txtVisibleStartDate">
		<input type="hidden" id="txtVisibleEndDate" name="txtVisibleEndDate">
		<input type="hidden" id="txtMode" name="txtMode">
		<input type="hidden" id="txtLoadCount" name="txtLoadCount" value="0">
		<input type="hidden" name="txtIncludeBankHolidays" id="txtIncludeBankHolidays" value="<%=Session("CALREP_IncludeBankHolidays").ToString()%>">
		<input type="hidden" name="txtIncludeWorkingDaysOnly" id="txtIncludeWorkingDaysOnly" value="<%=Session("CALREP_IncludeWorkingDaysOnly").ToString()%>">
		<input type="hidden" name="txtShowBankHolidays" id="txtShowBankHolidays" value="<%=Session("CALREP_ShowBankHolidays").ToString()%>">
		<input type="hidden" name="txtShowCaptions" id="txtShowCaptions" value="<%=Session("CALREP_ShowCaptions").ToString()%>">
		<input type="hidden" name="txtShowWeekends" id="txtShowWeekends" value="<%=Session("CALREP_ShowWeekends").ToString()%>">
		<input type="hidden" name="txtChangeOptions" id="txtChangeOptions"  value="<%=Session("CALREP_ChangeOptions").ToString()%>">
		<%=Html.AntiForgeryToken()%>
</form>

<form id="frmCalendarData" name="frmCalendarData" style="visibility: visible; display: block">
<%
	
	Dim sErrorDescription As String = ""

	Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
			
	Dim iLoop As Integer

	If Session("EmailGroupID") > 0 Then
		
		Try
			Dim rstEmailAddr = objDataAccess.GetDataTable("spASRIntGetEmailGroupAddresses", CommandType.StoredProcedure _
						, New SqlParameter("EmailGroupID", SqlDbType.Int) With {.Value = CleanNumeric(Session("EmailGroupID"))})

			iLoop = 1
			If Not rstEmailAddr Is Nothing Then
				Response.Write("<input id=txtEmailGroupAddr name=txtEmailGroupAddr value=""")

				For Each objRow In rstEmailAddr.Rows
					If iLoop > 1 Then
						Response.Write(";")
					End If

					Response.Write(Replace(objRow("Fixed").ToString(), """", "&quot;"))

				Next
				
				Response.Write(""">" & vbCrLf)
				
			End If

		Catch ex As Exception
			sErrorDescription = "Error getting the email addresses for group." & vbCrLf & FormatError(ex.Message)
		End Try
									
	Else
		Response.Write("<input type=hidden id=txtEmailGroupAddr name=txtEmailGroupAddr value=''>" & vbCrLf)
	End If

	If Not objCalendar Is Nothing Then
		sErrorDescription = objCalendar.ErrorString
	End If
	
	Response.Write("<input type='hidden' id=txtCalendarMode name=txtCalendarMode value=" & Session("CalRep_Mode") & ">" & vbCrLf)
	Response.Write("<input type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)

%>
</form>

<form action="util_run_calendarreport_download" method="post" id="frmExportData" name="frmExportData" target="submit-iframe">
	<input type="hidden" id="txtPreview" name="txtPreview" value="<%=objCalendar.OutputPreview%>">
	<input type="hidden" id="txtFormat" name="txtFormat" value="<%=objCalendar.OutputFormat%>">
	<input type="hidden" id="txtScreen" name="txtScreen" value="<%=objCalendar.OutputScreen%>">
	<input type="hidden" id="txtPrinter" name="txtPrinter" value="<%=objCalendar.OutputPrinter%>">
	<input type="hidden" id="txtPrinterName" name="txtPrinterName" value="<%=objCalendar.OutputPrinterName%>">
	<input type="hidden" id="txtSave" name="txtSave" value="<%=objCalendar.OutputSave%>">
	<input type="hidden" id="txtSaveExisting" name="txtSaveExisting" value="<%=objCalendar.OutputSaveExisting%>">
	<input type="hidden" id="txtEmail" name="txtEmail" value="<%=objCalendar.OutputEmail%>">
	<input type="hidden" id="txtEmailAddr" name="txtEmailAddr" value="<%=objCalendar.OutputEmailID%>">
	<input type="hidden" id="txtEmailAddrName" name="txtEmailAddrName" value="<%=Replace(objCalendar.OutputEmailGroupName, """", "&quot;")%>">
	<input type="hidden" id="txtEmailSubject" name="txtEmailSubject" value="<%=Replace(objCalendar.OutputEmailSubject, """", "&quot;")%>">
	<input type="hidden" id="txtEmailAttachAs" name="txtEmailAttachAs" value="<%=Replace(objCalendar.OutputEmailAttachAs, """", "&quot;")%>">
	<input type="hidden" id="txtEmailGroupID" name="txtEmailGroupID" value="<%=objCalendar.OutputEmailID%>">
	<input type="hidden" id="txtFileName" name="txtFileName" value="<%=objCalendar.OutputFilename%>">
	<input type="hidden" id="txtUtilType" name="txtUtilType" value="<%=session("utilType")%>">
	<input type="hidden" id="txtUtilID" name="txtUtilID" value="<%=Session("utilID")%>">
	<input type="hidden" id="download_token_value_id" name="download_token_value_id"/>
	<%=Html.AntiForgeryToken()%>
</form>

<form id="frmOriginalDefinition" style="visibility: hidden; display: none">
		<%
			Dim sReportName = objCalendar.Name
			Dim sErrMsg As String = ""
			Response.Write("	<input type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Replace(sReportName, """", "&quot;") & """>" & vbCrLf)
			Response.Write("	<input type='hidden' id=txtDefn_ErrMsg name=txtDefn_ErrMsg value=""" & sErrMsg & """>" & vbCrLf)
		%>
		<input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
		<input type="hidden" id="txtDateFormat" name="txtDateFormat" value="<%=session("LocaleDateFormat")%>">

		<input type="hidden" id="txtCancelPrint" name="txtCancelPrint">
		<input type="hidden" id="txtOptionsDone" name="txtOptionsDone">
		<input type="hidden" id="txtOptionsPortrait" name="txtOptionsPortrait">
		<input type="hidden" id="txtOptionsMarginLeft" name="txtOptionsMarginLeft">
		<input type="hidden" id="txtOptionsMarginRight" name="txtOptionsMarginRight">
		<input type="hidden" id="txtOptionsMarginTop" name="txtOptionsMarginTop">
		<input type="hidden" id="txtOptionsMarginBottom" name="txtOptionsMarginBottom">
		<input type="hidden" id="txtOptionsCopies" name="txtOptionsCopies">
		<input type="hidden" id="txtCalRep_UtilID" name="txtCalRep_UtilID" value='<%=Request("CalRepUtilID")%>'>
</form>

<div id="divEventDetail">
	<form id="frmEventDetails" name="frmEventDetails" action="util_run_calendarreport_breakdown" method="post" style="visibility: hidden; display: none">
		<input type="hidden" name="txtBreakdownCaption" id="txtBreakdownCaption">
		<input type="hidden" name="txtShowRegion" id="txtShowRegion">
		<input type="hidden" name="txtShowWorkingPattern" id="txtShowWorkingPattern">
		<input type="hidden" name="txtBaseIndex" id="txtBaseIndex">
		<input type="hidden" name="txtLabelIndex" id="txtLabelIndex">
		<%=Html.AntiForgeryToken()%>
	</form>
</div>

<% 

	Session("CALREP_Action") = ""
	Session("CalRep_Mode") = ""

	objCalendar = Nothing

%>