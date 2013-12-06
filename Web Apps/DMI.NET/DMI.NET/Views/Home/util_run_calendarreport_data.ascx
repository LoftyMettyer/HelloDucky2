<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="ADODB" %>
<%@ Import Namespace="System.Data" %>

<%@ Register Assembly="DayPilot" Namespace="DayPilot.Web.Ui" TagPrefix="DayPilot" %>

<link href="<%: Url.LatestContent("~/Themes/scheduler_white.css")%>" rel="stylesheet" type="text/css" />
<link href="<%: Url.LatestContent("~/Themes/calendar_white.css")%>" rel="stylesheet" type="text/css" />
<link href="<%: Url.LatestContent("~/Themes/layout.css")%>" rel="stylesheet" type="text/css" />
<link href="<%: Url.LatestContent("~/Content/MonthPicker.2.1.css")%>" rel="stylesheet" type="text/css" />
<script src="<%: Url.Content("~/scripts/jquery/jquery.maskedinput.js")%>" type="text/javascript"></script>
<script src="<%: Url.Content("~/scripts/jquery/MonthPicker.2.1.js")%>" type="text/javascript"></script>

<script type="text/javascript">

	$("#divReportButtons").css("visibility", "visible");
	$('#StartYearDemo').MonthPicker(
		{
			StartYear: $("#txtYear").val(),
			ShowIcon: false,
			UseInputMask: true,
			Speed: 10,
			OnAfterMenuClose: function () {

				var sMonthYear = $('#StartYearDemo').val();
				var frmGetDataForm = OpenHR.getForm("reportworkframe", "frmCalendarGetData");
				frmGetDataForm.txtMonth.value = sMonthYear.substring(0, 2);
				frmGetDataForm.txtYear.value = sMonthYear.substring(3, 7); 

				OpenHR.submitForm(frmGetDataForm);
			}
			
		});

	function eventCalendarClick(eventID) {

		var frmEvent = OpenHR.getForm("divEventDetail", "frmEventDetails");
		frmEvent.txtBaseIndex.value = eventID;
		OpenHR.submitForm(frmEvent, "CalendarEvent");
		$("#CalendarEvent").dialog("open");
		$("#CalendarEvent").dialog("option", "position", ['center', 'center']); //Center popup in screen
	}

	function todayClick() {

		var frmGetDataForm = OpenHR.getForm("reportworkframe", "frmCalendarGetData");
		var d = new Date();
		frmGetDataForm.txtMonth.value = d.getMonth() + 1;
		frmGetDataForm.txtYear.value = d.getFullYear();
		OpenHR.submitForm(frmGetDataForm);
		return true;
	}

	function ExportData(strMode) {

		var frmExport = OpenHR.getForm("reportworkframe", "frmExportData");
		frmExport.submit();

		return true;
	}

</script>

<div style="float: left; width: 80%">
	<DayPilot:DayPilotScheduler ID="DayPilotScheduler1" runat="server"
		HeaderFontSize="8pt" HeaderHeight="20"
		DataStartField="startdate"
		DataEndField="enddate"
		DataTextField="description"
		DataValueField="id"
		DataResourceField="resource"
		EventFontSize="11px"
		CellDuration="1440"
		NonBusinessBackColor="#FF0000"
		OnBeforeEventRender="DayPilotScheduler1_BeforeEventRender"
		EventClickHandling="JavaScript"
		EventClickJavaScript="eventCalendarClick({0});"
		TimeFormat="Clock24Hours" 
		CssOnly="True"
		CssClassPrefix="scheduler_white"
		EventHeight="25" RowHeaderColumnWidths="200">
		<Resources>
		</Resources>
	</DayPilot:DayPilotScheduler>
</div>


<%	
		Dim objCalendar As HR.Intranet.Server.CalendarReport
		objCalendar = Session("objCalendar" & Session("CalRepUtilID"))
	%>

	<div id="CalendarLegend" style="float:right;width:18%">
		
		<strong>Select Month :</strong>

		<input id="StartYearDemo" type="text"  value="	
		<%
		If objCalendar.StartOnCurrentMonth Then
			Session("CALREP_Year") = Date.Now.Year.ToString.PadLeft(4, "0"c)
			Session("CALREP_Month") = Date.Now.Month.ToString.PadLeft(2, "0"c)
			objCalendar.StartOnCurrentMonth = False
		ElseIf Session("CALREP_Year") Is Nothing Then
			Session("CALREP_Year") = objCalendar.ReportStartDate.Year.ToString.PadLeft(4, "0"c)
			Session("CALREP_Month") = objCalendar.ReportStartDate.Month.ToString.PadLeft(2, "0"c)
		End If
				
		Dim dStartDate = DateTime.Parse(String.Format("{0}-{1}-01", Session("CALREP_Year"), Session("CALREP_Month")))
	
		Response.Write(dStartDate.Month.ToString.PadLeft(2, "0"c) & "/" & dStartDate.Year)
		%>" />

		<input class="btn" type="button" id="cmdToday" name="cmdToday" value="Today" onclick="todayClick()" />

		<p></p>

		<strong>Legend :</strong>

		<%
			For Each objLegend In objCalendar.Legend
				If objLegend.Count > 0 Then

				%>
				<div class="scheduler_white_event_inner" style="position: relative; width: 50px; height: 20px">
					<div class="scheduler_white_event_bar_inner" style="background: <% =objLegend.HexColor %>; width: 100%">
						<% =objLegend.Text%>
					</div>
				</div>
				<%			
			
				End If
			Next
			
			objCalendar.IncludeBankHolidays = CBool(Session("CALREP_IncludeBankHolidays"))
			objCalendar.IncludeWorkingDaysOnly = CBool(Session("CALREP_IncludeWorkingDaysOnly"))
			objCalendar.ShowBankHolidays = CBool(Session("CALREP_ShowBankHolidays"))
			objCalendar.ShowCaptions = CBool(Session("CALREP_ShowCaptions"))
			objCalendar.ShowWeekends = CBool(Session("CALREP_ShowWeekends"))			
		%>

	</div>

<script runat="server">
	
	Private Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
			
		Dim objCalendar As HR.Intranet.Server.CalendarReport = CType(Session("objCalendar" & Session("CalRepUtilID")), HR.Intranet.Server.CalendarReport)
		Dim dStartDate As DateTime = objCalendar.ReportStartDate
	
		If Not Session("CALREP_Year") Is Nothing Then
			dStartDate = DateTime.Parse(String.Format("{0}-{1}-01", Session("CALREP_Year"), Session("CALREP_Month")))
		End If
		
		DayPilotScheduler1.StartDate = dStartDate
		DayPilotScheduler1.Days = DateTime.DaysInMonth(dStartDate.Year, dStartDate.Month)
		DayPilotScheduler1.DataSource = getData()
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
	
	Protected Function getData() As DataTable
		Dim dt As New DataTable()
		dt.Columns.Add("startdate", GetType(DateTime))
		dt.Columns.Add("enddate", GetType(DateTime))
		dt.Columns.Add("description", GetType(String))
		dt.Columns.Add("baseid", GetType(String))
		dt.Columns.Add("id", GetType(String))
		dt.Columns.Add("resource", GetType(String))
		dt.Columns.Add("color", GetType(String))
		
		Dim dr As DataRow

		Dim objCalendar As HR.Intranet.Server.CalendarReport
		objCalendar = Session("objCalendar" & Session("CalRepUtilID"))
			
		Dim sDescription As String
		Dim sPreviousDescription As String = ""
		Dim sEventDescription As String
		
		Dim dStart As Date
		Dim dEnd As Date
		
		If objCalendar.Events Is Nothing Then	'Report contains no records, return empty Data Table
			Return dt
		End If
		
		For Each objRow In objCalendar.Events.Rows

			sEventDescription = objRow("eventdescription1").ToString() & objRow("eventdescription2").ToString()
			sDescription = objCalendar.ConvertDescription(objRow("description1").ToString(), objRow("description2").ToString(), objRow("descriptionExpr").ToString())

			' Add to resource collection
			If Not sPreviousDescription = sDescription Then
				DayPilotScheduler1.Resources.Add(sDescription, objRow("baseid").ToString())
				sPreviousDescription = sDescription
			End If

			dr = dt.NewRow()
			dr("baseid") = objRow("baseid")
			dr("id") = objRow("id")
			
			If objRow("startsession") = "AM" Then
				dStart = CDate(objRow("startdate"))
			Else
				dStart = CDate(objRow("startdate")).AddHours(12)
			End If

			If objRow("endsession") = "AM" Then
				dEnd = CDate(objRow("enddate")).AddHours(12)
			Else
				dEnd = CDate(objRow("enddate")).AddDays(1)
			End If
			
			dr("startdate") = dStart
			dr("enddate") = dEnd
			dr("description") = sEventDescription
			dr("color") = objRow("eventcolor")
			dr("resource") = objRow("baseid")
			dt.Rows.Add(dr)
			
		Next
		
		Return dt
		
	End Function
	
</script>

<input type="hidden" name="txtFirstLoad" id="txtFirstLoad" value="<%=Session("CALREP_firstLoad").ToString()%>">

<form action="util_run_calendarreport_data_submit?CalRepUtilID=<%=Session("CalRepUtilID").ToString()%>" method="post" id="frmCalendarGetData" name="frmCalendarGetData">
		<input type="hidden" id="txtMonth" name="txtMonth" value="<%=Session("CALREP_Month").ToString()%>">
		<input type="hidden" id="txtYear" name="txtYear" value="<%=Session("CALREP_Year").ToString()%>">
		<input type="hidden" id="txtVisibleStartDate" name="txtVisibleStartDate">
		<input type="hidden" id="txtVisibleEndDate" name="txtVisibleEndDate">
		<input type="hidden" id="txtMode" name="txtMode">
		<input type="hidden" id="txtLoadCount" name="txtLoadCount" value="0">
		<input type="hidden" id="txtEmailGroupID" name="txtEmailGroupID" value="<%=Session("EmailGroupID").ToString()%>">		

		<input type="hidden" name="txtIncludeBankHolidays" id="txtIncludeBankHolidays" value="<%=Session("CALREP_IncludeBankHolidays").ToString()%>">
		<input type="hidden" name="txtIncludeWorkingDaysOnly" id="txtIncludeWorkingDaysOnly" value="<%=Session("CALREP_IncludeWorkingDaysOnly").ToString()%>">
		<input type="hidden" name="txtShowBankHolidays" id="txtShowBankHolidays" value="<%=Session("CALREP_ShowBankHolidays").ToString()%>">
		<input type="hidden" name="txtShowCaptions" id="txtShowCaptions" value="<%=Session("CALREP_ShowCaptions").ToString()%>">
		<input type="hidden" name="txtShowWeekends" id="txtShowWeekends" value="<%=Session("CALREP_ShowWeekends").ToString()%>">
		<input type="hidden" name="txtChangeOptions" id="txtChangeOptions"  value="<%=Session("CALREP_ChangeOptions").ToString()%>">
</form>

<form id="frmCalendarData" name="frmCalendarData" style="visibility: visible; display: block">
<%
	on error resume next
	
	Dim sErrorDescription As String = ""

	Dim cmdEmailGroup As Command
	Dim prmEmailGroupID As ADODB.Parameter
	Dim rstEmails As Recordset
	Dim iLoop As Integer

	If Session("EmailGroupID") > 0 Then
		cmdEmailGroup = New Command
		cmdEmailGroup.CommandText = "spASRIntGetEmailGroupAddresses"
		cmdEmailGroup.CommandType = CommandTypeEnum.adCmdStoredProc
		cmdEmailGroup.ActiveConnection = Session("databaseConnection")

		prmEmailGroupID = cmdEmailGroup.CreateParameter("EmailGroupID", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput)
		cmdEmailGroup.Parameters.Append(prmEmailGroupID)
		prmEmailGroupID.Value = CleanNumeric(Session("EmailGroupID"))

		Err.Clear()
		rstEmails = cmdEmailGroup.Execute

		If (Err.Number <> 0) Then
			sErrorDescription = "Error getting the email addresses for group." & vbCrLf & FormatError(Err.Description)
		End If

		If Len(sErrorDescription) = 0 Then
			iLoop = 1
			Response.Write("<INPUT id=txtEmailGroupAddr name=txtEmailGroupAddr value=""")
			Do While Not rstEmails.EOF
				If iLoop > 1 Then
					Response.Write(";")
				End If
				Response.Write(Replace(rstEmails.Fields("Fixed").Value, """", "&quot;"))
				rstEmails.MoveNext()
				iLoop = iLoop + 1
			Loop
			Response.Write(""">" & vbCrLf)

			' Release the ADO recordset object.
			rstEmails.Close()
		End If
					
		rstEmails = Nothing
		cmdEmailGroup = Nothing
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

<form action="util_run_calendarreport_download" method="post" id="frmExportData" name="frmExportData">
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
	<input type="hidden" id="txtEmailGroupID" name="txtEmailGroupID" value="<%=Session("EmailGroupID").ToString()%>">
	<input type="hidden" id="txtFileName" name="txtFileName" value="<%=objCalendar.OutputFilename%>">
	<input type="hidden" id="txtUtilType" name="txtUtilType" value="<%=session("utilType")%>">

</form>

<form id="frmOriginalDefinition" style="visibility: hidden; display: none">
		<%
				Dim sErrMsg As String = ""
			Response.Write("	<input type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Replace(Session("utilname"), """", "&quot;") & """>" & vbCrLf)
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
	</form>
</div>

<% 

	Session("CALREP_Action") = ""
	Session("CalRep_Mode") = ""

	objCalendar = Nothing

%>