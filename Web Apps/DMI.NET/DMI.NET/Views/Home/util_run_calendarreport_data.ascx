<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<object
    classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
    id="Microsoft_Licensed_Class_Manager_1_0">
    <param name="LPKPath" value="lpks/main.lpk">
</object>

<script type="text/javascript">

    function util_run_calendarreport_data_window_onload() {

        if (txtFirstLoad.value == 1) {
            loadAddRecords('data');
            return;
        }

        if (frmCalendarData.txtCalendarMode.value == "LOADCALENDARREPORTDATA") {
            fillCalBoxes();
        }
        else if (frmCalendarData.txtCalendarMode.value == "OUTPUTREPORT") {
            setGridFont(frmCalendarData.grdCalendarOutput);
            setGridFont(frmCalendarData.ssHiddenGrid);

            outputReport();
        }
    }


    function ExportData(strMode)
    {
        var frmGetDataForm = document.forms("frmCalendarGetData");
			
        frmGetDataForm.txtMode.value = "OUTPUTREPORT";

        refreshData();
    }
	
    function refreshData() {
        var frmGetData = OpenHR.getForm("dataframe", "frmCalendarGetData");
        OpenHR.submitForm(frmGetData);
    }
	
</script>

<input type="hidden" name="txtFirstLoad" id="txtFirstLoad" value="<%=Session("firstLoad").ToString()%>">

<form action="util_run_calendarreport_data_submit?CalRepUtilID=<%=Session("CalRepUtilID").ToString()%>&firstLoad=0" method="post" id="frmCalendarGetData" name="frmCalendarGetData">
    <input type="hidden" id="txtDaysInMonth" name="txtDaysInMonth">
    <input type="hidden" id="txtMonth" name="txtMonth">
    <input type="hidden" id="txtYear" name="txtYear">
    <input type="hidden" id="txtVisibleStartDate" name="txtVisibleStartDate">
    <input type="hidden" id="txtVisibleEndDate" name="txtVisibleEndDate">
    <input type="hidden" id="txtMode" name="txtMode">
    <input type="hidden" id="txtEmailGroupID" name="txtEmailGroupID">
</form>

<form id="frmCalendarData" name="frmCalendarData" style="visibility: visible; display: block">
<%
	on error resume next
	
    Dim sErrorDescription As String
    
	sErrorDescription = ""
	
    Dim fok As Boolean
    Dim fNotCancelled As Boolean
	
    Dim objCalendar As Object
    Dim rsEvents As Object
    Dim lngCurrentBaseID As Long
    Dim intBaseRecordIndex As Integer
	dim dtEventStartDate
	dim dtEventEndDate
	dim strEventStartSession
	dim strEventEndSession
    Dim strEventToolTip
	dim INPUT_VALUE 
	dim EVENT_DETAIL
	dim strKeyCode
	dim intEventCounter 
	dim intSessionStart
	dim intSessionEnd
	dim strEventID
	
	dim strEventDesc1Value_BD
	dim strEventDesc1ColumnName_BD
	dim strEventDesc2Value_BD
	dim strEventDesc2ColumnName_BD
	dim strBaseDescription_BD
	
    Dim lngMonth As Long
    Dim lngYear As Long
	dim dtMonth 
	dim dtVisibleStartDate
	dim dtVisibleEndDate
    Dim mintDaysInMonth As Integer

    Dim arrayDefinition
    Dim arrayColumnsDefinition
    Dim arrayDataDefinition

    Dim arrayStyles
    Dim arrayMerges
    
	fok = true
	fNotCancelled = true

	if Session("CalRep_Mode") = "EMAILGROUP" then
		Session("CalRep_Mode") = "OUTPUTREPORT"
	end if
	
	if Session("CalRep_Mode") = "LOADCALENDARREPORTDATA" then		

        objCalendar = Session("objCalendar" & Session("CalRepUtilID"))
		
        rsEvents = objCalendar.EventsRecordset
		
		intEventCounter = 0 
		
'********************************************************************
		lngMonth = CLng(Session("CALREP_Month"))
		lngYear = CLng(Session("CALREP_Year"))
		
		dtMonth = DateAdd("yyyy", CDbl(lngYear - Year(objCalendar.ReportStartDate_Calendar)), objCalendar.ReportStartDate_Calendar)
		dtMonth = DateAdd("m", CDbl(lngMonth - Month(objCalendar.ReportStartDate_Calendar)), dtMonth)
  
		mintDaysInMonth = Cint(Session("CALREP_DaysInMonth"))

		'Define the current visible Start and End Dates.
		dtVisibleEndDate = DateAdd("d", CDbl(mintDaysInMonth - Day(dtMonth)), dtMonth)
		dtVisibleStartDate = DateAdd("d", CDbl(-(mintDaysInMonth - 1)), dtVisibleEndDate)
  
'********************************************************************
		
		With rsEvents
			If Not (.BOF And .EOF) Then
			    
			  .MoveFirst
			  Do While Not .EOF
					
					lngCurrentBaseID = rsEvents.Fields(objCalendar.BaseIDColumn)
					strEventID = rsEvents.Fields(objCalendar.EventIDColumn)
					
					intBaseRecordIndex = objCalendar.BaseIndex_Get(CStr(lngCurrentBaseID))

					strBaseDescription_BD = objCalendar.ConvertDescription(.Fields("Description1").Value, .Fields("Description2").Value, .Fields("DescriptionExpr").Value)
					
                    If IsDBNull(.Fields("Legend").value) Then
                        strKeyCode = ""
                    Else
                        strKeyCode = Left(.Fields("Legend").value, 2)
                    End If
                    If IsDBNull(.Fields("EventDescription1Column").Value) Then
                        strEventDesc1ColumnName_BD = ""
                    Else
                        strEventDesc1ColumnName_BD = CStr(.Fields("EventDescription1Column").Value)
                    End If
					strEventDesc1Value_BD = objCalendar.ConvertEventDescription(.Fields("EventDescription1ColumnID").Value,.Fields("EventDescription1").Value)
					
                    If IsDBNull(.Fields("EventDescription2Column").Value) Then
                        strEventDesc2ColumnName_BD = ""
                    Else
                        strEventDesc2ColumnName_BD = CStr(.Fields("EventDescription2Column").Value)
                    End If
					strEventDesc2Value_BD = objCalendar.ConvertEventDescription(.Fields("EventDescription2ColumnID").Value,.Fields("EventDescription2").Value)
      
					EVENT_DETAIL = vbNullString 
					EVENT_DETAIL = EVENT_DETAIL & .Fields("Name") & vbTab
					EVENT_DETAIL = EVENT_DETAIL & convertSQLDateToLocale(.Fields("StartDate")) & vbTab
					EVENT_DETAIL = EVENT_DETAIL & UCase(.Fields("StartSession")) & vbTab
					EVENT_DETAIL = EVENT_DETAIL & convertSQLDateToLocale(.Fields("EndDate")) & vbTab
					EVENT_DETAIL = EVENT_DETAIL & UCase(.Fields("EndSession")) & vbTab
					EVENT_DETAIL = EVENT_DETAIL & FormatNumber(.Fields("Duration"),1,true) & vbTab
					EVENT_DETAIL = EVENT_DETAIL & strKeyCode & vbTab
					EVENT_DETAIL = EVENT_DETAIL & strEventDesc1ColumnName_BD & vbTab
					EVENT_DETAIL = EVENT_DETAIL & strEventDesc1Value_BD & vbTab
					EVENT_DETAIL = EVENT_DETAIL & strEventDesc2ColumnName_BD & vbTab
					EVENT_DETAIL = EVENT_DETAIL & strEventDesc2Value_BD & vbTab
					EVENT_DETAIL = EVENT_DETAIL & strBaseDescription_BD 
					
					INPUT_VALUE = vbNullString 
					
					'****************************************************************************
                    dtEventStartDate = .Fields("StartDate").Value
    
                    If IsDBNull(.Fields("EndDate").Value) Then
                        dtEventEndDate = dtEventStartDate
                    Else
                        dtEventEndDate = .Fields("EndDate").Value
                    End If
  
                    If IsDBNull(.Fields("StartSession").Value) And IsDBNull(.Fields("EndSession").Value) Then
                        strEventStartSession = "AM"
                        strEventEndSession = "PM"
                    ElseIf IsDBNull(.Fields("EndSession").Value) Then
                        strEventEndSession = strEventStartSession
                    Else
                        strEventStartSession = UCase(.Fields("StartSession").Value)
                        strEventEndSession = UCase(.Fields("EndSession").Value)
                    End If

					strEventToolTip = objCalendar.EventToolTipText(CDate(dtEventStartDate), CStr(strEventStartSession), Cdate(dtEventEndDate), CStr(strEventEndSession))
    
					'Force the Start & End Dates to be between the Report Start and End dates.
					If dtEventStartDate < objCalendar.ReportStartDate Then
					  dtEventStartDate = objCalendar.ReportStartDate
					End If
      
					If dtEventEndDate > objCalendar.ReportEndDate Then
					  dtEventEndDate = objCalendar.ReportEndDate
					End If

					'****************************************************************************
      
					' If the event start date is after the event end date, ignore the record
					If (dtEventStartDate > dtEventEndDate) Then
      
					' if the event is totally before the currently viewed timespan then do nothing
					ElseIf (dtEventStartDate < dtVisibleStartDate) And (dtEventEndDate < dtVisibleStartDate) Then
      
					' if the event is totally after the currently viewed timespan then do nothing
					ElseIf (dtEventStartDate > dtVisibleEndDate) And (dtEventEndDate > dtVisibleEndDate) Then
      
					' if the event starts before currently viewed timespan, but ends in the timspan then
					ElseIf (dtEventStartDate < dtVisibleStartDate) And (dtEventEndDate <= dtVisibleEndDate) Then
					  
					  dtEventStartDate = dtVisibleStartDate
					  strEventStartSession = "AM"
					  
						if strEventStartSession = "AM" then
							intSessionStart = 0
						else
							intSessionStart = 1
						end if
						
						if strEventEndSession = "AM" then
							intSessionEnd = 0
						else
							intSessionEnd = 1
						end if
						
						intEventCounter = intEventCounter + 1
						
						INPUT_VALUE = intBaseRecordIndex & "***" & strEventID & "***" & convertSQLDateToLocale(dtEventStartDate) & "***" & convertSQLDateToLocale(dtEventEndDate) & "***" & intSessionStart & "***" & intSessionEnd & "***" & strEventToolTip & "***" & strKeyCode
					  
                        Response.Write("<INPUT type=hidden name=Event_" & intEventCounter & " ID=Event_" & intEventCounter & " VALUE=""" & INPUT_VALUE & """>" & vbCrLf)
					  
                        Response.Write("<INPUT type=hidden name=EventDetail_" & intEventCounter & " ID=EventDetail_" & intEventCounter & " VALUE=""" & Server.HtmlEncode(EVENT_DETAIL) & """>" & vbCrLf)
					  
					' if the event starts in the currently viewed timespan, but ends after it then
					ElseIf (dtEventStartDate >= dtVisibleStartDate) And (dtEventEndDate > dtVisibleEndDate) Then
					  
					  dtEventEndDate = dtVisibleEndDate
					  strEventEndSession = "PM"

						if strEventStartSession = "AM" then
							intSessionStart = 0
						else
							intSessionStart = 1
						end if
						
						if strEventEndSession = "AM" then
							intSessionEnd = 0
						else
							intSessionEnd = 1
						end if
						
						intEventCounter = intEventCounter + 1
						
						INPUT_VALUE = intBaseRecordIndex & "***" & strEventID & "***" & convertSQLDateToLocale(dtEventStartDate) & "***" & convertSQLDateToLocale(dtEventEndDate) & "***" & intSessionStart & "***" & intSessionEnd & "***" & strEventToolTip & "***" & strKeyCode
					  
                        Response.Write("<INPUT type=hidden name=Event_" & intEventCounter & " ID=Event_" & intEventCounter & " VALUE=""" & INPUT_VALUE & """>" & vbCrLf)

                        Response.Write("<INPUT type=hidden name=EventDetail_" & intEventCounter & " ID=EventDetail_" & intEventCounter & " VALUE=""" & Server.HtmlEncode(EVENT_DETAIL) & """>" & vbCrLf)
  
					' if the event is enclosed within viewed timespan, and months are equal then
					ElseIf (dtEventStartDate >= dtVisibleStartDate) And (dtEventEndDate <= dtVisibleEndDate) And (Month(dtEventStartDate) = Month(dtEventEndDate)) Then
  
						if strEventStartSession = "AM" then
							intSessionStart = 0
						else
							intSessionStart = 1
						end if
						
						if strEventEndSession = "AM" then
							intSessionEnd = 0
						else
							intSessionEnd = 1
						end if
						
						intEventCounter = intEventCounter + 1
						
						INPUT_VALUE = intBaseRecordIndex & "***" & strEventID & "***" & convertSQLDateToLocale(dtEventStartDate) & "***" & convertSQLDateToLocale(dtEventEndDate) & "***" & intSessionStart & "***" & intSessionEnd & "***" & strEventToolTip & "***" & strKeyCode
					  
                        Response.Write("<INPUT type=hidden name=Event_" & intEventCounter & " ID=Event_" & intEventCounter & " VALUE=""" & INPUT_VALUE & """>" & vbCrLf)

                        Response.Write("<INPUT type=hidden name=EventDetail_" & intEventCounter & " ID=EventDetail_" & intEventCounter & " VALUE=""" & Server.HtmlEncode(EVENT_DETAIL) & """>" & vbCrLf)
     
					' if the event starts before the the viewed timespan and ends after the viewed timespan then
					ElseIf (dtEventStartDate < dtVisibleStartDate) And (dtEventEndDate > dtVisibleEndDate) Then
					  
					  dtEventStartDate = dtVisibleStartDate
					  strEventStartSession = "AM"
					  
					  dtEventEndDate = dtVisibleEndDate
					  strEventEndSession = "PM"
					          
						if strEventStartSession = "AM" then
							intSessionStart = 0
						else
							intSessionStart = 1
						end if
						
						if strEventEndSession = "AM" then
							intSessionEnd = 0
						else
							intSessionEnd = 1
						end if
						
						intEventCounter = intEventCounter + 1
						
						INPUT_VALUE = intBaseRecordIndex & "***" & strEventID & "***" & convertSQLDateToLocale(dtEventStartDate) & "***" & convertSQLDateToLocale(dtEventEndDate) & "***" & intSessionStart & "***" & intSessionEnd & "***" & strEventToolTip & "***" & strKeyCode
					  
                        Response.Write("<INPUT type=hidden name=Event_" & intEventCounter & " ID=Event_" & intEventCounter & " VALUE=""" & INPUT_VALUE & """>" & vbCrLf)

                        Response.Write("<INPUT type=hidden name=EventDetail_" & intEventCounter & " ID=EventDetail_" & intEventCounter & " VALUE=""" & Server.HtmlEncode(EVENT_DETAIL) & """>" & vbCrLf)
  
					End If
					
					.MoveNext
			  loop 
			  
			end if
			
		end with
		
	elseif Session("CalRep_Mode") = "OUTPUTREPORT" then		
        objCalendar = Session("objCalendar" & Session("CalRepUtilID"))
		
		if fok then 
			fok = objCalendar.OutputGridDefinition 
			fNotCancelled = Response.IsClientConnected 
			if fok then fok = fNotCancelled
		end if

		if fok then 
			fok = objCalendar.OutputGridColumns 
			fNotCancelled = Response.IsClientConnected 
			if fok then fok = fNotCancelled
		end if

		if fok then 
			fok = objCalendar.OutputReport(true) 
			fNotCancelled = Response.IsClientConnected 
			if fok then fok = fNotCancelled
		end if

		if fok then

		  arrayDefinition = objCalendar.OutputArray_Definition 
			arrayColumnsDefinition = objCalendar.OutputArray_Columns 
			arrayDataDefinition = objCalendar.OutputArray_Data 
		end if	
		
		if fok then
%>
	<TABLE WIDTH=100% HEIGHT=500>
		<TR>
			<TD>
				<OBJECT classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"
					id=grdCalendarOutput 
					name=grdCalendarOutput 
					codebase="cabs/COAInt_Grid.cab#version=3,1,3,6"
					style="BACKGROUND-COLOR: threedface; visibility: visible; display: block; HEIGHT: 0px; WIDTH: 0px"
					height="0"
					width="0" VIEWASTEXT>
<%
    For iCount = 1 To UBound(arrayDefinition)
        Response.Write(arrayDefinition(iCount))
    Next

    For iCount = 1 To UBound(arrayColumnsDefinition)
        Response.Write(arrayColumnsDefinition(iCount))
    Next
			
    For iCount = 1 To UBound(arrayDataDefinition)
        Response.Write(arrayDataDefinition(iCount))
    Next
%>
				</OBJECT>
			</TD>
		</TR>
	</TABLE>
<%
			if fok then
        arrayStyles = objCalendar.OutputArray_Styles
				arrayMerges = objCalendar.OutputArray_Merges
			end if	

			'************************* START OF HIDDEN GRID ******************************
%>
	<OBJECT classid="clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"   
		codebase="cabs/COAInt_Grid.cab#version=3,1,3,6" 
		id=ssHiddenGrid name=ssHiddenGrid 
		style="BACKGROUND-COLOR: threedface; visibility: visible; display: block; HEIGHT: 0px; WIDTH: 0px" 
		WIDTH=0 HEIGHT=0>
		
		<PARAM NAME="ScrollBars" VALUE="4">
		<PARAM NAME="_Version" VALUE="196617">
		<PARAM NAME="DataMode" VALUE="2">
		<PARAM NAME="Cols" VALUE="1">
		<PARAM NAME="Rows" VALUE="0">
		<PARAM NAME="BorderStyle" VALUE="1">
		<PARAM NAME="RecordSelectors" VALUE="0">
		<PARAM NAME="GroupHeaders" VALUE="0">
		<PARAM NAME="ColumnHeaders" VALUE="0">
		<PARAM NAME="GroupHeadLines" VALUE="0">
		<PARAM NAME="HeadLines" VALUE="0">
		<PARAM NAME="FieldDelimiter" VALUE="(None)">
		<PARAM NAME="FieldSeparator" VALUE="(Tab)">
		<PARAM NAME="Row.Count" VALUE="0">
		<PARAM NAME="Col.Count" VALUE="1">
		<PARAM NAME="stylesets.count" VALUE="0">
		<PARAM NAME="TagVariant" VALUE="EMPTY">
		<PARAM NAME="UseGroups" VALUE="0">
		<PARAM NAME="HeadFont3D" VALUE="0">
		<PARAM NAME="Font3D" VALUE="0">
		<PARAM NAME="DividerType" VALUE="0">
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
		<PARAM NAME="SelectTypeRow" VALUE="2">
		<PARAM NAME="SelectByCell" VALUE="0">
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
		<PARAM NAME="RowHeight" VALUE="238">
		<PARAM NAME="ExtraHeight" VALUE="0">
		<PARAM NAME="ActiveRowStyleSet" VALUE="">
		<PARAM NAME="CaptionAlignment" VALUE="2">
		<PARAM NAME="SplitterPos" VALUE="0">
		<PARAM NAME="SplitterVisible" VALUE="0">
		<PARAM NAME="Columns.Count" VALUE="1">
		<PARAM NAME="Columns(0).Width" VALUE="1000">
		<PARAM NAME="Columns(0).Visible" VALUE="0">
		<PARAM NAME="Columns(0).Columns.Count" VALUE="1">
		<PARAM NAME="Columns(0).Caption" VALUE="PageBreak">
		<PARAM NAME="Columns(0).Name" VALUE="PageBreak">
		<PARAM NAME="Columns(0).Alignment" VALUE="0">
		<PARAM NAME="Columns(0).CaptionAlignment" VALUE="2">
		<PARAM NAME="Columns(0).Bound" VALUE="0">
		<PARAM NAME="Columns(0).AllowSizing" VALUE="1">
		<PARAM NAME="Columns(0).DataField" VALUE="Column 0">
		<PARAM NAME="Columns(0).DataType" VALUE="8">
		<PARAM NAME="Columns(0).Level" VALUE="0">
		<PARAM NAME="Columns(0).NumberFormat" VALUE="">
		<PARAM NAME="Columns(0).Case" VALUE="0">
		<PARAM NAME="Columns(0).FieldLen" VALUE="4096">
		<PARAM NAME="Columns(0).VertScrollBar" VALUE="0">
		<PARAM NAME="Columns(0).Locked" VALUE="0">
		<PARAM NAME="Columns(0).Style" VALUE="0">
		<PARAM NAME="Columns(0).ButtonsAlways" VALUE="0">
		<PARAM NAME="Columns(0).RowCount" VALUE="0">
		<PARAM NAME="Columns(0).ColCount" VALUE="1">
		<PARAM NAME="Columns(0).HasHeadForeColor" VALUE="0">
		<PARAM NAME="Columns(0).HasHeadBackColor" VALUE="0">
		<PARAM NAME="Columns(0).HasForeColor" VALUE="0">
		<PARAM NAME="Columns(0).HasBackColor" VALUE="0">
		<PARAM NAME="Columns(0).HeadForeColor" VALUE="0">
		<PARAM NAME="Columns(0).HeadBackColor" VALUE="0">
		<PARAM NAME="Columns(0).ForeColor" VALUE="0">
		<PARAM NAME="Columns(0).BackColor" VALUE="0">
		<PARAM NAME="Columns(0).HeadStyleSet" VALUE="">
		<PARAM NAME="Columns(0).StyleSet" VALUE="">
		<PARAM NAME="Columns(0).Nullable" VALUE="1">
		<PARAM NAME="Columns(0).Mask" VALUE="">
		<PARAM NAME="Columns(0).PromptInclude" VALUE="0">
		<PARAM NAME="Columns(0).ClipMode" VALUE="0">
		<PARAM NAME="Columns(0).PromptChar" VALUE="95">
		<PARAM NAME="UseDefaults" VALUE="-1">
		<PARAM NAME="TabNavigation" VALUE="1">
		<PARAM NAME="BatchUpdate" VALUE="0">
		<PARAM NAME="_ExtentX" VALUE="0">
		<PARAM NAME="_ExtentY" VALUE="0">
		<PARAM NAME="_StockProps" VALUE="79">
		<PARAM NAME="Caption" VALUE="">
		<PARAM NAME="ForeColor" VALUE="0">
		<PARAM NAME="BackColor" VALUE="16777215">
		<PARAM NAME="Enabled" VALUE="-1">
		<PARAM NAME="DataMember" VALUE="">
	</OBJECT>
<%		
			'***************************** END OF HIDDEN GRID **************************************				
				
    Response.Write("<INPUT type='hidden' id=txtCalendarPageCount name=txtCalendarPageCount value=" & UBound(arrayMerges) & ">" & vbCrLf)
		end if
	end if
	
Dim cmdEmailGroup As Object
Dim prmEmailGroupID As Object
Dim rstEmails As Object
Dim iLoop As Integer

	if Session("EmailGroupID") > 0 then
cmdEmailGroup = CreateObject("ADODB.Command")
		cmdEmailGroup.CommandText = "spASRIntGetEmailGroupAddresses"
		cmdEmailGroup.CommandType = 4 ' Stored procedure
cmdEmailGroup.ActiveConnection = Session("databaseConnection")

prmEmailGroupID = cmdEmailGroup.CreateParameter("EmailGroupID", 3, 1) ' 3=integer, 1=input
cmdEmailGroup.Parameters.Append(prmEmailGroupID)
		prmEmailGroupID.value = cleanNumeric(Session("EmailGroupID"))

Err.Clear()
rstEmails = cmdEmailGroup.Execute

If (Err.Number <> 0) Then
    sErrorDescription = "Error getting the email addresses for group." & vbCrLf & formatError(Err.Description)
End If

		if len(sErrorDescription) = 0 then
			iLoop = 1
    Response.Write("<INPUT id=txtEmailGroupAddr name=txtEmailGroupAddr value=""")
			do while not rstEmails.EOF
				if iLoop > 1 then
            Response.Write(";")
				end if
        Response.Write(Replace(rstEmails.Fields("Fixed").Value, """", "&quot;"))
				rstEmails.MoveNext
				iLoop = iLoop + 1
			loop
    Response.Write(""">" & vbCrLf)

			' Release the ADO recordset object.
			rstEmails.close
		end if
					
rstEmails = Nothing
cmdEmailGroup = Nothing
Else
Response.Write("<INPUT type=hidden id=txtEmailGroupAddr name=txtEmailGroupAddr value=''>" & vbCrLf)
End If
		
sErrorDescription = objCalendar.ErrorString
	
Response.Write("<INPUT type='hidden' id=txtCalendarMode name=txtCalendarMode value=" & Session("CalRep_Mode") & ">" & vbCrLf)
Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
%>
</form>

<% 
if Session("CalRep_Mode") = "OUTPUTREPORT" then 

	dim iPage
	dim iStyle
	dim iMerge
	dim arrayPageStyles
	dim arrayPageMerges
	
        For iPage = 0 To UBound(arrayMerges)
            arrayPageMerges = arrayMerges(iPage)
            Response.Write("<FORM id=frmCalendarMerge_" & iPage & " name=frmCalendarMerge_" & iPage & ">" & vbCrLf)
            For iMerge = 0 To UBound(arrayPageMerges)
                INPUT_VALUE = arrayPageMerges(iMerge)
                Response.Write("	<INPUT type=hidden name=Merge_" & iPage & "_" & iMerge & " ID=Merge_" & iPage & "_" & iMerge & " VALUE=""" & INPUT_VALUE & """>" & vbCrLf)
            Next
            Response.Write("</FORM>" & vbCrLf)
        Next

        For iPage = 0 To UBound(arrayStyles)
            arrayPageStyles = arrayStyles(iPage)
            Response.Write("<FORM id=frmCalendarStyle_" & iPage & " name=frmCalendarStyle_" & iPage & ">" & vbCrLf)
            For iStyle = 0 To UBound(arrayPageStyles)
                INPUT_VALUE = arrayPageStyles(iStyle)
                Response.Write("	<INPUT type=hidden name=Style_" & iPage & "_" & iStyle & " ID=Style_" & iPage & "_" & iStyle & " VALUE=""" & INPUT_VALUE & """>" & vbCrLf)
            Next
            Response.Write("</FORM>" & vbCrLf)
        Next

end if 

if fok then
	objCalendar.OutputArray_Clear
end if

Session("CALREP_Action") = ""
Session("CalRep_Mode") = ""

    objCalendar = Nothing

%>

<form id="frmOriginalDefinition" style="visibility: hidden; display: none">
    <%
        Dim sErrMsg As String = ""
        Response.Write("	<INPUT type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Replace(Session("utilname"), """", "&quot;") & """>" & vbCrLf)
        Response.Write("	<INPUT type='hidden' id=txtDefn_ErrMsg name=txtDefn_ErrMsg value=""" & sErrMsg & """>" & vbCrLf)
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

<script runat="server" language="vb">

    Function convertSQLDateToLocale(psDate)
        Dim sLocaleFormat As String
        Dim iIndex As Integer
	
        If Len(psDate) > 0 Then
            sLocaleFormat = Session("LocaleDateFormat")
		
            iIndex = InStr(sLocaleFormat, "dd")
            If iIndex > 0 Then
                If Day(psDate) < 10 Then
                    sLocaleFormat = Left(sLocaleFormat, iIndex - 1) & _
                        "0" & Day(psDate) & Mid(sLocaleFormat, iIndex + 2)
                Else
                    sLocaleFormat = Left(sLocaleFormat, iIndex - 1) & _
                        Day(psDate) & Mid(sLocaleFormat, iIndex + 2)
                End If
            End If
		
            iIndex = InStr(sLocaleFormat, "mm")
            If iIndex > 0 Then
                If Month(psDate) < 10 Then
                    sLocaleFormat = Left(sLocaleFormat, iIndex - 1) & _
                        "0" & Month(psDate) & Mid(sLocaleFormat, iIndex + 2)
                Else
                    sLocaleFormat = Left(sLocaleFormat, iIndex - 1) & _
                        Month(psDate) & Mid(sLocaleFormat, iIndex + 2)
                End If
            End If
		
            iIndex = InStr(sLocaleFormat, "yyyy")
            If iIndex > 0 Then
                sLocaleFormat = Left(sLocaleFormat, iIndex - 1) & _
                    Year(psDate) & Mid(sLocaleFormat, iIndex + 4)
            End If

            convertSQLDateToLocale = sLocaleFormat
        Else
            convertSQLDateToLocale = ""
        End If
    End Function

    Function convertSQLDateToCalendar(psDate)
        Dim sCalendarFormat As String
        Dim iIndex As Integer
	
        If Len(psDate) > 0 Then
            'sLocaleFormat = session("LocaleDateFormat")
            sCalendarFormat = "dd/mm/yyyy"
		
            iIndex = InStr(sCalendarFormat, "dd")
            If iIndex > 0 Then
                If Day(psDate) < 10 Then
                    sCalendarFormat = Left(sCalendarFormat, iIndex - 1) & _
                        "0" & Day(psDate) & Mid(sCalendarFormat, iIndex + 2)
                Else
                    sCalendarFormat = Left(sCalendarFormat, iIndex - 1) & _
                        Day(psDate) & Mid(sCalendarFormat, iIndex + 2)
                End If
            End If
		
            iIndex = InStr(sCalendarFormat, "mm")
            If iIndex > 0 Then
                If Month(psDate) < 10 Then
                    sCalendarFormat = Left(sCalendarFormat, iIndex - 1) & _
                        "0" & Month(psDate) & Mid(sCalendarFormat, iIndex + 2)
                Else
                    sCalendarFormat = Left(sCalendarFormat, iIndex - 1) & _
                        Month(psDate) & Mid(sCalendarFormat, iIndex + 2)
                End If
            End If
		
            iIndex = InStr(sCalendarFormat, "yyyy")
            If iIndex > 0 Then
                sCalendarFormat = Left(sCalendarFormat, iIndex - 1) & _
                    Year(psDate) & Mid(sCalendarFormat, iIndex + 4)
            End If

            convertSQLDateToCalendar = sCalendarFormat
        Else
            convertSQLDateToCalendar = ""
        End If
    End Function

    Function formatError(psErrMsg As String)
        Dim iStart As Integer
        Dim iFound As Integer
  
        iFound = 0
        Do
            iStart = iFound
            iFound = InStr(iStart + 1, psErrMsg, "]")
        Loop While iFound > 0
  
        If (iStart > 0) And (iStart < Len(Trim(psErrMsg))) Then
            formatError = Trim(Mid(psErrMsg, iStart + 1))
        Else
            formatError = psErrMsg
        End If
    End Function

</script>


<script type="text/javascript">
    util_run_calendarreport_data_window_onload();
</script>


