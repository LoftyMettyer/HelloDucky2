<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script src="<%: Url.Content("~/bundles/utilities_calendarreport_run")%>" type="text/javascript"></script>  

<input type="hidden" name="txtFirstLoad" id="txtFirstLoad" value="<%=Session("CALREP_firstLoad").ToString()%>">

<form action="util_run_calendarreport_data_submit?CalRepUtilID=<%=Session("CalRepUtilID").ToString()%>" method="post" id="frmCalendarGetData" name="frmCalendarGetData">
    <input type="hidden" id="txtDaysInMonth" name="txtDaysInMonth">
    <input type="hidden" id="txtMonth" name="txtMonth">
    <input type="hidden" id="txtYear" name="txtYear">
    <input type="hidden" id="txtVisibleStartDate" name="txtVisibleStartDate">
    <input type="hidden" id="txtVisibleEndDate" name="txtVisibleEndDate">
    <input type="hidden" id="txtMode" name="txtMode">
    <input type="hidden" id="txtLoadCount" name="txtLoadCount" value="0">
    <input type="hidden" id="txtEmailGroupID" name="txtEmailGroupID" value="<%=Session("EmailGroupID").ToString()%>">
</form>

<form id="frmCalendarData" name="frmCalendarData" style="visibility: visible; display: block">
<%
	on error resume next
	
    Dim sErrorDescription As String = ""
    Dim fok As Boolean = True
    Dim fNotCancelled As Boolean
	
    Dim objCalendar As HR.Intranet.Server.CalendarReport
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
    Dim strKeyCode As String
    Dim intEventCounter As Integer
    Dim intSessionStart As Integer
    Dim intSessionEnd As Integer
    Dim strEventID As String
	
    Dim strEventDesc1Value_BD As String
    Dim strEventDesc1ColumnName_BD As String
    Dim strEventDesc2Value_BD As String
    Dim strEventDesc2ColumnName_BD As String
    Dim strBaseDescription_BD As String
	
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
        dtMonth = DateAdd("M", CDbl(lngMonth - Month(objCalendar.ReportStartDate_Calendar)), dtMonth)
  
		mintDaysInMonth = Cint(Session("CALREP_DaysInMonth"))

		'Define the current visible Start and End Dates.
		dtVisibleEndDate = DateAdd("d", CDbl(mintDaysInMonth - Day(dtMonth)), dtMonth)
		dtVisibleStartDate = DateAdd("d", CDbl(-(mintDaysInMonth - 1)), dtVisibleEndDate)
  
'********************************************************************
		
		With rsEvents
			If Not (.BOF And .EOF) Then
			    
			  .MoveFirst
			  Do While Not .EOF
					
                    lngCurrentBaseID = rsEvents.Fields(objCalendar.BaseIDColumn).value
                    strEventID = rsEvents.Fields(objCalendar.EventIDColumn).value
					
					intBaseRecordIndex = objCalendar.BaseIndex_Get(CStr(lngCurrentBaseID))

					strBaseDescription_BD = objCalendar.ConvertDescription(.Fields("Description1").Value, .Fields("Description2").Value, .Fields("DescriptionExpr").Value)
					
                    If IsDBNull(.Fields("Legend").value) Then
                        strKeyCode = ""
                    Else
                        strKeyCode = Left(.Fields("Legend").value, 2)
                    End If
                    If IsDBNull(.Fields("EventDescription1Column").Value) Then
                        strEventDesc1ColumnName_BD = vbNullString
                    Else
                        strEventDesc1ColumnName_BD = CStr(.Fields("EventDescription1Column").Value)
                    End If
                    
                    If IsDBNull(.Fields("EventDescription1ColumnID").value) Then
                        strEventDesc1Value_BD = vbNullString
                    Else
                        strEventDesc1Value_BD = objCalendar.ConvertEventDescription(.Fields("EventDescription1ColumnID").Value, .Fields("EventDescription1").Value)
                    End If
                    
                    If IsDBNull(.Fields("EventDescription2Column").Value) Then
                        strEventDesc2ColumnName_BD = vbNullString
                    Else
                        strEventDesc2ColumnName_BD = CStr(.Fields("EventDescription2Column").Value)
                    End If
                    
                    If Not IsDBNull(.Fields("EventDescription2ColumnID").Value) And Not IsDBNull(.Fields("EventDescription2").Value) Then
                        strEventDesc2Value_BD = objCalendar.ConvertEventDescription(.Fields("EventDescription2ColumnID").Value, .Fields("EventDescription2").Value)
                    Else
                        strEventDesc2Value_BD = vbNullString
                    End If
                    
                   
                    EVENT_DETAIL = vbNullString
                    EVENT_DETAIL = EVENT_DETAIL & .Fields("Name").value & vbTab
                    EVENT_DETAIL = EVENT_DETAIL & ConvertSQLDateToLocale(.Fields("StartDate").value) & vbTab
                    EVENT_DETAIL = EVENT_DETAIL & UCase(.Fields("StartSession").value) & vbTab
                    EVENT_DETAIL = EVENT_DETAIL & ConvertSQLDateToLocale(.Fields("EndDate").value) & vbTab
                    EVENT_DETAIL = EVENT_DETAIL & UCase(.Fields("EndSession").value) & vbTab
                    EVENT_DETAIL = EVENT_DETAIL & FormatNumber(.Fields("Duration").value, 1, True) & vbTab
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

                    strEventToolTip = objCalendar.EventToolTipText(CDate(dtEventStartDate), CStr(strEventStartSession), CDate(dtEventEndDate), CStr(strEventEndSession))
    
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
					  
                        If strEventStartSession = "AM" Then
                            intSessionStart = 0
                        Else
                            intSessionStart = 1
                        End If
						
                        If strEventEndSession = "AM" Then
                            intSessionEnd = 0
                        Else
                            intSessionEnd = 1
                        End If
						
                        intEventCounter = intEventCounter + 1
						
                        INPUT_VALUE = intBaseRecordIndex & "***" & strEventID & "***" & ConvertSQLDateToLocale(dtEventStartDate) & "***" & ConvertSQLDateToLocale(dtEventEndDate) & "***" & intSessionStart & "***" & intSessionEnd & "***" & strEventToolTip & "***" & strKeyCode
					  
                        Response.Write("<INPUT type=hidden name=Event_" & intEventCounter & " ID=Event_" & intEventCounter & " VALUE=""" & INPUT_VALUE & """>" & vbCrLf)
					  
                        Response.Write("<INPUT type=hidden name=EventDetail_" & intEventCounter & " ID=EventDetail_" & intEventCounter & " VALUE=""" & Server.HtmlEncode(EVENT_DETAIL) & """>" & vbCrLf)
					  
                        ' if the event starts in the currently viewed timespan, but ends after it then
                    ElseIf (dtEventStartDate >= dtVisibleStartDate) And (dtEventEndDate > dtVisibleEndDate) Then
					  
                        dtEventEndDate = dtVisibleEndDate
                        strEventEndSession = "PM"

                        If strEventStartSession = "AM" Then
                            intSessionStart = 0
                        Else
                            intSessionStart = 1
                        End If
						
                        If strEventEndSession = "AM" Then
                            intSessionEnd = 0
                        Else
                            intSessionEnd = 1
                        End If
						
                        intEventCounter = intEventCounter + 1
						
                        INPUT_VALUE = intBaseRecordIndex & "***" & strEventID & "***" & ConvertSQLDateToLocale(dtEventStartDate) & "***" & ConvertSQLDateToLocale(dtEventEndDate) & "***" & intSessionStart & "***" & intSessionEnd & "***" & strEventToolTip & "***" & strKeyCode
					  
                        Response.Write("<INPUT type=hidden name=Event_" & intEventCounter & " ID=Event_" & intEventCounter & " VALUE=""" & INPUT_VALUE & """>" & vbCrLf)

                        Response.Write("<INPUT type=hidden name=EventDetail_" & intEventCounter & " ID=EventDetail_" & intEventCounter & " VALUE=""" & Server.HtmlEncode(EVENT_DETAIL) & """>" & vbCrLf)
  
                        ' if the event is enclosed within viewed timespan, and months are equal then
                    ElseIf (dtEventStartDate >= dtVisibleStartDate) And (dtEventEndDate <= dtVisibleEndDate) And (Month(dtEventStartDate) = Month(dtEventEndDate)) Then
  
                        If strEventStartSession = "AM" Then
                            intSessionStart = 0
                        Else
                            intSessionStart = 1
                        End If
						
                        If strEventEndSession = "AM" Then
                            intSessionEnd = 0
                        Else
                            intSessionEnd = 1
                        End If
						
                        intEventCounter = intEventCounter + 1
						
                        INPUT_VALUE = intBaseRecordIndex & "***" & strEventID & "***" & ConvertSQLDateToLocale(dtEventStartDate) & "***" & ConvertSQLDateToLocale(dtEventEndDate) & "***" & intSessionStart & "***" & intSessionEnd & "***" & strEventToolTip & "***" & strKeyCode
					  
                        Response.Write("<INPUT type=hidden name=Event_" & intEventCounter & " ID=Event_" & intEventCounter & " VALUE=""" & INPUT_VALUE & """>" & vbCrLf)

                        Response.Write("<INPUT type=hidden name=EventDetail_" & intEventCounter & " ID=EventDetail_" & intEventCounter & " VALUE=""" & Server.HtmlEncode(EVENT_DETAIL) & """>" & vbCrLf)
     
                        ' if the event starts before the the viewed timespan and ends after the viewed timespan then
                    ElseIf (dtEventStartDate < dtVisibleStartDate) And (dtEventEndDate > dtVisibleEndDate) Then
					  
                        dtEventStartDate = dtVisibleStartDate
                        strEventStartSession = "AM"
					  
                        dtEventEndDate = dtVisibleEndDate
                        strEventEndSession = "PM"
					          
                        If strEventStartSession = "AM" Then
                            intSessionStart = 0
                        Else
                            intSessionStart = 1
                        End If
						
                        If strEventEndSession = "AM" Then
                            intSessionEnd = 0
                        Else
                            intSessionEnd = 1
                        End If
						
                        intEventCounter = intEventCounter + 1
						
                        INPUT_VALUE = intBaseRecordIndex & "***" & strEventID & "***" & ConvertSQLDateToLocale(dtEventStartDate) & "***" & ConvertSQLDateToLocale(dtEventEndDate) & "***" & intSessionStart & "***" & intSessionEnd & "***" & strEventToolTip & "***" & strKeyCode
					  
                        Response.Write("<INPUT type=hidden name=Event_" & intEventCounter & " ID=Event_" & intEventCounter & " VALUE=""" & INPUT_VALUE & """>" & vbCrLf)

                        Response.Write("<INPUT type=hidden name=EventDetail_" & intEventCounter & " ID=EventDetail_" & intEventCounter & " VALUE=""" & Server.HtmlEncode(EVENT_DETAIL) & """>" & vbCrLf)
  
                    End If
					
                    .MoveNext()
                Loop
			  
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

    Html.RenderPartial("Util_Def_CustomReports/ssHiddenGrid")
				
    Response.Write("<INPUT type='hidden' id=txtCalendarPageCount name=txtCalendarPageCount value=" & UBound(arrayMerges) & ">" & vbCrLf)
End If
End If
	
Dim cmdEmailGroup As Object
Dim prmEmailGroupID As Object
Dim rstEmails As Object
Dim iLoop As Integer

If Session("EmailGroupID") > 0 Then
cmdEmailGroup = CreateObject("ADODB.Command")
cmdEmailGroup.CommandText = "spASRIntGetEmailGroupAddresses"
cmdEmailGroup.CommandType = 4 ' Stored procedure
cmdEmailGroup.ActiveConnection = Session("databaseConnection")

prmEmailGroupID = cmdEmailGroup.CreateParameter("EmailGroupID", 3, 1) ' 3=integer, 1=input
cmdEmailGroup.Parameters.Append(prmEmailGroupID)
prmEmailGroupID.value = CleanNumeric(Session("EmailGroupID"))

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
    rstEmails.close()
End If
					
rstEmails = Nothing
cmdEmailGroup = Nothing
Else
Response.Write("<INPUT type=hidden id=txtEmailGroupAddr name=txtEmailGroupAddr value=''>" & vbCrLf)
End If

If Not objCalendar Is Nothing Then
sErrorDescription = objCalendar.ErrorString
End If
	
Response.Write("<INPUT type='hidden' id=txtCalendarMode name=txtCalendarMode value=" & Session("CalRep_Mode") & ">" & vbCrLf)
Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)


%>
</form>

<% 
    If Session("CalRep_Mode") = "OUTPUTREPORT" Then

        Dim iPage As Integer
        Dim iStyle As Integer
        Dim iMerge As Integer
        Dim arrayPageStyles
        Dim arrayPageMerges
	
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

    End If

    If fok And Not objCalendar Is Nothing Then
        objCalendar.OutputArray_Clear()
    End If

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


<script type="text/javascript">
    util_run_calendarreport_data_window_onload();
</script>