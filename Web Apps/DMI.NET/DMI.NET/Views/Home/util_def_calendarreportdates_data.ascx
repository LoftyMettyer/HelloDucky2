<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script type="text/javascript">
	function util_validate_calendarreportdates_data_window_onload() {
		//debugger;
		var frmCalendarData = document.getElementById("frmCalendarData");
		if (frmCalendarData.txtCalendarAction.value == "LOADCALENDAREVENTDETAILSCOLUMNS") {
			//window.parent.frames("calendarworkframe").loadAvailableEventColumns();
			loadAvailableEventColumns();
		}
		else if (frmCalendarData.txtCalendarAction.value == "LOADCALENDAREVENTKEYLOOKUPCOLUMNS") {
			//window.parent.frames("calendarworkframe").loadAvailableLookupColumns();
			loadAvailableLookupColumns();
		}
	}

	function data_refreshData() {
		var frmGetCalendarData = document.getElementById("frmGetCalendarData");
		OpenHR.submitForm(frmGetCalendarData);
	}
</script>

<div>
	<form action="util_def_calendarreportdates_data_submit" method="post" id="frmGetCalendarData" name="frmGetCalendarData">
		<input type="hidden" id="txtCalendarAction" name="txtCalendarAction">
		<input type="hidden" id="txtCalendarBaseTableID" name="txtCalendarBaseTableID" value="0">
		<input type="hidden" id="txtCalendarEventTableID" name="txtCalendarEventTableID" value="0">
		<input type="hidden" id="txtCalendarLookupTableID" name="txtCalendarLookupTableID" value="0">
	</form>

	<form id="frmCalendarData" name="frmCalendarData">
		<!-- INPUT element containing the required data for the calendar reports dates form -->
		<%
			Dim sErrorDescription = ""

			If Session("CalendarAction") = "LOADCALENDAREVENTDETAILSCOLUMNS" Then
				'Response.Write("<FONT COLOR=red><B>Base 1 Table : " & Session("CalendarBaseTableID") & "<B></FONT><BR>")
				'Response.Write("<FONT COLOR=red><B>Event Table : " & Session("CalendarEventTableID") & "<B></FONT>")
		
				Dim cmdEventCols = CreateObject("ADODB.Command")
				cmdEventCols.CommandText = "spASRIntGetCalendarReportColumns"
				cmdEventCols.CommandType = 4
				cmdEventCols.ActiveConnection = Session("databaseConnection")
				
				Dim prmBaseTableID = cmdEventCols.CreateParameter("baseTableID", 3, 1) ' 3=integer, 1=input
				cmdEventCols.Parameters.Append(prmBaseTableID)
				prmBaseTableID.value = CleanNumeric(Session("CalendarBaseTableID"))

				Dim prmEventTableID = cmdEventCols.CreateParameter("eventTableID", 3, 1) ' 3=integer, 1=input
				cmdEventCols.Parameters.Append(prmEventTableID)
				prmEventTableID.value = CleanNumeric(Session("CalendarEventTableID"))

				Err.Clear()
				Dim rstEventColumns = cmdEventCols.Execute

				If (Err.Number <> 0) Then
					sErrorDescription = "Error getting the calendar report event columns." & vbCrLf & FormatError(Err.Description)
				End If
		
				If Len(sErrorDescription) = 0 Then
					Dim iLoop = 1
					Do While Not rstEventColumns.EOF
						Response.Write("<INPUT type='hidden' id=txtRepCol_" & rstEventColumns.Fields("columnid").Value & " name=txtRepCol_" & rstEventColumns.Fields("columnid").Value & " value=" & Replace(rstEventColumns.Fields("columnName").Value, """", "&quot;") & ">" & vbCrLf)
						Response.Write("<INPUT type='hidden' id=txtRepColDataType_" & rstEventColumns.Fields("columnid").Value & " name=txtRepColDataType_" & rstEventColumns.Fields("columnid").Value & " value='" & Replace(rstEventColumns.Fields("datatype").Value, """", "&quot;") & "'>" & vbCrLf)
						Response.Write("<INPUT type='hidden' id=txtRepColSize_" & rstEventColumns.Fields("columnid").Value & " name=txtRepColSize_" & rstEventColumns.Fields("columnid").Value & " value='" & Replace(rstEventColumns.Fields("size").Value, """", "&quot;") & "'>" & vbCrLf)
						Response.Write("<INPUT type='hidden' id=txtRepColTableID_" & rstEventColumns.Fields("columnid").Value & " name=txtRepColTableID_" & rstEventColumns.Fields("columnid").Value & " value='" & Replace(rstEventColumns.Fields("tableid").Value, """", "&quot;") & "'>" & vbCrLf)
						Response.Write("<INPUT type='hidden' id=txtRepColTableName_" & rstEventColumns.Fields("columnid").Value & " name=txtRepColTableName_" & rstEventColumns.Fields("columnid").Value & " value='" & Replace(rstEventColumns.Fields("tablename").Value, """", "&quot;") & "'>" & vbCrLf)
						Response.Write("<INPUT type='hidden' id=txtRepColType_" & rstEventColumns.Fields("columnid").Value & " name=txtRepColType_" & rstEventColumns.Fields("columnid").Value & " value='" & Replace(rstEventColumns.Fields("columntype").Value, """", "&quot;") & "'>" & vbCrLf)
				
						rstEventColumns.MoveNext()
						iLoop = iLoop + 1
					Loop

					' Release the ADO recordset object.
					rstEventColumns.close()
				End If
			
				rstEventColumns = Nothing
				cmdEventCols = Nothing

			ElseIf Session("CalendarAction") = "LOADCALENDAREVENTKEYLOOKUPCOLUMNS" Then

				Dim cmdKeyLookupCols = CreateObject("ADODB.Command")
				cmdKeyLookupCols.CommandText = "spASRIntGetCalendarReportColumns"
				cmdKeyLookupCols.CommandType = 4 ' Stored procedure
				cmdKeyLookupCols.ActiveConnection = Session("databaseConnection")
								
				Dim prmBaseTableID = cmdKeyLookupCols.CreateParameter("baseTableID", 3, 1) ' 3=integer, 1=input
				cmdKeyLookupCols.Parameters.Append(prmBaseTableID)
				prmBaseTableID.value = CleanNumeric(Session("CalendarLookupTableID"))

				Dim prmEventTableID = cmdKeyLookupCols.CreateParameter("eventTableID", 3, 1) ' 3=integer, 1=input
				cmdKeyLookupCols.Parameters.Append(prmEventTableID)
				prmEventTableID.value = CleanNumeric(Session("CalendarLookupTableID"))
		
				Err.Clear()
				Dim rstLookupColumns = cmdKeyLookupCols.Execute
		
				If (Err.Number <> 0) Then
					sErrorDescription = "Error getting the calendar report columns." & vbCrLf & FormatError(Err.Description)
				End If
		
				If Len(sErrorDescription) = 0 Then
					Dim iLoop = 1
					Do While Not rstLookupColumns.EOF
						Response.Write("<INPUT type='hidden' id=txtRepCol_" & rstLookupColumns.Fields("columnid").Value & " name=txtRepCol_" & rstLookupColumns.Fields("columnid").Value & " value='" & Replace(rstLookupColumns.Fields("columnName").Value, """", "&quot;") & "'>" & vbCrLf)
						Response.Write("<INPUT type='hidden' id=txtRepColDataType_" & rstLookupColumns.Fields("columnid").Value & " name=txtRepColDataType_" & rstLookupColumns.Fields("columnid").Value & " value='" & Replace(rstLookupColumns.Fields("datatype").Value, """", "&quot;") & "'>" & vbCrLf)
						rstLookupColumns.MoveNext()
						iLoop = iLoop + 1
					Loop
		
					' Release the ADO recordset object.
					rstLookupColumns.close()
				End If
				
				rstLookupColumns = Nothing
				cmdKeyLookupCols = Nothing
		
			End If

			Response.Write("<INPUT type='hidden' id=txtCalendarAction name=txtCalendarAction value=" & Session("CalendarAction") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)

			Session("CalendarAction") = ""
		%>
	</form>
</div>

<script type="text/javascript">
	util_validate_calendarreportdates_data_window_onload();
</script>
