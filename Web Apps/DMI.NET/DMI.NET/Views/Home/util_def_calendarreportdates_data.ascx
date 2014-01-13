<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>

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
			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
			
			If Session("CalendarAction") = "LOADCALENDAREVENTDETAILSCOLUMNS" Then
			
				Try
			
					Dim rstEventColumns = objDataAccess.GetDataTable("spASRIntGetCalendarReportColumns", CommandType.StoredProcedure _
							, New SqlParameter("piBaseTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("CalendarBaseTableID"))} _
							, New SqlParameter("piEventTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("CalendarEventTableID"))})

					For Each objRow As DataRow In rstEventColumns.Rows
						Response.Write("<input type='hidden' id=txtRepCol_" & objRow("columnid") & " name=txtRepCol_" & objRow("columnid") & " value=" & Replace(objRow("columnName").ToString(), """", "&quot;") & ">" & vbCrLf)
						Response.Write("<input type='hidden' id=txtRepColDataType_" & objRow("columnid") & " name=txtRepColDataType_" & objRow("columnid") & " value='" & Replace(objRow("datatype").ToString(), """", "&quot;") & "'>" & vbCrLf)
						Response.Write("<input type='hidden' id=txtRepColSize_" & objRow("columnid") & " name=txtRepColSize_" & objRow("columnid") & " value='" & Replace(objRow("size").ToString(), """", "&quot;") & "'>" & vbCrLf)
						Response.Write("<input type='hidden' id=txtRepColTableID_" & objRow("columnid") & " name=txtRepColTableID_" & objRow("columnid") & " value='" & Replace(objRow("tableid").ToString(), """", "&quot;") & "'>" & vbCrLf)
						Response.Write("<input type='hidden' id=txtRepColTableName_" & objRow("columnid") & " name=txtRepColTableName_" & objRow("columnid") & " value='" & Replace(objRow("tablename").ToString(), """", "&quot;") & "'>" & vbCrLf)
						Response.Write("<input type='hidden' id=txtRepColType_" & objRow("columnid") & " name=txtRepColType_" & objRow("columnid") & " value='" & Replace(objRow("columntype").ToString(), """", "&quot;") & "'>" & vbCrLf)
					Next
			
				Catch ex As Exception
					sErrorDescription = "Error getting the calendar report columns." & vbCrLf & FormatError(ex.Message)

				End Try
				
			ElseIf Session("CalendarAction") = "LOADCALENDAREVENTKEYLOOKUPCOLUMNS" Then

				Try
			
					Dim rstLookupColumns = objDataAccess.GetDataTable("spASRIntGetCalendarReportColumns", CommandType.StoredProcedure _
							, New SqlParameter("piBaseTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("CalendarLookupTableID"))} _
							, New SqlParameter("piEventTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("CalendarLookupTableID"))})

					For Each objRow As DataRow In rstLookupColumns.Rows
						Response.Write("<input type='hidden' id='txtRepCol_" & objRow("columnid") & "' name='txtRepCol_" & objRow("columnid") & "' value='" & Replace(objRow("columnName").ToString(), """", "&quot;") & "'>" & vbCrLf)
						Response.Write("<input type='hidden' id='txtRepColDataType_" & objRow("columnid") & "' name='txtRepColDataType_" & objRow("columnid") & "' value='" & Replace(objRow("datatype").ToString(), """", "&quot;") & "'>" & vbCrLf)
					Next
			
				Catch ex As Exception
					sErrorDescription = "Error getting the calendar report columns." & vbCrLf & FormatError(ex.Message)

				End Try
					
			End If

			Response.Write("<input type='hidden' id=txtCalendarAction name=txtCalendarAction value=" & Session("CalendarAction") & ">" & vbCrLf)
			Response.Write("<input type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)

			Session("CalendarAction") = ""
		%>
	</form>
</div>

<script type="text/javascript">
	util_validate_calendarreportdates_data_window_onload();
</script>
