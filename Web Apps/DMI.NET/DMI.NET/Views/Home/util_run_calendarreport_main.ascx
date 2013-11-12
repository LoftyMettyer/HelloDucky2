<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script src="<%: Url.Content("~/bundles/utilities_calendarreport_run")%>" type="text/javascript"></script>

<%
	Dim fok As Boolean
	Dim objCalendar As HR.Intranet.Server.CalendarReport
	Dim fNotCancelled As Boolean
	Dim fBadUtilDef As Boolean
	Dim fNoRecords As Boolean
	Dim blnShowCalendar As Boolean
	Dim aPrompts
		
	fBadUtilDef = (Session("utiltype") = "") Or _
		 (Session("utilname") = "") Or _
		 (Session("utilid") = "") Or _
		 (Session("action") = "")
	
	fok = Not fBadUtilDef
	fNotCancelled = True
	
	'objCalendar = Nothing
	Session("objCalendar" & Session("UtilID")) = Nothing
	Session("objCalendar" & Session("UtilID")) = ""
	
	If fok Then
		' Create the reference to the DLL (Report Class)
		objCalendar = New HR.Intranet.Server.CalendarReport
				
		' Pass required info to the DLL
		objCalendar.Username = Session("username")
		CallByName(objCalendar, "Connection", CallType.Let, Session("databaseConnection"))
		objCalendar.CalendarReportID = Session("utilid")
		objCalendar.ClientDateFormat = Session("LocaleDateFormat")
		objCalendar.LocalDecimalSeparator = Session("LocaleDecimalSeparator")
		If CStr(Session("singleRecordID")) = "" Or CStr(Session("singleRecordID")) = "undefined" Then
			objCalendar.SingleRecordID = 0
		Else
			objCalendar.SingleRecordID = Session("singleRecordID")
		End If
		aPrompts = Session("Prompts_" & Session("utiltype") & "_" & Session("UtilID"))
		If fok Then
			fok = objCalendar.SetPromptedValues(aPrompts)
			fNotCancelled = Response.IsClientConnected
			If fok Then fok = fNotCancelled
		End If

		If fok Then
			fok = objCalendar.GetCalendarReportDefinition
			fNotCancelled = Response.IsClientConnected
			If fok Then fok = fNotCancelled
		End If
		
		If fok Then
			fok = objCalendar.GetEventsCollection
			fNotCancelled = Response.IsClientConnected
			If fok Then fok = fNotCancelled
		End If

		If fok Then
			fok = objCalendar.GetOrderArray
			fNotCancelled = Response.IsClientConnected
			If fok Then fok = fNotCancelled
		End If

		If fok Then
			fok = objCalendar.GenerateSQL
			fNotCancelled = Response.IsClientConnected
			If fok Then fok = fNotCancelled
		End If

		If fok Then
			fok = objCalendar.ExecuteSql
			fNotCancelled = Response.IsClientConnected
			If fok Then fok = fNotCancelled
		End If

		If fok Then
			fok = objCalendar.Initialise_WP_Region
			fNotCancelled = Response.IsClientConnected
			If fok Then fok = fNotCancelled
		End If
		
		objCalendar.SetLastRun()

		fNoRecords = objCalendar.NoRecords

		' Convert data over to DataTables (remove step at later date when rest of code converted)
		If fok Then		
			objCalendar.Events = RecordSetToDataTable(objCalendar.EventsRecordset)
		End If
		
		
		If fok Then
			If Response.IsClientConnected Then
				objCalendar.Cancelled = False
			Else
				objCalendar.Cancelled = True
			End If
		Else
			If Not fNoRecords Then
				If fNotCancelled Then
					objCalendar.FailedMessage = objCalendar.ErrorString
					objCalendar.Failed = True
				Else
					objCalendar.Cancelled = True
				End If
			End If
		End If
		
		blnShowCalendar = (objCalendar.OutputPreview Or (objCalendar.OutputFormat = 0 And objCalendar.OutputScreen))
		
		Session("objCalendar" & Session("UtilID")) = objCalendar

				
	End If

%>
<input type='hidden' id="txtLoadCount" name="txtLoadCount" value="0">
<input type='hidden' id="txtOK" name="txtOK" value="True">
<%
		
	Session("CalRepUtilID") = Request.Form("utilid")
		
	If blnShowCalendar Then
		Response.Write("<input type='hidden' id=txtPreview name=txtPreview value=1>" & vbCrLf)
	Else
		Response.Write("<input type='hidden' id=txtPreview name=txtPreview value=0>" & vbCrLf)
	End If
		
	If blnShowCalendar Then
%>

<div id="reportworkframe" data-framesource="util_run_calendarreport_data" style="display: inline-block; width: 100%">
	<%Html.RenderPartial("~/views/home/util_run_calendarreport_data.ascx")%>
</div>

<div id="reportdataframe" style="display: none;" />
	
<div id="outputoptions" data-framesource="util_run_outputoptions" style="display: none;">
	<%	Html.RenderPartial("~/Views/Home/util_run_outputoptions.ascx")%>
</div>

<%
Else
	If fBadUtilDef Then
%>

<input type='hidden' id="txtOK" name="txtOK" value="False">
<table align="center" class="outline" cellpadding="5" cellspacing="0">
	<tr>
		<td>
			<table class="invisible" cellspacing="0" cellpadding="0">
				<tr>
					<td colspan="3" height="10"></td>
				</tr>
				<tr>
					<td colspan="3" align="center">
						<h3>Error</h3>
					</td>
				</tr>
				<tr>
					<td width="20" height="10"></td>
					<td>
						<h4>Not all session variables found</h4>
					</td>
					<td width="20"></td>
				</tr>
				<tr>
					<td width="20" height="10"></td>
					<td>Type = <%Session("utiltype").ToString()%>
					</td>
					<td width="20"></td>
				</tr>
				<tr>
					<td width="20" height="10"></td>
					<td>Utility Name = <%Session("utilname").ToString()%>
					</td>
					<td width="20"></td>
				</tr>
				<tr>
					<td width="20" height="10"></td>
					<td>Utility ID = <%Session("utilid").ToString()%>
					</td>
					<td width="20"></td>
				</tr>
				<tr>
					<td width="20" height="10"></td>
					<td>Action = <%Session("action").ToString()%>
					</td>
					<td width="20"></td>
				</tr>
				<tr>
					<td colspan="3" height="10">&nbsp;</td>
				</tr>
				<tr>
					<td colspan="3" height="10" align="center">
						<input type="button" value="Close" name="cmdClose" style="WIDTH: 80px" width="80" id="cmdClose" class="btn"
							onclick="closeclick();" />
					</td>
				</tr>
				<tr>
					<td colspan="3" height="10"></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<input type="hidden" id="txtSuccessFlag" name="txtSuccessFlag" value="1">

	<%
Else
	%>

<input type='hidden' id="txtOK" name="txtOK" value="False">
<form id="frmPopup" name="frmPopup">
	<table align="center" class="outline" cellpadding="5" cellspacing="0">
		<tr>
			<td>
				<table class="invisible" cellspacing="0" cellpadding="0">
					<tr>
						<td colspan="3" height="10"></td>
					</tr>
					<%
						Dim sCloseFunction As String
		
						Response.Write("			  <tr> " & vbCrLf)
						Response.Write("			    <td width=20 height=10></td> " & vbCrLf)
						Response.Write("			    <td align=center> " & vbCrLf)

						If fNoRecords Then
							Response.Write("						<H4>Calendar Report '" & Session("utilname") & "' Completed successfully.</H4>" & vbCrLf)
							sCloseFunction = "closeclick();"
						Else
							Response.Write("						<H4>Calendar Report '" & Session("utilname") & "' Failed." & vbCrLf)
							sCloseFunction = "closeclick();"
						End If
						Response.Write("			    </td>" & vbCrLf)
						Response.Write("			    <td width=20></td> " & vbCrLf)
						Response.Write("			  </tr>" & vbCrLf)
					%>
					<tr>
						<td width="20" height="10"></td>
						<td align="center" nowrap>
							<%=objCalendar.ErrorString%>
						</td>
						<td width="20"></td>
					</tr>
					<tr>
						<td colspan="3" height="10">&nbsp;</td>
					</tr>
					<tr>
						<td colspan="3" height="10" align="center">
							<input type="button" value="Close" name="cmdClose" style="WIDTH: 80px" width="80" id="cmdClose" class="btn"
								onclick="closeclick();" />
						</td>
					</tr>
					<tr>
						<td colspan="3" height="10"></td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</form>
<input type="hidden" id="txtSuccessFlag" name="txtSuccessFlag" value="1">
<input type='hidden' id="txtPreview" name="txtPreview" value="0">
<%
End If
End If

Response.Write("<input type=hidden id=txtTitle name=txtTitle value=""" & Replace(objCalendar.CalendarReportName, """", "&quot;") & """>" & vbCrLf)
objCalendar = Nothing
%>

<form id="frmOriginalDefinition" style="visibility: hidden; display: none">
	<%
		Dim sErrMsg As String = ""
		Response.Write("	<input type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Replace(Session("utilname").ToString(), """", "&quot;") & """>" & vbCrLf)
		Response.Write("	<input type='hidden' id=txtDefn_ErrMsg name=txtDefn_ErrMsg value=""" & sErrMsg & """>" & vbCrLf)
	%>
	<input type="hidden" id="txtUserName" name="txtUserName" value="<%Session("username").ToString()%>">
	<input type="hidden" id="txtDateFormat" name="txtDateFormat" value="<%Session("LocaleDateFormat").ToString()%>">
	<input type="hidden" id="txtCancelPrint" name="txtCancelPrint">
	<input type="hidden" id="txtOptionsDone" name="txtOptionsDone">
	<input type="hidden" id="txtOptionsPortrait" name="txtOptionsPortrait">
	<input type="hidden" id="txtOptionsMarginLeft" name="txtOptionsMarginLeft">
	<input type="hidden" id="txtOptionsMarginRight" name="txtOptionsMarginRight">
	<input type="hidden" id="txtOptionsMarginTop" name="txtOptionsMarginTop">
	<input type="hidden" id="txtOptionsMarginBottom" name="txtOptionsMarginBottom">
	<input type="hidden" id="txtOptionsCopies" name="txtOptionsCopies">
	<input type="hidden" id="txtCalRep_UtilID" name="txtCalRep_UtilID" value="<%Session("UtilID").ToString()%>">
</form>

<script type="text/javascript">

	$("#reportframe").show();

	util_run_calendarreport_main_window_onload();
	$(".popup").dialog('option', 'title', $("#txtTitle").val());
	$("#top").hide();
	$("#calendarframeset").show();

</script>
