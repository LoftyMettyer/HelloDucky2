<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<%
	Dim fok As Boolean = True
	Dim objCalendar As New CalendarReport
	Dim fNotCancelled As Boolean = True
	Dim fNoRecords As Boolean
	Dim blnShowCalendar As Boolean
	Dim aPrompts
			
	'objCalendar = Nothing
	Session("objCalendar" & Session("UtilID")) = Nothing
	Session("objCalendar" & Session("UtilID")) = ""
	
	If fok Then

		objCalendar.SessionInfo = CType(Session("SessionContext"), SessionInfo)
					
		' Pass required info to the DLL
		objCalendar.Initialise()
		objCalendar.CalendarReportID = Session("utilid")
		objCalendar.ClientDateFormat = Session("LocaleDateFormat")
		objCalendar.LocalDecimalSeparator = Session("LocaleDecimalSeparator")
		objCalendar.SingleRecordID = Session("singleRecordID")
		
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
		
		fNoRecords = objCalendar.NoRecords

		' Convert data over to DataTables (remove step at later date when rest of code converted)
		If fok Then
			objCalendar.Events = objCalendar.EventsRecordset
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

		objCalendar.ClearUp()

		blnShowCalendar = (objCalendar.OutputPreview Or (objCalendar.OutputFormat = 0 And objCalendar.OutputScreen) Or objCalendar.OutputPrinter)
		
		Session("objCalendar" & Session("UtilID")) = objCalendar
				
	End If

%>
<input type='hidden' id="txtLoadCount" name="txtLoadCount" value="0">
<input type='hidden' id="txtOK" name="txtOK" value="True">
<%
		
	Session("CalRepUtilID") = Request.Form("utilid")
		
	Response.Write("<input type='hidden' id=txtPreview name=txtPreview value=" & blnShowCalendar & ">" & vbCrLf)
		
	If Not fNoRecords Then
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
	Session("CalendarReports_FailedOrNoRecords") = True
End If

%>

<input type='hidden' id="txtNoRecs" name="txtNoRecs" value="<%=objCalendar.NoRecords%>">

<form id="frmOriginalDefinition" style="visibility: hidden; display: none">
	<%
		Response.Write("	<input type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Replace(Session("utilname").ToString(), """", "&quot;") & """>" & vbCrLf)
		Response.Write("	<input type='hidden' id=txtDefn_ErrMsg name=txtDefn_ErrMsg value=""" & objCalendar.ErrorString & """>" & vbCrLf)
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
