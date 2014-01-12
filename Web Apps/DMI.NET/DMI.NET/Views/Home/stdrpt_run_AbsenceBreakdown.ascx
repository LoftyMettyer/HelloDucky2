<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="ADODB" %>

<script src="<%: Url.Content("~/bundles/OpenHR_General")%>" type="text/javascript"></script>
<script src="<%: Url.Content("~/bundles/utilities_crosstabs")%>" type="text/javascript"></script>

<%

	Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)

		Dim dtStartDate
		Dim dtEndDate
		Dim strAbsenceTypes As String
		Dim lngFilterID As Long
		Dim lngPicklistID As Long
		Dim lngPersonnelID As Long
		Dim bPrintFilterPickList As Boolean

	' Default output options
		Dim bOutputPreview As Boolean
	Dim lngOutputFormat As Integer
		Dim pblnOutputScreen As Boolean
		Dim pblnOutputPrinter As Boolean
		Dim pstrOutputPrinterName As String
		Dim pblnOutputSave As Boolean
		Dim plngOutputSaveExisting As Long
		Dim pblnOutputEmail As Boolean
		Dim plngOutputEmailID As Long
		Dim pstrOutputEmailName As String
		Dim pstrOutputEmailSubject As String
		Dim pstrOutputEmailAttachAs As String
		Dim pstrOutputFilename As String

	Dim lngHorColID As Integer
	Dim lngVerColID As Integer		
	Dim lngStartDateColID As Integer
	Dim lngStartSessionColID As Integer
	Dim lngEndDateColID As Integer
	Dim lngEndSessionColID As Integer
	Dim lngTypeColID As Integer
	Dim lngReasonColID As Integer
	Dim lngDurationColID As Integer
	Dim iParameterValue As Integer
	
	' Get variables for Absence Breakdown / Bradford Factor
	dtStartDate = convertLocaleDateToSQL(session("stdReport_StartDate"))
	dtEndDate = convertLocaleDateToSQL(session("stdReport_EndDate"))
	strAbsenceTypes = session("stdReport_AbsenceTypes")
	lngFilterID = session("stdReport_FilterID")
	lngPicklistID = session("stdReport_PicklistID")
	lngPersonnelID = session("optionRecordID")
	bPrintFilterPickList = session("stdReport_PrintFilterPicklistHeader")

	' Default output options
	bOutputPreview = session("stdReport_OutputPreview")
	lngOutputFormat = session("stdReport_OutputFormat")
	pblnOutputScreen = session("stdReport_OutputScreen")
	pblnOutputPrinter = session("stdReport_OutputPrinter")
	pstrOutputPrinterName = session("stdReport_OutputPrinterName")
	pblnOutputSave = session("stdReport_OutputSave")
	plngOutputSaveExisting = session("stdReport_OutputSaveExisting")
	pblnOutputEmail = session("stdReport_OutputEmail")
	plngOutputEmailID = session("stdReport_OutputEmailAddr")
	pstrOutputEmailSubject = session("stdReport_OutputEmailSubject")
	pstrOutputEmailAttachAs = session("stdReport_OutputEmailAttachAs")
	pstrOutputFilename = session("stdReport_OutputFilename")
	
	iParameterValue = CInt(objDatabase.GetModuleParameter("MODULE_ABSENCE", "Param_FieldStartSession"))		
	lngHorColID = iParameterValue
	lngStartSessionColID = iParameterValue
	
	iParameterValue = CInt(objDatabase.GetModuleParameter("MODULE_ABSENCE", "Param_FieldType"))
	lngVerColID = iParameterValue
	lngTypeColID = iParameterValue
		
	lngStartDateColID = CInt(objDatabase.GetModuleParameter("MODULE_ABSENCE", "Param_FieldStartDate"))
	lngEndDateColID = CInt(objDatabase.GetModuleParameter("MODULE_ABSENCE", "Param_FieldEndDate"))
	lngEndSessionColID = CInt(objDatabase.GetModuleParameter("MODULE_ABSENCE", "Param_FieldEndSession"))	
	lngReasonColID = CInt(objDatabase.GetModuleParameter("MODULE_ABSENCE", "Param_FieldReason"))
	lngDurationColID = CInt(objDatabase.GetModuleParameter("MODULE_ABSENCE", "Param_FieldDuration"))


		Dim fok As Boolean
		Dim objCrossTab As HR.Intranet.Server.CrossTab
		Dim fNotCancelled As Boolean
		Dim lngEventLogID As Long
		Dim blnNoDefinition As Boolean

		Session("objCrossTab" & Session("utilid")) = Nothing
	Session("CT_Mode") = ""
	Session("CT_PageNumber") = ""
	Session("CT_IntersectionType") = ""
	Session("CT_ShowPercentage") = ""
	Session("CT_PercentageOfPage") = ""
	Session("CT_SupressZeros") = ""

	if session("utiltype") = "" or _ 
		 session("utilname") = "" or _ 
		 session("utilid") = "" or _ 
		 session("action") = "" then 
				
				Response.Write("Error : Not all session variables found...<HR>")
				Response.Write("Type = " & Session("utiltype") & "<BR>")
				Response.Write("UtilName = " & Session("utilname") & "<BR>")
				Response.Write("UtilID = " & Session("utilid") & "<BR>")
				Response.Write("Action = " & Session("action") & "<BR>")
				Response.End()
	end if

	' Create the reference to the DLL (Report Class)
	objCrossTab = New CrossTab()
	objCrossTab.SessionInfo = CType(Session("SessionContext"), SessionInfo)

	Session("objCrossTab" & Session("utilid")) = Nothing

	' Pass required info to the DLL
	objCrossTab.CrossTabID = Session("utilid")
	objCrossTab.ClientDateFormat = session("localedateformat")
	objCrossTab.LocalDecimalSeparator = session("LocaleDecimalSeparator")

	fok = true
	blnNoDefinition = true

	Dim aPrompts
		Dim fModuleOk As Boolean
		
	aPrompts = Session("Prompts_" & session("utiltype") & "_" & session("utilid"))

	fModuleOK = true
	if lngStartDateColID = 0 or _
		lngStartSessionColID = 0 or _
		lngEndDateColID = 0 or _
		lngEndSessionColID = 0 or _
		lngTypeColID = 0 or _
		lngReasonColID = 0 or _
		lngDurationColID = 0 then
		
		fok = false
		fModuleOK = false
	end if

	if fok then 
		fok = objCrossTab.SetPromptedValues(aPrompts)
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok then
		fok = objCrossTab.SetAbsenceBreakDownDisplayOptions(bPrintFilterPickList)
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok then
		fok = objCrossTab.SetAbsenceBreakDownDisplayOptions(bPrintFilterPickList)
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok then
		fok = objCrossTab.AbsenceBreakdownRetreiveDefinition(dtStartDate, dtEndDate, lngHorColID, lngVerColID, lngPicklistID, lngFilterID, lngPersonnelID, strAbsenceTypes)
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok then
		blnNoDefinition = false
		lngEventLogID = objCrossTab.EventLogAddHeader
		fok = (lngEventLogID > 0)
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok then
		fok = objCrossTab.UDFFunctions(true)
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if
	
	if fok then
		fok = objCrossTab.CreateTempTable
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok then
		fok = objCrossTab.UDFFunctions(false)
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok then 
		fok = objCrossTab.AbsenceBreakdownRunStoredProcedure
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok then
		fok = objCrossTab.AbsenceBreakdownGetHeadingsAndSearches
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok then
		fok = objCrossTab.BuildTypeArray
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok then
		fok = objCrossTab.AbsenceBreakdownBuildDataArrays
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	' Need to pass in defined output options 
	'	(standard cross tab reads from definition, which of course we don't have in a standard report)
	if fok then
		fok = objCrossTab.SetAbsenceBreakDownDefaultOutputOptions(bOutputPreview, lngOutputFormat, pblnOutputScreen, pblnOutputPrinter, pstrOutputPrinterName, pblnOutputSave, plngOutputSaveExisting, pblnOutputEmail, plngOutputEmailID, pstrOutputEmailName, pstrOutputEmailSubject, pstrOutputEmailAttachAs, pstrOutputFilename)
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if		

		Session("objCrossTab" & Session("utilid")) = objCrossTab

		Response.Write("<script type=""text/javascript"">" & vbCrLf)
		Response.Write("function crosstab_loadAddRecords()" & vbCrLf)
		Response.Write("{" & vbCrLf)
		Response.Write("	var iCount;" & vbCrLf & vbCrLf)

		Response.Write("	iCount = new Number(txtLoadCount.value);" & vbCrLf)
		Response.Write("	txtLoadCount.value = iCount + 1;" & vbCrLf & vbCrLf)

		Response.Write("  if (iCount > 0) {	" & vbCrLf)
		Response.Write("    var frmGetData = OpenHR.getForm(""reportdataframe"",""frmGetReportData"");" & vbCrLf)
		Response.Write("    frmGetData.txtUtilID.value = """ & Session("utilid") & """;" & vbCrLf)
	Response.Write("    getCrossTabData(""LOAD"",0,0,0,0,0,0);" & vbCrLf & vbCrLf)
		Response.Write("  }" & vbCrLf & vbCrLf)

		Response.Write("}" & vbCrLf)

		Response.Write("</script>" & vbCrLf)

%>

<input type='hidden' id="txtLoadCount" name="txtLoadCount" value="0">
<input type='hidden' id="txtModuleOK" name="txtModuleOK" value="<%=fModuleOK%>">

<%	
if fModuleOK then
%>

<div id="main" data-framesource="stdrpt_run_AbsenceBreakdown" style="display: block;">
		<div id="reportworkframe" data-framesource="util_run_crosstabs" style="display: block;">
				<%Html.RenderPartial("~/views/home/util_run_crosstabs.ascx")%>
		</div>

		<div id="reportdataframe" data-framesource="util_run_crosstabsData" style="display: none;" accesskey="">
				<%Html.RenderPartial("~/views/home/util_run_crosstabsData.ascx")%>
		</div>
		
		<div id="reportbreakdownframe" data-framesource="util_run_crosstabsBreakdown" style="display: none;" accesskey="">
				<%Html.RenderPartial("~/views/home/util_run_crosstabsBreakdown.ascx")%>
		</div>

		<div id="outputoptions" data-framesource="util_run_outputoptions" style="display: none;">
				<% Html.RenderPartial("~/Views/Home/util_run_outputoptions.ascx")%>
		</div>
</div>


<form id="frmOutput" name="frmOutput">
		<input type="hidden" id="fok" name="fok" value="">
		<input type="hidden" id="cancelled" name="cancelled" value="">
		<input type="hidden" id="statusmessage" name="statusmessage" value="">
</form>

<%	
else
%>

<form Name=frmPopup ID=frmPopup>
<table align=center class="outline" cellPadding=5 cellSpacing=0>
	<TR>
		<TD>
			<table class="invisible" cellspacing=0 cellpadding=0>
				<tr>
					<td colspan=3 height=10></td>
				</tr>
				<tr> 
					<td width=20 height=10></td>
					<td align=center>
						<H4>Absence Breakdown Failed.</H4>
					</td>
					<td width=20></td>
				</tr>
				<tr>
					<td width=20 height=10></td>
					<td align=center nowrap>Module setup has not been completed.
					</td>
					<td width=20></td>
				</tr>
				<tr>
					<td colspan=3 height=10>&nbsp;</td>
				</tr>
				<tr> 
					<td colspan=3 height=10 align=center>
						<input type="button" value="Close" name="cmdClose" class="btn" style="WIDTH: 80px" width="80" id="cmdClose" onclick="closeclick();" />
					</td>
				</tr>
				<tr> 
					<td colspan=3 height=10></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</form>

<%
end if
%>


<script type="text/javascript">

	util_run_crosstabs_window_onload();

	$("#reportframe").show();
	$("#top").hide();
	$("#reportworkframe").show();

</script>
