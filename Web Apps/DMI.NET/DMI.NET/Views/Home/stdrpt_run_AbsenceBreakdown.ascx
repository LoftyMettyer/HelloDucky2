<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="System.Globalization" %>

<script src="<%: Url.LatestContent("~/bundles/OpenHR_General")%>" type="text/javascript"></script>
<script src="<%: Url.LatestContent("~/bundles/utilities_crosstabs")%>" type="text/javascript"></script>

<%

	Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)

	Dim dtStartDate As Date
	Dim dtEndDate As Date
	Dim strAbsenceTypes As String
	Dim lngFilterID As Long
	Dim lngPicklistID As Long
	Dim lngPersonnelID As Long
	Dim bPrintFilterPickList As Boolean

	' Default output options
		Dim bOutputPreview As Boolean
	Dim lngOutputFormat As Integer
		Dim pblnOutputScreen As Boolean
	Dim pblnOutputSave As Boolean
		Dim plngOutputSaveExisting As Long
		Dim pblnOutputEmail As Boolean
	Dim plngOutputEmailID As Integer
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
	Dim objCrossTab As CrossTab
	
	' Create the reference to the DLL (Report Class)
	objCrossTab = New CrossTab()
	objCrossTab.SessionInfo = CType(Session("SessionContext"), SessionInfo)

	Session("objCrossTab" & Session("utilid")) = Nothing

	' Get variables for Absence Breakdown / Bradford Factor
	dtStartDate = DateTime.ParseExact(Session("stdReport_StartDate").ToString(), "MM/dd/yyyy", CultureInfo.InvariantCulture)
	dtEndDate = DateTime.ParseExact(Session("stdReport_EndDate").ToString(), "MM/dd/yyyy", CultureInfo.InvariantCulture)
	
	strAbsenceTypes = session("stdReport_AbsenceTypes")
	lngFilterID = session("stdReport_FilterID")
	lngPicklistID = session("stdReport_PicklistID")
	lngPersonnelID = session("optionRecordID")
	bPrintFilterPickList = session("stdReport_PrintFilterPicklistHeader")

	' Default output options
	objCrossTab.Name = "Absence Breakdown"
	objCrossTab.OutputFormat = Session("stdReport_OutputFormat")
	objCrossTab.OutputPreview = Session("stdReport_OutputPreview")
	objCrossTab.OutputFilename = Session("stdReport_OutputFilename")
	
	bOutputPreview = objCrossTab.OutputPreview
	lngOutputFormat = objCrossTab.OutputFormat
	pblnOutputScreen = False
	pblnOutputSave = Session("stdReport_OutputSave")
	plngOutputSaveExisting = session("stdReport_OutputSaveExisting")
	pblnOutputEmail = session("stdReport_OutputEmail")
	plngOutputEmailID = session("stdReport_OutputEmailAddr")
	pstrOutputEmailSubject = session("stdReport_OutputEmailSubject")
	pstrOutputEmailAttachAs = session("stdReport_OutputEmailAttachAs")
	pstrOutputFilename = objCrossTab.OutputFilename
	
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
	Dim fNotCancelled As Boolean
		Dim lngEventLogID As Long
		Dim blnNoDefinition As Boolean

		Session("objCrossTab" & Session("utilid")) = Nothing
	Session("CT_Mode") = ""
	Session("CT_PageNumber") = 0
	Session("CT_IntersectionType") = ""
	Session("CT_ShowPercentage") = False
	Session("CT_PercentageOfPage") = False
	Session("CT_SupressZeros") = False

	If Session("utiltype") Is Nothing Or _
		 Session("utilname") Is Nothing Or _
		 Session("utilid") Is Nothing Or _
		 Session("action") Is Nothing Then
				
		Response.Write("Error : Not all session variables found...<HR>")
		Response.Write("Type = " & Session("utiltype") & "<BR>")
		Response.Write("UtilName = " & Session("utilname") & "<BR>")
		Response.Write("UtilID = " & Session("utilid") & "<BR>")
		Response.Write("Action = " & Session("action") & "<BR>")
		Response.End()
	End If

	' Pass required info to the DLL
	objCrossTab.CrossTabID = Session("utilid")

	fok = True
	blnNoDefinition = True

	Dim aPrompts
	Dim fModuleOk As Boolean

	Dim strEmailGroupName As String = ""
	If plngOutputEmailID > 0 Then strEmailGroupName = objCrossTab.GetEmailGroupName(plngOutputEmailID)
	
	aPrompts = Session("Prompts_" & Session("utiltype") & "_" & Session("utilid"))

	fModuleOk = True
	If lngStartDateColID = 0 Or _
		lngStartSessionColID = 0 Or _
		lngEndDateColID = 0 Or _
		lngEndSessionColID = 0 Or _
		lngTypeColID = 0 Or _
		lngReasonColID = 0 Or _
		lngDurationColID = 0 Then
		
		fok = False
		fModuleOk = False
	End If
	
	If fok Then
		fok = objCrossTab.SetPromptedValues(aPrompts)
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then
		fok = objCrossTab.SetAbsenceBreakDownDisplayOptions(bPrintFilterPickList)
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then
		fok = objCrossTab.SetAbsenceBreakDownDisplayOptions(bPrintFilterPickList)
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then
		fok = objCrossTab.AbsenceBreakdownRetreiveDefinition(dtStartDate, dtEndDate, lngHorColID, lngVerColID, lngPicklistID, lngFilterID, lngPersonnelID, strAbsenceTypes)
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then
		blnNoDefinition = False
		lngEventLogID = objCrossTab.EventLogAddHeader
		fok = (lngEventLogID > 0)
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then
		fok = objCrossTab.UDFFunctions(True)
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If
	
	If fok Then
		fok = objCrossTab.CreateTempTable
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then
		fok = objCrossTab.UDFFunctions(False)
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then
		fok = objCrossTab.AbsenceBreakdownRunStoredProcedure
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then
		fok = objCrossTab.CreatePivotDataset
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If
	
	If fok Then
		fok = objCrossTab.AbsenceBreakdownGetHeadingsAndSearches
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then
		fok = objCrossTab.BuildTypeArray
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then
		fok = objCrossTab.AbsenceBreakdownBuildDataArrays
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	objCrossTab.ClearUp()
	
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

			<form action="util_run_crosstab_downloadoutput" method="post" id="frmExportData" name="frmExportData" target="submit-iframe">
				<input type="hidden" id="txtPreview" name="txtPreview" value="<%=bOutputPreview%>">	
				<input type="hidden" id="txtFormat" name="txtFormat" value="<%=lngOutputFormat%>">
				<input type="hidden" id="txtScreen" name="txtScreen" value="<%=pblnOutputScreen%>">
				<input type="hidden" id="txtPrinter" name="txtPrinter" value="">
				<input type="hidden" id="txtPrinterName" name="txtPrinterName" value="">
				<input type="hidden" id="txtSave" name="txtSave" value="<%=pblnOutputSave%>">
				<input type="hidden" id="txtSaveExisting" name="txtSaveExisting" value="<%=plngOutputSaveExisting%>">
				<input type="hidden" id="txtEmail" name="txtEmail" value="<%=pblnOutputEmail%>">
				<input type="hidden" id="txtEmailAddr" name="txtEmailAddr" value="<%=plngOutputEmailID%>">
				<input type="hidden" id="txtEmailAddrName" name="txtEmailAddrName" value="<%=strEmailGroupName%>">
				<input type="hidden" id="txtEmailSubject" name="txtEmailSubject" value="<%=pstrOutputEmailSubject%>">
				<input type="hidden" id="txtEmailAttachAs" name="txtEmailAttachAs" value="<%=pstrOutputEmailAttachAs%>">
				<input type="hidden" id="txtEmailGroupAddr" name="txtEmailGroupAddr" value="">
				<input type="hidden" id="txtEmailGroupID" name="txtEmailGroupID" value="<%=plngOutputEmailID%>">
				<input type="hidden" id="txtFileName" name="txtFileName" value="<%=pstrOutputFilename%>">
				<input type="hidden" id="txtUtilType" name="txtUtilType" value="<%=session("utilType")%>">
				<input type="hidden" id="txtUtilID" name="txtUtilID" value="<%=Session("utilID")%>">
				<input type="hidden" id="txtMode" name="txtMode">
				<input type="hidden" id="download_token_value_id" name="download_token_value_id"/>
				<%=Html.AntiForgeryToken()%>
			</form>

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

<input type='hidden' id="txtNoRecs" name="txtNoRecs" value="<%=objCrossTab.NoRecords%>">

<%If Not objCrossTab.NoRecords Then%>
<script type="text/javascript">
	util_run_crosstabs_window_onload();
</script>
<%End If%>