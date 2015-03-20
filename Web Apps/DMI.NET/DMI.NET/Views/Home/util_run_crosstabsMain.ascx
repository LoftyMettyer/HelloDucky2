<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@Import namespace="DMI.NET" %>

<script src="<%: Url.LatestContent("~/bundles/utilities_crosstabs")%>" type="text/javascript"></script>

<%
	Dim fok As Boolean
	Dim objCrossTab As CrossTab
	Dim fNotCancelled As Boolean
	Dim lngEventLogID As Long
	Dim aPrompts As Object
		
	Session("objCrossTab" & Session("UtilID")) = Nothing
	Session("CT_Mode") = ""
	Session("CT_PageNumber") = 0
	Session("CT_IntersectionType") = ""
	Session("CT_ShowPercentage") = False
	Session("CT_PercentageOfPage") = False
	Session("CT_SupressZeros") = False
	Session("CT_Use1000") = False

	If Session("utiltype") Is Nothing Or _
		 Session("utilname") Is Nothing Or _
		 Session("utilid") Is Nothing Or _
		 Session("action") Is Nothing Then
				
		Response.Write("Error : Not all session variables found...<HR>")
		Response.Write("Type = " & Session("utiltype") & "<BR>")
		Response.Write("UtilName = " & Session("utilname") & "<BR>")
		Response.Write("UtilID = " & Session("utilid") & "<BR>")
		Response.Write("Action = " & Session("action") & "<BR>")
	End If

	' Create the reference to the DLL (Report Class)
	objCrossTab = New HR.Intranet.Server.CrossTab
	objCrossTab.SessionInfo = CType(Session("SessionContext"), SessionInfo)
	
	Session("objCrossTab" & Session("UtilID")) = Nothing

	' Pass required info to the DLL
	objCrossTab.CrossTabID = Session("utilid")

	fok = True

	aPrompts = Session("Prompts_" & Session("utiltype") & "_" & Session("utilid"))
	If fok Then
		fok = objCrossTab.SetPromptedValues(aPrompts)
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then
		fok = objCrossTab.RetreiveDefinition
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then
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
		fok = objCrossTab.CreatePivotDataset
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then
		fok = objCrossTab.GetHeadingsAndSearches
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then
		fok = objCrossTab.BuildTypeArray
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then
		fok = objCrossTab.BuildDataArrays
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	objCrossTab.ClearUp()
	
	Session("objCrossTab" & Session("UtilID")) = objCrossTab

%>

<script type="text/javascript">
	function crosstab_loadAddRecords() {

		var iCount;
		iCount = new Number(txtLoadCount.value);
		txtLoadCount.value = iCount + 1;
		if (iCount > 0) {
			var frmGetData = OpenHR.getForm("reportdataframe", "frmGetReportData");
			<% Response.Write("frmGetData.txtUtilID.value = """ & Session("utilid") & """;" & vbCrLf)%>
			getCrossTabData("LOAD", 0, 0, 0, 0, 0, 0);
		}
	}
</script>

<input type='hidden' id="txtLoadCount" name="txtLoadCount" value="0">
<input type='hidden' id="txtNoRecs" name="txtNoRecs" value="<%=objCrossTab.NoRecords%>">

<div id="reportworkframe" data-framesource="util_run_crosstabs" style="display: block">
	<%Html.RenderPartial("~/views/home/util_run_crosstabs.ascx")%>
		
	<form action="<%= IIf(objCrossTab.CrossTabType = CrossTabType.ctt9GridBox, "util_run_nineboxgrid_downloadoutput", "util_run_crosstab_downloadoutput")%>" method="post" id="frmExportData" name="frmExportData" target="submit-iframe">
		<input type="hidden" id="txtPreview" name="txtPreview" value="<%=objCrossTab.OutputPreview%>">	
		<input type="hidden" id="txtFormat" name="txtFormat" value="<%=objCrossTab.OutputFormat%>">
		<input type="hidden" id="txtScreen" name="txtScreen" value="<%=objCrossTab.OutputScreen%>">
		<input type="hidden" id="txtPrinter" name="txtPrinter" value="<%=objCrossTab.OutputPrinter%>">
		<input type="hidden" id="txtPrinterName" name="txtPrinterName" value="<%=objCrossTab.OutputPrinterName%>">
		<input type="hidden" id="txtSave" name="txtSave" value="<%=objCrossTab.OutputSave%>">
		<input type="hidden" id="txtSaveExisting" name="txtSaveExisting" value="<%=objCrossTab.OutputSaveExisting%>">
		<input type="hidden" id="txtEmail" name="txtEmail" value="<%=objCrossTab.OutputEmail%>">
		<input type="hidden" id="txtEmailAddr" name="txtEmailAddr" value="<%=objCrossTab.OutputEmailID%>">
		<input type="hidden" id="txtEmailAddrName" name="txtEmailAddrName" value="<%=objCrossTab.OutputEmailGroupName%>">
		<input type="hidden" id="txtEmailSubject" name="txtEmailSubject" value="<%=objCrossTab.OutputEmailSubject%>">
		<input type="hidden" id="txtEmailAttachAs" name="txtEmailAttachAs" value="<%=objCrossTab.OutputEmailAttachAs%>">
		<input type="hidden" id="txtEmailGroupAddr" name="txtEmailGroupAddr" value="">
		<input type="hidden" id="txtEmailGroupID" name="txtEmailGroupID" value="<%=objCrossTab.OutputEmailID%>">
		<input type="hidden" id="txtFileName" name="txtFileName" value="<%=objCrossTab.OutputFilename%>">
		<input type="hidden" id="txtUtilType" name="txtUtilType" value="<%=session("utilType")%>">
		<input type="hidden" id="txtUtilID" name="txtUtilID" value="<%=Session("utilID")%>">
		<input type="hidden" id="download_token_value_id" name="download_token_value_id"/>
		<%=Html.AntiForgeryToken()%>
	</form>

</div>

<div id="reportdataframe" data-framesource="util_run_crosstabsData" style="display: none;" accesskey="">
	<%Html.RenderPartial("~/views/home/util_run_crosstabsData.ascx")%>
</div>

<div id="reportbreakdownframe" data-framesource="util_run_crosstabsBreakdown" style="display: none;height: 100%;" accesskey="">
	<%Html.RenderPartial("~/views/home/util_run_crosstabsBreakdown.ascx")%>
</div>

<div id="outputoptions" data-framesource="util_run_outputoptions" style="display: none;">
	<%	Html.RenderPartial("~/Views/Home/util_run_outputoptions.ascx")%>
</div>

<form id="frmOutput" name="frmOutput">
	<input type="hidden" id="fok" name="fok" value="">
	<input type="hidden" id="cancelled" name="cancelled" value="">
	<input type="hidden" id="statusmessage" name="statusmessage" value="">
</form>



<%If Not objCrossTab.NoRecords Then%>
<script type="text/javascript">
	$(document).on('keydown', function (e) {
		if (e.currentTarget.activeElement.type == "checkbox") {
			if (e.keyCode == 32) {
				if (!/Firefox[\/\s](\d+\.\d+)/.test(navigator.userAgent)) { //test for Firefox
					e.currentTarget.activeElement.click();
				}
			}
		}
	});

	util_run_crosstabs_window_onload();

	$("#reportframe").show();
	$("#top").hide();
	$("#reportworkframe").show();
</script>
<%End If%>