<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="System.Globalization" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="HR.Intranet.Server.Structures" %>

<script src="<%: Url.LatestContent("~/bundles/utilities_customreports")%>" type="text/javascript"></script>
<script type="text/javascript">
	//$('#main').css('overflow', 'auto');
</script>
<% 
	Dim bBradfordFactor As Boolean
		
	bBradfordFactor = (Session("utiltype") = "16")

	Dim objReport As Report
	
	Dim fok As Boolean
	Dim fNotCancelled As Boolean

	Dim dtStartDate As Date
	Dim dtEndDate As Date
				Dim strAbsenceTypes As String = ""
				Dim lngFilterID As Long
				Dim lngPicklistID As Long
				Dim lngPersonnelID As Long
	
				Dim bBradford_SRV As Boolean
				Dim bBradford_ShowDurations As Boolean
				Dim bBradford_ShowInstances As Boolean
				Dim bBradford_ShowFormula As Boolean
				Dim bBradford_OmitBeforeStart As Boolean
				Dim bBradford_OmitAfterEnd As Boolean
				Dim bBradford_txtOrderBy1 As String
				Dim lngBradford_txtOrderBy1ID As String
				Dim bBradford_txtOrderBy1Asc As Boolean
				Dim bBradford_txtOrderBy2 As String
				Dim lngBradford_txtOrderBy2ID As String
				Dim bBradford_txtOrderBy2Asc As Boolean
				Dim bPrintFilterPickList As Boolean

	' Default output options
	             	 Dim bOutputPreview As Boolean
	             	 Dim lngOutputFormat As Long
	             	 Dim pblnOutputScreen As Boolean
	             	 Dim pblnOutputSave As Boolean
	             	 Dim pblnOutputEmail As Boolean
	Dim plngOutputEmailID As Integer
	             	 Dim pstrOutputEmailName As String
	             	 Dim pstrOutputEmailSubject As String
	             	 Dim pstrOutputEmailAttachAs As String
	             	 Dim pstrOutputFilename As String

	             	 Dim bMinBradford As Boolean
	             	 Dim lngMinBradfordAmount As Long
	             	 Dim pbDisplayBradfordDetail As Boolean
				
	             	 fok = True
	             	 fNotCancelled = True

	             	 ' Create the reference to the DLL (Report Class)
	objReport = New Report
	             	 objReport.SessionInfo = CType(Session("SessionContext"), SessionInfo)

				
	             	 ' Pass required info to the DLL			
	objReport.CustomReportID = Session("utilid")
	objReport.ClientDateFormat = Session("LocaleDateFormat")
	objReport.LocalDecimalSeparator = Session("LocaleDecimalSeparator")

	objReport.Name = "Bradford Factor"
	objReport.OutputFormat = Session("stdReport_OutputFormat")
	objReport.OutputPreview = Session("stdReport_OutputPreview")
	objReport.OutputFilename = Session("stdReport_OutputFilename")
	
	If fok And bBradfordFactor Then
		
		dtStartDate = DateTime.ParseExact(Session("stdReport_StartDate").ToString(), "MM/dd/yyyy", CultureInfo.InvariantCulture)
		dtEndDate = DateTime.ParseExact(Session("stdReport_EndDate").ToString(), "MM/dd/yyyy", CultureInfo.InvariantCulture)
						
		strAbsenceTypes = Session("stdReport_AbsenceTypes")
		lngFilterID = Session("stdReport_FilterID")
		lngPicklistID = Session("stdReport_PicklistID")
		lngPersonnelID = Session("optionRecordID")

		bBradford_SRV = Session("stdReport_Bradford_SRV")
		bBradford_ShowDurations = Session("stdReport_Bradford_ShowDurations")
		bBradford_ShowInstances = Session("stdReport_Bradford_ShowInstances")
		bBradford_ShowFormula = Session("stdReport_Bradford_ShowFormula")
		bBradford_OmitBeforeStart = Session("stdReport_Bradford_OmitBeforeStart")
		bBradford_OmitAfterEnd = Session("stdReport_Bradford_OmitAfterEnd")
		bBradford_txtOrderBy1 = Session("stdReport_Bradford_txtOrderBy1")
		lngBradford_txtOrderBy1ID = CLng(Session("stdReport_Bradford_txtOrderBy1ID"))
		bBradford_txtOrderBy1Asc = Session("stdReport_Bradford_txtOrderBy1Asc")
		bBradford_txtOrderBy2 = Session("stdReport_Bradford_txtOrderBy2")
		lngBradford_txtOrderBy2ID = CLng(Session("stdReport_Bradford_txtOrderBy2ID"))
		bBradford_txtOrderBy2Asc = Session("stdReport_Bradford_txtOrderBy2Asc")
		bPrintFilterPickList = Session("stdReport_PrintFilterPicklistHeader")

		bMinBradford = Session("stdReport_MinimumBradfordFactor")
		lngMinBradfordAmount = CLng(Session("stdReport_MinimumBradfordFactorAmount"))
		pbDisplayBradfordDetail = Session("stdReport_DisplayBradfordDetail")

		' Default output options
		bOutputPreview = objReport.OutputPreview
		lngOutputFormat = Session("stdReport_OutputFormat")
		pblnOutputScreen = False
		pblnOutputSave = Session("stdReport_OutputSave")
		pblnOutputEmail = Session("stdReport_OutputEmail")
		plngOutputEmailID = Session("stdReport_OutputEmailAddr")
		pstrOutputEmailName = Session("stdReport_OutputEmailName")
		pstrOutputEmailSubject = Session("stdReport_OutputEmailSubject")
		pstrOutputEmailAttachAs = Session("stdReport_OutputEmailAttachAs")
		pstrOutputFilename = objReport.OutputFilename
	End If

	
	Dim strEmailGroupName As String = ""
	If plngOutputEmailID > 0 Then strEmailGroupName = objReport.GetEmailGroupName(plngOutputEmailID)

	
	             	 If fok And Not bBradfordFactor Then
	             		 fok = objReport.GetCustomReportDefinition
		Session("utilname") = objReport.Name
	             		 fNotCancelled = Response.IsClientConnected
	             		 If fok Then fok = fNotCancelled
	             	 End If

	             	 If fok And Not bBradfordFactor Then
	             		 fok = objReport.GetDetailsRecordsets
	             		 fNotCancelled = Response.IsClientConnected
	             		 If fok Then fok = fNotCancelled
	             	 End If

	             	 If fok And bBradfordFactor Then
	             		 fok = objReport.SetBradfordDisplayOptions(bBradford_SRV, bBradford_ShowDurations, bBradford_ShowInstances, bBradford_ShowFormula, bPrintFilterPickList, pbDisplayBradfordDetail)

	             		 If lngPersonnelID = 0 Then
	             			 fok = objReport.SetBradfordOrders(bBradford_txtOrderBy1, bBradford_txtOrderBy2, bBradford_txtOrderBy1Asc, bBradford_txtOrderBy2Asc, lngBradford_txtOrderBy1ID, lngBradford_txtOrderBy2ID)
	             		 Else
	             			 fok = objReport.SetBradfordOrders("None", "None", False, False, 0, 0)
	             		 End If

	             		 fok = objReport.SetBradfordIncludeOptions(bBradford_OmitBeforeStart, bBradford_OmitAfterEnd, lngPersonnelID, lngFilterID, lngPicklistID, bMinBradford, lngMinBradfordAmount)
	             		 fNotCancelled = Response.IsClientConnected
	             		 If fok Then fok = fNotCancelled
	             	 End If

	             	 If fok And bBradfordFactor Then
	             		 fok = objReport.GetBradfordReportDefinition(dtStartDate, dtEndDate)
	             		 fNotCancelled = Response.IsClientConnected
	             		 If fok Then fok = fNotCancelled
	             	 End If

	             	 If fok And bBradfordFactor Then
	             		 fok = objReport.GetBradfordRecordSet
	             		 fNotCancelled = Response.IsClientConnected
	             		 If fok Then fok = fNotCancelled
	             	 End If

	             	 Dim aPrompts
				
	aPrompts = Session("Prompts_" & Session("utiltype") & "_" & Session("utilid"))
	             	 If fok Then
	             		 fok = objReport.SetPromptedValues(aPrompts)
	             		 fNotCancelled = Response.IsClientConnected
	             		 If fok Then fok = fNotCancelled
	             	 End If

	             	 If fok Then
	             		 fok = objReport.GenerateSQL
	             		 fNotCancelled = Response.IsClientConnected
	             		 If fok Then fok = fNotCancelled
	             	 End If

	             	 If fok And bBradfordFactor Then
	             		 fok = objReport.GenerateSQLBradford(strAbsenceTypes)
	             		 fNotCancelled = Response.IsClientConnected
	             		 If fok Then fok = fNotCancelled
	             	 End If

	             	 If fok Then
	             		 fok = objReport.AddTempTableToSQL
	             		 fNotCancelled = Response.IsClientConnected
	             		 If fok Then fok = fNotCancelled
	             	 End If

	             	 If fok Then
	             		 fok = objReport.MergeSQLStrings
	             		 fNotCancelled = Response.IsClientConnected
	             		 If fok Then fok = fNotCancelled
	             	 End If

	             	 If fok Then
	             		 fok = objReport.UDFFunctions(True)
	             		 fNotCancelled = Response.IsClientConnected
	             		 If fok Then fok = fNotCancelled
	             	 End If

	             	 If fok Then
	             		 fok = objReport.ExecuteSql
	             		 fNotCancelled = Response.IsClientConnected
	             		 If fok Then fok = fNotCancelled
	             	 End If

	             	 If fok Then
	             		 fok = objReport.UDFFunctions(False)
	             		 fNotCancelled = Response.IsClientConnected
	             		 If fok Then fok = fNotCancelled
	             	 End If

	             	 If fok And bBradfordFactor Then
		fok = objReport.CalculateBradfordFactors()
	             		 fNotCancelled = Response.IsClientConnected
	             		 If fok Then fok = fNotCancelled
	             	 End If

	             	 If fok And objReport.ChildCount > 1 And objReport.UsedChildCount > 1 Then
	             		 fok = objReport.CreateMutipleChildTempTable
	             		 fNotCancelled = Response.IsClientConnected
	             		 If fok Then fok = fNotCancelled
	             	 End If

	             	 If fok Then
	             		 fok = objReport.CheckRecordSet
	             		 fNotCancelled = Response.IsClientConnected
	             		 If fok Then fok = fNotCancelled
	             	 End If
		
	             	 If fok Then

	             		 If fok Then
	             			 fok = objReport.PopulateGrid_LoadRecords
	             			 fNotCancelled = Response.IsClientConnected
	             			 If fok Then fok = fNotCancelled
	             		 End If
		
	             		 If fok Then
	             			 fok = objReport.PopulateGrid_HideColumns
	             			 fNotCancelled = Response.IsClientConnected
	             			 If fok Then fok = fNotCancelled
	             		 End If
				
	             	 End If

	             	 Session("CustomReport") = objReport
	
	gridReportData.DataSource = objReport.ReportDataTable
	             	 gridReportData.DataBind()
		
		Dim fNoRecords As Boolean
		
		fNoRecords = objReport.NoRecords

		If fok Then
			If Response.IsClientConnected Then
				objReport.Cancelled = False
			Else
				objReport.Cancelled = True
			End If
		Else
			If Not fNoRecords Then
				If fNotCancelled Then
					objReport.FailedMessage = objReport.ErrorString
					objReport.Failed = True
				Else
					objReport.Cancelled = True
				End If
			End If
		End If
	
	If fok Then
		objReport.ClearUp()
	End If

	If fok Then
		Response.Write("<form name=frmOutput id=frmOutput method=post>" & vbCrLf)
		Response.Write("<div>")
		Response.Write("			<table name=tblGrid id=tblGrid height=100% width=100% class=""invisible"" cellspacing=0 cellpadding=0>" & vbCrLf)
		Response.Write("				<tr>" & vbCrLf)
		Response.Write("					<td ALIGN=center colspan=12 NAME='tdOutputMSG' ID='tdOutputMSG'>" & vbCrLf)
%>

<form id="formReportData" runat="server">
	<asp:GridView ID="gridReportData" runat="server"
		AllowPaging="False"
		GridLines="None"
		CssClass="visibletablecolumn"
		ClientIDMode="Static">
		<Columns>
			<asp:BoundField DataField="rowtype" HeaderText="rowType" />
		</Columns>
	</asp:GridView>
</form>

<%
	Response.Write("					</td>" & vbCrLf)
	Response.Write("				</tr>" & vbCrLf)
	Response.Write("				<tr>" & vbCrLf)
	Response.Write("					<td colspan=12 height=10></td>" & vbCrLf)
	Response.Write("				</tr>" & vbCrLf)

	Response.Write("				<tr height=25>" & vbCrLf)
	Response.Write("					<td width=20></td>" & vbCrLf)
	Response.Write("					<td colspan=8>" & vbCrLf)
	Response.Write("            <div>")
	Response.Write("						<table WIDTH=""100%"" class=""invisible"" CELLSPACING=0 CELLPADDING=0>" & vbCrLf)
	Response.Write("							<tr>" & vbCrLf)
	Response.Write("								<td>" & vbCrLf)
	Response.Write("								</td>" & vbCrLf)
	Response.Write("								<td>&nbsp;</td>" & vbCrLf)
	Response.Write("								<td width=20>" & vbCrLf)
	Response.Write("								</td>" & vbCrLf)
	Response.Write("							</tr>" & vbCrLf)
	Response.Write("						</table>" & vbCrLf)
	Response.Write("</div>")
	Response.Write("					</td>" & vbCrLf)
	Response.Write("					<td width=10></td>" & vbCrLf)
	Response.Write("					<td width=80> " & vbCrLf)
	Response.Write("					</td>" & vbCrLf)
	Response.Write("					<td width=20></td>" & vbCrLf)
	Response.Write("				</tr>" & vbCrLf)
	Response.Write("				<tr>" & vbCrLf)
	Response.Write("					<td colspan=12 height=10></td>" & vbCrLf)
	Response.Write("				</tr>" & vbCrLf)
	Response.Write("			</table>" & vbCrLf)
	Response.Write("      </div>")
	Response.Write("</form>" & vbCrLf)

End If
%>

<input type='hidden' id="txtDefn_Name" name="txtDefn_Name" value="<%:objReport.ReportCaption.ToString()%>">
<input type='hidden' id="txtDefn_ErrMsg" name="txtDefn_ErrMsg" value="<%:objReport.ErrorString%>">
<input type='hidden' id="txtNoRecs" name="txtNoRecs" value="<%:objReport.NoRecords%>">


		<script runat="server">

			Protected Sub gridReportData_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gridReportData.RowDataBound
		
				Dim objReport As Report = CType(Session("CustomReport"), Report)
				Dim objThisColumn As ReportDetailItem

				If e.Row.RowType = DataControlRowType.Header Or e.Row.RowType = DataControlRowType.Footer Then
					e.Row.CssClass = "header"
			
				Else

					If e.Row.Cells(0).Text = HR.Intranet.Server.Enums.RowType.GrandSummary Then
						e.Row.CssClass = "grandsummaryrow"
					
					ElseIf Not e.Row.Cells(0).Text = HR.Intranet.Server.Enums.RowType.Data Then
						e.Row.CssClass = "summarytablerow"
					
					End If
							
				End If

				For iCount = 1 To objReport.ReportDataTable.Columns.Count - 1
						
					objThisColumn = objReport.DisplayColumns(iCount)
					
					If objThisColumn.IsNumeric Then
						e.Row.Cells(iCount).HorizontalAlign = HorizontalAlign.Right
					Else
						e.Row.Cells(iCount).HorizontalAlign = HorizontalAlign.Left
					End If
	
				Next

			End Sub

			
		</script>
		
<form action="util_run_customreport_downloadoutput" method="post" id="frmExportData" name="frmExportData" target="submit-iframe">
	<input type="hidden" id="txtPreview" name="txtPreview" value="<%=bOutputPreview%>">
	<input type="hidden" id="txtFormat" name="txtFormat" value="<%=lngOutputFormat%>">
	<input type="hidden" id="txtScreen" name="txtScreen" value="<%=pblnOutputScreen%>">
	<input type="hidden" id="txtPrinter" name="txtPrinter" value="">
	<input type="hidden" id="txtPrinterName" name="txtPrinterName" value="">
	<input type="hidden" id="txtSave" name="txtSave" value="<%=pblnOutputSave%>">
	<input type="hidden" id="txtSaveExisting" name="txtSaveExisting" value="0">
	<input type="hidden" id="txtEmail" name="txtEmail" value="<%=pblnOutputEmail%>">
	<input type="hidden" id="txtEmailAddr" name="txtEmailAddr" value="<%=plngOutputEmailID%>">
	<input type="hidden" id="txtEmailAddrName" name="txtEmailAddrName" value="<%=strEmailGroupName%>">
	<input type="hidden" id="txtEmailSubject" name="txtEmailSubject" value="<%=pstrOutputEmailSubject%>">
	<input type="hidden" id="txtEmailAttachAs" name="txtEmailAttachAs" value="<%=pstrOutputEmailAttachAs%>">
	<input type="hidden" id="txtEmailGroupAddr" name="txtEmailGroupAddr" value="">
	<input type="hidden" id="txtFileName" name="txtFileName" value="<%=pstrOutputFilename%>">
	<input type="hidden" id="txtEmailGroupID" name="txtEmailGroupID" value=<%=plngOutputEmailID%>>
	<input type="hidden" id="txtUtilType" name="txtUtilType" value="<%=session("utilType")%>">
	<input type="hidden" id="txtUtilID" name="txtUtilID" value="<%=Session("utilID")%>">
	<input type="hidden" id="txtMode" name="txtMode">
	<input type="hidden" id="download_token_value_id" name="download_token_value_id"/>
	<%=Html.AntiForgeryToken()%>
</form>


<script type="text/javascript">
	var size = {
		MakeWidth: $('#divUtilRunForm').width(),
		MakeHeight: $('#reportworkframe').height()
	};

	//Shrink to fit, or set to 100px per column?
	var ShrinkToFit = true;
	var gridWidth;
	var gridHeight;

	//Get count of visible columns
	if (menu_isSSIMode()) {
		try {
			gridWidth = $('#reportworkframe').width();
			gridHeight = $('#reportworkframe').height() - 100;
		} catch (e) {
			gridWidth = 'auto';
			gridHeight = 'auto';
		}
		ShrinkToFit = true;
	} else {
		//DMI options.
		var iVisibleCount = 13;
		if (iVisibleCount < 8) ShrinkToFit = true;
		gridWidth = size.MakeWidth;
		gridHeight = size.MakeHeight;
	}


	tableToGrid("#gridReportData", {
		shrinkToFit: ShrinkToFit,
		width: gridWidth,
		height: gridHeight,
		ignoreCase: true,
		cmTemplate: { sortable: false },
		rowNum: 200000,
		loadComplete: function () {
			$('#gridReportData').hideCol("rowType");
			stylejqGrid();
			$('#gridReportData').setGridWidth($('#main').width());
		}
	});

	$('#gview_gridReportData td').css('white-space', 'pre-line');

	function stylejqGrid() {
		//jqGrid style overrides
		$('#gview_gridReportData tr.jqgrow td').css('vertical-align', 'top'); //float text to top, in case of multi-line cells
		$('#gview_gridReportData tr.footrow td').css('vertical-align', 'top'); //float text to top, in case of multi-line footers
		$('#gview_gridReportData .s-ico span').css('display', 'none'); //hide the sort order icons - they don't tie in to the dataview model.
	};
	
</script>
