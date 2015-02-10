<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="HR.Intranet.Server.Structures" %>

<script src="<%: Url.LatestContent("~/bundles/utilities_customreports")%>" type="text/javascript"></script>

<% 
	Dim bBradfordFactor As Boolean
	
	bBradfordFactor = (Session("utiltype") = "16")

	Dim objReport As Report
	
	Dim fok As Boolean
	Dim fNotCancelled As Boolean

	Dim dtStartDate As String
	Dim dtEndDate As String
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

	If fok Then
		fok = objReport.GetCustomReportDefinition
		Session("utilname") = objReport.Name
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then
		fok = objReport.GetDetailsRecordsets
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
	
	' Bind the data to the grid if atleast one non-hidden column available
	If ((objReport.ReportDataTable IsNot Nothing) AndAlso (objReport.ReportDataTable.Columns.Count > 1)) Then
		gridReportData.DataSource = objReport.ReportDataTable
		gridReportData.DataBind()
	Else
		Response.Write("No output generated. Check your data.")
	End If
	
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
	
	Response.Write("<input type=hidden id=txtSuccessFlag name=txtSuccessFlag value=2>" & vbCrLf)
Else%>

<input type="hidden" id="txtSuccessFlag" name="txtSuccessFlag" value="3">
<%
End If
%>

<input type='hidden' id="txtNoRecs" name="txtNoRecs" value="<%=objReport.NoRecords%>">

<form id="frmOriginalDefinition" style="visibility: hidden; display: none">
	<%
		Response.Write("	<input type='hidden' id='txtDefn_Name' name='txtDefn_Name' value='" & objReport.ReportCaption.ToString() & "'>" & vbCrLf)
		Response.Write("	<input type='hidden' id=txtDefn_ErrMsg name=txtDefn_ErrMsg value=""" & objReport.ErrorString & """>" & vbCrLf)
	%>
	<input type="hidden" id="txtUserName" name="txtUserName" value="<%=Session("username").ToString()%>">
	<input type="hidden" id="txtDateFormat" name="txtDateFormat" value="<%=Session("LocaleDateFormat").ToString()%>">

	<input type="hidden" id="txtCancelPrint" name="txtCancelPrint">
	<input type="hidden" id="txtOptionsDone" name="txtOptionsDone">
	<input type="hidden" id="txtOptionsPortrait" name="txtOptionsPortrait">
	<input type="hidden" id="txtOptionsMarginLeft" name="txtOptionsMarginLeft">
	<input type="hidden" id="txtOptionsMarginRight" name="txtOptionsMarginRight">
	<input type="hidden" id="txtOptionsMarginTop" name="txtOptionsMarginTop">
	<input type="hidden" id="txtOptionsMarginBottom" name="txtOptionsMarginBottom">
	<input type="hidden" id="txtOptionsCopies" name="txtOptionsCopies">
</form>

<script runat="server">

	Protected Sub gridReportData_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gridReportData.RowDataBound		
		
		Dim objReport As Report = CType(Session("CustomReport"), Report)

		If e.Row.RowType = DataControlRowType.Header Or e.Row.RowType = DataControlRowType.Footer Then
			e.Row.CssClass = "header"			
			For iCount = 1 To objReport.ReportDataTable.Columns.Count - 1
				e.Row.Cells(iCount).Text = e.Row.Cells(iCount).Text.Replace(" ", "_").Replace("&quot;", "_")
				
			Next
		Else

			If e.Row.Cells(0).Text = HR.Intranet.Server.Enums.RowType.GrandSummary Then
				e.Row.CssClass = "grandsummaryrow"
					
			ElseIf Not e.Row.Cells(0).Text = HR.Intranet.Server.Enums.RowType.Data Then
				e.Row.CssClass = "summarytablerow"
					
		End If
							
		End If

	End Sub	
	
</script>

<form action="util_run_customreport_downloadoutput" method="post" id="frmExportData" name="frmExportData" target="submit-iframe">
	<input type="hidden" id="txtPreview" name="txtPreview" value="<%=objReport.OutputPreview%>">
	<input type="hidden" id="txtFormat" name="txtFormat" value="<%=objReport.OutputFormat%>">
	<input type="hidden" id="txtScreen" name="txtScreen" value="<%=objReport.OutputScreen%>">
	<input type="hidden" id="txtPrinter" name="txtPrinter" value="<%=objReport.OutputPrinter%>">
	<input type="hidden" id="txtPrinterName" name="txtPrinterName" value="<%=objReport.OutputPrinterName%>">
	<input type="hidden" id="txtSave" name="txtSave" value="<%=objReport.OutputSave%>">
	<input type="hidden" id="txtSaveExisting" name="txtSaveExisting" value="<%=objReport.OutputSaveExisting%>">
	<input type="hidden" id="txtEmail" name="txtEmail" value="<%=objReport.OutputEmail%>">
	<input type="hidden" id="txtEmailAddr" name="txtEmailAddr" value="<%=objReport.OutputEmailID%>">
	<input type="hidden" id="txtEmailAddrName" name="txtEmailAddrName" value="<%=objReport.OutputEmailGroupName%>">
	<input type="hidden" id="txtEmailSubject" name="txtEmailSubject" value="<%=objReport.OutputEmailSubject%>">
	<input type="hidden" id="txtEmailAttachAs" name="txtEmailAttachAs" value="<%=objReport.OutputEmailAttachAs%>">
	<input type="hidden" id="txtEmailGroupAddr" name="txtEmailGroupAddr" value="">
	<input type="hidden" id="txtFileName" name="txtFileName" value="<%=objReport.OutputFilename%>">
	<input type="hidden" id="txtEmailGroupID" name="txtEmailGroupID" value="<%=objReport.OutputEmailID%>">
	<input type="hidden" id="txtUtilType" name="txtUtilType" value="<%=session("utilType")%>">
	<input type="hidden" id="txtUtilID" name="txtUtilID" value="<%=Session("utilID")%>">
	<input type="hidden" id="download_token_value_id" name="download_token_value_id"/>
	<%=Html.AntiForgeryToken()%>
</form>

<script type="text/javascript">
	//Shrink to fit, or set to 100px per column?
	var ShrinkToFit = false;
	var gridWidth;
	var gridHeight;
	// first get the size from the window
	// if that didn't work, get it from the body
	var size = {
		MakeWidth: $('#divUtilRunForm').width(),
		MakeHeight: $('#reportworkframe').height()
	};
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
		
		var iVisibleCount = Number("<%:objReport.DisplayColumns.Count%>");
		if ((iVisibleCount *100) < size.MakeWidth) ShrinkToFit = true;
		gridWidth = (size.MakeWidth);
		gridHeight = (size.MakeHeight);
	}

		var newFormat = OpenHR.getLocaleDateString();
		var srcFormat = newFormat;	
		
		tableToGrid("#gridReportData", {
			shrinkToFit: ShrinkToFit,
			width: gridWidth,
			height: gridHeight,
			ignoreCase: true,
			colNames: [
				<%Dim iColCount As Integer = 0
		For Each objItem In objReport.DisplayColumns
			Dim sColumnName = objReport.ReportDataTable.Columns(iColCount).ColumnName
			Response.Write(String.Format("{0}'{1}'", IIf(iColCount > 0, ", ", ""), sColumnName))
			iColCount += 1
		Next%>
			],
			colModel: [
				<%
	iColCount = 0
		For Each objItem In objReport.DisplayColumns
			Dim sColumnName = objReport.ReportDataTable.Columns(iColCount).ColumnName.Replace(" ", "_").Replace("""", "_")
			Dim iColumnWidth As Integer = 100
			If objItem.IsNumeric Then
				Response.Write(String.Format("{0}{{name:'", IIf(iColCount > 0, ", ", "")) & sColumnName & "',align:'right', width: '" & iColumnWidth.ToString() & "'}")
			ElseIf objItem.IsDateColumn Then
				Response.Write(String.Format("{0}{{name:'", IIf(iColCount > 0, ", ", "")) & sColumnName & "', edittype: 'date', align: 'center',  formatter: 'date', formatoptions: { srcformat: srcFormat, newformat: newFormat, disabled: true, width: '" & iColumnWidth.ToString() & "' }}")
			Else
				Response.Write(String.Format("{0}{{name:'", IIf(iColCount > 0, ", ", "")) & sColumnName & "', width: '" & iColumnWidth.ToString() & "'}")
			End If
			iColCount += 1
		Next
	%>
			],
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
		}
		
	if (menu_isSSIMode()) $('#gbox_gridReportData').css('margin', '0 auto'); //center the report in self-service screen.
</script>

