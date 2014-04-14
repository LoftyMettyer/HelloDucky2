﻿<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="HR.Intranet.Server.Structures" %>

<script src="<%: Url.LatestContent("~/bundles/utilities_customreports")%>" type="text/javascript"></script>

<% 
	Dim bBradfordFactor As Boolean
	Dim mstrCaption As String
	Dim sErrMsg As String
	
	bBradfordFactor = (Session("utiltype") = "16")

	Dim objReport As HR.Intranet.Server.Report
	
	If Session("utiltype") = "" Or _
		 Session("utilname") = "" Or _
		 Session("utilid") = "" Or _
		 Session("action") = "" Then

		Response.Write("<table align=center class=""outline"" cellPadding=5 cellSpacing=0>" & vbCrLf)
		Response.Write("	<tr>" & vbCrLf)
		Response.Write("		<td>" & vbCrLf)
		Response.Write("			<table class=""invisible"" cellspacing=0 cellpadding=0>" & vbCrLf)
		Response.Write("			  <tr>" & vbCrLf)
		Response.Write("			    <td colspan=3 height=10></td>" & vbCrLf)
		Response.Write("			  </tr>" & vbCrLf)
		Response.Write("			  <tr> " & vbCrLf)
		Response.Write("			    <td colspan=3 align=center> " & vbCrLf)
		Response.Write("						<H3>Error</H3>" & vbCrLf)
		Response.Write("			    </td>" & vbCrLf)
		Response.Write("			  </tr> " & vbCrLf)
		Response.Write("			  <tr> " & vbCrLf)
		Response.Write("			    <td width=20 height=10></td> " & vbCrLf)
		Response.Write("			    <td> " & vbCrLf)
		Response.Write("						<H4>Not all session variables found</H4>" & vbCrLf)
		Response.Write("			    </td>" & vbCrLf)
		Response.Write("			    <td width=20></td> " & vbCrLf)
		Response.Write("			  </tr>" & vbCrLf)
		Response.Write("			  <tr> " & vbCrLf)
		Response.Write("			    <td width=20 height=10></td> " & vbCrLf)
		Response.Write("			    <td>Type = " & Session("utiltype") & vbCrLf)
		Response.Write("			    </td>" & vbCrLf)
		Response.Write("			    <td width=20></td> " & vbCrLf)
		Response.Write("			  </tr>" & vbCrLf)
		Response.Write("			  <tr> " & vbCrLf)
		Response.Write("			    <td width=20 height=10></td> " & vbCrLf)
		Response.Write("			    <td>Utility Name = " & Session("utilname") & vbCrLf)
		Response.Write("			    </td>" & vbCrLf)
		Response.Write("			    <td width=20></td> " & vbCrLf)
		Response.Write("			  </tr>" & vbCrLf)
		Response.Write("			  <tr> " & vbCrLf)
		Response.Write("			    <td width=20 height=10></td> " & vbCrLf)
		Response.Write("			    <td>Utility ID = " & Session("utilid") & vbCrLf)
		Response.Write("			    </td>" & vbCrLf)
		Response.Write("			    <td width=20></td> " & vbCrLf)
		Response.Write("			  </tr>" & vbCrLf)
		Response.Write("			  <tr> " & vbCrLf)
		Response.Write("			    <td width=20 height=10></td> " & vbCrLf)
		Response.Write("			    <td>Action = " & Session("action") & vbCrLf)
		Response.Write("			    </td>" & vbCrLf)
		Response.Write("			    <td width=20></td> " & vbCrLf)
		Response.Write("			  </tr>" & vbCrLf)
		Response.Write("			  <tr>" & vbCrLf)
		Response.Write("			    <td colspan=3 height=10>&nbsp;</td>" & vbCrLf)
		Response.Write("			  </tr>" & vbCrLf)
		Response.Write("			  <tr> " & vbCrLf)
		Response.Write("			    <td colspan=3 height=10 align=center> " & vbCrLf)
		Response.Write("						<input type=button id=cmdClose name=cmdClose value=Close style=""WIDTH: 80px"" width=80 class=""btn""" & vbCrLf)		'1
		Response.Write("                      onclick=""closeclick();""" & vbCrLf)
		Response.Write("                      onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
		Response.Write("                      onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
		Response.Write("                      onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
		Response.Write("                      onblur=""try{button_onBlur(this);}catch(e){}"" />" & vbCrLf)
		Response.Write("			    </td>" & vbCrLf)
		Response.Write("			  </tr>" & vbCrLf)
		Response.Write("			  <tr> " & vbCrLf)
		Response.Write("			    <td colspan=3 height=10></td>" & vbCrLf)
		Response.Write("			  </tr>" & vbCrLf)
		Response.Write("			</table>" & vbCrLf)
		Response.Write("		</td>" & vbCrLf)
		Response.Write("	</tr>" & vbCrLf)
		Response.Write("</table>" & vbCrLf)
		Response.Write("<input type=hidden id=txtSuccessFlag name=txtSuccessFlag value=1>" & vbCrLf)
		Response.Write("</BODY>" & vbCrLf)
		
		Response.End()
	End If

	Dim icount As Integer
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
	objReport = New HR.Intranet.Server.Report
	objReport.SessionInfo = CType(Session("SessionContext"), SessionInfo)

				
	' Pass required info to the DLL			
	objReport.CustomReportID = Session("utilid")
	objReport.ClientDateFormat = Session("LocaleDateFormat")
	objReport.LocalDecimalSeparator = Session("LocaleDecimalSeparator")

	If fok And bBradfordFactor Then
		dtStartDate = ConvertLocaleDateToSQL(Session("stdReport_StartDate"))
		dtEndDate = ConvertLocaleDateToSQL(Session("stdReport_EndDate"))
						
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
		bOutputPreview = Session("stdReport_OutputPreview")
		lngOutputFormat = Session("stdReport_OutputFormat")
		pblnOutputScreen = False
		pblnOutputSave = Session("stdReport_OutputSave")
		'plngOutputSaveExisting = session("stdReport_OutputSaveExisting")
		pblnOutputEmail = Session("stdReport_OutputEmail")
		plngOutputEmailID = Session("stdReport_OutputEmailAddr")
		pstrOutputEmailName = Session("stdReport_OutputEmailName")
		pstrOutputEmailSubject = Session("stdReport_OutputEmailSubject")
		pstrOutputEmailAttachAs = Session("stdReport_OutputEmailAttachAs")
		pstrOutputFilename = Session("stdReport_OutputFilename")
	End If

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
	
	gridReportData.DataSource = objReport.datCustomReportOutput
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
		' Response.Write("					<td class=""reportgraphic""></td>" & vbCrLf)
		Response.Write("					<td ALIGN=center colspan=10 NAME='tdOutputMSG' ID='tdOutputMSG'>" & vbCrLf)
	
			
%>

<form id="formReportData" runat="server">
	<asp:GridView ID="gridReportData" runat="server"
		AllowPaging="False"
		GridLines="None"
		CssClass="visibletablecolumn"
		ClientIDMode="Static">
		<Columns>
			<asp:BoundField DataField="rowtype" ItemStyle-CssClass="hiddentablecolumn" HeaderText="" />
		</Columns>
	</asp:GridView>
</form>


<%		


	Response.Write("					</td>" & vbCrLf)
	Response.Write("					<td width=20></td>" & vbCrLf)
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

<form name="frmPopup" id="frmPopup">
	<table align="center" class="outline" cellpadding="5" cellspacing="0">
		<tr>
			<td>
				<table class="invisible" cellspacing="0" cellpadding="0">
					<tr>
						<td colspan="3" height="10"></td>
					</tr>
					<tr>
						<td width="20" height="10"></td>
						<td align="center">


							<%	If bBradfordFactor = True Then
									mstrCaption = "Bradford Factor"
								Else
									mstrCaption = "Custom Report '" & Session("utilname").ToString() & "'"
								End If

								If fNoRecords Then
									Response.Write("						<H4>" & mstrCaption & " Completed successfully.</H4>" & vbCrLf)
								Else
									Response.Write("						<H4>" & mstrCaption & " Failed." & vbCrLf)
								End If
							%>
						</td>
						<td width="20"></td>
					</tr>
					<tr>
						<td width="20" height="10"></td>
						<td align="center" nowrap><%=objReport.ErrorString%>
						</td>
						<td width="20"></td>
					</tr>
					<tr>
						<td colspan="3" height="10">&nbsp;</td>
					</tr>
				<%--	<tr>
						<td colspan="3" height="10" align="center">
							<input type="button" id="cmdClose" name="cmdClose" value="Close" style="WIDTH: 80px" width="80" class="btn"
								onclick="closeclick();" />
						</td>
					</tr>--%>
					<tr>
						<td colspan="3" height="10"></td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</form>

<input type="hidden" id="txtSuccessFlag" name="txtSuccessFlag" value="3">
<%
End If
%>

<input type='hidden' id="txtNoRecs" name="txtNoRecs" value="<%=objReport.NoRecords%>">

<form id="frmOriginalDefinition" style="visibility: hidden; display: none">
	<%
		Response.Write("	<input type='hidden' id='txtDefn_Name' name='txtDefn_Name' value='" & objReport.ReportCaption.ToString() & "'>" & vbCrLf)
		Response.Write("	<input type='hidden' id=txtDefn_ErrMsg name=txtDefn_ErrMsg value=""" & sErrMsg & """>" & vbCrLf)
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

	Protected Sub gridReportData_DataBound(sender As Object, e As EventArgs) Handles gridReportData.DataBound
	End Sub

	Protected Sub gridReportData_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gridReportData.RowDataBound		
		
		Dim objReport As Report = CType(Session("CustomReport"), Report)
		Dim objThisColumn As ReportDetailItem
		Dim bGroupWithNext As Boolean

		If e.Row.RowType = DataControlRowType.Header Or e.Row.RowType = DataControlRowType.Footer Then
			e.Row.CssClass = "header"
			ReportColumnCount = 0
			
			For iCount = 2 To objReport.datCustomReportOutput.Columns.Count - 1
				objThisColumn = objReport.DisplayColumns(iCount - 2)

				e.Row.Cells(iCount).Visible = Not objThisColumn.IsHidden And Not bGroupWithNext
				bGroupWithNext = objThisColumn.GroupWithNextColumn

				If e.Row.Cells(iCount).Visible Then ReportColumnCount += 1
			Next
					
		Else

			If e.Row.Cells(0).Text = HR.Intranet.Server.Enums.RowType.GrandSummary Then
				e.Row.CssClass = "grandsummaryrow"
					
			ElseIf Not e.Row.Cells(0).Text = HR.Intranet.Server.Enums.RowType.Data Then
				e.Row.CssClass = "summarytablerow"
			Else
				'e.Row.CssClass = "ui-widget-content jqgrow ui-row-ltr"
			End If
					
		End If
							
		e.Row.Cells(0).Visible = False
				
		If Not objReport.HasSummaryColumns Then
			e.Row.Cells(1).Visible = False
		End If

		For iCount = 2 To objReport.datCustomReportOutput.Columns.Count - 1
						
			objThisColumn = objReport.DisplayColumns(iCount - 2)
					
			If objThisColumn.IsNumeric Then
				e.Row.Cells(iCount).HorizontalAlign = HorizontalAlign.Right
			Else
				e.Row.Cells(iCount).HorizontalAlign = HorizontalAlign.Left
			End If
	
			If Session("utiltype") = UtilityType.utlBradfordFactor Then
				e.Row.Cells(iCount).Visible = Not objThisColumn.IsHidden
			Else
				e.Row.Cells(iCount).Visible = Not (objThisColumn.IDColumnName.StartsWith("?"))
			End If
			
		Next

	End Sub
		
	
	Public Property ReportColumnCount() As Integer
	
	
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

	<iframe name="submit-iframe" style="display: none;"></iframe>

</form>

<script type="text/javascript">
	

	//Shrink to fit, or set to 100px per column?
	var ShrinkToFit = false;
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
		var iVisibleCount = Number("<%:ReportColumnCount%>");
		if (iVisibleCount < 8) ShrinkToFit = true;
		gridWidth = 770;
		gridHeight = 390;
	}

	tableToGrid("#gridReportData", {
		shrinkToFit: ShrinkToFit,
		width: gridWidth,
		height: gridHeight,
		ignoreCase: true,
		cmTemplate: { sortable: false },
		rowNum: 200000,
		loadComplete: function () {
			stylejqGrid();
		}
	});



	function stylejqGrid() {
		//jqGrid style overrides
		$('#gview_gridReportData tr.jqgrow td').css('vertical-align', 'top'); //float text to top, in case of multi-line cells
		$('#gview_gridReportData tr.footrow td').css('vertical-align', 'top'); //float text to top, in case of multi-line footers
		$('#gview_gridReportData .s-ico span').css('display', 'none'); //hide the sort order icons - they don't tie in to the dataview model.
		$("#gview_gridReportData > .ui-jqgrid-titlebar").text("<%=objReport.ReportCaption%>"); //Activate title bar for the grid as this will then go naturally into the print functionality.
		$("#gview_gridReportData > .ui-jqgrid-titlebar").height("20px"); //no title bar; this is in the dialog title
		$("#gview_gridReportData .ui-jqgrid-titlebar").show();

	}
	if (menu_isSSIMode()) $('#gbox_gridReportData').css('margin', '0 auto'); //center the report in self-service screen.
	

</script>

