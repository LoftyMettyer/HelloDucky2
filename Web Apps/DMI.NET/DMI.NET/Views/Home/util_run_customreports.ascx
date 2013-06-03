<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script src="<%: Url.Content("~/bundles/utilities_customreports")%>" type="text/javascript"></script>

<object
    id="ClientDLL"
    classid="CLSID:40E1755A-5A2D-4AEE-99E7-65E7D455F799"
    codebase="cabs/COAInt_Client.CAB#version=1,0,0,147">
</object>

<% 
    Dim bBradfordFactor As Boolean
    Dim mstrCaption As String
    Dim sErrMsg As String
    
	bBradfordFactor = (session("utiltype") = "16")
%>
 
    <script type="text/javascript">
        function reports_window_onload() {
            customreport_loadAddRecords();
        }
    </script>
    

    <%
	if session("utiltype") = "" or _ 
	   session("utilname") = "" or _ 
	   session("utilid") = "" or _ 
	   session("action") = "" then

            Response.Write("<table align=center class=""outline"" cellPadding=5 cellSpacing=0>" & vbCrLf)
            Response.Write("	<TR>" & vbCrLf)
            Response.Write("		<TD>" & vbCrLf)
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
            Response.Write("						<input type=button id=cmdClose name=cmdClose value=Close style=""WIDTH: 80px"" width=80 class=""btn""" & vbCrLf)    '1
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
		
		Response.End
	end if

        Dim icount As Integer
        Dim fok As Boolean
        Dim objReport As HR.Intranet.Server.Report
        Dim fNotCancelled As Boolean

        Dim dtStartDate
        Dim dtEndDate
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
        
	fok = true
	fNotCancelled = true

	' Create the reference to the DLL (Report Class)
        objReport = New HR.Intranet.Server.Report
        
	' Pass required info to the DLL

        CallByName(objReport, "Connection", CallType.Let, Session("databaseConnection"))
        objReport.Username = Session("username")
        
	objReport.CustomReportID = session("utilid")
	objReport.ClientDateFormat = session("LocaleDateFormat")
	objReport.LocalDecimalSeparator = session("LocaleDecimalSeparator")

	if fok and bBradfordFactor then
            dtStartDate = convertLocaleDateToSQL(Session("stdReport_StartDate"))
            dtEndDate = convertLocaleDateToSQL(Session("stdReport_EndDate"))
            
		strAbsenceTypes = session("stdReport_AbsenceTypes")
		lngFilterID = session("stdReport_FilterID")
		lngPicklistID = session("stdReport_PicklistID")
		lngPersonnelID = session("optionRecordID")

		bBradford_SRV = session("stdReport_Bradford_SRV")
		bBradford_ShowDurations = session("stdReport_Bradford_ShowDurations")
		bBradford_ShowInstances = session("stdReport_Bradford_ShowInstances")
		bBradford_ShowFormula = session("stdReport_Bradford_ShowFormula")
		bBradford_OmitBeforeStart = session("stdReport_Bradford_OmitBeforeStart")
		bBradford_OmitAfterEnd = session("stdReport_Bradford_OmitAfterEnd")
		bBradford_txtOrderBy1 = session("stdReport_Bradford_txtOrderBy1")
		lngBradford_txtOrderBy1ID = Clng(session("stdReport_Bradford_txtOrderBy1ID"))
		bBradford_txtOrderBy1Asc = session("stdReport_Bradford_txtOrderBy1Asc")
		bBradford_txtOrderBy2 = session("stdReport_Bradford_txtOrderBy2")
		lngBradford_txtOrderBy2ID = Clng(session("stdReport_Bradford_txtOrderBy2ID"))
		bBradford_txtOrderBy2Asc = session("stdReport_Bradford_txtOrderBy2Asc")
		bPrintFilterPickList = session("stdReport_PrintFilterPicklistHeader")

		bMinBradford = session("stdReport_MinimumBradfordFactor")
            lngMinBradfordAmount = CLng(Session("stdReport_MinimumBradfordFactorAmount"))
		pbDisplayBradfordDetail	= session("stdReport_DisplayBradfordDetail")

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
		pstrOutputEmailName = session("stdReport_OutputEmailName")
		pstrOutputEmailSubject = session("stdReport_OutputEmailSubject")
		pstrOutputEmailAttachAs = session("stdReport_OutputEmailAttachAs")
		pstrOutputFilename = session("stdReport_OutputFilename")
	end if

	if fok and not bBradfordFactor then 
		fok = objReport.GetCustomReportDefinition
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok and not bBradfordFactor then 
		fok = objReport.GetDetailsRecordsets
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok and bBradfordFactor then 
		fok = objReport.SetBradfordDisplayOptions(bBradford_SRV, bBradford_ShowDurations, bBradford_ShowInstances, bBradford_ShowFormula, bPrintFilterPickList, pbDisplayBradfordDetail)

		if lngPersonnelID = 0 then
			fok = objReport.SetBradfordOrders(bBradford_txtOrderBy1, bBradford_txtOrderBy2, bBradford_txtOrderBy1Asc, bBradford_txtOrderBy2Asc, lngBradford_txtOrderBy1ID, lngBradford_txtOrderBy2ID)
		else
			fok = objReport.SetBradfordOrders("None", "None", false, false, 0, 0)		
		end if 

		fok = objReport.SetBradfordIncludeOptions(bBradford_OmitBeforeStart, bBradford_OmitAfterEnd, lngPersonnelID, lngFilterID, lngPicklistID, bMinBradford, lngMinBradfordAmount)
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok and bBradfordFactor then 
		fok = objReport.GetBradfordReportDefinition(dtStartDate, dtEndDate)
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok and bBradfordFactor then 
		fok = objReport.GetBradfordRecordSet
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

        Dim aPrompts
        
	aPrompts =  Session("Prompts_" & session("utiltype") & "_" & session("utilid"))
	if fok then 
		fok = objReport.SetPromptedValues(aPrompts)
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok then 
		fok = objReport.GenerateSQL 
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok and bBradfordFactor then 
		fok = objReport.GenerateSQLBradford(strAbsenceTypes)
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok then 
		fok = objReport.AddTempTableToSQL 
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok then 
		fok = objReport.MergeSQLStrings 
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok then
		fok = objReport.UDFFunctions(true)
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok then 
		fok = objReport.ExecuteSql 
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok then 
		fok = objReport.UDFFunctions(false)
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok and bBradfordFactor then 
		fok = objreport.CalculateBradfordFactors()
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

        If fok And objReport.ChildCount > 1 And objReport.UsedChildCount > 1 Then
            fok = objReport.CreateMutipleChildTempTable
            fNotCancelled = Response.IsClientConnected
            If fok Then fok = fNotCancelled
        End If

	if fok then 
		fok = objReport.CheckRecordSet 
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok then 
		fok = objreport.OutputGridDefinition 
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

        Dim arrayDefinition
        
	if fok then
	  arrayDefinition = objreport.OutputArray_Definition 
	end if	

	' Need to pass in defined output options 
	'	(standard report reads from definition, which of course we don't have in a standard report)
	if fok and bBradfordFactor then
		fok = objreport.SetBradfordDefaultOutputOptions(bOutputPreview, lngOutputFormat, pblnOutputScreen, pblnOutputPrinter, pstrOutputPrinterName, pblnOutputSave, plngOutputSaveExisting, pblnOutputEmail, plngOutputEmailID, pstrOutputEmailName, pstrOutputEmailSubject, pstrOutputEmailAttachAs, pstrOutputFilename)
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if		
	
	if fok then 
		fok = objreport.OutputGridColumns 
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

        Dim arrayColumnsDefinition
        Dim arrayPageBreakValues
        Dim arrayVisibleColumns

	if fok then
		arrayColumnsDefinition = objreport.OutputArray_Columns 

		if fok then 
			fok = objreport.PopulateGrid_LoadRecords 
			fNotCancelled = Response.IsClientConnected 
			if fok then fok = fNotCancelled
		end if

	  if fok then 
			fok = objreport.PopulateGrid_HideColumns 
			fNotCancelled = Response.IsClientConnected 
			if fok then fok = fNotCancelled
		end if

		arrayColumnsDefinition = objreport.OutputArray_Columns 
		            
		arrayPageBreakValues = objreport.OutputArray_PageBreakValues		
		arrayVisibleColumns = objreport.OutputArray_VisibleColumns
		
            If fok Then
                Response.Write(objReport.Output_GridForm)
            End If
        End If

	if fok then

            
            Response.Write("<FORM action=""util_run_outputoptions"" method=post id=frmExportData name=frmExportData>" & vbCrLf)
            Response.Write("  <INPUT type=""hidden"" id=txtPreview name=txtPreview value=""" & objReport.OutputPreview & """>" & vbCrLf)
            Response.Write("  <INPUT type=""hidden"" id=txtFormat name=txtFormat value=" & objReport.OutputFormat & ">" & vbCrLf)
            Response.Write("  <INPUT type=""hidden"" id=txtScreen name=txtScreen value=""" & objReport.OutputScreen & """>" & vbCrLf)
            Response.Write("  <INPUT type=""hidden"" id=txtPrinter name=txtPrinter value=""" & objReport.OutputPrinter & """>" & vbCrLf)
            Response.Write("  <INPUT type=""hidden"" id=txtPrinterName name=txtPrinterName value=""" & objReport.OutputPrinterName & """>" & vbCrLf)
            Response.Write("  <INPUT type=""hidden"" id=txtSave name=txtSave value=""" & objReport.OutputSave & """>" & vbCrLf)
            Response.Write("  <INPUT type=""hidden"" id=txtSaveExisting name=txtSaveExisting value=""" & CStr(objReport.OutputSaveExisting) & """>" & vbCrLf)
            Response.Write("  <INPUT type=""hidden"" id=txtEmail name=txtEmail value=""" & objReport.OutputEmail & """>" & vbCrLf)
            Response.Write("  <INPUT type=""hidden"" id=txtEmailAddr name=txtEmailAddr value=""" & CStr(objReport.OutputEmailID) & """>" & vbCrLf)
            Response.Write("  <INPUT type=""hidden"" id=txtEmailAddrName name=txtEmailAddrName value=""" & Replace(objReport.OutputEmailGroupName, """", "&quot;") & """>" & vbCrLf)
            Response.Write("  <INPUT type=""hidden"" id=txtEmailSubject name=txtEmailSubject value=""" & Replace(objReport.OutputEmailSubject, """", "&quot;") & """>" & vbCrLf)
            Response.Write("  <INPUT type=""hidden"" id=txtEmailAttachAs name=txtEmailAttachAs value=""" & Replace(objReport.OutputEmailAttachAs, """", "&quot;") & """>" & vbCrLf)
            Response.Write("  <INPUT type=""hidden"" id=txtFileName name=txtFileName value=""" & objReport.OutputFilename & """>" & vbCrLf)
            Response.Write("  <INPUT type=""hidden"" id=txtUtilType name=txtUtilType value=""" & Session("UtilType") & """>" & vbCrLf)

            
            
    For icount = 0 To (UBound(arrayPageBreakValues))
        Response.Write("	<INPUT type=hidden id=txtPageBreak_" & icount & " name=txtPageBreak_" & icount & " value=""" & Replace(arrayPageBreakValues(icount), """", "&quot;") & """>" & vbCrLf)
    Next

    Response.Write("			<input type=hidden id=pagebreak name=pagebreak value=""" & objReport.ReportHasPageBreak & """>" & vbCrLf)
    Response.Write("			<input type=hidden id=txtSummaryReport name=txtSummaryReport value=""" & objReport.ReportHasSummaryInfo & """>" & vbCrLf)
		
    For icount = 0 To (UBound(arrayVisibleColumns, 2))
        Response.Write("	<INPUT type=hidden id=txtVisHeading_" & icount & " name=txtVisHeading_" & icount & " value=""" & Replace(Replace(arrayVisibleColumns(0, icount), """", "&quot;"), "_", " ") & """>" & vbCrLf)
        Response.Write("	<INPUT type=hidden id=txtVisDataType_" & icount & " name=txtVisDataType_" & icount & " value=""" & arrayVisibleColumns(1, icount) & """>" & vbCrLf)
        Response.Write("	<INPUT type=hidden id=txtVisDecimals_" & icount & " name=txtVisDecimals_" & icount & " value=""" & arrayVisibleColumns(2, icount) & """>" & vbCrLf)
        Response.Write("	<INPUT type=hidden id=txtVis1000Separator_" & icount & " name=txtVis1000Separator_" & icount & " value=""" & arrayVisibleColumns(3, icount) & """>" & vbCrLf)
    Next
    Response.Write("	<INPUT type=hidden id=txtVisColCount name=txtVisColCount value=" & UBound(arrayVisibleColumns, 2) & ">" & vbCrLf)
    Response.Write("</FORM>" & vbCrLf)
	
    Response.Write("<script type=""text/javascript"">" & vbCrLf)

            ' Change the output text if Bradford Factor Report

            Response.Write("function ExportData(strMode) " & vbCrLf)
            Response.Write("	{" & vbCrLf & vbCrLf)
            
            Response.Write("	var bm;" & vbCrLf)
            Response.Write("    var fok;" & vbCrLf)
            Response.Write("	var sBreakValue = new String('');" & vbCrLf)
            Response.Write("	var blnBreakCheck = false;" & vbCrLf)
            Response.Write("    var frmExportData = OpenHR.getForm(""reportworkframe"",""frmExportData_ORIGINAL"");" & vbCrLf)
            
            Dim objUser As New HR.Intranet.Server.clsSettings          
            
            objReport.Username = Session("username").ToString()
            CallByName(objUser, "Connection", CallType.Let, Session("databaseConnection"))            
            
            'MH20031113 Fault 7606 Reset Columns and Styles...
            Response.Write("  ClientDLL.ResetColumns();" & vbCrLf)
            Response.Write("  ClientDLL.ResetStyles();" & vbCrLf)
            Response.Write("  ClientDLL.UserName = """ & CleanStringForJavaScript(Session("Username")) & """;" & vbCrLf)
            Response.Write("  ClientDLL.SaveAsValues = """ & CleanStringForJavaScript(Session("OfficeSaveAsValues")) & """;" & vbCrLf)

            Response.Write("  ClientDLL.SettingLocations(")
            Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleCol", "3")) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleRow", "2")) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataCol", "2")) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataRow", "4")) & ");" & vbCrLf)

            Response.Write("  ClientDLL.SettingTitle(")
            If (objUser.GetUserSetting("Output", "TitleGridLines", "0") = "1") Then
                Response.Write("true, ")
            Else
                Response.Write("false, ")
            End If

            If (objUser.GetUserSetting("Output", "TitleBold", "1") = "1") Then
                Response.Write("true, ")
            Else
                Response.Write("false, ")
            End If

            If (objUser.GetUserSetting("Output", "TitleUnderline", "0") = "1") Then
                Response.Write("true, ")
            Else
                Response.Write("false, ")
            End If

            Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleBackcolour", "16777215")) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleForecolour", "6697779")) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "TitleBackcolour", "16777215")))) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "TitleForecolour", "6697779")))) & ");" & vbCrLf)

            Response.Write("  ClientDLL.SettingHeading(")
            Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingGridLines", "1")) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingBold", "1")) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingUnderline", "0")) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingBackcolour", "16248553")) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingForecolour", "6697779")) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "HeadingBackcolour", "16248553")))) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "HeadingForecolour", "6697779")))) & ");" & vbCrLf)

            Response.Write("  ClientDLL.SettingData(")
            Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataGridLines", "1")) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataBold", "0")) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataUnderline", "0")) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataBackcolour", "15988214")) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetUserSetting("Output", "DataForecolour", "6697779")) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "DataBackcolour", "15988214")))) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "DataForecolour", "6697779")))) & ");" & vbCrLf)

            Response.Write("  ClientDLL.InitialiseStyles();" & vbCrLf)
            Response.Write("  ClientDLL.SettingOptions(")
            Response.Write("""" & CleanStringForJavaScript(objUser.GetUserSetting("Output", "WordTemplate", "")) & """, ")
            Response.Write("""" & CleanStringForJavaScript(objUser.GetUserSetting("Output", "ExcelTemplate", "")) & """, ")

            If (objUser.GetUserSetting("Output", "ExcelGridlines", "0") = "1") Then
                Response.Write("true, ")
            Else
                Response.Write("false, ")
            End If

            If (objUser.GetUserSetting("Output", "ExcelHeaders", "0") = "1") Then
                Response.Write("true, ")
            Else
                Response.Write("false, ")
            End If
		
            If (objUser.GetUserSetting("Output", "ExcelOmitSpacerRow", "0") = "1") Then
                Response.Write("true, ")
            Else
                Response.Write("false, ")
            End If
		
            If (objUser.GetUserSetting("Output", "ExcelOmitSpacerCol", "0") = "1") Then
                Response.Write("true, ")
            Else
                Response.Write("false, ")
            End If
		
            If (objUser.GetUserSetting("Output", "AutoFitCols", "1") = "1") Then
                Response.Write("true, ")
            Else
                Response.Write("false, ")
            End If

            If (objUser.GetUserSetting("Output", "Landscape", "1") = "1") Then
                Response.Write("true, " & vbCrLf)
            Else
                Response.Write("false, " & vbCrLf)
            End If

            Response.Write("document.all.item(""txtSysPerm_EMAILGROUPS_VIEW"").value);" & vbCrLf)

            'Set Options
            If Not objReport.OutputPreview Then
                Dim lngFormat As Long
                Dim blnScreen As Boolean
                Dim blnPrinter As Boolean
                Dim strPrinterName As String
                Dim blnSave As Boolean
                Dim lngSaveExisting As Long
                Dim blnEmail As Boolean
                Dim lngEmailGroupID As Long
                Dim strEmailSubject As String
                Dim strEmailAttachAs As String
                Dim strFileName As String
			
                lngFormat = CleanStringForJavaScript(objReport.OutputFormat)
                blnScreen = CleanStringForJavaScript(LCase(objReport.OutputScreen))
                blnPrinter = CleanStringForJavaScript(LCase(objReport.OutputPrinter))
                strPrinterName = CleanStringForJavaScript(objReport.OutputPrinterName)
                blnSave = CleanStringForJavaScript(LCase(objReport.OutputSave))
                lngSaveExisting = CleanStringForJavaScript(objReport.OutputSaveExisting)
                blnEmail = CleanStringForJavaScript(LCase(objReport.OutputEmail))
                lngEmailGroupID = CLng(objReport.OutputEmailID)
                strEmailSubject = CleanStringForJavaScript(objReport.OutputEmailSubject)
                strEmailAttachAs = CleanStringForJavaScript(objReport.OutputEmailAttachAs)
                'strFileName = objreport.OutputFilename 
                strFileName = CleanStringForJavaScript(objReport.OutputFilename)
			
                Dim cmdEmailAddr
                Dim prmEmailGroupID
                Dim rstEmailAddr
                Dim sErrorDescription As String = ""
                Dim iLoop As Integer
                Dim sEmailAddresses As String = ""
                
                If (objReport.OutputEmail) And (objReport.OutputEmailID > 0) Then
				
                    cmdEmailAddr = CreateObject("ADODB.Command")
                    cmdEmailAddr.CommandText = "spASRIntGetEmailGroupAddresses"
                    cmdEmailAddr.CommandType = 4 ' Stored procedure
                    cmdEmailAddr.ActiveConnection = Session("databaseConnection")

                    prmEmailGroupID = cmdEmailAddr.CreateParameter("EmailGroupID", 3, 1) ' 3=integer, 1=input
                    cmdEmailAddr.Parameters.Append(prmEmailGroupID)
                    prmEmailGroupID.value = CleanNumeric(lngEmailGroupID)

                    Err.Clear()
                    rstEmailAddr = cmdEmailAddr.Execute

                    If (Err.Number <> 0) Then
                        sErrorDescription = "Error getting the email addresses for group." & vbCrLf & FormatError(Err.Description)
                    End If

                    If Len(sErrorDescription) = 0 Then
                        iLoop = 1
                        Do While Not rstEmailAddr.EOF
                            If iLoop > 1 Then
                                sEmailAddresses = sEmailAddresses & ";"
                            End If
                            sEmailAddresses = sEmailAddresses & rstEmailAddr.Fields("Fixed").Value
                            rstEmailAddr.MoveNext()
                            iLoop = iLoop + 1
                        Loop
						
                        ' Release the ADO recordset object.
                        rstEmailAddr.close()
                    End If
							
                    rstEmailAddr = Nothing
                    cmdEmailAddr = Nothing
                End If
			
                Response.Write("  fok = ClientDLL.SetOptions(false, " & _
                                                        lngFormat & "," & blnScreen & ", " & _
                                                        blnPrinter & ",""" & strPrinterName & """, " & _
                                                        blnSave & "," & lngSaveExisting & ", " & _
                                                        blnEmail & ", """ & CleanStringForJavaScript(sEmailAddresses) & """, """ & _
                                                        strEmailSubject & """,""" & strEmailAttachAs & """,""" & strFileName & """);" & vbCrLf)
            Else
                Response.Write("  fok = ClientDLL.SetOptions(false, " & _
                                                "parseFloat(frmExportData.txtFormat.value), frmExportData.txtScreen.value, " & _
                                                "frmExportData.txtPrinter.value, frmExportData.txtPrinterName.value, " & _
                                                "frmExportData.txtSave.value, parseFloat(frmExportData.txtSaveExisting.value), " & _
                                                "frmExportData.txtEmail.value, frmDataFrame.txtEmailGroupAddr.value, " & _
                                                "frmExportData.txtEmailSubject.value, frmExportData.txtEmailAttachAs.value, frmExportData.txtFileName.value);" & vbCrLf)
		
            End If
		
            Response.Write("  if (fok == true) {" & vbCrLf)

            '(Chart or Pivot) and not summary report
            Response.Write("  blnIndicatorColumn = false;" & vbCrLf)
            Response.Write("  if ((frmExportData.txtFormat.value == ""5"") " & vbCrLf)
            Response.Write("   || (frmExportData.txtFormat.value == ""6"")) {" & vbCrLf)
            Response.Write("    blnIndicatorColumn = (frmExportData.txtSummaryReport.value == 'False')" & vbCrLf)
            Response.Write("  }" & vbCrLf)
		
            Response.Write("  ClientDLL.SizeColumnsIndependently = true;" & vbCrLf)

            Response.Write("  if (frmExportData.txtFormat.value == ""0"") {" & vbCrLf)
            Response.Write("    if (frmExportData.txtPrinter.value.toLowerCase() == ""true"") {" & vbCrLf)
            Response.Write("      ClientDLL.SetPrinter();" & vbCrLf)
            Response.Write("      dataOnlyPrint();" & vbCrLf)
            Response.Write("      ClientDLL.ResetDefaultPrinter();" & vbCrLf)
            Response.Write("    }" & vbCrLf)
            Response.Write("  }" & vbCrLf)
            Response.Write("  else {" & vbCrLf)
            Response.Write("  ClientDLL.HeaderRows = 1;" & vbCrLf)

            Response.Write("  if (ClientDLL.GetFile() == true) " & vbCrLf)
            Response.Write("		{" & vbCrLf)
			
            'Response.Write "		if (frmExportData.pagebreak.value == 'True') " & vbcrlf
            Response.Write("		if (frmExportData.pagebreak.value.toLowerCase() == ""true"") " & vbCrLf)
            Response.Write("			{" & vbCrLf)
            Response.Write("			var lngActualRow = new Number(0);" & vbCrLf)
            '	Response.Write "			ClientDLL.PageTitles = true;" & vbcrlf
			
            Response.Write("			ClientDLL.ArrayDim(frmExportData.txtVisColCount.value, 0);" & vbCrLf & vbCrLf)
            Response.Write("			lngActualRow = 0;" & vbCrLf)

            Response.Write("      frmOutput.ssOleDBGridDefSelRecords.MoveFirst();" & vbCrLf)
            Response.Write("      for (lngRow = 0; lngRow <= frmOutput.ssOleDBGridDefSelRecords.Rows; lngRow++)" & vbCrLf)
            Response.Write("				{" & vbCrLf)

            Response.Write("				lngActualRow = lngActualRow + 1; " & vbCrLf)
            Response.Write("				bm = frmOutput.ssOleDBGridDefSelRecords.AddItemBookmark(lngRow);" & vbCrLf)
            'Response.Write "				if (lngRow == (frmOutput.ssOleDBGridDefSelRecords.Rows - 1))" & vbcrlf
            Response.Write("				if (lngRow == (frmOutput.ssOleDBGridDefSelRecords.Rows))" & vbCrLf)
            Response.Write("					{" & vbCrLf)

            Response.Write("					if (frmExportData.txtSummaryReport.value == 'True') " & vbCrLf)
            Response.Write("						{" & vbCrLf)
            Response.Write("						sBreakValue = 'Grand Totals';" & vbCrLf)
            Response.Write("						}" & vbCrLf)
            Response.Write("					else" & vbCrLf)
            Response.Write("						{" & vbCrLf)
            Response.Write("						sBreakValue = document.getElementById('txtPageBreak_'+lngRow).value;" & vbCrLf)
            Response.Write("						}" & vbCrLf)
			
            Response.Write("				  if (lngActualRow > 0) {" & vbCrLf)
            If bBradfordFactor = True Then
                Response.Write("					ClientDLL.AddPage(replace(frmOutput.ssOleDBGridDefSelRecords.Caption,'&&','&'),'Bradford Factor');" & vbCrLf)
            Else
                Response.Write("					ClientDLL.AddPage(replace(frmOutput.ssOleDBGridDefSelRecords.Caption,'&&','&'),sBreakValue);" & vbCrLf)
            End If
            Response.Write("					var sColHeading = new String(''); " & vbCrLf)
            Response.Write("					var iColDataType = new Number(0); " & vbCrLf)
            Response.Write("					var iColDecimals = new Number(0); " & vbCrLf)
            Response.Write("					for (var lngCol = 0; lngCol<=frmExportData.txtVisColCount.value; lngCol++)" & vbCrLf)
            Response.Write("						{" & vbCrLf)
            Response.Write("						sColHeading = document.getElementById('txtVisHeading_'+lngCol).value;" & vbCrLf)
            Response.Write("						iColDataType = document.getElementById('txtVisDataType_'+lngCol).value;" & vbCrLf)
            Response.Write("						iColDecimals = document.getElementById('txtVisDecimals_'+lngCol).value;" & vbCrLf)
            Response.Write("						ClientDLL.AddColumn(sColHeading, iColDataType, iColDecimals, false);" & vbCrLf)
            Response.Write("						ClientDLL.ArrayAddTo(lngCol, 0, sColHeading);" & vbCrLf & vbCrLf)
            Response.Write("						}" & vbCrLf)
            Response.Write("                    ClientDLL.DataArray();" & vbCrLf)
            Response.Write("					lngActualRow = 0;" & vbCrLf)
            Response.Write("					blnBreakCheck = true;" & vbCrLf)
            Response.Write("					sBreakValue = '';" & vbCrLf)

            Response.Write("					}" & vbCrLf)
            Response.Write("				  }" & vbCrLf)
            Response.Write("        else if ((frmOutput.ssOleDBGridDefSelRecords.Columns(0).CellText(bm) == '*')" & vbCrLf)
            Response.Write("								&& (!blnBreakCheck))" & vbCrLf)
            Response.Write("					{" & vbCrLf)
			
            Response.Write("					sBreakValue = document.getElementById('txtPageBreak_'+lngRow).value;" & vbCrLf)

            If bBradfordFactor = True Then
                Response.Write("					ClientDLL.AddPage(replace(frmOutput.ssOleDBGridDefSelRecords.Caption,'&&','&'),'Bradford Factor');" & vbCrLf)
            Else
                Response.Write("					ClientDLL.AddPage(replace(frmOutput.ssOleDBGridDefSelRecords.Caption,'&&','&'),sBreakValue);" & vbCrLf)
            End If

            Response.Write("					var sColHeading = new String(''); " & vbCrLf)
            Response.Write("					var iColDataType = new Number(0); " & vbCrLf)
            Response.Write("					var iColDecimals = new Number(0); " & vbCrLf)
            Response.Write("					var iCol1000 = new Number(0); " & vbCrLf)
            Response.Write("					for (var lngCol = 0; lngCol<=frmExportData.txtVisColCount.value; lngCol++)" & vbCrLf)
            Response.Write("						{" & vbCrLf)
            Response.Write("						sColHeading = document.getElementById('txtVisHeading_'+lngCol).value;" & vbCrLf)
            Response.Write("						iColDataType = document.getElementById('txtVisDataType_'+lngCol).value;" & vbCrLf)
            Response.Write("						iColDecimals = document.getElementById('txtVisDecimals_'+lngCol).value;" & vbCrLf)
            Response.Write("						iCol1000 = document.getElementById('txtVis1000Separator_'+lngCol).value;" & vbCrLf)
            Response.Write("						ClientDLL.AddColumn(sColHeading, iColDataType, iColDecimals, iCol1000);" & vbCrLf)
            Response.Write("						ClientDLL.ArrayAddTo(lngCol, 0, sColHeading);" & vbCrLf & vbCrLf)
            Response.Write("						}" & vbCrLf)

            Response.Write("          ClientDLL.DataArray();" & vbCrLf)
            Response.Write("					ClientDLL.ArrayDim(frmExportData.txtVisColCount.value, 0);" & vbCrLf & vbCrLf)
            Response.Write("					lngActualRow = 0;" & vbCrLf)
            Response.Write("					blnBreakCheck = true;" & vbCrLf)
            Response.Write("					sBreakValue = '';" & vbCrLf)
            Response.Write("					ClientDLL.ResetColumns();" & vbCrLf)
            Response.Write("					ClientDLL.ResetStyles();" & vbCrLf)
			
            Response.Write("					}" & vbCrLf & vbCrLf)
            Response.Write("        else if (frmOutput.ssOleDBGridDefSelRecords.Columns(0).CellText(bm) != '*')" & vbCrLf)
            Response.Write("					{" & vbCrLf)
			
            Response.Write("					blnBreakCheck = false;" & vbCrLf)
            Response.Write("					lngCol = 0;" & vbCrLf)
            Response.Write("					ClientDLL.ArrayReDim();" & vbCrLf & vbCrLf)
            Response.Write("					for (var lngCount=0; lngCount<frmOutput.ssOleDBGridDefSelRecords.Columns.Count; lngCount++)" & vbCrLf)
            Response.Write("						{" & vbCrLf)
            Response.Write("						if (frmOutput.ssOleDBGridDefSelRecords.Columns(lngCount).Visible == true)" & vbCrLf)
            Response.Write("							{" & vbCrLf)
            Response.Write("						  ClientDLL.ArrayAddTo(lngCol, lngActualRow, frmOutput.ssOleDBGridDefSelRecords.Columns(lngCount).CellText(bm));" & vbCrLf)
            Response.Write("						  lngCol++;" & vbCrLf)
            Response.Write("							}" & vbCrLf)
            Response.Write("						}" & vbCrLf)

            Response.Write("					}" & vbCrLf)
            Response.Write("				}" & vbCrLf)
            Response.Write("			} " & vbCrLf)
            Response.Write("		else // no page break " & vbCrLf)
            Response.Write("			{ " & vbCrLf)

            Response.Write("      ClientDLL.ArrayDim(frmExportData.txtVisColCount.value, 0);" & vbCrLf & vbCrLf)
            If bBradfordFactor = True Then
                Response.Write("			ClientDLL.PageTitles = false;" & vbCrLf)
                Response.Write("      ClientDLL.AddPage(replace(frmOutput.ssOleDBGridDefSelRecords.Caption,'&&','&'),'" & "Bradford Factor" & "');" & vbCrLf)
            Else
                Response.Write("      ClientDLL.AddPage(replace(frmOutput.ssOleDBGridDefSelRecords.Caption,'&&','&'),'" & CleanStringForJavaScript(objReport.BaseTableName) & "');" & vbCrLf)
            End If

            Response.Write("			var sColHeading = new String(''); " & vbCrLf)
            Response.Write("			var iColDataType = new Number(0); " & vbCrLf)
            Response.Write("			var iColDecimals = new Number(0); " & vbCrLf)
            Response.Write("			var iCol1000 = new Number(0); " & vbCrLf)
            Response.Write("			for (var lngCol = 0; lngCol<=frmExportData.txtVisColCount.value; lngCol++)" & vbCrLf)
            Response.Write("				{" & vbCrLf)
            Response.Write("				sColHeading = document.getElementById('txtVisHeading_'+lngCol).value;" & vbCrLf)
            Response.Write("				iColDataType = document.getElementById('txtVisDataType_'+lngCol).value;" & vbCrLf)
            Response.Write("				iColDecimals = document.getElementById('txtVisDecimals_'+lngCol).value;" & vbCrLf)
            Response.Write("				iCol1000 = document.getElementById('txtVis1000Separator_'+lngCol).value;" & vbCrLf)
            Response.Write("				ClientDLL.AddColumn(sColHeading, iColDataType, iColDecimals, iCol1000);" & vbCrLf)
            Response.Write("				ClientDLL.ArrayAddTo(lngCol, 0, sColHeading);" & vbCrLf & vbCrLf)
            Response.Write("				}" & vbCrLf)
			
            Response.Write("	  lngActualRow = 0;" & vbCrLf)
            Response.Write("      for (lngRow = 0; lngRow < frmOutput.ssOleDBGridDefSelRecords.Rows; lngRow++)" & vbCrLf)
            Response.Write("				{" & vbCrLf)
            Response.Write("				bm = frmOutput.ssOleDBGridDefSelRecords.AddItemBookmark(lngRow);" & vbCrLf)

            'MH20040403
            Response.Write("    if (blnIndicatorColumn) " & vbCrLf)
            Response.Write("			{" & vbCrLf)
            Response.Write("			var sTestValue = new String(''); " & vbCrLf)
            Response.Write("			sTestValue = frmOutput.ssOleDBGridDefSelRecords.Columns(0).CellText(bm).substr(0,1); " & vbCrLf)
            Response.Write("			var blnIgnoreRow = (sTestValue == '*'); " & vbCrLf)
            Response.Write("			}" & vbCrLf)
            Response.Write("		else" & vbCrLf)
            Response.Write("			{" & vbCrLf)
            Response.Write("			blnIgnoreRow = false;" & vbCrLf)
            Response.Write("			}" & vbCrLf)
		
            Response.Write("      if (!blnIgnoreRow) {" & vbCrLf)
            Response.Write("				lngActualRow = lngActualRow + 1; " & vbCrLf)
            Response.Write("        lngCol = 0;" & vbCrLf)
            Response.Write("        ClientDLL.ArrayReDim();" & vbCrLf & vbCrLf)
            Response.Write("        for (var lngCount=0; lngCount<frmOutput.ssOleDBGridDefSelRecords.Columns.Count; lngCount++)" & vbCrLf)
            Response.Write("					{" & vbCrLf)
            Response.Write("          if (frmOutput.ssOleDBGridDefSelRecords.Columns(lngCount).Visible == true)" & vbCrLf)
            Response.Write("						{" & vbCrLf)
            Response.Write("            ClientDLL.ArrayAddTo(lngCol, lngActualRow, frmOutput.ssOleDBGridDefSelRecords.Columns(lngCount).CellText(bm));" & vbCrLf)
            Response.Write("            lngCol++;" & vbCrLf)
            Response.Write("						}" & vbCrLf)
            Response.Write("					}" & vbCrLf)
            Response.Write("				}" & vbCrLf)
            Response.Write("			}	// end (no page break) " & vbCrLf)
            Response.Write("			ClientDLL.DataArray();" & vbCrLf)
			
            Response.Write("		}" & vbCrLf)
            Response.Write("		}" & vbCrLf)
            Response.Write("	}" & vbCrLf)
            Response.Write("    ClientDLL.Complete();" & vbCrLf)

            Response.Write("	ShowDataFrame();" & vbCrLf)
            Response.Write("  }" & vbCrLf)

            If Not objReport.OutputPreview Then
                Response.Write("  frmError.txtEventLogID.value = """ & CleanStringForJavaScript(objReport.EventLogID) & """;" & vbCrLf)
                Response.Write("  if (frmOriginalDefinition.txtCancelPrint.value == 1) {" & vbCrLf)
                Response.Write("    raiseError('',false,true);" & vbCrLf)
                Response.Write("  }" & vbCrLf)
                Response.Write("  else if (ClientDLL.ErrorMessage != """") {" & vbCrLf)
                Response.Write("    raiseError(ClientDLL.ErrorMessage,false,false);" & vbCrLf)
                Response.Write("  }" & vbCrLf)
                Response.Write("  else {" & vbCrLf)
                Response.Write("    raiseError('',true,false);" & vbCrLf)
                Response.Write("  }" & vbCrLf)
            Else
                Response.Write("  sUtilTypeDesc = frames(""top"").frmPopup.txtUtilTypeDesc.value;" & vbCrLf)
                Response.Write("  if (frmOriginalDefinition.txtCancelPrint.value == 1) {" & vbCrLf)
                Response.Write("    OpenHR.messageBox(sUtilTypeDesc+"" output failed.\n\nCancelled by user."",64,sUtilTypeDesc);" & vbCrLf)
                Response.Write("  }" & vbCrLf)
                Response.Write("  else if (ClientDLL.ErrorMessage != """") {" & vbCrLf)
                Response.Write("    OpenHR.messageBox(sUtilTypeDesc+"" output failed.\n\n"" + ClientDLL.ErrorMessage,48,sUtilTypeDesc);" & vbCrLf)
                Response.Write("  }" & vbCrLf)
                Response.Write("  else {" & vbCrLf)
                Response.Write("    OpenHR.messageBox(sUtilTypeDesc+"" output complete."",64,sUtilTypeDesc);" & vbCrLf)
                Response.Write("  }" & vbCrLf)
            End If

            Response.Write("  }" & vbCrLf)
    Response.Write("</script>" & vbCrLf & vbCrLf)
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

        objReport.ClearUp()

        If fok Then
            Response.Write("<FORM name=frmOutput id=frmOutput method=post>" & vbCrLf)
            Response.Write("<table height=100% width=100% align=center class=""outline"" cellPadding=5 cellSpacing=0 >" & vbCrLf)
            Response.Write("	<TR>" & vbCrLf)
            Response.Write("		<TD>" & vbCrLf)
            Response.Write("			<table name=tblGrid id=tblGrid height=100% width=100% class=""invisible"" cellspacing=0 cellpadding=0>" & vbCrLf)
            Response.Write("				<tr>" & vbCrLf)
            Response.Write("					<td colspan=12 height=10></td>" & vbCrLf)
            Response.Write("				</tr>" & vbCrLf)
            Response.Write("				<tr>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("					<td ALIGN=center colspan=10 NAME='tdOutputMSG' ID='tdOutputMSG'>" & vbCrLf)

            For icount = 1 To UBound(arrayDefinition)
                Response.Write(arrayDefinition(icount))
            Next

            For icount = 1 To UBound(arrayColumnsDefinition)
                Response.Write(arrayColumnsDefinition(icount))
            Next

            'for iCount = 1 to UBound(arrayDataDefinition)
            '	if instr(arrayDataDefinition(icount),"<PARAM NAME=") then
            '		Response.Write "    " & arrayDataDefinition(icount) & vbcrlf
            '	end if
            'next 

            Response.Write("						</OBJECT>" & vbCrLf)

            Response.Write("					</td>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("				</tr>" & vbCrLf)
            Response.Write("				<tr>" & vbCrLf)
            Response.Write("					<td colspan=12 height=10></td>" & vbCrLf)
            Response.Write("				</tr>" & vbCrLf)

            Response.Write("				<tr height=25>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("					<td colspan=8>" & vbCrLf)
            Response.Write("						<TABLE WIDTH=""100%"" class=""invisible"" CELLSPACING=0 CELLPADDING=0>" & vbCrLf)
            Response.Write("							<TR>" & vbCrLf)
		
            ' Put the hidden grid (used for printing with page breaks)
            ' in here as we need it to be in the BODY (so that it picks up on the page font)
            ' but setting visibility to 'hidden' caused the printing to crash.
            Response.Write("								<TD>" & vbCrLf)
            Response.Write("									<OBJECT classid=""clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"" codebase=""cabs/COAInt_Grid.cab#version=3,1,3,6"" id=ssHiddenGrid name=ssHiddenGrid style=""HEIGHT: 0px; LEFT: 0px; TOP: 0px; WIDTH: 0px; POSITION: absolute"" VIEWASTEXT>" & vbCrLf)
            Response.Write("										<PARAM NAME=""ScrollBars"" VALUE=""4"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""_Version"" VALUE=""196617"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""DataMode"" VALUE=""2"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Cols"" VALUE=""1"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Rows"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""BorderStyle"" VALUE=""1"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""RecordSelectors"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""GroupHeaders"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""ColumnHeaders"" VALUE=""-1"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""GroupHeadLines"" VALUE=""1"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""HeadLines"" VALUE=""1"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""FieldDelimiter"" VALUE=""(None)"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""FieldSeparator"" VALUE=""(Tab)"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Row.Count"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Col.Count"" VALUE=""1"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""stylesets.count"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""TagVariant"" VALUE=""EMPTY"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""UseGroups"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""HeadFont3D"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Font3D"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""DividerType"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""DividerStyle"" VALUE=""1"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""DefColWidth"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""BeveColorScheme"" VALUE=""2"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""BevelColorFrame"" VALUE=""-2147483642"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""BevelColorHighlight"" VALUE=""-2147483628"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""BevelColorShadow"" VALUE=""-2147483632"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""BevelColorFace"" VALUE=""-2147483633"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""CheckBox3D"" VALUE=""-1"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""AllowAddNew"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""AllowDelete"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""AllowUpdate"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""MultiLine"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""ActiveCellStyleSet"" VALUE="""">" & vbCrLf)
            Response.Write("										<PARAM NAME=""RowSelectionStyle"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""AllowRowSizing"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""AllowGroupSizing"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""AllowColumnSizing"" VALUE=""-1"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""AllowGroupMoving"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""AllowColumnMoving"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""AllowGroupSwapping"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""AllowColumnSwapping"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""AllowGroupShrinking"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""AllowColumnShrinking"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""AllowDragDrop"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""UseExactRowCount"" VALUE=""-1"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""SelectTypeCol"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""SelectTypeRow"" VALUE=""2"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""SelectByCell"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""BalloonHelp"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""RowNavigation"" VALUE=""1"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""CellNavigation"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""MaxSelectedRows"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""HeadStyleSet"" VALUE="""">" & vbCrLf)
            Response.Write("										<PARAM NAME=""StyleSet"" VALUE="""">" & vbCrLf)
            Response.Write("										<PARAM NAME=""ForeColorEven"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""ForeColorOdd"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""BackColorEven"" VALUE=""16777215"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""BackColorOdd"" VALUE=""16777215"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Levels"" VALUE=""1"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""RowHeight"" VALUE=""238"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""ExtraHeight"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""ActiveRowStyleSet"" VALUE="""">" & vbCrLf)
            Response.Write("										<PARAM NAME=""CaptionAlignment"" VALUE=""2"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""SplitterPos"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""SplitterVisible"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns.Count"" VALUE=""1"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).Width"" VALUE=""1000"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).Visible"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).Columns.Count"" VALUE=""1"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).Caption"" VALUE=""PageBreak"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).Name"" VALUE=""PageBreak"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).Alignment"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).CaptionAlignment"" VALUE=""2"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).Bound"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).AllowSizing"" VALUE=""1"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).DataField"" VALUE=""Column 0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).DataType"" VALUE=""8"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).Level"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).NumberFormat"" VALUE="""">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).Case"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).FieldLen"" VALUE=""4096"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).VertScrollBar"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).Locked"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).Style"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).ButtonsAlways"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).RowCount"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).ColCount"" VALUE=""1"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).HasHeadForeColor"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).HasHeadBackColor"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).HasForeColor"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).HasBackColor"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).HeadForeColor"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).HeadBackColor"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).ForeColor"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).BackColor"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).HeadStyleSet"" VALUE="""">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).StyleSet"" VALUE="""">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).Nullable"" VALUE=""1"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).Mask"" VALUE="""">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).PromptInclude"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).ClipMode"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Columns(0).PromptChar"" VALUE=""95"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""UseDefaults"" VALUE=""-1"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""TabNavigation"" VALUE=""1"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""BatchUpdate"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""_ExtentX"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""_ExtentY"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""_StockProps"" VALUE=""79"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Caption"" VALUE="""">" & vbCrLf)
            Response.Write("										<PARAM NAME=""ForeColor"" VALUE=""0"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""BackColor"" VALUE=""16777215"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""Enabled"" VALUE=""-1"">" & vbCrLf)
            Response.Write("										<PARAM NAME=""DataMember"" VALUE="""">" & vbCrLf)
            Response.Write("									</OBJECT>" & vbCrLf)
            Response.Write("								</TD>" & vbCrLf)
            Response.Write("								<TD>&nbsp;</TD>" & vbCrLf)
            Response.Write("								<td width=20>" & vbCrLf)
            Response.Write("      						<input type=button id=output name=output value=Output style=""WIDTH: 80px"" class=""btn""" & vbCrLf)
            Response.Write("                            onclick=""ExportDataPrompt();""" & vbCrLf)
            Response.Write("                            onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
            Response.Write("                            onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
            Response.Write("                            onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
            Response.Write("                            onblur=""try{button_onBlur(this);}catch(e){}"" />" & vbCrLf)
            Response.Write("								</td>" & vbCrLf)
            Response.Write("							</tr>" & vbCrLf)
            Response.Write("						</table>" & vbCrLf)
            Response.Write("					</td>" & vbCrLf)
            Response.Write("					<td width=10></td>" & vbCrLf)
            Response.Write("					<td width=80> " & vbCrLf)
            Response.Write("      						<input type=button id=close name=close value=Close style=""WIDTH: 80px"" class=""btn""" & vbCrLf) '2
            Response.Write("                            onclick=""closeclick();""" & vbCrLf)
            Response.Write("                            onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
            Response.Write("                            onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
            Response.Write("                            onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
            Response.Write("                            onblur=""try{button_onBlur(this);}catch(e){}"" />" & vbCrLf)
            Response.Write("					</td>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("				</tr>" & vbCrLf)
            Response.Write("				<tr>" & vbCrLf)
            Response.Write("					<td colspan=12 height=10></td>" & vbCrLf)
            Response.Write("				</tr>" & vbCrLf)
            Response.Write("			</table>" & vbCrLf)
            Response.Write("		</td>" & vbCrLf)
            Response.Write("	</tr>" & vbCrLf)
            Response.Write("</table>" & vbCrLf)
            Response.Write("</FORM>" & vbCrLf)
  	
            Response.Write("<INPUT type='hidden' id=txtNoRecs name=txtNoRecs value=0>" & vbCrLf)
            Response.Write("<input type=hidden id=txtSuccessFlag name=txtSuccessFlag value=2>" & vbCrLf)
          Else%>
	
		<FORM Name=frmPopup ID=frmPopup>
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
                            
                            
<%	if bBradfordFactor = true then
			mstrCaption = "Bradford Factor"
		else
			mstrCaption = "Custom Report '" & session("utilname") & "'"
    End If

    If fNoRecords Then
        Response.Write("						<H4>" & mstrCaption & " Completed successfully.</H4>" & vbCrLf)
    Else
        Response.Write("						<H4>" & mstrCaption & " Failed." & vbCrLf)
    End If
%>
					    </td>
					    <td width=20></td>
					  </tr>
					  <tr>
					    <td width=20 height=10></td>
					    <td align=center nowrap><%=objReport.ErrorString%>
					    </td>
					    <td width=20></td>
					  </tr>
					  <tr>
					    <td colspan=3 height=10>&nbsp;</td>
					  </tr>
					  <tr> 
					    <td colspan=3 height=10 align=center>
                            <input type="button" id="cmdClose" name="cmdClose" value="Close" style="WIDTH: 80px" width="80" class="btn"
                                onclick="closeclick();" />
                        </td>
					  </tr>
					  <tr> 
					    <td colspan=3 height=10></td>
					  </tr>
					</table>
				</td>
			</tr>
		</table>
		</FORM>

<input type='hidden' id="txtNoRecs" name="txtNoRecs" value="1">
<input type="hidden" id="txtSuccessFlag" name="txtSuccessFlag" value="3">
<%
	end if
%>



    <form id="frmOriginalDefinition" style="visibility: hidden; display: none">
        <%
            Response.Write("	<INPUT type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Replace(Session("utilname"), """", "&quot;") & """>" & vbCrLf)
            Response.Write("	<INPUT type='hidden' id=txtDefn_ErrMsg name=txtDefn_ErrMsg value=""" & sErrMsg & """>" & vbCrLf)
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
    </form>

    <%
        Response.Write("<INPUT type=""hidden"" id=txtDatabase name=txtDatabase value=""" & Replace(Session("Database"), """", "&quot;") & """>")
    %>
