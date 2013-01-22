<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<% 
    Dim bBradfordFactor As Boolean
    Dim mstrCaption As String
    Dim sErrMsg As String
    
	bBradfordFactor = (session("utiltype") = "16")
%>


    <link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css" />
    <script src="<%: Url.Content("~/Scripts/jquery-1.8.2.js") %>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/Scripts/openhr.js") %>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/Scripts/ctl_SetFont.js") %>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/Scripts/ctl_SetStyles.js") %>" type="text/javascript"></script>

    <script src="<%: Url.Content("~/Scripts/jquery-ui-1.9.1.custom.min.js") %>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/Scripts/jquery.cookie.js") %>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/Scripts/menu.js")%>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/Scripts/jquery.ui.touch-punch.min.js") %>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/Scripts/jsTree/jquery.jstree.js") %>" type="text/javascript"></script>
    <script id="officebarscript" src="<%: Url.Content("~/Scripts/officebar/jquery.officebar.js") %>" type="text/javascript"></script>

    <object
        classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
        id="Microsoft_Licensed_Class_Manager_1_0"
        viewastext>
        <param name="LPKPath" value="lpks/main.lpk">
    </object>
    
    <script type="text/javascript">
        function customreports_window_onload() {
            $("#workframe").attr("data-framesource", "UTIL_RUN_CUSTOMREPORTS");
            loadAddRecords();
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
            Response.Write("                      onclick=""window.parent.parent.parent.self.close();""" & vbCrLf)
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

	dim icount
	dim definition
	dim fok
	dim objReport
	dim fNotCancelled

	dim dtStartDate
	dim dtEndDate
	dim strAbsenceTypes
	dim lngFilterID
	dim lngPicklistID
	dim lngPersonnelID
	
	dim bBradford_SRV
	dim bBradford_ShowDurations
	dim bBradford_ShowInstances
	dim bBradford_ShowFormula	
	dim bBradford_OmitBeforeStart	
	dim bBradford_OmitAfterEnd	
	dim bBradford_txtOrderBy1
	dim lngBradford_txtOrderBy1ID	
	dim bBradford_txtOrderBy1Asc
	dim bBradford_txtOrderBy2
	dim lngBradford_txtOrderBy2ID	
	dim bBradford_txtOrderBy2Asc
	dim bPrintFilterPickList

	' Default output options
	dim bOutputPreview
	dim lngOutputFormat
	dim pblnOutputScreen
	dim pblnOutputPrinter
	dim pstrOutputPrinterName
	dim pblnOutputSave
	dim plngOutputSaveExisting
	dim pblnOutputEmail
	dim plngOutputEmailID
	dim pstrOutputEmailName
	dim pstrOutputEmailSubject
	dim pstrOutputEmailAttachAs
	dim pstrOutputFilename	

        Dim bMinBradford As Boolean
        Dim lngMinBradfordAmount As Long
        Dim pbDisplayBradfordDetail As Boolean
        
	fok = true
	fNotCancelled = true

	' Create the reference to the DLL (Report Class)
        objReport = CreateObject("COAIntServer.Report")

	' Pass required info to the DLL
	objReport.Username = session("username")
        CallByName(objReport, "Connection", CallType.Let, Session("databaseConnection"))

	objReport.CustomReportID = session("utilid")
	objReport.ClientDateFormat = session("LocaleDateFormat")
	objReport.LocalDecimalSeparator = session("LocaleDecimalSeparator")

	if fok and bBradfordFactor then
            'dtStartDate = convertLocaleDateToSQL(session("stdReport_StartDate"))
            '   dtEndDate = convertLocaleDateToSQL(Session("stdReport_EndDate"))
            'TODO convertdate formats server side
            dtStartDate = ""
            dtEndDate = ""
            
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

	'**********************************************
	'TM20020809 Fault 4237 - check that at least one child table is used.
	if fok and objReport.ChildCount > 1 and objReport.UsedChildCount > 1 then
		fok = objReport.CreateMutipleChildTempTable
		fNotCancelled = Response.IsClientConnected
		if fok then fok = fNotCancelled
	end if
	'**********************************************
	
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
		
		if fok then
			'arrayDataDefinition = objreport.OutputArray_Data 

		  'Response.Write "		<FORM name=frmGridItems id=frmGridItems>" & vbcrlf
			'for iCount = 1 to UBound(arrayDataDefinition)
			'	if instr(arrayDataDefinition(icount),"<PARAM NAME=") = 0 then
			'		sGridItemName = "txtGridItem_" & iCount
			'	  Response.Write "			<input type=hidden id=" & sGridItemName & " name=" & sGridItemName & " value=""" & arrayDataDefinition(icount) & """>" & vbcrlf
			'	end if
			'next 
		  'Response.Write "		</FORM>" & vbcrlf
		  
                Response.Write(objReport.Output_GridForm)
            End If
        End If

	if fok then

            Response.Write("<FORM target=""Output"" action=""util_run_outputoptions.asp"" method=post id=frmExportData name=frmExportData>" & vbCrLf)
            Response.Write("  <INPUT type=""hidden"" id=txtPreview name=txtPreview value=""" & objReport.OutputPreview & """>" & vbCrLf)
            Response.Write("  <INPUT type=""hidden"" id=txtFormat name=txtFormat value=""" & objReport.OutputFormat & """>" & vbCrLf)
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
            Response.Write("<!--" & vbCrLf)
            
            ' Change the output text if Bradford Factor Report

            Response.Write("function ExportData(strMode) " & vbCrLf)
            Response.Write("	{" & vbCrLf & vbCrLf)
            Response.Write("	var bm;" & vbCrLf)
            Response.Write("	var sBreakValue = new String('');" & vbCrLf)
            Response.Write("	var blnBreakCheck = false;" & vbCrLf)
		
            Dim objUser
            objUser = CreateObject("COAIntServer.clsSettings")
            objReport.Username = Session("username")
            CallByName(objUser, "Connection", CallType.Let, Session("databaseConnection"))
            
            
            'MH20031113 Fault 7606 Reset Columns and Styles...
            Response.Write("  window.parent.parent.ASRIntranetOutput.ResetColumns();" & vbCrLf)
            Response.Write("  window.parent.parent.ASRIntranetOutput.ResetStyles();" & vbCrLf)
            Response.Write("  window.parent.parent.ASRIntranetOutput.UserName = """ & cleanStringForJavaScript(Session("Username")) & """;" & vbCrLf)
            Response.Write("  window.parent.parent.ASRIntranetOutput.SaveAsValues = """ & cleanStringForJavaScript(Session("OfficeSaveAsValues")) & """;" & vbCrLf)

            Response.Write("  window.parent.parent.ASRIntranetOutput.SettingLocations(")
            Response.Write(cleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleCol", "3")) & ", ")
            Response.Write(cleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleRow", "2")) & ", ")
            Response.Write(cleanStringForJavaScript(objUser.GetUserSetting("Output", "DataCol", "2")) & ", ")
            Response.Write(cleanStringForJavaScript(objUser.GetUserSetting("Output", "DataRow", "4")) & ");" & vbCrLf)

            Response.Write("  window.parent.parent.ASRIntranetOutput.SettingTitle(")
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

            Response.Write(cleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleBackcolour", "16777215")) & ", ")
            Response.Write(cleanStringForJavaScript(objUser.GetUserSetting("Output", "TitleForecolour", "6697779")) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "TitleBackcolour", "16777215")))) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "TitleForecolour", "6697779")))) & ");" & vbCrLf)

            Response.Write("  window.parent.parent.ASRIntranetOutput.SettingHeading(")
            Response.Write(cleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingGridLines", "1")) & ", ")
            Response.Write(cleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingBold", "1")) & ", ")
            Response.Write(cleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingUnderline", "0")) & ", ")
            Response.Write(cleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingBackcolour", "16248553")) & ", ")
            Response.Write(cleanStringForJavaScript(objUser.GetUserSetting("Output", "HeadingForecolour", "6697779")) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "HeadingBackcolour", "16248553")))) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "HeadingForecolour", "6697779")))) & ");" & vbCrLf)

            Response.Write("  window.parent.parent.ASRIntranetOutput.SettingData(")
            Response.Write(cleanStringForJavaScript(objUser.GetUserSetting("Output", "DataGridLines", "1")) & ", ")
            Response.Write(cleanStringForJavaScript(objUser.GetUserSetting("Output", "DataBold", "0")) & ", ")
            Response.Write(cleanStringForJavaScript(objUser.GetUserSetting("Output", "DataUnderline", "0")) & ", ")
            Response.Write(cleanStringForJavaScript(objUser.GetUserSetting("Output", "DataBackcolour", "15988214")) & ", ")
            Response.Write(cleanStringForJavaScript(objUser.GetUserSetting("Output", "DataForecolour", "6697779")) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "DataBackcolour", "15988214")))) & ", ")
            Response.Write(CleanStringForJavaScript(objUser.GetWordColourIndex(CLng(objUser.GetUserSetting("Output", "DataForecolour", "6697779")))) & ");" & vbCrLf)

            Response.Write("  frmMenuFrame = window.parent.parent.opener.window.parent.frames(""menuframe"");" & vbCrLf)

            Response.Write("  window.parent.parent.ASRIntranetOutput.InitialiseStyles();" & vbCrLf)
            Response.Write("  window.parent.parent.ASRIntranetOutput.SettingOptions(")
            Response.Write("""" & cleanStringForJavaScript(objUser.GetUserSetting("Output", "WordTemplate", "")) & """, ")
            Response.Write("""" & cleanStringForJavaScript(objUser.GetUserSetting("Output", "ExcelTemplate", "")) & """, ")

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

            Response.Write("frmMenuFrame.document.all.item(""txtSysPerm_EMAILGROUPS_VIEW"").value);" & vbCrLf)

            'Set Options
            Response.Write("  frmDataFrame = window.parent.frames(""dataframe"");" & vbCrLf)

		if not objreport.OutputPreview then	
			dim lngFormat
			dim blnScreen
			dim blnPrinter
			dim strPrinterName
			dim blnSave
			dim lngSaveExisting
			dim blnEmail
			dim lngEmailGroupID
			dim strEmailSubject
			dim strEmailAttachAs
			dim strFileName
			
			lngFormat = cleanStringForJavaScript(objreport.OutputFormat)
			blnScreen = cleanStringForJavaScript(LCase(objreport.OutputScreen))
			blnPrinter = cleanStringForJavaScript(LCase(objreport.OutputPrinter))
			strPrinterName = cleanStringForJavaScript(objreport.OutputPrinterName)
			blnSave = cleanStringForJavaScript(LCase(objreport.OutputSave))
			lngSaveExisting = cleanStringForJavaScript(objreport.OutputSaveExisting) 
			blnEmail = cleanStringForJavaScript(LCase(objreport.OutputEmail))
			lngEmailGroupID = CLng(objreport.OutputEmailID)
			strEmailSubject = cleanStringForJavaScript(objreport.OutputEmailSubject)
			strEmailAttachAs = cleanStringForJavaScript(objreport.OutputEmailAttachAs)
			'strFileName = objreport.OutputFilename 
			strFileName = cleanStringForJavaScript(objreport.OutputFilename)
			
                Dim cmdEmailAddr
                Dim prmEmailGroupID
                Dim rstEmailAddr
                Dim sErrorDescription
                Dim iLoop As Integer
                Dim sEmailAddresses As String
                
			if (objreport.OutputEmail) and (objreport.OutputEmailID > 0) then
				
                    cmdEmailAddr = Server.CreateObject("ADODB.Command")
				cmdEmailAddr.CommandText = "spASRIntGetEmailGroupAddresses"
				cmdEmailAddr.CommandType = 4 ' Stored procedure
                    cmdEmailAddr.ActiveConnection = Session("databaseConnection")

                    prmEmailGroupID = cmdEmailAddr.CreateParameter("EmailGroupID", 3, 1) ' 3=integer, 1=input
                    cmdEmailAddr.Parameters.Append(prmEmailGroupID)
				prmEmailGroupID.value = cleanNumeric(lngEmailGroupID)

                    Err.Clear()
                    rstEmailAddr = cmdEmailAddr.Execute

                    If (Err.Number <> 0) Then
                        sErrorDescription = "Error getting the email addresses for group." & vbCrLf & FormatError(Err.Description)
                    End If

				if len(sErrorDescription) = 0 then
					iLoop = 1
					do while not rstEmailAddr.EOF
						if iLoop > 1 then
							sEmailAddresses = sEmailAddresses & ";"
						end if
						sEmailAddresses = sEmailAddresses & rstEmailAddr.Fields("Fixed").Value
						rstEmailAddr.MoveNext
						iLoop = iLoop + 1
					loop
						
					' Release the ADO recordset object.
					rstEmailAddr.close
				end if
							
                    rstEmailAddr = Nothing
                    cmdEmailAddr = Nothing
			end if
			
                Response.Write("  fok = window.parent.parent.ASRIntranetOutput.SetOptions(false, " & _
                                                        lngFormat & "," & blnScreen & ", " & _
                                                        blnPrinter & ",""" & strPrinterName & """, " & _
                                                        blnSave & "," & lngSaveExisting & ", " & _
                                                        blnEmail & ", """ & cleanStringForJavaScript(sEmailAddresses) & """, """ & _
                                                        strEmailSubject & """,""" & strEmailAttachAs & """,""" & strFileName & """);" & vbCrLf)
            Else
                Response.Write("  fok = window.parent.parent.ASRIntranetOutput.SetOptions(false, " & _
                                                "frmExportData.txtFormat.value, frmExportData.txtScreen.value, " & _
                                                "frmExportData.txtPrinter.value, frmExportData.txtPrinterName.value, " & _
                                                "frmExportData.txtSave.value, frmExportData.txtSaveExisting.value, " & _
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
		
            Response.Write("  window.parent.parent.ASRIntranetOutput.SizeColumnsIndependently = true;" & vbCrLf)

            Response.Write("  if (frmExportData.txtFormat.value == ""0"") {" & vbCrLf)
            Response.Write("    if (frmExportData.txtPrinter.value.toLowerCase() == ""true"") {" & vbCrLf)
            Response.Write("      window.parent.parent.ASRIntranetOutput.SetPrinter();" & vbCrLf)
            Response.Write("      dataOnlyPrint();" & vbCrLf)
            Response.Write("      window.parent.parent.ASRIntranetOutput.ResetDefaultPrinter();" & vbCrLf)
            Response.Write("    }" & vbCrLf)
            Response.Write("  }" & vbCrLf)
            Response.Write("  else {" & vbCrLf)
            Response.Write("  window.parent.parent.ASRIntranetOutput.HeaderRows = 1;" & vbCrLf)

            Response.Write("  if (window.parent.parent.ASRIntranetOutput.GetFile() == true) " & vbCrLf)
            Response.Write("		{" & vbCrLf)
			
            'Response.Write "		if (frmExportData.pagebreak.value == 'True') " & vbcrlf
            Response.Write("		if (frmExportData.pagebreak.value.toLowerCase() == ""true"") " & vbCrLf)
            Response.Write("			{" & vbCrLf)
            Response.Write("			var lngActualRow = new Number(0);" & vbCrLf)
            '	Response.Write "			window.parent.parent.ASRIntranetOutput.PageTitles = true;" & vbcrlf
			
            Response.Write("			window.parent.parent.ASRIntranetOutput.ArrayDim(frmExportData.txtVisColCount.value, 0);" & vbCrLf & vbCrLf)
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
                Response.Write("					window.parent.parent.ASRIntranetOutput.AddPage(replace(frmOutput.ssOleDBGridDefSelRecords.Caption,'&&','&'),'Bradford Factor');" & vbCrLf)
            Else
                Response.Write("					window.parent.parent.ASRIntranetOutput.AddPage(replace(frmOutput.ssOleDBGridDefSelRecords.Caption,'&&','&'),sBreakValue);" & vbCrLf)
            End If
            Response.Write("					var sColHeading = new String(''); " & vbCrLf)
            Response.Write("					var iColDataType = new Number(0); " & vbCrLf)
            Response.Write("					var iColDecimals = new Number(0); " & vbCrLf)
            Response.Write("					for (var lngCol = 0; lngCol<=frmExportData.txtVisColCount.value; lngCol++)" & vbCrLf)
            Response.Write("						{" & vbCrLf)
            Response.Write("						sColHeading = document.getElementById('txtVisHeading_'+lngCol).value;" & vbCrLf)
            Response.Write("						iColDataType = document.getElementById('txtVisDataType_'+lngCol).value;" & vbCrLf)
            Response.Write("						iColDecimals = document.getElementById('txtVisDecimals_'+lngCol).value;" & vbCrLf)
            Response.Write("						window.parent.parent.ASRIntranetOutput.AddColumn(sColHeading, iColDataType, iColDecimals, false);" & vbCrLf)
            Response.Write("						window.parent.parent.ASRIntranetOutput.ArrayAddTo(lngCol, 0, sColHeading);" & vbCrLf & vbCrLf)
            Response.Write("						}" & vbCrLf)
            Response.Write("                    window.parent.parent.ASRIntranetOutput.DataArray();" & vbCrLf)
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
                Response.Write("					window.parent.parent.ASRIntranetOutput.AddPage(replace(frmOutput.ssOleDBGridDefSelRecords.Caption,'&&','&'),'Bradford Factor');" & vbCrLf)
            Else
                Response.Write("					window.parent.parent.ASRIntranetOutput.AddPage(replace(frmOutput.ssOleDBGridDefSelRecords.Caption,'&&','&'),sBreakValue);" & vbCrLf)
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
            Response.Write("						window.parent.parent.ASRIntranetOutput.AddColumn(sColHeading, iColDataType, iColDecimals, iCol1000);" & vbCrLf)
            Response.Write("						window.parent.parent.ASRIntranetOutput.ArrayAddTo(lngCol, 0, sColHeading);" & vbCrLf & vbCrLf)
            Response.Write("						}" & vbCrLf)

            Response.Write("          window.parent.parent.ASRIntranetOutput.DataArray();" & vbCrLf)
            Response.Write("					window.parent.parent.ASRIntranetOutput.ArrayDim(frmExportData.txtVisColCount.value, 0);" & vbCrLf & vbCrLf)
            Response.Write("					lngActualRow = 0;" & vbCrLf)
            Response.Write("					blnBreakCheck = true;" & vbCrLf)
            Response.Write("					sBreakValue = '';" & vbCrLf)
            Response.Write("					window.parent.parent.ASRIntranetOutput.ResetColumns();" & vbCrLf)
            Response.Write("					window.parent.parent.ASRIntranetOutput.ResetStyles();" & vbCrLf)
			
            Response.Write("					}" & vbCrLf & vbCrLf)
            Response.Write("        else if (frmOutput.ssOleDBGridDefSelRecords.Columns(0).CellText(bm) != '*')" & vbCrLf)
            Response.Write("					{" & vbCrLf)
			
            Response.Write("					blnBreakCheck = false;" & vbCrLf)
            Response.Write("					lngCol = 0;" & vbCrLf)
            Response.Write("					window.parent.parent.ASRIntranetOutput.ArrayReDim();" & vbCrLf & vbCrLf)
            Response.Write("					for (var lngCount=0; lngCount<frmOutput.ssOleDBGridDefSelRecords.Columns.Count; lngCount++)" & vbCrLf)
            Response.Write("						{" & vbCrLf)
            Response.Write("						if (frmOutput.ssOleDBGridDefSelRecords.Columns(lngCount).Visible == true)" & vbCrLf)
            Response.Write("							{" & vbCrLf)
            Response.Write("						  window.parent.parent.ASRIntranetOutput.ArrayAddTo(lngCol, lngActualRow, frmOutput.ssOleDBGridDefSelRecords.Columns(lngCount).CellText(bm));" & vbCrLf)
            Response.Write("						  lngCol++;" & vbCrLf)
            Response.Write("							}" & vbCrLf)
            Response.Write("						}" & vbCrLf)

            Response.Write("					}" & vbCrLf)
            Response.Write("				}" & vbCrLf)
            Response.Write("			} " & vbCrLf)
            Response.Write("		else // no page break " & vbCrLf)
            Response.Write("			{ " & vbCrLf)

            Response.Write("      window.parent.parent.ASRIntranetOutput.ArrayDim(frmExportData.txtVisColCount.value, 0);" & vbCrLf & vbCrLf)
            If bBradfordFactor = True Then
                Response.Write("			window.parent.parent.ASRIntranetOutput.PageTitles = false;" & vbCrLf)
                Response.Write("      window.parent.parent.ASRIntranetOutput.AddPage(replace(frmOutput.ssOleDBGridDefSelRecords.Caption,'&&','&'),'" & "Bradford Factor" & "');" & vbCrLf)
            Else
                Response.Write("      window.parent.parent.ASRIntranetOutput.AddPage(replace(frmOutput.ssOleDBGridDefSelRecords.Caption,'&&','&'),'" & cleanStringForJavaScript(objReport.BaseTableName) & "');" & vbCrLf)
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
            Response.Write("				window.parent.parent.ASRIntranetOutput.AddColumn(sColHeading, iColDataType, iColDecimals, iCol1000);" & vbCrLf)
            Response.Write("				window.parent.parent.ASRIntranetOutput.ArrayAddTo(lngCol, 0, sColHeading);" & vbCrLf & vbCrLf)
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
            Response.Write("        window.parent.parent.ASRIntranetOutput.ArrayReDim();" & vbCrLf & vbCrLf)
            Response.Write("        for (var lngCount=0; lngCount<frmOutput.ssOleDBGridDefSelRecords.Columns.Count; lngCount++)" & vbCrLf)
            Response.Write("					{" & vbCrLf)
            Response.Write("          if (frmOutput.ssOleDBGridDefSelRecords.Columns(lngCount).Visible == true)" & vbCrLf)
            Response.Write("						{" & vbCrLf)
            Response.Write("            window.parent.parent.ASRIntranetOutput.ArrayAddTo(lngCol, lngActualRow, frmOutput.ssOleDBGridDefSelRecords.Columns(lngCount).CellText(bm));" & vbCrLf)
            Response.Write("            lngCol++;" & vbCrLf)
            Response.Write("						}" & vbCrLf)
            Response.Write("					}" & vbCrLf)
            Response.Write("				}" & vbCrLf)
            Response.Write("			}	// end (no page break) " & vbCrLf)
            Response.Write("			window.parent.parent.ASRIntranetOutput.DataArray();" & vbCrLf)
			
            Response.Write("		}" & vbCrLf)
            Response.Write("		}" & vbCrLf)
            Response.Write("	}" & vbCrLf)
            Response.Write("    window.parent.parent.ASRIntranetOutput.Complete();" & vbCrLf)

            Response.Write("	window.parent.parent.ShowDataFrame();" & vbCrLf)
            Response.Write("  }" & vbCrLf)

            If Not objReport.OutputPreview Then
                Response.Write("  window.parent.parent.parent.frmError.txtEventLogID.value = """ & cleanStringForJavaScript(objReport.EventLogID) & """;" & vbCrLf)
                Response.Write("  if (frmOriginalDefinition.txtCancelPrint.value == 1) {" & vbCrLf)
                Response.Write("    window.parent.parent.raiseError('',false,true);" & vbCrLf)
                Response.Write("  }" & vbCrLf)
                Response.Write("  else if (window.parent.parent.ASRIntranetOutput.ErrorMessage != """") {" & vbCrLf)
                Response.Write("    window.parent.parent.raiseError(window.parent.parent.ASRIntranetOutput.ErrorMessage,false,false);" & vbCrLf)
                Response.Write("  }" & vbCrLf)
                Response.Write("  else {" & vbCrLf)
                Response.Write("    window.parent.parent.raiseError('',true,false);" & vbCrLf)
                Response.Write("  }" & vbCrLf)
            Else
                Response.Write("  sUtilTypeDesc = window.parent.parent.parent.frames(""top"").frmPopup.txtUtilTypeDesc.value;" & vbCrLf)
                Response.Write("  if (frmOriginalDefinition.txtCancelPrint.value == 1) {" & vbCrLf)
                Response.Write("    window.parent.parent.ASRIntranetFunctions.MessageBox(sUtilTypeDesc+"" output failed.\n\nCancelled by user."",64,sUtilTypeDesc);" & vbCrLf)
                Response.Write("  }" & vbCrLf)
                Response.Write("  else if (window.parent.parent.ASRIntranetOutput.ErrorMessage != """") {" & vbCrLf)
                Response.Write("    window.parent.parent.ASRIntranetFunctions.MessageBox(sUtilTypeDesc+"" output failed.\n\n"" + window.parent.parent.ASRIntranetOutput.ErrorMessage,48,sUtilTypeDesc);" & vbCrLf)
                Response.Write("  }" & vbCrLf)
                Response.Write("  else {" & vbCrLf)
                Response.Write("    window.parent.parent.ASRIntranetFunctions.MessageBox(sUtilTypeDesc+"" output complete."",64,sUtilTypeDesc);" & vbCrLf)
                Response.Write("  }" & vbCrLf)
            End If

            Response.Write("  }" & vbCrLf)
            Response.Write("-->" & vbCrLf)
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
            Response.Write("                            onclick=""ExportDataPrompt(false);""" & vbCrLf)
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
                <input type=button id=cmdClose name=cmdClose value=Close style="WIDTH: 80px" width=80 class="btn"
                    onclick="window.parent.parent.parent.self.close()"
                                      onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                                      onmouseout="try{button_onMouseOut(this);}catch(e){}"
                                      onfocus="try{button_onFocus(this);}catch(e){}"
                                      onblur="try{button_onBlur(this);}catch(e){}" />
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

		<INPUT type='hidden' id=txtNoRecs name=txtNoRecs value=1>
	  <input type=hidden id=txtSuccessFlag name=txtSuccessFlag value=3>
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

	<INPUT type="hidden" id=txtCancelPrint name=txtCancelPrint>
	<INPUT type="hidden" id=txtOptionsDone name=txtOptionsDone>
	<INPUT type="hidden" id=txtOptionsPortrait name=txtOptionsPortrait>
	<INPUT type="hidden" id=txtOptionsMarginLeft name=txtOptionsMarginLeft>
	<INPUT type="hidden" id=txtOptionsMarginRight name=txtOptionsMarginRight>
	<INPUT type="hidden" id=txtOptionsMarginTop name=txtOptionsMarginTop>
	<INPUT type="hidden" id=txtOptionsMarginBottom name=txtOptionsMarginBottom>
	<INPUT type="hidden" id=txtOptionsCopies name=txtOptionsCopies>
    </form>

    <%
        Response.Write("<INPUT type=""hidden"" id=txtDatabase name=txtDatabase value=""" & Replace(Session("Database"), """", "&quot;") & """>")
    %>


<script type="text/javascript">
<!--
    
    function addActiveXHandlers() {
        debugger;

        OpenHR.addActiveXHandler("tblGrid", "onresize", tblGrid_onresize);
        OpenHR.addActiveXHandler("ssOleDBGridDefSelRecords", "PrintInitialize", ssOleDBGridDefSelRecords_PrintInitialize);
        OpenHR.addActiveXHandler("ssOleDBGridDefselRecords", "PrintError", ssOleDBGridDefselRecords_PrintError);

        OpenHR.addActiveXHandler("ssHiddenGrid", "PrintInitialize", ssHiddenGrid_PrintInitialize);
        OpenHR.addActiveXHandler("ssHiddenGrid", "PrintBegin", ssHiddenGrid_PrintBegin);
        OpenHR.addActiveXHandler("ssHiddenGrid", "PrintError", ssHiddenGrid_PrintError);
    }

    function tblGrid_onresize() {

        try {
            if (txtNoRecs.value == 0) {
                frmOutput.ssOleDBGridDefSelRecords.Refresh();
                if ((frmOutput.ssOleDBGridDefSelRecords.visiblerows() + 1) >= frmOutput.ssOleDBGridDefSelRecords.rows()) {
                    frmOutput.ssOleDBGridDefSelRecords.FirstRow = frmOutput.ssOleDBGridDefSelRecords.AddItemBookmark(0);
                }
            }
        } catch(e) {
            return;
        }
    }

    function ssOleDBGridDefSelRecords_PrintInitialize(ssPrintInfo) {
        //Underline headings if not printing gridlines
        ssPrintInfo.PrintGridlines = 3; // 0 = none, 3 = all
    
        ssPrintInfo.PrintHeaders = 0; // 0 = top of every page, 1 = top of report
        ssPrintInfo.Portrait = false; 
        ssPrintInfo.Copies = 1; 
        ssPrintInfo.Collate = true; 
        ssPrintInfo.PrintColors = true; 

        ssPrintInfo.RowAutoSize = true;
        ssPrintInfo.PrintColumnHeaders = 1;
        ssPrintInfo.MaxLinesPerRow = 2;

        ssPrintInfo.PageHeader = "	" + frmOutput.ssOleDBGridDefSelRecords.Caption + "	";
        ssPrintInfo.PageFooter = "Printed on <date> at <time> by " + frmOriginalDefinition.txtUserName.value + "	" + "	" + "Page <page number>";
    }

    function ssOleDBGridDefselRecords_PrintError(lngPrintError, iResponse){
        if (lngPrintError == 30457) {
            frmOriginalDefinition.txtCancelPrint.value = 1;
        }
    }   

    function ssHiddenGrid_PrintInitialize(ssPrintInfo) {
    
        //Underline headings if not printing gridlines
        ssPrintInfo.PrintGridlines = 3; // 0 = none, 3 = all
    
        ssPrintInfo.PrintHeaders = 0; // 0 = top of every page, 1 = top of report
        ssPrintInfo.Portrait = false; 
        ssPrintInfo.Copies = 1; 
        ssPrintInfo.Collate = true; 
        ssPrintInfo.PrintColors = true; 

        ssPrintInfo.RowAutoSize = true;
        ssPrintInfo.PrintColumnHeaders = 1;
        ssPrintInfo.MaxLinesPerRow = 2;

        ssPrintInfo.PageHeader = "	" + frmOutput.ssOleDBGridDefSelRecords.Caption + "	";
        ssPrintInfo.PageFooter = "Printed on <date> at <time> by " + frmOriginalDefinition.txtUserName.value + "	" + "	" + "Page <page number>";    
    }
        
    function ssHiddenGrid_PrintBegin(ssPrintInfo) {
    
        if (frmOriginalDefinition.txtOptionsDone.value == 0) 
        {
            frmOriginalDefinition.txtOptionsPortrait.value = ssPrintInfo.Portrait;
            frmOriginalDefinition.txtOptionsMarginLeft.value = ssPrintInfo.MarginLeft;
            frmOriginalDefinition.txtOptionsMarginRight.value = ssPrintInfo.MarginRight;
            frmOriginalDefinition.txtOptionsMarginTop.value = ssPrintInfo.MarginTop;
            frmOriginalDefinition.txtOptionsMarginBottom.value = ssPrintInfo.MarginBottom;
            frmOriginalDefinition.txtOptionsCopies.value = ssPrintInfo.Copies;
        }
        else 
        {
            ssPrintInfo.Portrait = frmOriginalDefinition.txtOptionsPortrait.value;
            ssPrintInfo.MarginLeft = frmOriginalDefinition.txtOptionsMarginLeft.value;
            ssPrintInfo.MarginRight = frmOriginalDefinition.txtOptionsMarginRight.value;
            ssPrintInfo.MarginTop = frmOriginalDefinition.txtOptionsMarginTop.value;
            ssPrintInfo.MarginBottom = frmOriginalDefinition.txtOptionsMarginBottom.value;
            ssPrintInfo.Copies = frmOriginalDefinition.txtOptionsCopies.value;
        }    
    }

    function ssHiddenGrid_PrintError(lngPrintError, iResponse){
        if (lngPrintError == 30457) {
            frmOriginalDefinition.txtCancelPrint.value = 1;
        }
    }
    

    -->
</script>





<script type="text/javascript">
<!--

    function ShowReport() 
    {
        var iPollPeriod;
        var iPollCounter;
        var iDummy;

        iPollPeriod = 100;
        iPollCounter = iPollPeriod;

        debugger;

        // Expand the work frame and hide the option frame.
        //window.parent.parent.document.all.item("myframeset").rows = "0, *";
	
        if ((txtSuccessFlag.value == 1) || (txtSuccessFlag.value == 3)) 
        {

            // Resize the popup.
            iResizeByHeight = frmPopup.offsetParent.scrollHeight - window.parent.parent.parent.document.body.clientHeight;
            if (frmPopup.offsetParent.offsetHeight + iResizeByHeight > screen.height) 
            {
                try
                {
                    window.parent.window.parent.moveTo((screen.width - window.parent.parent.parent.document.body.offsetWidth) / 2, 0);
                    window.parent.window.parent.resizeTo(window.parent.parent.parent.document.body.offsetWidth, screen.height);
                }
                catch(e) {}
            }
            else 
            {
                try
                {
                    window.parent.window.parent.moveTo((screen.width - window.parent.parent.parent.document.body.offsetWidth) / 2, (screen.height - (window.parent.parent.parent.document.body.offsetHeight + iResizeByHeight)) / 2);
                    window.parent.window.parent.resizeBy(0, iResizeByHeight);
                }
                catch(e) {}
            }

            iResizeByWidth = frmPopup.offsetParent.scrollWidth - window.parent.parent.parent.document.body.clientWidth;
            if (frmPopup.offsetParent.offsetWidth + iResizeByWidth > screen.width) 
            {
                try
                {
                    window.parent.window.parent.moveTo(0, (screen.height - window.parent.parent.parent.document.body.offsetHeight) / 2);
                    window.parent.window.parent.resizeTo(screen.width, window.parent.parent.parent.document.body.offsetHeight);
                }
                catch(e) {}
            }
            else
            {
                try
                {
                    window.parent.window.parent.moveTo((screen.width - (window.parent.parent.parent.document.body.offsetWidth + iResizeByWidth)) / 2, (screen.height - window.parent.parent.parent.document.body.offsetHeight) / 2);
                    window.parent.window.parent.resizeBy(iResizeByWidth, 0);
                }
                catch(e) {}
            }		
        }
        else 
        {
            if (txtSuccessFlag.value == 2) 
            {
                debugger;

                var frmOutput = OpenHR.getForm("workframe", "frmOutput");

                setGridFont(frmOutput.ssHiddenGrid);
                setGridFont(frmOutput.ssOleDBGridDefSelRecords);

//                frmOutput.ssOleDBGridDefSelRecords.style.visibility = 'hidden';
  //              frmOutput.ssOleDBGridDefSelRecords.Redraw = false;
    //            frmOutput.ssOleDBGridDefSelRecords.style.visibility = 'visible';
      //          frmOutput.ssOleDBGridDefSelRecords.focus();
			
                var dataCollection = frmGridItems.elements;
                if (dataCollection!=null) 
                {
                    for (i=0; i<dataCollection.length; i++)  
                    {
                        if (i==iPollCounter) 
                        {			
                            try 
                            {
                                var frmRefresh = window.parent.parent.parent.window.opener.parent.frames("pollframe").document.forms("frmHit");	
                                var testDataCollection = frmRefresh.elements;
                                iDummy = testDataCollection.txtDummy.value;
                                frmRefresh.submit();
                                iPollCounter = iPollCounter + iPollPeriod;
                            }
                            catch(e) {}
                        }
				
                        sControlName = dataCollection.item(i).name;
                        sControlName = sControlName.substr(0, 12);
                        if (sControlName=="txtGridItem_") 
                        {
                            frmOutput.ssOleDBGridDefSelRecords.additem(dataCollection.item(i).value);
                        }
                    }
                }		
			
                // JPD 19/03/02 Fault 3665
                for (i=0; i<frmOutput.ssOleDBGridDefSelRecords.Columns.Count; i++) 
                {
                    if (frmOutput.ssOleDBGridDefSelRecords.Columns(i).Width > 32000) 
                    {
                        frmOutput.ssOleDBGridDefSelRecords.Columns(i).Width = 32000;
                    }
                }
			
                if (frmExportData.txtPreview.value == 'False')
                {
                    frmOutput.ssOleDBGridDefSelRecords.style.visibility = 'hidden';
                    frmOutput.ssOleDBGridDefSelRecords.Redraw = true;
                    ExportData("OUTPUTRUN");
                    document.getElementById('output').style.visibility = 'hidden';
                    document.getElementById('close').value = 'OK';
                    document.getElementById('tdOutputMSG').innerText = "Custom Report : '"+"' Completed Successfully.";
                    return;
                }
                else
                {
                    frmOutput.ssOleDBGridDefSelRecords.Redraw = true;
                    frmOutput.ssOleDBGridDefSelRecords.style.visibility = 'visible';
                    try
                    {
                        var lngPreviewWidth = new Number(850);
                        var lngPreviewHeight = new Number(550);
                        window.parent.parent.parent.moveTo((screen.width - lngPreviewWidth) / 2, (screen.height - lngPreviewHeight) / 2);
                        window.parent.parent.parent.resizeTo(lngPreviewWidth, lngPreviewHeight);
                    }
                    catch(e) {}
                }
            }
        }
			
        //window.parent.parent.document.all.item("myframeset").rows = "0, *, 0";
        //window.parent.parent.document.all.item("myframeset").rows = "0, *, 0";
        //$("#workframe").attr("myframeset", "REPORT");
        $("workframe").attr("data-framesource", "UTIL_RUN_CUSTOMREPORTSMAIN");
    }		

    function ExportDataPrompt() 
    {
        sURL = "util_run_outputoptions.asp" +
            "?txtUtilType=" + escape(frmExportData.txtUtilType.value) +
            "&txtPreview=" + escape(frmExportData.txtPreview.value) +
            "&txtFormat=" + escape(frmExportData.txtFormat.value) +
            "&txtScreen=" + escape(frmExportData.txtScreen.value) +
            "&txtPrinter=" + escape(frmExportData.txtPrinter.value) +
            "&txtPrinterName=" + escape(frmExportData.txtPrinterName.value) +
            "&txtSave=" + escape(frmExportData.txtSave.value) +
            "&txtSaveExisting=" + escape(frmExportData.txtSaveExisting.value) +
            "&txtEmail=" + escape(frmExportData.txtEmail.value) +
            "&txtEmailAddr=" + escape(frmExportData.txtEmailAddr.value) +
            "&txtEmailAddrName=" + escape(frmExportData.txtEmailAddrName.value) +
            "&txtEmailSubject=" + escape(frmExportData.txtEmailSubject.value) +
            "&txtEmailAttachAs=" + escape(frmExportData.txtEmailAttachAs.value) +
            "&txtFileName=" + escape(frmExportData.txtFileName.value);

        window.parent.parent.ShowOutputOptionsFrame(sURL);
    }

    function openDialog(pDestination, pWidth, pHeight, psResizable, psScroll)
    {
        dlgwinprops = "center:yes;" +
            "dialogHeight:" + pHeight + "px;" +
            "dialogWidth:" + pWidth + "px;" +
            "help:no;" +
            "resizable:" + psResizable + ";" +
            "scroll:" + psScroll + ";" +
            "status:no;";
        window.showModalDialog(pDestination, self, dlgwinprops);
    }

    function replace(sExpression, sFind, sReplace)
    {
        //gi (global search, ignore case)
        var re = new RegExp(sFind,"gi");
        sExpression = sExpression.replace(re, sReplace);
        return(sExpression);
    }
    
    function getData()
    {
        window.parent.parent.loadAddRecords();
    }
    
    function dataOnlyPrint()
    {
        // PageHeaderFont and PageFooterFont don't function in ASPs.
        //  frmOutput.ssOleDBGridDefSelRecords.PageHeaderFont.Name = "Verdana";
        //  frmOutput.ssOleDBGridDefSelRecords.PageHeaderFont.Size = 12;
        //  frmOutput.ssOleDBGridDefSelRecords.PageHeaderFont.Bold = true;
        //  frmOutput.ssOleDBGridDefSelRecords.PageHeaderFont.Underline = true;

        //  frmOutput.ssOleDBGridDefSelRecords.PageFooterFont.Name = "Verdana";
        //  frmOutput.ssOleDBGridDefSelRecords.PageFooterFont.Size = 8;
        //  frmOutput.ssOleDBGridDefSelRecords.PageFooterFont.Bold = false;
        //  frmOutput.ssOleDBGridDefSelRecords.PageFooterFont.Underline = false;

        frmOriginalDefinition.txtOptionsDone.value = 0;

        if (frmExportData.pagebreak.value == "True") 
        {
            // Need to loop through the grid, selecting rows until we find a '*' in
            // the first column ('PageBreak').  
            frmOriginalDefinition.txtCancelPrint.value = 0;
		
            frmOutput.ssHiddenGrid.Caption = frmOutput.ssOleDBGridDefSelRecords.caption;
            frmOutput.ssHiddenGrid.RemoveAll();
            frmOutput.ssHiddenGrid.Columns.RemoveAll();
		
            for (iColIndex = 0; iColIndex < frmOutput.ssOleDBGridDefSelRecords.Cols; iColIndex++) 
            {
                frmOutput.ssHiddenGrid.Columns.Add(iColIndex);
                frmOutput.ssHiddenGrid.Columns(iColIndex).Width = frmOutput.ssOleDBGridDefSelRecords.Columns(iColIndex).Width;
                frmOutput.ssHiddenGrid.Columns(iColIndex).Visible = frmOutput.ssOleDBGridDefSelRecords.Columns(iColIndex).Visible;
                frmOutput.ssHiddenGrid.Columns(iColIndex).Caption = frmOutput.ssOleDBGridDefSelRecords.Columns(iColIndex).Caption;
                frmOutput.ssHiddenGrid.Columns(iColIndex).Name = frmOutput.ssOleDBGridDefSelRecords.Columns(iColIndex).Name;
                frmOutput.ssHiddenGrid.Columns(iColIndex).Alignment = frmOutput.ssOleDBGridDefSelRecords.Columns(iColIndex).Alignment;
                frmOutput.ssHiddenGrid.Columns(iColIndex).CaptionAlignment = frmOutput.ssOleDBGridDefSelRecords.Columns(iColIndex).CaptionAlignment;
            }
		
            frmOutput.ssOleDBGridDefSelRecords.redraw = false;
            frmOutput.ssOleDBGridDefSelRecords.moveFirst();
		
            for (iIndex = 1; iIndex <= frmOutput.ssOleDBGridDefSelRecords.rows; iIndex++)	
            {		
                if (frmOutput.ssOleDBGridDefSelRecords.Columns(0).value == "*") 
                {
                    // NB. In DatMgr we just printSelectedRows. This doesn't work in an ASP
                    // so I copy the required rows to a hidden grid and do printAll on that.
                    if (frmOriginalDefinition.txtOptionsDone.value == 0) 
                    {
                        button_disable(window.parent.parent.parent.frames("top").frmPopup.Cancel, true);
                        frmOutput.ssHiddenGrid.PrintData(23,false,true);	
                        try 
                        {
                            button_disable(window.parent.parent.frames("top").frmPopup.Cancel, false);
                        }
                        catch(e) {}

                        frmOriginalDefinition.txtOptionsDone.value  = 1;
                        if (frmOriginalDefinition.txtCancelPrint.value == 1) 
                        {
                            frmOutput.ssOleDBGridDefSelRecords.redraw = true;
                            return;
                        }
                    }
                    else 
                    {
                        frmOutput.ssHiddenGrid.PrintData(23,false,false);	
                    }
                    frmOutput.ssHiddenGrid.RemoveAll();
                }
                else 
                {
                    sAddItem = new String("");
                    for (iColIndex = 0; iColIndex < frmOutput.ssOleDBGridDefSelRecords.Cols; iColIndex++) 
                    {
                        if(iColIndex > 0) 
                        {
                            sAddItem = sAddItem + "	";
                        }
                        sAddItem = sAddItem + frmOutput.ssOleDBGridDefSelRecords.Columns(iColIndex).value;
                    }
                    frmOutput.ssHiddenGrid.AddItem(sAddItem);
                }

                if (iIndex < frmOutput.ssOleDBGridDefSelRecords.rows) 
                {
                    frmOutput.ssOleDBGridDefSelRecords.MoveNext();
                }
                else 
                {
                    if (frmOriginalDefinition.txtOptionsDone.value == 0) 
                    {
                        button_disable(window.parent.parent.parent.frames("top").frmPopup.Cancel, true);
                        frmOutput.ssHiddenGrid.PrintData(23,false,true);	
                        try 
                        {
                            button_disable(window.parent.parent.frames("top").frmPopup.Cancel, false);
                        }
                        catch(e) {}
					
                        if(frmOriginalDefinition.txtCancelPrint.value == 1) 
                        {
                            frmOutput.ssOleDBGridDefSelRecords.redraw = true;
                            return;
                        }
                    }
                    else 
                    {
                        frmOutput.ssHiddenGrid.PrintData(23,false,false);	
                    }
                    break;
                }
            }
            frmOutput.ssOleDBGridDefSelRecords.redraw = true;
        }
        else 
        {
            button_disable(window.parent.parent.parent.frames("top").frmPopup.Cancel, true);
            frmOutput.ssOleDBGridDefSelRecords.PrintData(23,false,true);	
            try 
            {
                button_disable(window.parent.parent.frames("top").frmPopup.Cancel, false);
            }
            catch(e) {}
        }
    }

    function closeclick()
    {
        try
        {
            window.parent.parent.close();
        }
        catch (e)	{}
    }

-->
</script>


<script type="text/javascript">
    //addActiveXHandlers();
    customreports_window_onload();
</script>
