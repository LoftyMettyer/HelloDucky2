<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="HR.Intranet.Server.Structures" %>
<%@ Import Namespace="System.Data" %>

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
		pblnOutputScreen = Session("stdReport_OutputScreen")
		pblnOutputPrinter = Session("stdReport_OutputPrinter")
		pstrOutputPrinterName = Session("stdReport_OutputPrinterName")
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
		
	
	' this only needed for report output now.
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
	
	Dim fNoRecords As Boolean
	Dim sGroupingParams As String = ""
	Dim jsFooterFunction As String = ""
	Dim jsSrvFunction As New StringBuilder

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
		
		Dim colNamesArray As New StringBuilder
		Dim colModelArray As New StringBuilder
		Dim bGroupWithNext As Boolean
		Dim bGrouping As Boolean = False
		Dim sGroupFieldList As String = ""
		Dim sGroupOrder As String = ""
		Dim sGroupColumnShowList As String = ""
		Dim sGroupTextList As String = ""
		Dim bFooter As Boolean = False
		Dim sSortString As String = ""
		Dim iColIndex As Integer = 0
		
		' Configure COLUMNS Model for jqGrid
		For Each objRow As DataRow In objReport.mrstCustomReportsDetails.Rows
			
			Dim sColumnHeading As String = objReport.mrstCustomReportsOutput.Columns(iColIndex).ColumnName
			Dim sFooterText As String = ""
			Dim isVisibleString As String = ""
			Dim cellAttributes As String = ""
			Dim summaryType As String = ""
			Dim alignment As String = ""
			
			colNamesArray.Append(String.Format("{0}'{1}'", IIf(colNamesArray.Length > 0, ",", ""), sColumnHeading))
		
			' Report Configuration Options	
			' keep here; we use the groupwithnext flag.
			If CBool(objRow.Item("Boc")) Or CBool(objRow.Item("Poc")) Or CBool(objRow.Item("Voc")) Then
				bGrouping = True
				sGroupFieldList &= String.Format("{0}'{1}'", IIf(sGroupFieldList.Length > 0, ", ", ""), sColumnHeading)
				sGroupColumnShowList &= String.Format("{0}{1}", IIf(sGroupColumnShowList.Length > 0, ", ", ""), (bGroupWithNext = False).ToString().ToLower())
				sGroupTextList &= String.Format("{0}'{{0}}'", IIf(sGroupTextList.Length > 0, ", ", ""))
				sGroupOrder &= String.Format("{0}'{1}'", IIf(sGroupOrder.Length > 0, ", ", ""), objRow.Item("SortOrder").ToString().Trim().ToLower())
			End If
			
			' Suppress repeated values
			If CBool(objRow.Item("Srv")) Then
				jsSrvFunction.Append("var lastValue = $('#grdReport tr:first td:nth-child(" & iColIndex + 1.ToString() & ")').text();")
				jsSrvFunction.Append("$('#grdReport tr td:nth-child(" & iColIndex + 1.ToString() & ")').each(function () {")
				jsSrvFunction.Append("if($(this).text() == lastValue) {")
				jsSrvFunction.Append("	$(this).text('');")
				jsSrvFunction.Append("}")
				jsSrvFunction.Append("else {lastValue = $(this).text();}")
				jsSrvFunction.Append("});")
			End If
		
			If objRow.Item("Hidden") Or bGroupWithNext Then isVisibleString = ", hidden: true"
			If CBool(objRow.Item("GroupWithNextColumn")) Then cellAttributes = ", cellattr: function (rowId, tv, rawObject, cm, rdata) {return 'style=""white-space: normal;""';}"
		
			' Count, Sum, Average functions.
			Dim decimalPlaces As Integer = NullSafeInteger(objRow.Item("dp"))

			If CBool(objRow.Item("Avge")) Then
				summaryType = ", summaryType: ""avg"", summaryTpl: ""Sub Average: {0}"""
				jsFooterFunction &= String.Format("var avge_{0} = Number({1}).toFixed({2});", sColumnHeading.Replace(" ", "_"), objReport.mrstCustomReportsOutput.Compute("Avg(" & sColumnHeading & ")", ""), decimalPlaces)
				sFooterText = String.Format("'Average: ' + avge_{0}", sColumnHeading.Replace(" ", "_"))
			End If
			If CBool(objRow.Item("Cnt")) Then
				summaryType = ", summaryType: ""count"", summaryTpl: ""Sub Count: {0}"""
				jsFooterFunction &= String.Format("var cnt_{1} = {0};", objReport.mrstCustomReportsOutput.Compute("Count(" & sColumnHeading & ")", ""), sColumnHeading.Replace(" ", "_"))
				sFooterText &= String.Format("{0}'Count: ' + cnt_{1}", IIf(sFooterText.Length > 0, "+ '<br/>' + ", ""), sColumnHeading.Replace(" ", "_"))
			End If
			If CBool(objRow.Item("Tot")) Then
				summaryType = ", summaryType: ""sum"", summaryTpl: ""Sub Total: {0}"""
				jsFooterFunction &= String.Format("var sum_{0} = Number({1}).toFixed({2});", sColumnHeading.Replace(" ", "_"), objReport.mrstCustomReportsOutput.Compute("Sum(" & sColumnHeading & ")", ""), decimalPlaces)
				sFooterText &= String.Format("{0}'Total: ' + sum_{1}", IIf(sFooterText.Length > 0, "+ '<br/>' + ", ""), sColumnHeading.Replace(" ", "_"))
			End If
		
			If CBool(objRow.Item("IsNumeric")) And Not CBool(objRow.Item("GroupWithNextColumn")) Then alignment = ", align: ""right"", sorttype: ""integer"", formatter: ""number"", formatoptions:{ thousandsSeparator: """", defaultValue: """", decimalPlaces: " & decimalPlaces & "}"
					
			' This is a Date Column - format as such
			If objReport.mrstCustomReportsOutput.Columns(iColIndex).DataType = System.Type.GetType("System.DateTime") Then
				alignment = ", align: ""left"", sorttype: ""date"", formatter: ""date"", formatoptions: {srcformat: ""d/m/Y"", newformat: ""d/m/Y""}"
			End If
		
			' This is a Logic Column - format as such
			If objReport.mrstCustomReportsOutput.Columns(iColIndex).DataType = System.Type.GetType("System.Boolean") Then
				alignment = ", align: ""center"", formatter: ""checkbox"""
			End If

			' Footer row required?
			If bFooter = False Then bFooter = CBool(objRow.Item("Avge")) Or CBool(objRow.Item("Cnt")) Or CBool(objRow.Item("Tot"))
			If sFooterText.Length > 0 Then jsFooterFunction &= String.Format("jQuery('#grdReport').jqGrid('footerData', 'set', {{ '{0}': {1} }}, false);", sColumnHeading, sFooterText)
			
			' add column info to the colModel array.
			colModelArray.Append(String.Format("{{name:'{0}',index:'{0}', width: 100{1}{2}{3}{4}}},", sColumnHeading, isVisibleString, cellAttributes, summaryType, alignment))
		
			' set Group With Next flag.
			bGroupWithNext = CBool(objRow.Item("GroupWithNextColumn"))
		
			' Build sortOrder string for dataview when binding the DATA below..
			If objRow.Item("SortOrder").ToString().ToUpper() = "ASC" Or objRow.Item("SortOrder").ToString().ToUpper() = "DESC" Then
				sSortString &= String.Format("{0}{1} {2}", IIf(sSortString.Length > 0, ", ", ""), sColumnHeading, objRow.Item("SortOrder"))
			End If
			
			iColIndex += 1
		
		Next
	
		If bGrouping Then
			sGroupingParams = ",grouping: true," & _
		"groupingView : {groupField : [" & sGroupFieldList & "]," & _
		"groupColumnShow : [" & sGroupColumnShowList & "]," & _
		"groupText : [" & sGroupTextList & "]," & _
		"groupOrder: [" & sGroupOrder & "]," & _
		"groupCollapse : false," & _
		"groupSummary : [true]," & _
		"showSummaryOnHide: true," & _
		"groupDataSorted: false}"
		End If
		
		If bFooter = True Then
			sGroupingParams = ",footerrow: true, userDataOnFooter: true" & sGroupingParams
		End If
	
		' trailing comma removal
		colModelArray.Remove(colModelArray.Length - 1, 1)
	
		' Now for the DATA for jqGrid...
		Dim colData As New StringBuilder
	
		Dim dv As DataView = objReport.mrstCustomReportsOutput.DefaultView		
		
		' No multi-column grouping in current version of jqGrid, so sort on dataview
		If sSortString.Length > 0 Then dv.Sort = sSortString
		
		For Each objRow As DataRow In dv.ToTable().Rows		' objReport.mrstCustomReportsOutput.Rows
			colData.Append("{")
			
			bGroupWithNext = False
			Dim sColumnValue As String = ""
			Dim sColumnName As String = ""
			
			For iColIndex = 0 To objRow.ItemArray.Count() - 1
				Dim objThisColumn As ReportDetailItem = objReport.DisplayColumns(iColIndex)
				
				If Not bGroupWithNext Then
					sColumnValue = objRow.Item(iColIndex).ToString()
					sColumnName = objReport.mrstCustomReportsOutput.Columns(iColIndex).ColumnName
				Else
					' add next col too.
					If objRow.Item(iColIndex).ToString().Length > 0 Then
						sColumnValue &= vbNewLine & objRow.Item(iColIndex).ToString()
					End If
				End If

				bGroupWithNext = objThisColumn.GroupWithNextColumn
				
				' Bug JIRA#3767 - convert to proper case if grouping; jQGrid grouping is case sensitive ('fixed' in v4.6 I believe, but too much for v8.0?)
				If bGrouping Then sColumnValue = StrConv(sColumnValue, VbStrConv.ProperCase)
				
				If Not bGroupWithNext Then
					sColumnValue = Html.Encode(sColumnValue)
					sColumnValue = sColumnValue.Replace(vbNewLine, "<br/>")

					colData.Append(String.Format("'{0}':'{1}',", sColumnName, sColumnValue))
				End If
			Next

			' trailing comma removal
			colData.Remove(colData.Length - 1, 1)
			colData.Append("},")
		Next
		' trailing comma removal
		colData.Remove(colData.Length - 1, 1)

	
		' output to javascript variables.
		Dim cs As ClientScriptManager = Page.ClientScript
		cs.RegisterArrayDeclaration("g_colNamesArray", colNamesArray.ToString())
		cs.RegisterArrayDeclaration("g_colModelArray", colModelArray.ToString())
		cs.RegisterArrayDeclaration("g_colData", colData.ToString())
	End If
	
	If fok Then
		objReport.ClearUp()
	End If

	If fok Then
		Response.Write("<form name=frmOutput id=frmOutput method=post>" & vbCrLf)
		Response.Write("<div>")
		Response.Write("<table id='grdReport'></table>" & vbCrLf)
	
			
%>
<form id="formReportData" runat="server">
</form>
<%		
	Response.Write("      </div>")
	Response.Write("</form>" & vbCrLf)
		
	Response.Write("<input type='hidden' id=txtNoRecs name=txtNoRecs value=0>" & vbCrLf)
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
					<tr>
						<td colspan="3" height="10" align="center">
							<input type="button" id="cmdClose" name="cmdClose" value="Close" style="WIDTH: 80px" width="80" class="btn"
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

<input type='hidden' id="txtNoRecs" name="txtNoRecs" value="1">
<input type="hidden" id="txtSuccessFlag" name="txtSuccessFlag" value="3">
<%
End If
%>


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



<script type="text/javascript">

	<%If fok = True Then%>
	//Shrink to fit, or set to 100px per column?
	var ShrinkToFit = false;
	var gridWidth;
	var gridHeight;
	
	//Get count of visible columns
	var iVisibleCount = 0;
	for (var iArrayPos = 0; iArrayPos < g_colModelArray.length; iArrayPos++) {
		if (g_colModelArray[iArrayPos].hidden !== true) iVisibleCount++;
	}


	if (menu_isSSIMode()) {
		try {
			gridWidth = $('#reportworkframe').width();
			gridHeight = $('#reportworkframe').height() - 100;
		} catch(e) {
			gridWidth = 'auto';
			gridHeight = 'auto';
		}
		ShrinkToFit = true;
	} else {
		//DMI options.
		if (iVisibleCount < 8) ShrinkToFit = true;
		gridWidth = 770;
		gridHeight = 390;
	}
	


	jQuery("#grdReport").jqGrid({
		datatype: "local",
		shrinkToFit: ShrinkToFit,
		width: gridWidth,
		height: gridHeight,
		colNames: g_colNamesArray,
		colModel: g_colModelArray,
		data: g_colData,
		ignoreCase: true,
		rowNum: 200000		
			<%=sGroupingParams%>,
		loadComplete: function () {
			<%=jsFooterFunction%>;
			<%=jsSrvFunction.ToString()%>;
			stylejqGrid();
		}
	});

	function stylejqGrid() {
		//jqGrid style overrides
		//hide caption	
		$("#gview_grdReport > .ui-jqgrid-titlebar").hide(); //no title bar; this is in the dialog title
		$('#gview_grdReport tr.jqgrow td').css('vertical-align', 'top'); //float text to top, in case of multi-line cells
		$('#gview_grdReport .s-ico span').css('display', 'none'); //hide the sort order icons - they don't tie in to the dataview model.
		$('#gview_grdReport tr.footrow td').css('vertical-align', 'top'); //float text to top, in case of multi-line footers
	}

	if (menu_isSSIMode()) $('#gbox_grdReport').css('margin', '0 auto'); //center the report in self-service screen.	
	

	<%end if%>
	
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
	<input type="hidden" id="txtEmailGroupID" name="txtEmailGroupID" value="">
	<input type="hidden" id="txtUtilType" name="txtUtilType" value="<%=session("utilType")%>">

	<iframe name="submit-iframe" style="display: none;"></iframe>

</form>
