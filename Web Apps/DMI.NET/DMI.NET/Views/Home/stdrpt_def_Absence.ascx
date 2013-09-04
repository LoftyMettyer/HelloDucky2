﻿<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<%="" %>

<script src="<%: Url.Content("~/bundles/utilities_standardreports")%>" type="text/javascript"></script>

<%
	Dim sKey As String
		
	' Clear the session action which is used to botch the prompted values screen in
	Session("action") = ""
	Session("optionaction") = ""

	' Read the prompted start/end dates if there were any
	Dim aPrompts(1)

	For i = 0 To (Request.Form.Count) - 1
		sKey = Request.Form.Keys(i)
		If ((UCase(Left(sKey, 7)) = "PROMPT_") And (Mid(sKey, 8, 1) <> "3")) Or _
				(UCase(Left(sKey, 10)) = "PROMPTCHK_") Then
			
			If Mid(sKey, 8, 5) = "start" Then
				aPrompts(0) = Request.Form.Item(i)
			Else
				aPrompts(1) = Request.Form.Item(i)
			End If

		End If
	Next

	' Settings objects
	Dim objSettings As New HR.Intranet.Server.clsSettings
	objSettings.Connection = Session("databaseConnection")

	Dim aColumnNames
	Dim aAbsenceTypes
	Dim cmdReportsCols
	Dim prmBaseTableID
	Dim rstReportColumns
	Dim sErrorDescription As String
	Dim iCount As Integer
		
	' Load Absence Types and Personnel columns into array
	ReDim aColumnNames(1, 0)
	ReDim aAbsenceTypes(0)

	' Get the table records.
	cmdReportsCols = Server.CreateObject("ADODB.Command")
	cmdReportsCols.CommandText = "sp_ASRIntGetColumns"
	cmdReportsCols.CommandType = 4 ' Stored procedure
	cmdReportsCols.ActiveConnection = Session("databaseConnection")
																				
	prmBaseTableID = cmdReportsCols.CreateParameter("piTableID", 3, 1) ' 3=integer, 1=input
	cmdReportsCols.Parameters.Append(prmBaseTableID)
	prmBaseTableID.value = CleanNumeric(Session("Personnel_EmpTableID"))

	Err.Clear()
	rstReportColumns = cmdReportsCols.Execute
																																			
	If (Err.Number <> 0) Then
		sErrorDescription = "The personnel column information could not be retrieved." & vbCrLf & FormatError(Err.Description)
	End If

	If Len(sErrorDescription) = 0 Then
		iCount = 0
		Do While Not rstReportColumns.EOF
		
			If rstReportColumns.fields("OLEType").value <> 2 Then
		
				aColumnNames(0, iCount) = rstReportColumns.fields("ColumnID").value
				aColumnNames(1, iCount) = rstReportColumns.fields("ColumnName").value
			
				ReDim Preserve aColumnNames(1, UBound(aColumnNames, 2) + 1)
				iCount = iCount + 1
			End If
			
			rstReportColumns.MoveNext()
			
		Loop

		rstReportColumns.close()

	End If

	' Load absence types
	Dim cmdTables
	Dim rstTablesInfo
		
	cmdTables = Server.CreateObject("ADODB.Command")
	cmdTables.CommandText = "sp_ASRIntGetAbsenceTypes"
	cmdTables.CommandType = 4	' Stored Procedure
	cmdTables.ActiveConnection = Session("databaseConnection")

	rstTablesInfo = cmdTables.Execute
																			
	If (Err.Number <> 0) Then
		sErrorDescription = "The absence type information could not be retrieved." & vbCrLf & FormatError(Err.Description)
	End If

	If Len(sErrorDescription) = 0 Then
		iCount = 0
		Do While Not rstTablesInfo.EOF
			aAbsenceTypes(iCount) = rstTablesInfo.fields("Type").value
			ReDim Preserve aAbsenceTypes(UBound(aAbsenceTypes) + 1)
			iCount = iCount + 1
			rstTablesInfo.MoveNext()
		Loop

		rstTablesInfo.close()

	End If
												
	' Release the ADO objects.
	cmdTables = Nothing
	cmdReportsCols = Nothing


	' Set the default settings
	Dim strReportType As String = "AbsenceBreakdown"
	Dim strDate
	Dim strType
	Dim lngDefaultColumnID As Long
	Dim lngConfigColumnID As Long
	Dim strSaveExisting As String
	
	Response.Write("<script type=""text/javascript"">" & vbCrLf)

	Response.Write("function SetReportDefaults(){" & vbCrLf)
	Response.Write("   var frmAbsenceDefinition = OpenHR.getForm(""workframe"",""frmAbsenceDefinition"");" & vbCrLf)
		
	' Type of standard report being run
	If Session("StandardReport_Type") = 16 Then
		strReportType = "BradfordFactor"
	End If

	If Session("StandardReport_Type") = 15 Then
		strReportType = "AbsenceBreakdown"
		Response.Write("frmAbsenceDefinition.btnTab2.style.visibility = ""hidden"";" & vbCrLf)
	End If

	' Absence types
	For iCount = 0 To UBound(aAbsenceTypes) - 1
		If objSettings.GetSystemSetting(strReportType, "Absence Type " & aAbsenceTypes(iCount), "0") = "1" Then
			Response.Write("frmAbsenceDefinition.chkAbsenceType_" & iCount & ".checked = 1;" & vbCrLf)
		End If
	Next

	' Date range
	If Len(aPrompts(0)) = 0 Then
		strDate = ConvertSQLDateToLocale(objSettings.GetStandardReportDate(strReportType, "Start Date"))
	Else
		strDate = aPrompts(0)
	End If
	Response.Write("frmAbsenceDefinition.txtDateFrom.value = " & """" & CleanStringForJavaScript(strDate) & """" & ";" & vbCrLf)

	' Date range
	If Len(aPrompts(1)) = 0 Then
		strDate = ConvertSQLDateToLocale(objSettings.GetStandardReportDate(strReportType, "End Date"))
	Else
		strDate = aPrompts(1)
	End If
	Response.Write("frmAbsenceDefinition.txtDateTo.value = " & """" & CleanStringForJavaScript(strDate) & """" & ";" & vbCrLf)

	' Record Selection
	If Session("optionRecordID") = "0" Then

		strType = objSettings.GetSystemSetting(strReportType, "Type", "A")
		
		Select Case strType
			Case "A"
				Response.Write("frmAbsenceDefinition.optAllRecords.checked = 1;" & vbCrLf)
				Response.Write("frmAbsenceDefinition.optPickList.checked = 0;" & vbCrLf)
				Response.Write("frmAbsenceDefinition.optFilter.checked = 0;" & vbCrLf)
			Case "P"
				Response.Write("frmAbsenceDefinition.optAllRecords.checked = 0;" & vbCrLf)
				Response.Write("frmAbsenceDefinition.optPickList.checked = 1;" & vbCrLf)
				Response.Write("frmAbsenceDefinition.optFilter.checked = 0;" & vbCrLf)
				Response.Write("frmAbsenceDefinition.txtBasePicklist.value = " & """" & CleanStringForJavaScript(objSettings.GetPicklistFilterName(strReportType, strType)) & """" & ";" & vbCrLf)
				Response.Write("button_disable(frmAbsenceDefinition.cmdBasePicklist, false);" & vbCrLf)
				Response.Write("frmPostDefinition.txtBasePicklistID.value = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "ID", "0")) & ";" & vbCrLf)
			Case "F"
				Response.Write("frmAbsenceDefinition.optAllRecords.checked = 0;" & vbCrLf)
				Response.Write("frmAbsenceDefinition.optPickList.checked = 0;" & vbCrLf)
				Response.Write("frmAbsenceDefinition.optFilter.checked = 1;" & vbCrLf)
				Response.Write("frmAbsenceDefinition.txtBaseFilter.value = " & """" & CleanStringForJavaScript(objSettings.GetPicklistFilterName(strReportType, strType)) & """" & ";" & vbCrLf)
				Response.Write("button_disable(frmAbsenceDefinition.cmdBaseFilter, false);" & vbCrLf)
				Response.Write("frmPostDefinition.txtBaseFilterID.value = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "ID", "0")) & ";" & vbCrLf)
		End Select
	Else
		Response.Write("RecordSelection.style.visibility = ""hidden"";" & vbCrLf)
	End If

	' Display picklist in header
	Response.Write("frmAbsenceDefinition.chkPrintInReportHeader.checked =  " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "PrintFilterHeader", "0")) & vbCrLf)

	' Bradford Factor specific stuff
	If Session("StandardReport_Type") = 16 Then
		Response.Write("frmAbsenceDefinition.chkSRV.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "SRV", "0")) & ";" & vbCrLf)
		Response.Write("frmAbsenceDefinition.chkShowDurations.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Show Totals", "1")) & ";" & vbCrLf)
		Response.Write("frmAbsenceDefinition.chkShowInstances.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Show Count", "0")) & ";" & vbCrLf)
		Response.Write("frmAbsenceDefinition.chkShowFormula.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Show Workings", "0")) & ";" & vbCrLf)
		Response.Write("frmAbsenceDefinition.chkShowAbsenceDetails.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Display Absence Details", "1")) & ";" & vbCrLf)
		Response.Write("frmAbsenceDefinition.chkOmitBeforeStart.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Omit Before", "0")) & ";" & vbCrLf)
		Response.Write("frmAbsenceDefinition.chkOmitAfterEnd.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Omit After", "0")) & ";" & vbCrLf)
		Response.Write("frmAbsenceDefinition.chkMinimumBradfordFactor.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Minimum Bradford Factor", "0")) & ";" & vbCrLf)
		Response.Write("frmAbsenceDefinition.txtMinimumBradfordFactor.value = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Minimum Bradford Factor Amount", "0")) & ";" & vbCrLf)
		Response.Write("frmAbsenceDefinition.chkOrderBy1Asc.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Order By Asc", "1")) & ";" & vbCrLf)
		Response.Write("frmAbsenceDefinition.chkOrderBy2Asc.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Group By Asc", "1")) & ";" & vbCrLf)
			 
		lngDefaultColumnID = objSettings.GetModuleParameter("MODULE_PERSONNEL", "Param_FieldsSurname")
		lngConfigColumnID = objSettings.GetSystemSetting(strReportType, "Order By", lngDefaultColumnID)
		'Response.Write "frmAbsenceDefinition.cboOrderBy1.value = " & """" & sFieldName & """" & ";" & vbcrlf
		Response.Write("for (var i=0; i<frmAbsenceDefinition.cboOrderBy1.options.length; i++)" & vbCrLf)
		Response.Write("	{" & vbCrLf)
		Response.Write("	if (frmAbsenceDefinition.cboOrderBy1.options[i].value == " & lngConfigColumnID & ")" & vbCrLf)
		Response.Write("		{" & vbCrLf)
		Response.Write("		frmAbsenceDefinition.cboOrderBy1.selectedIndex = i; " & vbCrLf)
		Response.Write("		}" & vbCrLf)
		Response.Write("	}" & vbCrLf)
		
		lngDefaultColumnID = objSettings.GetModuleParameter("MODULE_PERSONNEL", "Param_FieldsForename")
		lngConfigColumnID = objSettings.GetSystemSetting(strReportType, "Group By", lngDefaultColumnID)
		'Response.Write "frmAbsenceDefinition.cboOrderBy2.value = " & """" & sFieldName & """" & ";" & vbcrlf
		Response.Write("for (var i=0; i<frmAbsenceDefinition.cboOrderBy2.options.length; i++)" & vbCrLf)
		Response.Write("	{" & vbCrLf)
		Response.Write("	if (frmAbsenceDefinition.cboOrderBy2.options[i].value == " & lngConfigColumnID & ")" & vbCrLf)
		Response.Write("		{" & vbCrLf)
		Response.Write("		frmAbsenceDefinition.cboOrderBy2.selectedIndex = i; " & vbCrLf)
		Response.Write("		}" & vbCrLf)
		Response.Write("	}" & vbCrLf)
	End If

	' Output Options
	Select Case objSettings.GetSystemSetting(strReportType, "Format", 0)
		Case "0"
			Response.Write("frmAbsenceDefinition.optOutputFormat0.checked = 1;" & vbCrLf)
		Case "1"
			Response.Write("frmAbsenceDefinition.optOutputFormat1.checked = 1;" & vbCrLf)
		Case "2"
			Response.Write("frmAbsenceDefinition.optOutputFormat2.checked = 1;" & vbCrLf)
		Case "3"
			Response.Write("frmAbsenceDefinition.optOutputFormat3.checked = 1;" & vbCrLf)
		Case "4"
			Response.Write("frmAbsenceDefinition.optOutputFormat4.checked = 1;" & vbCrLf)
		Case "5"
			Response.Write("frmAbsenceDefinition.optOutputFormat5.checked = 1;" & vbCrLf)
		Case "6"
			'MH20031211 Fault 7787
			'If Bradford then disallow Pivot (make it worksheet instead)
			If Session("StandardReport_Type") = 16 Then
				Response.Write("frmAbsenceDefinition.optOutputFormat4.checked = 1;" & vbCrLf)
			Else
				Response.Write("frmAbsenceDefinition.optOutputFormat6.checked = 1;" & vbCrLf)
			End If
		Case Else
			' Charts and pivot not in Intranet yet
			Response.Write("frmAbsenceDefinition.optOutputFormat0.checked = 1" & vbCrLf)
	End Select
	
	Response.Write("frmAbsenceDefinition.chkPreview.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Preview", 0)) & ";" & vbCrLf)
	Response.Write("frmAbsenceDefinition.chkDestination0.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Screen", 1)) & ";" & vbCrLf)

	Response.Write("frmAbsenceDefinition.chkDestination1.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Printer", 0)) & ";" & vbCrLf)

	Response.Write("strPrinterName = '" & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "PrinterName", "")) & "';" & vbCrLf)

	'Set the printer as defined in Report Configuration in DAT.
	'Response.Write "frmAbsenceDefinition.cboPrinterName.value = " & """" & objSettings.GetSystemSetting(strReportType, "PrinterName", "") & """" & ";" & vbcrlf
	Response.Write("for (var i=0; i<frmAbsenceDefinition.cboPrinterName.options.length; i++)" & vbCrLf)
	Response.Write("	{" & vbCrLf)
	Response.Write("	if (frmAbsenceDefinition.cboPrinterName.options[i].innerText.toLowerCase() == strPrinterName.toLowerCase())" & vbCrLf)
	Response.Write("		{" & vbCrLf)
	Response.Write("		frmAbsenceDefinition.cboPrinterName.selectedIndex = i; " & vbCrLf)
	Response.Write("		}" & vbCrLf)
	Response.Write("	}" & vbCrLf)

	'MH20040311
	Response.Write("if (frmAbsenceDefinition.chkDestination1.checked == true) " & vbCrLf)
	Response.Write("	{" & vbCrLf)
	Response.Write("	if (strPrinterName != """") " & vbCrLf)
	Response.Write("		{" & vbCrLf) '
	Response.Write("		if (frmAbsenceDefinition.cboPrinterName.options[frmAbsenceDefinition.cboPrinterName.selectedIndex].innerText != strPrinterName) " & vbCrLf)
	Response.Write("			{" & vbCrLf)
	Response.Write("			window.parent.frames(""menuframe"").ASRIntranetFunctions.MessageBox(""This definition is set to output to printer ""+strPrinterName+"" which is not set up on your PC."");" & vbCrLf)
	Response.Write("			var oOption = document.createElement(""OPTION"");" & vbCrLf)
	Response.Write("			frmAbsenceDefinition.cboPrinterName.options.add(oOption);" & vbCrLf)
	Response.Write("			oOption.innerText = strPrinterName;" & vbCrLf)
	Response.Write("			oOption.value = frmAbsenceDefinition.cboPrinterName.options.length-1;" & vbCrLf)
	Response.Write("			frmAbsenceDefinition.cboPrinterName.selectedIndex = oOption.value;" & vbCrLf)
	Response.Write("			}" & vbCrLf)
	Response.Write("		}" & vbCrLf)
	Response.Write("	}" & vbCrLf)
	
	Response.Write("frmAbsenceDefinition.chkDestination2.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Save", 0)) & ";" & vbCrLf)
	Response.Write("frmAbsenceDefinition.txtFilename.value = " & """" & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "FileName", "")) & """" & ";" & vbCrLf)

	Response.Write("populateSaveExisting();" & vbCrLf)
	Select Case objSettings.GetSystemSetting(strReportType, "SaveExisting", 0)
		Case 0
			strSaveExisting = "Overwrite"
			Response.Write("frmAbsenceDefinition.cboSaveExisting.selectedIndex = 0;" & vbCrLf)
		Case 1
			strSaveExisting = "Do not overwrite"
			Response.Write("frmAbsenceDefinition.cboSaveExisting.selectedIndex = 1;" & vbCrLf)
		Case 2
			strSaveExisting = "Add sequential number to name"
			Response.Write("frmAbsenceDefinition.cboSaveExisting.selectedIndex = 2;" & vbCrLf)
		Case 3
			strSaveExisting = "Append to file"
			Response.Write("frmAbsenceDefinition.cboSaveExisting.selectedIndex = 3;" & vbCrLf)
		Case 4
			strSaveExisting = "Create new sheet in workbook"
			Response.Write("frmAbsenceDefinition.cboSaveExisting.selectedIndex = 4;" & vbCrLf)
	End Select

	'Response.Write "frmAbsenceDefinition.cboSaveExisting.value = " & """" & strSaveExisting & """"  & ";" & vbcrlf
	Response.Write("frmAbsenceDefinition.chkDestination3.checked = " & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "Email", 0)) & ";" & vbCrLf)
	Response.Write("frmAbsenceDefinition.txtEmailGroup.value = " & """" & CleanStringForJavaScript(objSettings.GetEmailGroupName(objSettings.GetSystemSetting(strReportType, "EmailAddr", "0"))) & """" & ";" & vbCrLf)
	Response.Write("frmAbsenceDefinition.txtEmailGroupID.value = " & """" & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "EmailAddr", "")) & """" & ";" & vbCrLf)
	Response.Write("frmAbsenceDefinition.txtEmailAttachAs.value = " & """" & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "EmailAttachAs", "")) & """" & ";" & vbCrLf)
	Response.Write("frmAbsenceDefinition.txtEmailSubject.value = " & """" & CleanStringForJavaScript(objSettings.GetSystemSetting(strReportType, "EmailSubject", "")) & """" & ";" & vbCrLf)

	Response.Write(vbCrLf & "}")
	Response.Write("</script>" & vbCrLf)

	objSettings = Nothing

%>

<form id="frmAbsenceDefinition" name="frmAbsenceDefinition">
	<table cellpadding="5" cellspacing="0" width="700" height="60%">
		<tr>
			<td>
				<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
					<tr height="5">
						<td colspan="3"></td>
					</tr>

					<tr height="10">
						<td width="10"></td>
						<td>
							<input type="button" class="btn btndisabled" value="Definition" id="btnTab1" name="btnTab1" disabled="disabled"
								onclick="displayPage(1)"
								onmouseover="try{button_onMouseOver(this);}catch(e){}"
								onmouseout="try{button_onMouseOut(this);}catch(e){}"
								onfocus="try{button_onFocus(this);}catch(e){}"
								onblur="try{button_onBlur(this);}catch(e){}" />
							<%
								If Session("StandardReport_Type") = 16 Then
							%>
							<input type="button" class="btn" value="Options" id="btnTab2" name="btnTab2"
								onclick="displayPage(2)"
								onmouseover="try{button_onMouseOver(this);}catch(e){}"
								onmouseout="try{button_onMouseOut(this);}catch(e){}"
								onfocus="try{button_onFocus(this);}catch(e){}"
								onblur="try{button_onBlur(this);}catch(e){}" />
							<%
							End If
							%>
							<input type="button" class="btn" value="Output" id="btnTab3" name="btnTab3"
								onclick="displayPage(3)"
								onmouseover="try{button_onMouseOver(this);}catch(e){}"
								onmouseout="try{button_onMouseOut(this);}catch(e){}"
								onfocus="try{button_onFocus(this);}catch(e){}"
								onblur="try{button_onBlur(this);}catch(e){}" />
							<%
								' Causes problems if button isn't there
								If Session("StandardReport_Type") <> 16 Then
							%>
							<input type="button" class="btn" value="Options" id="btnTab2" name="btnTab2"
								onclick="displayPage(2)"
								onmouseover="try{button_onMouseOver(this);}catch(e){}"
								onmouseout="try{button_onMouseOut(this);}catch(e){}"
								onfocus="try{button_onFocus(this);}catch(e){}"
								onblur="try{button_onBlur(this);}catch(e){}" />
							<%
							End If
							%>

						</td>
						<td width="10"></td>
					</tr>

					<tr height="10">
						<td colspan="3"></td>
					</tr>

					<tr>
						<td width="10"></td>
						<td>
							<!-- First tab -->
							<div id="div1">
								<table width="100%" height="100%"  cellspacing="0" cellpadding="5">
									<tr valign="top">
										<td valign="top" rowspan="2" width="25%" height="100%">
											<table class="invisible" cellspacing="0" cellpadding="4" width="100%" height="100%">
												<tr style="height:40px">
													<td style="width: 400px; height:40px; text-align: left; vertical-align: top"><strong>Absence Types :</strong>
														<br><br>
														<span id="AbsenceTypes" style="width: 500px; height: 200px; overflow: auto;" >
															<table class="invisible" cellspacing="0" style="width: 200px;padding: 0">
																<tr style="padding-top: 20px">
																	<td>
																		<%
																			For iCount = 0 To UBound(aAbsenceTypes) - 1
																		%>

																		<tr>
																			<td>
																				<input id="chkAbsenceType_<%=iCount%>" name="chkAbsenceType_<%=iCount%>" type="checkbox" tagname="<%=aAbsenceTypes(iCount)%>" tabindex="-1"
																					onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																					onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																				<label
																					for="chkAbsenceType_<%=iCount%>"
																					class="checkbox"
																					tabindex="0"
																					onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																					onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																					onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																					onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																					onblur="try{checkboxLabel_onBlur(this);}catch(e){}">

																					<%=aAbsenceTypes(iCount)%>
																				</label>
																			</td>
																		</tr>
																		<%
																		Next
																		%>
																	</td>
																</tr>
															</table>
														</span>
													</td>

													<td height="10" align="left" valign="top">
														<table cellspacing="1" cellpadding="1" width="100%" >
															<tr>
																<td colspan="2" style="font-weight: bold">Date Range :</td>
															</tr>
															<tr style="height: 10px"> </tr>
															<tr>
																<td>
																	<table border="0">
																		<tr>
																			<td width="100">Start Date :
																			</td>
																			<td>
																				<input id="txtDateFrom" class="text" name="txtDateFrom" onblur="validateDate(this);">
																			</td>
																		</tr>
																		<tr>
																			<td width="100">End Date :
																			</td>
																			<td>
																				<input id="txtDateTo" class="text" name="txtDateTo" onblur="validateDate(this);">
																			</td>
																		</tr>
																	</table>
																</td>
															</tr>
														</table>

														&nbsp;

													<span id="RecordSelection">
														<table cellspacing="1" cellpadding="1" width="300" >
															<tr height="10">
																<td style="height:10px; text-align: left; vertical-align : top; font-weight: bold">Record Selection :</td>

																	<table class="invisible" cellspacing="0" cellpadding="3">
																	<tr style="height: 10px" </tr>
																		<tr>
																			<table width="325" height="80%" border="0" cellspacing="0" cellpadding="5">
																			<td>
																				<table width="360" class="invisible" cellspacing="0" cellpadding="0">
																					<tr>
																						<td style="width:95PX; padding-bottom: 6px" colspan="3">
																							<input checked id="optAllRecords" name="optAllRecords" type="radio"
																								onclick="changeRecordOptions('all')"
																								onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																								onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																								onfocus="try{radio_onFocus(this);}catch(e){}"
																								onblur="try{radio_onBlur(this);}catch(e){}" />
																							<label
																								tabindex="-1"
																								for="optAllRecords"
																								class="radio"
																								onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																								onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}" />
																							All
																								</label>
																						</td>
																					</tr>
																					<tr>
																						<td width="95">
																							<input id="optPickList" name="optPickList" type="radio"
																								onclick="changeRecordOptions('picklist')"
																								onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																								onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																								onfocus="try{radio_onFocus(this);}catch(e){}"
																								onblur="try{radio_onBlur(this);}catch(e){}" />
																							<label
																								tabindex="-1"
																								for="optPickList"
																								class="radio"
																								onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																								onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}" />
																							Picklist
																									</label>
																						</td>
																						<td>
																							<input id="txtBasePicklist" name="txtBasePicklist" class="text textdisabled" disabled="disabled" style="width=250">
																						</td>
																						<td width="15">
																							<input id="cmdBasePicklist" name="cmdBasePicklist" class="btn btndisabled" disabled="disabled" type="button" value="..."
																								onclick="selectRecordOption('picklist')"
																								onmouseover="try{button_onMouseOver(this);}catch(e){}"
																								onmouseout="try{button_onMouseOut(this);}catch(e){}"
																								onfocus="try{button_onFocus(this);refreshControls();}catch(e){}"
																								onblur="try{button_onBlur(this);}catch(e){}" />
																						</td>
																					</tr>
																					<tr>
																						<td>
																							<input id="optFilter" name="optFilter" type="radio"
																								onclick="changeRecordOptions('filter')"
																								onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																								onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																								onfocus="try{radio_onFocus(this);}catch(e){}"
																								onblur="try{radio_onBlur(this);}catch(e){}" />
																							<label
																								tabindex="-1"
																								for="optFilter"
																								class="radio"
																								onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																								onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}" />
																							Filter
																								</label>
																						</td>
																						<td>
																							<input id="txtBaseFilter" name="txtBaseFilter" class="text textdisabled" disabled="disabled" style="width=250">
																						</td>
																						<td>
																							<input id="cmdBaseFilter" name="cmdBaseFilter" class="btn btndisabled" disabled="disabled" type="button" value="..."
																								onclick="selectRecordOption('filter')"
																								onmouseover="try{button_onMouseOver(this);}catch(e){}"
																								onmouseout="try{button_onMouseOut(this);}catch(e){}"
																								onfocus="try{button_onFocus(this);refreshControls();}catch(e){}"
																								onblur="try{button_onBlur(this);}catch(e){}" />
																						</td>
																					</tr>
																				</table>

																			</td>
																		</tr>
															</tr>
															<tr>
																<td>
																	<table width="100%" class="invisible" cellspacing="0">
																		<tr>
																			<td>
																				<input id="chkPrintInReportHeader" name="chkPrintInReportHeader" type="checkbox" tabindex="-1"
																					onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																					onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																				<label
																					for="chkPrintInReportHeader"
																					class="checkbox"
																					tabindex="0"
																					onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																					onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																					onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																					onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																					onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																					Display filter or picklist title in the report header
																				</label>
																			</td>
																		</tr>
																	</table>
																</td>
															</tr>
														</table></td>
												</tr>
											</table>
											</SPAN>
										</td>
									</tr>
								</table>
						</td>
					</tr>
				</table>

				</div>
						<!-- Second Tab (Options) -->
				<div id="div2" style="display: none">
					<table width="100%"  cellspacing="0" cellpadding="5">
						<tr>
							<td>
								<table width="100%" class="invisible" cellspacing="0" cellpadding="5">
									<td>
										<table class="invisible" cellspacing="0" cellpadding="0">
											<tr><td style="font-weight:bold">Display :</td></tr>
											<tr style="height: 10px"></tr>
											<tr>
												<td>
													<input type="checkbox" id="chkSRV" name="chkSRV" tabindex="-1"
														onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
														onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
													<label
														for="chkSRV"
														class="checkbox"
														tabindex="0"
														onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
														onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
														onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
														onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
														onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
														Suppress Repeated Personnel Details
													</label>
												</td>
											</tr>
											<tr>
												<td>
													<input type="checkbox" id="chkShowDurations" name="chkShowDurations" tabindex="-1"
														onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
														onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
													<label
														for="chkShowDurations"
														class="checkbox"
														tabindex="0"
														onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
														onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
														onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
														onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
														onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
														Show Duration Totals
													</label>
												</td>
											</tr>
											<tr>
												<td>
													<input type="checkbox" id="chkShowInstances" name="chkShowInstances" tabindex="-1"
														onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
														onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
													<label
														for="chkShowInstances"
														class="checkbox"
														tabindex="0"
														onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
														onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
														onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
														onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
														onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
														Show Instances Count
													</label>
												</td>
											</tr>
											<tr>
												<td>
													<input type="checkbox" id="chkShowFormula" name="chkShowFormula" tabindex="-1"
														onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
														onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
													<label
														for="chkShowFormula"
														class="checkbox"
														tabindex="0"
														onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
														onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
														onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
														onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
														onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
														Show Bradford Factor Formula
													</label>
												</td>
											</tr>
											<tr>
												<td>
													<input type="checkbox" id="chkShowAbsenceDetails" name="chkAbsenceDetails" tabindex="-1"
														onclick="refreshTab2Controls();"
														onchange="refreshTab2Controls();"
														onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
														onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
													<label
														for="chkShowAbsenceDetails"
														class="checkbox"
														tabindex="0"
														onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
														onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
														onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
														onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
														onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
														Show Absence Details
													</label>
												</td>
											</tr>
										</table>
									</td>
								</table>
								<table width="100%" class="invisible" cellspacing="0" cellpadding="5">
									<td>
										<table class="invisible" cellspacing="0" cellpadding="0">
											<tr><td style="font-weight:bold">Record Selection :</td></tr>
											<tr style="height: 10px"></tr>
											<tr>
												<td>
													<input type="checkbox" id="chkOmitBeforeStart" name="chkOmitBeforeStart" tabindex="-1"
														onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
														onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
													<label
														for="chkOmitBeforeStart"
														class="checkbox"
														tabindex="0"
														onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
														onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
														onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
														onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
														onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
														Omit absences starting before the report start date
													</label>
												</td>
											</tr>
											<tr>
												<td>
													<input type="checkbox" id="chkOmitAfterEnd" name="chkOmitAfterEnd" tabindex="-1"
														onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
														onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
													<label
														for="chkOmitAfterEnd"
														class="checkbox"
														tabindex="0"
														onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
														onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
														onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
														onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
														onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
														Omit absences ending after the report end date
													</label>
												</td>
											</tr>
											<tr>
												<td>
													<table class="invisible" cellspacing="0" cellpadding="0">
														<tr>
															<td>
																<input type="checkbox" id="chkMinimumBradfordFactor" name="chkMinimumBradfordFactor" tabindex="-1"
																	onclick="refreshTab2Controls();"
																	onchange="refreshTab2Controls();"
																	onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																	onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																<label
																	for="chkMinimumBradfordFactor"
																	class="checkbox"
																	tabindex="0"
																	onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																	onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																	onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																	onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																	onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																	Minimum Bradford Factor
																</label>
																&nbsp;
																&nbsp;
																		<input id="txtMinimumBradfordFactor" name="txtMinimumBradfordFactor" class="text"
																			onblur="validateNumeric(this);">
															</td>
														</tr>
													</table>
												</td>
											</tr>
										</table>
									</td>
								</table>
								<table width="100%" class="invisible" cellspacing="0" cellpadding="5">
									<td>
										<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
											<tr><td style="font-weight:bold">Order :</td></tr>
											<tr style="height: 10px"></tr>
											<tr>
												<td style="white-space: nowrap;width:15%" >Order By :</td>
												<td style="width:60%">
													<select id="cboOrderBy1" name="cboOrderBy1" style="width: 60%" class="combo"
														onchange="refreshTab2Controls();">
														<option value="0">&lt;None&gt;</option>
														<%
															For iCount = 0 To UBound(aColumnNames, 2) - 1
																Response.Write("<OPTION VALUE = " & """" & aColumnNames(0, iCount) & """" & ">" & aColumnNames(1, iCount) & "</OPTION>")
															Next
														%>
													</select>
												<%--</td>
												<td>--%>
													<input type="checkbox" id="chkOrderBy1Asc" name="chkOrderBy1Asc" tabindex="-1"
														onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
														onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
													<label
														for="chkOrderBy1Asc"
														class="checkbox"
														tabindex="0"
														onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
														onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
														onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
														onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
														onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
														Ascending
													</label>
												</td>
												<td></td>
											</tr>
											<tr>
												<td style="white-space: nowrap;width:15%">Then : </td>
												<td width="60%">
													<select id="cboOrderBy2" name="cboOrderBy2" style="width: 60%" class="combo"
														onchange="refreshTab2Controls();">
														<option value="0">&lt;None&gt;</option>
														<%
															For iCount = 0 To UBound(aColumnNames, 2) - 1
																Response.Write("<OPTION VALUE = " & """" & aColumnNames(0, iCount) & """" & ">" & aColumnNames(1, iCount) & "</OPTION>")
															Next
														%>
													</select>
												<%--</td>
												<td>--%>
													<input type="checkbox" id="chkOrderBy2Asc" name="chkOrderBy2Asc" tabindex="-1"
														onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
														onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
													<label
														for="chkOrderBy2Asc"
														class="checkbox"
														tabindex="0"
														onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
														onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
														onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
														onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
														onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
														Ascending
													</label>
												</td>
												<td style="width:100%"> </td>
											</tr>
										</table>
									</td>
								</table>
							</td>
						</tr>
					</table>
				</div>

				<!-- Third tab -->
				<div id="div3" style="visibility: hidden; display: none">
					<table width="100%" height="100%"  cellspacing="0" cellpadding="5">
						<tr valign="top">
							<td>
								<table width="100%" class="invisible" cellspacing="10" cellpadding="0">
									<tr>
										<td valign="top" rowspan="2" width="25%" height="100%">
											<table  cellspacing="0"  style="width:200px; height:100%; padding:4px">
												<tr height="10">
													<td height="10" align="left" valign="top"><strong>Output Format :</strong>
														<br>
														<br>
														<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
															<tr height="20">
																<td width="5">&nbsp</td>
																<td align="left" width="15">
																	<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat0" value="0"
																		onclick="formatClick(0);"
																		onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																		onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																		onfocus="try{radio_onFocus(this);}catch(e){}"
																		onblur="try{radio_onBlur(this);}catch(e){}" />
																</td>
																<td align="left" nowrap>
																	<label
																		tabindex="-1"
																		for="optOutputFormat0"
																		class="radio"
																		onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																		onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}" />
																	Data Only
																				</label>
																</td>
																<td width="5">&nbsp</td>
															</tr>
															<%
																'MH20040705
																'Don't allow CSV for Bradford
																If Session("StandardReport_Type") = 16 Then
															%>
															<input type="hidden" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat1" value="1"
																onclick="formatClick(1);"
																onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																onfocus="try{radio_onFocus(this);}catch(e){}"
																onblur="try{radio_onBlur(this);}catch(e){}" />
															<%
															Else
															%>
															<tr height="10">
																<td colspan="4"></td>
															</tr>
															<tr height="20">
																<td width="5">&nbsp</td>
																<td align="left" width="15">
																	<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat1" value="1"
																		onclick="formatClick(1);"
																		onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																		onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																		onfocus="try{radio_onFocus(this);}catch(e){}"
																		onblur="try{radio_onBlur(this);}catch(e){}" />
																</td>
																<td align="left" nowrap>
																	<label
																		tabindex="-1"
																		for="optOutputFormat1"
																		class="radio"
																		onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																		onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}" />
																	CSV File
																				</label>
																</td>
																<td width="5">&nbsp</td>
															</tr>
															<%
															End If
															%>
															<tr height="10">
																<td colspan="4"></td>
															</tr>
															<tr height="20">
																<td width="5">&nbsp</td>
																<td align="left" width="15">
																	<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat2" value="2"
																		onclick="formatClick(2);"
																		onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																		onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																		onfocus="try{radio_onFocus(this);}catch(e){}"
																		onblur="try{radio_onBlur(this);}catch(e){}" />
																</td>
																<td align="left" nowrap>
																	<label
																		tabindex="-1"
																		for="optOutputFormat2"
																		class="radio"
																		onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																		onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}" />
																	HTML Document
																				</label>
																</td>
																<td width="5">&nbsp</td>
															</tr>
															<tr height="10">
																<td colspan="4"></td>
															</tr>
															<tr height="20">
																<td width="5">&nbsp</td>
																<td align="left" width="15">
																	<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat3" value="3"
																		onclick="formatClick(3);"
																		onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																		onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																		onfocus="try{radio_onFocus(this);}catch(e){}"
																		onblur="try{radio_onBlur(this);}catch(e){}" />
																</td>
																<td align="left" nowrap>
																	<label
																		tabindex="-1"
																		for="optOutputFormat3"
																		class="radio"
																		onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																		onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}" />
																	Word Document
																				</label>
																</td>
																<td width="5">&nbsp</td>
															</tr>
															<tr height="10">
																<td colspan="4"></td>
															</tr>
															<tr height="20">
																<td width="5">&nbsp</td>
																<td align="left" width="15">
																	<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat4" value="4"
																		onclick="formatClick(4);"
																		onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																		onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																		onfocus="try{radio_onFocus(this);}catch(e){}"
																		onblur="try{radio_onBlur(this);}catch(e){}" />
																</td>
																<td align="left" nowrap>
																	<label
																		tabindex="-1"
																		for="optOutputFormat4"
																		class="radio"
																		onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																		onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}" >Excel Worksheet
																				</label>
																</td>
																<td width="5">&nbsp</td>
															</tr>
															<tr height="10">
																<td colspan="4"></td>
															</tr>

															<tr height="5">
																<td width="5">&nbsp</td>
																<td align="left" width="15">
																	<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat5" value="5"
																		onclick="formatClick(5);"
																		onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																		onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																		onfocus="try{radio_onFocus(this);}catch(e){}"
																		onblur="try{radio_onBlur(this);}catch(e){}" />
																</td>
																<td>
																	<label
																		tabindex="-1"
																		for="optOutputFormat5"
																		class="radio"
																		onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																		onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}" />
																	Excel Chart
																				</label>
																</td>
																<td width="5">&nbsp</td>
															</tr>
															<%
																'MH20031211 Fault 7787
																'Don't allow Pivot for Bradford
																If Session("StandardReport_Type") = 16 Then
															%>
															<input type="hidden" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat6" value="6"
																onclick="formatClick(6);"
																onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																onfocus="try{radio_onFocus(this);}catch(e){}"
																onblur="try{radio_onBlur(this);}catch(e){}" />
															<%
															Else
															%>
															<tr height="10">
																<td colspan="4"></td>
															</tr>
															<tr height="5">
																<td width="5">&nbsp</td>
																<td align="left" width="15">
																	<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat6" value="6"
																		onclick="formatClick(6);"
																		onmouseover="try{radio_onMouseOver(this);}catch(e){}"
																		onmouseout="try{radio_onMouseOut(this);}catch(e){}"
																		onfocus="try{radio_onFocus(this);}catch(e){}"
																		onblur="try{radio_onBlur(this);}catch(e){}" />
																</td>
																<td>
																	<label
																		tabindex="-1"
																		for="optOutputFormat6"
																		class="radio"
																		onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
																		onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}" />
																	Excel Pivot Table
																			</label>
																</td>
																<td width="5">&nbsp</td>
															</tr>
															<%
															End If
															%>
															<tr height="5">
																<td colspan="4"></td>
															</tr>
														</table>
													</td>
												</tr>
											</table>
										</td>
										<td valign="top" width="75%">
											<table  cellspacing="0" cellpadding="4" width="100%" height="100%">
												<tr height="10">
													<td height="10" align="left" valign="top"><strong>Output Destination(s) :</strong>
														<br>
														<br>
														<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
															<tr height="20">
																<td width="5">&nbsp</td>
																<td align="left" colspan="6" nowrap>
																	<input name="chkPreview" id="chkPreview" type="checkbox" disabled="disabled" tabindex="-1"
																		onclick="refreshTab3Controls();"
																		onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																		onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																	<label
																		for="chkPreview"
																		class="checkbox checkboxdisabled"
																		tabindex="0"
																		onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																		onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																		onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																		onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																		onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																		Preview on screen
																	</label>
																</td>
																<td width="5">&nbsp</td>
															</tr>
															<tr height="10">
																<td colspan="8"></td>
															</tr>
															<tr height="20">
																<td width="5">&nbsp</td>
																<td align="left" colspan="6" nowrap>
																	<input name="chkDestination0" id="chkDestination0" type="checkbox" disabled="disabled" tabindex="-1"
																		onclick="refreshTab3Controls();"
																		onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																		onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																	<label
																		for="chkDestination0"
																		class="checkbox checkboxdisabled"
																		tabindex="0"
																		onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																		onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																		onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																		onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																		onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																		Display output on screen
																	</label>
																</td>
																<td width="5">&nbsp</td>
															</tr>
															<tr height="10">
																<td colspan="8"></td>
															</tr>
															<tr height="20">
																<td width="5">&nbsp</td>
																<td align="left" nowrap>
																	<input name="chkDestination1" id="chkDestination1" type="checkbox" disabled="disabled" tabindex="-1"
																		onclick="refreshTab3Controls();"
																		onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																		onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																	<label
																		for="chkDestination1"
																		class="checkbox checkboxdisabled"
																		tabindex="0"
																		onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																		onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																		onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																		onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																		onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																		Send to printer 
																	</label>
																</td>
																<td width="30" nowrap>&nbsp</td>
																<td align="left" nowrap>Printer location : 
																</td>
																<td width="15">&nbsp</td>
																<td colspan="2">
																	<select id="cboPrinterName" name="cboPrinterName" class="combo" width="100%" style="WIDTH: 100%">
																	</select>
																</td>
																<td width="5">&nbsp</td>
															</tr>
															<tr height="10">
																<td colspan="8"></td>
															</tr>
															<tr height="20">
																<td width="5">&nbsp</td>
																<td align="left" nowrap>
																	<input name="chkDestination2" id="chkDestination2" type="checkbox" disabled="disabled" tabindex="-1"
																		onclick="refreshTab3Controls();"
																		onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
																		onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
																	<label
																		for="chkDestination2"
																		class="checkbox checkboxdisabled"
																		tabindex="0"
																		onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
																		onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
																		onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
																		onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
																		onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
																		Save to file - 
																	</label>
																</td>
																<td width="30" nowrap>&nbsp</td>
																<td align="left" nowrap>File name :   
																</td>
																<td width="15" nowrap>&nbsp</td>
																<td colspan="2">
																	<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																		<tr>
																			<td>
																				<input id="txtFilename" name="txtFilename" class="text textdisabled" disabled="disabled" style="WIDTH: 200">
																			</td>
																			<td width="25">
																				<input id="cmdFilename" name="cmdFilename" class="btn" style="WIDTH: 100%" type="button" value="..."
																					onclick="saveFile();"
																					onmouseover="try{button_onMouseOver(this);}catch(e){}"
																					onmouseout="try{button_onMouseOut(this);}catch(e){}"
																					onfocus="try{button_onFocus(this);}catch(e){}"
																					onblur="try{button_onBlur(this);}catch(e){}" />
																			</td>
																</td>
														</table>
													</td>
													<td width="5">&nbsp</td>
												</tr>

												<tr height="10">
													<td colspan="8"></td>
												</tr>
												<tr height="20">
													<td width="5">&nbsp</td>
													<td align="left" nowrap></td>
													<td width="30" nowrap>&nbsp</td>
													<td align="left" nowrap>If existing file :
													</td>
													<td width="15" nowrap>&nbsp</td>
													<td colspan="2" width="100%" nowrap>
														<select id="cboSaveExisting" name="cboSaveExisting" class="combo" width="100%" style="WIDTH: 100%">
														</select>
													</td>
													<td width="5">&nbsp</td>
												</tr>

												<tr height="10">
													<td colspan="8"></td>
												</tr>
												<tr height="20">
													<td width="5">&nbsp</td>
													<td align="left" nowrap>
														<input name="chkDestination3" id="chkDestination3" type="checkbox" disabled="disabled" tabindex="-1"
															onclick="refreshTab3Controls();"
															onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
															onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
														<label
															for="chkDestination3"
															class="checkbox checkboxdisabled"
															tabindex="0"
															onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
															onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
															onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
															onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
															onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
															Send as email - 
														</label>
													</td>
													<td width="30" nowrap>&nbsp</td>
													<td align="left" nowrap>Email group :   
													</td>
													<td width="15" nowrap>&nbsp</td>
													<td colspan="2">
														<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
															<tr>
																<td>
																	<input id="txtEmailGroup" name="txtEmailGroup" class="text textdisabled" disabled="disabled" style="WIDTH: 200">
																	<input id="txtEmailGroupID" name="txtEmailGroupID" type="hidden">
																</td>
																<td width="25">
																	<input id="cmdEmailGroup" name="cmdEmailGroup" class="btn" style="WIDTH: 100%" type="button" value="..."
																		onclick="selectEmailGroup();"
																		onmouseover="try{button_onMouseOver(this);}catch(e){}"
																		onmouseout="try{button_onMouseOut(this);}catch(e){}"
																		onfocus="try{button_onFocus(this);}catch(e){}"
																		onblur="try{button_onBlur(this);}catch(e){}" />
																</td>
													</td>
											</table>
										</td>
										<td width="5">&nbsp</td>
									</tr>
									<tr height="10">
										<td colspan="8"></td>
									</tr>
									<tr height="20">
										<td width="5">&nbsp</td>
										<td align="left">&nbsp</td>
										<td width="30" nowrap>&nbsp</td>
										<td align="left" nowrap>Email subject :   
										</td>
										<td width="15">&nbsp</td>
										<td colspan="2" width="100%" nowrap>
											<input id="txtEmailSubject" class="text textdisabled" disabled="disabled" maxlength="255" name="txtEmailSubject" style="WIDTH: 100%">
										</td>
										<td width="5">&nbsp</td>
									</tr>
									<tr height="10">
										<td colspan="8"></td>
									</tr>
									<tr height="20">
										<td width="5">&nbsp</td>
										<td align="left">&nbsp</td>
										<td width="30" nowrap>&nbsp</td>
										<td align="left" nowrap>Attach as :   
										</td>
										<td width="15">&nbsp</td>
										<td colspan="2" width="100%" nowrap>
											<input id="txtEmailAttachAs" class="text textdisabled" disabled="disabled" maxlength="255" name="txtEmailAttachAs" style="WIDTH: 100%">
										</td>
										<td width="5">&nbsp</td>
									</tr>
									<tr height="10">
										<td colspan="8"></td>
									</tr>
								</table>
							</td>
						</tr>
					</table>
			</td>
		</tr>
	</table>
	</td>
								</tr>
							</TABLE>
						</DIV>
													
					</TD>
					<td width="10"></td>
	</TR> 

				<tr height="10">
					<td colspan="3"></td>
				</tr>

	<tr style="height:10px">
		<td>
			<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
				<tr>
					<td style="width:80px; text-align: left;padding-left: 20px">
						<input type="button" id="cmdOK" name="cmdOK" class="btn" value="Run"
							onclick="absence_okClick()" />
					</td>
					<td>&nbsp;</td>
					<td>
						<input type="button" id="cmdBack" name="cmdBack" class="btn" value="Back"
							onclick="absence_returnToRecEdit()" />
					</td>
					<td width="10"></td>
				</tr>
			</table>
		</td>
		<td width="10"></td>
		<td width="10"></td>
	</tr>

	<tr height="5">
		<td colspan="3"></td>
	</tr>
	</TABLE>
		</TD>
	</TR>
</TABLE>

		<input type='hidden' id="txtDatabase" name="txtDatabase" value="<%=session("Database")%>">
	<input type="hidden" id="txtWordVer" name="txtWordVer" value="<%=Session("WordVer")%>">
	<input type="hidden" id="txtExcelVer" name="txtExcelVer" value="<%=Session("ExcelVer")%>">
	<input type="hidden" id="txtWordFormats" name="txtWordFormats" value="<%=Session("WordFormats")%>">
	<input type="hidden" id="txtExcelFormats" name="txtExcelFormats" value="<%=Session("ExcelFormats")%>">
	<input type="hidden" id="txtWordFormatDefaultIndex" name="txtWordFormatDefaultIndex" value="<%=Session("WordFormatDefaultIndex")%>">
	<input type="hidden" id="txtExcelFormatDefaultIndex" name="txtExcelFormatDefaultIndex" value="<%=Session("ExcelFormatDefaultIndex")%>">
</form>

<form id="frmAbsenceUseful" name="frmAbsenceUseful" style="visibility: hidden; display: none">
	<input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
	<input type="hidden" id="txtLoading" name="txtLoading" value="Y">
	<input type="hidden" id="txtCurrentBaseTableID" name="txtCurrentBaseTableID">
	<input type="hidden" id="txtCurrentChildTableID" name="txtCurrentChildTableID" value="0">
	<input type="hidden" id="txtTablesChanged" name="txtTablesChanged">
	<input type="hidden" id="txtSelectedColumnsLoaded" name="txtSelectedColumnsLoaded" value="0">
	<input type="hidden" id="txtSortLoaded" name="txtSortLoaded" value="0">
	<input type="hidden" id="txtRepetitionLoaded" name="txtRepetitionLoaded" value="0">
	<input type="hidden" id="txtChildsLoaded" name="txtChildsLoaded" value="0">
	<input type="hidden" id="txtChanged" name="txtChanged" value="0">
	<input type="hidden" id="txtUtilID" name="txtUtilID" value='<%=session("utilid")%>'>
	<input type="hidden" id="txtChildCount" name="txtChildCount" value='<%=session("childcount")%>'>
	<input type="hidden" id="txtHiddenChildFilterCount" name="txtHiddenChildFilterCount" value='<%=session("hiddenfiltercount")%>'>
	<input type="hidden" id="txtLockGridEvents" name="txtLockGridEvents" value="0">
	<%
		Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
		Response.Write("<INPUT type='hidden' id=txtAction name=txtAction value=" & Session("action") & ">" & vbCrLf)
	%>
</form>

<form action="util_run_promptedvalues" target="string(15)" method="post" id="frmPostDefinition" name="frmPostDefinition">
	<input type="hidden" id="txtRecordSelectionType" name="txtRecordSelectionType">
	<input type="hidden" id="txtFromDate" name="txtFromDate">
	<input type="hidden" id="txtToDate" name="txtToDate">
	<input type="hidden" id="txtBasePicklistID" name="txtBasePicklistID" value="0">
	<input type="hidden" id="txtBasePicklist" name="txtBasePicklist">
	<input type="hidden" id="txtBaseFilterID" name="txtBaseFilterID" value="0">
	<input type="hidden" id="txtBaseFilter" name="txtBaseFilter">
	<input type="hidden" id="txtAbsenceTypes" name="txtAbsenceTypes">
	<input type="hidden" id="txtSRV" name="txtSRV">
	<input type="hidden" id="txtShowDurations" name="txtShowDurations">
	<input type="hidden" id="txtShowInstances" name="txtShowInstances">
	<input type="hidden" id="txtShowFormula" name="txtShowFormula">
	<input type="hidden" id="txtOmitBeforeStart" name="txtOmitBeforeStart">
	<input type="hidden" id="txtOmitAfterEnd" name="txtOmitAfterEnd">
	<input type="hidden" id="txtOrderBy1" name="txtOrderBy1">
	<input type="hidden" id="txtOrderBy1ID" name="txtOrderBy1ID">
	<input type="hidden" id="txtOrderBy1Asc" name="txtOrderBy1Asc">
	<input type="hidden" id="txtOrderBy2" name="txtOrderBy2">
	<input type="hidden" id="txtOrderBy2ID" name="txtOrderBy2ID">
	<input type="hidden" id="txtOrderBy2Asc" name="txtOrderBy2Asc">
	<input type="hidden" id="txtMinimumBradfordFactor" name="txtMinimumBradfordFactor">
	<input type="hidden" id="txtMinimumBradfordFactorAmount" name="txtMinimumBradfordFactorAmount">
	<input type="hidden" id="txtDisplayBradfordDetail" name="txtDisplayBradfordDetail">
	<input type="hidden" id="txtPrintFPinReportHeader" name="txtPrintFPinReportHeader">
	<input type="hidden" id="txtRecSelCurrentID" name="txtRecSelCurrentID" value='<%=Session("optionRecordID")%>'>
	<input type="hidden" id="utiltype" name="utiltype" value='<%=Session("StandardReport_Type")%>'>
	<input type="hidden" id="utilid" name="utilid" value="0">
	<input type="hidden" id="utilname" name="utilname" value="Standard Report">
	<input type="hidden" id="action" name="action" value="run">
	<input type="hidden" id="txtSend_OutputPreview" name="txtSend_OutputPreview">
	<input type="hidden" id="txtSend_OutputFormat" name="txtSend_OutputFormat">
	<input type="hidden" id="txtSend_OutputScreen" name="txtSend_OutputScreen">
	<input type="hidden" id="txtSend_OutputPrinter" name="txtSend_OutputPrinter">
	<input type="hidden" id="txtSend_OutputPrinterName" name="txtSend_OutputPrinterName">
	<input type="hidden" id="txtSend_OutputSave" name="txtSend_OutputSave">
	<input type="hidden" id="txtSend_OutputSaveExisting" name="txtSend_OutputSaveExisting">
	<input type="hidden" id="txtSend_OutputEmail" name="txtSend_OutputEmail">
	<input type="hidden" id="txtSend_OutputEmailAddr" name="txtSend_OutputEmailAddr">
	<input type="hidden" id="txtSend_OutputEmailSubject" name="txtSend_OutputEmailSubject">
	<input type="hidden" id="txtSend_OutputEmailAttachAs" name="txtSend_OutputEmailAttachAs">
	<input type="hidden" id="txtSend_OutputFilename" name="txtSend_OutputFilename">
</form>

<!-- Stuff required to make record selection stuff work -->
<form id="frmCustomReportStuff" name="frmCustomReportStuff">
	<input type="hidden" id="baseHidden" name="baseHidden">
</form>

<form id="frmEmailSelection" name="frmEmailSelection" target="emailSelection" action="util_emailSelection.asp" method="post" style="visibility: hidden; display: none">
	<input type="hidden" id="EmailSelCurrentID" name="EmailSelCurrentID">
</form>

<form id="frmRecordSelection" name="frmRecordSelection" target="recordSelection" action="util_recordSelection.asp" method="post" style="visibility: hidden; display: none">
	<input type="hidden" id="recSelTable" name="recSelTable" value='base'>
	<input type="hidden" id="recSelType" name="recSelType">
	<input type="hidden" id="recSelTableID" name="recSelTableID">
	<input type="hidden" id="recSelCurrentID" name="recSelCurrentID" value='<%=Session("optionRecordID")%>'>
</form>

<form action="default_submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
	<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</form>

<form action="tbAddFromWaitingListFind_Submit" method="post" id="frmGotoOption" name="frmGotoOption">
	<%Html.RenderPartial("~/Views/Shared/gotoOption.ascx")%>
</form>

<!-- Form to return to record edit screen -->
<form action="emptyoption" method="post" id="frmRecordEdit" name="frmRecordEdit">
</form>

<form id="frmSessionInformation" name="frmSessionInformation" style="visibility: hidden; display: none">
	<input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
	<input type="hidden" id="txtLoading" name="txtLoading" value="Y">
	<input type="hidden" id="txtChanged" name="txtChanged" value="0">
	<input type="hidden" id="txtUtilID" name="txtUtilID" value='<% =session("StandardReport_Type")%>'>
	<%
		Dim cmdDefinition
		Dim prmModuleKey
		Dim prmParameterKey
		Dim prmParameterValue
				
		cmdDefinition = CreateObject("ADODB.Command")
		cmdDefinition.CommandText = "sp_ASRIntGetModuleParameter"
		cmdDefinition.CommandType = 4	' Stored procedure.
		cmdDefinition.ActiveConnection = Session("databaseConnection")

		prmModuleKey = cmdDefinition.CreateParameter("moduleKey", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
		cmdDefinition.Parameters.Append(prmModuleKey)
		prmModuleKey.value = "MODULE_PERSONNEL"

		prmParameterKey = cmdDefinition.CreateParameter("paramKey", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
		cmdDefinition.Parameters.Append(prmParameterKey)
		prmParameterKey.value = "Param_TablePersonnel"

		prmParameterValue = cmdDefinition.CreateParameter("paramValue", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
		cmdDefinition.Parameters.Append(prmParameterValue)

		Err.Clear()
		cmdDefinition.Execute()

		Response.Write("<INPUT type='hidden' id=txtPersonnelTableID name=txtPersonnelTableID value=" & cmdDefinition.Parameters("paramValue").Value & ">" & vbCrLf)
			
		cmdDefinition = Nothing
	%>
</form>

<script type="text/javascript">
	stdrpt_def_absence_window_onload();
</script>
