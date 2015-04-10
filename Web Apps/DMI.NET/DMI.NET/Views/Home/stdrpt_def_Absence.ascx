<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>

<%="" %>

<script src="<%: Url.LatestContent("~/bundles/utilities_standardreports")%>" type="text/javascript"></script>

<%
		
	' Clear the session action which is used to botch the prompted values screen in
	Session("action") = ""
	Session("optionaction") = OptionActionType.Empty

	Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
	Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
	
	' Settings objects
	Dim objSettings As New HR.Intranet.Server.clsSettings
	objSettings.SessionInfo = CType(Session("SessionContext"), SessionInfo)
	
	Dim aColumnNames
	Dim aAbsenceTypes() As String
	Dim sErrorDescription As String
	Dim iCount As Integer
	Dim iUtilType = CType(Session("utiltype"), UtilityType)
	
	' Load Absence Types and Personnel columns into array
	ReDim aColumnNames(1, 0)
	ReDim aAbsenceTypes(0)
	
	' Retreive the absence options	
	Try
		
		Dim rstReportColumns = objDataAccess.GetDataTable("sp_ASRIntGetColumns", CommandType.StoredProcedure _
			, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = objDatabase.GetModuleParameter("MODULE_PERSONNEL", "Param_TablePersonnel")})

		iCount = 0
		For Each objRow As DataRow In rstReportColumns.Rows

			If CInt(objRow("OLEType")) <> 2 Then
		
				aColumnNames(0, iCount) = CInt(objRow("ColumnID"))
				aColumnNames(1, iCount) = objRow("ColumnName").ToString()
			
				ReDim Preserve aColumnNames(1, UBound(aColumnNames, 2) + 1)
				iCount += 1
			Else
				'MsgBox(objRow("ColumnName").ToString())
			End If
			
			
			
		Next

		Dim rstTablesInfo = objDataAccess.GetDataTable("sp_ASRIntGetAbsenceTypes", CommandType.StoredProcedure)

		iCount = 0
		For Each objRow As DataRow In rstTablesInfo.Rows
			aAbsenceTypes(iCount) = objRow(0).ToString()
			ReDim Preserve aAbsenceTypes(UBound(aAbsenceTypes) + 1)
			iCount += 1
		Next
		
		
	Catch ex As Exception
		sErrorDescription = "The personnel column information could not be retrieved." & vbCrLf & FormatError(ex.Message)
		
	End Try


	' Set the default settings
	
	Dim dtStartDate As Date
	Dim dtEndDate As Date
	
	Dim strReportType As String = "AbsenceBreakdown"
	Dim strType As String
	Dim lngDefaultColumnID As Integer
	Dim lngConfigColumnID As Integer
	Dim strSaveExisting As String
	
	Response.Write("<script type=""text/javascript"">" & vbCrLf)

	Response.Write("function SetReportDefaults(){" & vbCrLf)
	Response.Write("   	var frmAbsenceDefinition = $('#frmAbsenceDefinition')[0];" & vbCrLf)
	Response.Write("   	var frmPostDefinition = $('#frmPostDefinition')[0];" & vbCrLf)
	
	' Type of standard report being run
	If iUtilType = UtilityType.utlBradfordFactor Then
		strReportType = "BradfordFactor"
	End If

	If iUtilType = UtilityType.utlAbsenceBreakdown Then
		strReportType = "AbsenceBreakdown"
		Response.Write("frmAbsenceDefinition.btnTab2.style.visibility = ""hidden"";" & vbCrLf)
	End If

	' Absence types
	For iCount = 0 To UBound(aAbsenceTypes) - 1
		If objDatabase.GetSystemSetting(strReportType, "Absence Type " & aAbsenceTypes(iCount), "0") = "1" Then
			Response.Write("frmAbsenceDefinition.chkAbsenceType_" & iCount & ".checked = 1;" & vbCrLf)
		End If
	Next

	' Report period	
	Dim rstReportDates = objDataAccess.GetDataTable("spASRIntGetStandardReportDates", CommandType.StoredProcedure, _
					New SqlParameter("piReportType", SqlDbType.Int) With {.Value = iUtilType})

	If rstReportDates.Rows.Count > 0 Then
		dtStartDate = CalculatePromptedDate(rstReportDates.Rows(0))
		dtEndDate = CalculatePromptedDate(rstReportDates.Rows(1))
	Else
		Dim thisMonth As New DateTime(DateTime.Today.Year, DateTime.Today.Month, 1)
		dtStartDate = thisMonth.AddYears(-1)
		dtEndDate = dtStartDate.AddYears(1).AddDays(-1)
	End If
			
	Response.Write("frmAbsenceDefinition.txtDateFrom.value = " & """" & ConvertSQLDateToLocale(dtStartDate) & """" & ";" & vbCrLf)
	Response.Write("frmAbsenceDefinition.txtDateTo.value = " & """" & ConvertSQLDateToLocale(dtEndDate) & """" & ";" & vbCrLf)

	
	' Record Selection
	If Session("optionRecordID") = "0" Then

		strType = objDatabase.GetSystemSetting(strReportType, "Type", "A").ToString()
		
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
				Response.Write("frmPostDefinition.txtBasePicklistID.value = " & CleanStringForJavaScript(objDatabase.GetSystemSetting(strReportType, "ID", "0")) & ";" & vbCrLf)
			Case "F"
				Response.Write("frmAbsenceDefinition.optAllRecords.checked = 0;" & vbCrLf)
				Response.Write("frmAbsenceDefinition.optPickList.checked = 0;" & vbCrLf)
				Response.Write("frmAbsenceDefinition.optFilter.checked = 1;" & vbCrLf)
				Response.Write("frmAbsenceDefinition.txtBaseFilter.value = " & """" & CleanStringForJavaScript(objSettings.GetPicklistFilterName(strReportType, strType)) & """" & ";" & vbCrLf)
				Response.Write("button_disable(frmAbsenceDefinition.cmdBaseFilter, false);" & vbCrLf)
				Response.Write("frmPostDefinition.txtBaseFilterID.value = " & CleanStringForJavaScript(objDatabase.GetSystemSetting(strReportType, "ID", "0")) & ";" & vbCrLf)
		End Select
	Else
		Response.Write("RecordSelection.style.visibility = ""hidden"";" & vbCrLf)
	End If

	' Display picklist in header
	Response.Write("frmAbsenceDefinition.chkPrintInReportHeader.checked =  " & CleanStringForJavaScript(objDatabase.GetSystemSetting(strReportType, "PrintFilterHeader", "0")) & vbCrLf)

	' Bradford Factor specific stuff
	If iUtilType = UtilityType.utlBradfordFactor Then
		Response.Write("frmAbsenceDefinition.chkSRV.checked = " & CleanStringForJavaScript(objDatabase.GetSystemSetting(strReportType, "SRV", "0")) & ";" & vbCrLf)
		Response.Write("frmAbsenceDefinition.chkShowDurations.checked = " & CleanStringForJavaScript(objDatabase.GetSystemSetting(strReportType, "Show Totals", "1")) & ";" & vbCrLf)
		Response.Write("frmAbsenceDefinition.chkShowInstances.checked = " & CleanStringForJavaScript(objDatabase.GetSystemSetting(strReportType, "Show Count", "0")) & ";" & vbCrLf)
		Response.Write("frmAbsenceDefinition.chkShowFormula.checked = " & CleanStringForJavaScript(objDatabase.GetSystemSetting(strReportType, "Show Workings", "0")) & ";" & vbCrLf)
		Response.Write("frmAbsenceDefinition.chkShowAbsenceDetails.checked = " & CleanStringForJavaScript(objDatabase.GetSystemSetting(strReportType, "Display Absence Details", "1")) & ";" & vbCrLf)
		Response.Write("frmAbsenceDefinition.chkOmitBeforeStart.checked = " & CleanStringForJavaScript(objDatabase.GetSystemSetting(strReportType, "Omit Before", "0")) & ";" & vbCrLf)
		Response.Write("frmAbsenceDefinition.chkOmitAfterEnd.checked = " & CleanStringForJavaScript(objDatabase.GetSystemSetting(strReportType, "Omit After", "0")) & ";" & vbCrLf)
		Response.Write("frmAbsenceDefinition.chkMinimumBradfordFactor.checked = " & CleanStringForJavaScript(objDatabase.GetSystemSetting(strReportType, "Minimum Bradford Factor", "0")) & ";" & vbCrLf)
		Response.Write("frmAbsenceDefinition.txtMinimumBradfordFactor.value = " & CleanStringForJavaScript(objDatabase.GetSystemSetting(strReportType, "Minimum Bradford Factor Amount", "0")) & ";" & vbCrLf)
		Response.Write("frmAbsenceDefinition.chkOrderBy1Asc.checked = " & CleanStringForJavaScript(objDatabase.GetSystemSetting(strReportType, "Order By Asc", "1")) & ";" & vbCrLf)
		Response.Write("frmAbsenceDefinition.chkOrderBy2Asc.checked = " & CleanStringForJavaScript(objDatabase.GetSystemSetting(strReportType, "Group By Asc", "1")) & ";" & vbCrLf)
			 
		lngDefaultColumnID = CInt(objDatabase.GetModuleParameter("MODULE_PERSONNEL", "Param_FieldsSurname"))
		lngConfigColumnID = CInt(objDatabase.GetSystemSetting(strReportType, "Order By", lngDefaultColumnID))
		'Response.Write "frmAbsenceDefinition.cboOrderBy1.value = " & """" & sFieldName & """" & ";" & vbcrlf
		Response.Write("for (var i=0; i<frmAbsenceDefinition.cboOrderBy1.options.length; i++)" & vbCrLf)
		Response.Write("	{" & vbCrLf)
		Response.Write("	if (frmAbsenceDefinition.cboOrderBy1.options[i].value == " & lngConfigColumnID & ")" & vbCrLf)
		Response.Write("		{" & vbCrLf)
		Response.Write("		frmAbsenceDefinition.cboOrderBy1.selectedIndex = i; " & vbCrLf)
		Response.Write("		}" & vbCrLf)
		Response.Write("	}" & vbCrLf)
		
		lngDefaultColumnID = objDatabase.GetModuleParameter("MODULE_PERSONNEL", "Param_FieldsForename")
		lngConfigColumnID = objDatabase.GetSystemSetting(strReportType, "Group By", lngDefaultColumnID)
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
	Select Case objDatabase.GetSystemSetting(strReportType, "Format", 0)
		Case "0"
			Response.Write("frmAbsenceDefinition.optDefOutputFormat0.checked = 1;" & vbCrLf)
		Case "1"
			Response.Write("frmAbsenceDefinition.optDefOutputFormat1.checked = 1;" & vbCrLf)
		Case "2"
			Response.Write("frmAbsenceDefinition.optDefOutputFormat2.checked = 1;" & vbCrLf)
		Case "3"
			Response.Write("frmAbsenceDefinition.optDefOutputFormat3.checked = 1;" & vbCrLf)
		Case "4"
			Response.Write("frmAbsenceDefinition.optDefOutputFormat4.checked = 1;" & vbCrLf)
		Case "5"
			Response.Write("frmAbsenceDefinition.optDefOutputFormat5.checked = 1;" & vbCrLf)
		Case "6"
			'MH20031211 Fault 7787
			'If Bradford then disallow Pivot (make it worksheet instead)
			If iUtilType = UtilityType.utlBradfordFactor Then
				Response.Write("frmAbsenceDefinition.optDefOutputFormat4.checked = 1;" & vbCrLf)
			Else
				Response.Write("frmAbsenceDefinition.optDefOutputFormat6.checked = 1;" & vbCrLf)
			End If
		Case Else
			' Charts and pivot not in Intranet yet
			Response.Write("frmAbsenceDefinition.optDefOutputFormat0.checked = 1" & vbCrLf)
	End Select
	
	Response.Write("frmAbsenceDefinition.chkPreview.checked = " & CleanStringForJavaScript(objDatabase.GetSystemSetting(strReportType, "Preview", 0)) & ";" & vbCrLf)

	Response.Write("strPrinterName = '';" & vbCrLf)
	Response.Write("frmAbsenceDefinition.chkDestination2.checked = " & CleanStringForJavaScript(objDatabase.GetSystemSetting(strReportType, "Save", 0)) & ";" & vbCrLf)
	Response.Write("frmAbsenceDefinition.txtFilename.value = " & """" & CleanStringForJavaScript(objDatabase.GetSystemSetting(strReportType, "FileName", "")) & """" & ";" & vbCrLf)
	
	Select Case objDatabase.GetSystemSetting(strReportType, "SaveExisting", 0)
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

	Response.Write("frmAbsenceDefinition.chkDestination3.checked = " & CleanStringForJavaScript(objDatabase.GetSystemSetting(strReportType, "Email", 0)) & ";" & vbCrLf)
	Response.Write("frmAbsenceDefinition.txtAbsenceEmailGroup.value = " & """" & CleanStringForJavaScript(objSettings.GetEmailGroupName(objSettings.GetSystemSetting(strReportType, "EmailAddr", "0"))) & """" & ";" & vbCrLf)
	Response.Write("frmAbsenceDefinition.txtAbsenceEmailGroupID.value = " & """" & CleanStringForJavaScript(objDatabase.GetSystemSetting(strReportType, "EmailAddr", "")) & """" & ";" & vbCrLf)
	Response.Write("frmAbsenceDefinition.txtEmailAttachAs.value = " & """" & CleanStringForJavaScript(objDatabase.GetSystemSetting(strReportType, "EmailAttachAs", "")) & """" & ";" & vbCrLf)
	Response.Write("frmAbsenceDefinition.txtEmailSubject.value = " & """" & CleanStringForJavaScript(objDatabase.GetSystemSetting(strReportType, "EmailSubject", "")) & """" & ";" & vbCrLf)

	Response.Write(vbCrLf & "}")
	Response.Write("</script>" & vbCrLf)

	objSettings = Nothing

%>

<div id="frmAbsenceDefinitiontabs">
	<form id="frmAbsenceDefinition" name="frmAbsenceDefinition">

		<div id="thefrmAbsenceDefinitionButtons" style="padding: 10px 0 50px 10px">
			<div style="float: left">
				<input type="button" class="btn" value="Definition" id="btnTab1" name="btnTab1" disabled="disabled"
					onclick="display_Absence_Page(1);" />
				<%
					If iUtilType = UtilityType.utlBradfordFactor Then
				%>
				<input class="btn" id="btnTab2" name="btnTab2" onclick="display_Absence_Page(2);" type="button" value="Options" />
				<%
				End If
				%>
			</div>
			<div style="float: left; padding-left: 5px">
				<input type="button" class="btn" value="Output" id="btnTab3" name="btnTab3"
					onclick="display_Absence_Page(3);" />
				<%
					' Causes problems if button isn't there
					If iUtilType <> UtilityType.utlBradfordFactor Then
				%>
				<input class="btn" id="btnTab2" name="btnTab2" onclick="display_Absence_Page(2);" type="button" value="Options" />
				<%
				End If
				%>
			</div>
		</div>

		<!-- First tab -->
		<div id="div1">
			<div>
				<fieldset>
					<legend class="fontsmalltitle">Absence Types :</legend>
					<ul id="limheight">
						<%
							For iCount = 0 To UBound(aAbsenceTypes) - 1
						%>
						<li>
							<input id="chkAbsenceType_<%=iCount%>" name="chkAbsenceType_<%=iCount%>"
								type="checkbox" tagname="<%=aAbsenceTypes(iCount)%>" tabindex="0" />
							<label for="chkAbsenceType_<%=iCount%>"
								class="checkbox"
								tabindex="-1">
								<%=aAbsenceTypes(iCount)%>
							</label>
							</></li>
						<%
						Next
						%>
					</ul>
				</fieldset>
			</div>

			<div class="width30 floatleft">
				<fieldset>
					<legend class="fontsmalltitle">Date Range :</legend>
					<div style="padding-left: 20px">
						<div class="formField">
							<label>Start :</label>
							<input class="datepicker" id="txtDateFrom" name="txtDateFrom">
						</div>
						<div class="formField">
							<label>End :</label>
							<input class="datepicker" id="txtDateTo" name="txtDateTo">
						</div>
					</div>
				</fieldset>
			</div>

			<div class="width45 floatleft" id="RecordSelection">
				<fieldset>
					<legend class="fontsmalltitle">Record Selection :</legend>
					<div class="padleft20">
						<div class="padbot5" style="padding-top: 10px">
							<input checked id="optAllRecords" name="optRecordSelection" type="radio" onclick="changeRecordOptions('ALL')" />
							<label tabindex="-1" for="optAllRecords">All</label>
						</div>

						<div class="padbot10">
							<input id="optPickList" name="optRecordSelection" type="radio" onclick="changeRecordOptions('PICKLIST')" />
							<label tabindex="-1" for="optPickList">Picklist</label>
							<div class="floatright">
								<input id="txtBasePicklist" name="txtBasePicklist"  disabled="disabled" style="width: 250px">
								<input id="cmdBasePicklist" name="cmdBasePicklist" class="btn btndisabled" disabled="disabled" type="button" value="..." onclick="selectAbsencePicklist()" />
							</div>
						</div>

						<div class="padbot10">
							<input id="optFilter" name="optRecordSelection" type="radio" onclick="changeRecordOptions('FILTER')" />
							<label for="optFilter" tabindex="-1">Filter</label>
							<div class="floatright">
								<input id="txtBaseFilter" name="txtBaseFilter"  disabled="disabled" style="width: 250px">
								<input id="cmdBaseFilter" name="cmdBaseFilter" class="btn btndisabled" disabled="disabled" type="button" value="..." onclick="selectAbsenceFilter()" />
							</div>
						</div>

						<div>
							<input id="chkPrintInReportHeader" name="chkPrintInReportHeader" type="checkbox" tabindex="0" />
							<label for="chkPrintInReportHeader" class="checkbox" tabindex="-1">
								Display filter or picklist title in the report header
							</label>
						</div>
					</div>
				</fieldset>
			</div>
		</div>

		<!-- Second Tab (Options) -->
		<div id="div2" style="display: none; padding-left: 10px">
			<table style="width: 100%; border-collapse: collapse; padding: 5px">
				<tr>
					<td>
						<table class="invisible floatleft" style="width: 40%; border-collapse: collapse">
							<tr>
								<td>
									<table class="invisible" style="border-collapse: collapse; padding: 0">
										<tr>
											<td class="fontsmalltitle" colspan="2">Display :</td>
										</tr>
										<tr style="height: 5px"></tr>
										<tr>
											<td class="padleft20"></td>
											<td>
												<input type="checkbox" id="chkSRV" name="chkSRV" tabindex="0" />
												<label class="checkbox" for="chkSRV" tabindex="-1">
													Suppress Repeated Personnel Details
												</label>
											</td>
										</tr>
										<tr>
											<td class="padleft20"></td>
											<td>
												<input type="checkbox" id="chkShowDurations" name="chkShowDurations" tabindex="0">
												<label class="checkbox" for="chkShowDurations" tabindex="-1">
													Show Duration Totals
												</label>
											</td>
										</tr>
										<tr>
											<td class="padleft20"></td>
											<td>
												<input type="checkbox" id="chkShowInstances" name="chkShowInstances" tabindex="0" />
												<label class="checkbox" for="chkShowInstances" tabindex="-1">
													Show Instances Count
												</label>
											</td>
										</tr>
										<tr>
											<td class="padleft20"></td>
											<td>
												<input type="checkbox" id="chkShowFormula" name="chkShowFormula" tabindex="0" />
												<label class="checkbox" for="chkShowFormula" tabindex="-1">
													Show Bradford Factor Formula
												</label>
											</td>
										</tr>
										<tr>
											<td class="padleft20"></td>
											<td>
												<input id="chkShowAbsenceDetails" name="chkAbsenceDetails" onchange="absenceBreakdownRefreshTab2Controls();" 
													onclick="absenceBreakdownRefreshTab2Controls();" tabindex="0" type="checkbox" />
												<label class="checkbox" for="chkShowAbsenceDetails" tabindex="-1">
													Show Absence Details
												</label>
											</td>
										</tr>
									</table>
								</td>
								<td>
						</table>
						<table class="invisible" style="width: 45%; border-collapse: collapse; padding: 5px">
							<tr>
								<td>
									<table class="invisible" style="border-collapse: collapse; padding: 0">
										<tr>
											<td class="fontsmalltitle">Record Selection :</td>
										</tr>
										<tr style="height: 5px"></tr>
										<tr>
											<td class="padleft20">
												<input type="checkbox" id="chkOmitBeforeStart" name="chkOmitBeforeStart" tabindex="0" />
												<label
													for="chkOmitBeforeStart"
													class="checkbox"
													tabindex="-1">
													Omit absences starting before the report start date
												</label>
											</td>
										</tr>
										<tr>
											<td class="padleft20">
												<input type="checkbox" id="chkOmitAfterEnd" name="chkOmitAfterEnd" tabindex="0" />
												<label
													for="chkOmitAfterEnd"
													class="checkbox"
													tabindex="-1">
													Omit absences ending after the report end date
												</label>
											</td>
										</tr>
										<tr>
											<td class="padleft20">
												<input type="checkbox" id="chkMinimumBradfordFactor" name="chkMinimumBradfordFactor" tabindex="0"
													onclick="absenceBreakdownRefreshTab2Controls();"
													onchange="absenceBreakdownRefreshTab2Controls();" />
												<label
													for="chkMinimumBradfordFactor"
													class="checkbox"
													tabindex="-1">
													Minimum Bradford Factor
												</label>
												<input id="Text1" name="txtMinimumBradfordFactor" class="width20" onblur="validateNumeric(this);">
											</td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<table class="invisible" style="width: 100%; border-collapse: collapse; padding: 5px">
							<tr>
								<td>
									<table class="invisible" style="width: 100%; padding: 0; border-collapse: collapse;">
										<tr>
											<td class="fontsmalltitle" colspan="3">Order :</td>
										</tr>

										<tr style="height: 40px"> 
											<td class="padleft20"></td>
											<td class="width10">Order By :</td>
											<td class="width90">
												<select id="cboOrderBy1" name="cboOrderBy1" onchange="absenceBreakdownRefreshTab2Controls();">
													<option value="0">None</option>
													<%
														For iCount = 0 To UBound(aColumnNames, 2) - 1
															
															
															
															Response.Write("<OPTION VALUE = " & """" & aColumnNames(0, iCount) & """" & ">" & aColumnNames(1, iCount) & "</OPTION>")
														Next
													%>
												</select>
												<input type="checkbox" id="chkOrderBy1Asc" name="chkOrderBy1Asc" tabindex="0" />
												<label for="chkOrderBy1Asc" class="checkbox" tabindex="-1">
													Ascending
												</label>
											</td>
										</tr>

										<tr>
											<td></td>
											<td>Then : </td>
											<td>
												<select id="cboOrderBy2" name="cboOrderBy2"  onchange="absenceBreakdownRefreshTab2Controls();">
													<option value="0">None</option>
													<%
														For iCount = 0 To UBound(aColumnNames, 2) - 1
															Response.Write("<OPTION VALUE = " & """" & aColumnNames(0, iCount) & """" & ">" & aColumnNames(1, iCount) & "</OPTION>")
														Next
													%>
												</select>
												<input type="checkbox" id="chkOrderBy2Asc" name="chkOrderBy2Asc" tabindex="0" />
												<label for="chkOrderBy2Asc" class="checkbox" tabindex="-1">
													Ascending
												</label>
											</td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</div>

		<!-- Third tab -->
		<div id="div3"  style="display: none;padding-left: 10px; padding-right: 20px" >
			<table class="width100" style="border-collapse: collapse">
				<tr>
					<td colspan="2" class="fontsmalltitle width20">Output Format :</td>
					<td colspan="2" class="fontsmalltitle">Output Destination :</td>
				</tr>

				<tr class="hidden">
					<td style="width: 15px"></td>
					<td class="width30" ></td>
					<td style="width: 20px"></td>
					<td ></td>
				</tr>

				<tr>
					<td></td>
					<td  class="vertaligntop">
							<table >
								<tr>
									<td>
										<table class="invisible" style="border-collapse: collapse; padding: 0; width: 100%">
											<tr style="height: 20px">
												<td></td>
												<td style="text-align: left; width: 15px">
													<input type="radio" name="optDefOutputFormat" id="optDefOutputFormat0" value="0"
														style="width: 20px"
														onclick="formatAbsenceClick(0);" />
												</td>
												<td style="text-align: left; white-space: nowrap">
													<label for="optDefOutputFormat0" tabindex="-1">
														Data Only
													</label>
												</td>
												<td></td>
											</tr>
											<%
												'MH20040705
												'Don't allow CSV for Bradford
												If iUtilType = UtilityType.utlBradfordFactor Then
											%>
											<input id="optDefOutputFormat1" name="optDefOutputFormat" onclick="formatAbsenceClick(1);" style="width: 20px" type="hidden" value="1" />
											<%
											Else
											%>
											<tr style="height: 10px">
												<td colspan="4"></td>
											</tr>
											<tr style="height: 20px">
												<td></td>
												<td style="text-align: left; width: 15px">
													<input id="optDefOutputFormat1" name="optDefOutputFormat" onclick="formatAbsenceClick(1);" style="width: 20px" type="radio" value="1" />
												</td>
												<td style="text-align: left; white-space: nowrap">
													<label class="ui-state-error-text" for="optDefOutputFormat1" tabindex="-1">CSV File</label>
												</td>
												<td></td>
											</tr>
											<%
											End If
											%>
											<tr style="height: 10px">
												<td colspan="4"></td>
											</tr>
											<tr style="height: 20px">
												<td></td>
												<td style="text-align: left; width: 15px">
													<input type="radio" style="width: 20px" name="optDefOutputFormat" id="optDefOutputFormat2" value="2"
														onclick="formatAbsenceClick(2);" />
												</td>
												<td style="text-align: left; white-space: nowrap">
													<label
														tabindex="-1"
														for="optDefOutputFormat2"
														class="radio ui-state-error-text">
														HTML Document
													</label>
												</td>
												<td></td>
											</tr>
											<tr style="height: 10px">
												<td colspan="4"></td>
											</tr>
											<tr style="height: 20px">
												<td></td>
												<td style="text-align: left; width: 15px">
													<input type="radio" style="width: 20px" name="optDefOutputFormat" id="optDefOutputFormat3" value="3"
														onclick="	formatAbsenceClick(3);" />
												</td>
												<td style="text-align: left; white-space: nowrap">
													<label
														tabindex="-1"
														for="optDefOutputFormat3"
														class="radio ui-state-error-text">
														Word Document
													</label>
												</td>
												<td></td>
											</tr>
											<tr style="height: 10px">
												<td colspan="4"></td>
											</tr>
											<tr style="height: 20px">
												<td></td>
												<td style="text-align: left; width: 15px">
													<input type="radio" style="width: 20px" name="optDefOutputFormat" id="optDefOutputFormat4" value="4"
														onclick="	formatAbsenceClick(4);" />
												</td>
												<td style="text-align: left; white-space: nowrap">
													<label
														tabindex="-1"
														for="optDefOutputFormat4"
														>
														Excel Worksheet
													</label>
												</td>
												<td></td>
											</tr>
											<tr style="height: 10px">
												<td colspan="4"></td>
											</tr>

											<tr style="height: 5px">
												<td></td>
												<td style="text-align: left; width: 15px">
													<input type="radio" style="width: 20px" name="optDefOutputFormat" id="optDefOutputFormat5" value="5"
														onclick="formatAbsenceClick(5);" />
												</td>
												<td>
													<label
														tabindex="-1"
														for="optDefOutputFormat5"
														>
														Excel Chart
													</label>
												</td>
												<td></td>
											</tr>
											<%
												'MH20031211 Fault 7787
												'Don't allow Pivot for Bradford
												If iUtilType = UtilityType.utlBradfordFactor Then
											%>
											<input id="optDefOutputFormat6" name="optDefOutputFormat" onclick="formatAbsenceClick(6);" style="width: 20px" type="hidden" value="6" />
											<%
											Else
											%>
											<tr style="height: 10px">
												<td colspan="4"></td>
											</tr>
											<tr style="height: 5px">
												<td></td>
												<td style="text-align: left; width: 15px">
													<input id="optDefOutputFormat6" name="optDefOutputFormat" onclick="formatAbsenceClick(6);" style="width: 20px" type="radio" value="6" />
												</td>
												<td>
													<label
														for="optDefOutputFormat6" tabindex="-1">
														Excel Pivot Table
													</label>
												</td>
												<td></td>
											</tr>
											<%
											End If
											%>
											<tr style="height: 5px">
												<td colspan="4"></td>
											</tr>
										</table>
									</td>
								</tr>
							</table>
					</td>
					<td style="width:10px"></td>
					<td  class="vertaligntop">
						<table id="tableOutputDestination" class="invisible" style="border-collapse: collapse; padding: 0; width: 100%;">

							<tr class="hidden">
								<td class="width25" style="height: 30px"></td>
								<td class="width20"  >
								<td colspan="2" class="width100" >
								<td  >
							</tr>

						<tr>
							
							<td  colspan="5">
								<input name="chkPreview" id="chkPreview" type="checkbox" disabled="disabled" tabindex="0"
									onclick="absenceBreakdownRefreshTab3Controls();" />
								<label for="chkPreview" class="checkbox" tabindex="-1">
									Preview on screen
								</label>
							</td>
							<%--<td colspan="4"></td>--%>
						</tr>

						<tr>
							<td  colspan="1">
								<input name="chkDestination2" id="chkDestination2" type="checkbox" disabled="disabled" tabindex="0"
									onclick="absenceBreakdownRefreshTab3Controls();" />
								<label for="chkDestination2" class="checkbox" tabindex="-1">
									Save to file
								</label>
							</td>
							<td >File name :</td>
							<td  colspan="2">
								<input  disabled="disabled" id="txtFilename" name="txtFilename" type="text">
							</td>
							<td >
								<input id="cmdFilename" name="cmdFilename" class="btn hidden" type="button" value="..."
									onclick="populateAbsenceFileName(frmAbsenceDefinition);" />
							</td>
						</tr>

						<tr>
							<td  colspan="1">
							<td>
								<label class="ui-state-error-text">If file exists :</label>
							</td>
							<td colspan="2">
								<select id="cboSaveExisting" name="cboSaveExisting">
									<option value="Overwrite">Overwrite</option>
									<option value="Do not overwrite">Do not overwrite</option>
									<option value="Add sequential number to name">Add sequential number to name</option>
									<option value="Append to file">Append to file</option>
									<option value="CreateNewSheet">Create new sheet in workbook</option>
								</select>
												
							</td>
							<td></td>
						</tr>

						<tr>
							
							<td  colspan="1">
								<input disabled="disabled" id="chkDestination3" name="chkDestination3" onclick="absenceBreakdownRefreshTab3Controls();" tabindex="0" type="checkbox" />
								<label class="checkbox" for="chkDestination3" tabindex="-1">
									Send as email
								</label>
							</td>
							<td>Group :</td>
							<td  colspan="2">
								<input  disabled="disabled" id="txtAbsenceEmailGroup" name="txtAbsenceEmailGroup"  type="text">
								<input id="txtAbsenceEmailGroupID" name="txtAbsenceEmailGroupID" type="hidden">
							</td>
							<td >
								<input class="btn" id="cmdEmailGroup" name="cmdEmailGroup" onclick="selectAbsenceEmailGroup();" type="button" value="..." />
							</td>
						</tr>

						<tr>
							<td colspan="1"></td>
							<td >Subject :</td>
							<td colspan="2">
								<input  disabled="disabled" id="txtEmailSubject" maxlength="255" name="txtEmailSubject"  type="text">
							</td>
							 <td></td>
						</tr>

						<tr>
							<td colspan="1"></td>
							<td >Attach as :</td>
							<td colspan="2">
								<input  disabled="disabled" id="txtEmailAttachAs" maxlength="255" name="txtEmailAttachAs"  type="text">
							</td>
							<td></td>
						</tr>
					</table>
					</td>
				</tr>

				<tr>
					<td></td>
					<td colspan="3"><span class="DataManagerOnly ui-state-error-text">Note: Options marked in red are unavailable in OpenHR Web.</span></td>
				</tr>
			</table>
		</div>
		
		<div id="RunBackButtons" style="visibility: hidden; float: left; padding: 10px">
			<input type="button" id="cmdRun" name="cmdRun" class="btn" value="Run"
				onclick="absence_okClick()" />
		</div>

		<input type="hidden" id="txtWordFormats" name="txtWordFormats" value="<%=Session("WordFormats")%>">
		<input type="hidden" id="txtExcelFormats" name="txtExcelFormats" value="<%=Session("ExcelFormats")%>">
		<input type="hidden" id="txtWordFormatDefaultIndex" name="txtWordFormatDefaultIndex" value="<%=Session("WordFormatDefaultIndex")%>">
		<input type="hidden" id="txtExcelFormatDefaultIndex" name="txtExcelFormatDefaultIndex" value="<%=Session("ExcelFormatDefaultIndex")%>">

		<input type="hidden" id="txtAction" name="txtAction" value="<%=Session("action")%>">
		<input type="hidden" id="txtUtilID" name="txtUtilID" value='<%=session("utilid")%>'>
	</form>
</div>

<form method="post" id="frmPostDefinition" name="frmPostDefinition">
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
	<input type="hidden" id="txtRecSelCurrentID" name="txtRecSelCurrentID" value='<%:Session("optionRecordID")%>'>
	<input type="hidden" id="utiltype" name="utiltype" value='<%:iUtilType%>'>
	<input type="hidden" id="ID" name="utilid" value='<%:session("utilid")%>'>
	<input type="hidden" id="Name" name="utilname" value="Standard Report">
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
	<input type="hidden" id="txtFilterName" name="txtFilterName">
	<input type="hidden" id="txtPicklistName" name="txtPicklistName">
	<%=Html.AntiForgeryToken()%>
	<% 	
		Dim sParameterValue As String = objDatabase.GetModuleParameter("MODULE_PERSONNEL", "Param_TablePersonnel")
		Response.Write("<input type='hidden' id='txtPersonnelTableID' name='txtPersonnelTableID' value=" & sParameterValue & ">" & vbCrLf)
	%>
</form>

<div style='height: 0; width: 0; overflow: hidden;'>
	<input type="hidden" id="recSelTableID" name="recSelTableID" value="<%:SettingsConfig.Personnel_EmpTableID%>">
	<input type="hidden" id="recSelCurrentID" name="recSelCurrentID" value='<%=Session("optionRecordID")%>'>
	<input id="cmdGetFilename" name="cmdGetFilename" type="file" />
</div>

<!-- Form to return to record edit screen -->
<form action="emptyoption" method="post" id="frmRecordEdit" name="frmRecordEdit">
</form>


<script type="text/javascript">
	stdrpt_def_absence_window_onload();

	menu_setVisibleMenuItem("mnutoolNewReportFind", false);
	menu_setVisibleMenuItem("mnutoolCopyReportFind", false);
	menu_setVisibleMenuItem("mnutoolEditReportFind", false);
	menu_setVisibleMenuItem("mnutoolDeleteReportFind", false);
	menu_setVisibleMenuItem("mnutoolPropertiesReportFind", false);
	menu_toolbarEnableItem("mnutoolRunReportFind", true);
	menu_setVisibleMenuItem("mnutoolRunReportFind", true);
	menu_setVisibleMenuItem('mnutoolCloseReportFind', false);

	//only display the 'close' button for defsel when called from rec edit...
	<%	If Not Session("optionRecordID") = "0" Then%>
	menu_setVisibleMenuItem('mnutoolCloseReportFind', true);
	menu_toolbarEnableItem('mnutoolCloseReportFind', true);
	<%	End If%>

	// Show and select the tab
	$("#toolbarReportFind").parent().show();
	$("#toolbarReportFind").click();

	$(".datepicker").datepicker();

	$(document).on('keydown', '.datepicker', function (event) {

		switch (event.keyCode) {
			case 113:
				$(this).datepicker("setDate", new Date());
				$(this).datepicker('widget').hide('true');
				break;
		}
	});

	$(document).on('blur', '.datepicker', function (sender) {
		if (OpenHR.IsValidDate(sender.target.value) == false && sender.target.value != "") {
			OpenHR.modalMessage("Invalid date value entered");
			$(sender.target.id).focus();
		}
	});

	$('table').attr('border', '0');
	$('fieldset').css("border", '0');
</script>
