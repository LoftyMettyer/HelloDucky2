﻿<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>


<%		
	Dim objCrossTab As CrossTab
	Dim intCount As Integer
	Dim strCrossTabName As String
	Dim lngCount As Long
	Dim objUser As HR.Intranet.Server.clsSettings
	Dim lngLoopMin As Long
	Dim lngLoopMax As Long
	Dim strEmailAddresses As String
	Dim sErrorDescription As String

	Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

		
	Response.Write("<script type=""text/javascript"">" & vbCrLf)
	Response.Write("  //" & Session("CT_Mode") & vbCrLf)
	Response.Write("  function util_run_crosstabs_data_window_onload() {" & vbCrLf & vbCrLf)
	Response.Write("    $(""#reportdataframe"").attr(""data-framesource"", ""UTIL_RUN_CROSSTABSDATA"");" & vbCrLf & vbCrLf)	
	Response.Write("    $(""#reportframe"").show();" & vbCrLf)
	Response.Write("    $(""#reportdataframe"").hide();" & vbCrLf)
	Response.Write("    $(""#reportbreakdownframe"").hide();" & vbCrLf)
	Response.Write("    $(""#outputoptions"").hide();" & vbCrLf)
	Response.Write("    $(""#reportworkframe"").show();" & vbCrLf)
	Response.Write("    $(""#divReportButtons"").css(""visibility"", ""visible"");" & vbCrLf)

	Response.Write("    frmOriginalDefinition = OpenHR.getForm(""reportworkframe"",""frmOriginalDefinition"");" & vbCrLf)
	Response.Write("    frmExportData = OpenHR.getForm(""reportworkframe"",""frmExportData"");" & vbCrLf)
	Response.Write("    var ssOutputGrid;" & vbCrLf)
	Response.Write("    var fok;" & vbCrLf)
	
	Response.Write("    var colNames = [];" & vbCrLf)
	Response.Write("    var colData = [];" & vbCrLf)
	Response.Write("    var colMode = [];" & vbCrLf)
	Response.Write("    var value;" & vbCrLf)
	Response.Write("    var i;" & vbCrLf)
	Response.Write("    var sColumnName;" & vbCrLf)
	Response.Write("    var iCount2;" & vbCrLf)
	Response.Write("    var obj;" & vbCrLf)
	
	Response.Write("    ssOutputGrid = document.getElementById(""ssOutputGrid"");" & vbCrLf & vbCrLf)

	'**************************************
	' LOAD
	'**************************************

	If Session("CT_Mode") = "LOAD" Then
		'Populate Controls

		objCrossTab = Session("objCrossTab" & Session("CT_UtilID"))

		Response.Write("  AddToIntTypeCombo(""Count"",""0"");" & vbCrLf)
		Response.Write("  AddToIntTypeCombo(""Average"",""1"");" & vbCrLf)
		Response.Write("  AddToIntTypeCombo(""Maximum"",""2"");" & vbCrLf)
		Response.Write("  AddToIntTypeCombo(""Minimum"",""3"");" & vbCrLf)
		Response.Write("  AddToIntTypeCombo(""Total"",""4"");" & vbCrLf)
			 
		For intCount = 0 To objCrossTab.ColumnHeadingUbound(2)
			Response.Write("  AddToPgbCombo(""" & CleanStringForJavaScript(Left(objCrossTab.ColumnHeading(2, intCount), 255)) & """,""" & CStr(intCount) & """);" & vbCrLf)
		Next

		Response.Write("  txtPageColumn.value = """ & CleanStringForJavaScript(objCrossTab.PageBreakColumnName) & """;" & vbCrLf)
		Response.Write("  chkPercentType.checked = " & LCase(CStr(objCrossTab.ShowPercentage)) & ";" & vbCrLf)
		Response.Write("  chkPercentPage.checked = " & LCase(CStr(objCrossTab.PercentageOfPage)) & ";" & vbCrLf)
		Response.Write("  chkSuppressZeros.checked = " & LCase(CStr(objCrossTab.SuppressZeros)) & ";" & vbCrLf)
		Response.Write("  chkUse1000.checked = " & LCase(CStr(objCrossTab.Use1000Separator)) & ";" & vbCrLf)
		
		Response.Write("  txtIntersectionColumn.value = """ & CleanStringForJavaScript(objCrossTab.IntersectionColumnName) & """;" & vbCrLf)
		Response.Write("  cboIntersectionType.selectedIndex = " & CStr(objCrossTab.IntersectionType) & ";" & vbCrLf)

		If objCrossTab.PageBreakColumn = True Then
			Response.Write("  cboPage.selectedIndex = 0;" & vbCrLf)
		Else
			Response.Write("  control_disable(cboPage, true);" & vbCrLf)
		End If
			 
		Response.Write("  frmExportData.txtPreview.value = """ & objCrossTab.OutputPreview & """;" & vbCrLf)
		Response.Write("  frmExportData.txtFormat.value = " & objCrossTab.OutputFormat & ";" & vbCrLf)
		Response.Write("  frmExportData.txtScreen.value = """ & objCrossTab.OutputScreen & """;" & vbCrLf)
		Response.Write("  frmExportData.txtPrinter.value = """ & objCrossTab.OutputPrinter & """;" & vbCrLf)
		Response.Write("  frmExportData.txtPrinterName.value = """ & CleanStringForJavaScript(objCrossTab.OutputPrinterName) & """;" & vbCrLf)
		Response.Write("  frmExportData.txtSave.value = """ & objCrossTab.OutputSave & """;" & vbCrLf)
		Response.Write("  frmExportData.txtSaveExisting.value = """ & objCrossTab.OutputSaveExisting & """;" & vbCrLf)
		Response.Write("  frmExportData.txtEmail.value = """ & objCrossTab.OutputEmail & """;" & vbCrLf)
		Response.Write("  frmExportData.txtEmailAddr.value = " & objCrossTab.OutputEmailID & ";" & vbCrLf)
		Response.Write("  frmExportData.txtEmailAddrName.value = """ & CleanStringForJavaScript(objCrossTab.OutputEmailGroupName) & """;" & vbCrLf)
		Response.Write("  frmExportData.txtEmailSubject.value = """ & CleanStringForJavaScript(objCrossTab.OutputEmailSubject) & """;" & vbCrLf)
		Response.Write("  frmExportData.txtEmailAttachAs.value = """ & CleanStringForJavaScript(objCrossTab.OutputEmailAttachAs) & """;" & vbCrLf)
		Response.Write("  frmExportData.txtFileName.value = """ & CleanStringForJavaScript(objCrossTab.OutputFilename) & """;" & vbCrLf)
	End If

	
	'**************************************
	' LOAD / PRINT
	'**************************************
	'		 	If Session("CT_Mode") = "LOAD" Or _
	'
	If Session("CT_Mode") = "LOAD" Or _
			Session("CT_Mode") = "REFRESH" Or _
		 Session("CT_Mode") = "OUTPUTRUN" Or _
		 Session("CT_Mode") = "OUTPUTRUNTHENCLOSE" Then
		'Initalise Grid

		objCrossTab = Session("objCrossTab" & Session("CT_UtilID"))
		strCrossTabName = CleanStringForJavaScript(Replace(objCrossTab.CrossTabName, "&", "&&"))

		Response.Write("  colNames.push('');" & vbCrLf)
		Response.Write("	colMode.push({ name: '', classes: 'ui-state-default ui-widget-content ui-state-default ui-widget-header ui-state-default' });" & vbCrLf)
		
		For lngCount = 0 To objCrossTab.ColumnHeadingUbound(0)

			Dim headerCaption As String
			headerCaption = Replace(CleanStringForJavaScript(Left(objCrossTab.ColumnHeading(CLng(0), lngCount), 255)), "_", " ")
			headerCaption = CleanStringSpecialCharacters(headerCaption)

			Response.Write("  colNames.push('" & headerCaption & "');" & vbCrLf)
			Response.Write("	colMode.push({ name: '" & headerCaption & "', cellattr: function(rowId, value, rowObject, colModel, arrData) { return 'style=""text-align: right;""'; } });" & vbCrLf)
			
		Next

		If objCrossTab.CrossTabType <> 3 Then
			If Session("CT_Mode") = "LOAD" Then
				Session("CT_ShowPercentage") = objCrossTab.ShowPercentage
				Session("CT_PercentageOfPage") = objCrossTab.PercentageOfPage
				Session("CT_SuppressZeros") = objCrossTab.SuppressZeros
				Session("CT_IntersectionType") = objCrossTab.IntersectionType
				Session("CT_Use1000") = objCrossTab.Use1000Separator
				Session("CT_PageNumber") = 0
			End If
			
			Response.Write("  colNames.push('" & objCrossTab.IntersectionTypeValue(Session("CT_IntersectionType")) & "');" & vbCrLf)
			Response.Write("	colMode.push({ name: '" & objCrossTab.IntersectionTypeValue(Session("CT_IntersectionType")) & "', cellattr: function(rowId, value, rowObject, colModel, arrData) { return 'style=""text-align: right;""'; } });" & vbCrLf)

		End If
	End If


	'**************************************
	' LOAD / REFRESH
	'**************************************

	If Session("CT_Mode") = "LOAD" Or _
		 Session("CT_Mode") = "REFRESH" Then
		'PopulateScreen
		
		objCrossTab = Session("objCrossTab" & Session("CT_UtilID"))
		objCrossTab.IntersectionType = Session("CT_IntersectionType")
		objCrossTab.ShowPercentage = Session("CT_ShowPercentage")
		objCrossTab.PercentageOfPage = Session("CT_PercentageOfPage")
		objCrossTab.SuppressZeros = Session("CT_SuppressZeros")
		objCrossTab.Use1000Separator = Session("CT_Use1000")
		objCrossTab.BuildOutputStrings(CLng(Session("CT_PageNumber")))

		If objCrossTab.IntersectionColumn = True Then
			Response.Write("  control_disable(cboIntersectionType, false);" & vbCrLf)
		Else
			Response.Write("  cboIntersectionType.style.backgroundColor = 'threedface';" & vbCrLf)
		End If

		If objCrossTab.PageBreakColumn = True Then
			Response.Write("  control_disable(cboPage, false);" & vbCrLf)
			If objCrossTab.ShowPercentage = True Then
				Response.Write("  control_disable(chkPercentPage, false);" & vbCrLf)
			End If
		End If

		Response.Write("  control_disable(chkPercentType, false);" & vbCrLf)
		Response.Write("  control_disable(chkSuppressZeros, false);" & vbCrLf)
		Response.Write("  control_disable(chkUse1000, false);" & vbCrLf)
		
		Dim objData As String()
		For intCount = 1 To objCrossTab.OutputArrayDataUBound
								
			Response.Write("  obj = {};" & vbCrLf)
			objData = Split(objCrossTab.OutputArrayData(intCount), vbTab)
			For intCount2 = 0 To UBound(objData)
				If intCount2 = 0 Then	'The first column should be in bold
					Response.Write("  obj[colNames[" & intCount2 & "]] = '<span style=""font-weight: bold;"">" & objData(intCount2).Replace("<", "&lt;").Replace(">", "&gt;") & "</span>';" & vbCrLf)
				Else
					Response.Write("  obj[colNames[" & intCount2 & "]] = '" & objData(intCount2).Replace("<", "&lt;").Replace(">", "&gt;") & "';" & vbCrLf)
				End If
			Next
			Response.Write("  colData.push(obj);")
		Next


		' JDM - Fault 4849 - Disable these controls in absence breakdown mode
		If objCrossTab.CrossTabType = 3 Then
			Response.Write("  CrossTabPage.style.visibility = ""hidden"" ;" & vbCrLf)
			Response.Write("  chkPercentPage.style.visibility = ""hidden"" ;" & vbCrLf)
			Response.Write("  chkPercentType.style.visibility = ""hidden"" ;" & vbCrLf)
			Response.Write("  chkUse1000.style.visibility = ""hidden"" ;" & vbCrLf)
		End If

	End If


	If Session("CT_Mode") = "LOAD" Then
		If objCrossTab.OutputPreview = False Then
			Response.Write("        frmGetReportData.txtEmailGroupID.value = frmExportData.txtEmailAddr.value;" & vbCrLf)
			Response.Write("        ExportData(""EMAILGROUPTHENCLOSE"");" & vbCrLf)
		End If
	End If


	'**************************************
	' BREAKDOWN
	'**************************************

	If Session("CT_Mode") = "BREAKDOWN" Then

		objCrossTab = Session("objCrossTab" & Session("CT_UtilID"))
		objCrossTab.BuildBreakdownStrings(CLng(Session("CT_Hor")), CLng(Session("CT_Ver")), CLng(Session("CT_Pgb")))

		'Look up the Int Type Text from the Int Type Number...
		Response.Write("  var frmBreakdown = OpenHR.getForm(""reportbreakdownframe"", ""frmBreakdown"");" & vbCrLf)
		Response.Write("  document.getElementById('txtDataIntersectionType').value = cboIntersectionType.options[document.getElementById('txtDataIntersectionType').value].innerText;" & vbCrLf)
	
		Response.Write("  OpenHR.submitForm(frmBreakdown);" & vbCrLf)

		'**************************************
		' OUTPUTPROMPT
		'**************************************

	ElseIf Session("CT_Mode") = "OUTPUTPROMPT" Then

		Response.Write("  frmExportData.txtUtilType = " & Session("utiltype") & ");" & vbCrLf)
		Response.Write("  OpenHR.submitForm(frmExportData);" & vbCrLf)
				
				
		'**************************************
		' OUTPUTRUN
		'**************************************

	ElseIf Session("CT_Mode") = "OUTPUTPROMPT" Or _
				 Session("CT_Mode") = "OUTPUTRUN" Or _
				 Session("CT_Mode") = "OUTPUTRUNTHENCLOSE" Then
		
		If Session("CT_Mode") = "OUTPUTRUNTHENCLOSE" Then
			Response.Write("  try {" & vbCrLf)
			Response.Write("    if (frmOriginalDefinition.txtCancelPrint.value == 1) {" & vbCrLf)
			Response.Write("      window.parent.parent.raiseError('',false,true);" & vbCrLf)
			Response.Write("    }" & vbCrLf)
			Response.Write("    else if (ClientDLL.ErrorMessage != """") {" & vbCrLf)
			Response.Write("      window.parent.parent.raiseError(ClientDLL.ErrorMessage,false,false);" & vbCrLf)
			Response.Write("    }" & vbCrLf)
			Response.Write("    else {" & vbCrLf)
			Response.Write("      window.parent.parent.raiseError('',true,false);" & vbCrLf)
			Response.Write("    }" & vbCrLf)
			Response.Write("  }" & vbCrLf)
			Response.Write("  catch (e) {" & vbCrLf)
			Response.Write("  }" & vbCrLf)
		Else
			Response.Write("  sUtilTypeDesc = frmPopup.txtUtilTypeDesc.value;" & vbCrLf)
			Response.Write("  if (frmOriginalDefinition.txtCancelPrint.value == 1) {" & vbCrLf)
			Response.Write("    OpenHR.messageBox(sUtilTypeDesc+"" output failed.\n\nCancelled by user."",64,sUtilTypeDesc);" & vbCrLf)
			Response.Write("  }" & vbCrLf)
			Response.Write("  else if (ClientDLL.ErrorMessage == """") {" & vbCrLf)
			Response.Write("    OpenHR.messageBox(sUtilTypeDesc+"" output complete."",64,sUtilTypeDesc);" & vbCrLf)
			Response.Write("  }" & vbCrLf)
		End If


	ElseIf Session("CT_Mode") = "EMAILGROUP" Or _
				 Session("CT_Mode") = "EMAILGROUPTHENCLOSE" Then
		strEmailAddresses = ""
		If Session("CT_EmailGroupID") > 0 Then
					
			Try
				Dim rstEmailAddr = objDataAccess.GetDataTable("spASRIntGetEmailGroupAddresses", CommandType.StoredProcedure _
							, New SqlParameter("EmailGroupID", SqlDbType.Int) With {.Value = CleanNumeric(Session("CT_EmailGroupID"))})

				If Not rstEmailAddr Is Nothing Then
					For Each objRow In rstEmailAddr.Rows
						strEmailAddresses = strEmailAddresses & objRow(0).ToString() & ";"
					Next
				End If

			Catch ex As Exception
				sErrorDescription = "Error getting the email addresses for group." & vbCrLf & FormatError(ex.Message)
			End Try
			

		End If
		'Session("CT_EmailGroupAddr") = strEmailAddresses
		Response.Write("  frmExportData.txtEmailGroupAddr.value = """ & CleanStringForJavaScript(strEmailAddresses) & """;" & vbCrLf)

		If Session("CT_Mode") = "EMAILGROUPTHENCLOSE" Then
			Response.Write("  ExportData(""OUTPUTRUNTHENCLOSE"");" & vbCrLf)
		Else
			Response.Write("  ExportData(""OUTPUTRUN"");" & vbCrLf)
		End If
		 

	ElseIf Session("CT_Mode") = "" Then

		'Must be the first time this asp is called...
		Response.Write(" crosstab_loadAddRecords();" & vbCrLf)
	Else
		Response.Write("	$('#ssOutputGrid').jqGrid({data: colData, datatype: 'local', colNames: colNames, colModel: colMode, autowidth: true" & vbCrLf)
		Response.Write("    , rowNum:1000000")
		Response.Write("	  , ondblClickRow: function (rowId, iRow, iCol, e) {" & vbCrLf)
		Response.Write("	    	  if (iCol == 0) { return; } // Ignore double click on first column" & vbCrLf)
		Response.Write("	    	  var lngPage = cboPage.options[cboPage.selectedIndex].Value;" & vbCrLf)
		Response.Write("	    		var intType = cboIntersectionType.options[cboIntersectionType.selectedIndex].Value;" & vbCrLf)
		Response.Write("	    		var txtValue = $('#ssOutputGrid')[0].rows[iRow].cells[iCol].textContent;" & vbCrLf)
		Response.Write("	    		getBreakdown(iCol -1, iRow -1, lngPage, intType, txtValue);}" & vbCrLf)
		Response.Write("	, cmTemplate: { sortable: false, editable: true }});")
	End If
	
	Response.Write("  try {" & vbCrLf)
	Response.Write("    refreshCombo(""INTERSECTIONTYPE"");" & vbCrLf)
	Response.Write("    refreshCombo(""PAGE"");" & vbCrLf)
	Response.Write("    refreshCombo(""FILEFORMAT"");" & vbCrLf)
	Response.Write("  }" & vbCrLf)
	Response.Write("  catch(e) {" & vbCrLf)
	Response.Write("  }" & vbCrLf)

	Response.Write("}" & vbCrLf)
	Response.Write("</script>" & vbCrLf & vbCrLf)
%>

<script type="text/javascript">	

	function getCrossTabData(strMode, lngPageNumber, lngIntType, blnShowPer, blnPerPage, blnSupZeros, blnThousand) {

		control_disable(window.cboIntersectionType, true);
		control_disable(window.chkPercentPage, true);
		control_disable(window.chkPercentType, true);
		control_disable(window.chkSuppressZeros, true);
		control_disable(window.chkUse1000, true);
		control_disable(window.cboPage, true);
		
		$('#ssOutputGrid').jqGrid('clearGridData');
		$('#ssOutputGrid').jqGrid('GridUnload');
		
		var frmGetData = OpenHR.getForm("reportdataframe", "frmGetReportData");
		frmGetData.txtMode.value = strMode;
		frmGetData.txtPageNumber.value = lngPageNumber;
		frmGetData.txtIntersectionType.value = lngIntType;
		frmGetData.txtShowPercentage.value = blnShowPer;
		frmGetData.txtPercentageOfPage.value = blnPerPage;
		frmGetData.txtSuppressZeros.value = blnSupZeros;
		frmGetData.txtUse1000.value = blnThousand;
		OpenHR.submitForm(frmGetData);
	}

	function getBreakdown(lngHor, lngVer, lngPgb, txtIntType, txtCellValue) {

		var frmGetData = OpenHR.getForm("reportdataframe", "frmGetReportData");
		frmGetData.txtMode.value = "BREAKDOWN";
		frmGetData.txtHor.value = lngHor;
		frmGetData.txtVer.value = lngVer;
		frmGetData.txtPgb.value = lngPgb;
		frmGetData.txtIntersectionType.value = txtIntType;
		frmGetData.txtCellValue.value = txtCellValue;
		OpenHR.submitForm(frmGetData);
	}

	function ExportData(strMode) {
		var frmGetData = OpenHR.getForm("reportdataframe", "frmGetReportData");
		frmGetData.txtMode.value = strMode;
		OpenHR.submitForm(frmGetData);
	}

	function ViewExportOptions() {

		//var frmGetData = OpenHR.getForm("reportdataframe", "frmExportData");
		//OpenHR.submitForm(frmGetData,"outputoptions");
		output_setOptions();
				
		$("#reportworkframe").hide();
		$("#reportbreakdownframe").hide();
		$("#outputoptions").show();

	}

</script>

<form action="util_run_crosstabsDataSubmit" method="post" id="frmGetReportData" name="frmGetReportData">
	<input type="hidden" id="txtMode" name="txtMode" value="<%=Session("CT_Mode")%>">
	<input type="hidden" id="txtPageNumber" name="txtPageNumber" value="<%=Session("CT_PageNumber")%>">
	<input type="hidden" id="txtShowPercentage" name="txtShowPercentage" value="<%=Session("CT_ShowPercentage")%>">
	<input type="hidden" id="txtPercentageOfPage" name="txtPercentageOfPage" value="<%=Session("CT_PercentageOfPage")%>">
	<input type="hidden" id="txtSuppressZeros" name="txtSuppressZeros" value="<%=Session("CT_SuppressZeros")%>">
	<input type="hidden" id="txtUse1000" name="txtUse1000" value="<%=Session("CT_Use1000")%>">
	<input type="hidden" id="txtHor" name="txtHor" value="<%=Session("CT_Hor")%>">
	<input type="hidden" id="txtVer" name="txtVer" value="<%=Session("CT_Ver")%>">
	<input type="hidden" id="txtPgb" name="txtPgb" value="<%=Session("CT_Pgb")%>">
	<input type="hidden" id="txtIntersectionType" name="txtIntersectionType" value="<%=Session("CT_IntersectionType")%>">
	<input type="hidden" id="txtDataIntersectionType" name="txtDataIntersectionType" value="<%=Session("CT_IntersectionType")%>">
	<input type="hidden" id="txtCellValue" name="txtCellValue" value="<%=Session("CT_CellValue")%>">
	<input type="hidden" id="txtUtilID" name="txtUtilID" value="<%=Session("CT_UtilID")%>">
	<input type="hidden" id="txtEmailGroupID" name="txtEmailGroupID" value="<%=Session("CT_EmailGroupID")%>">
</form>


<textarea id="holdtext" style="display: none;"></textarea>

<iframe name="submit-iframe" style="display: none;"></iframe>


<script type="text/javascript">

	// Generated by the response.writes above
	util_run_crosstabs_data_window_onload();

</script>
