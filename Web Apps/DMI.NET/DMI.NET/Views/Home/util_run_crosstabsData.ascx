<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Activities.Statements" %>


<%		
	Dim objCrossTab As CrossTab
	Dim intCount As Integer
	Dim lngCount As Long
	Dim strEmailAddresses As String
	Dim sErrorDescription As String

	Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

		
	Response.Write("<script type=""text/javascript"">" & vbCrLf)
	Response.Write("  //" & Session("CT_Mode") & vbCrLf)
	Response.Write("  function util_run_crosstabs_data_window_onload() {" & vbCrLf & vbCrLf)
	Response.Write("    $(""#reportdataframe"").attr(""data-framesource"", ""UTIL_RUN_CROSSTABSDATA"");" & vbCrLf & vbCrLf)	

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

		If CleanStringForJavaScript(objCrossTab.PageBreakColumnName) <> "<None>" Then
			Response.Write("  $('#txtPageColumn').text(""" & CleanStringForJavaScript(objCrossTab.PageBreakColumnName) & " :  " & """);" & vbCrLf)
		End If
		Response.Write("  chkPercentType.checked = " & LCase(CStr(objCrossTab.ShowPercentage)) & ";" & vbCrLf)
		Response.Write("  chkPercentPage.checked = " & LCase(CStr(objCrossTab.PercentageOfPage)) & ";" & vbCrLf)
		Response.Write("  chkSuppressZeros.checked = " & LCase(CStr(objCrossTab.SuppressZeros)) & ";" & vbCrLf)
		Response.Write("  chkUse1000.checked = " & LCase(CStr(objCrossTab.Use1000Separator)) & ";" & vbCrLf)
		
		If CleanStringForJavaScript(objCrossTab.IntersectionColumnName) <> "<None>" Then
			Response.Write("  $('#txtIntersectionColumn').text(""" & CleanStringForJavaScript(objCrossTab.IntersectionColumnName) & " :  " & """);" & vbCrLf)
		End If
		
		Response.Write("  cboIntersectionType.selectedIndex = " & CStr(objCrossTab.IntersectionType) & ";" & vbCrLf)

		If objCrossTab.PageBreakColumn = True Then
			Response.Write("  cboPage.selectedIndex = 0;" & vbCrLf)
		Else
			Response.Write("  control_disable(cboPage, true);" & vbCrLf)
		End If
			 
	End If

	
	'**************************************
	' LOAD / PRINT
	'**************************************
	If Session("CT_Mode") = "LOAD" Or _
			Session("CT_Mode") = "REFRESH" Then
		'Initalise Grid

		objCrossTab = Session("objCrossTab" & Session("CT_UtilID"))

		Response.Write("  colNames.push('');" & vbCrLf)
		Response.Write("	colMode.push({ name: '', classes: 'ui-state-default ui-widget-content ui-state-default ui-widget-header ui-state-default' });" & vbCrLf)
		
		For lngCount = 0 To objCrossTab.ColumnHeadingUbound(0)

			Dim headerCaption As String
			headerCaption = Replace(CleanStringForJavaScript(Left(objCrossTab.ColumnHeading(CLng(0), lngCount), 255)), "_", " ")
			headerCaption = CleanStringSpecialCharacters(headerCaption)

			Response.Write("  colNames.push('" & headerCaption & "');" & vbCrLf)
			Response.Write("	colMode.push({ name: '" & headerCaption & "', cellattr: function(rowId, value, rowObject, colModel, arrData) { return 'style=""text-align: right;""'; } });" & vbCrLf)
			
		Next

		If objCrossTab.CrossTabType <> CrossTabType.cttAbsenceBreakdown Then
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
		If objCrossTab.CrossTabType = CrossTabType.cttAbsenceBreakdown Then
			Response.Write("  CrossTabPage.style.visibility = ""hidden"" ;" & vbCrLf)
			Response.Write("  chkPercentPage.style.visibility = ""hidden"" ;" & vbCrLf)
			Response.Write("  chkPercentType.style.visibility = ""hidden"" ;" & vbCrLf)
			Response.Write("  chkUse1000.style.visibility = ""hidden"" ;" & vbCrLf)
		End If

	End If


	'If Session("CT_Mode") = "LOAD" Then
	'	If objCrossTab.OutputPreview = False Then
	'		Response.Write("        frmGetReportData.txtEmailGroupID.value = frmExportData.txtEmailAddr.value;" & vbCrLf)
	'		Response.Write("        ExportData(""EMAILGROUPTHENCLOSE"");" & vbCrLf)
	'	End If
	'End If


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

 
	ElseIf Session("CT_Mode") = "" Then

		'Must be the first time this asp is called...
		Response.Write(" crosstab_loadAddRecords();" & vbCrLf)
	Else
		If Session("utiltype") <> "35" Then
			Response.Write("$('#ssOutputGrid').show();" & vbCrLf)
			Response.Write("	$('#ssOutputGrid').jqGrid({data: colData, datatype: 'local', colNames: colNames, height: $('#main').height() * 0.8, colModel: colMode, width: $('#main').width() * 0.99" & vbCrLf)
			Response.Write("    , rowNum:1000000")
			Response.Write("	  , ondblClickRow: function (rowId, iRow, iCol, e) {" & vbCrLf)
			Response.Write("	    	  if (iCol == 0) { return; } // Ignore double click on first column" & vbCrLf)
			Response.Write("	    	  var lngPage = cboPage.options[cboPage.selectedIndex].value;" & vbCrLf)
			Response.Write("	    		var intType = cboIntersectionType.options[cboIntersectionType.selectedIndex].value;" & vbCrLf)
			Response.Write("	    		var txtValue = $('#ssOutputGrid')[0].rows[iRow].cells[iCol].textContent;" & vbCrLf)
			Response.Write("	    		getBreakdown(iCol -1, iRow -1, lngPage, intType, txtValue);}" & vbCrLf)
			Response.Write("	, cmTemplate: { sortable: false, editable: true }});")
		Else
			Response.Write("var rowNum = 1;" & vbCrLf)
			Response.Write("for (var i = 0; i < colData.length; i++) {" & vbCrLf)
			Response.Write("	if ((colData[i][''].indexOf('&lt') < 0) && (colData[i][''].indexOf('&gt') < 0) && (colData[i][''].indexOf('Total') < 0)) {" & vbCrLf)
			Response.Write("		var colNum = 1;" & vbCrLf)
			Response.Write("		var originalColNum = 0;" & vbCrLf)
			Response.Write("		for (var key in colData[i]) {			" & vbCrLf)
			Response.Write("			if (!((key.indexOf('&lt') >= 0) || (colData[i][key].indexOf('&lt') >= 0) ||" & vbCrLf)
			Response.Write("				(key.indexOf('&gt') >= 0) || (colData[i][key].indexOf('&gt') >= 0) ||" & vbCrLf)
			Response.Write("				(key.length === 0))) {" & vbCrLf)
			Response.Write("				var gridRefID = 'nineBoxR' + rowNum + 'C' + colNum;" & vbCrLf)
			Response.Write("				$('#' + gridRefID + '>p:last').html(colData[i][key]);" & vbCrLf)
			Response.Write("				$('#' + gridRefID).attr('data-row', i);" & vbCrLf)
			Response.Write("				$('#' + gridRefID).attr('data-col', originalColNum);" & vbCrLf)
			Response.Write("				$('#' + gridRefID).attr('data-titlevalue', $('#' + gridRefID + '>p:first').html());" & vbCrLf)
			Response.Write("				$('#' + gridRefID).attr('data-countvalue', colData[i][key]);" & vbCrLf)
			Response.Write("				$('#' + gridRefID).off('click').on('click', function () {" & vbCrLf)
			Response.Write("					var iRow = $(this).attr('data-row');" & vbCrLf)
			Response.Write("					var iCol = $(this).attr('data-col');" & vbCrLf)
			Response.Write("	    		var lngPage = cboPage.options[cboPage.selectedIndex].value;" & vbCrLf)
			Response.Write("	    		var intType = cboIntersectionType.options[cboIntersectionType.selectedIndex].value;" & vbCrLf)
			Response.Write("	    		var countValue = $(this).attr('data-countvalue');" & vbCrLf)
			Response.Write("	    		if (countValue == '') countValue = '0';" & vbCrLf)
			Response.Write("					getBreakdown(iCol, iRow, lngPage, intType, countValue);" & vbCrLf)
			Response.Write("				}); //End of click event handler" & vbCrLf)
			Response.Write("				colNum += 1;" & vbCrLf)
			Response.Write("			}" & vbCrLf)
			Response.Write("			 if(key.length > 0) originalColNum += 1;" & vbCrLf)
			Response.Write("		}" & vbCrLf)
			Response.Write("    rowNum += 1;" & vbCrLf)
			Response.Write("	}" & vbCrLf)
			Response.Write("}" & vbCrLf)
			Response.Write("$( '#tblNineBox td[id^=\'nineBoxR\']' ).hover(function() {   $( this ).addClass('ui-state-highlight');}, function() {$( this ).removeClass('ui-state-highlight'); });" & vbCrLf)
			Response.Write("$('#tblNineBox').show();" & vbCrLf)
		End If	
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
	<input type="hidden" id="txtMode" name="txtMode" value="<%=ValidateFromWhiteList(Session("CT_Mode"), Code.InputValidation.WhiteListCollections.CT_Modes)%>">
	<input type="hidden" id="txtPageNumber" name="txtPageNumber" value="<%:Session("CT_PageNumber")%>">
	<input type="hidden" id="txtShowPercentage" name="txtShowPercentage" value="<%:Session("CT_ShowPercentage")%>">
	<input type="hidden" id="txtPercentageOfPage" name="txtPercentageOfPage" value="<%:Session("CT_PercentageOfPage")%>">
	<input type="hidden" id="txtSuppressZeros" name="txtSuppressZeros" value="<%:Session("CT_SuppressZeros")%>">
	<input type="hidden" id="txtUse1000" name="txtUse1000" value="<%:Session("CT_Use1000")%>">
	<input type="hidden" id="txtHor" name="txtHor" value="<%=Session("CT_Hor")%>">
	<input type="hidden" id="txtVer" name="txtVer" value="<%=Session("CT_Ver")%>">
	<input type="hidden" id="txtPgb" name="txtPgb" value="<%=Session("CT_Pgb")%>">
	<input type="hidden" id="txtIntersectionType" name="txtIntersectionType" value="<%=ValidateIntegerValue(Session("CT_IntersectionType"))%>">
	<input type="hidden" id="txtDataIntersectionType" name="txtDataIntersectionType" value="<%=Session("CT_IntersectionType")%>">
	<input type="hidden" id="txtCellValue" name="txtCellValue" value="<%=Session("CT_CellValue")%>">
	<input type="hidden" id="txtUtilID" name="txtUtilID" value="<%=ValidateIntegerValue(Session("CT_UtilID"))%>">
	<input type="hidden" id="txtEmailGroupID" name="txtEmailGroupID" value="<%=ValidateIntegerValue(Session("CT_EmailGroupID"))%>">
	<%=Html.AntiForgeryToken()%>
</form>


<textarea id="holdtext" style="display: none;"></textarea>

<script type="text/javascript">

	// Generated by the response.writes above
	util_run_crosstabs_data_window_onload();

</script>
