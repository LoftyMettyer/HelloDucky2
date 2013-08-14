<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="ADODB" %>

<script src="<%: Url.Content("~/Scripts/FormScripts/crosstabdef.js")%>" type="text/javascript"></script>

<object classid="clsid:F9043C85-F6F2-101A-A3C9-08002B2F49FB"
	id="dialog"
	codebase="cabs/comdlg32.cab#Version=1,0,0,0"
	style="LEFT: 0px; TOP: 0px"
	viewastext>
	<param name="_ExtentX" value="847">
	<param name="_ExtentY" value="847">
	<param name="_Version" value="393216">
	<param name="CancelError" value="0">
	<param name="Color" value="0">
	<param name="Copies" value="1">
	<param name="DefaultExt" value="">
	<param name="DialogTitle" value="">
	<param name="FileName" value="">
	<param name="Filter" value="">
	<param name="FilterIndex" value="0">
	<param name="Flags" value="0">
	<param name="FontBold" value="0">
	<param name="FontItalic" value="0">
	<param name="FontName" value="">
	<param name="FontSize" value="8">
	<param name="FontStrikeThru" value="0">
	<param name="FontUnderLine" value="0">
	<param name="FromPage" value="0">
	<param name="HelpCommand" value="0">
	<param name="HelpContext" value="0">
	<param name="HelpFile" value="">
	<param name="HelpKey" value="">
	<param name="InitDir" value="">
	<param name="Max" value="0">
	<param name="Min" value="0">
	<param name="MaxFileSize" value="260">
	<param name="PrinterDefault" value="1">
	<param name="ToPage" value="0">
	<param name="Orientation" value="1">
</object>

	<form id="frmTables" style="visibility: hidden; display: none">
		<%
			Dim sErrorDescription = ""

			' Get the table records.
			Dim cmdTables As Command = New Command()
			cmdTables.CommandText = "sp_ASRIntGetCrossTabTablesInfo"
			cmdTables.CommandType = CommandTypeEnum.adCmdStoredProc
			cmdTables.ActiveConnection = Session("databaseConnection")
	
			Response.Write("<B>Set Connection</B>")
	
			Err.Clear()
			Dim rstTablesInfo = cmdTables.Execute
	
			Response.Write("<B>Executed SP</B>")
	
			If (Err.Number <> 0) Then
				sErrorDescription = "The tables information could not be retrieved." & vbCrLf & FormatError(Err.Description)
			End If

			If Len(sErrorDescription) = 0 Then
				' Dim iCount = 0
				Do While Not rstTablesInfo.EOF
					Response.Write("<input type='hidden' id=txtTableName_" & rstTablesInfo.fields("tableID").value & " name=txtTableName_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("tableName").value & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtTableType_" & rstTablesInfo.fields("tableID").value & " name=txtTableType_" & rstTablesInfo.fields("tableID").value & " value=" & rstTablesInfo.fields("tableType").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtTableChildren_" & rstTablesInfo.fields("tableID").value & " name=txtTableChildren_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("childrenString").value & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtTableChildrenNames_" & rstTablesInfo.fields("tableID").value & " name=txtTableChildrenNames_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("childrenNames").value & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtTableParents_" & rstTablesInfo.fields("tableID").value & " name=txtTableParents_" & rstTablesInfo.fields("tableID").value & " value=""" & rstTablesInfo.fields("parentsString").value & """>" & vbCrLf)

					rstTablesInfo.MoveNext()
				Loop

				' Release the ADO recordset object.
				rstTablesInfo.close()
				rstTablesInfo = Nothing
			End If
	
			' Release the ADO command object.
			cmdTables = Nothing
		%>
	</form>

	<form id="frmOriginalDefinition" name="frmOriginalDefinition" style="visibility: hidden; display: none">
		<%
			Dim sErrMsg = ""
			Dim lngHStart = 0
			Dim lngHStop = 0
			Dim lngHStep = 0
			Dim lngVStart = 0
			Dim lngVStop = 0
			Dim lngVStep = 0
			Dim lngPStart = 0
			Dim lngPStop = 0
			Dim lngPStep = 0

			If Session("action") <> "new" Then
				Dim cmdDefn As Command = New Command()
				cmdDefn.CommandText = "sp_ASRIntGetCrossTabDefinition"
				cmdDefn.CommandType = CommandTypeEnum.adCmdStoredProc
				cmdDefn.ActiveConnection = Session("databaseConnection")
								
				Dim prmUtilDefnID = cmdDefn.CreateParameter("utilid", 3, 1)	' 3=integer, 1=input
				cmdDefn.Parameters.Append(prmUtilDefnID)
				prmUtilDefnID.value = CleanNumeric(Session("utilid"))
								
				Dim prmUser = cmdDefn.CreateParameter("user", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
				cmdDefn.Parameters.Append(prmUser)
				prmUser.value = Session("username")

				Dim prmAction = cmdDefn.CreateParameter("action", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
				cmdDefn.Parameters.Append(prmAction)
				prmAction.value = Session("action")

				Dim prmErrMsg = cmdDefn.CreateParameter("errMsg", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmErrMsg)

				Dim prmName = cmdDefn.CreateParameter("name", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmName)

				Dim prmOwner = cmdDefn.CreateParameter("owner", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmOwner)

				Dim prmDescription = cmdDefn.CreateParameter("description", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmDescription)

				Dim prmBaseTableID = cmdDefn.CreateParameter("baseTableID", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmBaseTableID)

				Dim prmAllRecords = cmdDefn.CreateParameter("allRecords", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmAllRecords)

				Dim prmPicklistID = cmdDefn.CreateParameter("picklistID", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmPicklistID)

				Dim prmPicklistName = cmdDefn.CreateParameter("picklistName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmPicklistName)

				Dim prmPicklistHidden = cmdDefn.CreateParameter("picklistHidden", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmPicklistHidden)

				Dim prmFilterID = cmdDefn.CreateParameter("filterID", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmFilterID)

				Dim prmFilterName = cmdDefn.CreateParameter("filterName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmFilterName)

				Dim prmFilterHidden = cmdDefn.CreateParameter("filterHidden", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmFilterHidden)
		
				Dim prmPrintFilter = cmdDefn.CreateParameter("PrintFilter", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmPrintFilter)

				Dim prmHColID = cmdDefn.CreateParameter("HColID", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmHColID)

				Dim prmHStart = cmdDefn.CreateParameter("HStart", 200, 2, 20)	'3=integer, 2=output, 20=size
				cmdDefn.Parameters.Append(prmHStart)

				Dim prmHStop = cmdDefn.CreateParameter("HStop", 200, 2, 20)	'3=integer, 2=output, 20=size
				cmdDefn.Parameters.Append(prmHStop)

				Dim prmHStep = cmdDefn.CreateParameter("HStep", 200, 2, 20)	'3=integer, 2=output, 20=size
				cmdDefn.Parameters.Append(prmHStep)

				Dim prmVColID = cmdDefn.CreateParameter("VColID", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmVColID)

				Dim prmVStart = cmdDefn.CreateParameter("VStart", 200, 2, 20)	'3=integer, 2=output, 20=size
				cmdDefn.Parameters.Append(prmVStart)

				Dim prmVStop = cmdDefn.CreateParameter("VStop", 200, 2, 20)	'3=integer, 2=output, 20=size
				cmdDefn.Parameters.Append(prmVStop)

				Dim prmVStep = cmdDefn.CreateParameter("VStep", 200, 2, 20)	'3=integer, 2=output, 20=size
				cmdDefn.Parameters.Append(prmVStep)

				Dim prmPColID = cmdDefn.CreateParameter("PColID", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmPColID)

				Dim prmPStart = cmdDefn.CreateParameter("PStart", 200, 2, 20)	'3=integer, 2=output, 20=size
				cmdDefn.Parameters.Append(prmPStart)

				Dim prmPStop = cmdDefn.CreateParameter("PStop", 200, 2, 20)	'3=integer, 2=output, 20=size
				cmdDefn.Parameters.Append(prmPStop)

				Dim prmPStep = cmdDefn.CreateParameter("PStep", 200, 2, 20)	'3=integer, 2=output, 20=size
				cmdDefn.Parameters.Append(prmPStep)

				Dim prmIType = cmdDefn.CreateParameter("IType", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmIType)

				Dim prmIColID = cmdDefn.CreateParameter("IColID", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmIColID)

				Dim prmPercentage = cmdDefn.CreateParameter("Percentage", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmPercentage)

				Dim prmPerPage = cmdDefn.CreateParameter("PerPage", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmPerPage)

				Dim prmSuppress = cmdDefn.CreateParameter("Suppress", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmSuppress)

				Dim prmThousand = cmdDefn.CreateParameter("Thousand", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmThousand)

				Dim prmOutputPreview = cmdDefn.CreateParameter("outputPreview", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmOutputPreview)
		
				Dim prmOutputFormat = cmdDefn.CreateParameter("outputFormat", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmOutputFormat)
		
				Dim prmOutputScreen = cmdDefn.CreateParameter("outputScreen", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmOutputScreen)
		
				Dim prmOutputPrinter = cmdDefn.CreateParameter("outputPrinter", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmOutputPrinter)
		
				Dim prmOutputPrinterName = cmdDefn.CreateParameter("outputPrinterName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmOutputPrinterName)
		
				Dim prmOutputSave = cmdDefn.CreateParameter("outputSave", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmOutputSave)
		
				Dim prmOutputSaveExisting = cmdDefn.CreateParameter("outputSaveExisting", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmOutputSaveExisting)
		
				Dim prmOutputEmail = cmdDefn.CreateParameter("outputEmail", 11, 2) '11=bit, 2=output
				cmdDefn.Parameters.Append(prmOutputEmail)
		
				Dim prmOutputEmailAddr = cmdDefn.CreateParameter("outputEmailAddr", 3, 2)	'3=integer, 2=output
				cmdDefn.Parameters.Append(prmOutputEmailAddr)

				Dim prmOutputEmailAddrName = cmdDefn.CreateParameter("outputEmailAddrName", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmOutputEmailAddrName)

				Dim prmOutputEmailSubject = cmdDefn.CreateParameter("outputEmailSubject", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmOutputEmailSubject)

				Dim prmOutputEmailAttachAs = cmdDefn.CreateParameter("outputEmailAttachAs", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmOutputEmailAttachAs)

				Dim prmOutputFilename = cmdDefn.CreateParameter("outputFilename", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefn.Parameters.Append(prmOutputFilename)

				Dim prmTimestamp = cmdDefn.CreateParameter("timestamp", 3, 2)	' 3=integer, 2=output
				cmdDefn.Parameters.Append(prmTimestamp)

				Err.Clear()
				cmdDefn.Execute()

				Dim iHiddenCalcCount As Integer = 0
				If (Err.Number <> 0) Then
					sErrMsg = "'" & Session("utilname") & "' cross tab definition could not be read." & vbCrLf & FormatError(Err.Description)
				Else

					'rstDefinition.close
					'set rstDefinition = nothing

					' NB. IMPORTANT ADO NOTE.
					' When calling a stored procedure which returns a recordset AND has output parameters
					' you need to close the recordset and set it to nothing before using the output parameters. 
					If Len(cmdDefn.Parameters("errMsg").value) > 0 Then
						sErrMsg = "'" & Session("utilname") & "' " & cmdDefn.Parameters("errMsg").value
					End If

					lngHStart = cmdDefn.Parameters("HStart").value
					lngHStop = cmdDefn.Parameters("HStop").value
					lngHStep = cmdDefn.Parameters("HStep").value
					lngVStart = cmdDefn.Parameters("VStart").value
					lngVStop = cmdDefn.Parameters("VStop").value
					lngVStep = cmdDefn.Parameters("VStep").value
					lngPStart = cmdDefn.Parameters("PStart").value
					lngPStop = cmdDefn.Parameters("PStop").value
					lngPStep = cmdDefn.Parameters("PStep").value

					Response.Write("<input type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Replace(cmdDefn.Parameters("name").value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Owner name=txtDefn_Owner value=""" & Replace(cmdDefn.Parameters("owner").value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Description name=txtDefn_Description value=""" & Replace(cmdDefn.Parameters("description").value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_BaseTableID name=txtDefn_BaseTableID value=" & cmdDefn.Parameters("baseTableID").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_AllRecords name=txtDefn_AllRecords value=" & cmdDefn.Parameters("allRecords").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_PicklistID name=txtDefn_PicklistID value=" & cmdDefn.Parameters("picklistID").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_PicklistName name=txtDefn_PicklistName value=""" & Replace(cmdDefn.Parameters("picklistName").value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_PicklistHidden name=txtDefn_PicklistHidden value=" & cmdDefn.Parameters("picklistHidden").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_FilterID name=txtDefn_FilterID value=" & cmdDefn.Parameters("filterID").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_FilterName name=txtDefn_FilterName value=""" & Replace(cmdDefn.Parameters("filterName").value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_FilterHidden name=txtDefn_FilterHidden value=" & cmdDefn.Parameters("filterHidden").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_FilterHeader name=txtDefn_FilterHeader value=" & cmdDefn.Parameters("PrintFilter").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_PrintFilter name=txtDefn_PrintFilter value=" & cmdDefn.Parameters("PrintFilter").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_HColID name=txtDefn_HColID value=" & cmdDefn.Parameters("HColID").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_HStart name=txtDefn_HStart value=" & cmdDefn.Parameters("HStart").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_HStop name=txtDefn_HStop value=" & cmdDefn.Parameters("HStop").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_HStep name=txtDefn_HStep value=" & cmdDefn.Parameters("HStep").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_VColID name=txtDefn_VColID value=" & cmdDefn.Parameters("VColID").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_VStart name=txtDefn_VStart value=" & cmdDefn.Parameters("VStart").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_VStop name=txtDefn_VStop value=" & cmdDefn.Parameters("VStop").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_VStep name=txtDefn_VStep value=" & cmdDefn.Parameters("VStep").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_PColID name=txtDefn_PColID value=" & cmdDefn.Parameters("PColID").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_PStart name=txtDefn_PStart value=" & cmdDefn.Parameters("PStart").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_PStop name=txtDefn_PStop value=" & cmdDefn.Parameters("PStop").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_PStep name=txtDefn_PStep value=" & cmdDefn.Parameters("PStep").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_IType name=txtDefn_IType value=" & cmdDefn.Parameters("IType").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_IColID name=txtDefn_IColID value=" & cmdDefn.Parameters("IColID").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Percentage name=txtDefn_Percentage value=" & cmdDefn.Parameters("Percentage").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_PerPage name=txtDefn_PerPage value=" & cmdDefn.Parameters("PerPage").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Suppress name=txtDefn_Suppress value=" & cmdDefn.Parameters("Suppress").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_Use1000 name=txtDefn_Use1000 value=" & cmdDefn.Parameters("Thousand").value & ">" & vbCrLf)

					Response.Write("<input type='hidden' id=txtDefn_OutputPreview name=txtDefn_OutputPreview value=" & cmdDefn.Parameters("OutputPreview").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputFormat name=txtDefn_OutputFormat value=" & cmdDefn.Parameters("OutputFormat").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputScreen name=txtDefn_OutputScreen value=" & cmdDefn.Parameters("OutputScreen").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputPrinter name=txtDefn_OutputPrinter value=" & cmdDefn.Parameters("OutputPrinter").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputPrinterName name=txtDefn_OutputPrinterName value=""" & cmdDefn.Parameters("OutputPrinterName").value & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputSave name=txtDefn_OutputSave value=" & cmdDefn.Parameters("OutputSave").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputSaveExisting name=txtDefn_OutputSaveExisting value=" & cmdDefn.Parameters("OutputSaveExisting").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputEmail name=txtDefn_OutputEmail value=" & cmdDefn.Parameters("OutputEmail").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputEmailAddr name=txtDefn_OutputEmailAddr value=" & cmdDefn.Parameters("OutputEmailAddr").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputEmailAddrName name=txtDefn_OutputEmailName value=""" & Replace(cmdDefn.Parameters("OutputEmailAddrName").value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputEmailSubject name=txtDefn_OutputEmailSubject value=""" & Replace(cmdDefn.Parameters("OutputEmailSubject").value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputEmailAttachAs name=txtDefn_OutputEmailAttachAs value=""" & Replace(cmdDefn.Parameters("OutputEmailAttachAs").value, """", "&quot;") & """>" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_OutputFilename name=txtDefn_OutputFilename value=""" & cmdDefn.Parameters("OutputFilename").value & """>" & vbCrLf)

					Response.Write("<input type='hidden' id=txtDefn_Timestamp name=txtDefn_Timestamp value=" & cmdDefn.Parameters("timestamp").value & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=txtDefn_HiddenCalcCount name=txtDefn_HiddenCalcCount value=" & iHiddenCalcCount & ">" & vbCrLf)
					Response.Write("<input type='hidden' id=session_action name=session_action value=" & Session("action") & ">" & vbCrLf)
					Response.Write("</form>" & vbCrLf)

				End If

				' Release the ADO command object.
				cmdDefn = Nothing

				If Len(sErrMsg) > 0 Then
					Session("confirmtext") = sErrMsg
					Session("confirmtitle") = "OpenHR Intranet"
					Session("followpage") = "defsel"
					Session("reaction") = "CROSSTABS"
					Response.Clear()
					Response.Redirect("confirmok")
				End If
	
			Else
				Session("childcount") = 0
				Session("hiddenfiltercount") = 0
			End If
		%>
	</form>

	<form id="frmDefinition" name="frmDefinition">
		<table valign="top" align="center" class="outline" cellpadding="5" width="100%" height="100%" cellspacing="0">
			<tr>
				<td colspan="2">
					<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
						<tr height="5"><td colspan="3"></td></tr>

						<tr height="10">
							<td width="10"></td>
							<td>
								<input type="button" value="Definition" id="btnTab1" name="btnTab1" class="btn btndisabled" disabled="disabled"
									onclick="displayPage(1)"/>
								<input type="button" value="Columns" id="btnTab2" name="btnTab2" class="btn btndisabled" disabled="disabled"
									onclick="displayPage(2)"/>
								<input type="button" value="Output" id="btnTab3" name="btnTab3" class="btn btndisabled" disabled="disabled"
									onclick="displayPage(3)"/>
							</td>
							<td width="10"></td>
						</tr>

						<tr height="10"><td colspan="3"></td></tr>

						<tr>
							<td width="10"></td>
							<td>
								<!-- First tab -->
								<div id="div1">
									<table width="100%" height="100%" class="outline" cellspacing="0" cellpadding="5">
										<tr valign="top">
											<td>
												<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
													<tr height="10">
														<td width="5">&nbsp;</td>
														<td width="10">Name :</td>
														<td width="5">&nbsp;</td>
														<td>
															<input id="txtName" name="txtName" maxlength="50" style="WIDTH: 100%" class="text"
																onkeyup="changeTab1Control()">
														</td>
														<td width="20">&nbsp;</td>
														<td width="10">Owner :</td>
														<td width="5">&nbsp;</td>
														<td width="40%">
															<input id="txtOwner" name="txtOwner" class="text textdisabled" style="WIDTH: 100%" disabled="disabled">
														</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr>
														<td colspan="9" height="5"></td>
													</tr>

													<tr height="60">
														<td width="5">&nbsp;</td>
														<td width="10" nowrap valign="top">Description :</td>
														<td width="5">&nbsp;</td>
														<td width="40%" rowspan="3">
															<textarea id="txtDescription" name="txtDescription" class="textarea" style="HEIGHT: 99%; WIDTH: 100%" wrap="VIRTUAL" height="0" maxlength="255"
																onkeyup="changeTab1Control()"
																onpaste="var selectedLength = document.selection.createRange().text.length;var pasteData = window.clipboardData.getData('Text');if ((this.value.length + pasteData.length - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}"
																onkeypress="var selectedLength = document.selection.createRange().text.length;if ((this.value.length + 1 - selectedLength) > parseInt(this.maxlength)) {return(false);}else {return(true);}">
															</textarea>
														</td>
														<td width="20" nowrap>&nbsp;</td>
														<td width="10" valign="top">Access :</td>
														<td width="5">&nbsp;</td>
														<td width="40%" rowspan="3" valign="top">
															<%Html.RenderPartial("Util_Def_CustomReports/grdaccess")%>    
														</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr height="10">
														<td colspan="7">&nbsp;</td>
													</tr>

													<tr height="10">
														<td colspan="7">&nbsp;</td>
													</tr>

													<tr>
														<td colspan="9">
															<hr>
														</td>
													</tr>

													<tr height="10">
														<td width="5">&nbsp;</td>
														<td width="100" nowrap valign="top">Base Table :</td>
														<td width="5">&nbsp;</td>
														<td width="40%" valign="top">
															<select id="cboBaseTable" name="cboBaseTable" style="WIDTH: 100%" class="combo combodisabled"
																onchange="changeBaseTable()" disabled="disabled">
															</select>
														</td>
														<td width="20" nowrap>&nbsp;</td>
														<td width="10" valign="top">Records :</td>
														<td width="5">&nbsp;</td>
														<td width="40%">
															<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
																<tr>
																	<td width="5">
																		<input checked id="optRecordSelection1" name="optRecordSelection" type="radio"
																			onclick="changeBaseTableRecordOptions()"/>
																	</td>
																	<td width="5">&nbsp;</td>
																	<td width="30">
																		<label
																			tabindex="-1"
																			for="optRecordSelection1"
																			class="radio">
																		All
																		</label>
																	</td>
																	<td>&nbsp;</td>
																</tr>
																<tr>
																	<td colspan="6" height="5"></td>
																</tr>
																<tr>
																	<td width="5">
																		<input id="optRecordSelection2" name="optRecordSelection" type="radio"
																			onclick="changeBaseTableRecordOptions()"/>
																	</td>
																	<td width="5">&nbsp;</td>
																	<td width="20">
																		<label
																			tabindex="-1"
																			for="optRecordSelection2"
																			class="radio">
																		Picklist
																		</label>
																	</td>
																	<td width="5">&nbsp;</td>
																	<td>
																		<input id="txtBasePicklist" name="txtBasePicklist" class="text textdisabled" disabled="disabled" style="WIDTH: 100%">
																	</td>
																	<td width="30">
																		<input id="cmdBasePicklist" name="cmdBasePicklist" style="WIDTH: 100%" type="button" disabled="disabled" value="..." class="btn btndisabled"
																			onclick="selectRecordOption('base', 'picklist')" />
																	</td>
																</tr>
																<tr>
																	<td colspan="6" height="5"></td>
																</tr>
																<tr>
																	<td width="5">
																		<input id="optRecordSelection3" name="optRecordSelection" type="radio"
																			onclick="changeBaseTableRecordOptions()" />
																	</td>
																	<td width="5">&nbsp;</td>
																	<td width="20">
																		<label
																			tabindex="-1"
																			for="optRecordSelection3"
																			class="radio">
																		Filter
																		</label>
																	</td>
																	<td width="5">&nbsp;</td>
																	<td>
																		<input id="txtBaseFilter" name="txtBaseFilter" disabled="disabled" class="text textdisabled" style="WIDTH: 100%">
																	</td>
																	<td width="30">
																		<input id="cmdBaseFilter" name="cmdBaseFilter" style="WIDTH: 100%" type="button" class="btn btndisabled" disabled="disabled" value="..."
																			onclick="selectRecordOption('base', 'filter')"/>
																	</td>
																</tr>
															</table>
														</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr>
														<td colspan="9" height="5">&nbsp;</td>
													</tr>

													<tr>
														<td colspan="5">&nbsp;</td>
														<td colspan="3">
															<input name="chkPrintFilter" id="chkPrintFilter" type="checkbox" disabled="disabled" tabindex="-1"
																onclick="changeTab1Control()" />
															<label
																for="chkPrintFilter"
																class="checkbox checkboxdisabled"
																tabindex="0">
																Display filter or picklist title in the report header
															</label>
														</td>
														<td width="5">&nbsp;</td>
													</tr>

													<tr>
														<td colspan="9" height="5">&nbsp;</td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
								</div>
								<div id="div2" style="visibility: hidden; display: none">
									<table width="100%" height="100%" class="outline" cellspacing="0" cellpadding="5">

										<tr valign="top">
											<td>
												<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
													<tr><td colspan="9" height="5"></td></tr>
													<tr height="10">
														<td width="5">&nbsp;</td>
														<td colspan="4" valign="top"><u>Headings & Breaks</u></td>
														<td width="15%" align="Center">Start</td>
														<td width="5">&nbsp;</td>
														<td width="15%" align="Center">Stop</td>
														<td width="5">&nbsp;</td>
														<td width="15%" align="Center">Increment</td>
														<td>&nbsp;</td>
													</tr>
													<tr><td colspan="9" height="5"></td></tr>
													<tr height="23">
														<td width="5">&nbsp;</td>
														<td width="80" nowrap valign="top">Horizontal :</td>
														<td width="5">&nbsp;</td>
														<td width="40%" valign="top">
															<select id="cboHor" name="cboHor" style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
																onchange="cboHor_Change();changeTab2Control(); ">
															</select>
														</td>
														<td width="15">&nbsp;</td>
														<td>
															<object classid="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830" 
																codebase="cabs/tinumb6.cab#version=6,0,1,1" 
																id="txtHorStart" name="txtHorStart" 
																style="height:24px; WIDTH:195px"
																onkeyup="changeTab2Control()">
																<param name="DisplayFormat" value="##########0.0000">
																<param name="Format" value="##########0.0000">
																<param name="MaxValue" value="99999999999.9999">
																<param name="Value" value="<%=lngHStart%>">
																<param name="Appearance" value="0">
																<param name="BackColor" value="15988214">
																<param name="ForeColor" value="6697779">
															</object>
														</td>
														<td width="5">&nbsp;</td>
														<td width="15%">
															<object classid="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830" 
																codebase="cabs/tinumb6.cab#version=6,0,1,1" 
																id="txtHorStop"
																name="txtHorStop" 
																style="height:24px; WIDTH:195px"
																onkeyup="changeTab2Control()">
																<param name="DisplayFormat" value="##########0.0000">
																<param name="Format" value="##########0.0000">
																<param name="MaxValue" value="99999999999.9999">
																<param name="Value" value="<%=lngHStop%>">
																<param name="Appearance" value="0">
																<param name="BackColor" value="15988214">
																<param name="ForeColor" value="6697779">
															</object>
														</td>
														<td width="5">&nbsp;</td>
														<td width="15%">
															<object classid="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"
																codebase="cabs/tinumb6.cab#version=6,0,1,1" 
																id="txtHorStep" 
																name="txtHorStep" 
																style="height:24px; WIDTH:195px"
																onkeyup="changeTab2Control()">
																<param name="DisplayFormat" value="##########0.0000">
																<param name="Format" value="##########0.0000">
																<param name="MaxValue" value="99999999999.9999">
																<param name="Value" value="<%=lngHStep%>">
																<param name="Appearance" value="0">
																<param name="BackColor" value="15988214">
																<param name="ForeColor" value="6697779">
															</object>
														</td>

														<td>&nbsp;</td>
													</tr>
													<tr><td colspan="9" height="5"></td></tr>
													<tr height="23">
														<td width="5">&nbsp;</td>
														<td width="80" nowrap valign="top">Vertical :</td>
														<td width="5">&nbsp;</td>
														<td width="40%" valign="top">
															<select id="cboVer" name="cboVer" style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
																onchange="cboVer_Change();changeTab2Control(); ">
															</select>
														</td>
														<td width="15">&nbsp;</td>
														<td width="15%">
															<object classid="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"
																codebase="cabs/tinumb6.cab#version=6,0,1,1" 
																id="txtVerStart" 
																name="txtVerStart" 
																style="height:24px; WIDTH:195px"
																onkeyup="changeTab2Control()">
																<param name="DisplayFormat" value="##########0.0000">
																<param name="Format" value="##########0.0000">
																<param name="MaxValue" value="99999999999.9999">
																<param name="Value" value="<%=lngVStart%>">
																<param name="Appearance" value="0">
																<param name="BackColor" value="15988214">
																<param name="ForeColor" value="6697779">
															</object>
														</td>
														<td width="5">&nbsp;</td>
														<td width="15%">
															<object classid="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"
																codebase="cabs/tinumb6.cab#version=6,0,1,1"
																id="txtVerStop"
																name="txtVerStop"
																style="height:24px; WIDTH:195px"
																onkeyup="changeTab2Control()">
																<param name="DisplayFormat" value="##########0.0000">
																<param name="Format" value="##########0.0000">
																<param name="MaxValue" value="99999999999.9999">
																<param name="Value" value="<%=lngVStop%>">
																<param name="Appearance" value="0">
																<param name="BackColor" value="15988214">
																<param name="ForeColor" value="6697779">
															</object>
														</td>
														<td width="5">&nbsp;</td>
														<td width="15%">
															<object classid="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830" 
																codebase="cabs/tinumb6.cab#version=6,0,1,1" 
																id="txtVerStep" 
																name="txtVerStep" 
																style="height:24px; WIDTH:195px"
																onkeyup="changeTab2Control()">
																<param name="DisplayFormat" value="##########0.0000">
																<param name="Format" value="##########0.0000">
																<param name="MaxValue" value="99999999999.9999">
																<param name="Value" value="<%=lngVStep%>">
																<param name="Appearance" value="0">
																<param name="BackColor" value="15988214">
																<param name="ForeColor" value="6697779">
															</object>
														</td>

														<td>&nbsp;</td>
													</tr>
													<tr><td colspan="9" height="5"></td></tr>
													<tr height="23">
														<td width="5">&nbsp;</td>
														<td width="100" nowrap valign="top">Page Break :</td>
														<td width="5">&nbsp;</td>
														<td width="40%" valign="top">
															<select id="cboPgb" name="cboPgb" style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
																onchange="cboPgb_Change();changeTab2Control(); ">
															</select>
														</td>
														<td width="15">&nbsp;</td>
														<td width="15%">
															<object classid="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"
																codebase="cabs/tinumb6.cab#version=6,0,1,1" id="txtPgbStart" name="txtPgbStart"
																style="LEFT: 0px; TOP: 0px; height:24px; WIDTH:195px"
																onkeyup="changeTab2Control()">
																<param name="DisplayFormat" value="##########0.0000">
																<param name="Format" value="##########0.0000">
																<param name="MaxValue" value="99999999999.9999">
																<param name="Value" value="<%=lngPStart%>">
																<param name="Appearance" value="0">
																<param name="BackColor" value="15988214">
																<param name="ForeColor" value="6697779">
															</object>
														</td>
														<td width="5">&nbsp;</td>
														<td width="15%">
															<object classid="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830"
																codebase="cabs/tinumb6.cab#version=6,0,1,1" id="txtPgbStop"
																name="txtPgbStop" style="LEFT: 0px; TOP: 0px; height:24px; WIDTH:195px"
																onkeyup="changeTab2Control()">
																<param name="DisplayFormat" value="##########0.0000">
																<param name="Format" value="##########0.0000">
																<param name="MaxValue" value="99999999999.9999">
																<param name="Value" value="<%=lngPStop%>">
																<param name="Appearance" value="0">
																<param name="BackColor" value="15988214">
																<param name="ForeColor" value="6697779">
															</object>
														</td>
														<td width="5">&nbsp;</td>
														<td width="15%">
															<object classid="clsid:49CBFCC2-1337-11D2-9BBF-00A024695830" 
																codebase="cabs/tinumb6.cab#version=6,0,1,1" 
																id="txtPgbStep" name="txtPgbStep" 
																style="LEFT: 0px; TOP: 0px; height:24px; WIDTH:195px"
																onkeyup="changeTab2Control()">
																<param name="DisplayFormat" value="##########0.0000">
																<param name="Format" value="##########0.0000">
																<param name="MaxValue" value="99999999999.9999">
																<param name="Value" value="<%=lngPStep%>">
																<param name="Appearance" value="0">
																<param name="BackColor" value="15988214">
																<param name="ForeColor" value="6697779">
															</object>
														</td>
														<td>&nbsp;</td>
													</tr>

													<tr height="40">
														<td colspan="11">
															<hr>
														</td>
													</tr>

													<tr height="10">
														<td width="5">&nbsp;</td>
														<td width="80" colspan="4" nowrap valign="top"><u>Intersection</u></td>
													</tr>
													<tr><td colspan="9" height="5"></td></tr>												

												</table>

												<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
													
													<tr>

												<td colspan="2">
													<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
														<tr height="10">
															<td width="5">&nbsp;</td>
															<td width="80" nowrap valign="top">Column :</td>
															<td width="5">&nbsp;</td>
															<td width="100%" valign="top">
																<select id="cboInt" name="cboInt" style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
																	onchange="cboInt_Change();changeTab2Control(); ">
																</select>
															</td>
														</tr>
														<tr><td colspan="9" height="5"></td></tr>
														<tr height="10">
															<td width="5">&nbsp;</td>
															<td width="80" nowrap valign="top">Type :</td>
															<td width="5">&nbsp;</td>
															<td width="100%" valign="top">
																<select id="cboIntType" name="cboIntType" 
																	style="WIDTH: 100%" class="combo combodisabled" disabled="disabled"
																	onchange="changeTab2Control()">
																	<option value="1">Average</option>
																	<option value="0" selected>Count</option>
																	<option value="2">Maximum</option>
																	<option value="3">Minimum</option>
																	<option value="4">Total</option>
																</select>
															</td>
														</tr>
													</table>
												</td>

											<td width="15">&nbsp;</td>
											<td>
												<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
													<tr><td>
															<input type="checkbox" id="chkPercentage" name="chkPercentage" tabindex="-1"
																onclick="changeTab2Control()" />
															<label
																for="chkPercentage"
																class="checkbox"
																tabindex="0">
																Percentage of Type
															</label>
														</td></tr>
													<tr><td height="5"></td></tr>
													<tr>
														<td>
															<input type="checkbox" id="chkPerPage" name="chkPerPage" tabindex="-1"
																onclick="changeTab2Control()" />
															<label
																for="chkPerPage"
																class="checkbox"
																tabindex="0">
																Percentage of Page
															</label>
														</td>
													</tr>
													<tr><td height="5"></td></tr>
													<tr>
														<td>
															<input type="checkbox" id="chkSuppress" name="chkSuppress" tabindex="-1"
																onclick="changeTab2Control()" />
															<label
																for="chkSuppress"
																class="checkbox"
																tabindex="0">
																Suppress Zeros
															</label>
														</td>
													</tr>
													<tr><td height="5"></td></tr>
													<tr>
														<td>
															<input type="checkbox" id="chkUse1000" name="chkUse1000" tabindex="-1"
																onclick="changeTab2Control()" />
															<label
																for="chkUse1000"
																class="checkbox"
																tabindex="0">
																Use 1000 Separators (,)
															</label>
														</td>
													</tr>
												</table>
											</td>
											
											</tr>

											<tr><td colspan="9" height="5"></td></tr>

										</table>

									</table>
								</div>							

<!-- OUTPUT OPTIONS -->
<div id="div3" style="visibility: hidden; display: none">
	<table width="100%" height="100%" class="outline" cellspacing="0" cellpadding="5">
		<tr valign="top">
			<td>
				<table width="100%" class="invisible" cellspacing="10" cellpadding="0">
					<tr>
						<td valign="top" rowspan="2" width="25%" height="100%">
							<table class="outline" cellspacing="0" cellpadding="4" width="100%" height="100%">
								<tr height="10">
									<td height="10" align="left" valign="top">Output Format :
										<br>
										<br>
										<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
											<tr height="20">
												<td width="5">&nbsp</td>
												<td align="left" width="15">
													<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat0" value="0"
														onclick="formatClick(0);"/>
												</td>
												<td align="left" nowrap>
													<label
														tabindex="-1"
														for="optOutputFormat0"
														class="radio">
														Data Only
													</label>
												</td>
												<td width="5">&nbsp</td>
											</tr>
											<tr height="10"><td colspan="4"></td></tr>
											<tr height="20">
												<td width="5">&nbsp</td>
												<td align="left" width="15">
													<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat1" value="1"
														onclick="formatClick(1);" />
												</td>
												<td align="left" nowrap>
													<label
														tabindex="-1"
														for="optOutputFormat1"
														class="radio">
														CSV File
													</label>
												</td>
												<td width="5">&nbsp</td>
											</tr>
											<tr height="10"><td colspan="4"></td></tr>
											<tr height="20">
												<td width="5">&nbsp</td>
												<td align="left" width="15">
													<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat2" value="2"
														onclick="formatClick(2);"/>
												</td>
												<td align="left" nowrap>
													<label
														tabindex="-1"
														for="optOutputFormat2"
														class="radio">
														HTML Document
													</label>
												</td>
												<td width="5">&nbsp</td>
											</tr>
											<tr height="10"><td colspan="4"></td></tr>
											<tr height="20">
												<td width="5">&nbsp</td>
												<td align="left" width="15">
													<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat3" value="3"
														onclick="formatClick(3);" />
												</td>
												<td align="left" nowrap>
													<label
														tabindex="-1"
														for="optOutputFormat3"
														class="radio">
														Word Document
													</label>
												</td>
												<td width="5">&nbsp</td>
											</tr>
											<tr height="10"><td colspan="4"></td></tr>
											<tr height="20">
												<td width="5">&nbsp</td>
												<td align="left" width="15">
													<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat4" value="4"
														onclick="formatClick(4);" />
												</td>
												<td align="left" nowrap>
													<label
														tabindex="-1"
														for="optOutputFormat4"
														class="radio">
														Excel Worksheet
													</label>
												</td>
												<td width="5">&nbsp</td>
											</tr>
											<tr height="10"><td colspan="4"></td></tr>
											<tr height="5">
												<td width="5">&nbsp</td>
												<td align="left" width="15">
													<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat5" value="5"
														onclick="formatClick(5);" />
												</td>
												<td>
													<label
														tabindex="-1"
														for="optOutputFormat5"
														class="radio">
														Excel Chart
													</label>
												</td>
												<td width="5">&nbsp</td>
											</tr>
											<tr height="10"><td colspan="4"></td></tr>
											<tr height="5">
												<td width="5">&nbsp</td>
												<td align="left" width="15">
													<input type="radio" width="20" style="WIDTH: 20px" name="optOutputFormat" id="optOutputFormat6" value="6"
														onclick="formatClick(6);" />
												</td>
												<td nowrap>
													<label
														tabindex="-1"
														for="optOutputFormat6"
														class="radio">
														Excel Pivot Table
													</label>
												</td>
												<td width="5">&nbsp</td>
											</tr>
											<tr height="5"><td colspan="4"></td></tr>
										</table>
									</td>
								</tr>
							</table>
						</td>
						<td valign="top" width="75%">
							<table class="outline" cellspacing="0" cellpadding="4" width="100%" height="100%">
								<tr height="10">
									<td height="10" align="left" valign="top">Output Destination(s) :
										<br>
										<br>
										<table class="invisible" cellspacing="0" cellpadding="0" width="100%">
											<tr height="20">
												<td width="5">&nbsp</td>
												<td align="left" colspan="6" nowrap>
													<input name="chkPreview" id="chkPreview" type="checkbox" disabled="disabled" tabindex="-1"
														onclick="changeTab3Control();" />
													<label
														for="chkPreview"
														class="checkbox checkboxdisabled"
														tabindex="0">
														Preview on screen
													</label>
												</td>
												<td width="5">&nbsp</td>
											</tr>
											<tr height="10"><td colspan="8"></td></tr>
											<tr height="20">
												<td></td>
												<td align="left" colspan="6" nowrap>
													<input name="chkDestination0" id="chkDestination0" type="checkbox" disabled="disabled" tabindex="-1"
														onclick="changeTab3Control();" />
													<label
														for="chkDestination0"
														class="checkbox checkboxdisabled"
														tabindex="0">
														Display output on screen 
													</label>
												</td>
												<td></td>
											</tr>
											<tr height="10"><td colspan="8"></td></tr>
											<tr height="20">
												<td></td>
												<td align="left" nowrap>
													<input name="chkDestination1" id="chkDestination1" type="checkbox" disabled="disabled" tabindex="-1"
														onclick="changeTab3Control();" />
													<label
														for="chkDestination1"
														class="checkbox checkboxdisabled"
														tabindex="0">
														Send to printer 
													</label>
												</td>
												<td width="30" nowrap>&nbsp</td>
												<td align="left" nowrap>Printer location : 
												</td>
												<td width="15">&nbsp</td>
												<td colspan="2">
													<select id="cboPrinterName" name="cboPrinterName" width="100%" style="WIDTH: 400px" class="combo"
														onchange="changeTab3Control()">
													</select>
												</td>
												<td></td>
											</tr>
											<tr height="10"><td colspan="8"></td></tr>
											<tr height="20">
												<td></td>
												<td align="left" nowrap>
													<input name="chkDestination2" id="chkDestination2" type="checkbox" disabled="disabled" tabindex="-1"
														onclick="changeTab3Control();" />
													<label
														for="chkDestination2"
														class="checkbox checkboxdisabled"
														tabindex="0">
														Save to file
													</label>
												</td>
												<td></td>
												<td align="left" nowrap>File name :   
												</td>
												<td></td>
												<td colspan="2">
													<table class="invisible" cellspacing="0" cellpadding="0" style="WIDTH: 400px">
														<tr>
															<td>
																<input id="txtFilename" name="txtFilename" class="text textdisabled" disabled="disabled" style="width: 375px">
															</td>
															<td width="25">
																<input id="cmdFilename" name="cmdFilename" style="WIDTH: 100%" type="button" class="btn" value="..."
																	onclick="saveFile(); changeTab3Control();"/>
															</td>

</tr>
</table>
															
												</td>

								</tr>
								<tr height="10"><td colspan="8"></td></tr>
								<tr height="20">
									<td colspan="3"></td>
									<td align="left" nowrap>If existing file :</td>
									<td></td>
									<td colspan="2" width="100%" nowrap>
										<select id="cboSaveExisting" name="cboSaveExisting" width="100%" style="WIDTH: 400px" class="combo"
											onchange="changeTab3Control()">
										</select>
									</td>
									<td></td>
								</tr>
								<tr height="10"><td colspan="8"></td></tr>
								<tr height="20">
									<td></td>
									<td align="left" nowrap>
										<input name="chkDestination3" id="chkDestination3" type="checkbox" disabled="disabled" tabindex="-1"
											onclick="changeTab3Control();"/>
										<label
											for="chkDestination3"
											class="checkbox checkboxdisabled"
											tabindex="0">
											Send as email 
										</label>
									</td>
									<td></td>
									<td align="left" nowrap>Email group : </td>
									<td></td>
									<td colspan="2">
										<table class="invisible" cellspacing="0" cellpadding="0" style="WIDTH: 400px">
											<tr>
												<td>
													<input id="txtEmailGroup" name="txtEmailGroup" class="text textdisabled" disabled="disabled" style="WIDTH: 100%">
													<input id="txtEmailGroupID" name="txtEmailGroupID" type="hidden">
												</td>
												<td width="25">
													<input id="cmdEmailGroup" name="cmdEmailGroup" style="WIDTH: 100%" type="button" class="btn" value="..."
														onclick="selectEmailGroup(); changeTab3Control();" />
												</td>
												
											</tr>	
										</table>

									</td>
					</tr>

					<tr height="10"><td colspan="8"></td></tr>

					<tr height="20">
						<td colspan="3"></td>
						<td align="left" nowrap>Email subject :</td>
						<td></td>
						<td colspan="2" width="100%" nowrap>
							<input id="txtEmailSubject" disabled="disabled" class="text textdisabled" maxlength="255" name="txtEmailSubject" style="WIDTH: 400px"
								onchange="frmUseful.txtChanged.value = 1;"
								onkeydown="frmUseful.txtChanged.value = 1;">
						</td>
						<td></td>
					</tr>

					<tr height="10"><td colspan="8"></td></tr>

					<tr height="20">
						<td colspan="3"></td>
						<td align="left" nowrap>Attach as : </td>
						<td></td>
						<td colspan="2" width="100%" nowrap>
							<input id="txtEmailAttachAs" disabled="disabled" maxlength="255" class="text textdisabled" name="txtEmailAttachAs" style="WIDTH: 400px"
								onchange="frmUseful.txtChanged.value = 1;"
								onkeydown="frmUseful.txtChanged.value = 1;">
						</td>
						<td></td>
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
										</table>
</div>					

							</td>
						</tr>
					</table>


				</td>
			</tr>

			<tr height="5"><td colspan="3"></td></tr>
		
			<tr height="10">
				<td width="10"></td>
				<td>
					<table style="width: 100%" class="invisible" cellspacing="0" cellpadding="0">
						<tr>
							<td>&nbsp;</td>
							<td width="80">
								<input type="button" id="cmdOK" name="cmdOK" value="OK" style="width: 100%" class="btn"
									onclick="okClick()" />
							</td>
							<td width="10"></td>
							<td width="80">
								<input type="button" id="cmdCancel" name="cmdCancel" value="Cancel" style="width: 100%" class="btn"
									onclick="cancelClick()" />
							</td>
						</tr>
					</table>
				</td>
				<td width="10"></td>
			</tr>

		</table>
	
		<input type='hidden' id="txtBasePicklistID" name="txtBasePicklistID">
		<input type='hidden' id="txtBaseFilterID" name="txtBaseFilterID">
		<input type='hidden' id="txtDatabase" name="txtDatabase" value="<%=session("Database")%>">

		<input type='hidden' id="txtWordVer" name="txtWordVer" value="<%=Session("WordVer")%>">
		<input type='hidden' id="txtExcelVer" name="txtExcelVer" value="<%=Session("ExcelVer")%>">
		<input type='hidden' id="txtWordFormats" name="txtWordFormats" value="<%=Session("WordFormats")%>">
		<input type='hidden' id="txtExcelFormats" name="txtExcelFormats" value="<%=Session("ExcelFormats")%>">
		<input type='hidden' id="txtWordFormatDefaultIndex" name="txtWordFormatDefaultIndex" value="<%=Session("WordFormatDefaultIndex")%>">
		<input type='hidden' id="txtExcelFormatDefaultIndex" name="txtExcelFormatDefaultIndex" value="<%=Session("ExcelFormatDefaultIndex")%>">
</form>

<form id="frmAccess">
	<%
		sErrorDescription = ""
	
		' Get the table records.
		Dim cmdAccess As Command = New Command()
		cmdAccess.CommandText = "spASRIntGetUtilityAccessRecords"
		cmdAccess.CommandType = CommandTypeEnum.adCmdStoredProc
		cmdAccess.ActiveConnection = Session("databaseConnection")

		Dim prmUtilType = cmdAccess.CreateParameter("utilType", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput)
		cmdAccess.Parameters.Append(prmUtilType)
		prmUtilType.Value = 1	' 1 = cross tabs

		Dim prmUtilID = cmdAccess.CreateParameter("utilID", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput)
		cmdAccess.Parameters.Append(prmUtilID)
		If UCase(Session("action")) = "NEW" Then
			prmUtilID.Value = 0
		Else
			prmUtilID.Value = CleanNumeric(Session("utilid"))
		End If

		Dim prmFromCopy = cmdAccess.CreateParameter("fromCopy", DataTypeEnum.adInteger, ParameterDirectionEnum.adParamInput)
		cmdAccess.Parameters.Append(prmFromCopy)
		If UCase(Session("action")) = "COPY" Then
			prmFromCopy.Value = 1
		Else
			prmFromCopy.Value = 0
		End If

		Err.Clear()
		Dim rstAccessInfo = cmdAccess.Execute
		If (Err.Number <> 0) Then
			sErrorDescription = "The access information could not be retrieved." & vbCrLf & FormatError(Err.Description)
		End If

		If Len(sErrorDescription) = 0 Then
			Dim iCount = 0
			Do While Not rstAccessInfo.EOF
				Response.Write("<input type='hidden' id=txtAccess_" & iCount & " name=txtAccess_" & iCount & " value=""" & rstAccessInfo.Fields("accessDefinition").Value & """>" & vbCrLf)

				iCount = iCount + 1
				rstAccessInfo.MoveNext()
			Loop

			' Release the ADO recordset object.
			rstAccessInfo.Close()
			rstAccessInfo = Nothing
		End If
	
		' Release the ADO command object.
		cmdAccess = Nothing
	%>
</form>

<form id="frmUseful" name="frmUseful" style="visibility: hidden; display: none">
	<input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
	<input type="hidden" id="txtLoading" name="txtLoading" value="Y">
	<input type="hidden" id="txtCurrentBaseTableID" name="txtCurrentBaseTableID">
	<input type="hidden" id="txtCurrentHColID" name="txtCurrentHColID">
	<input type="hidden" id="txtCurrentVColID" name="txtCurrentVColID">
	<input type="hidden" id="txtCurrentPColID" name="txtCurrentPColID">
	<input type="hidden" id="txtCurrentIColID" name="txtCurrentIColID">
	<input type="hidden" id="txtTablesChanged" name="txtTablesChanged">
	<input type="hidden" id="txtSelectedColumnsLoaded" name="txtSelectedColumnsLoaded" value="0">
	<input type="hidden" id="txtSortLoaded" name="txtSortLoaded" value="0">
	<input type="hidden" id="txtSecondTabShown" name="txtSecondTabShown" value="0">
	<input type="hidden" id="txtRepetitionLoaded" name="txtRepetitionLoaded" value="0">
	<input type="hidden" id="txtChanged" name="txtChanged" value="0">
	<input type="hidden" id="txtUtilID" name="txtUtilID" value='<%=session("utilid")%>'>
	<%
		Dim cmdDefinition As Command = New Command()
		cmdDefinition.CommandText = "sp_ASRIntGetModuleParameter"
		cmdDefinition.CommandType = CommandTypeEnum.adCmdStoredProc
		cmdDefinition.ActiveConnection = Session("databaseConnection")

		Dim prmModuleKey = cmdDefinition.CreateParameter("moduleKey", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
		cmdDefinition.Parameters.Append(prmModuleKey)
		prmModuleKey.value = "MODULE_PERSONNEL"

		Dim prmParameterKey = cmdDefinition.CreateParameter("paramKey", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
		cmdDefinition.Parameters.Append(prmParameterKey)
		prmParameterKey.value = "Param_TablePersonnel"

		Dim prmParameterValue = cmdDefinition.CreateParameter("paramValue", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
		cmdDefinition.Parameters.Append(prmParameterValue)

		Err.Clear()
		cmdDefinition.Execute()

		Response.Write("<input type='hidden' id=txtPersonnelTableID name=txtPersonnelTableID value=" & cmdDefinition.Parameters("paramValue").value & ">" & vbCrLf)
	
		cmdDefinition = Nothing

		Response.Write("<input type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
		Response.Write("<input type='hidden' id=txtAction name=txtAction value=" & Session("action") & ">" & vbCrLf)
	%>
</form>

<form id="frmValidate" name="frmValidate" target="validate" method="post" action="util_validate_crosstab" style="visibility: hidden; display: none">
	<input type="hidden" id="validateBaseFilter" name="validateBaseFilter" value="0">
	<input type="hidden" id="validateBasePicklist" name="validateBasePicklist" value="0">
	<input type="hidden" id="validateEmailGroup" name="validateEmailGroup" value="0">
	<input type="hidden" id="validateCalcs" name="validateCalcs" value=''>
	<input type="hidden" id="validateHiddenGroups" name="validateHiddenGroups" value=''>
	<input type="hidden" id="validateName" name="validateName" value=''>
	<input type="hidden" id="validateTimestamp" name="validateTimestamp" value=''>
	<input type="hidden" id="validateUtilID" name="validateUtilID" value=''>
</form>

<form id="frmSend" name="frmSend" method="post" action="util_def_crosstabs_submit" style="visibility: hidden; display: none">
	<input type="hidden" id="txtSend_ID" name="txtSend_ID" value="0">
	<input type="hidden" id="txtSend_name" name="txtSend_name" value=''>
	<input type="hidden" id="txtSend_description" name="txtSend_description" value=''>
	<input type="hidden" id="txtSend_baseTable" name="txtSend_baseTable" value="0">
	<input type="hidden" id="txtSend_allRecords" name="txtSend_allRecords" value="0">
	<input type="hidden" id="txtSend_picklist" name="txtSend_picklist" value="0">
	<input type="hidden" id="txtSend_filter" name="txtSend_filter" value="0">
	<input type="hidden" id="txtSend_PrintFilter" name="txtSend_PrintFilter" value="0">
	<input type="hidden" id="txtSend_access" name="txtSend_access" value=''>
	<input type="hidden" id="txtSend_userName" name="txtSend_userName" value=''>

	<input type="hidden" id="txtSend_HColID" name="txtSend_HColID" value="0">
	<input type="hidden" id="txtSend_HStart" name="txtSend_HStart" value=''>
	<input type="hidden" id="txtSend_HStop" name="txtSend_HStop" value=''>
	<input type="hidden" id="txtSend_HStep" name="txtSend_HStep" value=''>
	<input type="hidden" id="txtSend_VColID" name="txtSend_VColID" value="0">
	<input type="hidden" id="txtSend_VStart" name="txtSend_VStart" value=''>
	<input type="hidden" id="txtSend_VStop" name="txtSend_VStop" value=''>
	<input type="hidden" id="txtSend_VStep" name="txtSend_VStep" value=''>
	<input type="hidden" id="txtSend_PColID" name="txtSend_PColID" value="0">
	<input type="hidden" id="txtSend_PStart" name="txtSend_PStart" value=''>
	<input type="hidden" id="txtSend_PStop" name="txtSend_PStop" value=''>
	<input type="hidden" id="txtSend_PStep" name="txtSend_PStep" value=''>
	<input type="hidden" id="txtSend_IType" name="txtSend_IType" value="0">
	<input type="hidden" id="txtSend_IColID" name="txtSend_IColID" value="0">
	<input type="hidden" id="txtSend_Percentage" name="txtSend_Percentage" value="0">
	<input type="hidden" id="txtSend_PerPage" name="txtSend_PerPage" value="0">
	<input type="hidden" id="txtSend_Suppress" name="txtSend_Suppress" value="0">
	<input type="hidden" id="txtSend_Use1000Separator" name="txtSend_Use1000Separator" value="0">

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

	<input type="hidden" id="txtSend_reaction" name="txtSend_reaction">

	<input type="hidden" id="txtSend_jobsToHide" name="txtSend_jobsToHide">
	<input type="hidden" id="txtSend_jobsToHideGroups" name="txtSend_jobsToHideGroups">
</form>

<form id="frmRecordSelection" name="frmRecordSelection" target="recordSelection" action="util_recordSelection" method="post" style="visibility: hidden; display: none">
	<input type="hidden" id="recSelType" name="recSelType">
	<input type="hidden" id="recSelTableID" name="recSelTableID">
	<input type="hidden" id="recSelCurrentID" name="recSelCurrentID">
	<input type="hidden" id="recSelTable" name="recSelTable">
	<input type="hidden" id="recSelDefOwner" name="recSelDefOwner">
	<input type="hidden" id="recSelDefType" name="recSelDefType">
</form>

<form id="frmEmailSelection" name="frmEmailSelection" target="emailSelection" action="util_emailSelection" method="post" style="visibility: hidden; display: none">
	<input type="hidden" id="EmailSelCurrentID" name="EmailSelCurrentID">
</form>

<form id="frmSelectionAccess" name="frmSelectionAccess" style="visibility: hidden; display: none">
	<input type="hidden" id="forcedHidden" name="forcedHidden" value="N">
	<input type="hidden" id="baseHidden" name="baseHidden" value="N">

	<!-- need the count of hidden child filter access info -->
	<input type="hidden" id="childHidden" name="childHidden" value="N">
	<input type="hidden" id="calcsHiddenCount" name="calcsHiddenCount" value="0">
</form>

<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
	<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</form>

<input type='hidden' id="txtTicker" name="txtTicker" value="0">
<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">

<script type="text/javascript">
	util_def_crosstabs_window_onload();
	util_def_crosstabs_addhandlers();
</script>
