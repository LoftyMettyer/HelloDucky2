﻿<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<% 
	Dim bStandardReportPrompt As Boolean
		
	' This page is called from DefSel.asp.  If receives the following
	' information via the request object:
	'
	' ConfirmType - ok/yesno
	' UtilType    - 0-13 (see UtilityType code in DATMGR .exe
	' UtilName    - <the name of the utility>
	' UtilID      - <the id of the utility>
	' Action      - run/delete
	' FollowPage  - page to go to if YES is clicked <util_run.asp>
	If Session("action") = "STDREPORT_DATEPROMPT" Or Session("optionaction") = "STDREPORT_DATEPROMPT" Then
		bStandardReportPrompt = True
		Session("utilname") = ""
	Else
		bStandardReportPrompt = False
		Session("utiltype") = Request.Form("utiltype")
		Session("utilid") = Request.Form("utilid")
		Session("utilname") = Request.Form("utilname")
		Session("action") = Request.Form("action")
	End If
%>


<script type="text/javascript">

	function promptedvalues_window_onload() {

		var frmPromptedValues = document.getElementById("frmPromptedValues");

		frmPromptedValues.txtLocaleDateFormat.value = OpenHR.LocaleDateFormat;
		frmPromptedValues.txtLocaleDecimalSeparator.value = OpenHR.LocaleDecimalSeparator;
		frmPromptedValues.txtLocaleThousandSeparator.value = OpenHR.LocaleThousandSeparator;

		if (frmPromptedValues.RunInOptionFrame.value == "True") {
			$("#optionframe").attr("data-framesource", "UTIL_RUN_PROMPTEDVALUES");
			$("#workframe").hide();
			$("#optionframe").show();
		} else {
			$("#reportframe").attr("data-framesource", "UTIL_RUN_PROMPTEDVALUES");
			if (frmPromptedValues.StandardReportPrompt.value == "True") {
				$("#reportframe").show();
			}
		}

		if (frmPromptedValues.txtPromptCount.value == 0) {

			if (frmPromptedValues.StandardReportPrompt.value == "True") {
				OpenHR.submitForm(frmPromptedValues);
			} else {
				if (menu_isSSIMode() == true) {
					OpenHR.submitForm(frmPromptedValues, "workframe", true);
				} else {
					OpenHR.showInReportFrame(frmPromptedValues, true);
				}

			}

		} else {

			if (menu_isSSIMode() == false) {
				$(".popup").dialog("open");
			}

		// Set focus on the first prompt control.
			var controlCollection = frmPromptedValues.elements;
			if (controlCollection != null) {
				for (var i = 0; i < controlCollection.length; i++) {
					var sControlName = controlCollection.item(i).name;
					var sControlPrefix = sControlName.substr(0, 7);

					if ((sControlPrefix == "prompt_") || (sControlName.substr(0, 13) == "promptLookup_")) {
						controlCollection.item(i).focus();
						break;
					}
				}

			}
		}
	}
</script>

<div class="pageTitleDiv">
	<a href='javascript:loadPartialView("linksMain", "Home", "workframe", null);' title='Home'><i class='pageTitleIcon icon-arrow-left'></i></a>
	<h3 class="pageTitle"><% =Session("utilname")%></h3>
</div>

<form name="frmPromptedValues" id="frmPromptedValues" method="POST" action='<%
	If bStandardReportPrompt Then
		Response.Write("stdrpt_def_Absence")
	Else
		Response.Write("util_run")
	End If
		%>'>

	<%
		' Get variables for Absence Breakdown / Bradford Factor
		Session("stdReport_StartDate") = Request.Form("txtFromDate")
		Session("stdReport_EndDate") = Request.Form("txtToDate")
		Session("stdReport_AbsenceTypes") = Request.Form("txtAbsenceTypes")
		Session("stdReport_FilterID") = Request.Form("txtBaseFilterID")
		Session("stdReport_PicklistID") = Request.Form("txtBasePicklistID")
		Session("stdReport_Bradford_SRV") = Request.Form("txtSRV")
		Session("stdReport_Bradford_ShowDurations") = Request.Form("txtShowDurations")
		Session("stdReport_Bradford_ShowFormula") = Request.Form("txtShowFormula")
		Session("stdReport_Bradford_ShowInstances") = Request.Form("txtShowInstances")
		Session("stdReport_Bradford_OmitBeforeStart") = Request.Form("txtOmitBeforeStart")
		Session("stdReport_Bradford_OmitAfterEnd") = Request.Form("txtOmitAfterEnd")
		Session("stdReport_Bradford_txtOrderBy1") = Request.Form("txtOrderBy1")
		Session("stdReport_Bradford_txtOrderBy1ID") = Request.Form("txtOrderBy1ID")
		Session("stdReport_Bradford_txtOrderBy1Asc") = Request.Form("txtOrderBy1Asc")
		Session("stdReport_Bradford_txtOrderBy2") = Request.Form("txtOrderBy2")
		Session("stdReport_Bradford_txtOrderBy2ID") = Request.Form("txtOrderBy2ID")
		Session("stdReport_Bradford_txtOrderBy2Asc") = Request.Form("txtOrderBy2Asc")
		Session("stdReport_PrintFilterPicklistHeader") = Request.Form("txtPrintFPinReportHeader")
		Session("stdReport_MinimumBradfordFactor") = Request.Form("txtMinimumBradfordFactor")
		Session("stdReport_MinimumBradfordFactorAmount") = Request.Form("txtMinimumBradfordFactorAmount")
		Session("stdReport_DisplayBradfordDetail") = Request.Form("txtDisplayBradfordDetail")

		Session("stdReport_OutputPreview") = Request.Form("txtSend_OutputPreview")
		Session("stdReport_OutputFormat") = Request.Form("txtSend_OutputFormat")
		Session("stdReport_OutputScreen") = Request.Form("txtSend_OutputScreen")
		Session("stdReport_OutputPrinter") = Request.Form("txtSend_OutputPrinter")
		Session("stdReport_OutputPrinterName") = Request.Form("txtSend_OutputPrinterName")
		Session("stdReport_OutputSave") = Request.Form("txtSend_OutputSave")
		Session("stdReport_OutputSaveExisting") = Request.Form("txtSend_OutputSaveExisting")
		Session("stdReport_OutputEmail") = Request.Form("txtSend_OutputEmail")
		Session("stdReport_OutputEmailAddr") = Request.Form("txtSend_OutputEmailAddr")
		Session("stdReport_OutputEmailSubject") = Request.Form("txtSend_OutputEmailSubject")
		Session("stdReport_OutputEmailAttachAs") = Request.Form("txtSend_OutputEmailAttachAs")
		Session("stdReport_OutputFilename") = Request.Form("txtSend_OutputFilename")
	
		Dim iPromptCount As Integer
		Dim iPromptDateType As Integer
		Dim fDefaultFound As Boolean
		Dim fFirstValueDone As Boolean
		Dim sFirstValue As String
		Dim cmdDefn
		Dim prmUtilType
		Dim prmUtilID
		Dim prmRecordID
		Dim iValueType As Integer
	
		iPromptCount = 0
	
		'if Request.Form("utiltype") = 2 or Request.Form("utiltype") = 9 then
		If bStandardReportPrompt Then
			cmdDefn = CreateObject("ADODB.Command")
			cmdDefn.CommandText = "spASRIntGetStandardReportDates"
			cmdDefn.CommandType = 4	' Stored Procedure
			cmdDefn.ActiveConnection = Session("databaseConnection")

			prmUtilType = cmdDefn.CreateParameter("ReportType", 3, 1)	' 3=integer, 1=input
			cmdDefn.Parameters.Append(prmUtilType)
			prmUtilType.value = CleanNumeric(Session("StandardReport_Type"))
		Else
			cmdDefn = CreateObject("ADODB.Command")
			cmdDefn.CommandText = "sp_ASRIntGetUtilityPromptedValues"
			cmdDefn.CommandType = 4	' Stored Procedure
			cmdDefn.ActiveConnection = Session("databaseConnection")

			prmUtilType = cmdDefn.CreateParameter("utilType", 3, 1)	' 3=integer, 1=input
			cmdDefn.Parameters.Append(prmUtilType)
			prmUtilType.value = CleanNumeric(Session("utiltype"))

			prmUtilID = cmdDefn.CreateParameter("utilID", 3, 1)	' 3=integer, 1=input
			cmdDefn.Parameters.Append(prmUtilID)
			prmUtilID.value = CleanNumeric(Session("utilid"))

			prmRecordID = cmdDefn.CreateParameter("recordID", 3, 1)	' 3=integer, 1=input
			cmdDefn.Parameters.Append(prmRecordID)
			prmRecordID.value = CleanNumeric(CLng(Session("singleRecordID")))
		End If

		Err.Clear()
		Dim rstPromptedValue = cmdDefn.Execute

		If Not (rstPromptedValue.EOF And rstPromptedValue.BOF) Then
			Response.Write("<table align=center class=""outline"" cellPadding=5 cellSpacing=0 style=""width:100%;"">" & vbCrLf)
			Response.Write("  <tr>" & vbCrLf)
			Response.Write("	  <td>" & vbCrLf)
			Response.Write("			<table align=center class=""invisible"" cellspacing=0 cellpadding=0 style=""width:100%;"">" & vbCrLf)
			Response.Write("				<tr>" & vbCrLf)
			Response.Write("					<td colspan=5 align=center><H3 align=center>Prompted Values</H3></td>" & vbCrLf)
			Response.Write("				</tr>" & vbCrLf)

			Do While Not rstPromptedValue.EOF
				iPromptCount = iPromptCount + 1
				
				Response.Write("    <tr>" & vbCrLf)
				Response.Write("      <td width=20>&nbsp;</td>" & vbCrLf)
				Response.Write("      <td width='auto' nowrap>" & vbCrLf)

				If rstPromptedValue.fields("ValueType").value = 3 Then
					Response.Write("      <label " & vbCrLf)
					Response.Write("        for=""prompt_3_" & rstPromptedValue.fields("componentID").value & vbCrLf)
					Response.Write("        class=""checkbox""" & vbCrLf)
					Response.Write("        tabindex=0 " & vbCrLf)
					Response.Write("        onkeypress=""try{checkboxLabel_onKeyPress(this);}catch(e){}""" & vbCrLf)
					Response.Write("        onmouseover=""try{checkboxLabel_onMouseOver(this);}catch(e){}""" & vbCrLf)
					Response.Write("        onmouseout=""try{checkboxLabel_onMouseOut(this);}catch(e){}""" & vbCrLf)
					Response.Write("        onfocus=""try{checkboxLabel_onFocus(this);}catch(e){}""" & vbCrLf)
					Response.Write("        onblur=""try{checkboxLabel_onBlur(this);}catch(e){}"">" & vbCrLf)
				End If

				Response.Write("        " & rstPromptedValue.fields("PromptDescription").value & vbCrLf)

				If iValueType = 3 Then
					Response.Write("      </label>" & vbCrLf)
				End If

				Response.Write("      </td>" & vbCrLf)
				Response.Write("      <td width=20>&nbsp;&nbsp;</td>" & vbCrLf)
				Response.Write("      <td style='width:100%;'>" & vbCrLf)

				' Character Prompted Value
				If rstPromptedValue.fields("ValueType").value = 1 Then
					Response.Write("        <input type=text class=""text"" id=prompt_1_" & rstPromptedValue.fields("componentID").value & " name=prompt_1_" & rstPromptedValue.fields("componentID").value & " value=""" & Replace(rstPromptedValue.fields("valuecharacter").value, """", "&quot;") & """ maxlength=" & rstPromptedValue.fields("promptsize").value & " style=""WIDTH: 100%"">" & vbCrLf)
					Response.Write("        <input type=hidden id=promptMask_" & rstPromptedValue.fields("componentID").value & " name=promptMask_" & rstPromptedValue.fields("componentID").value & " value=""" & Replace(rstPromptedValue.fields("promptMask").value, """", "&quot;") & """>" & vbCrLf)

					' Numeric Prompted Value
				ElseIf rstPromptedValue.fields("ValueType").value = 2 Then
					Response.Write("        <input type=text class=""text"" id=prompt_2_" & rstPromptedValue.fields("componentID").value & " name=prompt_2_" & rstPromptedValue.fields("componentID").value & " value=""" & Replace(rstPromptedValue.fields("valuenumeric").value, ".", Session("LocaleDecimalSeparator")) & """ style=""WIDTH: 100%"">" & vbCrLf)
					Response.Write("        <input type=hidden id=promptSize_" & rstPromptedValue.fields("componentID").value & " name=promptSize" & rstPromptedValue.fields("componentID").value & " value=""" & rstPromptedValue.fields("promptSize").value & """>" & vbCrLf)
					Response.Write("        <input type=hidden id=promptDecs_" & rstPromptedValue.fields("componentID").value & " name=promptDecs" & rstPromptedValue.fields("componentID").value & " value=""" & rstPromptedValue.fields("promptDecimals").value & """>" & vbCrLf)

					' Logic Prompted Value
				ElseIf rstPromptedValue.fields("ValueType").value = 3 Then
					Response.Write("        <INPUT type=checkbox id=prompt_3_" & rstPromptedValue.fields("componentID").value & " name=prompt_3_" & rstPromptedValue.fields("componentID").value & " onclick=""checkboxClick(" & rstPromptedValue.fields("componentID").value & ")""")
					Response.Write("            onclick=""checkboxClick('" & rstPromptedValue.fields("componentID").value & "')""" & vbCrLf)
					Response.Write("            onmouseover=""try{checkbox_onMouseOver(this);}catch(e){}""" & vbCrLf)
					Response.Write("            onmouseout=""try{checkbox_onMouseOut(this);}catch(e){}""")
					If rstPromptedValue.fields("valuelogic").value Then
						Response.Write(" CHECKED/>" & vbCrLf)
					Else
						Response.Write("/>" & vbCrLf)
					End If
					Response.Write("        <input type=hidden id=promptChk_" & rstPromptedValue.fields("componentID").value & " name=promptChk_" & rstPromptedValue.fields("componentID").value & " value=" & rstPromptedValue.fields("valuelogic").value & ">" & vbCrLf)
							 
					' Date Prompted Value
				ElseIf rstPromptedValue.fields("ValueType").value = 4 Then

					If bStandardReportPrompt Then
						Response.Write("        <input type=text class=""text"" id=prompt_" & rstPromptedValue.fields("StartEndType").value & "_" & rstPromptedValue.fields("componentID").value & " name=prompt_" & rstPromptedValue.fields("StartEndType").value & "_" & rstPromptedValue.fields("componentID").value & " value=""")
					Else
						Response.Write("        <input type=text class=""text"" id=prompt_4_" & rstPromptedValue.fields("componentID").value & " name=prompt_4_" & rstPromptedValue.fields("componentID").value & " value=""")
					End If
					
					If (IsDBNull(rstPromptedValue.fields("promptDateType").value)) Or (rstPromptedValue.fields("promptDateType").value = vbNullString) Then
						iPromptDateType = 0
					Else
						iPromptDateType = rstPromptedValue.fields("promptDateType").value
					End If
				
					Dim iDay
					Dim dtDate
					Dim iMonth
								
					Select Case iPromptDateType
						Case 0
							' Explicit value
							If Not IsDBNull(rstPromptedValue.fields("valuedate").value) Then
								If (CStr(rstPromptedValue.fields("valuedate").value) <> "00:00:00") And _
										(CStr(rstPromptedValue.fields("valuedate").value) <> "12:00:00 AM") Then
									Response.Write(ConvertSQLDateToLocale(rstPromptedValue.fields("valuedate").value))
								End If
							End If
										
						Case 1
							' Current date
							Response.Write(ConvertSQLDateToLocale(Date.Now))
						
						Case 2
							' Start of current month
							iDay = (Day(Date.Now) * -1) + 1
							dtDate = DateAdd("d", iDay, Date.Now)
							Response.Write(ConvertSQLDateToLocale(dtDate))
						
						Case 3
							' End of current month
							iDay = (Day(Date.Now) * -1) + 1
							dtDate = DateAdd("d", iDay, Date.Now)
							dtDate = DateAdd("m", 1, dtDate)
							dtDate = DateAdd("d", -1, dtDate)
							Response.Write(ConvertSQLDateToLocale(dtDate))
						
						Case 4
							' Start of current year
							iDay = (Day(Date.Now) * -1) + 1
							iMonth = (Month(Date.Now) * -1) + 1
							dtDate = DateAdd("d", iDay, Date.Now)
							dtDate = DateAdd("m", iMonth, dtDate)
							Response.Write(ConvertSQLDateToLocale(dtDate))
						
						Case 5
							' End of current year
							iDay = (Day(Date.Now) * -1) + 1
							iMonth = (Month(Date.Now) * -1) + 1
							dtDate = DateAdd("d", iDay, Date.Now)
							dtDate = DateAdd("m", iMonth, dtDate)
							dtDate = DateAdd("yyyy", 1, dtDate)
							dtDate = DateAdd("d", -1, dtDate)
							Response.Write(ConvertSQLDateToLocale(dtDate))
					End Select
					Response.Write(""" style=""WIDTH: 100%"">" & vbCrLf)

					' Lookup Prompted Value
				ElseIf rstPromptedValue.fields("ValueType").value = 5 Then
					Response.Write("        <SELECT STYLE=""width:100%;"" id=promptLookup_" & rstPromptedValue.fields("componentID").value & " name=promptLookup_" & rstPromptedValue.fields("componentID").value & " class=""combo"" style=""WIDTH: 100%"" onchange=""comboChange(" & rstPromptedValue.fields("componentID").value & ")"">" & vbCrLf)

					fDefaultFound = False
					fFirstValueDone = False
					sFirstValue = ""
					Dim cmdLookupValues
					Dim prmColumnID
					Dim rstLookupValues

					' Get the lookup values.
					cmdLookupValues = CreateObject("ADODB.Command")
					cmdLookupValues.CommandText = "sp_ASRIntGetLookupValues"
					cmdLookupValues.CommandType = 4	' Stored Procedure
					cmdLookupValues.ActiveConnection = Session("databaseConnection")

					prmColumnID = cmdLookupValues.CreateParameter("columnID", 3, 1)
					cmdLookupValues.Parameters.Append(prmColumnID)
					prmColumnID.value = CleanNumeric(rstPromptedValue.fields("fieldColumnID").value)

					Err.Clear()
					rstLookupValues = cmdLookupValues.Execute

					Do While Not rstLookupValues.EOF
						Response.Write("          <OPTION")
						
						If Not fFirstValueDone Then
							sFirstValue = rstLookupValues.Fields(0).Value
							fFirstValueDone = True
						End If

						Dim sOptionValue
										
						If rstLookupValues.fields(0).type = 135 Then
							' Field is a date so format as such.
							sOptionValue = ConvertSQLDateToLocale(rstLookupValues.Fields(0).Value)
							If sOptionValue = ConvertSQLDateToLocale(rstPromptedValue.fields("valuecharacter").Value) Then
								Response.Write(" SELECTED")
								fDefaultFound = True
							End If
							Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
						ElseIf rstLookupValues.fields(0).type = 131 Then
							' Field is a numeric so format as such.
							sOptionValue = Replace(rstLookupValues.Fields(0).Value, ".", Session("LocaleDecimalSeparator"))
							If (Not IsDBNull(rstLookupValues.Fields(0).Value)) And (Not IsDBNull(rstPromptedValue.fields("valuecharacter").Value)) Then
								If FormatNumber(rstLookupValues.Fields(0).Value) = FormatNumber(rstPromptedValue.fields("valuecharacter").Value) Then
									Response.Write(" SELECTED")
									fDefaultFound = True
								End If
							End If
							Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
						ElseIf rstLookupValues.fields(0).type = 11 Then
							' Field is a logic so format as such.
							sOptionValue = rstLookupValues.Fields(0).Value
							If sOptionValue = rstPromptedValue.fields("valuecharacter").Value Then
								Response.Write(" SELECTED")
								fDefaultFound = True
							End If
							Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
						Else
							sOptionValue = RTrim(rstLookupValues.Fields(0).Value)
							If sOptionValue = rstPromptedValue.fields("valuecharacter").Value Then
								Response.Write(" SELECTED")
								fDefaultFound = True
							End If
							Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
						End If

						rstLookupValues.MoveNext()
					Loop

					Response.Write("        </SELECT>" & vbCrLf)

					Dim sDefaultValue
								
					If fDefaultFound Then
						sDefaultValue = rstPromptedValue.Fields("valuecharacter").Value
					Else
						sDefaultValue = sFirstValue
					End If

					If rstLookupValues.fields(0).type = 135 Then
						' Date.
						Response.Write("        <input type=hidden id=prompt_4_" & rstPromptedValue.fields("componentID").value & " name=prompt_4_" & rstPromptedValue.fields("componentID").value & " value=" & ConvertSQLDateToLocale(sDefaultValue) & ">" & vbCrLf)
					ElseIf rstLookupValues.fields(0).type = 131 Then
						' Numeric
						Response.Write("        <input type=hidden id=prompt_2_" & rstPromptedValue.fields("componentID").value & " name=prompt_2_" & rstPromptedValue.fields("componentID").value & " value=" & Replace(sDefaultValue, ".", Session("LocaleDecimalSeparator")) & ">" & vbCrLf)
					ElseIf rstLookupValues.fields(0).type = 11 Then
						' Logic
						Response.Write("        <input type=hidden id=prompt_3_" & rstPromptedValue.fields("componentID").value & " name=prompt_3_" & rstPromptedValue.fields("componentID").value & " value=" & sDefaultValue & ">" & vbCrLf)
					Else
						Response.Write("        <input type=hidden id=prompt_1_" & rstPromptedValue.fields("componentID").value & " name=prompt_1_" & rstPromptedValue.fields("componentID").value & " value=""" & Replace(sDefaultValue, """", "&quot;") & """>" & vbCrLf)
					End If

					' Release the ADO recordset object.
					rstLookupValues.close()
					rstLookupValues = Nothing
				End If
				
				Response.Write("					</td>" & vbCrLf)
				Response.Write("					<td width=20 height=10>&nbsp;</td>" & vbCrLf)
				Response.Write("				</tr>" & vbCrLf)

				rstPromptedValue.MoveNext()
			Loop
	%>
	<tr>
		<td colspan="5" height="10">&nbsp;</td>
	</tr>
	<tr height="20">
		<td width="20">&nbsp;</td>
		<td colspan="3" align='center'>
			<table class="invisible" cellspacing="0" cellpadding="0" align='center'>
				<td width="20">&nbsp;</td>
				<td width="80">

					<input type="button" class="btn" name="Submit" value="OK" style="WIDTH: 80px"
						onclick="SubmitPrompts()" />
				</td>
				<td width="20">&nbsp;</td>
				<td width="80">
					<input type="button" class="btn" name="Cancel" value="Cancel" style="WIDTH: 80px"
						onclick="CancelClick()" />
				</td>
			</table>
		</td>
		<td width="20">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="5" height="5">&nbsp;</td>
	</tr>
	</table>
		</td>
	</tr>
</table>
	<%
	End If
		
	rstPromptedValue.close()
	rstPromptedValue = Nothing

	Response.Write("<input type=""hidden"" id=""txtPromptCount"" name=""txtPromptCount"" value=" & iPromptCount & ">" & vbCrLf)
	%>

	<input type="hidden" id="utiltype" name="utiltype" value="<%=Session("utiltype")%>">
	<input type="hidden" id="utilid" name="utilid" value='<%=Session("utilid")%>'>
	<input type="hidden" id="utilname" name="utilname" value="<%=Replace(Session("utilname").ToString(), """", "&quot;")%>">
	<input type="hidden" id="action" name="action" value='<%=Session("action")%>'>
	<input type="hidden" id="lastPrompt" name="lastPrompt" value="">
	<input type="hidden" id="StandardReportPrompt" name="StandardReportPrompt" value="<%=bStandardReportPrompt%>">
	<input type="hidden" id="RunInOptionFrame" name="RunInOptionFrame" value='<%=(Session("optionAction") = "STDREPORT_DATEPROMPT") %>'>
	<input type="hidden" id="txtLocaleDateFormat" name="txtLocaleDateFormat" value="">
	<input type="hidden" id="txtLocaleDecimalSeparator" name="txtLocaleDecimalSeparator" value="">
	<input type="hidden" id="txtLocaleThousandSeparator" name="txtLocaleThousandSeparator" value="">
</form>



<script type="text/javascript">

	function SubmitPrompts() {

		// Validate the prompt values before submitting the form.
		var controlCollection = frmPromptedValues.elements;
		if (controlCollection != null) {
			for (var i = 0; i < controlCollection.length; i++) {
				sControlName = controlCollection.item(i).name;
				sControlPrefix = sControlName.substr(0, 7);

				if (sControlPrefix == "prompt_") {
					// Get the control's data type.
					iType = new Number(sControlName.substring(7, 8));
					if ((iType == 1) || (iType == 2) || (iType == 4)) {
						// Validate character, numeric and date prompts.
						// Logic and lookup prompts do not need validation.
						if (ValidatePrompt(controlCollection.item(i), iType) == false) {
							return;
						}
					}
				}
			}
		}

		// Everything OK. Submit the form.
		if (menu_isSSIMode() == true) {
			OpenHR.submitForm(frmPromptedValues, "workframe", true);
		} else {
			$(".popup").dialog("close");
			OpenHR.showInReportFrame(frmPromptedValues, true);
		}


	}

	function CancelClick() {

		if (frmPromptedValues.StandardReportPrompt.value == "True") {
			if (frmPromptedValues.RunInOptionFrame.value == "True") {
				var frmParent = window.dialogArguments.OpenHR.getForm("workframe", "frmRecordEditForm");

				window.parent.document.all.item("workframeset").cols = "*, 0";
				window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "visible";
				window.dialogArguments.OpenHR.submitForm(frmParent, null, false);
			}
			else {
				closeclick();
			}
		}
		else {
			closeclick();
		}
	}

	function ValidatePrompt(pctlPrompt, piDataType) {
		// Validate the given prompt value.
		var fOK;
		var reBackSlash = new RegExp("\\\\", "gi");
		var reDoubleBackSlash = new RegExp("\\\\\\\\", "gi");
		var sDecimalSeparator;
		var sThousandSeparator;
		var sPoint;
		var sConvertedValue;
		var sValue;

		sDecimalSeparator = "\\";
		sDecimalSeparator = sDecimalSeparator.concat(frmPromptedValues.txtLocaleDecimalSeparator.value);
		var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");

		sThousandSeparator = "\\";
		sThousandSeparator = sThousandSeparator.concat(frmPromptedValues.txtLocaleThousandSeparator.value);
		var reThousandSeparator = new RegExp(sThousandSeparator, "gi");

		sPoint = "\\.";
		var rePoint = new RegExp(sPoint, "gi");

		fOK = true;

		if ((fOK == true) && (piDataType == 2)) {
			// Numeric column.
			// Ensure that the value entered is numeric.
			sValue = pctlPrompt.value;

			if (sValue.length == 0) {
				sValue = "0";
				pctlPrompt.value = 0;
			}

			// Convert the value from locale to UK settings for use with the isNaN funtion.
			sConvertedValue = new String(sValue);
			// Remove any thousand separators.
			sConvertedValue = sConvertedValue.replace(reThousandSeparator, "");
			pctlPrompt.value = sConvertedValue;

			// Convert any decimal separators to '.'.
			if (frmPromptedValues.txtLocaleDecimalSeparator.value != ".") {
				// Remove decimal points.
				sConvertedValue = sConvertedValue.replace(rePoint, "A");
				// replace the locale decimal marker with the decimal point.
				sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
			}

			if (isNaN(sConvertedValue) == true) {
				fOK = false;
				OpenHR.messageBox("Invalid numeric value entered.");
				window.focus();
				pctlPrompt.focus();
			}
		}

		if ((fOK == true) && (piDataType == 4)) {
			// Date column.
			// Ensure that the value entered is a date.
			sValue = pctlPrompt.value;

			if (sValue.length == 0) {
				fOK = false;
			}
			else {
				// Convert the date to SQL format (use this as a validation check).
				// An empty string is returned if the date is invalid.
				//sValue = OpenHR.convertLocaleDateToSQL(sValue);
				sValue = localconvertLocaleDateToSQL(sValue);

				if (sValue.length == 0) {
					fOK = false;
				}
				else {
					pctlPrompt.value = OpenHR.ConvertSQLDateToLocale(sValue);
				}
			}

			if (fOK == false) {
				OpenHR.messageBox("Invalid date value entered.");
				window.focus();
				pctlPrompt.focus();
			}
		}

		if ((fOK == true) && (piDataType == 1)) {
			// Character column.
			// Ensure that the value entered matches the required mask (if there is one).
			sMaskCtlName = "promptMask_" + pctlPrompt.name.substring(9, pctlPrompt.name.length);
			//   sMaskCtlName = sMaskCtlName.toUpperCase();

			fFound = false;
			var controlCollection = frmPromptedValues.elements;
			if (controlCollection != null) {
				for (i = 0; i < controlCollection.length; i++) {
					if (controlCollection.item(i).name.toUpperCase() == sMaskCtlName) {
						fFound = true;
						break;
					}
				}
			}

			if (fFound == true) {
				sMask = frmPromptedValues.elements(sMaskCtlName).value;
				sValue = pctlPrompt.value;
				// Need to get rid of the backslash characters that precede literals.
				// But remember that two backslashes give a literal backslash that does not want
				// to be got rid of.
				sTemp = sMask.replace(reDoubleBackSlash, "a");
				sTemp = sTemp.replace(reBackSlash, "");
				if (sMask.length > 0) {
					if (sTemp.length != sValue.length) {
						fOK = false;
					}
					else {
						// Prompt values length matches mask length, so now check each character.
						fFollowingBackslash = false;
						iIndex = 0;
						for (i = 0; i < sMask.length; i++) {
							sValueChar = sValue.substring(iIndex, iIndex + 1);

							if (fFollowingBackslash == false) {
								switch (sMask.substring(i, i + 1)) {
									case "A":
										// Character must be uppercase.
										if (sValueChar.toUpperCase() != sValueChar) {
											fOK = false;
										}
										else {
											iNumber = new Number(sValueChar);
											if (isNaN(iNumber) == false) {
												fOK = true;
											}
										}
										iIndex = iIndex + 1;
										break;

									case "a":
										// Character must be lowercase.
										if (sValueChar.toLowerCase() != sValueChar) {
											fOK = false;
										}
										else {
											iNumber = new Number(sValueChar);
											if (isNaN(iNumber) == false) {
												fOK = false;
											}
										}
										iIndex = iIndex + 1;
										break;

									case "9":
										// Character must be numeric (0-9).
										iNumber = new Number(sValueChar);
										if (isNaN(iNumber) == true) {
											fOK = false;
										}
										iIndex = iIndex + 1;
										break;

									case "#":
										// Character must be numeric (0-9) or symbolic (+-%\).
										iNumber = new Number(sValueChar);
										if ((isNaN(iNumber) == true) &&
												(sValueChar != "+") &&
												(sValueChar != "-") &&
												(sValueChar != "%") &&
												(sValueChar != "\\")) {
											fOK = false;
										}
										iIndex = iIndex + 1;
										break;

									case "B":
										// Character must be logic (0 or 1).
										if ((sValueChar != "0") &&
												(sValueChar != "1")) {
											fOK = false;
										}
										iIndex = iIndex + 1;
										break;

									case "\\":
										// Following character is literal.
										fFollowingBackslash = true;
										break;

									default:
										// Literal.
										if (sMask.substring(i, i + 1) != sValueChar) {
											fOK = false;
										}
										iIndex = iIndex + 1;
								}
							}
							else {
								fFollowingBackslash = false;
								if (sMask.substring(i, i + 1) != sValueChar) {
									fOK = false;
								}
								iIndex = iIndex + 1;
							}

							if (fOK == false) {
								break;
							}
						}
					}
				}

				if (fOK == false) {
					OpenHR.messageBox("The entered value does not match the required format (" + sMask + ").");
					window.focus();
					pctlPrompt.focus();
				}
			}
		}

		return fOK;
	}

	function localconvertLocaleDateToSQL(psDateString) {
		/* Convert the given date string (in locale format) into 
		SQL format (mm/dd/yyyy). */
		var sDateFormat;
		var iDays;
		var iMonths;
		var iYears;
		var sDays;
		var sMonths;
		var sYears;
		var iValuePos;
		var sTempValue;
		var sValue;
		var iLoop;

		sDateFormat = frmPromptedValues.txtLocaleDateFormat.value;

		sDays = "";
		sMonths = "";
		sYears = "";
		iValuePos = 0;

		// Trim leading spaces.
		sTempValue = psDateString.substr(iValuePos, 1);
		while (sTempValue.charAt(0) == " ") {
			iValuePos = iValuePos + 1;
			sTempValue = psDateString.substr(iValuePos, 1);
		}

		for (iLoop = 0; iLoop < sDateFormat.length; iLoop++) {
			if ((sDateFormat.substr(iLoop, 1).toUpperCase() == 'D') && (sDays.length == 0)) {
				sDays = psDateString.substr(iValuePos, 1);
				iValuePos = iValuePos + 1;
				sTempValue = psDateString.substr(iValuePos, 1);

				if (isNaN(sTempValue) == false) {
					sDays = sDays.concat(sTempValue);
				}
				iValuePos = iValuePos + 1;
			}

			if ((sDateFormat.substr(iLoop, 1).toUpperCase() == 'M') && (sMonths.length == 0)) {
				sMonths = psDateString.substr(iValuePos, 1);
				iValuePos = iValuePos + 1;
				sTempValue = psDateString.substr(iValuePos, 1);

				if (isNaN(sTempValue) == false) {
					sMonths = sMonths.concat(sTempValue);
				}
				iValuePos = iValuePos + 1;
			}

			if ((sDateFormat.substr(iLoop, 1).toUpperCase() == 'Y') && (sYears.length == 0)) {
				sYears = psDateString.substr(iValuePos, 1);
				iValuePos = iValuePos + 1;
				sTempValue = psDateString.substr(iValuePos, 1);

				if (isNaN(sTempValue) == false) {
					sYears = sYears.concat(sTempValue);
				}
				iValuePos = iValuePos + 1;
				sTempValue = psDateString.substr(iValuePos, 1);

				if (isNaN(sTempValue) == false) {
					sYears = sYears.concat(sTempValue);
				}
				iValuePos = iValuePos + 1;
				sTempValue = psDateString.substr(iValuePos, 1);

				if (isNaN(sTempValue) == false) {
					sYears = sYears.concat(sTempValue);
				}
				iValuePos = iValuePos + 1;
			}

			// Skip non-numerics
			sTempValue = psDateString.substr(iValuePos, 1);
			while (isNaN(sTempValue) == true) {
				iValuePos = iValuePos + 1;
				sTempValue = psDateString.substr(iValuePos, 1);
			}
		}

		while (sDays.length < 2) {
			sTempValue = "0";
			sDays = sTempValue.concat(sDays);
		}

		while (sMonths.length < 2) {
			sTempValue = "0";
			sMonths = sTempValue.concat(sMonths);
		}

		while (sYears.length < 2) {
			sTempValue = "0";
			sYears = sTempValue.concat(sYears);
		}

		if (sYears.length == 2) {
			iValue = parseInt(sYears);
			if (iValue < 30) {
				sTempValue = "20";
			}
			else {
				sTempValue = "19";
			}

			sYears = sTempValue.concat(sYears);
		}

		while (sYears.length < 4) {
			sTempValue = "0";
			sYears = sTempValue.concat(sYears);
		}

		sTempValue = sMonths.concat("/");
		sTempValue = sTempValue.concat(sDays);
		sTempValue = sTempValue.concat("/");
		sTempValue = sTempValue.concat(sYears);

		sValue = OpenHR.ConvertSQLDateToLocale(sTempValue);

		iYears = parseInt(sYears);

		while (sMonths.substr(0, 1) == "0") {
			sMonths = sMonths.substr(1);
		}
		iMonths = parseInt(sMonths);

		while (sDays.substr(0, 1) == "0") {
			sDays = sDays.substr(1);
		}
		iDays = parseInt(sDays);

		var newDateObj = new Date(iYears, iMonths - 1, iDays);

		if ((newDateObj.getDate() != iDays) ||
				(newDateObj.getMonth() + 1 != iMonths) ||
				(newDateObj.getFullYear() != iYears)) {
			return "";
		}
		else {
			return sTempValue;
		}
	}

	function checkboxClick(piPromptID) {
		sSource = "prompt_3_" + piPromptID;
		sDest = "promptChk_" + piPromptID;

		frmPromptedValues.elements.item(sDest).value = frmPromptedValues.elements.item(sSource).checked;
	}

	function comboChange(piPromptID) {
		sSource = "promptLookup_" + piPromptID;
		ctlSource = frmPromptedValues.elements.item(sSource);

		var controlCollection = frmPromptedValues.elements;
		if (controlCollection != null) {
			for (i = 0; i < controlCollection.length; i++) {
				sControlName = controlCollection.item(i).name;
				sControlPrefix = sControlName.substr(0, 7);
				sControlID = sControlName.substr(9, sControlName.length);

				if ((sControlPrefix == "prompt_") && (sControlID == piPromptID)) {
					controlCollection.item(i).value = ctlSource.options[ctlSource.selectedIndex].text;
				}
			}
		}
	}

</script>

<script type="text/javascript">
	promptedvalues_window_onload();
</script>

