<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>


<script type="text/javascript">

	function promptedvalues_window_onload() {

		$('#frmPromptedValues').submit(function(e) {
			e.preventDefault();
			SubmitPrompts();
		});

		var frmPromptedValues = document.getElementById("frmPromptedValues");
		$(".datepicker").datepicker();

		$(document).on('keydown', '.datepicker', function (event) {

			switch (event.keyCode) {
				case 113:
				    $(this).datepicker("setDate", new Date())
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
		
		frmPromptedValues.txtLocaleDecimalSeparator.value = '<%:LocaleDecimalSeparator()%>';
		frmPromptedValues.txtLocaleThousandSeparator.value = '<%:Html.Raw(LocaleThousandSeparator())%>';

		if (frmPromptedValues.RunInOptionFrame.value == "True") {
			$("#optionframe").attr("data-framesource", "UTIL_RUN_PROMPTEDVALUES");
			$("#workframe").hide();
			$("#optionframe").show();
		} else {

			if (window.menu_isSSIMode() == true) {
				$("#workframe").attr("data-framesource", "UTIL_RUN_PROMPTEDVALUES");
			} else {
				$("#reportframe").attr("data-framesource", "UTIL_RUN_PROMPTEDVALUES");
			}

		}

		if (frmPromptedValues.txtPromptCount.value == 0) {

			if (menu_isSSIMode() == true) {
				OpenHR.submitForm(frmPromptedValues, "workframe", true);
			} else {
				OpenHR.showInReportFrame(frmPromptedValues, true);
			}

		} else {

			if (menu_isSSIMode() == false) {
				$(".popup").dialog('option', 'title', "Prompted Value"); 
				$(".popup").dialog("open");
			}

			// Set focus on the first prompt control.
			var controlCollection = frmPromptedValues.elements;
			if (controlCollection != null) {
				for (var i = 0; i < controlCollection.length; i++) {
					var sControlName = controlCollection.item(i).name;
					var sControlPrefix = sControlName.substr(0, 7);

					if ((sControlPrefix == "prompt_") || (sControlName.substr(0, 13) == "promptLookup_")) {

						if (sControlName.substr(0, 9) != "prompt_4_") {
							controlCollection.item(i).focus();
							break;
						}
					}
				}

			}
		}

	}

</script>


<div class="absolutefull">
		<div class="pageTitleDiv">
			<a href='javascript:loadPartialView("linksMain", "Home", "workframe", null);' title='Back'>
				<i class='pageTitleIcon icon-circle-arrow-left'></i>
			</a>
			<span class="pageTitle"><% =Session("utilname")%></span>
		</div>
		<br/>

	<div id="dataRow">

		<p tabindex="1"></p>

		<form name="frmPromptedValues" id="frmPromptedValues" method="POST" action="util_run_promptedvalues_submit">

			<%
				' Get variables for Absence Breakdown / Bradford Factor
				Session("stdReport_StartDate") = Request.Form("txtFromDate")
				Session("stdReport_EndDate") = Request.Form("txtToDate")
				Session("stdReport_AbsenceTypes") = Request.Form("txtAbsenceTypes")
				Session("stdReport_FilterID") = Request.Form("txtBaseFilterID")
				Session("stdReport_FilterName") = Request.Form("txtFilterName")
				Session("stdReport_PicklistID") = Request.Form("txtBasePicklistID")
				Session("stdReport_PicklistName") = Request.Form("txtPicklistName")
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
				Session("stdReport_OutputSave") = Request.Form("txtSend_OutputSave")
				Session("stdReport_OutputSaveExisting") = 0
				Session("stdReport_OutputEmail") = Request.Form("txtSend_OutputEmail")
				Session("stdReport_OutputEmailAddr") = Request.Form("txtSend_OutputEmailAddr")
				Session("stdReport_OutputEmailSubject") = Request.Form("txtSend_OutputEmailSubject")
				Session("stdReport_OutputEmailAttachAs") = Request.Form("txtSend_OutputEmailAttachAs")
				Session("stdReport_OutputFilename") = Request.Form("txtSend_OutputFilename")

				Dim objDatabaseAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
				Dim rstPromptedValue As DataTable
				Dim rstLookupValues As DataTable
				
				Dim fDefaultFound As Boolean
				Dim fFirstValueDone As Boolean
				Dim sFirstValue As String

				Dim iValueType As Integer
	
				Dim bAddUploadTemplate As Boolean = (CType(Session("utiltype"), UtilityType) = UtilityType.utlMailMerge)
				Dim iPromptCount = CInt(IIf(bAddUploadTemplate, 1, 0))
				
				rstPromptedValue = objDatabaseAccess.GetDataTable("spASRIntGetUtilityPromptedValues", CommandType.StoredProcedure, _
										New SqlParameter("piUtilType", SqlDbType.Int) With {.Value = CInt(CleanNumeric(Session("utiltype")))}, _
										New SqlParameter("piUtilID", SqlDbType.Int) With {.Value = CInt(CleanNumeric(Session("utilid")))}, _
										New SqlParameter("piRecordID", SqlDbType.Int) With {.Value = Session("singleRecordID")})
				
				If rstPromptedValue.Rows.Count > 0 Then

					Response.Write("			<table align=center class=""invisible"" cellspacing=5 cellpadding=0 style=""width:100%;"">" & vbCrLf)

					For Each objRow As DataRow In rstPromptedValue.Rows
					
						iPromptCount += 1
				
						Response.Write("    <tr>" & vbCrLf)
						Response.Write("      <td width='auto' nowrap>" & vbCrLf)

						If objRow("ValueType") = 3 Then
							Response.Write("      <label " & vbCrLf)
							Response.Write("        for=""prompt_3_" & objRow("componentID") & vbCrLf)
							Response.Write("        class=""checkbox""" & vbCrLf)
							Response.Write("        tabindex=0>" & vbCrLf)
						End If

						Response.Write("        " & objRow("PromptDescription") & vbCrLf)

						If iValueType = 3 Then
							Response.Write("      </label>" & vbCrLf)
						End If

						Response.Write("      </td>" & vbCrLf)
						Response.Write("      <td width=20>&nbsp;&nbsp;</td>" & vbCrLf)
						Response.Write("      <td style='width:100%;'>" & vbCrLf)

						' Character Prompted Value
						If objRow("ValueType") = 1 Then
							Response.Write("        <input type=text class=""text"" id=prompt_1_" & objRow("componentID") & " name=prompt_1_" & objRow("componentID") & " value=""" & Replace(CType(objRow("valuecharacter"), String), """", "&quot;") & """ maxlength=" & objRow("promptsize") & " style=""WIDTH: 100%"">" & vbCrLf)
							Response.Write("        <input type=hidden id=promptMask_" & objRow("componentID") & " name=promptMask_" & objRow("componentID") & " value=""" & Replace(CType(objRow("promptMask"), String), """", "&quot;") & """>" & vbCrLf)

							' Numeric Prompted Value
						ElseIf objRow("ValueType") = 2 Then
							Response.Write("        <input type=text class=""text"" id=prompt_2_" & objRow("componentID") & " name=prompt_2_" & objRow("componentID") & " value=""" & Replace(CType(objRow("valuenumeric"), String), ".", CType(Session("LocaleDecimalSeparator"), String)) & """ style=""WIDTH: 100%"">" & vbCrLf)
							Response.Write("        <input type=hidden id=promptSize_" & objRow("componentID") & " name=promptSize" & objRow("componentID") & " value=""" & objRow("promptSize") & """>" & vbCrLf)
							Response.Write("        <input type=hidden id=promptDecs_" & objRow("componentID") & " name=promptDecs" & objRow("componentID") & " value=""" & objRow("promptDecimals") & """>" & vbCrLf)

							' Logic Prompted Value
						ElseIf objRow("ValueType") = 3 Then
							Response.Write("        <input type=checkbox id=prompt_3_" & objRow("componentID") & " name=prompt_3_" & objRow("componentID") & " onclick=""checkboxClick(" & objRow("componentID") & ")""")
							Response.Write("            onclick=""checkboxClick('" & objRow("componentID") & "')""" & vbCrLf)
							Response.Write("            onmouseover=""try{checkbox_onMouseOver(this);}catch(e){}""" & vbCrLf)
							Response.Write("            onmouseout=""try{checkbox_onMouseOut(this);}catch(e){}""")
							If objRow("valuelogic") Then
								Response.Write(" CHECKED/>" & vbCrLf)
							Else
								Response.Write("/>" & vbCrLf)
							End If
							Response.Write("        <input type=hidden id=promptChk_" & objRow("componentID") & " name=promptChk_" & objRow("componentID") & " value=" & objRow("valuelogic") & ">" & vbCrLf)
							 
							' Date Prompted Value
						ElseIf objRow("ValueType") = 4 Then

							Response.Write("        <input type=text class=""datepicker"" id=prompt_4_" & objRow("componentID") & " name=prompt_4_" & objRow("componentID") & " value=""")
									
							Dim dtDate As Date = CalculatePromptedDate(objRow)
							Response.Write(ConvertSQLDateToLocale(dtDate))
							Response.Write(""" style=""WIDTH: 100%"">" & vbCrLf)

							' Lookup Prompted Value
						ElseIf objRow("ValueType") = 5 Then
							Response.Write("        <SELECT STYLE=""width:100%;"" id=promptLookup_" & objRow("componentID") & " name=promptLookup_" & objRow("componentID") & " class=""combo"" style=""WIDTH: 100%"" onchange=""comboChange(" & objRow("componentID") & ")"">" & vbCrLf)

							fDefaultFound = False
							fFirstValueDone = False
							sFirstValue = ""

							rstLookupValues = GetLookupValues(CInt(objRow("fieldColumnID")))
							
							For Each objLookupRow As DataRow In rstLookupValues.Rows
								
								Response.Write("          <OPTION")
						
								If Not fFirstValueDone Then
									sFirstValue = objLookupRow(0).ToString()
									fFirstValueDone = True
								End If

								Dim sOptionValue As String
										
								If rstLookupValues.Columns(0).DataType.Name.ToLower() = "datetime" Then
									' Field is a date so format as such.
									sOptionValue = ConvertSQLDateToLocale(objLookupRow(0))
									If sOptionValue = ConvertSQLDateToLocale(objRow("valuecharacter").ToString()) Then
										Response.Write(" SELECTED")
										fDefaultFound = True
									End If
									Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
								ElseIf rstLookupValues.Columns(0).DataType.Name.ToLower() = "decimal" Then
									' Field is a numeric so format as such.
									sOptionValue = Replace(CType(objLookupRow(0), String), ".", CType(Session("LocaleDecimalSeparator"), String))
									If (Not IsDBNull(objLookupRow(0))) And (Not IsDBNull(objRow("valuecharacter"))) Then
										If FormatNumber(objLookupRow(0)) = FormatNumber(objRow("valuecharacter")) Then
											Response.Write(" SELECTED")
											fDefaultFound = True
										End If
									End If
									Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
								ElseIf rstLookupValues.Columns(0).DataType.Name.ToLower() = "boolean" Then
									' Field is a logic so format as such.
									sOptionValue = objLookupRow(0).ToString()
									If sOptionValue = objRow("valuecharacter") Then
										Response.Write(" SELECTED")
										fDefaultFound = True
									End If
									Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
								Else
									sOptionValue = RTrim(objLookupRow(0).ToString())
									If sOptionValue = objRow("valuecharacter") Then
										Response.Write(" SELECTED")
										fDefaultFound = True
									End If
									Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
								End If

							Next
							
							Response.Write("        </SELECT>" & vbCrLf)

							Dim sDefaultValue As String
								
							If fDefaultFound Then
								sDefaultValue = objRow("valuecharacter").ToString()
							Else
								sDefaultValue = sFirstValue
							End If

							Select Case rstLookupValues.Columns(0).DataType.Name.ToLower()
								Case "datetime"
									Response.Write("        <input type=hidden id=prompt_4_" & objRow("componentID") & " name=prompt_4_" & objRow("componentID") & " value=" & ConvertSQLDateToLocale(sDefaultValue) & ">" & vbCrLf)

								Case "decimal"
									Response.Write("        <input type=hidden id=prompt_2_" & objRow("componentID") & " name=prompt_2_" & objRow("componentID") & " value=" & Replace(sDefaultValue, ".", CType(Session("LocaleDecimalSeparator"), String)) & ">" & vbCrLf)
									
								Case "boolean"
									Response.Write("        <input type=hidden id=prompt_3_" & objRow("componentID") & " name=prompt_3_" & objRow("componentID") & " value=" & sDefaultValue & ">" & vbCrLf)
									
								Case Else
									Response.Write("        <input type=hidden id=prompt_1_" & objRow("componentID") & " name=prompt_1_" & objRow("componentID") & " value=""" & Replace(sDefaultValue, """", "&quot;") & """>" & vbCrLf)
									
							End Select
							
							rstLookupValues = Nothing
						End If
				
						Response.Write("					</td>" & vbCrLf)
						Response.Write("					<td width=20 height=10>&nbsp;</td>" & vbCrLf)
						Response.Write("				</tr>" & vbCrLf)

					Next
					
					Response.Write("</table>" & vbCrLf)
					
			%>

	<%
	End If
		
	rstPromptedValue = Nothing

	Response.Write("<input type=""hidden"" id=""txtPromptCount"" name=""txtPromptCount"" value=" & iPromptCount & ">" & vbCrLf)
	%>

			<input type="hidden" id="utiltype" name="utiltype" value="<%=Session("utiltype")%>">
			<input type="hidden" id="utilid" name="utilid" value='<%=Session("utilid")%>'>
			<input type="hidden" id="utilname" name="utilname" value="<%=Replace(Session("utilname").ToString(), """", "&quot;")%>">
			<input type="hidden" id="action" name="action" value='<%=Session("action")%>'>
			<input type="hidden" id="lastPrompt" name="lastPrompt" value="">
			<input type="hidden" id="RunInOptionFrame" name="RunInOptionFrame" value='<%=(Session("optionAction") = "STDREPORT_DATEPROMPT") %>'>
			<input type="hidden" id="txtLocaleDateFormat" name="txtLocaleDateFormat" value="">
			<input type="hidden" id="txtLocaleDecimalSeparator" name="txtLocaleDecimalSeparator" value="">
			<input type="hidden" id="txtLocaleThousandSeparator" name="txtLocaleThousandSeparator" value="">

			<%=Html.AntiForgeryToken()%>
		</form>
		
		<%If bAddUploadTemplate Then%>
			<form name="frmTemplateFile" id="frmTemplateFile" method="post" enctype="multipart/form-data" action="util_run_uploadtemplate" target="submit-iframe">
				Template File: <input style="width: 500px" type="file" id="TemplateFile" name="TemplateFile" onchange="SubmitTemplate();" />
				<%=Html.AntiForgeryToken()%>
			</form>		
		<%End If%>
		
		<% If iPromptCount > 0 Then%>
			<br/>						
			<div class="centered">			
				<input type="button" id="butPromptedSubmit" class="btn" name="butPromptedSubmit" value="OK" style="WIDTH: 80px" onclick="SubmitPrompts()" />
				<input type="button" class="btn" name="Cancel" value="Cancel" style="WIDTH: 80px" onclick="closepromptedclick()" />
			</div>	
		<% End If%>

	</div>
</div>

<script type="text/javascript">

	function SubmitTemplate() {
		var frmTemplateFile = $("#frmTemplateFile")[0];
		frmTemplateFile.submit();
	}
	
	function SubmitPrompts() {
		var frmPromptedValues = document.getElementById('frmPromptedValues');

		if ($('#TemplateFile').val() == "") {
			OpenHR.modalMessage("No template file selected");
			return;
		}

		// Validate the prompt values before submitting the form.
		var controlCollection = frmPromptedValues.elements;
		if (controlCollection != null) {
			for (var i = 0; i < controlCollection.length; i++) {
				var sControlName = controlCollection.item(i).name;
				var sControlPrefix = sControlName.substr(0, 7);

				if (sControlPrefix == "prompt_") {
					// Get the control's data type.
					var iType = new Number(sControlName.substring(7, 8));
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
		var sMessage;
		var fFound;
		var sMaskCtlName;
		var iIndex;
		var frmPromptedValues = document.getElementById('frmPromptedValues');
		
		sDecimalSeparator = frmPromptedValues.txtLocaleDecimalSeparator.value;
		sThousandSeparator = frmPromptedValues.txtLocaleThousandSeparator.value;

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
			sConvertedValue = OpenHR.replaceAll(sConvertedValue, sThousandSeparator, "");
			pctlPrompt.value = sConvertedValue;

			// Convert any decimal separators to '.'.
			if (sDecimalSeparator != ".") {
				// Remove decimal points.
				sConvertedValue = sConvertedValue.replace(rePoint, "A");
				// replace the locale decimal marker with the decimal point.
				sConvertedValue = OpenHR.replaceAll(sConvertedValue, sDecimalSeparator, ".");
			}

			if (isNaN(sConvertedValue) == true) {
				fOK = false;
				sMessage = "Invalid numeric value entered.";
			}
		}

		if ((fOK == true) && (piDataType == 4)) {

			fOK = OpenHR.IsValidDate(pctlPrompt.value);

			if (fOK == false) {
				sMessage = "Invalid date value entered.";
			}
		}

		if ((fOK == true) && (piDataType == 1)) {
			// Character column.
			// Ensure that the value entered matches the required mask (if there is one).
			sMaskCtlName = "promptMask_" + pctlPrompt.name.substring(9, pctlPrompt.name.length);

			fFound = false;
			var controlCollection = frmPromptedValues.elements;
			if (controlCollection != null) {
				for (var i = 0; i < controlCollection.length; i++) {
					if (controlCollection.item(i).name.toUpperCase() == sMaskCtlName) {
						fFound = true;
						break;
					}
				}
			}

			if (fFound == true) {
				var sMask = frmPromptedValues.elements(sMaskCtlName).value;
				sValue = pctlPrompt.value;
				// Need to get rid of the backslash characters that precede literals.
				// But remember that two backslashes give a literal backslash that does not want
				// to be got rid of.
				var sTemp = sMask.replace(reDoubleBackSlash, "a");
				sTemp = sTemp.replace(reBackSlash, "");
				if (sMask.length > 0) {
					if (sTemp.length != sValue.length) {
						fOK = false;
					}
					else {
						// Prompt values length matches mask length, so now check each character.
						var fFollowingBackslash = false;
						iIndex = 0;
						for (i = 0; i < sMask.length; i++) {
							var sValueChar = sValue.substring(iIndex, iIndex + 1);

							if (fFollowingBackslash == false) {
								switch (sMask.substring(i, i + 1)) {
								case "A":
									// Character must be uppercase.
									if (sValueChar.toUpperCase() != sValueChar) {
										fOK = false;
									}
									else {
										var iNumber = new Number(sValueChar);
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
					sMessage = "The entered value does not match the required format (" + sMask + ").";
				}
			}
		}

		if (fOK == false) {
			OpenHR.modalMessage(sMessage);
			window.focus();
			pctlPrompt.focus();
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

		var frmPromptedValues = document.getElementById('frmPromptedValues');
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
			var iValue = parseInt(sYears);
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
		var sSource = "prompt_3_" + piPromptID;
		var sDest = "promptChk_" + piPromptID;
		var frmPromptedValues = document.getElementById('frmPromptedValues');
		
		frmPromptedValues.elements.item(sDest).value = frmPromptedValues.elements.item(sSource).checked;
	}

	function comboChange(piPromptID) {
		var frmPromptedValues = document.getElementById('frmPromptedValues');
		
		var sSource = "promptLookup_" + piPromptID;
		var ctlSource = frmPromptedValues.elements.item(sSource);

		var controlCollection = frmPromptedValues.elements;
		if (controlCollection != null) {
			for (var i = 0; i < controlCollection.length; i++) {
				var sControlName = controlCollection.item(i).name;
				var sControlPrefix = sControlName.substr(0, 7);
				var sControlID = sControlName.substr(9, sControlName.length);

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

