<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<% 	
	Session("filterID") = Request.Form("filterID")
%>

<script type="text/javascript">
	function promptedValues_onload() {
		
		var frmPromptedValues = document.getElementById("frmPromptedValues");
		
		if (frmPromptedValues.txtPromptCount.value == 0) {
			if ($('#tmpDialog').dialog('isOpen') == true) {
				OpenHR.submitForm(frmPromptedValues, 'tmpDialog');
			} else {
				OpenHR.submitForm(frmPromptedValues);
			}
		} else {

			$('#frmPromptedValues *[id^="prompt_"]:not([type="hidden"]), #frmPromptedValues *[id^="promptLookup_"]:not([type="hidden"])').first().focus();

			if ($('.popup').dialog('isOpen')) {
				var dialogWidth = screen.width / 3;

				$('.popup').dialog('option', 'height', 'auto');
				$('.popup').dialog('option', 'width', dialogWidth);
			}

			//create any jquery datepickers, but keep them closed.
			$('[data-type="date"]').datepicker();
			if ($('input[data-type="date"]').length > 0) setTimeout("hideDatePickers()", 100);
		}
	}

	function hideDatePickers() {
		$('input[data-type="date"]').datepicker('hide');
		if ($('#frmPromptedValues *[id^="prompt_"]:not([type="hidden"]), #frmPromptedValues *[id^="promptLookup_"]:not([type="hidden"])').first().hasClass('hasDatepicker') == true) $('#pv_cancel').focus();
	}

</script>

<script type="text/javascript">
				
	function SubmitPrompts() {
			
		// Validate the prompt values before submitting the form.
		var frmPromptedValues = document.getElementById("frmPromptedValues");
		var controlCollection = frmPromptedValues.elements;
		if (controlCollection!=null) {
			var sControlName, sControlPrefix;
			for (var i=0; i<controlCollection.length; i++)  {
				sControlName = controlCollection.item(i).name;
				sControlPrefix = sControlName.substr(0, 7);
	
				if (sControlPrefix=="prompt_") {

					// Get the control's data type.
					var iType = new Number(sControlName.substring(7,8));
					if ((iType==1) || (iType==2) || (iType==4)) {
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
		if ($('#tmpDialog').dialog('isOpen') == true) {
			//prompted Values for OpenHR.modalExpressionSelect screen.
			OpenHR.submitForm(frmPromptedValues, 'tmpDialog');
		} else {
			OpenHR.submitForm(frmPromptedValues);
		}
			
	}

	function pv_cancelClick() {			
		if($('#tmpDialog').dialog('isOpen') == true) {
			$("#tmpDialog").dialog('close');
		}
		else {
			if ($('.popup').dialog('isOpen')) {
				$(".popup").dialog('close');
			}
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

		sDecimalSeparator = "\\";
		sDecimalSeparator = sDecimalSeparator.concat(OpenHR.LocaleDecimalSeparator());
		var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");

		sThousandSeparator = "\\";
		sThousandSeparator = sThousandSeparator.concat(OpenHR.LocaleThousandSeparator());
		var reThousandSeparator = new RegExp(sThousandSeparator, "gi");

		sPoint = "\\.";
		var rePoint = new RegExp(sPoint, "gi");

		fOK = true;

		if ((fOK == true) && (piDataType == 2)) {
			// Numeric column.
			// Ensure that the value entered is numeric.
			var sValue = pctlPrompt.value;

			if (sValue.length == 0) {
				sValue = "0";
				pctlPrompt.value = 0;
			}

			// Convert the value from locale to UK settings for use with the isNaN funtion.
			var sConvertedValue = new String(sValue);
			// Remove any thousand separators.
			sConvertedValue = sConvertedValue.replace(reThousandSeparator, "");

			// Convert any decimal separators to '.'.
			if (OpenHR.LocaleDecimalSeparator != ".") {
				// Remove decimal points.
				sConvertedValue = sConvertedValue.replace(rePoint, "A");
				// replace the locale decimal marker with the decimal point.
				sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
			}

			if (isNaN(sConvertedValue) == true) {
				fOK = false;
				OpenHR.messageBox("Invalid numeric value entered.");
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

				sValue =  OpenHR.convertLocaleDateToSQL(sValue);
				if (sValue.length == 0) {
					fOK = false;
				}
				else {
					pctlPrompt.value = OpenHR.ConvertSQLDateToLocale(sValue);
				}
			}
			
			if (fOK == false) {
				OpenHR.messageBox("Invalid date value entered.");
				pctlPrompt.focus();
			}	
		}
	
		if ((fOK == true) && (piDataType == 1)) {
			// Character column.
			// Ensure that the value entered matches the required mask (if there is one).
			var frmPromptedValues = document.getElementById("frmPromptedValues");
			
			var sMaskCtlName = "promptMask_" + pctlPrompt.name.substring(9, pctlPrompt.name.length);

			var fFound = false;		
			var controlCollection = frmPromptedValues.elements;
			if (controlCollection!=null) {
				for (var i=0; i<controlCollection.length; i++)  {
					if (controlCollection.item(i).name == sMaskCtlName) {
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
						var iIndex = 0;
						for (i=0; i<sMask.length; i++)  {
							var sValueChar = sValue.substring(iIndex, iIndex+1);
						
							if (fFollowingBackslash == false) {
								switch (sMask.substring(i, i+1)) {
									case "A":
										// Character must be uppercase.
										if (sValueChar.toUpperCase() != sValueChar) {
											fOK = false;
										}
										else {
											var iNumber = new Number(sValueChar);
											if (isNaN(iNumber) == false) {
												fOK= true;
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
												fOK= false;
											}
										}
										iIndex = iIndex + 1;
										break;
									case "9":
										// Character must be numeric (0-9).
										iNumber = new Number(sValueChar);
										if (isNaN(iNumber) == true) {
											fOK= false;
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
											fOK= false;
										}
										iIndex = iIndex + 1;
										break;
									case "B":
										// Character must be logic (0 or 1).
										if ((sValueChar != "0") &&
											(sValueChar != "1")){
											fOK= false;
										}
										iIndex = iIndex + 1;
										break;
									case "\\":
										// Following character is literal.
										fFollowingBackslash = true;
										break;
									default:
										// Literal.
										if (sMask.substring(i, i+1) != sValueChar) {
											fOK = false;
										}
										iIndex = iIndex + 1;
								}
							}
							else {
								fFollowingBackslash = false;
								if (sMask.substring(i, i+1) != sValueChar) {
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
					pctlPrompt.focus();
				}	
			}
		}

		return fOK;
	}

	function checkboxClick(piPromptID) {
		var sSource = "prompt_3_" + piPromptID;
		var sDest = "promptChk_" + piPromptID;
		var frmPromptedValues = document.getElementById("frmPromptedValues");
	
		frmPromptedValues.elements.item(sDest).value = frmPromptedValues.elements.item(sSource).checked;
	}

	function comboChange(piPromptID) {
		var frmPromptedValues = document.getElementById("frmPromptedValues");
		var sSource = "promptLookup_" + piPromptID;
		var ctlSource = frmPromptedValues.elements.item(sSource);
	
		var controlCollection = frmPromptedValues.elements;
		if (controlCollection!=null) {
			var sControlName, sControlPrefix, sControlID;
			for (var i=0; i<controlCollection.length; i++)  {
				sControlName = controlCollection.item(i).name;
				sControlPrefix = sControlName.substr(0, 7);
				sControlID = sControlName.substr(9, sControlName.length);
	
				if ((sControlPrefix=="prompt_") && (sControlID == piPromptID)) {
					controlCollection.item(i).value = ctlSource.options[ctlSource.selectedIndex].text;
				}
			}
		}
	}
</script>

<script src="<%: Url.LatestContent("~/Scripts/ctl_SetStyles.js")%>" type="text/javascript"></script>


<form name="frmPromptedValues" id="frmPromptedValues" method="POST" action="<%:Url.Action("promptedValues_Submit")%>">

	<%		
		Dim objSession As SessionInfo = CType(Session("SessionContext"), SessionInfo)	'Set session info
		Dim objDataAccess As New clsDataAccess(objSession.LoginInfo) 'Instantiate DataAccess class
		Dim iPromptCount As Long
		Dim fDefaultFound As Boolean
		Dim fFirstValueDone As Boolean
		Dim sFirstValue As String
		Dim sDefaultValue As String
			
		Dim rstPromptedValue As DataTable = objDataAccess.GetDataTable( _
						"sp_ASRIntGetFilterPromptedValuesRecordset", _
						CommandType.StoredProcedure, _
						New SqlParameter("@piFilterID", SqlDbType.Int) With {.Value = CleanNumeric(Session("filterID"))} _
		)
			
		If rstPromptedValue.Rows.Count > 0 Then
	
			Response.Write("<h3>Prompted Values</h3>" & vbCrLf)
			
			For Each rowPromptedValues As DataRow In rstPromptedValue.Rows
				iPromptCount = iPromptCount + 1
				
				Dim componentID As String = rowPromptedValues("componentID").ToString
				Dim valueType As Integer = NullSafeInteger(rowPromptedValues("ValueType"))
				
				Response.Write("<div class='formField'>" & vbCrLf)
								
				If valueType = ExpressionValueTypes.giEXPRVALUE_LOGIC Then
					Response.Write(String.Format("<label for='prompt_3_{0}' class='checkbox' tabindex='0' style='width: 40%;'>", componentID) & vbCrLf)
				Else
					Response.Write("<label style='width:40%;'>")
				End If
				
				Response.Write(HttpUtility.HtmlEncode(rowPromptedValues("PromptDescription")))
				Response.Write("</label>")
								
				Select Case valueType
					Case ExpressionValueTypes.giEXPRVALUE_CHARACTER	' Character Prompted Value
						Response.Write(String.Format("<input type='text' id='prompt_1_{0}' name='prompt_1_{0}' value='{1}' maxlength={2} style='width: 58%;'>", componentID, Html.Encode(rowPromptedValues("valuecharacter").ToString), rowPromptedValues("promptsize").ToString) & vbCrLf)
						Response.Write(String.Format("<input type='hidden' id='promptMask_{0}' name='promptMask_{0}' value='{1}'>", componentID, Html.Encode(rowPromptedValues("promptMask").ToString)) & vbCrLf)
					
					Case ExpressionValueTypes.giEXPRVALUE_NUMERIC	' Numeric Prompted Value
						Response.Write(String.Format("<input type='text' id='prompt_2_{0}' name='prompt_2_{0}' value='{1}' style=""width: 58%;"">", componentID, Replace(rowPromptedValues("valuenumeric").ToString, ".", Session("LocaleDecimalSeparator").ToString)) & vbCrLf)
						Response.Write(String.Format("<input type='hidden' id='promptSize_{0}' name='promptSize{0}' value='{1}'>", componentID, rowPromptedValues("promptSize").ToString) & vbCrLf)
						Response.Write(String.Format("<input type='hidden' id='promptDecs_{0}' name='promptDecs{0}' value='{1}'>", componentID, rowPromptedValues("promptDecimals").ToString) & vbCrLf)
													 
					Case ExpressionValueTypes.giEXPRVALUE_LOGIC	' Logic Prompted Value
						Response.Write(String.Format("<input type='checkbox' id='prompt_3_{0}' name='prompt_3_{0}' {1} onclick='checkboxClick({0})' style='width: 1em;'/>", componentID, IIf(CBool(rowPromptedValues("valuelogic")), "checked", "")))
						Response.Write(String.Format("<input type='hidden' id='promptChk_{0}' name='promptChk_{0}' value='{1}'>", componentID, rowPromptedValues("valuelogic").ToString) & vbCrLf)
						
					Case ExpressionValueTypes.giEXPRVALUE_DATE	' Date Prompted Value
						
						Dim iDay As Integer, iMonth As Integer, dtDate As DateTime
						Dim dateString As String = ""
						
						Select Case rowPromptedValues("promptDateType")
							Case PromptedDateType.Explicit
								' Explicit value
								'If the explicit value is 1899-12-30 00:00:00.000 we need to display an empty date
								Dim is1899 As Boolean = rowPromptedValues("valuedate").Year.ToString() = "1899"
								If is1899 Then
									dateString = ""
								Else 'Display the explicit value coming down from the database
									dateString = ConvertSQLDateToLocale(rowPromptedValues("valuedate").ToString)
								End If
							Case PromptedDateType.Current
								' Current date
								dateString = ConvertSQLDateToLocale(Now())
							Case PromptedDateType.MonthStart
								' Start of current month
								iDay = (Day(Now()) * -1) + 1
								dtDate = DateAdd("d", iDay, Now())
								dateString = ConvertSQLDateToLocale(dtDate)
							Case PromptedDateType.MonthEnd
								' End of current month
								iDay = (Day(Now()) * -1) + 1
								dtDate = DateAdd("d", iDay, Now())
								dtDate = DateAdd("m", 1, dtDate)
								dtDate = DateAdd("d", -1, dtDate)
								dateString = ConvertSQLDateToLocale(dtDate)
							Case PromptedDateType.YearStart
								' Start of current year
								iDay = (Day(Now()) * -1) + 1
								iMonth = (Month(Now()) * -1) + 1
								dtDate = DateAdd("d", iDay, Now())
								dtDate = DateAdd("m", iMonth, dtDate)
								dateString = ConvertSQLDateToLocale(dtDate)
							Case PromptedDateType.YearEnd
								' End of current year
								iDay = (Day(Now()) * -1) + 1
								iMonth = (Month(Now()) * -1) + 1
								dtDate = DateAdd("d", iDay, Now())
								dtDate = DateAdd("m", iMonth, dtDate)
								dtDate = DateAdd("yyyy", 1, dtDate)
								dtDate = DateAdd("d", -1, dtDate)
								dateString = ConvertSQLDateToLocale(dtDate)
						End Select
						
						Response.Write(String.Format("<input type='text' data-type='date' id='prompt_4_{0}' name='prompt_4_{0}' value='{1}' style='width: 58%;'>", componentID, dateString) & vbCrLf)
						
					Case ExpressionValueTypes.giEXPRVALUE_TABLEVALUE
						Response.Write(String.Format("<select id='promptLookup_{0}' name='promptLookup_{0}' style='width: 58%;' class='combo' onchange='comboChange({0})'>", componentID) & vbCrLf)

						fDefaultFound = False
						fFirstValueDone = False
						sFirstValue = ""
					
						Dim rstLookupValues = GetLookupValues(CInt(CleanNumeric(rowPromptedValues("fieldColumnID").ToString)))

						For Each rowLookupValues As DataRow In rstLookupValues.Rows
							Response.Write("<option")
					
							If Not fFirstValueDone Then
								sFirstValue = rowLookupValues(0).ToString
								fFirstValueDone = True
							End If
						
							If rstLookupValues.Columns(0).DataType = GetType(DateTime) Then
								' Field is a date so format as such.
								Dim sOptionValue = ConvertSQLDateToLocale(rowLookupValues(0).ToString)
								If sOptionValue = ConvertSQLDateToLocale(rowPromptedValues("valuecharacter").ToString) Then
									Response.Write(" selected")
									fDefaultFound = True
								End If
								Response.Write(">" & sOptionValue & "</option>" & vbCrLf)
							ElseIf IsDataColumnDecimal(rstLookupValues.Columns(0)) Then
								' Field is a numeric so format as such.
								Dim sOptionValue = Replace(rowLookupValues(0).ToString, ".", Session("LocaleDecimalSeparator").ToString)
								If (Not IsDBNull(rstLookupValues(0))) And (Not IsDBNull(rowPromptedValues("valuecharacter").ToString)) Then
									If FormatNumber(rowLookupValues(0)) = FormatNumber(rowPromptedValues("valuecharacter").ToString) Then
										Response.Write(" selected")
										fDefaultFound = True
									End If
								End If
								Response.Write(">" & sOptionValue & "</option>" & vbCrLf)
							ElseIf rstLookupValues.Columns(0).DataType = GetType(System.Boolean) Then
								' Field is a logic so format as such.
								Dim sOptionValue As String = rowLookupValues(0).ToString
								If sOptionValue = rowPromptedValues("valuecharacter").ToString Then
									Response.Write(" selected")
									fDefaultFound = True
								End If
								Response.Write(">" & sOptionValue & "</option>" & vbCrLf)
							Else
								Dim sOptionValue As String = rowLookupValues(0).ToString
								If sOptionValue = rowPromptedValues("valuecharacter").ToString Then
									Response.Write(" selected")
									fDefaultFound = True
								End If
								Response.Write(">" & sOptionValue & "</option>" & vbCrLf)
							End If
						Next

						Response.Write("</select>" & vbCrLf)

						If fDefaultFound Then
							sDefaultValue = rowPromptedValues("valuecharacter").ToString
						Else
							sDefaultValue = sFirstValue
						End If
					
						If rstLookupValues.Columns(0).DataType = GetType(DateTime) Then
							' Date.
							Response.Write(String.Format("<input type='hidden' id='prompt_4_{0}' name='prompt_4_{0}' value='{1}'>", componentID, ConvertSQLDateToLocale(sDefaultValue)) & vbCrLf)
						ElseIf IsDataColumnDecimal(rstLookupValues.Columns(0)) Then
							' Numeric
							Response.Write(String.Format("<input type='hidden' id='prompt_2_{0}' name='prompt_2_{0}' value='{1}'>", componentID, Replace(sDefaultValue, ".", Session("LocaleDecimalSeparator").ToString)) & vbCrLf)
						ElseIf rstLookupValues.Columns(0).DataType = GetType(System.Boolean) Then
							' Logic
							Response.Write(String.Format("<input type='hidden' id='prompt_3_{0}' name='prompt_3_{0}' value='{1}'>", componentID, sDefaultValue) & vbCrLf)
						Else
							Response.Write(String.Format("<input type=hidden id=prompt_1_{0} name=prompt_1_{0} value='{1}'>", componentID, Html.Encode(sDefaultValue)) & vbCrLf)
						End If

				End Select
				
				Response.Write("</div>")

			Next
			
			Response.Write("<br/><br/>" & vbCrLf)
			Response.Write("<input type='button' class='btn' id='pv_cancel' name='Cancel' value='Cancel' style='width: 80px; float: right;' onclick='pv_cancelClick()' />" & vbCrLf)
			Response.Write("<input type='button' name='Submit' value='OK' style='width: 80px; float: right; margin-right: 10px;' class='btn' onclick='SubmitPrompts()' />" & vbCrLf)
			
		End If
		Response.Write(String.Format("<input type='hidden' id='txtPromptCount' name='txtPromptCount' value='{0}'>", iPromptCount) & vbCrLf)
		
		Response.Write(String.Format("<input type='hidden' id='filterID' name='filterID' value='{0}'>", Session("filterID")) & vbCrLf)
		
	%>
	
	<%=Html.AntiForgeryToken()%>
</form>

<script type="text/javascript">promptedValues_onload();</script>
