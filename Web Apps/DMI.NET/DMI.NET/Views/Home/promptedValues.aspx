<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

<%@ Import Namespace="DMI.NET" %>

<!DOCTYPE html>

<% 	
	Session("filterID") = Request.Form("filterID")
%>

<html>
<head>
	
	<link href="<%: Url.Content("~/Content/OpenHR.css")%>" rel="stylesheet" />

	<script src="<%: Url.Content("~/Scripts/jquery/jquery-1.8.3.js")%>"></script>
	<script src="<%: Url.Content("~/Scripts/jquery/jquery-ui-1.9.2.custom.js")%>"></script>
	<script src="<%: Url.Content("~/Scripts/OpenHR.js")%>" type="text/javascript"></script>

	<title>OpenHR Intranet</title>

	<script type="text/javascript">
		function promptedValues_onload() {
			var frmPromptedValues = document.getElementById("frmPromptedValues");
		
			if (frmPromptedValues.txtPromptCount.value == 0) {
				OpenHR.submitForm(frmPromptedValues);
			} else {
				// Set focus on the first prompt control.
				var controlCollection = frmPromptedValues.elements;
				if (controlCollection != null) {
					var sControlName, sControlPrefix;
					for (var i = 0; i < controlCollection.length; i++) {
						sControlName = controlCollection.item(i).name;
						sControlPrefix = sControlName.substr(0, 7);

						if ((sControlPrefix == "prompt_") || (sControlName.substr(0, 13) == "promptLookup_")) {
							controlCollection.item(i).focus();
							break;
						}
					}
				}

				// Resize the grid to show all prompted values.
				var iResizeBy = frmPromptedValues.offsetParent.scrollHeight - frmPromptedValues.offsetParent.clientHeight;
				if (frmPromptedValues.offsetParent.offsetHeight + iResizeBy > screen.height) {
					window.parent.dialogHeight = new String(screen.height) + "px";
				} else {
					var iNewHeight = new Number(window.parent.dialogHeight.substr(0, window.parent.dialogHeight.length - 2));
					iNewHeight = iNewHeight + iResizeBy;
					window.parent.dialogHeight = new String(iNewHeight) + "px";
				}
			}
		}
	</script>

	<script type="text/javascript">
		function SubmitPrompts()
		{
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
			OpenHR.submitForm(frmPromptedValues);
		}

		function CancelClick()
		{
			window.parent.self.close();
		}

		function ValidatePrompt(pctlPrompt, piDataType)
		{
			// Validate the given prompt value.
			var fOK;
			var reBackSlash = new RegExp("\\\\", "gi");
			var reDoubleBackSlash = new RegExp("\\\\\\\\", "gi");
			var sDecimalSeparator;
			var sThousandSeparator;
			var sPoint;

			sDecimalSeparator = "\\";
			sDecimalSeparator = sDecimalSeparator.concat(OpenHR.LocaleDecimalSeparator);
			var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");

			sThousandSeparator = "\\";
			sThousandSeparator = sThousandSeparator.concat(OpenHR.LocaleThousandSeparator);
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
					OpenHR.MessageBox("Invalid numeric value entered.");
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
					sValue = convertLocaleDateToSQL(sValue);
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
				sMaskCtlName = sMaskCtlName.toUpperCase();

				var fFound = false;		
				var controlCollection = frmPromptedValues.elements;
				if (controlCollection!=null) {
					for (var i=0; i<controlCollection.length; i++)  {
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

		function convertLocaleDateToSQL(psDateString)
		{ 
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
		
			sDateFormat = OpenHR.LocaleDateFormat;

			sDays="";
			sMonths="";
			sYears="";
			iValuePos = 0;

			// Trim leading spaces.
			sTempValue = psDateString.substr(iValuePos,1);
			while (sTempValue.charAt(0) == " ") 
			{
				iValuePos = iValuePos + 1;		
				sTempValue = psDateString.substr(iValuePos,1);
			}

			for (iLoop=0; iLoop<sDateFormat.length; iLoop++)  {
				if ((sDateFormat.substr(iLoop,1).toUpperCase() == 'D') && (sDays.length==0)){
					sDays = psDateString.substr(iValuePos,1);
					iValuePos = iValuePos + 1;
					sTempValue = psDateString.substr(iValuePos,1);

					if (isNaN(sTempValue) == false) {
						sDays = sDays.concat(sTempValue);			
					}
					iValuePos = iValuePos + 1;		
				}

				if ((sDateFormat.substr(iLoop,1).toUpperCase() == 'M') && (sMonths.length==0)){
					sMonths = psDateString.substr(iValuePos,1);
					iValuePos = iValuePos + 1;
					sTempValue = psDateString.substr(iValuePos,1);

					if (isNaN(sTempValue) == false) {
						sMonths = sMonths.concat(sTempValue);			
					}
					iValuePos = iValuePos + 1;
				}

				if ((sDateFormat.substr(iLoop,1).toUpperCase() == 'Y') && (sYears.length==0)){
					sYears = psDateString.substr(iValuePos,1);
					iValuePos = iValuePos + 1;
					sTempValue = psDateString.substr(iValuePos,1);

					if (isNaN(sTempValue) == false) {
						sYears = sYears.concat(sTempValue);			
					}
					iValuePos = iValuePos + 1;
					sTempValue = psDateString.substr(iValuePos,1);

					if (isNaN(sTempValue) == false) {
						sYears = sYears.concat(sTempValue);			
					}
					iValuePos = iValuePos + 1;
					sTempValue = psDateString.substr(iValuePos,1);

					if (isNaN(sTempValue) == false) {
						sYears = sYears.concat(sTempValue);			
					}
					iValuePos = iValuePos + 1;
				}

				// Skip non-numerics
				sTempValue = psDateString.substr(iValuePos,1);
				while (isNaN(sTempValue) == true) {
					iValuePos = iValuePos + 1;		
					sTempValue = psDateString.substr(iValuePos,1);
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

	<script src="<%: Url.Content("~/Scripts/ctl_SetStyles.js") %>" type="text/javascript"></script>

</head>

<body <%=session("BodyColour")%> leftmargin="20" topmargin="20" bottommargin="20" rightmargin="5">
	
	<form name="frmPromptedValues" id="frmPromptedValues" method="POST" action="promptedValues_submit">

		<%		
			Dim iPromptCount As Long
			Dim fDefaultFound As Boolean
			Dim fFirstValueDone As Boolean
			Dim sFirstValue As String
			Dim sDefaultValue As String
			
			Dim cmdDefn = CreateObject("ADODB.Command")
			cmdDefn.CommandText = "sp_ASRIntGetFilterPromptedValuesRecordset"
			cmdDefn.CommandType = 4	' Stored Procedure
			cmdDefn.ActiveConnection = Session("databaseConnection")

			Dim prmFilterID = cmdDefn.CreateParameter("filterID", 3, 1)	' 3=integer, 1=input
			cmdDefn.Parameters.Append(prmFilterID)
			prmFilterID.value = CleanNumeric(Session("filterID"))

			Err.Clear()
			Dim rstPromptedValue = cmdDefn.Execute

			If Not (rstPromptedValue.EOF And rstPromptedValue.BOF) Then
		%>
		<table align="center" class="outline" cellpadding="5" cellspacing="0">
			<tr>
				<td>
					<table align="center" class="invisible" cellspacing="0" cellpadding="0">
						<tr>
							<td colspan="5" align="center">
								<h3 align="center">Prompted Values</h3>
							</td>
						</tr>
						<%
							Do While Not rstPromptedValue.EOF
								iPromptCount = iPromptCount + 1
						%>
						<tr height="10">
							<td width="20" height="10"></td>
							<td nowrap height="10">
								<%
									If NullSafeString(rstPromptedValue.fields("ValueType").value) = "3" Then
								%>
								<label
									for="prompt_3_<%=rstPromptedValue.fields("componentID").value%>"
									class="checkbox"
									tabindex="0"
									onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
									onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
									onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
									onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
									onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
									<%
									End If
									%>

									<%=rstPromptedValue.fields("PromptDescription").value%>
									<%
										If NullSafeString(rstPromptedValue.fields("ValueType").value) = "3" Then
									%>
								</label>
								<%
								End If
								%>
							</td>
							<td width="20" height="10">&nbsp;</td>
							<td width="200" height="10">
								<%
									' Character Prompted Value
									If NullSafeString(rstPromptedValue.fields("ValueType").value) = "1" Then
								%>
								<input type="text" class="text" id='prompt_1_<%=rstPromptedValue.fields("componentID").value%>' name='prompt_1_<%=rstPromptedValue.fields("componentID").value%>'
									value="<%=replace(rstpromptedvalue.fields("valuecharacter"), """", "&quot;")%>" maxlength='<%=rstPromptedValue.fields("promptsize").value%>'
									style="WIDTH: 100%">
								<input type="hidden" id='promptMask_<%=rstPromptedValue.fields("componentID").value%>' name='promptMask_<%=rstPromptedValue.fields("componentID").value%>'
									value="<%=Replace(rstPromptedValue.fields("promptMask").value, """", "&quot;")%>">
								<%
									' Numeric Prompted Value
								ElseIf NullSafeString(rstPromptedValue.fields("ValueType").value) = "2" Then
								%>
								<input type="text" class="text" id='prompt_2_<%=rstPromptedValue.fields("componentID").value%>' name='prompt_2_<%=rstPromptedValue.fields("componentID").value%>'
									value="<%=Replace(rstPromptedValue.fields("valuenumeric").value, ".", Session("LocaleDecimalSeparator"))%>"
									style="WIDTH: 100%">
								<input type="hidden" id='promptSize_<%=rstPromptedValue.fields("componentID").value%>' name='promptSize<%=rstPromptedValue.fields("componentID").value%>'
									value="<%=rstPromptedValue.fields("promptSize").value%>">
								<input type="hidden" id='promptDecs_<%=rstPromptedValue.fields("componentID").value%>' name='promptDecs<%=rstPromptedValue.fields("componentID").value%>'
									value="<%=rstPromptedValue.fields("promptDecimals").value%>">
								<%
									' Logic Prompted Value
								ElseIf NullSafeString(rstPromptedValue.fields("ValueType").value) = "3" Then
								%>
								<input type="checkbox" id='prompt_3_<%=rstPromptedValue.fields("componentID").value%>' name='prompt_3_<%=rstPromptedValue.fields("componentID").value%>'
									<%If rstPromptedValue.fields("valuelogic").value Then%> checked <%End If%>
									onclick="checkboxClick(<%=rstPromptedValue.fields("componentID").value%>)"
									onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
									onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
								<input type="hidden" id='promptChk_<%=rstPromptedValue.fields("componentID").value%>' name='promptChk_<%=rstPromptedValue.fields("componentID").value%>'
									value='<%=rstPromptedValue.fields("valuelogic").value%>'>
								<%			  
									' Date Prompted Value
								ElseIf NullSafeString(rstPromptedValue.fields("ValueType").value) = "4" Then

									Response.Write("        <input type=text class=""text"" id=prompt_4_" & rstPromptedValue.fields("componentID").value & " name=prompt_4_" & rstPromptedValue.fields("componentID").value & " value=""")
	
									Dim iDay As Integer, iMonth As Integer, dtDate As DateTime
	
									Select Case rstPromptedValue.fields("promptDateType").value
										Case 0
											' Explicit value
											Response.Write(ConvertSqlDateToLocale(rstPromptedValue.fields("valuedate").value))
										Case 1
											' Current date
											Response.Write(ConvertSqlDateToLocale(Now()))
										Case 2
											' Start of current month
											iDay = (Day(Now()) * -1) + 1
											dtDate = DateAdd("d", iDay, Now())
											Response.Write(ConvertSqlDateToLocale(dtDate))
										Case 3
											' End of current month
											iDay = (Day(Now()) * -1) + 1
											dtDate = DateAdd("d", iDay, Now())
											dtDate = DateAdd("m", 1, dtDate)
											dtDate = DateAdd("d", -1, dtDate)
											Response.Write(ConvertSqlDateToLocale(dtDate))
										Case 4
											' Start of current year
											iDay = (Day(Now()) * -1) + 1
											iMonth = (Month(Now()) * -1) + 1
											dtDate = DateAdd("d", iDay, Now())
											dtDate = DateAdd("m", iMonth, dtDate)
											Response.Write(ConvertSqlDateToLocale(dtDate))
										Case 5
											' End of current year
											iDay = (Day(Now()) * -1) + 1
											iMonth = (Month(Now()) * -1) + 1
											dtDate = DateAdd("d", iDay, Now())
											dtDate = DateAdd("m", iMonth, dtDate)
											dtDate = DateAdd("yyyy", 1, dtDate)
											dtDate = DateAdd("d", -1, dtDate)
											Response.Write(ConvertSqlDateToLocale(dtDate))
									End Select
									Response.Write(""" style=""WIDTH: 100%"">" & vbCrLf)

									' Lookup Prompted Value
								ElseIf NullSafeString(rstPromptedValue.fields("ValueType").value) = "5" Then
									Response.Write("        		<SELECT id=promptLookup_" & rstPromptedValue.fields("componentID").value & " name=promptLookup_" & rstPromptedValue.fields("componentID").value & " style=""WIDTH: 100%"" class=""combo"" onchange=""comboChange(" & rstPromptedValue.fields("componentID").value & ")"">" & vbCrLf)

									fDefaultFound = False
									fFirstValueDone = False
									sFirstValue = ""
					
									' Get the lookup values.
									Dim cmdLookupValues = CreateObject("ADODB.Command")
									cmdLookupValues.CommandText = "sp_ASRIntGetLookupValues"
									cmdLookupValues.CommandType = 4	' Stored Procedure
									cmdLookupValues.ActiveConnection = Session("databaseConnection")

									Dim prmColumnID = cmdLookupValues.CreateParameter("columnID", 3, 1)
									cmdLookupValues.Parameters.Append(prmColumnID)
									prmColumnID.value = CleanNumeric(rstPromptedValue.fields("fieldColumnID").value)

									Err.Clear()
									Dim rstLookupValues = cmdLookupValues.Execute

									Do While Not rstLookupValues.EOF
										Response.Write("        		  <OPTION")
					
										If Not fFirstValueDone Then
											sFirstValue = rstLookupValues.Fields(0).Value
											fFirstValueDone = True
										End If
						
										If rstLookupValues.fields(0).type = 135 Then
											' Field is a date so format as such.
											Dim sOptionValue = ConvertSqlDateToLocale(rstLookupValues.Fields(0).Value)
											If sOptionValue = ConvertSqlDateToLocale(rstPromptedValue.fields("valuecharacter").value) Then
												Response.Write(" SELECTED")
												fDefaultFound = True
											End If
											Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
										ElseIf rstLookupValues.fields(0).type = 131 Then
											' Field is a numeric so format as such.
											Dim sOptionValue = Replace(rstLookupValues.Fields(0).Value, ".", Session("LocaleDecimalSeparator"))
											If (Not IsDBNull(rstLookupValues.Fields(0).Value)) And (Not IsDBNull(rstPromptedValue.fields("valuecharacter").value)) Then
												If FormatNumber(rstLookupValues.Fields(0).Value) = FormatNumber(rstPromptedValue.fields("valuecharacter").value) Then
													Response.Write(" SELECTED")
													fDefaultFound = True
												End If
											End If
											Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
										ElseIf rstLookupValues.fields(0).type = 11 Then
											' Field is a logic so format as such.
											Dim sOptionValue = rstLookupValues.Fields(0).Value
											If sOptionValue = rstPromptedValue.fields("valuecharacter").value Then
												Response.Write(" SELECTED")
												fDefaultFound = True
											End If
											Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
										Else
											Dim sOptionValue = rstLookupValues.Fields(0).Value
											If sOptionValue = rstPromptedValue.fields("valuecharacter").value Then
												Response.Write(" SELECTED")
												fDefaultFound = True
											End If
											Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
										End If

										rstLookupValues.MoveNext()
									Loop

									Response.Write("   		     </SELECT>" & vbCrLf)

									If fDefaultFound Then
										sDefaultValue = rstPromptedValue.Fields("valuecharacter").value
									Else
										sDefaultValue = sFirstValue
									End If
					
									If rstLookupValues.fields(0).type = 135 Then
										' Date.
										Response.Write("        <input type=hidden id=prompt_4_" & rstPromptedValue.fields("componentID").value & " name=prompt_4_" & rstPromptedValue.fields("componentID").value & " value=" & ConvertSqlDateToLocale(sDefaultValue) & ">" & vbCrLf)
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
								Response.Write("					<td width=20 height=10></td>" & vbCrLf)
								Response.Write("				</tr>" & vbCrLf)

								rstPromptedValue.MoveNext()
							Loop

							Response.Write("				<tr>" & vbCrLf)
							Response.Write("					<td colspan=5 height=10>&nbsp;</td>" & vbCrLf)
							Response.Write("				</tr>" & vbCrLf)
							Response.Write("				<tr height=20>" & vbCrLf)
							Response.Write("					<td width=20></td>" & vbCrLf)
							Response.Write("					<td colspan=3>" & vbCrLf)
							Response.Write("						<TABLE WIDTH=100% class=""invisible"" CELLSPACING=0 CELLPADDING=0>" & vbCrLf)
							Response.Write("							<TD>&nbsp;</TD>" & vbCrLf)
			
							Response.Write("							<td width=80>" & vbCrLf)
								%>
								<input type="button" name="Submit" value="OK" style="WIDTH: 80px" class="btn"
									onclick="SubmitPrompts()"
									onmouseover="try{button_onMouseOver(this);}catch(e){}"
									onmouseout="try{button_onMouseOut(this);}catch(e){}"
									onfocus="try{button_onFocus(this);}catch(e){}"
									onblur="try{button_onBlur(this);}catch(e){}" />
							</td>
							<td width="20"></td>
							<td width="80">
								<input type="button" class="btn" name="Cancel" value="Cancel" style="WIDTH: 80px"
									onclick="CancelClick()"
									onmouseover="try{button_onMouseOver(this);}catch(e){}"
									onmouseout="try{button_onMouseOut(this);}catch(e){}"
									onfocus="try{button_onFocus(this);}catch(e){}"
									onblur="try{button_onBlur(this);}catch(e){}" />
								<%
									Response.Write("							</td>" & vbCrLf)
									Response.Write("						</table>" & vbCrLf)
									Response.Write("					</td>" & vbCrLf)
									Response.Write("					<td width=20></td>" & vbCrLf)
									Response.Write("				</tr>" & vbCrLf)
									Response.Write("				<tr>" & vbCrLf)
									Response.Write("					<td colspan=5 height=5></td>" & vbCrLf)
									Response.Write("				</tr>" & vbCrLf)
									Response.Write("			</table>" & vbCrLf)
									Response.Write("		</td>" & vbCrLf)
									Response.Write("	</tr>" & vbCrLf)
									Response.Write("</table>" & vbCrLf)
								End If
		
								rstPromptedValue.close()
								rstPromptedValue = Nothing

								Response.Write("<input type=""hidden"" id=""txtPromptCount"" name=""txtPromptCount"" value=" & iPromptCount & ">" & vbCrLf)
								%>

								<input type="hidden" id="filterID" name="filterID" value="<%=Session("filterID")%>">
	</form>
</body>
</html>


