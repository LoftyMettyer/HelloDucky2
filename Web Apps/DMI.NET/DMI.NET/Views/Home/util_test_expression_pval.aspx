<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage"%>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="ADODB" %>

<!DOCTYPE html>

<link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css" />

	<%--External script resources--%>
<script src="<%: Url.Content("~/bundles/OpenHR_General")%>" type="text/javascript"></script>

<script id="officebarscript" src="<%: Url.Content("~/Scripts/officebar/jquery.officebar.js") %>" type="text/javascript"></script>
<link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css" />
<link href="<%: Url.LatestContent("~/Content/Site.css")%>" rel="stylesheet" type="text/css" />
<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />
<link id="DMIthemeLink" href="<%: Url.LatestContent("~/Content/themes/" & Session("ui-theme").ToString() & "/jquery-ui.min.css")%>" rel="stylesheet" type="text/css" />
<link href="<%= Url.LatestContent("~/Content/general_enclosed_foundicons.css")%>" rel="stylesheet" type="text/css" />
<link href="<%= Url.LatestContent("~/Content/font-awesome.css")%>" rel="stylesheet" type="text/css" />
<link href="<%= Url.LatestContent("~/Content/fonts/SSI80v194934/style.css")%>" rel="stylesheet" />

<html>
<head runat="server">
		<title>OpenHR Intranet</title>

		<script type="text/javascript" >

		function util_test_expression_pval_onload() {

			// Prevent default behaviour of empty form
			$(".text").keydown(function(e) {
				if (e.keyCode == 13) {
					e.preventDefault();
				}
			});

				if (frmPromptedValues.txtPromptCount.value == 0) {
						OpenHR.submitForm(frmPromptedValues);
				}
				else {
						// Set focus on the first prompt control.
						var controlCollection = frmPromptedValues.elements;
						if (controlCollection!=null) {
								for (i=0; i<controlCollection.length; i++)  {
										sControlName = controlCollection.item(i).name;
										sControlPrefix = sControlName.substr(0, 7);
	
										if ((sControlPrefix=="prompt_") || (sControlName.substr(0, 13)=="promptLookup_")) {
												controlCollection.item(i).focus();
												break;
										}
								}
						}

						// Resize the grid to show all prompted values.
						var iResizeBy = frmPromptedValues.offsetParent.scrollHeight	- frmPromptedValues.offsetParent.clientHeight;
						if (frmPromptedValues.offsetParent.offsetHeight + iResizeBy > screen.height) {
								window.parent.dialogHeight = new String(screen.height) + "px";
						}
						else {
								var iNewHeight = new Number(window.parent.dialogHeight.substr(0, window.parent.dialogHeight.length-2));
								iNewHeight = iNewHeight + iResizeBy;
								window.parent.dialogHeight = new String(iNewHeight) + "px";
						}
				}
		}

		function SubmitPrompts() {



				// Validate the prompt values before submitting the form.
				var controlCollection = frmPromptedValues.elements;
				if (controlCollection!=null) {
						for (i=0; i<controlCollection.length; i++)  {
								var sControlName = controlCollection.item(i).name;
								var sControlPrefix = sControlName.substr(0, 7);

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

		function CancelClick() {
			self.close();
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
				var sConvertedValue;

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
						var sValue = pctlPrompt.value;
			
						if (sValue.length == 0) {
								fOK = false;
						}
						else {
								// Convert the date to SQL format (use this as a validation check).
								// An empty string is returned if the date is invalid.
								sValue = OpenHR.convertLocaleDateToSQL(sValue);
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
						sMaskCtlName = "promptMask_" + pctlPrompt.name.substring(9, pctlPrompt.name.length);
						sMaskCtlName = sMaskCtlName.toUpperCase();

						fFound = false;		
						var controlCollection = frmPromptedValues.elements;
						if (controlCollection!=null) {
								for (i=0; i<controlCollection.length; i++)  {
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
												for (i=0; i<sMask.length; i++)  {
														sValueChar = sValue.substring(iIndex, iIndex+1);
						
														if (fFollowingBackslash == false) {
																switch (sMask.substring(i, i+1)) {
																		case "A":
																				// Character must be uppercase.
																				if (sValueChar.toUpperCase() != sValueChar) {
																						fOK = false;
																				}
																				else {
																						iNumber = new Number(sValueChar);
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

			frmPromptedValues.elements.item(sDest).value = frmPromptedValues.elements.item(sSource).checked;
		}

		function comboChange(piPromptID) {
				var sSource = "promptLookup_" + piPromptID;
				var ctlSource = frmPromptedValues.elements.item(sSource);
	
				var controlCollection = frmPromptedValues.elements;
				if (controlCollection!=null) {
						for (i=0; i<controlCollection.length; i++)  {
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

</head>
<body>
		
		<div data-framesource="util_test_expression_pval">

			<form name="frmPromptedValues" id="frmPromptedValues" method="POST" action="util_test_expression">
<%
	Dim iPromptCount As Integer
	Dim sPrompts As String
	Dim sNodeKey As String
	Dim sPromptDescription As String
	Dim iValueType As Integer
	Dim iPromptSize As Integer
	Dim iPromptDecimals As Integer
	Dim sPromptMask As String
	Dim lngTableID As Long
	Dim lngColumnID As Long
	Dim sValueCharacter As String
	Dim dblValueNumeric As Double
	Dim fValueLogic As Boolean
	Dim dtValueDate As String
	Dim iPromptDateType As Integer
	Dim iCharIndex As Integer
	Dim iParameterIndex As Integer
	Dim sTemp As String
	Dim sFiltersAndCalcs As String
	Dim sFilterCalcID As String
	Dim fDefaultFound As Boolean
	Dim fFirstValueDone As Boolean
	Dim sFirstValue As String

	Dim cmdDefn As Command
	Dim iDay As Integer
	Dim iMonth As Integer
	Dim dtDate As Date
	Dim prmFilterID As ADODB.Parameter
	Dim rstPromptedValue As Recordset
	Dim sOptionValue As String
	Dim sDefaultValue As String
	Dim cmdLookupValues As Command
	Dim prmColumnID As ADODB.Parameter
	Dim rstLookupValues As Recordset
		

	iPromptCount = 0
	sPrompts = Request.Form("prompts")
	sFiltersAndCalcs = Request.Form("filtersAndCalcs")
	
	If Len(sPrompts) > 0 Then
		Response.Write("<table align='center' class='outline' cellPadding='5' cellSpacing='0'>" & vbCrLf)
		Response.Write("  <tr>" & vbCrLf)
		Response.Write("	  <td>" & vbCrLf)
		Response.Write("			<table align='center' class='invisible' cellspacing='0' cellpadding='0'>" & vbCrLf)
		Response.Write("				<tr>" & vbCrLf)
		Response.Write("					<td colspan='5' align='center'><h3 align='center'>Prompted Values</h3></td>" & vbCrLf)
		Response.Write("				</tr>" & vbCrLf)

		iParameterIndex = 0
		Do While Len(sPrompts) > 0
			iCharIndex = InStr(sPrompts, "	")
			iParameterIndex = iParameterIndex + 1

			If iCharIndex >= 0 Then
				Select Case iParameterIndex
					Case 1
						sNodeKey = Left(sPrompts, iCharIndex - 1)
					Case 2
						sPromptDescription = Left(sPrompts, iCharIndex - 1)
					Case 3
						iValueType = IIf(Left(sPrompts, iCharIndex - 1) = "", 0, Left(sPrompts, iCharIndex - 1))
					Case 4
						iPromptSize = IIf(Left(sPrompts, iCharIndex - 1) = "", 0, Left(sPrompts, iCharIndex - 1))
					Case 5
						iPromptDecimals = IIf(Left(sPrompts, iCharIndex - 1) = "", 0, Left(sPrompts, iCharIndex - 1))
					Case 6
						sPromptMask = Left(sPrompts, iCharIndex - 1)
					Case 7
						lngTableID = IIf(Left(sPrompts, iCharIndex - 1) = "", 0, Left(sPrompts, iCharIndex - 1))
					Case 8
						lngColumnID = IIf(Left(sPrompts, iCharIndex - 1) = "", 0, Left(sPrompts, iCharIndex - 1))
					Case 9
						sValueCharacter = Left(sPrompts, iCharIndex - 1)
					Case 10
						dblValueNumeric = IIf(Left(sPrompts, iCharIndex - 1) = "", 0, Left(sPrompts, iCharIndex - 1))
					Case 11
						fValueLogic = IIf(Left(sPrompts, iCharIndex - 1) = "", False, Left(sPrompts, iCharIndex - 1))
					Case 12
						dtValueDate = Left(sPrompts, iCharIndex - 1)
					Case 13
						iParameterIndex = 0
						iPromptDateType = IIf(Left(sPrompts, iCharIndex - 1) = "", 0, Left(sPrompts, iCharIndex - 1))
						
						' Got all of the required prompt paramters, so display it.
						iPromptCount = iPromptCount + 1
						Response.Write("    <tr>" & vbCrLf)
						Response.Write("      <td width='20'></td>" & vbCrLf)
						Response.Write("      <td>" & vbCrLf)
						
						If iValueType = 3 Then
							Response.Write("      <label " & vbCrLf)
							Response.Write("      for='prompt_3_" & sNodeKey & "'" & vbCrLf)
							Response.Write("      class='checkbox'" & vbCrLf)
							Response.Write("      tabindex='0' >" & vbCrLf)
						End If
						
						Response.Write("      " & sPromptDescription & vbCrLf)

						If iValueType = 3 Then
							Response.Write("</label>" & vbCrLf)
						End If
						
						Response.Write("      </td>" & vbCrLf)
						Response.Write("      <td width='20'></td>" & vbCrLf)
						Response.Write("      <td width='200'>" & vbCrLf)

						' Character Prompted Value
						If iValueType = "1" Then
							Response.Write("        <input type='text' class='text' id='prompt_1_" & sNodeKey & "' name='prompt_1_" & sNodeKey & "' value='" & Replace(sValueCharacter, """", "&quot;") & "' maxlength='" & iPromptSize & "' style='width: 100%'>" & vbCrLf)
							Response.Write("        <input type='hidden' id='PROMPTMASK_" & sNodeKey & "' name='PROMPTMASK_" & sNodeKey & "' value='" & Replace(sPromptMask, """", "&quot;") & "'>" & vbCrLf)

							' Numeric Prompted Value
						ElseIf iValueType = 2 Then
							Response.Write("        <input type='text' class='text' id='prompt_2_" & sNodeKey & "' name='prompt_2_" & sNodeKey & "' value='" & Replace(dblValueNumeric, ".", Session("LocaleDecimalSeparator")) & "' style='width: 100%'>" & vbCrLf)
							Response.Write("        <input type='hidden' id='promptSize_" & sNodeKey & "' name='promptSize_" & sNodeKey & "' value='" & iPromptSize & "'>" & vbCrLf)
							Response.Write("        <input type='hidden' id='promptDecs_" & sNodeKey & "' name='promptDecs_" & sNodeKey & "' value='" & iPromptDecimals & "'>" & vbCrLf)

							' Logic Prompted Value
						ElseIf iValueType = 3 Then
							Response.Write("        <input type='checkbox' id='prompt_3_" & sNodeKey & "' name='prompt_3_" & sNodeKey & "'" & vbCrLf)
							Response.Write("            onclick=""checkboxClick('" & sNodeKey & "')""" & vbCrLf)
							If fValueLogic Then
								Response.Write(" checked/>" & vbCrLf)
							Else
								Response.Write("/>" & vbCrLf)
							End If
							
							Response.Write("        <input type='hidden' id='promptChk_" & sNodeKey & "' name='promptChk_" & sNodeKey & "' value='")
							If fValueLogic Then
								Response.Write("TRUE'>" & vbCrLf)
							Else
								Response.Write("FALSE'>" & vbCrLf)
							End If
							 
							' Date Prompted Value
						ElseIf iValueType = 4 Then
							Response.Write("        <input type='text' class='text' id='prompt_4_" & sNodeKey & "' name='prompt_4_" & sNodeKey & "' value='")
							Select Case iPromptDateType
								Case 0
									' Explicit value
									If (dtValueDate <> "12/30/1899") Then
										Response.Write(convertSQLDateToLocale(dtValueDate))
									End If
								Case 1
									' Current date
									sTemp = convertDateToSQLDate(Date.Now)
									Response.Write(convertSQLDateToLocale(sTemp))
								Case 2
									' Start of current month
									iDay = (Day(Date.Now) * -1) + 1
									dtDate = DateAdd("d", iDay, Date.Now)
									sTemp = convertDateToSQLDate(dtDate)
									Response.Write(convertSQLDateToLocale(sTemp))
								Case 3
									' End of current month
									iDay = (Day(Date.Now) * -1) + 1
									dtDate = DateAdd("d", iDay, Date.Now)
									dtDate = DateAdd("m", 1, dtDate)
									dtDate = DateAdd("d", -1, dtDate)
									sTemp = convertDateToSQLDate(dtDate)
									Response.Write(convertSQLDateToLocale(sTemp))
								Case 4
									' Start of current year
									iDay = (Day(Date.Now) * -1) + 1
									iMonth = (Month(Date.Now) * -1) + 1
									dtDate = DateAdd("d", iDay, Date.Now)
									dtDate = DateAdd("m", iMonth, dtDate)
									sTemp = convertDateToSQLDate(dtDate)
									Response.Write(convertSQLDateToLocale(sTemp))
								Case 5
									' End of current year
									iDay = (Day(Date.Now) * -1) + 1
									iMonth = (Month(Date.Now) * -1) + 1
									dtDate = DateAdd("d", iDay, Date.Now)
									dtDate = DateAdd("m", iMonth, dtDate)
									dtDate = DateAdd("yyyy", 1, dtDate)
									dtDate = DateAdd("d", -1, dtDate)
									sTemp = convertDateToSQLDate(dtDate)
									Response.Write(convertSQLDateToLocale(sTemp))
							End Select
							Response.Write("' style='width: 100%'>" & vbCrLf)

							' Lookup Prompted Value
						ElseIf iValueType = 5 Then
							Response.Write("        <select id='promptLookup_" & sNodeKey & "' class='combo' name='promptLookup_" & sNodeKey & "' style='width: 100%' onchange=""comboChange('" & sNodeKey & "')"">" & vbCrLf)

							fDefaultFound = False
							fFirstValueDone = False
							sFirstValue = ""

							' Get the lookup values.
							cmdLookupValues = New Command
							cmdLookupValues.CommandText = "sp_ASRIntGetLookupValues"
							cmdLookupValues.CommandType = CommandTypeEnum.adCmdStoredProc
							cmdLookupValues.ActiveConnection = Session("databaseConnection")

							prmColumnID = cmdLookupValues.CreateParameter("columnID", 3, 1)
							cmdLookupValues.Parameters.Append(prmColumnID)
							prmColumnID.value = CleanNumeric(lngColumnID)

							Err.Clear()
							rstLookupValues = cmdLookupValues.Execute
							Do While Not rstLookupValues.EOF
								Response.Write("          <option")

								If Not fFirstValueDone Then
									sFirstValue = rstLookupValues.Fields(0).Value
									fFirstValueDone = True
								End If
								
								If rstLookupValues.fields(0).type = 135 Then
									' Field is a date so format as such.
									sOptionValue = convertSQLDateToLocale2(rstLookupValues.Fields(0).Value)
									If sOptionValue = convertSQLDateToLocale(sValueCharacter) Then
										Response.Write(" selected")
										fDefaultFound = True
									End If
									Response.Write(">" & sOptionValue & "</option>" & vbCrLf)

								ElseIf rstLookupValues.fields(0).type = 131 Then
									' Field is a numeric so format as such.
									sOptionValue = Replace(rstLookupValues.Fields(0).Value, ".", Session("LocaleDecimalSeparator"))
									If (Not IsDBNull(rstLookupValues.Fields(0).Value)) And (Not IsDBNull(sValueCharacter)) Then
										If FormatNumber(rstLookupValues.Fields(0).Value) = FormatNumber(sValueCharacter) Then
											Response.Write(" selected")
											fDefaultFound = True
										End If
									End If
									Response.Write(">" & sOptionValue & "</option>" & vbCrLf)
								ElseIf rstLookupValues.fields(0).type = 11 Then
									' Field is a logic so format as such.
									sOptionValue = rstLookupValues.Fields(0).Value
									If sOptionValue = sValueCharacter Then
										Response.Write(" selected")
										fDefaultFound = True
									End If
									Response.Write(">" & sOptionValue & "</option>" & vbCrLf)
								Else
									sOptionValue = RTrim(rstLookupValues.Fields(0).Value)
									If sOptionValue = sValueCharacter Then
										Response.Write(" selected")
										fDefaultFound = True
									End If
									Response.Write(">" & sOptionValue & "</option>" & vbCrLf)
								End If

								rstLookupValues.MoveNext()
							Loop

							Response.Write("        </select>" & vbCrLf)

							If fDefaultFound Then
								sDefaultValue = sValueCharacter
							Else
								sDefaultValue = sFirstValue
							End If

							If rstLookupValues.fields(0).type = 135 Then
								' Date.
								Response.Write("        <input type='hidden' id='prompt_4_" & sNodeKey & "' name='prompt_4_" & sNodeKey & "' value='" & convertSQLDateToLocale(sDefaultValue) & "'>" & vbCrLf)
							ElseIf rstLookupValues.fields(0).type = 131 Then
								' Numeric
								Response.Write("        <input type='hidden' id='prompt_2_" & sNodeKey & "' name='prompt_2_" & sNodeKey & "' value='" & Replace(sDefaultValue, ".", Session("LocaleDecimalSeparator")) & "'>" & vbCrLf)
							ElseIf rstLookupValues.fields(0).type = 11 Then
								' Logic
								Response.Write("        <input type='hidden' id='prompt_3_" & sNodeKey & "' name='prompt_3_" & sNodeKey & "' value='" & sDefaultValue & "'>" & vbCrLf)
							Else
								Response.Write("        <input type='hidden' id='prompt_1_" & sNodeKey & "' name='prompt_1_" & sNodeKey & "' value='" & Replace(sDefaultValue, """", "&quot;") & "'>" & vbCrLf)
							End If

							' Release the ADO recordset object.
							rstLookupValues.close()
							rstLookupValues = Nothing
						End If
				
						Response.Write("					</td>" & vbCrLf)
						Response.Write("					<td width='20' height='10'></td>" & vbCrLf)
						Response.Write("				</tr>" & vbCrLf)
				End Select

				sPrompts = Mid(sPrompts, iCharIndex + 1)
			End If
		Loop
	End If

	If Len(sFiltersAndCalcs) > 0 Then
		Do While Len(sFiltersAndCalcs) > 0
			iCharIndex = InStr(sFiltersAndCalcs, "	")

			If iCharIndex >= 0 Then
				sFilterCalcID = Left(sFiltersAndCalcs, iCharIndex - 1)
				sFiltersAndCalcs = Mid(sFiltersAndCalcs, iCharIndex + 1)

				cmdDefn = New Command
				cmdDefn.CommandText = "sp_ASRIntGetFilterPromptedValuesRecordset"
				cmdDefn.CommandType = CommandTypeEnum.adCmdStoredProc
				cmdDefn.ActiveConnection = Session("databaseConnection")

				prmFilterID = cmdDefn.CreateParameter("filterID", 3, 1)	' 3=integer, 1=input
				cmdDefn.Parameters.Append(prmFilterID)
				prmFilterID.value = CleanNumeric(CLng(sFilterCalcID))

				Err.Clear()
				rstPromptedValue = cmdDefn.Execute

				If Not (rstPromptedValue.EOF And rstPromptedValue.BOF) Then
					If iPromptCount = 0 Then
						Response.Write("<table align='center' class='outline' cellPadding='5' cellSpacing='0'>" & vbCrLf)
						Response.Write("  <tr>" & vbCrLf)
						Response.Write("	  <td>" & vbCrLf)
						Response.Write("			<table align='center' class='invisible' cellspacing='0' cellpadding='0'>" & vbCrLf)
						Response.Write("				<tr>" & vbCrLf)
						Response.Write("					<td colspan='5' align='center'><h3 align='center'>Prompted Values</h3></td>" & vbCrLf)
						Response.Write("				</tr>" & vbCrLf)
					End If
					
					Do While Not rstPromptedValue.EOF
						iPromptCount = iPromptCount + 1
						Response.Write("				<tr height='10'>" & vbCrLf)
						Response.Write("					<td width='20' height='10'></td>" & vbCrLf)
						Response.Write("					<td nowrap height='10'>" & vbCrLf)
						
						If iValueType = 3 Then
							Response.Write("          <label " & vbCrLf)
							Response.Write("            for='prompt_3_C" & rstPromptedValue.fields("componentID").value & "'" & vbCrLf)
							Response.Write("            class='checkbox'" & vbCrLf)
							Response.Write("            tabindex='0' >" & vbCrLf)
						End If

						Response.Write("					  " & rstPromptedValue.fields("PromptDescription").value & vbCrLf)

						If iValueType = 3 Then
							Response.Write("          </label>" & vbCrLf)
						End If

						Response.Write("					</td>" & vbCrLf)
						Response.Write("					<td width='20' height='10'>&nbsp;</td>" & vbCrLf)
						Response.Write("   		    <td width='200' height='10'>" & vbCrLf)

						' Character Prompted Value
						If rstPromptedValue.fields("ValueType").value = 1 Then
							Response.Write("    		    <input type='text' class='text' id='prompt_1_C" & rstPromptedValue.fields("componentID").value & "' name='prompt_1_C" & rstPromptedValue.fields("componentID").value & "' value='" & Replace(rstPromptedValue.fields("valuecharacter").value, """", "&quot;") & "' maxlength='" & rstPromptedValue.fields("promptsize").value & "' style='width: 100%'>" & vbCrLf)
							Response.Write("    		    <input type='hidden' id='promptMask_C" & rstPromptedValue.fields("componentID").value & "' name='promptMask_C" & rstPromptedValue.fields("componentID").value & "' value='" & Replace(rstPromptedValue.fields("promptMask").value, """", "&quot;") & "'>" & vbCrLf)

							' Numeric Prompted Value
						ElseIf rstPromptedValue.fields("ValueType").value = 2 Then
							Response.Write("     		   <input type='text' class='text' id='prompt_2_C" & rstPromptedValue.fields("componentID").value & "' name='prompt_2_C" & rstPromptedValue.fields("componentID").value & "' value='" & Replace(rstPromptedValue.fields("valuenumeric").value, ".", Session("LocaleDecimalSeparator")) & "' style='width: 100%'>" & vbCrLf)
							Response.Write("     		   <input type='hidden' id='promptSize_C" & rstPromptedValue.fields("componentID").value & "' name='promptSize_C" & rstPromptedValue.fields("componentID").value & "' value='" & rstPromptedValue.fields("promptSize").value & "'>" & vbCrLf)
							Response.Write("     		   <input type='hidden' id='promptDecs_C" & rstPromptedValue.fields("componentID").value & "' name='promptDecs_C" & rstPromptedValue.fields("componentID").value & "' value='" & rstPromptedValue.fields("promptDecimals").value & "'>" & vbCrLf)

							' Logic Prompted Value
						ElseIf rstPromptedValue.fields("ValueType").value = 3 Then
							Response.Write("        <input type='checkbox' tabindex='-1' id='prompt_3_C" & rstPromptedValue.fields("componentID").value & "' name='prompt_3_C" & rstPromptedValue.fields("componentID").value & "'" & vbCrLf)
							Response.Write("            onclick=""checkboxClick('C" & rstPromptedValue.fields("componentID").value & "')""" & vbCrLf)
							If rstPromptedValue.fields("valuelogic").value Then
								Response.Write(" checked/>" & vbCrLf)
							Else
								Response.Write("/>" & vbCrLf)
							End If
							
							Response.Write("        <input type='hidden' id='promptChk_C" & rstPromptedValue.fields("componentID").value & "' name='promptChk_C" & rstPromptedValue.fields("componentID").value & "' value='" & rstPromptedValue.fields("valuelogic").value & "'>" & vbCrLf)
											 
							' Date Prompted Value
						ElseIf rstPromptedValue.Fields("ValueType").Value = 4 Then
							Response.Write("        <input type='text' class='text' id='prompt_4_C" & rstPromptedValue.Fields("componentID").Value & "' name='prompt_4_C" & rstPromptedValue.Fields("componentID").Value & "' value='")
							Select Case rstPromptedValue.Fields("promptDateType").Value
								Case 0
									' Explicit value
									If Not IsDBNull(rstPromptedValue.Fields("valuedate").Value) Then
										If (CStr(rstPromptedValue.Fields("valuedate").Value) <> "00:00:00") And _
												(CStr(rstPromptedValue.Fields("valuedate").Value) <> "12:00:00 AM") Then
											Response.Write(convertSQLDateToLocale2(rstPromptedValue.Fields("valuedate").Value))
										End If
									End If
								Case 1
									' Current date
									Response.Write(convertSQLDateToLocale2(Date.Now))
								Case 2
									' Start of current month
									iDay = (Day(Date.Now) * -1) + 1
									dtDate = DateAdd("d", iDay, Date.Now)
									Response.Write(convertSQLDateToLocale2(dtDate))
								Case 3
									' End of current month
									iDay = (Day(Date.Now) * -1) + 1
									dtDate = DateAdd("d", iDay, Date.Now)
									dtDate = DateAdd("m", 1, dtDate)
									dtDate = DateAdd("d", -1, dtDate)
									Response.Write(convertSQLDateToLocale2(dtDate))
								Case 4
									' Start of current year
									iDay = (Day(Date.Now) * -1) + 1
									iMonth = (Month(Date.Now) * -1) + 1
									dtDate = DateAdd("d", iDay, Date.Now)
									dtDate = DateAdd("m", iMonth, dtDate)
									Response.Write(convertSQLDateToLocale2(dtDate))
								Case 5
									' End of current year
									iDay = (Day(Date.Now) * -1) + 1
									iMonth = (Month(Date.Now) * -1) + 1
									dtDate = DateAdd("d", iDay, Date.Now)
									dtDate = DateAdd("m", iMonth, dtDate)
									dtDate = DateAdd("yyyy", 1, dtDate)
									dtDate = DateAdd("d", -1, dtDate)
									Response.Write(convertSQLDateToLocale2(dtDate))
							End Select
							Response.Write("' style='width: 100%'>" & vbCrLf)

							' Lookup Prompted Value
						ElseIf rstPromptedValue.Fields("ValueType").Value = 5 Then
							Response.Write("        		<select id='promptLookup_C" & rstPromptedValue.Fields("componentID").Value & "' name='promptLookup_C" & rstPromptedValue.Fields("componentID").Value & "' class='combo' style='width: 100%' onchange=""comboChange('C" & rstPromptedValue.Fields("componentID").Value & "')"">" & vbCrLf)

							fDefaultFound = False
							fFirstValueDone = False
							sFirstValue = ""

							' Get the lookup values.
							cmdLookupValues = New Command
							cmdLookupValues.CommandText = "sp_ASRIntGetLookupValues"
							cmdLookupValues.CommandType = CommandTypeEnum.adCmdStoredProc
							cmdLookupValues.ActiveConnection = Session("databaseConnection")

							prmColumnID = cmdLookupValues.CreateParameter("columnID", 3, 1)
							cmdLookupValues.Parameters.Append(prmColumnID)
							prmColumnID.Value = CleanNumeric(rstPromptedValue.Fields("fieldColumnID").Value)

							Err.Clear()
							rstLookupValues = cmdLookupValues.Execute

							Do While Not rstLookupValues.EOF
								Response.Write("        		  <option")

								If Not fFirstValueDone Then
									sFirstValue = rstLookupValues.Fields(0).Value
									fFirstValueDone = True
								End If

								If rstLookupValues.Fields(0).Type = 135 Then
									' Field is a date so format as such.
									sOptionValue = convertSQLDateToLocale2(rstLookupValues.Fields(0).Value)
									If sOptionValue = convertSQLDateToLocale2(rstPromptedValue.Fields("valuecharacter").Value) Then
										Response.Write(" selected")
										fDefaultFound = True
									End If
									Response.Write(">" & sOptionValue & "</option>" & vbCrLf)
								ElseIf rstLookupValues.Fields(0).Type = 131 Then
									' Field is a numeric so format as such.
									sOptionValue = Replace(rstLookupValues.Fields(0).Value, ".", Session("LocaleDecimalSeparator"))
									If (Not IsDBNull(rstLookupValues.Fields(0).Value)) And (Not IsDBNull(rstPromptedValue.Fields("valuecharacter").Value)) Then
										If FormatNumber(rstLookupValues.Fields(0).Value) = FormatNumber(rstPromptedValue.Fields("valuecharacter").Value) Then
											Response.Write(" selected")
											fDefaultFound = True
										End If
									End If
									Response.Write(">" & sOptionValue & "</option>" & vbCrLf)
								ElseIf rstLookupValues.Fields(0).Type = 11 Then
									' Field is a logic so format as such.
									sOptionValue = rstLookupValues.Fields(0).Value
									If sOptionValue = rstPromptedValue.Fields("valuecharacter").Value Then
										Response.Write(" selected")
										fDefaultFound = True
									End If
									Response.Write(">" & sOptionValue & "</option>" & vbCrLf)
								Else
									sOptionValue = RTrim(rstLookupValues.Fields(0).Value)
									If sOptionValue = rstPromptedValue.Fields("valuecharacter").Value Then
										Response.Write(" selected")
										fDefaultFound = True
									End If
									Response.Write(">" & sOptionValue & "</option>" & vbCrLf)
								End If

								rstLookupValues.MoveNext()
							Loop

							Response.Write("   		     </select>" & vbCrLf)

							If fDefaultFound Then
								sDefaultValue = rstPromptedValue.Fields("valuecharacter").Value
							Else
								sDefaultValue = sFirstValue
							End If

							If rstLookupValues.Fields(0).Type = 135 Then
								' Date.
								Response.Write("        <input type='hidden' id='prompt_4_C" & rstPromptedValue.Fields("componentID").Value & "' name='prompt_4_C" & rstPromptedValue.Fields("componentID").Value & "' value='" & convertSQLDateToLocale(sDefaultValue) & "'>" & vbCrLf)
							ElseIf rstLookupValues.Fields(0).Type = 131 Then
								' Numeric
								Response.Write("        <input type='hidden' id='prompt_2_C" & rstPromptedValue.Fields("componentID").Value & "' name='prompt_2_C" & rstPromptedValue.Fields("componentID").Value & "' value='" & Replace(sDefaultValue, ".", Session("LocaleDecimalSeparator")) & "'>" & vbCrLf)
							ElseIf rstLookupValues.Fields(0).Type = 11 Then
								' Logic
								Response.Write("        <input type='hidden' id='prompt_3_C" & rstPromptedValue.Fields("componentID").Value & "' name='prompt_3_C" & rstPromptedValue.Fields("componentID").Value & "' value='" & sDefaultValue & "'>" & vbCrLf)
							Else
								Response.Write("        <input type='hidden' id='prompt_1_C" & rstPromptedValue.Fields("componentID").Value & "' name='prompt_1_C" & rstPromptedValue.Fields("componentID").Value & "' value='" & Replace(sDefaultValue, """", "&quot;") & "'>" & vbCrLf)
							End If

							' Release the ADO recordset object.
							rstLookupValues.Close()
							rstLookupValues = Nothing
						End If
								
						Response.Write("					</td>" & vbCrLf)
						Response.Write("					<td width='20' height='10'></td>" & vbCrLf)
						Response.Write("				</tr>" & vbCrLf)

						rstPromptedValue.MoveNext()
					Loop
				End If

				rstPromptedValue.close()
				rstPromptedValue = Nothing
				cmdDefn = Nothing
			End If
		Loop
	End If

	If iPromptCount > 0 Then
		Response.Write("				<tr>" & vbCrLf)
		Response.Write("					<td colspan='5' height='10'>&nbsp;</td>" & vbCrLf)
		Response.Write("			  </tr>" & vbCrLf)
		Response.Write("				<tr height='20'>" & vbCrLf)
		Response.Write("					<td width='20'></td>" & vbCrLf)
		Response.Write("					<td colspan='3'>" & vbCrLf)
		Response.Write("						<table width='100%' class='invisible' cellspacing='0' cellpadding='0'>" & vbCrLf)
		Response.Write("							<td>&nbsp;</td>" & vbCrLf)
			
		Response.Write("							<td width='80'>" & vbCrLf)
		Response.Write("							    <input type='button' value='OK' name='Submit' class='btn' style='width: 80px'" & vbCrLf)
		Response.Write("									    onclick='SubmitPrompts();' />" & vbCrLf)
		Response.Write("							</td>")
		Response.Write("							<td width='20'></td>" & vbCrLf)
		Response.Write("							<td width='80'>" & vbCrLf)
		Response.Write("							    <input type='button' value='Cancel' name='Cancel' class='btn' value='Cancel' style='width: 80px'" & vbCrLf)
		Response.Write("									    onclick='CancelClick();' />" & vbCrLf)
		Response.Write("							</td>" & vbCrLf)
		Response.Write("						</table>" & vbCrLf)
		Response.Write("					</td>" & vbCrLf)
		Response.Write("					<td width='20'></td>" & vbCrLf)
		Response.Write("				</tr>" & vbCrLf)
		Response.Write("				<tr>" & vbCrLf)
		Response.Write("					<td colspan='5' height='5'></td>" & vbCrLf)
		Response.Write("				</tr>" & vbCrLf)
		Response.Write("			</table>" & vbCrLf)
		Response.Write("		</td>" & vbCrLf)
		Response.Write("	</tr>" & vbCrLf)
		Response.Write("</table>" & vbCrLf)
	End If

	Response.Write("<input type='hidden' id='txtPromptCount' name='txtPromptCount' value='" & iPromptCount & "'>" & vbCrLf)
%>
	<input type="hidden" id="type" name="type" value="<%=Request.Form("type")%>" />	
	<input type="hidden" id="components1" name="components1" value="<% =Request.Form("components1")%>" />
	<input type="hidden" id="tableID" name="tableID" value="<%=Request.Form("tableID")%>" />
</form>
		
		</div>
</body>
</html>


<script type="text/javascript">
		util_test_expression_pval_onload();
</script>



<script runat="server" language="vb">

	Function promptParameter(psDefnString As String, psParameter As String) As String
		Dim iCharIndex As Integer
		Dim sDefn As String
	
		sDefn = psDefnString

		iCharIndex = InStr(sDefn, "	")
		If iCharIndex >= 0 Then
			If psParameter = "NODEKEY" Then
				promptParameter = Left(sDefn, iCharIndex - 1)
				Exit Function
			End If
		
			sDefn = Mid(sDefn, iCharIndex + 1)
			iCharIndex = InStr(sDefn, "	")
			If iCharIndex >= 0 Then
				If psParameter = "PROMPTDESCRIPTION" Then
					promptParameter = Left(sDefn, iCharIndex - 1)
					Exit Function
				End If
			
				sDefn = Mid(sDefn, iCharIndex + 1)
				iCharIndex = InStr(sDefn, "	")
				If iCharIndex >= 0 Then
					If psParameter = "VALUETYPE" Then
						promptParameter = Left(sDefn, iCharIndex - 1)
						Exit Function
					End If
				
					sDefn = Mid(sDefn, iCharIndex + 1)
					iCharIndex = InStr(sDefn, "	")
					If iCharIndex >= 0 Then
						If psParameter = "PROMPTSIZE" Then
							promptParameter = Left(sDefn, iCharIndex - 1)
							Exit Function
						End If
					
						sDefn = Mid(sDefn, iCharIndex + 1)
						iCharIndex = InStr(sDefn, "	")
						If iCharIndex >= 0 Then
							If psParameter = "PROMPTDECIMALS" Then
								promptParameter = Left(sDefn, iCharIndex - 1)
								Exit Function
							End If
						
							sDefn = Mid(sDefn, iCharIndex + 1)
							iCharIndex = InStr(sDefn, "	")
							If iCharIndex >= 0 Then
								If psParameter = "PROMPTMASK" Then
									promptParameter = Left(sDefn, iCharIndex - 1)
									Exit Function
								End If
							
								sDefn = Mid(sDefn, iCharIndex + 1)
								iCharIndex = InStr(sDefn, "	")
								If iCharIndex >= 0 Then
									If psParameter = "FIELDTABLEID" Then
										promptParameter = Left(sDefn, iCharIndex - 1)
										Exit Function
									End If
								
									sDefn = Mid(sDefn, iCharIndex + 1)
									iCharIndex = InStr(sDefn, "	")
									If iCharIndex >= 0 Then
										If psParameter = "FIELDCOLUMNID" Then
											promptParameter = Left(sDefn, iCharIndex - 1)
											Exit Function
										End If
									
										sDefn = Mid(sDefn, iCharIndex + 1)
										iCharIndex = InStr(sDefn, "	")
										If iCharIndex >= 0 Then
											If psParameter = "VALUECHARACTER" Then
												promptParameter = Left(sDefn, iCharIndex - 1)
												Exit Function
											End If
										
											sDefn = Mid(sDefn, iCharIndex + 1)
											iCharIndex = InStr(sDefn, "	")
											If iCharIndex >= 0 Then
												If psParameter = "VALUENUMERIC" Then
													promptParameter = Left(sDefn, iCharIndex - 1)
													Exit Function
												End If
											
												sDefn = Mid(sDefn, iCharIndex + 1)
												iCharIndex = InStr(sDefn, "	")
												If iCharIndex >= 0 Then
													If psParameter = "VALUELOGIC" Then
														promptParameter = Left(sDefn, iCharIndex - 1)
														Exit Function
													End If
												
													sDefn = Mid(sDefn, iCharIndex + 1)
													iCharIndex = InStr(sDefn, "	")
													If iCharIndex >= 0 Then
														If psParameter = "VALUEDATE" Then
															promptParameter = Left(sDefn, iCharIndex - 1)
															Exit Function
														End If
													
														sDefn = Mid(sDefn, iCharIndex + 1)
														If psParameter = "PROMPTDATETYPE" Then
															promptParameter = Left(sDefn, iCharIndex - 1)
															Exit Function
														End If
													End If
												End If
											End If
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
	
		promptParameter = ""
	End Function

	Function convertDateToSQLDate(pdtDate) As String
		Dim iDays As Integer
		Dim iMonths As Integer
		Dim iYears As Integer
		Dim sResult As String
	
		sResult = ""
		iDays = Day(pdtDate)
		iMonths = Month(pdtDate)
		iYears = Year(pdtDate)

		If iMonths < 10 Then
			sResult = "0"
		End If
		sResult = sResult & iMonths & "/"
	
		If iDays < 10 Then
			sResult = sResult & "0"
		End If
		sResult = sResult & iDays & "/" & iYears
	
		convertDateToSQLDate = sResult
	End Function


	Function convertSQLDateToLocale2(psDate As String) As String
		Dim sLocaleFormat As String
		Dim iIndex As Integer

		If Len(psDate) > 0 Then
			sLocaleFormat = Session("LocaleDateFormat")
		
			iIndex = InStr(sLocaleFormat, "dd")
			If iIndex > 0 Then
				If Day(psDate) < 10 Then
					sLocaleFormat = Left(sLocaleFormat, iIndex - 1) & _
						"0" & Day(psDate) & Mid(sLocaleFormat, iIndex + 2)
				Else
					sLocaleFormat = Left(sLocaleFormat, iIndex - 1) & _
						Day(psDate) & Mid(sLocaleFormat, iIndex + 2)
				End If
			End If
		
			iIndex = InStr(sLocaleFormat, "mm")
			If iIndex > 0 Then
				If Month(psDate) < 10 Then
					sLocaleFormat = Left(sLocaleFormat, iIndex - 1) & _
						"0" & Month(psDate) & Mid(sLocaleFormat, iIndex + 2)
				Else
					sLocaleFormat = Left(sLocaleFormat, iIndex - 1) & _
						Month(psDate) & Mid(sLocaleFormat, iIndex + 2)
				End If
			End If
		
			iIndex = InStr(sLocaleFormat, "yyyy")
			If iIndex > 0 Then
				sLocaleFormat = Left(sLocaleFormat, iIndex - 1) & _
					Year(psDate) & Mid(sLocaleFormat, iIndex + 4)
			End If

			convertSQLDateToLocale2 = sLocaleFormat
		Else
			convertSQLDateToLocale2 = ""
		End If
	End Function
</script>
