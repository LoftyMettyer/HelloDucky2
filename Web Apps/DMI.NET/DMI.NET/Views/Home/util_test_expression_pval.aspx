<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage(of DMI.NET.Models.ObjectRequests.TestExpressionModel)" %>

<%@ Import Namespace="System.Globalization" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="DMI.NET.Helpers" %>

<!DOCTYPE html>


<script type="text/javascript">

	function util_test_expression_pval_onload() {

		window.UserLocale = "<%:Session("LocaleCultureName").ToString().ToLower()%>";
		window.parent.OpenHR.setDatepickerLanguage();
		
		var frmPromptedValues = document.getElementById('frmPromptedValues');			

		$(".datepicker").datepicker();
		$(document).on('keydown', '.datepicker', function (event) {
			switch (event.keyCode) {
				case 113:
					$(this).datepicker("setDate", new Date());
					$(this).datepicker('widget').hide('true');
					break;
			}
		});

		// Prevent default behaviour of empty form
		$("input[type='text']").keydown(function (e) {
			if (e.keyCode == 13) {
				e.preventDefault();
			}
		});
		
		if (frmPromptedValues.txtPromptCount.value == 0) {

			var postData = {
				UtilType: <%:CInt(Model.type)%>,
				components1: "<%:Model.components1%>",
				TableID: <%:Model.tableID%>,
				<%:Html.AntiForgeryTokenForAjaxPost() %> };
			OpenHR.submitForm(null, "divValidateExpression", null, postData, "util_test_expression");

		}
		else {
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

			// Resize the grid to show all prompted values.
			var iResizeBy = frmPromptedValues.offsetParent.scrollHeight - frmPromptedValues.offsetParent.clientHeight;
			if (frmPromptedValues.offsetParent.offsetHeight + iResizeBy > screen.height) {
				window.parent.outerHeight = new String(screen.height);
			}
			else {
				var iNewHeight = window.parent.outerHeight;
				iNewHeight = iNewHeight + iResizeBy;
				window.parent.outerHeight = new String(iNewHeight);
			}

			if ($('#divValidateExpression').dialog('isOpen')) {
				var dialogWidth = screen.width / 3;

				$('#divValidateExpression').dialog('option', 'height', 'auto');
				$('#divValidateExpression').dialog('option', 'width', dialogWidth);
			}
		}

		if (frmPromptedValues.txtPromptCount.value > 0) {
			
			var controlCollection = frmPromptedValues.elements;
			var icolumnSize =0; 
			var iDecimals =2; 
			var maximumValue = 9999999999;
			var sSource=[];
			var dDefaultValue =[];
			var icolumnSize=[];
			var iDecimals=[];			

			if (controlCollection != null) {				
				for (var i = 0; i < controlCollection.length; i++) 
				{					
					var sControlName = controlCollection.item(i).name;						
					if (sControlName.substr(0, 9) == "prompt_2_") 
					{					
						sSource.push(controlCollection.item(i).id);						
						dDefaultValue.push(controlCollection.item(i).defaultValue);
						$("#"+sSource[i]).autoNumeric('destroy');						
					}
				}
				
				for (var i = 0; i < controlCollection.length; i++) 
				{//get column size of Numeric prompted value
					var sControlName = controlCollection.item(i).name;					
					if(sControlName.substr(0, 11) == "promptSize_") {
						icolumnSize.push(controlCollection.item(i).value);						
					}					
				}

				for (var i = 0; i < controlCollection.length; i++) 
				{//get decimal size of Numeric prompted value
					var sControlName = controlCollection.item(i).name;
					if(sControlName.substr(0, 11) == "promptDecs_")
					{
						iDecimals.push(controlCollection.item(i).value);						
					}				
				}	
				
				for (var j=0; j < sSource.length; j++)
				{
					var maxValueAllowedBeforeDecimal = '9';
					for (var i = 1; i < parseInt(icolumnSize[j]) ; i++) {
						maxValueAllowedBeforeDecimal = maxValueAllowedBeforeDecimal + '9';
					}

					var maxValueAllowedAfterDecimal = '9';
					for (var i = 1; i < parseInt(iDecimals[j]) ; i++) {
						maxValueAllowedAfterDecimal = maxValueAllowedAfterDecimal + '9';
					}					
					// Set the maximum value
					maximumValue = parseFloat(maxValueAllowedBeforeDecimal + "." + maxValueAllowedAfterDecimal);

					//Set autonumeric for individual Numeric prompted value
					$('#'+sSource[j]).autoNumeric('init',{ aSep: ',', aNeg: '', vMax: maximumValue, mDec: parseInt(iDecimals[j]), aPad: false,mRound: 'S'});					
					$('#'+sSource[j]).val(dDefaultValue[j]);
					
				}				

		 }
	 }

}
	
	function SubmitPrompts() {

		var frmPromptedValues = document.getElementById('frmPromptedValues');

		// Validate the prompt values before submitting the form.
		var controlCollection = frmPromptedValues.elements;
		var submitElements = [];
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

					submitElements.push({
						Key: controlCollection.item(i).name,
						Type: iType,
						Value: controlCollection.item(i).value
					});

				}
			}
		}

		// Everything OK. Submit the form.
		var postData = {
			UtilType: <%:CInt(Model.type)%>,
			components1: "<%:Model.components1%>",
			TableID: <%:Model.tableID%>,
			PromptValues: submitElements,
			<%:Html.AntiForgeryTokenForAjaxPost() %> };
		OpenHR.submitForm(null, "divValidateExpression", null, postData, "util_test_expression");

	}

	function ute_cancelClick() {
		if ($('#divValidateExpression').dialog('isOpen') == true) {
			$('#divValidateExpression').dialog('close');
			$('#divValidateExpression').html();
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

		sDecimalSeparator = "\\";
		sDecimalSeparator = sDecimalSeparator.concat(OpenHR.LocaleDecimalSeparator());
		var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");

		sThousandSeparator = "\\";
		sThousandSeparator = sThousandSeparator.concat(OpenHR.LocaleThousandSeparator());
		var reThousandSeparator = new RegExp(sThousandSeparator, "gi");

		sPoint = "\\.";
		var rePoint = new RegExp(sPoint, "gi");

		var frmPromptedValues = document.getElementById('frmPromptedValues');

		fOK = true;
		var sValue;
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
			if (OpenHR.LocaleDecimalSeparator() != ".") {
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

			fOK = OpenHR.IsValidDate(pctlPrompt.value);

			if (fOK == false) {
				OpenHR.messageBox("Invalid date value entered.");
			}

		}

		if ((fOK == true) && (piDataType == 1)) {
			// Character column.
			// Ensure that the value entered matches the required mask (if there is one).
			var sMaskCtlName = "promptMask_" + pctlPrompt.name.substring(9, pctlPrompt.name.length);

			var fFound = false;
			var controlCollection = frmPromptedValues.elements;
			var i;
			if (controlCollection != null) {
				for (i = 0; i < controlCollection.length; i++) {
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
						for (i = 0; i < sMask.length; i++) {
							var sValueChar = sValue.substring(iIndex, iIndex + 1);

							if (fFollowingBackslash == false) {
								var iNumber;
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
							} else {
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
					pctlPrompt.focus();
				}
			}
		}

		return fOK;
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


<div data-framesource="util_test_expression_pval">

	<form name="frmPromptedValues" id="frmPromptedValues" method="POST" action="<%:Url.Action("util_test_expression")%>">
		<%
			Dim iPromptCount As Integer
			Dim sPrompts As String
			Dim sNodeKey As String = ""
			Dim sPromptDescription As String = ""
			Dim iValueType As Integer
			Dim iPromptSize As Integer
			Dim iPromptDecimals As Integer
			Dim sPromptMask As String = ""
			Dim lngTableID As Long
			Dim lngColumnID As Long
			Dim sValueCharacter As String = ""
			Dim dblValueNumeric As Double
			Dim fValueLogic As Boolean
			Dim dtValueDate As Date
			Dim iPromptDateType As Integer
			Dim iCharIndex As Integer
			Dim iParameterIndex As Integer
			Dim sTemp As String
			Dim sFiltersAndCalcs As String
			Dim sFilterCalcID As String
			Dim fDefaultFound As Boolean
			Dim fFirstValueDone As Boolean
			Dim sFirstValue As String

			Dim iDay As Integer
			Dim iMonth As Integer
			Dim dtDate As Date
			Dim rstPromptedValue As DataTable
			Dim sOptionValue As String
			Dim sDefaultValue As String
			Dim rstLookupValues As DataTable
			Dim rowLookupValues As DataRow
			Dim objSession As SessionInfo = CType(Session("SessionContext"), SessionInfo)	'Set session info
			Dim objDataAccess As New clsDataAccess(objSession.LoginInfo) 'Instantiate DataAccess class

			iPromptCount = 0
			sPrompts = Model.prompts
			sFiltersAndCalcs = Model.filtersAndCalcs
	
			If Len(sPrompts) > 0 Then
				iParameterIndex = 0
				Do While Len(sPrompts) > 0
					iCharIndex = InStr(sPrompts, "	")
					iParameterIndex = iParameterIndex + 1
										
					If iCharIndex >= 0 Then
						Select Case iParameterIndex
							Case ExpressionParameter.NodeKey
								sNodeKey = Left(sPrompts, iCharIndex - 1)
							Case ExpressionParameter.PromptDescription
								sPromptDescription = Left(sPrompts, iCharIndex - 1)
							Case ExpressionParameter.ValueType
								iValueType = CType(IIf(Left(sPrompts, iCharIndex - 1) = "", 0, Left(sPrompts, iCharIndex - 1)), Integer)
							Case ExpressionParameter.PromptSize
								iPromptSize = CType(IIf(Left(sPrompts, iCharIndex - 1) = "", 0, Left(sPrompts, iCharIndex - 1)), Integer)
							Case ExpressionParameter.PromptDecimals
								iPromptDecimals = CType(IIf(Left(sPrompts, iCharIndex - 1) = "", 0, Left(sPrompts, iCharIndex - 1)), Integer)
							Case ExpressionParameter.PromptMask
								sPromptMask = Left(sPrompts, iCharIndex - 1)
							Case ExpressionParameter.TableID
								lngTableID = CType(IIf(Left(sPrompts, iCharIndex - 1) = "", 0, Left(sPrompts, iCharIndex - 1)), Long)
							Case ExpressionParameter.ColumnID
								lngColumnID = CType(IIf(Left(sPrompts, iCharIndex - 1) = "", 0, Left(sPrompts, iCharIndex - 1)), Long)
							Case ExpressionParameter.ValueCharacter
								sValueCharacter = Left(sPrompts, iCharIndex - 1)
							Case ExpressionParameter.ValueNumeric
								sTemp = Left(sPrompts, iCharIndex - 1)
								If sTemp = "" Then
									dblValueNumeric = Nothing
								Else
									dblValueNumeric = Double.Parse(Left(sPrompts, iCharIndex - 1), CultureInfo.InvariantCulture)
								End If
							Case ExpressionParameter.ValueLogic
								fValueLogic = CType(IIf(Left(sPrompts, iCharIndex - 1) = "", False, Left(sPrompts, iCharIndex - 1)), Boolean)
							Case ExpressionParameter.ValueDate
								sTemp = Left(sPrompts, iCharIndex - 1)
								If sTemp = "null" Or sTemp = "12/30/1899" Or sTemp = "" Then
									dtValueDate = Nothing
								Else
									dtValueDate = DateTime.Parse(Left(sPrompts, iCharIndex - 1), CultureInfo.CreateSpecificCulture("en-US"))
								End If
							Case ExpressionParameter.PromptDateType
								iParameterIndex = 0
								iPromptDateType = CType(IIf(Left(sPrompts, iCharIndex - 1) = "", 0, Left(sPrompts, iCharIndex - 1)), Integer)
						
								' Got all of the required prompt paramters, so display it.
								If iPromptCount = 0 Then
									Response.Write("<h3>Prompted Values</h3>" & vbCrLf)
								End If
								
								iPromptCount = iPromptCount + 1
								
								Response.Write("<div class='formField'>" & vbCrLf)
								
								If iValueType = ExpressionValueTypes.giEXPRVALUE_LOGIC Then
									Response.Write(String.Format("<label for='prompt_3_{0}' class='checkbox' tabindex='0' style='width: 40%;'>", sNodeKey) & vbCrLf)
								Else
									Response.Write("<label style='width: 40%;'>")
								End If
						
								Response.Write(HttpUtility.HtmlEncode(sPromptDescription) & vbCrLf)
								Response.Write("</label>" & vbCrLf)
																					
								' Character Prompted Value
								If iValueType = ExpressionValueTypes.giEXPRVALUE_CHARACTER Then
									Response.Write("<input type='text' class='text' id='prompt_1_" & sNodeKey & "' name='prompt_1_" & sNodeKey & "' value='" & Replace(sValueCharacter, """", "&quot;") & "' maxlength='" & iPromptSize & "' style='width: 58%;'>" & vbCrLf)
									Response.Write("<input type='hidden' id='promptMask_" & sNodeKey & "' name='promptMask_" & sNodeKey & "' value='" & Replace(sPromptMask, """", "&quot;") & "'>" & vbCrLf)

									' Numeric Prompted Value
								ElseIf iValueType = ExpressionValueTypes.giEXPRVALUE_NUMERIC Then
									Response.Write(String.Format("<input type='text' class='text' id='prompt_2_{0}' name='prompt_2_{0}' value='{1}' style='width: 58%;'>", sNodeKey, Replace(dblValueNumeric.ToString, ".", Session("LocaleDecimalSeparator").ToString)) & vbCrLf)
									Response.Write(String.Format("<input type='hidden' id='promptSize_{0}' name='promptSize_{0}' value='{1}'>", sNodeKey, iPromptSize) & vbCrLf)
									Response.Write(String.Format("<input type='hidden' id='promptDecs_{0}' name='promptDecs_{0}' value='{1}'>", sNodeKey, iPromptDecimals) & vbCrLf)

									' Logic Prompted Value
								ElseIf iValueType = ExpressionValueTypes.giEXPRVALUE_LOGIC Then
									Response.Write(String.Format("<input type='checkbox' id='prompt_3_{0}' name='prompt_3_{0}' style='width: 1em;' onclick='checkboxClick(""{0}"")'", sNodeKey))
									If fValueLogic Then
										Response.Write(" checked/>" & vbCrLf)
									Else
										Response.Write("/>" & vbCrLf)
									End If
							
									Response.Write(String.Format("<input type='hidden' id='promptChk_{0}' name='promptChk_{0}' value='", sNodeKey))
									If fValueLogic Then
										Response.Write("TRUE'>" & vbCrLf)
									Else
										Response.Write("FALSE'>" & vbCrLf)
									End If
							 
									' Date Prompted Value
								ElseIf iValueType = ExpressionValueTypes.giEXPRVALUE_DATE Then
									Response.Write(String.Format("<input type='text' class='datepicker' id='prompt_4_{0}' name='prompt_4_{0}' value='", sNodeKey))
																												
									Select Case iPromptDateType
										Case PromptedDateType.Explicit
											' Explicit value
											If Not dtValueDate = Nothing Then
												Response.Write(ConvertSQLDateToLocale(dtValueDate))
											End If
																		
										Case PromptedDateType.Current
											' Current date
											Response.Write(ConvertSQLDateToLocale(Date.Now))
										Case PromptedDateType.MonthStart
											' Start of current month
											iDay = (Day(Date.Now) * -1) + 1
											dtDate = DateAdd("d", iDay, Date.Now)
											Response.Write(ConvertSQLDateToLocale(dtDate))
										Case PromptedDateType.MonthEnd
											' End of current month
											iDay = (Day(Date.Now) * -1) + 1
											dtDate = DateAdd("d", iDay, Date.Now)
											dtDate = DateAdd("m", 1, dtDate)
											dtDate = DateAdd("d", -1, dtDate)
											Response.Write(ConvertSQLDateToLocale(dtDate))
										Case PromptedDateType.YearStart
											' Start of current year
											iDay = (Day(Date.Now) * -1) + 1
											iMonth = (Month(Date.Now) * -1) + 1
											dtDate = DateAdd("d", iDay, Date.Now)
											dtDate = DateAdd("m", iMonth, dtDate)
											Response.Write(ConvertSQLDateToLocale(dtDate))
										Case PromptedDateType.YearEnd
											' End of current year
											iDay = (Day(Date.Now) * -1) + 1
											iMonth = (Month(Date.Now) * -1) + 1
											dtDate = DateAdd("d", iDay, Date.Now)
											dtDate = DateAdd("m", iMonth, dtDate)
											dtDate = DateAdd("yyyy", 1, dtDate)
											dtDate = DateAdd("d", -1, dtDate)
											Response.Write(ConvertSQLDateToLocale(dtDate))
									End Select
									Response.Write("' style='width: 58%;'>" & vbCrLf)

									' Lookup Prompted Value
								ElseIf iValueType = ExpressionValueTypes.giEXPRVALUE_TABLEVALUE Then
									Response.Write(String.Format("<select id='promptLookup_{0}' class='combo' name='promptLookup_{0}' style='width: 58%;' onchange=""comboChange('{0}')"">", sNodeKey) & vbCrLf)

									fDefaultFound = False
									fFirstValueDone = False
									sFirstValue = ""

									' Get the lookup values.
									rstLookupValues = GetLookupValues(CInt(CleanNumeric(lngColumnID)))
									For Each rowLookupValues In rstLookupValues.Rows
										Response.Write("<option")

										If Not fFirstValueDone Then
											sFirstValue = rowLookupValues(0).ToString
											fFirstValueDone = True
										End If
								
										If rstLookupValues.Columns(0).DataType = GetType(DateTime) Then
											' Field is a date so format as such.
											sOptionValue = ConvertSQLDateToLocale(rowLookupValues(0).ToString)
											If sOptionValue = ConvertSQLDateToLocale(sValueCharacter) Then
												Response.Write(" selected")
												fDefaultFound = True
											End If
											Response.Write(">" & sOptionValue & "</option>" & vbCrLf)
										ElseIf IsDataColumnDecimal(rstLookupValues.Columns(0)) Then
											' Field is a numeric so format as such.
											sOptionValue = Replace(rowLookupValues(0).ToString, ".", Session("LocaleDecimalSeparator").ToString)
											If (Not IsDBNull(rowLookupValues(0).ToString)) And (Not IsDBNull(sValueCharacter)) Then
												If FormatNumber(rowLookupValues(0).ToString) = FormatNumber(sValueCharacter) Then
													Response.Write(" selected")
													fDefaultFound = True
												End If
											End If
											Response.Write(">" & sOptionValue & "</option>" & vbCrLf)
										ElseIf rstLookupValues.Columns(0).DataType = GetType(Boolean) Then
											' Field is a logic so format as such.
											sOptionValue = rowLookupValues(0).ToString
											If sOptionValue = sValueCharacter Then
												Response.Write(" selected")
												fDefaultFound = True
											End If
											Response.Write(">" & sOptionValue & "</option>" & vbCrLf)
										Else
											sOptionValue = RTrim(rowLookupValues(0).ToString)
											If sOptionValue = sValueCharacter Then
												Response.Write(" selected")
												fDefaultFound = True
											End If
											Response.Write(">" & HttpUtility.HtmlEncode(sOptionValue) & "</option>" & vbCrLf)
										End If
									Next

									Response.Write("</select>" & vbCrLf)

									If fDefaultFound Then
										sDefaultValue = sValueCharacter
									Else
										sDefaultValue = sFirstValue
									End If

									If rstLookupValues.Columns(0).DataType = GetType(DateTime) Then
										' Date.
										Response.Write(String.Format("<input type='hidden' id='prompt_4_{0}' name='prompt_4_{0}' value='{1}'>", sNodeKey, ConvertSQLDateToLocale(sDefaultValue)) & vbCrLf)
									ElseIf IsDataColumnDecimal(rstLookupValues.Columns(0)) Then
										' Numeric
										Response.Write(String.Format("<input type='hidden' id='prompt_2_{0}' name='prompt_2_{0}' value='{1}'>", sNodeKey, Replace(sDefaultValue, ".", Session("LocaleDecimalSeparator").ToString)) & vbCrLf)
									ElseIf rstLookupValues.Columns(0).DataType = GetType(Boolean) Then
										' Logic
										Response.Write(String.Format("<input type='hidden' id='prompt_3_{0}' name='prompt_3_{0}' value='{1}'>", sNodeKey, sDefaultValue) & vbCrLf)
									Else
										Response.Write(String.Format("<input type='hidden' id='prompt_1_{0}' name='prompt_1_{0}' value='{1}'>", sNodeKey, Html.Encode(sDefaultValue)) & vbCrLf)
									End If
								End If
								
								Response.Write("</div>" & vbCrLf)
								
								
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

						rstPromptedValue = objDataAccess.GetDataTable("spASRIntGetUtilityPromptedValues", CommandType.StoredProcedure, _
									New SqlParameter("piUtilType", SqlDbType.Int) With {.Value = UtilityType.utlFilter}, _
									New SqlParameter("piUtilTableID", SqlDbType.Int) With {.Value = 0}, _
									New SqlParameter("piUtilID", SqlDbType.Int) With {.Value = CleanNumeric(CLng(sFilterCalcID))}, _
									New SqlParameter("piRecordID", SqlDbType.Int) With {.Value = 0})
							

						If rstPromptedValue.Rows.Count > 0 Then
							If iPromptCount = 0 Then
								Response.Write("<h3>Prompted Values</h3>" & vbCrLf)
							End If
					
							For Each rowPromptedValues1 As DataRow In rstPromptedValue.Rows
								
								Dim componentID = rowPromptedValues1("componentID").ToString
								
								iPromptCount = iPromptCount + 1
								
								Response.Write("<div class='formField'>" & vbCrLf)
								
								If NullSafeInteger(rowPromptedValues1("ValueType")) = 3 Then
									Response.Write(String.Format("<label for='prompt_3_C{0}' class='checkbox' tabindex='0' style='width: 40%;'>", componentID))
								Else
									Response.Write("<label style='width: 40%;'>")
								End If

								Response.Write(rowPromptedValues1("PromptDescription").ToString & vbCrLf)
								Response.Write("</label>" & vbCrLf)

								' Character Prompted Value
								If rowPromptedValues1("ValueType") = 1 Then
									Response.Write(String.Format("<input type='text' class='text' id='prompt_1_C{0}' name='prompt_1_C{0}' value='{1}' maxlength='" & rowPromptedValues1("promptsize").ToString & "' style='width: 58%;'>", componentID, Html.Encode(rowPromptedValues1("valuecharacter").ToString)) & vbCrLf)
									Response.Write(String.Format("<input type='hidden' id='promptMask_C{0}' name='promptMask_C{0}' value='{1}'>", componentID, Html.Encode(rowPromptedValues1("promptMask").ToString)) & vbCrLf)

									' Numeric Prompted Value
								ElseIf rowPromptedValues1("ValueType") = 2 Then
									Response.Write(String.Format("<input type='text' class='text' id='prompt_2_C{0}' name='prompt_2_C{0}' value='{1}' style='width: 58%;'>", componentID, Replace(rowPromptedValues1("valuenumeric").ToString, ".", Session("LocaleDecimalSeparator").ToString)) & vbCrLf)
									Response.Write(String.Format("<input type='hidden' id='promptSize_C{0}' name='promptSize_C{0}' value='{1}'>", componentID, rowPromptedValues1("promptSize").ToString) & vbCrLf)
									Response.Write(String.Format("<input type='hidden' id='promptDecs_C{0}' name='promptDecs_C{0}' value='{1}'>", componentID, rowPromptedValues1("promptDecimals").ToString) & vbCrLf)

									' Logic Prompted Value
								ElseIf rowPromptedValues1("ValueType") = 3 Then
									Response.Write(String.Format("<input type='checkbox' tabindex='-1' id='prompt_3_C{0}' name='prompt_3_C{0}' style='width: 1em;' onclick='checkboxClick(""C{0}"")'", componentID))
									If rowPromptedValues1("valuelogic") Then
										Response.Write(" checked/>" & vbCrLf)
									Else
										Response.Write("/>" & vbCrLf)
									End If
							
									Response.Write(String.Format("<input type='hidden' id='promptChk_C{0}' name='promptChk_C{0}' value='{1}'>", componentID, rowPromptedValues1("valuelogic").ToString) & vbCrLf)
											 
									' Date Prompted Value
								ElseIf rowPromptedValues1("ValueType") = 4 Then
														
									Response.Write(String.Format("<input type='text' class='text' id='prompt_4_C{0}' name='prompt_4_C{0}' value='", componentID))
									Select Case rowPromptedValues1("promptDateType")
										Case 0
											' Explicit value
											If Not IsDBNull(rowPromptedValues1("valuedate").ToString) Then
												If (CStr(rowPromptedValues1("valuedate").ToString) <> "00:00:00") And _
														(CStr(rowPromptedValues1("valuedate").ToString) <> "12:00:00 AM") Then
													Response.Write(ConvertSQLDateToLocale(rowPromptedValues1("valuedate").ToString))
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
									Response.Write("' style='width: 58%;'>" & vbCrLf)

									' Lookup Prompted Value
								ElseIf rowPromptedValues1("ValueType") = 5 Then
									Response.Write(String.Format("<select id='promptLookup_C{0}' name='promptLookup_C{0}' class='combo' style='width: 58%;' onchange=""comboChange('C{0}')"">", componentID) & vbCrLf)

									fDefaultFound = False
									fFirstValueDone = False
									sFirstValue = ""

									' Get the lookup values.
									rstLookupValues = GetLookupValues(NullSafeInteger(rowPromptedValues1("fieldColumnID")))

									For Each rowLookupValues In rstLookupValues.Rows
										Response.Write("<option")

										If Not fFirstValueDone Then
											sFirstValue = rowLookupValues(0).ToString
											fFirstValueDone = True
										End If

										If rstLookupValues.Columns(0).DataType = GetType(DateTime) Then
											' Field is a date so format as such.
											sOptionValue = ConvertSQLDateToLocale(rowLookupValues(0).ToString)
											If sOptionValue = ConvertSQLDateToLocale(rowPromptedValues1("valuecharacter").ToString) Then
												Response.Write(" selected")
												fDefaultFound = True
											End If
											Response.Write(">" & sOptionValue & "</option>" & vbCrLf)
										ElseIf IsDataColumnDecimal(rstLookupValues.Columns(0)) Then
											' Field is a numeric so format as such.
											sOptionValue = Replace(rowLookupValues(0).ToString, ".", Session("LocaleDecimalSeparator").ToString)
											If (Not IsDBNull(rowLookupValues(0))) And (Not IsDBNull(rstPromptedValue("valuecharacter"))) Then
												If FormatNumber(rowLookupValues(0)) = FormatNumber(rstPromptedValue("valuecharacter")) Then
													Response.Write(" selected")
													fDefaultFound = True
												End If
											End If
											Response.Write(">" & sOptionValue & "</option>" & vbCrLf)
										ElseIf rstLookupValues.Columns(0).DataType = GetType(Boolean) Then
											' Field is a logic so format as such.
											sOptionValue = rowLookupValues(0).ToString
											If sOptionValue = rowPromptedValues1("valuecharacter").ToString Then
												Response.Write(" selected")
												fDefaultFound = True
											End If
											Response.Write(">" & sOptionValue & "</option>" & vbCrLf)
										Else
											sOptionValue = RTrim(rowLookupValues(0).ToString)
											If sOptionValue = rowPromptedValues1("valuecharacter").ToString Then
												Response.Write(" selected")
												fDefaultFound = True
											End If
											Response.Write(">" & sOptionValue & "</option>" & vbCrLf)
										End If
									Next

									Response.Write("</select>" & vbCrLf)

									If fDefaultFound Then
										sDefaultValue = rowPromptedValues1("valuecharacter").ToString
									Else
										sDefaultValue = sFirstValue
									End If

									If rstLookupValues.Columns(0).DataType = GetType(DateTime) Then
										' Date.
										Response.Write(String.Format("<input type='hidden' id='prompt_4_C{0}' name='prompt_4_C{0}' value='{1}'>", componentID, ConvertSQLDateToLocale(sDefaultValue)) & vbCrLf)
									ElseIf IsDataColumnDecimal(rstLookupValues.Columns(0)) Then
										' Numeric
										Response.Write(String.Format("<input type='hidden' id='prompt_2_C{0}' name='prompt_2_C{0}' value='{1}'>", componentID, Replace(sDefaultValue, ".", Session("LocaleDecimalSeparator").ToString)) & vbCrLf)
									ElseIf rstLookupValues.Columns(0).DataType = GetType(Boolean) Then
										' Logic
										Response.Write(String.Format("<input type='hidden' id='prompt_3_C{0}' name='prompt_3_C{0}' value='{1}'>", componentID, sDefaultValue) & vbCrLf)
									Else
										Response.Write(String.Format("<input type='hidden' id='prompt_1_C{0}' name='prompt_1_C{0}' value='{1}'>", componentID, Html.Encode(sDefaultValue)) & vbCrLf)
									End If
								End If
								
								Response.Write("</div>" & vbCrLf)
								
							Next
						End If
					End If
				Loop
			End If

			If iPromptCount > 0 Then
				Response.Write("<br/><br/>" & vbCrLf)
				Response.Write("<input type='button' value='Cancel' name='Cancel' class='btn' value='Cancel' style='width: 80px; float: right;' onclick='ute_cancelClick();' />" & vbCrLf)
				Response.Write("<input type='button' value='OK' name='Submit' class='btn' style='width: 80px; float: right; margin-right: 10px;' onclick='SubmitPrompts();' />" & vbCrLf)
			End If

			Response.Write(String.Format("<input type='hidden' id='txtPromptCount' name='txtPromptCount' value='{0}'>", iPromptCount) & vbCrLf)
		%>
		<input type="hidden" id="type" name="type" value="<%:Model.type%>" />
		<input type="hidden" id="components1" name="components1" value="<%:Model.components1%>" />
		<input type="hidden" id="tableID" name="tableID" value="<%:Model.tableID%>" />
	</form>

</div>


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

	
</script>
