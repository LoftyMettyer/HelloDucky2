<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<!DOCTYPE html>

<% 	
	Session("filterID") = Request.Form("filterID")
%>

<html>
<head>

	<script src="<%: Url.LatestContent("~/bundles/jQuery")%>" type="text/javascript"></script>
	<script src="<%: Url.LatestContent("~/bundles/jQueryUI7")%>" type="text/javascript"></script>
	<script src="<%: Url.LatestContent("~/bundles/OpenHR_General")%>" type="text/javascript"></script>
	<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />

	<title>OpenHR</title>

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
				// sMaskCtlName = sMaskCtlName.toUpperCase();

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

</head>

<body <%=session("BodyColour")%> leftmargin="20" topmargin="20" bottommargin="20" rightmargin="5">
	
	<form name="frmPromptedValues" id="frmPromptedValues" method="POST" action="promptedValues_submit">

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
							For Each rowPromptedValues As DataRow In rstPromptedValue.Rows
								iPromptCount = iPromptCount + 1
						%>
						<tr height="10">
							<td width="20" height="10"></td>
							<td nowrap height="10">
								<%
									If NullSafeString(rowPromptedValues("ValueType")) = "3" Then
								%>
								<label
									for="prompt_3_<%=rowPromptedValues("componentID").ToString%>"
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

									<%=rowPromptedValues("PromptDescription").ToString%>
									<%
										If NullSafeString(rowPromptedValues("ValueType").ToString) = "3" Then
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
									If NullSafeString(rowPromptedValues("ValueType").ToString) = "1" Then
								%>
								<input type="text" class="text" id='prompt_1_<%=rowPromptedValues("componentID").ToString%>' name='prompt_1_<%=rowPromptedValues("componentID").ToString%>'
									value="<%=Replace(rowPromptedValues("valuecharacter").ToString, """", "&quot;")%>" maxlength='<%=rowPromptedValues("promptsize").ToString%>'
									style="WIDTH: 100%">
								<input type="hidden" id='promptMask_<%=rowPromptedValues("componentID").ToString%>' name='promptMask_<%=rowPromptedValues("componentID").ToString%>'
									value="<%=Replace(rowPromptedValues("promptMask").ToString, """", "&quot;")%>">
								<%
									' Numeric Prompted Value
								ElseIf NullSafeString(rowPromptedValues("ValueType").ToString) = "2" Then
								%>
								<input type="text" class="text" id='prompt_2_<%=rowPromptedValues("componentID").ToString%>' name='prompt_2_<%=rowPromptedValues("componentID").ToString%>'
									value="<%=Replace(rowPromptedValues("valuenumeric").ToString, ".", Session("LocaleDecimalSeparator"))%>"
									style="WIDTH: 100%">
								<input type="hidden" id='promptSize_<%=rowPromptedValues("componentID").ToString%>' name='promptSize<%=rowPromptedValues("componentID").ToString%>'
									value="<%=rowPromptedValues("promptSize").ToString%>">
								<input type="hidden" id='promptDecs_<%=rowPromptedValues("componentID").ToString%>' name='promptDecs<%=rowPromptedValues("componentID").ToString%>'
									value="<%=rowPromptedValues("promptDecimals").ToString%>">
								<%
									' Logic Prompted Value
								ElseIf NullSafeString(rowPromptedValues("ValueType").ToString) = "3" Then
								%>
								<input type="checkbox" id='prompt_3_<%=rowPromptedValues("componentID").ToString%>' name='prompt_3_<%=rowPromptedValues("componentID").ToString%>'
									<%If rowPromptedValues("valuelogic").ToString Then%> checked <%End If%>
									onclick="checkboxClick(<%=rowPromptedValues("componentID").ToString%>)" />
								<input type="hidden" id='promptChk_<%=rowPromptedValues("componentID").ToString%>' name='promptChk_<%=rowPromptedValues("componentID").ToString%>'
									value='<%=rowPromptedValues("valuelogic").ToString%>'>
								<%			  
									' Date Prompted Value
								ElseIf NullSafeString(rowPromptedValues("ValueType").ToString) = "4" Then

									Response.Write("        <input type=text class=""text"" id=prompt_4_" & rowPromptedValues("componentID").ToString & " name=prompt_4_" & rowPromptedValues("componentID").ToString & " value=""")
	
									Dim iDay As Integer, iMonth As Integer, dtDate As DateTime
	
									Select Case rowPromptedValues("promptDateType")
										Case 0
											' Explicit value
											Response.Write(ConvertSQLDateToLocale(rowPromptedValues("valuedate").ToString))
										Case 1
											' Current date
											Response.Write(ConvertSQLDateToLocale(Now()))
										Case 2
											' Start of current month
											iDay = (Day(Now()) * -1) + 1
											dtDate = DateAdd("d", iDay, Now())
											Response.Write(ConvertSQLDateToLocale(dtDate))
										Case 3
											' End of current month
											iDay = (Day(Now()) * -1) + 1
											dtDate = DateAdd("d", iDay, Now())
											dtDate = DateAdd("m", 1, dtDate)
											dtDate = DateAdd("d", -1, dtDate)
											Response.Write(ConvertSQLDateToLocale(dtDate))
										Case 4
											' Start of current year
											iDay = (Day(Now()) * -1) + 1
											iMonth = (Month(Now()) * -1) + 1
											dtDate = DateAdd("d", iDay, Now())
											dtDate = DateAdd("m", iMonth, dtDate)
											Response.Write(ConvertSQLDateToLocale(dtDate))
										Case 5
											' End of current year
											iDay = (Day(Now()) * -1) + 1
											iMonth = (Month(Now()) * -1) + 1
											dtDate = DateAdd("d", iDay, Now())
											dtDate = DateAdd("m", iMonth, dtDate)
											dtDate = DateAdd("yyyy", 1, dtDate)
											dtDate = DateAdd("d", -1, dtDate)
											Response.Write(ConvertSQLDateToLocale(dtDate))
									End Select
									Response.Write(""" style=""WIDTH: 100%"">" & vbCrLf)

									' Lookup Prompted Value
								ElseIf NullSafeString(rowPromptedValues("ValueType").ToString) = "5" Then
									Response.Write("        		<SELECT id=promptLookup_" & rowPromptedValues("componentID").ToString & " name=promptLookup_" & rowPromptedValues("componentID").ToString & " style=""WIDTH: 100%"" class=""combo"" onchange=""comboChange(" & rowPromptedValues("componentID").ToString & ")"">" & vbCrLf)

									fDefaultFound = False
									fFirstValueDone = False
									sFirstValue = ""
					
									Dim rstLookupValues = ASRIntranetFunctions.GetLookupValues(CInt(CleanNumeric(rowPromptedValues("fieldColumnID").ToString)))

									For Each rowLookupValues As DataRow In rstLookupValues.Rows
										Response.Write("        		  <OPTION")
					
										If Not fFirstValueDone Then
											sFirstValue = rowLookupValues(0).ToString
											fFirstValueDone = True
										End If
						
										If rstLookupValues.Columns(0).DataType = GetType(System.DateTime) Then
											' Field is a date so format as such.
											Dim sOptionValue = ConvertSQLDateToLocale(rowLookupValues(0).ToString)
											If sOptionValue = ConvertSQLDateToLocale(rowPromptedValues("valuecharacter").ToString) Then
												Response.Write(" SELECTED")
												fDefaultFound = True
											End If
											Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
										ElseIf GeneralUtilities.IsDataColumnDecimal(rstLookupValues.Columns(0)) Then
											' Field is a numeric so format as such.
											Dim sOptionValue = Replace(rowLookupValues(0).ToString, ".", Session("LocaleDecimalSeparator"))
											If (Not IsDBNull(rstLookupValues(0))) And (Not IsDBNull(rowPromptedValues("valuecharacter").ToString)) Then
												If FormatNumber(rowLookupValues(0)) = FormatNumber(rowPromptedValues("valuecharacter").ToString) Then
													Response.Write(" SELECTED")
													fDefaultFound = True
												End If
											End If
											Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
										ElseIf rstLookupValues.Columns(0).DataType = GetType(System.Boolean) Then
											' Field is a logic so format as such.
											Dim sOptionValue As String = rowLookupValues(0).ToString
											If sOptionValue = rowPromptedValues("valuecharacter").ToString Then
												Response.Write(" SELECTED")
												fDefaultFound = True
											End If
											Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
										Else
											Dim sOptionValue As String = rowLookupValues(0).ToString
											If sOptionValue = rowPromptedValues("valuecharacter").ToString Then
												Response.Write(" SELECTED")
												fDefaultFound = True
											End If
											Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
										End If
									Next

									Response.Write("   		     </SELECT>" & vbCrLf)

									If fDefaultFound Then
										sDefaultValue = rowPromptedValues("valuecharacter").ToString
									Else
										sDefaultValue = sFirstValue
									End If
					
									If rstLookupValues.Columns(0).DataType = GetType(System.DateTime) Then
										' Date.
										Response.Write("        <input type=hidden id=prompt_4_" & rowPromptedValues("componentID").ToString & " name=prompt_4_" & rowPromptedValues("componentID").ToString & " value=" & ConvertSQLDateToLocale(sDefaultValue) & ">" & vbCrLf)
									ElseIf GeneralUtilities.IsDataColumnDecimal(rstLookupValues.Columns(0)) Then
										' Numeric
										Response.Write("        <input type=hidden id=prompt_2_" & rowPromptedValues("componentID").ToString & " name=prompt_2_" & rowPromptedValues("componentID").ToString & " value=" & Replace(sDefaultValue, ".", Session("LocaleDecimalSeparator")) & ">" & vbCrLf)
									ElseIf rstLookupValues.Columns(0).DataType = GetType(System.Boolean) Then
										' Logic
										Response.Write("        <input type=hidden id=prompt_3_" & rowPromptedValues("componentID").value & " name=prompt_3_" & rowPromptedValues("componentID").value & " value=" & sDefaultValue & ">" & vbCrLf)
									Else
										Response.Write("        <input type=hidden id=prompt_1_" & rowPromptedValues("componentID").value & " name=prompt_1_" & rowPromptedValues("componentID").value & " value=""" & Replace(sDefaultValue, """", "&quot;") & """>" & vbCrLf)
									End If
								End If
				
								Response.Write("					</td>" & vbCrLf)
								Response.Write("					<td width=20 height=10></td>" & vbCrLf)
								Response.Write("				</tr>" & vbCrLf)
							Next 

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
									onclick="SubmitPrompts()" />
							</td>
							<td width="20"></td>
							<td width="80">
								<input type="button" class="btn" name="Cancel" value="Cancel" style="WIDTH: 80px"
									onclick="CancelClick()" />
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

								Response.Write("<input type=""hidden"" id=""txtPromptCount"" name=""txtPromptCount"" value=" & iPromptCount & ">" & vbCrLf)
								%>

								<input type="hidden" id="filterID" name="filterID" value="<%=Session("filterID")%>">
	</form>
	
	<script type="text/javascript"> promptedValues_onload();</script>
	

</body>
</html>


