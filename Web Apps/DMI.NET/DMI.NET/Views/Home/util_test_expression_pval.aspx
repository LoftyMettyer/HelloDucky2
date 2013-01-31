<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage"%>
<%@ Import Namespace="DMI.NET" %>

<!DOCTYPE html>

<link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css" />
<script src="<%: Url.Content("~/Scripts/jquery-1.8.2.js") %>" type="text/javascript"></script>
<script src="<%: Url.Content("~/Scripts/openhr.js") %>" type="text/javascript"></script>
<script src="<%: Url.Content("~/Scripts/ctl_SetFont.js") %>" type="text/javascript"></script>
<script src="<%: Url.Content("~/Scripts/ctl_SetStyles.js") %>" type="text/javascript"></script>
<script src="<%: Url.Content("~/Scripts/jquery-ui-1.9.1.custom.min.js") %>" type="text/javascript"></script>
<script src="<%: Url.Content("~/Scripts/jquery.cookie.js") %>" type="text/javascript"></script>
<script src="<%: Url.Content("~/Scripts/menu.js")%>" type="text/javascript"></script>
<script src="<%: Url.Content("~/Scripts/jquery.ui.touch-punch.min.js") %>" type="text/javascript"></script>
<script src="<%: Url.Content("~/Scripts/jsTree/jquery.jstree.js") %>" type="text/javascript"></script>
<script id="officebarscript" src="<%: Url.Content("~/Scripts/officebar/jquery.officebar.js") %>" type="text/javascript"></script>

<html>
<head runat="server">
    <title>OpenHR Intranet</title>

    <object
        classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
        id="Microsoft_Licensed_Class_Manager_1_0"
        viewastext>
        <param name="LPKPath" value="lpks/main.lpk">
    </object>

    <script type="text/javascript" >
<!--
    function util_test_expression_pval_onload() {

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
            iResizeBy = frmPromptedValues.offsetParent.scrollHeight	- frmPromptedValues.offsetParent.clientHeight;
            if (frmPromptedValues.offsetParent.offsetHeight + iResizeBy > screen.height) {
                window.parent.dialogHeight = new String(screen.height) + "px";
            }
            else {
                iNewHeight = new Number(window.parent.dialogHeight.substr(0, window.parent.dialogHeight.length-2));
                iNewHeight = iNewHeight + iResizeBy;
                window.parent.dialogHeight = new String(iNewHeight) + "px";
            }
        }
    }

    function SubmitPrompts()
    {
        // Validate the prompt values before submitting the form.
        var controlCollection = frmPromptedValues.elements;
        if (controlCollection!=null) {
            for (i=0; i<controlCollection.length; i++)  {
                sControlName = controlCollection.item(i).name;
                sControlPrefix = sControlName.substr(0, 7);

                if (sControlPrefix=="prompt_") {

                    // Get the control's data type.
                    iType = new Number(sControlName.substring(7,8));
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
        sDecimalSeparator = sDecimalSeparator.concat(ASRIntranetFunctions.LocaleDecimalSeparator);
        var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");

        sThousandSeparator = "\\";
        sThousandSeparator = sThousandSeparator.concat(ASRIntranetFunctions.LocaleThousandSeparator);
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
            if (ASRIntranetFunctions.LocaleDecimalSeparator != ".") {
                // Remove decimal points.
                sConvertedValue = sConvertedValue.replace(rePoint, "A");
                // replace the locale decimal marker with the decimal point.
                sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
            }

            if (isNaN(sConvertedValue) == true) {
                fOK = false;
                ASRIntranetFunctions.MessageBox("Invalid numeric value entered.");
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
                    pctlPrompt.value = ASRIntranetFunctions.ConvertSQLDateToLocale(sValue);
                }
            }
			
            if (fOK == false) {
                ASRIntranetFunctions.MessageBox("Invalid date value entered.");
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
                    ASRIntranetFunctions.MessageBox("The entered value does not match the required format (" + sMask + ").");
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
		
        sDateFormat = ASRIntranetFunctions.LocaleDateFormat;

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
	
        sValue = ASRIntranetFunctions.ConvertSQLDateToLocale(sTempValue);

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
    -->
</script>

</head>
<body>
    
    <div data-framesource="util_test_expression_pval">

        <FORM name=frmPromptedValues id=frmPromptedValues method=POST action=util_test_expression>
<%
	dim iPromptCount
	dim sPrompts
	dim sNodeKey
	dim sPromptDescription
	dim iValueType
	dim iPromptSize
	dim iPromptDecimals
	dim sPromptMask
	dim lngTableID
	dim lngColumnID
	dim sValueCharacter
	dim dblValueNumeric
	dim fValueLogic
	dim dtValueDate
	dim iPromptDateType
	dim iCharIndex
	dim iParameterIndex
	dim sTemp
	dim sFiltersAndCalcs
	dim sFilterCalcID
	dim fDefaultFound 
	dim fFirstValueDone 
	dim sFirstValue

    Dim cmdDefn
    Dim iDay As Integer
    Dim iMonth As Integer
    Dim dtDate As Date
    Dim prmFilterID
    Dim rstPromptedValue
    Dim sOptionValue As String
    Dim sDefaultValue As String
    Dim cmdLookupValues
    Dim prmColumnID
    Dim rstLookupValues
    

	iPromptCount = 0
	sPrompts = Request.Form("prompts")
	sFiltersAndCalcs = Request.Form("filtersAndCalcs")
	
	if len(sPrompts) > 0 then
        Response.Write("<table align=center class=""outline"" cellPadding=5 cellSpacing=0>" & vbCrLf)
        Response.Write("  <tr>" & vbCrLf)
        Response.Write("	  <td>" & vbCrLf)
        Response.Write("			<table align=center class=""invisible"" cellspacing=0 cellpadding=0>" & vbCrLf)
        Response.Write("				<tr>" & vbCrLf)
        Response.Write("					<td colspan=5 align=center><H3 align=center>Prompted Values</H3></td>" & vbCrLf)
        Response.Write("				</tr>" & vbCrLf)

        iParameterIndex = 0
		do while len(sPrompts) > 0
			iCharIndex = instr(sPrompts, "	")
			iParameterIndex = iParameterIndex + 1

			if iCharIndex >= 0 then
				select case iParameterIndex
					case 1
						sNodeKey = left(sPrompts, iCharIndex - 1)
					case 2
						sPromptDescription = left(sPrompts, iCharIndex - 1)
					case 3
						iValueType = left(sPrompts, iCharIndex - 1)
					case 4
						iPromptSize = left(sPrompts, iCharIndex - 1)
					case 5
						iPromptDecimals = left(sPrompts, iCharIndex - 1)
					case 6
						sPromptMask = left(sPrompts, iCharIndex - 1)
					case 7
						lngTableID = left(sPrompts, iCharIndex - 1)
					case 8
						lngColumnID = left(sPrompts, iCharIndex - 1)
					case 9
						sValueCharacter = left(sPrompts, iCharIndex - 1)
					case 10
						dblValueNumeric = left(sPrompts, iCharIndex - 1)
					case 11
						fValueLogic = left(sPrompts, iCharIndex - 1)
					case 12
						dtValueDate = left(sPrompts, iCharIndex - 1)
					case 13
						iParameterIndex = 0
						iPromptDateType = left(sPrompts, iCharIndex - 1)
						
                        ' Got all of the required prompt paramters, so display it.
                        iPromptCount = iPromptCount + 1
                        Response.Write("    <tr>" & vbCrLf)
                        Response.Write("      <td width=20></td>" & vbCrLf)
                        Response.Write("      <td>" & vbCrLf)
						
                        If iValueType = 3 Then
                            Response.Write("      <label " & vbCrLf)
                            Response.Write("      for=""prompt_3_" & sNodeKey & vbCrLf)
                            Response.Write("      class=""checkbox""" & vbCrLf)
                            Response.Write("      tabindex=0 " & vbCrLf)
                            Response.Write("      onkeypress=""try{checkboxLabel_onKeyPress(this);}catch(e){}""" & vbCrLf)
                            Response.Write("      onmouseover=""try{checkboxLabel_onMouseOver(this);}catch(e){}""" & vbCrLf)
                            Response.Write("      onmouseout=""try{checkboxLabel_onMouseOut(this);}catch(e){}""" & vbCrLf)
                            Response.Write("      onfocus=""try{checkboxLabel_onFocus(this);}catch(e){}""" & vbCrLf)
                            Response.Write("      onblur=""try{checkboxLabel_onBlur(this);}catch(e){}"">" & vbCrLf)
                        End If
						
                        Response.Write("      " & sPromptDescription & vbCrLf)

                        If iValueType = 3 Then
                            Response.Write("</label>" & vbCrLf)
                        End If
						
                        Response.Write("      </td>" & vbCrLf)
                        Response.Write("      <td width=20></td>" & vbCrLf)
                        Response.Write("      <td width=200>" & vbCrLf)

                        ' Character Prompted Value
                        If iValueType = "1" Then
                            Response.Write("        <input type=text class=""text"" id=prompt_1_" & sNodeKey & " name=prompt_1_" & sNodeKey & " value=""" & Replace(sValueCharacter, """", "&quot;") & """ maxlength=" & iPromptSize & " style=""WIDTH: 100%"">" & vbCrLf)
                            Response.Write("        <input type=hidden id=promptMask_" & sNodeKey & " name=promptMask_" & sNodeKey & " value=""" & Replace(sPromptMask, """", "&quot;") & """>" & vbCrLf)

                            ' Numeric Prompted Value
                        ElseIf iValueType = 2 Then
                            Response.Write("        <input type=text class=""text"" id=prompt_2_" & sNodeKey & " name=prompt_2_" & sNodeKey & " value=""" & Replace(dblValueNumeric, ".", Session("LocaleDecimalSeparator")) & """ style=""WIDTH: 100%"">" & vbCrLf)
                            Response.Write("        <input type=hidden id=promptSize_" & sNodeKey & " name=promptSize_" & sNodeKey & " value=""" & iPromptSize & """>" & vbCrLf)
                            Response.Write("        <input type=hidden id=promptDecs_" & sNodeKey & " name=promptDecs_" & sNodeKey & " value=""" & iPromptDecimals & """>" & vbCrLf)

                            ' Logic Prompted Value
                        ElseIf iValueType = 3 Then
                            Response.Write("        <input type=""checkbox"" id=prompt_3_" & sNodeKey & " name=prompt_3_" & sNodeKey & vbCrLf)
                            Response.Write("            onclick=""checkboxClick('" & sNodeKey & "')""" & vbCrLf)
                            Response.Write("            onmouseover=""try{checkbox_onMouseOver(this);}catch(e){}""" & vbCrLf)
                            Response.Write("            onmouseout=""try{checkbox_onMouseOut(this);}catch(e){}""")
                            If fValueLogic Then
                                Response.Write(" CHECKED/>" & vbCrLf)
                            Else
                                Response.Write("/>" & vbCrLf)
                            End If
						  
                            Response.Write("        <input type=hidden id=promptChk_" & sNodeKey & " name=promptChk_" & sNodeKey & " value=")
                            If fValueLogic Then
                                Response.Write("""TRUE"">" & vbCrLf)
                            Else
                                Response.Write("""FALSE"">" & vbCrLf)
                            End If
							 
                            ' Date Prompted Value
                        ElseIf iValueType = 4 Then
                            Response.Write("        <input type=text class=""text"" id=prompt_4_" & sNodeKey & " name=prompt_4_" & sNodeKey & " value=""")
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
                            Response.Write(""" style=""WIDTH: 100%"">" & vbCrLf)

                            ' Lookup Prompted Value
                        ElseIf iValueType = 5 Then
                            Response.Write("        <SELECT id=promptLookup_" & sNodeKey & " class=""combo"" name=promptLookup_" & sNodeKey & " style=""WIDTH: 100%"" onchange=""comboChange('" & sNodeKey & "')"">" & vbCrLf)

							fDefaultFound = false
							fFirstValueDone = false
							sFirstValue = ""

							' Get the lookup values.
                            cmdLookupValues = CreateObject("ADODB.Command")
							cmdLookupValues.CommandText = "sp_ASRIntGetLookupValues"
							cmdLookupValues.CommandType = 4 ' Stored Procedure
                            cmdLookupValues.ActiveConnection = Session("databaseConnection")

                            prmColumnID = cmdLookupValues.CreateParameter("columnID", 3, 1)
                            cmdLookupValues.Parameters.Append(prmColumnID)
							prmColumnID.value = cleanNumeric(lngColumnID)

                            Err.Clear()
                            rstLookupValues = cmdLookupValues.Execute
                            Do While Not rstLookupValues.EOF
                                Response.Write("          <OPTION")

                                If Not fFirstValueDone Then
                                    sFirstValue = rstLookupValues.Fields(0).Value
                                    fFirstValueDone = True
                                End If
								
                                If rstLookupValues.fields(0).type = 135 Then
                                    ' Field is a date so format as such.
                                    sOptionValue = convertSQLDateToLocale2(rstLookupValues.Fields(0).Value)
                                    If sOptionValue = convertSQLDateToLocale(sValueCharacter) Then
                                        Response.Write(" SELECTED")
                                        fDefaultFound = True
                                    End If
                                    Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)

                                ElseIf rstLookupValues.fields(0).type = 131 Then
                                    ' Field is a numeric so format as such.
                                    sOptionValue = Replace(rstLookupValues.Fields(0).Value, ".", Session("LocaleDecimalSeparator"))
                                    If (Not IsDBNull(rstLookupValues.Fields(0).Value)) And (Not IsDBNull(sValueCharacter)) Then
                                        If FormatNumber(rstLookupValues.Fields(0).Value) = FormatNumber(sValueCharacter) Then
                                            Response.Write(" SELECTED")
                                            fDefaultFound = True
                                        End If
                                    End If
                                    Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
								elseif rstLookupValues.fields(0).type = 11 then
									' Field is a logic so format as such.
                                    sOptionValue = rstLookupValues.Fields(0).Value
                                    If sOptionValue = sValueCharacter Then
                                        Response.Write(" SELECTED")
                                        fDefaultFound = True
                                    End If
                                    Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
                                Else
                                    sOptionValue = RTrim(rstLookupValues.Fields(0).Value)
                                    If sOptionValue = sValueCharacter Then
                                        Response.Write(" SELECTED")
                                        fDefaultFound = True
                                    End If
                                    Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
                                End If

                                rstLookupValues.MoveNext()
                            Loop

                            Response.Write("        </SELECT>" & vbCrLf)

                            If fDefaultFound Then
                                sDefaultValue = sValueCharacter
                            Else
                                sDefaultValue = sFirstValue
                            End If

                            If rstLookupValues.fields(0).type = 135 Then
                                ' Date.
                                Response.Write("        <input type=hidden id=prompt_4_" & sNodeKey & " name=prompt_4_" & sNodeKey & " value=" & convertSQLDateToLocale(sDefaultValue) & ">" & vbCrLf)
                            ElseIf rstLookupValues.fields(0).type = 131 Then
                                ' Numeric
                                Response.Write("        <input type=hidden id=prompt_2_" & sNodeKey & " name=prompt_2_" & sNodeKey & " value=" & Replace(sDefaultValue, ".", Session("LocaleDecimalSeparator")) & ">" & vbCrLf)
                            ElseIf rstLookupValues.fields(0).type = 11 Then
                                ' Logic
                                Response.Write("        <input type=hidden id=prompt_3_" & sNodeKey & " name=prompt_3_" & sNodeKey & " value=" & sDefaultValue & ">" & vbCrLf)
                            Else
                                Response.Write("        <input type=hidden id=prompt_1_" & sNodeKey & " name=prompt_1_" & sNodeKey & " value=""" & Replace(sDefaultValue, """", "&quot;") & """>" & vbCrLf)
                            End If

							' Release the ADO recordset object.
							rstLookupValues.close
                            rstLookupValues = Nothing
						end if
				
                        Response.Write("					</td>" & vbCrLf)
                        Response.Write("					<td width=20 height=10></td>" & vbCrLf)
                        Response.Write("				</tr>" & vbCrLf)
                End Select

				sPrompts = mid(sPrompts, iCharIndex + 1)
			end if
		loop
	end if

    If Len(sFiltersAndCalcs) > 0 Then
        Do While Len(sFiltersAndCalcs) > 0
            iCharIndex = InStr(sFiltersAndCalcs, "	")

            If iCharIndex >= 0 Then
                sFilterCalcID = Left(sFiltersAndCalcs, iCharIndex - 1)
                sFiltersAndCalcs = Mid(sFiltersAndCalcs, iCharIndex + 1)

                cmdDefn = CreateObject("ADODB.Command")
                cmdDefn.CommandText = "sp_ASRIntGetFilterPromptedValuesRecordset"
                cmdDefn.CommandType = 4 ' Stored Procedure
                cmdDefn.ActiveConnection = Session("databaseConnection")

                prmFilterID = cmdDefn.CreateParameter("filterID", 3, 1) ' 3=integer, 1=input
                cmdDefn.Parameters.Append(prmFilterID)
                prmFilterID.value = cleanNumeric(CLng(sFilterCalcID))

                Err.Clear()
                rstPromptedValue = cmdDefn.Execute

                If Not (rstPromptedValue.EOF And rstPromptedValue.BOF) Then
                    If iPromptCount = 0 Then
                        Response.Write("<table align=center class=""outline"" cellPadding=5 cellSpacing=0>" & vbCrLf)
                        Response.Write("  <tr>" & vbCrLf)
                        Response.Write("	  <td>" & vbCrLf)
                        Response.Write("			<table align=center class=""invisible"" cellspacing=0 cellpadding=0>" & vbCrLf)
                        Response.Write("				<tr>" & vbCrLf)
                        Response.Write("					<td colspan=5 align=center><H3 align=center>Prompted Values</H3></td>" & vbCrLf)
                        Response.Write("				</tr>" & vbCrLf)
                    End If
					
                    Do While Not rstPromptedValue.EOF
                        iPromptCount = iPromptCount + 1
                        Response.Write("				<tr height=10>" & vbCrLf)
                        Response.Write("					<td width=20 height=10></td>" & vbCrLf)
                        Response.Write("					<td nowrap height=10>" & vbCrLf)
						
                        If iValueType = 3 Then
                            Response.Write("          <label " & vbCrLf)
                            Response.Write("            for=""prompt_3_C" & rstPromptedValue.fields("componentID").value & vbCrLf)
                            Response.Write("            class=""checkbox""" & vbCrLf)
                            Response.Write("            tabindex=0 " & vbCrLf)
                            Response.Write("            onkeypress=""try{checkboxLabel_onKeyPress(this);}catch(e){}""" & vbCrLf)
                            Response.Write("            onmouseover=""try{checkboxLabel_onMouseOver(this);}catch(e){}""" & vbCrLf)
                            Response.Write("            onmouseout=""try{checkboxLabel_onMouseOut(this);}catch(e){}""" & vbCrLf)
                            Response.Write("            onfocus=""try{checkboxLabel_onFocus(this);}catch(e){}""" & vbCrLf)
                            Response.Write("            onblur=""try{checkboxLabel_onBlur(this);}catch(e){}"">" & vbCrLf)
                        End If

                        Response.Write("					  " & rstPromptedValue.fields("PromptDescription").value & vbCrLf)

                        If iValueType = 3 Then
                            Response.Write("          </label>" & vbCrLf)
                        End If

                        Response.Write("					</td>" & vbCrLf)
                        Response.Write("					<td width=20 height=10>&nbsp;</td>" & vbCrLf)
                        Response.Write("   		    <td width=200 height=10>" & vbCrLf)

                        ' Character Prompted Value
                        If rstPromptedValue.fields("ValueType").value = 1 Then
                            Response.Write("    		    <input type=text class=""text"" id=prompt_1_C" & rstPromptedValue.fields("componentID").value & " name=prompt_1_C" & rstPromptedValue.fields("componentID").value & " value=""" & Replace(rstpromptedvalue.fields("valuecharacter").value, """", "&quot;") & """ maxlength=" & rstPromptedValue.fields("promptsize").value & " style=""WIDTH: 100%"">" & vbCrLf)
                            Response.Write("    		    <input type=hidden id=promptMask_C" & rstPromptedValue.fields("componentID").value & " name=promptMask_C" & rstPromptedValue.fields("componentID").value & " value=""" & Replace(rstpromptedvalue.fields("promptMask").value, """", "&quot;") & """>" & vbCrLf)

                            ' Numeric Prompted Value
                        ElseIf rstPromptedValue.fields("ValueType").value = 2 Then
                            Response.Write("     		   <input type=text class=""text"" id=prompt_2_C" & rstPromptedValue.fields("componentID").value & " name=prompt_2_C" & rstPromptedValue.fields("componentID").value & " value=""" & Replace(rstpromptedvalue.fields("valuenumeric").value, ".", Session("LocaleDecimalSeparator")) & """ style=""WIDTH: 100%"">" & vbCrLf)
                            Response.Write("     		   <input type=hidden id=promptSize_C" & rstPromptedValue.fields("componentID").value & " name=promptSize_C" & rstPromptedValue.fields("componentID").value & " value=""" & rstpromptedvalue.fields("promptSize").value & """>" & vbCrLf)
                            Response.Write("     		   <input type=hidden id=promptDecs_C" & rstPromptedValue.fields("componentID").value & " name=promptDecs_C" & rstPromptedValue.fields("componentID").value & " value=""" & rstpromptedvalue.fields("promptDecimals").value & """>" & vbCrLf)

                            ' Logic Prompted Value
                        ElseIf rstPromptedValue.fields("ValueType").value = 3 Then
                            Response.Write("        <input type=""checkbox"" tabindex=""-1"" id=prompt_3_C" & rstPromptedValue.fields("componentID").value & " name=prompt_3_C" & rstPromptedValue.fields("componentID").value & vbCrLf)
                            Response.Write("            onclick=""checkboxClick('C" & rstPromptedValue.fields("componentID").value & "')""" & vbCrLf)
                            Response.Write("            onmouseover=""try{checkbox_onMouseOver(this);}catch(e){}""" & vbCrLf)
                            Response.Write("            onmouseout=""try{checkbox_onMouseOut(this);}catch(e){}""")
                            If rstPromptedvalue.fields("valuelogic").value Then
                                Response.Write(" CHECKED/>" & vbCrLf)
                            Else
                                Response.Write("/>" & vbCrLf)
                            End If
						  
                            Response.Write("        <input type=hidden id=promptChk_C" & rstPromptedValue.fields("componentID").value & " name=promptChk_C" & rstPromptedValue.fields("componentID").value & " value=" & rstPromptedvalue.fields("valuelogic").value & ">" & vbCrLf)
											 
                            ' Date Prompted Value
                        ElseIf rstPromptedValue.fields("ValueType") = 4 Then
                            Response.Write("        <input type=text class=""text"" id=prompt_4_C" & rstPromptedValue.fields("componentID").value & " name=prompt_4_C" & rstPromptedValue.fields("componentID").value & " value=""")
                            Select Case rstpromptedvalue.fields("promptDateType").value
                                Case 0
                                    ' Explicit value
                                    If Not IsDBNull(rstPromptedValue.fields("valuedate").value) Then
                                        If (CStr(rstPromptedValue.fields("valuedate").value) <> "00:00:00") And _
                                            (CStr(rstPromptedValue.fields("valuedate").value) <> "12:00:00 AM") Then
                                            Response.Write(convertSQLDateToLocale2(rstPromptedValue.fields("valuedate").value))
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
                            Response.Write(""" style=""WIDTH: 100%"">" & vbCrLf)

                            ' Lookup Prompted Value
                        ElseIf rstPromptedValue.fields("ValueType") = 5 Then
                            Response.Write("        		<SELECT id=promptLookup_C" & rstPromptedValue.fields("componentID").value & " name=promptLookup_C" & rstPromptedValue.fields("componentID").value & " class=""combo"" style=""WIDTH: 100%"" onchange=""comboChange('C" & rstPromptedValue.fields("componentID").value & "')"">" & vbCrLf)

                            fDefaultFound = False
                            fFirstValueDone = False
                            sFirstValue = ""

                            ' Get the lookup values.
                            cmdLookupValues = CreateObject("ADODB.Command")
                            cmdLookupValues.CommandText = "sp_ASRIntGetLookupValues"
                            cmdLookupValues.CommandType = 4 ' Stored Procedure
                            cmdLookupValues.ActiveConnection = Session("databaseConnection")

                            prmColumnID = cmdLookupValues.CreateParameter("columnID", 3, 1)
                            cmdLookupValues.Parameters.Append(prmColumnID)
                            prmColumnID.value = cleanNumeric(rstPromptedValue.fields("fieldColumnID").value)

                            Err.Clear()
                            rstLookupValues = cmdLookupValues.Execute

                            Do While Not rstLookupValues.EOF
                                Response.Write("        		  <OPTION")

                                If Not fFirstValueDone Then
                                    sFirstValue = rstLookupValues.Fields(0).Value
                                    fFirstValueDone = True
                                End If

                                If rstLookupValues.fields(0).type = 135 Then
                                    ' Field is a date so format as such.
                                    sOptionValue = convertSQLDateToLocale2(rstLookupValues.Fields(0).Value)
                                    If sOptionValue = convertSQLDateToLocale2(rstpromptedvalue.fields("valuecharacter")) Then
                                        Response.Write(" SELECTED")
                                        fDefaultFound = True
                                    End If
                                    Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
                                ElseIf rstLookupValues.fields(0).type = 131 Then
                                    ' Field is a numeric so format as such.
                                    sOptionValue = Replace(rstLookupValues.Fields(0).Value, ".", Session("LocaleDecimalSeparator"))
                                    If (Not IsDBNull(rstLookupValues.Fields(0).Value)) And (Not IsDBNull(rstPromptedValue.fields("valuecharacter"))) Then
                                        If FormatNumber(rstLookupValues.Fields(0).Value) = FormatNumber(rstPromptedValue.fields("valuecharacter")) Then
                                            Response.Write(" SELECTED")
                                            fDefaultFound = True
                                        End If
                                    End If
                                    Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
                                ElseIf rstLookupValues.fields(0).type = 11 Then
                                    ' Field is a logic so format as such.
                                    sOptionValue = rstLookupValues.Fields(0).Value
                                    If sOptionValue = rstpromptedvalue.fields("valuecharacter") Then
                                        Response.Write(" SELECTED")
                                        fDefaultFound = True
                                    End If
                                    Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
                                Else
                                    sOptionValue = RTrim(rstLookupValues.Fields(0).Value)
                                    If sOptionValue = rstpromptedvalue.fields("valuecharacter") Then
                                        Response.Write(" SELECTED")
                                        fDefaultFound = True
                                    End If
                                    Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
                                End If

                                rstLookupValues.MoveNext()
                            Loop

                            Response.Write("   		     </SELECT>" & vbCrLf)

                            If fDefaultFound Then
                                sDefaultValue = rstpromptedvalue.fields("valuecharacter").value
                            Else
                                sDefaultValue = sFirstValue
                            End If

                            If rstLookupValues.fields(0).type = 135 Then
                                ' Date.
                                Response.Write("        <input type=hidden id=prompt_4_C" & rstPromptedValue.fields("componentID").value & " name=prompt_4_C" & rstPromptedValue.fields("componentID").value & " value=" & convertSQLDateToLocale(sDefaultValue) & ">" & vbCrLf)
                            ElseIf rstLookupValues.fields(0).type = 131 Then
                                ' Numeric
                                Response.Write("        <input type=hidden id=prompt_2_C" & rstPromptedValue.fields("componentID").value & " name=prompt_2_C" & rstPromptedValue.fields("componentID").value & " value=" & Replace(sDefaultValue, ".", Session("LocaleDecimalSeparator")) & ">" & vbCrLf)
                            ElseIf rstLookupValues.fields(0).type = 11 Then
                                ' Logic
                                Response.Write("        <input type=hidden id=prompt_3_C" & rstPromptedValue.fields("componentID").value & " name=prompt_3_C" & rstPromptedValue.fields("componentID").value & " value=" & sDefaultValue & ">" & vbCrLf)
                            Else
                                Response.Write("        <input type=hidden id=prompt_1_C" & rstPromptedValue.fields("componentID").value & " name=prompt_1_C" & rstPromptedValue.fields("componentID").value & " value=""" & Replace(sDefaultValue, """", "&quot;") & """>" & vbCrLf)
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
                End If

                rstPromptedValue.close()
                rstPromptedValue = Nothing
                cmdDefn = Nothing
            End If
        Loop
    End If

    If iPromptCount > 0 Then
        Response.Write("				<tr>" & vbCrLf)
        Response.Write("					<td colspan=5 height=10>&nbsp;</td>" & vbCrLf)
        Response.Write("			  </tr>" & vbCrLf)
        Response.Write("				<tr height=20>" & vbCrLf)
        Response.Write("					<td width=20></td>" & vbCrLf)
        Response.Write("					<td colspan=3>" & vbCrLf)
        Response.Write("						<TABLE WIDTH=100% class=""invisible"" CELLSPACING=0 CELLPADDING=0>" & vbCrLf)
        Response.Write("							<TD>&nbsp;</TD>" & vbCrLf)
			
        Response.Write("							<td width=80>" & vbCrLf)
        Response.Write("							    <INPUT type=""button"" value=""OK"" name=""Submit"" class=""btn"" style=" & Chr(34) & "WIDTH: 80px" & Chr(34) & vbCrLf)
        Response.Write("									    onclick=""SubmitPrompts();""" & vbCrLf)
        Response.Write("                                        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
        Response.Write("                                        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
        Response.Write("                                        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
        Response.Write("                                        onblur=""try{button_onBlur(this);}catch(e){}"" />" & vbCrLf)
        Response.Write("							</td>")
        Response.Write("							<td width=20></td>" & vbCrLf)
        Response.Write("							<td width=80>" & vbCrLf)
        Response.Write("							    <INPUT type=""button"" value=""Cancel"" name=""Cancel"" class=""btn"" value=""Cancel"" style=" & Chr(34) & "WIDTH: 80px" & Chr(34) & vbCrLf)
        Response.Write("									    onclick=""CancelClick();""" & vbCrLf)
        Response.Write("                                        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
        Response.Write("                                        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
        Response.Write("                                        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
        Response.Write("                                        onblur=""try{button_onBlur(this);}catch(e){}"" />" & vbCrLf)
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
	<INPUT type="hidden" id=type name=type value=<%=Request.Form("type")%>>	
	<INPUT type="hidden" id=components1 name=components1 value="<% =Request.Form("components1")%>">
	<INPUT type="hidden" id=tableID name=tableID value=<%=Request.Form("tableID")%>>
</FORM>
    
    </div>
</body>
</html>


<script type="text/javascript">
    util_test_expression_pval_onload();
</script>



<script runat="server" language="vb">

function promptParameter(psDefnString, psParameter)
	dim iCharIndex
	dim sDefn
	
	sDefn = psDefnString

	iCharIndex = instr(sDefn, "	")
	if iCharIndex >= 0 then
		if psParameter = "NODEKEY" then
			promptParameter = left(sDefn, iCharIndex - 1)
			exit function
		end if
		
		sDefn = mid(sDefn, iCharIndex + 1)
		iCharIndex = instr(sDefn,  "	")
		if iCharIndex >= 0 then
			if psParameter = "PROMPTDESCRIPTION" then
				promptParameter = left(sDefn, iCharIndex - 1)
				exit function
			end if
			
			sDefn = mid(sDefn, iCharIndex + 1)
			iCharIndex = instr(sDefn, "	")
			if iCharIndex >= 0 then
				if psParameter = "VALUETYPE" then
					promptParameter = left(sDefn, iCharIndex - 1)
					exit function
				end if
				
				sDefn = mid(sDefn, iCharIndex + 1)
				iCharIndex = instr(sDefn, "	")
				if iCharIndex >= 0 then
					if psParameter = "PROMPTSIZE" then
						promptParameter = left(sDefn, iCharIndex - 1)
						exit function
					end if
					
					sDefn = mid(sDefn, iCharIndex + 1)
					iCharIndex = instr(sDefn, "	")
					if iCharIndex >= 0 then
						if psParameter = "PROMPTDECIMALS" then
							promptParameter = left(sDefn, iCharIndex - 1)
							exit function
						end if
						
						sDefn = mid(sDefn, iCharIndex + 1)
						iCharIndex = instr(sDefn, "	")
						if iCharIndex >= 0 then
							if psParameter = "PROMPTMASK" then
								promptParameter = left(sDefn, iCharIndex - 1)
								exit function
							end if
							
							sDefn = mid(sDefn, iCharIndex + 1)
							iCharIndex = instr(sDefn, "	")
							if iCharIndex >= 0 then
								if psParameter = "FIELDTABLEID" then
									promptParameter = left(sDefn, iCharIndex - 1)
									exit function
								end if
								
								sDefn = mid(sDefn, iCharIndex + 1)
								iCharIndex = instr(sDefn, "	")
								if iCharIndex >= 0 then
									if psParameter = "FIELDCOLUMNID" then
										promptParameter = left(sDefn, iCharIndex - 1)
										exit function
									end if
									
									sDefn = mid(sDefn, iCharIndex + 1)
									iCharIndex = instr(sDefn, "	")
									if iCharIndex >= 0 then
										if psParameter = "VALUECHARACTER" then
											promptParameter = left(sDefn, iCharIndex - 1)
											exit function
										end if
										
										sDefn = mid(sDefn, iCharIndex + 1)
										iCharIndex = instr(sDefn, "	")
										if iCharIndex >= 0 then
											if psParameter = "VALUENUMERIC" then
												promptParameter = left(sDefn, iCharIndex - 1)
												exit function
											end if
											
											sDefn = mid(sDefn, iCharIndex + 1)
											iCharIndex = instr(sDefn, "	")
											if iCharIndex >= 0 then
												if psParameter = "VALUELOGIC" then
													promptParameter = left(sDefn, iCharIndex - 1)
													exit function
												end if
												
												sDefn = mid(sDefn, iCharIndex + 1)
												iCharIndex = instr(sDefn, "	")
												if iCharIndex >= 0 then
													if psParameter = "VALUEDATE" then
														promptParameter = left(sDefn, iCharIndex - 1)
														exit function
													end if
													
													sDefn = mid(sDefn, iCharIndex + 1)
													if psParameter = "PROMPTDATETYPE" then
														promptParameter = left(sDefn, iCharIndex - 1)
														exit function
													end if
												end if	
											end if	
										end if	
									end if	
								end if	
							end if	
						end if	
					end if	
				end if	
			end if	
		end if	
	end if
	
	promptParameter = ""
end function

function convertDateToSQLDate(pdtDate)
	dim iDays
	dim iMonths
	dim iYears
	dim sResult
	
	sResult = ""
	iDays = day(pdtDate)
	iMonths = month(pdtDate)
	iYears = year(pdtDate)

	if iMonths < 10 then
		sResult = "0"
	end if
	sResult = sResult & iMonths & "/"
	
	if iDays < 10 then
		sResult = sResult & "0"
	end if
	sResult = sResult & iDays & "/" & iYears
	
	convertDateToSQLDate = sResult
end function

function convertSQLDateToLocale(psDate)
	dim sLocaleFormat
	dim iIndex
	
	if len(psDate) > 0 then	
		sLocaleFormat = session("LocaleDateFormat")
		
		iIndex = instr(sLocaleFormat,"dd")
		if iIndex > 0 then
			sLocaleFormat = left(sLocaleFormat, iIndex - 1) & _
				mid(psDate, 4, 2) & mid(sLocaleFormat, iIndex + 2)
		end if
		
		iIndex = instr(sLocaleFormat,"mm")
		if iIndex > 0 then
			sLocaleFormat = left(sLocaleFormat, iIndex - 1) & _
				left(psDate, 2) & mid(sLocaleFormat, iIndex + 2)
		end if
		
		iIndex = instr(sLocaleFormat,"yyyy")
		if iIndex > 0 then
			sLocaleFormat = left(sLocaleFormat, iIndex - 1) & _
				mid(psDate, 7, 4) & mid(sLocaleFormat, iIndex + 4)
		end if		

		convertSQLDateToLocale = sLocaleFormat
	else
		convertSQLDateToLocale = ""
	end if
end function

function convertSQLDateToLocale2(psDate)
	dim sLocaleFormat
	dim iIndex

	if len(psDate) > 0 then	
		sLocaleFormat = session("LocaleDateFormat")
		
		iIndex = instr(sLocaleFormat,"dd")
		if iIndex > 0 then
			if day(psDate) < 10 then
				sLocaleFormat = left(sLocaleFormat, iIndex - 1) & _
					"0" & day(psDate) & mid(sLocaleFormat, iIndex + 2)
			else
				sLocaleFormat = left(sLocaleFormat, iIndex - 1) & _
					day(psDate) & mid(sLocaleFormat, iIndex + 2)
			end if
		end if
		
		iIndex = instr(sLocaleFormat,"mm")
		if iIndex > 0 then
			if month(psDate) < 10 then
				sLocaleFormat = left(sLocaleFormat, iIndex - 1) & _
					"0" & month(psDate) & mid(sLocaleFormat, iIndex + 2)
			else
				sLocaleFormat = left(sLocaleFormat, iIndex - 1) & _
					month(psDate) & mid(sLocaleFormat, iIndex + 2)
			end if
		end if
		
		iIndex = instr(sLocaleFormat,"yyyy")
		if iIndex > 0 then
			sLocaleFormat = left(sLocaleFormat, iIndex - 1) & _
				year(psDate) & mid(sLocaleFormat, iIndex + 4)
		end if		

		convertSQLDateToLocale2 = sLocaleFormat
	else
		convertSQLDateToLocale2 = ""
	end if
end function
</script>
