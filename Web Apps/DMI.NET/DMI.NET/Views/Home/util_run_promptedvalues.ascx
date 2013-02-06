<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
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
	if session("action") = "STDREPORT_DATEPROMPT" or session("optionaction") = "STDREPORT_DATEPROMPT" then
		bStandardReportPrompt = true
	else
		bStandardReportPrompt = false
		session("utiltype") = Request.Form("utiltype")
		session("utilid") = Request.Form("utilid")
		session("utilname") = Request.Form("utilname")
		session("action") = Request.Form("action")
	end if
%>
        
<object
    classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
    id="Microsoft_Licensed_Class_Manager_1_0"
    viewastext>
    <param name="LPKPath" value="lpks/main.lpk">
</object>

<script type="text/javascript">    
 
    function promptedvalues_window_onload() {

           //remmed this - don't need to set current workframe source - leave as defsel.
        $("#workframe").attr("data-framesource", "UTIL_RUN_PROMPTEDVALUES");       

        var frmPromptedValues = OpenHR.getForm("workframe", "frmPromptedValues");

        //var frmPromptedValues = $("#frmPromptedValues");        // document.getElementById('frmPromptedValues');
        //var frmPromptedValues = document.getElementById('frmPromptedValues');



        frmPromptedValues.txtLocaleDateFormat.value = OpenHR.LocaleDateFormat;
        frmPromptedValues.txtLocaleDecimalSeparator.value = OpenHR.LocaleDecimalSeparator;
        frmPromptedValues.txtLocaleThousandSeparator.value = OpenHR.LocaleThousandSeparator;

        if (frmPromptedValues.RunInOptionFrame.value == "True") {
            $("#workframe").attr("data-framesource", "UTIL_RUN_PROMPTEDVALUES");

        } else {
            if (frmPromptedValues.StandardReportPrompt.value == "True") {
                $("#workframe").attr("data-framesource", "UTIL_RUN_PROMPTEDVALUES");                
            }
        }
        if (frmPromptedValues.txtPromptCount.value == 0) {
            OpenHR.submitForm(frmPromptedValues);

        } else {
            // Set focus on the first prompt control.
            var controlCollection = frmPromptedValues.elements;
            if (controlCollection != null) {
                for (i = 0; i < controlCollection.length; i++) {
                    sControlName = controlCollection.item(i).name;
                    sControlPrefix = sControlName.substr(0, 7);

                    if ((sControlPrefix == "prompt_") || (sControlName.substr(0, 13) == "promptLookup_")) {
                        controlCollection.item(i).focus();
                        break;
                    }
                }
            }

            // Resize the grid to show all propmted values.
            if (frmPromptedValues.StandardReportPrompt.value == "False") {
                iResizeBy = frmPromptedValues.offsetParent.scrollHeight - frmPromptedValues.offsetParent.clientHeight;
                if (frmPromptedValues.offsetParent.offsetHeight + iResizeBy > screen.height) {
                    try {
                        window.parent.moveTo((screen.width - frmPromptedValues.offsetParent.offsetWidth) / 2, 0);
                        window.parent.resizeTo(frmPromptedValues.offsetParent.offsetWidth, screen.height);
                    } catch (e) {
                    }
                } else {
                    try {
                        window.parent.moveTo((screen.width - frmPromptedValues.offsetParent.offsetWidth) / 2, (screen.height - (frmPromptedValues.offsetParent.offsetHeight + iResizeBy)) / 2);
                        window.parent.resizeBy(0, iResizeBy);
                    } catch (e) {
                    }
                }
            }
        }
    }
</script>

<div>
    
    
    <FORM name="frmPromptedValues" id="frmPromptedValues" method=POST action=
    <%
        If bStandardReportPrompt Then
            Response.Write("stdrpt_def_Absence")
        Else
            Response.Write("util_run")
        End If
    %>
    >

<%
	' Get variables for Absence Breakdown / Bradford Factor
	session("stdReport_StartDate") = Request.Form("txtFromDate")
	session("stdReport_EndDate") = Request.Form("txtToDate")
	session("stdReport_AbsenceTypes") = Request.Form("txtAbsenceTypes")
	session("stdReport_FilterID") = Request.Form("txtBaseFilterID")
	session("stdReport_PicklistID") = Request.Form("txtBasePicklistID")
	session("stdReport_Bradford_SRV") = Request.Form("txtSRV")
	session("stdReport_Bradford_ShowDurations") = Request.Form("txtShowDurations")
	session("stdReport_Bradford_ShowFormula") = Request.Form("txtShowFormula")	
	session("stdReport_Bradford_ShowInstances") = Request.Form("txtShowInstances")
	session("stdReport_Bradford_OmitBeforeStart") = Request.Form("txtOmitBeforeStart")
	session("stdReport_Bradford_OmitAfterEnd") = Request.Form("txtOmitAfterEnd")
	session("stdReport_Bradford_txtOrderBy1") = Request.Form("txtOrderBy1")		
	session("stdReport_Bradford_txtOrderBy1ID") = Request.Form("txtOrderBy1ID")
	session("stdReport_Bradford_txtOrderBy1Asc") = Request.Form("txtOrderBy1Asc")		
	session("stdReport_Bradford_txtOrderBy2") = Request.Form("txtOrderBy2")
	session("stdReport_Bradford_txtOrderBy2ID") = Request.Form("txtOrderBy2ID")		
	session("stdReport_Bradford_txtOrderBy2Asc") = Request.Form("txtOrderBy2Asc")		
	session("stdReport_PrintFilterPicklistHeader") = Request.Form("txtPrintFPinReportHeader")
	session("stdReport_MinimumBradfordFactor") = Request.Form("txtMinimumBradfordFactor")
	session("stdReport_MinimumBradfordFactorAmount") = Request.Form("txtMinimumBradfordFactorAmount")
	session("stdReport_DisplayBradfordDetail") = Request.Form("txtDisplayBradfordDetail")

	session("stdReport_OutputPreview") = Request.Form("txtSend_OutputPreview")
	session("stdReport_OutputFormat") = Request.Form("txtSend_OutputFormat")
	session("stdReport_OutputScreen") = Request.Form("txtSend_OutputScreen")
	session("stdReport_OutputPrinter") = Request.Form("txtSend_OutputPrinter")
	session("stdReport_OutputPrinterName") = Request.Form("txtSend_OutputPrinterName")
	session("stdReport_OutputSave") = Request.Form("txtSend_OutputSave")
	session("stdReport_OutputSaveExisting") = Request.Form("txtSend_OutputSaveExisting")
	session("stdReport_OutputEmail") = Request.Form("txtSend_OutputEmail")
	session("stdReport_OutputEmailAddr") = Request.Form("txtSend_OutputEmailAddr")
	session("stdReport_OutputEmailSubject") = Request.Form("txtSend_OutputEmailSubject")
	session("stdReport_OutputEmailAttachAs") = Request.Form("txtSend_OutputEmailAttachAs")
	session("stdReport_OutputFilename") = Request.Form("txtSend_OutputFilename")
	
	dim iPromptCount
	dim iPromptDateType
	dim fDefaultFound 
	dim fFirstValueDone 
    Dim sFirstValue
    Dim cmdDefn
    Dim prmUtilType
    Dim prmUtilID
    Dim prmRecordID
    ' Dim rstPromptedValue
    Dim iValueType
	
	iPromptCount = 0
	iPromptDateType = 0 
	
	'if Request.Form("utiltype") = 2 or Request.Form("utiltype") = 9 then
	if bStandardReportPrompt then
        cmdDefn = CreateObject("ADODB.Command")
		cmdDefn.CommandText = "spASRIntGetStandardReportDates"
		cmdDefn.CommandType = 4 ' Stored Procedure
        cmdDefn.ActiveConnection = Session("databaseConnection")

        prmUtilType = cmdDefn.CreateParameter("ReportType", 3, 1) ' 3=integer, 1=input
        cmdDefn.Parameters.Append(prmUtilType)
		prmUtilType.value = cleanNumeric(session("StandardReport_Type"))
	else
        cmdDefn = CreateObject("ADODB.Command")
		cmdDefn.CommandText = "sp_ASRIntGetUtilityPromptedValues"
		cmdDefn.CommandType = 4 ' Stored Procedure
        cmdDefn.ActiveConnection = Session("databaseConnection")

        prmUtilType = cmdDefn.CreateParameter("utilType", 3, 1) ' 3=integer, 1=input
        cmdDefn.Parameters.Append(prmUtilType)
		prmUtilType.value = cleanNumeric(session("utiltype"))

        prmUtilID = cmdDefn.CreateParameter("utilID", 3, 1) ' 3=integer, 1=input
        cmdDefn.Parameters.Append(prmUtilID)
		prmUtilID.value = cleanNumeric(session("utilid"))

        prmRecordID = cmdDefn.CreateParameter("recordID", 3, 1) ' 3=integer, 1=input
        cmdDefn.Parameters.Append(prmRecordID)
		prmRecordID.value = cleanNumeric(clng(session("singleRecordID")))
	end if

    Err.Clear()
    Dim rstPromptedValue = cmdDefn.Execute

	if not (rstPromptedValue.EOF and rstPromptedValue.BOF) then
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
                                Response.Write(ConvertSqlDateToLocale(rstPromptedValue.fields("valuedate").value))
                            End If
                        End If
										
                    Case 1
                        ' Current date
                        Response.Write(ConvertSqlDateToLocale(Date.Now))
						
                    Case 2
                        ' Start of current month
                        iDay = (Day(Date.Now) * -1) + 1
                        dtDate = DateAdd("d", iDay, Date.Now)
                        Response.Write(ConvertSqlDateToLocale(dtDate))
						
                    Case 3
                        ' End of current month
                        iDay = (Day(Date.Now) * -1) + 1
                        dtDate = DateAdd("d", iDay, Date.Now)
                        dtDate = DateAdd("m", 1, dtDate)
                        dtDate = DateAdd("d", -1, dtDate)
                        Response.Write(ConvertSqlDateToLocale(dtDate))
						
                    Case 4
                        ' Start of current year
                        iDay = (Day(Date.Now) * -1) + 1
                        iMonth = (Month(Date.Now) * -1) + 1
                        dtDate = DateAdd("d", iDay, Date.Now)
                        dtDate = DateAdd("m", iMonth, dtDate)
                        Response.Write(ConvertSqlDateToLocale(dtDate))
						
                    Case 5
                        ' End of current year
                        iDay = (Day(Date.Now) * -1) + 1
                        iMonth = (Month(Date.Now) * -1) + 1
                        dtDate = DateAdd("d", iDay, Date.Now)
                        dtDate = DateAdd("m", iMonth, dtDate)
                        dtDate = DateAdd("yyyy", 1, dtDate)
                        dtDate = DateAdd("d", -1, dtDate)
                        Response.Write(ConvertSqlDateToLocale(dtDate))
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
                cmdLookupValues.CommandType = 4 ' Stored Procedure
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
                        sOptionValue = ConvertSqlDateToLocale(rstLookupValues.Fields(0).Value)
                        If sOptionValue = ConvertSqlDateToLocale(rstPromptedValue.fields("valuecharacter")) Then
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
                        If sOptionValue = rstPromptedValue.fields("valuecharacter") Then
                            Response.Write(" SELECTED")
                            fDefaultFound = True
                        End If
                        Response.Write(">" & sOptionValue & "</OPTION>" & vbCrLf)
                    Else
                        sOptionValue = RTrim(rstLookupValues.Fields(0).Value)
                        If sOptionValue = rstPromptedValue.fields("valuecharacter") Then
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
            Response.Write("					<td width=20 height=10>&nbsp;</td>" & vbCrLf)
            Response.Write("				</tr>" & vbCrLf)

			rstPromptedValue.MoveNext
		loop
%>
				<tr>
					<td colspan=5 height=10>&nbsp;</td> 
			  </tr>
				<tr height=20>
					<td width=20>&nbsp;</td>
					<td colspan=3 align='center'>
						<TABLE class="invisible" CELLSPACING=0 CELLPADDING=0 ALIGN='center'>
							<TD width=20>&nbsp;</TD>
							<td width=80>

								<input type=button class="btn" name=Submit value=OK style="WIDTH: 80px" 
								    onclick="SubmitPrompts()"
	                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
	                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
	                                onfocus="try{button_onFocus(this);}catch(e){}"
	                                onblur="try{button_onBlur(this);}catch(e){}" />
							</td>
							<td width=20>&nbsp;</td>
							<td width=80>
								<input type=button class="btn" name=Cancel value=Cancel style="WIDTH: 80px"
								    onclick="CancelClick()" 
	                                onmouseover="try{button_onMouseOver(this);}catch(e){}" 
	                                onmouseout="try{button_onMouseOut(this);}catch(e){}"
	                                onfocus="try{button_onFocus(this);}catch(e){}"
	                                onblur="try{button_onBlur(this);}catch(e){}" />
							</td>
						</table>
					</td>
					<td width=20>&nbsp;</td>
				</tr>
				<tr>
					<td colspan=5 height=5>&nbsp;</td> 
				</tr>
			</table>
		</td>
	</tr>
</table>
<%
	end if
		
	rstPromptedValue.close
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
</FORM>

<!-- Form to return to record edit screen -->
<form action="emptyoption" method="post" id="frmRecordEdit" name="frmRecordEdit">
</form>

<FORM action="default_Submit" method=post id=frmGoto name=frmGoto style="visibility:hidden;display:none">
    <%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</FORM>

    </div>


<script type="text/javascript">
<!--
    function SubmitPrompts()
    {
        // Validate the prompt values before submitting the form.
        var controlCollection = frmPromptedValues.elements;
        if (controlCollection!=null) 
        {
            for (var i=0; i<controlCollection.length; i++)  
            {
                sControlName = controlCollection.item(i).name;
                sControlPrefix = sControlName.substr(0, 7);
	
                if (sControlPrefix=="prompt_") 
                {
                    // Get the control's data type.
                    iType = new Number(sControlName.substring(7,8));
                    if ((iType==1) || (iType==2) || (iType==4)) 
                    {
                        // Validate character, numeric and date prompts.
                        // Logic and lookup prompts do not need validation.
                        if (ValidatePrompt(controlCollection.item(i), iType) == false) 
                        {
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
        if (frmPromptedValues.StandardReportPrompt.value == "True")
        {
            if (frmPromptedValues.RunInOptionFrame.value == "True")
            {
                var frmParent =  window.dialogArguments.OpenHR.getForm("workframe","frmRecordEditForm");

                window.parent.document.all.item("workframeset").cols = "*, 0";		
                window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "visible";
                window.dialogArguments.OpenHR.submitForm(frmParent);
            }
            else
            {
                window.location.href="default";
            }
        }
        else
        {
            self.close();
        }	
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
        sDecimalSeparator = sDecimalSeparator.concat(frmPromptedValues.txtLocaleDecimalSeparator.value );
        var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");

        sThousandSeparator = "\\";
        sThousandSeparator = sThousandSeparator.concat(frmPromptedValues.txtLocaleThousandSeparator.value);
        var reThousandSeparator = new RegExp(sThousandSeparator, "gi");

        sPoint = "\\.";
        var rePoint = new RegExp(sPoint, "gi");

        fOK = true;

        if ((fOK == true) && (piDataType == 2)) 
        {
            // Numeric column.
            // Ensure that the value entered is numeric.
            sValue = pctlPrompt.value;

            if (sValue.length == 0) 
            {
                sValue = "0";
                pctlPrompt.value = 0;
            }

            // Convert the value from locale to UK settings for use with the isNaN funtion.
            sConvertedValue = new String(sValue);
            // Remove any thousand separators.
            sConvertedValue = sConvertedValue.replace(reThousandSeparator, "");
            pctlPrompt.value = sConvertedValue;

            // Convert any decimal separators to '.'.
            if (frmPromptedValues.txtLocaleDecimalSeparator.value  != ".") {
                // Remove decimal points.
                sConvertedValue = sConvertedValue.replace(rePoint, "A");
                // replace the locale decimal marker with the decimal point.
                sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
            }

            if (isNaN(sConvertedValue) == true) 
            {
                fOK = false;
                OpenHR.messageBox("Invalid numeric value entered.");
                window.focus();
                pctlPrompt.focus();
            }
        }
		
        if ((fOK == true) && (piDataType == 4)) 
        {
            // Date column.
            // Ensure that the value entered is a date.
            sValue = pctlPrompt.value;
			
            if (sValue.length == 0) 
            {
                fOK = false;
            }
            else 
            {
                // Convert the date to SQL format (use this as a validation check).
                // An empty string is returned if the date is invalid.
                sValue = convertLocaleDateToSQL(sValue);
                if (sValue.length == 0) 
                {
                    fOK = false;
                }
                else 
                {
                    pctlPrompt.value = OpenHR.ConvertSQLDateToLocale(sValue);
                }
            }
			
            if (fOK == false) 
            {
                OpenHR.messageBox("Invalid date value entered.");
                window.focus();
                pctlPrompt.focus();
            }	
        }
	
        if ((fOK == true) && (piDataType == 1)) 
        {
            // Character column.
            // Ensure that the value entered matches the required mask (if there is one).
            sMaskCtlName = "promptMask_" + pctlPrompt.name.substring(9, pctlPrompt.name.length);
            sMaskCtlName = sMaskCtlName.toUpperCase();

            fFound = false;		
            var controlCollection = frmPromptedValues.elements;
            if (controlCollection!=null) 
            {
                for (i=0; i<controlCollection.length; i++)  
                {
                    if (controlCollection.item(i).name.toUpperCase() == sMaskCtlName) 
                    {
                        fFound = true;
                        break;
                    }
                }
            }
		
            debugger;

            if (fFound == true) {
                sMask = frmPromptedValues.elements(sMaskCtlName).value;
                sValue = pctlPrompt.value;
                // Need to get rid of the backslash characters that precede literals.
                // But remember that two backslashes give a literal backslash that does not want
                // to be got rid of.
                sTemp = sMask.replace(reDoubleBackSlash, "a");
                sTemp = sTemp.replace(reBackSlash, "");
                if (sMask.length > 0) 
                {
                    if (sTemp.length != sValue.length) 
                    {
                        fOK = false;
                    }
                    else 
                    {
                        // Prompt values length matches mask length, so now check each character.
                        fFollowingBackslash = false;
                        iIndex = 0;
                        for (i=0; i<sMask.length; i++)  
                        {
                            sValueChar = sValue.substring(iIndex, iIndex+1);
						
                            if (fFollowingBackslash == false) 
                            {
                                switch (sMask.substring(i, i+1)) 
                                {
                                    case "A":
                                        // Character must be uppercase.
                                        if (sValueChar.toUpperCase() != sValueChar) 
                                        {
                                            fOK = false;
                                        }
                                        else 
                                        {
                                            iNumber = new Number(sValueChar);
                                            if (isNaN(iNumber) == false) 
                                            {
                                                fOK= true;
                                            }
                                        }
                                        iIndex = iIndex + 1;
                                        break;
									
                                    case "a":
                                        // Character must be lowercase.
                                        if (sValueChar.toLowerCase() != sValueChar) 
                                        {
                                            fOK = false;
                                        }
                                        else 
                                        {
                                            iNumber = new Number(sValueChar);
                                            if (isNaN(iNumber) == false) 
                                            {
                                                fOK= false;
                                            }
                                        }
                                        iIndex = iIndex + 1;
                                        break;
									
                                    case "9":
                                        // Character must be numeric (0-9).
                                        iNumber = new Number(sValueChar);
                                        if (isNaN(iNumber) == true) 
                                        {
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
                                            (sValueChar != "\\")) 
                                        {
                                            fOK= false;
                                        }
                                        iIndex = iIndex + 1;
                                        break;
									
                                    case "B":
                                        // Character must be logic (0 or 1).
                                        if ((sValueChar != "0") &&
                                            (sValueChar != "1"))
                                        {
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
                                        if (sMask.substring(i, i+1) != sValueChar) 
                                        {
                                            fOK = false;
                                        }
                                        iIndex = iIndex + 1;
                                }
                            }
                            else 
                            {
                                fFollowingBackslash = false;
                                if (sMask.substring(i, i+1) != sValueChar) 
                                {
                                    fOK = false;
                                }
                                iIndex = iIndex + 1;
                            }
						
                            if (fOK == false) 
                            {
                                break;
                            }
                        }
                    }
                }
		
                if (fOK == false) 
                {
                    OpenHR.messageBox("The entered value does not match the required format (" + sMask + ").");
                    window.focus();
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
		
        sDateFormat = frmPromptedValues.txtLocaleDateFormat.value;

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

        for (iLoop=0; iLoop<sDateFormat.length; iLoop++)  
        {
            if ((sDateFormat.substr(iLoop,1).toUpperCase() == 'D') && (sDays.length==0))
            {
                sDays = psDateString.substr(iValuePos,1);
                iValuePos = iValuePos + 1;
                sTempValue = psDateString.substr(iValuePos,1);

                if (isNaN(sTempValue) == false) 
                {
                    sDays = sDays.concat(sTempValue);			
                }
                iValuePos = iValuePos + 1;		
            }

            if ((sDateFormat.substr(iLoop,1).toUpperCase() == 'M') && (sMonths.length==0))
            {
                sMonths = psDateString.substr(iValuePos,1);
                iValuePos = iValuePos + 1;
                sTempValue = psDateString.substr(iValuePos,1);

                if (isNaN(sTempValue) == false) 
                {
                    sMonths = sMonths.concat(sTempValue);			
                }
                iValuePos = iValuePos + 1;
            }

            if ((sDateFormat.substr(iLoop,1).toUpperCase() == 'Y') && (sYears.length==0))
            {
                sYears = psDateString.substr(iValuePos,1);
                iValuePos = iValuePos + 1;
                sTempValue = psDateString.substr(iValuePos,1);

                if (isNaN(sTempValue) == false) 
                {
                    sYears = sYears.concat(sTempValue);			
                }
                iValuePos = iValuePos + 1;
                sTempValue = psDateString.substr(iValuePos,1);

                if (isNaN(sTempValue) == false) 
                {
                    sYears = sYears.concat(sTempValue);			
                }
                iValuePos = iValuePos + 1;
                sTempValue = psDateString.substr(iValuePos,1);

                if (isNaN(sTempValue) == false) 
                {
                    sYears = sYears.concat(sTempValue);			
                }
                iValuePos = iValuePos + 1;
            }

            // Skip non-numerics
            sTempValue = psDateString.substr(iValuePos,1);
            while (isNaN(sTempValue) == true) 
            {
                iValuePos = iValuePos + 1;		
                sTempValue = psDateString.substr(iValuePos,1);
            }
        }

        while (sDays.length < 2) 
        {
            sTempValue = "0";
            sDays = sTempValue.concat(sDays);
        }

        while (sMonths.length < 2) 
        {
            sTempValue = "0";
            sMonths = sTempValue.concat(sMonths);
        }

        while (sYears.length < 2) 
        {
            sTempValue = "0";
            sYears = sTempValue.concat(sYears);
        }

        if (sYears.length == 2) 
        {
            iValue = parseInt(sYears);
            if (iValue < 30) 
            {
                sTempValue = "20";
            }
            else 
            {
                sTempValue = "19";
            }
		
            sYears = sTempValue.concat(sYears);
        }

        while (sYears.length < 4) 
        {
            sTempValue = "0";
            sYears = sTempValue.concat(sYears);
        }

        sTempValue = sMonths.concat("/");
        sTempValue = sTempValue.concat(sDays);
        sTempValue = sTempValue.concat("/");
        sTempValue = sTempValue.concat(sYears);
	
        sValue = OpenHR.ConvertSQLDateToLocale(sTempValue);

        iYears = parseInt(sYears);
	
        while (sMonths.substr(0, 1) == "0") 
        {
            sMonths = sMonths.substr(1);
        }
        iMonths = parseInt(sMonths);
	
        while (sDays.substr(0, 1) == "0") 
        {
            sDays = sDays.substr(1);
        }
        iDays = parseInt(sDays);

        var newDateObj = new Date(iYears, iMonths - 1, iDays);

        if ((newDateObj.getDate() != iDays) || 
            (newDateObj.getMonth() + 1 != iMonths) || 
            (newDateObj.getFullYear() != iYears)) 
        {
            return "";
        }
        else 
        {
            return sTempValue;
        }
    }

    function checkboxClick(piPromptID) 
    {
        sSource = "prompt_3_" + piPromptID;
        sDest = "promptChk_" + piPromptID;
	
        frmPromptedValues.elements.item(sDest).value = frmPromptedValues.elements.item(sSource).checked;
    }

    function comboChange(piPromptID) 
    {
        sSource = "promptLookup_" + piPromptID;
        ctlSource = frmPromptedValues.elements.item(sSource);
	
        var controlCollection = frmPromptedValues.elements;
        if (controlCollection!=null) 
        {
            for (i=0; i<controlCollection.length; i++)  
            {
                sControlName = controlCollection.item(i).name;
                sControlPrefix = sControlName.substr(0, 7);
                sControlID = sControlName.substr(9, sControlName.length);
	 
                if ((sControlPrefix=="prompt_") && (sControlID == piPromptID)) 
                {
                    controlCollection.item(i).value = ctlSource.options[ctlSource.selectedIndex].text;
                }
            }
        }
    }
    -->
</script>
    
<script type="text/javascript">
    promptedvalues_window_onload();
</script>

