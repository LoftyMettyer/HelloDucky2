<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<object
    classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
    id="Microsoft_Licensed_Class_Manager_1_0">
    <param name="LPKPath" value="lpks/main.lpk">
</object>

<script type="text/javascript">

    function refreshInfo() {
        var frmUseful = window.parent.frames("workframe").document.forms("frmUseful");

        if (frmUseful.txtLoading.value == 0) {
            window.parent.frames("workframe").refreshDateSpecifics();
        }

        setOptions();
        return true;
    }

    function setOptions() {
        var frmNavFillerOptions = window.parent.frames("workframefiller").document.forms("frmOptions");

        with (frmNavFillerOptions) {
            txtIncludeBankHolidays.value = (frmOptions.chkIncludeBHols.checked);
            txtIncludeWorkingDaysOnly.value = (frmOptions.chkIncludeWorkingDaysOnly.checked);
            txtShowBankHolidays.value = (frmOptions.chkShadeBHols.checked);
            txtShowCaptions.value = (frmOptions.chkCaptions.checked);
            txtShowWeekends.value = (frmOptions.chkShadeWeekends.checked);

            submit();
            return;
        }
    }

</script>


<form name="frmOptions" id="frmOptions">
    <table align="center" class="outline" cellpadding="0" cellspacing="0" width="100%" height="100%">
        <tr>
            <td>
                <table class="invisible" cellspacing="0" cellpadding="2" width="100%" height="100%">
                    <tr height="5" valign="top">
                        <td width="5"></td>
                        <td height="100%" width="100%" align="left">Options :
                            <br>
                            <table width="100%" class="invisible" cellspacing="0" cellpadding="0">
                                <tr>
                                    <td width="5"></td>
                                    <td height="1">
                                        <%
                                            Dim objCalendar As Object
	
                                            objCalendar = Session("objCalendar" & Session("CalRepUtilID"))

                                            If objCalendar.IncludeBankHolidays_Enabled And objCalendar.IncludeBankHolidays Then
                                        %>
                                        <input name="chkIncludeBHols" id="chkIncludeBHols" type="checkbox" checked tabindex="-1"
                                            onclick="refreshInfo();"
                                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                        <label
                                            for="chkIncludeBHols"
                                            class="checkbox"
                                            tabindex="0"
                                            onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}" />
                                        <%
                                        ElseIf objCalendar.IncludeBankHolidays_Enabled And objCalendar.IncludeBankHolidays = False Then
                                        %>
                                        <input name="chkIncludeBHols" id="Checkbox1" type="checkbox" tabindex="-1"
                                            onclick="refreshInfo();"
                                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                        <label
                                            for="chkIncludeBHols"
                                            class="checkbox"
                                            tabindex="0"
                                            onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}" />
                                        <%
                                        Else
                                        %>
                                        <input name="chkIncludeBHols" id="Checkbox2" type="checkbox" disabled="disabled" tabindex="-1"
                                            onclick="refreshInfo();"
                                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                        <label
                                            for="chkIncludeBHols"
                                            class="checkbox checkboxdisabled"
                                            tabindex="0"
                                            onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}" />
                                        <%
                                        End If
                                        %>
											Include Bank Holidays 
                		    		        
                                    </td>
                                    <td width="5"></td>
                                </tr>
                                <tr>
                                    <td width="5"></td>
                                    <td height="1">
                                        <%
                                            If objCalendar.IncludeWorkingDaysOnly_Enabled And objCalendar.IncludeWorkingDaysOnly Then
                                        %>
                                        <input name="chkIncludeWorkingDaysOnly" id="chkIncludeWorkingDaysOnly" type="checkbox" checked tabindex="-1"
                                            onclick="refreshInfo();"
                                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                        <label
                                            for="chkIncludeWorkingDaysOnly"
                                            class="checkbox"
                                            tabindex="0"
                                            onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}" />
                                        <%
                                        ElseIf objCalendar.IncludeWorkingDaysOnly_Enabled And objCalendar.IncludeWorkingDaysOnly = False Then
                                        %>
                                        <input name="chkIncludeWorkingDaysOnly" id="Checkbox3" type="checkbox" tabindex="-1"
                                            onclick="refreshInfo();"
                                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                        <label
                                            for="chkIncludeWorkingDaysOnly"
                                            class="checkbox"
                                            tabindex="0"
                                            onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}" />
                                        <%
                                        Else
                                        %>
                                        <input name="chkIncludeWorkingDaysOnly" id="Checkbox4" type="checkbox" disabled="disabled" tabindex="-1"
                                            onclick="refreshInfo();"
                                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                        <label
                                            for="chkIncludeWorkingDaysOnly"
                                            class="checkbox checkboxdisabled"
                                            tabindex="0"
                                            onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}" />
                                        <%
                                        End If
                                        %>
											Working Days Only 
                		    		        
                                    </td>
                                    <td width="5"></td>
                                </tr>
                                <tr>
                                    <td width="5"></td>
                                    <td height="1">
                                        <%
                                            If objCalendar.ShowBankHolidays_Enabled And objCalendar.ShowBankHolidays Then
                                        %>

                                        <input name="chkShadeBHols" id="chkShadeBHols" type="checkbox" checked tabindex="-1"
                                            onclick="refreshInfo();"
                                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                        <label
                                            for="chkShadeBHols"
                                            class="checkbox"
                                            tabindex="0"
                                            onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}" />
                                        <%
                                        ElseIf objCalendar.ShowBankHolidays_Enabled And objCalendar.ShowBankHolidays = False Then
                                        %>
                                        <input name="chkShadeBHols" id="Checkbox5" type="checkbox" tabindex="-1"
                                            onclick="refreshInfo();"
                                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                        <label
                                            for="chkShadeBHols"
                                            class="checkbox"
                                            tabindex="0"
                                            onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}" />
                                        <%
                                        Else
                                        %>
                                        <input name="chkShadeBHols" id="Checkbox6" type="checkbox" disabled="disabled" tabindex="-1"
                                            onclick="refreshInfo();"
                                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                        <label
                                            for="chkShadeBHols"
                                            class="checkbox checkboxdisabled"
                                            tabindex="0"
                                            onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}" />
                                        <%
                                        End If
                                        %>
											Show Bank Holidays 
                		    		        
                                    </td>
                                    <td width="5"></td>
                                </tr>
                                <tr>
                                    <td width="5"></td>
                                    <td height="1" nowrap>
                                        <%
                                            If objCalendar.ShowCaptions Then
                                        %>
                                        <input name="chkCaptions" id="chkCaptions" type="checkbox" checked tabindex="-1"
                                            onclick="refreshInfo();"
                                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                        <%
                                        Else
                                        %>
                                        <input name="chkCaptions" id="Checkbox7" type="checkbox" tabindex="-1"
                                            onclick="refreshInfo();"
                                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />
                                        <%
                                        End If
                                        %>
                                        <label
                                            for="chkCaptions"
                                            class="checkbox"
                                            tabindex="0"
                                            onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
                                            Show Calendar Captions 
                                        </label>
                                    </td>
                                    <td width="5"></td>
                                </tr>
                                <tr>
                                    <td width="5"></td>
                                    <td height="1">
                                        <%
                                            If objCalendar.ShowWeekends Then
                                        %>
                                        <input name="chkShadeWeekends" id="chkShadeWeekends" type="checkbox" checked tabindex="-1"
                                            onclick="refreshInfo();" onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />

                                        <%
                                        Else
                                        %>
                                        <input name="chkShadeWeekends" id="Checkbox8" type="checkbox" tabindex="-1"
                                            onclick="refreshInfo();"
                                            onmouseover="try{checkbox_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkbox_onMouseOut(this);}catch(e){}" />

                                        <%
                                        End If
	
                                        objCalendar = Nothing
                                        %>
                                        <label
                                            for="chkShadeWeekends"
                                            class="checkbox"
                                            tabindex="0"
                                            onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}"
                                            onmouseover="try{checkboxLabel_onMouseOver(this);}catch(e){}"
                                            onmouseout="try{checkboxLabel_onMouseOut(this);}catch(e){}"
                                            onfocus="try{checkboxLabel_onFocus(this);}catch(e){}"
                                            onblur="try{checkboxLabel_onBlur(this);}catch(e){}">
                                            Show Weekends 
                                        </label>
                                    </td>
                                    <td width="5"></td>
                                </tr>
                            </table>
                        </td>
                        <td width="5"></td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>

    <input type="hidden" id="txtCalRep_UtilID" name="txtCalRep_UtilID" value='<%Session("CalRepUtilID").ToString()%>'>
</form>
