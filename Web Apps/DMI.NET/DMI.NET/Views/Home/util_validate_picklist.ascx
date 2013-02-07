<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>


<script type="text/javascript">

    function util_validate_picklist_window_onload() {

        debugger;

        $("#reportframe").attr("data-framesource", "UTIL_VALIDATE_PICKLIST");

        if (txtDisplay.value != "False") {
            // Hide the 'please wait' message.
            trPleaseWait1.style.visibility = 'hidden';
            trPleaseWait1.style.display = 'none';
            trPleaseWait2.style.visibility = 'hidden';
            trPleaseWait2.style.display = 'none';
            trPleaseWait3.style.visibility = 'hidden';
            trPleaseWait3.style.display = 'none';
            trPleaseWait4.style.visibility = 'hidden';
            trPleaseWait4.style.display = 'none';
            trPleaseWait5.style.visibility = 'hidden';
            trPleaseWait5.style.display = 'none';

        }
        else {
            nextPass();
        }
    }

    function nextPass() {
        var sURL;

        var frmValidate = OpenHR.getForm("reportframe", "frmValidatePicklist");

        iNextPass = new Number(frmValidate.validatePass.value);
        iNextPass = iNextPass + 1;

        if (iNextPass == 2) {
            frmValidate.validatePass.value = iNextPass;

            sURL = "util_validate_picklist" +
                "?validatePass=" + frmValidate.validatePass.value +
                "&validateName=" + escape(frmValidate.validateName.value) +
                "&validateTimestamp=" + frmValidate.validateTimestamp.value +
                "&validateUtilID=" + frmValidate.validateUtilID.value +
                "&validateAccess=" + frmValidate.validateAccess.value +
                "&validateBaseTableID=" + frmValidate.validateBaseTableID.value;

            //window.location.replace(sURL);                
            OpenHR.submitForm(frmValidate);            
        }
        else {
            var frmSend = OpenHR.getForm("workframe","frmSend");
            OpenHR.submitForm(frmSend);
        }
    }

    function overwrite() {
        nextPass();
    }

    function createNew() {
        window.dialogArguments.OpenHR.createNew(self);
    }

    function makeHidden() {
        nextPass();
    }

</script>


<table align="center" class="outline" cellpadding="5" cellspacing="0">
    <tr>
        <td>
            <table class="invisible" cellspacing="0" cellpadding="0">
                <tr>
                    <td colspan="5" height="10"></td>
                </tr>

                <tr id="trPleaseWait1">
                    <td width="20"></td>
                    <td align="center" colspan="3">Validating Picklist
                    </td>
                    <td width="20"></td>
                </tr>

                <tr id="trPleaseWait4" height="10">
                    <td colspan="5"></td>
                </tr>

                <tr id="trPleaseWait2">
                    <td width="20"></td>
                    <td align="center" colspan="3">Please Wait...
                    </td>
                    <td width="20"></td>
                </tr>

                <tr id="trPleaseWait5" height="20">
                    <td colspan="5"></td>
                </tr>

                <tr id="trPleaseWait3">
                    <td width="20"></td>
                    <td align="center" colspan="3">
                        <input type="button" value="Cancel" class="btn" name="Cancel" style="WIDTH: 80px" width="80" id="Cancel"
                            onclick="self.close()"
                            onmouseover="try{button_onMouseOver(this);}catch(e){}"
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
                    </td>
                    <td width="20"></td>
                </tr>


                <%
                    Dim fDisplay
                    fDisplay = False
	
                    Dim cmdValidate
                    Dim prmUtilName
                    Dim prmUtilID
                    Dim prmTimestamp
                    Dim prmAccess
                    Dim prmErrorMsg
                    Dim prmErrorCode
                    Dim prmBaseTableID
    
                    If Request("validatePass") = 1 Then
                        cmdValidate = Server.CreateObject("ADODB.Command")
                        cmdValidate.CommandText = "sp_ASRIntValidatePicklist"
                        cmdValidate.CommandType = 4 ' Stored Procedure
                        cmdValidate.ActiveConnection = Session("databaseConnection")

                        prmUtilName = cmdValidate.CreateParameter("utilName", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
                        cmdValidate.Parameters.Append(prmUtilName)
                        prmUtilName.value = Request("validateName")

                        prmUtilID = cmdValidate.CreateParameter("utilID", 3, 1) '3=integer, 1=input
                        cmdValidate.Parameters.Append(prmUtilID)
                        prmUtilID.value = CleanNumeric(Request("validateUtilID"))

                        prmTimestamp = cmdValidate.CreateParameter("timestamp", 3, 1) '3=integer, 1=input
                        cmdValidate.Parameters.Append(prmTimestamp)
                        prmTimestamp.value = CleanNumeric(Request("validateTimestamp"))

                        prmAccess = cmdValidate.CreateParameter("access", 200, 1, 8000) '200=varchar, 1=input, 8000=size
                        cmdValidate.Parameters.Append(prmAccess)
                        prmAccess.value = Request("validateAccess")

                        prmErrorMsg = cmdValidate.CreateParameter("errorMsg", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                        cmdValidate.Parameters.Append(prmErrorMsg)

                        prmErrorCode = cmdValidate.CreateParameter("errorCode", 3, 2) '3=integer, 2=output
                        cmdValidate.Parameters.Append(prmErrorCode)

                        Err.Clear()
                        cmdValidate.Execute()

                        If cmdValidate.Parameters("errorCode").Value = 1 Then
                            fDisplay = True
                            Response.Write("			  <tr>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                            Response.Write("						<H3>Error Saving Picklist</H3>" & vbCrLf)
                            Response.Write("			    </td>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			  </tr>" & vbCrLf)
                            Response.Write("			  <tr>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                            Response.Write("						" & cmdValidate.Parameters("errorMsg").Value & vbCrLf)
                            Response.Write("			    </td>" & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			  </tr>" & vbCrLf)
                            Response.Write("			  <tr>" & vbCrLf)
                            Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
                            Response.Write("			  </tr>" & vbCrLf)
                            Response.Write("			  <tr> " & vbCrLf)
                            Response.Write("					<td width=20></td>" & vbCrLf)
                            Response.Write("			    <td align=right> " & vbCrLf)
                %>
                <input type="button" value="Yes" class="btn" name="btnYes" style="WIDTH: 80px" width="80" id="btnYes"
                    onclick="createNew()"
                    onmouseover="try{button_onMouseOver(this);}catch(e){}"
                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                    onfocus="try{button_onFocus(this);}catch(e){}"
                    onblur="try{button_onBlur(this);}catch(e){}" />
                <%
                    Response.Write("			    </td>" & vbCrLf)
                    Response.Write("					<td width=20></td>" & vbCrLf)
                    Response.Write("			    <td align=left> " & vbCrLf)
                %>
                <input type="button" value="No" class="btn" name="btnNo" style="WIDTH: 80px" width="80" id="btnNo"
                    onclick="self.close()"
                    onmouseover="try{button_onMouseOver(this);}catch(e){}"
                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                    onfocus="try{button_onFocus(this);}catch(e){}"
                    onblur="try{button_onBlur(this);}catch(e){}" />
                <%
                    Response.Write("			    </td>" & vbCrLf)
                    Response.Write("					<td width=20></td>" & vbCrLf)
                    Response.Write("			  </tr>" & vbCrLf)
                Else
                    If cmdValidate.Parameters("errorCode").Value = 2 Then
                        fDisplay = True
                        Response.Write("			  <tr>" & vbCrLf)
                        Response.Write("					<td width=20></td>" & vbCrLf)
                        Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                        Response.Write("						<H3>Error Saving Picklist</H3>" & vbCrLf)
                        Response.Write("			    </td>" & vbCrLf)
                        Response.Write("					<td width=20></td>" & vbCrLf)
                        Response.Write("			  </tr>" & vbCrLf)
                        Response.Write("			  <tr>" & vbCrLf)
                        Response.Write("					<td width=20></td>" & vbCrLf)
                        Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                        Response.Write("						" & cmdValidate.Parameters("errorMsg").Value & vbCrLf)
                        Response.Write("			    </td>" & vbCrLf)
                        Response.Write("					<td width=20></td>" & vbCrLf)
                        Response.Write("			  </tr>" & vbCrLf)
                        Response.Write("			  <tr>" & vbCrLf)
                        Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
                        Response.Write("			  </tr>" & vbCrLf)
                        Response.Write("			  <tr> " & vbCrLf)
                        Response.Write("					<td width=20></td>" & vbCrLf)
                        Response.Write("			    <td align=right> " & vbCrLf)
                %>
                <input type="button" value="Yes" class="btn" name="btnYes" style="WIDTH: 80px" width="80" id="Button1"
                    onclick="overwrite()"
                    onmouseover="try{button_onMouseOver(this);}catch(e){}"
                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                    onfocus="try{button_onFocus(this);}catch(e){}"
                    onblur="try{button_onBlur(this);}catch(e){}" />
                <%
                    Response.Write("			    </td>" & vbCrLf)
                    Response.Write("					<td width=20></td>" & vbCrLf)
                    Response.Write("			    <td align=left> " & vbCrLf)
                %>
                <input type="button" value="No" class="btn" name="btnNo" style="WIDTH: 80px" width="80" id="Button2"
                    onclick="self.close()"
                    onmouseover="try{button_onMouseOver(this);}catch(e){}"
                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                    onfocus="try{button_onFocus(this);}catch(e){}"
                    onblur="try{button_onBlur(this);}catch(e){}" />
                <%
                    Response.Write("			    </td>" & vbCrLf)
                    Response.Write("					<td width=20></td>" & vbCrLf)
                    Response.Write("				</tr>" & vbCrLf)
                End If
            End If
	
            cmdValidate = Nothing
        Else
            If Request("validatePass") = 2 Then
                cmdValidate = Server.CreateObject("ADODB.Command")
                cmdValidate.CommandText = "sp_ASRIntValidatePicklist2"
                cmdValidate.CommandType = 4 ' Stored Procedure
                cmdValidate.ActiveConnection = Session("databaseConnection")

                prmUtilName = cmdValidate.CreateParameter("utilName", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
                cmdValidate.Parameters.Append(prmUtilName)
                prmUtilName.value = Request("validateName")

                prmUtilID = cmdValidate.CreateParameter("utilID", 3, 1) '3=integer, 1=input
                cmdValidate.Parameters.Append(prmUtilID)
                prmUtilID.value = CleanNumeric(Request("validateUtilID"))

                prmAccess = cmdValidate.CreateParameter("access", 200, 1, 8000) '200=varchar, 1=input, 8000=size
                cmdValidate.Parameters.Append(prmAccess)
                prmAccess.value = Request("validateAccess")

                prmBaseTableID = cmdValidate.CreateParameter("baseTableID", 3, 1) '3=integer, 1=input
                cmdValidate.Parameters.Append(prmBaseTableID)
                prmBaseTableID.value = CleanNumeric(Request("validateBaseTableID"))

                prmErrorMsg = cmdValidate.CreateParameter("errorMsg", 200, 2, 8000) '200=varchar, 2=output, 8000=size
                cmdValidate.Parameters.Append(prmErrorMsg)

                prmErrorCode = cmdValidate.CreateParameter("errorCode", 3, 2) '3=integer, 2=output
                cmdValidate.Parameters.Append(prmErrorCode)

                Err.Clear()
                cmdValidate.Execute()

                If cmdValidate.Parameters("errorCode").Value = 1 Then
                    fDisplay = True
                    Response.Write("			  <tr>" & vbCrLf)
                    Response.Write("					<td width=20></td>" & vbCrLf)
                    Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                    Response.Write("						<H3>Error Saving Picklist</H3>" & vbCrLf)
                    Response.Write("			    </td>" & vbCrLf)
                    Response.Write("					<td width=20></td>" & vbCrLf)
                    Response.Write("			  </tr>" & vbCrLf)
                    Response.Write("			  <tr>" & vbCrLf)
                    Response.Write("					<td width=20></td>" & vbCrLf)
                    Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                    Response.Write("						" & cmdValidate.Parameters("errorMsg").Value & vbCrLf)
                    Response.Write("			    </td>" & vbCrLf)
                    Response.Write("					<td width=20></td>" & vbCrLf)
                    Response.Write("			  </tr>" & vbCrLf)
                    Response.Write("			  <tr>" & vbCrLf)
                    Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
                    Response.Write("			  </tr>" & vbCrLf)
                    Response.Write("			  <tr> " & vbCrLf)
                    Response.Write("					<td width=20></td>" & vbCrLf)
                    Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                %>
                <input type="button" value="Close" class="btn" name="Cancel" style="WIDTH: 80px" width="80" id="Button3"
                    onclick="self.close()"
                    onmouseover="try{button_onMouseOver(this);}catch(e){}"
                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                    onfocus="try{button_onFocus(this);}catch(e){}"
                    onblur="try{button_onBlur(this);}catch(e){}" />
                <%
                    Response.Write("			    </td>" & vbCrLf)
                    Response.Write("					<td width=20></td>" & vbCrLf)
                    Response.Write("			  </tr>" & vbCrLf)
                Else
                    If cmdValidate.Parameters("errorCode").Value = 2 Then
                        fDisplay = True
                        Response.Write("			  <tr>" & vbCrLf)
                        Response.Write("					<td width=20></td>" & vbCrLf)
                        Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                        Response.Write("						<H3>Error Saving Picklist</H3>" & vbCrLf)
                        Response.Write("			    </td>" & vbCrLf)
                        Response.Write("					<td width=20></td>" & vbCrLf)
                        Response.Write("			  </tr>" & vbCrLf)
                        Response.Write("			  <tr>" & vbCrLf)
                        Response.Write("					<td width=20></td>" & vbCrLf)
                        Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                        Response.Write("						" & cmdValidate.Parameters("errorMsg").Value & vbCrLf)
                        Response.Write("			    </td>" & vbCrLf)
                        Response.Write("					<td width=20></td>" & vbCrLf)
                        Response.Write("			  </tr>" & vbCrLf)
                        Response.Write("			  <tr>" & vbCrLf)
                        Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
                        Response.Write("			  </tr>" & vbCrLf)
                        Response.Write("			  <tr> " & vbCrLf)
                        Response.Write("					<td width=20></td>" & vbCrLf)
                        Response.Write("			    <td align=right> " & vbCrLf)
                %>
                <input type="button" value="Yes" class="btn" name="btnYes" style="WIDTH: 80px" width="80" id="Button4"
                    onclick="makeHidden()"
                    onmouseover="try{button_onMouseOver(this);}catch(e){}"
                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                    onfocus="try{button_onFocus(this);}catch(e){}"
                    onblur="try{button_onBlur(this);}catch(e){}" />
                <%
                    Response.Write("			    </td>" & vbCrLf)
                    Response.Write("					<td width=20></td>" & vbCrLf)
                    Response.Write("			    <td align=left> " & vbCrLf)
                %>
                <input type="button" value="No" class="btn" name="btnNo" style="WIDTH: 80px" width="80" id="Button5"
                    onclick="self.close()"
                    onmouseover="try{button_onMouseOver(this);}catch(e){}"
                    onmouseout="try{button_onMouseOut(this);}catch(e){}"
                    onfocus="try{button_onFocus(this);}catch(e){}"
                    onblur="try{button_onBlur(this);}catch(e){}" />
                <%                            
                    Response.Write("			    </td>" & vbCrLf)
                    Response.Write("					<td width=20></td>" & vbCrLf)
                    Response.Write("				</tr>" & vbCrLf)
                End If
            End If
	
            cmdValidate = Nothing

        End If
    End If
	
    Response.Write("<INPUT type=hidden id=txtDisplay name=txtDisplay value=" & fDisplay & ">" & vbCrLf)
                %>
                <tr height="10">
                    <td colspan="5"></td>
                </tr>
            </table>
        </td>
    </tr>
</table>


<form id="frmValidatePicklist" name="frmValidatePicklist" method="post" action="util_validate_picklist" style="visibility: hidden; display: none">
    <input type="hidden" id="validatePass" name="validatePass" value='<%=Request("validatePass")%>'>
    <input type="hidden" id="validateName" name="validateName" value="<%=replace(Request("validateName"), """", "&quot;")%>">
    <input type="hidden" id="validateTimestamp" name="validateTimestamp" value='<%=Request("validateTimestamp")%>'>
    <input type="hidden" id="validateUtilID" name="validateUtilID" value='<%=Request("validateUtilID")%>'>
    <input type="hidden" id="validateAccess" name="validateAccess" value='<%=Request("validateAccess")%>'>
    <input type="hidden" id="validateBaseTableID" name="validateBaseTableID" value='<%=Request("validateBaseTableID")%>'>
    <input type="hidden" id="test" name="test" value="<%=Request.QueryString%>">
</form>

<script type="text/javascript">
    util_validate_picklist_window_onload();
</script>
