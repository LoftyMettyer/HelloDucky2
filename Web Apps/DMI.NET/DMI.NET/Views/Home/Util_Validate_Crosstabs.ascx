<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@Import namespace="DMI.NET" %>

<script type=”text/javascript">
	<%Html.RenderPartial("Util_Def_Crosstabs/dialog")%>
</script>

<div bgcolor='<%=session("ConvertedDesktopColour")%>' onload="return window_onload()" id=bdyMain leftmargin=20 topmargin=20 bottommargin=20 rightmargin=5>

<table align=center class="outline" cellPadding=5 cellSpacing=0>
	<TR>
		<TD>
			<table class="invisible" cellspacing="0" cellpadding="0">
				<tr> 
			    <td colspan=5 height=10></td>
			  </tr>

			  <tr id=trPleaseWait1> 
					<td width=20></td>
			    <td align=center colspan=3> 
						Validating Cross Tab
			    </td>
					<td width=20></td>
			  </tr>

			  <tr id=trPleaseWait4 height=10> 
					<td colspan=5></td>
			  </tr>

			  <tr id=trPleaseWait2> 
					<td width=20></td>
			    <td align=center colspan=3> 
						Please Wait...
			    </td>
					<td width=20></td>
			  </tr>

			  <tr id=trPleaseWait5 height=20> 
					<td colspan=5></td>
			  </tr>

			  <tr id=trPleaseWait3> 
					<td width=20></td>
			    <td align=center colspan=3> 
						<INPUT TYPE=button VALUE="Cancel" class="btn" NAME="Cancel" style="WIDTH: 80px" width=80 id=Cancel
						    OnClick="self.close()" 
                            onmouseover="try{button_onMouseOver(this);}catch(e){}" 
                            onmouseout="try{button_onMouseOut(this);}catch(e){}"
                            onfocus="try{button_onFocus(this);}catch(e){}"
                            onblur="try{button_onBlur(this);}catch(e){}" />
			    </td>
					<td width=20></td>
			  </tr>


<%
    Dim cmdValidate = CreateObject("ADODB.Command")
    cmdValidate.CommandText = "sp_ASRIntValidateCrossTab"
    cmdValidate.CommandType = 4 ' Stored Procedure
    cmdValidate.ActiveConnection = Session("databaseConnection")

    Dim prmUtilName = cmdValidate.CreateParameter("utilName", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdValidate.Parameters.Append(prmUtilName)
    prmUtilName.value = Request("validateName")

    Dim prmUtilID = cmdValidate.CreateParameter("utilID", 3, 1) '3=integer, 1=input
    cmdValidate.Parameters.Append(prmUtilID)
    prmUtilID.value = CleanNumeric(Request("validateUtilID"))

    Dim prmTimestamp = cmdValidate.CreateParameter("timestamp", 3, 1) '3=integer, 1=input
    cmdValidate.Parameters.Append(prmTimestamp)
    prmTimestamp.value = CleanNumeric(Request("validateTimestamp"))

    Dim prmBasePicklist = cmdValidate.CreateParameter("basePicklist", 3, 1) '3=integer, 1=input
    cmdValidate.Parameters.Append(prmBasePicklist)
    prmBasePicklist.value = CleanNumeric(Request("validateBasePicklist"))

    Dim prmBaseFilter = cmdValidate.CreateParameter("baseFilter", 3, 1) '3=integer, 1=input
    cmdValidate.Parameters.Append(prmBaseFilter)
    prmBaseFilter.value = CleanNumeric(Request("validateBaseFilter"))

    Dim prmEmailGroup = cmdValidate.CreateParameter("emailGroup", 3, 1) '3=integer, 1=input
    cmdValidate.Parameters.Append(prmEmailGroup)
    prmEmailGroup.value = CleanNumeric(Request("validateEmailGroup"))

    Dim prmHiddenGroups = cmdValidate.CreateParameter("hiddenGroups", 200, 1, 8000) '200=varchar, 1=input, 8000=size
    cmdValidate.Parameters.Append(prmHiddenGroups)
    prmHiddenGroups.value = Request("validateHiddenGroups")

    Dim prmErrorMsg = cmdValidate.CreateParameter("errorMsg", 200, 2, 8000) '200=varchar, 2=output, 8000=size
    cmdValidate.Parameters.Append(prmErrorMsg)

    Dim prmErrorCode = cmdValidate.CreateParameter("errorCode", 3, 2) '3=integer, 2=output
    cmdValidate.Parameters.Append(prmErrorCode)
	
    Dim prmDeletedFilters = cmdValidate.CreateParameter("deletedFilters", 200, 2, 8000) '200=varchar, 2=output, 8000=size
    cmdValidate.Parameters.Append(prmDeletedFilters)

    Dim prmHiddenFilters = cmdValidate.CreateParameter("hiddenFilters", 200, 2, 8000) '200=varchar, 2=output, 8000=size
    cmdValidate.Parameters.Append(prmHiddenFilters)

    Dim prmJobIDsToHide = cmdValidate.CreateParameter("jobsToHide", 200, 2, 8000) '200=varchar, 2=output, 8000=size
    cmdValidate.Parameters.Append(prmJobIDsToHide)

    Err.Clear()
    cmdValidate.Execute()

    Response.Write("<INPUT type=hidden id=txtErrorCode name=txtErrorCode value=" & cmdValidate.Parameters("errorCode").Value & ">" & vbCrLf)
    Response.Write("<INPUT type=hidden id=txtDeletedFilters name=txtDeletedFilters value=" & cmdValidate.Parameters("deletedFilters").Value & ">" & vbCrLf)
    Response.Write("<INPUT type=hidden id=txtHiddenFilters name=txtHiddenFilters value=" & cmdValidate.Parameters("hiddenFilters").Value & ">" & vbCrLf)
    Response.Write("<INPUT type=hidden id=txtJobIDsToHide name=txtJobIDsToHide value=""" & cmdValidate.Parameters("jobsToHide").Value & """>" & vbCrLf)

    If cmdValidate.Parameters("errorCode").Value = 1 Then
        Response.Write("			  <tr>" & vbCrLf)
        Response.Write("					<td width=20></td>" & vbCrLf)
        Response.Write("			    <td align=center colspan=3> " & vbCrLf)
        Response.Write("						<H3>Error Saving Cross Tab</H3>" & vbCrLf)
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
        Response.Write("    				    <INPUT TYPE=button VALUE=Close class=""btn"" NAME=Cancel style=""WIDTH: 80px"" width=80 id=Cancel" & vbCrLf)
        Response.Write("    				        OnClick=""self.close()""" & vbCrLf)
        Response.Write("    				        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
        Response.Write("    				        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
        Response.Write("    				        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
        Response.Write("    				        onblur=""try{button_onBlur(this);}catch(e){}""/>" & vbCrLf)
        Response.Write("			    </td>" & vbCrLf)
        Response.Write("					<td width=20></td>" & vbCrLf)
        Response.Write("			  </tr>" & vbCrLf)
    Else
        If cmdValidate.Parameters("errorCode").Value = 2 Then
            Response.Write("			  <tr>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			    <td align=center colspan=3> " & vbCrLf)
            Response.Write("						<H3>Error Saving Cross Tab</H3>" & vbCrLf)
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
            Response.Write("    				    <INPUT TYPE=button VALUE=Yes class=""btn"" NAME=btnYes style=""WIDTH: 80px"" width=80 id=btnYes" & vbCrLf)
            Response.Write("    				        OnClick=""createNew()""" & vbCrLf)
            Response.Write("    				        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
            Response.Write("    				        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
            Response.Write("    				        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
            Response.Write("    				        onblur=""try{button_onBlur(this);}catch(e){}""/>" & vbCrLf)
            Response.Write("			    </td>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			    <td align=left> " & vbCrLf)
            Response.Write("    				    <INPUT TYPE=button VALUE=No class=""btn"" NAME=btnNo style=""WIDTH: 80px"" width=80 id=btnNo" & vbCrLf)
            Response.Write("    				        OnClick=""self.close()""" & vbCrLf)
            Response.Write("    				        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
            Response.Write("    				        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
            Response.Write("    				        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
            Response.Write("    				        onblur=""try{button_onBlur(this);}catch(e){}""/>" & vbCrLf)
            Response.Write("			    </td>" & vbCrLf)
            Response.Write("					<td width=20></td>" & vbCrLf)
            Response.Write("			  </tr>" & vbCrLf)

        Else
            If cmdValidate.Parameters("errorCode").Value = 3 Then
                Response.Write("			  <tr>" & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                Response.Write("						<H3>Error Saving Cross Tab</H3>" & vbCrLf)
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
                Response.Write("    				    <INPUT TYPE=button VALUE=Yes class=""btn"" NAME=btnYes style=""WIDTH: 80px"" width=80 id=btnYes" & vbCrLf)
                Response.Write("    				        OnClick=""overwrite()""" & vbCrLf)
                Response.Write("    				        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
                Response.Write("    				        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
                Response.Write("    				        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
                Response.Write("    				        onblur=""try{button_onBlur(this);}catch(e){}""/>" & vbCrLf)
                Response.Write("			    </td>" & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("			    <td align=left> " & vbCrLf)
                Response.Write("    				    <INPUT TYPE=button VALUE=No class=""btn"" NAME=btnNo style=""WIDTH: 80px"" width=80 id=btnNo" & vbCrLf)
                Response.Write("    				        OnClick=""self.close()""" & vbCrLf)
                Response.Write("    				        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
                Response.Write("    				        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
                Response.Write("    				        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
                Response.Write("    				        onblur=""try{button_onBlur(this);}catch(e){}""/>" & vbCrLf)
                Response.Write("			    </td>" & vbCrLf)
                Response.Write("					<td width=20></td>" & vbCrLf)
                Response.Write("				</tr>" & vbCrLf)
            Else
                If cmdValidate.Parameters("errorCode").Value = 4 Then
                    Response.Write("			  <tr>" & vbCrLf)
                    Response.Write("					<td width=20></td>" & vbCrLf)
                    Response.Write("			    <td align=center colspan=3> " & vbCrLf)
                    Response.Write("						<H3>Error Saving Cross Tab</H3>" & vbCrLf)
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
                    Response.Write("    				    <INPUT TYPE=button VALUE=Yes class=""btn"" NAME=btnYes style=""WIDTH: 80px"" width=80 id=btnYes" & vbCrLf)
                    Response.Write("    				        OnClick=""continueSave()""" & vbCrLf)
                    Response.Write("    				        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
                    Response.Write("    				        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
                    Response.Write("    				        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
                    Response.Write("    				        onblur=""try{button_onBlur(this);}catch(e){}""/>" & vbCrLf)
                    Response.Write("			    </td>" & vbCrLf)
                    Response.Write("					<td width=20></td>" & vbCrLf)
                    Response.Write("			    <td align=left> " & vbCrLf)
                    Response.Write("    				    <INPUT TYPE=button VALUE=No class=""btn"" NAME=btnNo style=""WIDTH: 80px"" width=80 id=btnNo" & vbCrLf)
                    Response.Write("    				        OnClick=""self.close()""" & vbCrLf)
                    Response.Write("    				        onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
                    Response.Write("    				        onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
                    Response.Write("    				        onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
                    Response.Write("    				        onblur=""try{button_onBlur(this);}catch(e){}""/>" & vbCrLf)
                    Response.Write("			    </td>" & vbCrLf)
                    Response.Write("					<td width=20></td>" & vbCrLf)
                    Response.Write("				</tr>" & vbCrLf)
                End If
            End If
        End If
    End If
	
    cmdValidate = Nothing
%>
			  <tr height=10> 
					<td colspan=5></td>
			  </tr>
			</table>
		</TD>
	</TR>
</table>
</div>
