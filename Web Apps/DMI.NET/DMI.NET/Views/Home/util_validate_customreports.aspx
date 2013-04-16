<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>

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
</head>
<body id=bdyMain>
    
    
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
						Validating Report
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
	cmdValidate.CommandText = "sp_ASRIntValidateReport"
	cmdValidate.CommandType = 4 ' Stored Procedure
    cmdValidate.ActiveConnection = Session("databaseConnection")

    Dim prmUtilName = cmdValidate.CreateParameter("utilName", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
    cmdValidate.Parameters.Append(prmUtilName)
	prmUtilName.value = Request("validateName")

    Dim prmUtilID = cmdValidate.CreateParameter("utilID", 3, 1) '3=integer, 1=input
    cmdValidate.Parameters.Append(prmUtilID)
	prmUtilID.value = cleanNumeric(Request("validateUtilID"))

    Dim prmTimestamp = cmdValidate.CreateParameter("timestamp", 3, 1) '3=integer, 1=input
    cmdValidate.Parameters.Append(prmTimestamp)
	prmTimestamp.value = cleanNumeric(Request("validateTimestamp"))

    Dim prmBasePicklist = cmdValidate.CreateParameter("basePicklist", 3, 1) '3=integer, 1=input
    cmdValidate.Parameters.Append(prmBasePicklist)
	prmBasePicklist.value = cleanNumeric(Request("validateBasePicklist"))

    Dim prmBaseFilter = cmdValidate.CreateParameter("baseFilter", 3, 1) '3=integer, 1=input
    cmdValidate.Parameters.Append(prmBaseFilter)
	prmBaseFilter.value = cleanNumeric(Request("validateBaseFilter"))

    Dim prmEmailGroup = cmdValidate.CreateParameter("emailGroup", 3, 1) '3=integer, 1=input
    cmdValidate.Parameters.Append(prmEmailGroup)
	prmEmailGroup.value = cleanNumeric(Request("validateEmailGroup"))

    Dim prmParent1Picklist = cmdValidate.CreateParameter("parent1Picklist", 3, 1) '3=integer, 1=input
    cmdValidate.Parameters.Append(prmParent1Picklist)
	prmParent1Picklist.value = cleanNumeric(Request("validateP1Picklist"))

    Dim prmParent1Filter = cmdValidate.CreateParameter("parent1Filter", 3, 1) '3=integer, 1=input
    cmdValidate.Parameters.Append(prmParent1Filter)
	prmParent1Filter.value = cleanNumeric(Request("validateP1Filter"))

    Dim prmParent2Picklist = cmdValidate.CreateParameter("parent2Picklist", 3, 1) '3=integer, 1=input
    cmdValidate.Parameters.Append(prmParent2Picklist)
	prmParent2Picklist.value = cleanNumeric(Request("validateP2Picklist"))

    Dim prmParent2Filter = cmdValidate.CreateParameter("parent2Filter", 3, 1) '3=integer, 1=input
    cmdValidate.Parameters.Append(prmParent2Filter)
	prmParent2Filter.value = cleanNumeric(Request("validateP2Filter"))

    Dim prmChildFilter = cmdValidate.CreateParameter("childFilter", 200, 1, 8000) '200=varchar, 1=input, 8000=size
    cmdValidate.Parameters.Append(prmChildFilter)
	prmChildFilter.value = Request("validateChildFilter")

    Dim prmCalcs = cmdValidate.CreateParameter("calcs", 200, 1, 8000) '200=varchar, 1=input, 8000=size
    cmdValidate.Parameters.Append(prmCalcs)
	prmCalcs.value = Request("validateCalcs")

    Dim prmHiddenGroups = cmdValidate.CreateParameter("hiddenGroups", 200, 1, 8000) '200=varchar, 1=input, 8000=size
    cmdValidate.Parameters.Append(prmHiddenGroups)
	prmHiddenGroups.value = Request("validateHiddenGroups")

    Dim prmErrorMsg = cmdValidate.CreateParameter("errorMsg", 200, 2, 8000) '200=varchar, 2=output, 8000=size
    cmdValidate.Parameters.Append(prmErrorMsg)

    Dim prmErrorCode = cmdValidate.CreateParameter("errorCode", 3, 2) '3=integer, 2=output
    cmdValidate.Parameters.Append(prmErrorCode)
	
    Dim prmDeletedCalcs = cmdValidate.CreateParameter("deletedCalcs", 200, 2, 8000) '200=varchar, 2=output, 8000=size
    cmdValidate.Parameters.Append(prmDeletedCalcs)

    Dim prmHiddenCalcs = cmdValidate.CreateParameter("hiddenCalcs", 200, 2, 8000) '200=varchar, 2=output, 8000=size
    cmdValidate.Parameters.Append(prmHiddenCalcs)
	
    Dim prmDeletedFilters = cmdValidate.CreateParameter("deletedFilters", 200, 2, 8000) '200=varchar, 2=output, 8000=size
    cmdValidate.Parameters.Append(prmDeletedFilters)

    Dim prmHiddenFilters = cmdValidate.CreateParameter("hiddenFilters", 200, 2, 8000) '200=varchar, 2=output, 8000=size
    cmdValidate.Parameters.Append(prmHiddenFilters)

    Dim prmDeletedOrders = cmdValidate.CreateParameter("deletedOrders", 200, 2, 8000) '200=varchar, 2=output, 8000=size
    cmdValidate.Parameters.Append(prmDeletedOrders)
	
    Dim prmJobIDsToHide = cmdValidate.CreateParameter("jobsToHide", 200, 2, 8000) '200=varchar, 2=output, 8000=size
    cmdValidate.Parameters.Append(prmJobIDsToHide)

    Dim prmDeletedPicklists = cmdValidate.CreateParameter("deletedPicklists", 200, 2, 8000) '200=varchar, 2=output, 8000=size
    cmdValidate.Parameters.Append(prmDeletedPicklists)

    Dim prmHiddenPicklists = cmdValidate.CreateParameter("hiddenPicklists", 200, 2, 8000) '200=varchar, 2=output, 8000=size
    cmdValidate.Parameters.Append(prmHiddenPicklists)

    Err.Clear()
	cmdValidate.Execute

    Response.Write("<INPUT type=hidden id=txtErrorCode name=txtErrorCode value=" & cmdValidate.Parameters("errorCode").Value & ">" & vbCrLf)
    Response.Write("<INPUT type=hidden id=txtDeletedCalcs name=txtDeletedCalcs value=" & cmdValidate.Parameters("deletedCalcs").Value & ">" & vbCrLf)
    Response.Write("<INPUT type=hidden id=txtHiddenCalcs name=txtHiddenCalcs value=" & cmdValidate.Parameters("hiddenCalcs").Value & ">" & vbCrLf)
    Response.Write("<INPUT type=hidden id=txtDeletedFilters name=txtDeletedFilters value=" & cmdValidate.Parameters("deletedFilters").Value & ">" & vbCrLf)
    Response.Write("<INPUT type=hidden id=txtHiddenFilters name=txtHiddenFilters value=" & cmdValidate.Parameters("hiddenFilters").Value & ">" & vbCrLf)
    Response.Write("<INPUT type=hidden id=txtDeletedOrders name=txtDeletedOrders value=" & cmdValidate.Parameters("deletedOrders").Value & ">" & vbCrLf)
    Response.Write("<INPUT type=hidden id=txtJobIDsToHide name=txtJobIDsToHide value=""" & cmdValidate.Parameters("jobsToHide").Value & """>" & vbCrLf)
    Response.Write("<INPUT type=hidden id=txtDeletedPicklists name=txtDeletedPicklists value=" & cmdValidate.Parameters("deletedPicklists").Value & ">" & vbCrLf)
    Response.Write("<INPUT type=hidden id=txtHiddenPicklists name=txtHiddenPicklists value=" & cmdValidate.Parameters("hiddenPicklists").Value & ">" & vbCrLf)

    If cmdValidate.Parameters("errorCode").Value = 1 Then
        Response.Write("			  <tr>" & vbCrLf)
        Response.Write("					<td width=20></td>" & vbCrLf)
        Response.Write("			    <td align=center colspan=3> " & vbCrLf)
        Response.Write("						<H3>Error Saving Report</H3>" & vbCrLf)
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
            Response.Write("						<H3>Error Saving Report</H3>" & vbCrLf)
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
                Response.Write("						<H3>Error Saving Report</H3>" & vbCrLf)
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
                    Response.Write("						<H3>Error Saving Report</H3>" & vbCrLf)
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



    <script ID="clientEventHandlersJS" type="text/javascript">

        function validate_window_onload() {

            //// Hide the 'please wait' message.
            //trPleaseWait1.style.visibility='hidden';
            //trPleaseWait1.style.display='none';
            //trPleaseWait2.style.visibility='hidden';
            //trPleaseWait2.style.display='none';
            //trPleaseWait3.style.visibility='hidden';
            //trPleaseWait3.style.display='none';
            //trPleaseWait4.style.visibility='hidden';
            //trPleaseWait4.style.display='none';
            //trPleaseWait5.style.visibility='hidden';
            //trPleaseWait5.style.display='none';

            //// Resize the grid to show all prompted values.
            //iResizeBy = bdyMain.scrollWidth	- bdyMain.clientWidth;
            //if (bdyMain.offsetWidth + iResizeBy > screen.width) {
            //    window.dialogWidth = new String(screen.width) + "px";
            //}
            //else {
            //    iNewWidth = new Number(window.dialogWidth.substr(0, window.dialogWidth.length-2));
            //    iNewWidth = iNewWidth + iResizeBy;
            //    window.dialogWidth = new String(iNewWidth) + "px";
            //}

            //iResizeBy = bdyMain.scrollHeight	- bdyMain.clientHeight;
            //if (bdyMain.offsetHeight + iResizeBy > screen.height) {
            //    window.dialogHeight = new String(screen.height) + "px";
            //}
            //else {
            //    iNewHeight = new Number(window.dialogHeight.substr(0, window.dialogHeight.length-2));
            //    iNewHeight = iNewHeight + iResizeBy;
            //    window.dialogHeight = new String(iNewHeight) + "px";
            //}

            //iNewLeft = (screen.width - bdyMain.offsetWidth) / 2;
            //iNewTop = (screen.height - bdyMain.offsetHeight) / 2;
            //window.dialogLeft = new String(iNewLeft) + "px";
            //window.dialogTop = new String(iNewTop) + "px";

            if (txtErrorCode.value == 0) {
                var frmSubmit = window.dialogArguments.document.getElementById('frmSend');
                //window.dialogArguments.OpenHR(frmSubmit);
                OpenHR.submitForm(frmSubmit, null, false);
                //window.dialogArguments.document.getElementById('frmSend').submit();
                self.close();
                return;
            }

            /* TM - need to remove hidden filters from the definition */
            if (txtErrorCode.value == 1) {
                // Error, see if we need to remove any columns from the report.
                if (txtHiddenFilters.value.length > 0) {
                    window.dialogArguments.OpenHR.removeFilters(txtHiddenFilters.value);		                  
                }	
                if (txtDeletedFilters.value.length > 0) {
                    window.dialogArguments.OpenHR.removeFilters(txtDeletedFilters.value);		  
                }
            }	
	
            /* JPD - need to remove hidden picklists from the definition */
            if (txtErrorCode.value == 1) {
                // Error, see if we need to remove any columns from the report.
                if (txtHiddenPicklists.value.length > 0) {
                    window.dialogArguments.OpenHR.removePicklists(txtHiddenPicklists.value);		  
                }	
                if (txtDeletedPicklists.value.length > 0) {
                    window.dialogArguments.OpenHR.removePicklists(txtDeletedPicklists.value);		  
                }
            }	
	
            /* TM - need to remove hidden child orders from the definition */
            if (txtErrorCode.value == 1) {
                if (txtDeletedOrders.value.length > 0) {
                    window.dialogArguments.OpenHR.removeChildOrders(txtDeletedOrders.value);		  
                }
            }	
	
            if (txtErrorCode.value == 1) {
                // Error, see if we need to remove any columns from the report.
                if (txtHiddenCalcs.value.length > 0) {
                    window.dialogArguments.OpenHR.removeCalcs(txtHiddenCalcs.value);		  
                }	
                if (txtDeletedCalcs.value.length > 0) {
                    window.dialogArguments.OpenHR.removeCalcs(txtDeletedCalcs.value);		  
                }
            }
	
            //window.dialogArguments.window.parent.frames("menuframe").refreshMenu();
        }

        function overwrite()
        {
            window.dialogArguments.OpenHR.getElementById('frmSend').submit();
            self.close();
        }

        function createNew()
        {
            window.dialogArguments.OpenHR.createNew(self);		  
        }

        function continueSave()
        {
            window.dialogArguments.OpenHR.setJobsToHide(txtJobIDsToHide.value);		  
            window.dialogArguments.OpenHR.getElementById('frmSend').submit();
            self.close();
        }

</script>


<script type="text/javascript">    
    validate_window_onload();
</script>


    </body>

</html>
