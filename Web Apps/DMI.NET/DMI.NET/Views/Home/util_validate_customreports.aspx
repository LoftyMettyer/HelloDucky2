<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />

<script src="<%: Url.LatestContent("~/bundles/jQuery")%>" type="text/javascript"></script>
<script src="<%: Url.LatestContent("~/bundles/jQueryUI7")%>" type="text/javascript"></script>

<script src="<%: Url.LatestContent("~/bundles/OpenHR_General")%>" type="text/javascript"></script>

<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />
<link href="<%: Url.LatestContent("~/Content/Site.css")%>" rel="stylesheet" type="text/css" />
<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />
<link id="DMIthemeLink" href="<%: Url.LatestContent("~/Content/themes/" & Session("ui-admin-theme").ToString() & "/jquery-ui.min.css")%>" rel="stylesheet" type="text/css" />
<link href="<%= Url.LatestContent("~/Content/font-awesome.min.css")%>" rel="stylesheet" type="text/css" />
	


<!DOCTYPE html>
<html>
<head runat="server">
		<title>OpenHR</title>
</head>
<body id=bdyMain>
		
		
<table id="ValidatingMessageTable" align=center class="outline" cellPadding=5 cellSpacing=0>
	<tr>
		<td>
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
						Loading...
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
	Dim prmErrorCode As New SqlParameter("@piErrorCode", SqlDbType.VarChar) With {.Direction = ParameterDirection.Output, .Size = -1} 'We need this parameter outside the Try-Catch, see its use below
	Try
		Dim objSession As SessionInfo = CType(Session("SessionContext"), SessionInfo)	'Set session info
		Dim objDataAccess As New clsDataAccess(objSession.LoginInfo) 'Instantiate DataAccess class
	
		Dim prmErrorMsg As New SqlParameter("@psErrorMsg", SqlDbType.VarChar) With {.Direction = ParameterDirection.Output, .Size = -1}
		Dim prmDeletedCalcs As New SqlParameter("@psDeletedCalcs", SqlDbType.VarChar) With {.Direction = ParameterDirection.Output, .Size = -1}
		Dim prmHiddenCalcs As New SqlParameter("@psHiddenCalcs", SqlDbType.VarChar) With {.Direction = ParameterDirection.Output, .Size = -1}
		Dim prmDeletedFilters As New SqlParameter("@psDeletedFilters", SqlDbType.VarChar) With {.Direction = ParameterDirection.Output, .Size = -1}
		Dim prmHiddenFilters As New SqlParameter("@psHiddenFilters", SqlDbType.VarChar) With {.Direction = ParameterDirection.Output, .Size = -1}
		Dim prmDeletedOrders As New SqlParameter("@psDeletedOrders", SqlDbType.VarChar) With {.Direction = ParameterDirection.Output, .Size = -1}
		Dim prmJobIDsToHide As New SqlParameter("@psJobIDsToHide", SqlDbType.VarChar) With {.Direction = ParameterDirection.Output, .Size = -1}
		Dim prmDeletedPicklists As New SqlParameter("@psDeletedPicklists", SqlDbType.VarChar) With {.Direction = ParameterDirection.Output, .Size = -1}
		Dim prmHiddenPicklists As New SqlParameter("@psHiddenPicklists", SqlDbType.VarChar) With {.Direction = ParameterDirection.Output, .Size = -1}
        
		objDataAccess.ExecuteSP("sp_ASRIntValidateReport", _
						New SqlParameter("@psUtilName", SqlDbType.VarChar) With {.Value = Request("validateName"), .Size = 255}, _
						New SqlParameter("@piUtilID", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateUtilID"))}, _
						New SqlParameter("@piTimestamp", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateTimestamp"))}, _
						New SqlParameter("@piBasePicklistID", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateBasePicklist"))}, _
						New SqlParameter("@piBaseFilterID", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateBaseFilter"))}, _
						New SqlParameter("@piEmailGroupID", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateEmailGroup"))}, _
						New SqlParameter("@piParent1PicklistID", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateP1Picklist"))}, _
						New SqlParameter("@piParent1FilterID", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateP1Filter"))}, _
						New SqlParameter("@piParent2PicklistID", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateP2Picklist"))}, _
						New SqlParameter("@piParent2FilterID", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateP2Filter"))},
						New SqlParameter("@piChildFilterID", SqlDbType.VarChar) With {.Value = Request("validateChildFilter"), .Size = 100}, _
						New SqlParameter("@psCalculations", SqlDbType.VarChar) With {.Value = Request("validateCalcs"), .Size = -1}, _
						New SqlParameter("@psHiddenGroups ", SqlDbType.VarChar) With {.Value = Request("validateHiddenGroups"), .Size = -1}, _
						prmErrorMsg, _
						prmErrorCode, _
						prmDeletedCalcs, _
						prmHiddenCalcs, _
						prmDeletedFilters, _
						prmHiddenFilters, _
						prmDeletedOrders, _
						prmJobIDsToHide, _
						prmDeletedPicklists, _
						prmHiddenPicklists _
		)

		Response.Write("<INPUT type=hidden id=txtErrorCode name=txtErrorCode value=" & prmErrorCode.Value.ToString & ">" & vbCrLf)
		Response.Write("<INPUT type=hidden id=txtDeletedCalcs name=txtDeletedCalcs value=" & prmDeletedCalcs.Value.ToString & ">" & vbCrLf)
		Response.Write("<INPUT type=hidden id=txtHiddenCalcs name=txtHiddenCalcs value=" & prmHiddenCalcs.Value.ToString & ">" & vbCrLf)
		Response.Write("<INPUT type=hidden id=txtDeletedFilters name=txtDeletedFilters value=" & prmDeletedFilters.Value.ToString & ">" & vbCrLf)
		Response.Write("<INPUT type=hidden id=txtHiddenFilters name=txtHiddenFilters value=" & prmHiddenFilters.Value.ToString & ">" & vbCrLf)
		Response.Write("<INPUT type=hidden id=txtDeletedOrders name=txtDeletedOrders value=" & prmDeletedOrders.Value.ToString & ">" & vbCrLf)
		Response.Write("<INPUT type=hidden id=txtJobIDsToHide name=txtJobIDsToHide value=""" & prmJobIDsToHide.Value.ToString & """>" & vbCrLf)
		Response.Write("<INPUT type=hidden id=txtDeletedPicklists name=txtDeletedPicklists value=" & prmDeletedPicklists.Value.ToString & ">" & vbCrLf)
		Response.Write("<INPUT type=hidden id=txtHiddenPicklists name=txtHiddenPicklists value=" & prmHiddenPicklists.Value.ToString & ">" & vbCrLf)
	
		If CInt(prmErrorCode.Value) = 1 Then
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
			Response.Write("						" & prmErrorMsg.Value.ToString & vbCrLf)
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
		ElseIf CInt(prmErrorCode.Value) = 2 Then
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
			Response.Write("						" & prmErrorMsg.Value.ToString & vbCrLf)
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
		ElseIf CInt(prmErrorCode.Value) = 3 Then
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
			Response.Write("						" & prmErrorMsg.Value.ToString & vbCrLf)
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
		ElseIf CInt(prmErrorCode.Value) = 4 Then
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
			Response.Write("						" & prmErrorMsg.Value.ToString & vbCrLf)
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
	Catch ex As Exception
		
	End Try
%>
				<tr height=10> 
					<td colspan=5></td>
				</tr>
			</table>
		</td>
	</tr>
</table>

<%
	'Hide "Validating Report" message if we have an error condition, because we displayed the error message above
	If CInt(prmErrorCode.Value) > 0 Then
		Response.Write(String.Concat( _
							"<script text='text/javascript'>",
							"		$('#trPleaseWait1').hide();", _
							"		$('#trPleaseWait2').hide();", _
							"		$('#trPleaseWait3').hide();", _
							"</script>" _
							)
						)
	End If
%>

		<script ID="clientEventHandlersJS" type="text/javascript">

				function validate_window_onload() {

						if (txtErrorCode.value == 0) {
								var frmSubmit = window.dialogArguments.document.getElementById('frmSend');
								//window.dialogArguments.OpenHR(frmSubmit);
								window.dialogArguments.OpenHR.submitForm(frmSubmit, null, false);
								//window.dialogArguments.document.getElementById('frmSend').submit();
								self.close();
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
	$(function () {
		$("input[type=submit], input[type=button], button").button();
		$("input").addClass("ui-widget ui-corner-all");
		$("input").removeClass("text");
		$("input[type=submit], input[type=button], button").removeClass("ui-corner-all");
		$("input[type=submit], input[type=button], button").addClass("ui-corner-tl ui-corner-br");
	});

	validate_window_onload();
</script>


		</body>

</html>
