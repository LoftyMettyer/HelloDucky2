<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>

<link href="<%:Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />
<script src="<%:Url.LatestContent("~/Scripts/jquery/jquery-1.8.3.js")%>" type="text/javascript"></script>
<script src="<%:Url.LatestContent("~/Scripts/openhr.js")%>" type="text/javascript"></script>
<script src="<%:Url.LatestContent("~/Scripts/ctl_SetFont.js")%>" type="text/javascript"></script>
<script src="<%:Url.LatestContent("~/Scripts/ctl_SetStyles.js")%>" type="text/javascript"></script>
<script src="<%:Url.LatestContent("~/Scripts/jquery/jquery-ui-1.9.2.custom.js")%>" type="text/javascript"></script>
<script src="<%:Url.LatestContent("~/Scripts/jquery/jquery.cookie.js")%>" type="text/javascript"></script>
<script src="<%:Url.LatestContent("~/Scripts/FormScripts/menu.js")%>" type="text/javascript"></script>
<script src="<%:Url.LatestContent("~/Scripts/jquery/jquery.ui.touch-punch.min.js")%>" type="text/javascript"></script>
<script src="<%:Url.LatestContent("~/Scripts/jquery/jsTree/jquery.jstree.js")%>" type="text/javascript"></script>
<script id="officebarscript" src="<%: Url.LatestContent("~/Scripts/officebar/jquery.officebar.js")%>" type="text/javascript"></script>

	<script type="text/javascript">
		function util_validate_calendarreport_window_onload() {
			//// Hide the 'please wait' message.
			//window.trPleaseWait1.style.visibility = 'hidden';
			//window.trPleaseWait1.style.display = 'none';
			//window.trPleaseWait2.style.visibility = 'hidden';
			//window.trPleaseWait2.style.display = 'none';
			//window.trPleaseWait3.style.visibility = 'hidden';
			//window.trPleaseWait3.style.display = 'none';
			//window.trPleaseWait4.style.visibility = 'hidden';
			//window.trPleaseWait4.style.display = 'none';
			//window.trPleaseWait5.style.visibility = 'hidden';
			//window.trPleaseWait5.style.display = 'none';

			//// Resize the grid to show all prompted values.
			//var iResizeBy = bdyMain.scrollWidth - bdyMain.clientWidth;
			//if (bdyMain.offsetWidth + iResizeBy > screen.width) {
			//	window.dialogWidth = new String(screen.width) + "px";
			//} else {
			//	var iNewWidth = new Number(window.dialogWidth.substr(0, window.dialogWidth.length - 2));
			//	iNewWidth = iNewWidth + iResizeBy;
			//	window.dialogWidth = new String(iNewWidth) + "px";
			//}

			//iResizeBy = bdyMain.scrollHeight - bdyMain.clientHeight;
			//if (bdyMain.offsetHeight + iResizeBy > screen.height) {
			//	window.dialogHeight = new String(screen.height) + "px";
			//} else {
			//	var iNewHeight = new Number(window.dialogHeight.substr(0, window.dialogHeight.length - 2));
			//	iNewHeight = iNewHeight + iResizeBy;
			//	window.dialogHeight = new String(iNewHeight) + "px";
			//}

			//var iNewLeft = (screen.width - bdyMain.offsetWidth) / 2;
			//var iNewTop = (screen.height - bdyMain.offsetHeight) / 2;
			//window.dialogLeft = new String(iNewLeft) + "px";
			//window.dialogTop = new String(iNewTop) + "px";

			if (txtErrorCode.value == 0) {
				////OpenHR.getElementById("frmSend").submit();
				//var frmSubmit = window.dialogArguments.document.getElementById('frmSend');
				////window.dialogArguments.OpenHR(frmSubmit);
				//OpenHR.submitForm(frmSubmit, null, false);
				////window.dialogArguments.document.getElementById('frmSend').submit();
				//self.close();
				//return;

				var frmSend = window.dialogArguments.OpenHR.getForm("workframe", "frmSend");
				window.dialogArguments.OpenHR.submitForm(frmSend);
				self.close();

			}

			/* need to remove hidden event filters from the definition */
			if (txtErrorCode.value == 1) {
				if (txtHiddenFilters.value.length > 0) {
					window.dialogArguments.removeFilters(window.txtHiddenFilters.value);
				}
				if (txtDeletedFilters.value.length > 0) {
					window.dialogArguments.removeFilters(window.txtDeletedFilters.value);
				}
			}

			/* need to remove hidden picklists from the definition */
			if (txtErrorCode.value == 1) {
				// Error, see if we need to remove any columns from the report.
				if (txtHiddenPicklists.value.length > 0) {
					window.dialogArguments.removePicklists(txtHiddenPicklists.value);
				}
				if (txtDeletedPicklists.value.length > 0) {
					window.dialogArguments.removePicklists(txtDeletedPicklists.value);
				}
			}

			/* need to remove hidden calcs from the definition */
			if (txtErrorCode.value == 1) {
				if (txtHiddenCalcs.value.length > 0) {
					window.dialogArguments.removeCalcs(txtHiddenCalcs.value);
				}
				if (txtDeletedCalcs.value.length > 0) {
					window.dialogArguments.removeCalcs(txtDeleteCalcs.value);
				}
			}
			//window.dialogArguments.menu_refreshMenu();
			window.dialogArguments.OpenHR.refreshMenu("menuframe");
		}

		function overwrite() {
			window.dialogArguments.getElementById.getElementById('frmSend').submit();
			self.close();
		}

		function createNew() {
			window.dialogArguments.createNew(self);
		}

		function continueSave() {
			window.dialogArguments.setJobsToHide(window.txtJibIDsToHide.value);
			window.dialogArguments.submitForm(frmSend);
			self.close();
		}

	</script>

		<div id="Div1" leftmargin="20" topmargin="20" bottommargin="20" rightmargin="5">
		<table align="center" class="outline" cellpadding="5" cellspacing="0">
			<tr>
				<td>
					<table class="invisible" cellspacing="0" cellpadding="0">
						<tr>
							<td colspan="5" height="10"></td>
						</tr>

						<tr id="trPleaseWait1">
							<td width="20"></td>
							<td align="center" colspan="3">Validating Report
							</td>
							<td width="20"></td>
						</tr>

						<tr id="trPleaseWait4" height="10">
							<td colspan="5"></td>
						</tr>

						<tr id="trPleaseWait2">
							<td width="20"></td>
							<td align="center" colspan="3">Loading...
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
									onclick="self.close()" />
							</td>
							<td width="20"></td>
						</tr>


						<%
							
							Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
	
							Dim prmErrorMsg As New SqlParameter("psErrorMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
							Dim prmErrorCode As New SqlParameter("piErrorCode", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
							Dim prmDeletedFilters As New SqlParameter("psDeletedFilters", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
							Dim prmHiddenFilters As New SqlParameter("psHiddenFilters", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
							Dim prmDeletedCalcs As New SqlParameter("psDeletedCalcs", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
							Dim prmHiddenCalcs As New SqlParameter("psHiddenCalcs", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
							Dim prmDeletedPicklists As New SqlParameter("psDeletedPicklists", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
							Dim prmHiddenPicklists As New SqlParameter("psHiddenPicklists", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
							Dim prmJobIDsToHide As New SqlParameter("psJobIDsToHide", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
							
							objDataAccess.ExecuteSP("spASRIntValidateCalendarReport", _
									New SqlParameter("psUtilName", SqlDbType.VarChar, 255) With {.Value = Request("validateName")}, _
									New SqlParameter("piUtilID", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateUtilID"))}, _
									New SqlParameter("piTimestamp", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateTimestamp"))}, _
									New SqlParameter("piBasePicklistID", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateBasePicklist"))}, _
									New SqlParameter("piBaseFilterID", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateBaseFilter"))}, _
									New SqlParameter("piEmailGroupID", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateEmailGroup"))}, _
									New SqlParameter("piDescExprID", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateDescExpr"))}, _
									New SqlParameter("psEventFilterIDs", SqlDbType.VarChar, -1) With {.Value = Request("validateEventFilter")}, _
									New SqlParameter("piCustomStartID", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateCustomStart"))}, _
									New SqlParameter("piCustomEndID", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateCustomEnd"))}, _
									New SqlParameter("psHiddenGroups ", SqlDbType.VarChar, -1) With {.Value = Request("validateHiddenGroups")}, _
									prmErrorMsg, prmErrorCode, prmDeletedFilters, prmHiddenFilters, _
									prmDeletedCalcs, prmHiddenCalcs, prmDeletedPicklists, prmHiddenPicklists, prmJobIDsToHide)
	
							Response.Write("<input type=hidden id=txtErrorCode name=txtErrorCode value=" & prmErrorCode.Value & ">" & vbCrLf)
							Response.Write("<input type=hidden id=txtDeletedFilters name=txtDeletedFilters value=" & prmDeletedFilters.Value & ">" & vbCrLf)
							Response.Write("<input type=hidden id=txtHiddenFilters name=txtHiddenFilters value=" & prmHiddenFilters.Value & ">" & vbCrLf)
							Response.Write("<input type=hidden id=txtDeletedCalcs name=txtDeletedCalcs value=" & prmDeletedCalcs.Value & ">" & vbCrLf)
							Response.Write("<input type=hidden id=txtHiddenCalcs name=txtHiddenCalcs value=" & prmHiddenCalcs.Value & ">" & vbCrLf)
							Response.Write("<input type=hidden id=txtDeletedPicklists name=txtDeletedPicklists value=" & prmDeletedPicklists.Value & ">" & vbCrLf)
							Response.Write("<input type=hidden id=txtHiddenPicklists name=txtHiddenPicklists value=" & prmHiddenPicklists.Value & ">" & vbCrLf)
							Response.Write("<input type=hidden id=txtJobIDsToHide name=txtJobIDsToHide value=""" & prmJobIDsToHide.Value & """>" & vbCrLf)

							If prmErrorCode.Value = 1 Then
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
								Response.Write("						" & prmErrorMsg.Value & vbCrLf)
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
								If prmErrorCode.Value = 2 Then
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
									Response.Write("						" & prmErrorMsg.Value & vbCrLf)
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
									If prmErrorCode.Value = 3 Then
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
										Response.Write("						" & prmErrorMsg.Value & vbCrLf)
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
										If prmErrorCode.Value = 4 Then
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
											Response.Write("						" & prmErrorMsg.Value & vbCrLf)
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
	
							'cmdValidate = Nothing
						%>
						<tr height="10">
							<td colspan="5"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</div>

	<script type="text/javascript">
		util_validate_calendarreport_window_onload();
	</script>


<%--</body>
</html>--%>
