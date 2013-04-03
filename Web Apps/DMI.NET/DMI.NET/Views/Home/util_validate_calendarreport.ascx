<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script type="text/javascript">
	function util_validate_calendarreport_window_onload() {
		{
			// Hide the 'please wait' message.
			window.trPleaseWait1.style.visibility = 'hidden';
			window.trPleaseWait1.style.display = 'none';
			window.trPleaseWait2.style.visibility = 'hidden';
			window.trPleaseWait2.style.display = 'none';
			window.trPleaseWait3.style.visibility = 'hidden';
			window.trPleaseWait3.style.display = 'none';
			window.trPleaseWait4.style.visibility = 'hidden';
			window.trPleaseWait4.style.display = 'none';
			window.trPleaseWait5.style.visibility = 'hidden';
			window.trPleaseWait5.style.display = 'none';

			// Resize the grid to show all prompted values.
			iResizeBy = bdyMain.scrollWidth - bdyMain.clientWidth;
			if (bdyMain.offsetWidth + iResizeBy > screen.width) {
				window.dialogWidth = new String(screen.width) + "px";
			} else {
				iNewWidth = new Number(window.dialogWidth.substr(0, window.dialogWidth.length - 2));
				iNewWidth = iNewWidth + iResizeBy;
				window.dialogWidth = new String(iNewWidth) + "px";
			}

			iResizeBy = bdyMain.scrollHeight - bdyMain.clientHeight;
			if (bdyMain.offsetHeight + iResizeBy > screen.height) {
				window.dialogHeight = new String(screen.height) + "px";
			} else {
				iNewHeight = new Number(window.dialogHeight.substr(0, window.dialogHeight.length - 2));
				iNewHeight = iNewHeight + iResizeBy;
				window.dialogHeight = new String(iNewHeight) + "px";
			}

			iNewLeft = (screen.width - bdyMain.offsetWidth) / 2;
			iNewTop = (screen.height - bdyMain.offsetHeight) / 2;
			window.dialogLeft = new String(iNewLeft) + "px";
			window.dialogTop = new String(iNewTop) + "px";

			if (txtErrorCode.value == 0) {
				window.dialogArguments.document.getElementById('frmSend').submit();
				self.close();
				return;
			}

			/* need to remove hidden event filters from the definition */
			if (txtErrorCode.value == 1) {
				if (txtHiddenFilters.value.length > 0) {
					window.dialogArguments.OpenHR.removeFilters(window.txtHiddenFilters.value);
				}
				if (txtDeletedFilters.value.length > 0) {
					window.dialogArguments.OpenHR.removeFilters(window.txtDeletedFilters.value);
				}
			}

			/* need to remove hidden picklists from the definition */
			if (txtErrorCode.value == 1) {
				// Error, see if we need to remove any columns from the report.
				if (txtHiddenPicklists.value.length > 0) {
					window.dialogArguments.OpenHR.removePicklists(txtHiddenPicklists.value);
				}
				if (txtDeletedPicklists.value.length > 0) {
					window.dialogArguments.OpenHR.removePicklists(txtDeletedPicklists.value);
				}
			}

			/* need to remove hidden calcs from the definition */
			if (txtErrorCode.value == 1) {
				if (txtHiddenCalcs.value.length > 0) {
					window.dialogArguments.OpenHR.removeCalcs(txtHiddenCalcs.value);
				}
				if (txtDeletedCalcs.value.length > 0) {
					window.dialogArguments.OpenHR.removeCalcs(txtDeleteCalcs.value);
				}
			}
			window.dialogArguments.menu_refreshMenu();
		}

		function overwrite() {
			window.dialogArguments.OpenHR.getElementById.getElementById('frmSend').submit();
			self.close();
		}

		function createNew() {
			window.dialogArguments.OpenHR.createNew(self);
		}

		function continueSave() {
			window.dialogArguments.setJobsToHide(window.txtJibIDsToHide.value);
			window.dialogArguments.OpenHR.submitForm(frmSend);
			self.close();
		}
	}
</script>

<div bgcolor='<%=session("ConvertedDesktopColour")%>' onload="return window_onload()" id="bdyMain" leftmargin="20" topmargin="20" bottommargin="20" rightmargin="5">
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
						Dim cmdValidate = CreateObject("ADODB.Command")
						cmdValidate.CommandText = "spASRIntValidateCalendarReport"
						cmdValidate.CommandType = 4	' Stored Procedure
						cmdValidate.ActiveConnection = Session("databaseConnection")

						Dim prmUtilName = cmdValidate.CreateParameter("utilName", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
						cmdValidate.Parameters.Append(prmUtilName)
						prmUtilName.value = Request("validateName")

						Dim prmUtilID = cmdValidate.CreateParameter("utilID", 3, 1)	'3=integer, 1=input
						cmdValidate.Parameters.Append(prmUtilID)
						prmUtilID.value = CleanNumeric(Request("validateUtilID"))

						Dim prmTimestamp = cmdValidate.CreateParameter("timestamp", 3, 1)	'3=integer, 1=input
						cmdValidate.Parameters.Append(prmTimestamp)
						prmTimestamp.value = CleanNumeric(Request("validateTimestamp"))

						Dim prmBasePicklist = cmdValidate.CreateParameter("basePicklist", 3, 1)	'3=integer, 1=input
						cmdValidate.Parameters.Append(prmBasePicklist)
						prmBasePicklist.value = CleanNumeric(Request("validateBasePicklist"))

						Dim prmBaseFilter = cmdValidate.CreateParameter("baseFilter", 3, 1)	'3=integer, 1=input
						cmdValidate.Parameters.Append(prmBaseFilter)
						prmBaseFilter.value = CleanNumeric(Request("validateBaseFilter"))

						Dim prmEmailGroup = cmdValidate.CreateParameter("emailGroup", 3, 1)	'3=integer, 1=input
						cmdValidate.Parameters.Append(prmEmailGroup)
						prmEmailGroup.value = CleanNumeric(Request("validateEmailGroup"))

						Dim prmDescExpr = cmdValidate.CreateParameter("descExpr", 3, 1)	'3=integer, 1=input
						cmdValidate.Parameters.Append(prmDescExpr)
						prmDescExpr.value = CleanNumeric(Request("validateDescExpr"))

						Dim prmEventFilter = cmdValidate.CreateParameter("eventFilter", 200, 1, 8000)	'200=varchar, 1=input, 8000=size
						cmdValidate.Parameters.Append(prmEventFilter)
						prmEventFilter.value = Request("validateEventFilter")

						Dim prmCustomStart = cmdValidate.CreateParameter("customStart", 3, 1)	'3=integer, 1=input
						cmdValidate.Parameters.Append(prmCustomStart)
						prmCustomStart.value = CleanNumeric(Request("validateCustomStart"))

						Dim prmCustomEnd = cmdValidate.CreateParameter("customEnd", 3, 1)	'3=integer, 1=input
						cmdValidate.Parameters.Append(prmCustomEnd)
						prmCustomEnd.value = CleanNumeric(Request("validateCustomEnd"))
	
						Dim prmHiddenGroups = cmdValidate.CreateParameter("hiddenGroups", 200, 1, 8000)	'200=varchar, 1=input, 8000=size
						cmdValidate.Parameters.Append(prmHiddenGroups)
						prmHiddenGroups.value = Request("validateHiddenGroups")

						Dim prmErrorMsg = cmdValidate.CreateParameter("errorMsg", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
						cmdValidate.Parameters.Append(prmErrorMsg)

						Dim prmErrorCode = cmdValidate.CreateParameter("errorCode", 3, 2)	'3=integer, 2=output
						cmdValidate.Parameters.Append(prmErrorCode)
	
						Dim prmDeletedFilters = cmdValidate.CreateParameter("deletedFilters", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
						cmdValidate.Parameters.Append(prmDeletedFilters)

						Dim prmHiddenFilters = cmdValidate.CreateParameter("hiddenFilters", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
						cmdValidate.Parameters.Append(prmHiddenFilters)
	
						Dim prmDeletedCalcs = cmdValidate.CreateParameter("deletedCalcs", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
						cmdValidate.Parameters.Append(prmDeletedCalcs)

						Dim prmHiddenCalcs = cmdValidate.CreateParameter("hiddenCalcs", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
						cmdValidate.Parameters.Append(prmHiddenCalcs)

						Dim prmDeletedPicklists = cmdValidate.CreateParameter("deletedPicklists", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
						cmdValidate.Parameters.Append(prmDeletedPicklists)

						Dim prmHiddenPicklists = cmdValidate.CreateParameter("hiddenPicklists", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
						cmdValidate.Parameters.Append(prmHiddenPicklists)
	
						Dim prmJobIDsToHide = cmdValidate.CreateParameter("jobsToHide", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
						cmdValidate.Parameters.Append(prmJobIDsToHide)

						Err.Number = 0
						cmdValidate.Execute()

						Response.Write("<INPUT type=hidden id=txtErrorCode name=txtErrorCode value=" & cmdValidate.Parameters("errorCode").Value & ">" & vbCrLf)
						Response.Write("<INPUT type=hidden id=txtDeletedFilters name=txtDeletedFilters value=" & cmdValidate.Parameters("deletedFilters").Value & ">" & vbCrLf)
						Response.Write("<INPUT type=hidden id=txtHiddenFilters name=txtHiddenFilters value=" & cmdValidate.Parameters("hiddenFilters").Value & ">" & vbCrLf)
						Response.Write("<INPUT type=hidden id=txtDeletedCalcs name=txtDeletedCalcs value=" & cmdValidate.Parameters("deletedCalcs").Value & ">" & vbCrLf)
						Response.Write("<INPUT type=hidden id=txtHiddenCalcs name=txtHiddenCalcs value=" & cmdValidate.Parameters("hiddenCalcs").Value & ">" & vbCrLf)
						Response.Write("<INPUT type=hidden id=txtDeletedPicklists name=txtDeletedPicklists value=" & cmdValidate.Parameters("deletedPicklists").Value & ">" & vbCrLf)
						Response.Write("<INPUT type=hidden id=txtHiddenPicklists name=txtHiddenPicklists value=" & cmdValidate.Parameters("hiddenPicklists").Value & ">" & vbCrLf)
						Response.Write("<INPUT type=hidden id=txtJobIDsToHide name=txtJobIDsToHide value=""" & cmdValidate.Parameters("jobsToHide").Value & """>" & vbCrLf)

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

