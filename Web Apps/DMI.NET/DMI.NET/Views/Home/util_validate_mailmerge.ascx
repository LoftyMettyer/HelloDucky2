<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script type="text/javascript">
	//<script ID="clientEventHandlersJS" type="text/javascript">
	function util_validate_mailmerge_window_onload() {
		// Hide the 'please wait' message.
		//trPleaseWait1.style.visibility = 'hidden';
		//trPleaseWait1.style.display = 'none';
		//trPleaseWait2.style.visibility = 'hidden';
		//trPleaseWait2.style.display = 'none';
		//trPleaseWait3.style.visibility = 'hidden';
		//trPleaseWait3.style.display = 'none';
		//trPleaseWait4.style.visibility = 'hidden';
		//trPleaseWait4.style.display = 'none';
		//trPleaseWait5.style.visibility = 'hidden';
		//trPleaseWait5.style.display = 'none';

		//// Resize the grid to show all prompted values.
		//iResizeBy = bdyMain.scrollWidth - bdyMain.clientWidth;
		//if (bdyMain.offsetWidth + iResizeBy > screen.width) {
		//	window.dialogWidth = new String(screen.width) + "px";
		//}
		//else {
		//	iNewWidth = new Number(window.dialogWidth.substr(0, window.dialogWidth.length - 2));
		//	iNewWidth = iNewWidth + iResizeBy;
		//	window.dialogWidth = new String(iNewWidth) + "px";
		//}

		//iResizeBy = bdyMain.scrollHeight - bdyMain.clientHeight;
		//if (bdyMain.offsetHeight + iResizeBy > screen.height) {
		//	window.dialogHeight = new String(screen.height) + "px";
		//}
		//else {
		//	iNewHeight = new Number(window.dialogHeight.substr(0, window.dialogHeight.length - 2));
		//	iNewHeight = iNewHeight + iResizeBy;
		//	window.dialogHeight = new String(iNewHeight) + "px";
		//}

		//iNewLeft = (screen.width - bdyMain.offsetWidth) / 2;
		//iNewTop = (screen.height - bdyMain.offsetHeight) / 2;
		//window.dialogLeft = new String(iNewLeft) + "px";
		//window.dialogTop = new String(iNewTop) + "px";
		debugger;
		if (txtErrorCode.value == 0) {
			//window.dialogArguments.document.getElementById('frmSend').submit();
			//window.dialogArguments.OpenHR.getElementById("frmSend").submit();
			var frmSend = window.dialogArguments.OpenHR.getForm("workframe", "frmSend");
			window.dialogArguments.OpenHR.submitForm(frmSend);
			self.close();
			return;
		}

		if (txtErrorCode.value == 1) {
			// Error, see if we need to remove any columns from the mail merge.
			if (txtHiddenCalcs.value.length > 0) {
				//window.dialogArguments.window.parent.frames("workframe").removeCalcs(txtHiddenCalcs.value);
				window.dialogArguments.OpenHR.removeCalcs(txtHiddenCalcs.value);
				
			}
			if (txtDeletedCalcs.value.length > 0) {
				//window.dialogArguments.window.parent.frames("workframe").removeCalcs(txtDeletedCalcs.value);
				window.dialogArguments.OpenHR.removeCalcs(txtDeletedCalcs.value);
			}
		}

		try {
			//window.dialogArguments.window.parent.frames("menuframe").refreshMenu();
			window.dialogArguments.OpenHR.refreshMenu("menuframe");
		}
		catch (e) {
		}
	}

	function overwrite() {
		//window.dialogArguments.document.getElementById('frmSend').submit();
		window.dialogArguments.OpenHR.getElementById("frmSend").submit();
		
		self.close();
	}

	function createNew() {
		//window.dialogArguments.window.parent.frames("workframe").createNew(self);
		window.dialogArguments.OpenHR.getFrame("workframe").createNew(self);
	}

	function continueSave() {
		//window.dialogArguments.window.parent.frames("workframe").setJobsToHide(txtJobIDsToHide.value);
		window.dialogArguments.OpenHR.getFrame("workframe").setJobsToHide(txtJobIDsToHide.value);
		//window.dialogArguments.document.getElementById('frmSend').submit();
		window.dialogArguments.OpenHR.getElementById("frmSend").submit();
		self.close();
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
						<td align="center" colspan="3">Validating Mail Merge
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
						cmdValidate.CommandText = "sp_ASRIntValidateMailMerge"
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

						Dim prmCalcs = cmdValidate.CreateParameter("calcs", 200, 1, 8000)	'200=varchar, 1=input, 8000=size
						cmdValidate.Parameters.Append(prmCalcs)
						prmCalcs.value = Request("validateCalcs")

						Dim prmHiddenGroups = cmdValidate.CreateParameter("hiddenGroups", 200, 1, 8000)	'200=varchar, 1=input, 8000=size
						cmdValidate.Parameters.Append(prmHiddenGroups)
						prmHiddenGroups.value = Request("validateHiddenGroups")

						Dim prmErrorMsg = cmdValidate.CreateParameter("errorMsg", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
						cmdValidate.Parameters.Append(prmErrorMsg)

						Dim prmErrorCode = cmdValidate.CreateParameter("errorCode", 3, 2)	'3=integer, 2=output
						cmdValidate.Parameters.Append(prmErrorCode)

						Dim prmDeletedCalcs = cmdValidate.CreateParameter("deletedCalcs", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
						cmdValidate.Parameters.Append(prmDeletedCalcs)

						Dim prmHiddenCalcs = cmdValidate.CreateParameter("hiddenCalcs", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
						cmdValidate.Parameters.Append(prmHiddenCalcs)

						Dim prmJobIDsToHide = cmdValidate.CreateParameter("jobsToHide", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
						cmdValidate.Parameters.Append(prmJobIDsToHide)

						Err.Number = 0
						cmdValidate.Execute()

						Response.Write("<INPUT type=hidden id=txtErrorCode name=txtErrorCode value=" & cmdValidate.Parameters("errorCode").Value & ">" & vbCrLf)
						Response.Write("<INPUT type=hidden id=txtDeletedCalcs name=txtDeletedCalcs value=" & cmdValidate.Parameters("deletedCalcs").Value & ">" & vbCrLf)
						Response.Write("<INPUT type=hidden id=txtHiddenCalcs name=txtHiddenCalcs value=" & cmdValidate.Parameters("hiddenCalcs").Value & ">" & vbCrLf)
						Response.Write("<INPUT type=hidden id=txtJobIDsToHide name=txtJobIDsToHide value=""" & cmdValidate.Parameters("jobsToHide").Value & """>" & vbCrLf)

						If cmdValidate.Parameters("errorCode").Value = 1 Then
							Response.Write("			  </td>" & vbCrLf)
							Response.Write("					<td width=20></td>" & vbCrLf)
							Response.Write("			    <td align=center colspan=3> " & vbCrLf)
							Response.Write("						<H3>Error Saving Mail Merge</H3>" & vbCrLf)
							Response.Write("			    </td>" & vbCrLf)
							Response.Write("					<td width=20></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)
							Response.Write("			  </td>" & vbCrLf)
							Response.Write("					<td width=20></td>" & vbCrLf)
							Response.Write("			    <td align=center colspan=3> " & vbCrLf)
							Response.Write("						" & cmdValidate.Parameters("errorMsg").Value & vbCrLf)
							Response.Write("			    </td>" & vbCrLf)
							Response.Write("					<td width=20></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)
							Response.Write("			  </td>" & vbCrLf)
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
								Response.Write("			  </td>" & vbCrLf)
								Response.Write("					<td width=20></td>" & vbCrLf)
								Response.Write("			    <td align=center colspan=3> " & vbCrLf)
								Response.Write("						<H3>Error Saving Mail Merge</H3>" & vbCrLf)
								Response.Write("			    </td>" & vbCrLf)
								Response.Write("					<td width=20></td>" & vbCrLf)
								Response.Write("			  </tr>" & vbCrLf)
								Response.Write("			  </td>" & vbCrLf)
								Response.Write("					<td width=20></td>" & vbCrLf)
								Response.Write("			    <td align=center colspan=3> " & vbCrLf)
								Response.Write("						" & cmdValidate.Parameters("errorMsg").Value & vbCrLf)
								Response.Write("			    </td>" & vbCrLf)
								Response.Write("					<td width=20></td>" & vbCrLf)
								Response.Write("			  </tr>" & vbCrLf)
								Response.Write("			  </td>" & vbCrLf)
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
									Response.Write("			  </td>" & vbCrLf)
									Response.Write("					<td width=20></td>" & vbCrLf)
									Response.Write("			    <td align=center colspan=3> " & vbCrLf)
									Response.Write("						<H3>Error Saving Mail Merge</H3>" & vbCrLf)
									Response.Write("			    </td>" & vbCrLf)
									Response.Write("					<td width=20></td>" & vbCrLf)
									Response.Write("			  </tr>" & vbCrLf)
									Response.Write("			  </td>" & vbCrLf)
									Response.Write("					<td width=20></td>" & vbCrLf)
									Response.Write("			    <td align=center colspan=3> " & vbCrLf)
									Response.Write("						" & cmdValidate.Parameters("errorMsg").Value & vbCrLf)
									Response.Write("			    </td>" & vbCrLf)
									Response.Write("					<td width=20></td>" & vbCrLf)
									Response.Write("			  </tr>" & vbCrLf)
									Response.Write("			  </td>" & vbCrLf)
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
										Response.Write("			  </td>" & vbCrLf)
										Response.Write("					<td width=20></td>" & vbCrLf)
										Response.Write("			    <td align=center colspan=3> " & vbCrLf)
										Response.Write("						<H3>Error Saving Mail Merge</H3>" & vbCrLf)
										Response.Write("			    </td>" & vbCrLf)
										Response.Write("					<td width=20></td>" & vbCrLf)
										Response.Write("			  </tr>" & vbCrLf)
										Response.Write("			  </td>" & vbCrLf)
										Response.Write("					<td width=20></td>" & vbCrLf)
										Response.Write("			    <td align=center colspan=3> " & vbCrLf)
										Response.Write("						" & cmdValidate.Parameters("errorMsg").Value & vbCrLf)
										Response.Write("			    </td>" & vbCrLf)
										Response.Write("					<td width=20></td>" & vbCrLf)
										Response.Write("			  </tr>" & vbCrLf)
										Response.Write("			  </td>" & vbCrLf)
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
					<tr height="10">
						<td colspan="5"></td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</div>

<script type="text/javascript">
	util_validate_mailmerge_window_onload();
</script>