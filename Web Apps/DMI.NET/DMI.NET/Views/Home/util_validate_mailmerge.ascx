<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<script type="text/javascript">
	function util_validate_mailmerge_window_onload() {

		if (txtErrorCode.value == 0) {
			var frmSend = window.dialogArguments.OpenHR.getForm("workframe", "frmSend");
			window.dialogArguments.OpenHR.submitForm(frmSend);
			self.close();
			return;
		}

		if (txtErrorCode.value == 1) {
			if (txtHiddenCalcs.value.length > 0) {
				window.dialogArguments.OpenHR.removeCalcs(txtHiddenCalcs.value);
				
			}
			if (txtDeletedCalcs.value.length > 0) {
				window.dialogArguments.OpenHR.removeCalcs(txtDeletedCalcs.value);
			}
		}

		try {
			window.dialogArguments.OpenHR.refreshMenu("menuframe");
		}
		catch (e) {
		}
	}

	function overwrite() {
		window.dialogArguments.OpenHR.getElementById("frmSend").submit();	
		self.close();
	}

	function createNew() {
		window.dialogArguments.OpenHR.getFrame("workframe").createNew(self);
	}

	function continueSave() {
		window.dialogArguments.OpenHR.getFrame("workframe").setJobsToHide(txtJobIDsToHide.value);
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

						Dim prmErrorMsg = New SqlParameter("psErrorMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
						Dim prmErrorCode = New SqlParameter("piErrorCode", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
						Dim prmDeletedCalcs = New SqlParameter("psDeletedCalcs", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
						Dim prmHiddenCalcs = New SqlParameter("psHiddenCalcs", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
						Dim prmJobIDsToHide = New SqlParameter("psJobIDsToHide", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

						objDataAccess.ExecuteSP("sp_ASRIntValidateMailMerge" _
							, New SqlParameter("@psUtilName", SqlDbType.VarChar, 255) With {.Value = Request("validateName")} _
							, New SqlParameter("@piUtilID", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateUtilID"))} _
							, New SqlParameter("@piTimestamp", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateTimestamp"))} _
							, New SqlParameter("@piBasePicklistID", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateBasePicklist"))} _
							, New SqlParameter("@piBaseFilterID", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateBaseFilter"))} _
							, New SqlParameter("@psCalculations", SqlDbType.VarChar, -1) With {.Value = Request("validateCalcs")} _
							, New SqlParameter("@psHiddenGroups", SqlDbType.VarChar, -1) With {.Value = Request("validateHiddenGroups")} _
							, prmErrorMsg, prmErrorCode, prmDeletedCalcs, prmHiddenCalcs, prmJobIDsToHide)
												
						Response.Write("<input type=hidden id=txtErrorCode name=txtErrorCode value=" & prmErrorCode.Value & ">" & vbCrLf)
						Response.Write("<input type=hidden id=txtDeletedCalcs name=txtDeletedCalcs value=" & prmDeletedCalcs.Value & ">" & vbCrLf)
						Response.Write("<input type=hidden id=txtHiddenCalcs name=txtHiddenCalcs value=" & prmHiddenCalcs.Value & ">" & vbCrLf)
						Response.Write("<input type=hidden id=txtJobIDsToHide name=txtJobIDsToHide value=""" & prmJobIDsToHide.Value & """>" & vbCrLf)

						If prmErrorCode.Value = 1 Then
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
							Response.Write("						" & prmErrorMsg.Value & vbCrLf)
							Response.Write("			    </td>" & vbCrLf)
							Response.Write("					<td width=20></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)
							Response.Write("			  </td>" & vbCrLf)
							Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)
							Response.Write("			  <tr> " & vbCrLf)
							Response.Write("					<td width=20></td>" & vbCrLf)
							Response.Write("			    <td align=center colspan=3> " & vbCrLf)
							Response.Write("    				    <input TYPE=button VALUE=Close class=""btn"" NAME=Cancel style=""WIDTH: 80px"" width=80 id=Cancel" & vbCrLf)
							Response.Write("    				        OnClick=""self.close()""/>" & vbCrLf)
							Response.Write("			    </td>" & vbCrLf)
							Response.Write("					<td width=20></td>" & vbCrLf)
							Response.Write("			  </tr>" & vbCrLf)
						Else
							If prmErrorCode.Value = 2 Then
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
								Response.Write("						" & prmErrorMsg.Value & vbCrLf)
								Response.Write("			    </td>" & vbCrLf)
								Response.Write("					<td width=20></td>" & vbCrLf)
								Response.Write("			  </tr>" & vbCrLf)
								Response.Write("			  </td>" & vbCrLf)
								Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
								Response.Write("			  </tr>" & vbCrLf)
								Response.Write("			  <tr> " & vbCrLf)
								Response.Write("					<td width=20></td>" & vbCrLf)
								Response.Write("			    <td align=right> " & vbCrLf)
								Response.Write("    				    <input TYPE=button VALUE=Yes class=""btn"" NAME=btnYes style=""WIDTH: 80px"" width=80 id=btnYes" & vbCrLf)
								Response.Write("    				        OnClick=""createNew()""/>" & vbCrLf)
								Response.Write("			    </td>" & vbCrLf)
								Response.Write("					<td width=20></td>" & vbCrLf)
								Response.Write("			    <td align=left> " & vbCrLf)
								Response.Write("    				    <input TYPE=button VALUE=No class=""btn"" NAME=btnNo style=""WIDTH: 80px"" width=80 id=btnNo" & vbCrLf)
								Response.Write("    				        OnClick=""self.close()""/>" & vbCrLf)
								Response.Write("			    </td>" & vbCrLf)
								Response.Write("					<td width=20></td>" & vbCrLf)
								Response.Write("			  </tr>" & vbCrLf)

							Else
								If prmErrorCode.Value = 3 Then
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
									Response.Write("						" & prmErrorMsg.Value & vbCrLf)
									Response.Write("			    </td>" & vbCrLf)
									Response.Write("					<td width=20></td>" & vbCrLf)
									Response.Write("			  </tr>" & vbCrLf)
									Response.Write("			  </td>" & vbCrLf)
									Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
									Response.Write("			  </tr>" & vbCrLf)
									Response.Write("			  <tr> " & vbCrLf)
									Response.Write("					<td width=20></td>" & vbCrLf)
									Response.Write("			    <td align=right> " & vbCrLf)
									Response.Write("    				    <input TYPE=button VALUE=Yes class=""btn"" NAME=btnYes style=""WIDTH: 80px"" width=80 id=btnYes" & vbCrLf)
									Response.Write("    				        OnClick=""overwrite()""/>" & vbCrLf)
									Response.Write("			    </td>" & vbCrLf)
									Response.Write("					<td width=20></td>" & vbCrLf)
									Response.Write("			    <td align=left> " & vbCrLf)
									Response.Write("    				    <input TYPE=button VALUE=No class=""btn"" NAME=btnNo style=""WIDTH: 80px"" width=80 id=btnNo" & vbCrLf)
									Response.Write("    				        OnClick=""self.close()"" />" & vbCrLf)
									Response.Write("			    </td>" & vbCrLf)
									Response.Write("					<td width=20></td>" & vbCrLf)
									Response.Write("				</tr>" & vbCrLf)
								Else
									If prmErrorCode.Value = 4 Then
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
										Response.Write("						" & prmErrorMsg.Value & vbCrLf)
										Response.Write("			    </td>" & vbCrLf)
										Response.Write("					<td width=20></td>" & vbCrLf)
										Response.Write("			  </tr>" & vbCrLf)
										Response.Write("			  </td>" & vbCrLf)
										Response.Write("					<td height=20 colspan=5></td>" & vbCrLf)
										Response.Write("			  </tr>" & vbCrLf)
										Response.Write("			  <tr> " & vbCrLf)
										Response.Write("					<td width=20></td>" & vbCrLf)
										Response.Write("			    <td align=right> " & vbCrLf)
										Response.Write("    				    <input TYPE=button VALUE=Yes class=""btn"" NAME=btnYes style=""WIDTH: 80px"" width=80 id=btnYes" & vbCrLf)
										Response.Write("    				        OnClick=""continueSave()""/>" & vbCrLf)
										Response.Write("			    </td>" & vbCrLf)
										Response.Write("					<td width=20></td>" & vbCrLf)
										Response.Write("			    <td align=left> " & vbCrLf)
										Response.Write("    				    <input TYPE=button VALUE=No class=""btn"" NAME=btnNo style=""WIDTH: 80px"" width=80 id=btnNo" & vbCrLf)
										Response.Write("    				        OnClick=""self.close()""/>" & vbCrLf)
										Response.Write("			    </td>" & vbCrLf)
										Response.Write("					<td width=20></td>" & vbCrLf)
										Response.Write("				</tr>" & vbCrLf)
									End If
								End If
							End If
						End If

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