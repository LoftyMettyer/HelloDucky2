<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HR.Intranet.Server" %>


<script type="text/javascript">

		function util_validate_picklist_window_onload() {

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

						var frmSend = OpenHR.getForm("workframe", "frmSend");
						OpenHR.submitForm(frmSend);
					closeclick();

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
												<input type="button" value="Cancel" class="btn" name="Cancel" style="WIDTH: 80px" width="80" id="Cancel" />
										</td>
										<td width="20"></td>
								</tr>


								<%
									Dim fDisplay As Boolean = False
									
									Dim objSession As SessionInfo = CType(Session("SessionContext"), SessionInfo)
									Dim objDataAccess As New clsDataAccess(objSession.LoginInfo)

									Dim prmUtilName As New SqlParameter("psUtilName", SqlDbType.VarChar, 255)
									Dim prmUtilID As New SqlParameter("piUtilID", SqlDbType.Int)
									Dim prmTimestamp As New SqlParameter("piTimestamp", SqlDbType.Int)
									Dim prmAccess As New SqlParameter("psAccess", SqlDbType.VarChar, 255)
									Dim prmErrorMsg As New SqlParameter("psErrorMsg", SqlDbType.VarChar, 255)
									Dim prmErrorCode As New SqlParameter("piErrorCode", SqlDbType.VarChar, 255)
									Dim prmBaseTableID As New SqlParameter("piBaseTableID", SqlDbType.Int)
									
									If Request("validatePass") = 1 Then

										prmUtilName.Value = Request("validateName")
										prmUtilID.Value = CleanNumeric(Request("validateUtilID"))
										prmTimestamp.Value = CleanNumeric(Request("validateTimestamp"))
										prmAccess.Value = Request("validateAccess")
										prmErrorMsg.Direction = ParameterDirection.Output
										prmErrorCode.Direction = ParameterDirection.Output

										objDataAccess.ExecuteSP("sp_ASRIntValidatePicklist", prmUtilName, prmUtilID, prmTimestamp, prmAccess, prmErrorMsg, prmErrorCode)

										If prmErrorCode.Value = 1 Then
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
								%>
								<input type="button" value="Yes" class="btn" name="btnYes" style="WIDTH: 80px" width="80" id="btnYes"
										onclick="createNew()" />
								<%
										Response.Write("			    </td>" & vbCrLf)
										Response.Write("					<td width=20></td>" & vbCrLf)
										Response.Write("			    <td align=left> " & vbCrLf)
								%>
							
						

								<input type="button" value="No" class="btn" name="btnNo" style="WIDTH: 80px" width="80" id="btnNo"
										onclick="closeclick();" />
								<%
										Response.Write("			    </td>" & vbCrLf)
										Response.Write("					<td width=20></td>" & vbCrLf)
										Response.Write("			  </tr>" & vbCrLf)
								Else
									If prmErrorCode.Value = 2 Then
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
								%>
								<input type="button" value="Yes" class="btn" name="btnYes" style="WIDTH: 80px" width="80" id="Button1"
										onclick="overwrite()" />
								<%
										Response.Write("			    </td>" & vbCrLf)
										Response.Write("					<td width=20></td>" & vbCrLf)
										Response.Write("			    <td align=left> " & vbCrLf)
								%>
								<input type="button" value="No" class="btn" name="btnNo" style="WIDTH: 80px" width="80" id="Button2"
										onclick="closeclick();"/>
								<%
										Response.Write("			    </td>" & vbCrLf)
										Response.Write("					<td width=20></td>" & vbCrLf)
										Response.Write("				</tr>" & vbCrLf)
								End If
						End If
	

				Else
						If Request("validatePass") = 2 Then

								prmUtilName.Value = Request("validateName")
								prmUtilID.Value = CleanNumeric(Request("validateUtilID"))							
								prmAccess.Value = Request("validateAccess")
								prmBaseTableID.Value = CleanNumeric(Request("validateBaseTableID"))
								prmErrorMsg.Direction = ParameterDirection.Output
								prmErrorCode.Direction = ParameterDirection.Output

								objDataAccess.ExecuteSP("sp_ASRIntValidatePicklist2", prmUtilName, prmUtilID, prmAccess, prmBaseTableID, prmErrorMsg, prmErrorCode)

								If prmErrorCode.Value = 1 Then
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
								%>
								<input type="button" value="Close" class="btn" name="Cancel" style="WIDTH: 80px" width="80" id="Button3"
										onclick="closeclick();"/>
								<%
										Response.Write("			    </td>" & vbCrLf)
										Response.Write("					<td width=20></td>" & vbCrLf)
										Response.Write("			  </tr>" & vbCrLf)
								Else
									If prmErrorCode.Value = 2 Then
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
								%>
								<input type="button" value="Yes" class="btn" name="btnYes" style="WIDTH: 80px" width="80" id="Button4"
										onclick="makeHidden()" />
								<%
										Response.Write("			    </td>" & vbCrLf)
										Response.Write("					<td width=20></td>" & vbCrLf)
										Response.Write("			    <td align=left> " & vbCrLf)
								%>
								<input type="button" value="No" class="btn" name="btnNo" style="WIDTH: 80px" width="80" id="Button5"
										onclick="closeclick();" />
								<%                            
										Response.Write("			    </td>" & vbCrLf)
										Response.Write("					<td width=20></td>" & vbCrLf)
										Response.Write("				</tr>" & vbCrLf)
								End If
						End If
	
				End If
		End If
	
					Response.Write("<input type=hidden id=txtDisplay name=txtDisplay value=" & fDisplay & ">" & vbCrLf)
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
