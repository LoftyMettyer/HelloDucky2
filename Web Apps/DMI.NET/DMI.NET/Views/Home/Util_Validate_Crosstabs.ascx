<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@Import namespace="DMI.NET" %>

<script type="text/javascript">
<%--	<%Html.RenderPartial("Util_Def_Crosstabs/dialog")%>--%>
</script>

<script type="text/javascript">
	function util_validate_crosstab_window_onload() {
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
	}
</script>

<%--<div bgcolor='<%=session("ConvertedDesktopColour")%>' onload="return window_onload()" id=bdyMain leftmargin=20 topmargin=20 bottommargin=20 rightmargin=5> --%>
<div bgcolor='<%=session("ConvertedDesktopColour")%>' id="Div1" leftmargin="20" topmargin="20" bottommargin="20" rightmargin="5">
	<table align="center" class="outline" cellpadding="5" cellspacing="0">
		<tr>
			<td>
				<table class="invisible" cellspacing="0" cellpadding="0">
					<tr>
						<td colspan="5" height="10"></td>
					</tr>

					<tr id="trPleaseWait1">
						<td width="20"></td>
						<td align="center" colspan="3">Validating Cross Tab
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
							<input type="button" value="Cancel" class="btn" name="Cancel" style="WIDTH: 80px" width="80" id="Cancel" onclick="self.close()" />
						</td>
						<td width="20"></td>
					</tr>


<%
	Dim cmdValidate = CreateObject("ADODB.Command")
	cmdValidate.CommandText = "sp_ASRIntValidateCrossTab"
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

	Dim prmJobIDsToHide = cmdValidate.CreateParameter("jobsToHide", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
	cmdValidate.Parameters.Append(prmJobIDsToHide)

	Err.Clear()
	cmdValidate.Execute()

	Dim ResponseString As String = String.Concat( _
		"<input type='hidden' id='txtErrorCode' 'name='txtErrorCode' value='", cmdValidate.Parameters("errorCode").Value, "'>", vbCrLf, _
		"<input type='hidden' id='txtDeletedFilters' name='txtDeletedFilters' value='", cmdValidate.Parameters("deletedFilters").Value, "'>", vbCrLf, _
		"<input type='hidden' id='txtHiddenFilters' name='txtHiddenFilters' value='", cmdValidate.Parameters("hiddenFilters").Value, "'>", vbCrLf, _
		"<input type='hidden' id='txtJobIDsToHide' name='txtJobIDsToHide' value='", cmdValidate.Parameters("jobsToHide").Value, "'>", vbCrLf)
	
	'The following bit is common if an error code was returned
	If cmdValidate.Parameters("errorCode").Value <> 0 Then
		ResponseString = String.Concat(ResponseString, _
			"			  <tr>", vbCrLf, _
			"					<td width='20'></td>", vbCrLf, _
			"			    <td align='center' colspan='3'> ", vbCrLf, _
			"						<h3>Error Saving Cross Tab</h3>", vbCrLf, _
			"			    </td>", vbCrLf, _
			"					<td width='20'></td>", vbCrLf, _
			"			  </tr>", vbCrLf, _
			"			  <tr>", vbCrLf, _
			"					<td width='20'></td>", vbCrLf, _
			"			    <td align='center' colspan='3'> ", vbCrLf, _
			"						" & cmdValidate.Parameters("errorMsg").Value, vbCrLf, _
			"			    </td>", vbCrLf, _
			"					<td width='20'></td>", vbCrLf, _
			"			  </tr>", vbCrLf, _
			"			  <tr>", vbCrLf, _
			"					<td height='20' colspan='5'></td>", vbCrLf, _
			"			  </tr>", vbCrLf, _
			"			  <tr> ", vbCrLf, _
			"					<td width='20'></td>", vbCrLf)
	End If
	
	If cmdValidate.Parameters("errorCode").Value = 1 Then
		ResponseString = String.Concat(ResponseString, _
			"			    <td align='center' colspan='3'> ", vbCrLf, _
			"    				    <input type='button' value='Close' class='btn' name='Cancel' style='width: 80px' id='Cancel' OnClick='self.close()' />", vbCrLf, _
			"			    </td>", vbCrLf)
	ElseIf cmdValidate.Parameters("errorCode").Value = 2 Then
		ResponseString = String.Concat(ResponseString, _
			"			    <td align='right'>", vbCrLf, _
			"    				    <input type='button' value='Yes' class='btn' name='btnYes' style='width: 80px' id='btnYes' OnClick='createNew()' />", vbCrLf, _
			"			    </td>", vbCrLf, _
			"					<td width='20'></td>", vbCrLf, _
			"			    <td align='left'> ", vbCrLf, _
			"    				    <input type='button' value='No' class='btn' name='btnNo' style='width: 80px' id='btnNo' OnClick='self.close()' />", vbCrLf, _
			"			    </td>", vbCrLf)
	ElseIf cmdValidate.Parameters("errorCode").Value = 3 Then
		ResponseString = String.Concat(ResponseString, _
			"			    <td align='right'> ", vbCrLf, _
			"    				    <input type='button' value='Yes' class='btn' name='btnYes' style='width: 80px' id='btnYes' OnClick='overwrite()' />", vbCrLf, _
			"			    </td>", vbCrLf, _
			"					<td width='20'></td>", vbCrLf, _
			"			    <td align='left'> ", vbCrLf, _
			"    				    <input type='button' value='No' class='btn' name='btnNo' style='width: 80px' id='btnNo' OnClick='self.close()' />", vbCrLf, _
			"			    </td>", vbCrLf)
	ElseIf cmdValidate.Parameters("errorCode").Value = 4 Then
		ResponseString = String.Concat(ResponseString, _
			"			    <td align='right'> ", vbCrLf, _
			"    				    <input type='button' value='Yes' class='btn' name='btnYes' style='width: 80px' id='btnYes' OnClick='continueSave()' />", vbCrLf, _
			"			    </td>", vbCrLf, _
			"					<td width='20'></td>", vbCrLf, _
			"			    <td align='left'> ", vbCrLf, _
			"    				    <input type='button' value ='No' class='btn' name='btnNo' style='width: 80px' id='btnNo' OnClick='self.close() ' />", vbCrLf, _
			"			    </td>", vbCrLf)
	End If
	
	'The following bit is common if an error code was returned
	If cmdValidate.Parameters("errorCode").Value <> 0 Then
		ResponseString = String.Concat(ResponseString, _
			"					<td width='20'></td>", vbCrLf, _
			"			  </tr>", vbCrLf)
	End If
	
	Response.Write(ResponseString)
	
	Response.Write("<script type='text/javascript'>" & vbCrLf)
	Response.Write("	trPleaseWait1.style.display='none';" & vbCrLf)
	Response.Write("	trPleaseWait2.style.display='none';" & vbCrLf)
	Response.Write("	trPleaseWait3.style.display='none';" & vbCrLf)
	Response.Write("	trPleaseWait4.style.display='none';" & vbCrLf)
	Response.Write("	trPleaseWait5.style.display='none';" & vbCrLf)
	Response.Write("</script>" & vbCrLf)
	
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
	util_validate_crosstab_window_onload();
</script>
