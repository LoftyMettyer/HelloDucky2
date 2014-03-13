<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@Import namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>

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
						<td align="center" colspan="3">Validating Cross Tab
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
							<input type="button" value="Cancel" class="btn" name="Cancel" style="WIDTH: 80px" width="80" id="Cancel" onclick="self.close()" />
						</td>
						<td width="20"></td>
					</tr>


<%

	Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
	
	Dim prmErrorMsg As New SqlParameter("@psErrorMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
	Dim prmErrorCode As New SqlParameter("@piErrorCode", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
	Dim prmDeletedFilters As New SqlParameter("@psDeletedFilters", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
	Dim prmHiddenFilters As New SqlParameter("@psHiddenFilters", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
	Dim prmJobIDsToHide As New SqlParameter("@psJobIDsToHide", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
	      
	objDataAccess.ExecuteSP("sp_ASRIntValidateCrossTab", _
					New SqlParameter("psUtilName", SqlDbType.VarChar, 255) With {.Value = Request("validateName")}, _
					New SqlParameter("piUtilID", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateUtilID"))}, _
					New SqlParameter("piTimestamp", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateTimestamp"))}, _
					New SqlParameter("piBasePicklistID", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateBasePicklist"))}, _
					New SqlParameter("piBaseFilterID", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateBaseFilter"))}, _
					New SqlParameter("piEmailGroupID", SqlDbType.Int) With {.Value = CleanNumeric(Request("validateEmailGroup"))}, _
					New SqlParameter("psHiddenGroups ", SqlDbType.VarChar, -1) With {.Value = Request("validateHiddenGroups")}, _
					prmErrorMsg, prmErrorCode, prmDeletedFilters, prmHiddenFilters, prmJobIDsToHide)
	
	Dim ResponseString As String = String.Concat( _
		"<input type='hidden' id='txtErrorCode' 'name='txtErrorCode' value='", prmErrorCode.Value, "'>", vbCrLf, _
		"<input type='hidden' id='txtDeletedFilters' name='txtDeletedFilters' value='", prmDeletedFilters.Value, "'>", vbCrLf, _
		"<input type='hidden' id='txtHiddenFilters' name='txtHiddenFilters' value='", prmHiddenFilters.Value, "'>", vbCrLf, _
		"<input type='hidden' id='txtJobIDsToHide' name='txtJobIDsToHide' value='", prmJobIDsToHide.Value, "'>", vbCrLf)
	
	'The following bit is common if an error code was returned
	If prmErrorCode.Value <> 0 Then
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
			"						" & prmErrorMsg.Value, vbCrLf, _
			"			    </td>", vbCrLf, _
			"					<td width='20'></td>", vbCrLf, _
			"			  </tr>", vbCrLf, _
			"			  <tr>", vbCrLf, _
			"					<td height='20' colspan='5'></td>", vbCrLf, _
			"			  </tr>", vbCrLf, _
			"			  <tr> ", vbCrLf, _
			"					<td width='20'></td>", vbCrLf)
	End If
	
	If prmErrorCode.Value = 1 Then
		ResponseString = String.Concat(ResponseString, _
			"			    <td align='center' colspan='3'> ", vbCrLf, _
			"    				    <input type='button' value='Close' class='btn' name='Cancel' style='width: 80px' id='Cancel' OnClick='self.close()' />", vbCrLf, _
			"			    </td>", vbCrLf)
	ElseIf prmErrorCode.Value = 2 Then
		ResponseString = String.Concat(ResponseString, _
			"			    <td align='right'>", vbCrLf, _
			"    				    <input type='button' value='Yes' class='btn' name='btnYes' style='width: 80px' id='btnYes' OnClick='createNew()' />", vbCrLf, _
			"			    </td>", vbCrLf, _
			"					<td width='20'></td>", vbCrLf, _
			"			    <td align='left'> ", vbCrLf, _
			"    				    <input type='button' value='No' class='btn' name='btnNo' style='width: 80px' id='btnNo' OnClick='self.close()' />", vbCrLf, _
			"			    </td>", vbCrLf)
	ElseIf prmErrorCode.Value = 3 Then
		ResponseString = String.Concat(ResponseString, _
			"			    <td align='right'> ", vbCrLf, _
			"    				    <input type='button' value='Yes' class='btn' name='btnYes' style='width: 80px' id='btnYes' OnClick='overwrite()' />", vbCrLf, _
			"			    </td>", vbCrLf, _
			"					<td width='20'></td>", vbCrLf, _
			"			    <td align='left'> ", vbCrLf, _
			"    				    <input type='button' value='No' class='btn' name='btnNo' style='width: 80px' id='btnNo' OnClick='self.close()' />", vbCrLf, _
			"			    </td>", vbCrLf)
	ElseIf prmErrorCode.Value = 4 Then
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
	If prmErrorCode.Value <> 0 Then
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
