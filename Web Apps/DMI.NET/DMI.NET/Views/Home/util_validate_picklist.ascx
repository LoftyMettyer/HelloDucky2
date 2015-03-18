<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<script type="text/javascript">

	function util_validate_picklist_window_onload() {

		$("#reportframe").attr("data-framesource", "UTIL_VALIDATE_PICKLIST");

		if ($('#divUtil_validate_picklist #txtDisplay').val() != "False") {
			// Hide the 'Loading' message.
			$('#divUtil_validate_picklist h3[id^="trPleaseWait"]').hide();

			$('.popup').dialog('option', 'width', screen.width / 3);
			$('.popup').dialog('option', 'height', 'auto');
		}
		else {			
			nextPass();
		}
	}

	function nextPass() {
		var frmValidate = OpenHR.getForm("reportframe", "frmValidatePicklist");

		var iNextPass = new Number(frmValidate.validatePass.value);
		iNextPass = iNextPass + 1;

		if (iNextPass == 2) {
			frmValidate.validatePass.value = iNextPass;
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

	function makeHidden() {
		nextPass();
	}

</script>

<div id="divUtil_validate_picklist">
	<h3 id="trPleaseWait1">Validating Picklist</h3>
	<h3 id="trPleaseWait4"></h3>
	<h3 id="trPleaseWait2">Please wait...</h3>
	<h3 id="trPleaseWait5"></h3>
	<h3 id="trPleaseWait3"><input type="button" value="Cancel" class="btn" name="Cancel" style="width: 80px; float: right;" id="Cancel" /></h3>

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
				Response.Write("<h3>Error Saving Picklist</h3>" & vbCrLf)
				Response.Write(prmErrorMsg.Value & vbCrLf)
				Response.Write("<br/><br/>" & vbCrLf)
				Response.Write("<input type='button' value='No' class='btn' name='btnNo' style='width: 80px; float: right;' id='btnNo' onclick='closeclick();' />" & vbCrLf)
				Response.Write("<input type='button' value='Yes' class='btn' name='btnYes' style='width: 80px; float: right; margin-right: 10px;' id='btnYes' onclick='udp_createNew();' />" & vbCrLf)
			Else
				If prmErrorCode.Value = 2 Then
					fDisplay = True
					Response.Write("<h3>Error Saving Picklist</h3>" & vbCrLf)
					Response.Write(prmErrorMsg.Value & vbCrLf)
					Response.Write("<br/><br/>" & vbCrLf)
					Response.Write("<input type='button' value='No' class='btn' name='btnNo' style='width: 80px float: right;;' id='Button2' onclick='closeclick();'/>" & vbCrLf)
					Response.Write("<input type='button' value='Yes' class='btn' name='btnYes' style='width: 80px; float: right; margin-right: 10px;' id='Button1' onclick='overwrite();' />" & vbCrLf)
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
					Response.Write("<h3>Error Saving Picklist</h3>" & vbCrLf)
					Response.Write(prmErrorMsg.Value & vbCrLf)
					Response.Write("<br/><br/>" & vbCrLf)
					Response.Write("<input type='button' value='Close' class='btn' name='Cancel' style='width: 80px; float: right;' id='Button3' onclick='closeclick();'/>" & vbCrLf)
				Else
					If prmErrorCode.Value = 2 Then
						fDisplay = True
						Response.Write("<h3>Error Saving Picklist</h3>" & vbCrLf)
						Response.Write(prmErrorMsg.Value & vbCrLf)
						Response.Write("<br/><br/>" & vbCrLf)
						Response.Write("<input type='button' value='No' class='btn' name='btnNo' style='WIDTH: 80px; float: right;' id='Button5' onclick='closeclick();' />" & vbCrLf)
						Response.Write("<input type='button' value='Yes' class='btn' name='btnYes' style='width: 80px; float: right; margin-right: 10px;' id='Button4' onclick='makeHidden();' />" & vbCrLf)
					End If
				End If
	
			End If
			
		End If
		
	
		Response.Write(String.Format("<input type='hidden' id='txtDisplay' name='txtDisplay' value='{0}'>", fDisplay) & vbCrLf)
	%>
</div>

<form id="frmValidatePicklist" name="frmValidatePicklist" method="post" style="display: none;">
		<input type="hidden" id="validatePass" name="validatePass" value='<%:Request("validatePass")%>'>
		<input type="hidden" id="validateName" name="validateName" value="<%:replace(Request("validateName"), """", "&quot;")%>">
		<input type="hidden" id="validateTimestamp" name="validateTimestamp" value='<%:Request("validateTimestamp")%>'>
		<input type="hidden" id="validateUtilID" name="validateUtilID" value='<%:Request("validateUtilID")%>'>
		<input type="hidden" id="validateAccess" name="validateAccess" value='<%:Request("validateAccess")%>'>
		<input type="hidden" id="validateBaseTableID" name="validateBaseTableID" value='<%:Request("validateBaseTableID")%>'>
		<input type="hidden" id="test" name="test" value="<%:Request.QueryString%>">
</form>

<script type="text/javascript">
		util_validate_picklist_window_onload();
</script>
