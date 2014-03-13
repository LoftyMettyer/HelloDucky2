<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<%
	on error resume next
	
	Dim iOutOfOffice = 0
	Dim iRecordCount = 0

	Dim objSession As SessionInfo = CType(Session("SessionContext"), SessionInfo)
	Dim objDataAccess As New clsDataAccess(objSession.LoginInfo)

	If Session("action") = "WORKFLOWOUTOFOFFICE_CHECK" Then
		
		Dim prmOutOfOffice As SqlParameter = New SqlParameter("pfOutOfOffice", SqlDbType.Bit)
		prmOutOfOffice.Direction = ParameterDirection.Output

		Dim prmRecordCount As SqlParameter = New SqlParameter("piRecordCount", SqlDbType.Int)
		prmRecordCount.Direction = ParameterDirection.Output

		objDataAccess.ExecuteSP("spASRWorkflowOutOfOfficeCheck", prmOutOfOffice, prmRecordCount)

		iOutOfOffice = CInt(prmOutOfOffice.Value)
		iRecordCount = CInt(prmRecordCount.Value)		
		
	End If
	
	If Session("action") = "WORKFLOWOUTOFOFFICE_SET" Then
		
		Dim prmSetOffice As SqlParameter = New SqlParameter("pfOutOfOffice", SqlDbType.Bit)
		prmSetOffice.Value = CBool(Session("reset"))
		objDataAccess.ExecuteSP("spASRWorkflowOutOfOfficeSet", prmSetOffice)
		
		Session("reset") = 0
		
		' JIRA 3286 - reset the session variable, as we revisit this page as part of main.ascx refresh.
		Session("action") = ""
		
	End If
%>

<script type="text/javascript">
	function default_window_onload() {		
		try {			
		// Do nothing if the menu controls are not yet instantiated.
			if (OpenHR.getForm("menuframe", "frmMenuInfo") != null)
			{
			//window.parent.document.all.item("workframeset").cols = "*, 0";	
			$("#workframe").attr("data-framesource", "DEFAULT");
			
			if ("<%=Session("action")%>" == "WORKFLOWOUTOFOFFICE_CHECK")
			{
				if ($('#txtWorkflowRecordCount').val() == 0)
				{
					var answer = OpenHR.messageBox("Unable to set Workflow Out of Office.\nYou do not have an identifiable personnel record.",0); // 0 = OKOnly
				}
				else
				{
					if ($('#txtWorkflowOutOfOffice').val() == 1)
					{
						var sMsg = "Workflow Out of Office is currently on.\nWould you like to turn it off";
					
						if ($('#txtWorkflowRecordCount').val() > 1)
						{
							if ($('#txtWorkflowRecordCount').val() == 2)
							{
								sMsg = sMsg.concat(" for both");
							}
							else
							{
								sMsg = sMsg.concat(" for all ");
								sMsg = sMsg.concat($('#txtWorkflowRecordCount').val());
								}
									
								sMsg = sMsg.concat(" of your identified personnel records");
						}
			
						sMsg = sMsg.concat("?");
						answer =OpenHR.messageBox(sMsg,36); // 4 = Yes/No
						var iResetValue = 0;
					}
					else {
						sMsg = "Workflow Out of Office is currently off.\nWould you like to turn it on";
					
						if ($('#txtWorkflowRecordCount').val() > 1)
						{
							if ($('#txtWorkflowRecordCount').val() == 2)
							{
								sMsg = sMsg.concat(" for both");
							}
							else
							{
								sMsg = sMsg.concat(" for all ");
								sMsg = sMsg.concat($('#txtWorkflowRecordCount').val());
								}
									
								sMsg = sMsg.concat(" of your identified personnel records");
						}
			
						sMsg = sMsg.concat("?");
					
						answer = OpenHR.messageBox(sMsg,36); // 4 = Yes/No
						iResetValue = 1;
					}
				
					if (answer == 6) 
					{
						// Yes
						var frmGoto = document.getElementById('frmGoto');
						frmGoto.txtAction.value = "WORKFLOWOUTOFOFFICE_SET";
						frmGoto.txtGotoPage.value = "_default";
						frmGoto.txtReset.value = iResetValue;
						OpenHR.submitForm(frmGoto);
					}
				}
						
				return false;	
			}
		}
		else {
			$('#tblMsg').show();			
		}
	}
	catch(e) {}
}		
</script>

<input type="hidden" id="securitySettingFailure" name="securitySettingFailure" value="0">
<input type="hidden" id="txtWorkflowOutOfOffice" name="txtWorkflowOutOfOffice" value="<%=iOutOfOffice%>">
<input type="hidden" id="txtWorkflowRecordCount" name="txtWorkflowRecordCount" value="<%=iRecordCount%>">
<input type="hidden" id="txtWf_OutOfOffice" name="txtWf_OutOfOffice" value="<%=Session("WF_OutOfOffice")%>"/>
<div <%=session("BodyTag")%>>
	<table style="display: none;" id="tblMsg" width="100%" height="50%" class="invisible" cellspacing="0" cellpadding="0">
		<tr></tr>
		<tr>
			<td>
				<table class="outline" cellspacing="0" cellpadding="0" align="center">
					<tr>
						<td height="100" width="50%" align="middle">Loading menu...
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</div>

<script type="text/javascript">
		function hideMessage() {
				$('#tblMsg').hide();
		}
</script>

<form action="default_Submit" method="post" id="frmGoto" name="frmGoto">
	<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</form>

<script type="text/javascript">	default_window_onload();</script>