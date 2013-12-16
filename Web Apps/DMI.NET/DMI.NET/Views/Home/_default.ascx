<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<%
	on error resume next
	
	Dim iOutOfOffice = 0
	Dim iRecordCount = 0

	if Session("action") = "WORKFLOWOUTOFOFFICE_CHECK" then
		Dim cmdWorkflow = CreateObject("ADODB.Command")
		cmdWorkflow.CommandText = "spASRWorkflowOutOfOfficeCheck"
		cmdWorkflow.CommandType = 4	' Stored procedure
		cmdWorkflow.CommandTimeout = 180
		cmdWorkflow.ActiveConnection = Session("databaseConnection")
					
		Dim prmOutOfOffice = cmdWorkflow.CreateParameter("OutOfOffice", 11, 2) ' 11=boolean, 2=output
		cmdWorkflow.Parameters.Append(prmOutOfOffice)

		Dim prmRecordCount = cmdWorkflow.CreateParameter("RecordCount", 3, 2)	' 3=integer, 2=output
		cmdWorkflow.Parameters.Append(prmRecordCount)

		Err.Clear()
		cmdWorkflow.Execute()

		If cmdWorkflow.Parameters("OutOfOffice").Value Then
			iOutOfOffice = 1
		End If
		iRecordCount = CInt(cmdWorkflow.Parameters("RecordCount").Value)
			
		cmdWorkflow = Nothing
	End If
	
	If Session("action") = "WORKFLOWOUTOFOFFICE_SET" Then
		Dim cmdOutOfOffice = CreateObject("ADODB.Command")
		cmdOutOfOffice.CommandText = "spASRWorkflowOutOfOfficeSet"
		cmdOutOfOffice.CommandType = 4 ' Stored Procedure
		cmdOutOfOffice.ActiveConnection = Session("databaseConnection")

		Dim prmValue = cmdOutOfOffice.CreateParameter("value", 11, 1)	' 11=bit, 1=input
		cmdOutOfOffice.Parameters.Append(prmValue)
		prmValue.value = Session("reset")

		Err.Clear()
		cmdOutOfOffice.Execute()
				
		cmdOutOfOffice = Nothing

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
			
			// Get menu.asp to refresh the menu.
			//window.parent.frames("menuframe").refreshMenu();			
			menu_refreshMenu();

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
						<td height="100" width="50%" align="middle">Loading menu. Please wait ...
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