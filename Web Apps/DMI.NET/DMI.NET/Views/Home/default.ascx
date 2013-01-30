<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<%
	on error resume next
	
	Dim iOutOfOffice = 0
	Dim iRecordCount = 0

	if Session("action") = "WORKFLOWOUTOFOFFICE_CHECK" then
        Dim cmdWorkflow = CreateObject("ADODB.Command")
        cmdWorkflow.CommandText = "spASRWorkflowOutOfOfficeCheck"
        cmdWorkflow.CommandType = 4 ' Stored procedure
        cmdWorkflow.CommandTimeout = 180
        cmdWorkflow.ActiveConnection = Session("databaseConnection")
					
        Dim prmOutOfOffice = cmdWorkflow.CreateParameter("OutOfOffice", 11, 2) ' 11=boolean, 2=output
        cmdWorkflow.Parameters.Append(prmOutOfOffice)

        Dim prmRecordCount = cmdWorkflow.CreateParameter("RecordCount", 3, 2) ' 3=integer, 2=output
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

        Dim prmValue = cmdOutOfOffice.CreateParameter("value", 11, 1) ' 11=bit, 1=input
        cmdOutOfOffice.Parameters.Append(prmValue)
        prmValue.value = Session("reset")

        Err.Clear()
        cmdOutOfOffice.Execute()
        cmdOutOfOffice = Nothing

        Session("reset") = 0
    End If
%>

<script type="text/javascript">
function default_window_onload() {
	try
		{
		// Do nothing if the menu controls are not yet instantiated.
		if (window.parent.frames("menuframe").document.forms("frmWorkAreaInfo") != null) 
			{
			window.parent.document.all.item("workframeset").cols = "*, 0";	
			// Get menu.asp to refresh the menu.
			window.parent.frames("menuframe").refreshMenu();

			if ("<%=Session("action")%>" == "WORKFLOWOUTOFOFFICE_CHECK")
			{
				if (txtWorkflowRecordCount.value == 0)
				{
					answer = window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox("Unable to set Workflow Out of Office.\nYou do not have an identifiable personnel record.",0); // 0 = OKOnly
				}
				else
				{
					if (txtWorkflowOutOfOffice.value == 1)
					{
						sMsg = "Workflow Out of Office is currently on.\nWould you like to turn it off";
			 		
						if (txtWorkflowRecordCount.value > 1)
						{
							if (txtWorkflowRecordCount.value == 2)
							{
								sMsg = sMsg.concat(" for both");
							}
							else
							{
								sMsg = sMsg.concat(" for all ");
								sMsg = sMsg.concat(txtWorkflowRecordCount.value);
						    }
						      
						    sMsg = sMsg.concat(" of your identified personnel records");
						}
			
						sMsg = sMsg.concat("?");
			 		
			 			answer = window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sMsg,4); // 4 = Yes/No
			 			iResetValue = 0;
			 		}
			 		else
			 		{
						sMsg = "Workflow Out of Office is currently off.\nWould you like to turn it on";
			 		
						if (txtWorkflowRecordCount.value > 1)
						{
							if (txtWorkflowRecordCount.value == 2)
							{
								sMsg = sMsg.concat(" for both");
							}
							else
							{
								sMsg = sMsg.concat(" for all ");
								sMsg = sMsg.concat(txtWorkflowRecordCount.value);
						    }
						      
						    sMsg = sMsg.concat(" of your identified personnel records");
						}
			
						sMsg = sMsg.concat("?");
			 		
			 			answer = window.parent.frames("menuframe").ASRIntranetFunctions.MessageBox(sMsg,4); // 4 = Yes/No
			 			iResetValue = 1;
			 		}
			 	
					if (answer == 6) 
					{
						// Yes
						frmGoto.txtAction.value = "WORKFLOWOUTOFOFFICE_SET";
						frmGoto.txtGotoPage.value = "default.asp";
						frmGoto.txtReset.value = iResetValue;
						frmGoto.submit();
					}
				}
						
				return;	
			}
		}
		else 
		{
			tblMsg.style.visibility="visible";
			tblMsg.style.display="block";
		}
	}
	catch(e) {}
}		
</script>

<INPUT type=hidden id=securitySettingFailure name=securitySettingFailure value=0>	
<INPUT type=hidden id=txtWorkflowOutOfOffice name=txtWorkflowOutOfOffice value=<%=iOutOfOffice%>>
<INPUT type=hidden id=txtWorkflowRecordCount name=txtWorkflowRecordCount value=<%=iRecordCount%>>

<TABLE style="DISPLAY: none; VISIBILITY: hidden" id=tblMsg WIDTH="100%" height="50%" class="invisible" CELLSPACING=0 CELLPADDING=0>
    <tr></tr>
	<TR>
        <TD>
			<TABLE class="outline" CELLSPACING=0 CELLPADDING=0 align=center>
				<TR>
					<TD height=100 width="50%" ALIGN=middle>
						Loading menu. Please wait ...
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>


<script type="text/javascript">
<!--
    function hideMessage() {
        var fso;
        var sMsg;

        tblMsg.style.visibility = "hidden";
        tblMsg.style.display = "none";

    }
    -->
</script>

<FORM action="default_Submit" method=post id=frmGoto name=frmGoto>
	
<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>

</FORM>

<script type="text/javascript">	default_window_onload();</script>