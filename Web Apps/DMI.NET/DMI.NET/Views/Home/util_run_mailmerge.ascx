c<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %><%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<%
	
	Dim fok As Boolean = True
	Dim blnSuccess As Boolean
	Dim bDownloadFile As Boolean
	Dim objMailMerge As MailMerge
	Dim objMailMergeOutput As New Code.MailMergeRun
	Dim fNotCancelled As Boolean
	Dim lngEventLogID As Long
	Dim aPrompts

	' Create the reference to the DLL (Report Class)
	objMailMerge = New HR.Intranet.Server.MailMerge
	objMailMerge.SessionInfo = CType(Session("SessionContext"), SessionInfo)

	' Pass required info to the DLL
	objMailMerge.MailMergeID = CInt(Session("utilid"))
	objMailMerge.ClientDateFormat = Session("localedateformat")
	objMailMerge.SingleRecordID = Session("singleRecordID")
	
	If fok Then
		fok = objMailMerge.SQLGetMergeDefinition
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then
		lngEventLogID = objMailMerge.EventLogAddHeader
		fok = (lngEventLogID > 0)
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	aPrompts = Session("Prompts_" & Session("utiltype") & "_" & Session("utilid"))
	If fok Then
		fok = objMailMerge.SetPromptedValues(aPrompts)
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	objMailMergeOutput.Name = objMailMerge.DefName
	objMailMergeOutput.TemplateName = objMailMerge.DefTemplateFile
	objMailMergeOutput.OutputFileName = objMailMerge.DefOutputFileName
	objMailMergeOutput.EmailSubject = objMailMerge.DefEMailSubject
	objMailMergeOutput.EmailCalculationID = objMailMerge.DefEmailAddrCalc
	objMailMergeOutput.IsAttachment = objMailMerge.DefEMailAttachment
	objMailMergeOutput.AttachmentName = objMailMerge.DefAttachmentName
	objMailMergeOutput.Columns = objMailMerge.Columns
	
	If fok Then
		fok = objMailMergeOutput.ValidateTemplate()
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then
		fok = objMailMergeOutput.ValidateDefinition()
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	
	If fok Then
		fok = objMailMerge.SQLCodeCreate
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then
		fok = objMailMerge.UDFFunctions(True)
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then
		fok = objMailMerge.SQLGetMergeData
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled
	End If

	If fok Then
		fok = objMailMerge.UDFFunctions(False)
		fNotCancelled = Response.IsClientConnected
		If fok Then fok = fNotCancelled

		objMailMergeOutput.MergeData = objMailMerge.MergeData

		If objMailMerge.DefOutputFormat = MailMergeOutputTypes.WordDocument Then
			blnSuccess = objMailMergeOutput.ExecuteMailMerge()
			bDownloadFile = True
		Else
			blnSuccess = objMailMergeOutput.ExecuteToEmail()
			bDownloadFile = False
		End If

		Session("MailMerge_CompletedDocument") = objMailMergeOutput

	End If
	
	%>

<form action="util_run_mailmerge_completed" method="post" id="frmMailMergeOutput" name="frmMailMergeOutput">
	<input type="hidden" id="txtPreview" name="txtPreview" value="false">	
	<input type="hidden" id="deftitle" name="deftitle" value="false">
	<input type="hidden" id="fok" name="fok" value="false">
	<input type="hidden" id="cancelled" name="cancelled" value="false">
	<input type="hidden" id="statusmessage" name="statusmessage" value="">
	<input type="hidden" id="eventlogid" name="eventlogid" value="false">
	<input type="hidden" id="successcount" name="successcount" value="false">
	<input type="hidden" id="failcount" name="failcount" value="false">
	<input type="hidden" id="norecords" name="norecords" value="false">
	<input type="hidden" id="nodefinition" name="nodefinition" value="false">
</form>

<script type="text/javascript">
	
	<%
	Dim sErrorMessage As String
	
	' Errors during the merge
	If objMailMergeOutput.Errors.Count > 0 Then
		objMailMerge.EventLogChangeHeaderStatus(EventLog_Status.elsFailed)

		sErrorMessage = Join(objMailMergeOutput.Errors.ToArray())
		objMailMerge.FailedMessage = sErrorMessage
		sErrorMessage = HttpUtility.JavaScriptStringEncode(sErrorMessage)
		Response.Write(String.Format("raiseWarning(""{0}"", ""{1}"");", objMailMergeOutput.Name, sErrorMessage))
		
	Else
		objMailMerge.EventLogChangeHeaderStatus(EventLog_Status.elsSuccessful)
		
		' No data in result set
		If objMailMerge.NoRecords Then
			sErrorMessage = "Completed successfully, however there were no records that meet the selection criteria. No document has been produced."
			Response.Write(String.Format("raiseWarning(""{0}"", ""{1}"");", objMailMergeOutput.Name, sErrorMessage))
		End If
		
	End If
	%>

	<%	If bDownloadFile And blnSuccess Then%>
	document.getElementById("frmMailMergeOutput").submit();
	<% End If %>
	
	$(".popup").dialog("close");

	if (menu_isSSIMode()) {
		loadPartialView("linksMain", "Home", "workframe", null);
	}

</script>
