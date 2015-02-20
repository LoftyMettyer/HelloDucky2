<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<script src="<%:Url.Content("~/Scripts/jquery/jquery.cookie.js")%>"></script>

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
	objMailMergeOutput.PrinterName = objMailMerge.DefOutputPrinterName
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

		Select Case objMailMerge.DefOutputFormat
			Case MailMergeOutputTypes.WordDocument
				blnSuccess = objMailMergeOutput.ExecuteMailMerge(False)
				bDownloadFile = True
			
			Case MailMergeOutputTypes.IndividualEmail
				blnSuccess = objMailMergeOutput.ExecuteToEmail()
				bDownloadFile = False

			Case Else
				blnSuccess = objMailMergeOutput.ExecuteMailMerge(True)
				bDownloadFile = False			
				
		End Select
		
		
		Session("MailMerge_CompletedDocument") = objMailMergeOutput

	End If
	
	%>

<form action="util_run_mailmerge_completed" method="post" id="frmMailMergeOutput" name="frmMailMergeOutput">
	<input type="hidden" id="txtPreview" name="txtPreview" value="True">	
	<input type="hidden" id="deftitle" name="deftitle" value="false">
	<input type="hidden" id="fok" name="fok" value="false">
	<input type="hidden" id="cancelled" name="cancelled" value="false">
	<input type="hidden" id="statusmessage" name="statusmessage" value="">
	<input type="hidden" id="eventlogid" name="eventlogid" value="false">
	<input type="hidden" id="successcount" name="successcount" value="false">
	<input type="hidden" id="failcount" name="failcount" value="false">
	<input type="hidden" id="norecords" name="norecords" value="false">
	<input type="hidden" id="nodefinition" name="nodefinition" value="false">
	<input type="hidden" id="download_token_value_id" name="download_token_value_id"/>
	<%=Html.AntiForgeryToken()%>
</form>

<script type="text/javascript">
	
	<%
	Dim sErrorMessage As String
	
	' Errors during the merge
	If Len(objMailMerge.ErrorString) > 0 Then
		sErrorMessage = HttpUtility.JavaScriptStringEncode(objMailMerge.ErrorString)
		Response.Write(String.Format("OpenHR.modalPrompt(""{0}"",2,""{1}"");", sErrorMessage, objMailMerge.DefName))
	
	ElseIf objMailMergeOutput.Errors.Count > 0 Then
		objMailMerge.EventLogChangeHeaderStatus(EventLog_Status.elsFailed)

		sErrorMessage = Join(objMailMergeOutput.Errors.ToArray())
		objMailMerge.FailedMessage = sErrorMessage
		sErrorMessage = HttpUtility.JavaScriptStringEncode(sErrorMessage)
		Response.Write(String.Format("OpenHR.modalPrompt(""{0}"",2,""{1}"");", sErrorMessage, objMailMerge.DefName))
		Session("mailmergefail") = True
		
	Else
		objMailMerge.EventLogChangeHeaderStatus(EventLog_Status.elsSuccessful)
		
		' No data in result set
		If objMailMerge.NoRecords Then
			sErrorMessage = "Completed successfully, however there were no records that meet the selection criteria. No document has been produced."
			Response.Write(String.Format("OpenHR.modalPrompt(""{0}"",2,""{1}"");", sErrorMessage, objMailMerge.DefName))
		Else
			sErrorMessage = "Mail merge completed successfully."
			Response.Write(String.Format("OpenHR.modalPrompt(""{0}"",2,""{1}"");", sErrorMessage, objMailMerge.DefName))
		End If
        
		
	End If
	%>

	<%	If bDownloadFile And blnSuccess Then%>

	var fileDownloadCheckTimer;
	function blockUIForDownload() {
		
		var token = new Date().getTime(); //use the current timestamp as the token value		
		$('#download_token_value_id').val(token);
		menu_ShowWait('Generating output...');
		setTimeout('updateProgressMsg()', 50);
		$("body").addClass("loading");
		fileDownloadCheckTimer = window.setInterval(function () {
			var cookieValue = $.cookie('fileDownloadToken');
			if (cookieValue == token) {
				finishDownload();
			} else {
				$('#txtProgressMessage').val('Generating output...');
				$("body").addClass("loading");  //Overlapping ajax calls may have closed the spinner.
				updateProgressMsg();
			}
		}, 1000);
	}

	function finishDownload() {
		window.clearInterval(fileDownloadCheckTimer);
		$.removeCookie('fileDownloadToken'); //clears this cookie value		
		$("body").removeClass("loading");
		menu_ShowWait('Loading...');		
	}

	var frmMailMergeOutput = document.getElementById("frmMailMergeOutput");
	$(frmMailMergeOutput).submit(function () {
		blockUIForDownload();
	});

	$(frmMailMergeOutput).submit();
	<% End If %>
	
	$(".popup").dialog('option', 'title', "");
	$(".popup").dialog("close");

	if (menu_isSSIMode()) {
		loadPartialView("linksMain", "Home", "workframe", null);
	}
</script>
