c<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<%
	
	Dim ActiveXClientDLL As New HR.Intranet.Server.MailMergeClient
	
		Dim MergeFieldsData
		Dim OutputArrayData
		Dim fok As Boolean
		Dim objMailMerge As HR.Intranet.Server.MailMerge
		Dim fNotCancelled As Boolean
		Dim lngEventLogID As Long
		Dim blnNoDefinition As Boolean
		Dim aPrompts

	if session("utiltype") = "" or _ 
		 session("utilname") = "" or _ 
		 session("utilid") = "" or _ 
		 session("action") = "" then 
				
				Response.Write("Error : Not all session variables found...<HR>")
				Response.Write("Type = " & Session("utiltype") & "<BR>")
				Response.Write("UtilName = " & Session("utilname") & "<BR>")
				Response.Write("UtilID = " & Session("utilid") & "<BR>")
				Response.Write("Action = " & Session("action") & "<BR>")
				Response.End()
	end if

	' Create the reference to the DLL (Report Class)
		objMailMerge = New HR.Intranet.Server.MailMerge   
		
	' Pass required info to the DLL
	objMailMerge.Username = Session("username").ToString()
	objMailMerge.Connection = session("databaseConnection")
	objMailMerge.MailMergeID = session("utilid")
	objMailMerge.ClientDateFormat = session("localedateformat")
	objMailMerge.SingleRecordID = Session("singleRecordID")
	
	fok = true
	blnNoDefinition = true

	if fok then
		fok = objMailMerge.SQLGetMergeDefinition
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok then
		blnNoDefinition = false

		lngEventLogID = objMailMerge.EventLogAddHeader
		fok = (lngEventLogID > 0)
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	aPrompts =  Session("Prompts_" & session("utiltype") & "_" & session("utilid"))
	if fok then 
		fok = objMailMerge.SetPromptedValues(aPrompts)
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok then 
		fok = objMailMerge.SQLGetMergeColumns
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok then 
		fok = objMailMerge.SQLCodeCreate
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok then
		fok = objMailMerge.UDFFunctions(true)
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok then 
		fok = objMailMerge.SQLGetMergeData
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	if fok then 
		fok = objMailMerge.BuildOutputArray
		fNotCancelled = Response.IsClientConnected 
		if fok then fok = fNotCancelled
	end if

	ActiveXClientDLL.MM_DimArrays()

	OutputArrayData = objMailMerge.OutputArrayData

	For intCount = 0 To UBound(OutputArrayData)
		ActiveXClientDLL.MM_AddToOutputArrayData(OutputArrayData(intCount))
	Next
		
	MergeFieldsData = objMailMerge.MergeFieldsData

	For intCount = 0 To UBound(MergeFieldsData)
		ActiveXClientDLL.MM_AddToMergeFieldsData(MergeFieldsData(intCount))
	Next

	ActiveXClientDLL.MM_MergeFieldsUbound = objMailMerge.MergeFieldsUBound
	ActiveXClientDLL.MM_DefName = objMailMerge.DefName
	ActiveXClientDLL.MM_DefTemplateFile = objMailMerge.DefTemplateFile
	ActiveXClientDLL.MM_DefPauseBeforeMerge(objMailMerge.DefPauseBeforeMerge)
	ActiveXClientDLL.MM_DefSuppressBlankLines(objMailMerge.DefSuppressBlankLines)

	ActiveXClientDLL.MM_DefEmailSubject = objMailMerge.DefEMailSubject
	ActiveXClientDLL.MM_DefEmailAddrCalc = objMailMerge.DefEmailAddrCalc
	ActiveXClientDLL.MM_DefEMailAttachment(objMailMerge.DefEMailAttachment)
	ActiveXClientDLL.MM_DefAttachmentName = objMailMerge.DefAttachmentName

	ActiveXClientDLL.MM_DefOutputFormat = objMailMerge.DefOutputFormat
	ActiveXClientDLL.MM_DefOutputScreen(objMailMerge.DefOutputScreen)
	ActiveXClientDLL.MM_DefOutputPrinter(objMailMerge.DefOutputPrinter)
	ActiveXClientDLL.MM_DefOutputPrinterName = objMailMerge.DefOutputPrinterName
	ActiveXClientDLL.MM_DefOutputSave(objMailMerge.DefOutputSave)
	ActiveXClientDLL.MM_DefOutputFileName = objMailMerge.DefOutputFileName

	ActiveXClientDLL.SaveAsValues = Session("OfficeSaveAsValues")

	Dim blnSuccess = ActiveXClientDLL.MM_WORD_ExecuteMailMerge()
	Dim blnCancelled = ActiveXClientDLL.MM_Cancelled()
	
	Session("MailMerge_CompletedDocument") = ActiveXClientDLL
	
	fok = objMailMerge.UDFFunctions(false)
	fNotCancelled = Response.IsClientConnected 
	If fok Then fok = fNotCancelled
	
	If Not blnSuccess Then
		Response.Write(String.Format("Failure : {0}", ActiveXClientDLL.ErrorMessage))	
	End If
	
%>

<form action="util_run_mailmerge_completed" method="post" id="frmMailMergeOutput" name="frmMailMergeOutput">
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
	<%	If objMailMerge.DefOutputScreen And blnSuccess Then%>
		document.getElementById("frmMailMergeOutput").submit();
	<% End If %>
</script>
