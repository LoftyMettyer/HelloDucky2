<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<object
    id="ClientDLL"
    classid="CLSID:3A4EA159-1138-4AC3-B175-966CCB958820"
    codebase="cabs/COAInt_Client.CAB#version=1,0,0,147">
</object>

<%
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
	objMailMerge.Username = session("username")
	objMailMerge.Connection = session("databaseConnection")
	objMailMerge.MailMergeID = session("utilid")
	objMailMerge.ClientDateFormat = session("localedateformat")
	objMailMerge.LocalDecimalSeparator = session("LocaleDecimalSeparator")
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
%>

<script type="text/javascript">
    function util_run_mailmerge_onload() {

        var frmOutput = OpenHR.getForm("workframe", "frmMailMergeOutput");

        debugger;

        <%
    If fok = False Then
%>
        frmOutput.fok.value = "false";
        frmOutput.cancelled.value = "false";
        frmOutput.statusmessage.value = "<%=replace(cleanStringForJavaScript(objMailMerge.ErrorString), vbcrlf, "\n")%>";
        <%
Else
    'Check permission to email addresses
    If objMailMerge.DefOutput = 1 Then
%>

        if (window.document.all.item("txtSysPerm_EMAILADDRESSES_VIEW").value == 0) 
        {
            frmOutput.fok.value = "false";
            frmOutput.cancelled.value = "false";
            frmOutput.statusmessage.value = "You do not have permission to use email addresses.";
        }
        else 
        {
            <%
End If
%>
			
            frmOutput.fok.value = "true";
            frmOutput.cancelled.value = "false";
			
            ClientDLL.MM_DimArrays();

            <%  
    OutputArrayData = objMailMerge.OutputArrayData

    For intCount = 0 To UBound(OutputArrayData)
%>	
            ClientDLL.MM_AddToOutputArrayData("<%=cleanStringForJavaScript(OutputArrayData(intCount))%>");
            <%
Next

MergeFieldsData = objMailMerge.MergeFieldsData

For intCount = 0 To UBound(MergeFieldsData)
%>	
            ClientDLL.MM_AddToMergeFieldsData("<%=cleanStringForJavaScript(MergeFieldsData(intCount))%>");
            <%
Next
%>

            ClientDLL.MM_MergeFieldsUbound = <%=cleanStringForJavaScript(objMailMerge.MergeFieldsUbound)%>;
            ClientDLL.MM_DefName = "<%=cleanStringForJavaScript(objMailMerge.DefName)%>";
            ClientDLL.MM_DefTemplateFile = "<%=cleanStringForJavaScript(objMailMerge.DefTemplateFile)%>";
            ClientDLL.MM_DefPauseBeforeMerge(<%=cleanStringForJavaScript(LCase(CStr(objMailMerge.DefPauseBeforeMerge)))%>);
            ClientDLL.MM_DefSuppressBlankLines(<%=cleanStringForJavaScript(LCase(CStr(objMailMerge.DefSuppressBlankLines)))%>);

            ClientDLL.MM_DefEmailSubject = "<%=cleanStringForJavaScript(objMailMerge.DefEmailSubject)%>";
            ClientDLL.MM_DefEmailAddrCalc = <%=cleanStringForJavaScript(objMailMerge.DefEmailAddrCalc)%>;
            ClientDLL.MM_DefEMailAttachment(<%=cleanStringForJavaScript(LCase(CStr(objMailMerge.DefEMailAttachment)))%>);
            ClientDLL.MM_DefAttachmentName = "<%=cleanStringForJavaScript(objMailMerge.DefAttachmentName)%>";

            ClientDLL.MM_DefOutputFormat = <%=cleanStringForJavaScript(objMailMerge.DefOutputFormat)%>;
            ClientDLL.MM_DefOutputScreen(<%=cleanStringForJavaScript(LCase(objMailMerge.DefOutputScreen))%>);
            ClientDLL.MM_DefOutputPrinter(<%=cleanStringForJavaScript(LCase(objMailMerge.DefOutputPrinter))%>);
            ClientDLL.MM_DefOutputPrinterName = "<%=cleanStringForJavaScript(objMailMerge.DefOutputPrinterName)%>";
            ClientDLL.MM_DefOutputSave(<%=cleanStringForJavaScript(LCase(objMailMerge.DefOutputSave))%>);
            ClientDLL.MM_DefOutputFileName = "<%=cleanStringForJavaScript(objMailMerge.DefOutputFileName)%>";

            ClientDLL.MM_DefDocManMapID = <%=cleanStringForJavaScript(objMailMerge.DefDocManMapID)%>;
            ClientDLL.MM_DefDocManManualHeader(<%=cleanStringForJavaScript(LCase(objMailMerge.DefDocManManualHeader))%>);



            //Execute merge.
            var sOfficeSaveAsValues = '<%=session("OfficeSaveAsValues")%>';
            ClientDLL.SaveAsValues = sOfficeSaveAsValues;
            var blnTest = ClientDLL.MM_WORD_ExecuteMailMerge();
            var blnCancelled = ClientDLL.MM_Cancelled();
            if (blnTest == false)
            {
                if (blnCancelled == true)
                {
                    frmOutput.fok.value = 'true';
                    frmOutput.cancelled.value = 'true';
                }
                else
                {
                    frmOutput.fok.value = 'false';
                    frmOutput.cancelled.value = 'false';
                }
                frmOutput.statusmessage.value = ClientDLL.MM_StatusMessage;
            }
            <%
    If objMailMerge.DefOutput = 2 Then
%>
        }
        <% 
End If
End If
%>
        frmOutput.deftitle.value = "Mail Merge : '<%=cleanStringForJavaScript(objMailMerge.DefName)%>'";
        <%
    If lngEventLogID > 0 Then
%>
        frmOutput.eventlogid.value = "<%=cleanStringForJavaScript(cstr(lngEventLogID))%>";
        <%
Else
%>
        frmOutput.eventlogid.value = 0;
        <%
End If
%>
        frmOutput.failcount.value = <%=cleanStringForJavaScript(cstr(objMailMerge.FailCount))%>;
        frmOutput.successcount.value = <%=cleanStringForJavaScript(cstr(objMailMerge.SuccessCount))%>;
        frmOutput.norecords.value = <%=cleanStringForJavaScript(lcase(objMailMerge.NoRecords))%>;
        frmOutput.nodefinition.value = <%=cleanStringForJavaScript(lcase(blnNoDefinition))%>;
        OpenHR.submitForm(frmOutput);

    }
   
</script>

<%
	fok = objMailMerge.UDFFunctions(false)
	fNotCancelled = Response.IsClientConnected 
	if fok then fok = fNotCancelled
%>

<script type="text/javascript">

    function rtrim(strInput)
    {
        while (strInput.substr(strInput.length-1, 1) == ' ')
        {
            strInput = strInput.substr(0, strInput.length - 1);
        }
        return strInput;
    }

    function Replace(sExpression, sFind, sReplace)
    {
        //gi (global search, ignore case)
        var re = new RegExp(sFind,"gi");
        sExpression = sExpression.replace(re, sReplace);
        return(sExpression);
    }

</script>


<form id="frmOriginalDefinition">
    <%
        Dim sErrMsg As String
        sErrMsg = ""

        Response.Write("	<INPUT type='hidden' id=txtDefn_Name name=txtDefn_Name value=""" & Session("utilname") & """>" & vbCrLf)
        Response.Write("	<INPUT type='hidden' id=txtDefn_ErrMsg name=txtDefn_ErrMsg value=""" & sErrMsg & """>" & vbCrLf)
    %>
    <input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
    <input type="hidden" id="txtDateFormat" name="txtDateFormat" value="<%=session("LocaleDateFormat")%>">

    <input type="hidden" id="txtCancelPrint" name="txtCancelPrint">
    <input type="hidden" id="txtOptionsDone" name="txtOptionsDone">
    <input type="hidden" id="txtOptionsPortrait" name="txtOptionsPortrait">
    <input type="hidden" id="txtOptionsMarginLeft" name="txtOptionsMarginLeft">
    <input type="hidden" id="txtOptionsMarginRight" name="txtOptionsMarginRight">
    <input type="hidden" id="txtOptionsMarginTop" name="txtOptionsMarginTop">
    <input type="hidden" id="txtOptionsMarginBottom" name="txtOptionsMarginBottom">
    <input type="hidden" id="txtOptionsCopies" name="txtOptionsCopies">
</form>

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
    util_run_mailmerge_onload();
</script>
