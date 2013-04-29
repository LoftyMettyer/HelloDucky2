<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>


<script type="text/javascript">
{

<%
    
    If Session("utiltype") = "" Or _
       Session("utilname") = "" Or _
       Session("utilid") = "" Or _
       Session("action") = "" Then
       
        Response.Write("Error : Not all session variables found...<HR>")
        Response.Write("Type = " & Session("utiltype") & "<BR>")
        Response.Write("UtilName = " & Session("utilname") & "<BR>")
        Response.Write("UtilID = " & Session("utilid") & "<BR>")
        Response.Write("Action = " & Session("action") & "<BR>")
        Response.End()
    End If

    Dim objMailMerge As HR.Intranet.Server.MailMerge
    Dim strOutputMessage As String
    Dim fok As Boolean

    strOutputMessage = ""
    
    objMailMerge = New HR.Intranet.Server.MailMerge()

	objMailMerge.Username = session("username")
	objMailMerge.Connection = session("databaseConnection")
	objMailMerge.EventLogID = Request.Form("eventlogid")

    Response.Write("  sDefTitle = """ & cleanStringForJavaScript(Request.Form("deftitle")) & """;" & vbCrLf)

	objMailMerge.SuccessCount = Request.Form("successcount")
	objMailMerge.FailCount = Request.Form("failcount")
	fok = (Request.Form("fok") = "true")

    If Request.Form("nodefinition") = "false" Then
        If fok = False Then
            objMailMerge.FailCount = objMailMerge.FailCount + objMailMerge.SuccessCount
            objMailMerge.SuccessCount = 0
        End If

        strOutputMessage = CStr(objMailMerge.SuccessCount) & " record"
        If objMailMerge.SuccessCount <> 1 Then
            strOutputMessage = strOutputMessage & "s"
        End If
        strOutputMessage = strOutputMessage & " successful"
        If objMailMerge.FailCount > 0 Then
            strOutputMessage = strOutputMessage & "\n" & CStr(objMailMerge.FailCount) & " record"
            If objMailMerge.FailCount <> 1 Then
                strOutputMessage = strOutputMessage & "s"
            End If
            strOutputMessage = strOutputMessage & " failed"
            fok = False
        End If
    End If

	if Request.Form("norecords") then
		fok = true
	end if
	
	if Request.Form("cancelled") = "true" then
        objMailMerge.EventLogChangeHeaderStatus(1)   'Cancelled
        Response.Write("  OpenHR.messageBox(sDefTitle+"" Cancelled By User."", 48, ""Mail Merge"");" & vbCrLf)

	elseif fok = false then
        If Request.Form("nodefinition") = "false" Then
            objMailMerge.EventLogChangeHeaderStatus(2)   'Failed
            If Request.Form("statusmessage") <> "" Then
                objMailMerge.FailedMessage = Request.Form("statusmessage")
            End If
        End If
        strOutputMessage = CleanStringForJavaScript(Request.Form("statusmessage")) & "\n" & strOutputMessage
        strOutputMessage = " Failed.\n\n" & strOutputMessage
        Response.Write("  OpenHR.messageBox(sDefTitle+ """ & Replace(strOutputMessage, vbCrLf, "\n") & """, 48, ""Mail Merge"");" & vbCrLf)
		
	else
        objMailMerge.EventLogChangeHeaderStatus(3)   'Successful
        If objMailMerge.SuccessCount = 0 And objMailMerge.FailCount = 0 Then
            objMailMerge.FailedMessage = "Completed successfully" & vbCrLf & "No records meet selection criteria"
            strOutputMessage = "No records meet selection criteria."
        End If
        Response.Write("  OpenHR.messageBox(sDefTitle+"" Completed Successfully.\n" & strOutputMessage & """, 64, ""Mail Merge"");" & vbCrLf)
    End If

    
    If Request.Form("nodefinition") = "true" Then
        Response.Write("  try {" & vbCrLf)
        Response.Write("    ToggleCheck();" & vbCrLf)
        Response.Write("  }" & vbCrLf)
        Response.Write("  catch(e) {" & vbCrLf)
        Response.Write("  }" & vbCrLf)
    End If

    %>

    $(".popup").dialog("close");

    }
</script>

