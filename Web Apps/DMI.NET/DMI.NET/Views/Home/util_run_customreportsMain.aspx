<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

<!DOCTYPE html>

<html>
<head runat="server">
    <title>util_run_customreportsMain</title>
</head>
<body>
    
    <%

        Session("CR_Mode") = ""
        Response.Write("<script type=""text/javascript"">" & vbCrLf)
        Response.Write("<!--" & vbCrLf)
        Response.Write("function loadAddRecords()" & vbCrLf)
        Response.Write("{" & vbCrLf)
        Response.Write("  var iCount;" & vbCrLf & vbCrLf)

        Response.Write("  iCount = new Number(txtLoadCount.value);" & vbCrLf)
        Response.Write("  txtLoadCount.value = iCount + 1" & vbCrLf & vbCrLf)

        Response.Write("  if (iCount > 0) {	" & vbCrLf)
        Response.Write("    frmData = window.frames(""workframe"");" & vbCrLf)
        Response.Write("    frmData.ShowReport();" & vbCrLf & vbCrLf)
        Response.Write("  }" & vbCrLf & vbCrLf)

        Response.Write("}" & vbCrLf)
        Response.Write("-->" & vbCrLf)
        Response.Write("</script>" & vbCrLf)

    %>

</head>
    <INPUT type='hidden' id=txtLoadCount name=txtLoadCount value=0>

<frameset name=mainframeset rows="*,0" frameborder="0" framespacing="0">
  <frame src="util_run_customreports" name="workframe" noresize> 
  <frame src="util_run_customreportsData" name="dataframe" noresize> 
</frameset>


<FORM id=frmOutput name=frmOutput>
	<INPUT type="hidden" id=fok name=fok value="">
	<INPUT type="hidden" id=cancelled name=cancelled value="">
	<INPUT type="hidden" id=statusmessage name=statusmessage value="">
</FORM>


</body>
</html>
