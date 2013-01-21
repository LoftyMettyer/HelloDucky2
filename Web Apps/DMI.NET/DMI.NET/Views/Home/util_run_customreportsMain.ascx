<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

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
    'Response.Write("    var frmData = OpenHR.getForm(""workframe"", ""frmData"");")
    Response.Write("    ShowReport();" & vbCrLf & vbCrLf)
    Response.Write("  }" & vbCrLf & vbCrLf)

    Response.Write("}" & vbCrLf)
    Response.Write("-->" & vbCrLf)
    Response.Write("</script>" & vbCrLf)
  
%>

    
    <INPUT type='hidden' id=txtLoadCount name=txtLoadCount value=0>

<div id="mainframeset">
    <div id="workframe" style="display: none;">
         <%Html.RenderPartial("~/views/home/util_run_customreports.ascx")%> 
    </div>
    <div id="dataframe" style="display: none;">
         <%Html.RenderPartial("~/views/home/util_run_customreportsData.ascx")%>
    </div>
</div>

<FORM id=frmOutput name=frmOutput>
	<INPUT type="hidden" id=fok name=fok value="">
	<INPUT type="hidden" id=cancelled name=cancelled value="">
	<INPUT type="hidden" id=statusmessage name=statusmessage value="">
</FORM>
