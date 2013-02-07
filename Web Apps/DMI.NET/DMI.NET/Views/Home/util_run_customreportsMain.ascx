<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<%

    Session("CR_Mode") = ""
    Response.Write("<script type=""text/javascript"">" & vbCrLf)
    Response.Write("function loadAddRecords()" & vbCrLf)
    Response.Write("{" & vbCrLf)
    Response.Write("  var iCount;" & vbCrLf & vbCrLf)

    Response.Write("  iCount = new Number(txtLoadCount.value);" & vbCrLf)
    Response.Write("  txtLoadCount.value = iCount + 1" & vbCrLf & vbCrLf)

    Response.Write("  if (iCount > 0) {	" & vbCrLf)
    Response.Write("    ShowReport();" & vbCrLf & vbCrLf)
    Response.Write("  }" & vbCrLf & vbCrLf)

    Response.Write("}" & vbCrLf)
    Response.Write("</script>" & vbCrLf)
  
%>


<input type='hidden' id="txtLoadCount" name="txtLoadCount" value="0">

<div id="customreportmainframeset">
    <div id="reportworkframe" style="display: none;">
        <%Html.RenderPartial("~/views/home/util_run_customreports.ascx")%>
    </div>

    <div id="reportdataframe" style="display: none;">
        <%Html.RenderPartial("~/views/home/util_run_customreportsData.ascx")%>
    </div>
</div>

<form id="frmOutput" name="frmOutput">
    <input type="hidden" id="fok" name="fok" value="">
    <input type="hidden" id="cancelled" name="cancelled" value="">
    <input type="hidden" id="statusmessage" name="statusmessage" value="">
</form>


<script type="text/javascript">
    reports_window_onload();
    util_run_customreports_addActiveXHandlers();
</script>

