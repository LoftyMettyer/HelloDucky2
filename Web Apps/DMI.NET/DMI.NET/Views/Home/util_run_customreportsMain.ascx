<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<%

    Session("CR_Mode") = ""
    Response.Write("<script type=""text/javascript"">" & vbCrLf)
    Response.Write("function customreport_loadAddRecords()" & vbCrLf)
    Response.Write("{" & vbCrLf)
    Response.Write("  var iCount;" & vbCrLf & vbCrLf)
    
    Response.Write("  debugger;" & vbCrLf & vbCrLf)
    Response.Write("  iCount = new Number(txtLoadCount.value);" & vbCrLf)
    Response.Write("  txtLoadCount.value = iCount + 1" & vbCrLf & vbCrLf)
   
    Response.Write("  if (iCount > 0) {	" & vbCrLf)
    Response.Write("    ShowReport();" & vbCrLf & vbCrLf)   
    Response.Write("  }" & vbCrLf & vbCrLf)

    Response.Write("}" & vbCrLf)
    Response.Write("</script>" & vbCrLf)
  
%>


<input type='hidden' id="txtLoadCount" name="txtLoadCount" value="0">

<div id="reportworkframe" data-framesource="util_run_customreports" style="display: block;">
    <%Html.RenderPartial("~/views/home/util_run_customreports.ascx")%>
</div>

<div id="reportdataframe" data-framesource="util_run_customreportsData" style="display: none;" accesskey="">
    <%Html.RenderPartial("~/views/home/util_run_customreportsData.ascx")%>
</div>

<div id="outputoptions" data-framesource="util_run_outputoptions" style="display: none;">
    <% Html.RenderPartial("~/Views/Home/util_run_outputoptions.ascx")%>
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

