<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<input type='hidden' id="txtLoadCount" name="txtLoadCount" value="0">

<div id="reportworkframe" data-framesource="util_run_customreports" style="display: inline-block; width:100%; height: 100%;">
		<%
			If Session("utiltype") = "16" Then
				' Bradford Factor
				Html.RenderPartial("~/views/home/util_run_bradford_factor.ascx")
			Else
				Html.RenderPartial("~/views/home/util_run_customreports.ascx")
			End If%>
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
	$(".popup").dialog('option', 'title', $("#txtDefn_Name").val());
	//The next line was overriding the report title	
	if(menu_isSSIMode()) $("#PageDivTitle").html($("#txtDefn_Name").val());
	ShowCustomReport();
</script>