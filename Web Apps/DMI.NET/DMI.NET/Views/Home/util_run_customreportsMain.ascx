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
