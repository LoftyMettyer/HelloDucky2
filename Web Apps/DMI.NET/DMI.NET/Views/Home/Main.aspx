﻿<%@ Page Title="" Language="VB" MasterPageFile="~/Views/Shared/Site.Master" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">	
<%=DMI.NET.svrCleanup.GetPageTitle("") %>
</asp:Content>


<asp:Content runat="server" ID="Content1a" ContentPlaceHolderID="FixedLinksContent">
	<div id="fixedlinksframe" style="display: none;"><%	Html.RenderPartial("~/views/home/fixedlinks.ascx")%></div>	
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

<script type="text/javascript">
	
	function handleAjaxError(html) {

		//handle error
		OpenHR.messageBox(html.ErrorMessage.replace("<p>", "\n\n"), 48, html.ErrorTitle);

		window.location.href = "<%=Url.Action("Login", "Account")%>";
	}

	

	$(function () {

		<% 
	Response.Write("window.LocaleDateFormat = """ & Session("LocaleDateFormat") & """;")
		%>

			var SelfServiceUserType = '<%=ViewBag.SSIMode%>';

			if (SelfServiceUserType == 'True') {
				$("#workframeset").css("left", "0px");
				$("#reportframeset").css("left", "0px");
			}
			else {
				// ----  Apply jQuery functionality to the slide out CONTEXT MENU  ----
				var contextMenuTab = {
					speed: 300,
					//containerWidth: $('#menuframe').outerWidth() - 30,
					containerHeight: $('.ContextMenu-panel').outerHeight(),
					tabWidth: $('.ContextMenu-tab').outerWidth(),
					init: function () {
						$('.ContextMenu-tab').click(function (event) {
							var containerWidth = $('.ContextMenu-panel').outerWidth() - 30;

							if ($('#menuframe').hasClass('open')) {
								$('#menuframe').animate({ left: '-' + containerWidth }, contextMenuTab.speed)
									.removeClass('open');
								$("#workframeset").css("left", "30px");
								// $("#reportframeset").css("left", "30px");
								$('#ContextMenuIcon').attr('src', '<%= Url.Content("~/content/images/expand.png") %>');

							} else {
								$('#menuframe').animate({ left: '0' }, contextMenuTab.speed)
									.addClass('open');
								$("#workframeset").css("left", containerWidth);
								// $("#reportframeset").css("left", "350px");
								$('#ContextMenuIcon').attr('src', '<%= Url.Content("~/content/images/retract.png") %>');
							}
							
							//resize defsel/find screen accordingly.
							$('#findGridTable').setGridWidth($('#findGridRow').width());
							$('#DefSelRecords').setGridWidth($('#findGridRow').width());

							event.preventDefault();
						});
						
					}
				};
				contextMenuTab.init();
			}
		});

	$(document).ready(function() {

		$("#fixedlinksframe").show();
		$("#FixedLinksContent").fadeIn("slow");

		$(".popup").dialog({
			overflow: false,
			autoOpen: false,
			modal: true,
			height: 550,
			width: 800
		});
		
		$('#divPopupReportDefinition').dialog({
			overflow: false,
			autoOpen: false,
			width: '400px',
			height: 'auto',
			resizable: false,
			modal: true,
			title: 'OpenHR Web'
		});

		$('#divExpressionSelection').dialog({
			overflow: false,
			autoOpen: false,
			width: 'auto',
			height: 'auto',
			resizable: true,
			modal: true,
			title: 'OpenHR Web'
		});

		//load menu for dmi, or linksmain for ssi
		var SelfServiceUserType = '<%=ViewBag.SSIMode%>';

		if (SelfServiceUserType == 'True') {

			$.ajax({
				url: 'linksMain',
				dataType: 'html',
				type: 'POST',
				data: { psScreenInfo: '<%=session("SingleRecordTableID")%>!<%=session("SingleRecordViewID")%>_0' },
				success: function(html) {
					try {
						var jsonResponse = $.parseJSON(html);
						if (jsonResponse.ErrorMessage.length > 0) {
							handleAjaxError(jsonResponse);
							return false;
						}
					} catch(e) {
					}


					//$("#workframe").hide();
					$("#workframe").html(html).show();


					//final resize of the dashboard - for tiles, ensure width is sufficient
					resizeDashboard(); //site.master function.

				},
				error: function(req, status, errorObj) {

				}
			});
		} else {
			$("#menuframe").fadeIn("slow");
			$(".accordion").accordion("refresh");

			$('#officebar .button').addClass('ui-state-default');

			$('#officebar .button').hover(
				function() { if (!$(this).hasClass("disabled")) $(this).addClass('ui-state-hover'); },
				function() { if (!$(this).hasClass("disabled")) $(this).removeClass('ui-state-hover'); }
			);

		}


		$('header').show();
		var doit;
		var minHeight = $('#menuframe').height();
		$('.ContextMenu-panel').resizable({
			handles: 'e,w',
			resize: function() {
				clearTimeout(doit);
				doit = setTimeout(resizedw, 100);
			}
		});

		if (SelfServiceUserType == 'False') {
			var splitFunc = 'resizedw(' + getCookie('Intranet_MenuWidth') + ')';
			setTimeout(splitFunc, 50);
		}

		resetTimeout(); //Site.master function.

	});
	

	function resizedw(splitPos) {		
		if (!(Number(splitPos) > 0)) {
			splitPos = $('.ContextMenu-panel').width();
		} else {
			$('#menuframe').width(splitPos);
		}
		
		$("#workframeset").css("left", splitPos);

		//resize defsel/find screen accordingly.
		$('#findGridTable').setGridWidth($('#findGridRow').width());
		$('#DefSelRecords').setGridWidth($('#findGridRow').width());

		//save the width to a cookie for next time.
		setCookie('Intranet_MenuWidth', splitPos, 365); //setcookie is in site.master
	}


</script>



<%session("utilid")="" %>

	<div id="menuframe" class="open" style="display: none;">
		<div class="ContextMenu-panel">
			<div class="ContextMenu-tab ui-state-default" style="text-align: center; vertical-align: middle">
				<p style="display: none" class="rot-neg-90">Click here to show/hide this accordian menu...</p>
				<img id="ContextMenuIcon" src="<%= Url.Content("~/content/images/retract.png") %>" alt="x" />
			</div>
			<div class="ContextMenu-content">
				<%	Html.RenderPartial("~/views/home/menu.ascx")%>
			</div>
		</div>
	</div>

	<div id="mainframeset">

		<div id="workframeset" style="display: block;" class="ui-widget ui-widget-content">
			<div id="SSILinksFrame" style="display: none"></div>
			<div id="workframe" data-framesource="default.asp"><%Html.RenderPartial("~/views/home/_default.ascx")%></div>
			<div id="optionframe" data-framesource="emptyoption.asp" style="display: none"><%Html.RenderPartial("~/views/home/emptyoption.ascx")%></div>
		</div>

		<div id="optionframeset">
			<div id="dataframe" data-framesource="data.asp" style="display: none"><%Html.RenderPartial("~/views/home/data.ascx")%></div>
			<div id="optiondataframe" data-framesource="optionData.asp" style="display: none"><%Html.RenderPartial("~/views/home/optiondata.ascx")%></div>
		</div>

		<div id="reportframeset" class="popup" data-framesource="util_run" style="">
			<div id="reportframe" style="height: 100%"></div>
		</div>

	</div>

	<%Session("LoggingIn") = False%>	

</asp:Content>


