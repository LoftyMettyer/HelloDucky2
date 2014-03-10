<%@ Page Title="" Language="VB" MasterPageFile="~/Views/Shared/Site.Master" Inherits="System.Web.Mvc.ViewPage" %>

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

			var SelfServiceUserType = '<%=ViewBag.SSIMode%>';

			if (SelfServiceUserType == 'True') {
				$("#workframeset").css("left", "0px");
				$("#reportframeset").css("left", "0px");
			}
			else {

				// ----  Apply jQuery functionality to the slide out CONTEXT MENU  ----
				var ContextMenuTab = {
					speed: 300,
					containerWidth: $('.ContextMenu-panel').outerWidth() - 30,
					containerHeight: $('.ContextMenu-panel').outerHeight(),
					tabWidth: $('.ContextMenu-tab').outerWidth(),
					init: function () {
						//$('.ContextMenu-panel').css('height', ContextMenuTab.containerHeight + 'px');
						$('.ContextMenu-tab').click(function (event) {
							if ($('.ContextMenu-panel').hasClass('open')) {
								$('.ContextMenu-panel').animate({ left: '-' + ContextMenuTab.containerWidth }, ContextMenuTab.speed)
									.removeClass('open');
								$("#workframeset").css("left", "30px");
								// $("#reportframeset").css("left", "30px");
								$('#ContextMenuIcon').attr('src', '<%= Url.Content("~/content/images/expand.png") %>');

							} else {
								$('.ContextMenu-panel').animate({ left: '0' }, ContextMenuTab.speed)
									.addClass('open');
								$("#workframeset").css("left", "350px");
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
				ContextMenuTab.init();
			}
		});

	$(document).ready(function () {

			OpenHR.CheckForMessages();
		
				$("#fixedlinksframe").show();
				$("#FixedLinksContent").fadeIn("slow");

				$(".popup").dialog({
						overflow: false,
						autoOpen: false,
						modal: true,
						height: 550,
						width: 800
				});		

			//load menu for dmi, or linksmain for ssi
			var SelfServiceUserType = '<%=ViewBag.SSIMode%>';

			if (SelfServiceUserType == 'True') {
				
				$.ajax({
					url: 'linksMain',
					dataType: 'html',
					type: 'POST',
					data: { psScreenInfo: '<%=session("SingleRecordTableID")%>!<%=session("SingleRecordViewID")%>_0' },
					success: function (html) {
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
						if (window.currentLayout == 'tiles') {
							var pwfswidth = Number($('.pendingworkflowsframe').css('width').replace('px', ''));
							var hlwidth = Number(document.querySelector('.hypertextlinks').offsetWidth);
							var buttonwidth = Number($('.linkspagebutton').css('width').replace('px', ''));
							if ((pwfswidth > 0) && (hlwidth > 0) && (buttonwidth > 0)) {
								var requiredWidth = pwfswidth + hlwidth + buttonwidth + 300 + 300;								
								requiredWidth += 'px';								
								$('.tileContent').css('width', requiredWidth);
							}
						}

					},
					error: function (req, status, errorObj) {
						debugger;
					}
				});
			} else {
				$("#menuframe").fadeIn("slow");
				$(".accordion").accordion("resize");
				$('#officebar .button').addClass('ui-state-default');

				$('#officebar .button').hover(
					function () { if (!$(this).hasClass("disabled")) $(this).addClass('ui-state-hover'); },
					function () { if (!$(this).hasClass("disabled")) $(this).removeClass('ui-state-hover'); }
				);

			}

		//Timeout functionality
		try {
			window.timeoutMs = (Number('<%=Session("TimeoutSecs")%>') * 1000);
			}
		catch (e) {
			//default to 20 minutes.
			window.timeoutMs = 1200000;
		}

		window.timeoutHandle = window.setTimeout('OpenHR.SessionTimeout();', window.timeoutMs);

		});



</script>



<%session("utilid")="" %>


<div id="menuframe" style="display: none;">
<div class="ContextMenu-panel open">
<div class="ContextMenu-tab ui-state-default" style="text-align:center;vertical-align: middle">
<p style="display:none" class="rot-neg-90">Click here to show/hide this accordian menu...</p>
<img id="ContextMenuIcon" src="<%= Url.Content("~/content/images/retract.png") %>" alt="x"/>
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

</asp:Content>


