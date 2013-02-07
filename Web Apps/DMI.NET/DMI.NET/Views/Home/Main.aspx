<%@ Page Title="" Language="VB" MasterPageFile="~/Views/Shared/Site.Master" Inherits="System.Web.Mvc.ViewPage" %>

<asp:Content ID="Content1" ContentPlaceHolderID="TitleContent" runat="server">
<%
	Dim sReferringPage As String

	' If the database connection session variable does not exist, redirect to the login page.
	If Session("databaseConnection") Is Nothing Then
		Response.Redirect("../account/login")
		Response.Clear()
	Else
		' Only open the form if the referring page was the login page.
		' If it wasn't then redirect to the login page.
		sReferringPage = Request.ServerVariables("HTTP_REFERER")
		If InStrRev(sReferringPage, "/") > 0 Then
			sReferringPage = Mid(sReferringPage, InStrRev(sReferringPage, "/") + 1)
		End If

		If (UCase(sReferringPage) <> UCase("login")) And _
		 (UCase(sReferringPage) <> UCase("loginMessage")) Then

		End If

	End If

	' RH 18/04/01 - Clear this session variable
    Session("utilid") = ""        

%>
    
    <%=DMI.NET.svrCleanup.GetPageTitle("") %>

</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

<script type="text/javascript">

    $(function () {               

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
						$("#reportframeset").css("left", "30px");
						$('#ContextMenuIcon').attr('src', '<%= Url.Content("~/content/images/expand.png") %>');

					} else {
						$('.ContextMenu-panel').animate({ left: '0' }, ContextMenuTab.speed)
                            .addClass('open');
						$("#workframeset").css("left", "350px");
						$("#reportframeset").css("left", "350px");
						$('#ContextMenuIcon').attr('src', '<%= Url.Content("~/content/images/retract.png") %>');
					}
					event.preventDefault();
				});
			}
		};
		ContextMenuTab.init();
    });

    $(document).ready(function() {
        $("#fixedlinksframe").show();

        //Load Poll.asp, then reload every 30 seconds to keep
        //session alive, and check for server messages.
        loadPartialView(); // first time
        // re-call the function each 30 seconds
        window.setInterval("loadPartialView()", 30000);

    });

    function loadPartialView() {
        $.ajax({
            url: "<%:Url.Action("poll", "home")%>",
 		        type: "POST",
 		        success: function (html) {
 		            $("#poll").html(html);
 		        },
 		        error: function (req, status, errorObj) {
 		            //alert("OpenHR.submitForm ajax call to '" + url + "' failed with '" + errorObj + "'.");
 		        }
 		    });
    }


</script>



<%session("utilid")="" %>

<div id="menuframe">
<div class="ContextMenu-panel open">
<div class="ContextMenu-tab" style="text-align:center;vertical-align: middle">
<p style="display:none" class="rot-neg-90">Click here to show/hide this accordian menu...</p>
<img id="ContextMenuIcon" src="<%= Url.Content("~/content/images/retract.png") %>" alt="x"/>
</div>
<div class="ContextMenu-content">
<%	Html.RenderPartial("~/views/home/menu.ascx")%>
</div>
</div>
</div>

<div id="mainframeset">
	
	<div id="workframeset" style="display: block;">
		<div id="workframe" data-framesource="default.asp"><%Html.RenderPartial("~/views/home/default.ascx")%></div>
		<div id="optionframe" data-framesource="emptyoption.asp" style="display: none"><%Html.RenderPartial("~/views/home/emptyoption.ascx")%></div>
	</div>

	<div id="optionframeset">
		<div id="dataframe" data-framesource="data.asp" style="display: none"><%Html.RenderPartial("~/views/home/data.ascx")%></div>
		<div id="optiondataframe" data-framesource="optionData.asp" style="display: none"><%Html.RenderPartial("~/views/home/optiondata.ascx")%></div>
	</div>

	<div id="refresh" data-framesource="refresh.asp" style="display: none"><%Html.RenderPartial("~/views/home/refresh.ascx")%></div>

	<div id="pollframeset">
		<div id="poll" data-framesource="poll.asp" style="display: none"></div>
		<div id="pollmessageframe" data-framesource="pollmessage.asp" style="display: none"><%Html.RenderPartial("~/views/home/pollmessage.ascx")%></div>
	</div>
    
    <div id="reportframeset" class="popup" data-framesource="util_run" style="display: block;">
        <div id="reportframe"></div>
    </div>

    <div id="messageframe" style="display: none">Message Page</div>

	<div id="waitpage" data-framesource="WaitPage.asp" style="display: none">waitpage</div>

</div>

	
</asp:Content>
