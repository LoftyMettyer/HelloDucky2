<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<script type="text/javascript">
	function default_window_onload() {
		try {			
		// Do nothing if the menu controls are not yet instantiated.
		if (OpenHR.getForm("menuframe", "frmMenuInfo") != null)	{
			$("#workframe").attr("data-framesource", "DEFAULT");

		}
	}
	catch(e) {}
	}

</script>

<form action="default_Submit" method="post" id="frmGoto" name="frmGoto">
	<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</form>

<script type="text/javascript">default_window_onload();</script>