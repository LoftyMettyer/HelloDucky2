<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

<input type="hidden" id="txtDummyForJQuery" name="txtDummyForJQuery" value="0">

<script type="text/javascript">
	if ($('#tmpDialog').dialog('isOpen') == true) {
		//prompted Values for OpenHR.modalExpressionSelect screen.
		makeSelection('FILTER', '<%:Session("filterIDvalue")%>', '<%=Session("promptsvalue")%>');
		OpenHR.clearTmpDialog();
	} else {
		picklistdef_makeSelection('FILTER', '<%:Session("filterIDvalue")%>', '<%=Session("promptsvalue")%>');
	}
</script>
