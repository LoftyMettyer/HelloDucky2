@Imports DMI.NET.ViewModels.Reports
@Inherits System.Web.Mvc.WebViewPage(Of SaveWarningModel)

@Html.TextAreaFor(Function(m) m.ErrorCode)
@Html.TextAreaFor(Function(m) m.ErrorMessage)

Hello Ducky

<input type="button" value="Yes" onclick="commitThisSave();" />
<input type="button" value="No" onclick="cancelThisSave();" />


<script type="text/javascript">

	function commitThisSave() {

		var frmSubmit = $("#frmReportDefintion")[0];
		OpenHR.submitForm(frmSubmit);

		$("#divPopupReportDefinition").dialog("close");
		$("#divPopupReportDefinition").empty();

	}

	function cancelThisSave() {
		$("#divPopupReportDefinition").dialog("close");
		$("#divPopupReportDefinition").empty();
	}





</script>
