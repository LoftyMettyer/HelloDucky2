<div id="util_dialog_expression" data-framesource="util_dialog_expression">

	<form id="frmUseful" name="frmUseful" style="visibility: hidden; display: none">
		<input type="hidden" id="action" name="action" value='<%:ViewData("action")%>'>
	</form>

	<form id="frmValidate" name="frmValidate" method="post" action="<%:Url.Action("util_validate_expression", "Home")%>" style="visibility: hidden; display: none">
						<input type="hidden" id="validatePass" name="validatePass">
						<input type="hidden" id="validateName" name="validateName">
						<input type="hidden" id="validateOwner" name="validateOwner">
						<input type="hidden" id="validateTimestamp" name="validateTimestamp">
						<input type="hidden" id="validateUtilID" name="validateUtilID">
						<input type="hidden" id="validateUtilType" name="validateUtilType">
						<input type="hidden" id="validateAccess" name="validateAccess">
						<input type="hidden" id="components1" name="components1">
						<input type="hidden" id="validateBaseTableID" name="validateBaseTableID">>
						<input type="hidden" id="validateOriginalAccess" name="validateOriginalAccess">
						<%=Html.AntiForgeryToken()%>
				</form>

				<form id="frmTest" name="frmTest" method="post" action="<%:Url.Action("util_test_expression_pval", "Home")%>" style="visibility: hidden; display: none">
						<input type="hidden" id="type" name="type">
						<input type="hidden" id="components1" name="components1">
						<input type="hidden" id="tableID" name="tableID" value='<%:Session("utiltableid")%>'>
						<input type="hidden" id="prompts" name="prompts">
						<input type="hidden" id="filtersAndCalcs" name="filtersAndCalcs">
				</form>

		</div>

<script type="text/javascript">
	
	var frmUseful = $('#util_dialog_expression #frmUseful');
	if ($(frmUseful).find('#action').val() == "test") {
		var frmParentTest = $('#divDefExpression #frmTest');
		var frmTest = $('#util_dialog_expression #frmTest');

		frmTest.find('#type').val(frmParentTest.find('#type').val());
		frmTest.find('#components1').val(frmParentTest.find('#components1').val());
		frmTest.find('#tableID').val(frmParentTest.find('#tableID').val());
		frmTest.find('#prompts').val(frmParentTest.find('#prompts').val());
		frmTest.find('#filtersAndCalcs').val(frmParentTest.find('#filtersAndCalcs').val());
		OpenHR.submitForm(frmTest, 'tmpDialog');
	}
	else {
		var frmParentValidate = $("#divDefExpression #frmValidate");
		var frmValidate = $('#util_dialog_expression #frmValidate');

		frmValidate.find('#validatePass').val(frmParentValidate.find('#validatePass').val());
		frmValidate.find('#validateName').val(frmParentValidate.find('#validateName').val());
		frmValidate.find('#validateOwner').val(frmParentValidate.find('#validateOwner').val());
		frmValidate.find('#validateTimestamp').val(frmParentValidate.find('#validateTimestamp').val());
		frmValidate.find('#validateUtilID').val(frmParentValidate.find('#validateUtilID').val());
		frmValidate.find('#validateUtilType').val(frmParentValidate.find('#validateUtilType').val());
		frmValidate.find('#validateAccess').val(frmParentValidate.find('#validateAccess').val());
		frmValidate.find('#components1').val(frmParentValidate.find('#components1').val());
		frmValidate.find('#validateBaseTableID').val(frmParentValidate.find('#validateBaseTableID').val());
		frmValidate.find('#validateOriginalAccess').val(frmParentValidate.find('#validateOriginalAccess').val());

		OpenHR.submitForm(frmValidate, 'tmpDialog');
	}

</script>
