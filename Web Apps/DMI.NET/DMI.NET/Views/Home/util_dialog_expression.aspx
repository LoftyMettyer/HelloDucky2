<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>

<!DOCTYPE html>

<html>
<head runat="server">
    <title>OpenHR Intranet</title>

    <link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css" />

    <script src="<%: Url.Content("~/bundles/jQuery")%>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/bundles/OpenHR_General")%>" type="text/javascript"></script>

<script type="text/javascript">

    function util_dialog_expression_onload() {

        if(frmUseful.action.value == "test") {
            var frmParentTest = window.dialogArguments.document.forms("frmTest");
	
            frmTest.type.value = frmParentTest.type.value;
            frmTest.components1.value = frmParentTest.components1.value;
            frmTest.tableID.value = frmParentTest.tableID.value;
            frmTest.prompts.value = frmParentTest.prompts.value;
            frmTest.filtersAndCalcs.value = frmParentTest.filtersAndCalcs.value;
            OpenHR.submitForm(frmTest);

        }
        else {
            var frmParentValidate = window.dialogArguments.OpenHR.getForm("workframe","frmValidate");
	
            frmValidate.validatePass.value = frmParentValidate.validatePass.value;
            frmValidate.validateName.value = frmParentValidate.validateName.value;
            frmValidate.validateOwner.value = frmParentValidate.validateOwner.value;
            frmValidate.validateTimestamp.value = frmParentValidate.validateTimestamp.value;
            frmValidate.validateUtilID.value = frmParentValidate.validateUtilID.value;
            frmValidate.validateUtilType.value = frmParentValidate.validateUtilType.value;
            frmValidate.validateAccess.value = frmParentValidate.validateAccess.value;
            frmValidate.components1.value = frmParentValidate.components1.value;
            frmValidate.validateBaseTableID.value = frmParentValidate.validateBaseTableID.value;
            frmValidate.validateOriginalAccess.value = frmParentValidate.validateOriginalAccess.value;

            OpenHR.submitForm(frmValidate);
            }
        }
    
</script>


</head>
<body>
    <div id="util_dialog_expression" data-framesource="util_dialog_expression">

        <form id="frmUseful" name="frmUseful" style="visibility: hidden; display: none">
            <input type="hidden" id="action" name="action" value='<%=Request("action")%>'>
        </form>

        <form id="frmValidate" name="frmValidate" method="post" action="util_validate_expression" style="visibility: hidden; display: none">
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
        </form>

        <form id="frmTest" name="frmTest" method="post" action="util_test_expression_pval" style="visibility: hidden; display: none">
            <input type="hidden" id="type" name="type">
            <input type="hidden" id="Hidden1" name="components1">
            <input type="hidden" id="tableID" name="tableID">
            <input type="hidden" id="prompts" name="prompts">
            <input type="hidden" id="filtersAndCalcs" name="filtersAndCalcs">
        </form>

    </div>

</body>
</html>

<script type="text/javascript">
    util_dialog_expression_onload();
</script>
