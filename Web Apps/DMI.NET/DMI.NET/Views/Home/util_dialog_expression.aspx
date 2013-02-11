<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>

<!DOCTYPE html>

<html>
<head runat="server">
    <title>OpenHR Intranet</title>
    
    <link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css" />
    <script src="<%: Url.Content("~/Scripts/jquery-1.8.2.js") %>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/Scripts/openhr.js") %>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/Scripts/ctl_SetFont.js") %>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/Scripts/ctl_SetStyles.js") %>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/Scripts/jquery-ui-1.9.1.custom.min.js") %>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/Scripts/jquery.cookie.js") %>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/Scripts/menu.js")%>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/Scripts/jquery.ui.touch-punch.min.js") %>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/Scripts/jsTree/jquery.jstree.js") %>" type="text/javascript"></script>
    <script id="officebarscript" src="<%: Url.Content("~/Scripts/officebar/jquery.officebar.js") %>" type="text/javascript"></script>


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
                
        <FORM id=frmUseful name=frmUseful style="visibility:hidden;display:none">
	        <INPUT type=hidden id=action name=action value=<%=Request("action")%>>
        </FORM>

        <FORM id=frmValidate name=frmValidate method=post action=util_validate_expression style="visibility:hidden;display:none">
	        <INPUT type=hidden id=validatePass name=validatePass>
	        <INPUT type=hidden id=validateName name=validateName>
	        <INPUT type=hidden id=validateOwner name=validateOwner>
	        <INPUT type=hidden id=validateTimestamp name=validateTimestamp>
	        <INPUT type=hidden id=validateUtilID name=validateUtilID>
	        <INPUT type=hidden id=validateUtilType name=validateUtilType>
	        <INPUT type=hidden id=validateAccess name=validateAccess>
	        <INPUT type=hidden id=components1 name=components1>
	        <INPUT type=hidden id=validateBaseTableID name=validateBaseTableID>>
	        <INPUT type=hidden id=validateOriginalAccess name=validateOriginalAccess>
        </FORM>

        <FORM id=frmTest name=frmTest method=post action=util_test_expression_pval style="visibility:hidden;display:none">
	        <INPUT type="hidden" id=type name=type>	
	        <INPUT type="hidden" id=Hidden1 name=components1>
            <INPUT type="hidden" id=tableID name=tableID>
	        <INPUT type="hidden" id=prompts name=prompts>
	        <INPUT type="hidden" id=filtersAndCalcs name=filtersAndCalcs>
        </FORM>

    </div>

</body>
</html>

<script type="text/javascript">
    util_dialog_expression_onload();
</script>
