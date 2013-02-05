<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>

<!DOCTYPE html>

<html>
<head runat="server">
    <title>OpenHR Intranet</title>
    
    <link href="<%: Url.Content("~/Content/OpenHR.css") %>" rel="stylesheet" type="text/css" />
    <script src="<%: Url.Content("~/Scripts/jquery-1.8.2.js") %>" type="text/javascript"></script>
    <script src="<%: Url.Content("~/Scripts/openhr.js") %>" type="text/javascript"></script>

    <script type="text/javascript">
        function util_dialog_picklist_onload() {

            $("#picklistdialog").attr("data-framesource", "UTIL_DIALOG_PICKLIST");

            if (frmUseful.action.value == "add") {
                var frmParentAdd = window.dialogArguments.OpenHR.getForm("workframe","frmPicklistSelection");

                frmAdd.selectionType.value = frmParentAdd.selectionType.value;
                frmAdd.txtTableID.value = frmParentAdd.txtTableID.value;
                frmAdd.selectedIDs1.value = frmParentAdd.selectedIDs1.value;

                OpenHR.submitForm(frmAdd);
            } else {
                self.close;
            }
        }
    </script>

</head>
<body>
    <div id="picklistdialog" data-framesource="util_dialog_picklist">

        <form id="frmUseful" name="frmUseful" style="visibility: hidden; display: none">
            <input type="hidden" id="action" name="action" value='<%=Request("action")%>'>
        </form>

        <form id="frmAdd" name="frmAdd" method="post" action="picklistSelectionMain" style="visibility: hidden; display: none">
            <input type="hidden" id="selectionType" name="selectionType">
            <input type="hidden" id="txtTableID" name="txtTableID">
            <input type="hidden" id="selectedIDs1" name="selectedIDs1">
        </form>

    </div>
</body>
</html>


    <script type="text/javascript">
        util_dialog_picklist_onload();
    </script>