<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script type="text/javascript">
    function util_dialog_picklist_onload() {

        $("#reportframe").attr("data-framesource", "UTIL_DIALOG_PICKLIST");

        var frmParentAdd = OpenHR.getForm("workframe", "frmPicklistSelection");
        var frmAdd = document.getElementById("frmAdd");

        frmAdd.selectionType.value = frmParentAdd.selectionType.value;
        frmAdd.txtTableID.value = frmParentAdd.txtTableID.value;
        frmAdd.selectedIDs1.value = frmParentAdd.selectedIDs1.value;

        OpenHR.submitForm(frmAdd);
    }
</script>

<div id="picklistdialog" data-framesource="util_dialog_picklist">

    <form id="frmAdd" name="frmAdd" method="post" action="picklistSelectionMain" style="visibility: hidden; display: none">
        <input type="hidden" id="selectionType" name="selectionType">
        <input type="hidden" id="txtTableID" name="txtTableID">
        <input type="hidden" id="selectedIDs1" name="selectedIDs1">
    </form>

</div>


<script type="text/javascript">
    util_dialog_picklist_onload();
</script>
