<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<script type="text/javascript">
    function confirmok_window_onload() {

    	$("#workframe").attr("data-framesource", "CONFIRMOK");

        if (txtReloadMenu.value == 1) {
            window.parent.location.href = "main";
            return;
        }

        var sAction;
        if (txtReaction.length > 1) {
        	sAction = txtReaction[0].value;
        } else {
        	sAction = txtReaction.value;
        }

        if (sAction == "LOGOFF") {
            menu_logoffIntranet();
            return;
        }

        if (sAction == "EXIT") {
            window.parent.close();
        }

        //if (txtUtilType.value > 0) {
        //	window.parent.frames("menuframe").loadDefSelPage(txtUtilType.value, txtUtilID.value, true);
        //}

        if (sAction == "CROSSTABS" || sAction == "NINEBOXGRID") {
            menu_loadDefSelPage(txtUtilType.value, txtUtilID.value, 0, true);
        }

        if (sAction == "CUSTOMREPORTS") {
            menu_loadDefSelPage(2, txtUtilID.value, 0, true);
        }

        if (sAction == "CALENDARREPORTS") {
            menu_loadDefSelPage(17, txtUtilID.value, 0, true);
        }

        if (sAction == "MAILMERGE") {
            menu_loadDefSelPage(9, txtUtilID.value, 0, true);
        }

        if (sAction == "WORKFLOW") {
            menu_loadDefSelPage(25, txtUtilID.value, 0, true);
        }

        if (sAction == "WORKFLOWPENDINGSTEPS") {
            menu_autoLoadPage("workflowPendingSteps", false);
        }

        if (sAction == "WORKFLOWOUTOFOFFICE") {
            menu_WorkflowOutOfOffice();
        }

        if (sAction == "PICKLISTS") {
            menu_loadDefSelPage(10, txtUtilID.value, txtUtilTableID.value, true);
        }

        if (sAction == "FILTERS") {
            menu_loadDefSelPage(11, txtUtilID.value, txtUtilTableID.value, true);
        }

        if (sAction == "CALCULATIONS") {
            menu_loadDefSelPage(12, txtUtilID.value, txtUtilTableID.value, true);
        }

        if (sAction == "DEFAULT") {
            window.location.href = "main";  // "default.asp";
        }

        if (sAction.substring(0, 7) == "mnutool") {

            menu_loadPage(sAction.substring(7, sAction.length));
            return;
        }
        else {

            var frmData = OpenHR.getForm("dataframe", "frmData");
            var frmMenuInfo = OpenHR.getForm("menuframe", "frmMenuInfo");

            if ((sAction.substring(0, 3) == "PT_") ||
                (sAction.substring(0, 3) == "PV_")) {
                // PT_ = primary table
                // PV_ = primary table view
                if (frmMenuInfo.txtPrimaryStartMode.value == 3) {
	                frmData.txtRecordDescription.value = "";
                    menu_loadFindPageFirst(sAction);
                }
                else {
                    menu_loadRecordEditPage(sAction);
                }
                return;
            }

            if (sAction.substring(0, 3) == "TS_") {
                // TS_ = Table screen
                if (frmMenuInfo.txtLookupStartMode.value == 3) {
	                frmData.txtRecordDescription.value = "";
                    menu_loadFindPageFirst(sAction);
                }
                else {
                    menu_loadRecordEditPage(sAction);
                }
                return;
            }

            if (sAction.substring(0, 3) == "QE_") {
                // QE_ = quick entry screen
                if (frmMenuInfo.txtQuickAccessStartMode.value == 3) {
	                frmData.txtRecordDescription.value = "";
                    menu_loadFindPageFirst(sAction);
                }
                else {
                    menu_loadRecordEditPage(sAction);
                }
                return;
            }

            if (sAction.substring(0, 3) == "HT_") {
                // HT_ = history table
                if (frmMenuInfo.txtHistoryStartMode.value == 3) {
	                frmData.txtRecordDescription.value = "";
                    menu_loadFindPageFirst(sAction);
                }
                else {
                    menu_loadRecordEditPage(sAction);
                }
                return;
            }
        }
    }

</script>

<%
        Response.Write("<INPUT type='hidden' id=txtFollowPage name=txtFollowPage value=" & Session("followpage") & ">")
        Response.Write("<INPUT type='hidden' id=txtReaction name=txtReaction value=""" & Session("reaction") & """>")
        Response.Write("<INPUT type='hidden' id=txtUtilID name=txtUtilID value=" & Session("utilid") & ">")
        Response.Write("<INPUT type='hidden' id=txtUtilType name=txtUtilType value=" & Session("utiltype") & ">")
        Response.Write("<INPUT type='hidden' id=txtUtilTableID name=txtUtilTableID value=" & Session("utilTableID") & ">")
        Response.Write("<INPUT type='hidden' id=txtReloadMenu name=txtReloadMenu value=" & Session("reloadMenu") & ">")
        Session("confirmtitle") = Nothing
        Session("confirmtext") = Nothing
        Session("followpage") = Nothing
        Session("reaction") = Nothing
        Session("reloadMenu") = 0
%>

<form action="default_Submit" method="post" id="frmGoto" name="frmGoto" style="visibility: hidden; display: none">
    <%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</form>

<script type="text/javascript">
	confirmok_window_onload();
</script>
