<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>

<script type="text/javascript">

	function confirmok_window_onload() {
		
		$("#workframe").attr("data-framesource", "CONFIRMOK");

		if ($('#confirmOKParams #txtReloadMenu').val() == 1) {
			location.href = "main";
			return;
		}

		var sAction = $('#confirmOKParams #txtReaction').val();
		
		if (sAction == "LOGOFF") {
			menu_logoffIntranet();
			return;
		}

		if (sAction == "EXIT") {
			window.parent.close();
		}

		if (sAction == "CROSSTABS" || sAction == "NINEBOXGRID") {
			menu_loadDefSelPage($('#confirmOKParams #txtUtilType').val(), $('#confirmOKParams #txtUtilID').val(), 0, true);
		}

		if (sAction == "CUSTOMREPORTS") {
			menu_loadDefSelPage(2, $('#confirmOKParams #txtUtilID').val(), 0, true);
		}

		if (sAction == "CALENDARREPORTS") {
			menu_loadDefSelPage(17, $('#confirmOKParams #txtUtilID').val(), 0, true);
		}

		if (sAction == "MAILMERGE") {
			menu_loadDefSelPage(9, $('#confirmOKParams #txtUtilID').val(), 0, true);
		}

		if (sAction == "WORKFLOW") {
			menu_loadDefSelPage(25, $('#confirmOKParams #txtUtilID').val(), 0, true);
		}

		if (sAction == "WORKFLOWPENDINGSTEPS") {
			menu_autoLoadPage("workflowPendingSteps", false);
		}

		if (sAction == "WORKFLOWOUTOFOFFICE") {
			menu_WorkflowOutOfOffice();
		}

		if (sAction == "PICKLISTS") {
			menu_loadDefSelPage(10, $('#confirmOKParams #txtUtilID').val(), $('#confirmOKParams #txtUtilTableID').val(), true);
		}

		if (sAction == "FILTERS") {
			menu_loadDefSelPage(11, $('#confirmOKParams #txtUtilID').val(), $('#confirmOKParams #txtUtilTableID').val(), true);
			OpenHR.clearTmpDialog();
		}

		if (sAction == "CALCULATIONS") {
			menu_loadDefSelPage(12, $('#confirmOKParams #txtUtilID').val(), $('#confirmOKParams #txtUtilTableID').val(), true);
			OpenHR.clearTmpDialog();
		}

		if (sAction == "DEFAULT") {
			window.location.href = "main";
		}

		if (sAction.substring(0, 7) == "mnutool") {

			menu_loadPage(sAction.substring(7, sAction.length));
			return;
		}
		else {

			var frmData = OpenHR.getForm("dataframe", "frmData");
			var frmMenuInfo = $("#frmMenuInfo")[0].children;

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

		try {
			if ($('#confirmOKParams #txtErrorMessage').val().length > 0) {
				// An error returned from the server
				$('#errorMessages').html('').html($('#confirmOKParams #txtErrorMessage').val());
				$('#errorMessages').append('<br/><br/><input type="button" value="Close" style="float: right; width: 80px;" onclick="OpenHR.clearTmpDialog();"/>');
			}
		} catch (e) {}

	}

</script>

<div id="confirmOKParams">
<%
	Response.Write("<INPUT type='hidden' id=txtFollowPage name=txtFollowPage value=" & Session("followpage") & ">")
	Response.Write("<INPUT type='hidden' id=txtReaction name=txtReaction value=""" & Session("reaction") & """>")
	Response.Write("<INPUT type='hidden' id=txtUtilID name=txtUtilID value=" & Session("utilid") & ">")
	Response.Write("<INPUT type='hidden' id=txtUtilType name=txtUtilType value=" & Session("utiltype") & ">")
	Response.Write("<INPUT type='hidden' id=txtUtilTableID name=txtUtilTableID value=" & Session("utilTableID") & ">")
	Response.Write("<INPUT type='hidden' id=txtReloadMenu name=txtReloadMenu value=" & Session("reloadMenu") & ">")
	Response.Write("<INPUT type='hidden' id='txtErrorMessage' name='txtErrorMessage' value='" & Session("errorMessage") & "'>")
	Session("confirmtitle") = Nothing
	Session("confirmtext") = Nothing
	Session("followpage") = Nothing
	Session("reaction") = Nothing
	Session("reloadMenu") = 0
%>
</div>
<div id="errorMessages"></div>

<script type="text/javascript">
	confirmok_window_onload();
</script>
