<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

	<script type="text/javascript">
		function cancelClick() {
			$('#EventLogDelete').dialog("close");
		}

		function deleteClick() {
			var sEventIDs;
			var frmOpenerDelete = OpenHR.getForm("workframe", "frmDelete");
			var frmOpenerLog = OpenHR.getForm("workframe", "frmLog");
			var LogEvents = OpenHR.getForm("workframe", "LogEvents");
			
			sEventIDs = '';
			var i;
			if ($('#optSelection1').prop('checked') == true) { //Only the currently highlighted row(s)
				frmOpenerDelete.txtDeleteSel.value = 0;

				var eventID;
				var selectedRows = $(LogEvents).jqGrid('getGridParam', 'selarrrow');
				for (i = 0; i <= selectedRows.length - 1; i++) {
					var rowData = $(LogEvents).getRowData(selectedRows[i]);
					eventID = rowData["ID"];
					sEventIDs = sEventIDs + eventID + ",";
				}

				sEventIDs = sEventIDs.substr(0, sEventIDs.length - 1);
			} else if ($('#optSelection2').prop('checked') == true) { //All entries currently displayed
				frmOpenerDelete.txtDeleteSel.value = 1;

				var allRows = $(LogEvents).jqGrid('getGridParam', 'data');
				for (i = 0; i <= allRows.length - 1; i++) {
					sEventIDs = sEventIDs + allRows[i]["ID"] + ",";
				}

				sEventIDs = sEventIDs.substr(0, sEventIDs.length - 1);
			} else if ($('#optSelection3').prop('checked') == true) { //All entries (that the current user has permission to see)
				frmOpenerDelete.txtDeleteSel.value = 2;
			}
			
			frmOpenerDelete.txtSelectedIDs.value = sEventIDs;
			frmOpenerDelete.txtCurrentUsername.value = frmOpenerLog.cboUsername.options[frmOpenerLog.cboUsername.selectedIndex].value;
			frmOpenerDelete.txtCurrentType.value = frmOpenerLog.cboType.options[frmOpenerLog.cboType.selectedIndex].value;
			frmOpenerDelete.txtCurrentMode.value = frmOpenerLog.cboMode.options[frmOpenerLog.cboMode.selectedIndex].value;
			frmOpenerDelete.txtCurrentStatus.value = frmOpenerLog.cboStatus.options[frmOpenerLog.cboStatus.selectedIndex].value;
			frmOpenerDelete.txtViewAllPerm.value = frmOpenerLog.txtELViewAllPermission.value;

			OpenHR.submitForm(frmOpenerDelete);
			$('#EventLogDelete').dialog("close");
		}
	</script>

	<form id="frmEventSelection" name="frmEventSelection">
		<div>
			<div class="pageTitleDiv padbot15">
				<span class="pageTitle" id="PopupEventDeleteTitle">Delete Events</span>
			</div>

			<div class="padleft20 padbot10">
				<div class="padbot10">
					Please select the entries you wish to delete from the options below : 
				</div>

				<div class="padbot5">
					<input id="optSelection1" name="optSelection" type="radio" checked>
					<label for="optSelection1" tabindex="-1">Only the currently highlighted row(s)</label>
				</div>

				<div class="padbot5">
					<input id="optSelection2" name="optSelection" type="radio">
					<label for="optSelection2" tabindex="-1">All entries currently displayed</label>				
				</div>

				<div class="padbot5">
					<input id="optSelection3" name="optSelection" type="radio">
					<label for="optSelection3" tabindex="-1">All entries (that the current user has permission to see)</label>
				</div>
			</div>

			<div id="divEventLogDeleteButtons" class="clearboth">
				<input id="cmdDelete" type="button" value="Delete" name="cmdDelete" onclick="deleteClick();">
				<input id="cmdCancel" type="button" value="Cancel" name="cmdCancel" onclick="cancelClick();">
			</div>
		</div>
	</form>