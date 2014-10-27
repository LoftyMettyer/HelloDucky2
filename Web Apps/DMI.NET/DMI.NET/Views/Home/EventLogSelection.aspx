<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>

<!DOCTYPE html>
<html>
<head>
	<title>Event Log Selection - OpenHR</title>
	<script src="<%: Url.LatestContent("~/bundles/jQuery")%>" type="text/javascript"></script>
	<script src="<%: Url.LatestContent("~/bundles/jQueryUI7")%>" type="text/javascript"></script>
	<script src="<%: Url.LatestContent("~/bundles/OpenHR_General")%>" type="text/javascript"></script>
	<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />
	<script src="<%: Url.LatestContent("~/bundles/eventlog")%>" type="text/javascript"></script>

	<%--Here's the stylesheets for the font-icons displayed on the dashboard for wireframe and tile layouts--%>
	<link href="<%= Url.LatestContent("~/Content/general_enclosed_foundicons.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/font-awesome.min.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/fonts/SSI80v194934/style.css")%>" rel="stylesheet" />

	<%--Base stylesheets--%>
	<link href="<%: Url.LatestContent("~/Content/Site.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />

	<%--stylesheet for slide-out dmi menu--%>
	<link href="<%: Url.LatestContent("~/Content/contextmenustyle.css")%>" rel="stylesheet" type="text/css" />

	<%--ThemeRoller stylesheet--%>
	<link id="DMIthemeLink" href="<%: Url.LatestContent("~/Content/themes/" & Session("ui-admin-theme").ToString() & "/jquery-ui.min.css")%>" rel="stylesheet" type="text/css" />
	
</head>

<body>
	<script type="text/javascript">

		function cancelClick() {
			//self.close();
			$(this).dialog("close");
		}

		function deleteClick() {
			var sEventIDs;

			var frmOpenerDelete = window.dialogArguments.OpenHR.getForm("workframe", "frmDelete");
			var frmOpenerLog = window.dialogArguments.OpenHR.getForm("workframe", "frmLog");
			var LogEvents = window.dialogArguments.OpenHR.getForm("workframe", "LogEvents");
			
			sEventIDs = '';

			if (frmEventSelection.optSelection1.checked == true) { //Only the currently highlighted row(s)
				frmOpenerDelete.txtDeleteSel.value = 0;

				var eventID;
				var selectedRows = $(LogEvents).jqGrid('getGridParam', 'selarrrow');
				for (var i = 0; i <= selectedRows.length - 1; i++) {
					var rowData = $(LogEvents).getRowData(selectedRows[i]);
					eventID = rowData["ID"];
					sEventIDs = sEventIDs + eventID + ",";
				}

				sEventIDs = sEventIDs.substr(0, sEventIDs.length - 1);
			} else if (window.frmEventSelection.optSelection2.checked == true) { //All entries currently displayed
				frmOpenerDelete.txtDeleteSel.value = 1;

				var allRows = $(LogEvents).jqGrid('getGridParam', 'data');
				for (var i = 0; i <= allRows.length - 1; i++) {
					sEventIDs = sEventIDs + allRows[i]["ID"] + ",";
				}

				sEventIDs = sEventIDs.substr(0, sEventIDs.length - 1);
			} else if (window.frmEventSelection.optSelection3.checked == true) { //All entries (that the current user has permission to see)
				frmOpenerDelete.txtDeleteSel.value = 2;
			}
			
			frmOpenerDelete.txtSelectedIDs.value = sEventIDs;
			frmOpenerDelete.txtCurrentUsername.value = frmOpenerLog.cboUsername.options[frmOpenerLog.cboUsername.selectedIndex].value;
			frmOpenerDelete.txtCurrentType.value = frmOpenerLog.cboType.options[frmOpenerLog.cboType.selectedIndex].value;
			frmOpenerDelete.txtCurrentMode.value = frmOpenerLog.cboMode.options[frmOpenerLog.cboMode.selectedIndex].value;
			frmOpenerDelete.txtCurrentStatus.value = frmOpenerLog.cboStatus.options[frmOpenerLog.cboStatus.selectedIndex].value;

			frmOpenerDelete.txtViewAllPerm.value = frmOpenerLog.txtELViewAllPermission.value;

			window.dialogArguments.OpenHR.submitForm(frmOpenerDelete);
			self.close();
		}

	</script>


	<form id="frmEventSelection" name="frmEventSelection">
		<div>
			<div class="pageTitleDiv" style="margin-bottom: 15px">
				<span class="pageTitle" id="PopupReportDefinition_PageTitle">Delete Events</span>
			</div>

			<div class="padleft20 padbot10">
				<div class="padbot10">
					Please the select the entries you wish to delete from the options below : 
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

</body>
</html>
