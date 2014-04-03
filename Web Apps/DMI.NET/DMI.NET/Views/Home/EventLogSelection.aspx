<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>

<!DOCTYPE html>
<html>
<head>
	<title>Event Log Selection - OpenHR</title>
	<script src="<%: Url.LatestContent("~/bundles/jQuery")%>" type="text/javascript"></script>
	<script src="<%: Url.LatestContent("~/bundles/OpenHR_General")%>" type="text/javascript"></script>
	<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />

	<%--Here's the stylesheets for the font-icons displayed on the dashboard for wireframe and tile layouts--%>
	<link href="<%= Url.LatestContent("~/Content/general_enclosed_foundicons.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/font-awesome.css")%>" rel="stylesheet" type="text/css" />
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
			self.close();
		}

		function deleteClick() {
			var sEventIDs;

			var frmOpenerDelete = window.dialogArguments.OpenHR.getForm("workframe", "frmDelete");
			var frmOpenerLog = window.dialogArguments.OpenHR.getForm("workframe", "frmLog");

			sEventIDs = '';

			if (frmEventSelection.optSelection1.checked == true) {
				frmOpenerDelete.txtDeleteSel.value = 0;

				frmOpenerLog.ssOleDBGridEventLog.Redraw = false;

				for (var i = 0; i < frmOpenerLog.ssOleDBGridEventLog.selbookmarks.count; i++) {
					sEventIDs = sEventIDs + frmOpenerLog.ssOleDBGridEventLog.Columns("ID").cellvalue(frmOpenerLog.ssOleDBGridEventLog.selbookmarks(i)) + ',';
				}

				sEventIDs = sEventIDs.substr(0, sEventIDs.length - 1);

				frmOpenerLog.ssOleDBGridEventLog.Redraw = true;
			}


			else if (frmEventSelection.optSelection2.checked == true) {
				frmOpenerDelete.txtDeleteSel.value = 1;

				frmOpenerLog.ssOleDBGridEventLog.Redraw = false;

				frmOpenerLog.ssOleDBGridEventLog.MoveFirst();

				for (var i = 0; i < frmOpenerLog.ssOleDBGridEventLog.Rows; i++) {
					sEventIDs = sEventIDs + frmOpenerLog.ssOleDBGridEventLog.Columns("ID").cellvalue(frmOpenerLog.ssOleDBGridEventLog.AddItemBookmark(i)) + ',';
				}

				sEventIDs = sEventIDs.substr(0, sEventIDs.length - 1);

				frmOpenerLog.ssOleDBGridEventLog.Redraw = true;
			}


			else if (frmEventSelection.optSelection3.checked == true) {
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
		<table align="center" cellpadding="5" cellspacing="0" width="100%" height="100%">
			<tr>
				<td>
					<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
						<tr>
							<td>
								<table height="100%" width="100%" class="invisible" cellspacing="0" cellpadding="4">
									<tr height="30">
										<td colspan="4">Please the select the entries you wish to delete from the options below : 
										</td>
									</tr>
									<tr height="15">
										<td></td>
										<td width="8"></td>
										<td>
											<input id="optSelection1" name="optSelection" type="radio" checked>
										</td>
										<td>
											<label
												tabindex="-1"
												for="optSelection1"
												class="radio"/>
											Only the currently highlighted row(s)
										</td>
									</tr>
									<tr height="15">
										<td></td>
										<td width="8"></td>
										<td>
											<input id="optSelection2" name="optSelection" type="radio">
										</td>
										<td>
											<label
												tabindex="-1"
												for="optSelection2"
												class="radio" />
											All entries currently displayed
										</td>
									</tr>
									<tr height="15">
										<td></td>
										<td width="8"></td>
										<td>
											<input id="optSelection3" name="optSelection" type="radio">
										</td>
										<td nowrap>
											<label
												tabindex="-1"
												for="optSelection3"
												class="radio"/>
											All entries (that the current user has permission to see)
										</td>
									</tr>
									<tr height="5">
										<td colspan="4"></td>
									</tr>
									<tr>
										<td width="100%" colspan="4">
											<table height="100%" width="100%" class="invisible" cellspacing="0" cellpadding="4">
												<tr>
													<td></td>
													<td width="5">
														<input id="cmdDelete" type="button" value="Delete" name="cmdDelete" style="WIDTH: 80px" width="80" class="button ui-button ui-widget ui-state-default ui-widget-content ui-corner-tl ui-corner-br"
															onclick="deleteClick();">
													</td>
													<td width="5">
														<input id="cmdCancel" type="button" value="Cancel" name="cmdCancel" style="WIDTH: 80px" width="80" class="button ui-button ui-widget ui-state-default ui-widget-content ui-corner-tl ui-corner-br"
															onclick="cancelClick();">
													</td>
													<td></td>
												</tr>
											</table>
										</td>
									</tr>
								</table>
							</td>
							<td width="5"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</form>

</body>
</html>
