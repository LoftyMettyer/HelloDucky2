<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>

<script type="text/javascript">
	function tbBulkBooking_onload() {
		var frmBulkBooking = document.getElementById("frmBulkBooking");

		$("#optionframe").attr("data-framesource", "TBBULKBOOKING");
		$("#workframe").hide();
		$("#optionframe").show();

		frmBulkBooking.cmdCancel.focus();

		tbrefreshControls();
		menu_refreshMenu();		
	}

	function ok() {
		var sSelectedIDs = "";
		var frmBulkBooking = document.getElementById("frmBulkBooking");

		sSelectedIDs = $('#ssOleDBGridFindRecords').getDataIDs().join(",");

		var frmGotoOption = document.getElementById("frmGotoOption");

		frmGotoOption.txtGotoOptionAction.value = "SELECTBULKBOOKINGS";
		frmGotoOption.txtGotoOptionRecordID.value = $("#txtOptionRecordID").val();
		frmGotoOption.txtGotoOptionLinkRecordID.value = sSelectedIDs;
		<%If Session("TB_TBStatusPExists") Then%>
		frmGotoOption.txtGotoOptionLookupValue.value = frmBulkBooking.selStatus.options[frmBulkBooking.selStatus.selectedIndex].value;
		<%Else%>
		frmGotoOption.txtGotoOptionLookupValue.value = "B";
		<%End If%>

		frmGotoOption.txtGotoOptionPage.value = "emptyoption";
		OpenHR.submitForm(frmGotoOption);
	}

	function cancel() {
		$("#optionframe").hide();
		$("#workframe").show();

		var frmGotoOption = document.getElementById("frmGotoOption");

		frmGotoOption.txtGotoOptionAction.value = "CANCEL";
		frmGotoOption.txtGotoOptionLinkRecordID.value = 0;
		frmGotoOption.txtGotoOptionPage.value = "emptyoption";
		OpenHR.submitForm(frmGotoOption);
	}


	function tbrefreshControls() {
		var fNoneSelected;
		var frmBulkBooking = document.getElementById("frmBulkBooking");

		var selRowId = $("#ssOleDBGridFindRecords").jqGrid('getGridParam', 'selrow');

		fNoneSelected = (selRowId == null || selRowId == 'undefined');
		fGridHasRows = ($("#ssOleDBGridFindRecords").jqGrid('getGridParam', 'records') > 0);
	
		button_disable(frmBulkBooking.cmdRemove, fNoneSelected);
		button_disable(frmBulkBooking.cmdRemoveAll, !fGridHasRows);
		button_disable(frmBulkBooking.cmdOK, !fGridHasRows);

		$('#FindGridRow').css('border', (fGridHasRows ? 'none' : '1px solid silver'));
	}

	function add() {
		var sURL;
		var frmBookingSelection = document.getElementById("frmBookingSelection");

		frmBookingSelection.selectionType.value = "ALL";

		sURL = "tbBulkBookingSelectionMain" +
			"?selectionType=" + frmBookingSelection.selectionType.value;
		openDialog(sURL, (screen.width) / 3, (screen.height) / 2);
	}

	function filteredAdd() {
		var sURL;
		var frmBookingSelection = document.getElementById("frmBookingSelection");

		frmBookingSelection.selectionType.value = "FILTER";

		sURL = "tbBulkBookingSelectionMain" +
			"?selectionType=" + frmBookingSelection.selectionType.value;
		openDialog(sURL, (screen.width) / 3, (screen.height) / 2);
	}

	function addPicklist() {
		var sURL;
		var frmBookingSelection = document.getElementById("frmBookingSelection");
		frmBookingSelection.selectionType.value = "PICKLIST";

		sURL = "tbBulkBookingSelectionMain" +
			"?selectionType=" + frmBookingSelection.selectionType.value;
		openDialog(sURL, (screen.width) / 3, (screen.height) / 2);
	}

	function remove() {

		var grid = $("#ssOleDBGridFindRecords")
		var myDelOptions = {
			// because I use "local" data I don't want to send the changes
			// to the server so I use "processing:true" setting and delete
			// the row manually in onclickSubmit
			onclickSubmit: function (options) {
				var grid_id = $.jgrid.jqID(grid[0].id),
						grid_p = grid[0].p,
						newPage = grid_p.page,
						rowids = grid_p.multiselect ? grid_p.selarrrow : [grid_p.selrow];

				// reset the value of processing option which could be modified
				options.processing = true;

				// delete the row
				$.each(rowids, function () {
					grid.delRowData(this);
				});
				$.jgrid.hideModal("#delmod" + grid_id,
													{
														gb: "#gbox_" + grid_id,
														jqm: options.jqModal, onClose: options.onClose
													});

				if (grid_p.lastpage > 1) {// on the multipage grid reload the grid
					if (grid_p.reccount === 0 && newPage === grid_p.lastpage) {
						// if after deliting there are no rows on the current page
						// which is the last page of the grid
						newPage--; // go to the previous page
					}
					// reload grid to make the row from the next page visable.
					grid.trigger("reloadGrid", [{ page: newPage }]);
				}

				return true;
			},
			processing: true
		};

		grid.jqGrid('delGridRow', grid.jqGrid('getGridParam', 'selarrrow'), myDelOptions);

		moveFirst();
		tbrefreshControls();
	}

	function removeAll() {
		$('#ssOleDBGridFindRecords').jqGrid('clearGridData');		
		tbrefreshControls();
	}

	function makeSelection(psType, piID, psPrompts) {
		/* Get the current selected delegate IDs. */
		var frmBulkBooking = document.getElementById("frmBulkBooking");
		var sSelectedIDs = "";
		var sRecordID;

		sSelectedIDs = $('#ssOleDBGridFindRecords').getDataIDs().join(",");

		if ((psType == "ALL") && (psPrompts.length > 0)) {
			if (sSelectedIDs.length > 0) {
				sSelectedIDs = sSelectedIDs + ",";
			}
			sSelectedIDs = sSelectedIDs + psPrompts;
		}


		// Get the optionData.asp to get the required records.
		var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
		optionDataForm.txtOptionAction.value = "GETBULKBOOKINGSELECTION";
		optionDataForm.txtOptionPageAction.value = psType;
		optionDataForm.txtOptionRecordID.value = piID;
		optionDataForm.txtOptionValue.value = sSelectedIDs;
		optionDataForm.txtOptionPromptSQL.value = psPrompts;
		optionDataForm.txtOption1000SepCols.value = frmBulkBooking.txt1000SepCols.value;
		refreshOptionData(); //should be in scope.		
	}

	function moveFirst() {
		$('#ssOleDBGridFindRecords').jqGrid('setSelection', 1);
		menu_refreshMenu();
	}

	function bookmarksCount() {
		var selRowIds = $('#ssOleDBGridFindRecords').jqGrid('getGridParam', 'selarrrow');
		return selRowIds.length;
	}


	function openDialog(pDestination, pWidth, pHeight) {
		var dlgwinprops = "center:yes;" +
			"dialogHeight:" + pHeight + "px;" +
			"dialogWidth:" + pWidth + "px;" +
			"help:no;" +
			"resizable:no;" +
			"scroll:yes;" +
			"status:no;";
		window.showModalDialog(pDestination, self, dlgwinprops);
	}


</script>

<script src="<%: Url.LatestContent("~/Scripts/ctl_SetStyles.js")%>" type="text/javascript"></script>

<div <%=session("BodyTag")%>>
	<form name="frmBulkBooking" action="tbBulkBooking_Submit" method="post" id="frmBulkBooking" style="text-align: center">

		<table align="center" class="outline" cellpadding="5" cellspacing="0" width="100%" height="100%">
			<tr>
				<td>
					<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
						<tr>
							<td align="center" height="10" colspan="5">
								<span style="margin-left: 10px; float: left;" class="pageTitle">Bulk Booking</span>
							</td>
						</tr>

						<%
							Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
							Dim sErrorDescription = ""
	
							If Session("TB_TBStatusPExists") Then
								Response.Write("				<TR height=10>" & vbCrLf)
								Response.Write("					<TD width=20></TD>" & vbCrLf)
								Response.Write("					<TD colspan=3>" & vbCrLf)
								Response.Write("						<TABLE WIDTH=""100%"" class=""invisible"" CELLSPACING=0 CELLPADDING=0>" & vbCrLf)
								Response.Write("							<TR height=10>" & vbCrLf)
								Response.Write("								<TD  nowrap>Booking Status :</TD>" & vbCrLf)
								Response.Write("								<TD width=20>&nbsp;</TD>" & vbCrLf)
								Response.Write("								<TD>" & vbCrLf)
								Response.Write("									<SELECT id=selStatus name=selStatus class=""combo"">" & vbCrLf)
								Response.Write("										<OPTION value=B selected>Booked</OPTION>" & vbCrLf)
								Response.Write("										<OPTION value=P>Provisional</OPTION></SELECT>" & vbCrLf)
								Response.Write("								</TD>" & vbCrLf)
								Response.Write("								<TD style='width: 100%;'></TD>" & vbCrLf)
								Response.Write("								<TD ></TD>" & vbCrLf)
								Response.Write("							</TR>" & vbCrLf)
								Response.Write("						</TABLE>" & vbCrLf)
								Response.Write("					</TD>" & vbCrLf)
								Response.Write("					<TD width=20></TD>" & vbCrLf)
								Response.Write("				</TR>" & vbCrLf)
								Response.Write("				<TR>" & vbCrLf)
								Response.Write("				  <td height=10 colspan=5></td>" & vbCrLf)
								Response.Write("				</TR>" & vbCrLf)
							End If
						%>

						<tr>
							<td rowspan="13" width="20">&nbsp;&nbsp;&nbsp;&nbsp;</td>
							<td rowspan="13" width="100%">
								<%Try
										Dim prmErrorMsg As New SqlParameter("psErrorMsg", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}
										Dim prm1000SepCols As New SqlParameter("ps1000SeparatorCols", SqlDbType.VarChar, -1) With {.Direction = ParameterDirection.Output}

										Dim rstFindRecords = objDataAccess.GetFromSP("sp_ASRIntGetTBEmployeeColumns" _
														, prmErrorMsg _
														, prm1000SepCols)
										Response.Write("<INPUT type='hidden' id=txt1000SepCols name=txt1000SepCols value=""" & prm1000SepCols.Value & """>" & vbCrLf)
		
									Catch ex As Exception
										sErrorDescription = "The Employee table find columns could not be retrieved." & vbCrLf & FormatError(ex.Message)
											End Try%>
								<div id="FindGridRow" style="height: 400px; margin-bottom: 50px;">
									<table id="ssOleDBGridFindRecords" name="ssOleDBGridFindRecords" style="width: 100%"></table>
								</div>
							</td>
							<td rowspan="13" width="20">&nbsp;&nbsp;&nbsp;&nbsp;</td>
							<td width="100" height="10">
								<input type="button" id="cmdAdd" name="cmdAdd" value="Add" style="width: 100%" class="btn"
									onclick="add()" />
							</td>
							<td rowspan="13" width="20">&nbsp;&nbsp;&nbsp;&nbsp;</td>
						</tr>

						<tr>
							<td height="10"></td>
						</tr>

						<tr>
							<td height="10">
								<input type="button" name="cmdAddFilter" id="cmdAddFilter" value="Filtered Add" style="width: 100%; text-align: center" class="btn"
									onclick="filteredAdd()" />
							</td>
						</tr>

						<tr>
							<td height="10"></td>
						</tr>

						<tr>
							<td height="10">
								<input type="button" name="cmdAddPicklist" id="cmdAddPicklist" value="Picklist Add" style="width: 100%; text-align: center" class="btn"
									onclick="addPicklist()" />
							</td>
						</tr>

						<tr>
							<td height="10"></td>
						</tr>

						<tr>
							<td height="10">
								<input type="button" id="cmdRemove" name="cmdRemove" value="Remove" style="width: 100%" class="btn"
									onclick="remove()" />
							</td>
						</tr>

						<tr>
							<td height="10"></td>
						</tr>

						<tr>
							<td height="10">
								<input type="button" id="cmdRemoveAll" name="cmdRemoveAll" value="Remove All" style="width: 100%" class="btn"
									onclick="removeAll()" />
							</td>
						</tr>

						<tr>
							<td></td>
						</tr>

						<tr>
							<td height="10">
								<input type="button" name="cmdOK" value="OK" style="width: 100%" id="cmdOK" class="btn"
									onclick="ok()" />
							</td>
						</tr>

						<tr>
							<td height="10"></td>
						</tr>

						<tr>
							<td height="10">
								<input type="button" name="cmdCancel" value="Cancel" style="width: 100%" class="btn"
									onclick="cancel()" />
							</td>
						</tr>

						<tr>
							<td height="10" colspan="5"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</form>

	<input type='hidden' id="txtTicker" name="txtTicker" value="0">
	<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind">

	<input type='hidden' id="txtSelectionID" name="txtSelectionID" value="0">
	<input type='hidden' id="txtOptionRecordID" name="txtOptionRecordID" value='<%=session("optionRecordID")%>'>


	<form action="tbBulkBooking_Submit" method="post" id="frmGotoOption" name="frmGotoOption" style="visibility: hidden; display: none">
		<%Html.RenderPartial("~/Views/Shared/gotoOption.ascx")%>
	</form>

	<form id="frmBookingSelection" name="frmBookingSelection" target="tbBulkBookingSelection" action="tbBulkBookingSelectionMain" method="post" style="visibility: hidden; display: none">
		<input type="hidden" id="selectionType" name="selectionType">
	</form>

</div>

<script type="text/javascript">
	tbBulkBooking_onload();
</script>
