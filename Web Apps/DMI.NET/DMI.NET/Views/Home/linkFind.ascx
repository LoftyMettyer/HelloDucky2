<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="DMI.NET.Helpers" %>

<script type="text/javascript">
	
	function linkFind_removeAll(jqGridID) {
		//remove all rows from the jqGrid.
		$('#' + jqGridID).jqGrid('clearGridData');
	}

	function linkFind_window_onload() {
		var fOK;
		var frmLinkFindForm = document.getElementById('frmLinkFindForm');
		fOK = true;

		var sErrMsg = frmLinkFindForm.txtErrorDescription.value;
		if (sErrMsg.length > 0) {
			fOK = false;
			OpenHR.messageBox(sErrMsg);
			window.parent.location.replace("login");
		}

		if (fOK == true) {
			sErrMsg = frmLinkFindForm.txtFailureDescription.value;
			if (sErrMsg.length > 0) {
				fOK = false;
				OpenHR.messageBox(sErrMsg);
				CancelLink();
			}
		}

		if (fOK == true) {
			if (frmLinkFindForm.selectView.length == 0) {
				fOK = false;
				OpenHR.messageBox("You do not have permission to read the link table.");
				CancelLink();
			}
		}

		if (fOK == true) {
			if (frmLinkFindForm.selectOrder.length == 0) {
				fOK = false;
				OpenHR.messageBox("You do not have permission to use any of the link table orders.");
				CancelLink();
			}
		}

		if (fOK == true) {
			// Expand the option frame and hide the work frame.
			//window.parent.document.all.item("workframeset").cols = "0, *";
			$("#optionframe").attr("data-framesource", "LINKFIND");
			
	

			// Fault 3300
			frmLinkFindForm.txtOptionLinkViewID.value = frmLinkFindForm.selectView.options[frmLinkFindForm.selectView.selectedIndex].value;

			// Fault 3503
			//TODO: window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "hidden";

			// Get the optionData.asp to get the link find records.
			var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
			optionDataForm.txtOptionAction.value = "LOADFIND";
			optionDataForm.txtOptionTableID.value = frmLinkFindForm.txtOptionLinkTableID.value;
			optionDataForm.txtOptionViewID.value = frmLinkFindForm.selectView.options[frmLinkFindForm.selectView.selectedIndex].value;
			optionDataForm.txtOptionOrderID.value = frmLinkFindForm.selectOrder.options[frmLinkFindForm.selectOrder.selectedIndex].value;
			optionDataForm.txtOptionFirstRecPos.value = 1;
			optionDataForm.txtOptionCurrentRecCount.value = 0;
			optionDataForm.txtOptionPageAction.value = "LOAD";

			frmLinkFindForm.txtOptionLinkOrderID.value = optionDataForm.txtOptionOrderID.value;
			refreshOptionData();	//should be in scope now...
			
			//get width
			$("#optionframe").dialog({
				autoOpen: true,
				modal: true,
				width: 'auto',
				height: 'auto',
				resizable: false
			});

			// Set focus onto one of the form controls. 
			// NB. This needs to be done before making any reference to the grid
			frmLinkFindForm.cmdCancel.focus();		

		}
	}

	function SelectLink() {
	
		var selRowId = $("#ssOleDBGridLinkRecords").jqGrid('getGridParam', 'selrow');
		if (selRowId > 0) {
			$("#optionframe").dialog("close");
			$("#optionframe").html();

			var frmLinkFindForm = document.getElementById('frmLinkFindForm');
			var recordID = $("#ssOleDBGridLinkRecords").jqGrid('getCell', selRowId, 'ID');

			var postData = {
				Action: optionActionType.SELECTLINK,
				LinkRecordID: recordID,
				ScreenID: frmLinkFindForm.txtOptionScreenID.value,
				LinkTableID: frmLinkFindForm.txtOptionLinkTableID.value,
				<%:Html.AntiForgeryTokenForAjaxPost() %> };
			OpenHR.submitForm(null, "optionframe", null, postData, "linkFind_submit");

		}
	}

	function CancelLink() {		
		$("#optionframe").dialog("close");
		$("#optionframe").html();

		var postData = {
			Action: optionActionType.CANCEL,
			<%:Html.AntiForgeryTokenForAjaxPost() %> };
		OpenHR.submitForm(null, "optionframe", null, postData, "linkFind_submit");

	}
	/* Sequential search the grid for the required ID. */
	function locateRecord(psSearchFor, pfIDMatch) {
		var fFound;
		var iIndex;
		var iIDColumnIndex;
		var sColumnName;
		var frmLinkFindForm = document.getElementById('frmLinkFindForm');

		fFound = false;

		frmLinkFindForm.ssOleDBGridLinkRecords.redraw = false;

		if (pfIDMatch == true) {
			// Locate the ID column in the grid.
			iIDColumnIndex = -1;
			for (iIndex = 0; iIndex < frmLinkFindForm.ssOleDBGridLinkRecords.Cols; iIndex++) {
				sColumnName = frmLinkFindForm.ssOleDBGridLinkRecords.Columns(iIndex).Name;
				if (sColumnName.toUpperCase() == "ID") {
					iIDColumnIndex = iIndex;
					break;
				}
			}

			if (iIDColumnIndex >= 0) {
				frmLinkFindForm.ssOleDBGridLinkRecords.MoveLast();
				frmLinkFindForm.ssOleDBGridLinkRecords.MoveFirst();

				for (iIndex = 1; iIndex <= frmLinkFindForm.ssOleDBGridLinkRecords.rows; iIndex++) {
					if (frmLinkFindForm.ssOleDBGridLinkRecords.Columns(iIDColumnIndex).value == psSearchFor) {
						frmLinkFindForm.ssOleDBGridLinkRecords.SelBookmarks.Add(frmLinkFindForm.ssOleDBGridLinkRecords.Bookmark);
						fFound = true;
						break;
					}

					if (iIndex < frmLinkFindForm.ssOleDBGridLinkRecords.rows) {
						frmLinkFindForm.ssOleDBGridLinkRecords.MoveNext();
					}
					else {
						break;
					}
				}
			}
		}
		else {
			for (iIndex = 1; iIndex <= frmLinkFindForm.ssOleDBGridLinkRecords.rows; iIndex++) {
				var sGridValue = new String(frmLinkFindForm.ssOleDBGridLinkRecords.Columns(0).value);
				sGridValue = sGridValue.substr(0, psSearchFor.length).toUpperCase();
				if (sGridValue == psSearchFor.toUpperCase()) {
					frmLinkFindForm.ssOleDBGridLinkRecords.SelBookmarks.Add(frmLinkFindForm.ssOleDBGridLinkRecords.Bookmark);
					fFound = true;
					break;
				}

				if (iIndex < frmLinkFindForm.ssOleDBGridLinkRecords.rows) {
					frmLinkFindForm.ssOleDBGridLinkRecords.MoveNext();
				}
				else {
					break;
				}
			}
		}

		if ((fFound == false) && (frmLinkFindForm.ssOleDBGridLinkRecords.rows > 0)) {
			// Select the top row.
			frmLinkFindForm.ssOleDBGridLinkRecords.MoveFirst();
			frmLinkFindForm.ssOleDBGridLinkRecords.SelBookmarks.Add(frmLinkFindForm.ssOleDBGridLinkRecords.Bookmark);
		}

		frmLinkFindForm.ssOleDBGridLinkRecords.redraw = true;
	}
	
	function linkFind_refreshControls() {
		//linkFind...
		var frmLinkFindForm = document.getElementById("frmLinkFindForm");

		var selRowId = $("#ssOleDBGridLinkRecords").jqGrid('getGridParam', 'selrow');
		if (selRowId > 0) {
			button_disable(frmLinkFindForm.cmdSelectLink, false);
		}
		else {
			button_disable(frmLinkFindForm.cmdSelectLink, true);
		}

		if (frmLinkFindForm.selectOrder.length <= 1) {
			combo_disable(frmLinkFindForm.selectOrder, true);
			button_disable(frmLinkFindForm.btnGoLinkOrder, true);
		}

		if (frmLinkFindForm.selectView.length <= 1) {
			combo_disable(frmLinkFindForm.selectView, true);
			button_disable(frmLinkFindForm.btnGoLinkView, true);
		}
	}
	
	function goLinkView() {
		//need this as this grid won't accept live changes :/		
		$("#ssOleDBGridLinkRecords").jqGrid('GridUnload');
		// Get the optionData.asp to get the link find records.
		var frmLinkFindForm = document.getElementById('frmLinkFindForm');
		var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
		optionDataForm.txtOptionAction.value = "LOADFIND";
		optionDataForm.txtOptionTableID.value = frmLinkFindForm.txtOptionLinkTableID.value;
		optionDataForm.txtOptionViewID.value = frmLinkFindForm.selectView.options[frmLinkFindForm.selectView.selectedIndex].value;
		optionDataForm.txtOptionOrderID.value = frmLinkFindForm.selectOrder.options[frmLinkFindForm.selectOrder.selectedIndex].value;
		optionDataForm.txtOptionFirstRecPos.value = 1;
		optionDataForm.txtOptionCurrentRecCount.value = 0;

		frmLinkFindForm.txtOptionLinkViewID.value = optionDataForm.txtOptionViewID.value;
		frmLinkFindForm.txtOptionLinkOrderID.value = optionDataForm.txtOptionOrderID.value;

		refreshOptionData();
	}

	function goLinkOrder() {
		// Get the optionData.asp to get the link find records.

		//need this as this grid won't accept live changes :/		
		$("#ssOleDBGridLinkRecords").jqGrid('GridUnload');

		var frmLinkFindForm = document.getElementById('frmLinkFindForm');
		var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");

		optionDataForm.txtOptionAction.value = "LOADFIND";
		optionDataForm.txtOptionTableID.value = frmLinkFindForm.txtOptionLinkTableID.value;
		optionDataForm.txtOptionViewID.value = frmLinkFindForm.selectView.options[frmLinkFindForm.selectView.selectedIndex].value;
		optionDataForm.txtOptionOrderID.value = frmLinkFindForm.selectOrder.options[frmLinkFindForm.selectOrder.selectedIndex].value;
		optionDataForm.txtOptionFirstRecPos.value = 1;
		optionDataForm.txtOptionCurrentRecCount.value = 0;

		frmLinkFindForm.txtOptionLinkViewID.value = optionDataForm.txtOptionViewID.value;
		frmLinkFindForm.txtOptionLinkOrderID.value = optionDataForm.txtOptionOrderID.value;

		refreshOptionData();
	}

	function selectedOrderID() {
		var frmLinkFindForm = document.getElementById('frmLinkFindForm');
		return frmLinkFindForm.selectOrder.options[frmLinkFindForm.selectOrder.selectedIndex].value;
	}

	function selectedViewID() {
		var frmLinkFindForm = document.getElementById('frmLinkFindForm');
		return frmLinkFindForm.selectView.options[frmLinkFindForm.selectView.selectedIndex].value;
	}

	function doViewHelp() {
		var helpText = "The 'View' defines the subset of data from the table that is displayed in the grid." +
			"The name of the view should give an indication of which data is included in the subset.";
		OpenHR.messageBox(helpText, 48, "Information");
	}

	function doOrderHelp() {
		var helpText = "The 'Order' defines which columns are displayed in the grid, and the order in which the data is listed.";
		OpenHR.messageBox(helpText, 48, "Information");
	}

</script>

<script src="<%: Url.LatestContent("~/Scripts/ctl_SetStyles.js")%>" type="text/javascript"></script>

<div id="divLinkFindForm" <%=session("BodyTag")%>>
	<form action="" method="POST" id="frmLinkFindForm" name="frmLinkFindForm">
		<div class="" style="">
			<div class="pageTitleDiv" style="margin-bottom: 15px">
				<span class="pageTitle" id="PopupReportDefinition_PageTitle">Find Link Record</span>
			</div>

			<table  class="invisible" style="border-spacing: 0;">
				<tr>
					<td style="height:10px">
						<table class="width100 invisible" style="border-spacing: 0;">
							<tr>
								<td>View :</td>
								<td>
									<select id="selectView" name="selectView" class="width100" style="">
										<%
											Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
											Dim sErrorDescription = ""
											Dim sFailureDescription = ""
											Dim prmDfltOrderID As New SqlParameter("plngDfltOrderID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
											Dim rstViewRecords = objDatabase.DB.GetDataTable("sp_ASRIntGetLinkViews", CommandType.StoredProcedure _
													, New SqlParameter("plngTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionLinkTableID"))} _
													, prmDfltOrderID)

											If (Len(sErrorDescription) = 0) And (Len(sFailureDescription) = 0) Then
												For Each objRow As DataRow In rstViewRecords.Rows
													Response.Write("						<option value=" & objRow(0))
													If CInt(objRow(0)) = CInt(Session("optionLinkViewID")) Then
														Response.Write(" SELECTED")
													End If

													If objRow(0) = 0 Then
														Response.Write(">" & Replace(objRow(1).ToString(), "_", " ") & "</option>" & vbCrLf)
													Else
														Response.Write(">'" & Replace(objRow(1).ToString, "_", " ") & "' view</option>" & vbCrLf)
													End If
												Next
	
												If Session("optionLinkOrderID") <= 0 Then
													Session("optionLinkOrderID") = prmDfltOrderID.Value
												End If
											End If
										%>
									</select>
								</td>

								<td id="tdTViewHelp" name="tdTViewHelp" onclick="doViewHelp()" class="nowrap">
									<img id="imgTViewHelp" name="imgTViewHelp" alt="help"
										src="<%=Url.Content("~/Content/images/Help32.png")%>"
										title="What happens if I change the view?" style="vertical-align: middle;width: 17px; height: 17px; border: 0; cursor: pointer" />
								</td>

								<td>
									<input class="btn" id="btnGoLinkView" name="btnGoLinkView" onclick="goLinkView()" type="button" value="Go" />
								</td>

								<td style="text-align: right;">Order :</td>

								<td>
									<select id="selectOrder" name="selectOrder" class="width100" style="">
										<%
											If (Len(sErrorDescription) = 0) And (Len(sFailureDescription) = 0) Then
														
												Dim rstOrderRecords = objDatabase.GetTableOrders(CleanNumeric(Session("optionLinkTableID")), 0)
												For Each objRow As DataRow In rstOrderRecords.Rows
													Response.Write("						<option value=" & objRow(1))
													If objRow(1) = CInt(Session("optionLinkOrderID")) Then
														Response.Write(" SELECTED")
													End If
													Response.Write(">" & Replace(objRow(0).ToString(), "_", " ") & "</option>" & vbCrLf)
												Next

											End If
										%>
									</select>
								</td>

								<td id="tdTOrderHelp" name="tdTOrderHelp" onclick="doOrderHelp()" class="nowrap">
									<img id="imgTOrderHelp" name="imgTOrderHelp" alt="help"
										src="<%=Url.Content("~/Content/images/Help32.png")%>"
										title="What happens if I change the order?" style="vertical-align: middle; width: 17px; height: 17px; border: 0; cursor: pointer" />
								</td>
								<td>
									<input class="btn" id="btnGoLinkOrder" name="btnGoLinkOrder" onclick="goLinkOrder()" type="button" value="Go" />
								</td>
							</tr>
						</table>
					</td>
				</tr>

				<tr>
					<td>
						<%--<div id="linkFindGridRow" style="height: 75%; margin-bottom: 50px;">--%>
						<table id="ssOleDBGridLinkRecords" name="ssOleDBGridLinkRecords" class="width100 height100"></table>
						<%--</div>--%>
					</td>
				</tr>
			</table>

			<%
				Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtFailureDescription name=txtFailureDescription value=""" & sFailureDescription & """>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtOptionScreenID name=txtOptionScreenID value=" & Session("optionScreenID") & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtOptionLinkTableID name=txtOptionLinkTableID value=" & Session("optionLinkTableID") & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtOptionLinkViewID name=txtOptionLinkViewID value=" & Session("optionLinkViewID") & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtOptionLinkOrderID name=txtOptionLinkOrderID value=" & Session("optionLinkOrderID") & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtOptionLinkRecordID name=txtOptionLinkRecordID value=" & Session("optionLinkRecordID") & ">" & vbCrLf)
			%>
			<div id="divLinkFindButtons">
				<input class="btn" id="cmdSelectLink" name="cmdSelectLink" onclick="SelectLink()" type="button" value="Select" />
				<input class="btn" id="cmdCancel" name="cmdCancel" onclick="CancelLink()" type="button" value="Cancel" />
			</div>
		</div>
	</form>
	<input type="hidden" id="txtTicker" name="txtTicker" value="0">
	<input type="hidden" id="txtLastKeyFind" name="txtLastKeyFind" value="">

</div>

<script type="text/javascript">
	linkFind_window_onload();
	$('table').attr('border', '0');
</script>
