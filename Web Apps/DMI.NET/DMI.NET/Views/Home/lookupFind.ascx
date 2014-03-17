<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<script src="<%: Url.LatestContent("~/Scripts/ctl_SetFont.js")%>" type="text/javascript"></script>

<script type="text/javascript">

	function lookupFind_removeAll(jqGridID) {
		//remove all rows from the jqGrid.
		$('#' + jqGridID).jqGrid('clearGridData');
	}

	function lookupFind_window_onload() {
		var fOK;
		fOK = true;
		var frmLookupFindForm = document.getElementById("frmLookupFindForm");
		var sErrMsg = frmLookupFindForm.txtErrorDescription.value;
		if (sErrMsg.length > 0) {
			fOK = false;
			OpenHR.messageBox(sErrMsg);
			window.parent.location.replace("login");
		}

		if (fOK == true) {
			sErrMsg = frmLookupFindForm.txtFailureDescription.value;
			if (sErrMsg.length > 0) {
				fOK = false;
				OpenHR.messageBox(sErrMsg);
				window.parent.location.replace("login");
			}
		}

		if (fOK == true) {
			setGridFont(frmLookupFindForm.ssOleDBGrid);
			// Expand the option frame and hide the work frame.
			//window.parent.document.all.item("workframeset").cols = "0, *";	
			$("#optionframe").attr("data-framesource", "LOOKUPFIND");
			//$("#workframe").hide();
			//$("#optionframe").show();
			$("#optionframe").dialog({
				autoOpen: true,
				modal: true,
				width: 750,
				height: 800,
				close: function() {
					CancelLookup();
				}
			});

			// Set focus onto one of the form controls. 
			// NB. This needs to be done before making any reference to the grid
			frmLookupFindForm.cmdCancel.focus();

			// Fault 3503
			//todo: window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "hidden";

			// Get the optionData.asp to get the link find records.
			var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
			optionDataForm.txtOptionAction.value = "LOADLOOKUPFIND";
			optionDataForm.txtOptionColumnID.value = frmLookupFindForm.txtOptionColumnID.value;
			optionDataForm.txtOptionLookupColumnID.value = frmLookupFindForm.txtOptionLookupColumnID.value;
			optionDataForm.txtOptionLookupFilterValue.value = frmLookupFindForm.txtOptionLookupFilterValue.value;
			optionDataForm.txtOptionPageAction.value = "LOAD";
			optionDataForm.txtOptionFirstRecPos.value = 1;
			optionDataForm.txtOptionCurrentRecCount.value = 0;
			optionDataForm.txtOptionIsLookupTable.value = frmLookupFindForm.txtIsLookupTable.value;
			optionDataForm.txtOptionRecordID.value = OpenHR.getForm("workframe", "frmRecordEditForm").txtCurrentRecordID.value;

			optionDataForm.txtOptionParentTableID.value = OpenHR.getForm("workframe", "frmRecordEditForm").txtCurrentParentTableID.value;
			optionDataForm.txtOptionParentRecordID.value = OpenHR.getForm("workframe", "frmRecordEditForm").txtCurrentParentRecordID.value;

			if (frmLookupFindForm.txtIsLookupTable.value == "False") {
				optionDataForm.txtOptionTableID.value = frmLookupFindForm.txtOptionLinkTableID.value;
				optionDataForm.txtOptionViewID.value = frmLookupFindForm.selectView.options[frmLookupFindForm.selectView.selectedIndex].value;
				optionDataForm.txtOptionOrderID.value = frmLookupFindForm.selectOrder.options[frmLookupFindForm.selectOrder.selectedIndex].value;
			}
			else {
				optionDataForm.txtOptionTableID.value = 0;
				optionDataForm.txtOptionViewID.value = 0;
				optionDataForm.txtOptionOrderID.value = 0;
			}

			refreshOptionData();	//should be in scope now...
		}
	}

	function SelectLookup() {		
		var frmLookupFindForm = document.getElementById("frmLookupFindForm");
		var frmGotoOption = document.getElementById("frmGotoOption");

		var selRowId = $("#ssOleDBGrid").jqGrid('getGridParam', 'selrow');
		if (selRowId > 0) {
			//$("#optionframe").hide();
			//$("#workframe").show();
			$("#optionframe").dialog("destroy");

			frmGotoOption.txtGotoOptionColumnID.value = frmLookupFindForm.txtOptionColumnID.value;
			frmGotoOption.txtGotoOptionLookupColumnID.value = frmLookupFindForm.txtOptionLookupColumnID.value;
			frmGotoOption.txtGotoOptionLookupValue.value = selectedValue(selRowId);
			frmGotoOption.txtGotoOptionAction.value = "SELECTLOOKUP";
			frmGotoOption.txtGotoOptionPage.value = "emptyoption";
			OpenHR.submitForm(frmGotoOption);
		}
	}

	function ClearLookup() {
		var frmLookupFindForm = document.getElementById("frmLookupFindForm");
		var frmGotoOption = document.getElementById("frmGotoOption");
		// Fault 3503
		//TODO: ?window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "visible";
		//$("#optionframe").hide();
		//$("#workframe").show();
		$("#optionframe").dialog("destroy");

		frmGotoOption.txtGotoOptionColumnID.value = frmLookupFindForm.txtOptionColumnID.value;
		frmGotoOption.txtGotoOptionLookupColumnID.value = frmLookupFindForm.txtOptionLookupColumnID.value;
		frmGotoOption.txtGotoOptionLookupValue.value = "";
		frmGotoOption.txtGotoOptionAction.value = "SELECTLOOKUP";
		frmGotoOption.txtGotoOptionPage.value = "emptyoption";
		OpenHR.submitForm(frmGotoOption);
	}

	function CancelLookup() {
		// Fault 3503
		//TODO: ?window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "visible";
		//$("#optionframe").hide();
		//$("#workframe").show();
		$("#optionframe").dialog("destroy");

		var frmGotoOption = document.getElementById("frmGotoOption");

		frmGotoOption.txtGotoOptionAction.value = "CANCEL";
		frmGotoOption.txtGotoOptionPage.value = "emptyoption";
		OpenHR.submitForm(frmGotoOption);
	}

	/* Return the value of the record selected in the find form. */
	function selectedValue(selRowId) {
		var frmLookupFindForm = document.getElementById("frmLookupFindForm");
		var sValue;

		if (frmLookupFindForm.txtIsLookupTable.value == "False") {
			//sValue = frmLookupFindForm.ssOleDBGrid.Columns(parseInt(frmLookupFindForm.txtLookupColumnGridPosition.value)).value;
			var cellNumber = parseInt(frmLookupFindForm.txtLookupColumnGridPosition.value);
			sValue = $("#ssOleDBGrid").jqGrid('getCell', selRowId, cellNumber);
		}
		else {
			//sValue = frmLookupFindForm.ssOleDBGrid.Columns(0).Value;
			sValue = $("#ssOleDBGrid").jqGrid('getCell', selRowId, 0);
		}

		return (sValue);
	}

	/* Sequential search the grid for the required OLE. */
	function locateRecord(psFileName, pfExactMatch) {
		//select the grid row that contains the record with the passed in ID.
		var rowNumber = $("#ssOleDBGrid input[value='" + psFileName + "']").parent().parent().attr("id");
		if (rowNumber >= 0) {
			$("#ssOleDBGrid").jqGrid('setSelection', rowNumber);
		} else {
			$("#ssOleDBGrid").jqGrid('setSelection', 1);
		}
	}

	function lookupFind_refreshControls() {
		//lookupFind...
		var frmLookupFindForm = document.getElementById("frmLookupFindForm");

		var selRowId = $("#ssOleDBGrid").jqGrid('getGridParam', 'selrow');
		if (selRowId > 0) {
			button_disable(frmLookupFindForm.cmdSelectLookup, false);
		}
		else {
			button_disable(frmLookupFindForm.cmdSelectLookup, true);
		}

		if (frmLookupFindForm.txtOptionLookupMandatory.value == "true") {
			button_disable(frmLookupFindForm.cmdClearLookup, true);
		}
		else {
			button_disable(frmLookupFindForm.cmdClearLookup, false);
		}
	}

	function lookupFind_rowCount() {
		return $("#ssOleDBGrid tr").length - 1;
	}

	function lookupFind_moveFirst() {
		$("#ssOleDBGrid").jqGrid('setSelection', 1);
	}

	function goView() {

		//need this as this grid won't accept live changes :/		
		$("#ssOleDBGrid").jqGrid('GridUnload');

		var frmLookupFindForm = document.getElementById("frmLookupFindForm");

		var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
		optionDataForm.txtOptionAction.value = "LOADLOOKUPFIND";
		optionDataForm.txtOptionColumnID.value = frmLookupFindForm.txtOptionColumnID.value;
		optionDataForm.txtOptionLookupColumnID.value = frmLookupFindForm.txtOptionLookupColumnID.value;
		optionDataForm.txtOptionLookupFilterValue.value = frmLookupFindForm.txtOptionLookupFilterValue.value;
		optionDataForm.txtOptionPageAction.value = "LOAD";
		optionDataForm.txtOptionFirstRecPos.value = 1;
		optionDataForm.txtOptionCurrentRecCount.value = 0;
		optionDataForm.txtOptionIsLookupTable.value = frmLookupFindForm.txtIsLookupTable.value;
		optionDataForm.txtOptionRecordID.value = OpenHR.getForm("workframe", "frmRecordEditForm").txtCurrentRecordID.value;

		optionDataForm.txtOptionParentTableID.value = OpenHR.getForm("workframe", "frmRecordEditForm").txtCurrentParentTableID.value;
		optionDataForm.txtOptionParentRecordID.value = OpenHR.getForm("workframe", "frmRecordEditForm").txtCurrentParentRecordID.value;

		if (frmLookupFindForm.txtIsLookupTable.value == "False") {
			optionDataForm.txtOptionTableID.value = frmLookupFindForm.txtOptionLinkTableID.value;
			optionDataForm.txtOptionViewID.value = frmLookupFindForm.selectView.options[frmLookupFindForm.selectView.selectedIndex].value;
			optionDataForm.txtOptionOrderID.value = frmLookupFindForm.selectOrder.options[frmLookupFindForm.selectOrder.selectedIndex].value;
		}
		else {
			optionDataForm.txtOptionTableID.value = 0;
			optionDataForm.txtOptionViewID.value = 0;
			optionDataForm.txtOptionOrderID.value = 0;
		}

		refreshOptionData();	//should be in scope...
	}

	function goOrder() {
		//need this as this grid won't accept live changes :/		
		$("#ssOleDBGrid").jqGrid('GridUnload');

		var frmLookupFindForm = document.getElementById("frmLookupFindForm");

		var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
		optionDataForm.txtOptionAction.value = "LOADLOOKUPFIND";
		optionDataForm.txtOptionColumnID.value = frmLookupFindForm.txtOptionColumnID.value;
		optionDataForm.txtOptionLookupColumnID.value = frmLookupFindForm.txtOptionLookupColumnID.value;
		optionDataForm.txtOptionLookupFilterValue.value = frmLookupFindForm.txtOptionLookupFilterValue.value;
		optionDataForm.txtOptionPageAction.value = "LOAD";
		optionDataForm.txtOptionFirstRecPos.value = 1;
		optionDataForm.txtOptionCurrentRecCount.value = 0;
		optionDataForm.txtOptionIsLookupTable.value = frmLookupFindForm.txtIsLookupTable.value;
		optionDataForm.txtOptionRecordID.value = OpenHR.getForm("workframe", "frmRecordEditForm").txtCurrentRecordID.value;

		optionDataForm.txtOptionParentTableID.value = OpenHR.getForm("workframe", "frmRecordEditForm").txtCurrentParentTableID.value;
		optionDataForm.txtOptionParentRecordID.value = OpenHR.getForm("workframe", "frmRecordEditForm").txtCurrentParentRecordID.value;

		if (frmLookupFindForm.txtIsLookupTable.value == "False") {
			optionDataForm.txtOptionTableID.value = frmLookupFindForm.txtOptionLinkTableID.value;
			optionDataForm.txtOptionViewID.value = frmLookupFindForm.selectView.options[frmLookupFindForm.selectView.selectedIndex].value;
			optionDataForm.txtOptionOrderID.value = frmLookupFindForm.selectOrder.options[frmLookupFindForm.selectOrder.selectedIndex].value;
		}
		else {
			optionDataForm.txtOptionTableID.value = 0;
			optionDataForm.txtOptionViewID.value = 0;
			optionDataForm.txtOptionOrderID.value = 0;
		}

		refreshOptionData();	//should be in scope...
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

	function ssOleDBGrid_dblClick() {
		SelectLookup();
	}
</script>

<script src="<%: Url.LatestContent("~/Scripts/ctl_SetStyles.js")%>" type="text/javascript"></script>

<div <%=session("BodyTag")%>>
	<form action="" method="POST" id="frmLookupFindForm" name="frmLookupFindForm">
		<%
			
			Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
			Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

			Dim objTable = objDatabase.GetTableFromColumnID(CInt(Session("optionLookupColumnID")))
			Dim fIsLookupTable = (objTable.TableType = TableTypes.tabLookup)
			Dim lngLookupTableID = objTable.ID

			Dim sErrorDescription = ""
			Dim sFailureDescription = ""
	
		%>
		<div id="divFindForm" <%=session("BodyTag")%>>
			<div class="absolutefull">
				<div id="row1" style="margin-left: 20px;margin-right: 20px">
					<h3 class="pageTitle" align="left">Find Lookup Record</h3>
				</div>

				<%If Not fIsLookupTable Then%>
				<div id="row1a">
					<table align="center" class="outline" cellpadding="5" cellspacing="0" width="100%" height="100%">
						<tr>
							<td>
								<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
									<tr>
										<td height="10">&nbsp;&nbsp;</td>
										<td height="10">
											<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
												<tr>
													<td width="5%">View : </td>
													<td width="20%">
														<select id="selectView" name="selectView" class="combo" style="HEIGHT: 22px; WIDTH: 200px">
															<%

																' Get the view records.
																If (Len(sErrorDescription) = 0) And (Len(sFailureDescription) = 0) Then

																	Try

																		Dim prmDfltOrderID As New SqlParameter("plngDfltOrderID", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
																		Dim rstViewRecords = objDataAccess.GetDataTable("spASRIntGetLookupViews", CommandType.StoredProcedure _
																			, New SqlParameter("plngTableID", SqlDbType.Int) With {.Value = lngLookupTableID} _
																			, prmDfltOrderID _
																			, New SqlParameter("plngColumnID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionColumnID"))})
																				
																		For Each objRow As DataRow In rstViewRecords.Rows
																			Response.Write("						<option value=" & objRow(0))
																			If objRow(0) = CLng(Session("optionLinkViewID")) Then
																				Response.Write(" SELECTED")
																			End If

																			If objRow(0) = 0 Then
																				Response.Write(">" & Replace(objRow(1).ToString(), "_", " ") & "</option>" & vbCrLf)
																			Else
																				Response.Write(">'" & Replace(objRow(1).ToString(), "_", " ") & "' view</option>" & vbCrLf)
																			End If
																	
																		Next

																		If Session("optionLookupOrderID") <= 0 Then
																			Session("optionLookupOrderID") = prmDfltOrderID.Value
																		End If
																		
																	Catch ex As Exception
																		sErrorDescription = "The lookup view records could not be retrieved." & vbCrLf & FormatError(ex.Message)

																	End Try
		
																End If
															%>
														</select>
													</td>
																						
													<td width="10%" id="tdTViewHelp" name="tdTViewHelp" onclick="doViewHelp()" style="white-space: nowrap; text-align:center;" >
															
														<img id="imgTViewHelp" name="imgTViewHelp" alt="help"
														     src="<%=Url.Content("~/Content/images/Help32.png")%>"
														     
														     title="What happens if I change the view?" style="width:17px; height:17px; border:0; cursor: pointer"/>
															
													</td>
													<td  width="5%"></td>
													<td width="5%">
														<input type="button" value="Go" class="btn" id="btnGoView" name="btnGoView"
															onclick="goView()" />
													</td>
													<td  width="5%">
													</td>
													<td style="text-align: right; width:5%">Order :
													</td>
													<td width="5%">
													</td>
													<td width="20%">
														<select id="selectOrder" name="selectOrder" class="combo" style="HEIGHT: 22px; WIDTH: 200px">
															<%
																
																' Get the order records.																
																If (Len(sErrorDescription) = 0) And (Len(sFailureDescription) = 0) Then

																	Dim rstOrderRecords = objDatabase.GetTableOrders(lngLookupTableID, 0)																
																	For Each objRow As DataRow In rstOrderRecords.Rows
																		Response.Write("						<option value=" & objRow(1))
																		If objRow(1) = CInt(Session("optionOrderID")) Then
																			Response.Write(" SELECTED")
																		End If
																		Response.Write(">" & Replace(objRow(0).ToString(), "_", " ") & "</option>" & vbCrLf)
																	Next

																End If
															%>
														</select>
													</td>
																			
													<td width="10%" id="tdTOrderHelp" name="tdTOrderHelp" onclick="doOrderHelp()" style="white-space: nowrap; text-align:center;"  >
															
														<img id="imgTOrderHelp" name="imgTOrderHelp" alt="help"
														     src="<%=Url.Content("~/Content/images/Help32.png")%>"														     
														     title="What happens if I change the order?" style="width:17px; height:17px; border:0; cursor: pointer"/>
														
													</td>
													<td  width="5%"></td>
													<td width="10%">
														<input type="button" value="Go" class="btn" id="btnGoOrder" name="btnGoOrder"
															onclick="goOrder()" />
													</td>
												</tr>
											</table>
										</td>
										<td height="10">&nbsp;&nbsp;</td>
									</tr>
									<tr>
										<td height="10" colspan="3"></td>
									</tr>
								</table>
							</td>
						</tr>
					</table>
				</div>
				<%End If 'if fIsLookupTable then%>
				<div id="lookupFindGridRow" style="height: <%If fIsLookupTable Then%>60%<%Else%>55%<%End If%>; margin-bottom: 150px;">
					<table class="outline" style="width: 100%;" id="ssOleDBGrid"></table>
					<div id="ssOLEDBPager"></div>					
				</div>
				<div id="row3">
					<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
						<tr>
							<td colspan="6"></td>
						</tr>
						<tr>
							<td></td>
							<td width="10">
								<input id="cmdSelectLookup" name="cmdSelectLookup" type="button" value="Select" style="WIDTH: 75px" width="75" class="btn"
									onclick="SelectLookup()" />
							</td>
							<td width="40"></td>
							<td width="10">
								<input id="cmdClearLookup" name="cmdClearLookup" type="button" value="Clear" style="WIDTH: 75px" width="75" class="btn"
									onclick="ClearLookup()" />
							</td>
							<td width="40"></td>
							<td width="10">
								<input id="cmdCancel" name="cmdCancel" type="button" value="Cancel" style="WIDTH: 75px" width="75" class="btn"
									onclick="CancelLookup()" />
							</td>
						</tr>
					</table>
				</div>
			</div>
		</div>


		<input type='hidden' id="txtErrorDescription" name="txtErrorDescription" value="<%=sErrorDescription%>">
		<input type='hidden' id="txtFailureDescription" name="txtFailureDescription" value="<%=sFailureDescription%>">
		<input type='hidden' id="txtOptionColumnID" name="txtOptionColumnID" value='<%=session("optionColumnID")%>'>
		<input type='hidden' id="txtOptionLookupColumnID" name="txtOptionLookupColumnID" value='<%=session("optionLookupColumnID")%>'>
		<input type='hidden' id="txtOptionLookupMandatory" name="txtOptionLookupMandatory" value='<%=session("optionLookupMandatory")%>'>
		<input type='hidden' id="txtOptionLookupValue" name="txtOptionLookupValue" value='<%=session("optionLookupValue")%>'>
		<input type='hidden' id="txtOptionLookupFilterValue" name="txtOptionLookupFilterValue" value="<%=session("optionLookupFilterValue")%>">
		<input type='hidden' id="txtIsLookupTable" name="txtIsLookupTable" value="<%=fIsLookupTable%>">
		<input type='hidden' id="txtOptionLinkTableID" name="txtOptionLinkTableID" value="<%=lngLookupTableID%>">
		<input type='hidden' id="txtLookupColumnGridPosition" name="txtLookupColumnGridPosition" value="0">
	</form>

	<input type='hidden' id="txtTicker" name="txtTicker" value="0">
	<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">

	<form action="lookupFind_Submit" method="post" id="frmGotoOption" name="frmGotoOption">
		<%Html.RenderPartial("~/Views/Shared/gotoOption.ascx")%>
	</form>

</div>

<script type="text/javascript">
	lookupFind_window_onload();
</script>
