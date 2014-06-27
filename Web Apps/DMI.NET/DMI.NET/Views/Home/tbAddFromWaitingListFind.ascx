<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<%Dim sErrorDescription = ""
	Dim sFailureDescription = ""%>

<script type="text/javascript">
	
	function tbAddFromWaitingListFind_onload() {		
		var fOK;
		fOK = true;
		
		var frmtbFindForm = document.getElementById("frmtbFindForm");

		var sErrMsg = frmtbFindForm.txtErrorDescription.value;
		if (sErrMsg.length > 0) {
			fOK = false;
			OpenHR.messageBox(sErrMsg);
			window.parent.location.replace("login.asp");
		}

		if (fOK == true) {
			sErrMsg = frmtbFindForm.txtFailureDescription.value;
			if (sErrMsg.length > 0) {
				fOK = false;
				OpenHR.messageBox(sErrMsg);
				tbCancel();
			}
		}

		if (fOK == true) {
			if (frmtbFindForm.selectView.length == 0) {
				fOK = false;
				OpenHR.messageBox("You do not have permission to read the course table.");
				tbCancel();
			}
		}

		if (fOK == true) {
			if (frmtbFindForm.selectOrder.length == 0) {
				fOK = false;
				OpenHR.messageBox("You do not have permission to use any of the course table orders.");
				tbCancel();
			}
		}

		if (fOK == true) {

			// Expand the option frame and hide the work frame.
			//window.parent.document.all.item("workframeset").cols = "0, *";	
			$("#optionframe").attr("data-framesource", "TBADDFROMWAITINGLISTFIND");
			$("#workframe").hide();
			$("#optionframe").show();

			// Set focus onto one of the form controls. 
			// NB. This needs to be done before making any reference to the grid
			frmtbFindForm.cmdCancel.focus();

			//TODO: window.parent.frames("workframe").document.forms("frmtbFindForm").ssOleDBGridFindRecords.style.visibility = "hidden";

			// Get the optionData.asp to get the link find records.
			var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
			optionDataForm.txtOptionAction.value = "LOADADDFROMWAITINGLIST";
			optionDataForm.txtOptionTableID.value = frmtbFindForm.txtOptionLinkTableID.value;
			optionDataForm.txtOptionViewID.value = frmtbFindForm.selectView.options[frmtbFindForm.selectView.selectedIndex].value;
			optionDataForm.txtOptionOrderID.value = frmtbFindForm.selectOrder.options[frmtbFindForm.selectOrder.selectedIndex].value;
			optionDataForm.txtOptionRecordID.value = frmtbFindForm.txtOptionRecordID.value;
			optionDataForm.txtOptionFirstRecPos.value = 1;
			optionDataForm.txtOptionCurrentRecCount.value = 0;
			optionDataForm.txtOptionPageAction.value = "LOAD";

			frmtbFindForm.txtOptionLinkViewID.value = optionDataForm.txtOptionViewID.value;
			frmtbFindForm.txtOptionLinkOrderID.value = optionDataForm.txtOptionOrderID.value;

			refreshOptionData();	//should be in scope.
		}
	}
</script>

<script type="text/javascript">

	function tbSelect() {		
		var frmGotoOption = document.getElementById("frmGotoOption");
		var frmtbFindForm = document.getElementById("frmtbFindForm");

		if ($("#txtStatusPExists").val() != "True") {
			//TODO: window.parent.frames("workframe").document.forms("frmtbFindForm").ssOleDBGridFindRecords.style.visibility = "visible";
		}

		frmGotoOption.txtGotoOptionAction.value = "SELECTADDFROMWAITINGLIST_1";
		frmGotoOption.txtGotoOptionRecordID.value = frmtbFindForm.txtOptionRecordID.value;
		var selRowId = $("#ssOleDBGridRecords").jqGrid('getGridParam', 'selrow');
		var recordID = $("#ssOleDBGridRecords").jqGrid('getCell', selRowId, 'ID');
		frmGotoOption.txtGotoOptionLinkRecordID.value = recordID;	// ssselectedRecordID();
		frmGotoOption.txtGotoOptionPage.value = "emptyoption";
		OpenHR.submitForm(frmGotoOption);
	}

	function tbCancel() {
		var frmGotoOption = document.getElementById("frmGotoOption");

		//window.parent.frames("workframe").document.forms("frmtbFindForm").ssOleDBGridFindRecords.style.visibility = "visible";
		$("#optionframe").hide();
		$("#workframe").show();


		frmGotoOption.txtGotoOptionAction.value = "CANCEL";
		frmGotoOption.txtGotoOptionLinkRecordID.value = 0;
		frmGotoOption.txtGotoOptionPage.value = "emptyoption";
		OpenHR.submitForm(frmGotoOption);
	}

	function tbrefreshControls() {
		var frmtbFindForm = document.getElementById("frmtbFindForm");

		var selRowId = $("#ssOleDBGridRecords").jqGrid('getGridParam', 'selrow');

		button_disable(frmtbFindForm.cmdSelect, (selRowId == null));

		if (frmtbFindForm.selectOrder.length <= 1) {
			combo_disable(frmtbFindForm.selectOrder, true);
			button_disable(frmtbFindForm.btnGoOrder, true);
		}

		if (frmtbFindForm.selectView.length <= 1) {
			combo_disable(frmtbFindForm.selectView, true);
			button_disable(frmtbFindForm.btnGoView, true);
		}
	}

	function goView() {
		// Get the optionData.asp to get the link find records.
		var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
		var frmtbFindForm = document.getElementById("frmtbFindForm");
		optionDataForm.txtOptionAction.value = "LOADADDFROMWAITINGLIST";
		optionDataForm.txtOptionTableID.value = frmtbFindForm.txtOptionLinkTableID.value;
		optionDataForm.txtOptionViewID.value = frmtbFindForm.selectView.options[frmtbFindForm.selectView.selectedIndex].value;
		optionDataForm.txtOptionOrderID.value = frmtbFindForm.selectOrder.options[frmtbFindForm.selectOrder.selectedIndex].value;
		optionDataForm.txtOptionRecordID.value = frmtbFindForm.txtOptionRecordID.value;
		optionDataForm.txtOptionFirstRecPos.value = 1;
		optionDataForm.txtOptionCurrentRecCount.value = 0;

		frmtbFindForm.txtOptionLinkViewID.value = optionDataForm.txtOptionViewID.value;
		frmtbFindForm.txtOptionLinkOrderID.value = optionDataForm.txtOptionOrderID.value;

		refreshOptionData();	//should be in scope...
	}

	function goOrder() {
		// Get the optionData.asp to get the link find records.
		var frmtbFindForm = document.getElementById("frmtbFindForm");
		var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
		optionDataForm.txtOptionAction.value = "LOADADDFROMWAITINGLIST";
		optionDataForm.txtOptionTableID.value = frmtbFindForm.txtOptionLinkTableID.value;
		optionDataForm.txtOptionViewID.value = frmtbFindForm.selectView.options[frmtbFindForm.selectView.selectedIndex].value;
		optionDataForm.txtOptionOrderID.value = frmtbFindForm.selectOrder.options[frmtbFindForm.selectOrder.selectedIndex].value;
		optionDataForm.txtOptionRecordID.value = frmtbFindForm.txtOptionRecordID.value;
		optionDataForm.txtOptionFirstRecPos.value = 1;
		optionDataForm.txtOptionCurrentRecCount.value = 0;

		frmtbFindForm.txtOptionLinkViewID.value = optionDataForm.txtOptionViewID.value;
		frmtbFindForm.txtOptionLinkOrderID.value = optionDataForm.txtOptionOrderID.value;

		refreshOptionData();	//should be in scope.
	}

	function selectedOrderID() {
		var frmtbFindForm = document.getElementById("frmtbFindForm");
		return frmtbFindForm.selectOrder.options[frmtbFindForm.selectOrder.selectedIndex].value;
	}

	function selectedViewID() {
		var frmtbFindForm = document.getElementById("frmtbFindForm");
		return frmtbFindForm.selectView.options[frmtbFindForm.selectView.selectedIndex].value;
	}

</script>

<script src="<%: Url.LatestContent("~/Scripts/ctl_SetStyles.js")%>" type="text/javascript"></script>


<div <%=session("BodyTag")%>>
	<form action="" method="POST" id="frmtbFindForm" name="frmtbFindForm">

		<table align="center" class="outline" cellpadding="5" cellspacing="0" width="100%" height="100%">
			<tr>
				<td>
					<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
						<tr>
							<td height="10" colspan="3"></td>
						</tr>
						<tr>
							<td align="center" height="10" colspan="3">
								<h3 class="pageTitle" align="left">Add From Waiting List</h3>
							</td>
						</tr>
						<tr height="10">
							<td width="20"></td>
							<td>
								<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
									<tr>
										<td width="40">View :
										</td>
										<td width="10">&nbsp;
										</td>
										<td width="175">
											<select id="selectView" name="selectView" style="HEIGHT: 22px; WIDTH: 200px" class="combo">
												<%
													On Error Resume Next
													Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
													
													If (Len(sErrorDescription) = 0) And (Len(sFailureDescription) = 0) Then
														
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

													End If
												%>
											</select>
										</td>
										<td width="10">
											<input type="button" value="Go" id="btnGoView" class="btn" name="btnGoView"
												onclick="goView()"/>
										</td>
										<td>&nbsp;
										</td>
										<td width="40">Order :
										</td>
										<td width="10">&nbsp;
										</td>
										<td width="175">
											<select id="selectOrder" name="selectOrder" class="combo" style="HEIGHT: 22px; WIDTH: 200px">
												<%
													If (Len(sErrorDescription) = 0) And (Len(sFailureDescription) = 0) Then
														
														Dim rstTableOrderRecords = objDatabase.GetTableOrders(CInt(Session("optionLinkTableID")), 0)
														For Each objRow As DataRow In rstTableOrderRecords.Rows
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
										<td width="10">
											<input type="button" value="Go" id="btnGoOrder" name="btnGoOrder" class="btn"
												onclick="goOrder()" />
										</td>
									</tr>
								</table>
							</td>
							<td width="20"></td>
						</tr>
						<tr>
							<td height="10" colspan="3"></td>
						</tr>
						<tr>
							<td></td>
							<td>
								<div id="FindGridRow" style="height: 400px; margin-bottom: 50px;">
									<table id="ssOleDBGridRecords" name="ssOleDBGridRecords" style="width: 100%"></table>
								</div>
							</td>
							<td></td>
						</tr>
						<tr>
							<td height="10" colspan="3"></td>
						</tr>
						<tr>
							<td></td>
							<td height="10">
								<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
									<tr>
										<td colspan="4"></td>
									</tr>
									<tr>
										<td></td>
										<td width="10">
											<input id="cmdSelect" name="cmdSelect" type="button" value="Select" style="WIDTH: 75px" width="75" class="btn"
												onclick="tbSelect()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
										<td width="40"></td>
										<td width="10">
											<input id="cmdCancel" name="cmdCancel" type="button" value="Cancel" style="WIDTH: 75px" width="75" class="btn"
												onclick="tbCancel()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
									</tr>
								</table>
							</td>
							<td></td>
						</tr>
						<tr>
							<td height="10" colspan="3"></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<%
			Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtFailureDescription name=txtFailureDescription value=""" & sFailureDescription & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtOptionLinkTableID name=txtOptionLinkTableID value=" & Session("optionLinkTableID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtOptionLinkViewID name=txtOptionLinkViewID value=" & Session("optionLinkViewID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtOptionLinkOrderID name=txtOptionLinkOrderID value=" & Session("optionLinkOrderID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtOptionCourseTitle name=txtOptionCourseTitle value=" & Session("optionCourseTitle") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtOptionRecordID name=txtOptionRecordID value=" & Session("optionRecordID") & ">" & vbCrLf)
		%>
	</form>
	<input type="hidden" id="txtTicker" name="txtTicker" value="0">
	<input type="hidden" id="txtLastKeyFind" name="txtLastKeyFind" value="">
	<input type="hidden" id="txtStatusPExists" name="txtStatusPExists" value="<%=session("TB_TBStatusPExists")%>">

	<form action="tbAddFromWaitingListFind_Submit" method="post" id="frmGotoOption" name="frmGotoOption">
		<%Html.RenderPartial("~/Views/Shared/gotoOption.ascx")%>
	</form>

</div>

<script type="text/javascript">
	tbAddFromWaitingListFind_onload();
</script>

