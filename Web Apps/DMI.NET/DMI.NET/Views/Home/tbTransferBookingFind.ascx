<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<script type="text/javascript">
	function tbTransferBookingFind_onload() {
		var fOK;
		fOK = true;

		var frmtbFindForm = document.getElementById("frmtbFindForm");
		
		var sErrMsg = frmtbFindForm.txtErrorDescription.value;
		if (sErrMsg.length > 0) {
			fOK = false;
			OpenHR.messageBox(sErrMsg);
			window.parent.location.replace("login");
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
			$("#optionframe").attr("data-framesource", "TBTRANSFERBOOKINGFIND");
			$("#workframe").hide();
			$("#optionframe").show();

			// Set focus onto one of the form controls. 
			// NB. This needs to be done before making any reference to the grid
			frmtbFindForm.cmdCancel.focus();

			//TODO: window.parent.frames("workframe").document.forms("frmtbFindForm").ssOleDBGridFindRecords.style.visibility = "hidden";

			// Get the optionData.asp to get the link find records.
			var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
			optionDataForm.txtOptionAction.value = "LOADTRANSFERBOOKING";
			optionDataForm.txtOptionTableID.value = frmtbFindForm.txtOptionLinkTableID.value;
			optionDataForm.txtOptionViewID.value = frmtbFindForm.selectView.options[frmtbFindForm.selectView.selectedIndex].value;
			optionDataForm.txtOptionOrderID.value = frmtbFindForm.selectOrder.options[frmtbFindForm.selectOrder.selectedIndex].value;
			optionDataForm.txtOptionRecordID.value = frmtbFindForm.txtOptionRecordID.value;
			optionDataForm.txtOptionFirstRecPos.value = 1;
			optionDataForm.txtOptionCurrentRecCount.value = 0;
			optionDataForm.txtOptionPageAction.value = "LOAD";

			refreshOptionData();
		}
	}
</script>

<script type="text/javascript">
	function tbSelect()
	{  
		//TODO: window.parent.frames("workframe").document.forms("frmtbFindForm").ssOleDBGridFindRecords.style.visibility = "visible";
		var frmGotoOption = document.getElementById("frmGotoOption");
		var frmtbFindForm = document.getElementById("frmtbFindForm");		
		frmGotoOption.txtGotoOptionAction.value = "SELECTTRANSFERBOOKING_1";
		frmGotoOption.txtGotoOptionRecordID.value = frmtbFindForm.txtOptionRecordID.value;
		var selRowId = $("#ssOleDBGridRecords").jqGrid('getGridParam', 'selrow');
		var recordID = $("#ssOleDBGridRecords").jqGrid('getCell', selRowId, 'ID');
		frmGotoOption.txtGotoOptionLinkRecordID.value = recordID; //tbselectedRecordID();
		frmGotoOption.txtGotoOptionPage.value = "emptyoption";

		var optionDataForm = OpenHR.getForm("optiondataframe", "frmOptionData");
		frmGotoOption.txtGotoOptionLookupValue.value = optionDataForm.txtStatus.value;

		OpenHR.submitForm(frmGotoOption);
	}

	function tbCancel()
	{  
		//TODO: window.parent.frames("workframe").document.forms("frmtbFindForm").ssOleDBGridFindRecords.style.visibility = "visible";
		$("#optionframe").hide();
		$("#workframe").show();

		var frmGotoOption = document.getElementById("frmGotoOption");
		
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
	
		$("#txtLoading").val(0);
	}

	function goView() {
		// Get the optionData.asp to get the link find records.
		var frmtbFindForm = document.getElementById("frmtbFindForm");
		
		var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
		optionDataForm.txtOptionAction.value = "LOADTRANSFERBOOKING";
		optionDataForm.txtOptionTableID.value = frmtbFindForm.txtOptionLinkTableID.value;
		optionDataForm.txtOptionViewID.value = frmtbFindForm.selectView.options[frmtbFindForm.selectView.selectedIndex].value;
		optionDataForm.txtOptionOrderID.value = frmtbFindForm.selectOrder.options[frmtbFindForm.selectOrder.selectedIndex].value;
		optionDataForm.txtOptionRecordID.value = frmtbFindForm.txtOptionRecordID.value;
		optionDataForm.txtOptionFirstRecPos.value = 1;
		optionDataForm.txtOptionCurrentRecCount.value = 0;

		frmtbFindForm.txtOptionLinkViewID.value = optionDataForm.txtOptionViewID.value;
		frmtbFindForm.txtOptionLinkOrderID.value = optionDataForm.txtOptionOrderID.value;

		refreshOptionData(); //should be in scope...
	}

	function goOrder() {
		// Get the optionData.asp to get the link find records.
		var optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
		var frmtbFindForm = document.getElementById("frmtbFindForm");
		
		optionDataForm.txtOptionAction.value = "LOADTRANSFERBOOKING";
		optionDataForm.txtOptionTableID.value = frmtbFindForm.txtOptionLinkTableID.value;
		optionDataForm.txtOptionViewID.value = frmtbFindForm.selectView.options[frmtbFindForm.selectView.selectedIndex].value;
		optionDataForm.txtOptionOrderID.value = frmtbFindForm.selectOrder.options[frmtbFindForm.selectOrder.selectedIndex].value;
		optionDataForm.txtOptionRecordID.value = frmtbFindForm.txtOptionRecordID.value;
		optionDataForm.txtOptionFirstRecPos.value = 1;
		optionDataForm.txtOptionCurrentRecCount.value = 0;

		frmtbFindForm.txtOptionLinkViewID.value = optionDataForm.txtOptionViewID.value;
		frmtbFindForm.txtOptionLinkOrderID.value = optionDataForm.txtOptionOrderID.value;

		refreshOptionData();	//should be in scope
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
<FORM action="" method="POST" id="frmtbFindForm" name="frmtbFindForm">

<table align=center class="outline" cellPadding=5 cellSpacing=0 width=100% height=100%>
	<TR>
		<TD>
			<TABLE WIDTH="100%" height="100%" class="invisible" CELLSPACING=0 CELLPADDING=0>
				<TR>
					<TD height=10 colspan=3></td>
				</tr>
				<TR>
					<TD align=center height=10 colspan=3>
						<h3 class="pageTitle" align="left">Transfer Booking</h3>
					</td>
				</tr>
				<tr height=10>
					<td width=20></td>
					<td>
						<table WIDTH=100% class="invisible" CELLSPACING="0" CELLPADDING="0">
							<TR>
								<TD width=40>
									View :
								</TD>
								<TD width=10>
									&nbsp;
								</TD>
								<TD width=175>
									<SELECT id=selectView name=selectView class="combo" style="HEIGHT: 22px; WIDTH: 200px">
<%
	on error resume next
	Dim sErrorDescription = ""
	Dim sFailureDescription = ""

	Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
	
	if (len(sErrorDescription) = 0) and (len(sFailureDescription) = 0) then
		' Get the view records.
	
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

	end if
%>
									</SELECT>						
								</TD>
								<TD width=10>
									<INPUT type="button" value="Go" id=btnGoView name=btnGoView class="btn"
											onclick="goView()" />
								</TD>
								<TD>
									&nbsp;
								</TD>
								<TD width=40>
									Order :
								</TD>
								<TD width=10>
									&nbsp;
								</TD>
								<TD width=175>
									<SELECT id=selectOrder name=selectOrder class="combo" style="HEIGHT: 22px; WIDTH: 200px">
<%
	If (Len(sErrorDescription) = 0) And (Len(sFailureDescription) = 0) Then

		Dim rstOrderRecords = objDatabase.GetTableOrders(CInt(CleanNumeric(Session("optionLinkTableID"))), 0)
		For Each objRow As DataRow In rstOrderRecords.Rows
			Response.Write("						<option value=" & objRow(1))
			If objRow(1) = CInt(Session("optionLinkOrderID")) Then
				Response.Write(" SELECTED")
			End If
			Response.Write(">" & Replace(objRow(0).ToString(), "_", " ") & "</option>" & vbCrLf)
		Next
		
	End If
%>
									</SELECT>
								</TD>
								<TD width=10>
									<INPUT type="button" value="Go" id=btnGoOrder name=btnGoOrder class="btn"   
											onclick="goOrder()"
																				onmouseover="try{button_onMouseOver(this);}catch(e){}" 
																				onmouseout="try{button_onMouseOut(this);}catch(e){}"
																				onfocus="try{button_onFocus(this);}catch(e){}"
																				onblur="try{button_onBlur(this);}catch(e){}" />
								</TD>
							</TR>
						</table>
					</td>
					<td width=20></td>
				</tr>
				
				<TR>
					<TD height=10 colspan=3></td>
				</tr>
				
				<TR>
					<td></td>
					<TD>
						<div id="FindGridRow" style="height: 400px; margin-bottom: 50px;">
							<table id="ssOleDBGridRecords" name="ssOleDBGridRecords" style="width: 100%"></table>
						</div>
					</TD>
					<td></td>
				</TR>

				<TR>
					<TD height=10 colspan=3></td>
				</tr>

				<tr>
					<td></td>
					<td height="10">
						<table WIDTH=100% class="invisible" CELLSPACING="0" CELLPADDING="0">
							<TR>
								<TD colspan=4>
								</TD>
							</TR>
							<tr>	
								<td>
								</td>
								<td width=10>
									<input id="cmdSelect" name="cmdSelect" type="button" class="btn" value="Select" style="WIDTH: 75px" width="75" 
											onclick="tbSelect()"
																				onmouseover="try{button_onMouseOver(this);}catch(e){}" 
																				onmouseout="try{button_onMouseOut(this);}catch(e){}"
																				onfocus="try{button_onFocus(this);}catch(e){}"
																				onblur="try{button_onBlur(this);}catch(e){}" />
								</td>
								<td width=40>
								</td>
								<td width=10>
									<input id="cmdCancel" name="cmdCancel" type="button" class="btn" value="Cancel" style="WIDTH: 75px" width="75" 
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
				<TR>
					<TD height=10 colspan=3></td>
				</tr>
			</TABLE>
		</td>
	</tr>
</TABLE>
<%
	Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtFailureDescription name=txtFailureDescription value=""" & sFailureDescription & """>" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtOptionLinkTableID name=txtOptionLinkTableID value=" & Session("optionLinkTableID") & ">" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtOptionLinkViewID name=txtOptionLinkViewID value=" & Session("optionLinkViewID") & ">" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtOptionLinkOrderID name=txtOptionLinkOrderID value=" & Session("optionLinkOrderID") & ">" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtOptionCourseTitle name=txtOptionCourseTitle value=" & Session("optionCourseTitle") & ">" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtOptionRecordID name=txtOptionRecordID value=" & Session("optionRecordID") & ">" & vbCrLf)
%>
</FORM>
<INPUT type='hidden' id=txtTicker name=txtTicker value=0>
<INPUT type='hidden' id=txtLastKeyFind name=txtLastKeyFind value="">
<INPUT type='hidden' id=txtLoading name=txtLoading value=1>

<FORM action="tbTransferBookingFind_Submit" method=post id=frmGotoOption name=frmGotoOption style="visibility:hidden;display:none">
		<%Html.RenderPartial("~/Views/Shared/gotoOption.ascx")%>
</FORM>

</div>

<script type="text/javascript"> tbTransferBookingFind_onload();</script>
