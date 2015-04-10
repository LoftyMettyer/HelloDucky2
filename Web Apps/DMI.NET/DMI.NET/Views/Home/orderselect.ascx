<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="DMI.NET.Helpers" %>

<script type="text/javascript">

	function orderselect_window_onload() {

		// Expand the option frame and hide the work frame.
		//NOTE: keep this show/hide functionality before any DOM references.
		//window.parent.document.all.item("workframeset").cols = "0, *";
		$("#optionframe").attr("data-framesource", "SELECTORDER");
		$("#workframe").hide();
		$("#optionframe").show();

		var fOK;
		fOK = true;

		var frmOrderForm = document.getElementById("frmOrderForm");
		var sErrMsg = frmOrderForm.txtErrorDescription.value;
		if (sErrMsg.length > 0) {
			fOK = false;
			OpenHR.messageBox(sErrMsg);
			window.parent.location.replace("login");
		}

		//convert the table to a jqGrid with dblclick and keyboard interaction
		$(function () {
			tableToGrid("#ssOleDBGridOrderRecords", {
				ondblClickRow: function (rowID) {
					SelectOrder(rowID);
				},
				rowNum: 1000    //TODO set this to blocksize...
			});
		});

		$("#ssOleDBGridOrderRecords").jqGrid('bindKeys', {
			"onEnter": function (rowID) {
				SelectOrder(rowID);
			}
		});

		//resize the grid to the height of its container.
		$("#ssOleDBGridOrderRecords").jqGrid('setGridHeight', $("#orderGridRow").height());
		$("#div3").width($("#orderGridRow").width() + 11);

		if (fOK == true) {

			// Set focus onto one of the form controls. 
			// NB. This needs to be done before making any reference to the grid
			frmOrderForm.cmdCancel.focus();

			// Select the current record in the grid if its there, else select the top record if there is one.
			if (orderselect_rowCount() > 0) {
				if ($("#txtCurrentOrderID").val() > 0) {
					// Try to select the current record.
					locateRecord($("#txtCurrentOrderID").val(), true);
				} else {
					// Select the top row.
					orderselect_moveFirst();
				}
			}

			osrefreshControls();	// renamed to encapsulate.
		}
	}

	function SelectOrder() {
		//return selected orderID off to calling form.
		$("#optionframe").hide();
		$("#workframe").show();

		var frmOrderForm = document.getElementById("frmOrderForm");

		var postData = {		
			Action: optionActionType.SELECTORDER,
			ScreenID: frmOrderForm.txtOptionScreenID.value,
			TableID: frmOrderForm.txtOptionTableID.value,
			ViewID: frmOrderForm.txtOptionViewID.value,
			OrderID: orderselect_selectedRecordID(),
			<%:Html.AntiForgeryTokenForAjaxPost() %>};
		OpenHR.submitForm(null, "optionframe", false, postData, "orderselect_Submit");

	}


	function CancelOrder() {
		// Redisplay the workframe recedit control. 
		$("#optionframe").hide();
		$("#workframe").show();

		var sWorkPage = currentWorkFramePage();
		if (sWorkPage == "RECORDEDIT") {
			refreshData(); //should be in scope!
		}

		var postData = {
			Action: optionActionType.CANCEL,
			<%:Html.AntiForgeryTokenForAjaxPost() %>};
		OpenHR.submitForm(null, "optionframe", false, postData, "orderselect_Submit");

	}



	function orderselect_getRecordID(rowID) {
		//returns record ID for the selected row.
		return $("#ssOleDBGridOrderRecords").find("#" + rowID + " input[type=hidden]").val();
	}

	function orderselect_selectedRecordID() {
		/* Return the ID of the record selected in the find form. */
		var iRecordId;
		iRecordId = $("#ssOleDBGridOrderRecords").getGridParam('selrow');
		iRecordId = orderselect_getRecordID(iRecordId);

		return (iRecordId);
	}



	function orderselect_rowCount() {
		return $("#ssOleDBGridOrderRecords tr").length - 1;
	}



	function orderselect_moveFirst() {
		$("#ssOleDBGridOrderRecords").jqGrid('setSelection', 1);
	}



	/* Sequential search the grid for the required ID. */
	function locateRecord(psSearchFor) {
		var trID = $("#ssOleDBGridOrderRecords input[type=hidden]").filter(function () { return $(this).val() === psSearchFor; }).parent().parent().attr("id");

		if (Number(trID) > 0) {
			$("#ssOleDBGridOrderRecords").jqGrid('setSelection', trID);
		} else {
			//set top row.
			$("#ssOleDBGridOrderRecords").jqGrid('setSelection', 1);
		}
	}



	function osrefreshControls() {
		var frmOrderForm = document.getElementById("frmOrderForm");

		if (orderselect_rowCount() > 0) {
			if (orderselect_selectedRecordID() > 0) {
				button_disable(frmOrderForm.cmdSelectOrder, false);
			}
			else {
				button_disable(frmOrderForm.cmdSelectOrder, true);
			}
		}
		else {
			button_disable(frmOrderForm.cmdSelectOrder, true);
		}
	}



	function currentWorkFramePage() {
		var sCurrentPage = $("#workframe").attr("data-framesource");
		return (sCurrentPage);
	}



	function orderselect_addhandlers() {
		OpenHR.addActiveXHandler("ssOleDBGridOrderRecords", "dblClick", "ssOleDBGridOrderRecords_dblClick()");
		OpenHR.addActiveXHandler("ssOleDBGridOrderRecords", "KeyPress", "ssOleDBGridOrderRecords_KeyPress()");
	}



	function ssOleDBGridOrderRecords_dblClick() {
		SelectOrder();
	}



	function ssOleDBGridOrderRecords_KeyPress(iKeyAscii) {
		var iLastTick;
		var sFind;

		if ((iKeyAscii >= 32) && (iKeyAscii <= 255)) {
			var dtTicker = new Date();
			var iThisTick = new Number(dtTicker.getTime());
			if ($("#txtLastKeyFind").val().length > 0) {
				iLastTick = new Number($("#txtTicker").val());
			} else {
				iLastTick = new Number("0");
			}

			if (iThisTick > (iLastTick + 1500)) {
				sFind = String.fromCharCode(iKeyAscii);
			} else {
				sFind = $("#txtLastKeyFind").val() + String.fromCharCode(iKeyAscii);
			}

			$("#txtTicker").val(iThisTick);
			$("#txtLastKeyFind").val(sFind);

			locateRecord(sFind, false);
		}
	}

</script>


<script src="<%: Url.LatestContent("~/Scripts/ctl_SetStyles.js")%>" type="text/javascript"></script>

<div <%=session("BodyTag")%>>
	<form method="POST" id="frmOrderForm" name="frmOrderForm">
		<div class="absolutefull">
			<div class="pageTitleDiv" style="margin-bottom: 15px">
				<span class="pageTitle" id="PopupReportDefinition_PageTitle">Select Order</span>
			</div>
			<div id="orderGridRow" style="height: 70%; margin-right: 20px; margin-left: 20px;">
				<%Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
					Dim sErrorDescription = ""
	
					If Len(sErrorDescription) = 0 Then
										
						Dim rstOrderRecords = objDatabase.GetTableOrders(CleanNumeric(Session("optionTableID")), CleanNumeric(Session("optionViewID")))
											
						If Len(sErrorDescription) = 0 Then%>

				<table class="outline" style="width: 100%;" id="ssOleDBGridOrderRecords">
					<tr class="">
						<%For iLoop = 0 To (rstOrderRecords.Columns.Count - 1)
								Dim headerStyle As New StringBuilder
								Dim headerCaption As String
												
								headerStyle.Append("width: 373px; ")
												
								If rstOrderRecords.Columns(iLoop).ColumnName.ToUpper() = "ORDERID" Then
									headerStyle.Append("display: none; ")
								End If
	
								headerCaption = Replace(rstOrderRecords.Columns(iLoop).ColumnName.ToString(), "_", " ")
												
								headerStyle.Append("text-align: left; ")
										
								If rstOrderRecords.Columns(iLoop).ColumnName.ToUpper() <> "ORDERID" Then%>
						<th style="<%=headerStyle.ToString()%>"><%=headerCaption%></th>
						<%End If
						Next

						Dim lngRowCount = 0%>
					</tr>
					<%For Each objRow As DataRow In rstOrderRecords.Rows
							Dim iIDNumber As Integer = 0
												
							For iLoop = 0 To (rstOrderRecords.Columns.Count - 1)
								If rstOrderRecords.Columns(iLoop).ColumnName.ToUpper() = "ORDERID" Then
									iIDNumber = CInt(objRow(iLoop))
									Exit For
								End If
							Next
					%>

					<tr disabled="disabled" id="row_<%=iIDNumber.ToString()%>">
						<%
							For iLoop = 0 To (rstOrderRecords.Columns.Count - 1)
								If rstOrderRecords.Columns(iLoop).ColumnName.ToUpper() <> "ORDERID" Then
						%>
						<td class="" id="col_<%=NullSafeString(iIDNumber)%>"><%=Replace(NullSafeString(objRow(iLoop).ToString), "_", " ")%><input type='hidden' value='<%=NullSafeString(iIDNumber)%>'></td>
						<%
						End If
					Next
						%>
					</tr>
					<%
						lngRowCount = lngRowCount + 1
					Next
					%>
					<input type="hidden" id="txtCurrentOrderID" name="txtCurrentOrderID" value="<%=Session("optionOrderID")%>">

					<%
					End If
					%>
				</table>
				<%' Release the ADO command object.
					'cmdOrderRecords = Nothing
				End If
				%>
			</div>

			<div id='div3' class="clearboth" style='margin-top:50px'>
				<div id='divSelectOrder'>
					<input id="cmdSelectOrder" name="cmdSelectOrder" type="button" value="Select" class="btn" onclick="SelectOrder()" />
					<input id="cmdCancel" name="cmdCancel" type="button" value="Cancel" class="btn" onclick="CancelOrder()" />
				</div>
			</div>
			<%
				Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtOptionScreenID name=txtOptionScreenID value=" & Session("optionScreenID") & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtOptionTableID name=txtOptionTableID value=" & Session("optionTableID") & ">" & vbCrLf)
				Response.Write("<INPUT type='hidden' id=txtOptionViewID name=txtOptionViewID value=" & Session("optionViewID") & ">" & vbCrLf)
			%>
		</div>

	</form>
	<input type='hidden' id="txtTicker" name="txtTicker" value="0">
	<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">

</div>

<script type="text/javascript">
	orderselect_addhandlers();
	orderselect_window_onload();
</script>

