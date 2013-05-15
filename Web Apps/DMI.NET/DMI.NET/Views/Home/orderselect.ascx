<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script src="<%: Url.Content("~/Scripts/ctl_SetFont.js") %>" type="text/javascript"></script>

<script type="text/javascript">
	function orderselect_window_onload() {
		var fOK;
		fOK = true;

		var frmOrderForm = document.getElementById("frmOrderForm");
		var sErrMsg = frmOrderForm.txtErrorDescription.value;
		if (sErrMsg.length > 0) {
			fOK = false;
			OpenHR.messageBox(sErrMsg);
			window.parent.location.replace("login");
		}

		if (fOK == true) {
			setGridFont(frmOrderForm.ssOleDBGridOrderRecords);

			// Expand the option frame and hide the work frame.
			//window.parent.document.all.item("workframeset").cols = "0, *";
			$("#optionframe").attr("data-framesource", "SELECTORDER");
			$("#workframe").hide();
			$("#optionframe").show();

			// Set focus onto one of the form controls. 
			// NB. This needs to be done before making any reference to the grid
			frmOrderForm.cmdCancel.focus();

			// Select the current record in the grid if its there, else select the top record if there is one.
			if (frmOrderForm.ssOleDBGridOrderRecords.rows > 0) {
				if (frmOrderForm.txtCurrentOrderID.value > 0) {
					// Try to select the current record.
					locateRecord(frmOrderForm.txtCurrentOrderID.value, true);
				} else {
					// Select the top row.
					frmOrderForm.ssOleDBGridOrderRecords.MoveFirst();
					frmOrderForm.ssOleDBGridOrderRecords.SelBookmarks.Add(frmOrderForm.ssOleDBGridOrderRecords.Bookmark);
				}
			}

			// Get menu.asp to refresh the menu.
			// NPG20100824 Fault HRPRO1065 - leave menus disabled in these modal screens
			//window.parent.frames("menuframe").refreshMenu();

			// Hide the workframe recedit control. IE6 still displays it.
			var sWorkPage = currentWorkFramePage();
			if (sWorkPage == "RECORDEDIT") {
				//TODO: ??window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "hidden";
			} else {
				if (sWorkPage == "FIND") {
					//TODO: ??window.parent.frames("workframe").document.forms("frmFindForm").ssOleDBGridFindRecords.style.visibility = "hidden";
				}
			}

			osrefreshControls();	// renamed to encapsulate.
		}
	}
</script>

<script type="text/javascript">

	function SelectOrder() {
		// Redisplay the workframe recedit control. 
		var sWorkPage = currentWorkFramePage();
		if (sWorkPage == "RECORDEDIT") {
			//window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "visible";
			$("#optionframe").hide();
			$("#workframe").show();
		}
		else {
			if (sWorkPage == "FIND") {
				//window.parent.frames("workframe").document.forms("frmFindForm").ssOleDBGridFindRecords.style.visibility = "visible";
				$("#optionframe").hide();
				$("#workframe").show();

			}
		}

		var frmGotoOption = document.getElementById("frmGotoOption");
		var frmOrderForm = document.getElementById("frmOrderForm");

		frmGotoOption.txtGotoOptionScreenID.value = frmOrderForm.txtOptionScreenID.value;
		frmGotoOption.txtGotoOptionTableID.value = frmOrderForm.txtOptionTableID.value;
		frmGotoOption.txtGotoOptionViewID.value = frmOrderForm.txtOptionViewID.value;
		frmGotoOption.txtGotoOptionOrderID.value = osselectedRecordID();
		frmGotoOption.txtGotoOptionPage.value = "emptyoption";
		frmGotoOption.txtGotoOptionAction.value = "SELECTORDER";
		OpenHR.submitForm(frmGotoOption);
	}

	function CancelOrder() {		
		// Redisplay the workframe recedit control. 
		var sWorkPage = currentWorkFramePage();
		if (sWorkPage == "RECORDEDIT") {
			//window.parent.frames("workframe").document.forms("frmRecordEditForm").ctlRecordEdit.style.visibility = "visible";
			//window.parent.document.all.item("workframeset").cols = "*, 0";	
			$("#optionframe").hide();
			$("#workframe").show();
			refreshData(); //should be in scope!
		}
		else {
			if (sWorkPage == "FIND") {
				//window.parent.frames("workframe").document.forms("frmFindForm").ssOleDBGridFindRecords.style.visibility = "visible";
				$("#optionframe").hide();
				$("#workframe").show();
			}
		}
		var frmGotoOption = document.getElementById("frmGotoOption");

		frmGotoOption.txtGotoOptionAction.value = "CANCEL";
		frmGotoOption.txtGotoOptionPage.value = "emptyoption";
		OpenHR.submitForm(frmGotoOption);
	}

	/* Return the ID of the record selected in the find form. */
	function osselectedRecordID() {
		var iRecordID;
		var iIndex;
		var iIDColumnIndex;
		var sColumnName;

		iRecordID = 0;
		iIDColumnIndex = 0;
		var frmOrderForm = document.getElementById("frmOrderForm");

		if (frmOrderForm.ssOleDBGridOrderRecords.SelBookmarks.Count > 0) {
			for (iIndex = 0; iIndex < frmOrderForm.ssOleDBGridOrderRecords.Cols; iIndex++) {
				sColumnName = frmOrderForm.ssOleDBGridOrderRecords.Columns(iIndex).Name;
				if (sColumnName.toUpperCase() == "ORDERID") {
					iIDColumnIndex = iIndex;
					break;
				}
			}

			iRecordID = frmOrderForm.ssOleDBGridOrderRecords.Columns(iIDColumnIndex).Value;
		}

		return (iRecordID);
	}

	/* Sequential search the grid for the required ID. */
	function locateRecord(psSearchFor, pfIDMatch) {
		var fFound;
		var iIndex;
		var iIDColumnIndex;
		var sColumnName;
		var frmOrderForm = document.getElementById("frmOrderForm");

		fFound = false;

		frmOrderForm.ssOleDBGridOrderRecords.redraw = false;

		if (pfIDMatch == true) {
			// Locate the ID column in the grid.
			iIDColumnIndex = -1
			for (iIndex = 0; iIndex < frmOrderForm.ssOleDBGridOrderRecords.Cols; iIndex++) {
				sColumnName = frmOrderForm.ssOleDBGridOrderRecords.Columns(iIndex).Name;
				if (sColumnName.toUpperCase() == "ORDERID") {
					iIDColumnIndex = iIndex;
					break;
				}
			}

			if (iIDColumnIndex >= 0) {
				frmOrderForm.ssOleDBGridOrderRecords.MoveLast();
				frmOrderForm.ssOleDBGridOrderRecords.MoveFirst();

				for (iIndex = 1; iIndex <= frmOrderForm.ssOleDBGridOrderRecords.rows; iIndex++) {
					if (frmOrderForm.ssOleDBGridOrderRecords.Columns(iIDColumnIndex).value == psSearchFor) {
						frmOrderForm.ssOleDBGridOrderRecords.SelBookmarks.Add(frmOrderForm.ssOleDBGridOrderRecords.Bookmark);
						fFound = true;
						break;
					}

					if (iIndex < frmOrderForm.ssOleDBGridOrderRecords.rows) {
						frmOrderForm.ssOleDBGridOrderRecords.MoveNext();
					}
					else {
						break;
					}
				}
			}
		}
		else {
			for (iIndex = 1; iIndex <= frmOrderForm.ssOleDBGridOrderRecords.rows; iIndex++) {
				var sGridValue = new String(frmOrderForm.ssOleDBGridOrderRecords.Columns(0).value);
				sGridValue = sGridValue.substr(0, psSearchFor.length).toUpperCase();
				if (sGridValue == psSearchFor.toUpperCase()) {
					frmOrderForm.ssOleDBGridOrderRecords.SelBookmarks.Add(frmOrderForm.ssOleDBGridOrderRecords.Bookmark);
					fFound = true;
					break;
				}

				if (iIndex < frmOrderForm.ssOleDBGridOrderRecords.rows) {
					frmOrderForm.ssOleDBGridOrderRecords.MoveNext();
				}
				else {
					break;
				}
			}
		}

		if ((fFound == false) && (frmOrderForm.ssOleDBGridOrderRecords.rows > 0)) {
			// Select the top row.
			frmOrderForm.ssOleDBGridOrderRecords.MoveFirst();
			frmOrderForm.ssOleDBGridOrderRecords.SelBookmarks.Add(frmOrderForm.ssOleDBGridOrderRecords.Bookmark);
		}

		frmOrderForm.ssOleDBGridOrderRecords.redraw = true;
	}

	function osrefreshControls() {
		var frmOrderForm = document.getElementById("frmOrderForm");

		if (frmOrderForm.ssOleDBGridOrderRecords.rows > 0) {
			if (frmOrderForm.ssOleDBGridOrderRecords.SelBookmarks.Count > 0) {
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
		//// Return the current page in the workframeset.
		//sCols = window.parent.document.all.item("workframeset").cols;

		//re = / /gi;
		//sCols = sCols.replace(re, "");
		//sCols = sCols.substr(0, 1);

		//// Work frame is in view.
		//sCurrentPage = window.parent.frames("workframe").document.location;
		//sCurrentPage = sCurrentPage.toString();

		//if (sCurrentPage.lastIndexOf("/") > 0) {
		//	sCurrentPage = sCurrentPage.substr(sCurrentPage.lastIndexOf("/") + 1);
		//}

		//if (sCurrentPage.indexOf(".") > 0) {
		//	sCurrentPage = sCurrentPage.substr(0, sCurrentPage.indexOf("."));
		//}

		//re = / /gi;
		//sCurrentPage = sCurrentPage.replace(re, "");
		//sCurrentPage = sCurrentPage.toUpperCase();


		var sCurrentPage = $("#workframe").attr("data-framesource").replace(".asp", "");
		return (sCurrentPage);
	}

</script>

<script type="text/javascript">

	function orderselect_addhandlers() {
		OpenHR.addActiveXHandler("ssOleDBGridOrderRecords", "dblClick", ssOleDBGridOrderRecords_dblClick);
		OpenHR.addActiveXHandler("ssOleDBGridOrderRecords", "KeyPress", ssOleDBGridOrderRecords_KeyPress);
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


<script src="<%: Url.Content("~/Scripts/ctl_SetStyles.js") %>" type="text/javascript"></script>

<div <%=session("BodyTag")%>>
	<form action="" method="POST" id="frmOrderForm" name="frmOrderForm">
		<table align="center" class="outline" cellpadding="5" cellspacing="0" width="100%" height="100%">
			<tr>
				<td>
					<table id="orderTable" width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">
						<tr>
							<td height="10" colspan="3">
								<h3 align="center">Select Order</h3>
							</td>
						</tr>
						<tr>
							<td width="20"></td>
							<td>
								<%
									Dim sErrorDescription = ""
	
									If Len(sErrorDescription) = 0 Then
										' Get the order records.
										Dim cmdOrderRecords = CreateObject("ADODB.Command")
										cmdOrderRecords.CommandText = "sp_ASRIntGetTableOrders"
										cmdOrderRecords.CommandType = 4	' Stored Procedure
										cmdOrderRecords.ActiveConnection = Session("databaseConnection")

										Dim prmTableID = cmdOrderRecords.CreateParameter("tableID", 3, 1)
										cmdOrderRecords.Parameters.Append(prmTableID)
										prmTableID.value = CleanNumeric(Session("optionTableID"))

										Dim prmViewID = cmdOrderRecords.CreateParameter("viewID", 3, 1)
										cmdOrderRecords.Parameters.Append(prmViewID)
										prmViewID.value = CleanNumeric(Session("optionViewID"))

										Err.Clear()
										Dim rstOrderRecords = cmdOrderRecords.Execute

										If (Err.Number <> 0) Then
											sErrorDescription = "The order records could not be retrieved." & vbCrLf & FormatError(Err.Description)
										End If

										If Len(sErrorDescription) = 0 Then
											' Instantiate and initialise the grid. 
											Response.Write("			<OBJECT classid=""clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"" id=ssOleDBGridOrderRecords name=ssOleDBGridOrderRecords codebase=""cabs/COAInt_Grid.cab#version=3,1,3,6"" style=""LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:400px"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""ScrollBars"" VALUE=""4"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""_Version"" VALUE=""196617"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""DataMode"" VALUE=""2"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""Cols"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""Rows"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""BorderStyle"" VALUE=""1"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""RecordSelectors"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""GroupHeaders"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""ColumnHeaders"" VALUE=""1"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""GroupHeadLines"" VALUE=""1"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""HeadLines"" VALUE=""1"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""FieldDelimiter"" VALUE=""(None)"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""FieldSeparator"" VALUE=""(Tab)"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""Col.Count"" VALUE=""" & rstOrderRecords.fields.count & """>" & vbCrLf)
											Response.Write("				<PARAM NAME=""stylesets.count"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""TagVariant"" VALUE=""EMPTY"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""UseGroups"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""HeadFont3D"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""Font3D"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""DividerType"" VALUE=""3"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""DividerStyle"" VALUE=""1"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""DefColWidth"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""BeveColorScheme"" VALUE=""2"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""BevelColorFrame"" VALUE=""-2147483642"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""BevelColorHighlight"" VALUE=""-2147483628"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""BevelColorShadow"" VALUE=""-2147483632"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""BevelColorFace"" VALUE=""-2147483633"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""CheckBox3D"" VALUE=""-1"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""AllowAddNew"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""AllowDelete"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""AllowUpdate"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""MultiLine"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""ActiveCellStyleSet"" VALUE="""">" & vbCrLf)
											Response.Write("				<PARAM NAME=""RowSelectionStyle"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""AllowRowSizing"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""AllowGroupSizing"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""AllowColumnSizing"" VALUE=""-1"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""AllowGroupMoving"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""AllowColumnMoving"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""AllowGroupSwapping"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""AllowColumnSwapping"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""AllowGroupShrinking"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""AllowColumnShrinking"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""AllowDragDrop"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""UseExactRowCount"" VALUE=""-1"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""SelectTypeCol"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""SelectTypeRow"" VALUE=""1"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""SelectByCell"" VALUE=""-1"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""BalloonHelp"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""RowNavigation"" VALUE=""1"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""CellNavigation"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""MaxSelectedRows"" VALUE=""1"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""HeadStyleSet"" VALUE="""">" & vbCrLf)
											Response.Write("				<PARAM NAME=""StyleSet"" VALUE="""">" & vbCrLf)
											Response.Write("				<PARAM NAME=""ForeColorEven"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""ForeColorOdd"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""BackColorEven"" VALUE=""16777215"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""BackColorOdd"" VALUE=""16777215"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""Levels"" VALUE=""1"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""RowHeight"" VALUE=""503"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""ExtraHeight"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""ActiveRowStyleSet"" VALUE="""">" & vbCrLf)
											Response.Write("				<PARAM NAME=""CaptionAlignment"" VALUE=""2"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""SplitterPos"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""SplitterVisible"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""Columns.Count"" VALUE=""" & rstOrderRecords.fields.count & """>" & vbCrLf)

											For iLoop = 0 To (rstOrderRecords.fields.count - 1)

												If rstOrderRecords.fields(iLoop).name = "orderID" Then
													Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Width"" VALUE=""0"">" & vbCrLf)
													Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Visible"" VALUE=""0"">" & vbCrLf)
												Else
													Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Width"" VALUE=""100000"">" & vbCrLf)
													Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Visible"" VALUE=""-1"">" & vbCrLf)
												End If
	
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Columns.Count"" VALUE=""1"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Caption"" VALUE=""" & Replace(rstOrderRecords.fields(iLoop).name, "_", " ") & """>" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Name"" VALUE=""" & rstOrderRecords.fields(iLoop).name & """>" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Alignment"" VALUE=""0"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").CaptionAlignment"" VALUE=""3"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Bound"" VALUE=""0"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").AllowSizing"" VALUE=""1"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").DataField"" VALUE=""Column " & iLoop & """>" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").DataType"" VALUE=""8"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Level"" VALUE=""0"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").NumberFormat"" VALUE="""">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Case"" VALUE=""0"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").FieldLen"" VALUE=""4096"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").VertScrollBar"" VALUE=""0"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Locked"" VALUE=""0"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Style"" VALUE=""0"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").ButtonsAlways"" VALUE=""0"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").RowCount"" VALUE=""0"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").ColCount"" VALUE=""1"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").HasHeadForeColor"" VALUE=""0"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").HasHeadBackColor"" VALUE=""0"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").HasForeColor"" VALUE=""0"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").HasBackColor"" VALUE=""0"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").HeadForeColor"" VALUE=""0"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").HeadBackColor"" VALUE=""0"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").ForeColor"" VALUE=""0"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").BackColor"" VALUE=""0"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").HeadStyleSet"" VALUE="""">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").StyleSet"" VALUE="""">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Nullable"" VALUE=""1"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").Mask"" VALUE="""">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").PromptInclude"" VALUE=""0"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").ClipMode"" VALUE=""0"">" & vbCrLf)
												Response.Write("				<PARAM NAME=""Columns(" & iLoop & ").PromptChar"" VALUE=""95"">" & vbCrLf)
											Next

											Response.Write("				<PARAM NAME=""UseDefaults"" VALUE=""-1"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""TabNavigation"" VALUE=""1"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""_ExtentX"" VALUE=""17330"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""_ExtentY"" VALUE=""1323"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""_StockProps"" VALUE=""79"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""Caption"" VALUE="""">" & vbCrLf)
											Response.Write("				<PARAM NAME=""ForeColor"" VALUE=""0"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""BackColor"" VALUE=""16777215"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""Enabled"" VALUE=""-1"">" & vbCrLf)
											Response.Write("				<PARAM NAME=""DataMember"" VALUE="""">" & vbCrLf)

											Dim lngRowCount = 0
											Do While Not rstOrderRecords.EOF
												For iLoop = 0 To (rstOrderRecords.fields.count - 1)
													Response.Write("				<PARAM NAME=""Row(" & lngRowCount & ").Col(" & iLoop & ")"" VALUE=""" & Replace(rstOrderRecords.Fields(iLoop).Value, "_", " ") & """>" & vbCrLf)
												Next
												lngRowCount = lngRowCount + 1
												rstOrderRecords.MoveNext()
											Loop
											Response.Write("				<PARAM NAME=""Row.Count"" VALUE=""" & lngRowCount & """>" & vbCrLf)
											Response.Write("			</OBJECT>" & vbCrLf)
											Response.Write("			<INPUT type='hidden' id=txtCurrentOrderID name=txtCurrentOrderID value=" & Session("optionOrderID") & ">" & vbCrLf)

											' Release the ADO recordset object.
											rstOrderRecords.close()
											rstOrderRecords = Nothing
										End If
	
										' Release the ADO command object.
										cmdOrderRecords = Nothing
									End If
								%>
							</td>
							<td width="20"></td>
						</tr>
						<tr>
							<td height="10" colspan="3"></td>
						</tr>
						<tr>
							<td width="20"></td>
							<td height="10">
								<table width="100%" class="invisible" cellspacing="0" cellpadding="0">
									<tr>
										<td>&nbsp;</td>
										<td width="10">
											<input id="cmdSelectOrder" name="cmdSelectOrder" type="button" value="Select" style="WIDTH: 75px" width="75" class="btn"
												onclick="SelectOrder()"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
							</td>
							<td width="40"></td>
							<td width="10">
								<input id="cmdCancel" name="cmdCancel" type="button" value="Cancel" style="WIDTH: 75px" width="75" class="btn"
									onclick="CancelOrder()"
									onmouseover="try{button_onMouseOver(this);}catch(e){}"
									onmouseout="try{button_onMouseOut(this);}catch(e){}"
									onfocus="try{button_onFocus(this);}catch(e){}"
									onblur="try{button_onBlur(this);}catch(e){}" />
							</td>
						</tr>
					</table>
				</td>
				<td width="20"></td>
			</tr>
			<tr>
				<td height="10" colspan="3"></td>
			</tr>
		</table>
		</td>
	</tr>
</TABLE>
		<%
			Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtOptionScreenID name=txtOptionScreenID value=" & Session("optionScreenID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtOptionTableID name=txtOptionTableID value=" & Session("optionTableID") & ">" & vbCrLf)
			Response.Write("<INPUT type='hidden' id=txtOptionViewID name=txtOptionViewID value=" & Session("optionViewID") & ">" & vbCrLf)
		%>
	</form>
	<input type='hidden' id="txtTicker" name="txtTicker" value="0">
	<input type='hidden' id="txtLastKeyFind" name="txtLastKeyFind" value="">

	<form action="orderselect_Submit" method="post" id="frmGotoOption" name="frmGotoOption">
		<%Html.RenderPartial("~/Views/Shared/gotoOption.ascx")%>
	</form>

</div>

<script type="text/javascript">
	orderselect_addhandlers();
	orderselect_window_onload();
</script>

