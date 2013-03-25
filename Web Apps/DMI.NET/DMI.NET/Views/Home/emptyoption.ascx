<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<%

	'Clear the single record session variable
	Session("singleRecordID") = 0
	Session("optionDefSelRecordID") = 0

	If Len(Session("timestamp")) = 0 Then
		Session("timestamp") = 0
	End If

%>



<script type="text/javascript">
	function emptyoption_onload() {
		var fNoAction;
		var sCurrentWorkFramePage = $("#workframe").attr("data-framesource").replace(".asp", ""); //OpenHR.currentWorkPage();
		var frmMenu = OpenHR.getForm("menuframe", "frmWorkAreaInfo");

		// Do nothing if the menu controls are not yet instantiated.
		if (frmMenu != null) {
			if (OpenHR.currentWorkPage() != "DEFAULT") {
				fNoAction = true;

				if (txtAction.value == "SELECTORDER") {
					fNoAction = false;

					if (sCurrentWorkFramePage == "RECORDEDIT") {
						var frmRecEdit = OpenHR.getForm("workframe","frmRecordEditForm");
						frmRecEdit.txtRecEditFromDef.value = txtFromDef.value;
						frmRecEdit.txtCurrentOrderID.value = txtOrderID.value;
						OpenHR.getFrame("workframe").refreshData();
						$("#optionframe").attr("data-framesource", "EMPTYOPTION");
					} else {
						if (sCurrentWorkFramePage == "FIND") {
							var frmFind = OpenHR.getForm("workframe","frmFindForm");
							frmFind.txtCurrentOrderID.value = txtOrderID.value;
							menu_reloadFindPage("RELOAD", "");
						}
					}
				}

				if (txtAction.value == "SELECTFILTER") {
					fNoAction = false;

					if (sCurrentWorkFramePage == "RECORDEDIT") {
						var frmRecEdit = OpenHR.getForm("workframe","frmRecordEditForm");
						frmRecEdit.txtRecEditFilterSQL.value = txtFilterSQL.value;
						frmRecEdit.txtRecEditFilterDef.value = txtFilterDef.value;
						OpenHR.getFrame("workframe").refreshData();
						$("#optionframe").attr("data-framesource", "EMPTYOPTION");

					} else {
						if (sCurrentWorkFramePage == "FIND") {
							var frmFind = OpenHR.getForm("workframe","frmFindForm");
							frmFind.txtFilterSQL.value = txtFilterSQL.value;
							frmFind.txtFilterDef.value = txtFilterDef.value;
							menu_reloadFindPage("RELOAD", "");
						}
					}
				}

				if (txtAction.value == "QUICKFIND") {					
					fNoAction = false;
					var frmData = OpenHR.getForm("dataframe","frmData");
					frmData.txtRecordID.value = txtRecordID.value;					
					$("#workframe").attr("data-framesource", "RECORDEDIT");
					$("#optionframe").hide();
					$("#workframe").show();
	                
					//OpenHR.getFrame("workframe").refreshData();
					refreshData();	//recedit


				}

				if (txtAction.value == "SELECTLINK") {
					fNoAction = false;
					var recEditControl = OpenHR.getForm("workframe","frmRecordEditForm").ctlRecordEdit;
					var sControlName;
					var sColumnID;
					var dataCollection = frmEmptyOption.elements;
					if (dataCollection != null) {
						// Need to hide the popup in case setdata causes
						// the intrecedit control to display an error message.
						menu_refreshMenu();

						for (i = 0; i < dataCollection.length; i++) {
							sControlName = dataCollection.item(i).name;
							sControlName = sControlName.substr(0, 8);
							if (sControlName == "txtData_") {
								sColumnID = dataCollection.item(i).name;
								sColumnID = sColumnID.substr(8);
								recEditControl.setData(sColumnID, dataCollection.item(i).value);
							}
						}
						recEditControl.changed = true;
					}

					$("#optionframe").attr("data-framesource", "EMPTYOPTION");
					//window.setTimeout("window.parent.frames('menuframe').refreshMenu()", 100);
				}

				if (txtAction.value == "SELECTLOOKUP") {
					fNoAction = false;
					var recEditControl = OpenHR.getForm("workframe","frmRecordEditForm").ctlRecordEdit;
					recEditControl.setData(txtColumnID.value, txtValue.value);
					recEditControl.changed = true;

					$("#optionframe").attr("data-framesource", "EMPTYOPTION");
					//window.setTimeout("window.parent.frames('menuframe').refreshMenu()", 100);
				}

				if ((txtAction.value == "SELECTIMAGE") || (txtAction.value == "SELECTOLE")) {
					fNoAction = false;
					var recEditControl = OpenHR.getForm("workframe","frmRecordEditForm").ctlRecordEdit;
					recEditControl.setData(txtColumnID.value, txtFile.value);
					recEditControl.ChangedOLEPhoto(txtColumnID.value, "");
					recEditControl.changed = true;

					$("#optionframe").attr("data-framesource", "EMPTYOPTION");
					//window.setTimeout("window.parent.frames('menuframe').refreshMenu()", 100);
				}

				if (txtAction.value == "LINKOLE") {

					fNoAction = false;
					var recEditControl = OpenHR.getForm("workframe","frmRecordEditForm").ctlRecordEdit;
					recEditControl.setData(txtColumnID.value, txtFile.value);
					recEditControl.timestamp = <%=session("timestamp")%>;
                	//recEditControl.changed = true;

                    $("#optionframe").attr("data-framesource", "EMPTYOPTION");
                	//window.setTimeout("window.parent.frames('menuframe').refreshMenu()", 100);

									}

									if ((txtAction.value == "SELECTTRANSFERCOURSE") ||
											(txtAction.value == "SELECTBOOKCOURSE_2") ||
											(txtAction.value == "SELECTTRANSFERBOOKING_1") ||
											(txtAction.value == "SELECTADDFROMWAITINGLIST_2") ||
											(txtAction.value == "SELECTBULKBOOKINGS")) {

										if ((txtAction.value == "SELECTBOOKCOURSE_2") ||
                        (txtAction.value == "SELECTTRANSFERBOOKING_1") ||
                        (txtAction.value == "SELECTADDFROMWAITINGLIST_2")) {
											sPrefix = "The employee";
											sPrefix2 = "The employee is";
										} else {
											if (txtAction.value == "SELECTBULKBOOKINGS") {
												sPrefix = "A delegate";
												sPrefix2 = "Some delegates are";
											} else {
												sPrefix = "A delegate";
												sPrefix2 = "Some transferred delegates are";
											}
										}

										fNoAction = false;

										$("#optionframe").attr("data-framesource", "EMPTYOPTION");                    
										menu_refreshMenu();

										fTransferOK = true;

										iResultCode = txtResultCode.value;
										if (iResultCode > 0) {
											iOverlapCode = iResultCode % 10;
											iResultCode = iResultCode - iOverlapCode;
											iResultCode = iResultCode / 10;
											iAvailabilityCode = iResultCode % 10;
											iResultCode = iResultCode - iAvailabilityCode;
											iResultCode = iResultCode / 10;
											iPreReqCode = iResultCode % 10;
											iResultCode = iResultCode - iPreReqCode;
											iResultCode = iResultCode / 10;
											iOverbookCode = iResultCode;

											sTransferErrorMsg = ""
											sTransferWarningMsg = ""

											if (txtAction.value == "SELECTBULKBOOKINGS") {
												sPreReqFails = txtPreReqFails.value;
												sUnAvailFails = txtUnAvailFails.value;
												sOverlapFails = txtOverlapFails.value;
												sOverBookFails = txtOverBookFails.value;

												/*	alert('These delegates have failed the following checks: \n' +
														'\nCourse Prequisites - ' + sPreReqFails +
														'\nUnavailable - ' + sUnAvailFails +
														'\nOvelapping Course - ' + sOverlapFails);
																		
														alert('iResultCode = ' + iResultCode + 
														'\niPreReqCode = ' + iPreReqCode +
														'\niOverlapCode = ' + iOverlapCode +
														'\niAvailabilityCode = ' + iAvailabilityCode + 
														'\niOverbookCode = ' + iOverbookCode); 
												*/
											} else {
												sPreReqFails = "";
												sUnAvailFails = "";
												sOverlapFails = "";
												sOverBookFails = "";
											}

											if (iOverlapCode == 1) {
												if (sTransferErrorMsg.length > 0) sTransferErrorMsg = sTransferErrorMsg + "\n";
												if (sOverlapFails.length == 0) sTransferErrorMsg = sTransferErrorMsg + "This delegate is already booked on a course that overlaps with the selected course. \n";
												if (sOverlapFails.length > 0) sTransferErrorMsg = sTransferErrorMsg + "These delegates are already booked on a course that overlaps with the selected course. \n" + sOverlapFails + "\n";
											} else if (iOverlapCode == 2) {
												if (sTransferWarningMsg.length > 0) sTransferWarningMsg = sTransferWarningMsg + "\n";
												if (sOverlapFails.length == 0) sTransferWarningMsg = sTransferWarningMsg + "This delegate is already booked on a course that overlaps with the selected course. \n";
												if (sOverlapFails.length > 0) sTransferWarningMsg = sTransferWarningMsg + "These delegates are booked on a course that overlaps with the selected course. \n" + sOverlapFails + "\n";
											}

											if (iPreReqCode == 1) {
												if (sTransferErrorMsg.length > 0) sTransferErrorMsg = sTransferErrorMsg + "\n";
												if (sPreReqFails.length == 0) sTransferErrorMsg = sTransferErrorMsg + "The delegate has not met the pre-requisites for the course. \n";
												if (sPreReqFails.length > 0) sTransferErrorMsg = sTransferErrorMsg + "These delegates have not met the pre-requisites for the course: \n" + sPreReqFails + "\n";
											} else if (iPreReqCode == 2) {
												if (sTransferWarningMsg.length > 0) sTransferWarningMsg = sTransferWarningMsg + "\n";
												if (sPreReqFails.length == 0) sTransferWarningMsg = sTransferWarningMsg + "The delegate has not met the pre-requisites for the course. \n";
												if (sPreReqFails.length > 0) sTransferWarningMsg = sTransferWarningMsg + "These delegates have not met the pre-requisites for the course:  \n" + sPreReqFails + "\n";
											}

											if (iAvailabilityCode == 1) {
												if (sTransferErrorMsg.length > 0) sTransferErrorMsg = sTransferErrorMsg + "\n";
												if (sUnAvailFails.length == 0) sTransferErrorMsg = sTransferErrorMsg + "This delegate is unavailable for the selected course. \n";
												if (sUnAvailFails.length > 0) sTransferErrorMsg = sTransferErrorMsg + "These delegates are unavailable for the selected course. \n" + sUnAvailFails + "\n";
											} else if (iAvailabilityCode == 2) {
												if (sTransferWarningMsg.length > 0) sTransferWarningMsg = sTransferWarningMsg + "\n";
												if (sUnAvailFails.length == 0) sTransferWarningMsg = sTransferWarningMsg + "This delegate is unavailable for the selected course. \n";
												if (sUnAvailFails.length > 0) sTransferWarningMsg = sTransferWarningMsg + "These delegates are unavailable for the selected course. \n" + sUnAvailFails + "\n";
											}

											if (iOverbookCode == 1) {
												if (sTransferErrorMsg.length > 0) sTransferErrorMsg = sTransferErrorMsg + "\n";
												sTransferErrorMsg = sTransferErrorMsg + "The selected course is already fully booked.";
											} else if (iOverbookCode == 2) {
												if (sTransferWarningMsg.length > 0) sTransferWarningMsg = sTransferWarningMsg + "\n";
												sTransferWarningMsg = sTransferWarningMsg + "The selected course is already fully booked.";
											}

											if (sTransferErrorMsg.length > 0) {
												/* Error - not over-ridable. */
												if ((txtAction.value == "SELECTBOOKCOURSE_2") ||
														(txtAction.value == "SELECTTRANSFERBOOKING_1") ||
														(txtAction.value == "SELECTADDFROMWAITINGLIST_2")) {
													sTransferErrorMsg = sTransferErrorMsg + "\n\nUnable to make the booking."
												} else {
													if (txtAction.value == "SELECTBULKBOOKINGS") {
														sTransferErrorMsg = sTransferErrorMsg + "\n\nUnable to make the bookings."
													} else {
														sTransferErrorMsg = sTransferErrorMsg + "\n\nUnable to transfer the bookings."
													}
												}
												OpenHR.messageBox(sTransferErrorMsg);                            
												fTransferOK = false;
                            
											} else if (sTransferWarningMsg.length > 0) {
												/* Error - but over-ridable. */
												if ((txtAction.value == "SELECTBOOKCOURSE_2") ||
														(txtAction.value == "SELECTTRANSFERBOOKING_1") ||
														(txtAction.value == "SELECTADDFROMWAITINGLIST_2")) {
													sTransferWarningMsg = sTransferWarningMsg + "\n\nDo you still want to make the booking ?"
												} else {
													if (txtAction.value == "SELECTBULKBOOKINGS") {
														sTransferWarningMsg = sTransferWarningMsg + "\n\nDo you still want to make the bookings ?"
													} else {
														sTransferWarningMsg = sTransferWarningMsg + "\n\nDo you still want to transfer the bookings ?"
													}
												}
												iResponse = OpenHR.messageBox(sTransferWarningMsg, 36); // 36 = vbYesNo + vbQuestion

												if (iResponse == 7) { // 7 = vbNo
													fTransferOK = false;
												}
											}
										}

										if (txtAction.value == "SELECTBOOKCOURSE_2") {
											if (fTransferOK == true) {
												// Go ahead and book the course.
												var optionDataForm = OpenHR.getForm("optiondataframe","frmGetOptionData");
												optionDataForm.txtOptionAction.value = "SELECTBOOKCOURSE_3";
												optionDataForm.txtOptionRecordID.value = txtRecordID.value;
												optionDataForm.txtOptionLinkRecordID.value = txtLinkRecordID.value;
												optionDataForm.txtOptionValue.value = txtValue.value;

												window.parent.frames("optiondataframe").refreshOptionData();
											}
										} else {
											if (txtAction.value == "SELECTTRANSFERBOOKING_1") {
												if (fTransferOK == true) {
													// Go ahead and book the course.
													var optionDataForm = OpenHR.getForm("optiondataframe","frmGetOptionData");
													optionDataForm.txtOptionAction.value = "SELECTTRANSFERBOOKING_2";
													optionDataForm.txtOptionRecordID.value = txtRecordID.value;
													optionDataForm.txtOptionLinkRecordID.value = txtLinkRecordID.value;

													OpenHR.getFrame("optiondataframe").refreshOptionData();
												}
											} else {
												if (txtAction.value == "SELECTADDFROMWAITINGLIST_2") {
													if (fTransferOK == true) {
														// Go ahead and book the course.
														var optionDataForm = window.parent.frames("optiondataframe").document.forms("frmGetOptionData");
														optionDataForm.txtOptionAction.value = "SELECTADDFROMWAITINGLIST_3";
														optionDataForm.txtOptionRecordID.value = txtRecordID.value;
														optionDataForm.txtOptionLinkRecordID.value = txtLinkRecordID.value;
														optionDataForm.txtOptionValue.value = txtValue.value;

														OpenHR.getFrame("optiondataframe").refreshOptionData();
													}
												} else {
													if (txtAction.value == "SELECTBULKBOOKINGS") {
														if (fTransferOK == true) {
															// Go ahead and make the bookings.
															var optionDataForm = window.parent.frames("optiondataframe").document.forms("frmGetOptionData");
															optionDataForm.txtOptionAction.value = "SELECTBULKBOOKINGS_2";
															optionDataForm.txtOptionRecordID.value = txtRecordID.value;
															optionDataForm.txtOptionLinkRecordID.value = txtLinkRecordID.value;
															optionDataForm.txtOptionValue.value = txtValue.value;

															OpenHR.getFrame("optiondataframe").refreshOptionData();
														}
													} else {
														if (fTransferOK == true) {
															menu_transferCourse(txtLinkRecordID.value, true);
														} else {
															menu_transferCourse(0, true);                                        
														}
													}
												}
											}
										}
									}

									if (fNoAction == true) {
										$("#optionframe").attr("data-framesource", "EMPTYOPTION");

										// Get menu.asp to refresh the menu.
										menu_refreshMenu();
									}

            	// Fault 3503
									if (sCurrentWorkFramePage == "RECORDEDIT") {
										//OpenHR.getForm("workframe","frmRecordEditForm").ctlRecordEdit.refreshSize();
									}
								}
							}
						}
</script>



<%
	Response.Write("<INPUT type='hidden' id=txtAction name=txtAction value=""" & Replace(Session("optionAction"), """", "&quot;") & """>" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtErrorMessage name=txtErrorMessage value=""" & Replace(Session("errorMessage"), """", "&quot;") & """>" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtFromDef name=txtFromDef value=""" & Replace(Session("fromDef"), """", "&quot;") & """>" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtOrderID name=txtOrderID value=" & Session("orderID") & ">" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtFilterSQL name=txtFilterSQL value=""" & Replace(Session("optionFilterSQL"), """", "&quot;") & """>" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtFilterDef name=txtFilterDef value=""" & Replace(Session("optionFilterDef"), """", "&quot;") & """>" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtRecordID name=txtRecordID value=" & Session("optionRecordID") & ">" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtLinkRecordID name=txtLinkRecordID value=" & Session("optionLinkRecordID") & ">" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtLookupColumnID name=txtLookupColumnID value=" & Session("optionLookupColumnID") & ">" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtColumnID name=txtColumnID value=" & Session("optionColumnID") & ">" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtValue name=txtValue value=""" & Replace(Session("optionLookupValue"), """", "&quot;") & """>" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtFile name=txtFile value=""" & Replace(Session("optionFile"), """", "&quot;") & """>" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtResultCode name=txtResultCode value=" & Session("TBResultCode") & ">" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtPreReqFails name=txtPreReqFails value=""" & Replace(Session("PreReqFails"), """", "&quot;") & """>" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtUnAvailFails name=txtUnAvailFails value=""" & Replace(Session("UnAvailFails"), """", "&quot;") & """>" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtOverlapFails name=txtOverlapFails value=""" & Replace(Session("OverlapFails"), """", "&quot;") & """>" & vbCrLf)
	Response.Write("<INPUT type='hidden' id=txtOverBookFails name=txtOverBookFails value=""" & Replace(Session("OverBookFails"), """", "&quot;") & """>" & vbCrLf)
%>

<form id="frmEmptyOption" name="frmEmptyOption">
	<%	
		Dim cmdLinkValues
		Dim prmChildScreenID
		Dim prmTableID
		Dim prmRecordID
		Dim rstLinkValues
		Dim sErrorDescription As String
    
    
		If Session("optionAction") = "SELECTLINK" Then
			' Get the parent fields for the selected link.
			cmdLinkValues = CreateObject("ADODB.Command")
			cmdLinkValues.CommandText = "sp_ASRIntGetLinkParentValues"
			cmdLinkValues.CommandType = 4	' Stored Procedure
			cmdLinkValues.ActiveConnection = Session("databaseConnection")

			prmChildScreenID = cmdLinkValues.CreateParameter("childScreenID", 3, 1)
			cmdLinkValues.Parameters.Append(prmChildScreenID)
			prmChildScreenID.value = CleanNumeric(Session("optionScreenID"))

			prmTableID = cmdLinkValues.CreateParameter("tableID", 3, 1)
			cmdLinkValues.Parameters.Append(prmTableID)
			prmTableID.value = CleanNumeric(Session("optionLinkTableID"))

			prmRecordID = cmdLinkValues.CreateParameter("recordID", 3, 1)
			cmdLinkValues.Parameters.Append(prmRecordID)
			prmRecordID.value = CleanNumeric(Session("optionRecordID"))

			Err.Clear()
			rstLinkValues = cmdLinkValues.Execute

			If (Err.Number <> 0) Then
				sErrorDescription = "The link values could not be retrieved." & vbCrLf & FormatError(Err.Description)
			End If

			If Len(sErrorDescription) = 0 Then
				If Not (rstLinkValues.bof And rstLinkValues.eof) Then
					For iloop = 0 To (rstLinkValues.fields.count - 1)
						If IsDBNull(rstLinkValues.fields(iloop).value) Then
							Response.Write("<INPUT type='hidden' id=txtData_" & rstLinkValues.fields(iloop).name & " name=txtData_" & rstLinkValues.fields(iloop).name & " value="""">" & vbCrLf)
						Else
							Response.Write("<INPUT type='hidden' id=txtData_" & rstLinkValues.fields(iloop).name & " name=txtData_" & rstLinkValues.fields(iloop).name & " value=""" & Replace(rstLinkValues.fields(iloop).value, """", "&quot;") & """>" & vbCrLf)
						End If
					Next
				End If

				'	Release the ADO recordset object.
				rstLinkValues.close()
			End If
			
			rstLinkValues = Nothing

			' Release the ADO command object.
			cmdLinkValues = Nothing
		End If
	
		Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>")
	%>
</form>

<form action="emptyoption_Submit" method="post" id="frmGotoOption" name="frmGotoOption">
	<%Html.RenderPartial("~/Views/Shared/gotoOption.ascx")%>
</form>





<script type="text/javascript">
	emptyoption_onload();    
</script>

