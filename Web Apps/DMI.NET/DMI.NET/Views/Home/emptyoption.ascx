<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<%

	If Len(Session("timestamp")) = 0 Then
		Session("timestamp") = 0
	End If

%>

<script type="text/javascript">
	function emptyoption_onload() {		
		var fNoAction;
		var sCurrentWorkFramePage = $("#workframe").attr("data-framesource"); //OpenHR.currentWorkPage();
		var frmMenu = OpenHR.getForm("menuframe", "frmMenuInfo");

		// Do nothing if the menu controls are not yet instantiated.
		if (frmMenu != null) {
			if (OpenHR.currentWorkPage() != "DEFAULT") {
				fNoAction = true;				
					var div = document.getElementById("emptyoption_vars");
					var txtAction = div.querySelector("#txtAction");
					var txtFromDef = div.querySelector("#txtFromDef");
					var txtOrderID = div.querySelector("#txtOrderID");
					var txtFilterSQL = div.querySelector("#txtFilterSQL");
					var txtFilterDef = div.querySelector("#txtFilterDef");
					var txtRecordID = div.querySelector("#txtRecordID");
					var txtColumnID = div.querySelector("#txtColumnID");
					var txtValue = div.querySelector("#txtValue");
					var txtFile = div.querySelector("#txtFile");
					var txtFileValue = div.querySelector("#txtFileValue");
					var txtResultCode = div.querySelector("#txtResultCode");
					var txtPreReqFailsCount = div.querySelector("#txtPreReqFails");
					var txtUnAvailFailsCount = div.querySelector("#txtUnAvailFails");
					var txtOverlapFailsCount = div.querySelector("#txtOverlapFails");
					var txtCourseOverbooked = div.querySelector("#txtOverBooked");
					var txtLinkRecordID = div.querySelector("#txtLinkRecordID");
					var txtErrorMessage = div.querySelector("#txtErrorMessage");
				
				if (txtAction.value == "SELECTORDER") {
					fNoAction = false;

					if (sCurrentWorkFramePage == "RECORDEDIT") {						
						var frmRecEdit = OpenHR.getForm("workframe","frmRecordEditForm");
						frmRecEdit.txtRecEditFromDef.value = txtFromDef.value;
						frmRecEdit.txtCurrentOrderID.value = txtOrderID.value;
						refreshData(); 	//recedit
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
						frmRecEdit = OpenHR.getForm("workframe","frmRecordEditForm");
						frmRecEdit.txtRecEditFilterSQL.value = txtFilterSQL.value;
						frmRecEdit.txtRecEditFilterDef.value = txtFilterDef.value;
						$("#workframe").attr("data-framesource", "RECORDEDIT");
						$("#optionframe").hide();
						$("#workframe").show();
						refreshData(); 	//recedit
					} else {
						if (sCurrentWorkFramePage == "FIND") {
							frmFind = OpenHR.getForm("workframe","frmFindForm");
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
					var recEditControl;
					if (txtAction.value == "SELECTLINK") {
					fNoAction = false;
					recEditControl = OpenHR.getForm("workframe", "frmRecordEditForm").ctlRecordEdit;
						
					var sControlName;
					var sColumnID;
					var dataCollection = window.frmEmptyOption.elements;
					if (dataCollection != null) {
						// Need to hide the popup in case setdata causes
						// the intrecedit control to display an error message.
						menu_refreshMenu();

						for (var i = 0; i < dataCollection.length; i++) {
							sControlName = dataCollection.item(i).name;
							sControlName = sControlName.substr(0, 8);
							if (sControlName == "txtData_") {
								sColumnID = dataCollection.item(i).name;
								sColumnID = sColumnID.substr(8);
								recEdit_setData(sColumnID, dataCollection.item(i).value);
							}
						}
						enableSaveButton(); //recedit
					}

					$("#optionframe").attr("data-framesource", "EMPTYOPTION");
					//window.setTimeout("window.parent.frames('menuframe').refreshMenu()", 100);
				}

				if (txtAction.value == "SELECTLOOKUP") {
					fNoAction = false;
					recEditControl = OpenHR.getForm("workframe","frmRecordEditForm").ctlRecordEdit;
					recEdit_setData(txtColumnID.value, txtValue.value);
					enableSaveButton(); //recedit

					$("#optionframe").attr("data-framesource", "EMPTYOPTION");
					//window.setTimeout("window.parent.frames('menuframe').refreshMenu()", 100);
					menu_refreshMenu();
				}

				if ((txtAction.value == "SELECTIMAGE") || (txtAction.value == "SELECTOLE")) {
					fNoAction = false;
					recEditControl = OpenHR.getForm("workframe", "frmRecordEditForm").ctlRecordEdit;
						recEdit_setData(txtColumnID.value, txtFile.value);
						recEdit_ChangedOLEPhoto(txtColumnID.value, "");
					
						enableSaveButton(); //in scope.				
					
					$("#optionframe").attr("data-framesource", "EMPTYOPTION");
					//window.setTimeout("window.parent.frames('menuframe').refreshMenu()", 100);
					menu_refreshMenu();
				}

				if (txtAction.value == "LINKOLE") {
					if (txtErrorMessage.value == "") {
						fNoAction = false;
						recEditControl = OpenHR.getForm("workframe", "frmRecordEditForm").ctlRecordEdit;
						if (txtFileValue.value.length > 0) {
							// set the new photo value if applicable.
							$('#txtData_' + txtColumnID.value).attr('data-Img', txtFileValue.value);
						}

						recEdit_setData(txtColumnID.value, txtFile.value);

						$("#txtRecEditTimeStamp").val("<%=session("timestamp")%>");
						$("#optionframe").attr("data-framesource", "EMPTYOPTION");

						//Update the ID badge picture
						$("#UserPicture").attr("src", "<%=Session("SelfServicePhotograph_src")%>");

						menu_refreshMenu();
					}
				}


				if ((txtAction.value == "SELECTTRANSFERCOURSE") ||
						(txtAction.value == "SELECTBOOKCOURSE_2") ||
						(txtAction.value == "SELECTTRANSFERBOOKING_1") ||
						(txtAction.value == "SELECTADDFROMWAITINGLIST_2") ||
						(txtAction.value == "SELECTBULKBOOKINGS")) {
						var sPrefix;
						var sPrefix2;
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
					$("#optionframe").hide();
					$("#workframe").show();

					//menu_refreshMenu();

					var fTransferOK = true;

					if (txtResultCode.value.indexOf("\\") > 0) {
						// -------------  Results come from sp_ASRIntValidateBulkBookings, NOT sp_ASRIntValidateTransfers --------------
						var messageOverlapSingular = "This delegate is already booked on a course that overlaps with the selected course. \n";
						var messageOverlapPlural = "These delegates are already booked on a course that overlaps with the selected course. \n";
						var messagePrerequisitesSingular = "The delegate has not met the pre-requisites for the course. \n";
						var messagePrerequisitesPlural = "These delegates have not met the pre-requisites for the course. \n";
						var messageUnavailableSingular = "This delegate is unavailable for the selected course. \n";
						var messageUnavailablePlural = "These delegates are unavailable for the selected course. \n";
						//var OverlapCode;
						//var AvailabilityCode;
						//var PreReqCode;
						//var sOverBookFails;
						//var sPreReqFails;
						//var sUnAvailFails;
						//var sOverlapFails;
						var EmployeeName;
						var ResultCode;
						var ResultCodes = txtResultCode.value;
						var CourseOverbooked = txtCourseOverbooked.value;
						var EmployeesWithOverlapError = [];
						var EmployeesWithOverlapWarning = [];
						var EmployeesWithPreReqError = [];
						var EmployeesWithPreReqWarning = [];
						var EmployeesWithUnAvailError = [];
						var EmployeesWithUnAvailWarning = [];
						var j;
						var sTransferErrorMsg = "";
						var sTransferWarningMsg = "";

						if (ResultCodes.length > 0 && ResultCodes != 0) {
							if (ResultCodes.indexOf("|") != -1) { // Multiple ResultCodes returned, we need to parse them (Results come from sp_ASRIntValidateBulkBookings)
								ResultCodes = ResultCodes.split("|");

								//Loop over the results
								for (j = 0; j <= ResultCodes.length - 1; j++) {
									var EmployeeAndCode = ResultCodes[j].split("\\");
									EmployeeName = EmployeeAndCode[0];
									ResultCode = EmployeeAndCode[1];

									if (ResultCode[0] == 1) {
										EmployeesWithPreReqError.push(EmployeeName);
									} else if (ResultCode[0] == 2) {
										EmployeesWithPreReqWarning.push(EmployeeName);
									}

									if (ResultCode[2] == 1) {
										EmployeesWithOverlapError.push(EmployeeName);
									} else if (ResultCode[2] == 2) {
										EmployeesWithOverlapWarning.push(EmployeeName);
									}

									if (ResultCode[1] == 1) {
										EmployeesWithUnAvailError.push(EmployeeName);
									} else if (ResultCode[1] == 2) {
										EmployeesWithUnAvailWarning.push(EmployeeName);
									}
								}
							} else { // Single ResultCode returned, we need to parse it 														
								ResultCodes = ResultCodes.split("\\")[1];
								if (ResultCodes[0] == 1) {
									sTransferErrorMsg = messagePrerequisitesSingular + "\n";
								} else if (ResultCodes[0] == 2) {
									sTransferWarningMsg = messagePrerequisitesSingular + "\n";
								}

								if (ResultCodes[2] == 1) {
									sTransferErrorMsg = messageOverlapSingular + "\n";
								} else if (ResultCodes[2] == 2) {
									sTransferWarningMsg = messageOverlapSingular + "\n";
								}

								if (ResultCodes[1] == 1) {
									sTransferErrorMsg = messageUnavailableSingular + "\n";
								} else if (ResultCodes[1] == 2) {
									sTransferWarningMsg = messageUnavailableSingular + "\n";
								}
							}
						}

						if (CourseOverbooked == 1) {
							if (sTransferErrorMsg.length > 0) sTransferErrorMsg = sTransferErrorMsg + "\n";
							sTransferErrorMsg = sTransferErrorMsg + "The number of delegates selected would exceed the maximum number allowed on the course.";
						} else if (CourseOverbooked == 2) {
							if (sTransferWarningMsg.length > 0) sTransferWarningMsg = sTransferWarningMsg + "\n";
							sTransferWarningMsg = sTransferWarningMsg + "The number of delegates selected would exceed the maximum number allowed on the course.";
						}

						if (EmployeesWithPreReqError.length > 0) {
							if (EmployeesWithPreReqError.length == 1) sTransferErrorMsg = sTransferErrorMsg + messagePrerequisitesSingular + "\n";
							if (EmployeesWithPreReqError.length > 1) sTransferErrorMsg = sTransferErrorMsg + messagePrerequisitesPlural + "\n";
							for (j = 0; j <= EmployeesWithPreReqError.length - 1; j++) {
								sTransferErrorMsg += EmployeesWithPreReqError[j] + "\n";
							}
						}
						if (EmployeesWithPreReqWarning.length > 0) {
							if (EmployeesWithPreReqWarning.length == 1) sTransferWarningMsg = sTransferWarningMsg + messagePrerequisitesSingular + "\n";
							if (EmployeesWithPreReqWarning.length > 1) sTransferWarningMsg = sTransferWarningMsg + messagePrerequisitesPlural + "\n";
							for (j = 0; j <= EmployeesWithPreReqWarning.length - 1; j++) {
								sTransferWarningMsg += EmployeesWithPreReqWarning[j] + "\n";
							}
						}

						if (EmployeesWithOverlapError.length > 0) {
							if (EmployeesWithOverlapError.length == 1) sTransferErrorMsg = sTransferErrorMsg + messageOverlapSingular + "\n";
							if (EmployeesWithOverlapError.length > 1) sTransferErrorMsg = sTransferErrorMsg + messageOverlapPlural + "\n";
							for (j = 0; j <= EmployeesWithOverlapError.length - 1; j++) {
								sTransferErrorMsg += EmployeesWithOverlapError[j] + "\n";
							}
						}
						if (EmployeesWithOverlapWarning.length > 0) {
							if (EmployeesWithOverlapWarning.length == 1) sTransferWarningMsg = sTransferWarningMsg + messageOverlapSingular + "\n";
							if (EmployeesWithOverlapWarning.length > 1) sTransferWarningMsg = sTransferWarningMsg + messageOverlapPlural + "\n";
							for (j = 0; j <= EmployeesWithOverlapWarning.length - 1; j++) {
								sTransferWarningMsg += EmployeesWithOverlapWarning[j] + "\n";
							}
						}

						if (EmployeesWithUnAvailError.length > 0) {
							if (EmployeesWithUnAvailError.length == 1) sTransferErrorMsg = sTransferErrorMsg + messageUnavailableSingular + "\n";
							if (EmployeesWithUnAvailError.length > 1) sTransferErrorMsg = sTransferErrorMsg + messageUnavailablePlural + "\n";
							for (j = 0; j <= EmployeesWithUnAvailError.length - 1; j++) {
								sTransferErrorMsg += EmployeesWithUnAvailError[j] + "\n";
							}
						}
						if (EmployeesWithUnAvailWarning.length > 0) {
							if (EmployeesWithUnAvailWarning.length == 1) sTransferWarningMsg = sTransferWarningMsg + messageUnavailableSingular + "\n";
							if (EmployeesWithUnAvailWarning.length > 1) sTransferWarningMsg = sTransferWarningMsg + messageUnavailablePlural + "\n";
							for (j = 0; j <= EmployeesWithUnAvailWarning.length - 1; j++) {
								sTransferWarningMsg += EmployeesWithUnAvailWarning[j] + "\n";
							}
						}
						//-------------------------------------------------

					} else {
						// -------------  Results come from sp_ASRIntValidateTransfers, NOT sp_ASRIntValidateBulkBookings --------------
						
						var iResultCode = txtResultCode.value;
						//if (iResultCode > 0) {
						var iOverlapCode = iResultCode % 10;
						iResultCode = iResultCode - iOverlapCode;
						iResultCode = iResultCode / 10;
						var iAvailabilityCode = iResultCode % 10;
						iResultCode = iResultCode - iAvailabilityCode;
						iResultCode = iResultCode / 10;
						var iPreReqCode = iResultCode % 10;
						iResultCode = iResultCode - iPreReqCode;
						iResultCode = iResultCode / 10;
						var iOverbookCode = iResultCode;

						sTransferErrorMsg = "";
						sTransferWarningMsg = "";

						var sPreReqFails = "";
						var sUnAvailFails = "";
						var sOverlapFails = "";
						
						if (iOverlapCode == 1) {
							if (sTransferErrorMsg.length > 0) sTransferErrorMsg = sTransferErrorMsg + "\n";
							if (sOverlapFails.length == 0) sTransferErrorMsg = sTransferErrorMsg + "This delegate is already booked on a course that overlaps with the selected course. \n";
							if (sOverlapFails.length > 0) sTransferErrorMsg = sTransferErrorMsg + "These delegates are already booked on a course that overlaps with the selected course. \n" + sOverlapFails + "\n";
						}
						else if (iOverlapCode == 2) {
							if (sTransferWarningMsg.length > 0) sTransferWarningMsg = sTransferWarningMsg + "\n";
							if (sOverlapFails.length == 0) sTransferWarningMsg = sTransferWarningMsg + "This delegate is already booked on a course that overlaps with the selected course. \n";
							if (sOverlapFails.length > 0) sTransferWarningMsg = sTransferWarningMsg + "These delegates are booked on a course that overlaps with the selected course. \n" + sOverlapFails + "\n";
						}

						if (iPreReqCode == 1) {
							if (sTransferErrorMsg.length > 0) sTransferErrorMsg = sTransferErrorMsg + "\n";
							if (sPreReqFails.length == 0) sTransferErrorMsg = sTransferErrorMsg + "The delegate has not met the pre-requisites for the course. \n";
							if (sPreReqFails.length > 0) sTransferErrorMsg = sTransferErrorMsg + "These delegates have not met the pre-requisites for the course: \n" + sPreReqFails + "\n";
						}
						else if (iPreReqCode == 2) {
							if (sTransferWarningMsg.length > 0) sTransferWarningMsg = sTransferWarningMsg + "\n";
							if (sPreReqFails.length == 0) sTransferWarningMsg = sTransferWarningMsg + "The delegate has not met the pre-requisites for the course. \n";
							if (sPreReqFails.length > 0) sTransferWarningMsg = sTransferWarningMsg + "These delegates have not met the pre-requisites for the course:  \n" + sPreReqFails + "\n";
						}

						if (iAvailabilityCode == 1) {
							if (sTransferErrorMsg.length > 0) sTransferErrorMsg = sTransferErrorMsg + "\n";
							if (sUnAvailFails.length == 0) sTransferErrorMsg = sTransferErrorMsg + "This delegate is unavailable for the selected course. \n";
							if (sUnAvailFails.length > 0) sTransferErrorMsg = sTransferErrorMsg + "These delegates are unavailable for the selected course. \n" + sUnAvailFails + "\n";
						}
						else if (iAvailabilityCode == 2) {
							if (sTransferWarningMsg.length > 0) sTransferWarningMsg = sTransferWarningMsg + "\n";
							if (sUnAvailFails.length == 0) sTransferWarningMsg = sTransferWarningMsg + "This delegate is unavailable for the selected course. \n";
							if (sUnAvailFails.length > 0) sTransferWarningMsg = sTransferWarningMsg + "These delegates are unavailable for the selected course. \n" + sUnAvailFails + "\n";
						}

						if (iOverbookCode == 1) {
							if (sTransferErrorMsg.length > 0) sTransferErrorMsg = sTransferErrorMsg + "\n";
							sTransferErrorMsg = sTransferErrorMsg + "The selected course is already fully booked.";
						}
						else if (iOverbookCode == 2) {
							if (sTransferWarningMsg.length > 0) sTransferWarningMsg = sTransferWarningMsg + "\n";
							sTransferWarningMsg = sTransferWarningMsg + "The selected course is already fully booked.";
						}

					}

					if (sTransferErrorMsg.length > 0) {
							/* Error - not over-ridable. */
							if ((txtAction.value == "SELECTBOOKCOURSE_2") ||
									(txtAction.value == "SELECTTRANSFERBOOKING_1") ||
									(txtAction.value == "SELECTADDFROMWAITINGLIST_2")) {
								sTransferErrorMsg = sTransferErrorMsg + "\n\nUnable to make the booking.";
							} else {
								if (txtAction.value == "SELECTBULKBOOKINGS") {
									sTransferErrorMsg = sTransferErrorMsg + "\n\nUnable to make the bookings.";
								} else {
									sTransferErrorMsg = sTransferErrorMsg + "\n\nUnable to transfer the bookings.";
								}
							}
							OpenHR.messageBox(sTransferErrorMsg);                            
							fTransferOK = false;
														
						} else if (sTransferWarningMsg.length > 0) {
							/* Error - but over-ridable. */
							if ((txtAction.value == "SELECTBOOKCOURSE_2") ||
									(txtAction.value == "SELECTTRANSFERBOOKING_1") ||
									(txtAction.value == "SELECTADDFROMWAITINGLIST_2")) {
								sTransferWarningMsg = sTransferWarningMsg + "\nDo you still want to make the booking ?";
							} else {
								if (txtAction.value == "SELECTBULKBOOKINGS") {
									sTransferWarningMsg = sTransferWarningMsg + "\nDo you still want to make the bookings ?";
								} else {
									sTransferWarningMsg = sTransferWarningMsg + "\nDo you still want to transfer the bookings ?";
								}
							}
							var iResponse = OpenHR.messageBox(sTransferWarningMsg, 36); // 36 = vbYesNo + vbQuestion

							if (iResponse == 7) { // 7 = vbNo
								fTransferOK = false;
							}
						}
					
						var optionDataForm;
						if (txtAction.value == "SELECTBOOKCOURSE_2") {
						if (fTransferOK == true) {
							// Go ahead and book the course.
								optionDataForm = OpenHR.getForm("optiondataframe","frmGetOptionData");
								optionDataForm.txtOptionAction.value = "SELECTBOOKCOURSE_3";
							optionDataForm.txtOptionRecordID.value = txtRecordID.value;
							optionDataForm.txtOptionLinkRecordID.value = txtLinkRecordID.value;
							optionDataForm.txtOptionValue.value = txtValue.value;

							refreshOptionData(); //is in scope and unique anyhoo.
						}
					} else {
						if (txtAction.value == "SELECTTRANSFERBOOKING_1") {
							if (fTransferOK == true) {
								// Go ahead and book the course.
									optionDataForm = OpenHR.getForm("optiondataframe","frmGetOptionData");
									optionDataForm.txtOptionAction.value = "SELECTTRANSFERBOOKING_2";
								optionDataForm.txtOptionRecordID.value = txtRecordID.value;
								optionDataForm.txtOptionLinkRecordID.value = txtLinkRecordID.value;

								refreshOptionData();
							}
						} else {
							if (txtAction.value == "SELECTADDFROMWAITINGLIST_2") {
								if (fTransferOK == true) {
									// Go ahead and book the course.
										optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
										optionDataForm.txtOptionAction.value = "SELECTADDFROMWAITINGLIST_3";
									optionDataForm.txtOptionRecordID.value = txtRecordID.value;
									optionDataForm.txtOptionLinkRecordID.value = txtLinkRecordID.value;
									optionDataForm.txtOptionValue.value = txtValue.value;
									
									refreshOptionData(); //should be in scope!
								}
							} else {
								if (txtAction.value == "SELECTBULKBOOKINGS") {
									if (fTransferOK == true) {
										// Go ahead and make the bookings.
											optionDataForm = OpenHR.getForm("optiondataframe", "frmGetOptionData");
											optionDataForm.txtOptionAction.value = "SELECTBULKBOOKINGS_2";
										optionDataForm.txtOptionRecordID.value = txtRecordID.value;
										optionDataForm.txtOptionLinkRecordID.value = txtLinkRecordID.value;
										optionDataForm.txtOptionValue.value = txtValue.value;

										refreshOptionData();
									}
								} else {
									if (fTransferOK == true) {
										menu_transferCourse($("#txtLinkRecordID").val(), true);
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
					menu_refreshMenu(); //A second call to menu_RefreshMenu fixes the problem reported in the notes by Craig in Jira http://tcjira01:8080/browse/HRPRO-3140; don't ask me why it fixes it, it just does!
				}
				
				menu_refreshMenu();

				if (sCurrentWorkFramePage == "RECORDEDIT") {
					//display any errors that may have occurred:
					if (txtErrorMessage.value.length > 0) alert(txtErrorMessage.value);
				}
			}
		}
	}
</script>


<%
	Response.Write("<div id='emptyoption_vars'>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtAction' name='txtAction' value='" & Replace(Session("optionAction"), """", "&quot;") & "'>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtErrorMessage' name='txtErrorMessage' value='" & Replace(Session("errorMessage"), """", "&quot;") & "'>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtFromDef' name='txtFromDef' value='" & Replace(Session("fromDef"), """", "&quot;") & "'>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtOrderID' name='txtOrderID' value='" & Session("orderID") & "'>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtFilterSQL' name='txtFilterSQL' value='" & Replace(Session("optionFilterSQL"), """", "&quot;") & "'>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtFilterDef' name='txtFilterDef' value='" & Replace(Replace(Session("optionFilterDef"), """", "&quot;"), "'", "&#39;") & "'>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtRecordID' name='txtRecordID' value='" & Session("optionRecordID") & "'>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtLinkRecordID' name='txtLinkRecordID' value='" & Session("optionLinkRecordID") & "'>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtLookupColumnID' name='txtLookupColumnID' value='" & Session("optionLookupColumnID") & "'>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtColumnID' name='txtColumnID' value='" & Session("optionColumnID") & "'>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtValue' name='txtValue' value='" & Replace(Session("optionLookupValue"), """", "&quot;") & "'>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtFile' name='txtFile' value='" & Replace(Session("optionFile"), """", "&quot;") & "'>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtFileValue' name='txtFileValue' value='" & Replace(Session("optionFileValue"), """", "&quot;") & "'>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtResultCode' name='txtResultCode' value='" & Replace(Session("TBResultCode"), "'", "&#39;") & "'>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtPreReqFails' name='txtPreReqFails' value='" & Replace(Session("PreReqFails"), """", "&quot;") & "'>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtUnAvailFails' name='txtUnAvailFails' value='" & Replace(Session("UnAvailFails"), """", "&quot;") & "'>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtOverlapFails' name='txtOverlapFails' value='" & Replace(Session("OverlapFails"), """", "&quot;") & "'>" & vbCrLf)
	Response.Write("<input type='hidden' id='txtOverBooked' name='txtOverBooked' value='" & Replace(Session("Overbooked"), """", "&quot;") & "'>" & vbCrLf)
	Response.Write("</div>" & vbCrLf)
%>

<form id="frmEmptyOption" name="frmEmptyOption">
	<%
		
		Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
		Dim sErrorDescription As String = ""
			
		If Session("optionAction") = "SELECTLINK" Then
			
			Try
			
				Dim rstLinkValues = objDatabase.DB.GetDataTable("sp_ASRIntGetLinkParentValues", CommandType.StoredProcedure _
					, New SqlParameter("piChildScreenID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionScreenID"))} _
					, New SqlParameter("piTableID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionLinkTableID"))} _
					, New SqlParameter("piRecordID", SqlDbType.Int) With {.Value = CleanNumeric(Session("optionRecordID"))})

				If Not rstLinkValues Is Nothing Then
					For iloop = 0 To (rstLinkValues.Columns.Count - 1)
						If IsDBNull(rstLinkValues(0)(iloop)) Then
							Response.Write("<input type='hidden' id='txtData_" & rstLinkValues.Columns(iloop).ColumnName & "' name='txtData_" & rstLinkValues.Columns(iloop).ColumnName & "' value=''>" & vbCrLf)
						Else
							Response.Write("<input type='hidden' id='txtData_" & rstLinkValues.Columns(iloop).ColumnName & "' name='txtData_" & rstLinkValues.Columns(iloop).ColumnName & "' value='" & Replace(rstLinkValues(0)(iloop).ToString(), """", "&quot;") & "'>" & vbCrLf)
						End If
					Next
				End If

			Catch ex As Exception
				sErrorDescription = "The link values could not be retrieved." & vbCrLf & FormatError(Err.Description)

			End Try

		End If
	
		Response.Write("<input type='hidden' id='txtErrorDescription' name='txtErrorDescription' value='" & sErrorDescription & "'>")
	%>
</form>

<form action="emptyoption_Submit" method="post" id="frmGotoOption" name="frmGotoOption">
	<%Html.RenderPartial("~/Views/Shared/gotoOption.ascx")%>
	<%=Html.AntiForgeryToken()%>
</form>

<script type="text/javascript">
	emptyoption_onload();    
</script>
