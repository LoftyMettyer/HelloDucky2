<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>

<!DOCTYPE html>

<html>
<head runat="server">

	<title>OpenHR</title>
	<script src="<%: Url.LatestContent("~/bundles/jQuery")%>" type="text/javascript"></script>
	<script src="<%: Url.LatestContent("~/bundles/jQueryUI7")%>" type="text/javascript"></script>
	<script src="<%: Url.LatestContent("~/bundles/OpenHR_General")%>" type="text/javascript"></script>


	<%--Here's the stylesheets for the font-icons displayed on the dashboard for wireframe and tile layouts--%>
	<link href="<%= Url.LatestContent("~/Content/general_enclosed_foundicons.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/font-awesome.min.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/fonts/SSI80v194934/style.css")%>" rel="stylesheet" />

	<%--Base stylesheets--%>
	<link href="<%: Url.LatestContent("~/Content/Site.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />

	<%--stylesheet for slide-out dmi menu--%>
	<link href="<%: Url.LatestContent("~/Content/contextmenustyle.css")%>" rel="stylesheet" type="text/css" />

	<%--ThemeRoller stylesheet--%>
	<link id="DMIthemeLink" href="<%: Url.LatestContent("~/Content/themes/" & Session("ui-admin-theme").ToString() & "/jquery-ui.min.css")%>" rel="stylesheet" type="text/css" />

	<%--jQuery Grid Stylesheet--%>
	<link href="<%: Url.LatestContent("~/Content/ui.jqgrid.css")%>" rel="stylesheet" type="text/css" />

</head>
<body>

	<div>

	<div id="popout_Wrapper">

		<form id="frmEventDetails" name="frmEventDetails">

			<%
				
				Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)

				Dim i As Integer
				Dim sValue As String
			
				Session("eventName") = Request("txtEventName")
				Session("eventID") = Request("txtEventID")
				Session("cboString") = vbNullString

				If Request("txtEventMode") = "Batch" Then
					Session("eventBatch") = True
					Response.Write("<input type='hidden' Name='txtEventBatch' ID='txtEventBatch' value='1'>" & vbCrLf)
				Else
					Session("eventBatch") = False
					Response.Write("<input type='hidden' Name='txtEventBatch' ID='txtEventBatch' value='0'>" & vbCrLf)
				End If
				
				Dim rsAllBatchJobs = objDataAccess.GetDataTable("spASRIntGetEventLogBatchDetails", CommandType.StoredProcedure _
					, New SqlParameter("piBatchRunID", SqlDbType.Int) With {.Value = CleanNumeric(Request("txtEventBatchRunID"))} _
					, New SqlParameter("piEventID", SqlDbType.Int) With {.Value = CleanNumeric(Request("txtEventID"))})
		
				
				With rsAllBatchJobs
					
					i = 0
					For Each objRow As DataRow In .Rows

						i += 1

						Response.Write("<input type='hidden' Name='txtEventID_" & objRow("ID") & "' id='txtEventID_" & objRow("ID") & "' value='" & Replace(objRow("ID"), """", "&quot;") & "'>" & vbCrLf)
				
						sValue = objRow("Name").ToString()											'original value
						sValue = Replace(sValue, """", "&quot;")		'escape quotes
						sValue = Replace(sValue, "<", "&lt;")						'escape left angle bracket
						sValue = Replace(sValue, ">", "&gt;")						'escape right angle bracket
				
						Response.Write("<input type='hidden' Name='txtEventName_" & objRow("ID") & "' id='txtEventName_" & objRow("ID") & "' value='" & sValue & "'>" & vbCrLf)
						Response.Write("<input type='hidden' Name='txtEventMode_" & objRow("ID") & "' id='txtEventMode_" & objRow("ID") & "' value='" & Replace(objRow("Mode"), """", "&quot;") & "'>" & vbCrLf)
				
						Response.Write("<input type='hidden' Name='txtEventStartTime_" & objRow("ID") & "' id='txtEventStartTime_" & objRow("ID") & "' value='" & ConvertSQLDateToLocale(objRow("DateTime")) & " " & ConvertSqlDateToTime(objRow("DateTime")) & "'>" & vbCrLf)
				
						If IsDBNull(objRow("EndTime")) Then
							Response.Write("<input type='hidden' Name='txtEventEndTime_" & objRow("ID") & "' id='txtEventEndTime_" & objRow("ID") & "' value='" & vbNullString & "'>" & vbCrLf)
						Else
							Response.Write("<input type='hidden' Name='txtEventEndTime_" & objRow("ID") & "' id='txtEventEndTime_" & objRow("ID") & "' value='" & ConvertSQLDateToLocale(objRow("EndTime")) & " " & ConvertSqlDateToTime(objRow("EndTime")) & "'>" & vbCrLf)
						End If
				
						Response.Write("<input type='hidden' Name='txtEventDuration_" & objRow("ID") & "' id='txtEventDuration_" & objRow("ID") & "' value='" & FormatEventDuration(CLng(objRow("Duration"))) & "'>" & vbCrLf)

						Response.Write("<input type='hidden' Name='txtEventType_" & objRow("ID") & "' id='txtEventType_" & objRow("ID") & "' value='" & Replace(objRow("Type"), """", "&quot;") & "'>" & vbCrLf)
						Response.Write("<input type='hidden' Name='txtEventStatus_" & objRow("ID") & "' id='txtEventStatus_" & objRow("ID") & "' value='" & Replace(objRow("Status"), """", "&quot;") & "'>" & vbCrLf)
						Response.Write("<input type='hidden' Name='txtEventUser_" & objRow("ID") & "' id='txtEventUser_" & objRow("ID") & "' value='" & Replace(objRow("Username"), """", "&quot;") & "'>" & vbCrLf)
				
						Response.Write("<input type='hidden' Name='txtEventSuccessCount_" & objRow("ID") & "' id='txtEventSuccessCount_" & objRow("ID") & "' value='" & Replace(objRow("SuccessCount"), """", "&quot;") & "'>" & vbCrLf)
						Response.Write("<input type='hidden' Name='txtEventFailCount_" & objRow("ID") & "' id='txtEventFailCount_" & objRow("ID") & "' value='" & Replace(objRow("FailCount"), """", "&quot;") & "'>" & vbCrLf)
				
						Response.Write("<input type='hidden' Name='txtEventBatchRunID_" & objRow("ID") & "' id='txtEventBatchRunID_" & objRow("ID") & "' value='" & Replace(objRow("BatchRunID"), """", "&quot;") & "'>" & vbCrLf)
						Response.Write("<input type='hidden' Name='txtEventBatchName_" & objRow("ID") & "' id='txtEventBatchName_" & objRow("ID") & "' value='" & Replace(objRow("BatchName"), """", "&quot;") & "'>" & vbCrLf)
						Response.Write("<input type='hidden' Name='txtEventBatchJobID_" & objRow("ID") & "' id='txtEventBatchJobID_" & objRow("ID") & "' value='" & Replace(objRow("BatchJobID"), """", "&quot;") & "'>" & vbCrLf)
				
						If Session("eventBatch") = True Then
							If Session("eventID") = objRow("ID") Then
								Session("cboString") = Session("cboString") & "<option selected='selected' name='" & objRow("Name") & "' value='" & objRow("ID") & "'>" & objRow("Type") & " - " & objRow("Name") & "</option>" & vbCrLf
							Else
								Session("cboString") = Session("cboString") & "<option name='" & objRow("Name") & "' value='" & objRow("ID") & "'>" & objRow("Type") & " - " & objRow("Name") & "</option>" & vbCrLf
							End If
						End If
				
					Next
					Session("eventBatchName") = Request("txtEventBatchName")
			
					Session("cboString") = Session("cboString") & "</SELECT>" & vbCrLf
					If i <= 1 Then
						Session("cboString") = "<select disabled id=cboOtherJobs name=cboOtherJobs class=""combodisabled"" style=""WIDTH: 100%"" onchange='populateEventInfo();populateEventDetails();'>" & vbCrLf & Session("cboString")
					Else
						Session("cboString") = "<select id=cboOtherJobs name=cboOtherJobs class=""combo"" style=""WIDTH: 100%"" onchange='populateEventInfo();populateEventDetails();'>" & vbCrLf & Session("cboString")
					End If
				End With
	
				rsAllBatchJobs = Nothing
	
				Response.Write("<input type='hidden' Name='txtOriginalEventID' id='txtOriginalEventID' value='" & Request("txtEventID") & "'>" & vbCrLf)
			%>

			<div id="findGridRow" style="height: 70%; margin-right: 20px; margin-left: 20px;">

				<table align="center" cellpadding="5" cellspacing="0" width="100%" height="100%">
					<tr>
						<td>
							<table width="100%" height="100%" cellspacing="0" cellpadding="0">
								<tr height="5">
									<td colspan="3"></td>
								</tr>

								<tr>
									<td width="5"></td>
									<td>
										<table width="100%" height="100%" cellspacing="0" cellpadding="5">
											<tr valign="top">
												<td>
													<table height="100%" width="100%" class="invisible" cellspacing="0" cellpadding="0">
														<tr>
															<td>
																<table height="100%" width="100%" cellspacing="0" cellpadding="4">

																	<%					

																		If Session("eventBatch") = True Then
																			Response.Write("										<tr height='20'> " & vbCrLf)
																			Response.Write("												<td>" & vbCrLf)
																			Response.Write("													<table width='100%' class='invisible' cellspacing='0' cellpadding='4'>" & vbCrLf)
																			Response.Write("														<tr> " & vbCrLf)
																			Response.Write("															<td width='120' nowrap>" & vbCrLf)
																			Response.Write("																Batch Job Name :  " & vbCrLf)
																			Response.Write("															</td> " & vbCrLf)
																			Response.Write("															<td width='200' name='tdBatchJobName' id='tdBatchJobName'> " & vbCrLf)
																			Response.Write("																" & Session("eventBatchName") & vbCrLf)
																			Response.Write("															</td>" & vbCrLf)
																			Response.Write("															<td width='120' nowrap> " & vbCrLf)
																			Response.Write("																All Jobs in Batch :  " & vbCrLf)
																			Response.Write("															</td>" & vbCrLf)
																			Response.Write("															<td> " & vbCrLf)
		
																			Response.Write(Session("cboString"))
		
																			Response.Write("															</td> " & vbCrLf)
																			Response.Write("														</tr>" & vbCrLf)
																			Response.Write("													</table>" & vbCrLf)
																			Response.Write("												</td>" & vbCrLf)
																			Response.Write("											</tr>" & vbCrLf)
																			Response.Write("											<tr height='10'> " & vbCrLf)
																			Response.Write("												<td>" & vbCrLf)
																			Response.Write("													<hr> " & vbCrLf)
																			Response.Write("												</td>" & vbCrLf)
																			Response.Write("											</tr>" & vbCrLf)
																		End If
																	%>
																	<tr height="10">
																		<td>
																			<table width="100%" class="invisible" cellspacing="0" cellpadding="4">
																				<tr>
																					<td nowrap>Name : 
																					</td>
																					<td name="tdame" id="tdName" colspan="3"></td>

																					<td width="8%" nowrap>Mode : 
																					</td>
																					<td width="35%" name="tdMode" id="tdMode"></td>
																				</tr>

																				<tr height="5">
																					<td colspan="6"></td>
																				</tr>

																				<tr height="10">
																					<td width="8%" nowrap>Start : 
																					</td>
																					<td width="25%" name="tdStartTime" id="tdStartTime"></td>
																					<td width="8%" nowrap>End : 
																					</td>
																					<td width="25%" name="tdEndTime" id="tdEndTime"></td>
																					<td width="8%" nowrap>Duration : 
																					</td>
																					<td width="25%" name="tdDuration" id="tdDuration"></td>
																				</tr>

																				<tr height="5">
																					<td colspan="6"></td>
																				</tr>

																				<tr height="10">
																					<td nowrap>Type : 
																					</td>
																					<td name="tdType" id="tdType"></td>
																					<td nowrap>Status : 
																					</td>
																					<td name="tdStatus" id="tdStatus"></td>
																					<td width="9%" nowrap>User name : 
																					</td>
																					<td width="15%" name="tdUser" id="tdUser"></td>
																				</tr>
																			</table>
																		</td>
																	</tr>
																	<tr height="10">
																		<td>
																			<hr>
																		</td>
																	</tr>
																	<tr height="10">
																		<td>
																			<table width='100%' class="invisible" cellspacing="0" cellpadding="4">
																				<tr>
																					<td width="23%" nowrap>Records Successful : 
																					</td>
																					<td width="10%" name="tdSuccessCount" id="tdSuccessCount"></td>
																					<td width="20%" nowrap>Records Failed : 
																					</td>
																					<td width="50%" name="tdFailCount" id="tdFailCount"></td>
																				</tr>
																			</table>
																		</td>
																	</tr>
																	<tr height="5">
																		<td colspan="6"></td>
																	</tr>
																	<tr>
																		<td colspan="6" id="gridCell" name="gridCell">

																			<table id="ssOleDBGridEventLogDetails" class='outline' style="width: 100%">
																				<tr class='header'>
																					<th style='text-align: left;'>ID</th>
																					<th style='text-align: left;'>Details</th>
																				</tr>

																				<%
																					Dim iDetailCount As Integer = 0

																					Dim prmEventExists = New SqlParameter("piExists", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
																					Dim rsEventDetails = objDataAccess.GetDataTable("spASRIntGetEventLogDetails", CommandType.StoredProcedure _
																						, New SqlParameter("piBatchRunID", SqlDbType.Int) With {.Value = CleanNumeric(Request("txtEventBatchRunID"))} _
																						, New SqlParameter("piEventID", SqlDbType.Int) With {.Value = CleanNumeric(Request("txtEventID"))} _
																						, prmEventExists)
		
																					For Each objRow As DataRow In rsEventDetails.Rows
																						iDetailCount = iDetailCount + 1
				
																						sValue = objRow("Notes").ToString()
																						sValue = Replace(sValue, """", "&quot;")	'escape quotes

																						Response.Write("<tr disabled='disabled'>")
																						Response.Write("<td><input type='radio' value='row_" & objRow("EventLogID") & "'></td>")
																						Response.Write("<td class='findGridCell' id='col_" & iDetailCount.ToString() & "'>" & sValue & "<input id='detail_" & objRow("EventLogID") & "' type='hidden' value='" & sValue & "'></td>")
																						Response.Write("</tr>")
																			
																					Next
																				%>
																			</table>
																				<% 
																					If prmEventExists.Value > 0 Then
																						Response.Write("<input type='hidden' Name='txtEventExists' id='txtEventExists' value='1'>" & vbCrLf)
																					Else
																						Response.Write("<input type='hidden' Name='txtEventExists' id='txtEventExists' value='0'>" & vbCrLf)
																					End If
																				prmEventExists = Nothing
																				%>
																		</td>
																	</tr>
																</table>
															</td>
														</tr>
														<tr height="5">
															<td align="right" valign="bottom">
																<table class="invisible" cellspacing="0" cellpadding="4">
																	<tr>
																		<td width="10">
																			<input id="cmdEmail" type="button" class="button ui-button ui-widget ui-state-default ui-widget-content ui-corner-tl ui-corner-br" value="Email..." name="cmdEmail" style="width: 80px" onclick="emailEvent();" />
																			
																		</td>
																		<td width="5">
																			<input id="cmdPrint"  type="button" class="button ui-button ui-widget ui-state-default ui-widget-content ui-corner-tl ui-corner-br" value="Print..." name="cmdPrint" style="width: 80px" onclick="printEvent();" />
																		</td>
																		<td width="5">
																			<input id="cmdOK" type="button" class="button ui-button ui-widget ui-state-default ui-widget-content ui-corner-tl ui-corner-br" value="Close" name="cmdOK" style="width: 80px" onclick="okClick();" />
																		</td>
																	</tr>
																</table>
															</td>
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

			</div>

		</form>

		</div>
		<form id="frmUseful" name="frmUseful" style="visibility: hidden; display: none">
			<input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
			<%
					
				Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)

				Dim sParameterValue As String = objDatabase.GetModuleParameter("MODULE_PERSONNEL", "Param_TablePersonnel")
				Response.Write("<input type='hidden' id='txtPersonnelTableID' name='txtPersonnelTableID' value=" & sParameterValue & ">" & vbCrLf)
		
				Response.Write("<input type='hidden' id='txtErrorDescription' name='txtErrorDescription' value="""">" & vbCrLf)
				Response.Write("<input type='hidden' id='txtAction' name='txtAction' value=" & Session("action") & ">" & vbCrLf)
			
				
			%>
		</form>

		<form id="frmEmail" name="frmEmail" method="post" style="visibility: hidden; display: none" action="emailSelection">
			<input type="hidden" id="txtSelectedEventIDs" name="txtSelectedEventIDs">
			<input type="hidden" id="txtBatchInfo" name="txtBatchInfo">
			<input type="hidden" id="txtBatchy" name="txtBatchy" value="0">
			<input type="hidden" id="txtFromMain" name="txtFromMain" value="0">
		</form>


		<script type="text/javascript">

			function eventlogdetails_window_onload() {

				// Convert table to jQuery grid
				tableToGrid("#ssOleDBGridEventLogDetails", {
					cmTemplate: { sortable: false, editable: false },
					beforeSelectRow: function (rowid, e) {
						return false;
					},
					rowNum: 1000,
					height: 200
				});

				$('#ssOleDBGridEventLogDetails').hideCol("ID");

				$("#ssOleDBGridEventLogDetails").jqGrid('setGridHeight', $("#findGridRow").height());
				var y = $("#gbox_findGridTable").height();
				var z = $('#gbox_findGridTable .ui-jqgrid-bdiv').height();

				$("#ssOleDBGridEventLogDetails").setGridWidth($("#findGridRow").width() - 50);

				if ($("#txtEventExists").val() == 0) {

					var frmOpenerRefresh = OpenHR.getForm("workframe", "frmRefresh");
					var frmMainLog = OpenHR.getForm("workframe", "frmLog");



					OpenHR.messageBox("This record no longer exists in the event log.", 48, "Event Log");

					frmOpenerRefresh.txtCurrentUsername.value = frmMainLog.cboUsername.options[frmMainLog.cboUsername.selectedIndex].value;
					frmOpenerRefresh.txtCurrentType.value = frmMainLog.cboType.options[frmMainLog.cboType.selectedIndex].value;
					frmOpenerRefresh.txtCurrentMode.value = frmMainLog.cboMode.options[frmMainLog.cboMode.selectedIndex].value;
					frmOpenerRefresh.txtCurrentStatus.value = frmMainLog.cboStatus.options[frmMainLog.cboStatus.selectedIndex].value;

					frmOpenerRefresh.submit();

					self.close();
				} else {
					if (<%: Request("txtEmailPermission") %> == 1) {
						button_disable(frmEventDetails.cmdEmail, false);
					} else {
						button_disable(frmEventDetails.cmdEmail, true);
					}

					populateEventInfo();
					populateEventDetails();
					setGridCaption();
				}
			}

			function okClick() {
				self.close();
			}

			function emailEvent() {
				var sBatchInfo = "";
				var sURL;

				if (frmEventDetails.txtEventBatch.value == 1) {
					frmEmail.txtBatchy.value = 1;
					frmEmail.txtSelectedEventIDs.value = frmEventDetails.cboOtherJobs.options[frmEventDetails.cboOtherJobs.selectedIndex].value;

					sBatchInfo = sBatchInfo + "Batch Job Name :	" + document.getElementById('tdBatchJobName').innerText + String.fromCharCode(13) + String.fromCharCode(13);

					sBatchInfo = sBatchInfo + "All Jobs in Batch :	" + String.fromCharCode(13) + String.fromCharCode(13);

					for (var iCount = 0; iCount < frmEventDetails.cboOtherJobs.options.length; iCount++) {
						sBatchInfo = sBatchInfo + String(frmEventDetails.cboOtherJobs.options[iCount].text) + String.fromCharCode(13) + String.fromCharCode(13);
					}
				}
				else {
					frmEmail.txtBatchy.value = 0;
					frmEmail.txtSelectedEventIDs.value = frmEventDetails.txtOriginalEventID.value;
				}

				frmEmail.txtBatchInfo.value = sBatchInfo;

				sURL = "emailSelection" +
					"?txtSelectedEventIDs=" + frmEmail.txtSelectedEventIDs.value +
					"&txtEmailOrderColumn=" +
					"&txtEmailOrderOrder=" +
					"&txtFromMain=" + frmEmail.txtFromMain.value +
					"&txtBatchInfo=" + escape(frmEmail.txtBatchInfo.value) +
					"&txtBatchy=" + frmEmail.txtBatchy.value;
				OpenHR.windowOpen(sURL, (screen.width) / 2 + 45, (screen.height) / 2 - 85, "no", "no");
			}

			function printEvent() {
				//Hide the buttons before printing...
				$("#cmdEmail").hide();
				$("#cmdPrint").hide();
				$("#cmdOK").hide();
				OpenHR.printDiv("popout_Wrapper");
				//... and show them again
				$("#cmdEmail").show();
				$("#cmdPrint").show();
				$("#cmdOK").show();
			}

			function populateEventInfo() {
				var sNumber;

				if (frmEventDetails.txtEventBatch.value == true) {
					sNumber = frmEventDetails.cboOtherJobs.options[frmEventDetails.cboOtherJobs.selectedIndex].value;
				}
				else {
					sNumber = frmEventDetails.txtOriginalEventID.value;
				}

				document.getElementById('tdName').innerHTML = document.getElementById('txtEventName_' + sNumber).value;
				document.getElementById('tdMode').innerHTML = document.getElementById('txtEventMode_' + sNumber).value;

				document.getElementById('tdStartTime').innerHTML = document.getElementById('txtEventStartTime_' + sNumber).value;
				document.getElementById('tdEndTime').innerHTML = document.getElementById('txtEventEndTime_' + sNumber).value;
				document.getElementById('tdDuration').innerHTML = document.getElementById('txtEventDuration_' + sNumber).value;

				//document.getElementById('tdTime').innerHTML = ASRIntranetFunctions.ConvertSQLDateToTime(document.getElementById('txtEventTime_' + sNumber).value);
				document.getElementById('tdType').innerHTML = document.getElementById('txtEventType_' + sNumber).value;
				document.getElementById('tdStatus').innerHTML = document.getElementById('txtEventStatus_' + sNumber).value;
				document.getElementById('tdUser').innerHTML = document.getElementById('txtEventUser_' + sNumber).value;

				document.getElementById('tdSuccessCount').innerHTML = document.getElementById('txtEventSuccessCount_' + sNumber).value;
				document.getElementById('tdFailCount').innerHTML = document.getElementById('txtEventFailCount_' + sNumber).value;
			}

			function populateEventDetails() {

				if ($("#cboOtherJobs").length > 0) {
					var eventLogId = "row_" + $("#cboOtherJobs")[0].value;
					$("#ssOleDBGridEventLogDetails tr").hide();
					$('[id^="' + eventLogId + '" ]').show();
				}
				setGridCaption();

			}

			function setGridCaption() {
				var iTotalRec;
				var sCaption;

				if ($("#cboOtherJobs").length > 0) {
					iTotalRec = $('[id^="' + "row_" + $("#cboOtherJobs")[0].value + '" ]').length;
				} else {
					iTotalRec = $("#ssOleDBGridEventLogDetails tr").length;
				}
				iTotalRec--;

				// Update the grid caption after the user has used keys to view the details
				if (iTotalRec < 1) {
					sCaption = "No details exist for this entry";
				}
				else if (iTotalRec == 1) {
					sCaption = "Details (1 Entry)";

				} else {
					sCaption = "Details (" + iTotalRec + " Entries)";
				}

				$("#ssOleDBGridEventLogDetails").setLabel("Details", sCaption);

			}

		</script>


		<script type="text/javascript">
			eventlogdetails_window_onload();
		</script>

	</div>

</body>
</html>
