<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>

<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="ADODB" %>

<!DOCTYPE html>

<html>
<head runat="server">

	<title>OpenHR Intranet</title>
	<script src="<%: Url.Content("~/bundles/jQuery")%>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/bundles/jQueryUI7")%>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/bundles/OpenHR_General")%>" type="text/javascript"></script>


	<%--Here's the stylesheets for the font-icons displayed on the dashboard for wireframe and tile layouts--%>
	<link href="<%= Url.LatestContent("~/Content/general_enclosed_foundicons.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/font-awesome.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/fonts/SSI80v194934/style.css")%>" rel="stylesheet" />

	<%--Base stylesheets--%>
	<link href="<%: Url.LatestContent("~/Content/Site.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />

	<%--stylesheet for slide-out dmi menu--%>
	<link href="<%: Url.LatestContent("~/Content/contextmenustyle.css")%>" rel="stylesheet" type="text/css" />

	<%--ThemeRoller stylesheet--%>
	<link id="DMIthemeLink" href="<%: Url.LatestContent("~/Content/themes/" & Session("ui-theme").ToString() & "/jquery-ui.min.css")%>" rel="stylesheet" type="text/css" />

	<%--jQuery Grid Stylesheet--%>
	<link href="<%: Url.LatestContent("~/Content/ui.jqgrid.css")%>" rel="stylesheet" type="text/css" />

</head>
<body>

	<div>

		<%Html.RenderPartial("~/Views/Shared/ctl_ASRIntranetPrintFunctions.ascx")%>

		<form id="frmEventDetails" name="frmEventDetails">

			<%
				Dim rsAllBatchJobs As Recordset
				Dim i As Integer
				Dim sValue As String

				Dim objUtilities As HR.Intranet.Server.Utilities
		
				Dim cmdEventBatchJobs
				Dim prmBatchRunID
				Dim prmEventID
		
				objUtilities = Session("UtilitiesObject")
		
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

				cmdEventBatchJobs = CreateObject("ADODB.Command")
				cmdEventBatchJobs.CommandText = "spASRIntGetEventLogBatchDetails"
				cmdEventBatchJobs.CommandType = 4	' Stored procedure
				cmdEventBatchJobs.ActiveConnection = Session("databaseConnection")
								
				prmBatchRunID = cmdEventBatchJobs.CreateParameter("BatchRunID", 3, 1)	' 3=integer, 1=input
				cmdEventBatchJobs.Parameters.Append(prmBatchRunID)
				prmBatchRunID.value = CleanNumeric(Request("txtEventBatchRunID"))

				prmEventID = cmdEventBatchJobs.CreateParameter("EventID", 3, 1)	' 3=integer, 1=input
				cmdEventBatchJobs.Parameters.Append(prmEventID)
				prmEventID.value = CleanNumeric(Request("txtEventID"))

				Err.Clear()
				rsAllBatchJobs = cmdEventBatchJobs.Execute
	
				With rsAllBatchJobs
					If Not (.EOF And .BOF) Then
						i = 0
						Do Until .EOF
							i = i + 1

							Response.Write("<input type='hidden' Name='txtEventID_" & .Fields("ID").Value & "' id='txtEventID_" & .Fields("ID").Value & "' value='" & Replace(.Fields("ID").Value, """", "&quot;") & "'>" & vbCrLf)
				
							sValue = .Fields("Name").Value											'original value
							sValue = Replace(sValue, """", "&quot;")		'escape quotes
							sValue = Replace(sValue, "<", "&lt;")						'escape left angle bracket
							sValue = Replace(sValue, ">", "&gt;")						'escape right angle bracket
				
							Response.Write("<input type='hidden' Name='txtEventName_" & .Fields("ID").Value & "' id='txtEventName_" & .Fields("ID").Value & "' value='" & sValue & "'>" & vbCrLf)
							Response.Write("<input type='hidden' Name='txtEventMode_" & .Fields("ID").Value & "' id='txtEventMode_" & .Fields("ID").Value & "' value='" & Replace(.Fields("Mode").Value, """", "&quot;") & "'>" & vbCrLf)
				
							Response.Write("<input type='hidden' Name='txtEventStartTime_" & .Fields("ID").Value & "' id='txtEventStartTime_" & .Fields("ID").Value & "' value='" & ConvertSQLDateToLocale(.Fields("DateTime").Value) & " " & ConvertSqlDateToTime(.Fields("DateTime").Value) & "'>" & vbCrLf)
				
							If IsDBNull(.Fields("EndTime").Value) Then
								Response.Write("<input type='hidden' Name='txtEventEndTime_" & .Fields("ID").Value & "' id='txtEventEndTime_" & .Fields("ID").Value & "' value='" & vbNullString & "'>" & vbCrLf)
							Else
								Response.Write("<input type='hidden' Name='txtEventEndTime_" & .Fields("ID").Value & "' id='txtEventEndTime_" & .Fields("ID").Value & "' value='" & ConvertSQLDateToLocale(.Fields("EndTime").Value) & " " & ConvertSqlDateToTime(.Fields("EndTime").Value) & "'>" & vbCrLf)
							End If
				
							Response.Write("<input type='hidden' Name='txtEventDuration_" & .Fields("ID").Value & "' id='txtEventDuration_" & .Fields("ID").Value & "' value='" & objUtilities.FormatEventDuration(CLng(.Fields("Duration").Value)) & "'>" & vbCrLf)

							Response.Write("<input type='hidden' Name='txtEventType_" & .Fields("ID").Value & "' id='txtEventType_" & .Fields("ID").Value & "' value='" & Replace(.Fields("Type").Value, """", "&quot;") & "'>" & vbCrLf)
							Response.Write("<input type='hidden' Name='txtEventStatus_" & .Fields("ID").Value & "' id='txtEventStatus_" & .Fields("ID").Value & "' value='" & Replace(.Fields("Status").Value, """", "&quot;") & "'>" & vbCrLf)
							Response.Write("<input type='hidden' Name='txtEventUser_" & .Fields("ID").Value & "' id='txtEventUser_" & .Fields("ID").Value & "' value='" & Replace(.Fields("Username").Value, """", "&quot;") & "'>" & vbCrLf)
				
							Response.Write("<input type='hidden' Name='txtEventSuccessCount_" & .Fields("ID").Value & "' id='txtEventSuccessCount_" & .Fields("ID").Value & "' value='" & Replace(.Fields("SuccessCount").Value, """", "&quot;") & "'>" & vbCrLf)
							Response.Write("<input type='hidden' Name='txtEventFailCount_" & .Fields("ID").Value & "' id='txtEventFailCount_" & .Fields("ID").Value & "' value='" & Replace(.Fields("FailCount").Value, """", "&quot;") & "'>" & vbCrLf)
				
							Response.Write("<input type='hidden' Name='txtEventBatchRunID_" & .Fields("ID").Value & "' id='txtEventBatchRunID_" & .Fields("ID").Value & "' value='" & Replace(.Fields("BatchRunID").Value, """", "&quot;") & "'>" & vbCrLf)
							Response.Write("<input type='hidden' Name='txtEventBatchName_" & .Fields("ID").Value & "' id='txtEventBatchName_" & .Fields("ID").Value & "' value='" & Replace(.Fields("BatchName").Value, """", "&quot;") & "'>" & vbCrLf)
							Response.Write("<input type='hidden' Name='txtEventBatchJobID_" & .Fields("ID").Value & "' id='txtEventBatchJobID_" & .Fields("ID").Value & "' value='" & Replace(.Fields("BatchJobID").Value, """", "&quot;") & "'>" & vbCrLf)
				
							If Session("eventBatch") = True Then
								If Session("eventID") = .Fields("ID").Value Then
									Session("cboString") = Session("cboString") & "<option selected='selected' name='" & .Fields("Name").Value & "' value='" & .Fields("ID").Value & "'>" & .Fields("Type").Value & " - " & .Fields("Name").Value & "</option>" & vbCrLf
								Else
									Session("cboString") = Session("cboString") & "<option name='" & .Fields("Name").Value & "' value='" & .Fields("ID").Value & "'>" & .Fields("Type").Value & " - " & .Fields("Name").Value & "</option>" & vbCrLf
								End If
							End If
				
							.MoveNext()
						Loop
			
						Session("eventBatchName") = Request("txtEventBatchName")
			
						Session("cboString") = Session("cboString") & "</SELECT>" & vbCrLf
						If i <= 1 Then
							Session("cboString") = "<select disabled id=cboOtherJobs name=cboOtherJobs class=""combodisabled"" style=""WIDTH: 100%"" onchange='populateEventInfo();populateEventDetails();'>" & vbCrLf & Session("cboString")
						Else
							Session("cboString") = "<select id=cboOtherJobs name=cboOtherJobs class=""combo"" style=""WIDTH: 100%"" onchange='populateEventInfo();populateEventDetails();'>" & vbCrLf & Session("cboString")
						End If
					End If
				End With
	
				rsAllBatchJobs = Nothing
				cmdEventBatchJobs = Nothing
				prmEventID = Nothing
				prmBatchRunID = Nothing
				objUtilities = Nothing
	
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
																					Dim iDetailCount
																					Dim rsEventDetails
																					Dim cmdEventDetails
																					Dim prmEventExists

																					iDetailCount = 0
	
																					cmdEventDetails = CreateObject("ADODB.Command")
																					cmdEventDetails.CommandText = "spASRIntGetEventLogDetails"
																					cmdEventDetails.CommandType = 4	' Stored procedure
																					cmdEventDetails.ActiveConnection = Session("databaseConnection")
								
																					prmBatchRunID = cmdEventDetails.CreateParameter("BatchRunID", 3, 1)	' 3=integer, 1=input
																					cmdEventDetails.Parameters.Append(prmBatchRunID)
																					prmBatchRunID.value = CleanNumeric(Request("txtEventBatchRunID"))

																					prmEventID = cmdEventDetails.CreateParameter("EventID", 3, 1)	' 3=integer, 1=input
																					cmdEventDetails.Parameters.Append(prmEventID)
																					prmEventID.value = CleanNumeric(Request("txtEventID"))

																					prmEventExists = cmdEventDetails.CreateParameter("EventExists", 3, 2)	' 3=integer, 2=output
																					cmdEventDetails.Parameters.Append(prmEventExists)

																					Err.Clear()
																					rsEventDetails = cmdEventDetails.Execute

																					If Not (rsEventDetails.BOF And rsEventDetails.EOF) Then
																						Do While Not rsEventDetails.EOF
																							iDetailCount = iDetailCount + 1
				
																							sValue = rsEventDetails.Fields("Notes").value				'original value
																							sValue = Replace(sValue, """", "&quot;")	'escape quotes

																							Response.Write("<tr disabled='disabled'>")
																							Response.Write("<td><input type='radio' value='row_" & rsEventDetails.Fields("EventLogID").value & "'></td>")
																							Response.Write("<td class='findGridCell' id='col_" & iDetailCount.ToString() & "'>" & sValue & "<input id='detail_" & rsEventDetails.Fields("EventLogID").value & "' type='hidden' value='" & sValue & "'></td>")
																							Response.Write("</tr>")
																			
																							rsEventDetails.MoveNext()
																						Loop
																					End If
																					rsEventDetails.close()
																					rsEventDetails = Nothing
	
																					If cmdEventDetails.Parameters("EventExists").Value > 0 Then
																						Response.Write("<input type='hidden' Name='txtEventExists' id='txtEventExists' value='1'>" & vbCrLf)
																					Else
																						Response.Write("<input type='hidden' Name='txtEventExists' id='txtEventExists' value='0'>" & vbCrLf)
																					End If
	
																					cmdEventDetails = Nothing
																					prmEventID = Nothing
																					prmBatchRunID = Nothing
																					prmEventExists = Nothing
																				%>
																			</table>

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
																			<input id="cmdPrint"  type="button" class="button ui-button ui-widget ui-state-default ui-widget-content ui-corner-tl ui-corner-br" value="Print..." name="cmdPrint" style="width: 80px" onclick="printEvent(true);" />
																		</td>
																		<td width="5">
																			<input id="cmdOK" type="button" class="button ui-button ui-widget ui-state-default ui-widget-content ui-corner-tl ui-corner-br" value="OK" name="cmdOK" style="width: 80px" onclick="okClick();" />
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

		<form id="frmUseful" name="frmUseful" style="visibility: hidden; display: none">
			<input type="hidden" id="txtUserName" name="txtUserName" value="<%=session("username")%>">
			<%
				Dim cmdDefinition
				Dim prmModuleKey
				Dim prmParameterKey
				Dim prmParameterValue
				Dim sErrorDescription As String
		
				cmdDefinition = CreateObject("ADODB.Command")
				cmdDefinition.CommandText = "sp_ASRIntGetModuleParameter"
				cmdDefinition.CommandType = 4	' Stored procedure.
				cmdDefinition.ActiveConnection = Session("databaseConnection")

				prmModuleKey = cmdDefinition.CreateParameter("moduleKey", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
				cmdDefinition.Parameters.Append(prmModuleKey)
				prmModuleKey.value = "MODULE_PERSONNEL"

				prmParameterKey = cmdDefinition.CreateParameter("paramKey", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
				cmdDefinition.Parameters.Append(prmParameterKey)
				prmParameterKey.value = "Param_TablePersonnel"

				prmParameterValue = cmdDefinition.CreateParameter("paramValue", 200, 2, 8000)	'200=varchar, 2=output, 8000=size
				cmdDefinition.Parameters.Append(prmParameterValue)

				Err.Clear()
				cmdDefinition.Execute()

				Response.Write("<input type='hidden' id='txtPersonnelTableID' name='txtPersonnelTableID' value='" & cmdDefinition.Parameters("paramValue").Value & "'>" & vbCrLf)
	
				cmdDefinition = Nothing

				Response.Write("<input type='hidden' id='txtErrorDescription' name='txtErrorDescription' value='" & sErrorDescription & "'>" & vbCrLf)
				Response.Write("<input type='hidden' id='txtAction' name='txtAction' value='" & Session("action") & "'>" & vbCrLf)
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
					rowNum: 1000
				});

				$("#ssOleDBGridEventLogDetails").jqGrid('setGridHeight', $("#findGridRow").height());
				var y = $("#gbox_findGridTable").height();
				var z = $('#gbox_findGridTable .ui-jqgrid-bdiv').height();

				$("#DefSelRecords").setGridHeight($("#findGridRow").height());
				$("#DefSelRecords").setGridWidth($("#findGridRow").width());

				if ($("#txtEventExists").value == 0) {

					var frmOpenerRefresh = window.dialogArguments.OpenHR.getForm("workframe", "frmRefresh");
					var frmMainLog = window.dialogArguments.OpenHR.getForm("workframe", "frmLog");

					OpenHR.messageBox("This record no longer exists in the event log.", 48, "Event Log");

					frmOpenerRefresh.txtCurrentUsername.value = frmMainLog.cboUsername.options[frmMainLog.cboUsername.selectedIndex].value;
					frmOpenerRefresh.txtCurrentType.value = frmMainLog.cboType.options[frmMainLog.cboType.selectedIndex].value;
					frmOpenerRefresh.txtCurrentMode.value = frmMainLog.cboMode.options[frmMainLog.cboMode.selectedIndex].value;
					frmOpenerRefresh.txtCurrentStatus.value = frmMainLog.cboStatus.options[frmMainLog.cboStatus.selectedIndex].value;

					frmOpenerRefresh.submit();

					self.close();
				} else {
					var frmOpenerDetails = window.dialogArguments.OpenHR.getForm("workframe", "frmDetails");

					if (frmOpenerDetails.txtEmailPermission.value == 1) {
						button_disable(frmEventDetails.cmdEmail, false);
					} else {
						button_disable(frmEventDetails.cmdEmail, true);
					}

					populateEventInfo();
					populateEventDetails();
					setGridCaption();
				}
			}
		</script>

		<script type="text/javascript" id="scptGeneralFunctions">

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
				openDialog(sURL, 500, 400);
			}

			function printEvent(pfToPrinter) {
				var fOK = true;
				var sErrorString = new String("");
				var objPrinter = ASRIntranetPrintFunctions;
				var iDetailCount;

				if (pfToPrinter == true) {
					if (objPrinter.IsOK == false) {
						return;
					}
				}

				// OK so far.
				if (pfToPrinter == true) {
					fOK = objPrinter.PrintStart(false, frmUseful.txtUserName.value);
				}

				if (fOK == true) {

					if (pfToPrinter == true) {
						//print selected event information
						objPrinter.PrintHeader("Event Log : " + document.getElementById('tdName').innerText);
						objPrinter.PrintNonBold("Mode :	" + document.getElementById('tdMode').innerText);
						objPrinter.PrintNormal("");
						objPrinter.PrintNonBold("Start Time :	" + document.getElementById('tdStartTime').innerText);
						objPrinter.PrintNonBold("End Time :	" + document.getElementById('tdEndTime').innerText);
						objPrinter.PrintNonBold("End Time :	" + document.getElementById('tdDuration').innerText);
						objPrinter.PrintNormal("");
						objPrinter.PrintNonBold("Type :	" + document.getElementById('tdType').innerText);
						objPrinter.PrintNonBold("Status :	" + document.getElementById('tdStatus').innerText);
						objPrinter.PrintNonBold("User name :	" + document.getElementById('tdUser').innerText);
						objPrinter.PrintNormal("");
					}

					if (pfToPrinter == true && (frmEventDetails.txtEventBatch.value == 1)) {
						//print batch job information
						objPrinter.PrintNonBold("Batch Job Name :	" + document.getElementById('tdBatchJobName').innerText);
						objPrinter.PrintNormal("");
						objPrinter.PrintNormal("All Jobs in Batch :");
						objPrinter.PrintNormal("");

						for (var iCount = 0; iCount < frmEventDetails.cboOtherJobs.options.length; iCount++) {
							objPrinter.PrintNonBold(frmEventDetails.cboOtherJobs.options[iCount].text);
						}
					}

					if (pfToPrinter == true) {
						//print records summary information			
						objPrinter.PrintNormal("");
						objPrinter.PrintNonBold("Records Successful :	" + document.getElementById('tdSuccessCount').innerText);
						objPrinter.PrintNonBold("Records Failed :	" + document.getElementById('tdFailCount').innerText);
					}

					if (pfToPrinter == true) {
						//print selected event details
						objPrinter.PrintNormal("");
						objPrinter.PrintBold("Details : ");
						objPrinter.PrintNormal("");

						if ($("#cboOtherJobs").length > 0) {
							iDetailCount = $('[id^="' + "row_" + $("#cboOtherJobs")[0].value + '" ]').length;
						} else {
							iDetailCount = $("#ssOleDBGridEventLogDetails tr").length;
						}

						if (iDetailCount < 1) {
							objPrinter.PrintNonBold("There are no details for this event log entry");
						}
						else {

							var a;
							var rows;

							if ($("#cboOtherJobs").length > 0) {
								rows = $('[id^="' + "row_" + $("#cboOtherJobs")[0].value + '" ]');
							} else {
								rows = $("#ssOleDBGridEventLogDetails tr");
							}

							for (a = 1; a < rows.length; a++) {
								objPrinter.PrintBold("*** Log entry " + a + " of " + iDetailCount + " ***");

								sErrorString = rows[a].cells[1].innerText;

								objPrinter.PrintNonBold(sErrorString);
								objPrinter.PrintNormal("");

							}
						}
					}

					if (pfToPrinter == true) {
						objPrinter.PrintEnd();
						objPrinter.PrintConfirm("Event Log Details", "Event Log Details");
					}
				}
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

				//Update the grid caption after the user has used keys to view the details
				if (iTotalRec == 0) {
					sCaption = "No details exist for this entry";
				}
				else {
					sCaption = "Details (" + iTotalRec + " Entries)";
				}

				$("#ssOleDBGridEventLogDetails").setLabel("Details", sCaption);

			}

			function openDialog(pDestination, pWidth, pHeight) {
				dlgwinprops = "center:yes;" +
					"dialogHeight:" + pHeight + "px;" +
					"dialogWidth:" + pWidth + "px;" +
					"help:no;" +
					"resizable:yes;" +
					"scroll:yes;" +
					"status:no;";
				window.showModalDialog(pDestination, self, dlgwinprops);
			}

		</script>


		<script type="text/javascript">
			eventlogdetails_window_onload();
		</script>

	</div>

</body>
</html>
