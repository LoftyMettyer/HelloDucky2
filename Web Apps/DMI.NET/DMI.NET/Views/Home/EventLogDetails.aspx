<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage(of DMI.NET.Models.EventDetailModel)" %>

<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="DMI.NET.Helpers" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="DMI.NET.Models" %>

<div id="popout_Wrapper">
		<div class="pageTitleDiv margebot10">
			<span class="pageTitle" id="PopupEventDetail">Event Detail</span>
		</div>
		<form id="frmEventDetails" name="frmEventDetails">

			<%
				Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
				Dim i As Integer
				Dim detailsLabel1 As String = ""
				Dim detailsLabel2 As String = ""
				
				Session("cboString") = vbNullString

				If Model.IsBatch Then
					Response.Write("<input type='hidden' Name='txtEventBatch' ID='txtEventBatch' value='1'>" & vbCrLf)
					
					If Model.Mode = "Batch" Then
						detailsLabel1 = "Batch Job Name"
						detailsLabel2 = "All Jobs in Batch"
						Session("txtEventMode") = "Batch"
					Else
						detailsLabel1 = "Report Pack Name"
						detailsLabel2 = "All Reports in Pack"
						Session("txtEventMode") = "Pack"
					End If
				Else
					Session("txtEventMode") = "Manual"
					Response.Write("<input type='hidden' Name='txtEventBatch' ID='txtEventBatch' value='0'>" & vbCrLf)
				End If
				
				Dim rsAllBatchJobs = objDataAccess.GetDataTable("spASRIntGetEventLogBatchDetails", CommandType.StoredProcedure _
					, New SqlParameter("piBatchRunID", SqlDbType.Int) With {.Value = Model.BatchRunID} _
					, New SqlParameter("piEventID", SqlDbType.Int) With {.Value = Model.ID})
				
				With rsAllBatchJobs
					i = 0
					For Each objRow As DataRow In .Rows

						i += 1

						Response.Write("<input type='hidden' Name='txtEventID_" & objRow("ID") & "' id='txtEventID_" & objRow("ID") & "' value='" & Replace(CType(objRow("ID"), String), """", "&quot;") & "'>" & vbCrLf)
						Response.Write("<input type='hidden' Name='txtEventName_" & objRow("ID") & "' id='txtEventName_" & objRow("ID") & "' value='" & Html.Encode(objRow("Name")) & "'>" & vbCrLf)
						Response.Write("<input type='hidden' Name='txtEventMode_" & objRow("ID") & "' id='txtEventMode_" & objRow("ID") & "' value='" & Replace(CType(objRow("Mode"), String), """", "&quot;") & "'>" & vbCrLf)
				
						Response.Write("<input type='hidden' Name='txtEventStartTime_" & objRow("ID") & "' id='txtEventStartTime_" & objRow("ID") & "' value='" & ConvertSQLDateToLocale(objRow("DateTime")) & " " & ConvertSqlDateToTime(objRow("DateTime")) & "'>" & vbCrLf)
				
						If IsDBNull(objRow("EndTime")) Then
							Response.Write("<input type='hidden' Name='txtEventEndTime_" & objRow("ID") & "' id='txtEventEndTime_" & objRow("ID") & "' value='" & vbNullString & "'>" & vbCrLf)
						Else
							Response.Write("<input type='hidden' Name='txtEventEndTime_" & objRow("ID") & "' id='txtEventEndTime_" & objRow("ID") & "' value='" & ConvertSQLDateToLocale(objRow("EndTime")) & " " & ConvertSqlDateToTime(objRow("EndTime")) & "'>" & vbCrLf)
						End If
				
						Response.Write("<input type='hidden' Name='txtEventDuration_" & objRow("ID") & "' id='txtEventDuration_" & objRow("ID") & "' value='" & FormatEventDuration(CType(objRow("Duration"), Integer)) & "'>" & vbCrLf)

						Response.Write("<input type='hidden' Name='txtEventType_" & objRow("ID") & "' id='txtEventType_" & objRow("ID") & "' value='" & Replace(CType(objRow("Type"), String), """", "&quot;") & "'>" & vbCrLf)
						Response.Write("<input type='hidden' Name='txtEventStatus_" & objRow("ID") & "' id='txtEventStatus_" & objRow("ID") & "' value='" & Replace(CType(objRow("Status"), String), """", "&quot;") & "'>" & vbCrLf)
						Response.Write("<input type='hidden' Name='txtEventUser_" & objRow("ID") & "' id='txtEventUser_" & objRow("ID") & "' value='" & Replace(CType(objRow("Username"), String), """", "&quot;") & "'>" & vbCrLf)
				
						Response.Write("<input type='hidden' Name='txtEventSuccessCount_" & objRow("ID") & "' id='txtEventSuccessCount_" & objRow("ID") & "' value='" & Replace(CType(objRow("SuccessCount"), String), """", "&quot;") & "'>" & vbCrLf)
						Response.Write("<input type='hidden' Name='txtEventFailCount_" & objRow("ID") & "' id='txtEventFailCount_" & objRow("ID") & "' value='" & Replace(CType(objRow("FailCount"), String), """", "&quot;") & "'>" & vbCrLf)
				
						Response.Write("<input type='hidden' Name='txtEventBatchRunID_" & objRow("ID") & "' id='txtEventBatchRunID_" & objRow("ID") & "' value='" & Replace(CType(objRow("BatchRunID"), String), """", "&quot;") & "'>" & vbCrLf)
						Response.Write("<input type='hidden' Name='txtEventBatchName_" & objRow("ID") & "' id='txtEventBatchName_" & objRow("ID") & "' value='" & Replace(CType(objRow("BatchName"), String), """", "&quot;") & "'>" & vbCrLf)
						Response.Write("<input type='hidden' Name='txtEventBatchJobID_" & objRow("ID") & "' id='txtEventBatchJobID_" & objRow("ID") & "' value='" & Replace(CType(objRow("BatchJobID"), String), """", "&quot;") & "'>" & vbCrLf)
				
						If Model.IsBatch Then
							If Model.ID = objRow("ID") Then
								Session("cboString") = Session("cboString") & "<option selected='selected' name='" & objRow("Name") & "' value='" & objRow("ID") & "'>" & objRow("Type") & " - " & objRow("Name") & "</option>" & vbCrLf
							Else
								Session("cboString") = Session("cboString") & "<option name='" & objRow("Name") & "' value='" & objRow("ID") & "'>" & objRow("Type") & " - " & objRow("Name") & "</option>" & vbCrLf
							End If
						End If
					Next
					
					Try
						Session("eventBatchName") = .Rows(0)("BatchName")
					Catch ex As Exception
						Session("eventBatchName") = ""
					End Try
			
					Session("cboString") = Session("cboString") & "</SELECT>" & vbCrLf
					If i <= 1 Then
						Session("cboString") = "<select disabled id=cboOtherJobs name=cboOtherJobs class=""combodisabled"" style=""WIDTH: 100%"" onchange='populateEventInfo();populateEventDetails();'>" & vbCrLf & Session("cboString")
					Else
						Session("cboString") = "<select id=cboOtherJobs name=cboOtherJobs class=""combo"" style=""WIDTH: 100%"" onchange='populateEventInfo();populateEventDetails();'>" & vbCrLf & Session("cboString")
					End If
				End With
	
				Response.Write("<input type='hidden' Name='txtOriginalEventID' id='txtOriginalEventID' value='" & Model.ID & "'>" & vbCrLf)
			%>

			<div id="findGridRow" style="height: 30%; margin-right: 20px; margin-left: 20px;">
				<div>
					<table class="width100" style="padding: 4px; border-collapse: collapse">
						<%					

							If Model.IsBatch Then
								Response.Write("										<tr height='20'> " & vbCrLf)
								Response.Write("												<td>" & vbCrLf)
								Response.Write("													<table width='100%' class='invisible' cellspacing='0' cellpadding='4'>" & vbCrLf)
								Response.Write("														<tr> " & vbCrLf)
								Response.Write("															<td width='130px' class='fontsmalltitle' nowrap>" & vbCrLf)
								Response.Write("																" & detailsLabel1 & " :  " & vbCrLf)
								Response.Write("															</td> " & vbCrLf)
								Response.Write("															<td width='200' name='tdBatchJobName' id='tdBatchJobName'> " & vbCrLf)
								Response.Write("																" & Session("eventBatchName") & vbCrLf)
								Response.Write("															</td>" & vbCrLf)
								Response.Write("															<td width='130px' class='fontsmalltitle' nowrap> " & vbCrLf)
								Response.Write("																" & detailsLabel2 & " :  " & vbCrLf)
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
						<tr style="height:10px">
							<td>
								<table class="invisible"  style="width: 100%; padding: 4px; border-collapse: collapse">
									<tr>
										<td class="fontsmalltitle">Name : </td>
										<td id="tdName" colspan="3"></td>
										<td class="fontsmalltitle" style="width: 11%">Mode : </td>
										<td style="width: 35%"  id="tdMode"></td>
									</tr>

									<tr style="height:10px">
										<td class="fontsmalltitle nowrap" style="width: 8%">Start : </td>
										<td style="width: 25%" id="tdStartTime"></td>
										<td class="fontsmalltitle" style="width: 8%">End : </td>
										<td style="width: 25%" id="tdEndTime"></td>
										<td class="fontsmalltitle" style="width: 8%">Duration : </td>
										<td style="width: 25%" id="tdDuration"></td>
									</tr>

									<tr style="height:10px">
										<td class="fontsmalltitle">Type : </td>
										<td id="tdType"></td>
										<td class="fontsmalltitle">Status : </td>
										<td id="tdStatus"></td>
										<td class="fontsmalltitle" style="white-space: nowrap">User name : </td>
										<td style="width: 15%" id="tdUser"></td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td>
								<hr>
							</td>
						</tr>
						<tr style="height: 20px">
							<td class="padbot10">
								<table class="invisible"  style="padding: 4px; border-collapse: collapse">
									<tr class="padbot10">
										<td class="fontsmalltitle">Records Successful :</td>
										<td id="tdSuccessCount"></td>
										<td class="fontsmalltitle padleft45">Records Failed : </td>
										<td id="tdFailCount"></td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td colspan="6" id="gridCell">
								<table id="ssOleDBGridEventLogDetails" class='outline' style="width: 100%">
									<tr class='header'>
										<th style='text-align: left;'>ID</th>
										<th style='text-align: left;'>Details</th>
									</tr>

									<%
										Dim iDetailCount As Integer = 0

										Dim prmEventExists = New SqlParameter("piExists", SqlDbType.Int) With {.Direction = ParameterDirection.Output}
										Dim rsEventDetails = objDataAccess.GetDataTable("spASRIntGetEventLogDetails", CommandType.StoredProcedure _
											, New SqlParameter("piBatchRunID", SqlDbType.Int) With {.Value = Model.BatchRunID} _
											, New SqlParameter("piEventID", SqlDbType.Int) With {.Value = Model.ID} _
											, prmEventExists)
		
										For Each objRow As DataRow In rsEventDetails.Rows
											iDetailCount = iDetailCount + 1

											Response.Write("<tr disabled='disabled'>")
											Response.Write("<td><input type='radio' value='row_" & objRow("EventLogID") & "'></td>")
											Response.Write("<td class='findGridCell' id='col_" & iDetailCount.ToString() & "'>" & Html.Encode(objRow("Notes")) & "<input id='detail_" & objRow("EventLogID") & "' type='hidden' value='" & Html.Encode(objRow("Notes")) & "'></td>")
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
								%>
							</td>
						</tr>

					</table>
					<div id="divEventLogViewDetailsButtons" class="clearboth">
						<input id="cmdEmail" type="button"  value="Email..." name="cmdEmail"  onclick="emailDetailEvent();" />
						<input id="cmdPrint" type="button"  value="Print..." name="cmdPrint" onclick="printEvent();" />
						<input id="cmdOK" type="button"  value="Close" name="cmdOK" onclick="okClick();" />
					</div>
				</div>
			</div>			
		</form>
			<%--This is the popout for email selection--%>
			<div id="EventDetailEmailSelect"></div>
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

<div>
	<input type="hidden" id="txtSelectedEventIDs" name="txtSelectedEventIDs">
	<input type="hidden" id="txtBatchInfo" name="txtBatchInfo">
	<input type="hidden" id="txtBatchy" name="txtBatchy" value="0">
	<input type="hidden" id="txtFromMain" name="txtFromMain" value="0">
</div>


<script type="text/javascript">
			var frmDetails = OpenHR.getForm("workframe", "frmDetails");

			function eventlogdetails_window_onload() {
				// Convert table to jQuery grid				
				tableToGrid("#ssOleDBGridEventLogDetails", {
					height: '300',
					width: 'auto',
					cmTemplate: { sortable: false, editable: false },
					beforeSelectRow: function () {
						return false;
					},
					rowNum: 1000
				});
				
				$('#ssOleDBGridEventLogDetails').hideCol("ID");
				$("#ssOleDBGridEventLogDetails").setGridWidth($("#findGridRow").width());

				if ($('#txtEventExists').val() == 0) {
					var frmOpenerRefresh = OpenHR.getForm("workframe", "frmRefresh");
					var frmMainLog = OpenHR.getForm("workframe", "frmLog");

					OpenHR.messageBox("This record no longer exists in the event log.", 48, "Event Log");
					frmOpenerRefresh.txtCurrentUsername.value = frmMainLog.cboUsername.options[frmMainLog.cboUsername.selectedIndex].value;
					frmOpenerRefresh.txtCurrentType.value = frmMainLog.cboType.options[frmMainLog.cboType.selectedIndex].value;
					frmOpenerRefresh.txtCurrentMode.value = frmMainLog.cboMode.options[frmMainLog.cboMode.selectedIndex].value;
					frmOpenerRefresh.txtCurrentStatus.value = frmMainLog.cboStatus.options[frmMainLog.cboStatus.selectedIndex].value;
					frmOpenerRefresh.submit();
					okClick();

				} else {
					if ($("#txtEmailPermission").val() == 1) {
						button_disable($('#cmdEmail'), false);
					} else {
						button_disable($('#cmdEmail'), true);
					}
					
					populateEventInfo();
					populateEventDetails();
					setGridCaption();
				}

				$('#EventLogViewDetails').dialog(
				{
					position: ['center']
				});
			}

			function okClick() {
				$("#EventLogViewDetails").dialog("close");
			}

			function emailDetailEvent() {
				var sBatchInfo = "";
				var sURL;
				
				if ($('#txtEventBatch').val() == 1) {
					$("#txtBatchy").val(1);
					$('#txtSelectedEventIDs').val(frmEventDetails.cboOtherJobs.options[frmEventDetails.cboOtherJobs.selectedIndex].value);

					sBatchInfo = sBatchInfo + "<%:detailsLabel1%> :	" + $('#tdBatchJobName').html() + String.fromCharCode(13) + String.fromCharCode(13);
					sBatchInfo = sBatchInfo + "<%:detailsLabel2%> :	" + String.fromCharCode(13) + String.fromCharCode(13);

					for (var iCount = 0; iCount < $("#cboOtherJobs").length; iCount++) {
						sBatchInfo = sBatchInfo + String(frmEventDetails.cboOtherJobs.options[iCount].text) + String.fromCharCode(13) + String.fromCharCode(13);
					}
				}
				else {
					$("#txtBatchy").val(0); 
					$('#txtSelectedEventIDs').val($('#txtOriginalEventID').val());
				}
				
				$("#txtBatchInfo").val(sBatchInfo);

				var postData = {
					SelectedEventIDs: $('#txtSelectedEventIDs').val(),
					EmailOrderColumn: "",
					EmailOrderOrder: "",
					IsFromMain: $("#txtFromMain").val(),
					BatchInfo: escape($("#txtBatchInfo").val()),
					IsBatchy: $("#txtBatchy").val(),
					<%:Html.AntiForgeryTokenForAjaxPost() %> 
				};
			
				$('#EventLogEmailSelect').dialog("open");
				OpenHR.submitForm(null, "EventLogEmailSelect", null, postData, "EventLogEmail");

			}			

			function printEvent() {
				//Hide the buttons before printing...
				$("#divEventLogViewDetailsButtons").hide();

				OpenHR.printDiv("popout_Wrapper");
				
				//... and show them again
				$("#divEventLogViewDetailsButtons").show();
			}

			function populateEventInfo() {
				var sNumber;
				
				if ($('#txtEventBatch').val() == true) {
					sNumber = frmEventDetails.cboOtherJobs.options[frmEventDetails.cboOtherJobs.selectedIndex].value;
				}
				else {
					sNumber = $('#txtOriginalEventID').val();
				}

				$('#tdName').text($('#txtEventName_' + sNumber).val());
				$('#tdMode').html('<%:Session("txtEventMode")%>');
				$('#tdStartTime').html($('#txtEventStartTime_' + sNumber).val());
				$('#tdEndTime').html($('#txtEventEndTime_' + sNumber).val());
				$('#tdDuration').html($('#txtEventDuration_' + sNumber).val());
				$('#tdType').html($('#txtEventType_' + sNumber).val());
				$('#tdStatus').html($('#txtEventStatus_' + sNumber).val());
				$('#tdUser').html($('#txtEventUser_' + sNumber).val());
				$('#tdSuccessCount').html($('#txtEventSuccessCount_' + sNumber).val());
				$('#tdFailCount').html($('#txtEventFailCount_' + sNumber).val());
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
				var sCaption = "Details"; 
				//If one of the entries in the "All Reports in Batch" or "All Reports in Pack" contains additional details, then the grid will contain those details; however, 
				//if one of the entries does not contain any details, the details are only hidden (i.e. the table is not cleared), so the number of items is not reported correctly
				//So I've set the caption to "Details" only
				$("#ssOleDBGridEventLogDetails").setLabel("Details", sCaption);
			}
		</script>
	
		<script type="text/javascript">

			eventlogdetails_window_onload();
		</script>
