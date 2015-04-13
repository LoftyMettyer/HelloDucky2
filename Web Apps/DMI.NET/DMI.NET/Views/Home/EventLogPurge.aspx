<%@ OutputCache Duration="1" varyByParam="none"%>
<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HR.Intranet.Server" %>

<script type="text/javascript">
		function eventlogpurge_window_onload() {
			$('#optNoPurge').prop('checked', ($('#txtPurge').val() == 0));
			$('#optPurge').prop('checked', ($('#txtPurge').val() == 1));			
	
			refreshControls();
		}	

		function okClick() {
			var frmMainLog = OpenHR.getForm("workframe", "frmLog");
			var frmOpenerPurge = OpenHR.getForm("workframe", "frmPurge");
			
			if (($('#cboPeriod').val() == 3) && ($('#txtPeriod').val() > 200)) {
				OpenHR.modalPrompt("You cannot select a purge period of greater than 200 years.", 0, "Event Log");
			}
			else {
				if ($('#cboPeriod').val() == 0) {
					frmOpenerPurge.txtPurgePeriod.value = 'dd';
				}
				else if ($('#cboPeriod').val() == 1) {
					frmOpenerPurge.txtPurgePeriod.value = 'wk';
				}
				else if ($('#cboPeriod').val() == 2) {
					frmOpenerPurge.txtPurgePeriod.value = 'mm';
				}
				else if ($('#cboPeriod').val() == 3) {
					frmOpenerPurge.txtPurgePeriod.value = 'yy';
				}

				frmOpenerPurge.txtPurgeFrequency.value = $('#txtPeriod').val();
				if ($('#optPurge').prop('checked') == true) {
					frmOpenerPurge.txtDoesPurge.value = 1;
				}
				else {
					frmOpenerPurge.txtDoesPurge.value = 0;
				}

				frmOpenerPurge.txtCurrentUsername.value = frmMainLog.cboUsername.options[frmMainLog.cboUsername.selectedIndex].value;
				frmOpenerPurge.txtCurrentType.value = frmMainLog.cboType.options[frmMainLog.cboType.selectedIndex].value;
				frmOpenerPurge.txtCurrentMode.value = frmMainLog.cboMode.options[frmMainLog.cboMode.selectedIndex].value;
				frmOpenerPurge.txtCurrentStatus.value = frmMainLog.cboStatus.options[frmMainLog.cboStatus.selectedIndex].value;

				OpenHR.submitForm(frmOpenerPurge);
				$('#EventLogPurge').dialog("close");
			}
		}

		function cancelClick() {
			$('#EventLogPurge').dialog("close");
		}

		function spinRecords(pfUp) {
			var iRecords = $('#txtPeriod').val();
			if (pfUp == true) {
				iRecords = ++iRecords;
			}
			else {
				if (iRecords > 0) {
					iRecords = iRecords - 1;
				}
			}
			 $('#txtPeriod').val(iRecords);
		}

		function refreshControls() {
			
			if ($('#optNoPurge').prop('checked') == true) {
				text_disable($('#txtPeriod'), true);
				button_disable($('#cmdPeriodDown'), true);
				button_disable($('#cmdPeriodUp'), true);
				combo_disable($('#cboPeriod'), true);
				$('#cboPeriod').val('');
				$('#txtPeriodIndex').val(0);
				$('#txtPeriod').val(0);
				$('#txtFrequency').val(0);
			}
			if ($('#optPurge').prop('checked') == true) {
				text_disable($('#txtPeriod'), false);
				button_disable($('#cmdPeriodDown'), false);
				button_disable($('#cmdPeriodUp'), false);
				combo_disable($('#cboPeriod'), false);
				$('#cboPeriod').val($('#txtPeriodIndex').val());
				$('#txtPeriod').val($('#txtFrequency').val());
			}
		}

		function setRecordsNumeric() {
			var sConvertedValue;
			var sDecimalSeparator;
			var sThousandSeparator;
			var sPoint;

			sDecimalSeparator = "\\";
			sDecimalSeparator = sDecimalSeparator.concat(OpenHR.LocaleDecimalSeparator());
			var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");

			sThousandSeparator = "\\";
			sThousandSeparator = sThousandSeparator.concat(OpenHR.LocaleThousandSeparator());
			var reThousandSeparator = new RegExp(sThousandSeparator, "gi");

			sPoint = "\\.";
			var rePoint = new RegExp(sPoint, "gi");

			if ( $('#txtPeriod').val() == '') {
				 $('#txtPeriod').val(0);
			}

			// Convert the value from locale to UK settings for use with the isNaN funtion.
			sConvertedValue = new String($('#txtPeriod').val());

			// Remove any thousand separators.
			sConvertedValue = sConvertedValue.replace(reThousandSeparator, "");
			$('#txtPeriod').val(sConvertedValue);
			$('#txtFrequency').val(sConvertedValue);

			// Convert any decimal separators to '.'.
			if (OpenHR.LocaleDecimalSeparator() != ".") {
				// Remove decimal points.
				sConvertedValue = sConvertedValue.replace(rePoint, "A");
				// replace the locale decimal marker with the decimal point.
				sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
			}
			
			if (isNaN(sConvertedValue) == true) {
				OpenHR.modalPrompt("Invalid numeric value.", 0, "Event Log");
				 $('#txtPeriod').val(0);
			}
			else {
				if (sConvertedValue.indexOf(".") >= 0) {
					OpenHR.modalPrompt("Invalid integer value.", 0, "Event Log");
					$('#txtPeriod').val(0);
					$('#txtFrequency').val(0);
				}
				else {
					if ($('#txtPeriod').val() < 0) {
						OpenHR.modalPrompt("The value cannot be negative.", 0, "Event Log");
						$('#txtPeriod').val(0);
						$('#txtFrequency').val(0);
					}
					else {
						if ($('#txtPeriod').val() > 999) {
							OpenHR.modalPrompt("The value cannot be greater than 999.", 0, "Event Log");
							$('#txtPeriod').val(999);
							$('#txtFrequency').val(999);
						}
					}
				}
			}
		}
</script>

	<div>
		<div class="pageTitleDiv padbot15">
			<span class="pageTitle" id="PopupEventPurgeTitle">Purge Criteria</span>
		</div>

		<form id="frmEventPurge" name="frmEventPurge">
			<%
			
				Dim objDataAccess As clsDataAccess = CType(Session("DatabaseAccess"), clsDataAccess)
				Dim iPeriod As Integer
			
				Dim rsPurgeInfo = objDataAccess.GetDataTable("spASRIntGetEventLogPurgeDetails", CommandType.StoredProcedure)
		
				If rsPurgeInfo.Rows.Count = 0 Then

					Response.Write("<input type=hidden name=txtPurge id=txtPurge value=0>" & vbCrLf)
					Response.Write("<input type=hidden name=txtPeriodIndex id=txtPeriodIndex>" & vbCrLf)
					Response.Write("<input type=hidden name=txtFrequency id=txtFrequency>" & vbCrLf)
				Else
					Dim objRow = rsPurgeInfo.Rows(0)
					Response.Write("<input type=hidden name=txtPurge id=txtPurge value=1>" & vbCrLf)
					Response.Write("<input type=hidden name=txtFrequency id=txtFrequency value=" & objRow("Frequency") & ">" & vbCrLf)
		
					Select Case UCase(objRow("Period").ToString)
						Case "DD" : iPeriod = 0
						Case "WK" : iPeriod = 1
						Case "MM" : iPeriod = 2
						Case "YY" : iPeriod = 3
						Case Else : iPeriod = 0
					End Select
		
					Response.Write("<input type=hidden name=txtPeriodIndex id=txtPeriodIndex value=" & iPeriod & ">" & vbCrLf)
				End If
			%>
			<table class="invisible width100" style="border-spacing: 0; border-collapse: collapse">
				<tr>
					<td style="width: 8px"></td>
					<td>
						<input id="optNoPurge" name="optSelection" type="radio" checked="checked"
							onclick="$('#txtPurge').val(0); refreshControls();" />
					</td>
					<td class="alignleft" colspan="6">
						<label for="optNoPurge" tabindex="-1">
							Do not automatically purge the Event Log
						</label>
					</td>
				</tr>
				<tr style="height: 5px">
					<td colspan="8"></td>
				</tr>
				<tr>
					<td style="width: 8px"></td>
					<td>
						<input id="optPurge" name="optSelection" type="radio"
							onclick="$('#txtPurge').val(1); refreshControls();" />
					</td>
					<td class="alignleft">
						<label for="optPurge" tabindex="-1">
							Purge Event Log entries older than : 
						</label>
					</td>
					<td style="width: 15px">
						<input id="txtPeriod" name="txtPeriod" onchange="setRecordsNumeric();" onkeyup="setRecordsNumeric();" style="width: 40px" value="0">
					</td>
					<td style="width: 5px"></td>
					<td>
						<input class="button" id="cmdPeriodUp" name="cmdPeriodUp" onclick="spinRecords(true); setRecordsNumeric();" type="button" value="+" />
					</td>
					<td>
						<input class="button" id="cmdPeriodDown" name="cmdPeriodDown" onclick="spinRecords(false); setRecordsNumeric();" type="button" value="-" />
					</td>

					<td>
						<select class="floatright ui-widget-content ui-corner-tl ui-corner-bl" id="cboPeriod" name="cboPeriod">
							<option name="Day" value="0">Day(s)
							<option name="Week" value="1">Week(s)
							<option name="Month" value="2">Month(s)
							<option name="Year" value="3">Year(s)
						</select>
					</td>
				</tr>
			</table>
		</form>

		<div id="divEventLogPurgeButtons" class="clearboth">
			<input id="cmdOK" onclick="okClick();" type="button" value="OK" />
			<input id="okClick" onclick="cancelClick();" tabindex="1" type="button" value="Cancel" />
		</div>
	</div>

	<script type="text/javascript">
		eventlogpurge_window_onload();
	</script>
