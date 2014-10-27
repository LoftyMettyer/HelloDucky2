﻿<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="DMI.NET" %>

<!DOCTYPE html>
<html>
<head>
	<title>Event Log Selection - OpenHR</title>
	<script src="<%: Url.LatestContent("~/bundles/jQuery")%>" type="text/javascript"></script>
	<script src="<%: Url.LatestContent("~/bundles/jQueryUI7")%>" type="text/javascript"></script>
	<script src="<%: Url.LatestContent("~/bundles/OpenHR_General")%>" type="text/javascript"></script>
	<script src="<%: Url.LatestContent("~/bundles/eventlog")%>" type="text/javascript"></script>
	<script id="officebarscript" src="<%: Url.LatestContent("~/Scripts/officebar/jquery.officebar_MODIFIED.js")%>" type="text/javascript"></script>
	
	<link href="<%: Url.LatestContent("~/Content/OpenHR.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%: Url.LatestContent("~/Content/Site.css")%>" rel="stylesheet" type="text/css" />
	<link id="DMIthemeLink" href="<%: Url.LatestContent("~/Content/themes/" & Session("ui-admin-theme").ToString() & "/jquery-ui.min.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/general_enclosed_foundicons.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/font-awesome.min.css")%>" rel="stylesheet" type="text/css" />
	<link href="<%= Url.LatestContent("~/Content/fonts/SSI80v194934/style.css")%>" rel="stylesheet" />
</head>

	<script type="text/javascript">

		function eventlogpurge_window_onload() {
			//self.focus();
			$(".button").button();
			refreshControls();
		}	

		function okClick() {

			var frmOpenerPurge = window.dialogArguments.OpenHR.getForm("workframe", "frmPurge");
			var frmMainLog = window.dialogArguments.OpenHR.getForm("workframe", "frmLog");


			if ((frmEventPurge.cboPeriod.selectedIndex == 3) && (frmEventPurge.txtPeriod.value > 200)) {
				OpenHR.messageBox("You cannot select a purge period of greater than 200 years.", 48, "Event Log");
			}
			else {
				if (frmEventPurge.cboPeriod.selectedIndex == 0) {
					frmOpenerPurge.txtPurgePeriod.value = 'dd';
				}
				else if (frmEventPurge.cboPeriod.selectedIndex == 1) {
					frmOpenerPurge.txtPurgePeriod.value = 'wk';
				}
				else if (frmEventPurge.cboPeriod.selectedIndex == 2) {
					frmOpenerPurge.txtPurgePeriod.value = 'mm';
				}
				else if (frmEventPurge.cboPeriod.selectedIndex == 3) {
					frmOpenerPurge.txtPurgePeriod.value = 'yy';
				}

				frmOpenerPurge.txtPurgeFrequency.value = frmEventPurge.txtPeriod.value;
				if (frmEventPurge.optPurge.checked == true) {
					frmOpenerPurge.txtDoesPurge.value = 1;
				}
				else {
					frmOpenerPurge.txtDoesPurge.value = 0;
				}

				frmOpenerPurge.txtCurrentUsername.value = frmMainLog.cboUsername.options[frmMainLog.cboUsername.selectedIndex].value;
				frmOpenerPurge.txtCurrentType.value = frmMainLog.cboType.options[frmMainLog.cboType.selectedIndex].value;
				frmOpenerPurge.txtCurrentMode.value = frmMainLog.cboMode.options[frmMainLog.cboMode.selectedIndex].value;
				frmOpenerPurge.txtCurrentStatus.value = frmMainLog.cboStatus.options[frmMainLog.cboStatus.selectedIndex].value;

				window.dialogArguments.OpenHR.submitForm(frmOpenerPurge);
				self.close();
			}
		}

		function cancelClick() {
			//self.close();
			window.parent.closeclick();
			//$(this).dialog("close");
		}

		function spinRecords(pfUp) {
			var iRecords = frmEventPurge.txtPeriod.value;
			if (pfUp == true) {
				iRecords = ++iRecords;
			}
			else {
				if (iRecords > 0) {
					iRecords = iRecords - 1;
				}
			}
			frmEventPurge.txtPeriod.value = iRecords;
		}

		function refreshControls() {
			frmEventPurge.optNoPurge.checked = (frmEventPurge.txtPurge.value == 0);
			frmEventPurge.optPurge.checked = (frmEventPurge.txtPurge.value == 1);

			if (frmEventPurge.optNoPurge.checked == true) {
				text_disable(frmEventPurge.txtPeriod, true);
				button_disable(frmEventPurge.cmdPeriodDown, true);
				button_disable(frmEventPurge.cmdPeriodUp, true);
				combo_disable(frmEventPurge.cboPeriod, true);

				frmEventPurge.cboPeriod.value = '';
				frmEventPurge.txtPeriodIndex.value = 0;
				frmEventPurge.txtPeriod.value = 0;
				frmEventPurge.txtFrequency.value = 0;
			}
			else {
				text_disable(frmEventPurge.txtPeriod, false);
				button_disable(frmEventPurge.cmdPeriodDown, false);
				button_disable(frmEventPurge.cmdPeriodUp, false);
				combo_disable(frmEventPurge.cboPeriod, false);

				frmEventPurge.cboPeriod.selectedIndex = frmEventPurge.txtPeriodIndex.value;

				frmEventPurge.txtPeriod.value = frmEventPurge.txtFrequency.value;
			}
		}

		function setRecordsNumeric() {
			var sConvertedValue;
			var sDecimalSeparator;
			var sThousandSeparator;
			var sPoint;

			sDecimalSeparator = "\\";
			sDecimalSeparator = sDecimalSeparator.concat(OpenHR.LocaleDecimalSeparator);
			var reDecimalSeparator = new RegExp(sDecimalSeparator, "gi");

			sThousandSeparator = "\\";
			sThousandSeparator = sThousandSeparator.concat(OpenHR.LocaleThousandSeparator);
			var reThousandSeparator = new RegExp(sThousandSeparator, "gi");

			sPoint = "\\.";
			var rePoint = new RegExp(sPoint, "gi");

			if (frmEventPurge.txtPeriod.value == '') {
				frmEventPurge.txtPeriod.value = 0;
			}

			// Convert the value from locale to UK settings for use with the isNaN funtion.
			sConvertedValue = new String(frmEventPurge.txtPeriod.value);

			// Remove any thousand separators.
			sConvertedValue = sConvertedValue.replace(reThousandSeparator, "");
			frmEventPurge.txtPeriod.value = sConvertedValue;
			frmEventPurge.txtFrequency.value = sConvertedValue;

			// Convert any decimal separators to '.'.
			if (OpenHR.LocaleDecimalSeparator != ".") {
				// Remove decimal points.
				sConvertedValue = sConvertedValue.replace(rePoint, "A");
				// replace the locale decimal marker with the decimal point.
				sConvertedValue = sConvertedValue.replace(reDecimalSeparator, ".");
			}

			if (isNaN(sConvertedValue) == true) {
				OpenHR.messageBox("Invalid numeric value.", 48, "Event Log");
				frmEventPurge.txtPeriod.value = 0;
			}
			else {
				if (sConvertedValue.indexOf(".") >= 0) {
					OpenHR.messageBox("Invalid integer value.", 48, "Event Log");
					frmEventPurge.txtPeriod.value = 0;
					frmEventPurge.txtFrequency.value = 0;
				}
				else {
					if (frmEventPurge.txtPeriod.value < 0) {
						OpenHR.messageBox("The value cannot be negative.", 48, "Event Log");
						frmEventPurge.txtPeriod.value = 0;
						frmEventPurge.txtFrequency.value = 0;
					}
					else {
						if (frmEventPurge.txtPeriod.value > 999) {
							OpenHR.messageBox("The value cannot be greater than 999.", 48, "Event Log");
							frmEventPurge.txtPeriod.value = 999;
							frmEventPurge.txtFrequency.value = 999;
						}
					}
				}
			}
		}
	</script>

<body>
	<div>
		<div class="pageTitleDiv" style="margin-bottom: 15px">
			<span class="pageTitle" id="PopupReportDefinition_PageTitle">Purge Criteria</span>
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
	
				rsPurgeInfo = Nothing
			%>
			<table class="invisible" style="border-spacing: 0; border-collapse: collapse; width: 100%">
				<tr>
					<td style="width: 8px"></td>
					<td>
						<input id="optNoPurge" name="optSelection" type="radio" checked="checked"
							onclick="frmEventPurge.txtPurge.value = 0; refreshControls();" />
					</td>
					<td colspan="6" style="text-align: left">
						<label
							tabindex="-1"
							for="optNoPurge"
							class="radio">
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
							onclick="frmEventPurge.txtPurge.value = 1; refreshControls();" />
					</td>
					<td style="text-align: left; white-space: nowrap">
						<label
							tabindex="-1"
							for="optPurge"
							class="radio">
							Purge Event Log entries older than : 
						</label>
					</td>
					<td style="width: 15px">
						<input id="txtPeriod" class="text" name="txtPeriod" style="width: 40px" value="0"
							onkeyup="setRecordsNumeric();"
							onchange="setRecordsNumeric();">
					</td>
					<td style="width: 5px"></td>
					<td style="">
						<input style="" type="button" value="+" id="cmdPeriodUp" name="cmdPeriodUp" class="button"
							onclick="spinRecords(true); setRecordsNumeric();" />
					</td>
					<td style="">
						<input style="" type="button" value="-" id="cmdPeriodDown" name="cmdPeriodDown" class="button"
							onclick="spinRecords(false); setRecordsNumeric();" />
					</td>

					<td>
						<select name="cboPeriod" id="cboPeriod" class="floatright combo ui-widget-content ui-corner-tl ui-corner-bl">
							<option name="Day" value="0">
							Day(s)
												<option name="Week" value="1">
							Week(s)
												<option name="Month" value="2">
							Month(s)
												<option name="Year" value="3">
							Year(s)
						</select>
					</td>
				</tr>
			</table>
		</form>

		<div id="divEventLogPurgeButtons" class="clearboth">
			<input type="button" value="OK" id="cmdOK" onclick=" okClick() " />
			<input type="button" value="Cancel" id="okClick" onclick="cancelClick();;" />
		</div>
	</div>
</body>
</html>

	<script type="text/javascript">
		eventlogpurge_window_onload();
	</script>
