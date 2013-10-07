<%@ Page Language="VB" Inherits="System.Web.Mvc.ViewPage" %>
<%@ Import Namespace="DMI.NET" %>

<!DOCTYPE html>
<html>
<head>
	<title>Event Log Selection - OpenHR Intranet</title>
	<script src="<%: Url.Content("~/bundles/jQuery")%>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/bundles/jQueryUI7")%>" type="text/javascript"></script>
	<script src="<%: Url.Content("~/bundles/OpenHR_General")%>" type="text/javascript"></script>
	<link href="/Content/general_enclosed_foundicons.css?v=8.0.9.0" rel="stylesheet" type="text/css" />
	<link href="/Content/font-awesome.css?v=8.0.9.0" rel="stylesheet" type="text/css" />
	<link href="/Content/fonts/SSI80v194934/style.css?v=8.0.9.0" rel="stylesheet" />
	<link href="/Content/Site.css?v=8.0.9.0" rel="stylesheet" type="text/css" />
	<link href="/Content/OpenHR.css?v=8.0.9.0" rel="stylesheet" type="text/css" />
	<link href="/Content/contextmenustyle.css?v=8.0.9.0" rel="stylesheet" type="text/css" />
	<link id="Link1" href="/Content/themes/redmond/jquery-ui.min.css?v=8.0.9.0" rel="stylesheet" type="text/css" />
	<link href="/Content/ui.jqgrid.css?v=8.0.9.0" rel="stylesheet" type="text/css" />
	<style type="text/css">
		.auto-style1
		{
			width: 307px;
		}
  </style>
</head>

<body>
	<object classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
		id="Microsoft_Licensed_Class_Manager_1_0"
		viewastext>
		<param name="LPKPath" value="lpks/main.lpk">
	</object>

	<script type="text/javascript">
		function eventlogpurge_window_onload() {

			self.focus();

			// Resize the grid to show all prompted values.
			iResizeBy = frmEventPurge.offsetParent.scrollWidth - frmEventPurge.offsetParent.clientWidth;
			if (frmEventPurge.offsetParent.offsetWidth + iResizeBy > screen.width) {
				window.dialogWidth = new String(screen.width) + "px";
			} else {
				iNewWidth = new Number(window.dialogWidth.substr(0, window.dialogWidth.length - 2));
				iNewWidth = iNewWidth + iResizeBy;
				window.dialogWidth = new String(iNewWidth) + "px";
			}

			iResizeBy = frmEventPurge.offsetParent.scrollHeight - frmEventPurge.offsetParent.clientHeight;
			if (frmEventPurge.offsetParent.offsetHeight + iResizeBy > screen.height) {
				window.dialogHeight = new String(screen.height) + "px";
			} else {
				iNewHeight = new Number(window.dialogHeight.substr(0, window.dialogHeight.length - 2));
				iNewHeight = iNewHeight + iResizeBy;
				window.dialogHeight = new String(iNewHeight) + "px";
			}

			refreshControls();
		}
		
	</script>

	<script type="text/javascript" id="scptGeneralFunctions">

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
		self.close();
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

	<form id="frmEventPurge" name="frmEventPurge">
		<%
			Dim rsPurgeInfo
			Dim sSQL
			Dim iPeriod
			Dim cmdPurgeInfo
	
			cmdPurgeInfo = CreateObject("ADODB.Command")
			cmdPurgeInfo.CommandText = "spASRIntGetEventLogPurgeDetails"
			cmdPurgeInfo.CommandType = 4
			cmdPurgeInfo.ActiveConnection = Session("databaseConnection")

			Err.Clear()
			rsPurgeInfo = cmdPurgeInfo.Execute
	
			If rsPurgeInfo.BOF And rsPurgeInfo.EOF Then
				Response.Write("<INPUT type=hidden name=txtPurge id=txtPurge value=0>" & vbCrLf)
				Response.Write("<INPUT type=hidden name=txtPeriodIndex id=txtPeriodIndex>" & vbCrLf)
				Response.Write("<INPUT type=hidden name=txtFrequency id=txtFrequency>" & vbCrLf)
			Else
				Response.Write("<INPUT type=hidden name=txtPurge id=txtPurge value=1>" & vbCrLf)
				Response.Write("<INPUT type=hidden name=txtFrequency id=txtFrequency value=" & rsPurgeInfo.Fields("Frequency").Value & ">" & vbCrLf)
		
				Select Case UCase(rsPurgeInfo.Fields("Period").value)
					Case "DD" : iPeriod = 0
					Case "WK" : iPeriod = 1
					Case "MM" : iPeriod = 2
					Case "YY" : iPeriod = 3
					Case Else : iPeriod = 0
				End Select
		
				Response.Write("<INPUT type=hidden name=txtPeriodIndex id=txtPeriodIndex value=" & iPeriod & ">" & vbCrLf)
			End If
	
			rsPurgeInfo.close()
			rsPurgeInfo = Nothing
			cmdPurgeInfo = Nothing
		%>

		<%--<table align="center" cellpadding="5" cellspacing="0" width="100%" height="100%">--%>
			<table style="text-align : center; padding: 5px; border-spacing: 0;  width: 30%;  height: 100%; margin-right: 0px;">
			<tr>
				<td>
					<%--<table width="100%" height="100%" class="invisible" cellspacing="0" cellpadding="0">--%>
						<table class="invisible" style="width: 30%;  height: 100%; padding: 0; border-spacing: 0;">
						<tr style="height: 5px">
							<td colspan="2"></td>
						</tr>
						<tr>
							<td style="width:5px"></td>
							<td class="auto-style1">
								<table class="invisible" style="width: 86%;  height: 100%; padding: 0; border-spacing: 0;">
									<tr style="height: 5px">
										<td colspan="8" style="text-align: left">Purge Criteria : 
										</td>
									</tr>
									<tr style="height: 10px">
										<td colspan="8"></td>
									</tr>
									<tr>
										<td style="width: 8px"></td>
										<td>
											<input id="optNoPurge" name="optSelection" type="radio"
												onclick="frmEventPurge.txtPurge.value=0;refreshControls();"
												onmouseover="try{radio_onMouseOver(this);}catch(e){}"
												onmouseout="try{radio_onMouseOut(this);}catch(e){}"
												onfocus="try{radio_onFocus(this);}catch(e){}"
												onblur="try{radio_onBlur(this);}catch(e){}" />
										</td>
										<td colspan="6" style="text-align: left">
											<label
												tabindex="-1"
												for="optNoPurge"
												class="radio"
												onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
												onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
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
												onclick="frmEventPurge.txtPurge.value=1;refreshControls();"
												onmouseover="try{radio_onMouseOver(this);}catch(e){}"
												onmouseout="try{radio_onMouseOut(this);}catch(e){}"
												onfocus="try{radio_onFocus(this);}catch(e){}"
												onblur="try{radio_onBlur(this);}catch(e){}" />
										</td>
										<td style="text-align: left; white-space: nowrap">
											<label
												tabindex="-1"
												for="optPurge"
												class="radio"
												onmouseover="try{radioLabel_onMouseOver(this);}catch(e){}"
												onmouseout="try{radioLabel_onMouseOut(this);}catch(e){}">
											Purge Event Log entries older than : 
											</label>
										</td>
										<td style="width:15px">
											<input id="txtPeriod" class="text" name="txtPeriod" style="WIDTH: 40px" value="0" 
												onkeyup="setRecordsNumeric();"
												onchange="setRecordsNumeric();">
										</td>
										<td style="width:5px"></td>
										<td style="width: 15px">
											<input style="WIDTH: 15px" type="button" value="+" id="cmdPeriodUp" name="cmdPeriodUp" class="button ui-button ui-widget ui-state-default ui-widget-content ui-corner-tl ui-corner-br"
												onclick="spinRecords(true);setRecordsNumeric();"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
										<td style="width: 15px">
											<input style="WIDTH: 15px" type="button" value="-" id="cmdPeriodDown" name="cmdPeriodDown" class="button ui-button ui-widget ui-state-default ui-widget-content ui-corner-tl ui-corner-br"
												onclick="spinRecords(false);setRecordsNumeric();"
												onmouseover="try{button_onMouseOver(this);}catch(e){}"
												onmouseout="try{button_onMouseOut(this);}catch(e){}"
												onfocus="try{button_onFocus(this);}catch(e){}"
												onblur="try{button_onBlur(this);}catch(e){}" />
										</td>
										<td style="width: 10px">&nbsp;</td>
										<td>
											<select name="cboPeriod" id="cboPeriod" class="combo ui-widget ui-widget-content ui-corner-all">
												<option name="Day" value="0">Day(s)
												<option name="Week" value="1">Week(s)
												<option name="Month" value="2">Month(s)
												<option name="Year" value="3">Year(s)
											</select>
										</td>
									</tr>
									<tr style="height: 5px">
										<td colspan="8"></td>
									</tr>
									<tr style="height: 5px; text-align: right">
										<td style="text-align: right;  vertical-align: bottom" colspan="8">
											<table class="invisible" style="border-spacing: 0; padding: 4px">
												<tr>
													<td style="width: 10px">
														<input id="cmdOK" type="button" value="OK" name="cmdOk" style="WIDTH: 80px" class="button ui-button ui-widget ui-state-default ui-widget-content ui-corner-tl ui-corner-br"
															onclick="okClick();"
															onmouseover="try{button_onMouseOver(this);}catch(e){}"
															onmouseout="try{button_onMouseOut(this);}catch(e){}"
															onfocus="try{button_onFocus(this);}catch(e){}"
															onblur="try{button_onBlur(this);}catch(e){}" />
													</td>
													<td style="width: 5px">
														<input id="cmdCancel" type="button" value="Cancel" name="cmdCancel"  style="WIDTH: 80px" class="button ui-button ui-widget ui-state-default ui-widget-content ui-corner-tl ui-corner-br"
															onclick="cancelClick();"
															onmouseover="try{button_onMouseOver(this);}catch(e){}"
															onmouseout="try{button_onMouseOut(this);}catch(e){}"
															onfocus="try{button_onFocus(this);}catch(e){}"
															onblur="try{button_onBlur(this);}catch(e){}" />
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
			</tr>
		</table>
	</form>


	<script type="text/javascript">
		eventlogpurge_window_onload();
	</script>

	<p>
		&nbsp;</p>

</body>
</html>
