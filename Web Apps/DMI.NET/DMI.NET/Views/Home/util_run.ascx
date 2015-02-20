<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>

<%
	Session("CALREP_Year") = Nothing
	Session("CALREP_Month") = Nothing
	Session("CALREP_firstLoad") = 1
	Session("CALREP_IncludeBankHolidays") = 0
	Session("CALREP_IncludeWorkingDaysOnly") = 0
	Session("CALREP_ShowBankHolidays") = 0
	Session("CALREP_ShowCaptions") = 0
	Session("CALREP_ShowWeekends") = 0
	Session("CALREP_ChangeOptions") = 0
	
	
	Session("EmailGroupID") = 0
	Session("OutputOptions_Format") = 0
	Session("OutputOptions_Screen") = "true"
	Session("OutputOptions_Save") = "false"
	Session("OutputOptions_SaveExisting") = 0
		
	' following sessions vars:
	'
	' UtilType    - 0-13 (see UtilityType code in DATMGR .exe
	' UtilName    - <the name of the utility>
	' UtilID      - <the id of the utility>
	' Action      - run/delete

	' Write the prompted values from the calling form into a session variable.
	' NB. The prompts are written into an array and this array is written to a 
	' session variables with the name 'Prompts_<util type>_<util id>.
	Dim sKey As String

	Dim aPrompts(1, 0) As String
	Dim j = 0
	ReDim Preserve aPrompts(1, 0)
	For i = 0 To (Request.Form.Count) - 1
		sKey = Request.Form.Keys(i)
		If ((UCase(Left(sKey, 7)) = "PROMPT_") And (Mid(sKey, 8, 1) <> "3")) Or _
				(UCase(Left(sKey, 10)) = "PROMPTCHK_") Then
			ReDim Preserve aPrompts(1, j)
		
			If (UCase(Left(sKey, 10)) = "PROMPTCHK_") Then
				aPrompts(0, j) = "prompt_3_" & Mid(sKey, 11)
				aPrompts(1, j) = UCase(Request.Form.Item(i))
			Else
				aPrompts(0, j) = sKey
				Select Case Mid(sKey, 8, 1)
					Case "2"
						' Numeric. Replace locale decimal point with '.'
						aPrompts(1, j) = Replace(Request.Form.Item(i), CType(Session("LocaleDecimalSeparator"), String), ".")
					Case "4"
						' Date. Reformat to match SQL's mm/dd/yyyy format.
						aPrompts(1, j) = ConvertLocaleDateToSQL(Request.Form.Item(i))
					Case Else
						aPrompts(1, j) = Request.Form.Item(i)
				End Select
			End If
			j = j + 1
		End If
	Next
	sKey = "Prompts_" & Request.Form("utiltype") & "_" & Request.Form("utilid")
	Session(sKey) = aPrompts
%>

<script type="text/javascript">
	function raiseError(sErrorDesc, fok, fcancelled) {
		var frmError = document.getElementById("frmError");
		frmError.txtUtilTypeDesc.value = window.frames("top").frmPopup.txtUtilTypeDesc.value;
		frmError.txtErrorDesc.value = sErrorDesc;
		frmError.txtOK.value = fok;
		frmError.txtUserCancelled.value = fcancelled;
		var sTarget = new String("errorMessage");
		frmError.target = sTarget;
		NewWindow('', sTarget, '500', '200', 'no');
		frmError.submit();
		self.close();
		return;
	}

	function pausecomp(millis) {
		var date = new Date();
		var curDate;

		do {
			curDate = new Date();
		} while (curDate - date < millis);
	}

	function NewWindow(mypage, myname, w, h, scroll) {
		var winl = (screen.width - w) / 2;
		var wint = (screen.height - h) / 2;
		var winprops = 'height=' + h + ',width=' + w + ',top=' + wint + ',left=' + winl + ',scrollbars=' + scroll + ',resizable';
		var win = window.open(mypage, myname, winprops);

		if (parseInt(navigator.appVersion) >= 4) {
			pausecomp(300);
			win.window.focus();
		}
	}

	function ShowWaitFrame() {

		var fs = window.parent.document.all.item("reportframe");

		if (fs) {
			fs.rows = "*,0,0";
		}

	}

	function ShowDataFrame() {
		
		$("#cmdOK").hide();
		//$("#cmdCancel").hide();
		$("#cmdCancel").removeClass('ui-state-focus');
		$("#cmdCancel").button({ disabled: true });
		$("#cmdOutput").show();
		var isMobileDevice = ('<%=Session("isMobileDevice")%>' == 'True');
		$("#cmdOutput").prop('disabled', isMobileDevice);

		if (menu_isSSIMode() == true) {
			$("#divReportButtons #cmdClose").hide();  // Don't show the Close button in SSI

		} else {
			$(".ui-dialog-buttonpane #cmdClose").show();
		}

		$("#reportbreakdownframe").hide();
		$("#top").hide();
		$("#outputoptions").hide();
		$("#optionsframeset").show();
		$("#reportframe").show();
		$("#reportworkframe").show();
	}

	function outputOptionsPrintClick() {
		//Creates a new window, copies the report grid to it, formats the grid and sends to print.

		// Only custom reports have print functionality so I have selected in a bit closer here to make the report tidier re centering etc
		// If this gets reinstated applpication wide you may need to isolate this call for 
		// Custom Reports and reinstate the line commmented below //var divToPrint = document.getElementById('reportworkframe');
		//var divToPrint = document.getElementById('reportworkframe');
		var divToPrint = document.getElementById('gview_gridReportData');
		
		var ReportTitleFromTitleBar = $(".popup").dialog('option', 'title');
		var newWin = window.open("", "_blank", 'toolbar=no,status=no,menubar=no,scrollbars=yes,resizable=yes, width=1, height=1, visible=none', "");
		
		newWin.document.write('<sty');
		newWin.document.write('le>.ui-jqgrid-bdiv {height:auto!important;}');
		newWin.document.write('body {font-family:verdana;}');
		newWin.document.write('tr.ui-jqgrid-labels {display:none;}');
		newWin.document.write('tr.jqgfirstrow {background-color:lightgray;}');
		newWin.document.write('tr.jqgrow>td {border-bottom: 1px solid lightgray; padding-right: 5px;}');
		newWin.document.write('</sty');
		newWin.document.write('le>');
		newWin.document.write("<h3>" + ReportTitleFromTitleBar + "</h3><br/>");
		newWin.document.write(divToPrint.innerHTML);
		newWin.document.write('<scri');
		newWin.document.write('pt type="text/javascript">');
		newWin.document.write("var headerCells = document.querySelectorAll('.ui-th-column>div');");
		newWin.document.write("for (var i = 0, len = headerCells.length; i < len; i++) {");
		newWin.document.write("	document.querySelector('tr.jqgfirstrow>td:nth-child(' + (i + 1) + ')').innerText = headerCells[i].innerText.replace('_', ' ');");
		newWin.document.write("	}");
		newWin.document.write('</scri');
		newWin.document.write('pt>');
		
		newWin.document.close();
		newWin.focus();
		newWin.print();
		newWin.close();
	}
	
	function ExportDataPrompt() {
		//var frmExportData = OpenHR.getForm("reportworkframe", "frmExportData");
		//OpenHR.submitForm(frmExportData, "outputoptions");
		$("#reportworkframe").hide();
		$("#reportbreakdownframe").hide();
		$("#outputoptions").show();
		$("#cmdOK").show();
		$("#cmdCancel").show();
		$("#cmdCancel").button({ disabled: false });
		$("#cmdOutput").hide();
	}
</script>

<form id="frmError" name="frmError" action="util_run_error" method="post">
	<input type="hidden" id="txtUtilTypeDesc" name="txtUtilTypeDesc">
	<input type="hidden" id="txtEventLogID" name="txtEventLogID">
	<input type="hidden" id="txtOK" name="txtOK">
	<input type="hidden" id="txtUserCancelled" name="txtUserCancelled">
	<input type="hidden" id="txtErrorDesc" name="txtErrorDesc">
</form>

<div id="divUtilRunForm">
	<div class="absolutefull">
		<div class="pageTitleDiv">
			<a href='javascript:loadPartialView("linksMain", "Home", "workframe", null);' title='Back'>
				<i class='pageTitleIcon icon-circle-arrow-left'></i>
			</a>
			<span class="pageTitleSmaller" id="PageDivTitle">
				<% 
					If Session("StandardReport_Type") <> "" Then
						Response.Write(GetReportNameByReportType(Session("StandardReport_Type")))
						If Not Session("stdReport_StartDate") Is Nothing And Not Session("stdReport_EndDate") Is Nothing Then
							Response.Write(" (" & Session("stdReport_StartDate").ToString.Replace(" ", "") & " -> " & Session("stdReport_EndDate").ToString.Replace(" ", "") & ")")
						End If
					End If
					If CBool(Session("stdReport_PrintFilterPicklistHeader")) = True Then
						If Not Session("stdReport_PicklistName") Is Nothing Then
							If Session("stdReport_PicklistName").ToString <> "" Then
								Response.Write(" (Base Table Picklist: " & Session("stdReport_PicklistName") & ")")
							End If
						End If
						If Not Session("stdReport_FilterName") Is Nothing Then
							If Session("stdReport_FilterName").ToString <> "" Then
								Response.Write(" (Base Table Filter: " & Session("stdReport_FilterName") & ")")
							End If
						End If
					End If
				%>
			</span>
		</div>

		<div id="main" data-framesource="util_run" style="height: 80%; margin: 0 0 0 0; visibility: hidden">
			<%   
				Dim sPrintButtonLabel As String = "Print"
				If Session("utiltype") = "1" Then
					Html.RenderPartial("~/Views/Home/util_run_crosstabsMain.ascx")
				ElseIf Session("utiltype") = "2" Then
					Html.RenderPartial("~/Views/Home/util_run_customreportsMain.ascx")
				ElseIf Session("utiltype") = "3" Then
					'Html.RenderPartial("~/Views/Home/util_run_datatransfer.ascx")
				ElseIf Session("utiltype") = "4" Then
					'Html.RenderPartial("~/Views/Home/util_run_export.ascx")
				ElseIf Session("utiltype") = "5" Then
					'Html.RenderPartial("~/Views/Home/util_run_globaladd.ascx")
				ElseIf Session("utiltype") = "6" Then
					'Html.RenderPartial("~/Views/Home/util_run_globalupdate.ascx")
				ElseIf Session("utiltype") = "7" Then
					'Html.RenderPartial("~/Views/Home/util_run_globaldelete.ascx")
				ElseIf Session("utiltype") = "8" Then
					'Html.RenderPartial("~/Views/Home/util_run_import.ascx")
				ElseIf Session("utiltype") = "9" Then
					Html.RenderPartial("~/Views/Home/util_run_mailmerge.ascx")
				ElseIf Session("utiltype") = "15" Then
					Html.RenderPartial("~/Views/Home/stdrpt_run_AbsenceBreakdown.ascx")
				ElseIf Session("utiltype") = "16" Then
					Html.RenderPartial("~/Views/Home/util_run_customreportsMain.ascx")
				ElseIf Session("utiltype") = "17" Then
					Html.RenderPartial("~/Views/Home/util_run_calendarreport_main.ascx")
				ElseIf Session("utiltype") = "35" Then
					Html.RenderPartial("~/Views/Home/util_run_crosstabsMain.ascx")
				Else
					' blah.
				End If
			%>
		</div>
		<br/>
		<div id="divReportButtons" style="margin: 0; visibility: hidden; padding-top: 0; float: right">
			<%If Session("SSIMode") = True Then%>
				<%If (Session("utiltype") = "2") Then%> 
					<input class="btn minwidth100" type="button" id="cmdPrint" name="cmdPrint" value="<%=sPrintButtonLabel%>" onclick="outputOptionsPrintClick()" />
				<%End If%>
				<input class="btn minwidth100" type="button" id="cmdOK" name="cmdOK" value="Export" onclick="outputOptionsOKClick()" />
				<input class="btn minwidth100" type="button" id="cmdOutput" name="cmdOutput" value="Output" onclick="ExportDataPrompt();" />
				<input class="btn minwidth100" type="button" id="cmdCancel" name="cmdCancel" value="Preview" onclick="ShowDataFrame();" />
				<input class="btn minwidth100" type="button" id="cmdClose" name="cmdClose" value="Close" onclick="closeclick();" />
			<%End If%>
			</div>
		</div>
	</div>

<script type="text/javascript">
	
	<%If Session("mailmergefail") Then%>
		closeclick();
	<%Else%>

	var isMobileDevice = ('<%=Session("isMobileDevice")%>' == 'True');
	// first get the size from the window
	// if that didn't work, get it from the body
	var size = {};
	
	<%If Session("utiltype") = UtilityType.utlNineBoxGrid Then%>
		size.width = (screen.width) / 2;
		size.height = (window.innerHeight || document.body.clientHeight) - 100;
	<%Else%>
		size.width = (window.innerWidth || document.body.clientWidth) - 200;
		size.height = (window.innerHeight || document.body.clientHeight) - 200;
	<%End If%>

	if ($('#txtNoRecs').val() == "True") {
		OpenHR.modalPrompt($("#txtDefn_ErrMsg").val(), 2, $("#txtDefn_Name").val(), "");
		closeclick();
		
		if (menu_isSSIMode()) {
			loadPartialView("linksMain", "Home", "workframe", null);
		}

	} else {
		
		if ($("#txtPreview").val() == "True") {

			<%If Session("utiltype") = UtilityType.utlCalendarReport Then%>
			$(".popup").dialog({
				width: 1100,
				height: 720,
				resizable: true
			});
			$('#main').css('overflow', 'auto');
			
			<%Else%>
			$(".popup").dialog({
				title: "",
				width: size.width,
				height: size.height,
				resizable: true,
				resize: function() {
					var doit = 0;
					clearTimeout(doit);
					doit = setTimeout(resizeGrid, 100);
				}
			});

			<%End If%>


			var newButtons = [
				<%If (Session("utiltype") = "2") Then%> 
				{
					text: "<%=sPrintButtonLabel%>",
					click: function() { outputOptionsPrintClick(); },
					"class": "minwidth100",
					"id": "cmdPrint"
				},
				<%End If%>
				{
					text: "Export",
					click: function() { outputOptionsOKClick(); },
					"class": "minwidth100",
					"id": "cmdOK"
				},
				{
					text: "Output",
					click: function() { ExportDataPrompt(); },
					"class": "minwidth100",
					"id": "cmdOutput"
				}
				,
				{
					text: "Preview",
					click: function() { ShowDataFrame(); },
					"class": "minwidth100",
					"id": "cmdCancel"
				},				
				{
					text: "Close",
					click: function() { closeclick(); },
					"class": "minwidth100",
					"id": "cmdClose"
				}
			];

			$('.popup').dialog('option', 'buttons', newButtons);

			function resizeGrid() {
				var newHeight = $('#reportworkframe').height();
				var newWidth = window.innerWidth || document.body.clientWidth;
				$('#gridReportData').setGridHeight(newHeight);
				$('#gridReportData').setGridWidth($('#reportframeset').width() * 0.95);
			}
			

			$('#cmdOutput').prop('disabled', isMobileDevice);
			$(".popup").dialog("option", "position", ['center', 'center']); //Center popup in screen
			$('.popup').bind('dialogclose', function() {
				closeclick();
			});

			if (menu_isSSIMode() == false) {
				$('#main').css('marginTop', '30px'); //.css('borderTop', '1px solid rgb(206, 206, 206)');
			}

			$("#PageDivTitle").html($("#txtDefn_Name").val());
			$(".popup").dialog('option', 'title', $("#txtDefn_Name").val());
			$("#outputoptions").hide();

			if (menu_isSSIMode() == false) {
				resizeGrid();
			}
			
			$("#reportworkframe").show();
			$("#divReportButtons").css("visibility", "visible");
			$("#divCrossTabOptions").css("visibility", "visible");

			$("#main").css("visibility", "visible");
			ShowDataFrame();

		} else {
			doExport();

			if (menu_isSSIMode()) {
				loadPartialView("linksMain", "Home", "workframe", null);
			} else {
				closeclick();
			}
		}
	}
	<%End If%>
</script>
