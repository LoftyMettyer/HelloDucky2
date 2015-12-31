<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl(of DMI.NET.Models.ObjectRequests.PromptedValuesModel)" %>
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
	<%=Html.AntiForgeryToken()%>
</form>

<div id="divUtilRunForm">
	<div class="absolutefull">
		<div class="pageTitleDiv">
			<a href='javascript:loadPartialView("linksMain", "Home", "workframe", null);' title='Back'>
				<i class='pageTitleIcon icon-circle-arrow-left'></i>
			</a>
			<span class="pageTitleSmaller" id="PageDivTitle">
				<% 
					
					If Model.UtilType = UtilityType.utlAbsenceBreakdown Or Model.UtilType = UtilityType.utlBradfordFactor Then
						Response.Write(GetReportNameByReportType(Model.UtilType))
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
				If Model.UtilType = UtilityType.utlCrossTab Then
					Html.RenderPartial("~/Views/Home/util_run_crosstabsMain.ascx")
				ElseIf Model.UtilType = UtilityType.utlCustomReport Then
					Html.RenderPartial("~/Views/Home/util_run_customreportsMain.ascx")
				ElseIf Model.UtilType = UtilityType.utlMailMerge Then
					Html.RenderPartial("~/Views/Home/util_run_mailmerge.ascx")
				ElseIf Model.UtilType = UtilityType.utlAbsenceBreakdown Then
					Html.RenderPartial("~/Views/Home/stdrpt_run_AbsenceBreakdown.ascx")
				ElseIf Model.UtilType = UtilityType.utlBradfordFactor Then
					Html.RenderPartial("~/Views/Home/util_run_customreportsMain.ascx")
				ElseIf Model.UtilType = UtilityType.utlCalendarReport Then
					Html.RenderPartial("~/Views/Home/util_run_calendarreport_main.ascx")
				ElseIf Model.UtilType = UtilityType.utlNineBoxGrid Then
					Html.RenderPartial("~/Views/Home/util_run_crosstabsMain.ascx")
				ElseIf Model.UtilType = UtilityType.TalentReport Then
					Html.RenderPartial("~/Views/Home/util_run_talentreport.ascx")
				ElseIf Model.UtilType = UtilityType.utlMatchReport Then
					Html.RenderPartial("~/Views/Home/util_run_talentreport.ascx")
				End If
			%>
		</div>
		<br/>
		<div id="divReportButtons" style="margin: 0; visibility: hidden; padding-top: 0; float: right">
			<%If Session("SSIMode") = True Then%>
				<%If Model.UtilType = UtilityType.utlCustomReport Then%> 
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


<%If Model.UtilType = UtilityType.utlMailMerge Then%>
	closeclick();

	<%ElseIf Model.UtilType = UtilityType.utlFilter And Model.FilteredAdd Then%>		
		if (OpenHR.currentWorkPage() === "UTIL_DEF_PICKLIST") {
			picklistdef_makeSelection('FILTER', <%:Model.ID%>, '<%=Session("promptsvalue")%>');
		} else {
			bulkbooking_makeSelection('FILTER', <%:Model.ID%>, '<%=Session("promptsvalue")%>');	
		}
	<%Else%>

	var isMobileDevice = ('<%=Session("isMobileDevice")%>' == 'True');
	// first get the size from the window
	// if that didn't work, get it from the body
	var size = {};
	
	<%If Model.UtilType = UtilityType.utlNineBoxGrid Then%>
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

	} 

	else {

		if ($("#txtPreview").val() == "True") {

			<%If Model.UtilType = UtilityType.utlCalendarReport Then%>
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
				<%If Model.UtilType = UtilityType.utlCustomReport Then%> 
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
				$('#gridReportData').setGridHeight(newHeight);
				$('#gridReportData').setGridWidth($('#reportframe').width() * 0.95);
				$('#gbox_gridReportData').css('margin', '0 auto'); //center grid in parent.
			}
			

			$('#cmdOutput').prop('disabled', isMobileDevice);
			$(".popup").dialog("option", "position", ['center', 'center']); //Center popup in screen
			$('.popup').bind('dialogclose', function() {
				closeclick();
			});

			if (menu_isSSIMode() == false) {
				$(".popup").dialog("open");
				$(".popup").dialog({ dialogClass: 'no-close' });
				$('#main').css('marginTop', '30px'); //.css('borderTop', '1px solid rgb(206, 206, 206)');
			}

			$("#PageDivTitle").text($("#txtDefn_Name").val());
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
