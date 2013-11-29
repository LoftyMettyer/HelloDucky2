﻿<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<script src="<%: Url.Content("~/Scripts/ctl_SetFont.js") %>" type="text/javascript"></script>

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
		Session("OutputOptions_Printer") = "false"
		Session("OutputOptions_Save") = "false"
		Session("OutputOptions_SaveExisting") = 0
		
	' following sessions vars:
	'
	' UtilType    - 0-13 (see UtilityType code in DATMGR .exe
	' UtilName    - <the name of the utility>
	' UtilID      - <the id of the utility>
	' Action      - run/delete
	Session("utiltype") = Request.Form("utiltype")
	Session("utilid") = Request.Form("utilid")
	Session("utilname") = Request.Form("utilname")
	Session("action") = Request.Form("action")
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
						aPrompts(1, j) = convertLocaleDateToSQL(Request.Form.Item(i))
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
		$("#cmdCancel").hide();
		$("#cmdOutput").show();

		$("#reportbreakdownframe").hide();
		$("#top").hide();
		$("#outputoptions").hide();
		$("#optionsframeset").show();
		$("#reportframe").show();
		$("#reportworkframe").show();

	}

	function ExportDataPrompt() {

		//var frmExportData = OpenHR.getForm("reportworkframe", "frmExportData");
		//OpenHR.submitForm(frmExportData, "outputoptions");

		$("#reportworkframe").hide();
		$("#reportbreakdownframe").hide();
		$("#outputoptions").show();
		
		$("#cmdOK").show();
		$("#cmdCancel").show();
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
			<a href='javascript:loadPartialView("linksMain", "Home", "workframe", null);' title='Home'>
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
		
		<div id="main" data-framesource="util_run" style="height: 80%; margin: 0 0 0 0; overflow-y:scroll; overflow-x: scroll">

			<%   
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
				Else
					' blah.
				End If
			%>
		</div>

		<div id="divReportButtons" style="margin: 5px 20px 0 25px; visibility: hidden">
			<div style="float: right;">
				<input class="btn" type="button" id="cmdOK" name="cmdOK" value="Output" onclick="outputOptionsOKClick()" />
				<input class="btn" type="button" id="cmdCancel" name="cmdCancel" value="Preview" onclick="ShowDataFrame();" />
				<input class="btn" type="button" id="cmdOutput" name="cmdOutput" value="Options" onclick="ExportDataPrompt();" />
				<input class="btn" type="button" id="cmdClose" name="cmdClose" value="Close" onclick="closeclick();"
				/>
			</div>
		</div>

	</div>
</div>



<script type="text/javascript">

	<%
	If Session("utiltype") = "17" Then
		If Session("CalendarReports_FailedOrNoRecords") = True Then 'We need a smaller popup window because the report failed or has no records
	%>
	$(".popup").dialog({ width: 750, height: 400, resizable: true });
	<%Else%>
	$(".popup").dialog({ width: 1100, height: 720, resizable: true });
	<%End If%>
	
	<%
	Session.Remove("CalendarReports_FailedOrNoRecords")
Else
	%>
	$(".popup").dialog({ width: 850, height: 720, resizable: true });
	<%
End If
	%>

	$(".popup").dialog("option", "position", ['center', 'center']); //Center popup in screen
	
	//$('.popup').css({ top: '50%', left: '50%', margin: '-' + ($('.popup').height() / 2) + 'px 0 0 -' + ($('.popup').width() / 2) + 'px' });

	$('.popup').bind('dialogclose', function (event) {
		closeclick();
	});

	if (menu_isSSIMode() == false) {
		$('#main').css('marginTop', '30px').css('borderTop', '1px solid rgb(206, 206, 206)');
	}

	$("#outputoptions").hide();
	$("#reportworkframe").show();

	menu_refreshMenu();

	$("#cmdCancel").hide();
	$("#cmdOK").hide();

</script>
