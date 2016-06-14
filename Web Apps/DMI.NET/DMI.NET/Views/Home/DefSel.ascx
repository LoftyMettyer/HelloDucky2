<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="HR.Intranet.Server.Enums" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="HR.Intranet.Server.Extensions" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.Linq" %>


<%="" %>

<%
    Dim SelectedTableID As String = Request.Form("SelectedTableID")
    Dim fGotId As Boolean
    Dim iBaseTableID As Integer
		
    Dim iDefSelType = CType(Session("defseltype"), UtilityType)
    Dim isLoadedFromReportDefiniton As Boolean = CType(Session("IsLoadedFromReportDefinition"), Boolean)

    Dim objSession As SessionInfo = CType(Session("sessionContext"), SessionInfo)
	
    Session("objCalendar" & Session("UtilID")) = Nothing

    If Session("fromMenu") = 0 Then
        If Session("singleRecordID") < 1 Then
            If Not String.IsNullOrEmpty(Request.Form("txtTableID")) Then
                iBaseTableID = Request.Form("txtTableID")
            Else

                If Len(Session("tableID")) > 0 Then
                    If CLng(Session("tableID")) > 0 Then
                        iBaseTableID = Session("tableID")
                        fGotId = True
                    End If
                End If

                If fGotId = False Then
                    If (Session("singleRecordID") > 0) Then
                        iBaseTableID = SettingsConfig.Personnel_EmpTableID
                    End If
                End If
            End If
        Else
            If Len(Session("tableID")) > 0 Then
                iBaseTableID = Session("tableID")
            End If
        End If
    End If
	
    If Session("singleRecordID") = 0 Then
        If Session("optionTableID") > 0 Then
            iBaseTableID = Session("optionTableID")
        End If
        Session("tableID") = Session("utilTableID")
    End If
	
    Session("optionTableID") = 0
	
    If iDefSelType = UtilityType.utlPicklist Or iDefSelType = UtilityType.utlFilter Or iDefSelType = UtilityType.utlCalculation Then
        iBaseTableID = CInt(Session("utilTableID"))
    End If
%>

<script type="text/javascript">

	var isNewPermitted = eval("<%:objSession.IsPermissionGranted(iDefSelType.ToSecurityPrefix, "NEW").ToString.ToLower%>");
	var isEditPermitted = eval("<%:objSession.IsPermissionGranted(iDefSelType.ToSecurityPrefix, "EDIT").ToString.ToLower%>");
	var isViewPermitted = eval("<%:objSession.IsPermissionGranted(iDefSelType.ToSecurityPrefix, "VIEW").ToString.ToLower%>");
	var isDeletePermitted = eval("<%:objSession.IsPermissionGranted(iDefSelType.ToSecurityPrefix, "DELETE").ToString.ToLower%>");
	var isRunPermitted = eval("<%:objSession.IsPermissionGranted(iDefSelType.ToSecurityPrefix, "RUN").ToString.ToLower%>");
	var isLoadedFromReportDefiniton = <%:isLoadedFromReportDefiniton.ToString.ToLower%>;
	
	var defSelType = "<%:iDefSelType%>";
	var menuSection = "Report";
	if ((defSelType === "utlMailMerge") || (defSelType === "utlWorkflow")) menuSection = "Utilities";
	if ((defSelType === "utlPicklist") || (defSelType === "utlFilter") || (defSelType === "utlCalculation")) menuSection = "Tools";

	function ssOleDBGridDefSelRecords_dblClick() {

		var frmDefSel = document.getElementById("frmDefSel");

		if ((frmDefSel.utiltype.value == 10) || (frmDefSel.utiltype.value == 11) || (frmDefSel.utiltype.value == 12)) {
			menu_MenuClick('mnutoolEditToolsFind');

		}
		else {
			// DblClick triggers Run after prompting for confirmation. 
			if (!isRunPermitted) {
				return (false);
			}

			var answer = 0;
			var sCaption = "";

			if (frmDefSel.utiltype.value == 1) {
				sCaption = "Cross Tab";
			}

			if (frmDefSel.utiltype.value == 2) {
				sCaption = "Custom Report";
			}
			if (frmDefSel.utiltype.value == 9) {
				sCaption = "Mail Merge";
			}
			if (frmDefSel.utiltype.value == 17) {
				sCaption = "Calendar Report";
			}
			if (frmDefSel.utiltype.value == 14) {
				sCaption = "Match Report";
			}
			if (frmDefSel.utiltype.value == 38) {
				sCaption = "Talent Report";
			}
			if (frmDefSel.utiltype.value == 25) {
				sCaption = " Workflow";
			}
			if (frmDefSel.utiltype.value == 35) {
				sCaption = "9-Box Grid Report";
			}
			if (frmDefSel.utiltype.value == 3) {
				sCaption = "Data Transfer";
			}

			var sMessage = "Are you sure you want to run " + $("#utilname").val() + " ?";
			OpenHR.modalPrompt(sMessage, 4, sCaption, setrun);

		}
		return false;
	}

	function ssOleDBGridDefSelRecords_rowcolchange() {
		var rowId = $("#DefSelRecords").getGridParam("selrow");
		var gridData = $("#DefSelRecords").getRowData(rowId);
		var username = $("#frmDefSel #txtusername").val();

		$("#txtDescription").val(gridData.description);

		// Populate the hidden fields with the selected utils information       
		$("#utilid").val($("#DefSelRecords").getGridParam("selrow"));
		$("#utilname").val(gridData.Name);

		if (isEditPermitted && (gridData.Username === username || allowEdit())) {
			menu_SetmnutoolButtonCaption("mnutoolEdit" + menuSection + "Find", "Edit");
		} else {
			menu_SetmnutoolButtonCaption("mnutoolEdit" + menuSection + "Find", "View");
		}

		menu_toolbarEnableItem("mnutoolDelete" + menuSection + "Find", isDeletePermitted && allowEdit());

	}

	function defsel_window_onload() {	
		
		var frmDefSel = document.getElementById('frmDefSel');

		// Expand the option frame and hide the work frame.
		if ( (parseInt($("#txtSingleRecordID").val()) > 0) || (parseInt($("#txtMultipleRecordIDs").val().length) > 0 ) )  {
			$("#optionframe").attr("data-framesource", "DEFSEL");
			$("#workframe").hide();
			$("#ToolsFrame").hide();
			$("#optionframe").show();
		} else {
			if (isLoadedFromReportDefiniton) {
				$("#ToolsFrame").attr("data-framesource", "TOOLS_SCREEN_LOADED_FROM_REPORT_DEFINITION");
				$("#workframe").hide();
				$("#ToolsFrame").show();
			} else {
				$("#workframe").attr("data-framesource", "DEFSEL");
				$("#optionframe").hide();
				$("#ToolsFrame").hide();
				$("#workframe").show();
			}
		}


		$("#DefSelRecords").jqGrid('bindKeys', {
			"onEnter": function (rowid) {
				ssOleDBGridDefSelRecords_dblClick();
			}
		});

		refreshControls();

		// Navbar options = i.e. search, edit, save etc 
		$("#DefSelRecords").jqGrid('navGrid', '#pager-coldata-defsel', { del: false, add: false, edit: false, search: false, refresh: false }); // setup the buttons we want
		$("#DefSelRecords").jqGrid('filterToolbar', { stringResult: true, searchOnEnter: false });  //instantiate toolbar so we can use toggle.

		if ($('#pager-coldata-defsel :has(".ui-icon-search")').length == 0) {
			$("#DefSelRecords").jqGrid('navButtonAdd', "#pager-coldata-defsel", {
				caption: '',
				buttonicon: 'ui-icon-search',
				position: 'first',
				onClickButton: function () {
					this.clearToolbar();
					this.toggleToolbar();
					if ($('.ui-search-toolbar', this.grid.hDiv).is(':visible')) {
						$('.ui-search-toolbar', this.grid.fhDiv).show();
					} else {
						$('.ui-search-toolbar', this.grid.fhDiv).hide();
					}
				},
				title: 'Search',
				cursor: 'pointer'
			});

			$('.ui-search-toolbar').hide(); // Hide it on setting up the grid - NB Remove this line to have it open on setup
		}

		$("#findGridRowDefsel").height("60%");
		$(window).bind('resize', function () {
			$("#DefSelRecords").setGridWidth($('#findGridRowDefsel').width(), true);
		}).trigger('resize');

		$("#DefSelRecords").closest('.ui-jqgrid-bdiv').width($("#DefSelRecords").closest('.ui-jqgrid-bdiv').width() + 1);

		if (rowCount() > 0) {

			var isSingleRecord = (parseInt($("#txtSingleRecordID").val()) <= 0);
			var gotoID;

			if (isSingleRecord === true) {
				gotoID = $("#lastSelectedID")[0].value;
				if (Number(gotoID) == 0) gotoID = $("#DefSelRecords").getDataIDs()[0];
			} else {
				gotoID = $("#DefSelRecords").getDataIDs()[0];
			}
			$("#DefSelRecords").jqGrid("setSelection", gotoID);

			// If no row is selected then select first row
			if ($("#DefSelRecords").getGridParam('selrow') == null) {
				$("#DefSelRecords").jqGrid("setSelection", $("#DefSelRecords").getDataIDs()[0]);
			}

		} else {
			//If the table is empty disable Copy, Edit, Delete and Properties buttons
			menu_toolbarEnableItem("mnutoolCopyToolsFind", false);
			menu_toolbarEnableItem("mnutoolEditToolsFind", false);
			menu_toolbarEnableItem("mnutoolDeleteToolsFind", false);
			menu_toolbarEnableItem("mnutoolPropertiesToolsFind", false);
		}
	}

	function rowCount() {
		return $("#DefSelRecords").jqGrid('getGridParam', 'records');
	}

	function disableNonDefselTabs() {
		$("#toolbarRecordFind").parent().hide();
		$("#toolbarRecord").parent().hide();
		$("#toolbarRecordAbsence").parent().hide();
		$("#toolbarRecordQuickFind").parent().hide();
		$("#toolbarRecordSortOrder").parent().hide();
		$("#toolbarRecordFilter").parent().hide();
		$("#toolbarRecordMailMerge").parent().hide();
		$("#toolbarReportNewEditCopy").parent().hide();
		$("#toolbarReportRun").parent().hide();
		$("#toolbarUtilitiesNewEditCopy").parent().hide();
		$("#toolbarEventLogView").parent().hide();
		$("#toolbarAdminConfig").parent().hide();
	}

	function refreshControls() {

		//show the Defsel-Find menu block.
		disableNonDefselTabs();		
		var fFromMenu = !((parseInt($("#txtSingleRecordID").val()) > 0) || (parseInt($("#txtMultipleRecordIDs").val()) > 0));				
		var fHasRows = (rowCount() > 0);
		var isRunOnly = (defSelType === "utlWorkflow" || defSelType === "utlDataTransfer");

		try {
			if (!fFromMenu) OpenHR.resetSession();	//reset session timeout for record edit. Well, try to.
		}
		catch(e) {}

		$("#toolbarUtilitiesFind").parent().hide();
		$("#toolbarToolsFind").parent().hide();
		$("#toolbarEventLogFind").parent().hide();
		$("#toolbarWFPendingStepsFind").parent().hide();
		$("#toolbarReportFind").parent().hide();
		
		menu_setVisibleMenuItem("mnutoolNew" + menuSection + "Find", !isRunOnly && fFromMenu);
		menu_setVisibleMenuItem("mnutoolCopy" + menuSection + "Find", !isRunOnly && fFromMenu);
		menu_setVisibleMenuItem("mnutoolEdit" + menuSection + "Find", !isRunOnly && fFromMenu);
		menu_setVisibleMenuItem("mnutoolDelete" + menuSection + "Find", !isRunOnly && fFromMenu);
		menu_setVisibleMenuItem("mnutoolProperties" + menuSection + "Find", !isRunOnly && fFromMenu);
		menu_setVisibleMenuItem("mnutoolRun" + menuSection + "Find", (menuSection !== "Tools"));

		// Show the close button for the Calendar, Absence breakdown, Bradford Factor Reports and Mail Mearge defsel when it's loaded from the database section. (E.g. from personnal record)
		// Show the close button for the tools (Picklist, Filter, Calculation) section if it's loaded from the report definition
		menu_setVisibleMenuItem("mnutoolClose" + menuSection + "Find", !fFromMenu || isLoadedFromReportDefiniton);

		// Set the report find toolbar group name to 'find' and hide the picklist/filter menu items
		menu_setVisibletoolbarGroupById('mnuSectionReportToolsFind', false);
		$('#toolbarReportFind').text('Find');

		if (!isRunOnly) {
			menu_toolbarEnableItem("mnutoolNew" + menuSection + "Find", isNewPermitted && fFromMenu);
			menu_toolbarEnableItem("mnutoolCopy" + menuSection + "Find", fHasRows && isNewPermitted && fFromMenu);
			menu_toolbarEnableItem("mnutoolEdit" + menuSection + "Find", fHasRows && (isEditPermitted || isViewPermitted) && fFromMenu);
			menu_toolbarEnableItem("mnutoolDelete" + menuSection + "Find", fHasRows && isDeletePermitted && fFromMenu);
			menu_toolbarEnableItem("mnutoolProperties" + menuSection + "Find", fHasRows && fFromMenu);
			if (menuSection !== "Tools") menu_toolbarEnableItem("mnutoolRun" + menuSection + "Find", fHasRows && isRunPermitted);
			if (menuSection !== "Tools") menu_toolbarEnableItem("mnutoolClose" + menuSection + "Find", !fFromMenu);
			if (defSelType === "17") menu_toolbarEnableItem("mnutoolRunReportFind", fHasRows && isRunPermitted); //Calendar Reports
		} else {
			menu_toolbarEnableItem("mnutoolRun" + menuSection + "Find", fHasRows && isRunPermitted);
			menu_toolbarEnableItem("mnutoolRunUtilitiesFind", fFromMenu);
			if (menuSection !== "Tools") { menu_toolbarEnableItem("mnutoolClose" + menuSection + "Find", !fFromMenu);}
		}

		// Finally show and select the tab
		$("#toolbar" + menuSection + "Find").parent().show();
		$("#toolbar" + menuSection + "Find").click();

		// If delete permission is given for the report but the 'Read Only' permission has been given in Group Access then disable the delete button
		if (fHasRows && isDeletePermitted && fFromMenu) {
			DisableDeleteButtonIfDefinitionHasReadOnlyAccess("mnutoolDeleteReportFind");
		}

	}


	// If the selected record has Read Only permission given in the Group Access then disable the delete button
	function DisableDeleteButtonIfDefinitionHasReadOnlyAccess(menuItem) {
		var rowId = $("#DefSelRecords").getGridParam('selrow');
		if (rowId != null) {
			var gridData = $("#DefSelRecords").getRowData(rowId);
			if (gridData.Access == 'ro') {
				menu_toolbarEnableItem(menuItem, false);
			}
		}
	}

	function showproperties() {
		
		if (!$("#mnutoolPropertiesUtil").hasClass("disabled")) {
			var id = $("#DefSelRecords").getGridParam('selrow');
			var type = document.getElementById('frmDefSel').utiltype.value;
			var name = $("#utilname").val();
			OpenHR.OpenDialog("DefinitionProperties?type=" + type, "divPopupReportDefinition", { ID: id, Name: encodeURIComponent(name), __RequestVerificationToken: $('[name="__RequestVerificationToken"]').val() }, '900px');
		}
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
			// Delay fixes a problem with IE7 and Vista (don't know why though!)
			pausecomp(300);
			win.window.focus();
		}
	}

	function ReturnNewWindow(mypage, myname, w, h, scroll) {
		var winl = (screen.width - w) / 2;
		var wint = (screen.height - h) / 2;
		var winprops = 'height=' + h + ',width=' + w + ',top=' + wint + ',left=' + winl + ',scrollbars=' + scroll + ',resizable';
		var win = window.open(mypage, myname, winprops);

		if (parseInt(navigator.appVersion) >= 4) {
			// Delay fixes a problem with IE7 and Vista (don't know why though!)
			pausecomp(300);
			win.window.focus();
		}

		return win;

	}

	function ToggleCheck() {	
		
		var piTableid = 0;
		var frmDefSel = document.getElementById('frmDefSel');

		if ((frmDefSel.utiltype.value == 10) || (frmDefSel.utiltype.value == 11) || (frmDefSel.utiltype.value == 12)) {
			piTableid = frmDefSel.selectTable.options[frmDefSel.selectTable.selectedIndex].value;
		}

		// Load the required definition selection screen
		var displayDiv = (parseInt($("#txtSingleRecordID").val()) === 0 ? "workframe" : "optionframe");
		var singleRecordId = parseInt($("#txtSingleRecordID").val());
		var multipleRecordIds = "";

		var currentPage = defsel_currentWorkFramePage();
		if (currentPage === "FIND") {
			if ((parseInt($("#txtMultipleRecordIDs").val().length) > 0)) 
			{
				displayDiv = "optionframe";
				singleRecordId = 0;
				multipleRecordIds = $("#txtMultipleRecordIDs").val();

				var frmFind = OpenHR.getForm("workframe", "frmFindForm");
				piTableid = frmFind.txtCurrentTableID.value;
			}
		}

		// If definition is of tools type and loaded from the report definition then set load it inside the Tools frame
		if ((defSelType === "utlPicklist") || (defSelType === "utlFilter") || (defSelType === "utlCalculation")) {
			if (isLoadedFromReportDefiniton) {
				displayDiv = "ToolsFrame";
			}
		}

		var postData = {
			txtTableID: piTableid,
			utiltype: frmDefSel.utiltype.value,
			OnlyMine: false,
			RecordID: singleRecordId,
			MultipleRecordIDs: multipleRecordIds,
			__RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
		};

		OpenHR.submitForm(null, displayDiv, null, postData, "DefSel");

	}

	function setdelete() {
		if (!$("#mnutoolDeleteUtil").hasClass("disabled")) {
			var frmDefSel = document.getElementById('frmDefSel');

			OpenHR.modalPrompt("Delete '" + escapeHTML($("#utilname").val()) + "'. Are you sure ?", 4, "Confirm").then(function (answer) {
				if (answer === 6) {

					var postData = {
						Action: "delete",
						utiltype: frmDefSel.utiltype.value,
						utilID: frmDefSel.utilid.value,
						utilName: $("#utilname").val(),
						txtTableID: frmDefSel.txtTableID.value,
						categoryId:frmDefSel.categoryId.value,
						ownerName:frmDefSel.ownerName.value,
						__RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
					}

					OpenHR.submitForm(null, "divPopupReportDefinition", null, postData, "defsel_submit");

				}
			});

		}
	}

	function setrun(answer) {

		if (answer === 6) {

			var postData;

			if (!$("#mnutillRunUtil").hasClass("disabled")) {
				var frmDefSel = document.getElementById('frmDefSel');

				if (frmDefSel.utiltype.value == 25) {
					// Workflow
					postData = {
						ID: frmDefSel.utilid.value,
						Name: $("#utilname").val(),
						__RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
					}

					OpenHR.submitForm(null, "optionframe", null, postData, "util_run_workflow");

				} else {

					postData = {
						UtilType: frmDefSel.utiltype.value,
						ID: frmDefSel.utilid.value,
						Name: $("#utilname").val(),
						__RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
					}
					OpenHR.submitForm(null, "reportframe", false, postData, "util_run_promptedValues");

				}
			}
		}
	}

	function setnew() {
		if (!$("#mnutoolNewUtil").hasClass("disabled")) {
			var frmDefSel = document.getElementById('frmDefSel');
			frmDefSel.action.value = "new";
			OpenHR.submitForm(frmDefSel);
		}
	}

	function setcopy() {
		if (!$("#mnutoolCopyUtil").hasClass("disabled")) {
			var frmDefSel = document.getElementById('frmDefSel');

			frmDefSel.action.value = "copy";
			OpenHR.submitForm(frmDefSel);
		}
	}

	function setedit() {

		if (!$("#mnutoolEditUtil").hasClass("disabled")) {
			var frmDefSel = document.getElementById('frmDefSel');

			if (allowEdit() && isEditPermitted) {
				frmDefSel.action.value = "edit";
				OpenHR.submitForm(frmDefSel);
			} else {
				frmDefSel.action.value = "view";
				OpenHR.submitForm(frmDefSel);
			}
		}
	}

	function setcancel() {

		if ( (parseInt($("#txtSingleRecordID").val()) > 0) || (parseInt($("#txtMultipleRecordIDs").val()) > 0) ) {			
			refreshData();
			menu_disableMenu();

			$("#optionframe").hide();
			$("#workframe").show();
			$("#toolbarRecord").show();
			$("#toolbarRecord").click();

			menu_refreshMenu();
		}
	}


	function loadEmptyOption() {
		$.ajax({
			url: 'emptyoption',
			type: "POST",
			dataType: 'html',
			async: true,
			success: function (html) {
				try {
					$('#optionframe').html('');
					$('#optionframe').html(html);
				} catch (e) { }
			}
		});
	}



	function defsel_currentWorkFramePage() {
		var sCurrentPage = $("#workframe").attr("data-framesource");
		try {
			sCurrentPage = sCurrentPage.toUpperCase();
		} catch (e) { }

		return sCurrentPage;
	}


</script>

<div id="defsel" data-framesource="defsel" style="display: block; height: 100%; width: 99.9%">

    <form name="frmDefSel" class="absolutefull" action="defsel_submit" method="post" id="frmDefSel">
        <div id="findGridRowDefsel" style="height: 70%; margin-right: 20px; margin-left: 20px;">

            <table width="100%" height="100%" class="invisible">
                <tr>
                    <td colspan="5" height="10">
                        <span class="pageTitle">
                            <%
                                If iDefSelType = UtilityType.utlCrossTab Then
                                    Response.Write("Cross Tabs")
                                ElseIf iDefSelType = UtilityType.utlCustomReport Then
                                    Response.Write("Custom Reports")
                                ElseIf iDefSelType = UtilityType.utlMailMerge Then
                                    Response.Write("Mail Merge")
                                ElseIf iDefSelType = UtilityType.utlPicklist Then
                                    Response.Write("Picklists")
                                ElseIf iDefSelType = UtilityType.utlFilter Then
                                    Response.Write("Filters")
                                ElseIf iDefSelType = UtilityType.utlCalculation Then
                                    Response.Write("Calculations")
                                ElseIf iDefSelType = UtilityType.utlCalendarReport Then
                                    Response.Write("Calendar Reports")
                                ElseIf iDefSelType = UtilityType.utlWorkflow Then
                                    Response.Write("Workflow")
                                ElseIf iDefSelType = UtilityType.utlNineBoxGrid Then
                                    Response.Write("9-Box Grid Reports")
                                ElseIf iDefSelType = UtilityType.utlDataTransfer Then
                                    Response.Write("Data Transfer")
                              ElseIf iDefSelType = UtilityType.utlMatchReport Then
                                Response.Write("Match Reports")
                              ElseIf iDefSelType = UtilityType.TalentReport Then
                            		Response.Write("Talent Reports")
                                End If
                            %>
                        </span>
                    </td>
                </tr>
                <%
                	Dim sErrorDescription = ""
                	Dim objDatabase As Database = CType(Session("DatabaseFunctions"), Database)
                	
                	Dim notFromRecordEditscreen As Boolean
                	notFromRecordEditscreen = If((Not Session("multipleRecordIDs") Is Nothing AndAlso Session("multipleRecordIDs").ToString().Length > 0) Or (Not Session("singleRecordID") Is Nothing AndAlso Session("singleRecordID") > 0 ), False, True)

                	
                	
                	If iDefSelType <> UtilityType.utlWorkflow Then
                %>
                <tr>
                    <td height="10" colspan="5">
                        <table width="90%" class="invisible">
                            <tr>
                                <td>
                                    <div id="Container" class="width100 floatleft" style="width:800px;">
												
                                        <div id="divOwner" style="display: inline-block; margin-right: 50px">
                                            <label>Owner :</label>
                                            <select id="selectOwners" name="selectOwners" class="combo" style="margin-left: 4px; width: 200px">
                                                <%
                                                	If Len(sErrorDescription) = 0 Then
                                                		Dim rstOwners = objDatabase.DB.GetDataTable("sp_ASRGetAllObjectNames", CommandType.StoredProcedure)
                                                		Dim userName = LCase(Session("Username"))
                                                		Dim ownerName = Session("defsel_ownerName_" & iDefSelType)
                                                		Dim onlyMine = Session("OnlyMine") '
                                                		
                                                		' Delete the LoggedinUser from dataset(i.e from rstOwners) as we need to add 'Mine' as prefix to it.
                                                		' Display 'Mine(loggedinUser)' as selected in dropdown on defsel selection screen, if reports/utilities checkbox is checked in user configuration screen.
                                                		Dim foundRow As DataRow() = rstOwners.Select("Username='" + userName + "' ")
                                                		If foundRow.Count > 0 Then foundRow(0).Delete()
                                                		rstOwners.AcceptChanges()
                                                		                                                		
                                                		Response.Write("<option selected>All</option>" & vbCrLf)
                                                		
                                                		If ((onlyMine AndAlso ownerName = "") Or (onlyMine AndAlso Not notFromRecordEditscreen)) Then
                                                			'Display 'Mine(loggedinUser)' as selected in dropdown when user click on report/utilitie menu
                                                			Response.Write("<option SELECTED>Mine (" & userName & ")</option>" & vbCrLf)
                                                        Else
                                                            If (Not onlyMine AndAlso Not notFromRecordEditscreen) Then
                                                                ownerName = ""
                                                            End If
                                                            If Not userName Is Nothing AndAlso Len(userName) > 0 AndAlso userName <> LCase(ownerName) Then
                                                                Response.Write("<option>Mine (" & userName & ")</option>" & vbCrLf)
                                                            ElseIf (Not userName Is Nothing) AndAlso Not ownerName Is Nothing AndAlso userName = LCase(ownerName) Then
                                                                Response.Write("<option SELECTED>Mine (" & userName & ")</option>" & vbCrLf)
                                                            End If
                                                        End If
 		
                                                        Dim count As Integer = 0
                                                        For Each objRow As DataRow In rstOwners.Rows
                                                				
                                                            If Not userName Is Nothing _
                                                                AndAlso LCase(userName) <> LCase(Replace(objRow(0).ToString(), """", "&quot;")) Then
                                                                Response.Write("<option ")
                                                                If Not ownerName Is Nothing _
                                                                            AndAlso LCase(ownerName) = LCase(Replace(objRow(0).ToString(), """", "&quot;")) Then
                                                                    If (notFromRecordEditscreen) Then
                                                                        Response.Write(" SELECTED")
                                                                    End If
                                                                End If
                                                                Response.Write(">" & Replace(objRow(0).ToString(), """", "&quot;") & "</option>" & vbCrLf)
                                                				
                                                            End If
                                                			
                                                			
                                                        Next
                                                		
                                                    End If
                                                %>
                                            </select>
                                        </div>

                                        <%If notFromRecordEditscreen %>
                                        <div id="divCategory" style="display: inline-block; margin-right: 50px">
                                            <label>Category :</label>
                                            <%
                                            	Response.Write("<select id=selectCategories name=selectCategories class=combo style=margin-left: 4px; > ")
                                            	If Len(sErrorDescription) = 0 Then
                                            		Dim rstObjectCategories = objDatabase.DB.GetDataTable("spsys_getobjectcategories", CommandType.StoredProcedure _
                                            										, New SqlParameter("utilityType", SqlDbType.Int) With {.Value = 0} _
                                            										, New SqlParameter("UtilityID", SqlDbType.Int) With {.Value = 0} _
                                            										, New SqlParameter("tableID", SqlDbType.Int) With {.Value = 0})
												
                                            		Dim categoryID = Session("defsel_categoryId_" & iDefSelType)
                                            		
                                            		Response.Write("<option value=-1 selected>All</option>" & vbCrLf)
                                            		Response.Write("<option value=0")
                                            		If Not categoryID Is Nothing AndAlso CInt(categoryID) = 0 Then
                                            			Response.Write(" SELECTED")
                                            		End If
                                            		Response.Write(">None</option>" & vbCrLf)
												
                                            		For Each objRow As DataRow In rstObjectCategories.Rows
                                            			Response.Write("<option value=" & objRow(0))
                                            			If Not categoryID Is Nothing AndAlso CInt(objRow(0)) = CInt(categoryID) Then
                                            				Response.Write(" SELECTED")
                                            			End If
                                            			Response.Write(">" & Replace(objRow(1).ToString(), """", "&quot;") & "</option>" & vbCrLf)
                                            			
                                            		Next
                                            	
                                            	End If
                                            	
                                            	Response.Write("  </select>")
                                            
                                            %>
                                        </div>
										<%End If %>

                                        <% 							
                                        	If iDefSelType = UtilityType.utlPicklist Or iDefSelType = UtilityType.utlFilter Or iDefSelType = UtilityType.utlCalculation OR (Not notFromRecordEditscreen) Then
                                        %>

                                        <div id="divTable" style="display: inline-block;" class="invisible">
                                            <label>Table :</label>																					
																					<%
																						
																						Response.Write("<select id=selectTable name=selectTable style='height:22px;' width:'200px;' class='combo ")
																						If (Not notFromRecordEditscreen) Then
																							Response.Write("dropdownDisable' disabled=disabled >")
																					        
																						Else
																							Response.Write("' >")
																						End If
																						
																						Try

																							If iDefSelType = UtilityType.utlCalculation Then
																								Response.Write("<option value=0>None</option>")
																							End If

																							For Each objTable In objSession.Tables.OrderBy(Function(t) t.Name) 'Order by table name
																												
																								Response.Write("						<option value=" & objTable.ID)
																								If SelectedTableID Is Nothing Or SelectedTableID = "" Then
																									If objTable.ID = iBaseTableID Then
																										Response.Write(" SELECTED")
																									End If
																								Else
																									If objTable.ID = CLng(SelectedTableID) Then
																										Response.Write(" SELECTED")
																									End If
																								End If

																								Response.Write(">" & Replace(objTable.Name, "_", " ") & "</option>" & vbCrLf)

																							Next
				
																						Catch ex As Exception
																							sErrorDescription = "The table records could not be retrieved." & vbCrLf & ex.Message

																						End Try
																						Response.Write("  </select>")
																					End If
                                            %>
                                        </div>
                                        <br />
                                        <br />
                                    </div>
                                </td>
                            </tr>
                        </table>
                    </td>

                </tr>
                <%
                End If
                %>

							<tr>

								<td width="100%">
									<table height="100%" width="100%">
										<tr>
											<td width="100%">
												<table id="DefSelRecords"></table>
												<div id='pager-coldata-defsel'></div>
											</td>
										</tr>

										<tr height="10">
											<td></td>
										</tr>

										<tr>
											<td height="70">												
												<textarea cols="20" class="txtareaDisable" name="txtDescription" id="txtDescription" rows="4" tabindex="-1" 
													readonly="readonly">
													</textarea>
											</td>
										</tr>
									</table>
								</td>

								<td width="80" style="display: none;"></td>
								<td width="20">&nbsp;&nbsp;&nbsp;&nbsp;</td>
							</tr>

                <tr>
                    <td colspan="5" height="10"
                        <%
                    	If iDefSelType = UtilityType.utlWorkflow Then
                    		Response.Write(" style=""visibility:hidden""")
                    	End If%>>
                        <input type='hidden' id="txtusername" name="txtusername" value="<%=lcase(session("Username"))%>">
                    </td>
                </tr>
            </table>

					<input type="hidden" id="utiltype" name="utiltype" value="<%:CInt(iDefSelType)%>">
					<input type="hidden" id="utilid" name="utilid" value='<%:Session("utilid")%>'>
					<input type="hidden" id="utilname" name="utilname">
					<input type="hidden" id="action" name="action">
					<input type="hidden" id="txtTableID" name="txtTableID" value='<%=iBaseTableID%>'>
					<input type="hidden" id="categoryId" name="categoryId" value='<%:Session("defsel_categoryId_" & iDefSelType)%>'>
					<input type="hidden" id="ownerName" name="ownerName" value='<%:Session("defsel_ownerName_" & iDefSelType)%>'>
        </div>
        <%=Html.AntiForgeryToken()%>
    </form>


    <input type="hidden" id="txtSingleRecordID" name="txtSingleRecordID" value='<%:session("singleRecordID")%>'>
    <input type="hidden" id="txtMultipleRecordIDs" name="txtMultipleRecordIDs" value='<%:session("multipleRecordIDs")%>'>
    <input type="hidden" id="txtTicker" name="txtTicker" value="0">
    <input type="hidden" id="txtLastKeyFind" name="txtLastKeyFind" value="">

    <input type="hidden" id="lastSelectedID" name="lastSelectedID" value='<%=Session("utilid")%>'>
</div>


<script>
	$("#DefSelRecords").keydown(function (event) {
		//Add first letter search to the grid...
		try {
			var id = $('#DefSelRecords td:visible').filter(function () {
				return $(this).text().substring(0, 1).toLowerCase() == String.fromCharCode(event.which).toLowerCase();
			}).first().closest('tr').attr('id');
			if (Number(id) > 0)
				$("#DefSelRecords").jqGrid('setSelection', id);
		}
		catch (e) { }
	});

	function attachDefSelGrid() {	 
			
		var onlyMine = false;        
		var categoryID =  $("#selectCategories option:selected").val();	
		var ownerName = $.trim($("#selectOwners option:selected").text());
		if (ownerName.indexOf('Mine (') > -1)
		{			
			ownerName = ownerName.replace('Mine (', '').replace(')','');			
		}		
		if (categoryID === undefined && ownerName === '')
		{
			categoryID = -1;
			ownerName = '';
		}
		else if (categoryID === undefined && ownerName != '')
		{
			//come from recordedit screen
			categoryID = -1;
		}
		var defSelType = "<%:iDefSelType%>";		
		
    	if ( defSelType === "utlCustomReport" || defSelType === "utlCrossTab" || defSelType === "utlCalendarReport" || defSelType === "utlMailMerge" || defSelType === "utlNineBoxGrid" || defSelType === "TalentReport" || defSelType === "utlMatchReport")
    	{
    		$("#divCategory").show();
    	}
    	else
    	{
    		$("#divCategory").hide();
    	}

    	$("#categoryId")[0].value = categoryID;
    	$("#ownerName")[0].value = ownerName;

    	//resize grid		
    	var gridWidth = $("#findGridRowDefsel").width();
    	var gridHeight = $("#workframeset").height() * 0.6;	//findGridRow is hardcoded to 60% of workframeset.
		 
    	$("#DefSelRecords").jqGrid({			
    		url: 'GetDefinitionsForType?UtilType=' + <%:CInt(iDefSelType)%> + '&&TableID=' + <%=iBaseTableID%> + '&&OnlyMine=' + onlyMine + '&&CategoryID=' + categoryID +'&&OwnerName=' + ownerName,
        	datatype: 'json',
        	mtype: 'GET',
        	jsonReader: {
        		root: "rows", //array containing actual data
        		page: "page", //current page
        		total: "total", //total pages for the query
        		records: "records", //total number of records
        		repeatitems: false,
        		id: "ID"
        	},
        	colNames: ['ID', 'Name', 'description', 'Username', 'Access'],
        	colModel: [
				{ name: 'ID', index: 'ID', hidden: true },
				{ name: 'Name', index: 'Name', width: 40, sortable: false },
				{ name: 'description', index: 'description', hidden: true },
				{ name: 'Username', index: 'Username', hidden: true },
				{ name: 'Access', index: 'Access', hidden: true }],
        	viewrecords: false,
        	width: gridWidth,
        	height: gridHeight,
        	sortname: 'Name',
        	sortorder: "asc",
        	rowNum: 10000,
        	cmTemplate: { sortable: false },
        	ignoreCase: true,
        	onSelectRow: function (rowID) {
        		ssOleDBGridDefSelRecords_rowcolchange();
        	},
        	ondblClickRow: function (rowID) {
        		ssOleDBGridDefSelRecords_dblClick();
        	},
        	loadComplete: function (json) {				
        		defsel_window_onload();				
        	},
        	rowTotal: 50,
        	rowList: [],
        	pager: $('#pager-coldata-defsel'),
        	pgbuttons: false,
        	pgtext: null,
        	loadonce: true,
        	autoencode: true,
        	loadui: "disable"			
        });

    	// If no records found than clear the previously selected definition description
		if ($("#DefSelRecords").getGridParam("reccount")  == 0 )
		{
			$("#txtDescription").val('');
		}
	}

	$(function () {

		attachDefSelGrid();

		$("#selectTable").change(function () {			
			ToggleCheck();
		});

		$("#selectCategories").change(function () {				
			$('#DefSelRecords').jqGrid('GridUnload');
			attachDefSelGrid();
		});

		$("#selectOwners").change(function () {				
			$('#DefSelRecords').jqGrid('GridUnload');
			attachDefSelGrid();
		});

		$('#txtDescription').on('mouseover', function (e) {
			$(this).is('[readonly]') ? $(this).addClass("cursorNotAllowed") : $(this).removeClass("cursorNotAllowed");
		});

		$('#txtDescription').focus(function(){
			$(this).is('[readonly]') ? this.blur() : "";
		});

	});

	function allowEdit() {
		var rowId = $("#DefSelRecords").getGridParam("selrow");
		var gridData = $("#DefSelRecords").getRowData(rowId);
		if (gridData.Access === "ro") {
			return false;
		}
		return true;
	}


	// Close the Tools Screen (Picklists/Filter/Calculation) & clear the tools frame and return the user to the same point they left in the original report definition screen.
	function closeTools() {
		
		if (isLoadedFromReportDefiniton && !$("#mnutoolCloseToolsFind").hasClass("disabled")) {

			var absenceBreakdownOrBreadfordFactorForm = OpenHR.getForm("workframe", "frmAbsenceDefinition");
			var reportDefinitionForm = OpenHR.getForm("workframe", "frmReportDefintion");
			var utilType = null;
			
			// Hide the find Group			
			$("#toolbarToolsFind").parent().hide();

			if (absenceBreakdownOrBreadfordFactorForm != null && reportDefinitionForm == null) {

				// Refresh ribbon toolbar for the bradford factor and absence breakdown reports 			
				$("#toolbarReportFind").parent().show();
				$("#toolbarReportFind").click();

				// Sets the ribbon buttons
				SetsRibbonButtonsForAbsenceBreakdownAndBradfordFactor();
			} 
			else if(reportDefinitionForm != null && absenceBreakdownOrBreadfordFactorForm == null) 
			{
				// Refresh ribbon toolbar 
				$("#toolbarReportNewEditCopy").parent().show();
				$("#toolbarReportNewEditCopy").click();

				// Reset the utility type to source report definition type (From where we came from to the tools screen)
				utilType = reportDefinitionForm.txtReportType.value;
			}

			// Post data to reset the session variables which indicates that we have loaded the defsel into tools frame from the report definition.
			var postSessionData = {
				utiltype: utilType,
				isLoadedFromReportDefinition: false,
				__RequestVerificationToken: $('[name="__RequestVerificationToken"]').val()
			};

			OpenHR.submitForm(null, null, false, postSessionData, "ResetPageSourceFlag", function() {

				// Hide & Clear the Toolsframe and show the Workframe
				$("#ToolsFrame").html('');
				$("#ToolsFrame").hide();
				$("#workframe").show();

				if (absenceBreakdownOrBreadfordFactorForm == null && reportDefinitionForm != null) {

					// Show/Hide the picklist/filter/calculation ribbon button
					ShowHideToolsButtons();

					// Enable/Disable the picklist/filter/calculation ribbon button buttons based on the permissions granted		  
					EnableDisableToolsButtons();
				}
			});
		}
	}

</script>