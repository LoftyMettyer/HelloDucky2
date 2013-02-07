<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<%
	Dim fGotId As Boolean
	Dim sTemp As String
	
	Session("objCalendar" & Session("UtilID")) = Nothing
	
	If Not String.IsNullOrEmpty(Request.Form("OnlyMine")) Then
		Session("OnlyMine") = Request.Form("OnlyMine")
	Else
		If Session("fromMenu") = 1 Then
			' Read the defSel 'only mine' setting from the database.
			sTemp = "onlymine "
			Select Case Session("defseltype")
				Case 0
					sTemp = sTemp & "BatchJobs"
				Case 1
					sTemp = sTemp & "CrossTabs"
				Case 2
					sTemp = sTemp & "CustomReports"
				Case 3
					sTemp = sTemp & "DataTransfer"
				Case 4
					sTemp = sTemp & "Export"
				Case 5
					sTemp = sTemp & "GlobalAdd"
				Case 6
					sTemp = sTemp & "GlobalDelete"
				Case 7
					sTemp = sTemp & "GlobalUpdate"
				Case 8
					sTemp = sTemp & "Import"
				Case 9
					sTemp = sTemp & "MailMerge"
				Case 10
					sTemp = sTemp & "Picklists"
				Case 11
					sTemp = sTemp & "Filters"
				Case 12
					sTemp = sTemp & "Calculations"
				Case 17
					sTemp = sTemp & "CalendarReports"
				Case 25
					sTemp = sTemp & "Workflow"
			End Select

			' NB. The numeric codes used for the utilities are taken from
			' the UtilityType Public Enum in Data Manager. The list is :
			'	utlBatchJob = 0
			'	utlCrossTab = 1
			'	utlCustomReport = 2
			'	utlDataTransfer = 3
			'	utlExport = 4
			'	utlGlobalAdd = 5
			'	utlGlobalDelete = 6
			'	utlGlobalUpdate = 7
			'	utlImport = 8
			'	utlMailMerge = 9
			'	utlPicklist = 10
			'	utlFilter = 11
			'	utlCalculation = 12
			'	utlOrder = 13
			'	utlMatchReport = 14
			'	utlAbsenceBreakdown = 15
			'	utlBradfordFactor = 16
			'	utlCalendarReport = 17
			'	utlLabel = 18
			'	utlLabelType = 19
			'	utlRecordProfile = 20
			'	utlEmailAddress = 21
			'	utlEmailGroup = 22
			'	utlWorkflow = 25
					
			Dim cmdDefSelOnlyMine = CreateObject("ADODB.Command")
			cmdDefSelOnlyMine.CommandText = "sp_ASRIntGetSetting"
			cmdDefSelOnlyMine.CommandType = 4 ' Stored procedure.
			cmdDefSelOnlyMine.ActiveConnection = Session("databaseConnection")

			Dim prmSection = cmdDefSelOnlyMine.CreateParameter("section", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
			cmdDefSelOnlyMine.Parameters.Append(prmSection)
			prmSection.value = "defsel"

			Dim prmKey = cmdDefSelOnlyMine.CreateParameter("key", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
			cmdDefSelOnlyMine.Parameters.Append(prmKey)
			prmKey.value = sTemp

			Dim prmDefault = cmdDefSelOnlyMine.CreateParameter("default", 200, 1, 8000)	' 200=varchar, 1=input, 8000=size
			cmdDefSelOnlyMine.Parameters.Append(prmDefault)
			prmDefault.value = "0"

			Dim prmUserSetting = cmdDefSelOnlyMine.CreateParameter("userSetting", 11, 1) ' 11=bit, 1=input
			cmdDefSelOnlyMine.Parameters.Append(prmUserSetting)
			prmUserSetting.value = 1

			Dim prmResult = cmdDefSelOnlyMine.CreateParameter("result", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
			cmdDefSelOnlyMine.Parameters.Append(prmResult)

			Err.Clear()
			cmdDefSelOnlyMine.Execute()
			Session("OnlyMine") = (CLng(cmdDefSelOnlyMine.Parameters("result").Value) = 1)

			cmdDefSelOnlyMine = Nothing
		Else
			If CStr(Session("OnlyMine")) = "" Then Session("OnlyMine") = False
		End If
	End If
	Session("fromMenu") = 0

	If (CStr(Session("singleRecordID")) = "") Or (Session("singleRecordID") < 1) Then
		If Not String.IsNullOrEmpty(Request.Form("txtTableID")) Then
			Session("utilTableID") = Request.Form("txtTableID")
		Else
			If Len(Session("tableID")) > 0 Then
				If CLng(Session("tableID")) > 0 Then
					Session("utilTableID") = Session("tableID")
					fGotId = True
				End If
			End If

			If fGotId = False Then
				If (CStr(Session("optionDefSelRecordID")) <> "") Then
					If (Session("optionDefSelRecordID") > 0) Then
						Session("utilTableID") = Session("Personnel_EmpTableID")
					Else
						If (Session("Personnel_EmpTableID") > 0) Then
							Session("utilTableID") = Session("Personnel_EmpTableID")
						Else
							Session("utilTableID") = 0
						End If
					End If
				Else
					If (Session("Personnel_EmpTableID") > 0) Then
						Session("utilTableID") = Session("Personnel_EmpTableID")
					Else
						Session("utilTableID") = 0
					End If
				End If
			End If
		End If
	End If
	
	If CStr(Session("optionDefSelType")) <> "" Then
		Session("defseltype") = Session("optionDefSelType")
	End If

	session("tableID") = session("utilTableID")
	
	If CStr(Session("singleRecordID")) <> "" Then
		If Session("singleRecordID") < 1 Then
			If CStr(Session("optionDefSelRecordID")) <> "" Then
				If Session("optionDefSelRecordID") > 0 Then
					Session("singleRecordID") = Session("optionDefSelRecordID")
				End If
			End If
		Else
			If CStr(Session("optionTableID")) <> "" Then
				If Session("optionTableID") > 0 Then
					Session("utilTableID") = Session("optionTableID")
				End If
			End If
			Session("tableID") = Session("utilTableID")
		End If
	Else
		Session("singleRecordID") = 0
	End If
	
	Session("optionDefSelType") = ""
	Session("optionTableID") = ""
	Session("optionDefSelRecordID") = ""
	
	If CStr(Session("utilTableID")) = "" Then
		Session("utilTableID") = 0
	End If

	if (session("defseltype") <> 10) and (session("defseltype") <> 11) and (session("defseltype") <> 12) then
		if (session("singleRecordID") < 1) then
			session("utilTableID") = 0
		end if
	end if 
%>

<script type="text/javascript">
<!--
	function defsel_window_onload() {
		
		var frmDefSel = document.getElementById('frmDefSel');
		
		$("#workframe").attr("data-framesource", "DEFSEL");
	    
		if (frmDefSel.txtSingleRecordID.value > 0) {
			// Expand the option frame and hide the work frame.
			//TODO
			//window.parent.document.all.item("workframeset").cols = "1, *";

		}

		var sControlName;
		var sControlPrefix;
		
		frmDefSel.ssOleDBGridDefSelRecords.focus();
		frmDefSel.cmdCancel.focus();

		setGridFont(frmDefSel.ssOleDBGridDefSelRecords);

		var controlCollection = frmDefSel.elements;
		if (controlCollection != null) {
			for (var i = 0; i < controlCollection.length; i++) {
				sControlName = controlCollection.item(i).name;
				sControlPrefix = sControlName.substr(0, 13);

				if (sControlPrefix == "txtAddString_") {
					frmDefSel.ssOleDBGridDefSelRecords.AddItem(controlCollection.item(i).value);
				}
			}
		}

		if (frmDefSel.ssOleDBGridDefSelRecords.rows > 0) {
			// Need to refresh the grid before we movefirst.
			frmDefSel.ssOleDBGridDefSelRecords.refresh();

			if (frmDefSel.utilid.value > 0) {
				// Try to select the current record.
				locateRecordID(frmDefSel.utilid.value, true);
			} else {
				// Select the top row.
				frmDefSel.ssOleDBGridDefSelRecords.MoveFirst();
				frmDefSel.ssOleDBGridDefSelRecords.SelBookmarks.Add(frmDefSel.ssOleDBGridDefSelRecords.Bookmark);
			}
		}

		if ((frmDefSel.txtTableID.value == 0) && ((frmDefSel.utiltype.value == 10) || (frmDefSel.utiltype.value == 11) || (frmDefSel.utiltype.value == 12))) {
			frmDefSel.txtTableID.value = frmDefSel.selectTable.options[frmDefSel.selectTable.selectedIndex].value;
		}

		refreshControls();

		if (frmDefSel.txtSingleRecordID.value > 0) {
			// Expand the option frame and hide the work frame.
			//TODO
			//window.parent.frames("menuframe").disableMenu();
		} else {
			//TODO
			//window.parent.frames("menuframe").refreshMenu();
		}

		if (frmDefSel.txtSingleRecordID.value > 0) {
			// Expand the option frame and hide the work frame.
			//TODO
			//window.parent.document.all.item("workframeset").cols = "1, *";
		} else {
			//TODO
			//window.parent.document.all.item("workframeset").cols = "*, 0";
		}
	}
-->	
</script>

<script type="text/javascript">
	/* Sequential search the grid for the required ID. */
	function locateRecordID(psSearchFor, pfIdMatch) {
		var fFound;
		var iIndex;
		var iIdColumnIndex;
		var sColumnName;
		var frmDefSel = document.getElementById('frmDefSel');
		
		fFound = false;

		frmDefSel.ssOleDBGridDefSelRecords.redraw = false;

		if (pfIdMatch == true) {
			// Locate the ID column in the grid.
			iIdColumnIndex = -1;
			for (iIndex = 0; iIndex < frmDefSel.ssOleDBGridDefSelRecords.Cols; iIndex++) {
				sColumnName = frmDefSel.ssOleDBGridDefSelRecords.Columns(iIndex).Name;
				if (sColumnName.toUpperCase() == "ID") {
					iIdColumnIndex = iIndex;
					break;
				}
			}

			if (iIdColumnIndex >= 0) {
				frmDefSel.ssOleDBGridDefSelRecords.MoveLast();
				frmDefSel.ssOleDBGridDefSelRecords.MoveFirst();

				for (iIndex = 1; iIndex <= frmDefSel.ssOleDBGridDefSelRecords.rows; iIndex++) {
					if (frmDefSel.ssOleDBGridDefSelRecords.Columns(iIdColumnIndex).value == psSearchFor) {
						frmDefSel.ssOleDBGridDefSelRecords.SelBookmarks.Add(frmDefSel.ssOleDBGridDefSelRecords.Bookmark);
						fFound = true;
						break;
					}

					if (iIndex < frmDefSel.ssOleDBGridDefSelRecords.rows) {
						frmDefSel.ssOleDBGridDefSelRecords.MoveNext();
					}
					else {
						break;
					}
				}
			}
		}
		else {
			for (iIndex = 1; iIndex <= frmDefSel.ssOleDBGridDefSelRecords.rows; iIndex++) {
				var sGridValue = new String(frmDefSel.ssOleDBGridDefSelRecords.Columns(0).value);
				sGridValue = sGridValue.substr(0, psSearchFor.length).toUpperCase();
				if (sGridValue == psSearchFor.toUpperCase()) {
					frmDefSel.ssOleDBGridDefSelRecords.SelBookmarks.Add(frmDefSel.ssOleDBGridDefSelRecords.Bookmark);
					fFound = true;
					break;
				}

				if (iIndex < frmDefSel.ssOleDBGridDefSelRecords.rows) {
					frmDefSel.ssOleDBGridDefSelRecords.MoveNext();
				}
				else {
					break;
				}
			}
		}

		if ((fFound == false) && (frmDefSel.ssOleDBGridDefSelRecords.rows > 0)) {
			// Select the top row.
			frmDefSel.ssOleDBGridDefSelRecords.MoveFirst();
			frmDefSel.ssOleDBGridDefSelRecords.SelBookmarks.Add(frmDefSel.ssOleDBGridDefSelRecords.Bookmark);
		}

		frmDefSel.ssOleDBGridDefSelRecords.redraw = true;
	}

	function refreshControls() {
	    
	    //show the utilities menu block.
	    $("#mnuSectionUtilities").show();
	    $("#toolbarHome").click();

		var fNoneSelected;
		var frmpermissions = document.getElementById('frmpermissions');
		var frmDefSel = document.getElementById('frmDefSel');
		
		fNoneSelected = (frmDefSel.ssOleDBGridDefSelRecords.SelBookmarks.Count == 0);

		button_disable(frmDefSel.cmdEdit, (fNoneSelected ||
			((frmpermissions.grantedit.value == 0) && (frmpermissions.grantview.value == 0))));
		button_disable(frmDefSel.cmdNew, (frmpermissions.grantnew.value == 0));
		button_disable(frmDefSel.cmdCopy, (fNoneSelected || (frmpermissions.grantnew.value == 0)));
		button_disable(frmDefSel.cmdDelete, (fNoneSelected ||
			(frmpermissions.grantdelete.value == 0) ||
				(frmDefSel.cmdEdit.value.toUpperCase() == "VIEW")));

		if (((frmpermissions.grantedit.value == 0) &&
			(frmpermissions.grantview.value == 1)) ||
				(frmDefSel.cmdEdit.value.toUpperCase() == "VIEW")) {
			frmDefSel.cmdEdit.value = "View";
		}
		else {
			frmDefSel.cmdEdit.value = "Edit";
		}

		button_disable(frmDefSel.cmdProperties, (fNoneSelected ||
			((frmpermissions.grantnew.value == 0) &&
				(frmpermissions.grantedit.value == 0) &&
					(frmpermissions.grantview.value == 0) &&
						(frmpermissions.grantdelete.value == 0) &&
							(frmpermissions.grantrun.value == 0))));
		button_disable(frmDefSel.cmdRun, (fNoneSelected || (frmpermissions.grantrun.value == 0)));
	}

	function showproperties() {
		var sUrl;
		var frmDefSel = document.getElementById('frmDefSel');
		
		if (frmDefSel.ssOleDBGridDefSelRecords.selbookmarks.count > 0) {
			var frmProp = document.getElementById('frmProp');
			frmProp.prop_id.value = frmDefSel.ssOleDBGridDefSelRecords.Columns("id").Value;
			frmProp.prop_name.value = frmDefSel.ssOleDBGridDefSelRecords.Columns("name").Value;
			frmProp.utiltype.value = frmDefSel.utiltype.value;

			sUrl = "defselproperties" +
				"?prop_name=" + escape(frmProp.prop_name.value) +
					"&prop_id=" + frmProp.prop_id.value +
						"&utiltype=" + frmProp.utiltype.value;
			openDialog(sUrl, 500, 230);
			return false;
		}
		else {
			OpenHR.messageBox("You must select a definition to view", 48, "OpenHR Intranet");
		}
		return false;
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

	function openDialog(pDestination, pWidth, pHeight) {
		var dlgwinprops = "center:yes;" +
			"dialogHeight:" + pHeight + "px;" +
				"dialogWidth:" + pWidth + "px;" +
					"help:no;" +
						"resizable:yes;" +
							"scroll:yes;" +
								"status:no;";
		window.showModalDialog(pDestination, self, dlgwinprops);
	}

	function ToggleCheck() {
		
		var frmOnlyMine = document.getElementById('frmOnlyMine');
		var frmDefSel = document.getElementById('frmDefSel');
		
		if ((frmDefSel.utiltype.value == 10) || (frmDefSel.utiltype.value == 11) || (frmDefSel.utiltype.value == 12)) {
			frmOnlyMine.txtTableID.value = frmDefSel.selectTable.options[frmDefSel.selectTable.selectedIndex].value;
			frmDefSel.txtTableID.value = frmOnlyMine.txtTableID.value;
		}

		frmOnlyMine.OnlyMine.value = frmDefSel.checkbox.checked;

		OpenHR.submitForm(frmOnlyMine);
	}

	function setdelete() {

		var frmDefSel = document.getElementById('frmDefSel');
		if (frmDefSel.ssOleDBGridDefSelRecords.selbookmarks.count > 0) {
			var answer = OpenHR.messageBox("Delete this definition. Are you sure ?", 36, "Confirmation");

			if (answer == 6) {
				document.frmDefSel.action.value = "delete";
				OpenHR.submitForm(document.frmDefSel);
			}
		}
		else {
			OpenHR.messageBox("You must select a definition to delete", 48, "OpenHR Intranet");
		}
	}

	function setrun() {

		var frmDefSel = document.getElementById('frmDefSel');
		if (frmDefSel.ssOleDBGridDefSelRecords.selbookmarks.count > 0) {
			frmDefSel.action.value = "run";

			var sUtilId;

			if (frmDefSel.utiltype.value == 25) {
				// Workflow
				var frmWorkflow = document.getElementById('frmWorkflow');
				frmWorkflow.utiltype.value = frmDefSel.utiltype.value;
				frmWorkflow.utilid.value = frmDefSel.utilid.value;
				frmWorkflow.utilname.value = frmDefSel.utilname.value;
				frmWorkflow.action.value = frmDefSel.action.value;
				sUtilId = new String(frmDefSel.utilid.value);

				frmWorkflow.target = sUtilId;
				NewWindow('', sUtilId, '500', '200', 'yes');
				OpenHR.submitForm(frmWorkflow);
			}
			else {

			    //var frmPrompt = document.getElementById('frmPrompt');
		//	    var frmPrompt = OpenHR.getForm("workframe", "frmPrompt");
			    var frmPrompt = document.getElementById('frmPrompt');

				frmPrompt.utilid.value = frmDefSel.utilid.value;
				frmPrompt.utilname.value = frmDefSel.utilname.value;
				frmPrompt.action.value = frmDefSel.action.value;
				sUtilId = new String(frmDefSel.utilid.value);

		//	    frmPrompt.target = sUtilId;
	//		    var newWin = ReturnNewWindow('', sUtilId, '500', '200', 'yes');
			 //   OpenHR.submitForm(frmPrompt, newWin);
			    //OpenHR.submitForm(document.frmPrompt);
				//OpenHR.submitForm(document.frmPrompt);                
				OpenHR.showInReportFrame(document.frmPrompt); 

			}
		//	return false;
		}
		else {
			OpenHR.messageBox("You must select a definition to run", 48, "OpenHR Intranet");
		}
	//	return false;
	}

	function setnew() {

	    OpenHR.showPopup("Loading form. Please wait...");
		document.frmDefSel.action.value = "new";
		OpenHR.submitForm(document.frmDefSel);
	}

	function setcopy() {
		var frmDefSel = document.getElementById('frmDefSel');
		if (frmDefSel.ssOleDBGridDefSelRecords.selbookmarks.count > 0) {
			OpenHR.showPopup("Copying definition. Please wait...");
			document.frmDefSel.action.value = "copy";
			OpenHR.submitForm(document.frmDefSel);
		}
		else {
			OpenHR.messageBox("You must select a definition to copy", 48, "OpenHR Intranet");
		}
	}

	function setedit() {
		var frmDefSel = document.getElementById('frmDefSel');
		if (frmDefSel.ssOleDBGridDefSelRecords.selbookmarks.count > 0) {
			OpenHR.showPopup("Loading definition. Please wait...");

			if (frmDefSel.cmdEdit.value == "Edit") {
				document.frmDefSel.action.value = "edit";
				OpenHR.submitForm(document.frmDefSel);
			}
			else {
				document.frmDefSel.action.value = "view";
				OpenHR.submitForm(document.frmDefSel);
			}
		}
		else {
			OpenHR.messageBox("You must select a definition to edit", 48, "OpenHR Intranet");
		}	    

	}

	function setcancel() {
		var frmDefSel = document.getElementById('frmDefSel');
		if (frmDefSel.txtSingleRecordID.value > 0) {
			var sWorkPage = currentWorkFramePage();
			if (sWorkPage == "RECORDEDIT") {
				refreshData(); //workframe
			}

			window.location.href = "emptyoption";
			menu_disableMenu();
			//TODO
			//window.parent.document.all.item("workframeset").cols = "*, 0";
		}
		else {
			window.location.href = "default";
		}
	}

	/* Sequential search the grid for the required OLE. */
	function locateRecord(psFileName, pfExactMatch) {
		var fFound = false;
		var iIndex;
		
		var frmDefSel = document.getElementById('frmDefSel');

		frmDefSel.ssOleDBGridDefSelRecords.redraw = false;

		frmDefSel.ssOleDBGridDefSelRecords.MoveLast();
		frmDefSel.ssOleDBGridDefSelRecords.MoveFirst();

		for (iIndex = 1; iIndex <= frmDefSel.ssOleDBGridDefSelRecords.rows; iIndex++) {
			if (pfExactMatch == true) {
				if (frmDefSel.ssOleDBGridDefSelRecords.Columns(0).value == psFileName) {
					frmDefSel.ssOleDBGridDefSelRecords.SelBookmarks.Add(frmDefSel.ssOleDBGridDefSelRecords.Bookmark);
					fFound = true;
					break;
				}
			}
			else {
				var sGridValue = new String(frmDefSel.ssOleDBGridDefSelRecords.Columns(0).value);
				sGridValue = sGridValue.substr(0, psFileName.length).toUpperCase();
				if (sGridValue == psFileName.toUpperCase()) {
					frmDefSel.ssOleDBGridDefSelRecords.SelBookmarks.Add(frmDefSel.ssOleDBGridDefSelRecords.Bookmark);
					fFound = true;
					break;
				}
			}

			if (iIndex < frmDefSel.ssOleDBGridDefSelRecords.rows) {
				frmDefSel.ssOleDBGridDefSelRecords.MoveNext();
			}
			else {
				break;
			}
		}

		if ((fFound == false) && (frmDefSel.ssOleDBGridDefSelRecords.rows > 0)) {
			// Select the top row.
			frmDefSel.ssOleDBGridDefSelRecords.MoveFirst();
			frmDefSel.ssOleDBGridDefSelRecords.SelBookmarks.Add(frmDefSel.ssOleDBGridDefSelRecords.Bookmark);
		}

		frmDefSel.ssOleDBGridDefSelRecords.redraw = true;
	}

	//TODO
	function currentWorkFramePage() {
		// Work frame is in view.
		var sCurrentPage = window.parent.frames("workframe").document.location;
		sCurrentPage = sCurrentPage.toString();

		if (sCurrentPage.lastIndexOf("/") > 0) {
			sCurrentPage = sCurrentPage.substr(sCurrentPage.lastIndexOf("/") + 1);
		}

		if (sCurrentPage.indexOf(".") > 0) {
			sCurrentPage = sCurrentPage.substr(0, sCurrentPage.indexOf("."));
		}

		sCurrentPage = sCurrentPage.replace(/ /gi, "");
		sCurrentPage = sCurrentPage.toUpperCase();

		return (sCurrentPage);
	}

</script>

<script type="text/javascript">
<!--
	OpenHR.addActiveXHandler("ssOleDBGridDefSelRecords", "dblClick", ssOleDBGridDefSelRecords_dblClick);
	
	function ssOleDBGridDefSelRecords_dblClick() {

		var frmDefSel = document.getElementById("frmDefSel");

		if ((frmDefSel.utiltype.value == 10) || (frmDefSel.utiltype.value == 11) || (frmDefSel.utiltype.value == 12)) 
		{
			// DblClick triggers Edit.
			setedit();
		}
		else 
		{
			// DblClick triggers Run after prompting for confirmation. 
			if (frmDefSel.cmdRun.disabled == true) 
			{
				return(false);
			}

			var answer = 0;

		    if (frmDefSel.utiltype.value == 1) {
		        answer = OpenHR.messageBox("Are you sure you want to run the '" + frmDefSel.utilname.value + "' Cross Tab ?", 36, "Confirmation...");
		    }

		    if (frmDefSel.utiltype.value == 2) 
		    {
		        answer = OpenHR.messageBox("Are you sure you want to run the '" + frmDefSel.utilname.value + "' Custom Report ?",36,"Confirmation...");
		    }
		    if (frmDefSel.utiltype.value == 9) 
		    {
		        answer = OpenHR.messageBox("Are you sure you want to run the '" + frmDefSel.utilname.value + "' Mail Merge ?",36,"Confirmation...");
		    }
		    if (frmDefSel.utiltype.value == 17) 
		    {
		        answer = OpenHR.messageBox("Are you sure you want to run the '" + frmDefSel.utilname.value + "' Calendar Report ?",36,"Confirmation...");
		    }
			if (frmDefSel.utiltype.value == 25) 
			{
				answer = OpenHR.messageBox("Are you sure you want to run the '" + frmDefSel.utilname.value + "' Workflow ?",36,"Confirmation...");
			}
			
			if (answer == 6) 
			{
				setrun();
			}
		}
		return false;
	}
-->
</script>

<script type="text/javascript">
<!--
	OpenHR.addActiveXHandler("ssOleDBGridDefSelRecords", "KeyPress", ssOleDBGridDefSelRecords_KeyPress);
	
	function ssOleDBGridDefSelRecords_KeyPress(iKeyAscii) {

		var txtTicker = document.getElementById("txtTicker");
		var txtLastKeyFind = document.getElementById("txtLastKeyFind");
		
		if ((iKeyAscii >= 32) && (iKeyAscii <= 255)) {
			var dtTicker = new Date();
			var iThisTick = new Number(dtTicker.getTime());
			var sFind;
			var iLastTick;
			if (txtLastKeyFind.value.length > 0) {
				iLastTick = new Number(txtTicker.value);
			} else {
				iLastTick = new Number("0");
			}

			if (iThisTick > (iLastTick + 1500)) {
				sFind = String.fromCharCode(iKeyAscii);
			} else {
				sFind = txtLastKeyFind.value + String.fromCharCode(iKeyAscii);
			}

			txtTicker.value = iThisTick;
			txtLastKeyFind.value = sFind;

			locateRecord(sFind, false);
		}
	}
-->
</script>

<script type="text/javascript">
<!--
	OpenHR.addActiveXHandler("ssOleDBGridDefSelRecords", "rowcolchange", ssOleDBGridDefSelRecords_rowcolchange);
	
	function ssOleDBGridDefSelRecords_rowcolchange() {

		var frmDefSel = document.getElementById("frmDefSel");
		var frmpermissions = document.getElementById("frmpermissions");

		// Populate the textbox with the definitions description
		frmDefSel.txtDescription.value = frmDefSel.ssOleDBGridDefSelRecords.Columns("description").Value;

		// Populate the hidden fields with the selected utils information
		frmDefSel.utilid.value = frmDefSel.ssOleDBGridDefSelRecords.Columns("id").Value;
		frmDefSel.utilname.value = frmDefSel.ssOleDBGridDefSelRecords.Columns("name").Value;

		// Check for RO access and set EDIT/VIEW caption as appropriate
		var username = frmDefSel.ssOleDBGridDefSelRecords.Columns("username").Value;
		var access = frmDefSel.ssOleDBGridDefSelRecords.Columns("access").Value;

		button_disable(frmDefSel.cmdRun, (frmpermissions.grantrun.value == 0));
		button_disable(frmDefSel.cmdNew, (frmpermissions.grantnew.value == 0));
		button_disable(frmDefSel.cmdCopy, (frmpermissions.grantnew.value == 0));
		button_disable(frmDefSel.cmdEdit, (frmpermissions.grantedit.value == 0));

		if (username != frmDefSel.txtusername.value) {
			if (access == 'ro') {
				frmDefSel.cmdEdit.value = 'View';
				button_disable(frmDefSel.cmdDelete, true);
			} else {
				frmDefSel.cmdEdit.value = 'Edit';

				if (frmpermissions.grantdelete.value == 1) {
					button_disable(frmDefSel.cmdDelete, false);
				} else {
					button_disable(frmDefSel.cmdDelete, true);
				}
			}
		} else {
			frmDefSel.cmdEdit.value = 'Edit';

			if (frmpermissions.grantdelete.value == 1) {
				button_disable(frmDefSel.cmdDelete, false);
			} else {
				button_disable(frmDefSel.cmdDelete, true);
			}
		}

		refreshControls();
	}
-->
</script>

<form id=frmpermissions name=frmpermissions style="visibility:hidden;display:none">
<%
	Dim cmdDefSelAccess = CreateObject("ADODB.Command")
	cmdDefSelAccess.CommandText = "sp_ASRIntGetSystemPermissions"
	cmdDefSelAccess.CommandType = 4 ' Stored Procedure
	cmdDefSelAccess.ActiveConnection = Session("databaseConnection")

	Err.Clear()
	Dim rstDefSelAccess = cmdDefSelAccess.Execute
	
	Dim fNewGranted = 0
	Dim fEditGranted = 0
	Dim fDeleteGranted = 0
	Dim fRunGranted = 0
	Dim fViewGranted = 0

  'MH20011008 We should probably change this so that we pass
  'over the category key and get back a smaller recordset

	do until rstdefselaccess.eof
		if session("defseltype") = 1 then
			if rstdefselaccess.fields(0).value = "CROSSTABS_NEW" then
				fNewGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CROSSTABS_EDIT" then
				fEditGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CROSSTABS_DELETE" then
				fDeleteGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CROSSTABS_RUN" then
				fRunGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CROSSTABS_VIEW" then
				fViewGranted = CInt(rstDefSelAccess.fields(1).value)
			end if
		end if

		if session("defseltype") = 2 then
			if rstdefselaccess.fields(0).value = "CUSTOMREPORTS_NEW" then
				fNewGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CUSTOMREPORTS_EDIT" then
				fEditGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CUSTOMREPORTS_DELETE" then
				fDeleteGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CUSTOMREPORTS_RUN" then
				fRunGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CUSTOMREPORTS_VIEW" then
				fViewGranted = CInt(rstDefSelAccess.fields(1).value)
			end if
		end if

		if session("defseltype") = 9 then
			if rstdefselaccess.fields(0).value = "MAILMERGE_NEW" then
				fNewGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "MAILMERGE_EDIT" then
				fEditGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "MAILMERGE_DELETE" then
				fDeleteGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "MAILMERGE_RUN" then
				fRunGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "MAILMERGE_VIEW" then
				fViewGranted = CInt(rstDefSelAccess.fields(1).value)
			end if
		end if

		if session("defseltype") = 10 then
			if rstdefselaccess.fields(0).value = "PICKLISTS_NEW" then
				fNewGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "PICKLISTS_EDIT" then
				fEditGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "PICKLISTS_DELETE" then
				fDeleteGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "PICKLISTS_VIEW" then
				fViewGranted = CInt(rstDefSelAccess.fields(1).value)
			end if
		end if

		if session("defseltype") = 11 then
			if rstdefselaccess.fields(0).value = "FILTERS_NEW" then
				fNewGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "FILTERS_EDIT" then
				fEditGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "FILTERS_DELETE" then
				fDeleteGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "FILTERS_VIEW" then
				fViewGranted = CInt(rstDefSelAccess.fields(1).value)
			end if
		end if
  
		if session("defseltype") = 12 then
			if rstdefselaccess.fields(0).value = "CALCULATIONS_NEW" then
				fNewGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CALCULATIONS_EDIT" then
				fEditGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CALCULATIONS_DELETE" then
				fDeleteGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CALCULATIONS_VIEW" then
				fViewGranted = CInt(rstDefSelAccess.fields(1).value)
			end if
		end if
		
		if session("defseltype") = 17 then
			if rstdefselaccess.fields(0).value = "CALENDARREPORTS_NEW" then
				fNewGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CALENDARREPORTS_EDIT" then
				fEditGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CALENDARREPORTS_DELETE" then
				fDeleteGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CALENDARREPORTS_RUN" then
				fRunGranted = CInt(rstDefSelAccess.fields(1).value)
			elseif rstdefselaccess.fields(0).value = "CALENDARREPORTS_VIEW" then
				fViewGranted = CInt(rstDefSelAccess.fields(1).value)
			end if
		end if
		
		if session("defseltype") = 25 then
			fNewGranted = 0
			fEditGranted = 0
			fDeleteGranted = 0
			fViewGranted = 0

			if rstdefselaccess.fields(0).value = "WORKFLOW_RUN" then
				fRunGranted = CInt(rstDefSelAccess.fields(1).value)
			end if
		end if

		rstdefselaccess.movenext
	loop
	
	rstDefSelAccess = Nothing
	cmdDefSelAccess = Nothing
	
	Response.Write("<INPUT type=hidden id=grantnew name=grantnew value = " & fNewGranted & ">" & vbCrLf)
	Response.Write("<INPUT type=hidden id=grantedit name=grantedit value = " & fEditGranted & ">" & vbCrLf)
	Response.Write("<INPUT type=hidden id=grantdelete name=grantdelete value = " & fDeleteGranted & ">" & vbCrLf)
	Response.Write("<INPUT type=hidden id=grantrun name=grantrun value = " & fRunGranted & ">" & vbCrLf)
	Response.Write("<INPUT type=hidden id=grantview name=grantview value = " & fViewGranted & ">" & vbCrLf)
%>
</form>

<div <%=session("BodyTag")%>>

<form name="frmDefSel" action="defsel_submit" method="post" id="frmDefSel">

<table class="outline" height="100%" width=100%>
	<TR>
		<TD>
    	    <table width="100%" height="100%" class="invisible">
	            <tr> 
		            <td colspan=5 height=10>
						<H3 class="pageTitle">
<%
	if session("defseltype") = 0 then	        'BATCH JOB
		Response.write("Batch Jobs")
	elseif session("defseltype") = 1 then	    'CROSS TAB
		Response.write("Cross Tabs")
	elseif session("defseltype") = 2 then	    'CUSTOM REPORTS
		Response.write("Custom Reports")
	elseif session("defseltype") = 3 then	    'DATA TRANSFER
		Response.write("Data Transfer")
	elseif session("defseltype") = 4 then   	'EXPORT
		Response.write("Export")
	elseif session("defseltype") = 5 then   	'GLOBAL ADD
		Response.write("Global Add")
	elseif session("defseltype") = 6 then	    'GLOBAL UPDATE
		Response.write("Global Update")
	elseif session("defseltype") = 7 then   	'GLOBAL DELETE
		Response.write("Global Delete")
	elseif session("defseltype") = 8 then	    'IMPORT
		Response.write("Import")
	elseif session("defseltype") = 9 then		'MAIL MERGE
		Response.write("Mail Merge")
	elseif session("defseltype") = 10 then		'PICKLIST
		Response.write("Picklists")
	elseif session("defseltype") = 11 then		'FILTERS
		Response.write("Filters")
	elseif session("defseltype") = 12 then		'CALCULATIONS
		Response.write("Calculations")
	elseif session("defseltype") = 17 then		'CALENDAR REPORTS
		Response.write("Calendar Reports")
	elseif session("defseltype") = 25 then		'WORKFLOW
		Response.write("Workflow")
	end if
%>
						</H3>
					</td>
				</tr>

<% 
	Dim sErrorDescription = ""
	
    if session("defseltype") = 10 or session("defseltype") = 11 or session("defseltype") = 12 then 
%>
				<tr height=10>
					<td width=20>&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td height="10" colspan=3>
						<table width=100% class="invisible">
							<TR>
								<TD width=40>
									Table :
								</TD>
								<TD width=10>
									&nbsp;
								</TD>
								<TD width=175>
									<SELECT id=selectTable name=selectTable class="combo" style="HEIGHT: 22px; WIDTH: 200px">
<%
	    on error resume next
	
	If (Len(sErrorDescription) = 0) Then
		' Get the view records.
		Dim cmdTableRecords = CreateObject("ADODB.Command")
		cmdTableRecords.CommandText = "sp_ASRIntGetTables"
		cmdTableRecords.CommandType = 4 ' Stored Procedure
		cmdTableRecords.ActiveConnection = Session("databaseConnection")

		Err.Clear()
		Dim rstTableRecords = cmdTableRecords.Execute

		If (Err.Number <> 0) Then
			sErrorDescription = "The table records could not be retrieved." & vbCrLf & formatError(Err.Description)
		End If

		If (Len(sErrorDescription) = 0) Then
			Do While Not rstTableRecords.EOF
				Response.Write("						<OPTION value=" & rstTableRecords.Fields(0).Value)
				If rstTableRecords.Fields(0).Value = CLng(Session("utilTableID")) Then
					Response.Write(" SELECTED")
				End If

				Response.Write(">" & Replace(CStr(rstTableRecords.Fields(1).Value), "_", " ") & "</OPTION>" & vbCrLf)

				rstTableRecords.MoveNext()
			Loop
    			
			' Release the ADO recordset object.
			rstTableRecords.close()
			rstTableRecords = Nothing
		End If

		' Release the ADO command object.
		cmdTableRecords = Nothing
	End If
%>
									</SELECT>						
								</TD>
								<TD width=10>
									<INPUT type="button" value="Go" id=btnGoTable name=btnGoTable class="btn" onclick="ToggleCheck();" />
								</TD>
								<TD>&nbsp;
								</TD>
							</TR>
						</table>
					</td>
					<td width=20>&nbsp;&nbsp;&nbsp;&nbsp;</td>
				</tr>
		        <tr> 
				    <td colspan=5 height=10></td>
				</tr>
<%
    end if
%>
				
				<tr> 
					<td width=20>&nbsp;&nbsp;&nbsp;&nbsp;</td>
					<td width=100%>
						<table height=100% width=100%>
							<tr>
								<td width=100%>
<%
    if len(sErrorDescription) = 0 then
        ' Get the records.
		Dim cmdDefSelRecords = CreateObject("ADODB.Command")
        cmdDefSelRecords.CommandText = "sp_ASRIntPopulateDefSel"
        cmdDefSelRecords.CommandType = 4 ' Stored Procedure
		cmdDefSelRecords.ActiveConnection = Session("databaseConnection")

		Dim prmType = cmdDefSelRecords.CreateParameter("type", 3, 1)
		cmdDefSelRecords.Parameters.Append(prmType)
        prmType.value = cleanNumeric(session("defseltype"))

		Dim prmOnlyMine = cmdDefSelRecords.CreateParameter("onlymine", 11, 1)
		cmdDefSelRecords.Parameters.Append(prmOnlyMine)
        prmOnlyMine.value = cleanBoolean(session("OnlyMine")) ' 0 '1

		Dim prmTableId = cmdDefSelRecords.CreateParameter("tableID", 3, 1)
		cmdDefSelRecords.Parameters.Append(prmTableId)
		prmTableId.value = cleanNumeric(Session("utilTableID"))

		Err.Clear()
		Dim rstDefSelRecords = cmdDefSelRecords.Execute

		If (Err.Number <> 0) Then
			sErrorDescription = "The Defsel records could not be retrieved." & vbCrLf & formatError(Err.Description)
		End If

        if len(sErrorDescription) = 0 then
	        ' Instantiate and initialise the grid. 
			Response.Write("								<OBJECT classid=""clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"" id=ssOleDBGridDefSelRecords name=ssOleDBGridDefselRecords codebase=""cabs/COAInt_Grid.cab#version=3,1,3,6"" style=""LEFT: 0px; TOP: 0px; WIDTH:100%; HEIGHT:400px"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""ScrollBars"" VALUE=""4"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""_Version"" VALUE=""196616"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""DataMode"" VALUE=""2"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""Cols"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""Rows"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""BorderStyle"" VALUE=""1"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""RecordSelectors"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""GroupHeaders"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""ColumnHeaders"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""GroupHeadLines"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""HeadLines"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""FieldDelimiter"" VALUE=""(None)"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""FieldSeparator"" VALUE=""(Tab)"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""Col.Count"" VALUE=""" & rstDefSelRecords.fields.count & """>" & vbCrLf)
			Response.Write("									<PARAM NAME=""stylesets.count"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""TagVariant"" VALUE=""EMPTY"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""UseGroups"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""HeadFont3D"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""Font3D"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""DividerType"" VALUE=""3"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""DividerStyle"" VALUE=""1"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""DefColWidth"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""BeveColorScheme"" VALUE=""2"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""BevelColorFrame"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""BevelColorHighlight"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""BevelColorShadow"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""BevelColorFace"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""CheckBox3D"" VALUE=""-1"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""AllowAddNew"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""AllowDelete"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""AllowUpdate"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""MultiLine"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""ActiveCellStyleSet"" VALUE="""">" & vbCrLf)
			Response.Write("									<PARAM NAME=""RowSelectionStyle"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""AllowRowSizing"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""AllowGroupSizing"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""AllowColumnSizing"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""AllowGroupMoving"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""AllowColumnMoving"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""AllowGroupSwapping"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""AllowColumnSwapping"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""AllowGroupShrinking"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""AllowColumnShrinking"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""AllowDragDrop"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""UseExactRowCount"" VALUE=""-1"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""SelectTypeCol"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""SelectTypeRow"" VALUE=""1"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""SelectByCell"" VALUE=""-1"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""BalloonHelp"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""RowNavigation"" VALUE=""1"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""CellNavigation"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""MaxSelectedRows"" VALUE=""1"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""HeadStyleSet"" VALUE="""">" & vbCrLf)
			Response.Write("									<PARAM NAME=""StyleSet"" VALUE="""">" & vbCrLf)
			Response.Write("									<PARAM NAME=""ForeColorEven"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""ForeColorOdd"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""BackColorEven"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""BackColorOdd"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""Levels"" VALUE=""1"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""RowHeight"" VALUE=""503"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""ExtraHeight"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""ActiveRowStyleSet"" VALUE="""">" & vbCrLf)
			Response.Write("									<PARAM NAME=""CaptionAlignment"" VALUE=""2"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""SplitterPos"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""SplitterVisible"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""Columns.Count"" VALUE=""" & rstDefSelRecords.fields.count & """>" & vbCrLf)

	        for iLoop = 0 to (rstDefSelRecords.fields.count - 1)

		        if rstDefSelRecords.fields(iLoop).name <> "Name" then
					Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").Width"" VALUE=""0"">" & vbCrLf)
					Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").Visible"" VALUE=""0"">" & vbCrLf)
		        else
					Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").Width"" VALUE=""100000"">" & vbCrLf)
					Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").Visible"" VALUE=""-1"">" & vbCrLf)
		        end if
        			
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").Columns.Count"" VALUE=""1"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").Caption"" VALUE=""" & Replace(CStr(rstDefSelRecords.fields(iLoop).name), "_", " ") & """>" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").Name"" VALUE=""" & rstDefSelRecords.fields(iLoop).name & """>" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").Alignment"" VALUE=""0"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").CaptionAlignment"" VALUE=""3"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").Bound"" VALUE=""0"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").AllowSizing"" VALUE=""1"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").DataField"" VALUE=""Column " & iLoop & """>" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").DataType"" VALUE=""8"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").Level"" VALUE=""0"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").NumberFormat"" VALUE="""">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").Case"" VALUE=""0"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").FieldLen"" VALUE=""4096"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").VertScrollBar"" VALUE=""0"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").Locked"" VALUE=""0"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").Style"" VALUE=""0"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").ButtonsAlways"" VALUE=""0"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").RowCount"" VALUE=""0"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").ColCount"" VALUE=""1"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").HasHeadForeColor"" VALUE=""0"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").HasHeadBackColor"" VALUE=""0"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").HasForeColor"" VALUE=""0"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").HasBackColor"" VALUE=""0"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").HeadForeColor"" VALUE=""0"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").HeadBackColor"" VALUE=""0"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").ForeColor"" VALUE=""0"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").BackColor"" VALUE=""0"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").HeadStyleSet"" VALUE="""">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").StyleSet"" VALUE="""">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").Nullable"" VALUE=""1"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").Mask"" VALUE="""">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").PromptInclude"" VALUE=""0"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").ClipMode"" VALUE=""0"">" & vbCrLf)
				Response.Write("									<PARAM NAME=""Columns(" & iLoop & ").PromptChar"" VALUE=""95"">" & vbCrLf)
	        next 

			Response.Write("									<PARAM NAME=""UseDefaults"" VALUE=""-1"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""TabNavigation"" VALUE=""1"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""_ExtentX"" VALUE=""17330"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""_ExtentY"" VALUE=""1323"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""_StockProps"" VALUE=""79"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""Caption"" VALUE="""">" & vbCrLf)
			Response.Write("									<PARAM NAME=""ForeColor"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""BackColor"" VALUE=""0"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""Enabled"" VALUE=""-1"">" & vbCrLf)
			Response.Write("									<PARAM NAME=""DataMember"" VALUE="""">" & vbCrLf)
			Response.Write("									<PARAM NAME=""Row.Count"" VALUE=""0"">" & vbCrLf)
			Response.Write("								</OBJECT>" & vbCrLf)
        			
			Dim lngRowCount = 0
	        do while not rstDefSelRecords.EOF
				Dim sAddString = ""

		        for iLoop = 0 to (rstDefSelRecords.fields.count - 1)							
					sAddString = sAddString & Replace(Replace(CStr(rstDefSelRecords.Fields(iLoop).Value), "_", " "), Chr(34), "&quot;") & "	"
		        next 				

				Response.Write("<INPUT type='hidden' id=txtAddString_" & lngRowCount & " name=txtAddString_" & lngRowCount & " value=""" & sAddString & """>" & vbCrLf)

		        lngRowCount = lngRowCount + 1
		        rstDefSelRecords.MoveNext
	        loop

	        ' Release the ADO recordset object.
	        rstDefSelRecords.close
			rstDefSelRecords = Nothing
        end if
        			
        ' Release the ADO command object.
		cmdDefSelRecords = Nothing
        end if
		%>
								</td>
							</tr>
							
							<tr height=10>
								<td></td>
							</tr>
							
							<tr>
								<td height="70">
									<textarea cols="20" class="disabled" style="WIDTH: 100%;" name="txtDescription" rows="4" readonly="readonly" tabindex="-1">
									</textarea>
								</td>
							</tr>
						</table>							
					</td>
					
					<td width=20>&nbsp;&nbsp;&nbsp;&nbsp;</td>
	        
	                <td width=80 style="display: none;"> 
						<table height=100% class="invisible">
							<tr>
								<td>
        					        <input type="button" id="cmdNew" class="btn" name="cmdNew" value="New" style="WIDTH: 80px" width="80"
<% 
if (session("singleRecordID") > 0) or session("defseltype") = 25 then 
		Response.Write(" style=""visibility:hidden""")
end if
%>							  
									    onclick="setnew();" />
							    </td>
							</tr>
							<tr height=10>
								<td></td>
							</tr>
							<tr>
								<td>
								    <input type="button" name="cmdEdit" class="btn" value="Edit" style="WIDTH: 80px" width="80"
<% 
if (session("singleRecordID") > 0) or session("defseltype") = 25 then 
		Response.Write(" style=""visibility:hidden""")
end if
%>							  
                                        onclick="setedit();" />
							  </td>
							</tr>
							<tr height=10>
								<td></td>
							</tr>
							<tr>
								<td>
								    <input type="button" name="cmdCopy" class="btn" id="cmdCopy" value="Copy" style="WIDTH: 80px" width="80"
<% 
if (session("singleRecordID") > 0) or session("defseltype") = 25 then 
		Response.Write(" style=""visibility:hidden""")
end if
%>							  
									    onclick="setcopy();" />
							    </td>
							</tr>
							<tr height=10>
								<td></td>
							</tr>
							<tr>
								<td>
								    <input type="button" name="cmdDelete" class="btn" value="Delete" style="WIDTH: 80px" width="80"
<% 
if (session("singleRecordID") > 0) or session("defseltype") = 25 then 
		Response.Write(" style=""visibility:hidden""")
end if
%>							  
							            onclick="setdelete();" />
							  </td>
							</tr>
							<tr height=10>
								<td></td>
							</tr>
							<tr>
								<td>
								    <input type="button" name="cmdPrint" class="btn btndisabled" value="Print" style="WIDTH: 80px" width="80" disabled
<% 
if (session("singleRecordID") > 0) or session("defseltype") = 25 then 
		Response.Write(" style=""visibility:hidden""")
end if
%>							  />
							  </td>
							</tr>
							<tr height=10>
								<td></td>
							</tr>
							<tr>
								<td>
									<input type="button" name="cmdProperties" class="btn" value="Properties" style="WIDTH: 80px" width="80" 
										<% 
if (session("singleRecordID") > 0) or session("defseltype") = 25 then 
	Response.Write(" style=""visibility:hidden""")
											end if
%>							  
									       onclick="showproperties();" />
							  </td>
							</tr>
							<tr height=10>
								<td></td>
							</tr>
							<tr height=100%>
								<td></td>
							</tr>
							<tr>
								<td>
									<input type="button" name=cmdRun class="btn" value="Run" style="WIDTH: 80px" width="80" id=cmdRun
<% 
	if session("defseltype") = 10 or session("defseltype") = 11 or session("defseltype") = 12 then 
			Response.Write(" style=""visibility:hidden""")
	end if
%>
									    onclick="setrun();" />
							    </td>
							</tr>
							<tr height=10>
								<td></td>
							</tr>
							<tr>
								<td>
									<input type="button" name="cmdCancel" class="btn" value=
<% 
	if session("defseltype") = 10 or session("defseltype") = 11 or session("defseltype") = 12 then
		Response.Write("""OK""")
	else
		Response.Write("""Cancel""")
	end if
%>
										style="WIDTH: 80px" width="80"
										onclick="setcancel()" />
							    </td>
							</tr>
						</table>	
					</td>
					<td width=20>&nbsp;&nbsp;&nbsp;&nbsp;</td>
	            </tr>
  
		        <tr> 
				    <td colspan=5 height=10
<%
if session("defseltype") = 25 then 
		Response.Write(" style=""visibility:hidden""")
end if
%>				    >
						<INPUT type='hidden' id=txtusername name=txtusername value="<%=lcase(session("Username"))%>">
				    </td>
				</tr>

				<tr> 
					<td width=20>&nbsp;</td>
				    <td colspan=4 height="10"
<%
if session("defseltype") = 25 then 
		Response.Write(" style=""visibility:hidden""")
end if
%>
				    > 
				        <input <% if session("OnlyMine") then Response.Write("checked") %> type="checkbox" tabindex="-1" id="checkbox" name="checkbox" value="checkbox" 
				            onclick="ToggleCheck();" />
                        <label for="checkbox" class="checkbox" tabindex=0 onkeypress="try{checkboxLabel_onKeyPress(this);}catch(e){}">
				            Only show definitions where owner is '<%=Session("Username")%>'
	    		        </label>
					</td>
				</tr>
	        </table>
		</td>
	</tr>
</table>
<input type="hidden" id="utiltype" name="utiltype" value="<%=Session("defseltype")%>">
<input type="hidden" id="utilid" name="utilid"  value=<%=Session("utilid")%>>
<input type="hidden" id="utilname" name="utilname">
<input type="hidden" id="action" name="action">
<input type="hidden" id="txtTableID" name="txtTableID" value=<%=session("utilTableID")%>>
<input type="hidden" id="txtSingleRecordID" name="txtSingleRecordID" value=<%=session("singleRecordID")%>>
</form>
    
<form name=frmPrompt method=post action=util_run_promptedValues id=frmPrompt style="visibility:hidden;display:none">
	<input type="hidden" id="utiltype" name="utiltype" value="<%=Session("defseltype")%>">
	<input type="hidden" id="utilid" name="utilid"  value=<%=Session("utilid")%>>
	<input type="hidden" id="utilname" name="utilname" >
	<input type="hidden" id="action" name="action" >
</form>

<form name=frmWorkflow method=post action=util_run_workflow id=frmWorkflow style="visibility:hidden;display:none">
	<input type="hidden" id="utiltype" name="utiltype">
	<input type="hidden" id="utilid" name="utilid">
	<input type="hidden" id="utilname" name="utilname" >
	<input type="hidden" id="action" name="action" >
</form>

<form action="defsel" method=post id=frmOnlyMine name=frmOnlyMine style="visibility:hidden;display:none">
	<INPUT type="hidden" id=OnlyMine name=OnlyMine value=<%=Session("OnlyMine")%>>
	<INPUT type="hidden" id=txtTableID name=txtTableID value=<%=Session("utilTableID")%>>
</form>

<form target="properties" action="defselproperties" method=post id=frmProp name=frmProp style="visibility:hidden;display:none">
	<INPUT type="hidden" id=prop_name name=prop_name>
	<INPUT type="hidden" id=prop_id name=prop_id>
	<INPUT type="hidden" id=utiltype name=utiltype>
</form>

<INPUT type='hidden' id=txtTicker name=txtTicker value=0>
<INPUT type='hidden' id=txtLastKeyFind name=txtLastKeyFind value="">

<form action="default_Submit" method=post id=frmGoto name=frmGoto style="visibility:hidden;display:none">
	<%Html.RenderPartial("~/Views/Shared/gotoWork.ascx")%>
</form>

</div>

<script type="text/javascript">
	defsel_window_onload();
</script>
