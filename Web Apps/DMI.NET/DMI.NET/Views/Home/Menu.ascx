﻿<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="System.Collections.ObjectModel" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="HR.Intranet.Server.Structures" %>

<%
	On Error Resume Next

	Dim sErrorDescription As String
	Dim avPrimaryMenuInfo As List(Of MenuInfo)
	Dim avQuickEntryMenuInfo As List(Of MenuInfo)
	Dim avTableMenuInfo As List(Of TableScreen)
	Dim avHistoryMenuInfo As List(Of HR.Intranet.Server.Structures.HistoryScreen)
	Dim iLoop As Integer
	Dim objMenu As HR.Intranet.Server.Menu
	Dim sToolCaption As String
	Dim sToolID As String
	
	Dim objSession As SessionInfo = CType(Session("sessionContext"), SessionInfo)

	sErrorDescription = ""
	
	objMenu = New HR.Intranet.Server.Menu()
	objMenu.SessionInfo = objSession
		
	Response.Write(vbCrLf & "<script type=""text/javascript"">" & vbCrLf)

	' ------------------------------------------------------------------------------
	' Create the sub-routine to populate the database menu with the tables available
	' to the current user.
	' ------------------------------------------------------------------------------

	
	Response.Write("function refreshDatabaseMenu() {" & vbCrLf)
	Response.Write("  var objFileTool;" & vbCrLf)
	Response.Write("  var sLastToolName;" & vbCrLf)
	Response.Write("  var lngLastScreenID;" & vbCrLf & vbCrLf)
	Response.Write("  var frmMenuInfo = document.getElementById('frmMenuInfo');" & vbCrLf)
	Response.Write("  if (frmMenuInfo.txtDoneDatabaseMenu.value == 1) {" & vbCrLf)
	Response.Write("    return;" & vbCrLf)
	Response.Write("  }" & vbCrLf & vbCrLf)
	Response.Write("  frmMenuInfo.txtDoneDatabaseMenu.value = 1;" & vbCrLf & vbCrLf)
	
	If Session("avPrimaryMenuInfo") Is Nothing Then
		avPrimaryMenuInfo = objMenu.GetPrimaryTableMenu
		Session("avPrimaryMenuInfo") = avPrimaryMenuInfo
	Else
		avPrimaryMenuInfo = Session("avPrimaryMenuInfo")
	End If
	
	For Each objMenuItem In avPrimaryMenuInfo
		
		If objMenuItem.TableScreenID > 0 Then
			' The user has 'read' permission on the table, and no views on the table.
			' There is only one screen defined for the table.
				
			' Add a menu option to call up the primary table screen.
			' new method to insert a new menu item.
			Response.Write("  menu_insertMenuItem('mnubandDatabase', '" & CleanStringForJavaScript(Replace(objMenuItem.TableName, "_", " ")) & "..." & "', 'PT_" & objMenuItem.TableID & "_0_" & objMenuItem.TableScreenID & "');" & vbCrLf)
		ElseIf objMenuItem.ViewID > 0 Then
			' The user does NOT have 'read' permission on the table, but does have
			' 'read' permission on one view of the table.
			' There is only one screen defined for the view.
			' new method to insert a new menu item.
			Response.Write("  menu_insertMenuItem('mnubandDatabase', '" & CleanStringForJavaScript(Replace(objMenuItem.TableName, "_", " ")) & " (" & CleanStringForJavaScript(Replace(objMenuItem.ViewName, "_", " ")) & " view)..." & "', 'PV_" & objMenuItem.TableID & "_" & objMenuItem.ViewID & "_" & objMenuItem.ViewScreenID & "');" & vbCrLf)
		ElseIf (objMenuItem.ViewScreenCount > 0) Or ((objMenuItem.TableReadable = True) And (objMenuItem.TableScreenCount > 0)) Then
			' The user has 'read' permission on the table, and the table has more than one screen defined for it.
			' Or there are views on the table.
			'Instantiate the submenu heading tool and set properties

			' new method to insert a new submenu item.
			Response.Write("  menu_insertSubMenuItem('mnubandDatabase', '" & CleanStringForJavaScript(Replace(objMenuItem.TableName, "_", " ")) & "', 'PS_" & objMenuItem.TableID & "', 'mnusubband_" & CleanStringForJavaScript(objMenuItem.TableName) & "');" & vbCrLf & vbCrLf)
			
			' Add the submenu.			
			Response.Write("  lngLastScreenID = 0;" & vbCrLf)
			Response.Write("  sLastToolName = """";" & vbCrLf)
			
			For Each objSubMenu In objMenuItem.SubItems
						
				If objSubMenu.ViewID > 0 Then
					sToolID = "PV_" & objMenuItem.TableID & "_" & objSubMenu.ViewID & "_" & objSubMenu.TableScreenID
					sToolCaption = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & CleanStringForJavaScript(Replace(objSubMenu.ViewName, "_", " ")) & " view..."
				Else
					sToolID = "PT_" & objMenuItem.TableID & "_0_" & objSubMenu.TableScreenID
					sToolCaption = CleanStringForJavaScript(Replace(objSubMenu.ScreenName, "_", " ")) & "..."
				End If

				Response.Write("  lngLastScreenID = " & objSubMenu.TableScreenID & ";" & vbCrLf)
				Response.Write("	sLastToolName = '" & sToolID & "'" & vbCrLf)

				' new method to insert a new menu item.
				Response.Write("  menu_insertMenuItem('mnusubband_" & CleanStringForJavaScript(objMenuItem.TableName) & "', '" & sToolCaption & "', '" & sToolID & "','someClass');" & vbCrLf & vbCrLf)
			Next
		End If
	Next
	
	Response.Write("}" & vbCrLf & vbCrLf)
	
	
	' ------------------------------------------------------------------------------
	' Create the sub-routine to populate the quick entry menu.
	' ------------------------------------------------------------------------------
	Response.Write("function refreshQuickEntryMenu() {" & vbCrLf)
	Response.Write("  var objFileTool;" & vbCrLf)
	Response.Write("  var lngQuickEntryCount;" & vbCrLf & vbCrLf)
	Response.Write("  var frmMenuInfo = document.getElementById('frmMenuInfo');" & vbCrLf)
	Response.Write("  frmMenuInfo.txtDoneQuickEntryMenu.value = 1;" & vbCrLf & vbCrLf)
	
	If Session("avQuickEntryMenuInfo") Is Nothing Then
		avQuickEntryMenuInfo = objMenu.GetQuickEntryScreens
		Session("avQuickEntryMenuInfo") = avQuickEntryMenuInfo
	Else
		avQuickEntryMenuInfo = Session("avQuickEntryMenuInfo")
	End If
	
	For Each objMenuItem In avQuickEntryMenuInfo
		Response.Write("  menu_insertMenuItem('mnubandQuickEntry', '" & CleanStringForJavaScript(Replace(objMenuItem.TableName, "_", " ")) & "..." & "', 'QE_" & CleanStringForJavaScript(objMenuItem.TableID) & "_0_" & CleanStringForJavaScript(objMenuItem.TableScreenID) & "');" & vbCrLf)		
	Next
	
	Response.Write("  if (lngQuickEntryCount == 0) {" & vbCrLf)
	Response.Write("  }" & vbCrLf)

	' Sort the items.
	Response.Write("	menu_sortULMenuItems('mnubandQuickEntry');")
	
	Response.Write("}" & vbCrLf & vbCrLf)
		
	
	' ------------------------------------------------------------------------------
	' Create the sub-routine to populate the table screens menu.
	' ------------------------------------------------------------------------------
	Response.Write("function refreshTableScreensMenu() {" & vbCrLf)
	Response.Write("  var objFileTool;" & vbCrLf)
	Response.Write("  var lngTableScreensCount;" & vbCrLf & vbCrLf)
	Response.Write("  var frmMenuInfo = document.getElementById('frmMenuInfo');" & vbCrLf)
	Response.Write("  if (frmMenuInfo.txtDoneTableScreensMenu.value == 1) {" & vbCrLf)
	Response.Write("	  return;" & vbCrLf)
	Response.Write("  }" & vbCrLf & vbCrLf)
	Response.Write("  frmMenuInfo.txtDoneTableScreensMenu.value = 1;" & vbCrLf & vbCrLf)
	Response.Write("  lngTableScreensCount = 0;" & vbCrLf)

	If Session("avTableMenuInfo") Is Nothing Then
		avTableMenuInfo = objMenu.GetTableScreens
		Session("avTableMenuInfo") = avTableMenuInfo
	Else
		avTableMenuInfo = Session("avTableMenuInfo")
	End If

	For Each objTableScreen In avTableMenuInfo
		Response.Write("  lngTableScreensCount = lngTableScreensCount + 1;" & vbCrLf)
		Response.Write("  menu_insertMenuItem('mnubandTableScreens', '" & CleanStringForJavaScript(Replace(objTableScreen.TableName, "_", " ")) & "..." & "', 'TS_" & CleanStringForJavaScript(objTableScreen.TableID) & "_0_" & CleanStringForJavaScript(objTableScreen.ScreenID) & "');" & vbCrLf & vbCrLf)
	Next
	
	Response.Write("}" & vbCrLf & vbCrLf)
	
	' ------------------------------------------------------------------------------
	' Create the sub-routine to populate the history screens menu.
	' ------------------------------------------------------------------------------
	Response.Write("function menu_refreshHistoryScreensMenu(pParentScreenID) {" & vbCrLf)

	' Clear out any existing history sub-menus.
	Response.Write("  var objFileTool;" & vbCrLf)
	Response.Write("  var fDone = false;" & vbCrLf)
	Response.Write("  $('[aria-labelledby=""mnutoolHistory""] ul:first').empty();" & vbCrLf)
	
	Dim iLastParentScreenID = 0
	Dim iDoneCount = 0
	Dim iLastChildTableID = 0
	Dim iNextChildTableID As Integer = 0
	Dim sBand As String = ""
		
	If Session("avTableHistoryMenuInfo") Is Nothing Then
		avHistoryMenuInfo = objMenu.GetHistoryScreens
		Session("avTableHistoryMenuInfo") = avHistoryMenuInfo
	Else
		avHistoryMenuInfo = Session("avTableHistoryMenuInfo")
	End If
	
	iLoop = 0
	For Each objHistoryScreen In avHistoryMenuInfo.OrderBy(Function(n) n.parentScreenID)

		If iLastParentScreenID <> objHistoryScreen.parentScreenID Then
			If iDoneCount > 0 Then
				Response.Write("    fDone = true;" & vbCrLf)
				Response.Write("	}" & vbCrLf & vbCrLf)
			End If
			
			iLastChildTableID = 0
			iDoneCount = iDoneCount + 1
			Response.Write("  if (pParentScreenID == " & objHistoryScreen.parentScreenID & ") {" & vbCrLf)
		End If

		' Create the history screen menu item (without placing it in the menu).
		Response.Write("    objFileToolID = ""HT_" & CleanStringForJavaScript(objHistoryScreen.childTableID) & "_0_" & CleanStringForJavaScript(objHistoryScreen.childScreenID) & """;" & vbCrLf)
		Response.Write("    objFileToolCaption = """ & CleanStringForJavaScript(Replace(objHistoryScreen.childScreenName, "_", " ")) & "..."";" & vbCrLf)
		Response.Write("    objFileToolStyle = 0;" & vbCrLf)

		If iLoop < avHistoryMenuInfo.Count() - 1 Then
			If objHistoryScreen.parentScreenID = avHistoryMenuInfo(iLoop + 1).parentScreenID Then
				iNextChildTableID = avHistoryMenuInfo(iLoop + 1).childTableID
			Else
				iNextChildTableID = 0
			End If
		Else
			iNextChildTableID = 0
		End If
		
		If (iLastChildTableID = objHistoryScreen.childTableID) Then
			' The current screen is for the same table as the last screen added to the menu
			' which will have created the sub-menu, so just add it to the sub-menu.
			sBand = "mnuhistorysubband_" & CleanStringForJavaScript(objHistoryScreen.childTableName)
			Response.Write("    menu_insertMenuItem(""" & sBand & """, objFileToolCaption.replace(""&&"", ""&""), objFileToolID);" & vbCrLf & vbCrLf)
			
		Else
			If (iNextChildTableID = objHistoryScreen.childTableID) Then
				' The current screen is for the same table as the next screen to be added
				' but is for a different table to the last screen added to the menu
				' so create a sub-menu, and add this screen to the sub-menu.
				sBand = "mnuhistorysubband_" & CleanStringForJavaScript(objHistoryScreen.childTableName)
				Response.Write("    objBandToolCaption = """ & CleanStringForJavaScript(Replace(objHistoryScreen.childTableName, "_", " ")) & """;" & vbCrLf)
				Response.Write("    objBandToolSubBand = """ & sBand & """;" & vbCrLf)
					
				Response.Write("    menu_insertSubMenuItem(""mnubandHistory"", objBandToolCaption.replace(""&&"", ""&""), ""0"", objBandToolSubBand);" & vbCrLf)
				Response.Write("    menu_insertMenuItem(objBandToolSubBand, objFileToolCaption.replace(""&&"", ""&""), objFileToolID);" & vbCrLf & vbCrLf)
			Else
				' The current screen is for a different table/view to the next and last screens
				' added to the menu so just add this screen to the main menu as normal.
				Response.Write("   menu_insertMenuItem(""mnubandHistory"", objFileToolCaption.replace(""&&"", ""&""), objFileToolID);" & vbCrLf)
			End If
		End If

		iLastParentScreenID = objHistoryScreen.parentScreenID
		iLastChildTableID = objHistoryScreen.childTableID
		iLoop += 1
	Next

	If iDoneCount > 0 Then
		Response.Write("    fDone = true;" & vbCrLf)
		Response.Write("  }" & vbCrLf & vbCrLf)
	End If

	Response.Write("  if (fDone == false) {" & vbCrLf)
	Response.Write("      $('#mnubandHistory').empty();" & vbCrLf & vbCrLf)		' hack!
	Response.Write("	  $('#mnutoolHistory').hide();" & vbCrLf)
	Response.Write("      $('#mnutoolDatabase').click();")
	Response.Write("  }" & vbCrLf)
	Response.Write("  else {" & vbCrLf)
	Response.Write("	    // Disable the history menu for new records" & vbCrLf)
	Response.Write("	    var frmRecEdit = OpenHR.getForm('workframe', 'frmRecordEditForm');" & vbCrLf)
	Response.Write("	    if (frmRecEdit.txtCurrentRecordID.value <= 0) {" & vbCrLf)
	Response.Write("				$('[id^=""HT_""]').each(function () {" & vbCrLf)
	Response.Write("					menu_enableMenuItem($(this).attr(""id""), false);" & vbCrLf)
	Response.Write("				});" & vbCrLf)
	Response.Write("	    };" & vbCrLf)
	Response.Write("      applyJSTree('[aria-labelledby=""mnutoolHistory""]');" & vbCrLf)
	Response.Write("	  $(""#mnutoolHistory"").show();" & vbCrLf)
	Response.Write("      $('#mnutoolHistory').click();")
	Response.Write("  }" & vbCrLf)
	Response.Write("}" & vbCrLf & vbCrLf)
	
	Response.Write("</script>" & vbCrLf)

	objMenu = Nothing

	Dim objUtilities As New HR.Intranet.Server.Utilities
	objUtilities.SessionInfo = CType(Session("SessionContext"), SessionInfo)
	Session("UtilitiesObject") = objUtilities
	
	Dim objOLE As New HR.Intranet.Server.Ole
	objOLE.SessionInfo = CType(Session("SessionContext"), SessionInfo)

	objOLE.TempLocationPhysical = "\\" & Request.ServerVariables("SERVER_NAME") & "\HRProTemp$\"
	Session("OLEObject") = objOLE
	objOLE = Nothing
			
	' ------------------------------------------------------------------------------
	' Check what permissions the user has.
	' ------------------------------------------------------------------------------
	Dim iCustomReportsGranted As Integer = 0
	Dim iCrossTabsGranted As Integer = 0
	Dim iCalendarReportsGranted As Integer = 0
	Dim iMailMergeGranted As Integer = 0
	Dim iWorkflowGranted As Integer = 0
	Dim iCalculationsGranted As Integer = 0
	Dim iFiltersGranted As Integer = 0
	Dim iPicklistsGranted As Integer = 0
	Dim iNewUserGranted As Integer = 0
		
	For Each objPermission In objSession.Permissions

		Response.Write("<input type='hidden' id=txtSysPerm_" & Replace(objPermission.Key, " ", "_") & " name=txtSysPerm_" & Replace(objPermission.Key, " ", "_") & " value=""" & IIf(objPermission.IsPermitted, "1", "0") & """>" & vbCrLf)
		If Left(objPermission.Key, 13) = "CUSTOMREPORTS" And objPermission.IsPermitted Then iCustomReportsGranted = 1
		If Left(objPermission.Key, 9) = "CROSSTABS" And objPermission.IsPermitted Then iCrossTabsGranted = 1
		If Left(objPermission.Key, 15) = "CALENDARREPORTS" And objPermission.IsPermitted Then iCalendarReportsGranted = 1
		If Left(objPermission.Key, 9) = "MAILMERGE" And objPermission.IsPermitted Then iMailMergeGranted = 1
		If objPermission.Key = "WORKFLOW_RUN" And objPermission.IsPermitted Then iWorkflowGranted = 1
		If Left(objPermission.Key, 12) = "CALCULATIONS" And objPermission.IsPermitted Then iCalculationsGranted = 1
		If Left(objPermission.Key, 7) = "FILTERS" And objPermission.IsPermitted Then iFiltersGranted = 1
		If Left(objPermission.Key, 9) = "PICKLISTS" And objPermission.IsPermitted Then iPicklistsGranted = 1
		If (objPermission.Key = "MODULEACCESS_SYSTEMMANAGER" Or objPermission.Key = "MODULEACCESS_SECURITYMANAGER") And objPermission.IsPermitted Then iNewUserGranted = 1
			
	Next
	
	Dim bAbsenceEnabled = objSession.IsModuleEnabled("ABSENCE")

	Response.Write("<input type='hidden' id=txtAbsenceEnabled name=txtAbsenceEnabled value=" & IIf(bAbsenceEnabled, "1", "0") & ">")
	Response.Write("<input type='hidden' id=txtCustomReportsGranted name=txtCustomReportsGranted value=" & iCustomReportsGranted & ">")
	Response.Write("<input type='hidden' id=txtCrossTabsGranted name=txtCrossTabsGranted value=" & iCrossTabsGranted & ">")
	Response.Write("<input type='hidden' id=txtCalendarReportsGranted name=txtCalendarReportsGranted value=" & iCalendarReportsGranted & ">")
	Response.Write("<input type='hidden' id=txtMailMergeGranted name=txtMailMergeGranted value=" & iMailMergeGranted & ">")
	Response.Write("<input type='hidden' id=txtWorkflowGranted name=txtWorkflowGranted value=" & iWorkflowGranted & ">")
	Response.Write("<input type='hidden' id=txtCalculationsGranted name=txtCalculationsGranted value=" & iCalculationsGranted & ">")
	Response.Write("<input type='hidden' id=txtFiltersGranted name=txtFiltersGranted value=" & iFiltersGranted & ">")
	Response.Write("<input type='hidden' id=txtPicklistsGranted name=txtPicklistsGranted value=" & iPicklistsGranted & ">")
	Response.Write("<input type='hidden' id=txtNewUserGranted name=txtNewUserGranted value=" & iNewUserGranted & ">")

	Response.Write("<input type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>")
%>

<div id="contextmenu" class="accordion" style="display: none;">
		<h3 id="mnutoolDatabase">Database</h3>
		<div>
			<ul id="mnubandDatabase">
				<li id="mnutoolQuickEntry"><a href="#">Quick Access Screens</a>
					<ul id="mnubandQuickEntry"></ul>
				</li>
				<li id="mnutoolTableScreens"><a href="#">Lookup Table Screens</a>
					<ul id="mnubandTableScreens"></ul>
				</li>
				<li id="mnutoolLogoff"><a href="#">Log Off...</a></li>
			</ul>
		</div>
		<h3 id="mnutoolRecord" class="hidden">Record</h3>
		<div>
			<ul id="mnubandRecord"></ul>
		</div>
		<h3 id="mnutoolHistory" class="hidden">History</h3>
		<div>
			<ul id="mnubandHistory"></ul>
		</div>
		<h3 id="mnutoolReports">Reports</h3>
		<div>
			<ul id="mnubandReports">
				<li id="mnutoolCustomReports"><a href="#">Custom Reports</a></li>
				<li id="mnutoolCalendarReports"><a href="#">Calendar Reports</a></li>
				<li id="mnutoolCrossTabs"><a href="#">Cross Tabs</a></li>
				<li class="hidden" id="mnutoolStdRpt_AbsenceBreakdown"><a href="#">Absence Breakdown...</a></li>
				<li class="hidden" id="mnutoolStdRpt_BradfordFactor"><a href="#">Bradford Factor...</a></li>
				<li class="hidden" id="mnutoolStdRpt_StabilityReport"><a href="#">Stability Report...</a></li>
				<li class="hidden" id="mnutoolStdRpt_TurnoverReport"><a href="#">Turnover...</a></li>
				<li id="mnutoolOrgChart"><a href="#">Organisation Chart</a></li>
			</ul>
		</div>
		<h3 id="mnutoolUtilities">Utilities</h3>
		<div>
			<ul id="mnubandUtilities">
				<li class="hidden" id="mnutoolBatchJob"><a href="#">Batch Job</a></li>
				<li class="hidden" id="mnutoolDiary"><a href="#">Diary</a></li>
				<li id="mnutoolMailMerge"><a href="#">Mail Merge</a></li>
				<li id="mnutoolWorkflow"><a href="#">Workflow</a></li>
				<li class="hidden" id="mnutoolGlobalAdd"><a href="#">Global Add</a></li>
				<li class="hidden" id="mnutoolGlobalUpdate"><a href="#">Global Update</a></li>
				<li class="hidden" id="mnutoolGlobalDelete"><a href="#">Global Delete</a></li>
				<li class="hidden" id="mnutoolDataTransfer"><a href="#">Data Transfer</a></li>
				<li class="hidden" id="mnutoolImport"><a href="#">Import</a></li>
				<li class="hidden" id="mnutoolExport"><a href="#">Export</a></li>
			</ul>
		</div>
		<h3 id="mnutoolTools">Tools</h3>
		<div>
			<ul id="mnubandTools">
				<li id="mnutoolCalculations"><a href="#">Calculations</a></li>
				<li id="mnutoolFilters"><a href="#">Filters</a></li>
				<li id="mnutoolPicklists"><a href="#">Picklists</a></li>
			</ul>
		</div>
		<h3 id="mnutoolAdministration">Administration</h3>
		<div>
			<ul id="mnubandAdministration">
				<li id="mnutoolEventLog"><a href="#">Event Log</a></li>
				<li id="mnutoolWorkflowPopup"><a href="#">Workflow</a>
					<ul id="mnubandWorkflowPopup">
						<li id="mnutoolWorkflowPendingSteps"><a href="#">Pending Steps...</a></li>
						<li id="mnutoolWorkflowOutOfOffice"><a href="#">Out of Office...</a></li>
					</ul>
				</li>
				<li id="mnutoolPasswordChange"><a href="#">Change Password...</a></li>
				<li id="mnutoolNewUser"><a href="#">New User...</a></li>
				<li id="mnutoolConfiguration"><a href="#">User Configuration...</a></li>
				<li id="mnutoolPCConfiguration"><a href="#">PC Configuration...</a></li>
			</ul>
		</div>
		<h3 id="mnutoolHelp">Help</h3>
		<div>
			<ul id="mnubandHelp">
				<li id="mnutoolAboutHRPro"><a href="#">About OpenHR</a></li>
				<li class="hidden" id="mnutoolContentsAndIndex"><a href="#">Contents and Index</a></li>
			</ul>
		</div>
</div>

<div class="accessibility ui-accordion-header ui-helper-reset ui-state-default ui-corner-all">
	<ul>
		<li><a class="size-big" href="#" id="FontSizeBig" title="Large font"><span>A</span></a></li>
		<li><a class="size-default" href="#" id="FontSizeDefault" title="Default font"><span>A</span></a></li>
		<li><a class="size-small" href="#" id="FontSizeSmall" title="Small font"><span>A</span></a></li>
	</ul>
</div>

<script type="text/javascript">
	function documentReady()
	{
		//Change style of the third-level Database section leafs
		$('[aria-labelledby="mnutoolDatabase"] [id^="PV_"] .ui-state-default').css('font-weight', 'normal');

		//Get current accessibility settings (if any)
		var accordionFontSize = OpenHR.GetRegistrySetting("HR Pro", "AccordionAccessibilityFontSizeSize", "accordion-font-size"); //Font size

		if (accordionFontSize != ""){
			$(".accordion").css("font-size", accordionFontSize);
		}
	}

	$(document).ready(function() {
		setTimeout("documentReady()", 500);
	});


	//Change fonts on clicking the appropriate link
	$("#FontSizeBig").click(function () {
		$(".accordion").css("font-size", "large");
		//$(".accordion h3").css("font-size", "16pt");
		OpenHR.SaveRegistrySetting("HR Pro", "AccordionAccessibilityFontSize", "accordion-font-size", "large");
	});

	$("#FontSizeDefault").click(function (){
		$(".accordion").css("font-size", "1em");
		//$(".accordion h3").css("font-size", "14pt");
		OpenHR.SaveRegistrySetting("HR Pro", "AccordionAccessibilityFontSize", "accordion-font-size", "1em");
	});

	$("#FontSizeSmall").click(function (){
		$(".accordion").css("font-size", "small");
		//$(".accordion h3").css("font-size", "12pt");
		OpenHR.SaveRegistrySetting("HR Pro", "AccordionAccessibilityFontSize", "accordion-font-size", "small");
	});
</script>

<FORM action="" method=POST id=frmMenuInfo name=frmMenuInfo>
<%
	Response.Write("<INPUT type=""hidden"" id=txtDefaultStartPage name=txtDefaultStartPage value=""" & Replace(Session("DefaultStartPage"), """", "&quot;") & """>")
	Response.Write("<INPUT type=""hidden"" id=txtDatabase name=txtDatabase value=""" & Replace(Session("Database"), """", "&quot;") & """>")
%>
	<input type="hidden" id="txtUserType" name="txtUserType" value='<%=session("userType")%>'>

	<input type="hidden" id="txtPersonnel_EmpTableID" name="txtPersonnel_EmpTableID" value='<%=session("Personnel_EmpTableID")%>'>

	<input type="hidden" id="txtTB_EmpTableID" name="txtTB_EmpTableID" value='<%=session("TB_EmpTableID")%>'>
	<input type="hidden" id="txtTB_CourseTableID" name="txtTB_CourseTableID" value='<%=session("TB_CourseTableID")%>'>
	<input type="hidden" id="txtTB_CourseCancelDateColumnID" name="txtTB_CourseCancelDateColumnID" value='<%=session("TB_CourseCancelDateColumnID")%>'>
	<input type="hidden" id="txtWaitListOverRideColumnID" name="txtWaitListOverRideColumnID" value='<%=session("TB_WaitListOverRideColumnID")%>'>
	<input type="hidden" id="txtTB_TBTableID" name="txtTB_TBTableID" value='<%=session("TB_TBTableID")%>'>
	<input type="hidden" id="txtTB_TBTableSelect" name="txtTB_TBTableSelect" value='<%=session("TB_TBTableSelect")%>'>
	<input type="hidden" id="txtTB_TBTableInsert" name="txtTB_TBTableInsert" value='<%=session("TB_TBTableInsert")%>'>
	<input type="hidden" id="txtTB_TBTableUpdate" name="txtTB_TBTableUpdate" value='<%=session("TB_TBTableUpdate")%>'>
	<input type="hidden" id="txtTB_TBStatusColumnID" name="txtTB_TBStatusColumnID" value='<%=session("TB_TBStatusColumnID")%>'>
	<input type="hidden" id="txtTB_TBStatusColumnUpdate" name="txtTB_TBStatusColumnUpdate" value='<%=session("TB_TBStatusColumnUpdate")%>'>
	<input type="hidden" id="txtTB_TBCancelDateColumnID" name="txtTB_TBCancelDateColumnID" value='<%=session("TB_TBCancelDateColumnID")%>'>
	<input type="hidden" id="txtTB_TBCancelDateColumnUpdate" name="txtTB_TBCancelDateColumnUpdate" value='<%=session("TB_TBCancelDateColumnUpdate")%>'>
	<input type="hidden" id="txtTB_TBStatusPExists" name="txtTB_TBStatusPExists" value='<%=session("TB_TBStatusPExists")%>'>
	<input type="hidden" id="txtTB_WaitListTableID" name="txtTB_WaitListTableID" value='<%=session("TB_WaitListTableID")%>'>
	<input type="hidden" id="txtTB_WaitListTableInsert" name="txtTB_WaitListTableInsert" value='<%=session("TB_WaitListTableInsert")%>'>
	<input type="hidden" id="txtTB_WaitListTableDelete" name="txtTB_WaitListTableDelete" value='<%=session("TB_WaitListTableDelete")%>'>
	<input type="hidden" id="txtTB_WaitListCourseTitleColumnID" name="txtTB_WaitListCourseTitleColumnID" value='<%=session("TB_WaitListCourseTitleColumnID")%>'>
	<input type="hidden" id="txtTB_WaitListCourseTitleColumnUpdate" name="txtTB_WaitListCourseTitleColumnUpdate" value='<%=session("TB_WaitListCourseTitleColumnUpdate")%>'>
	<input type="hidden" id="txtTB_WaitListCourseTitleColumnSelect" name="txtTB_WaitListCourseTitleColumnSelect" value='<%=session("TB_WaitListCourseTitleColumnSelect")%>'>
	<input type="hidden" id="txtPrimaryStartMode" name="txtPrimaryStartMode" value='<%=session("PrimaryStartMode")%>'>
	<input type="hidden" id="txtHistoryStartMode" name="txtHistoryStartMode" value='<%=session("HistoryStartMode")%>'>
	<input type="hidden" id="txtLookupStartMode" name="txtLookupStartMode" value='<%=session("LookupStartMode")%>'>
	<input type="hidden" id="txtQuickAccessStartMode" name="txtQuickAccessStartMode" value='<%=session("QuickAccessStartMode")%>'>
	<input type="hidden" id="txtDesktopColour" name="txtDesktopColour" value='<%=session("DesktopColour")%>'>

	<input type="hidden" id="txtWFEnabled" name="txtWFEnabled" value='<%=session("WF_Enabled")%>'>
	<input type="hidden" id="txtWFOutOfOfficeEnabled" name="txtWFOutOfOfficeEnabled" value='<%=session("WF_OutOfOfficeConfigured")%>'>
	<input type="hidden" id="txtWFShowOutOfOffice" name="txtWFShowOutOfOffice" value='<%=Session("WF_ShowOutOfOffice")%>'>

	<input type="hidden" id="txtDoneDatabaseMenu" name="txtDoneDatabaseMenu" value="0">
	<input type="hidden" id="txtDoneQuickEntryMenu" name="txtDoneQuickEntryMenu" value="0">
	<input type="hidden" id="txtDoneTableScreensMenu" name="txtDoneTableScreensMenu" value="0">
	<input type="hidden" id="txtDoneSelfServiceStart" name="txtDoneSelfServiceStart" value="0">

	<input type="hidden" id="txtMenuSaved" name="txtMenuSaved" value="0">
	
	<input type="hidden" id="txtProgressMessage" name="txtProgressMessage" value=""/>
</FORM>

<script type="text/javascript">

	menu_window_onload();
	$("#contextmenu").fadeIn("slow");
	$(".accordion").accordion("resize");

</script>
