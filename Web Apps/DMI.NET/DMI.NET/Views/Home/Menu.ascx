<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>
<%@ Import Namespace="System.Collections.ObjectModel" %>
<%@ Import Namespace="HR.Intranet.Server" %>
<%@ Import Namespace="HR.Intranet.Server.Structures" %>

<%
	On Error Resume Next

	Dim sErrorDescription As String
	Dim avPrimaryMenuInfo
	Dim avSubMenuInfo
	Dim avQuickEntryMenuInfo
	Dim avTableMenuInfo As Collection(Of HR.Intranet.Server.Structures.TableScreen)
	Dim avHistoryMenuInfo As Collection(Of HR.Intranet.Server.Structures.HistoryScreen)
	Dim iLoop As Integer
	Dim iLoop2 As Integer
	Dim iCount As Integer
	Dim objMenu As HR.Intranet.Server.Menu
	Dim sToolCaption As String
	Dim sToolID As String
	
	Dim objSession As SessionInfo = CType(Session("sessionContext"), SessionInfo)

	sErrorDescription = ""
	
	objMenu = New HR.Intranet.Server.Menu()
		
	objMenu.Username = Session("username")
	
	If Session("avPrimaryMenuInfo") Is Nothing Then
		' only call 'setuptablescollection' if not already done.
		CallByName(objMenu, "Connection", CallType.Let, Session("databaseConnection"))
	End If
	
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
	
	For iLoop = 1 To UBound(avPrimaryMenuInfo, 2)
		If avPrimaryMenuInfo(4, iLoop) > 0 Then
			' The user has 'read' permission on the table, and no views on the table.
			' There is only one screen defined for the table.
				
			' Add a menu option to call up the primary table screen.
			' new method to insert a new menu item.
			Response.Write("  menu_insertMenuItem('mnubandDatabase', '" & CleanStringForJavaScript(Replace(avPrimaryMenuInfo(2, iLoop), "_", " ")) & "..." & "', 'PT_" & CleanStringForJavaScript(avPrimaryMenuInfo(1, iLoop)) & "_0_" & CleanStringForJavaScript(avPrimaryMenuInfo(4, iLoop)) & "');" & vbCrLf & vbCrLf)
		ElseIf avPrimaryMenuInfo(7, iLoop) > 0 Then
			' The user does NOT have 'read' permission on the table, but does have
			' 'read' permission on one view of the table.
			' There is only one screen defined for the view.
			' new method to insert a new menu item.
			Response.Write("  menu_insertMenuItem('mnubandDatabase', '" & CleanStringForJavaScript(Replace(avPrimaryMenuInfo(2, iLoop), "_", " ")) & " (" & CleanStringForJavaScript(Replace(avPrimaryMenuInfo(8, iLoop), "_", " ")) & " view)..." & "', 'PV_" & CleanStringForJavaScript(avPrimaryMenuInfo(1, iLoop)) & "_" & CleanStringForJavaScript(avPrimaryMenuInfo(7, iLoop)) & "_" & CleanStringForJavaScript(avPrimaryMenuInfo(10, iLoop)) & "');" & vbCrLf & vbCrLf)
		ElseIf (avPrimaryMenuInfo(9, iLoop) > 0) Or ((avPrimaryMenuInfo(5, iLoop) = True) And (avPrimaryMenuInfo(3, iLoop) > 0)) Then
			' The user has 'read' permission on the table, and the table has more than one screen defined for it.
			' Or there are views on the table.
			'Instantiate the submenu heading tool and set properties

			' new method to insert a new submenu item.
			Response.Write("  menu_insertSubMenuItem('mnubandDatabase', '" & CleanStringForJavaScript(Replace(avPrimaryMenuInfo(2, iLoop), "_", " ")) & "', 'PS_" & CleanStringForJavaScript(avPrimaryMenuInfo(1, iLoop)) & "', 'mnusubband_" & CleanStringForJavaScript(avPrimaryMenuInfo(2, iLoop)) & "');" & vbCrLf & vbCrLf)
			
			' Add the submenu.
			'If Session("avSubMenuInfo") Is Nothing Then
			avSubMenuInfo = objMenu.GetPrimaryTableSubMenu(CLng(avPrimaryMenuInfo(1, iLoop)))
			'Session("avSubMenuInfo") = avSubMenuInfo
			'Else
			'avSubMenuInfo = Session("avSubMenuInfo")
			'End If
			
			Response.Write("  lngLastScreenID = 0;" & vbCrLf)
			Response.Write("  sLastToolName = """";" & vbCrLf)
			
			For iLoop2 = 1 To UBound(avSubMenuInfo, 2)
				sToolCaption = ""
				sToolID = ""
				
				If avSubMenuInfo(3, iLoop2) > 0 Then
					sToolID = "PV_" & CleanStringForJavaScript(avPrimaryMenuInfo(1, iLoop)) & "_" & CleanStringForJavaScript(avSubMenuInfo(3, iLoop2)) & "_" & CleanStringForJavaScript(avSubMenuInfo(1, iLoop2))
					sToolCaption = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & CleanStringForJavaScript(Replace(avSubMenuInfo(4, iLoop2), "_", " ")) & " view..."
				Else
					sToolID = "PT_" & CleanStringForJavaScript(avPrimaryMenuInfo(1, iLoop)) & "_0_" & CleanStringForJavaScript(avSubMenuInfo(1, iLoop2))
					sToolCaption = CleanStringForJavaScript(Replace(avSubMenuInfo(2, iLoop2), "_", " ")) & "..."
				End If

				Response.Write("  lngLastScreenID = " & CleanStringForJavaScript(avSubMenuInfo(1, iLoop2)) & ";" & vbCrLf & vbCrLf)
				Response.Write("	sLastToolName = '" & sToolID & "'" & vbCrLf)

				' new method to insert a new menu item.
				Response.Write("  menu_insertMenuItem('mnusubband_" & CleanStringForJavaScript(avPrimaryMenuInfo(2, iLoop)) & "', '" & sToolCaption & "', '" & sToolID & "','someClass');" & vbCrLf & vbCrLf)
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
	Response.Write("  if (frmMenuInfo.txtDoneQuickEntryMenu.value == 1) {" & vbCrLf)
	Response.Write("	//  return;" & vbCrLf)
	Response.Write("  }" & vbCrLf & vbCrLf)
	Response.Write("  frmMenuInfo.txtDoneQuickEntryMenu.value = 1;" & vbCrLf & vbCrLf)
	Response.Write("  //lngQuickEntryCount = 0;" & vbCrLf)
	
	If Session("avQuickEntryMenuInfo") Is Nothing Then
		avQuickEntryMenuInfo = objMenu.GetQuickEntryScreens
		Session("avQuickEntryMenuInfo") = avQuickEntryMenuInfo
	Else
		avQuickEntryMenuInfo = Session("avQuickEntryMenuInfo")
	End If
	
	For iCount = 1 To UBound(avQuickEntryMenuInfo, 2)
		Response.Write("  //lngQuickEntryCount = lngQuickEntryCount + 1;" & vbCrLf)
		Response.Write("  //objFileTool = abMainMenu.Tools.add(0, ""QE_" & CleanStringForJavaScript(avQuickEntryMenuInfo(1, iCount)) & "_0_" & CleanStringForJavaScript(avQuickEntryMenuInfo(3, iCount)) & """);" & vbCrLf)
		Response.Write("  //objFileTool.Caption = """ & CleanStringForJavaScript(Replace(avQuickEntryMenuInfo(2, iCount), "_", " ")) & "..."";" & vbCrLf)
		Response.Write("  //objFileTool.Style = 0;" & vbCrLf)

		Response.Write("    //iIndex = -1;" & vbCrLf)
		Response.Write("    //for (iLoop2=0; iLoop2 < abMainMenu.Bands(""mnubandQuickEntry"").Tools.Count(); iLoop2++) {" & vbCrLf)
		Response.Write("		//	sCaption1 = abMainMenu.Bands(""mnubandQuickEntry"").Tools(iLoop2).Caption.toLowerCase();" & vbCrLf)
		Response.Write("		//	sCaption1 = sCaption1.substr(0, sCaption1.length - 3);" & vbCrLf)
		Response.Write("		//	sCaption2 = objFileTool.Caption.toLowerCase();" & vbCrLf)
		Response.Write("		//	sCaption2 = sCaption2.substr(0, sCaption2.length - 3);" & vbCrLf)
		Response.Write("    //  if (sCaption1 > sCaption2) {" & vbCrLf)
		Response.Write("    //    iIndex = iLoop2;" & vbCrLf)
		Response.Write("    //    break;" & vbCrLf)
		Response.Write("    //  }" & vbCrLf)
		Response.Write("    //}" & vbCrLf)
		Response.Write("    //abMainMenu.Bands(""mnubandQuickEntry"").Tools.Insert(iIndex, objFileTool);" & vbCrLf & vbCrLf)
		
		' new method to insert a new menu item.
		Response.Write("  menu_insertMenuItem('mnubandQuickEntry', '" & CleanStringForJavaScript(Replace(avQuickEntryMenuInfo(2, iCount), "_", " ")) & "..." & "', 'QE_" & CleanStringForJavaScript(avQuickEntryMenuInfo(1, iCount)) & "_0_" & CleanStringForJavaScript(avQuickEntryMenuInfo(3, iCount)) & "');" & vbCrLf & vbCrLf)
		
	Next

	Response.Write("  if (lngQuickEntryCount == 0) {" & vbCrLf)
	Response.Write("	//	abMainMenu.Bands(""mnubandDatabase"").Tools(""mnutoolQuickEntry"").enabled = false;" & vbCrLf)
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
		Response.Write("  //objFileTool = abMainMenu.Tools.add(0, ""TS_" & CleanStringForJavaScript(objTableScreen.TableID) & "_0_" & CleanStringForJavaScript(objTableScreen.ScreenID) & """);" & vbCrLf)
		Response.Write("  //objFileTool.Caption = """ & CleanStringForJavaScript(Replace(objTableScreen.TableName, "_", " ")) & "..."";" & vbCrLf)
		Response.Write("  //objFileTool.Style = 0;" & vbCrLf)
		Response.Write("  //abMainMenu.Bands(""mnubandTableScreens"").Tools.insert(0, objFileTool);" & vbCrLf & vbCrLf)
		Response.Write("  menu_insertMenuItem('mnubandTableScreens', '" & CleanStringForJavaScript(Replace(objTableScreen.TableName, "_", " ")) & "..." & "', 'TS_" & CleanStringForJavaScript(objTableScreen.TableID) & "_0_" & CleanStringForJavaScript(objTableScreen.ScreenID) & "');" & vbCrLf & vbCrLf)
	Next
	
	Response.Write("  if (lngTableScreensCount == 0) {" & vbCrLf)
	Response.Write("	//	abMainMenu.Bands(""mnubandDatabase"").Tools(""mnutoolTableScreens"").enabled = false;" & vbCrLf)
	Response.Write("  }" & vbCrLf)

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
		
	avHistoryMenuInfo = objMenu.GetHistoryScreens
	
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
	CallByName(objUtilities, "Connection", CallType.Let, Session("databaseConnection"))
	Session("UtilitiesObject") = objUtilities
	
	Dim objOLE As New HR.Intranet.Server.Ole
	CallByName(objOLE, "Connection", CallType.Let, Session("databaseConnection"))
	objOLE.TempLocationPhysical = "\\" & Request.ServerVariables("SERVER_NAME") & "\HRProTemp$\"
	Session("OLEObject") = objOLE
	objOLE = Nothing
	
	If Len(sErrorDescription) = 0 Then
		Dim prm1 As ADODB.Parameter
		Dim prm2 As ADODB.Parameter
		Dim prm3 As ADODB.Parameter
		Dim prm4 As ADODB.Parameter
		
		Dim cmdMisc = New ADODB.Command
		cmdMisc.CommandText = "spASRIntGetMiscParameters"
		cmdMisc.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
		cmdMisc.ActiveConnection = Session("databaseConnection")

		prm1 = cmdMisc.CreateParameter("param1", 200, 2, 8000) '200=varchar, 2=output, 8000=size
		cmdMisc.Parameters.Append(prm1)

		prm2 = cmdMisc.CreateParameter("param2", 200, 2, 8000) '200=varchar, 2=output, 8000=size
		cmdMisc.Parameters.Append(prm2)

		prm3 = cmdMisc.CreateParameter("param3", 200, 2, 8000) '200=varchar, 2=output, 8000=size
		cmdMisc.Parameters.Append(prm3)
		
		prm4 = cmdMisc.CreateParameter("param4", 200, 2, 8000) '200=varchar, 2=output, 8000=size
		cmdMisc.Parameters.Append(prm4)
		
		Err.Clear()
		cmdMisc.Execute()
		
		Response.Write("<input TYPE=Hidden NAME=txtCFG_PCL ID=txtCFG_PCL VALUE='" & cmdMisc.Parameters("param1").Value & "'>" & vbCrLf)
		Response.Write("<input TYPE=Hidden NAME=txtCFG_BA ID=txtCFG_BA VALUE='" & cmdMisc.Parameters("param2").Value & "'>" & vbCrLf)
		Response.Write("<input TYPE=Hidden NAME=txtCFG_LD ID=txtCFG_LD VALUE='" & cmdMisc.Parameters("param3").Value & "'>" & vbCrLf)
		Response.Write("<input TYPE=Hidden NAME=txtCFG_RT ID=txtCFG_RT VALUE='" & cmdMisc.Parameters("param4").Value & "'>" & vbCrLf)
	End If
		
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
	

	Dim iAbsenceEnabled = 0
	If Len(sErrorDescription) = 0 Then
		Dim cmdAbsenceModule = New ADODB.Command
		cmdAbsenceModule.CommandText = "spASRIntActivateModule"
		cmdAbsenceModule.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
		cmdAbsenceModule.ActiveConnection = Session("databaseConnection")
		cmdAbsenceModule.CommandTimeout = 300

		Dim prmModuleKey = cmdAbsenceModule.CreateParameter("moduleKey", 200, 1, 50) '200=varchar, 1=input, 50=size
		cmdAbsenceModule.Parameters.Append(prmModuleKey)
		prmModuleKey.Value = "ABSENCE"

		Dim prmEnabled = cmdAbsenceModule.CreateParameter("enabled", 11, 2)	' 11=bit, 2=output
		cmdAbsenceModule.Parameters.Append(prmEnabled)

		Err.Clear()
		cmdAbsenceModule.Execute()

		iAbsenceEnabled = CInt(cmdAbsenceModule.Parameters("enabled").Value)
		If iAbsenceEnabled < 0 Then
			iAbsenceEnabled = 1
		End If
		cmdAbsenceModule = Nothing
	End If

	Response.Write("<input type='hidden' id=txtAbsenceEnabled name=txtAbsenceEnabled value=" & iAbsenceEnabled & ">")
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
				<li id="mnutoolTableScreens"><a href="#">Table Screens</a>
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
	<INPUT type="hidden" id=txtIEVersion name=txtIEVersion value=<%=session("IEVersion")%>>
	<INPUT type="hidden" id=txtUserType name=txtUserType value=<%=session("userType")%>>

	<INPUT type="hidden" id=txtPersonnel_EmpTableID name=txtPersonnel_EmpTableID value=<%=session("Personnel_EmpTableID")%>>

	<INPUT type="hidden" id=txtTB_EmpTableID name=txtTB_EmpTableID value=<%=session("TB_EmpTableID")%>>
	<INPUT type="hidden" id=txtTB_CourseTableID name=txtTB_CourseTableID value=<%=session("TB_CourseTableID")%>>
	<INPUT type="hidden" id=txtTB_CourseCancelDateColumnID name=txtTB_CourseCancelDateColumnID value=<%=session("TB_CourseCancelDateColumnID")%>>
	<INPUT type="hidden" id=txtWaitListOverRideColumnID name=txtWaitListOverRideColumnID value=<%=session("TB_WaitListOverRideColumnID")%>>
	<INPUT type="hidden" id=txtTB_TBTableID name=txtTB_TBTableID value=<%=session("TB_TBTableID")%>>
	<INPUT type="hidden" id=txtTB_TBTableSelect name=txtTB_TBTableSelect value=<%=session("TB_TBTableSelect")%>>
	<INPUT type="hidden" id=txtTB_TBTableInsert name=txtTB_TBTableInsert value=<%=session("TB_TBTableInsert")%>>
	<INPUT type="hidden" id=txtTB_TBTableUpdate name=txtTB_TBTableUpdate value=<%=session("TB_TBTableUpdate")%>>
	<INPUT type="hidden" id=txtTB_TBStatusColumnID name=txtTB_TBStatusColumnID value=<%=session("TB_TBStatusColumnID")%>>
	<INPUT type="hidden" id=txtTB_TBStatusColumnUpdate name=txtTB_TBStatusColumnUpdate value=<%=session("TB_TBStatusColumnUpdate")%>>
	<INPUT type="hidden" id=txtTB_TBCancelDateColumnID name=txtTB_TBCancelDateColumnID value=<%=session("TB_TBCancelDateColumnID")%>>
	<INPUT type="hidden" id=txtTB_TBCancelDateColumnUpdate name=txtTB_TBCancelDateColumnUpdate value=<%=session("TB_TBCancelDateColumnUpdate")%>>
	<INPUT type="hidden" id=txtTB_TBStatusPExists name=txtTB_TBStatusPExists value=<%=session("TB_TBStatusPExists")%>>
	<INPUT type="hidden" id=txtTB_WaitListTableID name=txtTB_WaitListTableID value=<%=session("TB_WaitListTableID")%>>
	<INPUT type="hidden" id=txtTB_WaitListTableInsert name=txtTB_WaitListTableInsert value=<%=session("TB_WaitListTableInsert")%>>
	<INPUT type="hidden" id=txtTB_WaitListTableDelete name=txtTB_WaitListTableDelete value=<%=session("TB_WaitListTableDelete")%>>
	<INPUT type="hidden" id=txtTB_WaitListCourseTitleColumnID name=txtTB_WaitListCourseTitleColumnID value=<%=session("TB_WaitListCourseTitleColumnID")%>>
	<INPUT type="hidden" id=txtTB_WaitListCourseTitleColumnUpdate name=txtTB_WaitListCourseTitleColumnUpdate value=<%=session("TB_WaitListCourseTitleColumnUpdate")%>>
	<INPUT type="hidden" id=txtTB_WaitListCourseTitleColumnSelect name=txtTB_WaitListCourseTitleColumnSelect value=<%=session("TB_WaitListCourseTitleColumnSelect")%>>
	<INPUT type="hidden" id=txtPrimaryStartMode name=txtPrimaryStartMode value=<%=session("PrimaryStartMode")%>>
	<INPUT type="hidden" id=txtHistoryStartMode name=txtHistoryStartMode value=<%=session("HistoryStartMode")%>>
	<INPUT type="hidden" id=txtLookupStartMode name=txtLookupStartMode value=<%=session("LookupStartMode")%>>
	<INPUT type="hidden" id=txtQuickAccessStartMode name=txtQuickAccessStartMode value=<%=session("QuickAccessStartMode")%>>
	<INPUT type="hidden" id=txtDesktopColour name=txtDesktopColour value=<%=session("DesktopColour")%>>

	<INPUT type="hidden" id=txtWFEnabled name=txtWFEnabled value=<%=session("WF_Enabled")%>>
	<INPUT type="hidden" id=txtWFOutOfOfficeEnabled name=txtWFOutOfOfficeEnabled value=<%=session("WF_OutOfOfficeConfigured")%>>
	<input type="hidden" id="txtWFShowOutOfOffice" name="txtWFShowOutOfOffice" value=<%=Session("WF_ShowOutOfOffice")%>>
	
	<INPUT type="hidden" id=txtDoneDatabaseMenu name=txtDoneDatabaseMenu value=0>
	<INPUT type="hidden" id=txtDoneQuickEntryMenu name=txtDoneQuickEntryMenu value=0>
	<INPUT type="hidden" id=txtDoneTableScreensMenu name=txtDoneTableScreensMenu value=0>
	<INPUT type="hidden" id=txtDoneSelfServiceStart name=txtDoneSelfServiceStart value=0>

	<INPUT type="hidden" id=txtMenuSaved name=txtMenuSaved value=0>
</FORM>

<script type="text/javascript">

	menu_window_onload();
	$("#contextmenu").fadeIn("slow");
	$(".accordion").accordion("resize");

</script>
