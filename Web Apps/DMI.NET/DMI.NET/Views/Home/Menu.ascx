﻿<%@ Control Language="VB" Inherits="System.Web.Mvc.ViewUserControl" %>
<%@ Import Namespace="DMI.NET" %>

<%
	'NOT USED - This snippet has no return value.
	'Dim sReferringPage
	'' Only open the form if there was a referring page.
	'' If it wasn't then redirect to the login page.
	'sReferringPage = Request.ServerVariables("HTTP_REFERER") 
	'if inStrRev(sReferringPage, "/") > 0 then
	'	sReferringPage = mid(sReferringPage, inStrRev(sReferringPage, "/") + 1)
	'end if
	
	'if ucase(sReferringPage) <> ucase("clear.asp") then
	'	'Response.Redirect("login.asp")
	'end if	

%>

<script src="<%: Url.Content("~/Include/ctl_SetFont.txt") %>" type="text/javascript"></script>

<script src="<%: Url.Content("~/Scripts/FormScripts/menu.js?x=1") %>" type="text/javascript"></script>

<%
	On Error Resume Next

	Dim sErrorDescription As String
	Dim avPrimaryMenuInfo
	Dim avSubMenuInfo
	Dim avQuickEntryMenuInfo
	Dim avTableMenuInfo
	Dim avHistoryMenuInfo
	Dim iLoop As Integer
	Dim iLoop2 As Integer
	Dim iNextIndex As Integer
	Dim iCount As Integer
	Dim objMenu
	Dim sToolCaption As String
	Dim sToolID As String
	
	sErrorDescription = ""
	
	objMenu = CreateObject("COAIntServer.Menu")
	objMenu.Username = Session("username")
	CallByName(objMenu, "Connection", CallType.Let, Session("databaseConnection"))
	
	Response.Write(vbCrLf & "<SCRIPT LANGUAGE=javascript>" & vbCrLf)
	Response.Write("<!--" & vbCrLf)

	' ------------------------------------------------------------------------------
	' Create the sub-routine to populate the database menu with the tables available
	' to the current user.
	' ------------------------------------------------------------------------------

	
	Response.Write("function refreshDatabaseMenu() {" & vbCrLf)
	Response.Write("  var objFileTool;" & vbCrLf)
	Response.Write("  var sLastToolName;" & vbCrLf)
	Response.Write("  var lngLastScreenID;" & vbCrLf & vbCrLf)
	Response.Write("  if (frmMenuInfo.txtDoneDatabaseMenu.value == 1) {" & vbCrLf)
	Response.Write("    return;" & vbCrLf)
	Response.Write("  }" & vbCrLf & vbCrLf)
	Response.Write("  frmMenuInfo.txtDoneDatabaseMenu.value = 1;" & vbCrLf & vbCrLf)
	
	avPrimaryMenuInfo = objMenu.GetPrimaryTableMenu

	For iLoop = 1 To UBound(avPrimaryMenuInfo, 2)
		If avPrimaryMenuInfo(4, iLoop) > 0 Then
			' The user has 'read' permission on the table, and no views on the table.
			' There is only one screen defined for the table.
        
			' Add a menu option to call up the primary table screen.
			Response.Write("  //objFileTool = abMainMenu.Tools.add(0, ""PT_" & cleanStringForJavaScript(avPrimaryMenuInfo(1, iLoop)) & "_0_" & cleanStringForJavaScript(avPrimaryMenuInfo(4, iLoop)) & """);" & vbCrLf)
			Response.Write("  //objFileTool.Caption = """ & cleanStringForJavaScript(Replace(avPrimaryMenuInfo(2, iLoop), "_", " ")) & "..."";" & vbCrLf)
			Response.Write("  //objFileTool.Style = 0;" & vbCrLf)
			Response.Write("  //abMainMenu.Bands(""mnubandDatabase"").Tools.insert(0, objFileTool);" & vbCrLf & vbCrLf)
			
			' new method to insert a new menu item.
			Response.Write("  menu_insertMenuItem('mnubandDatabase', '" & cleanStringForJavaScript(Replace(avPrimaryMenuInfo(2, iLoop), "_", " ")) & "..." & "', 'PT_" & cleanStringForJavaScript(avPrimaryMenuInfo(1, iLoop)) & "_0_" & cleanStringForJavaScript(avPrimaryMenuInfo(4, iLoop)) & "');" & vbCrLf & vbCrLf)
			
		ElseIf avPrimaryMenuInfo(7, iLoop) > 0 Then
			' The user does NOT have 'read' permission on the table, but does have
			' 'read' permission on one view of the table.
			' There is only one screen defined for the view.
			Response.Write("  //objFileTool = abMainMenu.Tools.add(0, ""PV_" & cleanStringForJavaScript(avPrimaryMenuInfo(1, iLoop)) & "_" & cleanStringForJavaScript(avPrimaryMenuInfo(7, iLoop)) & "_" & cleanStringForJavaScript(avPrimaryMenuInfo(10, iLoop)) & """);" & vbCrLf)
			Response.Write("  //objFileTool.Caption = """ & cleanStringForJavaScript(Replace(avPrimaryMenuInfo(2, iLoop), "_", " ")) & " (" & cleanStringForJavaScript(Replace(avPrimaryMenuInfo(8, iLoop), "_", " ")) & " view)..."";" & vbCrLf)
			Response.Write("  //objFileTool.Style = 0;" & vbCrLf)
			Response.Write("  //abMainMenu.Bands(""mnubandDatabase"").Tools.insert(0, objFileTool);" & vbCrLf & vbCrLf)

			' new method to insert a new menu item.
			Response.Write("  menu_insertMenuItem('mnubandDatabase', '" & cleanStringForJavaScript(Replace(avPrimaryMenuInfo(2, iLoop), "_", " ")) & " (" & cleanStringForJavaScript(Replace(avPrimaryMenuInfo(8, iLoop), "_", " ")) & " view)..." & "', 'PV_" & cleanStringForJavaScript(avPrimaryMenuInfo(1, iLoop)) & "_" & cleanStringForJavaScript(avPrimaryMenuInfo(7, iLoop)) & "_" & cleanStringForJavaScript(avPrimaryMenuInfo(10, iLoop)) & "');" & vbCrLf & vbCrLf)


		ElseIf (avPrimaryMenuInfo(9, iLoop) > 0) Or _
  ((avPrimaryMenuInfo(5, iLoop) = True) And (avPrimaryMenuInfo(3, iLoop) > 0)) Then
			' The user has 'read' permission on the table, and the table has more than one screen defined for it.
			' Or there are views on the table.
  
			Response.Write("  //abMainMenu.Bands.add(""mnusubband_" & cleanStringForJavaScript(avPrimaryMenuInfo(2, iLoop)) & """);" & vbCrLf)
			Response.Write("  //abMainMenu.Bands(""mnusubband_" & cleanStringForJavaScript(avPrimaryMenuInfo(2, iLoop)) & """).type = 2;" & vbCrLf & vbCrLf)
        
			'Instantiate the submenu heading tool and set properties
			Response.Write("  //objFileTool = abMainMenu.Tools.add(0, ""PS_" & cleanStringForJavaScript(avPrimaryMenuInfo(1, iLoop)) & """);" & vbCrLf)
			Response.Write("  //objFileTool.Caption = """ & cleanStringForJavaScript(Replace(avPrimaryMenuInfo(2, iLoop), "_", " ")) & """;" & vbCrLf)
			Response.Write("  //objFileTool.Style = 0;" & vbCrLf)
			Response.Write("  //objFileTool.SubBand = ""mnusubband_" & cleanStringForJavaScript(avPrimaryMenuInfo(2, iLoop)) & """;" & vbCrLf)
			Response.Write("  //abMainMenu.Bands(""mnubandDatabase"").Tools.insert(0, objFileTool);" & vbCrLf & vbCrLf)

			' new method to insert a new submenu item.
			Response.Write("  menu_insertSubMenuItem('mnubandDatabase', '" & cleanStringForJavaScript(Replace(avPrimaryMenuInfo(2, iLoop), "_", " ")) & "', 'PS_" & cleanStringForJavaScript(avPrimaryMenuInfo(1, iLoop)) & "', 'mnusubband_" & cleanStringForJavaScript(avPrimaryMenuInfo(2, iLoop)) & "');" & vbCrLf & vbCrLf)
			
			' Add the submenu.
			avSubMenuInfo = objMenu.GetPrimaryTableSubMenu(CLng(avPrimaryMenuInfo(1, iLoop)))

			Response.Write("  lngLastScreenID = 0;" & vbCrLf)
			Response.Write("  sLastToolName = """";" & vbCrLf)
      
			
			For iLoop2 = 1 To UBound(avSubMenuInfo, 2)
				
				sToolCaption = ""
				sToolID = ""
				
				If avSubMenuInfo(3, iLoop2) > 0 Then
					Response.Write("  //objFileTool = abMainMenu.Tools.add(0, ""PV_" & cleanStringForJavaScript(avPrimaryMenuInfo(1, iLoop)) & "_" & cleanStringForJavaScript(avSubMenuInfo(3, iLoop2)) & "_" & cleanStringForJavaScript(avSubMenuInfo(1, iLoop2)) & """);" & vbCrLf)
					sToolID = "PV_" & cleanStringForJavaScript(avPrimaryMenuInfo(1, iLoop)) & "_" & cleanStringForJavaScript(avSubMenuInfo(3, iLoop2)) & "_" & cleanStringForJavaScript(avSubMenuInfo(1, iLoop2))
					Response.Write("  //objFileTool.Caption = """ & cleanStringForJavaScript(Replace(avSubMenuInfo(2, iLoop2), "_", " ")) & " (" & cleanStringForJavaScript(Replace(avSubMenuInfo(4, iLoop2), "_", " ")) & " view)..."";" & vbCrLf)
					sToolCaption = cleanStringForJavaScript(Replace(avSubMenuInfo(2, iLoop2), "_", " ")) & " (" & cleanStringForJavaScript(Replace(avSubMenuInfo(4, iLoop2), "_", " ")) & " view)..."
				Else
					Response.Write("  //objFileTool = abMainMenu.Tools.add(0, ""PT_" & cleanStringForJavaScript(avPrimaryMenuInfo(1, iLoop)) & "_0_" & cleanStringForJavaScript(avSubMenuInfo(1, iLoop2)) & """);" & vbCrLf)
					sToolID = "PT_" & cleanStringForJavaScript(avPrimaryMenuInfo(1, iLoop)) & "_0_" & cleanStringForJavaScript(avSubMenuInfo(1, iLoop2))
					Response.Write("  //objFileTool.Caption = """ & cleanStringForJavaScript(Replace(avSubMenuInfo(2, iLoop2), "_", " ")) & "..."";" & vbCrLf)
					sToolCaption = cleanStringForJavaScript(Replace(avSubMenuInfo(2, iLoop2), "_", " ")) & "..."
				End If

				Response.Write("  //objFileTool.Style = 0;" & vbCrLf)

				Response.Write("  //if ((lngLastScreenID > 0) &&" & vbCrLf)
				Response.Write("  //  (lngLastScreenID != " & cleanStringForJavaScript(avSubMenuInfo(1, iLoop2)) & ")) {" & vbCrLf)
				Response.Write("  // abMainMenu.Bands(""mnusubband_" & cleanStringForJavaScript(avPrimaryMenuInfo(2, iLoop)) & """).Tools(sLastToolName).beginGroup = true;" & vbCrLf)
				Response.Write("  //}" & vbCrLf)
				Response.Write("  lngLastScreenID = " & cleanStringForJavaScript(avSubMenuInfo(1, iLoop2)) & ";" & vbCrLf & vbCrLf)
				Response.Write("  //sLastToolName = objFileTool.name;" & vbCrLf & vbCrLf)
				Response.Write("	sLastToolName = '" & sToolID & "'" & vbCrLf)

				Response.Write("  //abMainMenu.Bands(""mnusubband_" & cleanStringForJavaScript(avPrimaryMenuInfo(2, iLoop)) & """).Tools.insert(0, objFileTool);" & vbCrLf & vbCrLf)
				
				' new method to insert a new menu item.
				Response.Write("  menu_insertMenuItem('mnusubband_" & cleanStringForJavaScript(avPrimaryMenuInfo(2, iLoop)) & "', '" & sToolCaption & "', '" & sToolID & "');" & vbCrLf & vbCrLf)
				
				
			Next
		End If
	Next
	
	Response.Write("  //abMainMenu.RecalcLayout();" & vbCrLf)
	Response.Write("  //abMainMenu.ResetHooks();" & vbCrLf)
	Response.Write("  //abMainMenu.Refresh();" & vbCrLf)
	Response.Write("}" & vbCrLf & vbCrLf)
	
	
	' ------------------------------------------------------------------------------
	' Create the sub-routine to populate the quick entry menu.
	' ------------------------------------------------------------------------------
	Response.Write("function refreshQuickEntryMenu() {" & vbCrLf)
	Response.Write("  var objFileTool;" & vbCrLf)
	Response.Write("  var lngQuickEntryCount;" & vbCrLf & vbCrLf)
	Response.Write("  if (frmMenuInfo.txtDoneQuickEntryMenu.value == 1) {" & vbCrLf)
	Response.Write("	//  return;" & vbCrLf)
	Response.Write("  }" & vbCrLf & vbCrLf)
	Response.Write("  frmMenuInfo.txtDoneQuickEntryMenu.value = 1;" & vbCrLf & vbCrLf)
	Response.Write("  //lngQuickEntryCount = 0;" & vbCrLf)
	
	avQuickEntryMenuInfo = objMenu.GetQuickEntryScreens

	For iCount = 1 To UBound(avQuickEntryMenuInfo, 2)
		Response.Write("  //lngQuickEntryCount = lngQuickEntryCount + 1;" & vbCrLf)
		Response.Write("  //objFileTool = abMainMenu.Tools.add(0, ""QE_" & cleanStringForJavaScript(avQuickEntryMenuInfo(1, iCount)) & "_0_" & cleanStringForJavaScript(avQuickEntryMenuInfo(3, iCount)) & """);" & vbCrLf)
		Response.Write("  //objFileTool.Caption = """ & cleanStringForJavaScript(Replace(avQuickEntryMenuInfo(2, iCount), "_", " ")) & "..."";" & vbCrLf)
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
		Response.Write("  menu_insertMenuItem('mnubandQuickEntry', '" & cleanStringForJavaScript(Replace(avQuickEntryMenuInfo(2, iCount), "_", " ")) & "..." & "', 'QE_" & cleanStringForJavaScript(avQuickEntryMenuInfo(1, iCount)) & "_0_" & cleanStringForJavaScript(avQuickEntryMenuInfo(3, iCount)) & "');" & vbCrLf & vbCrLf)
		
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
	Response.Write("  if (frmMenuInfo.txtDoneTableScreensMenu.value == 1) {" & vbCrLf)
	Response.Write("	  return;" & vbCrLf)
	Response.Write("  }" & vbCrLf & vbCrLf)
	Response.Write("  frmMenuInfo.txtDoneTableScreensMenu.value = 1;" & vbCrLf & vbCrLf)
	Response.Write("  lngTableScreensCount = 0;" & vbCrLf)

	avTableMenuInfo = objMenu.GetTableScreens

	For iCount = 1 To UBound(avTableMenuInfo, 2)
		Response.Write("  lngTableScreensCount = lngTableScreensCount + 1;" & vbCrLf)
		Response.Write("  //objFileTool = abMainMenu.Tools.add(0, ""TS_" & cleanStringForJavaScript(avTableMenuInfo(1, iCount)) & "_0_" & cleanStringForJavaScript(avTableMenuInfo(3, iCount)) & """);" & vbCrLf)
		Response.Write("  //objFileTool.Caption = """ & cleanStringForJavaScript(Replace(avTableMenuInfo(2, iCount), "_", " ")) & "..."";" & vbCrLf)
		Response.Write("  //objFileTool.Style = 0;" & vbCrLf)
		Response.Write("  //abMainMenu.Bands(""mnubandTableScreens"").Tools.insert(0, objFileTool);" & vbCrLf & vbCrLf)

		' new method to insert a new menu item.
		Response.Write("  menu_insertMenuItem('mnubandTableScreens', '" & cleanStringForJavaScript(Replace(avTableMenuInfo(2, iCount), "_", " ")) & "..." & "', 'TS_" & cleanStringForJavaScript(avTableMenuInfo(1, iCount)) & "_0_" & cleanStringForJavaScript(avTableMenuInfo(3, iCount)) & "');" & vbCrLf & vbCrLf)
		
	Next
	
	Response.Write("  if (lngTableScreensCount == 0) {" & vbCrLf)
	Response.Write("	//	abMainMenu.Bands(""mnubandDatabase"").Tools(""mnutoolTableScreens"").enabled = false;" & vbCrLf)
	Response.Write("  }" & vbCrLf)

	Response.Write("}" & vbCrLf & vbCrLf)
	
	
	
	
	Response.Write("-->" & vbCrLf)
	Response.Write("</SCRIPT>" & vbCrLf)

	objMenu = Nothing

	Dim objUtilities
	objUtilities = CreateObject("COAIntServer.Utilities")
	CallByName(objUtilities, "Connection", CallType.Let, Session("databaseConnection"))
	Session("UtilitiesObject") = objUtilities
	
	Dim objOLE
	objOLE = CreateObject("COAIntServer.clsOLE")
	CallByName(objOLE, "Connection", CallType.Let, Session("databaseConnection"))
	objOLE.TempLocationPhysical = "\\" & Request.ServerVariables("SERVER_NAME") & "\HRProTemp$\"
	Session("OLEObject") = objOLE
	objOLE = Nothing
	
	If Len(sErrorDescription) = 0 Then
		Dim cmdMisc
		Dim prm1
		Dim prm2
		Dim prm3
		Dim prm4
		
		cmdMisc = CreateObject("ADODB.Command")
		cmdMisc.CommandText = "spASRIntGetMiscParameters"
		cmdMisc.CommandType = 4	' Stored Procedure
		cmdMisc.ActiveConnection = Session("databaseConnection")

		prm1 = cmdMisc.CreateParameter("param1", 200, 2, 8000) '200=varchar, 2=output, 8000=size
		cmdMisc.Parameters.Append(prm1)

		prm2 = cmdMisc.CreateParameter("param2", 200, 2, 8000) '200=varchar, 2=output, 8000=size
		cmdMisc.Parameters.Append(prm2)

		prm3 = cmdMisc.CreateParameter("param3", 200, 2, 8000) '200=varchar, 2=output, 8000=size
		cmdMisc.Parameters.Append(prm3)
		
		prm4 = cmdMisc.CreateParameter("param4", 200, 2, 8000) '200=varchar, 2=output, 8000=size
		cmdMisc.Parameters.Append(prm4)
		
		' Err = 0
		cmdMisc.Execute()
		
		Response.Write("<INPUT TYPE=Hidden NAME=txtCFG_PCL ID=txtCFG_PCL VALUE='" & cmdMisc.Parameters("param1").value & "'>" & vbCrLf)
		Response.Write("<INPUT TYPE=Hidden NAME=txtCFG_BA ID=txtCFG_BA VALUE='" & cmdMisc.Parameters("param2").value & "'>" & vbCrLf)
		Response.Write("<INPUT TYPE=Hidden NAME=txtCFG_LD ID=txtCFG_LD VALUE='" & cmdMisc.Parameters("param3").value & "'>" & vbCrLf)
		Response.Write("<INPUT TYPE=Hidden NAME=txtCFG_RT ID=txtCFG_RT VALUE='" & cmdMisc.Parameters("param4").value & "'>" & vbCrLf)
	End If
		
	' ------------------------------------------------------------------------------
	' Check what permissions the user has.
	' ------------------------------------------------------------------------------
	Dim fCustomReportsGranted = False
	Dim fCrossTabsGranted = False
	Dim fCalendarReportsGranted = False
	Dim fMailMergeGranted = False
	Dim fWorkflowGranted = False
	Dim fCalculationsGranted = False
	Dim fFiltersGranted = False
	Dim fPicklistsGranted = False
	Dim fNewUserGranted = False

	If Len(sErrorDescription) = 0 Then
		Dim cmdSystemPermissions = CreateObject("ADODB.Command")
		cmdSystemPermissions.CommandText = "sp_ASRIntGetSystemPermissions"
		cmdSystemPermissions.CommandType = 4 ' Stored Procedure
		cmdSystemPermissions.ActiveConnection = Session("databaseConnection")
		cmdSystemPermissions.CommandTimeout = 300

		' Err = 0
		Dim rstSystemPermissions = cmdSystemPermissions.Execute
		
		If (Err.Number <> 0) Then
			sErrorDescription = "The system permissions could not be read." & vbCrLf & formatError(Err.Description)
		End If
		
		If Len(sErrorDescription) = 0 Then
			Do While Not rstSystemPermissions.EOF
				Response.Write("<INPUT type='hidden' id=txtSysPerm_" & Replace(rstSystemPermissions.fields("KEY").value, " ", "_") & " name=txtSysPerm_" & Replace(rstSystemPermissions.fields("KEY").value, " ", "_") & " value=""" & rstSystemPermissions.fields("PERMITTED").value & """>" & vbCrLf)

				If (Left(rstSystemPermissions.fields("KEY").value, 13) = "CUSTOMREPORTS") And _
				 (rstSystemPermissions.fields("PERMITTED").value = 1) Then
					fCustomReportsGranted = True
				End If
				If (Left(rstSystemPermissions.fields("KEY").value, 9) = "CROSSTABS") And _
				 (rstSystemPermissions.fields("PERMITTED").value = 1) Then
					fCrossTabsGranted = True
				End If
				If (Left(rstSystemPermissions.fields("KEY").value, 15) = "CALENDARREPORTS") And _
				 (rstSystemPermissions.fields("PERMITTED").value = 1) Then
					fCalendarReportsGranted = True
				End If
				If (Left(rstSystemPermissions.fields("KEY").value, 9) = "MAILMERGE") And _
				 (rstSystemPermissions.fields("PERMITTED").value = 1) Then
					fMailMergeGranted = True
				End If
				If (Left(rstSystemPermissions.fields("KEY").value, 12) = "WORKFLOW_RUN") And _
				 (rstSystemPermissions.fields("PERMITTED").value = 1) Then
					fWorkflowGranted = True
				End If
				If (Left(rstSystemPermissions.fields("KEY").value, 12) = "CALCULATIONS") And _
				 (rstSystemPermissions.fields("PERMITTED").value = 1) Then
					fCalculationsGranted = True
				End If
				If (Left(rstSystemPermissions.fields("KEY").value, 7) = "FILTERS") And _
				 (rstSystemPermissions.fields("PERMITTED").value = 1) Then
					fFiltersGranted = True
				End If
				If (Left(rstSystemPermissions.fields("KEY").value, 9) = "PICKLISTS") And _
				 (rstSystemPermissions.fields("PERMITTED").value = 1) Then
					fPicklistsGranted = True
				End If
				If ((rstSystemPermissions.fields("KEY").value = "MODULEACCESS_SYSTEMMANAGER") Or _
				  (rstSystemPermissions.fields("KEY").value = "MODULEACCESS_SECURITYMANAGER")) And _
				 (rstSystemPermissions.fields("PERMITTED").value = 1) Then
					fNewUserGranted = True
				End If

				rstSystemPermissions.MoveNext()
			Loop

			' Release the ADO recordset and command objects.
			rstSystemPermissions.close()
		End If
	
		rstSystemPermissions = Nothing
		cmdSystemPermissions = Nothing
	End If

	Dim iAbsenceEnabled = 0
	If Len(sErrorDescription) = 0 Then
		Dim cmdAbsenceModule = CreateObject("ADODB.Command")
		cmdAbsenceModule.CommandText = "spASRIntActivateModule"
		cmdAbsenceModule.CommandType = 4	' Stored Procedure
		cmdAbsenceModule.ActiveConnection = Session("databaseConnection")
		cmdAbsenceModule.CommandTimeout = 300

		Dim prmModuleKey = cmdAbsenceModule.CreateParameter("moduleKey", 200, 1, 50) '200=varchar, 1=input, 50=size
		cmdAbsenceModule.Parameters.Append(prmModuleKey)
		prmModuleKey.value = "ABSENCE"

		Dim prmEnabled = cmdAbsenceModule.CreateParameter("enabled", 11, 2) ' 11=bit, 2=output
		cmdAbsenceModule.Parameters.Append(prmEnabled)

		' Err = 0
		cmdAbsenceModule.Execute()

		iAbsenceEnabled = CInt(cmdAbsenceModule.Parameters("enabled").Value)
		If iAbsenceEnabled < 0 Then
			iAbsenceEnabled = 1
		End If
		cmdAbsenceModule = Nothing
	End If

	Response.Write("<INPUT type='hidden' id=txtAbsenceEnabled name=txtAbsenceEnabled value=" & iAbsenceEnabled & ">")
	Response.Write("<INPUT type='hidden' id=txtCustomReportsGranted name=txtCustomReportsGranted value=""" & fCustomReportsGranted & """>")
	Response.Write("<INPUT type='hidden' id=txtCrossTabsGranted name=txtCrossTabsGranted value=""" & fCrossTabsGranted & """>")
	Response.Write("<INPUT type='hidden' id=txtCalendarReportsGranted name=txtCalendarReportsGranted value=""" & fCalendarReportsGranted & """>")
	Response.Write("<INPUT type='hidden' id=txtMailMergeGranted name=txtMailMergeGranted value=""" & fMailMergeGranted & """>")
	Response.Write("<INPUT type='hidden' id=txtWorkflowGranted name=txtWorkflowGranted value=""" & fWorkflowGranted & """>")
	Response.Write("<INPUT type='hidden' id=txtCalculationsGranted name=txtCalculationsGranted value=""" & fCalculationsGranted & """>")
	Response.Write("<INPUT type='hidden' id=txtFiltersGranted name=txtFiltersGranted value=""" & fFiltersGranted & """>")
	Response.Write("<INPUT type='hidden' id=txtPicklistsGranted name=txtPicklistsGranted value=""" & fPicklistsGranted & """>")
	Response.Write("<INPUT type='hidden' id=txtNewUserGranted name=txtNewUserGranted value=""" & fNewUserGranted & """>")

	Response.Write("<INPUT type='hidden' id=txtErrorDescription name=txtErrorDescription value=""" & sErrorDescription & """>")
%>

<div id="contextmenu" class="accordion">
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
		<h3 id="mnutoolRecord">Record</h3>
		<div>
			<ul id="mnubandRecord"></ul>
		</div>
		<h3 id="mnutoolHistory">History</h3>
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
				<li id="mnutoolDownloadControls"><a href="#">Download Controls</a></li>
			</ul>
		</div>
		<h3 id="mnutoolHelp">Help</h3>
		<div>
			<ul id="mnubandHelp">
				<li id="mnutoolAboutHRPro"><a href="#">About OpenHR</a></li>
				<li id="mnutoolVersionInfo"><a href="#">Version Information</a></li>
				<li class="hidden" id="mnutoolContentsAndIndex"><a href="#">Contents and Index</a></li>
			</ul>
		</div>
</div>




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

	<INPUT type="hidden" id=txtDoneDatabaseMenu name=txtDoneDatabaseMenu value=0>
	<INPUT type="hidden" id=txtDoneQuickEntryMenu name=txtDoneQuickEntryMenu value=0>
	<INPUT type="hidden" id=txtDoneTableScreensMenu name=txtDoneTableScreensMenu value=0>
	<INPUT type="hidden" id=txtDoneSelfServiceStart name=txtDoneSelfServiceStart value=0>

	<INPUT type="hidden" id=txtMenuSaved name=txtMenuSaved value=0>
</FORM>

<FORM action="" method=POST id=frmWorkAreaInfo name=frmWorkAreaInfo>
<INPUT type="hidden" id=txtHRProNavigation name=txtHRProNavigation value=0>
</FORM>

<script type="text/javascript">

	menu_window_onload();

</script>