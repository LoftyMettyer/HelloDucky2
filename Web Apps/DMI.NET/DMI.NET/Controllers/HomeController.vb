Imports System.Web.Mvc
Imports System.IO
Imports System.Web

Namespace Controllers

  Public Class HomeController
    Inherits Controller

    Function Configuration() As ActionResult
      Return View()
    End Function


    <HttpPost()>
    Function passwordChange_Submit(value As FormCollection) As JsonResult

      On Error Resume Next

      Dim sReferringPage = ""
      Dim fSubmitPasswordChange = ""
      Dim sErrorText = ""

      ' Only process the form submission if the referring page was the newUser page.
      ' If it wasn't then redirect to the login page.
      'sReferringPage = Request.ServerVariables("HTTP_REFERER")
      'If InStrRev(sReferringPage, "/") > 0 Then
      '	sReferringPage = Mid(sReferringPage, InStrRev(sReferringPage, "/") + 1)
      'End If

      'If UCase(sReferringPage) <> UCase("passwordChange") Then
      '	Return RedirectToAction("login")
      'Else
      If True Then
        fSubmitPasswordChange = (Len(Request.Form("txtGotoPage")) = 0)

        If fSubmitPasswordChange Then
          ' Force password change only if there are no other users logged in with the same name.
          Dim cmdCheckUserSessions = CreateObject("ADODB.Command")
          cmdCheckUserSessions.CommandText = "spASRGetCurrentUsersCountOnServer"
          cmdCheckUserSessions.CommandType = 4 ' Stored procedure.
          cmdCheckUserSessions.ActiveConnection = Session("databaseConnection")

          Dim prmCount = cmdCheckUserSessions.CreateParameter("count", 3, 2) ' 3=integer, 2=output
          cmdCheckUserSessions.Parameters.Append(prmCount)

          Dim prmUserName = cmdCheckUserSessions.CreateParameter("userName", 200, 1, 8000)   ' 200=varchar, 1=input, 8000=size
          cmdCheckUserSessions.Parameters.Append(prmUserName)
          prmUserName.value = Session("Username")

          Err.Clear()
          cmdCheckUserSessions.Execute()

          Dim iUserSessionCount = CLng(cmdCheckUserSessions.Parameters("count").Value)
          cmdCheckUserSessions = Nothing

          If iUserSessionCount < 2 Then
            ' Read the Password details from the Password form.
            Dim sCurrentPassword = Request.Form("txtCurrentPassword")
            Dim sNewPassword = Request.Form("txtPassword1")

            ' Attempt to change the password on the SQL Server.
            Dim cmdChangePassword = CreateObject("ADODB.Command")
            cmdChangePassword.CommandText = "sp_password"
            cmdChangePassword.CommandType = 4 ' Stored Procedure
            cmdChangePassword.ActiveConnection = Session("databaseConnection")

            Dim prmCurrentPassword = cmdChangePassword.CreateParameter("currentPassword", 200, 1, 255)
            cmdChangePassword.Parameters.Append(prmCurrentPassword)
            If Len(sCurrentPassword) > 0 Then
              prmCurrentPassword.value = sCurrentPassword
            Else
              prmCurrentPassword.value = DBNull.Value
            End If

            Dim prmNewPassword = cmdChangePassword.CreateParameter("newPassword", 200, 1, 255)
            cmdChangePassword.Parameters.Append(prmNewPassword)
            If Len(sNewPassword) > 0 Then
              prmNewPassword.value = sNewPassword
            Else
              prmNewPassword.value = DBNull.Value
            End If

            Err.Clear()
            cmdChangePassword.Execute()

            ' Release the ADO command object.
            cmdChangePassword = Nothing

            If Err.Number <> 0 Then
              Session("ErrorTitle") = "Change Password Page"
              Session("ErrorText") = "You could not change your password because of the following error:<p>" & FormatError(Err.Description)
              Dim data = New JsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = ""}
              Return Json(data, JsonRequestBehavior.AllowGet)
              ' Return RedirectToAction("error", "home")
            Else
              ' Password changed okay. Update the appropriate record in the ASRSysPasswords table.
              Dim cmdPasswordOK = CreateObject("ADODB.Command")
              cmdPasswordOK.CommandText = "sp_ASRIntPasswordOK"
              cmdPasswordOK.CommandType = 4 ' Stored Procedure
              cmdPasswordOK.ActiveConnection = Session("databaseConnection")

              Err.Clear()
              cmdPasswordOK.Execute()
              If Err.Number <> 0 Then
                Session("ErrorTitle") = "Change Password Page"
                Session("ErrorText") = "You could not change your password because of the following error:<p>" & FormatError(Err.Description)
                Dim data1 = New JsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = ""}
                Return Json(data1, JsonRequestBehavior.AllowGet)
                ' Return RedirectToAction("error", "Account")
              End If

              ' Release the ADO command object.
              cmdPasswordOK = Nothing

              ' Close and reopen the connection object.
              Dim conX = Session("databaseConnection")
              Dim sConnString = conX.connectionString

              Dim iPos1 = InStr(UCase(sConnString), UCase(";PWD=" & sCurrentPassword))
              If iPos1 > 0 Then
                conX.close()
                conX = Nothing
                Session("databaseConnection") = ""


                Dim sNewConnString = Left(sConnString, iPos1 + 4) & sNewPassword & Mid(sConnString, iPos1 + 5 + Len(sCurrentPassword))
                ' Open a connection to the database.
                conX = CreateObject("ADODB.Connection")
                conX.open(sNewConnString)

                If Err.Number <> 0 Then
                  Session("ErrorTitle") = "Change Password Page"
                  Session("ErrorText") = "You could not change your password because of the following error:<p>" & FormatError(Err.Description)
                  Dim data1 = New JsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = ""}
                  Return Json(data1, JsonRequestBehavior.AllowGet)
                  ' Return RedirectToAction("error", "Account")
                End If

                Session("databaseConnection") = conX

              End If

              ' Create the cached system tables on the server - Don;t do it in a stored procedure because the #temp will then only be visible to that stored procedure
              Dim cmdCreateCache = CreateObject("ADODB.Command")
              cmdCreateCache.CommandText = "DECLARE @iUserGroupID	integer, " & vbNewLine & _
                              "	@sUserGroupName		sysname, " & vbNewLine & _
                              "	@sActualLoginName	varchar(250) " & vbNewLine & _
                              "-- Get the current user's group ID. " & vbNewLine & _
                              "EXEC spASRIntGetActualUserDetails " & vbNewLine & _
                              "	@sActualLoginName OUTPUT, " & vbNewLine & _
                              "	@sUserGroupName OUTPUT, " & vbNewLine & _
                              "	@iUserGroupID OUTPUT " & vbNewLine & _
                              "-- Create the SysProtects cache table " & vbNewLine & _
                              "IF OBJECT_ID('tempdb..#SysProtects') IS NOT NULL " & vbNewLine & _
                              "	DROP TABLE #SysProtects " & vbNewLine & _
                              "CREATE TABLE #SysProtects(ID int, Action tinyint, Columns varbinary(8000), ProtectType int) " & vbNewLine & _
                              "	INSERT #SysProtects " & vbNewLine & _
                              "	SELECT ID, Action, Columns, ProtectType " & vbNewLine & _
                              "       FROM sysprotects " & vbNewLine & _
                              "       WHERE uid = @iUserGroupID"
              'cmdCreateCache.CommandType = 4 ' Stored Procedure
              cmdCreateCache.ActiveConnection = conX
              cmdCreateCache.execute()
              cmdCreateCache = Nothing

              ' Tell the user that the password was changed okay.
              Session("MessageTitle") = "Change Password Page"
              Session("MessageText") = "Password changed successfully."
              Dim data = New JsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = "main"}
              Return Json(data, JsonRequestBehavior.AllowGet)
              ' Return RedirectToAction("message", "Account")

            End If
          Else
            Session("ErrorTitle") = "Change Password Page"
            sErrorText = "You could not change your password.<p>The account is currently being used by "
            If iUserSessionCount > 2 Then
              sErrorText = sErrorText & iUserSessionCount & " users"
            Else
              sErrorText = sErrorText & "another user"
            End If
            sErrorText = sErrorText & " in the system."
            Session("ErrorText") = sErrorText

            ' Return RedirectToAction("Loginerror", "Account")
          End If
        Else
          ' Save the required table/view and screen IDs in session variables.
          Session("action") = Request.Form("txtAction")
          Session("tableID") = Request.Form("txtGotoTableID")
          Session("viewID") = Request.Form("txtGotoViewID")
          Session("screenID") = Request.Form("txtGotoScreenID")
          Session("orderID") = Request.Form("txtGotoOrderID")
          Session("recordID") = Request.Form("txtGotoRecordID")
          Session("parentTableID") = Request.Form("txtGotoParentTableID")
          Session("parentRecordID") = Request.Form("txtGotoParentRecordID")
          Session("realSource") = Request.Form("txtGotoRealSource")
          Session("filterDef") = Request.Form("txtGotoFilterDef")
          Session("filterSQL") = Request.Form("txtGotoFilterSQL")
          Session("lineage") = Request.Form("txtGotoLineage")
          Session("defseltype") = Request.Form("txtGotoDefSelType")
          Session("utilID") = Request.Form("txtGotoUtilID")
          Session("locateValue") = Request.Form("txtGotoLocateValue")
          Session("firstRecPos") = Request.Form("txtGotoFirstRecPos")
          Session("currentRecCount") = Request.Form("txtGotoCurrentRecCount")
          Session("fromMenu") = Request.Form("txtGotoFromMenu")

          ' Go to the requested page.
          ' Return RedirectToAction(Request.Form("txtGotoPage"))
          Session("txtGotoPage") = Request.Form("txtGotoPage")
        End If
      End If
    End Function

    <HttpPost()>
    Function Configuration_Submit(value As FormCollection)

      On Error Resume Next

      Dim sTemp
      Dim sType = ""
      Dim sControlName
      Dim cmdPrimaryStartMode
      Dim prmSection
      Dim prmKey
      Dim prmUserSetting
      Dim prmValue
      Dim cmdHistoryStartMode

      ' Save the user configuration settings.

      '--------------------------------------------
      ' Save the Primary Record Editing Start Mode.
      '--------------------------------------------
      cmdPrimaryStartMode = CreateObject("ADODB.Command")
      cmdPrimaryStartMode.CommandText = "sp_ASRIntSaveSetting"
      cmdPrimaryStartMode.CommandType = 4 ' Stored procedure.
      cmdPrimaryStartMode.ActiveConnection = Session("databaseConnection")

      prmSection = cmdPrimaryStartMode.CreateParameter("section", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
      cmdPrimaryStartMode.Parameters.Append(prmSection)
      prmSection.value = "RecordEditing"

      prmKey = cmdPrimaryStartMode.CreateParameter("key", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
      cmdPrimaryStartMode.Parameters.Append(prmKey)
      prmKey.value = "Primary"

      prmUserSetting = cmdPrimaryStartMode.CreateParameter("userSetting", 11, 1)  ' 11=bit, 1=input
      cmdPrimaryStartMode.Parameters.Append(prmUserSetting)
      prmUserSetting.value = 1

      prmValue = cmdPrimaryStartMode.CreateParameter("value", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
      cmdPrimaryStartMode.Parameters.Append(prmValue)
      prmValue.value = Request.Form("txtPrimaryStartMode")

      Err.Clear()
      cmdPrimaryStartMode.Execute()
      cmdPrimaryStartMode = Nothing
      Session("PrimaryStartMode") = Request.Form("txtPrimaryStartMode")

      '--------------------------------------------
      ' Save the History Record Editing Start Mode.
      '--------------------------------------------
      cmdHistoryStartMode = CreateObject("ADODB.Command")
      cmdHistoryStartMode.CommandText = "sp_ASRIntSaveSetting"
      cmdHistoryStartMode.CommandType = 4 ' Stored procedure.
      cmdHistoryStartMode.ActiveConnection = Session("databaseConnection")

      prmSection = cmdHistoryStartMode.CreateParameter("section", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
      cmdHistoryStartMode.Parameters.Append(prmSection)
      prmSection.value = "RecordEditing"

      prmKey = cmdHistoryStartMode.CreateParameter("key", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
      cmdHistoryStartMode.Parameters.Append(prmKey)
      prmKey.value = "History"

      prmUserSetting = cmdHistoryStartMode.CreateParameter("userSetting", 11, 1)  ' 11=bit, 1=input
      cmdHistoryStartMode.Parameters.Append(prmUserSetting)
      prmUserSetting.value = 1

      prmValue = cmdHistoryStartMode.CreateParameter("value", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
      cmdHistoryStartMode.Parameters.Append(prmValue)
      prmValue.value = Request.Form("txtHistoryStartMode")

      Err.Clear()
      cmdHistoryStartMode.Execute()
      cmdHistoryStartMode = Nothing
      Session("HistoryStartMode") = Request.Form("txtHistoryStartMode")

      '--------------------------------------------
      ' Save the Lookup Record Editing Start Mode.
      '--------------------------------------------
      Dim cmdLookupStartMode
      cmdLookupStartMode = CreateObject("ADODB.Command")
      cmdLookupStartMode.CommandText = "sp_ASRIntSaveSetting"
      cmdLookupStartMode.CommandType = 4 ' Stored procedure.
      cmdLookupStartMode.ActiveConnection = Session("databaseConnection")

      prmSection = cmdLookupStartMode.CreateParameter("section", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
      cmdLookupStartMode.Parameters.Append(prmSection)
      prmSection.value = "RecordEditing"

      prmKey = cmdLookupStartMode.CreateParameter("key", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
      cmdLookupStartMode.Parameters.Append(prmKey)
      prmKey.value = "LookUp"

      prmUserSetting = cmdLookupStartMode.CreateParameter("userSetting", 11, 1) ' 11=bit, 1=input
      cmdLookupStartMode.Parameters.Append(prmUserSetting)
      prmUserSetting.value = 1

      prmValue = cmdLookupStartMode.CreateParameter("value", 200, 1, 8000)  ' 200=varchar, 1=input, 8000=size
      cmdLookupStartMode.Parameters.Append(prmValue)
      prmValue.value = Request.Form("txtLookupStartMode")

      Err.Clear()
      cmdLookupStartMode.Execute()
      cmdLookupStartMode = Nothing
      Session("LookupStartMode") = Request.Form("txtLookupStartMode")

      '--------------------------------------------
      ' Save the Quick Access Record Editing Start Mode.
      '--------------------------------------------
      Dim cmdQuickAccessStartMode
      cmdQuickAccessStartMode = CreateObject("ADODB.Command")
      cmdQuickAccessStartMode.CommandText = "sp_ASRIntSaveSetting"
      cmdQuickAccessStartMode.CommandType = 4 ' Stored procedure.
      cmdQuickAccessStartMode.ActiveConnection = Session("databaseConnection")

      prmSection = cmdQuickAccessStartMode.CreateParameter("section", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
      cmdQuickAccessStartMode.Parameters.Append(prmSection)
      prmSection.value = "RecordEditing"

      prmKey = cmdQuickAccessStartMode.CreateParameter("key", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
      cmdQuickAccessStartMode.Parameters.Append(prmKey)
      prmKey.value = "QuickAccess"

      prmUserSetting = cmdQuickAccessStartMode.CreateParameter("userSetting", 11, 1) ' 11=bit, 1=input
      cmdQuickAccessStartMode.Parameters.Append(prmUserSetting)
      prmUserSetting.value = 1

      prmValue = cmdQuickAccessStartMode.CreateParameter("value", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
      cmdQuickAccessStartMode.Parameters.Append(prmValue)
      prmValue.value = Request.Form("txtQuickAccessStartMode")

      Err.Clear()
      cmdQuickAccessStartMode.Execute()
      cmdQuickAccessStartMode = Nothing
      Session("QuickAccessStartMode") = Request.Form("txtQuickAccessStartMode")

      '--------------------------------------------
      ' Save the Expression Colour Mode.
      '--------------------------------------------
      Dim cmdExprColourMode
      cmdExprColourMode = CreateObject("ADODB.Command")
      cmdExprColourMode.CommandText = "sp_ASRIntSaveSetting"
      cmdExprColourMode.CommandType = 4 ' Stored procedure.
      cmdExprColourMode.ActiveConnection = Session("databaseConnection")

      prmSection = cmdExprColourMode.CreateParameter("section", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
      cmdExprColourMode.Parameters.Append(prmSection)
      prmSection.value = "ExpressionBuilder"

      prmKey = cmdExprColourMode.CreateParameter("key", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
      cmdExprColourMode.Parameters.Append(prmKey)
      prmKey.value = "ViewColours"

      prmUserSetting = cmdExprColourMode.CreateParameter("userSetting", 11, 1) ' 11=bit, 1=input
      cmdExprColourMode.Parameters.Append(prmUserSetting)
      prmUserSetting.value = 1

      prmValue = cmdExprColourMode.CreateParameter("value", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
      cmdExprColourMode.Parameters.Append(prmValue)
      prmValue.value = Request.Form("txtExprColourMode")

      Err.Clear()
      cmdExprColourMode.Execute()
      cmdExprColourMode = Nothing
      Session("ExprColourMode") = Request.Form("txtExprColourMode")

      '--------------------------------------------
      ' Save the Expression Node Mode.
      '--------------------------------------------
      Dim cmdExprNodeMode
      cmdExprNodeMode = CreateObject("ADODB.Command")
      cmdExprNodeMode.CommandText = "sp_ASRIntSaveSetting"
      cmdExprNodeMode.CommandType = 4 ' Stored procedure.
      cmdExprNodeMode.ActiveConnection = Session("databaseConnection")

      prmSection = cmdExprNodeMode.CreateParameter("section", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
      cmdExprNodeMode.Parameters.Append(prmSection)
      prmSection.value = "ExpressionBuilder"

      prmKey = cmdExprNodeMode.CreateParameter("key", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
      cmdExprNodeMode.Parameters.Append(prmKey)
      prmKey.value = "NodeSize"

      prmUserSetting = cmdExprNodeMode.CreateParameter("userSetting", 11, 1) ' 11=bit, 1=input
      cmdExprNodeMode.Parameters.Append(prmUserSetting)
      prmUserSetting.value = 1

      prmValue = cmdExprNodeMode.CreateParameter("value", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
      cmdExprNodeMode.Parameters.Append(prmValue)
      prmValue.value = Request.Form("txtExprNodeMode")

      Err.Clear()
      cmdExprNodeMode.Execute()
      cmdExprNodeMode = Nothing
      Session("ExprNodeMode") = Request.Form("txtExprNodeMode")

      '--------------------------------------------
      ' Save the Find Window Block Size.
      '--------------------------------------------
      Dim cmdFindSize
      cmdFindSize = CreateObject("ADODB.Command")
      cmdFindSize.CommandText = "sp_ASRIntSaveSetting"
      cmdFindSize.CommandType = 4 ' Stored procedure.
      cmdFindSize.ActiveConnection = Session("databaseConnection")

      prmSection = cmdFindSize.CreateParameter("section", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
      cmdFindSize.Parameters.Append(prmSection)
      prmSection.value = "IntranetFindWindow"

      prmKey = cmdFindSize.CreateParameter("key", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
      cmdFindSize.Parameters.Append(prmKey)
      prmKey.value = "BlockSize"

      prmUserSetting = cmdFindSize.CreateParameter("userSetting", 11, 1) ' 11=bit, 1=input
      cmdFindSize.Parameters.Append(prmUserSetting)
      prmUserSetting.value = 1

      prmValue = cmdFindSize.CreateParameter("value", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
      cmdFindSize.Parameters.Append(prmValue)
      prmValue.value = Request.Form("txtFindSize")

      Err.Clear()
      cmdFindSize.Execute()
      cmdFindSize = Nothing
      Session("FindRecords") = Request.Form("txtFindSize")

      '--------------------------------------------
      ' Save the DefSel 'only mine' settings.
      '--------------------------------------------
      For i = 0 To 20
        Select Case i
          Case 0
            sType = "BatchJobs"
          Case 1
            sType = "Calculations"
          Case 2
            sType = "CrossTabs"
          Case 3
            sType = "CustomReports"
          Case 4
            sType = "DataTransfer"
          Case 5
            sType = "Export"
          Case 6
            sType = "Filters"
          Case 7
            sType = "GlobalAdd"
          Case 8
            sType = "GlobalUpdate"
          Case 9
            sType = "GlobalDelete"
          Case 10
            sType = "Import"
          Case 11
            sType = "MailMerge"
          Case 12
            sType = "Picklists"
          Case 13
            sType = "CalendarReports"
          Case 14
            sType = "Labels"
          Case 15
            sType = "LabelDefinition"
          Case 16
            sType = "MatchReports"
          Case 17
            sType = "CareerProgression"
          Case 18
            sType = "EmailGroups"
          Case 19
            sType = "RecordProfile"
          Case 20
            sType = "SuccessionPlanning"
        End Select

        sControlName = "txtOwner_" & sType
        sTemp = "onlymine " & sType

        Dim cmdDefSelOnlyMine
        cmdDefSelOnlyMine = CreateObject("ADODB.Command")
        cmdDefSelOnlyMine.CommandText = "sp_ASRIntSaveSetting"
        cmdDefSelOnlyMine.CommandType = 4 ' Stored procedure.
        cmdDefSelOnlyMine.ActiveConnection = Session("databaseConnection")

        prmSection = cmdDefSelOnlyMine.CreateParameter("section", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
        cmdDefSelOnlyMine.Parameters.Append(prmSection)
        prmSection.value = "defsel"

        prmKey = cmdDefSelOnlyMine.CreateParameter("key", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
        cmdDefSelOnlyMine.Parameters.Append(prmKey)
        prmKey.value = sTemp

        prmUserSetting = cmdDefSelOnlyMine.CreateParameter("userSetting", 11, 1) ' 11=bit, 1=input
        cmdDefSelOnlyMine.Parameters.Append(prmUserSetting)
        prmUserSetting.value = 1

        prmValue = cmdDefSelOnlyMine.CreateParameter("value", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
        cmdDefSelOnlyMine.Parameters.Append(prmValue)
        prmValue.value = Request.Form(sControlName)

        Err.Clear()
        cmdDefSelOnlyMine.Execute()
        cmdDefSelOnlyMine = Nothing
      Next

      '--------------------------------------------
      ' Save the Utility Warning settings.
      '--------------------------------------------
      For i = 0 To 4
        Select Case i
          Case 0
            sType = "DataTransfer"
          Case 1
            sType = "GlobalAdd"
          Case 2
            sType = "GlobalUpdate"
          Case 3
            sType = "GlobalDelete"
          Case 4
            sType = "Import"
        End Select

        sControlName = "txtWarn_" & sType
        sTemp = "warning " & sType

        Dim cmdDefSelOnlyMine
        cmdDefSelOnlyMine = CreateObject("ADODB.Command")
        cmdDefSelOnlyMine.CommandText = "sp_ASRIntSaveSetting"
        cmdDefSelOnlyMine.CommandType = 4 ' Stored procedure.
        cmdDefSelOnlyMine.ActiveConnection = Session("databaseConnection")

        prmSection = cmdDefSelOnlyMine.CreateParameter("section", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
        cmdDefSelOnlyMine.Parameters.Append(prmSection)
        prmSection.value = "warningmsg"

        prmKey = cmdDefSelOnlyMine.CreateParameter("key", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
        cmdDefSelOnlyMine.Parameters.Append(prmKey)
        prmKey.value = sTemp

        prmUserSetting = cmdDefSelOnlyMine.CreateParameter("userSetting", 11, 1) ' 11=bit, 1=input
        cmdDefSelOnlyMine.Parameters.Append(prmUserSetting)
        prmUserSetting.value = 1

        prmValue = cmdDefSelOnlyMine.CreateParameter("value", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
        cmdDefSelOnlyMine.Parameters.Append(prmValue)
        prmValue.value = Request.Form(sControlName)

        Err.Clear()
        cmdDefSelOnlyMine.Execute()
        cmdDefSelOnlyMine = Nothing
      Next

      '--------------------------------------------
      ' Redirect to the save confirmation page.
      '--------------------------------------------
      Session("confirmtext") = "User Configuration has been saved successfully."
      Session("confirmtitle") = "User Configuration"
      Session("followpage") = "default"
      Session("reaction") = Request.Form("txtReaction")

      '    Return RedirectToAction("confirmok")

    End Function

    Function ConfirmOK() As ActionResult
      Return View()
    End Function

    ' GET: /Home
    Function Main() As ActionResult
      Return View()
    End Function

    Function Find() As ActionResult
      Return View()
    End Function

    <HttpPost()>
    Function default_Submit()

      ' Save the required table/view and screen IDs in session variables.
      Session("action") = Request.Form("txtAction")
      Session("tableID") = Request.Form("txtGotoTableID")
      Session("viewID") = Request.Form("txtGotoViewID")
      Session("screenID") = Request.Form("txtGotoScreenID")
      Session("orderID") = Request.Form("txtGotoOrderID")
      Session("recordID") = Request.Form("txtGotoRecordID")
      Session("parentTableID") = Request.Form("txtGotoParentTableID")
      Session("parentRecordID") = Request.Form("txtGotoParentRecordID")
      Session("realSource") = Request.Form("txtGotoRealSource")
      Session("filterDef") = Request.Form("txtGotoFilterDef")
      Session("filterSQL") = Request.Form("txtGotoFilterSQL")
      Session("lineage") = Request.Form("txtGotoLineage")
      Session("defseltype") = Request.Form("txtGotoDefSelType")
      Session("utilID") = Request.Form("txtGotoUtilID")
      Session("locateValue") = Request.Form("txtGotoLocateValue")
      Session("firstRecPos") = Request.Form("txtGotoFirstRecPos")
      Session("currentRecCount") = Request.Form("txtGotoCurrentRecCount")
      Session("fromMenu") = Request.Form("txtGotoFromMenu")
      Session("reset") = Request.Form("txtReset")

      Session("reloadMenu") = Request.Form("txtReloadMenu")

      Session("StandardReport_Type") = Request.Form("txtStandardReportType")
      Session("optionRecordID") = "0"
      Session("optionAction") = ""

      ' Go to the requested page.
      Return RedirectToAction(Request.Form("txtGotoPage").Replace(".asp", ""))

    End Function


    <HttpPost()>
    Function emptyoption_Submit()

      On Error Resume Next

      ' Save the required information in session variables.
      Session("optionScreenID") = Request.Form("txtGotoOptionScreenID")
      Session("optionTableID") = Request.Form("txtGotoOptionTableID")
      Session("optionViewID") = Request.Form("txtGotoOptionViewID")
      Session("optionOrderID") = Request.Form("txtGotoOptionOrderID")
      Session("optionRecordID") = Request.Form("txtGotoOptionRecordID")
      Session("optionFilterDef") = Request.Form("txtGotoOptionFilterDef")
      Session("optionFilterSQL") = Request.Form("txtGotoOptionFilterSQL")
      Session("optionValue") = Request.Form("txtGotoOptionValue")
      Session("optionLinkTableID") = Request.Form("txtGotoOptionLinkTableID")
      Session("optionLinkOrderID") = Request.Form("txtGotoOptionLinkOrderID")
      Session("optionLinkViewID") = Request.Form("txtGotoOptionLinkViewID")
      Session("optionLinkRecordID") = Request.Form("txtGotoOptionLinkRecordID")
      Session("optionColumnID") = Request.Form("txtGotoOptionColumnID")
      Session("optionLookupColumnID") = Request.Form("txtGotoOptionLookupColumnID")
      Session("optionLookupMandatory") = Request.Form("txtGotoOptionLookupMandatory")
      Session("optionLookupValue") = Request.Form("txtGotoOptionLookupValue")
      Session("optionLookupFilterValue") = Request.Form("txtGotoOptionLookupFilterValue")
      Session("optionFile") = Request.Form("txtGotoOptionFile")
      Session("optionExtension") = Request.Form("txtGotoOptionExtension")
      'Session("optionOLEOnServer") = Request.Form("txtGotoOptionOLEOnServer")
      Session("optionAction") = Request.Form("txtGotoOptionAction")
      Session("optionPageAction") = Request.Form("txtGotoOptionPageAction")
      Session("optionCourseTitle") = Request.Form("txtGotoOptionCourseTitle")
      Session("optionFirstRecPos") = Request.Form("txtGotoOptionFirstRecPos")
      Session("optionCurrentRecCount") = Request.Form("txtGotoOptionCurrentRecCount")
      Session("optionExprType") = Request.Form("txtGotoOptionExprType")
      Session("optionExprID") = Request.Form("txtGotoOptionExprID")
      Session("optionFunctionID") = Request.Form("txtGotoOptionFunctionID")
      Session("optionParameterIndex") = Request.Form("txtGotoOptionParameterIndex")
      Session("OptionRealsource") = Request.Form("txtGotoOptionRealsource")
      Session("StandardReport_Type") = Request.Form("txtStandardReportType")
      Session("optionDefSelType") = Request.Form("txtGotoOptionDefSelType")
      Session("optionDefSelRecordID") = Request.Form("txtGotoOptionDefSelRecordID")
      Session("optionOLEType") = Request.Form("txtGotoOptionOLEType")
      Session("optionOLEMaxEmbedSize") = Request.Form("txtGotoOptionOLEMaxEmbedSize")
      Session("optionOLEReadOnly") = Request.Form("txtGotoOptionOLEReadOnly")
      Session("optionOnlyNumerics") = CLng(Request.Form("txtOptionOnlyNumerics"))

      ' Go to the requested page.
      Return RedirectToAction(Request.Form("txtGotoOptionPage"))

    End Function

    Function DefSel() As ActionResult
      Return View()
    End Function

    <HttpPost()>
    Function DefSel(value As FormCollection)
      Return View()
    End Function

    <HttpPost()>
    Function DefSel_Submit(value As FormCollection)
      ' Set some session variables used by all the util pages
      Session("utiltype") = Request.Form("utiltype")
      Session("utilid") = Request.Form("utilid")
      Session("utilname") = Request.Form("utilname")
      Session("action") = Request.Form("action")
      Session("utiltableid") = Request.Form("txtTableID")

      ' Now examine what we are doing and redirect as appropriate
      If (Session("action") = "new") Or _
       (Session("action") = "edit") Or _
       (Session("action") = "view") Or _
       (Session("action") = "copy") Then
        Select Case Session("utiltype")
          Case 1 ' CROSS TABS
            Return RedirectToAction("util_def_crosstabs")
          Case 2 ' CUSTOM REPORTS
            Return RedirectToAction("util_def_customreports")
          Case 9 ' MAIL MERGE
            Return RedirectToAction("util_def_mailmerge")
          Case 10 ' PICKLISTS
            Return RedirectToAction("util_def_picklist")
          Case 11 ' FILTERS
            Return RedirectToAction("util_def_expression")
          Case 12 ' CALCULATIONS
            Return RedirectToAction("util_def_expression")
          Case 17 ' CALENDAR REPORTS
            Return RedirectToAction("util_def_calendarreport")
        End Select

      ElseIf Session("action") = "delete" Then
        Select Case Session("utiltype")
          Case 1  ' CROSS TABS
            Session("reaction") = "CROSSTABS"
          Case 2  ' CUSTOM REPORTS
            Session("reaction") = "CUSTOMREPORTS"
          Case 9  ' MAIL MERGE
            Session("reaction") = "MAILMERGE"
          Case 10 ' PICKLISTS
            Session("reaction") = "PICKLISTS"
          Case 11 ' FILTERS
            Session("reaction") = "FILTERS"
          Case 12 ' CALCULATIONS
            Session("reaction") = "CALCULATIONS"
          Case 17 ' CALENDAR REPORTS
            Session("reaction") = "CALENDARREPORTS"
        End Select
        Return RedirectToAction("checkforusage")
      End If

    End Function

    Function DefSelProperties() As ActionResult
      Return View()
    End Function

    Function Util_Def_CustomReports() As ActionResult
      Return View()
    End Function

    Function util_def_crosstabs() As ActionResult
      Return View()
    End Function

    Function CheckForUsage() As ActionResult
      Return View()
    End Function

    Function Data() As ActionResult
      Return View()
    End Function

    Function OptionData() As ActionResult
      Return View()
    End Function

    Function optionData_Submit() As ActionResult

      On Error Resume Next

      ' Read the information from the calling form.
      Session("optionAction") = Request.Form("txtOptionAction")
      Session("optionTableID") = Request.Form("txtOptionTableID")
      Session("optionViewID") = Request.Form("txtOptionViewID")
      Session("optionOrderID") = Request.Form("txtOptionOrderID")
      Session("optionColumnID") = Request.Form("txtOptionColumnID")
      Session("optionPageAction") = Request.Form("txtOptionPageAction")
      Session("optionFirstRecPos") = Request.Form("txtOptionFirstRecPos")
      Session("optionCurrentRecCount") = Request.Form("txtOptionCurrentRecCount")
      Session("optionLocateValue") = Request.Form("txtGotoLocateValue")
      Session("optionCourseTitle") = Request.Form("txtOptionCourseTitle")
      Session("optionRecordID") = Request.Form("txtOptionRecordID")
      Session("optionLinkRecordID") = Request.Form("txtOptionLinkRecordID")
      Session("optionValue") = Request.Form("txtOptionValue")
      Session("optionSQL") = Request.Form("txtOptionSQL")
      Session("optionPromptSQL") = Request.Form("txtOptionPromptSQL")
      Session("optionOnlyNumerics") = CLng(Request.Form("txtOptionOnlyNumerics"))
      Session("optionLookupColumnID") = Request.Form("txtOptionLookupColumnID")
      Session("optionFilterValue") = Request.Form("txtOptionLookupFilterValue")
      Session("IsLookupTable") = Request.Form("txtOptionIsLookupTable")
      Session("optionParentTableID") = Request.Form("txtOptionParentTableID")
      Session("optionParentRecordID") = Request.Form("txtOptionParentRecordID")
      Session("option1000SepCols") = Request.Form("txtOption1000SepCols")

      ' Go to the requested page.
      Return RedirectToAction("OptionData")

    End Function

    Function Data_Submit() As ActionResult

      On Error Resume Next

      Const DEADLOCK_ERRORNUMBER = -2147467259
      Const DEADLOCK_MESSAGESTART = "YOUR TRANSACTION (PROCESS ID #"
      Const DEADLOCK_MESSAGEEND = ") WAS DEADLOCKED WITH ANOTHER PROCESS AND HAS BEEN CHOSEN AS THE DEADLOCK VICTIM. RERUN YOUR TRANSACTION."
      Const DEADLOCK2_MESSAGESTART = "TRANSACTION (PROCESS ID "
      Const DEADLOCK2_MESSAGEEND = ") WAS DEADLOCKED ON "
      Const SQLMAILNOTSTARTEDMESSAGE = "SQL MAIL SESSION IS NOT STARTED."

      Dim iRETRIES = 5
      Dim iRetryCount = 0
      Dim sErrorMsg = "", sErrMsg = ""
      Dim fWarning = False
      Dim fOk = False
      Dim fTBOverride = False

      ' Read the information from the calling form.
      Dim sRealSource = Request.Form("txtRealSource")
      Dim lngTableID = Request.Form("txtCurrentTableID")
      Dim lngScreenID = Request.Form("txtCurrentScreenID")
      Dim lngViewID = Request.Form("txtCurrentViewID")
      Dim lngRecordID = Request.Form("txtRecordID")
      Dim sAction = Request.Form("txtAction")
      Dim sReaction = Request.Form("txtReaction")
      Dim sInsertUpdateDef = Request.Form("txtInsertUpdateDef")
      Dim iTimestamp = Request.Form("txtTimestamp")
      Dim iTBEmployeeRecordID = Request.Form("txtTBEmployeeRecordID")
      Dim iTBCourseRecordID = Request.Form("txtTBCourseRecordID")
      Dim sTBBookingStatusValue = Request.Form("txtTBBookingStatusValue")
      Dim fUserChoice = Request.Form("txtUserChoice")

      If Request.Form("txtTBOverride") = "" Then
        fTBOverride = False
      Else
        fTBOverride = CBool(Request.Form("txtTBOverride"))
      End If

      If sAction = "SAVE" Then
        Dim sTBErrorMsg = ""
        Dim sTBWarningMsg = ""
        Dim iTBResultCode = 0
        Dim sCode = ""

        If (Not fTBOverride) And (CLng(lngTableID) = CLng(Session("TB_TBTableID"))) Then
          ' Training Booking check.
          Dim cmdTBCheck = CreateObject("ADODB.Command")
          cmdTBCheck.CommandText = "sp_ASRIntValidateTrainingBooking"
          cmdTBCheck.CommandType = 4    ' Stored procedure
          cmdTBCheck.ActiveConnection = Session("databaseConnection")

          Dim prmResult = cmdTBCheck.CreateParameter("resultCode", 3, 2)    ' 3=integer, 2=output
          cmdTBCheck.Parameters.Append(prmResult)

          Dim prmTBEmployeeRecordID = cmdTBCheck.CreateParameter("empRecID", 3, 1)  '3=integer, 1=input
          cmdTBCheck.Parameters.Append(prmTBEmployeeRecordID)
          prmTBEmployeeRecordID.value = CleanNumeric(iTBEmployeeRecordID)

          Dim prmTBCourseRecordID = cmdTBCheck.CreateParameter("courseRecID", 3, 1) '3=integer, 1=input
          cmdTBCheck.Parameters.Append(prmTBCourseRecordID)
          prmTBCourseRecordID.value = CleanNumeric(iTBCourseRecordID)

          Dim prmTBStatus = cmdTBCheck.CreateParameter("tbStatus", 200, 1, 8000) '200=varchar, 1=input, 8000=size
          cmdTBCheck.Parameters.Append(prmTBStatus)
          prmTBStatus.value = sTBBookingStatusValue

          Dim prmTBRecordID = cmdTBCheck.CreateParameter("tbRecID", 3, 1) '3=integer, 1=input
          cmdTBCheck.Parameters.Append(prmTBRecordID)
          prmTBRecordID.value = CleanNumeric(lngRecordID)

          Err.Clear()
          cmdTBCheck.Execute()
          If (Err.Number <> 0) Then
            sErrorMsg = "Error validating training booking." & vbCrLf & FormatError(Err.Description)
          End If

          If Len(sErrorMsg) = 0 Then
            iTBResultCode = cmdTBCheck.Parameters("resultCode").Value
          End If
          cmdTBCheck = Nothing

          If Len(sErrorMsg) = 0 Then
            If iTBResultCode > 0 Then
              Dim sTBResultCode = CStr(iTBResultCode)
              If Len(sTBResultCode) = 4 Then
                ' Get the overbooking check code.
                sCode = Left(sTBResultCode, 1)
                If sCode = "1" Then
                  sTBErrorMsg = "The course is already fully booked. Unable to make the booking."
                Else
                  If sCode = "2" Then
                    sTBWarningMsg = "The course is already fully booked. Unable to make the booking."
                  End If
                End If
              End If

              If Len(sTBResultCode) >= 3 Then
                ' Get the pre-requisite check code.
                sCode = Mid(sTBResultCode, Len(sTBResultCode) - 2, 1)
                If sCode = "1" Then
                  If Len(sTBErrorMsg) > 0 Then
                    sTBErrorMsg = sTBErrorMsg & vbCrLf
                  End If
                  sTBErrorMsg = sTBErrorMsg & "The delegate has not met the pre-requisites for the course. Unable to make the booking."
                Else
                  If sCode = "2" Then
                    If Len(sTBWarningMsg) > 0 Then
                      sTBWarningMsg = sTBWarningMsg & vbCrLf
                    End If
                    sTBWarningMsg = sTBWarningMsg & "The delegate has not met the pre-requisites for the course."
                  End If
                End If
              End If

              If Len(sTBResultCode) >= 2 Then
                ' Get the availability check code.
                sCode = Mid(sTBResultCode, Len(sTBResultCode) - 1, 1)
                If sCode = "1" Then
                  If Len(sTBErrorMsg) > 0 Then
                    sTBErrorMsg = sTBErrorMsg & vbCrLf
                  End If
                  sTBErrorMsg = sTBErrorMsg & "The delegate is unavailable for the course."
                Else
                  If sCode = "2" Then
                    If Len(sTBWarningMsg) > 0 Then
                      sTBWarningMsg = sTBWarningMsg & vbCrLf
                    End If
                    sTBWarningMsg = sTBWarningMsg & "The delegate is unavailable for the course."
                  End If
                End If
              End If

              If Len(sTBResultCode) >= 1 Then
                ' Get the Overlapped Booking check code.
                sCode = Mid(sTBResultCode, Len(sTBResultCode), 1)
                If sCode = "1" Then
                  If Len(sTBErrorMsg) > 0 Then
                    sTBErrorMsg = sTBErrorMsg & vbCrLf
                  End If
                  sTBErrorMsg = sTBErrorMsg & "The delegate is already booked on a course that overlaps with this course. Unable to make the booking."
                Else
                  If sCode = "2" Then
                    If Len(sTBWarningMsg) > 0 Then
                      sTBWarningMsg = sTBWarningMsg & vbCrLf
                    End If
                    sTBWarningMsg = sTBWarningMsg & "The delegate is already booked on a course that overlaps with this course. Unable to make the booking."
                  End If
                End If
              End If
            End If
          End If
        End If

        If Len(sTBErrorMsg) > 0 Then
          ' Training Booking validation failure.	
          sErrorMsg = sTBErrorMsg
          sAction = "SAVEERROR"
        Else
          If Len(sTBWarningMsg) > 0 Then
            sErrorMsg = sTBWarningMsg
            sAction = sReaction
            fWarning = True
          Else
            ' Check if we're inserting or updating.
            If lngRecordID = 0 Then
              ' Inserting.

              ' The required stored procedure exists, so run it.
              Dim cmdInsertRecord = CreateObject("ADODB.Command")
              cmdInsertRecord.CommandText = "spASRIntInsertNewRecord"
              cmdInsertRecord.CommandType = 4 ' Stored procedure
              cmdInsertRecord.CommandTimeout = 180
              cmdInsertRecord.ActiveConnection = Session("databaseConnection")

              Dim prmNewID = cmdInsertRecord.CreateParameter("newID", 3, 2)
              cmdInsertRecord.Parameters.Append(prmNewID)

              Dim prmInsertSQL = cmdInsertRecord.CreateParameter("insertSQL", 201, 1, 2147483646)
              cmdInsertRecord.Parameters.Append(prmInsertSQL)
              prmInsertSQL.value = sInsertUpdateDef

              Dim fDeadlock = True
              Do While fDeadlock
                fDeadlock = False

                cmdInsertRecord.ActiveConnection.Errors.Clear()

                ' Run the insert stored procedure.
                cmdInsertRecord.Execute()

                If cmdInsertRecord.ActiveConnection.Errors.Count > 0 Then
                  For iLoop = 1 To cmdInsertRecord.ActiveConnection.Errors.Count
                    sErrMsg = FormatError(cmdInsertRecord.ActiveConnection.Errors.Item(iLoop - 1).Description)

                    If (cmdInsertRecord.ActiveConnection.Errors.Item(iLoop - 1).Number = DEADLOCK_ERRORNUMBER) And _
                     (((UCase(Left(sErrMsg, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
                    (UCase(Right(sErrMsg, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
                     ((UCase(Left(sErrMsg, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
                      (InStr(UCase(sErrMsg), DEADLOCK2_MESSAGEEND) > 0))) Then
                      ' The error is for a deadlock.
                      ' Sorry about having to use the err.description to trap the error but the err.number
                      ' is not specific and MSDN suggests using the err.description.
                      If (iRetryCount < iRETRIES) And (cmdInsertRecord.ActiveConnection.Errors.Count = 1) Then
                        iRetryCount = iRetryCount + 1
                        fDeadlock = True
                      Else
                        If Len(sErrorMsg) > 0 Then
                          sErrorMsg = sErrorMsg & vbCrLf
                        End If
                        sErrorMsg = sErrorMsg & "Another user is deadlocking the database. Try saving again."
                        fOk = False
                      End If
                    ElseIf UCase(cmdInsertRecord.ActiveConnection.Errors.Item(iLoop - 1).Description) = SQLMAILNOTSTARTEDMESSAGE Then
                      '"SQL Mail session is not started."
                      'Ignore this error
                      'ElseIf (cmdInsertRecord.ActiveConnection.Errors.Item(iloop - 1).Number = XP_SENDMAIL_ERRORNUMBER) And _
                      '	(UCase(Left(cmdInsertRecord.ActiveConnection.Errors.Item(iloop - 1).Description, Len(XP_SENDMAIL_MESSAGE))) = XP_SENDMAIL_MESSAGE) Then
                      '"EXECUTE permission denied on object 'xp_sendmail'"
                      'Ignore this error
                    ElseIf cmdInsertRecord.ActiveConnection.Errors.Item(iLoop - 1).nativeerror = 3609 Then
                      ' Ignore the follow on message that says "The transaction ended in the trigger."
                    Else
                      'NHRD 18082011 HRPRO-1572 Removed extra carriage return for this error msg
                      'sErrorMsg = sErrorMsg & vbcrlf & _
                      sErrorMsg = sErrorMsg & _
                       FormatError(cmdInsertRecord.ActiveConnection.Errors.Item(iLoop - 1).Description)
                      fOk = False
                    End If
                  Next

                  cmdInsertRecord.ActiveConnection.Errors.Clear()

                  If Not fOk Then
                    'JPD 20110705 HRPRO-1572
                    ' Now get validation failure message prefixed woth <record description> and <line of hyphens>.
                    ' Only add extra carriage return if required (ie. if there is a record description).
                    Dim sRecDescExists = ""
                    If Mid(sErrorMsg, 3, 5) <> "-----" Then
                      sRecDescExists = vbCrLf
                    End If

                    sErrorMsg = "The new record could not be created." & sRecDescExists & sErrorMsg
                    sAction = "SAVEERROR"
                  End If
                Else
                  lngRecordID = cmdInsertRecord.Parameters("newID").Value

                  If Len(sReaction) > 0 Then
                    sAction = sReaction
                  Else
                    sAction = "LOAD"
                  End If
                End If
              Loop
              cmdInsertRecord = Nothing


              'MH20001017 Immediate email stuff to go in v1.9.0
              Dim cmdInsertRecord2 = CreateObject("ADODB.Command")
              cmdInsertRecord2.CommandText = "spASREmailImmediate"
              cmdInsertRecord2.CommandType = 4    ' Stored procedure
              cmdInsertRecord2.CommandTimeout = 180
              cmdInsertRecord2.ActiveConnection = Session("databaseConnection")

              Dim prmInsertSQL2 = cmdInsertRecord2.CreateParameter("Username", 200, 1, 255)   ' 200=varchar, 1=input, 255=size
              cmdInsertRecord2.Parameters.Append(prmInsertSQL2)
              prmInsertSQL2.value = Session("Username")

              Err.Clear()
              cmdInsertRecord2.Execute()
              cmdInsertRecord2 = Nothing
            Else
              ' Updating.

              ' The required stored procedure exists, so run it.
              Dim cmdUpdateRecord = CreateObject("ADODB.Command")
              cmdUpdateRecord.CommandText = "spASRIntUpdateRecord"
              cmdUpdateRecord.CommandType = 4 ' Stored procedure
              cmdUpdateRecord.CommandTimeout = 180
              cmdUpdateRecord.ActiveConnection = Session("databaseConnection")

              Dim prmResultCode = cmdUpdateRecord.CreateParameter("resultCode", 3, 2) ' 3=integer, 2=output
              cmdUpdateRecord.Parameters.Append(prmResultCode)

              Dim prmUpdateSQL = cmdUpdateRecord.CreateParameter("updateSQL", 201, 1, 2147483646)
              cmdUpdateRecord.Parameters.Append(prmUpdateSQL)
              prmUpdateSQL.value = sInsertUpdateDef

              Dim prmTableID = cmdUpdateRecord.CreateParameter("tableID", 3, 1)
              cmdUpdateRecord.Parameters.Append(prmTableID)
              prmTableID.value = CLng(CleanNumeric(lngTableID))

              Dim prmRealSource = cmdUpdateRecord.CreateParameter("realSource", 200, 1, 255)
              cmdUpdateRecord.Parameters.Append(prmRealSource)
              prmRealSource.value = sRealSource

              Dim prmID = cmdUpdateRecord.CreateParameter("id", 3, 1) ' 3=integer, 1=input
              cmdUpdateRecord.Parameters.Append(prmID)
              prmID.value = CleanNumeric(lngRecordID)

              Dim prmTimestamp = cmdUpdateRecord.CreateParameter("timestamp", 3, 1) ' 3=integer, 1=input
              cmdUpdateRecord.Parameters.Append(prmTimestamp)
              prmTimestamp.value = CleanNumeric(iTimestamp)

              Dim fDeadlock = True
              Do While fDeadlock
                fDeadlock = False

                cmdUpdateRecord.ActiveConnection.Errors.Clear()

                ' Run the update stored procedure.
                cmdUpdateRecord.Execute()

                If cmdUpdateRecord.ActiveConnection.Errors.Count > 0 Then
                  For iLoop = 1 To cmdUpdateRecord.ActiveConnection.Errors.Count
                    sErrMsg = FormatError(cmdUpdateRecord.ActiveConnection.Errors.Item(iLoop - 1).Description)

                    If (cmdUpdateRecord.ActiveConnection.Errors.Item(iLoop - 1).Number = DEADLOCK_ERRORNUMBER) And _
                     (((UCase(Left(sErrMsg, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
                      (UCase(Right(sErrMsg, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
                     ((UCase(Left(sErrMsg, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
                     (InStr(UCase(sErrMsg), DEADLOCK2_MESSAGEEND) > 0))) Then
                      ' The error is for a deadlock.
                      ' Sorry about having to use the err.description to trap the error but the err.number
                      ' is not specific and MSDN suggests using the err.description.
                      If (iRetryCount < iRETRIES) And (cmdUpdateRecord.ActiveConnection.Errors.Count = 1) Then
                        iRetryCount = iRetryCount + 1
                        fDeadlock = True
                      Else
                        If Len(sErrorMsg) > 0 Then
                          sErrorMsg = sErrorMsg & vbCrLf
                        End If
                        sErrorMsg = sErrorMsg & "Another user is deadlocking the database. Try saving again."
                        fOk = False
                      End If
                    ElseIf UCase(cmdUpdateRecord.ActiveConnection.Errors.Item(iLoop - 1).Description) = SQLMAILNOTSTARTEDMESSAGE Then
                      '"SQL Mail session is not started."
                      'Ignore this error
                    ElseIf cmdUpdateRecord.ActiveConnection.Errors.Item(iLoop - 1).nativeerror = 3609 Then
                      ' Ignore the follow on message that says "The transaction ended in the trigger."
                    Else
                      sErrorMsg = sErrorMsg & vbCrLf & _
                       FormatError(cmdUpdateRecord.ActiveConnection.Errors.Item(iLoop - 1).Description)
                      fOk = False
                    End If
                  Next

                  cmdUpdateRecord.ActiveConnection.Errors.Clear()

                  If Not fOk Then
                    'JPD 20110705 HRPRO-1572
                    ' Now get validation failure message prefixed with <record description> and <line of hyphens>.
                    ' Only add extra carriage return if required (ie. if there is a record description).
                    Dim sRecDescExists = ""
                    If Mid(sErrorMsg, 3, 5) <> "-----" Then
                      sRecDescExists = vbCrLf
                    End If

                    sErrorMsg = "The record could not be updated." & sRecDescExists & sErrorMsg
                    sAction = "SAVEERROR"
                  End If
                Else
                  Select Case cmdUpdateRecord.Parameters("resultCode").Value
                    Case 1 ' Record changed by another user, and still in the current table/view.
                      sErrorMsg = "The record has been amended by another user and will be refreshed."
                    Case 2 ' Record changed by another user, and is no longer in the current table/view.
                      sErrorMsg = "The record has been amended by another user and is no longer in the current view."
                    Case 3 ' Record deleted by another user.
                      sErrorMsg = "The record has been deleted by another user."
                  End Select

                  If Len(sReaction) > 0 Then
                    sAction = sReaction
                  Else
                    sAction = "LOAD"
                  End If
                End If
              Loop
              cmdUpdateRecord = Nothing

              'MH20001017 Immediate email stuff to go in v1.9.0
              cmdUpdateRecord = CreateObject("ADODB.Command")
              cmdUpdateRecord.CommandText = "spASREmailImmediate"
              cmdUpdateRecord.CommandType = 4 ' Stored procedure
              cmdUpdateRecord.CommandTimeout = 180
              cmdUpdateRecord.ActiveConnection = Session("databaseConnection")

              prmUpdateSQL = cmdUpdateRecord.CreateParameter("Username", 200, 1, 255) ' 200=varchar, 1=input, 255=size
              cmdUpdateRecord.Parameters.Append(prmUpdateSQL)
              prmUpdateSQL.value = Session("Username")

              Err.Clear()
              cmdUpdateRecord.Execute()
              cmdUpdateRecord = Nothing
            End If
          End If
        End If
      ElseIf sAction = "DELETE" Then
        ' Deleting.

        ' The required stored procedure exists, so run it.
        Dim cmdDeleteRecord = CreateObject("ADODB.Command")
        cmdDeleteRecord.CommandText = "sp_ASRDeleteRecord"
        cmdDeleteRecord.CommandType = 4 ' Stored procedure
        cmdDeleteRecord.ActiveConnection = Session("databaseConnection")

        Dim prmResultCode = cmdDeleteRecord.CreateParameter("resultCode", 3, 2)
        cmdDeleteRecord.Parameters.Append(prmResultCode)

        Dim prmTableID = cmdDeleteRecord.CreateParameter("tableID", 3, 1)
        cmdDeleteRecord.Parameters.Append(prmTableID)
        prmTableID.value = CLng(CleanNumeric(lngTableID))

        Dim prmRealSource = cmdDeleteRecord.CreateParameter("realSource", 200, 1, 8000)
        cmdDeleteRecord.Parameters.Append(prmRealSource)
        prmRealSource.value = CleanString(sRealSource)

        Dim prmID = cmdDeleteRecord.CreateParameter("id", 3, 1)
        cmdDeleteRecord.Parameters.Append(prmID)
        prmID.value = CleanNumeric(lngRecordID)

        Dim fDeadlock = True
        Do While fDeadlock
          fDeadlock = False

          cmdDeleteRecord.ActiveConnection.Errors.Clear()

          ' Run the delete stored procedure.
          cmdDeleteRecord.Execute()

          If cmdDeleteRecord.ActiveConnection.Errors.Count > 0 Then
            For iLoop = 1 To cmdDeleteRecord.ActiveConnection.Errors.Count
              sErrMsg = FormatError(cmdDeleteRecord.ActiveConnection.Errors.Item(iLoop - 1).Description)

              If (cmdDeleteRecord.ActiveConnection.Errors.Item(iLoop - 1).Number = DEADLOCK_ERRORNUMBER) And _
               (((UCase(Left(sErrMsg, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
                (UCase(Right(sErrMsg, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
               ((UCase(Left(sErrMsg, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
               (InStr(UCase(sErrMsg), DEADLOCK2_MESSAGEEND) > 0))) Then

                ' The error is for a deadlock.
                ' Sorry about having to use the err.description to trap the error but the err.number
                ' is not specific and MSDN suggests using the err.description.
                If (iRetryCount < iRETRIES) And (cmdDeleteRecord.ActiveConnection.Errors.Count = 1) Then
                  iRetryCount = iRetryCount + 1
                  fDeadlock = True
                Else
                  If Len(sErrorMsg) > 0 Then
                    sErrorMsg = sErrorMsg & vbCrLf
                  End If
                  sErrorMsg = sErrorMsg & "Another user is deadlocking the database. Try saving again."
                  fOk = False
                End If
              ElseIf cmdDeleteRecord.ActiveConnection.Errors.Item(iLoop - 1).nativeerror = 3609 Then
                ' Ignore the follow on message that says "The transaction ended in the trigger."
              Else
                sErrorMsg = sErrorMsg & vbCrLf & _
                 FormatError(cmdDeleteRecord.ActiveConnection.Errors.Item(iLoop - 1).Description)
                fOk = False
              End If
            Next

            cmdDeleteRecord.ActiveConnection.Errors.Clear()

            If Not fOk Then
              sErrorMsg = "The record could not be deleted." & vbCrLf & sErrorMsg
              sAction = "SAVEERROR"
            End If
          Else
            Select Case cmdDeleteRecord.Parameters("resultCode").Value
              Case 2 ' Record changed by another user, and is no longer in the current table/view.
                sErrorMsg = "The record has been amended by another user and is no longer in the current view."
            End Select

            lngRecordID = 0

            If Len(sReaction) > 0 Then
              sAction = sReaction
            Else
              sAction = "LOAD"
            End If
          End If
        Loop
        cmdDeleteRecord = Nothing

        'MH20100609
        Dim cmdInsertRecord = CreateObject("ADODB.Command")
        cmdInsertRecord.CommandText = "spASREmailImmediate"
        cmdInsertRecord.CommandType = 4 ' Stored procedure
        cmdInsertRecord.CommandTimeout = 180
        cmdInsertRecord.ActiveConnection = Session("databaseConnection")

        Dim prmInsertSQL = cmdInsertRecord.CreateParameter("Username", 200, 1, 255) ' 200=varchar, 1=input, 255=size
        cmdInsertRecord.Parameters.Append(prmInsertSQL)
        prmInsertSQL.value = Session("Username")

        Err.Clear()
        cmdInsertRecord.Execute()
        cmdInsertRecord = Nothing

      ElseIf sAction = "CANCELCOURSE" Then
        ' Check number of bookings made.
        Dim cmdCancelCourse = CreateObject("ADODB.Command")
        cmdCancelCourse.CommandText = "sp_ASRIntCancelCourse"
        cmdCancelCourse.CommandType = 4 ' Stored procedure
        cmdCancelCourse.ActiveConnection = Session("databaseConnection")

        Dim prmNumberOfBookings = cmdCancelCourse.CreateParameter("numberOfBookings", 3, 2) ' 3=integer, 2=output
        cmdCancelCourse.Parameters.Append(prmNumberOfBookings)

        Dim prmCourseRecordID = cmdCancelCourse.CreateParameter("courseRecordID", 3, 1)
        cmdCancelCourse.Parameters.Append(prmCourseRecordID)
        prmCourseRecordID.value = CleanNumeric(lngRecordID)

        Dim prmTBTableID = cmdCancelCourse.CreateParameter("tbTableID", 3, 1)
        cmdCancelCourse.Parameters.Append(prmTBTableID)
        prmTBTableID.value = CleanNumeric(Session("TB_TBTableID"))

        Dim prmCourseTableID = cmdCancelCourse.CreateParameter("courseTableID", 3, 1)
        cmdCancelCourse.Parameters.Append(prmCourseTableID)
        prmCourseTableID.value = CleanNumeric(Session("TB_CourseTableID"))

        Dim prmTrainBookStatusColumnID = cmdCancelCourse.CreateParameter("tbStatusColumnID", 3, 1)
        cmdCancelCourse.Parameters.Append(prmTrainBookStatusColumnID)
        prmTrainBookStatusColumnID.value = CleanNumeric(Session("TB_TBStatusColumnID"))

        Dim prmCourseRealSource = cmdCancelCourse.CreateParameter("courseRealSource", 200, 1, 8000)
        cmdCancelCourse.Parameters.Append(prmCourseRealSource)
        prmCourseRealSource.value = sRealSource

        Dim prmErrorMessage = cmdCancelCourse.CreateParameter("errorMessage", 200, 2, 8000)
        cmdCancelCourse.Parameters.Append(prmErrorMessage)

        Dim prmCourseTitle = cmdCancelCourse.CreateParameter("courseTitle", 200, 2, 8000)
        cmdCancelCourse.Parameters.Append(prmCourseTitle)

        Err.Clear()
        cmdCancelCourse.Execute()
        If (Err.Number <> 0) Then
          sErrorMsg = "Error cancelling the course." & vbCrLf & FormatError(Err.Description)
          sAction = "SAVEERROR"
        Else
          sAction = "CANCELCOURSE_1"
          Session("numberOfBookings") = cmdCancelCourse.Parameters("numberOfBookings").Value
          Session("tbErrorMessage") = cmdCancelCourse.Parameters("errorMessage").Value
          Session("tbCourseTitle") = cmdCancelCourse.Parameters("courseTitle").Value
        End If

        cmdCancelCourse = Nothing
      ElseIf sAction = "CANCELCOURSE_2" Then
        Dim cmdCancelCourse = CreateObject("ADODB.Command")
        cmdCancelCourse.CommandText = "sp_ASRIntCancelCoursePart2"
        cmdCancelCourse.CommandType = 4 ' Stored procedure
        cmdCancelCourse.ActiveConnection = Session("databaseConnection")

        Dim prmEmployeeTableID = cmdCancelCourse.CreateParameter("employeeTableID", 3, 1)
        cmdCancelCourse.Parameters.Append(prmEmployeeTableID)
        prmEmployeeTableID.value = CleanNumeric(Session("TB_EmpTableID"))

        Dim prmCourseTableID = cmdCancelCourse.CreateParameter("courseTableID", 3, 1)
        cmdCancelCourse.Parameters.Append(prmCourseTableID)
        prmCourseTableID.value = CleanNumeric(Session("TB_CourseTableID"))

        Dim prmCourseRealSource = cmdCancelCourse.CreateParameter("courseRealSource", 200, 1, 8000)
        cmdCancelCourse.Parameters.Append(prmCourseRealSource)
        prmCourseRealSource.value = sRealSource

        Dim prmCourseRecordID = cmdCancelCourse.CreateParameter("courseRecordID", 3, 1)
        cmdCancelCourse.Parameters.Append(prmCourseRecordID)
        prmCourseRecordID.value = CleanNumeric(lngRecordID)

        Dim prmNewCourseRecordID = cmdCancelCourse.CreateParameter("newCourseRecordID", 3, 1)
        cmdCancelCourse.Parameters.Append(prmNewCourseRecordID)
        prmNewCourseRecordID.value = CleanNumeric(iTBCourseRecordID)

        Dim prmCourseCancelDateColumnID = cmdCancelCourse.CreateParameter("courseCancelDateColumnID", 3, 1)
        cmdCancelCourse.Parameters.Append(prmCourseCancelDateColumnID)
        prmCourseCancelDateColumnID.value = CleanNumeric(Session("TB_CourseCancelDateColumnID"))

        Dim prmCourseTitle = cmdCancelCourse.CreateParameter("courseTitle", 200, 1, 8000)
        cmdCancelCourse.Parameters.Append(prmCourseTitle)
        prmCourseTitle.value = Session("tbCourseTitle")

        Dim prmTBTableID = cmdCancelCourse.CreateParameter("tbTableID", 3, 1)
        cmdCancelCourse.Parameters.Append(prmTBTableID)
        prmTBTableID.value = CleanNumeric(Session("TB_TBTableID"))

        Dim prmTBTableInsert = cmdCancelCourse.CreateParameter("tbTableInsert", 11, 1)  ' 11=boolean, 1=input
        cmdCancelCourse.Parameters.Append(prmTBTableInsert)
        prmTBTableInsert.value = CleanBoolean(Session("TB_TBTableInsert"))

        Dim prmTrainBookStatusColumnID = cmdCancelCourse.CreateParameter("tbStatusColumnID", 3, 1)
        cmdCancelCourse.Parameters.Append(prmTrainBookStatusColumnID)
        prmTrainBookStatusColumnID.value = CleanNumeric(Session("TB_TBStatusColumnID"))

        Dim prmTrainBookCancelDateColumnID = cmdCancelCourse.CreateParameter("tbCancelDateColumnID", 3, 1)
        cmdCancelCourse.Parameters.Append(prmTrainBookCancelDateColumnID)
        prmTrainBookCancelDateColumnID.value = CleanNumeric(Session("TB_TBCancelDateColumnID"))

        Dim prmWLTableID = cmdCancelCourse.CreateParameter("wlTableID", 3, 1)
        cmdCancelCourse.Parameters.Append(prmWLTableID)
        prmWLTableID.value = CleanNumeric(Session("TB_WaitListTableID"))

        Dim prmWLTableInsert = cmdCancelCourse.CreateParameter("wlTableInsert", 11, 1)  ' 11=boolean, 1=input
        cmdCancelCourse.Parameters.Append(prmWLTableInsert)
        prmWLTableInsert.value = CleanBoolean(Session("TB_WaitListTableInsert"))

        Dim prmWLCourseTitleColumnID = cmdCancelCourse.CreateParameter("wlCourseTitleColumnID", 3, 1)
        cmdCancelCourse.Parameters.Append(prmWLCourseTitleColumnID)
        prmWLCourseTitleColumnID.value = CleanNumeric(Session("TB_WaitListCourseTitleColumnID"))

        Dim prmWLCourseTitleColumnUpdate = cmdCancelCourse.CreateParameter("wlCourseTitleColumnUpdate", 11, 1)  ' 11=boolean, 1=input
        cmdCancelCourse.Parameters.Append(prmWLCourseTitleColumnUpdate)
        prmWLCourseTitleColumnUpdate.value = CleanBoolean(Session("TB_WaitListCourseTitleColumnUpdate"))

        Dim prmCreateWaitListRecords = cmdCancelCourse.CreateParameter("createWaitListRecords", 11, 1) ' 11=boolean, 1=input
        cmdCancelCourse.Parameters.Append(prmCreateWaitListRecords)
        prmCreateWaitListRecords.value = CleanBoolean(Request.Form("txtTBCreateWLRecords"))

        Dim prmErrorMessage = cmdCancelCourse.CreateParameter("errorMessage", 200, 2, 8000)
        cmdCancelCourse.Parameters.Append(prmErrorMessage)

        Err.Clear()
        cmdCancelCourse.Execute()

        If (Err.Number <> 0) Then
          sErrorMsg = "Error cancelling the course." & vbCrLf & FormatError(Err.Description)
          sAction = "SAVEERROR"
        Else
          sErrorMsg = cmdCancelCourse.Parameters("errorMessage").Value

          If Len(sErrorMsg) > 0 Then
            sAction = "SAVEERROR"
          Else
            sAction = "LOAD"
          End If
        End If

        cmdCancelCourse = Nothing

      ElseIf sAction = "CANCELBOOKING" Then
        Dim cmdCancelBooking = CreateObject("ADODB.Command")
        cmdCancelBooking.CommandText = "sp_ASRIntCancelBooking"
        cmdCancelBooking.CommandType = 4    ' Stored procedure
        cmdCancelBooking.ActiveConnection = Session("databaseConnection")

        Dim prmTransferBooking = cmdCancelBooking.CreateParameter("transferBooking", 11, 1) '11=boolean, 1=input
        cmdCancelBooking.Parameters.Append(prmTransferBooking)
        prmTransferBooking.value = CleanBoolean(fUserChoice)

        Dim prmTBRecordID = cmdCancelBooking.CreateParameter("tbRecordID", 3, 1)  '3=integer, 1=input
        cmdCancelBooking.Parameters.Append(prmTBRecordID)
        prmTBRecordID.value = CleanNumeric(lngRecordID)

        Dim prmErrorMessage = cmdCancelBooking.CreateParameter("errorMessage", 200, 2, 8000)  '2=varchar, 2=output, 8000=size
        cmdCancelBooking.Parameters.Append(prmErrorMessage)

        Err.Clear()
        cmdCancelBooking.Execute()
        If (Err.Number <> 0) Then
          sErrorMsg = "Error cancelling the booking." & vbCrLf & FormatError(Err.Description)
          sAction = "SAVEERROR"
        Else
          sErrorMsg = cmdCancelBooking.Parameters("errorMessage").Value

          If Len(sErrorMsg) > 0 Then
            sAction = "SAVEERROR"
          Else
            sAction = "CANCELBOOKING_1"
          End If
        End If

        cmdCancelBooking = Nothing
      End If

      Session("selectSQL") = Request.Form("txtSelectSQL")
      Session("fromDef") = Request.Form("txtFromDef")
      Session("filterSQL") = Request.Form("txtFilterSQL")
      Session("filterDef") = Request.Form("txtFilterDef")
      Session("realSource") = sRealSource
      Session("tableID") = lngTableID
      Session("screenID") = lngScreenID
      Session("viewID") = lngViewID
      Session("recordID") = lngRecordID
      Session("action") = sAction
      Session("reaction") = ""
      Session("warningFlag") = fWarning
      Session("parentTableID") = Request.Form("txtParentTableID")
      Session("parentRecordID") = Request.Form("txtParentRecordID")
      Session("defaultCalcColumns") = Request.Form("txtDefaultCalcCols")
      Session("insertUpdateDef") = sInsertUpdateDef
      Session("errorMessage") = sErrorMsg
      Session("ReportBaseTableID") = Request.Form("txtReportBaseTableID")
      Session("ReportParent1TableID") = Request.Form("txtReportParent1TableID")
      Session("ReportParent2TableID") = Request.Form("txtReportParent2TableID")
      Session("ReportChildTableID") = Request.Form("txtReportChildTableID")
      Session("Param1") = Request.Form("txtParam1")

      'JDM - 24/07/02 - Fault 3917 - Reset year for absence calendar
      Session("stdrpt_AbsenceCalendar_StartYear") = Year(DateTime.Now())

      'JDM - 10/10/02 - Fault 4534 - Reset start month for absence calendar
      Session("stdrpt_AbsenceCalendar_StartMonth") = ""

      'TM - 05/09/02 - Store the event log parameters in session vaiables.
      Session("ELFilterUser") = Request.Form("txtELFilterUser")
      Session("ELFilterType") = Request.Form("txtELFilterType")
      Session("ELFilterStatus") = Request.Form("txtELFilterStatus")
      Session("ELFilterMode") = Request.Form("txtELFilterMode")
      Session("ELOrderColumn") = Request.Form("txtELOrderColumn")
      Session("ELOrderOrder") = Request.Form("txtELOrderOrder")

      Session("ELAction") = Request.Form("txtELAction")

      Session("ELCurrentRecCount") = Request.Form("txtELCurrRecCount")
      If Session("ELCurrentRecCount") < 1 Or Len(Session("ELCurrentRecCount")) < 1 Then
        Session("ELCurrentRecCount") = 0
      End If

      Session("ELFirstRecPos") = Request.Form("txtEL1stRecPos")
      If Session("ELFirstRecPos") < 1 Or Len(Session("ELFirstRecPos")) < 1 Then
        Session("ELFirstRecPos") = 1
      End If

      ' Go to the requested page.
      Return RedirectToAction("Data", "Home")

    End Function

    Function Util_RecordSelection() As ActionResult
      Return View()
    End Function

    Function Util_CustomReportChilds() As ActionResult
      Return View()
    End Function

    Function Util_EmailSelection() As ActionResult
      Return View()
    End Function

    Function Util_SortOrderSelection() As ActionResult
      Return View()
    End Function

    Function LinksMain() As ActionResult
      If Session("objButtonInfo") Is Nothing Or Session("objHypertextInfo") Is Nothing Or Session("objDropdownInfo") Is Nothing Then
        Return RedirectToAction("Login", "Account")
      End If

      Dim objHypertextInfo As VBA.Collection = Session("objHypertextInfo")
      Dim objButtonInfo As VBA.Collection = Session("objButtonInfo")
      Dim objDropdownInfo As VBA.Collection = Session("objDropdownInfo")

      Dim lstButtonInfo = (From collectionItem As Object In objHypertextInfo Select New navigationLink(collectionItem.ID, collectionItem.DrillDownHidden, collectionItem.LinkType, collectionItem.LinkOrder, collectionItem.Text, collectionItem.Text1, collectionItem.Text2, collectionItem.Prompt, collectionItem.ScreenID, collectionItem.TableID, collectionItem.ViewID, collectionItem.PageTitle, collectionItem.URL, collectionItem.UtilityType, collectionItem.UtilityID, collectionItem.NewWindow, collectionItem.BaseTable, collectionItem.LinkToFind, collectionItem.SingleRecord, collectionItem.PrimarySequence, collectionItem.SecondarySequence, collectionItem.FindPage, collectionItem.EmailAddress, collectionItem.EmailSubject, collectionItem.AppFilePath, collectionItem.AppParameters, collectionItem.DocumentFilePath, collectionItem.DisplayDocumentHyperlink, collectionItem.IsSeparator, collectionItem.Element_Type, collectionItem.SeparatorOrientation, collectionItem.PictureID, collectionItem.Chart_ShowLegend, collectionItem.Chart_Type, collectionItem.Chart_ShowGrid, collectionItem.Chart_StackSeries, collectionItem.Chart_ShowValues, collectionItem.Chart_ViewID, collectionItem.Chart_TableID, collectionItem.Chart_ColumnID, collectionItem.Chart_FilterID, collectionItem.Chart_AggregateType, collectionItem.Chart_ColumnName, collectionItem.Chart_ColumnName_2, collectionItem.UseFormatting, collectionItem.Formatting_DecimalPlaces, collectionItem.Formatting_Use1000Separator, collectionItem.Formatting_Prefix, collectionItem.Formatting_Suffix, collectionItem.UseConditionalFormatting, collectionItem.ConditionalFormatting_Operator_1, collectionItem.ConditionalFormatting_Value_1, collectionItem.ConditionalFormatting_Style_1, collectionItem.ConditionalFormatting_Colour_1, collectionItem.ConditionalFormatting_Operator_2, collectionItem.ConditionalFormatting_Value_2, collectionItem.ConditionalFormatting_Style_2, collectionItem.ConditionalFormatting_Colour_2, collectionItem.ConditionalFormatting_Operator_3, collectionItem.ConditionalFormatting_Value_3, collectionItem.ConditionalFormatting_Style_3, collectionItem.ConditionalFormatting_Colour_3, collectionItem.SeparatorColour, collectionItem.InitialDisplayMode, collectionItem.Chart_TableID_2, collectionItem.Chart_ColumnID_2, collectionItem.Chart_TableID_3, collectionItem.Chart_ColumnID_3, collectionItem.Chart_SortOrderID, collectionItem.Chart_SortDirection, collectionItem.Chart_ColourID, collectionItem.Chart_ShowPercentages)).ToList()
      lstButtonInfo.AddRange(From collectionItem As Object In objButtonInfo Select New navigationLink(collectionItem.ID, collectionItem.DrillDownHidden, collectionItem.LinkType, collectionItem.LinkOrder, collectionItem.Text, collectionItem.Text1, collectionItem.Text2, collectionItem.Prompt, collectionItem.ScreenID, collectionItem.TableID, collectionItem.ViewID, collectionItem.PageTitle, collectionItem.URL, collectionItem.UtilityType, collectionItem.UtilityID, collectionItem.NewWindow, collectionItem.BaseTable, collectionItem.LinkToFind, collectionItem.SingleRecord, collectionItem.PrimarySequence, collectionItem.SecondarySequence, collectionItem.FindPage, collectionItem.EmailAddress, collectionItem.EmailSubject, collectionItem.AppFilePath, collectionItem.AppParameters, collectionItem.DocumentFilePath, collectionItem.DisplayDocumentHyperlink, collectionItem.IsSeparator, collectionItem.Element_Type, collectionItem.SeparatorOrientation, collectionItem.PictureID, collectionItem.Chart_ShowLegend, collectionItem.Chart_Type, collectionItem.Chart_ShowGrid, collectionItem.Chart_StackSeries, collectionItem.Chart_ShowValues, collectionItem.Chart_ViewID, collectionItem.Chart_TableID, collectionItem.Chart_ColumnID, collectionItem.Chart_FilterID, collectionItem.Chart_AggregateType, collectionItem.Chart_ColumnName, collectionItem.Chart_ColumnName_2, collectionItem.UseFormatting, collectionItem.Formatting_DecimalPlaces, collectionItem.Formatting_Use1000Separator, collectionItem.Formatting_Prefix, collectionItem.Formatting_Suffix, collectionItem.UseConditionalFormatting, collectionItem.ConditionalFormatting_Operator_1, collectionItem.ConditionalFormatting_Value_1, collectionItem.ConditionalFormatting_Style_1, collectionItem.ConditionalFormatting_Colour_1, collectionItem.ConditionalFormatting_Operator_2, collectionItem.ConditionalFormatting_Value_2, collectionItem.ConditionalFormatting_Style_2, collectionItem.ConditionalFormatting_Colour_2, collectionItem.ConditionalFormatting_Operator_3, collectionItem.ConditionalFormatting_Value_3, collectionItem.ConditionalFormatting_Style_3, collectionItem.ConditionalFormatting_Colour_3, collectionItem.SeparatorColour, collectionItem.InitialDisplayMode, collectionItem.Chart_TableID_2, collectionItem.Chart_ColumnID_2, collectionItem.Chart_TableID_3, collectionItem.Chart_ColumnID_3, collectionItem.Chart_SortOrderID, collectionItem.Chart_SortDirection, collectionItem.Chart_ColourID, collectionItem.Chart_ShowPercentages))
      lstButtonInfo.AddRange(From collectionItem As Object In objDropdownInfo Select New navigationLink(collectionItem.ID, collectionItem.DrillDownHidden, collectionItem.LinkType, collectionItem.LinkOrder, collectionItem.Text, collectionItem.Text1, collectionItem.Text2, collectionItem.Prompt, collectionItem.ScreenID, collectionItem.TableID, collectionItem.ViewID, collectionItem.PageTitle, collectionItem.URL, collectionItem.UtilityType, collectionItem.UtilityID, collectionItem.NewWindow, collectionItem.BaseTable, collectionItem.LinkToFind, collectionItem.SingleRecord, collectionItem.PrimarySequence, collectionItem.SecondarySequence, collectionItem.FindPage, collectionItem.EmailAddress, collectionItem.EmailSubject, collectionItem.AppFilePath, collectionItem.AppParameters, collectionItem.DocumentFilePath, collectionItem.DisplayDocumentHyperlink, collectionItem.IsSeparator, collectionItem.Element_Type, collectionItem.SeparatorOrientation, collectionItem.PictureID, collectionItem.Chart_ShowLegend, collectionItem.Chart_Type, collectionItem.Chart_ShowGrid, collectionItem.Chart_StackSeries, collectionItem.Chart_ShowValues, collectionItem.Chart_ViewID, collectionItem.Chart_TableID, collectionItem.Chart_ColumnID, collectionItem.Chart_FilterID, collectionItem.Chart_AggregateType, collectionItem.Chart_ColumnName, collectionItem.Chart_ColumnName_2, collectionItem.UseFormatting, collectionItem.Formatting_DecimalPlaces, collectionItem.Formatting_Use1000Separator, collectionItem.Formatting_Prefix, collectionItem.Formatting_Suffix, collectionItem.UseConditionalFormatting, collectionItem.ConditionalFormatting_Operator_1, collectionItem.ConditionalFormatting_Value_1, collectionItem.ConditionalFormatting_Style_1, collectionItem.ConditionalFormatting_Colour_1, collectionItem.ConditionalFormatting_Operator_2, collectionItem.ConditionalFormatting_Value_2, collectionItem.ConditionalFormatting_Style_2, collectionItem.ConditionalFormatting_Colour_2, collectionItem.ConditionalFormatting_Operator_3, collectionItem.ConditionalFormatting_Value_3, collectionItem.ConditionalFormatting_Style_3, collectionItem.ConditionalFormatting_Colour_3, collectionItem.SeparatorColour, collectionItem.InitialDisplayMode, collectionItem.Chart_TableID_2, collectionItem.Chart_ColumnID_2, collectionItem.Chart_TableID_3, collectionItem.Chart_ColumnID_3, collectionItem.Chart_SortOrderID, collectionItem.Chart_SortDirection, collectionItem.Chart_ColourID, collectionItem.Chart_ShowPercentages))

      Dim viewModel = New NavLinksViewModel With {.NavigationLinks = lstButtonInfo, .NumberOfLinks = objDropdownInfo.Count}

      Return View(viewModel)
    End Function

    ' TODO
    Public Sub ShowPhoto(imageName As String)
      'TODO fetch path from registry
      Dim localImagesPath As String = HttpContext.Server.MapPath("~/pictures/profilephotos/")

      'TODO fetch imagename from db
      Dim file = localImagesPath & imageName
      Dim fStream As New FileStream(file, FileMode.Open, FileAccess.Read)
      Dim br As New BinaryReader(fStream)

      ' Show the number of bytes in the array.
      br.Close()
      fStream.Close()

      Response.ContentType = "image/png"
      Response.WriteFile(file)

    End Sub

    Function LogOff()
      Session("databaseConnection") = Nothing
      Return RedirectToAction("Login", "Account")
    End Function

    Function PasswordChange() As ActionResult
      Return View()
    End Function

    'Function ForcePasswordChange() As ActionResult
    '    Return View()
    'End Function


    Function Poll() As ActionResult
      Return PartialView()
    End Function

#Region "Event Log Forms"

    Function emailSelection() As ActionResult
      Return View()
    End Function

    Function EventLog() As ActionResult
      Return View()
    End Function

    Function EventLogDetails() As ActionResult
      Return View()
    End Function

    Function EventLogPurge() As ActionResult
      Return View()
    End Function

    Function EventLogSelection() As ActionResult
      Return View()
    End Function

#End Region

#Region "Running Reports"

    Function util_run_promptedvalues() As ActionResult
      Return View()
    End Function

    Function util_run() As ActionResult
      Return View()
    End Function

    Function util_run_customreports() As ActionResult
      Return PartialView()
    End Function

    Function util_run_customreportsData() As ActionResult
      Return PartialView()
    End Function

    Function util_run_customreportsMain() As ActionResult
      Return PartialView()
    End Function

    Function Progress() As ActionResult
      Return PartialView()
    End Function

    Function Refresh() As ActionResult
      Return View()
    End Function

    '  Function util_run_promptedvaluessubmit() As ActionResult
    '     Return RedirectToAction("util_run")
    '    End Function



#End Region

#Region "Defining Reports"

    Function util_def_customreports_submit()
      On Error Resume Next

      Dim cmdSave = CreateObject("ADODB.Command")
      cmdSave.CommandText = "sp_ASRIntSaveCustomReport"
      cmdSave.CommandType = 4 ' Stored Procedure
      cmdSave.ActiveConnection = Session("databaseConnection")

      Dim prmName = cmdSave.CreateParameter("name", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmName)
      prmName.value = Request.Form("txtSend_name")

      Dim prmDescription = cmdSave.CreateParameter("description", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmDescription)
      prmDescription.value = Request.Form("txtSend_description")

      Dim prmBaseTable = cmdSave.CreateParameter("baseTable", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmBaseTable)
      prmBaseTable.value = CleanNumeric(Request.Form("txtSend_baseTable"))

      Dim prmAllRecords = cmdSave.CreateParameter("allRecords", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmAllRecords)
      prmAllRecords.value = CleanBoolean(Request.Form("txtSend_allRecords"))

      Dim prmPicklistID = cmdSave.CreateParameter("picklistID", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmPicklistID)
      prmPicklistID.value = CleanNumeric(Request.Form("txtSend_picklist"))

      Dim prmFilterID = cmdSave.CreateParameter("filterID", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmFilterID)
      prmFilterID.value = CleanNumeric(Request.Form("txtSend_filter"))

      Dim prmParent1Table = cmdSave.CreateParameter("parent1Table", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmParent1Table)
      prmParent1Table.value = CleanNumeric(Request.Form("txtSend_parent1Table"))

      Dim prmParent1Filter = cmdSave.CreateParameter("parent1Filter", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmParent1Filter)
      prmParent1Filter.value = CleanNumeric(Request.Form("txtSend_parent1Filter"))

      Dim prmParent2Table = cmdSave.CreateParameter("parent2Table", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmParent2Table)
      prmParent2Table.value = CleanNumeric(Request.Form("txtSend_parent2Table"))

      Dim prmParent2Filter = cmdSave.CreateParameter("parent2Filter", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmParent2Filter)
      prmParent2Filter.value = CleanNumeric(Request.Form("txtSend_parent2Filter"))

      Dim prmSummary = cmdSave.CreateParameter("summary", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmSummary)
      prmSummary.value = CleanBoolean(Request.Form("txtSend_summary"))

      Dim prmPrintFilterHeader = cmdSave.CreateParameter("printFilterHeader", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmPrintFilterHeader)
      prmPrintFilterHeader.value = CleanBoolean(Request.Form("txtSend_printFilterHeader"))

      Dim prmUserName = cmdSave.CreateParameter("userName", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmUserName)
      prmUserName.value = Request.Form("txtSend_userName")

      Dim prmOutputPreview = cmdSave.CreateParameter("outputPreview", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmOutputPreview)
      prmOutputPreview.value = CleanBoolean(Request.Form("txtSend_OutputPreview"))

      Dim prmOutputFormat = cmdSave.CreateParameter("outputFormat", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmOutputFormat)
      prmOutputFormat.value = CleanNumeric(Request.Form("txtSend_OutputFormat"))

      Dim prmOutputScreen = cmdSave.CreateParameter("outputScreen", 11, 1)  ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmOutputScreen)
      prmOutputScreen.value = CleanBoolean(Request.Form("txtSend_OutputScreen"))

      Dim prmOutputPrinter = cmdSave.CreateParameter("outputPrinter", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmOutputPrinter)
      prmOutputPrinter.value = CleanBoolean(Request.Form("txtSend_OutputPrinter"))

      Dim prmOutputPrinterName = cmdSave.CreateParameter("outputPrinterName", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmOutputPrinterName)
      prmOutputPrinterName.value = Request.Form("txtSend_OutputPrinterName")

      Dim prmOutputSave = cmdSave.CreateParameter("outputSave", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmOutputSave)
      prmOutputSave.value = CleanBoolean(Request.Form("txtSend_OutputSave"))

      Dim prmOutputSaveExisting = cmdSave.CreateParameter("outputSaveExisting", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmOutputSaveExisting)
      prmOutputSaveExisting.value = CleanNumeric(Request.Form("txtSend_OutputSaveExisting"))

      Dim prmOutputEmail = cmdSave.CreateParameter("outputEmail", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmOutputEmail)
      prmOutputEmail.value = CleanBoolean(Request.Form("txtSend_OutputEmail"))

      Dim prmOutputEmailAddr = cmdSave.CreateParameter("outputEmailAddr", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmOutputEmailAddr)
      prmOutputEmailAddr.value = CleanNumeric(Request.Form("txtSend_OutputEmailAddr"))

      Dim prmOutputEmailSubject = cmdSave.CreateParameter("outputEmailSubject", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmOutputEmailSubject)
      prmOutputEmailSubject.value = Request.Form("txtSend_OutputEmailSubject")

      Dim prmOutputEmailAttachAs = cmdSave.CreateParameter("outputEmailAttachAs", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmOutputEmailAttachAs)
      prmOutputEmailAttachAs.value = Request.Form("txtSend_OutputEmailAttachAs")

      Dim prmOutputFilename = cmdSave.CreateParameter("outputFilename", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmOutputFilename)
      prmOutputFilename.value = Request.Form("txtSend_OutputFilename")

      Dim prmParent1AllRecords = cmdSave.CreateParameter("parent1AllRecords", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmParent1AllRecords)
      prmParent1AllRecords.value = CleanBoolean(Request.Form("txtSend_parent1AllRecords"))

      Dim prmParent1Picklist = cmdSave.CreateParameter("parent1Picklist", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmParent1Picklist)
      prmParent1Picklist.value = CleanNumeric(Request.Form("txtSend_parent1Picklist"))

      Dim prmParent2AllRecords = cmdSave.CreateParameter("parent2AllRecords", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmParent2AllRecords)
      prmParent2AllRecords.value = CleanBoolean(Request.Form("txtSend_parent2AllRecords"))

      Dim prmParent2Picklist = cmdSave.CreateParameter("parent2Picklist", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmParent2Picklist)
      prmParent2Picklist.value = CleanNumeric(Request.Form("txtSend_parent2Picklist"))

      Dim prmAccess = cmdSave.CreateParameter("access", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmAccess)
      prmAccess.value = Request.Form("txtSend_access")

      Dim prmJobToHide = cmdSave.CreateParameter("jobsToHide", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmJobToHide)
      prmJobToHide.value = Request.Form("txtSend_jobsToHide")

      Dim prmJobToHideGroups = cmdSave.CreateParameter("acess", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmJobToHideGroups)
      prmJobToHideGroups.value = Request.Form("txtSend_jobsToHideGroups")

      Dim prmColumns = cmdSave.CreateParameter("columns", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmColumns)
      prmColumns.value = Request.Form("txtSend_columns")

      Dim prmColumns2 = cmdSave.CreateParameter("columns2", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmColumns2)
      prmColumns2.value = Request.Form("txtSend_columns2")

      'pass the child string to the stored procedure, the stored procedure 
      'saves the child information to the ASRSysCustomReportChildDetails table.
      Dim prmChildTables = cmdSave.CreateParameter("childstring", 200, 1, 8000)
      cmdSave.Parameters.Append(prmChildTables)
      prmChildTables.value = Request.Form("txtSend_childTable")

      Dim prmID = cmdSave.CreateParameter("id", 3, 3) ' 3=integer,3=input/output
      cmdSave.Parameters.Append(prmID)
      prmID.value = CleanNumeric(Request.Form("txtSend_ID"))

      Dim prmIgnoreZeros = cmdSave.CreateParameter("ignoreZeros", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmIgnoreZeros)
      prmIgnoreZeros.value = CleanBoolean(Request.Form("txtSend_IgnoreZeros"))

      cmdSave.Execute()

      If Err.Number = 0 Then
        Session("confirmtext") = "Report has been saved successfully"
        Session("confirmtitle") = "Custom Reports"
        Session("followpage") = "defsel"
        Session("reaction") = Request.Form("txtSend_reaction")
        Session("utilid") = cmdSave.Parameters("id").Value

        Return RedirectToAction("confirmok")

      Else

        ' TO DO error reprting
        Response.Write("<HTML>" & vbCrLf)
        Response.Write("	<HEAD>" & vbCrLf)
        Response.Write("		<META NAME=""GENERATOR"" Content=""Microsoft Visual Studio 6.0"">" & vbCrLf)
        Response.Write("		<LINK href=""OpenHR.css"" rel=stylesheet type=text/css >" & vbCrLf)
        Response.Write("		<TITLE>" & vbCrLf)
        Response.Write("			OpenHR Intranet" & vbCrLf)
        Response.Write("		</TITLE>" & vbCrLf)
        Response.Write("  <!--#INCLUDE FILE=""include/ctl_SetStyles.txt"" -->")
        Response.Write("	</HEAD>" & vbCrLf)
        Response.Write("	<BODY id=bdyMainBody name=""bdyMainBody"" " & Session("BodyTag") & ">" & vbCrLf)

        Response.Write("	<table align=center class=""outline"" cellPadding=5 cellSpacing=0>" & vbCrLf)
        Response.Write("		<TR>" & vbCrLf)
        Response.Write("			<TD>" & vbCrLf)
        Response.Write("				<table class=""invisible"" cellspacing=0 cellpadding=0>" & vbCrLf)
        Response.Write("				  <tr> " & vbCrLf)
        Response.Write("				    <td colspan=3 height=10></td>" & vbCrLf)
        Response.Write("				  </tr>" & vbCrLf)
        Response.Write("				  <tr> " & vbCrLf)
        Response.Write("				    <td colspan=3 align=center> " & vbCrLf)
        Response.Write("							<H3>Error</H3>" & vbCrLf)
        Response.Write("				    </td>" & vbCrLf)
        Response.Write("				  </tr>" & vbCrLf)
        Response.Write("				  <tr> " & vbCrLf)
        Response.Write("				    <td width=20 height=10></td> " & vbCrLf)
        Response.Write("				    <td> " & vbCrLf)
        Response.Write("							<H4>Error saving report</H4>" & vbCrLf)
        Response.Write("				    </td>" & vbCrLf)
        Response.Write("				    <td width=20></td> " & vbCrLf)
        Response.Write("				  </tr>" & vbCrLf)
        Response.Write("				  <tr> " & vbCrLf)
        Response.Write("				    <td width=20 height=10></td> " & vbCrLf)
        Response.Write("				    <td> " & vbCrLf)
        Response.Write(Err.Description & vbCrLf)
        Response.Write("			    </td>" & vbCrLf)
        Response.Write("			    <td width=20></td> " & vbCrLf)
        Response.Write("			  </tr>" & vbCrLf)
        Response.Write("			  <tr> " & vbCrLf)
        Response.Write("			    <td colspan=3 height=20></td>" & vbCrLf)
        Response.Write("			  </tr>" & vbCrLf)
        Response.Write("			  <tr> " & vbCrLf)
        Response.Write("			    <td colspan=3 height=10 align=center>" & vbCrLf)
        Response.Write("						<INPUT TYPE=button VALUE=""Retry"" NAME=""GoBack"" class=""btn"" OnClick=""window.history.back(1)"" style=""WIDTH: 80px"" width=80 id=cmdGoBack>" & vbCrLf)
        Response.Write("                      onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
        Response.Write("                      onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
        Response.Write("		                  onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
        Response.Write("                      onblur=""try{button_onBlur(this);}catch(e){}"" />" & vbCrLf)
        Response.Write("			    </td>" & vbCrLf)
        Response.Write("			  </tr>" & vbCrLf)
        Response.Write("			  <tr>" & vbCrLf)
        Response.Write("			    <td colspan=3 height=10></td>" & vbCrLf)
        Response.Write("			  </tr>" & vbCrLf)
        Response.Write("			</table>" & vbCrLf)
        Response.Write("    </td>" & vbCrLf)
        Response.Write("  </tr>" & vbCrLf)
        Response.Write("</table>" & vbCrLf)
        Response.Write("	</BODY>" & vbCrLf)
        Response.Write("<HTML>" & vbCrLf)

        Return RedirectToAction("confirmok")

      End If

      cmdSave = Nothing

    End Function


    Function util_validate_customreports() As ActionResult
      Return View()
    End Function

    Function util_validate_crosstabs() As ActionResult
      Return View()
    End Function

#End Region

#Region "Expression Builder"

    Function util_def_expression() As ActionResult
      Return PartialView()
    End Function

    <HttpPost()>
    Function util_def_expression_submit(value As FormCollection)

      Dim objExpression
      Dim iExprType As Integer
      Dim iReturnType As Integer
      Dim sUtilType As String
      Dim sUtilType2 As String
      Dim fok As Boolean
      Dim cmdMakeHidden
      Dim prmUtilType
      Dim prmUtilID

      On Error Resume Next

      ' Get the server DLL to save the expression definition
      objExpression = CreateObject("COAIntServer.Expression")

      ' Pass required info to the DLL
      objExpression.Username = Session("username")
      CallByName(objExpression, "Connection", CallType.Let, Session("databaseConnection"))

      If Request.Form("txtSend_type") = 11 Then
        iExprType = 11
        iReturnType = 3
        sUtilType = "Filter"
        sUtilType2 = "filter"
      Else
        iExprType = 10
        iReturnType = 0
        sUtilType = "Calculation"
        sUtilType2 = "calculation"
      End If

      fok = objExpression.Initialise(CLng(Request.Form("txtSend_tableID")), _
        CLng(Request.Form("txtSend_ID")), CInt(iExprType), CInt(iReturnType))

      If fok Then
        fok = objExpression.SetExpressionDefinition(CStr(Request.Form("txtSend_components1")), _
          "", "", "", "", CStr(Request.Form("txtSend_names")))
      End If

      If fok Then
        fok = objExpression.SaveExpression(CStr(Request.Form("txtSend_name")), _
          CStr(Request.Form("txtSend_userName")), _
          CStr(Request.Form("txtSend_access")), _
          CStr(Request.Form("txtSend_description")))

        If fok Then
          If (Request.Form("txtSend_access") = "HD") And _
            (Request.Form("txtSend_ID") > 0) Then
            ' Hide any utilities that use this filter/calc.
            ' NB. The check to see if we can do this has already been done as part of the filter/calc validation. */
            cmdMakeHidden = CreateObject("ADODB.Command")
            cmdMakeHidden.CommandText = "sp_ASRIntMakeUtilitiesHidden"
            cmdMakeHidden.CommandType = 4 ' Stored procedure
            cmdMakeHidden.ActiveConnection = Session("databaseConnection")
            cmdMakeHidden.CommandTimeout = 180

            prmUtilType = cmdMakeHidden.CreateParameter("UtilType", 3, 1) ' 3 = integer, 1 = input
            cmdMakeHidden.Parameters.Append(prmUtilType)
            prmUtilType.value = CleanNumeric(Request.Form("txtSend_type"))

            prmUtilID = cmdMakeHidden.CreateParameter("UtilID", 3, 1) ' 3 = integer, 1 = input
            cmdMakeHidden.Parameters.Append(prmUtilID)
            prmUtilID.value = CleanNumeric(Request.Form("txtSend_ID"))

            Err.Clear()
            cmdMakeHidden.Execute()

            cmdMakeHidden = Nothing
          End If

          Session("confirmtext") = sUtilType & " has been saved successfully"
          Session("confirmtitle") = sUtilType & "s"
          Session("followpage") = "defsel"
          Session("reaction") = Request.Form("txtSend_reaction")
          Session("utilid") = objExpression.ExpressionID

        Else

          ' TODO ERROR REPORTING
          Response.Write("<HTML>" & vbCrLf)
          Response.Write("	<HEAD>" & vbCrLf)
          Response.Write("		<META NAME=""GENERATOR"" Content=""Microsoft Visual Studio 6.0"">" & vbCrLf)
          Response.Write("		<LINK href=""OpenHR.css"" rel=stylesheet type=text/css >" & vbCrLf)
          Response.Write("		<TITLE>" & vbCrLf)
          Response.Write("			OpenHR Intranet" & vbCrLf)
          Response.Write("		</TITLE>" & vbCrLf)
          Response.Write("  <!--#INCLUDE FILE=""include/ctl_SetStyles.txt"" -->")
          Response.Write("	</HEAD>" & vbCrLf)
          Response.Write("	<BODY id=bdyMainBody name=""bdyMainBody"" " & Session("BodyTag") & ">" & vbCrLf)

          Response.Write("	<table align=center class=""outline"" cellPadding=5 cellSpacing=0>" & vbCrLf)
          Response.Write("		<TR>" & vbCrLf)
          Response.Write("			<TD>" & vbCrLf)
          Response.Write("				<table class=""invisible"" cellspacing=0 cellpadding=0>" & vbCrLf)
          Response.Write("				  <tr> " & vbCrLf)
          Response.Write("				    <td colspan=3 height=10></td>" & vbCrLf)
          Response.Write("				  </tr>" & vbCrLf)
          Response.Write("				  <tr> " & vbCrLf)
          Response.Write("				    <td colspan=3 align=center> " & vbCrLf)
          Response.Write("							<H3>Error</H3>" & vbCrLf)
          Response.Write("				    </td>" & vbCrLf)
          Response.Write("				  </tr>" & vbCrLf)
          Response.Write("				  <tr> " & vbCrLf)
          Response.Write("				    <td width=20 height=10></td> " & vbCrLf)
          Response.Write("				    <td> " & vbCrLf)
          Response.Write("							<H4>Error saving " & sUtilType2 & "</H4>" & vbCrLf)
          Response.Write("				    </td>" & vbCrLf)
          Response.Write("				    <td width=20></td> " & vbCrLf)
          Response.Write("				  </tr>" & vbCrLf)
          Response.Write("				  <tr> " & vbCrLf)
          Response.Write("				    <td width=20 height=10></td> " & vbCrLf)
          Response.Write("				    <td> " & vbCrLf)
          Response.Write("							Unknown error" & vbCrLf)
          Response.Write("			    </td>" & vbCrLf)
          Response.Write("			    <td width=20></td> " & vbCrLf)
          Response.Write("			  </tr>" & vbCrLf)
          Response.Write("			  <tr> " & vbCrLf)
          Response.Write("			    <td colspan=3 height=20></td>" & vbCrLf)
          Response.Write("			  </tr>" & vbCrLf)
          Response.Write("			  <tr> " & vbCrLf)
          Response.Write("			    <td colspan=3 height=10 align=center>" & vbCrLf)
          Response.Write("						<INPUT TYPE=button VALUE=""Retry"" NAME=""GoBack"" class=""btn"" OnClick=""window.history.back(1)"" style=""WIDTH: 80px"" width=80 id=cmdGoBack>" & vbCrLf)
          Response.Write("                      onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
          Response.Write("                      onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
          Response.Write("		                  onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
          Response.Write("                      onblur=""try{button_onBlur(this);}catch(e){}"" />" & vbCrLf)
          Response.Write("			    </td>" & vbCrLf)
          Response.Write("			  </tr>" & vbCrLf)
          Response.Write("			  <tr>" & vbCrLf)
          Response.Write("			    <td colspan=3 height=10></td>" & vbCrLf)
          Response.Write("			  </tr>" & vbCrLf)
          Response.Write("			</table>" & vbCrLf)
          Response.Write("    </td>" & vbCrLf)
          Response.Write("  </tr>" & vbCrLf)
          Response.Write("</table>" & vbCrLf)
          Response.Write("	</BODY>" & vbCrLf)
          Response.Write("<HTML>" & vbCrLf)
        End If

      End If

      objExpression = Nothing

      'If fok Then
      'Return RedirectToAction("DefSel")
      ' Else
      'TODO - error message
      Return RedirectToAction("confirmok")
      ' End If

    End Function

    <HttpPost()>
    Function util_def_exprcomponent_submit(value As FormCollection)

      Dim sErrorMsg As String
      Dim sNextPage As String
      Dim sAction As String

      On Error Resume Next

      ' Read the information from the calling form.
      sNextPage = Request.Form("txtGotoOptionPage")
      sAction = Request.Form("txtGotoOptionAction")

      Session("optionScreenID") = Request.Form("txtGotoOptionScreenID")
      Session("optionTableID") = Request.Form("txtGotoOptionTableID")
      Session("optionViewID") = Request.Form("txtGotoOptionViewID")
      Session("optionOrderID") = Request.Form("txtGotoOptionOrderID")
      Session("optionRecordID") = Request.Form("txtGotoOptionRecordID")
      Session("optionFilterDef") = Request.Form("txtGotoOptionFilterDef")
      Session("optionFilterSQL") = Request.Form("txtGotoOptionFilterSQL")
      Session("optionValue") = Request.Form("txtGotoOptionValue")
      Session("optionLinkTableID") = Request.Form("txtGotoOptionLinkTableID")
      Session("optionLinkOrderID") = Request.Form("txtGotoOptionLinkOrderID")
      Session("optionLinkViewID") = Request.Form("txtGotoOptionLinkViewID")
      Session("optionLinkRecordID") = Request.Form("txtGotoOptionLinkRecordID")
      Session("optionColumnID") = Request.Form("txtGotoOptionColumnID")
      Session("optionLookupColumnID") = Request.Form("txtGotoOptionLookupColumnID")
      Session("optionLookupMandatory") = Request.Form("txtGotoOptionLookupMandatory")
      Session("optionLookupValue") = Request.Form("txtGotoOptionLookupValue")
      Session("optionFile") = Request.Form("txtGotoOptionFile")
      Session("optionExtension") = Request.Form("txtGotoOptionExtension")
      'Session("optionOLEOnServer") = Request.Form("txtGotoOptionOLEOnServer")
      Session("optionOLEType") = Request.Form("txtGotoOptionOLEType")
      Session("optionAction") = sAction
      Session("optionPageAction") = Request.Form("txtGotoOptionPageAction")
      Session("optionCourseTitle") = Request.Form("txtGotoOptionCourseTitle")
      Session("optionFirstRecPos") = Request.Form("txtGotoOptionFirstRecPos")
      Session("optionCurrentRecCount") = Request.Form("txtGotoOptionCurrentRecCount")
      Session("optionExprType") = Request.Form("txtGotoOptionExprType")
      Session("optionExprID") = Request.Form("txtGotoOptionExprID")
      Session("optionFunctionID") = Request.Form("txtGotoOptionFunctionID")
      Session("optionParameterIndex") = Request.Form("txtGotoOptionParameterIndex")
      Session("optionDefSelType") = Request.Form("txtGotoOptionDefSelType")
      Session("optionDefSelRecordID") = Request.Form("txtGotoOptionDefSelRecordID")

      If sAction = "CANCEL" Then
        ' Go to the requested page.
        Session("errorMessage") = sErrorMsg
      End If

      If sAction = "SELECTCOMPONENT" Then
        Session("errorMessage") = sErrorMsg
      End If

      ' Go to the requested page.
      Return RedirectToAction(sNextPage)


    End Function

    Function util_def_exprcomponent() As ActionResult
      Return PartialView()
    End Function

    Function util_test_expression() As ActionResult
      Return View()
    End Function

    Function util_test_expression_pval() As ActionResult
      Return View()
    End Function

    Function util_test_expression_submit(value As FormCollection)
      Return RedirectToAction("util_def_expression")
    End Function

    Function util_validate_expression() As ActionResult
      Return View()
    End Function

    Function util_dialog_expression() As ActionResult
      Return View()
    End Function

    Function FieldRec() As ActionResult
      Return View()
    End Function


#End Region

    Function recordEdit() As ActionResult
      Return PartialView()
    End Function

#Region "Picklists"

    Function util_def_picklist() As ActionResult
      Return PartialView()
    End Function

    Function util_def_picklist_submit(value As FormCollection)

      On Error Resume Next

      Dim cmdSave
      Dim prmName
      Dim prmDescription
      Dim prmAccess
      Dim prmUserName
      Dim prmColumns
      Dim prmColumns2
      Dim prmID
      Dim prmTableID

      cmdSave = Server.CreateObject("ADODB.Command")
      cmdSave.CommandText = "sp_ASRIntSavePicklist"
      cmdSave.CommandType = 4 ' Stored Procedure
      cmdSave.ActiveConnection = Session("databaseConnection")

      prmName = cmdSave.CreateParameter("name", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmName)
      prmName.value = Request.Form("txtSend_name")

      prmDescription = cmdSave.CreateParameter("description", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmDescription)
      prmDescription.value = Request.Form("txtSend_description")

      prmAccess = cmdSave.CreateParameter("access", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmAccess)
      prmAccess.value = Request.Form("txtSend_access")

      prmUserName = cmdSave.CreateParameter("user", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmUserName)
      prmUserName.value = Request.Form("txtSend_userName")

      prmColumns = cmdSave.CreateParameter("columns", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmColumns)
      prmColumns.value = Request.Form("txtSend_columns")

      prmColumns2 = cmdSave.CreateParameter("columns2", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmColumns2)
      prmColumns2.value = Request.Form("txtSend_columns2")

      prmID = cmdSave.CreateParameter("id", 3, 3) ' 3=integer,3=input/output
      cmdSave.Parameters.Append(prmID)
      prmID.value = CleanNumeric(Request.Form("txtSend_ID"))

      prmTableID = cmdSave.CreateParameter("tableID", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmTableID)
      prmTableID.value = CleanNumeric(Request.Form("txtSend_tableID"))

      Err.Clear()
      cmdSave.Execute()

      If Err.Number = 0 Then
        Session("confirmtext") = "Picklist has been saved successfully"
        Session("confirmtitle") = "Picklists"
        Session("followpage") = "defsel"
        Session("reaction") = Request.Form("txtSend_reaction")
        Session("utilid") = cmdSave.Parameters("id").Value

      Else
        Response.Write("<HTML>" & vbCrLf)
        Response.Write("	<HEAD>" & vbCrLf)
        Response.Write("		<META NAME=""GENERATOR"" Content=""Microsoft Visual Studio 6.0"">" & vbCrLf)
        Response.Write("		<LINK href=""OpenHR.css"" rel=stylesheet type=text/css >" & vbCrLf)
        Response.Write("		<TITLE>" & vbCrLf)
        Response.Write("			OpenHR Intranet" & vbCrLf)
        Response.Write("		</TITLE>" & vbCrLf)
        Response.Write("	</HEAD>" & vbCrLf)
        Response.Write("	<BODY id=bdyMainBody name=""bdyMainBody"" " & Session("BodyTag") & ">" & vbCrLf)

        Response.Write("	<table align=center class=""outline"" cellPadding=5 cellSpacing=0>" & vbCrLf)
        Response.Write("		<TR>" & vbCrLf)
        Response.Write("			<TD>" & vbCrLf)
        Response.Write("				<table class=""invisible"" cellspacing=0 cellpadding=0>" & vbCrLf)
        Response.Write("				  <tr> " & vbCrLf)
        Response.Write("				    <td colspan=3 height=10></td>" & vbCrLf)
        Response.Write("				  </tr>" & vbCrLf)
        Response.Write("				  <tr> " & vbCrLf)
        Response.Write("				    <td colspan=3 align=center> " & vbCrLf)
        Response.Write("							<H3>Error</H3>" & vbCrLf)
        Response.Write("				    </td>" & vbCrLf)
        Response.Write("				  </tr>" & vbCrLf)
        Response.Write("				  <tr> " & vbCrLf)
        Response.Write("				    <td width=20 height=10></td> " & vbCrLf)
        Response.Write("				    <td> " & vbCrLf)
        Response.Write("							<H4>Error saving picklist</H4>" & vbCrLf)
        Response.Write("				    </td>" & vbCrLf)
        Response.Write("				    <td width=20></td> " & vbCrLf)
        Response.Write("				  </tr>" & vbCrLf)
        Response.Write("				  <tr> " & vbCrLf)
        Response.Write("				    <td width=20 height=10></td> " & vbCrLf)
        Response.Write("				    <td> " & vbCrLf)
        Response.Write(Err.Description & vbCrLf)
        Response.Write("			    </td>" & vbCrLf)
        Response.Write("			    <td width=20></td> " & vbCrLf)
        Response.Write("			  </tr>" & vbCrLf)
        Response.Write("			  <tr> " & vbCrLf)
        Response.Write("			    <td colspan=3 height=20></td>" & vbCrLf)
        Response.Write("			  </tr>" & vbCrLf)
        Response.Write("			  <tr> " & vbCrLf)
        Response.Write("			    <td colspan=3 height=10 align=center>" & vbCrLf)
        Response.Write("						<INPUT TYPE=button VALUE=""Retry"" NAME=""GoBack"" OnClick=""window.history.back(1)"" class=""btn"" style=""WIDTH: 80px"" width=80 id=cmdGoBack>" & vbCrLf)
        Response.Write("                      onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
        Response.Write("                      onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
        Response.Write("		                  onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
        Response.Write("                      onblur=""try{button_onBlur(this);}catch(e){}"" />" & vbCrLf)
        Response.Write("			    </td>" & vbCrLf)
        Response.Write("			  </tr>" & vbCrLf)
        Response.Write("			  <tr>" & vbCrLf)
        Response.Write("			    <td colspan=3 height=10></td>" & vbCrLf)
        Response.Write("			  </tr>" & vbCrLf)
        Response.Write("			</table>" & vbCrLf)
        Response.Write("    </td>" & vbCrLf)
        Response.Write("  </tr>" & vbCrLf)
        Response.Write("</table>" & vbCrLf)
        Response.Write("	</BODY>" & vbCrLf)
        Response.Write("<HTML>" & vbCrLf)

      End If

      cmdSave = Nothing

      Return RedirectToAction("ConfirmOK")

    End Function

    Function util_dialog_picklist() As ActionResult
      Return View()
    End Function

    Function picklistSelectionMain() As ActionResult
      Return View()
    End Function

    Function picklistSelection() As ActionResult
      Return PartialView()
    End Function

    Function picklistSelectionData() As ActionResult
      Return PartialView()
    End Function

    Function picklistSelectionData_Submit(value As FormCollection)

      ' Read the information from the calling form.
      Session("tableID") = Request.Form("txtTableID")
      Session("viewID") = Request.Form("txtViewID")
      Session("orderID") = Request.Form("txtOrderID")
      Session("pageAction") = Request.Form("txtPageAction")
      Session("firstRecPos") = Request.Form("txtFirstRecPos")
      Session("currentRecCount") = Request.Form("txtCurrentRecCount")
      Session("locateValue") = Request.Form("txtGotoLocateValue")

      Session("picklistSelectionDataLoading") = False

      ' Go to the requested page.
      Return RedirectToAction("picklistSelectionData")

    End Function

#End Region

  End Class





  Public Class JsonAjaxResponse

    Public Property ErrorTitle As String
    Public Property ErrorMessage As String
    Public Property Redirect As String


  End Class
End Namespace




