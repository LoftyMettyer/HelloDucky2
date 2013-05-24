Imports System.Web.Mvc
Imports System.IO
Imports System.Web
Imports HR.Intranet.Server

Namespace Controllers

  Public Class HomeController
    Inherits Controller

#Region "Configuration"

    Function Configuration() As ActionResult
      Return View()
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

      Return RedirectToAction("confirmok")

    End Function

    Function PcConfiguration() As ActionResult
      Return View()
    End Function

#End Region

    <HttpPost()>
    Function newUser_Submit(value As FormCollection) As JsonResult
      On Error Resume Next

      Dim fSubmitNewUser = (Len(Request.Form("txtGotoPage")) = 0)

      If fSubmitNewUser Then
        ' Read the Password details from the Password form.
        Dim sNewUserLogin = Request.Form("selNewUser")

        ' Create an OpenHR user associated with the
        ' given SQL Server login.
        Dim cmdNewUser = CreateObject("ADODB.Command")
        cmdNewUser.CommandText = "sp_ASRIntNewUser"
        cmdNewUser.CommandType = 4 ' Stored Procedure
        cmdNewUser.ActiveConnection = Session("databaseConnection")

        Dim prmNewUser = cmdNewUser.CreateParameter("newUser", 200, 1, 255)
        cmdNewUser.Parameters.Append(prmNewUser)
        prmNewUser.value = sNewUserLogin

        Err.Clear()
        cmdNewUser.Execute()

        ' Release the ADO command object.
        cmdNewUser = Nothing

        If Err.Number <> 0 Then
          Session("ErrorTitle") = "New User Page"
          Session("ErrorText") = "You could not add the user because of the following error:<p>" & FormatError(Err.Description)
          Dim data1 = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = ""}
          Return Json(data1, JsonRequestBehavior.AllowGet)
          'Response.Redirect("error")
        Else
          Session("ErrorTitle") = "New User Page"
          Session("ErrorText") = "User added successfully."
          Dim data = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = ""}
          Return Json(data, JsonRequestBehavior.AllowGet)
          'Response.Redirect("message")
        End If
      Else
        ' Read the information from the calling form.
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
        ' Response.Redirect(Request.Form("txtGotoPage"))
        Session("txtGotoPage") = Request.Form("txtGotoPage")
      End If
    End Function

    <HttpPost()>
    Function passwordChange_Submit(value As FormCollection) As JsonResult

      On Error Resume Next

      Dim sReferringPage = ""
      Dim fSubmitPasswordChange = ""
      Dim sErrorText = ""

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
              Dim data = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = ""}
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
                Dim data1 = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = ""}
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
                  Dim data1 = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = ""}
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
              Session("ErrorTitle") = "Change Password Page"
              Session("ErrorText") = "Password changed successfully."

              Dim data = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = "main"}
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

    Function ConfirmOK() As ActionResult
      Return View()
    End Function

    ' GET: /Home
    Function Main() As ActionResult
      Return View()
    End Function

    Function Find(Optional sParameters As String = "") As ActionResult

      ' Additional controller actions for SSI view. Only SSI calls to this action have parameters.
      If sParameters.Length > 0 Then
        ' =========================
        ' Self-service Find request
        ' =========================
        Dim lngTopLevelRecordID
        Dim sTableName
        Dim sViewName
        'NPG20081401 Fault 12868
        Dim dblPreviousColumnWidth
        Dim objUser

        'NPG20081401 Fault 12868
        objUser = CreateObject("COAIntServer.clsSettings")
        CallByName(objUser, "Connection", CallType.Let, Session("databaseConnection"))

        Const DEADLOCK_ERRORNUMBER = -2147467259
        Const DEADLOCK_MESSAGESTART = "YOUR TRANSACTION (PROCESS ID #"
        Const DEADLOCK_MESSAGEEND = ") WAS DEADLOCKED WITH ANOTHER PROCESS AND HAS BEEN CHOSEN AS THE DEADLOCK VICTIM. RERUN YOUR TRANSACTION."
        Const DEADLOCK2_MESSAGESTART = "TRANSACTION (PROCESS ID "
        Const DEADLOCK2_MESSAGEEND = ") WAS DEADLOCKED ON "
        Const SQLMAILNOTSTARTEDMESSAGE = "SQL MAIL SESSION IS NOT STARTED."

        Dim iRETRIES = 5
        Dim iRetryCount = 0
        Dim sErrorDescription = ""

        Dim iRealTableID = 0
        Dim iRealViewID = 0

        lngTopLevelRecordID = Session("TopLevelRecID")

        If CLng(Session("tableType")) <> 2 Then
          ' Top Level table.
          'Response.Write "#<FONT COLOR='Red'><B>Top Level table.</B></FONT>#<BR>"

          Session("recordID") = lngTopLevelRecordID
          Session("parentTableID") = 0
          Session("parentRecordID") = 0
        Else
          ' Child table.
          ' Response.Write "#<FONT COLOR='Red'><B>Child table.</B></FONT>#<BR>"

          iRealTableID = Session("SSILinkTableID")
          iRealViewID = Session("SSILinkViewID")
          'session("tableID") = 0 
          Session("viewID") = 0
          Session("parentTableID") = Session("SSILinkTableID")
          Session("parentRecordID") = lngTopLevelRecordID
        End If

        ' Read the screen info from the query string.			

        'Response.Write "#<FONT COLOR='Red'><B>sParameters = " & sParameters & "</B></FONT>#<BR>"
        'Response.Write "#<FONT COLOR='Red'><B>parentTableID = " & session("parentTableID") & "</B></FONT>#<BR>"
        'Response.Write "#<FONT COLOR='Red'><B>parentRecordID = " & session("parentRecordID") & "</B></FONT>#<BR>"

        Session("action") = Left(sParameters, InStr(sParameters, "_") - 1)
        sParameters = Mid(sParameters, InStr(sParameters, "_") + 1)
        Session("firstRecPos") = Left(sParameters, InStr(sParameters, "_") - 1)
        sParameters = Mid(sParameters, InStr(sParameters, "_") + 1)
        Session("currentRecCount") = Left(sParameters, InStr(sParameters, "_") - 1)
        Session("locateValue") = Mid(sParameters, InStr(sParameters, "_") + 1)

        ' Flag an error if there is no current table or view is specified.
        If (Session("tableID") <= 0) Then
          'and (session("viewID") <= 0) then

          sErrorDescription = "The find page could not be loaded." & vbCrLf & "No table or view specified."
        End If

        If Len(sErrorDescription) = 0 Then
          ' Flag an error if there is no current screen is specified.
          If (Session("linkType") <> "multifind") And _
            (Session("screenID") <= 0) Then
            sErrorDescription = "The find page could not be loaded." & vbCrLf & "No screen specified."
          End If
        End If

        If Len(sErrorDescription) = 0 Then
          If (Session("linkType") = "multifind") Then
            Dim cmdOrder = CreateObject("ADODB.Command")
            cmdOrder.CommandText = "spASRIntGetDefaultOrder"
            cmdOrder.CommandType = 4 ' Stored Procedure
            cmdOrder.ActiveConnection = Session("databaseConnection")

            Dim prmTableID = cmdOrder.CreateParameter("tableID", 3, 1)
            cmdOrder.Parameters.Append(prmTableID)
            prmTableID.value = CleanNumeric(Session("tableID"))

            Dim prmOrderID = cmdOrder.CreateParameter("orderID", 3, 2)
            cmdOrder.Parameters.Append(prmOrderID)

            Err.Clear()
            cmdOrder.Execute()
            If (Err.Number <> 0) Then
              sErrorDescription = "The find page could not be loaded." & vbCrLf & "The default order for the table could not be determined :" & vbCrLf & FormatError(Err.Description)
            Else
              Session("orderID") = cmdOrder.Parameters("orderID").Value
            End If
            ' Release the ADO command object.
            cmdOrder = Nothing
          Else
            ' Get the screen's default order if none is already specified.
            Dim cmdScreenOrder = CreateObject("ADODB.Command")
            cmdScreenOrder.CommandText = "sp_ASRIntGetScreenOrder"
            cmdScreenOrder.CommandType = 4 ' Stored Procedure
            cmdScreenOrder.ActiveConnection = Session("databaseConnection")

            Dim prmOrderID = cmdScreenOrder.CreateParameter("orderID", 3, 2)
            cmdScreenOrder.Parameters.Append(prmOrderID)

            Dim prmScreenID = cmdScreenOrder.CreateParameter("screenID", 3, 1)
            cmdScreenOrder.Parameters.Append(prmScreenID)
            prmScreenID.value = CleanNumeric(Session("screenID"))

            Err.Clear()
            cmdScreenOrder.Execute()
            If (Err.Number <> 0) Then
              sErrorDescription = "The find page could not be loaded." & vbCrLf & "The default order for the table could not be determined :" & vbCrLf & FormatError(Err.Description)
            Else
              Session("orderID") = cmdScreenOrder.Parameters("orderID").Value
            End If
            ' Release the ADO command object.
            cmdScreenOrder = Nothing
          End If
        End If

        'Response.Write "#<FONT COLOR='Red'>session(SSILinkViewID) = <B>" & session("SSILinkViewID") & "</B></FONT>#<BR>"
        'Response.Write "#<FONT COLOR='Red'>session(SSILinkTableID) = <B>" & session("SSILinkTableID") & "</B></FONT>#<BR>"
        'Response.Write "#<FONT COLOR='Red'>session(PersonnelTableID) = <B>" & session("PersonnelTableID") & "</B></FONT>#<BR>"
        'Response.Write "#<FONT COLOR='Red'>session(TopLevelRecID) = <B>" & session("TopLevelRecID") & "</B></FONT>#<BR>"
        'Response.Write "#<FONT COLOR='Red'>session(SingleRecordViewID) = <B>" & session("SingleRecordViewID") & "</B></FONT>#<BR>"
        'Response.Write "#<FONT COLOR='Red'>session(tableID) = <B>" & session("tableID") & "</B></FONT>#<BR>"
        'Response.Write "#<FONT COLOR='Red'>session(viewID) = <B>" & session("viewID") & "</B></FONT>#<BR>"

        If Len(sErrorDescription) = 0 Then

          If CLng(Session("SSILinkViewID")) = CLng(Session("SingleRecordViewID")) Then
            lngTopLevelRecordID = Session("TopLevelRecID")
          End If

          If CLng(Session("tableType")) <> 2 Then
            ' Top Level table.
            Session("recordID") = 0 '  lngPersonnelRecordID			' never set???
            Session("parentTableID") = 0
            Session("parentRecordID") = 0
          Else
            ' Child table.
            Session("parentTableID") = Session("SSILinkTableID")
            Session("parentRecordID") = lngTopLevelRecordID
          End If

          ' Enable response buffering as we may redirect the response further down this page.
          Response.Buffer = True
        End If

        Dim sRecDesc = ""
        If CLng(Session("SSILinkViewID")) <> CLng(Session("SingleRecordViewID")) And _
          (Len(sErrorDescription) = 0) Then


          Dim cmdGetRecordDesc = CreateObject("ADODB.Command")
          cmdGetRecordDesc.CommandText = "spASRIntGetRecordDescriptionInView"
          cmdGetRecordDesc.CommandType = 4 ' Stored procedure
          cmdGetRecordDesc.ActiveConnection = Session("databaseConnection")

          Dim prmViewID = cmdGetRecordDesc.CreateParameter("viewID", 3, 1) ' 3 = integer, 1 = input
          cmdGetRecordDesc.Parameters.Append(prmViewID)
          prmViewID.value = CleanNumeric(Session("SSILinkViewID"))

          Dim prmTableID = cmdGetRecordDesc.CreateParameter("tableID", 3, 1) ' 3 = integer, 1 = input
          cmdGetRecordDesc.Parameters.Append(prmTableID)
          prmTableID.value = CleanNumeric(Session("tableID"))

          Dim prmRecordID = cmdGetRecordDesc.CreateParameter("recordID", 3, 1) ' 3 = integer, 1 = input
          cmdGetRecordDesc.Parameters.Append(prmRecordID)
          prmRecordID.value = 0

          Dim prmParentTableID = cmdGetRecordDesc.CreateParameter("parentTableID", 3, 1) ' 3 = integer, 1 = input
          cmdGetRecordDesc.Parameters.Append(prmParentTableID)
          prmParentTableID.value = CleanNumeric(Session("parentTableID"))

          Dim prmParentRecordID = cmdGetRecordDesc.CreateParameter("parentRecordID", 3, 1) ' 3=integer, 1=input
          cmdGetRecordDesc.Parameters.Append(prmParentRecordID)
          prmParentRecordID.value = CleanNumeric(Session("parentRecordID"))

          Dim prmRecordDesc = cmdGetRecordDesc.CreateParameter("recordDesc", 200, 2, 2147483646)
          cmdGetRecordDesc.Parameters.Append(prmRecordDesc)

          Dim prmErrorMessage = cmdGetRecordDesc.CreateParameter("errorMessage", 200, 2, 2147483646)
          cmdGetRecordDesc.Parameters.Append(prmErrorMessage)

          Dim fOK = True
          Dim fDeadlock = True
          Do While fDeadlock
            fDeadlock = False

            cmdGetRecordDesc.ActiveConnection.Errors.Clear()

            cmdGetRecordDesc.Execute()
            Dim sErrMsg As String

            If cmdGetRecordDesc.ActiveConnection.Errors.Count > 0 Then
              For iLoop = 1 To cmdGetRecordDesc.ActiveConnection.Errors.Count
                sErrMsg = FormatError(cmdGetRecordDesc.ActiveConnection.Errors.Item(iLoop - 1).Description)

                If (cmdGetRecordDesc.ActiveConnection.Errors.Item(iLoop - 1).Number = DEADLOCK_ERRORNUMBER) And _
                  (((UCase(Left(sErrMsg, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And _
                    (UCase(Right(sErrMsg, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or _
                  ((UCase(Left(sErrMsg, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And _
                    (InStr(UCase(sErrMsg), DEADLOCK2_MESSAGEEND) > 0))) Then
                  ' The error is for a deadlock.
                  ' Sorry about having to use the err.description to trap the error but the err.number
                  ' is not specific and MSDN suggests using the err.description.
                  If (iRetryCount < iRETRIES) And (cmdGetRecordDesc.ActiveConnection.Errors.Count = 1) Then
                    iRetryCount = iRetryCount + 1
                    fDeadlock = True
                  Else
                    If Len(sErrorDescription) > 0 Then
                      sErrorDescription = sErrorDescription & vbCrLf
                    End If
                    sErrorDescription = sErrorDescription & "Another user is deadlocking the database. Please try again."
                    fOK = False
                  End If
                Else
                  sErrorDescription = sErrorDescription & vbCrLf & _
                    FormatError(cmdGetRecordDesc.ActiveConnection.Errors.Item(iLoop - 1).Description)
                  fOK = False
                End If
              Next

              cmdGetRecordDesc.ActiveConnection.Errors.Clear()

              If Not fOK Then
                sErrorDescription = "Unable to get the record description." & vbCrLf & sErrorDescription
              End If
            End If
          Loop

          If (Len(sErrorDescription) = 0) Then
            If (Len(cmdGetRecordDesc.Parameters("errorMessage").Value) > 0) Then
              sErrorDescription = "Unable to get the record description." & vbCrLf & cmdGetRecordDesc.Parameters("errorMessage").Value
            Else
              sRecDesc = cmdGetRecordDesc.Parameters("recordDesc").Value
            End If
          End If

          cmdGetRecordDesc = Nothing
        End If

        If (Len(sErrorDescription) = 0) Then
          Dim sTitle As String = ""
          If (Session("linkType") <> "multifind") Then
            Dim cmdGetTableName = CreateObject("ADODB.Command")
            cmdGetTableName.CommandText = "sp_ASRIntGetTableName"
            cmdGetTableName.CommandType = 4 ' Stored procedure
            cmdGetTableName.ActiveConnection = Session("databaseConnection")

            Dim prmTableID = cmdGetTableName.CreateParameter("TableID", 3, 1)
            cmdGetTableName.Parameters.Append(prmTableID)
            prmTableID.value = CleanNumeric(Session("tableID"))

            Dim prmTableName = cmdGetTableName.CreateParameter("TableName", 200, 2, 255)
            cmdGetTableName.Parameters.Append(prmTableName)

            Err.Clear()
            cmdGetTableName.Execute()

            If (Err.Number <> 0) Then
              sErrorDescription = "Error getting the link table name." & vbCrLf & FormatError(Err.Description)
            Else
              sTableName = Replace(cmdGetTableName.Parameters("TableName").Value, "_", " ")
            End If

            cmdGetTableName = Nothing

            sTitle = "Select the required "

            If Len(sTableName) > 0 Then
              sTitle = sTitle & sTableName & " "
            End If

            sTitle = sTitle & "record"

            If Len(sRecDesc) > 0 Then
              sTitle = sTitle & " for " & sRecDesc
            End If
          Else
            Dim cmdGetPageTitle = CreateObject("ADODB.Command")
            cmdGetPageTitle.CommandText = "spASRIntGetPageTitle"
            cmdGetPageTitle.CommandType = 4 ' Stored procedure
            cmdGetPageTitle.ActiveConnection = Session("databaseConnection")

            Dim prmTableID = cmdGetPageTitle.CreateParameter("TableID", 3, 1)
            cmdGetPageTitle.Parameters.Append(prmTableID)
            prmTableID.value = CleanNumeric(Session("TableID"))

            Dim prmViewID = cmdGetPageTitle.CreateParameter("ViewID", 3, 1)
            cmdGetPageTitle.Parameters.Append(prmViewID)
            prmViewID.value = CleanNumeric(Session("ViewID"))

            Dim prmPageTitle = cmdGetPageTitle.CreateParameter("PageTitle", 200, 2, 200) ' 200=varchar, 2=output, 200=size
            cmdGetPageTitle.Parameters.Append(prmPageTitle)
            Err.Clear()
            cmdGetPageTitle.Execute()

            If (Err.Number <> 0) Then
              sErrorDescription = "Error getting the page title." & vbCrLf & FormatError(Err.Description)
            Else
              sTitle = Replace(cmdGetPageTitle.Parameters("PageTitle").Value, "_", " ")
            End If

            cmdGetPageTitle = Nothing
          End If

          sTitle = Server.UrlEncode(sTitle)
        End If

        If (Len(sErrorDescription) = 0) Then

          If CLng(Session("SSILinkViewID")) > -1 Then

            Dim cmdGetViewName = CreateObject("ADODB.Command")
            cmdGetViewName.CommandText = "spASRIntGetViewName"
            cmdGetViewName.CommandType = 4 ' Stored procedure
            cmdGetViewName.ActiveConnection = Session("databaseConnection")

            Dim prmViewID = cmdGetViewName.CreateParameter("ViewID", 3, 1)
            cmdGetViewName.Parameters.Append(prmViewID)

            If CLng(Session("SSILinkViewID")) <> CLng(Session("SingleRecordViewID")) And _
              (Session("linkType") <> "multifind") Then
              prmViewID.value = CleanNumeric(Session("SSILinkViewID"))
            Else
              prmViewID.value = CleanNumeric(Session("SingleRecordViewID"))
            End If

            Dim prmViewName = cmdGetViewName.CreateParameter("ViewName", 200, 2, 255)
            cmdGetViewName.Parameters.Append(prmViewName)

            Err.Clear()
            cmdGetViewName.Execute()

            If (Err.Number <> 0) Then
              sErrorDescription = "Error getting the link view name." & vbCrLf & FormatError(Err.Description)
            Else
              If Not IsDBNull(cmdGetViewName.Parameters("ViewName").Value) Then
                sViewName = Replace(cmdGetViewName.Parameters("ViewName").Value, "_", " ")
              Else
                sViewName = ""
              End If
            End If

            cmdGetViewName = Nothing

          Else

            Dim cmdGetTableName = CreateObject("ADODB.Command")
            cmdGetTableName.CommandText = "sp_ASRIntGetTableName"
            cmdGetTableName.CommandType = 4 ' Stored procedure
            cmdGetTableName.ActiveConnection = Session("databaseConnection")

            Dim prmTableID = cmdGetTableName.CreateParameter("TableID", 3, 1)
            cmdGetTableName.Parameters.Append(prmTableID)
            prmTableID.value = CleanNumeric(Session("SSILinkTableID"))

            Dim prmTableName = cmdGetTableName.CreateParameter("TableName", 200, 2, 255)
            cmdGetTableName.Parameters.Append(prmTableName)

            Err.Clear()
            cmdGetTableName.Execute()

            If (Err.Number <> 0) Then
              sErrorDescription = "Error getting the link table name." & vbCrLf & FormatError(Err.Description)
            Else
              If Not IsDBNull(cmdGetTableName.Parameters("TableName").Value) Then
                sTableName = Replace(cmdGetTableName.Parameters("TableName").Value, "_", " ")
              Else
                sTableName = ""
              End If
            End If

            cmdGetTableName = Nothing

          End If

          If (CLng(Session("SSILinkViewID")) = CLng(Session("SingleRecordViewID")) Or _
            (Session("linkType") = "multifind")) And _
            Session("SingleRecordViewID") = 0 Then

            sViewName = "single record"
          End If
        End If
      End If

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
    Function MailMerge_Submit()
      On Error Resume Next

      Dim cmdSave = CreateObject("ADODB.Command")
      cmdSave.CommandText = "sp_ASRIntSaveMailMerge"
      cmdSave.CommandType = 4 ' Stored Procedure
      cmdSave.ActiveConnection = Session("databaseConnection")

      Dim prmName = cmdSave.CreateParameter("name", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmName)
      prmName.value = Request.Form("txtSend_name")

      Dim prmDescription = cmdSave.CreateParameter("description", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmDescription)
      prmDescription.value = Request.Form("txtSend_description")

      Dim prmTableID = cmdSave.CreateParameter("tableID", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmTableID)
      prmTableID.value = CleanNumeric(Request.Form("txtSend_baseTable"))

      Dim prmSelection = cmdSave.CreateParameter("selection", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmSelection)
      prmSelection.value = CleanNumeric(Request.Form("txtSend_selection"))

      Dim prmPicklistID = cmdSave.CreateParameter("picklistID", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmPicklistID)
      prmPicklistID.value = CleanNumeric(Request.Form("txtSend_picklist"))

      Dim prmFilterID = cmdSave.CreateParameter("filterID", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmFilterID)
      prmFilterID.value = CleanNumeric(Request.Form("txtSend_filter"))

      Dim prmOutputFormat = cmdSave.CreateParameter("outputFormat", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmOutputFormat)
      prmOutputFormat.value = CleanNumeric(Request.Form("txtSend_outputformat"))

      Dim prmOutputSave = cmdSave.CreateParameter("outputSave", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmOutputSave)
      prmOutputSave.value = CleanBoolean(Request.Form("txtSend_outputsave"))

      Dim prmOutputFileName = cmdSave.CreateParameter("outputFileName", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmOutputFileName)
      prmOutputFileName.value = Request.Form("txtSend_outputfilename")

      Dim prmEmailAddrID = cmdSave.CreateParameter("emailAddrID", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmEmailAddrID)
      prmEmailAddrID.value = CleanNumeric(Request.Form("txtSend_emailaddrid"))

      Dim prmEmailSubject = cmdSave.CreateParameter("emailSubject", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmEmailSubject)
      prmEmailSubject.value = Request.Form("txtSend_emailsubject")

      Dim prmTemplateFileName = cmdSave.CreateParameter("templateFileName", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmTemplateFileName)
      prmTemplateFileName.value = Request.Form("txtSend_templatefilename")

      Dim prmOutputScreen = cmdSave.CreateParameter("outputScreen", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmOutputScreen)
      prmOutputScreen.value = CleanBoolean(Request.Form("txtSend_outputscreen"))

      Dim prmUserName = cmdSave.CreateParameter("userName", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmUserName)
      prmUserName.value = Request.Form("txtSend_userName")

      Dim prmEmailAsAttachment = cmdSave.CreateParameter("emailAsAttachment", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmEmailAsAttachment)
      prmEmailAsAttachment.value = CleanBoolean(Request.Form("txtSend_emailasattachment"))

      Dim prmEmailAttachmentName = cmdSave.CreateParameter("emailAttachmentName", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmEmailAttachmentName)
      prmEmailAttachmentName.value = Request.Form("txtSend_emailattachmentname")

      Dim prmSuppressBlanks = cmdSave.CreateParameter("suppressBlanks", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmSuppressBlanks)
      prmSuppressBlanks.value = CleanBoolean(Request.Form("txtSend_suppressblanks"))

      Dim prmPauseBeforeMerge = cmdSave.CreateParameter("pauseBeforeMerge", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmPauseBeforeMerge)
      prmPauseBeforeMerge.value = CleanBoolean(Request.Form("txtSend_pausebeforemerge"))

      Dim prmOutputPrinter = cmdSave.CreateParameter("outputPrinter", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmOutputPrinter)
      prmOutputPrinter.value = CleanBoolean(Request.Form("txtSend_outputprinter"))

      Dim prmOutputPrinterName = cmdSave.CreateParameter("outputPrinterName", 200, 1, 255) ' 200=varchar,1=input,255=size
      cmdSave.Parameters.Append(prmOutputPrinterName)
      prmOutputPrinterName.value = Request.Form("txtSend_outputprintername")

      Dim prmDocumentMapID = cmdSave.CreateParameter("documentMapID", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmDocumentMapID)
      prmDocumentMapID.value = CleanNumeric(Request.Form("txtSend_documentmapid"))

      Dim prmManualDocManHeader = cmdSave.CreateParameter("manualDocManHeader", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmManualDocManHeader)
      prmManualDocManHeader.value = CleanBoolean(Request.Form("txtSend_manualdocmanheader"))

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

      Dim prmID = cmdSave.CreateParameter("id", 3, 3) ' 3=integer,3=input/output
      cmdSave.Parameters.Append(prmID)
      prmID.value = CleanNumeric(Request.Form("txtSend_ID"))

      cmdSave.Execute()

      If Err.Number = 0 Then
        Session("confirmtext") = "Mail Merge has been saved successfully"
        Session("confirmtitle") = "Mail Merge"
        Session("followpage") = "defsel"
        Session("reaction") = Request.Form("txtSend_reaction")
        Session("utilid") = cmdSave.Parameters("id").Value

        Response.Redirect("confirmok")
      Else
        Response.Write("<HTML>" & vbCrLf)
        Response.Write("	<HEAD>" & vbCrLf)
        Response.Write("		<META NAME=""GENERATOR"" Content=""Microsoft Visual Studio 6.0"">" & vbCrLf)
        Response.Write("		<LINK href=""OpenHR.css"" rel=stylesheet type=text/css >" & vbCrLf)
        Response.Write("		<TITLE>" & vbCrLf)
        Response.Write("			OpenHR Intranet" & vbCrLf)
        Response.Write("		</TITLE>" & vbCrLf)
        Response.Write("		<meta http-equiv=""X-UA-Compatible"" content=""IE=5"">" & vbCrLf)
        Response.Write("  <!--#INCLUDE FILE=""include/ctl_SetStyles.txt"" -->")
        Response.Write("	</HEAD>" & vbCrLf)
        Response.Write("	<BODY>" & vbCrLf)
        Response.Write("Error saving definition : <BR>" & Err.Description & "<BR>" & vbCrLf)
        Response.Write("<INPUT TYPE=button VALUE=Retry NAME=GoBack OnClick=" & Chr(34) & "window.history.back(1)" & Chr(34) & " class=""btn"" style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & " width=100 id=cmdGoBack>")
        Response.Write("                      onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
        Response.Write("                      onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
        Response.Write("		                  onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
        Response.Write("                      onblur=""try{button_onBlur(this);}catch(e){}"" />" & vbCrLf)
        'Response.Write(vbCrLf & vbCrLf & sSQLString)
        Response.Write("	</BODY>" & vbCrLf)
        Response.Write("<HTML>" & vbCrLf)
      End If

      cmdSave = Nothing
      '%>	

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
            'Case 25	' WORKFLOW 
            'Return RedirectToAction("util_run_workflow")
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
            'Case 25	' WORKFLOW 
            '	Session("reaction") = "WORKFLOWS"
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

    Function util_delete() As ActionResult
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

    Function Util_CalcSelection() As ActionResult
      Return View()
    End Function

    Function Util_SortOrderSelection() As ActionResult
      Return View()
    End Function

    Function LinksMain() As ActionResult
      If Session("objButtonInfo") Is Nothing Or Session("objHypertextInfo") Is Nothing Or Session("objDropdownInfo") Is Nothing Then
        Return RedirectToAction("Login", "Account")
      End If

      Dim objHypertextInfo As Collection = Session("objHypertextInfo")
      Dim objButtonInfo As Collection = Session("objButtonInfo")
      Dim objDropdownInfo As Collection = Session("objDropdownInfo")

      Dim lstButtonInfo = (From collectionItem As Object In objHypertextInfo Select New navigationLink(collectionItem.ID, collectionItem.DrillDownHidden, collectionItem.LinkType, collectionItem.LinkOrder, collectionItem.Text, collectionItem.Text1, collectionItem.Text2, collectionItem.Prompt, collectionItem.ScreenID, collectionItem.TableID, collectionItem.ViewID, collectionItem.PageTitle, collectionItem.URL, collectionItem.UtilityType, collectionItem.UtilityID, collectionItem.NewWindow, collectionItem.BaseTable, collectionItem.LinkToFind, collectionItem.SingleRecord, collectionItem.PrimarySequence, collectionItem.SecondarySequence, collectionItem.FindPage, collectionItem.EmailAddress, collectionItem.EmailSubject, collectionItem.AppFilePath, collectionItem.AppParameters, collectionItem.DocumentFilePath, collectionItem.DisplayDocumentHyperlink, collectionItem.IsSeparator, collectionItem.Element_Type, collectionItem.SeparatorOrientation, collectionItem.PictureID, collectionItem.Chart_ShowLegend, collectionItem.Chart_Type, collectionItem.Chart_ShowGrid, collectionItem.Chart_StackSeries, collectionItem.Chart_ShowValues, collectionItem.Chart_ViewID, collectionItem.Chart_TableID, collectionItem.Chart_ColumnID, collectionItem.Chart_FilterID, collectionItem.Chart_AggregateType, collectionItem.Chart_ColumnName, collectionItem.Chart_ColumnName_2, collectionItem.UseFormatting, collectionItem.Formatting_DecimalPlaces, collectionItem.Formatting_Use1000Separator, collectionItem.Formatting_Prefix, collectionItem.Formatting_Suffix, collectionItem.UseConditionalFormatting, collectionItem.ConditionalFormatting_Operator_1, collectionItem.ConditionalFormatting_Value_1, collectionItem.ConditionalFormatting_Style_1, collectionItem.ConditionalFormatting_Colour_1, collectionItem.ConditionalFormatting_Operator_2, collectionItem.ConditionalFormatting_Value_2, collectionItem.ConditionalFormatting_Style_2, collectionItem.ConditionalFormatting_Colour_2, collectionItem.ConditionalFormatting_Operator_3, collectionItem.ConditionalFormatting_Value_3, collectionItem.ConditionalFormatting_Style_3, collectionItem.ConditionalFormatting_Colour_3, collectionItem.SeparatorColour, collectionItem.InitialDisplayMode, collectionItem.Chart_TableID_2, collectionItem.Chart_ColumnID_2, collectionItem.Chart_TableID_3, collectionItem.Chart_ColumnID_3, collectionItem.Chart_SortOrderID, collectionItem.Chart_SortDirection, collectionItem.Chart_ColourID, collectionItem.Chart_ShowPercentages)).ToList()
      lstButtonInfo.AddRange(From collectionItem As Object In objButtonInfo Select New navigationLink(collectionItem.ID, collectionItem.DrillDownHidden, collectionItem.LinkType, collectionItem.LinkOrder, collectionItem.Text, collectionItem.Text1, collectionItem.Text2, collectionItem.Prompt, collectionItem.ScreenID, collectionItem.TableID, collectionItem.ViewID, collectionItem.PageTitle, collectionItem.URL, collectionItem.UtilityType, collectionItem.UtilityID, collectionItem.NewWindow, collectionItem.BaseTable, collectionItem.LinkToFind, collectionItem.SingleRecord, collectionItem.PrimarySequence, collectionItem.SecondarySequence, collectionItem.FindPage, collectionItem.EmailAddress, collectionItem.EmailSubject, collectionItem.AppFilePath, collectionItem.AppParameters, collectionItem.DocumentFilePath, collectionItem.DisplayDocumentHyperlink, collectionItem.IsSeparator, collectionItem.Element_Type, collectionItem.SeparatorOrientation, collectionItem.PictureID, collectionItem.Chart_ShowLegend, collectionItem.Chart_Type, collectionItem.Chart_ShowGrid, collectionItem.Chart_StackSeries, collectionItem.Chart_ShowValues, collectionItem.Chart_ViewID, collectionItem.Chart_TableID, collectionItem.Chart_ColumnID, collectionItem.Chart_FilterID, collectionItem.Chart_AggregateType, collectionItem.Chart_ColumnName, collectionItem.Chart_ColumnName_2, collectionItem.UseFormatting, collectionItem.Formatting_DecimalPlaces, collectionItem.Formatting_Use1000Separator, collectionItem.Formatting_Prefix, collectionItem.Formatting_Suffix, collectionItem.UseConditionalFormatting, collectionItem.ConditionalFormatting_Operator_1, collectionItem.ConditionalFormatting_Value_1, collectionItem.ConditionalFormatting_Style_1, collectionItem.ConditionalFormatting_Colour_1, collectionItem.ConditionalFormatting_Operator_2, collectionItem.ConditionalFormatting_Value_2, collectionItem.ConditionalFormatting_Style_2, collectionItem.ConditionalFormatting_Colour_2, collectionItem.ConditionalFormatting_Operator_3, collectionItem.ConditionalFormatting_Value_3, collectionItem.ConditionalFormatting_Style_3, collectionItem.ConditionalFormatting_Colour_3, collectionItem.SeparatorColour, collectionItem.InitialDisplayMode, collectionItem.Chart_TableID_2, collectionItem.Chart_ColumnID_2, collectionItem.Chart_TableID_3, collectionItem.Chart_ColumnID_3, collectionItem.Chart_SortOrderID, collectionItem.Chart_SortDirection, collectionItem.Chart_ColourID, collectionItem.Chart_ShowPercentages))
      lstButtonInfo.AddRange(From collectionItem As Object In objDropdownInfo Select New navigationLink(collectionItem.ID, collectionItem.DrillDownHidden, collectionItem.LinkType, collectionItem.LinkOrder, collectionItem.Text, collectionItem.Text1, collectionItem.Text2, collectionItem.Prompt, collectionItem.ScreenID, collectionItem.TableID, collectionItem.ViewID, collectionItem.PageTitle, collectionItem.URL, collectionItem.UtilityType, collectionItem.UtilityID, collectionItem.NewWindow, collectionItem.BaseTable, collectionItem.LinkToFind, collectionItem.SingleRecord, collectionItem.PrimarySequence, collectionItem.SecondarySequence, collectionItem.FindPage, collectionItem.EmailAddress, collectionItem.EmailSubject, collectionItem.AppFilePath, collectionItem.AppParameters, collectionItem.DocumentFilePath, collectionItem.DisplayDocumentHyperlink, collectionItem.IsSeparator, collectionItem.Element_Type, collectionItem.SeparatorOrientation, collectionItem.PictureID, collectionItem.Chart_ShowLegend, collectionItem.Chart_Type, collectionItem.Chart_ShowGrid, collectionItem.Chart_StackSeries, collectionItem.Chart_ShowValues, collectionItem.Chart_ViewID, collectionItem.Chart_TableID, collectionItem.Chart_ColumnID, collectionItem.Chart_FilterID, collectionItem.Chart_AggregateType, collectionItem.Chart_ColumnName, collectionItem.Chart_ColumnName_2, collectionItem.UseFormatting, collectionItem.Formatting_DecimalPlaces, collectionItem.Formatting_Use1000Separator, collectionItem.Formatting_Prefix, collectionItem.Formatting_Suffix, collectionItem.UseConditionalFormatting, collectionItem.ConditionalFormatting_Operator_1, collectionItem.ConditionalFormatting_Value_1, collectionItem.ConditionalFormatting_Style_1, collectionItem.ConditionalFormatting_Colour_1, collectionItem.ConditionalFormatting_Operator_2, collectionItem.ConditionalFormatting_Value_2, collectionItem.ConditionalFormatting_Style_2, collectionItem.ConditionalFormatting_Colour_2, collectionItem.ConditionalFormatting_Operator_3, collectionItem.ConditionalFormatting_Value_3, collectionItem.ConditionalFormatting_Style_3, collectionItem.ConditionalFormatting_Colour_3, collectionItem.SeparatorColour, collectionItem.InitialDisplayMode, collectionItem.Chart_TableID_2, collectionItem.Chart_ColumnID_2, collectionItem.Chart_TableID_3, collectionItem.Chart_ColumnID_3, collectionItem.Chart_SortOrderID, collectionItem.Chart_SortDirection, collectionItem.Chart_ColourID, collectionItem.Chart_ShowPercentages))

      Dim viewModel = New NavLinksViewModel With {.NavigationLinks = lstButtonInfo, .NumberOfLinks = objDropdownInfo.Count}

      Session("SSILinkTableID") = Session("SingleRecordTableID")
      Session("SSILinkViewID") = Session("SingleRecordViewID")

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

    Public Sub ShowImageFromDb(imageID As String)

      imageID = CleanNumeric(imageID)

      ' Get the required picture using PictureID.
      Dim cmdReadPicture = CreateObject("ADODB.Command")
      cmdReadPicture.CommandText = "spASRIntGetPicture"
      cmdReadPicture.CommandType = 4 ' Stored Procedure
      cmdReadPicture.ActiveConnection = Session("databaseConnection")

      Dim prmPictureID = cmdReadPicture.CreateParameter("pictureid", 3, 1) ' 3=integer, 1=input
      cmdReadPicture.Parameters.Append(prmPictureID)
      prmPictureID.value = imageID

      Err.Clear()
      Dim objRs = cmdReadPicture.Execute

      If (Err.Number <> 0) Then
        Response.End()
      End If

      Dim image(-1) As Byte

      Do While Not objRs.EOF
        image = CType(objRs.fields(1).value, Byte())

        If image.Length > 0 Then Exit Do

        objRs.moveNext()
      Loop

      If image Is Nothing Then
        Throw New HttpException(404, "Image not found")
      End If

      Try
        Response.ContentType = "image/gif"
        Response.OutputStream.Write(image, 0, image.Length)
      Catch ex As Exception

      End Try

      objRs.close()
      objRs = Nothing

    End Sub


    Function LogOff()
      Session("databaseConnection") = Nothing
      Return RedirectToAction("Login", "Account")
    End Function

    Function PasswordChange() As ActionResult
      Return View()
    End Function

    Function NewUser() As ActionResult
      Return View()
    End Function

    'Function ForcePasswordChange() As ActionResult
    '    Return View()
    'End Function

    Function Poll() As PartialViewResult
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

    Function util_run_crosstabsMain() As ActionResult
      Return PartialView()
    End Function

    Function util_run_crosstabsData() As ActionResult
      Return PartialView()
    End Function

    Function util_run_crosstabsBreakdown() As ActionResult
      Return PartialView()
    End Function

    Function util_run_crosstabs() As ActionResult
      Return PartialView()
    End Function

    <HttpPost()>
    Function util_run_crosstabsDataSubmit()

      On Error Resume Next

      Session("CT_Mode") = Request.Form("txtMode")
      Session("CT_EmailGroupID") = Request.Form("txtEmailGroupID")
      Session("CT_EmailGroupAddr") = Request.Form("txtEmailGroupAddr")
      Session("CT_UtilID") = Request.Form("txtUtilID")

      If Session("CT_Mode") = "BREAKDOWN" Then
        Session("CT_Hor") = Request.Form("txtHor")
        Session("CT_Ver") = Request.Form("txtVer")
        Session("CT_Pgb") = Request.Form("txtPgb")
        Session("CT_IntersectionType") = Request.Form("txtIntersectionType")
        Session("CT_CellValue") = Request.Form("txtCellValue")
        Session("CT_Use1000") = Request.Form("txtUse1000")
      Else
        Session("CT_PageNumber") = Request.Form("txtPageNumber")
        Session("CT_IntersectionType") = Request.Form("txtIntersectionType")
        Session("CT_ShowPercentage") = Request.Form("txtShowPercentage")
        Session("CT_PercentageOfPage") = Request.Form("txtPercentageOfPage")
        Session("CT_SuppressZeros") = Request.Form("txtSuppressZeros")
        Session("CT_Use1000") = Request.Form("txtUse1000")
      End If

      ' Go to the requested page.
      Return RedirectToAction("util_run_crosstabsData")

    End Function

    <ValidateInput(False)>
    Function util_run_promptedvalues() As ActionResult
      Return View()
    End Function

    <ValidateInput(False)>
    Function util_run() As ActionResult
      Return PartialView()
    End Function

    <ValidateInput(False)>
    Function util_run_customreports() As ActionResult
      Return PartialView()
    End Function

    Function util_run_calendarreport_main() As ActionResult
      Return PartialView()
    End Function

    Public Function util_run_calendarreport_data() As ActionResult
      Return PartialView()
    End Function

    Function util_run_calendarreport_data_submit() As ActionResult

      On Error Resume Next

      Session("CALREP_Action") = Request.Form("txtAction")
      Session("CALREP_DaysInMonth") = Request.Form("txtDaysInMonth")
      Session("CALREP_Month") = Request.Form("txtMonth")
      Session("CALREP_Year") = Request.Form("txtYear")
      Session("CALREP_VisibleStartDate") = Request.Form("txtVisibleStartDate")
      Session("CALREP_VisibleEndDate") = Request.Form("txtVisibleEndDate")
      Session("CalRep_Mode") = Request.Form("txtMode")
      Session("EmailGroupID") = Request.Form("txtEmailGroupID")
      '  Session("firstLoad") = Request.Form("firstLoad")

      ' Go to the requested page.
      Return RedirectToAction("util_run_calendarreport_data")

    End Function

    <ValidateInput(False)>
    Function util_run_workflow() As ActionResult
      Return PartialView()
    End Function

    <ValidateInput(False)>
    Function WorkflowPendingSteps() As ActionResult
      Return PartialView()
    End Function

    <ValidateInput(False)>
    Function util_run_customreportsData() As ActionResult
      Return PartialView()
    End Function

    <ValidateInput(False)>
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

    Function util_def_calendarreportdates_data_submit()
      Session("CalendarAction") = Request.Form("txtCalendarAction")
      Session("CalendarBaseTableID") = Request.Form("txtCalendarBaseTableID")
      Session("CalendarEventTableID") = Request.Form("txtCalendarEventTableID")
      Session("CalendarLookupTableID") = Request.Form("txtCalendarLookupTableID")

      'Response.Redirect("util_def_calendarreportdates_data")
      Return RedirectToAction("util_def_calendarreportdates_data")
    End Function

    Function util_def_calendarreport_submit()
      On Error Resume Next

      Dim cmdSave = CreateObject("ADODB.Command")
      cmdSave.CommandText = "spASRIntSaveCalendarReport"
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

      Dim prmPicklist = cmdSave.CreateParameter("picklist", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmPicklist)
      prmPicklist.value = CleanNumeric(Request.Form("txtSend_picklist"))

      Dim prmFilter = cmdSave.CreateParameter("filt)er", 3, 1)  ' 3=integer,1=input
      cmdSave.Parameters.Append(prmFilter)
      prmFilter.value = CleanNumeric(Request.Form("txtSend_filter"))

      Dim prmPrintFilterHeader = cmdSave.CreateParameter("printFilterHeader", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmPrintFilterHeader)
      prmPrintFilterHeader.value = CleanBoolean(Request.Form("txtSend_printFilterHeader"))

      Dim prmUserName = cmdSave.CreateParameter("userName", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmUserName)
      prmUserName.value = Request.Form("txtSend_userName")

      Dim prmDescription1 = cmdSave.CreateParameter("description1", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmDescription1)
      prmDescription1.value = CleanNumeric(Request.Form("txtSend_desc1"))

      Dim prmDescription2 = cmdSave.CreateParameter("description2", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmDescription2)
      prmDescription2.value = CleanNumeric(Request.Form("txtSend_desc2"))

      Dim prmDescriptionExpr = cmdSave.CreateParameter("descriptionExpr", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmDescriptionExpr)
      prmDescriptionExpr.value = CleanNumeric(Request.Form("txtSend_descExpr"))

      Dim prmRegion = cmdSave.CreateParameter("region", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmRegion)
      prmRegion.value = CleanNumeric(Request.Form("txtSend_region"))

      Dim prmGroupByDesc = cmdSave.CreateParameter("groupByDesc", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmGroupByDesc)
      prmGroupByDesc.value = CleanBoolean(Request.Form("txtSend_groupbydesc"))

      Dim prmDescSeparator = cmdSave.CreateParameter("descSeparator", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmDescSeparator)
      prmDescSeparator.value = Request.Form("txtSend_descseparator")

      Dim prmStartType = cmdSave.CreateParameter("startType", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmStartType)
      prmStartType.value = CleanNumeric(Request.Form("txtSend_StartType"))

      Dim prmFixedStart = cmdSave.CreateParameter("fixedStart", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmFixedStart)
      If Len(Request.Form("txtSend_FixedStart")) > 0 Then
        prmFixedStart.value = Request.Form("txtSend_FixedStart")
      Else
        prmFixedStart.value = ""
      End If

      Dim prmStartFrequency = cmdSave.CreateParameter("startFrequency", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmStartFrequency)
      prmStartFrequency.value = CleanNumeric(Request.Form("txtSend_StartFrequency"))

      Dim prmStartPeriod = cmdSave.CreateParameter("startPeriod", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmStartPeriod)
      prmStartPeriod.value = CleanNumeric(Request.Form("txtSend_StartPeriod"))

      Dim prmStartDateExpr = cmdSave.CreateParameter("startDateExpr", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmStartDateExpr)
      prmStartDateExpr.value = CleanNumeric(Request.Form("txtSend_CustomStart"))

      Dim prmEndType = cmdSave.CreateParameter("endType", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmEndType)
      prmEndType.value = CleanNumeric(Request.Form("txtSend_EndType"))

      Dim prmFixedEnd = cmdSave.CreateParameter("fixedEnd", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmFixedEnd)
      If Len(Request.Form("txtSend_FixedEnd")) > 0 Then
        prmFixedEnd.value = Request.Form("txtSend_FixedEnd")
      Else
        prmFixedEnd.value = ""
      End If

      Dim prmEndFrequency = cmdSave.CreateParameter("endFrequency", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmEndFrequency)
      prmEndFrequency.value = CleanNumeric(Request.Form("txtSend_EndFrequency"))

      Dim prmEndPeriod = cmdSave.CreateParameter("endPeriod", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmEndPeriod)
      prmEndPeriod.value = CleanNumeric(Request.Form("txtSend_EndPeriod"))

      Dim prmEndDateExpr = cmdSave.CreateParameter("endDateExpr", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmEndDateExpr)
      prmEndDateExpr.value = CleanNumeric(Request.Form("txtSend_CustomEnd"))

      Dim prmShowBankHols = cmdSave.CreateParameter("showBankHols", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmShowBankHols)
      prmShowBankHols.value = CleanBoolean(Request.Form("txtSend_ShadeBHols"))

      Dim prmShowCaptions = cmdSave.CreateParameter("showCaptions", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmShowCaptions)
      prmShowCaptions.value = CleanBoolean(Request.Form("txtSend_Captions"))

      Dim prmShowWeekends = cmdSave.CreateParameter("showWeekends", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmShowWeekends)
      prmShowWeekends.value = CleanBoolean(Request.Form("txtSend_ShadeWeekends"))

      Dim prmStartOnCurrentMonth = cmdSave.CreateParameter("startOnCurrentMonth", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmStartOnCurrentMonth)
      prmStartOnCurrentMonth.value = CleanBoolean(Request.Form("txtSend_StartOnCurrentMonth"))

      Dim prmIncludeWorkdays = cmdSave.CreateParameter("includeWorkdays", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmIncludeWorkdays)
      prmIncludeWorkdays.value = CleanBoolean(Request.Form("txtSend_IncludeWorkingDaysOnly"))

      Dim prmIncludeBankHols = cmdSave.CreateParameter("includeBankHols", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmIncludeBankHols)
      prmIncludeBankHols.value = CleanBoolean(Request.Form("txtSend_IncludeBHols"))

      Dim prmOutputPreview = cmdSave.CreateParameter("outputPreview", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmOutputPreview)
      prmOutputPreview.value = CleanBoolean(Request.Form("txtSend_OutputPreview"))

      Dim prmOutputFormat = cmdSave.CreateParameter("outputFormat", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmOutputFormat)
      prmOutputFormat.value = CleanNumeric(Request.Form("txtSend_OutputFormat"))

      Dim prmOutputScreen = cmdSave.CreateParameter("outputScreen", 11, 1) ' 11=boolean, 1=input
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

      Dim prmAccess = cmdSave.CreateParameter("access", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmAccess)
      prmAccess.value = Request.Form("txtSend_access")

      Dim prmJobToHide = cmdSave.CreateParameter("jobsToHide", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmJobToHide)
      prmJobToHide.value = Request.Form("txtSend_jobsToHide")

      Dim prmJobToHideGroups = cmdSave.CreateParameter("acess", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmJobToHideGroups)
      prmJobToHideGroups.value = Request.Form("txtSend_jobsToHideGroups")

      Dim prmEvents = cmdSave.CreateParameter("events", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmEvents)
      prmEvents.value = Request.Form("txtSend_Events")
      Dim prmEvents2 = cmdSave.CreateParameter("events2", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmEvents2)
      prmEvents2.value = Request.Form("txtSend_Events2")

      'pass the order string to the stored procedure, the stored procedure 
      'saves the order information to the ASRSysCalendarReportOrder table.
      Dim prmOrderString = cmdSave.CreateParameter("orderstring", 200, 1, 8000)  ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmOrderString)
      prmOrderString.value = Request.Form("txtSend_OrderString")

      Dim prmID = cmdSave.CreateParameter("id", 3, 3) ' 3=integer,3=input/output
      cmdSave.Parameters.Append(prmID)
      prmID.value = CleanNumeric(Request.Form("txtSend_ID"))

      cmdSave.Execute()

      If Err.Number = 0 Then
        Session("confirmtext") = "Report has been saved successfully"
        Session("confirmtitle") = "Calendar Reports"
        Session("followpage") = "defsel"
        Session("reaction") = Request.Form("txtSend_reaction")
        Session("utilid") = cmdSave.Parameters("id").Value

        'Response.Redirect("confirmok.asp")
        Return RedirectToAction("ConfirmOK")
      Else
        Response.Write("<HTML>" & vbCrLf)
        Response.Write("	<HEAD>" & vbCrLf)
        Response.Write("		<META NAME=""GENERATOR"" Content=""Microsoft Visual Studio 6.0"">" & vbCrLf)
        Response.Write("		<LINK href=""AutoBG.css"" rel=stylesheet type=text/css >" & vbCrLf)
        Response.Write("		<TITLE>" & vbCrLf)
        Response.Write("			OpenHR Intranet" & vbCrLf)
        Response.Write("		</TITLE>" & vbCrLf)
        Response.Write("		<meta http-equiv=""X-UA-Compatible"" content=""IE=5"">" & vbCrLf)
        Response.Write("	</HEAD>" & vbCrLf)
        Response.Write("	<BODY id=bdyMainBody name=""bdyMainBody"" " & Session("BodyTag") & ">" & vbCrLf)

        Response.Write("	<table align=center border=1 cellPadding=5 cellSpacing=0>" & vbCrLf)
        Response.Write("		<TR>" & vbCrLf)
        Response.Write("			<TD bgcolor=threedface>" & vbCrLf)
        Response.Write("				<table border=0 cellspacing=0 cellpadding=0>" & vbCrLf)
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
        Response.Write("						<INPUT TYPE=button VALUE=""Retry"" NAME=""GoBack"" OnClick=""window.history.back(1)"" style=""WIDTH: 80px"" width=80 id=cmdGoBack>" & vbCrLf)
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

    End Function

    Public Function util_def_calendarreportdates() As ActionResult
      Return View()
    End Function

    Public Function util_def_calendarreportdates_main() As ActionResult
      Return View()
    End Function

    Public Function util_def_calendarreport() As ActionResult
      Return View()
    End Function

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

    Public Function util_def_calendarreportdates_data() As ActionResult
      Return View()
    End Function

    Function util_validate_customreports() As ActionResult
      Return View()
    End Function

    Function util_validate_calendarreport() As ActionResult
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

    <HttpPost(), ValidateInput(False)>
    Function util_def_expression_Submit()

      Dim objExpression As HR.Intranet.Server.Expression
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
      objExpression = New HR.Intranet.Server.Expression

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
    Function quickfind_Submit(value As FormCollection)
      Dim sErrorMsg = ""

      ' Only process the form submission if the referring page was the default page.
      ' If it wasn't then redirect to the login page.

      Dim sFilterSQL = Request.Form("txtGotoOptionFilterSQL")
      Dim sFilterDef = Request.Form("txtGotoOptionFilterDef")
      Dim sValue = Request.Form("txtGotoOptionValue")
      Dim sNextPage = Request.Form("txtGotoOptionPage")
      Dim sAction = Request.Form("txtGotoOptionAction")

      Dim lngRecordID = 0

      Session("optionScreenID") = Request.Form("txtGotoOptionScreenID")
      Session("optionTableID") = Request.Form("txtGotoOptionTableID")
      Session("optionViewID") = Request.Form("txtGotoOptionViewID")
      Session("optionOrderID") = Request.Form("txtGotoOptionOrderID")
      Session("optionRecordID") = Request.Form("txtGotoOptionRecordID")
      Session("optionFilterDef") = Request.Form("txtGotoOptionFilterDef")
      Session("optionFilterSQL") = sFilterSQL
      Session("optionFilterDef") = sFilterDef
      Session("optionValue") = sValue
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
      Session("optionFirstRecPos") = Request.Form("txtGotoOptionFirstRecPos")
      Session("optionCurrentRecCount") = Request.Form("txtGotoOptionCurrentRecCount")
      Session("optionPageAction") = Request.Form("txtGotoOptionPageAction")
      Session("optionCourseTitle") = Request.Form("txtGotoOptionCourseTitle")
      Session("optionExprType") = Request.Form("txtGotoOptionExprType")
      Session("optionExprID") = Request.Form("txtGotoOptionExprID")
      Session("optionFunctionID") = Request.Form("txtGotoOptionFunctionID")
      Session("optionParameterIndex") = Request.Form("txtGotoOptionParameterIndex")

      If sAction = "" Then
        ' Go to the requested page.
        Response.Redirect(sNextPage)
      End If

      If sAction = "CANCEL" Then
        ' Go to the requested page.
        Session("errorMessage") = sErrorMsg
        Response.Redirect(sNextPage)
      End If

      If sAction = "QUICKFIND" Then
        ' Try to get the record that matches the quick find criteria.
        Dim cmdQuickFind = CreateObject("ADODB.Command")
        cmdQuickFind.CommandText = "spASRIntGetQuickFindRecord"
        cmdQuickFind.CommandType = 4 ' Stored Procedure
        cmdQuickFind.ActiveConnection = Session("databaseConnection")

        Dim prmTableID = cmdQuickFind.CreateParameter("tableID", 3, 1)
        cmdQuickFind.Parameters.Append(prmTableID)
        prmTableID.value = CleanNumeric(Session("optionTableID"))

        Dim prmViewID = cmdQuickFind.CreateParameter("viewID", 3, 1)
        cmdQuickFind.Parameters.Append(prmViewID)
        prmViewID.value = CleanNumeric(Session("optionViewID"))

        Dim prmColumnID = cmdQuickFind.CreateParameter("columnID", 3, 1)
        cmdQuickFind.Parameters.Append(prmColumnID)
        prmColumnID.value = CleanNumeric(Session("optionColumnID"))

        Dim prmValue = cmdQuickFind.CreateParameter("value", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
        cmdQuickFind.Parameters.Append(prmValue)
        prmValue.value = sValue

        Dim prmFilterDef = cmdQuickFind.CreateParameter("filterDef", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
        cmdQuickFind.Parameters.Append(prmFilterDef)
        prmFilterDef.value = sFilterDef

        Dim prmResult = cmdQuickFind.CreateParameter("result", 3, 2)
        cmdQuickFind.Parameters.Append(prmResult)

        Dim prmDecSeparator = cmdQuickFind.CreateParameter("decSeparator", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
        cmdQuickFind.Parameters.Append(prmDecSeparator)
        prmDecSeparator.value = Session("LocaleDecimalSeparator")

        Dim prmDateFormat = cmdQuickFind.CreateParameter("dateFormat", 200, 1, 8000) ' 200=varchar, 1=input, 8000=size
        cmdQuickFind.Parameters.Append(prmDateFormat)
        prmDateFormat.value = Session("LocaleDateFormat")

        Err.Clear()
        cmdQuickFind.Execute()

        If Err.Number <> 0 Then
          sErrorMsg = "Error trying to run 'quick find'." & vbCrLf & FormatError(Err.Description)
        Else
          If (cmdQuickFind.Parameters("result").Value = 0) Then
            sErrorMsg = "No records match the criteria."

            If Len(sFilterDef) > 0 Then
              sErrorMsg = sErrorMsg & vbCrLf & _
                "Try removing the filter."
            End If
          Else
            ' A record has been found !
            lngRecordID = cmdQuickFind.Parameters("result").Value
          End If
        End If

        cmdQuickFind = Nothing

        Session("errorMessage") = sErrorMsg

        If Len(sErrorMsg) > 0 Then
          ' Go to the requested page.
          Return RedirectToAction("Quickfind")
        End If

      End If

      ' Go to the requested page.
      Session("optionRecordID") = lngRecordID
      Return RedirectToAction(sNextPage)

    End Function


    Function emptyoption() As ActionResult
      Return View()
    End Function


    <HttpPost()>
    Function util_def_exprcomponent_submit(value As FormCollection)

      Dim sErrorMsg As String = ""
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

    <ValidateInput(False)>
    Function util_test_expression() As ActionResult
      Return View()
    End Function

    <ValidateInput(False)>
    Function util_test_expression_pval() As ActionResult
      Return View()
    End Function

    <ValidateInput(False)>
    Function util_test_expression_submit(value As FormCollection)
      Return RedirectToAction("util_def_expression")
    End Function

    <ValidateInput(False)>
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

    <HttpPost()>
    Function recordEditMain(psScreenInfo As String) As ActionResult

      Session("action") = ""
      Session("parentTableID") = 0
      Session("parentRecordID") = 0
      Session("selectSQL") = ""
      Session("errorMessage") = ""
      Session("warningFlag") = ""
      Session("previousAction") = ""

      Dim sParameters As String = psScreenInfo

      Session("linkType") = Left(sParameters, InStr(sParameters, "_") - 1)

      sParameters = Mid(sParameters, InStr(sParameters, "_") + 1)

      Session("TopLevelRecID") = Left(sParameters, InStr(sParameters, "_") - 1)

      If Session("linkType") = "multifind" Then
        Session("screenID") = 0
        Session("title") = ""
        Session("startMode") = 0
        Session("tableID") = Mid(sParameters, InStr(sParameters, "_") + 1, ((InStr(sParameters, "!") - 1) - InStr(sParameters, "_")))
        Session("viewID") = Mid(sParameters, InStr(sParameters, "!") + 1)
        Session("tableType") = 1

      Else
        Session("linkID") = Mid(sParameters, InStr(sParameters, "_") + 1)

        Err.Clear()
        Dim cmdLinkInfo = CreateObject("ADODB.Command")
        cmdLinkInfo.CommandText = "spASRIntGetLinkInfo"
        cmdLinkInfo.CommandType = 4 ' Stored Procedure
        cmdLinkInfo.ActiveConnection = Session("databaseConnection")

        Dim prmLinkID = cmdLinkInfo.CreateParameter("linkID", 3, 1) ' 3=integer, 1=input
        cmdLinkInfo.Parameters.Append(prmLinkID)
        prmLinkID.value = CLng(CleanNumeric(Session("linkID")))

        Dim prmScreenID = cmdLinkInfo.CreateParameter("screenID", 3, 2) ' 3=integer, 2=output
        cmdLinkInfo.Parameters.Append(prmScreenID)

        Dim prmTableID = cmdLinkInfo.CreateParameter("tableID", 3, 2) ' 3=integer, 2=output
        cmdLinkInfo.Parameters.Append(prmTableID)

        Dim prmTitle = cmdLinkInfo.CreateParameter("title", 200, 2, 8000) ' 200=adVarChar, 2=output, 8000=size
        cmdLinkInfo.Parameters.Append(prmTitle)

        Dim prmStartMode = cmdLinkInfo.CreateParameter("startMode", 3, 2) ' 3=integer, 2=output
        cmdLinkInfo.Parameters.Append(prmStartMode)

        Dim prmTableType = cmdLinkInfo.CreateParameter("tableType", 3, 2) ' 3=integer, 2=output
        cmdLinkInfo.Parameters.Append(prmTableType)

        Err.Clear()
        cmdLinkInfo.Execute()

        Dim sErrorDescription As String = ""

        If (Err.Number <> 0) Then
          sErrorDescription = "Unable to get the link definition." & vbCrLf & FormatError(Err.Description)
        Else
          Session("screenID") = cmdLinkInfo.Parameters("screenID").Value
          Session("tableID") = cmdLinkInfo.Parameters("tableID").Value
          Session("title") = (cmdLinkInfo.Parameters("title").Value)
          Session("startMode") = cmdLinkInfo.Parameters("startMode").Value
          Session("tableType") = cmdLinkInfo.Parameters("tableType").Value
        End If

        cmdLinkInfo = Nothing

        'session("tableID") = session("SSILinkTableID") 
        Session("viewID") = Session("SSILinkViewID")
      End If

      ' recordEditMain.asp now replaced with the following server side code instead. So don't go looking for the form.
      If Session("linkType") = "multifind" Then
        Return RedirectToAction("Find", New With {.sParameters = "LOAD_0_0_"})
      Else
        If (CLng(Session("SSILinkTableID")) = CLng(Session("SingleRecordTableID"))) _
            And (CLng(Session("SSILinkViewID")) = CLng(Session("SingleRecordViewID"))) _
            And (CLng(Session("TopLevelRecID")) = 0) _
            And (CLng(Session("tableID")) <> CLng(Session("SingleRecordTableID"))) Then
          'TODO: error - no parent record in the current view.
          Stop
        End If
        If CleanNumeric(Session("startMode")) <> 3 Then
          Return View("recordEdit")
        Else
          Return RedirectToAction("Find", New With {.sParameters = "LOAD_0_0_"})
        End If
      End If

    End Function


    Function FormError() As JsonResult
      ' replaces response.redirect("error") 
      If NullSafeString(Session("ErrorTitle")).Length = 0 Then Session("ErrorTitle") = "Unspecified Form"
      If NullSafeString(Session("ErrorText")).Length = 0 Then Session("ErrorText") = "Unspecified Error (" & Session("ErrorTitle") & ")"

      Dim errorResponse = New ErrMsgJsonAjaxResponse() With {.ErrorTitle = Session("ErrorTitle"), .ErrorMessage = Session("ErrorText"), .Redirect = ""}
      Return Json(errorResponse, JsonRequestBehavior.AllowGet)

    End Function


#Region "Picklists"

    Function util_def_picklist() As ActionResult
      Return PartialView()
    End Function

    <HttpPost()>
    Function util_def_picklist_submit()

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
      Return View()
    End Function

    Function picklistSelectionData() As ActionResult
      Return View()
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

    Function util_validate_picklist() As ActionResult
      Return View()
    End Function

#End Region

#Region "Utilities"
    Function util_def_mailmerge() As ActionResult
      'Throw New NotImplementedException()
      Return View()
    End Function

    <HttpPost()>
    Function util_def_mailmerge_submit()
      On Error Resume Next

      Dim cmdSave = CreateObject("ADODB.Command")
      cmdSave.CommandText = "sp_ASRIntSaveMailMerge"
      cmdSave.CommandType = 4 ' Stored Procedure
      cmdSave.ActiveConnection = Session("databaseConnection")

      Dim prmName = cmdSave.CreateParameter("name", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmName)
      prmName.value = Request.Form("txtSend_name")

      Dim prmDescription = cmdSave.CreateParameter("description", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmDescription)
      prmDescription.value = Request.Form("txtSend_description")

      Dim prmTableID = cmdSave.CreateParameter("tableID", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmTableID)
      prmTableID.value = CleanNumeric(Request.Form("txtSend_baseTable"))

      Dim prmSelection = cmdSave.CreateParameter("selection", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmSelection)
      prmSelection.value = CleanNumeric(Request.Form("txtSend_selection"))

      Dim prmPicklistID = cmdSave.CreateParameter("picklistID", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmPicklistID)
      prmPicklistID.value = CleanNumeric(Request.Form("txtSend_picklist"))

      Dim prmFilterID = cmdSave.CreateParameter("filterID", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmFilterID)
      prmFilterID.value = CleanNumeric(Request.Form("txtSend_filter"))

      Dim prmOutputFormat = cmdSave.CreateParameter("outputFormat", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmOutputFormat)
      prmOutputFormat.value = CleanNumeric(Request.Form("txtSend_outputformat"))

      Dim prmOutputSave = cmdSave.CreateParameter("outputSave", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmOutputSave)
      prmOutputSave.value = CleanBoolean(Request.Form("txtSend_outputsave"))

      Dim prmOutputFileName = cmdSave.CreateParameter("outputFileName", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmOutputFileName)
      prmOutputFileName.value = Request.Form("txtSend_outputfilename")

      Dim prmEmailAddrID = cmdSave.CreateParameter("emailAddrID", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmEmailAddrID)
      prmEmailAddrID.value = CleanNumeric(Request.Form("txtSend_emailaddrid"))

      Dim prmEmailSubject = cmdSave.CreateParameter("emailSubject", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmEmailSubject)
      prmEmailSubject.value = Request.Form("txtSend_emailsubject")

      Dim prmTemplateFileName = cmdSave.CreateParameter("templateFileName", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmTemplateFileName)
      prmTemplateFileName.value = Request.Form("txtSend_templatefilename")

      Dim prmOutputScreen = cmdSave.CreateParameter("outputScreen", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmOutputScreen)
      prmOutputScreen.value = CleanBoolean(Request.Form("txtSend_outputscreen"))

      Dim prmUserName = cmdSave.CreateParameter("userName", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmUserName)
      prmUserName.value = Request.Form("txtSend_userName")

      Dim prmEmailAsAttachment = cmdSave.CreateParameter("emailAsAttachment", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmEmailAsAttachment)
      prmEmailAsAttachment.value = CleanBoolean(Request.Form("txtSend_emailasattachment"))

      Dim prmEmailAttachmentName = cmdSave.CreateParameter("emailAttachmentName", 200, 1, 8000) ' 200=varchar,1=input,8000=size
      cmdSave.Parameters.Append(prmEmailAttachmentName)
      prmEmailAttachmentName.value = Request.Form("txtSend_emailattachmentname")

      Dim prmSuppressBlanks = cmdSave.CreateParameter("suppressBlanks", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmSuppressBlanks)
      prmSuppressBlanks.value = CleanBoolean(Request.Form("txtSend_suppressblanks"))

      Dim prmPauseBeforeMerge = cmdSave.CreateParameter("pauseBeforeMerge", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmPauseBeforeMerge)
      prmPauseBeforeMerge.value = CleanBoolean(Request.Form("txtSend_pausebeforemerge"))

      Dim prmOutputPrinter = cmdSave.CreateParameter("outputPrinter", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmOutputPrinter)
      prmOutputPrinter.value = CleanBoolean(Request.Form("txtSend_outputprinter"))

      Dim prmOutputPrinterName = cmdSave.CreateParameter("outputPrinterName", 200, 1, 255) ' 200=varchar,1=input,255=size
      cmdSave.Parameters.Append(prmOutputPrinterName)
      prmOutputPrinterName.value = Request.Form("txtSend_outputprintername")

      Dim prmDocumentMapID = cmdSave.CreateParameter("documentMapID", 3, 1) ' 3=integer,1=input
      cmdSave.Parameters.Append(prmDocumentMapID)
      prmDocumentMapID.value = CleanNumeric(Request.Form("txtSend_documentmapid"))

      Dim prmManualDocManHeader = cmdSave.CreateParameter("manualDocManHeader", 11, 1) ' 11=boolean, 1=input
      cmdSave.Parameters.Append(prmManualDocManHeader)
      prmManualDocManHeader.value = CleanBoolean(Request.Form("txtSend_manualdocmanheader"))

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

      Dim prmID = cmdSave.CreateParameter("id", 3, 3) ' 3=integer,3=input/output
      cmdSave.Parameters.Append(prmID)
      prmID.value = CleanNumeric(Request.Form("txtSend_ID"))

      cmdSave.Execute()

      If Err.Number = 0 Then
        Session("confirmtext") = "Mail Merge has been saved successfully"
        Session("confirmtitle") = "Mail Merge"
        Session("followpage") = "defsel"
        Session("reaction") = Request.Form("txtSend_reaction")
        Session("utilid") = cmdSave.Parameters("id").Value

        Response.Redirect("confirmok")
      Else
        Response.Write("<HTML>" & vbCrLf)
        Response.Write("	<HEAD>" & vbCrLf)
        Response.Write("		<META NAME=""GENERATOR"" Content=""Microsoft Visual Studio 6.0"">" & vbCrLf)
        Response.Write("		<LINK href=""OpenHR.css"" rel=stylesheet type=text/css >" & vbCrLf)
        Response.Write("		<TITLE>" & vbCrLf)
        Response.Write("			OpenHR Intranet" & vbCrLf)
        Response.Write("		</TITLE>" & vbCrLf)
        Response.Write("		<meta http-equiv=""X-UA-Compatible"" content=""IE=5"">" & vbCrLf)
        Response.Write("  <!--#INCLUDE FILE=""include/ctl_SetStyles.txt"" -->")
        Response.Write("	</HEAD>" & vbCrLf)
        Response.Write("	<BODY>" & vbCrLf)
        Response.Write("Error saving definition : <BR>" & Err.Description & "<BR>" & vbCrLf)
        Response.Write("<INPUT TYPE=button VALUE=Retry NAME=GoBack OnClick=" & Chr(34) & "window.history.back(1)" & Chr(34) & " class=""btn"" style=" & Chr(34) & "WIDTH: 100px" & Chr(34) & " width=100 id=cmdGoBack>")
        Response.Write("                      onmouseover=""try{button_onMouseOver(this);}catch(e){}""" & vbCrLf)
        Response.Write("                      onmouseout=""try{button_onMouseOut(this);}catch(e){}""" & vbCrLf)
        Response.Write("		                  onfocus=""try{button_onFocus(this);}catch(e){}""" & vbCrLf)
        Response.Write("                      onblur=""try{button_onBlur(this);}catch(e){}"" />" & vbCrLf)
        'Response.Write(vbCrLf & vbCrLf & sSQLString)
        Response.Write("	</BODY>" & vbCrLf)
        Response.Write("<HTML>" & vbCrLf)
      End If

      cmdSave = Nothing
      '%>	

    End Function

    'ND my original call for reference later delete when approp
    '<ValidateInput(False)>
    Function util_validate_mailmerge() As ActionResult
      Return View()
    End Function

#End Region


    Function Quickfind() As ActionResult
      Return View()
    End Function

    Function Filterselect() As ActionResult
      Return View()
    End Function

    <HttpPost()>
    Function filterselect_Submit(value As FormCollection)
      Dim sErrorMsg = ""

      ' Only process the form submission if the referring page was the default page.
      ' If it wasn't then redirect to the login page.
      ' Read the information from the calling form.
      Dim sNextPage = Request.Form("txtGotoOptionPage")
      Dim sAction = Request.Form("txtGotoOptionAction")

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


      If sAction = "CANCEL" Then
        ' Go to the requested page.
        Session("errorMessage") = sErrorMsg
        Return RedirectToAction(sNextPage)
      End If

      If sAction = "SELECTFILTER" Then
        Session("errorMessage") = sErrorMsg

        ' Go to the requested page.
        Return RedirectToAction(sNextPage)
      End If

      Return RedirectToAction(sNextPage)

    End Function

    Function tbAddFromWaitingListFind() As ActionResult
      Return View()
    End Function

    <HttpPost()>
   Function tbAddFromWaitingListFind_Submit(value As FormCollection)

      On Error Resume Next

      Dim sErrorMsg = ""
      Dim iTBResultCode = 0
      Dim sPreReqFails = ""

      ' Only process the form submission if the referring page was the default page.
      ' If it wasn't then redirect to the login page.

      ' Read the information from the calling form.
      Dim sNextPage = Request.Form("txtGotoOptionPage")

      Dim sAction = Request.Form("txtGotoOptionAction")

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

      If (sAction = "SELECTADDFROMWAITINGLIST_1") Then
        If CLng(Session("optionRecordID")) > 0 Then
          ' First pass after selecting the employee to book.
          ' Get the user to choose whether to make the booking 'provisional'
          ' or confirmed.
          If Session("TB_TBStatusPExists") Then
            Return RedirectToAction("tbStatusPrompt")
          Else
            sAction = "SELECTADDFROMWAITINGLIST_2"
            Session("optionAction") = sAction
            Session("optionLookupValue") = "B"
          End If
        End If
      End If

      If (sAction = "SELECTADDFROMWAITINGLIST_2") Then
        If CLng(Session("optionRecordID")) > 0 Then
          If Len(sErrorMsg) = 0 Then
            ' Validate the booking.
            Dim sTBErrorMsg = ""
            Dim sTBWarningMsg = ""
            iTBResultCode = 0

            Dim cmdTBCheck = CreateObject("ADODB.Command")
            cmdTBCheck.CommandText = "sp_ASRIntValidateTrainingBooking"
            cmdTBCheck.CommandType = 4 ' Stored procedure
            cmdTBCheck.ActiveConnection = Session("databaseConnection")

            Dim prmResult = cmdTBCheck.CreateParameter("resultCode", 3, 2) ' 3=integer, 2=output
            cmdTBCheck.Parameters.Append(prmResult)

            Dim prmTBEmployeeRecordID = cmdTBCheck.CreateParameter("empRecID", 3, 1) '3=integer, 1=input
            cmdTBCheck.Parameters.Append(prmTBEmployeeRecordID)
            prmTBEmployeeRecordID.value = CleanNumeric(Session("optionLinkRecordID"))

            Dim prmTBCourseRecordID = cmdTBCheck.CreateParameter("courseRecID", 3, 1) '3=integer, 1=input
            cmdTBCheck.Parameters.Append(prmTBCourseRecordID)
            prmTBCourseRecordID.value = CleanNumeric(Session("optionRecordID"))

            Dim prmTBStatus = cmdTBCheck.CreateParameter("tbStatus", 200, 1, 8000) '200=varchar, 1=input, 8000=size
            cmdTBCheck.Parameters.Append(prmTBStatus)
            prmTBStatus.value = Session("optionLookupValue")

            Dim prmTBRecordID = cmdTBCheck.CreateParameter("tbRecID", 3, 1) '3=integer, 1=input
            cmdTBCheck.Parameters.Append(prmTBRecordID)
            prmTBRecordID.value = 0

            Err.Clear()
            cmdTBCheck.Execute()
            If (Err.Number <> 0) Then
              sErrorMsg = "Error validating training booking." & vbCrLf & FormatError(Err.Description)
            End If

            If Len(sErrorMsg) = 0 Then
              iTBResultCode = cmdTBCheck.Parameters("resultCode").Value
            End If
            cmdTBCheck = Nothing
          End If
        End If
      End If

      ' Go to the requested page.
      Session("TBResultCode") = iTBResultCode
      Session("errorMessage") = sErrorMsg
      Session("PreReqFails") = sPreReqFails ' This will be a sp output in the future along the lines of Bulkbooking
      Return RedirectToAction(sNextPage)

    End Function

    Function tbStatusPrompt() As ActionResult
      Return View()
    End Function

    Function tbBookCourseFind() As ActionResult
      Return View()
    End Function

    <HttpPost()>
    Function tbBookCourseFind_Submit(value As FormCollection)
      On Error Resume Next

      Dim sErrorMsg = ""
      Dim iTBResultCode = 0

      ' Only process the form submission if the referring page was the default page.
      ' If it wasn't then redirect to the login page.
      ' Read the information from the calling form.
      Dim sNextPage = Request.Form("txtGotoOptionPage")
      Dim sAction = Request.Form("txtGotoOptionAction")

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

      If (sAction = "SELECTBOOKCOURSE_1") Then
        If CLng(Session("optionRecordID")) > 0 Then
          ' First pass after selecting the course to book.
          ' Get the user to choose whether to make the booking 'provisional'
          ' or confirmed.
          If Session("TB_TBStatusPExists") Then
            Return RedirectToAction("tbStatusPrompt")
          Else
            sAction = "SELECTBOOKCOURSE_2"
            Session("optionAction") = sAction
            Session("optionLookupValue") = "B"
          End If
        End If
      End If

      If (sAction = "SELECTBOOKCOURSE_2") Then
        If CLng(Session("optionRecordID")) > 0 Then
          ' Get the employee record ID from the given Waiting List record.
          Dim iEmpRecID = 0

          Dim cmdEmpIDFromWLID = CreateObject("ADODB.Command")
          cmdEmpIDFromWLID.CommandText = "sp_ASRIntGetEmpIDFromWLID"
          cmdEmpIDFromWLID.CommandType = 4 ' Stored procedure
          cmdEmpIDFromWLID.ActiveConnection = Session("databaseConnection")

          Dim prmTBEmployeeRecordID = cmdEmpIDFromWLID.CreateParameter("empRecID", 3, 2) '3=integer, 2=output
          cmdEmpIDFromWLID.Parameters.Append(prmTBEmployeeRecordID)

          Dim prmTBWLRecordID = cmdEmpIDFromWLID.CreateParameter("WLRecID", 3, 1) '3=integer, 1=input
          cmdEmpIDFromWLID.Parameters.Append(prmTBWLRecordID)
          prmTBWLRecordID.value = CleanNumeric(CLng(Session("optionRecordID")))

          Err.Clear()
          cmdEmpIDFromWLID.Execute()
          If (Err.Number <> 0) Then
            sErrorMsg = "Error getting employee ID." & vbCrLf & FormatError(Err.Description)
          End If

          If Len(sErrorMsg) = 0 Then
            iEmpRecID = cmdEmpIDFromWLID.Parameters("empRecID").Value

            If iEmpRecID = 0 Then
              sErrorMsg = "Error getting employee ID."
            End If
          End If
          cmdEmpIDFromWLID = Nothing

          If Len(sErrorMsg) = 0 Then
            ' Validate the booking.
            Dim sTBErrorMsg = ""
            Dim sTBWarningMsg = ""
            iTBResultCode = 0

            Dim cmdTBCheck = CreateObject("ADODB.Command")
            cmdTBCheck.CommandText = "sp_ASRIntValidateTrainingBooking"
            cmdTBCheck.CommandType = 4 ' Stored procedure
            cmdTBCheck.ActiveConnection = Session("databaseConnection")

            Dim prmResult = cmdTBCheck.CreateParameter("resultCode", 3, 2) ' 3=integer, 2=output
            cmdTBCheck.Parameters.Append(prmResult)

            prmTBEmployeeRecordID = cmdTBCheck.CreateParameter("empRecID", 3, 1) '3=integer, 1=input
            cmdTBCheck.Parameters.Append(prmTBEmployeeRecordID)
            prmTBEmployeeRecordID.value = CleanNumeric(iEmpRecID)

            Dim prmTBCourseRecordID = cmdTBCheck.CreateParameter("courseRecID", 3, 1) '3=integer, 1=input
            cmdTBCheck.Parameters.Append(prmTBCourseRecordID)
            prmTBCourseRecordID.value = CleanNumeric(Session("optionLinkRecordID"))

            Dim prmTBStatus = cmdTBCheck.CreateParameter("tbStatus", 200, 1, 8000) '200=varchar, 1=input, 8000=size
            cmdTBCheck.Parameters.Append(prmTBStatus)
            prmTBStatus.value = Session("optionLookupValue")

            Dim prmTBRecordID = cmdTBCheck.CreateParameter("tbRecID", 3, 1) '3=integer, 1=input
            cmdTBCheck.Parameters.Append(prmTBRecordID)
            prmTBRecordID.value = 0

            Err.Clear()
            cmdTBCheck.Execute()
            If (Err.Number <> 0) Then
              sErrorMsg = "Error validating training booking." & vbCrLf & FormatError(Err.Description)
            End If

            If Len(sErrorMsg) = 0 Then
              iTBResultCode = cmdTBCheck.Parameters("resultCode").Value
            End If
            cmdTBCheck = Nothing
          End If
        End If
      End If

      ' Go to the requested page.
      Session("TBResultCode") = iTBResultCode
      Session("errorMessage") = sErrorMsg
      Return RedirectToAction(sNextPage)

    End Function

    Function tbBulkBooking() As ActionResult
      Return View()
    End Function

    <HttpPost()>
    Function tbBulkBooking_Submit(value As FormCollection)
      On Error Resume Next

      Dim sErrorMsg = ""
      Dim iTBResultCode = 0
      Dim sPreReqFails = ""
      Dim sUnAvailFails = ""
      Dim sOverlapFails = ""
      Dim sOverBookFails = ""

      ' Read the information from the calling form.
      Dim sNextPage = Request.Form("txtGotoOptionPage")
      Dim sAction = Request.Form("txtGotoOptionAction")

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

      If (sAction = "SELECTBULKBOOKINGS") Then
        If Len(Session("optionLinkRecordID")) > 0 Then
          ' Validate the bulk bookings.
          Dim cmdTBCheck = CreateObject("ADODB.Command")
          cmdTBCheck.CommandText = "sp_ASRIntValidateBulkBookings"
          cmdTBCheck.CommandType = 4 ' Stored procedure
          cmdTBCheck.ActiveConnection = Session("databaseConnection")

          Dim prmTBCourseRecordID = cmdTBCheck.CreateParameter("courseRecID", 3, 1) '3=integer, 1=input
          cmdTBCheck.Parameters.Append(prmTBCourseRecordID)
          prmTBCourseRecordID.value = CleanNumeric(Session("optionRecordID"))

          Dim prmTBEmployeeRecordIDs = cmdTBCheck.CreateParameter("employeeRecIDs", 200, 1, 8000) '200=varchar, 1=input, 8000=size
          cmdTBCheck.Parameters.Append(prmTBEmployeeRecordIDs)
          prmTBEmployeeRecordIDs.value = Session("optionLinkRecordID")

          Dim prmTBStatus = cmdTBCheck.CreateParameter("status", 200, 1, 8000) '200=varchar, 1=input, 8000=size
          cmdTBCheck.Parameters.Append(prmTBStatus)
          prmTBStatus.value = Session("optionLookupValue")

          Dim prmResult = cmdTBCheck.CreateParameter("resultCode", 3, 2) ' 3=integer, 2=output
          cmdTBCheck.Parameters.Append(prmResult)

          Dim prmErrorMsg = cmdTBCheck.CreateParameter("errorMessage", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
          cmdTBCheck.Parameters.Append(prmErrorMsg)

          Dim prmPreRequisites = cmdTBCheck.CreateParameter("PreRequisites", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
          cmdTBCheck.Parameters.Append(prmPreRequisites)

          Dim prmAvailability = cmdTBCheck.CreateParameter("Availability", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
          cmdTBCheck.Parameters.Append(prmAvailability)

          Dim prmOverLapping = cmdTBCheck.CreateParameter("Overlapping", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
          cmdTBCheck.Parameters.Append(prmOverLapping)

          Dim prmOverBooking = cmdTBCheck.CreateParameter("Overbooking", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
          cmdTBCheck.Parameters.Append(prmOverBooking)

          Err.Clear()
          cmdTBCheck.Execute()
          If (Err.Number <> 0) Then
            sErrorMsg = "Error validating training booking transfers." & vbCrLf & FormatError(Err.Description)
          End If

          iTBResultCode = cmdTBCheck.Parameters("resultCode").Value

          sPreReqFails = cmdTBCheck.Parameters("PreRequisites").Value
          sUnAvailFails = cmdTBCheck.Parameters("Availability").Value
          sOverlapFails = cmdTBCheck.Parameters("Overlapping").Value
          sOverBookFails = cmdTBCheck.Parameters("Overbooking").Value

          cmdTBCheck = Nothing
        End If
      End If

      ' Go to the requested page.
      Session("TBResultCode") = iTBResultCode
      Session("errorMessage") = sErrorMsg
      Session("PreReqFails") = sPreReqFails
      Session("UnAvailFails") = sUnAvailFails
      Session("OverlapFails") = sOverlapFails
      Session("OverBookFails") = sOverBookFails

      Return RedirectToAction(sNextPage)

    End Function

    Public Function tbBulkBookingSelectionMain() As ActionResult
      Return View()
    End Function


    <HttpPost()>
    Function tbBulkBookingSelectionData_Submit(value As FormCollection)

      On Error Resume Next

      Response.Expires = -1

      ' Read the information from the calling form.
      '		session("action") = Request.Form("txtAction")
      Session("tableID") = Request.Form("txtTableID")
      Session("viewID") = Request.Form("txtViewID")
      Session("orderID") = Request.Form("txtOrderID")
      '		Session("columnID") = Request.Form("txtColumnID")
      Session("pageAction") = Request.Form("txtPageAction")
      Session("firstRecPos") = Request.Form("txtFirstRecPos")
      Session("currentRecCount") = Request.Form("txtCurrentRecCount")
      Session("locateValue") = Request.Form("txtGotoLocateValue")
      '		session("recordID") = Request.Form("txtRecordID")
      '		session("linkRecordID") = Request.Form("txtLinkRecordID")
      '		session("value") = Request.Form("txtValue")
      '		session("SQL") = Request.Form("txtSQL")
      '		session("promptSQL") = Request.Form("txtPromptSQL")
      Session("fromMenu") = Request.Form("txtGotoFromMenu")

      Session("tbSelectionDataLoading") = False

      ' Go to the requested page.
      Return RedirectToAction("tbBulkBookingSelectionData")

    End Function

    Public Function tbBulkBookingSelectionData() As ActionResult
      Return View()
    End Function


    Function util_run_mailmerge_completed() As ActionResult
      Return View()
    End Function

    Function promptedValues() As ActionResult
      Return View()
    End Function


    <HttpPost()>
    Function promptedValues_Submit(value As FormCollection)
      On Error Resume Next

      Session("filterID") = Request.Form("filterID")
      'Response.Write("<input type=""hidden"" id=filterID name=filterID value=" & Request.Form("filterID") & ">" & vbCrLf)

      Dim sPrompts
      Dim aPrompts(1, 0)
      Dim j = 0
      sPrompts = ""
      ' ReDim Preserve aPrompts(1, 0)
      For i = 1 To (Request.Form.Count)
        Dim sKey = Request.Form.Keys(i)
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
                aPrompts(1, j) = Replace(Request.Form.Item(i), Session("LocaleDecimalSeparator"), ".")
              Case "4"
                ' Date. Reformat to match SQL's mm/dd/yyyy format.
                aPrompts(1, j) = convertLocaleDateToSQL(Request.Form.Item(i))
              Case Else
                aPrompts(1, j) = Request.Form.Item(i)
            End Select
          End If

          sPrompts = sPrompts & aPrompts(0, j) & vbTab & aPrompts(1, j) & vbTab

          j = j + 1
        End If
      Next

      Session("filterIDvalue") = Request.Form("filterID")
      Session("promptsvalue") = sPrompts

      'Response.Write("<input type=""hidden"" id=prompts name=prompts value=""" & sPrompts & """>" & vbCrLf)

      Return RedirectToAction("promptedValues_completed")

    End Function


    Function promptedValues_completed() As ActionResult
      Return View()
    End Function

    Function tbTransferBookingFind() As ActionResult
      Return View()
    End Function



    <HttpPost()>
    Function tbTransferBookingFind_Submit(value As FormCollection)
      On Error Resume Next

      Dim sErrorMsg = ""
      Dim iTBResultCode = 0

      ' Read the information from the calling form.
      Dim sNextPage = Request.Form("txtGotoOptionPage")
      Dim sAction = Request.Form("txtGotoOptionAction")

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

      If (sAction = "SELECTTRANSFERBOOKING_1") Then
        If CLng(Session("optionRecordID")) > 0 Then
          ' Get the employee record ID from the given Training Booking record.
          Dim iEmpRecID = 0

          Dim cmdEmpIDFromTBID = CreateObject("ADODB.Command")
          cmdEmpIDFromTBID.CommandText = "sp_ASRIntGetEmpIDFromTBID"
          cmdEmpIDFromTBID.CommandType = 4 ' Stored procedure
          cmdEmpIDFromTBID.ActiveConnection = Session("databaseConnection")

          Dim prmEmployeeRecordID = cmdEmpIDFromTBID.CreateParameter("empRecID", 3, 2) '3=integer, 2=output
          cmdEmpIDFromTBID.Parameters.Append(prmEmployeeRecordID)

          Dim prmTBRecordID = cmdEmpIDFromTBID.CreateParameter("TBRecID", 3, 1) '3=integer, 1=input
          cmdEmpIDFromTBID.Parameters.Append(prmTBRecordID)
          prmTBRecordID.value = CleanNumeric(CLng(Session("optionRecordID")))

          Err.Clear()
          cmdEmpIDFromTBID.Execute()
          If (Err.Number() <> 0) Then
            sErrorMsg = "Error getting employee ID." & vbCrLf & FormatError(Err.Description)
          End If

          If Len(sErrorMsg) = 0 Then
            iEmpRecID = cmdEmpIDFromTBID.Parameters("empRecID").Value

            If iEmpRecID = 0 Then
              sErrorMsg = "Error getting employee ID."
            End If
          End If
          cmdEmpIDFromTBID = Nothing

          If Len(sErrorMsg) = 0 Then
            ' Validate the booking.
            Dim sTBErrorMsg = ""
            Dim sTBWarningMsg = ""
            iTBResultCode = 0

            Dim cmdTBCheck = CreateObject("ADODB.Command")
            cmdTBCheck.CommandText = "sp_ASRIntValidateTrainingBooking"
            cmdTBCheck.CommandType = 4 ' Stored procedure
            cmdTBCheck.ActiveConnection = Session("databaseConnection")

            Dim prmResult = cmdTBCheck.CreateParameter("resultCode", 3, 2) ' 3=integer, 2=output
            cmdTBCheck.Parameters.Append(prmResult)

            Dim prmTBEmployeeRecordID = cmdTBCheck.CreateParameter("empRecID", 3, 1) '3=integer, 1=input
            cmdTBCheck.Parameters.Append(prmTBEmployeeRecordID)
            prmTBEmployeeRecordID.value = CleanNumeric(iEmpRecID)

            Dim prmTBCourseRecordID = cmdTBCheck.CreateParameter("courseRecID", 3, 1) '3=integer, 1=input
            cmdTBCheck.Parameters.Append(prmTBCourseRecordID)
            prmTBCourseRecordID.value = CleanNumeric(Session("optionLinkRecordID"))

            Dim prmTBStatus = cmdTBCheck.CreateParameter("tbStatus", 200, 1, 8000) '200=varchar, 1=input, 8000=size
            cmdTBCheck.Parameters.Append(prmTBStatus)
            prmTBStatus.value = Session("optionLookupValue")

            prmTBRecordID = cmdTBCheck.CreateParameter("tbRecID", 3, 1) '3=integer, 1=input
            cmdTBCheck.Parameters.Append(prmTBRecordID)
            prmTBRecordID.value = 0

            Err.Clear()
            cmdTBCheck.Execute()
            If (Err.Number() <> 0) Then
              sErrorMsg = "Error validating training booking." & vbCrLf & FormatError(Err.Description)
            End If

            If Len(sErrorMsg) = 0 Then
              iTBResultCode = cmdTBCheck.Parameters("resultCode").Value
            End If
            cmdTBCheck = Nothing
          End If
        End If
      End If

      ' Go to the requested page.
      Session("TBResultCode") = iTBResultCode
      Session("errorMessage") = sErrorMsg
      Return RedirectToAction(sNextPage)

    End Function

    Function util_run_outputoptions() As ActionResult

      Session("CT_Mode") = Request("txtMode")
      Session("OutputOptions_Format") = Request("txtFormat")
      Session("OutputOptions_Screen") = Request("txtScreen")
      Session("OutputOptions_Printer") = Request("txtPrinter")
      Session("OutputOptions_PrinterName") = Request("txtPrinterName")
      Session("OutputOptions_Save") = Request("txtSave")
      Session("OutputOptions_SaveExisting") = Request("txtSaveExisting")
      Session("OutputOptions_Email") = Request("txtEmail")
      Session("OutputOptions_EmailGroupID") = Request("txtEmailGroupID")
      Session("OutputOptions_EmailGroup") = Request("txtEmailGroup")
      Session("OutputOptions_EmailSubject") = Request("txtEmailSubject")
      Session("OutputOptions_EmailAttachAs") = Request("txtEmailAttachAs")
      Session("OutputOptions_Filename") = Request("txtFilename")

      Session("utiltype") = Request.Form("txtUtilType")

      Return View()
    End Function

		Function tbTransferCourseFind() As ActionResult
			Return View()
		End Function

		<HttpPost()>
	 Function tbTransferCourseFind_Submit(value As FormCollection)
			On Error Resume Next

			Dim sErrorMsg = ""
			Dim iTBResultCode = 0

			' Read the information from the calling form.
			Dim sNextPage = Request.Form("txtGotoOptionPage")
			Dim sAction = Request.Form("txtGotoOptionAction")

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
			Session("optionFirstRecPos") = Request.Form("txtGotoOptionFirstRecPos")
			Session("optionCurrentRecCount") = Request.Form("txtGotoOptionCurrentRecCount")
			Session("optionPageAction") = Request.Form("txtGotoOptionPageAction")
			Session("optionCourseTitle") = Request.Form("txtGotoOptionCourseTitle")
			Session("optionExprType") = Request.Form("txtGotoOptionExprType")
			Session("optionExprID") = Request.Form("txtGotoOptionExprID")
			Session("optionFunctionID") = Request.Form("txtGotoOptionFunctionID")
			Session("optionParameterIndex") = Request.Form("txtGotoOptionParameterIndex")
			Session("optionDefSelType") = Request.Form("txtGotoOptionDefSelType")
			Session("optionDefSelRecordID") = Request.Form("txtGotoOptionDefSelRecordID")

			If sAction = "" Then
				' Go to the requested page.
				Return RedirectToAction(sNextPage)
			End If

			If sAction = "SELECTTRANSFERCOURSE" Then

				If Session("optionLinkRecordID") > 0 Then
					' Validate the booking transfers.
					Dim cmdTBCheck = CreateObject("ADODB.Command")
					cmdTBCheck.CommandText = "sp_ASRIntValidateTransfers"
					cmdTBCheck.CommandType = 4 ' Stored procedure
					cmdTBCheck.ActiveConnection = Session("databaseConnection")

					Dim prmTBEmployeeTableID = cmdTBCheck.CreateParameter("empTableID", 3, 1)	'3=integer, 1=input
					cmdTBCheck.Parameters.Append(prmTBEmployeeTableID)
					prmTBEmployeeTableID.value = CleanNumeric(Session("TB_EmpTableID"))

					Dim prmTBCourseTableID = cmdTBCheck.CreateParameter("courseTableID", 3, 1) '3=integer, 1=input
					cmdTBCheck.Parameters.Append(prmTBCourseTableID)
					prmTBCourseTableID.value = CleanNumeric(Session("TB_CourseTableID"))

					Dim prmTBCourseRecordID = cmdTBCheck.CreateParameter("courseRecID", 3, 1)	'3=integer, 1=input
					cmdTBCheck.Parameters.Append(prmTBCourseRecordID)
					prmTBCourseRecordID.value = CleanNumeric(Session("optionRecordID"))

					Dim prmTBNewCourseRecordID = cmdTBCheck.CreateParameter("newCourseRecID", 3, 1)	'3=integer, 1=input
					cmdTBCheck.Parameters.Append(prmTBNewCourseRecordID)
					prmTBNewCourseRecordID.value = CleanNumeric(Session("optionLinkRecordID"))

					Dim prmTBTrainBookTableID = cmdTBCheck.CreateParameter("trainBookTableID", 3, 1) '3=integer, 1=input
					cmdTBCheck.Parameters.Append(prmTBTrainBookTableID)
					prmTBTrainBookTableID.value = CleanNumeric(Session("TB_TBTableID"))

					Dim prmTBTrainBookStatusColumnID = cmdTBCheck.CreateParameter("trainBookStatusColumnID", 3, 1) '3=integer, 1=input
					cmdTBCheck.Parameters.Append(prmTBTrainBookStatusColumnID)
					prmTBTrainBookStatusColumnID.value = CleanNumeric(Session("TB_TBStatusColumnID"))

					Dim prmResult = cmdTBCheck.CreateParameter("resultCode", 3, 2) ' 3=integer, 2=output
					cmdTBCheck.Parameters.Append(prmResult)

					Dim prmErrorMsg = cmdTBCheck.CreateParameter("errorMessage", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
					cmdTBCheck.Parameters.Append(prmErrorMsg)

					Err.Clear()
					cmdTBCheck.Execute()
					If (Err.Number <> 0) Then
						sErrorMsg = "Error validating training booking transfers." & vbCrLf & FormatError(Err.Description)
					End If

					If (Len(sErrorMsg) = 0) And Len(cmdTBCheck.Parameters("errorMessage").Value) > 0 Then
						sErrorMsg = "Error validating training booking transfers." & vbCrLf & cmdTBCheck.Parameters("errorMessage").Value
					End If

					iTBResultCode = cmdTBCheck.Parameters("resultCode").Value

					cmdTBCheck = Nothing
				End If

				Session("TBResultCode") = iTBResultCode
				Session("errorMessage") = sErrorMsg
				Return RedirectToAction(sNextPage)
			End If

		End Function

		Function orderselect() As ActionResult
			Return View()
		End Function

		<HttpPost()>
	 Function orderselect_Submit(value As FormCollection)
			On Error Resume Next

			Dim sErrorMsg = ""

			' Read the information from the calling form.
			Dim lngScreenID = Request.Form("txtGotoOptionScreenID")
			Dim lngViewID = Request.Form("txtGotoOptionViewID")
			Dim lngOrderID = Request.Form("txtGotoOptionOrderID")
			Dim sNextPage = Request.Form("txtGotoOptionPage")
			Dim sAction = Request.Form("txtGotoOptionAction")

			Session("optionScreenID") = lngScreenID
			Session("optionTableID") = Request.Form("txtGotoOptionTableID")
			Session("optionViewID") = lngViewID
			Session("optionOrderID") = lngOrderID
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
			Session("optionAction") = sAction
			Session("orderID") = lngOrderID
			Session("optionFirstRecPos") = Request.Form("txtGotoOptionFirstRecPos")
			Session("optionCurrentRecCount") = Request.Form("txtGotoOptionCurrentRecCount")
			Session("optionExprType") = Request.Form("txtGotoOptionExprType")
			Session("optionExprID") = Request.Form("txtGotoOptionExprID")
			Session("optionFunctionID") = Request.Form("txtGotoOptionFunctionID")
			Session("optionParameterIndex") = Request.Form("txtGotoOptionParameterIndex")


			If sAction = "CANCEL" Then
				' Go to the requested page.
				Session("errorMessage") = sErrorMsg
				Return RedirectToAction(sNextPage)
			End If

			If sAction = "SELECTORDER" Then
				' Get the SQL code for the selected order.
				Dim cmdOrder = CreateObject("ADODB.Command")
				cmdOrder.CommandText = "sp_ASRIntGetOrderSQL"
				cmdOrder.CommandType = 4 ' Stored Procedure
				cmdOrder.ActiveConnection = Session("databaseConnection")

				Dim prmScreenID = cmdOrder.CreateParameter("screenID", 3, 1)
				cmdOrder.Parameters.Append(prmScreenID)
				prmScreenID.value = CleanNumeric(lngScreenID)

				Dim prmViewID = cmdOrder.CreateParameter("viewID", 3, 1)
				cmdOrder.Parameters.Append(prmViewID)
				prmViewID.value = CleanNumeric(lngViewID)

				Dim prmOrderID = cmdOrder.CreateParameter("orderID", 3, 1)
				cmdOrder.Parameters.Append(prmOrderID)
				prmOrderID.value = CleanNumeric(lngOrderID)

				Dim prmFromDef = cmdOrder.CreateParameter("fromDef", 200, 2, 8000) ' 200=varchar, 2=output, 8000=size
				cmdOrder.Parameters.Append(prmFromDef)

				Err.Clear()
				cmdOrder.Execute()

				If (Err.Number <> 0) Then
					sErrorMsg = "Error retrieving the new order definition." & vbCrLf & FormatError(Err.Description)
				Else
					Session("fromDef") = cmdOrder.Parameters("fromDef").Value
				End If

				' Release the ADO command object.
				cmdOrder = Nothing

				Session("errorMessage") = sErrorMsg

				' Go to the requested page.
				Return RedirectToAction(sNextPage)
			End If

			Return RedirectToAction(sNextPage)

		End Function




		Private Function convertLocaleDateToSQL(psDate)
			Dim sLocaleFormat
			Dim sSQLFormat
			Dim iLocaleIndex

			If Len(psDate) > 0 Then
				sLocaleFormat = Session("LocaleDateFormat")

				Dim iIndex = InStr(sLocaleFormat, "mm")
				If iIndex > 0 Then
					sSQLFormat = Mid(psDate, iIndex, 2) & "/"
				End If

				iIndex = InStr(sLocaleFormat, "dd")
				If iIndex > 0 Then
					sSQLFormat = sSQLFormat & Mid(psDate, iIndex, 2) & "/"
				End If

				iIndex = InStr(sLocaleFormat, "yyyy")
				If iIndex > 0 Then
					sSQLFormat = sSQLFormat & Mid(psDate, iIndex, 4)
				End If

				convertLocaleDateToSQL = sSQLFormat
			Else
				convertLocaleDateToSQL = ""
			End If
    End Function


		Function lookupFind() As ActionResult
			Return View()
		End Function

		<HttpPost()>
	 Function lookupFind_Submit(value As FormCollection)

			On Error Resume Next

			Dim sErrorMsg = ""

			' Read the information from the calling form.
			Dim sNextPage = Request.Form("txtGotoOptionPage")
			Dim sAction = Request.Form("txtGotoOptionAction")

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
			Session("optionRecordID") = Request.Form("txtGotoOptionLinkRecordID")
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

			If sAction = "" Then
				' Go to the requested page.
				'Return RedirectToAction(sNextPage)
			End If

			If sAction = "CANCEL" Then
				' Go to the requested page.
				Session("errorMessage") = sErrorMsg
				'Return RedirectToAction(sNextPage)
			End If

			If sAction = "SELECTLOOKUP" Then
				Session("errorMessage") = sErrorMsg

				' Go to the requested page.
				'Return RedirectToAction(sNextPage)
			End If

			' Go to the requested page.
			Return RedirectToAction(sNextPage)

		End Function


#Region "Standard Reports"

		Public Function stdrpt_AbsenceCalendar() As ActionResult
			Return PartialView()
		End Function

		Public Function stdrpt_AbsenceCalendar_details() As ActionResult
			Return View()
		End Function

		<HttpPost()>
		Function stdrpt_AbsenceCalendar_submit(value As FormCollection)

			Session("stdrpt_AbsenceCalendar_StartMonth") = Request.Form("txtStartMonth")
			Session("stdrpt_AbsenceCalendar_StartYear") = Request.Form("txtStartYear")
			Session("stdrpt_AbsenceCalendar_IncludeBankHolidays") = Request.Form("txtIncludeBankHolidays")
			Session("stdrpt_AbsenceCalendar_IncludeWorkingDaysOnly") = Request.Form("txtIncludeWorkingDaysOnly")
			Session("stdrpt_AbsenceCalendar_ShowBankHolidays") = Request.Form("txtShowBankHolidays")
			Session("stdrpt_AbsenceCalendar_ShowCaptions") = Request.Form("txtShowCaptions")
			Session("stdrpt_AbsenceCalendar_ShowWeekends") = Request.Form("txtShowWeekends")
			Return RedirectToAction("stdrpt_AbsenceCalendar")

		End Function

		Public Function stdrpt_def_absence() As ActionResult
			Return View()
		End Function

		<HttpPost()>
		Public Function stdrpt_run_AbsenceBreakdown() As ActionResult
			Return View()
		End Function

#End Region


	End Class

  Public Class ErrMsgJsonAjaxResponse

    Public Property ErrorTitle As String
    Public Property ErrorMessage As String
    Public Property Redirect As String
  End Class







End Namespace




