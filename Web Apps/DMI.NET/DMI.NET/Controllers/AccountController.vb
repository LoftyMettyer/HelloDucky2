Imports ADODB
'Imports VBA

Namespace Controllers
  Public Class AccountController
    Inherits Controller

    Function GetWidgetData(values As FormCollection,
     widgetName As String,
     isWidgetLogin As Boolean,
     widgetUser As String,
     widgetPassword As String,
     widgetDatabase As String,
     widgetServer As String,
     widgetID As String) As JsonResult

      ' get the session("databaseConnection") object
      Dim dbConn As Object

      'If isWidgetLogin Then
      If (Session("databaseConnection")) Is Nothing Then
        Login(values, isWidgetLogin, widgetUser, widgetPassword, widgetDatabase, widgetServer)
      End If

      If Not (Session("databaseConnection")) Is Nothing Then
        dbConn = Session("databaseConnection")

        Select Case widgetName
          Case "DBValue"

            ' get collection of all links
            Dim objButtonInfo As Collection = Session("objButtonInfo")
            Dim sDBValue As String = ""
            Dim sCaption As String = ""
            Dim sFormattingSuffix As String = ""
            Dim sFormattingPrefix As String = ""

            ' will Linq this at some point.
            For Each collectionItem As Object In objButtonInfo
              If collectionItem.ID = widgetID Then

                ' sDBValue = DBValue(dbConn, 1, 163, 16771, 1, 4, 0, 0, 0)
                sDBValue = DBValue(dbConn, collectionItem.Chart_TableID, collectionItem.Chart_ColumnID,
                 collectionItem.Chart_FilterID,
                 collectionItem.Chart_AggregateType, collectionItem.Element_Type,
                 collectionItem.Chart_SortOrderID, collectionItem.Chart_SortDirection,
                 collectionItem.Chart_ColourID)

                sCaption = collectionItem.Text

                sFormattingSuffix = collectionItem.Formatting_Suffix

                sFormattingPrefix = collectionItem.Formatting_Prefix
                Exit For
              End If

            Next
            Dim iDBValue As Integer

            Try
              iDBValue = CInt(sDBValue)
            Catch ex As Exception
              iDBValue = 0
            End Try

            Dim data = New JsonData_DBValue() _
             With {.Caption = sCaption, .DBValue = iDBValue, .Formatting_Suffix = sFormattingSuffix,
             .Formatting_Prefix = sFormattingPrefix}
            Return Json(data, JsonRequestBehavior.AllowGet)

          Case "GetLinks"

            ' return a list of all ID's and element types and short descriptions
            Return Json(GetLinks, JsonRequestBehavior.AllowGet)


        End Select

      End If
      Return Json("Undefined")
    End Function

    Function GetLinks() As List(Of navigationLinks)

      Dim objButtonInfo As Collection = Session("objButtonInfo")

      Return (From collectionItem As Object In objButtonInfo Select New navigationLinks(collectionItem.linkType, collectionItem.linkOrder, collectionItem.prompt, collectionItem.text, collectionItem.element_Type, collectionItem.ID)).ToList()
    End Function

    Function DBValue(dbConn As Object,
     iChartTableID As Long,
     iChartColumnID As Long,
     iChartFilterID As Long,
     iChartAggregateType As Long,
     iChartElementType As Long,
     iChartSortOrderID As Long,
     iChartSortDirection As Long,
     iChartColourID As Long) As String

      Dim objChart = New Global.HR.Intranet.Server.clsChart

      ' reset the globals
      objChart.resetGlobals()


      ' Pass required info to the DLL
      objChart.Username = Session("username")
      objChart.Connection = dbConn
      'Session("databaseConnection")

      Dim mrstDBValueData

      mrstDBValueData = objChart.GetChartData(iChartTableID, iChartColumnID, iChartFilterID, iChartAggregateType,
        iChartElementType, iChartSortOrderID, iChartSortDirection, iChartColourID)

      If (Err.Number <> 0) Then
        Session("ErrorTitle") = "The Database Values could not be retrieved." & vbCrLf & formatError(Err.Description)
      Else
        Session("ErrorTitle") = ""
      End If
      Dim sText As String = ""

      If Len(Session("ErrorTitle")) = 0 Then
        Try
          If Not (mrstDBValueData.EOF And mrstDBValueData.bof) Then
            Dim iRecNum = 1
            Do While Not mrstDBValueData.EOF
              sText = mrstDBValueData.fields(0).value
              mrstDBValueData.movenext()
              iRecNum = iRecNum + 1
            Loop
            mrstDBValueData.close()
            'mrstDBValueData = Nothing
          Else ' no results - return zero
            sText = "No Data"
          End If
        Catch ex As Exception
          sText = "No Data"
        End Try


      End If


      Return sText
    End Function

    Function Login() As ActionResult
      Session("ErrorText") = Nothing
      Return View()
    End Function

    <HttpPost()>
    Function Login(values As FormCollection, Optional isWidgetLogin As Boolean = False,
 Optional widgetUser As String = "",
 Optional widgetPassword As String = "",
 Optional widgetDatabase As String = "",
 Optional widgetServer As String = "") As ActionResult


      'On Error Resume Next

      'Dim sReferringPage
      Dim sUserName
      Dim sPassword
      Dim sDatabaseName
      Dim sServerName
      Dim sLocaleDateFormat
      Dim sLocaleDateSeparator
      Dim sLocaleDecimalSeparator
      Dim sLocaleThousandSeparator
      Dim fForcePasswordChange
      Dim sConnectString
      Dim bWindowsAuthentication

      fForcePasswordChange = False
      Session("ConvertedDesktopColour") = "#f9f7fb"

      ' Only try to login if the referring page was the login page.
      ' If it wasn't then redirect to the login page.
      'sReferringPage = Request.ServerVariables("HTTP_REFERER")
      'If InStrRev(sReferringPage, "/") > 0 Then
      '	sReferringPage = Mid(sReferringPage, InStrRev(sReferringPage, "/") + 1)
      'End If

      If Not isWidgetLogin Then
        ' Read the User Name and Password from the Login form.
        sUserName = Request.Form("txtUserNameCopy")
        sPassword = Request.Form("txtPassword")
        sDatabaseName = Request.Form("txtDatabase")
        sServerName = Request.Form("txtServer")
        bWindowsAuthentication = Request.Form("chkWindowsAuthentication")
        sLocaleDateFormat = Request.Form("txtLocaleDateFormat")
        sLocaleDateSeparator = Request.Form("txtLocaleDateSeparator")

        sLocaleDecimalSeparator = Request.Form("txtLocaleDecimalSeparator")
        sLocaleThousandSeparator = Request.Form("txtLocaleThousandSeparator")

        Session("WordVer") = Request.Form("txtWordVer")
        Session("ExcelVer") = Request.Form("txtExcelVer")
      Else
        ' Read the User Name and Password from the Login form.
        sUserName = widgetUser
        sPassword = widgetPassword
        sDatabaseName = widgetDatabase
        sServerName = ".\sql2012"
        ' widgetServer
        bWindowsAuthentication = ""
        sLocaleDateFormat = "ddmmYYYY"
        ' Request.Form("txtLocaleDateFormat")
        sLocaleDateSeparator = "/"
        'Request.Form("txtLocaleDateSeparator")

        sLocaleDecimalSeparator = "."
        'Request.Form("txtLocaleDecimalSeparator")
        sLocaleThousandSeparator = ","
        'Request.Form("txtLocaleThousandSeparator")

        Session("WordVer") = "12"
        'Request.Form("txtWordVer")
        Session("ExcelVer") = "12"
        ' Request.Form("txtExcelVer")
      End If


      Session("LocaleDateFormat") = sLocaleDateFormat
      Session("LocaleDateSeparator") = sLocaleDateSeparator
      Session("LocaleDecimalSeparator") = sLocaleDecimalSeparator
      Session("LocaleThousandSeparator") = sLocaleThousandSeparator

      ' Store the username, for use in forcedchangepassword.
      Session("Username") = LCase(sUserName)

      ' Check if the server DLL is registered.
      Try
        Dim objMenu = New Menu
      Catch ex As Exception
        If Err.Number <> 0 Then
          Session("ErrorTitle") = "Login Page"
          Session("ErrorText") =
           "You could not login to the OpenHR database because of the following error:<p>COAInt_Server.DLL has not been registered on the IIS server.  Please contact support." & vbCrLf &
           "error: " & Err.Number.ToString & ": " & Err.Description
          Return RedirectToAction("Loginerror")
        End If
      End Try

      Dim objSettings = New Global.HR.Intranet.Server.clsSettings
      sConnectString = objSettings.GetSQLProviderString & "Data Source=" & sServerName & ";Initial Catalog=" &
       sDatabaseName & ";Application Name=OpenHR Intranet;DataTypeCompatibility=80;MARS Connection=True;"
      objSettings = Nothing

      ' Different connection string depending if use are using Windows Authentication
      If Not bWindowsAuthentication = "on" Then
        sConnectString = sConnectString & ";User ID=" & sUserName & ";Password=" & sPassword
      Else
        sConnectString = sConnectString & ";Trusted_Connection=yes;"
        sConnectString = sConnectString & ";Integrated Security=SSPI;"
      End If

      sConnectString = sConnectString & ";Persist Security Info=True;"

      ' Open a connection to the database.
      Dim conX = New Connection
      conX.ConnectionTimeout = 60

      Try
        conX.Open(sConnectString)
      Catch ex As Exception
        If InStr(1, Err.Description, "The password of the account must be changed") Or
         InStr(1, Err.Description, "The password for this login has expired") Or
         InStr(1, Err.Description, "The password of the account has expired") Then
          Session("SQL2005Force") = "Server=" & sServerName & ";UID=" & sUserName
          Return RedirectToAction("ForcedPasswordChange", "Account")
        End If

        If Err.Number <> 0 Then
          Session("ErrorTitle") = "Login Page"
          Session("ErrorText") =
           "The system could not log you on. Make sure your details are correct, then retype your password."
          Return RedirectToAction("Loginerror")
        End If
      End Try





      ' Set a 5 minute command timeout
      'conX.CommandTimeout = 300
      ' Set no command timeout
      conX.CommandTimeout = 0

      ' Enter the current session in the poll table. This will
      ' ensure that even if the login checks fail, the session will still be killed
      ' after 1 minute.
      Dim cmdHit = New Command
      cmdHit.CommandText = "sp_ASRIntPoll"
      cmdHit.CommandType = 4
      ' Stored Procedure
      cmdHit.ActiveConnection = conX
      Err.Number = 0
      cmdHit.Execute()
      If (Err.Number <> 0) Then
        Session("ErrorTitle") = "Login Page"
        Dim sErrorText = "You could not login to the OpenHR database because of the following error:<p>"

        If (Err.Number = -2147217900) _
         And (UCase(Left(formatError(Err.Description), 31)) = "COULD NOT FIND STORED PROCEDURE") Then
          sErrorText = sErrorText &
           "The database has not been scripted to run the intranet.<P>" &
           "Contact your system administrator."
        Else
          sErrorText = sErrorText & formatError(Err.Description)
        End If

        Session("ErrorText") = sErrorText
        Return RedirectToAction("Loginerror")
      End If
      cmdHit = Nothing

      Session("databaseConnection") = conX


      ' Successful login.


      ' Get the desktop colour.
      Dim cmdDesktopColour = New Command
      cmdDesktopColour.CommandText = "sp_ASRIntGetSetting"
      cmdDesktopColour.CommandType = 4
      ' Stored procedure.
      cmdDesktopColour.ActiveConnection = Session("databaseConnection")

      Dim prmSection = cmdDesktopColour.CreateParameter("section", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdDesktopColour.Parameters.Append(prmSection)
      prmSection.Value = "DesktopSetting"

      Dim prmKey = cmdDesktopColour.CreateParameter("key", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdDesktopColour.Parameters.Append(prmKey)
      prmKey.Value = "BackgroundColour"

      Dim prmDefault = cmdDesktopColour.CreateParameter("default", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdDesktopColour.Parameters.Append(prmDefault)
      prmDefault.Value = "2147483660"

      Dim prmUserSetting = cmdDesktopColour.CreateParameter("userSetting", 11, 1)
      ' 11=bit, 1=input
      cmdDesktopColour.Parameters.Append(prmUserSetting)
      prmUserSetting.Value = 0

      Dim prmResult = cmdDesktopColour.CreateParameter("result", 200, 2, 8000)
      ' 200=varchar, 2=output, 8000=size
      cmdDesktopColour.Parameters.Append(prmResult)

      Err.Number = 0
      cmdDesktopColour.Execute()
      Session("DesktopColour") = CLng(cmdDesktopColour.Parameters("result").Value)

      cmdDesktopColour = Nothing

      ' Convert the read Desktop colour into a value that can be used by the BODY tag.
      Select Case CStr(Session("DesktopColour"))
        Case "-2147483648"
          Session("ConvertedDesktopColour") = "scrollbar"

        Case "-2147483647"
          Session("ConvertedDesktopColour") = "background"

        Case "-2147483646"
          Session("ConvertedDesktopColour") = "activecaption"

        Case "-2147483645"
          Session("ConvertedDesktopColour") = "inactivecaption"

        Case "-2147483644"
          Session("ConvertedDesktopColour") = "menu"

        Case "-2147483643"
          Session("ConvertedDesktopColour") = "window"

        Case "-2147483642"
          Session("ConvertedDesktopColour") = "windowframe"

        Case "-2147483641"
          Session("ConvertedDesktopColour") = "menutext"

        Case "-2147483640"
          Session("ConvertedDesktopColour") = "windowtext"

        Case "-2147483639"
          Session("ConvertedDesktopColour") = "captiontext"

        Case "-2147483638"
          Session("ConvertedDesktopColour") = "activeborder"

        Case "-2147483637"
          Session("ConvertedDesktopColour") = "inactiveborder"

        Case "-2147483636"
          Session("ConvertedDesktopColour") = "appworkspace"

        Case "-2147483635"
          Session("ConvertedDesktopColour") = "highlight"

        Case "-2147483634"
          Session("ConvertedDesktopColour") = "highlighttext"

        Case "-2147483633"
          Session("ConvertedDesktopColour") = "threedface"

        Case "-2147483632"
          Session("ConvertedDesktopColour") = "threedshadow"

        Case "-2147483631"
          Session("ConvertedDesktopColour") = "graytext"

        Case "-2147483630"
          Session("ConvertedDesktopColour") = "buttontext"

        Case "-2147483629"
          Session("ConvertedDesktopColour") = "inactivecaptiontext"

        Case "-2147483628"
          Session("ConvertedDesktopColour") = "threedhighlight"

        Case "-2147483627"
          Session("ConvertedDesktopColour") = "threeddarkshadow"

        Case "-2147483626"
          Session("ConvertedDesktopColour") = "threedlightshadow"

        Case "-2147483625"
          Session("ConvertedDesktopColour") = "infotext"

        Case "-2147483624"
          Session("ConvertedDesktopColour") = "infobackground"

        Case Else
          Dim sColour = Hex(CLng(Session("DesktopColour")))

          Do While (Len(sColour) < 6)
            sColour = "0" & sColour
          Loop

          Dim sConvertedColour = "#"
          sConvertedColour = sConvertedColour & Mid(sColour, 5, 2)
          sConvertedColour = sConvertedColour & Mid(sColour, 3, 2)
          sConvertedColour = sConvertedColour & Mid(sColour, 1, 2)

          Session("ConvertedDesktopColour") = sConvertedColour
      End Select

      ' Check that its okay for the user to login.
      Dim cmdLoginCheck = New Command
      cmdLoginCheck.CommandText = "sp_ASRIntCheckLogin"
      cmdLoginCheck.CommandType = 4
      ' Stored Procedure
      cmdLoginCheck.ActiveConnection = conX

      Dim prmSuccessFlag = cmdLoginCheck.CreateParameter("SuccessFlag", 3, 2)
      ' 3 = integer, 2 = output
      cmdLoginCheck.Parameters.Append(prmSuccessFlag)

      Dim prmErrorMessage = cmdLoginCheck.CreateParameter("ErrorMessage", 200, 2, 8000)
      ' 200 = varchar, 2 = output, 8000 = size
      cmdLoginCheck.Parameters.Append(prmErrorMessage)

      Dim prmMinPasswordLength = cmdLoginCheck.CreateParameter("MinPasswordLength", 3, 2)
      ' 3 = integer, 2 = output
      cmdLoginCheck.Parameters.Append(prmMinPasswordLength)

      Dim prmIntranetVersion = cmdLoginCheck.CreateParameter("version", 200, 1, 50)
      ' 200 = varchar, 1 = input, 50 = size
      cmdLoginCheck.Parameters.Append(prmIntranetVersion)
      prmIntranetVersion.Value = Session("version")

      Dim prmPasswordLength = cmdLoginCheck.CreateParameter("pwdLength", 3, 1)
      ' 3 = integer, 1 = input
      cmdLoginCheck.Parameters.Append(prmPasswordLength)
      prmPasswordLength.Value = cleanNumeric(Len(sPassword))

      Dim prmUserType = cmdLoginCheck.CreateParameter("userType", 3, 2)
      ' 3 = integer, 2 = output
      cmdLoginCheck.Parameters.Append(prmUserType)

      Err.Number = 0
      cmdLoginCheck.Execute()

      If (Err.Number <> 0) Then
        Session("ErrorTitle") = "Login Page"

        If Err.Number = -2147217900 Then
          Session("ErrorText") = "Unable to login to the OpenHR database:<p>" &
           formatError(
            "Please ask the System Administrator to update the database in the System Manager.")
        Else
          Session("ErrorText") = "You could not login to the OpenHR database because of the following error:<p>" &
           formatError(Err.Description)
        End If
        Return RedirectToAction("Loginerror")
      End If

      Session("userType") = cmdLoginCheck.Parameters("userType").Value

      If cmdLoginCheck.Parameters("SuccessFlag").Value = 0 Then
        Session("ErrorTitle") = "Login Page"
        Session("ErrorText") = "You could not login to the OpenHR database because of the following error:<p>" &
         cmdLoginCheck.Parameters("ErrorMessage").Value
        Return RedirectToAction("Loginerror")
      ElseIf cmdLoginCheck.Parameters("SuccessFlag").Value = 2 Then
        '	Password expired.
        fForcePasswordChange = True
        Session("minPasswordLength") = cmdLoginCheck.Parameters("MinPasswordLength").Value
      End If

      ' NPG20091102 Fault HRPRO-354
      ' Retrieve the user's group name and assign it to a session variable
      Dim cmdDetail = New Command
      cmdDetail.CommandText = "spASRGetActualUserDetails"
      cmdDetail.CommandType = 4
      ' Stored Procedure      
      cmdDetail.ActiveConnection = conX

      Dim prmUserName = cmdDetail.CreateParameter("UserName", 200, 2, 255)
      ' 200=varchar, 2=output, 255=size
      cmdDetail.Parameters.Append(prmUserName)

      Dim prmUserGroup = cmdDetail.CreateParameter("Group", 200, 2, 255)
      ' 200=varchar, 2=output, 255=size
      cmdDetail.Parameters.Append(prmUserGroup)

      Dim prmGroupID = cmdDetail.CreateParameter("GroupID", 3, 2, 50)
      ' 3=integer, 2=output, 50=size
      cmdDetail.Parameters.Append(prmGroupID)

      Dim prmModuleKey = cmdDetail.CreateParameter("ModuleKey", 200, 1, 20, "INTRANET")
      ' 200=varchar, 1=input, 20=size
      cmdDetail.Parameters.Append(prmModuleKey)

      Err.Clear()
      cmdDetail.Execute()

      If (Err.Number <> 0) Then
        Session("ErrorTitle") = "Login Page"

        If Err.Number = -2147217900 Then
          Session("ErrorText") = "Unable to login to the OpenHR database:<p>" &
           formatError(
            "Please ask the System Administrator to update the database in the System Manager.")
        Else
          Session("ErrorText") = "You could not login to the OpenHR database because of the following error:<p>" &
           formatError(Err.Description)
        End If
        Return RedirectToAction("Loginerror")
      End If

      Session("Usergroup") = cmdDetail.Parameters("Group").Value



      ' Do we have DMI Multi-record access?	
      Dim cmdDmiUser = New Command
      cmdDmiUser.CommandText =
        "SELECT count(ItemID) as blah" & vbNewLine & _
        "FROM ASRSysGroupPermissions" & vbNewLine & _
        "WHERE ASRSysGroupPermissions.itemID = (SELECT ASRSysPermissionItems.itemID" & vbNewLine & _
        "FROM ASRSysPermissionItems" & vbNewLine & _
        "INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & vbNewLine & _
        "WHERE (ASRSysPermissionItems.itemKey = 'INTRANET'" & vbNewLine & _
        "AND ASRSysPermissionCategories.categoryKey = 'MODULEACCESS')" & vbNewLine & _
        "AND ASRSysGroupPermissions.groupName = '" & Session("Usergroup") & "'" & vbNewLine & _
        "AND ASRSysGroupPermissions.permitted = 1)" & vbNewLine
      cmdDmiUser.ActiveConnection = conX

      Err.Clear()
      Dim result = cmdDmiUser.Execute()

      If (Err.Number <> 0) Then
        Session("ErrorTitle") = "Login Page"

        If Err.Number = -2147217900 Then
          Session("ErrorText") = "Unable to login to the OpenHR database:<p>" &
           FormatError(
            "Please ask the System Administrator to update the database in the System Manager.")
        Else
          Session("ErrorText") = "You could not login to the OpenHR database because of the following error:<p>" &
           FormatError(Err.Description)
        End If
        Return RedirectToAction("Loginerror")
      End If

      Dim superUserAccessCount = result.Fields(0).Value

      If CLng(superUserAccessCount) > 0 Then
        ' We are a superUser.
        Session("SuperUser") = True
      Else
        Session("SuperUser") = False
      End If

      cmdDmiUser = Nothing


      ' If the users default database is not 'master' then make it so.
      Dim cmdDefaultDB = New Command
      cmdDefaultDB.CommandText =
       "IF EXISTS(SELECT 1 FROM master..syslogins WHERE loginname = SUSER_NAME() AND dbname <> 'master')" & vbNewLine &
       "	EXEC sp_defaultdb [" & sUserName & "], master"
      cmdDefaultDB.ActiveConnection = conX
      cmdDefaultDB.Execute()
      cmdDefaultDB = Nothing

      ' RH 22/03/01 - Put the username in a session variable	
      Session("Server") = sServerName
      ' Moved the following to above the call for forcedpasswordchange.
      ' Session("Username") = LCase(sUserName)
      Session("Database") = sDatabaseName
      Session("WinAuth") = (bWindowsAuthentication = "on")

      ' Release the ADO command object.
      cmdLoginCheck = Nothing

      ' Release the ADO command object.
      cmdDetail = Nothing


      ' Create the cached system tables on the server - Don;t do it in a stored procedure because the #temp will then only be visible to that stored procedure	
      Dim sString = "DECLARE @iUserGroupID	integer, " & vbNewLine &
        "	@sUserGroupName		sysname, " & vbNewLine &
        "	@sActualLoginName	varchar(250) " & vbNewLine &
        "-- Get the current user's group ID. " & vbNewLine &
        "EXEC spASRIntGetActualUserDetails " & vbNewLine &
        "	@sActualLoginName OUTPUT, " & vbNewLine &
        "	@sUserGroupName OUTPUT, " & vbNewLine &
        "	@iUserGroupID OUTPUT " & vbNewLine &
        "-- Create the SysProtects cache table " & vbNewLine &
        "IF OBJECT_ID('tempdb..#SysProtects') IS NOT NULL " & vbNewLine &
        "	DROP TABLE #SysProtects " & vbNewLine &
        "CREATE TABLE #SysProtects(ID int, Action tinyint, Columns varbinary(8000), ProtectType int) " &
        vbNewLine &
        "	INSERT #SysProtects " & vbNewLine &
        "	SELECT ID, Action, Columns, ProtectType " & vbNewLine &
        "       FROM sysprotects " & vbNewLine &
        "       WHERE uid = @iUserGroupID"

      'cmdCreateCache.CommandType = 4 ' Stored Procedure
      Dim cmdCreateCache = New Command
      cmdCreateCache.ActiveConnection = conX
      cmdCreateCache.CommandText = sString
      cmdCreateCache.Execute()
      cmdCreateCache = Nothing

      ' RH 18/04/01 - Put entry in the audit access log
      Dim cmdAudit = New Command
      cmdAudit.CommandText = "sp_ASRIntAuditAccess"
      cmdAudit.CommandType = 4
      ' Stored Procedure
      cmdAudit.ActiveConnection = conX

      Dim prmLoggingIn = cmdAudit.CreateParameter("LoggingIn", 11, 1, , True)
      ' 11 = boolean, 3 = int, 200 = varchar, 2 = output, 8000 = size
      cmdAudit.Parameters.Append(prmLoggingIn)

      Dim prmUser = cmdAudit.CreateParameter("Username", 200, 1, 1000)
      cmdAudit.Parameters.Append(prmUser)
      prmUser.Value = sUserName

      Err.Number = 0
      cmdAudit.Execute()

      If (Err.Number <> 0) Then
        Session("ErrorTitle") = "Login Page - Audit Access"
        Session("ErrorText") = "You could not login to the OpenHR database because of the following error:<p>" &
         FormatError(Err.Description)
        Return RedirectToAction("Loginerror")
      End If

      cmdAudit = Nothing

      ' Get Personnel module parameters		
      Dim cmdPersonnel = New Command
      cmdPersonnel.CommandText = "sp_ASRIntGetPersonnelParameters"
      cmdPersonnel.CommandType = 4
      ' Stored Procedure
      cmdPersonnel.ActiveConnection = conX

      Dim prmEmpTableID = cmdPersonnel.CreateParameter("empTableID", 3, 2)
      ' 3=integer, 2=output
      cmdPersonnel.Parameters.Append(prmEmpTableID)

      Err.Number = 0
      cmdPersonnel.Execute()

      If (Err.Number <> 0) Then
        Session("ErrorTitle") = "Login Page"
        Session("ErrorText") = "You could not login to the OpenHR database because of the following error:<p>" &
         FormatError(Err.Description)
        Return RedirectToAction("Loginerror")
      End If

      Session("Personnel_EmpTableID") = cmdPersonnel.Parameters("empTableID").Value

      cmdPersonnel = Nothing

      ' Get Workflow module parameters		
      Dim cmdWorkflow = New Command
      cmdWorkflow.CommandText = "spASRIntGetWorkflowParameters"
      cmdWorkflow.CommandType = 4
      ' Stored Procedure
      cmdWorkflow.ActiveConnection = conX

      Dim prmWFEnabled = cmdWorkflow.CreateParameter("wfEnabled", 11, 2)
      ' 11=boolean, 2=output
      cmdWorkflow.Parameters.Append(prmWFEnabled)

      Err.Number = 0
      cmdWorkflow.Execute()

      If (Err.Number <> 0) Then
        Session("ErrorTitle") = "Login Page"
        Session("ErrorText") = "You could not login to the OpenHR database because of the following error:<p>" &
         FormatError(Err.Description)
        Return RedirectToAction("Loginerror")
      End If

      Session("WF_Enabled") = cmdWorkflow.Parameters("wfEnabled").Value

      cmdWorkflow = Nothing

      ' Check if the OutOfOffice parameters have been configured.
      Dim fWorkflowOutOfOfficeConfigured = False
      Dim fWorkflowOutOfOffice = False
      Dim iWorkflowRecordCount = 0

      If Session("WF_Enabled") Then
        cmdWorkflow = New Command
        cmdWorkflow.CommandText = "spASRWorkflowOutOfOfficeConfigured"
        cmdWorkflow.CommandType = 4
        ' Stored Procedure
        cmdWorkflow.ActiveConnection = conX

        Dim prmWFOutOfOfficeConfig = cmdWorkflow.CreateParameter("wfOutOfOfficeConfig", 11, 2)
        ' 11=boolean, 2=output
        cmdWorkflow.Parameters.Append(prmWFOutOfOfficeConfig)

        Err.Number = 0
        cmdWorkflow.Execute()

        fWorkflowOutOfOfficeConfigured = cmdWorkflow.Parameters("wfOutOfOfficeConfig").Value

        cmdWorkflow = Nothing

        If fWorkflowOutOfOfficeConfigured Then
          ' Check if the current user OutOfOffice
          cmdWorkflow = New Command
          cmdWorkflow.CommandText = "spASRWorkflowOutOfOfficeCheck"
          cmdWorkflow.CommandType = 4
          ' Stored Procedure
          cmdWorkflow.ActiveConnection = conX

          Dim prmWFOutOfOffice = cmdWorkflow.CreateParameter("wfOutOfOffice", 11, 2)
          ' 11=boolean, 2=output
          cmdWorkflow.Parameters.Append(prmWFOutOfOffice)

          Dim prmWFRecordCount = cmdWorkflow.CreateParameter("wfRecordCount", 3, 2)
          ' 3=integer, 2=output
          cmdWorkflow.Parameters.Append(prmWFRecordCount)

          Err.Number = 0
          cmdWorkflow.Execute()

          fWorkflowOutOfOffice = cmdWorkflow.Parameters("wfOutOfOffice").Value
          iWorkflowRecordCount = cmdWorkflow.Parameters("wfRecordCount").Value

          cmdWorkflow = Nothing
        End If
      End If
      Session("WF_OutOfOfficeConfigured") = fWorkflowOutOfOfficeConfigured
      Session("WF_OutOfOffice") = fWorkflowOutOfOffice
      Session("WF_RecordCount") = iWorkflowRecordCount

      ' Get Training Booking module parameters		
      Dim cmdTrainingBooking = New Command
      cmdTrainingBooking.CommandText = "sp_ASRIntGetTrainingBookingParameters"
      cmdTrainingBooking.CommandType = 4
      ' Stored Procedure
      cmdTrainingBooking.ActiveConnection = conX

      prmEmpTableID = cmdTrainingBooking.CreateParameter("empTableID", 3, 2)
      ' 3=integer, 2=output
      cmdTrainingBooking.Parameters.Append(prmEmpTableID)

      Dim prmCourseTableID = cmdTrainingBooking.CreateParameter("courseTableID", 3, 2)
      ' 3=integer, 2=output
      cmdTrainingBooking.Parameters.Append(prmCourseTableID)

      Dim prmCourseCancelDateColumnID = cmdTrainingBooking.CreateParameter("courseCancelDateColumnID", 3, 2)
      ' 3=integer, 2=output
      cmdTrainingBooking.Parameters.Append(prmCourseCancelDateColumnID)

      Dim prmTBTableID = cmdTrainingBooking.CreateParameter("tbTableID", 3, 2)
      ' 3=integer, 2=output
      cmdTrainingBooking.Parameters.Append(prmTBTableID)

      Dim prmTBTableSelect = cmdTrainingBooking.CreateParameter("tbTableSelect", 11, 2)
      ' 11=boolean, 2=output
      cmdTrainingBooking.Parameters.Append(prmTBTableSelect)

      Dim prmTBTableInsert = cmdTrainingBooking.CreateParameter("tbTableInsert", 11, 2)
      ' 11=boolean, 2=output
      cmdTrainingBooking.Parameters.Append(prmTBTableInsert)

      Dim prmTBTableUpdate = cmdTrainingBooking.CreateParameter("tbTableUpdate", 11, 2)
      ' 11=boolean, 2=output
      cmdTrainingBooking.Parameters.Append(prmTBTableUpdate)

      Dim prmTBStatusColumnID = cmdTrainingBooking.CreateParameter("tbStatusColumnID", 3, 2)
      ' 3=integer, 2=output
      cmdTrainingBooking.Parameters.Append(prmTBStatusColumnID)

      Dim prmTBStatusColumnUpdate = cmdTrainingBooking.CreateParameter("tbStatusColumnUpdate", 11, 2)
      ' 11=boolean, 2=output
      cmdTrainingBooking.Parameters.Append(prmTBStatusColumnUpdate)

      Dim prmTBCancelDateColumnID = cmdTrainingBooking.CreateParameter("tbCancelDateColumnID", 3, 2)
      ' 3=integer, 2=output
      cmdTrainingBooking.Parameters.Append(prmTBCancelDateColumnID)

      Dim prmTBCancelDateColumnUpdate = cmdTrainingBooking.CreateParameter("tbCancelDateColumnUpdate", 11, 2)
      ' 11=boolean, 2=output
      cmdTrainingBooking.Parameters.Append(prmTBCancelDateColumnUpdate)

      Dim prmTBStatusPExists = cmdTrainingBooking.CreateParameter("tbStatusPExists", 11, 2)
      ' 11=boolean, 2=output
      cmdTrainingBooking.Parameters.Append(prmTBStatusPExists)

      Dim prmWaitListTableID = cmdTrainingBooking.CreateParameter("waitListTableID", 3, 2)
      ' 3=integer, 2=output
      cmdTrainingBooking.Parameters.Append(prmWaitListTableID)

      Dim prmWaitListTableInsert = cmdTrainingBooking.CreateParameter("waitListTableInsert", 11, 2)
      ' 11=boolean, 2=output
      cmdTrainingBooking.Parameters.Append(prmWaitListTableInsert)

      Dim prmWaitListTableDelete = cmdTrainingBooking.CreateParameter("waitListTableDelete", 11, 2)
      ' 11=boolean, 2=output
      cmdTrainingBooking.Parameters.Append(prmWaitListTableDelete)

      Dim prmWaitListCourseTitleColumnID = cmdTrainingBooking.CreateParameter("waitListCourseTitleColumnID", 3, 2)
      ' 3=integer, 2=output
      cmdTrainingBooking.Parameters.Append(prmWaitListCourseTitleColumnID)

      Dim prmWaitListCourseTitleColumnUpdate = cmdTrainingBooking.CreateParameter("waitListCourseTitleColumnUpdate", 11, 2)
      ' 11=boolean, 2=output
      cmdTrainingBooking.Parameters.Append(prmWaitListCourseTitleColumnUpdate)

      Dim prmWaitListCourseTitleColumnSelect = cmdTrainingBooking.CreateParameter("waitListCourseTitleColumnSelect", 11, 2)
      ' 11=boolean, 2=output
      cmdTrainingBooking.Parameters.Append(prmWaitListCourseTitleColumnSelect)

      Dim prmBulkBookingDefaultViewID = cmdTrainingBooking.CreateParameter("bulkBookingDefaultViewID", 3, 2)
      ' 3=integer, 2=output
      cmdTrainingBooking.Parameters.Append(prmBulkBookingDefaultViewID)

      '		Set prmWaitListOverRideColumnID = cmdTrainingBooking.CreateParameter("WaitListOverRideColumnID", 3, 2) ' 3=integer, 2=output
      '		cmdTrainingBooking.Parameters.Append prmWaitListOverRideColumnID

      Err.Number = 0
      cmdTrainingBooking.Execute()

      If (Err.Number <> 0) Then
        Session("ErrorTitle") = "Login Page"
        Session("ErrorText") = "You could not login to the OpenHR database because of the following error:<p>" &
         FormatError(Err.Description)
        Return RedirectToAction("Loginerror")
      End If

      Session("TB_EmpTableID") = cmdTrainingBooking.Parameters("empTableID").Value

      Session("TB_CourseTableID") = cmdTrainingBooking.Parameters("courseTableID").Value
      Session("TB_CourseCancelDateColumnID") = cmdTrainingBooking.Parameters("courseCancelDateColumnID").Value

      Session("TB_TBTableID") = cmdTrainingBooking.Parameters("tbTableID").Value
      Session("TB_TBTableSelect") = cmdTrainingBooking.Parameters("tbTableSelect").Value
      Session("TB_TBTableInsert") = cmdTrainingBooking.Parameters("tbTableInsert").Value
      Session("TB_TBTableUpdate") = cmdTrainingBooking.Parameters("tbTableUpdate").Value
      Session("TB_TBStatusColumnID") = cmdTrainingBooking.Parameters("tbStatusColumnID").Value
      Session("TB_TBStatusColumnUpdate") = cmdTrainingBooking.Parameters("tbStatusColumnUpdate").Value
      Session("TB_TBCancelDateColumnID") = cmdTrainingBooking.Parameters("tbCancelDateColumnID").Value
      Session("TB_TBCancelDateColumnUpdate") = cmdTrainingBooking.Parameters("tbCancelDateColumnUpdate").Value
      Session("TB_TBStatusPExists") = cmdTrainingBooking.Parameters("tbStatusPExists").Value

      Session("TB_WaitListTableID") = cmdTrainingBooking.Parameters("waitListTableID").Value
      Session("TB_WaitListTableInsert") = cmdTrainingBooking.Parameters("waitListTableInsert").Value
      Session("TB_WaitListTableDelete") = cmdTrainingBooking.Parameters("waitListTableDelete").Value
      Session("TB_WaitListCourseTitleColumnID") = cmdTrainingBooking.Parameters("waitListCourseTitleColumnID").Value
      Session("TB_WaitListCourseTitleColumnUpdate") =
       cmdTrainingBooking.Parameters("waitListCourseTitleColumnUpdate").Value
      Session("TB_WaitListCourseTitleColumnSelect") =
       cmdTrainingBooking.Parameters("waitListCourseTitleColumnSelect").Value

      Session("TB_BulkBookingDefaultViewID") = cmdTrainingBooking.Parameters("bulkBookingDefaultViewID").Value
      'session("TB_WaitListOverRideColumnID") = cmdTrainingBooking.Parameters("WaitListOverRideColumnID").Value

      cmdTrainingBooking = Nothing

      If CStr(Session("TB_TBTableID")) = "" Then Session("TB_TBTableID") = 0

      'MH 07/07/2004: Moved from default.asp so background stuff only gets called on
      'login and not every time you go back to default.asp (as per request from JPD).

      Dim sTempPath As String, sBGImage As String = "", intBGPos As Short = 2, strRepeat As String, strBGPos As String

      Dim objUtilities = New Global.HR.Intranet.Server.Utilities

      objUtilities.Connection = Session("databaseConnection")
      sTempPath = Server.MapPath("~/pictures")
      sBGImage = objUtilities.GetBackgroundPicture(CStr(sTempPath))
      intBGPos = objUtilities.GetBackgroundPosition

      objUtilities.OfficeInitialise(CInt(Session("WordVer")), CInt(Session("ExcelVer")))
      Session("WordFormats") = objUtilities.OfficeGetCommonDialogFormatsWord
      Session("ExcelFormats") = objUtilities.OfficeGetCommonDialogFormatsExcel
      Session("WordFormatDefaultIndex") = objUtilities.OfficeGetDefaultIndexWord
      Session("ExcelFormatDefaultIndex") = objUtilities.OfficeGetDefaultIndexExcel
      Session("OfficeSaveAsValues") = objUtilities.OfficeGetSaveAsValues

      objUtilities = Nothing

      Select Case intBGPos
        Case 0
          'Top Left
          strRepeat = "no-repeat"
          strBGPos = "top left"

        Case 1
          'Top Right
          strRepeat = "no-repeat"
          strBGPos = "top right"

        Case 2
          'Centre
          strRepeat = "no-repeat"
          strBGPos = "center"

        Case 3
          'Left Tile
          strRepeat = "repeat-y"
          strBGPos = "left"

        Case 4
          'Right Tile
          strRepeat = "repeat-y"
          strBGPos = "right"

        Case 5
          'Top Tile
          strRepeat = "repeat-x"
          strBGPos = "top"

        Case 6
          'Bottom Tile
          strRepeat = "repeat-x"
          strBGPos = "bottom"

        Case 7
          'Tile
          strRepeat = "repeat"
          strBGPos = "center"

        Case Else
          'Centre
          strRepeat = "no-repeat"
          strBGPos = "center"

      End Select

      sBGImage = Url.Content("~/Pictures/" & sBGImage)

      If (Len(sTempPath) > 0) And (Len(sBGImage) > 0) Then
        Session("BodyTag") = "bgcolor=" & Session("ConvertedDesktopColour") & " STYLE=""height: 100%; background-image:url('" & sBGImage &
         "'); background-repeat:" & strRepeat & "; background-position:" & strBGPos & """" & vbCrLf
      Else
        Session("BodyTag") = "height: 100%; bgcolor=" & Session("ConvertedDesktopColour") & vbCrLf
      End If
      Session("BodyColour") = "bgcolor=" & Session("ConvertedDesktopColour") & vbCrLf


      ' Get the Find Window Block Size.
      Dim cmdFindSize = New Command
      cmdFindSize.CommandText = "sp_ASRIntGetSetting"
      cmdFindSize.CommandType = 4
      ' Stored procedure.
      cmdFindSize.ActiveConnection = Session("databaseConnection")

      prmSection = cmdFindSize.CreateParameter("section", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdFindSize.Parameters.Append(prmSection)
      prmSection.Value = "IntranetFindWindow"

      prmKey = cmdFindSize.CreateParameter("key", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdFindSize.Parameters.Append(prmKey)
      prmKey.Value = "BlockSize"

      prmDefault = cmdFindSize.CreateParameter("default", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdFindSize.Parameters.Append(prmDefault)
      prmDefault.Value = "1000"

      prmUserSetting = cmdFindSize.CreateParameter("userSetting", 11, 1)
      ' 11=bit, 1=input
      cmdFindSize.Parameters.Append(prmUserSetting)
      prmUserSetting.Value = 1

      prmResult = cmdFindSize.CreateParameter("result", 200, 2, 8000)
      ' 200=varchar, 2=output, 8000=size
      cmdFindSize.Parameters.Append(prmResult)

      Err.Number = 0
      cmdFindSize.Execute()

      Dim sResult = Trim(cmdFindSize.Parameters("result").Value)
      Session("FindRecords") = 1000
      If (IsNumeric(sResult)) Then
        If (CLng(sResult) > 0) Then
          Session("FindRecords") = CLng(sResult)
        End If
      End If
      cmdFindSize = Nothing

      ' Get the Primary Record Editing Start Mode.
      Dim cmdPrimaryStartMode = New Command
      cmdPrimaryStartMode.CommandText = "sp_ASRIntGetSetting"
      cmdPrimaryStartMode.CommandType = 4
      ' Stored procedure.
      cmdPrimaryStartMode.ActiveConnection = Session("databaseConnection")

      prmSection = cmdPrimaryStartMode.CreateParameter("section", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdPrimaryStartMode.Parameters.Append(prmSection)
      prmSection.Value = "RecordEditing"

      prmKey = cmdPrimaryStartMode.CreateParameter("key", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdPrimaryStartMode.Parameters.Append(prmKey)
      prmKey.Value = "Primary"

      prmDefault = cmdPrimaryStartMode.CreateParameter("default", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdPrimaryStartMode.Parameters.Append(prmDefault)
      prmDefault.Value = "3"

      prmUserSetting = cmdPrimaryStartMode.CreateParameter("userSetting", 11, 1)
      ' 11=bit, 1=input
      cmdPrimaryStartMode.Parameters.Append(prmUserSetting)
      prmUserSetting.Value = 1

      prmResult = cmdPrimaryStartMode.CreateParameter("result", 200, 2, 8000)
      ' 200=varchar, 2=output, 8000=size
      cmdPrimaryStartMode.Parameters.Append(prmResult)

      Err.Number = 0
      cmdPrimaryStartMode.Execute()
      Session("PrimaryStartMode") = CLng(cmdPrimaryStartMode.Parameters("result").Value)
      cmdPrimaryStartMode = Nothing

      ' Get the History Record Editing Start Mode.
      Dim cmdHistoryStartMode = New Command
      cmdHistoryStartMode.CommandText = "sp_ASRIntGetSetting"
      cmdHistoryStartMode.CommandType = 4
      ' Stored procedure.
      cmdHistoryStartMode.ActiveConnection = Session("databaseConnection")

      prmSection = cmdHistoryStartMode.CreateParameter("section", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdHistoryStartMode.Parameters.Append(prmSection)
      prmSection.Value = "RecordEditing"

      prmKey = cmdHistoryStartMode.CreateParameter("key", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdHistoryStartMode.Parameters.Append(prmKey)
      prmKey.Value = "History"

      prmDefault = cmdHistoryStartMode.CreateParameter("default", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdHistoryStartMode.Parameters.Append(prmDefault)
      prmDefault.Value = "3"

      prmUserSetting = cmdHistoryStartMode.CreateParameter("userSetting", 11, 1)
      ' 11=bit, 1=input
      cmdHistoryStartMode.Parameters.Append(prmUserSetting)
      prmUserSetting.Value = 1

      prmResult = cmdHistoryStartMode.CreateParameter("result", 200, 2, 8000)
      ' 200=varchar, 2=output, 8000=size
      cmdHistoryStartMode.Parameters.Append(prmResult)

      Err.Number = 0
      cmdHistoryStartMode.Execute()
      Session("HistoryStartMode") = CLng(cmdHistoryStartMode.Parameters("result").Value)
      cmdHistoryStartMode = Nothing

      ' Get the Lookup Record Editing Start Mode.
      Dim cmdLookupStartMode = New Command
      cmdLookupStartMode.CommandText = "sp_ASRIntGetSetting"
      cmdLookupStartMode.CommandType = 4
      ' Stored procedure.
      cmdLookupStartMode.ActiveConnection = Session("databaseConnection")

      prmSection = cmdLookupStartMode.CreateParameter("section", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdLookupStartMode.Parameters.Append(prmSection)
      prmSection.Value = "RecordEditing"

      prmKey = cmdLookupStartMode.CreateParameter("key", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdLookupStartMode.Parameters.Append(prmKey)
      prmKey.Value = "LookUp"

      prmDefault = cmdLookupStartMode.CreateParameter("default", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdLookupStartMode.Parameters.Append(prmDefault)
      prmDefault.Value = "2"

      prmUserSetting = cmdLookupStartMode.CreateParameter("userSetting", 11, 1)
      ' 11=bit, 1=input
      cmdLookupStartMode.Parameters.Append(prmUserSetting)
      prmUserSetting.Value = 1

      prmResult = cmdLookupStartMode.CreateParameter("result", 200, 2, 8000)
      ' 200=varchar, 2=output, 8000=size
      cmdLookupStartMode.Parameters.Append(prmResult)

      Err.Number = 0
      cmdLookupStartMode.Execute()
      Session("LookupStartMode") = CLng(cmdLookupStartMode.Parameters("result").Value)
      cmdLookupStartMode = Nothing

      ' Get the Quick Access Record Editing Start Mode.
      Dim cmdQuickAccessStartMode = New Command
      cmdQuickAccessStartMode.CommandText = "sp_ASRIntGetSetting"
      cmdQuickAccessStartMode.CommandType = 4
      ' Stored procedure.
      cmdQuickAccessStartMode.ActiveConnection = Session("databaseConnection")

      prmSection = cmdQuickAccessStartMode.CreateParameter("section", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdQuickAccessStartMode.Parameters.Append(prmSection)
      prmSection.Value = "RecordEditing"

      prmKey = cmdQuickAccessStartMode.CreateParameter("key", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdQuickAccessStartMode.Parameters.Append(prmKey)
      prmKey.Value = "QuickAccess"

      prmDefault = cmdQuickAccessStartMode.CreateParameter("default", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdQuickAccessStartMode.Parameters.Append(prmDefault)
      prmDefault.Value = "1"

      prmUserSetting = cmdQuickAccessStartMode.CreateParameter("userSetting", 11, 1)
      ' 11=bit, 1=input
      cmdQuickAccessStartMode.Parameters.Append(prmUserSetting)
      prmUserSetting.Value = 1

      prmResult = cmdQuickAccessStartMode.CreateParameter("result", 200, 2, 8000)
      ' 200=varchar, 2=output, 8000=size
      cmdQuickAccessStartMode.Parameters.Append(prmResult)

      Err.Number = 0
      cmdQuickAccessStartMode.Execute()
      Session("QuickAccessStartMode") = CLng(cmdQuickAccessStartMode.Parameters("result").Value)
      cmdQuickAccessStartMode = Nothing

      ' Get the Expression Colour setting.
      Dim cmdExprColourMode = New Command
      cmdExprColourMode.CommandText = "sp_ASRIntGetSetting"
      cmdExprColourMode.CommandType = 4
      ' Stored procedure.
      cmdExprColourMode.ActiveConnection = Session("databaseConnection")

      prmSection = cmdExprColourMode.CreateParameter("section", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdExprColourMode.Parameters.Append(prmSection)
      prmSection.Value = "ExpressionBuilder"

      prmKey = cmdExprColourMode.CreateParameter("key", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdExprColourMode.Parameters.Append(prmKey)
      prmKey.Value = "ViewColours"

      prmDefault = cmdExprColourMode.CreateParameter("default", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdExprColourMode.Parameters.Append(prmDefault)
      prmDefault.Value = "1"

      prmUserSetting = cmdExprColourMode.CreateParameter("userSetting", 11, 1)
      ' 11=bit, 1=input
      cmdExprColourMode.Parameters.Append(prmUserSetting)
      prmUserSetting.Value = 1

      prmResult = cmdExprColourMode.CreateParameter("result", 200, 2, 8000)
      ' 200=varchar, 2=output, 8000=size
      cmdExprColourMode.Parameters.Append(prmResult)

      Err.Number = 0
      cmdExprColourMode.Execute()
      Session("ExprColourMode") = CLng(cmdExprColourMode.Parameters("result").Value)
      cmdExprColourMode = Nothing

      ' Get the Expression Node Expansion setting.
      Dim cmdExprNodeMode = New Command
      cmdExprNodeMode.CommandText = "sp_ASRIntGetSetting"
      cmdExprNodeMode.CommandType = 4
      ' Stored procedure.
      cmdExprNodeMode.ActiveConnection = Session("databaseConnection")

      prmSection = cmdExprNodeMode.CreateParameter("section", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdExprNodeMode.Parameters.Append(prmSection)
      prmSection.Value = "ExpressionBuilder"

      prmKey = cmdExprNodeMode.CreateParameter("key", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdExprNodeMode.Parameters.Append(prmKey)
      prmKey.Value = "NodeSize"

      prmDefault = cmdExprNodeMode.CreateParameter("default", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdExprNodeMode.Parameters.Append(prmDefault)
      prmDefault.Value = "1"

      prmUserSetting = cmdExprNodeMode.CreateParameter("userSetting", 11, 1)
      ' 11=bit, 1=input
      cmdExprNodeMode.Parameters.Append(prmUserSetting)
      prmUserSetting.Value = 1

      prmResult = cmdExprNodeMode.CreateParameter("result", 200, 2, 8000)
      ' 200=varchar, 2=output, 8000=size
      cmdExprNodeMode.Parameters.Append(prmResult)

      Err.Number = 0
      cmdExprNodeMode.Execute()
      Session("ExprNodeMode") = CLng(cmdExprNodeMode.Parameters("result").Value)
      cmdExprNodeMode = Nothing

      'Support details - tel no
      Dim cmdSupportInfo = New Command
      cmdSupportInfo.CommandText = "sp_ASRIntGetSetting"
      cmdSupportInfo.CommandType = 4
      ' Stored procedure.
      cmdSupportInfo.ActiveConnection = Session("databaseConnection")

      prmSection = cmdSupportInfo.CreateParameter("section", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdSupportInfo.Parameters.Append(prmSection)
      prmSection.Value = "Support"

      prmKey = cmdSupportInfo.CreateParameter("key", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdSupportInfo.Parameters.Append(prmKey)
      prmKey.Value = "Telephone No"

      prmDefault = cmdSupportInfo.CreateParameter("default", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdSupportInfo.Parameters.Append(prmDefault)
      prmDefault.Value = "+44 (0)1582 714820"

      prmUserSetting = cmdSupportInfo.CreateParameter("userSetting", 11, 1)
      ' 11=bit, 1=input
      cmdSupportInfo.Parameters.Append(prmUserSetting)
      prmUserSetting.Value = 0

      prmResult = cmdSupportInfo.CreateParameter("result", 200, 2, 8000)
      ' 200=varchar, 2=output, 8000=size
      cmdSupportInfo.Parameters.Append(prmResult)

      Err.Number = 0
      cmdSupportInfo.Execute()
      Session("SupportTelNo") = cmdSupportInfo.Parameters("result").Value
      cmdSupportInfo = Nothing

      'Support details - Fax
      cmdSupportInfo = New Command
      cmdSupportInfo.CommandText = "sp_ASRIntGetSetting"
      cmdSupportInfo.CommandType = 4
      ' Stored procedure.
      cmdSupportInfo.ActiveConnection = Session("databaseConnection")

      prmSection = cmdSupportInfo.CreateParameter("section", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdSupportInfo.Parameters.Append(prmSection)
      prmSection.Value = "Support"

      prmKey = cmdSupportInfo.CreateParameter("key", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdSupportInfo.Parameters.Append(prmKey)
      prmKey.Value = "Fax"

      prmDefault = cmdSupportInfo.CreateParameter("default", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdSupportInfo.Parameters.Append(prmDefault)
      prmDefault.Value = "+44 (0)1582 714814 "

      prmUserSetting = cmdSupportInfo.CreateParameter("userSetting", 11, 1)
      ' 11=bit, 1=input
      cmdSupportInfo.Parameters.Append(prmUserSetting)
      prmUserSetting.Value = 0

      prmResult = cmdSupportInfo.CreateParameter("result", 200, 2, 8000)
      ' 200=varchar, 2=output, 8000=size
      cmdSupportInfo.Parameters.Append(prmResult)

      Err.Number = 0
      cmdSupportInfo.Execute()
      Session("SupportFax") = cmdSupportInfo.Parameters("result").Value
      cmdSupportInfo = Nothing

      'Support details - Email
      cmdSupportInfo = New Command
      cmdSupportInfo.CommandText = "sp_ASRIntGetSetting"
      cmdSupportInfo.CommandType = 4
      ' Stored procedure.
      cmdSupportInfo.ActiveConnection = Session("databaseConnection")

      prmSection = cmdSupportInfo.CreateParameter("section", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdSupportInfo.Parameters.Append(prmSection)
      prmSection.Value = "Support"

      prmKey = cmdSupportInfo.CreateParameter("key", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdSupportInfo.Parameters.Append(prmKey)
      prmKey.Value = "Email"

      prmDefault = cmdSupportInfo.CreateParameter("default", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdSupportInfo.Parameters.Append(prmDefault)
      prmDefault.Value = "service.delivery@advancedcomputersoftware.com"

      prmUserSetting = cmdSupportInfo.CreateParameter("userSetting", 11, 1)
      ' 11=bit, 1=input
      cmdSupportInfo.Parameters.Append(prmUserSetting)
      prmUserSetting.Value = 0

      prmResult = cmdSupportInfo.CreateParameter("result", 200, 2, 8000)
      ' 200=varchar, 2=output, 8000=size
      cmdSupportInfo.Parameters.Append(prmResult)

      Err.Number = 0
      cmdSupportInfo.Execute()
      Session("SupportEmail") = cmdSupportInfo.Parameters("result").Value
      cmdSupportInfo = Nothing

      'Support details - Webpage
      cmdSupportInfo = New Command
      cmdSupportInfo.CommandText = "sp_ASRIntGetSetting"
      cmdSupportInfo.CommandType = 4
      ' Stored procedure.
      cmdSupportInfo.ActiveConnection = Session("databaseConnection")

      prmSection = cmdSupportInfo.CreateParameter("section", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdSupportInfo.Parameters.Append(prmSection)
      prmSection.Value = "Support"

      prmKey = cmdSupportInfo.CreateParameter("key", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdSupportInfo.Parameters.Append(prmKey)
      prmKey.Value = "Webpage"

      prmDefault = cmdSupportInfo.CreateParameter("default", 200, 1, 8000)
      ' 200=varchar, 1=input, 8000=size
      cmdSupportInfo.Parameters.Append(prmDefault)
      prmDefault.Value = "service.delivery@advancedcomputersoftware.com"

      prmUserSetting = cmdSupportInfo.CreateParameter("userSetting", 11, 1)
      ' 11=bit, 1=input
      cmdSupportInfo.Parameters.Append(prmUserSetting)
      prmUserSetting.Value = 0

      prmResult = cmdSupportInfo.CreateParameter("result", 200, 2, 8000)
      ' 200=varchar, 2=output, 8000=size
      cmdSupportInfo.Parameters.Append(prmResult)

      Err.Number = 0
      cmdSupportInfo.Execute()
      Session("SupportWebpage") = cmdSupportInfo.Parameters("result").Value
      cmdSupportInfo = Nothing


      ' Get the configured Single record view ID. 	For the dashboard.
      Dim cmdModuleInfo = New Command
      cmdModuleInfo.CommandText = "spASRIntGetSingleRecordViewID"
      cmdModuleInfo.CommandType = 4
      ' Stored Procedure
      cmdModuleInfo.ActiveConnection = conX

      Dim prmTableID = cmdModuleInfo.CreateParameter("tableID", 3, 2)
      ' 3=integer, 2=output
      cmdModuleInfo.Parameters.Append(prmTableID)

      Dim prmViewID = cmdModuleInfo.CreateParameter("viewID", 3, 2)
      ' 3=integer, 2=output
      cmdModuleInfo.Parameters.Append(prmViewID)

      Err.Number = 0
      cmdModuleInfo.Execute()

      If (Err.Number <> 0) Then
        Session("ErrorTitle") = "Login Page - Module setup"
        Session("ErrorText") = "You could not login to the OpenHR database because of the following error :<p>" &
         FormatError(Err.Description)
        Return RedirectToAction("Loginerror", "Account")
      End If

      Session("SingleRecordTableID") = cmdModuleInfo.Parameters("tableID").Value
      Session("SingleRecordViewID") = cmdModuleInfo.Parameters("viewID").Value

      ' SSI Stuff:
      Session("SSILinkTableID") = 0
      Session("SSILinkViewID") = 0

      cmdModuleInfo = Nothing

      ' Get dashboard items

      Dim objNavigation = New Global.HR.Intranet.Server.clsNavigationLinks
      objNavigation.Connection = Session("databaseConnection")
      objNavigation.ClearLinks()

      objNavigation.SSITableID = Session("SingleRecordTableID")
      objNavigation.SSIViewID = Session("SingleRecordViewID")
      objNavigation.LoadLinks()
      objNavigation.LoadNavigationLinks()

      Dim objHypertextInfo As Collection = objNavigation.GetLinks(0)
      Dim objButtonInfo As Collection = objNavigation.GetLinks(1)
      Dim objDropdownInfo As Collection = objNavigation.GetLinks(2)


      Session("objHypertextInfo") = objHypertextInfo
      Session("objButtonInfo") = objButtonInfo
      Session("objDropdownInfo") = objDropdownInfo

      objNavigation = Nothing
      objHypertextInfo = Nothing
      objButtonInfo = Nothing
      objDropdownInfo = Nothing

      ' Enable/Disable SQL 2000 functions
      'Dim cmdVersion = New ADODB.Command
      'cmdVersion.CommandText = "spASRIntUDFFunctionsEnabled"
      'cmdVersion.CommandType = 4	' Stored procedure.
      'cmdVersion.ActiveConnection = Session("databaseConnection")

      'prmResult = cmdVersion.CreateParameter("result", 11, 2) ' 11=bit, 2=output
      'cmdVersion.Parameters.Append(prmResult)

      'Err.Number = 0
      'cmdVersion.Execute()
      Session("EnableSQL2000Functions") = False
      ' cmdVersion.Parameters("result").Value
      ' cmdVersion = Nothing

      If Session("WinAuth") Then
        ' Do not force password change for Windows Authenticated users.
        fForcePasswordChange = False
      End If

      If fForcePasswordChange = True Then
        ' Force password change only if there are no other users logged in with the same name.
        Dim cmdCheckUserSessions = New Command
        cmdCheckUserSessions.CommandText = "spASRGetCurrentUsersCountOnServer"
        cmdCheckUserSessions.CommandType = 4
        ' Stored procedure.
        cmdCheckUserSessions.ActiveConnection = Session("databaseConnection")

        Dim prmCount = cmdCheckUserSessions.CreateParameter("count", 3, 2)
        ' 3=integer, 2=output
        cmdCheckUserSessions.Parameters.Append(prmCount)

        prmUserName = cmdCheckUserSessions.CreateParameter("userName", 200, 1, 8000)
        ' 200=varchar, 1=input, 8000=size
        cmdCheckUserSessions.Parameters.Append(prmUserName)
        prmUserName.Value = Session("Username")

        Err.Number = 0
        cmdCheckUserSessions.Execute()

        Dim iUserSessionCount = CLng(cmdCheckUserSessions.Parameters("count").Value)
        cmdCheckUserSessions = Nothing

        If iUserSessionCount < 2 Then
          Return RedirectToAction("ForcedPasswordChange")
        Else
          Return RedirectToAction("Main", "Home")
        End If
      Else
        If isWidgetLogin Then
          ' call the widget function if applicable
          'Dim str As String = ASRIntranetFunctions.GetDBValue(Session("databaseConnection"))

        Else
          Try

            ' grab some more info for the dashboard						
            Dim sErrorDescription = ""

            ' Get the self-service record ID.
            Dim cmdSSRecord = New Command
            cmdSSRecord.CommandText = "spASRIntGetSelfServiceRecordID"
            'Get Single Record ID
            cmdSSRecord.CommandType = 4
            ' Stored Procedure
            cmdSSRecord.ActiveConnection = Session("databaseConnection")

            Dim prmRecordID = cmdSSRecord.CreateParameter("@piRecordID", 3, 2)
            ' 3=integer, 2=output
            cmdSSRecord.Parameters.Append(prmRecordID)

            Dim prmRecordCount = cmdSSRecord.CreateParameter("@piRecordCount", 3, 2)
            ' 3=integer, 2=output
            cmdSSRecord.Parameters.Append(prmRecordCount)

            prmViewID = cmdSSRecord.CreateParameter("@piViewID", 3, 1)
            ' 3=integer, 1=input
            cmdSSRecord.Parameters.Append(prmViewID)
            prmViewID.Value = CleanNumeric(Session("SingleRecordViewID"))

            cmdSSRecord.Execute()

            If (Err.Number <> 0) Then
              sErrorDescription = "Unable to get the personnel record ID." & vbCrLf & FormatError(Err.Description)
            End If

            If Len(sErrorDescription) = 0 Then
              If cmdSSRecord.Parameters("@piRecordCount").Value = 1 Then
                ' Only one record.
                Session("TopLevelRecID") = CLng(cmdSSRecord.Parameters("@piRecordID").Value)
              Else
                If cmdSSRecord.Parameters("@piRecordCount").Value = 0 Then
                  ' No personnel record. 
                  Session("TopLevelRecID") = 0
                Else
                  ' More than one personnel record.
                  sErrorDescription = "You have access to more than one record in the defined Single-record view."
                End If
              End If
            End If

            cmdSSRecord = Nothing


            ' Get the record description.
            Dim sRecDesc = ""
            Dim cmdGetRecordDesc = New Command

            cmdGetRecordDesc.CommandText = "sp_ASRIntGetRecordDescription"
            cmdGetRecordDesc.CommandType = 4
            ' Stored procedure
            cmdGetRecordDesc.ActiveConnection = Session("databaseConnection")

            prmTableID = cmdGetRecordDesc.CreateParameter("tableID", 3, 1)
            ' 3 = integer, 1 = input
            cmdGetRecordDesc.Parameters.Append(prmTableID)
            prmTableID.Value = CleanNumeric(Session("SingleRecordTableID"))
            ' cleanNumeric(Session("tableID"))

            prmRecordID = cmdGetRecordDesc.CreateParameter("recordID", 3, 1)
            ' 3 = integer, 1 = input
            cmdGetRecordDesc.Parameters.Append(prmRecordID)
            prmRecordID.Value = CleanNumeric(Session("TopLevelRecID"))

            Dim prmParentTableID = cmdGetRecordDesc.CreateParameter("parentTableID", 3, 1)
            ' 3 = integer, 1 = input
            cmdGetRecordDesc.Parameters.Append(prmParentTableID)
            prmParentTableID.Value = CleanNumeric(Session("parentTableID"))

            Dim prmParentRecordID = cmdGetRecordDesc.CreateParameter("parentRecordID", 3, 1)
            ' 3=integer, 1=input
            cmdGetRecordDesc.Parameters.Append(prmParentRecordID)
            prmParentRecordID.Value = CleanNumeric(Session("parentRecordID"))

            Dim prmRecordDesc = cmdGetRecordDesc.CreateParameter("recordDesc", 200, 2, 8000)
            ' 200=varchar, 2=output, 8000=size
            cmdGetRecordDesc.Parameters.Append(prmRecordDesc)

            Const DEADLOCK_ERRORNUMBER = -2147467259
            Const DEADLOCK_MESSAGESTART = "YOUR TRANSACTION (PROCESS ID #"
            Const DEADLOCK_MESSAGEEND =
             ") WAS DEADLOCKED WITH ANOTHER PROCESS AND HAS BEEN CHOSEN AS THE DEADLOCK VICTIM. RERUN YOUR TRANSACTION."
            Const DEADLOCK2_MESSAGESTART = "TRANSACTION (PROCESS ID "
            Const DEADLOCK2_MESSAGEEND = ") WAS DEADLOCKED ON "

            Dim sErrMsg As String = ""
            Dim fOK = True
            Dim fDeadlock = True
            Dim iRetryCount = 0
            Dim iRETRIES = 0


            Do While fDeadlock
              fDeadlock = False

              cmdGetRecordDesc.ActiveConnection.Errors.Clear()

              cmdGetRecordDesc.Execute()

              If cmdGetRecordDesc.ActiveConnection.Errors.Count > 0 Then
                For iLoop = 1 To cmdGetRecordDesc.ActiveConnection.Errors.Count
                  sErrMsg = FormatError(cmdGetRecordDesc.ActiveConnection.Errors.Item(iLoop - 1).Description)

                  If (cmdGetRecordDesc.ActiveConnection.Errors.Item(iLoop - 1).Number = DEADLOCK_ERRORNUMBER) And
                   (((UCase(Left(sErrMsg, Len(DEADLOCK_MESSAGESTART))) = DEADLOCK_MESSAGESTART) And
                   (UCase(Right(sErrMsg, Len(DEADLOCK_MESSAGEEND))) = DEADLOCK_MESSAGEEND)) Or
                    ((UCase(Left(sErrMsg, Len(DEADLOCK2_MESSAGESTART))) = DEADLOCK2_MESSAGESTART) And
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
                    sErrorDescription = sErrorDescription & vbCrLf &
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

            If Len(sErrorDescription) = 0 Then
              Session("recdesc") = cmdGetRecordDesc.Parameters("recordDesc").Value
            End If

            cmdGetRecordDesc = Nothing

          Catch ex As Exception

          End Try

          Dim cookie = New HttpCookie("Login")
          cookie.Expires = DateTime.Now.AddYears(1)
          cookie.HttpOnly = True
          cookie("User") = Request.Form("txtUserNameCopy")
          cookie("Database") = Request.Form("txtDatabase")
          cookie("Server") = Request.Form("txtServer")
          cookie("WindowsAuthentication") = Request.Form("chkWindowsAuthentication")
          Response.Cookies.Add(cookie)

          If Session("SuperUser") = True Then
            Return RedirectToAction("Main", "Home")
          Else
            Return RedirectToAction("LinksMain", "Home")
          End If

        End If

      End If


      Return RedirectToAction("login", "account")
    End Function

    <HttpPost()>
    Function ForcedPasswordChange_Submit(value As FormCollection) As ActionResult

      On Error Resume Next

      Dim strErrorMessage As String = ""
      Dim sConnectString As String = ""

      Dim fSubmitPasswordChange = (Len(Request.Form("txtGotoPage")) = 0)

      If fSubmitPasswordChange Then
        ' Force password change only if there are no other users logged in with the same name.
        Dim cmdCheckUserSessions = CreateObject("ADODB.Command")
        cmdCheckUserSessions.CommandText = "spASRGetCurrentUsersCountOnServer"
        cmdCheckUserSessions.CommandType = 4 ' Stored procedure.
        cmdCheckUserSessions.ActiveConnection = Session("databaseConnection")

        Dim prmCount = cmdCheckUserSessions.CreateParameter("count", 3, 2) ' 3=integer, 2=output
        cmdCheckUserSessions.Parameters.Append(prmCount)

        Dim prmUserName = cmdCheckUserSessions.CreateParameter("userName", 200, 1, 8000)  ' 200=varchar, 1=input, 8000=size
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

          ' SQL Native Client Stuff
          If Err.Number = 3709 Then
            Err.Clear()

            Dim conX = CreateObject("ADODB.Connection")
            conX.ConnectionTimeout = 60

            Dim objSettings = New Global.HR.Intranet.Server.clsSettings

            Select Case objSettings.GetSQLNCLIVersion
              Case 9
                sConnectString = "Provider=SQLNCLI;"
              Case 10
                sConnectString = "Provider=SQLNCLI10;"
              Case 11
                sConnectString = "Provider=SQLNCLI11;"
            End Select
            objSettings = Nothing

            sConnectString = sConnectString & "DataTypeCompatibility=80;MARS Connection=True;" & Session("SQL2005Force") & _
               ";Old Password='" & Replace(sCurrentPassword, "'", "''") & "';Password='" & Replace(sNewPassword, "'", "''") & "'"

            conX.open(sConnectString)

            If Err.Number <> 0 Then
              If Err.Number <> 3706 Then   ' 3706 = Provider not found
                strErrorMessage = Err.Description
              End If
              Session("ErrorTitle") = "Change Password Page"
              Session("ErrorText") = strErrorMessage
              ' Return RedirectToAction("Loginerror", "Account")
              Return RedirectToAction("Loginerror", "Account")
            Else
              conX.close()
              Session("MessageTitle") = "Change Password Page"
              Session("MessageText") = "Password changed successfully. You may now login."
              ' Response.Redirect("loginmessage")
              Return RedirectToAction("Loginmessage", "Account")
            End If
          End If

          If Err.Number <> 0 Then
            Session("ErrorTitle") = "Change Password Page"
            Session("ErrorText") = "You could not change your password because of the following error:<p>" & Err.Description
            Return RedirectToAction("Loginerror", "Account")
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
              Session("ErrorText") = "You could not change your password because of the following error:<p>" & Err.Description
              Return RedirectToAction("Loginerror", "Account")
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
                Session("ErrorText") = "You could not change your password because of the following error:<p>" & Err.Description
                Return RedirectToAction("Loginerror", "Account")
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
            cmdCreateCache.ActiveConnection = conX
            cmdCreateCache.execute()
            cmdCreateCache = Nothing

            Session("MessageTitle") = "Change Password Page"
            Session("MessageText") = "Password changed successfully."
            Response.Redirect("loginmessage")
            'Response.Redirect("confirmok")
          End If
        Else
          Session("ErrorTitle") = "Change Password Page"
          Dim sErrorText = "You could not change your password.<p>The account is currently being used by "
          If iUserSessionCount > 2 Then
            sErrorText = sErrorText & iUserSessionCount & " users"
          Else
            sErrorText = sErrorText & "another user"
          End If
          sErrorText = sErrorText & " in the system."
          Session("ErrorText") = sErrorText

          Return RedirectToAction("Loginerror", "Account")
        End If
      Else
        ' Go to the main page.
        Response.Redirect("main")
      End If
    End Function

    Function ForcedPasswordChange() As ActionResult
      Return View()
    End Function

    Function Loginerror() As ActionResult
      Return View()
    End Function

    Function Loginmessage() As ActionResult
      Return View()
    End Function

    Function AboutHRPro() As ActionResult
      Return View()
    End Function

  End Class

  Public Class JsonData_DBValue
    Public Property Caption() As String
      Get
        Return m_Caption
      End Get
      Set(value As String)
        m_Caption = value
      End Set
    End Property

    Private m_Caption As String

    Public Property DBValue() As String
      Get
        Return m_DBValue
      End Get
      Set(value As String)
        m_DBValue = value
      End Set
    End Property

    Private m_DBValue As String

    Public Property Formatting_Suffix() As String
      Get
        Return m_Formatting_Suffix
      End Get
      Set(value As String)
        m_Formatting_Suffix = value
      End Set
    End Property

    Private m_Formatting_Suffix As String

    Public Property Formatting_Prefix() As String
      Get
        Return m_Formatting_Prefix
      End Get
      Set(value As String)
        m_Formatting_Prefix = value
      End Set
    End Property

    Private m_Formatting_Prefix As String
  End Class

  Public Class navigationLinks
    Private m_linkType As Integer
    Private m_linkOrder As Integer
    Private m_prompt As String
    Private m_text As String
    Private m_element_Type As Integer
    Private m_ID As Long

    Sub New(p_linkType As Integer, p_linkOrder As Integer, p_prompt As String, p_text As String, p_element_Type As Integer,
      p_ID As Long)

      linkType = p_linkType
      linkOrder = p_linkOrder
      prompt = p_prompt
      text = p_text
      element_Type = p_element_Type
      ID = p_ID
    End Sub


    Public Property linkType() As Integer
      Get
        Return m_linkType
      End Get
      Set(value As Integer)
        m_linkType = value
      End Set
    End Property

    Public Property linkOrder() As Integer
      Get
        Return m_linkOrder
      End Get
      Set(value As Integer)
        m_linkOrder = value
      End Set
    End Property

    Public Property prompt() As String
      Get
        Return m_prompt
      End Get
      Set(value As String)
        m_prompt = value
      End Set
    End Property

    Public Property text As String
      Get
        Return m_text
      End Get
      Set(value As String)
        m_text = value
      End Set
    End Property

    Public Property element_Type() As Integer
      Get
        Return m_element_Type
      End Get
      Set(value As Integer)
        m_element_Type = value
      End Set
    End Property

    Public Property ID() As Long
      Get
        Return m_ID
      End Get
      Set(value As Long)
        m_ID = value
      End Set
    End Property
  End Class
End Namespace

