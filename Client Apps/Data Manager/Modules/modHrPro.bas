Attribute VB_Name = "modHrPro"
Option Explicit

' Generic Postcode Structure (Moved from modAFDSpecifics)
Public Type PostCode
  PostCode As String * 8         ' Postcode
  FirstName As String * 30       'First Name
  Initial2 As String * 6         'Initial of Second Forename
  Surname As String * 30         'Surname
  Building As String * 60        'Building Name
  HouseNo As String * 10         'Building Number
  Street As String * 120         'Street or Thoroughfare (includes Dependant Thoroughfare)
  Locality As String * 70        'Locality (includes Double Dependant Locality)
  Town As String * 30            'Post Town
  County As String * 30          'County Name according to Postal Authority
  Phone As String * 20           'Telephone No where known, incl STD Code
  Organisation As String * 120   'Organisation Name (includes Department, if any)
End Type


'Public Constants
Public Const INT_MASK As String = "##########"
Public Const INT_MASK_NOBLANK As String = "#########0"
Public Const COL_GREY As Long = &H8000000F

'Public Const SQLMAILNOTSTARTEDMESSAGE = "SQL MAIL SESSION IS NOT STARTED."
'Public Const SQLMAILEXECUTEDENIED = "EXECUTE PERMISSION DENIED ON OBJECT 'XP_SENDMAIL', DATABASE 'MASTER', OWNER 'DBO'."

Public Const DEADLOCK_ERRORNUMBER = -2147467259
Public Const DEADLOCK_MESSAGESTART = "YOUR TRANSACTION (PROCESS ID #"
Public Const DEADLOCK_MESSAGEEND = ") WAS DEADLOCKED WITH ANOTHER PROCESS AND HAS BEEN CHOSEN AS THE DEADLOCK VICTIM. RERUN YOUR TRANSACTION."
Public Const DEADLOCK2_MESSAGESTART = "TRANSACTION (PROCESS ID "
Public Const DEADLOCK2_MESSAGEEND = ") WAS DEADLOCKED ON "

Public Const CONNECTIONBROKEN_MESSAGE = "GENERAL NETWORK ERROR. CHECK YOUR NETWORK DOCUMENTATION."
Public Const FRAMEWORK_MESSAGE = "An error occurred in the Microsoft .NET Framework while trying to load assembly id"

' Constants.
Public Const ODBCDRIVER As String = "SQL Server"

Private Declare Function GetModuleFileNameA Lib "kernel32" (ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hWnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long




Public Const gblnDEFAULTTITLEGRIDLINES As Boolean = False
Public Const gblnDEFAULTTITLEBOLD As Boolean = True
Public Const gblnDEFAULTTITLEUNDERLINE As Boolean = False
Public Const glngDEFAULTTITLEBACKCOLOUR As Long = 16777215    'vbWhite
Public Const glngDEFAULTTITLEFORECOLOUR As Long = 6697779     'GetColour("Midnight Blue")

Public Const gblnDEFAULTHEADINGGRIDLINES As Boolean = True
Public Const gblnDEFAULTHEADINGBOLD As Boolean = True
Public Const gblnDEFAULTHEADINGUNDERLINE As Boolean = False
Public Const glngDEFAULTHEADINGBACKCOLOUR As Long = 16248553  'GetColour("Dolphin Blue")
Public Const glngDEFAULTHEADINGFORECOLOUR As Long = 6697779   'GetColour("Midnight Blue")

Public Const gblnDEFAULTDATAGRIDLINES As Boolean = True
Public Const gblnDEFAULTDATABOLD As Boolean = False
Public Const gblnDEFAULTDATAUNDERLINE As Boolean = False
Public Const glngDEFAULTDATABACKCOLOUR As Long = 15988214     'GetColour("Pale Grey")
Public Const glngDEFAULTDATAFORECOLOUR As Long = 6697779      'GetColour("Midnight Blue")

Public gstrDebugOutputFile As String


'Public gsngTimer As Single

Public Sub DisplayApplication()
  'JPD 20030908 Fault 5756
  
  'JPD 20030917 Fault 6991
  If (glngWindowLeft = 0) And _
    (glngWindowTop = 0) And _
    (glngWindowWidth = 0) And _
    (glngWindowHeight = 0) Then
    
    Exit Sub
  End If
  
  frmMain.WindowState = IIf(giWindowState = vbMinimized, vbNormal, giWindowState)
  If frmMain.WindowState = vbNormal Then
    frmMain.Left = glngWindowLeft
    frmMain.Top = glngWindowTop
    frmMain.Width = glngWindowWidth
    frmMain.Height = glngWindowHeight
  End If

End Sub


Public Function EnableActiveBar(pobjAB As ActiveBar, pbEnable As Boolean) As Boolean

  'Function enables or disbales all the tools in the passed ActiveBar control according to the
  'passed pbEnable parameter.
  
  'TM20020612 Fault 2302
  Dim Tool As ActiveBarLibraryCtl.Tool


  'MH20040218 Fault 8080
  'Prevent MDI being inadvertidly reloaded
  If gcoTablePrivileges Is Nothing Then
    EnableActiveBar = True
    Exit Function
  End If


  On Error GoTo Error_Trap
  
  For Each Tool In pobjAB.Tools
    With Tool
      .Enabled = pbEnable
    End With
  Next Tool
  
  EnableActiveBar = True

TidyUpAndExit:
  Set Tool = Nothing
  Exit Function
  
Error_Trap:
  EnableActiveBar = False
  GoTo TidyUpAndExit
  
End Function

Public Function CheckVersion(psServerName As String) As Boolean
  ' Check that the database version is the right one for this application's version.
  ' If everything matches then return TRUE.
  ' If not, try to update the database.
  ' If the database can be updated return TRUE, else return FALSE.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fVersionOK As Boolean
  Dim fRefreshStoredProcedures As Boolean
  'Dim iPos1 As Integer
  'Dim iPos2 As Integer
  'Dim iResult As Integer
  Dim iMajorAppVersion As Integer
  Dim iMinorAppVersion As Integer
  Dim iRevisionAppVersion As Integer
  Dim lngDBVersion As Long
  'Dim lngLastVersion As Long
  'Dim sSQL As String
  'Dim sReadString As String
  'Dim sUpdateScript As String
  Dim sDBVersion As String
  Dim rsInfo As Recordset
  'Dim rsInfo2 As Recordset
  Dim blnNewStyleVersionNo As Boolean
  Dim iResponse As Integer

  fOK = True
  fVersionOK = False
  fRefreshStoredProcedures = False
    
    
  'sDBVersion = GetSystemSetting("Database", "Version", vbNullString)
  sDBVersion = GetDBVersion

  If Len(sDBVersion) = 0 Then
    fOK = False

    COAMsgBox "Error checking version compatibility." & vbCrLf & _
      "Version number not found.", _
      vbOKOnly + vbExclamation, Application.Name
  Else
    iMajorAppVersion = Val(Split(sDBVersion, ".")(0))
    iMinorAppVersion = Val(Split(sDBVersion, ".")(1))
    
    blnNewStyleVersionNo = (UBound(Split(sDBVersion, ".")) = 1)
    If Not blnNewStyleVersionNo Then
      iRevisionAppVersion = Val(Split(sDBVersion, ".")(2))
    End If
  End If

  fVersionOK = ASRDEVELOPMENT
  If fOK Then
    ' Check the application version against the one for the current database.
    If (app.Major = iMajorAppVersion) And _
      (app.Minor = iMinorAppVersion) And _
      (app.Revision = iRevisionAppVersion Or blnNewStyleVersionNo) Then
      ' Application and database versions match.
      fVersionOK = True
    End If
  End If
    
    
  If fOK Then
    ' Check the application version against the one for the current database.
    If (app.Major < iMajorAppVersion) Or _
      ((app.Major = iMajorAppVersion) And (app.Minor < iMinorAppVersion)) Or _
      ((app.Major = iMajorAppVersion) And (app.Minor = iMinorAppVersion) And (app.Revision < iRevisionAppVersion And Not blnNewStyleVersionNo)) Then
      ' Application is too old for the database.

      If ASRDEVELOPMENT Then
        iResponse = COAMsgBox("The application is out of date." & vbCrLf & _
          "Contact your administrator for a new version of the application." & vbCrLf & vbCrLf & _
          "Database Name : " & gsDatabaseName & vbCrLf & _
          "Database Version : " & sDBVersion & vbCrLf & vbCrLf & _
          "Application Version : " & CStr(app.Major) & "." & CStr(app.Minor) & vbCrLf & _
          "(ASRDevelopment - Continue anyway?)", vbExclamation + vbYesNo, Application.Name)

        fOK = (iResponse = vbYes)
      Else
        COAMsgBox "The application is out of date." & vbCrLf & _
          "Contact your administrator for a new version of the application." & vbCrLf & vbCrLf & _
          "Database Name : " & gsDatabaseName & vbCrLf & _
          "Database Version : " & sDBVersion & vbCrLf & vbCrLf & _
          "Application Version : " & CStr(app.Major) & "." & CStr(app.Minor), _
          vbExclamation + vbOKOnly, Application.Name

        fOK = False
      End If

    End If
  End If

  If fOK Then
    If (app.Major > iMajorAppVersion) Or _
      ((app.Major = iMajorAppVersion) And (app.Minor > iMinorAppVersion)) Or _
      ((app.Major = iMajorAppVersion) And (app.Minor = iMinorAppVersion) And (app.Revision > iRevisionAppVersion And Not blnNewStyleVersionNo)) Then
      ' Database is too old for the application. Try to update the database.
      
      If ASRDEVELOPMENT Then
        iResponse = COAMsgBox("The database is out of date." & vbCrLf & _
          "Please ask the System Administrator to update the database in the System Manager." & vbCrLf & vbCrLf & _
          "Database Name : " & gsDatabaseName & vbCrLf & _
          "Database Version : " & sDBVersion & vbCrLf & vbCrLf & _
          "Application Version : " & CStr(app.Major) & "." & CStr(app.Minor) & vbCrLf & _
          "(ASRDevelopment - Continue anyway?)", _
          vbExclamation + vbYesNo, Application.Name)
        
          fOK = (iResponse = vbYes)
      Else
         COAMsgBox "The database is out of date." & vbCrLf & _
          "Please ask the System Administrator to update the database in the System Manager." & vbCrLf & vbCrLf & _
          "Database Name : " & gsDatabaseName & vbCrLf & _
          "Database Version : " & sDBVersion & vbCrLf & vbCrLf & _
          "Application Version : " & CStr(app.Major) & "." & CStr(app.Minor), _
          vbExclamation + vbOKOnly, Application.Name
          
          fOK = False
      End If
    End If
  End If


  If fOK Then
    ' Check if a new version of the application is required due to an Intranet update
    
    sDBVersion = GetSystemSetting("Database", "Minimum Version", vbNullString)
    If Len(sDBVersion) > 0 Then
      
      iMajorAppVersion = Val(Split(sDBVersion, ".")(0))
      iMinorAppVersion = Val(Split(sDBVersion, ".")(1))
      
      blnNewStyleVersionNo = (UBound(Split(sDBVersion, ".")) = 1)
      If Not blnNewStyleVersionNo Then
        iRevisionAppVersion = Val(Split(sDBVersion, ".")(2))
      End If
      
      If (app.Major < iMajorAppVersion) Or _
        ((app.Major = iMajorAppVersion) And (app.Minor < iMinorAppVersion)) Or _
        ((app.Major = iMajorAppVersion) And (app.Minor = iMinorAppVersion) And (app.Revision < iRevisionAppVersion And Not blnNewStyleVersionNo)) Then

        COAMsgBox "The application is now out of date due to an update to the intranet module." & vbCrLf & _
          "Contact your administrator for a new version of the application.", _
          vbOKOnly + vbExclamation, Application.Name
        
        If Not ASRDEVELOPMENT Then
          fVersionOK = False
          fOK = fVersionOK
        End If

      End If
    End If
  End If

' Get the SQL Server version number.
  glngSQLVersion = 0
  Set rsInfo = datGeneral.GetReadOnlyRecords("master..xp_msver ProductVersion")
  With rsInfo
    If Not (.BOF And .EOF) Then
      glngSQLVersion = Val(!character_value)
    End If
    .Close
  End With
  Set rsInfo = Nothing

  ' If the platform has changed tag the refresh stored procedures flag
  If fOK Then
    fOK = CheckPlatform
  End If

  If fOK Then
    fRefreshStoredProcedures = (GetSystemSetting("Database", "RefreshStoredProcedures", 0) = 1)
  
    If fRefreshStoredProcedures Then
      ' Tell the user that the System manager needs to be run, and changes saved
      ' before this application can run.
      If Not ASRDEVELOPMENT Then
        fOK = False
      End If

      COAMsgBox "The database is out of date." & vbCrLf & _
        "Please ask the System Administrator to save the update in the System Manager.", _
        vbOKOnly + vbExclamation, Application.Name
    End If
  End If

  ' Are UDF functions enabled
  gbEnableUDFFunctions = datGeneral.EnableUDFFunctions

  ' If fOK and fVersionOK are true then the application and databases versions match.
TidyUpAndExit:
  If Not fOK And Not ASRDEVELOPMENT Then
    fVersionOK = False
    Screen.MousePointer = vbDefault
  End If
  
  CheckVersion = fVersionOK
  Exit Function
  
ErrorTrap:
  If (Err.Number = 75) Or (Err.Number = 76) Then
    COAMsgBox "The database is out of date." & vbCrLf & _
      "Unable to update the database as the required update script cannot be found.", _
      vbOKOnly + vbExclamation, Application.Name
  Else
    COAMsgBox "Error checking database and application versions." & vbCrLf & _
      Err.Description, _
      vbOKOnly + vbExclamation, Application.Name
  End If
  fOK = False
  Resume TidyUpAndExit
  
End Function

Sub Main()
  
  If (InStr(LCase(Command$), "/debug=true") > 0) Then
    gstrDebugOutputFile = app.Path & "\debug.txt"
    If Dir(gstrDebugOutputFile) <> vbNullString Then
      Kill gstrDebugOutputFile
    End If
  End If
  
  
  ' Allow initialisation of XP style controls
  DebugOutput "modHRPro.Main", "InitCommonControls"
  InitCommonControls
  
  'Instantiate Application class
  DebugOutput "modHRPro.Main", "Set Application"
  Set Application = New DataMgr.Application

  'Instantiate Application class
  DebugOutput "modHRPro.Main", "Set Database"
  Set Database = New DataMgr.Database
  
  DebugOutput "modHRPro.Main", "Set General"
  Set datGeneral = New DataMgr.clsGeneral
  
  DebugOutput "modHRPro.Main", "Set Email"
  Set objEmail = New clsEmail
  
  ' Create Current User class
  'Set gobjCurrentUser = New DataMgr.clsUser
  
  'Instantiate User Interface class
  DebugOutput "modHRPro.Main", "Set UI"
  Set UI = New DataMgr.UI
  
  'Instantiate Progress Bar class
  'Set gobjProgress = New COA_Progress
  DebugOutput "modHRPro.Main", "Set Progress"
  Set gobjProgress = New clsProgress
  gobjProgress.StyleResource = CodeJockStylePath
  gobjProgress.StyleIni = CodeJockStyleIni

  'Instantiate UtilityRunLog class
  DebugOutput "modHRPro.Main", "Set EventLog"
  Set gobjEventLog = New clsEventLog
  
  ' Initialise the standard error handler
  DebugOutput "modHRPro.Main", "Set ErrorStack"
  Set gobjErrorStack = New clsErrorStack
  
  ' Initialise the generic data access class
  DebugOutput "modHRPro.Main", "Set DataAccess"
  Set gobjDataAccess = New DataMgr.clsDataAccess
  
  ' Initialse the performance monitor
  DebugOutput "modHRPro.Main", "Set Performance"
  Set gobjPerformance = New DataMgr.clsPerformance
  
  ' Default logged on user information
  DebugOutput "modHRPro.Main", "Set WindowsCurrent"
  gstrWindowsCurrentDomain = Environ("USERDOMAIN")
  gstrWindowsCurrentUser = Environ("USERNAME")
  
  ' Do we read the default toolbars
  gbReadToolbarDefaults = False
  
  ' Options for which output types are available
'  gbAllowOutput_Word = Office_IsWordInstalled
'  gbAllowOutput_Excel = Office_IsExcelInstalled
'
'  ' Versions of Office
'  giOfficeVersion_Word = IIf(gbAllowOutput_Word, Office_WordVersion, 0)
'  giOfficeVersion_Excel = IIf(gbAllowOutput_Excel, Office_ExcelVersion, 0)
  
'  gblnStartupMSOffice = (InStr(LCase(Command$), "/msoffice=false") > 0)
'  If Not gblnStartupMSOffice Then
'    ' Versions of Office
'    giOfficeVersion_Word = Office_WordVersion
'    giOfficeVersion_Excel = Office_ExcelVersion
'  End If
'  gbAllowOutput_Word = IIf(giOfficeVersion_Word > 0, True, False)
'  gbAllowOutput_Excel = IIf(giOfficeVersion_Excel > 0, True, False)

  ' If we get problems, just in case...
  gbDisableCodeJock = (InStr(LCase(Command$), "/skin=false") > 0)
  gbActivateJobServer = (InStr(LCase(Command$), "/jobseek=true") > 0)

  If app.StartMode = vbSModeAutomation Then
    'If started via OLE automation, return control back to client application
    Exit Sub
  ElseIf app.StartMode = vbSModeStandalone Then
    'Login to database
    DebugOutput "modHRPro.Main", "Application.Login"
    If Application.Login Then
      'Display splash screen
      DebugOutput "modHRPro.Main", "Show frmSplash"
      frmSplash.Show
      frmSplash.Refresh
      
'      If gbAllowOutput_Word Then
'        'WdSaveFormat.wdFormatDocumentDefault
'        'WdSaveFormat.wdFormatDocument97
'        giOfficeSaveVersion_Word = GetSystemSetting("output", "save version word", WdSaveFormat.wdFormatDocument)
'        gsOfficeFileFilter_Word = GetSystemSetting("output", "file filter word", "Word Document (*.doc)|*.doc")
'        gsOfficeTemplateFilter_Word = GetSystemSetting("output", "template filter word", "Word Template (*.doc;*.dot)|*.doc;*.dot")
'      End If
'
'      If gbAllowOutput_Excel Then
'        'XlFileFormat.xlWorkbookDefault
'        'XlFileFormat.xlExcel8
'        giOfficeSaveVersion_Excel = GetSystemSetting("output", "save version excel", 56)  'XlFileFormat.xlExcel8)
'        gsOfficeFileFilter_Excel = GetSystemSetting("output", "file filter excel", "Excel Workbook (*.xls)|*.xls")
'        gsOfficeTemplateFilter_Excel = GetSystemSetting("output", "template filter excel", "Excel Template (*.xls;*.xlt)|*.xls;*.xlt")
'      End If
  
      'Activate The System
      DebugOutput "modHRPro.Main", "Application.Activate"
      Application.Activate
      
      'Unload splash screen
      DebugOutput "modHRPro.Main", "Unload frmSplash"
      Unload frmSplash
      
    End If
  End If
    
  DebugOutput "modHRPro.Main", "End"

End Sub


Public Function ConvertData(ByVal pvarData As Variant, ByVal pDatatype As SQLDataType) As Variant
'Public Function ConvertData(ByVal pvarData As Variant, ByVal pDataType As ADODB.DataTypeEnum) As Variant
  ' Convert the given variant value into the given data type.
  On Error GoTo ErrorTrap
  
  Dim varReturnData As Variant
  
  If IsNull(pvarData) Then
    varReturnData = Null
  Else
    Select Case pDatatype
      Case sqlBoolean
        varReturnData = CBool(pvarData)
            
      Case sqlVarChar, sqlLongVarChar
        varReturnData = RTrim(pvarData)
        If Len(varReturnData) = 0 Then
          varReturnData = Null
        End If
      
      Case sqlDate
        If IsDate(pvarData) Then
          varReturnData = CDate(pvarData)
        Else
          varReturnData = Null
        End If
              
'      Case adDecimal
'        If VarType(pvarData) = vbString And Len(Trim(pvarData)) = 0 Then
'          varReturnData = Null
'        Else
'          varReturnData = CDec(pvarData)
'        End If
              
      Case sqlNumeric
        If VarType(pvarData) = vbString And Len(Trim(pvarData)) = 0 Then
          varReturnData = Null
        Else
          varReturnData = CDbl(pvarData)
        End If
            
      Case sqlInteger
        If VarType(pvarData) = vbString And Len(Trim(pvarData)) = 0 Then
          varReturnData = Null
        Else
          varReturnData = CLng(pvarData)
        End If
              
'      Case adTinyInt
'        If VarType(pvarData) = vbString And Len(Trim(pvarData)) = 0 Then
'          varReturnData = Null
'        Else
'          varReturnData = CInt(pvarData)
'        End If
            
      Case Else
        varReturnData = CVar(pvarData)
          
    End Select
  End If
  
  ConvertData = varReturnData
  
  Exit Function
  
ErrorTrap:
  If (Err.Number = 6) And (pDatatype = sqlInteger) Then
    ConvertData = 0
  Else
    ConvertData = CVar(vbNullString)
  End If
  Err = False
  
End Function


Public Function DateFormat() As String
  ' Returns the date format.
  ' NB. Windows allows the user to configure totally stupid
  ' date formats (eg. d/M/yyMydy !). This function does not cater
  ' for such stupidity, and simply takes the first occurence of the
  ' 'd', 'M', 'y' characters.
  Dim sSysFormat As String
  Dim sSysDateSeparator As String
  Dim sDateFormat As String
  Dim iLoop As Integer
  Dim fDaysDone As Boolean
  Dim fMonthsDone As Boolean
  Dim fYearsDone As Boolean
  
  fDaysDone = False
  fMonthsDone = False
  fYearsDone = False
  sDateFormat = ""
    
  sSysFormat = UI.GetSystemDateFormat
  sSysDateSeparator = UI.GetSystemDateSeparator
    
  ' Loop through the string picking out the required characters.
  For iLoop = 1 To Len(sSysFormat)
      
    Select Case Mid(sSysFormat, iLoop, 1)
      Case "d"
        If Not fDaysDone Then
          ' Ensure we have two day characters.
          sDateFormat = sDateFormat & "dd"
          fDaysDone = True
        End If
          
      Case "M"
        If Not fMonthsDone Then
          ' Ensure we have two month characters.
          sDateFormat = sDateFormat & "mm"
          fMonthsDone = True
        End If
          
      Case "y"
        If Not fYearsDone Then
          ' Ensure we have four year characters.
          sDateFormat = sDateFormat & "yyyy"
          fYearsDone = True
        End If
          
      Case Else
        sDateFormat = sDateFormat & Mid(sSysFormat, iLoop, 1)
    End Select
      
  Next iLoop
    
  ' Ensure that all day, month and year parts of the date
  ' are present in the format.
  If Not fDaysDone Then
    If Mid(sDateFormat, Len(sDateFormat), 1) <> sSysDateSeparator Then
      sDateFormat = sDateFormat & sSysDateSeparator
    End If
      
    sDateFormat = sDateFormat & "dd"
  End If
    
  If Not fMonthsDone Then
    If Mid(sDateFormat, Len(sDateFormat), 1) <> sSysDateSeparator Then
      sDateFormat = sDateFormat & sSysDateSeparator
    End If
      
    sDateFormat = sDateFormat & "mm"
  End If
    
  If Not fYearsDone Then
    If Mid(sDateFormat, Len(sDateFormat), 1) <> sSysDateSeparator Then
      sDateFormat = sDateFormat & sSysDateSeparator
    End If
      
    sDateFormat = sDateFormat & "yyyy"
  End If
    
  ' Return the date format.
  DateFormat = sDateFormat
  
End Function


Public Function DBContains_DataType(pDatatype As SQLDataType) As Boolean

  '******************************************************************************
  ' DBContains_DataType - Returns a boolean indicating if the current DB has    *
  '                       any columns of the datatype indicated by pDataType.   *
  '                                                                             *
  ' Reason              - Created so the app only shows the Photo and OLE       *
  '                       path configuration screens, on login if nessecary.    *
  '******************************************************************************
  
  Dim rsDataType As ADODB.Recordset
  Dim sSQL As String
  
  On Error GoTo ErrorTrap
  
  sSQL = "SELECT C.* FROM ASRSysColumns C WHERE C.DataType = " & pDatatype

  Set rsDataType = datGeneral.GetRecords(sSQL)
  
  With rsDataType
    DBContains_DataType = Not (.EOF And .BOF)
    .Close
  End With
  
TidyUpAndExit:
  sSQL = vbNullString
  Set rsDataType = Nothing
  
  Exit Function
  
ErrorTrap:
  COAMsgBox "Error checking if the database " & gsDatabaseName & " contains columns of the specified datatype." _
          , vbExclamation + vbOKOnly, app.title
  DBContains_DataType = False
  Resume TidyUpAndExit
  
End Function

Public Function ValidateDate(psDate As String) As Variant
  ' Validate the given date with the system date format.
  ' Return vbNullString if the given date is invalid
  ' Else return the date value of the given date string.
  Dim sDay As String
  Dim sMonth As String
  Dim sYear As String
  Dim iYear As Integer
  Dim iDaysPerMonth As Integer
  Dim sDateFormat As String
  
  ' Get the system date format.
  sDateFormat = DateFormat
  
  ' Extract the year, month and day parts of the date.
  sYear = Replace(Mid(psDate, InStr(1, sDateFormat, "yyyy"), 4), " ", "")
  sMonth = Replace(Mid(psDate, InStr(1, sDateFormat, "mm"), 2), " ", "")
  sDay = Replace(Mid(psDate, InStr(1, sDateFormat, "dd"), 2), " ", "")

  ' Do not allow any part of the date to be empty.
  If Trim(sYear) = "" Or _
    Trim(sMonth) = "" Or _
    Trim(sDay) = "" Then
    
    ValidateDate = vbNullString
    Exit Function
    
  End If
  
  ' If the user entered less than two digits or less then assume
  ' the current century.
  If Len(Trim(sYear)) <= 2 Then
    ' Get the current century.
    sYear = Str((100 * Int(Year(Now) / 100)) + Val(sYear))
  End If
  
  ' Pad the year with zeroes to make it up to four digits long.
  sYear = Format(sYear, "0000")
  
  ' Validate the month part of the date.
  ' The date is invalid if the month part is empty while the other parts aren't.
  If Val(sMonth) > 12 Or _
    Val(sMonth) < 1 Then
    
    ValidateDate = vbNullString
    Exit Function
  End If
  
  ' Pad the month with zeroes to make it up to two digits long.
  sMonth = Format(sMonth, "00")
  
  ' Validate the day part of the date.
  ' Ensure we have a valid day value for the given month value.
  Select Case Val(sMonth)
    Case 4, 6, 9, 11 ' April, June, September, November (30 days in each)
      iDaysPerMonth = 30
      
    Case 1, 3, 5, 7, 8, 10, 12 'January, March, May, July, August, October, December (31 days in each)
      iDaysPerMonth = 31
    
    Case 2 ' February (28 days usually, 29 days in a leap year)
      iDaysPerMonth = 28
      iYear = Val(sYear)
      ' If the year is evenly divisible by 4 and not by 100
      ' then it is a leap year.
      If (iYear Mod 4 = 0) And _
        (iYear Mod 100 <> 0) Then
        iDaysPerMonth = 29
      Else
        ' If the year is evenly divisible by 4 and 100, then check to
        ' see if the quotient of year divided by 100 is also evenly
        ' divisible by 4. If it is, then this is a leap year.
        If (iYear Mod 4 = 0) And _
          (iYear Mod 100 = 0) And _
          (Int(iYear / 100) Mod 4 = 0) Then
          iDaysPerMonth = 29
        End If
      End If
      
  End Select

  If Val(sDay) < 1 Or Val(sDay) > iDaysPerMonth Then
    ValidateDate = vbNullString
    Exit Function
  End If
    
  ' Pad the day with zeroes to make it up to two digits long.
  sDay = Format(sDay, "00")
  
  ' Return the formatted date string.
  If InStr(1, sDateFormat, "d") < InStr(1, sDateFormat, "M") Then
    ValidateDate = DateValue(sDay & "/" & sMonth & "/" & sYear)
  Else
    ValidateDate = DateValue(sMonth & "/" & sDay & "/" & sYear)
  End If
  
End Function

Public Function GetTmpFNameInFolder(psPath As String) As String
  ' Same as the GetTmpFName function but works with a given folder,
  ' rather than the default temp folder.
  Dim lngCounter As Long
  Dim sTemp As String
  
  ' Create the given folder if it does not already exist.
  If Dir(psPath, vbDirectory) = vbNullString Then
    MkDir psPath
  End If
  
  ' Can't use the GetTempFileName API function as we're not
  ' using the default temp folder.
  ' Simply check if temp files exist, incrementing a counter suffix
  ' if the current attempt fails.
  lngCounter = 0
  sTemp = psPath & "\_T" & CStr(lngCounter) & ".tmp"
  Do While Dir(sTemp) <> vbNullString
    lngCounter = lngCounter + 1
    sTemp = psPath & "\_T" & CStr(lngCounter) & ".tmp"
  Loop
  
  GetTmpFNameInFolder = sTemp
  
End Function

Public Function GetTmpFName() As String

  Dim strTmpPath As String, strTmpName As String
  
  strTmpPath = Space(1024)
  strTmpName = Space(1024)

  Call GetTempPath(1024, strTmpPath)
  Call GetTempFileName(strTmpPath, "_T", 0, strTmpName)
  
  strTmpName = Trim(strTmpName)
  If Len(strTmpName) > 0 Then
    strTmpName = Left(strTmpName, Len(strTmpName) - 1)
    
    'MH20021227 For some reason a zero byte file is created... ANNOYING!
    If Dir(strTmpName) <> vbNullString Then
      Kill strTmpName
    End If
  
  Else
    strTmpName = vbNullString
  End If

  GetTmpFName = Trim(strTmpName)
  
End Function

Public Function GetSpecialFolderA(ByVal eSpecialFolder As mceIDLPaths) As String
  ' NPG20090305 Fault 13531
  ' Function returns the 'Temporary Internet Files' folder path if &H20 is passed in as eSpecialFolder.
  Dim iRet As Long
  Dim strPath As String
  
  strPath = Space(260)
    iRet = SHGetSpecialFolderPath(0, strPath, eSpecialFolder, False)
    If Trim(strPath) <> Chr(0) Then
        strPath = Left(strPath, InStr(strPath, Chr(0)) - 1) & "\"
        ' COAMsgBox (strPath)
    End If
    
    GetSpecialFolderA = strPath
    
End Function

Public Function GetTmpInternetFName() As String

  Dim strTmpPath As String, strTmpName As String
  
  strTmpPath = Space(1024)
  strTmpName = Space(1024)

  ' Call GetTempPath(1024, strTmpPath)
  strTmpPath = GetSpecialFolderA(CSIDL_INTERNET_CACHE)
  Call GetTempFileName(strTmpPath, "_T", 0, strTmpName)
  
  strTmpName = Trim(strTmpName)
  If Len(strTmpName) > 0 Then
    strTmpName = Left(strTmpName, Len(strTmpName) - 1)
    
    'MH20021227 For some reason a zero byte file is created... ANNOYING!
    If Dir(strTmpName) <> vbNullString Then
      Kill strTmpName
    End If
  
  Else
    strTmpName = vbNullString
  End If

  GetTmpInternetFName = Trim(strTmpName)
  
End Function

Public Function GetScreens() As Boolean
  ' Gets the list of screens the user can see, and populates the menu with them.
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iFileMenuCount As Integer
  Dim sBand As String
  Dim rsTemp As ADODB.Recordset
  Dim rsScreens As ADODB.Recordset
  Dim objFileTool As ActiveBarLibraryCtl.Tool
  Dim avPrimaryMenuInfo As Variant
  Dim avSubMenuInfo As Variant
  Dim sLastToolName As String
  Dim lngLastScreenID As Long
  Dim fViewLookupTableMenu As Boolean
  Dim objBandTool As ActiveBarLibraryCtl.Tool
  Dim sSQL As String
  Dim rsPictures As ADODB.Recordset
  Dim lngLastID As Long
  
  'Populate the menu with the screens available to the current user.
  With frmMain.abMain
    'JPD 20081202 Fault 13443
    ' Read all screen icons into a dummy menu band to improve performance later on (in LoadMenuPicture).
    sSQL = "SELECT ASRSysScreens.pictureID, ASRSysPictures.picture" & _
      " FROM ASRSysScreens" & _
      " INNER JOIN ASRSysPictures ON ASRSysScreens.pictureID = ASRSysPictures.pictureID" & _
      " WHERE ASRSysScreens.pictureID > 0" & _
      " ORDER BY ASRSysScreens.pictureID"
      
    Set rsPictures = gobjDataAccess.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    lngLastID = 0
    Do While Not rsPictures.EOF
      If lngLastID <> rsPictures!PictureID Then
        Set objBandTool = .Bands("bndIcons").Tools.Add(.Bands("bndIcons").Tools.Count + 1, "P" & CStr(rsPictures!PictureID))
        ReadPicture objBandTool, rsPictures!Picture, rsPictures!Picture.ActualSize, True

        lngLastID = rsPictures!PictureID
      End If

      rsPictures.MoveNext
    Loop
    rsPictures.Close
    Set rsPictures = Nothing
    
    'Get number of existing tools on the menu bar and add 1 for each new tool
    iFileMenuCount = .Tools.Count + 1

    'Get the parent table details.
    avPrimaryMenuInfo = datGeneral.GetPrimaryTableMenu

    For iLoop = 1 To UBound(avPrimaryMenuInfo, 2)
      GetColumnPrivileges (CStr(avPrimaryMenuInfo(2, iLoop)))
      
      If avPrimaryMenuInfo(4, iLoop) > 0 Then
        ' The user has 'read' permission on the table, and no views on the table.
        ' There is only one screen defined for the table.
        
        ' Add a menu option to call up the primary table screen.
        Set objFileTool = .Tools.Add(iFileMenuCount, "PT" & avPrimaryMenuInfo(4, iLoop))
        iFileMenuCount = iFileMenuCount + 1
        objFileTool.Caption = RemoveUnderScores(CStr(avPrimaryMenuInfo(2, iLoop))) & "..."
        objFileTool.Style = DDSStandard
            
        'Load the specified picture, or the default if none specified
        If avPrimaryMenuInfo(11, iLoop) > 0 Then
          LoadMenuPicture CLng(avPrimaryMenuInfo(11, iLoop)), objFileTool
        Else
          ' objFileTool.SetPicture 0, LoadResPicture("SCREEN", 0), COL_GREY
          objFileTool.SetPicture 0, LoadResPicture("SCREENICO", 1), COL_GREY
        End If
        
        'Add the new tool to the main menu
        .Bands("mnuFile").Tools.Insert 0, objFileTool
      
      ElseIf avPrimaryMenuInfo(7, iLoop) > 0 Then
        ' The user does NOT have 'read' permission on the table, but does have
        ' 'read' permission on one view of the table.
        ' There is only one screen defined for the view.
        
        ' Add a menu option to call up the primary table's view screen.
        Set objFileTool = .Tools.Add(iFileMenuCount, "PV" & avPrimaryMenuInfo(10, iLoop) & ":" & avPrimaryMenuInfo(7, iLoop))
        iFileMenuCount = iFileMenuCount + 1
        objFileTool.Caption = RemoveUnderScores(CStr(avPrimaryMenuInfo(2, iLoop))) & _
        " (" & RemoveUnderScores(CStr(avPrimaryMenuInfo(8, iLoop))) & " view)..."
        objFileTool.Style = DDSStandard
        If avPrimaryMenuInfo(12, iLoop) > 0 Then
          LoadMenuPicture CLng(avPrimaryMenuInfo(12, iLoop)), objFileTool
        Else
          objFileTool.SetPicture 0, LoadResPicture("VIEW", 0), COL_GREY
        End If
        .Bands("mnuFile").Tools.Insert 0, objFileTool
      
      ElseIf (avPrimaryMenuInfo(9, iLoop) > 0) Or _
        ((avPrimaryMenuInfo(5, iLoop) = True) And (avPrimaryMenuInfo(3, iLoop) > 0)) Then
        ' The user has 'read' permission on the table, and the table has more than one screen defined for it.
        ' Or there are views on the table.
  
        sBand = avPrimaryMenuInfo(2, iLoop)
        .Bands.Add sBand
        .Bands(sBand).Type = DDBTPopup
        
        'Instantiate the submenu heading tool and set properties
        Set objFileTool = .Tools.Add(iFileMenuCount, avPrimaryMenuInfo(1, iLoop))
        iFileMenuCount = iFileMenuCount + 1
        objFileTool.Caption = RemoveUnderScores(CStr(avPrimaryMenuInfo(2, iLoop)))
        objFileTool.SubBand = sBand
        objFileTool.SetPicture 0, LoadResPicture("TABLE", 0), COL_GREY
        
        'Add it to the main menu
        .Bands("mnuFile").Tools.Insert 0, objFileTool
      
        ' Add the submenu.
        avSubMenuInfo = datGeneral.GetPrimaryTableSubMenu(CLng(avPrimaryMenuInfo(1, iLoop)), CStr(avPrimaryMenuInfo(2, iLoop)))
        lngLastScreenID = 0
        sLastToolName = ""
        
        For iLoop2 = 1 To UBound(avSubMenuInfo, 2)
          If avSubMenuInfo(3, iLoop2) > 0 Then
            
            Set objFileTool = .Tools.Add(iFileMenuCount, "PV" & avSubMenuInfo(1, iLoop2) & ":" & avSubMenuInfo(3, iLoop2))
            iFileMenuCount = iFileMenuCount + 1
            objFileTool.Caption = RemoveUnderScores(CStr(avSubMenuInfo(2, iLoop2))) & _
              " (" & RemoveUnderScores(CStr(avSubMenuInfo(4, iLoop2))) & " view)..."
            objFileTool.Style = DDSStandard
            If avSubMenuInfo(5, iLoop2) > 0 Then
              LoadMenuPicture CLng(avSubMenuInfo(5, iLoop2)), objFileTool
            Else
              objFileTool.SetPicture 0, LoadResPicture("VIEW", 0), COL_GREY
            End If
            .Bands(sBand).Tools.Insert 0, objFileTool
          Else
            Set objFileTool = .Tools.Add(iFileMenuCount, "PT" & avSubMenuInfo(1, iLoop2))
            iFileMenuCount = iFileMenuCount + 1
            objFileTool.Caption = RemoveUnderScores(CStr(avSubMenuInfo(2, iLoop2))) & "..."
            objFileTool.Style = DDSStandard
            If avSubMenuInfo(5, iLoop2) > 0 Then
              LoadMenuPicture CLng(avSubMenuInfo(5, iLoop2)), objFileTool
            Else
              ' objFileTool.SetPicture 0, LoadResPicture("SCREEN", 0), COL_GREY
              objFileTool.SetPicture 0, LoadResPicture("SCREENICO", 1), COL_GREY
            End If
            .Bands(sBand).Tools.Insert 0, objFileTool
          End If
      
          If (lngLastScreenID > 0) And _
            (lngLastScreenID <> avSubMenuInfo(1, iLoop2)) Then
            .Bands(avPrimaryMenuInfo(2, iLoop)).Tools(sLastToolName).BeginGroup = True
          End If
            
          lngLastScreenID = avSubMenuInfo(1, iLoop2)
          sLastToolName = objFileTool.Name
        
        Next iLoop2
      End If
    Next iLoop

    ' Table screens sub-menu
    ' JPD 20030409 Fault 4093 - Lookup tables menu now a System Permission
    ' granted/denied in SecMgr.
    fViewLookupTableMenu = datGeneral.SystemPermission("MENU", "VIEWLOOKUPTABLES")
    
    If Not fViewLookupTableMenu Then
      .Tools("TableScreens").Visible = False
    Else
      Set rsTemp = datGeneral.GetTableScreens
      iFileMenuCount = iFileMenuCount + 1
      Do While Not rsTemp.EOF
        'First see if we have privileges to see this table
        If gcoTablePrivileges.item(rsTemp!TableName).AllowSelect _
          And Not gcoTablePrivileges.item(rsTemp!TableName).HideFromMenu Then
          Set objFileTool = .Tools.Add(iFileMenuCount, "TS" & rsTemp!TableID)
          iFileMenuCount = iFileMenuCount + 1
          objFileTool.Caption = RemoveUnderScores(rsTemp!TableName) & "..."
            If rsTemp!PictureID > 0 Then
              LoadMenuPicture rsTemp!PictureID, objFileTool
            Else
              If gcoTablePrivileges.item(rsTemp!TableName).TableType = tabLookup Then
                objFileTool.SetPicture 0, LoadResPicture("LOOKUP_TABLE", 0), COL_GREY
              Else
                ' objFileTool.SetPicture 0, LoadResPicture("SCREEN", 0), COL_GREY
                objFileTool.SetPicture 0, LoadResPicture("SCREENICO", 1), COL_GREY
              End If
            End If
          
          .Bands("bndTableScreens").Tools.Insert 0, objFileTool
        End If
        rsTemp.MoveNext
      Loop
      .Tools("TableScreens").Enabled = (.Bands("bndTableScreens").Tools.Count > 0)
      Set rsTemp = Nothing
    End If
    
    'Check if we have any Quick Entry screens to add to the menu.
    iFileMenuCount = .Tools.Count + 1
    
    Set rsScreens = datGeneral.GetQuickEntryScreens
    Do While Not rsScreens.EOF
      'First see if we have privileges to see this table
      If gcoTablePrivileges.item(rsScreens!TableName).AllowSelect Then

        ' Check that the current user has 'select' permission on at least one parent table,
        ' or at least one view of one parent table referenced by the quick entry screen.
        If ViewQuickEntry(rsScreens!ScreenID) Then
          Set objFileTool = .Tools.Add(iFileMenuCount, "QE" & rsScreens!ScreenID)
          iFileMenuCount = iFileMenuCount + 1
          objFileTool.Caption = RemoveUnderScores(rsScreens!Name) & "..."

          If rsScreens!PictureID > 0 Then
            LoadMenuPicture rsScreens!PictureID, objFileTool
          Else
            ' objFileTool.SetPicture 0, LoadResPicture("SCREEN", 0), COL_GREY
            objFileTool.SetPicture 0, LoadResPicture("SCREENICO", 1), COL_GREY
          End If

          .Bands("bndQuickEntry").Tools.Insert 0, objFileTool
        End If
      End If
      rsScreens.MoveNext
    Loop
    .Tools("QuickEntry").Enabled = (.Bands("bndQuickEntry").Tools.Count > 0)
  End With

End Function


Public Function GetColumnPrivileges(psTableViewName As String) As CColumnPrivileges
  ' Return the column privileges collection for the given table.
  On Error GoTo ErrorTrap
    
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim objColumnPrivileges As CColumnPrivileges
  Dim sTableViewName As String
  
  fOK = True
  sTableViewName = UCase$(psTableViewName)
  
  ' Instantiate  the Column Privileges collection if it does not already exist.
  If gcolColumnPrivilegesCollection Is Nothing Then
    Set gcolColumnPrivilegesCollection = New Collection
  End If
  
  ' If the given table/view's column privilege collection has already been
  ' read then simply return it.
  For iLoop = 1 To gcolColumnPrivilegesCollection.Count
    If UCase$(gcolColumnPrivilegesCollection(iLoop).Tag) = sTableViewName Then
      Set GetColumnPrivileges = gcolColumnPrivilegesCollection(iLoop)
      Exit Function
    End If
  Next iLoop
  
' JPD20020814 New child views - no longer required
'''  ' The given table/views column privileges have not been read, so read them now,
'''  ' and add the definition to the collection to speed up subsequent calls.
'''  ' Instantiate a new collection of column privileges.
'''  Set objColumnPrivileges = New CColumnPrivileges
'''  objColumnPrivileges.Tag = psTableViewName
'''
'''  datGeneral.GetColumnPermissions objColumnPrivileges
'''
'''  ' Add the column privileges collection to the collection of column privileges
'''  ' collection. Confused ?
'''  gcolColumnPrivilegesCollection.Add objColumnPrivileges, psTableViewName
    
TidyUpAndExit:
  If fOK Then
    Set GetColumnPrivileges = objColumnPrivileges
  Else
    Set GetColumnPrivileges = Nothing
  End If
  Exit Function
    
ErrorTrap:
    COAMsgBox Err.Description & " - GetColumnPrivileges"
    fOK = False
    Resume TidyUpAndExit
    
End Function


Public Function GetHistoryScreens(plngScreenID As Long) As clsHistoryScreens
  ' Return the history screens collection for the given screen.
  On Error GoTo ErrorTrap
    
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim rsScreens As ADODB.Recordset
  Dim objHistoryScreens As clsHistoryScreens
  Dim sSQL As String

  fOK = True

  ' Instantiate  the Column Privileges collection if it does not already exist.
  If gcolHistoryScreensCollection Is Nothing Then
    Set gcolHistoryScreensCollection = New Collection
  End If

  ' If the given screen's history screen collection has already been
  ' read then simply return it.
  For iLoop = 1 To gcolHistoryScreensCollection.Count
    If gcolHistoryScreensCollection(iLoop).Tag = Trim(Str(plngScreenID)) Then
      Set GetHistoryScreens = gcolHistoryScreensCollection(iLoop)
      Exit Function
    End If
  Next iLoop

  ' Instantiate a new collection of history screen.
  Set objHistoryScreens = New clsHistoryScreens
  objHistoryScreens.Tag = Trim(Str(plngScreenID))

  sSQL = "exec dbo.sp_ASRGetHistoryScreens " & Trim(Str(plngScreenID))
  Set rsScreens = gobjDataAccess.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  
  Do While Not rsScreens.EOF
    ' Check the screen is for a readable table.
    If gcoTablePrivileges.item(rsScreens!RealSource).AllowSelect Then
      objHistoryScreens.Add rsScreens!ScreenID, rsScreens!Name, rsScreens!PictureID, _
        rsScreens!TableID, 0, "", rsScreens!TableName
      End If
    rsScreens.MoveNext
  Loop
  rsScreens.Close
  Set rsScreens = Nothing

  ' Add the history screens collection to the collection of history screens
  ' collections. Confused ?
  gcolHistoryScreensCollection.Add objHistoryScreens

TidyUpAndExit:
  If fOK Then
    Set GetHistoryScreens = objHistoryScreens
  Else
    Set GetHistoryScreens = Nothing
  End If
  Exit Function
    
ErrorTrap:
    COAMsgBox Err.Description & " - GetHistoryScreens"
    fOK = False
    Resume TidyUpAndExit
    
End Function



Public Function IsTypeNumeric(ByVal DataType As DataTypeEnum) As Boolean

  Select Case DataType
    Case adBigInt, adInteger, adSmallInt, adTinyInt
      IsTypeNumeric = True
    Case adNumeric
      IsTypeNumeric = True
    Case adDecimal
      IsTypeNumeric = True
    Case adDouble
      IsTypeNumeric = True
    Case Else
      IsTypeNumeric = False
  End Select
  
End Function

Public Function EditForm_Load(ByVal ScreenID As Long, ByVal vScreenType As ScreenType, _
  Optional ByVal lViewID As Long = 0, Optional bNewTableEntry As Boolean) As Boolean
  ' Open a new record edit form for the given screen and table/view.
  Dim frmEditForm As frmRecEdit4
  Dim frmForm As Form
  
  ' Ensure that the current recEdit screen (if there is one) has changes saved.
  If Not vScreenType = screenLookup Then
    If Not frmMain.ActiveForm Is Nothing Then
      If TypeOf frmMain.ActiveForm Is frmRecEdit4 Then
        If Not frmMain.ActiveForm.SaveChanges Then
          Exit Function
        End If
      End If
    End If
  End If

  ' Show progress bar
  With gobjProgress
    '.AviFile = App.Path & "\videos\screen.avi"
    .AVI = dbScreenAutoLayout
    .Caption = "Loading screen..."
    .NumberOfBars = 0
    .Time = False
    .Cancel = False
    .OpenProgress
  End With

  'TM20021002 Fault 4465 - using the Screen.ActiveForm object whilst it was Nothing created
  'this error. If it is Nothing then set focus to the Main form, this will in turn set focus to
  'an Active MDI Child, if one exists.
  'TM20021030 Fault 4674
'  If Screen.ActiveForm Is Nothing Then
    frmMain.SetFocus
'  End If
  
  ' First check if it's a history screen and if it's already there, only one per parent allowed
  If (vScreenType = screenHistoryTable) Or _
    (vScreenType = screenHistoryView) Then

    For Each frmForm In Forms
      If frmForm.Name = "frmRecEdit4" Then
        If frmForm.ParentFormID = Screen.ActiveForm.FormID Then
          If frmForm.ScreenID = ScreenID Then
            If frmForm.ViewID = lViewID Then
              If frmForm.Visible Then
                frmForm.Enabled = True   'MH20001023
                frmForm.SetFocus
              Else
                frmForm.ShowHistorySummary
              End If

              'AE20080102 Fault #12733
              gobjProgress.CloseProgress
              Exit Function
            End If
          End If
        End If
      End If
    Next
  End If
  
'  Dim bEditFormFailed As Boolean
  
  ' Instantiate a new edit form
  Set frmEditForm = New frmRecEdit4
  With frmEditForm
    .ScreenType = vScreenType
    .FormID = Application.NextFormID
  End With

  ' Now decide what type of record editing form we
  ' are trying to load in order to set it's parent.
  Select Case vScreenType
    ' Parent screen so set the parent form id to 0
    Case screenParentTable, screenParentView
      frmEditForm.ParentFormID = 0
      
    ' History screen so set the parent form id property
    ' and increment the number of children that the parent has.
    Case screenHistoryTable, screenHistoryView
      frmEditForm.ParentFormID = Screen.ActiveForm.FormID
      
    ' Lookup screen.
    Case screenLookup
      frmEditForm.ParentFormID = 0
      frmEditForm.LookupLoading = bNewTableEntry
  End Select
  
  ' Now try to load the actual screen display.
  If Not frmEditForm.LoadScreen(ScreenID, lViewID) Then
    ' If the form fails to load then decrement the number
    ' of children if it was a history screen
    Unload frmEditForm
    Set frmEditForm = Nothing
    EditForm_Load = False
  Else
    ' The edit form has loaded correctly.
    EditForm_Load = True
    With frmEditForm
      .Enabled = True
      .LookupLoading = False
      
      ' RH 01/03/01 - Allow user definable startup screens for each
      '               type of table, eg Find/RecEdit New/RecEdit First
      
'      If (vScreenType = screenHistoryTable) Or _
'        (vScreenType = screenHistoryView) Then
'
'        .HistoryInitialise
'      Else
'        DoEvents
'        .Visible = True
'        DoEvents
'        .SetFocus
'      End If
      
      Select Case vScreenType
      
        Case screenHistoryTable, screenHistoryView
      
          If gcHistory = disFindWindow Then
'************************************************************
' Have commented out as fix 2568 created faults 3750 & 3751
'            'TM20020327 Fault 2568
'            'history may not have initialised successfully eg. the user
'            'does not have permissions on the columns in the sort order.
'            bEditFormFailed = Not .HistoryInitialise
'************************************************************

            .HistoryInitialise
          Else
            If gcHistory = disRecEdit_New Then .AddNew

            ' JPD20030226 Fault 5079
            If (.Recordset.EditMode = adEditAdd) And (Not .AllowInsert) Then
              .HistoryInitialise
            Else
              .Visible = True
              DoEvents
              .SetFocus
            End If
          End If
      
        Case screenParentTable, screenParentView
        
          If gcPrimary = disFindWindow Then
            .Find
          Else
            If gcPrimary = disRecEdit_New Then .AddNew
            
            ' JPD20030226 Fault 5079
            If (.Recordset.EditMode = adEditAdd) And (Not .AllowInsert) Then
              .Find
            Else
              .Visible = True
              DoEvents
              .SetFocus
            End If
          End If
        
        Case screenLookup
        
          If gcLookUp = disFindWindow Then
          
            If bNewTableEntry = False Then
              .Find
            Else
              .AddNew
              
              ' JPD20030226 Fault 5079
              If (.Recordset.EditMode = adEditAdd) And (Not .AllowInsert) Then
                .Find
              Else
                .Visible = True
                DoEvents
                .SetFocus
              End If
            End If
          
          Else
            If gcLookUp = disRecEdit_New Then .AddNew
            
            ' JPD20030226 Fault 5079
            If (.Recordset.EditMode = adEditAdd) And (Not .AllowInsert) Then
              .Find
            Else
              .Visible = True
              DoEvents
              .SetFocus
            End If
          End If
          
        Case screenQuickEntry
        
          If gcQuickAccess = disFindWindow Then
            .Find
          Else
            If gcQuickAccess = disRecEdit_New Then .AddNew
            
            ' JPD20030226 Fault 5079
            If (.Recordset.EditMode = adEditAdd) And (Not .AllowInsert) Then
              .Find
            Else
              .Visible = True
              DoEvents
              .SetFocus
            End If
          End If
        
      End Select
    End With
  End If
  
  gobjProgress.CloseProgress
  
End Function



Public Function ReadOLEData(OLEControl As OLE, OLEField As ADODB.Field) As Boolean
On Error GoTo ErrorTrap

  Dim strFileName As String
  Dim intFileNo As Integer
  Dim lngColSize As Long
  Dim ChunkSize As Long
  Dim Chunks As Integer
  Dim Fragment As Integer
  Dim Chunk() As Byte
  Dim i As Integer

  ChunkSize = 2 ^ 14

  If Not (OLEControl Is Nothing Or OLEField Is Nothing) Then
    lngColSize = OLEField.ActualSize
    If lngColSize > 0 Then
      strFileName = GetTmpFName
      intFileNo = FreeFile(1)
      Open strFileName For Binary Access Write As intFileNo
        
      Chunks = lngColSize \ ChunkSize
      Fragment = lngColSize Mod ChunkSize
        
      ReDim Chunk(Fragment)
      Chunk() = OLEField.GetChunk(Fragment)
      Put intFileNo, , Chunk()
        
      For i = 1 To Chunks
        ReDim Chunk(ChunkSize)
        Chunk() = OLEField.GetChunk(ChunkSize)
        Put intFileNo, , Chunk()
      Next i
        
      Close intFileNo
      Open strFileName For Binary Access Read As intFileNo
      OLEControl.ReadFromFile intFileNo
      OLEControl.DataChanged = False
    
      Close intFileNo
      Kill strFileName
      
      ReadOLEData = True
    End If
  End If
    
  Exit Function
  
ErrorTrap:
  ReadOLEData = False
  
End Function

Public Function BackupOLEData(pctlOLEControl As OLE) As String
  ' Write the OLE data to a file, and return the filename.
  On Error GoTo ErrorTrap
  
  Dim strFileName As String
  Dim intFileNo As Integer
  
  strFileName = vbNullString
  
  If Not (pctlOLEControl Is Nothing) Then
    If pctlOLEControl.OLEType <> vbOLENone Then
      strFileName = GetTmpFName
      intFileNo = FreeFile(1)
      Open strFileName For Binary Access Read Write As intFileNo
      pctlOLEControl.SaveToFile intFileNo
      Close intFileNo
      intFileNo = 0
    End If
  End If
  
TidyUpAndExit:
  On Error Resume Next
  If intFileNo > 0 Then Close intFileNo
  BackupOLEData = strFileName
  Exit Function
  
ErrorTrap:
  Err = False
  Resume TidyUpAndExit
  
End Function


Public Sub RestoreOLEData(pctlOLEControl As OLE, psFilename As String)
  ' Read the OLE data from a file.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim intFileNo As Integer

  fOK = (Not pctlOLEControl Is Nothing) And (Not psFilename = vbNullString)
  If fOK Then
    intFileNo = FreeFile(1)
    Open psFilename For Binary Access Read As intFileNo
    pctlOLEControl.ReadFromFile intFileNo
    pctlOLEControl.DataChanged = False
    
    Close intFileNo
    Kill psFilename
  End If
  
TidyUpAndExit:
  If Not fOK Then
    pctlOLEControl.Delete
    pctlOLEControl.Class = vbNullString
  End If
  Exit Sub

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub



Public Function ReadPicture(ByRef PictureObject As Object, ByRef PictureField As Object, ByVal PictureSize As Long, _
            Optional bActiveBar As Boolean) As Boolean
            
  Dim strTempName As String
  Dim intFileNo As Integer
  Dim lngColSize As Long
  Dim ChunkSize As Long
  Dim Chunks As Integer
  Dim Fragment As Integer
  Dim Chunk() As Byte
  Dim i As Integer

  ChunkSize = 2 ^ 14

  If Not (PictureObject Is Nothing Or PictureField Is Nothing) Then
    lngColSize = PictureSize
    If lngColSize > 0 Then
      strTempName = GetTmpFName
      intFileNo = FreeFile(1)
      Open strTempName For Binary Access Write As intFileNo
      
      Chunks = lngColSize \ ChunkSize
      Fragment = lngColSize Mod ChunkSize
        
      ReDim Chunk(Fragment)
      Chunk() = PictureField.GetChunk(Fragment)
      Put intFileNo, , Chunk()
        
      For i = 1 To Chunks
        ReDim Chunk(ChunkSize)
        Chunk() = PictureField.GetChunk(ChunkSize)
        Put intFileNo, , Chunk()
      Next i
      Close intFileNo
        
      If bActiveBar Then
        PictureObject.SetPicture 0, LoadPicture(strTempName), COL_GREY
      Else
        If TypeOf PictureObject Is Form Then
'            Set PictureObject.Icon = LoadPicture(strTempName)
            
            PictureObject.picIcon = LoadPicture(strTempName, vbLPSmall, vbLPColor)
            SendMessageLong PictureObject.hWnd, WM_SETICON, ICON_SMALL, PictureObject.picIcon.Picture.Handle
            
        Else
            Set PictureObject.Picture = LoadPicture(strTempName)
        End If
      End If
      Kill strTempName
    End If
  End If
  
End Function

Public Function GetPictureFromDatabase(plngImageID As Long) As String
  On Error GoTo ErrorTrap
  
  Dim strTempName As String
  Dim intFileNo As Integer
  Dim lngColSize As Long
  Dim ChunkSize As Long
  Dim Fragment As Integer
  Dim i As Integer
  Dim Chunks As Integer
  Dim Chunk() As Byte
  'Dim TempFile As Integer
  Dim recPictures As Recordset

  ChunkSize = 2 ^ 14
  strTempName = ""

  Set recPictures = datGeneral.GetPicture(plngImageID)

  If Not recPictures Is Nothing Then
  
    With recPictures
    
      If .BOF And .EOF Then
        ' Background image somehow deleted.
        SaveSystemSetting "DesktopSetting", "BitmapID", 0
        COAMsgBox "The background image no longer exists.", vbExclamation + vbOKOnly, app.ProductName
      Else
        lngColSize = !Picture.ActualSize
        If lngColSize > 0 Then
          strTempName = GetTmpFName
          intFileNo = FreeFile(1)
          Open strTempName For Binary Access Write As intFileNo
        
          Chunks = lngColSize \ ChunkSize
          Fragment = lngColSize Mod ChunkSize
          
          ReDim Chunk(Fragment)
          Chunk() = !Picture.GetChunk(Fragment)
          Put intFileNo, , Chunk()
          
          For i = 1 To Chunks
            ReDim Chunk(ChunkSize)
            Chunk() = !Picture.GetChunk(ChunkSize)
            Put intFileNo, , Chunk()
          Next i
  
          Close intFileNo
  
        End If
      End If
    End With
  End If

  recPictures.Close
  Set recPictures = Nothing

  GetPictureFromDatabase = strTempName
  
  Exit Function
  
ErrorTrap:
  GetPictureFromDatabase = ""
  
End Function
Public Sub SetCaption(Control As Object, ByRef Caption As String)

'  On Error Resume Next
'
'  'Attempt to set the Caption property
'  If TypeOf Control Is ASRDummyLabel Or TypeOf Control Is Frame Or TypeOf Control Is CheckBox Then
'    Control.Caption = Caption
'  Else
'    Control.Text = Caption
'  End If

  On Error Resume Next

  'Attempt to set the Caption property
  Control.Caption = Caption
  If Err Then
    'Attempt to set the Text property
    Control.Text = Caption
  End If
  Err = False



End Sub

Public Function ValidExcelPageName(sValue As String, sReplaceWith As String) As String

  'TM20020627 Fault 4060
  
  'returns a valid name that can be used as the name of an xl worksheet.
  
  'PRE-CONDITION: sReplaceWith cannot be : \ / ? * [ ]
  
  Dim sTempPageName As String
    
  'replace obliques
  sTempPageName = Replace(Replace(sValue, "/", sReplaceWith), "\", sReplaceWith)
  'replace astrixes (how do you spell *?)
  sTempPageName = Replace(sTempPageName, "*", sReplaceWith)
  'replace colons
  sTempPageName = Replace(sTempPageName, ":", sReplaceWith)
  'replace square brackets
  sTempPageName = Replace(Replace(sTempPageName, "[", sReplaceWith), "]", sReplaceWith)
  'replace question marks
  sTempPageName = Replace(sTempPageName, "?", sReplaceWith)

  ValidExcelPageName = sTempPageName

End Function

Public Function WriteOLEData(OLEControl As OLE, OLEField As ADODB.Field) As Boolean

  On Error GoTo ErrorTrap
  
  Dim strFileName As String
  Dim intFileNo As Integer
  Dim lngFileSize As Long
  Dim ChunkSize As Long
  Dim Chunks As Integer
  Dim Fragment As Integer
  Dim Chunk() As Byte
  Dim i As Integer
  Dim blnWriteOk As Boolean
  
  ChunkSize = 2 ^ 14

  If Not (OLEControl Is Nothing Or OLEField Is Nothing) Then
    If OLEControl.OLEType <> vbOLENone Then
      strFileName = GetTmpFName
      intFileNo = FreeFile(1)
      Open strFileName For Binary Access Read Write As intFileNo
    
      OLEControl.SaveToFile intFileNo
      Seek intFileNo, 1
    
      lngFileSize = LOF(intFileNo)
      Chunks = lngFileSize \ ChunkSize
      Fragment = lngFileSize Mod ChunkSize
        
      ReDim Chunk(Fragment)
      Get intFileNo, 1, Chunk()
      OLEField.AppendChunk Chunk()
        
      For i = 1 To Chunks
        ReDim Chunk(ChunkSize)
        Get intFileNo, , Chunk()
        OLEField.AppendChunk Chunk()
      Next i
      Close intFileNo
      intFileNo = 0
      
      blnWriteOk = True
    End If
  End If
  
ExitWriteOLEData:
  On Error Resume Next
  
  If intFileNo > 0 Then Close intFileNo
  If Len(Trim(strFileName)) > 0 Then Kill strFileName
  
  WriteOLEData = blnWriteOk
  Exit Function
  
ErrorTrap:
  blnWriteOk = False
  Err = False
  
  Resume ExitWriteOLEData
  
End Function

Public Sub SetFlatEditMask(ByRef Object As Object, ByVal sInputMask As String)

  Dim sConverted As String
  Dim sMask As String
  Dim sLiteral As String

  sMask = Replace(sInputMask, "A", ">")
  sMask = Replace(sMask, "a", "<")
  sMask = Replace(sMask, "9", "0")
  
  sLiteral = Replace(sMask, ">", " ")
  sLiteral = Replace(sLiteral, "<", " ")
  sLiteral = Replace(sLiteral, "0", " ")
  
  ' Set the mask
  Object.SetMask sMask, sLiteral, " "

End Sub


Public Function ConvertMaskToNumeric(sMask As String) As String

    Dim lPos As Long
    
    lPos = InStr(1, sMask, "9")
    If lPos = 1 Then
        sMask = "#" & Mid$(sMask, 2, Len(sMask))
        lPos = InStr(1, sMask, "9")
    End If
    
    Do While lPos > 0
        sMask = Mid$(sMask, 1, lPos - 1) & "#" & Mid$(sMask, lPos + 1, Len(sMask))
        lPos = InStr(lPos + 1, sMask, "9")
    Loop
    ConvertMaskToNumeric = sMask

End Function

Public Function GetDisplayFormat(psMask As String) As String
  Dim lngPos As Long
    
  lngPos = InStr(1, psMask, ".")
  
  If lngPos > 0 Then
    psMask = "0" & Mid$(psMask, lngPos, Len(psMask))
    lngPos = InStr(lngPos, psMask, "#")
        
    Do While lngPos > 0
      psMask = Mid$(psMask, 1, lngPos - 1) & "0" & Mid$(psMask, lngPos + 1, Len(psMask))
      lngPos = InStr(lngPos + 1, psMask, "#")
    Loop
        
    GetDisplayFormat = psMask
  Else
    GetDisplayFormat = "0"
  End If

End Function

Public Function GetModuleParameter(psModuleKey As String, psParameterKey As String) As String
  ' Return the value of the given parameter.
  GetModuleParameter = datGeneral.GetModuleParameter(psModuleKey, psParameterKey)
  
End Function
Public Function GetModuleArray(psModuleKey As String, psParameterKey As String) As Recordset
  ' Return the value of the given parameter.
  Set GetModuleArray = datGeneral.GetModuleArray(psModuleKey, psParameterKey)
  
End Function

Public Sub InitialiseModules()
  
  '26/07/2001 MH
  '' Has the training booking module been authorised.
  'gfTrainingBookingEnabled = datGeneral.TrainingBookingEnabled
  gfTrainingBookingEnabled = IsModuleEnabled(modTraining)
  
  ' If so, then read the training booking parameters
  If gfTrainingBookingEnabled Then
    gfTrainingBookingEnabled = ReadTrainingBookingParameters
  End If
  
  '26/07/2001 MH
  '' Has the absence module been authorised.
  'gfAbsenceEnabled = datGeneral.AbsenceEnabled
  gfAbsenceEnabled = IsModuleEnabled(modAbsence)
  
  ' If so, then read the absence parameters
  If gfAbsenceEnabled Then
    ReadAbsenceParameters
  End If
  
  '26/07/2001 MH
  '' Has the personnel module been authorised.
  'gfPersonnelEnabled = datGeneral.PersonnelEnabled
  gfPersonnelEnabled = IsModuleEnabled(modPersonnel)

  ' If so, then read the personnel parameters
  If gfPersonnelEnabled Then
    ReadPersonnelParameters
    ReadPostParameters
  End If
  
  ' Has the AFD PC Names & Numbers module been authorised.
  gfAFDEnabled = datGeneral.AFDEnabled

  ' The the Quick Address module been authorised
  giQAddressEnabled = datGeneral.QAddressEnabled

  ' Check for any email links. If found, try and start the email service
  datGeneral.EmailGenerationEnabled

  ' Read the bank holiday parameters, then set flag. NB Do in this order!
  ReadBankHolidayParameters
  gfBankHolidaysEnabled = datGeneral.BankHolidaysEnabled

  ' Load the Desktop settings
  glngDesktopBitmapID = GetSystemSetting("DesktopSetting", "BitmapID", 0)
  glngDesktopBitmapLocation = GetSystemSetting("DesktopSetting", "BitmapLocation", 0)
  glngDeskTopColour = GetSystemSetting("DesktopSetting", "BackgroundColour", &H8000000C)

  '26/07/2001 MH
  '' Do we have access to the CMG export module
  'gbCMGEnabled = datGeneral.IsCMGEnabled
  gbCMGEnabled = IsModuleEnabled(modCMG)
  gbXMLExportEnabled = IsModuleEnabled(modXMLExport)

  ' Payroll Integration Module
  gbAccordEnabled = IsModuleEnabled(modAccord)

  ' Payroll Integration Module
  gbWorkflowEnabled = IsModuleEnabled(modWorkflow)
  gbWorkflowOutOfOfficeEnabled = OutOfOfficeEnabled

  gbVersion1Enabled = IsModuleEnabled(modVersionOne)

End Sub


Public Sub AddNewTableEntry(plngLookupTableID As Long)
  ' Display the screen for the given lookup table.
  'Dim sSQL As String
  Dim sTableName As String
  'Dim frmEdit As Form
  Dim rsTemp As Recordset
  Dim objTableView As CTablePrivilege
  
  Screen.MousePointer = vbHourglass
  
  ' Select the record from the screens table
  Set rsTemp = datGeneral.GetScreenScreens(plngLookupTableID)
  
  ' Get the lookup table name.
  For Each objTableView In gcoTablePrivileges.Collection
    If objTableView.TableID = plngLookupTableID Then
      sTableName = objTableView.TableName
      Exit For
    End If
  Next objTableView
  Set objTableView = Nothing
  
  ' Check that we only have one entry.
  If rsTemp.RecordCount > 0 Then
    rsTemp.MoveLast
    
    If rsTemp.RecordCount = 1 Then
      ' The lookup table has one screen defined for it. So display it.
      If Not EditForm_Load(rsTemp!ScreenID, screenLookup, , True) Then
        Set rsTemp = Nothing
        frmMain.ActiveForm.Enabled = True
        Screen.MousePointer = vbDefault
        Exit Sub
      End If
      Set rsTemp = Nothing
      
      'Set the edit form property to know that it's an Add New Table Entry edit form
      With frmMain.ActiveForm
        .NewTableEntry = True
        .AddNew
        .SetFocus
      End With
      
      Screen.MousePointer = vbDefault
      Exit Sub
    ElseIf rsTemp.RecordCount > 1 Then
      ' The lookup table has more than one screen defined for it. So tell the user.
      COAMsgBox "The lookup table '" & sTableName & "' has more than one screen definition." & vbCrLf & "Please contact your system administrator for more information.", vbExclamation, app.ProductName
      frmMain.ActiveForm.Enabled = True
    End If
  Else
    ' The lookup table has no screens defined for it. So tell the user.
    COAMsgBox "The lookup table '" & sTableName & "' does not have a screen definition." & vbCrLf & "Please contact your system administrator for more information.", vbExclamation, app.ProductName
    frmMain.ActiveForm.Enabled = True
  End If
  rsTemp.Close
  Set rsTemp = Nothing
  
  Screen.MousePointer = vbDefault

End Sub



Public Function GetParentFormParameter(plngParentFormID As Long, psKey As String) As String
  Dim frmForm As Form
    
  For Each frmForm In Forms
    If (frmForm.Name = "frmRecEdit4") Then
      If frmForm.FormID = plngParentFormID Then
        Select Case psKey
          Case "CAPTION":
            GetParentFormParameter = frmForm.Caption
          Case "STATUSCAPTION":
            GetParentFormParameter = frmForm.StatusCaption
          Case "PRINTHEADER":
            GetParentFormParameter = frmForm.FindPrintHeader
          Case Else:
            GetParentFormParameter = ""
        End Select
        
        Exit Function
      End If
    End If
  Next frmForm

  GetParentFormParameter = ""
  
End Function
Public Sub LoadMenuPicture(lPictureID As Long, objFileTool As Object)
  On Error GoTo ErrorTrap
  
  Dim rsPicture As Recordset
  Dim objBandTool As ActiveBarLibraryCtl.Tool

  objFileTool.SetPicture 0, frmMain.abMain.Bands("bndIcons").Tools("P" & CStr(lPictureID)).GetPicture(0), COL_GREY
   
  Exit Sub
  
ErrorTrap:
  If Err.Number = 2006 Then
    Set objBandTool = frmMain.abMain.Bands("bndIcons").Tools.Add(frmMain.abMain.Bands("bndIcons").Tools.Count + 1, "P" & CStr(lPictureID))
    
    Set rsPicture = datGeneral.GetPicture(lPictureID)
    ReadPicture objBandTool, rsPicture!Picture, rsPicture!Picture.ActualSize, True
    rsPicture.Close
    Set rsPicture = Nothing
  
    objFileTool.SetPicture 0, frmMain.abMain.Bands("bndIcons").Tools("P" & CStr(lPictureID)).GetPicture(0), COL_GREY
  End If
End Sub

Public Function RemoveUnderScores(ByRef sName As String) As String

    Dim lPos As Long
    Dim sTableName As String
    
    sTableName = sName
    lPos = InStr(1, sTableName, "_")
    If lPos = 0 Then
        RemoveUnderScores = sTableName
        Exit Function
    End If
    
    Do While lPos > 0
        sTableName = Mid$(sTableName, 1, lPos - 1) & " " & Mid$(sTableName, lPos + 1, Len(sTableName))
        lPos = InStr(lPos, sTableName, "_")
    Loop
    RemoveUnderScores = sTableName
    
End Function

Public Function AddUnderScores(ByRef sName As String) As String
'#RH 20/8/99

    Dim lPos As Long
    Dim sTableName As String
    
    sTableName = sName
    lPos = InStr(1, sTableName, " ")
    If lPos = 0 Then
        AddUnderScores = sTableName
        Exit Function
    End If
    
    Do While lPos > 0
        sTableName = Mid$(sTableName, 1, lPos - 1) & "_" & Mid$(sTableName, lPos + 1, Len(sTableName))
        lPos = InStr(lPos, sTableName, " ")
    Loop
    AddUnderScores = sTableName
    
End Function



Public Function XFrame() As Double
  ' Return the width of a control frame.
  XFrame = UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX

End Function

Public Sub SetupTablesCollection()
  ' Read the list of tables the current user has permission to see.
  Dim fSysSecManager As Boolean
  Dim iLoop As Long
  Dim lngRoleID As Long
  Dim lngChildViewID As Long
  Dim sSQL As String
  Dim sRealSourceList As String
  Dim sTableViewName As String
  Dim rsInfo As ADODB.Recordset
  Dim rsTables As ADODB.Recordset
  Dim rsViews As ADODB.Recordset
  Dim rsPermissions As ADODB.Recordset
  Dim objTableView As CTablePrivilege
  Dim objColumnPrivileges As CColumnPrivileges
  Dim avChildViews() As Variant
  Dim lngNextIndex As Long
  Dim iTemp As Integer
  Dim strSecurityGroupName As String
  Dim frmForm As Form
  Dim sPermissionName As String
  Dim sTableName As String
  Dim intAction As Integer
  
  DebugOutput "modHRPro.SetupTablesCollection", "Start"
    
    ' Instantiate a new collection of table privileges.
  Set gcoTablePrivileges = New CTablePrivileges
  
  ReDim avChildViews(3, 0)
    
  sSQL = "SELECT ASRSysChildViews2.childViewID, ASRSysTables.tableName, ASRSysChildViews2.type" & _
    " FROM ASRSysChildViews2" & _
    " INNER JOIN ASRSysTables ON ASRSysChildViews2.tableID = ASRSysTables.tableID" & _
    " WHERE ASRSysChildViews2.role = '" & Replace(gsUserGroup, "'", "''") & "'"
  Set rsInfo = New ADODB.Recordset
  rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  
  Do While Not rsInfo.EOF
    lngNextIndex = UBound(avChildViews, 2) + 1
    ReDim Preserve avChildViews(3, lngNextIndex)
    avChildViews(1, lngNextIndex) = rsInfo!childViewID
    avChildViews(2, lngNextIndex) = rsInfo!TableName
    avChildViews(3, lngNextIndex) = IIf(IsNull(rsInfo!Type), 0, rsInfo!Type)
  
    rsInfo.MoveNext
  Loop
  rsInfo.Close
  Set rsInfo = Nothing
  
  sSQL = "SELECT count(*) AS recCount" & _
    " FROM ASRSysGroupPermissions" & _
    " INNER JOIN ASRSysPermissionItems ON ASRSysGroupPermissions.itemID = ASRSysPermissionItems.itemID" & _
    " INNER JOIN ASRSysPermissionCategories ON ASRSysPermissionItems.categoryID = ASRSysPermissionCategories.categoryID" & _
    " INNER JOIN sysusers a ON ASRSysGroupPermissions.groupName = a.name" & _
    "   AND a.name = '" & gsUserGroup & "'" & _
    " WHERE (ASRSysPermissionItems.itemKey = 'SYSTEMMANAGER'" & _
    " OR ASRSysPermissionItems.itemKey = 'SECURITYMANAGER')" & _
    " AND ASRSysGroupPermissions.permitted = 1" & _
    " AND ASRSysPermissionCategories.categorykey = 'MODULEACCESS'"
  
  Set rsInfo = New ADODB.Recordset
  rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  fSysSecManager = (rsInfo!recCount > 0)
  rsInfo.Close
  Set rsInfo = Nothing

  
  
  sSQL = "SELECT Distinct ASRSysGroupPermissions.groupName" & _
    " FROM ASRSysGroupPermissions" & _
    " INNER JOIN sysusers a ON ASRSysGroupPermissions.groupName = a.name" & _
    "   AND a.name = '" & gsUserGroup & "'"
  
  Set rsInfo = New ADODB.Recordset
  rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
  If Not (rsInfo.EOF And rsInfo.BOF) Then
    rsInfo.MoveFirst
    strSecurityGroupName = rsInfo.Fields("GroupName").Value
  Else
    strSecurityGroupName = ""
  End If
  rsInfo.Close
  Set rsInfo = Nothing


  ' Initialise the collection with items for each TABLE in the system.
  Set rsTables = datGeneral.GetAllTables
  With rsTables
    Do While Not .EOF
      Set objTableView = gcoTablePrivileges.Add(!TableName, !TableID, !TableType, !DefaultOrderID, _
        !RecordDescExprID, True, 0, "", !IsRemoteView)

      objTableView.RealSource = !TableName

      .MoveNext
    Loop
    .Close
  End With
  Set rsTables = Nothing

  ' Initialise the collection with items for each VIEW in the system.
  Set rsViews = datGeneral.GetAllViews
  With rsViews
    Do While Not .EOF
      Set objTableView = gcoTablePrivileges.Add(!TableName, !TableID, !TableType, !DefaultOrderID, _
        !RecordDescExprID, False, !ViewID, !ViewName, False)

      objTableView.RealSource = !ViewName

      .MoveNext
    Loop
    .Close
  End With
  Set rsViews = Nothing


  ' Get the 'realSource' and permissions for each table or view.
  If fSysSecManager Then
    For Each objTableView In gcoTablePrivileges.Collection
      If objTableView.TableType = tabChild Then
        sSQL = "SELECT childViewID" & _
          " FROM ASRSysChildViews2" & _
          " WHERE tableID = " & Trim(Str(objTableView.TableID)) & _
          " AND role = '" & Replace(gsUserGroup, "'", "''") & "'"
        Set rsInfo = New ADODB.Recordset
        rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText

        If Not (rsInfo.BOF And rsInfo.EOF) Then
          objTableView.RealSource = Left("ASRSysCV" & Trim(Str(rsInfo!childViewID)) & "#" & Replace(objTableView.TableName, " ", "_") & "#" & Replace(gsUserGroup, " ", "_"), 255)
        End If
        
        rsInfo.Close
        Set rsInfo = Nothing

      Else
        objTableView.RealSource = IIf(objTableView.IsTable, objTableView.TableName, objTableView.ViewName)
      End If

      objTableView.AllowSelect = True
      objTableView.AllowUpdate = True
      objTableView.AllowDelete = True
      objTableView.AllowInsert = True
      objTableView.HideFromMenu = False
    Next objTableView
    Set objTableView = Nothing
  Else
    ' If the user is NOT a 'system manager' or 'security manager'
    ' read the table permissions from the server.
    sSQL = "exec dbo.sp_ASRAllTablePermissions '" & Replace(gsSQLUserName, "'", "''") & "'"
    Set rsPermissions = New ADODB.Recordset
    rsPermissions.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText

    Do While Not rsPermissions.EOF
      Set objTableView = Nothing

      sPermissionName = rsPermissions.Fields("Name").Value

      If UCase$(Left$(sPermissionName, 8)) = "ASRSYSCV" Then
        lngChildViewID = rsPermissions.Fields("BaseTableID").Value
        Set objTableView = gcoTablePrivileges.FindTableID(lngChildViewID)
      Else
        If gcoTablePrivileges.IsValid(sPermissionName) Then
          Set objTableView = gcoTablePrivileges.item(sPermissionName)
        End If
      End If

      If Not objTableView Is Nothing Then
        objTableView.RealSource = sPermissionName

        Select Case rsPermissions.Fields("Action").Value
          Case 193 ' Select permission.
            objTableView.AllowSelect = True
          Case 195 ' Insert permission.
            objTableView.AllowInsert = True
          Case 196 ' Delete permission.
            objTableView.AllowDelete = True
          Case 197 ' Update permission.
            objTableView.AllowUpdate = True
        End Select
      End If

      rsPermissions.MoveNext
      
    Loop
    rsPermissions.Close
    Set rsPermissions = Nothing
  
    ' Get the view menu permissions
    sSQL = "SELECT TableName, HideFromMenu FROM ASRSysViewMenuPermissions WHERE GroupName = '" & strSecurityGroupName & "'"
    Set rsPermissions = New Recordset
    rsPermissions.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText

    Do While Not rsPermissions.EOF
      'JPD 20040109 Fault 7624
      If gcoTablePrivileges.IsValid(rsPermissions!TableName) Then
        Set objTableView = gcoTablePrivileges.item(rsPermissions!TableName)
        objTableView.HideFromMenu = rsPermissions!HideFromMenu
      End If
      rsPermissions.MoveNext
    Loop
    rsPermissions.Close
    Set rsPermissions = Nothing
  
  End If

  ' Get the column permissions for each table/view.
  sRealSourceList = vbNullString
  Set gcolColumnPrivilegesCollection = New Collection
  
  For Each objTableView In gcoTablePrivileges.Collection
    If Len(objTableView.RealSource) > 0 Then
      sRealSourceList = sRealSourceList & IIf(Len(sRealSourceList) > 0, ", '", "'") & objTableView.RealSource & "'"
    End If
  
    ' Instantiate  the Column Privileges collection if it does not already exist.
    Set objColumnPrivileges = New CColumnPrivileges
    
    If objTableView.IsTable Then
      objColumnPrivileges.Tag = UCase$(objTableView.TableName)
    Else
      objColumnPrivileges.Tag = UCase$(objTableView.ViewName)
    End If
    gcolColumnPrivilegesCollection.Add objColumnPrivileges, objColumnPrivileges.Tag
  
  Next objTableView
  Set objTableView = Nothing


  If Len(sRealSourceList) > 0 Then
    
    ' Get the list of all columns in all tables/views.
    Set rsInfo = New ADODB.Recordset
    rsInfo.Open "dbo.spASRGetAllTableAndViewColumns", gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc
    Do While Not rsInfo.EOF

      sTableName = UCase(rsInfo.Fields("TableViewName").Value)
      Set objColumnPrivileges = gcolColumnPrivilegesCollection.item(sTableName)

      objColumnPrivileges.Add _
        fSysSecManager, _
        fSysSecManager, _
        IIf(IsNull(rsInfo!ColumnName), "", rsInfo!ColumnName), _
        IIf(IsNull(rsInfo!ColumnType), 0, rsInfo!ColumnType), _
        IIf(IsNull(rsInfo!DataType), 0, rsInfo!DataType), _
        IIf(IsNull(rsInfo!ColumnID), 0, rsInfo!ColumnID), _
        IIf(IsNull(rsInfo!UniqueCheckType), 0, rsInfo!UniqueCheckType), _
        IIf(IsNull(rsInfo!DefaultDisplayWidth), 0, rsInfo!DefaultDisplayWidth), _
        IIf(IsNull(rsInfo!Size), 0, rsInfo!Size), _
        IIf(IsNull(rsInfo!Decimals), 0, rsInfo!Decimals), _
        IIf(IsNull(rsInfo!Use1000Separator), False, rsInfo!Use1000Separator), _
        IIf(IsNull(rsInfo!OLEType), OLE_SERVER, rsInfo!OLEType)

      rsInfo.MoveNext

    Loop
    rsInfo.Close
    Set rsInfo = Nothing


    ' If the current user is not a system/security manager then read the column permissions from SQL.
    If Not fSysSecManager Then
      ' Get the SQL group id of the current user.
      sSQL = "SELECT gid" & _
        " FROM sysusers" & _
        " WHERE name = '" & Replace(gsUserGroup, "'", "''") & "'"
      Set rsInfo = New ADODB.Recordset
      rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
      lngRoleID = rsInfo!gid
      rsInfo.Close

      sSQL = "SELECT sysobjects.name AS tableViewName," & _
        " syscolumns.name AS columnName," & _
        " sysprotects.action," & _
        " CASE protectType" & _
        "   WHEN 205 THEN 1" & _
        "   WHEN 204 THEN 1" & _
        "   ELSE 0" & _
        " END AS permission" & _
        " FROM sysprotects" & _
        " INNER JOIN sysobjects ON sysprotects.id = sysobjects.id" & _
        " INNER JOIN syscolumns ON sysprotects.id = syscolumns.id" & _
        " WHERE sysprotects.uid = " & Trim(Str(lngRoleID)) & _
        " AND (sysprotects.action = 193 or sysprotects.action = 197)" & _
        " AND syscolumns.name <> 'timestamp' AND syscolumns.name <> '_UPDFLAG'" & _
        " AND sysobjects.name in (" & sRealSourceList & ")" & _
        " AND (((convert(tinyint,substring(sysprotects.columns,1,1))&1) = 0" & _
        " AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)" & _
        " OR ((convert(tinyint,substring(sysprotects.columns,1,1))&1) != 0" & _
        " AND (convert(int,substring(sysprotects.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0))"
      rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
      
      Do While Not rsInfo.EOF
        
        ' Get the current column's table/view name.
        sTableName = rsInfo.Fields(0).Value   ' TableViewName
        Set objTableView = gcoTablePrivileges.FindRealSource(sTableName)
        
        If objTableView.IsTable Then
          sTableViewName = objTableView.TableName
        Else
          sTableViewName = rsInfo!TableViewName
        End If

        Set objColumnPrivileges = gcolColumnPrivilegesCollection.item(UCase(sTableViewName))

        intAction = rsInfo.Fields("Action").Value

        If intAction = 193 Then
          objColumnPrivileges.item(rsInfo!ColumnName).AllowSelect = rsInfo!Permission
        ElseIf intAction = 197 Then
          objColumnPrivileges.item(rsInfo!ColumnName).AllowUpdate = rsInfo!Permission
        End If

        rsInfo.MoveNext
      Loop
      rsInfo.Close
      Set rsInfo = Nothing
    End If
   
  End If
 
  'JPD 20040625 Fault 8714
  For Each frmForm In Forms
    If (frmForm.Name = "frmMain") Then
      ' If the frmMain form is already loaded then we need to rerun its Load method.
      ' The frmMain_load method may not have been run properly as it may have been loaded
      ' before the gcoTablePrivileges collection was created. This happened if the path
      ' definition screens were displayed after logging out of DatMgr.
      frmMain.Reload
      Exit For
    End If
  Next frmForm
  Set frmForm = Nothing
    
  DebugOutput "modHRPro.SetupTablesCollection", "End"

End Sub

Private Function ViewQuickEntry(plngScreenID As Long) As Boolean
  ' Return TRUE if the current user can see at least one parent table (or view of a parent table)
  ' of given quick view screen.
  On Error GoTo ErrorTrap
  
  Dim fCanView As Boolean
  Dim rsTables As Recordset
  Dim rsViews As Recordset
    
  fCanView = False
    
  ' Get the list of parent tables used in the quick entry screen.
  Set rsTables = datGeneral.GetQuickEntryTables(plngScreenID)
    
  With rsTables
    If (.EOF And .BOF) Then
      fCanView = True
    End If
  
    ' Loop through parent tables, seeing if we have select permissions on these tables.
    Do While (Not .EOF) And (Not fCanView)
      ' Check if the current user has 'select' permission on the given table.
      If gcoTablePrivileges.item(!TableName).AllowSelect Then
        fCanView = True
      Else
        ' No select permissions, can we use a view instead ???
        Set rsViews = datGeneral.GetQuickEntryViews(!TableID)
            
        'Loop through the views, and see if we have permission on these
        Do While (Not rsViews.EOF) And (Not fCanView)
          If gcoTablePrivileges.item(rsViews!ViewName).AllowSelect Then
            'We have a view we can use, let's get outta here
            fCanView = True
          End If
        
          rsViews.MoveNext
        Loop
        rsViews.Close
        Set rsViews = Nothing
      End If
        
      .MoveNext
    Loop
    .Close
  End With
  Set rsTables = Nothing
    
  ViewQuickEntry = fCanView
  Exit Function
  
ErrorTrap:
  If Err = 457 Then
    Resume Next
  End If
  
End Function

Public Sub SetComboText(cboCombo As ComboBox, sText As String, Optional blnCaseInsensitive As Boolean)
  
  Dim lCount As Long

  With cboCombo
    For lCount = 1 To .ListCount
      If .List(lCount - 1) = sText Or (blnCaseInsensitive And LCase(.List(lCount - 1)) = LCase(sText)) Then
        .ListIndex = lCount - 1
        Exit For
      End If
    Next
  End With

End Sub

Public Function GetComboItem(cboTemp As ComboBox) As Long
  GetComboItem = 0
  If cboTemp.ListIndex <> -1 Then
    GetComboItem = cboTemp.ItemData(cboTemp.ListIndex)
  End If
End Function

'{MH20000202
Public Sub SetComboItem(cboCombo As ComboBox, lItem As Long)

    Dim lCount As Long
    
    With cboCombo
        For lCount = 1 To .ListCount
            If .ItemData(lCount - 1) = lItem Then
                .ListIndex = lCount - 1
                Exit For
            End If
        Next
    
    End With

End Sub
'MH20000202}


Public Function GetCountSql(psSQL As String) As String
  ' Return a SQL string that will return the number of records that
  ' would be returned by the given SQL string.
  Dim lngPosition1 As Long
  Dim lngPosition2 As Long
  Dim sSubSQL As String
  Dim sFromTable As String
  Dim sJoinWhereCode As String
  
  Const sFrom = " FROM "
  Const sORDERBY = " ORDER BY "
  
  lngPosition1 = InStr(1, UCase(psSQL), sFrom) + Len(sFrom)
  sSubSQL = Mid(psSQL, lngPosition1)
  lngPosition1 = InStr(1, UCase(sSubSQL), " ")
  sFromTable = Trim(Mid$(sSubSQL, 1, lngPosition1))
  sSubSQL = Mid(sSubSQL, lngPosition1)
  lngPosition1 = InStr(1, UCase(sSubSQL), sORDERBY)
  sJoinWhereCode = Mid$(sSubSQL, 1, lngPosition1)
  
  GetCountSql = "SELECT COUNT(" & sFromTable & ".id) FROM " & sFromTable & " " & sJoinWhereCode

End Function


Public Function FormatTime(strTime As String)

  'This will format a time string:
  'e.g. 12:4 will become 12:04 etc.
  'If an invalid time is passed then
  'this function will return vbNullString

  Dim intHours As Integer
  Dim intMinutes As Integer
  Dim intFound As Integer

  intFound = InStr(strTime, ":")
  If intFound > 0 Then
    intHours = Val(Left$(strTime, intFound - 1))
    intMinutes = Val(Mid$(strTime, intFound + 1))
  Else
    intHours = Val(Left$(strTime, 2))
    intMinutes = Val(Mid$(strTime, 3))
  End If

  If intHours < 24 And intMinutes < 60 Then
    FormatTime = ConvertMinsToTime((intHours * 60) + intMinutes)
  Else
    FormatTime = vbNullString
  End If

End Function


Public Function ConvertMinsToTime(lngMinutes As Long) As String
  'e.g. 570 becomes '09:30'
  ConvertMinsToTime = Right$("0" & CStr(lngMinutes \ 60), 2) & ":" & _
                      Right$("0" & CStr(lngMinutes Mod 60), 2)
End Function


Public Function ConvertTimeToMins(strTime As String) As Long
  'e.g. '09:30' becomes 570
  ConvertTimeToMins = (Val(Left$(strTime, 2)) * 60) + _
                       Val(Right$(strTime, 2))
End Function


Public Sub AddMinutes(dtDate As Date, strTime As String, lngAddMinutes As Long)

  Dim lngNewTime As Long

  lngNewTime = ConvertTimeToMins(strTime) + lngAddMinutes

  Do While lngNewTime < 0
    'Go back to previous day
    lngNewTime = lngNewTime + 1440
    dtDate = DateAdd("d", -1, dtDate)
  Loop

  Do While lngNewTime > 1439
    'Go forward to next day
    lngNewTime = lngNewTime - 1440
    dtDate = DateAdd("d", 1, dtDate)
  Loop

  strTime = ConvertMinsToTime(lngNewTime)

End Sub

Public Function ConvertSQLDateToLocale(psSQLDate As String) As String
  ' Convert the given date string (mm/dd/yyyy) into the locale format.
  ' NB. This function assumes a sensible locale format is used.
  Dim fDaysDone As Boolean
  Dim fMonthsDone As Boolean
  Dim fYearsDone As Boolean
  Dim sLocaleFormat As String
  Dim iLoop As Integer
  Dim sFormattedDate As String
  
  sFormattedDate = ""
  
  ' Get the locale's date format.
  sLocaleFormat = DateFormat
  
  fDaysDone = False
  fMonthsDone = False
  fYearsDone = False
  
  For iLoop = 1 To Len(sLocaleFormat)
    Select Case UCase(Mid(sLocaleFormat, iLoop, 1))
      Case "D"
        If Not fDaysDone Then
          sFormattedDate = sFormattedDate & Mid(psSQLDate, 4, 2)
          fDaysDone = True
        End If
        
      Case "M"
        If Not fMonthsDone Then
          sFormattedDate = sFormattedDate & Mid(psSQLDate, 1, 2)
          fMonthsDone = True
        End If
      
      Case "Y"
        If Not fYearsDone Then
          sFormattedDate = sFormattedDate & Mid(psSQLDate, 7, 4)
          fYearsDone = True
        End If
      
      Case Else
        sFormattedDate = sFormattedDate & Mid(sLocaleFormat, iLoop, 1)
    End Select
  Next iLoop
  
  ConvertSQLDateToLocale = sFormattedDate
  
End Function


Public Sub UtilityDefAmended(psTable As String, _
  psIDColumn As String, _
  plngRecordID As Long, _
  plngTimestamp As Long, _
  blnContinueSave As Boolean, _
  blnSaveAsNew As Boolean, _
  Optional strType As String)
  
  Dim datData As clsDataAccess
  Dim sSQL As String
  Dim rsCheck As Recordset
  Dim rsTemp As Recordset
  Dim sTemp As String
  Dim strMBText As String
  Dim intMBButtons As Long
  Dim intMBResponse As Integer
  
  Dim blnTimeStampChanged As Boolean
  Dim blnDeletedDef As Boolean
  Dim blnReadOnly As Boolean

  On Error GoTo Amended_ERROR
  
  If strType = vbNullString Then
    strType = "definition"
  End If

  blnSaveAsNew = False
  blnContinueSave = True
  
  If plngRecordID = 0 Then
    ' The record cannot have been modified by another user if it hs not yet been given a record ID.
    Exit Sub
  End If
  
  Set datData = New clsDataAccess
  ' Compare the given Timestamp with the Timestamp in the given record on the server.
  sSQL = "SELECT convert(int, timestamp) AS TimeStamp, Access, UserName " & _
         " FROM " & psTable & _
         " WHERE " & psIDColumn & " = " & Trim(Str(plngRecordID))
  Set rsCheck = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  blnTimeStampChanged = True
  blnReadOnly = False
  
  blnDeletedDef = (rsCheck.BOF And rsCheck.EOF)
  If Not blnDeletedDef Then
    blnTimeStampChanged = (plngTimestamp <> rsCheck!Timestamp)
    blnReadOnly = (LCase$(rsCheck!userName) <> LCase$(gsUserName) And rsCheck!Access <> ACCESS_READWRITE)
  End If
  
  rsCheck.Close
  Set rsCheck = Nothing
  Set datData = Nothing

  If blnTimeStampChanged = False Then
    Exit Sub
  End If
  
  If blnDeletedDef Or blnReadOnly Then
    
    'Unable to overwrite existing definition
    If blnDeletedDef Then
      strMBText = "The current " & strType & " has been deleted by another user."
    Else
      strMBText = "The current " & strType & " has been amended by another user and is now Read Only."
    End If
                  
    strMBText = strMBText & vbCrLf & _
                "Save as a new " & strType & "?"
    intMBButtons = vbExclamation + vbOKCancel
    intMBResponse = COAMsgBox(strMBText, intMBButtons, app.ProductName)
      
    Select Case intMBResponse
    Case vbOK         'save as new (but this may cause duplicate name message)
      blnContinueSave = True
      blnSaveAsNew = True
    Case vbCancel     'Do not save
      blnContinueSave = False
    End Select
      
  Else
    
    ' Use this to get the host name, as W95/8 doesnt like UI.GetHostName
    Set rsTemp = datGeneral.GetReadOnlyRecords("SELECT HOST_NAME()")
    sTemp = rsTemp.Fields(0)
    
    Set rsTemp = datGeneral.GetReadOnlyRecords("SELECT SavedHost " & _
                                               "FROM   ASRSysUtilAccessLog " & _
                                               "WHERE  Type IN (" & GetTypeCodeFromTable(psTable) & ") " & _
                                               "AND    UtilID = " & plngRecordID)
    
'    If LCase(rsTemp.Fields("SavedHost")) <> LCase((UI.GetHostName)) Then
    If LCase(rsTemp.Fields("SavedHost")) <> LCase((sTemp)) Then
      
      ' If the definition was last changed by somebody else (rather than by
      ' automatically due to the access rights of a plist/filter/calc being
      ' changed, then prompt, otherwise, just overwrite it.
      
      'Prompt to see if user should overwrite definition
      strMBText = "The current " & strType & " has been amended by another user. " & vbCrLf & _
                  "Would you like to overwrite this " & strType & "?" & vbCrLf
      intMBButtons = vbExclamation + vbYesNoCancel
      intMBResponse = COAMsgBox(strMBText, intMBButtons, app.ProductName)
      
      Select Case intMBResponse
      Case vbYes        'overwrite existing definition and any changes
        blnContinueSave = True
      Case vbNo         'save as new (but this may cause duplicate name message)
        blnContinueSave = True
        blnSaveAsNew = True
      Case vbCancel     'Do not save
        blnContinueSave = False
      End Select
      
    Else
      blnContinueSave = True
    End If
    
  End If

  Exit Sub
  
Amended_ERROR:
  
  COAMsgBox "Error whilst checking if utility definition has been amended." & vbCrLf & vbCrLf & "(" & Err.Number & " - " & Err.Description & ")", vbExclamation + vbOKOnly, app.title
  blnContinueSave = False
  
End Sub

Private Function GetTypeCodeFromTable(sTable As String) As String

  Select Case LCase(sTable)
  
  Case "asrsysbatchjob": GetTypeCodeFromTable = "0"
  Case "asrsyscalendarreports": GetTypeCodeFromTable = "15"
  Case "asrsyscrosstab": GetTypeCodeFromTable = "1"
  Case "asrsyscustomreportsname": GetTypeCodeFromTable = "2"
  Case "asrsysdatatransfername": GetTypeCodeFromTable = "3"
  Case "asrsysexportname": GetTypeCodeFromTable = "4"
  Case "asrsysglobalfunctions": GetTypeCodeFromTable = "5,6,7"  '### is this ok ?
  Case "asrsysimport": GetTypeCodeFromTable = "8"
  Case "asrsysmailmergename": GetTypeCodeFromTable = "9"
  Case "asrsysmatchreportname": GetTypeCodeFromTable = "14"
  Case "asrsyspicklistname": GetTypeCodeFromTable = "10"
  Case "asrsysexpressions": GetTypeCodeFromTable = "11,12"  '### is this ok ?
  Case "asrsysorders": GetTypeCodeFromTable = "13"
  Case "asrsysrecordprofilename": GetTypeCodeFromTable = "20"
  Case Else
    Stop
    
  End Select

'  utlBatchJob = 0
'  utlCrossTab = 1
'  utlCustomReport = 2
'  utlDataTransfer = 3
'  utlExport = 4
'  utlGlobalAdd = 5
'  utlGlobalDelete = 6
'  utlGlobalUpdate = 7
'  utlImport = 8
'  utlMailMerge = 9
'  utlPicklist = 10
'  utlFilter = 11
'  utlCalculation = 12
'  utlOrder = 13

End Function

Public Function FormatEventDuration(lngSeconds As Long) As String

  Dim strHours As String
  Dim strMins As String
  Dim strSeconds As String
  Dim dblRemainder As Double
  'Dim bob As String
  
  Dim TIME_SEPARATOR  As String
  
  'NHRD17082004 Fault 8808 Changed this from a constant.
  'This will enable duration in the event log to be displayed properly.
  'I don't think it will affect any other time calulations
  'Const TIME_SEPARATOR = ":"
  TIME_SEPARATOR = UI.GetSystemTimeSeparator
  
  If Not (lngSeconds < 0) Then
    strHours = CStr(Fix(lngSeconds / 3600))
    
    If Len(strHours) <= 2 Then
      strHours = String((2 - Len(strHours)), "0") & strHours
    End If
    
    dblRemainder = CDbl(lngSeconds Mod 3600)
    
    strMins = CStr(Fix((dblRemainder / 60)))
    strMins = String((2 - Len(strMins)), "0") & strMins
    dblRemainder = CDbl(dblRemainder Mod 60)
    
    strSeconds = CStr(Fix(dblRemainder))
    strSeconds = String((2 - Len(strSeconds)), "0") & strSeconds
    
    FormatEventDuration = strHours & TIME_SEPARATOR & strMins & TIME_SEPARATOR & strSeconds
  Else
    FormatEventDuration = ""
  End If
  
End Function

Public Sub SetDateComboFormat(cboDate As GTMaskDate.GTMaskDate)

  Dim sDateFormat As String
  
  sDateFormat = DateFormat
  
  cboDate.Format = sDateFormat
  cboDate.DisplayFormat = sDateFormat
  
  sDateFormat = Replace(sDateFormat, "d", "_")
  sDateFormat = Replace(sDateFormat, "m", "_")
  sDateFormat = Replace(sDateFormat, "y", "_")
  
  cboDate.Text = sDateFormat

End Sub

'Public Sub ControlsDisableAll(frmCurrent As Form)
'
'  'Not all controls have a backcolor !
'  On Local Error Resume Next
'
'  Dim ctl As Control
'  For Each ctl In frmCurrent
'
'    If TypeOf ctl Is MSComctlLib.TabStrip Or _
'           TypeOf ctl Is ComctlLib.TabStrip Or _
'           TypeOf ctl Is SSTab Then
'      'Stop
'
'    ElseIf TypeOf ctl Is Frame Or _
'           TypeOf ctl Is PictureBox Or _
'           TypeOf ctl Is MSComctlLib.ListView Or _
'           TypeOf ctl Is ComctlLib.ListView Or _
'           TypeOf ctl Is ListBox Then
'      'Just make container controls and scroll-able controls look disabled...
'      '(NOTE: Code must be placed in drag-drop events etc. to disable it)
'      ctl.ForeColor = vbGrayText
'      ctl.BackColor = vbButtonFace
'
'      'TM20010822 Fault 2566
'      'Controls of this type need to be enabled so that scroll/click functionality
'      'is allowed.
'      ctl.Enabled = True
'
'    ElseIf TypeOf ctl Is SSDBGrid Then
'
'      If Not (TypeOf frmCurrent Is frmImport) And _
'         Not (TypeOf frmCurrent Is frmExport) Then
'        ctl.Enabled = False
'
'      End If
'
'      ctl.BackColorEven = vbButtonFace    'SSDBGrid
'      ctl.BackColorOdd = vbButtonFace     'SSDBGrid
'
'    ElseIf TypeOf ctl Is CommandButton Then
'      'Disable all CommandButtons except cancel...
'
'      If ctl.Cancel = False Then
'        ctl.Enabled = False
'      Else
'        ctl.Enabled = True
'      End If
'
'    Else
'      ctl.ForeColor = vbApplicationWorkspace
'      ctl.BackColor = vbButtonFace
'      'ctl.Enabled = False
'
'      'Maybe put this in ????
'      If TypeOf ctl Is TextBox Then
'        ctl.Locked = True
'        ctl.TabStop = False
'      Else
'        ctl.Enabled = False
'      End If
'    End If
'
'  Next
'
'End Sub

Public Sub ControlsDisableAll(objCurrent As Object, Optional blnEnabled As Boolean = False)

  Dim ctl As Control

  'Not all controls have a backcolor !
  On Local Error Resume Next

  If TypeOf objCurrent Is Form Then
    For Each ctl In objCurrent
      
      EnableControl ctl, blnEnabled
    Next

  Else
    EnableControl objCurrent, blnEnabled
    For Each ctl In objCurrent.Parent
      If ctl.Container.Name = objCurrent.Name Then
        EnableControl ctl, blnEnabled
      End If
    Next

  End If

End Sub

Public Function EnableControl(ctl As Control, blnEnabled As Boolean)

  ' JPD20020920 Added ActiveBar to the list of controls that are not enabled/disabled, as it has no
  ' 'enabled' property. Instead use the 'EnableActiveBar' method.
  If TypeOf ctl Is MSComctlLib.TabStrip Or _
         TypeOf ctl Is ComctlLib.TabStrip Or _
         TypeOf ctl Is ActiveBar Or _
         TypeOf ctl Is SSTab Then
    'Stop

  ElseIf TypeOf ctl Is Frame Or _
         TypeOf ctl Is PictureBox Or _
         TypeOf ctl Is MSComctlLib.ListView Or _
         TypeOf ctl Is ComctlLib.ListView Or _
         TypeOf ctl Is ListBox Or _
         TypeOf ctl Is Label Then
         
    'Just make container controls and scroll-able controls look disabled...
    '(NOTE: Code must be placed in drag-drop events etc. to disable it)
    ctl.ForeColor = IIf(blnEnabled, vbWindowText, vbApplicationWorkspace)
    If (TypeOf ctl Is ListBox) Or _
        (TypeOf ctl Is MSComctlLib.ListView) Or _
        (TypeOf ctl Is ComctlLib.ListView) Then
        
      ctl.BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)    'SSDBGrid
    End If

    'TM20010822 Fault 2566
    'Controls of this type need to be enabled so that scroll/click functionality
    'is allowed.
    ctl.Enabled = True

  ElseIf TypeOf ctl Is SSDBGrid Then

    ctl.Enabled = blnEnabled
    ctl.BackColorEven = IIf(blnEnabled, vbWindowBackground, vbButtonFace)   'SSDBGrid
    ctl.BackColorOdd = IIf(blnEnabled, vbWindowBackground, vbButtonFace)   'SSDBGrid

  ElseIf TypeOf ctl Is CommandButton Then   'Or _
         TypeOf ctl Is SSCommand Then
    'Disable all CommandButtons except cancel...

    If ctl.Cancel = False Then
      ctl.Enabled = blnEnabled
    Else
      ctl.Enabled = True
    End If

  'ElseIf (TypeOf ctl Is SSCheck) Or (TypeOf ctl Is CheckBox) Then
  ElseIf (TypeOf ctl Is CheckBox) Then

    ctl.ForeColor = IIf(blnEnabled, vbWindowText, vbApplicationWorkspace)
    ctl.BackColor = vbButtonFace
    ctl.Enabled = blnEnabled

  ElseIf (TypeOf ctl Is UpDown) Then
    ctl.Enabled = blnEnabled

  ElseIf (TypeOf ctl Is TextBox) Then
    ctl.ForeColor = IIf(blnEnabled, vbWindowText, vbApplicationWorkspace)
    ctl.BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)
    ctl.Locked = Not blnEnabled
    ctl.TabStop = blnEnabled
    ctl.Enabled = blnEnabled
  
  Else
    ctl.ForeColor = IIf(blnEnabled, vbWindowText, vbApplicationWorkspace)
    ctl.BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)
    ctl.Enabled = blnEnabled
  
  End If

End Function


Public Function vbCompiled() As Boolean

  'Dim nRtn As Long
  'Dim Buffer As String
  'Buffer = Space$(256)
  'nRtn = GetModuleFileNameA(0&, Buffer, Len(Buffer))
  'Buffer = UCase(Left(Buffer, nRtn))
  'vbCompiled = (Right(Buffer, 8) <> "\VB6.EXE")

  'Much better (and clever-er) !
  On Local Error Resume Next
  Err.Clear
  Debug.Print 1 / 0
  vbCompiled = (Err.Number = 0)

  'JDM - 26/09/01 - Fault 1924 - I want to run code with break on all errors
  'vbCompiled = IIf(VB.App.EXEName = "Prjhrpro", False, True)

End Function



Private Function GetFullAccessChildView(plngTableID As Long) As Long
'  ' Return the child view that gives full access to the given table.
'  Dim iNextIndex As Integer
'  Dim sSQL As String
'  Dim sParentSQL As String
'  Dim rsInfo As Recordset
'  Dim avParents() As Variant
'
'  ' Construct an array of the required child view's parents.
'  ' Column 1 is the parent type - UT = user-defined top-level table
'  '                               UV = user-defined top-level view
'  '                               SV = system generated child view
'  ' Column 2 is the parent ID.
'  ReDim avParents(2, 0)
'
'  sSQL = "SELECT ASRSysRelations.parentID,  ASRSysTables.tableType" & _
'    " FROM ASRSysRelations" & _
'    " INNER JOIN ASRSysTables ON ASRSysRelations.parentID = ASRSysTables.tableID" & _
'    " WHERE ASRSysRelations.childID = " & Trim(Str(plngTableID))
'  Set rsInfo = New Recordset
'  rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'  Do While Not rsInfo.EOF
'    iNextIndex = UBound(avParents, 2) + 1
'    ReDim Preserve avParents(2, iNextIndex)
'
'   If rsInfo!TableType = tabTopLevel Then
'      avParents(1, iNextIndex) = "UT"
'      avParents(2, iNextIndex) = rsInfo!ParentID
'    Else
'      avParents(1, iNextIndex) = "SV"
'      avParents(2, iNextIndex) = GetFullAccessChildView(rsInfo!ParentID)
'    End If
'
'    rsInfo.MoveNext
'  Loop
'  rsInfo.Close
'  Set rsInfo = Nothing
'
'  sParentSQL = ""
'  For iNextIndex = 1 To UBound(avParents, 2)
'    If avParents(1, iNextIndex) = "UT" Then
'      sParentSQL = sParentSQL & _
'        " INNER JOIN ASRSysChildViewParents tmpTable_" & Trim(Str(iNextIndex)) & _
'        " ON (ASRSysChildViews.childViewID = tmpTable_" & Trim(Str(iNextIndex)) & ".childViewID" & _
'        " AND tmpTable_" & Trim(Str(iNextIndex)) & ".parentType = 'UT'" & _
'        " AND tmpTable_" & Trim(Str(iNextIndex)) & ".parentID = " & Trim(Str(avParents(2, iNextIndex))) & ")"
'    Else
'      sParentSQL = sParentSQL & _
'        " INNER JOIN ASRSysChildViewParents tmpTable_" & Trim(Str(iNextIndex)) & _
'        " ON (ASRSysChildViews.childViewID = tmpTable_" & Trim(Str(iNextIndex)) & ".childViewID" & _
'        " AND tmpTable_" & Trim(Str(iNextIndex)) & ".parentType = 'SV'" & _
'        " AND tmpTable_" & Trim(Str(iNextIndex)) & ".parentID = " & Trim(Str(avParents(2, iNextIndex))) & ")"
'    End If
'  Next iNextIndex
'
'  sSQL = "SELECT ASRSysChildViews.childViewID" & _
'    " FROM ASRSysChildViews" & _
'    sParentSQL & _
'    " INNER JOIN ASRSysChildViewParents parentCount" & _
'    " ON (ASRSysChildViews.childViewID = parentCount.childViewID)" & _
'    " GROUP BY ASRSysChildViews.childViewID, ASRSysChildViews.tableID, ASRSysChildViews.type" & _
'    " HAVING ASRSysChildViews.tableID = " & Trim(Str(plngTableID)) & _
'    " AND (ASRSysChildViews.type = 0 OR ASRSysChildViews.type IS NULL)" & _
'    " AND COUNT(parentCount.childViewID) = " & Trim(Str(UBound(avParents, 2)))
'  Set rsInfo = New Recordset
'  rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'  GetFullAccessChildView = 0
'
'  If Not rsInfo.EOF Then
'    GetFullAccessChildView = IIf(IsNull(rsInfo!childViewID), 0, rsInfo!childViewID)
'  End If
'
'  rsInfo.Close
'  Set rsInfo = Nothing
'
End Function

Public Function GetLastField(strInput As String, strDelim As String) As String
  GetLastField = Mid(strInput, InStrRev(strInput, strDelim) + 1)
End Function

'Public Sub FormStateSave(frm As Form)
'
'  With frm
'    SavePCSetting .Name, "Left", .Left
'    SavePCSetting .Name, "Top", .Top
'    SavePCSetting .Name, "Width", .Width
'    SavePCSetting .Name, "Height", .Height
'    SavePCSetting .Name, "State", .WindowState
'  End With
'
'End Sub
'
'Public Sub FormStateRestore(frm As Form)
'
'  With frm
'    .Left = GetPCSetting( .Name, "Left", .Left)
'    .Top = GetPCSetting( .Name, "Top", .Top)
'    .Width = GetPCSetting( .Name, "Width", .Width)
'    .Height = GetPCSetting( .Name, "Height", .Height)
'    .WindowState = GetPCSetting( .Name, "State", .WindowState)
'  End With
'
'End Sub

Public Function IsPicklistValid(varID As Variant) As String
  IsPicklistValid = IsSelectionValid(varID, "picklist")
End Function

Public Function IsFilterValid(varID As Variant) As String

  Dim objExpr As clsExprExpression
  'Dim strRuntimeCode As String
  Dim strFilterName As String
  Dim astrUDFsRequired() As String

  On Local Error GoTo LocalErr

  ReDim astrUDFsRequired(0)

  strFilterName = vbNullString
  IsFilterValid = IsSelectionValid(varID, "filter")

  If IsFilterValid = vbNullString Then
    Set objExpr = New clsExprExpression
    With objExpr
      'JPD 20030324 Fault 5160
      .ExpressionID = CLng(varID)
      .ConstructExpression
      If (.ValidateExpression(True) <> giEXPRVALIDATION_NOERRORS) Then
        IsFilterValid = "The filter '" & strFilterName & "' used in this definition is invalid."
      End If
      
'      If .Initialise(0, CLng(varID), giEXPR_RUNTIMEFILTER, giEXPRVALUE_LOGIC) Then
'        strFilterName = .Name
'        If objExpr.RuntimeFilterCode(strRunTimeCode, True, True) Then
'
'          If gbEnableUDFFunctions Then
'            objExpr.UDFFilterCode astrUDFsRequired(), True, True
'          End If
'
'          UDFFunctions astrUDFsRequired, True
'          datGeneral.GetReadOnlyRecords strRunTimeCode
'          UDFFunctions astrUDFsRequired, False
'
'        End If
'      End If

    End With
    Set objExpr = Nothing
  End If

Exit Function

LocalErr:
  If strFilterName <> vbNullString Then
    IsFilterValid = "'" & strFilterName & "' "
  End If
  IsFilterValid = "The filter " & IsFilterValid & "used in this definition is invalid."

End Function

Public Function IsCalcValid(varID As Variant) As String
  IsCalcValid = IsSelectionValid(varID, "calculation")
End Function

Private Function IsSelectionValid(varID As Variant, strType As String) As String

  Dim rsTemp As Recordset

  IsSelectionValid = vbNullString
  If Val(varID) = 0 Then Exit Function

  Set rsTemp = GetSelectionAccess(varID, strType)
  
  If strType = "picklist" Then
    If rsTemp.BOF And rsTemp.EOF Then
      IsSelectionValid = _
        "The " & strType & " used in this definition has been " & _
        "deleted by another user."
    
    ElseIf LCase(Trim(rsTemp!userName)) <> LCase(Trim(gsUserName)) And _
          (rsTemp!Access = ACCESS_HIDDEN) Then
      IsSelectionValid = _
        "The " & strType & " used in this definition has been made " & _
        "hidden by another user."
    End If
  Else
    'TM20010807 Fault 2656 - If not a picklist then check if the expression has
    'hidden components.
    If rsTemp.BOF And rsTemp.EOF Then
      IsSelectionValid = _
        "The " & strType & " used in this definition has been " & _
        "deleted by another user."
    
    ElseIf LCase(Trim(rsTemp!userName)) <> LCase(Trim(gsUserName)) And _
          (rsTemp!Access = ACCESS_HIDDEN Or HasHiddenComponents(CLng(varID))) Then
      IsSelectionValid = _
        "The " & strType & " used in this definition has been made " & _
        "hidden by another user."
    End If
  End If
  Set rsTemp = Nothing

End Function

Public Function GetSelectionAccess(varID As Variant, strType As String) As Recordset

  Dim strSQL As String
  'Dim rsTemp As Recordset

  If strType = "picklist" Then
    strSQL = "SELECT Access, UserName FROM AsrSysPicklistName " & _
             "WHERE PickListID = " & CStr(varID)
  Else
    strSQL = "SELECT Access, UserName FROM AsrSysExpressions " & _
             "WHERE ExprID = " & CStr(varID)
  End If
  Set GetSelectionAccess = datGeneral.GetReadOnlyRecords(strSQL)

End Function

Public Sub LoadTableCombo(cboTemp As ComboBox, Optional strSQL As String)
  
  Dim rsTemp As Recordset

  If strSQL = vbNullString Then
    'MH20011025 Fault 3030
    'strSQL = "SELECT TableID, TableName FROM ASRSysTables " & _
             "WHERE tableType = " & Trim(Str(tabTopLevel)) & _
             " OR tableType = " & Trim(Str(tabChild)) & _
             " ORDER BY TableName"
    strSQL = "SELECT TableID, TableName FROM ASRSysTables " & _
             " ORDER BY TableName"
  End If
  Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)
  
  With cboTemp
    .Clear
    
    Do While Not rsTemp.EOF
      .AddItem rsTemp.Fields(1)
      .ItemData(.NewIndex) = Val(rsTemp.Fields(0))
      rsTemp.MoveNext
    Loop

    If .ListCount > 0 Then
      SetComboItem cboTemp, glngPersonnelTableID
      If .ListIndex < 0 Then
        .ListIndex = 0
      End If
    End If
    
  End With

  rsTemp.Close
  Set rsTemp = Nothing

End Sub

Public Function LoadTableCombo2(cboTemp As ComboBox)

  Dim objTableView As CTablePrivilege
  
  'MH20020823
  'This should be a bit quicker than "LoadTableCombo" as it doesn't
  'go back and hit the server, instead reads from the collection.

  With cboTemp
    .Clear
  
    For Each objTableView In gcoTablePrivileges.Collection
      If objTableView.IsTable Then
        .AddItem objTableView.TableName
        .ItemData(.NewIndex) = objTableView.TableID
      End If
    Next
    
    If .ListCount > 0 Then
      SetComboItem cboTemp, glngPersonnelTableID
      If .ListIndex < 0 Then
        .ListIndex = 0
      End If
    End If
  End With

End Function

Public Function SetStringLength(ByVal psInputString As String, piLength As Integer) As String

  ' Sets an input field to the desired length
  ' i.e. trims or pads out with spaces

  If Len(psInputString) > piLength Then
    SetStringLength = Left$(psInputString, piLength)
  Else
    SetStringLength = psInputString & Space$(piLength - Len(psInputString))
  End If

End Function

Private Function GetDBVersion() As String

  Dim rsInfo As Recordset
  
  GetDBVersion = GetSystemSetting("Database", "Version", vbNullString)

  If GetDBVersion = vbNullString Then
    Set rsInfo = datGeneral.GetReadOnlyRecords("SELECT SystemManagerVersion FROM ASRSysConfig")
  
    If Not rsInfo.BOF And Not rsInfo.EOF Then
      GetDBVersion = rsInfo.Fields(0).Value
    End If
  
    rsInfo.Close
    Set rsInfo = Nothing
  
  End If

End Function

Private Function Office_IsWordInstalled() As Boolean

  On Error GoTo NotInstalled

  Dim app As New Word.Application
  Set app = CreateObject("Word.Application")
  app.Quit

  Office_IsWordInstalled = True

TidyUpAndExit:
  Set app = Nothing
  Exit Function

NotInstalled:
  Office_IsWordInstalled = False
  Resume TidyUpAndExit

End Function

Private Function Office_IsExcelInstalled() As Boolean

  On Error GoTo NotInstalled

  Dim app As New Excel.Application
  Set app = CreateObject("Excel.Application")
  app.Quit

  Office_IsExcelInstalled = True
  
TidyUpAndExit:
  Set app = Nothing
  Exit Function

NotInstalled:
  Office_IsExcelInstalled = False
  Resume TidyUpAndExit
  
End Function

Public Function IsKeyword(ByVal strCheckWord As String) As Boolean

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "modHrPro.IsKeyword(strCheckWord)", Array(strCheckWord)
  
  Dim sSQL As String
  Dim rsKeywords As ADODB.Recordset
    
  ' Open the keywords resultset.
  sSQL = "SELECT keyword FROM ASRSysKeywords" & _
    " WHERE " & _
    " keyword='" & strCheckWord & "'"
    
  Set rsKeywords = New ADODB.Recordset
  rsKeywords.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  
  ' If the keywords resultset contains any records then the word is a keyword.
  IsKeyword = (Not (rsKeywords.BOF And rsKeywords.EOF))
  
  'Close and release keywords recordset
  rsKeywords.Close
  Set rsKeywords = Nothing

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Function
ErrorTrap:
  gobjErrorStack.HandleError

End Function


Public Function ValidateGTMaskDate(dtTemp As GTMaskDate.GTMaskDate) As Boolean

  Dim blnYearOkay As Boolean
  Dim sDateSeparator As String
  
  Dim strDateFormat As String
  Dim intDayIndex As Integer
  Dim intMonthIndex As Integer
  Dim intYearIndex As Integer
  
  ValidateGTMaskDate = True

  strDateFormat = DateFormat
  
  intDayIndex = InStr(1, strDateFormat, "dd")
  intMonthIndex = InStr(1, strDateFormat, "mm")
  intYearIndex = InStr(1, strDateFormat, "yyyy")
  
  'TM20020610 Fault 3855 - use the User Inteface System Date Separator.
  sDateSeparator = UI.GetSystemDateSeparator
  
  With dtTemp
    If Trim(Replace(.Text, sDateSeparator, "")) <> vbNullString Then
  
      'MH20020423 Fault 3760 (Avoid changing 01/13/2002 to 13/01/2002)
      'If Not IsDate(.DateValue) Or .DateValue < #1/1/1753# Then
      'If Not IsDate(.DateValue) Or .DateValue < #1/1/1753# Or Left(.Text, 5) <> Left(.DateValue, 5) Then
      
      'MH20020423 Fault 3543 Also make sure that they enter a valid year
      blnYearOkay = (Val(Mid(.Text, intYearIndex, 4)) >= 1753)
      
      If (Not IsDate(.DateValue)) Or (.DateValue < #1/1/1753#) Or (Not blnYearOkay) _
            Or Format(.DateValue, DateFormat) <> .Text Or IsNull(.DateValue) Then   'MH20030905 Fault 6290

        Clipboard.Clear
        Clipboard.SetText .Text
        .DateValue = Null
        .Paste
  
        .ForeColor = vbRed

        'MH20020712 Fault 4131
        'COAMsgBox sometimes causes run time error but DoEvents prevents this!
        DoEvents

        COAMsgBox "You have entered an invalid date.", vbOKOnly + vbExclamation, app.title
        .ForeColor = vbWindowText
        .DateValue = Null
        If .Visible And .Enabled Then
          .SetFocus
        End If
        ValidateGTMaskDate = False
      End If
    Else
'      'NHRD09092004 Fault 8895
'      If .Text = "  /  /" Then
'        Clipboard.Clear
'        Clipboard.SetText .Text
'        .DateValue = Null
'        .Paste
'        .ForeColor = vbRed
'
'        DoEvents
'
'        COAMsgBox "No date has been selected.", vbOKOnly + vbExclamation, App.Title
'        .ForeColor = vbWindowText
'        .DateValue = Null
'        If .Visible And .Enabled Then
'          .SetFocus
'        End If
'        ValidateGTMaskDate = False
'      End If
    End If
    
  End With

End Function


Public Sub EnableCombo(cboTemp As ComboBox, ByVal blnEnabled As Boolean)
  blnEnabled = (blnEnabled And cboTemp.ListCount > 0)
  cboTemp.Enabled = blnEnabled
  cboTemp.BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)
  cboTemp.ListIndex = IIf(blnEnabled, 0, -1)
End Sub


Public Function GetColour(ByVal sColour As String) As Long

  Dim rsTemp As Recordset
  Dim strSQL As String

  On Local Error GoTo LocalErr

  strSQL = "SELECT ColValue FROM ASRSysColours " & _
           " WHERE ColDesc = '" & LCase(sColour) & "'"
  Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)

  With rsTemp
    If Not .BOF And Not .EOF Then
      GetColour = rsTemp.Fields("ColValue").Value
    Else
      GetColour = 0
    End If
  End With

  rsTemp.Close
  Set rsTemp = Nothing

Exit Function

LocalErr:
  GetColour = 0
End Function


Public Function GetWordColourIndex(lngColourValue As Long) As Long

  Dim rsTemp As Recordset
  Dim strSQL As String

  On Local Error GoTo LocalErr

  strSQL = "SELECT WordColourIndex FROM ASRSysColours " & _
           " WHERE ColValue = " & CStr(lngColourValue)
  Set rsTemp = datGeneral.GetReadOnlyRecords(strSQL)
  
  With rsTemp
    If Not .BOF And Not .EOF Then
      GetWordColourIndex = rsTemp.Fields("WordColourIndex").Value
    Else
      GetWordColourIndex = 0
    End If
  End With

  rsTemp.Close
  Set rsTemp = Nothing

Exit Function

LocalErr:
  GetWordColourIndex = 0

End Function

Public Function CheckPlatform() As Boolean

  Dim sSQL As String, sMsg As String
  Dim lngSQLVersion As Double
  
  Dim strLastSQLServerVersion As String
  Dim strLastDatabaseName As String
  Dim strLastServerName As String
  Dim strOldServerName As String
  
  Dim rsSQLInfo As ADODB.Recordset
  
  ' Get the SQL Server version number.
  lngSQLVersion = 0
  sSQL = "master..xp_msver ProductVersion"
  Set rsSQLInfo = New ADODB.Recordset
  rsSQLInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsSQLInfo
    If Not (.BOF And .EOF) Then
      lngSQLVersion = Val(.Fields("character_value").Value)
    End If
    .Close
  End With
  Set rsSQLInfo = Nothing

  sMsg = vbNullString
  strLastSQLServerVersion = GetSystemSetting("Platform", "SQLServerVersion", 0)
  strLastDatabaseName = UCase$(GetSystemSetting("Platform", "DatabaseName", ""))
  strLastServerName = UCase$(GetSystemSetting("Platform", "ServerName", ""))
  If strLastServerName = "." Then strLastServerName = UCase$(UI.GetHostName)
  strOldServerName = ""

  If GetOldServerName <> GetServerName Then
    sMsg = "The Microsoft SQL Server has been renamed but the operation is incomplete."
    
    COAMsgBox sMsg & vbCrLf & _
            "Please contact your System Administrator", _
            vbOKOnly + vbExclamation, Application.Name
    
    CheckPlatform = False
      
    Exit Function
  Else
    If Val(strLastSQLServerVersion) <> lngSQLVersion Then
      sMsg = "The Microsoft SQL Version has been upgraded."
    ElseIf strLastServerName <> GetServerName() Then
      sMsg = "The database has moved to a different Microsoft SQL Server."
    ElseIf strLastDatabaseName <> GetDBName() Then
      sMsg = "The database name has changed."
  '  ElseIf Not FrameworkVersionOK Then
  '    sMsg = "The Microsoft .NET Framework version has changed on the server."
    End If
    
    If sMsg <> vbNullString Then
      COAMsgBox sMsg & vbCrLf & _
            "Please ask the System Administrator to update the database in the System Manager.", _
            vbOKOnly + vbExclamation, Application.Name
    
      CheckPlatform = False
      
      Exit Function
    End If
End If

  CheckPlatform = True
    
End Function

Public Function UDFFunctions(pastrUDFFunctions() As String, pbCreate As Boolean) As Boolean

  On Error GoTo UDFFunctions_ERROR

  Dim iCount As Integer
  Dim strDropCode As String
  Dim strFunctionName As String
  Dim sUDFCode As String
  Dim datData As clsDataAccess
  Dim iStart As Integer
  Dim iEnd As Integer
  Dim strFunctionNumber As String

  Const FUNCTIONPREFIX = "udf_ASRSys_"
  
  Set datData = New clsDataAccess
                       
  If gbEnableUDFFunctions Then
            
    For iCount = 1 To UBound(pastrUDFFunctions)
    
      'JPD 20060109 Fault 10509
      'iStart = Len("CREATE FUNCTION udf_ASRSys_") + 1
      iStart = InStr(pastrUDFFunctions(iCount), FUNCTIONPREFIX) + Len(FUNCTIONPREFIX)
      iEnd = InStr(1, Mid(pastrUDFFunctions(iCount), 1, 1000), "(@Pers")
      strFunctionNumber = Mid(pastrUDFFunctions(iCount), iStart, iEnd - iStart)
      strFunctionName = FUNCTIONPREFIX & strFunctionNumber
    
      'Drop existing function (could exist if the expression is used more than once in a report)
      strDropCode = "IF EXISTS" & _
        " (SELECT Name" & _
        "   FROM sysobjects" & _
        "   WHERE id = object_id('[" & datGeneral.UserNameForSQL & "]." & strFunctionName & "')" & _
        "     AND sysstat & 0xf = 0)" & _
        " DROP FUNCTION [" & gsUserName & "]." & strFunctionName
      datData.ExecuteSql strDropCode
       
      ' Create the new function
      If pbCreate Then
        sUDFCode = pastrUDFFunctions(iCount)
        datData.ExecuteSql sUDFCode
      End If
    
    Next iCount
  End If

  UDFFunctions = True
  Exit Function
  
UDFFunctions_ERROR:
  UDFFunctions = False
  
End Function

Public Function GetUtilityStatus(pintID As Integer) As String

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "modHRPro.GetUtilityStatus(pintID)", Array(pintID)

  ' Returns the user friendly description of status give the code
  Select Case pintID
    Case 0: GetUtilityStatus = "Pending"
    Case 1: GetUtilityStatus = "Cancelled"
    Case 2: GetUtilityStatus = "Failed"
    Case 3: GetUtilityStatus = "Successful"
    Case 4: GetUtilityStatus = "Skipped"
    Case 5: GetUtilityStatus = "Error"
    Case Else: GetUtilityStatus = "Unknown"
  End Select

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Function
ErrorTrap:
  gobjErrorStack.HandleError
  
End Function

Public Function GetUtilityType(pintID As Integer) As String

  'Dim lngIndex As Long

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "modHRPro.GetUtilityType(pintID)", Array(pintID)
  
  ' Returns the user friendly description of utility type give the code
  Select Case pintID
    Case eltCrossTab: GetUtilityType = "Cross Tab"
    Case eltCustomReport: GetUtilityType = "Custom Report"
    Case eltDataTransfer: GetUtilityType = "Data Transfer"
    Case eltExport: GetUtilityType = "Export"
    Case eltGlobalAdd: GetUtilityType = "Global Add"
    Case eltGlobalDelete: GetUtilityType = "Global Delete"
    Case eltGlobalUpdate: GetUtilityType = "Global Update"
    Case eltImport: GetUtilityType = "Import"
    Case eltMailMerge: GetUtilityType = "Mail Merge"
    Case eltDiaryDelete: GetUtilityType = "Diary Delete"
    Case eltDiaryRebuild: GetUtilityType = "Diary Rebuild"
    Case eltEmailRebuild: GetUtilityType = "Email Rebuild"
    Case eltStandardReport: GetUtilityType = "Standard Report"
    Case eltRecordEditing: GetUtilityType = "Record Editing"
    Case eltSystemError: GetUtilityType = "System Error"
    Case eltMatchReport: GetUtilityType = "Match Report"
    Case eltCalandarReport: GetUtilityType = "Calendar Report"
    Case eltLabel: GetUtilityType = "Envelopes & Labels"
    Case eltLabelType: GetUtilityType = "Label Definition"
    Case eltRecordProfile: GetUtilityType = "Record Profile"
    Case eltSuccessionPlanning: GetUtilityType = "Succession Planning"
    Case eltCareerProgression: GetUtilityType = "Career Progression"
    Case eltAccordImport: GetUtilityType = "Payroll Transfer (In)"
    Case eltAccordExport: GetUtilityType = "Payroll Transfer (Out)"
    Case eltWorkflowRebuild: GetUtilityType = "Workflow Rebuild"
    Case elt9BoxGrid: GetUtilityType = "9-Box Grid Report"
    Case eltTalentReport: GetUtilityType = "Talent Report"
    Case Else: GetUtilityType = "Unknown"
  End Select

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Function
  
ErrorTrap:
  gobjErrorStack.HandleError
  
End Function

Public Function GetJobName(pstrJobType As String, plngJobID As Long) As String

  Dim pstrSQL As String
  Dim prstTemp As Recordset
  
  pstrSQL = vbNullString
  
  Select Case UCase(pstrJobType)
  Case "CALENDAR REPORT"
    pstrSQL = "SELECT Name From ASRSysCalendarReports WHERE ID = " & plngJobID
  Case "CROSS TAB"
    pstrSQL = "SELECT Name From ASRSysCrossTab WHERE CrossTabID = " & plngJobID
  Case "CUSTOM REPORT"
    pstrSQL = "SELECT Name From ASRSysCustomReportsName WHERE ID = " & plngJobID
  Case "DATA TRANSFER"
    pstrSQL = "SELECT Name From ASRSysDataTransferName WHERE DataTransferID = " & plngJobID
  Case "EXPORT"
    pstrSQL = "SELECT Name From ASRSysExportName WHERE ID = " & plngJobID
  Case "GLOBAL ADD"
    pstrSQL = "SELECT Name From ASRSysGlobalFunctions WHERE FunctionID = " & plngJobID & " AND Type = 'A'"
  Case "GLOBAL DELETE"
    pstrSQL = "SELECT Name From ASRSysGlobalFunctions WHERE FunctionID = " & plngJobID & " AND Type = 'D'"
  Case "GLOBAL UPDATE"
    pstrSQL = "SELECT Name From ASRSysGlobalFunctions WHERE FunctionID = " & plngJobID & " AND Type = 'U'"
  Case "IMPORT"
    pstrSQL = "SELECT Name From ASRSysImportName WHERE ID = " & plngJobID
  Case "MAIL MERGE"
    pstrSQL = "SELECT Name From ASRSysMailMergeName WHERE MailMergeID = " & plngJobID
  Case "MATCH REPORT", "SUCCESSION PLANNING", "CAREER PROGRESSION"
    pstrSQL = "SELECT Name From ASRSysMatchReportName WHERE MatchReportID = " & plngJobID
  Case "RECORD PROFILE"
    pstrSQL = "SELECT Name From ASRSysRecordProfileName WHERE recordProfileID = " & plngJobID
  Case "ENVELOPES & LABELS"
    pstrSQL = "SELECT Name From ASRSysMailMergeName WHERE MailMergeID = " & plngJobID
  End Select
  
  'pstrSQL = pstrSQL & " AND (Access <> 'HD') OR (Username <> '" & gsUserName & "')"
  
  If Trim(pstrSQL) <> vbNullString Then
    Set prstTemp = datGeneral.GetReadOnlyRecords(pstrSQL)
    
    If Not prstTemp.BOF And Not prstTemp.EOF Then
      GetJobName = prstTemp.Fields(0)
      Set prstTemp = Nothing
      Exit Function
    End If
  End If
  
  GetJobName = ""
  Set prstTemp = Nothing

End Function

' Load the filename from a given stream
Public Function LoadFileNameFromStream(ByRef pobjStream As ADODB.Stream, pbFullPath As Boolean, pbDisplayOLEType As Boolean) As String

  Dim objTextStream As TextStream
  Dim strTempFileName As String
  Dim strProperties As String
  Dim objDocumentStream As New ADODB.Stream
  Dim objFileSystem As New FileSystemObject
  Dim strFileName As String
  Dim strPath As String
  Dim strUNC As String
  Dim iType As DataMgr.OLEType

  strTempFileName = GetTmpFName

  If Not pobjStream.State = adStateClosed Then
    If pobjStream.Size > 0 Then

      ' Setup new document stream
      objDocumentStream.Type = adTypeBinary
      objDocumentStream.Open
      
      ' Copy out the header part of the stream
      pobjStream.Position = 0
      pobjStream.CopyTo objDocumentStream, 400
      objDocumentStream.SaveToFile strTempFileName, adSaveCreateOverWrite
    
      ' Read in the header
      Set objTextStream = objFileSystem.OpenTextFile(strTempFileName, ForReading)
      strProperties = Trim(objTextStream.Read(400))

      iType = Val(Mid(strProperties, 9, 2))
      strFileName = Trim(Mid(strProperties, 11, 70))
      strPath = Trim(Mid(strProperties, 81, 210))
      strUNC = Trim(Mid(strProperties, 291, 60))
  
      objTextStream.Close
      objFileSystem.DeleteFile strTempFileName, True
      
      objDocumentStream.Close
      Set objDocumentStream = Nothing

      ' Append full path
      If pbFullPath And iType = OLE_UNC Then
        strFileName = strUNC & strPath & "\" & strFileName
      End If

      ' Append OLE Type
      If pbDisplayOLEType And Len(strFileName) > 0 Then
        If iType = OLE_EMBEDDED Then
          strFileName = strFileName & " (Embedded)"
        Else
          strFileName = strFileName & " (Linked)"
        End If
      End If

      LoadFileNameFromStream = strFileName

    End If
  End If

End Function

' Load a picture/photo from a given stream
Public Function LoadPictureFromStream(ByRef pobjStream As ADODB.Stream) As IPictureDisp

  On Error GoTo ErrorTrap

  'Dim objTextStream As TextStream
  Dim strTempFileName As String
  Dim objDocumentStream As New ADODB.Stream
  'Dim iOLEType As DataMgr.OLEType
  
  ' Save the document information to file to read in.
  strTempFileName = GetTmpFName

  If Not pobjStream.State = adStateClosed Then
    If pobjStream.Size > 0 Then
    
      ' Is stream a link or embedded
      If pobjStream.Size = 400 Then
      
        ' Linked photo
        strTempFileName = LoadFileNameFromStream(pobjStream, True, False)
        Set LoadPictureFromStream = LoadPicture(strTempFileName)
     
      Else
      
        ' Embedded - Setup new document stream
        objDocumentStream.Type = adTypeBinary
        objDocumentStream.Open
        
        ' Copy out the document part of the stream
        pobjStream.Position = 400
        pobjStream.CopyTo objDocumentStream, pobjStream.Size - 400
        objDocumentStream.SaveToFile strTempFileName, adSaveCreateOverWrite
      
        Set LoadPictureFromStream = LoadPicture(strTempFileName)
        
        ' Get rid of temporary file
        Kill strTempFileName
        
        objDocumentStream.Close
      
      End If
    End If
  Else
    
    ' No picture there
    Set LoadPictureFromStream = Nothing
  
  End If

TidyUpAndExit:
  Set objDocumentStream = Nothing
  Exit Function

ErrorTrap:
  Set LoadPictureFromStream = Nothing
  GoTo TidyUpAndExit

End Function

' Get the OLE Type froma stream (linked/embedded)
Public Function LoadOLETypeFromStream(ByRef pobjStream As ADODB.Stream) As DataMgr.OLEType

  Dim objTextStream As TextStream
  Dim strTempFileName As String
  Dim strProperties As String
  Dim objDocumentStream As New ADODB.Stream
  Dim objFileSystem As New FileSystemObject
  Dim strType As String

  strTempFileName = GetTmpFName

  If Not pobjStream.State = adStateClosed Then
    If pobjStream.Size > 0 Then

      ' Setup new document stream
      objDocumentStream.Type = adTypeBinary
      objDocumentStream.Open
      
      ' Copy out the header part of the stream
      pobjStream.Position = 0
      pobjStream.CopyTo objDocumentStream, 400
      objDocumentStream.SaveToFile strTempFileName, adSaveCreateOverWrite
    
      ' Read in the header
      Set objTextStream = objFileSystem.OpenTextFile(strTempFileName, ForReading)
      strProperties = Trim(objTextStream.Read(400))
    
      ' Set the OLE type
      strType = Val(Mid(strProperties, 9, 2))
 
      objTextStream.Close
      objFileSystem.DeleteFile strTempFileName, True
      
      objDocumentStream.Close
      Set objDocumentStream = Nothing

      LoadOLETypeFromStream = strType

    End If
  End If

End Function


Public Function DiaryFormat(dtInput As Date, strFormat As String) As String

  Dim lngIndex As Long
  
  DiaryFormat = Format(dtInput, strFormat)

  For lngIndex = 1 To Len(strFormat)
    If Mid(strFormat, lngIndex, 1) = ":" Then
      Mid(DiaryFormat, lngIndex, 1) = ":"
    End If
  Next

End Function

Public Function GetConnectionString(strParameter As String, strNewValue As String) As String

  'MH20010704 Will subsitute a parameter in the connection string for a new value
  'e.g. pass in "APP=" and a new value and function will return new connection string

  Dim strParamArray As Variant
  Dim strTemp As String
  Dim lngCount As Long
  
  'strTemp = gADOCon.ConnectionString
  strTemp = gsConnectionString
  'If InStr(strTemp, "Extended Properties=") > 0 Then
  If InStr(strTemp, strParameter) > 0 Then
    'strTemp = Mid(strTemp, InStr(strTemp, "Extended Properties="))
    'strTemp = Split(strTemp, Chr(34))(1)
    
    strParamArray = Split(strTemp, ";")
  
    For lngCount = 0 To UBound(strParamArray)
      
      If Left(strParamArray(lngCount), Len(strParameter)) = strParameter Then
        strParamArray(lngCount) = strParameter & strNewValue
        Exit For
      End If
    
    Next
    
    strTemp = Join(strParamArray, ";")
  
  End If

  GetConnectionString = strTemp

End Function

Public Function GetEmailAddress(lngEmailAddrID As Long, lngRecordID As Long) As String

  ' Return TRUE if the user has been granted the given permission.
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter

  On Error GoTo LocalErr
  
  ' Check if the user can create New instances of the given category.
  Set cmADO = New ADODB.Command
  With cmADO
    .CommandText = "dbo.spASRSysEmailAddr"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon

    Set pmADO = .CreateParameter("Result", adVarChar, adParamOutput, 8000)
    .Parameters.Append pmADO

    Set pmADO = .CreateParameter("EmailID", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.Value = lngEmailAddrID

    Set pmADO = .CreateParameter("RecordID", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.Value = lngRecordID

    cmADO.Execute

    GetEmailAddress = IIf(IsNull(.Parameters(0).Value), vbNullString, .Parameters(0).Value)
  End With
  Set cmADO = Nothing

Exit Function

LocalErr:
  COAMsgBox "Error reading email details" & vbCr & "(" & Err.Description & ")", vbExclamation
  Set cmADO = Nothing

End Function

Public Function EvaluateRecordDescription(plngRecordID As Long, plngRecDescID As Long) As String
  ' Return the evaluated record description for the given record.
  ' Used in frmRecEdit4, frmBulkbooking, frmAddFromWaitingList and frmRecordProfilePreview...so far..
  On Error GoTo ErrorTrap
  
  Dim sRecordDescription As String
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
    
  sRecordDescription = ""
  
  Set cmADO = New ADODB.Command
  With cmADO
    .CommandText = "sp_ASRExpr_" & Trim(Str(plngRecDescID))
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon
          
    Set pmADO = .CreateParameter("recordDescription", adVarChar, adParamOutput, VARCHAR_MAX_Size)
    .Parameters.Append pmADO
        
    Set pmADO = .CreateParameter("currentID", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.Value = plngRecordID
          
    Set pmADO = Nothing
        
    cmADO.Execute
        
    sRecordDescription = .Parameters("recordDescription").Value
  End With
  Set cmADO = Nothing
  
  EvaluateRecordDescription = sRecordDescription
  Exit Function
  
ErrorTrap:
  EvaluateRecordDescription = ""
  
End Function


Public Function COAMsgBox(sPrompt As String, Optional iButtons As VbMsgBoxStyle, Optional sTitle As String) As VbMsgBoxResult
  
  On Local Error GoTo LocalErr
  
  If sTitle = vbNullString Then sTitle = app.ProductName
  gobjProgress.Visible = False
  
  DebugOutput "modHRPro.COAMsgBox", Replace(sPrompt, vbCrLf, " ")
  If Not gblnBatchJobsOnly Then
    COAMsgBox = MsgBox(sPrompt, iButtons, sTitle)
  Else
    Open app.Path & "\batcherr.txt" For Append As #1
    Print #1, Format(Now, DateFormat & " hh:nn")
    Print #1, sPrompt
    Print #1, ""
    Close #1
  End If

Exit Function

LocalErr:
  Close #1

End Function


Public Function IsFileCompatibleWithWordVersion(strFileName As String)
  IsFileCompatibleWithWordVersion = (GetOfficeSaveAsFormat(strFileName, GetOfficeWordVersion, oaWord) <> "")
End Function

Public Function IsFileCompatibleWithExcelVersion(strFileName As String)
  IsFileCompatibleWithExcelVersion = (GetOfficeSaveAsFormat(strFileName, GetOfficeExcelVersion, oaExcel) <> "")
End Function


Public Function GetOfficeSaveAsFormat(strFileName As String, intOfficeVersion As Integer, app As OfficeApp) As String
  
  Dim rsTemp As ADODB.Recordset
  Dim sSQL As String
  Dim strExtension As String
  
  On Local Error GoTo LocalErr
  
  GetOfficeSaveAsFormat = ""

  If intOfficeVersion > 0 And InStr(strFileName, ".") Then
    strExtension = Mid(strFileName, InStrRev(strFileName, ".") + 1)
  
    sSQL = "SELECT " & IIf(intOfficeVersion < 12, "Office2003", "Office2007") & _
           " FROM ASRSysFileFormats WHERE Extension = '" & strExtension & "' AND Destination LIKE '" & IIf(app = oaWord, "WORD", "EXCEL") & "%'" & _
           " ORDER BY ID"
    Set rsTemp = datGeneral.GetReadOnlyRecords(sSQL)
    GetOfficeSaveAsFormat = IIf(IsNull(rsTemp(0).Value), "", rsTemp(0).Value)
    rsTemp.Close
    Set rsTemp = Nothing
  End If

Exit Function

LocalErr:
  GetOfficeSaveAsFormat = ""

End Function


Public Function GetOfficeWordVersion() As Integer

  Dim app As Word.Application
  
  On Error GoTo NotInstalled

  If giOfficeVersion_Word = 0 Then
    giOfficeVersion_Word = -1
    If InStr(LCase(Command$), "/msoffice=false") = 0 Then
      Set app = CreateObject("Word.Application")
      giOfficeVersion_Word = Val(app.Version)
      app.Quit
    End If
  End If

TidyUpAndExit:
  GetOfficeWordVersion = giOfficeVersion_Word
  Set app = Nothing
  
Exit Function

NotInstalled:
  Resume TidyUpAndExit

End Function


Public Function GetOfficeExcelVersion() As Integer

  Dim app As Excel.Application
  
  On Error GoTo NotInstalled

  If giOfficeVersion_Excel = 0 Then
    giOfficeVersion_Excel = -1
    If InStr(LCase(Command$), "/msoffice=false") = 0 Then
      Set app = CreateObject("Excel.Application")
      giOfficeVersion_Excel = Val(app.Version)
      app.Quit
    End If
  End If

TidyUpAndExit:
  GetOfficeExcelVersion = giOfficeVersion_Excel
  Set app = Nothing

Exit Function

NotInstalled:
  Resume TidyUpAndExit

End Function

Public Function InitialiseCommonDialogFormats(cd1 As CommonDialog, strDestin As String, intOfficeVersion As Integer, Direction As FileFormatDirection) As Boolean

  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim strFormatField As String
  Dim intUserDefault As Integer
  Dim intCount As Integer
  Dim blnResult As Boolean
  
  Dim strFilter As String
  Dim intFilterIndex As Integer
  
  On Local Error GoTo LocalErr
  
  blnResult = False
    
  strFormatField = IIf(intOfficeVersion < 12, "Office2003", "Office2007")
  
  strSQL = "SELECT * " & _
           "FROM   ASRSysFileFormats " & _
           "WHERE  Destination = '" & Replace(strDestin, "'", "''") & "' " & _
           "  AND  NOT " & strFormatField & " IS NULL " & _
           IIf(Direction = DirectionInput, " AND [direction] IN (0,2)", " AND [direction] IN (1,2)") & _
           " ORDER BY ID"
  Set rsTemp = datGeneral.GetRecords(strSQL)

  intUserDefault = GetUserSetting("Output", strDestin & "Format", 0)
  
  strFilter = vbNullString
  intFilterIndex = 0
  intCount = 1
  Do While Not rsTemp.EOF
    
    strFilter = strFilter & _
      IIf(strFilter <> vbNullString, "|", "") & _
      rsTemp.Fields("Description").Value & "|*." & rsTemp.Fields("Extension").Value
    
    If intUserDefault = rsTemp.Fields(strFormatField).Value Then
      intFilterIndex = intCount
    ElseIf rsTemp.Fields("Default").Value = True Then
      If intFilterIndex = 0 Then
        intFilterIndex = intCount
      End If
    End If
    
    intCount = intCount + 1
    rsTemp.MoveNext
  Loop

  cd1.Filter = strFilter
  cd1.FilterIndex = intFilterIndex

  blnResult = True

LocalErr:
  If Not rsTemp Is Nothing Then
    If rsTemp.State <> adStateClosed Then
      rsTemp.Close
    End If
    Set rsTemp = Nothing
  End If
    
  InitialiseCommonDialogFormats = blnResult

End Function


Public Sub DebugOutput(strWhere As String, strWhat As String)
  
  Dim lngFile As Long
  
  'Debug.Print CStr(Now) & "  " & Left(strWhere & Space(32), 32) & strWhat
  If gstrDebugOutputFile <> vbNullString Then
    On Local Error Resume Next
    lngFile = FreeFile
    Open gstrDebugOutputFile For Append As #lngFile
    Print #lngFile, CStr(Now) & "  " & Left(strWhere & Space(32), 32) & strWhat
    Close #lngFile
  End If

End Sub

Public Sub UpdateUsage(ByRef lngTYPE As utilityType, ByRef lngUtilityID As Long, lngAction As EditOptions)

  Dim cmdUsage As New ADODB.Command
  Dim pmADO As ADODB.Parameter

  If lngAction <> edtPrint Then

    Set cmdUsage = New ADODB.Command
    With cmdUsage
      .CommandText = "dbo.spstat_updateobjectusage"
      .CommandType = adCmdStoredProc
      .CommandTimeout = 0
      Set .ActiveConnection = gADOCon
  
      Set pmADO = .CreateParameter("objecttype", adInteger, adParamInput, 50)
      .Parameters.Append pmADO
      pmADO.Value = lngTYPE
  
      Set pmADO = .CreateParameter("objectid", adInteger, adParamInput, 50)
      .Parameters.Append pmADO
      pmADO.Value = lngUtilityID
  
      Set pmADO = .CreateParameter("lastaction", adInteger, adParamInput, 50)
      .Parameters.Append pmADO
      pmADO.Value = lngAction
  
      .Execute
  
    End With
    
  End If
  
  Set cmdUsage = Nothing

End Sub

Public Function SaveObjectCategories(ByRef theCombo As ComboBox, utilityType As utilityType, UtilityID As Long) As Boolean

  On Error GoTo ErrorTrap

  Dim bOK As Boolean
  Dim iLoop As Integer
  Dim iSelectedID As Integer
  
  bOK = True
  iSelectedID = GetComboItem(theCombo)
  
  gobjDataAccess.ExecuteSql "EXEC dbo.spsys_saveobjectcategories " & CStr(utilityType) & ", " & CStr(UtilityID) & ", " & CStr(iSelectedID)
  
TidyUpAndExit:
  SaveObjectCategories = bOK
  Exit Function
  
ErrorTrap:
bOK = False
  GoTo TidyUpAndExit

End Function

Public Function GetObjectCategory(utilityType As utilityType, UtilityID As Long) As String

  On Error GoTo ErrorTrap

  Dim rsTemp As ADODB.Recordset
  
  GetObjectCategory = "<None>"

  Set rsTemp = gobjDataAccess.OpenRecordset("EXEC dbo.spsys_getobjectcategories " & CStr(utilityType) & ", " & CStr(UtilityID) & ", 0" _
      , adOpenForwardOnly, adLockReadOnly)
  If Not rsTemp.BOF And Not rsTemp.EOF Then
    rsTemp.MoveFirst
    
    ' This loop may be a bit ineffcient, but it is proposed at some point to have multiple categories, so that's why it loops!
    Do While Not rsTemp.EOF
      If rsTemp.Fields("Selected").Value = 1 Then
        GetObjectCategory = rsTemp.Fields("category_name").Value
      End If
    rsTemp.MoveNext
    Loop
  End If

TidyUpAndExit:
  Set rsTemp = Nothing
  Exit Function
  
ErrorTrap:
  GoTo TidyUpAndExit

End Function

Public Sub GetObjectCategories(ByRef theCombo As ComboBox, utilityType As utilityType, UtilityID As Long, Optional TableID As Long)

  On Error GoTo ErrorTrap

  Dim rsTemp As ADODB.Recordset
  Dim iListIndex As Integer
  Dim bFound As Boolean
  
  bFound = False
  
  ' Add <none>
  theCombo.AddItem "<None>"
  theCombo.ItemData(theCombo.NewIndex) = 0
  iListIndex = theCombo.NewIndex
     
  Set rsTemp = gobjDataAccess.OpenRecordset("EXEC dbo.spsys_getobjectcategories " & CStr(utilityType) & ", " & CStr(UtilityID) & ", " & CStr(TableID) _
      , adOpenForwardOnly, adLockReadOnly)
  
  If Not rsTemp.BOF And Not rsTemp.EOF Then
    rsTemp.MoveFirst
    Do While Not rsTemp.EOF
      theCombo.AddItem rsTemp.Fields("category_name").Value
      theCombo.ItemData(theCombo.NewIndex) = rsTemp.Fields("ID").Value
      
      If rsTemp.Fields("Selected").Value = 1 Then
        iListIndex = theCombo.NewIndex
        bFound = True
      End If
      rsTemp.MoveNext
    Loop
  End If
  
  theCombo.Enabled = (theCombo.ListCount > 0)
    
  If iListIndex > -1 And UtilityID > 0 Then
    theCombo.ListIndex = iListIndex
  ElseIf TableID > -1 Then
    SetComboItem theCombo, TableID
  End If
  
TidyUpAndExit:
  Set rsTemp = Nothing
  Exit Sub
  
ErrorTrap:
  GoTo TidyUpAndExit

End Sub


Public Sub GetObjectOwners(ByRef theCombo As ComboBox, utilityType As String)

  On Error GoTo ErrorTrap

  Dim rsTemp As ADODB.Recordset
  Dim iListIndex As Integer
  Dim lngCount As Long
       
  theCombo.Clear
       
  ' Add <All>
  theCombo.AddItem "<All>"
  theCombo.ItemData(theCombo.NewIndex) = 0
  
  ' Add <Mine>
  theCombo.AddItem "<Mine> (" + StrConv(gsUserName, vbProperCase) + ")"
  theCombo.ItemData(theCombo.NewIndex) = 1
  
  iListIndex = theCombo.NewIndex
  lngCount = 2
       
  Set rsTemp = gobjDataAccess.OpenRecordset("SELECT DISTINCT username FROM ASRSysAllObjectNames WHERE NOT NULLIF(username,'') = '' AND username <> '" & gsUserName & "' ORDER BY username" _
      , adOpenForwardOnly, adLockReadOnly)
  
  If Not rsTemp.BOF And Not rsTemp.EOF Then
    rsTemp.MoveFirst
    Do While Not rsTemp.EOF
      theCombo.AddItem rsTemp.Fields("username").Value
      theCombo.ItemData(theCombo.NewIndex) = lngCount
      
      rsTemp.MoveNext
      lngCount = lngCount + 1
    Loop
  End If
  
  theCombo.Enabled = (theCombo.ListCount > 0)
    
  If GetUserSetting("DefSel", "OnlyMine " & utilityType, 0) = 1 Then
    SetComboItem theCombo, 1
  Else
    SetComboItem theCombo, 0
  End If
  
TidyUpAndExit:
  Set rsTemp = Nothing
  Exit Sub
  
ErrorTrap:
  GoTo TidyUpAndExit

End Sub

' Converts utilityID to text format because the batch jobs store it in a testual format (Why? For god's sake why???)
Public Function GetBatchJobType(ByVal Utility As utilityType) As String

  Select Case Utility
  
    Case utlBatchJob
      GetBatchJobType = "Batch Job"
  
    Case utlCalendarReport
      GetBatchJobType = "Calendar Report"
  
    Case utlCalculation
      GetBatchJobType = "Calculation"
  
    Case utlCrossTab
      GetBatchJobType = "Cross Tab"
  
    Case utlCustomReport
      GetBatchJobType = "Custom Report"
  
    Case utlDataTransfer
      GetBatchJobType = "Data Transfer"
  
    Case utlEmailAddress
      GetBatchJobType = "Email Address"
  
    Case utlEmailGroup
      GetBatchJobType = "Email Group"
  
    Case utlExport
      GetBatchJobType = "Export"
  
    Case utlFilter
      GetBatchJobType = "Filter"
      
    Case UtlGlobalAdd
      GetBatchJobType = "Global Add"
    
    Case utlGlobalDelete
      GetBatchJobType = "Global Delete"
    
    Case utlGlobalUpdate
      GetBatchJobType = "Global Update"
     
    Case utlImport
      GetBatchJobType = "Import"
  
    Case utlMatchReport
      GetBatchJobType = "Match Report"
    
    Case utlSuccession
      GetBatchJobType = "Succession Planning"
    
    Case utlCareer
      GetBatchJobType = "Career Progression"
  
    Case utlMailMerge
      GetBatchJobType = "Mail Merge"
      
    Case utlLabel
      GetBatchJobType = "Envelopes & Labels"
    
    Case utlLabelType
      GetBatchJobType = "Envelope & Label Template"
  
    Case utlDocumentMapping
      GetBatchJobType = "Document Type"
  
    Case utlPicklist
      GetBatchJobType = "Picklist"
    
    Case utlRecordProfile
      GetBatchJobType = "Record Profile"
      
    Case utlWorkflow
      GetBatchJobType = "Workflow"
      
    Case utlReportPack
      GetBatchJobType = "Report Pack"
      
    Case Else
      GetBatchJobType = ""

  End Select


End Function

Public Function ContainsInvalidXML(ByVal NodeText, AllowForwardSlash As Boolean) As Boolean

  Dim bOK As Boolean

  bOK = True
  If InStr(1, NodeText, " ") > 0 Then bOK = False
  If InStr(1, NodeText, ":") > 0 Then bOK = False
  If InStr(1, NodeText, "~") > 0 Then bOK = False
  If InStr(1, NodeText, "\") > 0 Then bOK = False
  If InStr(1, NodeText, "/") > 0 Then bOK = AllowForwardSlash
  If InStr(1, NodeText, ";") > 0 Then bOK = False
  If InStr(1, NodeText, "?") > 0 Then bOK = False
  If InStr(1, NodeText, "$") > 0 Then bOK = False
  If InStr(1, NodeText, "&") > 0 Then bOK = False
  If InStr(1, NodeText, "%") > 0 Then bOK = False
  If InStr(1, NodeText, "@") > 0 Then bOK = False
  If InStr(1, NodeText, "^") > 0 Then bOK = False
  If InStr(1, NodeText, "=") > 0 Then bOK = False
  If InStr(1, NodeText, "*") > 0 Then bOK = False
  If InStr(1, NodeText, "+") > 0 Then bOK = False
  If InStr(1, NodeText, "(") > 0 Then bOK = False
  If InStr(1, NodeText, Chr(34)) > 0 Then bOK = False
  If InStr(1, NodeText, ")") > 0 Then bOK = False
  If InStr(1, NodeText, "|") > 0 Then bOK = False
  If InStr(1, NodeText, "'") > 0 Then bOK = False
  If InStr(1, NodeText, "`") > 0 Then bOK = False
  If InStr(1, NodeText, "{") > 0 Then bOK = False
  If InStr(1, NodeText, "}") > 0 Then bOK = False
  If InStr(1, NodeText, "[") > 0 Then bOK = False
  If InStr(1, NodeText, "]") > 0 Then bOK = False
  If InStr(1, NodeText, "<") > 0 Then bOK = False
  If InStr(1, NodeText, ">") > 0 Then bOK = False

  ContainsInvalidXML = Not bOK
  
End Function

Public Function IsModuleEnabled(lngModuleCode As Module) As Boolean
  IsModuleEnabled = (gobjLicence.Modules And lngModuleCode)
End Function
