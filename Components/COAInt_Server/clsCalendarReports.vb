Option Strict Off
Option Explicit On

Imports System.Globalization
Imports VB = Microsoft.VisualBasic
Public Class clsCalendarReportsRUN

  Private mstrSQLSelect_RegInfoRegion As String
  Private mstrSQLSelect_BankHolDate As String
  Private mstrSQLSelect_BankHolDesc As String

  Private mstrSQLSelect_PersonnelStaticRegion As String
  Private mstrSQLSelect_PersonnelHRegion As String
  Private mstrSQLSelect_PersonnelHDate As String
  Private mstrSQLSelect_PersonnelStaticWP As String

  Private mstrBaseTableName As String

  Private mvarTableViews(,) As Object
  Private mstrTempRealSource As String

  'TableViews
  Private mstrRealSource As String
  Private mstrBaseTableRealSource As String
  Private mlngTableViews(,) As Integer
  Private mstrViews() As String
  Private mobjTableView As CTablePrivilege
  Private mobjColumnPrivileges As CColumnPrivileges

  Private mvarEventColumnViews(,) As Object
  Private mlngEventViewColumn As Integer

  '************************************************************

  Private mlngCalendarReportID As Integer
  Private mstrErrorString As String
  Private mstrTempTableName As String
  Private mblnUDFsCreated As Boolean
  Private mblnTempTableCreated As Boolean
  Private mblnOrderByCreated As Boolean
  Private mlngBaseDescriptionType As Integer

  Private mblnRegions As Boolean
  Private mblnWorkingPatterns As Boolean
  Private mblnStaticReg As Boolean
  Private mblnStaticWP As Boolean

  'Variables to store definition (report level variables)
  Private mstrCalendarReportsName As String
  Private mlngCalendarReportsBaseTable As Integer
  Private mstrCalendarReportsBaseTableName As String
  Private mlngCalendarReportsAllRecords As Integer
  Private mlngCalendarReportsPickListID As Integer
  Private mlngCalendarReportsFilterID As Integer
  Private mlngDescription1 As Integer
  Private mstrDescription1 As String
  Private mblnDesc1IsDate As Boolean
  Private mlngDescription2 As Integer
  Private mstrDescription2 As String
  Private mblnDesc2IsDate As Boolean
  Private mlngDescriptionExpr As Integer
  Private mstrDescriptionExpr As String
  Private mblnDescExprIsDate As Boolean

  Private mstrDescriptionSeparator As String

  Private mstrBaseIDColumn As String
  Private mstrEventIDColumn As String

  Private mblnDescCalcCode As Boolean
  Private mstrDescCalcCode As String

  Private mlngRegion As Integer
  Private mstrRegion As String
  Private mstrRegionColumnRealSource As String
  Private mblnGroupByDescription As Boolean

  Private mlngStartDateExpr As Integer
  Private mstrStartDate As String
  Private mdtStartDate As Date
  Private mlngEndDateExpr As Integer
  Private mstrEndDate As String
  Private mdtEndDate As String

  Private mblnShowBankHolidays As Boolean
  Private mblnShowCaptions As Boolean
  Private mblnShowWeekends As Boolean
  Private mbStartOnCurrentMonth As Boolean
  Private mblnIncludeWorkingDaysOnly As Boolean
  Private mblnIncludeBankHolidays As Boolean
  Private mblnCustomReportsPrintFilterHeader As Boolean
  Private mstrFilteredIDs As String

  'New Default Output Variables
  Private mblnOutputPreview As Boolean
  Private mlngOutputFormat As Integer
  Private mblnOutputScreen As Boolean
  Private mblnOutputPrinter As Boolean
  Private mstrOutputPrinterName As String
  Private mblnOutputSave As Boolean
  Private mlngOutputSaveExisting As Integer
  Private mblnOutputEmail As Boolean
  Private mlngOutputEmailID As Integer
  Private mstrOutputEmailName As String
  Private mstrOutputEmailSubject As String
  Private mstrOutputEmailAttachAs As String
  Private mstrOutputFilename As String

  'Recordset to store the final data from the temp table
  Private mrsCalendarReportsOutput As ADODB.Recordset
  Private mrsCalendarBaseInfo As ADODB.Recordset

  'Array to store data for each session label object,
  'array also holds information on working days and bank holidays etc.
  Private mvarDateLabelInfo() As Object

  Private mstrClientDateFormat As String
  Private mstrLocalDecimalSeparator As String
  Private mlngColumnLimit As Integer

  'Strings to hold the SQL statement
  Private mstrSQLEvent As String
  Private mstrSQLSelect As String
  Private mstrSQLFrom As String
  Private mstrSQLJoin As String
  Private mstrSQLWhere As String
  Private mstrSQLOrderBy As String
  Private mstrSQL As String
  Private mstrSQLBaseData As String
  Private mstrSQLBaseDateClause As String
  Private mstrSQLOrderList As String
  Private mstrSQLIDs As String
  Private mstrSQLDynamicLegendWhere As String
  Private mintDynamicEventCount As Short
  Private mstrSQLCreateTable As String

  Private mblnHasEventFilterIDs As Boolean
  Private mstrEventFilterIDs As String

  'used to temporarily store the Base table Start & End date table.columnname for the
  'current event. Then used when creating the mstrSQLBaseDateClause.
  Private mstrSQLBaseStartDateColumn As String
  Private mstrSQLBaseStartSessionColumn As String
  Private mstrSQLBaseEndDateColumn As String
  Private mstrSQLBaseEndSessionColumn As String
  Private mstrSQLBaseDurationColumn As String

  ' Classes
  Private mclsData As clsDataAccess
  Private mclsGeneral As clsGeneral
  Private mclsUI As clsUI
  Private mobjEventLog As clsEventLog

  'Array holding the columns to sort the report by
  Private mvarSortOrder(,) As Object
  Private mvarPrompts(,) As Object

  Private mcolEvents As clsCalendarEvents

  'Instance of the previewform
  'Private mfrmOutput As frmCalendarReportPreview


  'Batch Job Mode ?
  Private mblnBatchMode As Boolean

  'Has the user cancelled the report ?
  Private mblnUserCancelled As Boolean

  'Does the report generate no records ?
  Private mblnNoRecords As Boolean

  'Is the current user the definition owner ?
  Private mblnDefinitionOwner As Boolean

  'Runnning report for single record only!
  Private mlngSingleRecordID As Integer

  ' Array holding the User Defined functions that are needed for this report
  Private mastrUDFsRequired() As String

  Private mcolStaticBankHolidays As Collection
  Private mcolHistoricBankHolidays As Collection
  Private mcolStaticWorkingPatterns As Collection
  Private mcolHistoricWorkingPatterns As Collection

  Private mcolBaseDescIndex As Collection
  'Private mcolDateControlEvents As Collection

  Private mblnPersonnelBase As Boolean

  Private mstrRegionFormString As String
  Private mstrBHolFormString As String
  Private mstrWPFormString As String

  '****************************************************
  'variables for outputting
  Private mavOutputDateIndex(,) As Object

  Private mintFirstDayOfMonth_Output As Short
  Private mintDaysInMonth_Output As Short

  Private mintRangeStartIndex_Output As Short
  Private mintRangeEndIndex_Output As Short

  Private mdtVisibleStartDate_Output As Date
  Private mdtVisibleEndDate_Output As Date

  Private mdtEventStartDate_Output As Date
  Private mstrEventStartSession_Output As String
  Private mdtEventEndDate_Output As Date
  Private mstrEventEndSession_Output As String
  Private mstrDuration_Output As String
  Private mstrEventLegend_Output As String

  Private mlngMonth_Output As Integer
  Private mlngYear_Output As Integer

  Private mintCurrentBaseIndex_Output As Short
  Private mintBaseCount_Output As Short
  Private mstrBaseRecDesc_Output As String
  Private mintBaseRecordCount_Output As Short

  Private mcolBaseDescIndex_Output As Collection

  Private mlngGridRowIndex As Integer
  '****************************************************

  Private mvarOutputArray_Definition() As Object
  Private mvarOutputArray_Columns() As Object
  Private mvarOutputArray_Data() As Object
  Private mvarOutputArray_Styles() As Object
  Private mvarOutputArray_Merges() As Object

  Private mavLegend(,) As Object
  Private mstrLegend() As String
  Private mavAvailableColours(,) As Object
  Private mstrExcludedColours As String

  '****************************************************
  'variables for checking for multiple events
  Private mavLegendDateIndex() As Object

  Private mintFirstDayOfMonth_Legend As Short
  Private mintDaysInMonth_Legend As Short

  Private mintRangeStartIndex_Legend As Short
  Private mintRangeEndIndex_Legend As Short

  Private mdtVisibleStartDate_Legend As Date
  Private mdtVisibleEndDate_Legend As Date

  Private mdtEventStartDate_Legend As Date
  Private mstrEventStartSession_Legend As String
  Private mdtEventEndDate_Legend As Date
  Private mstrEventEndSession_Legend As String
  Private mstrDuration_Legend As String
  Private mstrEventLegend_Legend As String

  Private mlngMonth_Legend As Integer
  Private mlngYear_Legend As Integer

  Private mintCurrentBaseIndex_Legend As Short
  Private mintBaseCount_Legend As Short
  Private mstrBaseRecDesc_Legend As String
  Private mintBaseRecordCount_Legend As Short

  Private mcolBaseDescIndex_Legend As Collection

  Private mblnHasMultipleEvents As Boolean
  '****************************************************

  Private mblnDisableRegions As Boolean
  Private mblnDisableWPs As Boolean

  Private mstrCurrentEventKey As String
  Private mstrBaseRecDesc As String
  Private mlngCurrentRecordID As Integer
  Private mstrCurrentBaseRegion As String

  Private Const CALREP_DATEFORMAT As String = "dd/mm/yyyy"

  'default output colours
  Private mlngBC_Data As Integer
  Private mlngFC_Data As Integer
  Private mlngBC_Heading As Integer
  Private mlngFC_Heading As Integer
  Private mlngColor_Weekend As Integer
  Private mlngColor_Disabled As Integer
  Private mlngColor_RangeDisabled As Integer

  Private mavCareerRanges(,) As Object

  Private Const DAY_CONTROL_COUNT As Short = 37

  Private mintLegendCount As Short

  Private mintType_BaseDesc1 As Short
  Private mintType_BaseDesc2 As Short
  Private mintType_BaseDescExpr As Short
  Private mstrFormat_BaseDesc1 As String
  Private mstrFormat_BaseDesc2 As String

  Private mlngMergePageArrayIndex As Integer
  Private mlngStylePageArrayIndex As Integer

  'TM01042004 Fault 8428
  Private mblnCheckingRegionColumn As Boolean
  Private mblnCheckingDescColumn As Boolean

  Private Function SQLDateConvertToLocale(ByRef pstrTableColumn As String) As String

    'Takes the Column value and Returns a string with the SQL Code to format the
    'SQL date value into the known locale.

    Dim strSQL As String
    Dim strDateFormat As String

    Dim blnDateComplete As Boolean
    Dim blnMonthDone As Boolean
    Dim blnDayDone As Boolean
    Dim blnYearDone As Boolean

    Dim strShortDate As String

    Dim strDateSeparator As String

    Dim i As Short

    ' eg. DateFormat = "mm/dd/yyyy"
    '     Calendar   = "dd/mm/yyyy"
    '     DateString = "06/02/2000"
    '     Compare to = 02/06/2000

    strDateFormat = mstrClientDateFormat

    strDateSeparator = mstrLocalDecimalSeparator

    blnDateComplete = False
    blnMonthDone = False
    blnDayDone = False
    blnYearDone = False

    strShortDate = vbNullString

    For i = 1 To Len(strDateFormat) Step 1

      If (LCase(Mid(strDateFormat, i, 1)) = "d") And (Not blnDayDone) Then
        strShortDate = strShortDate & LCase(Mid(strDateFormat, i, 1))
        blnDayDone = True
      End If

      If (LCase(Mid(strDateFormat, i, 1)) = "m") And (Not blnMonthDone) Then
        strShortDate = strShortDate & LCase(Mid(strDateFormat, i, 1))
        blnMonthDone = True
      End If

      If (LCase(Mid(strDateFormat, i, 1)) = "y") And (Not blnYearDone) Then
        strShortDate = strShortDate & LCase(Mid(strDateFormat, i, 1))
        blnYearDone = True
      End If

      If blnDayDone And blnMonthDone And blnYearDone Then
        blnDateComplete = True
        Exit For
      End If

    Next i

    Select Case strShortDate
      Case "dmy"
        SQLDateConvertToLocale = "CONVERT(varchar, DATEPART(d," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & "CONVERT(varchar, DATEPART(m," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & "CONVERT(varchar, DATEPART(yyyy," & pstrTableColumn & ")) "
      Case "mdy"
        SQLDateConvertToLocale = "CONVERT(varchar, DATEPART(m," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & "CONVERT(varchar, DATEPART(d," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & "CONVERT(varchar, DATEPART(yyyy," & pstrTableColumn & ")) "
      Case "ydm"
        SQLDateConvertToLocale = "CONVERT(varchar, DATEPART(yyyy," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & "CONVERT(varchar, DATEPART(d," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & "CONVERT(varchar, DATEPART(m," & pstrTableColumn & ")) "
      Case "myd"
        SQLDateConvertToLocale = "CONVERT(varchar, DATEPART(m," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & "CONVERT(varchar, DATEPART(yyyy," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & "CONVERT(varchar, DATEPART(d," & pstrTableColumn & ")) "
      Case "ymd"
        SQLDateConvertToLocale = "CONVERT(varchar, DATEPART(yyyy," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & "CONVERT(varchar, DATEPART(m," & pstrTableColumn & ")) + '" & strDateSeparator & "' + " & "CONVERT(varchar, DATEPART(d," & pstrTableColumn & ")) "
    End Select


  End Function

  Public Property EventLogID() As Integer
    Get
      EventLogID = mobjEventLog.EventLogID
    End Get
    Set(ByVal Value As Integer)
      mobjEventLog.EventLogID = Value
    End Set
  End Property

  Public WriteOnly Property Cancelled() As Boolean
    Set(ByVal Value As Boolean)

      ' Connection object passed in from the asp page
      If Value = True Then
        mobjEventLog.ChangeHeaderStatus(clsEventLog.EventLog_Status.elsCancelled)
      Else
        mobjEventLog.ChangeHeaderStatus(clsEventLog.EventLog_Status.elsSuccessful)
      End If

    End Set
  End Property

  Public WriteOnly Property Failed() As Boolean
    Set(ByVal Value As Boolean)

      ' Connection object passed in from the asp page
      If Value = True Then
        mobjEventLog.ChangeHeaderStatus(clsEventLog.EventLog_Status.elsFailed)
      End If

    End Set
  End Property

  Public WriteOnly Property FailedMessage() As String
    Set(ByVal Value As String)
      mobjEventLog.AddDetailEntry(Value)
    End Set
  End Property


  Public ReadOnly Property NoRecords() As Boolean
    Get
      ' Does the report have any records ?
      NoRecords = mblnNoRecords
    End Get
  End Property

  Public ReadOnly Property OutputPreview() As Boolean
    Get
      OutputPreview = mblnOutputPreview
    End Get
  End Property

  Public ReadOnly Property OutputFormat() As Integer
    Get
      OutputFormat = mlngOutputFormat
    End Get
  End Property

  Public ReadOnly Property OutputScreen() As Boolean
    Get
      OutputScreen = mblnOutputScreen
    End Get
  End Property

  Public ReadOnly Property OutputPrinter() As Boolean
    Get
      OutputPrinter = mblnOutputPrinter
    End Get
  End Property

  Public ReadOnly Property OutputPrinterName() As String
    Get
      OutputPrinterName = mstrOutputPrinterName
    End Get
  End Property

  Public ReadOnly Property OutputSave() As Boolean
    Get
      OutputSave = mblnOutputSave
    End Get
  End Property

  Public ReadOnly Property OutputSaveExisting() As Integer
    Get
      OutputSaveExisting = mlngOutputSaveExisting
    End Get
  End Property

  Public ReadOnly Property OutputEmail() As Boolean
    Get
      OutputEmail = mblnOutputEmail
    End Get
  End Property

  Public ReadOnly Property OutputEmailID() As Integer
    Get
      OutputEmailID = mlngOutputEmailID
    End Get
  End Property

  Public ReadOnly Property OutputEmailGroupName() As String
    Get
      OutputEmailGroupName = mstrOutputEmailName
    End Get
  End Property

  Public ReadOnly Property OutputEmailSubject() As String
    Get
      OutputEmailSubject = mstrOutputEmailSubject
    End Get
  End Property

  Public ReadOnly Property OutputEmailAttachAs() As String
    Get
      OutputEmailAttachAs = mstrOutputEmailAttachAs
    End Get
  End Property

  Public ReadOnly Property OutputFilename() As String
    Get
      OutputFilename = mstrOutputFilename
    End Get
  End Property




  Public ReadOnly Property HasMultipleEvents() As Boolean
    Get
      HasMultipleEvents = mblnHasMultipleEvents
    End Get
  End Property

  Public ReadOnly Property IncludeBankHolidays_Enabled() As Boolean
    Get
      If (Not mblnGroupByDescription) And (Not mblnDisableRegions) And ((PersonnelBase And (Len(Trim(gsPersonnelRegionColumnName)) > 0) And (glngBHolRegionID > 0)) Or (PersonnelBase And (Len(Trim(gsPersonnelHRegionColumnName)) > 0) And (glngBHolRegionID > 0)) Or (mlngRegion > 0)) Then

        IncludeBankHolidays_Enabled = True
      Else
        IncludeBankHolidays_Enabled = False
      End If
    End Get
  End Property

  Public ReadOnly Property IncludeWorkingDaysOnly_Enabled() As Boolean
    Get
      If (Not mblnGroupByDescription) And (Not mblnDisableWPs) And ((PersonnelBase And (Len(Trim(gsPersonnelWorkingPatternColumnName)) > 0)) Or (PersonnelBase And (Len(Trim(gsPersonnelHWorkingPatternColumnName)) > 0))) Then

        IncludeWorkingDaysOnly_Enabled = True
      Else
        IncludeWorkingDaysOnly_Enabled = False
      End If
    End Get
  End Property





  Public WriteOnly Property LocalDecimalSeparator() As String
    Set(ByVal Value As String)

      ' Clients date format passed in from the asp page
      mstrLocalDecimalSeparator = Value

    End Set
  End Property

  Public WriteOnly Property ClientDateFormat() As String
    Set(ByVal Value As String)

      ' Clients date format passed in from the asp page
      mstrClientDateFormat = Value

    End Set
  End Property

  Public ReadOnly Property ShowBankHolidays_Enabled() As Boolean
    Get
      If (Not mblnGroupByDescription) And (Not mblnDisableRegions) And ((PersonnelBase And (Len(Trim(gsPersonnelRegionColumnName)) > 0) And (glngBHolRegionID > 0)) Or (PersonnelBase And (Len(Trim(gsPersonnelHRegionColumnName)) > 0) And (glngBHolRegionID > 0)) Or (mlngRegion > 0)) Then
        ShowBankHolidays_Enabled = True
      Else
        ShowBankHolidays_Enabled = False
      End If
    End Get
  End Property

  Public ReadOnly Property SQLCalendarBaseInfo() As String
    Get
      SQLCalendarBaseInfo = mstrSQLBaseData
    End Get
  End Property

  Public ReadOnly Property SQLOutput() As String
    Get
      SQLOutput = mstrSQL
    End Get
  End Property
  Public ReadOnly Property StaticWP() As Boolean
    Get
      StaticWP = mblnStaticWP
    End Get
  End Property
  Public ReadOnly Property StaticReg() As Boolean
    Get
      StaticReg = mblnStaticReg
    End Get
  End Property

  Public ReadOnly Property UserCancelled() As Boolean
    Get
      UserCancelled = mblnUserCancelled
    End Get
  End Property
  Public Property CalendarReportID() As Integer
    Get
      CalendarReportID = mlngCalendarReportID
    End Get
    Set(ByVal Value As Integer)
      mlngCalendarReportID = Value
    End Set
  End Property
  Public WriteOnly Property SingleRecordID() As Integer
    Set(ByVal Value As Integer)
      mlngSingleRecordID = Value
    End Set
  End Property
  Public Property ErrorString() As String
    Get
      ErrorString = mstrErrorString
    End Get
    Set(ByVal Value As String)
      mstrErrorString = Value
    End Set
  End Property

  Public WriteOnly Property Username() As String
    Set(ByVal Value As String)
      ' Username passed in from the asp page
      gsUsername = Value
    End Set
  End Property
  Public WriteOnly Property Connection() As Object
    Set(ByVal Value As Object)

      ' JDM - Create connection object differently if we are in development mode (i.e. debug mode)
      If ASRDEVELOPMENT Then
        gADOCon = New ADODB.Connection
        'UPGRADE_WARNING: Couldn't resolve default property of object vConnection. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        gADOCon.Open(Value)

        CreateASRDev_SysProtects(gADOCon)
      Else
        gADOCon = Value
      End If

      SetupTablesCollection()

      ReadPersonnelParameters()

      ReadBankHolidayParameters()

    End Set
  End Property

  Public ReadOnly Property BaseIDColumn() As String
    Get
      BaseIDColumn = "?ID_" & mstrCalendarReportsBaseTableName
    End Get
  End Property
  Public ReadOnly Property EventIDColumn() As String
    Get
      EventIDColumn = "?ID_EventID"
    End Get
  End Property
  Public ReadOnly Property PersonnelBase() As Boolean
    Get
      PersonnelBase = (mlngCalendarReportsBaseTable = glngPersonnelTableID)
    End Get
  End Property
  Public ReadOnly Property BaseTableRealSource() As String
    Get
      BaseTableRealSource = mstrBaseTableRealSource
    End Get
  End Property
  Public ReadOnly Property BaseTableID() As String
    Get
      BaseTableID = CStr(mlngCalendarReportsBaseTable)
    End Get
  End Property
  Public ReadOnly Property BaseDesc1IsDate() As Boolean
    Get
      BaseDesc1IsDate = mblnDesc1IsDate
    End Get
  End Property

  Public ReadOnly Property BaseDesc2IsDate() As Boolean
    Get
      BaseDesc2IsDate = mblnDesc2IsDate
    End Get
  End Property
  Public ReadOnly Property BaseDescExprIsDate() As Boolean
    Get
      BaseDescExprIsDate = mblnDescExprIsDate
    End Get
  End Property
  Public ReadOnly Property SQLIDs() As String
    Get
      SQLIDs = mstrSQLIDs
    End Get
  End Property
  Public ReadOnly Property StaticRegionColumn() As String
    Get
      StaticRegionColumn = mstrRegion
    End Get
  End Property
  Public ReadOnly Property StaticRegionColumnID() As Integer
    Get
      StaticRegionColumnID = mlngRegion
    End Get
  End Property
  Public ReadOnly Property ReportStartDate() As Date
    Get
      ReportStartDate = mdtStartDate
    End Get
  End Property
  Public ReadOnly Property ReportEndDate() As Date
    Get
      ReportEndDate = CDate(mdtEndDate)
    End Get
  End Property
  Public ReadOnly Property ReportStartDate_US() As String
    Get
      ReportStartDate_US = VB6.Format(mdtStartDate, "mm/dd/yyyy")
    End Get
  End Property
  Public ReadOnly Property ReportStartDate_Calendar() As Date
    Get
      ReportStartDate_Calendar = CDate(VB6.Format(mdtStartDate, CALREP_DATEFORMAT))
    End Get
  End Property

  Public ReadOnly Property ReportEndDate_Calendar() As Date
    Get
      ReportEndDate_Calendar = CDate(VB6.Format(mdtEndDate, CALREP_DATEFORMAT))
    End Get
  End Property

  Public ReadOnly Property ReportEndDate_CalendarString() As String
    Get
      ReportEndDate_CalendarString = Replace(VB6.Format(mdtEndDate, CALREP_DATEFORMAT), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/")
    End Get
  End Property
  Public ReadOnly Property ReportStartDate_CalendarString() As String
    Get
      ReportStartDate_CalendarString = Replace(VB6.Format(mdtStartDate, CALREP_DATEFORMAT), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/")
    End Get
  End Property

  Public ReadOnly Property ReportEndDate_US() As String
    Get
      ReportEndDate_US = VB6.Format(mdtEndDate, "mm/dd/yyyy")
    End Get
  End Property
  Public ReadOnly Property CalendarReportTitle() As String
    Get
      If mblnCustomReportsPrintFilterHeader Then
        If (mlngCalendarReportsFilterID > 0) Then
          CalendarReportTitle = mstrCalendarReportsName & " (Base Table filter : " & datGeneral.GetFilterName(mlngCalendarReportsFilterID) & ")"
        ElseIf (mlngCalendarReportsPickListID > 0) Then
          CalendarReportTitle = mstrCalendarReportsName & " (Base Table picklist : " & datGeneral.GetPicklistName(mlngCalendarReportsPickListID) & ")"
        End If
      Else
        CalendarReportTitle = mstrCalendarReportsName
      End If
    End Get
  End Property
  Public ReadOnly Property CalendarReportName() As String
    Get
      CalendarReportName = mstrCalendarReportsName
    End Get
  End Property
  Public ReadOnly Property EventsRecordset() As ADODB.Recordset
    Get
      EventsRecordset = mrsCalendarReportsOutput
    End Get
  End Property
  Public ReadOnly Property BaseRecordset() As ADODB.Recordset
    Get
      BaseRecordset = mrsCalendarBaseInfo
    End Get
  End Property
  Public ReadOnly Property GroupByDescription() As Boolean
    Get
      GroupByDescription = mblnGroupByDescription
    End Get
  End Property

  Public Property ShowBankHolidays() As Boolean
    Get
      ShowBankHolidays = mblnShowBankHolidays
    End Get
    Set(ByVal Value As Boolean)
      mblnShowBankHolidays = Value
    End Set
  End Property
  Public Property ShowCaptions() As Boolean
    Get
      ShowCaptions = mblnShowCaptions
    End Get
    Set(ByVal Value As Boolean)
      mblnShowCaptions = Value
    End Set
  End Property
  Public Property ShowWeekends() As Boolean
    Get
      ShowWeekends = mblnShowWeekends
    End Get
    Set(ByVal Value As Boolean)
      mblnShowWeekends = Value
    End Set
  End Property
  Public Property StartOnCurrentMonth() As Boolean
    Get
      StartOnCurrentMonth = mbStartOnCurrentMonth
    End Get
    Set(ByVal Value As Boolean)
      mbStartOnCurrentMonth = Value
    End Set
  End Property
  Public Property IncludeWorkingDaysOnly() As Boolean
    Get
      IncludeWorkingDaysOnly = mblnIncludeWorkingDaysOnly
    End Get
    Set(ByVal Value As Boolean)
      mblnIncludeWorkingDaysOnly = Value
    End Set
  End Property
  Public Property IncludeBankHolidays() As Boolean
    Get
      IncludeBankHolidays = mblnIncludeBankHolidays
    End Get
    Set(ByVal Value As Boolean)
      mblnIncludeBankHolidays = Value
    End Set
  End Property

  Public ReadOnly Property OutputArray_Definition() As Object
    Get

      ' Holds the HTML for the grid definition (object tag etc)
      OutputArray_Definition = VB6.CopyArray(mvarOutputArray_Definition)

    End Get
  End Property

  Public ReadOnly Property OutputArray_Columns() As Object
    Get

      ' Holds the HTML for the columns in the grid (2 + No. fields on report)
      OutputArray_Columns = VB6.CopyArray(mvarOutputArray_Columns)

    End Get
  End Property

  Public ReadOnly Property OutputArray_Merges() As Object
    Get

      OutputArray_Merges = VB6.CopyArray(mvarOutputArray_Merges)

    End Get
  End Property

  Public ReadOnly Property OutputArray_Styles() As Object
    Get

      OutputArray_Styles = VB6.CopyArray(mvarOutputArray_Styles)

    End Get
  End Property

  Public ReadOnly Property OutputArray_Data() As Object
    Get

      ' Holds the HTML for the actual data (and closes object tag)
      OutputArray_Data = VB6.CopyArray(mvarOutputArray_Data)

    End Get
  End Property

  Public Sub DEBUG_TEXT(ByRef DEBUG_STRING As String)
    FileOpen(1, My.Application.Info.DirectoryPath & "\calrep.txt", OpenMode.Append)
    PrintLine(1, VB6.Format(Now, "dd mmm yyyy    hh:mm:ss") & vbTab & DEBUG_STRING & vbNewLine & vbNewLine)
    FileClose(1)
  End Sub

  Public Function OutputArray_Clear() As Object
    ReDim mvarOutputArray_Definition(0)
    ReDim mvarOutputArray_Columns(0)
    ReDim mvarOutputArray_Data(0)
    ReDim mvarOutputArray_Styles(0)
    ReDim mvarOutputArray_Merges(0)
    mlngGridRowIndex = 0
  End Function

  Private Function CreateTempTable() As Boolean

    Dim strSQL As String

    mstrTempTableName = datGeneral.UniqueSQLObjectName("ASRSysTempCalendarReport", 3)

    strSQL = vbNullString
    strSQL = strSQL & "CREATE TABLE [" & mstrTempTableName & "] ("
    strSQL = strSQL & mstrSQLCreateTable
    strSQL = strSQL & ")"

    mclsData.ExecuteSql(strSQL)

    mblnTempTableCreated = True
    CreateTempTable = True

TidyUpAndExit:
    Exit Function

CreateTempTable_ERROR:
    CreateTempTable = False
    mstrErrorString = "Error whilst creating Temporary Table." & vbNewLine & Err.Description
    GoTo TidyUpAndExit

  End Function

  Public Function ConvertEventDescription(ByRef plngColumnID As Integer, ByRef pvarValue As Object) As String

    Dim strTempEventDesc As String
    Dim iDecimals As Short
    Dim strFormat As String

    'get the datatype/properties for the desc1 column
    'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
    If (plngColumnID > 0) And (Not IsDBNull(pvarValue)) Then
      If datGeneral.DoesColumnUseSeparators(plngColumnID) Then
        iDecimals = datGeneral.GetDecimalsSize(plngColumnID)
        strFormat = "#,0" & IIf(iDecimals > 0, "." & New String("#", iDecimals), "")
        'UPGRADE_WARNING: Couldn't resolve default property of object pvarValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        strTempEventDesc = VB6.Format(pvarValue, strFormat)

      ElseIf datGeneral.GetColumnDataType(plngColumnID) = Declarations.SQLDataType.sqlBoolean Then
        'UPGRADE_WARNING: Couldn't resolve default property of object pvarValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        strTempEventDesc = pvarValue

      ElseIf datGeneral.GetColumnDataType(plngColumnID) = Declarations.SQLDataType.sqlDate Then
        'UPGRADE_WARNING: Couldn't resolve default property of object pvarValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        strTempEventDesc = VB6.Format(pvarValue, mstrClientDateFormat)

      Else
        'UPGRADE_WARNING: Couldn't resolve default property of object pvarValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        strTempEventDesc = pvarValue

      End If
    Else
      strTempEventDesc = vbNullString
    End If

    ConvertEventDescription = strTempEventDesc

  End Function

  Private Function GetDescriptionDataTypes() As Boolean

    Dim iDecimals As Short

    mintType_BaseDesc1 = -1
    mintType_BaseDesc2 = -1
    mintType_BaseDescExpr = -1
    mstrFormat_BaseDesc1 = vbNullString
    mstrFormat_BaseDesc2 = vbNullString

    'get the datatype/properties for the desc1 column
    If (mlngDescription1 > 0) Then
      If datGeneral.DoesColumnUseSeparators(mlngDescription1) Then
        mintType_BaseDesc1 = 3
        iDecimals = datGeneral.GetDecimalsSize(mlngDescription1)
        mstrFormat_BaseDesc1 = "#,0" & IIf(iDecimals > 0, "." & New String("#", iDecimals), "")
      ElseIf datGeneral.BitColumn("C", mlngCalendarReportsBaseTable, mlngDescription1) Then
        mintType_BaseDesc1 = 2
      ElseIf datGeneral.DateColumn("C", mlngCalendarReportsBaseTable, mlngDescription1) Then
        mintType_BaseDesc1 = 1
      Else
        mintType_BaseDesc1 = 0
      End If
    End If
    'get the datatype/properties for the desc2 column
    If (mlngDescription2 > 0) Then
      If datGeneral.DoesColumnUseSeparators(mlngDescription2) Then
        mintType_BaseDesc2 = 3
        iDecimals = datGeneral.GetDecimalsSize(mlngDescription2)
        mstrFormat_BaseDesc2 = "#,0" & IIf(iDecimals > 0, "." & New String("#", iDecimals), "")
      ElseIf datGeneral.BitColumn("C", mlngCalendarReportsBaseTable, mlngDescription2) Then
        mintType_BaseDesc2 = 2
      ElseIf datGeneral.DateColumn("C", mlngCalendarReportsBaseTable, mlngDescription2) Then
        mintType_BaseDesc2 = 1
      Else
        mintType_BaseDesc2 = 0
      End If
    End If
    'get the datatype/properties for the descexpr column
    If (mlngDescriptionExpr > 0) Then
      If datGeneral.BitColumn("X", mlngCalendarReportsBaseTable, mlngDescriptionExpr) Then
        mintType_BaseDescExpr = 2
      ElseIf datGeneral.DateColumn("X", mlngCalendarReportsBaseTable, mlngDescriptionExpr) Then
        mintType_BaseDescExpr = 1
      Else
        mintType_BaseDescExpr = 0
      End If
    End If

  End Function

  Private Function InsertIntoTempTable(ByRef pstrSelectString As String) As Boolean

    Dim strSQL As String
    Dim fOK As Boolean

    fOK = True

    If Not mblnTempTableCreated Then
      If Not CreateTempTable() Then
        InsertIntoTempTable = False
        mstrErrorString = "Error creating the temporary table"
        GoTo TidyUpAndExit
      End If
    End If

    If (Not mblnUDFsCreated) Then
      fOK = UDFFunctions(mastrUDFsRequired, True)
      mblnUDFsCreated = fOK
      If Not fOK Then
        InsertIntoTempTable = False
        mstrErrorString = "Error creating SQL User Defined Functions"
        GoTo TidyUpAndExit
      End If
    End If

    strSQL = vbNullString
    strSQL = strSQL & "INSERT INTO [" & mstrTempTableName & "] "
    strSQL = strSQL & pstrSelectString

    mclsData.ExecuteSql(strSQL)

    InsertIntoTempTable = True

TidyUpAndExit:
    Exit Function

ErrorTrap:
    InsertIntoTempTable = False
    mstrErrorString = "Error inserting into the temporary table"
    GoTo TidyUpAndExit

  End Function

  Private Function AddToArray_Columns(ByRef pstrRowToAdd As String) As Boolean

    On Error GoTo AddError

    ReDim Preserve mvarOutputArray_Columns(UBound(mvarOutputArray_Columns) + 1)
    'UPGRADE_WARNING: Couldn't resolve default property of object mvarOutputArray_Columns(UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    mvarOutputArray_Columns(UBound(mvarOutputArray_Columns)) = pstrRowToAdd

    AddToArray_Columns = True
    Exit Function

AddError:

    AddToArray_Columns = False
    mstrErrorString = "Error adding to columns array:" & vbNewLine & Err.Description

  End Function

  Private Function AddToArray_Styles(ByRef pstrRowToAdd As String) As Boolean

    On Error GoTo AddError

    Dim varTempArray() As Object

    ReDim Preserve mvarOutputArray_Styles(mlngStylePageArrayIndex)

    'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
    If IsNothing(mvarOutputArray_Styles(mlngStylePageArrayIndex)) Then
      ReDim varTempArray(0)
    Else
      'UPGRADE_WARNING: Couldn't resolve default property of object mvarOutputArray_Styles(mlngStylePageArrayIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      varTempArray = mvarOutputArray_Styles(mlngStylePageArrayIndex)
    End If

    ReDim Preserve varTempArray(UBound(varTempArray) + 1)
    'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    varTempArray(UBound(varTempArray)) = pstrRowToAdd

    'UPGRADE_WARNING: Couldn't resolve default property of object mvarOutputArray_Styles(mlngStylePageArrayIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    mvarOutputArray_Styles(mlngStylePageArrayIndex) = VB6.CopyArray(varTempArray)

    AddToArray_Styles = True
    Exit Function

AddError:
    AddToArray_Styles = False
    mstrErrorString = "Error adding to styles array:" & vbNewLine & Err.Description

  End Function

  Private Function AddToArray_Merges(ByRef pstrRowToAdd As String) As Boolean

    On Error GoTo AddError

    Dim varTempArray() As Object

    ReDim Preserve mvarOutputArray_Merges(mlngMergePageArrayIndex)

    'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
    If IsNothing(mvarOutputArray_Merges(mlngMergePageArrayIndex)) Then
      ReDim varTempArray(0)
    Else
      'UPGRADE_WARNING: Couldn't resolve default property of object mvarOutputArray_Merges(mlngMergePageArrayIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      varTempArray = mvarOutputArray_Merges(mlngMergePageArrayIndex)
    End If

    ReDim Preserve varTempArray(UBound(varTempArray) + 1)
    'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    varTempArray(UBound(varTempArray)) = pstrRowToAdd

    'UPGRADE_WARNING: Couldn't resolve default property of object mvarOutputArray_Merges(mlngMergePageArrayIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    mvarOutputArray_Merges(mlngMergePageArrayIndex) = VB6.CopyArray(varTempArray)

    AddToArray_Merges = True
    Exit Function

AddError:
    AddToArray_Merges = False
    mstrErrorString = "Error adding to merge array:" & vbNewLine & Err.Description

  End Function

  Private Function AddToArray_Data(ByRef pintRow As Integer, ByRef pintCol As Short, ByRef pstrValue As String, Optional ByRef pblnLastValue As Boolean = False) As Boolean

    Dim lngRowCount As Integer

    On Error GoTo AddError


    'adds a single value (pstrValue) to a position in the grid denoted by the x (pintRow) and y (pintCol) indicies.
    If Not pblnLastValue Then
      ReDim Preserve mvarOutputArray_Data(UBound(mvarOutputArray_Data) + 1)
      'UPGRADE_WARNING: Couldn't resolve default property of object mvarOutputArray_Data(UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      mvarOutputArray_Data(UBound(mvarOutputArray_Data)) = "        <PARAM NAME=""Row(" & mlngGridRowIndex & ").Col(" & pintCol & ")"" VALUE=""" & pstrValue & """>" & vbNewLine

    Else
      ReDim Preserve mvarOutputArray_Data(UBound(mvarOutputArray_Data) + 1)
      'UPGRADE_WARNING: Couldn't resolve default property of object mvarOutputArray_Data(UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      mvarOutputArray_Data(UBound(mvarOutputArray_Data)) = "        <PARAM NAME=""Row(" & mlngGridRowIndex & ").Col(" & pintCol & ")"" VALUE=""" & pstrValue & """>" & vbNewLine

      lngRowCount = mlngGridRowIndex + 1
      ReDim Preserve mvarOutputArray_Data(UBound(mvarOutputArray_Data) + 1)
      'UPGRADE_WARNING: Couldn't resolve default property of object mvarOutputArray_Data(UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      mvarOutputArray_Data(UBound(mvarOutputArray_Data)) = "        <PARAM NAME=""Row.Count"" VALUE=""" & lngRowCount & """>" & vbNewLine

    End If

    AddToArray_Data = True
    Exit Function

AddError:

    AddToArray_Data = False
    mstrErrorString = "Error adding to data array:" & vbNewLine & Err.Description

  End Function


  Private Function AddToArray_Definition(ByRef pstrRowToAdd As String) As Boolean

    On Error GoTo AddError

    ReDim Preserve mvarOutputArray_Definition(UBound(mvarOutputArray_Definition) + 1)
    'UPGRADE_WARNING: Couldn't resolve default property of object mvarOutputArray_Definition(UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    mvarOutputArray_Definition(UBound(mvarOutputArray_Definition)) = pstrRowToAdd

    AddToArray_Definition = True
    Exit Function

AddError:

    AddToArray_Definition = False
    mstrErrorString = "Error adding to definition array:" & vbNewLine & Err.Description

  End Function


  Private Function DaysInMonth(ByRef pdtMonth As Date) As Short

    'Return the number of days in the month

    Dim dtNextMonth As Date

    dtNextMonth = DateAdd(Microsoft.VisualBasic.DateInterval.Month, 1, pdtMonth)
    DaysInMonth = VB.Day(DateAdd(Microsoft.VisualBasic.DateInterval.Day, VB.Day(dtNextMonth) * -1, dtNextMonth))

  End Function



  Private Sub DebugMSG(ByRef strInput As String, Optional ByRef blnOverwriteExisting As Boolean = False)

    'Ignore any errors in here...
    On Error GoTo LocalErr

    Dim strFileName As String

    strFileName = My.Application.Info.DirectoryPath & "\debug.txt"

    If blnOverwriteExisting Then
      'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
      If Dir(strFileName) <> vbNullString Then
        Kill(strFileName)
      End If
    End If

    FileOpen(99, strFileName, OpenMode.Append)

    PrintLine(99, Now & "  " & strInput)
    FileClose(99)

LocalErr:
    Err.Clear()

  End Sub


  Public Function OutputReport(ByRef blnPrompt As Boolean) As Boolean

    Dim intMonth As Short
    Dim intMonthCount As Short
    Dim dtMonth As Date
    Dim fOK As Boolean
    Dim strPageName As String

    On Error GoTo ErrorTrap

    Dim mavOutputDataIndex(2, 0) As Object
    ReDim mvarOutputArray_Styles(0)
    ReDim mvarOutputArray_Merges(0)

    fOK = True

    '  DebugMSG "Started @ " & Now(), True

    mlngBC_Data = 13434879
    mlngFC_Data = 0
    mlngBC_Heading = 13395456
    mlngFC_Heading = 16777215

    mlngColor_Weekend = 12632256
    mlngColor_RangeDisabled = 9868950
    mlngColor_Disabled = 8421504

    mstrExcludedColours = CStr(mlngBC_Data)
    'UPGRADE_WARNING: Couldn't resolve default property of object GetUserSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    mlngBC_Data = GetUserSetting("output", "databackcolour", 13434879)
    mstrExcludedColours = mstrExcludedColours & ", " & CStr(mlngBC_Data)

    GetAvailableColours(mstrExcludedColours)

    mlngGridRowIndex = 0

    Load_Legend()

    intMonthCount = DateDiff(Microsoft.VisualBasic.DateInterval.Month, mdtStartDate, CDate(mdtEndDate))

    '***********************************************************
    'get an array for the key
    mlngStylePageArrayIndex = 0
    mlngMergePageArrayIndex = 0
    OutputArray_GetLegendArray()
    AddToArray_Data(CShort(UBound(mvarOutputArray_Data)), 0, "*")
    AddToArray_Data(CShort(UBound(mvarOutputArray_Data)), 1, "Key")
    mlngGridRowIndex = mlngGridRowIndex + 1
    '***********************************************************

    For intMonth = 0 To intMonthCount Step 1
      mlngStylePageArrayIndex = mlngStylePageArrayIndex + 1
      mlngMergePageArrayIndex = mlngMergePageArrayIndex + 1

      mcolBaseDescIndex_Output = New Collection

      dtMonth = DateAdd(Microsoft.VisualBasic.DateInterval.Month, intMonth, mdtStartDate)
      mlngYear_Output = Year(dtMonth)
      mlngMonth_Output = Month(dtMonth)

      mintDaysInMonth_Output = DaysInMonth(dtMonth)

      'Define the current visible Start and End Dates.
      mdtVisibleEndDate_Output = DateAdd(Microsoft.VisualBasic.DateInterval.Day, CDbl(mintDaysInMonth_Output - VB.Day(dtMonth)), dtMonth)
      mdtVisibleStartDate_Output = DateAdd(Microsoft.VisualBasic.DateInterval.Day, CDbl(-(mintDaysInMonth_Output - 1)), mdtVisibleEndDate_Output)

      mintFirstDayOfMonth_Output = Weekday(mdtVisibleStartDate_Output, FirstDayOfWeek.Sunday)

      OutputArray_GetArray()

      strPageName = MonthName(mlngMonth_Output) & " " & mlngYear_Output

      AddToArray_Data(CInt(UBound(mvarOutputArray_Data)), 1, strPageName, False)
      AddToArray_Data(CInt(UBound(mvarOutputArray_Data)), 0, "*", IIf((intMonth = intMonthCount), True, False))
      mlngGridRowIndex = mlngGridRowIndex + 1

      'UPGRADE_NOTE: Object mcolBaseDescIndex_Output may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      mcolBaseDescIndex_Output = Nothing

    Next intMonth

    OutputReport = True

TidyUpAndExit:
    Exit Function

ErrorTrap:
    OutputReport = False
    GoTo TidyUpAndExit

  End Function

  Private Function OutputArray_AddCalendar() As Boolean

    On Error GoTo ErrorTrap

    Dim iNewIndex As Short
    Dim intDateValue As Short
    Dim intDateCount As Short
    Dim iCount As Short
    Dim iCount2 As Short
    Dim intCurrentIndex As Short
    Dim intControlCount As Short
    Dim intSessionCount As Short
    Dim intNextIndex As Short

    Dim dtLabelsDate As Date

    Dim lngBaseID As Integer
    Dim strDate As String
    Dim strSession As String
    Dim strIsBankHoliday As String
    Dim strIsWeekend As String
    Dim strIsWorkingDay As String
    Dim intHasEvent As Short
    Dim strCaption As String
    Dim strBackColour As String
    Dim strForeColour As String

    Dim strRegion As String
    Dim strWorkingPattern As String

    Dim varTempArray(,) As Object

    Dim blnNewBaseRecord As Boolean
    Dim strTempRecordDesc As String
    Dim intDescEmpty As Short
    Dim blnDescEmpty As Boolean

    Dim strBaseDescription1, strBaseDescription2 As Object
    Dim strBaseDescriptionExpr As String
    Dim iDecimals As Short
    Dim strFormat As String

    intDateCount = 0
    mstrBaseRecDesc_Output = vbNullString
    mintBaseRecordCount_Output = 0
    mstrBaseRecDesc = vbNullString
    mintCurrentBaseIndex_Output = 0
    blnNewBaseRecord = True
    mintRangeStartIndex_Output = 0
    mintRangeEndIndex_Output = 0
    mlngCurrentRecordID = -1

    With mrsCalendarBaseInfo
      If .BOF And .EOF Then
        OutputArray_AddCalendar = False
        GoTo TidyUpAndExit
      End If

      mintBaseCount_Output = .RecordCount
      ReDim mavOutputDateIndex(2, 0)

      .MoveFirst()
      Do While Not .EOF

        ' Get base description 1
        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        If Not IsDBNull(.Fields("Description1").Value) Then
          If datGeneral.DoesColumnUseSeparators(mlngDescription1) Then
            iDecimals = datGeneral.GetDecimalsSize(mlngDescription1)
            strFormat = "#,0" & IIf(iDecimals > 0, "." & New String("#", iDecimals), "")
            'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            strBaseDescription1 = VB6.Format(.Fields("Description1").Value, strFormat)
          ElseIf datGeneral.BitColumn("C", mlngCalendarReportsBaseTable, mlngDescription1) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            strBaseDescription1 = IIf(.Fields("Description1").Value, "Y", "N")
          ElseIf datGeneral.DateColumn("C", mlngCalendarReportsBaseTable, mlngDescription1) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            strBaseDescription1 = VB6.Format(.Fields("Description1").Value, mstrClientDateFormat)
          Else
            'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            strBaseDescription1 = .Fields("Description1").Value
          End If
        Else
          'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          strBaseDescription1 = vbNullString
        End If

        ' Get base description 2
        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        If Not IsDBNull(.Fields("Description2").Value) Then
          If datGeneral.DoesColumnUseSeparators(mlngDescription2) Then
            iDecimals = datGeneral.GetDecimalsSize(mlngDescription2)
            strFormat = "#,0" & IIf(iDecimals > 0, "." & New String("#", iDecimals), "")
            'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            strBaseDescription2 = VB6.Format(.Fields("Description2").Value, strFormat)
          ElseIf datGeneral.BitColumn("C", mlngCalendarReportsBaseTable, mlngDescription2) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            strBaseDescription2 = IIf(.Fields("Description2").Value, "Y", "N")
          ElseIf datGeneral.DateColumn("C", mlngCalendarReportsBaseTable, mlngDescription1) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            strBaseDescription2 = VB6.Format(.Fields("Description2").Value, mstrClientDateFormat)
          Else
            'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            strBaseDescription2 = .Fields("Description2").Value
          End If
        Else
          'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          strBaseDescription2 = vbNullString
        End If

        ' Get base description expression
        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        If Not IsDBNull(.Fields("DescriptionExpr").Value) Then
          If datGeneral.BitColumn("X", mlngCalendarReportsBaseTable, mlngDescriptionExpr) Then
            strBaseDescriptionExpr = IIf(.Fields("DescriptionExpr").Value, "Y", "N")
          ElseIf datGeneral.DateColumn("X", mlngCalendarReportsBaseTable, mlngDescriptionExpr) Then
            strBaseDescriptionExpr = IIf(.Fields("DescriptionExpr").Value, "Y", "N")
          Else
            strBaseDescriptionExpr = .Fields("DescriptionExpr").Value
          End If
        Else
          strBaseDescriptionExpr = vbNullString
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        strTempRecordDesc = strBaseDescription1
        'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        strTempRecordDesc = strTempRecordDesc & IIf((Len(strTempRecordDesc) > 0) And (Len(strBaseDescription2) > 0), mstrDescriptionSeparator, "") & strBaseDescription2
        strTempRecordDesc = strTempRecordDesc & IIf((Len(strTempRecordDesc) > 0) And (Len(strBaseDescriptionExpr) > 0), mstrDescriptionSeparator, "") & strBaseDescriptionExpr

        blnDescEmpty = (strTempRecordDesc = vbNullString)
        If blnDescEmpty Then
          intDescEmpty = intDescEmpty + 1
        Else
          intDescEmpty = 0
        End If

        If mblnGroupByDescription Then
          If ((strTempRecordDesc) <> mstrBaseRecDesc) Or (blnDescEmpty And Int(CDbl(intDescEmpty = 1))) Then
            blnNewBaseRecord = True
            blnDescEmpty = False

            mstrBaseRecDesc = strTempRecordDesc

            If Len(Trim(mstrRegionColumnRealSource)) > 0 Then
              'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
              mstrCurrentBaseRegion = IIf(IsDBNull(.Fields("Region").Value), "", .Fields("Region").Value)
            End If
            mintBaseRecordCount_Output = mintBaseRecordCount_Output + 1
          End If
          mlngCurrentRecordID = .Fields(mstrBaseIDColumn).Value

        Else
          If .Fields(mstrBaseIDColumn).Value <> mlngCurrentRecordID Then
            blnNewBaseRecord = True

            mstrBaseRecDesc = strTempRecordDesc

            mlngCurrentRecordID = .Fields(mstrBaseIDColumn).Value
            If Len(Trim(mstrRegionColumnRealSource)) > 0 Then
              'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
              mstrCurrentBaseRegion = IIf(IsDBNull(.Fields("Region").Value), "", .Fields("Region").Value)
            End If
            mintBaseRecordCount_Output = mintBaseRecordCount_Output + 1
          End If

        End If

        intSessionCount = 0
        mintCurrentBaseIndex_Output = mintCurrentBaseIndex_Output + 1

        ReDim Preserve mavOutputDateIndex(2, mintBaseRecordCount_Output)
        ReDim varTempArray(9, 74)

        If blnNewBaseRecord Then

          For iCount2 = 1 To 74 Step 1
            intSessionCount = intSessionCount + 1

            If intSessionCount = 1 Then
              intControlCount = intControlCount + 1
            End If

            lngBaseID = CInt(.Fields(mstrBaseIDColumn).Value)

            If (intControlCount >= mintFirstDayOfMonth_Output) And (intControlCount < (mintFirstDayOfMonth_Output + mintDaysInMonth_Output)) Then
              strSession = IIf(intSessionCount = 2, " PM", " AM")
              If Trim(strSession) = "AM" Then
                intDateCount = intDateCount + 1
              End If
              '            dtLabelsDate = CDate(intDateCount & "/" & mlngMonth_Output & "/" & CStr(mlngYear_Output))
              dtLabelsDate = DateAdd(Microsoft.VisualBasic.DateInterval.Day, CDbl(intDateCount - 1), mdtVisibleStartDate_Output)

              'calculate the indices of the out of report range bounaries.
              If dtLabelsDate < mdtStartDate Then
                mintRangeStartIndex_Output = intControlCount
              End If
              If dtLabelsDate = CDate(mdtEndDate) Then
                mintRangeEndIndex_Output = intControlCount + 1
              End If

              strDate = VB6.Format(dtLabelsDate, CALREP_DATEFORMAT)
              '            strBackColour = HexValue(lblCalDates(0).BackColor)
              strBackColour = CStr(0)

              strCaption = vbNullString

            Else
              strDate = "  /  /    "
              strSession = vbNullString
              '            strBackColour = HexValue(lblDisabled.BackColor)
              strBackColour = CStr(0)
              strCaption = vbNullString

            End If

            If Trim(strSession) <> vbNullString Then
              strIsBankHoliday = IIf(IsBankHoliday(dtLabelsDate, lngBaseID, strRegion), "1", "0")

              'flag if the date is a weekend
              strIsWeekend = IIf(IsWeekend(dtLabelsDate), "1", "0")

              'flag if the date & session is in the current personnel's working pattern.
              strIsWorkingDay = IIf(IsWorkingDay(dtLabelsDate, lngBaseID, Trim(strSession), strWorkingPattern), "1", "0")

            Else
              strIsBankHoliday = "0"
              strIsWeekend = "0"
              strIsWorkingDay = "0"

            End If

            'Add values to Date Index array
            'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(0, iCount2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            varTempArray(0, iCount2) = lngBaseID
            'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(1, iCount2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            varTempArray(1, iCount2) = strDate
            'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(2, iCount2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            varTempArray(2, iCount2) = strSession
            'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(3, iCount2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            varTempArray(3, iCount2) = strIsBankHoliday
            'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(4, iCount2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            varTempArray(4, iCount2) = strIsWeekend
            'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(5, iCount2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            varTempArray(5, iCount2) = strIsWorkingDay
            'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(6, iCount2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            varTempArray(6, iCount2) = 0
            'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(7, iCount2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            varTempArray(7, iCount2) = strCaption
            'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(8, iCount2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            varTempArray(8, iCount2) = strBackColour
            '          varTempArray(9, iCount2) = HexValue(vbBlack)
            'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(9, iCount2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            varTempArray(9, iCount2) = 0

            If intSessionCount = 2 Then
              intSessionCount = 0
            End If

          Next iCount2

          'UPGRADE_WARNING: Couldn't resolve default property of object mavOutputDateIndex(0, mintBaseRecordCount_Output). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          mavOutputDateIndex(0, mintBaseRecordCount_Output) = .Fields(mstrBaseIDColumn).Value
          'UPGRADE_WARNING: Couldn't resolve default property of object mavOutputDateIndex(1, mintBaseRecordCount_Output). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          mavOutputDateIndex(1, mintBaseRecordCount_Output) = mstrBaseRecDesc
          'UPGRADE_WARNING: Couldn't resolve default property of object mavOutputDateIndex(2, mintBaseRecordCount_Output). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          mavOutputDateIndex(2, mintBaseRecordCount_Output) = VB6.CopyArray(varTempArray)

        End If

        mcolBaseDescIndex_Output.Add(mintBaseRecordCount_Output, CStr(.Fields(mstrBaseIDColumn).Value))

        ReDim varTempArray(9, 0)

        intControlCount = 0
        intDateCount = 0

        blnNewBaseRecord = False

        .MoveNext()
      Loop

    End With

    OutputArray_AddCalendar = True

TidyUpAndExit:
    Exit Function

ErrorTrap:
    OutputArray_AddCalendar = False
    GoTo TidyUpAndExit

  End Function

  Private Function IsBankHoliday(ByRef pdtDate As Date, ByRef plngBaseID As Integer, ByRef pstrRegion As String) As Boolean

    On Error GoTo ErrorTrap

    Dim colBankHolidays As clsBankHolidays
    Dim objBankHoliday As clsBankHoliday

    If mblnPersonnelBase And (modPersonnelSpecifics.grtRegionType = modPersonnelSpecifics.RegionType.rtHistoricRegion) And (Not mblnGroupByDescription) And (mlngRegion < 1) Then

      'Need to get the current region from the previously populated.
      'NB. cant get the region from the collection as the current region is required even
      'when the date is NOT a bank holiday
      pstrRegion = GetCurrentRegion(plngBaseID, pdtDate)

      'Historic Region Bank Holidays
      colBankHolidays = mcolHistoricBankHolidays.Item(CStr(plngBaseID))

      For Each objBankHoliday In colBankHolidays.Collection
        With objBankHoliday
          If pdtDate = .HolidayDate Then
            'pstrRegion = .Region
            IsBankHoliday = True
            GoTo TidyUpAndExit
          End If
        End With
      Next objBankHoliday

    ElseIf ((mlngRegion > 0) Or (mblnPersonnelBase And (modPersonnelSpecifics.grtRegionType = modPersonnelSpecifics.RegionType.rtStaticRegion))) And (Not mblnGroupByDescription) Then

      'Static Region Bank Holidays
      colBankHolidays = mcolStaticBankHolidays.Item(CStr(plngBaseID))

      For Each objBankHoliday In colBankHolidays.Collection
        With objBankHoliday
          If pdtDate = .HolidayDate Then
            pstrRegion = .Region
            IsBankHoliday = True
            GoTo TidyUpAndExit
          End If
        End With
      Next objBankHoliday

    End If

    IsBankHoliday = False

TidyUpAndExit:
    'UPGRADE_NOTE: Object objBankHoliday may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    objBankHoliday = Nothing
    'UPGRADE_NOTE: Object colBankHolidays may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    colBankHolidays = Nothing
    Exit Function

ErrorTrap:
    IsBankHoliday = False
    GoTo TidyUpAndExit

  End Function

  Private Function GetCurrentRegion(ByRef plngBaseRecordID As Integer, ByRef pdtDate As Date) As String

    Dim intCount As Short

    On Error GoTo ErrorTrap

    For intCount = 1 To UBound(mavCareerRanges, 2) Step 1
      'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      If plngBaseRecordID = CInt(mavCareerRanges(0, intCount)) Then
        'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(2, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If mavCareerRanges(2, intCount) <> "" Then
          'has a career change in the past
          'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          If (pdtDate >= CDate(mavCareerRanges(1, intCount))) And (pdtDate < CDate(mavCareerRanges(2, intCount))) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(3, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            GetCurrentRegion = mavCareerRanges(3, intCount)
            Exit Function
          End If
        Else
          'has a effective start date but has no end date. (most recent career change)
          'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          If (pdtDate >= CDate(mavCareerRanges(1, intCount))) Then
            'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(3, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            GetCurrentRegion = mavCareerRanges(3, intCount)
            Exit Function
          End If
        End If
      End If
    Next intCount

TidyUpAndExit:
    Exit Function

ErrorTrap:
    GetCurrentRegion = vbNullString
    GoTo TidyUpAndExit

  End Function

  Private Function IsWeekend(ByRef pdtDate As Date) As Boolean
    If (Weekday(pdtDate, FirstDayOfWeek.Sunday) = FirstDayOfWeek.Saturday) Or (Weekday(pdtDate, FirstDayOfWeek.Sunday) = FirstDayOfWeek.Sunday) Then
      IsWeekend = True
    Else
      IsWeekend = False
    End If
  End Function

  Private Function IsWorkingDay(ByRef pdtDate As Date, ByRef plngBaseID As Integer, ByRef pstrSession As String, ByRef pstrWorkingPattern As String) As Boolean

    On Error GoTo ErrorTrap

    Dim colWorkingPatterns As clsCalendarEvents
    Dim objWorkingPattern As clsCalendarEvent

    Dim strWorkingPattern As String
    Dim intWeekDay As String

    Const WORKINGPATTERN_LENGTH As Short = 14

    strWorkingPattern = "              " 'empty working pattern
    intWeekDay = CStr(Weekday(pdtDate, FirstDayOfWeek.Sunday))

    If mblnPersonnelBase And (modPersonnelSpecifics.gwptWorkingPatternType = modPersonnelSpecifics.WorkingPatternType.wptHistoricWPattern) And (Not mblnGroupByDescription) Then

      'Historic Working Pattern

      colWorkingPatterns = mcolHistoricWorkingPatterns.Item(CStr(plngBaseID))
      For Each objWorkingPattern In colWorkingPatterns.Collection
        With objWorkingPattern

          'TM02072004 Fault 8851 - Force the working pattern length to be 14 characters!
          If Len(.WorkingPattern) < WORKINGPATTERN_LENGTH Then
            .WorkingPattern = .WorkingPattern & New String(" ", WORKINGPATTERN_LENGTH - Len(.WorkingPattern))
          ElseIf Len(.WorkingPattern) > WORKINGPATTERN_LENGTH Then
            .WorkingPattern = Left(.WorkingPattern, WORKINGPATTERN_LENGTH)
          End If

          If (.EndDateName <> vbNullString) Then
            If (pdtDate >= CDate(.StartDateName)) And (pdtDate < CDate(.EndDateName)) Then
              Select Case UCase(pstrSession)
                Case "AM"
                  If Mid(.WorkingPattern, (CDbl(intWeekDay) * 2) - 1, 1) = " " Then
                    pstrWorkingPattern = .WorkingPattern
                    IsWorkingDay = False
                    GoTo TidyUpAndExit
                  Else
                    pstrWorkingPattern = .WorkingPattern
                    IsWorkingDay = True
                    GoTo TidyUpAndExit
                  End If
                Case "PM"
                  If Mid(.WorkingPattern, CDbl(intWeekDay) * 2, 1) = " " Then
                    pstrWorkingPattern = .WorkingPattern
                    IsWorkingDay = False
                    GoTo TidyUpAndExit
                  Else
                    pstrWorkingPattern = .WorkingPattern
                    IsWorkingDay = True
                    GoTo TidyUpAndExit
                  End If
              End Select
            End If
          Else
            If (pdtDate >= CDate(.StartDateName)) Then
              Select Case UCase(pstrSession)
                Case "AM"
                  If Mid(.WorkingPattern, (CDbl(intWeekDay) * 2) - 1, 1) = " " Then
                    pstrWorkingPattern = .WorkingPattern
                    IsWorkingDay = False
                    GoTo TidyUpAndExit
                  Else
                    pstrWorkingPattern = .WorkingPattern
                    IsWorkingDay = True
                    GoTo TidyUpAndExit
                  End If
                Case "PM"
                  If Mid(.WorkingPattern, CDbl(intWeekDay) * 2, 1) = " " Then
                    pstrWorkingPattern = .WorkingPattern
                    IsWorkingDay = False
                    GoTo TidyUpAndExit
                  Else
                    pstrWorkingPattern = .WorkingPattern
                    IsWorkingDay = True
                    GoTo TidyUpAndExit
                  End If
              End Select
            End If
          End If
        End With
      Next objWorkingPattern

    ElseIf mblnPersonnelBase And (modPersonnelSpecifics.gwptWorkingPatternType = modPersonnelSpecifics.WorkingPatternType.wptStaticWPattern) And (Not mblnGroupByDescription) Then

      'Static Working Pattern

      colWorkingPatterns = mcolStaticWorkingPatterns.Item(CStr(plngBaseID))
      For Each objWorkingPattern In colWorkingPatterns.Collection
        With objWorkingPattern

          'TM02072004 Fault 8851 - Force the working pattern length to be 14 characters!
          If Len(.WorkingPattern) < WORKINGPATTERN_LENGTH Then
            .WorkingPattern = .WorkingPattern & New String(" ", WORKINGPATTERN_LENGTH - Len(.WorkingPattern))
          ElseIf Len(.WorkingPattern) > WORKINGPATTERN_LENGTH Then
            .WorkingPattern = Left(.WorkingPattern, WORKINGPATTERN_LENGTH)
          End If

          strWorkingPattern = .WorkingPattern

          Select Case UCase(pstrSession)
            Case "AM"
              If Mid(strWorkingPattern, (CDbl(intWeekDay) * 2) - 1, 1) = " " Then
                pstrWorkingPattern = strWorkingPattern
                IsWorkingDay = False
                GoTo TidyUpAndExit
              Else
                pstrWorkingPattern = strWorkingPattern
                IsWorkingDay = True
                GoTo TidyUpAndExit
              End If
            Case "PM"
              If Mid(strWorkingPattern, CDbl(intWeekDay) * 2, 1) = " " Then
                pstrWorkingPattern = strWorkingPattern
                IsWorkingDay = False
                GoTo TidyUpAndExit
              Else
                pstrWorkingPattern = strWorkingPattern
                IsWorkingDay = True
                GoTo TidyUpAndExit
              End If
          End Select
        End With
      Next objWorkingPattern
    End If

    pstrWorkingPattern = "              "
    IsWorkingDay = False

TidyUpAndExit:
    'UPGRADE_NOTE: Object objWorkingPattern may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    objWorkingPattern = Nothing
    'UPGRADE_NOTE: Object colWorkingPatterns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    colWorkingPatterns = Nothing
    Exit Function

ErrorTrap:
    pstrWorkingPattern = "              "
    IsWorkingDay = False
    GoTo TidyUpAndExit

  End Function

  Private Function OutputArray_AddDates() As Boolean

    On Error GoTo ErrorTrap

    Dim intControlCount As Short
    Dim intDateCount As Short

    AddToArray_Data(1, 0, "")

    For intControlCount = 1 To DAY_CONTROL_COUNT Step 1

      If (intControlCount >= mintFirstDayOfMonth_Output) And (intControlCount < (mintFirstDayOfMonth_Output + mintDaysInMonth_Output)) Then
        intDateCount = intDateCount + 1
        AddToArray_Data(1, intControlCount, CStr(intDateCount))
      Else
        'Add a blank date box
        AddToArray_Data(1, intControlCount, "")
      End If

    Next intControlCount

    mlngGridRowIndex = mlngGridRowIndex + 1

    OutputArray_AddDates = True

TidyUpAndExit:
    Exit Function

ErrorTrap:
    OutputArray_AddDates = False
    GoTo TidyUpAndExit

  End Function

  Private Function OutputArray_AddDays() As Boolean

    On Error GoTo ErrorTrap

    Dim iDayCount As Short
    Dim sDay As String
    Dim intCount As Short

    iDayCount = 1
    sDay = vbNullString
    intCount = 0

    '  mobjOutput.AddColumn "", sqlVarChar, 0
    AddToArray_Data(0, 0, "")

    For intCount = 1 To DAY_CONTROL_COUNT Step 1

      Select Case iDayCount
        Case 1 : sDay = "S"
        Case 2 : sDay = "M"
        Case 3 : sDay = "T"
        Case 4 : sDay = "W"
        Case 5 : sDay = "T"
        Case 6 : sDay = "F"
        Case 7 : sDay = "S"
      End Select

      '    mobjOutput.AddColumn sDay, sqlVarChar, 0
      AddToArray_Data(0, intCount, sDay)

      If iDayCount = 7 Then
        iDayCount = 0
      End If
      iDayCount = iDayCount + 1

    Next intCount

    mlngGridRowIndex = mlngGridRowIndex + 1

    OutputArray_AddDays = True

TidyUpAndExit:
    Exit Function

ErrorTrap:
    OutputArray_AddDays = False
    GoTo TidyUpAndExit

  End Function

  Private Function OutputArray_AddEvents() As Boolean

    On Error GoTo ErrorTrap

    Dim lngStart As Integer
    Dim lngEnd As Integer
    Dim lngCurrentBaseID As Integer
    Dim intBaseRecordIndex As Short

    Dim fOK As Boolean

    Dim sSQL As String

    fOK = True

    With mrsCalendarReportsOutput

      ' If there are no event records, skip this bit
      ' this bit (but still show the form)
      If .BOF And .EOF Then
        Exit Function
      End If

      .MoveFirst()
      ' Loop through the events recordset
      Do Until .EOF

        lngCurrentBaseID = .Fields(mstrBaseIDColumn).Value

        'UPGRADE_WARNING: Couldn't resolve default property of object mcolBaseDescIndex_Output.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        intBaseRecordIndex = mcolBaseDescIndex_Output.Item(CStr(lngCurrentBaseID))

        ' Load each event record data into variables
        ' (has to be done because start/end dates may be modified by code to fill grid correctly)
        mstrCurrentEventKey = .Fields(mstrEventIDColumn).Value

        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        mstrEventLegend_Output = IIf(IsDBNull(.Fields("Legend").Value), "", Left(.Fields("Legend").Value, 2))

        '****************************************************************************
        mdtEventStartDate_Output = .Fields("StartDate").Value

        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        If IsDBNull(.Fields("EndDate").Value) Then
          mdtEventEndDate_Output = mdtEventStartDate_Output
        Else
          mdtEventEndDate_Output = .Fields("EndDate").Value
        End If

        'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        If IsDBNull(.Fields("StartSession").Value) And IsDBNull(.Fields("EndSession").Value) Then
          mstrEventStartSession_Output = "AM"
          mstrEventEndSession_Output = "PM"
          'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
        ElseIf IsDBNull(.Fields("EndSession").Value) Then
          mstrEventEndSession_Output = mstrEventStartSession_Output
        Else
          mstrEventStartSession_Output = UCase(.Fields("StartSession").Value.ToString())
          mstrEventEndSession_Output = UCase(.Fields("EndSession").Value.ToString())
        End If

        mstrDuration_Output = .Fields("Duration").Value

        '****************************************************************************

        ' If the event start date is after the event end date, ignore the record
        If (mdtEventStartDate_Output > mdtEventEndDate_Output) Then

          ' if the event is totally before the currently viewed timespan then do nothing
        ElseIf (mdtEventStartDate_Output < mdtVisibleStartDate_Output) And (mdtEventEndDate_Output < mdtVisibleStartDate_Output) Then

          ' if the event is totally after the currently viewed timespan then do nothing
        ElseIf (mdtEventStartDate_Output > mdtVisibleEndDate_Output) And (mdtEventEndDate_Output > mdtVisibleEndDate_Output) Then

          ' if the event starts before currently viewed timespan, but ends in the timspan then
        ElseIf (mdtEventStartDate_Output < mdtVisibleStartDate_Output) And (mdtEventEndDate_Output <= mdtVisibleEndDate_Output) Then

          mdtEventStartDate_Output = mdtVisibleStartDate_Output
          mstrEventStartSession_Output = "AM"

          lngStart = Output_GetCalArrayIndex(intBaseRecordIndex, mdtEventStartDate_Output, IIf(mstrEventStartSession_Output = "AM", False, True))
          lngEnd = Output_GetCalArrayIndex(intBaseRecordIndex, mdtEventEndDate_Output, IIf(mstrEventEndSession_Output = "AM", False, True))

          fOK = OutputArray_FillEvents(intBaseRecordIndex, lngStart, lngEnd)

          ' if the event starts in the currently viewed timespan, but ends after it then
        ElseIf (mdtEventStartDate_Output >= mdtVisibleStartDate_Output) And (mdtEventEndDate_Output > mdtVisibleEndDate_Output) Then

          mdtEventEndDate_Output = mdtVisibleEndDate_Output
          mstrEventEndSession_Output = "PM"

          lngStart = Output_GetCalArrayIndex(intBaseRecordIndex, mdtEventStartDate_Output, IIf(mstrEventStartSession_Output = "AM", False, True))
          lngEnd = Output_GetCalArrayIndex(intBaseRecordIndex, mdtEventEndDate_Output, IIf(mstrEventEndSession_Output = "AM", False, True))

          fOK = OutputArray_FillEvents(intBaseRecordIndex, lngStart, lngEnd)

          ' if the event is enclosed within viewed timespan, and months are equal then
        ElseIf (mdtEventStartDate_Output >= mdtVisibleStartDate_Output) And (mdtEventEndDate_Output <= mdtVisibleEndDate_Output) And (Month(mdtEventStartDate_Output) = Month(mdtEventEndDate_Output)) Then

          lngStart = Output_GetCalArrayIndex(intBaseRecordIndex, mdtEventStartDate_Output, IIf(mstrEventStartSession_Output = "AM", False, True))
          lngEnd = Output_GetCalArrayIndex(intBaseRecordIndex, mdtEventEndDate_Output, IIf(mstrEventEndSession_Output = "AM", False, True))

          fOK = OutputArray_FillEvents(intBaseRecordIndex, lngStart, lngEnd)

          ' if the event starts before the the viewed timespan and ends after the viewed timespan then
        ElseIf (mdtEventStartDate_Output < mdtVisibleStartDate_Output) And (mdtEventEndDate_Output > mdtVisibleEndDate_Output) Then

          mdtEventStartDate_Output = mdtVisibleStartDate_Output
          mstrEventStartSession_Output = "AM"

          mdtEventEndDate_Output = mdtVisibleEndDate_Output
          mstrEventEndSession_Output = "PM"

          lngStart = Output_GetCalArrayIndex(intBaseRecordIndex, mdtEventStartDate_Output, IIf(mstrEventStartSession_Output = "AM", False, True))
          lngEnd = Output_GetCalArrayIndex(intBaseRecordIndex, mdtEventEndDate_Output, IIf(mstrEventEndSession_Output = "AM", False, True))

          fOK = OutputArray_FillEvents(intBaseRecordIndex, lngStart, lngEnd)

        End If

        If fOK = False Then
          Exit Do
        End If

        .MoveNext()
      Loop
    End With

    OutputArray_AddEvents = True

TidyUpAndExit:
    Exit Function

ErrorTrap:
    OutputArray_AddEvents = False
    GoTo TidyUpAndExit

  End Function

  Public Function Output_GetCalArrayIndex(ByRef pintBaseRecordIndex As Short, ByRef pdtDate As Date, ByRef pblnSession As Boolean) As Short

    ' This function returns the index value for the specified date and session.

    Dim dtFirstDate As Date
    Dim dtLastDate As Date

    Dim iCount As Short
    Dim varTempArray As Object

    '  dtFirstDate = CDate("01/" & mlngMonth_Output & "/" & mlngYear_Output)
    '  dtLastDate = CDate(DaysInMonth(dtFirstDate) & "/" & mlngMonth_Output & "/" & mlngYear_Output)
    dtFirstDate = mdtVisibleStartDate_Output
    dtLastDate = mdtVisibleEndDate_Output

    If (pdtDate < dtFirstDate) Or (pdtDate > dtLastDate) Then
      Output_GetCalArrayIndex = -1
      Exit Function
    End If

    'UPGRADE_WARNING: Couldn't resolve default property of object mavOutputDateIndex(2, pintBaseRecordIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    varTempArray = mavOutputDateIndex(2, pintBaseRecordIndex)

    For iCount = 1 To 74 Step 2

      'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(1, iCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      If varTempArray(1, iCount) <> "  /  /    " Then
        'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(1, iCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If (varTempArray(1, iCount) = VB6.Format(pdtDate, CALREP_DATEFORMAT)) Then
          Output_GetCalArrayIndex = IIf(pblnSession, iCount + 1, iCount)
          Exit Function
        End If
      End If

    Next iCount

    Output_GetCalArrayIndex = -1

  End Function

  Private Function OutputArray_FillEvents(ByRef plngCalDatIndex As Short, ByRef plngStart As Integer, ByRef plngEnd As Integer) As Boolean

    ' This function actually fills the cal boxes between the indexes specified
    ' according to the options selected by the user.

    On Error GoTo ErrorTrap

    Dim colEvents As clsCalendarEvents

    Dim intCount As Short

    Dim strCurrentRegion_BD As String
    Dim strCurrentWorkingPattern_BD As String

    Dim varTempArray(,) As Object

    Dim intStartCount As Short
    Dim intEndCount As Short

    'UPGRADE_WARNING: Couldn't resolve default property of object mavOutputDateIndex(2, plngCalDatIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    varTempArray = mavOutputDateIndex(2, plngCalDatIndex)

    ' Loop through the indexes as specified.
    For intCount = plngStart To plngEnd Step 1

      'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(6, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      If varTempArray(6, intCount) = 0 Then
        'Date & Session clear
        'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(6, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        varTempArray(6, intCount) = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(7, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        varTempArray(7, intCount) = mstrEventLegend_Output
        'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(8, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        varTempArray(8, intCount) = GetLegendColour(mstrCurrentEventKey)
        'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(9, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        varTempArray(9, intCount) = HexValue(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))

      Else
        'Date & Session already has an event, set it as Multiple.
        'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(6, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        varTempArray(6, intCount) = 2
        'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(7, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        varTempArray(7, intCount) = "."
        'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(8, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        varTempArray(8, intCount) = HexValue(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White))
        'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(9, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        varTempArray(9, intCount) = HexValue(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))

      End If

    Next intCount

    'UPGRADE_WARNING: Couldn't resolve default property of object mavOutputDateIndex(2, plngCalDatIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    mavOutputDateIndex(2, plngCalDatIndex) = VB6.CopyArray(varTempArray)

    OutputArray_FillEvents = True

TidyUpAndExit:
    Exit Function

ErrorTrap:
    OutputArray_FillEvents = False
    GoTo TidyUpAndExit

  End Function

  Private Function HexValue(ByRef plngColour As Integer) As String

    Dim strHEX As String

    strHEX = Hex(plngColour)

    If Len(strHEX) < 6 Then
      strHEX = New String("0", 6 - Len(strHEX)) & strHEX
    End If

    HexValue = "&H" & strHEX

  End Function

  Private Function GetLegendColour(ByRef pstrEventKey As String) As String

    Dim i As Short
    Dim lngTemp As Integer

    For i = 0 To UBound(mavLegend, 2) Step 1
      'UPGRADE_WARNING: Couldn't resolve default property of object mavLegend(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      If UCase(RTrim(mavLegend(0, i))) = UCase(RTrim(pstrEventKey)) Then
        'UPGRADE_WARNING: Couldn't resolve default property of object mavLegend(3, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        lngTemp = mavLegend(3, i)
        GetLegendColour = HexValue(lngTemp)
        Exit Function
      End If
    Next i

    GetLegendColour = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black).ToString

  End Function
  Private Function OutputArray_GetArray() As Boolean

    On Error GoTo ErrorTrap

    Dim fOK As Boolean

    fOK = True

    If fOK Then fOK = OutputArray_AddDays()

    If fOK Then fOK = OutputArray_AddDates()

    If fOK Then fOK = OutputArray_AddCalendar()

    If fOK Then fOK = OutputArray_AddEvents()

    If fOK Then fOK = OutputArray_RefreshDateSpecifics()

    OutputArray_GetArray = True

TidyUpAndExit:
    Exit Function

ErrorTrap:
    OutputArray_GetArray = False
    GoTo TidyUpAndExit

  End Function

  Public Sub SetLastRun()

    Call UtilUpdateLastRun(modUtilAccessLog.UtilityType.utlCalendarReport, mlngCalendarReportID)

  End Sub

  Private Function SetOutputStyles() As Boolean

    Dim intBaseRowCount As Short

    intBaseRowCount = mintBaseRecordCount_Output

    'add merge for the empty top left cells
    'mobjOutput.AddMerge 0, 0, 0, 1

    AddToArray_Merges("0" & vbTab & "0" & vbTab & "0" & vbTab & "1")

    '******************************************************************************
    'add style for the weekend ranges if required

    If ShowWeekends Then
      'first Sunday column (Sunday only)
      '    mobjOutput.AddStyle "", 1, 2, _
      ''                          1, CLng((2 * intBaseRowCount) + 1), _
      ''                            CLng(lblWeekend.BackColor), lblWeekend.ForeColor, False, False, True
      AddToArray_Styles("" & vbTab & "1" & vbTab & "2" & vbTab & "1" & vbTab & CStr(CInt((2 * intBaseRowCount) + 1)) & vbTab & CStr(mlngColor_Weekend) & vbTab & CStr(mlngColor_Weekend) & vbTab & "false" & vbTab & "false" & vbTab & "true")

      'first Sat, second Sunday
      '    mobjOutput.AddStyle "", 7, 2, _
      ''                          8, CLng((2 * intBaseRowCount) + 1), _
      ''                            CLng(lblWeekend.BackColor), lblWeekend.ForeColor, False, False, True
      AddToArray_Styles("" & vbTab & "7" & vbTab & "2" & vbTab & "8" & vbTab & CStr(CInt((2 * intBaseRowCount) + 1)) & vbTab & CStr(mlngColor_Weekend) & vbTab & CStr(mlngColor_Weekend) & vbTab & "false" & vbTab & "false" & vbTab & "true")

      'second Sat, third Sunday
      '    mobjOutput.AddStyle "", 14, 2, _
      ''                          15, CLng((2 * intBaseRowCount) + 1), _
      ''                            CLng(lblWeekend.BackColor), lblWeekend.ForeColor, False, False, True
      AddToArray_Styles("" & vbTab & "14" & vbTab & "2" & vbTab & "15" & vbTab & CStr(CInt((2 * intBaseRowCount) + 1)) & vbTab & CStr(mlngColor_Weekend) & vbTab & CStr(mlngColor_Weekend) & vbTab & "false" & vbTab & "false" & vbTab & "true")

      'third Sat, fourth Sunday
      '    mobjOutput.AddStyle "", 21, 2, _
      ''                          22, CLng((2 * intBaseRowCount) + 1), _
      ''                            CLng(lblWeekend.BackColor), lblWeekend.ForeColor, False, False, True
      AddToArray_Styles("" & vbTab & "21" & vbTab & "2" & vbTab & "22" & vbTab & CStr(CInt((2 * intBaseRowCount) + 1)) & vbTab & CStr(mlngColor_Weekend) & vbTab & CStr(mlngColor_Weekend) & vbTab & "false" & vbTab & "false" & vbTab & "true")

      'fourth Sat, fifth Sunday
      '    mobjOutput.AddStyle "", 28, 2, _
      ''                          29, CLng((2 * intBaseRowCount) + 1), _
      ''                            CLng(lblWeekend.BackColor), lblWeekend.ForeColor, False, False, True
      AddToArray_Styles("" & vbTab & "28" & vbTab & "2" & vbTab & "29" & vbTab & CStr(CInt((2 * intBaseRowCount) + 1)) & vbTab & CStr(mlngColor_Weekend) & vbTab & CStr(mlngColor_Weekend) & vbTab & "false" & vbTab & "false" & vbTab & "true")

      'fifth Sat, sixth Sunday
      '    mobjOutput.AddStyle "", 35, 2, _
      ''                          36, CLng((2 * intBaseRowCount) + 1), _
      ''                            CLng(lblWeekend.BackColor), lblWeekend.ForeColor, False, False, True
      AddToArray_Styles("" & vbTab & "35" & vbTab & "2" & vbTab & "36" & vbTab & CStr(CInt((2 * intBaseRowCount) + 1)) & vbTab & CStr(mlngColor_Weekend) & vbTab & CStr(mlngColor_Weekend) & vbTab & "false" & vbTab & "false" & vbTab & "true")
    End If

    '******************************************************************************


    'add style for the outside of report date boundaries
    'first out of range (if required)
    If (mintRangeStartIndex_Output > 0) Then
      '    mobjOutput.AddStyle "", 1, 2, _
      ''                         CLng(mintRangeStartIndex_Output), CLng((2 * intBaseRowCount) + 1), _
      ''                           CLng(lblRangeDisabled.BackColor), lblRangeDisabled.ForeColor, False, False, True
      AddToArray_Styles("" & vbTab & "1" & vbTab & "2" & vbTab & CStr(CInt(mintRangeStartIndex_Output)) & vbTab & CStr(CInt((2 * intBaseRowCount) + 1)) & vbTab & CStr(mlngColor_RangeDisabled) & vbTab & CStr(mlngColor_RangeDisabled) & vbTab & "false" & vbTab & "false" & vbTab & "true")
    End If

    'second out of range (if required)
    If (mintRangeEndIndex_Output > 0) And (mintRangeEndIndex_Output < 38) Then
      '    mobjOutput.AddStyle "", CLng(mintRangeEndIndex_Output), 2, _
      ''                         37, CLng((2 * intBaseRowCount) + 1), _
      ''                            CLng(lblRangeDisabled.BackColor), lblRangeDisabled.ForeColor, False, False, True
      AddToArray_Styles("" & vbTab & CStr(CInt(mintRangeEndIndex_Output)) & vbTab & "2" & vbTab & "37" & vbTab & CStr(CInt((2 * intBaseRowCount) + 1)) & vbTab & CStr(mlngColor_RangeDisabled) & vbTab & CStr(mlngColor_RangeDisabled) & vbTab & "false" & vbTab & "false" & vbTab & "true")
    End If


    'add style for the disabled ranges
    'first disabled range (if required)
    If (mintFirstDayOfMonth_Output > 1) Then
      '    mobjOutput.AddStyle "", 1, 2, _
      ''                         (mintFirstDayOfMonth_Output - 1), CLng((2 * intBaseRowCount) + 1), _
      ''                           CLng(lblDisabled.BackColor), lblDisabled.ForeColor, False, False, True
      AddToArray_Styles("" & vbTab & "1" & vbTab & "2" & vbTab & CStr(mintFirstDayOfMonth_Output - 1) & vbTab & CStr(CInt((2 * intBaseRowCount) + 1)) & vbTab & CStr(mlngColor_Disabled) & vbTab & CStr(mlngColor_Disabled) & vbTab & "false" & vbTab & "false" & vbTab & "true")
    End If

    'second disabled range (if required)
    If ((mintFirstDayOfMonth_Output + mintDaysInMonth_Output) <= 37) Then
      '    mobjOutput.AddStyle "", (mintFirstDayOfMonth_Output + mintDaysInMonth_Output), 2, _
      ''                         37, CLng((2 * intBaseRowCount) + 1), _
      ''                            CLng(lblDisabled.BackColor), lblDisabled.ForeColor, False, False, True
      AddToArray_Styles("" & vbTab & CStr(mintFirstDayOfMonth_Output + mintDaysInMonth_Output) & vbTab & "2" & vbTab & "37" & vbTab & CStr(CInt((2 * intBaseRowCount) + 1)) & vbTab & CStr(mlngColor_Disabled) & vbTab & CStr(mlngColor_Disabled) & vbTab & "false" & vbTab & "false" & vbTab & "true")
    End If

  End Function

  Private Function OutputArray_GetLegendArray() As Boolean

    On Error GoTo ErrorTrap

    Dim fOK As Boolean
    Dim i As Integer
    Dim iLegendCount As Integer
    Dim iNewIndex As Short

    fOK = True

    iLegendCount = 0

    'add the header row for the Key page
    iNewIndex = 0
    AddToArray_Data(0, 0, "Event Name")
    AddToArray_Data(0, 1, "    ")
    AddToArray_Data(0, 2, " ")
    mlngGridRowIndex = mlngGridRowIndex + 1

    iNewIndex = iNewIndex + 1
    AddToArray_Data(0, 0, "")
    AddToArray_Data(0, 1, "    ")
    AddToArray_Data(0, 2, " ")
    mlngGridRowIndex = mlngGridRowIndex + 1

    For i = 1 To (UBound(mavLegend, 2) * 2) Step 2
      iLegendCount = iLegendCount + 1
      iNewIndex = iNewIndex + 1
      'UPGRADE_WARNING: Couldn't resolve default property of object mavLegend(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      AddToArray_Data(0, 0, CStr(mavLegend(1, iLegendCount)))
      AddToArray_Data(0, 1, "    ")
      'UPGRADE_WARNING: Couldn't resolve default property of object mavLegend(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      AddToArray_Data(0, 2, Replace(mavLegend(2, iLegendCount), "&&", "&"))
      mlngGridRowIndex = mlngGridRowIndex + 1

      iNewIndex = iNewIndex + 1
      AddToArray_Data(0, 0, "")
      AddToArray_Data(0, 1, "    ")
      AddToArray_Data(0, 2, "")
      mlngGridRowIndex = mlngGridRowIndex + 1
    Next i

    iLegendCount = 0
    For i = 1 To (UBound(mavLegend, 2) * 2) Step 2
      iLegendCount = iLegendCount + 1
      '    mobjOutput.AddStyle "", 2, (i + 1), 2, (i + 1), CLng(lblLegend(iLegendCount).BackColor), CLng(lblLegend(iLegendCount).ForeColor), False, False, True
      'UPGRADE_WARNING: Couldn't resolve default property of object mavLegend(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      AddToArray_Styles("" & vbTab & "2" & vbTab & CStr(i + 1) & vbTab & "2" & vbTab & CStr(i + 1) & vbTab & CStr(CInt(mavLegend(3, iLegendCount))) & vbTab & IIf(mblnShowCaptions, CStr(CInt(0)), CStr(CInt(mavLegend(3, iLegendCount)))) & vbTab & "false" & vbTab & "false" & vbTab & "true")
    Next i

    OutputArray_GetLegendArray = True

TidyUpAndExit:
    Exit Function

ErrorTrap:
    OutputArray_GetLegendArray = False
    GoTo TidyUpAndExit

  End Function

  Private Function Load_Legend() As Boolean

    On Error GoTo ErrorTrap

    Dim intNewIndex As Short
    Dim intCount As Short
    Dim lngWidth As Integer

    Dim strEventID As String

    Dim blnNewEvent As Boolean

    Dim intColourIndex As Short
    Dim intColourMax As Short

    Dim lngFC_Data As Integer
    Dim lngBD_Data As Integer
    Dim lngFC_Header As Integer
    Dim lngBC_Header As Integer

    Const LEGEND_COLS As Short = 2

    strEventID = vbNullString

    ReDim mavLegend(3, 0)

    mintLegendCount = 0

    intColourMax = UBound(mavAvailableColours, 2)

    With mrsCalendarReportsOutput
      If Not (.BOF And .EOF) Then

        .MoveFirst()
        Do While Not .EOF
          If strEventID <> .Fields(mstrEventIDColumn).Value Then
            strEventID = .Fields(mstrEventIDColumn).Value

            blnNewEvent = True
            For intCount = 1 To UBound(mavLegend, 2) Step 1
              'UPGRADE_WARNING: Couldn't resolve default property of object mavLegend(0, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              If mavLegend(0, intCount) = strEventID Then
                blnNewEvent = False
              End If
            Next intCount

            If blnNewEvent Then
              intNewIndex = UBound(mavLegend, 2) + 1

              ReDim Preserve mavLegend(3, intNewIndex)
              'UPGRADE_WARNING: Couldn't resolve default property of object mavLegend(0, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              mavLegend(0, intNewIndex) = strEventID
              'UPGRADE_WARNING: Couldn't resolve default property of object mavLegend(1, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              mavLegend(1, intNewIndex) = Left(.Fields("Name").Value, 50)
              'UPGRADE_WARNING: Couldn't resolve default property of object mavLegend(2, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              mavLegend(2, intNewIndex) = Left(.Fields("Legend").Value, 2)

              intColourIndex = (intNewIndex - 1) Mod intColourMax
              'UPGRADE_WARNING: Couldn't resolve default property of object mavAvailableColours(1, intColourIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              'UPGRADE_WARNING: Couldn't resolve default property of object mavLegend(3, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              mavLegend(3, intNewIndex) = mavAvailableColours(1, intColourIndex)

            End If
          End If

          .MoveNext()
        Loop

        ' Sort the Array here - then add the Multiple events item to the end.
        SortLegend(mavLegend, 1)

        If mblnHasMultipleEvents Then
          intNewIndex = UBound(mavLegend, 2) + 1
          ReDim Preserve mavLegend(3, intNewIndex)
          'UPGRADE_WARNING: Couldn't resolve default property of object mavLegend(0, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          mavLegend(0, intNewIndex) = "EVENT_MULTIPLE"
          'UPGRADE_WARNING: Couldn't resolve default property of object mavLegend(1, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          mavLegend(1, intNewIndex) = "Multiple Events"
          'UPGRADE_WARNING: Couldn't resolve default property of object mavLegend(2, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          mavLegend(2, intNewIndex) = "."
          'UPGRADE_WARNING: Couldn't resolve default property of object mavLegend(3, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          mavLegend(3, intNewIndex) = "&HFFFFFF"
        End If

        mintLegendCount = UBound(mavLegend, 2)

      Else
        Load_Legend = False
        GoTo TidyUpAndExit

      End If
    End With

    Load_Legend = True

TidyUpAndExit:
    Exit Function

ErrorTrap:
    mstrErrorString = "Error creating Calendar Report Key."
    Load_Legend = False
    GoTo TidyUpAndExit

  End Function

  Private Function SortLegend(ByRef pavLegend As Object, ByRef pintIndex As Short) As Boolean

    On Error GoTo ErrorTrap

    Dim lngCount As Integer
    Dim lngRestOfArray As Integer
    Dim lngRowIndex As Integer
    Dim intStrComp As Short
    Dim i As Short

    Dim varTemp As Object

    For lngCount = 1 To UBound(pavLegend, 2) Step 1
      lngRowIndex = lngCount

      For lngRestOfArray = (lngCount + 1) To UBound(pavLegend, 2) Step 1
        'UPGRADE_WARNING: Couldn't resolve default property of object pavLegend(pintIndex, lngRestOfArray). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object pavLegend(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        intStrComp = StrComp(pavLegend(pintIndex, lngRowIndex), pavLegend(pintIndex, lngRestOfArray), CompareMethod.Text)
        If intStrComp = 1 Then
          lngRowIndex = lngRestOfArray
        End If
      Next lngRestOfArray

      'put the new lowest in position
      For i = 0 To UBound(pavLegend) Step 1
        'UPGRADE_WARNING: Couldn't resolve default property of object pavLegend(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object varTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        varTemp = pavLegend(i, lngRowIndex)
        'UPGRADE_WARNING: Couldn't resolve default property of object pavLegend(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        pavLegend(i, lngRowIndex) = pavLegend(i, lngCount)
        'UPGRADE_WARNING: Couldn't resolve default property of object varTemp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object pavLegend(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        pavLegend(i, lngCount) = varTemp
      Next i
    Next lngCount

    SortLegend = True

TidyUpAndExit:
    Exit Function

ErrorTrap:
    SortLegend = False
    GoTo TidyUpAndExit

  End Function

  Private Function GetAvailableColours(ByRef pstrExcludedColours As String) As Boolean

    On Error GoTo ErrorTrap

    Dim rsColours As ADODB.Recordset

    Dim intColourCount As Short
    Dim intNextIndex As Short

    intColourCount = 0
    intNextIndex = 0
    ReDim mavAvailableColours(3, intNextIndex)

    Dim strSQL As String

    strSQL = vbNullString
    strSQL = strSQL & "SELECT ASRSysColours.ColOrder, ASRSysColours.ColValue, "
    strSQL = strSQL & "       ASRSysColours.ColDesc, ASRSysColours.WordColourIndex, "
    strSQL = strSQL & "       ASRSysColours.CalendarLegendColour "
    strSQL = strSQL & "FROM ASRSysColours "
    strSQL = strSQL & "WHERE (CalendarLegendColour = 1) "
    strSQL = strSQL & "  AND (ASRSysColours.ColValue NOT IN ( " & pstrExcludedColours & ")) "
    strSQL = strSQL & "ORDER BY ASRSysColours.ColOrder "

    rsColours = datGeneral.GetRecords(strSQL)

    With rsColours
      If .BOF And .EOF Then
        GetAvailableColours = False
        GoTo TidyUpAndExit
      End If

      .MoveFirst()
      Do While Not .EOF
        ReDim Preserve mavAvailableColours(3, intNextIndex)

        'UPGRADE_WARNING: Couldn't resolve default property of object mavAvailableColours(0, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        mavAvailableColours(0, intNextIndex) = .Fields("ColValue").Value
        'UPGRADE_WARNING: Couldn't resolve default property of object mavAvailableColours(1, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        mavAvailableColours(1, intNextIndex) = HexValue(CInt(.Fields("ColValue").Value))
        'UPGRADE_WARNING: Couldn't resolve default property of object mavAvailableColours(2, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        mavAvailableColours(2, intNextIndex) = .Fields("ColDesc").Value
        'UPGRADE_WARNING: Couldn't resolve default property of object mavAvailableColours(3, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        mavAvailableColours(3, intNextIndex) = .Fields("WordColourIndex").Value

        intNextIndex = UBound(mavAvailableColours, 2) + 1

        .MoveNext()
      Loop

    End With
    rsColours.Close()

    GetAvailableColours = True

TidyUpAndExit:
    'UPGRADE_NOTE: Object rsColours may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsColours = Nothing
    Exit Function

ErrorTrap:
    GetAvailableColours = False
    GoTo TidyUpAndExit

  End Function

  Private Function OutputArray_RefreshDateSpecifics() As Boolean

    On Error GoTo ErrorTrap

    Dim intCount As Short

    'following variables used to establish required back & fore color for the label
    Dim blnIsWeekend As Boolean
    Dim blnIsBankHoliday As Boolean
    Dim blnIsWorkingDay As Boolean
    Dim blnIncBankHoliday As Boolean
    Dim blnIncWorkingDays As Boolean
    Dim blnShadeBankHolidays As Boolean
    Dim blnShadeWeekends As Boolean
    Dim blnShowCaptions As Boolean 'different use than blnShowCaption
    Dim blnHasEvent As Boolean
    Dim blnShowCaption As Boolean
    Dim intDefinedColourStyle As Short

    Dim strColour As String
    Dim intThisStartCount As Short
    Dim intThisEndCount As Short
    Dim intNextStartCount As Short
    Dim intNext2StartCount As Short
    Dim intIndexModulus As Short
    Dim intCurrentStartCount As Short
    Dim intCurrentEndCount As Short
    Dim intBaseCount As Short

    Dim strSession As String

    Dim blnNextHasEvent As Boolean
    Dim blnNext2HasEvent As Boolean
    Dim blnPrevHasEvent As Boolean

    Dim intSessionCount As Short

    Dim varTempArray As Object

    Dim strBaseDesc As String
    Dim strBackColour As String
    Dim strForeColour As String
    Dim strCaption As String

    Dim dtConvertedDate As Date

    Dim lngFirstRowIndex As Integer
    Dim lngSecondRowIndex As Integer

    intSessionCount = 0

    If mintBaseRecordCount_Output < 1 Then
      Exit Function
    End If

    'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

    blnIncBankHoliday = mblnIncludeBankHolidays
    blnIncWorkingDays = mblnIncludeWorkingDaysOnly
    blnShadeBankHolidays = mblnShowBankHolidays
    blnShadeWeekends = mblnShowWeekends
    blnShowCaptions = mblnShowCaptions

    '  DebugMSG "OutputArray_RefreshDateSpecifics()"

    SetOutputStyles()

    '  DebugMSG "SetOutputStyles (completed)"

    For intBaseCount = 1 To mintBaseRecordCount_Output Step 1

      'UPGRADE_WARNING: Couldn't resolve default property of object mavOutputDateIndex(1, intBaseCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      strBaseDesc = mavOutputDateIndex(1, intBaseCount)

      'add the Description values to the array
      AddToArray_Data(intBaseCount * 2, 0, strBaseDesc)
      lngFirstRowIndex = mlngGridRowIndex
      mlngGridRowIndex = mlngGridRowIndex + 1

      AddToArray_Data((intBaseCount * 2) + 1, 0, "")
      lngSecondRowIndex = mlngGridRowIndex
      mlngGridRowIndex = mlngGridRowIndex + 1

      'mobjOutput.AddMerge 0, (intBaseCount * 2), 0, ((intBaseCount * 2) + 1)
      AddToArray_Merges("0" & vbTab & CStr(intBaseCount * 2) & vbTab & "0" & vbTab & CStr((intBaseCount * 2) + 1))

      'UPGRADE_WARNING: Couldn't resolve default property of object mavOutputDateIndex(2, intBaseCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      varTempArray = mavOutputDateIndex(2, intBaseCount)

      For intCount = 1 To 74 Step 1

        intSessionCount = intSessionCount + 1

        mlngGridRowIndex = IIf((intSessionCount = 1), lngFirstRowIndex, lngSecondRowIndex)

        'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(1, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If varTempArray(1, intCount) = "  /  /    " Then
          'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          varTempArray(8, intCount) = HexValue(mlngColor_Disabled)
          'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          varTempArray(7, intCount) = ""

          If intSessionCount = 2 Then
            intSessionCount = 0
          End If

        Else
          'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          dtConvertedDate = ConvertCalendarDateToDateFormat(CStr(varTempArray(1, intCount)))
          If (dtConvertedDate >= mdtStartDate) And (dtConvertedDate <= CDate(mdtEndDate)) Then

            'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            blnIsBankHoliday = IIf(varTempArray(3, intCount) = "1", True, False)
            'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            blnIsWeekend = IIf(varTempArray(4, intCount) = "1", True, False)
            'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            strColour = varTempArray(8, intCount)
            'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            blnHasEvent = IIf(varTempArray(6, intCount) > 0, True, False)
            'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            blnIsWorkingDay = IIf(varTempArray(5, intCount) = "1", True, False)

            intDefinedColourStyle = 0 'Default Colour
            '          intDefinedColourStyle = 1   'Weekend/Bank Holiday Colour
            '          intDefinedColourStyle = 2   'Event Key Colour

            If blnHasEvent Then
              'Event
              intDefinedColourStyle = 2

              If (blnIsWorkingDay) Then
                'Event + Working Day

                If (blnIsBankHoliday) And (Not blnIncBankHoliday) And (Not blnShadeBankHolidays) And (Not ((blnIsWeekend) And (blnShadeWeekends))) Then
                  'Event + Working Day + ((Bank Holiday + Not Inc. Working Days Only + Not Shade Bank Holidays)))
                  intDefinedColourStyle = 0
                ElseIf (blnIsBankHoliday) And (Not blnIncBankHoliday) And ((blnShadeBankHolidays) Or ((blnIsWeekend) And (blnShadeWeekends))) Then
                  'Event + Working Day + ((Bank Holiday + Not Inc. Working Days Only + Shade Bank Holidays)))
                  intDefinedColourStyle = 1
                End If

              Else
                'Event + Not Working Day

                If (blnIncWorkingDays) And ((blnIsBankHoliday And Not blnIncBankHoliday) Or (Not blnIsBankHoliday)) And ((blnIsWeekend And Not blnShadeWeekends) Or (Not blnIsWeekend)) Then
                  'Event + Not Working Day + ((Bank Holiday + Not Inc. Working Days Only) || Not Bank Holiday) + ((Weekend + Not Show Weekends) || Not Weekend))
                  intDefinedColourStyle = 0
                End If

                If (blnIsBankHoliday) And (blnShadeBankHolidays) And (blnIncWorkingDays) And (Not blnIncBankHoliday) Then
                  'Event + Not Working Day + Bank Holiday + Shade Bank Holidays + Inc. Working Days Only + Not Inc. Bank Holidays
                  intDefinedColourStyle = 1
                ElseIf (blnIsWeekend) And (blnShadeWeekends) And (blnIncWorkingDays) And (Not blnIncBankHoliday) Then
                  'Event + Not Working Day + Weekend + Show Weekends + Inc. Working Days Only + Not Inc. Bank Holidays
                  intDefinedColourStyle = 1
                ElseIf (blnIsWeekend) And (Not blnIsBankHoliday) And (blnShadeWeekends) And (blnIncWorkingDays) And (blnIncBankHoliday) Then
                  'Event + Not Working Day + Weekend + Show Weekends + Inc. Working Days Only + Inc. Bank Holidays
                  intDefinedColourStyle = 1
                End If

                If (blnIsBankHoliday) And (blnIsWeekend) And (blnShadeWeekends) And (blnIncWorkingDays) And (Not blnIncBankHoliday) Then
                  'Event + Not Working Day + Bank Holiday + Weekend + Show Weekends + Inc. Working Days Only + Not Inc. Bank Holidays
                  intDefinedColourStyle = 1
                End If

                If (blnIsBankHoliday) And (Not blnIncBankHoliday) And (Not blnShadeBankHolidays) And (Not ((blnIsWeekend) And (blnShadeWeekends))) Then
                  'Event + Not Working Day + ((Bank Holiday + Not Inc. Working Days Only + Not Shade Bank Holidays)))
                  intDefinedColourStyle = 0
                ElseIf (blnIsBankHoliday) And (Not blnIncBankHoliday) And ((blnShadeBankHolidays) Or ((blnIsWeekend) And (blnShadeWeekends))) Then
                  'Event + Not Working Day + ((Bank Holiday + Not Inc. Working Days Only + Shade Bank Holidays)))
                  intDefinedColourStyle = 1
                End If

              End If

            Else
              'Not Event
              intDefinedColourStyle = 0

              If (blnIsWeekend) And (blnShadeWeekends) Then
                'Not Event + Weekend + Show Weekends
                intDefinedColourStyle = 1
              End If

              If (blnIsBankHoliday) And (blnShadeBankHolidays) Then
                'Not Event + Bank Holiday + Show Bank Holidays
                intDefinedColourStyle = 1
              End If

            End If

            Select Case intDefinedColourStyle
              Case 0
                'Show the default colour
                'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                varTempArray(8, intCount) = HexValue(mlngBC_Data)
                'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                varTempArray(9, intCount) = HexValue(mlngBC_Data)
                'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strBackColour = varTempArray(8, intCount)
                'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strForeColour = varTempArray(9, intCount)
                blnShowCaption = False

              Case 1
                'Show the Weekend/Bank Holiday colour
                'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                varTempArray(8, intCount) = HexValue(mlngColor_Weekend)
                'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                varTempArray(9, intCount) = HexValue(mlngColor_Weekend)
                'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strBackColour = varTempArray(8, intCount)
                'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strForeColour = varTempArray(9, intCount)
                blnShowCaption = False

              Case 2
                'Show the colour from the Event Key!
                'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                varTempArray(8, intCount) = strColour
                'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                varTempArray(9, intCount) = HexValue(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
                'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strBackColour = varTempArray(8, intCount)
                'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strForeColour = varTempArray(9, intCount)
                blnShowCaption = True

            End Select

            'set key character OR NOT.
            'TM17122003 Faults 7760 & 7761 fixed.
            'if the caption is not to be shown then set the caption to null string
            'rather than hide by making the forecolor the same as the backcolor.
            If ((blnShowCaptions) And (blnShowCaption)) Then
              'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              varTempArray(9, intCount) = HexValue(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black))
              'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              strForeColour = varTempArray(9, intCount)
              'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              strCaption = varTempArray(7, intCount)
            Else
              'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              varTempArray(9, intCount) = varTempArray(8, intCount)
              'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              strForeColour = varTempArray(9, intCount)
              strCaption = vbNullString
            End If

            '          If ((Not blnShowCaption) Or (Not blnShowCaptions)) Then
            '            strCaption = vbNullString
            '          Else
            '            strCaption = varTempArray(7, intCount)
            '          End If

            If intSessionCount = 1 Then

              If blnHasEvent Or ((blnIsBankHoliday) And (blnShadeBankHolidays)) Then
                '              mobjOutput.AddStyle "", CLng((intCount + 1) / 2), CLng(intBaseCount * 2), _
                ''                                    CLng((intCount + 1) / 2), CLng(intBaseCount * 2), _
                ''                                    CLng(varTempArray(8, intCount)), CLng(varTempArray(9, intCount)), False, False, True
                '              DebugMSG "Adding style for " & Format(dtConvertedDate, CALREP_DATEFORMAT) & "..." & "" & vbTab & CStr(CLng((intCount + 1) / 2)) & vbTab & CStr(CLng(intBaseCount * 2)) _
                ''                                & vbTab & CStr(CLng((intCount + 1) / 2)) & vbTab & CStr(CLng(intBaseCount * 2)) _
                ''                                & vbTab & CStr(CLng(varTempArray(8, intCount))) & vbTab & CStr(CLng(varTempArray(9, intCount))) & vbTab & "false" & vbTab & "false" & vbTab & "true"
                'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                AddToArray_Styles("" & vbTab & CStr(CInt((intCount + 1) / 2)) & vbTab & CStr(CInt(intBaseCount * 2)) & vbTab & CStr(CInt((intCount + 1) / 2)) & vbTab & CStr(CInt(intBaseCount * 2)) & vbTab & CStr(CInt(varTempArray(8, intCount))) & vbTab & CStr(CInt(varTempArray(9, intCount))) & vbTab & "false" & vbTab & "false" & vbTab & "true")

              End If

              AddToArray_Data(CShort(intBaseCount * 2), CShort((intCount + 1) / 2), strCaption)

            ElseIf intSessionCount = 2 Then

              If blnHasEvent Or ((blnIsBankHoliday) And (blnShadeBankHolidays)) Then
                '              mobjOutput.AddStyle "", CLng(intCount / 2), CLng((intBaseCount * 2) + 1), _
                ''                                    CLng(intCount / 2), CLng((intBaseCount * 2) + 1), _
                ''                                    CLng(varTempArray(8, intCount)), CLng(varTempArray(9, intCount)), False, False, True
                '              DebugMSG "Adding style for " & Format(dtConvertedDate, CALREP_DATEFORMAT) & "..." & "" & vbTab & CStr(CLng(intCount / 2)) & vbTab & CStr(CLng((intBaseCount * 2) + 1)) _
                ''                                & vbTab & CStr(CLng(intCount / 2)) & vbTab & CStr(CLng((intBaseCount * 2) + 1)) _
                ''                                & vbTab & CStr(CLng(varTempArray(8, intCount))) & vbTab & CStr(CLng(varTempArray(9, intCount))) & vbTab & "false" & vbTab & "false" & vbTab & "true"
                'UPGRADE_WARNING: Couldn't resolve default property of object varTempArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                AddToArray_Styles("" & vbTab & CStr(CInt(intCount / 2)) & vbTab & CStr(CInt((intBaseCount * 2) + 1)) & vbTab & CStr(CInt(intCount / 2)) & vbTab & CStr(CInt((intBaseCount * 2) + 1)) & vbTab & CStr(CInt(varTempArray(8, intCount))) & vbTab & CStr(CInt(varTempArray(9, intCount))) & vbTab & "false" & vbTab & "false" & vbTab & "true")
              End If

              AddToArray_Data(CShort((intBaseCount * 2) + 1), CShort(intCount / 2), strCaption)

              intSessionCount = 0

            End If
          Else
            If intSessionCount = 2 Then
              intSessionCount = 0
            End If
          End If
        End If

      Next intCount

      mlngGridRowIndex = lngSecondRowIndex + 1

    Next intBaseCount

    OutputArray_RefreshDateSpecifics = True

TidyUpAndExit:
    Exit Function

ErrorTrap:
    '  DebugMSG "********** ERROR **************"
    '  DebugMSG Err.Number & " " & Err.Description

    OutputArray_RefreshDateSpecifics = False
    GoTo TidyUpAndExit

  End Function

  Private Function ConvertCalendarDateToDateFormat(ByRef pstrDateString As String) As Date

    Dim dtTemp As Date
    Dim strDateFormat As String
    Dim lngDay_CR As Integer
    Dim lngMonth_CR As Integer
    Dim lngYear_CR As Integer

    Dim blnDateComplete As Boolean
    Dim blnMonthDone As Boolean
    Dim blnDayDone As Boolean
    Dim blnYearDone As Boolean

    Dim strShortDate As String

    Dim strDateSeparator As String

    Dim i As Short

    ' eg. DateFormat = "mm/dd/yyyy"
    '     Calendar   = "dd/mm/yyyy"
    '     DateString = "06/02/2000"
    '     Compare to = 02/06/2000

    strDateFormat = mstrClientDateFormat

    strDateSeparator = CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator

    blnDateComplete = False
    blnMonthDone = False
    blnDayDone = False
    blnYearDone = False

    lngDay_CR = CInt(Mid(pstrDateString, 1, 2))
    lngMonth_CR = CInt(Mid(pstrDateString, 4, 2))
    lngYear_CR = CInt(Mid(pstrDateString, 7, 4))

    strShortDate = vbNullString

    For i = 1 To Len(strDateFormat) Step 1

      If (LCase(Mid(strDateFormat, i, 1)) = "d") And (Not blnDayDone) Then
        strShortDate = strShortDate & LCase(Mid(strDateFormat, i, 1))
        blnDayDone = True
      End If

      If (LCase(Mid(strDateFormat, i, 1)) = "m") And (Not blnMonthDone) Then
        strShortDate = strShortDate & LCase(Mid(strDateFormat, i, 1))
        blnMonthDone = True
      End If

      If (LCase(Mid(strDateFormat, i, 1)) = "y") And (Not blnYearDone) Then
        strShortDate = strShortDate & LCase(Mid(strDateFormat, i, 1))
        blnYearDone = True
      End If

      If blnDayDone And blnMonthDone And blnYearDone Then
        blnDateComplete = True
        Exit For
      End If

    Next i

    Select Case strShortDate
      Case "dmy" : dtTemp = CDate(VB6.Format(lngDay_CR & strDateSeparator & lngMonth_CR & strDateSeparator & lngYear_CR, CALREP_DATEFORMAT))
      Case "mdy" : dtTemp = CDate(lngMonth_CR & strDateSeparator & lngDay_CR & strDateSeparator & lngYear_CR)
      Case "ydm" : dtTemp = CDate(lngYear_CR & strDateSeparator & lngDay_CR & strDateSeparator & lngMonth_CR)
      Case "myd" : dtTemp = CDate(lngMonth_CR & strDateSeparator & lngYear_CR & strDateSeparator & lngDay_CR)
      Case "ymd" : dtTemp = CDate(lngYear_CR & strDateSeparator & lngMonth_CR & strDateSeparator & lngDay_CR)
    End Select

    ConvertCalendarDateToDateFormat = dtTemp

  End Function

  Public Function BaseIndex_Add(ByRef pintCurrentBaseIndex As Short, ByRef plngCurrentRecordID As Integer) As Boolean
    mcolBaseDescIndex.Add(pintCurrentBaseIndex, CStr(plngCurrentRecordID))
  End Function

  Public Function BaseIndex_Get(ByRef pstrCurrentRecordID As String) As Object
    'UPGRADE_WARNING: Couldn't resolve default property of object mcolBaseDescIndex.Item(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    'UPGRADE_WARNING: Couldn't resolve default property of object BaseIndex_Get. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    BaseIndex_Get = mcolBaseDescIndex.Item(pstrCurrentRecordID)
  End Function

  Private Function CheckColumnPermissions(ByRef plngTableID As Integer, ByRef pstrTableName As String, ByRef pstrColumnName As String, ByRef strSQLRef As String) As Boolean

    'This function checks if the current user has read(select) permissions
    'on this column. If the user only has access through views then the
    'relevent views are added to the mlngTableViews() array which in turn
    'are used to create the join part of the query.

    Dim lngTempTableID As Integer
    Dim strTempTableName As String
    Dim strTempColumnName As String
    Dim blnColumnOK As Boolean
    Dim blnFound As Boolean
    Dim blnNoSelect As Boolean
    Dim iLoop1 As Short
    Dim intLoop As Short
    Dim strColumnCode As String
    Dim strSource As String
    Dim intNextIndex As Short
    Dim blnOK As Boolean
    Dim strTable As String
    Dim strColumn As String

    Dim pintNextIndex As Short

    Dim bDateColumn As Boolean

    ' Set flags with their starting values
    blnOK = True
    blnNoSelect = False
    bDateColumn = False

    strTable = vbNullString
    strColumn = vbNullString

    ' Load the temp variables
    lngTempTableID = plngTableID
    strTempTableName = pstrTableName
    strTempColumnName = pstrColumnName

    ' Check permission on that column
    mobjColumnPrivileges = GetColumnPrivileges(strTempTableName)
    mstrRealSource = gcoTablePrivileges.Item(strTempTableName).RealSource

    blnColumnOK = mobjColumnPrivileges.IsValid(strTempColumnName)

    If blnColumnOK Then
      blnColumnOK = mobjColumnPrivileges.Item(strTempColumnName).AllowSelect
    End If

    If mobjColumnPrivileges.Item(strTempColumnName).DataType = Declarations.SQLDataType.sqlDate Then
      bDateColumn = True
    End If

    If blnColumnOK Then
      ' this column can be read direct from the tbl/view or from a parent table
      strTable = mstrRealSource
      strColumn = strTempColumnName

      ' If the table isnt the base table (or its realsource) then
      ' Check if it has already been added to the array. If not, add it.
      If lngTempTableID <> mlngCalendarReportsBaseTable Then
        blnFound = False
        For intNextIndex = 1 To UBound(mlngTableViews, 2)
          If mlngTableViews(1, intNextIndex) = 0 And mlngTableViews(2, intNextIndex) = lngTempTableID Then
            blnFound = True
            Exit For
          End If
        Next intNextIndex

        If Not blnFound Then
          intNextIndex = UBound(mlngTableViews, 2) + 1
          ReDim Preserve mlngTableViews(2, intNextIndex)
          mlngTableViews(1, intNextIndex) = 0
          mlngTableViews(2, intNextIndex) = lngTempTableID
        End If
      End If

      If bDateColumn And mblnCheckingDescColumn Then
        strSQLRef = SQLDateConvertToLocale(strTable & "." & strColumn)
      Else
        strSQLRef = strTable & "." & strColumn
      End If

    Else

      ' this column cannot be read direct. If its from a parent, try parent views
      ' Loop thru the views on the table, seeing if any have read permis for the column

      ReDim mstrViews(0)
      For Each mobjTableView In gcoTablePrivileges.Collection
        If (Not mobjTableView.IsTable) And (mobjTableView.TableID = lngTempTableID) And (mobjTableView.AllowSelect) Then

          strSource = mobjTableView.ViewName
          mstrRealSource = gcoTablePrivileges.Item(strSource).RealSource

          ' Get the column permission for the view
          mobjColumnPrivileges = GetColumnPrivileges(strSource)

          ' If we can see the column from this view
          If mobjColumnPrivileges.IsValid(strTempColumnName) Then
            If mobjColumnPrivileges.Item(strTempColumnName).AllowSelect Then

              ReDim Preserve mstrViews(UBound(mstrViews) + 1)
              mstrViews(UBound(mstrViews)) = mobjTableView.ViewName

              If mlngEventViewColumn > 0 Then
                ReDim Preserve mvarEventColumnViews(1, UBound(mvarEventColumnViews, 2) + 1)
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarEventColumnViews(0, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarEventColumnViews(0, UBound(mvarEventColumnViews, 2)) = mobjTableView.ViewID
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarEventColumnViews(1, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarEventColumnViews(1, UBound(mvarEventColumnViews, 2)) = mlngEventViewColumn
              End If

              ' Check if view has already been added to the array
              blnFound = False
              For intNextIndex = 0 To UBound(mlngTableViews, 2)
                If mlngTableViews(1, intNextIndex) = 1 And mlngTableViews(2, intNextIndex) = mobjTableView.ViewID Then
                  blnFound = True
                  Exit For
                End If
              Next intNextIndex

              If Not blnFound Then
                ' View hasnt yet been added, so add it !
                intNextIndex = UBound(mlngTableViews, 2) + 1
                ReDim Preserve mlngTableViews(2, intNextIndex)
                mlngTableViews(1, intNextIndex) = 1
                mlngTableViews(2, intNextIndex) = mobjTableView.ViewID
                '              Exit For
              Else
                '              Exit For
              End If

            End If
          End If
        End If

      Next mobjTableView

      'UPGRADE_NOTE: Object mobjTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      mobjTableView = Nothing

      ' Does the user have select permission thru ANY views ?
      If UBound(mstrViews) = 0 Then
        blnNoSelect = True
      Else
        strSQLRef = ""
        For pintNextIndex = 1 To UBound(mstrViews)
          If pintNextIndex = 1 Then
            strSQLRef = "CASE"
          End If

          strSQLRef = strSQLRef & " WHEN NOT " & mstrViews(pintNextIndex) & "." & strTempColumnName & " IS NULL THEN "

          If bDateColumn And mblnCheckingDescColumn Then
            strSQLRef = strSQLRef & SQLDateConvertToLocale(mstrViews(pintNextIndex) & "." & strTempColumnName)
          Else
            strSQLRef = strSQLRef & mstrViews(pintNextIndex) & "." & strTempColumnName
          End If

        Next pintNextIndex

        If Len(strSQLRef) > 0 Then
          strSQLRef = strSQLRef & " ELSE NULL" & " END "
        End If

        '      strTable = mstrRealSource
        '      strColumn = strTempColumnName
      End If

      ' If we cant see a column, then get outta here
      If blnNoSelect Then
        strSQLRef = vbNullString
        CheckColumnPermissions = False
        mstrErrorString = vbNewLine & vbNewLine & "You do not have permission to see the column '" & pstrColumnName & "'" & vbNewLine & "either directly or through any views."
        Exit Function
      End If

      If Not blnOK Then
        strSQLRef = vbNullString
        CheckColumnPermissions = False
        Exit Function
      End If

    End If

    'TM01042004 Fault 8428
    If mblnCheckingRegionColumn = True Then
      mstrRegionColumnRealSource = mstrRealSource
    End If

    mlngEventViewColumn = 0

    '  strSQLRef = strTable & "." & strColumn
    CheckColumnPermissions = True

  End Function

  Private Function GenerateSQLEvent(ByRef pstrEventKey As String, ByRef pstrDynamicKey As String, ByRef pstrDynamicName As String) As Boolean

    Dim fOK As Boolean

    fOK = True

    If fOK Then fOK = GenerateSQLSelect(pstrEventKey, pstrDynamicKey, pstrDynamicName)
    If fOK Then fOK = GenerateSQLFrom()
    If fOK Then fOK = GenerateSQLJoin(pstrEventKey, pstrDynamicKey)
    If fOK Then fOK = GenerateSQLWhere(pstrEventKey, pstrDynamicKey, pstrDynamicName)

    If fOK Then
      mstrSQLEvent = mstrSQLSelect & vbNewLine & mstrSQLFrom & vbNewLine & mstrSQLJoin & vbNewLine & mstrSQLWhere & vbNewLine
    End If

    ' reset strings to hold the SQL statement
    mstrSQLSelect = vbNullString
    mstrSQLFrom = vbNullString
    mstrSQLJoin = vbNullString
    mstrSQLWhere = vbNullString

    GenerateSQLEvent = fOK

  End Function

  Public Function GetOrderArray() As Boolean

    On Error GoTo Error_Trap

    Dim rsTemp As ADODB.Recordset

    Dim sSQL As String

    Dim intTemp As Short

    ' Get columns defined as a SortOrder and load into array
    sSQL = "SELECT * FROM ASRSysCalendarReportOrder WHERE " & "CalendarReportID = " & mlngCalendarReportID & " " & "ORDER BY [OrderSequence]"

    rsTemp = datGeneral.GetReadOnlyRecords(sSQL)

    With rsTemp
      If .BOF And .EOF Then
        GetOrderArray = False
        mstrErrorString = "No columns have been defined as a sort order for the specified Calendar Report definition." & vbNewLine & "Please remove this definition and create a new one."
        Exit Function
      End If
      Do Until .EOF
        intTemp = UBound(mvarSortOrder, 2) + 1
        ReDim Preserve mvarSortOrder(2, intTemp)

        'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(0, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        mvarSortOrder(0, intTemp) = .Fields("ColumnID").Value
        'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(1, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        mvarSortOrder(1, intTemp) = datGeneral.GetColumnName(.Fields("ColumnID").Value)
        'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(2, intTemp). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        mvarSortOrder(2, intTemp) = .Fields("OrderType").Value

        .MoveNext()
      Loop
    End With

    GetOrderArray = True

TidyUpAndExit:
    'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsTemp = Nothing
    Exit Function

Error_Trap:
    GetOrderArray = False
    mstrErrorString = "Error whilst retrieving the event details recordsets'." & vbNewLine & Err.Description

  End Function

  Public Function SetPromptedValues(ByRef pavPromptedValues As Object) As Boolean

    ' Purpose : This function calls the individual functions that
    '           generate the components of the main SQL string.
    On Error GoTo ErrorTrap

    Dim fOK As Boolean
    Dim iLoop As Short
    Dim iDataType As Short
    Dim lngComponentID As Integer

    fOK = True

    ReDim mvarPrompts(1, 0)

    If IsArray(pavPromptedValues) Then
      ReDim mvarPrompts(1, UBound(pavPromptedValues, 2))

      For iLoop = 0 To UBound(pavPromptedValues, 2)
        ' Get the prompt data type.
        'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Len(Trim(Mid(pavPromptedValues(0, iLoop), 10))) > 0 Then
          'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          lngComponentID = CInt(Mid(pavPromptedValues(0, iLoop), 10))
          'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          iDataType = CShort(Mid(pavPromptedValues(0, iLoop), 8, 1))

          'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          mvarPrompts(0, iLoop) = lngComponentID

          ' NB. Locale to server conversions are done on the client.
          Select Case iDataType
            Case 2
              ' Numeric.
              'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              mvarPrompts(1, iLoop) = CDbl(pavPromptedValues(1, iLoop))
            Case 3
              ' Logic.
              'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              mvarPrompts(1, iLoop) = (UCase(CStr(pavPromptedValues(1, iLoop))) = "TRUE")
            Case 4
              ' Date.
              ' JPD 20040212 Fault 8082 - DO NOT CONVERT DATE PROMPTED VALUES
              ' THEY ARE PASSED IN FROM THE ASPs AS STRING VALUES IN THE CORRECT
              ' FORMAT (mm/dd/yyyy) AND DOING ANY KIND OF CONVERSION JUST SCREWS
              ' THINGS UP.
              'mvarPrompts(1, iLoop) = CDate(Format(pavPromptedValues(1, iLoop), "mm/dd/yyyy"))
              'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              mvarPrompts(1, iLoop) = pavPromptedValues(1, iLoop)
            Case Else
              'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              mvarPrompts(1, iLoop) = CStr(pavPromptedValues(1, iLoop))
          End Select
        Else
          'UPGRADE_WARNING: Couldn't resolve default property of object mvarPrompts(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          mvarPrompts(0, iLoop) = 0
        End If
      Next iLoop
    End If

    SetPromptedValues = fOK

    Exit Function

ErrorTrap:
    mstrErrorString = "Error setting prompted values." & vbNewLine & Err.Description
    mobjEventLog.AddDetailEntry(mstrErrorString)
    mobjEventLog.ChangeHeaderStatus(clsEventLog.EventLog_Status.elsFailed)
    SetPromptedValues = False

  End Function

  Public Function WorkingPatternTitle() As String

    Dim strTemp As String

    strTemp = vbNullString
    strTemp = strTemp & "<TR align=middle>" & vbNewLine
    strTemp = strTemp & "   <TD ALIGN=center VALIGN=middle></TD>" & vbNewLine
    strTemp = strTemp & "   <TD ALIGN=center VALIGN=middle>" & Left(WeekdayName(1, True, FirstDayOfWeek.Sunday), 1) & "</TD>" & vbNewLine
    strTemp = strTemp & "   <TD ALIGN=center VALIGN=middle>" & Left(WeekdayName(2, True, FirstDayOfWeek.Sunday), 1) & "</TD>" & vbNewLine
    strTemp = strTemp & "   <TD ALIGN=center VALIGN=middle>" & Left(WeekdayName(3, True, FirstDayOfWeek.Sunday), 1) & "</TD>" & vbNewLine
    strTemp = strTemp & "   <TD ALIGN=center VALIGN=middle>" & Left(WeekdayName(4, True, FirstDayOfWeek.Sunday), 1) & "</TD>" & vbNewLine
    strTemp = strTemp & "   <TD ALIGN=center VALIGN=middle>" & Left(WeekdayName(5, True, FirstDayOfWeek.Sunday), 1) & "</TD>" & vbNewLine
    strTemp = strTemp & "   <TD ALIGN=center VALIGN=middle>" & Left(WeekdayName(6, True, FirstDayOfWeek.Sunday), 1) & "</TD>" & vbNewLine
    strTemp = strTemp & "   <TD ALIGN=center VALIGN=middle>" & Left(WeekdayName(7, True, FirstDayOfWeek.Sunday), 1) & "</TD>" & vbNewLine
    strTemp = strTemp & "</TR>" & vbNewLine

    WorkingPatternTitle = strTemp

  End Function

  Public Function Write_Static_Historic_Forms() As String

    Write_Static_Historic_Forms = mstrWPFormString & vbNewLine & vbNewLine

    Write_Static_Historic_Forms = Write_Static_Historic_Forms & mstrBHolFormString & vbNewLine & vbNewLine

    Write_Static_Historic_Forms = Write_Static_Historic_Forms & mstrRegionFormString & vbNewLine & vbNewLine

  End Function

  'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
  Private Sub Class_Initialize_Renamed()

    ' Purpose : Sets references to other classes and redimensions arrays
    '           used for table usage information

    mclsData = New clsDataAccess
    mclsGeneral = New clsGeneral
    mclsUI = New clsUI
    mobjEventLog = New clsEventLog
    mcolBaseDescIndex = New Collection

    ReDim mvarSortOrder(2, 0)
    ReDim mlngTableViews(2, 0)
    ReDim mstrViews(0)
    ReDim mastrUDFsRequired(0)
    ReDim mvarTableViews(3, 0)

    ReDim mvarOutputArray_Definition(0)
    ReDim mvarOutputArray_Columns(0)
    ReDim mvarOutputArray_Data(0)
    ReDim mvarOutputArray_Styles(0)
    ReDim mvarOutputArray_Merges(0)

    ReDim mvarEventColumnViews(1, 0)

  End Sub
  Public Sub New()
    MyBase.New()
    Class_Initialize_Renamed()
  End Sub

  Private Function IsColumnInView(ByRef plngViewID As Integer, ByRef plngColumnID As Integer) As Boolean

    Dim lngCount As Integer

    For lngCount = 1 To UBound(mvarEventColumnViews, 2) Step 1
      'UPGRADE_WARNING: Couldn't resolve default property of object mvarEventColumnViews(1, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      'UPGRADE_WARNING: Couldn't resolve default property of object mvarEventColumnViews(0, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      If (mvarEventColumnViews(0, lngCount) = plngViewID) And (mvarEventColumnViews(1, lngCount) = plngColumnID) Then
        IsColumnInView = True
        Exit Function
      End If
    Next lngCount

  End Function

  'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
  Private Sub Class_Terminate_Renamed()

    ' Purpose : Clears references to other classes.
    'UPGRADE_NOTE: Object mclsData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mclsData = Nothing
    'UPGRADE_NOTE: Object mclsGeneral may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mclsGeneral = Nothing
    'Set mfrmOutput = Nothing
    'UPGRADE_NOTE: Object mcolEvents may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mcolEvents = Nothing
    'UPGRADE_NOTE: Object mobjTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mobjTableView = Nothing
    'UPGRADE_NOTE: Object mobjColumnPrivileges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mobjColumnPrivileges = Nothing
    'UPGRADE_NOTE: Object mcolBaseDescIndex may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mcolBaseDescIndex = Nothing

    'UPGRADE_NOTE: Object mcolHistoricBankHolidays may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mcolHistoricBankHolidays = Nothing
    'UPGRADE_NOTE: Object mcolStaticBankHolidays may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mcolStaticBankHolidays = Nothing
    'UPGRADE_NOTE: Object mcolHistoricWorkingPatterns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mcolHistoricWorkingPatterns = Nothing
    'UPGRADE_NOTE: Object mcolStaticWorkingPatterns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mcolStaticWorkingPatterns = Nothing

    ClearUp()

  End Sub
  Protected Overrides Sub Finalize()
    Class_Terminate_Renamed()
    MyBase.Finalize()
  End Sub

  Public Function ExecuteSql() As Boolean

    ' Purpose : This function executes the SQL string 'into' a recordset.

    On Error GoTo ExecuteSQL_ERROR

    '  'get all the base & event data into a recordset
    mstrSQL = vbNullString
    mstrSQL = mstrSQL & "SELECT * FROM [" & mstrTempTableName & "] "

    'get the ORDER BY statement which applies to the entire UNIONed query.
    GenerateSQLOrderBy()
    mstrSQL = mstrSQL & mstrSQLOrderBy

    mrsCalendarReportsOutput = mclsData.OpenRecordset(mstrSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

    If mrsCalendarReportsOutput.BOF And mrsCalendarReportsOutput.EOF Then
      ExecuteSql = False
      mstrErrorString = "No records meet selection criteria."
      mblnNoRecords = True
      mobjEventLog.ChangeHeaderStatus(clsEventLog.EventLog_Status.elsSuccessful)
      mobjEventLog.AddDetailEntry(mstrErrorString)
      Exit Function
    End If

    MultipleCheck()

    'get only the base table info into a recordset
    mrsCalendarBaseInfo = mclsData.OpenRecordset(mstrSQLBaseData, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

    If mrsCalendarBaseInfo.BOF And mrsCalendarBaseInfo.EOF Then
      ExecuteSql = False
      mstrErrorString = "No records meet selection criteria."
      mblnNoRecords = True
      mobjEventLog.ChangeHeaderStatus(clsEventLog.EventLog_Status.elsSuccessful)
      mobjEventLog.AddDetailEntry(mstrErrorString)
      Exit Function
    End If

    GetDescriptionDataTypes()

    'TM08102003
    UDFFunctions(mastrUDFsRequired, False)

    ExecuteSql = True
    Exit Function

ExecuteSQL_ERROR:

    mstrErrorString = "Error whilst executing SQL statement." & vbNewLine & Err.Description
    ExecuteSql = False

  End Function

  Private Function MultipleCheck() As Boolean

    Dim rsMultiple As ADODB.Recordset
    Dim sSQL As String
    Dim dtSD As Date
    Dim dtED As Date
    Dim strStartSession As String
    Dim strEndSession As String
    Dim lngBaseID As Integer
    Dim strDescription1 As String
    Dim strDescription2 As String
    Dim strDescriptionExpr As String
    Dim lngCurrentBaseID As Integer
    Dim avDateRanges(,) As Object
    Dim i As Short
    Dim blnHasOverlap As Boolean
    Dim intNewIndex As Short
    Dim strFullDesc As String
    Dim strCurrentDesc As String
    Dim blnFirstCalendarRecord As Boolean

    blnFirstCalendarRecord = True

    ReDim avDateRanges(6, 0)

    sSQL = vbNullString
    sSQL = sSQL & "SELECT [BaseID], [Description1], [Description2], [DescriptionExpr], [StartDate], [StartSession], [EndDate], [EndSession] " & vbNewLine
    sSQL = sSQL & "FROM [" & gsUsername & "].[" & mstrTempTableName & "]" & vbNewLine
    sSQL = sSQL & mstrSQLOrderBy

    rsMultiple = datGeneral.GetReadOnlyRecords(sSQL)

    If Not rsMultiple Is Nothing Then
      With rsMultiple
        If Not (.BOF And .EOF) Then
          Do Until .EOF
            dtSD = .Fields("StartDate").Value
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            dtED = IIf(IsDBNull(.Fields("EndDate").Value), .Fields("StartDate").Value, .Fields("EndDate").Value)
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            strStartSession = IIf(IsDBNull(.Fields("StartSession").Value), "AM", .Fields("StartSession").Value)
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            If ((IsDBNull(.Fields("EndDate").Value)) And (IsDBNull(.Fields("EndSession").Value)) And (IsDBNull(.Fields("StartSession").Value))) Then
              strEndSession = "PM"
              'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            ElseIf ((IsDBNull(.Fields("EndDate").Value)) And (IsDBNull(.Fields("EndSession").Value))) Then
              strEndSession = strStartSession
              'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            ElseIf IsDBNull(.Fields("EndSession").Value) Then
              strEndSession = "PM"
            Else
              strEndSession = .Fields("EndSession").Value
            End If

            If mblnGroupByDescription Then
              'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
              strDescription1 = IIf(IsDBNull(.Fields("Description1").Value), "", .Fields("Description1").Value)
              'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
              strDescription2 = IIf(IsDBNull(.Fields("Description2").Value), "", .Fields("Description2").Value)
              'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
              strDescriptionExpr = IIf(IsDBNull(.Fields("DescriptionExpr").Value), "", .Fields("DescriptionExpr").Value)
              strFullDesc = strDescription1 & mstrDescriptionSeparator & strDescription2 & mstrDescriptionSeparator & strDescriptionExpr

              If (strFullDesc <> strCurrentDesc) Or blnFirstCalendarRecord Then
                strCurrentDesc = strFullDesc
                blnFirstCalendarRecord = False

                ReDim avDateRanges(6, 0)
                'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(0, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                avDateRanges(0, 0) = strFullDesc
                'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(1, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                avDateRanges(1, 0) = dtSD
                'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(2, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                avDateRanges(2, 0) = dtED
                'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(3, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                avDateRanges(3, 0) = strStartSession
                'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(4, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                avDateRanges(4, 0) = strEndSession

              Else
                'Loop through the array for the current calendar row, checking if any dates overlap.
                For i = 0 To UBound(avDateRanges, 2) Step 1

                  'if the start or end dates 'equal' any other start orend dates then check if the sessions are also equal.
                  'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(4, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                  'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(2, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                  'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(3, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                  'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(1, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                  If ((dtSD = avDateRanges(1, i)) And (strStartSession = avDateRanges(3, i))) Or ((dtSD = avDateRanges(2, i)) And (strStartSession = avDateRanges(4, i))) Or ((dtED = avDateRanges(1, i)) And (strEndSession = avDateRanges(3, i))) Or ((dtED = avDateRanges(2, i)) And (strEndSession = avDateRanges(4, i))) Then
                    mblnHasMultipleEvents = True
                    MultipleCheck = True
                    GoTo TidyUpAndExit
                  End If

                  'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(2, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                  'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(1, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                  If ((dtSD > avDateRanges(1, i)) And (dtED < avDateRanges(2, i))) Or ((dtSD > avDateRanges(1, i)) And (dtSD < avDateRanges(2, i)) And (dtED > avDateRanges(2, i))) Or ((dtED > avDateRanges(1, i)) And (dtED < avDateRanges(2, i)) And (dtSD < avDateRanges(1, i))) Or ((dtSD < avDateRanges(1, i)) And (dtED > avDateRanges(2, i))) Then
                    mblnHasMultipleEvents = True
                    MultipleCheck = True
                    GoTo TidyUpAndExit
                  End If
                Next i

                intNewIndex = UBound(avDateRanges, 2) + 1
                ReDim Preserve avDateRanges(6, intNewIndex)
                'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(0, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                avDateRanges(0, intNewIndex) = lngBaseID
                'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(1, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                avDateRanges(1, intNewIndex) = dtSD
                'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(2, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                avDateRanges(2, intNewIndex) = dtED
                'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(3, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                avDateRanges(3, intNewIndex) = strStartSession
                'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(4, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                avDateRanges(4, intNewIndex) = strEndSession

              End If
            Else
              lngBaseID = .Fields("BaseID").Value

              If (lngBaseID <> lngCurrentBaseID) Or blnFirstCalendarRecord Then
                lngCurrentBaseID = lngBaseID
                blnFirstCalendarRecord = False

                ReDim avDateRanges(6, 0)
                'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(0, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                avDateRanges(0, 0) = lngBaseID
                'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(1, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                avDateRanges(1, 0) = dtSD
                'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(2, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                avDateRanges(2, 0) = dtED
                'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(3, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                avDateRanges(3, 0) = strStartSession
                'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(4, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                avDateRanges(4, 0) = strEndSession

              Else
                'Loop through the array for the current calendar row, checking if any dates overlap.
                For i = 0 To UBound(avDateRanges, 2) Step 1

                  'if the start or end dates 'equal' any other start orend dates then check if the sessions are also equal.
                  'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(4, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                  'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(2, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                  'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(3, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                  'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(1, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                  If ((dtSD = avDateRanges(1, i)) And (strStartSession = avDateRanges(3, i))) Or ((dtSD = avDateRanges(2, i)) And (strStartSession = avDateRanges(4, i))) Or ((dtED = avDateRanges(1, i)) And (strEndSession = avDateRanges(3, i))) Or ((dtED = avDateRanges(2, i)) And (strEndSession = avDateRanges(4, i))) Then
                    mblnHasMultipleEvents = True
                    MultipleCheck = True
                    GoTo TidyUpAndExit
                  End If

                  'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(2, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                  'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(1, i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                  If ((dtSD > avDateRanges(1, i)) And (dtED < avDateRanges(2, i))) Or ((dtSD > avDateRanges(1, i)) And (dtSD < avDateRanges(2, i)) And (dtED > avDateRanges(2, i))) Or ((dtED > avDateRanges(1, i)) And (dtED < avDateRanges(2, i)) And (dtSD < avDateRanges(1, i))) Or ((dtSD < avDateRanges(1, i)) And (dtED > avDateRanges(2, i))) Then
                    mblnHasMultipleEvents = True
                    MultipleCheck = True
                    GoTo TidyUpAndExit
                  End If
                Next i

                intNewIndex = UBound(avDateRanges, 2) + 1
                ReDim Preserve avDateRanges(6, intNewIndex)
                'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(0, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                avDateRanges(0, intNewIndex) = lngBaseID
                'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(1, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                avDateRanges(1, intNewIndex) = dtSD
                'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(2, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                avDateRanges(2, intNewIndex) = dtED
                'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(3, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                avDateRanges(3, intNewIndex) = strStartSession
                'UPGRADE_WARNING: Couldn't resolve default property of object avDateRanges(4, intNewIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                avDateRanges(4, intNewIndex) = strEndSession

              End If
            End If

            .MoveNext()
          Loop
        End If
        .Close()
      End With
    End If

    mblnHasMultipleEvents = False

    MultipleCheck = True

TidyUpAndExit:
    'UPGRADE_NOTE: Object rsMultiple may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsMultiple = Nothing
    Exit Function

ErrorTrap:
    MultipleCheck = False
    GoTo TidyUpAndExit

  End Function

  Public Function GetCalendarReportDefinition() As Boolean

    ' Purpose : This function retrieves the basic definition details
    '           and stores it in module level variables

    On Error GoTo Error_Trap

    Dim rsTemp As ADODB.Recordset

    Dim sSQL As String
    Dim sTable As String
    Dim sColumn As String
    Dim sDateInterval As String

    Dim i As Short

    Dim rsIDs As ADODB.Recordset
    Dim blnOK As Boolean

    Dim intExprReturnType As Short

    mstrSQLIDs = vbNullString

    sSQL = "SELECT * FROM ASRSYSCalendarReports " & "WHERE ID = " & CStr(mlngCalendarReportID) & " "

    rsTemp = mclsData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

    Dim pblnOK As Object
    Dim objTableView As CTablePrivilege
    Dim objExpression As clsExprExpression
    With rsTemp

      If .BOF And .EOF Then
        GetCalendarReportDefinition = False
        mstrErrorString = "Could not find specified Calendar Report definition."
        GoTo TidyUpAndExit
      End If

      'JPD 20040729 Fault 8972 & Fault 8990
      If LCase(.Fields("Username").Value.ToString()) <> LCase(gsUsername) And CurrentUserAccess(modUtilAccessLog.UtilityType.utlCalendarReport, mlngCalendarReportID) = ACCESS_HIDDEN Then
        GetCalendarReportDefinition = False
        mstrErrorString = "Report has been made hidden by another user."
        Exit Function
      End If

      mstrCalendarReportsName = .Fields("Name").Value
      mobjEventLog.AddHeader(clsEventLog.EventLog_Type.eltCalandarReport, mstrCalendarReportsName)
      mlngCalendarReportsBaseTable = .Fields("BaseTable").Value
      mstrCalendarReportsBaseTableName = datGeneral.GetTableName(mlngCalendarReportsBaseTable)

      ' Check the user has permission to read the base table.
      'UPGRADE_WARNING: Couldn't resolve default property of object pblnOK. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      pblnOK = False
      For Each objTableView In gcoTablePrivileges.Collection
        If (objTableView.TableID = mlngCalendarReportsBaseTable) And (objTableView.AllowSelect) Then
          'UPGRADE_WARNING: Couldn't resolve default property of object pblnOK. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          pblnOK = True
          Exit For
        End If
      Next objTableView
      'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      objTableView = Nothing
      If Not pblnOK Then
        GetCalendarReportDefinition = False
        mstrErrorString = "You do not have permission to read the base table" & vbNewLine & "either directly or through any views."
        GoTo TidyUpAndExit
      End If


      mlngCalendarReportsAllRecords = .Fields("AllRecords").Value
      mlngCalendarReportsPickListID = .Fields("picklist").Value
      mlngCalendarReportsFilterID = .Fields("Filter").Value

      'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
      mlngDescription1 = IIf(IsDBNull(.Fields("Description1").Value), 0, .Fields("Description1").Value)
      If mlngDescription1 > 0 Then
        mstrDescription1 = datGeneral.GetColumnName(.Fields("Description1").Value)
        mblnDesc1IsDate = (datGeneral.GetDataType(mlngCalendarReportsBaseTable, mlngDescription1) = Declarations.SQLDataType.sqlDate)
      Else
        mstrDescription1 = vbNullString
        mblnDesc1IsDate = False
      End If

      'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
      mlngDescription2 = IIf(IsDBNull(.Fields("Description2").Value), 0, .Fields("Description2").Value)
      If mlngDescription2 > 0 Then
        mstrDescription2 = datGeneral.GetColumnName(.Fields("Description2").Value)
        mblnDesc2IsDate = (datGeneral.GetDataType(mlngCalendarReportsBaseTable, mlngDescription2) = Declarations.SQLDataType.sqlDate)
      Else
        mstrDescription2 = vbNullString
        mblnDesc2IsDate = False
      End If

      'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
      mlngDescriptionExpr = IIf(IsDBNull(.Fields("DescriptionExpr").Value), 0, .Fields("DescriptionExpr").Value)
      If mlngDescriptionExpr > 0 Then

        objExpression = New clsExprExpression
        objExpression.ExpressionID = mlngDescriptionExpr
        objExpression.ConstructExpression()
        objExpression.ValidateExpression(True)
        If objExpression.ReturnType = 4 Then ' its date
          mblnDescExprIsDate = True
        Else
          mblnDescExprIsDate = False
        End If
        mlngBaseDescriptionType = objExpression.ReturnType
        mstrDescriptionExpr = objExpression.Name
      Else
        mlngBaseDescriptionType = -1
        mstrDescriptionExpr = vbNullString
        mblnDescExprIsDate = False
      End If
      'UPGRADE_NOTE: Object objExpression may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      objExpression = Nothing

      mlngRegion = .Fields("Region").Value
      If mlngRegion > 0 Then
        mstrRegion = datGeneral.GetColumnName(.Fields("Region").Value)

      ElseIf (mlngCalendarReportsBaseTable = glngPersonnelTableID) And (modPersonnelSpecifics.grtRegionType = modPersonnelSpecifics.RegionType.rtStaticRegion) Then

        mlngRegion = glngBHolRegionID
        mstrRegion = gsBHolRegionColumnName

      Else
        mstrRegion = vbNullString

      End If

      mblnGroupByDescription = IIf(.Fields("GroupByDesc").Value, True, False)
      'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
      mstrDescriptionSeparator = IIf(IsDBNull(.Fields("DescriptionSeparator").Value), " ", .Fields("DescriptionSeparator").Value)

      'create the events collection here so that the event filters can bee checked
      If Not GetEventsCollection() Then
        GetCalendarReportDefinition = False
        GoTo TidyUpAndExit
      End If

      'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
      mlngStartDateExpr = IIf(IsDBNull(.Fields("StartDateExpr").Value), 0, .Fields("StartDateExpr").Value)
      'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
      mlngEndDateExpr = IIf(IsDBNull(.Fields("EndDateExpr").Value), 0, .Fields("EndDateExpr").Value)

      'validate all filters/picklist & calculationshere before, actually using them
      mblnDefinitionOwner = (LCase(Trim(gsUsername)) = LCase(Trim(.Fields("Username").Value)))
      If Not IsRecordSelectionValid() Then
        GetCalendarReportDefinition = False
        GoTo TidyUpAndExit
      End If

      '************** Must do the dates stuff here *****************
      'calculate and store the start and end dates

      'START DATE
      Select Case .Fields("StartType").Value
        Case 0
          mdtStartDate = .Fields("FixedStart").Value
        Case 1
          'JPD 20041119 Faults 9510 & 9511
          'mdtStartDate = CDate(Format(Now, mstrClientDateFormat))
          mdtStartDate = Today
        Case 3
          'UPGRADE_WARNING: Couldn't resolve default property of object datGeneral.GetValueForRecordIndependantCalc(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          mdtStartDate = datGeneral.GetValueForRecordIndependantCalc(mlngStartDateExpr, mvarPrompts)
      End Select

      'END DATE
      Select Case .Fields("EndType").Value
        Case 0
          mdtEndDate = .Fields("FixedEnd").Value
        Case 1
          'JPD 20041119 Faults 9510 & 9511
          'mdtEndDate = CDate(Format(Now, mstrClientDateFormat))
          mdtEndDate = CStr(Today)
        Case 3
          'UPGRADE_WARNING: Couldn't resolve default property of object datGeneral.GetValueForRecordIndependantCalc(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          mdtEndDate = datGeneral.GetValueForRecordIndependantCalc(mlngEndDateExpr, mvarPrompts)
      End Select


      If (.Fields("StartType").Value = 2) And (.Fields("EndType").Value = 2) Then
        'START DATE
        Select Case .Fields("StartPeriod").Value
          Case 0 : sDateInterval = "d"
          Case 1 : sDateInterval = "ww"
          Case 2 : sDateInterval = "m"
          Case 3 : sDateInterval = "yyyy"
        End Select
        'JPD 20041119 Faults 9510 & 9511
        'mdtStartDate = DateAdd(sDateInterval, CDbl(!StartFrequency), CDate(Format(Now, mstrClientDateFormat)))
        mdtStartDate = DateAdd(sDateInterval, CDbl(.Fields("StartFrequency").Value), Today)

        'END DATE
        Select Case .Fields("EndPeriod").Value
          Case 0 : sDateInterval = "d"
          Case 1 : sDateInterval = "ww"
          Case 2 : sDateInterval = "m"
          Case 3 : sDateInterval = "yyyy"
        End Select
        'JPD 20041119 Faults 9510 & 9511
        'mdtEndDate = DateAdd(sDateInterval, CDbl(!EndFrequency), CDate(Format(Now, mstrClientDateFormat)))
        mdtEndDate = CStr(DateAdd(sDateInterval, CDbl(.Fields("EndFrequency").Value), Today))

      ElseIf .Fields("StartType").Value = 2 And .Fields("EndType").Value <> 2 Then
        'START DATE
        Select Case .Fields("StartPeriod").Value
          Case 0 : sDateInterval = "d"
          Case 1 : sDateInterval = "ww"
          Case 2 : sDateInterval = "m"
          Case 3 : sDateInterval = "yyyy"
        End Select
        mdtStartDate = DateAdd(sDateInterval, CDbl(.Fields("StartFrequency").Value), CDate(mdtEndDate))

      ElseIf .Fields("EndType").Value = 2 And .Fields("StartType").Value <> 2 Then
        'END DATE
        Select Case .Fields("EndPeriod").Value
          Case 0 : sDateInterval = "d"
          Case 1 : sDateInterval = "ww"
          Case 2 : sDateInterval = "m"
          Case 3 : sDateInterval = "yyyy"
        End Select
        mdtEndDate = CStr(DateAdd(sDateInterval, CDbl(.Fields("EndFrequency").Value), mdtStartDate))

      End If

      If mdtStartDate > CDate(mdtEndDate) Then
        mstrErrorString = "The report end date is before the report start date."
        GetCalendarReportDefinition = False
        GoTo TidyUpAndExit
      End If

      '************************************************

      mblnShowBankHolidays = .Fields("ShowBankHolidays").Value
      mblnShowCaptions = .Fields("ShowCaptions").Value
      mblnShowWeekends = .Fields("ShowWeekends").Value
      mbStartOnCurrentMonth = .Fields("StartOnCurrentMonth").Value
      mblnIncludeWorkingDaysOnly = .Fields("IncludeWorkingDaysOnly").Value
      mblnIncludeBankHolidays = .Fields("IncludeBankHolidays").Value
      mblnCustomReportsPrintFilterHeader = .Fields("PrintFilterHeader").Value

      mblnOutputPreview = .Fields("OutputPreview").Value
      mlngOutputFormat = .Fields("OutputFormat").Value
      mblnOutputScreen = .Fields("OutputScreen").Value
      mblnOutputPrinter = .Fields("OutputPrinter").Value
      mstrOutputPrinterName = .Fields("OutputPrinterName").Value
      mblnOutputSave = .Fields("OutputSave").Value
      mlngOutputSaveExisting = .Fields("OutputSaveExisting").Value
      mblnOutputEmail = .Fields("OutputEmail").Value
      mlngOutputEmailID = .Fields("OutputEmailAddr").Value
      mstrOutputEmailName = GetEmailGroupName(.Fields("OutputEmailAddr").Value)
      mstrOutputEmailSubject = .Fields("OutputEmailSubject").Value
      'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
      mstrOutputEmailAttachAs = IIf(IsDBNull(.Fields("OutputEmailAttachAs").Value), vbNullString, .Fields("OutputEmailAttachAs").Value)
      mstrOutputFilename = .Fields("OutputFilename").Value

      mblnPersonnelBase = (mlngCalendarReportsBaseTable = glngPersonnelTableID)

      If mblnCustomReportsPrintFilterHeader And (mlngSingleRecordID < 1) Then
        If (mlngCalendarReportsFilterID > 0) Then
          mstrCalendarReportsName = mstrCalendarReportsName & " (Base Table filter : " & datGeneral.GetFilterName(mlngCalendarReportsFilterID) & ")"
        ElseIf (mlngCalendarReportsPickListID > 0) Then
          mstrCalendarReportsName = mstrCalendarReportsName & " (Base Table picklist : " & datGeneral.GetPicklistName(mlngCalendarReportsPickListID) & ")"
        End If
      End If

      If mlngSingleRecordID > 0 Then
        'DebugMSG "Single Record ID = " & CStr(mlngSingleRecordID), True
        mstrSQLIDs = CStr(mlngSingleRecordID)

      ElseIf mlngCalendarReportsPickListID > 0 Then
        rsIDs = mclsData.OpenRecordset("EXEC sp_ASRGetPickListRecords " & mlngCalendarReportsPickListID, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

        If rsIDs.BOF And rsIDs.EOF Then
          mobjEventLog.ChangeHeaderStatus(clsEventLog.EventLog_Status.elsSuccessful)
          mobjEventLog.AddDetailEntry(mstrErrorString)
          mstrErrorString = "The selected picklist contains no records."
          GetCalendarReportDefinition = False
          GoTo TidyUpAndExit
        End If

        Do While Not rsIDs.EOF
          mstrSQLIDs = mstrSQLIDs & IIf(Len(mstrSQLIDs) > 0, ", ", "") & rsIDs.Fields(0).Value
          rsIDs.MoveNext()
        Loop
        rsIDs.Close()
        'UPGRADE_NOTE: Object rsIDs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rsIDs = Nothing

      ElseIf mlngCalendarReportsFilterID > 0 Then
        blnOK = datGeneral.FilteredIDs(mlngCalendarReportsFilterID, mstrFilteredIDs, mvarPrompts)

        ' Generate any UDFs that are used in this filter
        If blnOK And gbEnableUDFFunctions Then
          datGeneral.FilterUDFs(mlngCalendarReportsFilterID, mastrUDFsRequired)
        End If

        If blnOK Then
          blnOK = UDFFunctions(mastrUDFsRequired, True)
          If blnOK Then
            rsIDs = mclsData.OpenRecordset(mstrFilteredIDs, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
          End If

          If rsIDs.BOF And rsIDs.EOF Then
            GetCalendarReportDefinition = False
            mstrErrorString = "The base table filter returned no records."
            mobjEventLog.ChangeHeaderStatus(clsEventLog.EventLog_Status.elsSuccessful)
            mobjEventLog.AddDetailEntry(mstrErrorString)
            mblnNoRecords = True
            GoTo TidyUpAndExit
          End If

          Do While Not rsIDs.EOF
            mstrSQLIDs = mstrSQLIDs & IIf(Len(mstrSQLIDs) > 0, ", ", "") & rsIDs.Fields(0).Value
            rsIDs.MoveNext()
          Loop
          rsIDs.Close()
          'UPGRADE_NOTE: Object rsIDs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
          rsIDs = Nothing

          blnOK = UDFFunctions(mastrUDFsRequired, False)

        Else
          ' Permission denied on something in the filter.
          mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(mlngCalendarReportsFilterID) & "' filter."
          mobjEventLog.ChangeHeaderStatus(clsEventLog.EventLog_Status.elsSuccessful)
          mobjEventLog.AddDetailEntry(mstrErrorString)
          GetCalendarReportDefinition = False
          GoTo TidyUpAndExit
        End If

      End If

    End With
    rsTemp.Close()

    mstrBaseIDColumn = "?ID_" & mstrCalendarReportsBaseTableName
    mstrEventIDColumn = "?ID_EventID"

    mstrBaseTableRealSource = gcoTablePrivileges.Item(mstrCalendarReportsBaseTableName).RealSource

    GetCalendarReportDefinition = True

TidyUpAndExit:
    'UPGRADE_NOTE: Object objExpression may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    objExpression = Nothing
    'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsTemp = Nothing
    Exit Function

Error_Trap:
    '  DebugMSG "*************  ERROR : " & Err.Number & ":" & Err.Description & " *******************", False
    GetCalendarReportDefinition = False
    mstrErrorString = "Error whilst reading the Calendar Report definition." & vbNewLine & Err.Description
    Resume TidyUpAndExit

  End Function

  Public Function Initialise_WP_Region() As Boolean

    Dim fOK As Boolean
    Dim blnRegionEnabled As Boolean
    Dim blnWorkingPatternEnabled As Boolean

    mcolHistoricBankHolidays = New Collection
    mcolStaticBankHolidays = New Collection
    mcolHistoricWorkingPatterns = New Collection
    mcolStaticWorkingPatterns = New Collection

    fOK = True

    blnRegionEnabled = False
    blnWorkingPatternEnabled = False

    If (fOK And mblnPersonnelBase And (modPersonnelSpecifics.grtRegionType = modPersonnelSpecifics.RegionType.rtHistoricRegion) And (Not mblnGroupByDescription) And (mlngRegion < 1)) Or (fOK And ((mlngRegion > 0) Or (mblnPersonnelBase And (modPersonnelSpecifics.grtRegionType = modPersonnelSpecifics.RegionType.rtStaticRegion))) And (Not mblnGroupByDescription)) Then

      blnRegionEnabled = CheckPermission_RegionInfo()
    End If

    If blnRegionEnabled Then
      If fOK And mblnPersonnelBase And (modPersonnelSpecifics.grtRegionType = modPersonnelSpecifics.RegionType.rtHistoricRegion) And (Not mblnGroupByDescription) And (mlngRegion < 1) Then

        'get historical bank holidays
        'UPGRADE_WARNING: Couldn't resolve default property of object Get_HistoricBankHolidays. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        fOK = Get_HistoricBankHolidays()

        If fOK Then mblnRegions = True

      ElseIf fOK And ((mlngRegion > 0) Or (mblnPersonnelBase And (modPersonnelSpecifics.grtRegionType = modPersonnelSpecifics.RegionType.rtStaticRegion))) And (Not mblnGroupByDescription) Then

        'get static bank holidays collection
        'UPGRADE_WARNING: Couldn't resolve default property of object Get_StaticBankHolidays. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        fOK = Get_StaticBankHolidays()

        If fOK Then
          mblnRegions = True
          mblnStaticReg = True
        End If

      Else
        mblnDisableRegions = True

      End If
    End If



    If (fOK And mblnPersonnelBase And (modPersonnelSpecifics.gwptWorkingPatternType = modPersonnelSpecifics.WorkingPatternType.wptHistoricWPattern) And (Not mblnGroupByDescription)) Or (fOK And (mblnPersonnelBase And (modPersonnelSpecifics.gwptWorkingPatternType = modPersonnelSpecifics.WorkingPatternType.wptStaticWPattern) And (Not mblnGroupByDescription))) Then

      blnWorkingPatternEnabled = CheckPermission_WPInfo()
    End If

    If blnWorkingPatternEnabled Then
      If fOK And mblnPersonnelBase And (modPersonnelSpecifics.gwptWorkingPatternType = modPersonnelSpecifics.WorkingPatternType.wptHistoricWPattern) And (Not mblnGroupByDescription) Then

        'get historical working patterns
        'UPGRADE_WARNING: Couldn't resolve default property of object Get_HistoricWorkingPatterns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        fOK = Get_HistoricWorkingPatterns()

        If fOK Then mblnWorkingPatterns = True

      ElseIf fOK And (mblnPersonnelBase And (modPersonnelSpecifics.gwptWorkingPatternType = modPersonnelSpecifics.WorkingPatternType.wptStaticWPattern) And (Not mblnGroupByDescription)) Then

        'get static working patterns
        'UPGRADE_WARNING: Couldn't resolve default property of object Get_StaticWorkingPatterns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        fOK = Get_StaticWorkingPatterns()

        If fOK Then
          mblnWorkingPatterns = True
          mblnStaticWP = True
        End If

      Else
        mblnDisableWPs = True

      End If
    End If

    Initialise_WP_Region = True

  End Function

  Public Function Get_HistoricWorkingPatterns() As Object

    On Error GoTo ErrorTrap

    If mblnDisableWPs Then
      'UPGRADE_WARNING: Couldn't resolve default property of object Get_HistoricWorkingPatterns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      Get_HistoricWorkingPatterns = False
      Exit Function
    End If

    Dim rsCC As ADODB.Recordset 'career change data for base records
    Dim colWorkingPatterns As clsCalendarEvents

    Dim strSQLCC As String 'sql for retieving career change data

    Dim dtStartDate As Date
    Dim dtEndDate As Date

    Dim avCareerRanges(,) As String
    Dim intNextIndex As Short

    Dim blnNewBaseRecord As Boolean
    Dim lngBaseRecordID As Integer

    Dim intCount As Short

    ReDim avCareerRanges(4, 0)

    strSQLCC = vbNullString
    strSQLCC = strSQLCC & "SELECT " & vbNewLine
    strSQLCC = strSQLCC & "     " & gsPersonnelHWorkingPatternTableRealSource & ".ID_" & mlngCalendarReportsBaseTable & "," & vbNewLine
    strSQLCC = strSQLCC & "     " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternDateColumnName & ", " & vbNewLine
    strSQLCC = strSQLCC & "     " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternColumnName & ", " & vbNewLine
    strSQLCC = strSQLCC & "     (SELECT COUNT(B.ID) FROM " & gsPersonnelHWorkingPatternTableRealSource & " B WHERE B.ID_" & mlngCalendarReportsBaseTable & " = " & gsPersonnelHWorkingPatternTableRealSource & ".ID_" & mlngCalendarReportsBaseTable & " AND B." & gsPersonnelHWorkingPatternDateColumnName & " IS NOT NULL) AS 'CareerChanges' " & vbNewLine
    strSQLCC = strSQLCC & "FROM " & gsPersonnelHWorkingPatternTableRealSource & " " & vbNewLine
    If Len(Trim(mstrSQLIDs)) > 0 Then
      strSQLCC = strSQLCC & "WHERE " & vbNewLine
      strSQLCC = strSQLCC & "     " & gsPersonnelHWorkingPatternTableRealSource & ".ID_" & mlngCalendarReportsBaseTable & " IN (" & mstrSQLIDs & ") " & vbNewLine
      strSQLCC = strSQLCC & " AND " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternDateColumnName & " IS NOT NULL " & vbNewLine
    Else
      strSQLCC = strSQLCC & "WHERE " & vbNewLine
      strSQLCC = strSQLCC & "      " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternDateColumnName & " IS NOT NULL " & vbNewLine
    End If
    strSQLCC = strSQLCC & "ORDER BY "
    strSQLCC = strSQLCC & "     " & gsPersonnelHWorkingPatternTableRealSource & ".ID_" & mlngCalendarReportsBaseTable & ", "
    strSQLCC = strSQLCC & "     " & gsPersonnelHWorkingPatternTableRealSource & "." & gsPersonnelHWorkingPatternDateColumnName & " "

    rsCC = datGeneral.GetRecords(strSQLCC)

    lngBaseRecordID = -1
    blnNewBaseRecord = False

    '******************************************************************************
    'Create an array containing the ranges of career change period
    With rsCC

      If Not (.BOF And .EOF) Then

        Do While Not .EOF
          intNextIndex = UBound(avCareerRanges, 2) + 1
          ReDim Preserve avCareerRanges(4, intNextIndex)

          If lngBaseRecordID <> .Fields("ID_" & CStr(mlngCalendarReportsBaseTable)).Value Then
            lngBaseRecordID = .Fields("ID_" & CStr(mlngCalendarReportsBaseTable)).Value
            blnNewBaseRecord = True
            dtStartDate = .Fields(gsPersonnelHWorkingPatternDateColumnName).Value

            avCareerRanges(0, intNextIndex) = CStr(lngBaseRecordID) 'BaseRecordID
            avCareerRanges(1, intNextIndex) = CStr(dtStartDate) 'Start Date
            avCareerRanges(2, intNextIndex) = "" 'End Date
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            avCareerRanges(3, intNextIndex) = IIf(IsDBNull(.Fields(gsPersonnelHWorkingPatternColumnName).Value), "", .Fields(gsPersonnelHWorkingPatternColumnName).Value) 'Working Pattern???
            avCareerRanges(4, intNextIndex) = .Fields("CareerChanges").Value 'Career Change Count

          Else
            dtStartDate = .Fields(gsPersonnelHWorkingPatternDateColumnName).Value
            dtEndDate = dtStartDate
            avCareerRanges(2, intNextIndex - 1) = CStr(dtEndDate) 'End Date

            avCareerRanges(0, intNextIndex) = CStr(lngBaseRecordID) 'BaseRecordID
            avCareerRanges(1, intNextIndex) = CStr(dtStartDate) 'Start Date
            avCareerRanges(2, intNextIndex) = "" 'End Date
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            avCareerRanges(3, intNextIndex) = IIf(IsDBNull(.Fields(gsPersonnelHWorkingPatternColumnName).Value), "", .Fields(gsPersonnelHWorkingPatternColumnName).Value) 'Working Pattern???
            avCareerRanges(4, intNextIndex) = .Fields("CareerChanges").Value 'Career Change Count

          End If

          blnNewBaseRecord = False
          .MoveNext()
        Loop

      Else
        'UPGRADE_WARNING: Couldn't resolve default property of object Get_HistoricWorkingPatterns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Get_HistoricWorkingPatterns = True
        GoTo TidyUpAndExit

      End If

    End With
    '******************************************************************************

    lngBaseRecordID = -1
    blnNewBaseRecord = False

    '##############################################################################
    'populate form WP string with form data

    Dim INPUT_STRING As String
    Dim intRecordWP As Short

    INPUT_STRING = vbNullString
    intRecordWP = 0

    mstrWPFormString = "<FORM id=frmWP name=frmWP style=""visibility:hidden;display:none"">" & vbNewLine

    For intCount = 1 To UBound(avCareerRanges, 2) Step 1

      If lngBaseRecordID <> CInt(avCareerRanges(0, intCount)) Then
        If Not (colWorkingPatterns Is Nothing) Then
          mcolHistoricWorkingPatterns.Add(colWorkingPatterns, CStr(lngBaseRecordID))
          'UPGRADE_NOTE: Object colWorkingPatterns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
          colWorkingPatterns = Nothing
        End If
        colWorkingPatterns = New clsCalendarEvents

        lngBaseRecordID = CInt(avCareerRanges(0, intCount))
        blnNewBaseRecord = True
        intRecordWP = 0
        mstrWPFormString = mstrWPFormString & vbNewLine & vbTab & "<INPUT NAME=txtWPCOUNT_" & lngBaseRecordID & " ID=txtWPCOUNT_" & lngBaseRecordID & " VALUE=""" & avCareerRanges(4, intCount) & """>" & vbNewLine
      End If

      colWorkingPatterns.Add(CStr(colWorkingPatterns.Count), CStr(lngBaseRecordID), , , CShort(avCareerRanges(4, intCount)), , avCareerRanges(1, intCount), , , , avCareerRanges(2, intCount), , , , , , , , , , , , , , , , , , , , , avCareerRanges(3, intCount))

      intRecordWP = intRecordWP + 1

      INPUT_STRING = vbNullString
      INPUT_STRING = INPUT_STRING & VB6.Format(avCareerRanges(1, intCount), mstrClientDateFormat) & "_"
      INPUT_STRING = INPUT_STRING & VB6.Format(avCareerRanges(2, intCount), mstrClientDateFormat) & "_"
      INPUT_STRING = INPUT_STRING & avCareerRanges(3, intCount)

      mstrWPFormString = mstrWPFormString & vbTab & "<INPUT NAME=txtWP_" & lngBaseRecordID & "_" & intRecordWP & " ID=txtWP_" & lngBaseRecordID & "_" & intRecordWP & " VALUE=""" & Replace(INPUT_STRING, """", "&quot;") & """>" & vbNewLine

      If (intCount = UBound(avCareerRanges, 2)) And Not (colWorkingPatterns Is Nothing) Then
        mcolHistoricWorkingPatterns.Add(colWorkingPatterns, CStr(lngBaseRecordID))
        'UPGRADE_NOTE: Object colWorkingPatterns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        colWorkingPatterns = Nothing
      End If

      blnNewBaseRecord = False
    Next intCount

    mstrWPFormString = mstrWPFormString & "</FORM>" & vbNewLine & vbNewLine

    '##############################################################################

    'UPGRADE_WARNING: Couldn't resolve default property of object Get_HistoricWorkingPatterns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    Get_HistoricWorkingPatterns = True

TidyUpAndExit:
    'UPGRADE_NOTE: Object rsCC may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsCC = Nothing
    'UPGRADE_NOTE: Object colWorkingPatterns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    colWorkingPatterns = Nothing
    Exit Function

ErrorTrap:
    'UPGRADE_WARNING: Couldn't resolve default property of object Get_HistoricWorkingPatterns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    Get_HistoricWorkingPatterns = False
    GoTo TidyUpAndExit

  End Function

  Public Function Get_HistoricBankHolidays() As Object

    On Error GoTo ErrorTrap

    If mblnDisableRegions Then
      'UPGRADE_WARNING: Couldn't resolve default property of object Get_HistoricBankHolidays. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      Get_HistoricBankHolidays = False
      Exit Function
    End If

    Dim rsCC As ADODB.Recordset 'career change data for base records
    Dim rsPersonnelBHols As ADODB.Recordset
    Dim colBankHolidays As clsBankHolidays

    Dim strSQLCC As String 'sql for retieving career change data
    Dim strSQLAllBHols As String
    Dim strSQLSelect As String
    Dim strSQLWhere As String
    Dim strSQLDateRegion As String
    Dim strSQLOrder As String

    Dim dtStartDate As Date
    Dim dtEndDate As Date

    Dim intNextIndex As Short

    Dim blnNewBaseRecord As Boolean
    Dim lngBaseRecordID As Integer

    Dim lng100Counter As Integer
    Dim lngBaseRowCount As Integer
    Dim lngMainBaseCounter As Integer
    Dim lngTotalCareerChanges As Integer

    Dim intCount As Short
    Dim intBHolCount As Short
    Dim lngCount As Integer
    Dim fFinalCareerChange As Boolean

    ReDim mavCareerRanges(4, 0)

    Dim INPUT_STRING As String
    Dim intRecordBHol As Short

    intRecordBHol = 0
    mstrBHolFormString = "<FORM id=frmBHol name=frmBHol style=""visibility:hidden;display:none"">" & vbNewLine

    strSQLCC = vbNullString
    strSQLCC = strSQLCC & "SELECT " & vbNewLine
    strSQLCC = strSQLCC & "     " & gsPersonnelHRegionTableRealSource & ".ID_" & mlngCalendarReportsBaseTable & "," & vbNewLine
    strSQLCC = strSQLCC & "     " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & ", " & vbNewLine
    strSQLCC = strSQLCC & "     " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionColumnName & ", " & vbNewLine
    strSQLCC = strSQLCC & "     (SELECT COUNT(B.ID) FROM " & gsPersonnelHRegionTableRealSource & " B WHERE B.ID_" & mlngCalendarReportsBaseTable & " = " & gsPersonnelHRegionTableRealSource & ".ID_" & mlngCalendarReportsBaseTable & " AND B." & gsPersonnelHRegionDateColumnName & " IS NOT NULL) AS 'CareerChanges' " & vbNewLine
    strSQLCC = strSQLCC & "FROM " & gsPersonnelHRegionTableRealSource & " " & vbNewLine

    If Len(Trim(mstrSQLIDs)) > 0 Then
      strSQLCC = strSQLCC & "WHERE " & vbNewLine
      strSQLCC = strSQLCC & "     " & gsPersonnelHRegionTableRealSource & ".ID_" & mlngCalendarReportsBaseTable & " IN (" & mstrSQLIDs & ") " & vbNewLine
      strSQLCC = strSQLCC & " AND " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " IS NOT NULL " & vbNewLine
    Else
      strSQLCC = strSQLCC & "WHERE " & vbNewLine
      strSQLCC = strSQLCC & "      " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " IS NOT NULL " & vbNewLine
    End If

    strSQLCC = strSQLCC & "ORDER BY "
    strSQLCC = strSQLCC & "     " & gsPersonnelHRegionTableRealSource & ".ID_" & mlngCalendarReportsBaseTable & ", "
    strSQLCC = strSQLCC & "     " & gsPersonnelHRegionTableRealSource & "." & gsPersonnelHRegionDateColumnName & " "

    rsCC = datGeneral.GetRecords(strSQLCC)

    lngBaseRecordID = -1
    blnNewBaseRecord = False
    lng100Counter = 0
    lngMainBaseCounter = 0

    '******************************************************************************
    'Create an array containing the ranges of career change period

    With rsCC

      If Not (.BOF And .EOF) Then

        Do While Not .EOF
          intNextIndex = UBound(mavCareerRanges, 2) + 1
          ReDim Preserve mavCareerRanges(4, intNextIndex)

          If lngBaseRecordID <> .Fields("ID_" & CStr(mlngCalendarReportsBaseTable)).Value Then
            lngBaseRecordID = .Fields("ID_" & CStr(mlngCalendarReportsBaseTable)).Value
            blnNewBaseRecord = True
            lngBaseRowCount = lngBaseRowCount + 1
            '          dtStartDate = Format(.Fields(gsPersonnelHRegionDateColumnName).Value, "mm/dd/yyyy")
            dtStartDate = .Fields(gsPersonnelHRegionDateColumnName).Value

            'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(0, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mavCareerRanges(0, intNextIndex) = lngBaseRecordID 'BaseRecordID
            'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(1, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mavCareerRanges(1, intNextIndex) = dtStartDate 'Start Date
            'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(2, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mavCareerRanges(2, intNextIndex) = "" 'End Date
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(3, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mavCareerRanges(3, intNextIndex) = IIf(IsDBNull(.Fields(gsPersonnelHRegionColumnName).Value), "", .Fields(gsPersonnelHRegionColumnName).Value) 'Region
            'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(4, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mavCareerRanges(4, intNextIndex) = .Fields("CareerChanges").Value 'Career Change Count

          Else
            '          dtStartDate = Format(.Fields(gsPersonnelHRegionDateColumnName).Value, "mm/dd/yyyy")
            dtStartDate = .Fields(gsPersonnelHRegionDateColumnName).Value

            dtEndDate = dtStartDate
            'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(2, intNextIndex - 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mavCareerRanges(2, intNextIndex - 1) = dtEndDate 'End Date

            'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(0, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mavCareerRanges(0, intNextIndex) = lngBaseRecordID 'BaseRecordID
            'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(1, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mavCareerRanges(1, intNextIndex) = dtStartDate 'Start Date
            'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(2, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mavCareerRanges(2, intNextIndex) = "" 'End Date
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(3, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mavCareerRanges(3, intNextIndex) = IIf(IsDBNull(.Fields(gsPersonnelHRegionColumnName).Value), "", .Fields(gsPersonnelHRegionColumnName).Value) 'Region
            'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(4, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mavCareerRanges(4, intNextIndex) = .Fields("CareerChanges").Value 'Career Change Count

          End If

          blnNewBaseRecord = False
          .MoveNext()
        Loop

      Else
        'UPGRADE_WARNING: Couldn't resolve default property of object Get_HistoricBankHolidays. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Get_HistoricBankHolidays = True
        GoTo TidyUpAndExit

      End If

    End With

    lngTotalCareerChanges = UBound(mavCareerRanges, 2)

    '******************************************************************************

    lngBaseRecordID = -1
    blnNewBaseRecord = False

    INPUT_STRING = vbNullString
    intRecordBHol = 0

    mstrRegionFormString = "<FORM id=frmRegion name=frmRegion style=""visibility:hidden;display:none"">" & vbNewLine

    For intCount = 1 To UBound(mavCareerRanges, 2) Step 1

      'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      If lngBaseRecordID <> CInt(mavCareerRanges(0, intCount)) Then
        'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(4, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(0, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        mstrRegionFormString = mstrRegionFormString & vbNewLine & vbTab & "<INPUT NAME=txtRegionCOUNT_" & mavCareerRanges(0, intCount) & " ID=txtRegionCOUNT_" & mavCareerRanges(0, intCount) & " VALUE=""" & mavCareerRanges(4, intCount) & """>" & vbNewLine
        'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(0, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        lngBaseRecordID = mavCareerRanges(0, intCount)
        intRecordBHol = 0
      End If

      intRecordBHol = intRecordBHol + 1

      INPUT_STRING = vbNullString
      'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      INPUT_STRING = INPUT_STRING & mavCareerRanges(1, intCount) & "_"
      'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      INPUT_STRING = INPUT_STRING & mavCareerRanges(2, intCount) & "_"
      'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      INPUT_STRING = INPUT_STRING & mavCareerRanges(3, intCount)

      'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(0, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      mstrRegionFormString = mstrRegionFormString & vbTab & "<INPUT NAME=txtRegion_" & mavCareerRanges(0, intCount) & "_" & intRecordBHol & " ID=txtRegion_" & mavCareerRanges(0, intCount) & "_" & intRecordBHol & " VALUE=""" & INPUT_STRING & """>" & vbNewLine

    Next intCount

    mstrRegionFormString = mstrRegionFormString & "</FORM>" & vbNewLine


    lngBaseRecordID = -1
    blnNewBaseRecord = False

    '------------------------------------------------------------------------------
    'Create and execute a 'single' sql string which returns all the bank holidays
    'for all the selcted base table records.

    strSQLAllBHols = vbNullString
    strSQLSelect = vbNullString
    strSQLWhere = vbNullString
    strSQLDateRegion = vbNullString

    For intCount = 1 To UBound(mavCareerRanges, 2) Step 1

      'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(0, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      If lngBaseRecordID <> mavCareerRanges(0, intCount) Then
        'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(0, intCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        lngBaseRecordID = mavCareerRanges(0, intCount)
        blnNewBaseRecord = True
        lng100Counter = lng100Counter + 1
        lngMainBaseCounter = lngMainBaseCounter + 1
        strSQLSelect = vbNullString
        strSQLDateRegion = strSQLDateRegion & "         ( " & vbNewLine

        strSQLWhere = "WHERE " & vbNewLine

        intBHolCount = 0
      End If

      intBHolCount = intBHolCount + 1

      'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      strSQLSelect = vbNewLine & "SELECT  '" & mavCareerRanges(0, intCount) & "' AS 'ID' , " & vbNewLine
      strSQLSelect = strSQLSelect & "       " & mstrSQLSelect_RegInfoRegion & " AS 'Region', " & vbNewLine
      strSQLSelect = strSQLSelect & "       " & mstrSQLSelect_BankHolDate & " , " & vbNewLine
      strSQLSelect = strSQLSelect & "       " & mstrSQLSelect_BankHolDesc & " " & vbNewLine
      strSQLSelect = strSQLSelect & "FROM " & gsBHolRegionTableName & " " & vbNewLine

      For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
        '<REGIONAL CODE>
        'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If mvarTableViews(0, lngCount) = glngBHolRegionTableID Then
          'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          strSQLSelect = strSQLSelect & "           LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbNewLine
          'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          strSQLSelect = strSQLSelect & "           ON  " & gsBHolRegionTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbNewLine
        End If
      Next lngCount

      strSQLSelect = strSQLSelect & "           INNER JOIN " & gsBHolTableRealSource & vbNewLine
      strSQLSelect = strSQLSelect & "           ON  " & gsBHolRegionTableName & ".ID = " & gsBHolTableRealSource & ".ID_" & glngBHolRegionTableID & vbNewLine

      If intBHolCount > 1 Then
        strSQLDateRegion = strSQLDateRegion & " OR " & vbNewLine
      End If

      'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      fFinalCareerChange = (intBHolCount = CShort(mavCareerRanges(4, intCount)))

      If fFinalCareerChange Then
        strSQLDateRegion = strSQLDateRegion & "( " & vbNewLine
        'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        strSQLDateRegion = strSQLDateRegion & "(" & gsBHolTableRealSource & "." & gsBHolDateColumnName & " >= CONVERT(datetime, '" & VB6.Format(mavCareerRanges(1, intCount), "mm/dd/yyyy") & "')) " & vbNewLine
        strSQLDateRegion = strSQLDateRegion & " AND (" & gsBHolTableRealSource & "." & gsBHolDateColumnName & " >= '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "') " & vbNewLine
        strSQLDateRegion = strSQLDateRegion & " AND (" & gsBHolTableRealSource & "." & gsBHolDateColumnName & " <= '" & VB6.Format(mdtEndDate, "mm/dd/yyyy") & "') " & vbNewLine
        strSQLDateRegion = strSQLDateRegion & " AND " & vbNewLine
        'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        strSQLDateRegion = strSQLDateRegion & "(" & mstrSQLSelect_RegInfoRegion & " = '" & mavCareerRanges(3, intCount) & "') " & vbNewLine
        strSQLDateRegion = strSQLDateRegion & ") " & vbNewLine
        strSQLDateRegion = strSQLDateRegion & ") " & vbNewLine
      Else
        strSQLDateRegion = strSQLDateRegion & "( " & vbNewLine
        'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        strSQLDateRegion = strSQLDateRegion & "(" & gsBHolTableRealSource & "." & gsBHolDateColumnName & " >= CONVERT(datetime, '" & VB6.Format(mavCareerRanges(1, intCount), "mm/dd/yyyy") & "') " & vbNewLine & " AND (" & gsBHolTableRealSource & "." & gsBHolDateColumnName & " < CONVERT(datetime, '" & VB6.Format(mavCareerRanges(1, intCount + 1), "mm/dd/yyyy") & "'))) " & vbNewLine
        strSQLDateRegion = strSQLDateRegion & " AND (" & gsBHolTableRealSource & "." & gsBHolDateColumnName & " >= '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "') " & vbNewLine
        strSQLDateRegion = strSQLDateRegion & " AND (" & gsBHolTableRealSource & "." & gsBHolDateColumnName & " <= '" & VB6.Format(mdtEndDate, "mm/dd/yyyy") & "') " & vbNewLine
        'UPGRADE_WARNING: Couldn't resolve default property of object mavCareerRanges(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        strSQLDateRegion = strSQLDateRegion & " AND (" & mstrSQLSelect_RegInfoRegion & " = '" & mavCareerRanges(3, intCount) & "') " & vbNewLine
        strSQLDateRegion = strSQLDateRegion & ") " & vbNewLine
      End If

      If fFinalCareerChange Then
        strSQLAllBHols = strSQLAllBHols & strSQLSelect & vbNewLine
        strSQLAllBHols = strSQLAllBHols & strSQLWhere & vbNewLine
        strSQLAllBHols = strSQLAllBHols & strSQLDateRegion & vbNewLine
        strSQLAllBHols = strSQLAllBHols & " UNION ALL "
        strSQLWhere = vbNullString
        strSQLDateRegion = vbNullString
      End If

      'Send the query to SQL Server in batches of approximately 100, to avoid 256(260) Table/Views limit.
      'Do not split base records in to more than one batch!
      If ((lng100Counter = lngBaseRowCount) And fFinalCareerChange) Or ((lng100Counter > 100) And fFinalCareerChange) Or ((lngMainBaseCounter = lngBaseRowCount) And fFinalCareerChange) Then

        strSQLAllBHols = Left(strSQLAllBHols, Len(strSQLAllBHols) - 11)
        strSQLOrder = " ORDER BY 'ID', 'Region' " & vbNewLine
        strSQLAllBHols = strSQLAllBHols & strSQLOrder

        'Open App.Path & "\calrep.txt" For Output As #1
        'Print #1, strSQLAllBHols
        'Close #1

        rsPersonnelBHols = datGeneral.GetRecords(strSQLAllBHols)

        lngBaseRecordID = -1
        blnNewBaseRecord = False
        intRecordBHol = 0

        '##############################################################################
        'populate collections with new data
        With rsPersonnelBHols

          INPUT_STRING = vbNullString

          If Not (.BOF And .EOF) Then

            Do While Not .EOF

              If lngBaseRecordID <> CInt(.Fields("ID").Value) Then

                If Not (colBankHolidays Is Nothing) Then
                  mcolHistoricBankHolidays.Add(colBankHolidays, CStr(lngBaseRecordID))
                  'UPGRADE_NOTE: Object colBankHolidays may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                  colBankHolidays = Nothing
                End If
                colBankHolidays = New clsBankHolidays

                lngBaseRecordID = CInt(.Fields("ID").Value)

              End If

              'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
              colBankHolidays.Add(IIf(IsDBNull(.Fields("Region").Value), "", .Fields("Region").Value), IIf(IsDBNull(.Fields(gsBHolDescriptionColumnName).Value), "", .Fields(gsBHolDescriptionColumnName).Value), IIf(IsDBNull(.Fields(gsBHolDateColumnName).Value), "", .Fields(gsBHolDateColumnName).Value))

              intRecordBHol = intRecordBHol + 1

              INPUT_STRING = vbNullString
              'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
              INPUT_STRING = INPUT_STRING & IIf(IsDBNull(.Fields(gsBHolDateColumnName).Value), "", VB6.Format(.Fields(gsBHolDateColumnName).Value, mstrClientDateFormat)) & "_"
              'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
              INPUT_STRING = INPUT_STRING & IIf(IsDBNull(.Fields("Region").Value), "", .Fields("Region").Value)

              mstrBHolFormString = mstrBHolFormString & vbTab & "<INPUT NAME=txtBHol_" & lngBaseRecordID & "_" & intRecordBHol & " ID=txtBHol_" & lngBaseRecordID & "_" & intRecordBHol & " VALUE=""" & Replace(INPUT_STRING, """", "&quot;") & """>" & vbNewLine

              .MoveNext()

              If Not .EOF Then
                If lngBaseRecordID <> CInt(.Fields("ID").Value) Then
                  mstrBHolFormString = mstrBHolFormString & vbTab & "<INPUT NAME=txtBHolCOUNT_" & lngBaseRecordID & " ID=txtBHolCOUNT_" & lngBaseRecordID & " VALUE=""" & intRecordBHol & """>" & vbNewLine & vbNewLine
                  intRecordBHol = 0
                End If
              Else
                mstrBHolFormString = mstrBHolFormString & vbTab & "<INPUT NAME=txtBHolCOUNT_" & lngBaseRecordID & " ID=txtBHolCOUNT_" & lngBaseRecordID & " VALUE=""" & intRecordBHol & """>" & vbNewLine & vbNewLine
                intRecordBHol = 0
              End If

              If .EOF And Not (colBankHolidays Is Nothing) Then
                mcolHistoricBankHolidays.Add(colBankHolidays, CStr(lngBaseRecordID))
                'UPGRADE_NOTE: Object colBankHolidays may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                colBankHolidays = Nothing
              End If

            Loop


          Else
            'UPGRADE_WARNING: Couldn't resolve default property of object Get_HistoricBankHolidays. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Get_HistoricBankHolidays = True

          End If

        End With
        '##############################################################################

        'Reset SQL string variables ready for next batch to be created.
        strSQLAllBHols = vbNullString
        strSQLSelect = vbNullString
        strSQLWhere = vbNullString
        strSQLDateRegion = vbNullString
        lng100Counter = 0
        lngBaseRecordID = -1
      End If

      blnNewBaseRecord = False

    Next intCount

    mstrBHolFormString = mstrBHolFormString & "</FORM>" & vbNewLine & vbNewLine

    'UPGRADE_WARNING: Couldn't resolve default property of object Get_HistoricBankHolidays. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    Get_HistoricBankHolidays = True

TidyUpAndExit:
    'UPGRADE_NOTE: Object rsCC may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsCC = Nothing
    'UPGRADE_NOTE: Object rsPersonnelBHols may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsPersonnelBHols = Nothing
    'UPGRADE_NOTE: Object colBankHolidays may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    colBankHolidays = Nothing
    Exit Function

ErrorTrap:
    'UPGRADE_WARNING: Couldn't resolve default property of object Get_HistoricBankHolidays. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    Get_HistoricBankHolidays = False
    GoTo TidyUpAndExit

  End Function

  Public Function Get_StaticBankHolidays() As Object

    On Error GoTo ErrorTrap

    If mblnDisableRegions Then
      'UPGRADE_WARNING: Couldn't resolve default property of object Get_StaticBankHolidays. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      Get_StaticBankHolidays = False
      Exit Function
    End If

    Dim rsPersonnelBHols As ADODB.Recordset
    Dim colBankHolidays As clsBankHolidays

    Dim strSQLAllBHols As String

    Dim blnNewBaseRecord As Boolean
    Dim lngBaseRecordID As Integer

    Dim intCount As Short
    Dim intBHolCount As Short
    Dim lngCount As Integer
    Dim lngView As Integer

    Dim INPUT_STRING As String

    strSQLAllBHols = vbNullString
    strSQLAllBHols = strSQLAllBHols & "SELECT DISTINCT  [Base].ID, " & vbNewLine
    strSQLAllBHols = strSQLAllBHols & "                 [RegionInfo].Region, " & vbNewLine
    strSQLAllBHols = strSQLAllBHols & "                 [RegionInfo].Holiday_Date, " & vbNewLine
    strSQLAllBHols = strSQLAllBHols & "                 [RegionInfo].Description " & vbNewLine

    'gsBHolTableRealSource
    'gsBHolRegionTableName
    strSQLAllBHols = strSQLAllBHols & "FROM (SELECT DISTINCT " & vbNewLine
    strSQLAllBHols = strSQLAllBHols & "             " & mstrCalendarReportsBaseTableName & ".ID AS 'ID', " & vbNewLine
    strSQLAllBHols = strSQLAllBHols & "             " & mstrSQLSelect_PersonnelStaticRegion & " AS 'Region' " & vbNewLine

    If mlngRegion > 0 Then
      strSQLAllBHols = strSQLAllBHols & "      FROM " & mstrCalendarReportsBaseTableName & vbNewLine
    Else
      strSQLAllBHols = strSQLAllBHols & "      FROM " & gsPersonnelTableName & vbNewLine
    End If

    For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
      '<PERSONNEL CODE>
      If mlngRegion > 0 Then
        'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If mvarTableViews(0, lngCount) = mlngCalendarReportsBaseTable Then
          'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          strSQLAllBHols = strSQLAllBHols & "           LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbNewLine
          'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          strSQLAllBHols = strSQLAllBHols & "           ON  " & mstrCalendarReportsBaseTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbNewLine
        End If
      Else
        'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If mvarTableViews(0, lngCount) = mlngCalendarReportsBaseTable Then
          'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          strSQLAllBHols = strSQLAllBHols & "           LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbNewLine
          'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          strSQLAllBHols = strSQLAllBHols & "           ON  " & gsPersonnelTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbNewLine
        End If
      End If
    Next lngCount

    If Len(Trim(mstrSQLIDs)) > 0 Then
      strSQLAllBHols = strSQLAllBHols & "      WHERE " & mstrCalendarReportsBaseTableName & ".ID IN (" & mstrSQLIDs & ") " & vbNewLine
    End If

    strSQLAllBHols = strSQLAllBHols & "      ) AS [Base] " & vbNewLine

    strSQLAllBHols = strSQLAllBHols & "   INNER JOIN " & vbNewLine

    strSQLAllBHols = strSQLAllBHols & "   (SELECT DISTINCT " & vbNewLine
    strSQLAllBHols = strSQLAllBHols & "   " & gsBHolRegionTableName & ".ID AS 'ID', " & vbNewLine
    strSQLAllBHols = strSQLAllBHols & "   " & mstrSQLSelect_RegInfoRegion & " AS 'Region', " & vbNewLine
    strSQLAllBHols = strSQLAllBHols & "   " & mstrSQLSelect_BankHolDate & ", " & vbNewLine
    strSQLAllBHols = strSQLAllBHols & "   " & mstrSQLSelect_BankHolDesc & " " & vbNewLine

    strSQLAllBHols = strSQLAllBHols & "      FROM " & gsBHolRegionTableName & vbNewLine

    For lngCount = 0 To UBound(mvarTableViews, 2) Step 1
      '<REGIONAL CODE>
      'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, lngCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      If mvarTableViews(0, lngCount) = glngBHolRegionTableID Then
        'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        strSQLAllBHols = strSQLAllBHols & "           LEFT OUTER JOIN " & mvarTableViews(3, lngCount) & vbNewLine
        'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        strSQLAllBHols = strSQLAllBHols & "           ON  " & gsBHolRegionTableName & ".ID = " & mvarTableViews(3, lngCount) & ".ID" & vbNewLine
      End If
    Next lngCount

    strSQLAllBHols = strSQLAllBHols & "           INNER JOIN " & gsBHolTableRealSource & vbNewLine
    strSQLAllBHols = strSQLAllBHols & "           ON  " & gsBHolRegionTableName & ".ID = " & gsBHolTableRealSource & ".ID_" & glngBHolRegionTableID & vbNewLine

    If Len(Trim(mstrSQLIDs)) > 0 Then
      strSQLAllBHols = strSQLAllBHols & "     WHERE (" & gsBHolTableRealSource & "." & gsBHolDateColumnName & " >= '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "') " & vbNewLine
      strSQLAllBHols = strSQLAllBHols & "         AND (" & gsBHolTableRealSource & "." & gsBHolDateColumnName & " <= '" & VB6.Format(mdtEndDate, "mm/dd/yyyy") & "') " & vbNewLine
    Else
      strSQLAllBHols = strSQLAllBHols & "     WHERE (" & gsBHolTableRealSource & "." & gsBHolDateColumnName & " >= '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "') " & vbNewLine
      strSQLAllBHols = strSQLAllBHols & "         AND (" & gsBHolTableRealSource & "." & gsBHolDateColumnName & " <= '" & VB6.Format(mdtEndDate, "mm/dd/yyyy") & "') " & vbNewLine
    End If

    strSQLAllBHols = strSQLAllBHols & "    ) AS [RegionInfo] " & vbNewLine
    strSQLAllBHols = strSQLAllBHols & "    ON [Base].Region = [RegionInfo].Region " & vbNewLine
    strSQLAllBHols = strSQLAllBHols & "ORDER BY [Base].ID " & vbNewLine

    rsPersonnelBHols = datGeneral.GetRecords(strSQLAllBHols)

    lngBaseRecordID = -1
    blnNewBaseRecord = False

    '##############################################################################
    'populate collections with new data
    With rsPersonnelBHols

      INPUT_STRING = vbNullString
      intBHolCount = 0

      mstrBHolFormString = "<FORM id=frmBHol name=frmBHol style=""visibility:hidden;display:none"">" & vbNewLine

      If Not (.BOF And .EOF) Then

        Do While Not .EOF

          If lngBaseRecordID <> .Fields("ID").Value Then
            intBHolCount = 0
            If Not (colBankHolidays Is Nothing) Then
              mcolStaticBankHolidays.Add(colBankHolidays, CStr(lngBaseRecordID))
              'UPGRADE_NOTE: Object colBankHolidays may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
              colBankHolidays = Nothing
            End If
            colBankHolidays = New clsBankHolidays

            lngBaseRecordID = .Fields("ID").Value
            blnNewBaseRecord = True

          End If

          intBHolCount = intBHolCount + 1

          'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
          colBankHolidays.Add(IIf(IsDBNull(.Fields("Region").Value), "", .Fields("Region").Value), IIf(IsDBNull(.Fields(gsBHolDescriptionColumnName).Value), "", .Fields(gsBHolDescriptionColumnName).Value), IIf(IsDBNull(.Fields(gsBHolDateColumnName).Value), "", .Fields(gsBHolDateColumnName).Value))

          INPUT_STRING = vbNullString
          'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
          INPUT_STRING = INPUT_STRING & IIf(IsDBNull(.Fields(gsBHolDateColumnName).Value), "", VB6.Format(.Fields(gsBHolDateColumnName).Value, mstrClientDateFormat)) & "_"
          'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
          INPUT_STRING = INPUT_STRING & IIf(IsDBNull(.Fields("Region").Value), "", .Fields("Region").Value)
          '        INPUT_STRING = INPUT_STRING & IIf(IsNull(.Fields(gsBHolDescriptionColumnName).Value), "", .Fields(gsBHolDescriptionColumnName).Value) & "_"

          mstrBHolFormString = mstrBHolFormString & vbTab & "<INPUT NAME=txtBHol_" & .Fields("ID").Value & "_" & intBHolCount & " ID=txtBHol_" & .Fields("ID").Value & "_" & intBHolCount & " VALUE=""" & Replace(INPUT_STRING, """", "&quot;") & """>" & vbNewLine

          blnNewBaseRecord = False

          .MoveNext()

          If Not .EOF Then
            If lngBaseRecordID <> CInt(.Fields("ID").Value) Then
              mstrBHolFormString = mstrBHolFormString & vbTab & "<INPUT NAME=txtBHolCOUNT_" & lngBaseRecordID & " ID=txtBHolCOUNT_" & lngBaseRecordID & " VALUE=""" & intBHolCount & """>" & vbNewLine
              intBHolCount = 0
            End If
          Else
            mstrBHolFormString = mstrBHolFormString & vbTab & "<INPUT NAME=txtBHolCOUNT_" & lngBaseRecordID & " ID=txtBHolCOUNT_" & lngBaseRecordID & " VALUE=""" & intBHolCount & """>" & vbNewLine
            intBHolCount = 0
          End If


          If .EOF And Not (colBankHolidays Is Nothing) Then
            mcolStaticBankHolidays.Add(colBankHolidays, CStr(lngBaseRecordID))
            'UPGRADE_NOTE: Object colBankHolidays may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colBankHolidays = Nothing
          End If
        Loop

      Else
        mstrBHolFormString = mstrBHolFormString & vbTab & "<INPUT NAME=txtBHolCOUNT_" & lngBaseRecordID & " ID=txtBHolCOUNT_" & lngBaseRecordID & " VALUE=""" & intBHolCount & """>" & vbNewLine

      End If

      mstrBHolFormString = mstrBHolFormString & "</FORM>" & vbNewLine & vbNewLine

    End With
    '##############################################################################

    'UPGRADE_WARNING: Couldn't resolve default property of object Get_StaticBankHolidays. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    Get_StaticBankHolidays = True

TidyUpAndExit:
    'UPGRADE_NOTE: Object rsPersonnelBHols may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsPersonnelBHols = Nothing
    'UPGRADE_NOTE: Object colBankHolidays may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    colBankHolidays = Nothing
    Exit Function

ErrorTrap:
    'UPGRADE_WARNING: Couldn't resolve default property of object Get_StaticBankHolidays. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    Get_StaticBankHolidays = False
    GoTo TidyUpAndExit

  End Function

  Public Function HTML_MonthCombo(ByVal piStartMonth As Object) As Object

    'Build month selection dropdown combo
    Dim iCount As Short
    Dim strHTML As String

    'UPGRADE_WARNING: Couldn't resolve default property of object piStartMonth. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    piStartMonth = IIf(IsNumeric(piStartMonth), piStartMonth, 1)

    strHTML = "<select name='cboMonth' id='cboMonth' class='combo' style='WIDTH: 100px' onChange='monthChange();'>" & vbNewLine

    For iCount = 1 To 12

      If iCount = 1 Then
        strHTML = strHTML & "<OPTION selected value=" & Trim(Str(iCount)) & ">" & StrConv(MonthName(iCount), VbStrConv.ProperCase) & vbNewLine
      Else
        strHTML = strHTML & "<OPTION value=" & Trim(Str(iCount)) & ">" & StrConv(MonthName(iCount), VbStrConv.ProperCase) & vbNewLine
      End If

    Next iCount

    'UPGRADE_WARNING: Couldn't resolve default property of object HTML_MonthCombo. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    HTML_MonthCombo = strHTML & "  </SELECT>"

  End Function

  Public Function Get_StaticWorkingPatterns() As Object

    On Error GoTo ErrorTrap

    If mblnDisableWPs Then
      'UPGRADE_WARNING: Couldn't resolve default property of object Get_StaticWorkingPatterns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      Get_StaticWorkingPatterns = False
      Exit Function
    End If

    Dim colWorkingPatterns As clsCalendarEvents

    Dim strSQLAllBHols As String

    Dim blnNewBaseRecord As Boolean
    Dim lngBaseRecordID As Integer

    Dim intCount As Short
    Dim intBHolCount As Short

    Dim INPUT_STRING As String

    INPUT_STRING = vbNullString

    mstrWPFormString = "<FORM id=frmWP name=frmWP style=""visibility:hidden;display:none"">" & vbNewLine

    lngBaseRecordID = -1
    blnNewBaseRecord = False

    '##############################################################################
    'populate collections with new data
    With mrsCalendarBaseInfo

      If Not (.BOF And .EOF) Then
        .MoveFirst()

        Do While Not .EOF

          If lngBaseRecordID <> .Fields(mstrBaseIDColumn).Value Then

            If Not (colWorkingPatterns Is Nothing) Then
              mstrWPFormString = mstrWPFormString & vbNewLine & vbTab & "<INPUT NAME=txtWPCOUNT_" & lngBaseRecordID & " ID=txtWPCOUNT_" & lngBaseRecordID & " VALUE=1>" & vbNewLine

              mcolStaticWorkingPatterns.Add(colWorkingPatterns, CStr(lngBaseRecordID))
              'UPGRADE_NOTE: Object colWorkingPatterns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
              colWorkingPatterns = Nothing
            End If
            colWorkingPatterns = New clsCalendarEvents

            lngBaseRecordID = .Fields(mstrBaseIDColumn).Value
            blnNewBaseRecord = True

          End If

          'lngBaseRecordID = .Fields(mstrBaseIDColumn).Value

          INPUT_STRING = vbNullString
          INPUT_STRING = INPUT_STRING & .Fields(gsPersonnelWorkingPatternColumnName).Value

          mstrWPFormString = mstrWPFormString & vbTab & "<INPUT NAME=txtWP_" & lngBaseRecordID & " ID=txtBHol_" & lngBaseRecordID & " VALUE=""" & Replace(INPUT_STRING, """", "&quot;") & """>" & vbNewLine

          'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
          colWorkingPatterns.Add(CStr(colWorkingPatterns.Count), CStr(lngBaseRecordID), , , , , , , , , , , , , , , , , , , , , , , , , , , , , , IIf(IsDBNull(.Fields(gsPersonnelWorkingPatternColumnName).Value), "              ", .Fields(gsPersonnelWorkingPatternColumnName).Value))

          blnNewBaseRecord = False

          .MoveNext()

          If .EOF And Not (colWorkingPatterns Is Nothing) Then
            mstrWPFormString = mstrWPFormString & vbNewLine & vbTab & "<INPUT NAME=txtWPCOUNT_" & lngBaseRecordID & " ID=txtWPCOUNT_" & lngBaseRecordID & " VALUE=1>" & vbNewLine
            mcolStaticWorkingPatterns.Add(colWorkingPatterns, CStr(lngBaseRecordID))
            'UPGRADE_NOTE: Object colWorkingPatterns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            colWorkingPatterns = Nothing
          End If

        Loop

      Else
        'UPGRADE_WARNING: Couldn't resolve default property of object Get_StaticWorkingPatterns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Get_StaticWorkingPatterns = True
        GoTo TidyUpAndExit

      End If

      mstrWPFormString = mstrWPFormString & "</FORM>" & vbNewLine

    End With
    '##############################################################################

    'UPGRADE_WARNING: Couldn't resolve default property of object Get_StaticWorkingPatterns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    Get_StaticWorkingPatterns = True

TidyUpAndExit:
    'UPGRADE_NOTE: Object colWorkingPatterns may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    colWorkingPatterns = Nothing
    Exit Function

ErrorTrap:
    'UPGRADE_WARNING: Couldn't resolve default property of object Get_StaticWorkingPatterns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    Get_StaticWorkingPatterns = False
    GoTo TidyUpAndExit

  End Function

  Private Function CheckPermission_RegionInfo() As Boolean

    Dim strTableColumn As String


    'Check the  Bank Holiday Region Table - Region Table
    '           Bank Holiday Region Table - Region Column
    '           Bank Holidays Table - Bank Holiday Table
    '           Bank Holidays Table - Date Column
    '           Bank Holidays Table - Descripiton Column
    '...Bank Holiday module setup information.
    'If any are blank then we need to allow the report to run, but disable the Bank Holiday Display Options.
    If gsBHolRegionTableName = "" Or gsBHolRegionColumnName = "" Or gsBHolTableName = "" Or gsBHolDateColumnName = "" Or gsBHolDescriptionColumnName = "" Then

      GoTo DisableRegions
    End If

    'Check the  Career Change Region - Static Region Column
    '           Career Change Region - Historic Region Table
    '           Career Change Region - Historic Region Column
    '           Career Change Region - Historic Region Effective Date Column
    '...Personnel - Career Change module setup information.
    'If any are blank then we need to allow the report to run, but disable the Bank Holiday Display Options.
    If gsPersonnelRegionColumnName = "" Then
      If gsPersonnelHRegionTableName = "" Or gsPersonnelHRegionColumnName = "" Or gsPersonnelHRegionDateColumnName = "" Then

        GoTo DisableRegions
      End If
    End If




    '*******************************************************************
    ' All Region module information is setup                           *
    ' Now check the permissions on the Region module setup information *
    '*******************************************************************
    'Bank Holiday Region Table - Region Table (Regional Information)
    'Bank Holiday Region Table - Region Column
    '///////////////////////////////////////////////
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If CheckPermission_Columns(glngBHolRegionTableID, gsBHolRegionTableName, gsBHolRegionColumnName, strTableColumn) Then
      mstrSQLSelect_RegInfoRegion = strTableColumn
      strTableColumn = vbNullString
    Else
      GoTo DisableRegions
    End If
    '///////////////////////////////////////////////
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


    'Bank Holidays Table - Bank Holiday Table (Region History)
    'Bank Holidays Table - Date Column
    '///////////////////////////////////////////////
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If CheckPermission_Columns(glngBHolTableID, gsBHolTableName, gsBHolDateColumnName, strTableColumn) Then
      mstrSQLSelect_BankHolDate = strTableColumn
      strTableColumn = vbNullString
    Else
      GoTo DisableRegions
    End If
    '///////////////////////////////////////////////
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


    'Bank Holidays Table - Bank Holiday Table (Region History)
    'Bank Holidays Table - Descripiton Column
    '///////////////////////////////////////////////
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If CheckPermission_Columns(glngBHolTableID, gsBHolTableName, gsBHolDescriptionColumnName, strTableColumn) Then
      mstrSQLSelect_BankHolDesc = strTableColumn
      strTableColumn = vbNullString
    Else
      GoTo DisableRegions
    End If
    '///////////////////////////////////////////////
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\




    '*******************************************************************
    ' Permission granted on all Region module information.             *
    ' Now check the permissions on the                                 *
    ' Personnel Career Change Region module setup information          *
    '*******************************************************************
    If mlngRegion > 0 Then
      '///////////////////////////////////////////////
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      If CheckPermission_Columns(mlngCalendarReportsBaseTable, mstrCalendarReportsBaseTableName, mstrRegion, strTableColumn) Then
        mstrSQLSelect_PersonnelStaticRegion = strTableColumn
        strTableColumn = vbNullString
      Else
        GoTo DisableRegions
      End If
      '///////////////////////////////////////////////
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

    Else
      'Check Career Change Region access
      If gsPersonnelRegionColumnName <> "" Then
        'Personnel Table
        'Career Change Region - Static Region Column
        '///////////////////////////////////////////////
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        If CheckPermission_Columns(glngPersonnelTableID, gsPersonnelTableName, gsPersonnelRegionColumnName, strTableColumn) Then
          mstrSQLSelect_PersonnelStaticRegion = strTableColumn
          strTableColumn = vbNullString
        Else
          GoTo DisableRegions
        End If
        '///////////////////////////////////////////////
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

      Else
        'Career Change Region - Historic Region Table
        'Career Change Region - Historic Region Column
        '///////////////////////////////////////////////
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        If CheckPermission_Columns(glngPersonnelHRegionTableID, gsPersonnelHRegionTableName, gsPersonnelHRegionColumnName, strTableColumn) Then
          mstrSQLSelect_PersonnelHRegion = strTableColumn
          strTableColumn = vbNullString
        Else
          GoTo DisableRegions
        End If
        '///////////////////////////////////////////////
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

        'Career Change Region - Historic Region Table
        'Career Change Region - Historic Region Effective Date Column
        '///////////////////////////////////////////////
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        If CheckPermission_Columns(glngPersonnelHRegionTableID, gsPersonnelHRegionTableName, gsPersonnelHRegionDateColumnName, strTableColumn) Then
          mstrSQLSelect_PersonnelHDate = strTableColumn
          strTableColumn = vbNullString
        Else
          GoTo DisableRegions
        End If
        '///////////////////////////////////////////////
        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

      End If
    End If

    CheckPermission_RegionInfo = True

TidyUpAndExit:
    Exit Function

DisableRegions:
    mblnDisableRegions = True
    '  ShowBankHolidays = False
    '  IncludeBankHolidays = False
    '  mblnShowBankHols = False
    mblnRegions = False
    CheckPermission_RegionInfo = False
    GoTo TidyUpAndExit

  End Function

  Private Function CheckPermission_Columns(ByRef plngTableID As Integer, ByRef pstrTableName As String, ByRef pstrColumnName As String, ByRef strSQLRef As String) As Boolean

    'This function checks if the current user has read(select) permissions
    'on this column. If the user only has access through views then the
    'relevent views are added to the mvarTableViews() array which in turn
    'are used to create the join part of the query.

    Dim lngTempTableID As Integer
    Dim strTempTableName As String
    Dim strTempColumnName As String
    Dim blnColumnOK As Boolean
    Dim blnFound As Boolean
    Dim blnNoSelect As Boolean
    Dim iLoop1 As Short
    Dim intLoop As Short
    Dim strColumnCode As String
    Dim strSource As String
    Dim intNextIndex As Short
    Dim blnOK As Boolean
    Dim strTable As String
    Dim strColumn As String

    Dim pintNextIndex As Short

    ' Set flags with their starting values
    blnOK = True
    blnNoSelect = False

    strTable = vbNullString
    strColumn = vbNullString

    ' Load the temp variables
    lngTempTableID = plngTableID
    strTempTableName = pstrTableName
    strTempColumnName = pstrColumnName

    ' Check permission on that column
    mobjColumnPrivileges = GetColumnPrivileges(strTempTableName)
    mstrRealSource = gcoTablePrivileges.Item(strTempTableName).RealSource

    blnColumnOK = mobjColumnPrivileges.IsValid(strTempColumnName)

    If blnColumnOK Then
      blnColumnOK = mobjColumnPrivileges.Item(strTempColumnName).AllowSelect
    End If

    If blnColumnOK Then
      ' this column can be read direct from the tbl/view or from a parent table
      strTable = mstrRealSource
      strColumn = strTempColumnName

      '    ' If the table isnt the base table (or its realsource) then
      '    ' Check if it has already been added to the array. If not, add it.
      '    If lngTempTableID <> mlngCalendarReportsBaseTable Then
      blnFound = False
      For intNextIndex = 1 To UBound(mvarTableViews, 2)
        'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(2, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(1, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If mvarTableViews(1, intNextIndex) = 0 And mvarTableViews(2, intNextIndex) = lngTempTableID Then
          blnFound = True
          Exit For
        End If
      Next intNextIndex

      If Not blnFound Then
        intNextIndex = UBound(mvarTableViews, 2) + 1
        ReDim Preserve mvarTableViews(3, intNextIndex)
        'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(1, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        mvarTableViews(1, intNextIndex) = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(2, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        mvarTableViews(2, intNextIndex) = lngTempTableID
      End If
      '    End If

      strSQLRef = strTable & "." & strColumn
    Else

      ' this column cannot be read direct. If its from a parent, try parent views
      ' Loop thru the views on the table, seeing if any have read permis for the column

      ReDim mstrViews(0)
      For Each mobjTableView In gcoTablePrivileges.Collection
        If (Not mobjTableView.IsTable) And (mobjTableView.TableID = lngTempTableID) And (mobjTableView.AllowSelect) Then

          strSource = mobjTableView.ViewName
          mstrRealSource = gcoTablePrivileges.Item(strSource).RealSource

          ' Get the column permission for the view
          mobjColumnPrivileges = GetColumnPrivileges(strSource)

          ' If we can see the column from this view
          If mobjColumnPrivileges.IsValid(strTempColumnName) Then
            If mobjColumnPrivileges.Item(strTempColumnName).AllowSelect Then

              ReDim Preserve mstrViews(UBound(mstrViews) + 1)
              mstrViews(UBound(mstrViews)) = mobjTableView.ViewName

              ' Check if view has already been added to the array
              blnFound = False
              For intNextIndex = 0 To UBound(mvarTableViews, 2)
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(2, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(1, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If mvarTableViews(1, intNextIndex) = 1 And mvarTableViews(2, intNextIndex) = mobjTableView.ViewID Then
                  blnFound = True
                  Exit For
                End If
              Next intNextIndex

              If Not blnFound Then
                ' View hasnt yet been added, so add it !
                intNextIndex = UBound(mvarTableViews, 2) + 1
                ReDim Preserve mvarTableViews(3, intNextIndex)
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(0, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarTableViews(0, intNextIndex) = mobjTableView.TableID
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(1, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarTableViews(1, intNextIndex) = 1
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(2, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarTableViews(2, intNextIndex) = mobjTableView.ViewID
                'UPGRADE_WARNING: Couldn't resolve default property of object mvarTableViews(3, intNextIndex). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mvarTableViews(3, intNextIndex) = mobjTableView.ViewName
              End If

            End If
          End If
        End If

      Next mobjTableView
      'UPGRADE_NOTE: Object mobjTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
      mobjTableView = Nothing

      ' Does the user have select permission thru ANY views ?
      If UBound(mstrViews) = 0 Then
        blnNoSelect = True
      Else
        strSQLRef = ""
        For pintNextIndex = 1 To UBound(mstrViews)
          If pintNextIndex = 1 Then
            strSQLRef = "CASE"
          End If

          strSQLRef = strSQLRef & " WHEN NOT " & mstrViews(pintNextIndex) & "." & strTempColumnName & " IS NULL THEN " & mstrViews(pintNextIndex) & "." & strTempColumnName
        Next pintNextIndex

        If Len(strSQLRef) > 0 Then
          strSQLRef = strSQLRef & " ELSE NULL" & " END "
        End If

      End If

      ' If we cant see a column, then get outta here
      If blnNoSelect Then
        strSQLRef = vbNullString
        CheckPermission_Columns = False
        Exit Function
      End If

      If Not blnOK Then
        strSQLRef = vbNullString
        CheckPermission_Columns = False
        Exit Function
      End If

    End If

    '  'TM01042004 Fault 8428
    '  If mblnCheckingRegionColumn = True Then
    '    mstrRegionColumnRealSource = mstrRealSource
    '  End If

    CheckPermission_Columns = True

  End Function

  Private Function CheckPermission_WPInfo() As Boolean

    Dim objTable As CTablePrivilege
    Dim objColumn As CColumnPrivileges
    Dim pblnColumnOK As Boolean
    Dim strTableColumn As String

    'Check the  Career Change Working Pattern - Static Working Pattern Column
    '           Career Change Working Pattern - Historic Working Pattern Table
    '           Career Change Working Pattern - Historic Working Pattern Column
    '           Career Change Working Pattern - Historic Working Pattern Effective Date Column
    '...Personnel - Career Change module setup information.
    'If any are blank then we need to allow the report to run, but disable the Working Dys Display Option.
    If gsPersonnelWorkingPatternColumnName = "" Then
      If gsPersonnelHWorkingPatternTableName = "" Or gsPersonnelHWorkingPatternColumnName = "" Or gsPersonnelHWorkingPatternDateColumnName = "" Then

        GoTo DisableWPs
      End If
    End If

    '****************************************************************************
    ' All Working Pattern module information is setup                           *
    ' Now check the permissions on the Working Pattern module setup information *
    '****************************************************************************

    'Check Career Change Working Pattern access
    If gsPersonnelWorkingPatternColumnName <> "" Then
      'Career Change Working Pattern - Static Working Pattern Column
      '///////////////////////////////////////////////
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      If CheckPermission_Columns(glngPersonnelTableID, gsPersonnelTableName, gsPersonnelWorkingPatternColumnName, strTableColumn) Then
        mstrSQLSelect_PersonnelStaticWP = strTableColumn
        strTableColumn = vbNullString
      Else
        GoTo DisableWPs
      End If
      '///////////////////////////////////////////////
      '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

    Else
      'Career Change Working Pattern - Historic Working Pattern Table
      objColumn = GetColumnPrivileges(gsPersonnelHWorkingPatternTableName)

      'Career Change Working Pattern - Historic Working Pattern Column
      pblnColumnOK = objColumn.IsValid(gsPersonnelHWorkingPatternColumnName)
      If pblnColumnOK Then
        pblnColumnOK = objColumn.Item(gsPersonnelHWorkingPatternColumnName).AllowSelect
      End If
      If pblnColumnOK = False Then
        GoTo DisableWPs
      End If

      'Career Change Working Pattern - Historic Working Pattern Effective Date Column
      pblnColumnOK = objColumn.IsValid(gsPersonnelHWorkingPatternDateColumnName)
      If pblnColumnOK Then
        pblnColumnOK = objColumn.Item(gsPersonnelHWorkingPatternDateColumnName).AllowSelect
      End If
      If pblnColumnOK = False Then
        GoTo DisableWPs
      End If

    End If

    CheckPermission_WPInfo = True

TidyUpAndExit:
    'UPGRADE_NOTE: Object objTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    objTable = Nothing
    'UPGRADE_NOTE: Object objColumn may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    objColumn = Nothing
    Exit Function

DisableWPs:
    mblnDisableWPs = True
    IncludeWorkingDaysOnly = False
    mblnWorkingPatterns = False
    CheckPermission_WPInfo = False
    GoTo TidyUpAndExit

  End Function


  Public Function GetEventsCollection() As Boolean

    On Error GoTo Error_Trap

    Dim sSQL As String
    Dim intTemp As Short
    Dim rsTemp As ADODB.Recordset
    Dim lngTableID As Integer

    Dim sTempTableName As String
    Dim sTempStartDateName As String
    Dim sTempStartSessionName As String
    Dim sTempEndDateName As String
    Dim sTempEndSessionName As String
    Dim sTempDurationName As String
    Dim sTempLegendTableName As String
    Dim sTempLegendColumnName As String
    Dim sTempLegendCodeName As String
    Dim sTempLegendEventTypeName As String
    Dim sTempDesc1Name As String
    Dim sTempDesc2Name As String

    ' Get the column information from the Details table, in order
    sSQL = "SELECT * FROM AsrSysCalendarReportEvents WHERE " & "CalendarReportID = " & CStr(mlngCalendarReportID) & " ORDER BY Name ASC "

    rsTemp = mclsData.OpenRecordset(sSQL, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

    With rsTemp
      If .BOF And .EOF Then
        mstrErrorString = "No events found in the specified Calendar Report definition." & vbNewLine & "Please remove this definition and create a new one."
        GetEventsCollection = False
        GoTo TidyUpAndExit
      End If

      mcolEvents = New clsCalendarEvents

      Do Until .EOF

        sTempTableName = datGeneral.GetTableName(.Fields("TableID").Value)

        If .Fields("EventStartDateID").Value > 0 Then
          sTempStartDateName = datGeneral.GetColumnName(.Fields("EventStartDateID").Value)
        Else
          GetEventsCollection = False
          GoTo TidyUpAndExit
        End If

        If .Fields("EventStartSessionID").Value > 0 Then
          sTempStartSessionName = datGeneral.GetColumnName(.Fields("EventStartSessionID").Value)
        Else
          sTempStartSessionName = vbNullString
        End If

        If .Fields("EventEndDateID").Value > 0 Then
          sTempEndDateName = datGeneral.GetColumnName(.Fields("EventEndDateID").Value)
        Else
          sTempEndDateName = vbNullString
        End If

        If .Fields("EventEndSessionID").Value > 0 Then
          sTempEndSessionName = datGeneral.GetColumnName(.Fields("EventEndSessionID").Value)
        Else
          sTempEndSessionName = vbNullString
        End If

        If .Fields("EventDurationID").Value > 0 Then
          sTempDurationName = datGeneral.GetColumnName(.Fields("EventDurationID").Value)
        Else
          sTempDurationName = vbNullString
        End If

        If .Fields("LegendLookupTableID").Value > 0 Then
          sTempLegendTableName = datGeneral.GetTableName(.Fields("LegendLookupTableID").Value)
        Else
          sTempLegendTableName = vbNullString
        End If

        If .Fields("LegendLookupColumnID").Value > 0 Then
          sTempLegendColumnName = datGeneral.GetColumnName(.Fields("LegendLookupColumnID").Value)
        Else
          sTempLegendColumnName = vbNullString
        End If

        If .Fields("LegendLookupCodeID").Value > 0 Then
          sTempLegendCodeName = datGeneral.GetColumnName(.Fields("LegendLookupCodeID").Value)
        Else
          sTempLegendCodeName = vbNullString
        End If

        If .Fields("LegendEventColumnID").Value > 0 Then
          sTempLegendEventTypeName = datGeneral.GetColumnName(.Fields("LegendEventColumnID").Value)
        Else
          sTempLegendEventTypeName = vbNullString
        End If

        If .Fields("EventDesc1ColumnID").Value > 0 Then
          sTempDesc1Name = datGeneral.GetColumnName(.Fields("EventDesc1ColumnID").Value)
        Else
          sTempDesc1Name = vbNullString
        End If

        If .Fields("EventDesc2ColumnID").Value > 0 Then
          sTempDesc2Name = datGeneral.GetColumnName(.Fields("EventDesc2ColumnID").Value)
        Else
          sTempDesc2Name = vbNullString
        End If

        mcolEvents.Add(.Fields("EventKey").Value, .Fields("Name").Value, .Fields("TableID").Value, sTempTableName, .Fields("FilterID").Value, .Fields("EventStartDateID").Value, sTempStartDateName, .Fields("EventStartSessionID").Value, sTempStartSessionName, .Fields("EventEndDateID").Value, sTempEndDateName, .Fields("EventEndSessionID").Value, sTempEndSessionName, .Fields("EventDurationID").Value, sTempDurationName, .Fields("LegendType").Value, .Fields("LegendCharacter").Value, .Fields("LegendLookupTableID").Value, sTempLegendTableName, .Fields("LegendLookupColumnID").Value, sTempLegendColumnName, .Fields("LegendLookupCodeID").Value, sTempLegendCodeName, .Fields("LegendEventColumnID").Value, sTempLegendEventTypeName, .Fields("EventDesc1ColumnID").Value, sTempDesc1Name, .Fields("EventDesc2ColumnID").Value, sTempDesc2Name)

        .MoveNext()
      Loop
      .Close()
    End With

    GetEventsCollection = True

TidyUpAndExit:
    'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsTemp = Nothing
    Exit Function

Error_Trap:
    GetEventsCollection = False
    mstrErrorString = "Error whilst retrieving the event details recordsets'." & vbNewLine & Err.Description
    GoTo TidyUpAndExit

  End Function

  Public Function GenerateSQL() As Boolean

    Dim fOK As Boolean
    Dim objEvent As clsCalendarEvent
    Dim rsLegendBreakdown As ADODB.Recordset

    Dim strSQL As String
    Dim strDynamicKey As String
    Dim strDynamicName As String

    fOK = True
    mintDynamicEventCount = 0

    'loop through the events col and UNION the Event queries together
    For Each objEvent In mcolEvents.Collection

      mblnHasEventFilterIDs = False
      mstrEventFilterIDs = vbNullString

      With objEvent
        If (.LegendType = 1) And (.LegendTableID > 0) Then
          'Event is using a lookup table to find the calendar code for the event.
          'Therefore use the unique types from the legend information.

          strSQL = vbNullString
          strSQL = "SELECT DISTINCT " & .LegendTableName & "." & .LegendColumnName & vbNewLine
          strSQL = strSQL & " FROM " & .LegendTableName & vbNewLine

          rsLegendBreakdown = datGeneral.GetRecords(strSQL)

          If rsLegendBreakdown.BOF And rsLegendBreakdown.EOF Then
            mstrErrorString = "The '" & .LegendTableName & "' event lookup table contains no records."
            GenerateSQL = False
            Exit Function
          End If

          rsLegendBreakdown.MoveFirst()
          Do While Not rsLegendBreakdown.EOF
            mintDynamicEventCount = mintDynamicEventCount + 1

            strDynamicKey = "DYNAMICEVENT" & CStr(mintDynamicEventCount)
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            strDynamicName = Replace(IIf(IsDBNull(rsLegendBreakdown.Fields(CStr(.LegendColumnName)).Value), "", rsLegendBreakdown.Fields(CStr(.LegendColumnName)).Value), "'", "''")

            mstrSQLDynamicLegendWhere = vbNullString

            If fOK Then fOK = GenerateSQLEvent((objEvent.Key), strDynamicKey, strDynamicName)

            If Not fOK Then
              GenerateSQL = False
              Exit Function
            End If

            'mstrSQL = mstrSQL & mstrSQLEvent & " UNION "
            fOK = InsertIntoTempTable(mstrSQLEvent)

            mstrSQLEvent = vbNullString

            rsLegendBreakdown.MoveNext()
          Loop
          rsLegendBreakdown.Close()
          'UPGRADE_NOTE: Object rsLegendBreakdown may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
          rsLegendBreakdown = Nothing

        Else
          If fOK Then fOK = GenerateSQLEvent((objEvent.Key), vbNullString, vbNullString)

          If Not fOK Then
            GenerateSQL = False
            Exit Function
          End If

          'mstrSQL = mstrSQL & mstrSQLEvent & " UNION "
          If fOK Then fOK = InsertIntoTempTable(mstrSQLEvent)

          mstrSQLEvent = vbNullString

        End If
      End With

    Next objEvent

    'remove the last UNION command from the SQL string
    '  mstrSQL = Left(mstrSQL, Len(mstrSQL) - 7)

    GenerateSQL = fOK

TidyUpAndExit:
    'UPGRADE_NOTE: Object objEvent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    objEvent = Nothing

  End Function

  Private Function GenerateSQLSelect(ByRef pstrEventKey As String, ByRef pstrDynamicKey As String, ByRef pstrDynamicName As String) As Boolean

    ' Purpose : This function compiles the SQLSelect string looping
    '           thru the column details recordset.

    On Error GoTo GenerateSQLSelect_ERROR

    Dim objEvent As clsCalendarEvent

    Dim strColList As String
    Dim strBaseColList As String

    Dim strLookupTableName As String
    Dim strLookupColumnName As String
    Dim strLookupCodeName As String
    Dim strEventType As String
    Dim strLegendSQL As String
    Dim strTableColumn As String
    Dim strRegionSQL As String
    Dim lngTempTableID As Integer
    Dim strTempTableName As String
    Dim strTempColumnName As String

    Dim strTempStartSession As String
    Dim strTempEndSession As String

    'Get the Base ID column values so that these can be used in the group by clause when checking
    'for multiple events in MultipleCheck().
    mstrSQLCreateTable = mstrSQLCreateTable & "[BaseID] [Integer] NOT NULL, "
    strColList = strColList & "[" & mstrBaseTableRealSource & "].[ID] AS 'BaseID', " & vbNewLine

    If mlngDescription1 > 0 Then
      If CheckColumnPermissions(mlngCalendarReportsBaseTable, mstrCalendarReportsBaseTableName, mstrDescription1, strTableColumn) Then
        strColList = strColList & " CONVERT(varchar," & strTableColumn & ") AS 'Description1', " & vbNewLine
        strBaseColList = strBaseColList & " CONVERT(varchar," & strTableColumn & ") AS 'Description1', " & vbNewLine
        strTableColumn = vbNullString
      Else
        GenerateSQLSelect = False
        GoTo TidyUpAndExit
      End If
    Else
      strBaseColList = strBaseColList & "NULL AS 'Description1', " & vbNewLine
      strColList = strColList & "NULL AS 'Description1', " & vbNewLine
    End If
    mstrSQLCreateTable = mstrSQLCreateTable & "[Description1] [varchar] (MAX) NULL, "

    If mlngDescription2 > 0 Then
      If CheckColumnPermissions(mlngCalendarReportsBaseTable, mstrCalendarReportsBaseTableName, mstrDescription2, strTableColumn) Then
        strColList = strColList & "CONVERT(varchar," & strTableColumn & ") AS 'Description2', " & vbNewLine
        strBaseColList = strBaseColList & "CONVERT(varchar," & strTableColumn & ") AS 'Description2', " & vbNewLine
        strTableColumn = vbNullString
      Else
        GenerateSQLSelect = False
        GoTo TidyUpAndExit
      End If
    Else
      strBaseColList = strBaseColList & "NULL AS 'Description2', " & vbNewLine
      strColList = strColList & "NULL AS 'Description2', " & vbNewLine
    End If
    mstrSQLCreateTable = mstrSQLCreateTable & "[Description2] [varchar] (MAX) NULL, "

    If mlngDescriptionExpr > 0 Then
      If mblnDescCalcCode Then
        strColList = strColList & " " & mstrDescCalcCode & " AS 'DescriptionExpr', " & vbNewLine
        strBaseColList = strBaseColList & " " & mstrDescCalcCode & " AS 'DescriptionExpr', " & vbNewLine

      Else
        If CheckCalculationPermissions(mlngCalendarReportsBaseTable, mlngDescriptionExpr, strTableColumn) Then
          mstrDescCalcCode = strTableColumn
          mblnDescCalcCode = True
          strColList = strColList & " " & strTableColumn & " AS 'DescriptionExpr', " & vbNewLine
          strBaseColList = strBaseColList & " " & strTableColumn & " AS 'DescriptionExpr', " & vbNewLine
          strTableColumn = vbNullString
        Else
          GenerateSQLSelect = False
          GoTo TidyUpAndExit
        End If

      End If
    Else
      strBaseColList = strBaseColList & "NULL AS 'DescriptionExpr', " & vbNewLine
      strColList = strColList & "NULL AS 'DescriptionExpr', " & vbNewLine
    End If

    'need to set the type of the expression column for the CREAT TABLE...statement.
    Select Case mlngBaseDescriptionType
      Case modExpression.ExpressionValueTypes.giEXPRVALUE_NUMERIC, modExpression.ExpressionValueTypes.giEXPRVALUE_BYREF_NUMERIC
        mstrSQLCreateTable = mstrSQLCreateTable & "[DescriptionExpr] [float] NULL, "
      Case Else
        mstrSQLCreateTable = mstrSQLCreateTable & "[DescriptionExpr] [varchar] (MAX) NULL, "
    End Select


    objEvent = mcolEvents.Item(pstrEventKey)
    With objEvent
      If pstrDynamicKey <> vbNullString Then
        strColList = strColList & "'" & Replace(pstrDynamicKey, "'", "''") & "' AS '?ID_EventID', " & vbNewLine
      Else
        strColList = strColList & "'" & Replace(.Key, "'", "''") & "' AS '?ID_EventID', " & vbNewLine
      End If
      mstrSQLCreateTable = mstrSQLCreateTable & "[?ID_EventID] [varchar] (255) NULL, "

      If pstrDynamicName <> vbNullString Then
        strColList = strColList & "'" & Replace(pstrDynamicName, "'", "''") & "' AS 'Name', " & vbNewLine
      Else
        strColList = strColList & "'" & Replace(.Name, "'", "''") & "' AS 'Name', " & vbNewLine
      End If
      mstrSQLCreateTable = mstrSQLCreateTable & "[Name] [varchar] (255) NULL, "


      '****************************************************************************
      mlngEventViewColumn = .StartDateID
      If CheckColumnPermissions(.TableID, .TableName, .StartDateName, strTableColumn) Then
        strColList = strColList & strTableColumn & " AS 'StartDate', " & vbNewLine
        mstrSQLBaseStartDateColumn = strTableColumn
        strTableColumn = vbNullString
      Else
        GenerateSQLSelect = False
        Exit Function
      End If

      mstrSQLCreateTable = mstrSQLCreateTable & "[StartDate] [datetime] NULL, "

      If .StartSessionID > 0 Then
        If CheckColumnPermissions(.TableID, .TableName, .StartSessionName, strTableColumn) Then
          mstrSQLBaseStartSessionColumn = strTableColumn
          strTableColumn = vbNullString
        Else
          GenerateSQLSelect = False
          Exit Function
        End If
      End If

      If .EndDateID > 0 Then
        mlngEventViewColumn = .EndDateID
        If CheckColumnPermissions(.TableID, .TableName, .EndDateName, strTableColumn) Then
          mstrSQLBaseEndDateColumn = strTableColumn
          strTableColumn = vbNullString
        Else
          GenerateSQLSelect = False
          Exit Function
        End If
      End If

      If .EndSessionID > 0 Then
        If CheckColumnPermissions(.TableID, .TableName, .EndSessionName, strTableColumn) Then
          mstrSQLBaseEndSessionColumn = strTableColumn
          strTableColumn = vbNullString
        Else
          GenerateSQLSelect = False
          Exit Function
        End If
      End If

      If .DurationID > 0 Then
        mlngEventViewColumn = .DurationID
        If CheckColumnPermissions(.TableID, .TableName, .DurationName, strTableColumn) Then
          mstrSQLBaseDurationColumn = strTableColumn
          strTableColumn = vbNullString
        Else
          GenerateSQLSelect = False
          Exit Function
        End If

        If .StartSessionID > 0 Then
          strColList = strColList & mstrSQLBaseStartSessionColumn & " AS 'StartSession', " & vbNewLine
        Else
          strColList = strColList & "'AM' AS 'StartSession', " & vbNewLine
          mstrSQLBaseStartSessionColumn = "'AM'"
        End If

        '      strColList = strColList & " DATEADD(dd, " & mstrSQLBaseDurationColumn & " - 1 , " & mstrSQLBaseStartDateColumn & ") AS 'EndDate', " & vbNewLine
        strColList = strColList & "CASE " & vbNewLine
        strColList = strColList & "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '1' " & vbNewLine
        strColList = strColList & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '2' " & vbNewLine
        strColList = strColList & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '3' " & vbNewLine
        strColList = strColList & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '4' " & vbNewLine
        strColList = strColList & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '5') THEN " & vbNewLine

        strColList = strColList & "         DATEADD(dd " & vbNewLine
        strColList = strColList & "                 , CONVERT(integer,LEFT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))) " & vbNewLine
        strColList = strColList & "                           , CHARINDEX('.',LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))))- 1 )) " & vbNewLine
        strColList = strColList & "                 , " & mstrSQLBaseStartDateColumn & ") " & vbNewLine

        strColList = strColList & "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '6' " & vbNewLine
        strColList = strColList & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '7' " & vbNewLine
        strColList = strColList & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '8' " & vbNewLine
        strColList = strColList & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '9') " & vbNewLine
        strColList = strColList & "          AND (" & mstrSQLBaseStartSessionColumn & " = 'AM') THEN " & vbNewLine

        strColList = strColList & "         DATEADD(dd " & vbNewLine
        strColList = strColList & "                 , CONVERT(integer,LEFT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))) " & vbNewLine
        strColList = strColList & "                         , CHARINDEX('.',LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))))- 1 )) " & vbNewLine
        strColList = strColList & "                 , " & mstrSQLBaseStartDateColumn & ") " & vbNewLine

        strColList = strColList & "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '6' " & vbNewLine
        strColList = strColList & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '7' " & vbNewLine
        strColList = strColList & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '8' " & vbNewLine
        strColList = strColList & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '9') " & vbNewLine
        strColList = strColList & "          AND (" & mstrSQLBaseStartSessionColumn & " = 'PM') THEN " & vbNewLine

        strColList = strColList & "           DATEADD(dd " & vbNewLine
        strColList = strColList & "                 , CONVERT(integer,LEFT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))) " & vbNewLine
        strColList = strColList & "                         , CHARINDEX('.',LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))))- 1 ) + 1) " & vbNewLine
        strColList = strColList & "                 , " & mstrSQLBaseStartDateColumn & ") " & vbNewLine

        strColList = strColList & "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '0') " & vbNewLine
        strColList = strColList & "          AND (" & mstrSQLBaseStartSessionColumn & " = 'AM') THEN " & vbNewLine

        strColList = strColList & "           DATEADD(dd " & vbNewLine
        strColList = strColList & "                   , CONVERT(integer,LEFT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))) " & vbNewLine
        strColList = strColList & "                         , CHARINDEX('.',LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))))- 1 ) - 1) " & vbNewLine
        strColList = strColList & "                   , " & mstrSQLBaseStartDateColumn & ") " & vbNewLine

        strColList = strColList & "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '0') " & vbNewLine
        strColList = strColList & "          AND (" & mstrSQLBaseStartSessionColumn & " = 'PM') THEN " & vbNewLine

        strColList = strColList & "           DATEADD(dd " & vbNewLine
        strColList = strColList & "                 , CONVERT(integer,LEFT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))) " & vbNewLine
        strColList = strColList & "                               , CHARINDEX('.',LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),10,1))))- 1 )) " & vbNewLine
        strColList = strColList & "                 , " & mstrSQLBaseStartDateColumn & ") " & vbNewLine

        strColList = strColList & "END AS 'EndDate', " & vbNewLine

        If .EndSessionID > 0 Then
          strColList = strColList & mstrSQLBaseEndSessionColumn & " AS 'EndSession', " & vbNewLine
        Else
          '        strColList = strColList & "'PM' AS 'EndSession'," & vbNewLine

          strColList = strColList & "CASE"
          strColList = strColList & "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '1' " & vbNewLine
          strColList = strColList & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '2' " & vbNewLine
          strColList = strColList & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '3' " & vbNewLine
          strColList = strColList & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '4' " & vbNewLine
          strColList = strColList & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '5') THEN " & vbNewLine
          '    -- End Session = Start Session
          strColList = strColList & "           " & mstrSQLBaseStartSessionColumn & " " & vbNewLine

          strColList = strColList & "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '6' " & vbNewLine
          strColList = strColList & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '7' " & vbNewLine
          strColList = strColList & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '8' " & vbNewLine
          strColList = strColList & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '9' " & vbNewLine
          strColList = strColList & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '0') " & vbNewLine
          strColList = strColList & "          AND (" & mstrSQLBaseStartSessionColumn & " = 'AM') THEN " & vbNewLine
          '    -- End Session = "PM"
          strColList = strColList & "           'PM'  " & vbNewLine

          strColList = strColList & "      WHEN  (RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '6' " & vbNewLine
          strColList = strColList & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '7' " & vbNewLine
          strColList = strColList & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '8' " & vbNewLine
          strColList = strColList & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '9' " & vbNewLine
          strColList = strColList & "        OR  RIGHT(LTRIM(RTRIM(STR(ROUND(" & mstrSQLBaseDurationColumn & ",1),28,1))),1) = '0') " & vbNewLine
          strColList = strColList & "          AND (" & mstrSQLBaseStartSessionColumn & " = 'PM') THEN " & vbNewLine
          '    -- End Session = "AM"
          strColList = strColList & "           'AM' " & vbNewLine

          strColList = strColList & "END AS 'EndSession', " & vbNewLine
        End If

        strColList = strColList & mstrSQLBaseDurationColumn & " AS 'Duration', " & vbNewLine

      ElseIf .EndDateID > 0 Then

        If .StartSessionID > 0 Then
          strColList = strColList & mstrSQLBaseStartSessionColumn & " AS 'StartSession', " & vbNewLine
          strTempStartSession = mstrSQLBaseStartSessionColumn
        Else
          strColList = strColList & "'AM' AS 'StartSession'," & vbNewLine
          strTempStartSession = "'AM'"
        End If

        strColList = strColList & mstrSQLBaseEndDateColumn & " AS 'EndDate', " & vbNewLine

        If .EndSessionID > 0 Then
          strColList = strColList & mstrSQLBaseEndSessionColumn & " AS 'EndSession', " & vbNewLine
          strTempEndSession = mstrSQLBaseEndSessionColumn
        Else
          strColList = strColList & "'PM' AS 'EndSession'," & vbNewLine
          strTempEndSession = "'PM'"
        End If

        strColList = strColList & " CASE " & vbNewLine
        strColList = strColList & " WHEN " & strTempStartSession & " = " & strTempEndSession & vbNewLine
        strColList = strColList & "   THEN CONVERT(float,(DATEDIFF(dd, " & mstrSQLBaseStartDateColumn & ", " & mstrSQLBaseEndDateColumn & ") + 0.5)) " & vbNewLine
        strColList = strColList & " ELSE " & vbNewLine
        strColList = strColList & "   CONVERT(float,(DATEDIFF(dd, " & mstrSQLBaseStartDateColumn & ", " & mstrSQLBaseEndDateColumn & ") + 1)) " & vbNewLine
        strColList = strColList & " END AS 'Duration'," & vbNewLine
      Else

        If .StartSessionID > 0 Then
          strColList = strColList & mstrSQLBaseStartSessionColumn & " AS 'StartSession', " & vbNewLine
        Else
          strColList = strColList & "'AM' AS 'StartSession'," & vbNewLine
        End If

        strColList = strColList & mstrSQLBaseStartDateColumn & " AS 'EndDate', " & vbNewLine

        If .StartSessionID > 0 Then
          strColList = strColList & mstrSQLBaseStartSessionColumn & " AS 'EndSession'," & vbNewLine
          strColList = strColList & " 0.5 AS 'Duration', " & vbNewLine
        Else
          strColList = strColList & "'PM' AS 'EndSession'," & vbNewLine
          strColList = strColList & " 1 AS 'Duration', " & vbNewLine
        End If

      End If
      mstrSQLCreateTable = mstrSQLCreateTable & "[StartSession] [varchar] (255) NULL, "
      mstrSQLCreateTable = mstrSQLCreateTable & "[EndDate] [datetime] NULL, "
      mstrSQLCreateTable = mstrSQLCreateTable & "[EndSession] [varchar] (255) NULL, "
      mstrSQLCreateTable = mstrSQLCreateTable & "[Duration] [float] NULL, "
      '****************************************************************************

      If .Description1ID > 0 Then
        lngTempTableID = datGeneral.GetColumnTable(.Description1ID)
        strTempTableName = datGeneral.GetColumnTableName(.Description1ID)
        strTempColumnName = datGeneral.GetColumnName(.Description1ID)
        If CheckColumnPermissions(lngTempTableID, strTempTableName, strTempColumnName, strTableColumn) Then
          strColList = strColList & .Description1ID & " AS 'EventDescription1ColumnID', " & vbNewLine
          strColList = strColList & "'" & .Description1Name & "' AS 'EventDescription1Column', " & vbNewLine

          'TM20030407 Fault 5259 - if logic field...convert to 'Y' or 'N' accordingly.
          If datGeneral.GetDataType(lngTempTableID, .Description1ID) = Declarations.SQLDataType.sqlBoolean Then
            strColList = strColList & "CASE " & strTableColumn & " WHEN 1 THEN 'Y' ELSE 'N' END AS 'EventDescription1', " & vbNewLine
          Else
            strColList = strColList & "CONVERT(varchar(MAX)," & strTableColumn & ") AS 'EventDescription1', " & vbNewLine
          End If

          strTableColumn = vbNullString
        Else
          GenerateSQLSelect = False
          Exit Function
        End If
      Else
        strColList = strColList & "NULL AS 'EventDescription1ColumnID', " & vbNewLine
        strColList = strColList & "NULL AS 'EventDescription1Column', " & vbNewLine
        strColList = strColList & "NULL AS 'EventDescription1', " & vbNewLine
      End If
      mstrSQLCreateTable = mstrSQLCreateTable & "[EventDescription1ColumnID] [int] NULL, "
      mstrSQLCreateTable = mstrSQLCreateTable & "[EventDescription1Column] [varchar] (MAX) NULL, "
      mstrSQLCreateTable = mstrSQLCreateTable & "[EventDescription1] [varchar] (MAX) NULL, "

      If .Description2ID > 0 Then
        lngTempTableID = datGeneral.GetColumnTable(.Description2ID)
        strTempTableName = datGeneral.GetColumnTableName(.Description2ID)
        strTempColumnName = datGeneral.GetColumnName(.Description2ID)
        If CheckColumnPermissions(lngTempTableID, strTempTableName, strTempColumnName, strTableColumn) Then
          strColList = strColList & .Description2ID & " AS 'EventDescription2ColumnID', " & vbNewLine
          strColList = strColList & "'" & .Description2Name & "' AS 'EventDescription2Column', " & vbNewLine

          'TM20030407 Fault 5259 - if logic field...convert to 'Y' or 'N' accordingly.
          If datGeneral.GetDataType(lngTempTableID, .Description2ID) = Declarations.SQLDataType.sqlBoolean Then
            strColList = strColList & "CASE " & strTableColumn & " WHEN 1 THEN 'Y' ELSE 'N' END AS 'EventDescription2', " & vbNewLine
          Else
            strColList = strColList & "CONVERT(varchar(MAX)," & strTableColumn & ") AS 'EventDescription2', " & vbNewLine
          End If

          strTableColumn = vbNullString
        Else
          GenerateSQLSelect = False
          Exit Function
        End If
      Else
        strColList = strColList & "NULL AS 'EventDescription2ColumnID', " & vbNewLine
        strColList = strColList & "NULL AS 'EventDescription2Column', " & vbNewLine
        strColList = strColList & "NULL AS 'EventDescription2', " & vbNewLine
      End If
      mstrSQLCreateTable = mstrSQLCreateTable & "[EventDescription2ColumnID] [int] NULL, "
      mstrSQLCreateTable = mstrSQLCreateTable & "[EventDescription2Column] [varchar] (MAX) NULL, "
      mstrSQLCreateTable = mstrSQLCreateTable & "[EventDescription2] [varchar] (MAX) NULL, "

      If .LegendType = 1 Then
        If CheckColumnPermissions(.LegendTableID, .LegendTableName, .LegendCodeName, strTableColumn) Then
          strLegendSQL = "LEFT((SELECT TOP 1 " & strTableColumn
          strTableColumn = vbNullString
        Else
          GenerateSQLSelect = False
          Exit Function
        End If

        strLegendSQL = strLegendSQL & " FROM " & .LegendTableName

        If CheckColumnPermissions(.LegendTableID, .LegendTableName, .LegendColumnName, strTableColumn) Then
          strLegendSQL = strLegendSQL & " WHERE " & strTableColumn
          strTableColumn = vbNullString
        Else
          GenerateSQLSelect = False
          Exit Function
        End If

        If CheckColumnPermissions(.TableID, .TableName, .LegendEventTypeName, strTableColumn) Then
          strLegendSQL = strLegendSQL & " = " & strTableColumn & "),2) AS 'Legend', "
          mstrSQLDynamicLegendWhere = strTableColumn & " = '" & pstrDynamicName & "' "
          strTableColumn = vbNullString
        Else
          GenerateSQLSelect = False
          Exit Function
        End If

      Else
        strLegendSQL = "'" & Replace(Left(.LegendCharacter, 2), "'", "''") & "' AS 'Legend', "

      End If

      strColList = strColList & strLegendSQL & vbNewLine
      mstrSQLCreateTable = mstrSQLCreateTable & "[Legend] [varchar] (MAX) NULL, "
    End With

    'Add the static region column if required.
    If mlngRegion > 0 Then
      'TM01042004 Fault 8428
      mblnCheckingRegionColumn = True
      If CheckColumnPermissions(mlngCalendarReportsBaseTable, mstrCalendarReportsBaseTableName, mstrRegion, strTableColumn) Then
        strColList = strColList & "CONVERT(varchar," & strTableColumn & ") AS 'Region', " & vbNewLine
        strBaseColList = strBaseColList & "CONVERT(varchar," & strTableColumn & ") AS 'Region', " & vbNewLine
        'TM01042004 Fault 8428
        '      mstrRegionColumnRealSource = Left(strTableColumn, InStr(1, strTableColumn, ".") - 1)
        strTableColumn = vbNullString
      Else
        'TM19112004 Fault 8942
        strColList = strColList & "NULL AS 'Region', "
        strBaseColList = strBaseColList & "NULL AS 'Region', "
        '      GenerateSQLSelect = False
        '      GoTo TidyUpAndExit
      End If
      'TM01042004 Fault 8428
      mblnCheckingRegionColumn = False
    Else
      strColList = strColList & "NULL AS 'Region', "
      strBaseColList = strBaseColList & "NULL AS 'Region', "
    End If
    mstrSQLCreateTable = mstrSQLCreateTable & "[Region] [varchar] (MAX) NULL, "

    If CheckColumnPermissions(mlngCalendarReportsBaseTable, mstrCalendarReportsBaseTableName, "ID", strTableColumn) Then
      strColList = strColList & strTableColumn & " AS '?ID_" & mstrCalendarReportsBaseTableName & "', " & vbNewLine
      strBaseColList = strBaseColList & strTableColumn & " AS '?ID_" & mstrCalendarReportsBaseTableName & "', " & vbNewLine
      strTableColumn = vbNullString
    Else
      GenerateSQLSelect = False
      GoTo TidyUpAndExit
    End If
    mstrSQLCreateTable = mstrSQLCreateTable & "[?ID_" & mstrCalendarReportsBaseTableName & "] [varchar] (255) NULL, "

    'Add the static Working Pattern column if required.
    If (mlngCalendarReportsBaseTable = glngPersonnelTableID) And (modPersonnelSpecifics.gwptWorkingPatternType = modPersonnelSpecifics.WorkingPatternType.wptStaticWPattern) And (Not mblnGroupByDescription) Then
      If CheckColumnPermissions(mlngCalendarReportsBaseTable, mstrCalendarReportsBaseTableName, gsPersonnelWorkingPatternColumnName, strTableColumn) Then
        strColList = strColList & "CONVERT(varchar," & strTableColumn & ") AS 'Working_Pattern', " & vbNewLine
        strBaseColList = strBaseColList & "CONVERT(varchar," & strTableColumn & ") AS 'Working_Pattern', " & vbNewLine
        strTableColumn = vbNullString
      Else
        strColList = strColList & "NULL AS 'Working_Pattern', "
        strBaseColList = strBaseColList & "NULL AS 'Working_Pattern', "
        '      GenerateSQLSelect = False
        '      GoTo TidyUpAndExit
      End If
    Else
      strColList = strColList & "NULL AS 'Working_Pattern', "
      strBaseColList = strBaseColList & "NULL AS 'Working_Pattern', "
    End If
    mstrSQLCreateTable = mstrSQLCreateTable & "[Working_Pattern] [varchar] (255) NULL, "

    Dim intOrderCount As Short
    Dim strOrderColumn As String

    mstrSQLOrderList = vbNullString

    For intOrderCount = 1 To UBound(mvarSortOrder, 2) Step 1
      'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(1, intOrderCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      strOrderColumn = mvarSortOrder(1, intOrderCount)
      strTableColumn = vbNullString
      If CheckColumnPermissions(mlngCalendarReportsBaseTable, mstrCalendarReportsBaseTableName, strOrderColumn, strTableColumn) Then
        strColList = strColList & strTableColumn & " AS 'ORDER_" & CStr(intOrderCount) & "',"
        strBaseColList = strBaseColList & strTableColumn & " AS 'ORDER_" & CStr(intOrderCount) & "',"
        'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        mstrSQLOrderList = mstrSQLOrderList & strTableColumn & " " & mvarSortOrder(2, intOrderCount) & ","
        strTableColumn = vbNullString
        If Not mblnOrderByCreated Then
          'UPGRADE_WARNING: Couldn't resolve default property of object mvarSortOrder(2, intOrderCount). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          mstrSQLOrderBy = mstrSQLOrderBy & " [ORDER_" & CStr(intOrderCount) & "] " & mvarSortOrder(2, intOrderCount) & ","
        End If
      Else
        GenerateSQLSelect = False
        GoTo TidyUpAndExit
      End If
      mstrSQLCreateTable = mstrSQLCreateTable & "[ORDER_" & CStr(intOrderCount) & "] [varchar] (MAX) NULL,"
    Next intOrderCount

    If Not mblnOrderByCreated Then
      mstrSQLOrderBy = Left(mstrSQLOrderBy, Len(mstrSQLOrderBy) - 1)
      mblnOrderByCreated = True
    End If
    mstrSQLCreateTable = Left(mstrSQLCreateTable, Len(mstrSQLCreateTable) - 1)

    strColList = Left(strColList, Len(strColList) - 1)
    strBaseColList = Left(strBaseColList, Len(strBaseColList) - 1)
    mstrSQLOrderList = Left(mstrSQLOrderList, Len(mstrSQLOrderList) - 1)

    ' Start off the select statement
    mstrSQLSelect = "SELECT "
    mstrSQLSelect = mstrSQLSelect & strColList

    mstrSQLBaseData = "SELECT "
    mstrSQLBaseData = mstrSQLBaseData & strBaseColList

    GenerateSQLSelect = True

TidyUpAndExit:
    'UPGRADE_NOTE: Object objEvent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    objEvent = Nothing
    Exit Function

GenerateSQLSelect_ERROR:
    GenerateSQLSelect = False
    mstrErrorString = "Error whilst generating SQL Select statement." & vbNewLine & Err.Description
    GoTo TidyUpAndExit

  End Function

  Private Function CheckCalculationPermissions(ByRef plngBaseTableID As Integer, ByRef plngExprID As Integer, ByRef strSQLRef As String) As Boolean

    'This function checks if the current user has read(select) permissions
    'on this calculation. If the user only has access through views then the
    'relevent views are added to the mlngTableViews() array which in turn
    'are used to create the join part of the query.

    Dim lngTempTableID As Integer
    Dim strTempTableName As String
    Dim strTempColumnName As String
    Dim blnColumnOK As Boolean
    Dim blnFound As Boolean
    Dim blnNoSelect As Boolean
    Dim iLoop1 As Short
    Dim intLoop As Short
    Dim strColumnCode As String
    Dim strSource As String
    Dim intNextIndex As Short
    Dim blnOK As Boolean
    Dim strTable As String
    Dim strColumn As String
    Dim sCalcCode As String
    Dim alngSourceTables(,) As Integer
    Dim objCalcExpr As clsExprExpression

    ' Set flags with their starting values
    blnOK = True

    ' Get the calculation SQL, and the array of tables/views that are used to create it.
    ' Column 1 = 0 if this row is for a table, 1 if it is for a view.
    ' Column 2 = table/view ID.
    ReDim alngSourceTables(2, 0)
    objCalcExpr = New clsExprExpression
    blnOK = objCalcExpr.Initialise(plngBaseTableID, plngExprID, modExpression.ExpressionTypes.giEXPR_RUNTIMECALCULATION, modExpression.ExpressionValueTypes.giEXPRVALUE_UNDEFINED)
    If blnOK Then
      blnOK = objCalcExpr.RuntimeCalculationCode(alngSourceTables, sCalcCode, True, False, mvarPrompts)

      If blnOK And gbEnableUDFFunctions Then
        blnOK = objCalcExpr.UDFCalculationCode(alngSourceTables, mastrUDFsRequired, True)
      End If
    End If

    'The "SELECT ... INTO..." statement errors when it trys to create a column for
    'and empty string. Therefore wrap this empty sting in a CONVERT(varchar... clause if an sql empty string
    'is returned.
    'TM20030521 Fault 5702 - Compare the empty string with the calc code value converted to varchar
    sCalcCode = "CASE WHEN CONVERT(varchar," & sCalcCode & ") = '' " & "THEN CONVERT(varchar," & sCalcCode & ") " & "ELSE " & sCalcCode & " END"

    If blnOK Then
      strSQLRef = sCalcCode

      ' Add the required views to the JOIN code.
      For iLoop1 = 1 To UBound(alngSourceTables, 2)
        If alngSourceTables(1, iLoop1) = 1 Then
          ' Check if view has already been added to the array
          blnFound = False
          For intNextIndex = 1 To UBound(mlngTableViews, 2)
            If mlngTableViews(1, intNextIndex) = 1 And mlngTableViews(2, intNextIndex) = alngSourceTables(2, iLoop1) Then
              blnFound = True
              Exit For
            End If
          Next intNextIndex

          If Not blnFound Then

            ' View hasnt yet been added, so add it !
            intNextIndex = UBound(mlngTableViews, 2) + 1
            ReDim Preserve mlngTableViews(2, intNextIndex)
            mlngTableViews(1, intNextIndex) = 1
            mlngTableViews(2, intNextIndex) = alngSourceTables(2, iLoop1)

          End If
          '********************************************************************************
        ElseIf alngSourceTables(1, iLoop1) = 0 Then
          ' Check if table has already been added to the array
          blnFound = False
          For intNextIndex = 1 To UBound(mlngTableViews, 2)
            If mlngTableViews(1, intNextIndex) = 0 And mlngTableViews(2, intNextIndex) = alngSourceTables(2, iLoop1) Then
              blnFound = True
              Exit For
            End If
          Next intNextIndex

          'Only want to check if the source table is the base table
          ' if we have NOT just found the source table in the array of joined tables.
          If Not blnFound Then
            blnFound = (alngSourceTables(2, iLoop1) = mlngCalendarReportsBaseTable)
          End If

          If Not blnFound Then
            ' table hasnt yet been added, so add it !
            intNextIndex = UBound(mlngTableViews, 2) + 1
            ReDim Preserve mlngTableViews(2, intNextIndex)
            mlngTableViews(1, intNextIndex) = 0
            mlngTableViews(2, intNextIndex) = alngSourceTables(2, iLoop1)
          End If
          '********************************************************************************
        End If
      Next iLoop1
    Else
      ' Permission denied on something in the calculation.
      mstrErrorString = "You do not have permission to use the '" & objCalcExpr.Name & "' calculation."
      CheckCalculationPermissions = False
      Exit Function
    End If

    'UPGRADE_NOTE: Object objCalcExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    objCalcExpr = Nothing

    CheckCalculationPermissions = True

  End Function

  Private Function GenerateSQLFrom() As Boolean
    Dim iLoop As Short
    Dim pobjTableView As CTablePrivilege

    pobjTableView = New CTablePrivilege

    mstrSQLFrom = "FROM " & mstrBaseTableRealSource & vbNewLine
    mstrSQLBaseData = mstrSQLBaseData & " FROM " & mstrBaseTableRealSource & vbNewLine

    GenerateSQLFrom = True

TidyUpAndExit:
    'UPGRADE_NOTE: Object pobjTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    pobjTableView = Nothing
    Exit Function

GenerateSQLFrom_ERROR:
    GenerateSQLFrom = False
    mstrErrorString = "Error in GenerateSQLFrom." & vbNewLine & Err.Description
    GoTo TidyUpAndExit

  End Function

  Private Function GenerateSQLJoin(ByRef pstrEventKey As String, ByRef pstrDynamicKey As String) As Boolean

    On Error GoTo GenerateSQLJoin_ERROR

    Dim objTableView As CTablePrivilege
    Dim objChildTable As CTablePrivilege
    Dim rsTemp As ADODB.Recordset
    Dim objEvent As clsCalendarEvent

    Dim sChildJoinCode As String
    Dim strFilterIDs As String
    Dim sChildJoin As String

    Dim blnOK As Boolean

    Dim i As Short
    Dim intLoop As Short

    Dim bViewContains_StartColumn As Boolean
    Dim bViewContains_EndColumn As Boolean
    Dim bViewContains_DurationColumn As Boolean

    ' First, do the join for all the views etc...

    objEvent = mcolEvents.Item(pstrEventKey)

    For intLoop = 1 To UBound(mlngTableViews, 2)

      If mlngTableViews(1, intLoop) = 0 Then
        objTableView = gcoTablePrivileges.FindTableID(mlngTableViews(2, intLoop))
        bViewContains_StartColumn = True
        bViewContains_EndColumn = True
        bViewContains_DurationColumn = True
      Else
        objTableView = gcoTablePrivileges.FindViewID(mlngTableViews(2, intLoop))
        bViewContains_StartColumn = IsColumnInView((objTableView.ViewID), (objEvent.StartDateID))
        bViewContains_EndColumn = IsColumnInView((objTableView.ViewID), (objEvent.EndDateID))
        bViewContains_DurationColumn = IsColumnInView((objTableView.ViewID), (objEvent.DurationID))
      End If

      If (objTableView.TableID = mlngCalendarReportsBaseTable) And (objTableView.ViewID > 0) Or (datGeneral.IsAParentOf((objTableView.TableID), mlngCalendarReportsBaseTable)) Then
        ' Get the table/view object from the id stored in the array

        If (datGeneral.IsAParentOf((objTableView.TableID), mlngCalendarReportsBaseTable)) Then
          mstrSQLJoin = mstrSQLJoin & " LEFT OUTER JOIN " & objTableView.RealSource & " ON " & mstrBaseTableRealSource & ".ID_" & objTableView.TableID & " = " & objTableView.RealSource & ".ID"

          mstrSQLBaseData = mstrSQLBaseData & " LEFT OUTER JOIN " & objTableView.RealSource & " ON " & mstrBaseTableRealSource & ".ID_" & objTableView.TableID & " = " & objTableView.RealSource & ".ID"

        Else
          mstrSQLJoin = mstrSQLJoin & " LEFT OUTER JOIN " & objTableView.RealSource & " ON " & mstrBaseTableRealSource & ".ID = " & objTableView.RealSource & ".ID"

          mstrSQLBaseData = mstrSQLBaseData & " LEFT OUTER JOIN " & objTableView.RealSource & " ON " & mstrBaseTableRealSource & ".ID = " & objTableView.RealSource & ".ID"
        End If

        If (objTableView.TableID = objEvent.TableID) Then
          'add clause to SQL, so that only dates within the specified range are retrieved.
          If (objEvent.StartDateID > 0) And (objEvent.EndDateID > 0) And bViewContains_StartColumn And bViewContains_EndColumn Then
            'event is defined by start date and end date
            mstrSQLJoin = mstrSQLJoin & " AND ((" & objTableView.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "') AND " & objTableView.RealSource & "." & objEvent.EndDateName & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "'))" & vbNewLine & " OR (" & objTableView.RealSource & "." & objEvent.StartDateName & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "') AND " & objTableView.RealSource & "." & objEvent.EndDateName & " <= convert(datetime, '" & VB6.Format(mdtEndDate, "mm/dd/yyyy") & "'))" & vbNewLine & " OR (((" & objTableView.RealSource & "." & objEvent.StartDateName & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "')) AND (" & objTableView.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & VB6.Format(mdtEndDate, "mm/dd/yyyy") & "'))) AND " & objTableView.RealSource & "." & objEvent.EndDateName & " >= convert(datetime, '" & VB6.Format(mdtEndDate, "mm/dd/yyyy") & "'))" & vbNewLine & " OR (" & objTableView.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "') AND " & objTableView.RealSource & "." & objEvent.EndDateName & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "'))" & vbNewLine
            mstrSQLJoin = mstrSQLJoin & ")" & vbNewLine
            mstrSQLJoin = mstrSQLJoin & " AND (" & objTableView.RealSource & "." & objEvent.EndDateName & " >= " & objTableView.RealSource & "." & objEvent.StartDateName & ")" & vbNewLine

          ElseIf (objEvent.StartDateID > 0) And (objEvent.DurationID > 0) And bViewContains_StartColumn And bViewContains_DurationColumn Then
            'event is defined by start date and duration
            mstrSQLJoin = mstrSQLJoin & " OR (" & objTableView.RealSource & "." & objEvent.StartDateName & " IS NOT NULL AND " & objTableView.RealSource & "." & objEvent.DurationName & " > 0)" & vbNewLine

          ElseIf (objEvent.StartDateID) > 0 And (objEvent.EndDateID < 1) And (objEvent.DurationID < 1) And bViewContains_StartColumn Then
            'event is defined by just the start date - one off event with a range of one
            mstrSQLJoin = mstrSQLJoin & " AND ((" & objTableView.RealSource & "." & objEvent.StartDateName & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "') AND " & objTableView.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & VB6.Format(mdtEndDate, "mm/dd/yyyy") & "'))" & vbNewLine
            mstrSQLJoin = mstrSQLJoin & ")" & vbNewLine

          End If
        End If

      ElseIf (datGeneral.IsAChildOf(mlngTableViews(2, intLoop), mlngCalendarReportsBaseTable)) And (objEvent.TableID = objTableView.TableID) Then
        objChildTable = gcoTablePrivileges.FindTableID(mlngTableViews(2, intLoop))

        If objChildTable.AllowSelect Then
          sChildJoinCode = sChildJoinCode & " INNER JOIN " & objChildTable.RealSource & " ON " & mstrBaseTableRealSource & ".ID = " & objChildTable.RealSource & ".ID_" & mlngCalendarReportsBaseTable

          If (objEvent.FilterID > 0) Then

            'TM20030407 Fault 5257 - only get the filter string once for each event to avoid being prompted
            'more tahn once for the save event if the event is split into dynamic events.
            If mblnHasEventFilterIDs Then
              blnOK = True
            Else
              blnOK = datGeneral.FilteredIDs((objEvent.FilterID), strFilterIDs, mvarPrompts)
              mblnHasEventFilterIDs = blnOK
              mstrEventFilterIDs = strFilterIDs
            End If

            ' Generate any UDFs that are used in this filter
            If blnOK And gbEnableUDFFunctions Then
              datGeneral.FilterUDFs((objEvent.FilterID), mastrUDFsRequired)
            End If

            If blnOK Then
              sChildJoinCode = sChildJoinCode & " AND " & objChildTable.RealSource & ".ID IN (" & mstrEventFilterIDs & ")"
            Else
              ' Permission denied on something in the filter.
              mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(objEvent.FilterID) & "' filter."
              GenerateSQLJoin = False
              GoTo TidyUpAndExit
            End If
          End If

          'add clause to SQL, so that only dates within the specified range are retrieved.
          If (objEvent.StartDateID > 0 And objEvent.EndDateID > 0) Then
            'event is defined by start date and end date
            sChildJoinCode = sChildJoinCode & " AND ((" & objChildTable.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "') AND " & objChildTable.RealSource & "." & objEvent.EndDateName & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "'))" & vbNewLine & " OR (" & objChildTable.RealSource & "." & objEvent.StartDateName & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "') AND " & objChildTable.RealSource & "." & objEvent.EndDateName & " <= convert(datetime, '" & VB6.Format(mdtEndDate, "mm/dd/yyyy") & "'))" & vbNewLine & " OR (((" & objChildTable.RealSource & "." & objEvent.StartDateName & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "')) AND (" & objChildTable.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & VB6.Format(mdtEndDate, "mm/dd/yyyy") & "'))) AND " & objChildTable.RealSource & "." & objEvent.EndDateName & " >= convert(datetime, '" & VB6.Format(mdtEndDate, "mm/dd/yyyy") & "'))" & vbNewLine & " OR (" & objChildTable.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "') AND " & objChildTable.RealSource & "." & objEvent.EndDateName & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "'))" & vbNewLine
            sChildJoinCode = sChildJoinCode & ")" & vbNewLine
            sChildJoinCode = sChildJoinCode & " AND (" & objChildTable.RealSource & "." & objEvent.EndDateName & " >= " & objChildTable.RealSource & "." & objEvent.StartDateName & ") "

          ElseIf (objEvent.StartDateID > 0) And (objEvent.DurationID > 0) Then
            'event is defined by start date and duration
            sChildJoinCode = sChildJoinCode & " AND (" & objChildTable.RealSource & "." & objEvent.StartDateName & " IS NOT NULL AND " & objChildTable.RealSource & "." & objEvent.DurationName & " > 0)" & vbNewLine

          ElseIf objEvent.StartDateID > 0 And (objEvent.EndDateID < 1) And (objEvent.DurationID < 1) Then
            'event is defined by just the start date - one off event with a range of one
            sChildJoinCode = sChildJoinCode & " AND ((" & objChildTable.RealSource & "." & objEvent.StartDateName & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "') AND " & objChildTable.RealSource & "." & objEvent.StartDateName & " <= convert(datetime, '" & VB6.Format(mdtEndDate, "mm/dd/yyyy") & "'))" & vbNewLine
            sChildJoinCode = sChildJoinCode & ")" & vbNewLine

          End If
        End If
      End If

    Next intLoop

    mstrSQLJoin = mstrSQLJoin & sChildJoinCode
    '  mstrSQLBaseData = mstrSQLBaseData & mstrSQLJoin

    GenerateSQLJoin = True

TidyUpAndExit:
    strFilterIDs = vbNullString
    'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    objTableView = Nothing
    'UPGRADE_NOTE: Object objChildTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    objChildTable = Nothing
    'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsTemp = Nothing
    'UPGRADE_NOTE: Object objEvent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    objEvent = Nothing
    Exit Function

GenerateSQLJoin_ERROR:
    GenerateSQLJoin = False
    mstrErrorString = "Error in GenerateSQLJoin." & vbNewLine & Err.Description
    GoTo TidyUpAndExit

  End Function

  Private Function GenerateSQLWhere(ByRef pstrEventKey As String, ByRef pstrDynamicKey As String, ByRef pstrDynamicName As String) As Boolean

    ' Purpose : Generate the where clauses that cope with the joins
    '           NB Need to add the where clauses for filters/picklists etc

    On Error GoTo GenerateSQLWhere_ERROR

    Dim objExpr As clsExprExpression
    Dim rsTemp As New ADODB.Recordset
    Dim objEvent As clsCalendarEvent

    Dim strPickListIDs As String
    Dim strFilterIDs As String

    Dim blnOK As Boolean

    objEvent = mcolEvents.Item(pstrEventKey)

    '*******************************************************************************
    Dim pintLoop As Short
    Dim pobjTableView As CTablePrivilege

    pobjTableView = gcoTablePrivileges.FindTableID(mlngCalendarReportsBaseTable)
    If pobjTableView.AllowSelect = False Then

      ' First put the where clauses in for the joins...only if base table is a top level table
      If UCase(Left(mstrBaseTableRealSource, 6)) <> "ASRSYS" Then

        For pintLoop = 1 To UBound(mlngTableViews, 2)
          ' Get the table/view object from the id stored in the array
          If mlngTableViews(1, pintLoop) = 0 Then
            pobjTableView = gcoTablePrivileges.FindTableID(mlngTableViews(2, pintLoop))
          Else
            pobjTableView = gcoTablePrivileges.FindViewID(mlngTableViews(2, pintLoop))
          End If

          ' dont add where clause for the base/chil/p1/p2 TABLES...only add views here
          If (mlngTableViews(1, pintLoop) = 1) Then
            mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " OR ", " WHERE (") & mstrBaseTableRealSource & ".ID IN (SELECT ID FROM " & pobjTableView.RealSource & ")"
          End If

        Next pintLoop

        If Len(mstrSQLWhere) > 0 Then mstrSQLWhere = mstrSQLWhere & ")"
      End If

    End If
    '*******************************************************************************
    If mlngSingleRecordID > 0 Then
      mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrBaseTableRealSource & ".ID IN (" & mstrSQLIDs & ") "

      ' Now if we are using a picklist, add a where clause for that
      'Get List of IDs from Picklist
    ElseIf mlngCalendarReportsPickListID > 0 Then
      mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrBaseTableRealSource & ".ID IN (" & mstrSQLIDs & ")"

    ElseIf mlngCalendarReportsFilterID > 0 Then
      mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrBaseTableRealSource & ".ID IN (" & mstrSQLIDs & ")"

    End If

    'TM03032004 Fault 8184
    mstrSQLBaseData = mstrSQLBaseData & mstrSQLWhere

    'add clause to SQL, so that only dates within the specified range are retieved.
    If objEvent.TableID = mlngCalendarReportsBaseTable Then
      mstrSQLBaseDateClause = vbNullString

      'add clause to SQL, so that only dates within the specified range are retrieved.
      If objEvent.StartDateID > 0 And objEvent.EndDateID > 0 Then
        'event is defined by start date and end date
        mstrSQLBaseDateClause = mstrSQLBaseDateClause & "((" & mstrSQLBaseStartDateColumn & " <= convert(datetime, '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "') AND " & mstrSQLBaseEndDateColumn & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "'))" & vbNewLine
        mstrSQLBaseDateClause = mstrSQLBaseDateClause & " OR (" & mstrSQLBaseStartDateColumn & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "') AND " & mstrSQLBaseEndDateColumn & " <= convert(datetime, '" & VB6.Format(mdtEndDate, "mm/dd/yyyy") & "'))" & vbNewLine
        mstrSQLBaseDateClause = mstrSQLBaseDateClause & " OR (((" & mstrSQLBaseStartDateColumn & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "')) AND (" & mstrSQLBaseStartDateColumn & " <= convert(datetime, '" & VB6.Format(mdtEndDate, "mm/dd/yyyy") & "'))) AND " & mstrSQLBaseEndDateColumn & " >= convert(datetime, '" & VB6.Format(mdtEndDate, "mm/dd/yyyy") & "'))" & vbNewLine
        mstrSQLBaseDateClause = mstrSQLBaseDateClause & " OR (" & mstrSQLBaseStartDateColumn & " <= convert(datetime, '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "') AND " & mstrSQLBaseEndDateColumn & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "')))" & vbNewLine
        mstrSQLBaseDateClause = mstrSQLBaseDateClause & " AND (" & mstrSQLBaseEndDateColumn & ">=" & mstrSQLBaseStartDateColumn & ")"

      ElseIf (objEvent.StartDateID > 0) And (objEvent.DurationID > 0) Then
        'TM 25/04/2005 - Faults 10039 & 10040 - Check if the Start Date + Duration puts event in the report range.
        'event is defined by start date and duration
        mstrSQLBaseDateClause = mstrSQLBaseDateClause & "    (" & mstrSQLBaseDurationColumn & " > 0)" & vbNewLine
        mstrSQLBaseDateClause = mstrSQLBaseDateClause & "    AND (" & vbNewLine & vbNewLine

        ' 1 Event Start Date before Report Start Date, Duration carrys event into, but not beyond the Report Range.
        mstrSQLBaseDateClause = mstrSQLBaseDateClause & "        (" & mstrSQLBaseStartDateColumn & " < convert(datetime, '" & Replace(VB6.Format(mdtStartDate, "mm/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "')" & vbNewLine
        mstrSQLBaseDateClause = mstrSQLBaseDateClause & "      AND (DATEADD(day, " & mstrSQLBaseDurationColumn & ", " & mstrSQLBaseStartDateColumn & ") >= convert(datetime, '" & Replace(VB6.Format(mdtStartDate, "mm/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "'))" & vbNewLine
        mstrSQLBaseDateClause = mstrSQLBaseDateClause & "        AND (DATEADD(day, " & mstrSQLBaseDurationColumn & ", " & mstrSQLBaseStartDateColumn & ") <= convert(datetime, '" & Replace(VB6.Format(mdtEndDate, "mm/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "')))" & vbNewLine & vbNewLine

        ' 2 Event Start Date within Report Range, Duration carrys event beyond Report End Date.
        mstrSQLBaseDateClause = mstrSQLBaseDateClause & "     OR ((" & mstrSQLBaseStartDateColumn & " >= convert(datetime, '" & Replace(VB6.Format(mdtStartDate, "mm/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "'))" & vbNewLine
        mstrSQLBaseDateClause = mstrSQLBaseDateClause & "        AND (" & mstrSQLBaseStartDateColumn & " <= convert(datetime, '" & Replace(VB6.Format(mdtEndDate, "mm/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "'))" & vbNewLine
        mstrSQLBaseDateClause = mstrSQLBaseDateClause & "      AND (DATEADD(day, " & mstrSQLBaseDurationColumn & ", " & mstrSQLBaseStartDateColumn & ") > convert(datetime, '" & Replace(VB6.Format(mdtEndDate, "mm/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "')))" & vbNewLine & vbNewLine

        ' 3 Event Start Date within Report Range and Duration keeps event within Report Range.
        mstrSQLBaseDateClause = mstrSQLBaseDateClause & "     OR ((" & mstrSQLBaseStartDateColumn & " >= convert(datetime, '" & Replace(VB6.Format(mdtStartDate, "mm/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "'))" & vbNewLine
        mstrSQLBaseDateClause = mstrSQLBaseDateClause & "        AND (" & mstrSQLBaseStartDateColumn & " <= convert(datetime, '" & Replace(VB6.Format(mdtEndDate, "mm/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "'))" & vbNewLine
        mstrSQLBaseDateClause = mstrSQLBaseDateClause & "      AND (DATEADD(day, " & mstrSQLBaseDurationColumn & ", " & mstrSQLBaseStartDateColumn & ") <= convert(datetime, '" & Replace(VB6.Format(mdtEndDate, "mm/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "')))" & vbNewLine & vbNewLine

        ' 4 Event Start Date before Report Start Date and Duration carrys event beyond Report End Date.
        mstrSQLBaseDateClause = mstrSQLBaseDateClause & "     OR ((" & mstrSQLBaseStartDateColumn & " < convert(datetime, '" & Replace(VB6.Format(mdtStartDate, "mm/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "'))" & vbNewLine
        mstrSQLBaseDateClause = mstrSQLBaseDateClause & "      AND (DATEADD(day, " & mstrSQLBaseDurationColumn & ", " & mstrSQLBaseStartDateColumn & ") > convert(datetime, '" & Replace(VB6.Format(mdtEndDate, "mm/dd/yyyy"), CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator, "/") & "')))" & vbNewLine & vbNewLine

        mstrSQLBaseDateClause = mstrSQLBaseDateClause & "        )" & vbNewLine

      ElseIf objEvent.StartDateID > 0 And (objEvent.EndDateID < 1) And (objEvent.DurationID < 1) Then
        'event is defined by just the start date - one off event with a range of one
        mstrSQLBaseDateClause = mstrSQLBaseDateClause & "(" & mstrSQLBaseStartDateColumn & " >= convert(datetime, '" & VB6.Format(mdtStartDate, "mm/dd/yyyy") & "') AND " & mstrSQLBaseStartDateColumn & " <= convert(datetime, '" & VB6.Format(mdtEndDate, "mm/dd/yyyy") & "')) "

      End If

      mstrSQLBaseDateClause = mstrSQLBaseDateClause & " AND (" & mstrSQLBaseStartDateColumn & " IS NOT NULL)"

      If objEvent.FilterID > 0 Then
        blnOK = datGeneral.FilteredIDs((objEvent.FilterID), strFilterIDs, mvarPrompts)

        ' Generate any UDFs that are used in this filter
        If blnOK And gbEnableUDFFunctions Then
          datGeneral.FilterUDFs((objEvent.FilterID), mastrUDFsRequired)
        End If

        If blnOK Then
          mstrSQLWhere = mstrSQLWhere & IIf(Len(mstrSQLWhere) > 0, " AND ", " WHERE ") & mstrBaseTableRealSource & ".ID IN (" & strFilterIDs & ")"
        Else
          ' Permission denied on something in the filter.
          mstrErrorString = "You do not have permission to use the '" & datGeneral.GetFilterName(objEvent.FilterID) & "' filter."
          GenerateSQLWhere = False
          GoTo TidyUpAndExit
        End If
      End If

      If Len(mstrSQLWhere) > 0 Then
        mstrSQLWhere = mstrSQLWhere & " AND (" & mstrSQLBaseDateClause & ") "
      Else
        mstrSQLWhere = mstrSQLWhere & " WHERE " & mstrSQLBaseDateClause
      End If
    End If

    If (Len(mstrSQLDynamicLegendWhere) > 0) Then
      If (Len(mstrSQLWhere) > 0) Then
        mstrSQLWhere = mstrSQLWhere & " AND (" & mstrSQLDynamicLegendWhere & ") "
      Else
        mstrSQLWhere = mstrSQLWhere & " WHERE " & mstrSQLDynamicLegendWhere
      End If
    End If

    mstrSQLDynamicLegendWhere = vbNullString
    mstrSQLBaseStartDateColumn = vbNullString
    mstrSQLBaseEndDateColumn = vbNullString
    mstrSQLBaseDurationColumn = vbNullString

    GenerateSQLWhere = True

TidyUpAndExit:
    'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    rsTemp = Nothing
    Exit Function

GenerateSQLWhere_ERROR:
    GenerateSQLWhere = False
    mstrErrorString = "Error in GenerateSQLWhere." & vbNewLine & Err.Description
    GoTo TidyUpAndExit

  End Function

  Private Function GenerateSQLOrderBy() As Boolean

    ' Purpose : Returns order by string from the sort order array

    On Error GoTo GenerateSQLOrderBy_ERROR

    mstrSQLOrderBy = " ORDER BY " & mstrSQLOrderBy

    mstrSQLBaseData = mstrSQLBaseData & mstrSQLOrderBy

    GenerateSQLOrderBy = True
    Exit Function

GenerateSQLOrderBy_ERROR:

    GenerateSQLOrderBy = False
    mstrErrorString = "Error in GenerateSQLOrderBy." & vbNewLine & Err.Description

  End Function

  Private Function ClearUp() As Boolean

    ' Purpose : To clear all variables/recordsets/references and drops temptable
    ' Input   : None
    ' Output  : True/False success

    ' Definition variables

    On Error GoTo ClearUp_ERROR

    mstrCalendarReportsName = vbNullString
    mlngCalendarReportsBaseTable = 0
    mstrCalendarReportsBaseTableName = vbNullString
    mlngCalendarReportsAllRecords = 0
    mlngCalendarReportsPickListID = 0
    mlngCalendarReportsFilterID = 0

    mlngDescription1 = 0
    mstrDescription1 = vbNullString
    mlngDescription2 = 0
    mstrDescription2 = vbNullString
    mlngRegion = 0
    mstrRegion = vbNullString
    mblnGroupByDescription = False

    mstrStartDate = vbNullString
    mstrEndDate = vbNullString

    mblnShowBankHolidays = False
    mblnShowCaptions = False
    mblnShowWeekends = False
    mblnIncludeWorkingDaysOnly = False
    mblnIncludeBankHolidays = False

    'New Default Output Variables
    mblnOutputPreview = False
    mlngOutputFormat = 0
    mblnOutputScreen = True
    mblnOutputPrinter = False
    mstrOutputPrinterName = vbNullString
    mblnOutputSave = False
    mlngOutputSaveExisting = 0
    mblnOutputEmail = False
    mlngOutputEmailID = 0
    mstrOutputEmailName = vbNullString
    mstrOutputEmailSubject = vbNullString
    mstrOutputEmailAttachAs = vbNullString
    mstrOutputFilename = vbNullString

    mblnDefinitionOwner = False

    ' Recordsets
    'UPGRADE_NOTE: Object mrsCalendarReportsOutput may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mrsCalendarReportsOutput = Nothing
    'UPGRADE_NOTE: Object mrsCalendarBaseInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mrsCalendarBaseInfo = Nothing

    ' SQL strings
    mstrSQLEvent = vbNullString
    mstrSQLSelect = vbNullString
    mstrSQLFrom = vbNullString
    mstrSQLWhere = vbNullString
    mstrSQLJoin = vbNullString
    mstrSQLOrderBy = vbNullString
    mstrSQL = vbNullString
    mstrSQLBaseData = vbNullString
    mstrSQLBaseDateClause = vbNullString
    mstrSQLOrderList = vbNullString
    mstrSQLIDs = vbNullString
    mstrSQLDynamicLegendWhere = vbNullString
    mintDynamicEventCount = 0

    ' Class references
    'UPGRADE_NOTE: Object mclsData may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mclsData = Nothing
    'UPGRADE_NOTE: Object mclsGeneral may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mclsGeneral = Nothing
    'UPGRADE_NOTE: Object mclsUI may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mclsUI = Nothing
    'UPGRADE_NOTE: Object mobjEventLog may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mobjEventLog = Nothing
    'UPGRADE_NOTE: Object mcolBaseDescIndex may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mcolBaseDescIndex = Nothing
    'UPGRADE_NOTE: Object mcolEvents may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mcolEvents = Nothing

    ' Arrays
    Dim mavEventDetails(21, 0) As Object
    Dim mavSortOrder(2, 0) As Object
    ReDim mvarPrompts(1, 0)
    ReDim mastrUDFsRequired(0)
    ReDim mlngTableViews(2, 0)
    ReDim mstrViews(0)
    ReDim mvarTableViews(3, 0)

    ' Column Privilege arrays / collections / variables
    mstrBaseTableRealSource = vbNullString
    mstrRealSource = vbNullString
    'UPGRADE_NOTE: Object mobjTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mobjTableView = Nothing
    'UPGRADE_NOTE: Object mobjColumnPrivileges may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    mobjColumnPrivileges = Nothing

    ClearUp = True

    Exit Function

ClearUp_ERROR:

    mstrErrorString = "Error whilst clearing data." & vbNewLine & "(" & Err.Description & ")"
    ClearUp = False

  End Function

  Public Function IsRecordSelectionValid() As Boolean

    Dim i As Short
    Dim lngFilterID As Integer
    Dim objEvent As clsCalendarEvent
    Dim iResult As modUtilityAccess.RecordSelectionValidityCodes
    Dim fCurrentUserIsSysSecMgr As Boolean

    fCurrentUserIsSysSecMgr = CurrentUserIsSysSecMgr()

    ' Base Table First
    If mlngCalendarReportsFilterID > 0 Then
      iResult = ValidateRecordSelection(modUtilityAccess.RecordSelectionTypes.REC_SEL_FILTER, mlngCalendarReportsFilterID)
      Select Case iResult
        Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
          mstrErrorString = "The base table filter used in this definition has been deleted by another user."
        Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
          mstrErrorString = "The base table filter used in this definition is invalid."
        Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
          If Not fCurrentUserIsSysSecMgr Then
            mstrErrorString = "The base table filter used in this definition has been made hidden by another user."
          End If
      End Select
    ElseIf mlngCalendarReportsPickListID > 0 Then
      iResult = ValidateRecordSelection(modUtilityAccess.RecordSelectionTypes.REC_SEL_PICKLIST, mlngCalendarReportsPickListID)
      Select Case iResult
        Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
          mstrErrorString = "The base table picklist used in this definition has been deleted by another user."
        Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
          mstrErrorString = "The base table picklist used in this definition is invalid."
        Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
          If Not fCurrentUserIsSysSecMgr Then
            mstrErrorString = "The base table picklist used in this definition has been made hidden by another user."
          End If
      End Select
    End If

    'Description Calculation
    If Len(mstrErrorString) = 0 Then
      If mlngDescriptionExpr > 0 Then
        iResult = ValidateCalculation(mlngDescriptionExpr)
        Select Case iResult
          Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
            mstrErrorString = "The base description calculation used in this definition has been deleted by another user."
          Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
            mstrErrorString = "The base description calculation used in this definition is invalid."
          Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
            If Not fCurrentUserIsSysSecMgr Then
              mstrErrorString = "The base description calculation used in this definition has been made hidden by another user."
            End If
        End Select
      End If
    End If

    'Events Filters
    For Each objEvent In mcolEvents.Collection
      lngFilterID = objEvent.FilterID
      If lngFilterID > 0 Then
        iResult = ValidateRecordSelection(modUtilityAccess.RecordSelectionTypes.REC_SEL_FILTER, lngFilterID)
        Select Case iResult
          Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
            mstrErrorString = "An event table filter used in this definition has been deleted by another user."
          Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
            mstrErrorString = "An event table filter used in this definition is invalid."
          Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
            If Not fCurrentUserIsSysSecMgr Then
              mstrErrorString = "An event table filter used in this definition has been made hidden by another user."
            End If
        End Select
      End If
    Next objEvent
    'UPGRADE_NOTE: Object objEvent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
    objEvent = Nothing

    'Report Start Date Calculation
    If Len(mstrErrorString) = 0 Then
      If mlngStartDateExpr > 0 Then
        iResult = ValidateCalculation(mlngStartDateExpr)
        Select Case iResult
          Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
            mstrErrorString = "The report start date calculation used in this definition has been deleted by another user."
          Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
            mstrErrorString = "The report start date calculation used in this definition is invalid."
          Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
            If Not fCurrentUserIsSysSecMgr Then
              mstrErrorString = "The report start date calculation used in this definition has been made hidden by another user."
            End If
        End Select
      End If
    End If

    'Report End Date Calculation
    If Len(mstrErrorString) = 0 Then
      If mlngEndDateExpr > 0 Then
        iResult = ValidateCalculation(mlngEndDateExpr)
        Select Case iResult
          Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_DELETED
            mstrErrorString = "The report end date calculation used in this definition has been deleted by another user."
          Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_INVALID
            mstrErrorString = "The report end date calculation used in this definition is invalid."
          Case modUtilityAccess.RecordSelectionValidityCodes.REC_SEL_VALID_HIDDENBYOTHER
            If Not fCurrentUserIsSysSecMgr Then
              mstrErrorString = "The report end date calculation used in this definition has been made hidden by another user."
            End If
        End Select
      End If
    End If

    IsRecordSelectionValid = (Len(mstrErrorString) = 0)

  End Function

  Public Function EventToolTipText(ByRef pdtStartDate As Date, ByRef pstrStartSession As String, ByRef pdtEndDate As Date, ByRef pstrEndSession As String) As String

    Dim strToolTip As String

    strToolTip = vbNullString
    strToolTip = strToolTip & "Start Date: " & VB6.Format(pdtStartDate, "dd-mmm-yyyy ")
    strToolTip = strToolTip & LCase(pstrStartSession)
    strToolTip = strToolTip & "  --->  "
    strToolTip = strToolTip & "End Date: " & VB6.Format(pdtEndDate, "dd-mmm-yyyy ")
    strToolTip = strToolTip & LCase(pstrEndSession)

    EventToolTipText = strToolTip

  End Function

  Public Function OutputGridColumns() As Boolean

    On Error GoTo ErrTrap

    Dim iLoop As Short
    Dim pblnOK As Boolean
    Dim intColCounter As Short

    pblnOK = True

    ' Now loop through the recordset fields, adding the data columns
    For iLoop = 0 To DAY_CONTROL_COUNT Step 1

      intColCounter = iLoop

      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").Columns.Count"" VALUE=""1"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").Caption"" VALUE=""Column_" & iLoop & """>" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").Name"" VALUE=""Column_" & iLoop & """>" & vbNewLine)

      ' left align strings/dates, centre align logics, right align numerics
      If iLoop = 1 Then
        If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").Alignment"" VALUE=""0"">" & vbNewLine)
      Else
        If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").Alignment"" VALUE=""2"">" & vbNewLine)
      End If

      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").CaptionAlignment"" VALUE=""2"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").Bound"" VALUE=""0"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").AllowSizing"" VALUE=""1"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").DataField"" VALUE=""Column " & iLoop & """>" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").DataType"" VALUE=""8"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").Level"" VALUE=""0"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").NumberFormat"" VALUE="""">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").Case"" VALUE=""0"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").FieldLen"" VALUE=""4096"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").VertScrollBar"" VALUE=""0"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").Locked"" VALUE=""0"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").Style"" VALUE=""0"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").ButtonsAlways"" VALUE=""0"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").RowCount"" VALUE=""0"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").ColCount"" VALUE=""1"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").HasHeadForeColor"" VALUE=""0"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").HasHeadBackColor"" VALUE=""0"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").HasForeColor"" VALUE=""0"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").HasBackColor"" VALUE=""0"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").HeadForeColor"" VALUE=""0"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").HeadBackColor"" VALUE=""0"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").ForeColor"" VALUE=""0"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").BackColor"" VALUE=""0"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").HeadStyleSet"" VALUE="""">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").StyleSet"" VALUE="""">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").Nullable"" VALUE=""1"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").Mask"" VALUE="""">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").PromptInclude"" VALUE=""0"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").ClipMode"" VALUE=""0"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").PromptChar"" VALUE=""95"">" & vbNewLine)
      If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Columns(" & intColCounter & ").Width"" VALUE=""575"">" & vbNewLine)

    Next iLoop

    If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""UseDefaults"" VALUE=""-1"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""TabNavigation"" VALUE=""1"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""_ExtentX"" VALUE=""17330"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""_ExtentY"" VALUE=""1323"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""_StockProps"" VALUE=""79"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Caption"" VALUE="""">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""ForeColor"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""BackColor"" VALUE=""16777215"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""Enabled"" VALUE=""-1"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Columns("        <PARAM NAME=""DataMember"" VALUE="""">" & vbNewLine)

    OutputGridColumns = True

    Exit Function

ErrTrap:

    OutputGridColumns = False
    mstrErrorString = "Error with OutputGridColumns: " & vbNewLine & Err.Description

  End Function

  Public Function OutputGridDefinition() As Boolean

    Dim pblnOK As Boolean

    On Error GoTo ErrTrap

    pblnOK = True

    '  If pblnOK Then pblnOK = AddToArray_Definition("      <OBJECT classid=""clsid:4A4AA697-3E6F-11D2-822F-00104B9E07A1"" id=grdCalendarOutput name=grdCalendarOutput codebase=""cabs/COAInt_Grid.cab#version=1,0,0,0"" style=""LEFT: 0px; TOP: 0px; WIDTH:400; HEIGHT:300"">" & vbNewLine)

    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""FontName"" VALUE=""Tahoma"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""FontSize"" VALUE=""8.25"">" & vbNewLine)

    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""ScrollBars"" VALUE=""4"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""_Version"" VALUE=""196616"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""DataMode"" VALUE=""2"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""Caption"" VALUE=""" & Replace(mstrCalendarReportsName, """", "&quot;") & """>" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""Cols"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""Rows"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""BorderStyle"" VALUE=""1"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""RecordSelectors"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""GroupHeaders"" VALUE=""-1"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""ColumnHeaders"" VALUE=""-1"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""GroupHeadLines"" VALUE=""1"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""HeadLines"" VALUE=""1"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""FieldDelimiter"" VALUE=""(None)"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""FieldSeparator"" VALUE=""(Tab)"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""Col.Count"" VALUE=""" & DAY_CONTROL_COUNT + 2 & """>" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""stylesets.count"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""TagVariant"" VALUE=""EMPTY"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""UseGroups"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""HeadFont3D"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""Font3D"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""DividerType"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""DividerStyle"" VALUE=""1"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""DefColWidth"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""BeveColorScheme"" VALUE=""2"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""BevelColorFrame"" VALUE=""-2147483642"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""BevelColorHighlight"" VALUE=""-2147483628"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""BevelColorShadow"" VALUE=""-2147483632"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""BevelColorFace"" VALUE=""-2147483633"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""CheckBox3D"" VALUE=""-1"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""AllowAddNew"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""AllowDelete"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""AllowUpdate"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""MultiLine"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""ActiveCellStyleSet"" VALUE="""">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""RowSelectionStyle"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""AllowRowSizing"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""AllowGroupSizing"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""AllowColumnSizing"" VALUE=""-1"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""AllowGroupMoving"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""AllowColumnMoving"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""AllowGroupSwapping"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""AllowColumnSwapping"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""AllowGroupShrinking"" VALUE=""-1"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""AllowColumnShrinking"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""AllowDragDrop"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""UseExactRowCount"" VALUE=""-1"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""SelectTypeCol"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""SelectTypeRow"" VALUE=""1"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""SelectByCell"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""BalloonHelp"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""RowNavigation"" VALUE=""1"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""CellNavigation"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""MaxSelectedRows"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""HeadStyleSet"" VALUE="""">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""StyleSet"" VALUE="""">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""ForeColorEven"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""ForeColorOdd"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""BackColorEven"" VALUE=""16777215"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""BackColorOdd"" VALUE=""16777215"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""Levels"" VALUE=""1"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""RowHeight"" VALUE=""239"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""ExtraHeight"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""ActiveRowStyleSet"" VALUE="""">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""CaptionAlignment"" VALUE=""2"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""SplitterPos"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""SplitterVisible"" VALUE=""0"">" & vbNewLine)
    If pblnOK Then pblnOK = AddToArray_Definition("        <PARAM NAME=""Columns.Count"" VALUE=""" & (DAY_CONTROL_COUNT + 1) & """>" & vbNewLine)

    OutputGridDefinition = pblnOK

    Exit Function

ErrTrap:

    OutputGridDefinition = False
    mstrErrorString = "Error with OutputGridDefinition: " & vbNewLine & Err.Description

  End Function

  Public Function ConvertDescription(ByRef pvarDesc1 As Object, ByRef pvarDesc2 As Object, ByRef pvarDesc3 As Object) As String

    Dim strBaseDescription1, strBaseDescription2 As Object
    Dim strBaseDescriptionExpr As String
    Dim strTempRecordDesc As String

    'Get base description 1
    'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
    If Not IsDBNull(pvarDesc1) Then
      Select Case mintType_BaseDesc1
        Case 3
          'UPGRADE_WARNING: Couldn't resolve default property of object pvarDesc1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          strBaseDescription1 = VB6.Format(pvarDesc1, mstrFormat_BaseDesc1)
        Case 2
          'UPGRADE_WARNING: Couldn't resolve default property of object pvarDesc1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          strBaseDescription1 = IIf(pvarDesc1, "Y", "N")
        Case 1
          'UPGRADE_WARNING: Couldn't resolve default property of object pvarDesc1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          strBaseDescription1 = VB6.Format(pvarDesc1, mstrClientDateFormat)
        Case 0
          'UPGRADE_WARNING: Couldn't resolve default property of object pvarDesc1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          strBaseDescription1 = pvarDesc1
      End Select
    Else
      'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      strBaseDescription1 = vbNullString
    End If
    ' Get base description 2
    'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
    If Not IsDBNull(pvarDesc2) Then
      Select Case mintType_BaseDesc2
        Case 3
          'UPGRADE_WARNING: Couldn't resolve default property of object pvarDesc2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          strBaseDescription2 = VB6.Format(pvarDesc2, mstrFormat_BaseDesc2)
        Case 2
          'UPGRADE_WARNING: Couldn't resolve default property of object pvarDesc2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          strBaseDescription2 = IIf(pvarDesc2, "Y", "N")
        Case 1
          'UPGRADE_WARNING: Couldn't resolve default property of object pvarDesc2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          strBaseDescription2 = VB6.Format(pvarDesc2, mstrClientDateFormat)
        Case 0
          'UPGRADE_WARNING: Couldn't resolve default property of object pvarDesc2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          strBaseDescription2 = pvarDesc2
      End Select
    Else
      'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
      strBaseDescription2 = vbNullString
    End If
    ' Get base description expression
    'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
    If Not IsDBNull(pvarDesc3) Then
      Select Case mintType_BaseDescExpr
        Case 2
          'UPGRADE_WARNING: Couldn't resolve default property of object pvarDesc3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          strBaseDescriptionExpr = IIf(pvarDesc3, "Y", "N")
        Case 1
          'UPGRADE_WARNING: Couldn't resolve default property of object pvarDesc3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          strBaseDescriptionExpr = VB6.Format(pvarDesc3, mstrClientDateFormat)
        Case 0
          'UPGRADE_WARNING: Couldn't resolve default property of object pvarDesc3. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
          strBaseDescriptionExpr = pvarDesc3
      End Select
    Else
      strBaseDescriptionExpr = vbNullString
    End If

    'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    strTempRecordDesc = strBaseDescription1
    'UPGRADE_WARNING: Couldn't resolve default property of object strBaseDescription2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    strTempRecordDesc = strTempRecordDesc & IIf((Len(strTempRecordDesc) > 0) And (Len(strBaseDescription2) > 0), mstrDescriptionSeparator, "") & strBaseDescription2
    strTempRecordDesc = strTempRecordDesc & IIf((Len(strTempRecordDesc) > 0) And (Len(strBaseDescriptionExpr) > 0), mstrDescriptionSeparator, "") & strBaseDescriptionExpr

    ConvertDescription = strTempRecordDesc

  End Function
End Class