Attribute VB_Name = "modEnums"

Public Enum OfficeApp
  oaWord = 0
  oaExcel = 1
End Enum

' Microsoft Word Output Types
Public Enum WordOutputType
  wdFormatDocument = 0
  wdFormatDOSText = 4
  wdFormatDOSTextLineBreaks = 5
  wdFormatEncodedText = 7
  wdFormatFilteredHTML = 10
  wdFormatHTML = 8
  wdFormatRTF = 6
  wdFormatTemplate = 1
  wdFormatText = 2
  wdFormatTextLineBreaks = 3
  wdFormatUnicodeText = 7
  wdFormatWebArchive = 9
  wdFormatXML = 11
  wdFormatDocument97 = 0
  wdFormatDocumentDefault = 16
  wdFormatPDF = 17
  wdFormatTemplate97 = 1
  wdFormatXMLDocument = 12
  wdFormatXMLDocumentMacroEnabled = 13
  wdFormatXMLTemplate = 14
  wdFormatXMLTemplateMacroEnabled = 15
  wdFormatXPS = 18
End Enum

' Microsoft Excel output types
Public Enum ExcelOutputType
  xlHtml = 44
  xlOpenXMLWorkbook = 51
  xlExcel8 = 56
End Enum

' Utility Types
Public Enum UtilityType
  utlAll = -1
  utlBatchJob = 0
  utlCrossTab = 1
  utlCustomReport = 2
  utlDataTransfer = 3
  utlExport = 4
  UtlGlobalAdd = 5
  utlGlobalDelete = 6
  utlGlobalUpdate = 7
  utlImport = 8
  utlMailMerge = 9
  utlPicklist = 10
  utlFilter = 11
  utlCalculation = 12
  utlOrder = 13
  utlMatchReport = 14
  utlAbsenceBreakdown = 15
  utlBradfordFactor = 16
  utlCalendarReport = 17
  utlLabel = 18
  utlLabelType = 19
  utlRecordProfile = 20
  utlEmailAddress = 21
  utlEmailGroup = 22
  utlSuccession = 23
  utlCareer = 24
  utlWorkflow = 25
  utlWorkFlowPendingSteps = 26
  utlOrderDefinition = 27
  utlDocumentMapping = 28
  utlReportPack = 29
  utlTurnover = 30
  utlStability = 31
  utlScreen = 32
  utlTable = 33
  utlColumn = 34
  utlNineBoxGrid = 35
End Enum

' Screen Types
Public Enum ScreenType
  screenParentTable = 1
  screenParentView = 2
  screenHistoryTable = 4
  screenHistoryView = 8
  screenLookup = 16
  screenFind = 32
  screenHistorySummary = 64
  screenQuickEntry = 128
  screenPickList = 256
End Enum

Public Enum GlobalType
    glAdd = 1
    glUpdate = 2
    glDelete = 3
End Enum

'Edit option constants
Public Enum EditOptions
  edtCancel = 0
  edtAdd = 2 ^ 10
  edtDelete = 2 ^ 11
  edtEdit = 2 ^ 12
  edtCopy = 2 ^ 13
  edtSelect = 2 ^ 14
  edtDeselect = 2 ^ 15
  edtPrint = 2 ^ 16
  edtProperties = 2 ^ 17
  edtRefresh = 2 ^ 18
End Enum

'Standard report constants
Public Enum ReportOptions
  rptCancel = 0
  rptOK = 1
  rptRun = 2
End Enum

'Table type constants
Public Enum TableTypes
  tabTopLevel = 1
  tabChild = 2
  tabLookup = 3
End Enum

'Control type constants
Public Enum ControlTypes
  ctlCheck = 1
  ctlCombo = 2
  ctlImage = 4
  ctlOle = 8
  ctlRadio = 16
  ctlSpin = 32
  ctlText = 64
  ctlTab = 128
  ctlLabel = 256
  ctlFrame = 512
  ctlPhoto = 1024
  ctlCommand = 2048
  ctlWorkingPattern = 4096
  ctlLine = 2 ^ 13
  ctlNavigation = 2 ^ 14
  ctlColourPicker = 2 ^ 15
End Enum
  
'Column type constants
Public Enum ColumnTypes
  ColData = 0
  colLookup = 1
  colCalc = 2
  colSystem = 3
  colLink = 4
'  colWorkingPattern = 5
End Enum

'SQL DatType
Public Enum SQLDataType
  sqlUnknown = 0      ' ?
  sqlOle = -4         ' OLE columns
  sqlBoolean = -7     ' Logic columns
  sqlNumeric = 2      ' Numeric columns
  sqlInteger = 4      ' Integer columns
  sqlDate = 11        ' Date columns
  sqlVarChar = 12     ' Character columns
  sqlVarBinary = -3   ' Photo columns
  sqlLongVarChar = -1 ' Working Pattern columns
End Enum

'Case Conversion Types
Public Enum CaseConvert
    convNone = 0
    convUpper = 1
    convLower = 2
    convProper = 3
End Enum
  
Public Enum test
  globfuncvaltyp_STRAIGHTVALUE = 1
  globfuncvaltyp_LOOKUPTABLE = 2
  globfuncvaltyp_FIELD = 3
  globfuncvaltyp_CALCULATION = 4
End Enum

Public Enum FilterOperators
  giFILTEROP_UNDEFINED = 0
  giFILTEROP_EQUALS = 1
  giFILTEROP_NOTEQUALTO = 2
  giFILTEROP_ISATMOST = 3
  giFILTEROP_ISATLEAST = 4
  giFILTEROP_ISMORETHAN = 5
  giFILTEROP_ISLESSTHAN = 6
  giFILTEROP_ON = 7
  giFILTEROP_NOTON = 8
  giFILTEROP_AFTER = 9
  giFILTEROP_BEFORE = 10
  giFILTEROP_ONORAFTER = 11
  giFILTEROP_ONORBEFORE = 12
  giFILTEROP_CONTAINS = 13
  giFILTEROP_IS = 14
  giFILTEROP_DOESNOTCONTAIN = 15
  giFILTEROP_ISNOT = 16
End Enum

' Record profile orientation constants
Public Enum OrientationTypes
  giHORIZONTAL = 0
  giVERTICAL = 1
End Enum

' Navigation execution types
Public Enum NavigateIn
  URL = 0
  MenuBar = 1
  Db = 2
End Enum


' ------------
' clsEventLog
' ------------
Public Enum EventLog_Type
  eltCrossTab = 1
  eltCustomReport = 2
  eltDataTransfer = 3
  eltExport = 4
  eltGlobalAdd = 5
  eltGlobalDelete = 6
  eltGlobalUpdate = 7
  eltImport = 8
  eltMailMerge = 9
  eltDiaryDelete = 10
  eltDiaryRebuild = 11
  eltEmailRebuild = 12
  eltStandardReport = 13  'MH20010305
  eltRecordEditing = 14
  eltSystemError = 15
  eltMatchReport = 16
  eltCalandarReport = 17
  eltLabel = 18
  eltLabelType = 19
  eltRecordProfile = 20
  eltSuccessionPlanning = 21
  eltCareerProgression = 22
  eltAccordImport = 23
  eltAccordExport = 24
  eltWorkflowRebuild = 25
  elt9BoxGrid = 35
  
End Enum

Public Enum EventLog_Status
  elsPending = 0
  elsCancelled = 1
  elsFailed = 2
  elsSuccessful = 3
  elsSkipped = 4
  elsError = 5
End Enum


' ------------
' clsLicence
' ------------
Public Enum Module
  modPersonnel = 1
  modRecruitment = 2
  modAbsence = 4
  modTraining = 8
  modIntranet = 16
  modAFD = 32
  modFullSysMgr = 64
  modCMG = 128
  modQAddress = 256
  modAccord = 512
  modWorkflow = 1024
  modVersionOne = 2048
  modMobile = 4096
  modFusion = 8192
  modXMLExport = 16384
  mod3rdPartyTables = 32768
  modNineBoxGrid = 2 ^ 16
  modEditableGrids = 2 ^ 17
  modCustomisationPowerPack = 2 ^ 18
End Enum


' ------------
' frmAccordViewTransfers
' ------------
Public Enum AccordViewMode
  iLIVE_ALL = 0
  iARCHIVE_ALL = 1
  iCURRENT_RECORD = 2
End Enum


' ------------
' clsODBC
' ------------
Public Enum SQLInfo
  SQL_USER_NAME = 47
  SQL_DATABASE_NAME = 16
  SQL_DBMS_NAME = 17
  SQL_DBMS_VER = 18
  SQL_KEYWORDS = 89
  SQL_MAX_COLUMN_NAME_LEN = 30
  SQL_MAX_COLUMNS_IN_TABLE = 101
  SQL_MAX_TABLE_NAME_LEN = 35
  SQL_SERVER_NAME = 13
End Enum


' ------------
' clsUI
' ------------
Public Enum SystemMetrics
  SM_CXVSCROLL = 2
  SM_CYCAPTION = 4
  SM_CXBORDER = 5
  SM_CYBORDER = 6
  SM_CXFRAME = 32
  SM_CYFRAME = 33
  SM_CYSMCAPTION = 51
End Enum


' ------------
' modHRPro
' ------------
Public Enum LockType
  giLOCKTYPE_PHOTO = 1
  giLOCKTYPE_OLE = 2
  giLOCKTYPE_CRYSTAL = 4
End Enum

Public Enum DefaultDisplay
  disRecEdit_New = 1
  disRecEdit_First = 2
  disFindWindow = 3
End Enum

Public Enum OutputFormats
  fmtDataOnly = 0
  fmtCSV = 1
  fmtHTML = 2
  fmtWordDoc = 3
  fmtExcelWorksheet = 4
  fmtExcelchart = 5
  fmtExcelPivotTable = 6
  fmtFixedLengthFile = 7
  fmtCMGFile = 8
  fmtXML = 9
  fmtSQLTable = 99
End Enum

Public Enum OutputDestinations
  desScreen = 0
  desPrinter = 1
  desSave = 2
  desEmail = 3
End Enum

Public Enum OLEType
  OLE_LOCAL = 0
  OLE_SERVER = 1
  OLE_EMBEDDED = 2
  OLE_UNC = 3
End Enum

Public Enum mceIDLPaths
  CSIDL_INTERNET_CACHE = &H20 ' * CSIDL_INTERNET_CACHE - File system directory for temporary Internet files.
End Enum

'Background type constants
Public Enum BackgroundLocationTypes
  giLOCATION_TOPLEFT = 0
  giLOCATION_TOPRIGHT = 1
  giLOCATION_CENTRE = 2
  giLOCATION_LEFTTILE = 3
  giLOCATION_RIGHTTILE = 4
  giLOCATION_TOPTILE = 5
  giLOCATION_BOTTOMTILE = 6
  giLOCATION_TILE = 7
End Enum

Public Enum CrossTabType
  cttNormal = 0
  cttTurnover = 1
  cttStability = 2
  cttAbsenceBreakdown = 3
  ctt9GridBox = 4
End Enum

Public Enum MatchReportType
  mrtNormal = 0
  mrtSucession = 1
  mrtCareer = 2
End Enum

' Character trimming types
Public Enum TrimmingTypes
  giTRIMMING_NONE = 0
  giTRIMMING_LEFTRIGHT = 1
  giTRIMMING_LEFTONLY = 2
  giTRIMMING_RIGHTONLY = 3
End Enum

Public Enum ReturnPrintDateType
  RETURN_DEFAULT = 0
  RETURN_MANUAL = 1
  RETURN_CALCULATION = 2
End Enum

Public Enum ToolbarPositions
  giTOOLBAR_NONE = 0
  giTOOLBAR_TOP = 1
  giTOOLBAR_BOTTOM = 2
  giTOOLBAR_LEFT = 4
  giTOOLBAR_RIGHT = 8
  giTOOLBAR_FLOAT = 16
  giTOOLBAR_POPUP = 32
End Enum

' Navigation Display Types
Public Enum NavigationDisplayType
  Hyperlink = 0
  Button = 1
  Browser = 2
  Hidden = 3
End Enum

' Direction of file formats (used to initialise common dialog)
Public Enum FileFormatDirection
  DirectionInput = 0
  DirectionOutput = 1
  DirectionBoth = 2
End Enum

Public Enum LicenceType
  Concurrency = 0
  P14Headcount = 1
  Headcount = 2
  DMIConcurrencyAndP14 = 3
  DMIConcurrencyAndHeadcount = 4
End Enum

Public Enum WarningType
  Headcount95Percent = 0
  Licence5DayExpiry = 1
End Enum
