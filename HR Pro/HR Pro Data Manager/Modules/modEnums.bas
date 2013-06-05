Attribute VB_Name = "modEnums"
' ------------
' clsApplication
' ------------

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

'DefSel Screen Enum
Public Enum DefSelScreen
    screenNone = 0
    screenDataTransfer = 1
    screenGlobalAdd
    screenGlobalUpdate
    screenGlobalDelete
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
  fmtExcelChart = 5
  fmtExcelPivotTable = 6
  fmtFixedLengthFile = 7
  fmtCMGFile = 8
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
