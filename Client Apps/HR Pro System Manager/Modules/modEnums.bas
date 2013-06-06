Attribute VB_Name = "modEnums"
Option Explicit

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
End Enum
    
'VB Message Box Icons constants
Public Enum picImageIcon
  giCritical = 0
  giInformation = 2 ^ 10
  giExclamation = 2 ^ 11
  giQuestion = 2 ^ 12
End Enum

'List/Tree view item type constants
Public Enum ViewItemTypes
  giNODE_TABLEGROUP = (2 ^ 0 Or edtAdd)
  giNODE_RELATIONGROUP = (2 ^ 1 Or edtAdd)
  giNODE_TABLE = (2 ^ 2 Or edtAdd Or edtEdit Or edtDelete Or edtCopy)
  giNODE_COLUMN = (2 ^ 3 Or edtAdd Or edtEdit Or edtDelete)
  giNODE_RELATION = (2 ^ 4 Or edtAdd Or edtEdit Or edtDelete)
  giNODE_RELATIONCHILD = (2 ^ 5 Or edtAdd Or edtEdit Or edtDelete)
End Enum

' Control type constants
Public Enum ControlTypes
  giCTRL_CHECKBOX = 2 ^ 0
  giCTRL_COMBOBOX = 2 ^ 1
  giCTRL_IMAGE = 2 ^ 2
  giCTRL_OLE = 2 ^ 3
  giCTRL_OPTIONGROUP = 2 ^ 4
  giCTRL_SPINNER = 2 ^ 5
  giCTRL_TEXTBOX = 2 ^ 6
  giCTRL_TAB = 2 ^ 7
  giCTRL_LABEL = 2 ^ 8
  giCTRL_FRAME = 2 ^ 9
  giCTRL_PHOTO = 2 ^ 10
  giCTRL_LINK = 2 ^ 11
  giCTRL_WORKINGPATTERN = 2 ^ 12
  giCTRL_LINE = 2 ^ 13
  giCTRL_NAVIGATION = 2 ^ 14
  giCTRL_COLOURPICKER = 2 ^ 15
End Enum

Public Enum PropertyConstants
  propCaption = 2 ^ 0
  propBackColor = 2 ^ 1
  propBorderStyle = 2 ^ 2
  propDisplayType = 2 ^ 3
  propFont = 2 ^ 4
  propForeColor = 2 ^ 5
  propPicture = 2 ^ 6
  propBold = 2 ^ 7
  propItalic = 2 ^ 8
  propUnderline = 2 ^ 9
  propStrikeThru = 2 ^ 10
  propLeft = 2 ^ 11
  propTop = 2 ^ 12
  propWidth = 2 ^ 13
  propHeight = 2 ^ 14
End Enum

Public Enum AccessModes
  accNone = 0
  accFull = 1
  accSupportMode = 2
  accLimited = 3
  accSystemReadOnly = 4
End Enum

Public Enum LockTypes
  lckNone = 0
  lckSaving = 1
  lckManual = 2
  lckReadWrite = 3
End Enum

' Expression Type constants
Public Enum ExpressionTypes
  giEXPR_UNKNOWNTYPE = 0
  giEXPR_COLUMNCALCULATION = 1
  giEXPR_GOTFOCUS = 2           ' Not used.
  giEXPR_RECORDVALIDATION = 3
  giEXPR_DEFAULTVALUE = 4       ' Not used.
  giEXPR_STATICFILTER = 5
  giEXPR_PAGEBREAK = 6          ' Not used.
  giEXPR_ORDER = 7              ' Not used.
  giEXPR_RECORDDESCRIPTION = 8
  giEXPR_VIEWFILTER = 9
  giEXPR_RUNTIMECALCULATION = 10
  giEXPR_RUNTIMEFILTER = 11
  giEXPR_EMAIL = 12
  giEXPR_LINKFILTER = 13
  giEXPR_UTILRUNTIMEFILTER = 14     'Data Manager Only
  giEXPR_MATCHJOINEXPRESSION = 15   'Data Manager Only
  giEXPR_MATCHSCOREEXPRESSION = 16  'Data Manager Only
  giEXPR_MATCHWHEREEXPRESSION = 17
  giEXPR_RECORDINDEPENDANTCALC = 18
  giEXPR_OUTLOOKFOLDER = 19
  giEXPR_OUTLOOKSUBJECT = 20
  giEXPR_WORKFLOWCALCULATION = 21
  giEXPR_WORKFLOWSTATICFILTER = 22
  giEXPR_WORKFLOWRUNTIMEFILTER = 23
End Enum

' Expression Value types
Public Enum ExpressionValueTypes
  giEXPRVALUE_UNDEFINED = 0
  giEXPRVALUE_CHARACTER = 1
  giEXPRVALUE_NUMERIC = 2
  giEXPRVALUE_LOGIC = 3
  giEXPRVALUE_DATE = 4
  giEXPRVALUE_TABLEVALUE = 5
  giEXPRVALUE_OLE = 6
  giEXPRVALUE_PHOTO = 7
  giEXPRVALUE_BYREF_UNDEFINED = 100
  giEXPRVALUE_BYREF_CHARACTER = 101
  giEXPRVALUE_BYREF_NUMERIC = 102
  giEXPRVALUE_BYREF_LOGIC = 103
  giEXPRVALUE_BYREF_DATE = 104
  giEXPRVALUE_BYREF_TABLEVALUE = 105 ' Not used.
  giEXPRVALUE_BYREF_OLE = 106 ' Not used.
  giEXPRVALUE_BYREF_PHOTO = 107 ' Not used.
End Enum

' Expression Value types
Public Enum ExpressionComponentTypes
  giCOMPONENT_FIELD = 1
  giCOMPONENT_FUNCTION = 2
  giCOMPONENT_CALCULATION = 3
  giCOMPONENT_VALUE = 4
  giCOMPONENT_OPERATOR = 5
  giCOMPONENT_TABLEVALUE = 6
  giCOMPONENT_PROMPTEDVALUE = 7
  giCOMPONENT_CUSTOMCALC = 8  ' Not used.
  giCOMPONENT_EXPRESSION = 9
  giCOMPONENT_FILTER = 10
  giCOMPONENT_WORKFLOWVALUE = 11
  giCOMPONENT_WORKFLOWFIELD = 12
End Enum

' Field pass types
Public Enum FieldPassTypes
  giPASSBY_VALUE = 1
  giPASSBY_REFERENCE = 2
End Enum

' Field selection types
Public Enum FieldSelectionTypes
  giSELECT_FIRSTRECORD = 1
  giSELECT_LASTRECORD = 2
  giSELECT_SPECIFICRECORD = 3
  giSELECT_RECORDTOTAL = 4
  giSELECT_RECORDCOUNT = 5
End Enum

' Order object constants.
Public Enum OrderTypes
  giORDERTYPE_STATIC = 0
  giORDERTYPE_DYNAMIC = 1
End Enum

'Table types
Public Enum TableTypes
  iTabView = 0
  iTabParent = 1
  iTabChild = 2
  iTabLookup = 3
End Enum

'Copy Permission types
Public Enum CopySecurityType
  giTABLEPARENT = 1
  giTABLECHILD = 2
  giTABLELOOKUP = 3
  giVIEW = 4
  giCOLUMN = 5
End Enum

' Calculation Trigger types
Public Enum CalculationTriggers
  iCalcTriggerDepFields = 0
  iCalcTriggerBeforeSave = 1
  iCalcTriggerAfterSave = 2
End Enum

'Time Period Types
Public Enum TimePeriods
  iTimePeriodDays = 0
  iTimePeriodWeeks = 1
  iTimePeriodMonths = 2
  iTimePeriodYears = 3
End Enum
                
' Expression Validation codes
Public Enum ExprValidationCodes
  giEXPRVALIDATION_NOERRORS = 0
  giEXPRVALIDATION_MISSINGOPERAND = 1
  giEXPRVALIDATION_SYNTAXERROR = 2
  giEXPRVALIDATION_EXPRTYPEMISMATCH = 3
  giEXPRVALIDATION_UNKNOWNERROR = 4
  giEXPRVALIDATION_OPERANDTYPEMISMATCH = 5
  giEXPRVALIDATION_PARAMETERTYPEMISMATCH = 6
  giEXPRVALIDATION_NOCOMPONENTS = 7
  giEXPRVALIDATION_PARAMETERSYNTAXERROR = 8
  giEXPRVALIDATION_PARAMETERNOCOMPONENTS = 9
  giEXPRVALIDATION_FILTEREVALUATION = 10
  giEXPRVALIDATION_SQLERROR = 11          ' JPD20020419 Fault 3687
  giEXPRVALIDATION_ASSOCSQLERROR = 12     ' JPD20020419 Fault 3687
  giEXPRVALIDATION_INVALIDRECORDIDENTIFICATION = 13
End Enum

Public Enum UndoActionFlags
  giACTION_NOACTION = 0
  giACTION_DROPTABPAGE = 1
  giACTION_DROPCONTROL = 2
  giACTION_CUTCONTROLS = 3
  giACTION_PASTECONTROLS = 4
  giACTION_DELETETABPAGE = 5
  giACTION_DELETECONTROLS = 6
  giACTION_MOVECONTROLS = 7
  giACTION_STRETCHCONTROLS = 8
  giACTION_AUTOFORMAT = 9
  giACTION_DROPCONTROLAUTOLABEL = 10
  giACTION_SWAPCONTROL = 11
End Enum

Public Enum AccessCodes
  giACCESS_READWRITE = 0
  giACCESS_READONLY = 1
  giACCESS_HIDDEN = 2
End Enum

' Column SELECT and UPDATE privilege constants
Public Enum ColumnPrivilegeStates
  giPRIVILEGES_NONEGRANTED = 0
  giPRIVILEGES_ALLGRANTED = 1
  giPRIVILEGES_SOMEGRANTED = 2
End Enum

' Used to output definition of table and column types
Public Enum OutputDefintionTypes
  giEXPORT_TO_PRINTER = 0
  giEXPORT_TO_CLIPBOARD = 1
  giEXPORT_TO_WORD = 2
End Enum

'Used to show which trim type has been applied
Public Enum TrimTypes
  giTRIM_NONE = 0
  giTRIM_BOTHSIDES = 1
  giTRIM_LEFTSIDE = 2
  giTRIM_RIGHTSIDE = 3
End Enum

' Mapping types used for Payroll Transfer
Public Enum AccordMapType
  MAPTYPE_COLUMN = 0
  MAPTYPE_EXPRESSION = 1
  MAPTYPE_VALUE = 2
End Enum

' Navigation Display Types
Public Enum NavigationDisplayType
  Hyperlink = 0
  Button = 1
  Browser = 2
  Hidden = 3
End Enum

' Navigation execution types
Public Enum NavigateIn
  URL = 0
  MenuBar = 1
  DB = 2
End Enum

Public Enum UsageButtonOptions
  USAGEBUTTONS_OK = 2 ^ 0
  USAGEBUTTONS_YES = 2 ^ 1
  USAGEBUTTONS_NO = 2 ^ 2
  USAGEBUTTONS_PRINT = 2 ^ 3
  USAGEBUTTONS_SELECT = 2 ^ 4
  USAGEBUTTONS_FIX = 2 ^ 5
  USAGEBUTTONS_COPY = 2 ^ 6
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

Public Enum ExpressionColour
  EXPRESSIONBUILDER_COLOUROFF = 1
  EXPRESSIONBUILDER_COLOURON = 2
End Enum

Public Enum ExpressionSaveView
  EXPRESSIONBUILDER_NODESMINIMIZE = 1
  EXPRESSIONBUILDER_NODESEXPAND = 2
  EXPRESSIONBUILDER_NODESTOPLEVEL = 3
End Enum

Public Enum AccordTransactionStatus
  ACCORD_STATUS_PENDING = 1
  ACCORD_STATUS_PENDING_CHANGED = 2
  ACCORD_STATUS_SUCCESS = 10
  ACCORD_STATUS_SUCCESS_WARNINGS = 11
  ACCORD_STATUS_FAILURE_UNKNOWN = 20
  ACCORD_STATUS_FAILURE_RESOLVED = 21
  ACCORD_STATUS_BLOCKED = 30
End Enum

Public Enum PasswordChangeReason
  giPasswordChange_None = 0
  giPasswordChange_MinLength = 1
  giPasswordChange_Expired = 2
  giPasswordChange_AdminRequested = 3
  giPasswordChange_LastChangeUnknown = 4
  giPasswordChange_ComplexitySettings = 5
End Enum

Public Enum HotfixType
  BEFORELOAD = 0
  BEFORESAVE = 1
  AFTERSAVE = 2
End Enum

Public Enum PrintFontFormat
  pffNormal = 0
  pffBold = 1
  pffNonBold = 2
End Enum

Public Enum UtilityType
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
End Enum

' Workflow Edit
Public Enum DataAction
  DATAACTION_INSERT = 0
  DATAACTION_UPDATE = 1
  DATAACTION_DELETE = 2
End Enum

' Numbering out of sequence as we started using the WorkflowWebFormItemTypes
' enum for this, which included differenty items. Sorry.
Public Enum WorkflowEmailItemTypes
  giWFEMAILITEM_UNKNOWN = -1
  giWFEMAILITEM_DBVALUE = 1
  giWFEMAILITEM_LABEL = 2
  giWFEMAILITEM_WFVALUE = 4
  giWFEMAILITEM_FORMATCODE = 12 ' NB. Only used in emails.
  giWFEMAILITEM_CALCULATION = 16
  giWFEMAILITEM_FILEATTACHMENT = 18 ' NB. Only used in emails.
End Enum

Public Enum WFTriggerRelatedRecord
  WFRELATEDRECORD_INSERT = 0
  WFRELATEDRECORD_UPDATE = 1
  WFRELATEDRECORD_DELETE = 2
End Enum

Public Enum WorkflowTriggerLinkType
  WORKFLOWTRIGGERLINKTYPE_COLUMN = 0
  WORKFLOWTRIGGERLINKTYPE_RECORD = 1
  WORKFLOWTRIGGERLINKTYPE_DATE = 2
End Enum

Public Enum WorkflowTriggerOffsetPeriod
  WORKFLOWTRIGGERLINKOFFESTPERIOD_DAY = 0
  WORKFLOWTRIGGERLINKOFFESTPERIOD_WEEK = 1
  WORKFLOWTRIGGERLINKOFFESTPERIOD_MONTH = 2
  WORKFLOWTRIGGERLINKOFFESTPERIOD_YEAR = 3
End Enum

Public Enum ProcessAdminConfig
  iPROCESSADMIN_DISABLED = 0
  iPROCESSADMIN_SERVICEACCOUNT = 1
  iPROCESSADMIN_SQLACCOUNT = 2
  iPROCESSADMIN_EVERYONE = 3
End Enum


