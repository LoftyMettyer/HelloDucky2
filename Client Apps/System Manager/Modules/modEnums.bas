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

Public Enum CrossTabType
    cttNormal = 0
    cttTurnover = 1
    cttStability = 2
    cttAbsenceBreakdown = 3
    ctt9GridBox = 4
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
  lckSaveRequest = 4
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
Public Enum enum_TableTypes
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

Public Enum FusionTransactionStatus
  FUSION_STATUS_PENDING = 1
  FUSION_STATUS_PENDING_CHANGED = 2
  FUSION_STATUS_SUCCESS = 10
  FUSION_STATUS_SUCCESS_WARNINGS = 11
  FUSION_STATUS_FAILURE_UNKNOWN = 20
  FUSION_STATUS_FAILURE_RESOLVED = 21
  FUSION_STATUS_BLOCKED = 30
End Enum

'Mapping types used for Fusion Transfer
Public Enum FusionMapType
  FUSION_MAPTYPE_COLUMN = 0
  FUSION_MAPTYPE_EXPRESSION = 1
  FUSION_MAPTYPE_VALUE = 2
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
  utlCalendarreport = 17
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

Public Enum WorkflowElementProperties
  WORKFLOWELEMENTPROP_ITEMVALUE = 0
  WORKFLOWELEMENTPROP_COMPETIONCOUNT = 1
  WORKFLOWELEMENTPROP_FAILURECOUNT = 2
  WORKFLOWELEMENTPROP_TIMEOUTCOUNT = 3
  WORKFLOWELEMENTPROP_MESSAGE = 4
End Enum

Public Enum WorkflowWebFormItemTypes
  giWFFORMITEM_FORM = -2
  giWFFORMITEM_UNKNOWN = -1
  giWFFORMITEM_BUTTON = 0
  giWFFORMITEM_DBVALUE = 1
  giWFFORMITEM_LABEL = 2
  giWFFORMITEM_INPUTVALUE_CHAR = 3
  giWFFORMITEM_WFVALUE = 4
  giWFFORMITEM_INPUTVALUE_NUMERIC = 5
  giWFFORMITEM_INPUTVALUE_LOGIC = 6
  giWFFORMITEM_INPUTVALUE_DATE = 7
  giWFFORMITEM_FRAME = 8
  giWFFORMITEM_LINE = 9
  giWFFORMITEM_IMAGE = 10
  giWFFORMITEM_INPUTVALUE_GRID = 11
  giWFFORMITEM_FORMATCODE = 12 ' NB. Only used in emails.
  giWFFORMITEM_INPUTVALUE_DROPDOWN = 13
  giWFFORMITEM_INPUTVALUE_LOOKUP = 14
  giWFFORMITEM_INPUTVALUE_OPTIONGROUP = 15
  giWFFORMITEM_CALC = 16
  giWFFORMITEM_INPUTVALUE_FILEUPLOAD = 17
  giWFFORMITEM_FILEATTACHMENT = 18
  giWFFORMITEM_DBFILE = 19
  giWFFORMITEM_WFFILE = 20
  giWFFORMITEM_PAGETAB = 21
End Enum

Public Enum WorkflowInstanceStatus
  giWFSTATUS_UNKNOWN = -1
  giWFSTATUS_INPROGRESS = 0
  giWFSTATUS_CANCELLED = 1 ' No longer used
  giWFSTATUS_ERROR = 2
  giWFSTATUS_COMPLETE = 3
  giWFSTATUS_SCHEDULED = 4
End Enum

Public Enum WorkflowStoredDataValueTypes
  giWFDATAVALUE_UNKNOWN = -1
  giWFDATAVALUE_FIXED = 0
  giWFDATAVALUE_WFVALUE = 1
  giWFDATAVALUE_DBVALUE = 2
  giWFDATAVALUE_CALC = 3
End Enum

Public Enum WorkflowRecordSelectorTypes
  giWFRECSEL_UNKNOWN = -1         ' Oops, something's wrong
  giWFRECSEL_INITIATOR = 0        ' Initiator's personnel table record
  giWFRECSEL_IDENTIFIEDRECORD = 1 ' Identified via WebForm RecordSelector or StoredData Inserted/Updated Record
  giWFRECSEL_ALL = 2              ' Show all records from the table in a WebForm RecordSelector
  giWFRECSEL_UNIDENTIFIED = 3     ' Used when StoredData Inserts into a top-level table
  giWFRECSEL_TRIGGEREDRECORD = 4  ' Triggered Base table record
End Enum

Public Enum WorkflowFindUsageOption
  wfNone = 0
  wfRecSelType = 1
  wfElement = 2
  wfWebFormItem = 3
End Enum

Public Enum WorkflowInitiationTypes
  WORKFLOWINITIATIONTYPE_MANUAL = 0
  WORKFLOWINITIATIONTYPE_TRIGGERED = 1
  WORKFLOWINITIATIONTYPE_EXTERNAL = 2
End Enum

Public Enum WorkflowWebFormValidationTypes
  WORKFLOWWFVALIDATIONTYPE_ERROR = 0
  WORKFLOWWFVALIDATIONTYPE_WARNING = 1
End Enum

Public Enum DecisionCaptionType
  decisionCaption_T_F = 0
  decisionCaption_Y_N = 1
  decisionCaption_1_0 = 2
  decisionCaption_tick_cross = 3
End Enum

Public Enum WFItemPropertyOrientation
  wfItemPropertyOrientation_Vertical = 0
  wfItemPropertyOrientation_Horizontal = 1
End Enum

Public Enum WFItemPropertyState
  wfItemPropertyState_No = 0
  wfItemPropertyState_ReadOnly = 1
  wfItemPropertyState_ReadWrite = 2
End Enum

Public Enum WorkflowButtonAction
  WORKFLOWBUTTONACTION_SUBMIT = 0
  WORKFLOWBUTTONACTION_SAVEFORLATER = 1
  WORKFLOWBUTTONACTION_CANCEL = 2
End Enum

Public Enum WorkflowWebFormMessageType
  WORKFLOWWEBFORMMESSAGE_COMPLETION = 0
  WORKFLOWWEBFORMMESSAGE_SAVEDFORLATER = 1
  WORKFLOWWEBFORMMESSAGE_FOLLOWONFORMS = 2
End Enum

' Property type constants.
Public Enum WFItemProperty
  WFITEMPROP_NONE = -1 ' Used for the category markers in the property grid
  WFITEMPROP_UNKNOWN = 0
  WFITEMPROP_ALIGNMENT = 1
  WFITEMPROP_BACKCOLOR = 2
  WFITEMPROP_BORDERSTYLE = 3
  WFITEMPROP_CAPTION = 4
  WFITEMPROP_FONT = 5
  WFITEMPROP_FORECOLOR = 6
  WFITEMPROP_HEIGHT = 7
  WFITEMPROP_LEFT = 8
  WFITEMPROP_PICTURE = 9
  WFITEMPROP_TOP = 10
  WFITEMPROP_WIDTH = 11
  WFITEMPROP_WFIDENTIFIER = 12  ' Identifier of the Input/RecSel control in this WebForm
  WFITEMPROP_PICTURELOCATION = 13
  WFITEMPROP_DEFAULTVALUE_CHAR = 14
  WFITEMPROP_DBRECORD = 15 ' RecordSelection type of a DBValue control
  WFITEMPROP_SIZE = 16
  WFITEMPROP_DECIMALS = 17
  WFITEMPROP_DEFAULTVALUE_DATE = 18
  WFITEMPROP_DEFAULTVALUE_LOGIC = 19
  WFITEMPROP_DEFAULTVALUE_NUMERIC = 20
  WFITEMPROP_BACKSTYLE = 21
  WFITEMPROP_BACKCOLOREVEN = 22
  WFITEMPROP_BACKCOLORODD = 23
  WFITEMPROP_COLUMNHEADERS = 24
  WFITEMPROP_FORECOLOREVEN = 25
  WFITEMPROP_FORECOLORODD = 26
  WFITEMPROP_HEADERBACKCOLOR = 27
  WFITEMPROP_HEADFONT = 28
  WFITEMPROP_HEADLINES = 29
  WFITEMPROP_TABLEID = 30
  WFITEMPROP_ELEMENTIDENTIFIER = 31 ' Identifier of a preceding record identifying element (StoredData or WebForm with RecSel)
  WFITEMPROP_RECORDSELECTOR = 32  ' Identifier of a RecSel control in the WebForm identified by WFITEMPROP_ELEMENTIDENTIFIER
  WFITEMPROP_RECSELTYPE = 33 ' RecordSelection type of a RecSel control
  WFITEMPROP_BACKCOLORHIGHLIGHT = 34
  WFITEMPROP_FORECOLORHIGHLIGHT = 35
  WFITEMPROP_TIMEOUT = 36
  WFITEMPROP_CONTROLVALUELIST = 37
  WFITEMPROP_DEFAULTVALUE_LIST = 38
  WFITEMPROP_LOOKUPTABLEID = 39
  WFITEMPROP_LOOKUPCOLUMNID = 40
  WFITEMPROP_DEFAULTVALUE_LOOKUP = 41
  WFITEMPROP_RECORDTABLEID = 42 ' Table identified by WFITEMPROP_ELEMENTIDENTIFIER/WFITEMPROP_RECORDSELECTOR (can be ascendant table of the one in the element/recsel)
  WFITEMPROP_DESCRIPTION = 43
  WFITEMPROP_ORIENTATION = 44
  WFITEMPROP_RECORDORDER = 45
  WFITEMPROP_RECORDFILTER = 46
  WFITEMPROP_VALIDATION = 47
  WFITEMPROP_MANDATORY = 48
  WFITEMPROP_DEFAULTVALUE_EXPRID = 49
  WFITEMPROP_DEFAULTVALUE_WORKPATTERN = 50
  WFITEMPROP_DESCRIPTION_WORKFLOWNAME = 51
  WFITEMPROP_DESCRIPTION_ELEMENTCAPTION = 52
  WFITEMPROP_SUBMITTYPE = 53
  WFITEMPROP_CALCULATION = 54
  WFITEMPROP_CAPTIONTYPE = 55
  WFITEMPROP_DEFAULTVALUETYPE = 56
  WFITEMPROP_VERTICALOFFSETBEHAVIOUR = 57
  WFITEMPROP_HORIZONTALOFFSETBEHAVIOUR = 58
  WFITEMPROP_VERTICALOFFSET = 59
  WFITEMPROP_HORIZONTALOFFSET = 60
  WFITEMPROP_HEIGHTBEHAVIOUR = 61
  WFITEMPROP_WIDTHBEHAVIOUR = 62
  WFITEMPROP_PASSWORDTYPE = 63
  WFITEMPROP_COMPLETIONMESSAGETYPE = 64
  WFITEMPROP_COMPLETIONMESSAGE = 65
  WFITEMPROP_SAVEDFORLATERMESSAGETYPE = 66
  WFITEMPROP_SAVEDFORLATERMESSAGE = 67
  WFITEMPROP_FOLLOWONFORMSMESSAGETYPE = 68
  WFITEMPROP_FOLLOWONFORMSMESSAGE = 69
  WFITEMPROP_FILEEXTENSIONS = 70
  WFITEMPROP_LOOKUPFILTER = 71
  WFITEMPROP_LOOKUPFILTERCOLUMN = 72
  WFITEMPROP_LOOKUPFILTEROPERATOR = 73
  WFITEMPROP_LOOKUPFILTERVALUE = 74
  WFITEMPROP_TABNUMBER = 75
  WFITEMPROP_TABCAPTION = 76
  WFITEMPROP_LOOKUPORDER = 77
  WFITEMPROP_HOTSPOT = 78
  WFITEMPROP_USEASTARGETIDENTIFIER = 79
End Enum

Public Enum ElementType
  elem_Begin = 0
  elem_Terminator = 1
  elem_WebForm = 2
  elem_Email = 3
  elem_Decision = 4
  elem_StoredData = 5
  elem_SummingJunction = 6
  elem_Or = 7
  elem_Connector1 = 8
  elem_Connector2 = 9
End Enum

Public Enum LineDirection
  lineDirection_Down = 0
  lineDirection_Left = 1
  lineDirection_Right = 2
  lineDirection_Up = 3
End Enum

Public Enum VerticalOffset
  offsetTop = 0
  offsetBottom = 1
End Enum

Public Enum HorizontalOffset
  offsetLeft = 0
  offsetRight = 1
End Enum

Public Enum ControlSizeBehaviour
  behaveFixed = 0
  behaveFull = 1
End Enum

Public Enum DecisionOutboundFlows
  decisionOutFlow_False = 0
  decisionOutFlow_True = 1
End Enum

Public Enum WebFormOutboundFlows
  webFormOutFlow_Normal = 0
  webFormOutFlow_Timeout = 1
End Enum

Public Enum StoredDataOutboundFlows
  storedDataOutFlow_Success = 0
  storedDataOutFlow_Failure = 1
End Enum

Public Enum ProcessAdminConfig
  iPROCESSADMIN_DISABLED = 0
  iPROCESSADMIN_SERVICEACCOUNT = 1
  iPROCESSADMIN_SQLACCOUNT = 2
  iPROCESSADMIN_EVERYONE = 3
End Enum

Public Enum enum_EmailType
  LinkRecord = 0
  LinkColumn = 1
  LinkOffset = 2
  LinkAmendment = 3
  LinkRebuild = 4
End Enum

Public Enum enum_Module
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
  modNineBoxGrid = 65536
  modEditableGrids = 131072
End Enum

'System metrics constants
Public Enum SystemMetrics
  SM_CYCAPTION = 4
  SM_CXBORDER = 5
  SM_CYBORDER = 6
  SM_CXFRAME = 32
  SM_CYFRAME = 33
  SM_CXICON = 11
  SM_CXICONSPACING = 38
  SM_CYICON = 12
  SM_CYICONSPACING = 39
  SM_CYSMCAPTION = 51
  SM_CXVSCROLL = 2
End Enum

' GetWindow constants.
Public Enum GetWindowConstants
  GW_HWNDFIRST = 0
  GW_HWNDLAST = 1
  GW_HWNDNEXT = 2
  GW_HWNDPREV = 3
  GW_OWNER = 4
  GW_CHILD = 5
  GW_MAX = 5
End Enum

Public Enum DataTypes
  dtBIT = -7
  dtLONGVARBINARY = -4
  dtVARBINARY = -3
  dtBINARY = -2
  dtLONGVARCHAR = -1
  dtNUMERIC = 2
  dtINTEGER = 4
  dtTIMESTAMP = 11
  dtVARCHAR = 12
  dtGUID = 15
  ' These are temp for frmworkflowelementitem - TM has the file checked out and JED needs a build.
  ' To be removed when frmworkflowelementitem is available
  rdTypeLONGVARBINARY = -4
  rdTypeVARBINARY = -3
End Enum

Public Enum SSINTRANETLINKTYPES
  SSINTLINK_HYPERTEXT = 0
  SSINTLINK_BUTTON = 1
  SSINTLINK_DROPDOWNLIST = 2
  SSINTLINK_DOCUMENT = 3
End Enum

Public Enum SSINTRANETSCREENTYPES
  SSINTLINKSCREEN_HRPRO = 0
  SSINTLINKSCREEN_URL = 1
  SSINTLINKSCREEN_UTILITY = 2
  SSINTLINKSCREEN_EMAIL = 3
  SSINTLINKSCREEN_APPLICATION = 4
  SSINTLINKSEPARATOR = 5
  SSINTLINKCHART = 6
  SSINTLINKDB_VALUE = 7
  SSINTLINKPWFSTEPS = 8
  SSINTLINKTODAYS_EVENTS = 9
  SSINTLINKORGCHART = 10
  SSINTLINKSCREEN_DOCUMENT = 11
End Enum

'Column type constants
Public Enum ColumnTypes
  giCOLUMNTYPE_DATA = 0
  giCOLUMNTYPE_LOOKUP = 1
  giCOLUMNTYPE_CALCULATED = 2
  giCOLUMNTYPE_SYSTEM = 3
  giCOLUMNTYPE_LINK = 4
End Enum

Public Enum UniqueCheckTypes
  giUNIQUECHECKTYPE_NONE = 0
  giUNIQUECHECKTYPE_ENTIRE = -1
  giUNIQUECHECKTYPE_SIBLINGSALL = -2
End Enum

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


Public Enum UsageCheckObject
  Table = 1
  Relationship = 2
  Column = 3
  Calculation = 4
  Order = 5
  Picture = 6
  Filter = 7
  ChildTable = 8
  LookupTable = 9
  Email = 10
  OutlookFolder = 11
  View = 12
  Workflow = 13
  Form = 14
End Enum

Public Enum SQLServerAuthenticationType
  iWINDOWSONLY = 1
  iMIXEDMODE = 2
End Enum

Public Enum UsageCheckType
  Delete = 1
  Edit = 2
  Copy = 3
End Enum

Public Enum OLEType
  OLE_LOCAL = 0
  OLE_SERVER = 1
  OLE_EMBEDDED = 2
End Enum

Public Enum enum_ValidationType
  VALIDATION_MANDATORY = 0
  VALIDATION_UNIQUE = 1
  VALIDATION_DUPLICATE = 2
  VALIDATION_OVERLAP = 3
  VALIDATION_CUSTOM = 4
End Enum

Public Enum enum_Severity
  Severity_Warning = 0
  SEVERITY_FAILURE = 1
  SEVERITY_INFORMATION = 2
End Enum

Public Enum MobileElementItemTypes
  giMOBILEITEM_UNKNOWN = -1
  giMOBILEITEM_BANNER_BG = 0
  giMOBILEITEM_BANNER_RIGHTLOGO = 1
  giMOBILEITEM_BODY_BG = 2
End Enum

' Property type constants.
Public Enum MOBITEMProperty
  MOBITEMPROP_NONE = -1 ' Used for the category markers in the property grid
  MOBITEMPROP_UNKNOWN = 0
  MOBITEMPROP_ALIGNMENT = 1
  MOBITEMPROP_BACKCOLOR = 2
  MOBITEMPROP_BORDERSTYLE = 3
  MOBITEMPROP_CAPTION = 4
  MOBITEMPROP_FONT = 5
  MOBITEMPROP_FORECOLOR = 6
  MOBITEMPROP_HEIGHT = 7
  MOBITEMPROP_LEFT = 8
  MOBITEMPROP_PICTURE = 9
  MOBITEMPROP_TOP = 10
  MOBITEMPROP_WIDTH = 11
  MOBITEMPROP_WFIDENTIFIER = 12  ' Identifier of the Input/RecSel control in this WebForm
  MOBITEMPROP_PICTURELOCATION = 13
  MOBITEMPROP_DEFAULTVALUE_CHAR = 14
  MOBITEMPROP_DBRECORD = 15 ' RecordSelection type of a DBValue control
  MOBITEMPROP_SIZE = 16
  MOBITEMPROP_DECIMALS = 17
  MOBITEMPROP_DEFAULTVALUE_DATE = 18
  MOBITEMPROP_DEFAULTVALUE_LOGIC = 19
  MOBITEMPROP_DEFAULTVALUE_NUMERIC = 20
  MOBITEMPROP_BACKSTYLE = 21
  MOBITEMPROP_BACKCOLOREVEN = 22
  MOBITEMPROP_BACKCOLORODD = 23
  MOBITEMPROP_COLUMNHEADERS = 24
  MOBITEMPROP_FORECOLOREVEN = 25
  MOBITEMPROP_FORECOLORODD = 26
  MOBITEMPROP_HEADERBACKCOLOR = 27
  MOBITEMPROP_HEADFONT = 28
  MOBITEMPROP_HEADLINES = 29
  MOBITEMPROP_TABLEID = 30
  MOBITEMPROP_ELEMENTIDENTIFIER = 31 ' Identifier of a preceding record identifying element (StoredData or WebForm with RecSel)
  MOBITEMPROP_RECORDSELECTOR = 32  ' Identifier of a RecSel control in the WebForm identified by MOBITEMPROP_ELEMENTIDENTIFIER
  MOBITEMPROP_RECSELTYPE = 33 ' RecordSelection type of a RecSel control
  MOBITEMPROP_BACKCOLORHIGHLIGHT = 34
  MOBITEMPROP_FORECOLORHIGHLIGHT = 35
  MOBITEMPROP_TIMEOUT = 36
  MOBITEMPROP_CONTROLVALUELIST = 37
  MOBITEMPROP_DEFAULTVALUE_LIST = 38
  MOBITEMPROP_LOOKUPTABLEID = 39
  MOBITEMPROP_LOOKUPCOLUMNID = 40
  MOBITEMPROP_DEFAULTVALUE_LOOKUP = 41
  MOBITEMPROP_RECORDTABLEID = 42 ' Table identified by MOBITEMPROP_ELEMENTIDENTIFIER/MOBITEMPROP_RECORDSELECTOR (can be ascendant table of the one in the element/recsel)
  MOBITEMPROP_DESCRIPTION = 43
  MOBITEMPROP_ORIENTATION = 44
  MOBITEMPROP_RECORDORDER = 45
  MOBITEMPROP_RECORDFILTER = 46
  MOBITEMPROP_VALIDATION = 47
  MOBITEMPROP_MANDATORY = 48
  MOBITEMPROP_DEFAULTVALUE_EXPRID = 49
  MOBITEMPROP_DEFAULTVALUE_WORKPATTERN = 50
  MOBITEMPROP_DESCRIPTION_WORKFLOWNAME = 51
  MOBITEMPROP_DESCRIPTION_ELEMENTCAPTION = 52
  MOBITEMPROP_SUBMITTYPE = 53
  MOBITEMPROP_CALCULATION = 54
  MOBITEMPROP_CAPTIONTYPE = 55
  MOBITEMPROP_DEFAULTVALUETYPE = 56
  MOBITEMPROP_VERTICALOFFSETBEHAVIOUR = 57
  MOBITEMPROP_HORIZONTALOFFSETBEHAVIOUR = 58
  MOBITEMPROP_VERTICALOFFSET = 59
  MOBITEMPROP_HORIZONTALOFFSET = 60
  MOBITEMPROP_HEIGHTBEHAVIOUR = 61
  MOBITEMPROP_WIDTHBEHAVIOUR = 62
  MOBITEMPROP_PASSWORDTYPE = 63
  MOBITEMPROP_COMPLETIONMESSAGETYPE = 64
  MOBITEMPROP_COMPLETIONMESSAGE = 65
  MOBITEMPROP_SAVEDFORLATERMESSAGETYPE = 66
  MOBITEMPROP_SAVEDFORLATERMESSAGE = 67
  MOBITEMPROP_FOLLOWONFORMSMESSAGETYPE = 68
  MOBITEMPROP_FOLLOWONFORMSMESSAGE = 69
  MOBITEMPROP_FILEEXTENSIONS = 70
  MOBITEMPROP_LOOKUPFILTER = 71
  MOBITEMPROP_LOOKUPFILTERCOLUMN = 72
  MOBITEMPROP_LOOKUPFILTEROPERATOR = 73
  MOBITEMPROP_LOOKUPFILTERVALUE = 74
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

Public Enum TriggerCodePosition
  AfterU02Update = 0
  AfterI02Insert = 1
  AfterD01Delete = 2
End Enum
