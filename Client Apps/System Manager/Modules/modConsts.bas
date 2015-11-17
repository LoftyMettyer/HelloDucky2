Attribute VB_Name = "modConsts"
Option Explicit

Public Const VARCHAR_MAX_Size = 2147483646 'Yup one below the actual max, needs to be otherwise things go so awfully wrong, you don't believe me, well go on then, change it, see if I care!!!)

' Window formatting
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const WS_THICKFRAME As Long = &H40000
Public Const WS_MAXIMIZE As Long = &H1000000
Public Const WS_MAXIMIZEBOX As Long = &H10000
Public Const WS_MINIMIZE As Long = &H20000000
Public Const WS_MINIMIZEBOX As Long = &H20000
Public Const WS_EX_WINDOWEDGE As Long = &H100
Public Const WS_EX_APPWINDOW As Long = &H40000
Public Const WS_EX_DLGMODALFRAME As Long = &H1
Public Const GWL_EXSTYLE As Long = (-20)
Public Const GWL_STYLE As Long = (-16)

' Advanced database settings to control the recursion levels in the database
Public Const giDefaultRecursionLevel = 8

Public Const gsDUMMY_CHARACTER = "ASRDUMMYCHARVALUE"
Public Const gsDUMMY_NUMERIC = 1
Public Const gsDUMMY_LOGIC = True
Public Const gsDUMMY_DATE = #1/1/1998#
Public Const gsDUMMY_BYREF_CHARACTER = dtVARCHAR & vbTab & "a"
Public Const gsDUMMY_BYREF_NUMERIC = dtNUMERIC & vbTab & "1"
Public Const gsDUMMY_BYREF_LOGIC = dtBIT & vbTab & "0"
Public Const gsDUMMY_BYREF_DATE = dtTIMESTAMP & vbTab & "1/1/1998"

Public Const giPRINT_XINDENT = 1000
Public Const giPRINT_YINDENT = 1000
Public Const giPRINT_XSPACE = 500
Public Const giPRINT_YSPACE = 100

Public Const gsOLEDISPLAYTYPE_ICON = " (icon)"
Public Const gsOLEDISPLAYTYPE_CONTENTS = " (contents)"

Public Const gsVALIDATIONSPPREFIX = "udfvalid_"
Public Const gsVIEWVALIDATIONSPPREFIX = "sp_ASRValidateView_"

Public Const gsEMAILADDR = "dbo.spASRSysEmailAddr_"
Public Const gsEMAILSEND = "dbo.spASRSysEmailSend_"

Public Const gstrWindowsAuthentication_DNSName = "HRPRO"

' WORKFLOW MODULE CONSTANTS
Public Const gsMODULEKEY_WORKFLOW = "MODULE_WORKFLOW"
' Parameter Type constants.
Public Const gsPARAMETERKEY_URL = "Param_URL"
Public Const gsPARAMETERKEY_WEBPARAM1 = "Param_Web1"
Public Const gsPARAMETERKEY_EMAILCOLUMN = "Param_EmailColumn"
Public Const gsPARAMETERKEY_DELEGATIONACTIVATEDCOLUMN = "Param_DelegationActivatedColumn"
Public Const gsPARAMETERKEY_DELEGATEEMAIL = "Param_DelegateEmail"
Public Const gsPARAMETERKEY_COPYDELEGATEEMAIL = "Param_CopyDelegateEmail"
Public Const gsPARAMETERKEY_LOGINDETAILS = "Param_FieldsLoginDetails"
Public Const gsPARAMETERKEY_REQUIRESAUTHORIZATION = "Param_FieldsAuthorization"

Public Const gsWORKFLOWAPPLICATIONPREFIX = "OpenHR Workflow"

' MOBILE MODULE CONSTANTS
Public Const gsMODULEKEY_MOBILE = "MODULE_MOBILE"
' Parameter Type constants.
' fault 2303:
' Public Const gsPARAMETERKEY_UNIQUEEMAILCOLUMN = "Param_UniqueEmailColumn"
Public Const gsPARAMETERKEY_MOBILELOGIN = "Param_LoginName"
' Public Const gsPARAMETERKEY_LEAVINGDATE = "Param_FieldsLeavingDate"
Public Const gsPARAMETERKEY_MOBILEACTIVATED = "Param_MobileActivated"

Public Const gsMOBILEAPPLICATIONPREFIX = "OpenHR MOBILE"

'CATEGORY MODULE CONSTANTS
Public Const gsMODULEKEY_CATEGORY = "MODULE_CATEGORY"
Public Const gsPARAMETERKEY_CATEGORYTABLE = "Param_CategoryTable"
Public Const gsPARAMETERKEY_CATEGORYNAMECOLUMN = "Param_CatageoryNameColumn"

' TRAINING BOOKING MODULE CONSTANTS
Public Const gsMODULEKEY_TRAININGBOOKING = "MODULE_TRAININGBOOKING"
Public Const gsMODULEKEY_PERSONNEL = "MODULE_PERSONNEL"

' Parameter Type constants.
Public Const gsPARAMETERTYPE_TABLEID = "PType_TableID"
Public Const gsPARAMETERTYPE_COLUMNID = "PType_ColumnID"
Public Const gsPARAMETERTYPE_ORDERID = "PType_OrderID"
Public Const gsPARAMETERTYPE_OTHER = "PType_Other"
Public Const gsPARAMETERTYPE_SCREENID = "PType_ScreenID"
Public Const gsPARAMETERTYPE_VIEWID = "PType_ViewID"
Public Const gsPARAMETERTYPE_ENCYPTED = "PType_Encrypted"
Public Const gsPARAMETERTYPE_EMAILID = "PType_EmailID"

' Course Records constants.
Public Const gsPARAMETERKEY_COURSETABLE = "Param_CourseTable"
Public Const gsPARAMETERKEY_COURSETITLE = "Param_CourseTitle"
Public Const gsPARAMETERKEY_COURSESTARTDATE = "Param_CourseStartDate"
Public Const gsPARAMETERKEY_COURSEENDDATE = "Param_CourseEndDate"
Public Const gsPARAMETERKEY_COURSENUMBERBOOKED = "Param_CourseNumberBooked"
Public Const gsPARAMETERKEY_COURSEMAXNUMBER = "Param_CourseMaxNumber"
Public Const gsPARAMETERKEY_COURSECANCELLATIONDATE = "Param_CourseCancelDate"
Public Const gsPARAMETERKEY_COURSECANCELLEDBY = "Param_CourseCancelledBy"
Public Const gsPARAMETERKEY_COURSETRANSFERPROVISIONALS = "Param_CourseTransferProvisionals"
Public Const gsPARAMETERKEY_COURSEINCLUDEPROVISIONALS = "Param_CourseIncludeProvisionals"
Public Const gsPARAMETERKEY_COURSEOVERBOOKINGNOTIFICATION = "Param_CourseOverbookingNotification"
Public Const gsPARAMETERKEY_COURSEORDER = "Param_CourseOrder"
' Pre-requisite constants.
Public Const gsPARAMETERKEY_PREREQTABLE = "Param_PreReqTable"
Public Const gsPARAMETERKEY_PREREQCOURSETITLE = "Param_PreReqCourseTitle"
Public Const gsPARAMETERKEY_PREREQGROUPING = "Param_PreReqGrouping"
Public Const gsPARAMETERKEY_PREREQFAILURE = "Param_PreReqFailure"
Public Const gsPARAMETERKEY_PREREQDFLTFAILURE = "Param_PreReqDfltFailure"
' Employee Records constants.
Public Const gsPARAMETERKEY_EMPLOYEETABLE = "Param_EmployeeTable"
Public Const gsPARAMETERKEY_EMPLOYEEORDER = "Param_EmployeeOrder"
Public Const gsPARAMETERKEY_BULKBOOKINGDEFAULTVIEW = "Param_BulkBookingDefaultView" 'NHRD01052003 Fault 4687
' Unavailability constants.
Public Const gsPARAMETERKEY_UNAVAILTABLE = "Param_UnavailTable"
Public Const gsPARAMETERKEY_UNAVAILFROMDATE = "Param_UnavailFromDate"
Public Const gsPARAMETERKEY_UNAVAILTODATE = "Param_UnavailToDate"
Public Const gsPARAMETERKEY_UNAVAILFAILURE = "Param_UnavailFailure"
Public Const gsPARAMETERKEY_UNAVAILDFLTFAILURE = "Param_UnavailDfltFailure"
' Waiting List constants.
Public Const gsPARAMETERKEY_WAITLISTTABLE = "Param_WaitListTable"
Public Const gsPARAMETERKEY_WAITLISTCOURSETITLE = "Param_WaitListCourseTitle"
Public Const gsPARAMETERKEY_WAITLISTOVERRIDECOLUMN = "Param_WaitListOverRideColumn"
' Training Booking constants.
Public Const gsPARAMETERKEY_TRAINBOOKTABLE = "Param_TrainBookTable"
Public Const gsPARAMETERKEY_TRAINBOOKCOURSETITLE = "Param_TrainBookCourseTitle"
Public Const gsPARAMETERKEY_TRAINBOOKSTATUS = "Param_TrainBookStatus"
Public Const gsPARAMETERKEY_TRAINBOOKCANCELDATE = "Param_TrainBookCancelDate"
Public Const gsPARAMETERKEY_TRAINBOOKOVERLAPNOTIFICATION = "Param_TrainBookOverlapNotification"
' Related Column constants.
Public Const gsPARAMETERKEY_TBWLRELATEDCOLUMNS = "Param_TBWLRelatedColumns"

' PERSONNEL MODULE CONSTANTS

Public Const gsPARAMETERKEY_PERSONNELTABLE = "Param_TablePersonnel"
Public Const gsPARAMETERKEY_EMPLOYEENUMBER = "Param_FieldsEmployeeNumber"
Public Const gsPARAMETERKEY_FORENAME = "Param_FieldsForename"
Public Const gsPARAMETERKEY_SURNAME = "Param_FieldsSurname"
Public Const gsPARAMETERKEY_SSIWELCOME = "Param_FieldsSSIWelcome"
Public Const gsPARAMETERKEY_SSIPHOTOGRAPH = "Param_FieldsSSIPhotograph"
Public Const gsPARAMETERKEY_STARTDATE = "Param_FieldsStartDate"
Public Const gsPARAMETERKEY_LEAVINGDATE = "Param_FieldsLeavingDate"
Public Const gsPARAMETERKEY_FULLPARTTIME = "Param_FieldsFullPartTime"
Public Const gsPARAMETERKEY_EMAIL = "Param_FieldsEmail"
Public Const gsPARAMETERKEY_WORKEMAIL = "Param_FieldsWorkEmail"
Public Const gsPARAMETERKEY_DEPARTMENT = "Param_FieldsDepartment"
Public Const gsPARAMETERKEY_DATEOFBIRTH = "Param_FieldsDateOfBirth"
Public Const gsPARAMETERKEY_LOGINNAME = "Param_FieldsLoginName"
Public Const gsPARAMETERKEY_SECONDLOGINNAME = "Param_FieldsSecondLoginName"
Public Const gsPARAMETERKEY_GRADE = "Param_FieldsGrade"
Public Const gsPARAMETERKEY_MANAGERSTAFFNO = "Param_FieldsManagerStaffNo"
Public Const gsPARAMETERKEY_JOBTITLE = "Param_FieldsJobTitle"

'Region Constants - The following key is used for static region field
Public Const gsPARAMETERKEY_REGION = "Param_FieldsRegion"
'Region Constants - The following keys are used for historical region fields
Public Const gsPARAMETERKEY_HREGIONTABLE = "Param_FieldsHRegionTable"
Public Const gsPARAMETERKEY_HREGIONFIELD = "Param_FieldsHRegion"
Public Const gsPARAMETERKEY_HREGIONDATE = "Param_FieldsHRegionDate"

'WP Constants - The following key is used for static WP field
Public Const gsPARAMETERKEY_WORKINGPATTERN = "Param_FieldsWorkingPattern"
'WP Constants - The following keys are used for historical WP fields
Public Const gsPARAMETERKEY_HWORKINGPATTERNTABLE = "Param_FieldsHWorkingPatternTable"
Public Const gsPARAMETERKEY_HWORKINGPATTERNFIELD = "Param_FieldsHWorkingPattern"
Public Const gsPARAMETERKEY_HWORKINGPATTERNDATE = "Param_FieldsHWorkingPatternDate"

' ABSENCE MODULE CONSTANTS - MODULE KEY
Public Const gsMODULEKEY_ABSENCE = "MODULE_ABSENCE"

' ABSENCE MODULE CONSTANTS - PARAMETER KEYS
Public Const gsPARAMETERKEY_ABSENCETABLE = "Param_TableAbsence"
Public Const gsPARAMETERKEY_ABSENCETYPETABLE = "Param_TableAbsenceType"
Public Const gsPARAMETERKEY_ABSENCESCREEN = "Param_ScreenAbsence"
Public Const gsPARAMETERKEY_ABSENCESTARTDATE = "Param_FieldStartDate"
Public Const gsPARAMETERKEY_ABSENCESTARTSESSION = "Param_FieldStartSession"
Public Const gsPARAMETERKEY_ABSENCEENDDATE = "Param_FieldEndDate"
Public Const gsPARAMETERKEY_ABSENCEENDSESSION = "Param_FieldEndSession"
Public Const gsPARAMETERKEY_ABSENCEREASON = "Param_FieldReason"
Public Const gsPARAMETERKEY_ABSENCEDURATION = "Param_FieldDuration"
Public Const gsPARAMETERKEY_ABSENCECONTINUOUS = "Param_FieldContinuous"

Public Const gsPARAMETERKEY_ABSENCESSPAPPLIES = "Param_FieldSSPApplies"
Public Const gsPARAMETERKEY_ABSENCESSPQUALIFYINGDAYS = "Param_FieldQualifyingDays"
Public Const gsPARAMETERKEY_ABSENCESSPWAITINGDAYS = "Param_FieldWaitingDays"
Public Const gsPARAMETERKEY_ABSENCESSPPAIDDAYS = "Param_FieldPaidDays"

Public Const gsPARAMETERKEY_ABSENCEWORKINGDAYSTYPE = "Param_WorkingDaysType"
Public Const gsPARAMETERKEY_ABSENCEWORKINGDAYSNUMERICVALUE = "Param_WorkingDaysNum"
Public Const gsPARAMETERKEY_ABSENCEWORKINGDAYSPATTERNVALUE = "Param_WorkingDaysPattern"
Public Const gsPARAMETERKEY_ABSENCEWORKINGDAYSFIELD = "Param_FieldWorkingDays"

Public Const gsPARAMETERKEY_ABSENCETYPE = "Param_FieldType"
Public Const gsPARAMETERKEY_ABSENCETYPETYPE = "Param_FieldTypeType"
Public Const gsPARAMETERKEY_ABSENCETYPECODE = "Param_FieldTypeCode"
Public Const gsPARAMETERKEY_ABSENCETYPESSP = "Param_FieldTypeSSP"
Public Const gsPARAMETERKEY_ABSENCETYPECALCODE = "Param_FieldTypeCalCode"
'Public Const gsPARAMETERKEY_ABSENCETYPEINCLUDE = "Param_FieldTypeInclude"
'Public Const gsPARAMETERKEY_ABSENCETYPEBRADFORDINDEX = "Param_FieldTypeBradfordIndex"

Public Const gsPARAMETERKEY_ABSENCECALSTARTMONTH = "Param_FieldStartMonth"
Public Const gsPARAMETERKEY_ABSENCECALWEEKENDSHADING = "Param_OtherWeekendShading"
Public Const gsPARAMETERKEY_ABSENCECALBHOLSHADING = "Param_OtherBHolShading"
Public Const gsPARAMETERKEY_ABSENCECALINCLUDEWORKINGDAYSONLY = "Param_OtherIncludeWorkingsDaysOnly"
Public Const gsPARAMETERKEY_ABSENCECALBHOLINCLUDE = "Param_OtherBHolInclude"
Public Const gsPARAMETERKEY_ABSENCECALSHOWCAPTIONS = "Param_OtherShowCaptions"

Private Declare Function GetModuleFileNameA Lib "kernel32" (ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long

'Currency Module setup keys
Public Const gsMODULEKEY_CURRENCY = "MODULE_CURRENCY"

Public Const gsPARAMETERKEY_CONVERSIONTABLE = "Param_ConversionTable"
Public Const gsPARAMETERKEY_CURRENCYNAMECOLUMN = "Param_CurrencyNameColumn"
Public Const gsPARAMETERKEY_CONVERSIONVALUECOLUMN = "Param_ConversionValueColumn"
Public Const gsPARAMETERKEY_DECIMALCOLUMN = "Param_DecimalColumn"

Public Const gsMODULEKEY_POST = "MODULE_POST"
Public Const gsPARAMETERKEY_POSTTABLE = "Param_PostTable"
Public Const gsPARAMETERKEY_POSTJOBTITLECOLUMN = "Param_PostJobTitleColumn"
Public Const gsPARAMETERKEY_POSTGRADECOLUMN = "Param_PostGradeColumn"
Public Const gsPARAMETERKEY_GRADETABLE = "Param_GradeTable"
Public Const gsPARAMETERKEY_GRADECOLUMN = "Param_GradeColumn"
Public Const gsPARAMETERKEY_NUMLEVELCOLUMN = "Param_NumLevelColumn"
Public Const gsPARAMETERKEY_SUCCESSIONDEF = "Param_SuccessionDef"
Public Const gsPARAMETERKEY_SUCCESSIONALLOWEQUAL = "Param_SuccessionAllowEqual"
Public Const gsPARAMETERKEY_SUCCESSIONRESTRICT = "Param_SuccessionRestrict"
Public Const gsPARAMETERKEY_SUCCESSIONLEVELS = "Param_SuccessionLevels"
Public Const gsPARAMETERKEY_CAREERDEF = "Param_CareerDef"
Public Const gsPARAMETERKEY_CAREERALLOWEQUAL = "Param_CareerAllowEqual"
Public Const gsPARAMETERKEY_CAREERRESTRICT = "Param_CareerRestrict"
Public Const gsPARAMETERKEY_CAREERLEVELS = "Param_CareerLevels"

Public Const gsPARAMETERKEY_ABSENCEPARENTALLEAVETYPE = "Param_AbsenceParentalLeaveType"
Public Const gsPARAMETERKEY_ABSENCECHILDNO = "Param_AbsenceChildNo"

Public Const gsPARAMETERKEY_DEPENDANTSTABLE = "Param_TableDependants"
Public Const gsPARAMETERKEY_DEPENDANTSCHILDNO = "Param_DependantsChildNo"
Public Const gsPARAMETERKEY_DEPENDANTSDATEOFBIRTH = "Param_DependantsDateOfBirth"
Public Const gsPARAMETERKEY_DEPENDANTSADOPTEDDATE = "Param_DependantsAdoptedDate"
Public Const gsPARAMETERKEY_DEPENDANTSDISABLED = "Param_DependantsDisabled"

Public Const gsMODULEKEY_MATERNITY = "MODULE_MATERNITY"
Public Const gsPARAMETERKEY_MATERNITYTABLE = "Param_MaternityTable"
Public Const gsPARAMETERKEY_MATERNITYEWCDATECOLUMN = "Param_MaternityEWCDate"
Public Const gsPARAMETERKEY_MATERNITYLEAVETYPECOLUMN = "Param_MaternityLeaveType"
Public Const gsPARAMETERKEY_MATERNITYLEAVESTARTCOLUMN = "Param_MaternityLeaveStart"
Public Const gsPARAMETERKEY_MATERNITYBABYBIRTHCOLUMN = "Param_MaternityBabyBirth"

' SSINTRANET MODULE CONSTANTS
Public Const gsMODULEKEY_SSINTRANET = "MODULE_SSINTRANET"
Public Const gsPARAMETERKEY_SSINTRANETVIEW = "Param_SelfServiceView"
Public Const gsPARAMETERKEY_LMINTRANETVIEW = "Param_LineManagerView"

' HIERARCHY MODULE CONSTANTS
Public Const gsMODULEKEY_HIERARCHY = "MODULE_HIERARCHY"
Public Const gsPARAMETERKEY_HIERARCHYTABLE = "Param_TableHierarchy"
Public Const gsPARAMETERKEY_IDENTIFIER = "Param_FieldIdentifier"
Public Const gsPARAMETERKEY_REPORTSTO = "Param_FieldReportsTo"
Public Const gsPARAMETERKEY_POSTALLOCATIONTABLE = "Param_TablePostAllocation"
Public Const gsPARAMETERKEY_POSTALLOCSTARTDATE = "Param_FieldStartDate"
Public Const gsPARAMETERKEY_POSTALLOCENDDATE = "Param_FieldEndDate"

' PAYROLL MODULE CONSTANTS
Public Const gsMODULEKEY_ACCORD = "MODULE_ACCORD"
Public Const gsPARAMETERKEY_PURGEOPTION = "Param_PurgeOption"
Public Const gsPARAMETERKEY_PURGEOPTIONPERIOD = "Param_PurgeOptionPeriod"
Public Const gsPARAMETERKEY_PURGEOPTIONPERIODTYPE = "Param_PurgeOptionPeriodType"
Public Const gsPARAMETERKEY_DEFAULTSTATUS = "Param_DefaultStatus"
Public Const gsPARAMETERKEY_STATUSFORUTILITIES = "Param_StatusForUtilities"
Public Const gsPARAMETERKEY_ALLOWDELETE = "Param_AllowDelete"
Public Const gsPARAMETERKEY_ALLOWSTATUSCHANGE = "Param_AllowStatusChange"

' FUSION MODULE CONSTANTS
' This can be deleted when it is confirmed its not needed
'NHRD Prototype Fusion code
'Public Const gsMODULEKEY_FUSION = "MODULE_FUSION"
'Public Const gsPARAMETERKEY_FUSION_PURGEOPTION = "Param_PurgeOption"
'Public Const gsPARAMETERKEY_FUSION_PURGEOPTIONPERIOD = "Param_PurgeOptionPeriod"
'Public Const gsPARAMETERKEY_FUSION_PURGEOPTIONPERIODTYPE = "Param_PurgeOptionPeriodType"
'Public Const gsPARAMETERKEY_FUSION_DEFAULTSTATUS = "Param_DefaultStatus"
'Public Const gsPARAMETERKEY_FUSION_STATUSFORUTILITIES = "Param_StatusForUtilities"
'Public Const gsPARAMETERKEY_FUSION_ALLOWDELETE = "Param_AllowDelete"
'Public Const gsPARAMETERKEY_FUSION_ALLOWSTATUSCHANGE = "Param_AllowStatusChange"

' AUDIT TABLE MODULE CONSTANTS
Public Const gsMODULEKEY_AUDIT = "MODULE_AUDIT"
Public Const gsPARAMETERKEY_AUDITTABLE = "Param_AuditTable"
Public Const gsPARAMETERKEY_AUDITDATECOLUMN = "Param_AuditDateColumn"
Public Const gsPARAMETERKEY_AUDITTIMECOLUMN = "Param_AuditTimeColumn"
Public Const gsPARAMETERKEY_AUDITUSERCOLUMN = "Param_AuditUserColumn"
Public Const gsPARAMETERKEY_AUDITTABLECOLUMN = "Param_AuditTableColumn"
Public Const gsPARAMETERKEY_AUDITCOLUMNCOLUMN = "Param_AuditColumnColumn"
Public Const gsPARAMETERKEY_AUDITOLDVALUECOLUMN = "Param_AuditOldValueColumn"
Public Const gsPARAMETERKEY_AUDITNEWVALUECOLUMN = "Param_AuditNewValueColumn"
Public Const gsPARAMETERKEY_AUDITMODULECOLUMN = "Param_AuditModuleColumn"
Public Const gsPARAMETERKEY_AUDITDESCRIPTIONCOLUMN = "Param_AuditDescriptionColumn"
Public Const gsPARAMETERKEY_AUDITIDCOLUMN = "Param_AuditIDColumn"


' DOCUMENT MANAGEMENT MODULE CONSTANTS
Public Const gsMODULEKEY_DOCMANAGEMENT = "MODULE_DOCUMENTMANAGEMENT"
Public Const gsPARAMETERKEY_DOCMAN_CATEGORYTABLE = "Param_DocmanCatageoryTable"
Public Const gsPARAMETERKEY_DOCMAN_CATEGORYCOLUMN = "Param_DocManCatageoryColumn"
Public Const gsPARAMETERKEY_DOCMAN_TYPETABLE = "Param_DocmanTypeTable"
Public Const gsPARAMETERKEY_DOCMAN_TYPECOLUMN = "Param_DocManTypeColumn"
Public Const gsPARAMETERKEY_DOCMAN_TYPECATEGORYCOLUMN = "Param_DocManTypeCategoryColumn"

Public Const gsMODULEKEY_SQL = "MODULE_SQL"
Public Const giEXPRVALUE_BYREF_OFFSET = 100

Public Const gsSSP_PROCEDURENAME = "spsys_absencessp"
Public Const gsWorkingDaysBetween2Dates_PROCEDURENAME = "sp_ASRFn_WorkingDaysBetweenTwoDates"

' Pictures
Public Const ChunkSize = 2 ^ 14

' Workflow
Public Const WFITEMPROPERTYCOUNT = 82
Public Const WORKFLOWWEBFORM_MINSIZE_FILEUPLOAD = 1
Public Const WORKFLOWWEBFORM_MAXSIZE_FILEUPLOAD = 8000
Public Const WORKFLOWWEBFORM_MAXSIZE_CHARINPUT = VARCHAR_MAX_Size
Public Const WORKFLOWWEBFORM_MAXSIZE_NUMINPUT = 15

' Email links
Public Const strDelimStart As String = "«"   'asc = 171
Public Const strDelimStop As String = "»"    'asc = 187

