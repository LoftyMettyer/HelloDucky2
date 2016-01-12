Attribute VB_Name = "modGlobals"
'Public classes
Public Application As DataMgr.Application
Public Database As DataMgr.Database
Public datGeneral As DataMgr.clsGeneral
Public objEmail As clsEmail
Public gobjDataAccess As DataMgr.clsDataAccess
Public gobjPerformance As DataMgr.clsPerformance
Public gsConnectionString As String

' Holds postcode for AFD and Quick Address
Public oPostCode As DataMgr.PostCode

' New Progress Bar - Global class to be used for progress bars
Public gobjProgress As New clsProgress

' Utility Run Log - Global class used by utilities and batch jobs
Public gobjEventLog As New clsEventLog

'Diary Stuff (MH19991021)
Public gobjDiary As New DataMgr.clsDiary

'Hold the current database name
Public gsDatabaseName As String
Public gsServerName As String
Public gsCustomerName As String

' Windows authentication stuff
Public gbUseWindowsAuthentication As Boolean
Public gstrWindowsCurrentDomain As String
Public gstrWindowsCurrentUser As String

Public gblnBatchMode As Boolean
Public gblnReportPackMode As Boolean
Public gstrHTMLOutputTOC As String
'Just do batch jobs (without prompting) then log off
Public gblnBatchJobsOnly As Boolean

'Don't check printer or MS Office etc..
Public gblnStartupPrinter As Boolean
Public gblnStartupMSOffice As Boolean

' Automatic logon
Public gblnAutomaticLogon As Boolean

Public gblnResetPrinterDefaultBack As Boolean

Public gobjOperatorDefs As New clsOperatorDefs
Public gobjFunctionDefs As New clsFunctionDefs

' Display Find Window or RecEdit when selected from the database menu
' and also go straight to new record if we are displaying recedit first
Public gcPrimary As DefaultDisplay
Public gcHistory As DefaultDisplay
Public gcLookUp As DefaultDisplay
Public gcQuickAccess As DefaultDisplay

Public UI As DataMgr.UI

Public gcoLookupValues As CLookupValues
Public gcoTablePrivileges As CTablePrivileges
Public gcolColumnPrivilegesCollection As Collection
Public gcolHistoryScreensCollection As Collection
Public gcolSummaryFieldsCollection As Collection
Public gcolScreens As clsScreens
Public gcolScreenControls As Collection

Public gsPhotoPath As String
Public gsOLEPath As String
Public gsLocalOLEPath As String
Public gsCrystalPath As String
Public gsDocumentsPath As String

Public gsUserName As String       ' NB. This is actually the SQL LOGIN name
Public gsSQLUserName As String    ' NB. This is actually the SQL USER name
Public gsPassword As String
Public gfCurrentUserIsSysSecMgr As Boolean


Public gADOCon As ADODB.Connection

Public gbForceLogonScreen As Boolean
Public gsConnectString As String
Public gsUserGroup As String


Public glngDesktopBitmapID As Long
Public glngDesktopBitmapLocation As BackgroundLocationTypes
Public glngDeskTopColour As Long

Public gbActivateJobServer As Boolean       ' Allow the Data Manager to run the job seeker

Public gbAccordEnabled As Boolean           ' Is the Payroll Transfer module enabled
Public gbCMGEnabled As Boolean              ' Is the CMG module enabled
Public gbXMLExportEnabled As Boolean        ' Are XML exports licensed
Public gbTableExportEnabled As Boolean      ' Are exports direct to tables licensed
Public gstrDefaultPrinterName As String     ' Default printer device name

Public gobjErrorStack As clsErrorStack      ' Standard Error Handler

Public glngSQLVersion As Double             ' SQL Version

' Output options
Public giOfficeVersion_Word As Integer      ' Microsoft Word Version
Public giOfficeVersion_Excel As Integer     ' Microsoft Excel Version

Public gcolSystemPermissions As Collection  ' Holds system permissions for this user
Public gbEnableUDFFunctions As Boolean


Public gbReadToolbarDefaults As Boolean
Public gbCloseDefSelAfterRun As Boolean
Public gbRecentDisplayDefSel As Boolean
Public glngCurrentCategoryID As Long
Public gbRememberDefSelID As Boolean

Public gdUtilityRunDate As Date
Public glngUtilityRunID As Long

Public giWeekdayStart As VbDayOfWeek

Public giWindowState As FormWindowStateConstants
Public glngWindowLeft As Long
Public glngWindowTop As Long
Public glngWindowHeight As Long
Public glngWindowWidth As Long

Public gbWorkflowEnabled As Boolean           ' Is the Workflow module enabled
Public gbWorkflowOutOfOfficeEnabled As Boolean ' Is the Workflow module enabled AND have the OutOfOffice parameters been configured

Public gbVersion1Enabled As Boolean

Public ASRDEVELOPMENT As Boolean

Public gobjLicence As New clsLicence
