Attribute VB_Name = "modGlobals"
Option Explicit

Public gsDatabaseName As String

' Globals for the desktop settings
Public glngDesktopBitmapID As Long
Public glngDesktopBitmapLocation As BackgroundLocationTypes
Public glngDeskTopColour As Long

' Globals for the Audit (CMG Export)
Public gbCMGExportUseCSV As Boolean
Public gbCMGIgnoreBlanks As Boolean
Public gbCMGReverseDateChanged As Boolean
Public giCMGExportFileCodeOrderID As Integer
Public giCMGEXportRecordIDOrderID As Integer
Public giCMGExportFieldCodeOrderID As Integer
Public giCMGExportOutputColumnOrderID As Integer
Public giCMGExportLastChangeDateOrderID As Integer

Public gbCMGExportFileCode As Boolean
Public gbCMGExportFieldCode As Boolean
Public gbCMGExportLastChangeDate As Boolean
Public giCMGExportFileCodeSize As Integer
Public giCMGEXportRecordIDSize As Integer
Public giCMGExportFieldCodeSize As Integer
Public giCMGExportOutputColumnSize As Integer
Public giCMGExportLastChangeDateSize As Integer

Public glngExpressionViewColours As ExpressionColour
Public glngExpressionViewNodes As ExpressionSaveView

Public gbMaximizeScreens As Boolean

Public gbRememberDBColumnsView As Boolean
Public gpropShowColumns_DataMgr As SystemMgr.Properties
Public gpropShowColumns_DataMgrTable As SystemMgr.Properties
Public gpropShowColumns_PictMgr As SystemMgr.Properties
Public gpropShowColumns_ViewMgr As SystemMgr.Properties

' Postcode modules
Public gbAFDEnabled As Boolean
Public gbQAddressEnabled As Boolean

'' Fusion module
'Public gbFusionModule As Boolean

' Payroll module
Public gbOpenPayModule As Boolean

Public gbManualRecursionLevel As Boolean
Public giManualRecursionLevel As Integer

Public gbDisableSpecialFunctionAutoUpdate As Boolean
Public gbReorganiseIndexesInOvernightJob As Boolean

Public ASRDEVELOPMENT As Boolean
Public mblnAutoLabelling As Boolean
Public mblnDisplayScrOpen As Boolean
Public mblnDisplayWorkflowOpen As Boolean

Public glngSQLVersion As Double
Public gstrSQLFullVersion As String

Public gbCanUseWindowsAuthentication As Boolean
Public gbUseWindowsAuthentication As Boolean
Public gstrWindowsCurrentDomain As String
Public gstrWindowsCurrentUser As String
Public giSQLServerAuthenticationType As SystemMgr.SQLServerAuthenticationType
Public gbAttemptRecovery As Boolean

Public gADOCon As ADODB.Connection
Public gobjHRProEngine As SystemFramework.SysMgr

Public daoWS As DAO.Workspace
Public daoDb As DAO.Database
Public recTabEdit As DAO.Recordset
Public recTableValidationEdit As DAO.Recordset
Public recSummaryEdit As DAO.Recordset
Public recColEdit As DAO.Recordset
Public recDiaryEdit As DAO.Recordset
Public recContValEdit As DAO.Recordset
Public recRelEdit As DAO.Recordset
Public recHistScrEdit As DAO.Recordset
Public recScrEdit As DAO.Recordset
Public recPageCaptEdit As DAO.Recordset
Public recCtrlEdit As DAO.Recordset
Public recPictEdit As DAO.Recordset
Public recOrdEdit As DAO.Recordset
Public recOrdItemEdit As DAO.Recordset
Public recWorkflowEdit As DAO.Recordset
Public recWorkflowElementEdit As DAO.Recordset
Public recWorkflowLinkEdit As DAO.Recordset
Public recWorkflowElementItemEdit As DAO.Recordset
Public recWorkflowElementItemValuesEdit As DAO.Recordset
Public recWorkflowElementColumnEdit As DAO.Recordset
Public recWorkflowElementValidationEdit As DAO.Recordset
Public recWorkflowTriggeredLinks As DAO.Recordset
Public recWorkflowTriggeredLinkColumns As DAO.Recordset

Public recEmailAddrEdit As DAO.Recordset
Public recEmailLinksEdit As DAO.Recordset
Public recEmailRecipientsEdit As DAO.Recordset
Public recEmailLinksColumnsEdit As DAO.Recordset

Public recLinkContentEdit As DAO.Recordset
Public recOutlookFolders As DAO.Recordset
Public recOutlookLinks As DAO.Recordset
Public recOutlookLinksColumns As DAO.Recordset
Public recOutlookLinksDestinations As DAO.Recordset

Public glngAMStartTime As Long
Public glngAMEndTime As Long
Public glngPMStartTime As Long
Public glngPMEndTime As Long

Public recExprEdit As DAO.Recordset
Public recCompEdit As DAO.Recordset
Public recViewEdit As DAO.Recordset
Public recViewColEdit As DAO.Recordset
Public recViewScreens As DAO.Recordset
Public recModuleSetup As DAO.Recordset
Public recModuleRelatedColumns As DAO.Recordset
Public recMailMerge As DAO.Recordset
Public gsUserName As String
Public gsActualSQLLogin As String
Public gbCurrentUserIsSysSecMgr As Boolean
Public gbIsUserSystemAdmin As Boolean
Public gsSecurityGroup As String

Public Application As SystemMgr.clsApplication
Public ODBC As SystemMgr.clsODBC
Public UI As New SystemMgr.clsUI

Public gsTempDatabaseName As String

Public gobjProgress As clsProgress
Public gfRefreshStoredProcedures As Boolean
Public gobjOperatorDefs As New clsOperatorDefs
Public gobjFunctionDefs As New clsFunctionDefs

Public gbEnableUDFFunctions As Boolean
Public glngPageNum As Long
Public glngBottom As Long
Public gstrDefaultPrinterName As String     ' Default printer device name

Public giWindowState As FormWindowStateConstants
Public glngWindowLeft As Long
Public glngWindowTop As Long
Public glngWindowHeight As Long
Public glngWindowWidth As Long

' Automatic logon
Public gblnAutomaticLogon As Boolean
Public gblnAutomaticScript As Boolean
Public gblnAutomaticSave As Boolean

' Email links
Public glngExpressionTableIDForDeleteTrigger As Long
Public glngEmailMethod As Long
Public gstrEmailProfile As String
Public gstrEmailServer As String
Public gstrEmailAccount As String

Public glngEmailDateFormat As Long
Public gstrEmailAttachmentPath As String
Public gstrEmailTestAddr As String

Public gstrUpdateEmailCode As String
Public gstrDeleteEmailCode As String

' Directory paths
Public gsLogDirectory As String
Public gsApplicationPath As String

Public gobjLicence As New clsLicence
Public gbLicenceExpired As Boolean

Public gobjDefaultScreenFont As New StdFont
Public glngDefaultScreenForeColor As Long

