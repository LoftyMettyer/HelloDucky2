Attribute VB_Name = "modGlobals"
Option Explicit

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

Public gbCMGExportEnabled As Boolean
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

' Payroll module
Public gbAccordPayrollModule As Boolean
Public gbOpenPayModule As Boolean

Public gbManualRecursionLevel As Boolean
Public giManualRecursionLevel As Integer

Public gbDisableSpecialFunctionAutoUpdate As Boolean
Public gbReorganiseIndexesInOvernightJob As Boolean

