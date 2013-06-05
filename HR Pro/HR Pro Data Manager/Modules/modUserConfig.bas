Attribute VB_Name = "modUserConfig"
Option Explicit

'Should check for diary events on startup
Public Const gblnDiaryStartUpCheck As Boolean = True

'Should check for diary events throughout the day

' 12/06/00 - RH. Changed from a constant to a public variable
'            This variable is changed when the user toggles the
'            diary check status from within the Administration/Diary
'            menu option on frmMain.
'Public Const gblnDiaryConstCheck As Boolean = True
Public gblnDiaryConstCheck As Boolean

'Number of minutes between each diary check for new entries
Public Const glngDiaryIntervalCheck As Long = 1

Public Const SQLWhereMandatoryColumn = _
  "(Rtrim(DefaultValue) = '' OR (Rtrim(DefaultValue) = '__/__/____') and DataType = 11)" & _
  " AND Convert(int,isnull(dfltValueExprID,0)) = 0" & _
  " AND CalcExprID = 0" & _
  " AND Mandatory = '1'" & _
  " AND ColumnType <> 4 "

Public gbPrinterPrompt As Boolean
Public gbPrinterConfirm As Boolean
