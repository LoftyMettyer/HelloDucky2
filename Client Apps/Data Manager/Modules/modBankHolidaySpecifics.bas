Attribute VB_Name = "modBankHolidaySpecifics"
Option Explicit

Private Const gsPARAMETERKEY_BHOLREGIONTABLE = "Param_TableBHolRegion"
Private Const gsPARAMETERKEY_BHOLREGION = "Param_FieldBHolRegion"
Private Const gsPARAMETERKEY_BHOLTABLE = "Param_TableBHol"
Private Const gsPARAMETERKEY_BHOLDATE = "Param_FieldBHolDate"
Private Const gsPARAMETERKEY_BHOLDESCRIPTION = "Param_FieldBHolDescription"

Public gfBankHolidaysEnabled As Boolean

' Bank Holiday Region Table
Public glngBHolRegionTableID As Long
Public gsBHolRegionTableName As String

' Bank Holiday Region Column
Public glngBHolRegionID As Long
Public gsBHolRegionColumnName As String

' Bank Holiday Instances Table
Public glngBHolTableID As Long
Public gsBHolTableName As String
Public gsBHolTableRealSource As String

' Bank Holiday Instances Date Column
Public glngBHolDateID As Long
Public gsBHolDateColumnName As String

' Bank Holiday Instances Description Column
Public glngBHolDescriptionID As Long
Public gsBHolDescriptionColumnName As String

Public Sub ReadBankHolidayParameters()
  
  Dim objTable As CTablePrivilege
  
  On Error GoTo ReadParametersERROR
  
  ' Bank Holiday Region Table and Column
  glngBHolRegionTableID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLREGIONTABLE))
  If glngBHolRegionTableID > 0 Then
    gsBHolRegionTableName = datGeneral.GetTableName(glngBHolRegionTableID)
  Else
    gsBHolRegionTableName = ""
  End If

  glngBHolRegionID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLREGION))
  If glngBHolRegionID > 0 Then
    gsBHolRegionColumnName = datGeneral.GetColumnName(glngBHolRegionID)
  Else
    gsBHolRegionColumnName = ""
  End If

  ' Bank Holiday Instance Table and Columns

  glngBHolTableID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLTABLE))
  If glngBHolTableID > 0 Then
    gsBHolTableName = datGeneral.GetTableName(glngBHolTableID)
  
    ' Get the realsource into a variable too
    Set objTable = gcoTablePrivileges.FindTableID(glngBHolTableID)
    gsBHolTableRealSource = objTable.RealSource
    Set objTable = Nothing
  
  Else
    gsBHolTableName = ""
  End If

  glngBHolDateID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLDATE))
  If glngBHolDateID > 0 Then
    gsBHolDateColumnName = datGeneral.GetColumnName(glngBHolDateID)
  Else
    gsBHolDateColumnName = ""
  End If

  glngBHolDescriptionID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLDESCRIPTION))
  If glngBHolDescriptionID > 0 Then
    gsBHolDescriptionColumnName = datGeneral.GetColumnName(glngBHolDescriptionID)
  Else
    gsBHolDescriptionColumnName = ""
  End If
  
  Set objTable = Nothing
  
  Exit Sub
  
ReadParametersERROR:

  COAMsgBox "Error reading the Bank Holiday parameters." & vbCrLf & _
         Err.Description, vbExclamation + vbOKOnly, App.Title
  gfBankHolidaysEnabled = False
  Set objTable = Nothing

End Sub


Public Function ValidateBankHolidayParameters() As Boolean
  
' RH 01/12/00
' There is no real need for this, because Bank Holidays should
' be an optional thing, ie, the calcs/calendar should still
' function even if bank hols are not set up.
  
'  On Error GoTo ValidateERROR
'
'  ' Validate the configuration of the Bank Holiday parameters
'  Dim fValid As Boolean
'
'  ' Default to true
'  fValid = True
'
'  ' Now check the bank holiday module setup
'
'  If fValid Then
'    fValid = (glngBHolTableID > 0)
'    If Not fValid Then
'      COAMsgBox "Bank Holidays are not properly configured." & vbCrLf & _
'         "The Bank Holiday table is not defined.", vbOKOnly, App.ProductName
'    End If
'  End If
'
'  If fValid Then
'    fValid = (glngBholRegionTableID > 0)
'    If Not fValid Then
'      COAMsgBox "Bank Holidays are not properly configured." & vbCrLf & _
'         "The Bank Holiday Region table is not defined.", vbOKOnly, App.ProductName
'    End If
'  End If
'
'  If fValid Then
'    fValid = (glngBHolRegionID > 0)
'    If Not fValid Then
'      COAMsgBox "Bank Holidays are not properly configured." & vbCrLf & _
'         "The Bank Holiday Region column is not defined.", vbOKOnly, App.ProductName
'    End If
'  End If
'
'  If fValid Then
'    fValid = (glngBHolDateID > 0)
'    If Not fValid Then
'      COAMsgBox "Bank Holidays are not properly configured." & vbCrLf & _
'         "The Bank Holiday Date column is not defined.", vbOKOnly, App.ProductName
'    End If
'  End If
'
'  If fValid Then
'    fValid = (glngBHolDescriptionID > 0)
'    If Not fValid Then
'      COAMsgBox "Bank Holidays are not properly configured." & vbCrLf & _
'         "The Bank Holiday Description column is not defined.", vbOKOnly, App.ProductName
'    End If
'  End If
'
'ResumePoint:
'
'  ValidateBankHolidayParameters = fValid
'
'ValidateERROR:
'
'  COAMsgBox "Error whilst validating Bank Holiday parameters." & vbCrLf & _
'         Err.Description, vbExclamation + vbOKOnly, App.Title
'  fValid = False
'  Resume ResumePoint
  
End Function







