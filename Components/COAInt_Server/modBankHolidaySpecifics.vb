Option Strict Off
Option Explicit On
Module modBankHolidaySpecifics
	
	Private Const gsPARAMETERKEY_BHOLREGIONTABLE As String = "Param_TableBHolRegion"
	Private Const gsPARAMETERKEY_BHOLREGION As String = "Param_FieldBHolRegion"
	Private Const gsPARAMETERKEY_BHOLTABLE As String = "Param_TableBHol"
	Private Const gsPARAMETERKEY_BHOLDATE As String = "Param_FieldBHolDate"
	Private Const gsPARAMETERKEY_BHOLDESCRIPTION As String = "Param_FieldBHolDescription"
	
	Public gfBankHolidaysEnabled As Boolean
	
	' Bank Holiday Region Table
	Public glngBHolRegionTableID As Integer
	Public gsBHolRegionTableName As String
	
	' Bank Holiday Region Column
	Public glngBHolRegionID As Integer
	Public gsBHolRegionColumnName As String
	
	' Bank Holiday Instances Table
	Public glngBHolTableID As Integer
	Public gsBHolTableName As String
	Public gsBHolTableRealSource As String
	
	' Bank Holiday Instances Date Column
	Public glngBHolDateID As Integer
	Public gsBHolDateColumnName As String
	
	' Bank Holiday Instances Description Column
	Public glngBHolDescriptionID As Integer
	Public gsBHolDescriptionColumnName As String
	
	Public Sub ReadBankHolidayParameters()
		
		Dim objTable As CTablePrivilege
		
		On Error GoTo ReadParametersERROR
		
		gfBankHolidaysEnabled = True
		
		' Bank Holiday Region Table and Column
		glngBHolRegionTableID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLREGIONTABLE))
		If glngBHolRegionTableID > 0 Then
			gsBHolRegionTableName = datGeneral.GetTableName(glngBHolRegionTableID)
		Else
			gsBHolRegionTableName = ""
			gfBankHolidaysEnabled = False
		End If
		
		glngBHolRegionID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLREGION))
		If glngBHolRegionID > 0 Then
			gsBHolRegionColumnName = datGeneral.GetColumnName(glngBHolRegionID)
		Else
			gsBHolRegionColumnName = ""
			gfBankHolidaysEnabled = False
		End If
		
		' Bank Holiday Instance Table and Columns
		
		glngBHolTableID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLTABLE))
		If glngBHolTableID > 0 Then
			gsBHolTableName = datGeneral.GetTableName(glngBHolTableID)
			
			' Get the realsource into a variable too
			objTable = gcoTablePrivileges.FindTableID(glngBHolTableID)
			gsBHolTableRealSource = objTable.RealSource
			'UPGRADE_NOTE: Object objTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objTable = Nothing
			
		Else
			gsBHolTableName = ""
			gfBankHolidaysEnabled = False
		End If
		
		glngBHolDateID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLDATE))
		If glngBHolDateID > 0 Then
			gsBHolDateColumnName = datGeneral.GetColumnName(glngBHolDateID)
		Else
			gsBHolDateColumnName = ""
			gfBankHolidaysEnabled = False
		End If
		
		glngBHolDescriptionID = Val(GetModuleParameter(gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLDESCRIPTION))
		If glngBHolDescriptionID > 0 Then
			gsBHolDescriptionColumnName = datGeneral.GetColumnName(glngBHolDescriptionID)
		Else
			gsBHolDescriptionColumnName = ""
			gfBankHolidaysEnabled = False
		End If
		
		'UPGRADE_NOTE: Object objTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objTable = Nothing
		
		Exit Sub
		
ReadParametersERROR: 
		
		'NO MSGBOX ON THE SERVER ! - MsgBox "Error reading the Bank Holiday parameters." & vbNewLine & _
		'Err.Description, vbExclamation + vbOKOnly, App.Title
		gfBankHolidaysEnabled = False
		'UPGRADE_NOTE: Object objTable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objTable = Nothing
		
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
		'      MsgBox "Bank Holidays are not properly configured." & vbNewLine & _
		''         "The Bank Holiday table is not defined.", vbOKOnly, App.ProductName
		'    End If
		'  End If
		'
		'  If fValid Then
		'    fValid = (glngBholRegionTableID > 0)
		'    If Not fValid Then
		'      MsgBox "Bank Holidays are not properly configured." & vbNewLine & _
		''         "The Bank Holiday Region table is not defined.", vbOKOnly, App.ProductName
		'    End If
		'  End If
		'
		'  If fValid Then
		'    fValid = (glngBHolRegionID > 0)
		'    If Not fValid Then
		'      MsgBox "Bank Holidays are not properly configured." & vbNewLine & _
		''         "The Bank Holiday Region column is not defined.", vbOKOnly, App.ProductName
		'    End If
		'  End If
		'
		'  If fValid Then
		'    fValid = (glngBHolDateID > 0)
		'    If Not fValid Then
		'      MsgBox "Bank Holidays are not properly configured." & vbNewLine & _
		''         "The Bank Holiday Date column is not defined.", vbOKOnly, App.ProductName
		'    End If
		'  End If
		'
		'  If fValid Then
		'    fValid = (glngBHolDescriptionID > 0)
		'    If Not fValid Then
		'      MsgBox "Bank Holidays are not properly configured." & vbNewLine & _
		''         "The Bank Holiday Description column is not defined.", vbOKOnly, App.ProductName
		'    End If
		'  End If
		'
		'ResumePoint:
		'
		'  ValidateBankHolidayParameters = fValid
		'
		'ValidateERROR:
		'
		'  MsgBox "Error whilst validating Bank Holiday parameters." & vbNewLine & _
		''         Err.Description, vbExclamation + vbOKOnly, App.Title
		'  fValid = False
		'  Resume ResumePoint
		
	End Function
End Module