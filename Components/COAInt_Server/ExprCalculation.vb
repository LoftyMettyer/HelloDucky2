Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class clsExprCalculation
	
	' Component definition variables.
	Private mlngCalculationID As Integer
	Private msCalculationName As String
	Private miReturnType As Short
	
	' Class handling variables.
	Private mobjBaseComponent As clsExprComponent
	
	
	Public Function ContainsExpression(ByRef plngExprID As Integer) As Boolean
		' Retrun TRUE if the current expression (or any of its sub expressions)
		' contains the given expression. This ensures no cyclic expressions get created.
		'JPD 20040507 Fault 8600
		On Error GoTo ErrorTrap
		
		' Check if the calc component IS the one we're checking for.
		ContainsExpression = (plngExprID = mlngCalculationID)
		
		If Not ContainsExpression Then
			' The calc component IS NOT the one we're checking for.
			' Check if it contains the one we're looking for.
			ContainsExpression = HasExpressionComponent(mlngCalculationID, plngExprID)
		End If
		
TidyUpAndExit: 
		Exit Function
		
ErrorTrap: 
		ContainsExpression = True
		Resume TidyUpAndExit
		
	End Function
	
	
	
	
	
	Public Function UDFCode(ByRef psRuntimeCode() As String, ByRef palngSourceTables As Object, ByRef pfApplyPermissions As Boolean, ByRef pfValidating As Boolean, Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean
		
		Dim objCalc As clsExprExpression
		
		If mlngCalculationID = plngFixedExprID Then
			UDFCode = True
		Else
			' Instantiate the calculation expression.
			objCalc = New clsExprExpression
			
			With objCalc
				' Construct the calculation expression.
				.ExpressionID = mlngCalculationID
				.ConstructExpression()
				UDFCode = .UDFCode(psRuntimeCode, palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
			End With
			
			'UPGRADE_NOTE: Object objCalc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objCalc = Nothing
		End If
		
	End Function
	
	Public Function RuntimeCode(ByRef psRuntimeCode As String, ByRef palngSourceTables As Object, ByRef pfApplyPermissions As Boolean, ByRef pfValidating As Boolean, ByRef pavPromptedValues As Object, Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean
		
		Dim objCalc As clsExprExpression
		
		If mlngCalculationID = plngFixedExprID Then
			RuntimeCode = True
			psRuntimeCode = psFixedSQLCode
		Else
			' Instantiate the calculation expression.
			objCalc = New clsExprExpression
			
			With objCalc
				' Construct the calculation expression.
				.ExpressionID = mlngCalculationID
				.ConstructExpression()
				RuntimeCode = .RuntimeCode(psRuntimeCode, palngSourceTables, pfApplyPermissions, pfValidating, pavPromptedValues, plngFixedExprID, psFixedSQLCode)
			End With
			
			'UPGRADE_NOTE: Object objCalc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objCalc = Nothing
		End If
		
	End Function
	
	
	
	
	
	Public Function PrintComponent(ByRef piLevel As Short) As Boolean
		Dim Printer As New Printer
		' Print the component definition to the printer object.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		
		fOK = True
		
		' Position the printing.
		With Printer
			.CurrentX = giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
			.CurrentY = .CurrentY + giPRINT_YSPACE
			Printer.Print("Calculation : " & ComponentDescription)
		End With
		
TidyUpAndExit: 
		PrintComponent = fOK
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	Public Function CopyComponent() As Object
		' Copies the selected component.
		' When editting a component we actually copy the component first
		' and edit the copy. If the changes are confirmed then the copy
		' replaces the original. If the changes are cancelled then the
		' copy is discarded.
		Dim objCalcCopy As New clsExprCalculation
		
		' Copy the component's basic properties.
		objCalcCopy.CalculationID = mlngCalculationID
		
		CopyComponent = objCalcCopy
		
		' Disassociate object variables.
		'UPGRADE_NOTE: Object objCalcCopy may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objCalcCopy = Nothing
		
	End Function
	
	Public ReadOnly Property ComponentType() As Short
		Get
			' Return the component type.
			ComponentType = modExpression.ExpressionComponentTypes.giCOMPONENT_CALCULATION
			
		End Get
	End Property
	
	
	Public ReadOnly Property ReturnType() As Short
		Get
			' Return the calculation's return type.
			Dim objCalc As clsExprExpression
			
			' Instantiate the calculation expression.
			objCalc = New clsExprExpression
			
			With objCalc
				' Construct the calculation expression.
				.ExpressionID = mlngCalculationID
				.ConstructExpression()
				.ValidateExpression(False)
				miReturnType = .ReturnType
			End With
			
			'UPGRADE_NOTE: Object objCalc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objCalc = Nothing
			
			ReturnType = miReturnType
			
		End Get
	End Property
	
	
	
	
	
	Public Property BaseComponent() As clsExprComponent
		Get
			' Return the component's base component object.
			BaseComponent = mobjBaseComponent
			
		End Get
		Set(ByVal Value As clsExprComponent)
			' Set the component's base component object property.
			mobjBaseComponent = Value
			
		End Set
	End Property
	
	Public ReadOnly Property ComponentDescription() As String
		Get
			' Return the calculation's name.
			ComponentDescription = msCalculationName
			
		End Get
	End Property
	
	
	
	
	Public Property CalculationID() As Integer
		Get
			' Return the calculation ID property.
			CalculationID = mlngCalculationID
			
		End Get
		Set(ByVal Value As Integer)
			' Set the calculation ID property.
			mlngCalculationID = Value
			
			ReadCalculation()
			
		End Set
	End Property
	
	
	
	
	
	Public Function WriteComponent() As Object
		' Write the component definition to the component recordset.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		Dim sSQL As String
		
		fOK = True
		
		sSQL = "INSERT INTO ASRSysExprComponents" & " (componentID, exprID, type, calculationID, valueLogic)" & " VALUES(" & Trim(Str(mobjBaseComponent.ComponentID)) & "," & " " & Trim(Str(mobjBaseComponent.ParentExpression.ExpressionID)) & "," & " " & Trim(Str(modExpression.ExpressionComponentTypes.giCOMPONENT_CALCULATION)) & "," & " " & Trim(Str(mlngCalculationID)) & ", " & " 0)"
		gADOCon.Execute(sSQL,  , ADODB.CommandTypeEnum.adCmdText)
		
TidyUpAndExit: 
		'UPGRADE_WARNING: Couldn't resolve default property of object WriteComponent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		WriteComponent = fOK
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	Private Sub ReadCalculation()
		' Read the calculation definition from the database.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		Dim sSQL As String
		Dim rsCalculation As ADODB.Recordset
		
		' Set default values.
		msCalculationName = "<unknown>"
		
		' Get the calculation definition.
		sSQL = "SELECT name, returnType" & " FROM ASRSysExpressions" & " WHERE exprID = " & Trim(Str(mlngCalculationID))
		rsCalculation = datGeneral.GetRecords(sSQL)
		With rsCalculation
			fOK = Not (.EOF And .BOF)
			
			If fOK Then
				msCalculationName = .Fields("Name").Value
				miReturnType = .Fields("ReturnType").Value
			End If
			
			.Close()
		End With
		
TidyUpAndExit: 
		'UPGRADE_NOTE: Object rsCalculation may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsCalculation = Nothing
		Exit Sub
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Sub
End Class