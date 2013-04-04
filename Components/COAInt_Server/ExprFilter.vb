Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class clsExprFilter
	
	' Component definition variables.
	Private mlngFilterID As Integer
	Private msFilterName As String
	Private miReturnType As Short
	
	' Class handling variables.
	Private mobjBaseComponent As clsExprComponent
	
	
	
	
	Public Function RuntimeCode(ByRef psRuntimeCode As String, ByRef palngSourceTables As Object, ByRef pfApplyPermissions As Boolean, ByRef pfValidating As Boolean, ByRef pavPromptedValues As Object, Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean
		
		Dim objFilter As clsExprExpression
		
		Dim strRuntimeCode As String
		Dim bOK As Boolean
		
		' Instantiate and generate the runtime for the filter expression.
		objFilter = New clsExprExpression
		With objFilter
			.ExpressionID = mlngFilterID
			.ConstructExpression()
			bOK = .RuntimeCode(strRuntimeCode, palngSourceTables, pfApplyPermissions, pfValidating, pavPromptedValues, plngFixedExprID, psFixedSQLCode)
		End With
		
		' Return different value depending on passed in parameters
		If mlngFilterID = plngFixedExprID Then
			psRuntimeCode = psFixedSQLCode
			RuntimeCode = True
		Else
			psRuntimeCode = strRuntimeCode
			RuntimeCode = bOK
		End If
		
		'UPGRADE_NOTE: Object objFilter may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objFilter = Nothing
		
		
		'  If mlngFilterID = plngFixedExprID Then
		'    psRuntimeCode = psFixedSQLCode
		'    RuntimeCode = True
		'  Else
		'    ' Instantiate the filter expression.
		'    Set objFilter = New clsExprExpression
		'
		'    With objFilter
		'      ' Construct the filter expression.
		'      .ExpressionID = mlngFilterID
		'      .ConstructExpression
		'      RuntimeCode = .RuntimeCode(psRuntimeCode, palngSourceTables, pfApplyPermissions, pfValidating, pavPromptedValues, plngFixedExprID, psFixedSQLCode)
		'    End With
		'
		'    Set objFilter = Nothing
		'  End If
		
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
			Printer.Print("Filter : " & ComponentDescription)
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
		Dim objFilterCopy As New clsExprFilter
		
		' Copy the component's basic properties.
		objFilterCopy.FilterID = mlngFilterID
		
		CopyComponent = objFilterCopy
		
		' Disassociate object variables.
		'UPGRADE_NOTE: Object objFilterCopy may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objFilterCopy = Nothing
		
	End Function
	
	Public ReadOnly Property ComponentType() As Short
		Get
			' Return the component type.
			ComponentType = modExpression.ExpressionComponentTypes.giCOMPONENT_FILTER
			
		End Get
	End Property
	
	
	Public ReadOnly Property ReturnType() As Short
		Get
			' Return the filter's return type.
			Dim objFilter As clsExprExpression
			
			' Instantiate the filter expression.
			objFilter = New clsExprExpression
			
			With objFilter
				' Construct the filter expression.
				.ExpressionID = mlngFilterID
				.ConstructExpression()
				.ValidateExpression(False)
				miReturnType = .ReturnType
			End With
			
			'UPGRADE_NOTE: Object objFilter may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objFilter = Nothing
			
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
			' Return the filter's name.
			ComponentDescription = msFilterName
			
		End Get
	End Property
	
	
	
	
	Public Property FilterID() As Integer
		Get
			' Return the filter ID property.
			FilterID = mlngFilterID
			
		End Get
		Set(ByVal Value As Integer)
			' Set the filter ID property.
			mlngFilterID = Value
			
			ReadFilter()
			
		End Set
	End Property
	
	
	Public Function ContainsExpression(ByRef plngExprID As Integer) As Boolean
		' Retrun TRUE if the current expression (or any of its sub expressions)
		' contains the given expression. This ensures no cyclic expressions get created.
		'JPD 20040507 Fault 8600
		On Error GoTo ErrorTrap
		
		' Check if the calc component IS the one we're checking for.
		ContainsExpression = (plngExprID = mlngFilterID)
		
		If Not ContainsExpression Then
			' The calc component IS NOT the one we're checking for.
			' Check if it contains the one we're looking for.
			ContainsExpression = HasExpressionComponent(mlngFilterID, plngExprID)
		End If
		
TidyUpAndExit: 
		Exit Function
		
ErrorTrap: 
		ContainsExpression = True
		Resume TidyUpAndExit
		
	End Function
	
	
	
	
	
	Public Function WriteComponent() As Object
		' Write the component definition to the component recordset.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		Dim sSQL As String
		
		fOK = True
		
		sSQL = "INSERT INTO ASRSysExprComponents" & " (componentID, exprID, type, FilterID, valueLogic)" & " VALUES(" & Trim(Str(mobjBaseComponent.ComponentID)) & "," & " " & Trim(Str(mobjBaseComponent.ParentExpression.ExpressionID)) & "," & " " & Trim(Str(modExpression.ExpressionComponentTypes.giCOMPONENT_FILTER)) & "," & " " & Trim(Str(mlngFilterID)) & ", " & " 0)"
		gADOCon.Execute(sSQL,  , ADODB.CommandTypeEnum.adCmdText)
		
TidyUpAndExit: 
		'UPGRADE_WARNING: Couldn't resolve default property of object WriteComponent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		WriteComponent = fOK
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	Private Sub ReadFilter()
		' Read the filter definition from the database.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		Dim sSQL As String
		Dim rsFilter As ADODB.Recordset
		
		' Set default values.
		msFilterName = "<unknown>"
		
		' Get the filter definition.
		sSQL = "SELECT name, returnType" & " FROM ASRSysExpressions" & " WHERE exprID = " & Trim(Str(mlngFilterID))
		rsFilter = datGeneral.GetRecords(sSQL)
		With rsFilter
			fOK = Not (.EOF And .BOF)
			
			If fOK Then
				msFilterName = .Fields("Name").Value
				miReturnType = .Fields("ReturnType").Value
			End If
			
			.Close()
		End With
		
TidyUpAndExit: 
		'UPGRADE_NOTE: Object rsFilter may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsFilter = Nothing
		Exit Sub
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Sub
	
	Public Function UDFCode(ByRef psRuntimeCode() As String, ByRef palngSourceTables As Object, ByRef pfApplyPermissions As Boolean, ByRef pfValidating As Boolean, Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean
		
		Dim objFilter As clsExprExpression
		
		If mlngFilterID = plngFixedExprID Then
			UDFCode = True
		Else
			' Instantiate the filter expression.
			objFilter = New clsExprExpression
			
			With objFilter
				' Construct the filter expression.
				.ExpressionID = mlngFilterID
				.ConstructExpression()
				UDFCode = .UDFCode(psRuntimeCode, palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
			End With
			
			'UPGRADE_NOTE: Object objFilter may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objFilter = Nothing
		End If
		
	End Function
End Class