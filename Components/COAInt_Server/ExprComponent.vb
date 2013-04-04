Option Strict Off
Option Explicit On
Friend Class clsExprComponent
	
	' Component definition variables.
	Private mlngComponentID As Integer
	Private miComponentType As modExpression.ExpressionComponentTypes
	
	' Class handling variables.
	Private mobjParentExpression As clsExprExpression
	Private mvComponent As Object
	
	' Definition for expanded/unexpanded status of the component
	Private mbExpanded As Boolean
	
	Public Function ContainsExpression(ByRef plngExprID As Integer) As Boolean
		' Retrun TRUE if the current expression (or any of its sub expressions)
		' contains the given expression. This ensures no cyclic expressions get created.
		'JPD 20040507 Fault 8600
		On Error GoTo ErrorTrap
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.ContainsExpression. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ContainsExpression = mvComponent.ContainsExpression(plngExprID)
		
TidyUpAndExit: 
		Exit Function
		
ErrorTrap: 
		ContainsExpression = True
		Resume TidyUpAndExit
		
	End Function
	
	
	
	
	' UDF code for this component
	Public Function UDFCode(ByRef psRuntimeCode() As String, ByRef palngSourceTables As Object, ByRef pfApplyPermissions As Boolean, ByRef pfValidating As Boolean, Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean
		
		' Return the runtime filter SQL code for the component.
		'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.UDFCode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		UDFCode = mvComponent.UDFCode(psRuntimeCode, palngSourceTables, pfApplyPermissions, pfValidating, plngFixedExprID, psFixedSQLCode)
		
	End Function
	
	Public Property ExpandedNode() As Boolean
		Get
			'Return whether this node is expanded or not
			ExpandedNode = mbExpanded
			
		End Get
		Set(ByVal Value As Boolean)
			'Set whether this component node is expanded or not
			mbExpanded = Value
			
			Select Case Me.ComponentType
				Case modExpression.ExpressionComponentTypes.giCOMPONENT_FUNCTION
					'UPGRADE_WARNING: Couldn't resolve default property of object Me.Component.ExpandedNode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Me.Component.ExpandedNode = Value
					
				Case modExpression.ExpressionComponentTypes.giCOMPONENT_EXPRESSION
					'UPGRADE_WARNING: Couldn't resolve default property of object Me.Component.ExpandedNode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Me.Component.ExpandedNode = Value
					
			End Select
			
		End Set
	End Property
	
	
	Public Property Component() As Object
		Get
			' Return the real component object.
			Component = mvComponent
			
		End Get
		Set(ByVal Value As Object)
			' Set the real component object.
			If Not Value Is Nothing Then
				mvComponent = Value
				'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.ComponentType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				miComponentType = mvComponent.ComponentType
			End If
			
		End Set
	End Property
	
	Public ReadOnly Property ComponentDescription() As String
		Get
			' Return a text description of the component.
			'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.ComponentDescription. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ComponentDescription = mvComponent.ComponentDescription
			
		End Get
	End Property
	
	
	
	Public Property ComponentType() As Short
		Get
			' Return the component type property.
			ComponentType = miComponentType
			
		End Get
		Set(ByVal Value As Short)
			' Set the component type property.
			If miComponentType <> Value Then
				miComponentType = Value
				
				'UPGRADE_NOTE: Object mvComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				mvComponent = Nothing
				
				' Instantiate the correct type of component object for
				' the given component type.
				Select Case miComponentType
					
					Case modExpression.ExpressionComponentTypes.giCOMPONENT_FIELD
						mvComponent = New clsExprField
						
					Case modExpression.ExpressionComponentTypes.giCOMPONENT_FUNCTION
						mvComponent = New clsExprFunction
						
					Case modExpression.ExpressionComponentTypes.giCOMPONENT_CALCULATION
						mvComponent = New clsExprCalculation
						
					Case modExpression.ExpressionComponentTypes.giCOMPONENT_VALUE
						mvComponent = New clsExprValue
						
					Case modExpression.ExpressionComponentTypes.giCOMPONENT_OPERATOR
						mvComponent = New clsExprOperator
						
					Case modExpression.ExpressionComponentTypes.giCOMPONENT_TABLEVALUE
						mvComponent = New clsExprTableLookup
						
					Case modExpression.ExpressionComponentTypes.giCOMPONENT_PROMPTEDVALUE
						mvComponent = New clsExprPromptedValue
						
					Case modExpression.ExpressionComponentTypes.giCOMPONENT_CUSTOMCALC
						' Not required.
						
					Case modExpression.ExpressionComponentTypes.giCOMPONENT_EXPRESSION
						mvComponent = New clsExprExpression
						
					Case modExpression.ExpressionComponentTypes.giCOMPONENT_FILTER
						mvComponent = New clsExprFilter
						
				End Select
				
				If Not mvComponent Is Nothing Then
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.BaseComponent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mvComponent.BaseComponent = Me
				End If
			End If
			
		End Set
	End Property
	
	
	Public ReadOnly Property ReturnType() As Short
		Get
			' Return the component's return type.
			'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.ReturnType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ReturnType = mvComponent.ReturnType
			
		End Get
	End Property
	
	
	
	Public Property ComponentID() As Integer
		Get
			' Return the component id property.
			ComponentID = mlngComponentID
			
		End Get
		Set(ByVal Value As Integer)
			' Set the component id property.
			mlngComponentID = Value
			
		End Set
	End Property
	
	
	
	Public Property ParentExpression() As clsExprExpression
		Get
			' Return the component's parent expression.
			ParentExpression = mobjParentExpression
			
		End Get
		Set(ByVal Value As clsExprExpression)
			' Set the component's parent expression property.
			mobjParentExpression = Value
			
		End Set
	End Property
	
	
	
	Public Function RuntimeCode(ByRef psRuntimeCode As String, ByRef palngSourceTables As Object, ByRef pfApplyPermissions As Boolean, ByRef pfValidating As Boolean, ByRef pavPromptedValues As Object, Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean
		' Return the runtime filter SQL code for the component.
		'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.RuntimeCode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RuntimeCode = mvComponent.RuntimeCode(psRuntimeCode, palngSourceTables, pfApplyPermissions, pfValidating, pavPromptedValues, plngFixedExprID, psFixedSQLCode)
		
	End Function
	
	
	
	
	Public Function PrintComponent(ByRef piLevel As Short) As Boolean
		' Print the component definition.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Component.PrintComponent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fOK = Component.PrintComponent(piLevel)
		
TidyUpAndExit: 
		PrintComponent = fOK
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	
	
	Public Function WriteComponent() As Boolean
		' Write the component definition to the component recordset.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		Dim lngNewID As Integer
		
		
		' Update the real component expression id property, and give it
		' a unique component id.
		
		'MH20010712 Need keep manual record of allocated IDs incase users
		'in SYS MGR have created expressions but not yet saved changes
		'lngNewID = UniqueColumnValue("ASRSysExprComponents", "componentID")
		lngNewID = GetUniqueID("ExprComponents", "ASRSysExprComponents", "componentID")
		
		
		
		fOK = (lngNewID > 0)
		
		If fOK Then
			mlngComponentID = lngNewID
			'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.BaseComponent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			mvComponent.BaseComponent = Me
			
			' Instruct the real component to write its definition to the
			' component recordset.
			'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.WriteComponent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			fOK = mvComponent.WriteComponent
		End If
		
TidyUpAndExit: 
		WriteComponent = fOK
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	Public Function NewComponent() As Boolean
		' Define a new component.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		'  Dim frmEdit As frmExprComponent
		
		fOK = True
		
		' Initialize the properties for a new expression.
		InitializeComponent()
		
		'  ' Display the component definition form.
		'  Set frmEdit = New frmExprComponent
		'  With frmEdit
		'    Set .Component = Me
		'    .Show vbModal
		'    fOK = Not .Cancelled
		'  End With
		'
TidyUpAndExit: 
		' Disassociate object variables.
		'  Set frmEdit = Nothing
		NewComponent = fOK
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	
	Public Function CopyComponent() As clsExprComponent
		' Copies the selected component.
		' When editting a component we actually copy the component first
		' and edit the copy. If the changes are confirmed then the copy
		' replaces the original. If the changes are cancelled then the
		' copy is discarded.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		Dim objCopyComponent As New clsExprComponent
		
		' Copy the component's basic properties.
		With objCopyComponent
			.ComponentType = miComponentType
			.ParentExpression = mobjParentExpression
			
			' Instruct the original component to copy itself.
			'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.CopyComponent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.Component = mvComponent.CopyComponent
			'UPGRADE_WARNING: Couldn't resolve default property of object objCopyComponent.Component.BaseComponent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.Component.BaseComponent = objCopyComponent
			
			fOK = Not .Component Is Nothing
		End With
		
		'Copy whether this object is in expanded mode.
		objCopyComponent.ExpandedNode = mbExpanded
		
		
TidyUpAndExit: 
		If fOK Then
			CopyComponent = objCopyComponent
		Else
			'UPGRADE_NOTE: Object CopyComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			CopyComponent = Nothing
		End If
		'UPGRADE_NOTE: Object objCopyComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objCopyComponent = Nothing
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	
	Public Function ModifyComponent() As Boolean
		' Edit the component.
		'  On Error GoTo ErrorTrap
		
		'  Dim fOK As Boolean
		'  Dim frmEdit As frmExprComponent
		'
		'  ' Display the component definition form.
		'  Set frmEdit = New frmExprComponent
		'  With frmEdit
		'    Set .Component = Me
		'    .Show vbModal
		'    fOK = Not .Cancelled
		'  End With
		'
		'TidyUpAndExit:
		'  Set frmEdit = Nothing
		'  ModifyComponent = fOK
		'  Exit Function
		'
		'ErrorTrap:
		'  fOK = False
		'  Resume TidyUpAndExit
		
	End Function
	
	
	
	Private Sub InitializeComponent()
		' Initialize the properties for a new component.
		mlngComponentID = 0
		ComponentType = modExpression.ExpressionComponentTypes.giCOMPONENT_FIELD
		'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.BaseComponent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvComponent.BaseComponent = Me
		
	End Sub
	
	Public Function ConstructComponent(ByRef prsComponents As ADODB.Recordset) As Boolean
		' Read the component definition from the database.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		
		fOK = True
		
		' Initialise the component with the definition from the database.
		ComponentType = prsComponents.Fields("Type").Value
		
		With mvComponent
			Select Case miComponentType
				Case modExpression.ExpressionComponentTypes.giCOMPONENT_FIELD
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.TableID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.TableID = prsComponents.Fields("fieldTableID").Value
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.ColumnID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.ColumnID = prsComponents.Fields("fieldColumnID").Value
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.FieldPassType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.FieldPassType = prsComponents.Fields("fieldPassBy").Value
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.SelectionType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.SelectionType = prsComponents.Fields("fieldSelectionRecord").Value
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.SelectionLine. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.SelectionLine = prsComponents.Fields("fieldSelectionLine").Value
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.SelectionOrderID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.SelectionOrderID = prsComponents.Fields("fieldSelectionOrderID").Value
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.SelectionFilterID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.SelectionFilterID = prsComponents.Fields("FieldSelectionFilter").Value
					
				Case modExpression.ExpressionComponentTypes.giCOMPONENT_FUNCTION
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.FunctionID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.FunctionID = prsComponents.Fields("FunctionID").Value
					
					' Allow user preference of how expression builder initially loads
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.ExpandedNode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					.ExpandedNode = IIf(IsDbNull(prsComponents.Fields("ExpandedNode").Value), False, prsComponents.Fields("ExpandedNode").Value)
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.ExpandedNode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Me.ExpandedNode = .ExpandedNode
					
				Case modExpression.ExpressionComponentTypes.giCOMPONENT_CALCULATION
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.CalculationID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.CalculationID = prsComponents.Fields("CalculationID").Value
					
				Case modExpression.ExpressionComponentTypes.giCOMPONENT_VALUE
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.ReturnType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.ReturnType = prsComponents.Fields("ValueType").Value
					Select Case prsComponents.Fields("ValueType").Value
						Case modExpression.ExpressionValueTypes.giEXPRVALUE_CHARACTER
							'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							.Value = IIf(IsDbNull(prsComponents.Fields("valueCharacter").Value), "", prsComponents.Fields("valueCharacter").Value)
						Case modExpression.ExpressionValueTypes.giEXPRVALUE_NUMERIC
							'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							.Value = IIf(IsDbNull(prsComponents.Fields("valueNumeric").Value), 0, prsComponents.Fields("valueNumeric").Value)
						Case modExpression.ExpressionValueTypes.giEXPRVALUE_LOGIC
							'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							.Value = IIf(IsDbNull(prsComponents.Fields("valueLogic").Value), True, prsComponents.Fields("valueLogic").Value)
						Case modExpression.ExpressionValueTypes.giEXPRVALUE_DATE
							
							'MH20010201 Fault 1576
							'.Value = IIf(IsNull(prsComponents!valueDate), Date, prsComponents!valueDate)
							'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							.Value = prsComponents.Fields("valueDate").Value
					End Select
					
				Case modExpression.ExpressionComponentTypes.giCOMPONENT_OPERATOR
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.OperatorID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.OperatorID = prsComponents.Fields("OperatorID").Value
					
				Case modExpression.ExpressionComponentTypes.giCOMPONENT_TABLEVALUE
					' Do nothing as Table Value components are treated as Value components.
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.TableID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.TableID = prsComponents.Fields("LookupTableID").Value
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.ColumnID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.ColumnID = prsComponents.Fields("LookupColumnID").Value
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.ReturnType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.ReturnType = prsComponents.Fields("ValueType").Value
					
					Select Case prsComponents.Fields("ValueType").Value
						Case modExpression.ExpressionValueTypes.giEXPRVALUE_CHARACTER
							'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							.Value = IIf(IsDbNull(prsComponents.Fields("valueCharacter").Value), "", prsComponents.Fields("valueCharacter").Value)
						Case modExpression.ExpressionValueTypes.giEXPRVALUE_NUMERIC
							'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							.Value = IIf(IsDbNull(prsComponents.Fields("valueNumeric").Value), 0, prsComponents.Fields("valueNumeric").Value)
						Case modExpression.ExpressionValueTypes.giEXPRVALUE_LOGIC
							'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							.Value = IIf(IsDbNull(prsComponents.Fields("valueLogic").Value), True, prsComponents.Fields("valueLogic").Value)
						Case modExpression.ExpressionValueTypes.giEXPRVALUE_DATE
							'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							.Value = prsComponents.Fields("valueDate").Value
					End Select
					
				Case modExpression.ExpressionComponentTypes.giCOMPONENT_PROMPTEDVALUE
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.Prompt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					.Prompt = IIf(IsDbNull(prsComponents.Fields("promptDescription").Value), "", prsComponents.Fields("promptDescription").Value)
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.ValueType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					.ValueType = IIf(IsDbNull(prsComponents.Fields("ValueType").Value), modExpression.ExpressionValueTypes.giEXPRVALUE_CHARACTER, prsComponents.Fields("ValueType").Value)
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.ReturnSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					.ReturnSize = IIf(IsDbNull(prsComponents.Fields("promptSize").Value), 1, prsComponents.Fields("promptSize").Value)
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.ReturnDecimals. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					.ReturnDecimals = IIf(IsDbNull(prsComponents.Fields("promptDecimals").Value), 0, prsComponents.Fields("promptDecimals").Value)
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.ValueFormat. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					.ValueFormat = IIf(IsDbNull(prsComponents.Fields("promptMask").Value), "", prsComponents.Fields("promptMask").Value)
					
					Select Case prsComponents.Fields("ValueType").Value
						Case modExpression.ExpressionValueTypes.giEXPRVALUE_CHARACTER
							'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.DefaultValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							.DefaultValue = IIf(IsDbNull(prsComponents.Fields("valueCharacter").Value), "", prsComponents.Fields("valueCharacter").Value)
						Case modExpression.ExpressionValueTypes.giEXPRVALUE_NUMERIC
							'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.DefaultValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							.DefaultValue = IIf(IsDbNull(prsComponents.Fields("valueNumeric").Value), 0, prsComponents.Fields("valueNumeric").Value)
						Case modExpression.ExpressionValueTypes.giEXPRVALUE_LOGIC
							'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.DefaultValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							.DefaultValue = IIf(IsDbNull(prsComponents.Fields("valueLogic").Value), False, prsComponents.Fields("valueLogic").Value)
						Case modExpression.ExpressionValueTypes.giEXPRVALUE_DATE
							'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.DefaultValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							.DefaultValue = IIf(IsDbNull(prsComponents.Fields("valueDate").Value), Today, prsComponents.Fields("valueDate").Value)
						Case modExpression.ExpressionValueTypes.giEXPRVALUE_TABLEVALUE
							'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.DefaultValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							.DefaultValue = IIf(IsDbNull(prsComponents.Fields("valueCharacter").Value), "", prsComponents.Fields("valueCharacter").Value)
					End Select
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.LookupColumn. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					.LookupColumn = IIf(IsDbNull(prsComponents.Fields("fieldColumnID").Value), 0, prsComponents.Fields("fieldColumnID").Value)
					
				Case modExpression.ExpressionComponentTypes.giCOMPONENT_CUSTOMCALC
					' Not required.
					
				Case modExpression.ExpressionComponentTypes.giCOMPONENT_EXPRESSION
					' Sub-expressions are handled via the Function component class.
					
				Case modExpression.ExpressionComponentTypes.giCOMPONENT_FILTER
					' Load information for filters
					'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.FilterID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.FilterID = prsComponents.Fields("FilterID").Value
					
			End Select
			
		End With
		
TidyUpAndExit: 
		ConstructComponent = fOK
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	Public Function RootExpressionID() As Integer
		' Return the id of the expression which contains this component.
		' NB. We are not returning the id of the immediate parent expression;
		' rather the top-level parent expression. Return 0 if we are unable to
		' determine the root expression.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		Dim lngRootExprID As Integer
		Dim sSQL As String
		Dim objComp As clsExprComponent
		Dim rsExpressions As ADODB.Recordset
		
		sSQL = "SELECT ASRSysExpressions.parentComponentID, ASRSysExpressions.exprID" & " FROM ASRSysExpressions" & " JOIN ASRSysExprComponents ON ASRSysExpressions.exprID = ASRSysExprComponents.exprID" & " WHERE ASRSysExprComponents.componentID = " & Trim(Str(mlngComponentID))
		rsExpressions = datGeneral.GetRecords(sSQL)
		With rsExpressions
			fOK = Not (.EOF And .BOF)
			
			If fOK Then
				' See if the parent expression is a top level expression.
				If .Fields("ParentComponentID").Value = 0 Then
					lngRootExprID = .Fields("ExprID").Value
				Else
					' If the parent expression is not a top-level expression then
					' find the parent expression's parent expression. Confused yet ?
					objComp = New clsExprComponent
					objComp.ComponentID = .Fields("ParentComponentID").Value
					lngRootExprID = objComp.RootExpressionID
					'UPGRADE_NOTE: Object objComp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					objComp = Nothing
				End If
			End If
			
			.Close()
		End With
		
TidyUpAndExit: 
		'UPGRADE_NOTE: Object rsExpressions may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsExpressions = Nothing
		'UPGRADE_NOTE: Object objComp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objComp = Nothing
		If fOK Then
			RootExpressionID = lngRootExprID
		Else
			RootExpressionID = 0
		End If
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
End Class