Option Strict Off
Option Explicit On

Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums

Namespace Expressions
	Public Class clsExprComponent
		Inherits BaseExpressionComponent

		' Component definition variables.
		Private mlngComponentID As Integer
		Private miComponentType As ExpressionComponentTypes

		' Class handling variables.
		Private mobjParentExpression As clsExprExpression
		Private mvComponent As Object

		' Definition for expanded/unexpanded status of the component
		Private mbExpanded As Boolean

		Public Sub New(ByVal Value As SessionInfo)
			MyBase.New(Value)
		End Sub

		Public Function ContainsExpression(plngExprID As Integer) As Boolean
			' Retrun TRUE if the current expression (or any of its sub expressions)
			' contains the given expression. This ensures no cyclic expressions get created.
			Return mvComponent.ContainsExpression(plngExprID)

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
					Case ExpressionComponentTypes.giCOMPONENT_FUNCTION
						'UPGRADE_WARNING: Couldn't resolve default property of object Me.Component.ExpandedNode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Me.Component.ExpandedNode = Value

					Case ExpressionComponentTypes.giCOMPONENT_EXPRESSION
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


		Public Property ComponentType() As ExpressionComponentTypes
			Get
				Return miComponentType
			End Get

			Set(ByVal Value As ExpressionComponentTypes)
				' Set the component type property.
				If miComponentType <> Value Then
					miComponentType = Value

					'UPGRADE_NOTE: Object mvComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					mvComponent = Nothing

					' Instantiate the correct type of component object for
					' the given component type.
					Select Case miComponentType

						Case ExpressionComponentTypes.giCOMPONENT_FIELD
							mvComponent = New clsExprField(SessionInfo)

						Case ExpressionComponentTypes.giCOMPONENT_FUNCTION
							mvComponent = New clsExprFunction(SessionInfo)

						Case ExpressionComponentTypes.giCOMPONENT_CALCULATION
							mvComponent = New clsExprCalculation(SessionInfo)

						Case ExpressionComponentTypes.giCOMPONENT_VALUE
							mvComponent = New clsExprValue(SessionInfo)

						Case ExpressionComponentTypes.giCOMPONENT_OPERATOR
							mvComponent = New clsExprOperator(SessionInfo)

						Case ExpressionComponentTypes.giCOMPONENT_TABLEVALUE
							mvComponent = New clsExprTableLookup(SessionInfo)

						Case ExpressionComponentTypes.giCOMPONENT_PROMPTEDVALUE
							mvComponent = New clsExprPromptedValue(SessionInfo)

						Case ExpressionComponentTypes.giCOMPONENT_CUSTOMCALC
							' Not required.

						Case ExpressionComponentTypes.giCOMPONENT_EXPRESSION
							mvComponent = New clsExprExpression(SessionInfo)

						Case ExpressionComponentTypes.giCOMPONENT_FILTER
							mvComponent = New clsExprFilter(SessionInfo)

					End Select

					If Not mvComponent Is Nothing Then
						'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.BaseComponent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mvComponent.BaseComponent = Me
					End If
				End If

			End Set
		End Property


		Public ReadOnly Property ReturnType() As ExpressionValueTypes
			Get
				' Return the component's return type.
				'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.ReturnType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Return mvComponent.ReturnType

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


		Public Function RuntimeCode(ByRef psRuntimeCode As String, ByRef palngSourceTables(,) As Integer, pfApplyPermissions As Boolean _
																, pfValidating As Boolean, ByRef pavPromptedValues As Object _
																, ByRef psUDFs() As String _
																, Optional plngFixedExprID As Integer = 0, Optional psFixedSQLCode As String = "") As Boolean

			' Return the runtime filter SQL code for the component.
			'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.RuntimeCode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RuntimeCode = mvComponent.RuntimeCode(psRuntimeCode, palngSourceTables, pfApplyPermissions, pfValidating, pavPromptedValues, psUDFs, plngFixedExprID, psFixedSQLCode)

		End Function

		Public Function WriteComponent() As Boolean
			' Write the component definition to the component recordset.

			Dim fOK As Boolean
			Dim lngNewID As Integer

			Try

				' Update the real component expression id property, and give it a unique component id.

				'MH20010712 Need keep manual record of allocated IDs incase users
				'in SYS MGR have created expressions but not yet saved changes
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

			Catch ex As Exception
				Throw

			End Try

			Return True

		End Function

		Public Function NewComponent() As Boolean

			Try

				' Initialize the properties for a new component.
				mlngComponentID = 0
				ComponentType = ExpressionComponentTypes.giCOMPONENT_FIELD
				'	'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.BaseComponent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				mvComponent.BaseComponent = Me
				Return True

			Catch ex As Exception
				Return False

			End Try

		End Function


		Public Function CopyComponent() As clsExprComponent
			' Copies the selected component.
			' When editting a component we actually copy the component first
			' and edit the copy. If the changes are confirmed then the copy
			' replaces the original. If the changes are cancelled then the
			' copy is discarded.

			Dim fOK As Boolean
			Dim objCopyComponent As New clsExprComponent(SessionInfo)

			Try


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

			Catch ex As Exception
				Return Nothing

			End Try

			If fOK Then
				Return objCopyComponent
			Else
				Return Nothing
			End If

		End Function


		'Private Sub InitializeComponent()
		'	' Initialize the properties for a new component.
		'	mlngComponentID = 0
		'	ComponentType = ExpressionComponentTypes.giCOMPONENT_FIELD
		'	'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.BaseComponent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'	mvComponent.BaseComponent = Me

		'End Sub

		Public Function ConstructComponent(ByRef prsComponents As DataRow) As Boolean
			' Read the component definition from the datarow.

			Dim fOK As Boolean = True

			Try

				' Initialise the component with the definition from the database.
				ComponentType = CType(prsComponents("Type"), ExpressionComponentTypes)

				With mvComponent
					Select Case miComponentType
						Case ExpressionComponentTypes.giCOMPONENT_FIELD
							.TableID = prsComponents("fieldTableID")
							.ColumnID = prsComponents("fieldColumnID")
							.FieldPassType = prsComponents("fieldPassBy")
							.SelectionType = prsComponents("fieldSelectionRecord")
							.SelectionLine = prsComponents("fieldSelectionLine")
							.SelectionOrderID = prsComponents("fieldSelectionOrderID")
							.SelectionFilterID = prsComponents("FieldSelectionFilter")

						Case ExpressionComponentTypes.giCOMPONENT_FUNCTION
							.FunctionID = prsComponents("FunctionID")

						Case ExpressionComponentTypes.giCOMPONENT_CALCULATION
							.CalculationID = prsComponents("CalculationID")

						Case ExpressionComponentTypes.giCOMPONENT_VALUE
							.ReturnType = prsComponents("ValueType")
							Select Case CType(prsComponents("ValueType"), ExpressionValueTypes)
								Case ExpressionValueTypes.giEXPRVALUE_CHARACTER
									'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
									.Value = IIf(IsDBNull(prsComponents("valueCharacter")), "", prsComponents("valueCharacter"))
								Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
									'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
									.Value = IIf(IsDBNull(prsComponents("valueNumeric")), 0, prsComponents("valueNumeric"))
								Case ExpressionValueTypes.giEXPRVALUE_LOGIC
									'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
									.Value = IIf(IsDBNull(prsComponents("valueLogic")), True, prsComponents("valueLogic"))
								Case ExpressionValueTypes.giEXPRVALUE_DATE
									.Value = prsComponents("valueDate")
							End Select

						Case ExpressionComponentTypes.giCOMPONENT_OPERATOR
							.OperatorID = prsComponents("OperatorID")

						Case ExpressionComponentTypes.giCOMPONENT_TABLEVALUE
							' Do nothing as Table Value components are treated as Value components.
							.TableID = prsComponents("LookupTableID")
							.ColumnID = prsComponents("LookupColumnID")
							.ReturnType = prsComponents("ValueType")

							Select Case CType(prsComponents("ValueType"), ExpressionValueTypes)
								Case ExpressionValueTypes.giEXPRVALUE_CHARACTER
									'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
									.Value = IIf(IsDBNull(prsComponents("valueCharacter")), "", prsComponents("valueCharacter"))
								Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
									'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
									.Value = IIf(IsDBNull(prsComponents("valueNumeric")), 0, prsComponents("valueNumeric"))
								Case ExpressionValueTypes.giEXPRVALUE_LOGIC
									'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
									.Value = IIf(IsDBNull(prsComponents("valueLogic")), True, prsComponents("valueLogic"))
								Case ExpressionValueTypes.giEXPRVALUE_DATE
									.Value = prsComponents("valueDate")
							End Select

						Case ExpressionComponentTypes.giCOMPONENT_PROMPTEDVALUE
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							.Prompt = IIf(IsDBNull(prsComponents("promptDescription")), "", prsComponents("promptDescription"))
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							.ValueType = IIf(IsDBNull(prsComponents("ValueType")), ExpressionValueTypes.giEXPRVALUE_CHARACTER, prsComponents("ValueType"))
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							.ReturnSize = IIf(IsDBNull(prsComponents("promptSize")), 1, prsComponents("promptSize"))
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							.ReturnDecimals = IIf(IsDBNull(prsComponents("promptDecimals")), 0, prsComponents("promptDecimals"))
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							.ValueFormat = IIf(IsDBNull(prsComponents("promptMask")), "", prsComponents("promptMask"))

							Select Case CType(prsComponents("ValueType"), ExpressionValueTypes)
								Case ExpressionValueTypes.giEXPRVALUE_CHARACTER
									'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
									.DefaultValue = IIf(IsDBNull(prsComponents("valueCharacter")), "", prsComponents("valueCharacter"))
								Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
									'UPGRADE_WARNING: Couldn't resolve default property of object mvComponent.DefaultValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
									.DefaultValue = IIf(IsDBNull(prsComponents("valueNumeric")), 0, prsComponents("valueNumeric"))
								Case ExpressionValueTypes.giEXPRVALUE_LOGIC
									'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
									.DefaultValue = IIf(IsDBNull(prsComponents("valueLogic")), False, prsComponents("valueLogic"))
								Case ExpressionValueTypes.giEXPRVALUE_DATE
									'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
									.DefaultValue = IIf(IsDBNull(prsComponents("valueDate")), Today, prsComponents("valueDate"))
								Case ExpressionValueTypes.giEXPRVALUE_TABLEVALUE
									'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
									.DefaultValue = IIf(IsDBNull(prsComponents("valueCharacter")), "", prsComponents("valueCharacter"))
							End Select
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							.LookupColumn = IIf(IsDBNull(prsComponents("fieldColumnID")), 0, prsComponents("fieldColumnID"))

						Case ExpressionComponentTypes.giCOMPONENT_CUSTOMCALC
							' Not required.

						Case ExpressionComponentTypes.giCOMPONENT_EXPRESSION
							' Sub-expressions are handled via the Function component class.

						Case ExpressionComponentTypes.giCOMPONENT_FILTER
							' Load information for filters
							.FilterID = prsComponents("FilterID")

					End Select

				End With

			Catch ex As Exception
				fOK = False

			End Try

			Return fOK

		End Function

		Public Function RootExpressionID() As Integer
			' Return the id of the expression which contains this component.
			' NB. We are not returning the id of the immediate parent expression;
			' rather the top-level parent expression. Return 0 if we are unable to
			' determine the root expression.			

			Dim fOK As Boolean
			Dim lngRootExprID As Integer
			Dim sSQL As String
			Dim objComp As clsExprComponent
			Dim rsExpressions As DataTable
			Dim objRow As DataRow

			Try

				sSQL = "SELECT ASRSysExpressions.parentComponentID, ASRSysExpressions.exprID FROM ASRSysExpressions JOIN ASRSysExprComponents ON ASRSysExpressions.exprID = ASRSysExprComponents.exprID WHERE ASRSysExprComponents.componentID = " & Trim(Str(mlngComponentID))
				rsExpressions = DB.GetDataTable(sSQL)
				With rsExpressions
					fOK = (.Rows.Count > 0)

					If fOK Then
						objRow = .Rows(0)

						' See if the parent expression is a top level expression.
						If CInt(objRow("ParentComponentID")) = 0 Then
							lngRootExprID = CInt(.Rows(0)("ExprID"))
						Else
							' If the parent expression is not a top-level expression then
							' find the parent expression's parent expression. Confused yet ?
							objComp = New clsExprComponent(SessionInfo)
							objComp.ComponentID = CInt(objRow("ParentComponentID"))
							lngRootExprID = objComp.RootExpressionID
							'UPGRADE_NOTE: Object objComp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
							objComp = Nothing
						End If
					End If

				End With

			Catch ex As Exception
				Return 0

			End Try

			If fOK Then
				Return lngRootExprID
			Else
				Return 0
			End If

		End Function
	End Class
End Namespace