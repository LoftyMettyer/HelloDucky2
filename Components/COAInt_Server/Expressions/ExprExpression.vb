Option Strict Off
Option Explicit On

Imports System.Collections.Generic
Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Metadata
Imports HR.Intranet.Server.Structures
Imports System.Data.SqlClient

Namespace Expressions
	Public Class clsExprExpression
		Inherits BaseExpressionComponent

		' Expression definition variables.
		Private mlngExpressionID As Integer
		Private msExpressionName As String
		Private mlngBaseTableID As Integer
		Private miReturnType As ExpressionValueTypes
		Private miExpressionType As ExpressionTypes
		Private mlngParentComponentID As Integer
    Private mlngSecondTableID as Integer
		Private msOwner As String
		Private msAccess As String
		Private msDescription As String
		Private mlngTimeStamp As Integer
		Private msBaseTableName As String
		Private mbViewInColour As Boolean
		Private mbExpandedNode As Boolean

		Public mfDontUpdateTimeStamp As Boolean

		' Class handling variables.
		Private mfConstructed As Boolean
		Private mcolComponents As Collection
		Private mobjBadComponent As clsExprComponent
		Private mobjBaseComponent As clsExprComponent

		Private msErrorMessage As String

		' Array holding the User Defined functions that are needed for this expression
		Private mastrUDFsRequired() As String
    
	Public Property SecondTableID() As Integer
		Get
			SecondTableID = mlngSecondTableID		
		End Get
		Set(ByVal Value As Integer)		
      mlngSecondTableID = value
		End Set
	End Property

		Public ReadOnly Property ComponentDescription() As String
			Get
				Return msExpressionName
			End Get
		End Property

		Public Property ExpressionID() As Integer
			Get
				' Return the expression ID.
				ExpressionID = mlngExpressionID

			End Get
			Set(ByVal Value As Integer)
				' Set the expression ID.
				If mlngExpressionID <> Value Then
					mlngExpressionID = Value
					mfConstructed = False
				End If

			End Set
		End Property

		Public Property BaseTableID() As Integer
			Get
				' Return the expressions base table ID.
				BaseTableID = mlngBaseTableID

			End Get
			Set(ByVal Value As Integer)
				' Set the expression base table property.
				If mlngBaseTableID <> Value Then
					mlngBaseTableID = Value
					msBaseTableName = Tables.GetById(mlngBaseTableID).Name
				End If
			End Set
		End Property

		Public Property ReturnType() As ExpressionValueTypes
			Get
				Return miReturnType
			End Get
			Set(ByVal Value As ExpressionValueTypes)
				miReturnType = Value
			End Set
		End Property

		Public Property ExpressionType() As ExpressionTypes
			Get
				' Return the expression's parent type property.
				Return miExpressionType

			End Get
			Set(ByVal Value As ExpressionTypes)
				' Set the expression's type property.
				miExpressionType = Value

			End Set
		End Property

		Public Property Name() As String
			Get
				' Return the expression name.
				If Not mfConstructed Then
					ConstructExpression()
				End If

				Name = msExpressionName

			End Get
			Set(ByVal Value As String)
				' Set the expression name.
				msExpressionName = Value

			End Set
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

		Public ReadOnly Property ErrorMessage() As String
			Get
				Return msErrorMessage
			End Get
		End Property

		Public WriteOnly Property ParentComponentID() As Integer
			Set(ByVal Value As Integer)
				' Set the Parent component ID.
				mlngParentComponentID = Value

			End Set
		End Property

		Public ReadOnly Property ComponentType() As ExpressionComponentTypes
			Get
				Return ExpressionComponentTypes.giCOMPONENT_EXPRESSION
			End Get
		End Property

		Public Property Components() As Collection
			Get
				' Return the component collection.
				Components = mcolComponents

			End Get
			Set(ByVal Value As Collection)
				' Set the component collection.
				mcolComponents = Value

			End Set
		End Property


		Public Property Owner() As String
			Get
				' Return the expression owner.
				Owner = msOwner

			End Get
			Set(ByVal Value As String)
				' Set the expression owner.
				msOwner = Value

			End Set
		End Property

		Public ReadOnly Property BadComponent() As clsExprComponent
			Get
				' Return the component last caused the expression to fail its validity check.
				BadComponent = mobjBadComponent

			End Get
		End Property

		Public Property Access() As String
			Get
				' Return the access code.
				Access = msAccess

			End Get
			Set(ByVal Value As String)
				' Set the access code.
				msAccess = Value

			End Set
		End Property

		Public Property Description() As String
			Get
				' Return the expression's description.
				Description = msDescription

			End Get
			Set(ByVal Value As String)
				' Set the expression's descriprion property.
				msDescription = Value

			End Set
		End Property

		Public Property Timestamp() As Integer
			Get
				' Return the expression's timestamp value.
				Timestamp = mlngTimeStamp

			End Get
			Set(ByVal Value As Integer)
				' Set the expression's timestamp property.
				mlngTimeStamp = Value

			End Set
		End Property

		Public Property BaseTableName() As String
			Get
				' Return the name of the expression's base table.
				BaseTableName = msBaseTableName

			End Get
			Set(ByVal Value As String)
				' Set the name of the expression's base table.
				msBaseTableName = Value

			End Set
		End Property

		Public Property ViewInColour() As Boolean
			Get

				ViewInColour = mbViewInColour

			End Get
			Set(ByVal Value As Boolean)

				mbViewInColour = Value

			End Set
		End Property

		Public Property ExpandedNode() As Boolean
			Get

				ExpandedNode = mbExpandedNode

			End Get
			Set(ByVal Value As Boolean)

				mbExpandedNode = Value

			End Set
		End Property

		Public WriteOnly Property Constructed() as Boolean
    	Set(ByVal Value As Boolean)
				mfConstructed = Value
			End Set
		End Property

		Public Sub New(ByVal Value As SessionInfo)

			MyBase.New(Value)

			' Create a new collection to hold the expression's components.
			mcolComponents = New Collection
			mfConstructed = False
			mbExpandedNode = False
			ReDim mastrUDFsRequired(0)

		End Sub



		'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Private Sub Class_Terminate_Renamed()
			' Disassociate object variables.
			'UPGRADE_NOTE: Object mcolComponents may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			mcolComponents = Nothing
			'UPGRADE_NOTE: Object mobjBadComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			mobjBadComponent = Nothing
			'UPGRADE_NOTE: Object mobjBaseComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			mobjBaseComponent = Nothing

		End Sub
		Protected Overrides Sub Finalize()
			Class_Terminate_Renamed()
			MyBase.Finalize()
		End Sub

		Public Function DeleteComponent(ByRef pobjComponent As clsExprComponent) As Boolean
			' Remove the given component from the expression.

			Dim iLoop As Integer
			Dim iIndex As Integer = 0

			Try
				' Find the given component in the component collection.
				For iLoop = 1 To mcolComponents.Count()
					If pobjComponent Is mcolComponents.Item(iLoop) Then
						iIndex = iLoop
						Exit For
					End If
				Next iLoop

				' Delete the current component if it has been found.
				If iIndex > 0 Then
					mcolComponents.Remove(iIndex)
				End If

			Catch ex As Exception
				Return False

			End Try

			Return True

		End Function

    Public Function AddComponent() As clsExprComponent

      ' Instantiate a component object.
			Dim objComponent = New clsExprComponent(SessionInfo)
      objComponent.NewComponent
      Return AddComponent(objComponent)

    End Function

		Public Function AddComponent(objComponent As clsExprComponent) As clsExprComponent
			' Add a new component to the expression.
			' Returns the new component object.

			Try

				' Initialse the new component's properties.
				objComponent.ParentExpression = Me
        objComponent.Component.BaseComponent = objComponent
				mcolComponents.Add(objComponent)

			Catch ex As Exception
				Return Nothing

			End Try

  		Return objComponent

		End Function

		Public Function SelectExpression(ByRef pfLockTable As Boolean, Optional ByRef plngOptions As Integer = 0) As Boolean
		End Function

		Public Function CopyComponent() As clsExprExpression
		End Function

		Public Function DeleteExpression() As Boolean
		End Function

		Public Function ValidityMessage(piInvalidityCode As ExprValidationCodes) As String
			' Return the text nmessage that describes the given expression invalidity code.

			Select Case piInvalidityCode

				Case ExprValidationCodes.giEXPRVALIDATION_NOERRORS
					ValidityMessage = "No errors."

				Case ExprValidationCodes.giEXPRVALIDATION_MISSINGOPERAND
					ValidityMessage = "Missing operand."

				Case ExprValidationCodes.giEXPRVALIDATION_SYNTAXERROR
					ValidityMessage = "Syntax error."

				Case ExprValidationCodes.giEXPRVALIDATION_EXPRTYPEMISMATCH
					ValidityMessage = "Return type mismatch."

				Case ExprValidationCodes.giEXPRVALIDATION_UNKNOWNERROR
					ValidityMessage = "Unknown error."

				Case ExprValidationCodes.giEXPRVALIDATION_OPERANDTYPEMISMATCH
					ValidityMessage = "Operand type mismatch."

				Case ExprValidationCodes.giEXPRVALIDATION_PARAMETERTYPEMISMATCH
					ValidityMessage = "Parameter type mismatch."

				Case ExprValidationCodes.giEXPRVALIDATION_NOCOMPONENTS
					ValidityMessage = "The " & LCase(ExpressionTypeName(miExpressionType)) & " must have at least one component."

				Case ExprValidationCodes.giEXPRVALIDATION_PARAMETERSYNTAXERROR
					ValidityMessage = "Function parameter syntax error."

				Case ExprValidationCodes.giEXPRVALIDATION_PARAMETERNOCOMPONENTS
					ValidityMessage = "The function parameter expression must have at least one component."

				Case ExprValidationCodes.giEXPRVALIDATION_FILTEREVALUATION
					ValidityMessage = "Error creating SQL runtime code."

				Case ExprValidationCodes.giEXPRVALIDATION_SQLERROR
					ValidityMessage = "The complexity of the " & LCase(ExpressionTypeName(miExpressionType)) & " has caused the following SQL error : " & vbNewLine & vbNewLine & "'" & msErrorMessage & "'" & vbNewLine & vbNewLine & "Try simplifying the " & LCase(ExpressionTypeName(miExpressionType)) & "."

				Case ExprValidationCodes.giEXPRVALIDATION_ASSOCSQLERROR
					ValidityMessage = "The complexity of this " & LCase(ExpressionTypeName(miExpressionType)) & " would cause an expression that uses this " & LCase(ExpressionTypeName(miExpressionType)) & " to suffer from the following SQL error : " & vbNewLine & vbNewLine & "'" & msErrorMessage & "'" & vbNewLine & vbNewLine & "Try simplifying this " & LCase(ExpressionTypeName(miExpressionType)) & "."

				Case ExprValidationCodes.giEXPRVALIDATION_CYCLIC
					ValidityMessage = "Invalid definition due to cyclic reference."

				Case Else
					ValidityMessage = "The function parameter expression must have at least one component."

			End Select

		End Function

		Public Function ContainsExpression(plngExprID As Integer) As Boolean
			' Retrun TRUE if the current expression (or any of its sub expressions)
			' contains the given expression. This ensures no cyclic expressions get created.
			'JPD 20040507 Fault 8600

			Dim iLoop1 As Integer
			Dim bContainsExpression = False

			Try

				For iLoop1 = 1 To mcolComponents.Count()
					If bContainsExpression Then
						Exit For
					End If

					With mcolComponents.Item(iLoop1)
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().ContainsExpression. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						bContainsExpression = .ContainsExpression(plngExprID)
					End With
				Next iLoop1

			Catch ex As Exception
				Return True

			End Try

			Return bContainsExpression

		End Function

		Public Function WriteExpression() As Boolean
			' Write the expression definition to the database.

			Dim bOK As Boolean
			Dim objComponent As clsExprComponent

			Try

				Dim prmID = New SqlParameter("ExpressionID", SqlDbType.Int) With {.Direction = ParameterDirection.InputOutput, .Value = mlngExpressionID}

				DB.ExecuteSP("spASRIntSaveExpression", _
					prmID,
					New SqlParameter("Name", SqlDbType.VarChar, 255) With {.Value = msExpressionName}, _
					New SqlParameter("TableID", SqlDbType.Int) With {.Value = mlngBaseTableID}, _
					New SqlParameter("returnType", SqlDbType.Int) With {.Value = IIf(miExpressionType = ExpressionTypes.giEXPR_RUNTIMECALCULATION, ExpressionValueTypes.giEXPRVALUE_UNDEFINED, miReturnType)}, _
					New SqlParameter("returnSize", SqlDbType.Int) With {.Value = 0}, _
					New SqlParameter("returnDecimals", SqlDbType.Int) With {.Value = 0}, _
					New SqlParameter("type", SqlDbType.Int) With {.Value = miExpressionType}, _
					New SqlParameter("parentComponentID", SqlDbType.Int) With {.Value = mlngParentComponentID}, _
					New SqlParameter("Username", SqlDbType.VarChar, 50) With {.Value = msOwner}, _
					New SqlParameter("access", SqlDbType.VarChar, 2) With {.Value = msAccess}, _
					New SqlParameter("description", SqlDbType.VarChar, 255) With {.Value = msDescription})

				mlngExpressionID = CInt(prmID.Value)

				' Delete the expression's existing components from the database.
				bOK = DeleteExistingComponents()

				If bOK Then
					' Add any components for this expression.
					For Each objComponent In mcolComponents
						objComponent.ParentExpression = Me
						bOK = objComponent.WriteComponent

						If Not bOK Then
							Exit For
						End If
					Next objComponent
				End If

			Catch ex As Exception
				Return False

			End Try

			Return True

		End Function

		Public Function RuntimeCode(ByRef psRuntimeCode As String, ByRef palngSourceTables(,) As Integer, ByRef pfApplyPermissions As Boolean _
																, ByRef pfValidating As Boolean, ByRef pavPromptedValues As Object _
																, ByRef psUDFs() As String _
																, Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean

			' Return the SQL code that defines the expression.
			' Used when creating the 'where clause' for view definitions.

			Dim fOK As Boolean = True
			Dim iLoop1 As Integer
			Dim iLoop2 As Integer
			Dim iLoop3 As Integer
			Dim iParameter1Index As Integer
			Dim iParameter2Index As Integer
			Dim iMinOperatorPrecedence As Integer = -1
			Dim iMaxOperatorPrecedence As Integer = -1
			Dim sCode As String = ""
			Dim sComponentCode As String
			Dim vParameter1 As Object
			Dim vParameter2 As Object
			Dim avValues(,) As Object

			Try

				' Create an array of the components in the expression.
				' Column 1 = operator id.
				' Column 2 = component where clause code.
				ReDim avValues(2, mcolComponents.Count())
				For iLoop1 = 1 To mcolComponents.Count()
					With mcolComponents.Item(iLoop1)
						.SessionInfo = SessionInfo

						' If the current component is an operator then read the operator id into the array.
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop1).ComponentType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If .ComponentType = ExpressionComponentTypes.giCOMPONENT_OPERATOR Then
							'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object avValues(1, iLoop1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avValues(1, iLoop1) = .Component.OperatorID
							'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop1).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							iMinOperatorPrecedence = IIf(iMinOperatorPrecedence > .Component.Precedence Or iMinOperatorPrecedence = -1, .Component.Precedence, iMinOperatorPrecedence)
							'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop1).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							iMaxOperatorPrecedence = IIf(iMaxOperatorPrecedence < .Component.Precedence Or iMaxOperatorPrecedence = -1, .Component.Precedence, iMaxOperatorPrecedence)
						End If

						' JPD20020419 Fault 3687
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().RuntimeCode. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						fOK = .RuntimeCode(sComponentCode, palngSourceTables, pfApplyPermissions, pfValidating, pavPromptedValues, psUDFs, plngFixedExprID, psFixedSQLCode)

						If fOK Then
							'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avValues(2, iLoop1) = sComponentCode
						End If
					End With

					If Not fOK Then
						Exit For
					End If
				Next iLoop1

				If fOK Then
					' Loop throught the expression's components checking that they are valid.
					' Evaluate operators in the correct order.
					For iLoop1 = iMinOperatorPrecedence To iMaxOperatorPrecedence
						For iLoop2 = 1 To mcolComponents.Count()
							With mcolComponents.Item(iLoop2)
								.SessionInfo = SessionInfo
								'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).ComponentType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If .ComponentType = ExpressionComponentTypes.giCOMPONENT_OPERATOR Then
									'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If .Component.Precedence = iLoop1 Then
										' Check that the operator has the correct parameter types.
										' Read the value that follows the current operator.
										iParameter1Index = 0
										iParameter2Index = 0

										' Read the index of the first parameter.
										For iLoop3 = iLoop2 + 1 To UBound(avValues, 2)
											'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											If avValues(2, iLoop3) <> vbNullString Then
												iParameter1Index = iLoop3
												Exit For
											End If
										Next iLoop3

										' If a parameter has been found then read its value.
										' Otherwise the expression is invalid.
										If iParameter1Index > 0 Then
											'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iParameter1Index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											vParameter1 = avValues(2, iParameter1Index)
										End If

										' Read a second parameter if required.
										'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										If (.Component.OperandCount = 2) Then
											'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											'UPGRADE_WARNING: Couldn't resolve default property of object vParameter2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											vParameter2 = vParameter1
											iParameter2Index = iParameter1Index
											iParameter1Index = 0

											' Read the index of the parameter's value if there is one.
											For iLoop3 = iLoop2 - 1 To 1 Step -1
												'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												If avValues(2, iLoop3) <> vbNullString Then
													iParameter1Index = iLoop3
													Exit For
												End If
											Next iLoop3

											' If a parameter has been found then read its value.
											' Otherwise the expression is invalid.
											If iParameter1Index > 0 Then
												'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iParameter1Index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												vParameter1 = avValues(2, iParameter1Index)

												' JPD20020415 Fault 3662 - Need to cast values as float for division operators
												'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												If .Component.CastAsFloat Then
													'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													vParameter1 = "Cast(" & vParameter1 & " As Float)"
												End If

											End If

											' Update the array to reflect the constructed SQL code.
											'UPGRADE_WARNING: Couldn't resolve default property of object avValues(1, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											avValues(1, iLoop2) = vbNullString
											'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											If .Component.SQLType = "O" Then
												'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												If .Component.CheckDivideByZero Then
													' JPD20020415 Fault 3638
													'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													Select Case .Component.OperatorID
														Case 16	'Modulus
															'UPGRADE_WARNING: Couldn't resolve default property of object vParameter2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
															'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
															'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
															avValues(2, iLoop2) = "(case when " & vParameter2 & " = 0 then 0 else (" & vbNewLine & vParameter1 & " - (CAST((" & vParameter1 & " / " & vParameter2 & ") AS INT) * " & vParameter2 & ")" & vbNewLine & ") end)"
														Case Else
															'UPGRADE_WARNING: Couldn't resolve default property of object vParameter2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
															'UPGRADE_WARNING: Couldn't resolve default property of object avValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
															'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
															'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
															avValues(2, iLoop2) = "(case when " & vParameter2 & " = 0 then 0 else (" & vbNewLine & vParameter1 & " " & vbNewLine & avValues(2, iLoop2) & " " & vbNewLine & vParameter2 & vbNewLine & ") end)"
													End Select
												Else
													'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
													Select Case .Component.OperatorID
														Case 5 'And
															'UPGRADE_WARNING: Couldn't resolve default property of object vParameter2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
															'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
															'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
															avValues(2, iLoop2) = "(CASE WHEN (" & vParameter1 & " = 1) AND (" & vParameter2 & " = 1) THEN 1 ELSE 0 END)"
														Case 6 'Or
															'UPGRADE_WARNING: Couldn't resolve default property of object vParameter2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
															'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
															'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
															avValues(2, iLoop2) = "(CASE WHEN (" & vParameter1 & " = 1) OR (" & vParameter2 & " = 1) THEN 1 ELSE 0 END)"
														Case Else
															'UPGRADE_WARNING: Couldn't resolve default property of object vParameter2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
															'UPGRADE_WARNING: Couldn't resolve default property of object avValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
															'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
															'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
															avValues(2, iLoop2) = "(" & vbNewLine & vParameter1 & " " & vbNewLine & avValues(2, iLoop2) & " " & vbNewLine & vParameter2 & vbNewLine & ")"
													End Select
												End If
											Else
												'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												'UPGRADE_WARNING: Couldn't resolve default property of object vParameter2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												avValues(2, iLoop2) = IIf(Len(.Component.SQLFixedParam1) > 0, "(", "") & avValues(2, iLoop2) & vbNewLine & "(" & vbNewLine & vParameter1 & vbNewLine & ", " & vbNewLine & vParameter2 & vbNewLine & ")" & IIf(Len(.Component.SQLFixedParam1) > 0, " " & .Component.SQLFixedParam1 & ")", "")
											End If
											'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iParameter1Index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											avValues(2, iParameter1Index) = vbNullString
											'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iParameter2Index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											avValues(2, iParameter2Index) = vbNullString
										Else
											' Update the array to reflect the constructed SQL code.
											'UPGRADE_WARNING: Couldn't resolve default property of object avValues(1, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											avValues(1, iLoop2) = vbNullString
											'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											If .Component.SQLType = "O" Then
												'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												Select Case .Component.OperatorID
													Case 13	'Not
														'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														avValues(2, iLoop2) = "(CASE WHEN " & vParameter1 & " = 1 THEN 0 ELSE 1 END)"
													Case Else
														'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														'UPGRADE_WARNING: Couldn't resolve default property of object avValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
														avValues(2, iLoop2) = "(" & vbNewLine & avValues(2, iLoop2) & " " & vbNewLine & vParameter1 & vbNewLine & ")"
												End Select
											Else
												'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												'UPGRADE_WARNING: Couldn't resolve default property of object vParameter1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												avValues(2, iLoop2) = IIf(Len(.Component.SQLFixedParam1) > 0, "(", "") & avValues(2, iLoop2) & vbNewLine & "(" & vbNewLine & vParameter1 & vbNewLine & ")" & IIf(Len(.Component.SQLFixedParam1) > 0, " " & .Component.SQLFixedParam1 & ")", "")
											End If
											'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iParameter1Index). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
											avValues(2, iParameter1Index) = vbNullString
										End If

										If (miExpressionType = ExpressionTypes.giEXPR_RUNTIMECALCULATION) _
                      Or (miExpressionType = ExpressionTypes.giEXPR_RUNTIMEFILTER) _
                      Or (miExpressionType = ExpressionTypes.giEXPR_MATCHJOINEXPRESSION) _
                      Or (miExpressionType = ExpressionTypes.giEXPR_MATCHWHEREEXPRESSION) _
                      Or (miExpressionType = ExpressionTypes.giEXPR_MATCHSCOREEXPRESSION) _
                      Or (miExpressionType = ExpressionTypes.giEXPR_LINKFILTER) Then

											If (.Component.ReturnType = ExpressionValueTypes.giEXPRVALUE_LOGIC) And ((.Component.OperatorID <> 5) And (.Component.OperatorID <> 6) And (.Component.OperatorID <> 13)) Then
												avValues(2, iLoop2) = "(CASE WHEN (" & avValues(2, iLoop2) & ") THEN 1 ELSE 0 END)"
											End If
										End If
									End If
								End If
							End With
						Next iLoop2
					Next iLoop1

					For iLoop1 = 1 To UBound(avValues, 2)
						'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If avValues(2, iLoop1) <> vbNullString Then
							'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sCode = avValues(2, iLoop1).ToString()
							Exit For
						End If
					Next iLoop1

				End If

			Catch ex As Exception
				fOK = False

			Finally

				If fOK Then
					psRuntimeCode = sCode
				Else
					psRuntimeCode = ""
				End If

			End Try

			Return fOK

		End Function



		Public Function RuntimeFilterCode(ByRef psFilterCode As String, ByRef pfApplyPermissions As Boolean _
																			, ByRef psUDFs() As String _
																			, Optional ByRef pfValidating As Boolean = False, Optional ByRef pavPromptedValues As Object = Nothing, Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean
			' Return TRUE if the filter code was created okay.
			' Return the runtime filter SQL code in the parameter 'psFilterCode'.
			' Apply permissions to the filter code only if the 'pfApplyPermissions' parameter is TRUE.
			' The filter code is to be used to validate the expression if the 'pfValidating' parameter is TRUE.
			' This is used to suppress prompting the user for promted values, when we are only validating the expression.

			Dim fOK As Boolean
			Dim iLoop1 As Integer
			Dim sWhereCode As String
			Dim sBaseTableSource As String
			Dim sRuntimeFilterSQL As String
			Dim alngSourceTables(,) As Integer
			Dim objTableView As TablePrivilege
			Dim listRelatedTables As New List(Of TableRelation)
			Dim objTableRelation As TableRelation

			Try

				' Check if the 'validating' parameter is set.
				' If not, set it to FALSE.
				'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
				If IsNothing(pfValidating) Then
					pfValidating = False
				End If

				' Construct the expression from the database definition.
				fOK = ConstructExpression()

				If fOK Then
					sBaseTableSource = msBaseTableName
					If pfApplyPermissions Then
						' Get the 'realSource' of the table.
						objTableView = gcoTablePrivileges.Item(msBaseTableName)
						If objTableView.TableType = TableTypes.tabChild Then
							sBaseTableSource = objTableView.RealSource
						End If
						'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						objTableView = Nothing
					End If

					sRuntimeFilterSQL = "SELECT DISTINCT " & sBaseTableSource & ".id FROM " & sBaseTableSource & " " & vbNewLine

					' Create an array of the IDs of the tables/view referred to in the expression.
					' This is used for joining all of the tables/views used.
					' Column 1 = 0 if this row is for a table, 1 if it is for a view.
					' Column 2 = table/view ID.
					ReDim alngSourceTables(2, 0)

					' Get the filter code.
					fOK = RuntimeCode(sWhereCode, alngSourceTables, pfApplyPermissions, pfValidating, pavPromptedValues, psUDFs, plngFixedExprID, psFixedSQLCode)
				End If

				If fOK Then

           If miExpressionType = ExpressionTypes.giEXPR_MATCHJOINEXPRESSION Or _
             miExpressionType = ExpressionTypes.giEXPR_MATCHWHEREEXPRESSION Or _
             miExpressionType = ExpressionTypes.giEXPR_MATCHSCOREEXPRESSION Then

            sRuntimeFilterSQL = sWhereCode
    
          Else
					  ' Create an array of the tables related to the expression base table.
					  ' Used when Joining any other tables/view used.
					  ' Column 1 = 'parent' if the expression's base table is the parent of the other table
					  '            'child' if the expression's base table is the child of the other table
					  ' Column 2 = ID of the other table

					  For Each objRelation In Relations.FindAll(Function(n) n.ParentID = mlngBaseTableID)
						  objTableRelation = New TableRelation
						  objTableRelation.RelationType = RelationType.Parent
						  objTableRelation.TableID = objRelation.ChildID
						  listRelatedTables.Add(objTableRelation)
					  Next

					  For Each objRelation In Relations.FindAll(Function(n) n.ChildID = mlngBaseTableID)
						  objTableRelation = New TableRelation
						  objTableRelation.RelationType = RelationType.Child
						  objTableRelation.TableID = objRelation.ParentID
						  listRelatedTables.Add(objTableRelation)
					  Next


					  ' Join any other tables/views used.
					  For iLoop1 = 1 To UBound(alngSourceTables, 2)
						  If alngSourceTables(1, iLoop1) = 0 Then
							  objTableView = gcoTablePrivileges.FindTableID(alngSourceTables(2, iLoop1))
						  Else
							  objTableView = gcoTablePrivileges.FindViewID(alngSourceTables(2, iLoop1))
						  End If

						  If objTableView.TableID = mlngBaseTableID Then
							  ' Join a view on the base table.
							  If Not pfApplyPermissions Then
								  sRuntimeFilterSQL = sRuntimeFilterSQL & "LEFT OUTER JOIN " & objTableView.TableName & " ON " & sBaseTableSource & ".id = " & objTableView.TableName & ".id" & vbNewLine
							  Else
								  sRuntimeFilterSQL = sRuntimeFilterSQL & "LEFT OUTER JOIN " & objTableView.RealSource & " ON " & sBaseTableSource & ".id = " & objTableView.RealSource & ".id" & vbNewLine
							  End If
						  Else
							  ' Join a table/view on a parent/child related to the base table.
							  For Each objTableRelation In listRelatedTables.FindAll(Function(n) n.TableID = objTableView.TableID)


								  If Not pfApplyPermissions Then
									  'UPGRADE_WARNING: Couldn't resolve default property of object avRelatedTables(1, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									  If objTableRelation.RelationType = RelationType.Parent Then
										  sRuntimeFilterSQL = sRuntimeFilterSQL & "LEFT OUTER JOIN " & objTableView.TableName & " ON " & sBaseTableSource & ".id = " & objTableView.TableName & ".id_" & Trim(Str(mlngBaseTableID)) & " " & vbNewLine
									  Else
										  'UPGRADE_WARNING: Couldn't resolve default property of object avRelatedTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										  sRuntimeFilterSQL = sRuntimeFilterSQL & "LEFT OUTER JOIN " & objTableView.TableName & " ON " & sBaseTableSource & ".id_" & Trim(Str(objTableRelation.TableID)) & " = " & objTableView.TableName & ".id " & vbNewLine
									  End If
								  Else
									  'UPGRADE_WARNING: Couldn't resolve default property of object avRelatedTables(1, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									  If objTableRelation.RelationType = RelationType.Parent Then
										  sRuntimeFilterSQL = sRuntimeFilterSQL & "LEFT OUTER JOIN " & objTableView.RealSource & " ON " & sBaseTableSource & ".id = " & objTableView.RealSource & ".id_" & Trim(Str(mlngBaseTableID)) & " " & vbNewLine
									  Else
										  'UPGRADE_WARNING: Couldn't resolve default property of object avRelatedTables(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										  sRuntimeFilterSQL = sRuntimeFilterSQL & "LEFT OUTER JOIN " & objTableView.RealSource & " ON " & sBaseTableSource & ".id_" & Trim(Str(objTableRelation.TableID)) & " = " & objTableView.RealSource & ".id " & vbNewLine
									  End If
								  End If

								  Exit For

							  Next
						  End If

						  'UPGRADE_NOTE: Object objTableView may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						  objTableView = Nothing
					  Next iLoop1

					  ' Add the filter 'where' clause code.
					  If Len(sWhereCode) > 0 Then
						  sWhereCode = sWhereCode & " = 1"

						  sRuntimeFilterSQL = sRuntimeFilterSQL & "WHERE " & vbNewLine & sWhereCode & vbNewLine
					  End If
				  End If

        End If

			Catch ex As Exception
				fOK = False

			End Try


			If fOK Then
				psFilterCode = sRuntimeFilterSQL
			Else
				psFilterCode = ""
			End If

			Return fOK


		End Function

		Public Function RuntimeCalculationCode(ByRef palngSourceTables(,) As Integer, ByRef psCalcCode As String, ByRef pastrUDFsRequired() As String _
																					 , ByRef pfApplyPermissions As Boolean _
																					 , Optional ByRef pfValidating As Boolean = False, Optional ByRef pavPromptedValues As Object = Nothing _
																					 , Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean
			' Return TRUE if the Calculation code was created okay.
			' Return the runtime Calculation SQL code in the parameter 'psCalcCode'.
			' Apply permissions to the Calculation code only if the 'pfApplyPermissions' parameter is TRUE.

			Dim fOK As Boolean
			Dim sRuntimeSQL As String
			Dim avDummyPrompts(,) As Object

			Try

				' Check if the 'validating' parameter is set.
				' If not, set it to FALSE.
				'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
				If IsNothing(pfValidating) Then
					pfValidating = False
				End If

				' Construct the expression from the database definition.
				fOK = ConstructExpression()

				If fOK Then
					' Get the Calculation code.
					'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
					If IsNothing(pavPromptedValues) Then
						ReDim avDummyPrompts(1, 0)
						fOK = RuntimeCode(sRuntimeSQL, palngSourceTables, pfApplyPermissions, pfValidating, avDummyPrompts, pastrUDFsRequired, plngFixedExprID, psFixedSQLCode)
					Else
						fOK = RuntimeCode(sRuntimeSQL, palngSourceTables, pfApplyPermissions, pfValidating, pavPromptedValues, pastrUDFsRequired, plngFixedExprID, psFixedSQLCode)
					End If
				End If

				If fOK Then
					If pfApplyPermissions Then
						fOK = (ValidateExpression(True) = ExprValidationCodes.giEXPRVALIDATION_NOERRORS)
					End If
				End If

				If fOK And (miReturnType = ExpressionValueTypes.giEXPRVALUE_LOGIC) Then
					sRuntimeSQL = "convert(bit, " & sRuntimeSQL & ")"
				End If

			Catch ex As Exception
				fOK = False

			Finally
				If fOK Then
					psCalcCode = sRuntimeSQL
				Else
					psCalcCode = ""
				End If

			End Try

			Return fOK

		End Function

		Friend Function DeleteExistingComponents() As Boolean
			' Delete the expression's components and sub-expression's
			' (ie. function parameter expressions) from the database.

			Dim fOK As Boolean = True
			Dim sSQL As String
			Dim sDeletedExpressionIDs As String = ""
			Dim rsSubExpressions As DataTable
			Dim objExpr As clsExprExpression

			Try

				' Get the expression's function components from the database.
				sSQL = "SELECT ASRSysExpressions.exprID FROM ASRSysExpressions INNER JOIN ASRSysExprComponents ON ASRSysExpressions.parentComponentID = ASRSysExprComponents.componentID AND ASRSysExprComponents.exprID = " & Trim(Str(mlngExpressionID))
				rsSubExpressions = DB.GetDataTable(sSQL)
				With rsSubExpressions
					For Each objRow As DataRow In .Rows
						If Not fOK Then Exit For

						' Instantiate each function parameter expression.
						' Instruct the function parameter expression to delete its components.
						objExpr = New clsExprExpression(SessionInfo)
						objExpr.ExpressionID = CInt(objRow("ExprID"))
						fOK = objExpr.DeleteExistingComponents
						'UPGRADE_NOTE: Object objExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						objExpr = Nothing

						' Add the ID of the sub-expression to the string of sub-expressions to be deleted.
						sDeletedExpressionIDs = sDeletedExpressionIDs & IIf(Len(sDeletedExpressionIDs) > 0, ", ", "") & Trim(Str(objRow("ExprID")))

					Next
				End With
				'UPGRADE_NOTE: Object rsSubExpressions may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rsSubExpressions = Nothing

				If Len(sDeletedExpressionIDs) > 0 Then
					' Delete all existing sub-expressions for this expression from the database.
					sSQL = "DELETE FROM ASRSysExpressions WHERE exprID IN (" & sDeletedExpressionIDs & ")"
					DB.ExecuteSql(sSQL)
				End If

				' Delete all existing components for this expression from the database.
				sSQL = "DELETE FROM ASRSysExprComponents WHERE exprID = " & Trim(Str(mlngExpressionID))
				DB.ExecuteSql(sSQL)

			Catch ex As Exception
				Return False

			End Try

			Return fOK

		End Function

		Public Function ValidateExpression(ByRef pfTopLevel As Boolean) As ExprValidationCodes
			' Validate the expression. Return a code defining the validity of the expression.
			' NB. This function is also good for evaluating the return type of an expression
			' which has definite return type (eg. function sub-expressions, runtime calcs, etc).

			Dim iLoop1 As Integer
			Dim iLoop2 As Integer
			Dim iLoop3 As Integer
			Dim iParam1Type As Short
			Dim iParam2Type As Short
			Dim iParameter1Index As Integer
			Dim iParameter2Index As Integer
			Dim iParam1ReturnType As Short
			Dim iParam2ReturnType As Short
			Dim iOperatorReturnType As ExpressionValueTypes
			Dim iBadLogicColumnIndex As Integer
			Dim iMinOperatorPrecedence As Short
			Dim iMaxOperatorPrecedence As Short
			Dim iValidationCode As ExprValidationCodes
			Dim iEvaluatedReturnType As ExpressionValueTypes
			Dim aiDummyValues(,) As Integer
			Dim avDummyPrompts(,) As Object
			Dim iTempReturnType As Integer

			Try

				ReDim avDummyPrompts(1, 0)

				' Initialise variables.
				iBadLogicColumnIndex = 0
				iMinOperatorPrecedence = -1
				iMaxOperatorPrecedence = -1
				iValidationCode = ExprValidationCodes.giEXPRVALIDATION_NOERRORS
				'UPGRADE_NOTE: Object mobjBadComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				mobjBadComponent = Nothing

				' If there are no expression components then report the error.
				If mcolComponents.Count() = 0 Then
					iValidationCode = ExprValidationCodes.giEXPRVALIDATION_NOCOMPONENTS
				End If

				' Create an array of the component return types and operator ids.
				' Index 1 = operator id.
				' Index 2 = return type.
				'
				' Eg. the expression
				' 'abc'
				' Concatenated with
				' Function 'uppercase'
				'   <parameter>
				'      Field 'personnel.surname'
				'
				' will be represented in the array as
				' null,  giEXPRVALUE_CHARACTER
				'   17,  giEXPRVALUE_CHARACTER
				' null,  giEXPRVALUE_CHARACTER
				'
				' The operators are then evaluated to leave the array as :
				' null,  null
				' null,  giEXPRVALUE_CHARACTER
				' null,  null
				'
				' The one remaining value in the second column, after all operators have been evaluated
				' is the return type of the expression.
				ReDim aiDummyValues(2, mcolComponents.Count())

				For iLoop1 = 1 To mcolComponents.Count()
					' Stop validating the expression if we already know its invalid.
					If iValidationCode <> ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
						Exit For
					End If

					With mcolComponents.Item(iLoop1)
						.SessionInfo = SessionInfo
						' If the current component is an operator then read the operator id into the array.
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop1).ComponentType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If .ComponentType = ExpressionComponentTypes.giCOMPONENT_OPERATOR Then
							'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							aiDummyValues(1, iLoop1) = .Component.OperatorID

							' Remember the min and max operator precedence levels for later.
							'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop1).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							iMinOperatorPrecedence = IIf((iMinOperatorPrecedence > .Component.Precedence) Or (iMinOperatorPrecedence = -1), .Component.Precedence, iMinOperatorPrecedence)
							'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop1).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							iMaxOperatorPrecedence = IIf((iMaxOperatorPrecedence < .Component.Precedence) Or (iMaxOperatorPrecedence = -1), .Component.Precedence, iMaxOperatorPrecedence)
						Else
							aiDummyValues(1, iLoop1) = -1
						End If

						'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop1).ComponentType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If .ComponentType = ExpressionComponentTypes.giCOMPONENT_FUNCTION Then
							.Component.SessionInfo = SessionInfo
							' Validate the function.
							' NB. This also determines the function's return type if not already known.
							'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							iValidationCode = .Component.ValidateFunction
							If iValidationCode <> ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
								'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop1).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If .Component.BadComponent Is Nothing Then
									mobjBadComponent = mcolComponents.Item(iLoop1)
								Else
									'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									mobjBadComponent = .Component.BadComponent
								End If
								Exit For
							End If
						End If

						' Read the component return type into the array.
						'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().ReturnType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						aiDummyValues(2, iLoop1) = .ReturnType
					End With
				Next iLoop1

				' Loop throught the expression's components checking that they are valid.
				' Evaluate operators in the correct order of precedence.
				For iLoop1 = iMinOperatorPrecedence To iMaxOperatorPrecedence
					' Stop validating the expression if we already know it's invalid.
					If iValidationCode <> ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
						Exit For
					End If

					For iLoop2 = 1 To mcolComponents.Count()
						' Stop validating the expression if we already know it's invalid.
						If iValidationCode <> ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
							Exit For
						End If

						With mcolComponents.Item(iLoop2)
							.SessionInfo = SessionInfo
							'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).ComponentType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If .ComponentType = ExpressionComponentTypes.giCOMPONENT_OPERATOR Then
								'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If .Component.Precedence = iLoop1 Then
									' Check that the operator has the correct parameter types.
									' Read the dummy value that follows the current operator.
									iParameter1Index = 0
									iParameter2Index = 0
									For iLoop3 = iLoop2 + 1 To UBound(aiDummyValues, 2)
										' If an operator follows the operator then the expression is invalid.
										If aiDummyValues(1, iLoop3) > 0 Then
											mobjBadComponent = mcolComponents.Item(iLoop2)
											iValidationCode = ExprValidationCodes.giEXPRVALIDATION_MISSINGOPERAND
											Exit For
										End If

										' Read the index of the parameter.
										If aiDummyValues(2, iLoop3) > -1 Then
											iParameter1Index = iLoop3
											Exit For
										End If
									Next iLoop3

									' If a parameter has been found then read its dummy value.
									' Otherwise the expression is invalid.
									If iParameter1Index = 0 Then
										mobjBadComponent = mcolComponents.Item(iLoop2)
										iValidationCode = ExprValidationCodes.giEXPRVALIDATION_MISSINGOPERAND
									End If

									' Read a second parameter if required.
									'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If (.Component.OperandCount = 2) And (iValidationCode = ExprValidationCodes.giEXPRVALIDATION_NOERRORS) Then

										iParameter2Index = iParameter1Index

										' Read the dummy value that precedes the current operator.
										iParameter1Index = 0
										For iLoop3 = iLoop2 - 1 To 1 Step -1
											' If an operator follows the operator then the expression is invalid.
											If aiDummyValues(1, iLoop3) > 0 Then
												mobjBadComponent = mcolComponents.Item(iLoop2)
												iValidationCode = ExprValidationCodes.giEXPRVALIDATION_MISSINGOPERAND
												Exit For
											End If

											' Read the index of the parameter.
											If aiDummyValues(2, iLoop3) > -1 Then
												iParameter1Index = iLoop3
												Exit For
											End If
										Next iLoop3

										' If a parameter has been found then read its dummy value.
										' Otherwise the expression is invalid.
										If iParameter1Index = 0 Then
											mobjBadComponent = mcolComponents.Item(iLoop2)
											iValidationCode = ExprValidationCodes.giEXPRVALIDATION_MISSINGOPERAND
										End If
									End If

									' Validate the operator by evaluating it with the dummy parmameters.
									' NB. This also determines the operator's return type if not already known.
									' Only try to evaluate the dummy operation if we still think
									' it is valid.
									If iValidationCode = ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
										'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										If Not ValidateOperatorParameters(.Component.OperatorID, iOperatorReturnType, aiDummyValues(2, iParameter1Index), IIf(.Component.OperandCount = 2, aiDummyValues(2, iParameter2Index), ExpressionValueTypes.giEXPRVALUE_UNDEFINED)) Then

											iValidationCode = ExprValidationCodes.giEXPRVALIDATION_OPERANDTYPEMISMATCH
											mobjBadComponent = mcolComponents.Item(iLoop2)
										Else
											' Check that operators with logic parameters are valid.
											If (iBadLogicColumnIndex = 0) And (iParameter2Index > 0) Then
												'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().ComponentType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												iParam1Type = mcolComponents.Item(iParameter1Index).ComponentType
												'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().ReturnType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												iParam1ReturnType = mcolComponents.Item(iParameter1Index).ReturnType
												'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().ComponentType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												iParam2Type = mcolComponents.Item(iParameter2Index).ComponentType
												'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().ReturnType. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
												iParam2ReturnType = mcolComponents.Item(iParameter2Index).ReturnType

												If ((iParam1Type = ExpressionComponentTypes.giCOMPONENT_FIELD) And (iParam1ReturnType = ExpressionValueTypes.giEXPRVALUE_LOGIC)) And (((iParam2Type <> ExpressionComponentTypes.giCOMPONENT_FIELD) And (iParam2Type <> ExpressionComponentTypes.giCOMPONENT_VALUE)) Or (iParam2ReturnType <> ExpressionValueTypes.giEXPRVALUE_LOGIC)) Then

													iBadLogicColumnIndex = iParameter1Index
												End If

												If ((iParam2Type = ExpressionComponentTypes.giCOMPONENT_FIELD) And (iParam2ReturnType = ExpressionValueTypes.giEXPRVALUE_LOGIC)) And (((iParam1Type <> ExpressionComponentTypes.giCOMPONENT_FIELD) And (iParam1Type <> ExpressionComponentTypes.giCOMPONENT_VALUE)) Or (iParam1ReturnType <> ExpressionValueTypes.giEXPRVALUE_LOGIC)) Then

													iBadLogicColumnIndex = iParameter2Index
												End If
											End If

											' Update the array to reflect the evaluated operation.
											aiDummyValues(1, iLoop2) = -1
											aiDummyValues(2, iParameter1Index) = -1
											aiDummyValues(2, iParameter2Index) = -1
										End If
									End If
								End If
							End If
						End With
					Next iLoop2
				Next iLoop1

				' Check the expression has valid syntax (ie. if the components have evaluated to a single value).
				' Get the evaluated return type while we're at it.
				If iValidationCode = ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
					iEvaluatedReturnType = ExpressionValueTypes.giEXPRVALUE_UNDEFINED

					For iLoop1 = 1 To UBound(aiDummyValues, 2)
						If aiDummyValues(2, iLoop1) > -1 Then
							' If the expression has more than one component after evaluating
							' all of the operators then the expression is invalid.
							If iEvaluatedReturnType <> ExpressionValueTypes.giEXPRVALUE_UNDEFINED Then
								iValidationCode = ExprValidationCodes.giEXPRVALIDATION_SYNTAXERROR
								Exit For
							End If

							iEvaluatedReturnType = aiDummyValues(2, iLoop1)
						End If
					Next iLoop1
				End If

				' Set the expression's return type if it is not already set.
				If iValidationCode = ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
					If (miReturnType = ExpressionValueTypes.giEXPRVALUE_UNDEFINED) Or (miReturnType = ExpressionValueTypes.giEXPRVALUE_BYREF_UNDEFINED) Then
						miReturnType = iEvaluatedReturnType
					End If
				End If

				' Check the evaluated return type matches the pre-set return type.
				If (iValidationCode = ExprValidationCodes.giEXPRVALIDATION_NOERRORS) And (iEvaluatedReturnType <> miReturnType) Then
					iValidationCode = ExprValidationCodes.giEXPRVALIDATION_EXPRTYPEMISMATCH
				End If

				' JPD20020419 Fault 3687
				' Run the filter's SQL code to see if it is valid.
				If (iValidationCode = ExprValidationCodes.giEXPRVALIDATION_NOERRORS) And pfTopLevel Then
					iTempReturnType = miReturnType
					iValidationCode = ValidateSQLCode()
					miReturnType = iTempReturnType
				End If

			Catch ex As Exception
				Return ExprValidationCodes.giEXPRVALIDATION_UNKNOWNERROR

			End Try

			Return iValidationCode

		End Function

		Private Function ValidateSQLCode(Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As ExprValidationCodes
			' Validate the expression's SQL code. This picks up on exceptions such as too many nested levels of the CASE statement.

			Dim lngCalcViews(,) As Integer
			Dim intCount As Integer
			Dim sSource As String
			Dim sSPCode As String
			Dim strJoinCode As String
			Dim iValidationCode As ExprValidationCodes = ExprValidationCodes.giEXPRVALIDATION_NOERRORS
			Dim sSQLCode As String
			Dim lngOriginalExprID As Integer
			Dim sOriginalSQLCode As String
			Dim alngSourceTables(,) As Integer
			Dim sProcName As String
			Dim avDummyPrompts(,) As Object

			Try

				ReDim avDummyPrompts(1, 0)

				If ((Not ExprDeleted(ExpressionID)) Or (mlngExpressionID = 0)) And ((miExpressionType = ExpressionTypes.giEXPR_VIEWFILTER) Or (miExpressionType = ExpressionTypes.giEXPR_RUNTIMEFILTER) Or (miExpressionType = ExpressionTypes.giEXPR_RUNTIMECALCULATION)) Then

					mfConstructed = True

					If ((miExpressionType = ExpressionTypes.giEXPR_VIEWFILTER) Or (miExpressionType = ExpressionTypes.giEXPR_RUNTIMEFILTER)) Then
						If RuntimeFilterCode(sSQLCode, False, mastrUDFsRequired, True, avDummyPrompts, plngFixedExprID, psFixedSQLCode) Then

							Try
								sProcName = General.UniqueSQLObjectName("tmpsp_ASRExprTest", 4)

								' Create the test stored procedure to see if the filter expression is valid.
								sSPCode = " CREATE PROCEDURE " & sProcName & " AS " & sSQLCode
								DB.ExecuteSql(sSPCode)

								General.DropUniqueSQLObject(sProcName, 4)

							Catch ex As Exception
								iValidationCode = ExprValidationCodes.giEXPRVALIDATION_SQLERROR
								msErrorMessage = ex.Message.RemoveSensitive()

							End Try

						Else
							iValidationCode = ExprValidationCodes.giEXPRVALIDATION_FILTEREVALUATION
						End If
					End If

					If (miExpressionType = ExpressionTypes.giEXPR_RUNTIMECALCULATION) Then
						ReDim lngCalcViews(2, 0)
						If RuntimeCalculationCode(lngCalcViews, sSQLCode, mastrUDFsRequired, False, True, avDummyPrompts, plngFixedExprID, psFixedSQLCode) Then
							' Add the required views to the JOIN code.
							strJoinCode = vbNullString
							For intCount = 1 To UBound(lngCalcViews, 2)
								' JPD20020513 Fault 3871 - Join parent tables as well as views.
								If lngCalcViews(1, intCount) = 1 Then
									sSource = gcoTablePrivileges.FindViewID(lngCalcViews(2, intCount)).RealSource
								Else
									sSource = gcoTablePrivileges.FindTableID(lngCalcViews(2, intCount)).RealSource
								End If

								strJoinCode = strJoinCode & vbNewLine & " LEFT OUTER JOIN " & sSource & " ON " & msBaseTableName & ".ID = " & sSource & ".ID"
							Next

							Try

								sSQLCode = "SELECT " & sSQLCode & " FROM " & msBaseTableName & strJoinCode

								sProcName = General.UniqueSQLObjectName("tmpsp_ASRExprTest", 4)
								sSPCode = " CREATE PROCEDURE " & sProcName & " AS " & sSQLCode
								DB.ExecuteSql(sSPCode)

								' Drop the test stored procedure.
								General.DropUniqueSQLObject(sProcName, 4)

							Catch ex As Exception
								iValidationCode = ExprValidationCodes.giEXPRVALIDATION_SQLERROR
								msErrorMessage = ex.Message.RemoveSensitive()

							End Try

						Else
							iValidationCode = ExprValidationCodes.giEXPRVALIDATION_FILTEREVALUATION
						End If
					End If

					If iValidationCode = ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
						' Need to check if all calcs/filters that use this filter are still okay.
						'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
						If (IsNothing(plngFixedExprID) And IsNothing(psFixedSQLCode)) Or ((plngFixedExprID = 0) And (psFixedSQLCode = "")) Then
							lngOriginalExprID = mlngExpressionID

							' Create an array of the IDs of the tables/view referred to in the expression.
							' This is used for joining all of the tables/views used.
							' Column 1 = 0 if this row is for a table, 1 if it is for a view.
							' Column 2 = table/view ID.
							ReDim alngSourceTables(2, 0)

							RuntimeCode(sSQLCode, alngSourceTables, False, True, avDummyPrompts, mastrUDFsRequired, plngFixedExprID, psFixedSQLCode)
							sOriginalSQLCode = sSQLCode
						Else
							lngOriginalExprID = plngFixedExprID
							sOriginalSQLCode = psFixedSQLCode
						End If

						iValidationCode = ValidateAssociatedExpressionsSQLCode(lngOriginalExprID, sOriginalSQLCode)
					End If
				End If


			Catch ex As Exception
				iValidationCode = ExprValidationCodes.giEXPRVALIDATION_UNKNOWNERROR

			End Try

			Return iValidationCode

		End Function

		Private Function ValidateAssociatedExpressionsSQLCode(ByRef plngFixedExpressionID As Integer, ByRef psFixedSQLCode As String) As ExprValidationCodes
			' Validate the SQL code for any expressions that use this expression.
			' This picks up on exceptions such as too many nested levels of the CASE statement.
			Dim iValidationCode As ExprValidationCodes
			Dim sSQL As String
			Dim rsTemp As DataTable
			Dim objComp As clsExprComponent
			Dim objExpr As clsExprExpression

			iValidationCode = ExprValidationCodes.giEXPRVALIDATION_NOERRORS

			' Do nothing if this is a new expression
			If mlngExpressionID = 0 Then
				ValidateAssociatedExpressionsSQLCode = iValidationCode
				Exit Function
			End If

			sSQL = String.Format("SELECT componentID FROM ASRSysExprComponents WHERE calculationID = {0} OR filterID = {0} OR (fieldSelectionFilter = {0} AND type = {1})" _
				, mlngExpressionID, CStr(ExpressionComponentTypes.giCOMPONENT_FIELD))
			rsTemp = DB.GetDataTable(sSQL)

			With rsTemp
				For Each objRow As DataRow In .Rows
					If iValidationCode = ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
						Exit For
					End If

					objComp = New clsExprComponent(SessionInfo)
					objComp.ComponentID = CInt(objRow("ComponentID"))

					objExpr = New clsExprExpression(SessionInfo)
					objExpr.ExpressionID = objComp.RootExpressionID
					objExpr.ConstructExpression()
					iValidationCode = objExpr.ValidateSQLCode(plngFixedExpressionID, psFixedSQLCode)
					If iValidationCode <> ExprValidationCodes.giEXPRVALIDATION_NOERRORS Then
						msErrorMessage = objExpr.ErrorMessage
					End If
					'UPGRADE_NOTE: Object objExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					objExpr = Nothing

					'UPGRADE_NOTE: Object objComp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					objComp = Nothing

				Next
			End With
			'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsTemp = Nothing

			ValidateAssociatedExpressionsSQLCode = iValidationCode

		End Function


		Public Function ConstructExpression() As Boolean
			' Read the expression definition from the database and
			' construct the hierarchy of component class objects.

			Dim dsExpression As DataSet

			Dim fOK As Boolean = True
			Dim sSQL As String
			Dim objComponent As clsExprComponent
			Dim rsExpression As DataTable

			Try

				' Do nothing if the expression is already constructed.
				If mfConstructed Then
					If miExpressionType = ExpressionTypes.giEXPR_RUNTIMECALCULATION Then
						miReturnType = ExpressionValueTypes.giEXPRVALUE_UNDEFINED
					End If

					If mlngExpressionID > 0 Then
						' Get the expression timestamp.
						sSQL = String.Format("SELECT CONVERT(integer, ASRSysExpressions.timestamp) AS intTimestamp FROM ASRSysExpressions WHERE exprID = {0}", mlngExpressionID)
						rsExpression = DB.GetDataTable(sSQL)

						With rsExpression
							fOK = (.Rows.Count > 0)
							If fOK Then
								'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
								If Not mfDontUpdateTimeStamp Then mlngTimeStamp = IIf(IsDBNull(.Rows(0)("intTimestamp")), 0, .Rows(0)("intTimestamp"))
							End If
						End With
						'UPGRADE_NOTE: Object rsExpression may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						rsExpression = Nothing
					End If
				Else

					dsExpression = DB.GetDataSet("spASRIntGetExpressionAndComponents" _
							, New SqlParameter("ExpressionID", mlngExpressionID), New SqlParameter("ExpressionType", miExpressionType))

          fOK = (dsExpression.Tables(0).Rows.Count > 0 )

          If fOK then
					  Dim rowExpression = dsExpression.Tables(0).Rows(0)

					  If rowExpression Is Nothing Then
						  InitialiseExpression()
					  Else

						  msExpressionName = rowExpression("Name").ToString()
						  mlngBaseTableID = CInt(rowExpression("TableID"))
						  miReturnType = CType(rowExpression("ReturnType"), ExpressionValueTypes)
						  miExpressionType = CType(rowExpression("Type"), ExpressionTypes)

						  If miExpressionType = ExpressionTypes.giEXPR_RUNTIMECALCULATION Then
							  miReturnType = ExpressionValueTypes.giEXPRVALUE_UNDEFINED
						  End If

						  mlngParentComponentID = CInt(rowExpression("ParentComponentID"))
						  msOwner = rowExpression("Username").ToString()
						  msAccess = rowExpression("Access").ToString()
						  msDescription = rowExpression("Description").ToString()
						  mlngTimeStamp = CInt(rowExpression("intTimestamp"))
						  msBaseTableName = rowExpression("TableName").ToString()
						  mbViewInColour = CBool(rowExpression("ViewInColour"))

					  End If
          End If

					If fOK Then
						' Clear the expressions collection of components.
						ClearComponents()

						' Get the expression definition.
						For Each objRow As DataRow In dsExpression.Tables(1).Rows

							' Instantiate a new component object.
							objComponent = New clsExprComponent(SessionInfo)

							With objComponent
								' Initialise the new component's properties.
								.ParentExpression = Me
								.ComponentID = objRow("ComponentID")

								' Instruct the new component to read it's own definition from the database.
								fOK = .ConstructComponent(objRow)
							End With

							If fOK Then
								' If the component definition was read correctly then
								' add the new component to the expression's component collection.
								mcolComponents.Add(objComponent)
							End If

							' Disassociate object variables.
							'UPGRADE_NOTE: Object objComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
							objComponent = Nothing

						Next

					End If
				End If

			Catch ex As Exception
				Return False

			End Try

			Return fOK

		End Function


		Private Sub InitialiseExpression()
			' Initialize the properties for a new expression,
			' and clear the expression's component collection.
			ExpressionID = 0

			msExpressionName = ""
			mlngParentComponentID = 0
			msOwner = Login.Username
			msAccess = "RW"
			msDescription = ""
			mlngTimeStamp = 0

			mfConstructed = True

			' Clear any existing components from
			' the expression's component collection.
			ClearComponents()

		End Sub


		Public Sub ClearComponents()

			mcolComponents.Clear()
			mcolComponents = New Collection

		End Sub

		Public Function Initialise(plngBaseTableID As Integer, plngExpressionID As Integer, piType As ExpressionTypes, piReturnType As ExpressionValueTypes, Optional ByRef plngSecondTableID As Integer = 0) As Boolean
			BaseTableID = plngBaseTableID
			ExpressionID = plngExpressionID
			miExpressionType = piType
			miReturnType = piReturnType
   		SecondTableID = plngSecondTableID

			Return True
		End Function

	End Class
End Namespace