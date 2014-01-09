Option Strict Off
Option Explicit On

Imports ADODB
Imports System.Collections.Generic
Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Metadata
Imports HR.Intranet.Server.Structures
Imports System.Data.SqlClient

Friend Class clsExprExpression
	Inherits BaseExpressionComponent

	' Expression definition variables.
	Private mlngExpressionID As Integer
	Private msExpressionName As String
	Private mlngBaseTableID As Integer
	Private miReturnType As ExpressionValueTypes
	Private miExpressionType As Short
	Private mlngParentComponentID As Integer
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

	Public Property ExpressionType() As Short
		Get
			' Return the expression's parent type property.
			ExpressionType = miExpressionType

		End Get
		Set(ByVal Value As Short)
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

	Public Sub ResetConstructedFlag(ByRef fValue As Object)

		'UPGRADE_WARNING: Couldn't resolve default property of object fValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mfConstructed = fValue

	End Sub

	Public Sub New(ByVal Value As LoginInfo)

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
		On Error GoTo ErrorTrap

		Dim fOK As Boolean
		Dim iLoop As Short
		Dim iIndex As Short

		fOK = True
		iIndex = 0

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

TidyUpAndExit:
		DeleteComponent = fOK
		Exit Function

ErrorTrap:
		fOK = False
		Resume TidyUpAndExit

	End Function

	Public Function AddComponent() As clsExprComponent
		' Add a new component to the expression.
		' Returns the new component object.
		On Error GoTo ErrorTrap

		Dim fOK As Boolean
		Dim objComponent As clsExprComponent

		' Instantiate a component object.
		objComponent = New clsExprComponent(Login)

		' Initialse the new component's properties.
		objComponent.ParentExpression = Me

		' Get the new component to handle its own definition.
		fOK = objComponent.NewComponent

		If fOK Then
			' If the component definition was confirmed then
			' add the new component to the expression's component
			' collection.
			mcolComponents.Add(objComponent)
		End If

TidyUpAndExit:
		If fOK Then
			AddComponent = objComponent
		Else
			'UPGRADE_NOTE: Object AddComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			AddComponent = Nothing
		End If
		' Disassociate object variables.
		'UPGRADE_NOTE: Object objComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objComponent = Nothing
		Exit Function

ErrorTrap:
		fOK = False
		Resume TidyUpAndExit

	End Function

	Public Function SelectExpression(ByRef pfLockTable As Boolean, Optional ByRef plngOptions As Integer = 0) As Boolean
	End Function

	Public Function PrintComponent(ByRef piLevel As Short) As Boolean
	End Function

	Public Function CopyComponent() As clsExprExpression
	End Function

	Public Function DeleteExpression() As Boolean
	End Function

	Public Function ValidityMessage(ByRef piInvalidityCode As Short) As String
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

	Public Function ContainsExpression(ByRef plngExprID As Integer) As Boolean
		' Retrun TRUE if the current expression (or any of its sub expressions)
		' contains the given expression. This ensures no cyclic expressions get created.
		'JPD 20040507 Fault 8600
		On Error GoTo ErrorTrap

		Dim iLoop1 As Short

		ContainsExpression = False

		For iLoop1 = 1 To mcolComponents.Count()
			If ContainsExpression Then
				Exit For
			End If

			With mcolComponents.Item(iLoop1)
				'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item().ContainsExpression. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ContainsExpression = .ContainsExpression(plngExprID)
			End With
		Next iLoop1

TidyUpAndExit:
		Exit Function

ErrorTrap:
		ContainsExpression = True
		Resume TidyUpAndExit

	End Function

	Public Function WriteExpression_Transaction() As Boolean
		' Transaction wrapper for the 'WriteExpression' function.
		On Error GoTo ErrorTrap

		Dim fOK As Boolean

		' Begin the transaction of data.
		'gADOCon.BeginTrans()

		fOK = WriteExpression()

TidyUpAndExit:
		' Commit the data transaction if everything was okay.
		'If fOK Then
		'	gADOCon.CommitTrans()
		'Else
		'	gADOCon.RollbackTrans()
		'End If
		WriteExpression_Transaction = fOK
		Exit Function

ErrorTrap:
		fOK = False
		Resume TidyUpAndExit

	End Function

	Public Function WriteExpression() As Boolean
		'  ' Write the expression definition to the database.
		On Error GoTo ErrorTrap

		Dim fOK As Boolean
		Dim sSQL As String
		Dim objComponent As clsExprComponent

		fOK = True

		If mlngExpressionID = 0 Then

			mlngExpressionID = GetUniqueID("Expressions", "ASRSysExpressions", "exprID")

			' Add a record for the new expression.
			fOK = (mlngExpressionID > 0)

			If fOK Then
				sSQL = "INSERT INTO ASRSysExpressions" & " (exprID, name, TableID, returnType, returnSize, returnDecimals, " & " type, parentComponentID, Username, access, description, ViewInColour, ExpandedNode)" & " VALUES(" & Trim(Str(mlngExpressionID)) & ", " & "'" & Replace(Trim(msExpressionName), "'", "''") & "', " & Trim(Str(mlngBaseTableID)) & ", " & IIf(miExpressionType = ExpressionTypes.giEXPR_RUNTIMECALCULATION, Trim(Str(ExpressionValueTypes.giEXPRVALUE_UNDEFINED)), Trim(Str(miReturnType))) & ", " & "0,0, " & Trim(Str(miExpressionType)) & ", " & Trim(Str(mlngParentComponentID)) & ", " & "'" & Replace(Trim(msOwner), "'", "''") & "', " & "'" & Trim(msAccess) & "', " & "'" & Replace(Trim(msDescription), "'", "'") & "', " & IIf(mbViewInColour, "1, ", "0, ") & IIf(mbExpandedNode, "1", "0") & ")"
				gADOCon.Execute(sSQL, , ADODB.CommandTypeEnum.adCmdText)

			End If
		Else
			sSQL = "UPDATE ASRSysExpressions" & " SET name = '" & Replace(Trim(msExpressionName), "'", "''") & "'," & " TableID = " & Trim(Str(mlngBaseTableID)) & "," & " returnType = " & IIf(miExpressionType = ExpressionTypes.giEXPR_RUNTIMECALCULATION, Trim(Str(ExpressionValueTypes.giEXPRVALUE_UNDEFINED)), Trim(Str(miReturnType))) & "," & " returnSize = 0," & " returnDecimals = 0," & " type = " & Trim(Str(miExpressionType)) & "," & " parentComponentID = " & Trim(Str(mlngParentComponentID)) & "," & " Username = '" & Replace(Trim(msOwner), "'", "''") & "'," & " access = '" & Trim(msAccess) & "'," & " description = '" & Replace(Trim(msDescription), "'", "''") & "', " & " ViewInColour = " & IIf(mbViewInColour, "1", "0") & " WHERE exprID = " & Trim(Str(mlngExpressionID))

			'" owner = '" & Trim(msOwner) & "'," & _
			'
			gADOCon.Execute(sSQL, , ADODB.CommandTypeEnum.adCmdText)
		End If

		If fOK Then
			' Delete the expression's existing components from the database.
			fOK = DeleteExistingComponents()

			If fOK Then
				' Add any components for this expression.
				For Each objComponent In mcolComponents
					objComponent.ParentExpression = Me
					fOK = objComponent.WriteComponent

					If Not fOK Then
						Exit For
					End If
				Next objComponent
			End If
		End If

TidyUpAndExit:
		'UPGRADE_NOTE: Object objComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objComponent = Nothing
		WriteExpression = fOK
		Exit Function

ErrorTrap:
		'NO MSGBOX ON THE SERVER ! - MsgBox "Error saving the expression.", _
		'vbOKOnly + vbExclamation, App.ProductName
		fOK = False
		Resume TidyUpAndExit

	End Function

	Public Function RuntimeCode(ByRef psRuntimeCode As String, ByRef palngSourceTables(,) As Integer, ByRef pfApplyPermissions As Boolean _
															, ByRef pfValidating As Boolean, ByRef pavPromptedValues As Object _
															, ByRef psUDFs() As String _
															, Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean

		' Return the SQL code that defines the expression.
		' Used when creating the 'where clause' for view definitions.
		On Error GoTo ErrorTrap

		Dim fOK As Boolean
		Dim iLoop1 As Short
		Dim iLoop2 As Short
		Dim iLoop3 As Short
		Dim iParameter1Index As Short
		Dim iParameter2Index As Short
		Dim iMinOperatorPrecedence As Short
		Dim iMaxOperatorPrecedence As Short
		Dim sCode As String
		Dim sComponentCode As String
		Dim vParameter1 As Object
		Dim vParameter2 As Object
		Dim avValues(,) As Object

		fOK = True
		sCode = ""

		iMinOperatorPrecedence = -1
		iMaxOperatorPrecedence = -1

		' Create an array of the components in the expression.
		' Column 1 = operator id.
		' Column 2 = component where clause code.
		ReDim avValues(2, mcolComponents.Count())
		For iLoop1 = 1 To mcolComponents.Count()
			With mcolComponents.Item(iLoop1)
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

								If (miExpressionType = ExpressionTypes.giEXPR_RUNTIMECALCULATION) Or (miExpressionType = ExpressionTypes.giEXPR_RUNTIMEFILTER) Or (miExpressionType = ExpressionTypes.giEXPR_LINKFILTER) Then
									'UPGRADE_WARNING: Couldn't resolve default property of object mcolComponents.Item(iLoop2).Component. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									If (.Component.ReturnType = ExpressionValueTypes.giEXPRVALUE_LOGIC) And ((.Component.OperatorID <> 5) And (.Component.OperatorID <> 6) And (.Component.OperatorID <> 13)) Then
										'UPGRADE_WARNING: Couldn't resolve default property of object avValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
										'UPGRADE_WARNING: Couldn't resolve default property of object avValues(2, iLoop2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
					sCode = avValues(2, iLoop1)
					Exit For
				End If
			Next iLoop1

		End If

TidyUpAndExit:
		If fOK Then
			psRuntimeCode = sCode
		Else
			psRuntimeCode = ""
		End If
		RuntimeCode = fOK

		Exit Function

ErrorTrap:
		fOK = False
		Resume TidyUpAndExit

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
		Dim iLoop1 As Short
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

	Friend Function RuntimeCalculationCode(ByRef palngSourceTables(,) As Integer, ByRef psCalcCode As String, ByRef pastrUDFsRequired() As String _
																				 , ByRef pfApplyPermissions As Boolean _
																				 , Optional ByRef pfValidating As Boolean = False, Optional ByRef pavPromptedValues As Object = Nothing _
																				 , Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean
		' Return TRUE if the Calculation code was created okay.
		' Return the runtime Calculation SQL code in the parameter 'psCalcCode'.
		' Apply permissions to the Calculation code only if the 'pfApplyPermissions' parameter is TRUE.
		On Error GoTo ErrorTrap

		Dim fOK As Boolean
		Dim sRuntimeSQL As String
		Dim avDummyPrompts(,) As Object

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

TidyUpAndExit:
		If fOK Then
			psCalcCode = sRuntimeSQL
		Else
			psCalcCode = ""
		End If
		RuntimeCalculationCode = fOK

		Exit Function

ErrorTrap:
		fOK = False
		Resume TidyUpAndExit

	End Function

	Friend Function DeleteExistingComponents() As Boolean
		' Delete the expression's components and sub-expression's
		' (ie. function parameter expressions) from the database.
		On Error GoTo ErrorTrap

		Dim fOK As Boolean
		Dim sSQL As String
		Dim sDeletedExpressionIDs As String
		Dim rsSubExpressions As ADODB.Recordset
		Dim objExpr As clsExprExpression

		fOK = True
		sDeletedExpressionIDs = ""

		' Get the expression's function components from the database.
		sSQL = "SELECT ASRSysExpressions.exprID" & " FROM ASRSysExpressions" & " INNER JOIN ASRSysExprComponents" & "   ON ASRSysExpressions.parentComponentID = ASRSysExprComponents.componentID" & " AND ASRSysExprComponents.exprID = " & Trim(Str(mlngExpressionID))
		rsSubExpressions = General.GetRecordsInTransaction(sSQL)
		With rsSubExpressions
			Do While (Not .EOF) And fOK
				' Instantiate each function parameter expression.
				' Instruct the function parameter expression to delete its components.
				objExpr = New clsExprExpression(Login)
				objExpr.ExpressionID = .Fields("ExprID").Value
				fOK = objExpr.DeleteExistingComponents
				'UPGRADE_NOTE: Object objExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objExpr = Nothing

				' Add the ID of the sub-expression to the string of sub-expressions to be deleted.
				sDeletedExpressionIDs = sDeletedExpressionIDs & IIf(Len(sDeletedExpressionIDs) > 0, ", ", "") & Trim(Str(.Fields("ExprID").Value))

				.MoveNext()
			Loop

			.Close()
		End With
		'UPGRADE_NOTE: Object rsSubExpressions may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsSubExpressions = Nothing

		If Len(sDeletedExpressionIDs) > 0 Then
			' Delete all existing sub-expressions for this expression from the database.
			sSQL = "DELETE FROM ASRSysExpressions" & " WHERE exprID IN (" & sDeletedExpressionIDs & ")"
			gADOCon.Execute(sSQL, , ADODB.CommandTypeEnum.adCmdText)
		End If

		' Delete all existing components for this expression from the database.
		sSQL = "DELETE FROM ASRSysExprComponents" & " WHERE exprID = " & Trim(Str(mlngExpressionID))
		gADOCon.Execute(sSQL, , ADODB.CommandTypeEnum.adCmdText)

TidyUpAndExit:
		'UPGRADE_NOTE: Object objExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objExpr = Nothing

		'UPGRADE_NOTE: Object rsSubExpressions may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsSubExpressions = Nothing
		DeleteExistingComponents = fOK
		Exit Function

ErrorTrap:
		fOK = False
		Resume TidyUpAndExit

	End Function

	Public Function ValidateExpression(ByRef pfTopLevel As Boolean) As ExprValidationCodes
		' Validate the expression. Return a code defining the validity of the expression.
		' NB. This function is also good for evaluating the return type of an expression
		' which has definite return type (eg. function sub-expressions, runtime calcs, etc).
		On Error GoTo ErrorTrap

		Dim iLoop1 As Short
		Dim iLoop2 As Short
		Dim iLoop3 As Short
		Dim iParam1Type As Short
		Dim iParam2Type As Short
		Dim iParameter1Index As Short
		Dim iParameter2Index As Short
		Dim iParam1ReturnType As Short
		Dim iParam2ReturnType As Short
		Dim iOperatorReturnType As ExpressionValueTypes
		Dim iBadLogicColumnIndex As Short
		Dim iMinOperatorPrecedence As Short
		Dim iMaxOperatorPrecedence As Short
		Dim iValidationCode As ExprValidationCodes
		Dim iEvaluatedReturnType As ExpressionValueTypes
		Dim aiDummyValues(,) As Short
		Dim avDummyPrompts(,) As Object
		Dim iTempReturnType As Short

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

		Return iValidationCode

ErrorTrap:
		Return ExprValidationCodes.giEXPRVALIDATION_UNKNOWNERROR

	End Function

	Private Function ValidateSQLCode(Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As ExprValidationCodes
		' Validate the expression's SQL code. This picks up on errors such as too many nested levels of the CASE statement.
		On Error GoTo ErrorTrap

		Dim lngCalcViews(,) As Integer
		Dim intCount As Short
		Dim sSource As String
		Dim sSPCode As String
		Dim strJoinCode As String
		Dim iValidationCode As ExprValidationCodes
		Dim sSQLCode As String
		Dim lngOriginalExprID As Integer
		Dim sOriginalSQLCode As String
		Dim alngSourceTables(,) As Integer
		Dim sProcName As String
		Dim avDummyPrompts(,) As Object
		Dim intStart As Short
		Dim intFound As Short

		ReDim avDummyPrompts(1, 0)

		iValidationCode = ExprValidationCodes.giEXPRVALIDATION_NOERRORS

		If ((Not ExprDeleted(ExpressionID)) Or (mlngExpressionID = 0)) And ((miExpressionType = ExpressionTypes.giEXPR_VIEWFILTER) Or (miExpressionType = ExpressionTypes.giEXPR_RUNTIMEFILTER) Or (miExpressionType = ExpressionTypes.giEXPR_RUNTIMECALCULATION)) Then

			mfConstructed = True

			If ((miExpressionType = ExpressionTypes.giEXPR_VIEWFILTER) Or (miExpressionType = ExpressionTypes.giEXPR_RUNTIMEFILTER)) Then
				If RuntimeFilterCode(sSQLCode, False, mastrUDFsRequired, True, avDummyPrompts, plngFixedExprID, psFixedSQLCode) Then

					On Error GoTo SQLCodeErrorTrap

					sProcName = General.UniqueSQLObjectName("tmpsp_ASRExprTest", 4)

					' Create the test stored procedure to see if the filter expression is valid.
					sSPCode = " CREATE PROCEDURE " & sProcName & " AS " & sSQLCode
					gADOCon.Execute(sSPCode, , CommandTypeEnum.adCmdText)

					General.DropUniqueSQLObject(sProcName, 4)

					On Error GoTo ErrorTrap
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

					sSQLCode = "SELECT " & sSQLCode & " FROM " & msBaseTableName & strJoinCode

					On Error GoTo SQLCodeErrorTrap

					sProcName = General.UniqueSQLObjectName("tmpsp_ASRExprTest", 4)

					' Create the test stored procedure to see if the filter expression is valid.
					sSPCode = " CREATE PROCEDURE " & sProcName & " AS " & sSQLCode
					gADOCon.Execute(sSPCode, , CommandTypeEnum.adCmdText)

					' Drop the test stored procedure.
					General.DropUniqueSQLObject(sProcName, 4)

					On Error GoTo ErrorTrap
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

TidyUpAndExit:
		ValidateSQLCode = iValidationCode
		Exit Function

SQLCodeErrorTrap:
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If (IsNothing(plngFixedExprID) And IsNothing(psFixedSQLCode)) Or ((plngFixedExprID = 0) And (psFixedSQLCode = "")) Then
			iValidationCode = ExprValidationCodes.giEXPRVALIDATION_SQLERROR
		Else
			iValidationCode = ExprValidationCodes.giEXPRVALIDATION_ASSOCSQLERROR
		End If
		msErrorMessage = Err.Description

		Do
			intStart = intFound
			intFound = InStr(intStart + 1, msErrorMessage, "]")
		Loop While intFound > 0

		If intStart > 0 And intStart < Len(Trim(msErrorMessage)) Then
			msErrorMessage = Trim(Mid(msErrorMessage, intStart + 1))
		End If

		'UPGRADE_NOTE: Object mobjBadComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mobjBadComponent = Nothing
		Resume TidyUpAndExit

ErrorTrap:
		iValidationCode = ExprValidationCodes.giEXPRVALIDATION_UNKNOWNERROR
		Resume TidyUpAndExit

	End Function

	Private Function ValidateAssociatedExpressionsSQLCode(ByRef plngFixedExpressionID As Integer, ByRef psFixedSQLCode As String) As ExprValidationCodes
		' Validate the SQL code for any expressions that use this expression.
		' This picks up on errors such as too many nested levels of the CASE statement.
		Dim iValidationCode As ExprValidationCodes
		Dim sSQL As String
		Dim rsTemp As Recordset
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
		rsTemp = dataAccess.OpenRecordset(sSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)

		With rsTemp
			Do While (Not .EOF) And (iValidationCode = ExprValidationCodes.giEXPRVALIDATION_NOERRORS)
				objComp = New clsExprComponent(Login)
				objComp.ComponentID = .Fields("ComponentID").Value

				objExpr = New clsExprExpression(Login)
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

				.MoveNext()
			Loop
			.Close()
		End With
		'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsTemp = Nothing

		ValidateAssociatedExpressionsSQLCode = iValidationCode

	End Function


	Public Function ConstructExpression() As Boolean
		' Read the expression definition from the database and
		' construct the hierarchy of component class objects.
		On Error GoTo ErrorTrap

		Dim dsExpression As DataSet

		Dim fOK As Boolean
		Dim sSQL As String
		Dim objComponent As clsExprComponent
		Dim rsExpression As Recordset

		fOK = True

		' Do nothing if the expression is already constructed.
		If mfConstructed Then
			If miExpressionType = ExpressionTypes.giEXPR_RUNTIMECALCULATION Then
				miReturnType = ExpressionValueTypes.giEXPRVALUE_UNDEFINED
			End If

			If mlngExpressionID > 0 Then
				' Get the expression timestamp.
				sSQL = String.Format("SELECT CONVERT(integer, ASRSysExpressions.timestamp) AS intTimestamp FROM ASRSysExpressions WHERE exprID = {0}", mlngExpressionID)
				rsExpression = dataAccess.OpenRecordset(sSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)

				With rsExpression
					fOK = Not (.EOF And .BOF)
					If fOK Then
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						If Not mfDontUpdateTimeStamp Then mlngTimeStamp = IIf(IsDBNull(.Fields("intTimestamp").Value), 0, .Fields("intTimestamp").Value)
					End If
					.Close()
				End With
				'UPGRADE_NOTE: Object rsExpression may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				rsExpression = Nothing
			End If
		Else

			dsExpression = dataAccess.GetDataSet("spASRIntGetExpressionAndComponents" _
					, New SqlParameter("ExpressionID", mlngExpressionID), New SqlParameter("ExpressionType", miExpressionType))

			Dim rowExpression = dsExpression.Tables(0).Rows(0)

			If rowExpression Is Nothing Then
				InitialiseExpression()
			Else

				msExpressionName = rowExpression("Name")
				mlngBaseTableID = rowExpression("TableID")
				miReturnType = rowExpression("ReturnType")
				miExpressionType = rowExpression("Type")

				If miExpressionType = ExpressionTypes.giEXPR_RUNTIMECALCULATION Then
					miReturnType = ExpressionValueTypes.giEXPRVALUE_UNDEFINED
				End If

				mlngParentComponentID = rowExpression("ParentComponentID")
				msOwner = rowExpression("Username")
				msAccess = rowExpression("Access")
				msDescription = rowExpression("Description")
				mlngTimeStamp = rowExpression("intTimestamp")
				msBaseTableName = rowExpression("TableName")
				mbViewInColour = rowExpression("ViewInColour")

			End If


			If fOK Then
				' Clear the expressions collection of components.
				ClearComponents()

				' Get the expression definition.
				For Each objRow As DataRow In dsExpression.Tables(1).Rows

					' Instantiate a new component object.
					objComponent = New clsExprComponent(Login)

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

TidyUpAndExit:
		mfConstructed = fOK
		'UPGRADE_NOTE: Object rsExpression may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsExpression = Nothing
		'UPGRADE_NOTE: Object objComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objComponent = Nothing
		Return fOK

ErrorTrap:
		fOK = False
		'NO MSGBOX ON THE SERVER ! - MsgBox "Error constructing the expression.", _
		'vbOKOnly + vbExclamation, App.ProductName
		Err.Number = False
		Resume TidyUpAndExit

	End Function


	Private Sub InitialiseExpression()
		' Initialize the properties for a new expression,
		' and clear the expression's component collection.
		ExpressionID = 0

		msExpressionName = ""
		mlngParentComponentID = 0
		msOwner = gsUsername
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

	Public Function Initialise(ByRef plngBaseTableID As Integer, ByRef plngExpressionID As Integer, ByRef piType As Short, ByRef piReturnType As Short) As Boolean
		' Initialise the expression object.
		' Return TRUE if everything was initialised okay.
		On Error GoTo ErrorTrap

		Dim fOK As Boolean

		fOK = True

		BaseTableID = plngBaseTableID
		ExpressionID = plngExpressionID
		miExpressionType = piType
		miReturnType = piReturnType

TidyUpAndExit:
		Initialise = fOK
		Exit Function

ErrorTrap:
		fOK = False
		Resume TidyUpAndExit

	End Function


	Public Function ValidateSelection() As Boolean
		' Validate the expression section.
		On Error GoTo ErrorTrap

		Dim fOK As Boolean

		fOK = ConstructExpression()

TidyUpAndExit:
		ValidateSelection = fOK
		Exit Function

ErrorTrap:
		fOK = False
		'NO MSGBOX ON THE SERVER ! - MsgBox Err.Description, vbExclamation + vbOKOnly, App.ProductName
		Err.Number = False
		Resume TidyUpAndExit

	End Function

End Class