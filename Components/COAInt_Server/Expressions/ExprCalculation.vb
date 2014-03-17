Option Strict On
Option Explicit On

Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Structures

Friend Class clsExprCalculation
	Inherits BaseExpressionComponent

	' Component definition variables.
	Private mlngCalculationID As Integer
	Private msCalculationName As String
	Private miReturnType As ExpressionValueTypes

	' Class handling variables.
	Private mobjBaseComponent As clsExprComponent

	Public Sub New(ByVal Value As SessionInfo)
		MyBase.New(Value)
	End Sub

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

	Public Function RuntimeCode(ByRef psRuntimeCode As String, ByRef palngSourceTables(,) As Integer, pfApplyPermissions As Boolean _
															, pfValidating As Boolean, ByRef pavPromptedValues As Object _
															, ByRef psUDFs() As String _
															, Optional plngFixedExprID As Integer = 0, Optional psFixedSQLCode As String = "") As Boolean

		Dim objCalc As clsExprExpression

		If mlngCalculationID = plngFixedExprID Then
			RuntimeCode = True
			psRuntimeCode = psFixedSQLCode
		Else
			' Instantiate the calculation expression.
			objCalc = New clsExprExpression(SessionInfo)
			With objCalc
				' Construct the calculation expression.
				.ExpressionID = mlngCalculationID
				.ConstructExpression()
				RuntimeCode = .RuntimeCode(psRuntimeCode, palngSourceTables, pfApplyPermissions, pfValidating, pavPromptedValues, psUDFs, plngFixedExprID, psFixedSQLCode)
			End With

			'UPGRADE_NOTE: Object objCalc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objCalc = Nothing
		End If

	End Function

	Public Function PrintComponent(ByRef piLevel As Short) As Boolean
		'Dim Printer As New Printing.PrinterSettings
		' Print the component definition to the printer object.
		On Error GoTo ErrorTrap

		Dim fOK As Boolean

		fOK = True

		' Position the printing.
		' TODO: Implement printing
		'With Printer
		'	.CurrentX = giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
		'	.CurrentY = .CurrentY + giPRINT_YSPACE
		'	Printer.Print("Calculation : " & ComponentDescription)
		'End With

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
		Dim objCalcCopy As New clsExprCalculation(SessionInfo)

		' Copy the component's basic properties.
		objCalcCopy.CalculationID = mlngCalculationID

		CopyComponent = objCalcCopy

		' Disassociate object variables.
		'UPGRADE_NOTE: Object objCalcCopy may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objCalcCopy = Nothing

	End Function

	Public ReadOnly Property ComponentType() As ExpressionComponentTypes
		Get
			Return ExpressionComponentTypes.giCOMPONENT_CALCULATION
		End Get
	End Property


	Public ReadOnly Property ReturnType() As ExpressionValueTypes
		Get
			' Return the calculation's return type.
			Dim objCalc As clsExprExpression

			' Instantiate the calculation expression.
			objCalc = New clsExprExpression(SessionInfo)
			With objCalc
				' Construct the calculation expression.
				.ExpressionID = mlngCalculationID
				.ConstructExpression()
				.ValidateExpression(False)
				miReturnType = .ReturnType
			End With

			'UPGRADE_NOTE: Object objCalc may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objCalc = Nothing

			Return miReturnType

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
		Dim sSQL As String


		Try
			sSQL = "INSERT INTO ASRSysExprComponents" & " (componentID, exprID, type, calculationID, valueLogic)" & " VALUES(" & Trim(Str(mobjBaseComponent.ComponentID)) & "," & " " & Trim(Str(mobjBaseComponent.ParentExpression.ExpressionID)) & "," & " " & Trim(Str(ExpressionComponentTypes.giCOMPONENT_CALCULATION)) & "," & " " & Trim(Str(mlngCalculationID)) & ", " & " 0)"
			DB.ExecuteSql(sSQL)
			Return True

		Catch ex As Exception
			Return False

		End Try

	End Function

	Private Sub ReadCalculation()
		' Read the calculation definition from the database.
		On Error GoTo ErrorTrap

		Dim fOK As Boolean
		Dim sSQL As String
		Dim rsCalculation As DataTable
		Dim objRow As DataRow

		' Set default values.
		msCalculationName = "<unknown>"

		' Get the calculation definition.
		sSQL = "SELECT name, returnType FROM ASRSysExpressions WHERE exprID = " & Trim(Str(mlngCalculationID))
		rsCalculation = DB.GetDataTable(sSQL)

		fOK = (rsCalculation.Rows.Count > 0)

		If fOK Then
			objRow = rsCalculation.Rows(0)
			msCalculationName = objRow("Name").ToString()
			miReturnType = CType(objRow("ReturnType"), ExpressionValueTypes)
		End If


TidyUpAndExit:
		'UPGRADE_NOTE: Object rsCalculation may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsCalculation = Nothing
		Exit Sub

ErrorTrap:
		fOK = False
		Resume TidyUpAndExit

	End Sub
End Class