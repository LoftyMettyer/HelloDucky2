Option Strict On
Option Explicit On

Imports System.Data.SqlClient
Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums

Namespace Expressions
	Public Class clsExprCalculation
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

		Public Function ContainsExpression(plngExprID As Integer) As Boolean
			' Retrun TRUE if the current expression (or any of its sub expressions)
			' contains the given expression. This ensures no cyclic expressions get created.
			'JPD 20040507 Fault 8600

			Dim bContains As Boolean

			Try

				' Check if the calc component IS the one we're checking for.
				bContains = (plngExprID = mlngCalculationID)

				If Not bContains Then
					' The calc component IS NOT the one we're checking for.
					' Check if it contains the one we're looking for.
					bContains = HasExpressionComponent(mlngCalculationID, plngExprID)
				End If

			Catch ex As Exception
				Return True

			End Try

			Return bContains

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

			End If

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

			Return objCalcCopy

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
				Return msCalculationName
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

		Public Function WriteComponent() As Boolean

			Try
				DB.ExecuteSP("spASRIntSaveComponent", _
						New SqlParameter("componentID", SqlDbType.Int) With {.Value = mobjBaseComponent.ComponentID}, _
						New SqlParameter("expressionID", SqlDbType.Int) With {.Value = mobjBaseComponent.ParentExpression.ExpressionID}, _
						New SqlParameter("type", SqlDbType.TinyInt) With {.Value = ExpressionComponentTypes.giCOMPONENT_CALCULATION}, _
						New SqlParameter("calculationID", SqlDbType.Int) With {.Value = mlngCalculationID}, _
						New SqlParameter("filterID", SqlDbType.Int), _
						New SqlParameter("functionID", SqlDbType.Int), _
						New SqlParameter("operatorID", SqlDbType.Int), _
						New SqlParameter("valueType", SqlDbType.TinyInt), _
						New SqlParameter("valueCharacter", SqlDbType.VarChar, 255), _
						New SqlParameter("valueNumeric", SqlDbType.Float), _
						New SqlParameter("valueLogic", SqlDbType.Bit), _
						New SqlParameter("valueDate", SqlDbType.DateTime), _
						New SqlParameter("LookupTableID", SqlDbType.Int), _
						New SqlParameter("LookupColumnID", SqlDbType.Int), _
						New SqlParameter("fieldTableID", SqlDbType.Int), _
						New SqlParameter("fieldColumnID", SqlDbType.Int), _
						New SqlParameter("fieldPassBy", SqlDbType.TinyInt), _
						New SqlParameter("fieldSelectionRecord", SqlDbType.TinyInt), _
						New SqlParameter("fieldSelectionLine", SqlDbType.Int), _
						New SqlParameter("fieldSelectionOrderID", SqlDbType.Int), _
						New SqlParameter("fieldSelectionFilter", SqlDbType.Int), _
						New SqlParameter("promptDescription", SqlDbType.VarChar, 255), _
						New SqlParameter("promptSize", SqlDbType.SmallInt), _
						New SqlParameter("promptDecimals", SqlDbType.SmallInt), _
						New SqlParameter("promptMask", SqlDbType.VarChar, 255), _
						New SqlParameter("promptDateType", SqlDbType.Int))

				Return True

			Catch ex As Exception
				Return False

			End Try

		End Function

		Private Sub ReadCalculation()
			' Read the calculation definition from the database.

			Try

				Dim fOK As Boolean
				Dim sSQL As String
				Dim rsCalculation As DataTable
				Dim objRow As DataRow

				' Set default values.
				msCalculationName = "<unknown>"

				' Get the calculation definition.
				sSQL = "SELECT name, returnType FROM ASRSysExpressions WHERE exprID = " & mlngCalculationID.ToString()
				rsCalculation = DB.GetDataTable(sSQL)

				fOK = (rsCalculation.Rows.Count > 0)

				If fOK Then
					objRow = rsCalculation.Rows(0)
					msCalculationName = objRow("Name").ToString()
					miReturnType = CType(objRow("ReturnType"), ExpressionValueTypes)
				End If

			Catch ex As Exception
				Throw

			End Try

		End Sub
	End Class

End Namespace