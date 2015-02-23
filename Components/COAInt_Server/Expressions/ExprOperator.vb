Option Strict On
Option Explicit On

Imports System.Data.SqlClient
Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums

Namespace Expressions
	Friend Class clsExprOperator
		Inherits BaseExpressionComponent

		' Component definition variables.
		Private mlngOperatorID As Integer
		Private msOperatorName As String
		Private miReturnType As ExpressionValueTypes
		Private miOperandCount As Integer
		Private miPrecedence As Integer
		Private msSPName As String
		Private msSQLCode As String
		Private msSQLType As String
		Private mfUnknownParameterTypes As Boolean
		Private mfCheckDivideByZero As Boolean
		Private msSQLFixedParam1 As String
		Private mbCastAsFloat As Boolean

		' Class handling variables.
		Private mobjBaseComponent As clsExprComponent

		Public Sub New(ByVal Value As SessionInfo)
			MyBase.New(Value)
		End Sub

		Public Function ContainsExpression(plngExprID As Integer) As Boolean
			Return False
		End Function

		Public ReadOnly Property SQLType() As String
			Get
				' Return the operator SQL Type property.
				SQLType = msSQLType

			End Get
		End Property

		Public ReadOnly Property ReturnType() As ExpressionValueTypes
			Get
				Return miReturnType
			End Get
		End Property

		Public ReadOnly Property ComponentType() As ExpressionComponentTypes
			Get
				Return ExpressionComponentTypes.giCOMPONENT_OPERATOR
			End Get
		End Property

		Public ReadOnly Property Precedence() As Integer
			Get
				' Return the operator precedence property.
				Precedence = miPrecedence

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

		Public Property OperatorID() As Integer
			Get
				' Return the operator ID property.
				OperatorID = mlngOperatorID

			End Get
			Set(ByVal Value As Integer)
				' Setthe operator ID property.
				mlngOperatorID = Value

				' Read the operator definition from the database.
				ReadOperator()

			End Set
		End Property

		Public ReadOnly Property CastAsFloat() As Boolean
			Get

				' Property used to force surrounding values into using numeric (i.e. 0.00 as opposed to 0)
				' Is necessary in order to get correct values back from SQL when dealing with division signs
				CastAsFloat = mbCastAsFloat

			End Get
		End Property

		Public Property UnknownParameterTypes() As Boolean
			Get
				UnknownParameterTypes = mfUnknownParameterTypes

			End Get
			Set(ByVal Value As Boolean)
				mfUnknownParameterTypes = Value

			End Set
		End Property

		Public Property SQLCode() As String
			Get
				' Return the operator's SQL code.
				SQLCode = msSQLCode

			End Get
			Set(ByVal Value As String)
				' Set the operator's SQL code.
				msSQLCode = Value

			End Set
		End Property

		Public Property SPName() As String
			Get
				' Return the operator's stored procedure name.
				SPName = msSPName

			End Get
			Set(ByVal Value As String)
				' Set the operator's stored procedure name.
				msSPName = Value

			End Set
		End Property

		Public ReadOnly Property OperandCount() As Integer
			Get
				' Return the operator's operand count.
				OperandCount = miOperandCount

			End Get
		End Property

		Public ReadOnly Property ComponentDescription() As String
			Get
				' Return the operator description.
				ComponentDescription = msOperatorName

			End Get
		End Property

		Public Property CheckDivideByZero() As Boolean
			Get
				' Return the 'check for divide by zero' flag.
				CheckDivideByZero = mfCheckDivideByZero

			End Get
			Set(ByVal Value As Boolean)
				' Set the 'check for divide by zero' flag.
				mfCheckDivideByZero = Value

			End Set
		End Property

		Public Property SQLFixedParam1() As String
			Get
				' Return the first fixed SQL parameter.
				SQLFixedParam1 = msSQLFixedParam1

			End Get
			Set(ByVal Value As String)
				' Set the first fixed SQL parameter.
				msSQLFixedParam1 = Value

			End Set
		End Property

		Public Function WriteComponent() As Boolean

			Try
				DB.ExecuteSP("spASRIntSaveComponent", _
						New SqlParameter("componentID", SqlDbType.Int) With {.Value = mobjBaseComponent.ComponentID}, _
						New SqlParameter("expressionID", SqlDbType.Int) With {.Value = mobjBaseComponent.ParentExpression.ExpressionID}, _
						New SqlParameter("type", SqlDbType.TinyInt) With {.Value = ExpressionComponentTypes.giCOMPONENT_OPERATOR}, _
						New SqlParameter("calculationID", SqlDbType.Int), _
						New SqlParameter("filterID", SqlDbType.Int), _
						New SqlParameter("functionID", SqlDbType.Int), _
						New SqlParameter("operatorID", SqlDbType.Int) With {.Value = mlngOperatorID},
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

		Public Function CopyComponent() As Object
			' Copies the selected component.
			' When editting a component we actually copy the component first
			' and edit the copy. If the changes are confirmed then the copy
			' replaces the original. If the changes are cancelled then the
			' copy is discarded.
			Dim objOperatorCopy As New clsExprOperator(SessionInfo)

			' Copy the component's basic properties.
			objOperatorCopy.OperatorID = mlngOperatorID

			CopyComponent = objOperatorCopy

			' Disassociate object variables.
			'UPGRADE_NOTE: Object objOperatorCopy may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objOperatorCopy = Nothing

		End Function

		Private Sub ReadOperator()

			Try
				Dim objOperator = Operators.GetById(mlngOperatorID)
				msOperatorName = objOperator.Name
				miReturnType = objOperator.ReturnType
				miPrecedence = objOperator.Precedence
				miOperandCount = objOperator.OperandCount
				msSPName = IIf(IsDBNull(objOperator.SPName), "", objOperator.SPName).ToString()
				msSQLCode = objOperator.SQLCode
				msSQLType = objOperator.SQLType
				mfCheckDivideByZero = objOperator.CheckDivideByZero
				msSQLFixedParam1 = objOperator.SQLFixedParam1
				mbCastAsFloat = objOperator.CastAsFloat

				mfUnknownParameterTypes = (objOperator.Parameters.Count = 0)

			Catch ex As Exception
				Throw

			End Try

		End Sub

		Public Function RuntimeCode(ByRef psRuntimeCode As String, palngSourceTables(,) As Integer, pfApplyPermissions As Boolean _
															, pfValidating As Boolean, pavPromptedValues As Object _
															, psUDFs() As String _
															, Optional plngFixedExprID As Integer = 0, Optional psFixedSQLCode As String = "") As Boolean

			psRuntimeCode = msSQLCode
			Return True
		End Function

	End Class
End Namespace