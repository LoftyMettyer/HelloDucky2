Option Strict On
Option Explicit On

Imports System.Data.SqlClient
Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums

Namespace Expressions
	Friend Class clsExprFilter
		Inherits BaseExpressionComponent

		' Component definition variables.
		Private mlngFilterID As Integer
		Private msFilterName As String
		Private miReturnType As ExpressionValueTypes

		' Class handling variables.
		Private mobjBaseComponent As clsExprComponent

		Public Sub New(ByVal Value As SessionInfo)
			MyBase.New(Value)
		End Sub

		Public Function RuntimeCode(ByRef psRuntimeCode As String, ByRef palngSourceTables(,) As Integer, ByRef pfApplyPermissions As Boolean _
																, ByRef pfValidating As Boolean, ByRef pavPromptedValues As Object _
																, ByRef psUDFs() As String _
																, Optional plngFixedExprID As Integer = 0, Optional psFixedSQLCode As String = "") As Boolean

			Dim objFilter As clsExprExpression

			Dim strRuntimeCode As String = ""
			Dim bOK As Boolean

			' Instantiate and generate the runtime for the filter expression.
			objFilter = New clsExprExpression(SessionInfo)
			With objFilter
				.ExpressionID = mlngFilterID
				.ConstructExpression()
				bOK = .RuntimeCode(strRuntimeCode, palngSourceTables, pfApplyPermissions, pfValidating, pavPromptedValues, psUDFs, plngFixedExprID, psFixedSQLCode)
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
			'	Printer.Print("Filter : " & ComponentDescription)
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
			Dim objFilterCopy As New clsExprFilter(SessionInfo)

			' Copy the component's basic properties.
			objFilterCopy.FilterID = mlngFilterID

			CopyComponent = objFilterCopy

			' Disassociate object variables.
			'UPGRADE_NOTE: Object objFilterCopy may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objFilterCopy = Nothing

		End Function

		Public ReadOnly Property ComponentType() As ExpressionComponentTypes
			Get
				Return ExpressionComponentTypes.giCOMPONENT_FILTER
			End Get
		End Property

		Public ReadOnly Property ReturnType() As ExpressionValueTypes
			Get
				' Return the filter's return type.
				Dim objFilter As clsExprExpression

				' Instantiate the filter expression.
				objFilter = New clsExprExpression(SessionInfo)

				With objFilter
					' Construct the filter expression.
					.ExpressionID = mlngFilterID
					.ConstructExpression()
					.ValidateExpression(False)
					miReturnType = .ReturnType
				End With

				'UPGRADE_NOTE: Object objFilter may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				objFilter = Nothing

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

		Public Function WriteComponent() As Boolean

			Try
				DB.ExecuteSP("spASRIntSaveComponent", _
						New SqlParameter("componentID", SqlDbType.Int) With {.Value = mobjBaseComponent.ComponentID}, _
						New SqlParameter("expressionID", SqlDbType.Int) With {.Value = mobjBaseComponent.ParentExpression.ExpressionID}, _
						New SqlParameter("type", SqlDbType.TinyInt) With {.Value = ExpressionComponentTypes.giCOMPONENT_FILTER}, _
						New SqlParameter("calculationID", SqlDbType.Int), _
						New SqlParameter("filterID", SqlDbType.Int) With {.Value = mlngFilterID}, _
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

		Private Sub ReadFilter()
			' Read the filter definition from the database.
			On Error GoTo ErrorTrap

			Dim fOK As Boolean
			Dim sSQL As String
			Dim rsFilter As DataTable

			' Set default values.
			msFilterName = "<unknown>"

			' Get the filter definition.
			sSQL = "SELECT name, returnType FROM ASRSysExpressions WHERE exprID = " & Trim(Str(mlngFilterID))
			rsFilter = DB.GetDataTable(sSQL)
			With rsFilter
				fOK = (.Rows.Count > 0)

				If fOK Then
					msFilterName = .Rows(0)("Name").ToString()
					miReturnType = CType(.Rows(0)("ReturnType"), ExpressionValueTypes)
				End If

			End With

TidyUpAndExit:
			'UPGRADE_NOTE: Object rsFilter may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsFilter = Nothing
			Exit Sub

ErrorTrap:
			fOK = False
			Resume TidyUpAndExit

		End Sub

	End Class
End Namespace