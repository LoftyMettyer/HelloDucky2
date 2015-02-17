Option Strict On
Option Explicit On

Imports System.Data.SqlClient
Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums

Namespace Expressions
	Friend Class clsExprTableLookup
		Inherits BaseExpressionComponent

		' Component definition variables.
		Private miType As ExpressionValueTypes
		Private mdblNumericValue As Double
		Private msCharacterValue As String
		Private mfLogicValue As Boolean

		Private miTableID As Short
		Private miColumnID As Short

		'MH20010201 Fault 1576
		'Private mdtDateValue As Date
		Private mdtDateValue As Object

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

			ContainsExpression = False

TidyUpAndExit:
			Exit Function

ErrorTrap:
			ContainsExpression = True
			Resume TidyUpAndExit

		End Function

		Public Function RuntimeCode(ByRef psRuntimeCode As String, palngSourceTables(,) As Integer, pfApplyPermissions As Boolean _
																, pfValidating As Boolean, pavPromptedValues As Object _
																, psUDFs() As String _
																, Optional plngFixedExprID As Integer = 0, Optional psFixedSQLCode As String = "") As Boolean
			' Return the SQL code for the component.
			On Error GoTo ErrorTrap

			Dim fOK As Boolean
			Dim sCode As String

			fOK = True
			sCode = ""

			Select Case miType
				Case ExpressionValueTypes.giEXPRVALUE_CHARACTER
					sCode = "'" & Replace(msCharacterValue, "'", "''") & "'"
				Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
					sCode = Trim(Str(mdblNumericValue))
				Case ExpressionValueTypes.giEXPRVALUE_LOGIC
					sCode = IIf(mfLogicValue, "1", "0").ToString()
				Case ExpressionValueTypes.giEXPRVALUE_DATE
					'MH20010201 Fault 1576
					'sCode = "convert(datetime, '" & Format(mdtDateValue, "MM/dd/yyyy") & "')"
					'UPGRADE_WARNING: Couldn't resolve default property of object mdtDateValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					sCode = IIf(IsDBNull(mdtDateValue), "null", "convert(datetime, '" & VB6.Format(mdtDateValue, "MM/dd/yyyy") & "')").ToString()
			End Select

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
			'	Printer.Print(ComponentDescription)
			'End With

TidyUpAndExit:
			PrintComponent = fOK
			Exit Function

ErrorTrap:
			fOK = False
			Resume TidyUpAndExit

		End Function

		Public Function WriteComponent() As Boolean

			Try
				DB.ExecuteSP("spASRIntSaveComponent", _
						New SqlParameter("componentID", SqlDbType.Int) With {.Value = mobjBaseComponent.ComponentID}, _
						New SqlParameter("expressionID", SqlDbType.Int) With {.Value = mobjBaseComponent.ParentExpression.ExpressionID}, _
						New SqlParameter("type", SqlDbType.TinyInt) With {.Value = ExpressionComponentTypes.giCOMPONENT_TABLEVALUE}, _
						New SqlParameter("calculationID", SqlDbType.Int), _
						New SqlParameter("filterID", SqlDbType.Int), _
						New SqlParameter("functionID", SqlDbType.Int), _
						New SqlParameter("operatorID", SqlDbType.Int), _
						New SqlParameter("valueType", SqlDbType.TinyInt) With {.Value = miType}, _
						New SqlParameter("valueCharacter", SqlDbType.VarChar, 255) With {.Value = msCharacterValue}, _
						New SqlParameter("valueNumeric", SqlDbType.Float) With {.Value = mdblNumericValue}, _
						New SqlParameter("valueLogic", SqlDbType.Bit) With {.Value = mfLogicValue}, _
						New SqlParameter("valueDate", SqlDbType.DateTime) With {.Value = mdtDateValue}, _
						New SqlParameter("LookupTableID", SqlDbType.Int) With {.Value = miTableID},
						New SqlParameter("LookupColumnID", SqlDbType.Int) With {.Value = miColumnID},
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
			Dim objValueCopy As New clsExprTableLookup(SessionInfo)

			' Copy the component's basic properties.
			With objValueCopy
				.ReturnType = miType
				'UPGRADE_WARNING: Couldn't resolve default property of object Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Value = Value
				.ColumnID = ColumnID
				.TableID = TableID
			End With

			CopyComponent = objValueCopy

			' Disassociate object variables.
			'UPGRADE_NOTE: Object objValueCopy may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objValueCopy = Nothing

		End Function

		Public ReadOnly Property ComponentType() As ExpressionComponentTypes
			Get
				Return ExpressionComponentTypes.giCOMPONENT_TABLEVALUE
			End Get
		End Property


		Public Property Value() As Object
			Get
				' Return the value property.
				Select Case miType
					Case ExpressionValueTypes.giEXPRVALUE_CHARACTER
						'UPGRADE_WARNING: Couldn't resolve default property of object Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Value = msCharacterValue
					Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
						'UPGRADE_WARNING: Couldn't resolve default property of object Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Value = mdblNumericValue
					Case ExpressionValueTypes.giEXPRVALUE_LOGIC
						'UPGRADE_WARNING: Couldn't resolve default property of object Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Value = mfLogicValue
					Case ExpressionValueTypes.giEXPRVALUE_DATE
						'UPGRADE_WARNING: Couldn't resolve default property of object mdtDateValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Value = mdtDateValue
					Case Else
						'UPGRADE_WARNING: Couldn't resolve default property of object Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Value = ""
				End Select

			End Get
			Set(ByVal Value As Object)
				' Set the value property.
				Select Case miType
					Case ExpressionValueTypes.giEXPRVALUE_CHARACTER
						'UPGRADE_WARNING: Couldn't resolve default property of object pvNewValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						msCharacterValue = Value.ToString()
					Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
						'UPGRADE_WARNING: Couldn't resolve default property of object pvNewValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mdblNumericValue = CDbl(Value)
					Case ExpressionValueTypes.giEXPRVALUE_LOGIC
						'UPGRADE_WARNING: Couldn't resolve default property of object pvNewValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mfLogicValue = CBool(Value)
					Case ExpressionValueTypes.giEXPRVALUE_DATE
						'UPGRADE_WARNING: Couldn't resolve default property of object pvNewValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object mdtDateValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mdtDateValue = Value
				End Select

			End Set
		End Property

		Public ReadOnly Property ComponentDescription() As String
			Get
				' Return the component description.
				Select Case miType
					Case ExpressionValueTypes.giEXPRVALUE_CHARACTER
						ComponentDescription = Chr(34) & msCharacterValue & Chr(34)
					Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
						'MH20010130 Fault 1610
						'ComponentDescription = Trim(Str(mdblNumericValue))
						ComponentDescription = ConvertNumberForDisplay(CStr(mdblNumericValue))
					Case ExpressionValueTypes.giEXPRVALUE_LOGIC
						ComponentDescription = IIf(mfLogicValue, "True", "False").ToString()
					Case ExpressionValueTypes.giEXPRVALUE_DATE
						'MH20010201 Fault 1576
						'ComponentDescription = Format(mdtDateValue, "Long Date")
						'UPGRADE_WARNING: Couldn't resolve default property of object mdtDateValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						ComponentDescription = IIf(IsDBNull(mdtDateValue), "Empty Date", VB6.Format(mdtDateValue, "Long Date")).ToString()
					Case Else
						ComponentDescription = ""
				End Select

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


		Public Property ReturnType() As ExpressionValueTypes
			Get
				Return miType
			End Get
			Set(ByVal Value As ExpressionValueTypes)
				miType = Value
			End Set
		End Property


		Public Property ColumnID() As Short
			Get

				ColumnID = miColumnID

			End Get
			Set(ByVal Value As Short)

				miColumnID = Value

			End Set
		End Property


		Public Property TableID() As Short
			Get

				TableID = miTableID

			End Get
			Set(ByVal Value As Short)

				miTableID = Value

			End Set
		End Property

	End Class
End Namespace