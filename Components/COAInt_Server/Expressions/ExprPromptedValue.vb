Option Strict Off
Option Explicit On

Imports System.Data.SqlClient
Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums

Imports VB = Microsoft.VisualBasic

Namespace Expressions
	Friend Class clsExprPromptedValue
		Inherits BaseExpressionComponent

		' Component definition variables.
		Private msPrompt As String
		Private miType As ExpressionValueTypes
		Private miReturnSize As Short
		Private miReturnDecimals As Short
		Private msFormat As String
		Private mlngLookupColumnID As Integer

		Private msDefaultCharacterValue As String
		Private mdblDefaultNumericValue As Double
		Private mfDefaultLogicValue As Boolean
		Private mdtDefaultDateValue As Date?
		Private miDefaultDateType As Short

		' Class handling variables.
		Private mobjBaseComponent As clsExprComponent

		Public Sub New(ByVal Value As SessionInfo)
			MyBase.New(Value)
		End Sub

		Public Function ContainsExpression(ByRef plngExprID As Integer) As Boolean
			Return False
		End Function

		Public Function RuntimeCode(ByRef psRuntimeCode As String, palngSourceTables(,) As Integer, pfApplyPermissions As Boolean _
																, pfValidating As Boolean, pavPromptedValues As Object _
																, psUDFs() As String _
																, Optional plngFixedExprID As Integer = 0, Optional psFixedSQLCode As String = "") As Boolean

			Dim fOK As Boolean = True
			Dim fFound As Boolean
			Dim iLoop As Short
			Dim sCode As String = ""

			Try

				' Do not display the prompt form if we are just validating the expression.
				If pfValidating Then
					Select Case ReturnType
						Case ExpressionValueTypes.giEXPRVALUE_CHARACTER
							sCode = "'" & gsDUMMY_CHARACTER & "'"
						Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
							sCode = Trim(Str(gsDUMMY_NUMERIC))
						Case ExpressionValueTypes.giEXPRVALUE_LOGIC
							sCode = IIf(gsDUMMY_LOGIC, "1", "0").ToString()
						Case ExpressionValueTypes.giEXPRVALUE_DATE
							sCode = "convert(datetime, '" & VB6.Format(CDate(gsDUMMY_DATE), "MM/dd/yyyy") & "')"
					End Select
				Else

					fFound = False
					For iLoop = 0 To UBound(pavPromptedValues, 2)
						'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If pavPromptedValues(0, iLoop) = mobjBaseComponent.ComponentID Then
							Select Case ReturnType
								Case ExpressionValueTypes.giEXPRVALUE_CHARACTER
									'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sCode = "'" & Replace(pavPromptedValues(1, iLoop), "'", "''") & "'"
								Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
									'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sCode = Trim(Str(pavPromptedValues(1, iLoop)))
								Case ExpressionValueTypes.giEXPRVALUE_LOGIC
									'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sCode = IIf(pavPromptedValues(1, iLoop), "1", "0")
								Case ExpressionValueTypes.giEXPRVALUE_DATE
									' JPD 20040212 Fault 8082 - DO NOT CONVERT DATE PROMPTED VALUES
									' THEY ARE PASSED IN FROM THE ASPs AS STRING VALUES IN THE CORRECT
									' FORMAT (mm/dd/yyyy) AND DOING ANY KIND OF CONVERSION JUST SCREWS
									' THINGS UP.
									'sCode = "convert(datetime, '" & Format(pavPromptedValues(1, iLoop), "MM/dd/yyyy") & "')"
									'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									sCode = "convert(datetime, '" & pavPromptedValues(1, iLoop) & "')"
							End Select

							fFound = True
						End If
					Next iLoop
					fOK = fFound

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

		Public Function WriteComponent() As Boolean

			Try
				DB.ExecuteSP("spASRIntSaveComponent", _
						New SqlParameter("componentID", SqlDbType.Int) With {.Value = mobjBaseComponent.ComponentID}, _
						New SqlParameter("expressionID", SqlDbType.Int) With {.Value = mobjBaseComponent.ParentExpression.ExpressionID}, _
						New SqlParameter("type", SqlDbType.TinyInt) With {.Value = ExpressionComponentTypes.giCOMPONENT_PROMPTEDVALUE}, _
						New SqlParameter("calculationID", SqlDbType.Int), _
						New SqlParameter("filterID", SqlDbType.Int), _
						New SqlParameter("functionID", SqlDbType.Int), _
						New SqlParameter("operatorID", SqlDbType.Int), _
						New SqlParameter("valueType", SqlDbType.TinyInt) With {.Value = miType}, _
						New SqlParameter("valueCharacter", SqlDbType.VarChar, 255) With {.Value = msDefaultCharacterValue}, _
						New SqlParameter("valueNumeric", SqlDbType.Float) With {.Value = mdblDefaultNumericValue}, _
						New SqlParameter("valueLogic", SqlDbType.Bit) With {.Value = mfDefaultLogicValue}, _
						New SqlParameter("valueDate", SqlDbType.DateTime) With {.Value = mdtDefaultDateValue}, _
						New SqlParameter("LookupTableID", SqlDbType.Int), _
						New SqlParameter("LookupColumnID", SqlDbType.Int), _
						New SqlParameter("fieldTableID", SqlDbType.Int), _
						New SqlParameter("fieldColumnID", SqlDbType.Int) With {.Value = mlngLookupColumnID}, _
						New SqlParameter("fieldPassBy", SqlDbType.TinyInt), _
						New SqlParameter("fieldSelectionRecord", SqlDbType.TinyInt), _
						New SqlParameter("fieldSelectionLine", SqlDbType.Int), _
						New SqlParameter("fieldSelectionOrderID", SqlDbType.Int), _
						New SqlParameter("fieldSelectionFilter", SqlDbType.Int), _
						New SqlParameter("promptDescription", SqlDbType.VarChar, 255) With {.Value = msPrompt}, _
						New SqlParameter("promptSize", SqlDbType.SmallInt) With {.Value = miReturnSize}, _
						New SqlParameter("promptDecimals", SqlDbType.SmallInt) With {.Value = miReturnDecimals}, _
						New SqlParameter("promptMask", SqlDbType.VarChar, 255) With {.Value = msFormat}, _
						New SqlParameter("promptDateType", SqlDbType.Int) With {.Value = miDefaultDateType})

				Return True

			Catch ex As Exception
				Return False

			End Try

		End Function

		Public Function CopyComponent() As Object
			' Copies the selected component.
			' When editing a component we actually copy the component first
			' and edit the copy. If the changes are confirmed then the copy
			' replaces the original. If the changes are cancelled then the
			' copy is discarded.
			Dim objPromptedValueCopy As New clsExprPromptedValue(SessionInfo)

			' Copy the component's basic properties.
			With objPromptedValueCopy
				.Prompt = msPrompt
				.ValueType = miType
				.ReturnSize = miReturnSize
				.ReturnDecimals = miReturnDecimals
				.ValueFormat = msFormat
				'UPGRADE_WARNING: Couldn't resolve default property of object DefaultValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.DefaultValue = DefaultValue
				.LookupColumn = mlngLookupColumnID
				.DefaultDateType = miDefaultDateType
			End With

			CopyComponent = objPromptedValueCopy

			' Disassociate object variables.
			'UPGRADE_NOTE: Object objPromptedValueCopy may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objPromptedValueCopy = Nothing

		End Function


		Public Property LookupColumn() As Integer
			Get
				' Return the Lookup Column ID.
				LookupColumn = mlngLookupColumnID

			End Get
			Set(ByVal Value As Integer)
				' Set the Lookup Column ID.
				mlngLookupColumnID = Value

			End Set
		End Property


		Public Property ValueFormat() As String
			Get
				' Return the ValueFormat property.
				ValueFormat = msFormat

			End Get
			Set(ByVal Value As String)
				' Set the ValueFormat property.
				msFormat = Value

			End Set
		End Property

		Public Property DefaultValue() As Object
			Get

				' Return the default value property.
				Select Case miType
					Case ExpressionValueTypes.giEXPRVALUE_CHARACTER
						'UPGRADE_WARNING: Couldn't resolve default property of object DefaultValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						DefaultValue = msDefaultCharacterValue
					Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
						'UPGRADE_WARNING: Couldn't resolve default property of object DefaultValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						DefaultValue = mdblDefaultNumericValue
					Case ExpressionValueTypes.giEXPRVALUE_LOGIC
						'UPGRADE_WARNING: Couldn't resolve default property of object DefaultValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						DefaultValue = mfDefaultLogicValue
					Case ExpressionValueTypes.giEXPRVALUE_DATE

						' What type of default date is it?
						Select Case miDefaultDateType
							Case 0
								'UPGRADE_WARNING: Couldn't resolve default property of object DefaultValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								DefaultValue = mdtDefaultDateValue
							Case 1
								'UPGRADE_WARNING: Couldn't resolve default property of object DefaultValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								DefaultValue = Now
							Case 2
								'UPGRADE_WARNING: Couldn't resolve default property of object DefaultValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								DefaultValue = DateTime.FromOADate(Now.ToOADate - VB.Day(Now) + 1)
							Case 3
								DefaultValue = DateSerial(Year(Now), Month(Now), DateSerial(Year(Now), Month(Now) + 1, 1).ToOADate - DateSerial(Year(Now), Month(Now), 1).ToOADate)
							Case 4
								DefaultValue = DateSerial(Year(Now), 1, 1)
							Case 5
								DefaultValue = DateSerial(Year(Now), 12, 31)
						End Select

					Case ExpressionValueTypes.giEXPRVALUE_TABLEVALUE
						'UPGRADE_WARNING: Couldn't resolve default property of object DefaultValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						DefaultValue = msDefaultCharacterValue
					Case Else
						'UPGRADE_WARNING: Couldn't resolve default property of object DefaultValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						DefaultValue = ""
				End Select

			End Get

			Set(ByVal Value As Object)
				' Set the value property.
				Dim dtdummydate As Date

				Select Case miType
					Case ExpressionValueTypes.giEXPRVALUE_CHARACTER
						'UPGRADE_WARNING: Couldn't resolve default property of object pvNewValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						msDefaultCharacterValue = Value.ToString()
					Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
						'UPGRADE_WARNING: Couldn't resolve default property of object pvNewValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mdblDefaultNumericValue = CDbl(Value)
					Case ExpressionValueTypes.giEXPRVALUE_LOGIC
						'UPGRADE_WARNING: Couldn't resolve default property of object pvNewValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						mfDefaultLogicValue = CBool(Value)
					Case ExpressionValueTypes.giEXPRVALUE_DATE
						'UPGRADE_WARNING: Couldn't resolve default property of object pvNewValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						If IsDBNull(Value) Or (Value = dtdummydate) Then
							mdtDefaultDateValue = DateTime.FromOADate(0)
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object pvNewValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							mdtDefaultDateValue = Value
						End If
					Case ExpressionValueTypes.giEXPRVALUE_TABLEVALUE
						'UPGRADE_WARNING: Couldn't resolve default property of object pvNewValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						msDefaultCharacterValue = Value.ToString()
				End Select

			End Set
		End Property

		Public Property DefaultDateType() As Short
			Get
				DefaultDateType = miDefaultDateType
			End Get
			Set(ByVal Value As Short)
				miDefaultDateType = Value
			End Set
		End Property

		Public ReadOnly Property ComponentType() As ExpressionComponentTypes
			Get
				Return ExpressionComponentTypes.giCOMPONENT_PROMPTEDVALUE
			End Get
		End Property

		Public Property ReturnType() As ExpressionValueTypes
			Get

				Dim iType As ExpressionValueTypes = ExpressionValueTypes.giEXPRVALUE_UNDEFINED

				Select Case miType
					Case ExpressionValueTypes.giEXPRVALUE_CHARACTER
						iType = ExpressionValueTypes.giEXPRVALUE_CHARACTER

					Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
						iType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

					Case ExpressionValueTypes.giEXPRVALUE_LOGIC
						iType = ExpressionValueTypes.giEXPRVALUE_LOGIC

					Case ExpressionValueTypes.giEXPRVALUE_DATE
						iType = ExpressionValueTypes.giEXPRVALUE_DATE

					Case ExpressionValueTypes.giEXPRVALUE_TABLEVALUE

						Select Case Columns.GetById(mlngLookupColumnID).DataType
							Case ColumnDataType.sqlNumeric, ColumnDataType.sqlInteger
								iType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
							Case ColumnDataType.sqlDate
								iType = ExpressionValueTypes.giEXPRVALUE_DATE
							Case ColumnDataType.sqlVarChar, ColumnDataType.sqlLongVarChar
								iType = ExpressionValueTypes.giEXPRVALUE_CHARACTER
							Case ColumnDataType.sqlBoolean
								iType = ExpressionValueTypes.giEXPRVALUE_LOGIC
							Case ColumnDataType.sqlOle
								iType = ExpressionValueTypes.giEXPRVALUE_OLE
							Case ColumnDataType.sqlVarBinary
								iType = ExpressionValueTypes.giEXPRVALUE_PHOTO
						End Select

				End Select

				Return iType


			End Get

			Set(ByVal Value As ExpressionValueTypes)
				' Set the return type.
				miType = Value

			End Set
		End Property

		Public Property ReturnDecimals() As Short
			Get
				' Return the return number of decimals.
				ReturnDecimals = miReturnDecimals

			End Get
			Set(ByVal Value As Short)
				' Set the return number of decimals.
				miReturnDecimals = Value

			End Set
		End Property



		Public Property ValueType() As ExpressionValueTypes
			Get
				' Return the type property.
				ValueType = miType

			End Get
			Set(ByVal Value As ExpressionValueTypes)
				' Set the type property.
				miType = Value

			End Set
		End Property


		Public Property ReturnSize() As Short
			Get
				' Return the return size.
				ReturnSize = miReturnSize

			End Get
			Set(ByVal Value As Short)
				' Set the return size.
				miReturnSize = Value

			End Set
		End Property

		Public ReadOnly Property ComponentDescription() As String
			Get
				' Return the component description.
				Dim sDescription As String

				sDescription = msPrompt & " : "

				Select Case miType
					Case ExpressionValueTypes.giEXPRVALUE_CHARACTER
						sDescription = sDescription & "<string>"

					Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
						sDescription = sDescription & "<numeric>"

					Case ExpressionValueTypes.giEXPRVALUE_LOGIC
						sDescription = sDescription & "<logic>"

					Case ExpressionValueTypes.giEXPRVALUE_DATE
						sDescription = sDescription & "<date>"

					Case ExpressionValueTypes.giEXPRVALUE_TABLEVALUE
						sDescription = sDescription & "<table value>"
				End Select

				ComponentDescription = sDescription

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

		Public Property Prompt() As String
			Get
				' Return the Prompt property.
				Prompt = msPrompt

			End Get
			Set(ByVal Value As String)
				' Set the Prompt property.
				msPrompt = Value

			End Set
		End Property

		Public Function UDFCode(ByRef psRuntimeCode() As String, ByRef palngSourceTables(,) As Integer, ByRef pfApplyPermissions As Boolean, ByRef pfValidating As Boolean, Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean

			Return True

		End Function

	End Class
End Namespace