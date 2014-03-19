Option Strict On
Option Explicit On

Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums
Imports HR.Intranet.Server.Structures

Friend Class clsExprValue
	Inherits BaseExpressionComponent

	' Component definition variables.
	Private miType As ExpressionValueTypes
	Private mdblNumericValue As Double
	Private msCharacterValue As String
	Private mfLogicValue As Boolean
	Private mdtDateValue As Object
	
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
		' Return the SQL code for the component.
		On Error GoTo ErrorTrap

		Dim fOK As Boolean = True
		Dim sCode As String = ""

		Select Case miType
			Case ExpressionValueTypes.giEXPRVALUE_CHARACTER
				sCode = "'" & Replace(msCharacterValue, "'", "''") & "'"
			Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
				sCode = Trim(Str(mdblNumericValue))
			Case ExpressionValueTypes.giEXPRVALUE_LOGIC
				sCode = IIf(mfLogicValue, "1", "0").ToString()
			Case ExpressionValueTypes.giEXPRVALUE_DATE
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

		Return fOK

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

	Public Function WriteComponent() As Object
		' Write the component definition to the component recordset.
		Dim sSQL As String

		Try
			sSQL = "INSERT INTO ASRSysExprComponents (componentID, exprID, type, valueType, valueCharacter, valueNumeric, valueLogic, valuedate)" _
				& " VALUES(" & Trim(Str(mobjBaseComponent.ComponentID)) & ", " & Trim(Str(mobjBaseComponent.ParentExpression.ExpressionID)) & ", " _
				& Trim(Str(ExpressionComponentTypes.giCOMPONENT_VALUE)) & ", " & Trim(Str(miType)) & ", '" & Replace(msCharacterValue, "'", "''") & "', " & Trim(Str(mdblNumericValue)) & ", " _
				& IIf(mfLogicValue, "1", "0").ToString() & ", " & IIf(IsDBNull(mdtDateValue), "null", "'" & VB6.Format(mdtDateValue, "MM/dd/yyyy") & "'").ToString() & ")"
			DB.ExecuteSql(sSQL)
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
		Dim objValueCopy As New clsExprValue(SessionInfo)

		' Copy the component's basic properties.
		With objValueCopy
			.ReturnType = miType
			'UPGRADE_WARNING: Couldn't resolve default property of object Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.Value = Value
		End With

		CopyComponent = objValueCopy

		' Disassociate object variables.
		'UPGRADE_NOTE: Object objValueCopy may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objValueCopy = Nothing

	End Function

	Public ReadOnly Property ComponentType() As ExpressionComponentTypes
		Get
			Return ExpressionComponentTypes.giCOMPONENT_VALUE
		End Get
	End Property


	Public Property Value() As Object
		Get
			' Return the value property.
			Select Case miType
				Case ExpressionValueTypes.giEXPRVALUE_CHARACTER
					Return msCharacterValue
				Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
					Return mdblNumericValue
				Case ExpressionValueTypes.giEXPRVALUE_LOGIC
					Return mfLogicValue
				Case ExpressionValueTypes.giEXPRVALUE_DATE
					Return mdtDateValue
				Case Else
					Return ""
			End Select

		End Get

		Set(ByVal pValue As Object)
			' Set the value property.
			Select Case miType
				Case ExpressionValueTypes.giEXPRVALUE_CHARACTER
					'UPGRADE_WARNING: Couldn't resolve default property of object pvNewValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					msCharacterValue = pValue.ToString()
				Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
					'UPGRADE_WARNING: Couldn't resolve default property of object pvNewValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mdblNumericValue = CDbl(pValue)
				Case ExpressionValueTypes.giEXPRVALUE_LOGIC
					'UPGRADE_WARNING: Couldn't resolve default property of object pvNewValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mfLogicValue = CBool(pValue)
				Case ExpressionValueTypes.giEXPRVALUE_DATE
					'UPGRADE_WARNING: Couldn't resolve default property of object pvNewValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object mdtDateValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mdtDateValue = pValue
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
					ComponentDescription = ConvertNumberForDisplay(CStr(mdblNumericValue))
				Case ExpressionValueTypes.giEXPRVALUE_LOGIC
					ComponentDescription = IIf(mfLogicValue, "True", "False").ToString()
				Case ExpressionValueTypes.giEXPRVALUE_DATE
					ComponentDescription = IIf(IsDBNull(mdtDateValue), "Empty Date", VB6.Format(mdtDateValue, "Long Date")).ToString()
				Case Else
					ComponentDescription = ""
			End Select

		End Get
	End Property

	Public Property BaseComponent() As clsExprComponent
		Get
			Return mobjBaseComponent
		End Get
		Set(ByVal pValue As clsExprComponent)
			mobjBaseComponent = pValue
		End Set
	End Property

	Public Property ReturnType() As ExpressionValueTypes
		Get
			Return miType
		End Get
		Set(ByVal pValue As ExpressionValueTypes)
			miType = pValue
		End Set
	End Property

End Class