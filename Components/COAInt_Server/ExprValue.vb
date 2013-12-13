Option Strict Off
Option Explicit On

Imports HR.Intranet.Server.Enums

Friend Class clsExprValue
	
	' Component definition variables.
	Private miType As ExpressionValueTypes
	Private mdblNumericValue As Double
	Private msCharacterValue As String
	Private mfLogicValue As Boolean
	Private mdtDateValue As Object
	
	' Class handling variables.
	Private mobjBaseComponent As clsExprComponent
		
	Public Function ContainsExpression(ByRef plngExprID As Integer) As Boolean
		Return False
	End Function
	
	Public Function RuntimeCode(ByRef psRuntimeCode As String, ByRef palngSourceTables(,) As Integer, ByRef pfApplyPermissions As Boolean, ByRef pfValidating As Boolean, ByRef pavPromptedValues As Object, Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean
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
				sCode = IIf(mfLogicValue, "1", "0")
			Case ExpressionValueTypes.giEXPRVALUE_DATE
				'UPGRADE_WARNING: Couldn't resolve default property of object mdtDateValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				sCode = IIf(IsDBNull(mdtDateValue), "null", "convert(datetime, '" & VB6.Format(mdtDateValue, "MM/dd/yyyy") & "')")
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
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		Dim sSQL As String
		
		fOK = True
		
		'UPGRADE_WARNING: Couldn't resolve default property of object mdtDateValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		sSQL = "INSERT INTO ASRSysExprComponents" & " (componentID, exprID, type," & " valueType, valueCharacter, valueNumeric, valueLogic, valuedate)" & " VALUES(" & Trim(Str(mobjBaseComponent.ComponentID)) & "," & " " & Trim(Str(mobjBaseComponent.ParentExpression.ExpressionID)) & "," & " " & Trim(Str(ExpressionComponentTypes.giCOMPONENT_VALUE)) & "," & " " & Trim(Str(miType)) & "," & " '" & Replace(msCharacterValue, "'", "''") & "'," & " " & Trim(Str(mdblNumericValue)) & "," & " " & IIf(mfLogicValue, "1", "0") & "," & " " & IIf(IsDBNull(mdtDateValue), "null", "'" & VB6.Format(mdtDateValue, "MM/dd/yyyy") & "'") & ")"
		
		'MH20010201 Fault 1576
		'" '" & Format(mdtDateValue, "MM/dd/yyyy") & "')"
		
		gADOCon.Execute(sSQL,  , ADODB.CommandTypeEnum.adCmdText)
		
TidyUpAndExit: 
		'UPGRADE_WARNING: Couldn't resolve default property of object WriteComponent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		WriteComponent = fOK
		Exit Function
		
ErrorTrap: 
		'  If ASRDEVELOPMENT Then
		'    MsgBox Err.Description, vbCritical, "ASRDEVELOPMENT"
		'  End If
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	Public Function CopyComponent() As Object
		' Copies the selected component.
		' When editting a component we actually copy the component first
		' and edit the copy. If the changes are confirmed then the copy
		' replaces the original. If the changes are cancelled then the
		' copy is discarded.
		Dim objValueCopy As New clsExprValue
		
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
					msCharacterValue = pValue
				Case ExpressionValueTypes.giEXPRVALUE_NUMERIC
					'UPGRADE_WARNING: Couldn't resolve default property of object pvNewValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mdblNumericValue = pValue
				Case ExpressionValueTypes.giEXPRVALUE_LOGIC
					'UPGRADE_WARNING: Couldn't resolve default property of object pvNewValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					mfLogicValue = pValue
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
					ComponentDescription = datGeneral.ConvertNumberForDisplay(CStr(mdblNumericValue))
				Case ExpressionValueTypes.giEXPRVALUE_LOGIC
					ComponentDescription = IIf(mfLogicValue, "True", "False")
				Case ExpressionValueTypes.giEXPRVALUE_DATE
					'MH20010201 Fault 1576
					'ComponentDescription = Format(mdtDateValue, "Long Date")
					'UPGRADE_WARNING: Couldn't resolve default property of object mdtDateValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					ComponentDescription = IIf(IsDBNull(mdtDateValue), "Empty Date", VB6.Format(mdtDateValue, "Long Date"))
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
	
	Public Function UDFCode(ByRef psRuntimeCode() As String, ByRef palngSourceTables(,) As Integer, ByRef pfApplyPermissions As Boolean, ByRef pfValidating As Boolean, Optional ByRef plngFixedExprID As Integer = 0, Optional ByRef psFixedSQLCode As String = "") As Boolean
		Return True
	End Function

End Class