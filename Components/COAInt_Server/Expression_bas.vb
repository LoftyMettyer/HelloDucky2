Option Strict Off
Option Explicit On
Module modExpression
	
	' Expression Type constants
	' NB. If you modify this enum, you'll need to do the same to the mathcing enums in:
	'     System Manager - Application
	'     Security Manager - modExpression
	'     Data Manager - modExpression
	Public Enum ExpressionTypes
		giEXPR_UNKNOWNTYPE = 0
		giEXPR_COLUMNCALCULATION = 1
		giEXPR_GOTFOCUS = 2 ' Not used.
		giEXPR_RECORDVALIDATION = 3
		giEXPR_DEFAULTVALUE = 4 ' Not used.
		giEXPR_STATICFILTER = 5
		giEXPR_PAGEBREAK = 6 ' Not used.
		giEXPR_ORDER = 7 ' Not used.
		giEXPR_RECORDDESCRIPTION = 8
		giEXPR_VIEWFILTER = 9
		giEXPR_RUNTIMECALCULATION = 10
		giEXPR_RUNTIMEFILTER = 11
		giEXPR_EMAIL = 12 ' System Manager Only
		giEXPR_LINKFILTER = 13 ' System Manager Only
		giEXPR_UTILRUNTIMEFILTER = 14 'Import filter
		giEXPR_MATCHJOINEXPRESSION = 15
		giEXPR_MATCHSCOREEXPRESSION = 16
		giEXPR_MATCHWHEREEXPRESSION = 17
		giEXPR_RECORDINDEPENDANTCALC = 18
		giEXPR_OUTLOOKFOLDER = 19 'System Manager Only
		giEXPR_OUTLOOKSUBJECT = 20 'System Manager Only
		giEXPR_WORKFLOWCALCULATION = 21 'System Manager Only
		giEXPR_WORKFLOWSTATICFILTER = 22 'System Manager Only
		giEXPR_WORKFLOWRUNTIMEFILTER = 23 'System Manager Only
	End Enum
	
	Public Enum AccessCodes
		giACCESS_READWRITE = 0
		giAccess_READONLY = 1
		giACCESS_HIDDEN = 2
	End Enum
	
	' Expression Value types
	' NB. If you modify this enum, you'll need to do the same to the mathcing enums in:
	'     System Manager - Application
	'     Security Manager - modExpression
	'     Data Manager - modExpression
	Public Enum ExpressionComponentTypes
		giCOMPONENT_FIELD = 1
		giCOMPONENT_FUNCTION = 2
		giCOMPONENT_CALCULATION = 3
		giCOMPONENT_VALUE = 4
		giCOMPONENT_OPERATOR = 5
		giCOMPONENT_TABLEVALUE = 6
		giCOMPONENT_PROMPTEDVALUE = 7
		giCOMPONENT_CUSTOMCALC = 8 ' Not used.
		giCOMPONENT_EXPRESSION = 9
		giCOMPONENT_FILTER = 10
		giCOMPONENT_WORKFLOWVALUE = 11
		giCOMPONENT_WORKFLOWFIELD = 12
	End Enum
	
	Public Enum FieldSelectionTypes
		giSELECT_FIRSTRECORD = 1
		giSELECT_LASTRECORD = 2
		giSELECT_SPECIFICRECORD = 3
		giSELECT_RECORDTOTAL = 4
		giSELECT_RECORDCOUNT = 5
	End Enum
	
	Public Enum FieldPassTypes
		giPASSBY_VALUE = 1
		giPASSBY_REFERENCE = 2
	End Enum
	
	Public Const giPRINT_XINDENT As Short = 1000
	Public Const giPRINT_YINDENT As Short = 1000
	Public Const giPRINT_XSPACE As Short = 500
	Public Const giPRINT_YSPACE As Short = 100
	
	Public Enum ExprValidationCodes
		giEXPRVALIDATION_NOERRORS = 0
		giEXPRVALIDATION_MISSINGOPERAND = 1
		giEXPRVALIDATION_SYNTAXERROR = 2
		giEXPRVALIDATION_EXPRTYPEMISMATCH = 3
		giEXPRVALIDATION_UNKNOWNERROR = 4
		giEXPRVALIDATION_OPERANDTYPEMISMATCH = 5
		giEXPRVALIDATION_PARAMETERTYPEMISMATCH = 6
		giEXPRVALIDATION_NOCOMPONENTS = 7
		giEXPRVALIDATION_PARAMETERSYNTAXERROR = 8
		giEXPRVALIDATION_PARAMETERNOCOMPONENTS = 9
		giEXPRVALIDATION_FILTEREVALUATION = 10
		giEXPRVALIDATION_SQLERROR = 11 ' JPD20020419 Fault 3687
		giEXPRVALIDATION_ASSOCSQLERROR = 12 ' JPD20020419 Fault 3687
		giEXPRVALIDATION_CYCLIC = 13 'JPD 20040507 Fault 8600
	End Enum
	
	Public Enum ExpressionValueTypes
		giEXPRVALUE_UNDEFINED = 0
		giEXPRVALUE_CHARACTER = 1
		giEXPRVALUE_NUMERIC = 2
		giEXPRVALUE_LOGIC = 3
		giEXPRVALUE_DATE = 4
		giEXPRVALUE_TABLEVALUE = 5
		giEXPRVALUE_OLE = 6
		giEXPRVALUE_PHOTO = 7
		giEXPRVALUE_BYREF_UNDEFINED = 100
		giEXPRVALUE_BYREF_CHARACTER = 101
		giEXPRVALUE_BYREF_NUMERIC = 102
		giEXPRVALUE_BYREF_LOGIC = 103
		giEXPRVALUE_BYREF_DATE = 104
		giEXPRVALUE_BYREF_TABLEVALUE = 105 ' Not used.
		giEXPRVALUE_BYREF_OLE = 106 ' Not used.
		giEXPRVALUE_BYREF_PHOTO = 107 ' Not used.
	End Enum
	Public Const giEXPRVALUE_BYREF_OFFSET As Short = 100
	
	Public Const gsDUMMY_CHARACTER As String = "ASRDUMMYCHARVALUE"
	Public Const gsDUMMY_NUMERIC As Short = 1
	Public Const gsDUMMY_LOGIC As Boolean = True
	Public Const gsDUMMY_DATE As Date = #1/1/1998#
  Public Const gsDUMMY_BYREF_CHARACTER As String = "12" & vbTab & "a"
  Public Const gsDUMMY_BYREF_NUMERIC As String = "2" & vbTab & "1"
  Public Const gsDUMMY_BYREF_LOGIC As String = "-7" & vbTab & "0"
  Public Const gsDUMMY_BYREF_DATE As String = "11" & vbTab & "1/1/1998"
	
	' Order object constants.
	Public Enum OrderTypes
		giORDERTYPE_STATIC = 0
		giORDERTYPE_DYNAMIC = 1
	End Enum
	
	' Parameter Type constants.
	Public Const gsPARAMETERTYPE_ORDERID As String = "PType_OrderID"
	
	
	
	
	
	
	
	
	
	
	
	Public Function ExprDeleted(ByRef lngExprID As Object) As Boolean
		
		Dim rsExprTemp As ADODB.Recordset
		Dim sSQL As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object lngExprID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sSQL = "SELECT * FROM ASRSysExpressions WHERE ExprID = " & lngExprID
		
		rsExprTemp = datGeneral.GetRecords(sSQL)
		
		With rsExprTemp
			If .BOF And .EOF Then ExprDeleted = True
			.Close()
		End With
		
		sSQL = vbNullString
		'UPGRADE_NOTE: Object rsExprTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsExprTemp = Nothing
		
	End Function
	Public Function ExpressionTypeName(ByRef piType As Short) As String
		' Return the description of the expression type.
		Select Case piType
			Case ExpressionTypes.giEXPR_COLUMNCALCULATION
				ExpressionTypeName = "Column Calculation"
				
			Case ExpressionTypes.giEXPR_GOTFOCUS ' NOT USED.
				ExpressionTypeName = "Field Entry Validation Clause"
				
			Case ExpressionTypes.giEXPR_RECORDVALIDATION
				ExpressionTypeName = "Field Validation"
				
			Case ExpressionTypes.giEXPR_DEFAULTVALUE ' NOT USED.
				ExpressionTypeName = "Default Value"
				
			Case ExpressionTypes.giEXPR_STATICFILTER
				ExpressionTypeName = "Filter"
				
			Case ExpressionTypes.giEXPR_PAGEBREAK ' NOT USED.
				ExpressionTypeName = "Page Break"
				
			Case ExpressionTypes.giEXPR_ORDER ' NOT USED.
				ExpressionTypeName = "Order"
				
			Case ExpressionTypes.giEXPR_RECORDDESCRIPTION
				ExpressionTypeName = "Record Description"
				
			Case ExpressionTypes.giEXPR_VIEWFILTER
				ExpressionTypeName = "View Filter"
				
			Case ExpressionTypes.giEXPR_RUNTIMECALCULATION
				ExpressionTypeName = "Runtime Calculation"
				
			Case ExpressionTypes.giEXPR_RUNTIMEFILTER
				ExpressionTypeName = "Filter"
				
			Case Else
				ExpressionTypeName = "Expression"
		End Select
		
	End Function
	
	
	
	
	
	Public Function UniqueColumnValue(ByRef sTableName As String, ByRef sColumnName As String) As Integer
		On Error GoTo ErrorTrap
		
		Dim lngUniqueValue As Integer
		Dim sSQL As String
		Dim rsUniqueValue As ADODB.Recordset
		
		' Create a record set with a unique value for the given table and column.
		sSQL = "SELECT MAX(" & sColumnName & ") + 1 AS newValue" & " FROM " & sTableName
		rsUniqueValue = datGeneral.GetRecords(sSQL)
		With rsUniqueValue
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If IsDbNull(.Fields("newValue").Value) Then
				lngUniqueValue = 1
			Else
				lngUniqueValue = .Fields("newValue").Value
			End If
			
			.Close()
		End With
		
TidyUpAndExit: 
		' Disassociate object variables.
		'UPGRADE_NOTE: Object rsUniqueValue may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsUniqueValue = Nothing
		'Return the unique column value.
		UniqueColumnValue = lngUniqueValue
		Exit Function
		
ErrorTrap: 
		lngUniqueValue = 1
		Resume TidyUpAndExit
		
	End Function
	
	
	Public Function ValidNameChar(ByVal piAsciiCode As Short, ByVal piPosition As Short) As Short
		' Validate the characters used to create table and column names.
		On Error GoTo ErrorTrap
		
		If piAsciiCode = Asc(" ") Then
			' Substitute underscores for spaces.
			If piPosition <> 0 Then
				piAsciiCode = Asc("_")
			Else
				piAsciiCode = 0
			End If
		Else
			' Allow only pure alpha-numerics and underscores.
			' Do not allow numerics in the first chracter position.
			'    If Not (piAsciiCode = 8 Or piAsciiCode = Asc("_") Or _
			''      (piAsciiCode >= Asc("0") And piAsciiCode <= Asc("9") And piPosition <> 0) Or _
			''      (piAsciiCode >= Asc("A") And piAsciiCode <= Asc("Z")) Or _
			''      (piAsciiCode >= Asc("a") And piAsciiCode <= Asc("z"))) Then
			'      piAsciiCode = 0
			'    End If
			'  End If
			
			' RH 15/08/2000 - BUG...we should be able to start filter/calcs with a number char
			If Not (piAsciiCode = 8 Or piAsciiCode = Asc("_") Or (piAsciiCode >= Asc("0") And piAsciiCode <= Asc("9")) Or (piAsciiCode >= Asc("A") And piAsciiCode <= Asc("Z")) Or (piAsciiCode >= Asc("a") And piAsciiCode <= Asc("z"))) Then
				piAsciiCode = 0
			End If
		End If
		
		ValidNameChar = piAsciiCode
		Exit Function
		
ErrorTrap: 
		ValidNameChar = 0
		Err.Number = False
		
	End Function
	
	
	
	
	
	
	
	Public Function ValidateOperatorParameters(ByRef plngOperatorID As Integer, ByRef piResultType As ExpressionValueTypes, ByRef piParam1Type As Short, ByRef piParam2Type As Short) As Boolean
		' Validate the given operator with the given parameters.
		' Return the result type in the piResultType parameter.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		
		fOK = True
		
		' Validate the parameter types for the given operator.
		Select Case plngOperatorID
			Case 1 ' PLUS
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 2 ' MINUS
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 3 ' TIMES BY
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 4 ' DIVIDED BY
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 5 ' AND
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_LOGIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_LOGIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC
				
			Case 6 ' OR
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_LOGIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_LOGIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC
				
			Case 7 ' IS EQUAL TO
				fOK = ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_LOGIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_LOGIC)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER))
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC
				
			Case 8 ' IS NOT EQUAL TO
				fOK = ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_LOGIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_LOGIC)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER))
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC
				
			Case 9 ' IS LESS THAN
				fOK = ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER))
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC
				
			Case 10 ' IS GREATER THAN
				fOK = ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER))
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC
				
			Case 11 ' IS LESS THAN OR EQUAL TO
				fOK = ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER))
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC
				
			Case 12 ' IS GREATER THAN OR EQUAL TO
				fOK = ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER))
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC
				
			Case 13 ' NOT
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_LOGIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC
				
			Case 14 ' IS CONTAINED WITHIN
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC
				
			Case 15 ' TO THE POWER OF
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 16 ' MODULAS
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 17 ' CONCATENATED WITH
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER
				
			Case Else ' Unknown operator
				fOK = False
		End Select
		
TidyUpAndExit: 
		If Not fOK Then
			piResultType = ExpressionTypes.giEXPR_UNKNOWNTYPE
		End If
		
		ValidateOperatorParameters = fOK
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	Public Function ValidateFunctionParameters(ByRef plngFunctionID As Object, ByRef piResultType As ExpressionValueTypes, Optional ByRef piParam1Type As Short = 0, Optional ByRef piParam2Type As Short = 0, Optional ByRef piParam3Type As Short = 0, Optional ByRef piParam4Type As Short = 0, Optional ByRef piParam5Type As Short = 0, Optional ByRef piParam6Type As Short = 0) As Boolean
		' Validate the given function with the given parameters.
		' Return the result type in the piResultType parameter.
		On Error GoTo ErrorTrap
		
		Dim fOK As Boolean
		
		fOK = True
		
		' Get the parameter types.
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		piParam1Type = IIf(IsNothing(piParam1Type), ExpressionValueTypes.giEXPRVALUE_UNDEFINED, piParam1Type)
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		piParam2Type = IIf(IsNothing(piParam2Type), ExpressionValueTypes.giEXPRVALUE_UNDEFINED, piParam2Type)
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		piParam3Type = IIf(IsNothing(piParam3Type), ExpressionValueTypes.giEXPRVALUE_UNDEFINED, piParam3Type)
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		piParam4Type = IIf(IsNothing(piParam4Type), ExpressionValueTypes.giEXPRVALUE_UNDEFINED, piParam4Type)
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		piParam5Type = IIf(IsNothing(piParam5Type), ExpressionValueTypes.giEXPRVALUE_UNDEFINED, piParam5Type)
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		piParam6Type = IIf(IsNothing(piParam6Type), ExpressionValueTypes.giEXPRVALUE_UNDEFINED, piParam6Type)
		
		' Validate the parameter types for the given function.
		Select Case plngFunctionID
			Case 1 ' SYSTEM DATE
				fOK = True
				piResultType = ExpressionValueTypes.giEXPRVALUE_DATE
				
			Case 2 ' CONVERT TO UPPERCASE
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER
				
			Case 3 ' CONVERT NUMERIC TO STRING
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER
				
			Case 4 ' IF, THEN, ELSE
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_LOGIC) And (piParam2Type = piParam3Type)
				piResultType = piParam2Type
				
			Case 5 ' REMOVE LEADING AND TRAINING SPACES
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER
				
			Case 6 ' EXTRACT CHARACTERS FROM THE LEFT
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER
				
			Case 7 ' LENGTH OF STRING
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 8 ' CONVERT TO LOWERCASE
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER
				
			Case 9 ' MAXIMUM
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 10 ' MINIMUM
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 11 ' SEARCH FOR CHARACTER STRING
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 12 ' CAPITALIZE INITIALS
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER
				
			Case 13 'EXTRACT CHARACTERS FROM THE RIGHT
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER
				
			Case 14 ' EXTRACT PART OF A STRING
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam3Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER
				
			Case 15 ' SYSTEM TIME
				fOK = True
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER
				
			Case 16 ' IS FIELD EMPTY
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) Or (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) Or (piParam1Type = ExpressionValueTypes.giEXPRVALUE_LOGIC) Or (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC
				
			Case 17 ' CURRENT USER
				fOK = True
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER
				
			Case 18 ' WHOLE YEARS UNTIL CURRENT DATE
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 19 ' REMAINING MONTHS SINCE WHOLE YEARS
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 20 ' INITIALS FROM FORENAMES
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER
				
			Case 21 ' FIRST NAME FROM FORENAMES
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER
				
			Case 22 ' WEEKDAYS FROM START AND END DATES
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 23 ' ADD MONTHS TO DATE
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_DATE
				
			Case 24 ' ADD YEARS TO DATE
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_DATE
				
			Case 25 ' CONVERT CHARACTER TO NUMERIC
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 26 ' WHOLE MONTHS BETWEEN TWO DATES
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 27 ' PARENTHESESES
				fOK = True
				piResultType = piParam1Type
				
			Case 28 ' DAY OF THE WEEK
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 29 ' NUMBER OF WORKING DAYS PER WEEK
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 30 ' ABSENCE DURATION
				'TM08102003
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam3Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam4Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 31 ' ROUND DOWN TO NEAREST WHOLE NUMBER
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 32 ' YEAR OF DATE
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 33 ' MONTH OF DATE
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 34 ' DAY OF DATE
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 35 ' NICE DATE
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER
				
			Case 36 ' NICE TIME
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER
				
			Case 37 ' ROUND DATE TO START OF NEAREST MONTH
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_DATE
				
			Case 38 ' IS BETWEEN
				fOK = ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam3Type = ExpressionValueTypes.giEXPRVALUE_DATE)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam3Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC))
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC
				
			Case 39 ' SERVICE YEARS
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 40 ' SERVICE MONTHS
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 41 ' STATUTORY REDUNDANCY PAY
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam3Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam4Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam5Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 42 ' GET FIELD FROM DATABASE RECORD
				fOK = (((piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_LOGIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_LOGIC)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE))) And ((piParam3Type = ExpressionValueTypes.giEXPRVALUE_BYREF_CHARACTER) Or (piParam3Type = ExpressionValueTypes.giEXPRVALUE_BYREF_NUMERIC) Or (piParam3Type = ExpressionValueTypes.giEXPRVALUE_BYREF_LOGIC) Or (piParam3Type = ExpressionValueTypes.giEXPRVALUE_BYREF_DATE))
				piResultType = (piParam3Type - giEXPRVALUE_BYREF_OFFSET)
				
			Case 43 ' GET UNIQUE CODE
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 44 ' ADD DAYS TO DATE
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_DATE
				
			Case 45 ' DAYS BETWEEN TWO DATES
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 46 ' WORKING DAYS BETWEEN TWO DATES (INC BHOLS)
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
				'JPD 20031017 Fault 7269
			Case 47 ' ABSENCE BETWEEN TWO DATES
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam3Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 48 ' ROUND UP TO NEAREST WHOLE NUMBER
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 49 ' ROUND TO NEAREST NUMBER
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 51 ' CONVERT CURRENCY
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam3Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 52 ' Field Last Changed Date
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_CHARACTER Or piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_NUMERIC Or piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_LOGIC Or piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_DATE
				
			Case 53 ' Field changed between two dates
				fOK = ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_CHARACTER Or piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_NUMERIC Or piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_LOGIC Or piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_DATE)) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC
				
			Case 54 'Whole years between two dates
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
				' JPD20021121 Fault 3177
			Case 55 ' First Day of Month - VERSION 2 FUNCTION
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_DATE
				
				' JPD20021121 Fault 3177
			Case 56 ' Last Day of Month - VERSION 2 FUNCTION
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_DATE
				
				' JPD20021121 Fault 3177
			Case 57 ' First Day of Year - VERSION 2 FUNCTION
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_DATE
				
				' JPD20021121 Fault 3177
			Case 58 ' Last Day of Year - VERSION 2 FUNCTION
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_DATE
				
				' JPD20021129 Fault 4337
			Case 59 ' NAME OF MONTH
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER
				
				' JPD20021129 Fault 4337
			Case 60 ' NAME OF DAY
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER
				
				' JPD20021129 Fault 3606
			Case 61 ' IS FIELD POPULATED
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) Or (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) Or (piParam1Type = ExpressionValueTypes.giEXPRVALUE_LOGIC) Or (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC
				
			Case 65 'IS POST SUBORDINATE OF
				fOK = False
				'Select Case IdentifyingColumnDataType
				'  Case sqlNumeric, sqlInteger
				'    fOK = (piParam1Type = giEXPRVALUE_NUMERIC)
				'  Case Else
				'    fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
				'End Select
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC
				
			Case 66 'IS POST SUBORDINATE OF USER
				fOK = True
				'fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
				''      (piParam2Type = giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC
				
			Case 67 'IS PERSONNEL SUBORDINATE OF
				fOK = False
				'Select Case IdentifyingColumnDataType
				'  Case sqlNumeric, sqlInteger
				'    fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
				''      (piParam2Type = giEXPRVALUE_DATE)
				'  Case Else
				'    fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
				''      (piParam2Type = giEXPRVALUE_DATE)
				'End Select
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC
				
			Case 68 'IS PERSONNEL SUBORDINATE OF USER
				fOK = True
				'fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
				''      (piParam2Type = giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC
				
			Case 69 'HAS POST SUBORDINATE
				fOK = False
				'Select Case IdentifyingColumnDataType
				'  Case sqlNumeric, sqlInteger
				'    fOK = (piParam1Type = giEXPRVALUE_NUMERIC)
				'  Case Else
				'    fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
				'End Select
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC
				
			Case 70 'HAS POST SUBORDINATE USER
				fOK = True
				'fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
				''      (piParam2Type = giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC
				
			Case 71 'HAS PERSONNEL SUBORDINATE
				fOK = False
				'Select Case IdentifyingColumnDataType
				'  Case sqlNumeric, sqlInteger
				'    fOK = (piParam1Type = giEXPRVALUE_NUMERIC) And _
				''      (piParam2Type = giEXPRVALUE_DATE)
				'  Case Else
				'    fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
				''      (piParam2Type = giEXPRVALUE_DATE)
				'End Select
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC
				
			Case 72 'HAS PERSONNEL SUBORDINATE USER
				fOK = True
				'fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
				''      (piParam2Type = giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC
				
			Case 73 'BRADFORD FACTOR
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam3Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC
				
			Case 77 ' Replace Characters within a Strin
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam3Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER
				
			Case Else ' Unknown function
				fOK = False
		End Select
		
TidyUpAndExit: 
		If Not fOK Then
			piResultType = ExpressionTypes.giEXPR_UNKNOWNTYPE
		End If
		
		ValidateFunctionParameters = fOK
		Exit Function
		
ErrorTrap: 
		fOK = False
		Resume TidyUpAndExit
		
	End Function
	
	
	Public Function IsFilterValid(ByRef varID As Object) As String
		
		Dim objExpr As clsExprExpression
		Dim strRuntimeCode As String
		Dim strFilterName As String
    Dim avDummyPrompts(,) As Object
		
		On Error GoTo LocalErr
		
		ReDim avDummyPrompts(1, 0)
		
		strFilterName = vbNullString
		IsFilterValid = IsSelectionValid(varID, "filter")
		
		If IsFilterValid = vbNullString Then
			objExpr = New clsExprExpression
			With objExpr
				'JPD 20030324 Fault 5161
				'UPGRADE_WARNING: Couldn't resolve default property of object varID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.ExpressionID = CInt(varID)
				.ConstructExpression()
				If (.ValidateExpression(True) <> ExprValidationCodes.giEXPRVALIDATION_NOERRORS) Then
					IsFilterValid = "The filter '" & strFilterName & "' used in this definition is invalid."
				End If
				
				'      If .Initialise(0, CLng(varID), giEXPR_RUNTIMEFILTER, giEXPRVALUE_LOGIC) Then
				'        strFilterName = .Name
				'        If objExpr.RuntimeFilterCode(strRunTimeCode, True, True, avDummyPrompts) Then
				'          datGeneral.GetReadOnlyRecords strRunTimeCode
				'        End If
				'      End If
				
			End With
			'UPGRADE_NOTE: Object objExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			objExpr = Nothing
		End If
		
		Exit Function
		
LocalErr: 
		If strFilterName <> vbNullString Then
			IsFilterValid = "'" & strFilterName & "' "
		End If
		IsFilterValid = "The filter " & IsFilterValid & "used in this definition is invalid"
		
	End Function
	
	Private Function IsSelectionValid(ByRef varID As Object, ByRef strType As String) As String
		
		Dim rsTemp As ADODB.Recordset
		
		IsSelectionValid = vbNullString
		'UPGRADE_WARNING: Couldn't resolve default property of object varID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Val(varID) = 0 Then Exit Function
		
		rsTemp = GetSelectionAccess(varID, strType)
		
		If rsTemp.BOF And rsTemp.EOF Then
			IsSelectionValid = "The " & strType & " used in this definition has been " & "deleted by another user"
			
		ElseIf LCase(Trim(rsTemp.Fields("Username").Value)) <> LCase(Trim(datGeneral.Username)) And rsTemp.Fields("Access").Value = "HD" Then 
			'JPD 20040706 Fault 8781
			If Not CurrentUserIsSysSecMgr Then
				IsSelectionValid = "The " & strType & " used in this definition has been " & "hidden by another user"
			End If
		End If
		'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsTemp = Nothing
		
	End Function
	
	
	Public Function GetSelectionAccess(ByRef varID As Object, ByRef strType As String) As ADODB.Recordset
		
		Dim strSQL As String
		Dim rsTemp As ADODB.Recordset
		
		If strType = "picklist" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object varID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strSQL = "SELECT Access, UserName FROM AsrSysPicklistName " & "WHERE PickListID = " & CStr(varID)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object varID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strSQL = "SELECT Access, UserName FROM AsrSysExpressions " & "WHERE ExprID = " & CStr(varID)
		End If
		GetSelectionAccess = datGeneral.GetReadOnlyRecords(strSQL)
		
	End Function
	
	
	Public Function IsPicklistValid(ByRef varID As Object) As String
		IsPicklistValid = IsSelectionValid(varID, "picklist")
	End Function
	
	Public Function IsCalcValid(ByRef varID As Object) As String
		IsCalcValid = IsSelectionValid(varID, "calculation")
	End Function
	
	
	Public Function GetExprField(ByRef lngExprID As Integer, ByRef sField As String) As Object
		
		Dim sSQL As String
		Dim rsExpr As ADODB.Recordset
		
		On Error GoTo ErrorTrap
		
		sSQL = "SELECT * FROM ASRSysExpressions WHERE ExprID = " & lngExprID
		
		rsExpr = datGeneral.GetRecords(sSQL)
		
		With rsExpr
			If .RecordCount > 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object GetExprField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetExprField = .Fields(sField).Value
			End If
		End With
		
		rsExpr.Close()
		
TidyUpAndExit: 
		'UPGRADE_NOTE: Object rsExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsExpr = Nothing
		Exit Function
		
ErrorTrap: 
		'NO MSGBOX ON THE SERVER ! - MsgBox "Error retrieving field value from database.", vbOKOnly + vbCritical, App.Title
		Resume TidyUpAndExit
		
	End Function
	
	
	Public Function HasHiddenComponents(ByRef lngExprID As Integer) As Boolean
		
		'********************************************************************************
		' HasHiddenComponents - Loops through the passed expression searching for       *
		'                       hidden expressions (calcs/filters).                     *
		'                       Note: This function calls itself and drills down the    *
		'                       expression checking for hidden calcs & filters, then    *
		'                       works its way up the expressions/components.            *
		'                                                                               *
		' 'TM20010802 Fault 2617                                                        *
		'********************************************************************************
		
		Dim rsExpr As ADODB.Recordset
		Dim rsExprComp As ADODB.Recordset
		Dim lngCalcFilterID As Integer
		Dim bHasHiddenComp As Boolean
		Dim sStartAccess As String
		Dim sSQL As String
		
		On Error GoTo ErrorTrap
		
		'  sSQL = "SELECT * FROM ASRSysExpressions WHERE ExprID = " & lngExprID
		'  Set rsExpr = datGeneral.GetRecords(sSQL)
		
		sSQL = "SELECT * FROM ASRSysExprComponents WHERE ExprID = " & lngExprID
		rsExprComp = datGeneral.GetRecords(sSQL)
		
		bHasHiddenComp = False
		
		With rsExprComp
			Do Until .EOF
				Select Case .Fields("Type").Value
					Case ExpressionComponentTypes.giCOMPONENT_CALCULATION
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						lngCalcFilterID = IIf(IsDbNull(.Fields("CalculationID").Value), 0, .Fields("CalculationID").Value)
						
						If lngCalcFilterID > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object GetExprField(lngCalcFilterID, Access). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If HasHiddenComponents(lngCalcFilterID) Or GetExprField(lngCalcFilterID, "Access") = ACCESS_HIDDEN Then
								bHasHiddenComp = True
								'TM20011003
								'Need this function to just find out if there are any hidden components,
								'it was also setting the access of the functions and therefore changing
								'time stamp.
								'SetExprAccess lngCalcFilterID, "HD"
							End If
						End If
						
					Case ExpressionComponentTypes.giCOMPONENT_FILTER
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						lngCalcFilterID = IIf(IsDbNull(.Fields("FilterID").Value), 0, .Fields("FilterID").Value)
						
						If lngCalcFilterID > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object GetExprField(lngCalcFilterID, Access). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If HasHiddenComponents(lngCalcFilterID) Or GetExprField(lngCalcFilterID, "Access") = ACCESS_HIDDEN Then
								bHasHiddenComp = True
								'TM20011003
								'Need this function to just find out if there are any hidden components,
								'it was also setting the access of the functions and therefore changing
								'time stamp.
								'SetExprAccess lngCalcFilterID, "HD"
							End If
						End If
						
					Case ExpressionComponentTypes.giCOMPONENT_FIELD
						'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
						lngCalcFilterID = IIf(IsDbNull(.Fields("FieldSelectionFilter").Value), 0, .Fields("FieldSelectionFilter").Value)
						
						If lngCalcFilterID > 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object GetExprField(lngCalcFilterID, Access). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If HasHiddenComponents(lngCalcFilterID) Or GetExprField(lngCalcFilterID, "Access") = ACCESS_HIDDEN Then
								bHasHiddenComp = True
								'TM20011003
								'Need this function to just find out if there are any hidden components,
								'it was also setting the access of the functions and therefore changing
								'time stamp.
								'SetExprAccess lngCalcFilterID, "HD"
							End If
						End If
						
					Case ExpressionComponentTypes.giCOMPONENT_FUNCTION
						sSQL = "SELECT exprID FROM ASRSysExpressions WHERE parentComponentID = " & CStr(.Fields("ComponentID").Value)
						rsExpr = datGeneral.GetRecords(sSQL)
						Do Until rsExpr.EOF
							'UPGRADE_WARNING: Couldn't resolve default property of object GetExprField(rsExpr!ExprID, Access). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
              If HasHiddenComponents(rsExpr.Fields("ExprID").Value) Or GetExprField(rsExpr.Fields("ExprID").Value, "Access") = ACCESS_HIDDEN Then
                bHasHiddenComp = True
                Exit Do
              End If
							
							rsExpr.MoveNext()
						Loop 
						rsExpr.Close()
						'UPGRADE_NOTE: Object rsExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
						rsExpr = Nothing
				End Select
				
				If bHasHiddenComp Then
					Exit Do
				End If
				
				.MoveNext()
			Loop 
		End With
		
		'TM20011003
		'Need this function to just find out if there are any hidden components,
		'it was also setting the access of the functions and therefore changing
		'time stamp.
		'  If bHasHiddenComp Then SetExprAccess lngExprID, "HD"
		HasHiddenComponents = bHasHiddenComp
		
		rsExprComp.Close()
		'  rsExpr.Close
		
TidyUpAndExit: 
		'  Set rsExpr = Nothing
		'UPGRADE_NOTE: Object rsExprComp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsExprComp = Nothing
		
		Exit Function
		
ErrorTrap: 
		HasHiddenComponents = False
		Resume TidyUpAndExit
		
	End Function
	
	
	Public Function GetPickListField(ByRef lngPicklistID As Integer, ByRef sField As String) As Object
		
		Dim sSQL As String
		Dim rsExpr As ADODB.Recordset
		
		On Error GoTo ErrorTrap
		
		sSQL = "SELECT * FROM ASRSysPickListName WHERE PickListID = " & lngPicklistID
		
		rsExpr = datGeneral.GetRecords(sSQL)
		
		With rsExpr
			If .RecordCount > 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object GetPickListField. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				GetPickListField = .Fields(sField).Value
			End If
		End With
		
		rsExpr.Close()
		
TidyUpAndExit: 
		'UPGRADE_NOTE: Object rsExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsExpr = Nothing
		Exit Function
		
ErrorTrap: 
		'NO MSGBOX ON THE SERVER ! - MsgBox "Error retrieving field value from database.", vbOKOnly + vbCritical, App.Title
		Resume TidyUpAndExit
		
	End Function
	
	Public Function HasExpressionComponent(ByRef plngExprIDBeingSearched As Integer, ByRef plngExprIDSearchedFor As Integer) As Boolean
		'JPD 20040507 Fault 8600
		On Error GoTo ErrorTrap
		
		Dim rsExprComp As ADODB.Recordset
		Dim rsExpr As ADODB.Recordset
		Dim fHasExpr As Boolean
		Dim sSQL As String
		Dim lngSubExprID As Integer
		
		HasExpressionComponent = (plngExprIDBeingSearched = plngExprIDSearchedFor)
		
		If Not HasExpressionComponent Then
			sSQL = "SELECT * FROM ASRSysExprComponents WHERE ExprID = " & CStr(plngExprIDBeingSearched)
			rsExprComp = datGeneral.GetRecords(sSQL)
			
			With rsExprComp
				Do Until .EOF
					Select Case .Fields("Type").Value
						Case ExpressionComponentTypes.giCOMPONENT_CALCULATION
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							lngSubExprID = IIf(IsDbNull(.Fields("CalculationID").Value), 0, .Fields("CalculationID").Value)
							
							If lngSubExprID > 0 Then
								HasExpressionComponent = HasExpressionComponent(lngSubExprID, plngExprIDSearchedFor)
							End If
							
						Case ExpressionComponentTypes.giCOMPONENT_FILTER
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							lngSubExprID = IIf(IsDbNull(.Fields("FilterID").Value), 0, .Fields("FilterID").Value)
							
							If lngSubExprID > 0 Then
								HasExpressionComponent = HasExpressionComponent(lngSubExprID, plngExprIDSearchedFor)
							End If
							
						Case ExpressionComponentTypes.giCOMPONENT_FIELD
							'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
							lngSubExprID = IIf(IsDbNull(.Fields("FieldSelectionFilter").Value), 0, .Fields("FieldSelectionFilter").Value)
							
							If lngSubExprID > 0 Then
								HasExpressionComponent = HasExpressionComponent(lngSubExprID, plngExprIDSearchedFor)
							End If
							
						Case ExpressionComponentTypes.giCOMPONENT_FUNCTION
							sSQL = "SELECT exprID FROM ASRSysExpressions WHERE parentComponentID = " & CStr(.Fields("ComponentID").Value)
							rsExpr = datGeneral.GetRecords(sSQL)
							Do Until rsExpr.EOF
								HasExpressionComponent = HasExpressionComponent(rsExpr.Fields("ExprID").Value, plngExprIDSearchedFor)
								
								If HasExpressionComponent Then
									Exit Do
								End If
								
								rsExpr.MoveNext()
							Loop 
							rsExpr.Close()
							'UPGRADE_NOTE: Object rsExpr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
							rsExpr = Nothing
					End Select
					
					If HasExpressionComponent Then
						Exit Do
					End If
					
					.MoveNext()
				Loop 
			End With
			
			rsExprComp.Close()
		End If
		
TidyUpAndExit: 
		'UPGRADE_NOTE: Object rsExprComp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsExprComp = Nothing
		
		Exit Function
		
ErrorTrap: 
		Resume TidyUpAndExit
		
	End Function
End Module