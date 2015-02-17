Option Strict Off
Option Explicit On

Imports HR.Intranet.Server.Enums

Module modExpression


	Public Const giPRINT_XINDENT As Short = 1000
	Public Const giPRINT_YINDENT As Short = 1000
	Public Const giPRINT_XSPACE As Short = 500
	Public Const giPRINT_YSPACE As Short = 100

	Public Const giEXPRVALUE_BYREF_OFFSET As Short = 100

	Public Const gsDUMMY_CHARACTER As String = "ASRDUMMYCHARVALUE"
	Public Const gsDUMMY_NUMERIC As Short = 1
	Public Const gsDUMMY_LOGIC As Boolean = True
	Public Const gsDUMMY_DATE As Date = #1/1/1998#
	Public Const gsDUMMY_BYREF_CHARACTER As String = "12" & vbTab & "a"
	Public Const gsDUMMY_BYREF_NUMERIC As String = "2" & vbTab & "1"
	Public Const gsDUMMY_BYREF_LOGIC As String = "-7" & vbTab & "0"
	Public Const gsDUMMY_BYREF_DATE As String = "11" & vbTab & "1/1/1998"

	' Parameter Type constants.
	Public Const gsPARAMETERTYPE_ORDERID As String = "PType_OrderID"

	Public Function ExpressionTypeName(ByRef piType As ExpressionTypes) As String
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

			Case ExpressionTypes.giEXPR_PAGEBREAK	' NOT USED.
				ExpressionTypeName = "Page Break"

			Case ExpressionTypes.giEXPR_ORDER	' NOT USED.
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

	Public Function ValidNameChar(ByVal piAsciiCode As Integer, ByVal piPosition As Short) As Integer
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

		Return piAsciiCode

ErrorTrap:
		ValidNameChar = 0
		Err.Clear()

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

			Case 10	' IS GREATER THAN
				fOK = ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER))
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC

			Case 11	' IS LESS THAN OR EQUAL TO
				fOK = ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER))
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC

			Case 12	' IS GREATER THAN OR EQUAL TO
				fOK = ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER))
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC

			Case 13	' NOT
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_LOGIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC

			Case 14	' IS CONTAINED WITHIN
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC

			Case 15	' TO THE POWER OF
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 16	' MODULAS
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 17	' CONCATENATED WITH
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER

			Case Else	' Unknown operator
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

	Public Function ValidateFunctionParameters(ByRef plngFunctionID As Integer, ByRef piResultType As ExpressionValueTypes, Optional ByRef piParam1Type As Integer = 0, Optional ByRef piParam2Type As Integer = 0, Optional ByRef piParam3Type As Integer = 0, Optional ByRef piParam4Type As Integer = 0, Optional ByRef piParam5Type As Integer = 0, Optional ByRef piParam6Type As Integer = 0) As Boolean
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

			Case 10	' MINIMUM
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 11	' SEARCH FOR CHARACTER STRING
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 12	' CAPITALIZE INITIALS
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER

			Case 13	'EXTRACT CHARACTERS FROM THE RIGHT
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER

			Case 14	' EXTRACT PART OF A STRING
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam3Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER

			Case 15	' SYSTEM TIME
				fOK = True
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER

			Case 16	' IS FIELD EMPTY
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) Or (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) Or (piParam1Type = ExpressionValueTypes.giEXPRVALUE_LOGIC) Or (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC

			Case 17	' CURRENT USER
				fOK = True
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER

			Case 18	' WHOLE YEARS UNTIL CURRENT DATE
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 19	' REMAINING MONTHS SINCE WHOLE YEARS
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 20	' INITIALS FROM FORENAMES
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER

			Case 21	' FIRST NAME FROM FORENAMES
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER

			Case 22	' WEEKDAYS FROM START AND END DATES
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 23	' ADD MONTHS TO DATE
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_DATE

			Case 24	' ADD YEARS TO DATE
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_DATE

			Case 25	' CONVERT CHARACTER TO NUMERIC
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 26	' WHOLE MONTHS BETWEEN TWO DATES
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 27	' PARENTHESESES
				fOK = True
				piResultType = piParam1Type

			Case 28	' DAY OF THE WEEK
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 29	' NUMBER OF WORKING DAYS PER WEEK
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 30	' ABSENCE DURATION
				'TM08102003
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam3Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam4Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 31	' ROUND DOWN TO NEAREST WHOLE NUMBER
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 32	' YEAR OF DATE
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 33	' MONTH OF DATE
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 34	' DAY OF DATE
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 35	' NICE DATE
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER

			Case 36	' NICE TIME
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER

			Case 37	' ROUND DATE TO START OF NEAREST MONTH
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_DATE

			Case 38	' IS BETWEEN
				fOK = ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam3Type = ExpressionValueTypes.giEXPRVALUE_DATE)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam3Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC))
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC

			Case 39	' SERVICE YEARS
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 40	' SERVICE MONTHS
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 41	' STATUTORY REDUNDANCY PAY
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam3Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam4Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam5Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 42	' GET FIELD FROM DATABASE RECORD
				fOK = (((piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_LOGIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_LOGIC)) Or ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE))) And ((piParam3Type = ExpressionValueTypes.giEXPRVALUE_BYREF_CHARACTER) Or (piParam3Type = ExpressionValueTypes.giEXPRVALUE_BYREF_NUMERIC) Or (piParam3Type = ExpressionValueTypes.giEXPRVALUE_BYREF_LOGIC) Or (piParam3Type = ExpressionValueTypes.giEXPRVALUE_BYREF_DATE))
				piResultType = (piParam3Type - giEXPRVALUE_BYREF_OFFSET)

			Case 43	' GET UNIQUE CODE
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 44	' ADD DAYS TO DATE
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_DATE

			Case 45	' DAYS BETWEEN TWO DATES
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 46	' WORKING DAYS BETWEEN TWO DATES (INC BHOLS)
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

				'JPD 20031017 Fault 7269
			Case 47	' ABSENCE BETWEEN TWO DATES
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam3Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 48	' ROUND UP TO NEAREST WHOLE NUMBER
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 49	' ROUND TO NEAREST NUMBER
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 51	' CONVERT CURRENCY
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam3Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 52	' Field Last Changed Date
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_CHARACTER Or piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_NUMERIC Or piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_LOGIC Or piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_DATE

			Case 53	' Field changed between two dates
				fOK = ((piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_CHARACTER Or piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_NUMERIC Or piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_LOGIC Or piParam1Type = ExpressionValueTypes.giEXPRVALUE_BYREF_DATE)) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC

			Case 54	'Whole years between two dates
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

				' JPD20021121 Fault 3177
			Case 55	' First Day of Month - VERSION 2 FUNCTION
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_DATE

				' JPD20021121 Fault 3177
			Case 56	' Last Day of Month - VERSION 2 FUNCTION
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_DATE

				' JPD20021121 Fault 3177
			Case 57	' First Day of Year - VERSION 2 FUNCTION
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_DATE

				' JPD20021121 Fault 3177
			Case 58	' Last Day of Year - VERSION 2 FUNCTION
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_DATE

				' JPD20021129 Fault 4337
			Case 59	' NAME OF MONTH
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER

				' JPD20021129 Fault 4337
			Case 60	' NAME OF DAY
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER

				' JPD20021129 Fault 3606
			Case 61	' IS FIELD POPULATED
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) Or (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) Or (piParam1Type = ExpressionValueTypes.giEXPRVALUE_LOGIC) Or (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC

			Case 65	'IS POST SUBORDINATE OF
				fOK = False
				'Select Case IdentifyingColumnDataType
				'  Case sqlNumeric, sqlInteger
				'    fOK = (piParam1Type = giEXPRVALUE_NUMERIC)
				'  Case Else
				'    fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
				'End Select
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC

			Case 66	'IS POST SUBORDINATE OF USER
				fOK = True
				'fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
				''      (piParam2Type = giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC

			Case 67	'IS PERSONNEL SUBORDINATE OF
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

			Case 68	'IS PERSONNEL SUBORDINATE OF USER
				fOK = True
				'fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
				''      (piParam2Type = giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC

			Case 69	'HAS POST SUBORDINATE
				fOK = False
				'Select Case IdentifyingColumnDataType
				'  Case sqlNumeric, sqlInteger
				'    fOK = (piParam1Type = giEXPRVALUE_NUMERIC)
				'  Case Else
				'    fOK = (piParam1Type = giEXPRVALUE_CHARACTER)
				'End Select
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC

			Case 70	'HAS POST SUBORDINATE USER
				fOK = True
				'fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
				''      (piParam2Type = giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC

			Case 71	'HAS PERSONNEL SUBORDINATE
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

			Case 72	'HAS PERSONNEL SUBORDINATE USER
				fOK = True
				'fOK = (piParam1Type = giEXPRVALUE_CHARACTER) And _
				''      (piParam2Type = giEXPRVALUE_DATE)
				piResultType = ExpressionValueTypes.giEXPRVALUE_LOGIC

			Case 73	'BRADFORD FACTOR
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_DATE) And (piParam3Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 77	' Replace Characters within a Strin
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam2Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER) And (piParam3Type = ExpressionValueTypes.giEXPRVALUE_CHARACTER)
				piResultType = ExpressionValueTypes.giEXPRVALUE_CHARACTER

			Case 78
				fOK = True
				piResultType = ExpressionValueTypes.giEXPRVALUE_DATE

			Case 79
				fOK = True
				piResultType = ExpressionValueTypes.giEXPRVALUE_NUMERIC

			Case 80
				fOK = (piParam1Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And _
					(piParam2Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC) And _
					(piParam3Type = ExpressionValueTypes.giEXPRVALUE_NUMERIC)
				piResultType = ExpressionValueTypes.giEXPRVALUE_DATE

			Case Else	' Unknown function
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

End Module