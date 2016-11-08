Option Strict Off
Option Explicit On

Imports HR.Intranet.Server.BaseClasses
Imports HR.Intranet.Server.Enums
Imports System.IO

Public Class Utilities
	Inherits BaseForDMI

	Private mstrCommonDialogFormatsWord As String
	Private mstrCommonDialogFormatsExcel As String
	Private mintDefaultIndexWord As Integer
	Private mintDefaultIndexExcel As Integer
	Private mstrOfficeSaveAsValues As String

	' Array holding the User Defined functions that are needed for this report
	Private mastrUDFsRequired() As String

	Public Function UDFFunctions(ByRef pbCreate As Boolean) As Boolean
		Return General.UDFFunctions(mastrUDFsRequired, pbCreate)
	End Function

	Public Function GetFilteredIDs(ByRef plngExprID As Integer, ByRef pavPromptedValues As Object) As String
		' Return a string describing the record IDs from the given table
		' that satisfy the given criteria.
		Dim sIDSQL As String = ""
		Dim avPrompts(,) As Object
		Dim iDataType As Short
		Dim lngComponentID As Integer
		Dim iLoop As Short

		ReDim avPrompts(1, 0)

		If IsArray(pavPromptedValues) Then
			ReDim avPrompts(1, UBound(pavPromptedValues, 2))

			For iLoop = 0 To UBound(pavPromptedValues, 2)
				' Get the prompt data type.
				'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Len(Trim(Mid(pavPromptedValues(0, iLoop), 10))) > 0 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					lngComponentID = CInt(Mid(pavPromptedValues(0, iLoop), 10))
					'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					iDataType = CShort(Mid(pavPromptedValues(0, iLoop), 8, 1))

					'UPGRADE_WARNING: Couldn't resolve default property of object avPrompts(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					avPrompts(0, iLoop) = lngComponentID

					' NB. Locale to server conversions are done on the client.
					Select Case iDataType
						Case 2
							' Numeric.
							'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object avPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avPrompts(1, iLoop) = Double.Parse(pavPromptedValues(1, iLoop), System.Globalization.CultureInfo.InvariantCulture)
						Case 3
							' avPrompts.
							'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object avPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avPrompts(1, iLoop) = (UCase(CStr(pavPromptedValues(1, iLoop))) = "TRUE")
						Case 4
							' Date.
							' JPD 20040212 Fault 8082 - DO NOT CONVERT DATE PROMPTED VALUES
							' THEY ARE PASSED IN FROM THE ASPs AS STRING VALUES IN THE CORRECT
							' FORMAT (mm/dd/yyyy) AND DOING ANY KIND OF CONVERSION JUST SCREWS
							' THINGS UP.
							'avPrompts(1, iLoop) = CDate(pavPromptedValues(1, iLoop))
							'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object avPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avPrompts(1, iLoop) = pavPromptedValues(1, iLoop)
						Case Else
							'UPGRADE_WARNING: Couldn't resolve default property of object pavPromptedValues(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object avPrompts(1, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							avPrompts(1, iLoop) = CStr(pavPromptedValues(1, iLoop))
					End Select
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object avPrompts(0, iLoop). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					avPrompts(0, iLoop) = 0
				End If
			Next iLoop
		End If

		Dim blnOK As Boolean

		ReDim mastrUDFsRequired(0)

		blnOK = FilteredIDs(plngExprID, sIDSQL, mastrUDFsRequired, avPrompts)

		Return sIDSQL

	End Function

	Public ReadOnly Property OfficeGetDefaultIndexWord() As Integer
		Get
			Return mintDefaultIndexWord
		End Get
	End Property

	Public ReadOnly Property OfficeGetDefaultIndexExcel() As Integer
		Get
			Return mintDefaultIndexExcel
		End Get
	End Property

	Public ReadOnly Property OfficeGetCommonDialogFormatsWord() As String
		Get
			Return mstrCommonDialogFormatsWord
		End Get
	End Property

	Public ReadOnly Property OfficeGetCommonDialogFormatsExcel() As String
		Get
			Return mstrCommonDialogFormatsExcel
		End Get
	End Property

	Public ReadOnly Property OfficeGetSaveAsValues() As String
		Get
			Return mstrOfficeSaveAsValues
		End Get
	End Property

	Public Sub OfficeInitialise(ByRef intWordVer As Short, ByRef intExcelVer As Short)

		OfficeGetCommonDialogFormats("Word", intWordVer)
		OfficeGetCommonDialogFormats("Excel", intExcelVer)

		mstrOfficeSaveAsValues = OfficeGetSaveAsValuesForDestin("Word", intWordVer) & "|" & OfficeGetSaveAsValuesForDestin("WordTemplate", intWordVer) & "|" & OfficeGetSaveAsValuesForDestin("Excel", intExcelVer) & "|" & OfficeGetSaveAsValuesForDestin("ExcelTemplate", intExcelVer)

	End Sub

	Private Function OfficeGetSaveAsValuesForDestin(strDestin As String, intOfficeVersion As Short) As String

		Dim rsTemp As DataTable
		Dim strSQL As String
		Dim strFormatField As String
		Dim strOutput As String = ""

		strFormatField = IIf(intOfficeVersion < 12, "Office2003", "Office2007").ToString()

		strSQL = "SELECT * FROM ASRSysFileFormats WHERE  Destination = '" _
				& Replace(strDestin, "'", "''") & "' AND  NOT " & strFormatField & " IS NULL ORDER BY ID"
		rsTemp = DB.GetDataTable(strSQL)

		For Each objRow As DataRow In rsTemp.Rows
			If strOutput <> vbNullString Then
				strOutput = strOutput & "|"
			End If

			strOutput = strOutput & objRow("Extension").ToString() & "|" & objRow(strFormatField).ToString()

		Next

		Return strOutput

	End Function


	Private Sub OfficeGetCommonDialogFormats(strDestin As String, intOfficeVersion As Short)

		Dim objSettings As clsSettings
		Dim rsTemp As DataTable
		Dim strSQL As String
		Dim strOutput As String
		Dim strFormatField As String
		Dim intCount As Integer
		Dim intDefaultFormat As Integer
		Dim intDefaultFormatIndex As Integer

		objSettings = New clsSettings
		objSettings.SessionInfo = SessionInfo
		'UPGRADE_WARNING: Couldn't resolve default property of object objSettings.GetUserSetting(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		intDefaultFormat = CInt(objSettings.GetUserSetting("Output", strDestin & "Format", -1))
		'UPGRADE_NOTE: Object objSettings may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		objSettings = Nothing

		strFormatField = IIf(intOfficeVersion < 12, "Office2003", "Office2007").ToString()

		strSQL = "SELECT * FROM ASRSysFileFormats WHERE  Destination = '" & Replace(strDestin, "'", "''") _
					 & "' AND NOT " & strFormatField & " IS NULL ORDER BY ID"
		rsTemp = DB.GetDataTable(strSQL)

		intCount = 1
		intDefaultFormatIndex = -1
		strOutput = vbNullString
		For Each objRow As DataRow In rsTemp.Rows
			If strOutput <> vbNullString Then
				strOutput = strOutput & "|"
			End If

			If CInt(objRow(strFormatField)) = intDefaultFormat Then
				intDefaultFormatIndex = intCount

			ElseIf CBool(objRow("Default")) Then
				If intDefaultFormatIndex = -1 Then
					intDefaultFormatIndex = intCount
				End If

			End If

			strOutput = strOutput & objRow("Description").ToString() & "|*." & objRow("Extension").ToString()

			intCount += 1
		Next


		If strDestin = "Word" Then
			mstrCommonDialogFormatsWord = strOutput
			mintDefaultIndexWord = intDefaultFormatIndex
		Else
			mstrCommonDialogFormatsExcel = strOutput
			mintDefaultIndexExcel = intDefaultFormatIndex
		End If

	End Sub
End Class