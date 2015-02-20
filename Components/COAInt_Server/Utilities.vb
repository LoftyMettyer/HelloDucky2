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
							avPrompts(1, iLoop) = CDbl(pavPromptedValues(1, iLoop))
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

	Public Function GetPictures(ByRef plngScreenID As Integer, ByRef psTempPath As String) As Object
		Dim avPictures(,) As Object
		Dim sSQL As String
		Dim rsTemp As DataTable
		Dim sFileName As String

		ReDim avPictures(2, 0)

		sSQL = "SELECT DISTINCT ASRSysControls.pictureID, ASRSysPictures.name FROM ASRSysControls INNER JOIN ASRSysPictures ON ASRSysControls.pictureID = ASRSysPictures.pictureID" & _
			" WHERE screenID = " & Trim(Str(plngScreenID)) & " AND controlType = " & Trim(Str(ControlTypes.ctlImage))
		rsTemp = DB.GetDataTable(sSQL)

		With rsTemp
			For Each objRow As DataRow In rsTemp.Rows


				sFileName = LoadScreenControlPicture(CInt(objRow("PictureID")), psTempPath, objRow("Name").ToString())

				If Len(sFileName) > 0 Then
					ReDim Preserve avPictures(2, UBound(avPictures, 2) + 1)
					'UPGRADE_WARNING: Couldn't resolve default property of object avPictures(1, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					avPictures(1, UBound(avPictures, 2)) = CInt(objRow("PictureID"))
					'UPGRADE_WARNING: Couldn't resolve default property of object avPictures(2, UBound()). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					avPictures(2, UBound(avPictures, 2)) = Mid(sFileName, InStrRev(sFileName, "\") + 1)
				End If

			Next
		End With

		'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsTemp = Nothing

		'UPGRADE_WARNING: Couldn't resolve default property of object GetPictures. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetPictures = VB6.CopyArray(avPictures)

	End Function

	Public Function GetBackgroundPicture(ByRef psTempPath As String) As String

		' We are not currently using this functionality - disbale. To review!
		Return ""

		'Dim sSQL As String
		'Dim rsTemp As Recordset
		'Dim sFileName As String = ""
		'Dim lngPictureID As Short

		'sSQL = "SELECT DISTINCT ASRSysSystemSettings.settingValue FROM ASRSysSystemSettings WHERE ASRSysSystemSettings.section = 'desktopsetting' " _
		'		 & " AND  ASRSysSystemSettings.settingKey = 'bitmapid'"
		'rsTemp = mclsData.OpenRecordset(sSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)
		'With rsTemp
		'	If Not (.BOF And .EOF) Then
		'		lngPictureID = .Fields("settingValue").Value
		'	End If
		'	.Close()
		'End With
		''UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		'rsTemp = Nothing

		'sSQL = "SELECT DISTINCT ASRSysPictures.pictureID, ASRSysPictures.name" & " FROM ASRSysPictures " & " WHERE ASRSysPictures.pictureID = " & CStr(lngPictureID)
		'rsTemp = mclsData.OpenRecordset(sSQL, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)
		'With rsTemp
		'	If Not (.BOF And .EOF) Then
		'		sFileName = LoadScreenControlPicture(.Fields("PictureID").Value, psTempPath, .Fields("Name").Value)
		'	End If
		'	.Close()
		'End With
		''UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		'rsTemp = Nothing

		'GetBackgroundPicture = Mid(sFileName, InStrRev(sFileName, "\") + 1)

	End Function

	Public Function GetBackgroundPosition() As Integer

		Const sSQL As String = "SELECT DISTINCT settingValue FROM ASRSysSystemSettings WHERE section = 'desktopsetting' AND settingKey = 'bitmaplocation'"
		With DB.GetDataTable(sSQL)
			If .Rows.Count > 0 Then
				Return CInt(.Rows(0)(0))
			Else
				Return 0
			End If
		End With

	End Function

	Private Function LoadScreenControlPicture(plngPictureID As Integer, psTempPath As String, psName As String) As String
		' Read the given picture from the database.
		Dim sTempName As String
		Dim sPictureFile As String
		Dim rsPictures As DataTable
		Dim objRowData As DataRow
		Dim sSQL As String

		Try
			sSQL = "SELECT picture FROM ASRSysPictures WHERE pictureID = " & Trim(Str(plngPictureID))
			rsPictures = DB.GetDataTable(sSQL)

			sPictureFile = ""

			If rsPictures.Rows.Count > 0 Then

				objRowData = rsPictures.Rows(0)
				sTempName = psTempPath & "\" & psName

				Dim fs = New FileStream(sTempName, FileMode.Create)
				Dim objPicture As Byte() = objRowData(0)
				fs.Write(objPicture, 0, objPicture.Length)
				fs.Close()

			End If

		Catch ex As Exception
			Throw

		End Try

		Return sPictureFile

	End Function

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