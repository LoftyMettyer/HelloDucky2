Option Strict Off
Option Explicit On
Module modIntClient

	Public Enum OfficeApp
		oaWord = 0
		oaExcel = 1
	End Enum

	' Microsoft Word Output Types
	Public Enum WordOutputType
		wdFormatDocument = 0
		wdFormatDOSText = 4
		wdFormatDOSTextLineBreaks = 5
		wdFormatEncodedText = 7
		wdFormatFilteredHTML = 10
		wdFormatHTML = 8
		wdFormatRTF = 6
		wdFormatTemplate = 1
		wdFormatText = 2
		wdFormatTextLineBreaks = 3
		wdFormatUnicodeText = 7
		wdFormatWebArchive = 9
		wdFormatXML = 11
		wdFormatDocument97 = 0
		wdFormatDocumentDefault = 16
		wdFormatPDF = 17
		wdFormatTemplate97 = 1
		wdFormatXMLDocument = 12
		wdFormatXMLDocumentMacroEnabled = 13
		wdFormatXMLTemplate = 14
		wdFormatXMLTemplateMacroEnabled = 15
		wdFormatXPS = 18
	End Enum
	
	'Public Enum OutputDestinations
	'	desScreen = 0
	'	desPrinter = 1
	'	desSave = 2
	'	desEmail = 3
	'End Enum

	'Public gobjProgress As New clsHRProProgress
	'Public gsDatabaseName As String
	'Public gsServerName As String
	'Public gsUserName As String
	'Public gsDocumentsPath As String

	'Public strSettingWordTemplate As String
	'Public strSettingExcelTemplate As String
	'Public blnSettingExcelGridlines As Boolean
	'Public blnSettingExcelHeaders As Boolean
	'Public blnSettingAutoFitCols As Boolean
	'Public blnSettingLandscape As Boolean
	'
	'Public lngSettingTitleCol As Long
	'Public lngSettingTitleRow As Long
	'Public blnSettingTitleGridlines As Boolean
	'Public blnSettingTitleBold As Boolean
	'Public blnSettingTitleUnderline As Boolean
	'Public lngSettingTitleBackcolour As Long
	'Public lngSettingTitleForecolour As Long
	'
	'Public lngSettingHeadingCol As Long
	'Public lngSettingHeadingRow As Long
	'Public blnSettingHeadingGridlines As Boolean
	'Public blnSettingHeadingBold As Boolean
	'Public blnSettingHeadingUnderline As Boolean
	'Public lngSettingHeadingBackcolour As Long
	'Public lngSettingHeadingForecolour As Long
	'
	'Public lngSettingDataCol As Long
	'Public lngSettingDataRow As Long
	'Public blnSettingDataGridlines As Boolean
	'Public blnSettingDataBold As Boolean
	'Public blnSettingDataUnderline As Boolean
	'Public lngSettingDataBackcolour As Long
	'Public lngSettingDataForecolour As Long


	Private Const LOCALE_SYSTEM_DEFAULT As Integer = &H800
	Private Const LOCALE_USER_DEFAULT As Integer = &H400
	Private Const LOCALE_SDATE As Integer = &H1D ' date separator
	Private Const LOCALE_SSHORTDATE As Integer = &H1F	' short date format string
	Private Const LOCALE_SDECIMAL As Integer = &HE ' decimal separator
	Private Const LOCALE_STHOUSAND As Integer = &HF	' thousand separator
	Private Const LOCALE_IMEASURE As Integer = &HD ' Measurement System


	Public Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, ByRef cbRemoteName As Integer) As Integer
	Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Integer, ByVal lpTempFileName As String) As Integer
'	Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Integer, ByVal lpBuffer As String) As Integer
	Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal locale As Integer, ByVal LCType As Integer, ByVal lpLCData As String, ByVal cchData As Integer) As Integer

	'Public Function GetTmpFName() As String

	'	Dim strTmpPath, strTmpName As String

	'	strTmpPath = Space(1024)
	'	strTmpName = Space(1024)

	'	Call GetTempPath(1024, strTmpPath)
	'	Call GetTempFileName(strTmpPath, "_T", 0, strTmpName)

	'	strTmpName = Trim(strTmpName)
	'	If Len(strTmpName) > 0 Then
	'		strTmpName = Left(strTmpName, Len(strTmpName) - 1)

	'		'MH20021227 For some reason a zero byte file is created... ANNOYING!
	'		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
	'		If Dir(strTmpName) <> vbNullString Then
	'			Kill(strTmpName)
	'		End If

	'	Else
	'		strTmpName = vbNullString
	'	End If

	'	GetTmpFName = Trim(strTmpName)

	'End Function


	'Public Function DateFormat() As String
	'	' Returns the date format.
	'	' NB. Windows allows the user to configure totally stupid
	'	' date formats (eg. d/M/yyMydy !). This function does not cater
	'	' for such stupidity, and simply takes the first occurence of the
	'	' 'd', 'M', 'y' characters.
	'	Dim sSysFormat As String
	'	Dim sSysDateSeparator As String
	'	Dim sDateFormat As String
	'	Dim iLoop As Short
	'	Dim fDaysDone As Boolean
	'	Dim fMonthsDone As Boolean
	'	Dim fYearsDone As Boolean

	'	fDaysDone = False
	'	fMonthsDone = False
	'	fYearsDone = False
	'	sDateFormat = ""

	'	sSysFormat = GetSystemDateFormat
	'	sSysDateSeparator = GetSystemDateSeparator

	'	' Loop through the string picking out the required characters.
	'	For iLoop = 1 To Len(sSysFormat)

	'		Select Case Mid(sSysFormat, iLoop, 1)
	'			Case "d"
	'				If Not fDaysDone Then
	'					' Ensure we have two day characters.
	'					sDateFormat = sDateFormat & "dd"
	'					fDaysDone = True
	'				End If

	'			Case "M"
	'				If Not fMonthsDone Then
	'					' Ensure we have two month characters.
	'					sDateFormat = sDateFormat & "mm"
	'					fMonthsDone = True
	'				End If

	'			Case "y"
	'				If Not fYearsDone Then
	'					' Ensure we have four year characters.
	'					sDateFormat = sDateFormat & "yyyy"
	'					fYearsDone = True
	'				End If

	'			Case Else
	'				sDateFormat = sDateFormat & Mid(sSysFormat, iLoop, 1)
	'		End Select

	'	Next iLoop

	'	' Ensure that all day, month and year parts of the date
	'	' are present in the format.
	'	If Not fDaysDone Then
	'		If Mid(sDateFormat, Len(sDateFormat), 1) <> sSysDateSeparator Then
	'			sDateFormat = sDateFormat & sSysDateSeparator
	'		End If

	'		sDateFormat = sDateFormat & "dd"
	'	End If

	'	If Not fMonthsDone Then
	'		If Mid(sDateFormat, Len(sDateFormat), 1) <> sSysDateSeparator Then
	'			sDateFormat = sDateFormat & sSysDateSeparator
	'		End If

	'		sDateFormat = sDateFormat & "mm"
	'	End If

	'	If Not fYearsDone Then
	'		If Mid(sDateFormat, Len(sDateFormat), 1) <> sSysDateSeparator Then
	'			sDateFormat = sDateFormat & sSysDateSeparator
	'		End If

	'		sDateFormat = sDateFormat & "yyyy"
	'	End If

	'	' Return the date format.
	'	DateFormat = sDateFormat

	'End Function


	'Public Sub SetComboText(ByRef cboCombo As System.Windows.Forms.ComboBox, ByRef sText As String)

	'	Dim lCount As Integer

	'	With cboCombo
	'		For lCount = 1 To .Items.Count
	'			If VB6.GetItemString(cboCombo, lCount - 1) = sText Then
	'				.SelectedIndex = lCount - 1
	'				Exit For
	'			End If
	'		Next
	'	End With

	'End Sub


	'Public Sub SetComboItem(ByRef cboCombo As System.Windows.Forms.ComboBox, ByRef lItem As Integer)

	'	Dim lCount As Integer

	'	With cboCombo
	'		For lCount = 1 To .Items.Count
	'			If VB6.GetItemData(cboCombo, lCount - 1) = lItem Then
	'				.SelectedIndex = lCount - 1
	'				Exit For
	'			End If
	'		Next
	'	End With

	'End Sub

	'Public Sub EnableCombo(ByRef cboTemp As System.Windows.Forms.ComboBox, ByRef blnEnabled As Boolean)
	'	blnEnabled = (blnEnabled And cboTemp.Items.Count > 0)
	'	cboTemp.Enabled = blnEnabled
	'	cboTemp.BackColor = System.Drawing.ColorTranslator.FromOle(IIf(blnEnabled, System.Drawing.ColorTranslator.ToOle(System.Drawing.SystemColors.Window), System.Drawing.ColorTranslator.ToOle(System.Drawing.SystemColors.Control)))
	'	cboTemp.SelectedIndex = IIf(blnEnabled, 0, -1)
	'End Sub

	Function GetSystemMeasurement() As String

		On Error GoTo ErrorTrap

		Dim lngLength As Integer
		Dim sBuffer As New String(" ", 100)

		lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_IMEASURE, sBuffer, 99)
		GetSystemMeasurement = Left(sBuffer, lngLength - 1)

		If CDbl(GetSystemMeasurement) = 1 Then
			GetSystemMeasurement = "us"
		Else
			GetSystemMeasurement = "metric"
		End If

TidyUpAndExit:
		Exit Function
ErrorTrap:

	End Function

	Function GetSystemDateSeparator() As String
		' Return the system data separator.
		Dim lngLength As Integer
		Dim sBuffer As New String(" ", 100)

		lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDATE, sBuffer, 99)
		GetSystemDateSeparator = Left(sBuffer, lngLength - 1)

	End Function

	Function GetSystemDateFormat() As String
		' Return the system data format.
		Dim lngLength As Integer
		Dim sBuffer As New String(" ", 100)

		lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE, sBuffer, 99)
		GetSystemDateFormat = Left(sBuffer, lngLength - 1)

	End Function


	Public Function IsFileCompatibleWithWordVersion(ByRef strFilename As String, ByRef intOfficeVersion As Short) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object IsFileCompatibleWithWordVersion. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		IsFileCompatibleWithWordVersion = (GetOfficeSaveAsFormat(strFilename, intOfficeVersion, OfficeApp.oaWord) <> "")
	End Function

	Public Function IsFileCompatibleWithExcelVersion(ByRef strFilename As String, ByRef intOfficeVersion As Short) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object IsFileCompatibleWithExcelVersion. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		IsFileCompatibleWithExcelVersion = (GetOfficeSaveAsFormat(strFilename, intOfficeVersion, OfficeApp.oaExcel) <> "")
	End Function

	Public Function GetOfficeSaveAsFormat(ByRef strFilename As String, ByRef intOfficeVersion As Short, ByRef app As OfficeApp) As String

		Dim strOutput As String
		Dim strExtension As String
		Dim bln2007 As Boolean

		On Error GoTo LocalErr

		strOutput = ""

		If intOfficeVersion > 0 And InStr(strFilename, ".") Then
			strExtension = Trim(LCase(Mid(strFilename, InStrRev(strFilename, ".") + 1)))
			bln2007 = (intOfficeVersion >= 12)

			Select Case strExtension
				Case "doc" : strOutput = IIf(bln2007, "0", "0")
				Case "dot" : strOutput = IIf(bln2007, "1", "0")
				Case "xls" : strOutput = IIf(bln2007, "56", "-4143")
				Case "xlt" : strOutput = IIf(bln2007, "17", "17")
				Case "docx" : strOutput = IIf(bln2007, "16", "")
				Case "dotx" : strOutput = IIf(bln2007, "14", "")
				Case "xlsx" : strOutput = IIf(bln2007, "51", "")
				Case "xltx" : strOutput = IIf(bln2007, "17", "")
				Case "pdf" : strOutput = IIf(bln2007, "17", "")
				Case "txt" : strOutput = IIf(bln2007, "2", "")
				Case "rtf" : strOutput = IIf(bln2007, "6", "")
				Case "xml" : strOutput = IIf(bln2007, "12", "")	'not in table
				Case "xps" : strOutput = IIf(bln2007, "18", "")	'not in table
				Case "html"
					Select Case app
						Case OfficeApp.oaWord
							strOutput = IIf(bln2007, "8", "")
						Case OfficeApp.oaExcel
							strOutput = IIf(bln2007, "44", "")
					End Select
			End Select

		End If

		GetOfficeSaveAsFormat = strOutput

		Exit Function

LocalErr:
		GetOfficeSaveAsFormat = ""

	End Function


	Public Function GetSaveAsFormat2(ByRef strFilename As String, ByRef strSaveAsValues As String) As String

		Dim strArray() As String
		Dim intIndex As Short
		Dim strExtension As String
		Dim strResult As String


		strExtension = LCase(Mid(strFilename, InStrRev(strFilename, ".") + 1))
		strArray = Split(strSaveAsValues, "|")

		strResult = ""
		For intIndex = 0 To UBound(strArray) - 1 'Step 2
			If LCase(strArray(intIndex)) = strExtension Then
				strResult = strArray(intIndex + 1)
				Exit For
			End If
		Next

		GetSaveAsFormat2 = strResult

	End Function
End Module