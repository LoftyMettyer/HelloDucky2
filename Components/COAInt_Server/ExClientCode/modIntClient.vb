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

	'Public Function GetSaveAsFormat2(ByRef strFilename As String, ByRef strSaveAsValues As String) As String

	'	Dim strArray() As String
	'	Dim intIndex As Short
	'	Dim strExtension As String
	'	Dim strResult As String


	'	strExtension = LCase(Mid(strFilename, InStrRev(strFilename, ".") + 1))
	'	strArray = Split(strSaveAsValues, "|")

	'	strResult = ""
	'	For intIndex = 0 To UBound(strArray) - 1 'Step 2
	'		If LCase(strArray(intIndex)) = strExtension Then
	'			strResult = strArray(intIndex + 1)
	'			Exit For
	'		End If
	'	Next

	'	GetSaveAsFormat2 = strResult

	'End Function

End Module