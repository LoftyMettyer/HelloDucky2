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
	
	Public Function IsFileCompatibleWithWordVersion(strFilename As String, intOfficeVersion As Short) As Boolean
		Return (GetOfficeSaveAsFormat(strFilename, intOfficeVersion, OfficeApp.oaWord) <> "")
	End Function

	Public Function IsFileCompatibleWithExcelVersion(strFilename As String, intOfficeVersion As Short) As Boolean
		Return (GetOfficeSaveAsFormat(strFilename, intOfficeVersion, OfficeApp.oaExcel) <> "")
	End Function

	Public Function GetOfficeSaveAsFormat(strFilename As String, intOfficeVersion As Short, app As OfficeApp) As String

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

End Module