Option Strict Off
Option Explicit On

Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6

Public Class MailMergeClient

	' General form handling variables.
	Private msLocaleDateFormat As String
	Private msLocaleDecimalSeparator As String
	Private msLocaleThousandSeparator As String

	' Windows API function constants
	Const LOCALE_USER_DEFAULT As Integer = &H400
	Const LOCALE_SDATE As Integer = &H1D '  date separator
	Const LOCALE_SSHORTDATE As Integer = &H1F	'  short date format string
	Const LOCALE_SDECIMAL As Integer = &HE '  decimal separator
	Const LOCALE_STHOUSAND As Integer = &HF	'  thousand separator

	' Windows API functions
	Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal locale As Integer, ByVal LCType As Integer, ByVal lpLCData As String, ByVal cchData As Integer) As Integer
	Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Integer, ByVal lpTempFileName As String) As Integer
	Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Integer, ByVal lpBuffer As String) As Integer

	Private mobjFileSystem As New Scripting.FileSystemObject
	Private mobjFileInfo As Scripting.File

	'MM = Mail Merge
	Private mstrMMStatusMessage As String
	Private mlngMMMergeFieldsUbound As Integer
	Private maryMMOutputArrayData() As String
	Private maryMMMergeFieldsData() As String
	Private mstrMMDataSourceName As String
	Private mblnMMCancelled As Boolean
	Private mstrMMTempFileName As String

	Private mstrMMDefName As String
	Private mstrMMDefTemplateFile As String
	Private mstrMMDefTemplateSaveAs As String
	Private mblnMMDefPauseBeforeMerge As Boolean
	Private mblnMMDefSuppressBlankLines As Boolean

	Private mstrMMDefEmailSubject As String
	Private mlngMMDefEmailAddrCalc As Integer
	Private mblnMMDefEMailAttachment As Boolean
	Private mstrMMDefAttachmentName As String
	Private mstrMMDefAttachmentFormat As String

	Private mintMMDefOutputFormat As Short
	Private mblnMMDefOutputScreen As Boolean
	Private mblnMMDefOutputPrinter As Boolean
	Private mstrMMDefOutputPrinterName As String
	Private mblnMMDefOutputSave As Boolean
	Private mstrMMDefOutputFileName As String
	Private mstrMMDefOutputSaveAs As String

	Private miOfficeVersion_Word As Short
	Private miOfficeVersion_Excel As Short
	Private mstrSaveAsValues As String

	'Word Objects
	Dim mwrdApp As Microsoft.Office.Interop.Word.Application
	Dim mdocDataSource As Microsoft.Office.Interop.Word.Document
	Dim mdocTemplate As Microsoft.Office.Interop.Word.Document
	Dim mdocOutput As Microsoft.Office.Interop.Word.Document

	Public ReadOnly Property ErrorMessage As String
		Get
			Return mstrMMStatusMessage
		End Get
	End Property


	Public Function CheckString(ByRef psString As String) As String
		' Create a checksum of the given string
		Dim iKeyIndex As Short
		Dim sCheckString As String
		Dim sngRandom As Single
		Dim lngIndex As Integer
		Dim iCharCode As Short
		Dim lngTotalChar As Integer
		Dim aiKeys() As Short
		Dim fModify As Boolean
		Dim sTemp As String

		Const KEYCOUNT As Short = 10
		Const PREFIXKEYCOUNT As Short = 5

		ReDim aiKeys(KEYCOUNT)

		lngTotalChar = 0
		iKeyIndex = 1
		sCheckString = ""

		' Calculate the key digits.
		For lngIndex = 1 To UBound(aiKeys)
			Randomize()
			sngRandom = (Rnd() * 10) - 0.5
			aiKeys(lngIndex) = CShort(sngRandom)
		Next lngIndex

		' Do the checksum of the given string
		For lngIndex = 1 To Len(psString)
			iCharCode = Asc(Mid(psString, lngIndex, 1))

			iCharCode = iCharCode + (aiKeys(iKeyIndex) * (iKeyIndex + 1))

			lngTotalChar = lngTotalChar + iCharCode

			iKeyIndex = iKeyIndex + 1
			If iKeyIndex > UBound(aiKeys) Then
				iKeyIndex = 1
			End If
		Next lngIndex

		sCheckString = CStr(lngTotalChar)

		fModify = True
		sTemp = ""
		For lngIndex = 1 To UBound(aiKeys)
			If lngIndex <= PREFIXKEYCOUNT Then
				sTemp = sTemp & CStr(IIf(fModify, 9 - aiKeys(lngIndex), aiKeys(lngIndex)))
			Else
				sCheckString = sCheckString & CStr(IIf(fModify, 9 - aiKeys(lngIndex), aiKeys(lngIndex)))
			End If

			fModify = Not fModify
		Next lngIndex

		CheckString = sTemp & sCheckString

	End Function

	Public Function MM_AddToOutputArrayData(ByRef pstrValue As String) As Object
		Dim iNewIndex As Short
		iNewIndex = UBound(maryMMOutputArrayData) + 1
		ReDim Preserve maryMMOutputArrayData(iNewIndex)
		maryMMOutputArrayData(iNewIndex) = pstrValue
	End Function

	Public Function MM_AddToMergeFieldsData(ByRef pstrValue As String) As Object
		Dim iNewIndex As Short
		iNewIndex = UBound(maryMMMergeFieldsData) + 1
		ReDim Preserve maryMMMergeFieldsData(iNewIndex)
		maryMMMergeFieldsData(iNewIndex) = pstrValue
	End Function
	Public Function MM_DimArrays() As Object
		ReDim maryMMOutputArrayData(0)
		ReDim maryMMMergeFieldsData(0)
	End Function

Public CompletedDocumentName As String

	Public Function MM_DefCloseDoc(ByRef pblnValue As Boolean) As Object
		'mblnMMDefCloseDoc = pblnValue
	End Function

	Public WriteOnly Property MM_DefDocFile() As String
		Set(ByVal Value As String)
			'mstrMMDefDocFile = pstrValue
		End Set
	End Property

	Public WriteOnly Property MM_DefOutput() As Short
		Set(ByVal Value As Short)
			'mintMMDefOutput = pintValue
		End Set
	End Property

	Public ReadOnly Property MM_StatusMessage() As String
		Get
			MM_StatusMessage = mstrMMStatusMessage
		End Get
	End Property

	Public WriteOnly Property MM_MergeFieldsUbound() As Integer
		Set(ByVal Value As Integer)
			mlngMMMergeFieldsUbound = Value
		End Set
	End Property

	Public Property MM_DefName() As String
		Set(ByVal Value As String)
			mstrMMDefName = Value
		End Set
		Get
			Return mstrMMDefName
		End Get
	End Property

	Public WriteOnly Property MM_DefTemplateFile() As String
		Set(ByVal Value As String)
			mstrMMDefTemplateFile = Value
		End Set
	End Property

	Public WriteOnly Property MM_DefEmailSubject() As String
		Set(ByVal Value As String)
			mstrMMDefEmailSubject = Value
		End Set
	End Property

	Public WriteOnly Property MM_DefEmailAddrCalc() As Integer
		Set(ByVal Value As Integer)
			mlngMMDefEmailAddrCalc = Value
		End Set
	End Property

	Public WriteOnly Property MM_DefAttachmentName() As String
		Set(ByVal Value As String)
			mstrMMDefAttachmentName = Value
		End Set
	End Property

	Public WriteOnly Property MM_DefOutputFormat() As Short
		Set(ByVal Value As Short)
			mintMMDefOutputFormat = Value
		End Set
	End Property

	Public WriteOnly Property MM_DefOutputPrinterName() As String
		Set(ByVal Value As String)
			mstrMMDefOutputPrinterName = Value
		End Set
	End Property

	Public WriteOnly Property MM_DefOutputFileName() As String
		Set(ByVal Value As String)
			mstrMMDefOutputFileName = Value
		End Set
	End Property

	Public WriteOnly Property SaveAsValues() As String
		Set(ByVal Value As String)
			mstrSaveAsValues = Value
		End Set
	End Property

	Public ReadOnly Property LocaleDecimalSeparator() As String
		Get
			LocaleDecimalSeparator = msLocaleDecimalSeparator
		End Get
	End Property

	Public ReadOnly Property LocaleDateFormat() As String
		Get
			LocaleDateFormat = msLocaleDateFormat
		End Get
	End Property

	Public ReadOnly Property LocaleDateSeparator() As String
		Get
			LocaleDateSeparator = GetSystemDateSeparator()
		End Get
	End Property

	Public ReadOnly Property LocaleThousandSeparator() As String
		Get
			LocaleThousandSeparator = msLocaleThousandSeparator

		End Get
	End Property

	Public Function MM_DefDocSave(ByRef pblnValue As Boolean) As Object
		'mblnMMDefDocSave = pblnValue
	End Function

	Public Function MM_Cancelled() As Boolean
		MM_Cancelled = mblnMMCancelled
	End Function

	Public Function MM_DefPauseBeforeMerge(ByRef pblnValue As Object) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object pblnValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mblnMMDefPauseBeforeMerge = pblnValue
	End Function

	Public Function MM_DefSuppressBlankLines(ByRef pblnValue As Object) As Object
		mblnMMDefSuppressBlankLines = pblnValue
	End Function

	Public Function MM_DefEMailAttachment(ByRef pblnValue As Boolean) As Object
		mblnMMDefEMailAttachment = pblnValue
	End Function

	Public Function MM_DefOutputScreen(ByRef pblnValue As Boolean) As Object
		mblnMMDefOutputScreen = pblnValue
	End Function

	Public Function MM_DefOutputPrinter(ByRef pblnValue As Boolean) As Object
		mblnMMDefOutputPrinter = pblnValue
	End Function

	Public Function MM_DefOutputSave(ByRef pblnValue As Boolean) As Object
		mblnMMDefOutputSave = pblnValue
	End Function

	Private Function MM_WORDDOC_OpenTempate(ByRef pwrdApp As Microsoft.Office.Interop.Word.Application, ByRef pbSuppressPrompt As Boolean) As Boolean

		Dim bOK As Boolean = True

		Try

			If Not IO.File.Exists(mstrMMDefTemplateFile) Then
				mstrMMStatusMessage = "Error opening Word template file <" & Replace(mstrMMDefTemplateFile, "\", "\\") & ">."
				bOK = False
			Else

				mdocTemplate = pwrdApp.Documents.Open(mstrMMDefTemplateFile, False, False, False, "", "", False, "", "", 0)

				If mdocTemplate.MailMerge.MainDocumentType <> Microsoft.Office.Interop.Word.WdMailMergeMainDocType.wdNotAMergeDocument Then
					mstrMMStatusMessage = "The Word template file <" & mstrMMDefTemplateFile & "> has been saved referencing a different set of data." & vbCrLf & vbCrLf & "You need to remove this reference in order to proceed with this mail merge." & vbCrLf & vbCrLf
				End If

				mdocTemplate.MailMerge.MainDocumentType = Microsoft.Office.Interop.Word.WdMailMergeMainDocType.wdNotAMergeDocument

				pwrdApp.CustomizationContext = mdocTemplate
				mdocTemplate.Saved = True
			End If

		Catch ex As Exception
			mstrMMStatusMessage = "Error opening Word template file <" & Replace(mstrMMDefTemplateFile, "\", "\\") & ">."
			bOK = False

		End Try

		Return bOK

	End Function

	Public Function MM_WORD_CreateTemplateFile(ByRef psTemplatePath As String) As Boolean

		Dim bOK As Boolean = True

		Try

			mwrdApp = New Microsoft.Office.Interop.Word.Application
			mwrdApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone
			mdocTemplate = mwrdApp.Documents.Add

			mstrMMDefTemplateSaveAs = GetSaveAsFormat2(psTemplatePath, mstrSaveAsValues)

			If mstrMMDefTemplateSaveAs <> vbNullString Then
				mdocTemplate.SaveAs(psTemplatePath, Val(mstrMMDefTemplateSaveAs))
			Else
				mstrMMStatusMessage = "Error creating template file"
				bOK = False
			End If

			mdocTemplate.Close((False))
			mwrdApp.Quit((False))

		Catch ex As Exception
			mstrMMStatusMessage = "Error creating template file"
			bOK = False

		End Try

		Return bOK

	End Function

	Private Function MM_WORDDOC_ValidateTemplateMergeFields(ByRef pdocTemplate As Microsoft.Office.Interop.Word.Document, ByRef pblnAbortOnError As Boolean) As Boolean

		Dim fOk As Boolean
		Dim intCount As Short
		Dim strField As String
		Dim blnFieldOkay As Boolean

		On Error GoTo ErrorTrap

		fOk = True
		mstrMMStatusMessage = vbNullString

		If fOk Then
			intCount = 1
			While intCount <= pdocTemplate.Fields.Count
				strField = pdocTemplate.Fields.Item(intCount).Code.Text
				If Mid(strField, 2, 10) = "MERGEFIELD" Then
					strField = Mid(strField, 13)
					strField = RTrim(strField)
					strField = Replace(strField, """", "")
					strField = Replace(strField, " ", "_")
					strField = Left(strField, 39)

					blnFieldOkay = MM_FieldOK(strField)

					If Not blnFieldOkay Then
						mstrMMStatusMessage = "The Template Document '" & mstrMMDefTemplateFile & "' contains the merge field " & strField & " which was not selected in the mail merge definition."
						Return False

					End If
				End If

				intCount = intCount + 1
			End While

		End If

		Return True


ErrorTrap:
		Select Case Err.Number
			Case 462, -2147417848
				If Err.Number = 462 Then
					mwrdApp = New Microsoft.Office.Interop.Word.Application
				End If
				fOk = MM_WORDDOC_OpenTempate(mwrdApp, True)
				If fOk Then
					pdocTemplate = mwrdApp.ActiveDocument
					MM_WORDDOC_SetMergeOptions(pdocTemplate, mstrMMDataSourceName)
					pdocTemplate.Saved = True
				End If

			Case Else
				mstrMMStatusMessage = "Error Validating Merge Fields (" & Replace(Err.Description, "\", "\\") & ")"
		End Select

		Return False

	End Function

	Private Function MM_FieldOK(ByRef psFieldName As String) As Boolean

		Dim i As Short

		For i = 0 To UBound(maryMMMergeFieldsData)
			maryMMMergeFieldsData(i) = Trim(maryMMMergeFieldsData(i))
			maryMMMergeFieldsData(i) = Replace(maryMMMergeFieldsData(i), Chr(34), "")
			maryMMMergeFieldsData(i) = Replace(maryMMMergeFieldsData(i), " ", "_")
			maryMMMergeFieldsData(i) = Left(maryMMMergeFieldsData(i), 39)

			If mlngMMDefEmailAddrCalc > 0 Then
				If UCase(psFieldName) = UCase(maryMMMergeFieldsData(i)) Or UCase(psFieldName) = "EMAIL_ADDRESS" Then
					MM_FieldOK = True
					Exit Function
				End If
			Else
				If UCase(psFieldName) = UCase(maryMMMergeFieldsData(i)) Then
					MM_FieldOK = True
					Exit Function
				End If
			End If
		Next i

	End Function

	Private Function MM_WORDDOC_SetMergeOptions(ByRef pdocTemplate As Microsoft.Office.Interop.Word.Document, ByRef pstrDataSource As String) As Boolean

		Dim objTemp As Object

		pdocTemplate.MailMerge.OpenDataSource(mstrMMDataSourceName, , , , , False)
		pdocTemplate.MailMerge.SuppressBlankLines = mblnMMDefSuppressBlankLines

		If mintMMDefOutputFormat = 1 Then
			pdocTemplate.MailMerge.Destination = Microsoft.Office.Interop.Word.WdMailMergeDestination.wdSendToEmail
			pdocTemplate.MailMerge.MailAddressFieldName = "Email_Address"
			pdocTemplate.MailMerge.MailSubject = mstrMMDefEmailSubject
			pdocTemplate.MailMerge.MailAsAttachment = CBool(LCase(CStr(mblnMMDefEMailAttachment)))

			If Val(mwrdApp.Version) >= 12 Then
				pdocTemplate.MailMerge.DataSource.FirstRecord = Microsoft.Office.Interop.Word.WdMailMergeDefaultRecord.wdDefaultFirstRecord
				pdocTemplate.MailMerge.DataSource.LastRecord = Microsoft.Office.Interop.Word.WdMailMergeDefaultRecord.wdDefaultLastRecord

				If Not mblnMMDefEMailAttachment Then
					objTemp = pdocTemplate.MailMerge
					'UPGRADE_WARNING: Couldn't resolve default property of object objTemp.MailFormat. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					objTemp.MailFormat = 1 'wdMailFormatHTML
				End If
			End If

		Else
			pdocTemplate.MailMerge.Destination = Microsoft.Office.Interop.Word.WdMailMergeDestination.wdSendToNewDocument
		End If

		Return True

	End Function

	Public Function MM_WORD_ExecuteMailMerge() As Boolean

		On Error GoTo ErrorTrap

		Dim fOk As Boolean
		fOk = True

		mblnMMCancelled = False

		'Open Word application
		mstrMMStatusMessage = "Error opening Word application."
		mwrdApp = New Microsoft.Office.Interop.Word.Application
		mwrdApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone

		'Check Word version
		If Val(mwrdApp.Version) < 8 Then
			mstrMMStatusMessage = "You are running a version of Word which is not supported by HR-Pro."
			fOk = False
			GoTo TidyUpAndExit
		End If

		'Check Template File
		mstrMMDefTemplateSaveAs = GetSaveAsFormat2(mstrMMDefTemplateFile, mstrSaveAsValues)
		If mstrMMDefTemplateSaveAs = vbNullString Then
			mstrMMStatusMessage = "This definition is set to use a template file which is not compatible with your version of Microsoft Office."
			fOk = False
			GoTo TidyUpAndExit
		End If

		'Check Output Filename
		If Trim(mstrMMDefOutputFileName) <> vbNullString Then
			mstrMMDefOutputSaveAs = GetSaveAsFormat2(mstrMMDefOutputFileName, mstrSaveAsValues)
			If mstrMMDefOutputSaveAs = vbNullString Then
				mstrMMStatusMessage = "This definition is set to save in a file format which is not compatible with your version of Microsoft Office."
				fOk = False
				GoTo TidyUpAndExit
			End If
		End If


		If Not ValidPrinter(mstrMMDefOutputPrinterName) Then
			mstrMMStatusMessage = "This definition is set to output to printer " & mstrMMDefOutputPrinterName & " which is not set up on your PC."
			fOk = False
			GoTo TidyUpAndExit
		End If


		'Check Email AttachAs
		If Trim(mstrMMDefAttachmentName) <> vbNullString Then
			mstrMMDefAttachmentFormat = GetSaveAsFormat2(mstrMMDefAttachmentName, mstrSaveAsValues)
			If mstrMMDefAttachmentFormat = vbNullString Then
				mstrMMStatusMessage = "This definition is set to email an attachment in a file format which is not compatible with your version of Microsoft Office."
				fOk = False
				GoTo TidyUpAndExit
			End If
		End If

		If Not MM_WORD_Execute_Step1() Then
			fOk = False
			GoTo ErrorTrap
		End If

		'If Not MM_WORD_Execute_Step2() Then
		'	fOk = False
		'	GoTo ErrorTrap
		'End If

		If Not MM_WORD_Execute_Step3() Then
			fOk = False
			GoTo ErrorTrap
		End If

		If Not MM_WORD_Execute_Step4() Then
			fOk = False
			GoTo ErrorTrap
		End If

TidyUpAndExit:
		'UPGRADE_NOTE: Object mdocDataSource may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mdocDataSource = Nothing
		'UPGRADE_NOTE: Object mdocTemplate may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mdocTemplate = Nothing
		Return fOk

ErrorTrap:
		fOk = False
		'Kill Word Documents
		'UPGRADE_NOTE: Object mdocDataSource may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mdocDataSource = Nothing
		'UPGRADE_NOTE: Object mdocTemplate may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mdocTemplate = Nothing
		'Kill Word Application in correct order.
		If Not mwrdApp Is Nothing Then mwrdApp.Quit(False)
		'UPGRADE_NOTE: Object mwrdApp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mwrdApp = Nothing
		GoTo TidyUpAndExit

	End Function

	Private Function ValidPrinter(ByRef strName As String) As Boolean

		Dim objPrinter As Printer
		Dim blnFound As Boolean

		If strName <> vbNullString And strName <> "<Default Printer>" Then
			blnFound = False
			For Each objPrinter In Printers
				If objPrinter.DeviceName = strName Then
					blnFound = True
					Exit For
				End If
			Next objPrinter
		Else
			blnFound = True
		End If

		ValidPrinter = blnFound

	End Function

	Private Function MM_WORD_Execute_Step1() As Boolean

		On Error GoTo ErrorTrap

		Dim fOk As Boolean = True
		Dim strDataSourceFormat As String

		'Microsoft Word Declarations
		Dim intCount As Short
		Dim lngTableCols As Integer
		Dim intMBResponse As MsgBoxResult
		Dim blnMergeFieldExists As Boolean
		Dim strField As String

		'Create Data Source
		mstrMMStatusMessage = "Error creating data source."
		mstrMMDataSourceName = GetTempFile()
		mdocDataSource = mwrdApp.Documents.Add()

		strDataSourceFormat = GetSaveAsFormat2(mstrMMDataSourceName, mstrSaveAsValues)
		If strDataSourceFormat = "" Then strDataSourceFormat = "0"
		mdocDataSource.SaveAs(mstrMMDataSourceName, Val(strDataSourceFormat))

		'Populate Data Source
		mstrMMStatusMessage = "Error populating data source."
		If mlngMMMergeFieldsUbound < 3 Then
			mdocDataSource.ActiveWindow.Selection.Bookmarks.Add("TableStart", mdocDataSource.ActiveWindow.Selection.Range)
		End If

		For intCount = 1 To UBound(maryMMOutputArrayData) Step 1
			If intCount > 1 Then
				mdocDataSource.ActiveWindow.Selection.TypeParagraph()
			End If
			mdocDataSource.ActiveWindow.Selection.TypeText((maryMMOutputArrayData(intCount)))
		Next intCount

		lngTableCols = mlngMMMergeFieldsUbound
		If mintMMDefOutputFormat = 1 Then	'Email
			lngTableCols = lngTableCols + 1
		End If
		If lngTableCols < 3 Then
			mdocDataSource.ActiveWindow.Selection.GoTo(Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, , , "TableStart")
			mdocDataSource.ActiveWindow.Selection.EndKey(Microsoft.Office.Interop.Word.WdUnits.wdStory, Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
			mdocDataSource.ActiveWindow.Selection.ConvertToTable(Microsoft.Office.Interop.Word.WdTableFieldSeparator.wdSeparateByTabs, , , , Microsoft.Office.Interop.Word.WdTableFormat.wdTableFormatNone, , , False, False, , , , , False)
		End If
		mdocDataSource.SaveAs(FileFormat:=Val(strDataSourceFormat))
		mdocDataSource.Close((False))

		'Call open template
		fOk = MM_WORDDOC_OpenTempate(mwrdApp, False)
		If fOk Then
			mdocTemplate = mwrdApp.ActiveDocument
		End If

		If fOk Then
			If mdocTemplate.ReadOnly Then
				fOk = MM_WORDDOC_ValidateTemplateMergeFields(mdocTemplate, True)
				If fOk Then
					MM_WORDDOC_SetMergeOptions(mdocTemplate, mstrMMDataSourceName)
				End If
			Else
				fOk = MM_WORDDOC_ValidateTemplateMergeFields(mdocTemplate, False)
				If fOk Then
					MM_WORDDOC_SetMergeOptions(mdocTemplate, mstrMMDataSourceName)
				End If
			End If

			mdocTemplate.Saved = True

		End If

		'Pause before merge
		intMBResponse = MsgBoxResult.No
		'If mblnMMDefPauseBeforeMerge Then
		'	If fOk Then
		'		intMBResponse = MsgBox("Would you like to amend the document or merge fields?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Question, "Mail Merge")
		'		If intMBResponse = MsgBoxResult.Cancel Then
		'			mblnMMCancelled = True
		'			fOk = False
		'		End If
		'	End If
		'End If


CheckMergeFieldExists:

		'Check that there is at least one merge field
		If fOk Then

			intCount = 1
			blnMergeFieldExists = False

			Do While blnMergeFieldExists = False And mblnMMCancelled = False

				While ((intCount <= mdocTemplate.Fields.Count) And (Not blnMergeFieldExists))
					strField = mdocTemplate.Fields.Item(intCount).Code.Text
					If Mid(strField, 2, 10) = "MERGEFIELD" Then
						blnMergeFieldExists = True
					End If
					intCount = intCount + 1
				End While

				If Not blnMergeFieldExists Then
					mstrMMStatusMessage = "There are no merge fields specified in the template document."
					mblnMMCancelled = True
				End If
			Loop

			MM_WORDDOC_ValidateTemplateMergeFields(mdocTemplate, False)

		End If

TidyUpAndExit:
		Return fOk

ErrorTrap:

		Select Case Err.Number
			Case 462, -2147417848
				If Err.Number = 462 Then
					mwrdApp = New Microsoft.Office.Interop.Word.Application
				End If
				fOk = MM_WORDDOC_OpenTempate(mwrdApp, True)
				If fOk Then
					mdocTemplate = mwrdApp.ActiveDocument
					fOk = MM_WORDDOC_SetMergeOptions(mdocTemplate, mstrMMDataSourceName)

					'MH20050323 Fault 9909
					'mdocTemplate.Saved = True
					Resume CheckMergeFieldExists
				End If

			Case Else
				fOk = False
				'mstrMMStatusMessage = Replace(Err.Description, "\", "\\") & " (Step 1)"
				mstrMMStatusMessage = Err.Description
		End Select

		GoTo TidyUpAndExit

	End Function

'	Private Function MM_WORD_Execute_Step2() As Boolean

'		On Error GoTo ErrorTrap

'		Dim fOk As Boolean
'		Dim blnTemplateReadOnly As Boolean
'		Dim intMBResponse As MsgBoxResult
'		Dim intErrTrapCode As Short

'		fOk = True

'		intErrTrapCode = 0

'		mwrdApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsAll
'		blnTemplateReadOnly = False
'		intMBResponse = MsgBoxResult.Yes
'		While ((mdocTemplate.Saved = False) And (intMBResponse = MsgBoxResult.Yes))

'			intMBResponse = MsgBox("You have not saved changes to the template document, " & "Would you like to save changes now?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Mail Merge")
'			intErrTrapCode = 1
'			If intMBResponse = MsgBoxResult.Yes Then
'				mwrdApp.Visible = True
'				mwrdApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsAll
'				mdocTemplate.SaveAs(FileFormat:=Val(mstrMMDefTemplateSaveAs))
'			End If

'			If blnTemplateReadOnly = True Then
'				intMBResponse = MsgBox("This template file is currently in use by another user." & "Please save as a new file.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Mail Merge")
'				intErrTrapCode = 2
'				mdocTemplate.SaveAs(FileFormat:=Val(mstrMMDefTemplateSaveAs))
'				blnTemplateReadOnly = False
'				intMBResponse = MsgBoxResult.Yes
'			End If

'			mwrdApp.Visible = False
'			mwrdApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone

'		End While


'TidyUpAndExit:
'		MM_WORD_Execute_Step2 = fOk
'		Exit Function

'ErrorTrap:
'		Select Case Err.Number
'			Case 462, -2147417848
'				If Err.Number = 462 Then
'					mwrdApp = New Microsoft.Office.Interop.Word.Application
'				End If
'				fOk = MM_WORDDOC_OpenTempate(mwrdApp, True)
'				If fOk Then
'					mdocTemplate = mwrdApp.ActiveDocument
'					MM_WORDDOC_SetMergeOptions(mdocTemplate, mstrMMDataSourceName)
'					mdocTemplate.Saved = True
'					Resume Next
'				End If
'			Case 4198
'			Case 5155
'			Case 5356
'				If intErrTrapCode = 1 Then
'					blnTemplateReadOnly = True
'					Resume Next
'				End If
'			Case Else
'				fOk = False
'				mstrMMStatusMessage = "Error checking if the template file has been saved (" & intErrTrapCode & ")."
'		End Select

'		GoTo TidyUpAndExit

'	End Function

	Private Function MM_WORD_Execute_Step3() As Boolean

		Dim strField As String
		Dim blnMergeFieldExists As Boolean
		Dim intCount As Short
		Dim strOriginalDefaultPrinter As String

		'Run the Merge.

		On Error GoTo ErrorTrap

		Dim fOk As Boolean = True

		If fOk Then

			'MH20050323 Fault 9909
			intCount = 1
			blnMergeFieldExists = False
			While ((intCount <= mdocTemplate.Fields.Count) And (Not blnMergeFieldExists))
				strField = mdocTemplate.Fields.Item(intCount).Code.Text
				If Mid(strField, 2, 10) = "MERGEFIELD" Then
					blnMergeFieldExists = True
				End If
				intCount = intCount + 1
			End While

			If Not blnMergeFieldExists Then
				mstrMMStatusMessage = "No merge fields specified in the template document."
				fOk = False
				Exit Function
			End If



			mwrdApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsAll

			'Rename attachment
			'Rename attachment
			If Trim(mstrMMDefAttachmentName) <> vbNullString Then

				'Get temp path
				mstrMMTempFileName = Space(1024)
				mstrMMTempFileName = GetTempFile()

				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				If Dir(mstrMMDefAttachmentName) <> vbNullString Then
					Kill((mstrMMDefAttachmentName))
				End If

				mdocTemplate.SaveAs(mstrMMDefAttachmentName, Val(mstrMMDefAttachmentFormat))

			End If

			mdocTemplate.MailMerge.Execute(False)
			mdocOutput = mwrdApp.ActiveDocument
			mdocOutput.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView


			If mblnMMDefOutputSave Then
				'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				If Dir(mstrMMDefOutputFileName) <> vbNullString Then
					Kill((mstrMMDefOutputFileName))
				End If
				mdocOutput.SaveAs(mstrMMDefOutputFileName, Val(mstrMMDefOutputSaveAs))
				While (mwrdApp.BackgroundSavingStatus > 0)
					Windows.Forms.Application.DoEvents()
				End While
			End If


			If mblnMMDefOutputPrinter Then
				strOriginalDefaultPrinter = vbNullString
				If Trim(mstrMMDefOutputPrinterName) <> vbNullString Then
					strOriginalDefaultPrinter = mwrdApp.ActivePrinter
					mwrdApp.ActivePrinter = mstrMMDefOutputPrinterName
				End If

				mdocOutput.PrintOut()
				While (mwrdApp.BackgroundPrintingStatus > 0)
					Windows.Forms.Application.DoEvents()
				End While

				If Trim(strOriginalDefaultPrinter) <> vbNullString Then
					mwrdApp.ActivePrinter = strOriginalDefaultPrinter
				End If
			End If

		End If


TidyUpAndExit:
		MM_WORD_Execute_Step3 = fOk
		Exit Function

ErrorTrap:
		Select Case Err.Number
			Case 462
				mstrMMStatusMessage = "Microsoft Word has been closed by the user."
				mwrdApp = New Microsoft.Office.Interop.Word.Application

			Case -2147417848
				mstrMMStatusMessage = "Microsoft Word document has been closed by the user."

			Case 4605
				mstrMMStatusMessage = "No merge fields specified in the template document."

			Case 5152
				mstrMMStatusMessage = "Error saving Mail Merge to file" & vbCrLf & mstrMMDefOutputFileName & vbCrLf & "Please ensure that the output document path is correct within this definition."

			Case 5630
				mstrMMStatusMessage = "Email field is not specified in the merge data."

			Case Else
				mstrMMStatusMessage = Replace(Err.Description, "\", "\\")

		End Select
		fOk = False

		GoTo TidyUpAndExit

	End Function

	Private Function MM_WORD_Execute_Step4() As Boolean

		On Error GoTo ErrorTrap

		Dim fOk As Boolean
		Dim blnLeaveDocOpen As Boolean

		fOk = True

		mwrdApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone

		'Make sure that the template file does not still reference the data source
		'(close it and open it incase there have been changes to it or that the
		'the template is now the temporary email attachment name thingy!
		On Error Resume Next
		mdocTemplate.Close(False)
		If mstrMMDefTemplateSaveAs <> vbNullString Then
			mdocTemplate = mwrdApp.Documents.Open(mstrMMDefTemplateFile, False, False) ', False, False, False, False, False, False, False, False, 0)
			mdocTemplate.MailMerge.MainDocumentType = Microsoft.Office.Interop.Word.WdMailMergeMainDocType.wdNotAMergeDocument
			mdocTemplate.Saved = False 'MH20070227 Fault 12013 - Force save for Office 2007
			mdocTemplate.SaveAs(mstrMMDefTemplateFile, Val(mstrMMDefTemplateSaveAs))
			mdocTemplate.Close((False))
		End If

		'If the temorary email attachment name exists then kill it
		If mstrMMDefAttachmentName <> vbNullString Then
			Kill((mstrMMTempFileName))
		End If

		'mdocDataSource.Close()

		On Error GoTo ErrorTrap
		'Kill the data source
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Dir(mstrMMDataSourceName) <> vbNullString Then
			Kill((mstrMMDataSourceName))
		End If

		blnLeaveDocOpen = (mblnMMDefOutputScreen And fOk = True)

		' Save to temp object
		CompletedDocumentName = My.Computer.FileSystem.GetTempFileName
		mdocOutput.SaveAs(CompletedDocumentName)

		''On Error Resume Next
		'If blnLeaveDocOpen Then
		'	mwrdApp.Visible = True
		'	mwrdApp.WindowState = Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateNormal
		'	mwrdApp.Activate()
		'	mdocOutput.Activate()

		'Else
			mdocOutput.Close(False)
			'UPGRADE_NOTE: Object mdocOutput may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			mdocOutput = Nothing
			'UPGRADE_NOTE: Object mdocTemplate may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			mdocTemplate = Nothing
			mwrdApp.Quit()
			'UPGRADE_NOTE: Object mwrdApp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			mwrdApp = Nothing
	'	End If

TidyUpAndExit:
		MM_WORD_Execute_Step4 = fOk
		Exit Function

ErrorTrap:
		MsgBox(Err.Description)
		fOk = False
		mwrdApp.Quit()
		'UPGRADE_NOTE: Object mwrdApp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mwrdApp = Nothing
		GoTo TidyUpAndExit

	End Function

	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object mdocDataSource may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mdocDataSource = Nothing
		'UPGRADE_NOTE: Object mdocOutput may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mdocOutput = Nothing
		'UPGRADE_NOTE: Object mdocTemplate may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mdocTemplate = Nothing
		'UPGRADE_NOTE: Object mwrdApp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mwrdApp = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub

	Public Function SelectFolder(ByRef psInitDir As String, ByRef psPrompt As String, ByRef psTitle As String) As String
	Return psInitDir
'		Dim fOk As Boolean
'		Dim frmDir As New frmPathSel

'		frmDir.Text = psTitle
'		If ValidateDir(psInitDir) Then
'			frmDir.lblPrompt.Text = "Please select a folder for the '" & psTitle & "'"

'		Else
'			frmDir.lblPrompt.Text = "The '" & psTitle & "' has either not been set, or does not exist. " & "Please select a folder for the '" & psTitle & "'."
'			psInitDir = vbNullString
'		End If

'		fOk = frmDir.Initialise(psInitDir)

'		If fOk Then
'			frmDir.ShowDialog()
'			SelectFolder = frmDir.SelectedFolder
'		End If

'TidyUpAndExit:
'		'UPGRADE_NOTE: Object frmDir may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
'		frmDir = Nothing
'		Exit Function

'ErrorTrap:
'		SelectFolder = vbNullString
'		GoTo TidyUpAndExit

	End Function

	'Public Function SendMail(ByRef psTo As String, Optional ByRef psSubject As String = "", Optional ByRef psBody As String = "", Optional ByRef psCC As String = "", Optional ByRef psBCC As String = "") As Boolean

	'	Dim strError As String

	'	strError = SendEmailFromClientUsingMAPI(psTo, psCC, psBCC, psSubject, psBody, "", True)
	'	If strError <> vbNullString Then
	'		MsgBox(strError, MsgBoxStyle.Exclamation, "Send Mail Message")
	'		SendMail = False
	'	Else
	'		SendMail = True
	'	End If

	'End Function



'	Public Function SendEmailFromClientUsingMAPI(ByRef strTo As String, ByRef strCC As String, ByRef strBCC As String, ByRef strSubject As String, ByRef strMsgText As String, ByRef strAttachment As String, ByRef blnShowMessage As Boolean) As String


'		Dim objMapiSession As AxMSMAPI.AxMAPISession
'		Dim objMapiMessages As AxMSMAPI.AxMAPIMessages

'		Dim strError As String
'		Dim strRecipients As String
'		Dim lngRecipType As Integer

'		Dim strArray() As String
'		Dim lngIndex As Integer
'		'Dim blnShowMessage As Boolean

'		On Error GoTo LocalErr

'		objMapiSession = frmEmailSel.MAPISession1
'		objMapiMessages = frmEmailSel.MAPIMessages1


'		strError = vbNullString
'		'blnShowMessage = False

'		If Trim(Replace(strTo, ";", "")) = vbNullString Then
'			SendEmailFromClientUsingMAPI = "Please select recipient(s) to email"
'			Exit Function
'		End If


'		If objMapiSession.SessionID = 0 Then
'			'objMapiSession.DownLoadMail = False
'			objMapiSession.SignOn()
'			objMapiMessages.SessionID = objMapiSession.SessionID
'		End If


'		With objMapiMessages
'			.AddressResolveUI = True
'			.Compose()

'			For lngRecipType = 1 To 3

'				Select Case lngRecipType
'					Case 1 : strRecipients = strTo
'					Case 2 : strRecipients = strCC
'					Case 3 : strRecipients = strBCC
'				End Select

'				If (Trim(strRecipients) <> vbNullString) Then
'					strArray = Split(strRecipients, ";")
'					For lngIndex = LBound(strArray) To UBound(strArray)
'						If Trim(strArray(lngIndex)) <> vbNullString Then
'							.RecipIndex = .RecipCount

'							.RecipType = lngRecipType
'							.RecipDisplayName = Trim(strArray(lngIndex))
'							.RecipAddress = "smtp:" & Trim(strArray(lngIndex))
'							'.ResolveName



'						End If
'					Next
'				End If

'			Next

'			.MsgSubject = strSubject
'			.MsgNoteText = strMsgText & " "

'			If (Trim(strAttachment) <> vbNullString) Then
'				strArray = Split(strAttachment, ";")
'				For lngIndex = LBound(strArray) To UBound(strArray)
'					If Trim(strArray(lngIndex)) <> vbNullString Then
'						.AttachmentIndex = lngIndex
'						.AttachmentPathName = Trim(strArray(lngIndex))
'						'.AttachmentName = FileOnlyFromFullPath(Trim(strArray(lngIndex)))
'						'.AttachmentPosition = lngIndex
'						'.AttachmentType = 0
'					End If
'				Next
'			End If

'			.Send(blnShowMessage)

'		End With

'TidyAndExit:
'		If objMapiSession.SessionID <> 0 Then
'			objMapiSession.SignOff()
'		End If
'		objMapiMessages.SessionID = 0

'		SendEmailFromClientUsingMAPI = strError

'		Exit Function

'LocalErr:
'		If blnShowMessage = False Then
'			'Simple MAPI is no longer supported in Microsoft Outlook 2007. It is still supported by Exchange Server 2003.
'			'http://msdn.microsoft.com/en-us/library/cc815424.aspx
'			blnShowMessage = True
'			Resume

'		ElseIf Err.Number = 32001 Or Err.Number = 32003 Then
'			Resume Next

'		Else
'			strError = "Error sending email " & vbCrLf & Err.Description & " (MAPI)"

'			GoTo TidyAndExit

'		End If

'	End Function

	Function GetSystemDecimalSeparator() As String
		' Return the system decimal separator.
		Dim lngLength As Integer
		Dim sBuffer As String = StrDup(100, " ")

		lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, sBuffer, 99)
		GetSystemDecimalSeparator = Left(sBuffer, lngLength - 1)

	End Function

	Public Function ConvertSQLDateToLocale(ByRef psSQLDate As String) As String
		' Convert the given date string (mm/dd/yyyy) into the locale format.
		' NB. This function assumes a sensible locale format is used.
		Dim fDaysDone As Boolean
		Dim fMonthsDone As Boolean
		Dim fYearsDone As Boolean
		Dim iLoop As Short
		Dim sFormattedDate As String

		sFormattedDate = ""

		' Get the locale's date format.
		fDaysDone = False
		fMonthsDone = False
		fYearsDone = False

		For iLoop = 1 To Len(msLocaleDateFormat)
			Select Case UCase(Mid(msLocaleDateFormat, iLoop, 1))
				Case "D"
					If Not fDaysDone Then
						sFormattedDate = sFormattedDate & Mid(psSQLDate, 4, 2)
						fDaysDone = True
					End If

				Case "M"
					If Not fMonthsDone Then
						sFormattedDate = sFormattedDate & Mid(psSQLDate, 1, 2)
						fMonthsDone = True
					End If

				Case "Y"
					If Not fYearsDone Then
						sFormattedDate = sFormattedDate & Mid(psSQLDate, 7, 4)
						fYearsDone = True
					End If

				Case Else
					sFormattedDate = sFormattedDate & Mid(msLocaleDateFormat, iLoop, 1)
			End Select
		Next iLoop

		ConvertSQLDateToLocale = sFormattedDate

	End Function

	Public Function ConvertSQLDateToTime(ByRef psSQLDate As String) As String

		Dim sTempDate As String

		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		sTempDate = IIf(IsDBNull(psSQLDate), "", psSQLDate)

		ConvertSQLDateToTime = VB6.Format(sTempDate, "hh:nn")

	End Function

	Public Sub SaveRegistrySetting(ByRef psAppName As String, ByRef psSection As String, ByRef psKey As String, ByRef pvValue As Object)
		' Save the given value to the registry with the given registry key values.
		'UPGRADE_WARNING: Couldn't resolve default property of object pvValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SaveSetting(psAppName, psSection, psKey, pvValue)

	End Sub

	Public Function GetRegistrySetting(ByRef psAppName As String, ByRef psSection As String, ByRef psKey As String) As String
		' Get the required value from the registry with the given registry key values.
'		GetRegistrySetting = GetSetting(psAppName, psSection, psKey)
		Return ""
	End Function

	Public Function GetTempFile() As String
		' Return a temporary file name.
		Dim sTmpPath As String
		Dim sTmpName As String

		sTmpPath = Space(1024)
		sTmpName = Space(1024)

		Call GetTempPath(1024, sTmpPath)
		Call GetTempFileName(sTmpPath, "_T", 0, sTmpName)

		sTmpName = Trim(sTmpName)
		If Len(sTmpName) > 0 Then
			sTmpName = Left(sTmpName, Len(sTmpName) - 1)
		Else
			sTmpName = vbNullString
		End If

		GetTempFile = Trim(sTmpName)

	End Function

	Public Function ValidateDir(ByRef psDir As String) As Boolean

		Dim fso As New Scripting.FileSystemObject

		On Error Resume Next

		ValidateDir = False
		ValidateDir = fso.FolderExists(psDir)
		'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		fso = Nothing

	End Function

	Public Function ValidateFilePath(ByRef psDir As String) As Boolean

		Dim fso As New Scripting.FileSystemObject

		On Error Resume Next

		ValidateFilePath = False
		ValidateFilePath = fso.FileExists(psDir)
		'UPGRADE_NOTE: Object fso may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		fso = Nothing

	End Function

	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		GetLocaleDateFormat()
		GetLocaleNumberFormat()
		mstrSaveAsValues = "doc|0|dot|1|xls|56|xlt|17"
	End Sub

	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub

	Private Sub GetLocaleDateFormat()
		' Returns the date format.
		' NB. Windows allows the user to configure totally stupid
		' date formats (eg. d/M/yyMydy !). This function does not cater
		' for such stupidity, and simply takes the first occurence of the
		' 'd', 'M', 'y' characters.
		Dim sSysFormat As String
		Dim sSysDateSeparator As String
		Dim sDateFormat As String
		Dim iLoop As Short
		Dim fDaysDone As Boolean
		Dim fMonthsDone As Boolean
		Dim fYearsDone As Boolean

		fDaysDone = False
		fMonthsDone = False
		fYearsDone = False
		sDateFormat = ""

		sSysFormat = GetSystemDateFormat()
		sSysDateSeparator = GetSystemDateSeparator()

		' Loop through the string picking out the required characters.
		For iLoop = 1 To Len(sSysFormat)
			Select Case Mid(sSysFormat, iLoop, 1)
				Case "d"
					If Not fDaysDone Then
						' Ensure we have two day characters.
						sDateFormat = sDateFormat & "dd"
						fDaysDone = True
					End If

				Case "M"
					If Not fMonthsDone Then
						' Ensure we have two month characters.
						sDateFormat = sDateFormat & "mm"
						fMonthsDone = True
					End If

				Case "y"
					If Not fYearsDone Then
						' Ensure we have four year characters.
						sDateFormat = sDateFormat & "yyyy"
						fYearsDone = True
					End If

				Case Else
					sDateFormat = sDateFormat & Mid(sSysFormat, iLoop, 1)

			End Select
		Next iLoop

		' Ensure that all day, month and year parts of the date
		' are present in the format.
		If Not fDaysDone Then
			If Mid(sDateFormat, Len(sDateFormat), 1) <> sSysDateSeparator Then
				sDateFormat = sDateFormat & sSysDateSeparator
			End If

			sDateFormat = sDateFormat & "dd"
		End If

		If Not fMonthsDone Then
			If Mid(sDateFormat, Len(sDateFormat), 1) <> sSysDateSeparator Then
				sDateFormat = sDateFormat & sSysDateSeparator
			End If

			sDateFormat = sDateFormat & "mm"
		End If

		If Not fYearsDone Then
			If Mid(sDateFormat, Len(sDateFormat), 1) <> sSysDateSeparator Then
				sDateFormat = sDateFormat & sSysDateSeparator
			End If

			sDateFormat = sDateFormat & "yyyy"
		End If

		' Return the date format.
		msLocaleDateFormat = sDateFormat

	End Sub

	Private Sub GetLocaleNumberFormat()

		msLocaleDecimalSeparator = GetSystemDecimalSeparator()
		msLocaleThousandSeparator = GetSystemThousandSeparator()

	End Sub

	Function GetSystemThousandSeparator() As String
		' Return the system data separator.
		Dim lngLength As Integer
		Dim sBuffer As String = StrDup(100, " ")

		lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STHOUSAND, sBuffer, 99)
		GetSystemThousandSeparator = Left(sBuffer, lngLength - 1)

	End Function

	Private Function GetSystemDateSeparator() As String
		' Return the system data separator.
		Dim lngLength As Integer
		Dim sBuffer As String = StrDup(100, " ")

		lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDATE, sBuffer, 99)
		GetSystemDateSeparator = Left(sBuffer, lngLength - 1)

	End Function

	Private Function GetSystemDateFormat() As String
		' Return the system data format.
		Dim lngLength As Integer
		Dim sBuffer As String = StrDup(100, " ")

		lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE, sBuffer, 99)
		GetSystemDateFormat = Left(sBuffer, lngLength - 1)

	End Function

	Public Function PrinterCount() As Integer
		PrinterCount = Printers.Count
	End Function

	Public Function PrinterName(ByRef lngIndex As Integer) As String
		PrinterName = Printers(lngIndex).DeviceName
	End Function

	Public Function GetPCSetting(ByRef strSection As String, ByRef strKey As String, ByRef varDefault As Object) As String
		'UPGRADE_WARNING: Couldn't resolve default property of object varDefault. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
	'	GetPCSetting = GetSetting("HR Pro", strSection, strKey, varDefault)
	Return varDefault
	End Function

	Public Function SavePCSetting(ByRef strSection As String, ByRef strKey As String, ByRef varSetting As Object) As Boolean
		'Trap error in case user doesn't have permission to write to the registry
		On Error Resume Next
		'UPGRADE_WARNING: Couldn't resolve default property of object varSetting. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		SaveSetting("HR Pro", strSection, strKey, varSetting)
	End Function

	Private Function GetUNCOnly(ByVal pstrFileName As String) As String

		On Error GoTo GetUNCPath_Err
		Dim strMSG As String
		Dim lngReturn As Integer
		Dim strLocalName As String
		Dim strRemoteName As String
		Dim lngRemoteName As Integer
		Dim strUNCPath As String
		strLocalName = GetDriveOnly(pstrFileName)
		strRemoteName = New String(Chr(32), 255)
		lngRemoteName = Len(strRemoteName)

		'Attempt to grab UNC
		lngReturn = WNetGetConnection(strLocalName, strRemoteName, lngRemoteName)

		If lngReturn = 0 Then
			GetUNCOnly = Trim(Replace(strRemoteName, Chr(0), ""))

			' Local path
		ElseIf lngReturn = 2250 Then
			GetUNCOnly = GetDriveOnly(pstrFileName)
		Else
			GetUNCOnly = Trim(strLocalName)
		End If

GetUNCPath_End:
		Exit Function

GetUNCPath_Err:
		GetUNCOnly = Trim(strLocalName)
		Resume GetUNCPath_End
	End Function

	Public Function ConvertToUNC(ByVal pstrFileName As String) As String

		On Error GoTo ConvertUNCPath_Err
		Dim strMSG As String
		Dim lngReturn As Integer
		Dim strLocalName As String
		Dim strRemoteName As String
		Dim lngRemoteName As Integer
		Dim strUNCPath As String
		strLocalName = GetDriveOnly(pstrFileName)
		strRemoteName = New String(Chr(32), 255)
		lngRemoteName = Len(strRemoteName)

		'Attempt to grab UNC
		lngReturn = WNetGetConnection(strLocalName, strRemoteName, lngRemoteName)

		If lngReturn = 0 Then
			ConvertToUNC = Trim(Replace(strRemoteName, Chr(0), "")) & GetPathOnly(pstrFileName, True) & "\" & GetFileNameOnly(pstrFileName)

			' Local path
		ElseIf lngReturn = 2250 Then
			ConvertToUNC = pstrFileName
		Else
			ConvertToUNC = pstrFileName
		End If

ConvertUNCPath_End:
		Exit Function

ConvertUNCPath_Err:
		ConvertToUNC = pstrFileName
		Resume ConvertUNCPath_End
	End Function

	' Extracts just the filename from a path
	Public Function GetFileNameOnly(ByRef pstrFilePath As String) As String
		Dim astrPath() As String
		astrPath = Split(pstrFilePath, "\")
		GetFileNameOnly = astrPath(UBound(astrPath))
	End Function

	Public Function IsFileOnNetwork(ByVal pstrFileName As String) As Short

		Dim lngReturn As Integer
		Dim strLocalName As String
		Dim strRemoteName As String
		Dim lngRemoteName As Integer

		strLocalName = GetDriveOnly(pstrFileName)
		strRemoteName = New String(Chr(32), 255)
		lngRemoteName = Len(strRemoteName)

		'Attempt to grab UNC
		lngReturn = WNetGetConnection(strLocalName, strRemoteName, lngRemoteName)

		If lngReturn = 0 Then
			IsFileOnNetwork = 1

			' Local path
		ElseIf lngReturn = 2250 Then
			IsFileOnNetwork = 0
		Else
			IsFileOnNetwork = 0
		End If

	End Function

	Public Function GetDriveOnly(ByVal pstrFileName As String) As String

		If Mid(pstrFileName, 2, 1) = ":" Then
			GetDriveOnly = Mid(pstrFileName, 1, 1) & ":"
		Else
			GetDriveOnly = ""
		End If

	End Function

	' Extracts the path from a given filename
	Public Function GetPathOnly(ByRef pstrFilePath As String, ByRef pbStripDriveLetter As Boolean) As String

		Dim L As Short
		Dim tempchar As String
		Dim strPath As String

		L = Len(pstrFilePath)

		While L > 0
			tempchar = Mid(pstrFilePath, L, 1)
			If tempchar = "\" Then
				strPath = Mid(pstrFilePath, 1, L - 1)

				' Strip off drive letter
				If pbStripDriveLetter And Mid(strPath, 2, 1) = ":" Then
					strPath = Mid(strPath, 3, Len(strPath))
				End If

				GetPathOnly = strPath

				Exit Function
			End If
			L = L - 1
		End While

	End Function

	'' Wrapper for the encryption module as we don't want users to be able to see the super secret key
	'Public Function OLEEncryptFile(ByRef strInFile As String, ByRef strUploadPath As String, ByRef strKey As String) As String

	'	Dim objEncrypt As New clsEncryption
	'	Dim strOutFile As String

	'	strOutFile = Space(1024)
	'	Call GetTempFileName(strUploadPath, "_T", 0, strOutFile)
	'	strOutFile = Left(strOutFile, InStr(strOutFile, Chr(0)) - 1)

	'	objEncrypt.EncryptFile(strInFile, strOutFile, True, "230678" & strKey)
	'	'UPGRADE_NOTE: Object objEncrypt may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
	'	objEncrypt = Nothing

	'	OLEEncryptFile = strOutFile

	'End Function

	' Returns a message if the filename length is too long to embed/link

	Public Function CheckOLEFileNameLength(ByRef strFilename As String) As String

		Dim bOK As Boolean
		bOK = True

		' Defined maximum filename length of 50
		If Len(GetFileNameOnly(strFilename)) > 50 Then
			CheckOLEFileNameLength = "File name is too long." & vbCrLf & "Maximum file length is 50 characters."
			bOK = False
		End If

		' Defined maximum filename length of 50
		If Len(GetPathOnly(strFilename, True)) > 100 And bOK Then
			CheckOLEFileNameLength = "Directory structure is too long." & vbCrLf & "Maximum length is 100 characters."
		End If

		' Defined maximum filename length of 50
		If Len(Trim(GetUNCOnly(strFilename))) > 50 And bOK Then
			CheckOLEFileNameLength = "Network path is too long." & vbCrLf & "Maximum length is 50 characters."
		End If

	End Function

	' Is the passed in filename a valid picture extension
	Public Function CheckOLEPictureExtension(ByRef strFilename As String) As Boolean

		Dim strExtension As String

		If Len(strFilename) > 3 Then
			strExtension = UCase(Mid(strFilename, Len(strFilename) - 2, 3))

			CheckOLEPictureExtension = True

			If Not (strExtension = "JPG" Or strExtension = "GIF" Or strExtension = "BMP") Then
				CheckOLEPictureExtension = False
			End If
		Else
			CheckOLEPictureExtension = False
		End If

	End Function

	' Gets the size of a specified file
	Public Function FileSize(ByRef strFilename As String) As String

		On Error GoTo ErrorTrap
		Dim strSize As String

		mobjFileInfo = mobjFileSystem.GetFile(strFilename)
		'UPGRADE_WARNING: Couldn't resolve default property of object mobjFileInfo.Size. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		strSize = mobjFileInfo.Size

TidyUpAndExit:
		FileSize = strSize
		'UPGRADE_NOTE: Object mobjFileInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mobjFileInfo = Nothing
		Exit Function

ErrorTrap:
		strSize = "Unknown"
		GoTo TidyUpAndExit

	End Function

	Public Function FileLastModified(ByRef strFilename As String) As String

		On Error GoTo ErrorTrap
		Dim strDate As String

		mobjFileInfo = mobjFileSystem.GetFile(strFilename)
		strDate = CStr(mobjFileInfo.DateLastModified)

TidyUpAndExit:
		FileLastModified = strDate
		'UPGRADE_NOTE: Object mobjFileInfo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		mobjFileInfo = Nothing
		Exit Function

ErrorTrap:
		strDate = "Unknown"
		GoTo TidyUpAndExit

	End Function

	' Uploads a specified file to one of the ole paths
	Public Sub UploadFile(ByRef strFilename As String, ByRef strLocation As String)

		On Error GoTo ErrorTrap

		Dim strUploadPath As String

		' Upload the file
		mobjFileSystem.CopyFile(strFilename, strLocation & "\", True)

TidyUpAndExit:
		Exit Sub

ErrorTrap:
		GoTo TidyUpAndExit

	End Sub

	' Returns array of files in the ole path
	Public Function FolderList(ByRef pstrLocation As String) As Object

		On Error GoTo ErrorTrap

		FolderList = mobjFileSystem.GetFolder(pstrLocation)

TidyUpAndExit:
		Exit Function

ErrorTrap:
		GoTo TidyUpAndExit

	End Function

	Public Function NiceSize(ByRef pstrSize As String) As String

		Select Case Len(pstrSize)
			Case Is < 5
				NiceSize = pstrSize & " bytes"

			Case Is < 7
				NiceSize = Mid(pstrSize, 1, Len(pstrSize) - 3) & "KB"

			Case 7
				NiceSize = Mid(pstrSize, 1, 1) & "." & Mid(pstrSize, 2, 2) & "MB"

			Case Is < 10
				NiceSize = Mid(pstrSize, 1, Len(pstrSize) - 6) & "MB"

		End Select

	End Function

	Public Function GetOfficeWordVersion() As Short

		Dim App As Object

		On Error GoTo NotInstalled

		If miOfficeVersion_Word = 0 Then
			App = CreateObject("Word.Application")
			'UPGRADE_WARNING: Couldn't resolve default property of object App.Version. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			miOfficeVersion_Word = Val(App.Version)
			'UPGRADE_WARNING: Couldn't resolve default property of object App.Quit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			App.Quit()
		End If

TidyUpAndExit:
		GetOfficeWordVersion = miOfficeVersion_Word
		'UPGRADE_NOTE: Object App may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		App = Nothing

		Exit Function

NotInstalled:
		miOfficeVersion_Word = -1
		Resume TidyUpAndExit

	End Function

	Public Function GetOfficeExcelVersion() As Short

		Dim App As Object

		On Error GoTo NotInstalled

		If miOfficeVersion_Excel = 0 Then
			App = CreateObject("Excel.Application")
			'UPGRADE_WARNING: Couldn't resolve default property of object App.Version. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			miOfficeVersion_Excel = Val(App.Version)
			'UPGRADE_WARNING: Couldn't resolve default property of object App.Quit. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			App.Quit()
		End If

TidyUpAndExit:
		GetOfficeExcelVersion = miOfficeVersion_Excel
		'UPGRADE_NOTE: Object App may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		App = Nothing

		Exit Function

NotInstalled:
		miOfficeVersion_Excel = -1
		Resume TidyUpAndExit

	End Function


	''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''
	' Taken from modIntClient
	''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''

	Public Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, ByRef cbRemoteName As Integer) As Integer

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



End Class
