VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IObjectSafetyTLB.IObjectSafety

' General form handling variables.
'Private mobjPopup As prjIntranetPopup.clsIntranetPopup
Private msLocaleDateFormat As String
Private msLocaleDecimalSeparator As String
Private msLocaleThousandSeparator As String

'''Private mobjProc As Object

Dim mstrCurrentSessionKey As String

' Windows API function constants
Const LOCALE_USER_DEFAULT = &H400
Const LOCALE_SDATE = &H1D        '  date separator
Const LOCALE_SSHORTDATE = &H1F        '  short date format string
Const LOCALE_SDECIMAL = &HE         '  decimal separator
Const LOCALE_STHOUSAND = &HF         '  thousand separator

'''Const ERROR_SUCCESS = 0
'''Const HKEY_CLASSES_ROOT = &H80000000
'''Const KEY_ALL_ACCESS = &H3F
'''Const REG_SZ As Long = 1
'''Const REG_EXPAND_SZ As Long = 2

' Windows API functions
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'''Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'''Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
  (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
  ByVal samDesired As Long, phkResult As Long) As Long
'''Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName$, ByVal lpdwReserved As Long, lpdwType As Long, _
  lpData As Any, lpcbData As Long) As Long


Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Const PROCESS_QUERY_INFORMATION = &H400

Private MAPIMessages1 As MSMAPI.MAPIMessages
Private MAPISession1 As MSMAPI.MAPISession

'Security Settings - Login Checks i.e.Bad Attempts
Private gblnCFG_PCL As Boolean
Private gintCFG_BA As Integer
Private glngCFG_RT As Integer
Private glngCFG_LD As Long
Private gintPC_BA As Integer
Private gdtPC_LA As Date
Private gdtPC_LKD As Date

Private mobjFileSystem As New FileSystemObject
Private mobjFileInfo As File

'MM = Mail Merge
Private mstrMMStatusMessage As String
Private mlngMMMergeFieldsUbound As Long
Private maryMMOutputArrayData() As String
Private maryMMMergeFieldsData() As String
Private mstrMMDataSourceName As String
Private mblnMMCancelled As Boolean
Private mstrMMTempFileName As String

'Private mintMMDefOutput As Integer              'OutputFormat
'Private mstrMMDefDocFile As String              'OutputFilename
'Private mblnMMDefDocSave As Boolean             'OutputSave
'Private mblnMMDefCloseDoc As Boolean            'Not OutputScreen

Private mstrMMDefName As String
Private mstrMMDefTemplateFile As String
Private mstrMMDefTemplateSaveAs As String
Private mblnMMDefPauseBeforeMerge As Boolean
Private mblnMMDefSuppressBlankLines As Boolean

Private mstrMMDefEmailSubject As String
Private mlngMMDefEmailAddrCalc As Long
Private mblnMMDefEMailAttachment As Boolean
Private mstrMMDefAttachmentName As String
Private mstrMMDefAttachmentFormat As String

Private mintMMDefOutputFormat As Integer
Private mblnMMDefOutputScreen As Boolean
Private mblnMMDefOutputPrinter As Boolean
Private mstrMMDefOutputPrinterName As String
Private mblnMMDefOutputSave As Boolean
Private mstrMMDefOutputFileName As String
Private mstrMMDefOutputSaveAs As String

Private mlngDocManMapID As Long
Private mblnDocManManualHeader As Boolean

Private miOfficeVersion_Word As Integer
Private miOfficeVersion_Excel As Integer
Private mstrSaveAsValues As String

'Word Objects
Dim mwrdApp As Word.Application
Dim mdocDataSource As Word.Document
Dim mdocTemplate As Word.Document
Dim mdocOutput As Word.Document

Public Function CheckString(psString As String) As String
  ' Create a checksum of the given string
  Dim iKeyIndex As Integer
  Dim sCheckString As String
  Dim sngRandom As Single
  Dim lngIndex As Long
  Dim iCharCode As Integer
  Dim lngTotalChar As Long
  Dim aiKeys() As Integer
  Dim fModify As Boolean
  Dim sTemp As String

  Const KEYCOUNT = 10
  Const PREFIXKEYCOUNT = 5

  ReDim aiKeys(KEYCOUNT)

  lngTotalChar = 0
  iKeyIndex = 1
  sCheckString = ""

  ' Calculate the key digits.
  For lngIndex = 1 To UBound(aiKeys)
    Randomize
    sngRandom = (Rnd * 10) - 0.5
    aiKeys(lngIndex) = CInt(sngRandom)
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

Public Function MM_AddToOutputArrayData(pstrValue As String)
  Dim iNewIndex As Integer
  iNewIndex = UBound(maryMMOutputArrayData) + 1
  ReDim Preserve maryMMOutputArrayData(iNewIndex) As String
  maryMMOutputArrayData(iNewIndex) = pstrValue
End Function

Public Function MM_AddToMergeFieldsData(pstrValue As String)
  Dim iNewIndex As Integer
  iNewIndex = UBound(maryMMMergeFieldsData) + 1
  ReDim Preserve maryMMMergeFieldsData(iNewIndex) As String
  maryMMMergeFieldsData(iNewIndex) = pstrValue
End Function
Public Function MM_DimArrays()
  ReDim maryMMOutputArrayData(0) As String
  ReDim maryMMMergeFieldsData(0) As String
End Function


Public Function MM_DefCloseDoc(pblnValue As Boolean)
  'mblnMMDefCloseDoc = pblnValue
End Function

Public Property Let MM_DefDocFile(pstrValue As String)
  'mstrMMDefDocFile = pstrValue
End Property

Public Function MM_DefDocSave(pblnValue As Boolean)
  'mblnMMDefDocSave = pblnValue
End Function

Public Property Let MM_DefOutput(pintValue As Integer)
  'mintMMDefOutput = pintValue
End Property


Public Property Get MM_StatusMessage() As String
  MM_StatusMessage = mstrMMStatusMessage
End Property

Public Function MM_Cancelled() As Boolean
  MM_Cancelled = mblnMMCancelled
End Function

Public Property Let MM_MergeFieldsUbound(plngValue As Long)
  mlngMMMergeFieldsUbound = plngValue
End Property


Public Property Let MM_DefName(pstrValue As String)
  mstrMMDefName = pstrValue
End Property

Public Property Let MM_DefTemplateFile(pstrValue As String)
  mstrMMDefTemplateFile = pstrValue
End Property

Public Function MM_DefPauseBeforeMerge(pblnValue As Variant)
  mblnMMDefPauseBeforeMerge = pblnValue
End Function

Public Function MM_DefSuppressBlankLines(pblnValue As Boolean)
  mblnMMDefSuppressBlankLines = pblnValue
End Function


Public Property Let MM_DefEmailSubject(pstrValue As String)
  mstrMMDefEmailSubject = pstrValue
End Property

Public Property Let MM_DefEmailAddrCalc(plngValue As Long)
  mlngMMDefEmailAddrCalc = plngValue
End Property

Public Function MM_DefEMailAttachment(pblnValue As Boolean)
  mblnMMDefEMailAttachment = pblnValue
End Function

Public Property Let MM_DefAttachmentName(pstrValue As String)
  mstrMMDefAttachmentName = pstrValue
End Property


Public Property Let MM_DefOutputFormat(pintValue As Integer)
  mintMMDefOutputFormat = pintValue
End Property

Public Function MM_DefOutputScreen(pblnValue As Boolean)
  mblnMMDefOutputScreen = pblnValue
End Function

Public Function MM_DefOutputPrinter(pblnValue As Boolean)
  mblnMMDefOutputPrinter = pblnValue
End Function

Public Property Let MM_DefOutputPrinterName(pstrValue As String)
  mstrMMDefOutputPrinterName = pstrValue
End Property

Public Function MM_DefOutputSave(pblnValue As Boolean)
  mblnMMDefOutputSave = pblnValue
End Function

Public Property Let MM_DefOutputFileName(pstrValue As String)
  mstrMMDefOutputFileName = pstrValue
End Property


Public Property Let MM_DefDocManMapID(pintValue As Integer)
  mlngDocManMapID = pintValue
End Property

Public Function MM_DefDocManManualHeader(pblnValue As Boolean)
  mblnDocManManualHeader = pblnValue
End Function


Public Property Let SaveAsValues(strValue As String)
  mstrSaveAsValues = strValue
End Property



Private Function MM_WORDDOC_OpenTempate(pwrdApp As Word.Application, pbSuppressPrompt As Boolean) As Boolean
  
  Dim strMSG As String
  'Dim wrdDocTemplate As Word.Document
  
  Set mdocTemplate = pwrdApp.Documents.Open(mstrMMDefTemplateFile, False, False, False, "", "", False, "", "", 0)
  
  If mdocTemplate.MailMerge.MainDocumentType <> wdNotAMergeDocument Then
    strMSG = "The Word template file <" & mstrMMDefTemplateFile & "> has been saved referencing a different set of data." & vbCrLf & vbCrLf & "You need to remove this reference in order to proceed with this mail merge." & vbCrLf & vbCrLf & "Proceed ?"
    If Not pbSuppressPrompt Then
      If MsgBox(strMSG, vbYesNo + vbQuestion, "Mail Merge") <> vbYes Then
        mblnMMCancelled = True
        MM_WORDDOC_OpenTempate = False
        Exit Function
      End If
    End If
  End If
  
  mdocTemplate.MailMerge.MainDocumentType = wdNotAMergeDocument
  
  pwrdApp.CustomizationContext = mdocTemplate
  
  On Error Resume Next
  mdocTemplate.CommandBars("Mail Merge").Controls("Mail Me&rge...").Visible = False
  mdocTemplate.CommandBars("Mail Merge").Controls("&Error Check Merge").Visible = False
  mdocTemplate.CommandBars("Mail Merge").Controls("&Merge to Document").Visible = False
  mdocTemplate.CommandBars("Mail Merge").Controls("&Merge to Printer").Visible = False
  mdocTemplate.CommandBars("Mail Merge").Controls("&Merge Options").Visible = False
  mdocTemplate.CommandBars("Mail Merge").Controls("&Find in Field").Visible = False
  mdocTemplate.CommandBars("Mail Merge").Controls("&Data Form").Visible = False
  
  On Error GoTo ErrorTrap
  mdocTemplate.Saved = True
 
  MM_WORDDOC_OpenTempate = True
  
  Exit Function
  
ErrorTrap:
  'mstrMMStatusMessage = "Error opening Word template file <" & mstrMMDefTemplateFile & ">."
  mstrMMStatusMessage = "Error opening Word template file <" & Replace(mstrMMDefTemplateFile, "\", "\\") & ">."
  MM_WORDDOC_OpenTempate = False
  
End Function

Public Function MM_WORD_CreateTemplateFile(psTemplatePath As String) As Boolean

  On Error GoTo ErrorTrap
  
  ShowPopup ("Creating Word Document...")
  
  Set mwrdApp = New Word.Application
  mwrdApp.DisplayAlerts = wdAlertsNone
  Set mdocTemplate = mwrdApp.Documents.Add


  mstrMMDefTemplateSaveAs = GetSaveAsFormat2(psTemplatePath, mstrSaveAsValues)
  
  If mstrMMDefTemplateSaveAs <> vbNullString Then
    mdocTemplate.SaveAs psTemplatePath, Val(mstrMMDefTemplateSaveAs)
  Else
    MsgBox "Error creating template file", vbCritical + vbOKOnly, "OpenHR Intranet"
    MM_WORD_CreateTemplateFile = False
  End If
  
  
  On Error Resume Next
  mdocTemplate.Close (False)
  mwrdApp.Quit (False)

  MM_WORD_CreateTemplateFile = True

TidyUpAndExit:
  ClosePopup
  Set mdocTemplate = Nothing
  Set mwrdApp = Nothing
  Exit Function
  
ErrorTrap:
  MsgBox "Error creating template file", vbCritical + vbOKOnly, "OpenHR Intranet"
  MM_WORD_CreateTemplateFile = False
  GoTo TidyUpAndExit
  
End Function

Private Function MM_WORDDOC_ValidateTemplateMergeFields(pdocTemplate As Word.Document, pblnAbortOnError As Boolean) As Boolean
  
  Dim fOk As Boolean
  Dim intCount As Integer
  Dim strField As String
  Dim blnFieldOkay As Boolean
  Dim strTemp As String
  Dim intResponse As VbMsgBoxResult
  Dim wrdApp As Word.Application
  
  On Error GoTo ErrorTrap
  
  fOk = True
  blnFieldOkay = False
  mstrMMStatusMessage = vbNullString
  
  If fOk Then
    intCount = 1
    While intCount <= pdocTemplate.Fields.Count
      strField = pdocTemplate.Fields(intCount).Code
      If Mid(strField, 2, 10) = "MERGEFIELD" Then
        strField = Mid(strField, 13)
        strField = RTrim(strField)
        strField = Replace(strField, """", "")
        strField = Replace(strField, " ", "_")
        strField = Left(strField, 39)
      
        blnFieldOkay = MM_FieldOK(strField)
              
        If Not blnFieldOkay Then
          mstrMMStatusMessage = "The Template Document '" & mstrMMDefTemplateFile & "' contains the merge field " & strField & " which was not selected in the mail merge definition."
          If pblnAbortOnError Then
            MM_WORDDOC_ValidateTemplateMergeFields = False
            Exit Function
          End If
          
          If pblnAbortOnError Then
            MM_WORDDOC_ValidateTemplateMergeFields = False
            Exit Function
          End If
          
          strTemp = mstrMMStatusMessage & vbCrLf & vbCrLf & "Select OK to proceed with the merge and remove this " & _
                                                          "field from the template or select Cancel to abort this Mail Merge."
          intResponse = MsgBox(strTemp, vbOKCancel + vbQuestion, "Mail Merge")
          If intResponse = vbOK Then
            pdocTemplate.Fields(intCount).Delete
            pdocTemplate.Saved = False
            mstrMMStatusMessage = vbNullString
            intCount = intCount - 1
          Else
            mstrMMStatusMessage = "The Template Document '" & mstrMMDefTemplateFile & "' contains the merge field " & strField & " which was not selected in the mail merge definition."
            MM_WORDDOC_ValidateTemplateMergeFields = False
            mblnMMCancelled = True
            Exit Function
          End If
        
        End If
      End If
    
      intCount = intCount + 1
    Wend
   
  End If

  MM_WORDDOC_ValidateTemplateMergeFields = True

  Exit Function
  
ErrorTrap:
  Select Case Err.Number
    Case 462, -2147417848
      If Err.Number = 462 Then
        Set mwrdApp = New Word.Application
      End If
      fOk = MM_WORDDOC_OpenTempate(mwrdApp, True)
      If fOk Then
        Set pdocTemplate = mwrdApp.ActiveDocument
        MM_WORDDOC_SetMergeOptions pdocTemplate, mstrMMDataSourceName
        pdocTemplate.Saved = True
      End If
  
    Case Else
      fOk = False
      mstrMMStatusMessage = "Error Validating Merge Fields (" & Replace(Err.Description, "\", "\\") & ")"
      MM_WORDDOC_ValidateTemplateMergeFields = False
  End Select
  
End Function

  
  
Private Function MM_FieldOK(psFieldName As String) As Boolean
  
  Dim i As Integer
  
  For i = 0 To UBound(maryMMMergeFieldsData)
    maryMMMergeFieldsData(i) = Trim(maryMMMergeFieldsData(i))
    maryMMMergeFieldsData(i) = Replace(maryMMMergeFieldsData(i), Chr(34), "")
    maryMMMergeFieldsData(i) = Replace(maryMMMergeFieldsData(i), " ", "_")
    maryMMMergeFieldsData(i) = Left(maryMMMergeFieldsData(i), 39)
    
    If mlngMMDefEmailAddrCalc > 0 Then
      If UCase(psFieldName) = UCase(maryMMMergeFieldsData(i)) Or _
        UCase(psFieldName) = "EMAIL_ADDRESS" Then
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

Private Function MM_WORDDOC_SetMergeOptions(pdocTemplate As Word.Document, pstrDataSource As String) As Boolean
  
  Dim objTemp As Object
  
  pdocTemplate.MailMerge.OpenDataSource mstrMMDataSourceName, , , , , False
  pdocTemplate.MailMerge.SuppressBlankLines = mblnMMDefSuppressBlankLines
  
  If mintMMDefOutputFormat = 1 Then
    pdocTemplate.MailMerge.Destination = wdSendToEmail
    pdocTemplate.MailMerge.MailAddressFieldName = "Email_Address"
    pdocTemplate.MailMerge.MailSubject = mstrMMDefEmailSubject
    pdocTemplate.MailMerge.MailAsAttachment = LCase(mblnMMDefEMailAttachment)
    
    If Val(mwrdApp.Version) >= 12 Then
      pdocTemplate.MailMerge.DataSource.FirstRecord = wdDefaultFirstRecord
      pdocTemplate.MailMerge.DataSource.LastRecord = wdDefaultLastRecord
      
      If Not mblnMMDefEMailAttachment Then
        Set objTemp = pdocTemplate.MailMerge
        objTemp.MailFormat = 1   'wdMailFormatHTML
      End If
    End If

  Else
    pdocTemplate.MailMerge.Destination = wdSendToNewDocument
  End If

  MM_WORDDOC_SetMergeOptions = True
  
  Exit Function
  
ErrorTrap:
  mstrMMStatusMessage = "Error setting merge options (" & Replace(Err.Description, "\", "\\") & ")"
  MM_WORDDOC_SetMergeOptions = False
  
End Function

Public Function MM_WORD_ExecuteMailMerge() As Boolean

  On Error GoTo ErrorTrap
  
  Dim fOk As Boolean
  fOk = True
  
  mblnMMCancelled = False
  
  'Open Word application
  mstrMMStatusMessage = "Error opening Word application."
  Set mwrdApp = New Word.Application
  mwrdApp.DisplayAlerts = wdAlertsNone
  
  'Check Word version
  If Val(mwrdApp.Version) < 8 Then
    mstrMMStatusMessage = "You are running a version of Word which is not supported by HR-Pro."
    fOk = False
    GoTo TidyUpAndExit
  End If
  
  
  'Check Template File
'  If Not IsFileCompatibleWithWordVersion(mstrMMDefTemplateFile, Val(mwrdApp.Version)) Then
'    mstrMMStatusMessage = "This definition is set to use a template file which is not compatible with your version of Microsoft Office."
'    fOk = False
'    GoTo TidyUpAndExit
'  End If
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
    mstrMMStatusMessage = "This definition is set to output to printer " & mstrMMDefOutputPrinterName & _
                          " which is not set up on your PC."
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
  
  If Not MM_WORD_Execute_Step2() Then
    fOk = False
    GoTo ErrorTrap
  End If
  
  If Not MM_WORD_Execute_Step3() Then
    fOk = False
    GoTo ErrorTrap
  End If
  
  If Not MM_WORD_Execute_Step4() Then
    fOk = False
    GoTo ErrorTrap
  End If
  

TidyUpAndExit:
  Set mdocDataSource = Nothing
  Set mdocTemplate = Nothing
  MM_WORD_ExecuteMailMerge = fOk
  Exit Function

ErrorTrap:
  fOk = False
  'Kill Word Documents
  Set mdocDataSource = Nothing
  Set mdocTemplate = Nothing
  'Kill Word Application in correct order.
  If Not mwrdApp Is Nothing Then mwrdApp.Quit False
  Set mwrdApp = Nothing
  GoTo TidyUpAndExit
  
End Function

Private Function ValidPrinter(strName As String) As Boolean

  Dim objPrinter As Printer
  Dim blnFound As Boolean

  If strName <> vbNullString And strName <> "<Default Printer>" Then
    blnFound = False
    For Each objPrinter In Printers
      If objPrinter.DeviceName = strName Then
        blnFound = True
        Exit For
      End If
    Next
  Else
    blnFound = True
  End If

  ValidPrinter = blnFound

End Function

Private Function MM_WORD_Execute_Step1() As Boolean

  On Error GoTo ErrorTrap

  Dim fOk As Boolean
  Dim strDataSourceFormat As String
  
  fOk = True

  'Microsoft Word Declarations
  Dim intCount As Integer
  Dim lngTableCols As Long
  Dim intMBResponse As VbMsgBoxResult
  Dim blnMergeFieldExists As Boolean
  Dim strField As String
  
  
'  mstrMMDefTemplateSaveAs = GetSaveAsFormat2(mstrMMDefTemplateFile, mstrSaveAsValues)
'  If mstrMMDefTemplateSaveAs = vbNullString Then
'    fOk = False
'    mstrMMStatusMessage = "This definition is set to use a template file which is not compatible with your version of Microsoft Office."
'    GoTo TidyUpAndExit
'  End If
  
  
'  If mblnMMDefOutputSave Then
'    mstrMMDefOutputSaveAs = GetSaveAsFormat2(mstrMMDefOutputFileName, mstrSaveAsValues)
'    If mstrMMDefOutputSaveAs = vbNullString Then
'      fOk = False
'      mstrMMStatusMessage = "This definition is set to save in a file format which is not compatible with your version of Microsoft Office."
'      GoTo TidyUpAndExit
'    End If
'  End If
  
  
  
  'Create Data Source
  mstrMMStatusMessage = "Error creating data source."
  mstrMMDataSourceName = GetTempFile()
  Set mdocDataSource = mwrdApp.Documents.Add()


  strDataSourceFormat = GetSaveAsFormat2(mstrMMDataSourceName, mstrSaveAsValues)
  If strDataSourceFormat = "" Then strDataSourceFormat = "0"
  mdocDataSource.SaveAs mstrMMDataSourceName, Val(strDataSourceFormat)

  'Populate Data Source
  mstrMMStatusMessage = "Error populating data source."
  If mlngMMMergeFieldsUbound < 3 Then
    mdocDataSource.ActiveWindow.Selection.Bookmarks.Add "TableStart", mdocDataSource.ActiveWindow.Selection.Range
  End If
  
  For intCount = 1 To UBound(maryMMOutputArrayData) Step 1
    If intCount > 1 Then
      mdocDataSource.ActiveWindow.Selection.TypeParagraph
    End If
    mdocDataSource.ActiveWindow.Selection.TypeText (maryMMOutputArrayData(intCount))
  Next intCount
  
  lngTableCols = mlngMMMergeFieldsUbound
  If mintMMDefOutputFormat = 1 Then   'Email
    lngTableCols = lngTableCols + 1
  End If
  If lngTableCols < 3 Then
    mdocDataSource.ActiveWindow.Selection.Goto wdGoToBookmark, , , "TableStart"
    mdocDataSource.ActiveWindow.Selection.EndKey wdStory, wdExtend
    mdocDataSource.ActiveWindow.Selection.ConvertToTable wdSeparateByTabs, , , , wdTableFormatNone, _
                                                         , , False, False, , , , _
                                                         , False
  End If
  mdocDataSource.SaveAs FileFormat:=Val(strDataSourceFormat)
  mdocDataSource.Close (False)
  
  'Call open template
  fOk = MM_WORDDOC_OpenTempate(mwrdApp, False)
  If fOk Then
    Set mdocTemplate = mwrdApp.ActiveDocument
  End If
  
  If fOk Then
    If mdocTemplate.ReadOnly Then
      fOk = MM_WORDDOC_ValidateTemplateMergeFields(mdocTemplate, True)
      If fOk Then
        MM_WORDDOC_SetMergeOptions mdocTemplate, mstrMMDataSourceName
      End If
    Else
      fOk = MM_WORDDOC_ValidateTemplateMergeFields(mdocTemplate, False)
      If fOk Then
        MM_WORDDOC_SetMergeOptions mdocTemplate, mstrMMDataSourceName
      End If
    End If
  End If
  mdocTemplate.Saved = True
  
  'Pause before merge
  intMBResponse = vbNo
  If mblnMMDefPauseBeforeMerge Then
    If fOk Then
      intMBResponse = MsgBox("Would you like to amend the document or merge fields?", vbYesNoCancel + vbQuestion, "Mail Merge")
      If intMBResponse = vbCancel Then
        mblnMMCancelled = True
        fOk = False
      End If
    End If
  End If


CheckMergeFieldExists:

  'Check that there is at least one merge field
  If fOk Then
    If intMBResponse <> vbYes Then
      intCount = 1
      blnMergeFieldExists = False
      
      Do While blnMergeFieldExists = False And mblnMMCancelled = False

        While ((intCount <= mdocTemplate.Fields.Count) And (Not blnMergeFieldExists))
          strField = mdocTemplate.Fields(intCount).Code
          If Mid(strField, 2, 10) = "MERGEFIELD" Then
            blnMergeFieldExists = True
          End If
          intCount = intCount + 1
        Wend

        If Not blnMergeFieldExists Then
          intMBResponse = MsgBox("There are no merge fields specified in the template document." & vbCrLf & vbCrLf & _
                                  "Would you like to switch to word to insert merge fields now?", vbQuestion + vbOKCancel, "Mail Merge")
          If intMBResponse = vbCancel Then
            mblnMMCancelled = True
            fOk = False
          Else
            mwrdApp.Visible = True
            mwrdApp.WindowState = wdWindowStateNormal
            mwrdApp.Activate

            intMBResponse = MsgBox("Please insert merge fields on the template document then click OK to proceed", _
                                    vbInformation + vbOKCancel, "Mail Merge")

            If intMBResponse = vbCancel Then
              mblnMMCancelled = True
              fOk = False
            End If
          End If
        End If
      Loop

    End If
      
    If intMBResponse = vbYes Then
      mwrdApp.Visible = True
      mwrdApp.WindowState = wdWindowStateNormal
      mwrdApp.Activate
      
      intMBResponse = MsgBox("Click OK to proceed with the mail merge.", vbOKCancel + vbInformation, "Mail Merge")
      
      If intMBResponse = vbCancel Then
        mblnMMCancelled = True
        fOk = False
      End If
      
      MM_WORDDOC_ValidateTemplateMergeFields mdocTemplate, False
    End If
  End If

TidyUpAndExit:
  MM_WORD_Execute_Step1 = fOk
  Exit Function

ErrorTrap:
  
  Select Case Err.Number
  Case 462, -2147417848
    If Err.Number = 462 Then
      Set mwrdApp = New Word.Application
    End If
    fOk = MM_WORDDOC_OpenTempate(mwrdApp, True)
    If fOk Then
      Set mdocTemplate = mwrdApp.ActiveDocument
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

Private Function MM_WORD_Execute_Step2() As Boolean

  On Error GoTo ErrorTrap
  
  Dim fOk As Boolean
  Dim blnTemplateReadOnly As Boolean
  Dim intMBResponse As VbMsgBoxResult
  Dim intErrTrapCode As Integer
  
  fOk = True
  
  intErrTrapCode = 0
  
  mwrdApp.DisplayAlerts = wdAlertsAll
  blnTemplateReadOnly = False
  intMBResponse = vbYes
  While ((mdocTemplate.Saved = False) And (intMBResponse = vbYes))
    
    intMBResponse = MsgBox("You have not saved changes to the template document, " & _
                            "Would you like to save changes now?", vbQuestion + vbYesNo, "Mail Merge")
    intErrTrapCode = 1
    If intMBResponse = vbYes Then
      mwrdApp.Visible = True
      mwrdApp.DisplayAlerts = wdAlertsAll
      mdocTemplate.SaveAs FileFormat:=Val(mstrMMDefTemplateSaveAs)
    End If
    
    If blnTemplateReadOnly = True Then
      intMBResponse = MsgBox("This template file is currently in use by another user." & _
                              "Please save as a new file.", vbInformation + vbOKOnly, "Mail Merge")
      intErrTrapCode = 2
      mdocTemplate.SaveAs FileFormat:=Val(mstrMMDefTemplateSaveAs)
      blnTemplateReadOnly = False
      intMBResponse = vbYes
    End If
  
    mwrdApp.Visible = False
    mwrdApp.DisplayAlerts = wdAlertsNone
    
  Wend

  
TidyUpAndExit:
  MM_WORD_Execute_Step2 = fOk
  Exit Function
  
ErrorTrap:
  Select Case Err.Number
  Case 462, -2147417848
    If Err.Number = 462 Then
      Set mwrdApp = New Word.Application
    End If
    fOk = MM_WORDDOC_OpenTempate(mwrdApp, True)
    If fOk Then
      Set mdocTemplate = mwrdApp.ActiveDocument
      MM_WORDDOC_SetMergeOptions mdocTemplate, mstrMMDataSourceName
      mdocTemplate.Saved = True
      Resume Next
    End If
  Case 4198
  Case 5155
  Case 5356
    If intErrTrapCode = 1 Then
      blnTemplateReadOnly = True
      Resume Next
    End If
  Case Else
    fOk = False
    mstrMMStatusMessage = "Error checking if the template file has been saved (" & intErrTrapCode & ")."
  End Select
  
  GoTo TidyUpAndExit
  
End Function

Private Function MM_WORD_Execute_Step3() As Boolean

  Dim strField As String
  Dim blnMergeFieldExists As Boolean
  Dim intCount As Integer
  Dim strOriginalDefaultPrinter As String
  
  'Run the Merge.
  
  On Error GoTo ErrorTrap
  
  Dim fOk As Boolean
  
  fOk = True
  
  If fOk Then

    'MH20050323 Fault 9909
    intCount = 1
    blnMergeFieldExists = False
    While ((intCount <= mdocTemplate.Fields.Count) And (Not blnMergeFieldExists))
      strField = mdocTemplate.Fields(intCount).Code
      If Mid(strField, 2, 10) = "MERGEFIELD" Then
        blnMergeFieldExists = True
      End If
      intCount = intCount + 1
    Wend

    If Not blnMergeFieldExists Then
      mstrMMStatusMessage = "No merge fields specified in the template document."
      fOk = False
      Exit Function
    End If



    mwrdApp.DisplayAlerts = wdAlertsAll
    
    'Rename attachment
    'Rename attachment
    If Trim(mstrMMDefAttachmentName) <> vbNullString Then
      
      'Get temp path
      mstrMMTempFileName = Space(1024)
      mstrMMTempFileName = GetTempFile()

      If Dir(mstrMMDefAttachmentName) <> vbNullString Then
        Kill (mstrMMDefAttachmentName)
      End If
      
      mdocTemplate.SaveAs mstrMMDefAttachmentName, Val(mstrMMDefAttachmentFormat)
  
    End If
  
    mdocTemplate.MailMerge.Execute False
    Set mdocOutput = mwrdApp.ActiveDocument
    mdocOutput.ActiveWindow.View.Type = wdNormalView
    
    
    If mblnMMDefOutputSave Then
      If Dir(mstrMMDefOutputFileName) <> vbNullString Then
        Kill (mstrMMDefOutputFileName)
      End If
      mdocOutput.SaveAs mstrMMDefOutputFileName, Val(mstrMMDefOutputSaveAs)
      While (mwrdApp.BackgroundSavingStatus > 0)
        DoEvents
      Wend
    End If


    If mblnMMDefOutputPrinter Then
      strOriginalDefaultPrinter = vbNullString
      If Trim(mstrMMDefOutputPrinterName) <> vbNullString Then
        strOriginalDefaultPrinter = mwrdApp.ActivePrinter
        mwrdApp.ActivePrinter = mstrMMDefOutputPrinterName
      End If

      mdocOutput.PrintOut
      While (mwrdApp.BackgroundPrintingStatus > 0)
        DoEvents
      Wend
    
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
      Set mwrdApp = New Word.Application
      
    Case -2147417848
      mstrMMStatusMessage = "Microsoft Word document has been closed by the user."
    
    Case 4605
      mstrMMStatusMessage = "No merge fields specified in the template document."
    
    Case 5152
      mstrMMStatusMessage = "Error saving Mail Merge to file" & vbCrLf & _
                            mstrMMDefOutputFileName & vbCrLf & _
                            "Please ensure that the output document path is correct within this definition."
    
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
  
  mwrdApp.DisplayAlerts = wdAlertsNone
    
  'Make sure that the template file does not still reference the data source
  '(close it and open it incase there have been changes to it or that the
  'the template is now the temporary email attachment name thingy!
  On Error Resume Next
  mdocTemplate.Close False
  If mstrMMDefTemplateSaveAs <> vbNullString Then
    Set mdocTemplate = mwrdApp.Documents.Open(mstrMMDefTemplateFile, False, False, _
                                              False, vbNullString, vbNullString, False, vbNullString, vbNullString, 0)
    mdocTemplate.MailMerge.MainDocumentType = wdNotAMergeDocument
    mdocTemplate.Saved = False    'MH20070227 Fault 12013 - Force save for Office 2007
    mdocTemplate.SaveAs mstrMMDefTemplateFile, Val(mstrMMDefTemplateSaveAs)
    mdocTemplate.Close (False)
  End If
  
  'If the temorary email attachment name exists then kill it
  If mstrMMDefAttachmentName <> vbNullString Then
    Kill (mstrMMTempFileName)
  End If
  
  mdocDataSource.Close False
  
  On Error GoTo ErrorTrap
  'Kill the data source
  If Dir(mstrMMDataSourceName) <> vbNullString Then
    Kill (mstrMMDataSourceName)
  End If
  
  'blnLeaveDocOpen = ((mblnMMDefCloseDoc = False) And (mintMMDefOutput = OutputDestinations.desScreen))
  blnLeaveDocOpen = (mblnMMDefOutputScreen And fOk = True)
  
  'MsgBox fOk
  'MsgBox mblnMMDefOutputScreen
  'MsgBox blnLeaveDocOpen
  
  On Error Resume Next
  If blnLeaveDocOpen Then
    mwrdApp.Visible = True
    mwrdApp.WindowState = wdWindowStateNormal
    mwrdApp.Activate
    mdocOutput.Activate
  Else
    mdocOutput.Close False
    Set mdocOutput = Nothing
    Set mdocTemplate = Nothing
    mwrdApp.Quit
    Set mwrdApp = Nothing
  End If
  
TidyUpAndExit:
  MM_WORD_Execute_Step4 = fOk
  Exit Function
  
ErrorTrap:
  MsgBox Err.Description
  fOk = False
  mwrdApp.Quit
  Set mwrdApp = Nothing
  GoTo TidyUpAndExit

End Function

Private Sub Class_Terminate()
  Set mdocDataSource = Nothing
  Set mdocOutput = Nothing
  Set mdocTemplate = Nothing
  Set mwrdApp = Nothing
End Sub

Private Sub IObjectSafety_GetInterfaceSafetyOptions(ByVal riid As Long, _
                                                    pdwSupportedOptions As Long, _
                                                    pdwEnabledOptions As Long)

    Dim Rc      As Long
    Dim rClsId  As udtGUID
    Dim IID     As String
    Dim bIID()  As Byte

    pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or _
                          INTERFACESAFE_FOR_UNTRUSTED_DATA

    If (riid <> 0) Then
        CopyMemory rClsId, ByVal riid, Len(rClsId)

        bIID = String$(MAX_GUIDLEN, 0)
        Rc = StringFromGUID2(rClsId, VarPtr(bIID(0)), MAX_GUIDLEN)
        Rc = InStr(1, bIID, vbNullChar) - 1
        IID = Left$(UCase(bIID), Rc)

        Select Case IID
            Case IID_IDispatch
                pdwEnabledOptions = IIf(m_fSafeForScripting, _
              INTERFACESAFE_FOR_UNTRUSTED_CALLER, 0)
                Exit Sub
            Case IID_IPersistStorage, IID_IPersistStream, _
               IID_IPersistPropertyBag
                pdwEnabledOptions = IIf(m_fSafeForInitializing, _
              INTERFACESAFE_FOR_UNTRUSTED_DATA, 0)
                Exit Sub
            Case Else
                Err.Raise E_NOINTERFACE
                Exit Sub
        End Select
    End If

End Sub


Private Sub IObjectSafety_SetInterfaceSafetyOptions(ByVal riid As Long, _
                                                    ByVal dwOptionsSetMask As Long, _
                                                    ByVal dwEnabledOptions As Long)
    Dim Rc          As Long
    Dim rClsId      As udtGUID
    Dim IID         As String
    Dim bIID()      As Byte

    If (riid <> 0) Then
        CopyMemory rClsId, ByVal riid, Len(rClsId)

        bIID = String$(MAX_GUIDLEN, 0)
        Rc = StringFromGUID2(rClsId, VarPtr(bIID(0)), MAX_GUIDLEN)
        Rc = InStr(1, bIID, vbNullChar) - 1
        IID = Left$(UCase(bIID), Rc)

        Select Case IID
            Case IID_IDispatch
                If ((dwEnabledOptions And dwOptionsSetMask) <> _
             INTERFACESAFE_FOR_UNTRUSTED_CALLER) Then
                    Err.Raise E_FAIL
                    Exit Sub
                Else
                    If Not m_fSafeForScripting Then
                        Err.Raise E_FAIL
                    End If
                    Exit Sub
                End If

            Case IID_IPersistStorage, IID_IPersistStream, _
          IID_IPersistPropertyBag
                If ((dwEnabledOptions And dwOptionsSetMask) <> _
              INTERFACESAFE_FOR_UNTRUSTED_DATA) Then
                    Err.Raise E_FAIL
                    Exit Sub
                Else
                    If Not m_fSafeForInitializing Then
                        Err.Raise E_FAIL
                    End If
                    Exit Sub
                End If

            Case Else
                Err.Raise E_NOINTERFACE
                Exit Sub
        End Select
    End If

End Sub

Public Function SelectFolder(psInitDir As String, psPrompt As String, psTitle As String) As String
  
  Dim fOk As Boolean
  Dim frmDir As New frmPathSel
  
  frmDir.Caption = psTitle
  If ValidateDir(psInitDir) Then
    frmDir.lblPrompt.Caption = "Please select a folder for the '" & psTitle & "'"
    
  Else
    frmDir.lblPrompt = "The '" & psTitle & "' has either not been set, or does not exist. " & _
                       "Please select a folder for the '" & psTitle & "'."
    psInitDir = vbNullString
  End If
  
  fOk = frmDir.Initialise(psInitDir)
  
  If fOk Then
    frmDir.Show vbModal
    SelectFolder = frmDir.SelectedFolder
  End If
  
TidyUpAndExit:
  Set frmDir = Nothing
  Exit Function
  
ErrorTrap:
  SelectFolder = vbNullString
  GoTo TidyUpAndExit
  
End Function

Public Property Let DatabaseName(ByVal strNewValue As String)
  gsDatabaseName = strNewValue
End Property

Public Property Let ServerName(ByVal strNewValue As String)
  gsServerName = strNewValue
End Property

Public Function ClientMonthName(plngMonth As Long, pbAbbreviate As Boolean) As String
  ClientMonthName = MonthName(plngMonth, pbAbbreviate)
End Function

Public Function ClientWeekDayName(plngDay As Long, pbAbbreviate As Boolean, FirstDayofWeek As VbDayOfWeek) As String
  ClientWeekDayName = WeekdayName(plngDay, pbAbbreviate, FirstDayofWeek)
End Function

Public Function ClientWeekDayLetter(plngDay As Long, FirstDayofWeek As VbDayOfWeek) As String
  ClientWeekDayLetter = Left(WeekdayName(plngDay, True, vbSunday), 1)
End Function
' The current session key (used for encryption purposes)
Public Property Let CurrentSessionKey(pstrNewValue As String)
  mstrCurrentSessionKey = pstrNewValue
End Property

Private Function LC_FormatDateTimeMessage(plngSeconds As Long) As String

  Dim strTemp As String
  Dim intTemp As Integer
  Dim intFix As Integer
  
  intTemp = plngSeconds
  strTemp = vbNullString
  
  'Calculate days
  If intTemp > 86400 Then
    intFix = Fix(intTemp / 86400)
    intTemp = (intTemp Mod 86400)
    strTemp = strTemp & CStr(intFix) & IIf((intFix > 1), " days", " day")
  End If

  'Calculate months
  If intTemp > 3600 Then
    intFix = Fix(intTemp / 3600)
    intTemp = (intTemp Mod 3600)
    strTemp = strTemp & IIf((Len(strTemp) > 0), " ", "")
    strTemp = strTemp & CStr(intFix) & IIf((intFix > 1), " hours", " hour")
  End If
 
  'Calculate minutes
  If intTemp > 60 Then
    intFix = Fix(intTemp / 60)
    intTemp = (intTemp Mod 60)
    strTemp = strTemp & IIf((Len(strTemp) > 0), " ", "")
    strTemp = strTemp & CStr(intFix) & IIf((intFix > 1), " minutes", " minute")
  End If

  'Calculate seconds
  If intTemp > 0 Then
    strTemp = strTemp & IIf((Len(strTemp) > 0), " ", "")
    strTemp = strTemp & CStr(intTemp) & IIf((intTemp > 1), " seconds", " second")
  Else
    If (Len(strTemp) = 0) Then
      strTemp = strTemp & "0 seconds"
    End If
  End If

  LC_FormatDateTimeMessage = strTemp
  
End Function

Public Function LC_IncrementBadAttempt() As Boolean

  Dim intBadAttempts As Integer

  LC_Initialise
  
  'Exit the Login Check functions if the PC Lockout functionality has been disabled.
  If Not gblnCFG_PCL Then
    LC_IncrementBadAttempt = True
    Exit Function
  End If
  
  LC_ReadCurrentLockStatus
  
  intBadAttempts = gintPC_BA + 1
  
  If (intBadAttempts >= gintCFG_BA) Then
    LC_SaveCurrentLockStatus CStr(intBadAttempts), CStr(Format(Now(), "yyyy/mm/dd hh:mm:ss")), CStr(Format(Now(), "yyyy/mm/dd hh:mm:ss"))
    LC_PCLocked
  Else
    LC_SaveCurrentLockStatus CStr(intBadAttempts), CStr(Format(Now(), "yyyy/mm/dd hh:mm:ss")), "0"
  End If
 
  'Re-initialise PC variables as these are always refreshed by "LC_ReadCurrentLockStatus"
  gintPC_BA = 0
  gdtPC_LA = 0
  gdtPC_LKD = 0
 
End Function

Private Function LC_Initialise() As Boolean
 
  Dim strCFG_PCL As String
  Dim strCFG_BA As String
  Dim strCFG_RT As String
  Dim strCFG_LD As String
 
  Dim strInput As String
  Dim strEKey As String
  Dim strLens As String
  Dim lngStart As Long
  Dim lngFinish As Long
  
  Dim strOutput As String
  Dim lngCount As Long

  On Error GoTo ErrorTrap
  
  strInput = GetPCSetting("Misc", "ASRSysParam1", vbNullString)
  If strInput = vbNullString Then
    strCFG_PCL = "1"
    strCFG_BA = "3"
    strCFG_LD = "300"
  Else
    lngStart = Len(strInput) - 13
    strEKey = Mid(strInput, lngStart + 1, 10)
    strLens = Right(strInput, 3)
    strInput = XOREncript(Left(strInput, lngStart), strEKey)
  
    lngStart = 1
    lngFinish = Asc(Mid(strLens, 1, 1)) - 127
    strCFG_PCL = Mid(strInput, lngStart, lngFinish)
    
    lngStart = lngStart + lngFinish
    lngFinish = Asc(Mid(strLens, 2, 1)) - 127
    strCFG_BA = Mid(strInput, lngStart, lngFinish)
    
    lngStart = lngStart + lngFinish
    lngFinish = Asc(Mid(strLens, 3, 1)) - 127
    strCFG_LD = Mid(strInput, lngStart, lngFinish)
  End If

  gblnCFG_PCL = CBool(strCFG_PCL)
  gintCFG_BA = CInt(strCFG_BA)
  glngCFG_LD = CLng(strCFG_LD)
  
  
  strInput = GetPCSetting("Misc", "ASRSysParam3", vbNullString)
  If strInput = vbNullString Then
    strCFG_RT = "3600"
  Else
    lngStart = Len(strInput) - 11
    strEKey = Mid(strInput, lngStart + 1, 10)
    strLens = Right(strInput, 1)
    strInput = XOREncript(Left(strInput, lngStart), strEKey)
  
    lngStart = 1
    lngFinish = Asc(Mid(strLens, 1, 1)) - 127
    strCFG_RT = Mid(strInput, lngStart, lngFinish)
  End If
  
  glngCFG_RT = CLng(strCFG_RT)
  
  
  LC_Initialise = True
  
TidyUpAndExit:
  Exit Function

ErrorTrap:
  LC_Initialise = False
  'MsgBox "Error Reading Security Settings" & vbCrLf & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly + vbExclamation, App.Title
  MsgBox "Error Reading Security Settings" & vbCrLf & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly + vbExclamation, "OpenHR Intranet"
  GoTo TidyUpAndExit
 
End Function

Public Function LC_PCLocked() As Boolean
  
  Dim blnPCLocked As Boolean
  Dim strMessage As String
  Dim strTimeRemaining As String

  LC_Initialise
  
  strMessage = vbNullString
  strTimeRemaining = vbNullString
 
  'Exit the Login Check functions if the PC Lockout functionality has been disabled.
  If Not gblnCFG_PCL Then
    LC_PCLocked = False
    Exit Function
  End If

  LC_ReadCurrentLockStatus
  
  If (Now() >= DateAdd("s", glngCFG_RT, gdtPC_LA)) Then
    blnPCLocked = False
    LC_ResetLock
  Else
    If (gintPC_BA >= gintCFG_BA) Then
      If (DateAdd("s", glngCFG_LD, gdtPC_LKD) >= Now()) Then
        blnPCLocked = True
      Else
        blnPCLocked = False
        LC_ResetLock
      End If
    Else
      blnPCLocked = False
    End If
  End If
  
  If blnPCLocked Then
    strTimeRemaining = LC_FormatDateTimeMessage(DateDiff("s", Now(), DateAdd("s", glngCFG_LD, gdtPC_LKD)))
    strMessage = "This PC has been temporarily locked from using OpenHR." & vbCrLf & vbCrLf & "The lock will be removed in " & strTimeRemaining & "."
    'MsgBox strMessage, vbOKOnly + vbExclamation, App.Title
    MsgBox strMessage, vbOKOnly + vbExclamation, "OpenHR Intranet"
  End If
  
  'Re-initialise PC variables as these are always refreshed by "LC_ReadCurrentLockStatus"
  gintPC_BA = 0
  gdtPC_LA = 0
  gdtPC_LKD = 0
  
  LC_PCLocked = blnPCLocked
  
End Function

Private Function LC_ReadCurrentLockStatus() As Boolean

  Dim strPC_BA As String
  Dim strPC_LA As String
  Dim strPC_LKD As String

  Dim strInput As String
  Dim strEKey As String
  Dim strLens As String
  Dim lngStart As Long
  Dim lngFinish As Long
  
  On Error GoTo ErrorTrap
  
  strInput = GetPCSetting("Misc", "ASRSysParam2", vbNullString)
  If strInput = vbNullString Then
    strPC_BA = "0"
    strPC_LKD = "0"
  
  Else
    lngStart = Len(strInput) - 12
    strEKey = Mid(strInput, lngStart + 1, 10)
    strLens = Right(strInput, 2)
    strInput = XOREncript(Left(strInput, lngStart), strEKey)
  
    lngStart = 1
    lngFinish = Asc(Mid(strLens, 1, 1)) - 127
    strPC_BA = Mid(strInput, lngStart, lngFinish)
    
    lngStart = lngStart + lngFinish
    lngFinish = Asc(Mid(strLens, 2, 1)) - 127
    strPC_LKD = Mid(strInput, lngStart, lngFinish)
    
  End If
  
  gintPC_BA = strPC_BA
  gdtPC_LKD = strPC_LKD


  strInput = GetPCSetting("Misc", "ASRSysParam4", vbNullString)
  If strInput = vbNullString Then
    strPC_LA = "0"
  
  Else
    lngStart = Len(strInput) - 11
    strEKey = Mid(strInput, lngStart + 1, 10)
    strLens = Right(strInput, 1)
    strInput = XOREncript(Left(strInput, lngStart), strEKey)
  
    lngStart = 1
    lngFinish = Asc(Mid(strLens, 1, 1)) - 127
    strPC_LA = Mid(strInput, lngStart, lngFinish)
  
  End If
  
  gdtPC_LA = strPC_LA

  LC_ReadCurrentLockStatus = True
  
TidyUpAndExit:
  Exit Function

ErrorTrap:
  LC_ReadCurrentLockStatus = False
  'MsgBox "Error Reading PC Security Settings" & vbCrLf & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly + vbExclamation, App.Title
  MsgBox "Error Reading PC Security Settings" & vbCrLf & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly + vbExclamation, "OpenHR Intranet"
  GoTo TidyUpAndExit
    
End Function

Private Function LC_ResetLock() As Boolean
  
  'Exit the Login Check functions if the PC Lockout functionality has been disabled.
  If Not gblnCFG_PCL Then
    LC_ResetLock = True
    Exit Function
  End If
  
  LC_SaveCurrentLockStatus "0", "0", "0"
  
End Function

Private Function LC_SaveCurrentLockStatus(pstrBadAttempts As String, pstrLastAttempt As String, pstrLocked As String) As Boolean

  Dim strEKey As String
  Dim strLens As String
  Dim strOutput As String
  Dim lngCount As Long

  On Error GoTo ErrorTrap
  
  strOutput = pstrBadAttempts & pstrLocked
  strLens = Chr(Len(pstrBadAttempts) + 127) & Chr(Len(pstrLocked) + 127)
  strEKey = vbNullString
  For lngCount = 1 To 10
    strEKey = strEKey & Chr(Int(Rnd * 255) + 1)
  Next
  strOutput = XOREncript(strOutput, strEKey) & strEKey & strLens
  
  SavePCSetting "Misc", "ASRSysParam2", strOutput


  strOutput = pstrLastAttempt
  strLens = Chr(Len(pstrLastAttempt) + 127)
  strEKey = vbNullString
  For lngCount = 1 To 10
    strEKey = strEKey & Chr(Int(Rnd * 255) + 1)
  Next
  strOutput = XOREncript(strOutput, strEKey) & strEKey & strLens
  
  SavePCSetting "Misc", "ASRSysParam4", strOutput
  
  
  LC_SaveCurrentLockStatus = True
  
TidyUpAndExit:
  Exit Function

ErrorTrap:
  LC_SaveCurrentLockStatus = False
  'MsgBox "Error Saving PC Security Settings" & vbCrLf & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly + vbExclamation, App.Title
  MsgBox "Error Saving PC Security Settings" & vbCrLf & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly + vbExclamation, "OpenHR Intranet"
  GoTo TidyUpAndExit
  
End Function

Public Function LC_SaveSettingsToRegistry(strCFG_PCL As String, strCFG_BA As String, strCFG_RT As String, strCFG_LD As String) As Boolean
  
  Dim strEKey As String
  Dim strLens As String
  Dim strOutput As String
  Dim lngCount As Long

  On Error GoTo ErrorTrap
 
  LC_Initialise
  
  'Reset the PC Lock as there must have been a successful login!
  LC_ResetLock
  
  strOutput = strCFG_PCL & strCFG_BA & strCFG_LD
  strLens = Chr(Len(strCFG_PCL) + 127) & Chr(Len(strCFG_BA) + 127) & _
            Chr(Len(strCFG_LD) + 127)
  strEKey = vbNullString
  For lngCount = 1 To 10
    strEKey = strEKey & Chr(Int(Rnd * 255) + 1)
  Next
  strOutput = XOREncript(strOutput, strEKey) & strEKey & strLens
  
  SavePCSetting "Misc", "ASRSysParam1", strOutput

  
  strOutput = strCFG_RT
  strLens = Chr(Len(strCFG_RT) + 127)
  strEKey = vbNullString
  For lngCount = 1 To 10
    strEKey = strEKey & Chr(Int(Rnd * 255) + 1)
  Next
  strOutput = XOREncript(strOutput, strEKey) & strEKey & strLens
  
  SavePCSetting "Misc", "ASRSysParam3", strOutput

TidyUpAndExit:
  Exit Function

ErrorTrap:
  'MsgBox "Error Saving Security Settings" & vbCrLf & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly + vbExclamation, App.Title
  MsgBox "Error Saving Security Settings" & vbCrLf & vbCrLf & Err.Number & " " & Err.Description, vbOKOnly + vbExclamation, "OpenHR Intranet"
  GoTo TidyUpAndExit
  
End Function

Private Function MAPISignon() As Integer
  
  'Begin a MAPI session
  On Error Resume Next
  MAPISignon = True
  
  If MAPISession1.SessionID = 0 Then
    'No session currently exists
    MAPISession1.DownLoadMail = False
    MAPISession1.SignOn
    If Err > 0 Then
      If Err <> 32001 And Err <> 32003 Then
        MsgBox "Email not configured correctly." & vbCrLf & _
               IIf(Err.Description <> vbNullString, "(" & Trim(Err.Description) & ")", ""), _
               vbExclamation, "Mail Error"
      End If
      MAPISignon = False
    Else
      MAPIMessages1.SessionID = MAPISession1.SessionID
    End If
  End If

End Function

Private Function MAPIsignoff() As Integer
  
  'End a MAPI session
  On Error Resume Next
  MAPIsignoff = True
  If MAPISession1.SessionID <> 0 Then
    'Session currently exists
    MAPISession1.SignOff
    If Err > 0 Then
      MsgBox Error$, 48, "Mail Error"
      MAPIsignoff = False
    Else
      MAPIMessages1.SessionID = 0
    End If
  End If

End Function

'Public Function SendMail(psTo As String, Optional psSubject As String, _
'                            Optional psBody As String, Optional psCC As String, _
'                            Optional psBCC As String) As Boolean
'
'  Dim strAddress() As String
'  Dim lngCount As Long
'
'  On Error GoTo ErrorTrap
'
'  'Log onto Mail server if not already logged onto it
'  frmEmailSel.MAPISignon
'  If frmEmailSel.MAPISession1.SessionID <> 0 Then
'    With frmEmailSel.MAPIMessages1
'      .Compose
'
'      If Trim(psTo) <> vbNullString Then
'        strAddress = Split(psTo, ";")
'        For lngCount = 0 To UBound(strAddress)
'          If Trim(strAddress(lngCount)) <> vbNullString Then
'            .RecipIndex = .RecipCount
'            .RecipAddress = Trim(strAddress(lngCount))
'            .RecipType = mapToList
'            .ResolveName
'          End If
'        Next
'      End If
'
'      If Trim(psCC) <> vbNullString Then
'        strAddress = Split(psCC, ";")
'        For lngCount = 0 To UBound(strAddress)
'          If Trim(strAddress(lngCount)) <> vbNullString Then
'            .RecipIndex = .RecipCount
'            .RecipAddress = Trim(strAddress(lngCount))
'            .RecipType = mapCcList
'            .ResolveName
'          End If
'        Next
'      End If
'
'      If Trim(psBCC) <> vbNullString Then
'        strAddress = Split(psBCC, ";")
'        For lngCount = 0 To UBound(strAddress)
'          If Trim(strAddress(lngCount)) <> vbNullString Then
'            .RecipIndex = .RecipCount
'            .RecipAddress = Trim(strAddress(lngCount))
'            .RecipType = mapBccList
'            .ResolveName
'          End If
'        Next
'      End If
'
'      .MsgSubject = psSubject
'
'      ' Send the message
'      .MsgNoteText = psBody
'
'      On Error Resume Next
'      .Send True
'      On Error GoTo ErrorTrap
'
'    End With
'  End If
'
'  SendMail = True
'
'TidyUpAndExit:
'  frmEmailSel.MAPIsignoff
'  Exit Function
'
'ErrorTrap:
'
'  If Err.Number = 32026 Then  'Not supported
'    MsgBox "Error sending email", vbCritical, "Send Mail Message"
'  'TM20020920 Fault 4416 - catch the err if a recipient cannot be resolved.
'  ElseIf Err.Number = MAPIErrors.mapUnknownRecipient Then
'    MsgBox Err.Description, vbExclamation, "Send Mail Message"
'  Else
'    MsgBox Err.Description, vbExclamation, "Send Mail Message"
'  End If
'
'  SendMail = False
'  GoTo TidyUpAndExit
'
'End Function


Public Function SendMail(psTo As String, Optional psSubject As String, _
                            Optional psBody As String, Optional psCC As String, _
                            Optional psBCC As String) As Boolean
  
  Dim strError As String
  
  strError = SendEmailFromClientUsingMAPI(psTo, psCC, psBCC, psSubject, psBody, "", True)
  If strError <> vbNullString Then
    MsgBox strError, vbExclamation, "Send Mail Message"
    SendMail = False
  Else
    SendMail = True
  End If

End Function
  
  
  
Public Function SendEmailFromClientUsingMAPI(strTo As String, strCC As String, strBCC As String, strSubject As String, strMsgText As String, strAttachment As String, blnShowMessage As Boolean) As String

  
  Dim objMapiSession As MAPISession
  Dim objMapiMessages As MAPIMessages
  
  Dim strError As String
  Dim strRecipients As String
  Dim lngRecipType As Long
  
  Dim strArray() As String
  Dim lngIndex As Long
  'Dim blnShowMessage As Boolean
  
  On Error GoTo LocalErr

  Set objMapiSession = frmEmailSel.MAPISession1
  Set objMapiMessages = frmEmailSel.MAPIMessages1


  strError = vbNullString
  'blnShowMessage = False
  
  If Trim(Replace(strTo, ";", "")) = vbNullString Then
    SendEmailFromClientUsingMAPI = "Please select recipient(s) to email"
    Exit Function
  End If
  
  
  If objMapiSession.SessionID = 0 Then
    'objMapiSession.DownLoadMail = False
    objMapiSession.SignOn
    objMapiMessages.SessionID = objMapiSession.SessionID
  End If
    
    
  With objMapiMessages
    .AddressResolveUI = True
    .Compose

    For lngRecipType = 1 To 3
    
      Select Case lngRecipType
      Case 1: strRecipients = strTo
      Case 2: strRecipients = strCC
      Case 3: strRecipients = strBCC
      End Select
      
      If (Trim(strRecipients) <> vbNullString) Then
        strArray = Split(strRecipients, ";")
        For lngIndex = LBound(strArray) To UBound(strArray)
          If Trim(strArray(lngIndex)) <> vbNullString Then
            .RecipIndex = .RecipCount

            .RecipType = lngRecipType
            .RecipDisplayName = Trim(strArray(lngIndex))
            .RecipAddress = "smtp:" & Trim(strArray(lngIndex))
            '.ResolveName



          End If
        Next
      End If

    Next

    .MsgSubject = strSubject
    .MsgNoteText = strMsgText & " "

    If (Trim(strAttachment) <> vbNullString) Then
      strArray = Split(strAttachment, ";")
      For lngIndex = LBound(strArray) To UBound(strArray)
        If Trim(strArray(lngIndex)) <> vbNullString Then
          .AttachmentIndex = lngIndex
          .AttachmentPathName = Trim(strArray(lngIndex))
          '.AttachmentName = FileOnlyFromFullPath(Trim(strArray(lngIndex)))
          '.AttachmentPosition = lngIndex
          '.AttachmentType = 0
        End If
      Next
    End If

    .Send blnShowMessage

  End With

TidyAndExit:
  If objMapiSession.SessionID <> 0 Then
    objMapiSession.SignOff
  End If
  objMapiMessages.SessionID = 0
  
  SendEmailFromClientUsingMAPI = strError

Exit Function

LocalErr:
  If blnShowMessage = False Then
    'Simple MAPI is no longer supported in Microsoft Outlook 2007. It is still supported by Exchange Server 2003.
    'http://msdn.microsoft.com/en-us/library/cc815424.aspx
    blnShowMessage = True
    Resume 0

  ElseIf Err.Number = 32001 Or Err = 32003 Then
    Resume Next
  
  Else
    strError = "Error sending email " & vbCrLf & Err.Description & " (MAPI)"

    GoTo TidyAndExit
  
  End If

End Function


Public Function DaysInMonth(pdtDate) As Integer
  'Return the number of days in the month
  Dim dtNextMonth As Date
  dtNextMonth = DateAdd("m", 1, CDate(pdtDate))
  DaysInMonth = Day(DateAdd("d", Day(dtNextMonth) * -1, dtNextMonth))
End Function

Function GetSystemDecimalSeparator() As String
  ' Return the system decimal separator.
  Dim lngLength As Long
  Dim sBuffer As String * 100
  
  lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, sBuffer, 99)
  GetSystemDecimalSeparator = Left(sBuffer, lngLength - 1)
  
End Function

Public Function ConvertSQLDateToLocale(psSQLDate As String) As String
  ' Convert the given date string (mm/dd/yyyy) into the locale format.
  ' NB. This function assumes a sensible locale format is used.
  Dim fDaysDone As Boolean
  Dim fMonthsDone As Boolean
  Dim fYearsDone As Boolean
  Dim iLoop As Integer
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

Public Function ConvertSQLDateToTime(psSQLDate As String) As String
  
  Dim sTempDate As String
  
  sTempDate = IIf(IsNull(psSQLDate), "", psSQLDate)
  
  ConvertSQLDateToTime = Format(sTempDate, "hh:nn")
    
End Function

Public Sub SaveRegistrySetting(psAppName As String, psSection As String, psKey As String, pvValue As Variant)
  ' Save the given value to the registry with the given registry key values.
  SaveSetting psAppName, psSection, psKey, pvValue

End Sub

Public Function GetRegistrySetting(psAppName As String, psSection As String, psKey As String) As String
  ' Get the required value from the registry with the given registry key values.
  GetRegistrySetting = GetSetting(AppName:=psAppName, Section:=psSection, Key:=psKey)

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



Public Function ValidateDir(psDir As String) As Boolean
  
  Dim fso As New FileSystemObject
  
  On Error Resume Next
  
  ValidateDir = False
  ValidateDir = fso.FolderExists(psDir)
  Set fso = Nothing
  
End Function

Public Function ValidateFilePath(psDir As String) As Boolean
  
  Dim fso As New FileSystemObject
  
  On Error Resume Next
  
  ValidateFilePath = False
  ValidateFilePath = fso.FileExists(psDir)
  Set fso = Nothing
  
End Function

Private Sub Class_Initialize()
  GetLocaleDateFormat
  GetLocaleNumberFormat
  mstrSaveAsValues = "doc|0|dot|1|xls|56|xlt|17"
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
  Dim iLoop As Integer
  Dim fDaysDone As Boolean
  Dim fMonthsDone As Boolean
  Dim fYearsDone As Boolean
  
  fDaysDone = False
  fMonthsDone = False
  fYearsDone = False
  sDateFormat = ""
    
  sSysFormat = GetSystemDateFormat
  sSysDateSeparator = GetSystemDateSeparator
    
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
  
  msLocaleDecimalSeparator = GetSystemDecimalSeparator
  msLocaleThousandSeparator = GetSystemThousandSeparator
  
End Sub



Function GetSystemThousandSeparator() As String
  ' Return the system data separator.
  Dim lngLength As Long
  Dim sBuffer As String * 100
  
  lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STHOUSAND, sBuffer, 99)
  GetSystemThousandSeparator = Left(sBuffer, lngLength - 1)
  
End Function


Private Function GetSystemDateSeparator() As String
  ' Return the system data separator.
  Dim lngLength As Long
  Dim sBuffer As String * 100
  
  lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDATE, sBuffer, 99)
  GetSystemDateSeparator = Left(sBuffer, lngLength - 1)
  
End Function

Private Function GetSystemDateFormat() As String
  ' Return the system data format.
  Dim lngLength As Long
  Dim sBuffer As String * 100
  
  lngLength = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE, sBuffer, 99)
  GetSystemDateFormat = Left(sBuffer, lngLength - 1)
  
End Function


Public Function GetScreenTwipsPerPixelX() As Single
  GetScreenTwipsPerPixelX = Screen.TwipsPerPixelX
  
End Function
Public Function GetScreenTwipsPerPixelY() As Single
  GetScreenTwipsPerPixelY = Screen.TwipsPerPixelY
  
End Function


Public Function MessageBox(psPrompt As String, _
  Optional piButtons As VbMsgBoxStyle, _
  Optional psTitle As String) As VbMsgBoxResult
  
  ' Display the given message box.
  If Len(psTitle) = 0 Then
    'psTitle = App.Title
    psTitle = "OpenHR Intranet"
  End If
  
  MessageBox = MsgBox(psPrompt, piButtons, psTitle)
  
End Function

Public Sub OldEditFile(psFilename As String)
'''  ' Launch the default application for the given file.
'''  Dim fOK As Boolean
'''  Dim iLoop As Integer
'''  Dim iIndex As Integer
'''  Dim lngAPIResult As Long
'''  Dim lngKeyHandle As Long
'''  Dim lngTYPE As Long
'''  Dim sFileExtension As String
'''  Dim sRegValue As String
'''  Dim sApplication As String
'''  Dim pid&
'''
'''  Const iVALUESIZE = 1024
'''
'''  fOK = False
'''
'''  ' Get the extension of the file to open.
'''  iIndex = InStrRev(psFileName, ".")
'''  If iIndex > 0 Then
'''    fOK = True
'''    sFileExtension = Mid(psFileName, iIndex)
'''  Else
'''    ' No extension on the file.
'''    MsgBox "Error opening the selected file." & vbCrLf & _
'''      "The file has no extension.", _
'''      vbExclamation + vbOKOnly, App.Title
'''  End If
'''
'''  If fOK Then
'''    ' Open the registry key for the given file extension's class.
'''    lngAPIResult = RegOpenKeyEx(HKEY_CLASSES_ROOT, sFileExtension, 0, KEY_ALL_ACCESS, lngKeyHandle)
'''    fOK = (lngAPIResult = ERROR_SUCCESS)
'''
'''    If Not fOK Then
'''      MsgBox "Error opening the selected file." & vbCrLf & _
'''        "Unable to open the file extension's key in the registry.", _
'''        vbExclamation + vbOKOnly, App.Title
'''    End If
'''  End If
'''
'''  If fOK Then
'''    ' Get the class name of the given file extension.
'''    sRegValue = Space$(iVALUESIZE)
'''
'''    lngAPIResult = RegQueryValueEx(lngKeyHandle, "", 0, lngTYPE, ByVal sRegValue, iVALUESIZE)
'''    RegCloseKey (lngKeyHandle)
'''
'''    fOK = ((lngAPIResult = ERROR_SUCCESS) And (lngTYPE = REG_SZ))
'''
'''    If Not fOK Then
'''      MsgBox "Error opening the selected file." & vbCrLf & _
'''        "Unable to read the file extension's class from the registry.", _
'''        vbExclamation + vbOKOnly, App.Title
'''    End If
'''  End If
'''
'''  If fOK Then
'''    ' Extract the file extension's class from the value read from the registry.
'''    For iLoop = 1 To Len(sRegValue)
'''      If Mid(sRegValue, iLoop, 1) = vbNullChar Then
'''        sRegValue = Left(sRegValue, iLoop - 1)
'''        Exit For
'''      End If
'''    Next iLoop
'''
'''    fOK = (Len(sRegValue) > 0)
'''
'''    If Not fOK Then
'''      MsgBox "Error opening the selected file." & vbCrLf & _
'''        "Unable to read the file extension's class from the registry.", _
'''        vbExclamation + vbOKOnly, App.Title
'''    End If
'''  End If
'''
'''  If fOK Then
'''    ' Open the given file's class registry entry.
'''    lngAPIResult = RegOpenKeyEx(HKEY_CLASSES_ROOT, sRegValue & "\shell\open\command", 0, KEY_ALL_ACCESS, lngKeyHandle)
'''    sRegValue = Space$(iVALUESIZE)
'''
'''    fOK = (lngAPIResult = ERROR_SUCCESS)
'''
'''    If Not fOK Then
'''      MsgBox "Error opening the selected file." & vbCrLf & _
'''        "Unable to open the given file's class key in the registry.", _
'''        vbExclamation + vbOKOnly, App.Title
'''    End If
'''  End If
'''
'''  If fOK Then
'''    ' Read the given file's default application from the registry.
'''    lngAPIResult = RegQueryValueEx(lngKeyHandle, "", 0, lngTYPE, ByVal sRegValue, iVALUESIZE)
'''    RegCloseKey (lngKeyHandle)
'''
'''    fOK = ((lngAPIResult = ERROR_SUCCESS) And ((lngTYPE = REG_EXPAND_SZ) Or (lngTYPE = REG_SZ)))
'''
'''    If Not fOK Then
'''      MsgBox "Error opening the selected file." & vbCrLf & _
'''        "Unable to read the default application for the given file's class from the registry.", _
'''        vbExclamation + vbOKOnly, App.Title
'''    End If
'''  End If
'''
'''  If fOK Then
'''    ' Extract the file extension's class from the value read from the registry.
'''    For iLoop = 1 To Len(sRegValue)
'''      If Mid(sRegValue, iLoop, 1) = vbNullChar Then
'''        sRegValue = Left(sRegValue, iLoop - 1)
'''        Exit For
'''      End If
'''    Next iLoop
'''
'''    fOK = (Len(sRegValue) > 0)
'''    If Not fOK Then
'''      MsgBox "Error opening the selected file." & vbCrLf & _
'''        "Unable to read the default application for the given file's class from the registry.", _
'''        vbExclamation + vbOKOnly, App.Title
'''    End If
'''  End If
'''
'''  If fOK Then
'''    ' Remove any parameter definitions from the default application string.
'''    sApplication = sRegValue
'''    For iLoop = (Len(sRegValue) - 1) To 1 Step -1
'''      If (Mid(sRegValue, iLoop, 2) = " /") Or (Mid(sRegValue, iLoop, 2) = " """) Then
'''        sApplication = Left(sRegValue, iLoop - 1)
'''      End If
'''    Next iLoop
'''
'''    sApplication = sApplication & " """ & psFileName & """"
'''
'''    pid = Shell(sApplication, vbNormalFocus)
'''
'''    fOK = (pid <> 0)
'''
'''    If Not fOK Then
'''      MsgBox "Error opening the selected file." & vbCrLf & _
'''        "Unable to run the selected file's default application.", _
'''        vbExclamation + vbOKOnly, App.Title
'''    End If
'''  End If
'''
'''  If fOK Then
'''    ' Launch the default graphic editing package.
'''
'''    Set mobjProc = CreateObject("procWatcher.dwAppWatch")
'''    mobjProc.SetAppWatch pid
'''    mobjProc.SetAppCallback Me
'''
'''    ' Lock OpenHR until the photo edition has been done.
'''    frmSystemLocked.LockType = 1
'''    frmSystemLocked.Show vbModal
'''  End If
'''
End Sub
Public Sub EditFile(psFilename As String, pbFileEncrypted As Boolean, pstrRealName As String, pbIsReadOnly As Boolean)
  ' Launch the default application for the given file.
  On Error GoTo Edit_Error
  
  Dim fOk As Boolean
  Dim iIndex As Integer
  Dim lngTemp As Long
  Dim sPath As String
  Dim sFile As String
  Dim sFolder As String
  Dim strEditFile As String
  Dim objEncryption As New clsEncryption
  Dim strTempFileName As String
  Dim iPreviousAttributes As Integer
  Dim bResetFlags As Boolean

  fOk = False
    
  If pbFileEncrypted Then
    strTempFileName = Space(1024)
    Call GetTempPath(1024, strTempFileName)
    strEditFile = Left(strTempFileName, InStr(strTempFileName, Chr(0)) - 1) & GetFileNameOnly(pstrRealName)
    
    objEncryption.DecryptFile psFilename, strEditFile, True, "230678" & mstrCurrentSessionKey
  Else
    strEditFile = psFilename
  End If
    
  ' Force file to be read only if necessary
  If pbIsReadOnly Then
    Set mobjFileInfo = mobjFileSystem.GetFile(strEditFile)
    iPreviousAttributes = mobjFileInfo.Attributes
    bResetFlags = True
    mobjFileInfo.Attributes = ReadOnly
  End If
    
  iIndex = InStrRev(strEditFile, "\")
  If iIndex > 0 Then
    fOk = True
    sFile = Mid(strEditFile, iIndex + 1)
    sFolder = Left(strEditFile, iIndex - 1)
  Else
    'MsgBox "Error opening the selected file." & vbCrLf & _
      "The file has no folder.", _
      vbExclamation + vbOKOnly, App.Title
    MsgBox "Error opening the selected file." & vbCrLf & _
      "The file has no folder.", _
      vbExclamation + vbOKOnly, "OpenHR Intranet"
  End If
  
  If fOk Then
    ' Initialise an empty string to pass to the API call
    sPath = Space(255)
    ' Get the executables path for the path & document filename
    lngTemp = FindExecutable(sFile, sFolder, sPath)

    ' If we have got a valid executable to run the document with then continue
    If Len(Trim(sPath)) > 1 Then
      ' For some reason W95 adds /n or /e onto the end of the path, so lose anything
      ' after the xxx.exe
      sPath = Left(sPath, InStr(LCase(sPath), ".exe") + 3)

      ' Tidy up the path returned from the API. Trust me, its needed !
      sPath = Replace(Trim(sPath), Chr(0), "")
  
      ' If the executable path returned from the API is the same as the OLE
      ' documents path and filename, then we are running an exe, so empty
      ' the sPath variable, otherwise add a space.
      If sPath = sFolder & "\" & sFile Then sPath = "" Else sPath = sPath & " "
      
      ' JDM - 21/06/2004 - Fault 5314 - Botch to get Microsoft Word working
      If LCase(Trim(Right(sFile, 3))) = "doc" Then
        sPath = sPath + "/x "
      End If

      ' Shell out the process and capture the ID
      lngTemp = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(sPath & Chr(34) & sFolder & "\" & sFile & Chr(34), vbNormalFocus))
  
      frmSystemLocked.ProcessID = lngTemp
  
      ' Show the system locked form
      frmSystemLocked.Show vbModal
    Else
      fOk = False
      'MsgBox "No application is associated with this file.", _
        vbExclamation + vbOKOnly, App.Title
      MsgBox "No application is associated with this file.", _
        vbExclamation + vbOKOnly, "OpenHR Intranet"
    End If
  End If

  ' Need to re-encrypt document back to the original
  If pbFileEncrypted And fOk Then
    objEncryption.EncryptFile strEditFile, psFilename, True, "230678" & mstrCurrentSessionKey
  End If

  ' Reset the previous state flag
  If pbIsReadOnly And bResetFlags Then
    Set mobjFileInfo = mobjFileSystem.GetFile(strEditFile)
    mobjFileInfo.Attributes = iPreviousAttributes
  End If

  Set objEncryption = Nothing

  Exit Sub

Edit_Error:
  
  If LCase(Err.Description) = "permission denied" Then
    bResetFlags = False
  Else
    'MsgBox "Error attempting to invoke the default editor for this file type." & vbCrLf & vbCrLf & _
         "Trying to call : Shell " & sPath & Chr(34) & sFolder & "\" & sFile & Chr(34) & vbCrLf & _
         "(" & Err.Number & " - " & Err.Description & ")", _
         vbExclamation + vbOKOnly, App.Title
    MsgBox "Error attempting to invoke the default editor for this file type." & vbCrLf & vbCrLf & _
         "Trying to call : Shell " & sPath & Chr(34) & sFolder & "\" & sFile & Chr(34) & vbCrLf & _
         "(" & Err.Number & " - " & Err.Description & ")", _
         vbExclamation + vbOKOnly, "OpenHR Intranet"
  End If

End Sub

'''Public Sub dwAppTerminated(obj As Object)
'''  ' Pick up the signal that says that the graphic editing application
'''  ' has terminated.
'''  Dim lngCount As Long
'''
'''  ' Get rid of the locking form.
'''  Unload frmSystemLocked
'''
'''  ' Highlight the selected photo file in the listbox.
'''  Set mobjProc = Nothing
'''
'''End Sub

Public Sub ClosePopup()
'  On Error GoTo ErrorTrap
'
'  If Not mobjPopup Is Nothing Then
'    mobjPopup.ClosePopup
'    Set mobjPopup = Nothing
'  End If
'
'  Exit Sub
'
'ErrorTrap:

End Sub

Public Sub ShowPopup(psMessage As String)

'  On Error GoTo ErrorTrap
'
'  If mobjPopup Is Nothing Then
'   Set mobjPopup = New prjIntranetPopup.clsIntranetPopup
'  End If
'
'  mobjPopup.ShowPopup psMessage
'
'  Exit Sub
'
'ErrorTrap:

End Sub


Public Property Get LocaleDecimalSeparator() As String
  LocaleDecimalSeparator = msLocaleDecimalSeparator

End Property

Public Property Get LocaleDateFormat() As String
  LocaleDateFormat = msLocaleDateFormat

End Property

Public Property Get LocaleDateSeparator() As String
  LocaleDateSeparator = GetSystemDateSeparator

End Property

Public Property Get LocaleThousandSeparator() As String
  LocaleThousandSeparator = msLocaleThousandSeparator

End Property

Public Function PrinterCount() As Long
  PrinterCount = Printers.Count
End Function


Public Function PrinterName(lngIndex As Long) As String
  PrinterName = Printers(lngIndex).DeviceName
End Function

Private Function XOREncript(strInput, strKey) As String

  Dim lngCount As Long
  Dim strOutput As String
  Dim strChar As String
  
  For lngCount = 1 To Len(strInput)
    strChar = Mid(strKey, lngCount Mod Len(strKey) + 1, 1)
    strOutput = strOutput & Chr(Asc(strChar) Xor Asc(Mid(strInput, lngCount, 1)))
  Next

  XOREncript = strOutput

End Function

Public Function GetPCSetting(strSection As String, strKey As String, varDefault As Variant) As String
  GetPCSetting = GetSetting("HR Pro", strSection, strKey, varDefault)
End Function

Public Function SavePCSetting(strSection As String, strKey As String, varSetting As Variant) As Boolean
  'Trap error in case user doesn't have permission to write to the registry
  On Local Error Resume Next
  SaveSetting "HR Pro", strSection, strKey, varSetting
End Function

Private Function GetUNCOnly(ByVal pstrFileName As String) As String
    
  On Local Error GoTo GetUNCPath_Err
  Dim strMSG As String
  Dim lngReturn As Long
  Dim strLocalName As String
  Dim strRemoteName As String
  Dim lngRemoteName As Long
  Dim strUNCPath As String
  strLocalName = GetDriveOnly(pstrFileName)
  strRemoteName = String(255, Chr(32))
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
    
  On Local Error GoTo ConvertUNCPath_Err
  Dim strMSG As String
  Dim lngReturn As Long
  Dim strLocalName As String
  Dim strRemoteName As String
  Dim lngRemoteName As Long
  Dim strUNCPath As String
  strLocalName = GetDriveOnly(pstrFileName)
  strRemoteName = String(255, Chr(32))
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
Public Function GetFileNameOnly(pstrFilePath As String) As String
  Dim astrPath() As String
  astrPath = Split(pstrFilePath, "\")
  GetFileNameOnly = astrPath(UBound(astrPath))
End Function

Public Function IsFileOnNetwork(ByVal pstrFileName As String) As Integer

  Dim lngReturn As Long
  Dim strLocalName As String
  Dim strRemoteName As String
  Dim lngRemoteName As Long

  strLocalName = GetDriveOnly(pstrFileName)
  strRemoteName = String(255, Chr(32))
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
Public Function GetPathOnly(pstrFilePath As String, pbStripDriveLetter As Boolean) As String
   
  Dim L As Integer
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
  Wend
  
End Function

' Wrapper for the encryption module as we don't want users to be able to see the super secret key
Public Function OLEEncryptFile(strInFile As String, strUploadPath As String, strKey As String) As String

  Dim objEncrypt As New clsEncryption
  Dim strOutFile As String
  
  strOutFile = Space(1024)
  Call GetTempFileName(strUploadPath, "_T", 0, strOutFile)
  strOutFile = Left(strOutFile, InStr(strOutFile, Chr(0)) - 1)
  
  objEncrypt.EncryptFile strInFile, strOutFile, True, "230678" & strKey
  Set objEncrypt = Nothing

  OLEEncryptFile = strOutFile

End Function

' Returns a message if the filename length is too long to embed/link
Public Function CheckOLEFileNameLength(strFilename As String) As String

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
Public Function CheckOLEPictureExtension(strFilename As String) As Boolean

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
Public Function FileSize(strFilename As String) As String

  On Error GoTo ErrorTrap
  Dim strSize As String

  Set mobjFileInfo = mobjFileSystem.GetFile(strFilename)
  strSize = mobjFileInfo.Size

TidyUpAndExit:
  FileSize = strSize
  Set mobjFileInfo = Nothing
  Exit Function
  
ErrorTrap:
  strSize = "Unknown"
  GoTo TidyUpAndExit

End Function

Public Function FileLastModified(strFilename As String) As String

  On Error GoTo ErrorTrap
  Dim strDate As String

  Set mobjFileInfo = mobjFileSystem.GetFile(strFilename)
  strDate = mobjFileInfo.DateLastModified

TidyUpAndExit:
  FileLastModified = strDate
  Set mobjFileInfo = Nothing
  Exit Function
  
ErrorTrap:
  strDate = "Unknown"
  GoTo TidyUpAndExit

End Function

' Uploads a specified file to one of the ole paths
Public Sub UploadFile(strFilename As String, strLocation As String)

  On Error GoTo ErrorTrap

  Dim strUploadPath As String

  ' Upload the file
  mobjFileSystem.CopyFile strFilename, strLocation & "\", True

TidyUpAndExit:
  Exit Sub

ErrorTrap:
  GoTo TidyUpAndExit

End Sub

' Returns array of files in the ole path
Public Function FolderList(pstrLocation As String) As Variant

  On Error GoTo ErrorTrap

  Set FolderList = mobjFileSystem.GetFolder(pstrLocation)

TidyUpAndExit:
  Exit Function

ErrorTrap:
  GoTo TidyUpAndExit

End Function

Public Function NiceSize(pstrSize As String) As String

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


Public Function GetOfficeWordVersion() As Integer

  Dim App As Object

  On Error GoTo NotInstalled

  If miOfficeVersion_Word = 0 Then
    Set App = CreateObject("Word.Application")
    miOfficeVersion_Word = Val(App.Version)
    App.Quit
  End If

TidyUpAndExit:
  GetOfficeWordVersion = miOfficeVersion_Word
  Set App = Nothing

Exit Function

NotInstalled:
  miOfficeVersion_Word = -1
  Resume TidyUpAndExit

End Function


Public Function GetOfficeExcelVersion() As Integer

  Dim App As Object

  On Error GoTo NotInstalled

  If miOfficeVersion_Excel = 0 Then
    Set App = CreateObject("Excel.Application")
    miOfficeVersion_Excel = Val(App.Version)
    App.Quit
  End If

TidyUpAndExit:
  GetOfficeExcelVersion = miOfficeVersion_Excel
  Set App = Nothing

Exit Function

NotInstalled:
  miOfficeVersion_Excel = -1
  Resume TidyUpAndExit

End Function

