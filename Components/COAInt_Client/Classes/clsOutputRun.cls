VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOutputRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IObjectSafetyTLB.IObjectSafety

Private mobjOutputType As Object
'Private mfrmOutput As frmOutputOptions
Private mcolStyles As Collection
Private mcolMerges As Collection
Private mcolColumns As Collection
Private mstrErrorMessage As String

Private mlngFormat As Long
Private mstrFunction As String
Private mstrEmailAddresses As String
Private mstrEmailSubject As String
Private mstrEmailAttachAs As String
Private mblnPageRange As Boolean
Private mblnPrintData As Boolean
Private mstrPrinterName As String

Private mstrDefaultPrinter As String
Private mblnSizeColumnsIndependently As Boolean
Private mlngHeaderRows As Long
Private mlngHeaderCols As Long
Private mstrArray() As String

Private mstrSaveAsValues As String

Private mblnData As Boolean
Private mblnCSV As Boolean
Private mblnHTML As Boolean
Private mblnWord As Boolean
Private mblnExcel As Boolean
Private mblnChart As Boolean
Private mblnPivot As Boolean
Private mblnPivotSuppressBlanks As Boolean
Private mblnIndicatorColumn As Boolean
Private mblnPageTitles As Boolean

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
                pdwEnabledOptions = IIf(m_fSafeForScripting, INTERFACESAFE_FOR_UNTRUSTED_CALLER, 0)
            
            Case IID_IPersistStorage, IID_IPersistStream, IID_IPersistPropertyBag
                pdwEnabledOptions = IIf(m_fSafeForInitializing, INTERFACESAFE_FOR_UNTRUSTED_DATA, 0)
                       
            Case IID_IDispatchX
                pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_DATA Or INTERFACESAFE_FOR_UNTRUSTED_CALLER
            
            Case Else
                Err.Raise E_NOINTERFACE
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
                If ((dwEnabledOptions And dwOptionsSetMask) <> INTERFACESAFE_FOR_UNTRUSTED_CALLER) Then
                    Err.Raise E_FAIL
                    Exit Sub
                Else
                MsgBox "here2"
                    If Not m_fSafeForScripting Then
                        Err.Raise E_FAIL
                    End If
                    Exit Sub
                End If

            Case IID_IPersistStorage, IID_IPersistStream, IID_IPersistPropertyBag
                If ((dwEnabledOptions And dwOptionsSetMask) <> INTERFACESAFE_FOR_UNTRUSTED_DATA) Then
                    Err.Raise E_FAIL
                    Exit Sub
                Else
                    If Not m_fSafeForInitializing Then
                        Err.Raise E_FAIL
                    End If
                    Exit Sub
                End If

            Case IID_IDispatchX
             If (dwEnabledOptions And dwOptionsSetMask) = INTERFACESAFE_FOR_UNTRUSTED_CALLER Then
                Exit Sub
              End If

            Case Else
                Err.Raise E_NOINTERFACE
                Exit Sub
        End Select
    End If
    
End Sub

'''Private Function ConvertToGrid(pstrArray() As String, pGrdTemp As SSDBGrid) As SSDBGrid
'''
'''  Dim strAddLine As String
'''  Dim iCol As Long
'''  Dim iRow As Long
'''
'''  With pGrdTemp
'''    'initialise grid
'''    .RemoveAll
'''    .Columns.RemoveAll
'''    .ColumnHeaders = False
'''
'''    'add the columns to the grid
'''    For iCol = 0 To UBound(pstrArray) Step 1
'''      .Columns.Add iCol
'''    Next iCol
'''
'''    For iRow = 0 To UBound(pstrArray, 2) Step 1
'''
'''      strAddLine = vbNullString
'''
'''      For iCol = 0 To UBound(pstrArray) Step 1
'''        strAddLine = strAddLine & pstrArray(iCol, iRow)
'''
'''        If iCol < UBound(pstrArray) Then
'''          strAddLine = strAddLine & vbTab
'''        End If
'''      Next iCol
'''
'''      .AddItem strAddLine
'''
'''    Next iRow
'''
'''  End With
'''
'''  If mblnSizeColumnsIndependently Then
'''    Set ConvertToGrid = AutoSizeGridColumms(pGrdTemp)
'''  Else
'''    Set ConvertToGrid = pGrdTemp
'''  End If
'''
'''End Function
'''
'''Private Function AutoSizeGridColumms(pSSDBGrid As SSDBGrid) As SSDBGrid
'''
'''  'formats the size of the columns so all data is shown
'''
'''  Dim iCol As Long
'''  Dim iRow As Long
'''  Dim lngTextWidth As Long
'''  Dim varBookmark As Variant
'''
'''  Const TextWidthMultiplier = 80
'''
'''  lngTextWidth = 0
'''
'''  With pSSDBGrid
'''    .Redraw = False
'''    .MoveFirst
'''    For iCol = 0 To .Columns.Count - 1 Step 1
'''
'''      lngTextWidth = ((Len(Trim(.Columns(iCol).CellText(varBookmark)))) * TextWidthMultiplier)
'''
'''      If .Columns(iCol).Visible Then
'''        For iRow = 0 To .Rows - 1 Step 1
'''          varBookmark = .AddItemBookmark(iRow)
'''
'''          If ((Len(Trim(.Columns(iCol).CellText(varBookmark)))) * TextWidthMultiplier) > lngTextWidth Then
'''            lngTextWidth = ((Len(Trim(.Columns(iCol).CellText(varBookmark)))) * TextWidthMultiplier)
'''          End If
'''        Next iRow
'''
'''        .Columns(iCol).Width = lngTextWidth + 150
'''      End If
'''      lngTextWidth = 0
'''    Next iCol
'''    .Redraw = True
'''  End With
'''
'''  Set AutoSizeGridColumms = pSSDBGrid
'''
'''End Function
'''
'''
Private Sub Class_Initialize()

  'Set mfrmOutput = New frmOutputOptions
  Set mcolColumns = New Collection

  mlngHeaderRows = 1
  mlngHeaderCols = 0
  mblnPageTitles = True
  mblnPivotSuppressBlanks = True
  mblnIndicatorColumn = False

  mblnData = True
  mblnCSV = True
  mblnHTML = True
  mblnWord = True
  mblnExcel = True
  mblnChart = True
  mblnPivot = True

  'InitialiseStyles
  Set mcolMerges = New Collection

End Sub

Public Sub InitialiseStyles()

  Dim objStyle As clsOutputStyle
  
  Set mcolStyles = New Collection

  Set objStyle = New clsOutputStyle
  With objStyle
    .Name = "Title"
    .StartCol = glngSettingTitleCol
    .StartRow = glngSettingTitleRow
    .Gridlines = gblnSettingTitleGridlines
    .Bold = gblnSettingTitleBold
    .Underline = gblnSettingTitleUnderline
    .BackCol = glngSettingTitleBackcolour
    .ForeCol = glngSettingTitleForecolour
    .BackCol97 = glngSettingTitleBackcolour97
    .ForeCol97 = glngSettingTitleForecolour97
  End With

  mcolStyles.Add objStyle, objStyle.Name


  Set objStyle = New clsOutputStyle
  With objStyle
    .Name = "Heading"
    .Gridlines = gblnSettingHeadingGridlines
    .Bold = gblnSettingHeadingBold
    .Underline = gblnSettingHeadingUnderline
    .BackCol = glngSettingHeadingBackcolour
    .ForeCol = glngSettingHeadingForecolour
    .BackCol97 = glngSettingHeadingBackcolour97
    .ForeCol97 = glngSettingHeadingForecolour97
    .CenterText = True
  End With

  mcolStyles.Add objStyle, objStyle.Name
  
  
  Set objStyle = New clsOutputStyle
  With objStyle
    .Name = "HeadingCols"
    .StartCol = 0
    .StartRow = 0
    .Gridlines = gblnSettingHeadingGridlines
    .Bold = gblnSettingHeadingBold
    .Underline = gblnSettingHeadingUnderline
    .BackCol = glngSettingHeadingBackcolour
    .ForeCol = glngSettingHeadingForecolour
    .BackCol97 = glngSettingHeadingBackcolour97
    .ForeCol97 = glngSettingHeadingForecolour97
  End With

  mcolStyles.Add objStyle, objStyle.Name


  Set objStyle = New clsOutputStyle
  With objStyle
    .Name = "Data"
    .StartCol = glngSettingDataCol
    .StartRow = glngSettingDataRow
    .Gridlines = gblnSettingDataGridlines
    .Bold = gblnSettingDataBold
    .Underline = gblnSettingDataUnderline
    .BackCol = glngSettingDataBackcolour
    .ForeCol = glngSettingDataForecolour
    .BackCol97 = glngSettingDataBackcolour97
    .ForeCol97 = glngSettingDataForecolour97
  End With

  mcolStyles.Add objStyle, objStyle.Name


  Set objStyle = Nothing

End Sub


Public Function AddColumn(strHeading As String, lngDataType As Long, lngDecimals As Long, bThousandSeparator As Boolean) As Boolean

  Dim objColumn As clsColumn

  On Local Error GoTo LocalErr
  AddColumn = True
  
  Set objColumn = New clsColumn
  objColumn.Heading = strHeading
  objColumn.DataType = lngDataType
  objColumn.DecPlaces = lngDecimals
  objColumn.ThousandSeparator = bThousandSeparator

  mcolColumns.Add objColumn

Exit Function

LocalErr:
  AddColumn = False

End Function


Public Function AddStyle( _
  strType As String, _
  lngStartCol As Long, _
  lngStartRow As Long, _
  lngEndCol As Long, _
  lngEndRow As Long, _
  Optional lngBackCol As Variant, _
  Optional lngForeCol As Variant, _
  Optional blnBold As Variant, _
  Optional blnUnderline As Variant, _
  Optional blnGridLines As Variant, _
  Optional lngBackCol97 As Variant, _
  Optional lngForeCol97 As Variant) As Boolean

  Dim objStyle As clsOutputStyle

  On Local Error GoTo LocalErr
  AddStyle = True

  Set objStyle = New clsOutputStyle

  With objStyle
    .StartCol = lngStartCol '(mcolStyles("Data").StartCol)
    .StartRow = lngStartRow '(mcolStyles("Data").StartRow)
    .EndCol = lngEndCol   '(mcolStyles("Data").StartCol)
    .EndRow = lngEndRow   '(mcolStyles("Data").StartRow)

    Select Case strType
    Case "Title", "Heading", "Data"
      .BackCol = mcolStyles(strType).BackCol
      .ForeCol = mcolStyles(strType).ForeCol
      .Bold = mcolStyles(strType).Bold
      .Underline = mcolStyles(strType).Underline
      .Gridlines = mcolStyles(strType).Gridlines
    End Select

    If Not IsMissing(lngBackCol) Then
      .BackCol = lngBackCol
    End If

    If Not IsMissing(lngForeCol) Then
      .ForeCol = lngForeCol
    End If

    If Not IsMissing(blnBold) Then
      .Bold = blnBold
    End If

    If Not IsMissing(blnUnderline) Then
      .Underline = blnUnderline
    End If

    If Not IsMissing(blnGridLines) Then
      .Gridlines = blnGridLines
    End If

    If Not IsMissing(lngBackCol97) Then
      .BackCol97 = lngBackCol97
    End If

    If Not IsMissing(lngForeCol97) Then
      .ForeCol97 = lngForeCol97
    End If

  End With

  mcolStyles.Add objStyle
  Set objStyle = Nothing


Exit Function

LocalErr:
  AddStyle = False

End Function


Public Function AddMerge(lngStartCol As Long, lngStartRow As Long, _
      lngEndCol As Long, lngEndRow As Long) As Boolean

  Dim objMerge As clsOutputStyle

  On Local Error GoTo LocalErr
  AddMerge = True

  If lngStartCol <> lngEndCol Or lngStartRow <> lngEndRow Then
    Set objMerge = New clsOutputStyle
    objMerge.StartCol = lngStartCol
    objMerge.StartRow = lngStartRow
    objMerge.EndCol = lngEndCol
    objMerge.EndRow = lngEndRow
    
    mcolMerges.Add objMerge
    Set objMerge = Nothing
  End If

Exit Function

LocalErr:
  AddMerge = False

End Function

Public Function ResetColumns()
  Set mcolColumns = Nothing
  Set mcolColumns = New Collection
End Function

Public Function ResetStyles()
  Set mcolStyles = Nothing
  InitialiseStyles
End Function

Public Function ResetMerges()
  Set mcolMerges = Nothing
  Set mcolMerges = New Collection
End Function

Private Sub Class_Terminate()
  'Set mfrmOutput = Nothing
  Set mcolStyles = Nothing
  Set mcolMerges = Nothing
  Set mcolColumns = Nothing
  Set mobjOutputType = Nothing
End Sub


Public Function GetFile()
  On Local Error Resume Next
  GetFile = True
  GetFile = mobjOutputType.GetFile(Me, mcolStyles)
End Function
  
Public Function AddPage(strDefTitle As String, Optional mstrSheetName As String)
  On Local Error Resume Next
  mobjOutputType.AddPage strDefTitle, mstrSheetName, mcolStyles
End Function

Public Property Let SizeColumnsIndependently(ByVal blnNewValue As Boolean)
  mblnSizeColumnsIndependently = blnNewValue
End Property


Public Function DataArray()

  If UBound(mstrArray, 2) < 1 Then
    mstrErrorMessage = "No data to output."
    Exit Function
  End If

  'Not all classes support all properties so catch any errors...
  On Local Error Resume Next
  mobjOutputType.SizeColumnsIndependently = mblnSizeColumnsIndependently
  mobjOutputType.HeaderCols = mlngHeaderCols
  mobjOutputType.HeaderRows = mlngHeaderRows

  On Local Error GoTo 0
  mobjOutputType.DataArray mstrArray(), mcolColumns, mcolStyles, mcolMerges

End Function

''Public Function DataArrayToGrid(strArray() As String, pgrdNew As SSDBGrid)
''
''  'Not all classes support all properties so catch any errors...
''  On Local Error Resume Next
''  mobjOutputType.SizeColumnsIndependently = mblnSizeColumnsIndependently
''  mobjOutputType.HeaderCols = mlngHeaderCols
''  mobjOutputType.HeaderRows = mlngHeaderRows
''
''  'On Local Error GoTo 0
''  If mblnPrintData Then
''    DataGrid ConvertToGrid(strArray(), pgrdNew)
''  Else
''    mobjOutputType.DataArray strArray(), mcolColumns, mcolStyles, mcolMerges
''  End If
''
''End Function

'Public Function RecordProfilePage(pfrmRecProfile As Form, _
'  piPageNumber As Integer) As Boolean
'  'Not all classes support all properties so catch any errors...
'  On Local Error Resume Next
''''  mobjOutputType.SizeColumnsIndependently = mblnSizeColumnsIndependently
''''  mobjOutputType.HeaderCols = mlngHeaderCols
''''  mobjOutputType.HeaderRows = mlngHeaderRows
'
'  On Local Error GoTo 0
'
''''  mobjOutputType.DataArray strArray(), mcolColumns, mcolStyles, mcolMerges
'  RecordProfilePage = mobjOutputType.RecordProfilePage(pfrmRecProfile, piPageNumber, mcolStyles)
'
'End Function

Public Property Get PrintData() As Boolean
  PrintData = mblnPrintData
End Property

Public Property Get PrinterName() As String
  PrinterName = mstrPrinterName
End Property

Public Property Let HeaderRows(ByVal lngNewValue As Long)
  mlngHeaderRows = lngNewValue
End Property

Public Property Let HeaderCols(ByVal lngNewValue As Long)
  mlngHeaderCols = lngNewValue
End Property


'Public Function DataGrid(objNewValue As SSDBGrid)
'
'  Dim strArray() As String
'  Dim lngGridGrp As Long
'  Dim lngGridCol As Long
'  Dim lngGridRow As Long
'  Dim lngGrp As Long
'  Dim lngCol As Long
'  Dim lngRow As Long
'  Dim blnGroupHeaders As Boolean
'
'  On Local Error GoTo LocalErr
'
'  If mblnPrintData Then
'    mobjOutputType.DataGrid objNewValue
'    Exit Function
'  End If
'
'
'  With objNewValue
'
'    ResetMerges
'    blnGroupHeaders = (.GroupHeaders And .Groups.Count > 0)
'
'    .Redraw = False
'
'    'Get count of visible columns
'    lngCol = 0
'    For lngGridCol = 0 To .Cols - 1
'      If .Columns(lngGridCol).Visible Then
'        lngCol = lngCol + 1
'
'        'Check if this is a header column...
'        If .Columns(lngGridCol).ButtonsAlways Then
'          If mlngHeaderCols = lngCol - 1 Then
'            mlngHeaderCols = lngCol
'          End If
'        End If
'      End If
'    Next
'    ReDim Preserve strArray(lngCol - 1, .Rows + IIf(blnGroupHeaders, 1, 0))
'
'    'GROUP HEADERS
'    lngCol = 0
'    If blnGroupHeaders Then
'      mlngHeaderRows = 2
'      For lngGridGrp = 0 To .Groups.Count - 1
'        If .Groups(lngGridGrp).Visible Then
'          strArray(lngCol, 0) = .Groups(lngGridGrp).Caption
'          AddMerge lngCol, 0, lngCol + .Groups(lngGridGrp).Columns.Count - 1, 0
'          lngCol = lngCol + .Groups(lngGridGrp).Columns.Count
'        End If
'      Next
'    End If
'
'    'COLUMN HEADERS
'    lngCol = 0
'    lngRow = IIf(blnGroupHeaders, 1, 0)
'    For lngGridCol = 0 To .Cols - 1
'      If .Columns(lngGridCol).Visible Then
'        strArray(lngCol, lngRow) = .Columns(lngGridCol).Caption
'        lngCol = lngCol + 1
'      End If
'    Next
'
'    'DATA ROWS
'    For lngGridRow = 0 To .Rows - 1
'      lngCol = 0
'      lngRow = lngRow + 1
'      For lngGridCol = 0 To .Cols - 1
'
'        If .Columns(lngGridCol).Visible Then
'          strArray(lngCol, lngRow) = .Columns(lngGridCol).CellText(.AddItemBookmark(lngGridRow))
'          lngCol = lngCol + 1
'        End If
'
'      Next
'
'    Next
'    .Redraw = True
'
'  End With
'
'  DataArray strArray()
'
'Exit Function
'
'LocalErr:
'  mstrErrorMessage = Err.Description
'
'End Function

Public Function Complete()

  If mstrErrorMessage = vbNullString Then
    On Local Error Resume Next
    mobjOutputType.Complete
  End If
  ClearUp

End Function

Public Function ClearUp()
  On Local Error Resume Next
  mobjOutputType.ClearUp
End Function


Private Function CheckEmailAttachment(strExt As String)
  
  Dim lngFound As Long
  
  If mstrEmailAttachAs <> vbNullString Then
    lngFound = InStrRev(mstrEmailAttachAs, ".")
    If lngFound = 0 Then
      mstrEmailAttachAs = mstrEmailAttachAs & "." & strExt
    End If
  End If

End Function


Public Function SendEmail(strAttachment As String) As Boolean

  Dim objOutputEmail As clsGeneral
  'Dim strAddress() As String
  'Dim lngCount As Long

  On Error GoTo LocalErr

  If gblnEmailSystemPermission = False Then
    mstrErrorMessage = "You do not have permission to use email groups."
    SendEmail = False
    Exit Function
  End If
  
  
'  Screen.MousePointer = vbHourglass
'
'  If Trim(Replace(mstrEmailAddresses, ";", "")) = vbNullString Then
'    mstrErrorMessage = "Error sending email (invalid email address)"
'    SendEmail = False
'    Exit Function
'  End If
'
'  frmEmailSel.MAPISignon
'  If frmEmailSel.MAPISession1.SessionID <> 0 Then
'    With frmEmailSel.MAPIMessages1
'      .Compose
'
'      strAddress = Split(mstrEmailAddresses, ";")
'
'      For lngCount = 0 To UBound(strAddress)
'        If Trim(strAddress(lngCount)) <> vbNullString Then
'          .RecipIndex = .RecipCount
'          .RecipAddress = Trim(strAddress(lngCount))
'          .RecipType = mapToList
'          .ResolveName
'        End If
'      Next
'
'
'      .MsgSubject = mstrEmailSubject
'      .MsgNoteText = " "
'      If strAttachment <> "" Then
'        .AttachmentPosition = 0
'        .AttachmentType = 0
'        .AttachmentPathName = strAttachment
'        .AttachmentName = mstrEmailAttachAs
'      End If
'      .Send False
'    End With
'  End If
'  'frmEmailSel.MAPIsignoff
  
  Set objOutputEmail = New clsGeneral
  mstrErrorMessage = objOutputEmail.SendEmailFromClientUsingMAPI(mstrEmailAddresses, "", "", mstrEmailSubject, "", strAttachment, False)
  Set objOutputEmail = Nothing

  'Set frmEmailSel = Nothing
  SendEmail = True

Exit Function

LocalErr:
  mstrErrorMessage = "Error sending email" & _
    IIf(Err.Description <> vbNullString, " (" & Err.Description & ")", vbNullString)
  'On Error Resume Next
  'frmEmailSel.MAPIsignoff
  'Set frmEmailSel = Nothing
  SendEmail = False

End Function


Public Function GetTempFileName(strFilename As String) As String

  Dim strTempFileName As String

  On Local Error GoTo LocalErr

  'Get temp path
  strTempFileName = Space(1024)
  Call GetTempPath(1024, strTempFileName)
  strTempFileName = GetTmpFName
  If InStr(strTempFileName, Chr(0)) > 0 Then
    strTempFileName = Left(strTempFileName, InStr(strTempFileName, Chr(0)) - 1)
  End If

  'temp path + "\" + file name
  If strFilename <> vbNullString Then
    strFilename = Left(strTempFileName, InStrRev(strTempFileName, "\")) & _
                  Mid(strFilename, InStrRev(strFilename, "\") + 1)
  Else
    strFilename = strTempFileName
  End If

  If Dir(strFilename) <> vbNullString Then
    Kill strFilename
  End If

  GetTempFileName = strFilename

Exit Function

LocalErr:
  mstrErrorMessage = Err.Description

End Function


Public Sub ShowFormats(blnData As Boolean, blnCSV As Boolean, blnHTML As Boolean, _
  blnWord As Boolean, blnExcel As Boolean, blnChart As Boolean, blnPivot As Boolean)

  mblnData = blnData
  mblnCSV = blnCSV
  mblnHTML = blnHTML
  mblnWord = blnWord
  mblnExcel = blnExcel
  mblnChart = blnChart
  mblnPivot = blnPivot

End Sub


'Public Property Get cboPageBreak() As ComboBox
  'Set cboPageBreak = mfrmOutput.cboPageBreak
'  mblnPageRange = True
'End Property


Public Function SetOptions( _
    blnPrompt As Boolean, _
    lngFormat As Long, _
    blnScreen As Boolean, _
    blnPrinter As Boolean, _
    strPrinterName As String, _
    blnSave As Boolean, _
    lngSaveExisting As Long, _
    blnEmail As Boolean, _
    strEmailAddresses As String, _
    strEmailSubject As String, _
    strEmailAttachAs As String, _
    strFilename As String) As Boolean

  Dim Prtr As Printer
  Dim objColumn As clsColumn
  Dim blnCancelled As Boolean
  'Dim lngSaveExisting As Long

  On Local Error GoTo LocalErr


  blnCancelled = False

  'mlngEmailGroupID = lngEmailAddr
  mstrEmailAddresses = strEmailAddresses
  mstrEmailSubject = strEmailSubject
  mstrEmailAttachAs = strEmailAttachAs


  mlngFormat = lngFormat
  Select Case mlngFormat
  Case fmtDataOnly
    Set mobjOutputType = New clsOutputGrid

  Case fmtCSV, fmtFixedLengthFile
    Set mobjOutputType = New clsOutputCSV
    CheckEmailAttachment "csv"
    
  Case fmtHTML
    Set mobjOutputType = New clsOutputHTML
    CheckEmailAttachment "htm"

  Case fmtWordDoc
    Set mobjOutputType = New clsOutputWord
    CheckEmailAttachment "doc"

  Case fmtExcelWorksheet
    Set mobjOutputType = New clsOutputExcel
    CheckEmailAttachment "xls"

  Case fmtExcelChart
    Set mobjOutputType = New clsOutputExcel
    mobjOutputType.Chart = True
    CheckEmailAttachment "xls"

  Case fmtExcelPivotTable
    Set mobjOutputType = New clsOutputExcel
    mobjOutputType.PivotTable = True
    CheckEmailAttachment "xls"

  End Select


  mobjOutputType.Parent = Me
  mobjOutputType.Screen = blnScreen
  mobjOutputType.DestPrinter = blnPrinter
  mobjOutputType.Save = blnSave
  mobjOutputType.SaveExisting = lngSaveExisting
  mobjOutputType.Email = blnEmail
  mobjOutputType.FileName = strFilename

  mblnPrintData = (mlngFormat = fmtDataOnly And blnPrinter)
  
  If strPrinterName = "<Default Printer>" Then
    mstrPrinterName = Printer.DeviceName
  Else
    mstrPrinterName = strPrinterName
  End If
  mobjOutputType.PrinterName = mstrPrinterName

  If blnPrinter And Printers.Count = 0 Then
    mstrErrorMessage = "Unable to print as no printers are installed."
  End If


  'MH20040209 Fault 8024
  If Not ValidPrinter(mstrPrinterName) Then
    mstrErrorMessage = _
        "This definition is set to output to printer " & mstrPrinterName & _
        " which is not set up on your PC."
  End If
  
  
  SetOptions = (mstrErrorMessage = vbNullString And Not blnCancelled)

Exit Function

LocalErr:
  If mstrErrorMessage = vbNullString Then
    mstrErrorMessage = Err.Description
  End If

End Function


Public Sub SetPrinter()

  Dim objDefPrinter As cSetDfltPrinter

  If mstrPrinterName <> "<Default Printer>" Then
    mstrDefaultPrinter = Printer.DeviceName
    Set objDefPrinter = New cSetDfltPrinter
    objDefPrinter.SetPrinterAsDefault mstrPrinterName
    Set objDefPrinter = Nothing
  End If

End Sub

Public Sub ResetDefaultPrinter()

  Dim objDefPrinter As cSetDfltPrinter

  If mstrPrinterName <> "<Default Printer>" Then
    Set objDefPrinter = New cSetDfltPrinter
    objDefPrinter.SetPrinterAsDefault mstrDefaultPrinter
    Set objDefPrinter = Nothing
  End If

End Sub

Public Property Let FileDelimiter(ByVal strNewValue As String)
  'clsOutputCSV only
  On Local Error Resume Next
  mobjOutputType.FileDelimiter = strNewValue
End Property

Public Property Let EncloseInQuotes(ByVal blnNewValue As Boolean)
  'clsOutputCSV only
  On Local Error Resume Next
  mobjOutputType.EncloseInQuotes = blnNewValue
End Property

Public Property Let ApplyStyles(ByVal blnNewValue As Boolean)
  On Local Error Resume Next
  mobjOutputType.ApplyStyles = blnNewValue
End Property


Public Sub OpenProgress(strType As String, strName As String, lngMaxValue As Long)
'
'  On Local Error GoTo LocalErr
'
'  With gobjProgress
'
'    'If .Visible = False Then
'
'      Select Case mlngFormat
'      Case fmtExcelChart, OutputFormats.fmtExcelPivotTable, OutputFormats.fmtExcelWorksheet
'        .AviFile = App.Path & "\videos\excel.avi"
'      Case fmtWordDoc
'        .AviFile = App.Path & "\videos\word.avi"
'      Case fmtHTML
'        .AviFile = App.Path & "\videos\internet.avi"
'      Case Else
'        .AviFile = App.Path & "\videos\report.avi"
'      End Select
'
'      .NumberOfBars = 1
'      .Time = False
'      .Cancel = True
'      .Bar1MaxValue = lngMaxValue
'
'      .Caption = strType
'      .Bar1Value = 0
'      .Bar1Caption = strType & IIf(strName <> vbNullString, ": " & strName, vbNullString)
'      .Bar1MaxValue = lngMaxValue
'
'      .OpenProgress
'
'    'End If
'
'  End With
'
'Exit Sub
'
'LocalErr:
'  mstrErrorMessage = Err.Description

End Sub

Public Property Get ErrorMessage() As String
  If mstrErrorMessage <> vbNullString Then
    ErrorMessage = mstrErrorMessage
  Else
    If Not (mobjOutputType Is Nothing) Then
      ErrorMessage = mobjOutputType.ErrorMessage
    End If
  End If
End Property



Public Function SettingOptions( _
      strSettingWordTemplate As String, _
      strSettingExcelTemplate As String, _
      blnSettingExcelGridlines As Boolean, _
      blnSettingExcelHeaders As Boolean, _
      blnSettingExcelOmitSpacerRow As Boolean, _
      blnSettingExcelOmitSpacerCol As Boolean, _
      blnSettingAutoFitCols As Boolean, _
      blnSettingLandscape As Boolean, _
      blnEmailSystemPermission As Boolean) As Boolean _
      

  gstrSettingWordTemplate = strSettingWordTemplate
  gstrSettingExcelTemplate = strSettingExcelTemplate
  gblnSettingExcelGridlines = blnSettingExcelGridlines
  gblnSettingExcelHeaders = blnSettingExcelHeaders
  gblnSettingExcelOmitSpacerRow = blnSettingExcelOmitSpacerRow
  gblnSettingExcelOmitSpacerCol = blnSettingExcelOmitSpacerCol
  gblnSettingAutoFitCols = blnSettingAutoFitCols
  gblnSettingLandscape = blnSettingLandscape
  gblnEmailSystemPermission = blnEmailSystemPermission

End Function


Public Function SettingLocations( _
      lngSettingTitleCol As Long, _
      lngSettingTitleRow As Long, _
      lngSettingDataCol As Long, _
      lngSettingDataRow As Long) As Boolean

  glngSettingTitleCol = lngSettingTitleCol
  glngSettingTitleRow = lngSettingTitleRow
  glngSettingDataCol = lngSettingDataCol
  glngSettingDataRow = lngSettingDataRow

End Function


Public Function SettingTitle( _
      blnSettingTitleGridlines As Boolean, _
      blnSettingTitleBold As Boolean, _
      blnSettingTitleUnderline As Boolean, _
      lngSettingTitleBackcolour As Long, _
      lngSettingTitleForecolour As Long, _
      lngSettingTitleBackcolour97 As Long, _
      lngSettingTitleForecolour97 As Long) As Boolean

  gblnSettingTitleGridlines = blnSettingTitleGridlines
  gblnSettingTitleBold = blnSettingTitleBold
  gblnSettingTitleUnderline = blnSettingTitleUnderline
  glngSettingTitleBackcolour = lngSettingTitleBackcolour
  glngSettingTitleForecolour = lngSettingTitleForecolour
  glngSettingTitleBackcolour97 = lngSettingTitleBackcolour97
  glngSettingTitleForecolour97 = lngSettingTitleForecolour97

End Function


Public Function SettingHeading( _
      blnSettingHeadingGridlines As Boolean, _
      blnSettingHeadingBold As Boolean, _
      blnSettingHeadingUnderline As Boolean, _
      lngSettingHeadingBackcolour As Long, _
      lngSettingHeadingForecolour As Long, _
      lngSettingHeadingBackcolour97 As Long, _
      lngSettingHeadingForecolour97 As Long) As Boolean

  gblnSettingHeadingGridlines = blnSettingHeadingGridlines
  gblnSettingHeadingBold = blnSettingHeadingBold
  gblnSettingHeadingUnderline = blnSettingHeadingUnderline
  glngSettingHeadingBackcolour = lngSettingHeadingBackcolour
  glngSettingHeadingForecolour = lngSettingHeadingForecolour
  glngSettingHeadingBackcolour97 = lngSettingHeadingBackcolour97
  glngSettingHeadingForecolour97 = lngSettingHeadingForecolour97

End Function


Public Function SettingData( _
      blnSettingDataGridlines As Boolean, _
      blnSettingDataBold As Boolean, _
      blnSettingDataUnderline As Boolean, _
      lngSettingDataBackcolour As Long, _
      lngSettingDataForecolour As Long, _
      lngSettingDataBackcolour97 As Long, _
      lngSettingDataForecolour97 As Long) As Boolean

  gblnSettingDataGridlines = blnSettingDataGridlines
  gblnSettingDataBold = blnSettingDataBold
  gblnSettingDataUnderline = blnSettingDataUnderline
  glngSettingDataBackcolour = lngSettingDataBackcolour
  glngSettingDataForecolour = lngSettingDataForecolour
  glngSettingDataBackcolour97 = lngSettingDataBackcolour97
  glngSettingDataForecolour97 = lngSettingDataForecolour97

End Function


Public Function ArrayDim(lngCol As Long, lngRow As Long) As Boolean
  ReDim mstrArray(lngCol, lngRow) As String
End Function

Public Function ArrayReDim() As Boolean
  ReDim Preserve mstrArray(UBound(mstrArray, 1), UBound(mstrArray, 2) + 1) As String
End Function

Public Function ArrayAddTo(lngCol As Long, lngRow As Long, strInput) As Boolean
  mstrArray(lngCol, lngRow) = strInput
End Function

Public Property Get UserName() As String
  UserName = gsUserName
End Property

Public Property Let UserName(ByVal strNewValue As String)
  gsUserName = strNewValue
End Property

'Public Property Get Format() As OutputFormats
'  Format = mlngFormat
'End Property
'
'Public Property Let Format(lngNewValue As OutputFormats)
'  mlngFormat = Format
'End Property

Public Property Get PageTitles() As Boolean
  PageTitles = mblnPageTitles
End Property

Public Property Let PageTitles(ByVal blnNewValue As Boolean)
  mblnPageTitles = blnNewValue
End Property

Public Property Get EmailAttachAs() As String
  EmailAttachAs = mstrEmailAttachAs
End Property

Public Function KillFile(strFilename As String) As Boolean

  On Local Error GoTo LocalErr

  Kill strFilename
  KillFile = True

Exit Function

LocalErr:
  mstrErrorMessage = "Error overwriting file '" & strFilename & "'" & _
      IIf(Err.Description <> vbNullString, vbCrLf & "(" & Err.Description & ")", vbNullString)
  'mstrErrorMessage = "Cannot access read-only document '" & Mid(strFileName, InStrRev(strFileName, "\") + 1) & "'."
  KillFile = False

End Function


Public Function GetSequentialNumberedFile(ByVal strFilename As String) As String

  Dim lngFound As Long
  Dim lngCount As Long

  lngCount = 2
  lngFound = InStrRev(strFilename, ".")
  Do While Dir(Left(strFilename, lngFound - 1) & "(" & CStr(lngCount) & ")" & Mid(strFilename, lngFound)) <> vbNullString
    lngCount = lngCount + 1
  Loop
  GetSequentialNumberedFile = Left(strFilename, lngFound - 1) & "(" & CStr(lngCount) & ")" & Mid(strFilename, lngFound)

End Function
      

Public Property Get PrintPrompt() As Boolean

  'Dim blnPromptConfig As Boolean
  'blnPromptConfig = (GetPCSetting("Printer", "Prompt", False) = True)
  'PrintPrompt = (blnPromptConfig And Not gblnBatchMode)

End Property

Public Property Let PivotSuppressBlanks(strNewValue As Boolean)
  mblnPivotSuppressBlanks = strNewValue
End Property

Public Property Get PivotSuppressBlanks() As Boolean
  PivotSuppressBlanks = mblnPivotSuppressBlanks
End Property

Public Property Let PivotDataFunction(ByVal strNewValue As String)
  mstrFunction = strNewValue
End Property

Public Property Get PivotDataFunction() As String
  If mstrFunction = vbNullString Then mstrFunction = "Total"
  PivotDataFunction = mstrFunction
End Property

Public Property Get IndicatorColumn() As Boolean
  IndicatorColumn = mblnIndicatorColumn
End Property

Public Property Let IndicatorColumn(ByVal blnNewValue As Boolean)
  mblnIndicatorColumn = blnNewValue
End Property

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

Public Property Get SaveAsValues() As String
  SaveAsValues = mstrSaveAsValues
End Property

Public Property Let SaveAsValues(strValue As String)
  mstrSaveAsValues = strValue
End Property

Public Function GetSaveAsFormat(strFilename As String) As String
  GetSaveAsFormat = GetSaveAsFormat2(strFilename, mstrSaveAsValues)
End Function
