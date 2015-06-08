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
Attribute VB_Exposed = False
Option Explicit

Private mobjOutputType As Object
Private mfrmOutput As frmOutputOptions
Private mcolStyles As Collection
Private mcolMerges As Collection
Private mcolColumns As Collection
Private mstrErrorMessage As String

Private mlngFormat As OutputFormats
Private mstrFunction As String
Private mlngEmailGroupID As Long
Private mstrEmailSubject As String
Private mstrEmailAttachAs As String
Private mblnPageRange As Boolean
Private mblnPrintData As Boolean
Private mstrPrinterName As String
Private mblnPageTitles As Boolean

Private mstrDefaultPrinter As String
Private mblnSizeColumnsIndependently As Boolean
Private mlngHeaderRows As Long
Private mlngHeaderCols As Long
Private mblnPivotSuppressBlanks As Boolean

Private mblnData As Boolean
Private mblnCSV As Boolean
Private mblnHTML As Boolean
Private mblnWord As Boolean
Private mblnExcel As Boolean
Private mblnChart As Boolean
Private mblnPivot As Boolean

Private mblnUserCancelled As Boolean
Private mblnIndicatorColumn As Boolean
Private mblnUpdateProgress As Boolean

Private Function ConvertToGrid(pstrArray() As String, pGrdTemp As SSDBGrid) As SSDBGrid

  Dim strAddLine As String
  Dim icol As Long
  Dim iRow As Long
    
  With pGrdTemp
    'initialise grid
    .RemoveAll
    .Columns.RemoveAll
    .ColumnHeaders = False
      
    'add the columns to the grid
    For icol = 0 To UBound(pstrArray) Step 1
      .Columns.Add icol
    Next icol
    
    For iRow = 0 To UBound(pstrArray, 2) Step 1
        
      strAddLine = vbNullString
       
      For icol = 0 To UBound(pstrArray) Step 1
        strAddLine = strAddLine & pstrArray(icol, iRow)
        
        If icol < UBound(pstrArray) Then
          strAddLine = strAddLine & vbTab
        End If
      Next icol
        
      .AddItem strAddLine
      
    Next iRow
    
  End With
  
  If mblnSizeColumnsIndependently Then
    Set ConvertToGrid = AutoSizeGridColumms(pGrdTemp)
  Else
    Set ConvertToGrid = pGrdTemp
  End If
  
End Function

Private Function AutoSizeGridColumms(pSSDBGrid As SSDBGrid) As SSDBGrid
  
  'formats the size of the columns so all data is shown

  Dim icol As Long
  Dim iRow As Long
  Dim lngTextWidth As Long
  Dim varBookmark As Variant

  Const TextWidthMultiplier = 80
  
  lngTextWidth = 0
  
  With pSSDBGrid
    .Redraw = False
    .MoveFirst
    For icol = 0 To .Columns.Count - 1 Step 1
      
      lngTextWidth = ((Len(Trim(.Columns(icol).CellText(varBookmark)))) * TextWidthMultiplier)
      
      If .Columns(icol).Visible Then
        For iRow = 0 To .Rows - 1 Step 1
          varBookmark = .AddItemBookmark(iRow)
          
          If ((Len(Trim(.Columns(icol).CellText(varBookmark)))) * TextWidthMultiplier) > lngTextWidth Then
            lngTextWidth = ((Len(Trim(.Columns(icol).CellText(varBookmark)))) * TextWidthMultiplier)
          End If
        Next iRow
        
        .Columns(icol).Width = lngTextWidth + 150
      End If
      lngTextWidth = 0
    Next icol
    .Redraw = True
  End With
  
  Set AutoSizeGridColumms = pSSDBGrid
  
End Function


Public Property Let UserCancelled(pblnNewValue As Boolean)
  mblnUserCancelled = pblnNewValue
End Property

Public Property Get UserCancelled() As Boolean
  UserCancelled = (mblnUserCancelled Or gobjProgress.Cancelled)
End Property
Private Sub Class_Initialize()

  Set mfrmOutput = New frmOutputOptions
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

  InitialiseStyles
  Set mcolMerges = New Collection
  
  On Local Error Resume Next
  'MH20030922 Fault 6124 (Q257688)
  Printer.TrackDefault = True
  mstrDefaultPrinter = Printer.DeviceName

End Sub

Private Sub InitialiseStyles()

  Dim objStyle As clsOutputStyle
  
  Set mcolStyles = New Collection
  
  Set objStyle = New clsOutputStyle
  With objStyle
    .Name = "Title"
    .StartCol = Val(GetUserSetting("Output", "TitleCol", "3"))
    .StartRow = Val(GetUserSetting("Output", "TitleRow", "2"))
    .Gridlines = (GetUserSetting("Output", "TitleGridLines", "0") = "1")
    .Bold = (GetUserSetting("Output", "TitleBold", "1") = "1")
    .Underline = (GetUserSetting("Output", "TitleUnderline", "1") = "1")
    .BackCol = Val(GetUserSetting("Output", "TitleBackcolour", vbWhite))
    .ForeCol = Val(GetUserSetting("Output", "TitleForecolour", GetColour("Midnight Blue")))
  End With

  mcolStyles.Add objStyle, objStyle.Name
  
  Set objStyle = New clsOutputStyle
  With objStyle
    .Name = "Heading"
    .Gridlines = (GetUserSetting("Output", "HeadingGridLines", "1") = "1")
    .Bold = (GetUserSetting("Output", "HeadingBold", "1") = "1")
    .Underline = (GetUserSetting("Output", "HeadingUnderline", "0") = "1")
    .BackCol = Val(GetUserSetting("Output", "HeadingBackcolour", GetColour("Dolphin Blue")))
    .ForeCol = Val(GetUserSetting("Output", "HeadingForecolour", GetColour("Midnight Blue")))
    .CenterText = True
  End With

  mcolStyles.Add objStyle, objStyle.Name
  
  Set objStyle = New clsOutputStyle
  With objStyle
    .Name = "HeadingCols"
    .StartCol = 0
    .StartRow = 0
    .Gridlines = (GetUserSetting("Output", "HeadingGridLines", "1") = "1")
    .Bold = (GetUserSetting("Output", "HeadingBold", "1") = "1")
    .Underline = (GetUserSetting("Output", "HeadingUnderline", "0") = "1")
    .BackCol = Val(GetUserSetting("Output", "HeadingBackcolour", GetColour("Dolphin Blue")))
    .ForeCol = Val(GetUserSetting("Output", "HeadingForecolour", GetColour("Midnight Blue")))
  End With

  mcolStyles.Add objStyle, objStyle.Name

  Set objStyle = New clsOutputStyle
  With objStyle
    .Name = "Data"
    .StartCol = Val(GetUserSetting("Output", "DataCol", "2"))
    .StartRow = Val(GetUserSetting("Output", "DataRow", "4"))
    .Gridlines = (GetUserSetting("Output", "DataGridLines", "1") = "1")
    .Bold = (GetUserSetting("Output", "DataBold", "0") = "1")
    .Underline = (GetUserSetting("Output", "DataUnderline", "0") = "1")
    .BackCol = Val(GetUserSetting("Output", "DataBackcolour", GetColour("Pale Grey")))
    .ForeCol = Val(GetUserSetting("Output", "DataForecolour", GetColour("Midnight Blue")))
  End With

  mcolStyles.Add objStyle, objStyle.Name

'  'New StyleSet for Report Pack
'  Set objStyle = New clsOutputStyle
'  With objStyle
'    .Name = "Heading2"
'    .StartCol = Val(GetUserSetting("Output", "TitleCol", "3"))
'    .StartRow = Val(GetUserSetting("Output", "TitleRow", "2"))
'    .Gridlines = (GetUserSetting("Output", "TitleGridLines", "0") = "1")
'    .Bold = (GetUserSetting("Output", "TitleBold", "1") = "1")
'    .Underline = (GetUserSetting("Output", "TitleUnderline", "1") = "1")
'    .BackCol = Val(GetUserSetting("Output", "TitleBackcolour", vbWhite))
'    .ForeCol = Val(GetUserSetting("Output", "TitleForecolour", GetColour("Midnight Blue")))
'  End With
'
'  mcolStyles.Add objStyle, objStyle.Name


  Set objStyle = Nothing

End Sub


Public Function AddColumn(strHeading As String, lngDataType As Long, lngDecimals As Long, Optional blnThousandSeparator As Boolean) As Boolean

  Dim objColumn As clsColumn

  On Local Error GoTo LocalErr
  AddColumn = True
  
  Set objColumn = New clsColumn
  objColumn.Heading = strHeading
  objColumn.DataType = lngDataType
  objColumn.DecPlaces = lngDecimals
  objColumn.ThousandSeparator = blnThousandSeparator

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
  Optional blnGridlines As Variant) As Boolean

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

    If Not IsMissing(blnGridlines) Then
      .Gridlines = blnGridlines
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
  Set mfrmOutput = Nothing
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
  
Public Function AddPage(strDefTitle As String, Optional mstrSheetName As String, Optional mblnOutputCoverSheet As Boolean)
  On Local Error Resume Next
  
    mobjOutputType.AddPage strDefTitle, mstrSheetName, mcolStyles, mblnOutputCoverSheet

End Function

Public Property Let SizeColumnsIndependently(ByVal blnNewValue As Boolean)
  mblnSizeColumnsIndependently = blnNewValue
End Property

Public Function DataArray(strArray() As String)
  'Not all classes support all properties so catch any errors...
  On Local Error Resume Next
  
  Err.Clear
  If UBound(strArray, 2) < 0 Then
    mstrErrorMessage = "No data to output."
    Exit Function
  End If

  mobjOutputType.SizeColumnsIndependently = mblnSizeColumnsIndependently
  mobjOutputType.HeaderCols = mlngHeaderCols
  mobjOutputType.HeaderRows = mlngHeaderRows

  'On Local Error GoTo 0
  mobjOutputType.DataArray strArray(), mcolColumns, mcolStyles, mcolMerges

End Function

Public Function DataArrayToGrid(strArray() As String, pgrdNew As SSDBGrid)
  
  'Not all classes support all properties so catch any errors...
  On Local Error Resume Next
  mobjOutputType.SizeColumnsIndependently = mblnSizeColumnsIndependently
  mobjOutputType.HeaderCols = mlngHeaderCols
  mobjOutputType.HeaderRows = mlngHeaderRows

  'On Local Error GoTo 0
  If mblnPrintData Then
    DataGrid ConvertToGrid(strArray(), pgrdNew)
  Else
    mobjOutputType.DataArray strArray(), mcolColumns, mcolStyles, mcolMerges
  End If

End Function

Public Function RecordProfilePage(pfrmRecProfile As frmRecordProfilePreview, _
  piPageNumber As Integer) As Boolean
  'Not all classes support all properties so catch any errors...
  On Local Error Resume Next
'''  mobjOutputType.SizeColumnsIndependently = mblnSizeColumnsIndependently
'''  mobjOutputType.HeaderCols = mlngHeaderCols
'''  mobjOutputType.HeaderRows = mlngHeaderRows

  On Local Error GoTo 0

'''  mobjOutputType.DataArray strArray(), mcolColumns, mcolStyles, mcolMerges
  RecordProfilePage = mobjOutputType.RecordProfilePage(pfrmRecProfile, piPageNumber, mcolStyles)
  
End Function

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

Public Function DataGrid(objNewValue As SSDBGrid)

  Dim strArray() As String
  Dim lngGridGrp As Long
  Dim lngGridCol As Long
  Dim lngGridRow As Long
  Dim lngGrp As Long
  Dim lngCol As Long
  Dim lngRow As Long
  Dim lngVisibleCols As Long
  Dim blnGroupHeaders As Boolean
  Dim blnIgnoreRow As Boolean
  Dim varBookmark As Variant
  
  On Local Error GoTo LocalErr

  If mblnPrintData Then
    mobjOutputType.DataGrid objNewValue
    Exit Function
  End If
  
  
  With objNewValue

    If .Rows = 0 Or .Cols = 0 Then
      mstrErrorMessage = "No data in grid to output."
      Exit Function
    End If


    ResetMerges
    blnGroupHeaders = (.GroupHeaders And .Groups.Count > 0)

    'Get count of visible columns
    lngVisibleCols = -1
    For lngGridCol = 0 To .Cols - 1
      If .Columns(lngGridCol).Visible Then
        lngVisibleCols = lngVisibleCols + 1
        
        'Check if this is a header column...
        If .Columns(lngGridCol).ButtonsAlways Then
          If mlngHeaderCols = lngVisibleCols - 1 Then
            mlngHeaderCols = lngVisibleCols
          End If
        End If
      End If
    Next
    
    
    'GROUP HEADERS
    lngCol = 0
    lngRow = -1
    If blnGroupHeaders Then
      lngRow = lngRow + 1
      ReDim Preserve strArray(lngVisibleCols, lngRow)
      
      mlngHeaderRows = 2
      For lngGridGrp = 0 To .Groups.Count - 1
        If .Groups(lngGridGrp).Visible Then
          strArray(lngCol, lngRow) = .Groups(lngGridGrp).Caption
          AddMerge lngCol, 0, lngCol + .Groups(lngGridGrp).Columns.Count - 1, 0
          lngCol = lngCol + .Groups(lngGridGrp).Columns.Count
        End If
      Next
    End If
    
    'COLUMN HEADERS
    lngCol = 0
    lngRow = lngRow + 1
    ReDim Preserve strArray(lngVisibleCols, lngRow)
    For lngGridCol = 0 To .Cols - 1
      If .Columns(lngGridCol).Visible Then
        strArray(lngCol, lngRow) = .Columns(lngGridCol).Caption
        lngCol = lngCol + 1
      End If
    Next

    ' Remember the grid settings before we start scrolling through
    varBookmark = .FirstRow
    .Redraw = False
    .MoveFirst

    'DATA ROWS
    For lngGridRow = 0 To .Rows - 1
      
      blnIgnoreRow = False
      If mblnIndicatorColumn Then
        blnIgnoreRow = (Left(.Columns(0).CellText(.AddItemBookmark(lngGridRow)), 1) = "*")
      End If

      If Not blnIgnoreRow Then
        
        lngCol = 0
        lngRow = lngRow + 1
        If lngRow > UBound(strArray, 2) Then ReDim Preserve strArray(lngVisibleCols, lngRow + 100)

        For lngGridCol = 0 To .Cols - 1
          If .Columns(lngGridCol).Visible Then
            'strArray(lngCol, lngRow) = .Columns(lngGridCol).CellText(.AddItemBookmark(lngGridRow))
            strArray(lngCol, lngRow) = .Columns(lngGridCol).Text
            lngCol = lngCol + 1
          End If
        Next

      End If
      
      .MoveNext
      
    Next

     ' Reset the grid to how we found it
    .Bookmark = varBookmark
    .Redraw = True

  End With
  
  ReDim Preserve strArray(lngVisibleCols, lngRow)

  DataArray strArray()

Exit Function

LocalErr:
  mstrErrorMessage = Err.Description

End Function

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


Public Function SendEmail(strAttachment As String) As Boolean

  Dim objOutputEmail As clsOutputEMail
  Dim rsEmail As Recordset
  Dim strSQL As String
  
  Dim strErrorTemp As String
  Dim strAddress() As String
  Dim lngCount As Long
  Dim blnShowMessage As Boolean
  Dim blnProgressVisible As Boolean

  Dim strTo As String

  On Error GoTo LocalErr
  
  blnShowMessage = False
  
  If datGeneral.SystemPermission("EMAILGROUPS", "VIEW") = False Then
    mstrErrorMessage = "You do not have permission to use email groups."
    SendEmail = False
    Exit Function
  End If

  Screen.MousePointer = vbHourglass

      strSQL = "SELECT ASRSysEmailGroupItems.*," & _
               " ASRSysEmailAddress.Name as 'AddrName', ASRSysEmailAddress.Fixed as 'AddrFixed'" & _
               " FROM ASRSysEmailGroupItems" & _
               " JOIN ASRSysEmailAddress ON ASRSysEmailGroupItems.EmailDefID = ASRSysEmailAddress.EmailID" & _
               " WHERE EmailGroupID = " & CStr(mlngEmailGroupID) & _
               " ORDER BY AddrName"
      Set rsEmail = datGeneral.GetReadOnlyRecords(strSQL)

      If rsEmail.BOF And rsEmail.EOF Then
        mstrErrorMessage = "Error retrieving email recipient(s)"
        Exit Function
      End If

      blnProgressVisible = gobjProgress.Visible
      gobjProgress.Visible = False
      
      strTo = vbNullString
      strErrorTemp = "Error resolving email recipient(s)"
      Do While Not rsEmail.EOF
        strAddress = Split(rsEmail!AddrFixed, ";")
        
        For lngCount = 0 To UBound(strAddress)
          If Trim(strAddress(lngCount)) <> vbNullString Then
            strTo = strTo & IIf(strTo <> vbNullString, ";", "") & strAddress(lngCount)
          End If
        Next
        
        rsEmail.MoveNext
      Loop
      strErrorTemp = vbNullString
      
      rsEmail.Close
      Set rsEmail = Nothing
      Set objOutputEmail = New clsOutputEMail
      mstrErrorMessage = objOutputEmail.SendEmailFromClient(strTo, "", "", mstrEmailSubject, "", strAttachment, False)
      Set objOutputEmail = Nothing

      gobjProgress.Visible = blnProgressVisible
  SendEmail = True
Exit Function

LocalErr:
    If strErrorTemp = "" Then strErrorTemp = "Error sending email"
    mstrErrorMessage = strErrorTemp & _
      IIf(Err.Description <> vbNullString, " (" & Err.Description & ")", vbNullString)
    On Error Resume Next
    SendEmail = False
End Function

Public Function SendReportPackEmail(strAttachment As String) As Boolean
  Dim objOutputEmail As clsOutputEMail
  Dim rsEmail As Recordset
  Dim strSQL As String
  Dim strErrorTemp As String
  Dim strAddress() As String
  Dim lngCount As Long
  Dim blnShowMessage As Boolean
  Dim blnProgressVisible As Boolean

  Dim strTo As String

  On Error GoTo LocalErr
  blnShowMessage = False

  If datGeneral.SystemPermission("EMAILGROUPS", "VIEW") = False Then
    mstrErrorMessage = "You do not have permission to use email groups."
    SendReportPackEmail = False
    Exit Function
  End If

  Screen.MousePointer = vbHourglass
      strSQL = "SELECT ASRSysEmailGroupItems.*," & _
               " ASRSysEmailAddress.Name as 'AddrName', ASRSysEmailAddress.Fixed as 'AddrFixed'" & _
               " FROM ASRSysEmailGroupItems" & _
               " JOIN ASRSysEmailAddress ON ASRSysEmailGroupItems.EmailDefID = ASRSysEmailAddress.EmailID" & _
               " WHERE EmailGroupID = " & CStr(mlngEmailGroupID) & _
               " ORDER BY AddrName"
      Set rsEmail = datGeneral.GetReadOnlyRecords(strSQL)

      If rsEmail.BOF And rsEmail.EOF Then
        mstrErrorMessage = "Error retrieving email recipient(s)"
        Exit Function
      End If

      blnProgressVisible = gobjProgress.Visible
      gobjProgress.Visible = False
      
      strTo = vbNullString
      strErrorTemp = "Error resolving email recipient(s)"
      Do While Not rsEmail.EOF
        strAddress = Split(rsEmail!AddrFixed, ";")
        
        For lngCount = 0 To UBound(strAddress)
          If Trim(strAddress(lngCount)) <> vbNullString Then
            strTo = strTo & IIf(strTo <> vbNullString, ";", "") & strAddress(lngCount)
          End If
        Next
        
        rsEmail.MoveNext
      Loop
      strErrorTemp = vbNullString
      
      rsEmail.Close
      Set rsEmail = Nothing

      Set objOutputEmail = New clsOutputEMail
      mstrErrorMessage = objOutputEmail.SendEmailFromClient(strTo, "", "", mstrEmailSubject, "", strAttachment, False)
      Set objOutputEmail = Nothing

      gobjProgress.Visible = blnProgressVisible

  SendReportPackEmail = True

Exit Function

LocalErr:
    If strErrorTemp = "" Then strErrorTemp = "Error sending email"
    mstrErrorMessage = strErrorTemp & _
      IIf(Err.Description <> vbNullString, " (" & Err.Description & ")", vbNullString)
    On Error Resume Next
    SendReportPackEmail = False
End Function

Public Function GetTempFileName(strFileName As String) As String

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
  If strFileName <> vbNullString Then
    strFileName = Left(strTempFileName, InStrRev(strTempFileName, "\")) & _
                  Mid(strFileName, InStrRev(strFileName, "\") + 1)
  Else
    strFileName = strTempFileName
  End If

  If Dir(strFileName) <> vbNullString Then
    Kill strFileName
  End If

  GetTempFileName = strFileName

Exit Function

LocalErr:
  mstrErrorMessage = Err.Description

End Function

Public Function GetTempIntFileName(strFileName As String) As String
' NPG20090305 Fault 13531
' Generate temporary filename in 'Temporary Internet Files' folder
  Dim strTempFileName As String

  On Local Error GoTo LocalErr

  'Get temp path
  strTempFileName = Space(1024)
  ' Call GetTempPath(1024, strTempFileName)
  strTempFileName = GetTmpInternetFName
  If InStr(strTempFileName, Chr(0)) > 0 Then
    strTempFileName = Left(strTempFileName, InStr(strTempFileName, Chr(0)) - 1)
  End If

  'temp path + "\" + file name
  If strFileName <> vbNullString Then
    strFileName = Left(strTempFileName, InStrRev(strTempFileName, "\")) & _
                  Mid(strFileName, InStrRev(strFileName, "\") + 1)
  Else
    strFileName = strTempFileName
  End If

  If Dir(strFileName) <> vbNullString Then
    Kill strFileName
  End If

  GetTempIntFileName = strFileName

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


Public Property Get cboPageBreak() As ComboBox
  Set cboPageBreak = mfrmOutput.cboPageBreak
  mblnPageRange = True
End Property


Public Function SetOptions( _
    blnPrompt As Boolean, _
    lngFormat As Long, _
    blnScreen As Boolean, _
    blnPrinter As Boolean, _
    strPrinterName As String, _
    blnSave As Boolean, _
    lngSaveExisting As Long, _
    blnEmail As Boolean, _
    lngEmailAddr As Long, _
    strEmailSubject As String, _
    strEmailAttachAs As String, _
    strFileName As String, _
    Optional blnCustomReportsPrintFilterHeader As Boolean, _
    Optional blnPreviewOnScreen As Boolean, _
    Optional strOutputTitlePage As String, _
    Optional strOutputReportPackTitle As String, _
    Optional strOutputOverrideFilter As String, _
    Optional blnOutputTOC As Boolean, _
    Optional blnOutputCoverSheet As Boolean, _
    Optional lngOverrideFilterID As Long, _
    Optional blnOutputRetainPivotOrChart As Boolean, _
    Optional lngOriginalOutputFormat As Long)

  Dim objColumn As clsColumn
  Dim blnCancelled As Boolean
  'Dim lngSaveExisting As Long
  Dim strEmailFileExtension As String

  On Local Error GoTo LocalErr

  blnCancelled = False
  mlngEmailGroupID = lngEmailAddr
  mstrEmailSubject = strEmailSubject
  mstrEmailAttachAs = strEmailAttachAs

  If blnPrompt Then
    With mfrmOutput

      .optOutputFormat(lngFormat).value = True
      .chkDestination(desScreen).value = IIf(blnScreen And .chkDestination(desScreen).Enabled, vbChecked, vbUnchecked)

      .chkDestination(desPrinter).value = IIf(blnPrinter, vbChecked, vbUnchecked)
      SetComboText .cboPrinterName, strPrinterName
      If strPrinterName <> vbNullString Then
        If .cboPrinterName.Text <> strPrinterName Then
          .cboPrinterName.AddItem strPrinterName
          .cboPrinterName.ListIndex = .cboPrinterName.NewIndex
          COAMsgBox "This definition is set to output to printer " & strPrinterName & _
                 " which is not set up on your PC.", vbInformation, .Caption
        End If
      End If

      .chkDestination(desSave).value = IIf(blnSave, vbChecked, vbUnchecked)
      If blnSave Then
        SetComboItem .cboSaveExisting, lngSaveExisting
      End If

      .chkDestination(desEmail).value = IIf(blnEmail, vbChecked, vbUnchecked)
      If blnEmail Then
        .txtEmailGroup.Text = datGeneral.GetEmailGroupName(lngEmailAddr)
        .txtEmailGroup.Tag = lngEmailAddr
        .txtEmailSubject.Text = strEmailSubject
        .txtEmailAttachAs.Text = strEmailAttachAs
      End If
      .txtFilename.Text = strFileName
      .ShowFormats mblnData, mblnCSV, mblnHTML, mblnWord, mblnExcel, mblnChart, mblnPivot
      .PageRange = mblnPageRange
      .Show vbModal
      If .Cancelled Then
        mblnUserCancelled = True
        SetOptions = False
        Exit Function
      End If

      lngFormat = mfrmOutput.Format
      blnScreen = (.chkDestination(desScreen).value = vbChecked)
      blnPrinter = (.chkDestination(desPrinter).value = vbChecked)
      strPrinterName = .cboPrinterName.Text
      blnSave = (.chkDestination(desSave).value = vbChecked)

      blnEmail = (.chkDestination(desEmail).value = vbChecked)
      lngEmailAddr = Val(.txtEmailGroup.Tag)
      strEmailSubject = .txtEmailSubject.Text
      strEmailAttachAs = .txtEmailAttachAs.Text
      strFileName = .txtFilename.Text

      mlngEmailGroupID = .txtEmailGroup.Tag
      mstrEmailSubject = .txtEmailSubject.Text
      mstrEmailAttachAs = .txtEmailAttachAs.Text

      If .cboSaveExisting.ListIndex <> -1 Then
        lngSaveExisting = .cboSaveExisting.ItemData(.cboSaveExisting.ListIndex)
      End If

      Select Case mlngFormat
        Case fmtCSV, fmtFixedLengthFile
          CheckEmailAttachment "csv"
        Case fmtHTML
          CheckEmailAttachment "htm"
        Case fmtWordDoc
          CheckEmailAttachment strEmailFileExtension
        Case fmtExcelWorksheet, fmtExcelchart, fmtExcelPivotTable
          CheckEmailAttachment strEmailFileExtension
      End Select

      blnCancelled = mfrmOutput.Cancelled
    End With
  End If

  If Not ValidPrinter(strPrinterName) Then
    mstrErrorMessage = _
        "This definition is set to output to printer " & strPrinterName & _
        " which is not set up on your PC."
    SetOptions = False
    Exit Function
  End If

  mlngFormat = lngFormat
  
  If gblnReportPackMode And mlngFormat = fmtExcelWorksheet Then
    'Default should be excel if not set
    Select Case lngOriginalOutputFormat
      Case fmtExcelPivotTable
        mlngFormat = fmtExcelPivotTable
      Case fmtExcelchart
        mlngFormat = fmtExcelchart
    End Select
  End If
  
  Select Case mlngFormat
    Case fmtDataOnly
      If Not blnPrinter And Not blnSave And Not blnEmail Then
        mstrErrorMessage = "Invalid output destination for batch mode." & vbCrLf & _
                           "You must select a different output format or additional output destination."
        SetOptions = False
        Exit Function
      End If
      Set mobjOutputType = New clsOutputGrid
    Case fmtCSV, fmtFixedLengthFile
      Set mobjOutputType = New clsOutputCSV
    Case fmtHTML
      Set mobjOutputType = New clsOutputHTML
    Case fmtWordDoc
      Set mobjOutputType = New clsOutputWord
    Case fmtExcelWorksheet
      Set mobjOutputType = New clsOutputExcel
    Case fmtExcelchart
      Set mobjOutputType = New clsOutputExcel
      mobjOutputType.Chart = True
    Case fmtExcelPivotTable
      Set mobjOutputType = New clsOutputExcel
      mobjOutputType.PivotTable = True
  End Select

  mobjOutputType.Parent = Me
  mobjOutputType.Screen = blnScreen
  mobjOutputType.DestPrinter = blnPrinter
  mobjOutputType.Save = blnSave
  mobjOutputType.SaveExisting = lngSaveExisting
  mobjOutputType.Email = blnEmail
  mobjOutputType.FileName = strFileName

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

  SetOptions = Not blnCancelled

Exit Function

LocalErr:
  If mstrErrorMessage = vbNullString Then
    mstrErrorMessage = Err.Description
  End If

End Function


Public Sub SetPrinter()

  Dim objDefPrinter As cSetDfltPrinter

  If gstrDefaultPrinterName = vbNullString Then
    If Printer.DeviceName <> vbNullString Then
      gstrDefaultPrinterName = Printer.DeviceName
    End If
  End If


  If mstrPrinterName <> "<Default Printer>" Then
    If Printer.DeviceName <> mstrPrinterName Then
      Set objDefPrinter = New cSetDfltPrinter
      objDefPrinter.SetPrinterAsDefault mstrPrinterName
      Set objDefPrinter = Nothing
    End If
  End If

End Sub

Public Sub ResetDefaultPrinter()

  Dim objDefPrinter As cSetDfltPrinter
  Dim lngTimeOut As Long

  lngTimeOut = Timer + 10

  If mstrPrinterName <> "<Default Printer>" Then
    Set objDefPrinter = New cSetDfltPrinter
    If gstrDefaultPrinterName <> vbNullString Then
      Do While Printer.DeviceName <> gstrDefaultPrinterName 'And lngTimeOut > Timer
        objDefPrinter.SetPrinterAsDefault gstrDefaultPrinterName
      Loop
    End If
    Set objDefPrinter = Nothing
  End If

End Sub

Public Property Let FileDelimiter(ByVal strNewValue As String)
  'clsOutputCSV only
  On Local Error Resume Next
  mobjOutputType.FileDelimiter = strNewValue
End Property

Public Property Let DisableDelimiterCheck(ByVal blnNewValue As Boolean)
  'clsOutputCSV only
  On Local Error Resume Next
  mobjOutputType.DisableDelimiterCheck = blnNewValue
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

  On Local Error GoTo LocalErr

  With gobjProgress
    
    'If .Visible = False Then
    
      Select Case mlngFormat
      Case fmtExcelchart, fmtExcelPivotTable, fmtExcelWorksheet
        '.AviFile = App.Path & "\videos\excel.avi"
        .AVI = dbExcel
        .MainCaption = "Output to Excel"
      Case fmtWordDoc
        '.AviFile = App.Path & "\videos\word.avi"
        .AVI = dbWord
        .MainCaption = "Output to Word"
      Case fmtHTML
        '.AviFile = App.Path & "\videos\internet.avi"
        .AVI = dbInternet
        .MainCaption = "Output to HTML"
      Case Else
        '.AviFile = App.Path & "\videos\report.avi"
        .AVI = dbText
        .MainCaption = "Output to File"
      End Select
  
      .NumberOfBars = 1
      .Time = False
      .Cancel = True
      .Bar1MaxValue = lngMaxValue
      
      .Caption = strType
      .Bar1Value = 0
      .Bar1Caption = strType & IIf(strName <> vbNullString, ": " & strName, vbNullString)
      .Bar1MaxValue = lngMaxValue
  
      .OpenProgress

    'End If

  End With

Exit Sub

LocalErr:
  mstrErrorMessage = Err.Description

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

Public Property Get Format() As OutputFormats
  Format = mlngFormat
End Property

Public Property Let Format(lngNewValue As OutputFormats)
  mlngFormat = Format
End Property

Public Property Get PageTitles() As Boolean
  PageTitles = mblnPageTitles
End Property

Public Property Let PageTitles(ByVal blnNewValue As Boolean)
  mblnPageTitles = blnNewValue
End Property

Public Property Get EmailAttachAs() As String
  EmailAttachAs = mstrEmailAttachAs
End Property

Private Function CheckEmailAttachment(strExt As String)
  
  Dim lngFound As Long
  
  If mstrEmailAttachAs <> vbNullString Then
    lngFound = InStrRev(mstrEmailAttachAs, ".")
    If lngFound = 0 Then
      mstrEmailAttachAs = mstrEmailAttachAs & "." & strExt
    End If
  End If

End Function

Public Function KillFile(strFileName As String) As Boolean

  On Local Error GoTo LocalErr

  Kill strFileName
  KillFile = True

Exit Function

LocalErr:
  mstrErrorMessage = "Error overwriting file '" & strFileName & "'" & _
      IIf(Err.Description <> vbNullString, vbCrLf & "(" & Err.Description & ")", vbNullString)
  'mstrErrorMessage = "Cannot access read-only document '" & Mid(strFileName, InStrRev(strFileName, "\") + 1) & "'."
  KillFile = False

End Function


Public Function GetSequentialNumberedFile(ByVal strFileName As String) As String

  Dim lngFound As Long
  Dim lngCount As Long

  lngCount = 2
  
  ' NPG20081201 Fault 13340
  If InStrRev(strFileName, "\") > InStrRev(strFileName, ".") Then ' No filename extension specified
    Do While Dir(strFileName & "(" & CStr(lngCount) & ")") <> vbNullString
      lngCount = lngCount + 1
    Loop
    
    GetSequentialNumberedFile = strFileName & "(" & CStr(lngCount) & ")"
  Else  ' Specified filename has an extension
    lngFound = InStrRev(strFileName, ".")
    Do While Dir(Left(strFileName, lngFound - 1) & "(" & CStr(lngCount) & ")" & Mid(strFileName, lngFound)) <> vbNullString
      lngCount = lngCount + 1
    Loop
    
    GetSequentialNumberedFile = Left(strFileName, lngFound - 1) & "(" & CStr(lngCount) & ")" & Mid(strFileName, lngFound)
  End If
  

End Function
      

Public Property Get PrintPrompt() As Boolean

  Dim blnPromptConfig As Boolean
  blnPromptConfig = (GetPCSetting("Printer", "Prompt", True) = True)
  PrintPrompt = (blnPromptConfig And Not gblnBatchMode)

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


Public Property Let UpdateProgressPerRow(blnNewValue As Boolean)
  mblnUpdateProgress = blnNewValue
End Property

Public Property Get UpdateProgressPerRow() As Boolean
  UpdateProgressPerRow = mblnUpdateProgress
End Property

