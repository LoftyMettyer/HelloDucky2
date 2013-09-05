VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOutputHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mobjParent As clsOutputRun

Private mstrDefTitle As String
Private mblnScreen As Boolean
Private mblnPrinter As Boolean
Private mstrPrinterName As String
Private mblnSave As Boolean
Private mlngSaveExisting As Long
Private mblnEmail As Boolean
Private mstrFileName As String
Private mlngHeaderRows As Long
Private mlngHeaderCols As Long
Private mblnHeaderVertical As Boolean
Private mblnApplyStyles As Boolean

'Private mstrHTMLTemplate As String

Private mstrHTMLOutput As String
Private mlngPageCount As Long
Private mstrErrorMessage As String

Public Sub ClearUp()
  Close
End Sub

'Public Function RecordProfilePage(pfrmRecProfile As Form, _
'  piPageNumber As Integer, _
'  pcolStyles As Collection)
'  ' Output the record profile page to Excel.
'
'  On Error GoTo ErrorTrap
'  gobjErrorStack.PushStack "clsOutputHTML.RecordProfilePage()"
'
'  Dim fOK As Boolean
'  Dim iLoop As Integer
'  Dim iLoop2 As Integer
'  Dim iLoop3 As Integer
'  Dim iLoop4 As Integer
'  Dim ctlTemp As Control
'  Dim varBookmark As Variant
'  Dim fGridPreceded As Boolean
'  Dim fGridFollowed As Boolean
'  Dim sTemp As String
'  Dim fPhotoDone As Boolean
'  Dim objRecProfTable As clsRecordProfileTabDtl
'  Dim sTempName As String
'  Dim iTemp As Integer
'  Dim fIsHeading As Boolean
'  Dim sSubFolderPath As String
'
'  Const RECPROFFOLLOWONCORRECTION = 10
'
'  Const COLUMN_ISHEADING = "IsHeading"
'  Const COLUMN_ISPHOTO = "IsPhoto"
'  Const PHOTOSTYLESET = "PhotoSS_"
'
'  fOK = True
'
'  sSubFolderPath = mstrFileName
'  If InStrRev(sSubFolderPath, ".") > 0 Then
'    sSubFolderPath = Left(sSubFolderPath, InStrRev(sSubFolderPath, ".") - 1)
'  End If
'  sSubFolderPath = sSubFolderPath & "_files"
'
'  For Each ctlTemp In pfrmRecProfile.Controls
'    If ctlTemp.Container Is pfrmRecProfile.picOutput(piPageNumber) Then
'      '
'      ' LABEL control
'      '
'      If TypeOf ctlTemp Is Label Then
'        If ctlTemp.Visible Then
'          mstrHTMLOutput = mstrHTMLOutput & _
'            "<BR>" & _
'            HTMLText("SPAN", ctlTemp.Caption, pcolStyles("Title")) & _
'            "<BR>"
'        End If
'      End If
'
'      '
'      ' GRID control
'      '
'      If TypeOf ctlTemp Is SSDBGrid Then
'        Set objRecProfTable = pfrmRecProfile.Definition.Item(ctlTemp.Tag)
'
'        ' Check if this grid is preceded or followed IMMEDIATELY by other grids.
'        ' ie. if this grid is part of a group of grids that are used to display
'        ' data (including pictures) vertically.
'        ' NB. Grids have only one row height value. To display pictures with their
'        ' own row height, we actually put them in their own grid, and position this grid
'        ' IMMEDIATELY after the normal data grid. Subsequent data is put in its own grid
'        ' IMMEDIATELY after the picture's grid.
'        ' This is what's meant by 'following' and 'preceding' grids.
'        fGridFollowed = False
'        fGridPreceded = False
'
'        If ctlTemp.Index > 1 Then
'          If (ctlTemp.Container = pfrmRecProfile.grdOutput(ctlTemp.Index - 1).Container) And _
'            (ctlTemp.Top = (pfrmRecProfile.grdOutput(ctlTemp.Index - 1).Top + pfrmRecProfile.grdOutput(ctlTemp.Index - 1).Height - RECPROFFOLLOWONCORRECTION)) Then
'
'            fGridPreceded = True
'          End If
'        End If
'
'        If ctlTemp.Index < pfrmRecProfile.grdOutput.Count - 1 Then
'          If (ctlTemp.Container = pfrmRecProfile.grdOutput(ctlTemp.Index + 1).Container) And _
'            ((ctlTemp.Top + ctlTemp.Height - RECPROFFOLLOWONCORRECTION) = pfrmRecProfile.grdOutput(ctlTemp.Index + 1).Top) Then
'
'            fGridFollowed = True
'          End If
'        End If
'
'        If Not fGridPreceded Then
'          mstrHTMLOutput = mstrHTMLOutput & _
'              "<CENTER><TABLE border=1 cellspacing=0 cellpadding=1" & _
'              " bordercolordark=" & HexColour(pcolStyles("Data").BackCol) & _
'              " bordercolorlight=000000>" & vbCrLf
'        End If
'
'        ' Send the group headers to the HTML document.
'        If (ctlTemp.GroupHeaders) And (ctlTemp.ColumnHeaders) Then
'          mstrHTMLOutput = mstrHTMLOutput & "<TR>"
'          For iLoop = 0 To ctlTemp.Groups.Count - 1
'            If ctlTemp.Groups(iLoop).Visible Then
'              iTemp = 0
'              For iLoop2 = 0 To ctlTemp.Groups(iLoop).Columns.Count - 1
'                If (ctlTemp.Columns(iLoop).Visible) Then
'                  iTemp = iTemp + 1
'                End If
'              Next iLoop2
'
'              ' Send the group header to the HTML document.
'              mstrHTMLOutput = mstrHTMLOutput & _
'                  HTMLText("TD", ctlTemp.Groups(iLoop).Caption, pcolStyles("Heading"), IIf(iTemp > 1, " COLSPAN=" & CStr(iTemp), ""))
'            End If
'          Next iLoop
'          mstrHTMLOutput = mstrHTMLOutput & _
'              "</TR>" & vbCrLf
'        End If
'
'        ' Send the column headers to the HTML document.
'        If (ctlTemp.ColumnHeaders) Then
'          mstrHTMLOutput = mstrHTMLOutput & "<TR>"
'          For iLoop = 0 To ctlTemp.Columns.Count - 1
'            If (ctlTemp.Columns(iLoop).Visible) Then
'              ' Send the column header to the HTML document.
'              mstrHTMLOutput = mstrHTMLOutput & _
'                  HTMLText("TD", ctlTemp.Columns(iLoop).Caption, pcolStyles("Heading"), vbNullString)
'            End If
'          Next iLoop
'          mstrHTMLOutput = mstrHTMLOutput & "</TR>" & vbCrLf
'        End If
'
'        ' Send data rows and columns to Excel.
'        For iLoop = 0 To ctlTemp.Rows - 1
'          varBookmark = ctlTemp.AddItemBookmark(iLoop)
'
'          mstrHTMLOutput = mstrHTMLOutput & "<TR>"
'
'          For iLoop2 = 0 To ctlTemp.Columns.Count - 1
'            If ctlTemp.Columns(iLoop2).Visible Then
'              ' Send the text or picture to the HTML document.
'              fPhotoDone = False
'              If (ctlTemp.TagVariant = COLUMN_ISPHOTO) And _
'                (ctlTemp.Columns(iLoop2).Style <> 4) Then
'
'                For iLoop3 = 0 To ctlTemp.Columns.Count - 1
'                  If ctlTemp.Columns(iLoop3).Visible Then
'                    sTemp = PHOTOSTYLESET & CStr(iLoop3 + 1)
'
'                    For iLoop4 = 0 To ctlTemp.StyleSets.Count - 1
'                      If ctlTemp.StyleSets(iLoop4).Name = sTemp Then
'                        sTempName = GetTmpFNameInFolder(sSubFolderPath)
'
'                        SavePicture ctlTemp.StyleSets(iLoop4).Picture, sTempName
'                        mstrHTMLOutput = mstrHTMLOutput & _
'                              "<TD" & _
'                              " bgcolor=" & HexColour(pcolStyles("Data").BackCol) & ">" & _
'                              "<IMG alt="""" src=""file://" & sTempName & """>" & _
'                              "</TD>"
'
'                        fPhotoDone = True
'                        Exit For
'                      End If
'                    Next iLoop4
'                  End If
'
'                  If fPhotoDone Then
'                    Exit For
'                  End If
'                Next iLoop3
'              End If
'
'              If (Not fPhotoDone) And _
'                ctlTemp.Columns(iLoop2).TagVariant = COLUMN_ISPHOTO Then
'
'                sTemp = PHOTOSTYLESET & CStr(iLoop2 + 1) & "_" & ctlTemp.Columns(CStr(objRecProfTable.IDPosition)).Value
'
'                For iLoop4 = 0 To ctlTemp.StyleSets.Count - 1
'                  If ctlTemp.StyleSets(iLoop4).Name = sTemp Then
'                    sTempName = GetTmpFNameInFolder(sSubFolderPath)
'                    SavePicture ctlTemp.StyleSets(iLoop4).Picture, sTempName
'                    mstrHTMLOutput = mstrHTMLOutput & _
'                          "<TD" & _
'                          " bgcolor=" & HexColour(pcolStyles("Data").BackCol) & ">" & _
'                          "<IMG alt="""" src=""file://" & sTempName & """>" & _
'                          "</TD>"
'                    fPhotoDone = True
'                    Exit For
'                  End If
'                Next iLoop4
'              End If
'
'              If Not fPhotoDone Then
'                ' Send the data to the HTML document.
'                varBookmark = ctlTemp.AddItemBookmark(iLoop)
'
'                fIsHeading = ((ctlTemp.Columns(iLoop2).Style = 4) And _
'                  (ctlTemp.Columns(iLoop2).TagVariant <> COLUMN_ISPHOTO)) Or _
'                  (ctlTemp.Columns(iLoop2).StyleSet = "Separator")
'                If Not ctlTemp.ColumnHeaders Then
'                  If ctlTemp.Columns(COLUMN_ISHEADING).CellText(varBookmark) = "1" Then
'                    fIsHeading = True
'                  End If
'                End If
'
'                mstrHTMLOutput = mstrHTMLOutput & _
'                    HTMLText("TD", ctlTemp.Columns(iLoop2).CellText(varBookmark), IIf(fIsHeading, pcolStyles("HeadingCols"), pcolStyles("Data")), vbNullString)
'              End If
'
'            End If
'          Next iLoop2
'
'          mstrHTMLOutput = mstrHTMLOutput & "</TR>" & vbCrLf
'        Next iLoop
'
'        If Not fGridFollowed Then
'          mstrHTMLOutput = mstrHTMLOutput & _
'              "</TABLE></CENTER><BR>" & vbCrLf
'        End If
'      End If
'    End If
'  Next ctlTemp
'  Set ctlTemp = Nothing
'
'TidyUpAndExit:
'  gobjErrorStack.PopStack
'  RecordProfilePage = fOK
'  Exit Function
'
'ErrorTrap:
'  gobjErrorStack.HandleError
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function

Public Property Let Screen(blnScreen As Boolean)
  mblnScreen = blnScreen
End Property

Public Property Let DestPrinter(blnPrinter As Boolean)
  mblnPrinter = blnPrinter
End Property

Public Property Let PrinterName(strPrinterName As String)
  mstrPrinterName = strPrinterName
End Property

Public Property Let Save(blnSave As Boolean)
  mblnSave = blnSave
End Property

Public Property Let SaveExisting(lngSaveExisting As Long)
  mlngSaveExisting = lngSaveExisting
End Property

Public Property Get SaveExisting() As Long
  SaveExisting = mlngSaveExisting
End Property

Public Property Let Email(blnEmail As Boolean)
  mblnEmail = blnEmail
End Property

Public Property Let FileName(strFilename As String)
  mstrFileName = strFilename
End Property

Public Function GetFile(objParent As clsOutputRun, colSytles As Collection) As Boolean

  Dim strTempFileName As String
  Dim lngFound As Long
  Dim lngCount As Long
  
  Dim strLineInput As String
  Dim blnAppending As Boolean
  Dim blnFound As Boolean
  
  On Local Error GoTo LocalErr
  
  
  blnAppending = False
  
  'Just in case we are emailing but not saving...
  If Not mblnSave Then
    If mblnEmail Then
      mstrFileName = objParent.GetTempFileName(mobjParent.EmailAttachAs)
      If Dir(mstrFileName) <> vbNullString Then
        objParent.KillFile mstrFileName
      End If
    End If
  
  Else
  
    'Check if file already exists...
    If Dir(mstrFileName) <> vbNullString And mstrFileName <> vbNullString Then
  
      Select Case mlngSaveExisting
      Case 0  'Overwrite
        If Not objParent.KillFile(mstrFileName) Then
          GetFile = False
          Exit Function
        End If
  
      Case 1  'Do not overwrite (fail)
        mstrErrorMessage = "File already exists."
        GetFile = False
        Exit Function

      Case 2  'Add Sequential number to file
        mstrFileName = mobjParent.GetSequentialNumberedFile(mstrFileName)

      Case 3  'Append to existing file


      Case 4  'Create new worksheet within existing workbook...
        'N/A (EXCEL ONLY)

      End Select

    End If

    If OpenFile Then
      blnAppending = ((LOF(1) > 0) And (mlngSaveExisting = 3))
    End If

  End If

'  If mstrHTMLTemplate <> vbNullString Then
'    If Dir(mstrHTMLTemplate) <> vbNullString Then
'      Open mstrHTMLTemplate For Input As #2
'
'      blnFound = False
'      Do While Not blnFound And Not EOF(2)
'        Input #2, strLineInput
'        blnFound = (LCase(strLineInput) = "<hrprodata>")
'        If Not blnFound Then
'          mstrHTMLOutput = mstrHTMLOutput & strLineInput
'        End If
'      Loop
'    End If
'
'  Else
    If Not blnAppending Then
      mstrHTMLOutput = mstrHTMLOutput & _
          "<HTML><BODY>"
    End If
    mstrHTMLOutput = mstrHTMLOutput & _
        "<FONT face=Verdana size=2>"

'  End If

  GetFile = (mstrErrorMessage = vbNullString)

Exit Function

LocalErr:
  mstrErrorMessage = Err.Description
  GetFile = False

End Function


Public Sub AddPage(strDefTitle As String, mstrSheetName As String, colStyles As Collection)

  Dim strTitle As String

  On Local Error GoTo LocalErr

  mstrHTMLOutput = mstrHTMLOutput & "<BR>"
  
  mlngPageCount = mlngPageCount + 1
  If mlngPageCount = 1 Then
    mstrDefTitle = strDefTitle
    strTitle = HTMLText("SPAN", strDefTitle, colStyles("Title"))
    mstrHTMLOutput = mstrHTMLOutput & _
        "<CENTER>" & strTitle & "</CENTER>"
  Else
    mstrHTMLOutput = mstrHTMLOutput & "<HR>" & vbCrLf
  End If

  If mstrSheetName <> vbNullString And mobjParent.PageTitles Then
    mstrHTMLOutput = mstrHTMLOutput & _
        "<BR>" & HTMLText("SPAN", mstrSheetName, colStyles("Title")) & "<BR>"
  End If

  mstrHTMLOutput = mstrHTMLOutput & "<BR>"

Exit Sub

LocalErr:
  mstrErrorMessage = Err.Description

End Sub


Public Sub DataArray(strArray() As String, colColumns As Collection, colStyles As Collection, colMerges As Collection)

  Dim objColumn As clsColumn
  Dim strOutput As String
  Dim lngGridCol As Long
  Dim lngGridRow As Long
  
  On Local Error GoTo LocalErr
  
  
  With colStyles("Title")
    .StartCol = glngSettingTitleCol
    .StartRow = glngSettingTitleRow
    .EndCol = .StartCol
    .EndRow = .StartRow
  End With

  With colStyles("Heading")
    .StartCol = 0
    .StartRow = 0
    .EndCol = UBound(strArray, 1)
    .EndRow = mlngHeaderRows - 1
  End With

  If mlngHeaderCols > 0 Then
    With colStyles("HeadingCols")
      .StartCol = 0
      .StartRow = 0
      .EndCol = mlngHeaderCols - 1
      .EndRow = UBound(strArray, 2)
    End With
  End If

  With colStyles("Data")
    .StartCol = mlngHeaderCols
    .StartRow = mlngHeaderRows
    .EndCol = UBound(strArray, 1)
    .EndRow = UBound(strArray, 2)
  End With
  
    
  mstrHTMLOutput = mstrHTMLOutput & _
      "<CENTER><TABLE border=1 cellspacing=0 cellpadding=1" & _
      " bordercolordark=" & HexColour(colStyles("Data").BackCol) & _
      " bordercolorlight=000000>" & vbCrLf

  'strOutput = vbNullString
  'lngGridCol = 0
  'For Each objColumn In colColumns
  '  strOutput = strOutput & _
  '      HTMLText("TD", objColumn.Heading, colStyles("Heading"))
  '  lngGridCol = lngGridCol + 1
  'Next
  'mstrHTMLOutput = mstrHTMLOutput & "<TR>" & strOutput & "</TR>" & vbCrLf
  
  For lngGridRow = 0 To UBound(strArray, 2)
    strOutput = vbNullString
    For lngGridCol = 0 To UBound(strArray, 1)
      strOutput = strOutput & _
          CheckHTMLText("TD", strArray(lngGridCol, lngGridRow), lngGridCol, lngGridRow, colStyles, colMerges, colColumns)
      
'      If gobjProgress.Visible And gobjProgress.Cancelled Then
'        mstrErrorMessage = "Cancelled by user."
'        Exit Sub
'      End If
    
    Next
    mstrHTMLOutput = mstrHTMLOutput & _
        "<TR>" & strOutput & "</TR>" & vbCrLf
  Next

  mstrHTMLOutput = mstrHTMLOutput & _
      "</TABLE></CENTER>" & vbCrLf

Exit Sub

LocalErr:
  mstrErrorMessage = Err.Description

End Sub


Private Function OpenFile() As Boolean
  
  On Local Error GoTo LocalErr
  
  Open mstrFileName For Append As #1

  OpenFile = True

Exit Function

LocalErr:
  mstrErrorMessage = "Error saving file <" & mstrFileName & ">" & _
    IIf(Err.Description <> vbNullString, vbCrLf & " (" & Err.Description & ")", vbNullString)
  OpenFile = False

End Function


Private Function CheckHTMLText(strTag As String, strInput As String, lngCol As Long, lngRow As Long, colStyles As Collection, colMerges As Collection, colColumns As Collection) As String

  Dim objStyle As clsOutputStyle
  Dim objMerge As clsOutputStyle
  Dim objTemp As clsOutputStyle
  Dim objColumn As clsColumn
  Dim strTemp As String

  On Local Error GoTo LocalErr
  
  
  For Each objTemp In colStyles
    If (objTemp.StartCol <= lngCol And objTemp.EndCol >= lngCol) And _
       (objTemp.StartRow <= lngRow And objTemp.EndRow >= lngRow) Then
        Set objStyle = objTemp
    End If
  Next

  For Each objTemp In colMerges
    If (objTemp.StartCol <= lngCol And objTemp.EndCol >= lngCol) And _
       (objTemp.StartRow <= lngRow And objTemp.EndRow >= lngRow) Then
        Set objMerge = objTemp
    End If
  Next

  
  Select Case colColumns(lngCol + 1).DataType
  Case sqlNumeric, sqlInteger
    strTemp = " ALIGN=Right "
  Case sqlBoolean
    strTemp = " ALIGN=Center "
  Case Else
    strTemp = " ALIGN=Left "
  End Select

  If objMerge Is Nothing Then
    'No merging required...
    CheckHTMLText = HTMLText(strTag, strInput, objStyle, strTemp)

  Else
    If objMerge.StartCol = lngCol And objMerge.StartRow = lngRow Then
      'This is the top left of a merged range...
      strTemp = strTemp & _
                " COLSPAN=" & CStr(objMerge.EndCol - objMerge.StartCol + 1) & _
                " ROWSPAN=" & CStr(objMerge.EndRow - objMerge.StartRow + 1) & _
                " VALIGN=Top"
      CheckHTMLText = HTMLText(strTag, strInput, objStyle, strTemp)
    Else
      'part of a range so don't bother...
      CheckHTMLText = vbNullString
    End If
  End If

Exit Function

LocalErr:
  mstrErrorMessage = Err.Description

End Function


Private Function HTMLText(strTag As String, strInput As String, objStyle As clsOutputStyle, Optional strExtraTag As String) As String

  Dim strOutput As String

  On Local Error GoTo LocalErr

  strOutput = Replace(strInput, "<", "&LT;")
  strOutput = Replace(strOutput, ">", "&GT;")
  strOutput = Replace(strOutput, vbTab, "</TD><TD>")
  strOutput = Replace(strOutput, " ", "&nbsp;")

  If strOutput = vbNullString Then
    strOutput = "&nbsp;"
  End If

  If mblnApplyStyles And Not (objStyle Is Nothing) Then
    If objStyle.CenterText Then
      strOutput = "<CENTER>" & strOutput & "</CENTER>"
    End If
    If objStyle.Bold Then
      strOutput = "<B>" & strOutput & "</B>"
    End If
    If objStyle.Underline Then
      strOutput = "<U>" & strOutput & "</U>"
    End If

    strOutput = _
      "<" & strTag & strExtraTag & _
      " bgcolor=" & HexColour(objStyle.BackCol) & ">" & _
      "<FONT size=2 color=" & HexColour(objStyle.ForeCol) & ">" & _
      strOutput & _
      "</FONT>" & _
      "</" & strTag & ">"
  End If

  HTMLText = strOutput

Exit Function

LocalErr:
  mstrErrorMessage = Err.Description

End Function


Private Function HexColour(lngColour) As String
  HexColour = Right("0" & Hex(lngColour Mod 256), 2) & _
              Right("0" & Hex(lngColour \ 256), 2) & _
              Right("0" & Hex(lngColour \ 65536), 2)
End Function



Public Property Let HeaderRows(ByVal lngNewValue As Long)
  mlngHeaderRows = lngNewValue
End Property

Public Property Let HeaderCols(ByVal lngNewValue As Long)
  mlngHeaderCols = lngNewValue
End Property

Private Sub Class_Initialize()
  'mstrHTMLTemplate = GetUserSetting("Output", "HTMLTemplate", vbNullString)
  mblnApplyStyles = True
End Sub

Public Property Let Parent(objNewValue As clsOutputRun)
  Set mobjParent = objNewValue
End Property

Public Property Get ErrorMessage() As String
  ErrorMessage = mstrErrorMessage
End Property

Private Sub Class_Terminate()
  Close
End Sub

Public Sub Complete()

  Dim strLineInput As String
  Dim blnOK As Boolean


  If mstrErrorMessage <> vbNullString Then
    Close
    Exit Sub
  End If
  
  On Local Error GoTo LocalErr

  blnOK = True

'  If mstrHTMLTemplate <> vbNullString Then
'    If Dir(mstrHTMLTemplate) <> vbNullString Then
'      Do While Not EOF(2)
'        Input #2, strLineInput
'        mstrHTMLOutput = mstrHTMLOutput & strLineInput
'      Loop
'    End If
'
'  Else
    mstrHTMLOutput = mstrHTMLOutput & _
        "</CENTER><BR><HR>" & vbCrLf & _
        "Created on " & Format(Now, DateFormat & " hh:nn") & " by " & gsUserName & vbCrLf & _
        "</FONT></BODY></HTML>"
'  End If
  
  If mblnSave Then
    Print #1, mstrHTMLOutput
  End If
  
  Close


  'EMAIL
  If mblnEmail Then
    mstrErrorMessage = "Error sending email"
    
    If mblnSave Then
      mobjParent.SendEmail mstrFileName
    Else
      'mstrFileName = GetTmpFName
      'mstrFileName = Left(mstrFileName, Len(mstrFileName) - 3) & "htm"
      Open mstrFileName For Output As #1
      Print #1, mstrHTMLOutput
      Close
      mobjParent.SendEmail mstrFileName
      Kill mstrFileName
    End If

  End If


  If mblnScreen Then
    mstrErrorMessage = "Error displaying HTML"
    blnOK = DisplayInBrowser
  End If

  If blnOK Then
    mstrErrorMessage = vbNullString
  End If

TidyAndExit:

Exit Sub

LocalErr:
  mstrErrorMessage = mstrErrorMessage & _
    IIf(Err.Description <> vbNullString, " (" & Err.Description & ")", vbNullString)
  Resume TidyAndExit

End Sub

Private Function DisplayInBrowser() As Boolean

  Dim IE As SHDocVw.InternetExplorer
  Dim dblWait As Double
  Dim dblWait2 As Double
  Dim blnOK As Boolean

  On Error GoTo LocalErr

  blnOK = True
  dblWait = Timer + 10

  Set IE = New SHDocVw.InternetExplorer
  
  'JPD 20091007 Fault HRPRO-31, HRPRO-33, HRPRO-34
  ' New SHDocVw.InternetExplorer sometimes gets a handle on the existing SSI/DMI browser instance.
  ' If this happens, do it again toensure you get a fresh instance.
  If IE.LocationName = "OpenHR Self-service Intranet" _
    Or IE.LocationName = "OpenHR Intranet" Then
  
    Set IE = New SHDocVw.InternetExplorer
  End If
  
  If mblnSave Then
    IE.Navigate mstrFileName
    Do While IE.Busy
      DoEvents
    Loop

  Else
RetryDisplay:
    'AE20071129 Fault #12111 / #12112
'    IE.Navigate ""    'Creates a blank document
    IE.Navigate "about:blank", "_Blank"
    Do While IE.Busy
      DoEvents
    Loop

    'Keep trying for 10 seconds then error
    blnOK = False
    Do

      Err.Clear
      On Error GoTo LocalErr
      IE.AddressBar = False
      If Not IE.Document Is Nothing Then
        IE.Document.Title = mstrDefTitle
        If Not IE.Document.Body Is Nothing Then
          IE.Document.Body.InnerHtml = mstrHTMLOutput
          IE.Document.Body.Bgcolor = "white"
          blnOK = (Err.Number = 0)
        End If
      End If
      DoEvents

    Loop While Not blnOK

  End If
  
  If blnOK Then
    IE.Visible = True
  End If
  Set IE = Nothing

  DisplayInBrowser = blnOK

Exit Function

LocalErr:
  dblWait2 = Timer + 2
  Do While dblWait2 > Timer
    DoEvents
  Loop

  If dblWait > Timer Then
    Err.Clear
    On Error GoTo LocalErr
    GoTo RetryDisplay
  End If
  DisplayInBrowser = False
  Set IE = Nothing

End Function
