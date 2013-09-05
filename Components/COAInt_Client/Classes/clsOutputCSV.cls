VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOutputCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mobjParent As clsOutputRun

Private mblnScreen As Boolean
Private mblnPrinter As Boolean
Private mstrPrinterName As String
Private mblnSave As Boolean
Private mlngSaveExisting As Long
Private mblnEmail As Boolean
Private mstrFileName As String
Private mstrDelim As String
Private mblnQuotes As Boolean

Private mstrErrorMessage As String

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


Public Function GetFile(objParent As clsOutputRun, colStyles As Collection) As Boolean

  Dim strTempFileName As String
  Dim lngFound As Long
  Dim lngCount As Long

  On Local Error GoTo LocalErr


  If mstrFileName = vbNullString Then
    mstrFileName = objParent.GetTempFileName(mobjParent.EmailAttachAs)
    If Dir(mstrFileName) <> vbNullString Then
      objParent.KillFile mstrFileName
    End If
  
  
  ElseIf Dir(mstrFileName) <> vbNullString Then
    'Check if file already exists...

    Select Case mlngSaveExisting
    Case 0  'Overwrite
      Kill mstrFileName

    Case 1  'Do not overwrite (fail)
      mstrErrorMessage = "File already exists"
    
    Case 2  'Add Sequential number to file
      mstrFileName = mobjParent.GetSequentialNumberedFile(mstrFileName)

    
    Case 3  'Append to existing file
    
    
    Case 4  'Create new worksheet within existing workbook...
      'N/A (EXCEL ONLY)
    
    End Select

  End If

  GetFile = (mstrErrorMessage = vbNullString)

Exit Function

LocalErr:
  GetFile = False
  mstrErrorMessage = Err.Description

End Function


'Public Sub DataGrid(objNewValue As SSDBGrid, colColumns As Collection, colStyles As Collection)
'
'  Dim strOutput As String
'  Dim lngGridCol As Long
'  Dim lngGridRow As Long
'
'  Open mstrFileName For Append As #1
'
'  With objNewValue
'
'    .Redraw = False
'
'    'If not appending to existing file then add headers...
'    If LOF(1) = 0 Then
'      strOutput = vbNullString
'      For lngGridCol = 0 To .Cols - 1
'        If .Columns(lngGridCol).Visible Then
'          strOutput = _
'            IIf(strOutput <> vbNullString, strOutput & ",", "") & _
'            .Columns(lngGridCol).Caption
'        End If
'      Next
'      Print #1, strOutput
'    End If
'
'    For lngGridRow = 0 To .Rows - 1
'      strOutput = vbNullString
'      .Bookmark = .AddItemBookmark(lngGridRow)
'      For lngGridCol = 0 To .Cols - 1
'        If .Columns(lngGridCol).Visible Then
'          strOutput = _
'            IIf(strOutput <> vbNullString, strOutput & ",", "") & _
'            .Columns(lngGridCol).CellText(.Bookmark)
'        End If
'      Next
'      Print #1, strOutput
'    Next
'
'    .Redraw = True
'  End With
'
'  Close
'
'End Sub


Public Sub DataArray(strArray() As String, colColumns As Collection, colStyles As Collection, colMerges As Collection)

  Dim objColumn As clsColumn
  Dim strOutput As String
  Dim strTemp As String
  Dim lngGridCol As Long
  Dim lngGridRow As Long

  On Local Error GoTo LocalErr

  Open mstrFileName For Append As #1

  For lngGridRow = 0 To UBound(strArray, 2)
    strOutput = vbNullString
    For lngGridCol = 0 To UBound(strArray, 1)
      
      strTemp = strArray(lngGridCol, lngGridRow)

      If mblnQuotes Then
        If InStr(strTemp, ",") > 0 Or InStr(strTemp, Chr(34)) Then
          strTemp = Chr(34) & Replace(strTemp, String(1, 34), String(2, 34)) & Chr(34)
        End If
      End If

      strOutput = IIf(lngGridCol > 0, strOutput & mstrDelim, "") & strTemp

'      If gobjProgress.Visible And gobjProgress.Cancelled Then
'        mstrErrorMessage = "Cancelled by User"
'        Close
'        Exit Sub
'      End If

    Next
    
    Print #1, strOutput
  Next

  Close

Exit Sub

LocalErr:
  mstrErrorMessage = Err.Description

End Sub


Public Sub Complete()

  On Local Error GoTo LocalErr
  
  If mstrErrorMessage <> vbNullString Then
    Exit Sub
  End If
  
  'EMAIL
  If mblnEmail Then
    mstrErrorMessage = "Error sending email"
    mobjParent.SendEmail mstrFileName
  End If

  mstrErrorMessage = vbNullString

Exit Sub

LocalErr:
  mstrErrorMessage = mstrErrorMessage & _
    IIf(Err.Description <> vbNullString, " (" & Err.Description & ")", vbNullString)

End Sub

Private Sub Class_Initialize()
  mstrDelim = ","
  mblnQuotes = True
End Sub

Public Property Let FileDelimiter(ByVal strNewValue As String)
  mstrDelim = strNewValue
End Property

Public Property Let EncloseInQuotes(ByVal blnNewValue As Boolean)
  mblnQuotes = blnNewValue
End Property

Public Property Let Parent(objNewValue As clsOutputRun)
  Set mobjParent = objNewValue
End Property

Public Property Get ErrorMessage() As String
  ErrorMessage = mstrErrorMessage
End Property

Private Sub Class_Terminate()
  Close
End Sub

Public Sub ClearUp()
  Close
End Sub

