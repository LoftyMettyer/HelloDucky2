VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOutputGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mgrdPrintGrid As SSDBGrid
Attribute mgrdPrintGrid.VB_VarHelpID = -1
Private mobjPrintGrid As clsPrintGrid
Private mobjParent As clsOutputRun

Private mstrDefTitle As String
Private mstrErrorMessage As String
Private mlngPageCount As Long

Private mblnScreen As Boolean
Private mblnPrinter As Boolean
Private mstrPrinterName As String
Private mblnSave As Boolean
Private mlngSaveExisting As Long
Private mblnEmail As Boolean
Private mstrFileName As String

Public Sub ClearUp()
  Set mgrdPrintGrid = Nothing
  Set mobjPrintGrid = Nothing
End Sub

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

Public Property Let FileName(strFileName As String)
  mstrFileName = strFileName
End Property

Private Sub mgrdPrintGrid_PrintInitialize(ByVal ssPrintInfo As SSDataWidgets_B.ssPrintInfo)
  Call mobjPrintGrid.PrintInitialise(ssPrintInfo)
End Sub

Public Function AddPage(strDefTitle As String, mstrSheetName As String, colStyles As Collection)
  mstrDefTitle = strDefTitle
  mlngPageCount = mlngPageCount + 1
End Function

Public Sub DataGrid(objNewValue As SSDBGrid)

  Dim strDefaultPrinter As String
  
  Set mgrdPrintGrid = objNewValue
  
  If mstrErrorMessage <> vbNullString Then
    Exit Sub
  End If

  If mlngPageCount = 0 Or mlngPageCount = 1 Then
    Set mobjPrintGrid = New clsPrintGrid
    mobjParent.SetPrinter
  End If

  mobjPrintGrid.Heading = mstrDefTitle
  mobjPrintGrid.Grid = mgrdPrintGrid
  mobjPrintGrid.SuppressPrompt = (mlngPageCount > 1 Or Not mobjParent.PrintPrompt)

  If Not gobjProgress.Cancelled Then
    mobjPrintGrid.PrintGrid False
  End If

  If mobjPrintGrid.Cancelled Or gobjProgress.Cancelled Then
    mstrErrorMessage = "Cancelled by user."
    mobjParent.UserCancelled = True
  End If

End Sub

Public Sub Complete()
  mobjParent.ResetDefaultPrinter
End Sub

Public Function RecordProfilePage(pfrmRecProfile As frmRecordProfilePreview, _
  piPageNumber As Integer, _
  pcolStyles As Collection) As Boolean
  
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim strDefaultPrinter As String
  
  fOK = True
  
  If piPageNumber = 1 Then
    Set mobjPrintGrid = New clsPrintGrid
  
    mobjParent.SetPrinter
    
    mobjPrintGrid.Heading = mstrDefTitle
  End If
  mobjPrintGrid.SuppressPrompt = (piPageNumber > 1 Or Not mobjParent.PrintPrompt)
  
  fOK = mobjPrintGrid.PrintRecordProfilePage(pfrmRecProfile, piPageNumber)

  If fOK Then
    If Not gobjProgress.Cancelled Then
      fOK = Not mobjPrintGrid.Cancelled
    End If
  End If
  
  'JPD 20050810 Fault 10167
  If mobjPrintGrid.Cancelled Or gobjProgress.Cancelled Then
    mstrErrorMessage = "Cancelled by user."
    mobjParent.UserCancelled = True
  End If
  
  'JPD 20050524 Fault 10027
  'mobjParent.ResetDefaultPrinter
  
TidyUpAndExit:
  'Set mobjPrintGrid = Nothing
  RecordProfilePage = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Public Property Let Parent(objNewValue As clsOutputRun)
  Set mobjParent = objNewValue
End Property

Public Property Get ErrorMessage() As String
  ErrorMessage = mstrErrorMessage
End Property

