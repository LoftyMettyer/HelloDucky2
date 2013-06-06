VERSION 5.00
Begin VB.Form frmCustomReportsAddOrder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Sort Order"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1024
   Icon            =   "frmCustomReportsAddOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAction 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   2775
      TabIndex        =   7
      Top             =   2925
      Width           =   1200
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "&Cancel"
      Height          =   400
      Index           =   0
      Left            =   4020
      TabIndex        =   8
      Top             =   2925
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   2745
      Left            =   90
      TabIndex        =   9
      Top             =   45
      Width           =   5130
      Begin VB.OptionButton optDesc 
         Caption         =   "&Descending"
         Height          =   225
         Left            =   1110
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton optAsc 
         Caption         =   "&Ascending"
         Height          =   225
         Left            =   1110
         TabIndex        =   1
         Top             =   780
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.CheckBox chkSuppressRepeated 
         Caption         =   "&Suppress Repeated Values"
         Height          =   270
         Left            =   1110
         TabIndex        =   6
         Top             =   2325
         Width           =   2745
      End
      Begin VB.CheckBox chkValueOnChange 
         Caption         =   "&Value on Change"
         Height          =   270
         Left            =   1110
         TabIndex        =   5
         Top             =   2025
         Width           =   2130
      End
      Begin VB.CheckBox chkPageOnChange 
         Caption         =   "&Page on Change"
         Height          =   270
         Left            =   1110
         TabIndex        =   4
         Top             =   1725
         Width           =   2130
      End
      Begin VB.CheckBox chkBreakOnChange 
         Caption         =   "&Break on Change"
         Height          =   270
         Left            =   1110
         TabIndex        =   3
         Top             =   1425
         Width           =   2130
      End
      Begin VB.ComboBox cboColumns 
         Height          =   315
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   3870
      End
      Begin VB.Label lblOrder 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order :"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   780
         Width           =   525
      End
      Begin VB.Label lblColumns 
         BackStyle       =   0  'Transparent
         Caption         =   "Column :"
         Height          =   255
         Left            =   150
         TabIndex        =   10
         Top             =   360
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmCustomReportsAddOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'DataAccess Class
Private datData As DataMgr.clsDataAccess
Private mblnEditingExisting As Boolean
Private mfrmForm As frmCustomReports
Private mblnInvalidOnInitialise As Boolean
Private mbLoading As Boolean
Private mblnUserCancelled As Boolean


Private Sub CheckColumnOptions()

  Dim iLoop As Integer
  Dim objItem As ListItem
  Dim sKey As String
  Dim objColumm  As clsColumn

  sKey = cboColumns.ItemData(cboColumns.ListIndex)
  Set objColumm = mfrmForm.mcolCustomReportColDetails.Item("C" & sKey)
  If Not (objColumm Is Nothing) Then
    If objColumm.Hidden Then
      Me.chkValueOnChange.Value = vbUnchecked
      Me.chkValueOnChange.Enabled = False
    Else
      Me.chkValueOnChange.Enabled = True
    End If
  
    If objColumm.Hidden Or objColumm.Repetition Then
      Me.chkSuppressRepeated.Value = vbUnchecked
      Me.chkSuppressRepeated.Enabled = False
    Else
      Me.chkSuppressRepeated.Enabled = True
    End If
 
  Else
    Me.chkValueOnChange.Enabled = True
    Me.chkSuppressRepeated.Enabled = True
  End If

  Set objColumm = Nothing

End Sub

Private Sub SetBreaks(bBreak As Boolean, bPage As Boolean, BValue As Boolean, bSuppress As Boolean)

  'Function created to check that 'Break on Change' and 'Page on Change' are not
  'selected at the same time.
  
  If bBreak And bPage Then
    Me.chkBreakOnChange.Value = vbUnchecked
    Me.chkPageOnChange.Value = vbUnchecked
  ElseIf bBreak Then
    Me.chkBreakOnChange.Value = vbChecked
    Me.chkPageOnChange.Value = vbUnchecked
  ElseIf bPage Then
    Me.chkBreakOnChange.Value = vbUnchecked
    Me.chkPageOnChange.Value = vbChecked
  End If

  Me.chkValueOnChange.Value = IIf(BValue, vbChecked, vbUnchecked)
  Me.chkSuppressRepeated.Value = IIf(bSuppress, vbChecked, vbUnchecked)

End Sub

Public Function Initialise(sBaseTable As String, iCols As Integer, Optional blnEditing As Boolean, Optional pfrmForm As frmCustomReports) As Boolean

  ' Populate the combo box showing available columns then show the form
  Dim iLoop As Integer
  Dim sSelected As String
  
  mbLoading = True
  
  ' Set references to class modules
  Set datData = New DataMgr.clsDataAccess
    
  Set mfrmForm = pfrmForm
  
  mblnEditingExisting = blnEditing
  
  If mblnEditingExisting = True Then
  
    sSelected = mfrmForm.grdReportOrder.Columns("Column").Text
    Me.cboColumns.AddItem sSelected
    Me.cboColumns.ItemData(Me.cboColumns.NewIndex) = mfrmForm.grdReportOrder.Columns("ColumnID").Text
  
    If mfrmForm.grdReportOrder.Columns("Order").Text = "Asc" Then Me.optAsc.Value = True Else Me.optDesc.Value = True
    
  'TM20010821 Fault 2379
  'Now use the new SetBreaks() sub to define the values of the Break Check Boxes
  '  If mfrmForm.grdReportOrder.Columns("Break").Value Then Me.chkBreakOnChange.Value = 1
  '  If mfrmForm.grdReportOrder.Columns("Page").Value Then Me.chkPageOnChange.Value = 1
  '  If mfrmForm.grdReportOrder.Columns("Value").Value Then Me.chkValueOnChange.Value = 1
  '  If mfrmForm.grdReportOrder.Columns("Hide").Value Then Me.chkSuppressRepeated.Value = 1
    
    SetBreaks mfrmForm.grdReportOrder.Columns("Break").Value, mfrmForm.grdReportOrder.Columns("Page").Value, _
              mfrmForm.grdReportOrder.Columns("Value").Value, mfrmForm.grdReportOrder.Columns("Hide").Value
    
    If mfrmForm.grdReportOrder.Columns("Break").Value And mfrmForm.grdReportOrder.Columns("Page").Value Then
      mblnInvalidOnInitialise = True
    End If
    
  End If

  'AE20071005 Fault #10468
  mfrmForm.grdReportOrder.Redraw = False
  
  For iLoop = 1 To mfrmForm.ListView2.ListItems.Count
    If Left(mfrmForm.ListView2.ListItems(iLoop).Key, 1) = "C" Then
      
      If OkToAdd(mfrmForm.ListView2.ListItems(iLoop).Key) Then
        Me.cboColumns.AddItem mfrmForm.ListView2.ListItems(iLoop).Text
        Me.cboColumns.ItemData(Me.cboColumns.NewIndex) = Right(mfrmForm.ListView2.ListItems(iLoop).Key, Len(mfrmForm.ListView2.ListItems(iLoop).Key) - 1)
      End If
    
    End If
  Next iLoop
  
  'AE20071005 Fault #10468
  mfrmForm.grdReportOrder.Redraw = True
  
  If mblnEditingExisting = True Then SetComboText Me.cboColumns, sSelected
  
  If Me.cboColumns.ListCount = 0 Then
    If iCols = 0 Then
      COAMsgBox "You must add a column to the report before you can add to the sort order.", vbExclamation + vbOKOnly, "Custom Reports"
      Initialise = False
      Exit Function
    Else
      COAMsgBox "You must add more columns to the report before you can add to the sort order.", vbExclamation + vbOKOnly, "Custom Reports"
      Initialise = False
      Exit Function
    End If
  Else
    cboColumns.ListIndex = 0
    cboColumns.Enabled = (cboColumns.ListCount > 1)
    cboColumns.BackColor = IIf(cboColumns.Enabled, vbWindowBackground, vbButtonFace)
  End If

  CheckColumnOptions
  
  If mblnInvalidOnInitialise = True Then
    Me.chkBreakOnChange.Enabled = True
    Me.chkPageOnChange.Enabled = True
  End If
  
  Initialise = True
  mbLoading = False
  Me.Show vbModal

  
End Function

Private Sub cboColumns_Click()
  If Not mbLoading Then
    CheckColumnOptions
  End If
End Sub


Private Sub chkBreakOnChange_Click()

  If chkBreakOnChange.Value = vbChecked Then
    chkPageOnChange.Value = vbUnchecked ' vbChecked
    chkPageOnChange.Enabled = False
  Else
    chkPageOnChange.Enabled = True
  End If

End Sub

Private Sub chkPageOnChange_Click()

  If chkPageOnChange.Value = vbChecked Then
    chkBreakOnChange.Value = vbUnchecked ' vbChecked
    chkBreakOnChange.Enabled = False
  Else
    chkBreakOnChange.Enabled = True
  End If

End Sub


Private Sub cmdAction_Click(Index As Integer)

  Dim objColumn As clsColumn
 
 
  ' OK pressed
  If Index = 1 Then
    If mblnEditingExisting = True Then
      
      ' edit guff here
      mfrmForm.grdReportOrder.Columns("Column").Text = Me.cboColumns.Text
      mfrmForm.grdReportOrder.Columns("ColumnID").Text = Me.cboColumns.ItemData(Me.cboColumns.ListIndex)

      If Me.optAsc.Value = True Then mfrmForm.grdReportOrder.Columns("Order").Text = "Asc" Else mfrmForm.grdReportOrder.Columns("Order").Text = "Desc"
  
      If Me.chkBreakOnChange.Value = 1 Then
        mfrmForm.grdReportOrder.Columns("Break").Value = True
      Else
        mfrmForm.grdReportOrder.Columns("Break").Value = False
      End If
      If Me.chkPageOnChange.Value = 1 Then
        mfrmForm.grdReportOrder.Columns("Page").Value = True
      Else
        mfrmForm.grdReportOrder.Columns("Page").Value = False
      End If
      If Me.chkValueOnChange.Value = 1 Then
        mfrmForm.grdReportOrder.Columns("Value").Value = True
      Else
        mfrmForm.grdReportOrder.Columns("Value").Value = False
      End If
      If Me.chkSuppressRepeated.Value = 1 Then
        mfrmForm.grdReportOrder.Columns("Hide").Value = True
      Else
        mfrmForm.grdReportOrder.Columns("Hide").Value = False
      End If

    Else
      mfrmForm.grdReportOrder.MoveLast
      mfrmForm.grdReportOrder.AddItem _
        Me.cboColumns.ItemData(Me.cboColumns.ListIndex) & vbTab & _
        Me.cboColumns.Text & vbTab & _
        IIf(optAsc.Value, "Asc", "Desc") & vbTab & _
        IIf(Me.chkBreakOnChange.Value = vbChecked, True, False) & vbTab & _
        IIf(Me.chkPageOnChange.Value = vbChecked, True, False) & vbTab & _
        IIf(Me.chkValueOnChange.Value = vbChecked, True, False) & vbTab & _
        IIf(Me.chkSuppressRepeated.Value = vbChecked, True, False)
        
      mfrmForm.grdReportOrder.SelBookmarks.RemoveAll
      mfrmForm.grdReportOrder.MoveLast
      mfrmForm.grdReportOrder.SelBookmarks.Add mfrmForm.grdReportOrder.Bookmark
      
    End If
    
    Set objColumn = mfrmForm.mcolCustomReportColDetails.Item("C" & Me.cboColumns.ItemData(Me.cboColumns.ListIndex))
    objColumn.BreakOnChange = IIf(Me.chkBreakOnChange.Value = vbChecked, True, False)
    objColumn.PageOnChange = IIf(Me.chkPageOnChange.Value = vbChecked, True, False)
    objColumn.ValueOnChange = IIf(Me.chkValueOnChange.Value = vbChecked, True, False)
    objColumn.SurpressRepeatedValues = IIf(Me.chkSuppressRepeated.Value = vbChecked, True, False)
  'AE20071025 Fault #6797
  ElseIf Index = 0 Then
    mblnUserCancelled = True
  End If

  'AE20071025 Fault #6797
  'Unload Me
  'Set frmCustomReportsAddOrder = Nothing
  Me.Hide
  
  Set objColumn = Nothing
  
  

End Sub

Private Function OkToAdd(sKey As String) As Boolean

  ' Checks if the column has already been defined as a sort order.
  ' If so, do not add it to the combo box for selection (again).
  
  Dim pintLoop As Integer
  Dim pvarbookmark As Variant
  Dim plngPrevRow As Long
  
  ' RH 12/06/00 - FAULT 417
  
'  For iloop = 0 To frmCustomReports.grdReportOrder.Rows - 1
'    frmCustomReports.grdReportOrder.Row = iloop
'    If frmCustomReports.grdReportOrder.Columns(0).Text = Right(skey, Len(skey) - 1) Then
'      OkToAdd = False
'      Exit Function
'    End If
'  Next iloop
  
'  OkToAdd = True

  plngPrevRow = mfrmForm.grdReportOrder.AddItemRowIndex(mfrmForm.grdReportOrder.Bookmark)
  
  mfrmForm.grdReportOrder.MoveFirst
  Do Until pintLoop = mfrmForm.grdReportOrder.Rows
    pvarbookmark = mfrmForm.grdReportOrder.GetBookmark(pintLoop)
    If mfrmForm.grdReportOrder.Columns(0).CellText(pvarbookmark) = Right(sKey, Len(sKey) - 1) Then
      OkToAdd = False
      mfrmForm.grdReportOrder.Bookmark = mfrmForm.grdReportOrder.GetBookmark(plngPrevRow)
      mfrmForm.grdReportOrder.SelBookmarks.Add mfrmForm.grdReportOrder.Bookmark
      Exit Function
    End If
    pintLoop = pintLoop + 1
  Loop
  
  mfrmForm.grdReportOrder.Bookmark = mfrmForm.grdReportOrder.GetBookmark(plngPrevRow)
  mfrmForm.grdReportOrder.SelBookmarks.Add mfrmForm.grdReportOrder.Bookmark
  OkToAdd = True

End Function


Private Function GetTableNameFromColumn(lngColumnID As Long) As String

  Dim rsInfo As Recordset
  Dim strSQL As String
  
  strSQL = "SELECT ASRSysTables.TableName " & _
           "FROM ASRSysColumns JOIN ASRSysTables " & _
           "ON (ASRSysTables.TableID = ASRSysColumns.TableID) " & _
           "WHERE ColumnID = " & CStr(lngColumnID)

  Set rsInfo = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
          
  GetTableNameFromColumn = rsInfo!TableName
  
  Set rsInfo = Nothing

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Public Property Get UserCancelled() As Boolean
  UserCancelled = mblnUserCancelled
End Property


