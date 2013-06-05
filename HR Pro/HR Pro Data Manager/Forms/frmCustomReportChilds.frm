VERSION 5.00
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmCustomReportChilds 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Report Child Table"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1023
   Icon            =   "frmCustomReportChilds.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2010
      TabIndex        =   9
      Top             =   2520
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3270
      TabIndex        =   10
      Top             =   2520
      Width           =   1200
   End
   Begin VB.Frame fraChild 
      Caption         =   "Child :"
      Height          =   2265
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4380
      Begin VB.TextBox txtChildOrder 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   720
         Width           =   2505
      End
      Begin VB.CommandButton cmdChildOrder 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   3595
         TabIndex        =   3
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdOrderClear 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Wingdings 2"
            Size            =   20.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3945
         TabIndex        =   4
         ToolTipText     =   "Clear Order"
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdFilterClear 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Wingdings 2"
            Size            =   20.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3945
         TabIndex        =   7
         ToolTipText     =   "Clear Filter"
         Top             =   1140
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdChildFilter 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   3600
         TabIndex        =   6
         Top             =   1140
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.ComboBox cboChild 
         Height          =   315
         Left            =   1095
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   3150
      End
      Begin VB.TextBox txtChildFilter 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   1140
         Width           =   2505
      End
      Begin COASpinner.COA_Spinner spnMaxRecords 
         Height          =   315
         Left            =   1095
         TabIndex        =   8
         Top             =   1545
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   556
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         MaximumValue    =   999
         Text            =   "0"
      End
      Begin VB.Label lblChildOrder 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order :"
         Enabled         =   0   'False
         Height          =   195
         Left            =   195
         TabIndex        =   15
         Top             =   780
         Width           =   525
      End
      Begin VB.Label lblMaxRecordsAll 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(All Records)"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2205
         TabIndex        =   14
         Top             =   1605
         Width           =   1230
      End
      Begin VB.Label lblMaxRecords 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Records :"
         Enabled         =   0   'False
         Height          =   195
         Left            =   195
         TabIndex        =   13
         Top             =   1605
         Width           =   825
      End
      Begin VB.Label lblChildTable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Table :"
         Height          =   195
         Left            =   195
         TabIndex        =   12
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblChildFilter 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter :"
         Enabled         =   0   'False
         Height          =   195
         Left            =   195
         TabIndex        =   11
         Top             =   1200
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmCustomReportChilds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbCancelled As Boolean
Private mblnLoading As Boolean

Private mlngBaseTableID As Long

'DataAccess Class
Private datData As HRProDataMgr.clsDataAccess

Private mfrmParent As frmCustomReports

Private mblnNew As Boolean

Private mstrChildTable As String
Private mstrOrder As String
Private mlngOrderID As Long
Private mstrFilter As String
Private mlngFilterID As Long
Private mlngMaxRecords As Long
Private mlngChildTableID As Long

Private mbAllChildsSelected As Boolean

Public Property Let Cancelled(bCancelled As Boolean)
  mbCancelled = bCancelled
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mbCancelled
End Property

Private Function AlreadyUsedInReport(plngChildTableID As Long, Optional plngExclusion As Long) As Boolean

  Dim pintOldPosition As Integer
  Dim pvarbookmark As Variant
  Dim pintLoop As Integer
  Dim pvarOriginalBM As Variant
  
  With mfrmParent.grdChildren
    ' Loop thru the child tables grid.
    pvarOriginalBM = .Bookmark
    .Redraw = False
    .MoveFirst
    Do Until pintLoop = .Rows
      pvarbookmark = .GetBookmark(pintLoop)
      If .Columns("TableID").CellText(pvarbookmark) = plngChildTableID Then
        If plngExclusion = 0 Then
          AlreadyUsedInReport = True
          ' Loop thru the child tables grid.
          .MoveFirst
          Do Until pintLoop = .Rows
            pvarbookmark = .GetBookmark(pintLoop)
            If .Columns("TableID").CellText(pvarbookmark) = plngExclusion Then
              .Bookmark = pvarOriginalBM
              .Redraw = True
              Exit Function
            End If
            pintLoop = pintLoop + 1
          Loop
          .Bookmark = pvarOriginalBM
          .Redraw = True
          Exit Function
        Else
          If .Columns("TableID").CellText(pvarbookmark) = plngExclusion Then
            AlreadyUsedInReport = False
          Else
            AlreadyUsedInReport = True
            ' Loop thru the child tables grid.
            .MoveFirst
            Do Until pintLoop = .Rows
              pvarbookmark = .GetBookmark(pintLoop)
              If .Columns("TableID").CellText(pvarbookmark) = plngExclusion Then
                .Bookmark = pvarOriginalBM
                .Redraw = True
                Exit Function
              End If
              pintLoop = pintLoop + 1
            Loop
            .Bookmark = pvarOriginalBM
            .Redraw = True
            Exit Function
          End If
        End If
      End If
      pintLoop = pintLoop + 1
    Loop
  
    ' Loop thru the child tables grid.
    .MoveFirst
    Do Until pintLoop = .Rows
      pvarbookmark = .GetBookmark(pintLoop)
      If .Columns("TableID").CellText(pvarbookmark) = plngExclusion Then
        .Bookmark = pvarOriginalBM
        .Redraw = True
        Exit Function
      End If
      pintLoop = pintLoop + 1
    Loop
    
    .Bookmark = pvarOriginalBM
    .Redraw = True
  End With
  
  AlreadyUsedInReport = False

End Function

Private Function PopulateChildCombo() As Boolean

  Dim sSQL As String
  Dim rsChildren As ADODB.Recordset
  
  On Error GoTo Error_Trap
  
  ' Clear Child Combo and add <None> entry
  cboChild.Clear

  ' Get the children of the selected base table
  sSQL = "SELECT asrsystables.tablename, asrsystables.tableid " & _
         "FROM asrsystables " & _
         "WHERE asrsystables.tableid in " & _
         "(select childid from asrsysrelations " & _
         "WHERE parentid = " & CStr(mlngBaseTableID) & ") " & _
         "ORDER BY tablename"
  
  Set rsChildren = datData.OpenRecordset(sSQL, adOpenStatic, adLockReadOnly)
  
  If Not rsChildren.BOF And Not rsChildren.EOF Then
    If (rsChildren.RecordCount = mfrmParent.grdChildren.Rows) And mblnNew Then
      COAMsgBox "All child tables for the current base table have been added to the report definition." _
              , vbInformation + vbOKOnly, "Custom Reports"
      PopulateChildCombo = False
      Me.Cancelled = True
      Exit Function
    Else
      Do Until rsChildren.EOF
        If AlreadyUsedInReport(rsChildren!TableID, IIf(mblnNew, 0, mlngChildTableID)) = False Then
          cboChild.AddItem rsChildren!TableName
          cboChild.ItemData(cboChild.NewIndex) = rsChildren!TableID
        End If
        rsChildren.MoveNext
      Loop
    End If
  End If
  
  Select Case cboChild.ListCount
  Case 0:
    cboChild.Enabled = False
    cboChild.BackColor = IIf(cboChild.Enabled, vbWindowBackground, vbButtonFace)
    txtChildOrder.Text = ""
    txtChildOrder.Tag = 0
    txtChildFilter.Text = ""
    txtChildFilter.Tag = 0
    txtChildFilter.Enabled = False
    spnMaxRecords.Value = 0
    spnMaxRecords.Enabled = False
    spnMaxRecords.BackColor = IIf(spnMaxRecords.Enabled, vbWindowBackground, vbButtonFace)
    lblChildOrder.Enabled = False
    lblChildFilter.Enabled = False
    lblChildTable.Enabled = False
    lblMaxRecords.Enabled = False
    lblMaxRecordsAll.Enabled = False
    cmdChildFilter.Enabled = True
    cmdChildOrder.Enabled = True
    
  Case 1:
    cboChild.ListIndex = 0
    cboChild.Enabled = False
    cboChild.BackColor = IIf(cboChild.Enabled, vbWindowBackground, vbButtonFace)
    txtChildOrder.Text = ""
    txtChildOrder.Tag = 0
    txtChildFilter.Text = ""
    txtChildFilter.Tag = 0
    txtChildFilter.Enabled = False
    spnMaxRecords.Value = 0
    spnMaxRecords.Enabled = True
    spnMaxRecords.BackColor = IIf(spnMaxRecords.Enabled, vbWindowBackground, vbButtonFace)
    lblChildOrder.Enabled = True
    lblChildFilter.Enabled = True
    lblChildTable.Enabled = True
    lblMaxRecords.Enabled = True
    lblMaxRecordsAll.Enabled = True
    cmdChildFilter.Enabled = True
    cmdChildOrder.Enabled = True
    
  Case Is > 1:
    cboChild.ListIndex = 0
    cboChild.Enabled = True
    cboChild.BackColor = IIf(cboChild.Enabled, vbWindowBackground, vbButtonFace)
    txtChildOrder.Text = ""
    txtChildOrder.Tag = 0
    txtChildFilter.Text = ""
    txtChildFilter.Tag = 0
    txtChildFilter.Enabled = False
    spnMaxRecords.Value = 0
    spnMaxRecords.Enabled = True
    spnMaxRecords.BackColor = IIf(spnMaxRecords.Enabled, vbWindowBackground, vbButtonFace)
    lblChildOrder.Enabled = True
    lblChildFilter.Enabled = True
    lblChildTable.Enabled = True
    lblMaxRecords.Enabled = True
    lblMaxRecordsAll.Enabled = True
    cmdChildFilter.Enabled = True
    cmdChildOrder.Enabled = True
    
  Case Else:
    cboChild.Enabled = False
    cboChild.BackColor = IIf(cboChild.Enabled, vbWindowBackground, vbButtonFace)
    txtChildOrder.Text = ""
    txtChildOrder.Tag = 0
    txtChildFilter.Text = ""
    txtChildFilter.Tag = 0
    txtChildFilter.Enabled = False
    spnMaxRecords.Value = 0
    spnMaxRecords.Enabled = False
    spnMaxRecords.BackColor = IIf(spnMaxRecords.Enabled, vbWindowBackground, vbButtonFace)
    lblChildOrder.Enabled = False
    lblChildFilter.Enabled = False
    lblChildTable.Enabled = False
    lblMaxRecords.Enabled = False
    lblMaxRecordsAll.Enabled = False
    cmdChildFilter.Enabled = True
    cmdChildOrder.Enabled = True
  
  End Select
  
TidyUpAndExit:
  Set rsChildren = Nothing
  Exit Function
  
Error_Trap:
  COAMsgBox "Error populating childs dropdown box.", vbExclamation + vbOKOnly, "Custom Reports"
  PopulateChildCombo = False
  GoTo TidyUpAndExit
  
End Function



Private Function ValidateChildInfo() As Boolean

  Dim bOK As Boolean
  Dim prstTemp As ADODB.Recordset
  Dim sMessage As String
  
  On Error GoTo Error_Trap
  
  bOK = True
  
  If Me.txtChildFilter.Tag > 0 Then
    Set prstTemp = datGeneral.GetReadOnlyRecords("SELECT * FROM ASRSysExpressions WHERE exprID = " & Me.txtChildFilter.Tag)
    

    If prstTemp.BOF And prstTemp.EOF Then
      sMessage = "The '" & Me.txtChildFilter.Text & "' filter has been deleted by another user."
      COAMsgBox sMessage, vbExclamation + vbOKOnly, "Custom Reports"
      Me.txtChildFilter.Text = vbNullString
      Me.txtChildFilter.Tag = 0
      bOK = False
    Else
      sMessage = IsFilterValid(prstTemp!ExprID)
      
      If sMessage <> vbNullString Then
        sMessage = "The '" & Me.txtChildFilter.Text & "' filter has been made hidden by another user."
        COAMsgBox sMessage, vbExclamation + vbOKOnly, "Custom Reports"
        Me.txtChildFilter.Text = vbNullString
        Me.txtChildFilter.Tag = 0
        bOK = False
      End If
    End If
    
    prstTemp.Close
  End If
  
  ValidateChildInfo = bOK

TidyUpAndExit:
  Set prstTemp = Nothing
  Exit Function
  
Error_Trap:
  COAMsgBox "Error validating child table information.", vbExclamation + vbOKOnly, "Custom Reports"
  ValidateChildInfo = False
  GoTo TidyUpAndExit
  
End Function

Private Sub cmdCancel_Click()
  Me.Cancelled = True
  Unload Me
End Sub

Private Sub cmdChildFilter_Click()

  Dim lngOriginalChildFilterID As Long
  
  lngOriginalChildFilterID = Me.txtChildFilter.Tag
  
  GetFilter cboChild, txtChildFilter
  
  'Changed = (lngOriginalChildFilterID <> Me.txtChildFilter.Tag)

  Me.cmdFilterClear.Enabled = (Me.txtChildFilter.Tag > 0)
  
End Sub

Private Sub GetFilter(ctlSource As Control, ctlTarget As Control)
  ' Allow the user to select/create/modify a filter for the Data Transfer.
  Dim fOK As Boolean
  Dim objExpression As clsExprExpression
  
  ' Instantiate a new expression object.
  Set objExpression = New clsExprExpression
  
  With objExpression
    ' Initialise the expression object.
    If TypeOf ctlSource Is TextBox Then
      fOK = .Initialise(ctlSource.Tag, Val(ctlTarget.Tag), giEXPR_RUNTIMEFILTER, giEXPRVALUE_LOGIC)
    ElseIf TypeOf ctlSource Is ComboBox Then
      fOK = .Initialise(ctlSource.ItemData(ctlSource.ListIndex), Val(ctlTarget.Tag), giEXPR_RUNTIMEFILTER, giEXPRVALUE_LOGIC)
    End If
      
    If fOK Then
      ' Instruct the expression object to display the expression selection/creation/modification form.
      If .SelectExpression(True) = True Then
      
  '      If .Access = 2 Then
        If .Access = "HD" Then
          If Not mfrmParent.DefinitionOwner Then
            COAMsgBox "Unable to select this filter as it is a hidden filter and you are not the owner of this definition", vbExclamation
            If ctlTarget.Tag = .ExpressionID Or (.ExpressionID = 0) Then
              ctlTarget.Text = ""
              ctlTarget.Tag = 0
            End If
            Exit Sub
'          Else
'              If (mblnForceHidden = False) And (Me.optHidden = False) Then COAMsgBox "This definition will now be hidden as a hidden filter has been selected", vbInformation
          End If
'        Else
'          optReadWrite.Enabled = mblnDefinitionCreator
'          optReadOnly.Enabled = mblnDefinitionCreator
'          optHidden.Enabled = mblnDefinitionCreator
        End If
        
        ' Read the selected expression info.
        ctlTarget.Text = IIf(Len(.Name) = 0, "", .Name)
        ctlTarget.Tag = .ExpressionID
        
        Changed = True
        
      Else
        If ctlTarget.Tag = .ExpressionID Then
          If .Access = "HD" Then
            If Not mfrmParent.DefinitionOwner Then
              COAMsgBox "Unable to select this filter as it is a hidden filter and you are not the owner of this definition", vbExclamation
              ctlTarget.Text = ""
              ctlTarget.Tag = 0
              Exit Sub
            End If
          End If
        End If
      End If
      
    End If
    
    End With
  
  
  Set objExpression = Nothing

End Sub

Private Sub GetOrder(ctlSource As Control, ctlTarget As Control)
  ' Display the 'Order' selection form.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim sSQL As String
  Dim objOrder As clsOrder
  Dim rsOrders As Recordset
  
  fOK = True

  ' Instantiate an order object.
  Set objOrder = New clsOrder

  With objOrder
    ' Initialize the order object.
    .OrderID = ctlTarget.Tag
    .TableID = ctlSource.ItemData(ctlSource.ListIndex)
    
    .OrderType = giORDERTYPE_DYNAMIC
    
    ' Instruct the Order object to handle the selection.
    If .SelectOrder Then
      ctlTarget.Tag = .OrderID
      ctlTarget.Text = .OrderName
    Else
      ' Check in case the original expression has been deleted.
      sSQL = "SELECT *" & _
        " FROM ASRSysOrders" & _
        " WHERE orderID = " & Trim(Str(ctlTarget.Tag))
      Set rsOrders = datGeneral.GetRecords(sSQL)
      With rsOrders
        If (.EOF And .BOF) Then
          ctlTarget.Tag = 0
          ctlTarget.Text = ""
        End If
      
        .Close
      End With
      Set rsOrders = Nothing
    End If
  End With

TidyUpAndExit:
  Set objOrder = Nothing
  If Not fOK Then
    COAMsgBox "Error changing order ID.", vbExclamation + vbOKOnly, App.ProductName
  End If
  Exit Sub

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub

Private Sub cmdChildOrder_Click()
  Dim lngOriginalChildOrderID As Long
  
  lngOriginalChildOrderID = Me.txtChildOrder.Tag
  
  GetOrder cboChild, txtChildOrder
  
  'Changed = (lngOriginalChildOrderID <> Me.txtChildOrder.Tag)

  Me.cmdOrderClear.Enabled = (Me.txtChildOrder.Tag > 0)

End Sub

Private Sub cmdFilterClear_Click()

  With txtChildFilter
    .Text = vbNullString
    .Tag = 0
  End With
  
  Me.cmdFilterClear.Enabled = False
  
End Sub

Private Sub cmdOK_Click()

  If Trim(Me.cboChild.Text) = vbNullString Then
    COAMsgBox "You must select a table and column.", vbExclamation, Me.Caption
    Exit Sub
  End If

  If ValidateChildInfo Then
    Cancelled = False
    Me.Hide
  End If

End Sub


Private Sub cmdOrderClear_Click()
  
  With txtChildOrder
    .Text = vbNullString
    .Tag = 0
  End With
  
  Me.cmdOrderClear.Enabled = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then
    Me.Cancelled = True
  End If
End Sub


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


Private Sub Form_Unload(Cancel As Integer)
  Set datData = Nothing
End Sub


Private Sub spnMaxRecords_Change()

  Changed = True
  lblMaxRecordsAll.Visible = spnMaxRecords.Value = 0

End Sub

Private Sub cboChild_Click()

  If mblnLoading = True Then Exit Sub

  'TM20020517 Fault 3889
  ' If no change the exit sub, else set changed flag and reset the values
  ' for the filter and max records.
  If cboChild.Text = mstrChildTable Then
    Exit Sub
  Else
    Me.txtChildFilter.Tag = 0
    Me.txtChildFilter.Text = vbNullString
    'NHRD20100519 Jira 714
    Me.txtChildOrder.Tag = 0
    Me.txtChildOrder.Text = vbNullString
    
    Me.spnMaxRecords.Value = 0
    
    Me.Changed = True
  End If
  
End Sub

Public Property Get ChildTableID() As Long
  ChildTableID = CLng(Me.cboChild.ItemData(Me.cboChild.ListIndex))
End Property

Public Property Get ChildTable() As String
  ChildTable = Me.cboChild.Text
End Property

Public Property Get OrderID() As Long
  OrderID = IIf(Me.txtChildOrder.Tag = vbNullString, 0, CLng(Me.txtChildOrder.Tag))
End Property

Public Property Get Order() As String
  Order = Me.txtChildOrder.Text
End Property

Public Property Get FilterID() As Long
  FilterID = IIf(Me.txtChildFilter.Tag = vbNullString, 0, CLng(Me.txtChildFilter.Tag))
End Property

Public Property Get Filter() As String
  Filter = Me.txtChildFilter.Text
End Property

Public Property Get MaxRecords() As Long
  MaxRecords = Me.spnMaxRecords.Value
End Property

Public Function Initialize(bNew As Boolean, frmParentForm As frmCustomReports, Optional lngChildTableID As Long, _
                            Optional sChildTable As String, Optional lngFilterID As Long, _
                            Optional sFilter As String, Optional lngMaxRecords As Long, _
                            Optional lngOrderID As Long, Optional sOrder As String) As Boolean
  
  On Error GoTo Error_Trap

  ' Set references to class modules
  Set datData = New HRProDataMgr.clsDataAccess
  
  mblnNew = bNew
  
  Set mfrmParent = frmParentForm

  mlngBaseTableID = mfrmParent.cboBaseTable.ItemData(mfrmParent.cboBaseTable.ListIndex)
  mlngChildTableID = IIf(IsMissing(lngChildTableID), 0, lngChildTableID)
 
  PopulateChildCombo
  
  'TM25092003 Fault 7052 - this Doevents was doing the multiple clicks before showing the form as modal.
  'DoEvents
  
  If Not Me.Cancelled Then
    If Not mblnNew Then
      SetComboText cboChild, sChildTable
      txtChildOrder.Text = sOrder
      txtChildOrder.Tag = lngOrderID
      txtChildFilter.Text = sFilter
      txtChildFilter.Tag = lngFilterID
      spnMaxRecords.Value = lngMaxRecords
    Else
      Me.cboChild.ListIndex = 0
      Me.txtChildOrder.Text = vbNullString
      Me.txtChildOrder.Tag = 0
      Me.txtChildFilter.Text = vbNullString
      Me.txtChildFilter.Tag = 0
      Me.spnMaxRecords.Value = 0
    End If
  Else
    Initialize = False
    Exit Function
  End If

  mstrChildTable = Me.ChildTable
  Me.cmdOrderClear.Enabled = (Me.txtChildOrder.Tag > 0)
  Me.cmdFilterClear.Enabled = (Me.txtChildFilter.Tag > 0)
  Me.Changed = bNew
  
  Initialize = True
  
TidyUpAndExit:
  Exit Function
  
Error_Trap:
  COAMsgBox "Error initialising the the child tables form.", vbExclamation + vbOKOnly, "Custom Reports"
  Initialize = False
  GoTo TidyUpAndExit

End Function

Public Property Get Changed() As Boolean
  Changed = cmdOK.Enabled
End Property
Public Property Let Changed(ByVal pblnChanged As Boolean)
  cmdOK.Enabled = pblnChanged
End Property

Private Sub txtChildFilter_Change()

  Me.Changed = True
  
  Me.cmdFilterClear.Enabled = (Me.txtChildFilter.Tag > 0)
  
End Sub


Private Sub txtChildOrder_Change()
  Me.Changed = True
  
  Me.cmdOrderClear.Enabled = (Me.txtChildOrder.Tag > 0)
  
End Sub


