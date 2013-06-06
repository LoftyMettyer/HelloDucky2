VERSION 5.00
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmRecordProfileTable 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Record Profile Table"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4785
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1064
   Icon            =   "frmRecordProfileTable.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraChild 
      Caption         =   "Related Table :"
      Height          =   3100
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   4605
      Begin VB.CheckBox chkPageBreak 
         Caption         =   "Pag&e Break"
         Height          =   315
         Left            =   200
         TabIndex        =   14
         Top             =   1950
         Width           =   1380
      End
      Begin VB.OptionButton optOrientation 
         Caption         =   "&Vertical"
         Height          =   195
         Index           =   1
         Left            =   3045
         TabIndex        =   17
         Top             =   2560
         Width           =   1035
      End
      Begin VB.OptionButton optOrientation 
         Caption         =   "Hori&zontal"
         Height          =   195
         Index           =   0
         Left            =   1785
         TabIndex        =   16
         Top             =   2560
         Value           =   -1  'True
         Width           =   1230
      End
      Begin VB.TextBox txtOrder 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   1100
         Width           =   2730
      End
      Begin VB.CommandButton cmdOrder 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   3825
         TabIndex        =   9
         Top             =   1100
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
         Left            =   4125
         MaskColor       =   &H000000FF&
         TabIndex        =   10
         ToolTipText     =   "Clear Order"
         Top             =   1100
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.TextBox txtFilter 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   700
         Width           =   2730
      End
      Begin VB.ComboBox cboTable 
         Height          =   315
         Left            =   1095
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   3375
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   3825
         TabIndex        =   5
         Top             =   700
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
         Left            =   4125
         MaskColor       =   &H000000FF&
         TabIndex        =   6
         ToolTipText     =   "Clear Filter"
         Top             =   700
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin COASpinner.COA_Spinner spnMaxRecords 
         Height          =   315
         Left            =   1095
         TabIndex        =   12
         Top             =   1500
         Width           =   960
         _ExtentX        =   1693
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
      Begin VB.Label lblPageBreak 
         Caption         =   " (applies only for output to Word    and printing Data Only output)"
         Height          =   405
         Left            =   1620
         TabIndex        =   20
         Top             =   1995
         Width           =   2940
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblOrientation 
         Caption         =   "Data Orientation :"
         Height          =   195
         Left            =   195
         TabIndex        =   15
         Top             =   2550
         Width           =   1665
      End
      Begin VB.Label lblOrder 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order :"
         Enabled         =   0   'False
         Height          =   195
         Left            =   195
         TabIndex        =   7
         Top             =   1155
         Width           =   525
      End
      Begin VB.Label lblFilter 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter :"
         Enabled         =   0   'False
         Height          =   195
         Left            =   200
         TabIndex        =   3
         Top             =   760
         Width           =   465
      End
      Begin VB.Label lblTable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Table :"
         Height          =   195
         Left            =   200
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblMaxRecords 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Records :"
         Enabled         =   0   'False
         Height          =   195
         Left            =   195
         TabIndex        =   11
         Top             =   1560
         Width           =   825
      End
      Begin VB.Label lblMaxRecordsAll 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(All Records)"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2205
         TabIndex        =   13
         Top             =   1560
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3495
      TabIndex        =   19
      Top             =   3300
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2235
      TabIndex        =   18
      Top             =   3300
      Width           =   1200
   End
End
Attribute VB_Name = "frmRecordProfileTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfCancelled As Boolean
Private mfLoading As Boolean

Private mlngBaseTableID As Long

'DataAccess Class
Private datData As DataMgr.clsDataAccess

Private mfrmParent As frmRecordProfile

Private mfNew As Boolean

Private msRelatedTable As String
Private mlngRelatedTableID As Long
Private msFilter As String
Private mlngFilterID As Long
Private msOrder As String
Private mlngOrderID As Long
Private mlngMaxRecords As Long

Private mfAllRelatedTablesSelected As Boolean

Private mavTables() As Variant
Private miNumberOfTables As Integer

Private Function AlreadyUsedInReport(plngTableID As Long, Optional plngExclusionID As Long) As Boolean
  Dim varBookmark As Variant
  Dim iLoop As Integer
  
  With mfrmParent.grdRelatedTables
    ' Loop thru the related tables grid.
    .MoveFirst
    Do Until iLoop = .Rows
      varBookmark = .GetBookmark(iLoop)
      If .Columns("TableID").CellText(varBookmark) = plngTableID Then
        If plngExclusionID = 0 Then
          AlreadyUsedInReport = True
          ' Loop thru the related tables grid.
          .MoveFirst
          Do Until iLoop = .Rows
            varBookmark = .GetBookmark(iLoop)
            If .Columns("TableID").CellText(varBookmark) = plngExclusionID Then
              Exit Function
            End If
            iLoop = iLoop + 1
          Loop
          Exit Function
        Else
          If .Columns("TableID").CellText(varBookmark) = plngExclusionID Then
            AlreadyUsedInReport = False
          Else
            AlreadyUsedInReport = True
            ' Loop thru the related tables grid.
            .MoveFirst
            Do Until iLoop = .Rows
              varBookmark = .GetBookmark(iLoop)
              If .Columns("TableID").CellText(varBookmark) = plngExclusionID Then
                  Exit Function
              End If
              iLoop = iLoop + 1
            Loop
            Exit Function
          End If
        End If
      End If
      iLoop = iLoop + 1
    Loop

    ' Loop thru the related tables grid.
    .MoveFirst
    Do Until iLoop = .Rows
      varBookmark = .GetBookmark(iLoop)
      If .Columns("TableID").CellText(varBookmark) = plngExclusionID Then
        Exit Function
      End If
      iLoop = iLoop + 1
    Loop
  End With

  AlreadyUsedInReport = False

End Function




Private Sub GetRelatedTables(plngTableID As Long, psRelationship As String)
  Dim sSQL As String
  Dim rsTables As ADODB.Recordset
  Dim iLoop As Integer
  Dim fFound As Boolean
  
  If psRelationship = "CHILD" Then
    sSQL = "SELECT ASRSysTables.tableName, ASRSysTables.tableID" & _
      " FROM ASRSysTables " & _
      " INNER JOIN ASRSysRelations ON ASRSysTables.tableID = ASRSysRelations.childID" & _
      " WHERE ASRSysRelations.parentID = " & Trim(Str(plngTableID))
  Else
    sSQL = "SELECT ASRSysTables.tableName, ASRSysTables.tableID" & _
      " FROM ASRSysTables " & _
      " INNER JOIN ASRSysRelations ON ASRSysTables.tableID = ASRSysRelations.parentID" & _
      " WHERE ASRSysRelations.childID = " & Trim(Str(plngTableID))
  End If
  
  Set rsTables = datData.OpenRecordset(sSQL, adOpenStatic, adLockReadOnly)
  Do While Not rsTables.EOF
    fFound = False
    
    For iLoop = 1 To UBound(mavTables, 2)
      If mavTables(1, iLoop) = rsTables!TableID Then
        fFound = True
        Exit For
      End If
    Next iLoop
    
    If fFound = False Then
      ReDim Preserve mavTables(3, UBound(mavTables, 2) + 1)
      mavTables(1, UBound(mavTables, 2)) = rsTables!TableID
      mavTables(2, UBound(mavTables, 2)) = rsTables!TableName
      mavTables(3, UBound(mavTables, 2)) = (psRelationship = "PARENT")
      
      GetRelatedTables rsTables!TableID, psRelationship
    End If
    
    rsTables.MoveNext
  Loop
  rsTables.Close
  Set rsTables = Nothing

End Sub


Public Function Initialize(pfNew As Boolean, _
  pfrmParentForm As frmRecordProfile, _
  Optional plngRelatedTableID As Long, _
  Optional psRelatedTable As String, _
  Optional plngFilterID As Long, _
  Optional psFilter As String, _
  Optional plngOrderID As Long, _
  Optional psOrder As String, _
  Optional plngMaxRecords As Long, _
  Optional piOrientation As Integer, _
  Optional pfPageBreak As Boolean, _
  Optional piNumberOfTables As Integer) As Boolean
  
  On Error GoTo Error_Trap

  ' Set references to class modules
  Set datData = New DataMgr.clsDataAccess
  
  mfNew = pfNew
  
  Set mfrmParent = pfrmParentForm

  mlngBaseTableID = mfrmParent.cboBaseTable.ItemData(mfrmParent.cboBaseTable.ListIndex)
  mlngRelatedTableID = IIf(IsMissing(plngRelatedTableID), 0, plngRelatedTableID)
  miNumberOfTables = piNumberOfTables
  
  'JPD 20030911 Fault 6359
  Me.Caption = "Record Profile Table" & IIf(miNumberOfTables > 1, "s", "")
  
  PopulateTableCombo

  If Not Cancelled Then
    If Not mfNew Then
      If miNumberOfTables > 1 Then
      Else
        SetComboText cboTable, psRelatedTable
      End If
      
      txtFilter.Text = psFilter
      txtFilter.Tag = plngFilterID
      txtOrder.Text = psOrder
      txtOrder.Tag = plngOrderID
      spnMaxRecords.Value = plngMaxRecords
      optOrientation(0).Value = (piOrientation = giHORIZONTAL)
      optOrientation(1).Value = Not optOrientation(0).Value
      chkPageBreak.Value = IIf(pfPageBreak, vbChecked, vbUnchecked)
    Else
      cboTable.ListIndex = 0
      txtFilter.Text = vbNullString
      txtFilter.Tag = 0
      txtOrder.Text = vbNullString
      txtOrder.Tag = 0
      spnMaxRecords.Value = 0
      optOrientation(0).Value = True
      optOrientation(1).Value = Not optOrientation(0).Value
      chkPageBreak.Value = vbUnchecked
    End If
  Else
    Initialize = False
    Exit Function
  End If

  msRelatedTable = RelatedTable
  
  cmdFilterClear.Enabled = (txtFilter.Tag > 0)
  cmdOrderClear.Enabled = (txtOrder.Tag > 0)
  Changed = pfNew
  Initialize = True
  
TidyUpAndExit:
  Exit Function
  
Error_Trap:
  COAMsgBox "Error initialising the the related tables form.", vbExclamation + vbOKOnly, "Record Profile"
  Initialize = False
  GoTo TidyUpAndExit

End Function

Public Property Get MaxRecords() As Long
  MaxRecords = spnMaxRecords.Value
  
End Property
Private Function PopulateTableCombo() As Boolean
  Dim iLoop As Integer
  
  On Error GoTo Error_Trap

  ' Clear Table Combo
  cboTable.Clear

  If miNumberOfTables > 1 Then
    cboTable.AddItem Trim(Str(miNumberOfTables)) & " tables"
    cboTable.ItemData(cboTable.NewIndex) = 0
  Else
    ' Get the tables related to the selected base table
    ' Put the table info into an array
    '   Column 1 = table ID
    '   Column 2 = table name
    '   Column 3 = true if this table is an ASCENDENT of the base table
    '            = false if this table is an DESCENDENT of the base table
    ReDim mavTables(3, 0)
    
    GetRelatedTables mlngBaseTableID, "PARENT"
    GetRelatedTables mlngBaseTableID, "CHILD"
  
    If UBound(mavTables, 2) = 0 Then
      COAMsgBox "All related tables for the current base table have been added to the record profile definition." _
              , vbInformation + vbOKOnly, "Record Profile"
      PopulateTableCombo = False
      Cancelled = True
      Exit Function
    Else
      For iLoop = 1 To UBound(mavTables, 2)
        If Not AlreadyUsedInReport(CLng(mavTables(1, iLoop)), IIf(mfNew, 0, mlngRelatedTableID)) Then
          cboTable.AddItem mavTables(2, iLoop)
          cboTable.ItemData(cboTable.NewIndex) = mavTables(1, iLoop)
        End If
      Next iLoop
    End If
  
    If cboTable.ListCount = 0 Then
      COAMsgBox "All related tables for the current base table have been added to the record profile definition." _
              , vbInformation + vbOKOnly, "Record Profile"
      PopulateTableCombo = False
      Cancelled = True
      Exit Function
    End If
  End If
  
  cboTable.ListIndex = 0
  cboTable.Enabled = (cboTable.ListCount > 1)
  cboTable.BackColor = IIf(cboTable.Enabled, vbWindowBackground, vbButtonFace)
  
  txtFilter.Text = ""
  txtFilter.Tag = 0
  txtFilter.Enabled = False
  
  spnMaxRecords.Value = 0
  spnMaxRecords.BackColor = IIf(spnMaxRecords.Enabled, vbWindowBackground, vbButtonFace)
  
  txtOrder.Text = ""
  txtOrder.Tag = 0
  txtOrder.Enabled = False
  
  optOrientation(0).Value = True
  optOrientation(1).Value = Not optOrientation(0).Value

TidyUpAndExit:
  Exit Function

Error_Trap:
  COAMsgBox "Error populating related tables dropdown box.", vbExclamation + vbOKOnly, "Record Profile"
  PopulateTableCombo = False
  GoTo TidyUpAndExit
  
End Function




Public Property Get Cancelled() As Boolean
  Cancelled = mfCancelled
  
End Property


Public Property Let Cancelled(pfCancelled As Boolean)
  mfCancelled = pfCancelled
  
End Property


Public Property Get RelatedTable() As String
  RelatedTable = cboTable.Text
  
End Property


Public Property Get Changed() As Boolean
  Changed = cmdOK.Enabled
  
End Property

Public Property Let Changed(ByVal pfChanged As Boolean)
  cmdOK.Enabled = pfChanged
  
End Property


Private Sub cboTable_Click()
  Dim iLoop As Integer
  
  If mfLoading = True Then Exit Sub

  txtFilter.Tag = 0
  txtFilter.Text = vbNullString
  txtOrder.Tag = 0
  txtOrder.Text = vbNullString
  spnMaxRecords.Value = 0
  Changed = True

  If miNumberOfTables > 1 Then
    lblFilter.Enabled = False
    cmdFilter.Enabled = False
    lblOrder.Enabled = False
    cmdOrder.Enabled = False
    lblMaxRecords.Enabled = False
    spnMaxRecords.Enabled = False
    spnMaxRecords.BackColor = vbButtonFace
    lblMaxRecordsAll.Enabled = False
  Else
    For iLoop = 1 To UBound(mavTables, 2)
      If mavTables(1, iLoop) = RelatedTableID Then
        ' Disable the filter/order controls if the selected table is an ASCENDANT of the base table
        lblFilter.Enabled = Not mavTables(3, iLoop)
        cmdFilter.Enabled = Not mavTables(3, iLoop)
        lblOrder.Enabled = Not mavTables(3, iLoop)
        cmdOrder.Enabled = Not mavTables(3, iLoop)
        lblMaxRecords.Enabled = Not mavTables(3, iLoop)
        spnMaxRecords.Enabled = Not mavTables(3, iLoop)
        spnMaxRecords.BackColor = IIf(mavTables(3, iLoop), vbButtonFace, vbWindowBackground)
        lblMaxRecordsAll.Enabled = Not mavTables(3, iLoop)
        
        Exit For
      End If
    Next iLoop
  End If

End Sub


Private Sub chkPageBreak_Click()
  Changed = True

End Sub




Private Sub cmdCancel_Click()
  Cancelled = True
  Unload Me

End Sub


Private Sub cmdFilter_Click()
  GetFilter cboTable, txtFilter
  cmdFilterClear.Enabled = (txtFilter.Tag > 0)

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
        If .Access = ACCESS_HIDDEN Then
          If Not mfrmParent.DefinitionOwner Then
            COAMsgBox "Unable to select this filter as it is a hidden filter and you are not the owner of this definition", vbExclamation
            If ctlTarget.Tag = .ExpressionID Or (.ExpressionID = 0) Then
              ctlTarget.Text = ""
              ctlTarget.Tag = 0
            End If
            Exit Sub
          End If
        End If
        
        ' Read the selected expression info.
        ctlTarget.Text = IIf(Len(.Name) = 0, "", .Name)
        ctlTarget.Tag = .ExpressionID
        
        Changed = True
      Else
        If ctlTarget.Tag = .ExpressionID Then
          If .Access = ACCESS_HIDDEN Then
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
  ' Allow the user to select/create/modify a filter for the Data Transfer.
  On Error GoTo ErrorTrap

  Dim lngTableID As Long
  Dim objOrder As clsOrder
  Dim sSQL As String
  Dim rsOrders As Recordset
  Dim fOK As Boolean

  fOK = True
  
  Screen.MousePointer = vbHourglass

  If TypeOf ctlSource Is TextBox Then
    lngTableID = ctlSource.Tag
  Else
    lngTableID = ctlSource.ItemData(ctlSource.ListIndex)
  End If

  ' Instantiate an order object.
  Set objOrder = New clsOrder

  With objOrder
    ' Initialize the order object.
    .OrderID = Val(ctlTarget.Tag)
    .TableID = lngTableID
    .OrderType = giORDERTYPE_DYNAMIC
    
    ' Instruct the Order object to handle the selection.
    If .SelectOrder Then
      ctlTarget.Text = .OrderName
      ctlTarget.Tag = .OrderID
      
      Changed = True
    Else
      ' Check in case the original order has been deleted.
      sSQL = "SELECT *" & _
        " FROM ASRSysOrders" & _
        " WHERE orderID = " & Trim(Str(Val(ctlTarget.Tag)))
      Set rsOrders = datGeneral.GetRecords(sSQL)
      With rsOrders
        If (.EOF And .BOF) Then
          ctlTarget.Text = ""
          ctlTarget.Tag = 0
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

Public Property Get RelatedTableID() As Long
  RelatedTableID = CLng(cboTable.ItemData(cboTable.ListIndex))
  
End Property
Public Property Get Filter() As String
  Filter = txtFilter.Text
  
End Property

Public Property Get Order() As String
  Order = txtOrder.Text
  
End Property


Public Property Get Orientation() As OrientationTypes
  Orientation = IIf(optOrientation(0).Value, giHORIZONTAL, giVERTICAL)

End Property


Public Property Get PageBreak() As Integer
  PageBreak = (chkPageBreak.Value = vbChecked)
  
End Property



Public Property Get FilterID() As Long
  FilterID = IIf(txtFilter.Tag = vbNullString, 0, CLng(txtFilter.Tag))
  
End Property



Public Property Get OrderID() As Long
  OrderID = IIf(txtOrder.Tag = vbNullString, 0, CLng(txtOrder.Tag))
  
End Property


Private Sub cmdFilterClear_Click()
  With txtFilter
    .Text = vbNullString
    .Tag = 0
  End With
  
  cmdFilterClear.Enabled = False

End Sub


Private Sub cmdOK_Click()
  If Trim(cboTable.Text) = vbNullString Then
    COAMsgBox "You must select a table.", vbExclamation, Me.Caption
    Exit Sub
  End If

  If ValidateTableInfo Then
    Cancelled = False
    Me.Hide
  End If

End Sub


Private Function ValidateTableInfo() As Boolean

  Dim fOK As Boolean
  Dim rsTemp As ADODB.Recordset
  Dim sMessage As String

  On Error GoTo Error_Trap

  fOK = True

  If txtFilter.Tag > 0 Then
    Set rsTemp = datGeneral.GetReadOnlyRecords("SELECT * FROM ASRSysExpressions WHERE exprID = " & txtFilter.Tag)

    If rsTemp.BOF And rsTemp.EOF Then
      sMessage = "The '" & txtFilter.Text & "' filter no longer exists."
      COAMsgBox sMessage, vbExclamation + vbOKOnly, "Record Profile"
      txtFilter.Text = vbNullString
      txtFilter.Tag = 0
      fOK = False
    Else
      sMessage = IsFilterValid(rsTemp!ExprID)

      If sMessage <> vbNullString Then
        sMessage = "The '" & txtFilter.Text & "' filter has been made hidden by another user."
        COAMsgBox sMessage, vbExclamation + vbOKOnly, "Record Profile"
        txtFilter.Text = vbNullString
        txtFilter.Tag = 0
        fOK = False
      End If
    End If

    rsTemp.Close
  End If

  If txtOrder.Tag > 0 Then
    Set rsTemp = datGeneral.GetReadOnlyRecords("SELECT * FROM ASRSysOrders WHERE orderID = " & txtOrder.Tag)

    If rsTemp.BOF And rsTemp.EOF Then
      sMessage = "The '" & txtOrder.Text & "' order no longer exists."
      COAMsgBox sMessage, vbExclamation + vbOKOnly, "Record Profile"
      txtOrder.Text = vbNullString
      txtOrder.Tag = 0
      fOK = False
    End If

    rsTemp.Close
  End If

  ValidateTableInfo = fOK

TidyUpAndExit:
  Set rsTemp = Nothing
  Exit Function

Error_Trap:
  COAMsgBox "Error validating related table information.", vbExclamation + vbOKOnly, "Record Profile"
  ValidateTableInfo = False
  GoTo TidyUpAndExit
  
End Function


Private Sub cmdOrder_Click()
  GetOrder cboTable, txtOrder
  cmdOrderClear.Enabled = (txtOrder.Tag > 0)

End Sub

Private Sub cmdOrderClear_Click()
  With txtOrder
    .Text = vbNullString
    .Tag = 0
  End With
  
  cmdOrderClear.Enabled = False

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
    Cancelled = True
  End If

End Sub


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


Private Sub Form_Unload(Cancel As Integer)
  Set datData = Nothing

End Sub


Private Sub optOrientation_Click(Index As Integer)
  Changed = True
  
End Sub


Private Sub spnMaxRecords_Change()
  Changed = True
  lblMaxRecordsAll.Visible = (spnMaxRecords.Value = 0)

End Sub


Private Sub txtFilter_Change()
  Changed = True
  cmdFilterClear.Enabled = (txtFilter.Tag > 0)

End Sub


Private Sub txtOrder_Change()
  Changed = True
  cmdOrderClear.Enabled = (txtOrder.Tag > 0)

End Sub




