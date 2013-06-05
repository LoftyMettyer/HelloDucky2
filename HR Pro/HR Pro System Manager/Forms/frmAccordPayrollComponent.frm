VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{BE7AC23D-7A0E-4876-AFA2-6BAFA3615375}#1.0#0"; "COA_Spinner.ocx"
Begin VB.Form frmAccordPayrollComponent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Payroll Transfer Field"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6885
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5059
   Icon            =   "frmAccordPayrollComponent.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDefinition 
      Caption         =   "Definition : "
      Height          =   795
      Left            =   135
      TabIndex        =   24
      Top             =   135
      Width           =   6650
      Begin VB.CheckBox chkKeyField 
         Caption         =   "Key Field"
         Enabled         =   0   'False
         Height          =   240
         Left            =   5085
         TabIndex        =   2
         Top             =   315
         Width           =   1230
      End
      Begin VB.ComboBox cboTransferField 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1530
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   270
         Width           =   3120
      End
      Begin VB.Label lblTransferField 
         Caption         =   "Transfer Field : "
         Height          =   240
         Left            =   180
         TabIndex        =   25
         Top             =   315
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   5580
      TabIndex        =   20
      Top             =   4980
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   405
      Left            =   4275
      TabIndex        =   19
      Top             =   4980
      Width           =   1200
   End
   Begin VB.Frame fraComponentType 
      Caption         =   "Type :"
      Height          =   3930
      Left            =   135
      TabIndex        =   0
      Top             =   960
      Width           =   1890
      Begin VB.OptionButton optComponentType 
         Caption         =   "Colu&mn"
         Height          =   315
         Index           =   0
         Left            =   165
         TabIndex        =   3
         Tag             =   "COMP_FIELD"
         Top             =   300
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.OptionButton optComponentType 
         Caption         =   "C&alculation"
         Height          =   315
         Index           =   1
         Left            =   165
         TabIndex        =   4
         Tag             =   "COMP_OPERATOR"
         Top             =   650
         Width           =   1590
      End
      Begin VB.OptionButton optComponentType 
         Caption         =   "&Value"
         Height          =   315
         Index           =   2
         Left            =   165
         TabIndex        =   5
         Tag             =   "COMP_VALUE"
         Top             =   990
         Width           =   1110
      End
   End
   Begin VB.Frame fraComponent 
      Caption         =   "Column :"
      Height          =   3930
      Index           =   0
      Left            =   2115
      TabIndex        =   16
      Tag             =   "1"
      Top             =   960
      Width           =   4665
      Begin VB.CheckBox chkPreventModify 
         Caption         =   "&Prevent data updates once in Payroll"
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   3165
         Width           =   3660
      End
      Begin COASpinner.COA_Spinner spnGroup 
         Height          =   315
         Left            =   1875
         TabIndex        =   13
         Top             =   3495
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
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
         MaximumValue    =   99999
         Text            =   "0"
      End
      Begin VB.CheckBox chkGroupBy 
         Caption         =   "&Group Number"
         Height          =   225
         Left            =   180
         TabIndex        =   12
         Top             =   3525
         Width           =   2550
      End
      Begin VB.CheckBox chkAlwaysTransfer 
         Caption         =   "Al&ways Transfer"
         Height          =   240
         Left            =   180
         TabIndex        =   10
         Top             =   2835
         Width           =   4380
      End
      Begin VB.CheckBox chkConvertData 
         Caption         =   "Convert &Data"
         Height          =   285
         Left            =   180
         TabIndex        =   8
         Top             =   1125
         Width           =   1500
      End
      Begin VB.ComboBox cboFldTable 
         Height          =   315
         ItemData        =   "frmAccordPayrollComponent.frx":000C
         Left            =   990
         List            =   "frmAccordPayrollComponent.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   285
         Width           =   3525
      End
      Begin VB.ComboBox cboFldColumn 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   705
         Width           =   3525
      End
      Begin SSDataWidgets_B.SSDBGrid grdColumnMapping 
         Height          =   1260
         Left            =   900
         TabIndex        =   9
         Top             =   1440
         Width           =   3585
         ScrollBars      =   2
         _Version        =   196617
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   2
         AllowDelete     =   -1  'True
         MultiLine       =   0   'False
         RowSelectionStyle=   2
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   3
         SelectByCell    =   -1  'True
         BalloonHelp     =   0   'False
         MaxSelectedRows =   0
         ForeColorEven   =   -2147483640
         ForeColorOdd    =   -2147483640
         BackColorEven   =   -2147483643
         BackColorOdd    =   -2147483643
         RowHeight       =   423
         ExtraHeight     =   79
         Columns.Count   =   2
         Columns(0).Width=   3625
         Columns(0).Caption=   "HR Pro Value"
         Columns(0).Name =   "HRProValue"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2117
         Columns(1).Caption=   "Payroll Value"
         Columns(1).Name =   "AccordValue"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   6324
         _ExtentY        =   2222
         _StockProps     =   79
         Enabled         =   0   'False
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblColumn 
         Caption         =   "Column :"
         Height          =   285
         Left            =   135
         TabIndex        =   23
         Top             =   765
         Width           =   870
      End
      Begin VB.Label lblTable 
         Caption         =   "Table : "
         Height          =   240
         Left            =   135
         TabIndex        =   22
         Top             =   360
         Width           =   690
      End
   End
   Begin VB.Frame fraComponent 
      Caption         =   "Calculation :"
      Height          =   3930
      Index           =   1
      Left            =   2115
      TabIndex        =   17
      Tag             =   "1"
      Top             =   960
      Width           =   4665
      Begin VB.CommandButton cmdCalculation 
         Caption         =   "..."
         Height          =   315
         Left            =   4110
         Picture         =   "frmAccordPayrollComponent.frx":0010
         TabIndex        =   14
         Top             =   380
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.TextBox txtCalculation 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   21
         Top             =   380
         Width           =   3885
      End
   End
   Begin VB.Frame fraComponent 
      Caption         =   "Value :"
      Height          =   3930
      Index           =   2
      Left            =   2115
      TabIndex        =   18
      Tag             =   "1"
      Top             =   960
      Width           =   4665
      Begin VB.TextBox txtText 
         Height          =   285
         Left            =   225
         TabIndex        =   15
         Top             =   380
         Width           =   4155
      End
   End
End
Attribute VB_Name = "frmAccordPayrollComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbReadOnly As Boolean
Private mbLoading As Boolean

Private mlngTransferID As Long
Private mlngTransferFieldID As Long

Private miMapType As HRProSystemMgr.AccordMapType
Private mlngBaseTableID As Long
Private mlngTableID As Long
Private mlngColumnID As Long
Private mlngExprID As Long
Private mstrValue As String
Private mbCancelled As Boolean
Private mstrDescription As String
Private mbIsKeyField As Boolean
Private mbIsCompanyCode As Boolean
Private mbIsEmployeeCode As Boolean
Private mlngDirection As Long
Private mbAlwaysTransferField As Boolean
Private mbGroup As Boolean
Private mlngGroup As Long
Private mbConvertData As Boolean
Private mbUndefined As Boolean
Private mbForceAlwaysTransfer As Boolean
Private mbPreventModify As Boolean

Public Property Let Changed(pbNewValue As Boolean)
  cmdOK.Enabled = pbNewValue And Not mbLoading
End Property

Public Property Let TransferFieldID(ByVal plngNewValue As Long)
  mlngTransferFieldID = plngNewValue
End Property

Public Property Let TransferID(ByVal plngNewValue As Long)
  mlngTransferID = plngNewValue
End Property

Public Property Let ConvertData(ByVal pbNewValue As Boolean)
  mbConvertData = pbNewValue
End Property

Public Property Get ConvertData() As Boolean
  ConvertData = mbConvertData
End Property

Public Property Let AlwaysTransferField(ByVal pbNewValue As Boolean)
  mbAlwaysTransferField = pbNewValue
End Property

Public Property Get AlwaysTransferField() As Boolean
  AlwaysTransferField = mbAlwaysTransferField
End Property

Public Property Let PreventModify(ByVal pbNewValue As Boolean)
  mbPreventModify = pbNewValue
End Property

Public Property Get PreventModify() As Boolean
  PreventModify = mbPreventModify
End Property

Public Property Let BaseTableID(ByVal plngNewValue As Long)
  mlngBaseTableID = plngNewValue
End Property

Public Property Let Direction(ByVal plngNewValue As Long)
  mlngDirection = plngNewValue
End Property

Public Property Get Direction() As Long
  Direction = mlngDirection
End Property

Public Property Let IsKeyField(ByVal pbNewValue As Boolean)
  mbIsKeyField = pbNewValue
End Property

Public Property Get IsKeyField() As Boolean
  IsKeyField = mbIsKeyField
End Property

Public Property Let IsCompanyCode(ByVal pbNewValue As Boolean)
  mbIsCompanyCode = pbNewValue
End Property

Public Property Let IsEmployeeCode(ByVal pbNewValue As Boolean)
  mbIsEmployeeCode = pbNewValue
End Property

Public Property Let Description(ByVal pstrNewValue As String)
  mstrDescription = pstrNewValue
End Property

Public Property Get Description() As String
  Description = mstrDescription
End Property

Public Property Let MapType(ByVal piNewValue As Integer)
  miMapType = piNewValue
End Property

Public Property Get MapType() As Integer
  MapType = miMapType
End Property

Public Property Let TableID(ByVal plngNewValue As Long)
  mlngTableID = plngNewValue
End Property

Public Property Get TableID() As Long
  TableID = mlngTableID
End Property

Public Property Let ColumnID(ByVal plngNewValue As Long)
  mlngColumnID = plngNewValue
End Property

Public Property Get ColumnID() As Long
  ColumnID = mlngColumnID
End Property

Public Property Let ExprID(ByVal plngNewValue As Long)
  mlngExprID = plngNewValue
End Property

Public Property Get ExprID() As Long
  ExprID = mlngExprID
End Property

Public Property Let value(ByVal pstrNewValue As String)
  mstrValue = pstrNewValue
End Property

Public Property Get value() As String
  value = mstrValue
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mbCancelled
End Property

Private Sub cboFldColumn_Click()
  mlngColumnID = cboFldColumn.ItemData(cboFldColumn.ListIndex)
  Me.Changed = True
End Sub

Private Sub cboFldTable_Click()

  mlngTableID = GetComboItem(cboFldTable)
  cboFldColumn_Refresh
  Me.Changed = True
  
End Sub

Private Sub chkAlwaysTransfer_Click()
  mbAlwaysTransferField = chkAlwaysTransfer.value
  Me.Changed = True
End Sub

Private Sub chkConvertData_Click()
  mbConvertData = (chkConvertData.value = vbChecked)
  grdColumnMapping.Enabled = mbConvertData And Not mbReadOnly
  
  grdColumnMapping.BackColorOdd = IIf(grdColumnMapping.Enabled, vbWhite, vbButtonFace)
  grdColumnMapping.BackColorEven = IIf(grdColumnMapping.Enabled, vbWhite, vbButtonFace)
  grdColumnMapping.AllowAddNew = grdColumnMapping.Enabled
  grdColumnMapping.Refresh
 
  Me.Changed = True
End Sub

Private Sub chkGroupBy_Click()
  mbGroup = (chkGroupBy.value = vbChecked)
  If Not mbGroup Then
    spnGroup.value = 0
  End If
  spnGroup.Enabled = mbGroup And Not mbReadOnly

  Me.Changed = True
End Sub

Private Sub chkPreventModify_Click()
  mbPreventModify = chkPreventModify.value
  Me.Changed = True
End Sub

Private Sub cmdCancel_Click()
  
'  If cmdOk.Enabled Then
'    Select Case MsgBox("Apply changes ?", vbYesNo + vbQuestion, Me.Caption)
'
'      Case vbNo
'        mbCancelled = True
'        UnLoad Me
'
'      Case vbYes
'        If Validate Then
'          SaveMappings
'          mbCancelled = False
'          UnLoad Me
'        End If
'    End Select
'  Else
    mbCancelled = True
    UnLoad Me
'  End If
  
End Sub

Private Sub cmdOK_Click()
  
  If Validate Then
    SaveMappings
    mbCancelled = False
    UnLoad Me
  End If
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_Load()
  Const GRIDROWHEIGHT = 239
  
  grdColumnMapping.RowHeight = GRIDROWHEIGHT

  mbReadOnly = (Application.AccessMode <> accFull And Application.AccessMode <> accSupportMode)
  mbUndefined = (miMapType = MAPTYPE_COLUMN And mlngColumnID = 0)
  mbLoading = True
  cboTransferField.Clear
  cboTransferField.AddItem mstrDescription
  cboTransferField.Text = mstrDescription
  
  ControlsDisableAll Me, Not mbReadOnly
  ControlsDisableAll fraDefinition, False
  ControlsDisableAll txtCalculation, False
  ControlsDisableAll grdColumnMapping, False
  
  LoadMappings
  cboFldTable_Refresh
  cboFldColumn_Refresh
  
  optComponentType(miMapType).value = True
  optComponentType_Click (miMapType)
  
  Select Case miMapType
    Case MAPTYPE_COLUMN
      If mlngColumnID > 0 Then
        SetComboItem cboFldColumn, mlngColumnID
      End If
  
    Case MAPTYPE_EXPRESSION
      txtCalculation.Text = ""
      GetCalculationExpressionDetails
  
    Case MAPTYPE_VALUE
      txtText.Text = mstrValue
  
  End Select
  
  chkKeyField.value = IIf(mbIsKeyField, vbChecked, vbUnchecked)
  chkAlwaysTransfer.Enabled = Not (mbIsKeyField Or mbReadOnly Or mbForceAlwaysTransfer)
  chkAlwaysTransfer.value = IIf(mbAlwaysTransferField, vbChecked, vbUnchecked)
  chkConvertData.value = IIf(mbConvertData, vbChecked, vbUnchecked)
  chkPreventModify.value = IIf(mbPreventModify, vbChecked, vbUnchecked)
  
  spnGroup.Enabled = False
  chkGroupBy.value = IIf(mlngGroup > 0, vbChecked, vbUnchecked)
  spnGroup.value = mlngGroup
  
  optComponentType(1).Enabled = IIf(mbIsKeyField, False, True) And Not mbReadOnly
  optComponentType(2).Enabled = IIf(mbIsEmployeeCode, False, True) And Not mbReadOnly
   
  mbLoading = False
  cmdOK.Enabled = mbUndefined And Not mbReadOnly

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  ' If the user cancels or tries to close the form
  If UnloadMode <> vbFormCode And cmdOK.Enabled Then
    Select Case MsgBox("Apply changes ?", vbYesNoCancel + vbQuestion, Me.Caption)
      Case vbCancel
        Cancel = True
      Case vbNo
        mbCancelled = True
      Case vbYes
        If Validate Then
          SaveMappings
          mbCancelled = False
        End If
    End Select
  End If

End Sub

Private Sub grdColumnMapping_Change()
  Me.Changed = True
End Sub

Private Sub optComponentType_Click(Index As Integer)

  Dim iCount As Integer

  For iCount = fraComponent.LBound To fraComponent.UBound
    fraComponent(iCount).Enabled = (iCount = Index)
  Next iCount

  fraComponent(Index).ZOrder 0
  miMapType = Index
  Me.Changed = True
  
End Sub

Private Sub spnGroup_Change()
  mlngGroup = spnGroup.value
  Me.Changed = True
End Sub

Private Sub txtText_Change()
  mstrValue = txtText.Text
  Me.Changed = True
End Sub

Private Sub cboFldTable_Refresh()

  Dim iIndex As Integer
  Dim iDefaultIndex As Integer
  Dim bTableOK As Boolean

  ' Clear the current contents of the combo.
  cboFldTable.Clear
  iIndex = -1
  
  With recTabEdit
    .Index = "idxName"
    
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If

    Do While Not .EOF
      bTableOK = False
           
      If (Not .Fields("deleted")) Then
        If .Fields("tableID") = mlngBaseTableID Then
          bTableOK = True
        Else
          recRelEdit.Index = "idxParentID"
          recRelEdit.Seek "=", .Fields("tableID"), mlngBaseTableID
          
          If Not recRelEdit.NoMatch Then
            bTableOK = True
          End If
        End If
        
        If bTableOK Then
          cboFldTable.AddItem !TableName
          cboFldTable.ItemData(cboFldTable.NewIndex) = !TableID
          
          If !TableID = mlngTableID Then
            iIndex = cboFldTable.NewIndex
          End If
          
          If !TableID = mlngBaseTableID Then
            iDefaultIndex = cboFldTable.NewIndex
          End If
        End If
      End If
      
      .MoveNext
    Loop
  End With
            

  ' Enable the combo if there are items.
  With cboFldTable
    If .ListCount > 0 Then
      .Enabled = Not mbReadOnly
      If iIndex < 0 Then
        If iDefaultIndex >= 0 Then
          iIndex = iDefaultIndex
        Else
          iIndex = 0
        End If
      End If
      cboFldTable.ListIndex = iIndex
    Else
      .Enabled = False
      
      cboFldTable.AddItem "<no tables>"
      cboFldTable.ItemData(cboFldTable.NewIndex) = 0
      cboFldTable.ListIndex = 0
      cboFldColumn_Refresh
    End If
  End With


End Sub


Private Sub cboFldColumn_Refresh()
  ' Refresh the Columns Columns combo.
  Dim iIndex As Integer
  Dim lngTableID As Integer
  Dim bAdd As Boolean
  
  iIndex = 0
  lngTableID = GetComboItem(cboFldTable)
  
  ' Clear columns combo.
  cboFldColumn.Clear
  
  ' Loop through columns for selected lookup table.
  With recColEdit
    .Index = "idxName"
    .Seek ">=", lngTableID
    
    If Not .NoMatch Then
      Do While Not .EOF
        If .Fields("tableID") <> lngTableID Then
          Exit Do
        End If
        
        ' Add each column name to the lookup columns combo.
        ' NB. We only want to add certain types of column. There's not use in
        ' looking up OLE values.
        If (.Fields("columnType") <> giCOLUMNTYPE_SYSTEM) And _
          (.Fields("columnType") <> giCOLUMNTYPE_LINK) And _
          (Not .Fields("deleted")) And _
          (.Fields("dataType") <> dtLONGVARBINARY) And _
          (.Fields("dataType") <> dtVARBINARY) Then
          
'          If mbIsKeyField Then
'            If .Fields("Mandatory") = True Then
'              cboFldColumn.AddItem .Fields("columnName")
'              cboFldColumn.ItemData(cboFldColumn.NewIndex) = .Fields("columnID")
'            End If
'          Else
            cboFldColumn.AddItem .Fields("columnName")
            cboFldColumn.ItemData(cboFldColumn.NewIndex) = .Fields("columnID")
'          End If
      
          If .Fields("columnID") = mlngColumnID Then
            iIndex = cboFldColumn.NewIndex
          End If
        End If
    
        .MoveNext
      Loop
    End If
  End With
  
  ' Enable the combo if there are items.
  With cboFldColumn
    If .ListCount > 0 Then
      .ListIndex = iIndex
    Else
      .Enabled = False
    End If
  End With
  
  Exit Sub
  
End Sub


Private Sub cmdCalculation_Click()
  
  Dim objExpr As CExpression
  Dim fDataTypeChanged As Boolean
  
  ' Instantiate an expression object.
  Set objExpr = New CExpression
  
  With objExpr
    .Initialise mlngTableID, mlngExprID, giEXPR_COLUMNCALCULATION, giEXPRVALUE_CHARACTER
    
    ' Instruct the expression object to display the
    ' expression selection form.
    If .SelectExpression Then
      
      mlngExprID = .ExpressionID
      ' Read the selected expression info.
      GetCalculationExpressionDetails
    Else
      ' Check in case the original expression has been deleted.
      With recExprEdit
        .Index = "idxExprID"
        .Seek "=", mlngExprID, False

        If .NoMatch Then
          ' Read the selected expression info.
          mlngExprID = 0
          GetCalculationExpressionDetails
        End If
        
      End With
    End If
  End With
  
  Set objExpr = Nothing
  Me.Changed = True

End Sub

Private Sub GetCalculationExpressionDetails()
  Dim sExprName As String
  Dim objExpr As CExpression
  
  ' Initialise the default values.
  sExprName = vbNullString
  
  ' Instantiate the expression class.
  Set objExpr = New CExpression
  
  With objExpr
    ' Set the expression id.
    .ExpressionID = mlngExprID
    
    ' Read the required info from the expression.
    If .ReadExpressionDetails Then
      sExprName = .Name
    End If
  End With

  ' Disassociate object variables.
  Set objExpr = Nothing
  
  ' Update the calculation controls properties.
  txtCalculation.Text = sExprName

End Sub

Private Function GetComboItem(cboTemp As ComboBox) As Long
  GetComboItem = 0
  If cboTemp.ListIndex <> -1 Then
    GetComboItem = cboTemp.ItemData(cboTemp.ListIndex)
  End If
End Function

' Load the mapping values for this transfer field
Private Sub LoadMappings()

  Dim sSQL As String
  Dim strAddString As String
  Dim rsDefinition As DAO.Recordset

  sSQL = "SELECT *" & _
    " FROM tmpAccordTransferFieldMappings" & _
    " WHERE TransferID = " & CStr(mlngTransferID) & _
    " AND FieldID = " & CStr(mlngTransferFieldID) & _
    " ORDER BY HRProValue"
    
  Set rsDefinition = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  grdColumnMapping.RemoveAll

  While Not rsDefinition.EOF
    strAddString = IIf(IsNull(rsDefinition!HRProValue), "", rsDefinition!HRProValue) & vbTab & rsDefinition!AccordValue
    grdColumnMapping.AddItem strAddString
    rsDefinition.MoveNext
  Wend
  
  rsDefinition.Close
  Set rsDefinition = Nothing

End Sub

Private Sub SaveMappings()

  Dim iLoop As Integer
  Dim sSQL As String
  Dim varBookmark As Variant

  daoDb.Execute "DELETE FROM tmpAccordTransferFieldMappings" & _
                  " WHERE TransferID = " & CStr(mlngTransferID) & _
                  " AND FieldID = " & CStr(mlngTransferFieldID), dbFailOnError
  
  UI.LockWindow grdColumnMapping.hWnd
  
  With grdColumnMapping
    For iLoop = 0 To .Rows - 1
      .Bookmark = .AddItemBookmark(iLoop)
  
      If Not (Len(.Columns("HRProValue").value) = 0 And Len(.Columns("AccordValue").value) = 0) Then
  
        sSQL = "INSERT INTO tmpAccordTransferFieldMappings" & _
          " (TransferID, FieldID, HRProValue, AccordValue)" & _
          " VALUES (" & _
          CStr(mlngTransferID) & "," & _
          CStr(mlngTransferFieldID) & "," & _
          "'" & Replace(IIf(Len(.Columns("HRProValue").value) = 0, "", .Columns("HRProValue").value), "'", "''") & "'," & _
          "'" & Replace(IIf(Len(.Columns("AccordValue").value) = 0, "", .Columns("AccordValue").value), "'", "''") & "')"
  
        daoDb.Execute sSQL, dbFailOnError
      End If
      
    Next iLoop
  End With

  UI.UnlockWindow

  Application.Changed = True

End Sub

' Validate the selection
Private Function Validate() As Boolean
  
  Dim strMessage As String
  Dim bWarning As Boolean
  
  strMessage = ""
  bWarning = False
  
  ' If a calc make sure something is defined
  If optComponentType(1).value = True And mlngExprID < 1 Then
    strMessage = strMessage & "No calculation selected."
    bWarning = True
  End If

  If bWarning Then
    MsgBox strMessage, vbExclamation, Me.Caption
    Validate = False
  Else
    Validate = True
  End If

End Function

Public Property Let IsDepartmentCode(ByVal pbNewValue As Boolean)
  mbForceAlwaysTransfer = IIf(pbNewValue, True, mbForceAlwaysTransfer)
End Property

Public Property Let IsDepartmentName(ByVal pbNewValue As Boolean)
  mbForceAlwaysTransfer = IIf(pbNewValue, True, mbForceAlwaysTransfer)
End Property

Public Property Let IsPayrollCode(ByVal pbNewValue As Boolean)
  mbForceAlwaysTransfer = IIf(pbNewValue, True, mbForceAlwaysTransfer)
End Property

Public Property Let IsEmployeeName(ByVal pbNewValue As Boolean)
  mbForceAlwaysTransfer = IIf(pbNewValue, True, mbForceAlwaysTransfer)
End Property

Public Property Let Group(ByVal plngNewValue As Long)
  mlngGroup = plngNewValue
End Property

Public Property Get Group() As Long
  Group = mlngGroup
End Property
