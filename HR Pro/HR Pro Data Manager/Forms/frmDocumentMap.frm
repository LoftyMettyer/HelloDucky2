VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.1#0"; "Codejock.Controls.v13.1.0.ocx"
Begin VB.Form frmDocumentMap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Document Map"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9780
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   6855
      Left            =   90
      TabIndex        =   2
      Top             =   45
      Width           =   9555
      _Version        =   851969
      _ExtentX        =   16854
      _ExtentY        =   12091
      _StockProps     =   68
      ItemCount       =   2
      Item(0).Caption =   "Definition"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "fraDefinition"
      Item(0).Control(1)=   "fraClassification"
      Item(0).Control(2)=   "fraDestination"
      Item(1).Caption =   "Advanced"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "fraHeader"
      Begin VB.Frame fraHeader 
         Caption         =   "Header : "
         Height          =   4380
         Left            =   -69910
         TabIndex        =   30
         Top             =   405
         Visible         =   0   'False
         Width           =   9300
         Begin VB.TextBox txtHeader 
            Height          =   3525
            Left            =   225
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   32
            Top             =   630
            Width           =   8835
         End
         Begin VB.CheckBox chkManualHeader 
            Caption         =   "Manual Header"
            Height          =   285
            Left            =   225
            TabIndex        =   31
            Top             =   315
            Width           =   2355
         End
      End
      Begin VB.Frame fraDestination 
         Caption         =   "Destination : "
         Height          =   2625
         Left            =   90
         TabIndex        =   18
         Top             =   4095
         Width           =   9300
         Begin VB.CheckBox chkLockTablesUntilComplete 
            Caption         =   "Lock Tables Until Complete"
            Height          =   240
            Left            =   5310
            TabIndex        =   29
            Top             =   315
            Value           =   1  'Checked
            Width           =   3210
         End
         Begin VB.ComboBox cboParentKeyfield 
            Height          =   315
            Left            =   1845
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1890
            Width           =   3090
         End
         Begin VB.ComboBox cboParentTable 
            Height          =   315
            Left            =   1845
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1485
            Width           =   3090
         End
         Begin VB.ComboBox cboTargetKeyField 
            Height          =   315
            Left            =   1845
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1080
            Width           =   3090
         End
         Begin VB.ComboBox cboTargetColumn 
            Height          =   315
            Left            =   1845
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   675
            Width           =   3090
         End
         Begin VB.ComboBox cboTargetTable 
            Height          =   315
            Left            =   1845
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   270
            Width           =   3090
         End
         Begin VB.Label lblParentKeyfield 
            Caption         =   "Parent Key Field :"
            Height          =   375
            Left            =   225
            TabIndex        =   26
            Top             =   1935
            Width           =   1545
         End
         Begin VB.Label lblParentTable 
            Caption         =   "Parent Table :"
            Height          =   195
            Left            =   225
            TabIndex        =   25
            Top             =   1530
            Width           =   1275
         End
         Begin VB.Label lblTargetTable 
            Caption         =   "Table :"
            Height          =   285
            Left            =   225
            TabIndex        =   24
            Top             =   315
            Width           =   1365
         End
         Begin VB.Label lblTargetColumn 
            Caption         =   "Column : "
            Height          =   330
            Left            =   225
            TabIndex        =   23
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblTargetKeyfield 
            Caption         =   "Key Field :"
            Height          =   240
            Left            =   225
            TabIndex        =   22
            Top             =   1125
            Width           =   1230
         End
      End
      Begin VB.Frame fraClassification 
         Caption         =   "Classification : "
         Height          =   1410
         Left            =   90
         TabIndex        =   13
         Top             =   2520
         Width           =   9300
         Begin VB.ComboBox cboTypes 
            Height          =   315
            Left            =   1845
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   765
            Width           =   3090
         End
         Begin VB.ComboBox cboCategories 
            Height          =   315
            Left            =   1845
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   360
            Width           =   3090
         End
         Begin VB.Label Label2 
            Caption         =   "Type :"
            Height          =   285
            Left            =   315
            TabIndex        =   17
            Top             =   810
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "Category :"
            Height          =   285
            Index           =   0
            Left            =   270
            TabIndex        =   16
            Top             =   420
            Width           =   1095
         End
      End
      Begin VB.Frame fraDefinition 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   90
         TabIndex        =   3
         Top             =   405
         Width           =   9300
         Begin VB.TextBox txtUserName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6225
            MaxLength       =   30
            TabIndex        =   8
            Top             =   315
            Width           =   2820
         End
         Begin VB.OptionButton optReadWrite 
            Caption         =   "Read / &Write"
            Height          =   195
            Left            =   6225
            TabIndex        =   7
            Top             =   765
            Value           =   -1  'True
            Width           =   1650
         End
         Begin VB.OptionButton optReadOnly 
            Caption         =   "&Read Only"
            Height          =   195
            Left            =   6225
            TabIndex        =   6
            Top             =   1155
            Width           =   1470
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   1845
            MaxLength       =   50
            TabIndex        =   5
            Top             =   315
            Width           =   3090
         End
         Begin VB.TextBox txtDesc 
            Height          =   1080
            Left            =   1845
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   705
            Width           =   3090
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Owner :"
            Height          =   195
            Index           =   4
            Left            =   5325
            TabIndex        =   12
            Top             =   360
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            Height          =   195
            Index           =   2
            Left            =   315
            TabIndex        =   11
            Top             =   365
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description :"
            Height          =   195
            Index           =   1
            Left            =   315
            TabIndex        =   10
            Top             =   765
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Access :"
            Height          =   195
            Index           =   3
            Left            =   5325
            TabIndex        =   9
            Top             =   765
            Width           =   600
         End
      End
   End
   Begin XtremeSuiteControls.PushButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   8460
      TabIndex        =   1
      Top             =   7065
      Width           =   1200
      _Version        =   851969
      _ExtentX        =   2117
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Cancel"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdOK 
      Height          =   375
      Left            =   7200
      TabIndex        =   0
      Top             =   7065
      Width           =   1200
      _Version        =   851969
      _ExtentX        =   2117
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&OK"
      UseVisualStyle  =   -1  'True
   End
End
Attribute VB_Name = "frmDocumentMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngDocumentMapID As Long

Private mblnCancelled As Boolean
Private mblnReadOnly As Boolean
Private mblnFromCopy As Boolean
Private mblnDefinitionCreator As Boolean
Private mbChanged As Boolean

Private mdatData As HRProDataMgr.clsDataAccess
Private mclsGeneral As HRProDataMgr.clsGeneral

Private Const SQLTableDef = "ASRSysDocumentMapping"

' Module Setup Constants
Private Const MODULEKEY_DOCMANAGEMENT = "MODULE_DOCUMENTMANAGEMENT"
Private Const PARAMETERKEY_DOCMAN_CATEGORYTABLE = "Param_DocmanCatageoryTable"
Private Const PARAMETERKEY_DOCMAN_CATEGORYCOLUMN = "Param_DocManCatageoryColumn"
Private Const PARAMETERKEY_DOCMAN_TYPETABLE = "Param_DocmanTypeTable"
Private Const PARAMETERKEY_DOCMAN_TYPECOLUMN = "Param_DocManTypeColumn"
Private Const PARAMETERKEY_DOCMAN_TYPECATEGORYCOLUMN = "Param_DocManTypeCategoryColumn"


Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Public Property Let Cancelled(ByVal bCancel As Boolean)
  mblnCancelled = bCancel
End Property

Public Property Get SelectedID() As Long
  SelectedID = mlngDocumentMapID
End Property

Public Sub PrintDefinition(plngDocumentMapID)

End Sub


Public Function Initialise(bNew As Boolean, bCopy As Boolean, Optional lngDocumentMapID As Long, Optional bPrint As Boolean) As Boolean

  Set mdatData = New HRProDataMgr.clsDataAccess
  Set mclsGeneral = New HRProDataMgr.clsGeneral

  PopulateCategoriesCombo
  LoadTableCombo cboTargetTable

  If Not bNew Then
    mlngDocumentMapID = lngDocumentMapID
    
    RetreiveDefinition
  End If
  
  mbChanged = False
  
End Function

Private Function GetDefinition() As ADODB.Recordset

  Dim strSQL As String

  strSQL = "SELECT * FROM " & SQLTableDef & " " & _
           "WHERE [DocumentMapID] = " & CStr(mlngDocumentMapID)
  Set GetDefinition = mdatData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

End Function

Private Sub RetreiveDefinition()

  Dim bOK As Boolean
  Dim rsTemp As ADODB.Recordset

  Set rsTemp = GetDefinition
  If rsTemp.BOF And rsTemp.EOF Then
    MsgBox "This definition has been deleted by another user.", vbExclamation + vbOKOnly, "Label Definition"
    bOK = False
    Exit Sub
  End If

  txtDesc.Text = IIf(rsTemp!Description <> vbNullString, rsTemp!Description, vbNullString)
  
  If mblnFromCopy Then
    txtName.Text = "Copy of " & rsTemp!Name
    txtUserName = gsUserName
    mblnDefinitionCreator = True
  Else
    txtName.Text = rsTemp!Name
    txtUserName = StrConv(rsTemp!UserName, vbProperCase)
    mblnDefinitionCreator = (LCase$(rsTemp!UserName) = LCase$(gsUserName))
  End If

  mblnReadOnly = Not datGeneral.SystemPermission("LABELDEFINITION", "EDIT")
    
  SetComboItem cboCategories, IIf(IsNull(rsTemp.Fields("CategoryRecordID").Value), 0, rsTemp.Fields("CategoryRecordID").Value)
  SetComboItem cboTypes, IIf(IsNull(rsTemp.Fields("TypeRecordID").Value), 0, rsTemp.Fields("TypeRecordID").Value)
    
  SetComboItem cboTargetTable, IIf(IsNull(rsTemp.Fields("TargetTableID").Value), 0, rsTemp.Fields("TargetTableID").Value)
  SetComboItem cboTargetKeyField, IIf(IsNull(rsTemp.Fields("TargetKeyFieldColumnID").Value), 0, rsTemp.Fields("TargetKeyFieldColumnID").Value)
  SetComboItem cboTargetColumn, IIf(IsNull(rsTemp.Fields("TargetColumnID").Value), 0, rsTemp.Fields("TargetColumnID").Value)

  SetComboItem cboParentTable, IIf(IsNull(rsTemp.Fields("ParentTableID").Value), 0, rsTemp.Fields("ParentTableID").Value)
  SetComboItem cboParentKeyfield, IIf(IsNull(rsTemp.Fields("ParentKeyFieldID").Value), 0, rsTemp.Fields("ParentKeyFieldID").Value)

  chkManualHeader.Value = IIf(IsNull(rsTemp!ManualHeader), vbUnchecked, Abs(rsTemp!ManualHeader))
  txtHeader.Text = IIf(IsNull(rsTemp!HeaderText), vbNullString, rsTemp!HeaderText)
  RefreshHeaderText

End Sub

Private Sub RefreshHeaderText()

  Dim sHeader As String
  
  If chkManualHeader.Value = vbUnchecked Then
  
    'If cboTypes.Text = "<None>" Then
  
    sHeader = "<<category='" & cboCategories.Text & "'; type='" & cboTypes.Text & "'" & vbNewLine
    sHeader = sHeader & ";wordmergefield{{" & datGeneral.GetColumnName(GetComboItem(cboTargetKeyField), True) & "}}>>"
    
    txtHeader.Text = sHeader
    
  End If

End Sub


Private Function SaveDefinition() As Boolean

  Dim rsTemp As ADODB.Recordset
  Dim sSQL As String
  Dim bOK As Boolean
  
  bOK = True
  RefreshHeaderText
  
  
  
  
  
  If mlngDocumentMapID > 0 Then
    sSQL = "UPDATE dbo.[" & SQLTableDef & "] SET " _
              & "[Name] = '" & Replace(txtName.Text, "'", "''") & "', " _
              & "[Description] = '" & Replace(txtDesc.Text, "'", "''") & "', " _
              & "[Access] = " & IIf(optReadOnly.Value = True, "'RO'", "'RW'") & ", " _
              & "[TargetTableID] = " & GetComboItem(cboTargetTable) & ", " _
              & "[TargetKeyFieldColumnID] = " & GetComboItem(cboTargetKeyField) & ", " _
              & "[TargetColumnID] = " & GetComboItem(cboTargetColumn) & ", " _
              & "[ManualHeader] = " & CStr(Abs(chkManualHeader.Value <> 0)) & ", " _
              & "[HeaderText] = '" & Replace(txtHeader.Text, "'", "''") & "', " _
              & "[ParentTableID] = " & GetComboItem(cboParentTable) & ", " _
              & "[ParentKeyFieldID] = " & GetComboItem(cboParentKeyfield) & ", " _
              & "[CategoryRecordID] = " & GetComboItem(cboCategories) & ", " _
              & "[TypeRecordID] = " & GetComboItem(cboTypes) _
              & " WHERE [DocumentMapID] = " & CStr(mlngDocumentMapID) & ";"
    gADOCon.Execute sSQL, , adCmdText
    Call UtilUpdateLastSaved(utlDocumentMapping, mlngDocumentMapID)
               
  Else
  
    sSQL = "INSERT " & SQLTableDef & " (" _
              & " [Name], [Description]," _
              & " [UserName], [Access], " _
              & " [TargetTableID], [TargetKeyFieldColumnID], [TargetColumnID], [ParentTableID], [ParentKeyFieldID]," _
              & " [CategoryRecordID], [TypeRecordID], [ManualHeader], [HeaderText]) " _
              & " VALUES('" _
              & Replace(txtName.Text, "'", "''") & "', '" & Replace(txtDesc.Text, "'", "''") _
              & "', '" & datGeneral.UserNameForSQL & "', " & IIf(optReadOnly.Value = True, "'RO'", "'RW'") _
              & ", " & GetComboItem(cboTargetTable) & ", " & GetComboItem(cboTargetKeyField) & ", " & GetComboItem(cboTargetColumn) _
              & ", " & GetComboItem(cboParentTable) & ", " & GetComboItem(cboParentKeyfield) _
              & ", " & GetComboItem(cboCategories) & ", " & GetComboItem(cboTypes) _
              & ", " & CStr(Abs(chkManualHeader.Value <> 0)) & ", '" & Replace(txtHeader.Text, "'", "''") & "');"

    mlngDocumentMapID = InsertDocumentMap(sSQL)
    Call UtilCreated(utlDocumentMapping, mlngDocumentMapID)
  
  End If
  
  SaveDefinition = bOK

End Function

Private Sub cboCategories_Click()
  PopulateTypesCombo
  RefreshHeaderText
End Sub

Private Sub cboParentTable_Click()

  Dim rsCols As ADODB.Recordset

  ' Columns for the parent table
  If cboParentTable.Enabled Then
    Set rsCols = mclsGeneral.GetColumnNames(GetComboItem(cboParentTable))
    Do While Not rsCols.EOF
    
      With cboParentKeyfield
        .AddItem rsCols!ColumnName
        .ItemData(.NewIndex) = rsCols!ColumnID
      End With
      
      rsCols.MoveNext
    Loop
    rsCols.Close
  End If

End Sub

Private Sub cboTargetColumn_Click()
  RefreshHeaderText
End Sub

Private Sub cboTargetKeyField_Click()
  RefreshHeaderText
End Sub

Private Sub cboTargetTable_Click()

  Dim sSQL As String
  Dim rsCols As ADODB.Recordset
  Dim rsParents As ADODB.Recordset
  Dim rsTables As ADODB.Recordset
  Dim bHasParent As Boolean
  
  Screen.MousePointer = vbHourglass
  
  cboParentTable.Clear
  cboParentKeyfield.Clear
  cboTargetColumn.Clear
  cboTargetKeyField.Clear
  
  'Get all the columns for the selected table
  Set rsCols = datGeneral.GetColumnNames(cboTargetTable.ItemData(cboTargetTable.ListIndex))
  Do While Not rsCols.EOF
  
    With cboTargetColumn
      .AddItem rsCols!ColumnName
      .ItemData(.NewIndex) = rsCols!ColumnID
    End With
    
    With cboTargetKeyField
      .AddItem rsCols!ColumnName
      .ItemData(.NewIndex) = rsCols!ColumnID
    End With
    
    rsCols.MoveNext
  Loop
  rsCols.Close


  ' Get parent tables
  sSQL = "SELECT [ParentID] FROM dbo.[ASRSysRelations] WHERE [ChildID] = " & cboTargetTable.ItemData(cboTargetTable.ListIndex) & ";"
  Set rsParents = mdatData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Do While Not rsParents.EOF
    sSQL = "SELECT TableName, TableID FROM ASRSysTables " & _
           "WHERE TableID = " & rsParents!ParentID & " " & _
           ""
    Set rsTables = mdatData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    Do While Not rsTables.EOF
      cboParentTable.AddItem rsTables!TableName
      cboParentTable.ItemData(cboParentTable.NewIndex) = rsTables!TableID
      rsTables.MoveNext
    Loop
    rsParents.MoveNext
    rsTables.Close
  Loop
  rsParents.Close

  bHasParent = cboParentTable.ListCount > 0
  
  EnableControl lblParentTable, bHasParent
  EnableControl cboParentTable, bHasParent
  EnableControl lblParentKeyfield, bHasParent
  EnableControl cboParentKeyfield, bHasParent
  
  RefreshHeaderText

  Set rsParents = Nothing
  Set rsTables = Nothing
  Set rsCols = Nothing
  Screen.MousePointer = vbNormal

End Sub

Private Sub cboTypes_Click()
  RefreshHeaderText
End Sub

Private Sub chkManualHeader_Click()
  txtHeader.Enabled = (chkManualHeader.Value = vbChecked)
  RefreshHeaderText
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdOK_Click()

  Dim fOK As Boolean
  
  Screen.MousePointer = vbHourglass
    
  If Not SaveDefinition Then
    Screen.MousePointer = vbNormal
    Exit Sub
  End If
  
  Screen.MousePointer = vbNormal
  
  Me.Hide

End Sub


Private Function InsertDocumentMap(pstrSQL As String) As Long

  ' Insert definition into the name table and return the ID.

  On Error GoTo ErrorTrap

  Dim sSQL As String
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim fSavedOK As Boolean
  
  fSavedOK = True
  
  Set cmADO = New ADODB.Command
  
  With cmADO
    .CommandText = "sp_ASRInsertNewUtility"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
  
    Set .ActiveConnection = gADOCon
              
    Set pmADO = .CreateParameter("newID", adInteger, adParamOutput)
    .Parameters.Append pmADO
            
    Set pmADO = .CreateParameter("insertString", adLongVarChar, adParamInput, -1)
    .Parameters.Append pmADO
    pmADO.Value = pstrSQL
              
    Set pmADO = .CreateParameter("tablename", adVarChar, adParamInput, 255)
    .Parameters.Append pmADO
    pmADO.Value = SQLTableDef
              
    Set pmADO = .CreateParameter("idcolumnname", adVarChar, adParamInput, 30)
    .Parameters.Append pmADO
    pmADO.Value = "DocumentMapID"
              
    Set pmADO = Nothing
            
    cmADO.Execute
              
    If Not fSavedOK Then
      MsgBox "The new record could not be created." & vbCrLf & vbCrLf & _
        Err.Description, vbOKOnly + vbExclamation, App.ProductName
        InsertDocumentMap = 0
        Set cmADO = Nothing
        Exit Function
    End If
    
    InsertDocumentMap = IIf(IsNull(.Parameters(0).Value), 0, .Parameters(0).Value)
          
  End With
  
  Set cmADO = Nothing

  Exit Function
  
ErrorTrap:
  
  fSavedOK = False
  Resume Next
  
End Function

Private Sub Form_Load()

  ' Get rid of the icon off the form
  RemoveIcon Me

End Sub

Private Sub txtName_Change()
  mbChanged = True
End Sub

Private Sub PopulateCategoriesCombo()

  Dim rsCategories As ADODB.Recordset
  Dim sTableName As String
  Dim sColumnName As String
  
  sTableName = mclsGeneral.GetTableName(Val(GetModuleParameter(MODULEKEY_DOCMANAGEMENT, PARAMETERKEY_DOCMAN_CATEGORYTABLE)))
  sColumnName = mclsGeneral.GetColumnName(Val(GetModuleParameter(MODULEKEY_DOCMANAGEMENT, PARAMETERKEY_DOCMAN_CATEGORYCOLUMN)))

  cboCategories.Clear
  cboCategories.AddItem "<None>"
  cboCategories.ItemData(cboCategories.NewIndex) = 0

  Set rsCategories = mclsGeneral.GetRecords("SELECT DISTINCT [id], [" & sColumnName & "] FROM dbo.[" & sTableName & "]")
  With rsCategories
    Do While Not .EOF
    
      cboCategories.AddItem .Fields(1).Value
      cboCategories.ItemData(cboCategories.NewIndex) = .Fields(0).Value
      
      .MoveNext
    Loop
    
    .Close
  End With

  Set rsCategories = Nothing

  PopulateTypesCombo

End Sub

Private Sub PopulateTypesCombo()

  Dim rsTypes As ADODB.Recordset
  Dim sTableName As String
  Dim sColumnName As String
  Dim sSolumnCategoryName As String
  
  sTableName = mclsGeneral.GetTableName(Val(GetModuleParameter(MODULEKEY_DOCMANAGEMENT, PARAMETERKEY_DOCMAN_TYPETABLE)))
  sColumnName = mclsGeneral.GetColumnName(Val(GetModuleParameter(MODULEKEY_DOCMANAGEMENT, PARAMETERKEY_DOCMAN_TYPECOLUMN)))
  sSolumnCategoryName = mclsGeneral.GetColumnName(Val(GetModuleParameter(MODULEKEY_DOCMANAGEMENT, PARAMETERKEY_DOCMAN_TYPECATEGORYCOLUMN)))

  cboTypes.Clear
  cboTypes.AddItem "<None>"
  cboTypes.ItemData(cboTypes.NewIndex) = 0

  Set rsTypes = mclsGeneral.GetRecords("SELECT [id], [" & sColumnName & "] FROM dbo.[" & sTableName & "] " & _
                      "WHERE [" & sSolumnCategoryName & "] = '" & Replace(cboCategories.Text, "'", "''") & "'")
  With rsTypes
  
    Do While Not .EOF
      cboTypes.AddItem (.Fields(1).Value)
      cboTypes.ItemData(cboTypes.NewIndex) = .Fields(0).Value
      
      .MoveNext
    Loop
    
    .Close
  End With

  Set rsTypes = Nothing

End Sub

