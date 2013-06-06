VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.1#0"; "Codejock.Controls.v13.1.0.ocx"
Begin VB.Form frmDocumentMap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Document Management Type"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10755
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDocumentMap.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   6945
      Left            =   90
      TabIndex        =   2
      Top             =   45
      Width           =   10545
      _Version        =   851969
      _ExtentX        =   18600
      _ExtentY        =   12250
      _StockProps     =   68
      ItemCount       =   2
      Item(0).Caption =   "Definition"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "fraDefinition"
      Item(0).Control(1)=   "fraClassification"
      Item(0).Control(2)=   "fraDestination"
      Item(1).Caption =   "Advanced"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "fraHeader"
      Item(1).Control(1)=   "chkLockTablesUntilComplete"
      Begin VB.CheckBox chkLockTablesUntilComplete 
         Caption         =   "Lock Tables Until Complete"
         Height          =   240
         Left            =   -69865
         TabIndex        =   32
         Top             =   4995
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   3210
      End
      Begin VB.Frame fraHeader 
         Caption         =   "Header : "
         Height          =   4380
         Left            =   -69910
         TabIndex        =   29
         Top             =   405
         Visible         =   0   'False
         Width           =   10335
         Begin VB.TextBox txtHeader 
            Height          =   3525
            Left            =   225
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   31
            Top             =   630
            Width           =   9915
         End
         Begin VB.CheckBox chkManualHeader 
            Caption         =   "Manual Header"
            Height          =   285
            Left            =   225
            TabIndex        =   30
            Top             =   315
            Width           =   2355
         End
      End
      Begin VB.Frame fraDestination 
         Caption         =   "Destination : "
         Height          =   2715
         Left            =   90
         TabIndex        =   18
         Top             =   4095
         Width           =   10335
         Begin VB.ComboBox cboParent2Table 
            Height          =   315
            Left            =   7065
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   1080
            Width           =   3090
         End
         Begin VB.ComboBox cboParent2Keyfield 
            Height          =   315
            Left            =   7065
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   1485
            Width           =   3090
         End
         Begin VB.ComboBox cboTargetCategory 
            Height          =   315
            Left            =   1845
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   1485
            Width           =   3090
         End
         Begin VB.ComboBox cboTargetType 
            Height          =   315
            Left            =   1845
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   1890
            Width           =   3090
         End
         Begin VB.ComboBox cboTargetGUID 
            Height          =   315
            Left            =   1845
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   2295
            Width           =   3090
         End
         Begin VB.ComboBox cboParent1Keyfield 
            Height          =   315
            Left            =   7065
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   675
            Width           =   3090
         End
         Begin VB.ComboBox cboParent1Table 
            Height          =   315
            Left            =   7065
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   270
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
         Begin VB.Label lblParent2Table 
            Caption         =   "Parent 2 Table :"
            Height          =   195
            Left            =   5310
            TabIndex        =   42
            Top             =   1125
            Width           =   1725
         End
         Begin VB.Label lblParent2Keyfield 
            Caption         =   "Parent 2 Key Field :"
            Height          =   375
            Left            =   5310
            TabIndex        =   41
            Top             =   1530
            Width           =   1815
         End
         Begin VB.Label lblDestinationGUID 
            Caption         =   "Unique ID :"
            Height          =   195
            Left            =   225
            TabIndex        =   35
            Top             =   2340
            Width           =   1140
         End
         Begin VB.Label lblDestinationType 
            Caption         =   "Type : "
            Height          =   240
            Left            =   225
            TabIndex        =   34
            Top             =   1935
            Width           =   1095
         End
         Begin VB.Label lblDestinationCategory 
            Caption         =   "Category :"
            Height          =   285
            Left            =   225
            TabIndex        =   33
            Top             =   1530
            Width           =   1410
         End
         Begin VB.Label lblParent1Keyfield 
            Caption         =   "Parent 1 Key Field :"
            Height          =   375
            Left            =   5310
            TabIndex        =   26
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label lblParent1Table 
            Caption         =   "Parent 1 Table :"
            Height          =   195
            Left            =   5310
            TabIndex        =   25
            Top             =   315
            Width           =   1725
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
         Width           =   10335
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
         Width           =   10335
         Begin VB.TextBox txtUserName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   7065
            MaxLength       =   30
            TabIndex        =   8
            Top             =   315
            Width           =   3045
         End
         Begin VB.OptionButton optReadWrite 
            Caption         =   "Read / &Write"
            Height          =   195
            Left            =   7065
            TabIndex        =   7
            Top             =   765
            Value           =   -1  'True
            Width           =   1650
         End
         Begin VB.OptionButton optReadOnly 
            Caption         =   "&Read Only"
            Height          =   195
            Left            =   7065
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
            Left            =   5280
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
            Left            =   5280
            TabIndex        =   9
            Top             =   765
            Width           =   600
         End
      End
   End
   Begin XtremeSuiteControls.PushButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   9450
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
      Left            =   8190
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

Private Const SQLTableDef = "ASRSysDocumentManagementTypes"

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

  mblnReadOnly = Not datGeneral.SystemPermission("VERSION1", "EDIT")
    
  SetComboItem cboCategories, IIf(IsNull(rsTemp.Fields("CategoryRecordID").Value), 0, rsTemp.Fields("CategoryRecordID").Value)
  SetComboItem cboTypes, IIf(IsNull(rsTemp.Fields("TypeRecordID").Value), 0, rsTemp.Fields("TypeRecordID").Value)
    
  SetComboItem cboTargetTable, IIf(IsNull(rsTemp.Fields("TargetTableID").Value), 0, rsTemp.Fields("TargetTableID").Value)
  SetComboItem cboTargetKeyField, IIf(IsNull(rsTemp.Fields("TargetKeyFieldColumnID").Value), 0, rsTemp.Fields("TargetKeyFieldColumnID").Value)
  SetComboItem cboTargetColumn, IIf(IsNull(rsTemp.Fields("TargetColumnID").Value), 0, rsTemp.Fields("TargetColumnID").Value)
  SetComboItem cboTargetCategory, IIf(IsNull(rsTemp.Fields("TargetCategoryColumnID").Value), 0, rsTemp.Fields("TargetCategoryColumnID").Value)
  SetComboItem cboTargetType, IIf(IsNull(rsTemp.Fields("TargetTypeColumnID").Value), 0, rsTemp.Fields("TargetTypeColumnID").Value)
  SetComboItem cboTargetGUID, IIf(IsNull(rsTemp.Fields("TargetGUIDColumnID").Value), 0, rsTemp.Fields("TargetGUIDColumnID").Value)

  SetComboItem cboParent1Table, IIf(IsNull(rsTemp.Fields("Parent1TableID").Value), 0, rsTemp.Fields("Parent1TableID").Value)
  SetComboItem cboParent1Keyfield, IIf(IsNull(rsTemp.Fields("Parent1KeyFieldColumnID").Value), 0, rsTemp.Fields("Parent1KeyFieldColumnID").Value)

  SetComboItem cboParent2Table, IIf(IsNull(rsTemp.Fields("Parent2TableID").Value), 0, rsTemp.Fields("Parent2TableID").Value)
  SetComboItem cboParent2Keyfield, IIf(IsNull(rsTemp.Fields("Parent2KeyFieldColumnID").Value), 0, rsTemp.Fields("Parent2KeyFieldColumnID").Value)


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
              & "[TargetCategoryColumnID] = " & GetComboItem(cboTargetType) & ", " _
              & "[TargetTypeColumnID] = " & GetComboItem(cboTargetCategory) & ", " _
              & "[TargetGUIDColumnID] = " & GetComboItem(cboTargetGUID) & ", " _
              & "[ManualHeader] = " & CStr(Abs(chkManualHeader.Value <> 0)) & ", " _
              & "[HeaderText] = '" & Replace(txtHeader.Text, "'", "''") & "', " _
              & "[Parent1TableID] = " & GetComboItem(cboParent1Table) & ", " _
              & "[Parent1KeyFieldColumnID] = " & GetComboItem(cboParent1Keyfield) & ", " _
              & "[Parent2TableID] = " & GetComboItem(cboParent2Table) & ", " _
              & "[Parent2KeyFieldColumnID] = " & GetComboItem(cboParent2Keyfield) & ", " _
              & "[CategoryRecordID] = " & GetComboItem(cboCategories) & ", " _
              & "[TypeRecordID] = " & GetComboItem(cboTypes) _
              & " WHERE [DocumentMapID] = " & CStr(mlngDocumentMapID) & ";"
    gADOCon.Execute sSQL, , adCmdText
    Call UtilUpdateLastSaved(utlDocumentMapping, mlngDocumentMapID)
               
  Else
  
    sSQL = "INSERT " & SQLTableDef & " (" _
              & " [Name], [Description]," _
              & " [UserName], [Access], " _
              & " [TargetTableID], [TargetKeyFieldColumnID], [TargetColumnID], [Parent1TableID], [Parent1KeyFieldColumnID], [Parent2TableID], [Parent2KeyFieldColumnID]," _
              & " [TargetCategoryColumnID], [TargetTypeColumnID], [TargetGUIDColumnID], [CategoryRecordID], [TypeRecordID], [ManualHeader], [HeaderText]) " _
              & " VALUES('" _
              & Replace(txtName.Text, "'", "''") & "', '" & Replace(txtDesc.Text, "'", "''") _
              & "', '" & datGeneral.UserNameForSQL & "', " & IIf(optReadOnly.Value = True, "'RO'", "'RW'") _
              & ", " & GetComboItem(cboTargetTable) & ", " & GetComboItem(cboTargetKeyField) & ", " & GetComboItem(cboTargetColumn) _
              & ", " & GetComboItem(cboParent1Table) & ", " & GetComboItem(cboParent1Keyfield) _
              & ", " & GetComboItem(cboParent2Table) & ", " & GetComboItem(cboParent2Keyfield) _
              & ", " & GetComboItem(cboTargetCategory) & ", " & GetComboItem(cboTargetType) & ", " & GetComboItem(cboTargetGUID) _
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

Private Sub cboParent1Table_Click()

  Dim rsCols As ADODB.Recordset

  ' Columns for the parent table
  If cboParent1Table.Enabled Then
    Set rsCols = mclsGeneral.GetColumnNames(GetComboItem(cboParent1Table))
    Do While Not rsCols.EOF
    
      With cboParent1Keyfield
        .AddItem rsCols!ColumnName
        .ItemData(.NewIndex) = rsCols!ColumnID
      End With
      
      rsCols.MoveNext
    Loop
    rsCols.Close
  End If

End Sub

Private Sub cboParent2Table_Click()

  Dim rsCols As ADODB.Recordset

  ' Columns for the parent table
  If cboParent2Table.Enabled Then
    Set rsCols = mclsGeneral.GetColumnNames(GetComboItem(cboParent2Table))
    Do While Not rsCols.EOF
    
      With cboParent2Keyfield
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
  Dim iParents As Integer
  
  Screen.MousePointer = vbHourglass
  
  cboParent1Table.Clear
  cboParent1Keyfield.Clear
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
    
    With cboTargetCategory
      .AddItem rsCols!ColumnName
      .ItemData(.NewIndex) = rsCols!ColumnID
    End With
    
    With cboTargetType
      .AddItem rsCols!ColumnName
      .ItemData(.NewIndex) = rsCols!ColumnID
    End With
    
    With cboTargetGUID
      .AddItem rsCols!ColumnName
      .ItemData(.NewIndex) = rsCols!ColumnID
    End With
    
    rsCols.MoveNext
  Loop
  rsCols.Close


  ' Get parent tables
  sSQL = "SELECT [ParentID] FROM dbo.[ASRSysRelations] WHERE [ChildID] = " & cboTargetTable.ItemData(cboTargetTable.ListIndex) & ";"
  Set rsParents = mdatData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    
  iParents = rsParents.RecordCount

  Do While Not rsParents.EOF
    sSQL = "SELECT TableName, TableID FROM ASRSysTables " & _
           "WHERE TableID = " & rsParents!ParentID & " " & _
           ""
    Set rsTables = mdatData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    Do While Not rsTables.EOF
      cboParent1Table.AddItem rsTables!TableName
      cboParent1Table.ItemData(cboParent1Table.NewIndex) = rsTables!TableID
      
      cboParent2Table.AddItem rsTables!TableName
      cboParent2Table.ItemData(cboParent1Table.NewIndex) = rsTables!TableID
      
      rsTables.MoveNext
    Loop
    rsParents.MoveNext
    rsTables.Close
  Loop
  rsParents.Close

  
  EnableControl lblParent1Table, (iParents > 0)
  EnableControl cboParent1Table, (iParents > 0)
  EnableControl lblParent1Keyfield, (iParents > 0)
  EnableControl cboParent1Keyfield, (iParents > 0)
  
  EnableControl lblParent2Table, (iParents > 1)
  EnableControl cboParent2Table, (iParents > 1)
  EnableControl lblParent2Keyfield, (iParents > 1)
  EnableControl cboParent2Keyfield, (iParents > 1)
  
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

