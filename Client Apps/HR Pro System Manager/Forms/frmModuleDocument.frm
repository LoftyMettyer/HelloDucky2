VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmModuleDocument 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Document Management"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10065
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmModuleDocument.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   5640
      Left            =   90
      TabIndex        =   2
      Top             =   45
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   9948
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Types"
      TabPicture(0)   =   "frmModuleDocument.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraTypes"
      Tab(0).Control(1)=   "fraComponent(5)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Mail Merge"
      TabPicture(1)   =   "frmModuleDocument.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblTransferTable"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblColumnName"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "grdTransferDetails(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdDelete"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdEdit"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cboTransferTables"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtCategory"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.TextBox txtCategory 
         Height          =   315
         Left            =   930
         MaxLength       =   30
         TabIndex        =   20
         Text            =   "txtColName"
         Top             =   945
         Width           =   7110
      End
      Begin VB.ComboBox cboTransferTables 
         Height          =   315
         Left            =   945
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   540
         Width           =   3255
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit..."
         Enabled         =   0   'False
         Height          =   400
         Left            =   6885
         TabIndex        =   16
         Top             =   1665
         Width           =   1200
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Cle&ar"
         Enabled         =   0   'False
         Height          =   400
         Left            =   6885
         TabIndex        =   15
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Frame fraTypes 
         Caption         =   "Types : "
         Height          =   1755
         Left            =   -74865
         TabIndex        =   8
         Top             =   1935
         Width           =   5865
         Begin VB.ComboBox cboTypeTable 
            Height          =   315
            Left            =   2430
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   315
            Width           =   3255
         End
         Begin VB.ComboBox cboTypeColumn 
            Height          =   315
            Left            =   2430
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Tag             =   "cboTypeTable"
            Top             =   1215
            Width           =   3255
         End
         Begin VB.ComboBox cboTypeCategoryColumn 
            Height          =   315
            Left            =   2430
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Tag             =   "cboTypeTable"
            Top             =   765
            Width           =   3255
         End
         Begin VB.Label lblTypeTable 
            Caption         =   "Type Table : "
            Height          =   285
            Left            =   195
            TabIndex        =   14
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblTypeColumn 
            Caption         =   "Type Column : "
            Height          =   285
            Left            =   195
            TabIndex        =   13
            Top             =   1260
            Width           =   1410
         End
         Begin VB.Label lblTypeCategoryColumn 
            Caption         =   "Type Category Column : "
            Height          =   285
            Left            =   195
            TabIndex        =   12
            Top             =   810
            Width           =   2130
         End
      End
      Begin VB.Frame fraComponent 
         Caption         =   "Categories :"
         Height          =   1305
         Index           =   5
         Left            =   -74865
         TabIndex        =   3
         Tag             =   "6"
         Top             =   450
         Width           =   5865
         Begin VB.ComboBox cboCategoryTable 
            Height          =   315
            Left            =   2430
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   360
            Width           =   3255
         End
         Begin VB.ComboBox cboCategoryColumn 
            Height          =   315
            Left            =   2430
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Tag             =   "cboCategoryTable"
            Top             =   750
            Width           =   3255
         End
         Begin VB.Label lblCategoryTable 
            Caption         =   "Category Table : "
            Height          =   285
            Left            =   225
            TabIndex        =   7
            Top             =   405
            Width           =   1680
         End
         Begin VB.Label lblCatgeoryColumn 
            Caption         =   "Category Column : "
            Height          =   330
            Left            =   225
            TabIndex        =   6
            Top             =   795
            Width           =   1815
         End
      End
      Begin SSDataWidgets_B.SSDBGrid grdTransferDetails 
         Height          =   3255
         Index           =   0
         Left            =   135
         TabIndex        =   17
         Top             =   1665
         Width           =   6510
         _Version        =   196617
         DataMode        =   2
         RecordSelectors =   0   'False
         GroupHeaders    =   0   'False
         Col.Count       =   21
         stylesets.count =   2
         stylesets(0).Name=   "KeyField"
         stylesets(0).BackColor=   14024703
         stylesets(0).HasFont=   -1  'True
         BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(0).Picture=   "frmModuleDocument.frx":0044
         stylesets(1).Name=   "Mandatory"
         stylesets(1).BackColor=   15400959
         stylesets(1).HasFont=   -1  'True
         BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(1).Picture=   "frmModuleDocument.frx":0060
         AllowUpdate     =   0   'False
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
         Columns.Count   =   21
         Columns(0).Width=   5292
         Columns(0).Caption=   "Transfer Field"
         Columns(0).Name =   "Description"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   5741
         Columns(1).Caption=   "HR Pro Value"
         Columns(1).Name =   "Display_MapToValue"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3200
         Columns(2).Visible=   0   'False
         Columns(2).Caption=   "ASRMapType"
         Columns(2).Name =   "ASRMapType"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   3200
         Columns(3).Visible=   0   'False
         Columns(3).Caption=   "ASRTableID"
         Columns(3).Name =   "ASRTableID"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   3200
         Columns(4).Visible=   0   'False
         Columns(4).Caption=   "ASRColumnID"
         Columns(4).Name =   "ASRColumnID"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   3200
         Columns(5).Visible=   0   'False
         Columns(5).Caption=   "ASRExprID"
         Columns(5).Name =   "ASRExprID"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   3200
         Columns(6).Visible=   0   'False
         Columns(6).Caption=   "ASRValue"
         Columns(6).Name =   "ASRValue"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         Columns(7).Width=   3200
         Columns(7).Visible=   0   'False
         Columns(7).Caption=   "Mandatory"
         Columns(7).Name =   "Mandatory"
         Columns(7).DataField=   "Column 7"
         Columns(7).DataType=   8
         Columns(7).FieldLen=   256
         Columns(8).Width=   3200
         Columns(8).Visible=   0   'False
         Columns(8).Caption=   "TransferFieldID"
         Columns(8).Name =   "TransferFieldID"
         Columns(8).DataField=   "Column 8"
         Columns(8).DataType=   8
         Columns(8).FieldLen=   256
         Columns(9).Width=   3200
         Columns(9).Visible=   0   'False
         Columns(9).Caption=   "IsCompanyCode"
         Columns(9).Name =   "IsCompanyCode"
         Columns(9).DataField=   "Column 9"
         Columns(9).DataType=   8
         Columns(9).FieldLen=   256
         Columns(10).Width=   3200
         Columns(10).Visible=   0   'False
         Columns(10).Caption=   "IsEmployeeCode"
         Columns(10).Name=   "IsEmployeeCode"
         Columns(10).DataField=   "Column 10"
         Columns(10).DataType=   8
         Columns(10).FieldLen=   256
         Columns(11).Width=   3200
         Columns(11).Visible=   0   'False
         Columns(11).Caption=   "Direction"
         Columns(11).Name=   "Direction"
         Columns(11).DataField=   "Column 11"
         Columns(11).DataType=   8
         Columns(11).FieldLen=   256
         Columns(12).Width=   3200
         Columns(12).Visible=   0   'False
         Columns(12).Caption=   "IsKeyField"
         Columns(12).Name=   "IsKeyField"
         Columns(12).DataField=   "Column 12"
         Columns(12).DataType=   17
         Columns(12).FieldLen=   256
         Columns(13).Width=   3200
         Columns(13).Visible=   0   'False
         Columns(13).Caption=   "AlwaysTransfer"
         Columns(13).Name=   "AlwaysTransfer"
         Columns(13).DataField=   "Column 13"
         Columns(13).DataType=   17
         Columns(13).FieldLen=   256
         Columns(14).Width=   3200
         Columns(14).Visible=   0   'False
         Columns(14).Caption=   "ConvertData"
         Columns(14).Name=   "ConvertData"
         Columns(14).DataField=   "Column 14"
         Columns(14).DataType=   17
         Columns(14).FieldLen=   256
         Columns(15).Width=   3200
         Columns(15).Visible=   0   'False
         Columns(15).Caption=   "IsEmployeeName"
         Columns(15).Name=   "IsEmployeeName"
         Columns(15).DataField=   "Column 15"
         Columns(15).DataType=   8
         Columns(15).FieldLen=   256
         Columns(16).Width=   3200
         Columns(16).Visible=   0   'False
         Columns(16).Caption=   "IsDepartmentCode"
         Columns(16).Name=   "IsDepartmentCode"
         Columns(16).DataField=   "Column 16"
         Columns(16).DataType=   8
         Columns(16).FieldLen=   256
         Columns(17).Width=   3200
         Columns(17).Visible=   0   'False
         Columns(17).Caption=   "IsDepartmentName"
         Columns(17).Name=   "IsDepartmentName"
         Columns(17).DataField=   "Column 17"
         Columns(17).DataType=   8
         Columns(17).FieldLen=   256
         Columns(18).Width=   3200
         Columns(18).Visible=   0   'False
         Columns(18).Caption=   "IsPayrollCode"
         Columns(18).Name=   "IsPayrollCode"
         Columns(18).DataField=   "Column 18"
         Columns(18).DataType=   8
         Columns(18).FieldLen=   256
         Columns(19).Width=   3200
         Columns(19).Visible=   0   'False
         Columns(19).Caption=   "Group"
         Columns(19).Name=   "Group"
         Columns(19).DataField=   "Column 19"
         Columns(19).DataType=   8
         Columns(19).FieldLen=   256
         Columns(20).Width=   3200
         Columns(20).Visible=   0   'False
         Columns(20).Caption=   "PreventModify"
         Columns(20).Name=   "PreventModify"
         Columns(20).DataField=   "Column 20"
         Columns(20).DataType=   8
         Columns(20).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   11483
         _ExtentY        =   5741
         _StockProps     =   79
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
      Begin VB.Label lblColumnName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         Height          =   195
         Left            =   225
         TabIndex        =   21
         Top             =   1005
         Width           =   645
      End
      Begin VB.Label lblTransferTable 
         Caption         =   "Table : "
         Height          =   285
         Left            =   225
         TabIndex        =   19
         Top             =   585
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   5700
      TabIndex        =   1
      Top             =   5850
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   405
      Left            =   4410
      TabIndex        =   0
      Top             =   5850
      Width           =   1200
   End
End
Attribute VB_Name = "frmModuleDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnReadOnly As Boolean
Private mbChanged As Boolean

Private Sub cboCategoryTable_Click()

  Dim objctl As Control

  ' Clear the current contents of the combos.
  For Each objctl In Me
  
    If (TypeOf objctl Is ComboBox) And _
      (objctl.Tag = "cboCategoryTable") Then
      
        PopulateComboWithColumns objctl, GetComboItem(cboCategoryTable), True
      
    End If
  Next objctl

  mbChanged = True

  RefreshButtons

End Sub

Private Sub cboTypeTable_Click()

  Dim objctl As Control

  ' Clear the current contents of the combos.
  For Each objctl In Me
  
    If (TypeOf objctl Is ComboBox) And _
      (objctl.Tag = "cboTypeTable") Then
      
        PopulateComboWithColumns objctl, GetComboItem(cboTypeTable), True
      
    End If
  Next objctl

  mbChanged = True

  RefreshButtons

End Sub

Private Sub cboTypeColumn_Click()
  mbChanged = True
  RefreshButtons
End Sub

Private Sub cboTypeCategoryColumn_Click()
  mbChanged = True
  RefreshButtons
End Sub

Private Sub cmdCancel_Click()
  UnLoad Me
End Sub

Private Sub cmdOK_Click()

  If SaveChanges Then
    mbChanged = False
    UnLoad Me
  End If

End Sub


Private Function SaveChanges() As Boolean

  ' Category info
  SaveModuleSetting gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_CATEGORYTABLE, gsPARAMETERTYPE_TABLEID, GetComboItem(cboCategoryTable)
  SaveModuleSetting gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_CATEGORYCOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboCategoryColumn)

  ' Types info
  SaveModuleSetting gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_TYPETABLE, gsPARAMETERTYPE_TABLEID, GetComboItem(cboTypeTable)
  SaveModuleSetting gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_TYPECOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboTypeColumn)
  SaveModuleSetting gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_TYPECATEGORYCOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboTypeCategoryColumn)

  SaveChanges = True

End Function


Private Sub PopulateComboWithColumns(ByRef cboTemp As ComboBox, ByVal plngTableID As Long, ByVal AllowNone As Boolean)

  If AllowNone Then
    With cboTemp
      .Clear
      .AddItem "<None>"
      .ItemData(.NewIndex) = 0
    End With
  End If


  With recColEdit
    .Index = "idxName"
    .Seek ">=", plngTableID

    If Not .NoMatch Then
      Do While Not .EOF
        If !TableID <> plngTableID Then
          Exit Do
        End If

        If (Not !Deleted) And (!DataType = dtVARCHAR) Then

          cboTemp.AddItem (!ColumnName)
          cboTemp.ItemData(cboTemp.NewIndex) = !ColumnID

        End If

        .MoveNext
      Loop
    End If
  End With

End Sub

' Initialise the Base Table combo(s)
Private Sub InitialiseCombos()
   
  With cboCategoryTable
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
   
   
  With cboTypeTable
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
    
    
  ' Add items to the combo for each table that has not been deleted,
  ' and is a Lookup table.
  With recTabEdit
    .Index = "idxName"
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      If Not !Deleted And (!TableType = giTABLELOOKUP) Then
                
        cboTypeTable.AddItem !TableName
        cboTypeTable.ItemData(cboTypeTable.NewIndex) = !TableID
                
        cboCategoryTable.AddItem !TableName
        cboCategoryTable.ItemData(cboTypeTable.NewIndex) = !TableID
                
      End If
      .MoveNext
    Loop
  End With
    

End Sub

Private Sub RefreshButtons()
  cmdOk.Enabled = mbChanged
End Sub

Private Sub RetrieveDefinition()

  SetComboItem cboCategoryTable, GetModuleSetting(gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_CATEGORYTABLE, 0)
  SetComboItem cboCategoryColumn, GetModuleSetting(gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_CATEGORYCOLUMN, 0)

  SetComboItem cboTypeTable, GetModuleSetting(gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_TYPETABLE, 0)
  SetComboItem cboTypeCategoryColumn, GetModuleSetting(gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_TYPECATEGORYCOLUMN, 0)
  SetComboItem cboTypeColumn, GetModuleSetting(gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_TYPECOLUMN, 0)

End Sub

Private Sub Form_Load()

  mblnReadOnly = (Application.AccessMode <> accFull And _
                 Application.AccessMode <> accSupportMode)

  If mblnReadOnly Then
    ControlsDisableAll Me
  End If
  
  InitialiseCombos
  PopulateParentsCombo cboTransferTables


  RetrieveDefinition

  mbChanged = False

  ' Get rid of that pesky icon
  RemoveIcon Me

End Sub


Private Sub PopulateParentsCombo(ByRef objCombo As ComboBox)
  ' Clear the contents of the combo.
  objCombo.Clear
  
  With recTabEdit
    .Index = "idxName"
    
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      If !TableType = iTabParent And Not !Deleted Then
        objCombo.AddItem !TableName
        objCombo.ItemData(objCombo.NewIndex) = !TableID
      End If
      
      .MoveNext
    Loop
  End With
  
End Sub

