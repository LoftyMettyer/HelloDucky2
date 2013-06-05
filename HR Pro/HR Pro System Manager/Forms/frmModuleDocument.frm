VERSION 5.00
Begin VB.Form frmModuleDocument 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Document Management"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6060
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraComponent 
      Caption         =   "Categories :"
      Height          =   1305
      Index           =   5
      Left            =   90
      TabIndex        =   9
      Tag             =   "6"
      Top             =   45
      Width           =   5865
      Begin VB.ComboBox cboCategoryColumn 
         Height          =   315
         Left            =   2430
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Tag             =   "cboCategoryTable"
         Top             =   750
         Width           =   3255
      End
      Begin VB.ComboBox cboCategoryTable 
         Height          =   315
         Left            =   2430
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label lblCatgeoryColumn 
         Caption         =   "Category Column : "
         Height          =   330
         Left            =   225
         TabIndex        =   13
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label lblCategoryTable 
         Caption         =   "Category Table : "
         Height          =   285
         Left            =   225
         TabIndex        =   12
         Top             =   405
         Width           =   1680
      End
   End
   Begin VB.Frame fraTypes 
      Caption         =   "Types : "
      Height          =   1755
      Left            =   90
      TabIndex        =   2
      Top             =   1530
      Width           =   5865
      Begin VB.ComboBox cboTypeCategoryColumn 
         Height          =   315
         Left            =   2430
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "cboTypeTable"
         Top             =   765
         Width           =   3255
      End
      Begin VB.ComboBox cboTypeColumn 
         Height          =   315
         Left            =   2430
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "cboTypeTable"
         Top             =   1215
         Width           =   3255
      End
      Begin VB.ComboBox cboTypeTable 
         Height          =   315
         Left            =   2430
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   315
         Width           =   3255
      End
      Begin VB.Label lblTypeCategoryColumn 
         Caption         =   "Type Category Column : "
         Height          =   285
         Left            =   195
         TabIndex        =   8
         Top             =   810
         Width           =   2130
      End
      Begin VB.Label lblTypeColumn 
         Caption         =   "Type Column : "
         Height          =   285
         Left            =   195
         TabIndex        =   7
         Top             =   1260
         Width           =   1410
      End
      Begin VB.Label lblTypeTable 
         Caption         =   "Type Table : "
         Height          =   285
         Left            =   195
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4755
      TabIndex        =   1
      Top             =   3420
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   3465
      TabIndex        =   0
      Top             =   3420
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
  RetrieveDefinition

  mbChanged = False

  ' Get rid of that pesky icon
  RemoveIcon Me

End Sub
