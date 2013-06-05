VERSION 5.00
Begin VB.Form frmCategorySetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Categories"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   375
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
   Icon            =   "frmCategorySetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraComponent 
      Caption         =   "Categories :"
      Height          =   1440
      Index           =   1
      Left            =   90
      TabIndex        =   2
      Tag             =   "6"
      Top             =   45
      Width           =   5865
      Begin VB.ComboBox cboCategoryTable 
         Height          =   315
         Left            =   2430
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   3255
      End
      Begin VB.ComboBox cboCategoryColumn 
         Height          =   315
         Left            =   2430
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "cboCategoryTable"
         Top             =   795
         Width           =   3255
      End
      Begin VB.Label lblCategoryTable 
         Caption         =   "Category Table : "
         Height          =   285
         Left            =   225
         TabIndex        =   6
         Top             =   405
         Width           =   1680
      End
      Begin VB.Label lblCatgeoryColumn 
         Caption         =   "Name Column : "
         Height          =   330
         Left            =   225
         TabIndex        =   5
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4740
      TabIndex        =   1
      Top             =   1620
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   3420
      TabIndex        =   0
      Top             =   1620
      Width           =   1200
   End
End
Attribute VB_Name = "frmCategorySetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnReadOnly As Boolean

Private mbLoading As Boolean

Private mlngTableID As Long
Private mlngColumnID As Long
Private mfChanged As Boolean

Public Property Get Changed() As Boolean
  Changed = mfChanged
End Property
Public Property Let Changed(ByVal pblnChanged As Boolean)
  mfChanged = pblnChanged
  If Not mbLoading Then cmdOk.Enabled = True
End Property

Private Sub cboCategoryTable_Click()

  mlngTableID = cboCategoryTable.ItemData(cboCategoryTable.ListIndex)

  RefreshControls
  
  Changed = True

End Sub

Private Sub cmdCancel_Click()
  UnLoad Me
End Sub

Private Sub cmdOK_Click()

  If SaveChanges Then
    Changed = False
    UnLoad Me
  End If

End Sub

Private Function SaveChanges() As Boolean
  'AE20071119 Fault #12607
  SaveChanges = False

'  If Not ValidateSetup Then
'    Exit Function
'  End If

  Screen.MousePointer = vbHourglass
  ' Write the parameter values to the local database.


  With recModuleSetup
    .Index = "idxModuleParameter"

    ' Save the Conversion Table ID.
    .Seek "=", gsMODULEKEY_CATEGORY, gsPARAMETERKEY_CATEGORYTABLE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_CATEGORY
      !parameterkey = gsPARAMETERKEY_CATEGORYTABLE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_TABLEID
    !parametervalue = mlngTableID
    .Update

    ' Save the Currency Name column ID.
    .Seek "=", gsMODULEKEY_CATEGORY, gsPARAMETERKEY_CATEGORYNAMECOLUMN
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_CATEGORY
      !parameterkey = gsPARAMETERKEY_CATEGORYNAMECOLUMN
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mlngColumnID

    .Update

  End With

  'AE20071119 Fault #12607
  SaveChanges = True
  Application.Changed = True
  
  Screen.MousePointer = vbDefault
End Function

Private Sub InitialiseBaseTableCombos()
  
  ' Initialise the Base Table combo(s)
  Dim iTableListIndex As Integer
    
  iTableListIndex = 0
  
' Clear the combos, and add '<None>' items.
  cboCategoryTable.Clear
  cboCategoryTable.AddItem "<None>"
  cboCategoryTable.ItemData(cboCategoryTable.NewIndex) = 0
    
  ' Add items to the combo for each table that has not been deleted,
  ' and is a Lookup table.
  With recTabEdit
    .Index = "idxName"
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      If Not !Deleted And (!TableType = giTABLELOOKUP) Then
        
        cboCategoryTable.AddItem !TableName
        cboCategoryTable.ItemData(cboCategoryTable.NewIndex) = !TableID
        
        If !TableID = mlngTableID Then
          iTableListIndex = cboCategoryTable.NewIndex
        End If
      End If
      .MoveNext
    Loop
  End With
  
  cboCategoryTable.Enabled = Not mblnReadOnly
  cboCategoryTable.ListIndex = iTableListIndex

End Sub

Private Sub ReadParameters()
  
  ' Read the parameter values from the database into local variables.
  With recModuleSetup
    .Index = "idxModuleParameter"

    ' Get the Currency conversion table ID.
    .Seek "=", gsMODULEKEY_CATEGORY, gsPARAMETERKEY_CATEGORYTABLE
    If .NoMatch Then
      mlngTableID = 0
    Else
      mlngTableID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the Currency name column ID.
    .Seek "=", gsMODULEKEY_CATEGORY, gsPARAMETERKEY_CATEGORYNAMECOLUMN
    If .NoMatch Then
      mlngColumnID = 0
    Else
      mlngColumnID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

  End With

End Sub

Private Sub Form_Load()

  mbLoading = True
  Screen.MousePointer = vbHourglass

  mblnReadOnly = (Application.AccessMode <> accFull And _
                  Application.AccessMode <> accSupportMode)

  If mblnReadOnly Then
    ControlsDisableAll Me
  End If

  ' Read the current settings from the database.
  ReadParameters

  ' Initialise all controls with the current settings, or defaults.
  InitialiseBaseTableCombos

  cmdOk.Enabled = False
  Changed = False
  mbLoading = False
  Screen.MousePointer = vbDefault
 
End Sub

Private Sub RefreshControls()
  
  Dim iNameIndex As Integer
  Dim iValueIndex As Integer
  Dim iDecimalIndex As Integer
  
  cboCategoryColumn.Enabled = Not (cboCategoryTable.ListIndex = 0)
  cboCategoryColumn.Clear
  
  With recColEdit
    .Index = "idxName"
    .Seek ">=", mlngTableID

    If Not .NoMatch Then
      ' Add items to the combos for each column that has not been deleted,
      ' or is a system or link column.
      Do While Not .EOF
        If !TableID <> mlngTableID Then
          Exit Do
        End If

        If (Not !Deleted) And _
          (!columntype <> giCOLUMNTYPE_LINK) And _
          (!DataType = dtVARCHAR) And _
          (!columntype <> giCOLUMNTYPE_SYSTEM) Then

            cboCategoryColumn.AddItem !ColumnName
            cboCategoryColumn.ItemData(cboCategoryColumn.NewIndex) = !ColumnID

        End If

        .MoveNext
      Loop
    End If
  End With

  If cboCategoryColumn.ListCount > 0 Then
    cboCategoryColumn.ListIndex = 0
    mlngColumnID = cboCategoryColumn.ItemData(cboCategoryColumn.ListIndex)
  Else
    cboCategoryColumn.Enabled = False
    mlngColumnID = 0
  End If

  cboCategoryColumn.BackColor = IIf(cboCategoryColumn.Enabled, vbWindowBackground, vbButtonFace)

End Sub
