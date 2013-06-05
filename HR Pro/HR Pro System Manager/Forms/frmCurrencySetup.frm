VERSION 5.00
Object = "{1EE59219-BC23-4BDF-BB08-D545C8A38D6D}#1.0#0"; "COA_Line.ocx"
Begin VB.Form frmCurrencySetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Currency Setup"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1042
   Icon            =   "frmCurrencySetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   3075
      TabIndex        =   2
      Top             =   2820
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4395
      TabIndex        =   3
      Top             =   2820
      Width           =   1200
   End
   Begin VB.Frame fraCConv 
      Caption         =   "Currency Conversion Table Information :"
      Height          =   2565
      Left            =   120
      TabIndex        =   4
      Top             =   90
      Width           =   5475
      Begin VB.ComboBox cboCConvDecimal 
         Height          =   315
         Left            =   2865
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2025
         Width           =   2500
      End
      Begin VB.ComboBox cboCConvValue 
         Height          =   315
         Left            =   2865
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1485
         Width           =   2500
      End
      Begin VB.ComboBox cboCConvTable 
         Height          =   315
         Left            =   2865
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   2500
      End
      Begin VB.ComboBox cboCConvName 
         Height          =   315
         Left            =   2865
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   945
         Width           =   2500
      End
      Begin COALine.COA_Line ASRDummyLine3 
         Height          =   30
         Left            =   150
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   765
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   53
      End
      Begin VB.Label lblCConvDecimal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Decimal Column :"
         Height          =   195
         Left            =   195
         TabIndex        =   11
         Top             =   2085
         Width           =   1755
      End
      Begin VB.Label lblCConvValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conversion Value Column :"
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   1545
         Width           =   2460
      End
      Begin VB.Label lblCConvName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Currency Name Column :"
         Height          =   195
         Left            =   195
         TabIndex        =   7
         Top             =   1005
         Width           =   2325
      End
      Begin VB.Label lblCConvTable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Currency Conversion Table :"
         Height          =   195
         Left            =   195
         TabIndex        =   6
         Top             =   360
         Width           =   2595
      End
   End
End
Attribute VB_Name = "frmCurrencySetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvar_lngCConvTableID As Long
Private mvar_lngCConvNameColumnID As Long
Private mvar_lngCConvValueColumnID As Long
Private mvar_lngCConvDecimalColumnID As Long

Private mblnReadOnly As Boolean

Private mbLoading As Boolean
Private mfChanged As Boolean

Public Property Get Changed() As Boolean
  Changed = mfChanged
End Property
Public Property Let Changed(ByVal pblnChanged As Boolean)
  mfChanged = pblnChanged
  If Not mbLoading Then cmdOK.Enabled = True
End Property

Private Sub cboCConvDecimal_Click()

  With Me.cboCConvDecimal
    mvar_lngCConvDecimalColumnID = .ItemData(.ListIndex)
  End With
  
  Changed = True
  
End Sub

Private Sub cboCConvName_Change()
  Changed = True
End Sub

Private Sub cboCConvName_Click()

  With Me.cboCConvName
    mvar_lngCConvNameColumnID = .ItemData(.ListIndex)
  End With
  
  Changed = True
End Sub

Private Sub cboCConvTable_Change()
  Changed = True
End Sub

Private Sub cboCConvValue_Change()
  Changed = True
End Sub

Private Sub cboCConvValue_Click()
  With Me.cboCConvValue
    mvar_lngCConvValueColumnID = .ItemData(.ListIndex)
  End With
  
  Changed = True
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
  
  cmdOK.Enabled = False
  Changed = False
  mbLoading = False
  Screen.MousePointer = vbDefault
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  ' If the user cancels or tries to close the form
  'AE20071119 Fault #12607
  'If UnloadMode <> vbFormCode And cmdOK.Enabled Then
  If Changed Then
    Select Case MsgBox("Apply module changes ?", vbYesNoCancel + vbQuestion, Me.Caption)
      Case vbCancel
        Cancel = True
      Case vbYes
        'AE20071119 Fault #12607
        'SaveChanges
        Cancel = (Not SaveChanges)
    End Select
  End If
  
End Sub

Private Sub InitialiseBaseTableCombos()
  
  ' Initialise the Base Table combo(s)
  Dim iConversionTableListIndex As Integer
    
  iConversionTableListIndex = 0
  
' Clear the combos, and add '<None>' items.
  Me.cboCConvTable.Clear
  Me.cboCConvTable.AddItem "<None>"
  Me.cboCConvTable.ItemData(cboCConvTable.NewIndex) = 0
    
  ' Add items to the combo for each table that has not been deleted,
  ' and is a Lookup table.
  With recTabEdit
    .Index = "idxName"
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      If Not !Deleted And (!TableType = giTABLELOOKUP) Then
        
        cboCConvTable.AddItem !TableName
        cboCConvTable.ItemData(cboCConvTable.NewIndex) = !TableID
        
        If !TableID = mvar_lngCConvTableID Then
          iConversionTableListIndex = cboCConvTable.NewIndex
        End If
      End If
      .MoveNext
    Loop
  End With
  
  cboCConvTable.Enabled = Not mblnReadOnly
  cboCConvTable.ListIndex = iConversionTableListIndex

End Sub

Private Sub cboCConvTable_Click()

  With Me.cboCConvTable
    mvar_lngCConvTableID = .ItemData(.ListIndex)
  End With

  RefreshCConvControls
  
  Changed = True

End Sub

Private Sub RefreshCConvControls()
  
  Dim iNameIndex As Integer
  Dim iValueIndex As Integer
  Dim iDecimalIndex As Integer
  
  If cboCConvTable.ListIndex = 0 Then
    cboCConvName.Enabled = False
    cboCConvValue.Enabled = False
    cboCConvDecimal.Enabled = False
  End If
  
  cboCConvName.Clear
  cboCConvValue.Clear
  cboCConvDecimal.Clear
  
  With recColEdit
    .Index = "idxName"
    .Seek ">=", mvar_lngCConvTableID

    If Not .NoMatch Then
      ' Add items to the combos for each column that has not been deleted,
      ' or is a system or link column.
      Do While Not .EOF
        If !TableID <> mvar_lngCConvTableID Then
          Exit Do
        End If

        If (Not !Deleted) And _
          (!ColumnType <> giCOLUMNTYPE_LINK) And _
          (!ColumnType <> giCOLUMNTYPE_SYSTEM) Then

          ' Load varchar fields into the name combo.
          If !DataType = dtVARCHAR Then
            cboCConvName.AddItem !ColumnName
            cboCConvName.ItemData(cboCConvName.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngCConvNameColumnID Then
              iNameIndex = cboCConvName.NewIndex
            End If
          End If
          
          ' Load numeric fields into the conversion combo.
          If !DataType = adDecimal Or _
             !DataType = dtNUMERIC Or _
             !DataType = adDouble Then
            cboCConvValue.AddItem !ColumnName
            cboCConvValue.ItemData(cboCConvValue.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngCConvValueColumnID Then
              iValueIndex = cboCConvValue.NewIndex
            End If
          End If
          
          ' Load varchar fields into the name combo.
          If !DataType = dtINTEGER Then
            cboCConvDecimal.AddItem !ColumnName
            cboCConvDecimal.ItemData(cboCConvDecimal.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngCConvDecimalColumnID Then
              iDecimalIndex = cboCConvDecimal.NewIndex
            End If
          End If
          
          Changed = True

        End If

        .MoveNext
      Loop
    End If
  End With

  If cboCConvName.ListCount > 0 And cboCConvValue.ListCount > 0 _
      And cboCConvDecimal.ListCount > 0 Then
    cboCConvName.ListIndex = 0
    cboCConvValue.ListIndex = 0
    cboCConvDecimal.ListIndex = 0

    cboCConvName.Enabled = (cboCConvName.ListCount > 0)
    cboCConvValue.Enabled = (cboCConvValue.ListCount > 0)
    cboCConvDecimal.Enabled = (cboCConvDecimal.ListCount > 0)
  Else
    cboCConvName.Enabled = False
    cboCConvValue.Enabled = False
    cboCConvDecimal.Enabled = False
  End If

  cboCConvName.BackColor = IIf(cboCConvName.Enabled, vbWindowBackground, vbButtonFace)
  cboCConvValue.BackColor = IIf(cboCConvValue.Enabled, vbWindowBackground, vbButtonFace)
  cboCConvDecimal.BackColor = IIf(cboCConvDecimal.Enabled, vbWindowBackground, vbButtonFace)

End Sub

Private Sub cmdOK_Click()

  'AE20071119 Fault #12607
  'If ValidateSetup Then
    'SaveChanges
  If SaveChanges Then
    Changed = False
    UnLoad Me
  End If
  
End Sub

Private Sub cmdCancel_Click()
  'AE20071119 Fault #12607
'  Dim pintAnswer As Integer
'    If Changed = True And cmdOK.Enabled Then
'      pintAnswer = MsgBox("You have made changes...do you wish to save these changes ?", vbQuestion + vbYesNoCancel, App.Title)
'      If pintAnswer = vbYes Then
'        'AE20071108 Fault #12551
'        'Using Me.MousePointer = vbNormal forces the form to be reloaded
'        'after its been unloaded in cmdOK_Click, changed to Screen.MousePointer
'        'Me.MousePointer = vbHourglass
'        Screen.MousePointer = vbHourglass
'        cmdOK_Click 'This is just like saving
'        'Me.MousePointer = vbNormal
'        Screen.MousePointer = vbNormal
'        Exit Sub
'      ElseIf pintAnswer = vbCancel Then
'        Exit Sub
'      End If
'    End If
'TidyUpAndExit:
  UnLoad Me
End Sub

Private Function ValidateSetup() As Boolean
  ValidateSetup = True
End Function

' Save the CMG Export Details
Private Function SaveChanges() As Boolean
  'AE20071119 Fault #12607
  SaveChanges = False
  
  If Not ValidateSetup Then
    Exit Function
  End If
  
  Screen.MousePointer = vbHourglass
  ' Write the parameter values to the local database.

  With recModuleSetup
    .Index = "idxModuleParameter"

    ' Save the Conversion Table ID.
    .Seek "=", gsMODULEKEY_CURRENCY, gsPARAMETERKEY_CONVERSIONTABLE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_CURRENCY
      !parameterkey = gsPARAMETERKEY_CONVERSIONTABLE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_TABLEID
    !parametervalue = mvar_lngCConvTableID
    .Update

    ' Save the Currency Name column ID.
    .Seek "=", gsMODULEKEY_CURRENCY, gsPARAMETERKEY_CURRENCYNAMECOLUMN
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_CURRENCY
      !parameterkey = gsPARAMETERKEY_CURRENCYNAMECOLUMN
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngCConvNameColumnID
    .Update

    ' Save the Conversion Value column ID.
    .Seek "=", gsMODULEKEY_CURRENCY, gsPARAMETERKEY_CONVERSIONVALUECOLUMN
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_CURRENCY
      !parameterkey = gsPARAMETERKEY_CONVERSIONVALUECOLUMN
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngCConvValueColumnID
    .Update

    ' Save the Decimal column ID.
    .Seek "=", gsMODULEKEY_CURRENCY, gsPARAMETERKEY_DECIMALCOLUMN
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_CURRENCY
      !parameterkey = gsPARAMETERKEY_DECIMALCOLUMN
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngCConvDecimalColumnID
    .Update
  
  End With

  'AE20071119 Fault #12607
  SaveChanges = True
  Application.Changed = True
  
  Screen.MousePointer = vbNormal
End Function

Private Sub ReadParameters()
  
  ' Read the parameter values from the database into local variables.
  
  With recModuleSetup
    .Index = "idxModuleParameter"

    ' Get the Currency conversion table ID.
    .Seek "=", gsMODULEKEY_CURRENCY, gsPARAMETERKEY_CONVERSIONTABLE
    If .NoMatch Then
      mvar_lngCConvTableID = 0
    Else
      mvar_lngCConvTableID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the Currency name column ID.
    .Seek "=", gsMODULEKEY_CURRENCY, gsPARAMETERKEY_CURRENCYNAMECOLUMN
    If .NoMatch Then
      mvar_lngCConvNameColumnID = 0
    Else
      mvar_lngCConvNameColumnID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the Conversion Value column ID.
    .Seek "=", gsMODULEKEY_CURRENCY, gsPARAMETERKEY_CONVERSIONVALUECOLUMN
    If .NoMatch Then
      mvar_lngCConvValueColumnID = 0
    Else
      mvar_lngCConvValueColumnID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the Decimal column ID.
    .Seek "=", gsMODULEKEY_CURRENCY, gsPARAMETERKEY_DECIMALCOLUMN
    If .NoMatch Then
      mvar_lngCConvDecimalColumnID = 0
    Else
      mvar_lngCConvDecimalColumnID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

  End With

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


