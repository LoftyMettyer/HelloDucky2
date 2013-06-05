VERSION 5.00
Object = "{1EE59219-BC23-4BDF-BB08-D545C8A38D6D}#1.1#0"; "COA_Line.ocx"
Begin VB.Form frmBankHolidaySetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bank Holidays Setup"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5005
   Icon            =   "frmBankHolidaySetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2730
      TabIndex        =   5
      Top             =   3615
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4080
      TabIndex        =   6
      Top             =   3615
      Width           =   1200
   End
   Begin VB.Frame fraBankHolidays 
      Caption         =   "Bank Holidays Table :"
      Height          =   1860
      Left            =   120
      TabIndex        =   11
      Top             =   1620
      Width           =   5160
      Begin VB.ComboBox cboBHolTable 
         Height          =   315
         Left            =   2500
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   2500
      End
      Begin VB.ComboBox cboBHolDate 
         Height          =   315
         Left            =   2500
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   945
         Width           =   2500
      End
      Begin VB.ComboBox cboBHolDescription 
         Height          =   315
         Left            =   2500
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1350
         Width           =   2500
      End
      Begin COALine.COA_Line ASRDummyLine4 
         Height          =   30
         Left            =   150
         Top             =   765
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   53
      End
      Begin VB.Label lblBHolTable 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Holiday Table :"
         Height          =   195
         Left            =   195
         TabIndex        =   8
         Top             =   360
         Width           =   1905
      End
      Begin VB.Label lblBHolDate 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Column :"
         Height          =   195
         Left            =   195
         TabIndex        =   12
         Top             =   1005
         Width           =   1425
      End
      Begin VB.Label lblBHolDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Description Column :"
         Height          =   195
         Left            =   195
         TabIndex        =   13
         Top             =   1410
         Width           =   1875
      End
   End
   Begin VB.Frame fraBHolRegion 
      Caption         =   "Bank Holiday Region Table :"
      Height          =   1440
      Left            =   120
      TabIndex        =   7
      Top             =   90
      Width           =   5160
      Begin VB.ComboBox cboBHolRegionTable 
         Height          =   315
         Left            =   2500
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   2500
      End
      Begin VB.ComboBox cboBHolRegion 
         Height          =   315
         Left            =   2500
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   945
         Width           =   2500
      End
      Begin COALine.COA_Line ASRDummyLine3 
         Height          =   30
         Left            =   150
         Top             =   765
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   53
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Region Column :"
         Height          =   195
         Left            =   195
         TabIndex        =   10
         Top             =   1005
         Width           =   1170
      End
      Begin VB.Label lblBHolRegionTable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Region Table :"
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   360
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmBankHolidaySetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mvar_lngBHolRegionTableID As Long
Private mvar_lngBHolRegionID As Long
Private mvar_lngBHolTableID As Long
Private mvar_lngBHolDateID As Long
Private mvar_lngBHolDescriptionID As Long
Private mblnReadOnly As Boolean
Private mfChanged As Boolean

Public Property Get Changed() As Boolean
  Changed = mfChanged
End Property
Public Property Let Changed(ByVal pblnChanged As Boolean)
  mfChanged = pblnChanged
  cmdOk.Enabled = True
End Property

Private Sub cboBHolDate_Change()
  Changed = True
End Sub

Private Sub cboBHolDescription_Change()
  Changed = True
End Sub

Private Sub cboBHolRegion_Change()
  Changed = True
End Sub

Private Sub cboBHolRegionTable_Change()
  Changed = True
End Sub

Private Sub cboBHolRegionTable_Click()

  With cboBHolRegionTable
    mvar_lngBHolRegionTableID = .ItemData(.ListIndex)
  End With

  RefreshBHolRegionControls

End Sub

Private Sub cboBHolTable_Change()
  Changed = True
End Sub

Private Sub Form_Load()
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
  
  'AE20080204 Fault #12829
  mfChanged = False
  cmdOk.Enabled = False
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
  Dim iBHolRegionTableListIndex As Integer
    
  iBHolRegionTableListIndex = 0
  
  ' Clear the combos, and add '<None>' items.
    cboBHolRegionTable.Clear
    cboBHolRegionTable.AddItem "<None>"
    cboBHolRegionTable.ItemData(cboBHolRegionTable.NewIndex) = 0
    
  ' Add items to the combo for each table that has not been deleted,
  ' and is a child of the defined Personnel table.
  With recTabEdit
    .Index = "idxName"
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      If Not !Deleted Then
        
        cboBHolRegionTable.AddItem !TableName
        cboBHolRegionTable.ItemData(cboBHolRegionTable.NewIndex) = !TableID
        
        If !TableID = mvar_lngBHolRegionTableID Then
          iBHolRegionTableListIndex = cboBHolRegionTable.NewIndex
        End If
      End If
      .MoveNext
    Loop
  End With
  
  ' Select the appropriate combo items.
    'cboBHolTable.Enabled = True
    cboBHolTable.Enabled = Not mblnReadOnly
    cboBHolRegionTable.ListIndex = iBHolRegionTableListIndex

End Sub

Private Sub cboBHolTable_Click()
  With cboBHolTable
    mvar_lngBHolTableID = .ItemData(.ListIndex)
  End With

  RefreshBHolControls
  Changed = True
End Sub

Private Sub cboBHolDate_Click()
  With cboBHolDate
    mvar_lngBHolDateID = .ItemData(.ListIndex)
  End With
  Changed = True
End Sub

Private Sub cboBHolDescription_Click()
  With cboBHolDescription
    mvar_lngBHolDescriptionID = .ItemData(.ListIndex)
  End With
Changed = True
End Sub

Private Sub cboBHolRegion_Click()
  With cboBHolRegion
    mvar_lngBHolRegionID = .ItemData(.ListIndex)
  End With
Changed = True
End Sub

Private Sub RefreshBHolRegionControls()

  ' Refresh the BHol Region controls
  Dim iBHolRegionListIndex As Integer
  Dim iBHolTableListIndex As Integer
  
  iBHolRegionListIndex = 0
  iBHolTableListIndex = 0

  ' Clear the current contents of the BHolRegion field combo
  cboBHolRegion.Clear
  cboBHolRegion.AddItem "<None>"
  cboBHolRegion.ItemData(cboBHolRegion.NewIndex) = 0

  With recColEdit
    .Index = "idxName"
    .Seek ">=", mvar_lngBHolRegionTableID

    If Not .NoMatch Then
      ' Add non system/link cols to the combo that have not been deleted
      Do While Not .EOF
        If !TableID <> mvar_lngBHolRegionTableID Then
          Exit Do
        End If
        If (Not !Deleted) And _
          (!columnType <> giCOLUMNTYPE_LINK) And _
          (!columnType <> giCOLUMNTYPE_SYSTEM) Then

          ' Load varchar fields
          If !DataType = dtVARCHAR Then
            cboBHolRegion.AddItem !ColumnName
            cboBHolRegion.ItemData(cboBHolRegion.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngBHolRegionID Then
              iBHolRegionListIndex = cboBHolRegion.NewIndex
            End If
          End If
        End If
        .MoveNext
      Loop
    End If
  End With

  ' Select the appropriate combo items.
  cboBHolRegion.ListIndex = iBHolRegionListIndex

  ' Now populate the BankHolidayTable combo with children of the table selected
  ' in the BHolRegionTable combo
  
  ' Clear the current contents of the combo
  cboBHolTable.Clear
  cboBHolTable.AddItem "<None>"
  cboBHolTable.ItemData(cboBHolTable.NewIndex) = 0
  
  ' Add the tables that are children of the table selected in the BHolRegionTable
  ' combo
  
  With recTabEdit
    .Index = "idxName"
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      If Not !Deleted Then
        recRelEdit.Index = "idxParentID"
        recRelEdit.Seek "=", mvar_lngBHolRegionTableID, !TableID
        
        If Not recRelEdit.NoMatch Then
          cboBHolTable.AddItem !TableName
          cboBHolTable.ItemData(cboBHolTable.NewIndex) = !TableID
        
          If !TableID = mvar_lngBHolTableID Then
            iBHolTableListIndex = cboBHolTable.NewIndex
          End If
        End If
        
      End If
      .MoveNext
    Loop
  End With
  
  ' Select the appropriate combo items.
  cboBHolTable.ListIndex = iBHolTableListIndex

End Sub

Private Sub RefreshBHolControls()
  ' Refresh the BHol controls.
  Dim iBHolDateListIndex As Integer
  Dim iBHolDescriptionListIndex As Integer
  Dim objCtl As Control

  iBHolDateListIndex = 0
  iBHolDescriptionListIndex = 0

  ' Clear the current contents of the combos.
  For Each objCtl In Me
    If TypeOf objCtl Is ComboBox And _
      (objCtl.Name = "cboBHolDate" Or _
      objCtl.Name = "cboBHolDescription") Then
        
      With objCtl
        .Clear
        .AddItem "<None>"
        .ItemData(.NewIndex) = 0
      End With
    End If
  Next objCtl

  With recColEdit
    .Index = "idxName"
    .Seek ">=", mvar_lngBHolTableID

    If Not .NoMatch Then
      ' Add items to the combos for each column that has not been deleted,
      ' or is a system or link column.
      Do While Not .EOF
        If !TableID <> mvar_lngBHolTableID Then
          Exit Do
        End If

        If (Not !Deleted) And _
          (!columnType <> giCOLUMNTYPE_LINK) And _
          (!columnType <> giCOLUMNTYPE_SYSTEM) Then

          ' Load date fields
          If !DataType = dtTIMESTAMP Then
            cboBHolDate.AddItem !ColumnName
            cboBHolDate.ItemData(cboBHolDate.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngBHolDateID Then
              iBHolDateListIndex = cboBHolDate.NewIndex
            End If
          End If
          
          ' Load varchar fields
          If !DataType = dtVARCHAR Then
            cboBHolDescription.AddItem !ColumnName
            cboBHolDescription.ItemData(cboBHolDescription.NewIndex) = !ColumnID
            If !ColumnID = mvar_lngBHolDescriptionID Then
              iBHolDescriptionListIndex = cboBHolDescription.NewIndex
            End If

          End If
        End If

        .MoveNext
      Loop
    End If
  End With

  ' Select the appropriate combo items.
  cboBHolDate.ListIndex = iBHolDateListIndex
  cboBHolDescription.ListIndex = iBHolDescriptionListIndex

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

    ' Save the BHol Region Table ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLREGIONTABLE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_BHOLREGIONTABLE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_TABLEID
    !parametervalue = mvar_lngBHolRegionTableID
    .Update

    ' Save the Absence BHol Region column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLREGION
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_BHOLREGION
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngBHolRegionID
    .Update

    ' Save the Absence BHol table ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLTABLE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_BHOLTABLE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_TABLEID
    !parametervalue = mvar_lngBHolTableID
    .Update

    ' Save the Absence BHol Date column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLDATE
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_BHOLDATE
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngBHolDateID
    .Update

    ' Save the Absence BHol Description column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLDESCRIPTION
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_ABSENCE
      !parameterkey = gsPARAMETERKEY_BHOLDESCRIPTION
    Else
      .Edit
    End If
    !ParameterType = gsPARAMETERTYPE_COLUMNID
    !parametervalue = mvar_lngBHolDescriptionID
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

    ' Get the BHol Region table ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLREGIONTABLE
    If .NoMatch Then
      mvar_lngBHolRegionTableID = 0
    Else
      mvar_lngBHolRegionTableID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the BHol Region column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLREGION
    If .NoMatch Then
      mvar_lngBHolRegionID = 0
    Else
      mvar_lngBHolRegionID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the BHol table ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLTABLE
    If .NoMatch Then
      mvar_lngBHolTableID = 0
    Else
      mvar_lngBHolTableID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the BHol Date column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLDATE
    If .NoMatch Then
      mvar_lngBHolDateID = 0
    Else
      mvar_lngBHolDateID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

    ' Get the BHol Description column ID.
    .Seek "=", gsMODULEKEY_ABSENCE, gsPARAMETERKEY_BHOLDESCRIPTION
    If .NoMatch Then
      mvar_lngBHolDescriptionID = 0
    Else
      mvar_lngBHolDescriptionID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

  End With

End Sub


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


