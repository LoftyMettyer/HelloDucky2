VERSION 5.00
Begin VB.Form frmPostSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Post"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5045
   Icon            =   "frmPostSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Post / Job :"
      Height          =   1560
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4695
      Begin VB.ComboBox cboPostTable 
         Height          =   315
         ItemData        =   "frmPostSetup.frx":000C
         Left            =   1500
         List            =   "frmPostSetup.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   300
         Width           =   3000
      End
      Begin VB.ComboBox cboJobTitle 
         Height          =   315
         ItemData        =   "frmPostSetup.frx":0010
         Left            =   1500
         List            =   "frmPostSetup.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   705
         Width           =   3000
      End
      Begin VB.ComboBox cboGrade 
         Height          =   315
         ItemData        =   "frmPostSetup.frx":0014
         Left            =   1500
         List            =   "frmPostSetup.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1095
         Width           =   3000
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Job Title :"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   765
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Table :"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Grade :"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1160
         Width           =   540
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Grade :"
      Height          =   1560
      Left            =   120
      TabIndex        =   2
      Top             =   1710
      Width           =   4695
      Begin VB.ComboBox cboHeirarchy 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1100
         Width           =   3000
      End
      Begin VB.TextBox txtGradeTable 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1500
         TabIndex        =   4
         Top             =   300
         Width           =   3000
      End
      Begin VB.TextBox txtGradeColumn 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1500
         TabIndex        =   3
         Top             =   700
         Width           =   3000
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Hierarchy :"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1155
         Width           =   795
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Table :"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Grade :"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   765
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3615
      TabIndex        =   1
      Top             =   3375
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   2325
      TabIndex        =   0
      Top             =   3375
      Width           =   1200
   End
End
Attribute VB_Name = "frmPostSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnReadOnly As Boolean
Private mfChanged As Boolean
Private mbLoading As Boolean

Public Property Get Changed() As Boolean
  Changed = mfChanged
End Property
Public Property Let Changed(ByVal pblnChanged As Boolean)
  mfChanged = pblnChanged
  If Not mbLoading Then cmdOK.Enabled = True
End Property

Private Sub cboJobTitle_Click()
  Changed = True
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
'        Screen.MousePointer = vbDefault
'        Exit Sub
'      ElseIf pintAnswer = vbCancel Then
'        Exit Sub
'      End If
'    End If
'TidyUpAndExit:
  UnLoad Me
End Sub

Private Sub cmdOK_Click()

'  SaveParam gsPARAMETERKEY_POSTTABLE, gsPARAMETERTYPE_TABLEID, GetComboItem(cboPostTable)
'  SaveParam gsPARAMETERKEY_POSTJOBTITLECOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboJobTitle)
'  SaveParam gsPARAMETERKEY_POSTGRADECOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboGrade)
'  SaveParam gsPARAMETERKEY_GRADETABLE, gsPARAMETERTYPE_TABLEID, Val(txtGradeTable.Tag)
'  SaveParam gsPARAMETERKEY_GRADECOLUMN, gsPARAMETERTYPE_COLUMNID, Val(txtGradeColumn.Tag)
'  SaveParam gsPARAMETERKEY_NUMLEVELCOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboHeirarchy)
'  Application.Changed = True
'  UnLoad Me
  If SaveChanges Then
    Changed = False
    UnLoad Me
  End If
  

End Sub

Private Function SaveChanges() As Boolean
  SaveChanges = False
  
  SaveParam gsPARAMETERKEY_POSTTABLE, gsPARAMETERTYPE_TABLEID, GetComboItem(cboPostTable)
  SaveParam gsPARAMETERKEY_POSTJOBTITLECOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboJobTitle)
  SaveParam gsPARAMETERKEY_POSTGRADECOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboGrade)
  SaveParam gsPARAMETERKEY_GRADETABLE, gsPARAMETERTYPE_TABLEID, val(txtGradeTable.Tag)
  SaveParam gsPARAMETERKEY_GRADECOLUMN, gsPARAMETERTYPE_COLUMNID, val(txtGradeColumn.Tag)
  SaveParam gsPARAMETERKEY_NUMLEVELCOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboHeirarchy)
  
  SaveChanges = True
  Application.Changed = True
  
  Screen.MousePointer = vbDefault
End Function

Private Sub SaveParam(strKey As String, strType As String, lngValue As Long)

  With recModuleSetup

    .Index = "idxModuleParameter"
    .Seek "=", gsMODULEKEY_POST, strKey
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_POST
      !parameterkey = strKey
    Else
      .Edit
    End If
    !ParameterType = strType
    !parametervalue = lngValue
    .Update

  End With

End Sub

Private Function GetComboItem(cboTemp As ComboBox) As Long
  GetComboItem = 0
  If cboTemp.ListIndex <> -1 Then
    GetComboItem = cboTemp.ItemData(cboTemp.ListIndex)
  End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_Load()
  Screen.MousePointer = vbHourglass
  
  mbLoading = True
  Changed = False
  cmdOK.Enabled = False
  mblnReadOnly = (Application.AccessMode <> accFull And _
                  Application.AccessMode <> accSupportMode)

  If mblnReadOnly Then
    ControlsDisableAll Me
  End If

  PopulateBaseTableCombos
  InitialiseCombos
  
  'AE20080204 Fault #12829
  mfChanged = False
  mbLoading = False
  Screen.MousePointer = vbDefault
End Sub


Private Function ReadParam(strKey As String) As Long

  With recModuleSetup
    .Index = "idxModuleParameter"

    .Seek "=", gsMODULEKEY_POST, strKey
    If .NoMatch Then
      ReadParam = 0
    Else
      ReadParam = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

  End With

End Function


Private Sub InitialiseCombos()

  SetComboItem cboPostTable, ReadParam(gsPARAMETERKEY_POSTTABLE)
  SetComboItem cboJobTitle, ReadParam(gsPARAMETERKEY_POSTJOBTITLECOLUMN)
  SetComboItem cboGrade, ReadParam(gsPARAMETERKEY_POSTGRADECOLUMN)
  SetComboItem cboHeirarchy, ReadParam(gsPARAMETERKEY_NUMLEVELCOLUMN)

End Sub


Private Sub PopulateBaseTableCombos()
  
  Dim lngPostTable As Long
  
  cboPostTable.Clear
  cboPostTable.AddItem "<None>"
  cboPostTable.ItemData(cboPostTable.NewIndex) = 0
  cboPostTable.ListIndex = 0

  ' Add items to the combo for each table that has not been deleted,
  ' and is a Lookup table.
  With recTabEdit
    .Index = "idxName"
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If

    Do While Not .EOF
      If Not !Deleted Then

        If HasLookupColumn(!TableID) Then
          cboPostTable.AddItem !TableName
          cboPostTable.ItemData(cboPostTable.NewIndex) = !TableID
        End If

      End If
      .MoveNext
    Loop
  End With

End Sub


Private Sub cboPostTable_Click()

  Dim lngPostTable As Long
  Dim lngJobTitleCol As Long
  Dim lngGradeCol As Long
  Dim blnEnabled As Boolean

  Dim lngJobTitleListIndex As Long
  Dim lngGradeListIndex As Long
  

  lngJobTitleListIndex = -1
  lngGradeListIndex = -1


  cboJobTitle.Clear
  cboGrade.Clear
  txtGradeTable.Text = vbNullString
  txtGradeTable.Tag = 0
  txtGradeColumn.Text = vbNullString
  txtGradeColumn.Tag = 0
  
  If cboPostTable.ListIndex > 0 Then
    lngPostTable = cboPostTable.ItemData(cboPostTable.ListIndex)
  End If
  
  If lngPostTable > 0 Then
    
    lngJobTitleCol = ReadParam(gsPARAMETERKEY_POSTJOBTITLECOLUMN)
    lngGradeCol = ReadParam(gsPARAMETERKEY_POSTGRADECOLUMN)
    
    With recColEdit
      .Index = "idxName"
      .Seek ">=", lngPostTable
  
      If Not .NoMatch Then
        ' Add items to the combos for each column that has not been deleted,
        ' or is a system or link column.
        Do While Not .EOF
          If !TableID <> lngPostTable Then
            Exit Do
          End If

          If (Not !Deleted) And _
            (!columntype = giCOLUMNTYPE_LOOKUP) Then
  
            cboGrade.AddItem !ColumnName
            cboGrade.ItemData(cboGrade.NewIndex) = !ColumnID
            If !ColumnID = lngGradeCol Then
              lngGradeListIndex = cboGrade.NewIndex
            End If
  
            ' Load varchar fields into the name combo.
            If !DataType = dtVARCHAR Then
              cboJobTitle.AddItem !ColumnName
              cboJobTitle.ItemData(cboJobTitle.NewIndex) = !ColumnID
              If !ColumnID = lngJobTitleCol Then
                lngJobTitleListIndex = cboJobTitle.NewIndex
              End If
            End If
  
          End If
  
          .MoveNext
        Loop
      End If
    End With
  
  End If


  If lngJobTitleListIndex >= 0 Then
    cboJobTitle.ListIndex = lngJobTitleListIndex
  End If
  If lngGradeListIndex >= 0 Then
    cboGrade.ListIndex = lngGradeListIndex
  End If

  blnEnabled = (cboJobTitle.ListCount > 0)
  cboJobTitle.Enabled = blnEnabled
  cboJobTitle.BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)
  If blnEnabled And cboJobTitle.ListIndex < 0 Then
    cboJobTitle.ListIndex = 0
  End If
  
  cboGrade.ListIndex = lngGradeListIndex
  
  blnEnabled = (cboGrade.ListCount > 0)
  cboGrade.Enabled = blnEnabled
  cboGrade.BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)
  If blnEnabled And cboGrade.ListIndex < 0 Then
    cboGrade.ListIndex = 0
  End If

  Changed = True

End Sub

Private Sub cboGrade_Click()

  Dim lngGradeTable As Long
  Dim lngGradeColumn As Long
  Dim lngPostGradeColumn As Long
  Dim blnEnabled As Boolean

  lngPostGradeColumn = cboGrade.ItemData(cboGrade.ListIndex)
  
  With recColEdit
    .Index = "idxColumnID"
    .Seek "=", lngPostGradeColumn
    If Not .NoMatch Then
      lngGradeTable = !LookupTableID
      lngGradeColumn = !LookupColumnID
    End If
  End With
  
  
  With recTabEdit
    .Index = "idxTableID"
    .Seek "=", lngGradeTable
    If Not .NoMatch Then
      txtGradeTable.Text = !TableName
      txtGradeTable.Tag = lngGradeTable
    Else
      txtGradeTable.Text = vbNullString
      txtGradeTable.Tag = 0
    End If
  End With


  txtGradeColumn.Text = vbNullString
  txtGradeColumn.Tag = 0

  cboHeirarchy.Clear
  With recColEdit
    .Index = "idxName"
    .Seek ">=", lngGradeTable

    If Not .NoMatch Then
      ' Add items to the combos for each column that has not been deleted,
      ' or is a system or link column.
      Do While Not .EOF
        If !TableID <> lngGradeTable Then
          Exit Do
        End If

        
        If (Not !Deleted) And _
          (!columntype <> giCOLUMNTYPE_LINK) And _
          (!columntype <> giCOLUMNTYPE_SYSTEM) Then

          If !ColumnID = lngGradeColumn Then
            txtGradeColumn.Text = !ColumnName
            txtGradeColumn.Tag = lngGradeColumn
          End If

          If !DataType = dtNUMERIC Or !DataType = dtinteger Then
            cboHeirarchy.AddItem !ColumnName
            cboHeirarchy.ItemData(cboHeirarchy.NewIndex) = !ColumnID
          End If

        End If

        .MoveNext
      Loop
    End If
  End With
  
  blnEnabled = (cboHeirarchy.ListCount > 0)
  cboHeirarchy.Enabled = blnEnabled
  cboHeirarchy.BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)
  If blnEnabled And cboHeirarchy.ListIndex < 0 Then
    cboHeirarchy.ListIndex = 0
  End If

  Changed = True

End Sub

Private Sub cboHeirarchy_Click()
  Changed = True
End Sub


Private Function HasLookupColumn(lngTableID As Long) As Boolean
    
  HasLookupColumn = False
  
  With recColEdit
    .Index = "idxName"
    .Seek ">=", lngTableID

    If Not .NoMatch Then
      Do While Not .EOF

        If !TableID <> lngTableID Then
          Exit Do
        End If

        If (Not !Deleted) And (!columntype = giCOLUMNTYPE_LOOKUP) Then
          HasLookupColumn = True
          Exit Function
        End If

        .MoveNext
      Loop
    End If
  End With

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  ' If the user cancels or tries to close the form
  If Changed Then
    Select Case MsgBox("Apply module changes ?", vbYesNoCancel + vbQuestion, Me.Caption)
      Case vbCancel
        Cancel = True
      Case vbYes
        Cancel = (Not SaveChanges)
    End Select
  End If
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


