VERSION 5.00
Object = "{1EE59219-BC23-4BDF-BB08-D545C8A38D6D}#1.1#0"; "COA_Line.ocx"
Begin VB.Form frmMaternitySetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maternity Setup"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5047
   Icon            =   "frmMaternitySetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraAbsence 
      Caption         =   "Maternity :"
      Height          =   2620
      Left            =   120
      TabIndex        =   2
      Top             =   90
      Width           =   5160
      Begin VB.ComboBox cboLeaveType 
         Height          =   315
         Left            =   2505
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1320
         Width           =   2500
      End
      Begin VB.ComboBox cboBabyBirthDate 
         Height          =   315
         Left            =   2505
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2130
         Width           =   2500
      End
      Begin VB.ComboBox cboLeaveStart 
         Height          =   315
         Left            =   2505
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1725
         Width           =   2500
      End
      Begin VB.ComboBox cboEWCDate 
         Height          =   315
         Left            =   2505
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   920
         Width           =   2500
      End
      Begin VB.ComboBox cboMaternityTable 
         Height          =   315
         Left            =   2505
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   300
         Width           =   2500
      End
      Begin COALine.COA_Line ASRDummyLine1 
         Height          =   30
         Left            =   180
         Top             =   755
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   53
      End
      Begin VB.Label lblReason 
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Type Column :"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   1400
         Width           =   2010
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Baby Birth Date Column :"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   2190
         Width           =   1800
      End
      Begin VB.Label lblStartSession 
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Start Column :"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   1785
         Width           =   2070
      End
      Begin VB.Label lblStartDate 
         BackStyle       =   0  'Transparent
         Caption         =   "BDW Date Column :"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   975
         Width           =   1875
      End
      Begin VB.Label lblAbsenceTable 
         BackStyle       =   0  'Transparent
         Caption         =   "Maternity Table :"
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   360
         Width           =   1650
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4080
      TabIndex        =   1
      Top             =   2800
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2780
      TabIndex        =   0
      Top             =   2800
      Width           =   1200
   End
End
Attribute VB_Name = "frmMaternitySetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnReadOnly As Boolean
Private mvar_lngPersonnelTableID As Long

Private mvar_lngOriginalMaternityTableID As Long
Private mbLoading As Boolean
Private mfChanged As Boolean

Public Property Get Changed() As Boolean
  Changed = mfChanged
End Property
Public Property Let Changed(ByVal pblnChanged As Boolean)
  mfChanged = pblnChanged
  RefreshControls
End Property
Private Sub cboBabyBirthDate_Click()
Changed = True
End Sub

Private Sub cboEWCDate_Click()
Changed = True
End Sub
Private Sub cboLeaveStart_Click()
Changed = True
End Sub

Private Sub cboLeaveType_Click()
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

  'AE20071119 Fault #12607
  'If ValidateSetup Then
'    SaveParam gsPARAMETERKEY_MATERNITYTABLE, gsPARAMETERTYPE_TABLEID, GetComboItem(cboMaternityTable)
'    SaveParam gsPARAMETERKEY_MATERNITYEWCDATECOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboEWCDate)
'    SaveParam gsPARAMETERKEY_MATERNITYLEAVETYPECOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboLeaveType)
'    SaveParam gsPARAMETERKEY_MATERNITYLEAVESTARTCOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboLeaveStart)
'    SaveParam gsPARAMETERKEY_MATERNITYBABYBIRTHCOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboBabyBirthDate)
'    Application.Changed = True
  If SaveChanges Then
    Changed = False
    UnLoad Me
  End If
  
End Sub

Private Sub PopulateBaseTableCombos()
  
  Dim lngPostTable As Long
  
  cboMaternityTable.Clear
  cboMaternityTable.AddItem "<None>"
  cboMaternityTable.ItemData(cboMaternityTable.NewIndex) = 0
  cboMaternityTable.ListIndex = 0

  ' Add items to the combo for each table that has not been deleted,
  ' and is a Lookup table.
  With recTabEdit
    .Index = "idxName"
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If

    Do While Not .EOF
      If Not !Deleted Then
        recRelEdit.Index = "idxParentID"
        recRelEdit.Seek "=", mvar_lngPersonnelTableID, !TableID

        If Not recRelEdit.NoMatch Then

          'If HasTypeOfColumn(!TableID, dtTIMESTAMP) Then
          '  If HasTypeOfColumn(!TableID, rdTypeVARCHAR) Then
              cboMaternityTable.AddItem !TableName
              cboMaternityTable.ItemData(cboMaternityTable.NewIndex) = !TableID
          '  End If
          'End If

        End If
      End If

      .MoveNext
    Loop
  End With

End Sub


Private Sub cboMaternityTable_Click()

  Dim lngMaternityTable As Long
  Dim blnEnabled As Boolean

  Dim lngEWCDateCol As Long
  Dim lngLeaveTypeCol As Long
  Dim lngLeaveStartCol As Long
  Dim lngBabyBirthCol As Long


  cboEWCDate.Clear
  cboLeaveType.Clear
  cboLeaveStart.Clear
  cboBabyBirthDate.Clear
  
  If cboMaternityTable.ListIndex > 0 Then
    lngMaternityTable = cboMaternityTable.ItemData(cboMaternityTable.ListIndex)
  End If
  
  If lngMaternityTable > 0 Then
    
    lngEWCDateCol = ReadParam(gsPARAMETERKEY_MATERNITYEWCDATECOLUMN)
    lngLeaveTypeCol = ReadParam(gsPARAMETERKEY_MATERNITYLEAVETYPECOLUMN)
    lngLeaveStartCol = ReadParam(gsPARAMETERKEY_MATERNITYLEAVESTARTCOLUMN)
    lngBabyBirthCol = ReadParam(gsPARAMETERKEY_MATERNITYBABYBIRTHCOLUMN)
    
    
    With recColEdit
      .Index = "idxName"
      .Seek ">=", lngMaternityTable
  
      If Not .NoMatch Then
        ' Add items to the combos for each column that has not been deleted,
        ' or is a system or link column.
        Do While Not .EOF
          If !TableID <> lngMaternityTable Then
            Exit Do
          End If
  
          If (Not !Deleted) And _
            (!columntype <> giCOLUMNTYPE_LINK) And _
            (!columntype <> giCOLUMNTYPE_SYSTEM) Then

            If !DataType = dtTIMESTAMP Then
              cboEWCDate.AddItem !ColumnName
              cboEWCDate.ItemData(cboEWCDate.NewIndex) = !ColumnID
              If !ColumnID = lngEWCDateCol Then
                cboEWCDate.ListIndex = cboEWCDate.NewIndex
              End If
            
              cboLeaveStart.AddItem !ColumnName
              cboLeaveStart.ItemData(cboLeaveStart.NewIndex) = !ColumnID
              If !ColumnID = lngLeaveStartCol Then
                cboLeaveStart.ListIndex = cboLeaveStart.NewIndex
              End If
            
              cboBabyBirthDate.AddItem !ColumnName
              cboBabyBirthDate.ItemData(cboBabyBirthDate.NewIndex) = !ColumnID
              If !ColumnID = lngBabyBirthCol Then
                cboBabyBirthDate.ListIndex = cboBabyBirthDate.NewIndex
              End If
            End If
            
            
            If !DataType = dtVARCHAR Then
              cboLeaveType.AddItem !ColumnName
              cboLeaveType.ItemData(cboLeaveType.NewIndex) = !ColumnID
              If !ColumnID = lngLeaveTypeCol Then
                cboLeaveType.ListIndex = cboLeaveType.NewIndex
              End If
            End If
  
          End If
  
          .MoveNext
        Loop
      End If
    End With
  
  End If


  blnEnabled = (cboEWCDate.ListCount > 0)
  cboEWCDate.Enabled = blnEnabled
  cboEWCDate.BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)
  If blnEnabled And cboEWCDate.ListIndex < 0 Then
    cboEWCDate.ListIndex = 0
  End If
  
  blnEnabled = (cboLeaveType.ListCount > 0)
  cboLeaveType.Enabled = blnEnabled
  cboLeaveType.BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)
  If blnEnabled And cboLeaveType.ListIndex < 0 Then
    cboLeaveType.ListIndex = 0
  End If
  
  blnEnabled = (cboLeaveStart.ListCount > 0)
  cboLeaveStart.Enabled = blnEnabled
  cboLeaveStart.BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)
  If blnEnabled And cboLeaveStart.ListIndex < 0 Then
    cboLeaveStart.ListIndex = 0
  End If
  
  blnEnabled = (cboBabyBirthDate.ListCount > 0)
  cboBabyBirthDate.Enabled = blnEnabled
  cboBabyBirthDate.BackColor = IIf(blnEnabled, vbWindowBackground, vbButtonFace)
  If blnEnabled And cboBabyBirthDate.ListIndex < 0 Then
    cboBabyBirthDate.ListIndex = 0
  End If

Changed = True

End Sub


Private Sub SaveParam(strKey As String, strType As String, lngValue As Long)

  With recModuleSetup

    .Index = "idxModuleParameter"
    .Seek "=", gsMODULEKEY_MATERNITY, strKey
    If .NoMatch Then
      .AddNew
      !moduleKey = gsMODULEKEY_MATERNITY
      !parameterkey = strKey
    Else
      .Edit
    End If
    !ParameterType = strType
    !parametervalue = lngValue
    .Update

  End With

End Sub

Private Function ValidateSetup() As Boolean
  'JPD 20040106 Fault 7894
  
  On Error GoTo ValidateError

  Dim fSpecialFunctionUsed As Boolean
  Dim sSQL As String
  Dim rsCheck As DAO.Recordset
  Dim objComp As CExprComponent
  Dim lngExprID As Long
  Dim objExpr As CExpression
  Dim sExprName As String
  Dim sExprType As String
  Dim sExprParentTable As String
  Dim lngExprBaseTableID As Long
  Dim sFunctionName As String
  Dim lngFunctionID As Long
  Dim objFunctionDef As clsFunctionDef

  ' Don't allow the Maternity table to change if the special
  ' functions are used anywhere.
  fSpecialFunctionUsed = False

  If (mvar_lngOriginalMaternityTableID <> GetComboItem(cboMaternityTable)) Then
    ' Find any expression field components that use the special functions.
    sSQL = "SELECT tmpComponents.componentID, tmpComponents.functionID" & _
      " FROM tmpComponents" & _
      " WHERE tmpComponents.type = " & Trim(Str(giCOMPONENT_FUNCTION)) & _
      " AND tmpComponents.functionID IN (64)"
    Set rsCheck = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    If Not (rsCheck.BOF And rsCheck.EOF) Then
      fSpecialFunctionUsed = True

      Set objComp = New CExprComponent
      objComp.ComponentID = rsCheck!ComponentID
      lngExprID = objComp.RootExpressionID
      Set objComp = Nothing

      ' Get the expression name and type description.
      Set objExpr = New CExpression
      objExpr.ExpressionID = lngExprID

      If objExpr.ReadExpressionDetails Then
        sExprName = objExpr.Name
        sExprType = objExpr.ExpressionTypeName
        lngExprBaseTableID = objExpr.BaseTableID

        ' Get the expression's parent table name.
        recTabEdit.Index = "idxTableID"
        recTabEdit.Seek "=", lngExprBaseTableID

        If Not recTabEdit.NoMatch Then
          sExprParentTable = recTabEdit!TableName
        End If
      Else
        sExprName = "<unknown>"
        sExprType = "<unknown>"
        sExprParentTable = "<unknown>"
      End If

      ' Disassociate object variables.
      Set objExpr = Nothing

      gobjFunctionDefs.Initialise
      sFunctionName = "<unknown>"
      If gobjFunctionDefs.IsValidID(rsCheck!FunctionID) Then
        Set objFunctionDef = gobjFunctionDefs.Item("F" & Trim(Str(rsCheck!FunctionID)))
        sFunctionName = objFunctionDef.Name
        Set objFunctionDef = Nothing
      End If
    End If
    ' Close the recordset.
    rsCheck.Close

    If fSpecialFunctionUsed Then
      MsgBox "The 'Maternity' table cannot be changed." & vbCrLf & vbCrLf & _
        "It is used as the base table for the '" & sFunctionName & "' function which is used in the " & sExprType & " '" & sExprName & "', " & _
        "which is owned by the '" & sExprParentTable & "' table.", _
        vbExclamation + vbOKOnly, App.Title
      ValidateSetup = False
      Exit Function
    End If
  End If

  ValidateSetup = True
  Exit Function

ValidateError:

  MsgBox "Error validating the module setup." & vbCrLf & _
         Err.Description, vbExclamation + vbOKOnly, App.Title
  ValidateSetup = False

End Function

Private Function SaveChanges() As Boolean
  'AE20071119 Fault #12607
  SaveChanges = False
  
  If Not ValidateSetup Then
    Exit Function
  End If
  
  Screen.MousePointer = vbHourglass
  
  SaveParam gsPARAMETERKEY_MATERNITYTABLE, gsPARAMETERTYPE_TABLEID, GetComboItem(cboMaternityTable)
  SaveParam gsPARAMETERKEY_MATERNITYEWCDATECOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboEWCDate)
  SaveParam gsPARAMETERKEY_MATERNITYLEAVETYPECOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboLeaveType)
  SaveParam gsPARAMETERKEY_MATERNITYLEAVESTARTCOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboLeaveStart)
  SaveParam gsPARAMETERKEY_MATERNITYBABYBIRTHCOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboBabyBirthDate)
  
  'AE20071119 Fault #12607
  SaveChanges = True
  Application.Changed = True
  
  Screen.MousePointer = vbDefault
End Function

Private Function ReadParam(strKey As String) As Long

  With recModuleSetup
    .Index = "idxModuleParameter"

    .Seek "=", gsMODULEKEY_MATERNITY, strKey
    If .NoMatch Then
      ReadParam = 0
    Else
      ReadParam = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If

  End With

End Function

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

  mbLoading = True
  
  Screen.MousePointer = vbHourglass

  mblnReadOnly = (Application.AccessMode <> accFull And _
                  Application.AccessMode <> accSupportMode)

  If mblnReadOnly Then
    ControlsDisableAll Me
  End If

  If CheckPersonnelTable Then
    PopulateBaseTableCombos
    InitialiseCombos
  End If
  mbLoading = False
  Changed = False
  Screen.MousePointer = vbDefault

End Sub


Private Sub InitialiseCombos()

  SetComboItem cboMaternityTable, ReadParam(gsPARAMETERKEY_MATERNITYTABLE)
  SetComboItem cboEWCDate, ReadParam(gsPARAMETERKEY_MATERNITYEWCDATECOLUMN)
  SetComboItem cboLeaveType, ReadParam(gsPARAMETERKEY_MATERNITYLEAVETYPECOLUMN)
  SetComboItem cboLeaveStart, ReadParam(gsPARAMETERKEY_MATERNITYLEAVESTARTCOLUMN)
  SetComboItem cboBabyBirthDate, ReadParam(gsPARAMETERKEY_MATERNITYBABYBIRTHCOLUMN)

  mvar_lngOriginalMaternityTableID = GetComboItem(cboMaternityTable)
  
End Sub


Private Function CheckPersonnelTable() As Boolean
  
  ' Read the Absence parameter values from the database into local variables.
  Dim iTemp As Integer
  Dim iLoop As Integer
  
  With recModuleSetup
    .Index = "idxModuleParameter"

    ' Get the Personnel table ID.
    .Seek "=", gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE
    If .NoMatch Then
      mvar_lngPersonnelTableID = 0
    Else
      mvar_lngPersonnelTableID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, !parametervalue)
    End If
    
    If mvar_lngPersonnelTableID = 0 Then
      MsgBox "The Personnel module has not been configured." & vbCrLf & _
        "The Maternity module requires the Personnel table to be defined.", _
        vbExclamation + vbOKOnly, Application.Name
    End If

  End With

  CheckPersonnelTable = (mvar_lngPersonnelTableID > 0)

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


Private Sub RefreshControls()
If Not mbLoading Then cmdOK.Enabled = mfChanged
End Sub
