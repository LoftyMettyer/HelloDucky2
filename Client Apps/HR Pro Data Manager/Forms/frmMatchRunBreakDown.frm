VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmMatchRunBreakDown 
   Caption         =   "Match Breakdown"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1016
   Icon            =   "frmMatchRunBreakDown.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   120
      TabIndex        =   4
      Top             =   1000
      Width           =   5895
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   400
         Left            =   4695
         TabIndex        =   6
         Top             =   3195
         Width           =   1200
      End
      Begin VB.ComboBox cboRelation 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   0
         Width           =   3480
      End
      Begin SSDataWidgets_B.SSDBGrid grdBreakdown 
         Height          =   2655
         Left            =   0
         TabIndex        =   7
         Top             =   435
         Width           =   5865
         _Version        =   196617
         DataMode        =   2
         RecordSelectors =   0   'False
         GroupHeaders    =   0   'False
         Col.Count       =   2
         AllowUpdate     =   0   'False
         MultiLine       =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   1
         SelectByCell    =   -1  'True
         BalloonHelp     =   0   'False
         MaxSelectedRows =   1
         ForeColorEven   =   0
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Caption=   "Record Description"
         Columns(0).Name =   "Record Description"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   3200
         Columns(1).Caption=   "Intersection Value"
         Columns(1).Name =   "Intersection Value"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   10345
         _ExtentY        =   4683
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
      Begin VB.CheckBox chkAllRecords 
         Caption         =   "&Include non-matching records"
         Height          =   195
         Left            =   0
         TabIndex        =   8
         Top             =   3255
         Width           =   3435
      End
      Begin VB.Label lblTables 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tables :"
         Height          =   195
         Left            =   0
         TabIndex        =   9
         Top             =   60
         Width           =   570
      End
   End
   Begin VB.ComboBox cboTable1 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   200
      Width           =   3480
   End
   Begin VB.ComboBox cboTable2 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   3480
   End
   Begin VB.Label lblTable2Name 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<Table 2 Name> :"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   660
      Width           =   1320
   End
   Begin VB.Label lblTable1Name 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<Table 1 Name> :"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   255
      Width           =   1320
   End
End
Attribute VB_Name = "frmMatchRunBreakDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmParent As frmMatchRun
Private colRecDesc1 As Collection
Private colRecDesc2 As Collection
Private mlngTable1RecDescExprID As Long
Private mlngTable2RecDescExprID As Long
Private mblnLoading As Boolean

'Private mcolCrossRefArrays As Collection
Private mcolRecDesc1 As Collection
Private mcolRecDesc2 As Collection
Private mlngCrossRefArray() As Long

Private Const lngGap As Long = 120

Public Property Let ParentForm(frmParent As frmMatchRun)
  Set mfrmParent = frmParent
End Property

Public Property Let Loading(blnNewValue As Boolean)
  mblnLoading = blnNewValue
End Property

Public Property Let Table1RecDescExprID(ByVal lngNewValue As Long)
  mlngTable1RecDescExprID = lngNewValue
  Set colRecDesc1 = New Collection
End Property

Public Property Let Table2RecDescExprID(ByVal lngNewValue As Long)
  mlngTable2RecDescExprID = lngNewValue
  Set colRecDesc2 = New Collection
End Property

Public Sub AddToCrossRef(lngID1 As Long, lngID2 As Long)

  Dim strRecDesc As String
  'Dim lngTemp() As Long
  Dim lngIndex As Long
  
  On Local Error Resume Next
  
  
  lngIndex = UBound(mlngCrossRefArray, 2) + 1
  ReDim Preserve mlngCrossRefArray(1, lngIndex)
  mlngCrossRefArray(0, lngIndex) = lngID1
  mlngCrossRefArray(1, lngIndex) = lngID2
  
  
  strRecDesc = GetRecordDesc(mlngTable1RecDescExprID, lngID1)
  mcolRecDesc1.Add strRecDesc, "ID" & CStr(lngID1)
  If lngIndex = 1 Or mlngTable2RecDescExprID = 0 Then
    cboTable1.AddItem strRecDesc
    cboTable1.ItemData(cboTable1.NewIndex) = lngID1
  End If
  
  If lngID2 > 0 Then
    strRecDesc = GetRecordDesc(mlngTable2RecDescExprID, lngID2)
    mcolRecDesc2.Add strRecDesc, "ID" & CStr(lngID2)
  End If

  'If Not AlreadyInCollection(1, lngID1) Then
    'If strRecDesc <> vbNullString Then
    '  colRecDesc1.Add strRecDesc, "ID" & CStr(lngID1)
    '  cboTable1.AddItem strRecDesc
    '  cboTable1.ItemData(cboTable1.NewIndex) = lngID1
    'End If
    
    'ReDim lngTemp(0) As Long
    'lngTemp(0) = lngID2
    'mcolCrossRefArrays.Add lngTemp, "ID" & CStr(lngID1)
    'mcolRecDesc1.Add strRecDesc, "ID1-" & CStr(lngID1)
  
  'Else
    'lngTemp = mcolCrossRefArrays("ID" & CStr(lngID1))
    'lngIndex = UBound(lngTemp) + 1
    'ReDim Preserve lngTemp(lngIndex) As Long
    'lngTemp(lngIndex) = lngID2
    'mcolCrossRefArrays.Remove "ID" & CStr(lngID1)
    'mcolCrossRefArrays.Add lngTemp, "ID" & CStr(lngID1)

  'End If

  

  'If Not AlreadyInCollection(2, lngID2) Then
  '  strRecDesc = GetRecordDesc(mlngTable2RecDescExprID, lngID2)
  '  If strRecDesc <> vbNullString Then
  '    colRecDesc2.Add strRecDesc, "ID" & CStr(lngID2)
  '  End If
  'End If

End Sub

'Private Function AlreadyInCollection(lngIndex As Long, lngID As Long) As Boolean
'
'  'Dim strTemp As String
'
'  On Local Error Resume Next
'
'  'If lngIndex = 1 Then
'  '  strTemp = colRecDesc1("ID" & CStr(lngID))
'  'Else
'  '  strTemp = colRecDesc2("ID" & CStr(lngID))
'  'End If
'  'AlreadyInCollection = True
'
'  AlreadyInCollection = False
'  If lngIndex = 1 Then
'    AlreadyInCollection = (colRecDesc1("ID" & CStr(lngID)) <> vbNullString)
'  Else
'    AlreadyInCollection = (colRecDesc2("ID" & CStr(lngID)) <> vbNullString)
'  End If
'
'End Function

Private Function GetRecordDesc(lngRecDescExprID As Long, lngRecordID As Long) As String

  ' Return TRUE if the user has been granted the given permission.
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter

  On Error GoTo LocalErr
  
  If lngRecDescExprID < 1 Then
    'GetRecordDesc = "Record Description Undefined"
    GetRecordDesc = vbNullString
    Exit Function
  End If
  
  
  ' Check if the user can create New instances of the given category.
  Set cmADO = New ADODB.Command
  With cmADO
    .CommandText = "dbo.sp_ASRExpr_" & lngRecDescExprID
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon

    Set pmADO = .CreateParameter("Result", adVarChar, adParamOutput, VARCHAR_MAX_Size)
    .Parameters.Append pmADO

    Set pmADO = .CreateParameter("RecordID", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.Value = lngRecordID

    cmADO.Execute

    GetRecordDesc = .Parameters(0).Value
  
  End With
  Set cmADO = Nothing

Exit Function

LocalErr:
  'COAMsgBox "Error reading record description" & vbCr & _
         "(ID = " & CStr(lngRecordID) & ", Record Description = " & CStr(lngRecDescExprID)
  'fOK = False

End Function

Private Sub cboRelation_Click()
  PopulateGridBreakdown
End Sub

Private Sub cboTable1_Click()

  'Dim lngTemp() As Long
  Dim lngTempID As Long
  Dim lngIndex As Long

  'lngTemp = mcolCrossRefArrays("ID" & cboTable1.ItemData(cboTable1.ListIndex))
  If mblnLoading = True Then
    Exit Sub
  End If
  mblnLoading = True

  
  If mlngTable2RecDescExprID > 0 Then
    With cboTable2
      If .ListIndex <> -1 Then
        lngTempID = .ItemData(.ListIndex)
      End If
      .Clear

      'For lngIndex = 0 To UBound(lngTemp)
      '  .AddItem colRecDesc2("ID" & CStr(lngTemp(lngIndex)))
      '  .ItemData(.NewIndex) = CStr(lngTemp(lngIndex))
      'Next

      For lngIndex = 0 To UBound(mlngCrossRefArray, 2)
        
        If mlngCrossRefArray(0, lngIndex) = cboTable1.ItemData(cboTable1.ListIndex) Then
          .AddItem mcolRecDesc2("ID" & CStr(mlngCrossRefArray(1, lngIndex)))
          .ItemData(.NewIndex) = mlngCrossRefArray(1, lngIndex)
          If mlngCrossRefArray(1, lngIndex) = lngTempID Then
            .ListIndex = .NewIndex
          End If
        End If
      Next

      'If .ListCount > 0 Then
      '  .ListIndex = 0
      'End If
    End With
  End If

  mblnLoading = False
  PopulateGridBreakdown

End Sub

Private Sub cboTable2_Click()
  
  'Dim lngTemp() As Long
  Dim lngTempID As Long
  Dim lngIndex As Long

  'lngTemp = mcolCrossRefArrays("ID" & cboTable1.ItemData(cboTable1.ListIndex))
  If mblnLoading = True Then
    Exit Sub
  End If
  mblnLoading = True

  
  With cboTable1
    
    If .ListIndex <> -1 Then
      lngTempID = .ItemData(.ListIndex)
    End If
    
    .Clear

    'For lngIndex = 0 To UBound(lngTemp)
    '  .AddItem colRecDesc2("ID" & CStr(lngTemp(lngIndex)))
    '  .ItemData(.NewIndex) = CStr(lngTemp(lngIndex))
    'Next

    For lngIndex = 1 To UBound(mlngCrossRefArray, 2)
      If mlngCrossRefArray(1, lngIndex) = cboTable2.ItemData(cboTable2.ListIndex) Then
        .AddItem mcolRecDesc1("ID" & CStr(mlngCrossRefArray(0, lngIndex)))
        .ItemData(.NewIndex) = mlngCrossRefArray(0, lngIndex)
        If mlngCrossRefArray(0, lngIndex) = lngTempID Then
          .ListIndex = .NewIndex
        End If
      End If
    Next

    'If .ListCount > 0 Then
    '  .ListIndex = 0
    'End If
  End With

  mblnLoading = False
  PopulateGridBreakdown

End Sub

Private Sub chkAllRecords_Click()
  PopulateGridBreakdown
End Sub

Private Sub cmdOK_Click()
  Me.Hide
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
  'Set mcolCrossRefArrays = New Collection
  ReDim mlngCrossRefArray(1, 0)
  Set mcolRecDesc1 = New Collection
  Set mcolRecDesc2 = New Collection
  Frame1.BackColor = Me.BackColor
  
  Dim lngMinWidth As Long
  Dim lngMinHeight As Long
  Dim lngComboLeft As Long
  
  'JPD 20030908 Fault 5756
  DisplayApplication
  
  lngComboLeft = lblTables.Width
  If lngComboLeft < lblTable1Name.Width Then
    lngComboLeft = lblTable1Name.Width
  End If
  If lngComboLeft < lblTable2Name.Width Then
    lngComboLeft = lblTable2Name.Width
  End If
  lngComboLeft = lngComboLeft + lblTables.Left + (lngGap * 3)

  lngMinWidth = lngComboLeft + 3000 + lngGap
  lngMinHeight = cboRelation.Top + cmdOK.Height + cmdOK.Height + (lngGap * 2) + 2500

  Hook Me.hWnd, lngMinWidth, lngMinHeight
End Sub

Public Sub ShowBreakdown(lngTable1ID As Long, lngTable2ID As Long, lngMatchReportType As MatchReportType)

  Dim lngIndex As Long

  'mblnLoading = True

  If mlngTable2RecDescExprID > 0 Then
    With cboTable1
      .Clear
      .AddItem mcolRecDesc1("ID" & CStr(lngTable1ID))
      .ItemData(.NewIndex) = lngTable1ID
      .ListIndex = 0
    End With
  
    For lngIndex = 0 To cboTable2.ListCount - 1
      If lngTable2ID = cboTable2.ItemData(lngIndex) Then
        'MH20030529 Fault 5436
        If cboTable2.ListIndex = lngIndex Then
          cboTable2_Click
        Else
          cboTable2.ListIndex = lngIndex
        End If
        Exit For
      End If
    Next
  End If

  'mblnLoading = True
  
  
  For lngIndex = 0 To cboTable1.ListCount - 1
    If lngTable1ID = cboTable1.ItemData(lngIndex) Then
      'MH20030529 Fault 5436
      If cboTable1.ListIndex = lngIndex Then
        cboTable1_Click
      Else
        cboTable1.ListIndex = lngIndex
      End If
      Exit For
    End If
  Next

  'mblnLoading = False
  
  'EnableCombo cboTable1
  'EnableCombo cboTable2
  'EnableCombo cboRelation

  chkAllRecords.Value = IIf(lngMatchReportType <> mrtNormal, vbChecked, vbUnchecked)
  Me.Show vbModal

End Sub

Private Sub PopulateGridBreakdown()
  
  Dim lngRelationID As Long
  Dim lngTable1ID As Long
  Dim lngTable2ID As Long

  Screen.MousePointer = vbHourglass

  If mblnLoading = False And cboTable1.ListIndex <> -1 Then 'And cboTable2.ListIndex <> -1 Then
    lngRelationID = cboRelation.ItemData(cboRelation.ListIndex)
    lngTable1ID = cboTable1.ItemData(cboTable1.ListIndex)
    If cboTable2.ListIndex >= 0 Then
      lngTable2ID = cboTable2.ItemData(cboTable2.ListIndex)
    End If
    
    mfrmParent.GetRecordsetBreakdown lngRelationID, lngTable1ID, lngTable2ID
    mfrmParent.PopulateGridBreakdown lngRelationID
  End If

  Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    Cancel = True
    Me.Hide
  End If
End Sub

Private Sub Form_Resize()

'  Dim lngMinWidth As Long
'  Dim lngMinHeight As Long
  Dim lngComboLeft As Long
  
  'JPD 20030908 Fault 5756
  DisplayApplication
  
  lngComboLeft = lblTables.Width
  If lngComboLeft < lblTable1Name.Width Then
    lngComboLeft = lblTable1Name.Width
  End If
  If lngComboLeft < lblTable2Name.Width Then
    lngComboLeft = lblTable2Name.Width
  End If
  lngComboLeft = lngComboLeft + lblTables.Left + (lngGap * 3)

'  lngMinWidth = lngComboLeft + 3000 + lngGap
'  lngMinHeight = cboRelation.Top + cmdOK.Height + cmdOK.Height + (lngGap * 2) + 2500

'  If Me.WindowState = vbNormal Then
'    If Me.Height > Screen.Height Then Me.Height = lngMinHeight
'    If Me.Width > Screen.Width Then Me.Width = lngMinWidth
'    If Me.Height < lngMinHeight Then Me.Height = lngMinHeight
'    If Me.Width < lngMinWidth Then Me.Width = lngMinWidth
'  End If

  Frame1.Height = Me.ScaleHeight - (Frame1.Top + 120)
  Frame1.Width = Me.ScaleWidth - 240

  cboTable1.Left = lngComboLeft
  cboTable1.Width = Me.ScaleWidth - (lngComboLeft + 120)

  cboTable2.Left = lngComboLeft
  cboTable2.Width = cboTable1.Width

  cboRelation.Left = lngComboLeft - 120
  cboRelation.Width = cboTable1.Width

  'cmdOK.Left = Me.ScaleWidth - (cmdOK.Width + 120)
  'cmdOK.Top = Me.ScaleHeight - (cmdOK.Height + 120)
  cmdOK.Left = Frame1.Width - (cmdOK.Width + 10)
  cmdOK.Top = Frame1.Height - (cmdOK.Height + 10)

  grdBreakdown.Top = cboRelation.Top + 400
  'grdBreakdown.Width = Me.ScaleWidth - 240
  grdBreakdown.Width = Frame1.Width
  grdBreakdown.Height = cmdOK.Top - (grdBreakdown.Top + 120)

  chkAllRecords.Top = cmdOK.Top + 60

  ' Get rid of the icon off the form
  RemoveIcon Me

End Sub

'Private Function EnableCombo(cboTemp As ComboBox)
'  With cboTemp
'    .Enabled = (.ListCount > 1)
'    .BackColor = IIf(.ListCount > 1, vbWindowBackground, vbButtonFace)
'  End With
'End Function

Private Sub Form_Unload(Cancel As Integer)
  Unhook Me.hWnd
End Sub

