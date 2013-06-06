VERSION 5.00
Begin VB.Form frmMobileSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mobile Configuration"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMobileSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5865
      TabIndex        =   6
      Top             =   2535
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   4605
      TabIndex        =   5
      Top             =   2535
      Width           =   1200
   End
   Begin VB.Frame fraPersonnelTable 
      Caption         =   "Personnel Table :"
      Height          =   2250
      Left            =   150
      TabIndex        =   2
      Top             =   120
      Width           =   6945
      Begin VB.ComboBox cboLeavingDateColumn 
         Height          =   315
         Left            =   2820
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1530
         Width           =   3975
      End
      Begin VB.ComboBox cboEMailColumn 
         Height          =   315
         Left            =   2820
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   3975
      End
      Begin VB.ComboBox cboPersonnelTable 
         Height          =   315
         Left            =   2820
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   3975
      End
      Begin VB.ComboBox cboLoginName 
         Height          =   315
         Left            =   2820
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1110
         Width           =   3975
      End
      Begin VB.Label lblLeavingDateColumn 
         Caption         =   "Login Expiry Date :"
         Height          =   195
         Left            =   195
         TabIndex        =   10
         Top             =   1575
         Width           =   1995
      End
      Begin VB.Label lblPersonnelTable 
         Caption         =   "Personnel Table :"
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   360
         Width           =   1560
      End
      Begin VB.Label lblLoginNameColumn 
         AutoSize        =   -1  'True
         Caption         =   "Mobile Login Username :"
         Height          =   195
         Left            =   195
         TabIndex        =   8
         Top             =   1170
         Width           =   2115
      End
      Begin VB.Label lblEmailAddresses 
         Caption         =   "Registration Email Address :"
         Height          =   195
         Left            =   195
         TabIndex        =   7
         Top             =   765
         Width           =   2445
      End
   End
End
Attribute VB_Name = "frmMobileSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnReadOnly As Boolean
Private mfChanged As Boolean
Private mbLoading As Boolean

Private mlngPersonnelTableID As Long
Private mlngLoginColumnID As Long
Private mlngUniqueEmailColumnID As Long
Private mlngLeavingDateColumnID As Long

Public Property Get Changed() As Boolean
  Changed = mfChanged
End Property
Public Property Let Changed(ByVal pblnChanged As Boolean)
  mfChanged = pblnChanged
  If Not mbLoading Then cmdOk.Enabled = True
End Property



Private Sub cboEMailColumn_Click()
  With cboEMailColumn
    mlngUniqueEmailColumnID = .ItemData(.ListIndex)
  End With
  
  If Not mbLoading Then
    mbLoading = True
    RefreshPersonnelColumnControls
    mbLoading = False
    Changed = True
  End If
  
  RefreshControls
End Sub



Private Sub cboLeavingDateColumn_Click()
 With cboLeavingDateColumn
    mlngLeavingDateColumnID = .ItemData(.ListIndex)
  End With
  
  If Not mbLoading Then
    mbLoading = True
    RefreshPersonnelColumnControls
    mbLoading = False
    Changed = True
  End If
  
  RefreshControls

End Sub

Private Sub cboLoginName_Click()
  With cboLoginName
    mlngLoginColumnID = .ItemData(.ListIndex)
  End With
  
  If Not mbLoading Then
    mbLoading = True
    RefreshPersonnelColumnControls
    mbLoading = False
    Changed = True
  End If
  
  RefreshControls

End Sub


Private Sub cboPersonnelTable_Click()
  Dim iLoop As Integer
  Dim fAlreadyChanged As Boolean
  Dim fGoAhead As Boolean
  Dim objEmail As clsEmailAddr
  Dim fFixedDelegateEmail As Boolean
  
  fAlreadyChanged = mfChanged
  
  If mlngPersonnelTableID <> cboPersonnelTable.ItemData(cboPersonnelTable.ListIndex) Then
    
    fGoAhead = True
    If (mlngLoginColumnID > 0) _
      Or (mlngUniqueEmailColumnID > 0) _
      Or (mlngLeavingDateColumnID > 0) Then
      
      fGoAhead = (MsgBox("Warning: Changing the Personnel table will reset all other parameters." & vbCrLf & _
      "Are you sure you wish to continue?", _
      vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes)
    End If

    If fGoAhead Then
      mlngPersonnelTableID = cboPersonnelTable.ItemData(cboPersonnelTable.ListIndex)
           
      Changed = True
      RefreshControls
    Else
      For iLoop = 0 To cboPersonnelTable.ListCount - 1
        If mlngPersonnelTableID = cboPersonnelTable.ItemData(iLoop) Then
          cboPersonnelTable.ListIndex = iLoop
          mfChanged = fAlreadyChanged
          Exit For
        End If
      Next iLoop
    End If
  End If
  
  RefreshPersonnelColumnControls
  
End Sub


Private Sub cmdCancel_Click()
  Dim pintAnswer As Integer
    If Changed = True And cmdOk.Enabled Then
      pintAnswer = MsgBox("You have made changes...do you wish to save these changes ?", vbQuestion + vbYesNoCancel, App.Title)
      If pintAnswer = vbYes Then
        'AE20071108 Fault #12551
        'Using Me.MousePointer = vbNormal forces the form to be reloaded
        'after its been unloaded in cmdOK_Click, changed to Screen.MousePointer
        'Me.MousePointer = vbHourglass
        Screen.MousePointer = vbHourglass
        cmdOK_Click 'This is just like saving
        Screen.MousePointer = vbDefault
        'Me.MousePointer = vbNormal
        Exit Sub
      ElseIf pintAnswer = vbCancel Then
        Exit Sub
      End If
    End If
TidyUpAndExit:
  UnLoad Me
End Sub



Private Sub cmdOK_Click()
  Dim fSaveOK As Boolean

  
  fSaveOK = True
  


  
    SaveChanges
  
  UnLoad Me
End Sub

Private Sub SaveChanges()
  ' Save the parameter values to the local database.
  Dim iLoop As Integer
  Dim varBookmark As Variant
  Dim sColumnID As String
  Dim sSQL As String
  

  SaveModuleSetting gsMODULEKEY_MOBILE, gsPARAMETERKEY_PERSONNELTABLE, gsPARAMETERTYPE_TABLEID, mlngPersonnelTableID
  SaveModuleSetting gsMODULEKEY_MOBILE, gsPARAMETERKEY_LOGINNAME, gsPARAMETERTYPE_COLUMNID, mlngLoginColumnID
  SaveModuleSetting gsMODULEKEY_MOBILE, gsPARAMETERKEY_UNIQUEEMAILCOLUMN, gsPARAMETERTYPE_COLUMNID, mlngUniqueEmailColumnID
  SaveModuleSetting gsMODULEKEY_MOBILE, gsPARAMETERKEY_LEAVINGDATE, gsPARAMETERTYPE_COLUMNID, mlngLeavingDateColumnID
    
  Application.Changed = True

End Sub


Private Function ADOConError(objTestConn As ADODB.Connection) As String

  Dim strErrorDesc As String
  Dim lngCount As Long

  strErrorDesc = vbNullString
  If Not objTestConn Is Nothing Then
    If Not objTestConn.Errors Is Nothing Then
      For lngCount = 0 To objTestConn.Errors.Count - 1
        strErrorDesc = objTestConn.Errors(lngCount).Description
      Next
      strErrorDesc = Mid(strErrorDesc, InStrRev(strErrorDesc, "]") + 1)
    End If
  End If

  ADOConError = strErrorDesc

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
  cmdOk.Enabled = False
  
  ' Position the form in the same place it was last time.
  Me.Top = GetPCSetting(Me.Name, "Top", (Screen.Height - Me.Height) / 2)
  Me.Left = GetPCSetting(Me.Name, "Left", (Screen.Width - Me.Width) / 2)
  
  
  mblnReadOnly = (Application.AccessMode <> accFull And _
                  Application.AccessMode <> accSupportMode)
  
  If mblnReadOnly Then
    ControlsDisableAll Me
  End If

  ' Read the current settings from the database.
  ReadParameters
  cboPersonnelTable_Refresh
   
  mfChanged = False
  
  RefreshControls
  
  mbLoading = False
End Sub


Private Sub cboPersonnelTable_Refresh()
  ' Populate the tables combo.
  Dim sSQL As String
  Dim rsTables As DAO.Recordset
  Dim iDefaultItem As Integer
   
  cboPersonnelTable.Clear
  cboPersonnelTable.AddItem "<None>"
  cboPersonnelTable.ItemData(cboPersonnelTable.NewIndex) = 0
  
  iDefaultItem = 0
  
  ' Add the Personnel table and its children (not grand children).
  sSQL = "SELECT tmpTables.tableID, tmpTables.tableName" & _
    " FROM tmpTables" & _
    " WHERE (tmpTables.deleted = FALSE)" & _
    " ORDER BY tmpTables.tableName"
  Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  While Not rsTables.EOF
    cboPersonnelTable.AddItem rsTables!TableName
    cboPersonnelTable.ItemData(cboPersonnelTable.NewIndex) = rsTables!TableID
    
    If mlngPersonnelTableID = rsTables!TableID Then
      iDefaultItem = cboPersonnelTable.NewIndex
    End If
    
    rsTables.MoveNext
  Wend
  rsTables.Close
  Set rsTables = Nothing
      
  cboPersonnelTable.ListIndex = iDefaultItem

End Sub

Private Sub ReadParameters()
  
  ' Read the parameter values from the database into local variables.
  Dim sUser As String
  Dim lngColumnID As Long
  Dim lngPersModulePersonnelTableID As Long

  ' ------------------------------------------
  ' Read the Personnel Identification parameters
  ' ------------------------------------------
  mlngPersonnelTableID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE, 0)
  mlngLoginColumnID = GetModuleSetting(gsMODULEKEY_MOBILE, gsPARAMETERKEY_LOGINNAME, 0)
  mlngUniqueEmailColumnID = GetModuleSetting(gsMODULEKEY_MOBILE, gsPARAMETERKEY_UNIQUEEMAILCOLUMN, 0)
  mlngLeavingDateColumnID = GetModuleSetting(gsMODULEKEY_MOBILE, gsPARAMETERKEY_LEAVINGDATE, 0)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  If (UnloadMode <> vbFormCode) And mfChanged Then
    Select Case MsgBox("Apply changes ?", vbYesNoCancel + vbQuestion, Me.Caption)
      Case vbCancel
        Cancel = True
      Case vbYes
        cmdOK_Click
    End Select
  End If

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Save the form position to registry.
  If Me.WindowState = vbNormal Then
    SavePCSetting Me.Name, "Top", Me.Top
    SavePCSetting Me.Name, "Left", Me.Left
  End If

End Sub

Private Sub RefreshControls()
  Dim ctlCombo As ComboBox
  
  ' ------------------------------------------
  ' Refresh the Personnel Identification tab controls
  ' ------------------------------------------
  ' Refresh the Personnel Table frame controls
  cboPersonnelTable.Enabled = (cboPersonnelTable.ListCount > 1) And _
    (Not mblnReadOnly) And _
    (Not Application.PersonnelModule)
  cboPersonnelTable.BackColor = IIf(cboPersonnelTable.Enabled, vbWindowBackground, vbButtonFace)
  lblPersonnelTable.Enabled = cboPersonnelTable.Enabled

  cboLoginName.Enabled = (cboLoginName.ListCount > 1) And _
      (Not mblnReadOnly)
    cboLoginName.BackColor = IIf(cboLoginName.Enabled, vbWindowBackground, vbButtonFace)
  lblLoginNameColumn.Enabled = cboLoginName.Enabled
  
  cboEMailColumn.Enabled = (cboEMailColumn.ListCount > 1) And _
    (Not mblnReadOnly)
    cboEMailColumn.BackColor = IIf(cboEMailColumn.Enabled, vbWindowBackground, vbButtonFace)
  lblEmailAddresses.Enabled = cboEMailColumn.Enabled
  
  cboLeavingDateColumn.Enabled = (cboLeavingDateColumn.ListCount > 1) And _
    (Not mblnReadOnly)
    cboLeavingDateColumn.BackColor = IIf(cboLeavingDateColumn.Enabled, vbWindowBackground, vbButtonFace)
  lblLeavingDateColumn.Enabled = cboLeavingDateColumn.Enabled
  
  ' Disable the OK button as required.
  cmdOk.Enabled = mfChanged
  
End Sub




Private Sub RefreshPersonnelColumnControls()
  ' Refresh the Personnel column controls
  Dim iLoginColumnListIndex As Integer
  Dim iEmailColumnListIndex As Integer
  Dim iLeavingDateListIndex As Integer
  
  Dim objctl As Control

  iLoginColumnListIndex = 0
  
  UI.LockWindow Me.hWnd
  
  ' Clear the current contents of the combos.
  For Each objctl In Me
    If (TypeOf objctl Is ComboBox) And _
      (objctl.Name = "cboLoginName") Or _
      (objctl.Name = "cboEMailColumn") Or _
      (objctl.Name = "cboLeavingDateColumn") Then

      With objctl
        .Clear
        .AddItem "<None>"
        .ItemData(.NewIndex) = 0
      End With
    End If
  Next objctl

  With recColEdit
    .Index = "idxName"
    .Seek ">=", mlngPersonnelTableID

    If Not .NoMatch Then
      ' Add items to the combos for each column that has not been deleted,
      ' or is a system or link column.
      Do While Not .EOF
        If !TableID <> mlngPersonnelTableID Then
          Exit Do
        End If

        If (Not !Deleted) And _
          (!columntype <> giCOLUMNTYPE_LINK) And _
          (!columntype <> giCOLUMNTYPE_SYSTEM) Then

          If !DataType = dtVARCHAR Then
              cboLoginName.AddItem !ColumnName
              cboLoginName.ItemData(cboLoginName.NewIndex) = !ColumnID
              If !ColumnID = mlngLoginColumnID Then
                iLoginColumnListIndex = cboLoginName.NewIndex
              End If
  
              cboEMailColumn.AddItem !ColumnName
              cboEMailColumn.ItemData(cboEMailColumn.NewIndex) = !ColumnID
              If !ColumnID = mlngUniqueEmailColumnID Then
                iEmailColumnListIndex = cboEMailColumn.NewIndex
              End If
              
          ElseIf !DataType = dtTIMESTAMP Then
              cboLeavingDateColumn.AddItem !ColumnName
              cboLeavingDateColumn.ItemData(cboLeavingDateColumn.NewIndex) = !ColumnID
              If !ColumnID = mlngLeavingDateColumnID Then
                iLeavingDateListIndex = cboLeavingDateColumn.NewIndex
              End If
  
          End If
          
        End If

        .MoveNext
      Loop
    End If
  End With

  ' Select the appropriate combo items.
  cboLoginName.ListIndex = iLoginColumnListIndex
  cboEMailColumn.ListIndex = iEmailColumnListIndex
  cboLeavingDateColumn.ListIndex = iLeavingDateListIndex


  UI.UnlockWindow

End Sub




