VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPicklists 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Picklist Definition"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1051
   Icon            =   "frmPicklists.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraDefinition 
      Height          =   1950
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   8700
      Begin VB.TextBox txtDesc 
         Height          =   1080
         Left            =   1395
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   700
         Width           =   3090
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1395
         MaxLength       =   50
         TabIndex        =   2
         Top             =   240
         Width           =   3090
      End
      Begin VB.OptionButton optHidden 
         Caption         =   "&Hidden"
         Height          =   195
         Left            =   5500
         TabIndex        =   10
         Top             =   1480
         Width           =   1200
      End
      Begin VB.OptionButton optReadOnly 
         Caption         =   "&Read Only"
         Height          =   195
         Left            =   5500
         TabIndex        =   9
         Top             =   1130
         Width           =   1425
      End
      Begin VB.OptionButton optReadWrite 
         Caption         =   "Read / &Write"
         Height          =   195
         Left            =   5500
         TabIndex        =   8
         Top             =   780
         Value           =   -1  'True
         Width           =   1605
      End
      Begin VB.TextBox txtUserName 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   5500
         MaxLength       =   30
         TabIndex        =   6
         Top             =   250
         Width           =   3000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Access :"
         Height          =   195
         Index           =   3
         Left            =   4700
         TabIndex        =   7
         Top             =   760
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description :"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   3
         Top             =   760
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   1
         Top             =   310
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Owner :"
         Height          =   195
         Index           =   2
         Left            =   4700
         TabIndex        =   5
         Top             =   310
         Width           =   585
      End
   End
   Begin VB.Frame fraOKCancelButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   405
      Left            =   6300
      TabIndex        =   21
      Top             =   5000
      Width           =   2505
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   400
         Left            =   45
         TabIndex        =   19
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   400
         Left            =   1305
         TabIndex        =   20
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.Frame fraPicklistRecords 
      Height          =   2865
      Left            =   90
      TabIndex        =   11
      Top             =   2025
      Width           =   8700
      Begin VB.Frame fraAddRemoveButtons 
         BorderStyle     =   0  'None
         Caption         =   "fraButtons"
         Height          =   2400
         Left            =   7170
         TabIndex        =   13
         Top             =   250
         Width           =   1425
         Begin VB.CommandButton cmdNew 
            Caption         =   "&Add..."
            Height          =   400
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   1425
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "R&emove"
            Enabled         =   0   'False
            Height          =   400
            Left            =   0
            TabIndex        =   17
            Top             =   1500
            Width           =   1425
         End
         Begin VB.CommandButton cmdDeleteAll 
            Caption         =   "Re&move All"
            Enabled         =   0   'False
            Height          =   400
            Left            =   0
            TabIndex        =   18
            Top             =   2000
            Width           =   1425
         End
         Begin VB.CommandButton cmdAddAll 
            Caption         =   "A&dd All"
            Height          =   400
            Left            =   0
            TabIndex        =   15
            Top             =   500
            Width           =   1425
         End
         Begin VB.CommandButton cmdAddFilter 
            Caption         =   "&Filtered Add..."
            Height          =   400
            Left            =   0
            TabIndex        =   16
            Top             =   1000
            Width           =   1425
         End
      End
      Begin MSComctlLib.ListView lvRecords 
         Height          =   2445
         Left            =   225
         TabIndex        =   12
         Top             =   250
         Width           =   6850
         _ExtentX        =   12091
         _ExtentY        =   4313
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin ActiveBarLibraryCtl.ActiveBar abPicklist 
      Left            =   3405
      Top             =   5025
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "frmPicklists.frx":000C
   End
End
Attribute VB_Name = "frmPicklists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private datData As HRProDataMgr.clsDataAccess
Private fOK As Boolean
Private mblnReportedRestriction As Boolean

' Picklist definition variables.
Private mlngPicklistID As Long
Private mlngTableID As Long
Private mstrTableName As String
Private mavOrderDefinition() As Variant
Private mlngSelectedRecords As Long

' Form handling variables.
Private mfFromCopy As Boolean
Private mfReadOnly As Boolean
Private mfCancelled As Boolean
Private mfLoading As Boolean
Private mfSizing As Boolean
Private mfNoSelect As Boolean
Private mlngTimeStamp As Long

' Array holding the User Defined functions that are needed for this report
Private mastrUDFsRequired() As String


Public Property Get Changed() As Boolean
  Changed = cmdOK.Enabled
End Property

Public Property Let Changed(blnChanged As Boolean)
  cmdOK.Enabled = blnChanged
End Property


'Public Sub InitialisePickList(pfNew As Boolean, pfCopy As Boolean, Optional pvTableID As Variant, Optional plngPicklistID As Long)
Public Function InitialisePickList(pfNew As Boolean, pfCopy As Boolean, plngTableID As Long, Optional plngPicklistID As Long)
  ' Initialise the picklist definition form.
  Dim fLocked As Boolean
  Dim rsPicklist As Recordset
  Dim blnDefinitionCreator As Boolean
  Dim sAccess As String
  
  Screen.MousePointer = vbHourglass
  
  'Set mdatPick = New clsPicklists
  Set datData = New clsDataAccess
  mblnReportedRestriction = False

  ' Initialise variables.
  fOK = True
  mfCancelled = False
  mfFromCopy = pfCopy
  mfReadOnly = False
  
  ' Check if the table combo should be locked.
  'fLocked = (Not IsMissing(pvTableID))
  'If fLocked Then
  '  mlngTableID = pvTableID
  'Else
  '  mlngTableID = 0
  'End If
  
  mlngTableID = plngTableID
  mstrTableName = datGeneral.GetTableName(mlngTableID)
  
  sAccess = GetUserSetting("utils&reports", "dfltaccess picklists", ACCESS_READWRITE)
  
  If pfNew Then
    ' Creating a new picklist.
    mlngPicklistID = 0
    txtUserName.Text = gsUserName
  
    Select Case sAccess
      Case ACCESS_READWRITE
        optReadWrite.Value = True
      Case ACCESS_READONLY
        optReadOnly.Value = True
      Case Else
        optHidden.Value = True
    End Select
  Else
    mfLoading = True
    fLocked = True
    
    ' Editing an existing picklist.
    mlngPicklistID = plngPicklistID
    
    ' Read the picklist definition.
    Set rsPicklist = GetPicklist(mlngPicklistID)
    mlngTableID = rsPicklist!TableID
        
    txtDesc.Text = IIf(rsPicklist!Description <> vbNullString, rsPicklist!Description, vbNullString)
    
    If mfFromCopy Then
      txtName.Text = "Copy of " & rsPicklist!Name
      txtUserName.Text = gsUserName
      blnDefinitionCreator = True
    Else
      txtName.Text = rsPicklist!Name
      txtUserName.Text = rsPicklist!UserName
      blnDefinitionCreator = (LCase$(rsPicklist!UserName) = LCase$(gsUserName))
    
      sAccess = rsPicklist!Access
    End If
    
    mfReadOnly = Not datGeneral.SystemPermission("PICKLISTS", "EDIT")
    
    If Not blnDefinitionCreator Then
      optReadWrite.Enabled = False
      optReadOnly.Enabled = False
      optHidden.Enabled = False
    End If
    
    Select Case sAccess
      Case ACCESS_READWRITE
        optReadWrite.Value = True
      Case ACCESS_READONLY
        optReadOnly.Value = True
        mfReadOnly = (mfReadOnly Or Not blnDefinitionCreator) And (Not gfCurrentUserIsSysSecMgr)
      Case Else
        optHidden.Value = True
    End Select
        
    mlngTimeStamp = rsPicklist!intTimestamp
        
    rsPicklist.Close
    Set rsPicklist = Nothing
  
  End If
  
  ReadTableInfo
  
  mfLoading = False
  
  RefreshControls
  
  If mfFromCopy Then
    mlngPicklistID = 0
    Me.Changed = True
  Else
    Me.Changed = False
  End If

  InitialisePickList = fOK

  Screen.MousePointer = vbDefault
      
End Function


Private Sub abPicklist_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)

  cmdDelete_Click
  
End Sub

Private Sub abPicklist_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)

  Cancel = True
  
End Sub

'Private Sub cboTable_Populate()
'  ' Populate the tables combo.
'  Dim lngIndex As Long
'  Dim lngGoodIndex As Long
'  Dim objTable As CTablePrivilege
'
'  mfLoading = True
'
'  With cboTable
'    .Clear
'
'    For Each objTable In gcoTablePrivileges.Collection
'      If objTable.IsTable Then
'        .AddItem objTable.TableName
'        .ItemData(.NewIndex) = objTable.TableID
'      End If
'    Next objTable
'    Set objTable = Nothing
'
'    lngGoodIndex = 0
'    For lngIndex = 0 To (.ListCount - 1)
'      If .ItemData(lngIndex) = mlngTableID Then
'        lngGoodIndex = lngIndex
'        Exit For
'      End If
'    Next lngIndex
'
'    If .ListCount > 0 Then
'      .ListIndex = lngGoodIndex
'    Else
'      txtName.Enabled = False
'      cboTable.Enabled = False
'      cmdNew.Enabled = False
'      cmdAddAll.Enabled = False
'      cmdAddFilter.Enabled = False
'      cmdDelete.Enabled = False
'      cmdDeleteAll.Enabled = False
'      cmdOK.Enabled = False
'      optReadWrite.Enabled = False
'      optReadOnly.Enabled = False
'      optHidden.Enabled = False
'    End If
'  End With
'
'  mfLoading = False
'
'End Sub
'
'
'Private Sub cboTable_Click()
'  Dim lngIndex As Long
'
'  If Not mfLoading Then
'    If lvRecords.ListItems.Count > 0 Then
'      If COAMsgBox("Changing the table will remove ALL records from the picklist, continue ?", vbQuestion + vbOKCancel, Me.Caption) = vbOK Then
'        lvRecords.ListItems.Clear
'        RefreshControls
'      Else
'        mfLoading = True
'        For lngIndex = 0 To cboTable.ListCount - 1
'          If cboTable.ItemData(lngIndex) = mlngTableID Then
'            cboTable.ListIndex = lngIndex
'            Exit For
'          End If
'        Next lngIndex
'        mfLoading = False
'      End If
'    End If
'  End If
'
'  mlngTableID = cboTable.ItemData(cboTable.ListIndex)
'  ReadTableInfo
'  RefreshControls
'
'  Me.Changed = True
'
'End Sub

Private Sub cmdAddAll_Click()
  ' Add all records for the selected table into the picklist.
  On Error GoTo AddAllError
  
  AddItems blnCheckRecordCount:=False
  Me.Changed = True
  
  Exit Sub
  
AddAllError:
  
  COAMsgBox "Error whilst adding all records to the picklist." & vbCrLf & vbCrLf & Err.Description, vbExclamation + vbOKOnly, App.Title
    
End Sub

Private Sub cmdAddFilter_Click()
  ' Add all records for the selected table into the picklist.
  Dim fOK As Boolean
  Dim sFilteredIDs As String
  Dim objFilter As clsExprExpression
  
  ReDim mastrUDFsRequired(0)
  
  ' Add the filter 'where' clause code.
  Set objFilter = New clsExprExpression
  With objFilter
    'If .Initialise(cboTable.ItemData(cboTable.ListIndex), 0, giEXPR_RUNTIMEFILTER, giEXPRVALUE_LOGIC) Then
    If .Initialise(mlngTableID, 0, giEXPR_RUNTIMEFILTER, giEXPRVALUE_LOGIC) Then
      .SelectExpression True
  
      If (.ExpressionID > 0) Then
        
        fOK = datGeneral.FilteredIDs(.ExpressionID, sFilteredIDs)
        
        ' Generate any UDFs that are used in this filter
        If fOK Then
          datGeneral.FilterUDFs .ExpressionID, mastrUDFsRequired()
        End If
        
        If fOK Then
          
          lvRecords.Sorted = False
          
          If fOK Then fOK = UDFFunctions(mastrUDFsRequired, True)
          
          AddItems sFilteredIDs, False
          
          If fOK Then fOK = UDFFunctions(mastrUDFsRequired, False)

          lvRecords.Sorted = True
          
            Me.Changed = True

        Else
          COAMsgBox "You do not have permission to use this filter.", vbExclamation, Me.Caption
        End If
      End If
    End If
  End With
  Set objFilter = Nothing
        
End Sub


Private Sub cmdCancel_Click()
  ' Exit the form, without saving changes.
  If Me.Changed = True Then
    Select Case COAMsgBox("You have changed this picklist definition. Would you like to save changes ?", vbQuestion + vbYesNoCancel, "Picklists")
    Case vbYes
      cmdOK_Click
    Case vbNo
      Cancelled = False
    Case vbCancel
      Cancelled = True
    End Select
  End If
  
  If Cancelled = False Then
    Unload Me
  End If
  'Me.Hide

End Sub

Private Sub cmdDelete_Click()
  ' Remove the selected item from the picklist definition.
  Dim lngIndex As Long
  Dim lngNextIndex As Long
  Dim objNode As MSComctlLib.ListItem
  Dim alngIndices() As Long
  
  Screen.MousePointer = vbHourglass
  
  ' Construct an array of the item indices to be deleted.
  ReDim alngIndices(0)
  For Each objNode In lvRecords.ListItems
    If objNode.Selected Then
      lngNextIndex = UBound(alngIndices) + 1
      ReDim Preserve alngIndices(lngNextIndex)
      alngIndices(lngNextIndex) = objNode.Index
    End If
  Next objNode
  Set objNode = Nothing
  
  For lngIndex = UBound(alngIndices) To 1 Step -1
    lvRecords.ListItems.Remove alngIndices(lngIndex)
  Next lngIndex
  
  If lvRecords.ListItems.Count > 0 Then
    lvRecords.SelectedItem = lvRecords.ListItems(1)
    lvRecords.SelectedItem.EnsureVisible
  End If
  
  Screen.MousePointer = vbDefault
  
  RefreshControls
  
  Me.Changed = True

End Sub

Private Sub cmdDeleteAll_Click()
  ' Remove all items from the picklist definition.
  If COAMsgBox("Remove all records from the picklist, are you sure ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
    Screen.MousePointer = vbHourglass
    lvRecords.ListItems.Clear
    Screen.MousePointer = vbDefault
  
    Me.Changed = True
    
    RefreshControls
  End If

End Sub

Private Sub cmdNew_Click()
  ' Display the find form for selecting items to add to the picklist.
  Dim fApply As Boolean
  Dim lngIndex As Long
  Dim sFilterString As String
  Dim alngRecordIDs() As Long
  Dim frmPickFind As frmPicklistFind
  
  ' Display the form for the user to selected the required records.
  Set frmPickFind = New frmPicklistFind
  With frmPickFind
    
    If .Initialise(mlngTableID, CSVPicklistItems) = True Then
Screen.MousePointer = vbDefault
    
      .Show vbModal
        
      If Not .Cancelled Then
        ReDim alngRecordIDs(0)
        alngRecordIDs = .SelectedRecordIDs
        
        mlngSelectedRecords = UBound(alngRecordIDs)
        
        ' Add the selected records to the picklist listbox.
        If UBound(alngRecordIDs) > 0 Then
          Screen.MousePointer = vbHourglass
          lvRecords_ClearSelections
          sFilterString = ""
          
          For lngIndex = 1 To UBound(alngRecordIDs)
            sFilterString = sFilterString & _
              IIf(Len(sFilterString) > 0, ",", "") & _
              Trim(Str(alngRecordIDs(lngIndex)))
          Next lngIndex
          
          If Len(sFilterString) > 0 Then
            AddItems sFilterString, False
          End If
          
          Screen.MousePointer = vbDefault
          RefreshControls
        End If
      
        Me.Changed = True
      
      End If
    End If
  End With
        
  Unload frmPickFind
  Set frmPickFind = Nothing
  
End Sub

Private Sub cmdOK_Click()
  ' Validate and save the picklist definition, then exit.
  Dim sAccess As String
  Dim objNode As MSComctlLib.ListItem
  Dim pblnSaveAsNew As Boolean
  Dim pblnContinueSave As Boolean

  Cancelled = True

  ' Validate the picklist name.
  If Len(Trim(txtName.Text)) = 0 Then
    COAMsgBox "You must give this definition a name.", vbExclamation, "Picklists"
    txtName.SetFocus
    Exit Sub
  End If
  
  ' RH 29/08/00 - BUG 852 - Should allow save of empty picklist.
  ' RH 13/09/00 - JED request leave check in
  If lvRecords.ListItems.Count = 0 Then
    COAMsgBox "Picklists must contain at least one record.", vbExclamation, "Picklists"
    Exit Sub
  End If
  
  If optReadWrite.Value Then
    sAccess = ACCESS_READWRITE
  ElseIf optReadOnly.Value Then
    sAccess = ACCESS_READONLY
  Else
    sAccess = ACCESS_HIDDEN
  End If
  
  Screen.MousePointer = vbHourglass
  
  'Check if this definition has been changed by another user
  Call UtilityDefAmended("ASRSysPickListName", "PicklistID", mlngPicklistID, mlngTimeStamp, pblnContinueSave, pblnSaveAsNew)
  If pblnContinueSave = False Then
    Screen.MousePointer = vbDefault
    Exit Sub
  ElseIf pblnSaveAsNew Then
    txtUserName = gsUserName
    optReadWrite.Enabled = True
    optReadOnly.Enabled = True
    optHidden.Enabled = True
    mlngPicklistID = 0
  End If
  
  If mlngPicklistID > 0 Then
    
    ' Editing an existing Picklist, first delete all data related to this picklist, then Insert new.
    If CheckForExistingName(False) Then
      Exit Sub
    End If
      
    If optHidden.Value And (UCase(gsUserName) = UCase(txtUserName.Text)) Then
      If datGeneral.CheckCanMakeHidden("P", mlngPicklistID, gsUserName, "Picklist Validation") = False Then
        Exit Sub
      End If
    End If
    
    DeletePicklistItems mlngPicklistID
    UpdatePicklistName Trim(txtName.Text), txtDesc.Text, mlngTableID, sAccess, mlngPicklistID
      
    For Each objNode In lvRecords.ListItems
      InsertPicklistItem mlngPicklistID, Val(objNode.Tag)
    Next objNode
    Set objNode = Nothing
  
    Call UtilUpdateLastSaved(utlPicklist, mlngPicklistID)
  
  Else
    
    ' Add a new one.
    If CheckForExistingName(True) Then
      Exit Sub
    End If
          
    mlngPicklistID = InsertPicklistName(Trim(txtName.Text), txtDesc.Text, mlngTableID, sAccess)

    For Each objNode In lvRecords.ListItems
      InsertPicklistItem mlngPicklistID, Val(objNode.Tag)
    Next objNode
    Set objNode = Nothing
  
    Call UtilCreated(utlPicklist, mlngPicklistID)
  
  End If

  Cancelled = False
  Unload Me
  'Me.Hide

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 192 Then
    KeyCode = 0
  End If
  
End Sub

Private Sub Form_Load()
  Cancelled = False
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then
    Call cmdCancel_Click
    Cancel = Cancelled
  End If
End Sub


Public Property Get Cancelled() As Boolean
  Cancelled = mfCancelled

End Property

Public Property Let Cancelled(ByVal pfCancelled As Boolean)
  mfCancelled = pfCancelled

End Property


Private Function CheckForExistingName(pfNew As Boolean) As Boolean
  
  Dim rsName As Recordset
  Dim sSQL As String

  'MH20030516 Fault 3440
  'sSQL = "SELECT name FROM ASRSysPickListName" & _
         " WHERE name = '" & Replace(Trim(txtName.Text), "'", "''") & "'"
  sSQL = "SELECT name FROM ASRSysPickListName" & _
         " WHERE name = '" & Replace(Trim(txtName.Text), "'", "''") & "'" & _
         " AND tableID = " & CStr(mlngTableID)

  If Not pfNew Then
    sSQL = sSQL & " AND pickListID <> " & Trim(Str(mlngPicklistID))
  End If
  Set rsName = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  
  If Not (rsName.EOF And rsName.BOF) Then
    CheckForExistingName = True
        
    COAMsgBox "A picklist definition called '" & Trim(txtName.Text) & "' already exists.", vbExclamation, Me.Caption
    txtName.SetFocus
    Screen.MousePointer = vbDefault
  End If

  rsName.Close
  Set rsName = Nothing

End Function

'Private Sub Form_Resize()
'  ' Resize the form's controls as the form is itself resized.
'  Dim lCount As Long
'  Dim lWidth As Long
'  Dim iLastColumnIndex As Integer
'  Dim iMaxPosition As Integer
'
'  Const dblFORM_MINWIDTH = 6285
'  Const dblFORM_MINHEIGHT = 5400
'
'  Const dblCOORD_XGAP = 200
'  Const dblCOORD_SMALLXGAP = 150
'  Const dblCOORD_YGAP = 200
'  Const dblCOORD_SMALLYGAP = 100
'
'  If Me.WindowState = vbNormal Then
'
'    ' Ensure the form does not get narrower than the defined minimum for a Find window.
'    If Me.Width < dblFORM_MINWIDTH Then
'      Me.Width = dblFORM_MINWIDTH
'    End If
'
'    ' Ensure the form does not get wider than the screen.
'    If Me.Width > Screen.Width Then
'      Me.Width = Screen.Width
'    End If
'
'    ' Initialise the form height.
'    If Not mfSizing Then
'      mfSizing = True
'      Me.Height = Screen.Height / 3
'    End If
'
'    ' Ensure the form does not get shorter than the defined minimum for a Find window.
'    If Me.Height < dblFORM_MINHEIGHT Then
'      mfSizing = True
'      Me.Height = dblFORM_MINHEIGHT
'    End If
'
'    ' Ensure the form does not get taller than the screen.
'    If Me.Height > Screen.Height Then
'      Me.Height = Screen.Height
'    End If
'
'    ' Size the Picklisty Name/Table controls.
'    txtName.Width = Me.ScaleWidth - txtName.Left - fraAddRemoveButtons.Width - (2 * dblCOORD_XGAP)
'    cboTable.Width = txtName.Width
'
'    ' Size the listview.
'    With lvRecords
'      .Width = Me.ScaleWidth - .Left - fraAddRemoveButtons.Width - dblCOORD_XGAP - dblCOORD_SMALLXGAP
'      .Height = Me.ScaleHeight - .Top - fraAccess.Height - dblCOORD_YGAP - dblCOORD_SMALLYGAP
'    End With
'
'    ' Size the access frame.
'    With fraAccess
'      .Width = lvRecords.Width
'      .Top = lvRecords.Top + lvRecords.Height + dblCOORD_SMALLYGAP
'    End With
'
'    ' Size the frame with the Add/Remove command buttons in.
'    With fraAddRemoveButtons
'      .Left = lvRecords.Left + lvRecords.Width + dblCOORD_XGAP
'    End With
'
'    ' Size the frame with the OK/Cancel command buttons in.
'    With fraOKCancelButtons
'      .Top = Me.ScaleHeight - .Height - dblCOORD_YGAP
'      .Left = fraAddRemoveButtons.Left
'    End With
'
'    ' Stretch the last find column to fit the listview.
'    iLastColumnIndex = -1
'    iMaxPosition = -1
'    With lvRecords
'      If .ColumnHeaders(.ColumnHeaders.Count).Left + .ColumnHeaders(.ColumnHeaders.Count).Width < .Width Then
'        If .Width - .ColumnHeaders(.ColumnHeaders.Count).Left - _
'          ((UI.GetSystemMetrics(SM_CXFRAME) * 6) * Screen.TwipsPerPixelX) > 0 Then
'          .ColumnHeaders(.ColumnHeaders.Count).Width = .Width - .ColumnHeaders(.ColumnHeaders.Count).Left - _
'            ((UI.GetSystemMetrics(SM_CXFRAME) * 6) * Screen.TwipsPerPixelX)
'        End If
'      End If
'    End With
'  End If
'
'End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub Form_Unload(Cancel As Integer)
  lvRecords.ListItems.Clear
  Set datData = Nothing
End Sub




Private Sub lvRecords_ItemClick(ByVal Item As MSComctlLib.ListItem)
  
  RefreshControls
  
End Sub

Private Sub lvRecords_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

  If Button = vbRightButton Then
  
    With Me.abPicklist.Bands("bndPicklist")
    
      ' Enable/disable the required tools.
      .Tools("Remove").Enabled = Me.cmdDelete.Enabled
    
    End With

    abPicklist.Bands("bndPicklist").TrackPopup -1, -1
    
  End If

End Sub

Private Sub optHidden_Click()

  Me.Changed = True

End Sub

Private Sub optReadOnly_Click()

  Me.Changed = True

End Sub

Private Sub optReadWrite_Click()

  Me.Changed = True

End Sub

Private Sub txtDesc_Change()

  Me.Changed = True

End Sub

Private Sub txtDesc_GotFocus()
  UI.txtSelText
  cmdOK.Default = False
End Sub

Private Sub txtDesc_LostFocus()
  cmdOK.Default = True
End Sub

Private Sub txtName_Change()

  Me.Changed = True
  
End Sub

Private Sub txtName_GotFocus()
  UI.txtSelText
  
End Sub

Private Sub RefreshControls()
  ' Enable/Disable controls as required.
  
  If mfReadOnly Then
    ControlsDisableAll Me
    txtDesc.Enabled = True
    txtDesc.Locked = True
    txtDesc.BackColor = vbButtonFace
    txtDesc.ForeColor = vbGrayText
  Else
    cmdDelete.Enabled = (lvRecords_SelectedItemsCount > 0) And (Not mfNoSelect)
    cmdDeleteAll.Enabled = (lvRecords.ListItems.Count > 0) And (Not mfNoSelect)
  End If
  
  Me.Caption = "Picklist Definition (" & lvRecords.ListItems.Count & " record" & IIf(lvRecords.ListItems.Count <> 1, "s", "") & ")"
  
End Sub

Private Sub ReadTableInfo()
  ' Read the required information about the selected table.
'  Dim fNoSelect As Boolean
  Dim iNextIndex As Integer
'  Dim objColumnPrivileges As CColumnPrivileges
  Dim rsInfo As Recordset
  Dim objTable As CTablePrivilege
  
'  fNoSelect = False
  
  ' Initialise the order definition array.
  ' Index 1 = column name.
  ' Index 2 = table name.
  ' Index 3 = table ID.
  ' Index 4 = column size.
  ' Index 5 = decimals.
  ' Index 6 = uses separator.
  ReDim mavOrderDefinition(6, 0)
  
  ' Clear the listview headers.
  lvRecords.ColumnHeaders.Clear
  
  ' Get the default order items from the database.
  Set objTable = gcoTablePrivileges.FindTableID(mlngTableID)
  Set rsInfo = datGeneral.GetOrderDefinition(objTable.DefaultOrderID)
  Set objTable = Nothing
  
  If rsInfo.EOF And rsInfo.BOF Then
    COAMsgBox "No default order defined for this table." & vbCrLf & _
           "Unable to display the picklist records.", vbExclamation, "Security"
'    mfNoSelect = True
  Else
    ' Check the user's privilieges on the order columns.
    Do While Not rsInfo.EOF
      If rsInfo!Type = "F" Then
'        ' Get the column privileges collection for the given table.
'        Set objColumnPrivileges = GetColumnPrivileges(rsInfo!TableName)
'
'        If objColumnPrivileges.Item(rsInfo!ColumnName).AllowSelect Then
          ' Add the column name to the listview headers.
          lvRecords.ColumnHeaders.Add , , RemoveUnderScores(rsInfo!ColumnName)
          
          ' Add the column to the order definition array.
          iNextIndex = UBound(mavOrderDefinition, 2) + 1
          ReDim Preserve mavOrderDefinition(6, iNextIndex)
          mavOrderDefinition(1, iNextIndex) = Trim(rsInfo!ColumnName)
          mavOrderDefinition(2, iNextIndex) = Trim(rsInfo!TableName)
          mavOrderDefinition(3, iNextIndex) = rsInfo!TableID
          mavOrderDefinition(4, iNextIndex) = datGeneral.GetDataSize(rsInfo!ColumnID)
          mavOrderDefinition(5, iNextIndex) = datGeneral.GetDecimalsSize(rsInfo!ColumnID)
          mavOrderDefinition(6, iNextIndex) = datGeneral.DoesColumnUseSeparators(rsInfo!ColumnID)

'        Else
'          fNoSelect = True
'        End If
'
'        Set objColumnPrivileges = Nothing
      End If
      
      rsInfo.MoveNext
    Loop

'    ' Inform the user if they do not have permission to see the picklist data.
'    If fNoSelect Then
'      If UBound(mavOrderDefinition, 2) > 0 Then
'        COAMsgBox "You do not have 'read' permission on all of the columns in the selected table's default order." & _
'          vbCrLf & "Only permitted columns will be shown.", vbExclamation, "Security"
'        fNoSelect = False
'      Else
'        COAMsgBox "You do not have 'read' permission on any of the columns in the selected table's default order." & _
'          vbCrLf & "Unable to display the picklist records.", vbExclamation, "Security"
'      End If
'    End If
    
'    mfNoSelect = fNoSelect
  End If
  
  rsInfo.Close
  Set rsInfo = Nothing
  
  ReadPicklist

End Sub
Private Sub ReadPicklist()
  ' Read the picklist records.
  Dim sPicklistFilter As String

  Dim objTableView As Object

  'MH20000714 Only retreive picklist IDs which still exist in the base table

'  ' Populate the listview with the picklist records.
'  sPicklistFilter = "SELECT recordID" & _
'    " FROM ASRSysPickListItems" & _
'    " WHERE pickListID = " & Trim(Str(mlngPickListID))
  
  ' Populate the listview with the picklist records.
  'sPicklistFilter = "SELECT recordID"
  sPicklistFilter = "SELECT DISTINCT recordID" & _
    " FROM ASRSysPickListItems" & _
    " WHERE pickListID = " & Trim(Str(mlngPicklistID)) & _
    " AND RecordID IN" & _
    "(SELECT ID FROM " & gcoTablePrivileges.Item(mstrTableName).RealSource & ")"
  
  AddItems sPicklistFilter, True
  
  If lvRecords.ListItems.Count > 0 Then
    lvRecords_ClearSelections
    lvRecords.SelectedItem = lvRecords.ListItems(1)
    lvRecords.SelectedItem.EnsureVisible
  End If

End Sub

Private Function lvRecords_SelectedItemsCount() As Long
  ' Return the count of selected items in the listview.
  Dim lngCount As Long
  Dim objNode As MSComctlLib.ListItem
  
  lngCount = 0
  For Each objNode In lvRecords.ListItems
    If objNode.Selected Then lngCount = lngCount + 1
  Next objNode
  Set objNode = Nothing
  
  lvRecords_SelectedItemsCount = lngCount
  
End Function
Private Sub lvRecords_ClearSelections()
  ' Deselect any currently selected items in the listview.
  Dim objNode As MSComctlLib.ListItem
  
  For Each objNode In lvRecords.ListItems
    objNode.Selected = False
  Next objNode
  Set objNode = Nothing

End Sub


Private Function ItemInList(plngID As Long) As Boolean
  ' Return TRUE if the given ID is already in the picklist.
  Dim fInList As Boolean
  Dim objItem As MSComctlLib.ListItem
  
  fInList = False

  For Each objItem In lvRecords.ListItems
    If Trim(objItem.Tag) = Trim(Str(plngID)) Then
      fInList = True
      Exit For
    End If
  Next objItem
  Set objItem = Nothing

  ItemInList = fInList
  
End Function





Private Sub AddItems(Optional psFilterString As String, Optional blnCheckRecordCount As Boolean)
  Dim fApply As Boolean
  Dim fColumnOK As Boolean
  Dim fFound As Boolean
  Dim fNoSelect As Boolean
  Dim fSomeSelect As Boolean
  Dim fColumnDenied As Boolean
  Dim fRecordDenied As Boolean
  Dim fNoOrder As Boolean
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iNextIndex As Integer
  Dim lngCount As Long
  Dim sSQL As String
  Dim sSource As String
  Dim sRealSource As String
  Dim sColumnCode As String
  Dim sColumnList As String
  Dim sJoinCode As String
  Dim sWhereCode As String
  Dim sItemDesc As String
  Dim sRecordDesc As String
  Dim sRecord As String
  Dim objColumnPrivileges As CColumnPrivileges
  Dim rsItem As Recordset
  Dim objNode As MSComctlLib.ListItem
  Dim objTableView As CTablePrivilege
  Dim avTableViews() As Variant
  Dim asViews() As String
  Dim lngExpectNumRecords As Long
  Dim lngActualNumRecords As Long
  Dim lCountAddedAll As Long
  Dim rsFilterCount As Recordset
  Dim fUserCancelled As Boolean
  Dim fAlreadyCleared As Boolean
  Dim strFormat As String
  
  On Error GoTo AddItems_ERROR
  
  Screen.MousePointer = vbHourglass

  lvRecords_ClearSelections
  
  fColumnDenied = False
  fRecordDenied = False
  fNoSelect = False
  fSomeSelect = False
  sJoinCode = ""
  sColumnList = ""
  
  ' Dimension an array of tables/views joined to the base table/view.
  ' Column 1 = 0 if this row is for a table, 1 if it is for a view.
  ' Column 2 = table/view ID.
  ' Column 3 = view name.
  ReDim avTableViews(3, 0)
    
  fNoOrder = (UBound(mavOrderDefinition, 2) = 0)
  
  If Not fNoOrder Then
    
    For iLoop = 1 To UBound(mavOrderDefinition, 2)
      sSource = mavOrderDefinition(2, iLoop)
      
      Set objColumnPrivileges = GetColumnPrivileges(sSource)
      sRealSource = gcoTablePrivileges.Item(sSource).RealSource
    
      fColumnOK = objColumnPrivileges.IsValid(mavOrderDefinition(1, iLoop))
      
      If fColumnOK Then
        fColumnOK = objColumnPrivileges.Item(mavOrderDefinition(1, iLoop)).AllowSelect
      End If
      Set objColumnPrivileges = Nothing
    
      If fColumnOK Then
        ' The column can be read from the base table/view, or directly from a parent table.
        ' Add the column to the column list.
        sColumnList = sColumnList & _
          IIf(Len(sColumnList) > 0, ", ", "") & _
          sRealSource & "." & Trim(mavOrderDefinition(1, iLoop))
        fSomeSelect = True
  
        ' If the column comes from a parent table, then add the table to the Join code.
        If mavOrderDefinition(3, iLoop) <> mlngTableID Then
          ' Check if the table has already been added to the join code.
          fFound = False
          For iNextIndex = 1 To UBound(avTableViews, 2)
            If avTableViews(1, iNextIndex) = 0 And _
              avTableViews(2, iNextIndex) = mavOrderDefinition(3, iLoop) Then
              fFound = True
              Exit For
            End If
          Next iNextIndex
          
          If Not fFound Then
            ' The table has not yet been added to the join code, so add it to the array and the join code.
            iNextIndex = UBound(avTableViews, 2) + 1
            ReDim Preserve avTableViews(3, iNextIndex)
            avTableViews(1, iNextIndex) = 0
            avTableViews(2, iNextIndex) = mavOrderDefinition(3, iLoop)
            
            sJoinCode = sJoinCode & _
              " LEFT OUTER JOIN " & sRealSource & _
              " ON " & gcoTablePrivileges.Item(mstrTableName).RealSource & ".ID_" & Trim(Str(mavOrderDefinition(3, iLoop))) & _
              " = " & sRealSource & ".ID"
              '" ON " & gcoTablePrivileges.Item(cboTable.List(cboTable.ListIndex)).RealSource & ".ID_" & Trim(Str(mavOrderDefinition(3, iLoop))) &
          End If
        End If
      Else
        ' The column cannot be read from the base table/view, or directly from a parent table.
        ' Loop through the views on the column's table, seeing if any have 'read' permission granted on them.
        ReDim asViews(0)
        For Each objTableView In gcoTablePrivileges.Collection
          If (Not objTableView.IsTable) And _
            (objTableView.TableID = mavOrderDefinition(3, iLoop)) And _
            (objTableView.AllowSelect) Then
            
            sSource = objTableView.ViewName
            sRealSource = gcoTablePrivileges.Item(sSource).RealSource

            ' Get the column permission for the view.
            Set objColumnPrivileges = GetColumnPrivileges(sSource)

            If objColumnPrivileges.IsValid(mavOrderDefinition(1, iLoop)) Then
              If objColumnPrivileges.Item(mavOrderDefinition(1, iLoop)).AllowSelect Then
                ' Add the view info to an array to be put into the column list or order code below.
                iNextIndex = UBound(asViews) + 1
                ReDim Preserve asViews(iNextIndex)
                asViews(iNextIndex) = objTableView.ViewName
                
                ' Add the view to the Join code.
                ' Check if the view has already been added to the join code.
                fFound = False
                For iNextIndex = 1 To UBound(avTableViews, 2)
                  If avTableViews(1, iNextIndex) = 1 And _
                    avTableViews(2, iNextIndex) = objTableView.ViewID Then
                    fFound = True
                    Exit For
                  End If
                Next iNextIndex
        
                If Not fFound Then
                  ' The view has not yet been added to the join code, so add it to the array and the join code.
                  iNextIndex = UBound(avTableViews, 2) + 1
                  ReDim Preserve avTableViews(3, iNextIndex)
                  avTableViews(1, iNextIndex) = 1
                  avTableViews(2, iNextIndex) = objTableView.ViewID
                  avTableViews(3, iNextIndex) = objTableView.ViewName
        
                  If mavOrderDefinition(3, iLoop) = mlngTableID Then
                    sJoinCode = sJoinCode & _
                      " LEFT OUTER JOIN " & sRealSource & _
                      " ON " & gcoTablePrivileges.Item(mstrTableName).RealSource & ".ID" & _
                      " = " & sRealSource & ".ID"
                      '" ON " & gcoTablePrivileges.Item(cboTable.List(cboTable.ListIndex)).RealSource & ".ID" &
                  Else
                    sJoinCode = sJoinCode & _
                      " LEFT OUTER JOIN " & sRealSource & _
                      " ON " & gcoTablePrivileges.Item(mstrTableName).RealSource & ".ID_" & Trim(Str(objTableView.TableID)) & _
                      " = " & sRealSource & ".ID"
                      '" ON " & gcoTablePrivileges.Item(cboTable.List(cboTable.ListIndex)).RealSource & ".ID_" & Trim(Str(objTableView.TableID)) &
                  End If
                End If
              End If
            End If
            Set objColumnPrivileges = Nothing

          End If
        Next objTableView
        Set objTableView = Nothing
  
        ' The current user does have permission to 'read' the column through a/some view(s) on the
        ' table.
        If UBound(asViews) = 0 Then
          fNoSelect = True
          fColumnDenied = True
          'JPD 20030805 Fault 6568
          'sColumnList = sColumnList & _
            IIf(Len(sColumnList) > 0, ", ", "") & _
            "NULL"
          For iLoop2 = lvRecords.ColumnHeaders.Count To 1 Step -1
            If UCase(lvRecords.ColumnHeaders(iLoop2).Text) = UCase(RemoveUnderScores(Trim(mavOrderDefinition(1, iLoop)))) Then
              lvRecords.ColumnHeaders.Remove iLoop2
            End If
          Next

        Else
          ' Add the column to the column list.
          sColumnCode = ""
          For iNextIndex = 1 To UBound(asViews)
            If iNextIndex = 1 Then
              sColumnCode = "CASE "
            End If
              
            sColumnCode = sColumnCode & _
              " WHEN NOT " & asViews(iNextIndex) & "." & mavOrderDefinition(1, iLoop) & " IS NULL THEN " & asViews(iNextIndex) & "." & mavOrderDefinition(1, iLoop)
          Next iNextIndex
            
          If Len(sColumnCode) > 0 Then
            sColumnCode = sColumnCode & _
              " ELSE NULL" & _
              " END AS " & _
              mavOrderDefinition(1, iLoop)
              
            sColumnList = sColumnList & _
              IIf(Len(sColumnList) > 0, ", ", "") & _
              sColumnCode
          End If
        
          fSomeSelect = True
        End If
      End If
    Next iLoop
    
    ' Create the string for creating the items that will appear in the listbox.
    If fSomeSelect Then
      
      sWhereCode = vbNullString
      For iNextIndex = 1 To UBound(avTableViews, 2)
        If avTableViews(1, iNextIndex) = 1 Then

          sWhereCode = sWhereCode & _
            IIf(iNextIndex > 1, " OR ", vbNullString) & _
            gcoTablePrivileges.Item(mstrTableName).RealSource & ".ID IN (SELECT ID FROM " & avTableViews(3, iNextIndex) & ")"
        
        End If
      Next iNextIndex
      
      If sWhereCode <> vbNullString Then
        sWhereCode = " WHERE (" & sWhereCode & ")"
      End If
      
'' Generate any UDFs that are used in this filter
'If blnOK Then
'  datGeneral.FilterUDFs lngTempFilterID, mastrUDFsRequired()
'End If
            
      ' Add the filter code if required.
      lngExpectNumRecords = -1
      If Not IsMissing(psFilterString) Then
        If Len(psFilterString) > 0 Then
          sWhereCode = sWhereCode & IIf(sWhereCode = vbNullString, " WHERE ", " AND ") & _
              gcoTablePrivileges.Item(mstrTableName).RealSource & ".id IN (" & psFilterString & ")"
        End If
        
        If blnCheckRecordCount Then
          Set rsItem = datGeneral.GetRecords(psFilterString)
          lngExpectNumRecords = rsItem.RecordCount
        End If
      
      End If
      
      
      sSQL = "SELECT " & sColumnList & ", " & gcoTablePrivileges.Item(mstrTableName).RealSource & ".id" & _
        " FROM " & gcoTablePrivileges.Item(mstrTableName).RealSource & _
        " " & sJoinCode & sWhereCode
      
      ' Get the required recordset.
      Set rsItem = datGeneral.GetRecords(sSQL)

      If lngExpectNumRecords <> -1 Then
        'This will check whether or not the user has access
        'to all of the records included in this picklist.
        lngActualNumRecords = rsItem.RecordCount
        fRecordDenied = (lngActualNumRecords < lngExpectNumRecords)
      End If

      If Not mfLoading Then
        With gobjProgress
          '.AviFile = ""
          .AVI = dbPicklist
          .MainCaption = "Picklist"
          .Caption = "Adding Records To Picklist..."
          .NumberOfBars = 1
          .Bar1Value = 0
          .Bar1MaxValue = rsItem.RecordCount
          .Cancel = True
          .Time = False
          .OpenProgress
        End With
      End If
      
      With rsItem
        Do While Not .EOF
          sRecord = ""
          If IsNull(rsItem(0)) Then
            sRecord = ""
          Else
            If rsItem.Fields(0).Type = adDBTimeStamp Then
              sRecord = Format(rsItem(0), DateFormat)
            ElseIf rsItem.Fields(0).Type = adNumeric Then
              ' Are thousand separators used
              strFormat = "0"
              If mavOrderDefinition(6, 1) Then
                strFormat = "#,0"
              End If
              If mavOrderDefinition(5, 1) > 0 Then
                strFormat = strFormat & "." & String(mavOrderDefinition(5, 1), "0")
              End If
              
              sRecord = Format(rsItem(0), strFormat)
  
            Else
              sRecord = rsItem(0)
            End If
          End If

          ' RH 16/11/00 - Dont need to call ItemInList when we are adding all...just
          '               clear the listview, then add them all without checking
          If blnCheckRecordCount = False And psFilterString = "" Then
            If fAlreadyCleared = False Then
              ' Only clear the listview once
              Me.lvRecords.ListItems.Clear
              fAlreadyCleared = True
            End If
            fApply = True
          Else
            ' RH 25/01/01 - BUG 1501
            If mfLoading = True Then
              fApply = True
            Else
              fApply = Not ItemInList(rsItem(rsItem.Fields.Count - 1))
            End If
          End If
          
          ' Check if the current item is already in the picklist.
'          fApply = Not ItemInList(rsItem(rsItem.Fields.Count - 1))
          If Not fApply Then
            sRecordDesc = sRecord
            For lngCount = 1 To (rsItem.Fields.Count - 2)
              sRecordDesc = sRecordDesc & ", "
              If Not IsNull(rsItem(lngCount)) Then
                sItemDesc = rsItem(lngCount)
                sRecordDesc = sRecordDesc & Trim(sItemDesc)
              End If
            Next lngCount
    
            ' RH 06/09/00 - Do not add duplicate entries in the picklist
            'fApply = (COAMsgBox("The record '" & sRecordDesc & "' is already in this picklist, add again ?", vbQuestion + vbYesNo, Me.Caption) = vbYes)
            fApply = False
          
          End If

          If fApply Then
            
            Set objNode = lvRecords.ListItems.Add(, , sRecord)
            objNode.Selected = True

            For lngCount = 1 To (rsItem.Fields.Count - 1)
              If IsNull(rsItem(lngCount)) Then
                sRecord = ""
              Else
                If rsItem.Fields(lngCount).Type = adDBTimeStamp Then
                  sRecord = Format(rsItem(lngCount), DateFormat)
                ElseIf rsItem.Fields(lngCount).Type = adNumeric Then
                  ' Are thousand separators used
                  strFormat = "0"
                  If mavOrderDefinition(6, lngCount + 1) Then
                    strFormat = "#,0"
                  End If
                  If mavOrderDefinition(5, lngCount + 1) > 0 Then
                    strFormat = strFormat & "." & String(mavOrderDefinition(5, lngCount + 1), "0")
                  End If
                  
                  sRecord = Format(rsItem(lngCount), strFormat)
                Else
                  sRecord = rsItem(lngCount)
                End If
              End If
    
              If lngCount < (rsItem.Fields.Count - 1) Then
                objNode.SubItems(lngCount) = sRecord
              Else
                objNode.Tag = sRecord
              End If
            Next lngCount
            lCountAddedAll = lCountAddedAll + 1
          End If
    
          If Not mfLoading Then
            gobjProgress.UpdateProgress
            gobjProgress.Bar1Caption = "Added " & gobjProgress.Bar1Value & " Records"
            If gobjProgress.Cancelled Then
              fUserCancelled = True
              Exit Do
            End If
          End If
          
          .MoveNext
        Loop

        If gobjProgress.Visible = True Then gobjProgress.CloseProgress
        
        .Close
      End With
      Set rsItem = Nothing
    End If
  End If

  ' blncheckrecordcount is true if we are reading the definition only
  If blnCheckRecordCount = False Then
    If InStr(psFilterString, "SELECT") > 0 Then
      ' We are adding records from a filter
      Set rsFilterCount = datGeneral.GetRecords(psFilterString)
      If (lCountAddedAll <> rsFilterCount.RecordCount) And Not fUserCancelled Then
        COAMsgBox "One or more records returned by this filter have already been added" & vbCrLf & _
               "to the picklist.", vbInformation + vbOKOnly, "Picklists"
      End If
      Set rsFilterCount = Nothing
    Else
      If psFilterString = "" Then
        ' We are adding all records
        ' <is it worth telling them if some records from the table have already
        '  been added to the picklist ? i think not>
      Else
        ' We are adding just selected records
        If lCountAddedAll <> (UBound(Split(psFilterString, ",")) + 1) Then
          If (UBound(Split(psFilterString, ",")) + 1) = 1 Then
            COAMsgBox "The selected record has not been added to this picklist as it" & vbCrLf & _
                   "has been deleted by another user.", vbInformation + vbOKOnly, "Picklists"
          Else
            ' RH 21/11/00 - Bug 1404
            If mlngSelectedRecords <> (UBound(Split(psFilterString, ",")) + 1) Then
              COAMsgBox "One or more records have not been added to this picklist as they" & vbCrLf & _
                     "have been deleted by another user.", vbInformation + vbOKOnly, "Picklists"
            End If
          End If
        End If
      End If
    End If
  End If

  RefreshControls

  Screen.MousePointer = vbDefault

  If mblnReportedRestriction = False Then
    mblnReportedRestriction = True
  
    If fNoOrder Then
      COAMsgBox "No default order defined for this table." & vbCrLf & _
             "Unable to add records.", vbExclamation, "Security"
  
    ElseIf fColumnDenied Then
    
      ' Inform the user if they do not have permission to see the data.
      If fSomeSelect Then
        COAMsgBox "You do not have 'read' permission on all of the columns in the selected order." & vbCrLf & _
               "Only permitted columns will be shown.", vbExclamation, "Security"
      Else
        COAMsgBox "You do not have permission to read any of the columns in the default order for this table." & vbCrLf & _
               "Unable to display records.", vbExclamation, "Security"
        fOK = False
      End If
  
    End If


    If fOK Then
      If fRecordDenied Then

        If lngActualNumRecords > 0 Then
          COAMsgBox "You do not have 'read' permission on all of the records in the selected picklist." & vbCrLf & _
                 "Only permitted records will be shown" & _
                 IIf(Not mfFromCopy, " and the definition will be read only", vbNullString) & _
                 ".", vbExclamation, "Security"
          mfReadOnly = (Not mfFromCopy)
        Else
          COAMsgBox "You do not have 'read' permission on any of the records in the selected picklist." & vbCrLf & _
                 "Unable to display records.", vbExclamation, "Security"
          fOK = False
        End If
      
      End If
    End If
  
  End If

Exit Sub

AddItems_ERROR:

  If gobjProgress.Visible = True Then gobjProgress.CloseProgress
  COAMsgBox "Error whilst adding records to picklist : " & vbCrLf & vbCrLf & Err.Number & " - " & Err.Description, vbExclamation + vbOKOnly, App.Title
  Screen.MousePointer = vbDefault
  
End Sub

Public Property Get SelectedID() As Long
  SelectedID = mlngPicklistID
End Property

Public Property Let SelectedID(ByVal lngNewValue As Long)
  mlngPicklistID = lngNewValue
End Property


Public Sub PrintDef(plngTableID As Long, plngPicklistID As Long)

  Dim objPrintDef As clsPrintDef
  Dim rsTemp As Recordset
  Dim objNode As MSComctlLib.ListItem
  Dim lngCol As Long
  
  Dim strColumnHeader As String
  Dim strColumnData As String
  
  
  mlngTableID = plngTableID
  mstrTableName = datGeneral.GetTableName(mlngTableID)
  mlngPicklistID = plngPicklistID
  
  Set datData = New clsDataAccess
  Set rsTemp = GetPicklist(mlngPicklistID)
  
  Set objPrintDef = New HRProDataMgr.clsPrintDef

  If objPrintDef.IsOK Then
  
    With objPrintDef
      If .PrintStart(False) Then
        .PrintHeader "Picklist : " & rsTemp!Name
    
        .PrintNormal "Description : " & rsTemp!Description
        .PrintNormal
    
        .PrintNormal "Owner : " & rsTemp!UserName
        .PrintNormal "Access : " & AccessDescription(rsTemp!Access)
        'Select Case rsTemp!Access
        'Case "RW": .PrintNormal "Access : Read / Write"
        'Case "RO": .PrintNormal "Access : Read only"
        'Case "HD": .PrintNormal "Access : Hidden"
        'End Select
        .PrintNormal
        
        .PrintNormal "Base Table : " & mstrTableName
        
        'Select Case Val(rsTemp!Selection)
        'Case 0: .PrintNormal "Records : All"
        'Case 1: .PrintNormal "Records : Picklist '" & rsTemp!PickListName & "'"
        'Case 2: .PrintNormal "Records : Filter '" & rsTemp!FilterName & "'"
        'End Select
        
        .PrintNormal
        
        '--------
        
        .PrintTitle "Records"
      
        Call ReadTableInfo
    
        If lvRecords.ListItems.Count > 0 Then
        
          For Each objNode In lvRecords.ListItems
    
            For lngCol = 1 To lvRecords.ColumnHeaders.Count
            
              If lngCol = 1 Then
                strColumnData = objNode.Text
              Else
                strColumnData = objNode.SubItems(lngCol - 1)
              End If
              
              .PrintNormal lvRecords.ColumnHeaders(lngCol).Text & " : " & strColumnData
    
            Next
            .PrintNormal
        
          Next
    
        End If
    
        '--------
        
        .PrintEnd
        .PrintConfirm "Picklist : " & rsTemp!Name, "Picklist Definition"
      End If
    End With
  
  End If
  
  rsTemp.Close
  
  Set rsTemp = Nothing
  'Set mdatPick = Nothing

Exit Sub

LocalErr:
  COAMsgBox "Printing Picklist Definition Failed" & vbCrLf & Err.Description, vbCritical, "Picklist"

End Sub


Private Function GetPicklist(plngPicklistID As Long) As Recordset
  ' Get the picklist definition from the database.
  Dim sSQL As String
    
'  sSQL = "SELECT *" & _
'    " FROM ASRSysPickListName" & _
'    " WHERE pickListID = " & Trim(Str(plngPicklistID))
  
  sSQL = "SELECT *, CONVERT(integer, ASRSysPickListName.Timestamp) as intTimeStamp" & _
    " FROM ASRSysPickListName" & _
    " WHERE pickListID = " & Trim(Str(plngPicklistID))
    
  Set GetPicklist = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

End Function

Private Sub DeletePicklistItems(plngPicklistID As Long)
  ' Delete the given picklist definition from the database.
  Dim sSQL As String
    
  sSQL = "DELETE FROM ASRSysPickListItems" & _
    " WHERE pickListID = " & Trim(Str(plngPicklistID))
  datData.ExecuteSql sSQL

End Sub

Private Sub UpdatePicklistName(psName As String, psDesc As String, plngTableID As Long, psAccess As String, plngPicklistID As Long)
  ' Update the given picklist definition in the database.
  Dim sSQL As String
    
  sSQL = "UPDATE ASRSysPickListName SET" & _
    " name = '" & Replace(psName, "'", "''") & "'," & _
    " description = '" & Replace(psDesc, "'", "''") & "'," & _
    " tableID = " & Trim(Str(plngTableID)) & "," & _
    " access = '" & psAccess & "'" & _
    " WHERE pickListID = " & Trim(Str(plngPicklistID))
  datData.ExecuteSql sSQL

End Sub

Private Sub InsertPicklistItem(plngPicklistID As Long, plngRecordID As Long)
  ' Insert the given picklist item into the database.
  Dim sSQL As String
    
  sSQL = "INSERT INTO ASRSysPickListItems" & _
    " (pickListID, recordID)" & _
    " VALUES(" & Trim(Str(plngPicklistID)) & ", " & Trim(Str(plngRecordID)) & ")"
  datData.ExecuteSql sSQL

End Sub

Private Function InsertPicklistName(psName As String, psDesc As String, plngTableID As Long, psAccess As String) As Long
  ' Insert the given picklist definition into the database.
  ' Return the new record's ID.
  Dim sSQL As String
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim fSavedOK As Boolean

  On Error GoTo InsertPicklist_ERROR
  
'  Dim rsPicklist As Recordset
    
  sSQL = "INSERT ASRSysPickListName" & _
    " (name, description, tableID, access, userName)" & _
    " VALUES(" & _
    "'" & Replace(psName, "'", "''") & "', " & _
    "'" & Replace(psDesc, "'", "''") & "', " & _
    plngTableID & ", " & _
    "'" & psAccess & "', " & _
    "'" & datGeneral.UserNameForSQL & "')"
  
'  datData.ExecuteSql sSQL
    
'  sSQL = "SELECT MAX(pickListID)" & _
'    " FROM ASRSysPickListName"
'  Set rsPicklist = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
'  InsertPicklistName = rsPicklist(0)
    
'  rsPicklist.Close
'  Set rsPicklist = Nothing

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
    pmADO.Value = sSQL
              
    Set pmADO = .CreateParameter("tablename", adVarChar, adParamInput, 255)
    .Parameters.Append pmADO
    pmADO.Value = "AsrSysPickListName"
              
    Set pmADO = .CreateParameter("idcolumnname", adVarChar, adParamInput, 30)
    .Parameters.Append pmADO
    pmADO.Value = "PickListID"
              
    Set pmADO = Nothing
            
    cmADO.Execute
              
    If Not fSavedOK Then
      COAMsgBox "The new record could not be created." & vbCrLf & vbCrLf & _
        Err.Description, vbOKOnly + vbExclamation, App.ProductName
        InsertPicklistName = 0
        Set cmADO = Nothing
        Exit Function
    End If
    
    InsertPicklistName = IIf(IsNull(.Parameters(0).Value), 0, .Parameters(0).Value)
          
  End With
  
  Set cmADO = Nothing
  Exit Function
  
InsertPicklist_ERROR:
  
  fSavedOK = False
  Resume Next

End Function


Private Function CSVPicklistItems() As String
  
  ' RH 06/09/00 - Return list of IDs in the current picklist
  Dim objItem As MSComctlLib.ListItem
  Dim sList As String
  
  For Each objItem In lvRecords.ListItems
    sList = sList & IIf(Len(sList) > 0, ", ", "") & objItem.Tag
  Next objItem
  
  CSVPicklistItems = IIf(Len(sList) > 0, sList, vbNullString)
  Set objItem = Nothing
  
End Function


Private Function GetExpressionCount(plngID As Long) As Long

  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim fOK As Boolean
  
  Dim strFilterCode As String
  Dim lngRecs As Long
  Dim strMBText As String
  Dim objExpression As clsExprExpression
  
  Set objExpression = New clsExprExpression
  objExpression.ExpressionID = plngID
  
    fOK = objExpression.RuntimeFilterCode(strFilterCode, True, True)
    If fOK Then
  
      strSQL = "SELECT COUNT(*) FROM " & _
               gcoTablePrivileges.Item(mstrTableName).RealSource & _
               " WHERE ID IN (" & strFilterCode & ")"
      Set rsTemp = datGeneral.GetRecords(strSQL)
      
      lngRecs = rsTemp(0).Value
      
      rsTemp.Close
      Set rsTemp = Nothing
      Set objExpression = Nothing
      
      GetExpressionCount = lngRecs
    
    End If
    
End Function


