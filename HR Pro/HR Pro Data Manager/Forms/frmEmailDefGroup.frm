VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmEmailDefGroup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Email Group Definition"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1073
   Icon            =   "frmEmailDefGroup.frx":0000
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDefinition 
      Height          =   1950
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9180
      Begin VB.TextBox txtDesc 
         Height          =   1080
         Left            =   1620
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   705
         Width           =   3000
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1620
         MaxLength       =   50
         TabIndex        =   2
         Top             =   315
         Width           =   3000
      End
      Begin VB.OptionButton optReadOnly 
         Caption         =   "&Read Only"
         Height          =   195
         Left            =   6000
         TabIndex        =   9
         Top             =   1200
         Width           =   1695
      End
      Begin VB.OptionButton optReadWrite 
         Caption         =   "Read / &Write"
         Height          =   195
         Left            =   6000
         TabIndex        =   8
         Top             =   810
         Value           =   -1  'True
         Width           =   1875
      End
      Begin VB.TextBox txtUserName 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   6000
         MaxLength       =   30
         TabIndex        =   6
         Top             =   315
         Width           =   3000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Access :"
         Height          =   195
         Index           =   3
         Left            =   5100
         TabIndex        =   7
         Top             =   810
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
         Top             =   750
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
         Top             =   365
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Owner :"
         Height          =   195
         Index           =   2
         Left            =   5100
         TabIndex        =   5
         Top             =   365
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2310
      Left            =   120
      TabIndex        =   10
      Top             =   2100
      Width           =   9180
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Add..."
         Height          =   400
         Index           =   0
         Left            =   7800
         TabIndex        =   13
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit..."
         Enabled         =   0   'False
         Height          =   400
         Index           =   1
         Left            =   7800
         TabIndex        =   14
         Top             =   740
         Width           =   1200
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Re&move"
         Enabled         =   0   'False
         Height          =   400
         Left            =   7800
         TabIndex        =   15
         Top             =   1240
         Width           =   1200
      End
      Begin VB.CommandButton cmdRemoveAll 
         Caption         =   "Remo&ve All"
         Enabled         =   0   'False
         Height          =   400
         Left            =   7800
         TabIndex        =   16
         Top             =   1740
         Width           =   1200
      End
      Begin SSDataWidgets_B.SSDBGrid ssGrdRecipients 
         Height          =   1935
         Left            =   1620
         TabIndex        =   12
         Top             =   240
         Width           =   5895
         _Version        =   196617
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   2
         stylesets.count =   5
         stylesets(0).Name=   "ssetSelected"
         stylesets(0).ForeColor=   -2147483634
         stylesets(0).BackColor=   -2147483635
         stylesets(0).Picture=   "frmEmailDefGroup.frx":000C
         stylesets(1).Name=   "ssetHeaderDisabled"
         stylesets(1).ForeColor=   -2147483631
         stylesets(1).BackColor=   -2147483633
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
         stylesets(1).Picture=   "frmEmailDefGroup.frx":0028
         stylesets(2).Name=   "ssetEnabled"
         stylesets(2).ForeColor=   -2147483640
         stylesets(2).BackColor=   -2147483643
         stylesets(2).HasFont=   -1  'True
         BeginProperty stylesets(2).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(2).Picture=   "frmEmailDefGroup.frx":0044
         stylesets(3).Name=   "ssetHeaderEnabled"
         stylesets(3).ForeColor=   -2147483630
         stylesets(3).BackColor=   -2147483633
         stylesets(3).HasFont=   -1  'True
         BeginProperty stylesets(3).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(3).Picture=   "frmEmailDefGroup.frx":0060
         stylesets(4).Name=   "ssetDisabled"
         stylesets(4).ForeColor=   -2147483631
         stylesets(4).BackColor=   -2147483633
         stylesets(4).HasFont=   -1  'True
         BeginProperty stylesets(4).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(4).Picture=   "frmEmailDefGroup.frx":007C
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
         RowNavigation   =   1
         MaxSelectedRows =   1
         StyleSet        =   "ssetDisabled"
         ForeColorEven   =   0
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   10345
         Columns(0).Caption=   "Recipient"
         Columns(0).Name =   "Recipient"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "EmailID"
         Columns(1).Name =   "EmailID"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   10398
         _ExtentY        =   3413
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Addresses :"
         Height          =   195
         Index           =   4
         Left            =   225
         TabIndex        =   11
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   8100
      TabIndex        =   18
      Top             =   4515
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   6800
      TabIndex        =   17
      Top             =   4515
      Width           =   1200
   End
End
Attribute VB_Name = "frmEmailDefGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fOK As Boolean
Private mfCancelled As Boolean
Private mfFromCopy As Boolean
Private mfReadOnly As Boolean

Private mlngEmailGroupID As Long
Private mlngTimeStamp As Long

Private Sub RefreshRecipientsGrid()
  
  With ssGrdRecipients
    .Enabled = True
    .AllowUpdate = (False)
    
    If mfReadOnly Then
      .HeadStyleSet = "ssetHeaderDisabled"
      .StyleSet = "ssetDisabled"
      .ActiveRowStyleSet = "ssetDisabled"
      .SelectTypeRow = ssSelectionTypeNone
    Else
      .HeadStyleSet = "ssetHeaderEnabled"
      .StyleSet = "ssetEnabled"
      .ActiveRowStyleSet = "ssetSelected"
      .SelectTypeRow = ssSelectionTypeSingleSelect
      .RowNavigation = ssRowNavigationLRLock
      
      .SelBookmarks.RemoveAll
      .SelBookmarks.Add .Bookmark
    End If
    
  End With

  UpdateButtonStatus
  
End Sub

Public Property Let Changed(blnChanged As Boolean)
  cmdOK.Enabled = blnChanged
End Property

Public Property Get Changed() As Boolean
  Changed = cmdOK.Enabled
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mfCancelled
End Property

Public Property Get SelectedID() As Long
  SelectedID = mlngEmailGroupID
End Property

Public Function Initialise(pfNew As Boolean, pfCopy As Boolean, Optional plngEmailGroupID As Long)
  Dim sAccess As String
  
  Screen.MousePointer = vbHourglass
 
  fOK = True
  'mfCancelled = True
  mfFromCopy = pfCopy
  mfReadOnly = False
  mlngEmailGroupID = plngEmailGroupID

  
  If pfNew Then
    mlngEmailGroupID = 0
    
    sAccess = GetUserSetting("utils&reports", "dfltaccess emailgroups", ACCESS_READWRITE)
    
    Select Case sAccess
      Case ACCESS_READWRITE
        optReadWrite.Value = True
      Case Else
        optReadOnly.Value = True
    End Select
    
    txtUserName.Text = gsUserName
  Else
    'mfLoading = True
    RetrieveDefinition
    'mfLoading = False
  End If
  
  RefreshRecipientsGrid
  
  'RefreshControls
  
  If mfFromCopy Then
    mlngEmailGroupID = 0
    Me.Changed = True
  Else
    Me.Changed = False
  End If

  Initialise = True

  Screen.MousePointer = vbDefault
      
End Function


Private Sub cmdEdit_Click(Index As Integer)

  Dim frmDefinition As frmEmailDef
  Dim frmSelection As frmDefSel
  Dim lForms As Long
  Dim blnExit As Boolean
  Dim blnOK As Boolean
  Dim lngRow As Long

  Set frmSelection = New frmDefSel
  blnExit = False

  Set frmDefinition = New frmEmailDef
  
  With frmSelection
    Do While Not blnExit
      
      .HideDescription = True
      .Options = edtAdd + edtDelete + edtEdit + edtCopy + edtPrint + edtProperties + edtSelect
      .EnableRun = False
      .TableComboVisible = False
      
      If (Index = 1) And (ssGrdRecipients.Columns("EmailID").Value <> vbNullString) Then
        .SelectedID = ssGrdRecipients.Columns("EmailID").Value
      End If
      
      If .ShowList(utlEmailAddress) Then

        .Show vbModal
        Select Case .Action
        Case edtAdd
          Set frmDefinition = New frmEmailDef
          frmDefinition.Initialise True, .FromCopy
          frmDefinition.Show vbModal
          .SelectedID = frmDefinition.SelectedID
          Unload frmDefinition
          Set frmDefinition = Nothing

        'TM20010808 Fault 2656 - Must validate the definition before allowing the edit/copy.
        Case edtEdit
          Set frmDefinition = New frmEmailDef
          frmDefinition.Initialise False, .FromCopy, .SelectedID
          If Not frmDefinition.Cancelled Then
            frmDefinition.Show vbModal
            If .FromCopy And frmDefinition.SelectedID > 0 Then
              .SelectedID = frmDefinition.SelectedID
            End If
          End If
          Unload frmDefinition
          Set frmDefinition = Nothing

        Case edtPrint
          Set frmDefinition = New frmEmailDef
          frmDefinition.Initialise False, .FromCopy, .SelectedID
          If Not frmDefinition.Cancelled Then
            frmDefinition.PrintDef .SelectedID
          End If
          Unload frmDefinition
          Set frmDefinition = Nothing

        Case edtSelect

          If Index = 1 Then
            lngRow = ssGrdRecipients.AddItemRowIndex(ssGrdRecipients.Bookmark)
            If Not AlreadyUsed(.SelectedID, lngRow) Then
              ssGrdRecipients.RemoveItem lngRow
              ssGrdRecipients.AddItem GetEmailDescription(.SelectedID) & vbTab & CStr(.SelectedID), lngRow
              ssGrdRecipients.Bookmark = ssGrdRecipients.AddItemBookmark(lngRow)
              ssGrdRecipients.SelBookmarks.RemoveAll
              ssGrdRecipients.SelBookmarks.Add ssGrdRecipients.Bookmark
              Changed = True
            End If
          Else
            If Not AlreadyUsed(.SelectedID, -1) Then
              ssGrdRecipients.AddItem GetEmailDescription(.SelectedID) & vbTab & CStr(.SelectedID)
              Changed = True
            End If
          End If
  
          blnExit = True

        Case edtDeselect, 0
          blnExit = True  'cancel

        End Select

      End If

    Loop
  End With

  Unload frmSelection
  Set frmSelection = Nothing

  UpdateButtonStatus

End Sub

Private Sub Form_Activate()
  mfCancelled = True
End Sub

Private Sub RetrieveDefinition()

  Dim rsEmail As Recordset
  Dim strSQL As String
  Dim blnDefinitionCreator As Boolean

  strSQL = "SELECT *, " & _
           "CONVERT(integer,ASRSysEmailGroupName.TimeStamp) AS intTimeStamp " & _
           "FROM ASRSysEmailGroupName " & _
           "WHERE EmailGroupID = " & CStr(mlngEmailGroupID)
  Set rsEmail = datGeneral.GetReadOnlyRecords(strSQL)

  If rsEmail.BOF And rsEmail.EOF Then
    COAMsgBox "Error retriving email group definition", vbCritical, Me.Caption
    Exit Sub
  End If


  If mfFromCopy Then
    txtName.Text = "Copy of " & rsEmail!Name
    txtUserName.Text = gsUserName
    blnDefinitionCreator = True
  Else
    txtName.Text = rsEmail!Name
    txtUserName.Text = rsEmail!UserName
    blnDefinitionCreator = (LCase$(rsEmail!UserName) = LCase$(gsUserName))
  End If

  txtDesc.Text = IIf(rsEmail!Description <> vbNullString, rsEmail!Description, vbNullString)
  mfReadOnly = Not datGeneral.SystemPermission("EmailGroups", "EDIT")

  If Not blnDefinitionCreator Then
    optReadWrite.Enabled = False
    optReadOnly.Enabled = False
  End If

  Select Case rsEmail!Access
  Case ACCESS_READWRITE
    optReadWrite.Value = True
  Case ACCESS_READONLY
    optReadOnly.Value = True
    'MH20040122 Fault 7888
    'mfReadOnly = (mfReadOnly Or Not blnDefinitionCreator)
    mfReadOnly = ((mfReadOnly Or Not blnDefinitionCreator) And Not gfCurrentUserIsSysSecMgr)
  End Select

  If mfReadOnly Then
    ControlsDisableAll Me
    txtDesc.Enabled = True
    txtDesc.Locked = True
    txtDesc.BackColor = vbButtonFace
    txtDesc.ForeColor = vbGrayText
  End If

  mlngTimeStamp = rsEmail!intTimestamp

  rsEmail.Close
  Set rsEmail = Nothing


  strSQL = "SELECT ASRSysEmailGroupItems.*," & _
           " ASRSysEmailAddress.Name as 'AddrName', ASRSysEmailAddress.Fixed as 'AddrFixed'" & _
           " FROM ASRSysEmailGroupItems" & _
           " JOIN ASRSysEmailAddress ON ASRSysEmailGroupItems.EmailDefID = ASRSysEmailAddress.EmailID" & _
           " WHERE EmailGroupID = " & CStr(mlngEmailGroupID) & _
           " ORDER BY AddrName"
  Set rsEmail = datGeneral.GetReadOnlyRecords(strSQL)

  With ssGrdRecipients
    .RemoveAll

    Do While Not rsEmail.EOF
      .AddItem rsEmail!AddrName & " <" & rsEmail!AddrFixed & ">" & vbTab & _
               rsEmail!EmailDefID
      rsEmail.MoveNext
    Loop

  End With

  rsEmail.Close
  Set rsEmail = Nothing

End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdRemove_Click()
  
  Dim lngRow As Long
  
  With ssGrdRecipients
    If .Rows = 1 Then
      .RemoveAll
    Else
      
      lngRow = .AddItemRowIndex(.Bookmark)
      .RemoveItem .AddItemRowIndex(.Bookmark)
      If .Rows > 0 Then
        If lngRow < .Rows Then
          .SelBookmarks.Add .GetBookmark(lngRow)
        ElseIf lngRow = .Rows Then
          .MoveLast
          .SelBookmarks.Add .Bookmark
        End If
      End If
    
    End If
    Changed = True
  End With

  UpdateButtonStatus

End Sub

Private Sub cmdRemoveAll_Click()
  
  If COAMsgBox("Are you sure you want to remove all the Email Addresses from this definition?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
    ssGrdRecipients.RemoveAll
    UpdateButtonStatus
    
    Changed = True
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  Dim pintAnswer As Integer
  
  If (Changed) Then
    
    If (mfCancelled) Then
      pintAnswer = COAMsgBox("You have changed the current definition. Save changes ?", vbQuestion + vbYesNoCancel, Me.Caption)
        
      If pintAnswer = vbYes Then
        Cancel = 1
        cmdOK_Click
        Exit Sub
      ElseIf pintAnswer = vbCancel Then
        Cancel = 1
        Exit Sub
      End If
    End If
  End If

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub optReadOnly_Click()
  Changed = True
End Sub

Private Sub optReadWrite_Click()
  Changed = True
End Sub

Private Sub ssGrdRecipients_DblClick()
  If cmdEdit(1).Enabled Then
    cmdEdit_Click 1
  End If
End Sub

Private Sub ssGrdRecipients_RowLoaded(ByVal Bookmark As Variant)
  
  With ssGrdRecipients

    If (mfReadOnly) Then
      .Columns(0).CellStyleSet "ssetDisabled"
      .Columns(1).CellStyleSet "ssetDisabled"
    Else
      .Columns(0).CellStyleSet "ssetEnabled"
      .Columns(1).CellStyleSet "ssetEnabled"
    End If
   
  End With

End Sub

Private Sub ssGrdRecipients_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)

  ' Configure the grid for the currently selected row.
  On Error GoTo ErrorTrap
  
  Dim iLoop As Integer
  
  With ssGrdRecipients

    ' Set the styleSet of the rows to show which is selected.
    For iLoop = 0 To .Rows - 1
      If (mfReadOnly) Then
        .Columns(0).CellStyleSet "ssetDisabled", iLoop
        .Columns(1).CellStyleSet "ssetDisabled", iLoop
      Else
        If iLoop = .Row Then
          .Columns(0).CellStyleSet "ssetSelected", iLoop
          .Columns(1).CellStyleSet "ssetSelected", iLoop
        Else
          .Columns(0).CellStyleSet "ssetEnabled", iLoop
          .Columns(1).CellStyleSet "ssetEnabled", iLoop
        End If
      End If
    Next iLoop
    
    .SelBookmarks.RemoveAll
    .SelBookmarks.Add .Bookmark

  End With

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  Resume TidyUpAndExit
  
End Sub

Private Sub txtDesc_Change()
  Changed = True
End Sub

Private Sub txtDesc_GotFocus()
  With txtDesc
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtName_Change()
  Changed = True
End Sub

Private Sub txtName_GotFocus()
  With txtName
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub cmdOK_Click()

  If ValidateDefinition Then
    SaveDefinition
    Me.Hide
    mfCancelled = False
  End If

End Sub


Private Function ValidateDefinition() As Boolean

  ValidateDefinition = False
  
  If Trim(txtName.Text) = vbNullString Then
    COAMsgBox "You must give this definition a name.", vbExclamation, Me.Caption
    txtName.SetFocus
    Exit Function
  End If

  If ssGrdRecipients.Rows < 1 Then
    COAMsgBox "You must add at least one email definition.", vbExclamation, Me.Caption
    Exit Function
  End If
  
  If UniqueName(txtName.Text) = False Then
    COAMsgBox "An Email Group called '" & Trim(txtName.Text) & "' already exists.", vbExclamation, Me.Caption
    txtName.SetFocus
    Exit Function
  End If
  
  
  ValidateDefinition = True

End Function


Private Sub SaveDefinition()

  Dim objInsUpd As clsDataInsertUpdate
  Dim strSQL As String
  Dim lngRow As Long
  Dim varBookmark As Variant
  Dim strError As String
  Dim blnNew As Boolean

  blnNew = (mlngEmailGroupID = 0)

  Set objInsUpd = New clsDataInsertUpdate

  objInsUpd.AddColumn "Name", txtName.Text, True
  objInsUpd.AddColumn "Description", txtDesc.Text, True
  objInsUpd.AddColumn "UserName", txtUserName.Text, True
  objInsUpd.AddColumn "Access", IIf(optReadWrite.Value, ACCESS_READWRITE, ACCESS_READONLY), True
  
  mlngEmailGroupID = objInsUpd.InsertUpdate("ASRSysEmailGroupName", "EmailGroupID", mlngEmailGroupID)

  Set objInsUpd = Nothing

  If blnNew Then
    Call UtilCreated(utlEmailGroup, mlngEmailGroupID)
  Else
    Call UtilUpdateLastSaved(utlEmailGroup, mlngEmailGroupID)
  End If


  'Do recipients
  strSQL = "DELETE FROM ASRSysEmailGroupItems " & _
           "WHERE EmailGroupID = " & CStr(mlngEmailGroupID)
  datGeneral.ExecuteSql strSQL, vbNullString

  With ssGrdRecipients
    .MoveFirst

    lngRow = 0
    Do Until lngRow = .Rows

      varBookmark = .GetBookmark(lngRow)
      lngRow = lngRow + 1

      strSQL = "INSERT ASRSysEmailGroupItems " & _
               "(EmailGroupID, EmailDefID) " & _
               "VALUES " & _
               "(" & CStr(mlngEmailGroupID) & ", " & _
               .Columns("EmailID").CellText(varBookmark) & ")"

      If Not datGeneral.ExecuteSql(strSQL, strError) Then
        COAMsgBox "Error saving definition" & vbCrLf & strError, vbCritical, Me.Caption
        Exit Sub
      End If

    Loop

  End With

Exit Sub

LocalErr:
  COAMsgBox "Error saving definition" & vbCrLf & Err.Description, vbCritical, Me.Caption

End Sub


Private Function GetEmailDescription(lngID As Long) As String

  Dim rsEmail As Recordset
  Dim strSQL As String

  strSQL = "SELECT Name, Fixed " & _
           "FROM   ASRSysEmailAddress " & _
           "WHERE  EmailID = " & CStr(lngID)
  Set rsEmail = datGeneral.GetReadOnlyRecords(strSQL)

  GetEmailDescription = vbNullString
  If Not (rsEmail.BOF And rsEmail.EOF) Then
    rsEmail.MoveFirst
    GetEmailDescription = Trim(rsEmail!Name) & " <" & Trim(rsEmail!Fixed) & ">"
  End If

  rsEmail.Close
  Set rsEmail = Nothing

End Function



Private Sub UpdateButtonStatus()

  Dim blnRows As Boolean
  
  ssGrdRecipients.SelBookmarks.RemoveAll
  ssGrdRecipients.SelBookmarks.Add ssGrdRecipients.Bookmark

  blnRows = (ssGrdRecipients.Rows > 0 And Not mfReadOnly)
  cmdEdit(1).Enabled = blnRows
  cmdRemove.Enabled = blnRows
  cmdRemoveAll.Enabled = blnRows

End Sub


Public Sub PrintDef(lngEmailGroupID As Long)

  Dim objPrintDef As clsPrintDef
  Dim rsEmail As Recordset
  Dim strSQL As String
  Dim strName As String

  strSQL = "SELECT *, " & _
           "CONVERT(integer,ASRSysEmailGroupName.TimeStamp) AS intTimeStamp " & _
           "FROM ASRSysEmailGroupName " & _
           "WHERE EmailGroupID = " & CStr(lngEmailGroupID)
  Set rsEmail = datGeneral.GetReadOnlyRecords(strSQL)

  If rsEmail.BOF And rsEmail.EOF Then
    COAMsgBox "Error retriving email group definition", vbCritical, Me.Caption
    Exit Sub
  End If


  Set objPrintDef = New HRProDataMgr.clsPrintDef

  If objPrintDef.IsOK Then
  
    With objPrintDef
      
      If .PrintStart(False) Then
      
        strName = rsEmail!Name
        .PrintHeader "Email Group : " & strName
    
        .PrintNormal "Description : " & IIf(rsEmail!Description <> vbNullString, rsEmail!Description, vbNullString)
        .PrintNormal
    
        .PrintNormal "Owner : " & rsEmail!UserName
        
        .PrintNormal "Access : " & AccessDescription(rsEmail!Access)
        'Select Case rsEmail!Access
        'Case ACCESS_READWRITE: .PrintNormal "Access : Read / Write"
        'Case ACCESS_READONLY: .PrintNormal "Access : Read only"
        'End Select

        rsEmail.Close
        Set rsEmail = Nothing
        
        .PrintNormal
        .PrintTitle "Recipients"
        
        strSQL = "SELECT ASRSysEmailGroupItems.*," & _
                 " ASRSysEmailAddress.Name as 'AddrName', ASRSysEmailAddress.Fixed as 'AddrFixed'" & _
                 " FROM ASRSysEmailGroupItems" & _
                 " JOIN ASRSysEmailAddress ON ASRSysEmailGroupItems.EmailDefID = ASRSysEmailAddress.EmailID" & _
                 " WHERE EmailGroupID = " & CStr(lngEmailGroupID) & _
                 " ORDER BY AddrName"
        Set rsEmail = datGeneral.GetReadOnlyRecords(strSQL)

        Do While Not rsEmail.EOF
          .PrintNonBold rsEmail!AddrName & " <" & rsEmail!AddrFixed & ">"
          rsEmail.MoveNext
        Loop

        .PrintEnd
        .PrintConfirm "Email Group : " & strName, "Email Group Definition"
    
      End If
    End With
  
  End If
    
  rsEmail.Close
  Set rsEmail = Nothing

Exit Sub

LocalErr:
  COAMsgBox "Printing Email Group Definition Failed", vbCritical, "Email Group Definition"

End Sub


Private Function UniqueName(sName As String) As Boolean

  Dim rsName As Recordset
  Dim sSQL As String
    
  sSQL = "SELECT * FROM ASRSysEmailGroupName " & _
         " WHERE Name = '" & Replace(sName, "'", "''") & "' AND EmailGroupID <> " & CStr(mlngEmailGroupID)
    
  Set rsName = datGeneral.GetReadOnlyRecords(sSQL)
  UniqueName = (rsName.BOF And rsName.EOF)
  rsName.Close
    
  Set rsName = Nothing

End Function


Private Function AlreadyUsed(lngID As Long, lngIgnoreRow As Long) As Boolean

  Dim varBookmark As Variant
  Dim lngRow As Long
  
  AlreadyUsed = False
  With ssGrdRecipients
    For lngRow = 0 To .Rows - 1
      If lngRow <> lngIgnoreRow Then
        varBookmark = .AddItemBookmark(lngRow)
        If .Columns("EmailID").CellValue(varBookmark) = lngID Then
          COAMsgBox "This email address is already included in this email group.", vbInformation, "Email Group"
          AlreadyUsed = True
          Exit Function
        End If
      End If
    Next
  End With

End Function

