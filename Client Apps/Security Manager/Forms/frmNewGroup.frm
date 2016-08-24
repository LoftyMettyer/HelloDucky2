VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmNewGroup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New User Group"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   390
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
   HelpContextID   =   8014
   Icon            =   "frmNewGroup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4035
      TabIndex        =   8
      Top             =   4285
      Width           =   1200
   End
   Begin VB.Frame fraAccess 
      Caption         =   "Existing Report / Utility Access :"
      Height          =   3400
      Left            =   150
      TabIndex        =   2
      Top             =   700
      Width           =   5100
      Begin VB.OptionButton optAccess 
         Caption         =   "Copy access from existing user group"
         Height          =   315
         Index           =   0
         Left            =   200
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   3630
      End
      Begin VB.OptionButton optAccess 
         Caption         =   "Configure access for each report / utility type"
         Height          =   315
         Index           =   1
         Left            =   200
         TabIndex        =   5
         Top             =   1100
         Width           =   4725
      End
      Begin VB.ComboBox cboUserGroups 
         Height          =   315
         Left            =   495
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   705
         Width           =   4455
      End
      Begin SSDataWidgets_B.SSDBGrid grdAccess 
         Height          =   1695
         Left            =   495
         TabIndex        =   6
         Top             =   1500
         Width           =   4455
         ScrollBars      =   2
         _Version        =   196617
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   3
         stylesets.count =   2
         stylesets(0).Name=   "ActiveCheckbox"
         stylesets(0).BackColor=   -2147483635
         stylesets(0).HasFont=   -1  'True
         BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(0).Picture=   "frmNewGroup.frx":000C
         stylesets(1).Name=   "ActiveText"
         stylesets(1).ForeColor=   -2147483634
         stylesets(1).BackColor=   -2147483635
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
         stylesets(1).Picture=   "frmNewGroup.frx":0028
         MultiLine       =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   0
         BalloonHelp     =   0   'False
         MaxSelectedRows =   0
         ForeColorEven   =   0
         BackColorEven   =   -2147483643
         BackColorOdd    =   -2147483643
         RowHeight       =   423
         Columns.Count   =   3
         Columns(0).Width=   4630
         Columns(0).Caption=   "Report / Utility"
         Columns(0).Name =   "ReportUtility"
         Columns(0).AllowSizing=   0   'False
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   2752
         Columns(1).Caption=   "Access"
         Columns(1).Name =   "DefaultAccess"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).Locked=   -1  'True
         Columns(1).Style=   3
         Columns(1).Row.Count=   3
         Columns(1).Col.Count=   2
         Columns(1).Row(0).Col(0)=   "Read / Write"
         Columns(1).Row(1).Col(0)=   "Read Only"
         Columns(1).Row(2).Col(0)=   "Hidden"
         Columns(2).Width=   3200
         Columns(2).Visible=   0   'False
         Columns(2).Caption=   "ReportUtilityID"
         Columns(2).Name =   "ReportUtilityID"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   7858
         _ExtentY        =   2990
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
   End
   Begin VB.TextBox txtGroupName 
      Height          =   315
      Left            =   900
      MaxLength       =   30
      TabIndex        =   1
      Top             =   200
      Width           =   4305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2775
      TabIndex        =   7
      Top             =   4285
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      Height          =   195
      Left            =   200
      TabIndex        =   0
      Top             =   255
      Width           =   510
   End
End
Attribute VB_Name = "frmNewGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private miAction As groupAction

Private msAllAccessSetting As String


Public Function AccessConfiguration() As Variant
  ' Return the name of the group that we're using to copy the access for each utility/report type,
  ' or an array of the defined access for each utility/report type.
  Dim avAccess() As Variant
  Dim iLoop As Integer
  Dim varBookmark As Variant
  Dim sGroup As String
  Dim sCopyGroup As String
  Dim vArray As Variant
  
  ' Create an array of the reports/utilities and their defined access.
  ' Column 1 = report/utility type ID
  ' Column 2 = access code (RW/RO/HD)
  ReDim avAccess(2, 0)
  
  ' NB. We're only bothered about the access configuration if we're creating a new
  ' user group.
  If miAction = GROUPACTION_NEW Then
    If optAccess(0).Value And (cboUserGroups.ListCount > 0) Then
      ' NB. We're only bothered about the copy group if the user has
      ' chosen to copy the access from a different group.
      
      ' Check if the selected user group is a new one.
      ' If so, it won't have access granted on the utilities/reports yet.
      ' So get the configuration or group that it's using for it's access.
      sCopyGroup = cboUserGroups.Text
      
      Do While Len(sCopyGroup) > 0
        sGroup = sCopyGroup
        sCopyGroup = gObjGroups(sCopyGroup).AccessCopyGroup
      Loop
      
      AccessConfiguration = sGroup
      
      ' sGroup now gives the root group thats having it's access copied
      ' ie. one which is not copied from another group.
      ' It still might be a new group, so check if its got an access configuration array.
      vArray = gObjGroups(sGroup).AccessConfiguration
      If IsArray(vArray) Then
        If UBound(vArray, 2) > 0 Then
          AccessConfiguration = vArray
        End If
      End If
    
      Exit Function
    Else
      ' NB. We're only bothered about this array if the user's NOT chosen
      ' to copy the access from a different group.
      With grdAccess
        ' 'Update' the grid. This ensures that the correct text/value properties of the cells
        ' are returned when using the cellText/cellValue functions.
        .Update
        
        ' Loop through the grid rows, creating an array entry for each one.
        For iLoop = 1 To (.Rows - 1)
          varBookmark = .AddItemBookmark(iLoop)
          
          ReDim Preserve avAccess(2, UBound(avAccess, 2) + 1)
          avAccess(1, UBound(avAccess, 2)) = CInt(.Columns("ReportUtilityID").CellText(varBookmark))
          avAccess(2, UBound(avAccess, 2)) = CStr(AccessCode(.Columns("DefaultAccess").CellText(varBookmark)))
        Next iLoop
      End With
  
      AccessConfiguration = avAccess
      Exit Function
    End If
  End If
  
  AccessConfiguration = ""
  
End Function


Public Sub Initialise(piAction As groupAction)
  ' Display different controls if the user
  ' is creating a new group, editing an existing group,
  ' or creating a new group by copying an existing one.
  Dim fAccessFrameVisible As Boolean
  Dim objGroup As SecurityGroup
  
  Const GAPUNDERBUTTONS = 625
  Const GAPABOVEBUTTONS = 185
  
  miAction = piAction
  
  fAccessFrameVisible = (piAction = GROUPACTION_NEW)
  fraAccess.Visible = fAccessFrameVisible
  
  cmdOK.Top = IIf(fAccessFrameVisible, fraAccess.Top + fraAccess.Height, txtGroupName.Top + txtGroupName.Height) + GAPABOVEBUTTONS
  cmdCancel.Top = cmdOK.Top
  
  Me.Height = cmdOK.Top + cmdOK.Height + GAPUNDERBUTTONS
  
  If fAccessFrameVisible Then
    ' Populate the group combo.
    With cboUserGroups
      .Clear
      For Each objGroup In gObjGroups
        'JPD 20040109 Fault 7909
        If Not objGroup.DeleteGroup Then
          .AddItem objGroup.Name
        End If
      Next objGroup
      Set objGroup = Nothing
          
      If .ListCount > 0 Then
        .ListIndex = 0
        optAccess_Click (0)
      Else
        optAccess(1).Value = True
        .Enabled = False
        grdAccess.Enabled = True
      End If
    
      optAccess(0).Enabled = (.ListCount > 0)
    End With
    
    
    ' Populate the report/utility access grid.
    With grdAccess
    .RemoveAll
    
    ' NEWACCESS - needs to be updated as each report/utility is updated for the new access.
    .AddItem "<All Reports / Utilities>" & _
      vbTab & AccessDescription(ACCESS_HIDDEN) & _
      vbTab & "-1"
    .AddItem "9-Box Grid Report" & _
      vbTab & AccessDescription(ACCESS_HIDDEN) & _
      vbTab & utlNineBoxGrid
    .AddItem "Batch Job" & _
      vbTab & AccessDescription(ACCESS_HIDDEN) & _
      vbTab & utlBatchJob
    .AddItem "Calendar Report" & _
      vbTab & AccessDescription(ACCESS_HIDDEN) & _
      vbTab & utlCalendarReport
    .AddItem "Career Progression" & _
      vbTab & AccessDescription(ACCESS_HIDDEN) & _
      vbTab & utlCareer
    .AddItem "Cross Tab" & _
      vbTab & AccessDescription(ACCESS_HIDDEN) & _
      vbTab & utlCrossTab
    .AddItem "Custom Report" & _
      vbTab & AccessDescription(ACCESS_HIDDEN) & _
      vbTab & utlCustomReport
    .AddItem "Data Transfer" & _
      vbTab & AccessDescription(ACCESS_HIDDEN) & _
      vbTab & utlDataTransfer
    .AddItem "Envelopes & Labels" & _
      vbTab & AccessDescription(ACCESS_HIDDEN) & _
      vbTab & utlLabel
    .AddItem "Export" & _
      vbTab & AccessDescription(ACCESS_HIDDEN) & _
      vbTab & utlExport
    .AddItem "Global Add" & _
      vbTab & AccessDescription(ACCESS_HIDDEN) & _
      vbTab & UtlGlobalAdd
    .AddItem "Global Delete" & _
      vbTab & AccessDescription(ACCESS_HIDDEN) & _
      vbTab & utlGlobalDelete
    .AddItem "Global Update" & _
      vbTab & AccessDescription(ACCESS_HIDDEN) & _
      vbTab & utlGlobalUpdate
    .AddItem "Import" & _
      vbTab & AccessDescription(ACCESS_HIDDEN) & _
      vbTab & utlImport
    .AddItem "Mail Merge" & _
      vbTab & AccessDescription(ACCESS_HIDDEN) & _
      vbTab & utlMailMerge
    .AddItem "Match Report" & _
      vbTab & AccessDescription(ACCESS_HIDDEN) & _
      vbTab & utlMatchReport
    .AddItem "Organisation Report" & _
      vbTab & AccessDescription(ACCESS_HIDDEN) & _
      vbTab & utlOrganisation
    .AddItem "Record Profile" & _
      vbTab & AccessDescription(ACCESS_HIDDEN) & _
      vbTab & utlRecordProfile
    .AddItem "Report Pack" & _
      vbTab & AccessDescription(ACCESS_HIDDEN) & _
      vbTab & utlReportPack
    .AddItem "Succession Planning" & _
      vbTab & AccessDescription(ACCESS_HIDDEN) & _
      vbTab & utlSuccession
    .AddItem "Talent Report" & _
      vbTab & AccessDescription(ACCESS_HIDDEN) & _
      vbTab & utlTalent
    End With
    
    msAllAccessSetting = ACCESSDESC_HIDDEN
  End If
  
End Sub


Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdOK_Click()
  
  If Trim(txtGroupName.Text) = vbNullString Then
    MsgBox "You must give this group a name.", vbExclamation, Me.Caption
    txtGroupName.SetFocus
    Exit Sub
  End If
  
  If Left(UCase(Trim(txtGroupName.Text)), 6) = "ASRSYS" Then
    MsgBox "'ASRSys' is a reserved word." & vbCrLf & _
            "Please enter another name.", vbInformation + vbOKOnly, App.Title
    txtGroupName.SetFocus
    Exit Sub
  End If

  If Left(UCase(Trim(txtGroupName.Text)), 3) = "DB_" Then
    MsgBox "'DB_' is a reserved word." & vbCrLf & _
            "Please enter another name.", vbInformation + vbOKOnly, App.Title
    txtGroupName.SetFocus
    Exit Sub
  End If



  Me.Tag = "OK"
  Me.Hide

End Sub

Private Sub Form_Activate()
  
  'Assume cancel (unless click OK)
  Me.Tag = "Cancel"
  
  txtGroupName.SetFocus
  
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
  grdAccess.RowHeight = 239

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If Me.Visible Then
    Me.Hide
    Cancel = True
  End If
  
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


Private Sub grdAccess_ComboCloseUp()
    
  'JPD 20030728 Fault 6486
  If (grdAccess.AddItemRowIndex(grdAccess.Bookmark) = 0) And _
    (Len(grdAccess.Columns("DefaultAccess").Text) > 0) Then
    
    ' The '<All Reports / Utilities>' access has changed. Apply the selection to all other Reports / Utilities.
    ForceAccess AccessCode(grdAccess.Columns("DefaultAccess").Text)
    
    grdAccess.MoveFirst
    grdAccess.Col = 1
  End If

End Sub

Private Sub ForceAccess(Optional pvAccess As Variant)

  Dim iLoop As Integer
  Dim varBookmark As Variant
  
  UI.LockWindow grdAccess.hWnd
  
  With grdAccess
    .MoveFirst

    For iLoop = 0 To (.Rows - 1)
      varBookmark = .Bookmark
      .Columns("Access").Text = AccessDescription(CStr(pvAccess))
      .MoveNext
    Next iLoop
    
    .MoveFirst
  End With

  UI.UnlockWindow
End Sub

Private Sub optAccess_Click(Index As Integer)
  ' Enable/disable the controls no longer used.
  cboUserGroups.Enabled = optAccess(0).Value
  grdAccess.Enabled = optAccess(1).Value
  
End Sub

Private Sub txtGroupName_GotFocus()
  
  With txtGroupName
    .SelStart = 0
    .SelLength = Len(.Text)
  End With

End Sub

Private Sub txtGroupName_KeyPress(KeyAscii As Integer)
  ' Check that the char is valid
  KeyAscii = ValidNameChar(KeyAscii, txtGroupName.SelStart)
  
End Sub

Public Property Get GroupName() As String
  'Return the entered group name.
  GroupName = txtGroupName.Text
  
End Property

Public Property Let GroupName(sNewGroupName As String)
  'Sets the Group Name
  txtGroupName.Text = sNewGroupName
  Me.Caption = IIf(Len(sNewGroupName) > 0, "Rename User Group", "New User Group")
End Property

Public Sub CopyGroup()

  Me.Caption = "Copy User Group"

End Sub
