VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{1EE59219-BC23-4BDF-BB08-D545C8A38D6D}#1.1#0"; "COA_Line.ocx"
Begin VB.Form frmWorkflowSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Workflow Configuration"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5070
   Icon            =   "frmWorkflowSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab ssTabStrip 
      Height          =   5660
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   6800
      _ExtentX        =   11986
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&Web Site"
      TabPicture(0)   =   "frmWorkflowSetup.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraWebSite"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraWebSiteLogin"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraAuthenticate"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Personnel Identification"
      TabPicture(1)   =   "frmWorkflowSetup.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDelegation"
      Tab(1).Control(1)=   "fraPersonnelTable"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Service"
      TabPicture(2)   =   "frmWorkflowSetup.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraService"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&Mobile Specifics"
      TabPicture(3)   =   "frmWorkflowSetup.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraMobileKey"
      Tab(3).Control(1)=   "Frame1"
      Tab(3).ControlCount=   2
      Begin VB.Frame fraAuthenticate 
         Caption         =   "Authentication :"
         Height          =   900
         Left            =   150
         TabIndex        =   48
         Top             =   1500
         Width           =   6500
         Begin VB.CheckBox chkRequireAuthorization 
            Caption         =   "Authenticate after all email lin&ks"
            Height          =   480
            Left            =   135
            TabIndex        =   4
            Top             =   270
            Width           =   3195
         End
      End
      Begin VB.Frame fraMobileKey 
         Caption         =   "Custom.Web.Config :"
         Height          =   975
         Left            =   -74850
         TabIndex        =   44
         Top             =   3090
         Visible         =   0   'False
         Width           =   6500
         Begin VB.CommandButton cmdGenMobileKey 
            Caption         =   "&Generate"
            Enabled         =   0   'False
            Height          =   400
            Left            =   5100
            TabIndex        =   45
            Top             =   315
            Width           =   1200
         End
         Begin VB.Label lblGetMobileKey 
            AutoSize        =   -1  'True
            Caption         =   "Generate Mobile Web Config Keys :"
            Height          =   195
            Left            =   195
            TabIndex        =   46
            Top             =   420
            Width           =   3060
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Personnel Table :"
         Height          =   2445
         Left            =   -74850
         TabIndex        =   34
         Top             =   500
         Width           =   6500
         Begin VB.ComboBox cboMobUserActivated 
            Height          =   315
            Left            =   2370
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   1890
            Width           =   3975
         End
         Begin VB.ComboBox cboMobLoginName 
            Height          =   315
            Left            =   2370
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   1110
            Width           =   3975
         End
         Begin VB.ComboBox cboMobPersonnelTable 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2370
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   300
            Width           =   3975
         End
         Begin VB.ComboBox cboMobEMailColumn 
            Height          =   315
            Left            =   2370
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   720
            Width           =   3975
         End
         Begin VB.ComboBox cboMobLeavingDateColumn 
            Height          =   315
            Left            =   2370
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   1500
            Width           =   3975
         End
         Begin VB.Label lblMobUserActivated 
            Caption         =   "User Activated Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   47
            Top             =   1935
            Width           =   2115
         End
         Begin VB.Label lblMobEmailAddresses 
            Caption         =   "Registration Email Address :"
            Height          =   390
            Left            =   195
            TabIndex        =   43
            Top             =   660
            Width           =   1770
         End
         Begin VB.Label lblMobLoginNameColumn 
            AutoSize        =   -1  'True
            Caption         =   "Mobile Login Username :"
            Height          =   195
            Left            =   195
            TabIndex        =   42
            Top             =   1170
            Width           =   2115
         End
         Begin VB.Label Label1 
            Caption         =   "Personnel Table :"
            Enabled         =   0   'False
            Height          =   195
            Left            =   195
            TabIndex        =   41
            Top             =   360
            Width           =   1560
         End
         Begin VB.Label lblMobLeavingDateColumn 
            Caption         =   "Login Expiry Date :"
            Height          =   195
            Left            =   195
            TabIndex        =   40
            Top             =   1545
            Width           =   1995
         End
      End
      Begin VB.Frame fraWebSiteLogin 
         Caption         =   "Login :"
         Height          =   1700
         Left            =   150
         TabIndex        =   5
         Top             =   2565
         Visible         =   0   'False
         Width           =   6500
         Begin VB.CommandButton cmdTestLogon 
            Caption         =   "&Test Login"
            Height          =   400
            Left            =   5100
            TabIndex        =   10
            Top             =   1100
            Width           =   1200
         End
         Begin VB.TextBox txtUID 
            Height          =   315
            Left            =   1230
            MaxLength       =   128
            TabIndex        =   7
            Top             =   300
            Width           =   5070
         End
         Begin VB.TextBox txtPWD 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   1230
            MaxLength       =   128
            PasswordChar    =   "*"
            TabIndex        =   9
            Top             =   700
            Width           =   5070
         End
         Begin VB.Label lblUID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User :"
            Height          =   195
            Left            =   200
            TabIndex        =   6
            Top             =   360
            Width           =   435
         End
         Begin VB.Label lblPassword 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password :"
            Height          =   195
            Left            =   195
            TabIndex        =   8
            Top             =   765
            Width           =   930
         End
      End
      Begin VB.Frame fraService 
         Caption         =   "Service :"
         Height          =   925
         Left            =   -74850
         TabIndex        =   30
         Top             =   500
         Width           =   6500
         Begin VB.CommandButton cmdSuspendService 
            Caption         =   "S&uspend"
            Height          =   400
            Left            =   5100
            TabIndex        =   32
            Top             =   300
            Width           =   1200
         End
         Begin VB.Label lblServiceSuspended 
            AutoSize        =   -1  'True
            Caption         =   "The Workflow Service is (not) currently suspended."
            Height          =   195
            Left            =   195
            TabIndex        =   31
            Top             =   360
            Width           =   3690
         End
      End
      Begin VB.Frame fraDelegation 
         Caption         =   "Out of Office Delegation :"
         Height          =   1560
         Left            =   -74850
         TabIndex        =   23
         Top             =   3900
         Width           =   6500
         Begin VB.CheckBox chkCopyDelegateEmail 
            Caption         =   "Copy &email to original recipient"
            Height          =   315
            Left            =   200
            TabIndex        =   29
            Top             =   1100
            Width           =   3180
         End
         Begin VB.TextBox txtEmail 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   700
            Width           =   3585
         End
         Begin VB.CommandButton cmdEmail 
            Caption         =   "..."
            Height          =   315
            Left            =   5985
            TabIndex        =   28
            Top             =   700
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.ComboBox cboDelegationActivatedColumn 
            Height          =   315
            Left            =   2400
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   300
            Width           =   3900
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "'Delegate To' Email :"
            Height          =   195
            Left            =   195
            TabIndex        =   26
            Top             =   765
            Width           =   1785
         End
         Begin VB.Label lblDelegationActivatedColumn 
            BackStyle       =   0  'Transparent
            Caption         =   "Activation Column :"
            Height          =   195
            Left            =   195
            TabIndex        =   24
            Top             =   360
            Width           =   1755
         End
      End
      Begin VB.Frame fraPersonnelTable 
         Caption         =   "Personnel Table :"
         Height          =   3320
         Left            =   -74850
         TabIndex        =   11
         Top             =   500
         Width           =   6500
         Begin VB.ComboBox cboLoginName 
            Height          =   315
            Index           =   1
            Left            =   2370
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1100
            Width           =   3975
         End
         Begin VB.CommandButton cmdRemoveAllEmailAddressColumns 
            Caption         =   "Remo&ve All"
            Height          =   400
            Left            =   5100
            TabIndex        =   22
            Top             =   2700
            Width           =   1200
         End
         Begin VB.CommandButton cmdRemoveEmailAddressColumn 
            Caption         =   "&Remove"
            Height          =   400
            Left            =   5100
            TabIndex        =   21
            Top             =   2200
            Width           =   1200
         End
         Begin VB.CommandButton cmdAddEmailAddressColumn 
            Caption         =   "&Add ..."
            Height          =   400
            Left            =   5100
            TabIndex        =   20
            Top             =   1700
            Width           =   1200
         End
         Begin VB.ComboBox cboLoginName 
            Height          =   315
            Index           =   0
            Left            =   2370
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   700
            Width           =   3975
         End
         Begin VB.ComboBox cboPersonnelTable 
            Height          =   315
            Left            =   2370
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   300
            Width           =   3975
         End
         Begin SSDataWidgets_B.SSDBGrid grdEmailAddressColumns 
            Height          =   1395
            Left            =   2370
            TabIndex        =   19
            Top             =   1695
            Width           =   2625
            _Version        =   196617
            DataMode        =   2
            RecordSelectors =   0   'False
            GroupHeaders    =   0   'False
            Col.Count       =   2
            AllowUpdate     =   0   'False
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
            SelectTypeRow   =   3
            SelectByCell    =   -1  'True
            BalloonHelp     =   0   'False
            MaxSelectedRows =   0
            ForeColorEven   =   0
            BackColorEven   =   -2147483643
            BackColorOdd    =   -2147483643
            RowHeight       =   423
            ExtraHeight     =   79
            Columns.Count   =   2
            Columns(0).Width=   3200
            Columns(0).Visible=   0   'False
            Columns(0).Caption=   "ColumnID"
            Columns(0).Name =   "ColumnID"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   8281
            Columns(1).Caption=   "Column"
            Columns(1).Name =   "ColumnName"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            TabNavigation   =   1
            _ExtentX        =   4630
            _ExtentY        =   2461
            _StockProps     =   79
            BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin COALine.COA_Line ASRDummyLine1 
            Height          =   30
            Left            =   200
            Top             =   1560
            Width           =   6100
            _ExtentX        =   10769
            _ExtentY        =   53
         End
         Begin VB.Label lblEmailAddresses 
            Caption         =   "Email Addresses :"
            Height          =   195
            Left            =   195
            TabIndex        =   18
            Top             =   1755
            Width           =   1620
         End
         Begin VB.Label lblLoginNameColumn 
            AutoSize        =   -1  'True
            Caption         =   "Login Name Column(s) :"
            Height          =   195
            Left            =   195
            TabIndex        =   14
            Top             =   765
            Width           =   2100
         End
         Begin VB.Label lblPersonnelTable 
            Caption         =   "Personnel Table :"
            Height          =   195
            Left            =   195
            TabIndex        =   12
            Top             =   360
            Width           =   1560
         End
      End
      Begin VB.Frame fraWebSite 
         Caption         =   "Address :"
         Height          =   850
         Left            =   150
         TabIndex        =   1
         Top             =   500
         Width           =   6500
         Begin VB.TextBox txtURL 
            Height          =   315
            Left            =   1230
            MaxLength       =   200
            TabIndex        =   3
            Top             =   300
            Width           =   5070
         End
         Begin VB.Label lblURL 
            Caption         =   "URL :"
            Height          =   195
            Left            =   195
            TabIndex        =   2
            Top             =   360
            Width           =   570
         End
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   4485
      TabIndex        =   33
      Top             =   5960
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5750
      TabIndex        =   17
      Top             =   5960
      Width           =   1200
   End
End
Attribute VB_Name = "frmWorkflowSetup"
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
Private mlngSecondLoginColumnID As Long
Private mlngDelegationActivatedColumnID As Long
Private mlngDelegateEmailID As Long
Private mfCopyDelegateEmail As Boolean
Private mfRequiresAuthorization As Boolean
Private mfServiceSuspended As Boolean
Private mlngMobPersonnelTableID As Long
Private mlngMobLoginColumnID As Long
' Private mlngMobUniqueEmailColumnID As Long
Private mlngWorkEmailColumnID As Long
Private mlngMobLeavingDateColumnID As Long
Private mlngMobActivatedColumnID As Long

Private msOriginalURL As String
Private msOriginalUser As String
Private msOriginalPassword As String

' CONSTANTS
'
' Page number constants.
Private Const miPAGE_WEBSITE = 0
Private Const miPAGE_PERSONNELIDENTIFICATION = 1
Private Const miPAGE_SERVICE = 2


Public Property Get Changed() As Boolean
  Changed = mfChanged
End Property
Public Property Let Changed(ByVal pblnChanged As Boolean)
  mfChanged = pblnChanged
  If Not mbLoading Then cmdOK.Enabled = True
End Property
Private Sub RefreshServiceSuspended()

  lblServiceSuspended.Caption = "The Workflow Service is " & _
    IIf(mfServiceSuspended, "", "not ") & _
    "currently suspended."
  cmdSuspendService.Caption = IIf(mfServiceSuspended, "&Resume", "S&uspend")
  
End Sub

Private Sub cboDelegationActivatedColumn_Click()
  With cboDelegationActivatedColumn
    mlngDelegationActivatedColumnID = .ItemData(.ListIndex)
  End With

  Changed = True
  RefreshControls

End Sub




Private Sub cboLoginName_Click(Index As Integer)
  With cboLoginName(Index)
    If Index = 1 Then
      mlngSecondLoginColumnID = .ItemData(.ListIndex)
    Else
      mlngLoginColumnID = .ItemData(.ListIndex)
    End If
  End With
  
  If Not mbLoading Then
    mbLoading = True
    RefreshPersonnelColumnControls
    mbLoading = False
    Changed = True
  End If
  
  RefreshControls

End Sub


Private Sub cboMobEMailColumn_Click()
  With cboMobEMailColumn
    ' mlngMobUniqueEmailColumnID = .ItemData(.ListIndex)
    mlngWorkEmailColumnID = .ItemData(.ListIndex)
  End With
  
  If Not mbLoading Then
    mbLoading = True
    'RefreshPersonnelColumnControls
    mbLoading = False
    Changed = True
  End If
  
  RefreshControls

End Sub

Private Sub cboMobLeavingDateColumn_Click()
 With cboMobLeavingDateColumn
    mlngMobLeavingDateColumnID = .ItemData(.ListIndex)
  End With
  
  If Not mbLoading Then
    mbLoading = True
    'RefreshPersonnelColumnControls
    mbLoading = False
    Changed = True
  End If
  
  RefreshControls

End Sub

Private Sub cboMobLoginName_Click()
  With cboMobLoginName
    mlngMobLoginColumnID = .ItemData(.ListIndex)
  End With
  
  If Not mbLoading Then
    mbLoading = True
    'RefreshPersonnelColumnControls
    mbLoading = False
    Changed = True
  End If
  
  RefreshControls
End Sub

Private Sub cboMobUserActivated_Click()
  With cboMobUserActivated
    mlngMobActivatedColumnID = .ItemData(.ListIndex)
  End With
  
  If Not mbLoading Then
    mbLoading = True
    'RefreshPersonnelColumnControls
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

    ' Check if the Delegate Email is fixed.
    fFixedDelegateEmail = True
    If mlngDelegateEmailID > 0 Then
      ' Create a new Email object.
      Set objEmail = New clsEmailAddr

      With objEmail
        .EmailID = mlngDelegateEmailID
        .ConstructEmail
        
        fFixedDelegateEmail = (Len(.Fixed) > 0)
      End With
      Set objEmail = Nothing
    End If
    
    fGoAhead = True
    If (mlngLoginColumnID > 0) _
      Or (mlngSecondLoginColumnID > 0) _
      Or (grdEmailAddressColumns.Rows > 0) _
      Or (mlngDelegationActivatedColumnID > 0) _
      Or (Not fFixedDelegateEmail) Then
      
      fGoAhead = (MsgBox("Warning: Changing the Personnel table will reset all Personnel Identification and Delegation parameters." & vbCrLf & _
      "Are you sure you wish to continue?", _
      vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes)
    End If

    If fGoAhead Then
      mlngPersonnelTableID = cboPersonnelTable.ItemData(cboPersonnelTable.ListIndex)

      grdEmailAddressColumns.RemoveAll
      
      If Not fFixedDelegateEmail Then
        mlngDelegateEmailID = 0
        GetEmailAddressDetails
      End If
      
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

Private Sub chkCopyDelegateEmail_Click()
  mfCopyDelegateEmail = (chkCopyDelegateEmail.value = vbChecked)
  Changed = True
  RefreshControls

End Sub

Private Sub chkRequireAuthorization_Click()
  mfRequiresAuthorization = (chkRequireAuthorization.value = vbChecked)
  Changed = True
End Sub

Private Sub cmdAddEmailAddressColumn_Click()
  Dim sRow As String
  Dim frmColumn As New frmWorkflowEmailAddressColumn

  With frmColumn
    .Initialize 0, _
      mlngPersonnelTableID, _
      SelectedEmailColumns

    If Not .Cancelled Then
      .Show vbModal
    End If

    If Not .Cancelled Then
      sRow = CStr(.ColumnID) _
        & vbTab & .ColumnName
      
      With grdEmailAddressColumns
        .AddItem sRow

        .Bookmark = .AddItemBookmark(.Rows - 1)
        .SelBookmarks.RemoveAll
        .SelBookmarks.Add .Bookmark
      End With
      
      Changed = True
    End If
  End With

  UnLoad frmColumn
  Set frmColumn = Nothing

End Sub

Private Function SelectedEmailColumns() As String
  ' Return a string of the selected email column IDs
  Dim iLoop As Integer
  Dim varBookMark As Variant
  Dim sSelectedIDs As String
  
  sSelectedIDs = "0"
  
  With grdEmailAddressColumns
    For iLoop = 0 To (.Rows - 1)
      varBookMark = .AddItemBookmark(iLoop)

      sSelectedIDs = sSelectedIDs & "," & .Columns("ColumnID").CellText(varBookMark)
    Next iLoop
  End With

  SelectedEmailColumns = sSelectedIDs
  
End Function

Private Sub cmdCancel_Click()
  Dim pintAnswer As Integer
    If Changed = True And cmdOK.Enabled Then
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

Private Sub cmdEmail_Click()
  ' Display the Email selection form.
  Dim objEmail As clsEmailAddr
  Dim wfTemp As VB.Control
  
  ' Create a new Email object.
  Set objEmail = New clsEmailAddr

  ' Initialize the Email object.
  With objEmail
    .EmailID = mlngDelegateEmailID
    .TableID = mlngPersonnelTableID

    ' Instruct the Email object to handle the selection.
    If .SelectEmail Then
      mlngDelegateEmailID = .EmailID
      txtEmail.Text = .EmailName

      Changed = True
    Else
      ' Check in case the original Email has been deleted.
      With recEmailAddrEdit
        .Index = "idxID"
        .Seek "=", mlngDelegateEmailID

        If .NoMatch Then
          If mlngDelegateEmailID > 0 Then
            Changed = True
          End If
          
          mlngDelegateEmailID = 0
          txtEmail.Text = vbNullString
        Else
          If !Deleted Then
            If mlngDelegateEmailID > 0 Then
              Changed = True
            End If
            
            mlngDelegateEmailID = 0
            txtEmail.Text = vbNullString
          End If
        End If
      End With
    End If
  End With

  ' Disassociate object variables.
  Set objEmail = Nothing

  RefreshControls

End Sub

Private Sub cmdGenMobileKey_Click()
  Dim sNewQueryString As String
  Dim frmChangedPlatform As frmChangedPlatform
  Dim sURL As String
  Dim sNewString As String
  
  sNewQueryString = GetWorkflowQueryString(-1, -1, txtUID.Text, txtPWD.Tag)

  Set frmChangedPlatform = New frmChangedPlatform
  frmChangedPlatform.ResetList
  
  ' MobileKey key
  sNewString = "<add key=""MobileKey"" value=""" & sNewQueryString & """/>"
  frmChangedPlatform.AddToList sNewString
  
  ' Workflow URL key
  sURL = Trim(txtURL.Text)
  If UCase(Right(sURL, 5)) <> ".ASPX" _
    And Right(sURL, 1) <> "/" _
    And Len(sURL) > 0 Then

    sURL = sURL + "/"
  End If
  sNewString = "<add key=""WorkflowURL"" value=""" & sURL & """/>"
  frmChangedPlatform.AddToList sNewString
  

  frmChangedPlatform.Width = (3 * Screen.Width / 4)
  frmChangedPlatform.Height = (Screen.Height / 2)

  frmChangedPlatform.ShowMessage 2
    
  UnLoad frmChangedPlatform
  Set frmChangedPlatform = Nothing
     

End Sub

Private Sub cmdOK_Click()
  Dim fSaveOK As Boolean
  Dim rsWorkflows As DAO.Recordset
  Dim sSQL As String
  Dim sURL As String
  Dim sTemp As String
  Dim sNewQueryString As String
  Dim asWorkflows() As String
  Dim sOldString As String
  Dim sNewString As String
  Dim frmChangedPlatform As frmChangedPlatform
  Dim iLoop As Integer
  Dim fGoAhead As Boolean
  
  fSaveOK = True
  
  ReDim asWorkflows(2, 0)
  ' Column 0 = Workflow Name
  ' Column 1 = Original URL + QueryString
  ' Column 2 = New URL + QueryString
  
  If (msOriginalURL <> Trim(txtURL.Text)) _
    Or (msOriginalUser <> txtUID.Text) _
    Or (msOriginalPassword <> txtPWD.Tag) Then
    ' QueryString parameters changed. Inform user that this will affect external initiation URLs (if there are any).
    
    sSQL = "SELECT tmpWorkflows.name," & _
      "   tmpWorkflows.ID," & _
      "   tmpWorkflows.queryString" & _
      " FROM tmpWorkflows " & _
      " WHERE tmpWorkflows.initiationType = " & CStr(WORKFLOWINITIATIONTYPE_EXTERNAL) & _
      "   AND tmpWorkflows.deleted = FALSE"
           
    Set rsWorkflows = daoDb.OpenRecordset(sSQL)
    With rsWorkflows
      If Not (.EOF And .BOF) Then
        Do Until .EOF
        
          sOldString = GetWorkflowURL
          If Len(sOldString) > 0 Then
            sOldString = IIf(Len(!queryString) > 0, sOldString & "?" & !queryString, "")
          End If

          ' NPG20141118 - always regenerate querystrings if url changes
          sNewQueryString = GetWorkflowQueryString(!ID * -1, -1, txtUID.Text, txtPWD.Tag)
          
          sSQL = "UPDATE tmpWorkflows" & _
            " SET changed = TRUE," & _
            "   queryString = '" & Replace(sNewQueryString, "'", "''") & "'" & _
            " WHERE ID = " & CStr(!ID)
          daoDb.Execute sSQL
          
          If (msOriginalURL <> Trim(txtURL.Text)) Then
            sURL = Trim(txtURL.Text)
            If UCase(Right(sURL, 5)) <> ".ASPX" _
              And Right(sURL, 1) <> "/" _
              And Len(sURL) > 0 Then

              sURL = sURL + "/"
            End If
          Else
            sURL = GetWorkflowURL
          End If
         
          sNewString = IIf((Len(sURL) > 0) And (Len(sNewQueryString) > 0), sURL & "?" & sNewQueryString, "")
          
          If (sOldString <> sNewString) _
            And (Len(sOldString) + Len(sNewString) > 0) Then
            ReDim Preserve asWorkflows(2, UBound(asWorkflows, 2) + 1)
            asWorkflows(0, UBound(asWorkflows, 2)) = !Name
            asWorkflows(1, UBound(asWorkflows, 2)) = sOldString
            asWorkflows(2, UBound(asWorkflows, 2)) = sNewString
          End If
          
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rsWorkflows = Nothing
    
  End If
  
  ' NPG20120514 - Fault HRPRO-2236
  ' all mobile cdropdowns none or selected?
  If Application.MobileModule Then
    If (mlngMobLoginColumnID = 0 And _
      mlngMobLeavingDateColumnID = 0 And _
      mlngMobActivatedColumnID = 0) Or _
      (mlngMobLoginColumnID > 0 And _
      mlngMobLeavingDateColumnID > 0 And _
      mlngMobActivatedColumnID > 0) Then
      fSaveOK = True
    Else
      fSaveOK = False
    End If
  End If
  
  
  
  
  If fSaveOK Then
    If UBound(asWorkflows, 2) > 0 Then
      Set frmChangedPlatform = New frmChangedPlatform
      frmChangedPlatform.ResetList

      For iLoop = 1 To UBound(asWorkflows, 2)
        frmChangedPlatform.AddToList asWorkflows(0, iLoop), _
          IIf(Len(asWorkflows(1, iLoop)) > 0, asWorkflows(1, iLoop), "<none>"), _
          IIf(Len(asWorkflows(2, iLoop)) > 0, asWorkflows(2, iLoop), "<none>")
      Next iLoop

      frmChangedPlatform.Width = (3 * Screen.Width / 4)
      frmChangedPlatform.Height = (Screen.Height / 2)

      frmChangedPlatform.ShowMessage 1
      UnLoad frmChangedPlatform
      Set frmChangedPlatform = Nothing
     End If
     
    SaveChanges
  
    UnLoad Me
  
  Else
      MsgBox "Mobile specifics not correctly configured." & vbCrLf & vbCrLf & "All or none of the columns must be selected.", vbExclamation, Me.Caption

  End If
  

End Sub

Private Sub SaveChanges()
  ' Save the parameter values to the local database.
  Dim iLoop As Integer
  Dim varBookMark As Variant
  Dim sColumnID As String
  Dim sSQL As String
  
  SaveModuleSetting gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_URL, gsPARAMETERTYPE_OTHER, RTrim(Replace(txtURL.Text, "\", "/"))
  
  SaveWebLogon Replace(txtUID.Text, ";", ""), Replace(txtPWD.Tag, ";", "")
  
  SaveModuleSetting gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_PERSONNELTABLE, gsPARAMETERTYPE_TABLEID, mlngPersonnelTableID
  SaveModuleSetting gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_LOGINNAME, gsPARAMETERTYPE_COLUMNID, mlngLoginColumnID
  SaveModuleSetting gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_SECONDLOGINNAME, gsPARAMETERTYPE_COLUMNID, mlngSecondLoginColumnID
  
  SaveModuleSetting gsMODULEKEY_MOBILE, gsPARAMETERKEY_PERSONNELTABLE, gsPARAMETERTYPE_TABLEID, mlngMobPersonnelTableID
  SaveModuleSetting gsMODULEKEY_MOBILE, gsPARAMETERKEY_LOGINNAME, gsPARAMETERTYPE_COLUMNID, mlngMobLoginColumnID
  ' SaveModuleSetting gsMODULEKEY_MOBILE, gsPARAMETERKEY_UNIQUEEMAILCOLUMN, gsPARAMETERTYPE_COLUMNID, mlngMobUniqueEmailColumnID
  SaveModuleSetting gsMODULEKEY_MOBILE, gsPARAMETERKEY_WORKEMAIL, gsPARAMETERTYPE_COLUMNID, mlngWorkEmailColumnID
  SaveModuleSetting gsMODULEKEY_MOBILE, gsPARAMETERKEY_LEAVINGDATE, gsPARAMETERTYPE_COLUMNID, mlngMobLeavingDateColumnID
  SaveModuleSetting gsMODULEKEY_MOBILE, gsPARAMETERKEY_MOBILEACTIVATED, gsPARAMETERTYPE_COLUMNID, mlngMobActivatedColumnID
  
  ' Save the Email Address columns
  ' Clear the current database values.
  daoDb.Execute "DELETE FROM tmpModuleSetup" & _
    " WHERE moduleKey = '" & gsMODULEKEY_WORKFLOW & "'" & _
    " AND parameterkey = '" & gsPARAMETERKEY_EMAILCOLUMN & "'", _
    dbFailOnError

  With grdEmailAddressColumns
    For iLoop = 0 To (.Rows - 1)
      varBookMark = .AddItemBookmark(iLoop)

      sColumnID = .Columns("ColumnID").CellText(varBookMark)
    
      sSQL = "INSERT INTO tmpModuleSetup" & _
        " (moduleKey, parameterkey, parameterType, parametervalue)" & _
        " VALUES" & _
        " ('" & gsMODULEKEY_WORKFLOW & "','" & gsPARAMETERKEY_EMAILCOLUMN & "','" & gsPARAMETERTYPE_COLUMNID & "'," & sColumnID & ")"
      daoDb.Execute sSQL, dbFailOnError
    Next iLoop
  End With
    
  SaveModuleSetting gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_DELEGATIONACTIVATEDCOLUMN, gsPARAMETERTYPE_COLUMNID, mlngDelegationActivatedColumnID
  SaveModuleSetting gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_DELEGATEEMAIL, gsPARAMETERTYPE_EMAILID, mlngDelegateEmailID
  SaveModuleSetting gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_COPYDELEGATEEMAIL, gsPARAMETERTYPE_OTHER, IIf(mfCopyDelegateEmail, "TRUE", "FALSE")
  SaveModuleSetting gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_REQUIRESAUTHORIZATION, gsPARAMETERTYPE_OTHER, IIf(mfRequiresAuthorization, "TRUE", "FALSE")
   
  Application.Changed = True

End Sub

Private Sub cmdRemoveAllEmailAddressColumns_Click()
  
  grdEmailAddressColumns.RemoveAll
    
  Changed = True

  RefreshControls

End Sub

Private Sub cmdRemoveEmailAddressColumn_Click()
  Dim sRowsToDelete As String
  Dim iCount As Integer
  Dim ctlGrid As SSDBGrid
  Dim iIndex As Integer
  
  sRowsToDelete = ","

  With grdEmailAddressColumns
    If .Rows = 1 Then
      cmdRemoveAllEmailAddressColumns_Click
    Else
      For iCount = 0 To .SelBookmarks.Count - 1
        sRowsToDelete = sRowsToDelete & CStr(.AddItemRowIndex(.SelBookmarks(iCount))) & ","
      Next iCount
      
      For iCount = (.Rows - 1) To 0 Step -1
        If InStr(sRowsToDelete, "," & CStr(iCount) & ",") > 0 Then
          .Bookmark = .AddItemBookmark(iCount)
          .RemoveItem iCount
        End If
      Next iCount
    End If
    
    If .Rows > 0 Then
      .SelBookmarks.RemoveAll
      .MoveFirst
      .SelBookmarks.Add .Bookmark
    End If
  End With
  
  Changed = True

End Sub


Private Sub cmdSuspendService_Click()
  Dim sMsg As String
  
  sMsg = "This action will " & _
    IIf(mfServiceSuspended, "resume", "suspend") & _
    " the Workflow Service with immediate effect." & vbCrLf & _
    "Do you wish to continue ?"
    
  If MsgBox(sMsg, vbYesNo + vbQuestion, Me.Caption) = vbYes Then
    mfServiceSuspended = Not mfServiceSuspended
    SaveSystemSetting "workflow", "suspended", IIf(mfServiceSuspended, "1", "0")
    RefreshServiceSuspended
  End If
  
End Sub

Private Sub cmdTestLogon_Click()

  Dim objTestConn As ADODB.Connection
  Dim sConnect As String

  On Error GoTo LocalErr

  If Trim(txtUID.Text) = vbNullString Then
    MsgBox "You must enter a user name.", vbInformation, Me.Caption
    Exit Sub
  End If
  
  Screen.MousePointer = vbHourglass

  sConnect = "Driver=SQL Server;" & _
    "Server=" & Replace(gsServerName, ";", "") & ";" & _
    "UID=" & Replace(txtUID.Text, ";", "") & ";" & _
    "PWD=" & Replace(txtPWD.Tag, ";", "") & ";" & _
    "Database=" & gsDatabaseName & ";Pooling=false;App=Test OpenHR Workflow"

  Set objTestConn = New ADODB.Connection
  With objTestConn
    .ConnectionString = sConnect
    .Provider = "SQLOLEDB"
    .CommandTimeout = 10
    .ConnectionTimeout = 30   'MH20030911 Fault 6944
    .CursorLocation = adUseServer
    .Mode = adModeReadWrite
    .Properties("Packet Size") = 32767
    .Open
    .Close
  End With

  Set objTestConn = Nothing

  Screen.MousePointer = vbDefault
  MsgBox "Test completed successfully.", vbInformation, Me.Caption

Exit Sub

LocalErr:
  Screen.MousePointer = vbDefault
  
  MsgBox "Error during Workflow login test." & vbCrLf & _
    ADOConError(objTestConn), vbInformation, Me.Caption

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
  Const GRIDROWHEIGHT = 239
  
  grdEmailAddressColumns.RowHeight = GRIDROWHEIGHT
  
  mbLoading = True
  cmdOK.Enabled = False
  
  ' Position the form in the same place it was last time.
  Me.Top = GetPCSetting(Me.Name, "Top", (Screen.Height - Me.Height) / 2)
  Me.Left = GetPCSetting(Me.Name, "Left", (Screen.Width - Me.Width) / 2)
  
  
  mblnReadOnly = (Application.AccessMode <> accFull And _
                  Application.AccessMode <> accSupportMode)
  
  If mblnReadOnly Then
    ControlsDisableAll Me
    grdEmailAddressColumns.Enabled = True
  End If

  ' Read the current settings from the database.
  ReadParameters
  cboPersonnelTable_Refresh
  
  ssTabStrip.Tab = miPAGE_WEBSITE
  ssTabStrip_Click miPAGE_WEBSITE
  
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
  Dim sURL As String
  Dim sUser As String
  Dim sPassword As String
  Dim lngColumnID As Long
  Dim sRow As String
  Dim lngPersModulePersonnelTableID As Long
  Dim objEmail As clsEmailAddr
  
  ' ------------------------------------------
  ' Read the Web Site parameters
  ' ------------------------------------------
  ' Get the Web Site URL
  sURL = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_URL, "")
  txtURL.Text = sURL
  msOriginalURL = Trim(sURL)
  
  ' Get the Web Site login UID and password
  ReadWebLogon sUser, sPassword
  txtUID.Text = sUser
  txtPWD.Text = IIf(Len(sPassword) > 0, Space(20), "")     'Don't show the actual length of the password!
  txtPWD.Tag = sPassword
  msOriginalUser = Trim(sUser)
  msOriginalPassword = sPassword

  ' ------------------------------------------
  ' Read the Personnel Identification parameters
  ' ------------------------------------------
  mlngPersonnelTableID = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_PERSONNELTABLE, 0)
  mlngLoginColumnID = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_LOGINNAME, 0)
  mlngSecondLoginColumnID = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_SECONDLOGINNAME, 0)


  ' --------------------------------------------
  ' Read the Mobile Configuration parameters
  ' --------------------------------------------
  mlngMobPersonnelTableID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_PERSONNELTABLE, 0)
  mlngMobLoginColumnID = GetModuleSetting(gsMODULEKEY_MOBILE, gsPARAMETERKEY_LOGINNAME, 0)
  ' mlngWorkEmailColumnID = GetModuleSetting(gsMODULEKEY_MOBILE, gsPARAMETERKEY_WORKEMAIL, 0)
  mlngWorkEmailColumnID = GetModuleSetting(gsMODULEKEY_PERSONNEL, gsPARAMETERKEY_WORKEMAIL, 0)
  mlngMobLeavingDateColumnID = GetModuleSetting(gsMODULEKEY_MOBILE, gsPARAMETERKEY_LEAVINGDATE, 0)
  mlngMobActivatedColumnID = GetModuleSetting(gsMODULEKEY_MOBILE, gsPARAMETERKEY_MOBILEACTIVATED, 0)

  If (mlngLoginColumnID = 0) _
    And (mlngSecondLoginColumnID <> 0) Then
    
    mlngLoginColumnID = mlngSecondLoginColumnID
    mlngSecondLoginColumnID = 0
  End If
    
  ' Get the Email Address columns
  With recModuleSetup
    .Index = "idxModuleParameter"
    .Seek "=", gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_EMAILCOLUMN
  
    If Not .NoMatch Then
      Do While Not .EOF
        If (!moduleKey <> gsMODULEKEY_WORKFLOW) Or _
          (!parameterkey <> gsPARAMETERKEY_EMAILCOLUMN) Then

          Exit Do
        End If

        lngColumnID = IIf(IsNull(!parametervalue) Or Len(!parametervalue) = 0, 0, CLng(!parametervalue))

        sRow = CStr(lngColumnID) _
          & vbTab & GetColumnName(lngColumnID, True)
        
        With grdEmailAddressColumns
          .AddItem sRow
  
          .Bookmark = .AddItemBookmark(.Rows - 1)
          .SelBookmarks.RemoveAll
          .SelBookmarks.Add .Bookmark
        End With

        .MoveNext
      Loop
    End If
  End With
    
  mlngDelegationActivatedColumnID = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_DELEGATIONACTIVATEDCOLUMN, 0)
  
  mlngDelegateEmailID = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_DELEGATEEMAIL, 0)
  GetEmailAddressDetails
    
  mfCopyDelegateEmail = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_COPYDELEGATEEMAIL, True)
  chkCopyDelegateEmail.value = IIf(mfCopyDelegateEmail, vbChecked, vbUnchecked)

  mfRequiresAuthorization = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_REQUIRESAUTHORIZATION, False)
  chkRequireAuthorization.value = IIf(mfRequiresAuthorization, vbChecked, vbUnchecked)

  ' Get the ServiceSuspended flag
  mfServiceSuspended = GetSystemSetting("workflow", "suspended", "0")
  RefreshServiceSuspended
  
End Sub
Private Sub GetEmailAddressDetails()

  Dim strEmailName As String

  strEmailName = vbNullString

  If mlngDelegateEmailID > 0 Then
    With recEmailAddrEdit
      .Index = "idxID"
      .Seek "=", mlngDelegateEmailID

      ' Read the expression's name from the recordset.
      If Not .NoMatch Then
        strEmailName = !Name
      End If

    End With
  End If

  txtEmail.Text = strEmailName

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


Private Sub grdEmailAddressColumns_DblClick()
  If Not mblnReadOnly Then
    cmdAddEmailAddressColumn_Click
  End If

End Sub

Private Sub grdEmailAddressColumns_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  RefreshControls

End Sub

Private Sub ssTabStrip_Click(PreviousTab As Integer)
  ' Enable, and make visible the selected tab.
  
  If Not mblnReadOnly Then
    fraWebSite.Enabled = (ssTabStrip.Tab = miPAGE_WEBSITE)
    fraWebSiteLogin.Enabled = (ssTabStrip.Tab = miPAGE_WEBSITE)
    fraPersonnelTable.Enabled = (ssTabStrip.Tab = miPAGE_PERSONNELIDENTIFICATION)
    fraDelegation.Enabled = (ssTabStrip.Tab = miPAGE_PERSONNELIDENTIFICATION)
    fraService.Enabled = (ssTabStrip.Tab = miPAGE_SERVICE)
    
    RefreshControls
  End If
  
End Sub

Private Sub txtPWD_Change()
  Changed = True
  txtPWD.Tag = txtPWD.Text
  RefreshControls

End Sub

Private Sub txtPWD_GotFocus()
  With txtPWD
    .SelStart = 0
    .SelLength = Len(.Text)
  End With

End Sub


Private Sub txtUID_Change()
  Changed = True
  RefreshControls

End Sub

Private Sub txtUID_GotFocus()
  With txtUID
    .SelStart = 0
    .SelLength = Len(.Text)
  End With

End Sub


Private Sub txtURL_Change()
  Changed = True
  RefreshControls

End Sub


Private Sub RefreshControls()
  Dim ctlCombo As ComboBox
  
  ' ------------------------------------------
  ' Refresh the Web Site tab controls
  ' ------------------------------------------
  ' Refresh the Web Site frame controls
  txtURL.Enabled = (Not mblnReadOnly)
  txtURL.BackColor = IIf(txtURL.Enabled, vbWindowBackground, vbButtonFace)
  lblURL.Enabled = txtURL.Enabled
    
  txtUID.Enabled = (Not mblnReadOnly)
  txtUID.BackColor = IIf(txtUID.Enabled, vbWindowBackground, vbButtonFace)
  lblUID.Enabled = txtUID.Enabled
    
  txtPWD.Enabled = (Not mblnReadOnly)
  txtPWD.BackColor = IIf(txtPWD.Enabled, vbWindowBackground, vbButtonFace)
  lblPassword.Enabled = txtPWD.Enabled
  
  cmdTestLogon.Enabled = (Len(Trim(txtUID.Text)) > 0)
  
  ' ------------------------------------------
  ' Refresh the Personnel Identification tab controls
  ' ------------------------------------------
  ' Refresh the Personnel Table frame controls
  cboPersonnelTable.Enabled = (cboPersonnelTable.ListCount > 1) And _
    (Not mblnReadOnly) And _
    (Not Application.PersonnelModule)
  cboPersonnelTable.BackColor = IIf(cboPersonnelTable.Enabled, vbWindowBackground, vbButtonFace)
  lblPersonnelTable.Enabled = cboPersonnelTable.Enabled

  For Each ctlCombo In cboLoginName
    ctlCombo.Enabled = (ctlCombo.ListCount > 1) And _
      (Not mblnReadOnly) And _
      (Not Application.PersonnelModule)
    ctlCombo.BackColor = IIf(ctlCombo.Enabled, vbWindowBackground, vbButtonFace)
  Next ctlCombo
  Set ctlCombo = Nothing
  lblLoginNameColumn.Enabled = cboLoginName(0).Enabled
  
  ' Refresh the Email Addresses frame controls
  cmdAddEmailAddressColumn.Enabled = (Not mblnReadOnly) And _
    (cboPersonnelTable.ListCount >= 1) And _
    (cboPersonnelTable.ListIndex > 0)

  With grdEmailAddressColumns
    If .Rows = 0 Then
      cmdRemoveEmailAddressColumn.Enabled = False
      cmdRemoveAllEmailAddressColumns.Enabled = False
    Else
      If .SelBookmarks.Count > 0 Then
        cmdRemoveEmailAddressColumn.Enabled = Not mblnReadOnly
      Else
        cmdRemoveEmailAddressColumn.Enabled = False
      End If
      
      cmdRemoveAllEmailAddressColumns.Enabled = Not mblnReadOnly
    End If
  End With

  ' Refresh the Out of Office Delegation frame controls
  cboDelegationActivatedColumn.Enabled = (cboDelegationActivatedColumn.ListCount > 1) And _
    (Not mblnReadOnly)
  cboDelegationActivatedColumn.BackColor = IIf(cboDelegationActivatedColumn.Enabled, vbWindowBackground, vbButtonFace)
  lblDelegationActivatedColumn.Enabled = cboDelegationActivatedColumn.Enabled
  
  cmdEmail.Enabled = (cboDelegationActivatedColumn.ListCount > 1) And _
    (Not mblnReadOnly)
  lblEmail.Enabled = cmdEmail.Enabled
  
  chkCopyDelegateEmail.Enabled = (Not mblnReadOnly)
  
  ' ------------------------------------------
  ' Refresh the Mobile Key controls
  ' ------------------------------------------
  ' Refresh the Mobile Key controls
  
  ' All columns configured and module activated
  If Application.MobileModule _
    And cboMobEMailColumn.ListIndex > 0 _
    And cboMobLoginName.ListIndex > 0 _
    And cboMobLeavingDateColumn.ListIndex > 0 _
    And cboMobUserActivated.ListIndex > 0 Then
  
    cmdGenMobileKey.Enabled = True
  Else
    cmdGenMobileKey.Enabled = False
  End If
  
  lblGetMobileKey.Enabled = cmdGenMobileKey.Enabled
  
  cboMobPersonnelTable.Clear
  cboMobPersonnelTable.AddItem (cboPersonnelTable)
  cboMobPersonnelTable.ListIndex = 0
  
  
  cboMobLoginName.Enabled = (cboMobLoginName.ListCount > 1) And _
      (Not mblnReadOnly) And Application.MobileModule
    cboMobLoginName.BackColor = IIf(cboMobLoginName.Enabled, vbWindowBackground, vbButtonFace)
  lblMobLoginNameColumn.Enabled = cboMobLoginName.Enabled
  
  cboMobEMailColumn.Enabled = (cboMobEMailColumn.ListCount > 1) And _
    (Not mblnReadOnly) And _
    (Application.MobileModule) And _
    (Not Application.PersonnelModule)
    cboMobEMailColumn.BackColor = IIf(cboMobEMailColumn.Enabled, vbWindowBackground, vbButtonFace)
  lblMobEmailAddresses.Enabled = cboMobEMailColumn.Enabled
  
  cboMobLeavingDateColumn.Enabled = (cboMobLeavingDateColumn.ListCount > 1) And _
    (Not mblnReadOnly) And Application.MobileModule
    cboMobLeavingDateColumn.BackColor = IIf(cboMobLeavingDateColumn.Enabled, vbWindowBackground, vbButtonFace)
  lblMobLeavingDateColumn.Enabled = cboMobLeavingDateColumn.Enabled
  
  cboMobUserActivated.Enabled = (cboMobUserActivated.ListCount > 1) And _
    (Not mblnReadOnly) And Application.MobileModule
    cboMobUserActivated.BackColor = IIf(cboMobUserActivated.Enabled, vbWindowBackground, vbButtonFace)
  lblMobUserActivated.Enabled = cboMobUserActivated.Enabled
    
  ' Disable the OK button as required.
  cmdOK.Enabled = mfChanged
  
End Sub




Private Sub RefreshPersonnelColumnControls()
  ' Refresh the Personnel column controls
  Dim iLoginColumnListIndex As Integer
  Dim iSecondLoginColumnListIndex As Integer
  Dim iDelegationActivatedListIndex As Integer
  
  Dim iMobLoginColumnListIndex As Integer
  Dim iMobEmailColumnListIndex As Integer
  Dim iMobLeavingDateListIndex As Integer
  Dim iMobUserActivatedIndex As Integer

  Dim objctl As Control

  iLoginColumnListIndex = 0
  iSecondLoginColumnListIndex = 0
  iDelegationActivatedListIndex = 0

  UI.LockWindow Me.hWnd
  
  ' Clear the current contents of the combos.
  For Each objctl In Me
    If (TypeOf objctl Is ComboBox) And _
      ((objctl.Name = "cboDelegationActivatedColumn") Or _
      (objctl.Name = "cboLoginName") Or _
      (objctl.Name = "cboMobLoginName") Or _
      (objctl.Name = "cboMobEMailColumn") Or _
      (objctl.Name = "cboMobLeavingDateColumn") Or _
      (objctl.Name = "cboMobUserActivated")) Then

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
          
            If !ColumnID <> mlngSecondLoginColumnID Then
              cboLoginName(0).AddItem !ColumnName
              cboLoginName(0).ItemData(cboLoginName(0).NewIndex) = !ColumnID
              If !ColumnID = mlngLoginColumnID Then
                iLoginColumnListIndex = cboLoginName(0).NewIndex
              End If
            End If
  
            If !ColumnID <> mlngLoginColumnID Then
              cboLoginName(1).AddItem !ColumnName
              cboLoginName(1).ItemData(cboLoginName(1).NewIndex) = !ColumnID
              If !ColumnID = mlngSecondLoginColumnID Then
                iSecondLoginColumnListIndex = cboLoginName(1).NewIndex
              End If
            End If
            
            ' Mobile specifics
            cboMobLoginName.AddItem !ColumnName
            cboMobLoginName.ItemData(cboMobLoginName.NewIndex) = !ColumnID
            If !ColumnID = mlngMobLoginColumnID Then
              iMobLoginColumnListIndex = cboMobLoginName.NewIndex
            End If

            cboMobEMailColumn.AddItem !ColumnName
            cboMobEMailColumn.ItemData(cboMobEMailColumn.NewIndex) = !ColumnID
            If !ColumnID = mlngWorkEmailColumnID Then
              iMobEmailColumnListIndex = cboMobEMailColumn.NewIndex
            End If
            
          End If
          
          If !DataType = dtTIMESTAMP Then
              cboMobLeavingDateColumn.AddItem !ColumnName
              cboMobLeavingDateColumn.ItemData(cboMobLeavingDateColumn.NewIndex) = !ColumnID
              If !ColumnID = mlngMobLeavingDateColumnID Then
                iMobLeavingDateListIndex = cboMobLeavingDateColumn.NewIndex
              End If
          End If
          
                    
          If !DataType = dtBIT Then
            cboDelegationActivatedColumn.AddItem !ColumnName
            cboDelegationActivatedColumn.ItemData(cboDelegationActivatedColumn.NewIndex) = !ColumnID
            
            If !ColumnID = mlngDelegationActivatedColumnID Then
              iDelegationActivatedListIndex = cboDelegationActivatedColumn.NewIndex
            End If
            
            ' Mobile Activated Combo
            cboMobUserActivated.AddItem !ColumnName
            cboMobUserActivated.ItemData(cboMobUserActivated.NewIndex) = !ColumnID
            
            If !ColumnID = mlngMobActivatedColumnID Then
              iMobUserActivatedIndex = cboMobUserActivated.NewIndex
            End If
            
          End If
        
      End If
      .MoveNext
      Loop
    End If
  End With

  ' Select the appropriate combo items.
  cboLoginName(0).ListIndex = iLoginColumnListIndex
  cboLoginName(1).ListIndex = iSecondLoginColumnListIndex
  cboDelegationActivatedColumn.ListIndex = iDelegationActivatedListIndex

  cboMobLoginName.ListIndex = iMobLoginColumnListIndex
  cboMobEMailColumn.ListIndex = iMobEmailColumnListIndex
  cboMobLeavingDateColumn.ListIndex = iMobLeavingDateListIndex
  cboMobUserActivated.ListIndex = iMobUserActivatedIndex

  UI.UnlockWindow

End Sub



Private Sub txtURL_GotFocus()
  UI.txtSelText

End Sub



