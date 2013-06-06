VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNewMultipleUserReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Automatic User Add Report"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1030
   Icon            =   "frmNewMultipleUserReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   6720
      TabIndex        =   2
      Top             =   4350
      Width           =   1155
   End
   Begin VB.Frame frmDetails 
      Caption         =   "Details :"
      Height          =   705
      Left            =   105
      TabIndex        =   4
      Top             =   90
      Width           =   7770
      Begin VB.TextBox txtSecurityGroupName 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1320
         TabIndex        =   5
         Top             =   255
         Width           =   2700
      End
      Begin VB.Label lblSecurityGroup 
         Caption         =   "User Group :"
         Height          =   225
         Left            =   135
         TabIndex        =   6
         Top             =   315
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   405
      Left            =   105
      TabIndex        =   3
      Top             =   4350
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   405
      Left            =   5520
      TabIndex        =   1
      Top             =   4350
      Width           =   1155
   End
   Begin VB.Frame frmReport 
      Caption         =   "Users Added :"
      Height          =   3345
      Left            =   90
      TabIndex        =   0
      Top             =   915
      Width           =   7785
      Begin SSDataWidgets_B.SSDBGrid grdAddedUsers 
         Height          =   2925
         Left            =   150
         TabIndex        =   7
         Top             =   270
         Width           =   7515
         ScrollBars      =   0
         _Version        =   196617
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   5
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
         SelectTypeRow   =   0
         BalloonHelp     =   0   'False
         RowNavigation   =   1
         MaxSelectedRows =   0
         ForeColorEven   =   -2147483640
         ForeColorOdd    =   -2147483640
         BackColorEven   =   -2147483643
         BackColorOdd    =   -2147483643
         RowHeight       =   423
         Columns.Count   =   5
         Columns(0).Width=   3149
         Columns(0).Caption=   "User"
         Columns(0).Name =   "User"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2990
         Columns(1).Caption=   "User Name"
         Columns(1).Name =   "User Name"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   2646
         Columns(2).Caption=   "Password"
         Columns(2).Name =   "Password"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   4604
         Columns(3).Caption=   "Status"
         Columns(3).Name =   "Status"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   3200
         Columns(4).Caption=   "ICON"
         Columns(4).Name =   "ICON"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   13256
         _ExtentY        =   5159
         _StockProps     =   79
         BackColor       =   16777215
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   1380
      Top             =   4275
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewMultipleUserReport.frx":000C
            Key             =   "DATABASE"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewMultipleUserReport.frx":04E7
            Key             =   "STATUS_ALERT"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewMultipleUserReport.frx":0A85
            Key             =   "MODE"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewMultipleUserReport.frx":0DD7
            Key             =   "SERVER"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewMultipleUserReport.frx":1129
            Key             =   "STATUS_OK"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewMultipleUserReport.frx":1802
            Key             =   "GROUP"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewMultipleUserReport.frx":1E83
            Key             =   "STATUS_ERROR"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmNewMultipleUserReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mstrSecurityGroupName As String
Private mbCancelled As Boolean
Private mbUsersCreated As Boolean
Private miCreateMode As HrProSecurityMgr.CreateUserMode

Public Property Get Cancelled() As Boolean
  Cancelled = mbCancelled
End Property

Public Property Let SecurityGroupName(ByVal pstrNewValue As String)
  mstrSecurityGroupName = pstrNewValue
  txtSecurityGroupName.Text = mstrSecurityGroupName
End Property

Public Property Let ReportData(ByVal avUserStatus As Variant)

Dim iCount As Integer
Dim strStatus As String
Dim strPassword As String
Dim strStatusType As String

grdAddedUsers.RemoveAll

For iCount = 0 To UBound(avUserStatus, 1) - 1

  ' Add the report list
  Select Case avUserStatus(iCount, 3)
    Case iSUCCESS
      strStatus = IIf(miCreateMode = iUSERCREATE_SQLLOGIN, "SQL Login Created", "Windows Login Added")
      strPassword = avUserStatus(iCount, 2)
      mbUsersCreated = True
      strStatusType = "STATUS_OK"
    Case iFAILED_USERNAMEISBLANK
      strStatus = "Blank Username specified"
      strPassword = ""
      strStatusType = "STATUS_ERROR"
    Case iFAILED_USERNAMEISRESERVED
      strStatus = "Reserved Username"
      strPassword = ""
      strStatusType = "STATUS_ERROR"
    Case iFAILED_USERNAMEISKEYWORD
      strStatus = "Username is keyword"
      strPassword = ""
      strStatusType = "STATUS_ERROR"
    Case iFAILED_ILLEGALPASSWORD
      strStatus = "Password is illegal"
      strPassword = avUserStatus(iCount, 2)
      strStatusType = "STATUS_ERROR"
    Case iWARNING_LOGINEXISTS
      strStatus = "Login attached"
      strPassword = "***************"
      strStatusType = "STATUS_ALERT"
    Case iFAILED_USERNAMEISUSED
      strStatus = "User already has a login"
      strPassword = "***************"
      strStatusType = "STATUS_ALERT"
    Case iFAILED_USERNAMEISNUMERIC
      strStatus = "Username is numeric"
      strPassword = ""
      strStatusType = "STATUS_ERROR"
    'TM20020122 Fault 3350
    Case iFAILED_PASSWORDNOTMINIMUM
      strStatus = "Password less than minimum length"
      strPassword = ""
      strStatusType = "STATUS_ERROR"
    'TM20020122 Fault 3374
    Case iFAILED_USERNAMEGREATERTHANLOGINSIZE
      strStatus = "Username length greater than Login field length"
      strPassword = ""
      strStatusType = "STATUS_ERROR"
    'TM20020122 Fault 3229
    Case iFAILED_USERNAMEISTOOLONG
      strStatus = "Username length greater than max length of " & CStr(giMAXIMUMUSERNAMELENGTH)
      strPassword = ""
      strStatusType = "STATUS_ERROR"
    Case iFAILED_NTACCOUNTNOTEXIST
      strStatus = "Windows account does not exist"
      strPassword = ""
      strStatusType = "STATUS_ERROR"
    Case iFAILED_PASSWORDNOTCOMPLEX
      strStatus = "Password not complex enough"
      strPassword = "***************"
      strStatusType = "STATUS_ERROR"
    Case iSUCCESS_USERALREADYADDED
      strStatus = "Login field already populated"
      strPassword = "***************"
      strStatusType = "STATUS_ALERT"
    
  End Select
  
    '  grdAddedUsers.AddItem avUserStatus(iCount, 4) & vbTab _
    '    & avUserStatus(iCount, 1) & vbTab _
    '    & strPassword & vbTab _
    '    & strStatus

  'NHRD07032003 Fault 3508
  'Shuffled the order to make Username the first column
  'Also changed column order in the grdAddedUsers
  'otherwise column headers would not match.
  grdAddedUsers.AddItem avUserStatus(iCount, 4) & vbTab _
    & avUserStatus(iCount, 1) & vbTab _
    & strPassword & vbTab _
    & strStatus & vbTab _
    & strStatusType
    
Next iCount

End Property

Private Sub cmdCancel_Click()

  mbCancelled = True
  Unload Me

End Sub

Private Sub cmdOK_Click()

  mbCancelled = False
  Unload Me

End Sub

Private Sub PrintGrid()

  Dim strUserName As String
  Dim strUserDescription As String
  Dim strPassword As String
  Dim strStatus As String
  Dim objPrinter As New HrProSecurityMgr.clsPrintDef
  Dim objIcon As IPictureDisp

  On Error GoTo ErrorTrap

  Screen.MousePointer = vbHourglass

  With objPrinter
    If .IsOK Then
      If .PrintStart(True) Then
        .PrintHeader "Automatic User Add Report"
     
        Set objIcon = imlIcons.ListImages("GROUP").Picture
        .PrintNormal "User Group: " & mstrSecurityGroupName, objIcon
        
        Set objIcon = imlIcons.ListImages("DATABASE").Picture
        .PrintNormal "Database: " & gsDatabaseName, objIcon
        
        Set objIcon = imlIcons.ListImages("SERVER").Picture
        .PrintNormal "Server: " & gsServerName, objIcon
        
        Set objIcon = imlIcons.ListImages("MODE").Picture
        If miCreateMode = iUSERCREATE_SQLLOGIN Then
          .PrintNormal "Mode: SQL Server Authentication", objIcon
        Else
          .PrintNormal "Mode: Windows Domain Accounts", objIcon
        End If
        
        .PrintNormal ""
        .TabsOnPage = 10
        .PrintBold "User Name" & vbTab & vbTab & vbTab & "User" & vbTab & vbTab & "Password" & vbTab & vbTab & "Status"
  
        .FontSize = 8
        
        grdAddedUsers.MoveFirst
        For iCount = 1 To grdAddedUsers.Rows
          strUserName = grdAddedUsers.Columns("User Name").Text
          strUserDescription = grdAddedUsers.Columns("User").Text
          strPassword = grdAddedUsers.Columns("Password").Text
          strStatus = grdAddedUsers.Columns("Status").Text
          Set objIcon = imlIcons.ListImages(grdAddedUsers.Columns("ICON").Text).Picture
    
          .PrintNonBold strUserName & vbTab & vbTab & vbTab & strUserDescription & vbTab & vbTab & strPassword & vbTab & vbTab & strStatus, objIcon
          grdAddedUsers.MoveNext
        Next iCount
        
        .PrintEnd
  
      End If
    End If
  End With

  Screen.MousePointer = vbDefault

  Exit Sub
ErrorTrap:
  MsgBox ("Print Error")

End Sub

Private Sub cmdPrint_Click()

  PrintGrid

End Sub

Private Sub Form_Activate()

  ' JDM - 17/12/01 - Fault 3198 - Dont user automatic scrollbars
  DoEvents
  With grdAddedUsers
    If .VisibleRows < .Rows Then
      'NHRD08072003 Fault 3827
      .ScrollBars = ssScrollBarsAutomatic
    Else
      .ScrollBars = ssScrollBarsNone
      .Columns("Status").Width = .Columns("Status").Width + 240
    End If
  
    .Columns("Icon").Width = 0
  End With
  
End Sub

Private Sub Form_Load()

  mbCancelled = True
  mbUsersCreated = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Dim iResponse As Integer

  If mbCancelled And mbUsersCreated Then
    iResponse = MsgBox("Do you want to add these users ?", vbYesNoCancel + vbQuestion, "Automatic Add")
  
    Select Case iResponse
      Case vbYes
        mbCancelled = False
     
      Case vbCancel
        Cancel = 1
        Exit Sub
    
    End Select
  
  End If

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


Private Sub grdAddedUsers_ColResize(ByVal ColIndex As Integer, Cancel As Integer)

  If ColIndex > 3 Then
    Cancel = 1
  End If

End Sub

Private Sub grdAddedUsers_PrintError(ByVal PrintError As Long, Response As Integer)
'NHRD10032004 Fault 6348 Included to alter default message 30457
  If PrintError = 30457 Then '"cancelled by user" error
    MsgBox "Print job cancelled by user. ", vbOKOnly + vbExclamation, "Print Cancelled"
    'Set to 0 to prevent a default error message from being displayed
    Response = 0
  End If
  
End Sub

Public Property Let CreateMode(ByVal piNewValue As HrProSecurityMgr.CreateUserMode)
  miCreateMode = piNewValue
End Property
