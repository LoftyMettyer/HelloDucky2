VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNewTrustedUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Windows Login(s)"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7110
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   8037
   Icon            =   "frmNewTrustedUser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imlSmallIcons 
      Left            =   3180
      Top             =   3525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewTrustedUser.frx":000C
            Key             =   "TRUSTEDUSER"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewTrustedUser.frx":03F2
            Key             =   "TRUSTEDGROUP"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3030
      Left            =   105
      TabIndex        =   0
      Top             =   450
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   5345
      SortKey         =   3
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User Name"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "LoginType"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "SortColumn"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CheckBox chkShowMembers 
      Caption         =   "&Show Members"
      Height          =   225
      Left            =   150
      TabIndex        =   1
      Top             =   3630
      Width           =   2025
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   420
      Left            =   4515
      TabIndex        =   2
      Top             =   3555
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   420
      Left            =   5790
      TabIndex        =   3
      Top             =   3555
      Width           =   1200
   End
   Begin VB.Label lblSelectUsers 
      Caption         =   "Select users to add :"
      Height          =   270
      Left            =   135
      TabIndex        =   4
      Top             =   90
      Width           =   1890
   End
End
Attribute VB_Name = "frmNewTrustedUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mbCancelled As Boolean
Dim mstrDomainName As String
Dim mstrUsersSelected As String
Dim mstrUsersSelectedTypes As String
Dim mastrDomainUsers() As String
Dim mastrDomainGroups() As String

Dim mlngMinHeight As Long
Dim mlngMinWidth As Long

Public Property Get Cancelled() As Boolean
  Cancelled = mbCancelled
End Property

Public Property Let Cancelled(ByVal pbNewValue As Boolean)
  mbCancelled = pbNewValue
End Property

Public Property Let DomainName(ByVal pstrData As String)
  mstrDomainName = pstrData
End Property

Public Property Get UsersSelected() As String
  UsersSelected = mstrUsersSelected
End Property

Public Property Get UsersSelectedTypes() As String
  UsersSelectedTypes = mstrUsersSelectedTypes
End Property

Private Function PopulateUsersFromDomain() As Boolean

  On Error GoTo ErrorTrap

  Dim rsRecords As New ADODB.Recordset
  Dim iCount As Integer
  Dim strSQL As String
  Dim objNet As New Net
  Dim astrUserList() As String
  Dim iKey As Integer
  Dim bOK As Boolean

  bOK = False
  With gobjProgress
    '.AviFile = ""
    .AVI = dbLoadDomain
    .NumberOfBars = 1
    .Bar1MaxValue = 2
    .Caption = "Loading domain information, please wait..."
    .MainCaption = "Reading Domain"
    .Cancel = False
    .Time = False
    .HidePercentages = False
    .OpenProgress
  End With

  
  mastrDomainGroups = InitialiseWindowsGroups(mstrDomainName)
  bOK = (UBound(mastrDomainGroups, 2) > 0)

  ' Get the users
  If bOK Then
    mastrDomainUsers = InitialiseWindowsUsers(mstrDomainName)
  End If
  
  ' AE20080917 Fault #13372
  If glngSQLVersion > 8 Then
    mstrDomainName = GetDomainFromUser(mstrDomainName, mastrDomainGroups(0, 0))
  End If

'    ReDim mastrDomainUsers(1, 0)
'    objNet.GetUserList mstrDomainName, astrUserList
'
'    For iCount = LBound(astrUserList) To UBound(astrUserList)
'      ReDim Preserve mastrDomainUsers(1, iCount)
'      mastrDomainUsers(0, iCount) = astrUserList(iCount)
'      mastrDomainUsers(1, iCount) = ""
'    Next iCount
  'End If

TidyUpAndExit:
  gobjProgress.CloseProgress
  Set rsRecords = Nothing
  PopulateUsersFromDomain = bOK
  Exit Function

ErrorTrap:
  bOK = False
  GoTo TidyUpAndExit

End Function

Private Sub chkShowMembers_Click()

  RefreshGrid

End Sub

Private Sub cmdCancel_Click()

  mbCancelled = True
  Unload Me

End Sub

Private Sub cmdOK_Click()

  Dim iCount As Integer
  Dim strUsersSelected As String
  Dim strUsersSelectedTypes As String
  
  strUsersSelected = ""
  strUsersSelectedTypes = ""
  For iCount = 1 To lvList.ListItems.Count
    If lvList.ListItems(iCount).Selected Then
      strUsersSelected = IIf(Len(strUsersSelected) > 0, strUsersSelected & ";", "") & mstrDomainName & "\" & lvList.ListItems(iCount).Text
      strUsersSelectedTypes = IIf(Len(strUsersSelectedTypes) > 0, strUsersSelectedTypes & ";", "") & Str(lvList.ListItems(iCount).SubItems(2))
    End If
  Next iCount

  mstrUsersSelected = strUsersSelected
  mstrUsersSelectedTypes = strUsersSelectedTypes
  mbCancelled = False
  Unload Me
  
End Sub

Private Sub RefreshGrid()

  Dim objThisItem As Object ' ComctlLib.ListItem
  Dim iCount As Integer
  Dim iBaseKey As Integer
  
  With gobjProgress
    '.AviFile = ""
    .AVI = dbLoadDomain
    .NumberOfBars = 1
    .Bar1MaxValue = 2
    .Caption = "Loading domain information, please wait..."
    .MainCaption = "Reading Domain"
    .Cancel = False
    .Time = False
    .HidePercentages = False
    .OpenProgress
  End With
  
  ' Clear the existing data
  lvList.ListItems.Clear
  lvList.Sorted = False
  
  ' Add the groups to the grid
  For iCount = LBound(mastrDomainGroups, 2) To UBound(mastrDomainGroups, 2)
    ' AE20080917 Fault #13372
    mastrDomainGroups(0, iCount) = Replace(mastrDomainGroups(0, iCount), mstrDomainName & "\", "")
    Set objThisItem = lvList.ListItems.Add(, "key" & Str(iCount), mastrDomainGroups(0, iCount), "TRUSTEDGROUP", "TRUSTEDGROUP")
    objThisItem.SubItems(1) = mastrDomainGroups(1, iCount)
    objThisItem.SubItems(2) = iUSERTYPE_TRUSTEDGROUP
    objThisItem.SubItems(3) = "aaaa" & mastrDomainGroups(0, iCount)   ' Used for sorting
  Next iCount

' AE20080219 Fault #12892
'  ' Add the users
'  iBaseKey = UBound(mastrDomainGroups, 2) + 1
'  If chkShowMembers.Value = vbChecked Then
'    For icount = LBound(mastrDomainUsers) + 1 To UBound(mastrDomainUsers)
'      Set objThisItem = lvList.ListItems.Add(, "key" & Str(iBaseKey + icount), mastrDomainUsers(icount), "TRUSTEDUSER", "TRUSTEDUSER")
'      objThisItem.SubItems(1) = ""     ' FullName
'      objThisItem.SubItems(2) = iUSERTYPE_TRUSTEDUSER
'      objThisItem.SubItems(3) = "zzzz" & mastrDomainUsers(icount)   ' Used for sorting
'    Next icount
'  End If

  iBaseKey = UBound(mastrDomainGroups, 2) + 1
  If chkShowMembers.Value = vbChecked Then
    For iCount = LBound(mastrDomainUsers) To UBound(mastrDomainUsers)
      ' AE20080917 Fault #13372
      mastrDomainUsers(iCount) = Replace(mastrDomainUsers(iCount), mstrDomainName & "\", "")
      Set objThisItem = lvList.ListItems.Add(, "key" & Str(iBaseKey + iCount), mastrDomainUsers(iCount), "TRUSTEDUSER", "TRUSTEDUSER")
      objThisItem.SubItems(1) = ""     ' FullName
      objThisItem.SubItems(2) = iUSERTYPE_TRUSTEDUSER
      objThisItem.SubItems(3) = "zzzz" & mastrDomainUsers(iCount)   ' Used for sorting
    Next iCount
  End If
  
  lvList.Sorted = True

  gobjProgress.CloseProgress

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_Resize()

  'JPD 20030908 Fault 5756
  DisplayApplication

  lvList.Width = Me.Width - 300
  lvList.Height = Me.Height - 1700
  
  lvList.ColumnHeaders(2).Width = lvList.Width - (lvList.ColumnHeaders(1).Width + 325)
  
  chkShowMembers.Top = Me.Height - 1100
  cmdOK.Top = chkShowMembers.Top
  cmdOK.Left = Me.Width - ((cmdOK.Width + cmdCancel.Width) + 250)
  cmdCancel.Top = chkShowMembers.Top
  cmdCancel.Left = cmdOK.Left + cmdOK.Width + 50

End Sub

' Do the same as the OK button
Private Sub lvList_DblClick()
  cmdOK_Click
End Sub

Public Function Initialise() As Boolean

  ' Set smallest size
  mlngMinHeight = Me.Height
  mlngMinWidth = Me.Width

  ' Attach imagelists to listview
  lvList.SmallIcons = imlSmallIcons
  lvList.Icons = imlSmallIcons

  ' Get specified users
  If PopulateUsersFromDomain Then
    RefreshGrid
    Initialise = True
  Else
    Initialise = False
  End If

End Function
