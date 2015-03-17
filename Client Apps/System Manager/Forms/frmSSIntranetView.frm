VERSION 5.00
Begin VB.Form frmSSIntranetView 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Self-service Intranet Table (View)"
   ClientHeight    =   7605
   ClientLeft      =   1545
   ClientTop       =   435
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5058
   Icon            =   "frmSSIntranetView.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraWFOutOfOffice 
      Caption         =   "Display options :"
      Height          =   735
      Left            =   200
      TabIndex        =   26
      Top             =   6105
      Width           =   6000
      Begin VB.CheckBox chkWFOutOfOffice 
         Caption         =   "Display &Workflow Out Of Office In ID Badge Dropdown"
         Height          =   255
         Left            =   200
         TabIndex        =   21
         Top             =   300
         Width           =   5500
      End
   End
   Begin VB.ComboBox cboTable 
      Height          =   315
      Left            =   1790
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   4400
   End
   Begin VB.Frame fraLinksLink 
      Caption         =   "Hypertext Link for displaying the links page :"
      Height          =   800
      Left            =   200
      TabIndex        =   4
      Top             =   1485
      Width           =   6000
      Begin VB.TextBox txtLinksLinkText 
         Height          =   315
         Left            =   1590
         MaxLength       =   200
         TabIndex        =   6
         Top             =   300
         Width           =   4215
      End
      Begin VB.Label lblLinksLinkText 
         Caption         =   "Text :"
         Height          =   195
         Left            =   195
         TabIndex        =   5
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Frame fraFindPageLinks 
      Caption         =   "Record selection page :"
      Height          =   3600
      Left            =   200
      TabIndex        =   7
      Top             =   2400
      Width           =   6000
      Begin VB.TextBox txtPageTitle 
         Height          =   315
         Left            =   1590
         MaxLength       =   200
         TabIndex        =   9
         Top             =   300
         Width           =   4215
      End
      Begin VB.TextBox txtDropdownListLinkText 
         Height          =   315
         Left            =   1590
         MaxLength       =   200
         TabIndex        =   20
         Top             =   3100
         Width           =   4215
      End
      Begin VB.CheckBox chkDropdownListLink 
         Caption         =   "&Dropdown List Link"
         Height          =   195
         Left            =   200
         TabIndex        =   18
         Top             =   2760
         Width           =   2100
      End
      Begin VB.CheckBox chkButtonLink 
         Caption         =   "&Button Link"
         Height          =   195
         Left            =   200
         TabIndex        =   13
         Top             =   1560
         Width           =   1500
      End
      Begin VB.TextBox txtButtonLinkPromptText 
         Height          =   315
         Left            =   1590
         MaxLength       =   200
         TabIndex        =   15
         Top             =   1900
         Width           =   4215
      End
      Begin VB.TextBox txtButtonLinkButtonText 
         Height          =   315
         Left            =   1590
         MaxLength       =   200
         TabIndex        =   17
         Top             =   2300
         Width           =   4215
      End
      Begin VB.CheckBox chkHypertextLink 
         Caption         =   "&Hypertext Link"
         Height          =   195
         Left            =   200
         TabIndex        =   10
         Top             =   760
         Width           =   1800
      End
      Begin VB.TextBox txtHypertextLinkText 
         Height          =   315
         Left            =   1590
         MaxLength       =   200
         TabIndex        =   12
         Top             =   1100
         Width           =   4215
      End
      Begin VB.Label lblPageTitle 
         Caption         =   "Page Title :"
         Height          =   195
         Left            =   195
         TabIndex        =   8
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label lblDropdownListLinkText 
         Caption         =   "Text :"
         Height          =   195
         Left            =   405
         TabIndex        =   19
         Top             =   3165
         Width           =   705
      End
      Begin VB.Label lblButtonLinkPromptText 
         Caption         =   "Prompt :"
         Height          =   195
         Left            =   405
         TabIndex        =   14
         Top             =   1965
         Width           =   795
      End
      Begin VB.Label lblButtonLinkButtonText 
         Caption         =   "Button Text :"
         Height          =   195
         Left            =   405
         TabIndex        =   16
         Top             =   2355
         Width           =   1230
      End
      Begin VB.Label lblHypertextLinkText 
         Caption         =   "Text :"
         Height          =   195
         Left            =   405
         TabIndex        =   11
         Top             =   1155
         Width           =   705
      End
   End
   Begin VB.CheckBox chkSingleRecordView 
      Caption         =   "&Single Record View"
      Height          =   255
      Left            =   200
      TabIndex        =   3
      Top             =   1110
      Width           =   2145
   End
   Begin VB.ComboBox cboView 
      Height          =   315
      Left            =   1790
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   675
      Width           =   4400
   End
   Begin VB.Frame fraOKCancel 
      Height          =   400
      Left            =   3600
      TabIndex        =   22
      Top             =   7020
      Width           =   2600
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   400
         Left            =   1400
         TabIndex        =   24
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   400
         Left            =   135
         TabIndex        =   23
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Table : "
      Height          =   195
      Left            =   195
      TabIndex        =   25
      Top             =   300
      Width           =   540
   End
   Begin VB.Label lblView 
      Caption         =   "View :"
      Height          =   195
      Left            =   195
      TabIndex        =   0
      Top             =   735
      Width           =   705
   End
End
Attribute VB_Name = "frmSSIntranetView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnCancelled As Boolean
Private mfChanged As Boolean
Private mlngTableID As Long
Private mlngViewID As Long
'Private mlngPersonnelTableID As Long
Private msSelectedTables As String
Private msSelectedViews As String

Private mblnReadOnly As Boolean

Private mfDefaultPageTitle As Boolean
Private mcolSSITableViews As clsSSITableViews

Public Sub Initialize(plngViewID As Long, _
  plngTableID As Long, _
  psSelectedViews As String, _
  psSelectedTables As String, _
  pfSingleRecordView As Boolean, _
  psButtonLinkPromptText As String, _
  psButtonLinkButtonText As String, _
  psHypertextLinkText As String, _
  psDropdownListLinkText As String, _
  pfButtonLink As Boolean, _
  pfHypertextLink As Boolean, _
  pfDropdownListLink As Boolean, _
  psLinksLinkText As String, _
  psPageTitle As String, _
  pcolSSITableViews As clsSSITableViews, _
  pfWFOutOfOffice As Boolean)
  
  Set mcolSSITableViews = pcolSSITableViews
  
  mlngTableID = plngTableID
  mlngViewID = plngViewID
'  mlngPersonnelTableID = plngPersonnelTableID
  msSelectedTables = psSelectedTables
  msSelectedViews = psSelectedViews
  mfDefaultPageTitle = (plngTableID < 1)
  
  GetTables
  GetViews

  SingleRecordView = pfSingleRecordView
  
  If mlngTableID > 0 Then
    PageTitle = psPageTitle
  End If
  
  HypertextLink = pfHypertextLink
  ButtonLink = pfButtonLink
  DropdownListLink = pfDropdownListLink
  
  HypertextLinkText = psHypertextLinkText
  ButtonLinkPromptText = psButtonLinkPromptText
  ButtonLinkButtonText = psButtonLinkButtonText
  DropdownListLinkText = psDropdownListLinkText
  LinksLinkText = psLinksLinkText
  
  WFOutOfOffice = pfWFOutOfOffice
  
  mfChanged = False
  
  If Not IsModuleEnabled(modWorkflow) Then
    fraWFOutOfOffice.Visible = False
    fraOKCancel.Top = fraOKCancel.Top - fraWFOutOfOffice.Height
    Me.Height = Me.Height - fraWFOutOfOffice.Height
  End If
    
  RefreshControls
  
End Sub

Private Sub RefreshControls()

  Dim fValid As Boolean
  
  chkSingleRecordView.Enabled = (cboView.ListIndex > 0)
  
  If SingleRecordView Then
    HypertextLink = False
    ButtonLink = False
    DropdownListLink = False
  End If
  
  
  fraFindPageLinks.Enabled = Not SingleRecordView
  chkHypertextLink.Enabled = Not SingleRecordView
  chkButtonLink.Enabled = Not SingleRecordView
  chkDropdownListLink.Enabled = Not SingleRecordView
  
  ' Disable the controls as required.
  txtPageTitle.Enabled = Not SingleRecordView
  txtPageTitle.BackColor = IIf(txtPageTitle.Enabled, vbWindowBackground, vbButtonFace)
  lblPageTitle.Enabled = txtPageTitle.Enabled
  If SingleRecordView Then
    txtPageTitle.Text = ""
  End If
  
  txtHypertextLinkText.Enabled = HypertextLink
  txtHypertextLinkText.BackColor = IIf(txtHypertextLinkText.Enabled, vbWindowBackground, vbButtonFace)
  lblHypertextLinkText.Enabled = txtHypertextLinkText.Enabled
  If Not HypertextLink Then
    txtHypertextLinkText.Text = ""
  End If
  
  txtButtonLinkPromptText.Enabled = ButtonLink
  txtButtonLinkPromptText.BackColor = IIf(txtButtonLinkPromptText.Enabled, vbWindowBackground, vbButtonFace)
  lblButtonLinkPromptText.Enabled = txtButtonLinkPromptText.Enabled
  If Not ButtonLink Then
    txtButtonLinkPromptText.Text = ""
  End If
  
  txtButtonLinkButtonText.Enabled = ButtonLink
  txtButtonLinkButtonText.BackColor = IIf(txtButtonLinkButtonText.Enabled, vbWindowBackground, vbButtonFace)
  lblButtonLinkButtonText.Enabled = txtButtonLinkButtonText.Enabled
  If Not ButtonLink Then
    txtButtonLinkButtonText.Text = ""
  End If
  
  txtDropdownListLinkText.Enabled = DropdownListLink
  txtDropdownListLinkText.BackColor = IIf(txtDropdownListLinkText.Enabled, vbWindowBackground, vbButtonFace)
  lblDropdownListLinkText.Enabled = txtDropdownListLinkText.Enabled
  If Not DropdownListLink Then
    txtDropdownListLinkText.Text = ""
  End If
  
  fValid = SingleRecordView Or HypertextLink Or ButtonLink Or DropdownListLink
  
  If Not SingleRecordView _
    And (Len(Trim(txtPageTitle.Text)) = 0) Then
    
    fValid = False
  End If
  
  If HypertextLink _
    And (Len(Trim(txtHypertextLinkText.Text)) = 0) Then
    
    fValid = False
  End If
  
  If ButtonLink _
    And (Len(Trim(txtButtonLinkButtonText.Text)) = 0) Then
      
    fValid = False
  End If
  
  If DropdownListLink _
    And ((Len(Trim(txtDropdownListLinkText.Text)) = 0)) Then
      
    fValid = False
  End If
  
  If ((Len(Trim(txtLinksLinkText.Text)) = 0)) Then
    fValid = False
  End If
  
  ' Disable the OK button as required.
  cmdOk.Enabled = mfChanged And fValid
    
End Sub

Private Sub GetTables()

  ' Populate the views combo.
  Dim sSQL As String
  Dim rsTemp As DAO.Recordset
  Dim iDefaultItem As Integer
  Dim sTableName As String
  
  iDefaultItem = 0

  cboView.Clear

  sSQL = "SELECT tmpTables.tableID, tmpTables.tableName" & _
    " FROM tmpTables" & _
    " WHERE (tmpTables.deleted = FALSE)" & _
    " AND (tmpTables.tableType <> " & iTabChild & ") " & _
    " ORDER BY tmpTables.tableName"
  Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  While Not rsTemp.EOF
    cboTable.AddItem rsTemp!TableName
    cboTable.ItemData(cboTable.NewIndex) = rsTemp!TableID

    If mlngTableID = rsTemp!TableID Then
      iDefaultItem = cboTable.NewIndex
    End If

    rsTemp.MoveNext
  Wend
  rsTemp.Close
  Set rsTemp = Nothing

  If cboTable.ListCount = 0 Then
'    sSQL = "SELECT tmpTables.tableName" & _
'      " FROM tmpTables" & _
'      " WHERE tmpTables.tableID = " & CStr(mlngTableID)
'    sSQL = "SELECT tmpTables.tableID, tmpTables.tableName" & _
'      " FROM tmpTables" & _
'      " WHERE (tmpTables.deleted = FALSE)" & _
'      " AND (tmpTables.tableType <> " & iTabChild & ") " & _
'      " ORDER BY tmpTables.tableName"
'    Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
'
'    If Not (rsTemp.EOF And rsTemp.BOF) Then
'      sTableName = rsTemp!TableName
'    End If
'
'    rsTemp.Close
'    Set rsTemp = Nothing
'
'    If Len(sTableName) = 0 Then
'      MsgBox "All available views have already been selected.", vbOKOnly + vbExclamation, Application.Name
'    Else
'      MsgBox "All '" & sTableName & "' table views have already been selected.", vbOKOnly + vbExclamation, Application.Name
'    End If
'
'    cmdCancel_Click
  Else
    cboTable.ListIndex = iDefaultItem
  End If
  
End Sub

Private Sub GetViews()

  ' Populate the views combo.
  Dim sSQL As String
  Dim rsTemp As DAO.Recordset
  Dim iDefaultItem As Integer
  Dim sTableName As String
  
  iDefaultItem = 0

  cboView.Clear
  cboView.AddItem "<None>"
  cboView.ItemData(cboView.NewIndex) = -1

  If mlngTableID > 0 Then
    sSQL = "SELECT tmpViews.viewID, tmpViews.viewName" & _
      " FROM tmpViews" & _
      " WHERE (tmpViews.deleted = FALSE)" & _
      " AND (tmpViews.viewTableID = " & CStr(cboTable.ItemData(cboTable.ListIndex)) & ")" & _
      " AND ((tmpViews.viewID NOT IN (" & msSelectedViews & "))" & _
      "   OR (tmpViews.viewID = " & CStr(mlngViewID) & "))" & _
      " ORDER BY tmpViews.viewName"
    Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    

    While Not rsTemp.EOF
      cboView.AddItem rsTemp!ViewName
      cboView.ItemData(cboView.NewIndex) = rsTemp!ViewID

      If mlngViewID = rsTemp!ViewID Then
        iDefaultItem = cboView.NewIndex
      End If

      rsTemp.MoveNext
    Wend
    rsTemp.Close
    Set rsTemp = Nothing
  End If

  If cboView.ListCount = 1 Then
    cboView.ListIndex = 0
    cboView.Enabled = False
    cboView.ForeColor = vbGrayText
    cboView.BackColor = vbButtonFace
  Else
    cboView.Enabled = True
    cboView.ForeColor = vbWindowText
    cboView.BackColor = vbWindowBackground
    cboView.ListIndex = iDefaultItem
  End If
  
End Sub

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Public Property Let Cancelled(ByVal bCancel As Boolean)
  mblnCancelled = bCancel
End Property

Public Property Get DropdownListLink() As Boolean
  DropdownListLink = (chkDropdownListLink.value = vbChecked)
End Property

Public Property Get TableViewName() As String

  If cboTable.ListCount > 0 Then
    TableViewName = cboTable.List(cboTable.ListIndex) & IIf(cboView.ListIndex > 0, " (" & cboView.List(cboView.ListIndex) & " view)", vbNullString)
  Else
    TableViewName = ""
  End If
  
End Property

Public Property Get ButtonLinkPromptText() As String
  
  If ButtonLink Then
    ButtonLinkPromptText = txtButtonLinkPromptText.Text
  Else
    ButtonLinkPromptText = ""
  End If
  
End Property

Public Property Get ButtonLinkButtonText() As String
  
  If ButtonLink Then
    ButtonLinkButtonText = txtButtonLinkButtonText.Text
  Else
    ButtonLinkButtonText = ""
  End If
  
End Property

Public Property Get HypertextLinkText() As String

  If HypertextLink Then
    HypertextLinkText = txtHypertextLinkText.Text
  Else
    HypertextLinkText = ""
  End If
  
End Property

Public Property Get PageTitle() As String

  If Not SingleRecordView Then
    PageTitle = txtPageTitle.Text
  Else
    PageTitle = ""
  End If
  
End Property

Public Property Get DropdownListLinkText() As String

  If DropdownListLink Then
    DropdownListLinkText = txtDropdownListLinkText.Text
  Else
    DropdownListLinkText = ""
  End If
  
End Property

Public Property Get LinksLinkText() As String
  LinksLinkText = txtLinksLinkText.Text
End Property

Public Property Get TableID() As Long
  
  If cboTable.ListCount > 0 Then
    TableID = cboTable.ItemData(cboTable.ListIndex)
  Else
    TableID = 0
  End If
  
End Property

Public Property Get ViewID() As Long

  If cboView.ListCount > 0 Then
    ViewID = cboView.ItemData(cboView.ListIndex)
  Else
    ViewID = 0
  End If
  
End Property

Public Property Get ButtonLink() As Boolean
  ButtonLink = (chkButtonLink.value = vbChecked)
End Property

Public Property Get HypertextLink() As Boolean
  HypertextLink = (chkHypertextLink.value = vbChecked)
End Property

Public Property Let DropdownListLink(ByVal pfNewValue As Boolean)
  chkDropdownListLink.value = IIf(pfNewValue, vbChecked, vbUnchecked)
End Property

Public Property Let HypertextLink(ByVal pfNewValue As Boolean)
  chkHypertextLink.value = IIf(pfNewValue, vbChecked, vbUnchecked)
End Property

Public Property Let ButtonLink(ByVal pfNewValue As Boolean)
  chkButtonLink.value = IIf(pfNewValue, vbChecked, vbUnchecked)
End Property

Public Property Get WFOutOfOffice() As Boolean
  WFOutOfOffice = (chkWFOutOfOffice.value = vbChecked)
End Property

Public Property Let WFOutOfOffice(ByVal pfNewValue As Boolean)
  chkWFOutOfOffice.value = IIf(pfNewValue, vbChecked, vbUnchecked)
End Property

Public Property Let HypertextLinkText(ByVal psNewValue As String)

  If HypertextLink Then
    txtHypertextLinkText.Text = psNewValue
  Else
    txtHypertextLinkText.Text = ""
  End If

End Property

Public Property Let PageTitle(ByVal psNewValue As String)

  If Not SingleRecordView Then
    txtPageTitle.Text = psNewValue
  Else
    txtPageTitle.Text = ""
  End If

End Property

Public Property Let LinksLinkText(ByVal psNewValue As String)
  txtLinksLinkText.Text = psNewValue
End Property

Public Property Let ButtonLinkPromptText(ByVal psNewValue As String)

  If ButtonLink Then
    txtButtonLinkPromptText.Text = psNewValue
  Else
    txtButtonLinkPromptText.Text = ""
  End If

End Property

Public Property Let DropdownListLinkText(ByVal psNewValue As String)

  If DropdownListLink Then
    txtDropdownListLinkText.Text = psNewValue
  Else
    txtDropdownListLinkText.Text = ""
  End If

End Property

Public Property Let ButtonLinkButtonText(ByVal psNewValue As String)

  If ButtonLink Then
    txtButtonLinkButtonText.Text = psNewValue
  Else
    txtButtonLinkButtonText.Text = ""
  End If

End Property

Private Sub cboTable_Click()
  
  GetViews
  mfChanged = True
  RefreshControls
  
End Sub

Private Sub cboView_Click()

  If mfDefaultPageTitle Then
    If cboTable.ListCount = 0 Then
      PageTitle = ""
    Else
      PageTitle = "Select the required " & Replace(cboView.Text, "_", " ") _
        & " record"
    End If
  End If
  
  mfChanged = True
  RefreshControls

End Sub

Private Sub chkButtonLink_Click()

  mfChanged = True
  RefreshControls

End Sub

Private Sub chkDropdownListLink_Click()

  mfChanged = True
  RefreshControls

End Sub

Private Sub chkHypertextLink_Click()

  mfChanged = True
  RefreshControls

End Sub

Private Sub chkSingleRecordView_Click()

  If Not chkSingleRecordView.value Then
    PageTitle = "Select the required " & Replace(cboView.Text, "_", " ") _
      & " record"
    mfDefaultPageTitle = True
  End If
  
  mfChanged = True
  RefreshControls

End Sub

Private Sub chkWFOutOfOffice_Click()
  mfChanged = True
  RefreshControls

End Sub

Private Sub cmdCancel_Click()

  Cancelled = True
  UnLoad Me

End Sub

Private Sub cmdOK_Click()

  If ValidateTableView Then
    Cancelled = False
    Me.Hide
  End If

End Sub

Private Function TableViewExistsInCollection(ByVal plngTableID As Long, ByVal plngViewID As Long) As Boolean
  
  Dim oSSITableView As clsSSITableView
  
  If (plngTableID = mlngTableID) And (plngViewID = mlngViewID) Then
    TableViewExistsInCollection = False
    Exit Function
  End If
  
  For Each oSSITableView In mcolSSITableViews.Collection
  
    With oSSITableView
      
      If ((.TableID = plngTableID) And (IIf(.ViewID = 0, -1, .ViewID) = IIf(plngViewID = 0, -1, plngViewID))) Then
          
        TableViewExistsInCollection = True
        Exit Function
      End If
    
    End With
    
  Next oSSITableView
  
  TableViewExistsInCollection = False
  
End Function

Private Function ValidateTableView() As Boolean

  ' Return FALSE if the view definition is invalid.
  Dim fValid As Boolean

  fValid = True
  
  If TableViewExistsInCollection(Me.TableID, Me.ViewID) Then
    fValid = False
    MsgBox "The selected table and view have already been added to the Self-service Intranet Setup." & _
      vbNewLine & vbNewLine & "Please select a different table and view.", vbOKOnly + vbExclamation, Application.Name
    If cboView.Enabled Then
      cboView.SetFocus
    Else
      cboTable.SetFocus
    End If
  End If
  
  If (Not SingleRecordView) _
    And (Not HypertextLink) _
    And (Not ButtonLink) _
    And (Not DropdownListLink) Then
    
    fValid = False
    MsgBox "No link type has been selected.", vbOKOnly + vbExclamation, Application.Name
    chkHypertextLink.SetFocus
  End If
  
  If HypertextLink _
    And (Len(Trim(txtHypertextLinkText.Text)) = 0) Then
    
    fValid = False
    MsgBox "No Hypertext Link text has been entered.", vbOKOnly + vbExclamation, Application.Name
    txtHypertextLinkText.SetFocus
  End If
  
  If Not SingleRecordView _
    And (Len(Trim(txtPageTitle.Text)) = 0) Then
    
    fValid = False
    MsgBox "No page title has been entered.", vbOKOnly + vbExclamation, Application.Name
    txtPageTitle.SetFocus
  End If
  
  If ButtonLink _
    And (Len(Trim(txtButtonLinkButtonText.Text)) = 0) Then
      
    fValid = False
    MsgBox "No Button Link button text has been entered.", vbOKOnly + vbExclamation, Application.Name
    txtButtonLinkButtonText.SetFocus
  End If
  
  If DropdownListLink _
    And ((Len(Trim(txtDropdownListLinkText.Text)) = 0)) Then
      
    fValid = False
    MsgBox "No Dropdown List Link text has been entered.", vbOKOnly + vbExclamation, Application.Name
    txtDropdownListLinkText.SetFocus
  End If

  If ((Len(Trim(txtLinksLinkText.Text)) = 0)) Then
      
    fValid = False
    MsgBox "No Links Link text has been entered.", vbOKOnly + vbExclamation, Application.Name
    txtLinksLinkText.SetFocus
  End If

  ValidateTableView = fValid
  
End Function

Private Sub Form_Initialize()

  mblnReadOnly = (Application.AccessMode <> accFull And _
    Application.AccessMode <> accSupportMode)

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

  ' Position the form in the same place it was last time.
  Me.Top = GetPCSetting(Me.Name, "Top", (Screen.Height - Me.Height) / 2)
  Me.Left = GetPCSetting(Me.Name, "Left", (Screen.Width - Me.Width) / 2)
  
  fraOKCancel.BorderStyle = vbBSNone

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Cancelled = True
  
  If (UnloadMode <> vbFormCode) And mfChanged Then
    Select Case MsgBox("Apply changes ?", vbYesNoCancel + vbQuestion, Me.Caption)
      Case vbCancel
        Cancel = True
      Case vbYes
        cmdOK_Click
        Cancel = True   'MH20021105 Fault 4694
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

Private Sub txtButtonLinkButtonText_Change()

  mfChanged = True
  RefreshControls
  
End Sub

Private Sub txtButtonLinkButtonText_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtButtonLinkPromptText_Change()

  mfChanged = True
  RefreshControls
  
End Sub

Private Sub txtButtonLinkPromptText_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtDropdownListLinkText_Change()

  mfChanged = True
  RefreshControls
  
End Sub

Private Sub txtDropdownListLinkText_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtHypertextLinkText_Change()

  mfChanged = True
  RefreshControls
  
End Sub

Private Sub txtHypertextLinkText_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtLinksLinkText_Change()

  mfChanged = True
  RefreshControls
  
End Sub

Private Sub txtLinksLinkText_GotFocus()
  UI.txtSelText
End Sub

Public Property Get SingleRecordView() As Boolean
  SingleRecordView = (chkSingleRecordView.value = vbChecked)
End Property

Public Property Let SingleRecordView(ByVal pfNewValue As Boolean)
  chkSingleRecordView.value = IIf(pfNewValue, vbChecked, vbUnchecked)
End Property

Private Sub txtPageTitle_Change()

  mfChanged = True
  RefreshControls
  
End Sub

Private Sub txtPageTitle_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtPageTitle_KeyPress(KeyAscii As Integer)
  mfDefaultPageTitle = False
End Sub
