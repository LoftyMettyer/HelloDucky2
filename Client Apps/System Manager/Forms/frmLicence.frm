VERSION 5.00
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Begin VB.Form frmLicence 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Licence Information"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   8019
   Icon            =   "frmLicence.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCustomer 
      Caption         =   "Customer Details :"
      Height          =   1395
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   8865
      Begin VB.CommandButton cmdClipboard 
         Caption         =   "C&opy"
         Height          =   400
         Left            =   8175
         Picture         =   "frmLicence.frx":000C
         TabIndex        =   34
         ToolTipText     =   "Copy to clipboard"
         Top             =   735
         Width           =   570
      End
      Begin VB.TextBox txtLicence 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   6120
         MaxLength       =   6
         TabIndex        =   25
         Top             =   765
         Width           =   840
      End
      Begin VB.TextBox txtLicence 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   24
         Top             =   765
         Width           =   840
      End
      Begin VB.TextBox txtLicence 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   2880
         MaxLength       =   6
         TabIndex        =   23
         Top             =   765
         Width           =   840
      End
      Begin VB.TextBox txtLicence 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   3960
         MaxLength       =   6
         TabIndex        =   22
         Top             =   765
         Width           =   840
      End
      Begin VB.TextBox txtLicence 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   5040
         MaxLength       =   6
         TabIndex        =   21
         Top             =   765
         Width           =   840
      End
      Begin VB.TextBox txtLicence 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Index           =   5
         Left            =   7200
         MaxLength       =   6
         TabIndex        =   20
         Top             =   765
         Width           =   840
      End
      Begin VB.TextBox txtCustName 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   1
         Top             =   300
         Width           =   3135
      End
      Begin VB.TextBox txtCustNo 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   6780
         MaxLength       =   4
         TabIndex        =   2
         Top             =   300
         Width           =   1260
      End
      Begin VB.Label lblLicenceKey 
         Caption         =   "Licence Key :"
         Height          =   255
         Left            =   210
         TabIndex        =   33
         Top             =   810
         Width           =   1230
      End
      Begin VB.Label lblLicence 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   7035
         TabIndex        =   32
         Top             =   795
         Width           =   90
      End
      Begin VB.Label lblLicence 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   5955
         TabIndex        =   31
         Top             =   795
         Width           =   90
      End
      Begin VB.Label lblLicence 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   4860
         TabIndex        =   30
         Top             =   795
         Width           =   90
      End
      Begin VB.Label lblLicence 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   3780
         TabIndex        =   29
         Top             =   795
         Width           =   90
      End
      Begin VB.Label lblLicence 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   2715
         TabIndex        =   28
         Top             =   795
         Width           =   90
      End
      Begin VB.Label lblCustomerName 
         AutoSize        =   -1  'True
         Caption         =   "Customer Name :"
         Height          =   195
         Left            =   195
         TabIndex        =   10
         Top             =   360
         Width           =   1530
      End
      Begin VB.Label lblCustomerNo 
         AutoSize        =   -1  'True
         Caption         =   "Customer No. :"
         Height          =   195
         Left            =   5265
         TabIndex        =   11
         Top             =   360
         Width           =   1320
      End
   End
   Begin VB.Frame fraLicensedUsers 
      Caption         =   "Licence Details :"
      Height          =   5310
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   8865
      Begin VB.ComboBox cboType 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmLicence.frx":0316
         Left            =   1785
         List            =   "frmLicence.frx":0329
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   330
         Width           =   3105
      End
      Begin VB.TextBox txtHeadcount 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3990
         MaxLength       =   6
         TabIndex        =   6
         Top             =   2820
         Width           =   915
      End
      Begin VB.TextBox txtSSI 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   4005
         MaxLength       =   6
         TabIndex        =   5
         Top             =   2100
         Width           =   900
      End
      Begin VB.ListBox lstModules 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   4785
         Left            =   5010
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   315
         Width           =   3690
      End
      Begin VB.TextBox txtDMIM 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   4005
         MaxLength       =   3
         TabIndex        =   4
         Top             =   1695
         Width           =   900
      End
      Begin VB.TextBox txtDAT 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   4005
         MaxLength       =   3
         TabIndex        =   3
         Top             =   1290
         Width           =   900
      End
      Begin GTMaskDate.GTMaskDate txtExpiryDate 
         Height          =   315
         Left            =   1770
         TabIndex        =   17
         Top             =   720
         Width           =   3120
         _Version        =   65537
         _ExtentX        =   5503
         _ExtentY        =   556
         _StockProps     =   77
         Enabled         =   0   'False
         BeginProperty NullFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         BeginProperty CalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CalDayCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblModel 
         Caption         =   "Model :"
         Height          =   240
         Left            =   165
         TabIndex        =   19
         Top             =   390
         Width           =   795
      End
      Begin VB.Label lblExpiryDate 
         Caption         =   "Expiry Date :"
         Height          =   225
         Left            =   165
         TabIndex        =   18
         Top             =   780
         Width           =   1185
      End
      Begin VB.Label lblHeadcount 
         Caption         =   "Headcount :"
         Height          =   420
         Left            =   195
         TabIndex        =   16
         Top             =   2880
         Width           =   1065
      End
      Begin VB.Label lblSSI 
         AutoSize        =   -1  'True
         Caption         =   "Self-service Users :"
         Height          =   195
         Left            =   195
         TabIndex        =   15
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label lblDMIM 
         AutoSize        =   -1  'True
         Caption         =   "OpenHR Web Users :"
         Height          =   195
         Left            =   195
         TabIndex        =   14
         Top             =   1755
         Width           =   1800
      End
      Begin VB.Label lblDAT 
         AutoSize        =   -1  'True
         Caption         =   "Data Manager Users :"
         Height          =   195
         Left            =   195
         TabIndex        =   13
         Top             =   1350
         Width           =   1875
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   7755
      TabIndex        =   9
      Top             =   6990
      Width           =   1200
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Request..."
      Default         =   -1  'True
      Height          =   400
      Left            =   6390
      TabIndex        =   8
      Top             =   6990
      Width           =   1200
   End
   Begin VB.Label lblLicence 
      AutoSize        =   -1  'True
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   7140
      TabIndex        =   27
      Top             =   930
      Width           =   90
   End
   Begin VB.Label lblLicence 
      AutoSize        =   -1  'True
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   6045
      TabIndex        =   26
      Top             =   900
      Width           =   90
   End
End
Attribute VB_Name = "frmLicence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfLoading As Boolean

Private mblnReadOnly As Boolean
Private mstrLicenceKey As String

Private Sub cboType_Click()
 
  If cboType.ListIndex = 0 Then
    txtHeadcount.Text = 0
  Else
    txtSSI.Text = 0
  End If
 
End Sub

Private Sub cmdApply_Click()

  Dim objLicence As clsLicence
  Dim lngCount As Long
  Dim lngModules As Long
  Dim bForceSystemSave As Boolean
  Dim bWorkflowEnabled As Boolean
  Dim blnCorrectKey As Boolean
    
  bWorkflowEnabled = IsModuleEnabled(modWorkflow)
    
  With frmLicenceKey
    
    blnCorrectKey = False
    Do While Not blnCorrectKey And Not .Cancelled

      .Show vbModal
      If Not .Cancelled Then
      
        Screen.MousePointer = vbHourglass
    
        SaveSystemSetting "Licence", "Customer Name", txtCustName.Text
        SaveSystemSetting "Licence", "Customer No", txtCustNo.Text
        SaveSystemSetting "Licence", "Key", .LicenceKey
        
        gobjLicence.LicenceKey = .LicenceKey
        gbLicenceExpired = False
        CheckLicence
        
        LoadShowWhichColumns
        CreateSP_CalculateHeadcount
               
        bForceSystemSave = (bWorkflowEnabled <> IsModuleEnabled(modWorkflow))
        blnCorrectKey = True
        frmSysMgr.RefreshMenu
        
        Screen.MousePointer = vbDefault
       
        MsgBox "Licence details amended successfully", vbExclamation, "Licence Key"
        
        UnLoad Me

      End If
    
    Loop
  
  End With

  If bForceSystemSave Then
    Application.Changed = True
  End If

  UnLoad frmLicenceKey
  Set frmLicenceKey = Nothing

End Sub

Private Sub cmdCancel_Click()
  UnLoad Me
End Sub

Private Sub cmdClipboard_Click()
  Clipboard.Clear
  Clipboard.SetText mstrLicenceKey
End Sub

Private Sub Form_Initialize()

  mblnReadOnly = (Application.AccessMode = accSystemReadOnly)
  
  If mblnReadOnly Then
    ControlsDisableAll Me
  End If

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

  Dim objLicence As clsLicence
  Dim lngCustNo As Long
  Dim lngUsers As Long
  Dim lngModules As Long
  Dim lngCount As Long
  Dim ctlTemp As Control
  
  mstrLicenceKey = GetSystemSetting("Licence", "Key", vbNullString)
    
  PopulateModules
  DisplayLicence (mstrLicenceKey)
  
  Set objLicence = New clsLicence

  With objLicence
    .LicenceKey = mstrLicenceKey

    If .CustomerNo > 0 Then

      txtCustName.Text = GetSystemSetting("Licence", "Customer Name", "")
      cboType.ListIndex = .LicenceType
      txtCustNo.Text = CStr(.CustomerNo)
      txtDAT.Text = CStr(.DATUsers)
      txtDMIM.Text = CStr(.DMIMUsers)
      txtSSI.Text = CStr(.SSIUsers)
      txtHeadcount.Text = CStr(.Headcount)
      
      If Year(.ExpiryDate) > 1900 Then
        txtExpiryDate.DateValue = .ExpiryDate
      End If

      'TM20011218 Fault 3296 - set loading variables.
      mfLoading = True
      With lstModules
        For lngCount = .ListCount - 1 To 0 Step -1
          .Selected(lngCount) = (objLicence.Modules And .ItemData(lngCount))
        Next
        ' Deselect all rows in Module box for TFS 14733
        .ListIndex = -1
      End With
      mfLoading = False

    End If
  End With

  Set objLicence = Nothing

End Sub


Private Sub PopulateModules()

  Dim lngBit As Long
  
  lngBit = 1
  With lstModules
    .AddItem "Personnel": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Recruitment": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Absence": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Training": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Intranet": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "AFD": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Full System Manager": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "CMG": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Quick Address": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Payroll (Shared Table)": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Workflow": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "V1": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Mobile Interface": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Fusion Integration": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "XML Exports": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "OpenLMS Integration": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "9-Box Grid Reports": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Editable Grids": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Power Customisation Pack": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "Talent Management Reports": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    lngBit = lngBit * 2
  End With

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub lstModules_ItemCheck(Item As Integer)

  'TM20011218 Fault 3296
  'Ignore any clicks...If the user has read only access and we are not loading the form.
  If mblnReadOnly Then
    If Not mfLoading Then   'Use loading flag to prevent out of stack space error due to recursion...
      With lstModules
        mfLoading = True
        .Selected(.ListIndex) = Not .Selected(.ListIndex)
        mfLoading = False
      End With
    End If
    Exit Sub
  End If
  
End Sub

Private Sub lstModules_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Button = 0
    Shift = 0
End Sub

Private Sub txtCustName_GotFocus()
  With txtCustName
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtCustNo_GotFocus()
  With txtCustNo
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtDMIM_GotFocus()
  With txtDMIM
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtSSI_GotFocus()
  With txtSSI
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtDAT_GotFocus()
  With txtDAT
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub DisplayLicence(ByVal sLicence As String)
  txtLicence(0).Text = Mid(sLicence, 1, 6)
  txtLicence(1).Text = Mid(sLicence, 8, 6)
  txtLicence(2).Text = Mid(sLicence, 15, 6)
  txtLicence(3).Text = Mid(sLicence, 22, 6)
  txtLicence(4).Text = Mid(sLicence, 29, 6)
  txtLicence(5).Text = Mid(sLicence, 36, 6)
End Sub
