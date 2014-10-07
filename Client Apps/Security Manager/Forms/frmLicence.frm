VERSION 5.00
Object = "{AB3877A8-B7B2-11CF-9097-444553540000}#1.0#0"; "gtdate32.ocx"
Begin VB.Form frmLicence 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Licence Information"
   ClientHeight    =   7170
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
   ScaleHeight     =   7170
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCustomer 
      Caption         =   "Customer Details :"
      Height          =   1620
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   8850
      Begin VB.ComboBox cboType 
         Height          =   315
         ItemData        =   "frmLicence.frx":000C
         Left            =   1800
         List            =   "frmLicence.frx":001F
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   705
         Width           =   3120
      End
      Begin VB.TextBox txtCustName 
         Height          =   315
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   1
         Top             =   300
         Width           =   3135
      End
      Begin VB.TextBox txtCustNo 
         Height          =   315
         Left            =   6870
         MaxLength       =   4
         TabIndex        =   2
         Top             =   300
         Width           =   1260
      End
      Begin GTMaskDate.GTMaskDate txtExpiryDate 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   1110
         Width           =   3120
         _Version        =   65537
         _ExtentX        =   5503
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
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
      Begin VB.Label lblExpiryDate 
         Caption         =   "Expiry Date :"
         Height          =   225
         Left            =   195
         TabIndex        =   21
         Top             =   1170
         Width           =   1185
      End
      Begin VB.Label lblModel 
         Caption         =   "Model :"
         Height          =   240
         Left            =   195
         TabIndex        =   20
         Top             =   780
         Width           =   795
      End
      Begin VB.Label lblCustomerName 
         AutoSize        =   -1  'True
         Caption         =   "Customer Name :"
         Height          =   195
         Left            =   195
         TabIndex        =   13
         Top             =   360
         Width           =   1530
      End
      Begin VB.Label lblCustomerNo 
         AutoSize        =   -1  'True
         Caption         =   "Customer No. :"
         Height          =   195
         Left            =   5265
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraLicensedUsers 
      Caption         =   "Licence Details :"
      Height          =   4620
      Left            =   120
      TabIndex        =   15
      Top             =   1845
      Width           =   8850
      Begin VB.TextBox txtHeadcount 
         Height          =   315
         Left            =   3990
         MaxLength       =   6
         TabIndex        =   9
         Top             =   2220
         Width           =   915
      End
      Begin VB.TextBox txtSSI 
         Height          =   315
         Left            =   4000
         MaxLength       =   6
         TabIndex        =   8
         Top             =   1500
         Width           =   900
      End
      Begin VB.ListBox lstModules 
         Height          =   4110
         Left            =   5010
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   315
         Width           =   3690
      End
      Begin VB.TextBox txtDMIS 
         Height          =   315
         Left            =   4000
         MaxLength       =   3
         TabIndex        =   7
         Top             =   1100
         Width           =   900
      End
      Begin VB.TextBox txtDMIM 
         Height          =   315
         Left            =   4000
         MaxLength       =   3
         TabIndex        =   6
         Top             =   700
         Width           =   900
      End
      Begin VB.TextBox txtDAT 
         Height          =   315
         Left            =   4000
         MaxLength       =   3
         TabIndex        =   5
         Top             =   300
         Width           =   900
      End
      Begin VB.Label lblHeadcount 
         Caption         =   "Headcount :"
         Height          =   420
         Left            =   195
         TabIndex        =   22
         Top             =   2280
         Width           =   1065
      End
      Begin VB.Label lblSSI 
         AutoSize        =   -1  'True
         Caption         =   "Self-service Intranet :"
         Height          =   195
         Left            =   195
         TabIndex        =   19
         Top             =   1560
         Width           =   1905
      End
      Begin VB.Label lblDMIS 
         AutoSize        =   -1  'True
         Caption         =   "Data Manager Intranet (Single Record) :"
         Height          =   195
         Left            =   195
         TabIndex        =   18
         Top             =   1155
         Width           =   3450
      End
      Begin VB.Label lblDMIM 
         AutoSize        =   -1  'True
         Caption         =   "Data Manager Intranet (Multiple Records) :"
         Height          =   195
         Left            =   195
         TabIndex        =   17
         Top             =   765
         Width           =   3690
      End
      Begin VB.Label lblDAT 
         AutoSize        =   -1  'True
         Caption         =   "Data Manager :"
         Height          =   195
         Left            =   195
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   7755
      TabIndex        =   12
      Top             =   6600
      Width           =   1200
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   6390
      TabIndex        =   11
      Top             =   6600
      Width           =   1200
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

Public Property Let Changed(ByVal blnNewValue As Boolean)
  cmdApply.Enabled = blnNewValue
End Property

Private Sub cboType_Change()
  Changed = True
End Sub

Private Sub cmdApply_Click()

  Dim objLicence As clsLicence
  Dim lngCount As Long
  Dim lngModules As Long

  'Validate customer number...
  With txtCustNo
    If Len(.Text) <> 4 Or Val(.Text) < 1000 Then
      MsgBox "Invalid Customer Number", vbExclamation
      .SetFocus
      Exit Sub
    End If
  End With

  'Validate number of users...
  If Val(txtDAT.Text) = 0 And Val(txtDMIM.Text) = 0 Then
    MsgBox "Invalid Number of Users", vbExclamation
    txtDAT.SetFocus
    Exit Sub
  End If

  'Check with modules have been selected...
  With lstModules
    lngModules = 0
    For lngCount = 0 To .ListCount - 1
      If .Selected(lngCount) Then
        lngModules = lngModules + .ItemData(lngCount)
      End If
    Next

  End With


  Dim blnCorrectKey As Boolean
  
  
  With frmLicenceKey
    
    blnCorrectKey = False
    Do While Not blnCorrectKey And Not .Cancelled

      .Show vbModal
      If Not .Cancelled Then
    
        'If not cancelled then check they have entered the correct key
        Set objLicence = New clsLicence
        objLicence.LicenceKey = .LicenceKey

        If objLicence.CustomerNo = Val(txtCustNo.Text) And _
           objLicence.DATUsers = Val(txtDAT.Text) And _
           objLicence.DMIMUsers = Val(txtDMIM.Text) And _
           objLicence.DMISUsers = Val(txtDMIS.Text) And _
           objLicence.SSIUsers = Val(txtSSI.Text) And _
           objLicence.Headcount = Val(txtHeadcount.Text) And _
           objLicence.LicenceType = cboType.ListIndex And _
           objLicence.Modules = lngModules Then

              'MH20010910 Fault 2819
              SaveSystemSetting "Licence", "Customer Name", txtCustName.Text
              SaveSystemSetting "Licence", "Customer No", txtCustNo.Text
              SaveSystemSetting "Licence", "Key", .LicenceKey

              blnCorrectKey = True
              MsgBox "Licence details amended successfully", vbExclamation, "Licence Key"
              Unload Me

        Else
          MsgBox "Invalid licence key.", vbExclamation, "Licence Key"

        End If

        Set objLicence = Nothing

      End If
    
    Loop
  
  End With

  Unload frmLicenceKey
  Set frmLicenceKey = Nothing

End Sub

Private Sub cmdCancel_Click()
  Unload Me
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
  
  If Application.AccessMode <> accFull Then
    ControlsDisableAll Me
  End If
  
  PopulateModules
  
  Set objLicence = New clsLicence

  With objLicence
    .LicenceKey = GetSystemSetting("Licence", "Key", vbNullString)

    If .CustomerNo > 0 Then

      txtCustName.Text = GetSystemSetting("Licence", "Customer Name", "")
      cboType.ListIndex = .LicenceType
      txtCustNo.Text = CStr(.CustomerNo)
      txtDAT.Text = CStr(.DATUsers)
      txtDMIM.Text = CStr(.DMIMUsers)
      txtDMIS.Text = CStr(.DMISUsers)
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
      End With
      mfLoading = False

    End If
  End With

  Set objLicence = Nothing

  Changed = False

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
    .AddItem "3rd Party Tables": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    .AddItem "9-Box Grid Reports": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
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
  
  Changed = True
  
End Sub

'Private Sub tabLicences_Click(PreviousTab As Integer)
'  fraCustomer(0).Enabled = (tabLicences.Tab = 0)
'  fraCustomer(1).Enabled = (tabLicences.Tab = 1)
'  fraLicensedUsers.Enabled = (tabLicences.Tab = 0)
'  fraModules.Enabled = (tabLicences.Tab = 1)
'
'End Sub

Private Sub Text1_Change()

End Sub

Private Sub lstModules_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Button = 0
    Shift = 0
End Sub

Private Sub txtCustName_Change()
  Changed = True
End Sub

Private Sub txtCustName_GotFocus()
  With txtCustName
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtCustNo_Change()
  Changed = True
End Sub

Private Sub txtCustNo_GotFocus()
  With txtCustNo
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub txtDMIM_Change()
  Changed = True

End Sub

Private Sub txtDMIM_GotFocus()
  With txtDMIM
    .SelStart = 0
    .SelLength = Len(.Text)
  End With

End Sub


Private Sub txtDMIS_Change()
  Changed = True

End Sub


Private Sub txtDMIS_GotFocus()
  With txtDMIS
    .SelStart = 0
    .SelLength = Len(.Text)
  End With

End Sub

Private Sub txtHeadcount_Change()
  Changed = True
End Sub

Private Sub txtSSI_Change()
  Changed = True
End Sub

Private Sub txtSSI_GotFocus()
  With txtSSI
    .SelStart = 0
    .SelLength = Len(.Text)
  End With

End Sub

Private Sub txtDAT_Change()
  Changed = True
End Sub

Private Sub txtDAT_GotFocus()
  With txtDAT
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub
