VERSION 5.00
Begin VB.Form frmLicence 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Licence Information"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1019
   Icon            =   "frmLicence.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCustomer 
      Caption         =   "Customer Details :"
      Height          =   1200
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5090
      Begin VB.TextBox txtCustName 
         Height          =   315
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   2
         Top             =   300
         Width           =   3135
      End
      Begin VB.TextBox txtCustNo 
         Height          =   315
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   4
         Top             =   700
         Width           =   1000
      End
      Begin VB.Label lblCustomerName 
         AutoSize        =   -1  'True
         Caption         =   "Customer Name :"
         Height          =   195
         Left            =   195
         TabIndex        =   1
         Top             =   360
         Width           =   1530
      End
      Begin VB.Label lblCustomerNo 
         AutoSize        =   -1  'True
         Caption         =   "Customer No. :"
         Height          =   195
         Left            =   200
         TabIndex        =   3
         Top             =   760
         Width           =   1095
      End
   End
   Begin VB.Frame fraLicensedUsers 
      Caption         =   "Licence Details :"
      Height          =   4005
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   5090
      Begin VB.TextBox txtSSI 
         Height          =   315
         Left            =   4000
         MaxLength       =   3
         TabIndex        =   12
         Top             =   1500
         Width           =   900
      End
      Begin VB.ListBox lstModules 
         Height          =   1860
         Left            =   240
         Style           =   1  'Checkbox
         TabIndex        =   13
         Top             =   1920
         Width           =   4685
      End
      Begin VB.TextBox txtDMIS 
         Height          =   315
         Left            =   4000
         MaxLength       =   3
         TabIndex        =   11
         Top             =   1100
         Width           =   900
      End
      Begin VB.TextBox txtDMIM 
         Height          =   315
         Left            =   4000
         MaxLength       =   3
         TabIndex        =   9
         Top             =   700
         Width           =   900
      End
      Begin VB.TextBox txtDAT 
         Height          =   315
         Left            =   4000
         MaxLength       =   3
         TabIndex        =   7
         Top             =   300
         Width           =   900
      End
      Begin VB.Label lblSSI 
         AutoSize        =   -1  'True
         Caption         =   "Self-service Intranet :"
         Height          =   195
         Left            =   195
         TabIndex        =   16
         Top             =   1560
         Width           =   1905
      End
      Begin VB.Label lblDMIS 
         AutoSize        =   -1  'True
         Caption         =   "Data Manager Intranet (Single Record) :"
         Height          =   195
         Left            =   195
         TabIndex        =   10
         Top             =   1155
         Width           =   3450
      End
      Begin VB.Label lblDMIM 
         AutoSize        =   -1  'True
         Caption         =   "Data Manager Intranet (Multiple Records) :"
         Height          =   195
         Left            =   195
         TabIndex        =   8
         Top             =   765
         Width           =   3690
      End
      Begin VB.Label lblDAT 
         AutoSize        =   -1  'True
         Caption         =   "Data Manager :"
         Height          =   195
         Left            =   195
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4010
      TabIndex        =   15
      Top             =   5580
      Width           =   1200
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   2645
      TabIndex        =   14
      Top             =   5580
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


Private Sub cmdApply_Click()

  Dim objLicence As COALicence.clsLicence2
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

'    If lngModules = 0 Then
'      MsgBox "No Modules selected", vbExclamation
'      .SetFocus
'      Exit Sub
'    End If
  End With


  Dim blnCorrectKey As Boolean
  
  
  With frmLicenceKey
    
    blnCorrectKey = False
    Do While Not blnCorrectKey And Not .Cancelled

      .Show vbModal
      If Not .Cancelled Then
    
        'If not cancelled then check they have entered the correct key
        Set objLicence = New COALicence.clsLicence2
        objLicence.LicenceKey = .LicenceKey

        If objLicence.CustomerNo = Val(txtCustNo.Text) And _
           objLicence.DATUsers = Val(txtDAT.Text) And _
           objLicence.DMIMUsers = Val(txtDMIM.Text) And _
           objLicence.DMISUsers = Val(txtDMIS.Text) And _
           objLicence.SSIUsers = Val(txtSSI.Text) And _
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

Private Sub Form_Load()

  Dim objLicence As COALicence.clsLicence2
  Dim lngCustNo As Long
  Dim lngUsers As Long
  Dim lngModules As Long
  Dim lngCount As Long
  Dim ctlTemp As Control
  
  If Application.AccessMode <> accFull Then
    ControlsDisableAll Me
  End If
  
  PopulateModules
  
  Set objLicence = New COALicence.clsLicence2

  With objLicence
    .LicenceKey = GetLicenceKey

    If .CustomerNo > 0 Then

      txtCustName.Text = GetSystemSetting("Licence", "Customer Name", "")
      txtCustNo.Text = CStr(.CustomerNo)
      txtDAT.Text = CStr(.DATUsers)
      txtDMIM.Text = CStr(.DMIMUsers)
      txtDMIS.Text = CStr(.DMISUsers)
      txtSSI.Text = CStr(.SSIUsers)

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
    .AddItem "Version 1 Integration": .ItemData(.NewIndex) = lngBit: lngBit = lngBit * 2
    
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
