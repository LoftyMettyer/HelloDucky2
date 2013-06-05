VERSION 5.00
Begin VB.Form frmDefaultPermissions2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Apply Permissions & Copy Table Data"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5044
   Icon            =   "frmDefaultPermissions2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCopyData 
      Caption         =   "Copy Table :"
      Height          =   1215
      Left            =   105
      TabIndex        =   8
      Top             =   105
      Width           =   4770
      Begin VB.CheckBox chkCopyData 
         Caption         =   "Copy &table data"
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   750
         Width           =   3375
      End
      Begin VB.Label lblCheck1 
         BackStyle       =   0  'Transparent
         Caption         =   "Copying this table will also copy all child tables, screens, orders and expressions associated with it."
         Height          =   420
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4485
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2400
      TabIndex        =   1
      Top             =   5595
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3660
      TabIndex        =   2
      Top             =   5595
      Width           =   1200
   End
   Begin VB.Frame fraPermissions 
      Caption         =   "Permissions :"
      Height          =   4110
      Left            =   105
      TabIndex        =   0
      Top             =   1395
      Width           =   4770
      Begin VB.ListBox lstPermissions 
         Enabled         =   0   'False
         Height          =   1410
         ItemData        =   "frmDefaultPermissions2.frx":000C
         Left            =   240
         List            =   "frmDefaultPermissions2.frx":001C
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   1965
         Width           =   4200
      End
      Begin VB.OptionButton optNoPermission 
         Caption         =   "&No permission"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Value           =   -1  'True
         Width           =   1995
      End
      Begin VB.OptionButton optCustomisePermissions 
         Caption         =   "C&ustomise permissions :"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   2715
      End
      Begin VB.OptionButton optCopyPermissions 
         Caption         =   "Copy e&xisting permissions"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   2835
      End
      Begin VB.Label lblInstructions2 
         Caption         =   "NOTE: This is the last opportunity to change these defaults."
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3480
         Width           =   4320
      End
      Begin VB.Label lblInstructions1 
         Caption         =   "Please specify the default permissions that will be applied to all user groups for this"
         Height          =   585
         Left            =   120
         TabIndex        =   7
         Top             =   285
         Width           =   4320
      End
   End
End
Attribute VB_Name = "frmDefaultPermissions2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private miChoice As Integer
Private miGrantRead As Boolean
Private miGrantNew As Boolean
Private miGrantEdit As Boolean
Private miGrantDelete As Boolean
Private mblnCopyPermissions As Boolean
Private mblnNewView As Boolean
Private gfCopyData As Boolean

Private mblnUpdatingChk As Boolean
Private mblnPrevChk0 As Boolean
Private mblnPrevChk1 As Boolean
Private mblnPrevChk2 As Boolean
Private mblnPrevChk3 As Boolean

' Returns the selected property value for read permissions
Public Property Get GrantRead() As Integer
  GrantRead = miGrantRead
End Property
' Returns the selected property value for new permissions
Public Property Get GrantNew() As Integer
  GrantNew = miGrantNew
End Property
' Returns the selected property value for edit permissions
Public Property Get GrantEdit() As Integer
  GrantEdit = miGrantEdit
End Property
' Returns the selected property value for delete permissions
Public Property Get GrantDelete() As Integer
  GrantDelete = miGrantDelete
End Property
'Sets the property types
Public Sub SetType(ByVal sNewValue As String, ByVal sTableType As CopySecurityType, Optional objIcon As Object)
  
  mblnNewView = (sNewValue = "new")
  
  If sTableType = giTABLELOOKUP Then
    lblInstructions1.Caption = lblInstructions1.Caption + " lookup table"
    lstPermissions.Selected(2) = True
  End If

  If sTableType = giVIEW Then
    lblInstructions1.Caption = lblInstructions1.Caption + " view"
  End If
  
  If sTableType = giTABLEPARENT Or sTableType = giTABLECHILD Then
    lblInstructions1.Caption = lblInstructions1.Caption + " table"
  End If

  If sNewValue = "copy" And sTableType <> giVIEW Then
    lblInstructions1.Caption = lblInstructions1.Caption + ", its child tables and its associated views"
  End If
   
  ' Set the icon to whatever is passed in.
  If Not IsNull(objIcon) Then
    'NPG20080207 SUGG S000586 no icons now.
    ' Me.Icon = objIcon
  End If


  ' Hide the copy table stuff and resize the form, if referring to a view.
  If sTableType = giVIEW Or sNewValue = "new" Then
    Height = 5460
    Caption = "Apply Permissions"
       
    optNoPermission.Top = 840
    optCopyPermissions.Top = 1200
    optCustomisePermissions.Top = 1560
    
    fraCopyData.Visible = False
    
    fraPermissions.Top = 150
    
    cmdOK.Top = 4440
    cmdCancel.Top = 4440
    
  End If

End Sub

Public Property Get OkCancel() As Integer
  OkCancel = miChoice
End Property

Public Property Get CopyData() As Boolean
  ' Return whether or no the data is to be copied also.
  CopyData = gfCopyData
  
End Property


Public Property Get CopyPermissions() As Boolean
  'CopyPermissions = (optCopyPermissions.Value = True)
  CopyPermissions = mblnCopyPermissions
End Property


Private Sub chkCopyData_Click()
  ' Update the global variable.
  'gfCopyData = chkCopyData.Value
End Sub

'Private Sub chkDelete_Click()
''NHRD25072003 Fault 6274
''Delete Permission
''If chkDelete Then
''  chkSelect = vbChecked
''  chkSelect.Enabled = False
''Else
''  chkSelect = vbChecked
''  chkSelect.Enabled = ((chkUpdate = vbUnchecked))
''End If
'
'RefreshControls
''****************************************
'
''  miGrantDelete = chkDelete.Value
''
''  ' Force other options to be greyed on
''  chkSelect.Value = IIf(chkSelect.Value = vbUnchecked And miGrantDelete = vbChecked, vbGrayed, chkSelect.Value)
''
''  ' Turn other options off
''  chkSelect.Value = IIf(chkSelect.Value = vbGrayed And miGrantDelete = vbUnchecked, vbUnchecked, chkSelect.Value)
'
'End Sub
'Private Sub chkInsert_Click()
''NHRD25072003 Fault 6274
''New permission
'If chkInsert Then
'  chkUpdate = vbChecked
'  chkUpdate.Enabled = False
'
'  chkSelect = vbChecked
'  chkSelect.Enabled = False
'Else
'  chkUpdate = vbChecked
'  chkUpdate.Enabled = True
'
'  chkSelect = vbChecked
'  chkSelect.Enabled = False
'End If
'
'RefreshControls
'
'miGrantNew = chkInsert.Value
'
''  ' Force other options to be greyed on
''  chkUpdate.Value = IIf(chkUpdate.Value = vbUnchecked And miGrantNew = vbChecked, vbGrayed, chkUpdate.Value)
''  chkSelect.Value = IIf(chkSelect.Value = vbUnchecked And miGrantNew = vbChecked, vbGrayed, chkSelect.Value)
''
''  ' Turn other options off
''  chkUpdate.Value = IIf(chkUpdate.Value = vbGrayed And miGrantNew = vbUnchecked, vbUnchecked, chkUpdate.Value)
''  chkSelect.Value = IIf(chkSelect.Value = vbGrayed And miGrantNew = vbUnchecked, vbUnchecked, chkSelect.Value)
''
'End Sub
'Private Sub chkSelect_Click()
'  'Read Permission
'  miGrantRead = chkSelect.Value
'  RefreshControls
'End Sub
'Private Sub chkUpdate_Click()
''NHRD25072003 Fault 6274
''Edit Permission
'If chkUpdate Then
'  chkSelect = vbChecked
'  chkSelect.Enabled = False
'Else
'  chkSelect = vbChecked
'  chkSelect.Enabled = (chkDelete = vbUnchecked)
'End If
'
'miGrantEdit = chkUpdate.Value
'
''  ' Force other options to be greyed on
''  chkSelect.Value = IIf(chkSelect.Value = vbUnchecked And miGrantEdit = vbChecked, vbGrayed, chkSelect.Value)
''
''  ' Turn other options off
''  chkSelect.Value = IIf(chkSelect.Value = vbGrayed And miGrantEdit = vbUnchecked, vbUnchecked, chkSelect.Value)
'
'End Sub

Private Sub cmdCancel_Click()

  miChoice = vbCancel
  UnLoad Me

End Sub

Private Sub cmdOK_Click()

  Dim iInvalidCount As Integer
  Dim sMessage As String
    
  iInvalidCount = 0
  sMessage = ""
  
  ' Validate the permissions.
  If lstPermissions.Selected(2) = vbUnchecked Then
    ' If the 'UPDATE' permission is granted, but the 'SELECT' privilege is not,
    ' then inform the user.
    If (lstPermissions.Selected(1) = vbChecked) Then
      sMessage = "'Edit'"
      iInvalidCount = iInvalidCount + 1
    End If
  
    ' If the 'DELETE' permission is granted, but the 'SELECT' privilege is not,
    ' then inform the user.
    If lstPermissions.Selected(3) = vbChecked Then
      sMessage = sMessage & IIf(iInvalidCount > 0, ", ", "") & "'Delete'"
      iInvalidCount = iInvalidCount + 1
    End If
  
    ' If the 'INSERT' permission is granted, but the 'SELECT' privilege is not,
    ' then inform the user.
    If lstPermissions.Selected(0) = vbChecked Then
      sMessage = sMessage & IIf(iInvalidCount > 0, ", ", "") & "'New'"
      iInvalidCount = iInvalidCount + 1
    End If
    
    If iInvalidCount > 0 Then
      sMessage = "'Read' permission must be granted if " & sMessage & _
        " permission" & IIf(iInvalidCount > 1, "s are", " is") & " granted."
    End If
  End If
' chkInsert - lstPermissions.Selected(0)
' chkUpdate - lstPermissions.Selected(1)
' chkSelect - lstPermissions.Selected(2)
' chkDelete - lstPermissions.Selected(3)
  
  If lstPermissions.Selected(1) = vbUnchecked Then
    ' If the 'INSERT' permission is granted, but the 'UPDATE' privilege is not,
    ' then inform the user.
    If lstPermissions.Selected(0) = vbChecked Then
      sMessage = sMessage & IIf(Len(sMessage) > 0, vbCrLf, "") & _
        "'Edit' permission must be granted if 'New' permission is granted."
      iInvalidCount = iInvalidCount + 1
    End If
    
  End If
  
  ' JDM - 20/07/01 - Fault 2594 - In lookup the read checkbox is disabled.
  If iInvalidCount > 0 Then
    MsgBox sMessage, vbInformation + vbOKOnly, App.Title
    Exit Sub
  End If
  
  miChoice = vbOK
  UnLoad Me
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
  
  miGrantNew = False
  miGrantEdit = False
  miGrantRead = False
  miGrantDelete = False
  
  ' Disable 'Copy Existing' option for all 'New' transactions
  optCopyPermissions.Enabled = Not mblnNewView
  
  RefreshControls

End Sub

Private Sub RefreshControls()
' chkInsert - lstPermissions.Selected(0)
' chkUpdate - lstPermissions.Selected(1)
' chkSelect - lstPermissions.Selected(2)
' chkDelete - lstPermissions.Selected(3)
    
  lstPermissions.Enabled = optCustomisePermissions.value
  
  If optCustomisePermissions.value = False Then
    ' clear ticks from all items
    lstPermissions.Selected(1) = False
    lstPermissions.Selected(2) = False
    lstPermissions.Selected(3) = False
    lstPermissions.Selected(0) = False
  End If
    
End Sub

Private Sub UpdateBooleans()
  miGrantNew = lstPermissions.Selected(0)
  miGrantEdit = lstPermissions.Selected(1)
  miGrantRead = lstPermissions.Selected(2)
  miGrantDelete = lstPermissions.Selected(3)
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub lstPermissions_Click()

  If mblnUpdatingChk Then Exit Sub
  
  mblnUpdatingChk = True
  
  Dim piIndex As Integer
  
  piIndex = lstPermissions.ListIndex
  ' chkInsert - lstPermissions.Selected(0) - New
  ' chkUpdate - lstPermissions.Selected(1) - Edit
  ' chkSelect - lstPermissions.Selected(2) - Read
  ' chkDelete - lstPermissions.Selected(3) - Delete
      
  Select Case piIndex
    Case 0 ' New option selected
      If lstPermissions.Selected(piIndex) = True Then
        ' NPG20081114 Fault 13333
        If mblnPrevChk0 Then
          lstPermissions.Selected(0) = vbUnchecked
        End If
      Else
        If Not mblnPrevChk0 Then
          lstPermissions.Selected(0) = vbChecked
        End If
      End If
        
      lstPermissions.Selected(1) = vbChecked
      lstPermissions.Selected(2) = vbChecked
      
    Case 1 ' Edit option selected
      If lstPermissions.Selected(1) = True Then
        ' NPG20081114 Fault 13333
        If mblnPrevChk1 Then
          lstPermissions.Selected(1) = vbUnchecked
        End If
      Else
        If Not mblnPrevChk1 Then
          lstPermissions.Selected(1) = vbChecked
        End If
      End If
      lstPermissions.Selected(2) = vbChecked
      lstPermissions.Selected(0) = vbUnchecked
      
    Case 2 ' Read option selected
      If lstPermissions.Selected(2) = True Then
        ' NPG20081114 Fault 13333
        If mblnPrevChk2 Then
          lstPermissions.Selected(2) = vbUnchecked
        End If
      Else
        If Not mblnPrevChk2 Then
          lstPermissions.Selected(2) = vbChecked
        End If
     End If
      lstPermissions.Selected(0) = vbUnchecked
      lstPermissions.Selected(1) = vbUnchecked
      lstPermissions.Selected(3) = vbUnchecked
     
    Case 3 ' Delete option selected
      If lstPermissions.Selected(3) = True Then
        ' NPG20081114 Fault 13333
        If mblnPrevChk3 Then
          lstPermissions.Selected(3) = vbUnchecked
        End If
      Else
        If Not mblnPrevChk3 Then
          lstPermissions.Selected(3) = vbChecked
        End If
      End If
      lstPermissions.Selected(2) = vbChecked

  End Select
    
  Call UpdateBooleans
  Call RefreshControls

  mblnPrevChk0 = lstPermissions.Selected(0)
  mblnPrevChk1 = lstPermissions.Selected(1)
  mblnPrevChk2 = lstPermissions.Selected(2)
  mblnPrevChk3 = lstPermissions.Selected(3)
  
  mblnUpdatingChk = False
  
End Sub

Private Sub optCopyPermissions_Click()
  RefreshControls
  mblnCopyPermissions = optCopyPermissions.value
End Sub

Private Sub optCustomisePermissions_Click()
  RefreshControls
  mblnCopyPermissions = optCopyPermissions.value
End Sub

Private Sub optNoPermission_Click()
  RefreshControls
  mblnCopyPermissions = optCopyPermissions.value
End Sub



