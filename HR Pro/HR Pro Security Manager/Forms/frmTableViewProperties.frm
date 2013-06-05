VERSION 5.00
Begin VB.Form frmTableViewProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Table / View Properties"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1020
   Icon            =   "frmTableViewProperties.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraSelection 
      Caption         =   "Menu / Link Access :"
      Height          =   750
      Left            =   150
      TabIndex        =   14
      Top             =   4365
      Width           =   4750
      Begin VB.CheckBox chkHidefromMenu 
         Caption         =   "Hide from &menu and Self-service Intranet links"
         Height          =   240
         Left            =   180
         TabIndex        =   7
         Top             =   330
         Width           =   4365
      End
   End
   Begin VB.Frame fraRelatedToParents 
      Caption         =   "Child tables with more than one parent :"
      Height          =   1600
      Left            =   150
      TabIndex        =   12
      Top             =   2650
      Width           =   4750
      Begin VB.OptionButton optRelatedToParents 
         Caption         =   "A&ll permitted parents"
         Height          =   315
         Index           =   1
         Left            =   200
         TabIndex        =   6
         Top             =   1200
         Width           =   2160
      End
      Begin VB.OptionButton optRelatedToParents 
         Caption         =   "&Any permitted parent"
         Height          =   315
         Index           =   0
         Left            =   200
         TabIndex        =   5
         Top             =   800
         Width           =   2220
      End
      Begin VB.Label lblChildTables 
         Caption         =   "Users will only have access to records in the child table that are related to ..."
         Height          =   405
         Left            =   195
         TabIndex        =   13
         Top             =   300
         Width           =   4155
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraPermissions 
      Caption         =   "Permissions :"
      Height          =   1950
      Left            =   150
      TabIndex        =   10
      Top             =   600
      Width           =   4750
      Begin VB.CheckBox chkSelect 
         Caption         =   "'&Read' permission"
         Height          =   315
         Left            =   200
         TabIndex        =   3
         Top             =   1100
         Width           =   2100
      End
      Begin VB.CheckBox chkDelete 
         Caption         =   "'&Delete' permission"
         Height          =   315
         Left            =   200
         TabIndex        =   4
         Top             =   1500
         Width           =   2100
      End
      Begin VB.CheckBox chkUpdate 
         Caption         =   "'&Edit' permission"
         Height          =   315
         Left            =   200
         TabIndex        =   2
         Top             =   700
         Width           =   1700
      End
      Begin VB.CheckBox chkInsert 
         Caption         =   "'&New' permission"
         Height          =   315
         Left            =   200
         TabIndex        =   1
         Top             =   300
         Width           =   2100
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3700
      TabIndex        =   9
      Top             =   5580
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2300
      TabIndex        =   8
      Top             =   5580
      Width           =   1200
   End
   Begin VB.TextBox txtTableView 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H80000011&
      Height          =   315
      Left            =   1500
      TabIndex        =   0
      Top             =   200
      Width           =   3405
   End
   Begin VB.Label lblTableView 
      BackStyle       =   0  'Transparent
      Caption         =   "Table / View :"
      Height          =   195
      Left            =   195
      TabIndex        =   11
      Top             =   255
      Width           =   1200
   End
End
Attribute VB_Name = "frmTableViewProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msSecurityGroupName As String
Private msUserGroup As String
Private miLookupTableStatus As Integer

Private mbSelectionEnabled As Boolean
Private mbRelatedEnabled As Boolean

Public Enum LookupTablesStatus
  LOOKUPTABLES_NONE = 0
  LOOKUPTABLES_ALL = 1
  LOOKUPTABLES_SOME = 2
End Enum

' Specific permission items
Private Enum PermissionItems
  PERMISSIONITEM_VIEWLOOKUPS = 133
End Enum

Private mblnReadOnly As Boolean

Private Sub chkDelete_Click()
'Delete Permission
'NHRD02062003 Fault 50, 445
If chkDelete Then
  chkSelect = vbChecked
  chkSelect.Enabled = False
Else
  chkSelect = vbChecked
  chkSelect.Enabled = ((chkUpdate = vbUnchecked))
End If

RefreshControls
End Sub

Private Sub chkInsert_Click()
'New permission
'NHRD02062003 Fault 50, 445
If chkInsert Then
  chkUpdate = vbChecked
  chkUpdate.Enabled = False
  
  chkSelect = vbChecked
  chkSelect.Enabled = False
Else
  chkUpdate = vbChecked
  chkUpdate.Enabled = True
  
  chkSelect = vbChecked
  chkSelect.Enabled = False
End If

RefreshControls

End Sub

Private Sub chkSelect_Click()
'Read Permission
RefreshControls
End Sub

Private Sub chkUpdate_Click()
'Edit Permission
'NHRD02062003 Fault 50, 445
If chkUpdate Then
  chkSelect = vbChecked
  chkSelect.Enabled = False
Else
  chkSelect = vbChecked
  chkSelect.Enabled = (chkDelete = vbUnchecked)
End If

RefreshControls
End Sub

Private Sub RefreshControls()

  ' The OK command control is only enabled if the check boxes are either
  ' checked or unchecked (not greyed).
  cmdOK.Enabled = (chkSelect.Value <> vbGrayed) Or _
    (chkUpdate.Value <> vbGrayed) Or _
    (chkInsert.Value <> vbGrayed) Or _
    (chkDelete.Value <> vbGrayed) Or _
    (chkHidefromMenu.Value <> vbGrayed) Or _
    Not mblnReadOnly
    
End Sub

Private Sub chkHidefromMenu_Click()

  RefreshControls

End Sub

Private Sub cmdCancel_Click()

  Me.Tag = "Cancel"
  Me.Hide

End Sub

Private Sub cmdOK_Click()

  Dim iInvalidCount As Integer
  Dim sMessage As String
  
  iInvalidCount = 0
  sMessage = ""
  
  ' Validate the permissions.
  
  ' Check that 'read' permission is granted if the user is changing permission on lookup tables.
  If (chkSelect.Value = vbUnchecked) And _
    (miLookupTableStatus <> LOOKUPTABLES_NONE) Then
    MsgBox "'Read' permission cannot be revoked for lookup tables.", vbInformation + vbOKOnly, App.Title
    chkSelect.SetFocus
    Exit Sub
  End If
  
  If chkSelect.Value = vbUnchecked Then
    ' If the 'UPDATE' permission is granted, but the 'SELECT' privilege is not,
    ' then inform the user.
    If (chkUpdate.Value = vbChecked) Or _
      (chkUpdate.Value = vbGrayed) Then
      sMessage = "'Edit'"
      iInvalidCount = iInvalidCount + 1
    End If
  
    ' If the 'DELETE' permission is granted, but the 'SELECT' privilege is not,
    ' then inform the user.
    If chkDelete.Value = vbChecked Then
      sMessage = sMessage & IIf(iInvalidCount > 0, ", ", "") & "'Delete'"
      iInvalidCount = iInvalidCount + 1
    End If
  
    ' If the 'INSERT' permission is granted, but the 'SELECT' privilege is not,
    ' then inform the user.
    If chkInsert.Value = vbChecked Then
      sMessage = sMessage & IIf(iInvalidCount > 0, ", ", "") & "'New'"
      iInvalidCount = iInvalidCount + 1
    End If
    
    If iInvalidCount > 0 Then
      sMessage = "'Read' permission must be granted if " & sMessage & _
        " permission" & IIf(iInvalidCount > 1, "s are", " is") & " granted."
    End If
  End If
  
  If chkUpdate.Value = vbUnchecked Then
    ' If the 'INSERT' permission is granted, but the 'UPDATE' privilege is not,
    ' then inform the user.
    If chkInsert.Value = vbChecked Then
      sMessage = sMessage & IIf(Len(sMessage) > 0, vbCrLf, "") & _
        "'Edit' permission must be granted if 'New' permission is granted."
      iInvalidCount = iInvalidCount + 1
    End If
  End If
  
  If (chkUpdate.Value = vbChecked) And _
    (chkSelect.Value = vbGrayed) Then
    ' If the 'UPDATE' permission is fully granted, but the 'SELECT' privilege is only partially
    ' granted then inform the user.
    If chkInsert.Value = vbChecked Then
      sMessage = sMessage & IIf(Len(sMessage) > 0, vbCrLf, "") & _
        "'Read' permission must be fully granted if 'Edit' permission is fully granted."
      iInvalidCount = iInvalidCount + 1
    End If
  End If
  
  If iInvalidCount > 0 Then
    MsgBox sMessage, vbInformation + vbOKOnly, App.Title
    chkSelect.SetFocus
    Exit Sub
  End If
  
  ' Check that all permissions are granted if the user group has permission to
  ' run the System Manager or Security Manager.
  If (chkSelect.Value <> vbChecked) Or _
    (chkUpdate.Value <> vbChecked) Or _
    (chkInsert.Value <> vbChecked) Or _
    (chkDelete.Value <> vbChecked) Then
    
    sMessage = ""
    If gObjGroups(msUserGroup).SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed Then
      sMessage = "System Manager"
    End If
    
    If gObjGroups(msUserGroup).SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGER").Allowed Then
      If Len(sMessage) > 0 Then
        sMessage = "System & Security Managers"
      Else
        sMessage = "Security Manager"
      End If
    End If
  
    If Len(sMessage) > 0 Then
      'JPD 20050208 Fault 9790
      If MsgBox("Permission to run the " & sMessage & " requires the " & _
        "user group to have full access to all tables and views." & vbCrLf & vbCrLf & _
        "Permission to run the " & sMessage & " will be revoked automatically." & vbCrLf & vbCrLf & _
        "Are you sure you want to continue ?", vbYesNo + vbExclamation, App.ProductName) = vbNo Then
                      
        If chkSelect.Enabled Then
          chkSelect.SetFocus
        End If
        
        Exit Sub
      Else
        gObjGroups(msUserGroup).SystemPermissions.Item("P_MODULEACCESS_SYSTEMMANAGER").Allowed = False
        gObjGroups(msUserGroup).SystemPermissions.Item("P_MODULEACCESS_SECURITYMANAGER").Allowed = False
      End If
    End If
  End If
  
  Me.Tag = "OK"
  Me.Hide

End Sub

Public Property Get SelectPermission() As ColumnPrivilegeStates
  Select Case chkSelect.Value
    Case vbChecked
      SelectPermission = giPRIVILEGES_ALLGRANTED
      
    Case vbGrayed
      SelectPermission = giPRIVILEGES_SOMEGRANTED

    Case Else
      SelectPermission = giPRIVILEGES_NONEGRANTED
   End Select
  
End Property

Public Property Let SelectPermission(ByVal piNewValue As ColumnPrivilegeStates)
  Select Case piNewValue
    Case giPRIVILEGES_ALLGRANTED
      chkSelect.Value = vbChecked
      
    Case giPRIVILEGES_SOMEGRANTED
      chkSelect.Value = vbGrayed

    Case Else
      chkSelect.Value = vbUnchecked
   End Select

End Property

Public Property Get InsertPermission() As ColumnPrivilegeStates
  Select Case chkInsert.Value
    Case vbChecked
      InsertPermission = giPRIVILEGES_ALLGRANTED
      
    Case vbGrayed
      InsertPermission = giPRIVILEGES_SOMEGRANTED

    Case Else
      InsertPermission = giPRIVILEGES_NONEGRANTED
   End Select

End Property

Public Property Let InsertPermission(ByVal piNewValue As ColumnPrivilegeStates)
  Select Case piNewValue
    Case giPRIVILEGES_ALLGRANTED
      chkInsert.Value = vbChecked
     
    Case giPRIVILEGES_SOMEGRANTED
      chkInsert.Value = vbGrayed

    Case Else
      chkInsert.Value = vbUnchecked
  End Select

End Property

Public Property Get UpdatePermission() As ColumnPrivilegeStates
  Select Case chkUpdate.Value
    Case vbChecked
      UpdatePermission = giPRIVILEGES_ALLGRANTED
      
    Case vbGrayed
      UpdatePermission = giPRIVILEGES_SOMEGRANTED

    Case Else
      UpdatePermission = giPRIVILEGES_NONEGRANTED
   End Select

End Property

Public Property Let UpdatePermission(ByVal piNewValue As ColumnPrivilegeStates)
  Select Case piNewValue
    Case giPRIVILEGES_ALLGRANTED
      chkUpdate.Value = vbChecked
      
    Case giPRIVILEGES_SOMEGRANTED
      chkUpdate.Value = vbGrayed

    Case Else
      chkUpdate.Value = vbUnchecked
   End Select

End Property

Public Property Get DeletePermission() As ColumnPrivilegeStates
  Select Case chkDelete.Value
    Case vbChecked
      DeletePermission = giPRIVILEGES_ALLGRANTED
      
    Case vbGrayed
      DeletePermission = giPRIVILEGES_SOMEGRANTED

    Case Else
      DeletePermission = giPRIVILEGES_NONEGRANTED
   End Select
  
End Property

Public Property Get MultiParentJoinType() As Integer
  MultiParentJoinType = IIf(optRelatedToParents(0).Value, 0, 1)
  
End Property


Public Property Let DeletePermission(ByVal piNewValue As ColumnPrivilegeStates)
   Select Case piNewValue
   Case giPRIVILEGES_ALLGRANTED
     chkDelete.Value = vbChecked
     
   Case giPRIVILEGES_SOMEGRANTED
     chkDelete.Value = vbGrayed

   Case Else
     chkDelete.Value = vbUnchecked
  End Select

End Property

Public Property Let MultiParentJoinType(ByVal piNewValue As Integer)
  optRelatedToParents(piNewValue).Value = True

End Property


Public Property Let TableViewName(ByVal psTableViewName As String)
  txtTableView.Text = psTableViewName
  
End Property

Public Property Let MultiParentChildren(ByVal pfMultiParentChildren As Boolean)
  If Not pfMultiParentChildren Then
    fraRelatedToParents.Visible = False
    cmdOK.Top = IIf(mbSelectionEnabled, fraSelection.Top + fraSelection.Height + 150, fraPermissions.Top + fraPermissions.Height + 150)
    cmdCancel.Top = cmdOK.Top
  Else
    cmdOK.Top = fraRelatedToParents.Top + fraRelatedToParents.Height + 150
    cmdCancel.Top = cmdOK.Top
  End If
  
  Me.Height = cmdOK.Top + cmdOK.Height + 600
  
End Property


Public Property Let LookupTableStatus(ByVal piStatus As Integer)
  ' Flag to show if the table(s)/views(s) being modified are lookup tables.
  ' We need to know this as 'read' permission is automatically granted on lookup tables.
  ' 0 = No tables/views being modified are lookup tables.
  ' 1 = All tables being edited are lookup tables.
  ' 2 = Some tables/views being edited are lookup tables.
  miLookupTableStatus = piStatus
  
  Select Case miLookupTableStatus
    Case 1
      'frmSelection.Top = fraRelatedToParents.Top
      fraSelection.Visible = True
      fraSelection.Top = IIf(mbRelatedEnabled, fraRelatedToParents.Top + fraRelatedToParents.Height + 150, fraRelatedToParents.Top)
      cmdOK.Top = fraSelection.Top + fraSelection.Height + 150
      cmdCancel.Top = cmdOK.Top
      Me.Height = cmdOK.Top + cmdOK.Height + 600
      mbSelectionEnabled = True
    Case Else
'      cmdOK.Top = fraPermissions.Top + fraPermissions.Height + 150
'      cmdCancel.Top = cmdOK.Top
'      Me.Height = cmdOK.Top + cmdOK.Height + 500
      fraSelection.Visible = False
      mbSelectionEnabled = False
  End Select
  
End Property

Public Property Let UserGroup(ByVal psUserGroupName As String)
  msUserGroup = psUserGroupName
  
End Property



Private Sub Form_Activate()

Dim iCount As Integer

  'NHRD18022003 Fault 2978
  ' Also modified RefreshControls() above replaced AND's with OR's
  
  For iCount = 1 To gObjGroups.Item(msSecurityGroupName).SystemPermissions.Count
    If gObjGroups.Item(msSecurityGroupName).SystemPermissions.Item(iCount).ItemKey = "VIEWLOOKUPTABLES" Then
      fraSelection.Enabled = gObjGroups.Item(msSecurityGroupName).SystemPermissions.Item(iCount).Allowed
      If Not fraSelection.Enabled Then
        chkHidefromMenu.Value = vbGrayed
        chkHidefromMenu.ForeColor = vb3DShadow
      End If
    End If
  Next iCount
  
  cmdOK.Enabled = False
End Sub

Private Sub Form_Load()

  mblnReadOnly = (Application.AccessMode <> accFull)
  If mblnReadOnly Then
    'Disable everything
    ControlsDisableAll Me
  End If
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  If UnloadMode <> vbFormCode Then
    Me.Tag = "Cancel"
    Me.Hide
  End If

End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub optRelatedToParents_Click(Index As Integer)
  'NHRD11072003 Fault 6211
  RefreshControls
End Sub

Public Property Get HideFromMenu() As Integer
  Select Case chkHidefromMenu.Value
    Case vbChecked
      HideFromMenu = giPRIVILEGES_ALLGRANTED
    Case vbGrayed
      HideFromMenu = giPRIVILEGES_SOMEGRANTED
    Case Else
      HideFromMenu = giPRIVILEGES_NONEGRANTED
   End Select
End Property

Public Property Let HideFromMenu(ByVal piNewValue As Integer)
   Select Case piNewValue
   Case giPRIVILEGES_ALLGRANTED
     chkHidefromMenu.Value = vbChecked
   Case giPRIVILEGES_SOMEGRANTED
     chkHidefromMenu.Value = vbGrayed
   Case Else
     chkHidefromMenu.Value = vbUnchecked
  End Select
End Property

Public Property Let SecurityGroupName(ByVal psSecurityGroupName As String)
  msSecurityGroupName = psSecurityGroupName
End Property
