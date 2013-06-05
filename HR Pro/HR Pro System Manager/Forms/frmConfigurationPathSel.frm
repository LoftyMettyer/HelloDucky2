VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfigurationPathSel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Email Attachment Path"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1010
   Icon            =   "frmConfigurationPathSel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin MSComctlLib.TreeView trvAttachmentDir 
         Height          =   3975
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   7011
         _Version        =   393217
         Indentation     =   556
         LabelEdit       =   1
         Style           =   7
         SingleSel       =   -1  'True
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3135
      TabIndex        =   3
      Top             =   4560
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   1800
      TabIndex        =   2
      Top             =   4560
      Width           =   1200
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfigurationPathSel.frx":000C
            Key             =   "IMG_OPENFOLDER"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfigurationPathSel.frx":03D9
            Key             =   "IMG_CLOSEDFOLDER"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfigurationPathSel.frx":07C5
            Key             =   "IMG_SERVER"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConfigurationPathSel.frx":0B8C
            Key             =   "IMG_HARDDRIVE"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConfigurationPathSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrPath As String
Private mblnCancelled As String

Private Sub cmdCancel_Click()
  UnLoad Me
End Sub

Private Sub cmdOK_Click()

  If Trim(mstrPath) <> vbNullString Then
    mblnCancelled = False
    'frmConfiguration.txtAttachmentsPath.Text = mstrPath
    UnLoad Me
  End If

End Sub

Public Function Initialise() As Boolean

  Screen.MousePointer = vbHourglass
  Initialise = trvAttachmentDir_Populate
  Screen.MousePointer = vbDefault
  mblnCancelled = True

End Function


Private Function trvAttachmentDir_Populate() As Boolean

  Dim rsTemp As New ADODB.Recordset
  Dim strSQL As String
  Dim objCurrentNode As MSComctlLib.Node
  Dim lngCount As Long
  Dim objServerNode As MSComctlLib.Node
  
  On Error GoTo LocalErr
  
  Set objServerNode = trvAttachmentDir.Nodes.Add(, , , UCase(Database.ServerName), "IMG_SERVER", "IMG_SERVER")
  objServerNode.Expanded = True

  'JDM - 10/12/2004 - Faults 2936/6253/9226 - Limit selection to just local drives
  'strSQL = "exec sp_ASRServerDir 1, ''"
  strSQL = "master..xp_fixeddrives"
  
  rsTemp.Open strSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdStoredProc

  Do While Not rsTemp.EOF
    Set objCurrentNode = trvAttachmentDir.Nodes.Add(objServerNode, ssatChild, rsTemp.Fields(0).value & ":\", rsTemp.Fields(0).value & ":", "IMG_HARDDRIVE", "IMG_HARDDRIVE")
    PopulateFolder objCurrentNode
    rsTemp.MoveNext
  Loop
  rsTemp.Close
  Set rsTemp = Nothing

  Set objCurrentNode = Nothing
  trvAttachmentDir_Populate = True

Exit Function

LocalErr:
  MsgBox "Error setting the Email Attachment Path." & vbNewLine & "Please see your system administrator.", vbCritical, "Configuration"
  trvAttachmentDir_Populate = False

End Function


Private Sub PopulateFolder(objParentNode As MSComctlLib.Node) ', strCurrentDirectory As String)

  Dim rsTemp As New ADODB.Recordset
  Dim strSQL As String
  Dim objCurrentNode As MSComctlLib.Node

  On Error GoTo LocalErr

  'Only do each node once !!!
  If objParentNode.Tag = "1" Then
    Exit Sub
  End If
  objParentNode.Tag = "1"

  strSQL = "master..xp_subdirs '" & objParentNode.key & "'"
  rsTemp.Open strSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText

  If Not rsTemp.BOF And Not rsTemp.EOF Then
    Do While Not rsTemp.EOF
      'MH20040127 Fault 6955
      If Trim(rsTemp.Fields(0).value) <> vbNullString Then
        Set objCurrentNode = trvAttachmentDir.Nodes.Add(objParentNode, ssatChild, objParentNode.key & rsTemp.Fields(0).value & "\", rsTemp.Fields(0).value, "IMG_CLOSEDFOLDER", "IMG_OPENFOLDER")
      End If
      rsTemp.MoveNext
    Loop
  End If

LocalErr:
  If Not (rsTemp Is Nothing) Then
    If rsTemp.State = adStateOpen Then
      rsTemp.Close
    End If
    Set rsTemp = Nothing
  End If
  Set objCurrentNode = Nothing

End Sub


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


Private Sub trvAttachmentDir_Click()
  If Not (trvAttachmentDir.SelectedItem Is Nothing) Then
    'Check that the server name is not highlighted !
    If Not (trvAttachmentDir.SelectedItem.Parent Is Nothing) Then
      mstrPath = trvAttachmentDir.SelectedItem.key
      Me.Caption = "Attachment Path - " & mstrPath
    End If
  End If
End Sub

Private Sub trvAttachmentDir_DblClick()
  If Not (trvAttachmentDir.SelectedItem Is Nothing) Then
    'Check that the server name is not highlighted !
    If Not (trvAttachmentDir.SelectedItem.Parent Is Nothing) Then
      cmdOK_Click
    End If
  End If
End Sub

Private Sub trvAttachmentDir_Expand(ByVal Node As MSComctlLib.Node)

  Dim objCurrentNode As MSComctlLib.Node

  If Node.Children > 0 Then
    Set objCurrentNode = Node.Child.FirstSibling
    Do While Not (objCurrentNode Is Nothing)
      If Not (objCurrentNode.Parent Is Nothing) Then
        If objCurrentNode.Parent = Node Then
          PopulateFolder objCurrentNode
        End If
      End If
      Set objCurrentNode = objCurrentNode.Next
    Loop
  
    Set objCurrentNode = Nothing
  End If

End Sub

Public Property Get AttachmentPath() As String
  AttachmentPath = mstrPath
End Property

Public Property Let AttachmentPath(ByVal strNewValue As String)
  
  Dim objNode As MSComctlLib.Node
  Dim lngFound As Long

  On Local Error GoTo LocalErr

  mstrPath = strNewValue
  Me.Caption = "Attachment Path - " & mstrPath

  lngFound = InStr(strNewValue, "\")
  Do While lngFound > 0
    Set objNode = trvAttachmentDir.Nodes(Left(strNewValue, lngFound))
    Call trvAttachmentDir_Expand(objNode)
    objNode.Expanded = True
    lngFound = InStr(lngFound + 1, strNewValue, "\")
  Loop

  If Not (objNode Is Nothing) Then
    objNode.Selected = True
  End If

Exit Property

LocalErr:
  MsgBox "Unable to find path <" & strNewValue & ">", vbInformation, "Email Attachment Path"

End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property
