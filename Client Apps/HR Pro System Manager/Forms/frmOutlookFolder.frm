VERSION 5.00
Object = "{1C203F10-95AD-11D0-A84B-00A0247B735B}#1.0#0"; "SSTree.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOutlookFolder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Outlook Folder Definition"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5940
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1062
   Icon            =   "frmOutlookFolder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDefinition 
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5715
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   2
         Top             =   315
         Width           =   3885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   1
         Top             =   365
         Width           =   510
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   3300
      TabIndex        =   9
      Top             =   4725
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4650
      TabIndex        =   10
      Top             =   4725
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   3
      Top             =   1000
      Width           =   5715
      Begin VB.CommandButton cmdCalc 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   5205
         TabIndex        =   8
         Top             =   3120
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.TextBox txtCalc 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1665
         TabIndex        =   7
         Top             =   3120
         Width           =   3540
      End
      Begin VB.OptionButton optCalculated 
         Caption         =   "Calculated"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   3180
         Width           =   1260
      End
      Begin VB.OptionButton optFixed 
         Caption         =   "Fixed"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin SSActiveTreeView.SSTree Treeview1 
         Height          =   2655
         Left            =   1665
         TabIndex        =   5
         Top             =   360
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   4683
         _Version        =   65538
         LabelEdit       =   1
         LineStyle       =   1
         NodeSelectionStyle=   2
         Indentation     =   315
         AutoSearch      =   0   'False
         HideSelection   =   0   'False
         PictureBackgroundUseMask=   0   'False
         HasFont         =   -1  'True
         HasMouseIcon    =   0   'False
         HasPictureBackground=   0   'False
         ImageList       =   "ImageList1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblDisabledTreeView 
         BorderStyle     =   1  'Fixed Single
         Height          =   2655
         Left            =   1665
         TabIndex        =   11
         Top             =   360
         Width           =   3855
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   4560
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
            Picture         =   "frmOutlookFolder.frx":000C
            Key             =   "OPENFLDR"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutlookFolder.frx":03D9
            Key             =   "CALENDAR"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOutlookFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjOutlookFolder As clsOutlookFolder
Private mblnCancelled As Boolean
Private mlngTableID As Long
Private mblnReadOnly As Boolean

Public Property Get OutlookFolder() As clsOutlookFolder
  Set OutlookFolder = mobjOutlookFolder
End Property

Public Sub Initialise(ByVal objNewValue As clsOutlookFolder, blnCopy As Boolean)

  Set mobjOutlookFolder = objNewValue

  With mobjOutlookFolder

    TableID = .TableID

    If .FolderType = 0 Then
      optFixed.value = True

    Else
      optCalculated.value = True
      txtCalc.Tag = .CalcExprID
      txtCalc.Text = GetExpressionName(.CalcExprID)

    End If

    PopulateTreeView

    If blnCopy Then
      mobjOutlookFolder.FolderID = 0
      txtName.Text = "Copy of " & .Name
      Changed = True
    Else
      txtName.Text = .Name
      Changed = False
    End If

    If .FixedPath <> vbNullString And TreeView1.SelectedItem Is Nothing Then
      MsgBox "This definition is set to output to outlook folder '" & .FixedPath & _
       "' which you do not currently have access to.", vbExclamation, Me.Caption
    End If

  End With

End Sub


Public Sub PopulateTreeView()

  On Error GoTo LocalErr

  Dim olApp As Outlook.Application
  Dim olNameSpace As Outlook.NameSpace
  Dim objFolder As Outlook.MAPIFolder

  Dim lngIndex As Long

  Screen.MousePointer = vbHourglass

  lblDisabledTreeView.BackColor = vbWindowBackground
  TreeView1.Visible = False
  TreeView1.Nodes.Clear
  TreeView1.Sorted = True

  'If IsCalendarArrayPopulated Then
  '  For lngIndex = 0 To UBound(gstrOutlookCalendarsArray)
  '    AddCalendarToTreeView gstrOutlookCalendarsArray(lngIndex)
  '  Next
  'Else
    Set olApp = New Outlook.Application
    Set olNameSpace = olApp.GetNamespace("MAPI")
    For Each objFolder In olNameSpace.Folders
      ProcessFolder objFolder
    Next
    Set olNameSpace = Nothing
    Set olApp = Nothing
  'End If

  TreeView1.Visible = (optFixed.value = True)
  lblDisabledTreeView.BackColor = vbButtonFace
  Screen.MousePointer = vbDefault

Exit Sub

LocalErr:
  Screen.MousePointer = vbDefault
  MsgBox "Error populating outlook folders" & _
         IIf(Err.Description <> vbNullString, vbCrLf & "(" & Err.Description & ")", vbNullString), vbCritical

End Sub


Private Sub ProcessFolder(objParentFolder As MAPIFolder, Optional objParentNode As SSNode)

  Dim objFolder As MAPIFolder
  Dim objNode As SSNode
  Dim strIcon As String

  If ValidFolder(objParentFolder) Then

    strIcon = IIf(objParentFolder.DefaultItemType = 1, "CALENDAR", "OPENFLDR")
    If objParentNode Is Nothing Then
      Set objNode = TreeView1.Nodes.Add(, , , " " & objParentFolder.Name, strIcon, strIcon)
    Else
      Set objNode = TreeView1.Nodes.Add(objParentNode, tvwChild, , " " & objParentFolder.Name, strIcon, strIcon)
    End If

    For Each objFolder In objParentFolder.Folders
      ProcessFolder objFolder, objNode
    Next

    objNode.Tag = Replace("\\" & objNode.FullPath, "\ ", "\")
    objNode.Sorted = True
    objNode.Expanded = True

    'If objParentFolder.DefaultItemType = 1 Then
    '  AddToOutlookCalendarArray objNode.Tag
    'End If

    If mobjOutlookFolder.FolderType = 0 Then
      If objNode.Tag = mobjOutlookFolder.FixedPath Then
        objNode.Selected = True
      End If
    End If

  End If

End Sub


Private Function ValidFolder(objParentFolder As MAPIFolder) As Boolean

  Dim objFolder As MAPIFolder

  On Local Error GoTo TidyAndExit
  
  ValidFolder = False

  If objParentFolder.DefaultItemType = 1 Then
    ValidFolder = True
    Exit Function
  End If

  For Each objFolder In objParentFolder.Folders
    If ValidFolder(objFolder) Then
      ValidFolder = True
      Exit Function
    End If
  Next

TidyAndExit:

End Function


Private Sub cmdCalc_Click()

  ' Display the Record Description selection form.
  Dim fOK As Boolean
  Dim objExpr As CExpression

  ' Instantiate an expression object.
  Set objExpr = New CExpression

  With objExpr
    fOK = .Initialise(mlngTableID, Val(txtCalc.Tag), giEXPR_OUTLOOKFOLDER, giEXPRVALUE_CHARACTER)

    If fOK Then
      ' Instruct the expression object to display the
      ' expression selection form.
      If .SelectExpression Then
        txtCalc.Tag = .ExpressionID
        If txtCalc.Tag < 0 Then
          txtCalc.Tag = 0
        End If
        txtCalc.Text = GetExpressionName(.ExpressionID)
        Changed = True
      Else
        ' Check in case the original expression has been deleted.
        With recExprEdit
          .Index = "idxExprID"
          .Seek "=", Val(txtCalc.Tag), False

          If .NoMatch Then
            txtCalc.Tag = 0
            txtCalc.Text = vbNullString
          End If
        End With
      End If
    End If
  End With

  ' Disassociate object variables.
  Set objExpr = Nothing

End Sub

Private Sub cmdCancel_Click()
  
  If Me.Changed Then
    Select Case MsgBox("You have made changes...do you wish to save these changes ?", vbQuestion + vbYesNoCancel, App.Title)
    Case vbYes
      cmdOK_Click
      Exit Sub
    Case vbCancel
      Exit Sub
    End Select
  End If

  Me.Hide

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  If UnloadMode <> vbFormCode Then
    cmdCancel_Click
    Cancel = True
  End If

End Sub

Private Sub optCalculated_Click()
  'chkShowAll.Enabled = False
  TreeView1.Visible = False
  cmdCalc.Enabled = True
  Changed = True
End Sub

Private Sub optFixed_Click()
  'chkShowAll.Enabled = True
  TreeView1.Visible = True
  txtCalc.Tag = 0
  txtCalc.Text = vbNullString
  cmdCalc.Enabled = False
  Changed = True
End Sub

Private Sub TreeView1_NodeClick(Node As SSActiveTreeView.SSNode)
  mobjOutlookFolder.FixedPath = Node.Tag
  Changed = True
End Sub

Private Sub txtName_Change()
  Changed = True
End Sub

Private Sub txtName_GotFocus()
  With txtName
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Public Property Let Cancelled(ByVal blnNewValue As Boolean)
  mblnCancelled = blnNewValue
End Property

Private Sub cmdOK_Click()

  If ValidDefinition = False Then
    Exit Sub
  End If

  SaveDefinition
  mblnCancelled = False
  Me.Hide

  Application.ChangedOutlookLink = True

End Sub

Private Sub Form_Load()
  mblnCancelled = True

  mblnReadOnly = (Application.AccessMode = accSystemReadOnly)

  If mblnReadOnly Then
    txtName.BackColor = vbButtonFace
    txtName.Enabled = False
    optFixed.Enabled = False
    optCalculated.Enabled = False
    TreeView1.BackColor = vbButtonFace
    TreeView1.ForeColor = vbApplicationWorkspace
  End If

End Sub


Private Function ValidDefinition() As Boolean

  On Error GoTo ErrorTrap
  
  ValidDefinition = False

  If Trim(txtName.Text) = vbNullString Then
    MsgBox "You must give this definition a name.", vbExclamation, Me.Caption
    txtName.SetFocus
    Exit Function
  End If

  
  
  With recOutlookFolders
    If Not .BOF And Not .EOF Then
      .MoveFirst
      Do While Not .EOF
        
        If !TableID = mobjOutlookFolder.TableID Or !TableID = 0 Then
          If (Trim(!Name) = Trim(txtName.Text)) And _
            (!FolderID <> mobjOutlookFolder.FolderID) And _
            (!Deleted = False) Then
              MsgBox "An outlook folder named '" & Trim(txtName.Text) & "' already exists !", vbOKOnly + vbExclamation, Application.Name
            Exit Function
          End If
        End If
  
        .MoveNext
      Loop
    End If
  End With



  If optFixed.value = True Then
    If TreeView1.SelectedItem Is Nothing Then
      MsgBox "You must select an outlook folder.", vbExclamation, Me.Caption
      Exit Function
    End If
  Else
    If Val(txtCalc.Tag) = 0 Then
      MsgBox "You must select a calculation.", vbExclamation, Me.Caption
      Exit Function
    End If
  End If

  ValidDefinition = True

TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  MsgBox "Error validating Outlook folder", vbCritical, Me.Caption
  ValidDefinition = False
  Resume TidyUpAndExit

End Function


Private Function SaveDefinition() As Boolean

  With mobjOutlookFolder

    .Name = txtName.Text
    
    If optFixed.value = True Then
      .TableID = 0
      .FolderType = 0
      .FixedPath = TreeView1.SelectedItem.Tag
      .CalcExprID = 0
    Else
      .TableID = mlngTableID
      .FolderType = 1
      .FixedPath = vbNullString
      .CalcExprID = Val(txtCalc.Tag)
    End If
  
  End With

End Function


Public Property Get TableID() As Long
  TableID = mlngTableID
End Property

Public Property Let TableID(ByVal lngNewValue As Long)
  mlngTableID = lngNewValue
End Property


'Private Function IsCalendarArrayPopulated() As Boolean
'
'  Dim lngIndex As Long
'
'  On Local Error Resume Next
'  Err.Clear
'  lngIndex = UBound(gstrOutlookCalendarsArray)
'
'  IsCalendarArrayPopulated = (Err.Number = 0)
'
'End Function
'
'
'Private Function AddToOutlookCalendarArray(strCalendarPath As String) As Boolean
'
'  Dim lngIndex As Long
'
'  If Not IsCalendarArrayPopulated Then
'    ReDim gstrOutlookCalendarsArray(0)
'    lngIndex = 0
'  Else
'    lngIndex = lngIndex + 1
'    ReDim Preserve gstrOutlookCalendarsArray(lngIndex)
'  End If
'
'  gstrOutlookCalendarsArray(lngIndex) = strCalendarPath
'  Debug.Print strCalendarPath
'
'End Function
'
'
'Private Function AddCalendarToTreeView(strCalendarPath As String) As Boolean
'
'  Dim objParentNode As SSNode
'  Dim strFolders() As String
'  Dim lngIndex As Long
'
'  On Local Error Resume Next
'
'  Set objParentNode = Nothing
'  strFolders = Split(strCalendarPath)
'
'  For lngIndex = 0 To UBound(strFolders)
'
'    strIcon = IIf(lngIndex = UBound(strFolders), "CALENDAR", "OPENFLDR")
'
'    '1. Create from Root
'    '2. Create from Parent
'    '3. Get node
'
'
'    If objParentNode Is Nothing Then
'      Set objParentNode = TreeView1.Nodes.Add(, , , " " & strFolders(lngIndex), strIcon, strIcon)
'    Else
'      Set objParentNode = TreeView1.Nodes.Add(objParentNode, tvwChild, , " " & strFolders(lngIndex), strIcon, strIcon)
'    End If
'
'  Next
'
'End Function

Public Property Get Changed() As Boolean
  Changed = cmdOK.Enabled
End Property

Public Property Let Changed(ByVal blnNewValue As Boolean)
  cmdOK.Enabled = blnNewValue And Not mblnReadOnly
End Property

