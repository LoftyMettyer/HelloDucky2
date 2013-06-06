VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{1C203F10-95AD-11D0-A84B-00A0247B735B}#1.0#0"; "sstree.ocx"
Begin VB.Form frmExpr 
   Caption         =   "Expression Definition"
   ClientHeight    =   5955
   ClientLeft      =   495
   ClientTop       =   1560
   ClientWidth     =   8805
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5018
   Icon            =   "frmExpr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   8805
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraDefinition 
      Caption         =   "Definition :"
      Height          =   3945
      Index           =   1
      Left            =   100
      TabIndex        =   8
      Top             =   1900
      Width           =   8600
      Begin VB.Frame fraButtons 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   900
         Index           =   1
         Left            =   7200
         TabIndex        =   14
         Top             =   2760
         Width           =   1200
         Begin VB.CommandButton cmdOk 
            Caption         =   "&OK"
            Default         =   -1  'True
            Height          =   400
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   1200
         End
         Begin VB.CommandButton cmdCancel 
            Cancel          =   -1  'True
            Caption         =   "&Cancel"
            Height          =   400
            Left            =   0
            TabIndex        =   21
            Top             =   500
            Width           =   1200
         End
      End
      Begin VB.Frame fraButtons 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2400
         Index           =   0
         Left            =   7200
         TabIndex        =   13
         Top             =   240
         Width           =   1200
         Begin VB.CommandButton cmdDeleteComponent 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   0
            TabIndex        =   18
            Top             =   1500
            Width           =   1200
         End
         Begin VB.CommandButton cmdModifyComponent 
            Caption         =   "&Edit..."
            Height          =   400
            Left            =   0
            TabIndex        =   17
            Top             =   1000
            Width           =   1200
         End
         Begin VB.CommandButton cmdInsertComponent 
            Caption         =   "&Insert..."
            Height          =   400
            Left            =   0
            TabIndex        =   16
            Top             =   500
            Width           =   1200
         End
         Begin VB.CommandButton cmdAddComponent 
            Caption         =   "&Add..."
            Height          =   400
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   1200
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
            Height          =   400
            Left            =   0
            TabIndex        =   19
            Top             =   2000
            Width           =   1200
         End
      End
      Begin SSActiveTreeView.SSTree sstrvComponents 
         Height          =   3525
         Left            =   150
         TabIndex        =   6
         Top             =   255
         Width           =   6900
         _ExtentX        =   12171
         _ExtentY        =   6218
         _Version        =   65538
         NodeSelectionStyle=   2
         Style           =   6
         Indentation     =   315
         AutoSearch      =   0   'False
         HideSelection   =   0   'False
         PictureBackgroundUseMask=   0   'False
         HasFont         =   -1  'True
         HasMouseIcon    =   0   'False
         HasPictureBackground=   0   'False
         ImageList       =   "<None>"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LoadStyleRoot   =   1
      End
   End
   Begin VB.Frame fraDefinition 
      Height          =   1860
      Index           =   0
      Left            =   100
      TabIndex        =   7
      Top             =   0
      Width           =   8600
      Begin VB.TextBox txtDescription 
         Height          =   1000
         Left            =   1335
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   650
         Width           =   3135
      End
      Begin VB.TextBox txtOwner 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   5535
         TabIndex        =   2
         Top             =   250
         Width           =   2910
      End
      Begin VB.OptionButton optAccess 
         Caption         =   "Read / &Write"
         Height          =   315
         Index           =   0
         Left            =   5535
         TabIndex        =   3
         Top             =   675
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.OptionButton optAccess 
         Caption         =   "&Read only"
         Height          =   315
         Index           =   1
         Left            =   5535
         TabIndex        =   4
         Top             =   1000
         Width           =   1230
      End
      Begin VB.OptionButton optAccess 
         Caption         =   "&Hidden"
         Height          =   315
         Index           =   2
         Left            =   5535
         TabIndex        =   5
         Top             =   1325
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox txtExpressionName 
         Height          =   315
         Left            =   1335
         MaxLength       =   255
         TabIndex        =   0
         Top             =   250
         Width           =   3135
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Description :"
         Height          =   195
         Left            =   195
         TabIndex        =   12
         Top             =   705
         Width           =   1125
      End
      Begin VB.Label lblOwner 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Owner :"
         Height          =   195
         Left            =   4695
         TabIndex        =   11
         Top             =   315
         Width           =   720
      End
      Begin VB.Label lblAccess 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Access :"
         Height          =   195
         Left            =   4695
         TabIndex        =   10
         Top             =   705
         Width           =   690
      End
      Begin VB.Label lblExpressionName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         Height          =   195
         Left            =   200
         TabIndex        =   9
         Top             =   310
         Width           =   510
      End
   End
   Begin MSComDlg.CommonDialog comDlgBox 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontName        =   "Verdana"
   End
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   15
      Top             =   1035
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "frmExpr.frx":000C
   End
End
Attribute VB_Name = "frmExpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Expression definition variables.
Private mobjExpression As CExpression
Private mcolComponents As Collection

' Form handling variables.
Private mfModifiable As Boolean
Private mfCancelled As Boolean
Private mfChanged As Boolean
Private mfForcedReadOnly As Boolean

' Form handling constants.
Const ROOTKEY = "EXPRESSION_ROOT"

' Cut'n Paste Functionality
Private mcolClipboard As Collection
Private mbCanCut As Boolean
Private mbCanCopy As Boolean
Private mbCanPaste As Boolean
Private mbCanMoveUp As Boolean
Private mbCanMoveDown As Boolean
Private mbColoursOn As Boolean

Private mbCanDelete As Boolean
Private mbCanEdit As Boolean
Private mbCanInsert As Boolean

Private mblnReadOnly As Boolean

Private Enum UndoTypes
  giUNDO_DELETE = 1
  giUNDO_PASTE = 2
  giUNDO_CUT = 3
  giUNDO_ADD = 4
  giUNDO_INSERT = 5
  giUNDO_MOVEUP = 6
  giUNDO_MOVEDOWN = 7
  giUNDO_EDIT = 8
  giUNDO_RENAME = 9
End Enum

Private mcolUndoData() As CExpression
Private maUndoTypes() As UndoTypes
Private miUndoLevel As Integer

' JPD20021108 Fault 3287
Private msShortcutKeys As String

Private mfLabelEditing As Boolean

Public Property Get Expression() As CExpression
  ' Return the expression that is being editted.
  Set Expression = mobjExpression
  
End Property
Public Property Set Expression(pobjExpr As CExpression)
  ' Set the expression that is being editted.
  Set mobjExpression = pobjExpr

  ' Check if the user can modify the expression.
  mfModifiable = ((gsUserName = mobjExpression.Owner) _
      Or (mobjExpression.Access = ACCESS_READWRITE)) _
    And (Not mblnReadOnly) _
    And (Not mfForcedReadOnly)

  ' Update the form display for the new expression.
  ConfigureScreen
  
  ' Update the screen controls with the expression's properties.
  With mobjExpression
    txtExpressionName.Text = .Name
    txtDescription.Text = .Description
    txtOwner.Text = .Owner
    
    'optAccess(.Access).Value = True
    Select Case .Access
      Case ACCESS_READWRITE: optAccess(0).value = True
      Case ACCESS_READONLY: optAccess(1).value = True
      Case ACCESS_HIDDEN: optAccess(2).value = True
    End Select
  
  End With
  
  ' Populate the treeview with the expression definition.
  PopulateTreeView
  
  ' Set the initial view
  SetInitialExpandedNodes
  
  If (mobjExpression.ExpressionID = 0) And _
    (Len(mobjExpression.Name) > 0) Then
    ' ie. if we are copying an existing expression.
    mfChanged = True
  Else
    ClearChanged
  End If
  
End Property

Private Sub cmdAddComponent_Click()

  ' Place components on the undo collection
  CreateUndoView (giUNDO_ADD)

  AddComponent (True)

End Sub

Private Sub cmdCancel_Click()

  Dim intAnswer As Integer

  'JPD 20030730 Fault 5587
  mfLabelEditing = False
  
  ' Check if any changes have been made.
  If CheckChanged Then
    
    ' JDM - 22/08/01 - Fault 2714 - Changed the cancel message to be consistent.
    intAnswer = MsgBox("You have changed the current definition. Save changes ?", vbQuestion + vbYesNoCancel + vbDefaultButton1, App.ProductName)
    If intAnswer = vbYes Then
        Call cmdOk_Click
        Exit Sub
    ElseIf intAnswer = vbCancel Then
        Exit Sub
    End If
    
'    If MsgBox(" The " & LCase(mobjExpression.ExpressionTypeName) & " definition has changed.  Save changes ?", vbQuestion + vbYesNo + vbDefaultButton1, App.ProductName) = vbYes Then
'      Call cmdOK_Click
'      Exit Sub
'    End If
  End If
  
  'Me.Cancelled = True
  
  ' Unload the form.
  UnLoad Me

End Sub

Private Sub cmdDeleteComponent_Click()
   
  ' Place components on the undo collection
  CreateUndoView (giUNDO_DELETE)
   
  DeleteComponents
   
End Sub

Private Sub cmdInsertComponent_Click()

  ' Place components on the undo collection
  CreateUndoView (giUNDO_INSERT)

  InsertComponent (True)
  
End Sub

Private Sub cmdModifyComponent_Click()
  Dim objParentExpression As CExpression
  Dim objCurrentComponent As CExprComponent
  Dim objNewComponent As CExprComponent
  Dim sNewComponentKey As String
  Dim sNextNodeKey  As String
  
  ' Place components on the undo collection
  CreateUndoView (giUNDO_EDIT)
  
  ' Get the selected component, and it's parent expression.
  Set objCurrentComponent = SelectedComponent(sstrvComponents.SelectedItem)
  Set objParentExpression = SelectedExpression(sstrvComponents.SelectedItem)
  
  ' Let the parent expression handle the modification of a component.
  Set objNewComponent = objParentExpression.ModifyComponent(objCurrentComponent)
  If Not objNewComponent Is Nothing Then
    sNextNodeKey = sstrvComponents.SelectedItem.key

    ' Add the modified component to the treeview.
    sNewComponentKey = InsertComponentNode(objNewComponent, sNextNodeKey, True, False)
    
    ' Select the new component.
    sstrvComponents.SelectedItem = sstrvComponents.Nodes(sNewComponentKey)
    sstrvComponents.SelectedItem.Expanded = True
    
    ' Remove the old version of the component from the treeview.
    RemoveComponentNode sNextNodeKey
    
    mfChanged = True
    
    ' Ensure the command buttons are configured for the selected component.
    RefreshButtons
  End If

  ' Disassociate object variables.
  Set objCurrentComponent = Nothing
  Set objParentExpression = Nothing
  Set objNewComponent = Nothing
  
End Sub

Private Sub cmdOk_Click()
  ' If no changes have been made to the expression, treat the OK command as AccessCodes Cancel command.
  If CheckChanged Then
    ' Check expression
    If CheckExpression Then
      Cancelled = False
      UnLoad Me
    Else
      ' Ensure the command buttons are configured for the selected component.
      RefreshButtons
    End If
  Else
    cmdCancel_Click
  End If

End Sub

Private Sub RefreshButtons()
  ' Enable/Disable the button depending on what treenode is selected.
  
  Dim objNode As SSActiveTreeView.SSNode
  Dim iNodesSelected As Integer
  
  ' By default allow everything to be done
  mbCanEdit = True
  mbCanDelete = True
  mbCanInsert = True
  mbCanCut = True
  mbCanCopy = True
  mbCanPaste = True
  mbCanMoveDown = True
  mbCanMoveUp = True
  iNodesSelected = 0

  ' Loop through each selected node
  For Each objNode In sstrvComponents.Nodes
    If objNode.Selected = True Then
      iNodesSelected = iNodesSelected + 1
      
      ' If the root node is selected then disable the Insert/Modify/Delete buttons.
      If objNode.key = ROOTKEY Then
        mbCanInsert = False
        mbCanEdit = False
        mbCanDelete = False
        mbCanCut = False
        mbCanPaste = (mcolClipboard.Count > 0) And mbCanPaste
        mbCanCopy = False
        mbCanMoveDown = False
        mbCanMoveUp = False
      Else
        Select Case mcolComponents.Item(objNode.key).ComponentType
          ' Enable the Insert/Modify/Delete buttons for function components.
          Case giCOMPONENT_FUNCTION
            mbCanCut = True And mbCanCut
            mbCanPaste = (mcolClipboard.Count > 0) And mbCanPaste
            mbCanCopy = True And mbCanCopy
            mbCanMoveDown = Not (objNode.LastSibling.Index = objNode.Index) And mbCanMoveDown
            mbCanMoveUp = Not (objNode.FirstSibling.Index = objNode.Index) And mbCanMoveUp
          
          ' Disable the Insert/Modify/Delete buttons for function parameter expressions.
          ' Enable the Insert/Modify/Delete buttons for true expressions.
          Case giCOMPONENT_EXPRESSION
            mbCanEdit = Not (mcolComponents.Item(objNode.Parent.key).ComponentType = giCOMPONENT_FUNCTION) And mbCanEdit
            mbCanInsert = False
            mbCanDelete = Not (mcolComponents.Item(objNode.Parent.key).ComponentType = giCOMPONENT_FUNCTION) And mbCanDelete
            mbCanCut = False
            mbCanPaste = (mcolClipboard.Count > 0) And mbCanPaste
            mbCanCopy = False
            mbCanMoveDown = False
            mbCanMoveUp = False
  
          ' Enable the Insert/Modify/Delete buttons by default.
          Case Else
            mbCanDelete = True And mbCanDelete
            mbCanCut = True And mbCanCut
            mbCanPaste = (mcolClipboard.Count > 0) And mbCanPaste
            mbCanCopy = True And mbCanCopy
            mbCanMoveDown = Not (objNode.LastSibling.Index = objNode.Index) And mbCanMoveDown
            mbCanMoveUp = Not (objNode.FirstSibling.Index = objNode.Index) And mbCanMoveUp
        End Select
      End If
    End If
  Next objNode

  ' Only allow edit and insert when single nodes are selected
  mbCanMoveDown = (iNodesSelected = 1) And mbCanMoveDown And mfModifiable
  mbCanMoveUp = (iNodesSelected = 1) And mbCanMoveUp And mfModifiable
  mbCanInsert = (iNodesSelected = 1) And mbCanInsert And mfModifiable
  mbCanEdit = (iNodesSelected = 1) And mbCanEdit And mfModifiable
  mbCanDelete = (iNodesSelected > 0) And mbCanDelete And mfModifiable

  ' Enable/disable controls depending on the selected component.
  cmdInsertComponent.Enabled = mbCanInsert
  cmdModifyComponent.Enabled = mbCanEdit
  cmdDeleteComponent.Enabled = mbCanDelete

'  If sstrvComponents.Visible And sstrvComponents.Enabled Then
'    sstrvComponents.SetFocus
'  End If

End Sub

Private Sub cmdPrint_Click()

  Dim bCancelPrint As Boolean

  On Error GoTo Cancelled
  bCancelPrint = False
  
'  'JDM - 08/08/01 - Fault 2190 - Show the dialog box when printing
'  With comDlgBox
'    .Flags = cdlPDNoSelection Or cdlPDHidePrintToFile Or cdlPDReturnDC
'    .ShowPrinter
'    Printer.Copies = .Copies
'  End With
  
  If Not bCancelPrint Then
    Screen.MousePointer = vbHourglass
    mobjExpression.PrintExpressionWithoutConstructing
    Screen.MousePointer = vbDefault
  End If

  Exit Sub

Cancelled:
  bCancelPrint = True
  Resume Next

End Sub

Private Sub Form_Activate()
  'JPD 20040507 Fault 7094
  sstrvComponents.ApproximateNodeCount = sstrvComponents.Nodes.Count

End Sub

Private Sub Form_Initialize()

  ' Initialise cut'n paste options
  Set mcolClipboard = New Collection
  mbCanCut = False
  mbCanPaste = False
  mbCanCopy = False
  mbCanMoveUp = False
  mbCanMoveDown = False
  mbColoursOn = False
  
  ' Initialise the undo functionality
  ReDim mcolUndoData(0)
  ReDim maUndoTypes(0)
  miUndoLevel = 0

End Sub

Private Sub CopyComponents()

  Dim objNode As SSActiveTreeView.SSNode
  Dim iCount As Integer

  ' Clear the exiting clipboard
  For iCount = mcolClipboard.Count To 1 Step -1
    mcolClipboard.Remove (iCount)
  Next iCount

  ' Place selected components on the pasteboard
  For Each objNode In sstrvComponents.Nodes
    If objNode.Selected = True Then
      mcolClipboard.Add SelectedComponent(objNode).CopyComponent
    End If
  Next objNode
  
  ' Tidy up
  Set objNode = Nothing
  
End Sub

Private Sub CutComponents()

  ' Place components on the undo collection
  CreateUndoView (giUNDO_CUT)

  ' Place the selected components on the pasteboard
  CopyComponents
  DeleteComponents

End Sub

Private Sub PasteComponents()

  'Pastes the collection of clipboard components into the expression

  Dim objNewComponent As CExprComponent

  ' Place components on the undo collection
  CreateUndoView (giUNDO_PASTE)

  For Each objNewComponent In mcolClipboard
    If sstrvComponents.SelectedItem.key = ROOTKEY Then
      AddComponent False, objNewComponent
    Else
      If SelectedComponent(sstrvComponents.SelectedItem).ComponentType = giCOMPONENT_EXPRESSION Then
        AddComponent False, objNewComponent
      Else
        InsertComponent False, objNewComponent, True
      End If
    End If
  Next objNewComponent


End Sub

Public Sub MoveComponentDown()

Dim mobjComponent As CExprComponent
Set mobjComponent = New CExprComponent
Dim miOldNode As Integer

' Place components on the undo collection
CreateUndoView (giUNDO_MOVEDOWN)

'Ensure we are not the bottom node for this child
If Not sstrvComponents.SelectedItem.LastSibling.Index = sstrvComponents.SelectedItem.Index Then

  'Copy the current component, and node ID
  Set mobjComponent = SelectedComponent(sstrvComponents.SelectedItem)
  miOldNode = sstrvComponents.SelectedItem.key

  'Move the current treeview node selection to the next one down
  sstrvComponents.SelectedItem.Next.Selected = True

  'Paste the object into the expression
  InsertComponent False, mobjComponent, True

  ' Remove the old component and node
  If SelectedExpression(sstrvComponents.SelectedItem).DeleteComponent(mobjComponent) = True Then
    RemoveComponentNode (miOldNode)
  End If

End If

'Refresh the display
RefreshButtons

'Clear up memory
Set mobjComponent = Nothing

End Sub

Public Sub MoveComponentUp()

Dim mobjComponent As CExprComponent
Set mobjComponent = New CExprComponent
Dim miOldNode As Integer

' Place components on the undo collection
CreateUndoView (giUNDO_MOVEUP)

'Ensure we are not the bottom node for this child
If Not sstrvComponents.SelectedItem.FirstSibling.Index = sstrvComponents.SelectedItem.Index Then

  'Copy the current component, and node ID
  Set mobjComponent = SelectedComponent(sstrvComponents.SelectedItem)
  miOldNode = sstrvComponents.SelectedItem.key

  'Move the current treeview node selection to the next one down
  sstrvComponents.SelectedItem.Previous.Selected = True

  'Paste the object into the expression
  InsertComponent False, mobjComponent, False

  ' Remove the old component and node
  If SelectedExpression(sstrvComponents.SelectedItem).DeleteComponent(mobjComponent) = True Then
    RemoveComponentNode (miOldNode)
  End If

End If

'Refresh the display
RefreshButtons

'Clear up memory
Set mobjComponent = Nothing

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select

' JDM - 15/03/01 - Fault 1934 - Only do things if we have access
If mfModifiable = False Then
    KeyCode = 0
    Shift = 0
End If

' JDM - 15/02/01 - fault 1868 - Error when pressing CTRL-X on treeview control
' For some reason the Sheridan treeview control wants to fire off it own cutn'paste functionality
' must trap it here not in it's own keydown event
If ActiveControl.Name = "sstrvComponents" Then
    
  ' Cut component
  If (Shift And vbCtrlMask) And KeyCode = Asc("X") Then
    If mbCanCut Then
      CutComponents
    End If
    KeyCode = 0
    Shift = 0
  End If

  ' Copy component
  If (Shift And vbCtrlMask) And KeyCode = Asc("C") Then
    If mbCanCopy Then
      CopyComponents
    End If
    KeyCode = 0
    Shift = 0
  End If

  ' Paste component
  If (Shift And vbCtrlMask) And KeyCode = Asc("V") Then
      If mbCanPaste Then
        PasteComponents
      End If
      KeyCode = 0
      Shift = 0
  End If

  ' Delete components
  If KeyCode = vbKeyDelete And mbCanDelete Then
    cmdDeleteComponent_Click
    KeyCode = 0
    Shift = 0
  End If

  ' Insert components
  If KeyCode = vbKeyInsert And mbCanInsert Then
    cmdInsertComponent_Click
    KeyCode = 0
    Shift = 0
  End If

  ' Undo the last action
  'TM20020919 Fault 4408 - was not comparing vbCtrlMask to anything.
  'If KeyCode = vbKeyZ And vbCtrlMask Then
  If KeyCode = vbKeyZ And Shift = vbCtrlMask Then
    ExecuteUndo
    KeyCode = 0
    Shift = 0
  End If

End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  'JPD 20030909 Fault 6937
  If (Not TypeOf Me.ActiveControl Is TextBox) And _
    (Not TypeOf Me.ActiveControl Is SSTree) Then
    
    sstrvComponents_KeyPress KeyAscii
  End If
End Sub

Private Sub Form_Load()
  Dim objOperatorDef As clsOperatorDef
  Dim objFunctionDef As clsFunctionDef
  
  ' Hook the resize event handler
  Hook Me.hWnd, Me.Width, Me.Height
  
  ApplySkinToActiveBar ActiveBar1
  
  ' JPD20021108
  ' Initialise the collections of operators and functions
  ' if not already initialised.
  gobjOperatorDefs.Initialise
  gobjFunctionDefs.Initialise
  
  ' JPD20021108 Fault 3287
  msShortcutKeys = ""
  
  For Each objOperatorDef In gobjOperatorDefs
    If Len(objOperatorDef.ShortcutKeys) > 0 Then
      msShortcutKeys = msShortcutKeys & objOperatorDef.ShortcutKeys
    End If
  Next objOperatorDef
  Set objOperatorDef = Nothing
  
  For Each objFunctionDef In gobjFunctionDefs
    If Len(objFunctionDef.ShortcutKeys) > 0 Then
      msShortcutKeys = msShortcutKeys & objFunctionDef.ShortcutKeys
    End If
  Next objFunctionDef
  Set objFunctionDef = Nothing
  
  ' Clear the menu shortcuts. This needs to be done so that some shortcut keys
  ' (eg. DEL) will function normally in textboxes instead of triggering menu options.
  frmSysMgr.ClearMenuShortcuts
  
  mblnReadOnly = (Application.AccessMode = accSystemReadOnly Or _
                 (Application.AccessMode = accLimited And mobjExpression.ExpressionType <> giEXPR_VIEWFILTER))

  If mblnReadOnly Then
    ControlsDisableAll Me
    
    ' JDM - 15/03/01 - Fault 2792 - Allow user to expand / shrink nodes
    mfModifiable = False
    sstrvComponents.Enabled = True
    
    cmdPrint.Enabled = True
  End If

  fraButtons(0).BackColor = Me.BackColor
  fraButtons(1).BackColor = Me.BackColor

  Cancelled = True

End Sub

Private Sub ConfigureScreen()
  ' Configure the screen controls.
  Dim fUserIsCreator As Integer
  
  ' Configure the screen controls depending on the type of
  ' selection being made.
  Me.Caption = mobjExpression.ExpressionTypeName & " Definition"
  
  ' Only allow the access permmission to be changed by the original creator.
  fUserIsCreator = (gsUserName = mobjExpression.Owner)
  optAccess(0).Enabled = fUserIsCreator And (Not mblnReadOnly) And (Not mfForcedReadOnly)
  optAccess(1).Enabled = fUserIsCreator And (Not mblnReadOnly) And (Not mfForcedReadOnly)
  optAccess(2).Enabled = fUserIsCreator And (Not mblnReadOnly) And (Not mfForcedReadOnly)
    
  ' Do not allow the user to change the expression if they are not allowed to.
  If Not mfModifiable Then
    UI.ControlsDisableAll Me

    ' JDM - 15/03/01 - Fault 1934 - Allow user to expand / shrink nodes
    sstrvComponents.Enabled = True
    ActiveBar1.ForeColor = RGB(0, 0, 0)

    ' JDM - 15/03/01 - Fault 2004 - Enable the print button
    cmdPrint.Enabled = True
  
  End If
    
  ' JDM - 16/03/01 - Fault 2298 - Restore last view
  Select Case glngExpressionViewColours
    Case EXPRESSIONBUILDER_COLOUROFF
      mbColoursOn = False
    Case EXPRESSIONBUILDER_COLOURON
      mbColoursOn = True
  End Select

  ' Get rid of the icon off the form
  RemoveIcon Me
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'JDM - 07/03/01 - Fault 1936 - Ask user if they wish to save changes
    Dim intAnswer As Integer

  'JPD 20030730 Fault 5587
  mfLabelEditing = False
    
    If UnloadMode <> vbFormCode Then

        'Check if any changes have been made.
        If CheckChanged Then
            intAnswer = MsgBox("You have changed the current definition. Save changes ?", vbQuestion + vbYesNoCancel + vbDefaultButton1, App.ProductName)
            If intAnswer = vbYes Then
                Call cmdOk_Click
                If Me.Cancelled = True Then Cancel = 1
            ElseIf intAnswer = vbNo Then
                Me.Cancelled = True
            ElseIf intAnswer = vbCancel Then
                Cancel = 1
            End If
        Else
          Me.Cancelled = True
        End If
    End If

End Sub

Private Function CheckExpression() As Boolean
  ' Check that the expression information is valid.
  On Error GoTo ErrorTrap
  
  Dim fValid As Boolean
  Dim iLoop As Integer
  Dim objComponent As CExprComponent
  Dim objBadComponent As CExprComponent
  Dim iValidityCode As ExprValidationCodes
  Dim objCalcExpr As CExpression
  Dim alngColumns() As Long
  Dim sColumnName As String
  Dim sTableName As String
  
 
 
  ' Check that there is an expression name.
  fValid = (Len(Trim(txtExpressionName.Text)) > 0)
  If Not fValid Then
    MsgBox "The " & LCase(mobjExpression.ExpressionTypeName) & " must be given a name.", _
      vbExclamation + vbOKOnly, Application.Name
    sstrvComponents.SelectedItem = sstrvComponents.Nodes(ROOTKEY)
    mfCancelled = True
    Me.txtExpressionName.SetFocus
  End If
  
  ' Check that the expression is valid.
  If fValid Then
    iValidityCode = mobjExpression.ValidateExpression(True)
    
    fValid = (iValidityCode = giEXPRVALIDATION_NOERRORS)
  
    If Not fValid Then
      
      ' RH 23/10/00 - Allow saving of TypeMisMatch col calc expressions
      '               NOTE : In WriteExpression we write the
      '               ReturnedEvalType rather than the mobjexpression Return type
      '               Otherwise the defsel will show the invalid calc in the list
      '               once you come out of frmExpression.
      
      If (iValidityCode = giEXPRVALIDATION_EXPRTYPEMISMATCH) And _
        (mobjExpression.ExpressionType = giEXPR_COLUMNCALCULATION) Then
         
        ' JPD20020419 Fault 3786
        If (mobjExpression.EvaluatedReturnType = giEXPRVALUE_CHARACTER) Or _
          (mobjExpression.EvaluatedReturnType = giEXPRVALUE_DATE) Or _
          (mobjExpression.EvaluatedReturnType = giEXPRVALUE_LOGIC) Or _
          (mobjExpression.EvaluatedReturnType = giEXPRVALUE_NUMERIC) Then


          
          'NPG20080229 Fault 12944
       
          
'          If MsgBox(mobjExpression.ValidityMessage(iValidityCode) & vbCrLf & _
'                 "You may still save this calculation but you will not be able to select" & vbCrLf & _
'                 "it for a column of this type. Would you like to continue ?", vbExclamation + vbYesNo, Application.Name) = vbYes Then
'            fValid = True
'          End If
          
          
          'NPG20080421 Fault 13110
          If mobjExpression.ExpressionID <> 0 Then
            If Not mobjExpression.ShowExpressionUsage Then
              If MsgBox(mobjExpression.ValidityMessage(iValidityCode) & vbCrLf & _
                     "You may still save this calculation but you will not be able to select" & vbCrLf & _
                     "it for a column of this type. Would you like to continue ?", vbExclamation + vbYesNo, Application.Name) = vbYes Then
                fValid = True
              End If
            End If
          Else
              If MsgBox(mobjExpression.ValidityMessage(iValidityCode) & vbCrLf & _
                     "You may still save this calculation but you will not be able to select" & vbCrLf & _
                     "it for a column of this type. Would you like to continue ?", vbExclamation + vbYesNo, Application.Name) = vbYes Then
                fValid = True
              End If
          End If


        Else
          MsgBox mobjExpression.ValidityMessage(iValidityCode), _
            vbExclamation + vbOKOnly, Application.Name
        End If
      Else
        MsgBox mobjExpression.ValidityMessage(iValidityCode), _
          vbExclamation + vbOKOnly, Application.Name
      End If
      
      ' Set the invalid component to be the current component.
      If iValidityCode = giEXPRVALIDATION_EXPRTYPEMISMATCH Then
        sstrvComponents.SelectedItem = sstrvComponents.Nodes(ROOTKEY)
        sstrvComponents.SetFocus
      Else
        Set objBadComponent = mobjExpression.BadComponent
        If objBadComponent Is Nothing Then
          sstrvComponents.SelectedItem = sstrvComponents.Nodes(ROOTKEY)
          sstrvComponents.SetFocus
        Else
          For iLoop = 1 To mcolComponents.Count
            If mcolComponents.Item(iLoop) Is objBadComponent Then
              sstrvComponents.SelectedItem = sstrvComponents.Nodes(iLoop + 1)
              sstrvComponents.SetFocus
              Exit For
            End If
          Next iLoop
        End If
        Set objBadComponent = Nothing
      End If
    End If
  End If
  
  If fValid Then
    ' Check that the expression name is unique.
    recExprEdit.Index = "idxExprName"
    recExprEdit.Seek "=", mobjExpression.Name
      
    If Not recExprEdit.NoMatch Then
      Do While (Not recExprEdit.EOF) And fValid
        If recExprEdit!Name <> mobjExpression.Name Then
          Exit Do
        End If
        
        ' If there exists another expression with the same name then
        ' report this to the user.
        If (recExprEdit!ExprID <> mobjExpression.ExpressionID) And _
          (recExprEdit!ParentComponentID = 0) And _
          (Not recExprEdit!Deleted) And _
          (recExprEdit!TableID = mobjExpression.BaseTableID) And _
          (IIf(IsNull(recExprEdit!UtilityID), 0, recExprEdit!UtilityID) = mobjExpression.UtilityID) And _
          (recExprEdit!Type = mobjExpression.ExpressionType) Then
            
          MsgBox "A " & LCase(mobjExpression.ExpressionTypeName) & " called '" & mobjExpression.Name & "' already exists.", _
            vbExclamation + vbOKOnly, Application.Name
          sstrvComponents.SelectedItem = sstrvComponents.Nodes(ROOTKEY)
          txtExpressionName.SetFocus
          
          fValid = False
          Exit Do
        End If
          
        recExprEdit.MoveNext
      Loop
    End If
  End If
  
  'JPD 20040504 Fault 8599
  If fValid And (mobjExpression.ExpressionID > 0) Then
    fValid = Not mobjExpression.ContainsExpression(mobjExpression.ExpressionID)
    
    If Not fValid Then
     MsgBox "Invalid definition due to cyclic reference.", _
        vbExclamation + vbOKOnly, App.ProductName
      mfCancelled = True
    End If
  End If
  
'  If fValid And _
'    (mobjExpression.ExpressionID > 0) Then
'    ' If this expression is used in any other expressions, and these expressions are
'    ' used in column calculations, check that it does not reference
'    ' the column that is calculated, as this would cause recursion.
'    ReDim alngColumns(0)
'    mobjExpression.CalculatedColumnsThatUseThisExpression alngColumns
'
'    For iLoop = 1 To UBound(alngColumns)
'      fValid = Not mobjExpression.ExpressionContainsColumn(alngColumns(iLoop))
'      If Not fValid Then
'        ' Get the name of the column that uses the expression as its calculation.
'        recColEdit.Index = "idxColumnID"
'        recColEdit.Seek "=", Trim(Str(alngColumns(iLoop)))
'
'        If Not recColEdit.NoMatch Then
'          sColumnName = recColEdit!ColumnName
'
'          recTabEdit.Index = "idxTableID"
'          recTabEdit.Seek "=", recColEdit!TableID
'
'          If Not recTabEdit.NoMatch Then
'            sTableName = recTabEdit!TableName
'          Else
'            sTableName = "<unknown>"
'          End If
'        Else
'          sColumnName = "<unknown>"
'          sTableName = "<unknown>"
'        End If
'
'        MsgBox "This expression is used in the column calculation for the '" & sColumnName & "' column in the '" & sTableName & "' table," & vbCrLf & _
'          "but uses the column that is being calculated.", _
'          vbExclamation + vbOKOnly, App.ProductName
'        Exit For
'      End If
'    Next iLoop
'  End If
  
'  If fValid And _
'    (mobjExpression.ExpressionType = giEXPR_COLUMNCALCULATION) And _
'    (mobjExpression.ExpressionID > 0) Then
'    ' Check that any calulation expressions that are used in this expression
'    ' do not cause recursion.
'    For Each objComponent In mobjExpression.Components
'      If fValid And (objComponent.ComponentType = giCOMPONENT_CALCULATION) Then
'        Set objCalcExpr = New CExpression
'        objCalcExpr.ExpressionID = objComponent.Component.CalculationID
'        fValid = Not objCalcExpr.UsesCalculation(mobjExpression.ExpressionID)
'        Set objCalcExpr = Nothing
'      End If
'    Next objComponent
'    Set objComponent = Nothing
'
'    If Not fValid Then
'      MsgBox "This expression uses a calculation that uses this expression itself.", _
'        vbExclamation + vbOKOnly, App.ProductName
'    End If
'  End If

TidyUpAndExit:
  CheckExpression = fValid
  Exit Function

ErrorTrap:
  Err = False
  CheckExpression = False
  Exit Function
  
End Function

Private Sub Form_Terminate()

  Set mobjExpression = Nothing
  Set mcolComponents = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)

  ' Turn off form resize event handler
  Unhook Me.hWnd

End Sub

Private Sub optAccess_Click(Index As Integer)
  ' Update the expression object.
  'Select Case Index
  '  Case 0
  '    mobjExpression.Access = giACCESS_READWRITE
  '  Case 1
  '    mobjExpression.Access = giACCESS_READONLY
  '  Case 2
  '    mobjExpression.Access = giACCESS_HIDDEN
  'End Select
  
  Select Case Index
    Case 0: mobjExpression.Access = ACCESS_READWRITE
    Case 1: mobjExpression.Access = ACCESS_READONLY
    Case 2: mobjExpression.Access = ACCESS_HIDDEN
  End Select
  
  mfChanged = True

End Sub


'Private Sub sstbPopup_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
'  ' Process the tool click.
'  Select Case Tool.ID
'
'    Case "ID_Add"
'      cmdAddComponent_Click
'
'    Case "ID_Insert"
'      cmdInsertComponent_Click
'
'    Case "ID_Modify"
'      cmdModifyComponent_Click
'
'    Case "ID_Delete"
'      cmdDeleteComponent_Click
'
'    Case "ID_Rename"
'      sstrvComponents.StartLabelEdit
'
'  End Select
'
'End Sub

Private Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)

Dim iCount As Integer

  ' Process the tool click.
  Select Case Tool.Name
        
    Case "ID_Add"
      cmdAddComponent_Click
      
    Case "ID_Insert"
      cmdInsertComponent_Click
      
    Case "ID_Edit"
      cmdModifyComponent_Click
      
    Case "ID_Delete"
      cmdDeleteComponent_Click
      
    Case "ID_Rename"
      sstrvComponents.StartLabelEdit

    'Copy the component to the clipboard
    Case "ID_Copy"
        CopyComponents
    
    'Cut the component to the clipboard
    Case "ID_Cut"
        CutComponents

    'Paste the component from the clipboard
    Case "ID_Paste"
        PasteComponents

    'Move component one up the hierarchy
    Case "ID_MoveUp"
        MoveComponentUp

    'Move component one down the hierarchy
    Case "ID_MoveDown"
        MoveComponentDown

    'Expand all tree view nodes
    Case "ID_ExpandAll"
        For iCount = 1 To sstrvComponents.Nodes.Count
            sstrvComponents.Nodes(iCount).Expanded = True
            'JDM - 07/03/01 - Fault 1937 - Ensure scrollbars appear correctly
            sstrvComponents.Nodes(iCount).EnsureVisible
        Next iCount
        sstrvComponents.SelectedItem.EnsureVisible
    
    'Shrink all nodes in treeview
    Case "ID_ShrinkAll"
        For iCount = 1 To sstrvComponents.Nodes.Count
            sstrvComponents.Nodes(iCount).Expanded = False
        Next iCount
        sstrvComponents.SelectedItem.EnsureVisible

    'Enlarge font for all nodes
    Case "ID_ZoomIn"
        sstrvComponents.Font.Size = sstrvComponents.Font.Size + 2
        For iCount = 1 To sstrvComponents.Nodes.Count
            sstrvComponents.Nodes(iCount).Font.Size = sstrvComponents.Font.Size
        Next iCount
        sstrvComponents.SelectedItem.EnsureVisible

        Tool.Enabled = (sstrvComponents.Font.Size < 11)
        ActiveBar1.Tools("ID_ZoomOut").Enabled = True

    'Shrink font for all nodes
    Case "ID_ZoomOut"
        sstrvComponents.Font.Size = sstrvComponents.Font.Size - 2
        For iCount = 1 To sstrvComponents.Nodes.Count
            sstrvComponents.Nodes(iCount).Font.Size = sstrvComponents.Font.Size
        Next iCount
        sstrvComponents.SelectedItem.EnsureVisible

        Tool.Enabled = (sstrvComponents.Font.Size > 7)
        ActiveBar1.Tools("ID_ZoomIn").Enabled = True

    'Put all nodes to normal view
    Case "ID_ZoomNormal"
        sstrvComponents.Font.Size = 8
        For iCount = 1 To sstrvComponents.Nodes.Count
            sstrvComponents.Nodes(iCount).Font.Size = 8
        Next iCount
        sstrvComponents.SelectedItem.EnsureVisible

        ActiveBar1.Tools("ID_ZoomIn").Enabled = True
        ActiveBar1.Tools("ID_ZoomOut").Enabled = True

    'Add colour contouring (i.e. each level appears as a different colour)
    Case "ID_Colour"
        mbColoursOn = Not mbColoursOn
        Tool.Checked = mbColoursOn

        For iCount = 1 To sstrvComponents.Nodes.Count
            sstrvComponents.Nodes(iCount).ForeColor = GetNodeColour(sstrvComponents.Nodes(iCount).Level)
        Next iCount
  
    ' Send to printer
    Case "ID_OutputToPrinter"
      cmdPrint_Click
  
    ' Send to Clipboard
    Case "ID_OutputToClipboard"
      Clipboard.Clear
      mobjExpression.CopyExpressionToClipboard
    
    'Undo last action
    Case "ID_Undo"
      ExecuteUndo
  
  End Select

End Sub

Private Sub sstrvComponents_AfterLabelEdit(Cancel As SSActiveTreeView.SSReturnBoolean, NewString As String)
  'JPD 20030730 Fault 5587
  mfLabelEditing = False
  
  ' RH - Fault 1909 - Restore the default button
  cmdOk.Default = True
  
  ' Validate the entered label.
  If Len(NewString) = 0 Then
    MsgBox "You must enter a name.", vbExclamation + vbOKOnly, Application.Name
    Cancel = True
  Else
    SelectedComponent(sstrvComponents.SelectedItem).Component.Name = NewString
    mfChanged = True
  End If

End Sub

Private Sub sstrvComponents_BeforeLabelEdit(Cancel As SSActiveTreeView.SSReturnBoolean)
  
  ' RH - Fault 1909 - Stop the default button
  cmdOk.Default = False
  
  ' Only allow sub-expression labels to be edited.
  If sstrvComponents.SelectedItem.key = ROOTKEY Then
    Cancel = True
  Else
    If SelectedComponent(sstrvComponents.SelectedItem).ComponentType <> giCOMPONENT_EXPRESSION Then
      Cancel = True
    Else
      Cancel = Not mfModifiable
    End If
  End If

  'JPD 20030730 Fault 5587
  mfLabelEditing = Not Cancel

End Sub

Private Sub sstrvComponents_Collapse(Node As SSActiveTreeView.SSNode)
  
  ' Do not allow the root node to be collapsed.
  If Node.key = ROOTKEY Then
    Node.Expanded = True
  End If
  
  ' Set the expandednode property for the component
  If Node.Level > 1 Then
    mcolComponents(Node.key).ExpandedNode = False
  End If

  'JPD 20040507 Fault 7094
  sstrvComponents.ApproximateNodeCount = sstrvComponents.Nodes.Count

End Sub

Private Sub sstrvComponents_DblClick()

  If (sstrvComponents.Nodes.Count > 1) And (sstrvComponents.Nodes.Item(1).Selected = False) Then
    If SelectedComponent(sstrvComponents.SelectedItem).ComponentType <> giCOMPONENT_EXPRESSION And Me.cmdModifyComponent.Enabled = True Then
      cmdModifyComponent_Click
    End If
  End If

End Sub

Private Sub sstrvComponents_Expand(Node As SSActiveTreeView.SSNode)

  ' Set the expandednode property for the component
  If Node.Level > 1 Then
    mcolComponents(Node.key).ExpandedNode = True
  End If

  'JPD 20040507 Fault 7094
  sstrvComponents.ApproximateNodeCount = sstrvComponents.Nodes.Count

End Sub

Private Sub sstrvComponents_KeyPress(KeyAscii As Integer)
  ' JPD20021108 Fault 3287
  Dim objDummyComponent As CExprComponent
  Dim objOperatorDef As clsOperatorDef
  Dim objFunctionDef As clsFunctionDef
  Dim fFound As Boolean
  Dim IID As Integer
  Dim iComponentType As ExpressionComponentTypes
  
  'JPD 20030730 Fault 5587
  If mfLabelEditing Then
    Exit Sub
  End If
  
  If InStr(msShortcutKeys, UCase(Chr(KeyAscii))) > 0 Then
    ' Get the required operator/function.
    fFound = False
  
    For Each objOperatorDef In gobjOperatorDefs
      If InStr(objOperatorDef.ShortcutKeys, UCase(Chr(KeyAscii))) > 0 Then
        IID = objOperatorDef.ID
        iComponentType = giCOMPONENT_OPERATOR
        fFound = True
        Exit For
      End If
    Next objOperatorDef
    Set objOperatorDef = Nothing
    
    If Not fFound Then
      For Each objFunctionDef In gobjFunctionDefs
        If InStr(objFunctionDef.ShortcutKeys, UCase(Chr(KeyAscii))) > 0 Then
          IID = objFunctionDef.ID
          iComponentType = giCOMPONENT_FUNCTION
          fFound = True
          Exit For
        End If
      Next objFunctionDef
      Set objFunctionDef = Nothing
    End If
    
    If fFound Then
      CreateUndoView (giUNDO_ADD)
    
      AddComponent False, objDummyComponent, iComponentType, IID
    End If
  End If
  
End Sub


Private Sub sstrvComponents_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pop-up the menu.
  Dim fRenamable As Boolean
  
  ' Popup menu on right button.
  If Button = vbRightButton Then
  
    fRenamable = False
    If sstrvComponents.SelectedItem.key <> ROOTKEY Then
      If SelectedComponent(sstrvComponents.SelectedItem).ComponentType = giCOMPONENT_EXPRESSION Then
        fRenamable = True And mfModifiable
      End If
    End If
    
    With ActiveBar1.Bands("popup1")
  
      ' Enable/disable the required tools.
      .Tools("ID_Add").Enabled = cmdAddComponent.Enabled
      .Tools("ID_Insert").Enabled = cmdInsertComponent.Enabled
      .Tools("ID_Edit").Enabled = cmdModifyComponent.Enabled
      .Tools("ID_Delete").Enabled = cmdDeleteComponent.Enabled
      .Tools("ID_Rename").Enabled = fRenamable
      .Tools("ID_Cut").Enabled = mbCanCut And mfModifiable
      .Tools("ID_Copy").Enabled = mbCanCopy And mfModifiable
      .Tools("ID_Paste").Enabled = mbCanPaste And mfModifiable
      .Tools("ID_Copy").Enabled = mbCanCopy And mfModifiable
      .Tools("ID_MoveUp").Enabled = mbCanMoveUp And mfModifiable
      .Tools("ID_MoveDown").Enabled = mbCanMoveDown And mfModifiable
      .Tools("ID_Undo").Enabled = (miUndoLevel > 0)
      
      ' Set the undo text
      If miUndoLevel > 0 Then
        Select Case maUndoTypes(miUndoLevel)
          Case giUNDO_DELETE
            .Tools("ID_Undo").Caption = "Undo Delete"
          Case giUNDO_PASTE
            .Tools("ID_Undo").Caption = "Undo Paste"
          Case giUNDO_CUT
            .Tools("ID_Undo").Caption = "Undo Cut"
          Case giUNDO_ADD
            .Tools("ID_Undo").Caption = "Undo Add"
          Case giUNDO_INSERT
            .Tools("ID_Undo").Caption = "Undo Insert"
          Case giUNDO_MOVEUP
            .Tools("ID_Undo").Caption = "Undo Move Up"
          Case giUNDO_MOVEDOWN
            .Tools("ID_Undo").Caption = "Undo Move Down"
          Case giUNDO_EDIT
            .Tools("ID_Undo").Caption = "Undo Edit"
          Case giUNDO_RENAME
            .Tools("ID_Undo").Caption = "Undo Rename"
          Case Else
            .Tools("ID_Undo").Caption = "Undo"
        End Select
      Else
        .Tools("ID_Undo").Caption = "Undo"
      End If

      ' JDM - 15/03/01 - Fault 1934 - Allow user to expand / shrink nodes
      If mfModifiable Then
        ActiveBar1.RecalcLayout
        ActiveBar1.Bands("popup1").TrackPopup -1, -1
      Else
        ActiveBar1.Bands("PopupReadOnly").TrackPopup -1, -1
      End If

    End With
    
'    ' Configure the popup menu tools.
'    With sstbPopup
'      ' Enable/disable the required tools.
'      .Tools("ID_Add").Enabled = cmdAddComponent.Enabled
'      .Tools("ID_Insert").Enabled = cmdInsertComponent.Enabled
'      .Tools("ID_Modify").Enabled = cmdModifyComponent.Enabled
'      .Tools("ID_Delete").Enabled = cmdDeleteComponent.Enabled
'      .Tools("ID_Rename").Enabled = fRenamable
'
'      ' Display the popup menu.
'      .ToolBars("Popup").DockedStatus = ssFloating
'      .ToolBars("Popup").FloatingLeft = -10000
'      .Visible = True
'      .PopupMenu .Tools("ID_Popup")
'    End With
  
  End If
  
End Sub

Private Sub sstrvComponents_NodeClick(pcNode As SSActiveTreeView.SSNode)
  ' Enable/disable the command controls depending on the
  ' the current component selection.
  RefreshButtons

End Sub

Private Sub sstrvComponents_TopNodeChange(Node As SSActiveTreeView.SSNode)
  'JPD 20040507 Fault 7094
  sstrvComponents.ApproximateNodeCount = sstrvComponents.Nodes.Count

End Sub

Private Sub sstrvComponents_ValidateSelection(SelectionType As SSActiveTreeView.Constants_ValidateSelection, StartNode As SSActiveTreeView.SSNode, EndNode As SSActiveTreeView.SSNode, Cancel As SSActiveTreeView.SSReturnBoolean)
  'JPD 20040913 Fault 9127
'  If (SelectionType = ssatValidateToggle) _
'    And (sstrvComponents.SelectedNodes.Count = 1) Then Cancel = True

  ' JDM - 21/07/05 - Fault 10184 - Couldn't multi-select with the ctrl key on nodes
  If sstrvComponents.SelectedNodes.Count = 1 Then
    If SelectionType = ssatValidateToggle Then
      If sstrvComponents.SelectedNodes(1).key = StartNode.key Then
        Cancel = True
      End If
    End If
  End If

End Sub

Private Sub txtDescription_Change()
  ' Update the expression object.
  mobjExpression.Description = Trim(txtDescription.Text)

End Sub


Private Sub txtDescription_GotFocus()
  ' Select the entire contents of the textbox.
  UI.txtSelText

End Sub

Private Sub txtExpressionName_Change()
  Dim sValidatedName As String
  Dim iSelStart As Integer
  Dim iSelLen As Integer
  
  'JPD 20090102 Fault 13484
  sValidatedName = ValidateName(txtExpressionName.Text)
  
  If sValidatedName <> txtExpressionName.Text Then
    iSelStart = txtExpressionName.SelStart
    iSelLen = txtExpressionName.SelLength
    
    txtExpressionName.Text = sValidatedName
    
    txtExpressionName.SelStart = iSelStart
    txtExpressionName.SelLength = iSelLen
  End If
  
  ' Update the expression object.
  mobjExpression.Name = Trim(txtExpressionName.Text)
  
  ' Update the tree view display.
  If sstrvComponents.Nodes.Count > 0 Then
    sstrvComponents.Nodes(ROOTKEY).Text = mobjExpression.Name
  End If
  
End Sub

Private Sub PopulateTreeView()
  Dim objNode As SSActiveTreeView.SSNode
  Dim objComponent As CExprComponent

  ' Clear the treeview
  sstrvComponents.Nodes.Clear
  
  ' Clear the component collection and add the root expression.
  Set mcolComponents = Nothing
  Set mcolComponents = New Collection
  
  ' Add the expression root node.
  Set objNode = sstrvComponents.Nodes.Add(, , ROOTKEY, txtExpressionName.Text)
  With objNode
    .Font.Bold = True
    .Expanded = True
  End With
  Set objNode = Nothing
  
  ' Add nodes for each component in the expression.
  For Each objComponent In mobjExpression.Components
    AddComponentNode objComponent, ROOTKEY, objComponent.ExpandedNode, False
  Next objComponent
  Set objComponent = Nothing
  
  ' Select the root node.
  sstrvComponents.SelectedItem = sstrvComponents.Nodes(ROOTKEY)
  sstrvComponents.SelectedItem.Expanded = True
  
  ' Ensure the correct buttons are enabled for the selected node.
  RefreshButtons
  
End Sub

Public Property Get Cancelled() As Boolean
  ' Return the cancelled property.
  Cancelled = mfCancelled
  
End Property

Public Property Let ForcedReadOnly(ByVal pfNewValue As Boolean)
  ' Set the cancelled property.
  mfForcedReadOnly = pfNewValue
  
End Property

Public Property Let Cancelled(ByVal pfNewValue As Boolean)
  ' Set the cancelled property.
  mfCancelled = pfNewValue
  
End Property


Private Function CheckChanged() As Boolean
  'Have any of the controls datachanged flag been set ?
  Dim fChanged As Boolean
  Dim ctlScreenControl As Control
  
  If mfChanged Then
    CheckChanged = True
    Exit Function
  End If

  fChanged = False
  
  For Each ctlScreenControl In Me
    If TypeOf ctlScreenControl Is TextBox Or _
      TypeOf ctlScreenControl Is COA_Spinner Or _
      TypeOf ctlScreenControl Is CheckBox Or _
      TypeOf ctlScreenControl Is ComboBox Then
          
      If ctlScreenControl.DataChanged Then
        fChanged = True
        Exit For
      End If
    End If
  Next ctlScreenControl
  Set ctlScreenControl = Nothing
  
  CheckChanged = fChanged
  
End Function

Private Sub ClearChanged()
  ' Set all controls datachanged flags to false
  Dim ctlScreenControl As Control
  
  For Each ctlScreenControl In Me
    If TypeOf ctlScreenControl Is TextBox Or _
      TypeOf ctlScreenControl Is COA_Spinner Or _
      TypeOf ctlScreenControl Is CheckBox Or _
      TypeOf ctlScreenControl Is ComboBox Then
      
      If ctlScreenControl.DataChanged Then
        ctlScreenControl.DataChanged = False
      End If
    End If
  Next ctlScreenControl
  Set ctlScreenControl = Nothing
  
  mfChanged = False
  
End Sub

Private Function AddComponentNode(pobjComponent As CExprComponent, psParentNodeKey As String, pfExpanded As Boolean, pbFirstChild As Boolean) As String
  ' Populate the treeview with the given component's nodes.

  Dim sNodeKey As String
  Dim objNode As SSActiveTreeView.SSNode
  Dim objComponent As CExprComponent
  
  ' Create a unique key for the treeview node, and the associated
  ' object in the mcolComponents collection.
  sNodeKey = UniqueKey
  
  ' Add the component to the collection.
  mcolComponents.Add pobjComponent, sNodeKey

  ' Add the node to the treeview (Make first child if flag passed in as a parameter)
  If Not pbFirstChild Or sstrvComponents.Nodes(psParentNodeKey).Children = 0 Then
    Set objNode = sstrvComponents.Nodes.Add(psParentNodeKey, tvwChild, sNodeKey, pobjComponent.ComponentDescription)
  Else
    Set objNode = sstrvComponents.Nodes.Add(sstrvComponents.Nodes(psParentNodeKey).Child.key, tvwPrevious, sNodeKey, pobjComponent.ComponentDescription)
  End If

  ' If expanded node make sure it's visible
  If pfExpanded Then
    objNode.EnsureVisible
  End If

  'Set the colour of this particular node
  objNode.ForeColor = GetNodeColour(objNode.Level)

  ' Add sub-nodes for function parameters, and expression components.
  If pobjComponent.ComponentType = giCOMPONENT_FUNCTION Then
    For Each objComponent In pobjComponent.Component.Parameters
      AddComponentNode objComponent, sNodeKey, (pobjComponent.ExpandedNode And pfExpanded), False
    Next objComponent
    Set objComponent = Nothing
  ElseIf TypeOf pobjComponent.Component Is CExpression Then
    For Each objComponent In pobjComponent.Component.Components
      AddComponentNode objComponent, sNodeKey, (pobjComponent.ExpandedNode And pfExpanded), False
    Next objComponent
    Set objComponent = Nothing
  End If

  ' Disassociate the objNode variable.
  Set objNode = Nothing

  ' Return the key of the new node.
  AddComponentNode = sNodeKey
  
End Function

Private Function InsertComponentNode(pobjComponent As CExprComponent, psNodeKey As String, pfExpanded As Boolean, pbInsertBelow As Boolean) As String
  Dim objNode As SSActiveTreeView.SSNode
  Dim sNodeKey As String
  Dim objComponent As CExprComponent
  
  ' Create a unique key for the treeview node, and the associated
  ' object in the mcolComponents collection.
  sNodeKey = UniqueKey
  
  ' Add the component to the collection.
  mcolComponents.Add pobjComponent, sNodeKey

  ' Add the node to the treeview.
    If pbInsertBelow = True Then
        Set objNode = sstrvComponents.Nodes.Add(psNodeKey, tvwNext, sNodeKey, pobjComponent.ComponentDescription)
    Else
        Set objNode = sstrvComponents.Nodes.Add(psNodeKey, tvwPrevious, sNodeKey, pobjComponent.ComponentDescription)
    End If

  'Set the colour of this particular node
  objNode.ForeColor = GetNodeColour(objNode.Level)

  objNode.Expanded = pfExpanded
    
  ' Add sub-nodes for function parameters, and expression components.
  If pobjComponent.ComponentType = giCOMPONENT_FUNCTION Then
    For Each objComponent In pobjComponent.Component.Parameters
      AddComponentNode objComponent, sNodeKey, (pobjComponent.ExpandedNode And pfExpanded), False
    Next objComponent
    Set objComponent = Nothing
  ElseIf TypeOf pobjComponent.Component Is CExpression Then
    For Each objComponent In pobjComponent.Component.Components
      AddComponentNode objComponent, sNodeKey, (pobjComponent.ExpandedNode And pfExpanded), False
    Next objComponent
    Set objComponent = Nothing
  End If

  ' Disassociate the objNode variable.
  Set objNode = Nothing

  ' Return the key of the new node.
  InsertComponentNode = sNodeKey
  
End Function

Private Sub RemoveComponentNode(psNodeKey As String)
  Dim objNode As SSActiveTreeView.SSNode
  
  Set objNode = sstrvComponents.Nodes(psNodeKey)
  
  ' Remove any sub-nodes of the given node. This isn't strictly necessary as
  ' the removal of a parent node automatically removes the children. But we
  ' do need to remove the children from the collection.
  Do While Not objNode.Child Is Nothing
    RemoveComponentNode objNode.Child.key
  Loop

  ' Remove the component from the treeview and the collection.
  sstrvComponents.Nodes.Remove psNodeKey
  mcolComponents.Remove psNodeKey

  Set objNode = Nothing
  
End Sub

Private Sub txtExpressionName_GotFocus()
  ' Select all text upon focus.
  UI.txtSelText

End Sub

Private Function SelectedExpression(pobjNode As SSActiveTreeView.SSNode) As CExpression
  
  ' Return the parent expression of the treeview's selected component.
  Dim sParentKey As String
  
  ' Determine the key of the selected node's parent in the treeview.
  If pobjNode.key = ROOTKEY Then
    sParentKey = pobjNode.key
  Else
    sParentKey = pobjNode.Parent.key
  End If
    
  ' Get the selected component's parent expression.
  If sParentKey = ROOTKEY Then
    Set SelectedExpression = mobjExpression
  Else
    Set SelectedExpression = mcolComponents(sParentKey).Component
  End If

End Function

Private Function SelectedComponent(pobjNode As SSActiveTreeView.SSNode) As CExprComponent

  ' Return the treeview's selected component.
  If pobjNode.key = ROOTKEY Then
    If mcolComponents.Count = 0 Then
      Set SelectedComponent = New CExprComponent
    Else
      Set SelectedComponent = mcolComponents(pobjNode.Child.key)
    End If
  Else
    Set SelectedComponent = mcolComponents(pobjNode.key)
  End If

End Function

Private Function UniqueKey() As String
  ' Return a unique key for items in the treeview and component collection.
  Dim iLoop As Integer
  Dim iNextKey As Integer
  Dim iKey As Integer
  Dim sKey As String
  
  iNextKey = 1
  
  For iLoop = 1 To sstrvComponents.Nodes.Count
    sKey = sstrvComponents.Nodes(iLoop).key
    
    If sKey <> ROOTKEY Then
      iKey = val(sstrvComponents.Nodes(iLoop).key)
    
      If iKey >= iNextKey Then
        iNextKey = iKey + 1
      End If
    End If
  Next iLoop
  
  UniqueKey = Trim(Str(iNextKey))
  
End Function

Private Sub txtExpressionName_KeyPress(KeyAscii As Integer)
  ' Validate the character entered.
  KeyAscii = ValidNameChar(KeyAscii, txtExpressionName.SelStart)

End Sub

Public Function AddComponent(pbNewComponent As Boolean, _
  Optional pobjComponent As CExprComponent, _
  Optional piComponentType As ExpressionComponentTypes, _
  Optional piOpFuncID As Integer) As Boolean

  Dim sNewComponentKey As String
  Dim sParentExpressionKey As String
  Dim objParentExpression As CExpression
  Dim objNewComponent As CExprComponent
  Dim objCurrentComponent As CExprComponent
  Dim objPreviousComponent As CExprComponent
  Dim bMakeFirstChildNode As Boolean
  Dim iHiddenElements As Integer
  Dim bPasteBelow As Boolean
  
  ' If the root node is selected then we want to add the component to the
  ' root expression.
  If sstrvComponents.SelectedItem.key = ROOTKEY Then
    Set objParentExpression = mobjExpression
    sParentExpressionKey = ROOTKEY
  Else
    ' Get the selected component.
    Set objCurrentComponent = SelectedComponent(sstrvComponents.SelectedItem)

    ' Determine the parent expression of the selected component.
    If objCurrentComponent.ComponentType = giCOMPONENT_EXPRESSION Then
      Set objParentExpression = objCurrentComponent.Component
      sParentExpressionKey = sstrvComponents.SelectedItem.key
    Else
      Set objParentExpression = SelectedExpression(sstrvComponents.SelectedItem)
      sParentExpressionKey = sstrvComponents.SelectedItem.Parent.key
    End If

    Set objCurrentComponent = Nothing
  End If

  ' Get the expression to handle the addition of a new component.
  If pbNewComponent Then
    Set objNewComponent = objParentExpression.AddComponent
    bMakeFirstChildNode = False
  Else
    If pobjComponent Is Nothing Then
      Set objNewComponent = objParentExpression.AddOperatorFunctionComponent(piComponentType, piOpFuncID)
      bMakeFirstChildNode = False
    Else
      Set objPreviousComponent = SelectedComponent(sstrvComponents.SelectedItem)
      bPasteBelow = IIf(sParentExpressionKey = ROOTKEY, False, False)
      Set objNewComponent = objParentExpression.PasteComponent(pobjComponent, objPreviousComponent, bPasteBelow)
      bMakeFirstChildNode = True
    End If
  End If

  If Not objNewComponent Is Nothing Then
    ' Add the new component to the treeview.
    sNewComponentKey = AddComponentNode(objNewComponent, sParentExpressionKey, True, bMakeFirstChildNode)

    ' Select the new component.
    sstrvComponents.SelectedItem = sstrvComponents.Nodes(sNewComponentKey)
    
    'JDM - 04/12/01 - Fault 3124 - Strangely missing out the display of nodes. (I'm sure I did this before)
    DoEvents
   
    sstrvComponents.SelectedItem.Expanded = True
    sstrvComponents.Refresh

    mfChanged = True
    
    ' Ensure the command buttons are configured for the selected component.
    RefreshButtons
  End If

  ' Disassociate object variables.
  Set objParentExpression = Nothing
  Set objNewComponent = Nothing

End Function

Public Function GetNodeColour(piNodeLevel As Integer)

' Returns a different colour based on what node level is passed in

If mbColoursOn = False Then
    GetNodeColour = RGB(0, 0, 0)
Else
    Select Case piNodeLevel Mod 7
        'JDM - 07/03/01 - Fault 1943 - Fixed colour levels being messed up
        Case 0
            GetNodeColour = RGB(0, 15, 200)
        Case 1
            GetNodeColour = RGB(0, 0, 0)
        Case 2
            GetNodeColour = RGB(180, 0, 0)
        Case 3
            GetNodeColour = RGB(0, 125, 0)
        Case 4
            GetNodeColour = RGB(0, 0, 125)
        Case 5
            GetNodeColour = RGB(125, 125, 0)
        Case 6
            GetNodeColour = RGB(0, 125, 125)
        Case 7
            GetNodeColour = RGB(125, 0, 125)
    End Select

End If

End Function

Public Function InsertComponent(pbNewComponent As Boolean, Optional pobjComponent As CExprComponent, Optional lbInsertBelow As Boolean) As Boolean

  Dim objParentExpression As CExpression
  Dim objCurrentComponent As CExprComponent
  Dim objNewComponent As CExprComponent
  Dim sNextNodeKey As String
  Dim sNewComponentKey As String
  Dim bExpandedNode As Boolean
  
  ' Get the selected component,, and it's parent expression.
  Set objCurrentComponent = SelectedComponent(sstrvComponents.SelectedItem)
  Set objParentExpression = SelectedExpression(sstrvComponents.SelectedItem)
  
  sNextNodeKey = sstrvComponents.SelectedItem.key
  
  ' Instruct the parent expression to handle the insertion of a new component.
    If Not pbNewComponent Then
        Set objNewComponent = objParentExpression.PasteComponent(pobjComponent, objCurrentComponent, lbInsertBelow)
        bExpandedNode = pobjComponent.ExpandedNode
    Else
        Set objNewComponent = objParentExpression.InsertComponent(objCurrentComponent)
        lbInsertBelow = False
        bExpandedNode = True
    End If

  If Not objNewComponent Is Nothing Then
    ' Insert the new component in the treeview.
    sNewComponentKey = InsertComponentNode(objNewComponent, sNextNodeKey, bExpandedNode, lbInsertBelow)
    
    ' Select the new component.
    sstrvComponents.SelectedItem = sstrvComponents.Nodes(sNewComponentKey)
    sstrvComponents.SelectedItem.Expanded = bExpandedNode
    sstrvComponents.Refresh

    mfChanged = True
    
    ' Ensure the command buttons are configured for the selected component.
    RefreshButtons
  End If
  
  ' Disassociate object variables.
  Set objCurrentComponent = Nothing
  Set objParentExpression = Nothing
  Set objNewComponent = Nothing

End Function

Public Sub SetInitialExpandedNodes()

    Dim iCount As Integer
    Dim iLevelToExpandTo  As Integer

    iLevelToExpandTo = 2

    Select Case glngExpressionViewNodes

      ' Shrink all nodes
      Case EXPRESSIONBUILDER_NODESMINIMIZE
        For iCount = 1 To sstrvComponents.Nodes.Count
          sstrvComponents.Nodes(iCount).Expanded = False
        Next iCount

      ' Expand all nodes
      Case EXPRESSIONBUILDER_NODESEXPAND
        For iCount = 1 To sstrvComponents.Nodes.Count
          sstrvComponents.Nodes(iCount).Expanded = True
          sstrvComponents.Nodes(iCount).EnsureVisible
        Next iCount

      'Expand all specified levels
      Case EXPRESSIONBUILDER_NODESTOPLEVEL
        For iCount = 1 To sstrvComponents.Nodes.Count
          If sstrvComponents.Nodes(iCount).Level <= iLevelToExpandTo Then
            sstrvComponents.Nodes(iCount).Expanded = True
            sstrvComponents.Nodes(iCount).EnsureVisible
          Else
            sstrvComponents.Nodes(iCount).Expanded = False
          End If
        Next iCount

    End Select

    ' Ensure currently selected item is visible in the listbox
    sstrvComponents.SelectedItem.EnsureVisible

End Sub

Private Sub DeleteComponents()
  
  ' Deletes the selected nodes and their corresponding components
  On Error GoTo ErrorTrap:
  
  Dim iLoop As Integer
  Dim iOriginalNodeIndex As Integer
  Dim objComponent As CExprComponent
  Dim objExpression As CExpression
  Dim objNode As SSActiveTreeView.SSNode
  Dim bPositionNode As Boolean

  ' Save index of last selected item (used to correctly position pointer after deletion is complete)
  iOriginalNodeIndex = sstrvComponents.SelectedItem.Index - 1
  bPositionNode = False

  ' Loop through each selected node
  For Each objNode In sstrvComponents.Nodes
    If objNode.Selected = True Then
              
      'Get the selected component, and it's parent expression.
      Set objComponent = SelectedComponent(objNode)
      Set objExpression = SelectedExpression(objNode)
            
      ' Instruct the parent expression to handle the deletion of a component.
      If objExpression.DeleteComponent(objComponent) Then
        RemoveComponentNode objNode.key
        bPositionNode = True
      End If
      
    End If
  Next objNode
              
  ' Select the preceding visible component.
  If bPositionNode Then
    iOriginalNodeIndex = IIf(iOriginalNodeIndex > sstrvComponents.Nodes.Count, sstrvComponents.Nodes.Count, iOriginalNodeIndex)
    For iLoop = iOriginalNodeIndex To 1 Step -1
      If sstrvComponents.Nodes(iLoop).Visible Then
        Exit For
      End If
    Next iLoop
    
    ' JDM - Fault 10275 - 01/11/2005 - God knows how but nodeindex was somehow getting set to 0, I'm guessing its a faults with the
    '                                  selecteditem index. Made it happen once on my machine, but couldn't recreate,
    '                                  so I'm hoping this will do the necessary...
    ' Select node above last selected item
    If iLoop > 0 Then
      sstrvComponents.SelectedItem = sstrvComponents.Nodes(iLoop)
    End If

    ' Ensure the command buttons are configured for the selected component.
    RefreshButtons
    mfChanged = True
    
  End If
     
ErrorTrap:
     
  ' Disassociate object variables.
  Set objComponent = Nothing
  Set objExpression = Nothing

End Sub
' Place components on the undo collection
Private Sub CreateUndoView(ByVal iUndoType As UndoTypes)
  
  ' Set the current undo level
  miUndoLevel = miUndoLevel + 1
  
  ' Save the undo type
  ReDim Preserve maUndoTypes(miUndoLevel)
  maUndoTypes(UBound(maUndoTypes)) = iUndoType
  
  ' Save the current expression
  ReDim Preserve mcolUndoData(miUndoLevel)
  Set mcolUndoData(miUndoLevel) = mobjExpression.CopyComponent
  
End Sub

Private Sub ExecuteUndo()

  ' Set the current expression to be one from the undo array
  If miUndoLevel <= UBound(mcolUndoData) And miUndoLevel > 0 Then
    Set mobjExpression.Components = mcolUndoData(miUndoLevel).Components
    PopulateTreeView
    miUndoLevel = miUndoLevel - 1
  End If

End Sub

Private Sub Form_Resize()

  Dim lngWidth As Long

  'JPD 20030908 Fault 5756
  DisplayApplication

  fraDefinition(0).Move 100, 0, Me.ScaleWidth - 200, 1860
  
  lngWidth = fraDefinition(0).Width - (txtOwner.Left + 160)
  txtOwner.Width = IIf(lngWidth < 3000, lngWidth, 3000)

  fraDefinition(1).Move 100, 1900, Me.ScaleWidth - 200, Me.ScaleHeight - 2000

  With fraDefinition(1)
    fraButtons(0).Move .Width - (fraButtons(0).Width + 160), 240
    fraButtons(1).Move fraButtons(0).Left, .Height - (fraButtons(1).Height + 160)

    sstrvComponents.Move 150, 255, .Width - (fraButtons(0).Width + 450), .Height - 400
  End With

End Sub

