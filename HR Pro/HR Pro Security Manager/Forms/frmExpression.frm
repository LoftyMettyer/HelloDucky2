VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{1C203F10-95AD-11D0-A84B-00A0247B735B}#1.0#0"; "SSTree.ocx"
Begin VB.Form frmExpression 
   Caption         =   "Expression Definition"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
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
   HelpContextID   =   8035
   Icon            =   "frmExpression.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   8805
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraDefinition 
      Caption         =   "Definition :"
      Height          =   4305
      Index           =   1
      Left            =   100
      TabIndex        =   8
      Top             =   1890
      Width           =   8600
      Begin VB.Frame fraButtons 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   900
         Index           =   1
         Left            =   7200
         TabIndex        =   22
         Top             =   3240
         Width           =   1200
         Begin VB.CommandButton cmdCancel 
            Cancel          =   -1  'True
            Caption         =   "&Cancel"
            Height          =   400
            Left            =   0
            TabIndex        =   21
            Top             =   500
            Width           =   1200
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "&OK"
            Default         =   -1  'True
            Height          =   400
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   1200
         End
      End
      Begin VB.Frame fraButtons 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2915
         Index           =   0
         Left            =   7200
         TabIndex        =   13
         Top             =   240
         Width           =   1200
         Begin VB.CommandButton cmdAddComponent 
            Caption         =   "&Add..."
            Height          =   400
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   1200
         End
         Begin VB.CommandButton cmdInsertComponent 
            Caption         =   "&Insert..."
            Height          =   400
            Left            =   0
            TabIndex        =   15
            Top             =   500
            Width           =   1200
         End
         Begin VB.CommandButton cmdModifyComponent 
            Caption         =   "&Edit..."
            Height          =   400
            Left            =   0
            TabIndex        =   16
            Top             =   1000
            Width           =   1200
         End
         Begin VB.CommandButton cmdDeleteComponent 
            Caption         =   "&Delete"
            Height          =   400
            Left            =   0
            TabIndex        =   17
            Top             =   1500
            Width           =   1200
         End
         Begin VB.CommandButton cmdTest 
            Caption         =   "&Test"
            Height          =   400
            Left            =   0
            TabIndex        =   19
            Top             =   2500
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
            Height          =   400
            Left            =   0
            TabIndex        =   18
            Top             =   2000
            Width           =   1200
         End
      End
      Begin SSActiveTreeView.SSTree sstrvComponents 
         Height          =   3900
         Left            =   150
         TabIndex        =   6
         Top             =   255
         Width           =   6900
         _ExtentX        =   12171
         _ExtentY        =   6879
         _Version        =   65536
         NodeSelectionStyle=   2
         PictureAlignment=   0
         Style           =   6
         Indentation     =   315
         LoadStyleRoot   =   1
         AutoSearch      =   0   'False
         HideSelection   =   0   'False
         PictureBackgroundUseMask=   0   'False
         HasFont         =   -1  'True
         HasMouseIcon    =   0   'False
         HasPictureBackground=   0   'False
         ImageList       =   "(None)"
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
   End
   Begin VB.Frame fraDefinition 
      Height          =   1860
      Index           =   0
      Left            =   100
      TabIndex        =   7
      Top             =   0
      Width           =   8600
      Begin VB.OptionButton optAccess 
         Caption         =   "&Hidden"
         Height          =   345
         Index           =   2
         Left            =   5445
         TabIndex        =   5
         Top             =   1325
         Width           =   1200
      End
      Begin VB.OptionButton optAccess 
         Caption         =   "&Read Only"
         Height          =   315
         Index           =   1
         Left            =   5445
         TabIndex        =   4
         Top             =   1000
         Width           =   1200
      End
      Begin VB.OptionButton optAccess 
         Caption         =   "Read / &Write"
         Height          =   315
         Index           =   0
         Left            =   5445
         TabIndex        =   3
         Top             =   675
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.TextBox txtOwner 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   5445
         TabIndex        =   2
         Top             =   250
         Width           =   3000
      End
      Begin VB.TextBox txtDescription 
         Height          =   1000
         Left            =   1335
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   650
         Width           =   3090
      End
      Begin VB.TextBox txtExpressionName 
         Height          =   315
         Left            =   1335
         MaxLength       =   255
         TabIndex        =   0
         Top             =   250
         Width           =   3090
      End
      Begin VB.Label lblAccess 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Access :"
         Height          =   195
         Left            =   4700
         TabIndex        =   12
         Top             =   710
         Width           =   600
      End
      Begin VB.Label lblOwner 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Owner :"
         Height          =   195
         Left            =   4700
         TabIndex        =   11
         Top             =   310
         Width           =   585
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Description :"
         Height          =   195
         Left            =   195
         TabIndex        =   10
         Top             =   705
         Width           =   1125
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
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   30
      Top             =   1545
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
      Bands           =   "frmExpression.frx":000C
   End
End
Attribute VB_Name = "frmExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Expression definition variables.
Private mobjExpression As clsExprExpression
Private mcolComponents As Collection
Private mblnWasHidden As Boolean

'Indicating if the current user is denied edit/copy/delete privilages to the current expr.
Private mblnDenied As Boolean

' Form handling variables.
Private mfModifiable As Boolean
Private mfCancelled As Boolean
'Private mfChanged As Boolean
Private mfValid As Boolean

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

Private mcolUndoData() As clsExprExpression
Private maUndoTypes() As UndoTypes
Private miUndoLevel As Integer

' JPD20021108 Fault 3287
Private msShortcutKeys As String

' Form handling constants.
Const ROOTKEY = "EXPRESSION_ROOT"

Private mblnShownMessage As Boolean
Private mblnForcedHidden As Boolean
Private mblnHasChanged As Boolean

Private miOriginalReturnType As Integer

Private mvarUDFsRequired() As String
Private mfLabelEditing As Boolean

Private Function AccessState(lngExprID As Long) As String

 ' Returns access code for the current expression.
  Dim strSQL As String
  Dim rsTemp As ADODB.Recordset
  
  strSQL = "SELECT Access FROM ASRSysExpressions " & _
           "WHERE ExprID = " & lngExprID
           
  Set rsTemp = modExpression.OpenRecordset(strSQL, adOpenStatic, adLockReadOnly)

  With rsTemp
    If .RecordCount > 0 Then
      AccessState = !Access
    Else
      AccessState = ACCESS_READWRITE
    End If
  End With

End Function

Public Property Get Changed() As Boolean
  Changed = cmdOK.Enabled
End Property

Public Property Let Changed(blnChanged As Boolean)
  cmdOK.Enabled = blnChanged
End Property

Private Sub RemoveComponentNode(psNodeKey As String)
  Dim iLoop As Integer
  Dim objNode As SSActiveTreeView.SSNode

  Set objNode = sstrvComponents.Nodes(psNodeKey)

  ' Remove any sub-nodes of the given node. This isn't strictly necessary as
  ' the removal of a parent node automatically removes the children. But we
  ' do need to remove the children from the collection.
  Do While Not objNode.Child Is Nothing
    RemoveComponentNode objNode.Child.Key
  Loop

  ' Remove the component from the treeview and the collection.
  sstrvComponents.Nodes.Remove psNodeKey
  mcolComponents.Remove psNodeKey

  Set objNode = Nothing
    
End Sub

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

      ' JDM - 15/03/01 - Fault 1935 - Save the colour status of the expression
      mobjExpression.ViewInColour = mbColoursOn

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

Private Sub cmdAddComponent_Click()
  
  ' Place components on the undo collection
  CreateUndoView (giUNDO_ADD)

  AddComponent (True)

End Sub
Private Function SelectedExpression(pobjNode As SSActiveTreeView.SSNode) As clsExprExpression
  
  ' Return the parent expression of the treeview's selected component.
  Dim sParentKey As String
  
  ' Determine the key of the selected node's parent in the treeview.
  If pobjNode.Key = ROOTKEY Then
    sParentKey = pobjNode.Key
  Else
    sParentKey = pobjNode.Parent.Key
  End If
    
  ' Get the selected component's parent expression.
  If sParentKey = ROOTKEY Then
    Set SelectedExpression = mobjExpression
  Else
    Set SelectedExpression = mcolComponents(sParentKey).Component
  End If

End Function

Private Sub cmdCancel_Click()
  
  Dim intAnswer As Integer
  
  'JPD 20030730 Fault 5587
  mfLabelEditing = False
  
  ' Check if any changes have been made.
  If Me.Changed Then
    'intAnswer = MsgBox(" The " & LCase(ExpressionTypeName(mobjExpression.ExpressionType)) & " definition has changed.  Save changes ?", vbQuestion + vbYesNoCancel + vbDefaultButton1, App.ProductName)
    intAnswer = MsgBox("You have changed the current definition. Save changes ?", vbQuestion + vbYesNoCancel + vbDefaultButton1, App.ProductName)
    If intAnswer = vbYes Then
      Call cmdOK_Click
      Exit Sub
    ElseIf intAnswer = vbCancel Then
      Exit Sub
    End If
  End If
  
  ' Unload the form.
  mfCancelled = True
  
  ' Unload the form.
  Unload Me

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


Private Sub DeleteComponents()
  
  ' Deletes the selected nodes and their corresponding components
  On Error GoTo ErrorTrap
  
  Dim iLoop As Integer
  Dim iOriginalNodeIndex As Integer
  Dim objComponent As clsExprComponent
  Dim objExpression As clsExprExpression
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
        RemoveComponentNode objNode.Key
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
    
    ' Select node above last selected item
    If iLoop > 0 Then
      sstrvComponents.SelectedItem = sstrvComponents.Nodes(iLoop)
    End If

    ' Ensure the command buttons are configured for the selected component.
    RefreshButtons
    Me.Changed = True
    
  End If
     
  ' Check if there are hidden elements in the expression
  'SetAccessOptions HasHiddenComponents(mobjExpression.ExpressionID), Me.Expression.Access, mobjExpression.Owner
  SetAccessOptions HiddenElements, Me.Expression.Access, mobjExpression.Owner

ErrorTrap:

  ' Disassociate object variables.
  Set objComponent = Nothing
  Set objExpression = Nothing

End Sub

Private Sub cmdDeleteComponent_Click()
  
  ' Place components on the undo collection
  CreateUndoView (giUNDO_DELETE)

  ' Delete the selected components
  DeleteComponents

End Sub


Private Sub cmdInsertComponent_Click()

  ' Place components on the undo collection
  CreateUndoView (giUNDO_INSERT)

  InsertComponent (True)

End Sub

Private Sub cmdModifyComponent_Click()
  Dim objParentExpression As clsExprExpression
  Dim objCurrentComponent As clsExprComponent
  Dim objNewComponent As clsExprComponent
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
    sNextNodeKey = sstrvComponents.SelectedItem.Key

    ' Add the modified component to the treeview.
    sNewComponentKey = InsertComponentNode(objNewComponent, sNextNodeKey, True, False)

    ' Select the new component.
    sstrvComponents.SelectedItem = sstrvComponents.Nodes(sNewComponentKey)
    sstrvComponents.SelectedItem.Expanded = True

    ' Remove the old version of the component from the treeview.
    RemoveComponentNode sNextNodeKey

    Me.Changed = True
  
    ' Check if there are hidden elements in the expression
    'SetAccessOptions HasHiddenComponents(mobjExpression.ExpressionID), mobjExpression.Access, mobjExpression.Owner
    SetAccessOptions HiddenElements, mobjExpression.Access, mobjExpression.Owner

    ' Ensure the command buttons are configured for the selected component.
    RefreshButtons
  Else
    'JPD 20030819 Fault 6156. Need to check access options again as the user may have edited
    ' into a filter/calc/field component, changed the access of the filter/calc/field filter
    ' but then 'cancelled' out of the component edit screen.
    SetAccessOptions HiddenElements, mobjExpression.Access, mobjExpression.Owner, True
  End If

  ' Disassociate object variables.
  Set objCurrentComponent = Nothing
  Set objParentExpression = Nothing
  Set objNewComponent = Nothing

End Sub

Private Sub cmdOK_Click()
  
  ' Check the expression, and then unload the form.
  Dim bIsUsed As Boolean

  ' RH 30/01/01 - We dont want to update the timestamp, otherwise it will
  '               think the filter has never been updated by another user
  '               when it could well have been.
  mobjExpression.mfDontUpdateTimeStamp = True
  
  mfValid = False
  
  ' RH Fault 882 - Check for calcs used in the expression which have
  '                been deleted by another user before the current
  '                expression has been saved.
  If CheckForDeletedCalcs Then
    If CheckExpression Then
        
        'Only allow change of return type if expression is not used anywhere
        If Not mobjExpression.ExpressionID = 0 Then
          If Not mobjExpression.ReturnType = miOriginalReturnType Then
            If Not AllowChangeReturnType Then
              Exit Sub
            End If
          End If
        End If
        
        If mobjExpression.ExpressionID > 0 And LCase(mobjExpression.Owner) = LCase(gsUserName) And Me.optAccess(2).Value = True And mblnWasHidden = False Then
          If mobjExpression.ExpressionType = 10 Then    ' Its a Runtime Calc Expression
            If CheckCanMakeHidden("E", mobjExpression.ExpressionID, gsUserName, "Expression Validation") = False Then
              mobjExpression.mfDontUpdateTimeStamp = False
              Exit Sub
            End If
          Else                                          ' Its a Filter Expression
            If CheckCanMakeHidden("F", mobjExpression.ExpressionID, gsUserName, "Filter Validation") = False Then
              mobjExpression.mfDontUpdateTimeStamp = False
              Exit Sub
            End If
          End If
        End If
  
      mfValid = True
      
      Cancelled = False
      Unload Me
    Else
      ' Ensure the command buttons are configured for the selected component.
      RefreshButtons
    End If
  Else
    RefreshButtons
  End If

  mobjExpression.mfDontUpdateTimeStamp = False

End Sub

Private Function CheckForDeletedCalcs() As Boolean
  'TM20011001 Fault 2656
  'Adapted function to check for hidden elements by using the ValidComp
  'routine in clsExprComponent, by doing this we automatically validate
  'filters AND calcs.

  Dim iReturnCode As Integer
  Dim iLoop As Integer
  Dim sSQL As String
  Dim rsCheck As Recordset
  Dim objComponent As clsExprComponent

  On Error GoTo ErrorTrap

  ' Loop through each component in the expression
  For Each objComponent In mobjExpression.Components
    
    iReturnCode = ValidComponent(objComponent, False)
    
    Select Case iReturnCode
    Case 0:
    Case 1:
      MsgBox "The expression component '" & objComponent.ComponentDescription & _
              "' has been made hidden by another user." & vbCrLf & _
              "The expression will now be made hidden.", vbExclamation + vbOKOnly, App.Title
      SetExprAccess mobjExpression.ExpressionID, ACCESS_HIDDEN
      mobjExpression.Access = ACCESS_HIDDEN
      mblnHasChanged = True
      mblnForcedHidden = True
    Case 2:
      If Not isOwnerOfComp(objComponent) Then
        MsgBox "The expression component '" & objComponent.ComponentDescription & _
                "' has been made hidden by another user." & vbCrLf & _
                "Cannot make any modifications to this definition. " & vbCrLf _
                , vbExclamation + vbOKOnly, App.Title
        SetExprAccess mobjExpression.ExpressionID, ACCESS_HIDDEN
        mblnDenied = True
        Me.Cancelled = True
        Me.Changed = False
        Unload Me
        Exit Function

      Else
        MsgBox "The expression component '" & objComponent.ComponentDescription & _
                "' is hidden " & _
                "and cannot be added to another user's expression." _
                , vbExclamation + vbOKOnly, App.Title
        
      End If
    Case 3:
      MsgBox "The expression component '" & objComponent.ComponentDescription & "' has been deleted." & vbCrLf & _
             "Please remove it from the expression.", vbExclamation + vbOKOnly, App.Title
    Case 4:
      MsgBox "The expression contains a component which has been made hidden by another user." & vbCrLf & _
             "It will be removed from the expression.", vbExclamation + vbOKOnly, App.Title
      
      RemoveUnowned_HDComps mobjExpression
      PopulateTreeView
      mblnHasChanged = True
    End Select
    
    If (iReturnCode > 1) Then
'      MsgBox "The expression component '" & objComponent.ComponentDescription & "' has been deleted." & vbCrLf & _
'             "Please remove it from the expression.", vbExclamation + vbOKOnly, App.Title
'      ' Hilight the invalid calc in the treeview
      For iLoop = 1 To mobjExpression.Components.Count
        If mobjExpression.Components(iLoop) Is objComponent Then
          sstrvComponents.SelectedItem = sstrvComponents.Nodes(iLoop + 1)
          sstrvComponents.SetFocus
          Exit For
        End If
      Next iLoop
      Exit For
    End If
  Next objComponent
    
  Set rsCheck = Nothing
    
TidyUpAndExit:
  CheckForDeletedCalcs = (iReturnCode < 2)
  Exit Function

ErrorTrap:
  MsgBox "Error validating expression (checking for deleted calcs)." & vbCrLf & Err.Description, _
    vbExclamation + vbOKOnly, App.ProductName
  iReturnCode = 0
  Resume TidyUpAndExit
  
End Function

Private Sub SetAccessOptions(bHasHiddenComps As Boolean, _
  sAccessCurrentState As String, _
  sOwner As String, _
  Optional pvShowHiddenMessage As Variant)
  
' Checks the expression for hidden components and sets the enabled and selected
' properties of the Access option group.

  On Error GoTo ErrorTrap
  
  If Not bHasHiddenComps And sAccessCurrentState = ACCESS_READWRITE Then
    Me.optAccess(0).Value = True
    Me.optAccess(0).Enabled = (LCase(sOwner) = LCase(gsUserName))
    Me.optAccess(1).Enabled = (LCase(sOwner) = LCase(gsUserName))
    Me.optAccess(2).Enabled = (LCase(sOwner) = LCase(gsUserName))
    
  ElseIf Not bHasHiddenComps And sAccessCurrentState = ACCESS_READONLY Then
    Me.optAccess(1).Value = True
    Me.optAccess(0).Enabled = (LCase(sOwner) = LCase(gsUserName))
    Me.optAccess(1).Enabled = (LCase(sOwner) = LCase(gsUserName))
    Me.optAccess(2).Enabled = (LCase(sOwner) = LCase(gsUserName))
    
  ElseIf Not bHasHiddenComps And sAccessCurrentState = ACCESS_HIDDEN Then
    If (Not Me.optAccess(2).Enabled) And (LCase(sOwner) = LCase(gsUserName)) Then MsgBox "The expression no longer has to be hidden.", vbInformation + vbOKOnly, App.Title
    Me.optAccess(2).Value = True
    Me.optAccess(0).Enabled = (LCase(sOwner) = LCase(gsUserName))
    Me.optAccess(1).Enabled = (LCase(sOwner) = LCase(gsUserName))
    Me.optAccess(2).Enabled = (LCase(sOwner) = LCase(gsUserName))
      
  ElseIf bHasHiddenComps And sAccessCurrentState <> ACCESS_HIDDEN Then
    If Not IsMissing(pvShowHiddenMessage) Then
      If CBool(pvShowHiddenMessage) Then
        MsgBox "The selected component is hidden, the current expression will now be made hidden.", vbExclamation + vbOKOnly, App.Title
      End If
    End If
             
    Me.optAccess(2).Value = True
    Me.optAccess(0).Enabled = False
    Me.optAccess(1).Enabled = False
    Me.optAccess(2).Enabled = False
  
  ElseIf bHasHiddenComps And sAccessCurrentState = ACCESS_HIDDEN Then
    Me.optAccess(2).Value = True
    Me.optAccess(0).Enabled = False
    Me.optAccess(1).Enabled = False
    Me.optAccess(2).Enabled = False
    
  End If

TidyUpAndExit:
  Exit Sub

ErrorTrap:
  MsgBox "Error validating expression (checking for hidden components)." & vbCrLf & Err.Description, _
    vbExclamation + vbOKOnly, App.ProductName
  Resume TidyUpAndExit
  
End Sub


Private Function CheckExpression() As Boolean
  ' Check that the expression information is valid.
  On Error GoTo ErrorTrap

  Dim fValid As Boolean
  Dim fReadOnly As Boolean
  Dim fDeleted As Boolean
  Dim fTimeStampChanged As Boolean
  Dim fContinueSave As Boolean
  Dim fSaveAsNew As Boolean
  Dim iLoop As Integer
  Dim sSQL As String
  Dim sExpressionTypeName As String
  Dim sMBText As String
  Dim rsCheck As Recordset
  Dim objBadComponent As clsExprComponent
  Dim objComponent As clsExprComponent
  Dim objCalcExpr As clsExprExpression
  Dim iValidityCode As Integer
  Dim alngColumns() As Long
  Dim sAccess As String

  sExpressionTypeName = ExpressionTypeName(mobjExpression.ExpressionType)
  
  ' Check that there is an expression name.
  fValid = (Len(Trim(txtExpressionName.Text)) > 0)
  If Not fValid Then
    MsgBox "You must give this definition a name.", vbExclamation, Me.Caption
    sstrvComponents.SelectedItem = sstrvComponents.Nodes(ROOTKEY)
    mfCancelled = True
    txtExpressionName.SetFocus
  End If
  
  'JPD 20030818 Fault 6691 - check if the definition has been modified/deleted
  ' before trying to validate the expression components
  ' Check that the expression has not been modified by someone else.
  If fValid And (mobjExpression.ExpressionID > 0) Then
    fSaveAsNew = False
    fContinueSave = True
    
    sSQL = "SELECT convert(int, timestamp) AS timestamp, access, Username" & _
      " FROM ASRSysExpressions" & _
      " WHERE exprID = " & Trim(Str(mobjExpression.ExpressionID))
    Set rsCheck = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    fDeleted = (rsCheck.BOF And rsCheck.EOF)
    
    If fDeleted Then
      fTimeStampChanged = True
      fReadOnly = False
    Else
      fTimeStampChanged = (mobjExpression.Timestamp <> rsCheck!Timestamp)
      fReadOnly = (UCase(rsCheck!UserName) <> UCase(gsUserName) And rsCheck!Access <> ACCESS_READWRITE)
      
      'JPD 20030818 Fault 6693
      sAccess = rsCheck!Access
    End If
    
    rsCheck.Close
    Set rsCheck = Nothing
    
    If fTimeStampChanged Then
      If fDeleted Or fReadOnly Then
        ' Unable to overwrite existing definition
        If fDeleted Then
          sMBText = "This " & sExpressionTypeName & " has been deleted by another user."
        Else
          sMBText = "This " & sExpressionTypeName & " has been amended by another user and is now '" & AccessDescription(sAccess) & "'."
        End If
                      
        sMBText = sMBText & vbCrLf & "Save as a new definition ?"
        Select Case MsgBox(sMBText, vbExclamation + vbOKCancel, App.ProductName)
        Case vbOK         'save as new (but this may cause duplicate name message)
          fContinueSave = True
          fSaveAsNew = True
        Case vbCancel     'Do not save
          fContinueSave = False
        End Select
      Else
        ' Prompt to see if user should overwrite definition
        sMBText = "This " & sExpressionTypeName & " has been amended by another user. " & vbCrLf & _
          "Would you like to overwrite this definition?" & vbCrLf
        Select Case MsgBox(sMBText, vbExclamation + vbYesNoCancel, App.ProductName)
        Case vbYes        'overwrite existing definition and any changes
          fContinueSave = True
        Case vbNo         'save as new (but this may cause duplicate name message)
          fContinueSave = True
          fSaveAsNew = True
        Case vbCancel     'Do not save
          fContinueSave = False
        End Select
      End If
    End If
    
    If Not fContinueSave Then
      fValid = False
    ElseIf fSaveAsNew Then
      mobjExpression.ExpressionID = 0
      ' RH 30/01/01 - We are saving as new, but in order to save correctly, we
      '               need to reset the mfconstructed flag in mobjexpression
      '               otherwise it will reset the name property etc.
      mobjExpression.ResetConstructedFlag True
      mobjExpression.Owner = gsUserName
      txtOwner.Text = mobjExpression.Owner
      optAccess(0).Enabled = True
      optAccess(1).Enabled = True
      optAccess(2).Enabled = True
    End If
  End If
  
  ' Check that the expression is valid.
  If fValid Then
    mobjExpression.ResetConstructedFlag True
    mobjExpression.ConstructExpression
    iValidityCode = mobjExpression.ValidateExpression(True)
  
    fValid = (iValidityCode = giEXPRVALIDATION_NOERRORS)

    If Not fValid Then
      MsgBox mobjExpression.ValidityMessage(iValidityCode), _
        vbExclamation + vbOKOnly, App.ProductName
        mfCancelled = True

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
    
  If fValid And (mobjExpression.ExpressionType <> giEXPR_UTILRUNTIMEFILTER) Then
    ' Check that the expression name is unique for the table and expression type.
    sSQL = "SELECT *" & _
      " FROM ASRSysExpressions" & _
      " WHERE exprID <> " & Trim(Str(mobjExpression.ExpressionID)) & _
      " AND parentComponentID = 0" & _
      " AND name = '" & Replace(Trim(mobjExpression.Name), "'", "''") & "'" & _
      " AND TableID = " & Trim(Str(mobjExpression.BaseTableID)) & _
      " AND type = " & Trim(Str(mobjExpression.ExpressionType))
    Set rsCheck = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    With rsCheck
      fValid = .EOF And .BOF
      
      If Not fValid Then
        MsgBox "A " & LCase(sExpressionTypeName) & " called '" & mobjExpression.Name & "' already exists.", _
          vbExclamation + vbOKOnly, App.ProductName
        sstrvComponents.SelectedItem = sstrvComponents.Nodes(ROOTKEY)
        txtExpressionName.SetFocus
        mfCancelled = True
      End If
    
      .Close
    End With
    Set rsCheck = Nothing
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
    
TidyUpAndExit:
  'JPD 20031008 Fault 7081
  If Not fValid Then
    RefreshButtons
  End If
  
  CheckExpression = fValid
  Exit Function

ErrorTrap:
  MsgBox "Error validating expression." & vbCrLf & Err.Description, _
    vbExclamation + vbOKOnly, App.ProductName
  fValid = False
  Resume TidyUpAndExit
  
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
  
  Me.Changed = False
  
End Sub


Private Function InsertComponentNode(pobjComponent As clsExprComponent, psNodeKey As String, pfExpanded As Boolean, pbInsertBelow As Boolean) As String
  Dim objNode As SSActiveTreeView.SSNode
  Dim iLoop As Integer
  Dim sNodeKey As String
  Dim objComponent As clsExprComponent

  ' Create a unique key for the treeview node, and the associated
  ' object in the gcNodeComponents collection.
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

  ' If expanded node make sure it's visible
  If pfExpanded Then
    objNode.EnsureVisible
  End If

  ' Add sub-nodes for function parameters, and expression components.
  If pobjComponent.ComponentType = giCOMPONENT_FUNCTION Then
    For Each objComponent In pobjComponent.Component.Parameters
      AddComponentNode objComponent, sNodeKey, objComponent.ExpandedNode, False
    Next objComponent
    Set objComponent = Nothing
  ElseIf TypeOf pobjComponent.Component Is clsExprExpression Then
    For Each objComponent In pobjComponent.Component.Components
      AddComponentNode objComponent, sNodeKey, objComponent.ExpandedNode, False
    Next objComponent
    Set objComponent = Nothing
  End If

  ' Disassociate the objNode variable.
  Set objNode = Nothing

  ' Return the key of the new node.
  InsertComponentNode = sNodeKey
  
End Function

Private Sub PasteComponents()

  'Pastes the collection of clipboard components into the expression

  Dim objNewComponent As clsExprComponent

  ' Place components on the undo collection
  CreateUndoView (giUNDO_PASTE)

  For Each objNewComponent In mcolClipboard
    If sstrvComponents.SelectedItem.Key = ROOTKEY Then
      AddComponent False, objNewComponent
    Else
      If SelectedComponent(sstrvComponents.SelectedItem).ComponentType = giCOMPONENT_EXPRESSION Then
        AddComponent False, objNewComponent
      Else
        InsertComponent False, objNewComponent, True
      End If
    End If
  Next objNewComponent
  
  ' Ensure the command buttons are configured for the selected component.
  RefreshButtons

End Sub

Private Sub cmdPrint_Click()

  ' Print the expression1
  mobjExpression.PrintExpression

End Sub

Private Function AllowChangeReturnType() As Boolean

  Dim varWhereUsed As Variant
  Dim intCount As Integer
  Dim strUsage As String
  Dim strExpressionType As String

  If Me.Expression.ExpressionType = 10 Then
    strExpressionType = "CALCULATION"
  Else
    strExpressionType = "FILTER"
  End If

  On Error GoTo Prop_ERROR

  ' RH Show the user we are doing something...checking for usage could take a while
  Screen.MousePointer = vbHourglass

  Load frmDefProp

  With frmDefProp
    .SetChangeTypeError
    .Caption = Me.Caption
    .UtilName = Me.Expression.Name
    .PopulateUtil Me.Expression.ComponentType, Me.Expression.ExpressionID
    .CheckForUseage strExpressionType, Me.Expression.ExpressionID

    ' RH return the pointer to norma
    Screen.MousePointer = vbDefault

    If .UsageCount > 0 Then
      .Show vbModal
      AllowChangeReturnType = False
    Else
      AllowChangeReturnType = True
    End If
  End With

TidyUp:

  Unload frmDefProp
  Set frmDefProp = Nothing

  Exit Function

Prop_ERROR:

  Screen.MousePointer = vbDefault
  MsgBox "Error retrieving properties for this definition." & vbCrLf & "Please contact support stating : " & vbCrLf & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "Properties"
  Resume TidyUp

End Function

Private Sub cmdTest_Click()

  Dim rsTemp As ADODB.Recordset
  Dim strSQL As String
  Dim fOK As Boolean

  Dim strFilterCode As String
  Dim lngRecs As Long
  Dim strMBText As String

  ReDim mvarUDFsRequired(0)

  If CheckExpression = True Then

    fOK = mobjExpression.RuntimeFilterCode(strFilterCode, True, False)

    If fOK Then
      fOK = mobjExpression.UDFFilterCode(mvarUDFsRequired(), True, False)
    End If

    If fOK Then

      ' Create dynamic User defined functions
      UDFFunctions mvarUDFsRequired, True

      strSQL = "SELECT COUNT(*) FROM " & _
               mobjExpression.BaseTableName & _
               " WHERE ID IN (" & strFilterCode & ")"
      Set rsTemp = modExpression.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

      lngRecs = rsTemp(0).Value

      rsTemp.Close
      Set rsTemp = Nothing

      strMBText = "You have permission to view " & CStr(lngRecs) & " " & _
                  IIf(lngRecs <> 1, "records", "record") & _
                  " using this filter."

    End If

    strMBText = "Your filter is defined correctly." & vbCrLf & vbCrLf & strMBText
    MsgBox strMBText, vbInformation + vbOKOnly, "Filter Definition"

  End If

End Sub

Private Sub Form_Activate()

  ' JDM - 07/11/01 - Fault 3103 - Scrollbar is stuffed (bug within ActiveTreeBar control)
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
      ActiveBar1_Click ActiveBar1.Tools("ID_Cut")
    End If
    KeyCode = 0
    Shift = 0
  End If

  ' Copy component
  If (Shift And vbCtrlMask) And KeyCode = Asc("C") Then
    If mbCanCopy Then
      ActiveBar1_Click ActiveBar1.Tools("ID_Copy")
    End If
    KeyCode = 0
    Shift = 0
  End If

  ' Paste component
  If (Shift And vbCtrlMask) And KeyCode = Asc("V") Then
    If mbCanPaste Then
      ActiveBar1_Click ActiveBar1.Tools("ID_Paste")
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

  If KeyCode = vbKeyDown Then
    KeyCode = KeyCode
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
  Hook Me.hWnd, Me.Width, Me.Height, Screen.Width, Screen.Height
  
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
  
  fraButtons(0).BackColor = Me.BackColor
  fraButtons(1).BackColor = Me.BackColor

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  'JDM - 07/03/01 - Fault 1936 - Ask user if they wish to save changes
  Dim intAnswer As Integer
    
  'JPD 20030730 Fault 5587
  mfLabelEditing = False
    
  If UnloadMode <> vbFormCode Then
     
    'Check if any changes have been made.
    If Me.Changed Then
      intAnswer = MsgBox("You have changed the current definition. Save changes ?", vbQuestion + vbYesNoCancel + vbDefaultButton1, App.ProductName)
      If intAnswer = vbYes Then
          Call cmdOK_Click
          If Me.Cancelled Or Not mfValid Then Cancel = 1
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

Public Property Get Cancelled() As Boolean
  ' Return the cancelled property.
  Cancelled = mfCancelled
  
End Property

Public Property Let Cancelled(ByVal pfNewValue As Boolean)
  ' Set the cancelled property.
  mfCancelled = pfNewValue
  
End Property

Public Property Get Expression() As clsExprExpression
  ' Return the expression that is being editted.
  Set Expression = mobjExpression
  
End Property

Public Property Set Expression(pobjExpr As clsExprExpression)
  
  Dim lngHiddenComponents As Boolean
  Dim sAccessState As String
  Dim iValidationCode As Integer
  
  Screen.MousePointer = vbHourglass
  
  ' Set the expression that is being editted.
  Set mobjExpression = pobjExpr
  
  ' Update the screen controls with the expression's properties.
  With mobjExpression
    
    iValidationCode = ValidateExpr(mobjExpression, True)
    'Decode iValidationCode using the above return codes.
    'Codes:
    '
    '4 -->  Expression is owned by current user but it contains hidden
    '       components owned by another user.
    '3 -->  Expression is owned by current user has deleted components
    '       in the definition.
    '2 -->  Expression is NOT owned by current user but it contains hidden
    '       components, therefore should now be hidden to all but owner.
    '1 -->  Expression is owned by current user but it contains hidden
    '       components owned by the current user and is not already hidden,
    '       therefore the expression should made hidden.
    '0 -->  Expression has the correct access defined for the components
    '       within the definition, no message required!

    Select Case iValidationCode
    Case 0:
      'nothing required
    Case 1:
      SetExprAccess .ExpressionID, ACCESS_HIDDEN
      .Access = ACCESS_HIDDEN
      mblnHasChanged = True
      mblnForcedHidden = True
    Case 2:
      SetExprAccess .ExpressionID, ACCESS_HIDDEN
      mblnDenied = True
      Me.Cancelled = True
    Case 3:
      'nothing required
    Case 4:
      RemoveUnowned_HDComps mobjExpression
      mblnForcedHidden = False
      mblnHasChanged = True
    End Select

''********************************************************************************
'' 'TM20010801 Fault 2617                                                        *
'' 'Displys message to user if the expression should be hidden and the owner is  *
'' 'not the current user. If the expression does contain hidden elements then    *
'' 'the Access property is set to hidden, if the current user is not the owner   *
'' 'of the selected expression the form is cancelled.                            *
''********************************************************************************
'    sAccessState = AccessState(.ExpressionID)
'    If .ExpressionID <> 0 Then lngHiddenComponents = HasHiddenComponents(.ExpressionID)
'    If lngHiddenComponents And sAccessState <> "HD" And LCase(.Owner) <> LCase(gsUserName) Then
'      MsgBox "The selected expression contains hidden components and is owned by another user." & vbCrLf & vbCrLf & "The expression will now be made hidden.", vbExclamation + vbOKOnly, App.Title
'      sAccessState = "HD"
'      Me.Cancelled = True
'    ElseIf lngHiddenComponents And sAccessState <> "HD" _
'            And LCase(.Owner) = LCase(gsUserName) And .ActionType <> edtdelete Then
'      MsgBox "The selected expression contains hidden components!" & vbCrLf & vbCrLf & "The expression will now be made hidden.", vbExclamation + vbOKOnly, App.Title
'      sAccessState = "HD"
'    End If
    
    If Not Me.Cancelled Then
      'TM20010801 Fault 2617
      'If a copy is being made of the selected expression then the owner is changed to the current user.
      If .ActionType = edtCopy Then
        .Owner = gsUserName
      End If
      
      ' Check if the user can modify the expression.
      mfModifiable = (LCase(gsUserName) = LCase(.Owner)) Or _
        (.Access = ACCESS_READWRITE) Or _
        (gfCurrentUserIsSysSecMgr)
            
      ' Is it already hidden ? if so, set a flag
      mblnWasHidden = (.Access = ACCESS_HIDDEN)
      
      ' Treat the filter as read-only if the user does not have permission to edit them.
'      If mfModifiable And .ExpressionType = giEXPR_RUNTIMEFILTER Then
'        mfModifiable = SystemPermission("FILTERS", "EDIT")
'      End If
'
'      ' Treat the filter as read-only if the user does not have permission to edit them.
'      If mfModifiable And .ExpressionType = giEXPR_STATICFILTER Then
'        mfModifiable = SystemPermission("FILTERS", "EDIT")
'      End If
'
'      ' Treat the calculation as read-only if the user does not have permission to edit them.
'      If mfModifiable And .ExpressionType = giEXPR_RUNTIMECALCULATION Then
'        mfModifiable = SystemPermission("CALCULATIONS", "EDIT")
'      End If
    
      ' Enable/disable controls as required.
      ConfigureScreen
      
      'TM20010801 Fault 2617
      'Only try to change the access option group settings if we are not deleting the expression.
      'Note: The form would already of cancelled if the current user was not the owner of the expression,
      '      therefore not allowing a 'non-owner' to delete the expression.
      If .ActionType <> edtDelete Then
        'Set the properties of the Access option group.
        'SetAccessOptions HasHiddenComponents(mobjExpression.ExpressionID), sAccessState, .Owner
        SetAccessOptions HiddenElements, sAccessState, .Owner
      End If
      
      txtExpressionName.Text = .Name
      txtDescription.Text = .Description
      txtOwner.Text = .Owner
      
      .ValidateExpression True
      miOriginalReturnType = .ReturnType
      .ResetConstructedFlag True
      
      'optAccess(.Access).Value = True
      Select Case .Access
      Case ACCESS_READWRITE: optAccess(0).Value = True
      Case ACCESS_READONLY: optAccess(1).Value = True
      Case ACCESS_HIDDEN: optAccess(2).Value = True
      End Select
    
      ' Populate the treeview with the expression definition.
      PopulateTreeView
    
      ' JDM - 19/03/01 - Set the initially expanded/shrunk nodes
      SetInitialExpandedNodes
      
      ' JDM - 28/08/01 - Fault 2725 - Copy expression should make OK button enabled
      If .ActionType = edtCopy Or mblnHasChanged Then
        ' ie. if we are copying an existing expression.
        Me.Changed = True
      Else
        ClearChanged
      End If
    End If
  End With

  Screen.MousePointer = vbDefault
  
End Property

Private Sub PopulateTreeView()
  Dim objNode As SSActiveTreeView.SSNode
  Dim objComponent As clsExprComponent
  
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
    AddComponentNode objComponent, ROOTKEY, True, False
  Next objComponent
  Set objComponent = Nothing
  
  ' Select the root node.
  sstrvComponents.SelectedItem = sstrvComponents.Nodes(ROOTKEY)
  sstrvComponents.SelectedItem.Expanded = True
  
  ' Ensure the correct buttons are enabled for the selected node.
  RefreshButtons
  
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
      If objNode.Key = ROOTKEY Then
        mbCanInsert = False
        mbCanEdit = False
        mbCanDelete = False
        mbCanCut = False
        mbCanPaste = (mcolClipboard.Count > 0) And mbCanPaste
        mbCanCopy = False
        mbCanMoveDown = False
        mbCanMoveUp = False
      Else
        Select Case mcolComponents.Item(objNode.Key).ComponentType
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
            mbCanEdit = Not (mcolComponents.Item(objNode.Parent.Key).ComponentType = giCOMPONENT_FUNCTION) And mbCanEdit
            mbCanInsert = False
            mbCanDelete = Not (mcolComponents.Item(objNode.Parent.Key).ComponentType = giCOMPONENT_FUNCTION) And mbCanDelete
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


Private Function AddComponentNode(pobjComponent As clsExprComponent, psParentNodeKey As String, pfExpanded As Boolean, pbFirstChild As Boolean) As String
  ' Populate the treeview with the given component's nodes.
  Dim iLoop As Integer
  Dim sNodeKey As String
  Dim objNode As SSActiveTreeView.SSNode
  Dim objComponent As clsExprComponent
  
  ' Create a unique key for the treeview node, and the associated
  ' object in the gcNodeComponents collection.
  sNodeKey = UniqueKey
  
  ' Add the component to the collection.
  mcolComponents.Add pobjComponent, sNodeKey

  ' Add the node to the treeview (Make first child if flag passed in as a parameter)
  If Not pbFirstChild Or sstrvComponents.Nodes(psParentNodeKey).Children = 0 Then
    Set objNode = sstrvComponents.Nodes.Add(psParentNodeKey, tvwChild, sNodeKey, pobjComponent.ComponentDescription)
  Else
    Set objNode = sstrvComponents.Nodes.Add(sstrvComponents.Nodes(psParentNodeKey).Child.Key, tvwPrevious, sNodeKey, pobjComponent.ComponentDescription)
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
      AddComponentNode objComponent, sNodeKey, pobjComponent.ExpandedNode, False
    Next objComponent
    Set objComponent = Nothing
  ElseIf TypeOf pobjComponent.Component Is clsExprExpression Then
    For Each objComponent In pobjComponent.Component.Components
      AddComponentNode objComponent, sNodeKey, pobjComponent.ExpandedNode, False
    Next objComponent
    Set objComponent = Nothing
  End If

  ' Disassociate the objNode variable.
  Set objNode = Nothing

  ' Return the key of the new node.
  AddComponentNode = sNodeKey
  
End Function

Private Function UniqueKey() As String
  ' Return a unique key for items in the treeview and component collection.
  Dim iKey As Integer
  Dim iLoop As Integer
  Dim iNextKey As Integer
  Dim sKey As String
  
  iNextKey = 1
  
  For iLoop = 1 To sstrvComponents.Nodes.Count
    sKey = sstrvComponents.Nodes(iLoop).Key
    
    If sKey <> ROOTKEY Then
      iKey = Val(sstrvComponents.Nodes(iLoop).Key)
    
      If iKey >= iNextKey Then
        iNextKey = iKey + 1
      End If
    End If
  Next iLoop
  
  UniqueKey = Trim(Str(iNextKey))
  
End Function

Private Sub ConfigureScreen()
  ' Configure the screen controls.
  Dim fUserIsCreator As Integer
  
  ' Configure the screen controls depending on the type of
  ' selection being made.
  
  'JPD 20030911 Fault 6359
  ' RH 17/11/00 - request from pjc - dont say definition after the expr type in the caption
  ' Me.Caption = ExpressionTypeName(mobjExpression.ExpressionType) & " Definition"
  'Me.Caption = ExpressionTypeName(mobjExpression.ExpressionType)
  If mobjExpression.ExpressionType = giEXPR_RUNTIMECALCULATION Then
    Me.Caption = "Calculation Definition"
  Else
    Me.Caption = ExpressionTypeName(mobjExpression.ExpressionType) & " Definition"
  End If
  
  'TM20010926 Fault
  ' Only allow the access permission to be changed by the original creator.
  fUserIsCreator = (LCase(gsUserName) = LCase(mobjExpression.Owner))
  optAccess(0).Enabled = fUserIsCreator And Not mblnForcedHidden
  optAccess(1).Enabled = fUserIsCreator And Not mblnForcedHidden
  optAccess(2).Enabled = fUserIsCreator And Not mblnForcedHidden
    
  ' Do not allow the user to change the expression if they are not allowed to.
  If Not mfModifiable Then
    ControlsDisableAll Me

    ' JDM - 15/03/01 - Fault 1934 - Allow user to expand / shrink nodes
    sstrvComponents.Enabled = True
    ActiveBar1.ForeColor = RGB(0, 0, 0)

    ' JDM - 15/03/01 - Fault 2004 - Enable the test and print button
    cmdTest.Enabled = True
    cmdPrint.Enabled = True

  End If

  ' JDM - 26/11/01 - Fault 3196 - Restore last view
  Select Case GetSystemSetting("ExpressionBuilder", "ViewColours", EXPRESSIONBUILDER_COLOUROFF)
    Case EXPRESSIONBUILDER_COLOUROFF
      mbColoursOn = False
    Case EXPRESSIONBUILDER_COLOURON
      mbColoursOn = True
    Case Else
      mbColoursOn = False
  End Select
  
  ' Display the appropriate form icon.
 Select Case mobjExpression.ExpressionType
  Case giEXPR_COLUMNCALCULATION
  Case giEXPR_GOTFOCUS
  Case giEXPR_RECORDVALIDATION
  Case giEXPR_STATICFILTER
  Case giEXPR_RECORDDESCRIPTION
  Case giEXPR_VIEWFILTER
  Case giEXPR_RUNTIMECALCULATION
    Me.HelpContextID = 8025
  Case giEXPR_RUNTIMEFILTER
    cmdTest.Visible = True
    Me.HelpContextID = 8026
  Case giEXPR_UTILRUNTIMEFILTER
    lblOwner.Visible = False
    txtOwner.Visible = False
    lblAccess.Visible = False
    optAccess(0).Visible = False
    optAccess(1).Visible = False
    optAccess(2).Visible = False
    
  End Select
  
  ' Get rid of the icon off the form
  Me.Icon = Nothing
  SetWindowLong Me.hWnd, GWL_EXSTYLE, WS_EX_WINDOWEDGE Or WS_EX_APPWINDOW Or WS_EX_DLGMODALFRAME
  
End Sub

Private Sub Form_Terminate()

  ' Disassociate object variables.
  Set mobjExpression = Nothing
  Set mcolComponents = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)

  ' Turn off form resize event handler
  Unhook Me.hWnd

End Sub

Private Sub optAccess_Click(Index As Integer)
  ' Update the expression object.
  mobjExpression.Access = Choose((Index + 1), ACCESS_READWRITE, ACCESS_READONLY, ACCESS_HIDDEN)
  Me.Changed = True

End Sub
Private Sub sstrvComponents_AfterLabelEdit(Cancel As SSActiveTreeView.SSReturnBoolean, NewString As String)
  'JPD 20030730 Fault 5587
  mfLabelEditing = False
  
  ' RH - Fault 1909 - Put the default button back on
  cmdOK.Default = True
  
  ' Validate the entered label.
  If Len(NewString) = 0 Then
    MsgBox "You must enter a name.", vbExclamation + vbOKOnly, App.ProductName
    Cancel = True
  Else
    SelectedComponent(sstrvComponents.SelectedItem).Component.Name = NewString
    Me.Changed = True
  End If

End Sub

Private Function SelectedComponent(pobjNode As SSActiveTreeView.SSNode) As clsExprComponent

  ' Return the treeview's selected component.
  If pobjNode.Key = ROOTKEY Then
    If mcolComponents.Count = 0 Then
      Set SelectedComponent = New clsExprComponent
    Else
      Set SelectedComponent = mcolComponents(pobjNode.Child.Key)
    End If
  Else
    Set SelectedComponent = mcolComponents(pobjNode.Key)
  End If

End Function

Private Sub sstrvComponents_BeforeLabelEdit(Cancel As SSActiveTreeView.SSReturnBoolean)
  
  ' Place components on the undo collection
  CreateUndoView (giUNDO_RENAME)
  
  ' RH - Fault 1909 - Remove the default button
  cmdOK.Default = False
  
  ' Only allow sub-expression labels to be edited.
  If sstrvComponents.SelectedItem.Key = ROOTKEY Then
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
  If Node.Key = ROOTKEY Then
    Node.Expanded = True
  End If

  ' Set the expandednode property for the component
  If Node.Level > 1 Then
    mcolComponents(Node.Key).ExpandedNode = False
  End If

  'TM20020828 Fault 3834
  ' JDM - 07/11/01 - Fault 3103 - Scrollbar is stuffed (bug within ActiveTreeBar control)
  'sstrvComponents.ApproximateNodeCount = sstrvComponents.Nodes.Count - 1
  sstrvComponents.ApproximateNodeCount = sstrvComponents.Nodes.Count

End Sub

Private Sub sstrvComponents_DblClick()
  
  ' RH - Runtime error when doubleclicking treeview with no items in it.
  If (sstrvComponents.Nodes.Count > 1) And (sstrvComponents.Nodes.Item(1).Selected = False) Then
    If SelectedComponent(sstrvComponents.SelectedItem).ComponentType <> giCOMPONENT_EXPRESSION And Me.cmdModifyComponent.Enabled = True Then
      cmdModifyComponent_Click
    End If
  End If
  
End Sub

Private Sub sstrvComponents_Expand(Node As SSActiveTreeView.SSNode)

' Set the expandednode property for the component
If Node.Level > 1 Then
  mcolComponents(Node.Key).ExpandedNode = True
End If

  'TM20020828 Fault 3834
  ' JDM - 07/11/01 - Fault 3103 - Scrollbar is stuffed (bug within ActiveTreeBar control)
  'sstrvComponents.ApproximateNodeCount = sstrvComponents.Nodes.Count - 1
  sstrvComponents.ApproximateNodeCount = sstrvComponents.Nodes.Count

End Sub

Private Sub sstrvComponents_KeyDown(KeyCode As Integer, Shift As Integer)

  ' Control = crap, have to force it to do this...
  Form_KeyDown KeyCode, Shift

End Sub

Private Sub sstrvComponents_KeyPress(KeyAscii As Integer)
  ' JPD20021108 Fault 3287
  Dim objDummyComponent As clsExprComponent
  Dim objOperatorDef As clsOperatorDef
  Dim objFunctionDef As clsFunctionDef
  Dim fFound As Boolean
  Dim iID As Integer
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
        iID = objOperatorDef.ID
        iComponentType = giCOMPONENT_OPERATOR
        fFound = True
        Exit For
      End If
    Next objOperatorDef
    Set objOperatorDef = Nothing
    
    If Not fFound Then
      For Each objFunctionDef In gobjFunctionDefs
        If InStr(objFunctionDef.ShortcutKeys, UCase(Chr(KeyAscii))) > 0 Then
          iID = objFunctionDef.ID
          iComponentType = giCOMPONENT_FUNCTION
          fFound = True
          Exit For
        End If
      Next objFunctionDef
      Set objFunctionDef = Nothing
    End If
    
    If fFound Then
      CreateUndoView (giUNDO_ADD)
    
      AddComponent False, objDummyComponent, iComponentType, iID
    End If
  End If
  
End Sub

Private Sub sstrvComponents_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Pop-up the menu.
  Dim fRenamable As Boolean

  ' Popup menu on right button.
  If Button = vbRightButton Then
    
    fRenamable = False
    If sstrvComponents.SelectedItem.Key <> ROOTKEY Then
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

  End If

End Sub


Private Sub sstrvComponents_NodeClick(Node As SSActiveTreeView.SSNode)
  
  ' Enable/disable the command controls depending on the
  ' the current compoennt selection.
  RefreshButtons

End Sub

Private Sub sstrvComponents_TopNodeChange(Node As SSActiveTreeView.SSNode)

  ' JDM - 07/11/01 - Fault 3103 - Scrollbar is stuffed (bug within ActiveTreeBar control)
'  sstrvComponents.ApproximateNodeCount = sstrvComponents.Nodes.Count - 1
  sstrvComponents.ApproximateNodeCount = sstrvComponents.Nodes.Count

End Sub

Private Sub sstrvComponents_ValidateSelection(SelectionType As SSActiveTreeView.Constants_ValidateSelection, StartNode As SSActiveTreeView.SSNode, EndNode As SSActiveTreeView.SSNode, Cancel As SSActiveTreeView.SSReturnBoolean)
  'JPD 20040913 Fault 9129
'  If (SelectionType = ssatValidateToggle) _
'    And (sstrvComponents.SelectedNodes.Count = 1) Then Cancel = True

  ' JDM - 21/07/05 - Fault 10183 - Couldn't multi-select with the ctrl key on nodes
  If sstrvComponents.SelectedNodes.Count = 1 Then
    If SelectionType = ssatValidateToggle Then
      If sstrvComponents.SelectedNodes(1).Key = StartNode.Key Then
        Cancel = True
      End If
    End If
  End If

End Sub

Private Sub txtDescription_Change()
  ' Update the expression object.
  mobjExpression.Description = Trim(txtDescription.Text)
  Me.Changed = True

End Sub


Private Sub txtDescription_GotFocus()
  ' Select the entire contents of the textbox.
  UI.txtSelText
  cmdOK.Default = False

End Sub

Private Sub txtDescription_LostFocus()
  cmdOK.Default = True
End Sub

Private Sub txtExpressionName_Change()
  ' Update the expression object.
  mobjExpression.Name = Trim(txtExpressionName.Text)
  
  ' Update the tree view display.
  If sstrvComponents.Nodes.Count > 0 Then
    sstrvComponents.Nodes(ROOTKEY).Text = mobjExpression.Name
  End If

  Me.Changed = True

End Sub

Private Sub txtExpressionName_GotFocus()
  ' Select the entire contents of the textbox.
  UI.txtSelText

End Sub

Private Sub txtExpressionName_KeyPress(KeyAscii As Integer)
  ' Validate the character entered.
  KeyAscii = ValidNameChar(KeyAscii, txtExpressionName.SelStart)

End Sub

Public Function AddComponent(pbNewComponent As Boolean, _
  Optional pobjComponent As clsExprComponent, _
  Optional piComponentType As ExpressionComponentTypes, _
  Optional piOpFuncID As Integer) As Boolean

  Dim sNewComponentKey As String
  Dim sParentExpressionKey As String
  Dim objParentExpression As clsExprExpression
  Dim objNewComponent As clsExprComponent
  Dim objCurrentComponent As clsExprComponent
  Dim objPreviousComponent As clsExprComponent
  Dim bMakeFirstChildNode As Boolean
  Dim iHiddenElements As Integer
  Dim bPasteBelow As Boolean
  
  ' If the root node is selected then we want to add the component to the
  ' root expression.
  If sstrvComponents.SelectedItem.Key = ROOTKEY Then
    Set objParentExpression = mobjExpression
    sParentExpressionKey = ROOTKEY
  Else
    ' Get the selected component.
    Set objCurrentComponent = SelectedComponent(sstrvComponents.SelectedItem)

    ' Determine the parent expression of the selected component.
    If objCurrentComponent.ComponentType = giCOMPONENT_EXPRESSION Then
      Set objParentExpression = objCurrentComponent.Component
      sParentExpressionKey = sstrvComponents.SelectedItem.Key
    Else
      Set objParentExpression = SelectedExpression(sstrvComponents.SelectedItem)
      sParentExpressionKey = sstrvComponents.SelectedItem.Parent.Key
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
    
    'JDM - 08/11/01 - Fault 3124 - Strangely missing out the display of nodes.
    DoEvents
    
    sstrvComponents.SelectedItem.Expanded = True
    sstrvComponents.Refresh

    Me.Changed = True
    
    ' Check if there are hidden elements in the expression
    'SetAccessOptions HasHiddenComponents(mobjExpression.ExpressionID), Me.Expression.Access, mobjExpression.Owner
    SetAccessOptions HiddenElements, Me.Expression.Access, mobjExpression.Owner

    ' Ensure the command buttons are configured for the selected component.
    RefreshButtons
  End If

  ' Disassociate object variables.
  Set objParentExpression = Nothing
  Set objNewComponent = Nothing

End Function

Private Function HiddenElements() As Boolean
  ' Checks the expression for hidden components.
  HiddenElements = mobjExpression.HiddenElements
  
End Function

Public Function InsertComponent(pbNewComponent As Boolean, Optional pobjComponent As clsExprComponent, Optional lbInsertBelow As Boolean) As Boolean

  Dim objParentExpression As clsExprExpression
  Dim objCurrentComponent As clsExprComponent
  Dim objNewComponent As clsExprComponent
  Dim sNextNodeKey As String
  Dim sNewComponentKey As String
  Dim bExpandedNode As Boolean

  ' Get the selected component,, and it's parent expression.
  Set objCurrentComponent = SelectedComponent(sstrvComponents.SelectedItem)
  Set objParentExpression = SelectedExpression(sstrvComponents.SelectedItem)

  sNextNodeKey = sstrvComponents.SelectedItem.Key

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

    Me.Changed = True
    
    ' Check if there are hidden elements in the expression
    'SetAccessOptions HasHiddenComponents(mobjExpression.ExpressionID), Me.Expression.Access, mobjExpression.Owner
    SetAccessOptions HiddenElements, Me.Expression.Access, mobjExpression.Owner

    ' Ensure the command buttons are configured for the selected component.
    RefreshButtons
  End If

  ' Disassociate object variables.
  Set objCurrentComponent = Nothing
  Set objParentExpression = Nothing
  Set objNewComponent = Nothing

End Function

Public Sub MoveComponentDown()

Dim mobjComponent As clsExprComponent
Set mobjComponent = New clsExprComponent
Dim miOldNode As Integer

' Place components on the undo collection
CreateUndoView (giUNDO_MOVEDOWN)

'Ensure we are not the bottom node for this child
If Not sstrvComponents.SelectedItem.LastSibling.Index = sstrvComponents.SelectedItem.Index Then

  'Copy the current component, and node ID
  Set mobjComponent = SelectedComponent(sstrvComponents.SelectedItem)
  miOldNode = sstrvComponents.SelectedItem.Key

  'Move the current treeview node selection to the next one down
  sstrvComponents.SelectedItem.Next.Selected = True

  'Paste the object into the expression
  InsertComponent False, mobjComponent, True

  ' Remove the old component and node
  If SelectedExpression(sstrvComponents.SelectedItem).DeleteComponent(mobjComponent) = True Then
    RemoveComponentNode (miOldNode)
  End If

End If

'Refresh tht display
RefreshButtons

'Clear up memory
Set mobjComponent = Nothing

End Sub

Public Sub MoveComponentUp()

Dim mobjComponent As clsExprComponent
Set mobjComponent = New clsExprComponent
Dim miOldNode As Integer

' Place components on the undo collection
CreateUndoView (giUNDO_MOVEUP)

'Ensure we are not the bottom node for this child
If Not sstrvComponents.SelectedItem.FirstSibling.Index = sstrvComponents.SelectedItem.Index Then

  'Copy the current component, and node ID
  Set mobjComponent = SelectedComponent(sstrvComponents.SelectedItem)
  miOldNode = sstrvComponents.SelectedItem.Key

  'Move the current treeview node selection to the next one down
  sstrvComponents.SelectedItem.Previous.Selected = True

  'Paste the object into the expression
  InsertComponent False, mobjComponent, False

  ' Remove the old component and node
  If SelectedExpression(sstrvComponents.SelectedItem).DeleteComponent(mobjComponent) = True Then
    RemoveComponentNode (miOldNode)
  End If

End If

'Refresh tht display
RefreshButtons

'Clear up memory
Set mobjComponent = Nothing

End Sub

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


Public Sub SetInitialExpandedNodes()

  Dim iCount As Integer
  Dim iLevelToExpandTo  As Integer

  iLevelToExpandTo = 2

  Select Case GetSystemSetting("ExpressionBuilder", "NodeSize", EXPRESSIONBUILDER_NODESMINIMIZE)

    Case EXPRESSIONBUILDER_NODESMINIMIZE
      ' Shrink all nodes
      For iCount = 1 To sstrvComponents.Nodes.Count
        sstrvComponents.Nodes(iCount).Expanded = False
      Next iCount

    Case EXPRESSIONBUILDER_NODESEXPAND
      ' Expand all nodes
      For iCount = 1 To sstrvComponents.Nodes.Count
        sstrvComponents.Nodes(iCount).Expanded = True
        sstrvComponents.Nodes(iCount).EnsureVisible
      Next iCount

    Case EXPRESSIONBUILDER_NODESLASTSAVE
      ' Do nothing as this is the default

    Case EXPRESSIONBUILDER_NODESTOPLEVEL
      'Expand all specified levels
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

