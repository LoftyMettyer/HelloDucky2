VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{1C203F10-95AD-11D0-A84B-00A0247B735B}#1.0#0"; "SSTree.ocx"
Begin VB.Form frmViewMgr 
   Caption         =   "View Manager"
   ClientHeight    =   3825
   ClientLeft      =   390
   ClientTop       =   1830
   ClientWidth     =   7470
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5037
   Icon            =   "frmViewMgr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3825
   ScaleWidth      =   7470
   WindowState     =   2  'Maximized
   Begin VB.ListBox lstColumns 
      Height          =   1410
      Left            =   4000
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   1900
      Visible         =   0   'False
      Width           =   3000
   End
   Begin ComctlLib.ListView lstViews 
      Height          =   1500
      Left            =   4000
      TabIndex        =   1
      Top             =   315
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   2646
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      _Version        =   327682
      Icons           =   "imglstLargeIcons"
      SmallIcons      =   "imglstIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Description"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.Frame fraSplit 
      BorderStyle     =   0  'None
      Height          =   3000
      Left            =   3500
      MousePointer    =   9  'Size W E
      TabIndex        =   3
      Top             =   200
      Width           =   200
   End
   Begin SSActiveTreeView.SSTree trvTables 
      Height          =   1500
      Left            =   0
      TabIndex        =   0
      Top             =   315
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   2646
      _Version        =   65538
      LabelEdit       =   1
      LineStyle       =   1
      Indentation     =   315
      AutoSearch      =   0   'False
      HideSelection   =   0   'False
      PictureBackgroundUseMask=   0   'False
      HasFont         =   -1  'True
      HasMouseIcon    =   0   'False
      HasPictureBackground=   0   'False
      ImageList       =   "imglstIcons"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Sorted          =   1
   End
   Begin ComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   3540
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   10081
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
   Begin ActiveBarLibraryCtl.ActiveBar abViewMgr 
      Left            =   2085
      Top             =   2340
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
      Bands           =   "frmViewMgr.frx":000C
   End
   Begin ComctlLib.ImageList imglstLargeIcons 
      Left            =   915
      Top             =   2310
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmViewMgr.frx":ECDC
            Key             =   "IMG_VIEW"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblRightPaneCaption 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Right Pane Caption"
      Height          =   285
      Left            =   4000
      TabIndex        =   5
      Top             =   0
      Width           =   3000
   End
   Begin VB.Label lblLeftPaneCaption 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Left Pane Caption"
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3000
   End
   Begin ComctlLib.ImageList imglstIcons 
      Left            =   200
      Top             =   2300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmViewMgr.frx":F52E
            Key             =   "IMG_TABLE"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmViewMgr.frx":FA80
            Key             =   "IMG_VIEW"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmViewMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declare events
Event Activate()
Event Deactivate()
Event UnLoad()

'Local variables
Private gctlActiveView As Control
Private gfSplitMoving As Boolean
Private gSngSplitStartX As Single
Private gfMenuActionKey As Boolean
Private gsTreeViewNodeKey As String
Private gfRefreshing As Boolean

Private mblnReadOnly As Boolean

Private Const MIN_FORM_HEIGHT = 5000
Private Const MIN_FORM_WIDTH = 6000

Private Sub SplitMove()
  
  ' Limit the minimum size of the tree and list views.
  With fraSplit
    If .Left < 810 Then
      .Left = 810
    ElseIf .Left + .Width > Me.ScaleWidth - 2000 Then
      .Left = Me.ScaleWidth - (2000 + .Width)
    End If
  End With
  
  ' Resize the tree view.
  With trvTables
    .Width = fraSplit.Left - .Left
    lblLeftPaneCaption.Width = .Width
  End With
  
  ' Resize the listview and grid controls.
  With lstViews
    .Left = fraSplit.Left + fraSplit.Width
    .Width = Me.ScaleWidth - .Left
    lstColumns.Left = .Left
    lstColumns.Width = .Width
    lblRightPaneCaption.Left = .Left
    lblRightPaneCaption.Width = .Width
  End With
  
  ' Flag that the split move has ended.
  gfSplitMoving = False

End Sub



Private Sub RefreshStatusBar()
  Dim iItems As Integer
  Dim iSelections As Integer
  Dim sMessage As String
  
  If trvTables.SelectedItem Is Nothing Then
    sMessage = ""
  Else
    If trvTables.SelectedItem.DataKey = "TABLE" Then
      iItems = lstViews.ListItems.Count
      iSelections = lstViews_SelectedCount
  
      sMessage = Trim(Str(iItems)) & " view" & IIf(iItems <> 1, "s", "") & _
        ", " & Trim(Str(iSelections)) & " selected."
    ElseIf trvTables.SelectedItem.DataKey = "VIEW" Then
      iItems = lstColumns.ListCount - 1
      iSelections = lstColumns_SelectedCount
      
      sMessage = Trim(Str(iItems)) & " column" & IIf(iItems <> 1, "s", "") & _
        ", " & Trim(Str(iSelections)) & " selected."
    Else
      sMessage = ""
    End If
  End If
  
  sbStatus.Panels(1).Text = sMessage
  
End Sub

Private Sub abViewMgr_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)

  EditMenu Tool.Name
  
End Sub

Private Sub abViewMgr_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)

  ' Do not let the user modify the layout.
  Cancel = True

End Sub

Private Sub Form_Activate()
  RaiseEvent Activate

End Sub

Private Sub Form_Deactivate()
  RaiseEvent Deactivate
  
End Sub

Private Sub Form_GotFocus()
  ' Set focus on the treeview.
  trvTables.SetFocus

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  'TM20020102 Fault 2879
  Dim bHandled As Boolean
  
  
  Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select

  bHandled = frmSysMgr.tbMain.OnKeyDown(KeyCode, Shift)
  If bHandled Then
    KeyCode = 0
    Shift = 0
  End If

  'MH20010702
  'Still allow F1 key for help though
  If KeyCode <> vbKeyF1 And Not bHandled Then

    ' JDM - 19/02/01 - Fault 1869 - Error when pressing CTRL-X on treeview control
    ' For some reason the Sheridan treeview control wants to fire off it own cutn'paste functionality
    ' must trap it here not in it's own keydown event
    If ActiveControl.Name = "trvTables" Then
      KeyCode = 0
      Shift = 0
    End If
  End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

  'TM20020102 Fault 2879
  Dim bHandled As Boolean
  
  bHandled = frmSysMgr.tbMain.OnKeyUp(KeyCode, Shift)
  If bHandled Then
    KeyCode = 0
    Shift = 0
  End If

End Sub

Private Sub Form_Load()
  Dim sAppName  As String
  Dim sSection  As String

  Hook Me.hWnd, MIN_FORM_WIDTH, MIN_FORM_HEIGHT

  mblnReadOnly = (Application.AccessMode = accSystemReadOnly)

  ' Initialise form properties from registry settings.
  With Me
    sAppName = App.ProductName
    sSection = .Name
  End With
  
  If gbMaximizeScreens Then
    Me.WindowState = vbMaximized
  End If
  
  gsTreeViewNodeKey = ""
  gfRefreshing = False
  
  lstViews.View = GetPCSetting(sSection, "View", lvwIcon)

  ' Position controls.
  With lblLeftPaneCaption
    .Left = 0
    .Caption = " Tables"
  End With
  lstViews.Top = trvTables.Top
  lstColumns.Top = trvTables.Top
  fraSplit.Width = UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX
  
  ' Populate the treeview.
  trvTables_Initialise
  ' Populate the listview or grid.
  UpdateRightPane
    
  ChangeView lstViews.View
  
End Sub

Private Sub Form_Resize()

  Dim lngHeight As Long

  ' If the form is minimized the do nothing.
  If Me.WindowState <> vbMinimized Then
  
    ' Position the label controls on the form.
    With lblLeftPaneCaption
      .Top = 0
      lblRightPaneCaption.Top = .Top
      trvTables.Top = .Top + .Height
    End With
    lstViews.Top = trvTables.Top
    lstColumns.Top = trvTables.Top
    
    ' Size the tree and list view controls.
    lngHeight = Me.ScaleHeight - (trvTables.Top + sbStatus.Height)
    If lngHeight < 0 Then lngHeight = 0
    trvTables.Height = lngHeight
    lstViews.Height = lngHeight
    lstColumns.Height = lngHeight
  
    ' Position and size the split frame.
    With fraSplit
      .Width = UI.GetSystemMetrics(SM_CXFRAME) * Screen.TwipsPerPixelX
      .Top = lblLeftPaneCaption.Top
      .Height = lblLeftPaneCaption.Height + trvTables.Height
      If .Left + .Width > Me.ScaleWidth - 810 Then
        .Left = Me.ScaleWidth - (810 + .Width)
      End If
    End With
    
    ' Call the routine to size the tree, list view and gridcontrols.
    SplitMove
    
    ' Refresh the form display.
    Me.Refresh
    
  End If
  
  frmSysMgr.RefreshMenu

  ' Get rid of the icon off the form
  If Me.WindowState = vbMaximized Then
    SetBlankIcon Me
  Else
    RemoveIcon Me
    Me.BorderStyle = vbSizable
  End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim sAppName As String
  Dim sSection As String
   
  ' Save form size and position to the registry.
  With Me
    sAppName = App.ProductName
    sSection = .Name
  End With
    
  ' Ensure the menu is updated.
  frmSysMgr.RefreshMenu True
  
  Unhook Me.hWnd
  
End Sub



Private Sub trvTables_Initialise()
  Dim nodX As SSNode
  Dim sSQL As String
  Dim sNodeLabel As String
  
  ' Clear the treeview.
  trvTables.Nodes.Clear
  
  ' Get a list of the tables and add them as nodes to the tree
  With recTabEdit
    .Index = "idxTableID"
    
    If Not (.BOF And .EOF) Then
      .MoveFirst
      
      While Not .EOF
        ' Check that it has not been deleted
        If Not .Fields("Deleted") Then
          ' JPD 13/1/00 Only allow views to be created on top-level tables.
          If !TableType = iTabParent Then
            Set nodX = trvTables.Nodes.Add(, , "T" & .Fields("TableID"), _
              .Fields("TableName"), "IMG_TABLE", , "TABLE")
            nodX.Sorted = True
            Set nodX = Nothing
          End If
        End If
        
        .MoveNext
      Wend
    End If
  End With
  
  ' Now get a list of the views that are associated with these tables
  With recViewEdit
  
    .Index = "idxViewID"
    On Error Resume Next  ' Ignore an error on movefirst if no records in table
    .MoveFirst
    On Error GoTo 0
    
    While Not .EOF
      ' Check that it is not marked for deletion
      If Not .Fields("Deleted") Then
        sNodeLabel = Trim(.Fields("ViewName"))
        Set nodX = trvTables.Nodes.Add("T" & .Fields("ViewTableID"), ssatChild, _
          "V" & .Fields("ViewID"), sNodeLabel, "IMG_VIEW", , "VIEW")
        Set nodX = Nothing
      End If
      
      .MoveNext
    Wend
  
  End With
  
End Sub


Private Sub fraSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Record the split move start position.
  gSngSplitStartX = X
  
  ' Flag that the split is being moved.
  gfSplitMoving = True
  
End Sub


Private Sub fraSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' If we are moving the split then move it.
  If gfSplitMoving Then
    fraSplit.Left = fraSplit.Left + (X - gSngSplitStartX)
  End If
  
End Sub


Private Sub fraSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' If the split is being moved then call the routine that resizes the
  ' tree and list views accordingly.
  If gfSplitMoving Then
    SplitMove
  End If
  
End Sub


Private Sub lstColumns_GotFocus()
  Set gctlActiveView = lstColumns
  frmSysMgr.RefreshMenu

End Sub


Private Sub lstColumns_ItemCheck(Item As Integer)
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim fNewValue As Boolean
  Dim fAllColumns As Boolean
  Dim iCount As Integer
  Dim lngViewID As Long
  Dim lngColumnID As Long
  
  fOK = True
  
  If gfRefreshing Then
    Exit Sub
  End If
  gfRefreshing = True
  
  
  If mblnReadOnly Then
    With lstColumns
      .Selected(.ListIndex) = Not .Selected(.ListIndex)
    End With
    gfRefreshing = False
    Exit Sub
  End If
  
  
  ' Get the view ID from the selected treeview node.
  lngViewID = Right(trvTables.SelectedItem.key, Len(trvTables.SelectedItem.key) - 1)
  
  fNewValue = lstColumns.Selected(Item)

  UI.LockWindow lstColumns.hWnd

  ' Check if the user has selected the all columns.
  If Item = 0 Then
    For iCount = 1 To lstColumns.ListCount - 1
      ' Update all rows in the listbox.
      lstColumns.Selected(iCount) = fNewValue
      lngColumnID = lstColumns.ItemData(iCount)
      fOK = ChangeViewColumn_Transaction(lngViewID, lngColumnID, fNewValue)
      
      If Not fOK Then
        Exit For
      End If
    Next iCount
  Else
    lngColumnID = lstColumns.ItemData(Item)
        
    fOK = ChangeViewColumn_Transaction(lngViewID, lngColumnID, fNewValue)
        
    If fOK Then
      ' Check for all columns now being selected or deselected
      fAllColumns = True
      For iCount = 1 To lstColumns.ListCount - 1
        If Not lstColumns.Selected(iCount) Then
          fAllColumns = False
          Exit For
        End If
      Next iCount
      ' Update the 'All' row in the listbox.
      lstColumns.Selected(0) = fAllColumns
    End If
  End If
    
  If fOK Then
    frmSysMgr.RefreshMenu
    RefreshStatusBar
  End If
  
TidyUpAndExit:
  ' Select the original item.
  lstColumns.ListIndex = Item
  gfRefreshing = False
  UI.UnlockWindow
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub


Public Function ChangeViewColumn_Transaction(plngViewID As Long, plngColumnID As Long, pfValue As Boolean) As Boolean
  ' Transaction wrapper for the 'DeleteColumn' function.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  fOK = True
  
  ' Update the database
  With recViewColEdit
    
    .Index = "idxViewColID"
    .Seek "=", plngViewID, plngColumnID
      
    If .NoMatch Then
      ' Add the column to the view.
      .AddNew
      .Fields("ViewID") = plngViewID
      .Fields("ColumnID") = plngColumnID
      .Fields("InView") = pfValue
      .Fields("New") = True
      .Fields("Deleted") = False
      .Fields("Changed") = False
    Else
      ' Update the column record in the view.
      .Edit
      .Fields("InView") = pfValue
      If Not .Fields("New") Then .Fields("Changed") = True
    End If
      
    .Update
  End With

  With recViewEdit
    .Index = "idxViewID"
    .Seek "=", plngViewID
    If Not .NoMatch Then
      If Not .Fields("New") Then
        .Edit
        .Fields("Changed") = True
        .Update
      End If
    End If
  End With
  
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  ChangeViewColumn_Transaction = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Sub lstViews_DblClick()
  Dim ThisNode As SSNode
  
  ' If we have some listview items ...
  If lstViews.ListItems.Count > 0 Then
  
    ' If the select edlistview item has children ...
    If trvTables.SelectedItem.Children > 0 Then
      
      ' Set the selected item to be the selected item in the treeview.
      Set ThisNode = trvTables.Nodes(lstViews.SelectedItem.key)
      ThisNode.EnsureVisible
      trvTables.SelectedItem = ThisNode
      
      gsTreeViewNodeKey = ThisNode.key
      
      ' Disassociate object variables.
      Set ThisNode = Nothing
      
      ' Populate the listview with the children of the selected item.
      UpdateRightPane
    Else
      ' If the selected item does not have children then display its
      ' property page.
      EditMenu "ID_Properties"
    End If
  End If

End Sub


Private Sub lstViews_GotFocus()
  Set gctlActiveView = lstViews
  frmSysMgr.RefreshMenu

End Sub


Private Sub lstViews_ItemClick(ByVal Item As ComctlLib.ListItem)
'  Set gctlActiveView = lstViews
'  frmSysMgr.RefreshMenu

End Sub


Private Sub lstViews_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyInsert
      EditMenu "ID_New"
      lstViews.SetFocus
    Case vbKeyReturn
      If (Not lstViews.SelectedItem Is Nothing) And (lstViews = 1) Then
        lstViews_DblClick
      End If
  End Select

End Sub


Private Sub lstViews_KeyPress(KeyAscii As Integer)
  ' If we have just pressed a menu hot-key then do not process
  ' the key press as a jump the next listview item beginning
  ' with that letter.
  If gfMenuActionKey Then
    KeyAscii = 0
    gfMenuActionKey = False
  End If

End Sub


Private Sub lstViews_KeyUp(KeyCode As Integer, Shift As Integer)
  ' Refresh the status bar.
  RefreshStatusBar

End Sub

Private Sub lstViews_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim lXMouse As Long
  Dim lYMouse As Long
  
  ' Display a pop-up menu.
  If Button = vbRightButton Then
  
    frmSysMgr.RefreshMenu
    
    With frmSysMgr.tbMain
      'If .Tools("ID_New").Enabled Then
      If .Tools("ID_New").Enabled Or .Tools("ID_Properties").Enabled Then
        UI.GetMousePos lXMouse, lYMouse
'        .PopupMenu "ID_mnuEdit", ssPopupMenuLeftAlign, lXMouse, lYMouse
        .Bands("ID_mnuEdit").TrackPopup -1, -1
      End If
    End With
    
  End If
      
  ' Refresh the status bar.
  RefreshStatusBar
  
  ' Refresh the menu.
  frmSysMgr.RefreshMenu

End Sub


Private Sub trvTables_Collapse(Node As SSActiveTreeView.SSNode)
  ' Ensure the specified node is selected.
  Node.Selected = True
  
  ' Populate the listview with the children of the specified node.
  If gsTreeViewNodeKey <> Node.key Then
    UpdateRightPane
    gsTreeViewNodeKey = Node.key
  
    ' Refresh the menu.
    frmSysMgr.RefreshMenu
  End If
  
End Sub

Private Sub trvTables_DblClick()
  
  If Not trvTables.SelectedItem Is Nothing Then
    If trvTables.SelectedItem.DataKey = "VIEW" Then
      ShowViewProperties
    End If
  End If

End Sub


Private Sub trvTables_Expand(Node As SSActiveTreeView.SSNode)
  ' Ensure the specified node is selected.
  Node.Selected = True
  
  ' Populate the listview with the children of the specified node.
  If gsTreeViewNodeKey <> Node.key Then
    UpdateRightPane
    gsTreeViewNodeKey = Node.key
  
    ' Refresh the menu.
    frmSysMgr.RefreshMenu
  End If
  
End Sub

Private Sub trvTables_GotFocus()
  Set gctlActiveView = trvTables
  frmSysMgr.RefreshMenu

End Sub

Private Sub trvTables_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyInsert Then
    EditMenu "ID_New"
  End If
  
  ' Refresh the status bar.
  RefreshStatusBar

End Sub


Private Sub trvTables_KeyPress(KeyAscii As Integer)
  ' If we have just pressed menu hot-key then do not process
  ' the key press as a jump the next listview item beginning
  ' with that letter.
  If gfMenuActionKey Or (KeyAscii = vbKeyReturn) Then
    KeyAscii = 0
    gfMenuActionKey = False
  End If

End Sub


Private Sub trvTables_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim lXMouse As Long
  Dim lYMouse As Long
  Dim nodX As SSNode

  ' Pop up a menu if the right mouse button is pressed.
  If Button = vbRightButton Then
  
    ' Ensure the menu is up to date.
    frmSysMgr.RefreshMenu
    
    ' Call the activebar to display the popup menu.
    ' Check that we are over a node
    Set nodX = trvTables.HitTest(X, Y)
    If Not nodX Is Nothing Then
      Set nodX = Nothing
      UI.GetMousePos lXMouse, lYMouse
'        frmSysMgr.tbMain.PopupMenu "ID_mnuEdit", ssPopupMenuLeftAlign, lXMouse, lYMouse
      frmSysMgr.tbMain.Bands("ID_mnuEdit").TrackPopup -1, -1
    End If
  End If

End Sub

Private Sub trvTables_NodeClick(Node As SSActiveTreeView.SSNode)
  ' Update the display for the selected node.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  ' If we are changing node then clear any selections from the listview.
  If Node.key <> gsTreeViewNodeKey Then
    UpdateRightPane
    gsTreeViewNodeKey = Node.key
  
    ' Update the menu.
    frmSysMgr.RefreshMenu
  End If

TidyUpAndExit:
  If Not fOK Then
    MsgBox "Error refreshing display.", vbExclamation + vbOKOnly, App.ProductName
  End If
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub


Public Sub EditMenu(psMenuItem As String)
  ' Perform the required operation for the menu selection.
  
  Select Case psMenuItem
    
    Case "ID_SaveChanges"
      View_SaveChanges
    
    Case "ID_New"
      ' A new view is required
      AddNewView

    Case "ID_Delete"
      ' Delete the current selected view(s)
      DeleteViews
    
    Case "ID_CopyDef"
      ' Copy the selected view(s)
      CopyView
      
    Case "ID_Properties"
      ' The property page for the view is required
      ShowViewProperties

    Case "ID_SelectAll"
      ' Select all columns
      gfMenuActionKey = True
      lstViews_SelectAll
      lstViews.SetFocus
      
    Case "ID_LargeIcons"
      ' Change the view to display large icons.
      ChangeView 0 'lvwIcon

    Case "ID_SmallIcons"
      ' Change the view to display small icons.
      ChangeView 1 'lvwSmallIcon
  
    Case "ID_List"
      ' Change the view to display a list.
      ChangeView 2 'lvwList

    Case "ID_Details"
      ' Change the view to display details.
      ChangeView 3 'lvwReport
      
    Case "ID_CustomiseColumns"
      ' Customise column view...
      Set frmShowColumns = New SystemMgr.frmShowColumns
      frmShowColumns.PropertySet = gpropShowColumns_ViewMgr
      frmShowColumns.Show vbModal
      SetColumnSizes
      Exit Sub
      
  End Select
      
End Sub


Private Sub View_SaveChanges()

  Dim frmPrompt As frmSaveChangesPrompt
  
  ' Save changes without exiting.
  Set frmPrompt = New frmSaveChangesPrompt
  frmPrompt.Buttons = vbOKCancel
  frmPrompt.Show vbModal
  If frmPrompt.Choice = vbOK Then
    Application.Changed = Not (SaveChanges(frmPrompt.RefreshDatabase))
    Me.SetFocus
    frmSysMgr.RefreshMenu
  End If
  Set frmPrompt = Nothing

End Sub
      
Private Sub lstViews_SelectAll()
  Dim iLoop As Integer
  
  ' Loop through the list view items marking each one as selected.
  For iLoop = 1 To lstViews.ListItems.Count
    lstViews.ListItems(iLoop).Selected = True
  Next iLoop

End Sub


Private Sub AddNewView()
  ' Add a new view for the selected table.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim lngTableID As Long
  Dim sNodeLabel As String
  Dim frmViewProperties As frmViewProp
  Dim nodX As SSNode
  
  ' Get the ID of the selected table.
  lngTableID = 0
  If gctlActiveView Is trvTables Then
    If Not trvTables.SelectedItem Is Nothing Then
      If trvTables.SelectedItem.DataKey = "TABLE" Then
        lngTableID = Right(trvTables.SelectedItem.key, Len(trvTables.SelectedItem.key) - 1)
      Else
        lngTableID = Right(trvTables.SelectedItem.Parent.key, Len(trvTables.SelectedItem.Parent.key) - 1)
      End If
    End If
  Else
    If gctlActiveView Is lstViews Then
      lngTableID = Right(trvTables.SelectedItem.key, Len(trvTables.SelectedItem.key) - 1)
    End If
  End If
  
  fOK = (lngTableID > 0)
  
  If fOK Then
    ' Display the view property form.
    Set frmViewProperties = New frmViewProp
    
    With frmViewProperties
      .TableID = lngTableID
      .ViewID = 0
      .Show vbModal
      fOK = Not .Cancelled
    End With
    
    ' Decide if OK or Cancel was chosen.
    If fOK Then
      
      ' Save the properites.
      With recViewEdit
        .Index = "idxViewID"
        .Seek "=", frmViewProperties.ViewID
        
        .Bookmark = .LastModified
        
        ' Now add it to the tree view
        sNodeLabel = Trim(.Fields("ViewName"))
        Set nodX = trvTables.Nodes.Add("T" & .Fields("ViewTableID"), ssatChild, _
          "V" & .Fields("ViewID"), sNodeLabel, "IMG_VIEW", , "VIEW")
      End With
      
      ' Now add all the columns to the ASRSysViewColumns
      fOK = AddViewColumns_Transaction(recViewEdit.Fields("ViewID"), recViewEdit!ViewTableID)
      
      ' Finally select the new node and then release the reference
      nodX.EnsureVisible
      Set trvTables.SelectedItem = nodX
      trvTables_NodeClick trvTables.SelectedItem
    
      Set nodX = Nothing
  
    End If
  End If
  
TidyUpAndExit:
  Set frmViewProperties = Nothing
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub

'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@
'@@ Procedure : ShowViewProperties
'@@
'@@ Desc      : This procedure will show the selected views property page
'@@             form
'@@
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'@@ Changes   :
'@@ 13/08/1998  RJB   Created

Private Sub ShowViewProperties()
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim lngViewID As Long
  Dim sNodeLabel As String
  Dim frmViewProperties As frmViewProp
  
  lngViewID = 0
  If gctlActiveView Is trvTables Then
    If Not trvTables.SelectedItem Is Nothing Then
      lngViewID = Right(trvTables.SelectedItem.key, Len(trvTables.SelectedItem.key) - 1)
    End If
  Else
    If gctlActiveView Is lstViews Then
      If Not lstViews.SelectedItem Is Nothing Then
        lngViewID = Right(lstViews.SelectedItem.key, Len(lstViews.SelectedItem.key) - 1)
      End If
    End If
  End If
  
  fOK = (lngViewID > 0)
  
  If fOK Then
    ' Find the record in the table.
    With recViewEdit
      .Index = "idxViewID"
      .Seek "=", lngViewID
      
      fOK = Not .NoMatch
    End With
  End If
    
  If fOK Then
    ' Display the view property form.
    Set frmViewProperties = New frmViewProp
    frmViewProperties.ViewID = lngViewID
      
    frmViewProperties.Show vbModal
      
    'NHRD25072003 Fault 6309
    lstViews_Refresh
    
    ' Decide if OK or Cancel was chosen.
    If Not frmViewProperties.Cancelled Then
      With recViewEdit
        .Index = "idxViewID"
        .Seek "=", lngViewID
        ' Update the tree view description
        sNodeLabel = Trim(.Fields("ViewName"))
      End With
      'NHRD10092003 Fault 4457 Commented out line below as this was causing the problem for this fault
      'trvTables.SelectedItem.Text = sNodeLabel

      'JPD 20050113 Fault 9314
      trvTables.Nodes("V" & CStr(lngViewID)).Text = sNodeLabel
    End If
  Else
    ' Display an error message as for some reason we can not find the view.
    MsgBox "Unable to locate the view in the local database.", vbCritical + vbOKOnly, App.Title
  End If
  
TidyUpAndExit:
  Set frmViewProperties = Nothing
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub

Private Sub DeleteViews()
  ' Delete the selected view(s).
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fDeleteAll As Boolean
  Dim fConfirmed As Boolean
  Dim iLoop As Integer
  Dim iNextIndex As Integer
  Dim lngViewID As Long
  Dim lngTableID As Long
  Dim sViewName As String
  Dim aLngViewID() As Long
  Dim sSQL As String
  Dim rsModules As DAO.Recordset
  Dim rsTemp As DAO.Recordset
  Dim sModuleName As String
  Dim frmUse As frmUsage
  Dim fUsed As Boolean
  Dim sColumnName As String
  Dim sTableName As String
  Dim rsUtils1 As ADODB.Recordset
  
  ReDim aLngViewID(0)
  fDeleteAll = False
  fConfirmed = False
  fOK = True
  
  ' If we have more than one selection then question the multi-deletion.
  If gctlActiveView Is lstViews Then
    
    lngTableID = Right(trvTables.SelectedItem.key, Len(trvTables.SelectedItem.key) - 1)
    
    If (lstViews_SelectedCount > 1) Then
      If MsgBox("Are you sure you want to delete these " & _
        Trim(Str(lstViews_SelectedCount)) & _
        " views ?", vbQuestion + vbYesNo, App.ProductName) = vbYes Then
        
        fDeleteAll = True
        fConfirmed = True
      Else
        Exit Sub
      End If
      
      ' Read the ids of the views to be deleted from the listview into an array.
      For iLoop = 1 To lstViews.ListItems.Count
        If lstViews.ListItems(iLoop).Selected = True Then
          iNextIndex = UBound(aLngViewID) + 1
          ReDim Preserve aLngViewID(iNextIndex)
          aLngViewID(iNextIndex) = Right(lstViews.ListItems(iLoop).key, Len(lstViews.ListItems(iLoop).key) - 1)
        End If
      Next iLoop
    Else
      ' Read the ids of the view to be deleted from the treeview into an array.
      ReDim aLngViewID(1)
      aLngViewID(1) = Right(lstViews.SelectedItem.key, Len(lstViews.SelectedItem.key) - 1)
    End If
    
  ElseIf gctlActiveView Is trvTables Then
    ReDim aLngViewID(1)
    aLngViewID(1) = Right(trvTables.SelectedItem.key, Len(trvTables.SelectedItem.key) - 1)
    lngTableID = Right(trvTables.SelectedItem.Parent.key, Len(trvTables.SelectedItem.Parent.key) - 1)
  End If

  ' Delete all of the views in the array..
  For iLoop = 1 To UBound(aLngViewID)
  
    lngViewID = aLngViewID(iLoop)
    
    With recViewEdit
      .Index = "idxViewID"
      .Seek "=", lngViewID
      
      If Not .NoMatch Then
        sViewName = Trim(.Fields("viewName"))
      End If
    End With
    
    If Not fDeleteAll Then
      ' Prompt the user to confirm the deletion.
      If MsgBox("Are you sure you want to delete the view '" & _
        sViewName & "' ?", vbYesNo + vbDefaultButton2 + _
        vbQuestion, Application.Name) = vbYes Then
                          
        fConfirmed = True
      Else
        fConfirmed = False
      End If
    End If
     
    If fConfirmed Then
      fUsed = False
      
      Set frmUse = New frmUsage
      frmUse.ResetList
      
      ' Check if the view is used anywhere.
      sSQL = "SELECT moduleKey" & _
        " FROM tmpModuleSetup" & _
        " WHERE parameterType = '" & gsPARAMETERTYPE_VIEWID & "'" & _
        " AND parameterValue = '" & Trim(Str(lngViewID)) & "'"
      Set rsModules = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
          
      If Not (rsModules.BOF And rsModules.EOF) Then
        Select Case rsModules!moduleKey
          Case gsMODULEKEY_TRAININGBOOKING
            sModuleName = "Training Booking"
          Case gsMODULEKEY_PERSONNEL
            sModuleName = "Personnel"
          Case gsMODULEKEY_ABSENCE
            sModuleName = "Absence"
          Case gsMODULEKEY_SSINTRANET
            sModuleName = "Self-service Intranet"
          Case Else
            sModuleName = "<unknown>"
        End Select
          
        frmUse.AddToList "Module : " & sModuleName
        fUsed = True
      End If
      ' Close the recordset.
      rsModules.Close
      Set rsModules = Nothing
      
      ' Now Check if the view is used in a Organisation Report
      sSQL = "SELECT DISTINCT r.Name, r.ID, r.Username" & _
             " FROM ASRSysOrganisationReport r" & _
             " WHERE r.BaseViewId = " & Trim(Str(lngViewID))
      Set rsUtils1 = New ADODB.Recordset
      rsUtils1.Open sSQL, gADOCon, adOpenStatic, adLockReadOnly
      With rsUtils1
        If Not (.EOF And .BOF) Then
          fUsed = True
          Do Until .EOF
            frmUse.AddToList ("Organisation Report : " & !Name)
            .MoveNext
          Loop
        End If
        .Close
      End With
      Set rsUtils1 = Nothing
      
      ' Find any columns that use this expression as a link - default view value.
      sSQL = "SELECT DISTINCT tmpColumns.columnName, tmpColumns.tableID" & _
        " FROM tmpColumns" & _
        " WHERE deleted = FALSE" & _
        " AND tmpColumns.columnType = 4 " & _
        " AND tmpColumns.linkViewID = " & Trim(Str(lngViewID))
      Set rsTemp = daoDb.OpenRecordset(sSQL, _
        dbOpenForwardOnly, dbReadOnly)
      If Not (rsTemp.BOF And rsTemp.EOF) Then
        fUsed = True
        Do Until rsTemp.EOF
          ' Get the column and table names.
          sColumnName = rsTemp!ColumnName
            
          recTabEdit.Index = "idxTableID"
          recTabEdit.Seek "=", rsTemp!TableID
            
          If Not recTabEdit.NoMatch Then
            sTableName = recTabEdit!TableName
          Else
            sTableName = "<unknown>"
          End If
          frmUse.AddToList ("Link Column : " & sColumnName & " <" & sTableName & ">")
          rsTemp.MoveNext
        Loop
      End If
      rsTemp.Close
      
      sSQL = "SELECT COUNT(*) AS result" & _
        " FROM tmpSSIViews" & _
        " WHERE viewID = " & Trim(Str(lngViewID))
      Set rsTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

      If (rsTemp!result > 0) Then
        frmUse.AddToList "Module : Self-service Intranet links"
        fUsed = True
      End If
      rsTemp.Close
      Set rsTemp = Nothing
      
      If fUsed Then
        Screen.MousePointer = vbDefault
        frmUse.ShowMessage sViewName & " View", "The view cannot be deleted as it is used by the following:", UsageCheckObject.View
      Else
        fOK = DeleteView_Transaction(lngViewID)
  
        ' Remove the node from the tree view
        If fOK Then
          trvTables.Nodes.Remove "V" & Trim(Str(lngViewID))
        End If
      End If
        
      UnLoad frmUse
      Set frmUse = Nothing
    End If
  Next iLoop

  If fConfirmed Then
    Set trvTables.SelectedItem = trvTables.Nodes("T" & Trim(Str(lngTableID)))
    trvTables_NodeClick trvTables.SelectedItem
  End If
  
  UpdateRightPane
  
  If Not gctlActiveView Is Nothing Then
    gctlActiveView.SetFocus
  End If

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub

Public Function DeleteView_Transaction(plngViewID As Long) As Boolean
  ' Delete the given view.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  ' Find the view in the database and mark it for deletion.
  With recViewEdit
    .Index = "idxViewID"
    .Seek "=", plngViewID
        
    fOK = Not .NoMatch
    If fOK Then
      .Edit
      .Fields("Deleted") = True
      .Update
    End If
  End With
  
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  DeleteView_Transaction = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function CopyView() As Boolean
  ' Add a new copy of the selected view.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim lngViewID As Long
  Dim sNodeLabel As String
  Dim frmViewProperties As frmViewProp
  Dim nodX As SSNode
  
  ' Get the ID of the selected view.
  lngViewID = 0
  If gctlActiveView Is trvTables Then
    If Not trvTables.SelectedItem Is Nothing Then
      If trvTables.SelectedItem.DataKey = "VIEW" Then
        lngViewID = Right(trvTables.SelectedItem.key, Len(trvTables.SelectedItem.key) - 1)
      End If
    End If
  Else
    If gctlActiveView Is lstViews Then
      'NPG20080211 Fault 12874
      lngViewID = Right(lstViews.SelectedItem.key, Len(lstViews.SelectedItem.key) - 1)
    End If
  End If
  
  fOK = (lngViewID > 0)
  
  If fOK Then
    ' Display the view property form.
    Set frmViewProperties = New frmViewProp
    
    With frmViewProperties
      .Copy = True
      .ViewID = lngViewID
      .Show vbModal
      fOK = Not .Cancelled
    End With
      
    ' Decide if OK or Cancel was chosen.
    If fOK Then
      ' Save the properites.
      With recViewEdit
        .Index = "idxViewID"
        .Seek "=", frmViewProperties.ViewID
        
        .Bookmark = .LastModified
        
        ' Now add it to the tree view
        sNodeLabel = Trim(.Fields("ViewName"))
        Set nodX = trvTables.Nodes.Add("T" & .Fields("ViewTableID"), ssatChild, _
          "V" & .Fields("ViewID"), sNodeLabel, "IMG_VIEW", , "VIEW")
      End With
      
      ' Now add all the columns to the ASRSysViewColumns
      fOK = CopyViewColumns_Transaction(lngViewID, frmViewProperties.ViewID)
      
      ' Finally select the new node and then release the reference
      nodX.EnsureVisible
      Set trvTables.SelectedItem = nodX
      trvTables_NodeClick trvTables.SelectedItem
    
      Set nodX = Nothing
  
    End If
  End If
  
TidyUpAndExit:
  Set frmViewProperties = Nothing
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
End Function

Private Function UpdateRightPane() As Boolean
  ' Update the right pane (views listview or columns grid) depending on the
  ' treeview seelction.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  '# Roy - temp fix as when form first loads, trvtables.selecteditem is nothing, despite node
  '# 1 actually being selected
  
  If trvTables.Nodes.Count > 0 And trvTables.SelectedItem Is Nothing Then trvTables.SelectedItem = trvTables.Nodes(1)
  
  If trvTables.SelectedItem Is Nothing Then
    lstViews.Visible = False
    lstColumns.Visible = False
    lblRightPaneCaption.Caption = ""
  Else
    Select Case trvTables.SelectedItem.DataKey
      Case "TABLE"
        ' Display the views listview.
        lstColumns.Visible = False
        lstViews.Visible = True
        lstViews_Refresh
      
      Case "VIEW"
        ' Display the columns grid.
        lstViews.Visible = False
        lstColumns.Visible = True
        lstColumns_Refresh
    End Select
  End If
  
TidyUpAndExit:
  UpdateRightPane = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Sub lstViews_Refresh()
  ' Populate the listview with the views for the selected table.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim lngTableID As Long
  Dim sTableName As String
  Dim ThisItem As ComctlLib.ListItem

  fOK = True
  
  ' Clear the list view items.
  lstViews.ListItems.Clear
  
  If Not trvTables.SelectedItem Is Nothing Then
    lngTableID = Right(trvTables.SelectedItem.key, Len(trvTables.SelectedItem.key) - 1)
  
    ' Get the selected table's details..
    With recTabEdit
      .Index = "idxTableID"
      .Seek "=", lngTableID
      
      fOK = Not .NoMatch
      If fOK Then
        sTableName = .Fields("tableName")
      End If
    End With
    
    If fOK Then
      lblRightPaneCaption.Caption = " '" & sTableName & "' table views"
  
      With recViewEdit
      
        .Index = "idxViewID"
        On Error Resume Next  ' Ignore an error on movefirst if no records in table
        .MoveFirst
        On Error GoTo 0
        
        While Not .EOF
          ' Check that it is not marked for deletion
          If (Not .Fields("Deleted")) And _
            (.Fields("viewTableID") = lngTableID) Then
            Set ThisItem = lstViews.ListItems.Add(, _
              "V" & .Fields("viewID"), .Fields("viewName"), "IMG_VIEW", "IMG_VIEW")
            'NHRD09092003 Fault 6505 Replaces carriages return boxes with spaces.
            'ThisItem.SubItems(1) = .Fields("viewDescription")
            ThisItem.SubItems(1) = Replace(.Fields("viewDescription"), vbCrLf, "  ")
            
          End If
          
          .MoveNext
        Wend
      
      End With
 
      ' If no items are selected then try to select the first one.
      If (lstViews_SelectedCount = 0) And (lstViews.ListItems.Count > 0) Then
        lstViews.SelectedItem = lstViews.ListItems(1)
        lstViews.SelectedItem.EnsureVisible
      End If
      
    End If
  End If
  
  ' Refresh the status bar.
  RefreshStatusBar

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub

Private Function lstColumns_Refresh() As Boolean
  ' Populate the listbox with the columns for the selected view.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fInView As Boolean
  Dim fAllColumns As Boolean
  Dim lngViewID As Long
  Dim lngTableID As Long
  Dim sSQL As String
  Dim sMessage As String
  Dim rsColumns As DAO.Recordset
  
  sMessage = ""
  
  gfRefreshing = True
  UI.LockWindow lstColumns.hWnd
  
  ' Remove all the columns from the listbox.
  lstColumns.Clear
  
  ' Get the view ID and the associated table ID for the selected view.
  lngViewID = Right(trvTables.SelectedItem.key, Len(trvTables.SelectedItem.key) - 1)
  With recViewEdit
    .Index = "idxViewID"
    .Seek "=", lngViewID
    fOK = Not .NoMatch
    If fOK Then
      lngTableID = .Fields("viewTableID")
      sMessage = " '" & Trim(.Fields("viewName")) & "' view columns"
    End If
  End With
      
  ' Update the caption.
  lblRightPaneCaption.Caption = sMessage
  
  ' Get the columns for the selected table.
  If fOK Then
    sSQL = "SELECT tmpColumns.*" & _
      " FROM tmpColumns" & _
      " WHERE tmpColumns.tableID = " & Trim(Str(lngTableID)) & _
      " AND tmpColumns.columnType <> " & Trim(Str(giCOLUMNTYPE_SYSTEM)) & _
      " AND tmpColumns.columnType <> " & Trim(Str(giCOLUMNTYPE_LINK)) & _
      " AND tmpColumns.deleted = FALSE" & _
      " ORDER BY columnName"
    Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    
    ' Add the columns to the grid.
    fAllColumns = True
    With rsColumns
      While Not .EOF
      
        recViewColEdit.Index = "idxViewColID"
        recViewColEdit.Seek "=", lngViewID, .Fields("columnID")
        fInView = Not recViewColEdit.NoMatch
        If fInView Then
          fInView = recViewColEdit("InView")
        End If
        lstColumns.AddItem .Fields("ColumnName")
        lstColumns.ItemData(lstColumns.NewIndex) = .Fields("ColumnID")
        lstColumns.Selected(lstColumns.NewIndex) = fInView
        
        If Not fInView Then fAllColumns = False
        
        .MoveNext
      Wend
      .Close
    End With
    Set rsColumns = Nothing

    fAllColumns = fAllColumns And (lstColumns.ListCount > 0)
    ' Add the 'all columns' column.
    lstColumns.AddItem "<All>", 0
    lstColumns.ItemData(lstColumns.NewIndex) = 0
    lstColumns.Selected(lstColumns.NewIndex) = fAllColumns
    ' See if all the screens are all selected.
    lstColumns.Enabled = (lstColumns.ListCount > 1)

    ' Select the first item.
    If lstColumns.Enabled Then
      lstColumns.ListIndex = 0
    End If

  End If
  
  ' Refresh the status bar.
  RefreshStatusBar

TidyUpAndExit:
  UI.UnlockWindow
  gfRefreshing = False
  ' Disassociate object variables.
  Set rsColumns = Nothing
  lstColumns_Refresh = fOK
  Exit Function
  
ErrorTrap:
  ' indicate that the function has failed
  fOK = False
  MsgBox Err.Description, vbCritical + vbOKOnly, App.Title
  Resume TidyUpAndExit
  
End Function


Public Function lstViews_SelectedCount() As Integer
  ' Return the count of views selected in the listview.
  Dim iLoop As Integer
  
  lstViews_SelectedCount = 0
  
  ' Loop through the list view items counting how many
  ' are currently selected.
  For iLoop = 1 To lstViews.ListItems.Count
    If lstViews.ListItems(iLoop).Selected = True Then
      lstViews_SelectedCount = lstViews_SelectedCount + 1
    End If
  Next iLoop

End Function


Public Function lstColumns_SelectedCount() As Integer
  ' Return the count of columns selected in the listbox.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iCount As Integer
  Dim iSelectedCount As Integer
  
  fOK = True
  iSelectedCount = 0
  
  ' Loop through the columns grid counting those selected.
  For iCount = 1 To lstColumns.ListCount - 1
    If lstColumns.Selected(iCount) Then
      iSelectedCount = iSelectedCount + 1
    End If
  Next iCount

TidyUpAndExit:
  If fOK Then
    lstColumns_SelectedCount = iSelectedCount
  Else
    lstColumns_SelectedCount = 0
  End If
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function




Public Property Get ActiveView() As Control
  ' Return the active view control.
  Set ActiveView = gctlActiveView
  
End Property

Private Function AddViewColumns_Transaction(plngViewID As Long, plngViewTableID As Long) As Boolean
  ' Add the table's columns to the view definition in the local database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iCount As Integer
  Dim lngNewID As Long
  Dim lngTableID As Long
  Dim sSQL As String
  Dim sNodeLabel As String
  Dim frmViewProperties As frmViewProp
  Dim nodX As SSNode
  Dim rsColumns As DAO.Recordset
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  recColEdit.Index = "idxTableID"
  recColEdit.Seek "=", plngViewTableID
  
  fOK = Not recColEdit.NoMatch
  
  If fOK Then
        
    Do While Not recColEdit.EOF

      ' If no more columns for this table exit loop
      If recColEdit!TableID <> plngViewTableID Then
        Exit Do
      End If
  
      ' Don't add deleted columns
      If recColEdit!Deleted <> True Then
      
        'Add column details to Columns table
        With recViewColEdit
          .AddNew
          .Fields("ViewID") = plngViewID
          .Fields("ColumnID") = recColEdit.Fields("ColumnID")
          .Fields("InView") = True
          .Fields("New") = True
          .Fields("Deleted") = False
          .Fields("Changed") = False
          .Update
        End With
      End If
      
      recColEdit.MoveNext
    Loop
  End If

TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  AddViewColumns_Transaction = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function CopyViewColumns_Transaction(plngOriginalViewID As Long, plngNewViewID As Long) As Boolean
  ' Add the table's columns to the view definition in the local database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  sSQL = "INSERT INTO tmpViewColumns ( ViewID, ColumnID, InView, changed, new, deleted )"
  sSQL = sSQL & "SELECT " & plngNewViewID
  sSQL = sSQL & "     , tmpViewColumns.ColumnID "
  sSQL = sSQL & "     , tmpViewColumns.InView "
  sSQL = sSQL & "     , tmpViewColumns.changed "
  sSQL = sSQL & "     , tmpViewColumns.new "
  sSQL = sSQL & "     , tmpViewColumns.deleted "
  sSQL = sSQL & "FROM tmpViewColumns "
  sSQL = sSQL & "WHERE Deleted = 0 "
  sSQL = sSQL & "AND tmpViewColumns.ViewID = " & plngOriginalViewID

  daoDb.Execute sSQL, dbFailOnError

  fOK = (daoDb.RecordsAffected > 0)
  
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  CopyViewColumns_Transaction = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function ChangeView(ViewStyle As ComctlLib.ListViewConstants) As Boolean

  Static InChangeView As Boolean
  On Error GoTo ErrorTrap

  If InChangeView Then Exit Function

  InChangeView = True

'TM20010914 Fault 1753
'As ActiveBar does not support mutual exclusivity on its tools, the following code
'ensures only one of the view options is selected at any one time.

  Me.lstViews.View = ViewStyle
  With abViewMgr
    Select Case ViewStyle
      Case lvwIcon
        .Tools("ID_LargeIcons").Checked = True
        .Tools("ID_SmallIcons").Checked = False
        .Tools("ID_List").Checked = False
        .Tools("ID_Details").Checked = False
        frmSysMgr.tbMain.Tools("ID_LargeIcons").Checked = True
        frmSysMgr.tbMain.Tools("ID_SmallIcons").Checked = False
        frmSysMgr.tbMain.Tools("ID_List").Checked = False
        frmSysMgr.tbMain.Tools("ID_Details").Checked = False
        .Tools("ID_CustomiseColumns").Enabled = False
      
      Case lvwSmallIcon
        .Tools("ID_LargeIcons").Checked = False
        .Tools("ID_SmallIcons").Checked = True
        .Tools("ID_List").Checked = False
        .Tools("ID_Details").Checked = False
        frmSysMgr.tbMain.Tools("ID_LargeIcons").Checked = False
        frmSysMgr.tbMain.Tools("ID_SmallIcons").Checked = True
        frmSysMgr.tbMain.Tools("ID_List").Checked = False
        frmSysMgr.tbMain.Tools("ID_Details").Checked = False
        .Tools("ID_CustomiseColumns").Enabled = False
        
      Case lvwList
        .Tools("ID_LargeIcons").Checked = False
        .Tools("ID_SmallIcons").Checked = False
        .Tools("ID_List").Checked = True
        .Tools("ID_Details").Checked = False
        frmSysMgr.tbMain.Tools("ID_LargeIcons").Checked = False
        frmSysMgr.tbMain.Tools("ID_SmallIcons").Checked = False
        frmSysMgr.tbMain.Tools("ID_List").Checked = True
        frmSysMgr.tbMain.Tools("ID_Details").Checked = False
        .Tools("ID_CustomiseColumns").Enabled = False
        
      Case lvwReport
        .Tools("ID_LargeIcons").Checked = False
        .Tools("ID_SmallIcons").Checked = False
        .Tools("ID_List").Checked = False
        .Tools("ID_Details").Checked = True
        frmSysMgr.tbMain.Tools("ID_LargeIcons").Checked = False
        frmSysMgr.tbMain.Tools("ID_SmallIcons").Checked = False
        frmSysMgr.tbMain.Tools("ID_List").Checked = False
        frmSysMgr.tbMain.Tools("ID_Details").Checked = True
        .Tools("ID_CustomiseColumns").Enabled = True
    
    End Select
  End With

  InChangeView = False

  ChangeView = True

  Exit Function

ErrorTrap:
  ChangeView = False
  Err = False

End Function

Private Sub SetColumnSizes()

  Dim iCount As Integer
 
  With Me.lstViews
    For iCount = 2 To .ColumnHeaders.Count Step 1
    
      If gpropShowColumns_ViewMgr(.ColumnHeaders.Item(iCount).Text) = True Then
        .ColumnHeaders(iCount).Width = (Len(.ColumnHeaders(iCount).Text) + 1) * UI.GetAvgCharWidth(Me.hDC)
      Else
        .ColumnHeaders(iCount).Width = 0
      End If
    Next iCount
  End With
  
End Sub
