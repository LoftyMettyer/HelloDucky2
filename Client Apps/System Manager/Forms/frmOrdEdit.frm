VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmOrdEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Order Definition"
   ClientHeight    =   8010
   ClientLeft      =   -390
   ClientTop       =   1950
   ClientWidth     =   10905
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5019
   Icon            =   "frmOrdEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picDocument 
      Height          =   510
      Left            =   1980
      Picture         =   "frmOrdEdit.frx":000C
      ScaleHeight     =   450
      ScaleWidth      =   465
      TabIndex        =   27
      Top             =   7485
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   9600
      TabIndex        =   8
      Top             =   7485
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   8340
      TabIndex        =   7
      Top             =   7485
      Width           =   1200
   End
   Begin TabDlg.SSTab sstabOrderDefinition 
      Height          =   7335
      Left            =   45
      TabIndex        =   9
      Top             =   45
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   12938
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Find Window Columns"
      TabPicture(0)   =   "frmOrdEdit.frx":0596
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraPageContainer(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Sort Order Columns"
      TabPicture(1)   =   "frmOrdEdit.frx":05B2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraPageContainer(1)"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraPageContainer 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   6800
         Index           =   1
         Left            =   -74850
         TabIndex        =   23
         Top             =   450
         Width           =   10545
         Begin VB.Frame fraSortColumns 
            Caption         =   "Sort Order Columns :"
            Height          =   6300
            Left            =   6000
            TabIndex        =   25
            Top             =   400
            Width           =   4400
            Begin ComctlLib.TreeView trvSelectedSortColumns 
               DragIcon        =   "frmOrdEdit.frx":05CE
               Height          =   5910
               Left            =   150
               TabIndex        =   18
               Top             =   250
               Width           =   4150
               _ExtentX        =   7329
               _ExtentY        =   10425
               _Version        =   327682
               HideSelection   =   0   'False
               LabelEdit       =   1
               Style           =   1
               ImageList       =   "ImageList1"
               Appearance      =   1
            End
         End
         Begin VB.TextBox txtOrderName 
            Height          =   315
            Index           =   1
            Left            =   780
            TabIndex        =   10
            Top             =   0
            Width           =   2670
         End
         Begin VB.Frame fraSortColumnsSource 
            Caption         =   "Columns :"
            Height          =   6300
            Left            =   0
            TabIndex        =   24
            Top             =   400
            Width           =   4400
            Begin ComctlLib.TreeView trvSortColumns 
               DragIcon        =   "frmOrdEdit.frx":0718
               Height          =   5910
               Left            =   150
               TabIndex        =   11
               Top             =   250
               Width           =   4150
               _ExtentX        =   7329
               _ExtentY        =   10425
               _Version        =   327682
               HideSelection   =   0   'False
               LabelEdit       =   1
               Style           =   7
               ImageList       =   "imglstTreeviewImages"
               Appearance      =   1
            End
         End
         Begin VB.CommandButton cmdSortColumnAscDesc 
            Caption         =   "Asc. / D&esc."
            Height          =   400
            Left            =   4500
            TabIndex        =   17
            Top             =   3800
            Width           =   1380
         End
         Begin VB.CommandButton sscmdAddSortColumn 
            Caption         =   "&Add"
            Height          =   405
            Left            =   4500
            TabIndex        =   12
            Top             =   495
            Width           =   1380
         End
         Begin VB.CommandButton sscmdInsertSortColumn 
            Caption         =   "&Insert"
            Height          =   405
            Left            =   4500
            TabIndex        =   13
            Top             =   1005
            Width           =   1380
         End
         Begin VB.CommandButton sscmdRemoveSortColumn 
            Caption         =   "&Remove"
            Height          =   405
            Left            =   4500
            TabIndex        =   14
            Top             =   1500
            Width           =   1380
         End
         Begin VB.CommandButton sscmdMoveUpSortColumn 
            Caption         =   "Move &Up"
            Height          =   405
            Left            =   4500
            TabIndex        =   15
            Top             =   2400
            Width           =   1380
         End
         Begin VB.CommandButton sscmdMoveDownSortColumn 
            Caption         =   "Move &Down"
            Height          =   405
            Left            =   4500
            TabIndex        =   16
            Top             =   2895
            Width           =   1380
         End
         Begin VB.Label lblOrderName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            Height          =   195
            Index           =   1
            Left            =   45
            TabIndex        =   26
            Top             =   60
            Width           =   510
         End
      End
      Begin VB.Frame fraPageContainer 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   6800
         Index           =   0
         Left            =   150
         TabIndex        =   19
         Top             =   450
         Width           =   10545
         Begin VB.Frame fraFindColumns 
            Caption         =   "Find Window Columns :"
            Height          =   6300
            Left            =   6000
            TabIndex        =   21
            Top             =   400
            Width           =   4400
            Begin VB.ListBox lstSelectedFindColumns 
               Height          =   5910
               Left            =   150
               Style           =   1  'Checkbox
               TabIndex        =   28
               Top             =   255
               Width           =   4150
            End
         End
         Begin VB.TextBox txtOrderName 
            Height          =   315
            Index           =   0
            Left            =   780
            TabIndex        =   0
            Top             =   0
            Width           =   2670
         End
         Begin VB.Frame fraFindColumnsSource 
            Caption         =   "Columns :"
            Height          =   6300
            Left            =   0
            TabIndex        =   20
            Top             =   400
            Width           =   4400
            Begin ComctlLib.TreeView trvFindColumns 
               DragIcon        =   "frmOrdEdit.frx":0B5A
               Height          =   5910
               Left            =   150
               TabIndex        =   1
               Top             =   250
               Width           =   4150
               _ExtentX        =   7329
               _ExtentY        =   10425
               _Version        =   327682
               HideSelection   =   0   'False
               LabelEdit       =   1
               Style           =   7
               ImageList       =   "imglstTreeviewImages"
               Appearance      =   1
            End
         End
         Begin VB.CommandButton sscmdRemoveFindColumn 
            Caption         =   "&Remove"
            Height          =   405
            Left            =   4500
            TabIndex        =   4
            Top             =   1500
            Width           =   1380
         End
         Begin VB.CommandButton sscmdAddFindColumn 
            Caption         =   "&Add"
            Height          =   405
            Left            =   4500
            TabIndex        =   2
            Top             =   495
            Width           =   1380
         End
         Begin VB.CommandButton sscmdInsertFindColumn 
            Caption         =   "&Insert"
            Height          =   405
            Left            =   4500
            TabIndex        =   3
            Top             =   1005
            Width           =   1380
         End
         Begin VB.CommandButton sscmdMoveUpFindColumn 
            Caption         =   "Move &Up"
            Height          =   405
            Left            =   4500
            TabIndex        =   5
            Top             =   2400
            Width           =   1380
         End
         Begin VB.CommandButton sscmdMoveDownFindColumn 
            Caption         =   "Move &Down"
            Height          =   405
            Left            =   4500
            TabIndex        =   6
            Top             =   2895
            Width           =   1380
         End
         Begin VB.Label lblMakeEditable 
            Caption         =   "Check box to enable in-line column editing"
            Height          =   270
            Left            =   6135
            TabIndex        =   29
            Top             =   60
            Width           =   4005
         End
         Begin VB.Label lblOrderName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   22
            Top             =   60
            Width           =   510
         End
      End
   End
   Begin ComctlLib.ImageList imglstTreeviewImages 
      Left            =   1110
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmOrdEdit.frx":0CA4
            Key             =   "IMG_TABLE"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmOrdEdit.frx":11F6
            Key             =   "IMG_COLUMN"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4215
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmOrdEdit.frx":1548
            Key             =   "IMG_DOWN"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmOrdEdit.frx":1A9A
            Key             =   "IMG_UP"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOrdEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Order definition variables.
Private mobjOrder As Order

' Form handling variables.
Private mfCancelled As Boolean
Private mfColumnDrag As Boolean
Private mfChanged As Boolean

' Form handling constants.
Const PAGE_FINDCOLUMNS = 0
Const PAGE_SORTCOLUMNS = 1

Private mblnReadOnly As Boolean

Public Property Get Cancelled() As Boolean
  ' Return the Cancelled property.
  Cancelled = mfCancelled
  
End Property

Public Property Get Order() As Order
  ' Return the Order object.
  Set Order = mobjOrder
  
End Property

Public Property Set Order(pobjOrder As Order)
  ' Set the Order object.
  Dim sIcon As String
  Dim objOrdItem As OrderItem
  Dim objNewNode As ComctlLib.Node
  'Dim iSequence As Integer

  ' Set the Order object global variable.
  Set mobjOrder = pobjOrder

  ' Set the order caption dependent on whether it is a new order or not.
  Me.Caption = "Order Definition"
  
  ' Display the order name.
  txtOrderName(PAGE_FINDCOLUMNS).Text = Trim(mobjOrder.OrderName)
  txtOrderName(PAGE_SORTCOLUMNS).Text = Trim(mobjOrder.OrderName)
  
  ' Populate the Find and Sort Columns treeviews.
  PopulateTreeViews
  
  ' Add items to the order and find columns listviews as required.
 ' For iSequence = 0 To mobjOrder.OrderItems.Count
    For Each objOrdItem In mobjOrder.OrderItems
     ' If objOrdItem.Sequence = iSequence Then
        If objOrdItem.ItemType = "O" Then
          sIcon = IIf(objOrdItem.Ascending, "IMG_UP", "IMG_DOWN")
          Set objNewNode = trvSelectedSortColumns.Nodes.Add(, , , objOrdItem.FullColumnName, sIcon)
          objNewNode.Tag = objOrdItem.ColumnID
          Set objNewNode = Nothing
        
          ' Remove the item from the Sort Coumns treeview.
          trvSortColumns.Nodes.Remove "C" & Trim(Str(objOrdItem.ColumnID))
        Else
                    
          With lstSelectedFindColumns
            .AddItem objOrdItem.FullColumnName
            .ItemData(.NewIndex) = objOrdItem.ColumnID
            .Selected(.NewIndex) = objOrdItem.Editable
          End With
      
          ' Remove the item from the Find Coumns treeview.
          trvFindColumns.Nodes.Remove "C" & Trim(Str(objOrdItem.ColumnID))
        End If
      'End If
    Next objOrdItem
    Set objOrdItem = Nothing
  'Next iSequence
  
  lstSelectedFindColumns.ListIndex = 0
  
  With trvSelectedSortColumns
    If .Nodes.Count > 0 Then
      .Nodes.Item(1).Selected = True
      'TM20010921 Fault 2030 & 'TM20010921 Fault 2031
      'The following line was off-centring the items
'      .SelectedItem.EnsureVisible
    End If
  End With
  
  ' Ensure the first page is selected.
  sstabOrderDefinition.Tab = PAGE_FINDCOLUMNS
  FindColumns_RefreshControls
  
  If (mobjOrder.OrderID = 0) And _
    (Len(mobjOrder.OrderName) > 0) Then
    ' ie. if we are copying an existing expression.
    mfChanged = True
  Else
    mfChanged = False
  End If
  
End Property

Private Sub cmdCancel_Click()

  Dim intAnswer As Integer
  
  ' Check if any changes have been made.
  If mfChanged Then
    intAnswer = MsgBox("The order definition has changed.  Save changes ?", vbQuestion + vbYesNoCancel + vbDefaultButton1, App.ProductName)
    If intAnswer = vbYes Then
      If Me.cmdOK.Enabled Then
        Call cmdOK_Click
        Exit Sub
      Else
        If (Len(Me.txtOrderName(0).Text) = 0) Then
          MsgBox "Invalid Order Name", vbExclamation + vbOKOnly, App.Title
        Else
          MsgBox "You must define both a find window order and a sort order" & vbCrLf & _
                 "for this table.", vbExclamation + vbOKOnly, App.Title
        End If
        Exit Sub
      End If
    ElseIf intAnswer = vbCancel Then
      Exit Sub
    End If
  End If
  
  ' Set the Cancelled property and unload the form.
  mfCancelled = True
  UnLoad Me
  
End Sub

Private Sub cmdOK_Click()
  ' Confirm the order.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iSequence As Integer
  Dim objNode As ComctlLib.Node
  
  fOK = True
  
  If mfChanged Then
    ' Reset the Cancelled property.
    mfCancelled = False
    
    ' Validate the order name.
    fOK = (Len(Trim(txtOrderName(0).Text)) > 1)
    If Not fOK Then
      MsgBox "Invalid order name.", vbOKOnly + vbExclamation, Application.Name
    Else
      ' Check that there are no other orders for this table with this name.
      With recOrdEdit
        .Index = "idxTableID"
        .Seek "=", Order.TableID
        
        If Not .NoMatch Then
          Do While Not .EOF
            
            If !TableID <> Order.TableID Then
              Exit Do
            End If
            
            If (!Name = Trim(txtOrderName(0).Text)) And _
              (!OrderID <> Order.OrderID) And _
              (!Type = Order.OrderType) Then
              MsgBox "An order named '" & Trim(txtOrderName(0).Text) & "' already exists !", vbOKOnly + vbExclamation, Application.Name
              fOK = False
              Exit Do
            End If
            
            .MoveNext
          Loop
        End If
      End With
    End If
  
    If fOK Then
      ' Write the changes to the order object.
      mobjOrder.OrderName = Trim(txtOrderName(0).Text)
      mobjOrder.ClearOrderItems
      
      If trvSelectedSortColumns.Nodes.Count > 0 Then
        Set objNode = trvSelectedSortColumns.Nodes.Item(1).FirstSibling
        iSequence = 0
        Do While Not objNode Is Nothing
          iSequence = iSequence + 1
          mobjOrder.AddOrderItem objNode.Tag, "O", iSequence, (objNode.Image = "IMG_UP"), objNode.Text, False
          Set objNode = objNode.Next
        Loop
      End If
            
      With lstSelectedFindColumns
        For iSequence = 0 To .ListCount - 1
          mobjOrder.AddOrderItem .ItemData(iSequence), "F", iSequence, True, .List(iSequence), .Selected(iSequence)
        Next
      End With
           
    End If
    
  Else
    mfCancelled = True
  End If
  
TidyUpAndExit:
  If fOK Then
    ' Unload the form.
    UnLoad Me
  End If
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub

Private Sub cmdSortColumnAscDesc_Click()
  ' Toggle the order of the selected item in the listview.
  If Not trvSelectedSortColumns.SelectedItem Is Nothing Then
    With trvSelectedSortColumns.SelectedItem
      .Image = IIf(.Image = "IMG_UP", "IMG_DOWN", "IMG_UP")
    End With
  
    mfChanged = True
  End If

End Sub

Private Sub Form_Initialize()
  
  mblnReadOnly = (Application.AccessMode <> accFull And _
                  Application.AccessMode <> accSupportMode)
  
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
  Dim iLoop As Integer

  ' Change the colur of the page frame containers.
  ' They are dark grey so that you developers can see them in VB, but
  ' need to be changed to be the same colour as the form so that the
  ' user doesn't notice them.
  For iLoop = fraPageContainer.LBound To fraPageContainer.UBound
    fraPageContainer(iLoop).BackColor = Me.BackColor
  Next iLoop
  
  trvFindColumns.DragIcon = picDocument.Picture
  trvSelectedSortColumns.DragIcon = picDocument.Picture
  trvSortColumns.DragIcon = picDocument.Picture

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  Dim intAnswer As Integer
  
  ' Set the Cancelled property.
  If UnloadMode <> vbFormCode Then
    'Check if any changes have been made.
    If mfChanged Then
      intAnswer = MsgBox("The order definition has changed.  Save changes ?", vbQuestion + vbYesNoCancel + vbDefaultButton1, App.ProductName)
      If intAnswer = vbYes Then
        If Me.cmdOK.Enabled Then
          Call cmdOK_Click
          If mfCancelled = True Then Cancel = 1
        Else
          If (Len(Me.txtOrderName(0).Text) = 0) Then
            MsgBox "Invalid Order Name", vbExclamation + vbOKOnly, App.Title
          Else
            MsgBox "You must define both a find window order and a sort order" & vbCrLf & _
                   "for this table.", vbExclamation + vbOKOnly, App.Title
          End If
          Cancel = 1
        End If
      ElseIf intAnswer = vbNo Then
        mfCancelled = True
      ElseIf intAnswer = vbCancel Then
        Cancel = 1
      End If
    Else
      mfCancelled = True
    End If
  End If
    
End Sub

Private Sub PopulateTreeViews()
  ' Populate the Find Columns treeview.
  Dim lngTableID As Long
  Dim sSQL As String
  Dim rsInfo As DAO.Recordset
  Dim objNode As ComctlLib.Node
  
  lngTableID = 0
  
  ' Clear the treeview.
  trvFindColumns.Nodes.Clear
  trvSortColumns.Nodes.Clear
  
  ' Get the list of columns for the order's base table.
  sSQL = "SELECT tmpColumns.tableID, tmpColumns.columnID, tmpColumns.columnName, tmpTables.tableName, tmpColumns.MaxOleSizeEnabled, tmpColumns.OLEType, tmpColumns.DataType" & _
    " FROM tmpColumns, tmpTables" & _
    " WHERE tmpColumns.tableID = " & Trim(Str(mobjOrder.TableID)) & _
    " AND tmpColumns.deleted = FALSE" & _
    " AND tmpColumns.ColumnType <> " & Trim(Str(giCOLUMNTYPE_SYSTEM)) & _
    " AND tmpColumns.ColumnType <> " & Trim(Str(giCOLUMNTYPE_LINK)) & _
    " AND tmpTables.tableID = tmpColumns.tableID"
  Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  

  Do While Not rsInfo.EOF
    ' Add the table root node if it hasn't already been added.
    If lngTableID <> rsInfo!TableID Then
      lngTableID = rsInfo!TableID
    
      Set objNode = trvFindColumns.Nodes.Add(, tvwChild, _
        "T" & Trim(Str(mobjOrder.TableID)), rsInfo!TableName, "IMG_TABLE", "IMG_TABLE")
      objNode.Sorted = True
      objNode.Expanded = True
      Set objNode = Nothing
    
      Set objNode = trvSortColumns.Nodes.Add(, tvwChild, _
        "T" & Trim(Str(mobjOrder.TableID)), rsInfo!TableName, "IMG_TABLE", "IMG_TABLE")
      objNode.Sorted = True
      objNode.Expanded = True
      Set objNode = Nothing
    End If
    
    ' Add items to the treeview for each column in the order's base table.
    If rsInfo!DataType = -3 Or rsInfo!DataType = -4 Then
      If rsInfo!OLEType <> 2 Or rsInfo!MaxOleSizeEnabled <> True Then GoTo continueloop
    End If
    
    Set objNode = trvFindColumns.Nodes.Add("T" & Trim(Str(rsInfo!TableID)), _
      tvwChild, "C" & Trim(Str(rsInfo!ColumnID)), _
      rsInfo!ColumnName, "IMG_COLUMN", "IMG_COLUMN")
    objNode.Tag = rsInfo!ColumnID
    Set objNode = Nothing
    
    Set objNode = trvSortColumns.Nodes.Add("T" & Trim(Str(rsInfo!TableID)), _
      tvwChild, "C" & Trim(Str(rsInfo!ColumnID)), _
      rsInfo!ColumnName, "IMG_COLUMN", "IMG_COLUMN")
    objNode.Tag = rsInfo!ColumnID
    Set objNode = Nothing
    
continueloop:
  
    
    
    rsInfo.MoveNext
  Loop
  rsInfo.Close
  Set rsInfo = Nothing
  
  ' Get the list of columns for the order's parent tables.
  sSQL = "SELECT tmpColumns.tableID, tmpColumns.columnID, tmpColumns.columnName, tmpTables.tableName" & _
    " FROM tmpColumns, tmpTables, tmpRelations" & _
    " WHERE tmpRelations.childID = " & Trim(Str(mobjOrder.TableID)) & _
    " AND tmpRelations.parentID = tmpColumns.tableID" & _
    " AND tmpColumns.deleted = FALSE" & _
    " AND tmpColumns.ColumnType <> " & Trim(Str(giCOLUMNTYPE_SYSTEM)) & _
    " AND tmpColumns.ColumnType <> " & Trim(Str(giCOLUMNTYPE_LINK)) & _
    " AND tmpTables.tableID = tmpColumns.tableID" & _
    " ORDER BY tmpTables.tableName, tmpColumns.columnName"
  Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    
  Do While Not rsInfo.EOF
    ' Add the table root node if it hasn't already been added.
    If lngTableID <> rsInfo!TableID Then
      lngTableID = rsInfo!TableID
    
      Set objNode = trvFindColumns.Nodes.Add(, tvwChild, _
        "T" & Trim(Str(rsInfo!TableID)), rsInfo!TableName, "IMG_TABLE", "IMG_TABLE")
      objNode.Sorted = True
      objNode.Expanded = False
      Set objNode = Nothing
    
      Set objNode = trvSortColumns.Nodes.Add(, tvwChild, _
        "T" & Trim(Str(rsInfo!TableID)), rsInfo!TableName, "IMG_TABLE", "IMG_TABLE")
      objNode.Sorted = True
      objNode.Expanded = False
      Set objNode = Nothing
    End If
    
    ' Add items to the treeview for each column in the order's base table.
    Set objNode = trvFindColumns.Nodes.Add("T" & Trim(Str(rsInfo!TableID)), _
      tvwChild, "C" & Trim(Str(rsInfo!ColumnID)), _
      rsInfo!ColumnName, "IMG_COLUMN", "IMG_COLUMN")
    objNode.Tag = rsInfo!ColumnID
    Set objNode = Nothing
    
    Set objNode = trvSortColumns.Nodes.Add("T" & Trim(Str(rsInfo!TableID)), _
      tvwChild, "C" & Trim(Str(rsInfo!ColumnID)), _
      rsInfo!ColumnName, "IMG_COLUMN", "IMG_COLUMN")
    objNode.Tag = rsInfo!ColumnID
    Set objNode = Nothing
    
    rsInfo.MoveNext
  Loop
  rsInfo.Close
  Set rsInfo = Nothing
    
  ' Select the first item in each treeview.
  With trvFindColumns
    If .Nodes.Count > 0 Then
      .Nodes(1).Selected = True
      .SelectedItem.EnsureVisible
    End If
  End With
  
  With trvSortColumns
    If .Nodes.Count > 0 Then
      .Nodes(1).Selected = True
      'TM20010921 Fault 2030 & 'TM20010921 Fault 2031
      'The following line was off-centring the items
'      .SelectedItem.EnsureVisible
    End If
  End With

End Sub

Private Sub FindColumns_RefreshControls()
  ' Refesh the controls whose status is variable.
  Dim fSelectedFirstNode As Boolean
  Dim fSelectedLastNode As Boolean
  Dim fSelectedColumnValid As Boolean
  Dim fSelectedFindColumnValid As Boolean

  If mblnReadOnly Then
    Exit Sub
  End If

  ' Check that we have a valid column selected.
  fSelectedColumnValid = Not (trvFindColumns.SelectedItem Is Nothing) And Not mblnReadOnly
  If fSelectedColumnValid Then
    fSelectedColumnValid = (Left(trvFindColumns.SelectedItem.key, 1) = "C")
  End If
  
  sscmdAddFindColumn.Enabled = fSelectedColumnValid
  sscmdInsertFindColumn.Enabled = fSelectedColumnValid
  
  ' Check if we have a 'selected column' selected.
  fSelectedFirstNode = (lstSelectedFindColumns.ListIndex = 0)
  fSelectedLastNode = (lstSelectedFindColumns.ListIndex = lstSelectedFindColumns.ListCount - 1)
  fSelectedFindColumnValid = (lstSelectedFindColumns.ListCount > 0)
  
  sscmdRemoveFindColumn.Enabled = fSelectedFindColumnValid
  sscmdMoveUpFindColumn.Enabled = fSelectedFindColumnValid And _
    (lstSelectedFindColumns.ListCount > 1) And _
    (Not fSelectedFirstNode)
  sscmdMoveDownFindColumn.Enabled = fSelectedFindColumnValid And _
    (lstSelectedFindColumns.ListCount > 1) And _
    (Not fSelectedLastNode)
    
  ' Disable the OK command control if there are no order items specified.
  cmdOK.Enabled = (lstSelectedFindColumns.ListCount > 0) And _
    (trvSelectedSortColumns.Nodes.Count > 0) And _
    (Len(Trim(txtOrderName(0).Text)) > 0) And Not mblnReadOnly

End Sub

Private Sub SortColumns_RefreshControls()
  ' Refesh the controls whose status is variable.
  Dim fSelectedFirstNode As Boolean
  Dim fSelectedLastNode As Boolean
  Dim fSelectedColumnValid As Boolean
  Dim fSelectedSortColumnValid As Boolean
  
  ' Check that we have a valid column selected.
  fSelectedColumnValid = Not (trvSortColumns.SelectedItem Is Nothing) And Not mblnReadOnly
  If fSelectedColumnValid Then
    fSelectedColumnValid = (Left(trvSortColumns.SelectedItem.key, 1) = "C")
  End If
  
  sscmdAddSortColumn.Enabled = fSelectedColumnValid
  sscmdInsertSortColumn.Enabled = fSelectedColumnValid
  
  ' Check if we have a 'selected column' selected.
  fSelectedFirstNode = False
  fSelectedLastNode = False
  fSelectedSortColumnValid = Not (trvSelectedSortColumns.SelectedItem Is Nothing) And Not mblnReadOnly
  If fSelectedSortColumnValid Then
    fSelectedFirstNode = (trvSelectedSortColumns.SelectedItem.Tag = trvSelectedSortColumns.SelectedItem.FirstSibling.Tag)
    fSelectedLastNode = (trvSelectedSortColumns.SelectedItem.Tag = trvSelectedSortColumns.SelectedItem.LastSibling.Tag)
  End If
  
  cmdSortColumnAscDesc.Enabled = fSelectedSortColumnValid
  sscmdRemoveSortColumn.Enabled = fSelectedSortColumnValid
  sscmdMoveUpSortColumn.Enabled = fSelectedSortColumnValid And _
    (trvSelectedSortColumns.Nodes.Count > 1) And _
    (Not fSelectedFirstNode)
  sscmdMoveDownSortColumn.Enabled = fSelectedSortColumnValid And _
    (trvSelectedSortColumns.Nodes.Count > 1) And _
    (Not fSelectedLastNode)
    
  ' Disable the OK command control if there are no order items specified.
  cmdOK.Enabled = (lstSelectedFindColumns.ListCount > 0) And _
    (trvSelectedSortColumns.Nodes.Count > 0) And _
    (Len(Trim(txtOrderName(0).Text)) > 0) And Not mblnReadOnly

End Sub































Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub lstSelectedFindColumns_Click()
  FindColumns_RefreshControls
End Sub

Private Sub lstSelectedFindColumns_ItemCheck(Item As Integer)
  mfChanged = True
End Sub

Private Sub lstSelectedFindColumns_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  FindColumns_RefreshControls
End Sub

Private Sub sscmdAddFindColumn_Click()
  ' Add the selected column to the find columns listview in the last position.
  Dim nodSelection As Node
  Dim objNewNode As ComctlLib.Node
         
  ' Do nothing if there is no selection in the treeview.
  If trvFindColumns.SelectedItem Is Nothing Then
    Exit Sub
  End If
  
  Set nodSelection = trvFindColumns.SelectedItem
  
  ' Do nothing if the selected node is a table node.
  If Left(nodSelection.key, 1) <> "C" Then
    Exit Sub
  End If
  
  ' Add the selected column to the Find Columns listview.
  With lstSelectedFindColumns
    .AddItem nodSelection.Parent.Text & "." & nodSelection.Text
    .ItemData(.NewIndex) = nodSelection.Tag
  End With

  ' Remove the column from the treeview.
  trvFindColumns.Nodes.Remove "C" & Trim(Str(nodSelection.Tag))
  
  ' Disassociate object variables.
  Set nodSelection = Nothing
  
  mfChanged = True
    
  FindColumns_RefreshControls
  
End Sub

Private Sub sscmdAddSortColumn_Click()
  ' Add the selected column to the sort columns listview in the last position.
  Dim objNewNode As ComctlLib.Node
  Dim nodSelection As Node
  
  ' Do nothing if there is no selection in the treeview.
  If trvSortColumns.SelectedItem Is Nothing Then
    Exit Sub
  End If

  Set nodSelection = trvSortColumns.SelectedItem
  
  ' Do nothing if the selected node is a table node.
  If Left(nodSelection.key, 1) <> "C" Then
    Exit Sub
  End If
    
  ' Add the selected column to the Sort Order columns listview.
  Set objNewNode = trvSelectedSortColumns.Nodes.Add(, , , nodSelection.Parent.Text & "." & nodSelection.Text, "IMG_UP")
  objNewNode.Tag = nodSelection.Tag
  objNewNode.Selected = True
  
  'TM20010921 Fault 2030 & 'TM20010921 Fault 2031
  'The following line was off-centring the items
  'objNewNode.EnsureVisible
  Set objNewNode = Nothing
  
  ' Remove the column from the treeview.
  trvSortColumns.Nodes.Remove "C" & Trim(Str(nodSelection.Tag))
  
  ' Disassociate object variables.
  Set nodSelection = Nothing
  
  mfChanged = True
    
  SortColumns_RefreshControls

End Sub

Private Sub sscmdInsertFindColumn_Click()
  ' Insert the selected column to the find columns listview in the selected position.
  Dim nodSelection As Node
  Dim objNewNode As ComctlLib.Node
  Dim objHighlightNode As ComctlLib.Node
  
  ' Do nothing if there is no selection in the treeview.
  If trvFindColumns.SelectedItem Is Nothing Then
    Exit Sub
  End If
  
  Set nodSelection = trvFindColumns.SelectedItem
  
  ' Do nothing if the selected node is a table node.
  If Left(nodSelection.key, 1) <> "C" Then
    Exit Sub
  End If
  
  ' Deselect the current listview selection.
  ' Add the selected column to the Find Columns listview.
  With lstSelectedFindColumns
    .AddItem nodSelection.Parent.Text & "." & nodSelection.Text
    .ItemData(.NewIndex) = nodSelection.Tag
  End With

  ' Ensure that it appears inserted
  MoveItemInListBox lstSelectedFindColumns, lstSelectedFindColumns.NewIndex, lstSelectedFindColumns.ListIndex

  ' Remove the column from the treeview.
  trvFindColumns.Nodes.Remove "C" & Trim(Str(nodSelection.Tag))
    
  ' Disassociate object variables.
  Set objNewNode = Nothing
  Set objHighlightNode = Nothing
  Set nodSelection = Nothing
  
  mfChanged = True
    
  FindColumns_RefreshControls
  
End Sub

Private Sub sscmdInsertSortColumn_Click()
  ' Insert the selected column to the Sort Order columns listview in the selected position.
  Dim nodSelection As Node
  Dim objNewNode As ComctlLib.Node
  Dim objHighlightNode As ComctlLib.Node
  
  ' Do nothing if there is no selection in the treeview.
  If trvSortColumns.SelectedItem Is Nothing Then
    Exit Sub
  End If

  Set nodSelection = trvSortColumns.SelectedItem
    
  ' Do nothing if the selected node is a table node.
  If Left(nodSelection.key, 1) <> "C" Then
    Exit Sub
  End If
    
  ' Deselect the current listview selection.
  ' Add the selected column to the Sort Columns listview.
  Set objHighlightNode = trvSelectedSortColumns.SelectedItem
  If Not objHighlightNode Is Nothing Then
    objHighlightNode.Selected = False
    Set objNewNode = trvSelectedSortColumns.Nodes.Add(objHighlightNode.Index, tvwPrevious, , nodSelection.Parent.Text & "." & nodSelection.Text, "IMG_UP")
  Else
    Set objNewNode = trvSelectedSortColumns.Nodes.Add(, , , nodSelection.Parent.Text & "." & nodSelection.Text, "IMG_UP")
  End If
  objNewNode.Tag = nodSelection.Tag
  objNewNode.Selected = True
  'TM20010921 Fault 2030 & 'TM20010921 Fault 2031
  'The following line was off-centring the items
'  objNewNode.EnsureVisible
    
  ' Remove the column from the treeview.
  trvSortColumns.Nodes.Remove "C" & Trim(Str(nodSelection.Tag))
  
  ' Disassociate object variables.
  Set objNewNode = Nothing
  Set objHighlightNode = Nothing
  Set nodSelection = Nothing
  
  mfChanged = True
    
  SortColumns_RefreshControls

End Sub

Private Sub sscmdMoveDownFindColumn_Click()
  
  ' Move the selected treeview item DOWN one position.
  MoveItemInListBox lstSelectedFindColumns, lstSelectedFindColumns.ListIndex, lstSelectedFindColumns.ListIndex + 1
  FindColumns_RefreshControls
  mfChanged = True
  
End Sub

Private Sub sscmdMoveDownSortColumn_Click()
  ' Move the selected treeview item DOWN one position.
  Dim objNode As ComctlLib.Node
  Dim objNewNode As ComctlLib.Node
  
  Set objNode = trvSelectedSortColumns.SelectedItem
  If Not objNode Is Nothing Then
    ' Move the selected item down one position in the treeview.
    If objNode.Tag <> objNode.LastSibling.Tag Then
      Set objNewNode = trvSelectedSortColumns.Nodes.Add(objNode.Next, tvwNext, , objNode.Text, objNode.Image)
      objNewNode.Tag = objNode.Tag
      objNewNode.Selected = True
      'TM20010921 Fault 2030 & 'TM20010921 Fault 2031
      'The following line was off-centring the items
'      objNewNode.EnsureVisible
      Set objNewNode = Nothing

      trvSelectedSortColumns.Nodes.Remove objNode.Index
      
      mfChanged = True
    End If
  End If
    
  ' Disassociate object variables.
  Set objNode = Nothing

  SortColumns_RefreshControls
  
End Sub

Private Sub sscmdMoveUpFindColumn_Click()

  ' Move the selected treeview item UP one position.
  MoveItemInListBox lstSelectedFindColumns, lstSelectedFindColumns.ListIndex, lstSelectedFindColumns.ListIndex - 1
  FindColumns_RefreshControls
  mfChanged = True

End Sub

Private Sub sscmdMoveUpSortColumn_Click()
  ' Move the selected treeview item DOWN one position.
  Dim objNode As ComctlLib.Node
  Dim objNewNode As ComctlLib.Node
  
  Set objNode = trvSelectedSortColumns.SelectedItem
  If Not objNode Is Nothing Then
    ' Move the selected item up one position in the treeview.
    If objNode.Tag <> objNode.FirstSibling.Tag Then
      Set objNewNode = trvSelectedSortColumns.Nodes.Add(objNode.Previous, tvwPrevious, , objNode.Text, objNode.Image)
      objNewNode.Tag = objNode.Tag
      objNewNode.Selected = True
      'TM20010921 Fault 2030 & 'TM20010921 Fault 2031
      'The following line was off-centring the items
'      objNewNode.EnsureVisible
      Set objNewNode = Nothing

      trvSelectedSortColumns.Nodes.Remove objNode.Index
      
      mfChanged = True
    End If
  End If
    
  ' Disassociate object variables.
  Set objNode = Nothing

  SortColumns_RefreshControls

End Sub

Private Sub sscmdRemoveFindColumn_Click()
  ' Remove the selected item from the Find Columns listview.
  Dim lngColumnID As Long
  Dim sSQL As String
  Dim objNode As ComctlLib.Node
  Dim rsInfo As DAO.Recordset
  
  lngColumnID = lstSelectedFindColumns.ItemData(lstSelectedFindColumns.ListIndex)
  lstSelectedFindColumns.RemoveItem (lstSelectedFindColumns.ListIndex)
  
  ' Get the columnName and tableID of the selected column.
  sSQL = "SELECT tmpColumns.tableID, tmpColumns.columnName" & _
    " FROM tmpColumns" & _
    " WHERE tmpColumns.columnID = " & Trim(Str(lngColumnID))
  Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  If Not (rsInfo.EOF And rsInfo.BOF) Then
    ' Add the column back into the treeview.
    Set objNode = trvFindColumns.Nodes.Add("T" & Trim(Str(rsInfo!TableID)), _
      tvwChild, "C" & Trim(Str(lngColumnID)), rsInfo!ColumnName, "IMG_COLUMN", "IMG_COLUMN")
    objNode.Tag = lngColumnID
    Set objNode = Nothing
  End If

  mfChanged = True
         
  FindColumns_RefreshControls

End Sub

Private Sub sscmdRemoveSortColumn_Click()
  ' Remove the selected item from the Sort Columns listview.
  Dim lngColumnID As Long
  Dim sSQL As String
  Dim objNode As ComctlLib.Node
  Dim rsInfo As DAO.Recordset
  
  With trvSelectedSortColumns
    If Not .SelectedItem Is Nothing Then
      lngColumnID = .SelectedItem.Tag
      
      .Nodes.Remove .SelectedItem.Index
      
      ' Get the columnName and tableID of the selected column.
      sSQL = "SELECT tmpColumns.tableID, tmpColumns.columnName" & _
        " FROM tmpColumns" & _
        " WHERE tmpColumns.columnID = " & Trim(Str(lngColumnID))
      Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    
      If Not (rsInfo.EOF And rsInfo.BOF) Then
        ' Add the column back into the treeview.
        Set objNode = trvSortColumns.Nodes.Add("T" & Trim(Str(rsInfo!TableID)), _
          tvwChild, "C" & Trim(Str(lngColumnID)), rsInfo!ColumnName, "IMG_COLUMN", "IMG_COLUMN")
        objNode.Tag = lngColumnID
        Set objNode = Nothing
      End If
    
      mfChanged = True
    End If
      
    If Not .SelectedItem Is Nothing Then
      .SelectedItem.Selected = True
    End If
  End With
  
  SortColumns_RefreshControls

End Sub

Private Sub sstabOrderDefinition_Click(PreviousTab As Integer)

  fraPageContainer(PAGE_FINDCOLUMNS).Enabled = (sstabOrderDefinition.Tab = PAGE_FINDCOLUMNS)
  fraPageContainer(PAGE_SORTCOLUMNS).Enabled = (sstabOrderDefinition.Tab = PAGE_SORTCOLUMNS)

  If sstabOrderDefinition.Tab = PAGE_FINDCOLUMNS Then
    FindColumns_RefreshControls
  Else
    SortColumns_RefreshControls
  End If

End Sub

Private Sub trvFindColumns_Click()
  If Not mblnReadOnly Then
    FindColumns_RefreshControls
  End If
End Sub

Private Sub trvFindColumns_DblClick()
  If Not mblnReadOnly Then
    sscmdAddFindColumn_Click
  End If
End Sub

Private Sub trvFindColumns_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  FindColumns_RefreshControls
End Sub

Private Sub trvFindColumns_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  FindColumns_RefreshControls
End Sub

Private Sub trvFindColumns_NodeClick(ByVal Node As ComctlLib.Node)
  FindColumns_RefreshControls
End Sub

Private Sub trvSelectedSortColumns_DblClick()
  If Not mblnReadOnly Then
    cmdSortColumnAscDesc_Click
  End If
End Sub

Private Sub trvSelectedSortColumns_DragDrop(Source As Control, x As Single, y As Single)
  ' Drop a selected item from the columns treeview into the selected columns treeview.
  Dim fDropOk As Boolean
  Dim objHighlightNode As ComctlLib.Node
  Dim objNewNode As ComctlLib.Node
  Dim objOldNode As ComctlLib.Node

  If mblnReadOnly Then
    Exit Sub
  End If
  
  
  Set objHighlightNode = trvSelectedSortColumns.DropHighlight

  If Source Is trvSortColumns Then
    fDropOk = Not (trvSortColumns.SelectedItem Is Nothing)

    If fDropOk Then
      If Not trvSelectedSortColumns.SelectedItem Is Nothing Then
        trvSelectedSortColumns.SelectedItem.Selected = False
      End If

      If objHighlightNode Is Nothing Then
        Set objNewNode = trvSelectedSortColumns.Nodes.Add(, , , trvSortColumns.SelectedItem.Parent.Text & "." & trvSortColumns.SelectedItem.Text, "IMG_UP")
      Else
        Set objNewNode = trvSelectedSortColumns.Nodes.Add(objHighlightNode, tvwPrevious, , trvSortColumns.SelectedItem.Parent.Text & "." & trvSortColumns.SelectedItem.Text, "IMG_UP")
      End If
      
      objNewNode.Tag = trvSortColumns.SelectedItem.Tag
      objNewNode.Selected = True
      'TM20010921 Fault 2030 & 'TM20010921 Fault 2031
      'The following line was off-centring the items
      'objNewNode.EnsureVisible
      
      ' Remove the column from the treeview.
      trvSortColumns.Nodes.Remove trvSortColumns.SelectedItem.Index
      
      trvSortColumns.Drag vbEndDrag
      Set objNewNode = Nothing
      trvSelectedSortColumns.SetFocus
      mfChanged = True
    Else
      trvSortColumns.Drag vbCancel
    End If
  ElseIf Source Is trvSelectedSortColumns Then
    If (objHighlightNode Is Nothing) Or _
      (trvSelectedSortColumns.SelectedItem Is objHighlightNode) Then

      trvSelectedSortColumns.Drag vbCancel
    Else
      Set objOldNode = trvSelectedSortColumns.SelectedItem
      Set objNewNode = trvSelectedSortColumns.Nodes.Add(objHighlightNode, tvwPrevious, , objOldNode.Text, objOldNode.Image)
      objNewNode.Tag = objOldNode.Tag
      objNewNode.Selected = True
      objNewNode.EnsureVisible

      trvSelectedSortColumns.Nodes.Remove objOldNode.Index
      
      Set objNewNode = Nothing
      Set objOldNode = Nothing
      mfChanged = True
    End If
  End If

  Set objHighlightNode = Nothing
  Set trvSelectedSortColumns.DropHighlight = Nothing

  SortColumns_RefreshControls

End Sub

Private Sub trvSelectedSortColumns_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  Dim objNode As ComctlLib.Node

  If mblnReadOnly Then
    Exit Sub
  End If

  'Get the item at the mouse's coordinates.
  Set objNode = trvSelectedSortColumns.HitTest(x, y)

  'TM20010921 Fault 2030 & 'TM20010921 Fault 2031
  'The following lines were off-centring the items
  ' Check if the item at the mouse's coordinates is a control.
  '  If Not objNode Is Nothing Then
  '    objNode.EnsureVisible
  '  End If

  ' Set the DropHighlight node
  Set trvSelectedSortColumns.DropHighlight = objNode

  Set objNode = Nothing

End Sub

Private Sub trvSelectedSortColumns_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyDelete Then
    sscmdRemoveSortColumn_Click
  End If

End Sub

Private Sub trvSelectedSortColumns_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
    sscmdRemoveSortColumn_Click
  End If

End Sub

Private Sub trvSelectedSortColumns_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim objNode As ComctlLib.Node

  If mblnReadOnly Then
    Exit Sub
  End If
  
  If Button = vbLeftButton Then
    ' Get the item at the mouse position
    Set objNode = trvSelectedSortColumns.HitTest(x, y)
    If Not objNode Is Nothing Then
      ' If this node is not the selected node, make it
      If Not objNode Is trvSelectedSortColumns.SelectedItem Then
        trvSelectedSortColumns.SelectedItem.Selected = False
        Set trvSelectedSortColumns.SelectedItem = objNode
      End If
    End If

    If trvSelectedSortColumns.Nodes.Count > 0 Then
      mfColumnDrag = True
      trvSelectedSortColumns.Drag vbBeginDrag
    End If
  End If

  SortColumns_RefreshControls

End Sub

Private Sub trvSelectedSortColumns_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  If mblnReadOnly Then
    Exit Sub
  End If
  
  If mfColumnDrag Then
    ' Reset the flag that shows that a column is being dragged.
    trvSelectedSortColumns.Drag vbCancel
    mfColumnDrag = False
  End If

  SortColumns_RefreshControls
  
End Sub

Private Sub trvSelectedSortColumns_NodeClick(ByVal Node As ComctlLib.Node)
  Node.Selected = True
  SortColumns_RefreshControls

End Sub

Private Sub trvSortColumns_Click()
  If Not mblnReadOnly Then
    SortColumns_RefreshControls
  End If
End Sub

Private Sub trvSortColumns_DblClick()
  If Not mblnReadOnly Then
    sscmdAddSortColumn_Click
  End If
End Sub

Private Sub trvSortColumns_DragDrop(Source As Control, x As Single, y As Single)
  
  If Not mblnReadOnly Then
    ' Remove the selected item from the columns listview.
    If Source Is trvSelectedSortColumns Then
      sscmdRemoveSortColumn_Click
    End If
    SortColumns_RefreshControls
  End If

End Sub

Private Sub trvSortColumns_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Start the drag-drop operation.
  Dim fGoodColumn As Boolean
  Dim objItem As ComctlLib.ListItem
  Dim nodSelection As Node

  If mblnReadOnly Then
    Exit Sub
  End If

  If Button = vbLeftButton Then
    'Get the item at the mouse position
    Set nodSelection = trvSortColumns.HitTest(x, y)
    If Not nodSelection Is Nothing Then
      'If this item is not the selected item, make it
      If Not nodSelection Is trvSortColumns.SelectedItem Then
        Set trvSortColumns.SelectedItem = nodSelection
      End If
    End If
  End If
  
  ' Do not drag anything if there is no selected item in the treeview.
  fGoodColumn = Not (trvSortColumns.SelectedItem Is Nothing)
  
  ' Do not drag the selected item if is not a column.
  If fGoodColumn Then
    fGoodColumn = (Left(trvSortColumns.SelectedItem.key, 1) = "C")
  End If
  
  If fGoodColumn Then
    ' Set the flag to show that a column is being dragged.
    mfColumnDrag = True
    trvSortColumns.Drag vbBeginDrag
  End If
  
  SortColumns_RefreshControls
  
End Sub

Private Sub trvSortColumns_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  If mblnReadOnly Then
    Exit Sub
  End If

  If mfColumnDrag Then
    ' Reset the flag that shows that a column is being dragged.
    trvSortColumns.Drag vbCancel
    mfColumnDrag = False
  End If

  SortColumns_RefreshControls

End Sub

Private Sub trvSortColumns_NodeClick(ByVal Node As ComctlLib.Node)
  SortColumns_RefreshControls

End Sub

Private Sub txtOrderName_Change(Index As Integer)
  Dim iLoop As Integer
  Dim sValidatedName As String
  Dim iSelStart As Integer
  Dim iSelLen As Integer
  
  'JPD 20090102 Fault 13484
  sValidatedName = ValidateName(txtOrderName(Index).Text)
  
  If sValidatedName <> txtOrderName(Index).Text Then
    iSelStart = txtOrderName(Index).SelStart
    iSelLen = txtOrderName(Index).SelLength
    
    txtOrderName(Index).Text = sValidatedName
    
    txtOrderName(Index).SelStart = iSelStart
    txtOrderName(Index).SelLength = iSelLen
  End If
  
  For iLoop = txtOrderName.LBound To txtOrderName.UBound
    If iLoop <> Index Then
      txtOrderName(iLoop).Text = txtOrderName(Index).Text
    End If
  Next iLoop

  mfChanged = True
  
  ' Disable the OK command control if there are no order items specified.
  cmdOK.Enabled = (lstSelectedFindColumns.ListCount > 0) And _
    (trvSelectedSortColumns.Nodes.Count > 0) And _
    (Len(Trim(txtOrderName(0).Text)) > 0) And Not mblnReadOnly

End Sub

Private Sub txtOrderName_GotFocus(Index As Integer)
  ' Select all text upon focus.
  UI.txtSelText

End Sub

Private Sub txtOrderName_KeyPress(Index As Integer, KeyAscii As Integer)
  ' Validate the character entered.
  KeyAscii = ValidNameChar(KeyAscii, txtOrderName(Index).SelStart)

End Sub

Private Sub MoveItemInListBox(ByRef objListbox As ListBox, iOldIndex As Integer, iNewIndex As Integer)
  
  Dim sItem As String
  Dim iItemData As Integer
  Dim bSelected As Boolean
  
  With objListbox
    bSelected = .Selected(iOldIndex)
    sItem = .List(iOldIndex)
    iItemData = .ItemData(iOldIndex)
    
    .RemoveItem iOldIndex
    .AddItem sItem, iNewIndex
    .Selected(iNewIndex) = bSelected
    .ItemData(iNewIndex) = iItemData
    
    .ListIndex = iNewIndex
  End With

End Sub
