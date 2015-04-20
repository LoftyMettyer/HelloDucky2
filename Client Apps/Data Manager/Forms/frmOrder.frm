VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOrder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Order Definition"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1048
   Icon            =   "frmOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picDocument 
      Height          =   510
      Left            =   2040
      Picture         =   "frmOrder.frx":000C
      ScaleHeight     =   450
      ScaleWidth      =   465
      TabIndex        =   28
      Top             =   4920
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   5385
      TabIndex        =   8
      Top             =   5000
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   6650
      TabIndex        =   9
      Top             =   5000
      Width           =   1200
   End
   Begin TabDlg.SSTab sstabOrderDefinition 
      Height          =   4800
      Left            =   45
      TabIndex        =   10
      Top             =   45
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   8467
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Find Window Columns"
      TabPicture(0)   =   "frmOrder.frx":0596
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraPageContainer(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Sort Order Columns"
      TabPicture(1)   =   "frmOrder.frx":05B2
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraPageContainer(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fraPageContainer 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   4300
         Index           =   1
         Left            =   150
         TabIndex        =   24
         Top             =   450
         Width           =   7500
         Begin VB.Frame fraSortColumns 
            Caption         =   "Sort Order Columns :"
            Height          =   3800
            Left            =   4500
            TabIndex        =   26
            Top             =   400
            Width           =   3000
            Begin ComctlLib.TreeView trvSelectedSortColumns 
               DragIcon        =   "frmOrder.frx":05CE
               Height          =   3375
               Left            =   165
               TabIndex        =   19
               Top             =   270
               Width           =   2700
               _ExtentX        =   4763
               _ExtentY        =   5953
               _Version        =   327682
               HideSelection   =   0   'False
               LabelEdit       =   1
               Style           =   1
               ImageList       =   "ImageList1"
               Appearance      =   1
            End
         End
         Begin VB.Frame fraSortColumnsSource 
            Caption         =   "Columns :"
            Height          =   3800
            Left            =   0
            TabIndex        =   25
            Top             =   400
            Width           =   3000
            Begin ComctlLib.TreeView trvSortColumns 
               DragIcon        =   "frmOrder.frx":0718
               Height          =   3375
               Left            =   150
               TabIndex        =   12
               Top             =   250
               Width           =   2700
               _ExtentX        =   4763
               _ExtentY        =   5953
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
            Left            =   3105
            TabIndex        =   18
            Top             =   3800
            Width           =   1335
         End
         Begin VB.TextBox txtOrderName 
            Height          =   315
            Index           =   1
            Left            =   645
            TabIndex        =   11
            Top             =   0
            Width           =   2355
         End
         Begin VB.CommandButton sscmdAddSortColumn 
            Caption         =   "&Add"
            Height          =   405
            Left            =   3105
            TabIndex        =   13
            Top             =   495
            Width           =   1335
         End
         Begin VB.CommandButton sscmdInsertSortColumn 
            Caption         =   "&Insert"
            Height          =   405
            Left            =   3105
            TabIndex        =   14
            Top             =   1005
            Width           =   1335
         End
         Begin VB.CommandButton sscmdRemoveSortColumn 
            Caption         =   "&Remove"
            Height          =   405
            Left            =   3105
            TabIndex        =   15
            Top             =   1500
            Width           =   1335
         End
         Begin VB.CommandButton sscmdMoveUpSortColumn 
            Caption         =   "Move &Up"
            Height          =   405
            Left            =   3105
            TabIndex        =   16
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton sscmdMoveDownSortColumn 
            Caption         =   "Move &Down"
            Height          =   405
            Left            =   3105
            TabIndex        =   17
            Top             =   2895
            Width           =   1335
         End
         Begin VB.Label lblOrderName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   27
            Top             =   60
            Width           =   510
         End
      End
      Begin VB.Frame fraPageContainer 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   4300
         Index           =   0
         Left            =   -74850
         TabIndex        =   20
         Top             =   450
         Width           =   7500
         Begin VB.Frame fraFindColumnsSource 
            Caption         =   "Columns :"
            Height          =   3800
            Left            =   0
            TabIndex        =   22
            Top             =   400
            Width           =   3000
            Begin ComctlLib.TreeView trvFindColumns 
               DragIcon        =   "frmOrder.frx":0B5A
               Height          =   3375
               Left            =   150
               TabIndex        =   1
               Top             =   250
               Width           =   2700
               _ExtentX        =   4763
               _ExtentY        =   5953
               _Version        =   327682
               HideSelection   =   0   'False
               LabelEdit       =   1
               Style           =   7
               ImageList       =   "imglstTreeviewImages"
               Appearance      =   1
            End
         End
         Begin VB.Frame fraFindColumns 
            Caption         =   "Find Window Columns :"
            Height          =   3800
            Left            =   4500
            TabIndex        =   21
            Top             =   400
            Width           =   3000
            Begin ComctlLib.TreeView trvSelectedFindColumns 
               DragIcon        =   "frmOrder.frx":0F9C
               Height          =   3375
               Left            =   150
               TabIndex        =   7
               Top             =   255
               Width           =   2700
               _ExtentX        =   4763
               _ExtentY        =   5953
               _Version        =   327682
               HideSelection   =   0   'False
               LabelEdit       =   1
               ImageList       =   "imglstTreeviewImages"
               Appearance      =   1
            End
         End
         Begin VB.TextBox txtOrderName 
            Height          =   315
            Index           =   0
            Left            =   645
            TabIndex        =   0
            Top             =   0
            Width           =   2355
         End
         Begin VB.CommandButton sscmdRemoveFindColumn 
            Caption         =   "&Remove"
            Height          =   405
            Left            =   3105
            TabIndex        =   4
            Top             =   1500
            Width           =   1335
         End
         Begin VB.CommandButton sscmdAddFindColumn 
            Caption         =   "&Add"
            Height          =   405
            Left            =   3105
            TabIndex        =   2
            Top             =   495
            Width           =   1335
         End
         Begin VB.CommandButton sscmdInsertFindColumn 
            Caption         =   "&Insert"
            Height          =   405
            Left            =   3105
            TabIndex        =   3
            Top             =   1005
            Width           =   1335
         End
         Begin VB.CommandButton sscmdMoveUpFindColumn 
            Caption         =   "Move &Up"
            Height          =   405
            Left            =   3105
            TabIndex        =   5
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton sscmdMoveDownFindColumn 
            Caption         =   "Move &Down"
            Height          =   405
            Left            =   3105
            TabIndex        =   6
            Top             =   2895
            Width           =   1335
         End
         Begin VB.Label lblOrderName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   23
            Top             =   60
            Width           =   510
         End
      End
   End
   Begin ComctlLib.ImageList imglstTreeviewImages 
      Left            =   3645
      Top             =   4875
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmOrder.frx":10E6
            Key             =   "IMG_COLUMN"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmOrder.frx":1680
            Key             =   "IMG_TABLE"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmOrder.frx":1BD2
            Key             =   "IMG_CALC"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3015
      Top             =   4875
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
            Picture         =   "frmOrder.frx":2124
            Key             =   "IMG_DOWN"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmOrder.frx":24EC
            Key             =   "IMG_UP"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Order definition variables.
Private mobjOrder As clsOrder

' Form handling variables.
Private mfCancelled As Boolean
Private mfColumnDrag As Boolean
Private mfReadOnly As Boolean
Private mfChanged As Boolean
Private mavColumns() As Variant

' Form handling constants.
Const PAGE_FINDCOLUMNS = 0
Const PAGE_SORTCOLUMNS = 1

Public Property Get Cancelled() As Boolean
  ' Return the Cancelled property.
  Cancelled = mfCancelled
  
End Property


Public Property Get Order() As clsOrder
  ' Return the Order object.
  Set Order = mobjOrder
  
End Property

Public Property Set Order(pobjOrder As clsOrder)
  ' Set the Order object.
  Dim sIcon As String
  Dim objOrdItem As clsOrderItem
  Dim objNewNode As ComctlLib.Node
  Dim iSequence As Integer

  ' Set the Order object global variable.
  Set mobjOrder = pobjOrder

  ' Display the order name.
  txtOrderName(PAGE_FINDCOLUMNS).Text = Trim(mobjOrder.OrderName)
  txtOrderName(PAGE_SORTCOLUMNS).Text = Trim(mobjOrder.OrderName)

  ' Populate the Find and Sort Columns treeviews.
  PopulateTreeViews

  ' Add items to the order and find columns listviews as required.
  For iSequence = 1 To mobjOrder.OrderItems.Count
    For Each objOrdItem In mobjOrder.OrderItems
      If objOrdItem.Sequence = iSequence Then
        If objOrdItem.ItemType = "O" Then
          sIcon = IIf(objOrdItem.Ascending, "IMG_UP", "IMG_DOWN")
          Set objNewNode = trvSelectedSortColumns.Nodes.Add(, , , objOrdItem.FullColumnName, sIcon)
          objNewNode.Tag = objOrdItem.ColumnID
          Set objNewNode = Nothing
        
          ' Remove the item from the Sort Coumns treeview.
          trvSortColumns.Nodes.Remove "C" & Trim(Str(objOrdItem.ColumnID))
        Else
          Set objNewNode = trvSelectedFindColumns.Nodes.Add(, , , objOrdItem.FullColumnName)
          objNewNode.Tag = objOrdItem.ColumnID
          Set objNewNode = Nothing
      
          ' Remove the item from the Find Coumns treeview.
          trvFindColumns.Nodes.Remove "C" & Trim(Str(objOrdItem.ColumnID))
        End If
      End If
    Next objOrdItem
    Set objOrdItem = Nothing
  Next iSequence
  
  
  ' Select the first item in each treeview.
  With trvSelectedFindColumns
    If .Nodes.Count > 0 Then
      .Nodes.item(1).Selected = True
      .SelectedItem.EnsureVisible
    End If
  End With
  
  With trvSelectedSortColumns
    If .Nodes.Count > 0 Then
      .Nodes.item(1).Selected = True
      .SelectedItem.EnsureVisible
    End If
  End With

  ' Disable controls if the user does not have permission to edit orders.
  mfReadOnly = Not datGeneral.SystemPermission("ORDERS", "EDIT")
  
  If pobjOrder.ContainsEditableColumns And mobjOrder.OrderID <> 0 Then
    mfReadOnly = True
    COAMsgBox "This order has been defined as a system order and cannot be edited.", vbInformation
  End If
  
  If mfReadOnly Then
    DisableAll
  End If
  
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






Private Sub DisableAll()
  ' Disable all controls.
  Dim ctlTemp As Control
  
  For Each ctlTemp In Me.Controls
    If (Not TypeOf ctlTemp Is Label) And _
      (Not TypeOf ctlTemp Is ImageList) And _
      (Not TypeOf ctlTemp Is Frame) And _
      (Not TypeOf ctlTemp Is TreeView) And _
      (Not TypeOf ctlTemp Is ListView) And _
      (Not TypeOf ctlTemp Is SSTab) Then
      
      ctlTemp.Enabled = False
    End If
  Next ctlTemp
  Set ctlTemp = Nothing
  
  cmdCancel.Enabled = True

End Sub


Private Sub PopulateTreeViews()
  ' Populate the Find Columns treeview.
  Dim iNextIndex As Integer
  Dim lngTableID As Long
  Dim sSQL As String
  Dim rsInfo As Recordset
  Dim objNode As ComctlLib.Node

  lngTableID = 0

  ' Clear the treeview controls.
  trvFindColumns.Nodes.Clear
  trvSortColumns.Nodes.Clear

  ' Construct an array of column info.
  'This is to avoid having to hit the server everytime we need info on a column.
  ' Column 1 = Column ID
  ' Column 2 = Column name
  ' Column 3 = Table ID
  ' Column 4 = table name
  ReDim mavColumns(4, 0)
  
  ' Get the list of columns for the order's base table.
  sSQL = "SELECT ASRSysColumns.tableID, ASRSysColumns.columnID, ASRSysColumns.columnName, ASRSysTables.tableName" & _
    " FROM ASRSysColumns" & _
    " INNER JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & _
    " WHERE ASRSysColumns.tableID = " & Trim(Str(mobjOrder.TableID)) & _
    " AND ASRSysColumns.ColumnType <> " & Trim(Str(colSystem)) & _
    " AND ASRSysColumns.ColumnType <> " & Trim(Str(colLink))
      Set rsInfo = datGeneral.GetRecords(sSQL)
  With rsInfo
    Do While Not .EOF
      ' Add the table root node if it hasn't already been added.
      If lngTableID <> !TableID Then
        lngTableID = !TableID

        Set objNode = trvFindColumns.Nodes.Add(, tvwChild, _
          "T" & Trim(Str(mobjOrder.TableID)), !TableName, "IMG_TABLE", "IMG_TABLE")
        objNode.Sorted = True
        objNode.Expanded = True
        Set objNode = Nothing
  
        Set objNode = trvSortColumns.Nodes.Add(, tvwChild, _
          "T" & Trim(Str(mobjOrder.TableID)), !TableName, "IMG_TABLE", "IMG_TABLE")
        objNode.Sorted = True
        objNode.Expanded = True
        Set objNode = Nothing
      End If
  
      ' Add items to the treeview for each column in the order's base table.
      Set objNode = trvFindColumns.Nodes.Add("T" & Trim(Str(!TableID)), _
        tvwChild, "C" & Trim(Str(!ColumnID)), _
        !ColumnName, "IMG_COLUMN", "IMG_COLUMN")
      objNode.Tag = !ColumnID
      Set objNode = Nothing
  
      Set objNode = trvSortColumns.Nodes.Add("T" & Trim(Str(!TableID)), _
        tvwChild, "C" & Trim(Str(!ColumnID)), _
        !ColumnName, "IMG_COLUMN", "IMG_COLUMN")
      objNode.Tag = !ColumnID
      Set objNode = Nothing
  
      ' Add the column to the srray of column info.
      iNextIndex = UBound(mavColumns, 2) + 1
      ReDim Preserve mavColumns(4, iNextIndex)
      mavColumns(1, iNextIndex) = !ColumnID
      mavColumns(2, iNextIndex) = !ColumnName
      mavColumns(3, iNextIndex) = mobjOrder.TableID
      mavColumns(4, iNextIndex) = !TableName
      
      .MoveNext
    Loop
    
    .Close
  End With
  Set rsInfo = Nothing

  ' Get the list of columns for the order's parent tables.
  sSQL = "SELECT ASRSysColumns.tableID, ASRSysColumns.columnID, ASRSysColumns.columnName, ASRSysTables.tableName" & _
    " FROM ASRSysColumns" & _
    " JOIN ASRSysTables ON ASRSysColumns.tableID = ASRSysTables.tableID" & _
    " JOIN ASRSysRelations ON ASRSysColumns.tableID = ASRSysRelations.parentID" & _
    " WHERE ASRSysRelations.childID = " & Trim(Str(mobjOrder.TableID)) & _
    " AND ASRSysColumns.ColumnType <> " & Trim(Str(colSystem)) & _
    " AND ASRSysColumns.ColumnType <> " & Trim(Str(colLink)) & _
    " AND ASRSysColumns.DataType <> " & Trim(Str(sqlOle)) & _
    " AND ASRSysColumns.DataType <> " & Trim(Str(sqlVarBinary)) & _
    " ORDER BY ASRSysTables.tableName, ASRSysColumns.columnName"
  Set rsInfo = datGeneral.GetRecords(sSQL)
  With rsInfo
    Do While Not .EOF
      ' Add the table root node if it hasn't already been added.
      If lngTableID <> !TableID Then
        lngTableID = !TableID
  
        Set objNode = trvFindColumns.Nodes.Add(, tvwChild, _
          "T" & Trim(Str(!TableID)), !TableName, "IMG_TABLE", "IMG_TABLE")
        objNode.Sorted = True
        objNode.Expanded = False
        Set objNode = Nothing
  
        Set objNode = trvSortColumns.Nodes.Add(, tvwChild, _
          "T" & Trim(Str(!TableID)), !TableName, "IMG_TABLE", "IMG_TABLE")
        objNode.Sorted = True
        objNode.Expanded = False
        Set objNode = Nothing
      End If
  
      ' Add items to the treeview for each column in the order's base table.
      Set objNode = trvFindColumns.Nodes.Add("T" & Trim(Str(!TableID)), _
        tvwChild, "C" & Trim(Str(!ColumnID)), _
        !ColumnName, "IMG_COLUMN", "IMG_COLUMN")
      objNode.Tag = !ColumnID
      Set objNode = Nothing
  
      Set objNode = trvSortColumns.Nodes.Add("T" & Trim(Str(!TableID)), _
        tvwChild, "C" & Trim(Str(!ColumnID)), _
        rsInfo!ColumnName, "IMG_COLUMN", "IMG_COLUMN")
      objNode.Tag = !ColumnID
      Set objNode = Nothing
  
      ' Add the column to the srray of column info.
      iNextIndex = UBound(mavColumns, 2) + 1
      ReDim Preserve mavColumns(4, iNextIndex)
      mavColumns(1, iNextIndex) = !ColumnID
      mavColumns(2, iNextIndex) = !ColumnName
      mavColumns(3, iNextIndex) = !TableID
      mavColumns(4, iNextIndex) = !TableName
      
      rsInfo.MoveNext
    Loop

    .Close
  End With
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
      .SelectedItem.EnsureVisible
    End If
  End With

End Sub


Private Sub SortColumns_RefreshControls()
  ' Refesh the controls whose status is variable.
  Dim fSelectedFirstNode As Boolean
  Dim fSelectedLastNode As Boolean
  Dim fSelectedColumnValid As Boolean
  Dim fSelectedSortColumnValid As Boolean
  
  ' Check that we have a valid column selected.
  fSelectedColumnValid = Not (trvSortColumns.SelectedItem Is Nothing)
  If fSelectedColumnValid Then
    fSelectedColumnValid = (Left(trvSortColumns.SelectedItem.Key, 1) = "C")
  End If
  
  sscmdAddSortColumn.Enabled = fSelectedColumnValid And _
    (Not mfReadOnly)
  sscmdInsertSortColumn.Enabled = fSelectedColumnValid And _
    (Not mfReadOnly)
  
  ' Check if we have a 'selected column' selected.
  fSelectedFirstNode = False
  fSelectedLastNode = False
  fSelectedSortColumnValid = Not (trvSelectedSortColumns.SelectedItem Is Nothing)
  If fSelectedSortColumnValid Then
    fSelectedFirstNode = (trvSelectedSortColumns.SelectedItem.Tag = trvSelectedSortColumns.SelectedItem.FirstSibling.Tag)
    fSelectedLastNode = (trvSelectedSortColumns.SelectedItem.Tag = trvSelectedSortColumns.SelectedItem.LastSibling.Tag)
  End If
  
  cmdSortColumnAscDesc.Enabled = fSelectedSortColumnValid And _
    (Not mfReadOnly)
  sscmdRemoveSortColumn.Enabled = fSelectedSortColumnValid And _
    (Not mfReadOnly)
  sscmdMoveUpSortColumn.Enabled = fSelectedSortColumnValid And _
    (trvSelectedSortColumns.Nodes.Count > 1) And _
    (Not fSelectedFirstNode) And _
    (Not mfReadOnly)
  sscmdMoveDownSortColumn.Enabled = fSelectedSortColumnValid And _
    (trvSelectedSortColumns.Nodes.Count > 1) And _
    (Not fSelectedLastNode) And _
    (Not mfReadOnly)
    
  ' Disable the OK command control if there are no order items specified.
  cmdOK.Enabled = (trvSelectedFindColumns.Nodes.Count > 0) And _
    (trvSelectedSortColumns.Nodes.Count > 0) And _
    (Len(Trim(txtOrderName(0).Text)) > 0) And _
    (Not mfReadOnly)

End Sub



Private Sub cmdCancel_Click()

  Dim intAnswer As Integer
  
  ' Check if any changes have been made.
  If mfChanged Then
    intAnswer = COAMsgBox("The order definition has changed.  Save changes ?", vbQuestion + vbYesNoCancel + vbDefaultButton1, app.ProductName)
    If intAnswer = vbYes Then
      If Me.cmdOK.Enabled Then
        Call cmdOK_Click
        Exit Sub
      Else
        If (Len(Me.txtOrderName(0).Text) = 0) Then
          COAMsgBox "Invalid Order Name", vbExclamation + vbOKOnly, app.title
        Else
          COAMsgBox "You must define both a find window order and a sort order" & vbCrLf & _
                 "for this table.", vbExclamation + vbOKOnly, app.title
        End If
        Exit Sub
      End If
    ElseIf intAnswer = vbCancel Then
      Exit Sub
    End If
  End If
  
  ' Set the Cancelled property and unload the form.
  mfCancelled = True
  Unload Me

End Sub






Private Sub FindColumns_RefreshControls()
  ' Refesh the controls whose status is variable.
  Dim fSelectedFirstNode As Boolean
  Dim fSelectedLastNode As Boolean
  Dim fSelectedColumnValid As Boolean
  Dim fSelectedFindColumnValid As Boolean
  
  ' Check that we have a valid column selected.
  fSelectedColumnValid = Not (trvFindColumns.SelectedItem Is Nothing)
  If fSelectedColumnValid Then
    fSelectedColumnValid = (Left(trvFindColumns.SelectedItem.Key, 1) = "C")
  End If
  
  sscmdAddFindColumn.Enabled = (fSelectedColumnValid And Not mfReadOnly)
  sscmdInsertFindColumn.Enabled = (fSelectedColumnValid And Not mfReadOnly)
  
  ' Check if we have a 'selected column' selected.
  fSelectedFirstNode = False
  fSelectedLastNode = False
  fSelectedFindColumnValid = Not (trvSelectedFindColumns.SelectedItem Is Nothing)
  If fSelectedFindColumnValid Then
    fSelectedFirstNode = (trvSelectedFindColumns.SelectedItem.Tag = trvSelectedFindColumns.SelectedItem.FirstSibling.Tag)
    fSelectedLastNode = (trvSelectedFindColumns.SelectedItem.Tag = trvSelectedFindColumns.SelectedItem.LastSibling.Tag)
  End If
  
  sscmdRemoveFindColumn.Enabled = (fSelectedFindColumnValid And Not mfReadOnly)
  sscmdMoveUpFindColumn.Enabled = fSelectedFindColumnValid And _
    (trvSelectedFindColumns.Nodes.Count > 1) And _
    (Not fSelectedFirstNode) And _
    (Not mfReadOnly)
  sscmdMoveDownFindColumn.Enabled = fSelectedFindColumnValid And _
    (trvSelectedFindColumns.Nodes.Count > 1) And _
    (Not fSelectedLastNode) And _
    (Not mfReadOnly)
    
  ' Disable the OK command control if there are no order items specified.
  cmdOK.Enabled = (trvSelectedFindColumns.Nodes.Count > 0) And _
    (trvSelectedSortColumns.Nodes.Count > 0) And _
    (Len(Trim(txtOrderName(0).Text)) > 0) And _
    (Not mfReadOnly)

End Sub

Private Sub cmdOK_Click()
  ' Confirm the order.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fDeleted As Boolean
  Dim fTimeStampChanged As Boolean
  Dim fContinueSave As Boolean
  Dim fSaveAsNew As Boolean
  Dim iLoop As Integer
  Dim iSequence As Integer
  Dim sSQL As String
  Dim sMBText As String
  Dim sTableName As String
  Dim objNode As ComctlLib.Node
  Dim rsOrders As Recordset
  
  fOK = True
  
  If mfChanged Then
    ' Reset the Cancelled property.
    mfCancelled = False
    
    ' Validate the order name.
    fOK = Len(Trim(txtOrderName(0).Text)) > 0
    If Not fOK Then
      COAMsgBox "Invalid order name.", vbOKOnly + vbExclamation, app.ProductName
    Else
      ' Check that the order has not been modified by someone else.
      If mobjOrder.OrderID > 0 Then
        fSaveAsNew = False
        fContinueSave = True
    
        sSQL = "SELECT convert(int, timestamp) AS timestamp" & _
          " FROM ASRSysOrders" & _
          " WHERE orderID = " & Trim(Str(mobjOrder.OrderID))
        Set rsOrders = datGeneral.GetRecords(sSQL)
    
        fDeleted = (rsOrders.BOF And rsOrders.EOF)
    
        If fDeleted Then
          fTimeStampChanged = True
        Else
          fTimeStampChanged = (mobjOrder.Timestamp <> rsOrders!Timestamp)
        End If
    
        rsOrders.Close
        Set rsOrders = Nothing
    
        If fTimeStampChanged Then
          If fDeleted Then
            ' Unable to overwrite existing definition
            sMBText = "This order has been deleted by another user." & vbCrLf & _
              "Save as a new definition ?"
          
            Select Case COAMsgBox(sMBText, vbExclamation + vbOKCancel, app.ProductName)
              Case vbOK         'save as new (but this may cause duplicate name message)
                fContinueSave = True
                fSaveAsNew = True
              Case vbCancel     'Do not save
                fContinueSave = False
            End Select
          Else
            ' Prompt to see if user should overwrite definition
            sMBText = "This order has been amended by another user. " & vbCrLf & _
              "Would you like to overwrite this definition?" & vbCrLf
            Select Case COAMsgBox(sMBText, vbExclamation + vbYesNoCancel, app.ProductName)
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
          fOK = False
        ElseIf fSaveAsNew Then
          mobjOrder.OrderID = 0
        End If
      End If
    End If
    
    If fOK Then
      ' Check that there are no other orders for this table with this name.
      sSQL = "SELECT *" & _
        " FROM ASRSysOrders" & _
        " WHERE orderID <> " & Trim(Str(mobjOrder.OrderID)) & _
        " AND tableID = " & Trim(Str(mobjOrder.TableID)) & _
        " AND name = '" & Replace(Trim(txtOrderName(0).Text), "'", "''") & "'" & _
        " AND type = " & Trim(Str(mobjOrder.OrderType))
      Set rsOrders = datGeneral.GetRecords(sSQL)
      With rsOrders
        fOK = .EOF And .BOF
        
        If Not fOK Then
          COAMsgBox "An order named '" & Trim(txtOrderName(0).Text) & "' already exists !", vbOKOnly + vbExclamation, app.ProductName
        End If
      
        .Close
      End With
      Set rsOrders = Nothing
    End If
  
    If fOK Then
      ' Write the changes to the order object.
      mobjOrder.OrderName = Trim(txtOrderName(0).Text)
      fOK = mobjOrder.ClearOrderItems
      
      If fOK Then
        If trvSelectedSortColumns.Nodes.Count > 0 Then
          Set objNode = trvSelectedSortColumns.Nodes.item(1).FirstSibling
          iSequence = 0
          Do While Not objNode Is Nothing
            iSequence = iSequence + 1
            sTableName = ""
            For iLoop = 1 To UBound(mavColumns, 2)
              If mavColumns(1, iLoop) = objNode.Tag Then
                sTableName = mavColumns(4, iLoop)
                Exit For
              End If
            Next iLoop

            mobjOrder.AddOrderItem objNode.Tag, "O", iSequence, (objNode.Image = "IMG_UP"), objNode.Text, sTableName, False
            Set objNode = objNode.Next
          Loop
        End If
        
        If trvSelectedFindColumns.Nodes.Count > 0 Then
          Set objNode = trvSelectedFindColumns.Nodes.item(1).FirstSibling
          iSequence = 0
          Do While Not objNode Is Nothing
            iSequence = iSequence + 1
            sTableName = ""
            For iLoop = 1 To UBound(mavColumns, 2)
              If mavColumns(1, iLoop) = objNode.Tag Then
                sTableName = mavColumns(4, iLoop)
                Exit For
              End If
            Next iLoop
            
            mobjOrder.AddOrderItem objNode.Tag, "F", iSequence, True, objNode.Text, sTableName, False
            Set objNode = objNode.Next
          Loop
        End If
      End If
    End If
  Else
    mfCancelled = True
  End If
  
TidyUpAndExit:
  ' Disassociate object variables.
  If fOK Then
    Unload Me
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

  trvSelectedFindColumns.DragIcon = picDocument.Picture
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
        intAnswer = COAMsgBox("You have changed the current definition. Save changes ?", vbQuestion + vbYesNoCancel + vbDefaultButton1, app.ProductName)
        If intAnswer = vbYes Then
          If Me.cmdOK.Enabled Then
            Call cmdOK_Click
            If mfCancelled = True Then Cancel = 1
          Else
            If (Len(Me.txtOrderName(0).Text) = 0) Then
              COAMsgBox "Invalid Order Name", vbExclamation + vbOKOnly, app.title
            Else
              COAMsgBox "You must define both a find window order and a sort order" & vbCrLf & _
                   "for this table.", vbExclamation + vbOKOnly, app.title
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


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

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
  If Left(nodSelection.Key, 1) <> "C" Then
    Exit Sub
  End If
  
  ' Do nothing if the selected column is already in the list of Find columns.
  Set objNewNode = trvSelectedFindColumns.Nodes.Add(, , , nodSelection.Parent.Text & "." & nodSelection.Text)
  objNewNode.Tag = nodSelection.Tag
  objNewNode.Selected = True
  objNewNode.EnsureVisible
  Set objNewNode = Nothing
    
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
  If Left(nodSelection.Key, 1) <> "C" Then
    Exit Sub
  End If
    
  ' Add the selected column to the Sort Order columns listview.
  Set objNewNode = trvSelectedSortColumns.Nodes.Add(, , , nodSelection.Parent.Text & "." & nodSelection.Text, "IMG_UP")
  objNewNode.Tag = nodSelection.Tag
  objNewNode.Selected = True
  objNewNode.EnsureVisible
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
  If Left(nodSelection.Key, 1) <> "C" Then
    Exit Sub
  End If
  
  ' Deselect the current listview selection.
  ' Add the selected column to the Find Columns listview.
  Set objHighlightNode = trvSelectedFindColumns.SelectedItem
  If Not objHighlightNode Is Nothing Then
    objHighlightNode.Selected = False
    Set objNewNode = trvSelectedFindColumns.Nodes.Add(objHighlightNode.Index, tvwPrevious, , nodSelection.Parent.Text & "." & nodSelection.Text)
  Else
    Set objNewNode = trvSelectedFindColumns.Nodes.Add(, , , nodSelection.Parent.Text & "." & nodSelection.Text)
  End If
  objNewNode.Tag = nodSelection.Tag
  objNewNode.Selected = True
  objNewNode.EnsureVisible
  
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
  If Left(nodSelection.Key, 1) <> "C" Then
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
  objNewNode.EnsureVisible
  
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
  ' Move the selected listview item DOWN one position.
  Dim objNode As ComctlLib.Node
  Dim objNewNode As ComctlLib.Node
  
  Set objNode = trvSelectedFindColumns.SelectedItem
  If Not objNode Is Nothing Then
    ' Move the selected item down one position in the treeview.
    If objNode.Tag <> objNode.LastSibling.Tag Then
      Set objNewNode = trvSelectedFindColumns.Nodes.Add(objNode.Next, tvwNext, , objNode.Text)
      objNewNode.Tag = objNode.Tag
      objNewNode.Selected = True
      objNewNode.EnsureVisible
      Set objNewNode = Nothing

      trvSelectedFindColumns.Nodes.Remove objNode.Index
      
      mfChanged = True
    End If
  End If
    
  ' Disassociate object variables.
  Set objNode = Nothing

  FindColumns_RefreshControls

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
      objNewNode.EnsureVisible
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
  ' Move the selected treeview item DOWN one position.
  Dim objNode As ComctlLib.Node
  Dim objNewNode As ComctlLib.Node
  
  Set objNode = trvSelectedFindColumns.SelectedItem
  If Not objNode Is Nothing Then
    ' Move the selected item up one position in the treeview.
    If objNode.Tag <> objNode.FirstSibling.Tag Then
      Set objNewNode = trvSelectedFindColumns.Nodes.Add(objNode.Previous, tvwPrevious, , objNode.Text)
      objNewNode.Tag = objNode.Tag
      objNewNode.Selected = True
      objNewNode.EnsureVisible
      Set objNewNode = Nothing

      trvSelectedFindColumns.Nodes.Remove objNode.Index
      
      mfChanged = True
    End If
  End If
    
  ' Disassociate object variables.
  Set objNode = Nothing

  FindColumns_RefreshControls
  
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
      objNewNode.EnsureVisible
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
  Dim iLoop As Integer
  Dim lngColumnID As Long
  Dim objNode As ComctlLib.Node

  With trvSelectedFindColumns
    If Not .SelectedItem Is Nothing Then
      lngColumnID = .SelectedItem.Tag

      .Nodes.Remove .SelectedItem.Index

      ' Get the columnName and tableID of the selected column.
      For iLoop = 1 To UBound(mavColumns, 2)
        If mavColumns(1, iLoop) = lngColumnID Then
          ' Add the column back into the treeview.
          Set objNode = trvFindColumns.Nodes.Add("T" & Trim(Str(mavColumns(3, iLoop))), _
            tvwChild, "C" & Trim(Str(lngColumnID)), mavColumns(2, iLoop), "IMG_COLUMN", "IMG_COLUMN")
          objNode.Tag = lngColumnID
          Set objNode = Nothing
        
          Exit For
        End If
      Next iLoop
    
      mfChanged = True
    End If

    If Not .SelectedItem Is Nothing Then
      .SelectedItem.Selected = True
    End If
  End With

  FindColumns_RefreshControls

End Sub


Private Sub sscmdRemoveSortColumn_Click()
  ' Remove the selected item from the Find Columns listview.
  Dim iLoop As Integer
  Dim lngColumnID As Long
  Dim objNode As ComctlLib.Node

  With trvSelectedSortColumns
    If Not .SelectedItem Is Nothing Then
      lngColumnID = .SelectedItem.Tag

      .Nodes.Remove .SelectedItem.Index

      ' Get the columnName and tableID of the selected column.
      For iLoop = 1 To UBound(mavColumns, 2)
        If mavColumns(1, iLoop) = lngColumnID Then
          ' Add the column back into the treeview.
          Set objNode = trvSortColumns.Nodes.Add("T" & Trim(Str(mavColumns(3, iLoop))), _
            tvwChild, "C" & Trim(Str(lngColumnID)), mavColumns(2, iLoop), "IMG_COLUMN", "IMG_COLUMN")
          objNode.Tag = lngColumnID
          Set objNode = Nothing
        
          Exit For
        End If
      Next iLoop
    
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
  FindColumns_RefreshControls

End Sub


Private Sub trvFindColumns_DblClick()
  If Not mfReadOnly Then
    sscmdAddFindColumn_Click
  End If
  
End Sub


Private Sub trvFindColumns_DragDrop(Source As Control, x As Single, y As Single)
  ' Remove the selected item from the columns listview.
  If Source Is trvSelectedFindColumns Then
    sscmdRemoveFindColumn_Click
  End If

  FindColumns_RefreshControls

End Sub


Private Sub trvFindColumns_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Start the drag-drop operation.
  Dim fGoodColumn As Boolean
  Dim nodSelection As Node
  
  If Not mfReadOnly Then
    If Button = vbLeftButton Then
      'Get the item at the mouse position
      Set nodSelection = trvFindColumns.HitTest(x, y)
      If Not nodSelection Is Nothing Then
        'If this item is not the selected item, make it
        If Not nodSelection Is trvFindColumns.SelectedItem Then
          Set trvFindColumns.SelectedItem = nodSelection
        End If
      End If
    End If
    
    ' Do not drag anything if there is no selected item in the treeview.
    fGoodColumn = Not (trvFindColumns.SelectedItem Is Nothing)
    
    ' Do not drag the selected item if is not a column.
    If fGoodColumn Then
      fGoodColumn = (Left(trvFindColumns.SelectedItem.Key, 1) = "C")
    End If
    
    If fGoodColumn Then
      ' Set the flag to show that a column is being dragged.
      mfColumnDrag = True
      trvFindColumns.Drag vbBeginDrag
    End If
    
    FindColumns_RefreshControls
  End If
  
End Sub


Private Sub trvFindColumns_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If mfColumnDrag Then
    ' Reset the flag that shows that a column is being dragged.
    trvFindColumns.Drag vbCancel
    mfColumnDrag = False
  End If

  FindColumns_RefreshControls

End Sub


Private Sub trvFindColumns_NodeClick(ByVal Node As ComctlLib.Node)
  FindColumns_RefreshControls

End Sub


Private Sub trvSelectedFindColumns_DblClick()
  If Not mfReadOnly Then
    sscmdRemoveFindColumn_Click
  End If
  
End Sub


Private Sub trvSelectedFindColumns_DragDrop(Source As Control, x As Single, y As Single)
  ' Drop a selected item from the columns listbox into the listview.
  Dim fDropOk As Boolean
  Dim objHighlightNode As ComctlLib.Node
  Dim objNewNode As ComctlLib.Node
  Dim objOldNode As ComctlLib.Node

  Set objHighlightNode = trvSelectedFindColumns.DropHighlight

  If Source Is trvFindColumns Then
    fDropOk = Not (trvFindColumns.SelectedItem Is Nothing)

    If fDropOk Then
      If Not trvSelectedFindColumns.SelectedItem Is Nothing Then
        trvSelectedFindColumns.SelectedItem.Selected = False
      End If

      If objHighlightNode Is Nothing Then
        Set objNewNode = trvSelectedFindColumns.Nodes.Add(, , , trvFindColumns.SelectedItem.Parent.Text & "." & trvFindColumns.SelectedItem.Text)
      Else
        Set objNewNode = trvSelectedFindColumns.Nodes.Add(objHighlightNode, tvwPrevious, , trvFindColumns.SelectedItem.Parent.Text & "." & trvFindColumns.SelectedItem.Text)
      End If
      
      objNewNode.Tag = trvFindColumns.SelectedItem.Tag
      objNewNode.Selected = True
      objNewNode.EnsureVisible

      ' Remove the column from the treeview.
      trvFindColumns.Nodes.Remove trvFindColumns.SelectedItem.Index


      trvFindColumns.Drag vbEndDrag
      Set objNewNode = Nothing
      trvSelectedFindColumns.SetFocus
      mfChanged = True
    Else
      trvFindColumns.Drag vbCancel
    End If
  ElseIf Source Is trvSelectedFindColumns Then
    If (objHighlightNode Is Nothing) Or _
      (trvSelectedFindColumns.SelectedItem Is objHighlightNode) Then

      trvSelectedFindColumns.Drag vbCancel
    Else
      Set objOldNode = trvSelectedFindColumns.SelectedItem
      Set objNewNode = trvSelectedFindColumns.Nodes.Add(objHighlightNode, tvwPrevious, , objOldNode.Text)
      objNewNode.Tag = objOldNode.Tag
      objNewNode.Selected = True
      objNewNode.EnsureVisible

      trvSelectedFindColumns.Nodes.Remove objOldNode.Index
      
      Set objNewNode = Nothing
      Set objOldNode = Nothing
      mfChanged = True
    End If
  End If

  Set objHighlightNode = Nothing
  Set trvSelectedFindColumns.DropHighlight = Nothing

  FindColumns_RefreshControls

End Sub


Private Sub trvSelectedFindColumns_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  Dim objNode As ComctlLib.Node

  'Get the item at the mouse's coordinates.
  Set objNode = trvSelectedFindColumns.HitTest(x, y)

  ' Check if the item at the mouse's coordinates is a control.
  If Not objNode Is Nothing Then
    objNode.EnsureVisible
  End If

  ' Set the DropHighlight node
  Set trvSelectedFindColumns.DropHighlight = objNode

  Set objNode = Nothing

End Sub

Private Sub trvSelectedFindColumns_KeyUp(KeyCode As Integer, Shift As Integer)
  If Not mfReadOnly Then
    If KeyCode = vbKeyDelete Then
      sscmdRemoveFindColumn_Click
    End If
  End If

End Sub

Private Sub trvSelectedFindColumns_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim objNode As ComctlLib.Node

  If Not mfReadOnly Then
    If Button = vbLeftButton Then
      ' Get the item at the mouse position
      Set objNode = trvSelectedFindColumns.HitTest(x, y)
      If Not objNode Is Nothing Then
        ' If this node is not the selected node, make it
        If Not objNode Is trvSelectedFindColumns.SelectedItem Then
          trvSelectedFindColumns.SelectedItem.Selected = False
          Set trvSelectedFindColumns.SelectedItem = objNode
        End If
      End If
  
      If trvSelectedFindColumns.Nodes.Count > 0 Then
        mfColumnDrag = True
        trvSelectedFindColumns.Drag vbBeginDrag
      End If
    End If
    
    FindColumns_RefreshControls
  End If
  
End Sub

Private Sub trvSelectedFindColumns_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  If mfColumnDrag Then
    ' Reset the flag that shows that a column is being dragged.
    trvSelectedFindColumns.Drag vbCancel
    mfColumnDrag = False
  End If

  FindColumns_RefreshControls

End Sub

Private Sub trvSelectedFindColumns_NodeClick(ByVal Node As ComctlLib.Node)
  Node.Selected = True
  FindColumns_RefreshControls

End Sub

Private Sub trvSelectedSortColumns_DblClick()
  If Not mfReadOnly Then
    cmdSortColumnAscDesc_Click
  End If

End Sub


Private Sub trvSelectedSortColumns_DragDrop(Source As Control, x As Single, y As Single)
  ' Drop a selected item from the columns listbox into the listview.
  Dim fDropOk As Boolean
  Dim objHighlightNode As ComctlLib.Node
  Dim objNewNode As ComctlLib.Node
  Dim objOldNode As ComctlLib.Node

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
      objNewNode.EnsureVisible
      
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

  'Get the item at the mouse's coordinates.
  Set objNode = trvSelectedSortColumns.HitTest(x, y)

  ' Check if the item at the mouse's coordinates is a control.
  If Not objNode Is Nothing Then
    objNode.EnsureVisible
  End If

  ' Set the DropHighlight node
  Set trvSelectedSortColumns.DropHighlight = objNode

  Set objNode = Nothing

End Sub

Private Sub trvSelectedSortColumns_KeyUp(KeyCode As Integer, Shift As Integer)
  If Not mfReadOnly Then
    If KeyCode = vbKeyDelete Then
      sscmdRemoveSortColumn_Click
    End If
  End If

End Sub

Private Sub trvSelectedSortColumns_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim objNode As ComctlLib.Node

  If Not mfReadOnly Then
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
  End If

End Sub

Private Sub trvSelectedSortColumns_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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
  SortColumns_RefreshControls

End Sub


Private Sub trvSortColumns_DblClick()
  If Not mfReadOnly Then
    sscmdAddSortColumn_Click
  End If
  
End Sub


Private Sub trvSortColumns_DragDrop(Source As Control, x As Single, y As Single)
  ' Remove the selected item from the columns listview.
  If Source Is trvSelectedSortColumns Then
    sscmdRemoveSortColumn_Click
  End If

  SortColumns_RefreshControls

End Sub


Private Sub trvSortColumns_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Start the drag-drop operation.
  Dim fGoodColumn As Boolean
  Dim nodSelection As Node
  
  If Not mfReadOnly Then
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
      fGoodColumn = (Left(trvSortColumns.SelectedItem.Key, 1) = "C")
    End If
    
    If fGoodColumn Then
      ' Set the flag to show that a column is being dragged.
      mfColumnDrag = True
      trvSortColumns.Drag vbBeginDrag
    End If
    
    SortColumns_RefreshControls
  End If
  
End Sub


Private Sub trvSortColumns_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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
  
  For iLoop = txtOrderName.LBound To txtOrderName.UBound
    If iLoop <> Index Then
      txtOrderName(iLoop).Text = txtOrderName(Index).Text
    End If
  Next iLoop

  mfChanged = True
  
  ' Disable the OK command control if there are no order items specified.
  cmdOK.Enabled = (trvSelectedFindColumns.Nodes.Count > 0) And _
    (trvSelectedSortColumns.Nodes.Count > 0) And _
    (Len(Trim(txtOrderName(0).Text)) > 0) And _
    (Not mfReadOnly)

End Sub

Private Sub txtOrderName_GotFocus(Index As Integer)
  UI.txtSelText
  
End Sub


Private Sub txtOrderName_KeyPress(Index As Integer, KeyAscii As Integer)
  ' Validate the character entered.
  KeyAscii = ValidNameChar(KeyAscii, txtOrderName(Index).SelStart)

End Sub



