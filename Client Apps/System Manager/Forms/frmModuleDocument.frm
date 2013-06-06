VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmModuleDocument 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Document Management"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmModuleDocument.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   3885
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   6853
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Types"
      TabPicture(0)   =   "frmModuleDocument.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraComponent(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraTypes"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Mail Merge"
      TabPicture(1)   =   "frmModuleDocument.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblTransferTable"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblColumnName"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "grdTransferDetails(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdDelete"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdEdit"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cboTables"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cboCategory"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.ComboBox cboCategory 
         Height          =   315
         Left            =   -74055
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   495
         Width           =   3255
      End
      Begin VB.ComboBox cboTables 
         Height          =   315
         Left            =   -74040
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   900
         Width           =   3255
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit..."
         Enabled         =   0   'False
         Height          =   400
         Left            =   -68115
         TabIndex        =   17
         Top             =   1665
         Width           =   1200
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Cle&ar"
         Enabled         =   0   'False
         Height          =   400
         Left            =   -68115
         TabIndex        =   18
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Frame fraTypes 
         Caption         =   "Types : "
         Height          =   1755
         Left            =   135
         TabIndex        =   5
         Top             =   1935
         Width           =   5865
         Begin VB.ComboBox cboTypeTable 
            Height          =   315
            Left            =   2430
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   315
            Width           =   3255
         End
         Begin VB.ComboBox cboTypeColumn 
            Height          =   315
            Left            =   2430
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Tag             =   "cboTypeTable"
            Top             =   1215
            Width           =   3255
         End
         Begin VB.ComboBox cboTypeCategoryColumn 
            Height          =   315
            Left            =   2430
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Tag             =   "cboTypeTable"
            Top             =   765
            Width           =   3255
         End
         Begin VB.Label lblTypeTable 
            Caption         =   "Type Table : "
            Height          =   285
            Left            =   195
            TabIndex        =   6
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblTypeColumn 
            Caption         =   "Type Column : "
            Height          =   285
            Left            =   195
            TabIndex        =   10
            Top             =   1260
            Width           =   1410
         End
         Begin VB.Label lblTypeCategoryColumn 
            Caption         =   "Type Category Column : "
            Height          =   285
            Left            =   195
            TabIndex        =   8
            Top             =   810
            Width           =   2130
         End
      End
      Begin VB.Frame fraComponent 
         Caption         =   "Categories :"
         Height          =   1305
         Index           =   1
         Left            =   135
         TabIndex        =   12
         Tag             =   "6"
         Top             =   450
         Width           =   5865
         Begin VB.ComboBox cboCategoryTable 
            Height          =   315
            Left            =   2430
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   360
            Width           =   3255
         End
         Begin VB.ComboBox cboCategoryColumn 
            Height          =   315
            Left            =   2430
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Tag             =   "cboCategoryTable"
            Top             =   750
            Width           =   3255
         End
         Begin VB.Label lblCategoryTable 
            Caption         =   "Category Table : "
            Height          =   285
            Left            =   225
            TabIndex        =   1
            Top             =   405
            Width           =   1680
         End
         Begin VB.Label lblCatgeoryColumn 
            Caption         =   "Category Column : "
            Height          =   330
            Left            =   225
            TabIndex        =   3
            Top             =   795
            Width           =   1815
         End
      End
      Begin SSDataWidgets_B.SSDBGrid grdTransferDetails 
         Height          =   3255
         Index           =   0
         Left            =   -74865
         TabIndex        =   16
         Top             =   1665
         Width           =   6510
         ScrollBars      =   2
         _Version        =   196617
         DataMode        =   2
         RecordSelectors =   0   'False
         GroupHeaders    =   0   'False
         Col.Count       =   4
         stylesets.count =   2
         stylesets(0).Name=   "KeyField"
         stylesets(0).BackColor=   14024703
         stylesets(0).HasFont=   -1  'True
         BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(0).Picture=   "frmModuleDocument.frx":0044
         stylesets(1).Name=   "Mandatory"
         stylesets(1).BackColor=   15400959
         stylesets(1).HasFont=   -1  'True
         BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(1).Picture=   "frmModuleDocument.frx":0060
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   3
         SelectByCell    =   -1  'True
         BalloonHelp     =   0   'False
         MaxSelectedRows =   0
         ForeColorEven   =   -2147483640
         ForeColorOdd    =   -2147483640
         BackColorEven   =   -2147483643
         BackColorOdd    =   -2147483643
         RowHeight       =   423
         ExtraHeight     =   79
         Columns.Count   =   4
         Columns(0).Width=   3200
         Columns(0).Visible=   0   'False
         Columns(0).Caption=   "CategoryID"
         Columns(0).Name =   "CategoryID"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   5292
         Columns(1).Caption=   "Heading"
         Columns(1).Name =   "Heading"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3200
         Columns(2).Visible=   0   'False
         Columns(2).Caption=   "ColumnID"
         Columns(2).Name =   "ColumnID"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   5636
         Columns(3).Caption=   "Value"
         Columns(3).Name =   "Value"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   11492
         _ExtentY        =   5741
         _StockProps     =   79
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblColumnName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         Height          =   195
         Left            =   -74775
         TabIndex        =   15
         Top             =   555
         Width           =   645
      End
      Begin VB.Label lblTransferTable 
         Caption         =   "Table : "
         Height          =   285
         Left            =   -74775
         TabIndex        =   13
         Top             =   945
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   5070
      TabIndex        =   20
      Top             =   4050
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   405
      Left            =   3780
      TabIndex        =   19
      Top             =   4050
      Width           =   1200
   End
End
Attribute VB_Name = "frmModuleDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnReadOnly As Boolean
Private mbChanged As Boolean

Private mavarCategoryTableIDs() As Variant


Private Sub cboCategory_Click()

  Dim iCount As Integer
  Dim iIndex As Integer
  'Set the base table
  SetComboItem cboTables, CLng(mavarCategoryTableIDs(2, cboCategory.ListIndex))
  
  For iCount = grdTransferDetails.LBound To grdTransferDetails.UBound
    grdTransferDetails(iCount).Visible = (cboCategory.ListIndex = iCount)
    grdTransferDetails.Item(iCount).SelBookmarks.RemoveAll
    GoTopOfGrid CLng(iCount), (cboTables = "<None>")
  Next iCount
  
  mbChanged = True
  RefreshButtons

End Sub

Private Sub cboCategoryTable_Click()

  Dim objctl As Control

  ' Clear the current contents of the combos.
  For Each objctl In Me
  
    If (TypeOf objctl Is ComboBox) And _
      (objctl.Tag = "cboCategoryTable") Then
      
        PopulateComboWithColumns objctl, GetComboItem(cboCategoryTable), True
      
    End If
  Next objctl

  mbChanged = True

  RefreshButtons

End Sub

Private Sub cboTables_Click()

  Dim lngIndex As Long

  If SelectedComboItem(cboTables) <> mavarCategoryTableIDs(2, cboCategory.ListIndex) Then
    
    If mavarCategoryTableIDs(2, cboCategory.ListIndex) > 0 Then
    
      If MsgBox("Changing the base table will reset all the columns for this document type." & vbCrLf _
        & "Are you sure you want to continue?", vbYesNo + vbQuestion, "Document Management Setup") = vbYes Then
        
        PopulateTransferDetails cboCategory.ListIndex, True
        mavarCategoryTableIDs(2, cboCategory.ListIndex) = SelectedComboItem(cboTables)
      Else
        lngIndex = mavarCategoryTableIDs(2, cboCategory.ListIndex)
        SetComboItem cboTables, lngIndex
      End If
    Else
      mavarCategoryTableIDs(2, cboCategory.ListIndex) = SelectedComboItem(cboTables)
      GoTopOfGrid 0, (cboTables = "<None>")
      cmdEdit.Enabled = (cboTables = "<None>")
    End If
    
    mbChanged = True
    
  End If
  
  RefreshButtons

End Sub

Private Sub cboTypeTable_Click()

  Dim objctl As Control

  ' Clear the current contents of the combos.
  For Each objctl In Me
  
    If (TypeOf objctl Is ComboBox) And _
      (objctl.Tag = "cboTypeTable") Then
      
        PopulateComboWithColumns objctl, GetComboItem(cboTypeTable), True
      
    End If
  Next objctl

  mbChanged = True

  RefreshButtons

End Sub

Private Sub cboTypeColumn_Click()
  mbChanged = True
  RefreshButtons
End Sub

Private Sub cboTypeCategoryColumn_Click()
  mbChanged = True
  RefreshButtons
End Sub

Private Sub cmdCancel_Click()
  UnLoad Me
End Sub

Private Sub cmdOK_Click()

  If SaveChanges Then
    mbChanged = False
    UnLoad Me
  End If

End Sub


Private Function SaveChanges() As Boolean

  Dim iLoop As Integer
  Dim sSQL As String
  Dim iLoopTypes As Integer
  Dim iCategory As Integer

  Screen.MousePointer = vbHourglass

  ' Category info
  SaveModuleSetting gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_CATEGORYTABLE, gsPARAMETERTYPE_TABLEID, GetComboItem(cboCategoryTable)
  SaveModuleSetting gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_CATEGORYCOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboCategoryColumn)

  ' Types info
  SaveModuleSetting gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_TYPETABLE, gsPARAMETERTYPE_TABLEID, GetComboItem(cboTypeTable)
  SaveModuleSetting gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_TYPECOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboTypeColumn)
  SaveModuleSetting gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_TYPECATEGORYCOLUMN, gsPARAMETERTYPE_COLUMNID, GetComboItem(cboTypeCategoryColumn)


  ' Mail Merge Stuff
'  daoDb.Execute "DELETE FROM tmpDocumentManagementCategories", dbFailOnError
'
'  For iLoop = LBound(mavarCategoryTableIDs, 2) To UBound(mavarCategoryTableIDs, 2) - 1
'    sSQL = "INSERT INTO tmpDocumentManagementCategories" & _
'      " (CategoryID, Category, [TableID])" & _
'      " VALUES (" & _
'      CStr(mavarCategoryTableIDs(0, iLoop)) & "," & _
'      "'" & CStr(mavarCategoryTableIDs(1, iLoop)) & "', " & _
'      mavarCategoryTableIDs(2, iLoop) & ")"
'
'    daoDb.Execute sSQL, dbFailOnError
'  Next iLoop


  ' Store the transfer details
'  daoDb.Execute "DELETE FROM tmpDocumentManagementHeaderInfo WHERE Type = 1", dbFailOnError
'  For iLoopTypes = 0 To cboCategory.ListCount - 1
'    With grdTransferDetails(iLoopTypes)
'      .Redraw = False
'      .MoveFirst
'
'      iCategory = cboCategory.ItemData(iLoopTypes)
'
'      For iLoop = 0 To (.Rows - 1)
'
'        sSQL = "INSERT INTO tmpDocumentManagementHeaderInfo" & _
'          " (CategoryID, Heading, ColumnID, Type)" & _
'          " VALUES (" & _
'          iCategory & "," & _
'          "'" & .Columns("Heading").value & "'," & _
'          .Columns("ColumnID").value & "," & _
'          "1)"
'
'        daoDb.Execute sSQL, dbFailOnError
'        .MoveNext
'      Next iLoop
'    End With
'  Next iLoopTypes



  Screen.MousePointer = vbNormal
  Application.Changed = True
  SaveChanges = True

End Function


Private Sub PopulateComboWithColumns(ByRef cboTemp As ComboBox, ByVal plngTableID As Long, ByVal AllowNone As Boolean)

  If AllowNone Then
    With cboTemp
      .Clear
      .AddItem "<None>"
      .ItemData(.NewIndex) = 0
    End With
  End If


  With recColEdit
    .Index = "idxName"
    .Seek ">=", plngTableID

    If Not .NoMatch Then
      Do While Not .EOF
        If !TableID <> plngTableID Then
          Exit Do
        End If

        If (Not !Deleted) And (!DataType = dtVARCHAR) Then

          cboTemp.AddItem (!ColumnName)
          cboTemp.ItemData(cboTemp.NewIndex) = !ColumnID

        End If

        .MoveNext
      Loop
    End If
  End With

End Sub

' Initialise the Base Table combo(s)
Private Sub InitialiseCombos()
   
  With cboCategoryTable
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
   
   
  With cboTypeTable
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
    
    
  ' Add items to the combo for each table that has not been deleted,
  ' and is a Lookup table.
  With recTabEdit
    .Index = "idxName"
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      If Not !Deleted And (!TableType = giTABLELOOKUP) Then
                
        cboTypeTable.AddItem !TableName
        cboTypeTable.ItemData(cboTypeTable.NewIndex) = !TableID
                
        cboCategoryTable.AddItem !TableName
        cboCategoryTable.ItemData(cboTypeTable.NewIndex) = !TableID
                
      End If
      .MoveNext
    Loop
  End With
    

End Sub

Private Sub RefreshButtons()
  cmdOK.Enabled = mbChanged
End Sub

Private Sub RetrieveDefinition()

  SetComboItem cboCategoryTable, GetModuleSetting(gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_CATEGORYTABLE, 0)
  SetComboItem cboCategoryColumn, GetModuleSetting(gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_CATEGORYCOLUMN, 0)

  SetComboItem cboTypeTable, GetModuleSetting(gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_TYPETABLE, 0)
  SetComboItem cboTypeCategoryColumn, GetModuleSetting(gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_TYPECATEGORYCOLUMN, 0)
  SetComboItem cboTypeColumn, GetModuleSetting(gsMODULEKEY_DOCMANAGEMENT, gsPARAMETERKEY_DOCMAN_TYPECOLUMN, 0)

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

  ReDim mavarCategoryTableIDs(2, 0)

  SSTab1.TabVisible(0) = True
  SSTab1.TabVisible(1) = False

  mblnReadOnly = (Application.AccessMode <> accFull And _
                 Application.AccessMode <> accSupportMode)

  If mblnReadOnly Then
    ControlsDisableAll Me
  End If
  
  InitialiseCombos
  PopulateParentsCombo cboTables
  'PopulateCategories

  ' Load the transfer types
  For iLoop = 0 To cboCategory.ListCount - 1
    PopulateTransferDetails iLoop, False
  Next iLoop

  RetrieveDefinition

  mbChanged = False

  ' Get rid of that pesky icon
  RemoveIcon Me

End Sub


Private Sub PopulateParentsCombo(ByRef objCombo As ComboBox)
  ' Clear the contents of the combo.
  objCombo.Clear
  
  With recTabEdit
    .Index = "idxName"
    
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      If !TableType = iTabParent And Not !Deleted Then
        objCombo.AddItem !TableName
        objCombo.ItemData(objCombo.NewIndex) = !TableID
      End If
      
      .MoveNext
    Loop
  End With
  
End Sub

Private Sub PopulateCategories()

  Dim sSQL As String
  Dim rsData As DAO.Recordset


  sSQL = "SELECT CategoryID, Category, TableID FROM tmpDocumentManagementCategories" _
      & " ORDER BY CategoryID"
  Set rsData = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  With rsData
    While Not .EOF

      mavarCategoryTableIDs(0, UBound(mavarCategoryTableIDs, 2)) = Trim(!CategoryID)
      mavarCategoryTableIDs(1, UBound(mavarCategoryTableIDs, 2)) = !Category
      mavarCategoryTableIDs(2, UBound(mavarCategoryTableIDs, 2)) = !TableID
      ReDim Preserve mavarCategoryTableIDs(2, UBound(mavarCategoryTableIDs, 2) + 1)

      cboCategory.AddItem Trim(!Category)
      cboCategory.ItemData(cboCategory.NewIndex) = !CategoryID
      
      .MoveNext
    Wend
   
    .Close
  End With
       
  Set rsData = Nothing

  ' Set to the top
  cboCategory.ListIndex = 0

End Sub


Private Sub PopulateTransferDetails(ByVal plngTransferGrid As Long, pbReset As Boolean)

  Dim sSQL As String
  Dim strAddString As String
  Dim strMapToDescription As String
  Dim rsDefinition As DAO.Recordset
  Dim ctlGrid As SSDBGrid
  Dim iCategoryID As Integer

  iCategoryID = cboCategory.ItemData(plngTransferGrid)

  ' Unload grid if resetting
  If pbReset Then
    grdTransferDetails(plngTransferGrid).RemoveAll
  Else

    ' Load up a grid for this definition
    If plngTransferGrid > 0 Then
      Load grdTransferDetails(plngTransferGrid)
      grdTransferDetails(plngTransferGrid).RemoveAll
    End If
  End If

  sSQL = "SELECT *" & _
    " FROM tmpDocumentManagementHeaderInfo" & _
    " WHERE [CategoryID] = " & CStr(iCategoryID) & " AND [Type] = 1" & _
    " ORDER BY CategoryID;"
    
  Set rsDefinition = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  While Not rsDefinition.EOF
    
    If Not pbReset Then
      If IsNull(rsDefinition!ColumnID) Then
        strMapToDescription = ""
      Else
        strMapToDescription = GetColumnName(rsDefinition!ColumnID, False)
      End If
    Else
      strMapToDescription = ""
    End If
    
    strAddString = rsDefinition!CategoryID & vbTab & rsDefinition!Heading & vbTab & rsDefinition!ColumnID & vbTab & strMapToDescription
                        
    grdTransferDetails(plngTransferGrid).AddItem strAddString
    rsDefinition.MoveNext
    
  Wend
  GoTopOfGrid plngTransferGrid, (cboCategory.Text = "<None>")
  cmdEdit.Enabled = (cboCategory.Text = "<None>")
  
  rsDefinition.Close
  Set rsDefinition = Nothing

End Sub

Private Sub GoTopOfGrid(lngIndex As Long, fClearBookmarks As Boolean)

  If fClearBookmarks Then
    grdTransferDetails.Item(lngIndex).SelBookmarks.RemoveAll
  Else
    grdTransferDetails.Item(lngIndex).SelBookmarks.Add grdTransferDetails.Item(lngIndex).Bookmark
    grdTransferDetails.Item(lngIndex).Bookmark = grdTransferDetails.Item(lngIndex).SelBookmarks(lngIndex)
  End If

End Sub

Private Function SelectedComboItem(cboTemp As ComboBox) As Long
  With cboTemp
    If .ListIndex >= 0 Then
      SelectedComboItem = .ItemData(.ListIndex)
    Else
      SelectedComboItem = 0
    End If
  End With
End Function

