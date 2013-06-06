VERSION 5.00
Begin VB.Form frmRelate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relationships"
   ClientHeight    =   4770
   ClientLeft      =   1215
   ClientTop       =   1575
   ClientWidth     =   4305
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5025
   Icon            =   "frmRelate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   1710
      TabIndex        =   3
      Top             =   4215
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   2985
      TabIndex        =   4
      Top             =   4215
      Width           =   1200
   End
   Begin VB.Frame fraChildren 
      Caption         =   "Child Tables :"
      Height          =   3500
      Left            =   150
      TabIndex        =   1
      Top             =   600
      Width           =   4035
      Begin VB.ListBox lstRelations 
         Height          =   2985
         Index           =   1
         Left            =   200
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   300
         Width           =   3675
      End
   End
   Begin VB.ComboBox cboParents 
      Height          =   315
      Left            =   1530
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   200
      Width           =   2685
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parent Table :"
      Height          =   195
      Left            =   195
      TabIndex        =   5
      Top             =   255
      Width           =   1245
   End
End
Attribute VB_Name = "frmRelate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Local variables to hold property values
Private objParentTable As HRProSystemMgr.Table

'Local variables
Private gfNew As Boolean
Private gfInitialising As Boolean
Private mfLoading As Boolean

Private mblnReadOnly As Boolean

Private Function ParentCount(pLngChildTableID As Long, plngExcludedTableID As Long, ByRef pLngParentTableID As Long) As Integer
  ' Return the number of parents the given table has (excluding the given table).
  ' Return the ID of the parent table in the plngParenTableID parameter.
  ' NB. This is only useful if the
  Dim iLoop As Integer
  Dim iListPtr  As Integer
  Dim iParentCount As Integer
  
  iParentCount = 0
  pLngParentTableID = 0
  
  ' Loop through the parent tables.
  For iLoop = 1 To lstRelations.Count
    ' Do not bother with the excluded table.
    If val(lstRelations(iLoop).Tag) <> plngExcludedTableID Then
      ' Loop through the possible children of the parent table.
      ' ie. loop through the items in the listview associated with the parent table.
      For iListPtr = 0 To (lstRelations(iLoop).ListCount - 1)
        ' Check if the listview node is for the given child table.
        If (pLngChildTableID = lstRelations(iLoop).ItemData(iListPtr)) Then
          ' Check if the node is selected (ie. the table is selected as a child of the parent).
          If lstRelations(iLoop).Selected(iListPtr) Then
            iParentCount = iParentCount + 1
            pLngParentTableID = val(lstRelations(iLoop).Tag)
          End If
              
          Exit For
        End If
      Next iListPtr
    End If
  Next iLoop
  
  ParentCount = iParentCount
  
End Function

Public Property Get ParentTable() As HRProSystemMgr.Table
  
  ' Return the current parent table.
  If objParentTable Is Nothing Then
    Set objParentTable = New HRProSystemMgr.Table
  End If
  
  Set ParentTable = objParentTable
  
End Property

Public Property Set ParentTable(TableObject As HRProSystemMgr.Table)

  ' Set the parent table.
  If TableObject Is Nothing Then
    gfNew = gfInitialising
  Else
  
    Set objParentTable = TableObject
    
    If gfInitialising Then
      gfNew = (objParentTable.TableID < 1)
    End If
    
  End If
  
  gfInitialising = False
  
End Property

Private Sub cboParents_Click()
  Dim iLoop As Integer
  
  ' Update the parent table property.
  ParentTable.TableID = cboParents.ItemData(cboParents.ListIndex)
  ParentTable.ReadTable
  
  ' Only display the required listbox.
  For iLoop = 1 To lstRelations.UBound
    If lstRelations(iLoop).Tag = ParentTable.TableID Then
      lstRelations(iLoop).Visible = True
    Else
      lstRelations(iLoop).Visible = False
    End If
  Next iLoop
  
End Sub

Private Sub cmdCancel_Click()
  UnLoad Me
  
End Sub


Private Sub cmdOK_Click()
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fRelExists As Boolean
  Dim fChangesMade As Boolean
  Dim fRelRequired As Boolean
  Dim iLoop As Integer
  Dim iListPtr As Integer
  Dim objRelatedTable As New HRProSystemMgr.Table
  
  ' Start transaction for temporary database.
  daoWS.BeginTrans
  
  fOK = True
  fChangesMade = False
  
  For iLoop = 1 To cboParents.ListCount
    
    ' Update the parent table property.
    ParentTable.TableID = cboParents.ItemData(iLoop - 1)
    
    fOK = ParentTable.ReadTable
    
    If fOK Then
      ' Update relation definitions for the parent table.
      For iListPtr = 0 To (lstRelations(iLoop).ListCount - 1)
      
        objRelatedTable.TableID = lstRelations(iLoop).ItemData(iListPtr)
        fRelRequired = lstRelations(iLoop).Selected(iListPtr)
        
        With recRelEdit
          .Index = "idxParentID"
          .Seek "=", ParentTable.TableID, objRelatedTable.TableID
          fRelExists = (Not .NoMatch)
        End With
          
        ' Is this a new relation?
        If fRelRequired And Not fRelExists Then
        
          fOK = objRelatedTable.ReadTable
          
          If fOK Then
            'Add new relation definition
            With recRelEdit
              .AddNew
              !parentID = ParentTable.TableID
              !childID = objRelatedTable.TableID
              .Update
            End With
                        
            ' Add foreign key column definition for related table
            With objRelatedTable
              fOK = .AddIDColumn(ParentTable.TableID)
              
              If fOK Then
                fOK = .WriteTable
                fChangesMade = fOK
              End If
              
              'TM20010724 Fault 1239
              If fOK Then
                'Set the !Changed field property to True for all Default expressions
                'in the child table. So that SaveChanges refreshes the 'DfltExpr' SP.
                fOK = .DefaultExprChange(.TableID)
              End If
            End With
            
          End If
        
        ElseIf fRelExists And Not fRelRequired Then
          'Delete old relation definition
          fOK = ParentTable.DeleteRelation(objRelatedTable.TableID)
          fChangesMade = fOK
        End If
        
        If Not fOK Then
          Exit For
        End If
      Next iListPtr
    
    End If
  
    If Not fOK Then
      Exit For
    End If
  Next iLoop
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set objRelatedTable = Nothing
  
  If fOK And fChangesMade Then
    'Commit changes to temporary database
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
    Application.ChangedDiaryLink = True
  Else
    'Rollback changes to temporary database
    daoWS.Rollback
  End If
  
  UnLoad Me
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub

Private Sub Form_Activate()
  'Dim iLoop As Integer
  
  ' If we are not editing the properties of a specific table
  ' then select the top item.
  If (ParentTable.TableID < 1 Or gfNew) And _
    cboParents.ListCount > 0 Then
    ParentTable.TableID = cboParents.ItemData(0)
    ParentTable.ReadTable
  End If

  cboParents.Text = ParentTable.TableName
  
End Sub



Private Function IsIndirectlyRelated(pLngChildID As Long, plngParentID As Long) As Boolean
  ' Returns TRUE if the given child table is indirectly to the given parent table.
  Dim fFound As Boolean
  Dim iLoop1 As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim iNextIndex As Integer
  Dim alngParentsParents() As Long
  Dim alngChildsParents() As Long
  Dim alngLowestChildren() As Long
  Dim alngHighestParents() As Long

  '
  ' Construct an array the given parent table's top-level parents.
  '
  ReDim alngParentsParents(0)
  
  ' Call the function to return an array of the given parent tables top-level parents.
  alngParentsParents = GetHighestParents(plngParentID, pLngChildID, plngParentID)
  ' If the given parent table has no top-level parents then it must be one itself.
  ' So add it to the array.
  If UBound(alngParentsParents) = 0 Then
    ReDim Preserve alngParentsParents(1)
    alngParentsParents(1) = plngParentID
  End If
      
  '
  ' Construct an array the given child table's top-level parents.
  '
  ReDim alngChildsParents(0)
  
  ' Call the function to return an array of the given child tables bottom-level children.
  ReDim alngLowestChildren(0)
  alngLowestChildren = GetLowestChildren(pLngChildID, pLngChildID, plngParentID)
  ' If the given child table has no bottom-level children then it must be one itself.
  ' So add it to the array.
  If UBound(alngLowestChildren) = 0 Then
    ReDim Preserve alngLowestChildren(1)
    alngLowestChildren(1) = pLngChildID
  End If
  
  ' Call the function to return an array of each bottom-level child table's top-level parents.
  For iLoop1 = 1 To UBound(alngLowestChildren)
    ReDim alngHighestParents(0)
    alngHighestParents = GetHighestParents(alngLowestChildren(iLoop1), pLngChildID, plngParentID)
    ' If the bottom-level child table has no top-level parents then it must be one itself.
    ' So add it to the array.
    If UBound(alngHighestParents) = 0 Then
      ReDim Preserve alngHighestParents(1)
      alngHighestParents(1) = alngLowestChildren(iLoop1)
    End If
    
    ' Add the bottom-level child table's top-level parents to the array of the
    ' given child table's top-level parents if it is not already in the array.
    For iLoop2 = 1 To UBound(alngHighestParents)
      fFound = False
      For iLoop3 = 1 To UBound(alngChildsParents)
        If alngChildsParents(iLoop3) = alngHighestParents(iLoop2) Then
          fFound = True
          Exit For
        End If
      Next iLoop3
      
      If Not fFound Then
        iNextIndex = UBound(alngChildsParents) + 1
        ReDim Preserve alngChildsParents(iNextIndex)
        alngChildsParents(iNextIndex) = alngHighestParents(iLoop2)
      End If
    Next iLoop2
  Next iLoop1
  
  ' Check if any of the given parent table's top-level parents match
  ' with the given child table's top-level parents.
  ' If any do then the two given tables are related.
  For iLoop1 = 1 To UBound(alngParentsParents)
    For iLoop2 = 1 To UBound(alngChildsParents)
      If alngParentsParents(iLoop1) = alngChildsParents(iLoop2) Then
        IsIndirectlyRelated = True
        Exit Function
      End If
    Next iLoop2
  Next iLoop1
  
  IsIndirectlyRelated = False

End Function



Private Sub Form_Initialize()

  gfInitialising = True
  
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
  Screen.MousePointer = vbHourglass
 
  ' Clear the menu shortcuts. This needs to be done so that some shortcut keys
  ' (eg. DEL) will function normally in textboxes instead of triggering menu options.
  frmSysMgr.ClearMenuShortcuts
  
  ' Populate the combo with a list of parent and child databases.
  PopulateParentsCombo
  
  '  Populate the listboxes with the possible relations for each table.
  mfLoading = True
  PopulateListBoxes
  mfLoading = False
  
  Screen.MousePointer = vbNormal
  
End Sub


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Disassociate object variables.
  Set objParentTable = Nothing

End Sub


Private Sub PopulateParentsCombo()
  ' Clear the contents of the combo.
  cboParents.Clear
  
  With recTabEdit
    .Index = "idxName"
    
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    
    Do While Not .EOF
      If !TableType <> iTabLookup And Not !Deleted Then
        cboParents.AddItem !TableName
        cboParents.ItemData(cboParents.NewIndex) = !TableID
      End If
      
      .MoveNext
    Loop
  End With
  
End Sub

Private Sub PopulateListBoxes()
  Dim lngTableID As Long
  Dim fRelExists As Boolean
  Dim iLoop As Integer
  Dim objTable As Table
  
  ' For each table in the parentcombo ...
  For iLoop = 1 To cboParents.ListCount
  
    ' Create a new listbox if required.
    If iLoop > lstRelations.UBound Then
      Load lstRelations(iLoop)
    End If
    lstRelations(iLoop).Tag = cboParents.ItemData(iLoop - 1)
 
    ' Clear the listview of all items.
    lstRelations(iLoop).Clear
    lstRelations(iLoop).Visible = False
  
    ' Set the parent table property to be the one selected in the combo.
    lngTableID = cboParents.ItemData(iLoop - 1)
    Set objTable = New Table
    objTable.TableID = lngTableID
    
    With recTabEdit
      .Index = "idxName"
      
      If Not (.BOF And .EOF) Then
        .MoveFirst
      End If
      
      ' Add all child tables to the listbox.
      ' NB. Do not add deleted tables, and do not add the listbox's base table.
      Do While Not .EOF()
        If (Not !Deleted) And _
          (!TableID <> lngTableID) And _
          (!TableType = iTabChild) Then
            
          recRelEdit.Index = "idxParentID"
          recRelEdit.Seek "=", lngTableID, !TableID
          fRelExists = (Not recRelEdit.NoMatch)
              
          ' Add the table to the listview
          lstRelations(iLoop).AddItem (!TableName)
          lstRelations(iLoop).ItemData(lstRelations(iLoop).NewIndex) = !TableID
          lstRelations(iLoop).Selected(lstRelations(iLoop).NewIndex) = fRelExists
        End If
        
        .MoveNext
      Loop
    End With
    
    If lstRelations(iLoop).ListCount > 0 Then
      lstRelations(iLoop).ListIndex = 0
    End If
    
    Set objTable = Nothing
  Next iLoop

End Sub













Private Sub lstRelations_ItemCheck(Index As Integer, Item As Integer)
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim iListPtr As Integer
  'Dim iListPtr2 As Integer
  Dim iParentCount As Integer
  Dim iGrandParentCount As Integer
  Dim iChildParentCount As Integer
  Dim lngSelectedTableID As Long
  Dim lngParentTableID As Long
  Dim lngTemp As Long
  Dim lngChildTableID As Long
  Dim sSelectedTableName As String
  Dim sParentTableName As String

  Const MAXPARENTS = 2
  
  
  'Ignore any clicks...
  If mblnReadOnly Then
    If Not mfLoading Then   'Use loading flag to prevent out of stack space error due to recursion...
      With lstRelations(Index)
        mfLoading = True
        .Selected(.ListIndex) = Not .Selected(.ListIndex)
        mfLoading = False
      End With
    End If
    Exit Sub
  End If
  
  
  fOK = True
  
  If lstRelations(Index).Selected(lstRelations(Index).ListIndex) And _
    Not mfLoading Then
    
    ' Get the ID and name of the table just selected.
    lngSelectedTableID = lstRelations(Index).ItemData(lstRelations(Index).ListIndex)
    sSelectedTableName = lstRelations(Index).List(lstRelations(Index).ListIndex)
    sParentTableName = cboParents.List(cboParents.ListIndex)
    
    ' Get the count of the selected table's parent.
    iParentCount = ParentCount(lngSelectedTableID, val(lstRelations(Index).Tag), lngParentTableID)

    ' Check if the selected table already has 2 parents.
    fOK = (iParentCount < MAXPARENTS)
    
    If Not fOK Then
      ' Max. parent count has already been met so tell the user what's happended,
      ' and reset the selection just made.
      MsgBox "The '" & sSelectedTableName & "' table already has the maximum number of parent tables.", _
        vbOKOnly + vbExclamation, Application.Name
      
      lstRelations(Index).Selected(lstRelations(Index).ListIndex) = False
    End If
    
    ' Check if the selected table already has a parent table,
    ' and the parent table has more than one parent.
    If fOK And (iParentCount = 1) Then
      ' Get the parent's number of parents.
      iGrandParentCount = ParentCount(lngParentTableID, 0, lngTemp)

      fOK = (iGrandParentCount < MAXPARENTS)

      If Not fOK Then
        ' The selected table already has a parent table,
        ' and the parent table has more than one parent.
        ' Tell the user what's happended, and reset the selection just made.
        MsgBox "The '" & sSelectedTableName & "' table already has a parent table with more than one parent.", _
          vbOKOnly + vbExclamation, Application.Name

        lstRelations(Index).Selected(lstRelations(Index).ListIndex) = False
      End If
    End If

    If fOK And (iParentCount = 1) Then
      ' Check if the selected table already has a parent table,
      ' and the selected table already has a child,
      ' and the child table has more than one parent (including the selected table).
      For iLoop = 1 To lstRelations.Count
        If lstRelations(iLoop).Tag = lngSelectedTableID Then
          For iListPtr = 0 To (lstRelations(iLoop).ListCount - 1)
            If lstRelations(iLoop).Selected(iListPtr) Then
              
              lngChildTableID = lstRelations(iLoop).ItemData(iListPtr)
              iChildParentCount = ParentCount(lngChildTableID, 0, lngTemp)
              
              fOK = (iChildParentCount < MAXPARENTS)

              If Not fOK Then
                MsgBox "The '" & sSelectedTableName & "' table already has a parent table, and a child with the maximum number of parents.", _
                  vbOKOnly + vbExclamation, Application.Name
        
                lstRelations(Index).Selected(lstRelations(Index).ListIndex) = False
              End If
            End If
            
            If Not fOK Then
              Exit For
            End If
          Next iListPtr

          Exit For
        End If
      
        If Not fOK Then
          Exit For
        End If
      Next iLoop
    End If

    If fOK And (iParentCount = 1) Then
      ' Check if the selected table already has a parent table,
      ' and the new parent table has more than one parent.
    
      ' Get the new parent's number of parents.
      iGrandParentCount = ParentCount(val(lstRelations(Index).Tag), 0, lngTemp)
    
      fOK = (iGrandParentCount < MAXPARENTS)

      If Not fOK Then
        ' The selected table already has a parent table,
        ' and the new parent table has more than one parent.
        ' Tell the user what's happended, and reset the selection just made.
        MsgBox "The '" & sSelectedTableName & "' table already has a parent table, and the new parent table has more than one parent.", _
          vbOKOnly + vbExclamation, Application.Name

        lstRelations(Index).Selected(lstRelations(Index).ListIndex) = False
      End If
    End If
    
    If fOK Then
      fOK = Not IsDescendant(lstRelations(Index).Tag, lngSelectedTableID)
      If Not fOK Then
        MsgBox "The '" & sParentTableName & "' table is already a descendant of the '" & sSelectedTableName & "' table.", _
          vbOKOnly + vbExclamation, Application.Name
    
        lstRelations(Index).Selected(lstRelations(Index).ListIndex) = False
      End If
    End If
    
    If fOK Then
      fOK = Not IsIndirectlyRelated(lngSelectedTableID, lstRelations(Index).Tag)
      If Not fOK Then
        MsgBox "The '" & sSelectedTableName & "' table is already indirectly related to the '" & sParentTableName & "' table.", _
          vbOKOnly + vbExclamation, Application.Name
    
        lstRelations(Index).Selected(lstRelations(Index).ListIndex) = False
      End If
    End If
  End If
  
End Sub


Private Function GetHighestParents(plngTableID As Long, plngIgnoreChildID As Long, plngIgnoreParentID As Long) As Variant
  ' Return an array of the given tables top generation ascendants.
  ' Ignore the direct relationship between the given child and parent table.
  Dim iLoop1 As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim iNextIndex As Integer
  Dim alngParents() As Long
  Dim alngHighestParents() As Long
  
  ReDim alngHighestParents(0)
  
  ' Loop through the relationship listboxes.
  For iLoop1 = 1 To lstRelations.UBound
    ' If the listbox does NOT list the given table's children...
    If lstRelations(iLoop1).Tag <> plngTableID Then
      ' Loop through the items in the listbox.
      For iLoop2 = 0 To lstRelations(iLoop1).ListCount - 1
        ' If the given table is a child of the current listbox's table,
        ' And ignoring the given relationship...
        If (lstRelations(iLoop1).ItemData(iLoop2) = plngTableID) And _
          (lstRelations(iLoop1).Selected(iLoop2)) And _
          ((lstRelations(iLoop1).Tag <> plngIgnoreParentID) Or _
          (plngTableID <> plngIgnoreChildID)) Then

          ' Call the function to get the given parent table's top-level parents.
          ReDim alngParents(0)
          alngParents = GetHighestParents(lstRelations(iLoop1).Tag, plngIgnoreChildID, plngIgnoreParentID)
          
          ' If the given parent table has parents add them to the array of top-level parents.
          ' Else, the parent table is a top-level table, so add it to the array.
          If UBound(alngParents) > 0 Then
            For iLoop3 = 1 To UBound(alngParents)
              iNextIndex = UBound(alngHighestParents) + 1
              ReDim Preserve alngHighestParents(iNextIndex)
              alngHighestParents(iNextIndex) = alngParents(iLoop3)
            Next iLoop3
          Else
            iNextIndex = UBound(alngHighestParents) + 1
            ReDim Preserve alngHighestParents(iNextIndex)
            alngHighestParents(iNextIndex) = lstRelations(iLoop1).Tag
          End If
          
          Exit For
        End If
      Next iLoop2
    End If
  Next iLoop1

  GetHighestParents = alngHighestParents
  
End Function
Private Function IsDescendant(pLngChildID As Long, plngParentID As Long) As Boolean
  ' Returns TRUE if the given Child table is a descendant
  ' of the given Parent table id.
  ' NB. A version of this function exists in the Table class (IsDescendantOf)
  ' which runs off the recRelEdit table. This version of the function runs of
  ' the relationships that have been set up in the local ListBoxes.
  Dim iLoop1 As Integer
  Dim iLoop2 As Integer
  
  IsDescendant = False
  
  ' Analyse the listboxes for each table to determine if the given
  ' Child table is in any way a descendant of the given Parent table.
  For iLoop1 = 1 To lstRelations.UBound
    ' If they current listbox defines the given Parent table ...
    If lstRelations(iLoop1).Tag = plngParentID Then
      ' Check to see if the given child table is selected in the listbox.
      ' ie. is a child of the parent table.
      For iLoop2 = 0 To lstRelations(iLoop1).ListCount - 1
        If lstRelations(iLoop1).ItemData(iLoop2) = pLngChildID Then
          If lstRelations(iLoop1).Selected(iLoop2) Then
            IsDescendant = True
          End If
            
          Exit For
        End If
      Next iLoop2
    Else
      ' Check if the given parent table is the parent of any
      ' of the given child table's parents.
      For iLoop2 = 0 To lstRelations(iLoop1).ListCount - 1
        If (lstRelations(iLoop1).ItemData(iLoop2) = pLngChildID) Then
          If (lstRelations(iLoop1).Selected(iLoop2)) Then
            If IsDescendant(lstRelations(iLoop1).Tag, plngParentID) Then
              IsDescendant = True
            End If
          End If
            
          Exit For
        End If
      Next iLoop2
    End If
      
    If IsDescendant Then
      Exit For
    End If
  Next iLoop1
  
End Function



Private Function GetLowestChildren(plngTableID As Long, plngIgnoreChildID As Long, plngIgnoreParentID As Long) As Variant
  ' Return an array of the given tables bottom generation descendants.
  ' Ignore the direct relationship between the given child and parent table.
  Dim iLoop1 As Integer
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim iNextIndex As Integer
  Dim lngChildID As Long
  Dim alngChildren() As Long
  Dim alngLowestChildren() As Long
  
  ReDim alngLowestChildren(0)
  
  ' Loop through the relationship listboxes.
  For iLoop1 = 1 To lstRelations.UBound
    ' If the listbox lists the given table's children...
    If lstRelations(iLoop1).Tag = plngTableID Then
      ' Loop through the items in the listbox.
      For iLoop2 = 0 To lstRelations(iLoop1).ListCount - 1
        ' If the item is selected (ie. relationship exists...
        If lstRelations(iLoop1).Selected(iLoop2) Then
          ' Get the child table's ID.
          lngChildID = lstRelations(iLoop1).ItemData(iLoop2)

          ' Ignore the given relationship.
          If (lngChildID <> plngIgnoreChildID) Or _
            (plngTableID <> plngIgnoreParentID) Then
          
            ' Call the function to get the given child table's bottom-level children.
            ReDim alngChildren(0)
            alngChildren = GetLowestChildren(lngChildID, plngIgnoreChildID, plngIgnoreParentID)

            ' If the given child table has children add them to the array of bottom-level children.
            ' Else, the child table is a bottom-level table, so add it to the array.
            If UBound(alngChildren) > 0 Then
              For iLoop3 = 1 To UBound(alngChildren)
                iNextIndex = UBound(alngLowestChildren) + 1
                ReDim Preserve alngLowestChildren(iNextIndex)
                alngLowestChildren(iNextIndex) = alngChildren(iLoop3)
              Next iLoop3
            Else
              iNextIndex = UBound(alngLowestChildren) + 1
              ReDim Preserve alngLowestChildren(iNextIndex)
              alngLowestChildren(iNextIndex) = lngChildID
            End If
          End If
        End If
      Next iLoop2

      Exit For
    End If
  Next iLoop1

  GetLowestChildren = alngLowestChildren
  
End Function


Private Sub lstRelations_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  'If mblnReadOnly Then
    Index = 0
    Button = 0
    Shift = 0
  'End If
End Sub
