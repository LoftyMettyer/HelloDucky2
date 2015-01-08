VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
 Option Explicit

'Local variables to hold property values
Private gLngTableID As Long
Private gsTableName As String
Private giTableType As Integer
Private gDatUpdated As Date
Private gfNew As Boolean
Private gfChanged As Boolean
Private gfDeleted As Boolean
Private gLngPrimaryOrderID As Long
Private gLngRecordDescriptionID As Long
Private glngPrimaryEmailID As Long
Private gobjSummaryFields As Collection
Private mstrExprNotCopied As String
Private miGrantRead As Integer
Private miGrantEdit As Integer
Private miGrantNew As Integer
Private miGrantDelete As Integer
Private mbPermissionsPrompted As Boolean
Private mbIsCopy As Boolean
Private mlngCopyTableID As Long
Private msCopyTableName As String
Private mbManualColumnBreaks As Boolean
Private mbAuditInsert As Boolean
Private mbAuditDelete As Boolean
'Private mlngEmailInsert As Long
'Private mlngEmailDelete As Long
Private mfrmUse As frmUsage

Private mcolEmailLinks As Collection
Private mcolOutlookLinks As Collection
Private mcolWorkflowTriggeredLinks As Collection
Private mcolTableValidations As Collection
Private mcolTableTriggers As Collection
Private mbIsRemoteView As Boolean
Private mbLocked As Boolean

Private mbInsertTriggerDisabled As Boolean
Private mbUpdateTriggerDisabled As Boolean
Private mbDeleteTriggerDisabled As Boolean

Public Property Get InsertTriggerDisabled() As Boolean
  InsertTriggerDisabled = mbInsertTriggerDisabled
End Property

Public Property Let InsertTriggerDisabled(ByVal bNewValue As Boolean)
  mbInsertTriggerDisabled = bNewValue
End Property

Public Property Get UpdateTriggerDisabled() As Boolean
  UpdateTriggerDisabled = mbUpdateTriggerDisabled
End Property

Public Property Let UpdateTriggerDisabled(ByVal bNewValue As Boolean)
  mbUpdateTriggerDisabled = bNewValue
End Property

Public Property Get DeleteTriggerDisabled() As Boolean
  DeleteTriggerDisabled = mbDeleteTriggerDisabled
End Property

Public Property Let DeleteTriggerDisabled(ByVal bNewValue As Boolean)
  mbDeleteTriggerDisabled = bNewValue
End Property

Public Property Get PermissionsPrompted() As Boolean
    PermissionsPrompted = mbPermissionsPrompted
End Property

Public Property Let PermissionsPrompted(ByVal bNewValue As Boolean)
  mbPermissionsPrompted = bNewValue
End Property

Public Property Let GrantRead(ByVal iNewValue As Integer)
  miGrantRead = iNewValue
End Property

Public Property Let GrantEdit(ByVal iNewValue As Integer)
  miGrantEdit = iNewValue
End Property

Public Property Let GrantNew(ByVal iNewValue As Integer)
  miGrantNew = iNewValue
End Property

Public Property Let GrantDelete(ByVal iNewValue As Integer)
  miGrantDelete = iNewValue
End Property

Public Property Get ExpressionsNotCopied() As String
  ExpressionsNotCopied = mstrExprNotCopied
End Property

Public Property Get PrimaryEmailID() As Long
  PrimaryEmailID = glngPrimaryEmailID
End Property

Public Property Let PrimaryEmailID(ByVal lngNewValue As Long)
  glngPrimaryEmailID = lngNewValue
End Property


Public Property Get TableID() As Long
  
  ' Return the current table id.
  TableID = gLngTableID

End Property

Public Property Let TableID(plngNewTableID As Long)
  
  ' Set the table id.
  gLngTableID = plngNewTableID
  
End Property

Public Property Get TableName() As String
  ' Return the current table name.
  TableName = gsTableName
  
End Property

Public Property Let TableName(psNewName As String)
  
  ' Set the table name.
  gsTableName = psNewName
  
End Property

Public Property Get TableType() As enum_TableTypes

  ' Return the current table type.
  TableType = giTableType
  
End Property

Public Property Let TableType(piNewType As enum_TableTypes)

  ' Set the table type.
  giTableType = piNewType
  
End Property

Public Property Get LastUpdated() As Date

  ' Return the data the table was last updated.
  LastUpdated = gDatUpdated
  
End Property

Public Property Let LastUpdated(pDatUpdated As Date)

  ' Set the date the table was last updated.
  gDatUpdated = pDatUpdated
  
End Property

Public Property Get IsNew() As Boolean

  ' Return whether or not the table is new.
  IsNew = gfNew
  
End Property

Public Property Get IsChanged() As Boolean

  ' Return whether or not the table has been changed.
  IsChanged = gfChanged
  
End Property

Public Property Let IsChanged(pfSetting As Boolean)
  ' Set whether or not the table has been changed.
  gfChanged = pfSetting
  
End Property

Public Property Let IsDeleted(pfSetting As Boolean)
  ' Set whether or not the table has been deleted.
  gfDeleted = pfSetting
  
End Property
Public Property Let IsNew(pfSetting As Boolean)
  ' Set whether or not the table is new.
  gfNew = pfSetting
  
End Property


Public Property Get IsDeleted() As Boolean

  ' Return whether or not the table has been deleted.
  IsDeleted = gfDeleted
  
End Property

Public Function DeleteTable() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsColumns As DAO.Recordset
  Dim rsOrders As DAO.Recordset
  Dim rsExpressions As DAO.Recordset
  Dim rsViews As DAO.Recordset
  Dim rsRelatedTables As DAO.Recordset
  Dim objColumn As Column
  Dim objOrder As Order
  Dim objTable As Table
  Dim objExpression As CExpression
  
  Dim objEmail As clsEmailAddr
  Dim rsEmailAddrs As DAO.Recordset
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
 
  fOK = ReadTable

  'TM19012004
  ' Check that the table is not used.
  If fOK Then
    Set mfrmUse = New frmUsage
    mfrmUse.ResetList
    If TableIsUsed(mfrmUse) Then
      Screen.MousePointer = vbDefault
      Select Case Me.TableType
        Case enum_TableTypes.iTabParent
          mfrmUse.ShowMessage Me.TableName & " Table", "The table cannot be deleted as the table is used by the following:", UsageCheckObject.Table
        Case enum_TableTypes.iTabChild
          mfrmUse.ShowMessage Me.TableName & " Child Table", "The table cannot be deleted as the table is used by the following:", UsageCheckObject.ChildTable
        Case enum_TableTypes.iTabLookup
          mfrmUse.ShowMessage Me.TableName & " Lookup Table", "The table cannot be deleted as the table is used by the following:", UsageCheckObject.LookupTable
      End Select
      
      fOK = False
    End If
    UnLoad mfrmUse
    Set mfrmUse = Nothing
  End If


  ' Find & read table record to delete.
  If fOK Then
    
    ' Check if table is already deleted.
    If Not IsDeleted Then
      
'      fOK = Not TableIsUsed
      If fOK Then
        'Set record deleted flag
        gfDeleted = True
      
        'Write table record
        fOK = WriteTable
      End If
    
      ' Delete any child relations of this table
      If fOK Then
        fOK = DeleteAllRelations
      End If
    
      ' Delete any parent relations of this table (if they are not used elsewhere).
      If fOK Then
'        daoDb.Execute "DELETE FROM tmpRelations" & _
          " WHERE childID=" & gLngTableID, dbFailOnError
        sSQL = "SELECT tmpRelations.parentID" & _
          " FROM tmpRelations" & _
          " WHERE tmpRelations.childID = " & Trim(Str(gLngTableID))
        
        Set rsRelatedTables = daoDb.OpenRecordset(sSQL, _
          dbOpenForwardOnly, dbReadOnly)
                
        With rsRelatedTables
          Do While (Not .EOF) And fOK
            Set objTable = New Table
            objTable.TableID = rsRelatedTables.Fields("parentID")
            
            If objTable.ReadTable Then
'              fOK = Not objTable.RelationIsUsed(gLngTableID)
              'TM19012004
              ' Check that the relationship is not used.
              Set mfrmUse = New frmUsage
              mfrmUse.ResetList
              If RelationIsUsed(gLngTableID, mfrmUse) Then
                Screen.MousePointer = vbDefault
                mfrmUse.ShowMessage Me.TableName & " - " & GetTableName(gLngTableID) & " Relationship", "The relationship cannot be deleted as the relationship is used by the following:", UsageCheckObject.Relationship
                fOK = False
              End If
              UnLoad mfrmUse
              Set mfrmUse = Nothing

            End If
            
            Set objTable = Nothing
                    
            .MoveNext
          Loop
                  
          .Close
        End With
        
        If fOK Then
          daoDb.Execute "DELETE FROM tmpRelations" & _
            " WHERE childID = " & Trim(Str(gLngTableID)), dbFailOnError
        End If
      End If
    
      ' Delete all of the columns of this table.
      If fOK Then
        sSQL = "SELECT tmpColumns.columnID" & _
          " FROM tmpColumns" & _
          " WHERE tmpColumns.deleted = FALSE" & _
          " AND tmpColumns.tableID = " & Trim(Str(gLngTableID))
  
        Set rsColumns = daoDb.OpenRecordset(sSQL, _
          dbOpenForwardOnly, dbReadOnly)
        
        With rsColumns
          Do While (Not .EOF) And fOK
          
            Set objColumn = New Column
            objColumn.ColumnID = rsColumns.Fields("columnID")
            objColumn.TableID = gLngTableID
            fOK = objColumn.DeleteColumn
            Set objColumn = Nothing
            
            .MoveNext
          Loop
          
          .Close
        End With
      End If
    
      ' Delete all of the screens for this table (and all of the screens controls
      ' and history screens).
      If fOK Then
      
        ' Mark this table's screens as deleted.
        daoDb.Execute "UPDATE tmpScreens" & _
          " SET deleted = TRUE" & _
          " WHERE tableID = " & Trim(Str(gLngTableID)), dbFailOnError
        
        ' Delete controls for each deleted screen.
        sSQL = "DELETE tmpControls.*" & _
          " FROM tmpControls, tmpScreens" & _
          " WHERE tmpScreens.tableID = " & Trim(Str(gLngTableID)) & _
          " AND tmpScreens.screenID = tmpControls.screenID"
        daoDb.Execute sSQL
        
        ' Delete history screens for each deleted screen.
        sSQL = "DELETE tmpHistoryScreens.*" & _
          " FROM tmpHistoryScreens, tmpScreens" & _
          " WHERE tmpScreens.tableID = " & Trim(Str(gLngTableID)) & _
          " AND (tmpHistoryScreens.parentScreenID = tmpScreens.screenID" & _
          " OR tmpHistoryScreens.historyScreenID = tmpScreens.screenID)"
        daoDb.Execute sSQL

        ' Delete page captions for each deleted screen.
        sSQL = "DELETE tmpPageCaptions.*" & _
          " FROM tmpPageCaptions, tmpScreens" & _
          " WHERE tmpScreens.tableID = " & Trim(Str(gLngTableID)) & _
          " AND tmpScreens.screenID = tmpPageCaptions.screenID"
        daoDb.Execute sSQL

      End If
      
      ' Delete all of the orders for this table.
      If fOK Then
      
        sSQL = "SELECT tmpOrders.orderID" & _
          " FROM tmpOrders" & _
          " WHERE tmpOrders.deleted = FALSE" & _
          " AND tmpOrders.tableID = " & Trim(Str(gLngTableID))
  
        Set rsOrders = daoDb.OpenRecordset(sSQL, _
          dbOpenForwardOnly, dbReadOnly)
        
        With rsOrders
          Do While (Not .EOF) And fOK
          
            Set objOrder = New Order
            objOrder.OrderID = rsOrders.Fields("orderID")
            fOK = objOrder.DeleteOrder
            Set objOrder = Nothing
            
            .MoveNext
          Loop
          
          .Close
        End With
      End If
      
      
      'MH20000727
      ' Delete all of the email addresses for this table.
      If fOK Then
      
        sSQL = "SELECT tmpEMailAddresses.EmailID" & _
          " FROM tmpEMailAddresses" & _
          " WHERE tmpEMailAddresses.deleted = FALSE" & _
          " AND tmpEMailAddresses.tableID = " & Trim(Str(gLngTableID))
  
        Set rsEmailAddrs = daoDb.OpenRecordset(sSQL, _
          dbOpenForwardOnly, dbReadOnly)
        
        With rsEmailAddrs
          Do While (Not .EOF) And fOK
          
            Set objEmail = New clsEmailAddr
            objEmail.EmailID = rsEmailAddrs.Fields("EmailID")
            fOK = objEmail.DeleteEmail
            Set objEmail = Nothing
            
            .MoveNext
          Loop
          
          .Close
        End With
      End If
    
      
      ' Delete all of the expressions for this table.
      If fOK Then
        sSQL = "SELECT tmpExpressions.exprID" & _
          " FROM tmpExpressions" & _
          " WHERE tmpExpressions.deleted = FALSE" & _
          " AND tmpExpressions.TableID = " & Trim(Str(gLngTableID))
  
        Set rsExpressions = daoDb.OpenRecordset(sSQL, _
          dbOpenForwardOnly, dbReadOnly)
        
        With rsExpressions
          Do While (Not .EOF) And fOK
          
            Set objExpression = New CExpression
            objExpression.ExpressionID = rsExpressions.Fields("exprID")
            fOK = objExpression.DeleteExpression
            Set objExpression = Nothing
            
            .MoveNext
          Loop
        
          .Close
        End With
      End If
      
      ' Delete all of the views for this table.
      If fOK Then
        sSQL = "SELECT tmpViews.viewID" & _
          " FROM tmpViews" & _
          " WHERE viewTableID = " & Trim(Str(gLngTableID))
  
        Set rsViews = daoDb.OpenRecordset(sSQL, _
          dbOpenForwardOnly, dbReadOnly)
        
        With rsViews
          Do While (Not .EOF)
          
            ' Mark the view screen records as deleted.
            sSQL = "UPDATE tmpViewScreens" & _
              " SET deleted = TRUE" & _
              " WHERE tmpViewScreens.viewID = " & Trim(Str(.Fields("viewID")))
            daoDb.Execute sSQL, dbFailOnError
            
            ' Mark the view column records as deleted.
            sSQL = "UPDATE tmpViewColumns" & _
              " SET deleted = TRUE" & _
              " WHERE tmpViewColumns.viewID = " & Trim(Str(.Fields("viewID")))
            daoDb.Execute sSQL, dbFailOnError
        
            .MoveNext
          Loop
          
          .Close
        End With
              
        ' Mark the view master records as deleted.
        sSQL = "UPDATE tmpViews" & _
          " SET deleted = TRUE" & _
          " WHERE viewTableID = " & Trim(Str(gLngTableID))
        daoDb.Execute sSQL, dbFailOnError
      End If
      
      
      ' Delete any summary field records for this table.
      If fOK Then
        fOK = ClearSummaryFields
        
        If fOK Then
          sSQL = "DELETE FROM tmpSummary" & _
            " WHERE tmpSummary.historyTableID = " & Trim(Str(gLngTableID))
          daoDb.Execute sSQL
        End If
      End If
    
    
      If fOK Then
        sSQL = "UPDATE tmpEmailLinks" & _
          " SET deleted = TRUE " & _
          " WHERE tmpEmailLinks.TableID = " & CStr(gLngTableID)
        daoDb.Execute sSQL
      End If
    
    
    End If
  End If
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set rsColumns = Nothing
  Set rsOrders = Nothing
  Set rsExpressions = Nothing
  Set rsViews = Nothing
  Set rsRelatedTables = Nothing
  Set objColumn = Nothing
  Set objOrder = Nothing
  Set objExpression = Nothing
  
  ' Commit the data transaction if everything was okay.
  If fOK Then
    gfDeleted = True
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If

  DeleteTable = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function EditTable() As Boolean
  ' Edit the table's properties.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim frmEdit As frmTabEdit
  
  fOK = ReadTable
  
  ' Find & read table record to edit
  If fOK Then
  
    Set frmEdit = New frmTabEdit
    Set frmEdit.Table = Me
    frmEdit.Show vbModal
    
    fOK = Not frmEdit.Cancelled
    
    If fOK Then
      fOK = WriteTable_Transaction
    End If
    
  End If
  
TidyUpAndExit:
  Set frmEdit = Nothing
  EditTable = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Public Function CloneTable(pavCloneRegister As Variant, pfCopyData As Boolean, pfCopySecurity As Boolean _
  , pbGrantRead As Boolean _
  , pbGrantNew As Boolean _
  , pbGrantEdit As Boolean _
  , pbGrantDelete As Boolean) As Table
  
  ' Copy the table.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fGoodName As Boolean
  Dim fParentCloned As Boolean
  Dim iIndex As Integer
  Dim iCounter As Integer
  Dim lngViewID As Long
  Dim lngOrderId As Long
  Dim lngScreenID As Long
  Dim lngColumnID As Long
  'NPG20080507 Fault 12974
  Dim lngTableID As Long
  Dim lngCalcExprID As Long
  Dim lngExpressionID As Long
  Dim lngOldIDColumnID As Long
  Dim lngNewIDColumnID As Long
  Dim lngGotFocusExprID As Long
  Dim lngDfltValueExprID As Long  'MH20060613
  Dim lngLostFocusExprID As Long
  Dim lngLookupFilterValueID As Long
  Dim lngHistoryScreenID As Long
  Dim sSQL As String
  Dim sViewName As String
  Dim sTableName As String
  Dim sScreenName As String
  Dim rsViews As DAO.Recordset
  Dim rsOrders As DAO.Recordset
  Dim rsScreens As DAO.Recordset
  Dim rsWorkflows As DAO.Recordset
  Dim rsColumns As DAO.Recordset
  Dim rsFilters As DAO.Recordset
  Dim rsChildren As DAO.Recordset
  Dim rsControls As DAO.Recordset
  Dim rsExpressions As DAO.Recordset
  Dim rsViewScreens As DAO.Recordset
  Dim rsViewColumns As DAO.Recordset
  Dim rsPageCaptions As DAO.Recordset
  Dim rsHistoryScreens As DAO.Recordset
  Dim objNewTable As Table
  Dim objChildTable As Table
  Dim objCopyChildTable As Table
  Dim objNewOrder As Order
  Dim objSourceOrder As Order
  Dim objNewColumn As Column
  Dim objSourceColumn As Column
  Dim objNewExpr As CExpression
  Dim objSourceExpr As CExpression
  Dim rsParents As DAO.Recordset
  Dim rsSummaryFields As DAO.Recordset
  Dim objSourceSummary As cSummaryField
  Dim objNewSummary As cSummaryField
  Dim lngOriginalTableID As Long
  
  Dim objSourceEmail As clsEmailAddr
  Dim objNewEmail As clsEmailAddr
  Dim objSourceEmailLink As clsEmailLink
  Dim objNewEmailLink As clsEmailLink
  Dim rsEmailDef As DAO.Recordset
  
  'NPG20071120 Fault 12614
  Dim objSourceDiaryLink As cDiaryLink
  Dim objNewDiaryLink As cDiaryLink
  Dim rsDiaryDef As DAO.Recordset
  
  
  Dim mstrEmailNotCopied As String
  Dim i As Integer
  'Dim lngNewEmailID As Long
  'Dim objEmailLinks As Collection

  Dim rsOutlookDef As DAO.Recordset
  Dim objSourceOutlookFolder As clsOutlookFolder
  Dim objNewOutlookFolder As clsOutlookFolder
  Dim objSourceOutlookLink As clsOutlookLink
  Dim objNewOutlookLink As clsOutlookLink
  Dim objSourceOutlookLinkColumn As clsOutlookLinkColumn
  Dim objNewOutlookLinkColumn As clsOutlookLinkColumn

  Dim rsWorkflowTriggeredLinkDef As DAO.Recordset
  Dim objSourceWorkflowTriggeredLink As clsWorkflowTriggeredLink
  Dim objNewWorkflowTriggeredLink As clsWorkflowTriggeredLink
  Dim objSourceWorkflowTriggeredLinkColumn As clsWFTriggeredLinkColumn
  Dim objNewWorkflowTriggeredLinkColumn As clsWFTriggeredLinkColumn

  Dim bGrantRead As Boolean
  Dim bGrantNew As Boolean
  Dim bGrantEdit As Boolean
  Dim bGrantDelete As Boolean
  Dim lngCopyTableID As Long
  Dim strCopyTableName As String

  ' Find & read the current table record.
  fOK = ReadTable
  
  ' Calculate the default permissions that are applied to copied objects
  If Not pfCopySecurity Then
    bGrantRead = pbGrantRead
    bGrantNew = pbGrantNew
    bGrantEdit = pbGrantEdit
    bGrantDelete = pbGrantDelete
    lngCopyTableID = 0
    strCopyTableName = ""
  Else
    'MH20100625
    'bGrantRead = 0
    'bGrantNew = 0
    'bGrantEdit = 0
    'bGrantDelete = 0
    bGrantRead = miGrantRead
    bGrantNew = miGrantNew
    bGrantEdit = miGrantEdit
    bGrantDelete = miGrantDelete
    
    lngCopyTableID = TableID
    strCopyTableName = TableName
  End If
  
  ' Instantiate a new table object, and copy the source table's definition to it.
  If fOK Then

    Set objNewTable = New Table
    
    With objNewTable
      .IsNew = True
      .IsChanged = False
      .IsDeleted = False
      .TableID = 0
      
      'JPD 20030829 Fault 5538 - set the flag to stop the user being prompted for
      ' default permissions next time they go into the table properties screen.
      ' The default permissions are defined when copying the tables.
      .PermissionsPrompted = True

      ' Copy the table's expression and order IDs later as they require the column IDs.
      .RecordDescriptionID = 0
      .PrimaryOrderID = 0
          
      ' Create a new table name.
      sTableName = "Copy_of_" & TableName
      ' Check that the table name is not already used.
      iCounter = 1
      fGoodName = False
      Do While Not fGoodName
        With recTabEdit
          .Index = "idxName"
          .Seek "=", sTableName, False
          If Not .NoMatch Then
            iCounter = iCounter + 1
            sTableName = "Copy_" & Trim(Str(iCounter)) & "_of_" & TableName
          Else
            'AE20071205 Fault #12679
'            fGoodName = True
'          End If
            
            recViewEdit.Index = "idxViewName"
            recViewEdit.Seek "=", sTableName, False
            
            If Not recViewEdit.NoMatch Then
              iCounter = iCounter + 1
              sTableName = "Copy_" & Trim(Str(iCounter)) & "_of_" & TableName
            Else
              fGoodName = True
            End If
          End If
        End With
      Loop
      .TableName = sTableName
      .TableType = TableType
      
      'Add the default security permissions
      .GrantRead = bGrantRead
      .GrantNew = bGrantNew
      .GrantEdit = bGrantEdit
      .GrantDelete = bGrantDelete
      .IsCopy = True
      .CopyTableID = lngCopyTableID
      .CopyTableName = strCopyTableName
      .ManualSummaryColumnBreaks = mbManualColumnBreaks
      .AuditDelete = mbAuditDelete
      .AuditInsert = mbAuditInsert
      
      'NPG20071120 Fault 12615
      'Default email to source ID. This will NOT be overwritten for Fixed email addresses
      '.EmailDeleteID = mlngEmailDelete
      '.EmailInsertID = mlngEmailInsert
      .PrimaryEmailID = glngPrimaryEmailID
      
      .InsertTriggerDisabled = mbInsertTriggerDisabled
      .UpdateTriggerDisabled = mbUpdateTriggerDisabled
      .DeleteTriggerDisabled = mbDeleteTriggerDisabled
      
      ' Write the new table definition to the local tables.
      fOK = .WriteTable
            
      If fOK Then
        ' Remember the table IDs of the original and copied tables.
        iIndex = UBound(pavCloneRegister, 2) + 1
        ReDim Preserve pavCloneRegister(3, iIndex)
        pavCloneRegister(1, iIndex) = "TABLE"
        pavCloneRegister(2, iIndex) = TableID
        pavCloneRegister(3, iIndex) = .TableID
      End If
      
    End With
  End If
  
  ' Remember the source table ID if we need to copy data from it.
  If fOK Then
    If pfCopyData Then
      With recTabEdit
        .Index = "idxTableID"
        .Seek "=", objNewTable.TableID
        .Edit
        .Fields("copyDataTableID") = TableID
        .Update
      End With
    End If
  End If
  
  ' Remember the source table's 'ID' column ID.
  If fOK Then
    With recColEdit
      .Index = "idxName"
      .Seek "=", TableID, "ID", False
      
      fOK = Not .NoMatch
      
      If fOK Then
        lngOldIDColumnID = !ColumnID
      
        .Seek "=", objNewTable.TableID, "ID", False
        fOK = Not .NoMatch
        
        If fOK Then
          lngNewIDColumnID = !ColumnID
        
          iIndex = UBound(pavCloneRegister, 2) + 1
          ReDim Preserve pavCloneRegister(3, iIndex)
          pavCloneRegister(1, iIndex) = "COLUMN"
          pavCloneRegister(2, iIndex) = lngOldIDColumnID
          pavCloneRegister(3, iIndex) = lngNewIDColumnID
        End If
      End If
    End With
  End If
    
  ' Ensure that the relationships between the cloned table and its parents
  ' are set up.
  ' NB. We only need to clone the relationship to parenttables not
  ' already cloned in this operation.
  If fOK Then
    sSQL = "SELECT tmpRelations.parentID" & _
      " FROM tmpRelations" & _
      " WHERE tmpRelations.childID = " & Trim(Str(gLngTableID))
    Set rsParents = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    With rsParents
      ' For each parent table definition ...
      Do While (Not .EOF) And fOK
        ' Check that the relationship has not already been cloned.
        fParentCloned = False
        For iIndex = 1 To UBound(pavCloneRegister, 2)
          If pavCloneRegister(1, iIndex) = "TABLE" And _
            pavCloneRegister(2, iIndex) = !parentID Then
            
            fParentCloned = True
            Exit For
          End If
        Next iIndex
        
        If Not fParentCloned Then
          With recRelEdit
            ' Create the relationship record in the database.
            .AddNew
            !parentID = rsParents.Fields("parentID")
            !childID = objNewTable.TableID
            .Update
      
            ' Create the column in the child table that links it to the ID column in the
            ' parent table.
            With objNewTable
              If .AddIDColumn(rsParents.Fields("parentID")) Then
                fOK = .WriteTable
              End If
            End With
          End With
         End If
         
        .MoveNext
      Loop
    
      .Close
    End With
    ' Disassociate object variables.
    Set rsParents = Nothing
  End If
  
  '
  ' Copy the column definitions.
  '
  If fOK Then
  
    ' Get the table's column definitions from the database.
    sSQL = "SELECT tmpColumns.columnID" & _
      " FROM tmpColumns" & _
      " WHERE tmpColumns.deleted = FALSE" & _
      " AND tmpColumns.tableID = " & Trim(Str(gLngTableID)) & _
      " AND tmpColumns.columnType <> " & Trim(Str(giCOLUMNTYPE_SYSTEM))
    Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
        
    With rsColumns
      ' For each column definition ...
      Do While (Not .EOF) And fOK
      
        ' Instantiate the original column object.
        Set objSourceColumn = New Column
        objSourceColumn.ColumnID = .Fields("columnID").value
        objSourceColumn.TableID = gLngTableID

        ' Instruct the original column object to clone itself.
        Set objNewColumn = objSourceColumn.CloneColumn(False)
        fOK = Not objNewColumn Is Nothing

        If fOK Then
          With objNewColumn
            ' Copy properties from the original column to the copy.
            .TableID = objNewTable.TableID
            ' Reset expression IDs. These will be done later when the expression are copied.
            .Properties("calcExprID").value = 0
            .Properties("gotFocusExprID").value = 0
            .Properties("dfltValueExprID").value = 0  'MH20060613
            .Properties("lostFocusExprID").value = 0
            
            ' Update the link table if the column is a link column.
            If .Properties("columnType").value = giCOLUMNTYPE_LINK Then
              For iIndex = 1 To UBound(pavCloneRegister, 2)
                If pavCloneRegister(1, iIndex) = "TABLE" And _
                  pavCloneRegister(2, iIndex) = .Properties("linkTableID").value Then
                  .Properties("linkTableID").value = pavCloneRegister(3, iIndex)
                  Exit For
                End If
              Next iIndex

              For iIndex = 1 To UBound(pavCloneRegister, 2)
                If pavCloneRegister(1, iIndex) = "VIEW" And _
                  pavCloneRegister(2, iIndex) = .Properties("linkViewID").value Then
                  .Properties("linkViewID").value = pavCloneRegister(3, iIndex)
                  Exit For
                End If
              Next iIndex
            End If

            'JPD20010730
            If .Properties("uniqueCheckType").value > 0 Then
              For iIndex = 1 To UBound(pavCloneRegister, 2)
                If pavCloneRegister(1, iIndex) = "TABLE" And _
                  pavCloneRegister(2, iIndex) = .Properties("uniqueCheckType").value Then
                  .Properties("uniqueCheckType").value = pavCloneRegister(3, iIndex)
                  Exit For
                End If
              Next iIndex
            End If
            
            ' Write the copied column definition to the database.
            fOK = .WriteColumn
          End With
          
          ' Remember the IDs of the original and copied columns.
          If fOK Then
            iIndex = UBound(pavCloneRegister, 2) + 1
            ReDim Preserve pavCloneRegister(3, iIndex)
            pavCloneRegister(1, iIndex) = "COLUMN"
            pavCloneRegister(2, iIndex) = objSourceColumn.ColumnID
            pavCloneRegister(3, iIndex) = objNewColumn.ColumnID
          End If
        End If
          
        ' Disassociate object variables.
        Set objSourceColumn = Nothing
        Set objNewColumn = Nothing
            
        .MoveNext
      Loop
          
      .Close
    End With
    ' Disassociate object variables.
    Set rsColumns = Nothing
  End If
    
  '
  ' Copy the order definitions.
  '
  ' Get the table's order definitions from the database.
  If fOK Then
  
    sSQL = "SELECT tmpOrders.orderID" & _
      " FROM tmpOrders" & _
      " WHERE tmpOrders.deleted = FALSE" & _
      " AND tmpOrders.tableID = " & Trim(Str(gLngTableID))
    Set rsOrders = daoDb.OpenRecordset(sSQL, _
      dbOpenForwardOnly, dbReadOnly)
        
    With rsOrders
      ' For each order definition ...
      Do While (Not .EOF) And fOK
      
        ' Instantiate the original order object.
        Set objSourceOrder = New Order
        objSourceOrder.OrderID = .Fields("orderID").value
        ' Instruct the original order object to clone itself.
        Set objNewOrder = objSourceOrder.CloneOrder(pavCloneRegister)
               
        fOK = Not objNewOrder Is Nothing
        
        If fOK Then
          With objNewOrder
            ' Copy properties from the original column to the copy.
            .TableID = objNewTable.TableID
            ' Write the copied order definition to the database.
            fOK = .WriteOrder
          End With
        
          ' Remember the IDs of the original and copied orders.
          If fOK Then
            iIndex = UBound(pavCloneRegister, 2) + 1
            ReDim Preserve pavCloneRegister(3, iIndex)
            pavCloneRegister(1, iIndex) = "ORDER"
            pavCloneRegister(2, iIndex) = objSourceOrder.OrderID
            pavCloneRegister(3, iIndex) = objNewOrder.OrderID
          End If
        End If
                
        ' Disassociate object variables.
        Set objSourceOrder = Nothing
        Set objNewOrder = Nothing
                
        .MoveNext
      Loop
              
      .Close
    End With
    ' Disassociate object variables.
    Set rsOrders = Nothing

    ' Update the copied table with the new order ID of it's primary order.
    If fOK Then
      If gLngPrimaryOrderID > 0 Then
        For iIndex = 1 To UBound(pavCloneRegister, 2)
          If pavCloneRegister(1, iIndex) = "ORDER" And _
            pavCloneRegister(2, iIndex) = gLngPrimaryOrderID Then
            With objNewTable
              .PrimaryOrderID = pavCloneRegister(3, iIndex)
              .WriteTable
            End With
            Exit For
          End If
        Next iIndex
      End If
    End If
  End If
    
  '
  ' Copy the summary field definitions.
  '
  If fOK Then
    ' Get the table's Summary Field definitions from the database.
    sSQL = "SELECT tmpSummary.ID" & _
      " FROM tmpSummary" & _
      " WHERE tmpSummary.historyTableID = " & Trim(Str(gLngTableID))
    Set rsSummaryFields = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
        
    With rsSummaryFields
      ' For each Summary Field definition ...
      Do While (Not .EOF) And fOK
        ' Instantiate the original Summary Field object.
        Set objSourceSummary = New cSummaryField
        objSourceSummary.ID = .Fields("ID")
        Set objNewSummary = objSourceSummary.CloneSummaryField
           
        fOK = Not objNewSummary Is Nothing
        
        If fOK Then
          ' Copy properties from the original Summary Field to the copy.
          objNewSummary.HistoryTableID = objNewTable.TableID
          objNewTable.SummaryFields.Add objNewSummary
          ' Write the copied Summary Field definition to the database.
          fOK = objNewSummary.WriteSummaryField
        End If
        
        ' Disassociate object variables.
        Set objSourceSummary = Nothing
        Set objNewSummary = Nothing
            
        .MoveNext
      Loop
          
      .Close
    End With
    ' Disassociate object variables.
    Set rsSummaryFields = Nothing
  
  End If
  
  '
  ' Copy the child tables.
  '
  If fOK Then
  
    ' Get the table's relation definitions from the database.
    sSQL = "SELECT tmpRelations.childID" & _
      " FROM tmpRelations" & _
      " WHERE tmpRelations.parentID = " & Trim(Str(gLngTableID))
    Set rsChildren = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
        
    With rsChildren
      ' For each relation definition ...
      Do While (Not .EOF) And fOK
        ' Instantiate the original child table object.
        Set objChildTable = New SystemMgr.Table
        objChildTable.TableID = .Fields("childID")
        ' Instruct the original child table object to clone itself.
        Set objCopyChildTable = objChildTable.CloneTable(pavCloneRegister, pfCopyData, pfCopySecurity, bGrantRead, bGrantNew, bGrantEdit, bGrantDelete)
            
        fOK = Not objCopyChildTable Is Nothing

        If fOK Then
          mstrExprNotCopied = mstrExprNotCopied & objChildTable.ExpressionsNotCopied
        End If

        ' Add the copied relation to the database.
        If fOK Then
          With recRelEdit
            .AddNew
            !parentID = objNewTable.TableID
            !childID = objCopyChildTable.TableID
            .Update
          End With
          
          ' Add an ID column to the copied child table for the copied relation.
          With objCopyChildTable
            If .AddIDColumn(objNewTable.TableID) Then

              ''MH20010417
              'iIndex = UBound(pavCloneRegister, 2) + 1
              'ReDim Preserve pavCloneRegister(3, iIndex)
              'pavCloneRegister(1, iIndex) = "COLUMN"
              'pavCloneRegister(2, iIndex) = "ID_" & TableID
              'pavCloneRegister(3, iIndex) = "ID_" & objNewTable.TableID
              
              fOK = .WriteTable
            End If
          End With
        End If
  
  
        ''MH20010417
        '' Remember the source table ID if we need to copy data from it.
        'If fOK Then
        '  If pfCopyData Then
        '    With recTabEdit
        '      .Index = "idxTableID"
        '      .Seek "=", objCopyChildTable.TableID
        '      .Edit
        '      .Fields("copyDataTableID") = objChildTable.TableID
        '      .Update
        '    End With
        '  End If
        'End If
        
        
        ' Disassociate object variables.
        Set objChildTable = Nothing
        Set objCopyChildTable = Nothing
              
        .MoveNext
      Loop
            
      .Close
    End With
    ' Disassociate object variables.
    Set rsChildren = Nothing
  End If
  
  '
  ' Copy the expression definitions.
  '
  If fOK Then
    ' Get the table's expression definitions from the database.
    'sSQL = "SELECT tmpExpressions.exprID" & _
      " FROM tmpExpressions" & _
      " WHERE tmpExpressions.deleted = FALSE" & _
      " AND tmpExpressions.TableID = " & Trim(Str(gLngTableID)) & _
      " AND tmpExpressions.parentComponentID = 0"
    ' NB. Workflow expressions are cloned as part of the CloneWorkflow method.
    sSQL = "SELECT tmpExpressions.exprID" & _
      " FROM tmpExpressions" & _
      " WHERE tmpExpressions.deleted = FALSE" & _
      " AND tmpExpressions.TableID = " & Trim(Str(gLngTableID)) & _
      " AND tmpExpressions.parentComponentID = 0 " & _
      " AND tmpExpressions.Type <> " & CStr(giEXPR_RUNTIMECALCULATION) & _
      " AND tmpExpressions.Type <> " & CStr(giEXPR_RUNTIMEFILTER) & _
      " AND tmpExpressions.Type <> " & CStr(giEXPR_WORKFLOWCALCULATION) & _
      " AND tmpExpressions.Type <> " & CStr(giEXPR_WORKFLOWSTATICFILTER) & _
      " AND tmpExpressions.Type <> " & CStr(giEXPR_WORKFLOWRUNTIMEFILTER)
    Set rsExpressions = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    With rsExpressions
      ' For each expression definition ...
      Do While (Not .EOF) And fOK
        ' Instantiate the original expression object.
        Set objSourceExpr = New CExpression
        objSourceExpr.ExpressionID = .Fields("exprID")

        Set objNewExpr = objSourceExpr.CloneExpression(pavCloneRegister)
        fOK = Not objNewExpr Is Nothing
        
        If fOK Then
          ' Copy properties from the original expression to the copy.
          objNewExpr.BaseTableID = objNewTable.TableID
          ' Write the copied expession definition to the database.
          fOK = objNewExpr.WriteExpression
        End If
        
        ' Remember the IDs of the original and copied orders.
        'If fOK Then
        iIndex = UBound(pavCloneRegister, 2) + 1
        ReDim Preserve pavCloneRegister(3, iIndex)
        pavCloneRegister(1, iIndex) = "EXPRESSION"
        pavCloneRegister(2, iIndex) = objSourceExpr.ExpressionID
        
        If fOK Then
          pavCloneRegister(3, iIndex) = objNewExpr.ExpressionID
        Else
          pavCloneRegister(3, iIndex) = 0
          
          mstrExprNotCopied = mstrExprNotCopied & _
              objSourceExpr.ExpressionTypeName & _
              ": '" & objSourceExpr.Name & "'" & vbCrLf
          fOK = True
        End If

        ' Disassociate object variables.
        Set objSourceExpr = Nothing
        Set objNewExpr = Nothing

        .MoveNext
      Loop
          
      .Close
    End With
    ' Disassociate object variables.
    Set rsExpressions = Nothing
    
  End If
  
  'TM20020430 Fault 3814
  '
  ' Copy the email definitions.
  '
  If fOK Then
    ' Get the table's email definitions from the database.
    sSQL = "SELECT tmpEmailAddresses.EmailID" & _
      " FROM tmpEmailAddresses" & _
      " WHERE tmpEmailAddresses.deleted = FALSE" & _
      " AND tmpEmailAddresses.TableID = " & CStr(gLngTableID)
      
    Set rsEmailDef = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    With rsEmailDef
      ' For each email definition ...
      Do While (Not .EOF) And fOK
        ' Instantiate the original email object.
        Set objSourceEmail = New clsEmailAddr
        objSourceEmail.EmailID = .Fields("EmailID")

        Set objNewEmail = objSourceEmail.CloneEmail(pavCloneRegister)
        fOK = Not objNewEmail Is Nothing
        
        If fOK Then
          ' Copy properties from the original email to the copy.
          objNewEmail.TableID = objNewTable.TableID
          ' Write the copied email definition to the database.
          fOK = objNewEmail.WriteEmail
        End If
        
        ' Remember the IDs of the original and copied emails.
        'If fOK Then
        iIndex = UBound(pavCloneRegister, 2) + 1
        ReDim Preserve pavCloneRegister(3, iIndex)
        pavCloneRegister(1, iIndex) = "EMAIL"
        pavCloneRegister(2, iIndex) = objSourceEmail.EmailID
        
        If fOK Then
          pavCloneRegister(3, iIndex) = objNewEmail.EmailID
        Else
          pavCloneRegister(3, iIndex) = 0
          
          mstrEmailNotCopied = mstrEmailNotCopied & objSourceEmail.EmailName & vbCrLf
          fOK = True
        End If
      
        ' Disassociate object variables.
        Set objSourceEmail = Nothing
        Set objNewEmail = Nothing
        Set rsColumns = Nothing
        
        .MoveNext
      Loop
          
      .Close
    End With
    ' Disassociate object variables.
    Set rsEmailDef = Nothing
    
  End If

''''''''''''  If fOK Then
''''''''''''    'Update the copied columns's email links with the new IDs of their columns and calculations.
''''''''''''    If fOK Then
''''''''''''      sSQL = "SELECT tmpEmailLinks.columnID, tmpEmailLinks.linkID " & _
''''''''''''            " FROM tmpEmailLinks " & _
''''''''''''            " WHERE tmpEmailLinks.deleted = FALSE" & _
''''''''''''               " AND tmpEmailLinks.columnID IN (SELECT tmpColumns.columnID" & _
''''''''''''                                              " FROM tmpColumns" & _
''''''''''''                                              " WHERE tmpColumns.deleted = FALSE" & _
''''''''''''                                                " AND tmpColumns.tableID = " & Trim(Str(gLngTableID)) & ")"
''''''''''''
''''''''''''      Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
''''''''''''
''''''''''''      With rsColumns
''''''''''''        Do While (Not .EOF) And fOK
''''''''''''          lngColumnID = 0
''''''''''''          For iIndex = 1 To UBound(pavCloneRegister, 2)
''''''''''''            If pavCloneRegister(1, iIndex) = "COLUMN" And _
''''''''''''              pavCloneRegister(2, iIndex) = .Fields("columnID") Then
''''''''''''              lngColumnID = pavCloneRegister(3, iIndex)
''''''''''''              Exit For
''''''''''''            End If
''''''''''''          Next iIndex
''''''''''''
''''''''''''          Set objNewColumn = New Column
''''''''''''          With objNewColumn
''''''''''''            .ColumnID = lngColumnID
''''''''''''
''''''''''''            fOK = .ReadColumn
''''''''''''
''''''''''''            If fOK Then
''''''''''''              Set objEmailLinks = New Collection
''''''''''''
''''''''''''              If .EmailLinks.Count > 0 Then
''''''''''''                For Each objSourceEmailLink In .EmailLinks
''''''''''''
''''''''''''                  Set objNewEmailLink = New clsEmailLink
''''''''''''                  Set objNewEmailLink = objSourceEmailLink.CloneLink(pavCloneRegister)
''''''''''''
''''''''''''                  objEmailLinks.Add objNewEmailLink
''''''''''''
''''''''''''                  Set objNewEmailLink = Nothing
''''''''''''
''''''''''''                Next objSourceEmailLink
''''''''''''
''''''''''''                Set objNewColumn.EmailLinks = objEmailLinks
''''''''''''
''''''''''''                fOK = .WriteColumn
''''''''''''              End If
''''''''''''            End If
''''''''''''          End With
''''''''''''          ' Disassociate object variables.
''''''''''''          Set objEmailLinks = Nothing
''''''''''''          Set objNewColumn = Nothing
''''''''''''          Set objSourceEmailLink = Nothing
''''''''''''          .MoveNext
''''''''''''        Loop
''''''''''''      End With
''''''''''''      Set rsColumns = Nothing
''''''''''''    End If
''''''''''''  End If
''''''''''''  'End of copying email.



  'EmailLinks

  Dim objOriginalLink As clsEmailLink
  Dim objNewLink As clsEmailLink
  Dim lngID As Long

  For Each objOriginalLink In mcolEmailLinks

    Set objNewLink = objOriginalLink.CloneLink

    objNewLink.TableID = objNewTable.TableID
    objNewLink.FilterID = GetClone(pavCloneRegister, "EXPRESSION", objOriginalLink.FilterID)
    objNewLink.DateColumnID = GetClone(pavCloneRegister, "COLUMN", objOriginalLink.DateColumnID)
    
    'Set objNewLink.SubjectContent = New clsLinkContent
    objNewLink.SubjectContent.CloneDetail pavCloneRegister, objOriginalLink.SubjectContent.GetID
    
    'Set objNewLink.BodyContent = New clsLinkContent
    objNewLink.BodyContent.CloneDetail pavCloneRegister, objOriginalLink.BodyContent.GetID
    
    
    'Columns
    For iIndex = 1 To objOriginalLink.Columns.Count
      lngID = GetClone(pavCloneRegister, "COLUMN", objOriginalLink.Columns.Item(iIndex))
      objNewLink.Columns.Add lngID, CStr(lngID)
    Next

    'To recipients
    For iIndex = 1 To objOriginalLink.RecipientsTo.Count
      lngID = GetClone(pavCloneRegister, "EMAIL", objOriginalLink.RecipientsTo.Item(iIndex))
      objNewLink.RecipientsTo.Add lngID, CStr(lngID)
    Next

    'Cc recipients
    For iIndex = 1 To objOriginalLink.RecipientsCc.Count
      lngID = GetClone(pavCloneRegister, "EMAIL", objOriginalLink.RecipientsCc.Item(iIndex))
      objNewLink.RecipientsCc.Add lngID, CStr(lngID)
    Next

    'Bcc recipients
    For iIndex = 1 To objOriginalLink.RecipientsBcc.Count
      lngID = GetClone(pavCloneRegister, "EMAIL", objOriginalLink.RecipientsBcc.Item(iIndex))
      objNewLink.RecipientsBcc.Add lngID, CStr(lngID)
    Next

    objNewTable.EmailLinks.Add objNewLink

  Next
  
  
  
  
  
  'NPG20071120 Fault 12614
  '
  ' Copy the Diary Filters
  '
  If fOK Then
    ' Get the table's diary definitions from the database.
    sSQL = "SELECT tmpDiary.DiaryID, tmpDiary.FilterID" & _
      " FROM tmpDiary" & _
      " WHERE tmpDiary.ColumnID IN (SELECT ColumnID FROM tmpColumns WHERE TableID = " & CStr(objNewTable.TableID) & ")"

    Set rsDiaryDef = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    With rsDiaryDef
      ' For each diary definition ...
      Do While (Not .EOF) And fOK
        ' Instantiate the original diary object.
        
        If fOK Then
          ' Copy properties from the original Diary Link to the copy.
          For iIndex = 1 To UBound(pavCloneRegister, 2)
            If pavCloneRegister(1, iIndex) = "EXPRESSION" And _
               pavCloneRegister(2, iIndex) = .Fields("FilterID") Then

                recDiaryEdit.Index = "idxDiaryLinkID"
                recDiaryEdit.Seek "=", .Fields("DiaryID")

                If Not recDiaryEdit.NoMatch Then
                  recDiaryEdit.Edit
                  recDiaryEdit!FilterID = pavCloneRegister(3, iIndex)
                  recDiaryEdit.Update
                End If
                Exit For

            End If
          Next iIndex
        End If

        .MoveNext
      Loop

      .Close
    End With

    Set rsDiaryDef = Nothing

  End If
   
  ' Update the copied table with the new ID of it's record description expression.
  If fOK Then
  
    If gLngRecordDescriptionID > 0 Then
      For iIndex = 1 To UBound(pavCloneRegister, 2)
        If pavCloneRegister(1, iIndex) = "EXPRESSION" And _
          pavCloneRegister(2, iIndex) = gLngRecordDescriptionID Then
          objNewTable.RecordDescriptionID = pavCloneRegister(3, iIndex)
          fOK = objNewTable.WriteTable
          Exit For
        End If
      Next iIndex
    End If
  
    'NPG20071121 Fault 12615
    ' If it's a column or calc email then update the EmailID to the new definition
    'If glngPrimaryEmailID > 0 Then
      For iIndex = 1 To UBound(pavCloneRegister, 2)
        If pavCloneRegister(1, iIndex) = "EMAIL" Then
        
          If pavCloneRegister(2, iIndex) = glngPrimaryEmailID Then
            objNewTable.PrimaryEmailID = pavCloneRegister(3, iIndex)
            fOK = objNewTable.WriteTable
          End If

'          If pavCloneRegister(2, iIndex) = mlngEmailInsert Then
'            objNewTable.EmailInsertID = pavCloneRegister(3, iIndex)
'            fOK = objNewTable.WriteTable
'          End If
'
'          If pavCloneRegister(2, iIndex) = mlngEmailDelete Then
'            objNewTable.EmailDeleteID = pavCloneRegister(3, iIndex)
'            fOK = objNewTable.WriteTable
'          End If
        
        End If
      Next iIndex
    'End If

  End If
  
  If fOK Then
    ' Update the copied table's columns with the new IDs of their calculation, gotFocus, and lostFocus expressions.
    'JPD 20050810 Fault 10249 Also update the lookupFilterValueID if required.
    
'MH20060613
'    sSQL = "SELECT tmpColumns.columnID, tmpColumns.calcExprID, tmpColumns.gotFocusExprID, tmpColumns.lostFocusExprID, tmpColumns.lookupFilterValueID" & _
      " FROM tmpColumns" & _
      " WHERE tmpColumns.deleted = FALSE" & _
      " AND tmpColumns.tableID = " & Trim(Str(gLngTableID)) & _
      " AND tmpColumns.columnType <> " & Trim(Str(giCOLUMNTYPE_SYSTEM)) & _
      " AND ((tmpColumns.calcExprID <> 0) OR (tmpColumns.gotFocusExprID <> 0) OR (tmpColumns.dfltValueExprID <> 0) OR (tmpColumns.lostFocusExprID <> 0) OR (tmpColumns.lookupFilterValueID <> 0))"
    sSQL = "SELECT tmpColumns.columnID, tmpColumns.calcExprID, tmpColumns.gotFocusExprID, tmpColumns.dfltValueExprID, tmpColumns.lostFocusExprID, tmpColumns.lookupFilterValueID" & _
      " FROM tmpColumns" & _
      " WHERE tmpColumns.deleted = FALSE" & _
      " AND tmpColumns.tableID = " & Trim(Str(gLngTableID)) & _
      " AND tmpColumns.columnType <> " & Trim(Str(giCOLUMNTYPE_SYSTEM)) & _
      " AND ((tmpColumns.calcExprID <> 0) OR (tmpColumns.gotFocusExprID <> 0) OR (tmpColumns.dfltValueExprID <> 0) OR (tmpColumns.lostFocusExprID <> 0) OR (tmpColumns.lookupFilterValueID <> 0))"
    
    Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
        
    With rsColumns
      Do While (Not .EOF) And fOK
        
        lngColumnID = 0
        lngCalcExprID = 0
        lngGotFocusExprID = 0
        lngDfltValueExprID = 0  'MH20060613
        lngLostFocusExprID = 0
        lngLookupFilterValueID = 0
        
        For iIndex = 1 To UBound(pavCloneRegister, 2)
          
          If pavCloneRegister(1, iIndex) = "COLUMN" And _
            pavCloneRegister(2, iIndex) = .Fields("columnID") Then
            lngColumnID = pavCloneRegister(3, iIndex)
          End If
          If pavCloneRegister(1, iIndex) = "EXPRESSION" And _
            pavCloneRegister(2, iIndex) = .Fields("calcExprID") Then
            lngCalcExprID = pavCloneRegister(3, iIndex)
          End If
          If pavCloneRegister(1, iIndex) = "EXPRESSION" And _
            pavCloneRegister(2, iIndex) = .Fields("gotFocusExprID") Then
            lngGotFocusExprID = pavCloneRegister(3, iIndex)
          End If

          'MH20060613
          If pavCloneRegister(1, iIndex) = "EXPRESSION" And _
            pavCloneRegister(2, iIndex) = .Fields("dfltValueExprID") Then
            lngDfltValueExprID = pavCloneRegister(3, iIndex)
          End If

          If pavCloneRegister(1, iIndex) = "EXPRESSION" And _
            pavCloneRegister(2, iIndex) = .Fields("lostFocusExprID").value Then
            lngLostFocusExprID = pavCloneRegister(3, iIndex)
          End If
          If pavCloneRegister(1, iIndex) = "COLUMN" And _
            pavCloneRegister(2, iIndex) = .Fields("lookupFilterValueID").value Then
            lngLookupFilterValueID = pavCloneRegister(3, iIndex)
          End If
        
          
'MH20060613
          If (lngColumnID > 0) And _
            ((.Fields("calcExprID").value = 0) Or (lngCalcExprID > 0)) And _
            ((.Fields("gotFocusExprID").value = 0) Or (lngGotFocusExprID > 0)) And _
            ((.Fields("dfltValueExprID").value = 0) Or (lngDfltValueExprID > 0)) And _
            ((.Fields("lostFocusExprID").value = 0) Or (lngLostFocusExprID > 0)) And _
            ((.Fields("lookupFilterValueID").value = 0) Or (lngLookupFilterValueID > 0)) Then
            Exit For
          End If
        Next iIndex
            
        Set objNewColumn = New Column
        With objNewColumn
          .ColumnID = lngColumnID
          
          fOK = .ReadColumn
          
          If fOK Then
            .Properties("calcExprID").value = lngCalcExprID
            .Properties("gotFocusExprID").value = lngGotFocusExprID
            .Properties("dfltValueExprID").value = lngDfltValueExprID 'MH20060613
            .Properties("lostFocusExprID").value = lngLostFocusExprID
            .Properties("lookupFilterValueID").value = lngLookupFilterValueID
            fOK = .WriteColumn
          End If
        End With
        ' Disassociate object variables.
        Set objNewColumn = Nothing
              
        .MoveNext
      Loop
      
      .Close
    End With
  
    ' Disassociate object variables.
    Set rsColumns = Nothing
  End If
  
  If fOK Then
  
  'NPG20071119 Fault 12376
  
    ' Update the copied table's expression field components with the new IDs of their filter expressions.
    'sSQL = "SELECT tmpComponents.componentID, tmpComponents.fieldSelectionFilter" & _
      " FROM tmpComponents, tmpExpressions " & _
      " WHERE tmpExpressions.TableID = " & Trim(Str(objNewTable.TableID)) & _
      " AND tmpComponents.exprID = tmpExpressions.exprID" & _
      " AND tmpComponents.type = " & Trim(Str(giCOMPONENT_FIELD)) & _
      " AND tmpComponents.FieldPassBy = " & Trim(Str(giPASSBY_VALUE)) & _
      " AND tmpComponents.fieldSelectionFilter > 0"


    sSQL = "SELECT tmpComponents.componentID, tmpComponents.fieldSelectionFilter, tmpComponents.calculationID" & _
      " FROM tmpComponents, tmpExpressions " & _
      " WHERE tmpExpressions.TableID = " & Trim(Str(objNewTable.TableID)) & _
      " AND tmpComponents.exprID = tmpExpressions.exprID" & _
      " AND (" & _
          "(tmpComponents.type = " & Trim(Str(giCOMPONENT_FIELD)) & _
          " AND tmpComponents.FieldPassBy = " & Trim(Str(giPASSBY_VALUE)) & _
          " AND tmpComponents.fieldSelectionFilter > 0)" & _
        " OR " & _
          "(tmpComponents.type = " & Trim(Str(giCOMPONENT_CALCULATION)) & _
          " AND tmpComponents.calculationID > 0)" & _
      ")"

    Set rsFilters = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
            
    With rsFilters
      Do While (Not .EOF)
        
        For iIndex = 1 To UBound(pavCloneRegister, 2)
        
'NPG20071119 Fault 12376
'          If pavCloneRegister(1, iIndex) = "EXPRESSION" And _
'            pavCloneRegister(2, iIndex) = .Fields("fieldSelectionFilter") Then
'
'            recCompEdit.Index = "idxCompID"
'            recCompEdit.Seek "=", .Fields("componentID")
'            If Not recCompEdit.NoMatch Then
'              recCompEdit.Edit
'              recCompEdit.Fields("fieldSelectionFilter") = pavCloneRegister(3, iIndex)
'              recCompEdit.Update
'            End If
'            Exit For
'          End If
          If pavCloneRegister(1, iIndex) = "EXPRESSION" Then
            If pavCloneRegister(2, iIndex) = .Fields("fieldSelectionFilter") Then
              recCompEdit.Index = "idxCompID"
              recCompEdit.Seek "=", .Fields("componentID")
              If Not recCompEdit.NoMatch Then
                recCompEdit.Edit
                recCompEdit.Fields("fieldSelectionFilter") = pavCloneRegister(3, iIndex)
                recCompEdit.Update
              End If
            End If
          
            If pavCloneRegister(2, iIndex) = .Fields("calculationID") Then
              recCompEdit.Index = "idxCompID"
              recCompEdit.Seek "=", .Fields("componentID")
              If Not recCompEdit.NoMatch Then
                recCompEdit.Edit
                recCompEdit.Fields("calculationID") = pavCloneRegister(3, iIndex)
                recCompEdit.Update
              End If
            End If
          
          End If
        
        Next iIndex
              
        .MoveNext
      Loop
      
      .Close
    End With
  
    ' Disassociate object variables.
    Set rsFilters = Nothing
  End If
  
  
  '
  ' Copy Screens.
  '
  If fOK Then
  
    ' Get the table's screen definitions from the database.
    sSQL = "SELECT *" & _
      " FROM tmpScreens" & _
      " WHERE tmpScreens.deleted = FALSE" & _
      " AND tmpScreens.tableID = " & Trim(Str(gLngTableID))
    Set rsScreens = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
        
    With rsScreens
      ' For each screen definition ...
      Do While (Not .EOF)
      
        ' Get the copied screen's new order ID.
        lngOrderId = 0
        If .Fields("OrderID") > 0 Then
          For iIndex = 1 To UBound(pavCloneRegister, 2)
            If pavCloneRegister(1, iIndex) = "ORDER" And _
              pavCloneRegister(2, iIndex) = .Fields("OrderID") Then
              lngOrderId = pavCloneRegister(3, iIndex)
              Exit For
            End If
          Next iIndex
        End If
        
        ' Create a new screen name.
        sScreenName = "Copy_of_" & .Fields("Name")
        ' Check that the screen name is not already used.
        iCounter = 1
        fGoodName = False
        Do While Not fGoodName
          With recScrEdit
            .Index = "idxName"
            .Seek "=", sScreenName, False
            If Not .NoMatch Then
              iCounter = iCounter + 1
              sScreenName = "Copy_" & Trim(Str(iCounter)) & "_of_" & rsScreens.Fields("Name")
            Else
              fGoodName = True
            End If
          End With
        Loop
        ' Get a unique ID for the new record.
        lngScreenID = UniqueColumnValue("tmpScreens", "screenID")
        
        ' Add a new record in the database for the copied screen definition.
        recScrEdit.AddNew
        
        recScrEdit!ScreenID = lngScreenID
        recScrEdit!Changed = False
        recScrEdit!New = True
        recScrEdit!Deleted = False
        recScrEdit!Name = sScreenName
        recScrEdit!TableID = objNewTable.TableID
        recScrEdit!OrderID = lngOrderId
        recScrEdit!Height = .Fields("Height")
        recScrEdit!Width = .Fields("Width")
        recScrEdit!PictureID = .Fields("PictureID")
        recScrEdit!FontName = .Fields("FontName")
        recScrEdit!FontSize = .Fields("FontSize")
        recScrEdit!FontBold = .Fields("FontBold")
        recScrEdit!FontItalic = .Fields("FontItalic")
        recScrEdit!FontStrikethru = .Fields("FontStrikeThru")
        recScrEdit!FontUnderline = .Fields("FontUnderline")
        recScrEdit!GridX = .Fields("GridX")
        recScrEdit!GridY = .Fields("GridY")
        recScrEdit!AlignToGrid = .Fields("AlignToGrid")
        recScrEdit!dfltForeColour = .Fields("DfltForeColour")
        recScrEdit!dfltFontName = .Fields("DfltFontName")
        recScrEdit!dfltFontSize = .Fields("DfltFontSize")
        recScrEdit!dfltFontBold = .Fields("DfltFontBold")
        recScrEdit!dfltFontItalic = .Fields("DfltFontItalic")
        
        ' JPD - Fault 10524
        recScrEdit!QuickEntry = .Fields("quickEntry")
        recScrEdit!SSIntranet = .Fields("ssIntranet")
      
        recScrEdit.Update
      
        ' Remember the IDs of the original and copied screens.
        iIndex = UBound(pavCloneRegister, 2) + 1
        ReDim Preserve pavCloneRegister(3, iIndex)
        pavCloneRegister(1, iIndex) = "SCREEN"
        pavCloneRegister(2, iIndex) = .Fields("ScreenID")
        pavCloneRegister(3, iIndex) = lngScreenID
      
        ' Copy the screen control definitions.
        sSQL = "SELECT *" & _
          " FROM tmpControls" & _
          " WHERE tmpControls.screenID = " & Trim(Str(.Fields("ScreenID")))
        Set rsControls = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
            
        With rsControls
          ' For each screen control definition ...
          Do While (Not .EOF)
            ' Add a new record in the database for the copied screen control definition.
            recCtrlEdit.AddNew
            
            ' Get the copied screen control's new column ID.
            lngColumnID = 0
            If .Fields("ColumnID") > 0 Then
              For iIndex = 1 To UBound(pavCloneRegister, 2)
                If pavCloneRegister(1, iIndex) = "COLUMN" And _
                  pavCloneRegister(2, iIndex) = .Fields("ColumnID") Then
                  lngColumnID = pavCloneRegister(3, iIndex)
                  Exit For
                End If
              Next iIndex
            End If
            
'            'NPG20080507 Fault 12974
'            ' Parent table ID's are to be retained.
'            lngTableID = IIf(IsNull(.Fields("TableID")), vbNull, .Fields("TableID"))
'
'            If .Fields("TableID") = TableID Then ' Not a parent table ID, so assign new tableID
'              lngTableID = objNewTable.TableID
'            End If
            
            ' Get the copied screen control's new column ID.
            lngTableID = IIf(IsNull(.Fields("TableID")), 0, .Fields("TableID"))
            If lngTableID > 0 Then
              For iIndex = 1 To UBound(pavCloneRegister, 2)
                If pavCloneRegister(1, iIndex) = "TABLE" And _
                  pavCloneRegister(2, iIndex) = .Fields("TableID") Then
                  lngTableID = pavCloneRegister(3, iIndex)
                  Exit For
                End If
              Next iIndex
            End If
            
            
            recCtrlEdit!ScreenID = lngScreenID
            recCtrlEdit!PageNo = .Fields("PageNo")
            recCtrlEdit!ControlLevel = .Fields("ControlLevel")
            'NPG20080507 Fault 12974
            'recCtrlEdit!TableID = objNewTable.TableID
            If lngTableID > 0 Then
              recCtrlEdit!TableID = lngTableID
            End If
            recCtrlEdit!ColumnID = lngColumnID
            recCtrlEdit!ControlType = .Fields("ControlType")
            recCtrlEdit!ControlIndex = .Fields("ControlIndex")
            recCtrlEdit!TopCoord = .Fields("TopCoord")
            recCtrlEdit!LeftCoord = .Fields("LeftCoord")
            recCtrlEdit!Height = .Fields("Height")
            recCtrlEdit!Width = .Fields("Width")
            recCtrlEdit!Caption = .Fields("Caption")
            recCtrlEdit!BackColor = .Fields("BackColor")
            recCtrlEdit!ForeColor = .Fields("ForeColor")
            recCtrlEdit!FontName = .Fields("FontName")
            recCtrlEdit!FontSize = .Fields("FontSize")
            recCtrlEdit!FontBold = .Fields("FontBold")
            recCtrlEdit!FontItalic = .Fields("FontItalic")
            recCtrlEdit!FontStrikethru = .Fields("FontStrikeThru")
            recCtrlEdit!FontUnderline = .Fields("FontUnderline")
            recCtrlEdit!PictureID = .Fields("PictureID")
            recCtrlEdit!DisplayType = .Fields("DisplayType")
            recCtrlEdit!TabIndex = .Fields("TabIndex")
            recCtrlEdit!BorderStyle = .Fields("BorderStyle")
            recCtrlEdit!Alignment = .Fields("Alignment")
            recCtrlEdit!ReadOnly = .Fields("ReadOnly")      'NPG20071023
          
            recCtrlEdit.Update
            
            .MoveNext
          Loop
        End With
        ' Disassociate object variables.
        Set rsControls = Nothing
        
        ' Copy the screen page caption definitions.
        sSQL = "SELECT *" & _
          " FROM tmpPageCaptions" & _
          " WHERE tmpPageCaptions.screenID = " & Trim(Str(.Fields("ScreenID")))
        Set rsPageCaptions = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
            
        With rsPageCaptions
          ' For each screen page caption definition ...
          Do While (Not .EOF)
            ' Add a new record in the database for the copied screen page caption definition.
            recPageCaptEdit.AddNew
            
            recPageCaptEdit!ScreenID = lngScreenID
            recPageCaptEdit!PageIndexID = .Fields("PageIndexID")
            recPageCaptEdit!Caption = .Fields("Caption")
          
            recPageCaptEdit.Update
            
            .MoveNext
          Loop
        End With
        ' Disassociate object variables.
        Set rsPageCaptions = Nothing
            
        ' Copy the history screen records.
        sSQL = "SELECT *" & _
          " FROM tmpHistoryScreens" & _
          " WHERE tmpHistoryScreens.parentScreenID = " & Trim(Str(.Fields("ScreenID")))
        Set rsHistoryScreens = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
            
        With rsHistoryScreens
          ' For each history screen definition ...
          Do While (Not .EOF)
            ' Get the ID of the copy of the child table.
            lngHistoryScreenID = 0
            For iIndex = 1 To UBound(pavCloneRegister, 2)
              If pavCloneRegister(1, iIndex) = "SCREEN" And _
                pavCloneRegister(2, iIndex) = .Fields("HistoryScreenID") Then
                lngHistoryScreenID = pavCloneRegister(3, iIndex)
                Exit For
              End If
            Next iIndex
            
            ' Add a new record in the database for the copied history screen definition.
            recHistScrEdit.AddNew
            
            recHistScrEdit!ID = UniqueColumnValue("tmpHistoryScreens", "ID")
            recHistScrEdit!parentScreenID = lngScreenID
            recHistScrEdit!historyScreenID = lngHistoryScreenID
          
            recHistScrEdit.Update
            
            .MoveNext
          Loop
        End With
        ' Disassociate object variables.
        Set rsHistoryScreens = Nothing
              
        .MoveNext
      Loop
              
      .Close
    End With
    ' Disassociate object variables.
    Set rsScreens = Nothing
    
  End If
  
  '
  ' Copy Views.
  '
  If fOK Then
  
    ' Get the table's view definitions from the database.
    sSQL = "SELECT *" & _
      " FROM tmpViews" & _
      " WHERE tmpViews.deleted = FALSE" & _
      " AND tmpViews.viewTableID = " & Trim(Str(gLngTableID))
    Set rsViews = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
        
    With rsViews
      ' For each view definition ...
      Do While (Not .EOF)
        ' Get the copied view's 'where clause' expression ID.
        lngExpressionID = 0
        If .Fields("ExpressionID") > 0 Then
          For iIndex = 1 To UBound(pavCloneRegister, 2)
            If pavCloneRegister(1, iIndex) = "EXPRESSION" And _
              pavCloneRegister(2, iIndex) = .Fields("ExpressionID") Then
              lngExpressionID = pavCloneRegister(3, iIndex)
              Exit For
            End If
          Next iIndex
        End If

        ' Create a new view name.
        sViewName = "Copy_of_" & .Fields("ViewName")
        ' Check that the view name is not already used.
        iCounter = 1
        fGoodName = False
        Do While Not fGoodName
          With recViewEdit
            .Index = "idxViewName"
            .Seek "=", sViewName, False
            If Not .NoMatch Then
              iCounter = iCounter + 1
              sViewName = "Copy_" & Trim(Str(iCounter)) & "_of_" & rsViews.Fields("ViewName")
            Else
              fGoodName = True
            End If
          End With
        Loop
        ' Get a unique ID for the new record.
        lngViewID = UniqueColumnValue("tmpViews", "ViewID")
        
        ' Add a new record in the database for the copied view definition.
        recViewEdit.AddNew
        
        recViewEdit!ViewID = lngViewID
        'NPG20080207 Fault 12874
        ' recViewEdit!OriginalViewName = sViewName
        recViewEdit!OriginalViewName = .Fields("OriginalViewName")
        
        recViewEdit!ViewName = sViewName
        recViewEdit!ViewDescription = .Fields("ViewDescription")
        recViewEdit!ViewTableID = objNewTable.TableID
        recViewEdit!ViewSQL = .Fields("ViewSQL")
        
        ' MH20020808 Remove reference to "viewAlternativeName"
        'recViewEdit!ViewAlternativeName = .Fields("ViewAlternativeName")
        recViewEdit!ExpressionID = lngExpressionID
        recViewEdit!Changed = False
        recViewEdit!New = True
        recViewEdit!Deleted = False
        
        ' Copy existing security permissions
        If pfCopySecurity Then
          recViewEdit!GrantRead = .Fields("GrantRead")
          recViewEdit!GrantNew = .Fields("GrantNew")
          recViewEdit!GrantEdit = .Fields("GrantEdit")
          recViewEdit!GrantDelete = .Fields("GrantDelete")
        Else
          recViewEdit!GrantRead = bGrantRead
          recViewEdit!GrantNew = bGrantNew
          recViewEdit!GrantEdit = bGrantEdit
          recViewEdit!GrantDelete = bGrantDelete
        End If
        
        recViewEdit.Update
      
        ' Remember the IDs of the original and copied screens.
        iIndex = UBound(pavCloneRegister, 2) + 1
        ReDim Preserve pavCloneRegister(3, iIndex)
        pavCloneRegister(1, iIndex) = "VIEW"
        pavCloneRegister(2, iIndex) = .Fields("ViewID")
        pavCloneRegister(3, iIndex) = lngViewID
      
        ' Copy the view screen definitions.
        sSQL = "SELECT *" & _
          " FROM tmpViewScreens" & _
          " WHERE tmpViewScreens.viewID = " & Trim(Str(.Fields("ViewID")))
        Set rsViewScreens = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
            
        With rsViewScreens
          ' For each view screen definition ...
          Do While (Not .EOF)
            ' Get the copied screen control's new column ID.
            lngScreenID = 0
            If .Fields("ScreenID") > 0 Then
              For iIndex = 1 To UBound(pavCloneRegister, 2)
                If pavCloneRegister(1, iIndex) = "SCREEN" And _
                  pavCloneRegister(2, iIndex) = .Fields("ScreenID") Then
                  lngScreenID = pavCloneRegister(3, iIndex)
                  Exit For
                End If
              Next iIndex
            End If
            
            ' Add a new record in the database for the copied view screen definition.
            recViewScreens.AddNew
            
            recViewScreens!ScreenID = lngScreenID
            recViewScreens!ViewID = lngViewID
            recViewScreens!New = True
            recViewScreens!Deleted = False
          
            recViewScreens.Update
            
            .MoveNext
          Loop
        End With
        ' Disassociate object variables.
        Set rsViewScreens = Nothing
        
        ' Copy the view column definitions.
        sSQL = "SELECT tmpViewColumns.*" & _
          " FROM tmpViewColumns, tmpColumns" & _
          " WHERE tmpViewColumns.viewID = " & Trim(Str(.Fields("ViewID"))) & _
          " AND tmpViewColumns.columnID = tmpColumns.columnID" '& _
          " AND tmpColumns.columnType <> " & Trim(Str(giCOLUMNTYPE_SYSTEM)) & _
          " AND tmpColumns.columnType <> " & Trim(Str(giCOLUMNTYPE_LINK))

        Set rsViewColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
            
        With rsViewColumns
          ' For each View Column definition ...
          Do While (Not .EOF)
            ' Get the copied screen control's new column ID.
            lngColumnID = 0
            If .Fields("ColumnID") > 0 Then
              For iIndex = 1 To UBound(pavCloneRegister, 2)
                If pavCloneRegister(1, iIndex) = "COLUMN" And _
                  pavCloneRegister(2, iIndex) = .Fields("ColumnID") Then
                  lngColumnID = pavCloneRegister(3, iIndex)
                  Exit For
                End If
              Next iIndex
            End If
            
            ' Add a new record in the database for the copied view column definition.
            recViewColEdit.AddNew
            
            recViewColEdit!ViewID = lngViewID
            recViewColEdit!ColumnID = lngColumnID
            recViewColEdit!InView = .Fields("InView")
            recViewColEdit!New = True
            recViewColEdit!Changed = False
            recViewColEdit!Deleted = False
          
            recViewColEdit.Update
            
            .MoveNext
          Loop
        End With
        ' Disassociate object variables.
        Set rsViewColumns = Nothing
              
        .MoveNext
      Loop
              
      .Close
    End With
    ' Disassociate object variables.
    Set rsScreens = Nothing
    
  End If
  
  '
  ' Copy Table Based Workflows. ie. ones that can be used in triggered links
  '
  If fOK Then
    ' Get the table's workflow definitions from the database.
    sSQL = "SELECT *" & _
      " FROM tmpWorkflows" & _
      " WHERE tmpWorkflows.deleted = FALSE" & _
      " AND tmpWorkflows.initiationType = " & Trim(Str(WORKFLOWINITIATIONTYPE_TRIGGERED)) & _
      " AND tmpWorkflows.baseTable = " & Trim(Str(gLngTableID))
    Set rsWorkflows = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    With rsWorkflows
      ' For each workflow definition ...
      Do While (Not .EOF) And fOK
        fOK = CloneWorkflow(.Fields("ID"), pavCloneRegister)
        
        .MoveNext
      Loop

      .Close
    End With
    ' Disassociate object variables.
    Set rsWorkflows = Nothing
  End If
  
  If fOK Then
    ' Get the table's Outlook definitions from the database.
    sSQL = "SELECT tmpOutlookFolders.FolderID" & _
      " FROM tmpOutlookFolders" & _
      " WHERE tmpOutlookFolders.deleted = FALSE" & _
      " AND tmpOutlookFolders.TableID = " & CStr(gLngTableID)

    Set rsOutlookDef = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    With rsOutlookDef
      ' For each Outlook definition ...
      Do While (Not .EOF) And fOK
        
        'With recOutlookLinks
        '  .Index = "idxTableID"
        '  .Seek "=", gLngTableID

        '  If Not .NoMatch Then

            ' Instantiate the original Outlook object.
        Set objSourceOutlookFolder = New clsOutlookFolder

        objSourceOutlookFolder.FolderID = .Fields("FolderID")

        Set objNewOutlookFolder = objSourceOutlookFolder.CloneOutlookFolder(pavCloneRegister)
        fOK = Not objNewOutlookFolder Is Nothing

        If fOK Then
          ' Copy properties from the original Outlook to the copy.
          objNewOutlookFolder.TableID = objNewTable.TableID
          ' Write the copied Outlook definition to the database.
          fOK = objNewOutlookFolder.WriteFolder
        End If

        ' Remember the IDs of the original and copied Outlooks.
        'If fOK Then
        iIndex = UBound(pavCloneRegister, 2) + 1
        ReDim Preserve pavCloneRegister(3, iIndex)
        pavCloneRegister(1, iIndex) = "OUTLOOKFOLDER"
        pavCloneRegister(2, iIndex) = objSourceOutlookFolder.FolderID

        If fOK Then
          pavCloneRegister(3, iIndex) = objNewOutlookFolder.FolderID
        Else
          pavCloneRegister(3, iIndex) = 0

          'mstrOutlookNotCopied = mstrOutlookNotCopied & objSourceOutlook.OutlookName & vbCrLf
          fOK = True
        End If

        ' Disassociate object variables.
        Set objSourceOutlookFolder = Nothing
        Set objNewOutlookFolder = Nothing
        
        .MoveNext
      Loop
          
      .Close
    End With
    ' Disassociate object variables.
    Set rsOutlookDef = Nothing

  End If

  If fOK Then
    ' Get the table's Outlook definitions from the database.
    sSQL = "SELECT tmpOutlookLinks.LinkID" & _
      " FROM tmpOutlookLinks" & _
      " WHERE tmpOutlookLinks.deleted = FALSE" & _
      " AND tmpOutlookLinks.TableID = " & CStr(gLngTableID)

    Set rsOutlookDef = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    With rsOutlookDef
      ' For each Outlook definition ...
      Do While (Not .EOF) And fOK
        ' Instantiate the original Outlook object.
        Set objSourceOutlookLink = New clsOutlookLink
        objSourceOutlookLink.LinkID = .Fields("LinkID")

        Set objNewOutlookLink = objSourceOutlookLink.CloneOutlookLink(pavCloneRegister)
        fOK = Not objNewOutlookLink Is Nothing
        
        If fOK Then
          ' Copy properties from the original Outlook to the copy.
          objNewOutlookLink.TableID = objNewTable.TableID
          ' Write the copied Outlook definition to the database.
          fOK = objNewOutlookLink.WriteLink
        End If
        
        iIndex = UBound(pavCloneRegister, 2) + 1
        ReDim Preserve pavCloneRegister(3, iIndex)
        pavCloneRegister(1, iIndex) = "OUTLOOKLINK"
        pavCloneRegister(2, iIndex) = objSourceOutlookLink.LinkID
        
        If fOK Then
          pavCloneRegister(3, iIndex) = objNewOutlookLink.LinkID
        Else
          pavCloneRegister(3, iIndex) = 0
        
          'mstrOutlookNotCopied = mstrOutlookNotCopied & objSourceOutlook.OutlookName & vbCrLf
          'fOK = True
        End If
      
        ' Disassociate object variables.
        Set objSourceOutlookLink = Nothing
        Set objNewOutlookLink = Nothing
        
        .MoveNext
      Loop
          
      .Close
    End With
    ' Disassociate object variables.
    Set rsOutlookDef = Nothing

    objNewTable.ReadOutlookLinks
  
  End If

  If fOK Then
    ' Get the table's Workflow link definitions from the database.
    sSQL = "SELECT tmpWorkflowTriggeredLinks.LinkID" & _
      " FROM tmpWorkflowTriggeredLinks" & _
      " WHERE tmpWorkflowTriggeredLinks.deleted = FALSE" & _
      " AND tmpWorkflowTriggeredLinks.TableID = " & CStr(gLngTableID)

    Set rsWorkflowTriggeredLinkDef = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    With rsWorkflowTriggeredLinkDef
      ' For each Workflow Triggered Link definition ...
      Do While (Not .EOF) And fOK
        Set objSourceWorkflowTriggeredLink = New clsWorkflowTriggeredLink
        objSourceWorkflowTriggeredLink.LinkID = .Fields("LinkID")

        Set objNewWorkflowTriggeredLink = objSourceWorkflowTriggeredLink.CloneLink(pavCloneRegister)
        fOK = Not objNewWorkflowTriggeredLink Is Nothing

        If fOK Then
          ' Write the copied Workflow definition to the database.
          fOK = objNewWorkflowTriggeredLink.WriteLink
        End If

        iIndex = UBound(pavCloneRegister, 2) + 1
        ReDim Preserve pavCloneRegister(3, iIndex)
        pavCloneRegister(1, iIndex) = "WORKFLOWLINK"
        pavCloneRegister(2, iIndex) = objSourceWorkflowTriggeredLink.LinkID
        
        If fOK Then
          pavCloneRegister(3, iIndex) = objNewWorkflowTriggeredLink.LinkID
        Else
          pavCloneRegister(3, iIndex) = 0
        End If

        ' Disassociate object variables.
        Set objSourceWorkflowTriggeredLink = Nothing
        Set objNewWorkflowTriggeredLink = Nothing

        .MoveNext
      Loop

      .Close
    End With
    ' Disassociate object variables.
    Set rsWorkflowTriggeredLinkDef = Nothing

    objNewTable.ReadWorkflowLinks
    objNewTable.ReadTableValidations

  End If


TidyUpAndExit:
  If fOK Then
    Set CloneTable = objNewTable
  Else
    Set CloneTable = Nothing
  End If
  
  ' Disassociate object variables.
  Set rsViews = Nothing
  Set rsOrders = Nothing
  Set rsScreens = Nothing
  Set rsColumns = Nothing
  Set rsFilters = Nothing
  Set rsChildren = Nothing
  Set rsControls = Nothing
  Set rsViewScreens = Nothing
  Set rsViewColumns = Nothing
  Set rsExpressions = Nothing
  Set rsPageCaptions = Nothing
  Set rsHistoryScreens = Nothing
  Set objNewTable = Nothing
  Set objChildTable = Nothing
  Set objCopyChildTable = Nothing
  Set objNewOrder = Nothing
  Set objSourceOrder = Nothing
  Set objNewColumn = Nothing
  Set objSourceColumn = Nothing
  Set objNewExpr = Nothing
  Set objSourceExpr = Nothing
  
  Exit Function

ErrorTrap:
  gobjProgress.Visible = False
  MsgBox "Unable to copy the table." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  gobjProgress.Visible = True
  fOK = False
Resume Next
  Resume TidyUpAndExit

End Function


Public Function FindTable() As Boolean
  On Error GoTo ErrorTrap
  
  With recTabEdit
    .Index = "idxTableID"
    .Seek "=", TableID
    FindTable = (Not .NoMatch)
  End With
  
  Exit Function
  
ErrorTrap:
  FindTable = False

End Function

Public Function NewTable() As Boolean
  ' Create a new table.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim frmEdit As frmTabEdit
  
  gfNew = True
  gfChanged = False
  gfDeleted = False
  gLngTableID = 0
  
  Set frmEdit = New frmTabEdit
  Set frmEdit.Table = Me
  frmEdit.Show vbModal
  
  fOK = Not frmEdit.Cancelled
  
  If fOK Then
    fOK = WriteTable_Transaction
    
    If Not fOK Then
      gfNew = True
      gfChanged = False
      gfDeleted = False
      gLngTableID = 0
    End If
  End If

TidyUpAndExit:
  Set frmEdit = Nothing
  
  NewTable = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function ReadTable() As Boolean
  ' Read the table's details from the local tables.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sIndex As String
  
  With recTabEdit
    'Store current index
    sIndex = .Index
    
    'Find the table record
    fOK = FindTable
    
    If fOK Then
    
      'Set table object properties
      gsTableName = !TableName
      giTableType = !TableType
      
      mbLocked = !Locked
      
      If Not IsNull(!LastUpdated) Then
        gDatUpdated = !LastUpdated
      End If
      
      gfNew = !New
      gfChanged = !Changed
      gfDeleted = !Deleted
      
      miGrantRead = !GrantRead
      miGrantNew = !GrantNew
      miGrantEdit = !GrantEdit
      miGrantDelete = !GrantDelete
      
      'JPD 20030829 Fault 5538
      mbIsCopy = !IsCopy
      mlngCopyTableID = !copySecurityTableID
      msCopyTableName = !copySecurityTableName

      mbPermissionsPrompted = !PermissionsPrompted
      mbManualColumnBreaks = !ManualSummaryColumnBreaks
      mbAuditInsert = IIf(IsNull(!AuditInsert), False, !AuditInsert)
      mbAuditDelete = IIf(IsNull(!AuditDelete), False, !AuditDelete)
      'mlngEmailInsert = IIf(IsNull(!EmailInsert), False, !EmailInsert)
      'mlngEmailDelete = IIf(IsNull(!EmailDelete), False, !EmailDelete)
      
      If Not IsNull(!defaultOrderID) Then
        gLngPrimaryOrderID = !defaultOrderID
      Else
        gLngPrimaryOrderID = 0
      End If
      
      If Not IsNull(!RecordDescExprID) Then
        gLngRecordDescriptionID = !RecordDescExprID
      Else
        gLngRecordDescriptionID = 0
      End If
      
      'MH20000727
      If Not IsNull(!DefaultEmailID) Then
        glngPrimaryEmailID = !DefaultEmailID
      Else
        glngPrimaryEmailID = 0
      End If
      
      ' Read the summary field definition.
      fOK = ReadSummaryFields
    
      
      If fOK Then
        fOK = ReadEmailLinks
      End If
      
      'MH20040322
      If fOK Then
        fOK = ReadOutlookLinks
      End If
    
      If fOK Then
        fOK = ReadWorkflowLinks
      End If
      
      If fOK Then
        fOK = ReadTableValidations
      End If
      
      If fOK Then
        fOK = ReadTableTriggers
      End If
      
      InsertTriggerDisabled = !InsertTriggerDisabled
      UpdateTriggerDisabled = !UpdateTriggerDisabled
      DeleteTriggerDisabled = !DeleteTriggerDisabled
      
    End If
  End With
  
TidyUpAndExit:
  'Restore original index
  recTabEdit.Index = sIndex
  ReadTable = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function ReadSummaryFields() As Boolean
  ' Read the table's summary fields from the database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objSummaryField As cSummaryField
  
  fOK = ClearSummaryFields
  
  With recSummaryEdit
    .Index = "idxRealOrder"
    .Seek ">=", TableID
    
    If Not .NoMatch Then
      Do While Not .EOF
    
        If !HistoryTableID <> TableID Then
          Exit Do
        End If
        
        Set objSummaryField = New cSummaryField
        objSummaryField.ID = !ID
        If objSummaryField.ReadSummaryField Then
          gobjSummaryFields.Add objSummaryField
        End If
    
        .Index = "idxRealOrder"
        .Seek ">=", TableID, objSummaryField.Sequence, objSummaryField.SummaryColumnID
        
        Set objSummaryField = Nothing
        
        .MoveNext
      Loop
    End If
  End With
  
TidyUpAndExit:
  ReadSummaryFields = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
    
End Function
Public Property Get SummaryFields() As Collection
  ' Return the collection of Summary Fields.
  Set SummaryFields = gobjSummaryFields
  
End Property
Public Property Set SummaryFields(pobjNewCollection As Collection)
  ' Set the collection of diary links.
  Set gobjSummaryFields = pobjNewCollection
  
End Property


Public Function ClearSummaryFields() As Boolean
  ' Clear the table's summary fields collection.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  ' Remove all summary fields objects from the collection.
  Do While gobjSummaryFields.Count > 0
    gobjSummaryFields.Remove 1
  Loop
  Set gobjSummaryFields = Nothing
  
  ' Re-instantiate the collection.
  Set gobjSummaryFields = New Collection
  
TidyUpAndExit:
  ClearSummaryFields = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Public Function WriteTable() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fNewTable As Boolean
  Dim sIndex As String
  Dim sSQL As String
  Dim rsViews As DAO.Recordset
  
  With recTabEdit
    'Store current index
    sIndex = .Index
    
    'Check if a new table is being written
    fNewTable = (IsNew And Not IsChanged)
    
    If fNewTable Then
      'Get new table id
'      TableID = UniqueColumnValue("tmpTables", "tableID")
      TableID = GetNextObjectIdentitySeed("ASRSysTables")
      
      'Add a record for new table
      fOK = (TableID > 0)
      If fOK Then
        .AddNew
        !TableID = TableID
      End If
    Else
      'Find existing table record
      fOK = FindTable
      If fOK Then
        'Edit table record
        .Edit
      End If
    End If
    
    ' Write table record.
    If fOK Then
      gDatUpdated = Now()
      gfChanged = True
      !TableName = TableName
      
      'MH20010329
      If IsNull(!OriginalTableName) Then
        !OriginalTableName = TableName
      End If
      
      !TableType = TableType
      !LastUpdated = LastUpdated
      !New = IsNew
      !Changed = IsChanged
      !Deleted = IsDeleted
      !defaultOrderID = PrimaryOrderID
      !RecordDescExprID = RecordDescriptionID
      !DefaultEmailID = PrimaryEmailID
      !GrantRead = miGrantRead
      !GrantEdit = miGrantEdit
      !GrantNew = miGrantNew
      !GrantDelete = miGrantDelete
      !IsCopy = mbIsCopy
      !copySecurityTableID = IIf(IsNull(CopyTableID), 0, CopyTableID)
      !copySecurityTableName = IIf(IsNull(CopyTableName), 0, CopyTableName)
      !PermissionsPrompted = mbPermissionsPrompted
      !ManualSummaryColumnBreaks = mbManualColumnBreaks
      !AuditInsert = mbAuditInsert
      !AuditDelete = mbAuditDelete
      '!EmailInsert = mlngEmailInsert
      '!EmailDelete = mlngEmailDelete
      !InsertTriggerDisabled = InsertTriggerDisabled
      !UpdateTriggerDisabled = UpdateTriggerDisabled
      !DeleteTriggerDisabled = DeleteTriggerDisabled

      .Update
      
      'If this is a new table add an ID column
      If fNewTable Then
        fOK = AddIDColumn(TableID)
      End If
    
      If fOK Then
        fOK = WriteSummaryFields
      End If
      
      ' JPD 13/1/00 Do not allow non-parent tables to have views defined.
      If fOK Then
        If TableType <> iTabParent Then
          ' Mark the view records as deleted.
          sSQL = "SELECT tmpViews.viewID" & _
            " FROM tmpViews" & _
            " WHERE viewTableID = " & Trim(Str(TableID))
        
          Set rsViews = daoDb.OpenRecordset(sSQL, _
            dbOpenForwardOnly, dbReadOnly)
              
          With rsViews
            Do While (Not .EOF)
              ' Mark the view screen records as deleted.
              sSQL = "UPDATE tmpViewScreens" & _
                " SET deleted = TRUE" & _
                " WHERE tmpViewScreens.viewID = " & Trim(Str(.Fields("viewID")))
              daoDb.Execute sSQL, dbFailOnError
                  
              ' Mark the view column records as deleted.
              sSQL = "UPDATE tmpViewColumns" & _
                " SET deleted = TRUE" & _
                " WHERE tmpViewColumns.viewID = " & Trim(Str(.Fields("viewID")))
              daoDb.Execute sSQL, dbFailOnError
              
              .MoveNext
            Loop
                
            .Close
          End With
                    
          ' Mark the view master records as deleted.
          sSQL = "UPDATE tmpViews" & _
            " SET deleted = TRUE" & _
            " WHERE viewTableID = " & Trim(Str(TableID))
          daoDb.Execute sSQL, dbFailOnError
        End If
      End If

    End If
  End With

  WriteEmailLinks
  WriteOutlookLinks
  WriteWorkflowLinks
  WriteTableValidations
  WriteTableTriggers

TidyUpAndExit:
  ' Restore original index.
  recTabEdit.Index = sIndex
  WriteTable = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Public Function WriteTable_Transaction() As Boolean
  ' Transaction wrapper for the 'WriteTable' function.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  fOK = WriteTable
  
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  WriteTable_Transaction = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function WriteSummaryFields() As Boolean
  ' Write the Summary Field values to the database.
  On Error GoTo ErrorTrap
  
  Dim fWriteOk As Boolean
  Dim objSummaryField As cSummaryField
  
  fWriteOk = True
  
  ' Delete all existing control values for the current column.
  daoDb.Execute "DELETE FROM tmpSummary WHERE historyTableID=" & TableID, dbFailOnError
  
  ' Add records for each Summary Field in the collection.
  For Each objSummaryField In gobjSummaryFields
    If fWriteOk Then
      fWriteOk = objSummaryField.WriteSummaryField
    End If
  Next objSummaryField
  Set objSummaryField = Nothing

TidyUpAndExit:
  'Check if column was written successfully
  WriteSummaryFields = fWriteOk
  Exit Function

ErrorTrap:
  fWriteOk = False
  Err = False
  Resume TidyUpAndExit

End Function



Public Function AddIDColumn(pLngParentTableID As Long) As Boolean
  On Error GoTo ErrorTrap
  
  Dim IDColumn As New SystemMgr.Column
  Dim fAddOk As Boolean
  
  fAddOk = False
  
  'Add ID column definition for new table if it doesn't already exist.
  With IDColumn
    .TableID = TableID
    .InitColumnProps
    .Properties("new").value = True
    .Properties("columnName").value = IIf(pLngParentTableID = TableID, "ID", "ID_" & pLngParentTableID)
    .Properties("dataType").value = dtINTEGER
    .Properties("size").value = 0
    .Properties("decimals").value = 0
    .Properties("columnType").value = giCOLUMNTYPE_SYSTEM
    .Properties("Afdenabled").value = False
    .Properties("Afdindividual").value = False
    .Properties("Afdforename").value = 0
    .Properties("Afdsurname").value = 0
    .Properties("Afdinitial").value = 0
    .Properties("Afdtelephone").value = 0
    .Properties("Afdaddress").value = 0
    .Properties("Afdproperty").value = 0
    .Properties("Afdstreet").value = 0
    .Properties("Afdlocality").value = 0
    .Properties("Afdtown").value = 0
    .Properties("Afdcounty").value = 0
    .Properties("QAddressEnabled").value = False
    .Properties("QAindividual").value = False
    .Properties("QAaddress").value = 0
    .Properties("QAproperty").value = 0
    .Properties("QAstreet").value = 0
    .Properties("QAlocality").value = 0
    .Properties("QAtown").value = 0
    .Properties("QAcounty").value = 0
    
    recColEdit.Index = "idxName"
    recColEdit.Seek "=", TableID, .Properties("columnName").value, False
    
    If recColEdit.NoMatch Then
      fAddOk = .WriteColumn
    Else
      fAddOk = True
    End If
  
  End With
  
ExitAddIDColumn:
  Set IDColumn = Nothing
  AddIDColumn = fAddOk
  
  Exit Function
  
ErrorTrap:
  fAddOk = False
  Err = False
  
  Resume ExitAddIDColumn

End Function

Public Function HasRelations() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fHasRelations As Boolean
  
  fHasRelations = False
  
  With recRelEdit
    .Index = "idxParentID"
    .Seek ">=", gLngTableID
    
    If Not .NoMatch Then
      If !parentID = gLngTableID Then
        fHasRelations = True
      End If
    End If
  End With
  
TidyUpAndExit:
  HasRelations = fHasRelations
  Exit Function
  
ErrorTrap:
  fHasRelations = False
  Resume TidyUpAndExit
  
End Function
Public Function IsDescendantOf(plngTableID As Long) As Boolean
  ' Returns TRUE if the current table object is a descendant
  ' of the given table id.
  Dim objParentTable As Table
  Dim varBookMark As Variant
  
  With recRelEdit
    .Index = "idxParentID"
    .MoveFirst
    .Seek "=", plngTableID, gLngTableID
    IsDescendantOf = Not .NoMatch
  End With
    
  If Not IsDescendantOf Then
    ' Check if the given table is a parent of any of the parents of the current table object.
    With recRelEdit
      .Index = "idxChildID"
      .MoveFirst
      .Seek "=", gLngTableID
    
      If Not .NoMatch Then
        Do While (Not .EOF) And (Not IsDescendantOf)
          If !childID <> gLngTableID Then
            Exit Do
          End If
        
          varBookMark = .Bookmark
          Set objParentTable = New Table
          objParentTable.TableID = !parentID
          IsDescendantOf = objParentTable.IsDescendantOf(plngTableID)
          Set objParentTable = Nothing
          .Index = "idxChildID"
          .Bookmark = varBookMark

          .MoveNext
        Loop
      End If
    End With
  End If
  
End Function

Public Function DeleteAllRelations() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsRelations As DAO.Recordset
  Dim lngTempTableID As Long
  
  fOK = True
  
  sSQL = "SELECT childID" & _
    " FROM tmpRelations" & _
    " WHERE parentID=" & Trim(Str(gLngTableID))
  Set rsRelations = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  
  With rsRelations
    Do While (Not .EOF) And fOK
      lngTempTableID = .Fields("childID")
      fOK = DeleteRelation(lngTempTableID)
      
      'TM20010724 Fault 1239
      If fOK Then
        'Set the !Changed field property to True for all Default expressions
        'in the child table. So that SaveChanges refreshes the 'DfltExpr' SP.
        fOK = DefaultExprChange(lngTempTableID)
      End If
      
      .MoveNext
    Loop
  End With
  
TidyUpAndExit:
  DeleteAllRelations = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Public Function DefaultExprChange(lngTableID As Long) As Boolean

  '******************************************************************************
  ' DefaultExprChange - Sets the Changed fields to true for all default         *
  '                     expressions in the passed tableID. This function is     *
  '                     called when the relation informaiton is upadted and the *
  '                     'DfltExpr' SPs need to be refreshed.                    *
  '                                                                             *
  'TM20010724 Fault 1239                                                        *
  '******************************************************************************
  
  On Error GoTo ErrorTrap
  
  Dim lngExprID As Long
  
  With recColEdit
    .MoveFirst
    Do While Not .EOF
    
      If !TableID = lngTableID And !dfltValueExprID > 0 And Not IsNull(!dfltValueExprID) Then
        lngExprID = !dfltValueExprID
        
        With recExprEdit
          .Index = "idxExprID"
          .Seek "=", lngExprID, False
          If Not .NoMatch Then
            .Edit
            !Changed = True
            .Update
          End If
        End With
        
      End If
      .MoveNext
      
    Loop
  End With
            
  DefaultExprChange = True
  Exit Function
  
ErrorTrap:
  DefaultExprChange = False
  
End Function

Public Function DeleteRelation(ByVal plngRelatedTableID As Long) As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  'TM19012004
  ' Check that the relationship is not used.
  Set mfrmUse = New frmUsage
  mfrmUse.ResetList
  If RelationIsUsed(plngRelatedTableID, mfrmUse) Then
    Screen.MousePointer = vbDefault
    mfrmUse.ShowMessage Me.TableName & " - " & GetTableName(plngRelatedTableID) & " Relationship", "The relationship cannot be deleted as the relationship is used by the following:", UsageCheckObject.Relationship
    fOK = False
  End If
  UnLoad mfrmUse
  Set mfrmUse = Nothing
  
 
  If fOK Then
  
    ' Delete foreign key column for this relation
    daoDb.Execute "UPDATE tmpColumns" & _
      " SET deleted=TRUE" & _
      " WHERE columnName='ID_" & TableID & "'" & _
      " AND tableID=" & plngRelatedTableID, dbFailOnError
    
    ' Set related table as changed
    daoDb.Execute "UPDATE tmpTables" & _
      " SET changed=TRUE, lastUpdated=Now()" & _
      " WHERE tableID=" & plngRelatedTableID, dbFailOnError
      
    ' Delete relation
    daoDb.Execute "DELETE FROM tmpRelations" & _
      " WHERE parentID=" & TableID & _
      " AND childID=" & plngRelatedTableID, dbFailOnError
    
  End If
  
TidyUpAndExit:
  DeleteRelation = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Public Function RelationIsUsed(pLngChildTableID As Long, pfrmUsage As frmUsage) As Boolean
  
  ' Return true if the relationship is used somewhere and
  ' therefore cannot be deleted.
  '
  ' Relationships may be used in the following context :
  '
  '   As the basis of a Field Component in an expression.
  '   ie. a field component in an expression can only refer to
  '   a field in a table that is related directly to the
  '   expression's parent table. If the relationship between the
  '   two tables is deleted then the field component cannot
  '   make any connection between the two tables, and so cannot
  '   be evaluated.
  '
  '   A screen may contain controls for columns from its associated
  '   table, and also for columns in its parent table. If the table's
  '   columns are used in any other screens then do not allow the
  '   user to delete it
  '
  '   Module definitions.
  '
  '   Sort Order/Find Window definitions can include columns from the
  '   parent tables of the order's base table.
  '
  '   Each screens can have a set of associated history screens. The set of possible
  '   history screens is the set of screens associated with child tables of the
  '   base screen's table.
  '
  '   Unique (within sibling records) checks.
  '
  '   Utility/Report definitions
  On Error GoTo ErrorTrap
  
  Set mfrmUse = pfrmUsage
    
  Dim fUsed As Boolean
  Dim lngExprParentTableID As Long
  Dim sSQL As String
  Dim sExprName As String
  Dim sExprParentTable As String
  Dim sExprType As String
  Dim sChildTable As String
  Dim sParentTable As String
  Dim rsExpressions As DAO.Recordset
  Dim rsScreens As DAO.Recordset
  Dim rsLinks As DAO.Recordset
  Dim rsCheck1 As ADODB.Recordset
  Dim rsCheck2 As ADODB.Recordset
  Dim rsCheck3 As ADODB.Recordset
  Dim rsCheck4 As ADODB.Recordset
  Dim rsCheck5 As ADODB.Recordset
  Dim rsCheck6 As ADODB.Recordset
  Dim rsCheck7 As ADODB.Recordset
  Dim rsCheck8 As ADODB.Recordset
  Dim rsCheck9 As ADODB.Recordset
  Dim rsCheck10 As ADODB.Recordset
  Dim objExpr As CExpression
  Dim rsModules As DAO.Recordset
  Dim rsInfo As DAO.Recordset
  Dim sModuleName As String
  Dim sType As String
  Dim objComp As CExprComponent
  Dim lngExprID As Long
  Dim sAccess As String
  Dim rsDAOTemp As DAO.Recordset
  Dim rsDAOTemp2 As DAO.Recordset
  Dim lngPersonnelTableID As Long

  lngPersonnelTableID = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_PERSONNELTABLE, 0)
  
  ' Get the child table name.
  recTabEdit.Index = "idxTableID"
  recTabEdit.Seek "=", pLngChildTableID
  If Not recTabEdit.NoMatch Then
    sChildTable = recTabEdit!TableName
  Else
    sChildTable = "<unknown>"
  End If

  ' Check that the relation is not used as the basis for an expression component.
  sSQL = "SELECT DISTINCT tmpExpressions.exprID, tmpExpressions.Type, tmpExpressions.Name, " & _
                          "tmpExpressions.TableID, tmpTables.TableName " & _
        " FROM tmpExpressions, tmpTables " & _
        " WHERE (tmpExpressions.TableID = tmpTables.TableID )" & _
          " AND (tmpExpressions.Type IN (" & giEXPR_LINKFILTER & "," & giEXPR_COLUMNCALCULATION & "," & _
                                      giEXPR_DEFAULTVALUE & "," & giEXPR_STATICFILTER & "," & _
                                      giEXPR_GOTFOCUS & "," & giEXPR_RECORDVALIDATION & "," & _
                                      giEXPR_VIEWFILTER & "," & giEXPR_EMAIL & "," & giEXPR_RECORDDESCRIPTION & "," & _
                                      giEXPR_OUTLOOKSUBJECT & "," & giEXPR_OUTLOOKFOLDER & "))" & _
          " AND (tmpExpressions.ParentComponentID = 0)"

  Set rsExpressions = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  With rsExpressions
    If Not (.BOF And .EOF) Then
  '    fUsed = True
      Do Until .EOF
        If ExpressionUsesRelationship(!ExprID, gLngTableID, pLngChildTableID, !TableID) Then
          fUsed = True
            
          Select Case !Type
            Case ExpressionTypes.giEXPR_LINKFILTER
              sExprType = "Link Filter"
            
            Case ExpressionTypes.giEXPR_COLUMNCALCULATION
              sExprType = "Column Calculation"
            
            Case ExpressionTypes.giEXPR_DEFAULTVALUE
              sExprType = "Default Value"
            
            Case ExpressionTypes.giEXPR_STATICFILTER
              sExprType = "Static Filter"
            
            Case ExpressionTypes.giEXPR_GOTFOCUS
              sExprType = "Field Entry Validation Clause"
            
            Case ExpressionTypes.giEXPR_RECORDVALIDATION
              sExprType = "Field Entry Validation"
            
            Case ExpressionTypes.giEXPR_EMAIL
              sExprType = "Calculated Email Address"
            
            Case ExpressionTypes.giEXPR_RECORDDESCRIPTION
              sExprType = "Record Description"
            
            Case ExpressionTypes.giEXPR_OUTLOOKSUBJECT
              sExprType = "Subject Calculation"
            
            Case ExpressionTypes.giEXPR_OUTLOOKFOLDER
              sExprType = "Calendar Folder"
            
            Case Else
              sExprType = "Expression"
              
          End Select
          
          mfrmUse.AddToList (sExprType & " : " & !Name & " <" & !TableName & ">")
  
        End If
  
        .MoveNext
      Loop
    End If
    'Close temporary recordset
    .Close
  End With
 
  
  ' Find any screens that contain controls for
  ' columns based on this relationship.
  sSQL = "SELECT DISTINCT tmpScreens.name, tmpScreens.tableID" & _
    " FROM tmpScreens, tmpControls" & _
    " WHERE tmpScreens.deleted = FALSE" & _
    " AND tmpScreens.screenID = tmpControls.screenID" & _
    " AND tmpScreens.tableID = " & Trim(Str(pLngChildTableID)) & _
    " AND tmpControls.tableID = " & Trim(Str(gLngTableID))
  Set rsScreens = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  If Not (rsScreens.BOF And rsScreens.EOF) Then
    fUsed = True
    Do Until rsScreens.EOF
      ' Get the expression's parent table name.
      recTabEdit.Index = "idxTableID"
      recTabEdit.Seek "=", rsScreens.Fields("tableID")

      If Not recTabEdit.NoMatch Then
        sParentTable = recTabEdit!TableName
      Else
        sParentTable = "<unknown>"
      End If
      mfrmUse.AddToList ("Screen : " & rsScreens.Fields("name") & " <" & sParentTable & ">")
      rsScreens.MoveNext
    Loop
  End If
  rsScreens.Close
  
  
  ' Find any Link columns that are based on this relationship.
  sSQL = "SELECT DISTINCT tmpColumns.columnName, tmpTables.tableName" & _
    " FROM tmpColumns, tmpTables" & _
    " WHERE tmpTables.deleted = FALSE" & _
    " AND tmpColumns.deleted = FALSE" & _
    " AND tmpColumns.columnType = " & Trim(Str(giCOLUMNTYPE_LINK)) & _
    " AND tmpColumns.tableID = " & Trim(Str(pLngChildTableID)) & _
    " AND tmpColumns.linkTableID = " & Trim(Str(gLngTableID)) & _
    " AND tmpTables.tableID = " & Trim(Str(pLngChildTableID))
  Set rsLinks = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  If Not (rsLinks.BOF And rsLinks.EOF) Then
    fUsed = True
    Do Until rsLinks.EOF
      mfrmUse.AddToList ("Link Column : " & rsLinks.Fields("columnName") & " <" & rsLinks.Fields("tableName") & ">")
      rsLinks.MoveNext
    Loop
  End If
  ' Close temporary recordset.
  rsLinks.Close

  
  ' Find any Module Definitions that use this relationship.
  sSQL = "SELECT DISTINCT a.moduleKey" & _
    " FROM tmpModuleSetup a, tmpModuleSetup b" & _
    " WHERE ((a.parameterKey = '" & gsPARAMETERKEY_COURSETABLE & "' AND a.parameterValue = '" & Trim(Str(gLngTableID)) & "')" & _
    "   AND (b.parameterKey = '" & gsPARAMETERKEY_PREREQTABLE & "' AND b.parameterValue = '" & Trim(Str(pLngChildTableID)) & "'))" & _
    " OR" & _
    " ((a.parameterKey = '" & gsPARAMETERKEY_COURSETABLE & "' AND a.parameterValue = '" & Trim(Str(gLngTableID)) & "')" & _
    "   AND (b.parameterKey = '" & gsPARAMETERKEY_TRAINBOOKTABLE & "' AND b.parameterValue = '" & Trim(Str(pLngChildTableID)) & "'))" & _
    " OR" & _
    " ((a.parameterKey = '" & gsPARAMETERKEY_EMPLOYEETABLE & "' AND a.parameterValue = '" & Trim(Str(gLngTableID)) & "')" & _
    "   AND (b.parameterKey = '" & gsPARAMETERKEY_TRAINBOOKTABLE & "' AND b.parameterValue = '" & Trim(Str(pLngChildTableID)) & "'))" & _
    " OR" & _
    " ((a.parameterKey = '" & gsPARAMETERKEY_EMPLOYEETABLE & "' AND a.parameterValue = '" & Trim(Str(gLngTableID)) & "')" & _
    "   AND (b.parameterKey = '" & gsPARAMETERKEY_UNAVAILTABLE & "' AND b.parameterValue = '" & Trim(Str(pLngChildTableID)) & "'))" & _
    " OR" & _
    " ((a.parameterKey = '" & gsPARAMETERKEY_EMPLOYEETABLE & "' AND a.parameterValue = '" & Trim(Str(gLngTableID)) & "')" & _
    " AND (b.parameterKey = '" & gsPARAMETERKEY_WAITLISTTABLE & "' AND b.parameterValue = '" & Trim(Str(pLngChildTableID)) & "'))"
  Set rsModules = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  If Not (rsModules.BOF And rsModules.EOF) Then
    fUsed = True
    Do Until rsModules.EOF
      Select Case rsModules!moduleKey
        Case gsMODULEKEY_TRAININGBOOKING
          sModuleName = "Training Booking"
        Case gsMODULEKEY_PERSONNEL
          sModuleName = "Personnel"
        Case gsMODULEKEY_ABSENCE
          sModuleName = "Absence"
        Case gsMODULEKEY_CURRENCY
          sModuleName = "Currency"
        Case gsMODULEKEY_POST
          sModuleName = "Post"
        Case gsMODULEKEY_MATERNITY
          sModuleName = "Maternity"
        Case gsMODULEKEY_SSINTRANET
          sModuleName = "Self Service Intranet"
        Case gsMODULEKEY_HIERARCHY
          sModuleName = "Hierachy"
        Case Else
          sModuleName = "<Unknown>"
      End Select
      mfrmUse.AddToList (sModuleName & " Configuration")
      rsModules.MoveNext
    Loop
  End If
  ' Close the recordset.
  rsModules.Close
  
  ' Self-service Intranet
  sSQL = "SELECT COUNT(*) AS result" & _
    " FROM tmpSSIntranetLinks" & _
    " INNER JOIN tmpScreens ON tmpSSIntranetLinks.screenID = tmpScreens.screenID" & _
    " WHERE tmpScreens.deleted = FALSE" & _
    " AND tmpSSIntranetLinks.screenID > 0" & _
    " AND tmpScreens.tableid = " & CStr(pLngChildTableID) & _
    " AND tmpSSIntranetLinks.tableid = " & CStr(gLngTableID)
  Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  If rsInfo!result > 0 Then
    fUsed = True
    mfrmUse.AddToList ("Self-service Intranet module")
  End If
  ' Close temporary recordset.
  rsInfo.Close
    
  ' Find any Sort Order/Find Window columns that are based on this relationship.
  sSQL = "SELECT DISTINCT tmpOrders.name" & _
    " FROM tmpOrders, tmpOrderItems, tmpColumns" & _
    " WHERE tmpOrders.deleted = FALSE" & _
    " AND tmpOrders.tableID = " & Trim(Str(pLngChildTableID)) & _
    " AND tmpOrders.orderID = tmpOrderItems.orderID" & _
    " AND tmpOrderItems.columnID = tmpColumns.columnID" & _
    " AND tmpColumns.tableID = " & Trim(Str(gLngTableID))
  Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  If Not (rsInfo.BOF And rsInfo.EOF) Then
    fUsed = True
    Do Until rsInfo.EOF
      mfrmUse.AddToList ("Order : " & rsInfo!Name & " <" & sChildTable & ">")
      rsInfo.MoveNext
    Loop
  End If
  rsInfo.Close
        
    
  ' Find any screen/history screen associations that are based on this relationship.
  sSQL = "SELECT DISTINCT a.name AS parentScreen, b.name AS childScreen" & _
    " FROM tmpScreens a, tmpScreens b, tmpHistoryScreens" & _
    " WHERE a.screenID = tmpHistoryScreens.parentScreenID" & _
    "   AND b.screenID = tmpHistoryScreens.historyScreenID" & _
    "   AND a.tableID = " & Trim(Str(gLngTableID)) & _
    "   AND b.tableID = " & Trim(Str(pLngChildTableID))
  Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  If Not (rsInfo.BOF And rsInfo.EOF) Then
    fUsed = True
    Do Until rsInfo.EOF
      mfrmUse.AddToList ("History Screen : " & rsInfo!childScreen & " <" & rsInfo!parentScreen & ">")
      rsInfo.MoveNext
    Loop
  End If
  rsInfo.Close
  
  
  
  ' Find any email links which contain columns that are based on this relationship.
  sSQL = "SELECT el.Title " & _
         "FROM   tmpEmailLinks el, tmpLinkContent lc, tmpColumns c " & _
         "WHERE  (el.SubjectContentID = lc.ContentID OR el.BodyContentID = lc.ContentID) " & _
         "   AND lc.FieldCode = 'C'" & _
         "   AND lc.FieldID = c.columnid" & _
         "   AND c.tableID = " & CStr(gLngTableID) & _
         "   AND el.TableID = " & CStr(pLngChildTableID)
  Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  If Not (rsInfo.BOF And rsInfo.EOF) Then
    fUsed = True
    Do Until rsInfo.EOF
      mfrmUse.AddToList ("Email Content : " & rsInfo!Title & " <" & sChildTable & ">")
      rsInfo.MoveNext
    Loop
  End If
  rsInfo.Close
  
  
  
  ' Find any column definitions that use this relationship
  ' as part of the 'unique within sibling records' checks.
  sSQL = "SELECT DISTINCT tmpColumns.columnName, tmpTables.tableName" & _
    " FROM tmpColumns, tmpTables" & _
    " WHERE tmpTables.deleted = FALSE" & _
    " AND tmpColumns.deleted = FALSE" & _
    " AND tmpColumns.uniqueCheckType = " & Trim(Str(gLngTableID)) & _
    " AND tmpColumns.tableID = " & Trim(Str(pLngChildTableID)) & _
    " AND tmpTables.tableID = " & Trim(Str(pLngChildTableID))
  Set rsLinks = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  If Not (rsLinks.BOF And rsLinks.EOF) Then
    fUsed = True
    Do Until rsLinks.EOF
      mfrmUse.AddToList ("Unique (within sibling records) check : " & rsLinks.Fields("columnName") & " <" & rsLinks.Fields("tableName") & ">")
      rsInfo.MoveNext
    Loop
  End If
  ' Close temporary recordset.
  rsLinks.Close
    
    
  ' Find any Custom Reports that use this relationship.
  ' Base Table - Parent 1 (if any columns/calcs from Parent 1 are in the report)
  ' Base Table - Parent 2 (if any columns/calcs from Parent 2 are in the report)
  ' Base Table - any related child table in the report
  sSQL = "SELECT DISTINCT ASRSysCustomReportsName.ID, ASRSysCustomReportsName.name" & _
    " FROM ASRSysCustomReportsName" & _
    "   INNER JOIN ASRSysCustomReportsDetails " & _
    "   ON ASRSysCustomReportsName.ID = ASRSysCustomReportsDetails.customReportID" & _
    "      AND ASRSysCustomReportsDetails.type = 'C'" & _
    "         INNER JOIN ASRSysColumns " & _
    "         ON ASRSysCustomReportsDetails.colExprID = ASRSysColumns.columnID" & _
    " WHERE ((ASRSysCustomReportsName.baseTable = " & Trim(Str(gLngTableID)) & ") AND (ASRSysColumns.tableID = " & Trim(Str(pLngChildTableID)) & "))" & _
    "     OR ((ASRSysCustomReportsName.baseTable = " & Trim(Str(pLngChildTableID)) & ") AND (ASRSysColumns.tableID = " & Trim(Str(gLngTableID)) & "))"
  
  Set rsCheck1 = New ADODB.Recordset
  rsCheck1.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsCheck1
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Custom Report : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsCheck1 = Nothing


  ' Find any Record Profiles that use this relationship.
  ' Base Table - any related parent/child table in the report
  sSQL = "SELECT DISTINCT aRPN.name" & _
    " FROM ASRSysRecordProfileName aRPN" & _
    " INNER JOIN ASRSysRecordProfileDetails aRPD ON aRPN.recordProfileID = aRPD.recordProfileID" & _
    "   AND aRPD.type = 'C'" & _
    " INNER JOIN ASRSysColumns aC ON aRPD.columnID = aC.columnID" & _
    "   AND aC.tableID = " & CStr(gLngTableID) & _
    " INNER JOIN ASRSysRecordProfileName bRPN ON aRPN.recordProfileID = bRPN.recordProfileID" & _
    " INNER JOIN ASRSysRecordProfileDetails bRPD ON bRPN.recordProfileID = bRPD.recordProfileID" & _
    "   AND bRPD.type = 'C'" & _
    " INNER JOIN ASRSysColumns bC ON bRPD.columnID = bC.columnID" & _
    "   AND bC.tableID = " & CStr(pLngChildTableID)
  
  Set rsCheck2 = New ADODB.Recordset
  rsCheck2.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsCheck2
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Record Profile : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsCheck2 = Nothing


  ' Find any Mail Merge/Envelopes & Labels that use this relationship.
  ' Base Table - any related parent table with columns/calcs in the report
  sSQL = "SELECT DISTINCT ASRSysMailMergeName.name, ASRSysMailMergeName.isLabel" & _
    " FROM ASRSysMailMergeName" & _
    " INNER JOIN ASRSysMailMergeColumns ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeColumns.mailMergeID" & _
    "   AND ASRSysMailMergeColumns.type = 'C'" & _
    " INNER JOIN ASRSysColumns ON ASRSysMailMergeColumns.columnID = ASRSysColumns.columnID" & _
    "   AND ASRSysColumns.tableid = " & CStr(gLngTableID) & _
    " WHERE ASRSysMailMergeName.tableID = " & CStr(pLngChildTableID) & _
    " UNION" & _
    " SELECT ASRSysMailMergeName.name, ASRSysMailMergeName.isLabel" & _
    " FROM ASRSysMailMergeName" & _
    " INNER JOIN ASRSysMailMergeColumns ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeColumns.mailMergeID" & _
    "   AND ASRSysMailMergeColumns.type = 'E'" & _
    " INNER JOIN ASRSysExpressions ON ASRSysMailMergeColumns.columnid = ASRSysExpressions.exprID" & _
    "   AND ASRSysExpressions.tableID = " & CStr(gLngTableID) & _
    " WHERE ASRSysMailMergeName.tableID = " & CStr(pLngChildTableID)
  
  Set rsCheck3 = New ADODB.Recordset
  rsCheck3.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsCheck3
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
      
        mfrmUse.AddToList (IIf(!IsLabel, "Envelopes & Labels", "Mail Merge") & " : " & !Name)
'            ' Tell the user why the relation cannot be deleted.
'          MsgBox "The relationship between the '" & TableName & "' and '" & sChildTable & "' tables cannot be deleted." & vbCr & vbCr & _
'            "It is used in the '" & !Name & "' " & IIf(!IsLabel, "Envelopes & Labels", "Mail Merge") & ".", _
'            vbExclamation + vbOKOnly, Application.Name
          
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsCheck3 = Nothing


  
  ' Find any Calendar Reports that use this relationship.
  ' Base Table - any related child table in the report
  sSQL = "SELECT DISTINCT ASRSysCalendarReports.name" & _
    " FROM ASRSysCalendarReports" & _
    " INNER JOIN ASRSysCalendarReportEvents ON ASRSysCalendarReports.ID = ASRSysCalendarReportEvents.calendarReportID" & _
    "   AND tableID = " & CStr(pLngChildTableID) & _
    " WHERE ASRSysCalendarReports.baseTable = " & CStr(gLngTableID)

  Set rsCheck4 = New ADODB.Recordset
  rsCheck4.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsCheck4
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Calendar Report : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsCheck4 = Nothing


  ' Find any Match/Career/Succession Reports that use this relationship.
  ' Base Table - any related child table in the report
  ' Match Table - any related child table in the report
  sSQL = "SELECT DISTINCT ASRSysMatchReportName.name, ASRSysMatchReportName.matchReportType" & _
    " FROM ASRSysMatchReportName" & _
    " INNER JOIN ASRSysMatchReportTables ON ASRSysMatchReportName.matchReportID = ASRSysMatchReportTables.matchReportID" & _
    "   AND ASRSysMatchReportTables.table1ID = " & CStr(pLngChildTableID) & _
    " WHERE ASRSysMatchReportName.table1ID = " & CStr(gLngTableID) & _
    " UNION" & _
    " SELECT ASRSysMatchReportName.name, ASRSysMatchReportName.matchReportType" & _
    " FROM ASRSysMatchReportName" & _
    " INNER JOIN ASRSysMatchReportTables ON ASRSysMatchReportName.matchReportID = ASRSysMatchReportTables.matchReportID" & _
    "   AND ASRSysMatchReportTables.table2ID = " & CStr(pLngChildTableID) & _
    " WHERE ASRSysMatchReportName.table2ID = " & CStr(gLngTableID)
  
  Set rsCheck5 = New ADODB.Recordset
  rsCheck5.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsCheck5
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
        ' Tell the user why the relation cannot be deleted.
        Select Case !matchReportType
          Case 1
            sType = "Succession Planning"
          Case 2
            sType = "Career Progression"
          Case Else
            sType = "Match Report"
        End Select
        mfrmUse.AddToList (sType & " : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsCheck5 = Nothing

  
  ' Find any Global Add Functions that use this relationship.
  ' Base Table - child table
  sSQL = "SELECT DISTINCT name" & _
    " FROM ASRSysGlobalFunctions" & _
    " WHERE tableID = " & CStr(gLngTableID) & _
    "   AND childTableID = " & CStr(pLngChildTableID)
  
  Set rsCheck6 = New ADODB.Recordset
  rsCheck6.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsCheck6
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Global Add : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsCheck6 = Nothing

  
  ' Find any Data Transfers that use this relationship.
  ' Source Table - source column might be child of the source table
  ' Destination Table - destination column might be child of the destination table
  sSQL = "SELECT DISTINCT ASRSysDataTransferName.name" & _
    " FROM ASRSysDataTransferName" & _
    " INNER JOIN ASRSysDataTransferColumns ON ASRSysDataTransferName.dataTransferID = ASRSysDataTransferColumns.dataTransferID" & _
    "   AND ASRSysDataTransferColumns.fromTableID = " & CStr(pLngChildTableID) & _
    " WHERE ASRSysDataTransferName.fromTableID = " & CStr(gLngTableID) & _
    " UNION" & _
    " SELECT ASRSysDataTransferName.name" & _
    " FROM ASRSysDataTransferName" & _
    " INNER JOIN ASRSysDataTransferColumns ON ASRSysDataTransferName.dataTransferID = ASRSysDataTransferColumns.dataTransferID" & _
    "   AND ASRSysDataTransferColumns.toTableID = " & CStr(pLngChildTableID) & _
    " WHERE ASRSysDataTransferName.toTableID = " & CStr(gLngTableID)
  
  Set rsCheck7 = New ADODB.Recordset
  rsCheck7.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsCheck7
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Data Transfer : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsCheck7 = Nothing


  ' Find any Imports that use this relationship.
  ' Base Table - import column might be parent of the base table
  sSQL = "SELECT DISTINCT ASRSysImportName.name" & _
    " FROM ASRSysImportName" & _
    "     INNER JOIN ASRSysImportDetails " & _
    "     ON ASRSysImportName.ID = ASRSysImportDetails.importID" & _
    " WHERE ASRSysImportName.baseTable = " & CStr(pLngChildTableID) & _
    "   AND ASRSysImportDetails.tableid = " & CStr(gLngTableID)
  
  Set rsCheck8 = New ADODB.Recordset
  rsCheck8.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsCheck8
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Import : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsCheck8 = Nothing

  
  ' Find any Exports that use this relationship.
  ' Base Table - emport column might be parent or child of the base table
  sSQL = "SELECT DISTINCT ASRSysExportName.name" & _
    " FROM ASRSysExportName" & _
    " INNER JOIN ASRSysExportDetails ON ASRSysExportName.ID = ASRSysExportDetails.exportID" & _
    "   AND ASRSysExportDetails.tableID = " & CStr(gLngTableID) & _
    " WHERE ASRSysExportName.baseTable = " & CStr(pLngChildTableID) & _
    " UNION" & _
    " SELECT ASRSysExportName.name" & _
    " FROM ASRSysExportName" & _
    " INNER JOIN ASRSysExportDetails ON ASRSysExportName.ID = ASRSysExportDetails.exportID" & _
    "   AND ASRSysExportDetails.tableID = " & CStr(pLngChildTableID) & _
    " WHERE ASRSysExportName.baseTable = " & CStr(gLngTableID)
  
  Set rsCheck9 = New ADODB.Recordset
  rsCheck9.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsCheck9
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Export : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsCheck9 = Nothing
  

  ' Now Check if the poor Relation is used in a Runtime Calculation OR Runtime Filter.
  sSQL = "SELECT DISTINCT ASRSysExpressions.exprID, ASRSysExpressions.Type, ASRSysExpressions.Name, ASRSysTables.TableName " & _
        " FROM ASRSysExpressions " & _
        "       LEFT OUTER JOIN  ASRSysExprComponents " & _
        "       ON ASRSysExpressions.ExprID = ASRSysExprComponents.ExprID " & _
        "       LEFT OUTER JOIN ASRSysTables" & _
        "       ON ASRSysExpressions.TableID = ASRSysTables.TableID" & _
        " WHERE ASRSysExpressions.Type IN (" & giEXPR_RUNTIMECALCULATION & "," & giEXPR_RUNTIMEFILTER & ") " & _
        "   AND ASRSysExpressions.ParentComponentID = 0"
  Set rsCheck10 = New ADODB.Recordset
  rsCheck10.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsCheck10
    If Not (.BOF And .EOF) Then
'      fUsed = True
      Do Until .EOF
        If ExpressionUsesRelationship_SQL(!ExprID, gLngTableID, pLngChildTableID) Then
          fUsed = True
        
          Select Case !Type
            Case ExpressionTypes.giEXPR_RUNTIMEFILTER
              sExprType = "Runtime Filter"
              mfrmUse.AddToList (sExprType & " : " & !Name & " <" & !TableName & ">")
            Case ExpressionTypes.giEXPR_RUNTIMECALCULATION
              sExprType = "Runtime Calculation"
              mfrmUse.AddToList (sExprType & " : " & !Name & " <" & !TableName & ">")
            Case Else
              sExprType = "Expression"
          End Select
          
        End If
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsCheck10 = Nothing
  
  ' Check that it is not used by a Workflow storedData element.
  sSQL = "SELECT tmpWorkflowElements.workflowID," & _
    "   tmpWorkflowElements.caption," & _
    "   tmpWorkflowElements.dataRecord," & _
    "   tmpWorkflowElements.recSelWebFormIdentifier," & _
    "   tmpWorkflowElements.recSelIdentifier," & _
    "   tmpWorkflows.name" & _
    " FROM tmpWorkflowElements" & _
    " INNER JOIN tmpWorkflows ON tmpWorkflowElements.workflowID = tmpWorkflows.ID" & _
    " WHERE tmpWorkflowElements.type = 5" & _
    "   AND tmpWorkflowElements.dataTableID = " & CStr(pLngChildTableID) & _
    "   AND tmpWorkflowElements.dataAction = 0" & _
    "   AND tmpWorkflows.deleted = false"

  Set rsDAOTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  If Not (rsDAOTemp.BOF And rsDAOTemp.EOF) Then
    Do Until rsDAOTemp.EOF
      Select Case rsDAOTemp.Fields("dataRecord")
        Case giWFRECSEL_INITIATOR
          ' Stored data element is INSERTing into the child table
          ' based on its relationship to the initiator's record (ie. the personnel table)
          If gLngTableID = lngPersonnelTableID Then
            fUsed = True
            mfrmUse.AddToList ("Workflow : " & rsDAOTemp.Fields("name") & " <'" & rsDAOTemp.Fields("caption") & "' stored data element>")
          End If
          
        Case giWFRECSEL_IDENTIFIEDRECORD
          ' Stored data element is INSERTing into the child table
          ' based on its relationship to a previously selected record (using record selector)
          sSQL = "SELECT COUNT(*) AS [result]" & _
            " FROM tmpWorkflowElementItems" & _
            " INNER JOIN tmpWorkflowElements ON tmpWorkflowElementItems.elementID = tmpWorkflowElements.id" & _
            " WHERE tmpWorkflowElements.workflowID = " & CStr(rsDAOTemp.Fields("workflowID")) & _
            "   AND tmpWorkflowElements.identifier = '" & rsDAOTemp.Fields("recSelWebFormIdentifier") & "'" & _
            "   AND tmpWorkflowElementItems.identifier = '" & rsDAOTemp.Fields("recSelIdentifier") & "'" & _
            "   AND tmpWorkflowElementItems.tableID = " & CStr(gLngTableID)
        
          Set rsDAOTemp2 = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
          If rsDAOTemp2.Fields("result") > 0 Then
            fUsed = True
            mfrmUse.AddToList ("Workflow : " & rsDAOTemp.Fields("name") & " <'" & rsDAOTemp.Fields("caption") & "' stored data element>")
          End If
      End Select
      
      rsDAOTemp.MoveNext
    Loop
  End If
  'Close temporary recordset
  rsDAOTemp.Close
  
  ' Check that it is not used by a Workflow WebForm element.
  sSQL = "SELECT tmpWorkflowElementItems.dbRecord," & _
    "   tmpWorkflowElements.workflowID," & _
    "   tmpWorkflowElements.caption," & _
    "   tmpWorkflowElementItems.wfFormIdentifier," & _
    "   tmpWorkflowElementItems.wfValueIdentifier" & _
    " FROM tmpWorkflowElementItems" & _
    " INNER JOIN tmpWorkflowElements ON tmpWorkflowElementItems.elementID = tmpWorkflowElements.ID" & _
    " WHERE tmpWorkflowElementItems.itemType = 11" & _
    "   AND tmpWorkflowElementItems.tableID = " & CStr(pLngChildTableID)

  Set rsDAOTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  If Not (rsDAOTemp.BOF And rsDAOTemp.EOF) Then
    Do Until rsDAOTemp.EOF
      
      recWorkflowEdit.Index = "idxWorkflowID"
      recWorkflowEdit.Seek "=", rsDAOTemp.Fields("workflowID")

      If Not recWorkflowEdit.NoMatch Then
        If recWorkflowEdit.Fields("deleted").value = False Then
      
          Select Case rsDAOTemp.Fields("dbRecord")
            Case giWFRECSEL_INITIATOR
              ' Web form element record selector is for child records of the initiator's record (ie. the personnel table)
              If gLngTableID = lngPersonnelTableID Then
                fUsed = True
                mfrmUse.AddToList ("Workflow : " & rsDAOTemp.Fields("name") & " <'" & rsDAOTemp.Fields("caption") & "' web form record selector>")
              End If
    
            Case giWFRECSEL_IDENTIFIEDRECORD
              ' Web form element record selector is for child records of a previously selected record (using record selector)
              sSQL = "SELECT COUNT(*) AS [result]" & _
                " FROM tmpWorkflowElementItems" & _
                " INNER JOIN tmpWorkflowElements ON tmpWorkflowElementItems.elementID = tmpWorkflowElements.id" & _
                " WHERE tmpWorkflowElements.workflowID = " & CStr(rsDAOTemp.Fields("workflowID")) & _
                "   AND tmpWorkflowElements.identifier = '" & rsDAOTemp.Fields("wfFormIdentifier") & "'" & _
                "   AND tmpWorkflowElementItems.identifier = '" & rsDAOTemp.Fields("wfValueIdentifier") & "'" & _
                "   AND tmpWorkflowElementItems.tableID = " & CStr(gLngTableID)
    
              Set rsDAOTemp2 = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
              If rsDAOTemp2.Fields("result") > 0 Then
                fUsed = True
                mfrmUse.AddToList ("Workflow : " & recWorkflowEdit.Fields("name").value & " <'" & rsDAOTemp.Fields("caption") & "' web form record selector>")
              End If
          End Select
        End If
      End If

      rsDAOTemp.MoveNext
    Loop
  End If
  'Close temporary recordset
  rsDAOTemp.Close
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set rsDAOTemp = Nothing
  Set rsModules = Nothing
  Set rsExpressions = Nothing
  Set rsScreens = Nothing
  Set rsLinks = Nothing
  Set rsInfo = Nothing
  Set objExpr = Nothing
  
  RelationIsUsed = fUsed
  Exit Function

ErrorTrap:
  fUsed = True
  Resume TidyUpAndExit
  
End Function


Public Function TableIsUsed(pfrmUsage As frmUsage) As Boolean
 
  ' Return true if the table is used somewhere and
  ' therefore cannot be deleted.
  '
  ' Tables may be used in the following context :
  '
  '   As part of a Field Component in an expression.
  '   As the lookup table for a column.
  '   Module definitions.
  On Error GoTo ErrorTrap
  
  Set mfrmUse = pfrmUsage
  
  Dim fUsed As Boolean
  Dim lngExprParentTableID As Long
  Dim sSQL As String
  Dim sExprType As String
  Dim sExprName As String
  Dim sParentTable As String
  Dim objExpr As CExpression
  Dim rsExpressions As DAO.Recordset
  Dim rsColumns As DAO.Recordset
  Dim rsModules As DAO.Recordset
  Dim rsDAOTemp As DAO.Recordset
  Dim sModuleName As String
  Dim rsGlobalFunctions As ADODB.Recordset
  Dim rsUtils1 As ADODB.Recordset
  Dim rsUtils2 As ADODB.Recordset
  Dim rsUtils3 As ADODB.Recordset
  Dim rsUtils4 As ADODB.Recordset
  Dim rsUtils5 As ADODB.Recordset
  Dim rsUtils6 As ADODB.Recordset
  Dim rsUtils7 As ADODB.Recordset
  Dim rsUtils8 As ADODB.Recordset
  Dim rsUtils9 As ADODB.Recordset
  Dim rsUtils10 As ADODB.Recordset
  Dim rsUtils11 As ADODB.Recordset
  Dim rsUtils12 As ADODB.Recordset
  Dim rsUtils13 As ADODB.Recordset
  
  Dim rsUtilsDocTypes As ADODB.Recordset
  Dim sGlobalFunctionType As String
  Dim rsEmail As DAO.Recordset
  Dim strSYSOrderIDs As String
  Dim strSYSEmailIDs As String
  Dim rsOrders As DAO.Recordset
  Dim strSYSCheckedOrderIDs As String
  
  Dim bTR As Boolean
  Dim bPE As Boolean
  Dim bAB As Boolean
  Dim bCU As Boolean
  Dim bPO As Boolean
  Dim bMA As Boolean
  Dim bSS As Boolean
  Dim bHI As Boolean
  Dim bBA As Boolean
  Dim bUN As Boolean
  Dim bWF As Boolean
  Dim bAu As Boolean
  
  fUsed = False
  
  strSYSOrderIDs = vbNullString
  strSYSEmailIDs = vbNullString
  
  ' Find any expressions that use any columns from this table
  ' as the basis for a field component.
  sSQL = "SELECT DISTINCT tmpExpressions.exprID" & _
    " FROM tmpExpressions, tmpComponents" & _
    " WHERE tmpExpressions.deleted = FALSE" & _
    " AND tmpExpressions.exprID = tmpComponents.exprID" & _
    " AND tmpComponents.type = " & Trim(Str(giCOMPONENT_FIELD)) & _
    " AND tmpExpressions.TableID = " & Trim(Str(gLngTableID)) & _
    " AND tmpComponents.fieldTableID = " & Trim(Str(gLngTableID)) & _
    " AND tmpExpressions.ParentComponentID = 0" & _
    " AND ((tmpExpressions.Type = " & giEXPR_LINKFILTER & ")" & _
    " OR  (tmpExpressions.Type = " & giEXPR_COLUMNCALCULATION & ")" & _
    " OR  (tmpExpressions.Type = " & giEXPR_DEFAULTVALUE & ")" & _
    " OR  (tmpExpressions.Type = " & giEXPR_STATICFILTER & ")" & _
    " OR  (tmpExpressions.Type = " & giEXPR_GOTFOCUS & ")" & _
    " OR  (tmpExpressions.Type = " & giEXPR_RECORDVALIDATION & ")" & _
    " OR  (tmpExpressions.Type = " & giEXPR_VIEWFILTER & ")" & _
    " OR  (tmpExpressions.Type = " & giEXPR_EMAIL & ")" & _
    " OR  (tmpExpressions.Type = " & giEXPR_RECORDDESCRIPTION & ")" & _
    " OR  (tmpExpressions.Type = " & giEXPR_OUTLOOKSUBJECT & ")" & _
    " OR  (tmpExpressions.Type = " & giEXPR_OUTLOOKFOLDER & "))"
  Set rsExpressions = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  If Not (rsExpressions.BOF And rsExpressions.EOF) Then
    fUsed = True
    Do Until rsExpressions.EOF
      
      ' Get the expression name and type description.
      Set objExpr = New CExpression
      objExpr.ExpressionID = rsExpressions.Fields("exprId")

      If objExpr.ReadExpressionDetails Then

        sExprName = objExpr.Name
        sExprType = objExpr.ExpressionTypeName
        lngExprParentTableID = objExpr.BaseTableID

        ' Get the expression's parent table name.
        recTabEdit.Index = "idxTableID"
        recTabEdit.Seek "=", lngExprParentTableID

        If Not recTabEdit.NoMatch Then
          sParentTable = recTabEdit!TableName
        Else
          sParentTable = "<unknown>"
        End If
      End If
      ' Disassociate object variables.
      Set objExpr = Nothing
      mfrmUse.AddToList (sExprType & " : " & sExprName & " <" & sParentTable & ">")
      rsExpressions.MoveNext
    Loop
  End If
  ' Close temporary recordset.
  rsExpressions.Close
  

  ' Find any columns that use this table
  ' as the a lookup table.
  sSQL = "SELECT DISTINCT tmpColumns.columnName, tmpColumns.tableID" & _
    " FROM tmpColumns" & _
    " WHERE tmpColumns.deleted = FALSE" & _
    " AND tmpColumns.lookupTableID = " & Trim(Str(gLngTableID)) & _
    " AND tmpColumns.columnType = " & Trim(Str(giCOLUMNTYPE_LOOKUP)) & _
    " AND tmpColumns.tableID <> " & Trim(Str(gLngTableID))
  Set rsColumns = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  If Not (rsColumns.BOF And rsColumns.EOF) Then
    fUsed = True
    Do Until rsColumns.EOF
      ' Get the expression's parent table name.
      recTabEdit.Index = "idxTableID"
      recTabEdit.Seek "=", rsColumns.Fields("tableID")
      If Not recTabEdit.NoMatch Then
        sParentTable = recTabEdit!TableName
      Else
        sParentTable = "<unknown>"
      End If
      mfrmUse.AddToList ("Lookup Table : " & rsColumns.Fields("columnName") & " <" & sParentTable & ">")
      rsColumns.MoveNext
    Loop
  End If
  ' Close temporary recordset.
  rsColumns.Close

  
  ' Check that the table is not used in any Module definitions.
  sSQL = "SELECT DISTINCT moduleKey, parameterkey" & _
    " FROM tmpModuleSetup" & _
    " WHERE parameterType = '" & gsPARAMETERTYPE_TABLEID & "'" & _
    " AND parameterValue = '" & Trim(Str(gLngTableID)) & "'"
  Set rsModules = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  If Not (rsModules.BOF And rsModules.EOF) Then
    fUsed = True
    Do Until rsModules.EOF
      If rsModules!moduleKey = gsMODULEKEY_TRAININGBOOKING And (Not bTR) Then
          bTR = True
          sModuleName = "Training Booking"
          mfrmUse.AddToList (sModuleName & " Configuration")
      ElseIf ((rsModules!moduleKey = gsMODULEKEY_PERSONNEL) Or (rsModules!moduleKey = gsMODULEKEY_HIERARCHY)) And (Not bPE) Then
          bPE = True
          sModuleName = "Personnel"
          mfrmUse.AddToList (sModuleName & " Configuration")
      ElseIf (rsModules!moduleKey = gsMODULEKEY_WORKFLOW) And (Not bWF) Then
        If Application.WorkflowModule Then
          bWF = True
          sModuleName = "Workflow"
          mfrmUse.AddToList (sModuleName & " Configuration")
        End If
      ElseIf rsModules!moduleKey = gsMODULEKEY_ABSENCE Then
          
          If (rsModules!parameterkey = gsPARAMETERKEY_BHOLREGIONTABLE) _
            Or (rsModules!parameterkey = gsPARAMETERKEY_BHOLREGION) _
            Or (rsModules!parameterkey = gsPARAMETERKEY_BHOLTABLE) _
            Or (rsModules!parameterkey = gsPARAMETERKEY_BHOLDATE) _
            Or (rsModules!parameterkey = gsPARAMETERKEY_BHOLDESCRIPTION) Then
            sModuleName = "Bank Holiday"
            If Not bBA Then
              bBA = True
              mfrmUse.AddToList (sModuleName & " Configuration")
            End If
          Else
            sModuleName = "Absence"
            If Not bAB Then
              bAB = True
              mfrmUse.AddToList (sModuleName & " Configuration")
            End If
          End If
          
      ElseIf rsModules!moduleKey = gsMODULEKEY_CURRENCY And (Not bCU) Then
          bCU = True
          sModuleName = "Currency"
          mfrmUse.AddToList (sModuleName & " Configuration")
      ElseIf rsModules!moduleKey = gsMODULEKEY_POST And (Not bPO) Then
          bPO = True
          sModuleName = "Post"
          mfrmUse.AddToList (sModuleName & " Configuration")
      ElseIf rsModules!moduleKey = gsMODULEKEY_MATERNITY And (Not bMA) Then
          bMA = True
          sModuleName = "Maternity"
          mfrmUse.AddToList (sModuleName & " Configuration")
      ElseIf rsModules!moduleKey = gsMODULEKEY_AUDIT And (Not bAu) Then
          bAu = True
          sModuleName = "Audit"
          mfrmUse.AddToList (sModuleName & " Configuration")
      ElseIf rsModules!moduleKey = gsMODULEKEY_SSINTRANET And (Not bSS) Then
          bSS = True
          sModuleName = "Self Service Intranet"
          mfrmUse.AddToList (sModuleName & " Configuration")
      End If
      
      rsModules.MoveNext
    Loop
  End If
  ' Close the recordset.
  rsModules.Close
  
  
  sSQL = "SELECT COUNT(*) AS result" & _
    " FROM tmpSSIViews" & _
    " WHERE tableID = " & Trim(Str(gLngTableID))
  Set rsModules = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  If (rsModules!result > 0) Then
    mfrmUse.AddToList "Module : Self-service Intranet links"
    fUsed = True
  End If
  rsModules.Close
  Set rsModules = Nothing

  
  ' Check that the table is not used in any orders/find column collections (TEMP TABLES).
  sSQL = "SELECT DISTINCT tmpOrders.orderID, tmpOrders.name, tmpTables.tableName, tmpOrders.Type " & _
    " FROM tmpOrders, tmpOrderItems, tmpTables, tmpColumns " & _
    " WHERE tmpOrders.deleted = FALSE" & _
    "   AND tmpTables.deleted = FALSE" & _
    "   AND tmpOrders.orderID = tmpOrderItems.orderID " & _
    "   AND tmpOrderItems.ColumnID = tmpColumns.ColumnID " & _
    "   AND tmpOrders.tableID = tmpTables.TableID " & _
    "   AND tmpColumns.TableID = " & Trim(Str(gLngTableID))
  Set rsOrders = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  If Not (rsOrders.BOF And rsOrders.EOF) Then
    fUsed = True
    Do Until rsOrders.EOF
      strSYSOrderIDs = strSYSOrderIDs & IIf(Len(strSYSOrderIDs) > 0, ",", "") & rsOrders!OrderID
            
      If rsOrders!Type = 1 Then
        mfrmUse.AddToList ("Order : " & rsOrders!Name & " <" & rsOrders!TableName & ">")
      Else
        mfrmUse.AddToList ("Order : " & rsOrders!Name & " <" & rsOrders!TableName & ">")
      End If
      
      rsOrders.MoveNext
    Loop
  End If
  'Close temporary recordset
  rsOrders.Close
  
  
  ' ### We now need to check every single util/report/tool to see if the Table is used.
  
  ' NHRD09112010 JIRA HRPRO-1083
  ' NHRD07012011 JIRA HRPRO-1083 - had to change search value from parent to [TargetTableID]
  ' Check if the column is used in any Document Types
  sSQL = "SELECT DISTINCT DocumentMapID, Name " & _
          "FROM ASRSysDocumentManagementTypes " & _
          "WHERE ASRSysDocumentManagementTypes.[TargetTableID] = " & gLngTableID

  Set rsUtilsDocTypes = New ADODB.Recordset
  
  rsUtilsDocTypes.Open sSQL, gADOCon, adOpenStatic, adLockReadOnly
  With rsUtilsDocTypes
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Document Type : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsUtilsDocTypes = Nothing
  
  ' Now Check if the column is used in a Custom Report
  sSQL = "SELECT DISTINCT ASRSysCustomReportsName.Name, ASRSysCustomReportsName.ID, ASRSysCustomReportsName.Username" & _
         " FROM ASRSysCustomReportsName" & _
         "    LEFT OUTER JOIN ASRSysCustomReportsChildDetails" & _
         "    ON ASRSysCustomReportsName.ID = ASRSysCustomReportsChildDetails.CustomReportID" & _
         " WHERE ASRSysCustomReportsChildDetails.ChildTable = " & Trim(Str(gLngTableID)) & _
         "   OR ASRSysCustomReportsName.BaseTable = " & Trim(Str(gLngTableID)) & _
         "   OR ASRSysCustomReportsName.Parent1Table = " & Trim(Str(gLngTableID)) & _
         "   OR ASRSysCustomReportsName.Parent2Table = " & Trim(Str(gLngTableID))
  Set rsUtils1 = New ADODB.Recordset
  rsUtils1.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  
  With rsUtils1
    If Not (rsUtils1.BOF And rsUtils1.EOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Custom Report : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsUtils1 = Nothing
  
  
  ' Now Check if the column is used in a Calendar Report
  sSQL = "SELECT DISTINCT ASRSysCalendarReports.Name, ASRSysCalendarReports.Username, ASRSysCalendarReports.ID" & _
        " FROM ASRSysCalendarReports" & _
        "     LEFT OUTER JOIN ASRSysCalendarReportEvents " & _
        "     ON ASRSysCalendarReports.ID = ASRSysCalendarReportEvents.calendarReportID" & _
        " WHERE  ASRSysCalendarReports.BaseTable = " & Trim(Str(gLngTableID)) & " " & _
        "     OR ASRSysCalendarReportEvents.TableID = " & Trim(Str(gLngTableID)) & " " & _
        "     OR ASRSysCalendarReportEvents.LegendLookupTableID = " & Trim(Str(gLngTableID))
  Set rsUtils2 = New ADODB.Recordset
  rsUtils2.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsUtils2
   If Not (.BOF And .EOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Calendar Report : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsUtils2 = Nothing
  
 
  ' Now Check if the column is used in a Cross Tab
  sSQL = "SELECT DISTINCT ASRSysCrossTab.CrossTabID, ASRSysCrossTab.Name, ASRSysCrossTab.Username" & _
         " FROM ASRSysCrossTab " & _
         " WHERE ASRSysCrossTab.TableID = " & Trim(Str(gLngTableID)) & _
         " AND CrossTabType <> " & ctt9GridBox
    
  Set rsUtils3 = New ADODB.Recordset
  rsUtils3.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsUtils3
    If Not (rsUtils3.BOF And rsUtils3.EOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Cross Tab : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsUtils3 = Nothing
 
   ' Now Check if the column is used in a 9-Box Grid Reports
  sSQL = "SELECT DISTINCT ASRSysCrossTab.CrossTabID, ASRSysCrossTab.Name, ASRSysCrossTab.Username" & _
         " FROM ASRSysCrossTab " & _
         " WHERE ASRSysCrossTab.TableID = " & Trim(Str(gLngTableID)) & _
         " AND CrossTabType = " & ctt9GridBox
    
  Set rsUtils13 = New ADODB.Recordset
  rsUtils13.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsUtils13
    If Not (rsUtils13.BOF And rsUtils13.EOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("9-Box Grid Report : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsUtils13 = Nothing
 
  ' Now Check if the column is used in a Record Profile
  sSQL = "SELECT DISTINCT ASRSysRecordProfileName.Name, ASRSysRecordProfileName.Username, ASRSysRecordProfileName.recordProfileID " & _
        "FROM ASRSysRecordProfileName " & _
        "   LEFT OUTER JOIN ASRSysRecordProfileDetails " & _
        "   ON ASRSysRecordProfileName.recordProfileID = ASRSysRecordProfileDetails.recordProfileID " & _
        "WHERE ASRSysRecordProfileName.BaseTable = " & Trim(Str(gLngTableID)) & " " & _
        " OR  ASRSysRecordProfileDetails.TableID = " & Trim(Str(gLngTableID)) & " "
  Set rsUtils4 = New ADODB.Recordset
  rsUtils4.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsUtils4
    If Not (.BOF And .EOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Record Profile : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsUtils4 = Nothing
 
 
  'Now Check if the column is used in a Match Report/Career Progression/Succession Planning
  sSQL = "SELECT DISTINCT ASRSysMatchReportName.Name, ASRSysMatchReportName.Username, ASRSysMatchReportName.matchReportID, ASRSysMatchReportName.matchReportType" & _
    " FROM ASRSysMatchReportName" & _
    "     LEFT OUTER JOIN ASRSysMatchReportTables " & _
    "     ON ASRSysMatchReportName.matchReportID = ASRSysMatchReportTables.matchReportID" & _
    " WHERE ASRSysMatchReportName.Table1ID = " & Trim(Str(gLngTableID)) & _
    "   OR ASRSysMatchReportName.Table2ID = " & Trim(Str(gLngTableID)) & _
    "   OR ASRSysMatchReportTables.Table1ID = " & Trim(Str(gLngTableID)) & _
    "   OR ASRSysMatchReportTables.Table2ID = " & Trim(Str(gLngTableID))
  Set rsUtils5 = New ADODB.Recordset
  rsUtils5.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsUtils5
   If Not (.BOF And .EOF) Then
      fUsed = True
      Do Until .EOF
        Select Case !matchReportType
          Case 1
            mfrmUse.AddToList ("Succession Planning : " & !Name)
          Case 2
            mfrmUse.AddToList ("Career Progression : " & !Name)
          Case Else
            mfrmUse.AddToList ("Match Report : " & !Name)
        End Select
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsUtils5 = Nothing
 
 
  ' Now Check if the column is used in a Mail Merge
  sSQL = "SELECT DISTINCT ASRSysMailMergeName.Name, ASRSysMailMergeName.IsLabel, ASRSysMailMergeName.mailMergeID, ASRSysMailMergeName.Username " & _
         "FROM  ASRSysMailMergeName " & _
         "      INNER JOIN ASRSysMailMergeColumns " & _
         "      ON ASRSysMailMergeName.MailMergeID = ASRSysMailMergeColumns.MailMergeID " & _
         "      INNER JOIN ASRSysColumns " & _
         "      ON ASRSysMailMergeColumns.ColumnID = ASRSysColumns.ColumnID " & _
         "WHERE  ASRSysColumns.TableID = " & Trim(Str(gLngTableID)) & " " & _
         " AND UPPER(ASRSysMailMergeColumns.type) = 'C'"
  Set rsUtils6 = New ADODB.Recordset
  rsUtils6.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsUtils6
    If Not (rsUtils6.BOF And rsUtils6.EOF) Then
      fUsed = True
      Do Until .EOF
        If Not IsNull(!IsLabel) Then
          If !IsLabel Then
            mfrmUse.AddToList ("Envelope & Label : " & !Name)
          Else
            mfrmUse.AddToList ("Mail Merge : " & !Name)
          End If
        Else
          mfrmUse.AddToList ("Mail Merge : " & !Name)
        End If
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsUtils6 = Nothing
 
 
  'Check that the column is not used in any Global Add/Update definitions.
  sSQL = "SELECT DISTINCT ASRSysGlobalFunctions.functionID, ASRSysGlobalFunctions.name, ASRSysGlobalFunctions.type, ASRSysGlobalFunctions.username" & _
        " FROM ASRSysGlobalFunctions " & _
        "       LEFT OUTER JOIN ASRSysGlobalItems " & _
        "       ON ASRSysGlobalFunctions.FunctionID = ASRSysGlobalItems.FunctionID " & _
        " WHERE ASRSysGlobalFunctions.TableID = " & Trim(Str(gLngTableID)) & _
        "   OR ASRSysGlobalFunctions.ChildTableID = " & Trim(Str(gLngTableID)) & _
        "   OR ASRSysGlobalItems.LookupTableID = " & Trim(Str(gLngTableID))
  
  Set rsGlobalFunctions = New ADODB.Recordset
  rsGlobalFunctions.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsGlobalFunctions
    If Not (rsGlobalFunctions.BOF And rsGlobalFunctions.EOF) Then
      Do Until rsGlobalFunctions.EOF
        fUsed = True
        ' Tell the user why the expression cannot be deleted.
        Select Case UCase(!Type)
          Case "U"
            sGlobalFunctionType = "Update"
          Case "A"
            sGlobalFunctionType = "Add"
          Case "D"
            sGlobalFunctionType = "Delete"
          Case Else
            sGlobalFunctionType = "Function"
        End Select
        mfrmUse.AddToList ("Global " & sGlobalFunctionType & " : " & !Name)
        .MoveNext
      Loop
    End If
  .Close
  End With
  Set rsGlobalFunctions = Nothing
    
    
  ' Now Check if the column is used in a Data Transfer
  sSQL = "SELECT DISTINCT ASRSysDataTransferName.Name, ASRSysDataTransferName.DataTransferID, ASRSysDataTransferName.Username " & _
         "FROM  ASRSysDataTransferName " & _
         "        LEFT OUTER JOIN ASRSysDataTransferColumns " & _
         "        ON ASRSysDataTransferName.DataTransferID = ASRSysDataTransferColumns.DataTransferID " & _
         "WHERE ASRSysDataTransferName.FromTableID = " & Trim(Str(gLngTableID)) & _
         "   OR ASRSysDataTransferName.ToTableID = " & Trim(Str(gLngTableID)) & _
         "   OR ASRSysDataTransferColumns.FromTableID = " & Trim(Str(gLngTableID)) & _
         "   OR ASRSysDataTransferColumns.ToTableID = " & Trim(Str(gLngTableID))
  
  Set rsUtils7 = New ADODB.Recordset
  rsUtils7.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsUtils7
    If Not (rsUtils7.BOF And rsUtils7.EOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Data Transfer : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsUtils7 = Nothing
  

  ' Now Check if the column is used in an Import
  sSQL = " SELECT DISTINCT     ASRSysImportName.Name, ASRSysImportName.ID, ASRSysImportName.Username " & _
         " FROM ASRSysImportName " & _
         "      INNER JOIN ASRSysImportDetails " & _
         "      ON ASRSysImportName.ID = ASRSysImportDetails.ImportID " & _
         " WHERE ASRSysImportDetails.TableID = " & Trim(Str(gLngTableID)) & _
         "      OR ASRSysImportName.BaseTable = " & Trim(Str(gLngTableID))
  
  Set rsUtils8 = New ADODB.Recordset
  rsUtils8.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsUtils8
    If Not (rsUtils8.BOF And rsUtils8.EOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Import : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsUtils8 = Nothing


  ' Now Check if the column is used in a Export
  sSQL = "SELECT DISTINCT   ASRSysExportName.Name, ASRSysExportName.ID, ASRSysExportName.Username " & _
         "FROM  ASRSysExportName " & _
         "         LEFT OUTER JOIN ASRSysExportDetails " & _
         "         ON ASRSysExportName.ID = ASRSysExportDetails.ExportID " & _
         "WHERE      ASRSysExportName.BaseTable = " & Trim(Str(gLngTableID)) & " " & _
         "  OR       ASRSysExportName.Parent1Table = " & Trim(Str(gLngTableID)) & " " & _
         "  OR       ASRSysExportName.Parent2Table = " & Trim(Str(gLngTableID)) & " " & _
         "  OR       ASRSysExportName.ChildTable = " & Trim(Str(gLngTableID)) & " " & _
         "  OR       ASRSysExportDetails.TableID = " & Trim(Str(gLngTableID)) & " "
  
  Set rsUtils9 = New ADODB.Recordset
  rsUtils9.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsUtils9
    If Not (rsUtils9.BOF And rsUtils9.EOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Export : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsUtils9 = Nothing


  ' Now Check if the column is used in a Runtime Calculation OR Runtime Filter.
  sSQL = "SELECT DISTINCT ASRSysExpressions.exprID, ASRSysExpressions.Type, ASRSysExpressions.Name, ASRSysTables.TableName " & _
        " FROM ASRSysExpressions " & _
        "       LEFT OUTER JOIN  ASRSysExprComponents " & _
        "       ON ASRSysExpressions.ExprID = ASRSysExprComponents.ExprID " & _
        "       LEFT OUTER JOIN ASRSysTables" & _
        "       ON ASRSysExpressions.TableID = ASRSysTables.TableID" & _
        " WHERE ASRSysExpressions.Type IN (" & giEXPR_RUNTIMECALCULATION & "," & giEXPR_RUNTIMEFILTER & ", " & giEXPR_RECORDINDEPENDANTCALC & ") " & _
        "   AND ASRSysExpressions.ParentComponentID = 0" & _
        "   AND (ASRSysExpressions.TableID = " & Trim(Str(gLngTableID)) & _
        "      OR ASRSysExprComponents.FieldTableID = " & Trim(Str(gLngTableID)) & _
        "      OR ASRSysExprComponents.LookupTableID = " & Trim(Str(gLngTableID)) & ")"
  
  Set rsUtils10 = New ADODB.Recordset
  rsUtils10.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsUtils10
    If Not (rsUtils10.BOF And rsUtils10.EOF) Then
      fUsed = True
      Do Until .EOF
        Select Case !Type
          Case ExpressionTypes.giEXPR_RUNTIMEFILTER
            sExprType = "Runtime Filter"
            mfrmUse.AddToList (sExprType & " : " & !Name & " <" & !TableName & ">")
          Case ExpressionTypes.giEXPR_RUNTIMECALCULATION
            sExprType = "Runtime Calculation"
            mfrmUse.AddToList (sExprType & " : " & !Name & " <" & !TableName & ">")
          Case ExpressionTypes.giEXPR_RECORDINDEPENDANTCALC
            sExprType = "Record Independent Calculation"
            mfrmUse.AddToList (sExprType & " : " & !Name & " <Report Configuration/Calendar Reports>")
          Case Else
            sExprType = "Expression"
        End Select
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsUtils10 = Nothing


  ' Check that the table is not used in any orders/find column collections (SQL TABLES).
  strSYSCheckedOrderIDs = "0"
  sSQL = "SELECT DISTINCT tmpOrders.orderID" & _
    " FROM tmpOrders" & _
    " WHERE ((tmpOrders.deleted = TRUE)" & _
    "   OR (tmpOrders.changed = TRUE))"
  Set rsOrders = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  If Not (rsOrders.BOF And rsOrders.EOF) Then
    Do Until rsOrders.EOF
      strSYSCheckedOrderIDs = strSYSCheckedOrderIDs & "," & rsOrders!OrderID
      rsOrders.MoveNext
    Loop
  End If
  'Close temporary recordset
  rsOrders.Close
  
  sSQL = "SELECT DISTINCT ASRSysOrders.orderID, ASRSysOrders.name, ASRSysTables.tableName" & _
        " FROM ASRSysOrders, ASRSysOrderItems, ASRSysTables, ASRSysColumns" & _
        " WHERE ASRSysOrders.orderID = ASRSysOrderItems.orderID " & _
        "   AND ASRSysOrderItems.ColumnID = ASRSysColumns.ColumnID " & _
        "   AND ASRSysOrders.tableID = ASRSysTables.TableID " & _
        "   AND ASRSysColumns.TableID = " & Trim(Str(gLngTableID)) & _
        "   AND ASRSysOrders.orderID NOT IN (" & strSYSCheckedOrderIDs & ")"
  If Len(strSYSOrderIDs) > 0 Then
    sSQL = sSQL & _
        "   AND ASRSysOrders.OrderID NOT IN (" & strSYSOrderIDs & ")"
  End If
  
  Set rsUtils11 = New ADODB.Recordset
  rsUtils11.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsUtils11
    If Not (rsUtils11.BOF And rsUtils11.EOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Order : " & !Name & " <" & !TableName & ">")
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsUtils11 = Nothing

  
  ' Now Check if the column is used in a Picklist.
  sSQL = "SELECT DISTINCT ASRSysPicklistName.picklistID, ASRSysPicklistName.Name " & _
        " FROM ASRSysPicklistName " & _
        " WHERE ASRSysPicklistName.TableID = " & Trim(Str(gLngTableID))
  Set rsUtils12 = New ADODB.Recordset
  rsUtils12.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsUtils12
    If Not (rsUtils12.BOF And rsUtils12.EOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Picklist : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsUtils12 = Nothing


  'Make check for the columns use in email definitions.
  sSQL = "SELECT DISTINCT tmpEmailAddresses.name " & _
         "FROM tmpEmailAddresses " & _
         "WHERE tmpEmailAddresses.TableID = " & gLngTableID & _
         " AND tmpEmailAddresses.Deleted = FALSE "
  Set rsEmail = daoDb.OpenRecordset(sSQL)
  With rsEmail
    If Not (rsEmail.BOF And rsEmail.EOF) Then
      fUsed = True
      .MoveFirst
      Do Until .EOF
        mfrmUse.AddToList ("Email Address : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsEmail = Nothing

  If Application.WorkflowModule Then
    ' Check that it is not used as by a recordSelector on a Web form.
    sSQL = "SELECT DISTINCT tmpWorkflowElements.workflowID," & _
      "   tmpWorkflowElements.identifier, " & _
      "   tmpWorkflowElementItems.itemType " & _
      " FROM tmpWorkflowElementItems" & _
      " INNER JOIN tmpWorkflowElements ON tmpWorkflowElementItems.elementID = tmpWorkflowElements.id" & _
      " WHERE tmpWorkflowElementItems.tableID = " & Trim(Str(gLngTableID)) & _
      " OR tmpWorkflowElementItems.LookupTableID = " & Trim(Str(gLngTableID))
    Set rsDAOTemp = daoDb.OpenRecordset(sSQL, _
      dbOpenForwardOnly, dbReadOnly)
    If Not (rsDAOTemp.BOF And rsDAOTemp.EOF) Then
      Do Until rsDAOTemp.EOF
        recWorkflowEdit.Index = "idxWorkflowID"
        recWorkflowEdit.Seek "=", rsDAOTemp.Fields("workflowID")
  
        If Not recWorkflowEdit.NoMatch Then
          If recWorkflowEdit.Fields("deleted").value = False Then
            fUsed = True
            mfrmUse.AddToList ("Workflow : " & recWorkflowEdit.Fields("name").value & " <'" & rsDAOTemp.Fields("identifier") & "' web form " & IIf(rsDAOTemp.Fields("itemtype") = 14, "lookup", "record selector") & " table>")
          End If
        End If
  
        rsDAOTemp.MoveNext
      Loop
    End If
    'Close temporary recordset
    rsDAOTemp.Close

    ' Check that it is not used as the Base Table for a Workflow.
    sSQL = "SELECT DISTINCT tmpWorkflows.name" & _
      " FROM tmpWorkflows" & _
      " WHERE tmpWorkflows.deleted = FALSE" & _
      " AND tmpWorkflows.initiationType = " & Trim(Str(WORKFLOWINITIATIONTYPE_TRIGGERED)) & _
      " AND tmpWorkflows.baseTable = " & Trim(Str(gLngTableID))
  
    Set rsDAOTemp = daoDb.OpenRecordset(sSQL, _
      dbOpenForwardOnly, dbReadOnly)
    If Not (rsDAOTemp.BOF And rsDAOTemp.EOF) Then
      Do Until rsDAOTemp.EOF
        fUsed = True
        mfrmUse.AddToList ("Workflow : " & rsDAOTemp.Fields("name"))
        
        rsDAOTemp.MoveNext
      Loop
    End If
    'Close temporary recordset
    rsDAOTemp.Close
  End If
  
  ' NPG20100427 Fault HRPRO-912
  If Application.SelfServiceIntranetModule Then
    ' Check that it is not used in SSI Charting.
    sSQL = "SELECT DISTINCT tmpSSIntranetLinks.ID," & _
      "   tmpSSIntranetLinks.Element_Type," & _
      "   tmpSSIntranetLinks.text" & _
      " FROM tmpSSIntranetLinks" & _
      " WHERE tmpSSIntranetLinks.Chart_tableID = " & Trim(Str(gLngTableID)) & _
      " AND tmpSSIntranetLinks.Element_Type in (2, 4)"
  
    Set rsDAOTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    If Not (rsDAOTemp.BOF And rsDAOTemp.EOF) Then
      fUsed = True
      Do Until rsDAOTemp.EOF
          Select Case rsDAOTemp.Fields("Element_Type")
          Case 2  ' Chart
            mfrmUse.AddToList ("Self Service Intranet Chart : " & rsDAOTemp.Fields("text"))
          Case 4  ' Database Value
            mfrmUse.AddToList ("Self Service Intranet Database Value : " & rsDAOTemp.Fields("text"))
          Case Else
            mfrmUse.AddToList ("Self Service Intranet Configuration : " & rsDAOTemp.Fields("text"))
          End Select
        rsDAOTemp.MoveNext
      Loop
    End If
    'Close temporary recordset
    rsDAOTemp.Close
  End If
 
TidyUpAndExit:
  ' Disassociate object variables.
  Set rsDAOTemp = Nothing
  Set rsModules = Nothing
  Set rsExpressions = Nothing
  Set rsColumns = Nothing
  Set objExpr = Nothing
  Set rsUtils1 = Nothing
  Set rsUtils2 = Nothing
  Set rsUtils3 = Nothing
  Set rsUtils4 = Nothing
  Set rsUtils5 = Nothing
  Set rsUtils6 = Nothing
  Set rsUtils7 = Nothing
  Set rsUtils8 = Nothing
  Set rsUtils9 = Nothing
  Set rsUtils10 = Nothing
  Set rsUtils11 = Nothing
  Set rsUtils12 = Nothing
  Set rsUtils13 = Nothing
  
  TableIsUsed = fUsed
  Exit Function

ErrorTrap:
  fUsed = True
  Resume TidyUpAndExit
  
End Function

Public Property Get PrimaryOrderID() As Long

  ' Return the primary order id.
  PrimaryOrderID = gLngPrimaryOrderID
  
End Property
Public Property Get RecordDescriptionID() As Long

  ' Return the Record Description Expression id.
  RecordDescriptionID = gLngRecordDescriptionID
  
End Property

Public Property Let PrimaryOrderID(plngNewValue As Long)

  ' Set the primary order id.
  gLngPrimaryOrderID = plngNewValue
  
End Property
Public Property Let RecordDescriptionID(plngNewValue As Long)

  ' Set the Record Description Expression ID.
  gLngRecordDescriptionID = plngNewValue
  
End Property

Private Sub Class_Initialize()
  Set gobjSummaryFields = New Collection
  Set mcolEmailLinks = New Collection
  Set mcolOutlookLinks = New Collection
  Set mcolWorkflowTriggeredLinks = New Collection

  miGrantRead = 0
  miGrantEdit = 0
  miGrantNew = 0
  miGrantDelete = 0
  mbPermissionsPrompted = False
  
End Sub


Private Sub Class_Terminate()
  Set gobjSummaryFields = Nothing
  Set mcolEmailLinks = Nothing
  Set mcolOutlookLinks = Nothing
  Set mcolWorkflowTriggeredLinks = Nothing
  
End Sub


Public Function CloneTable_Transaction(pfCopyData As Boolean, pfCopySecurity As Boolean _
  , pbGrantRead As Boolean _
  , pbGrantNew As Boolean _
  , pbGrantEdit As Boolean _
  , pbGrantDelete As Boolean) As Table
  
  ' Transaction wrapper for the 'CloneTable' function.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim sSQL As String
  Dim objCloneTable As Table
  Dim rsParents As DAO.Recordset
  Dim rsColumns As DAO.Recordset
  Dim avCloneRegister() As Variant
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  Screen.MousePointer = vbHourglass
  
  ReDim avCloneRegister(3, 0)

  ' If the table being cloned has any parent tables, then remember their IDs
  ' and their column IDs in the clone register.
  sSQL = "SELECT tmpRelations.parentID" & _
    " FROM tmpRelations" & _
    " WHERE tmpRelations.childID = " & Trim(Str(gLngTableID))
  Set rsParents = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
            
  With rsParents
    ' For each parent table definition ...
    Do While (Not .EOF)
      ' Remember the IDs of the parent tables.
      iIndex = UBound(avCloneRegister, 2) + 1
      ReDim Preserve avCloneRegister(3, iIndex)
      avCloneRegister(1, iIndex) = "TABLE"
      avCloneRegister(2, iIndex) = .Fields("parentID")
      avCloneRegister(3, iIndex) = .Fields("parentID")
      
      sSQL = "SELECT tmpColumns.columnID" & _
        " FROM tmpColumns" & _
        " WHERE tmpColumns.deleted = FALSE" & _
        " AND tmpColumns.tableID = " & Trim(Str(.Fields("parentID")))
      Set rsColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
            
      With rsColumns
        ' For each parent table definition ...
        Do While (Not .EOF)
          ' Remember the IDs of the parent tables.
          iIndex = UBound(avCloneRegister, 2) + 1
          ReDim Preserve avCloneRegister(3, iIndex)
          avCloneRegister(1, iIndex) = "COLUMN"
          avCloneRegister(2, iIndex) = .Fields("columnID")
          avCloneRegister(3, iIndex) = .Fields("columnID")
                    
          .MoveNext
        Loop
                  
        .Close
      End With
      ' Disassociate object variables.
      Set rsColumns = Nothing
  
      .MoveNext
    Loop
              
    .Close
  End With
  ' Disassociate object variables.
  Set rsParents = Nothing

  mstrExprNotCopied = vbNullString

  ' Clone the table.
  Set objCloneTable = CloneTable(avCloneRegister, pfCopyData, pfCopySecurity, pbGrantRead, pbGrantNew, pbGrantEdit, pbGrantDelete)
  fOK = Not objCloneTable Is Nothing

  If fOK Then
    ' Ensure that the relationships between the cloned table and its parents
    ' are set up.
    sSQL = "SELECT tmpRelations.parentID" & _
      " FROM tmpRelations" & _
      " WHERE tmpRelations.childID = " & Trim(Str(gLngTableID))
    Set rsParents = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
    With rsParents
      ' For each parent table definition ...
      Do While (Not .EOF) And fOK
        With recRelEdit
          ' Create the relationship record in the database.
          .AddNew
          !parentID = rsParents.Fields("parentID")
          !childID = objCloneTable.TableID
          .Update
  
          ' Create the column in the child table that links it to the ID column in the
          ' parent table.
          With objCloneTable
            If .AddIDColumn(rsParents.Fields("parentID")) Then
              fOK = .WriteTable
            End If
          End With
        End With
  
        .MoveNext
      Loop
  
      .Close
    End With
    ' Disassociate object variables.
    Set rsParents = Nothing
  End If
  
TidyUpAndExit:
  Screen.MousePointer = vbDefault
    
  Set rsParents = Nothing
  Set rsColumns = Nothing
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True

    'MH20010402
    If mstrExprNotCopied <> vbNullString Then
      gobjProgress.Visible = False
      MsgBox "The following expression(s) could not be copied as they rely upon module setup:" & vbCrLf & vbCrLf & _
             mstrExprNotCopied, vbInformation, "Copy Tables"
    End If

  Else
    daoWS.Rollback
    gobjProgress.Visible = False
    MsgBox "Error copying table.", _
      vbExclamation + vbOKOnly, Application.Name
  End If
  Set CloneTable_Transaction = objCloneTable
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function
Public Function DeleteAllRelations_Transaction() As Boolean
  ' Transaction wrapper for the 'DeleteAllRelations' function.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  fOK = DeleteAllRelations
  
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  DeleteAllRelations_Transaction = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function DeleteRelation_Transaction(ByVal plngRelatedTableID As Long) As Boolean
  ' Transaction wrapper for the 'DeleteRelation' function.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  fOK = DeleteRelation(plngRelatedTableID)
  
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  DeleteRelation_Transaction = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Property Get IsCopy() As Boolean
  IsCopy = mbIsCopy
End Property

Public Property Let IsCopy(pbNewValue As Boolean)
  mbIsCopy = pbNewValue
End Property

Public Property Get CopyTableID() As Long
  CopyTableID = mlngCopyTableID
End Property

Public Property Let CopyTableID(plngNewValue As Long)
  mlngCopyTableID = plngNewValue
End Property

Public Property Get CopyTableName() As String
  CopyTableName = msCopyTableName
End Property

Public Property Let CopyTableName(psNewValue As String)
  msCopyTableName = psNewValue
End Property

Public Property Get ManualSummaryColumnBreaks() As Boolean
    ManualSummaryColumnBreaks = mbManualColumnBreaks
End Property

Public Property Let ManualSummaryColumnBreaks(ByVal bNewValue As Boolean)
  mbManualColumnBreaks = bNewValue
End Property

Public Function PrintDefinition(piOutputType As SystemMgr.OutputDefintionTypes) As Boolean
  On Error GoTo ErrorTrap

  Dim bOK As Boolean
  Dim iCount As Integer
  
  bOK = True
  
  ' Edit the table's properties.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim frmEdit As frmTabEdit
  
  fOK = ReadTable
  
  ' Find & read table record to edit
  If fOK Then
  
    Set frmEdit = New frmTabEdit
    Set frmEdit.Table = Me
    
    Select Case piOutputType
    
      Case giEXPORT_TO_PRINTER
        frmEdit.PrintDefinition
        
      Case giEXPORT_TO_CLIPBOARD
        frmEdit.CopyDefinitionToClipboard
    
      Case giEXPORT_TO_WORD
    
    End Select
    
    'JPD 20031008 Fault 7174
    UnLoad frmEdit
  End If
  
TidyUpAndExit:
  Set frmEdit = Nothing
  PrintDefinition = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Public Property Get AuditInsert() As Boolean
  AuditInsert = mbAuditInsert
End Property
Public Property Let AuditInsert(ByVal bNewValue As Boolean)
  mbAuditInsert = bNewValue
End Property
Public Property Get AuditDelete() As Boolean
  AuditDelete = mbAuditDelete
End Property
Public Property Let AuditDelete(ByVal bNewValue As Boolean)
  mbAuditDelete = bNewValue
End Property
'Public Property Get EmailInsertID() As Long
'  EmailInsertID = mlngEmailInsert
'End Property
'Public Property Let EmailInsertID(ByVal lngNewValue As Long)
'  mlngEmailInsert = lngNewValue
'End Property
'Public Property Get EmailDeleteID() As Long
'  EmailDeleteID = mlngEmailDelete
'End Property
'Public Property Let EmailDeleteID(ByVal lngNewValue As Long)
'  mlngEmailDelete = lngNewValue
'End Property

Private Function WriteTableValidations() As Boolean

  On Error GoTo ErrorTrap

  Dim objValidation As clsTableValidation
  Dim strSQL As String
  Dim fOK As Boolean

  fOK = True

  strSQL = "DELETE FROM tmpTableValidations WHERE tableID = " & CStr(gLngTableID)
  daoDb.Execute strSQL, dbFailOnError

  ' Write the column control values to the database.
  If mcolTableValidations.Count > 0 Then
    ' Add records for each Workflow Triggered Link in the collection.
    For Each objValidation In mcolTableValidations
      objValidation.TableID = gLngTableID
      fOK = objValidation.WriteValidation

      If Not fOK Then
        Exit For
      End If
    Next objValidation
    Set objValidation = Nothing
  End If
  
TidyUpAndExit:
  WriteTableValidations = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit


End Function

Private Function WriteTableTriggers() As Boolean

  On Error GoTo ErrorTrap

  Dim objTrigger As clsTableTrigger
  Dim strSQL As String
  Dim fOK As Boolean

  fOK = True

  strSQL = "DELETE FROM tmpTableTriggers WHERE tableID = " & CStr(gLngTableID)
  daoDb.Execute strSQL, dbFailOnError

  ' Write the column control values to the database.
  If mcolTableTriggers.Count > 0 Then
    For Each objTrigger In mcolTableTriggers
      objTrigger.TableID = gLngTableID
      fOK = objTrigger.WriteTrigger

      If Not fOK Then
        Exit For
      End If
    Next objTrigger
    Set objTrigger = Nothing
  End If
  
TidyUpAndExit:
  WriteTableTriggers = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Private Function WriteWorkflowLinks() As Boolean
  On Error GoTo ErrorTrap

  Dim objLink As clsWorkflowTriggeredLink
  Dim strSQL As String
  Dim fOK As Boolean

  fOK = True

  strSQL = "DELETE FROM tmpWorkflowTriggeredLinks WHERE tableID = " & CStr(gLngTableID)
  daoDb.Execute strSQL, dbFailOnError

  ' Write the column control values to the database.
  If mcolWorkflowTriggeredLinks.Count > 0 Then
    ' Add records for each Workflow Triggered Link in the collection.
    For Each objLink In mcolWorkflowTriggeredLinks
      objLink.TableID = gLngTableID
      fOK = objLink.WriteLink

      If Not fOK Then
        Exit For
      End If
    Next objLink
    Set objLink = Nothing
  End If
  
TidyUpAndExit:
  WriteWorkflowLinks = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function


Private Function WriteEmailLinks() As Boolean
  On Error GoTo ErrorTrap

  Dim objLink As clsEmailLink
  Dim strSQL As String
  Dim strSearch As String
  Dim fOK As Boolean
  
  fOK = True
  
  strSQL = "DELETE FROM tmpEmailLinks WHERE TableID = " & CStr(gLngTableID)
  daoDb.Execute strSQL, dbFailOnError
  
  ' Write the column control values to the database.
  If mcolEmailLinks.Count > 0 Then
    ' Add records for each Email Link in the collection.
    For Each objLink In mcolEmailLinks
      objLink.TableID = gLngTableID
      fOK = objLink.WriteEMailLink
      
      If Not fOK Then
        Exit For
      End If
    Next objLink
    Set objLink = Nothing
  End If
  
TidyUpAndExit:
  WriteEmailLinks = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function


Private Function WriteOutlookLinks() As Boolean
  On Error GoTo ErrorTrap

  Dim objLink As clsOutlookLink
  Dim strSQL As String
  Dim strSearch As String
  Dim fOK As Boolean
  
  fOK = True
  
  strSQL = "DELETE FROM tmpOutlookLinks WHERE TableID = " & CStr(gLngTableID)
  daoDb.Execute strSQL, dbFailOnError
  
  ' Write the column control values to the database.
  If mcolOutlookLinks.Count > 0 Then
    ' Add records for each Email Link in the collection.
    For Each objLink In mcolOutlookLinks
      objLink.TableID = gLngTableID
      fOK = objLink.WriteLink
      
      If Not fOK Then
        Exit For
      End If
    Next objLink
    Set objLink = Nothing
  End If
  
TidyUpAndExit:
  WriteOutlookLinks = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function



Public Property Get EmailLinks() As Collection
  Set EmailLinks = mcolEmailLinks
End Property

Public Property Get OutlookLinks() As Collection
  Set OutlookLinks = mcolOutlookLinks
End Property

Public Property Get WorkflowTriggeredLinks() As Collection
  Set WorkflowTriggeredLinks = mcolWorkflowTriggeredLinks
End Property

Public Property Get TableValidations() As Collection
  Set TableValidations = mcolTableValidations
End Property

Public Property Get TableTriggers() As Collection
  Set TableTriggers = mcolTableTriggers
End Property

Public Property Let EmailLinks(ByVal colNewValue As Collection)
  Set mcolEmailLinks = colNewValue
End Property

Public Property Let OutlookLinks(ByVal colNewValue As Collection)
  Set mcolOutlookLinks = colNewValue
End Property

Public Property Let WorkflowTriggeredLinks(ByVal colNewValue As Collection)
  Set mcolWorkflowTriggeredLinks = colNewValue
End Property

Public Property Let TableValidations(ByVal colNewValue As Collection)
  Set mcolTableValidations = colNewValue
End Property

Public Property Let TableTriggers(ByVal colNewValue As Collection)
  Set mcolTableTriggers = colNewValue
End Property

Public Function ReadEmailLinks()

  Dim objEmailLink As clsEmailLink

  On Error GoTo ErrorTrap
  
  Set mcolEmailLinks = Nothing
  Set mcolEmailLinks = New Collection
  
  With recEmailLinksEdit
    .Index = "idxTableID"
    .Seek "=", gLngTableID

    If Not .NoMatch Then

      Do While !TableID = gLngTableID

        If Not !Deleted Then
          Set objEmailLink = New clsEmailLink
          objEmailLink.ReadEmailLink
          mcolEmailLinks.Add objEmailLink, "ID" & CStr(objEmailLink.LinkID)
          Set objEmailLink = Nothing
        End If

        .MoveNext
        If .EOF Then
          Exit Do
        End If
      Loop

    End If

  End With

  ReadEmailLinks = True
  Exit Function

ErrorTrap:
  ReadEmailLinks = False

End Function


Public Function ReadOutlookLinks()

  Dim objOutlookLink As clsOutlookLink

  On Error GoTo ErrorTrap
  
  Set mcolOutlookLinks = Nothing
  Set mcolOutlookLinks = New Collection
  
  With recOutlookLinks
    .Index = "idxTableID"
    .Seek "=", gLngTableID

    If Not .NoMatch Then

      Do While !TableID = gLngTableID

        If Not !Deleted Then
          Set objOutlookLink = New clsOutlookLink
          objOutlookLink.ReadLink
          mcolOutlookLinks.Add objOutlookLink, "ID" & CStr(objOutlookLink.LinkID)
          Set objOutlookLink = Nothing
        End If

        .MoveNext
        If .EOF Then
          Exit Do
        End If
      Loop

    End If

  End With

  ReadOutlookLinks = True
  Exit Function

ErrorTrap:
  ReadOutlookLinks = False

End Function


Public Function ReadWorkflowLinks() As Boolean

  Dim objLink As clsWorkflowTriggeredLink
  
  On Error GoTo ErrorTrap

  Set mcolWorkflowTriggeredLinks = Nothing
  Set mcolWorkflowTriggeredLinks = New Collection

  With recWorkflowTriggeredLinks
    .Index = "idxTableID"
    .Seek "=", gLngTableID

    If Not .NoMatch Then
      Do While !TableID = gLngTableID
        If Not !Deleted Then
          Set objLink = New clsWorkflowTriggeredLink
          objLink.ReadLink
          mcolWorkflowTriggeredLinks.Add objLink, "ID" & CStr(objLink.LinkID)
          Set objLink = Nothing
        End If

        .MoveNext
        If .EOF Then
          Exit Do
        End If
      Loop
    End If
  End With

  ReadWorkflowLinks = True
  Exit Function

ErrorTrap:
  ReadWorkflowLinks = False

End Function

Public Function ReadTableValidations() As Boolean

  Dim objValidation As clsTableValidation
  
  On Error GoTo ErrorTrap

  Set mcolTableValidations = New Collection

  With recTableValidationEdit
    .Index = "idxTableID"
    .Seek "=", gLngTableID

    If Not .NoMatch Then
      Do While !TableID = gLngTableID
        If Not !Deleted Then
          Set objValidation = New clsTableValidation
          objValidation.ValidationID = .Fields("ValidationID").value
          objValidation.ReadValidation
          mcolTableValidations.Add objValidation, "ID" & CStr(objValidation.ValidationID)
          Set objValidation = Nothing
        End If

        .MoveNext
        If .EOF Then
          Exit Do
        End If
      Loop
    End If
  End With

  ReadTableValidations = True
  Exit Function

ErrorTrap:
  ReadTableValidations = False

End Function


Public Function ReadTableTriggers() As Boolean

  Dim objTrigger As clsTableTrigger
  
  On Error GoTo ErrorTrap

  Set mcolTableTriggers = New Collection

  With recTableTriggerEdit
    .Index = "idxTableID"
    .Seek "=", gLngTableID

    If Not .NoMatch Then
      Do While !TableID = gLngTableID
        If Not !Deleted Then
          Set objTrigger = New clsTableTrigger
          objTrigger.TriggerID = .Fields("TriggerID").value
          objTrigger.ReadTrigger
          mcolTableTriggers.Add objTrigger, "ID" & CStr(objTrigger.TriggerID)
          Set objTrigger = Nothing
        End If

        .MoveNext
        If .EOF Then
          Exit Do
        End If
      Loop
    End If
  End With

  ReadTableTriggers = True
  Exit Function

ErrorTrap:
  ReadTableTriggers = False

End Function

Public Property Get IsRemoteView() As Boolean
  IsRemoteView = mbIsRemoteView
End Property

Public Property Let IsRemoteView(pbNewValue As Boolean)
  mbIsRemoteView = pbNewValue
End Property

Public Property Get Locked() As Boolean
  Locked = mbLocked
End Property
