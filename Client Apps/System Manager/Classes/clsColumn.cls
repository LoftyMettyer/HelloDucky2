VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Column"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Properties"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' Local variables to hold property values
Private gLngColumnID As Long
Private gLngTableID As Long
Private gsControlValues As String

' Local variable to hold collection
Private gMVarProperties As SystemMgr.Properties
Private mobjDiaryLinks As Collection
'Private mobjEmailLinks As Collection

Private mfrmUse As frmUsage
Public Property Get ColumnID() As Long
  
  ' Return the current column id.
  ColumnID = gLngColumnID
  
End Property

Public Property Let ColumnID(pLngNewID As Long)
  
  ' Set the current column id.
  gLngColumnID = pLngNewID
  
  If Properties("columnID") Is Nothing Then
    Properties.Add "columnID"
  End If
  
  Properties("columnID").value = gLngColumnID

End Property

Public Property Get TableID() As Long

  ' Return the current table id.
  TableID = gLngTableID
  
  If Properties("tableID") Is Nothing Then
    TableID = 0
  Else
    TableID = Properties("tableID").value
  End If

End Property

Public Property Let TableID(plngID As Long)

  ' Set the current table id.
  gLngTableID = plngID
  
  If Properties("tableID") Is Nothing Then
    Properties.Add "tableID"
  End If
  
  Properties("tableID").value = gLngTableID

End Property

Public Property Get Properties() As SystemMgr.Properties

  ' Return the current instance of the properties class.
  If gMVarProperties Is Nothing Then
    Set gMVarProperties = New SystemMgr.Properties
  End If
  
  Set Properties = gMVarProperties
  
End Property

Public Property Set Properties(pvData As SystemMgr.Properties)
  ' Set the properties class property.
  Set gMVarProperties = pvData
  
End Property

Public Property Get IsNew() As Boolean
  
  ' Return the 'new column' flag.
  If Properties("New") Is Nothing Then
    IsNew = False
  Else
    IsNew = Properties("New").value
  End If
  
End Property

Public Property Let IsNew(pfSetting As Boolean)
  
  ' Set the 'new column' flag.
  If Properties("New") Is Nothing Then
    Properties.Add "New"
  End If
  
  Properties("New").value = pfSetting

End Property

Public Property Get IsChanged() As Boolean
  
  ' Return the 'changed' flag.
  If Properties("Changed") Is Nothing Then
    IsChanged = False
  Else
    IsChanged = Properties("Changed").value
  End If

End Property

Public Property Let IsChanged(pfSetting As Boolean)
  
  ' Set the 'changed' flag.
  If Properties("Changed") Is Nothing Then
    Properties.Add "Changed"
  End If
  
  Properties("Changed").value = pfSetting
  
End Property

Public Property Get IsDeleted() As Boolean
  
  ' Return the 'deleted' flag.
  If Properties("deleted") Is Nothing Then
    IsDeleted = False
  Else
    IsDeleted = Properties("deleted").value
  End If

End Property

Public Property Let IsDeleted(pfSetting As Boolean)
  
  ' Set the 'deleted' flag.
  If Properties("deleted") Is Nothing Then
    Properties.Add "deleted"
  End If
  
  Properties("deleted").value = pfSetting
  
End Property

Public Function DeleteColumn() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  
  'Find & read column record to delete
  fOK = ReadColumn
  
  'TM19012004
  ' Check that the column is not used.
  If fOK Then
    Set mfrmUse = New frmUsage
    mfrmUse.ResetList
    If ColumnIsUsed(mfrmUse) Then
      Screen.MousePointer = vbDefault
      mfrmUse.ShowMessage GetTableName(Me.TableID) & "." & Me.Properties("ColumnName").value & " Column", "The column cannot be deleted as the column is used by the following:", UsageCheckObject.Column
      fOK = False
    End If
    UnLoad mfrmUse
    Set mfrmUse = Nothing
  End If
  
  If fOK Then
'    'Find & read column record to delete
'    fOK = ReadColumn
    
    If fOK Then
      'Check if column is already deleted
      If Not IsDeleted Then
        'Set record deleted flag
        IsDeleted = True
        
        'Write column record
        fOK = WriteColumn
        
        ' Delete any other table records that hang off the column record.
        If fOK Then
          ' Delete any associated column control value records.
          sSQL = "DELETE FROM tmpControlValues" & _
            " WHERE tmpControlValues.columnID = " & Trim(Str(gLngColumnID))
          daoDb.Execute sSQL
        
          ' Delete any associated diary link records.
          fOK = ClearDiaryLinks
          'fOK = ClearEmailLinks
        End If
        
        If fOK Then
          sSQL = "DELETE FROM tmpDiary" & _
            " WHERE tmpDiary.columnID = " & CStr(gLngColumnID)
          daoDb.Execute sSQL

'          sSQL = "DELETE FROM tmpEmailLinksRecipients " & _
'            " WHERE tmpEmailLinksRecipients.LinkID IN " & _
'            "(SELECT LinkID FROM tmpEmailLinks " & _
'            " WHERE tmpEmailLinks.ColumnID = " & CStr(gLngColumnID) & ")"
'          daoDb.Execute sSQL
'
'          'sSQL = "DELETE FROM tmpEmailLinksAttachments " & _
'            " WHERE tmpEmailLinksAttachments.LinkID IN " & _
'            "(SELECT LinkID FROM tmpEmailLinks " & _
'            " WHERE tmpEmailLinks.ColumnID = " & CStr(gLngColumnID) & ")"
'          'daoDb.Execute sSQL
'
'          sSQL = "DELETE FROM tmpEmailLinks" & _
'            " WHERE tmpEmailLinks.ColumnID = " & CStr(gLngColumnID)
'          daoDb.Execute sSQL
        End If
        
      End If
    End If
  End If
  
TidyUpAndExit:
  DeleteColumn = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function DeleteColumn_Transaction() As Boolean
  ' Transaction wrapper for the 'DeleteColumn' function.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  fOK = DeleteColumn
  
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  DeleteColumn_Transaction = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function ColumnIsUsed(pfrmUsage As frmUsage) As Boolean
  ' Return true if the column is used somewhere and
  ' therefore cannot be deleted.
  '
  ' Columns may be used in the following contexts :
  '   Lookup columns for other columns.
  '   Field components in expressions.
  '   Lookup field component in expressions.
  '   Lookup columns for prompted value components in expressions.
  '   Screen controls.
  '   Order/Find Columns.
  '   Summary Field definitions.
  '   Module definitions.
  '   Global Update/Add definitions.
  '   Chart Definitions.
  '   Document Types

  On Error GoTo ErrorTrap
  
  Set mfrmUse = pfrmUsage
  
  Dim fUsed As Boolean
  Dim sSQL As String
  Dim rsOtherColumns As DAO.Recordset
  Dim rsExpressions As DAO.Recordset
  Dim rsScreens As DAO.Recordset
  Dim rsOrders As DAO.Recordset
  Dim rsViews As DAO.Recordset
  Dim rsEmail As DAO.Recordset
  Dim rsOutlook As DAO.Recordset
  Dim rsSummaryFields As DAO.Recordset
  Dim rsModules As DAO.Recordset
  Dim rsAccord As DAO.Recordset
'NHRD Prototype Fusion Code***********************************************************
  'Dim rsFusion As DAO.Recordset ' This can be deleted when it is confirmed its not needed
  Dim rsValidations As DAO.Recordset
  Dim sColumnName As String
  Dim sTableName As String
  Dim sName As String
  Dim sModuleName As String
  Dim sGlobalFunctionType As String
  Dim sAccess As String
  Dim strSYSOrderIDs As String
  Dim strSYSCheckedOrderIDs As String
  Dim rsDAOTemp As DAO.Recordset

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
  
  Dim rsUtilsDocTypes As ADODB.Recordset
  
  Dim rsExprs As ADODB.Recordset
  
  Dim sMessage As String
  
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
  Dim bDM As Boolean
  Dim bMo As Boolean
  
  Dim strExpressions() As String
  Dim intIndex As Integer
  Dim blnFound As Boolean
  Dim objComp As CExprComponent
  Dim lngExprID As Long
  Dim fExprUsed As Boolean
  
  strSYSOrderIDs = vbNullString
  'strSYSEmailIDs = vbNullString
  
  fUsed = False
  
  ' Get the column's name.
  If ReadColumn Then
    sName = Me.Properties("ColumnName").value
  Else
    sName = "<unknown>"
  End If
  
  
  'NOTE: Item zero is delibrately left blank to
  'avoid blank items being added to the grid
  ReDim strExpressions(0)

  
  'Check Access db
  sSQL = "SELECT ExprID " & _
         "FROM   tmpComponents " & _
         "WHERE  FieldColumnID = " & CStr(gLngColumnID) & _
         "   OR  LookupColumnID = " & CStr(gLngColumnID)
  Set rsExpressions = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  
  If Not (rsExpressions.BOF And rsExpressions.EOF) Then
    'fUsed = True   'MH20061110 Fault 11670
    Do Until rsExpressions.EOF
      
      sMessage = (GetExpressionUsageDesc(rsExpressions!ExprID))
      
      blnFound = False
      For intIndex = 0 To UBound(strExpressions)
        If strExpressions(intIndex) = sMessage Then
          blnFound = True
          Exit For
        End If
      Next

      If Not blnFound Then
        fUsed = True
        mfrmUse.AddToList sMessage

        intIndex = UBound(strExpressions) + 1
        ReDim strExpressions(intIndex)
        strExpressions(intIndex) = sMessage
      End If
      
      rsExpressions.MoveNext
    Loop
  End If
  rsExpressions.Close


  'Check SQL db
  'NOTE: Need to check email address definitions in both SYS and DAT
  sSQL = "SELECT componentID " & _
         "FROM   ASRSysExprComponents " & _
         "WHERE  FieldColumnID = " & CStr(gLngColumnID) & _
         "    OR LookupColumnID = " & CStr(gLngColumnID)
  Set rsExprs = New ADODB.Recordset
  ' AE20080326 Fault #13043 - KB272358
  'rsExprs.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  rsExprs.Open sSQL, gADOCon, adOpenStatic, adLockReadOnly
  If Not (rsExprs.BOF And rsExprs.EOF) Then
    'fUsed = True     'MH20061110 Fault 11670
    Do Until rsExprs.EOF
      'JPD 20080425 Fault 12884
      Set objComp = New CExprComponent
      objComp.ComponentID = rsExprs!ComponentID
      lngExprID = objComp.RootExpressionID
      Set objComp = Nothing
      
      sSQL = "SELECT changed, deleted " & _
             "FROM   tmpExpressions " & _
             "WHERE  exprID = " & CStr(lngExprID)
      Set rsExpressions = daoDb.OpenRecordset(sSQL, _
        dbOpenForwardOnly, dbReadOnly)
        
      fExprUsed = (rsExpressions.BOF And rsExpressions.EOF)
      If Not fExprUsed Then
        fExprUsed = (Not rsExpressions!Changed) _
          And (Not rsExpressions!Deleted)
      End If
      
      If fExprUsed Then
        sMessage = (GetExpressionUsageDescFromSQL(lngExprID))
        
        blnFound = False
        For intIndex = 0 To UBound(strExpressions)
          If strExpressions(intIndex) = sMessage Then
            blnFound = True
            Exit For
          End If
        Next
  
        If Not blnFound Then
          fUsed = True
          mfrmUse.AddToList sMessage
  
          intIndex = UBound(strExpressions) + 1
          ReDim strExpressions(intIndex)
          strExpressions(intIndex) = sMessage
        End If
      End If
      rsExpressions.Close
      
      rsExprs.MoveNext
    Loop
  End If
  rsExprs.Close
  Set rsExprs = Nothing

  ' Find any columns that use this column for lookup values.
  sSQL = "SELECT DISTINCT tmpColumns.columnName, tmpTables.tableName" & _
    " FROM tmpColumns, tmpTables" & _
    " WHERE tmpTables.deleted = FALSE" & _
    "   AND tmpColumns.deleted = FALSE" & _
    "   AND tmpColumns.tableID = tmpTables.tableID" & _
    "   AND tmpColumns.lookupColumnID = " & Trim(Str(gLngColumnID))
  Set rsOtherColumns = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  If Not (rsOtherColumns.BOF And rsOtherColumns.EOF) Then
    fUsed = True
    Do Until rsOtherColumns.EOF
      ' Get the column and table names.
      sColumnName = rsOtherColumns.Fields("columnName")
      sTableName = rsOtherColumns.Fields("tableName")
      mfrmUse.AddToList ("Lookup Column : " & sColumnName & " <" & sTableName & ">")
      rsOtherColumns.MoveNext
    Loop
  End If
  'Close temporary recordset
  rsOtherColumns.Close
  
  ' Find any columns that use this column for lookup filter columns.
  sSQL = "SELECT DISTINCT tmpColumns.columnName, tmpTables.tableName" & _
    " FROM tmpColumns, tmpTables" & _
    " WHERE tmpTables.deleted = FALSE" & _
    "   AND tmpColumns.deleted = FALSE" & _
    "   AND tmpColumns.tableID = tmpTables.tableID" & _
    "   AND tmpColumns.lookupFilterColumnID = " & Trim(Str(gLngColumnID))
  Set rsOtherColumns = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  If Not (rsOtherColumns.BOF And rsOtherColumns.EOF) Then
    fUsed = True
    Do Until rsOtherColumns.EOF
      ' Get the column and table names.
      sColumnName = rsOtherColumns.Fields("columnName")
      sTableName = rsOtherColumns.Fields("tableName")
      mfrmUse.AddToList ("Lookup Filter Column : " & sColumnName & " <" & sTableName & ">")
      rsOtherColumns.MoveNext
    Loop
  End If
  'Close temporary recordset
  rsOtherColumns.Close
  
  ' Find any columns that use this column for lookup filter values.
  sSQL = "SELECT DISTINCT tmpColumns.columnName, tmpTables.tableName" & _
    " FROM tmpColumns, tmpTables" & _
    " WHERE tmpTables.deleted = FALSE" & _
    "   AND tmpColumns.deleted = FALSE" & _
    "   AND tmpColumns.tableID = tmpTables.tableID" & _
    "   AND tmpColumns.lookupFilterValueID = " & Trim(Str(gLngColumnID))
  Set rsOtherColumns = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  If Not (rsOtherColumns.BOF And rsOtherColumns.EOF) Then
    fUsed = True
    Do Until rsOtherColumns.EOF
      ' Get the column and table names.
      sColumnName = rsOtherColumns.Fields("columnName")
      sTableName = rsOtherColumns.Fields("tableName")
      mfrmUse.AddToList ("Lookup Filter Value : " & sColumnName & " <" & sTableName & ">")
      rsOtherColumns.MoveNext
    Loop
  End If
  'Close temporary recordset
  rsOtherColumns.Close
  
  
''  ' Check that the column is not used for a lookup prompted value in an expression.
''  sSQL = "SELECT DISTINCT tmpExpressions.exprID " & _
''    " FROM tmpExpressions " & _
''    " WHERE (tmpExpressions.type NOT IN (" & giEXPR_RUNTIMECALCULATION & "," & giEXPR_RUNTIMEFILTER & "," & giEXPR_RECORDINDEPENDANTCALC & "," & _
''                                             giEXPR_UTILRUNTIMEFILTER & "," & giEXPR_MATCHJOINEXPRESSION & "," & giEXPR_MATCHSCOREEXPRESSION & "," & _
''                                             giEXPR_MATCHWHEREEXPRESSION & ")) " & _
''    "   AND (tmpExpressions.parentComponentID = 0)"
'''    "   AND ((tmpComponents.fieldColumnID = " & Trim(Str(gLngColumnID)) & ") " & _
'''    "   OR  (tmpComponents.lookupColumnID = " & Trim(Str(gLngColumnID)) & ")) " & _
''
''  Set rsExpressions = daoDb.OpenRecordset(sSQL, _
''    dbOpenForwardOnly, dbReadOnly)
''  If Not (rsExpressions.BOF And rsExpressions.EOF) Then
'''    fUsed = True
''    Do Until rsExpressions.EOF
''      sExprName = vbNullString
''      sExprParentTable = vbNullString
''
''      ' Get the expression name and type description.
''      Set objExpr = New CExpression
''      lngExprID = rsExpressions!ExprID
''      objExpr.ExpressionID = lngExprID
''
''      If objExpr.ConstructExpression Then
''        sExprName = objExpr.Name
''        sExprType = objExpr.ExpressionTypeName
''        lngExprParentTableID = objExpr.BaseTableID
''        If objExpr.ExpressionContainsColumn(gLngColumnID) Then
''          fUsed = True
''          If objExpr.ExpressionType = giEXPR_RECORDINDEPENDANTCALC Then
''            sExprParentTable = vbNullString
''          Else
''            ' Get the expression's parent table name.
''            recTabEdit.Index = "idxTableID"
''            recTabEdit.Seek "=", lngExprParentTableID
''            If Not recTabEdit.NoMatch Then
''              sExprParentTable = recTabEdit!TableName
''            Else
''              sExprParentTable = "<unknown>"
''            End If
''          End If
''          mfrmUse.AddToList (sExprType & " : " & sExprName & " <" & sExprParentTable & ">")
''        End If
''      End If
''      ' Disassociate object variables.
''      'objExpr.Terminate      'MH20060727
''      Set objExpr = Nothing
''
''      rsExpressions.MoveNext
''    Loop
''  End If
''  'Close temporary recordset
''  rsExpressions.Close


  ' Check that the column is not used in any screens.
  sSQL = "SELECT DISTINCT tmpScreens.name, tmpTables.tableName" & _
    " FROM tmpScreens, tmpControls, tmpTables" & _
    " WHERE tmpScreens.deleted = FALSE" & _
    " AND tmpTables.deleted = FALSE" & _
    " AND tmpScreens.screenID = tmpControls.screenID" & _
    " AND tmpTables.tableID = tmpControls.tableID" & _
    " AND tmpControls.columnID = " & Trim(Str(gLngColumnID))
  Set rsScreens = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  If Not (rsScreens.BOF And rsScreens.EOF) Then
    fUsed = True
    Do Until rsScreens.EOF
      mfrmUse.AddToList ("Screen : " & rsScreens!Name & " <" & rsScreens!TableName & ">")
      rsScreens.MoveNext
    Loop
  End If
  'Close temporary recordset
  rsScreens.Close

  
  ' Check that the column is not used in any orders/find column collections.
  sSQL = "SELECT DISTINCT tmpOrders.orderID, tmpOrders.name, tmpTables.tableName" & _
    " FROM tmpOrders, tmpOrderItems, tmpTables" & _
    " WHERE tmpOrders.deleted = FALSE" & _
    " AND tmpTables.deleted = FALSE" & _
    " AND tmpOrders.orderID = tmpOrderItems.orderID" & _
    " AND tmpTables.tableID = tmpOrders.tableID" & _
    " AND tmpOrderItems.columnID = " & Trim(Str(gLngColumnID))
  Set rsOrders = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  If Not (rsOrders.BOF And rsOrders.EOF) Then
    fUsed = True
    Do Until rsOrders.EOF
      strSYSOrderIDs = strSYSOrderIDs & IIf(Len(strSYSOrderIDs) > 0, ",", "") & rsOrders!OrderID
      mfrmUse.AddToList ("Order : " & rsOrders!Name & " <" & rsOrders!TableName & ">")
      rsOrders.MoveNext
    Loop
  End If
  'Close temporary recordset
  rsOrders.Close
    
    
  ' Check that the column is not used in any views.
  sSQL = "SELECT DISTINCT tmpViews.viewName, tmpTables.tableName" & _
    " FROM tmpViewColumns, tmpViews, tmpTables" & _
    " WHERE tmpViewColumns.deleted = FALSE" & _
    " AND tmpViews.deleted = FALSE" & _
    " AND tmpTables.deleted = FALSE" & _
    " AND tmpViewColumns.columnID = " & Trim(Str(gLngColumnID)) & _
    " AND tmpViewColumns.viewID = tmpViews.viewID" & _
    " AND tmpViewColumns.inView <> FALSE" & _
    " AND tmpViews.viewTableID = tmptables.tableID"
  Set rsViews = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  If Not (rsViews.BOF And rsViews.EOF) Then
    fUsed = True
    Do Until rsViews.EOF
      mfrmUse.AddToList ("View : " & rsViews!ViewName & " <" & rsViews!TableName & ">")
      rsViews.MoveNext
    Loop
  End If
  'Close temporary recordset
  rsViews.Close
    
    
  ' Table validations
  sSQL = "SELECT Message FROM tmpTableValidations" & _
    " WHERE tmpTableValidations.ColumnID = " & Trim(Str(gLngColumnID)) & _
    " OR tmpTableValidations.EventStartDateColumnID =" & Trim(Str(gLngColumnID)) & _
    " OR tmpTableValidations.EventStartSessionColumnID =" & Trim(Str(gLngColumnID)) & _
    " OR tmpTableValidations.EventEndDateColumnID =" & Trim(Str(gLngColumnID)) & _
    " OR tmpTableValidations.EventEndSessionColumnID =" & Trim(Str(gLngColumnID)) & _
    " OR tmpTableValidations.EventTypeColumnID =" & Trim(Str(gLngColumnID))
  Set rsValidations = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  If Not (rsValidations.BOF And rsValidations.EOF) Then
    fUsed = True
    Do Until rsValidations.EOF
      mfrmUse.AddToList ("Validation : " & rsValidations!Message)
      rsValidations.MoveNext
    Loop
  End If
  'Close temporary recordset
  rsValidations.Close
    
    
  ' Check that the column is not used in any Summary Field definitions.
  sSQL = "SELECT DISTINCT tmpTables.tableName" & _
    " FROM tmpSummary, tmpTables" & _
    " WHERE tmpTables.deleted = FALSE" & _
    " AND tmpSummary.parentColumnID = " & Trim(Str(gLngColumnID)) & _
    " AND tmpSummary.historyTableID = tmpTables.tableID"
  Set rsSummaryFields = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  If Not (rsSummaryFields.BOF And rsSummaryFields.EOF) Then
    fUsed = True
    Do Until rsSummaryFields.EOF
      mfrmUse.AddToList ("Summary Field : " & rsSummaryFields.Fields("tableName"))
      rsSummaryFields.MoveNext
    Loop
  End If
  ' Close the recordset.
  rsSummaryFields.Close
  
  
  ' Check that the column is not used in any Module definitions.
  sSQL = "SELECT DISTINCT moduleKey, parameterKey" & _
    " FROM tmpModuleSetup" & _
    " WHERE parameterType = '" & gsPARAMETERTYPE_COLUMNID & "'" & _
    " AND parameterValue = '" & Trim(Str(gLngColumnID)) & "'" & _
    " UNION" & _
    " SELECT moduleKey, parameterKey" & _
    " FROM tmpModuleRelatedColumns" & _
    " WHERE sourceColumnID = " & Trim(Str(gLngColumnID)) & _
    " OR destColumnID = " & Trim(Str(gLngColumnID))
    
  Set rsModules = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  If Not (rsModules.BOF And rsModules.EOF) Then
    Do Until rsModules.EOF
      If rsModules!moduleKey = gsMODULEKEY_TRAININGBOOKING And (Not bTR) Then
        fUsed = True
        bTR = True
        sModuleName = "Training Booking"
        mfrmUse.AddToList (sModuleName & " Configuration")
      
      ElseIf ((rsModules!moduleKey = gsMODULEKEY_PERSONNEL) Or (rsModules!moduleKey = gsMODULEKEY_HIERARCHY)) And (Not bPE) Then
        fUsed = True
        bPE = True
        sModuleName = "Personnel"
        mfrmUse.AddToList (sModuleName & " Configuration")
      
      ElseIf (rsModules!moduleKey = gsMODULEKEY_WORKFLOW) And (Not bWF) Then
        If Application.WorkflowModule Then
          fUsed = True
          bWF = True
          sModuleName = "Workflow"
          mfrmUse.AddToList (sModuleName & " Configuration")
        End If
        
      ElseIf rsModules!moduleKey = gsMODULEKEY_ABSENCE Then
        fUsed = True
        
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
        fUsed = True
        bCU = True
        sModuleName = "Currency"
        mfrmUse.AddToList (sModuleName & " Configuration")
      
      ElseIf rsModules!moduleKey = gsMODULEKEY_POST And (Not bPO) Then
        fUsed = True
        bPO = True
        sModuleName = "Post"
        mfrmUse.AddToList (sModuleName & " Configuration")
      
      ElseIf rsModules!moduleKey = gsMODULEKEY_MATERNITY And (Not bMA) Then
        fUsed = True
        bMA = True
        sModuleName = "Maternity"
        mfrmUse.AddToList (sModuleName & " Configuration")
      
      ElseIf rsModules!moduleKey = gsMODULEKEY_AUDIT And (Not bAu) Then
        fUsed = True
        bAu = True
        sModuleName = "Audit"
        mfrmUse.AddToList (sModuleName & " Configuration")
      
      ElseIf rsModules!moduleKey = gsMODULEKEY_SSINTRANET And (Not bSS) Then
        fUsed = True
        bSS = True
        sModuleName = "Self Service Intranet"
        mfrmUse.AddToList (sModuleName & " Configuration")
      
      ElseIf rsModules!moduleKey = gsMODULEKEY_DOCMANAGEMENT And (Not bDM) Then
        fUsed = True
        bDM = True
        sModuleName = "Document Management"
        mfrmUse.AddToList (sModuleName & " Configuration")
      
      ElseIf rsModules!moduleKey = gsMODULEKEY_MOBILE And (Not bMo) Then
        fUsed = True
        bMo = True
        sModuleName = "Mobile Workflow"
        mfrmUse.AddToList (sModuleName & " Configuration")
      End If
      
      rsModules.MoveNext
    Loop
  End If
  ' Close the recordset.
  rsModules.Close

  
  ' JDM - 10/11/2004 - Fault 9415 - Payroll module setup
  'AE20071204 Fault #12658
  If IsModuleEnabled(modAccord) Then
    sSQL = "SELECT rtrim(t.TransferType) + ' - ' + f.Description AS Description" & _
           " FROM tmpAccordTransferFieldDefinitions f " & _
           " INNER JOIN tmpAccordTransferTypes t ON t.TransferTypeID = f.TransferTypeID" & _
           " WHERE ASRColumnID = " & Me.ColumnID & " AND ASRMapType = 0"
    Set rsAccord = daoDb.OpenRecordset(sSQL)
    With rsAccord
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          mfrmUse.AddToList "Payroll Configuration : " & Trim(!Description)
          .MoveNext
        Loop
      End If
      .Close
    End With
  End If

'NHRD Prototype Fusion Code***********************************************************
' This can be deleted when it is confirmed its not needed
'  If gbFusionModule Then
'    sSQL = "SELECT rtrim(t.FusionType) + ' - ' + f.Description AS Description" & _
'           " FROM tmpFusionFieldDefinitions f " & _
'           " INNER JOIN tmpFusionTypes t ON t.FusionTypeID = f.FusionTypeID" & _
'           " WHERE ASRColumnID = " & Me.ColumnID & " AND ASRMapType = 0"
'    Set rsFusion = daoDb.OpenRecordset(sSQL)
'    With rsFusion
'      If Not (.EOF And .BOF) Then
'        fUsed = True
'        Do Until .EOF
'          mfrmUse.AddToList "Fusion Configuration : " & Trim(!Description)
'          .MoveNext
'        Loop
'      End If
'      .Close
'    End With
'  End If
'NHRD Prototype Fusion Code***********************************************************
  
  'TM20011003 Fault 2191 - 1
  'Make check for the columns use in email definitions.
  sSQL = "SELECT DISTINCT tmpEmailAddresses.name " & _
         "FROM tmpEmailAddresses " & _
         "WHERE tmpEmailAddresses.ColumnID = " & Me.ColumnID & _
         " AND tmpEmailAddresses.Deleted = FALSE "
  Set rsEmail = daoDb.OpenRecordset(sSQL)
  With rsEmail
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Email definition : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsEmail = Nothing


  'MH20095027 Email Link content
  sSQL = "SELECT DISTINCT tmpTables.TableName, tmpEmailLinks.Title " & _
         "FROM tmpLinkContent, tmpEmailLinksColumns, tmpEmailLinks, tmpTables " & _
         "WHERE ( ( tmpLinkContent.FieldID = " & CStr(Me.ColumnID) & _
         "          AND tmpLinkContent.FieldCode = 'C' " & _
         "          AND ( tmpLinkContent.ContentID = tmpEmailLinks.SubjectContentID " & _
         "                OR tmpLinkContent.ContentID = tmpEmailLinks.BodyContentID " & _
         "              )" & _
         "        ) " & _
         "        OR " & _
         "        ( tmpEmailLinksColumns.ColumnID = " & CStr(Me.ColumnID) & _
         "          AND tmpEmailLinksColumns.LinkID = tmpEmailLinks.LinkID" & _
         "        )" & _
         "      ) " & _
         "  AND tmpEmailLinks.TableID = tmpTables.TableID " & _
         "  AND tmpEmailLinks.Deleted = False "
  Set rsEmail = daoDb.OpenRecordset(sSQL)
  With rsEmail
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList "Email Link : " & !Title & " <" & !TableName & ">"
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsEmail = Nothing



  'MH20040324 Outlook Calendar Link date/time range
  sSQL = "SELECT DISTINCT tmpOutlookLinks.Title, tmpTables.TableName " & _
         "FROM tmpOutlookLinks, tmpTables " & _
         "WHERE tmpOutlookLinks.TableID = tmpTables.TableID " & _
         " AND tmpOutlookLinks.Deleted = False " & _
         " AND (StartDate = " & CStr(Me.ColumnID) & " OR EndDate = " & CStr(Me.ColumnID) & " OR ColumnStartTime = " & CStr(Me.ColumnID) & " OR ColumnEndTime = " & CStr(Me.ColumnID) & ")"
  Set rsOutlook = daoDb.OpenRecordset(sSQL)
  With rsOutlook
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList "Outlook Calendar Link : " & !Title & " <" & !TableName & ">"
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsOutlook = Nothing



  'MH20040324 Outlook Calendar Link on the columns tab
  sSQL = "SELECT DISTINCT tmpOutlookLinks.Title, tmpTables.TableName " & _
         "FROM tmpOutlookLinksColumns, tmpOutlookLinks, tmpTables " & _
         "WHERE tmpOutlookLinksColumns.ColumnID = " & Me.ColumnID & _
         " AND tmpOutlookLinks.LinkID = tmpOutlookLinksColumns.LinkID " & _
         " AND tmpOutlookLinks.TableID = tmpTables.TableID " & _
         " AND tmpOutlookLinks.Deleted = False " & _
         " AND (StartDate <> " & CStr(Me.ColumnID) & " AND EndDate <> " & CStr(Me.ColumnID) & " AND ColumnStartTime <> " & CStr(Me.ColumnID) & " AND ColumnEndTime <> " & CStr(Me.ColumnID) & ")"
  Set rsOutlook = daoDb.OpenRecordset(sSQL)
  With rsOutlook
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList "Outlook Calendar Link : " & !Title & " <" & !TableName & ">"
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsOutlook = Nothing

  ' ### We now need to check every single util/report/tool for column usage
  
  ' NHRD09112010 JIRA HRPRO-1083
  ' Check if the column is used in any Document Types
  sSQL = "SELECT DocumentMapID, Name" & _
            ",ASRSysDocumentManagementTypes.TargetKeyFieldColumnID" & _
            ",ASRSysDocumentManagementTypes.TargetColumnID" & _
            ",ASRSysDocumentManagementTypes.TargetCategoryColumnID" & _
            ",ASRSysDocumentManagementTypes.TargetTypeColumnID" & _
            ",ASRSysDocumentManagementTypes.Parent1KeyFieldColumnID " & _
         "FROM ASRSysDocumentManagementTypes " & _
         "WHERE ASRSysDocumentManagementTypes.TargetKeyFieldColumnID = " & gLngColumnID & _
            " OR ASRSysDocumentManagementTypes.TargetColumnID = " & gLngColumnID & _
            " OR ASRSysDocumentManagementTypes.TargetCategoryColumnID = " & gLngColumnID & _
            " OR ASRSysDocumentManagementTypes.TargetTypeColumnID = " & gLngColumnID & _
            " OR ASRSysDocumentManagementTypes.Parent1KeyFieldColumnID = " & gLngColumnID
        
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
         " INNER JOIN ASRSysCustomReportsDetails" & _
         "   ON ASRSysCustomReportsName.ID = ASRSysCustomReportsDetails.CustomReportID" & _
         " WHERE ASRSysCustomReportsDetails.ColExprID = " & Trim(Str(gLngColumnID)) & _
         "   AND ASRSysCustomReportsDetails.Type = 'C'"
  Set rsUtils1 = New ADODB.Recordset
  ' AE20080326 Fault #13043 - KB272358
  'rsUtils1.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  rsUtils1.Open sSQL, gADOCon, adOpenStatic, adLockReadOnly
  With rsUtils1
    If Not (.EOF And .BOF) Then
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
    "   LEFT OUTER JOIN ASRSysCalendarReportEvents " & _
    "   ON ASRSysCalendarReports.ID = ASRSysCalendarReportEvents.calendarReportID " & _
    "   LEFT OUTER JOIN ASRSysCalendarReportOrder " & _
    "   ON ASRSysCalendarReports.ID = ASRSysCalendarReportOrder.calendarReportID " & _
    " WHERE ASRSysCalendarReports.description1 = " & Trim(Str(gLngColumnID)) & _
    "   OR ASRSysCalendarReports.description2 = " & Trim(Str(gLngColumnID)) & _
    "   OR ASRSysCalendarReports.region = " & Trim(Str(gLngColumnID)) & _
    "   OR ASRSysCalendarReportEvents.eventStartDateID = " & Trim(Str(gLngColumnID)) & _
    "   OR ASRSysCalendarReportEvents.eventStartSessionID = " & Trim(Str(gLngColumnID)) & _
    "   OR ASRSysCalendarReportEvents.eventEndDateID = " & Trim(Str(gLngColumnID)) & _
    "   OR ASRSysCalendarReportEvents.eventEndSessionID = " & Trim(Str(gLngColumnID)) & _
    "   OR ASRSysCalendarReportEvents.eventDurationID = " & Trim(Str(gLngColumnID)) & _
    "   OR ASRSysCalendarReportEvents.legendLookupColumnID = " & Trim(Str(gLngColumnID)) & _
    "   OR ASRSysCalendarReportEvents.legendLookupCodeID = " & Trim(Str(gLngColumnID)) & _
    "   OR ASRSysCalendarReportEvents.legendEventColumnID = " & Trim(Str(gLngColumnID)) & _
    "   OR ASRSysCalendarReportEvents.eventDesc1ColumnID = " & Trim(Str(gLngColumnID)) & _
    "   OR ASRSysCalendarReportEvents.eventDesc2ColumnID = " & Trim(Str(gLngColumnID)) & _
    "   OR ASRSysCalendarReportOrder.columnID = " & Trim(Str(gLngColumnID))
  Set rsUtils2 = New ADODB.Recordset
  ' AE20080326 Fault #13043 - KB272358
  'rsUtils2.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  rsUtils2.Open sSQL, gADOCon, adOpenStatic, adLockReadOnly
  With rsUtils2
    If Not (.EOF And .BOF) Then
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
         " WHERE crossTabType <> " & ctt9GridBox & " AND " & _
         " (ASRSysCrossTab.HorizontalColID = " & Trim(Str(gLngColumnID)) & _
         "   OR ASRSysCrossTab.VerticalColID = " & Trim(Str(gLngColumnID)) & _
         "   OR ASRSysCrossTab.PageBreakColID = " & Trim(Str(gLngColumnID)) & _
         "   OR ASRSysCrossTab.IntersectionColID = " & Trim(Str(gLngColumnID)) & ")"
  Set rsUtils3 = New ADODB.Recordset
  ' AE20080326 Fault #13043 - KB272358
  'rsUtils3.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  rsUtils3.Open sSQL, gADOCon, adOpenStatic, adLockReadOnly
  With rsUtils3
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Cross Tab : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsUtils3 = Nothing

  ' Now Check if the column is used in a 9-Box Grid Report
  sSQL = "SELECT DISTINCT ASRSysCrossTab.CrossTabID, ASRSysCrossTab.Name, ASRSysCrossTab.Username" & _
         " FROM ASRSysCrossTab " & _
         " WHERE crossTabType = " & ctt9GridBox & " AND " & _
         " (ASRSysCrossTab.HorizontalColID = " & Trim(Str(gLngColumnID)) & _
         "   OR ASRSysCrossTab.VerticalColID = " & Trim(Str(gLngColumnID)) & _
         "   OR ASRSysCrossTab.PageBreakColID = " & Trim(Str(gLngColumnID)) & _
         "   OR ASRSysCrossTab.IntersectionColID = " & Trim(Str(gLngColumnID)) & ")"

  Set rsUtils12 = New ADODB.Recordset
  rsUtils12.Open sSQL, gADOCon, adOpenStatic, adLockReadOnly
  With rsUtils12
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("9-Box Grid Report : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsUtils12 = Nothing
  

  ' Now Check if the column is used in a Record Profile
  sSQL = "SELECT DISTINCT ASRSysRecordProfileName.Name, ASRSysRecordProfileName.Username, ASRSysRecordProfileName.recordProfileID " & _
    "FROM ASRSysRecordProfileName " & _
    "   INNER JOIN ASRSysRecordProfileDetails " & _
    "   ON ASRSysRecordProfileName.recordProfileID = ASRSysRecordProfileDetails.recordProfileID " & _
    "WHERE ASRSysRecordProfileDetails.columnID = " & Trim(Str(gLngColumnID)) & " " & _
    "AND ASRSysRecordProfileDetails.Type = 'C'"
  Set rsUtils4 = New ADODB.Recordset
  ' AE20080326 Fault #13043 - KB272358
  'rsUtils4.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  rsUtils4.Open sSQL, gADOCon, adOpenStatic, adLockReadOnly
  With rsUtils4
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Record Profile : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsUtils4 = Nothing
  

  ' Now Check if the column is used in a Match Report/Career Progression/Succession Planning
  sSQL = "SELECT DISTINCT ASRSysMatchReportName.Name, ASRSysMatchReportName.Username, ASRSysMatchReportName.matchReportID, ASRSysMatchReportName.matchReportType" & _
    " FROM ASRSysMatchReportName" & _
    "     LEFT OUTER JOIN ASRSysMatchReportDetails " & _
    "     ON ASRSysMatchReportName.matchReportID = ASRSysMatchReportDetails.matchReportID" & _
    "     LEFT OUTER JOIN ASRSysMatchReportBreakdown " & _
    "     ON ASRSysMatchReportName.matchReportID = ASRSysMatchReportBreakdown.matchReportID" & _
    " WHERE (ASRSysMatchReportDetails.colExprID = " & Trim(Str(gLngColumnID)) & _
    "   AND ASRSysMatchReportDetails.colType = 'C') " & _
    " OR (ASRSysMatchReportBreakdown.colExprID = " & Trim(Str(gLngColumnID)) & _
    "   AND ASRSysMatchReportBreakdown.colType = 'C')"
  Set rsUtils5 = New ADODB.Recordset
  ' AE20080326 Fault #13043 - KB272358
  'rsUtils5.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  rsUtils5.Open sSQL, gADOCon, adOpenStatic, adLockReadOnly
  With rsUtils5
    If Not (.EOF And .BOF) Then
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
         "FROM ASRSysMailMergeName " & _
         "    INNER JOIN ASRSysMailMergeColumns " & _
         "    ON ASRSysMailMergeName.MailMergeID = ASRSysMailMergeColumns.MailMergeID " & _
         "WHERE ASRSysMailMergeColumns.ColumnID = " & Trim(Str(gLngColumnID)) & " " & _
         "  AND ASRSysMailMergeColumns.Type = 'C'"
  Set rsUtils6 = New ADODB.Recordset
  ' AE20080326 Fault #13043 - KB272358
  'rsUtils6.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  rsUtils6.Open sSQL, gADOCon, adOpenStatic, adLockReadOnly
  With rsUtils6
    If Not (.EOF And .BOF) Then
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


  ' Check that the column is not used in any Global Add/Update definitions.
  sSQL = "SELECT DISTINCT ASRSysGlobalFunctions.functionID, ASRSysGlobalFunctions.name, ASRSysGlobalFunctions.type, ASRSysGlobalFunctions.username" & _
        " FROM ASRSysGlobalFunctions" & _
        "     LEFT OUTER JOIN ASRSysGlobalItems " & _
        "     ON ASRSysGlobalFunctions.functionID = ASRSysGlobalItems.functionID" & _
        " WHERE (ASRSysGlobalItems.refColumnID = " & Trim(Str(gLngColumnID)) & ") " & _
        "   OR (ASRSysGlobalItems.LookupColumnID = " & Trim(Str(gLngColumnID)) & ") " & _
        "   OR (ASRSysGlobalItems.ColumnID = " & Trim(Str(gLngColumnID)) & ")"
  Set rsUtils7 = New ADODB.Recordset
  ' AE20080326 Fault #13043 - KB272358
  'rsUtils7.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  rsUtils7.Open sSQL, gADOCon, adOpenStatic, adLockReadOnly
  With rsUtils7
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
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
  Set rsUtils7 = Nothing

    
  ' Now Check if the column is used in a Data Transfer
  sSQL = "SELECT DISTINCT ASRSysDataTransferName.Name, ASRSysDataTransferName.DataTransferID, ASRSysDataTransferName.Username " & _
         "FROM ASRSysDataTransferName " & _
         "      INNER JOIN ASRSysDataTransferColumns " & _
         "      ON ASRSysDataTransferName.DataTransferID = ASRSysDataTransferColumns.DataTransferID " & _
         "WHERE ASRSysDataTransferColumns.FromColumnID = " & Trim(Str(gLngColumnID)) & " " & _
         "    OR ASRSysDataTransferColumns.ToColumnID = " & Trim(Str(gLngColumnID))
  Set rsUtils8 = New ADODB.Recordset
  ' AE20080326 Fault #13043 - KB272358
  'rsUtils8.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  rsUtils8.Open sSQL, gADOCon, adOpenStatic, adLockReadOnly
  With rsUtils8
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Data Transfer : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsUtils8 = Nothing
  
  
  ' Now Check if the column is used in a Export
  sSQL = "SELECT DISTINCT ASRSysExportName.Name, ASRSysExportName.ID, ASRSysExportName.Username " & _
         "FROM ASRSysExportName " & _
         "    INNER JOIN ASRSysExportDetails " & _
         "    ON ASRSysExportName.ID = ASRSysExportDetails.ExportID " & _
         "WHERE ASRSysExportDetails.ColExprID = " & Trim(Str(gLngColumnID)) & " " & _
         "  AND ASRSysExportDetails.Type = 'C'"
  Set rsUtils9 = New ADODB.Recordset
  ' AE20080326 Fault #13043 - KB272358
  'rsUtils9.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  rsUtils9.Open sSQL, gADOCon, adOpenStatic, adLockReadOnly
  With rsUtils9
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Export : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsUtils9 = Nothing
  
  
  ' Now Check if the column is used in an Import
  sSQL = "SELECT DISTINCT ASRSysImportName.Name, ASRSysImportName.ID, ASRSysImportName.Username " & _
         "FROM ASRSysImportName " & _
         "    INNER JOIN ASRSysImportDetails " & _
         "    ON ASRSysImportName.ID = ASRSysImportDetails.ImportID " & _
         "WHERE ASRSysImportDetails.ColExprID = " & Trim(Str(gLngColumnID)) & " " & _
         "  AND ASRSysImportDetails.Type = 'C'"
  Set rsUtils10 = New ADODB.Recordset
  ' AE20080326 Fault #13043 - KB272358
  'rsUtils10.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  rsUtils10.Open sSQL, gADOCon, adOpenStatic, adLockReadOnly
  With rsUtils10
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Import : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsUtils10 = Nothing


'  ' Now Check if the column is used in a Runtime Calculation OR Runtime Filter.
'  sSQL = "SELECT DISTINCT ASRSysExpressions.exprID, ASRSysExpressions.Type, ASRSysExpressions.Name, ASRSysTables.TableName " & _
'        " FROM ASRSysExpressions " & _
'        "       LEFT OUTER JOIN  ASRSysExprComponents " & _
'        "       ON ASRSysExpressions.ExprID = ASRSysExprComponents.ExprID " & _
'        "       LEFT OUTER JOIN ASRSysTables" & _
'        "       ON ASRSysExpressions.TableID = ASRSysTables.TableID" & _
'        " WHERE ASRSysExpressions.Type IN (" & giEXPR_RUNTIMECALCULATION & "," & giEXPR_RUNTIMEFILTER & ", " & giEXPR_RECORDINDEPENDANTCALC & ") " & _
'        "   AND ASRSysExpressions.ParentComponentID = 0"
''        "   AND (ASRSysExprComponents.FieldColumnID = " & Trim(Str(gLngColumnID)) & _
''        "      OR ASRSysExprComponents.LookupColumnID = " & Trim(Str(gLngColumnID)) & ")"
'  Set rsUtils = New ADODB.Recordset
'  rsUtils.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
'  With rsUtils
'    If Not (rsUtils.BOF And rsUtils.EOF) Then
''      fUsed = True
'      Do Until .EOF
'
'        If ExpressionUsesColumn(!ExprID, gLngColumnID) Then
'          fUsed = True
'
'          Select Case !Type
'            Case ExpressionTypes.giEXPR_RUNTIMEFILTER
'              sExprType = "Runtime Filter"
'              mfrmUse.AddToList (sExprType & " : " & !Name & " <" & !TableName & ">")
'            Case ExpressionTypes.giEXPR_RUNTIMECALCULATION
'              sExprType = "Runtime Calculation"
'              mfrmUse.AddToList (sExprType & " : " & !Name & " <" & !TableName & ">")
'            Case ExpressionTypes.giEXPR_RECORDINDEPENDANTCALC
'              sExprType = "Record Independent Calculation"
'              mfrmUse.AddToList (sExprType & " : " & !Name & " <Report Configuration/Calendar Reports>")
'            Case Else
'              sExprType = "Expression"
'          End Select
'
'        End If
'
'        .MoveNext
'      Loop
'    End If
'    .Close
'  End With
'  Set rsUtils = Nothing

  
  ' Check that the column is not used in any orders/find column collections (SQL TABLES).
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
        "   AND ASRSysColumns.ColumnID = " & Trim(Str(gLngColumnID)) & _
        "   AND ASRSysOrders.orderID NOT IN (" & strSYSCheckedOrderIDs & ")"
  
  'JPD 20040227 Fault 8163
  If Len(strSYSOrderIDs) > 0 Then
    sSQL = sSQL & _
        "   AND ASRSysOrders.OrderID NOT IN (" & strSYSOrderIDs & ")"
  End If
  
  Set rsUtils11 = New ADODB.Recordset
  ' AE20080326 Fault #13043 - KB272358
  'rsUtils11.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  rsUtils11.Open sSQL, gADOCon, adOpenStatic, adLockReadOnly
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

  If Application.WorkflowModule Then
    ' Check that it is not used by a Workflow StoredData element.
    sSQL = "SELECT DISTINCT tmpWorkflowElements.workflowID," & _
      "   tmpWorkflowElements.caption" & _
      " FROM tmpWorkflowElementColumns" & _
      " INNER JOIN tmpWorkflowElements ON tmpWorkflowElementColumns.elementID = tmpWorkflowElements.id" & _
      " WHERE tmpWorkflowElementColumns.columnID = " & Trim(Str(gLngColumnID))
  
    Set rsDAOTemp = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    If Not (rsDAOTemp.BOF And rsDAOTemp.EOF) Then
      Do Until rsDAOTemp.EOF
        recWorkflowEdit.Index = "idxWorkflowID"
        recWorkflowEdit.Seek "=", rsDAOTemp.Fields("workflowID")
  
        If Not recWorkflowEdit.NoMatch Then
          If recWorkflowEdit.Fields("deleted").value = False Then
            fUsed = True
            mfrmUse.AddToList ("Workflow : " & recWorkflowEdit.Fields("name").value & " <'" & rsDAOTemp.Fields("caption") & "' stored data element>")
          End If
        End If
  
        rsDAOTemp.MoveNext
      Loop
    End If
    'Close temporary recordset
    rsDAOTemp.Close
  
    ' Check that it is not used by a Workflow WebForm or Email element.
    sSQL = "SELECT DISTINCT tmpWorkflowElements.workflowID," & _
      "   tmpWorkflowElements.caption," & _
      "   tmpWorkflowElements.type " & _
      " FROM tmpWorkflowElementItems" & _
      " INNER JOIN tmpWorkflowElements ON tmpWorkflowElementItems.elementID = tmpWorkflowElements.id" & _
      " WHERE tmpWorkflowElementItems.dbColumnID = " & Trim(Str(gLngColumnID)) & _
      "   OR tmpWorkflowElementItems.LookupColumnID = " & Trim(Str(gLngColumnID)) & _
      "   OR tmpWorkflowElementItems.LookupFilterColumnID = " & Trim(Str(gLngColumnID)) & _
      " ORDER BY tmpWorkflowElements.type DESC"
    Set rsDAOTemp = daoDb.OpenRecordset(sSQL, _
      dbOpenForwardOnly, dbReadOnly)
    If Not (rsDAOTemp.BOF And rsDAOTemp.EOF) Then
      Do Until rsDAOTemp.EOF
        recWorkflowEdit.Index = "idxWorkflowID"
        recWorkflowEdit.Seek "=", rsDAOTemp.Fields("workflowID")
  
        If Not recWorkflowEdit.NoMatch Then
          If recWorkflowEdit.Fields("deleted").value = False Then
            fUsed = True
            mfrmUse.AddToList ("Workflow : " & recWorkflowEdit.Fields("name").value & " <'" & rsDAOTemp.Fields("caption") & "' " & IIf(rsDAOTemp.Fields("type") = 3, "email", "web form") & " element>")
          End If
        End If
  
        rsDAOTemp.MoveNext
      Loop
    End If
    'Close temporary recordset
    rsDAOTemp.Close
  
    ' Workflow Triggered Link date range
    sSQL = "SELECT DISTINCT [tmpWorkflows].[name], [tmpTables].[TableName] " & _
           "FROM [tmpWorkflowTriggeredLinks], [tmpWorkflows], [tmpTables] " & _
           "WHERE [tmpWorkflowTriggeredLinks].[TableID] = [tmpTables].[TableID] " & _
           " AND [tmpWorkflowTriggeredLinks].[workflowID] = [tmpWorkflows].[ID] " & _
           " AND [tmpWorkflowTriggeredLinks].[deleted] = FALSE " & _
           " AND [tmpWorkflowTriggeredLinks].[dateColumn] = " & Trim(Str(Me.ColumnID))
    Set rsDAOTemp = daoDb.OpenRecordset(sSQL, _
      dbOpenForwardOnly, dbReadOnly)
    With rsDAOTemp
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          mfrmUse.AddToList "Workflow Link : " & !Name & " <" & !TableName & ">"
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rsDAOTemp = Nothing
  
    ' Workflow Link on the columns tab
    sSQL = "SELECT DISTINCT tmpWorkflows.name, tmpTables.TableName " & _
           "FROM tmpWorkflowTriggeredLinkColumns, tmpWorkflowTriggeredLinks, tmpTables, tmpWorkflows " & _
           "WHERE tmpWorkflowTriggeredLinkColumns.ColumnID = " & Trim(Str(Me.ColumnID)) & _
           " AND tmpWorkflowTriggeredLinks.LinkID = tmpWorkflowTriggeredLinkColumns.LinkID " & _
           " AND tmpWorkflowTriggeredLinks.workflowID = tmpWorkflows.ID " & _
           " AND tmpWorkflowTriggeredLinks.TableID = tmpTables.TableID " & _
           " AND tmpWorkflowTriggeredLinks.Deleted = FALSE " & _
           " AND tmpWorkflowTriggeredLinks.dateColumn <> " & Trim(Str(Me.ColumnID))
    Set rsDAOTemp = daoDb.OpenRecordset(sSQL, _
      dbOpenForwardOnly, dbReadOnly)
    With rsDAOTemp
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          mfrmUse.AddToList "Workflow Link : " & !Name & " <" & !TableName & ">"
          .MoveNext
        Loop
      End If
      .Close
    End With
    Set rsDAOTemp = Nothing
  End If
  
  
  If Application.SelfServiceIntranetModule Then
    ' Check that it is not used in SSI Charting.
    sSQL = "SELECT DISTINCT tmpSSIntranetLinks.ID," & _
      "   tmpSSIntranetLinks.Element_Type," & _
      "   tmpSSIntranetLinks.text" & _
      " FROM tmpSSIntranetLinks" & _
      " WHERE tmpSSIntranetLinks.Chart_columnID = " & Trim(Str(gLngColumnID)) & _
      " OR tmpSSIntranetLinks.Chart_ColumnID_2 = " & Trim(Str(gLngColumnID)) & _
      " OR tmpSSIntranetLinks.Chart_ColumnID_3 = " & Trim(Str(gLngColumnID)) & _
      " OR tmpSSIntranetLinks.Chart_ColourID = " & Trim(Str(gLngColumnID)) & _
      " OR tmpSSIntranetLinks.Chart_FilterID = " & Trim(Str(gLngColumnID))
  
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
  Set rsOtherColumns = Nothing
  Set rsExpressions = Nothing
  Set rsScreens = Nothing
  Set rsOrders = Nothing
  Set rsViews = Nothing
  Set rsSummaryFields = Nothing
  Set rsModules = Nothing
  'Set objExpr = Nothing
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
  
  ColumnIsUsed = fUsed
  Exit Function

ErrorTrap:
  MsgBox "Error checking if the column is used." & vbCr & vbCr & _
    Err.Description, vbExclamation + vbOKOnly, App.ProductName
  fUsed = True
  Resume TidyUpAndExit
  
End Function






Public Function EditColumn(pbSaveAsNew As Boolean) As Boolean

  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim frmEdit As frmColEdit
  
' jpd - Already read.
'  fOK = ReadColumn
  
'  If fOK Then
    Set frmEdit = New frmColEdit
    Set frmEdit.Column = Me
    

    ' JDM - 27/03/01 - Clear the changed flag
    If pbSaveAsNew Then
        frmEdit.SetAsNew
    End If

    'MH20010110
    'Renaming a column currently produces an error so stop
    'users renaming the column for the moment
    'With frmEdit.txtColumnName
    '  .Locked = True
    '  .BackColor = vbButtonFace
    'End With

    Screen.MousePointer = vbDefault
    
    frmEdit.Show vbModal
    
    fOK = Not frmEdit.Cancelled And frmEdit.Changed

    If fOK Then
      fOK = WriteColumn_Transaction
    End If
'  End If
  
TidyUpAndExit:
  Set frmEdit = Nothing
  EditColumn = fOK
  Exit Function

ErrorTrap:
  If ASRDEVELOPMENT And Len(Err.Description) > 0 Then
    MsgBox Err.Description, vbCritical, "ASR DEVELOPMENT"
    Stop
  End If
  fOK = False
  Resume TidyUpAndExit

End Function


Public Function FindColumn() As Boolean
  On Error GoTo ErrorTrap
  
  With recColEdit
    .Index = "idxColumnID"
    .Seek "=", ColumnID
    FindColumn = (Not .NoMatch)
  End With
  
  Exit Function
  
ErrorTrap:
  FindColumn = False

End Function

Public Function InitColumnProps() As Boolean
  Dim ThisField As DAO.Field
  Dim fOK As Boolean
  
  Set Me.Properties = Nothing
  
  For Each ThisField In recColEdit.Fields
    Me.Properties.Add ThisField.Name, Null
  Next
  
  Me.Properties("tableID").value = gLngTableID
  Me.Properties("columnID").value = ColumnID
  Me.Properties("new").value = False
  Me.Properties("changed").value = False
  Me.Properties("deleted").value = False
  
  Me.Properties("columnType").value = giCOLUMNTYPE_DATA
  Me.Properties("dataType").value = dtVARCHAR
  Me.Properties("size").value = 1
  Me.Properties("controlType").value = giCTRL_TEXTBOX
  
  Me.Properties("defaultValue").value = ""
  
  Me.Properties("lookupTableID").value = 0
  Me.Properties("lookupColumnID").value = 0
  Me.Properties("lookupFilterValueID").value = 0
  Me.Properties("lookupFilterOperator").value = 0
  Me.Properties("lookupFilterColumnID").value = 0
  
  'Me.Properties("diaryPeriod") = iTimePeriodDays
  'Me.Properties("diaryOffset") = 0
  'Me.Properties("diaryRemind") = False
  'Me.Properties("diaryComment") = ""
  
  Me.Properties("spinnerMinimum").value = 0
  Me.Properties("spinnerMaximum").value = 10
  Me.Properties("spinnerIncrement").value = 1
  
  'TM20020204 Fault 3404 - set the default to 'No Conversion' not 'Proper Case'.
  'Me.Properties("convertCase") = 3
  Me.Properties("convertCase").value = 0
  Me.Properties("Use1000Separator").value = 0
  Me.Properties("Trimming").value = 1
  
  Me.Properties("mask").value = ""
  Me.Properties("alignment").value = vbLeftJustify
  
  Me.Properties("calcExprID").value = 0
  Me.Properties("gotFocusExprID").value = 0
  Me.Properties("lostFocusExprID").value = 0
  Me.Properties("locked").value = 0
  
  gsControlValues = vbNullString
  
    
  fOK = True
  If fOK Then fOK = ClearDiaryLinks
  'If fOK Then fOK = ClearEmailLinks

  InitColumnProps = fOK

End Function

Public Function NewColumn() As Boolean
  ' Create a new column.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim frmEdit As frmColEdit
  
  'Initialise column
  ColumnID = 0
  fOK = InitColumnProps
  IsNew = True
  
  'Display column edit form
  If fOK Then
    Set frmEdit = New frmColEdit
    Set frmEdit.Column = Me
    frmEdit.Show vbModal
    fOK = Not frmEdit.Cancelled
  End If
  
  If fOK Then
    fOK = WriteColumn_Transaction
  End If
  
TidyUpAndExit:
  Set frmEdit = Nothing
  
  ' Commit the data transaction if everything was okay.
  If Not fOK Then
    ColumnID = 0
    InitColumnProps
    IsNew = True
  End If
  
  NewColumn = fOK
  Exit Function
  
ErrorTrap:
  If ASRDEVELOPMENT And Len(Err.Description) > 0 Then
    MsgBox Err.Description, vbCritical, "ASR DEVELOPMENT"
    Stop
  End If
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function ReadColumn() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sIndex As String
  Dim objField As DAO.Field
  
  With recColEdit
    'Store current index and bookmark
    sIndex = .Index
    
    'Find the column record
    fOK = FindColumn
    
    If fOK Then
    
      'Set column object properties
      For Each objField In .Fields
        If Me.Properties(objField.Name) Is Nothing Then
          Me.Properties.Add objField.Name
        End If
        Me.Properties(objField.Name).value = objField.value
      Next
      Set objField = Nothing
      
      ' Read the control values.
      fOK = ReadControlValues
    
      If fOK Then
        fOK = ReadDiaryLinks
        'fOK = ReadEmailLinks
      End If
      
    End If
  End With
  
TidyUpAndExit:
  Set objField = Nothing
  'Restore original index and bookmark
  recColEdit.Index = sIndex
  ReadColumn = fOK
  Exit Function

ErrorTrap:
  If ASRDEVELOPMENT And Len(Trim(Err.Description)) > 0 Then
    MsgBox Err.Description, vbExclamation, "ASR DEVELOPMENT ONLY"
  End If
  
  fOK = False
  Resume TidyUpAndExit
  
End Function
Private Function ReadControlValues() As Boolean
  ' Read the column control values from the local tables.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sValues As String
  
  fOK = True
  sValues = vbNullString
  
  With recContValEdit
    .Index = "idxColumnID"
    .Seek ">=", ColumnID
    
    If Not .NoMatch Then
      ' Read the current control values into a string representation
      ' to be used as the list property for a combo, or the items
      ' in a radio group.
      Do While Not .EOF
    
        ' If no more control values for this column exit loop
        If !ColumnID <> ColumnID Then
          Exit Do
        End If
        
        If LenB(Trim(!value)) <> 0 Then
          sValues = sValues & Trim(!value) & vbNewLine
        End If
        
        .MoveNext
      Loop
      
      If LenB(sValues) <> 0 Then
        ' Remove the trailing carriage return and line feed characters.
        sValues = Left$(sValues, Len(sValues) - 2)
      End If
          
    End If
  
  End With
  
  gsControlValues = sValues
  
TidyUpAndExit:
  ReadControlValues = fOK
  Exit Function
  
ErrorTrap:
  gsControlValues = vbNullString
  fOK = False
  Resume TidyUpAndExit
  
End Function

Private Function ReadDiaryLinks() As Boolean
  ' Read the column's diary links from the database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objDiaryLink As cDiaryLink
  
  fOK = ClearDiaryLinks
  
  With recDiaryEdit
    .Index = "idxColumnID"
    .Seek "=", ColumnID
    
    If Not .NoMatch Then
      ' Read the current control values into a string representation
      ' to be used as the list property for a combo, or the items
      ' in a radio group.
      Do While Not .EOF
    
        ' If no more control values for this column exit loop
        If !ColumnID <> ColumnID Then
          Exit Do
        End If
        
        Set objDiaryLink = New cDiaryLink
        objDiaryLink.DiaryLinkId = !diaryID
        If objDiaryLink.ReadDiaryLink Then
          mobjDiaryLinks.Add objDiaryLink
        End If
         
        Set objDiaryLink = Nothing
        
        .MoveNext
      Loop
    End If
  End With
  
TidyUpAndExit:
  ReadDiaryLinks = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

'Private Function ReadEmailLinks() As Boolean
'  ' Read the column's Email links from the database.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim objEmailLink As clsEmailLink
'
'  fOK = ClearEmailLinks
'
'  With recEmailLinksEdit
'    .Index = "idxTableID"
'    .Seek "=", TableID
'
'    If Not .NoMatch Then
'      ' Read the current control values into a string representation
'      ' to be used as the list property for a combo, or the items
'      ' in a radio group.
'
'      Do While Not .EOF
'
'        ' If no more control values for this column exit loop
'        If !ColumnID = ColumnID Then
'
'          Set objEmailLink = New clsEmailLink
'          objEmailLink.LinkID = !LinkID
'          If objEmailLink.ReadEmailLink Then
'            mobjEmailLinks.Add objEmailLink, "ID" & objEmailLink.LinkID
'          End If
'          Set objEmailLink = Nothing
'
'        End If
'
'        .MoveNext
'      Loop
'    End If
'  End With
'
'TidyUpAndExit:
'  ReadEmailLinks = fOK
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function

Public Function ClearDiaryLinks() As Boolean
  ' Clear the column's diary links collection.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  ' Remove all diary link objects from the collection.
  Do While mobjDiaryLinks.Count > 0
    Application.ChangedDiaryLink = True     'MH20080528 Fault 13188
    mobjDiaryLinks.Remove 1
  Loop
  Set mobjDiaryLinks = Nothing
  
  ' Re-instantiate the collection.
  Set mobjDiaryLinks = New Collection
  
TidyUpAndExit:
  ClearDiaryLinks = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


'Public Function ClearEmailLinks() As Boolean
'  ' Clear the column's Email links collection.
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'
'  fOK = True
'
'  ' Remove all Email link objects from the collection.
'  Do While mobjEmailLinks.Count > 0
'    Application.ChangedEmailLink = True   'MH20080528 Fault 13188
'    mobjEmailLinks.Remove 1
'  Loop
'  Set mobjEmailLinks = Nothing
'
'  ' Re-instantiate the collection.
'  Set mobjEmailLinks = New Collection
'
'TidyUpAndExit:
'  ClearEmailLinks = fOK
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function

Public Function WriteColumn() As Boolean
  ' Write the column definition to the local tables.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sIndex As String
  Dim objProperty As Property
  Dim objTable As Table
  
  With recColEdit
    ' Remember the current index.
    sIndex = .Index
    
    If (IsNew And Not IsChanged) Then
      ' Get a new column ID value.
      'ColumnID = UniqueColumnValue("tmpColumns", "columnID")
      ColumnID = GetNextObjectIdentitySeed("ASRSysColumns")
            
      ' Add a new record for the Column.
      fOK = (ColumnID > 0)
      If fOK Then
        .AddNew
        .Fields("Guid").value = "{" & Mid(CreateGUID(), 2, 36) & "}"
        .Fields("Deleted") = False
      End If
            
      ' Create entries for this column in all views on the current table.
      If Not (recViewEdit.BOF And recViewEdit.EOF) Then
        recViewEdit.MoveFirst
        Do While Not recViewEdit.EOF
          If recViewEdit!ViewTableID = gLngTableID Then
          
            With recViewColEdit
          
              .Index = "idxViewColID"
              .Seek "=", recViewEdit!ViewID, gLngColumnID
            
              If .NoMatch Then
                ' Add the column to the view.
                .AddNew
                .Fields("ViewID") = recViewEdit!ViewID
                .Fields("ColumnID") = gLngColumnID
                .Fields("InView") = False
                .Fields("New") = True
                .Fields("Deleted") = False
                .Fields("Changed") = False
                .Update
                
                If Not recViewEdit.Fields("New") Then
                  recViewEdit.Edit
                  recViewEdit.Fields("Changed") = True
                  recViewEdit.Update
                End If
              End If
            
            End With
          End If
        
          recViewEdit.MoveNext
        Loop
      End If

    Else
      ' Find the existing column record.
      fOK = FindColumn
      
      If fOK Then
        ' Edit the column record.
        .Edit
      End If
    End If
    
    ' Write the column definition to the local tables.
    If fOK Then
      IsChanged = True
      
      For Each objProperty In Me.Properties
        If Not IsNull(objProperty.value) Then
          .Fields(objProperty.Name) = objProperty.value
        End If
      Next
      Set objProperty = Nothing
      
      .Update
    End If
  
  End With
      
  If fOK Then
    ' Write the table record.
    Set objTable = New SystemMgr.Table
    objTable.TableID = TableID
    
    fOK = objTable.ReadTable
    
    If fOK Then
      objTable.IsChanged = True
      objTable.LastUpdated = Now
      fOK = objTable.WriteTable
    End If
  
    Set objTable = Nothing
  End If
  
  ' Write the Column Control Values to the local tables.
  If fOK Then
    fOK = WriteControlValues
  End If
      
  ' Write the Column Diary Links to the local tables.
  If fOK Then
    fOK = WriteDiaryLinks
  End If
      
'  ' Write the Column Email Links to the local tables.
'  If fOK Then
'    fOK = WriteEmailLinks
'  End If
      
TidyUpAndExit:
  ' Disassociate object variables.
  Set objProperty = Nothing
  Set objTable = Nothing
  
  'Restore original index.
  recColEdit.Index = sIndex
  
  WriteColumn = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Public Function WriteColumn_Transaction() As Boolean
  ' Transaction wrapper for the 'WriteColumn' function.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  fOK = WriteColumn
  
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  WriteColumn_Transaction = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Sub Class_Initialize()

  ' Create a new collection to hold the columns diary links.
  Set mobjDiaryLinks = New Collection
  'Set mobjEmailLinks = New Collection

End Sub

Private Sub Class_Terminate()

  ' Disassociate any object variables.
  Set gMVarProperties = Nothing
  Set mobjDiaryLinks = Nothing
  'Set mobjEmailLinks = Nothing

End Sub


Public Property Get ColumnDataType() As ExpressionValueTypes

  Select Case Me.Properties("dataType").value
    Case dtVARCHAR
      ColumnDataType = giEXPRVALUE_CHARACTER
  
    Case dtNUMERIC
      ColumnDataType = giEXPRVALUE_NUMERIC
 
    Case dtINTEGER
      ColumnDataType = giEXPRVALUE_NUMERIC
    
    Case dtBIT
      ColumnDataType = giEXPRVALUE_LOGIC
 
    Case dtTIMESTAMP
      ColumnDataType = giEXPRVALUE_DATE
    
    Case dtLONGVARBINARY
      ColumnDataType = giEXPRVALUE_OLE
 
    Case dtVARBINARY
      ColumnDataType = giEXPRVALUE_PHOTO
     
    Case dtLONGVARCHAR
      ColumnDataType = giEXPRVALUE_CHARACTER
  
    Case Else
      ColumnDataType = giEXPRVALUE_UNDEFINED
  End Select
  
End Property


Public Property Get ControlValuesString() As String

  ' Return the string that represents the column control
  ' values in a combo or radio control.
  ControlValuesString = gsControlValues
  
End Property

Public Property Let ControlValuesString(ByVal psNewValue As String)

  ' Set the string that represents the column control
  ' values in a combo or radio control.
  gsControlValues = psNewValue
  
End Property

Private Function WriteControlValues() As Boolean
  ' Write the Column's control values to the local tables.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sValues As String
  Dim sSubString As String
  Dim iIndex As Integer
  Dim iSequence As Integer
  
  fOK = True
  iSequence = 1
  
  ' JPD - Fault 10520 (Jonny assures me he's tested this change. I'll trust him...)
  ' Delete all existing control values for the current column.
  daoDb.Execute "DELETE FROM tmpControlValues WHERE columnID=" & ColumnID, dbFailOnError
  
  ' Write the column control values to the database.
  If Len(Trim(gsControlValues)) > 0 Then
  
    sValues = Trim(gsControlValues)
    
    With recContValEdit
          
      ' Add records for each control value in the string.
      While Len(sValues) > 0
        
        iIndex = InStr(sValues, vbCr & vbLf)
      
        If iIndex > 0 Then
          sSubString = Left(sValues, iIndex - 1)
          sValues = Mid(sValues, iIndex + 2)
        Else
          sSubString = sValues
          sValues = ""
        End If
          
        If Len(sSubString) > 0 Then
          .AddNew
          .Fields("columnID") = ColumnID
          .Fields("value") = sSubString
          .Fields("sequence") = iSequence
          .Update
        
          iSequence = iSequence + 1
        End If
        
      Wend
      
    End With
  End If
  
ExitWriteControlValues:
  WriteControlValues = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume ExitWriteControlValues

End Function

Private Function WriteDiaryLinks() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objLink As cDiaryLink
  
  fOK = True
  
  ' Write the column control values to the database.
  If mobjDiaryLinks.Count > 0 Then
  
    ' Delete all existing control values for the current column.
    daoDb.Execute "DELETE FROM tmpDiary WHERE columnID=" & ColumnID, dbFailOnError
    
    ' Add records for each Diary Link in the collection.
    For Each objLink In mobjDiaryLinks
      objLink.ColumnID = gLngColumnID
      fOK = objLink.WriteDiaryLink
      
      If Not fOK Then
        Exit For
      End If
    Next objLink
    Set objLink = Nothing
  End If
  
TidyUpAndExit:
  WriteDiaryLinks = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function


'Private Function WriteEmailLinks() As Boolean
'  On Error GoTo ErrorTrap
'
'  Dim objLink As clsEmailLink
'  Dim strSQL As String
'  Dim strSearch As String
'  Dim fOK As Boolean
'
'  fOK = True
'
'    strSearch = "FROM tmpEMailLinks WHERE columnID = " & CStr(ColumnID)
'
'
'    'MH20090520
'    strSQL = "DELETE FROM tmpLinkContent WHERE ContentID IN (SELECT SubjectContentID " & strSearch & ")"
'    daoDb.Execute strSQL, dbFailOnError
'
'    strSQL = "DELETE FROM tmpLinkContent WHERE ContentID IN (SELECT BodyContentID " & strSearch & ")"
'    daoDb.Execute strSQL, dbFailOnError
'
'
'
'    strSQL = "DELETE FROM tmpEmailLinksRecipients WHERE LinkID IN (SELECT LinkID " & strSearch & ")"
'    daoDb.Execute strSQL, dbFailOnError
'
'
'    'strSQL = "DELETE FROM tmpEmailLinksAttachments WHERE LinkID IN (SELECT LinkID " & strSearch & ")"
'    'daoDb.Execute strSQL, dbFailOnError
'
'    ' Delete all existing control values for the current column.
'    strSQL = "DELETE " & strSearch
'    daoDb.Execute strSQL, dbFailOnError
'
'
''    With recEmailLinksEdit
''      If Not .BOF And .EOF Then
''        .MoveFirst
''        Do While Not .EOF
''          If !ColumnID = ColumnID Then
''
''            With recEmailAttachmentsEdit
''              If Not .BOF And .EOF Then
''                .MoveFirst
''                Do While Not .EOF
''                  If !LinkID = recEmailLinksEdit!LinkID Then
''                    .Delete
''                  Else
''                    .MoveNext
''                  End If
''                Loop
''              End If
''            End With
''
''            With recEmailRecipientsEdit
''              If Not .BOF And .EOF Then
''                .MoveFirst
''                Do While Not .EOF
''                  If !LinkID = recEmailLinksEdit!LinkID Then
''                    .Delete
''                  Else
''                    .MoveNext
''                  End If
''                Loop
''              End If
''            End With
''
''            .Delete
''          Else
''            .MoveNext
''          End If
''        Loop
''      End If
''    End With
'
'
'
'
'
'
'
'
'
'
'
'  ' Write the column control values to the database.
'  If mobjEmailLinks.Count > 0 Then
'
'    ' Add records for each Email Link in the collection.
'    For Each objLink In mobjEmailLinks
'      objLink.ColumnID = ColumnID
'      fOK = objLink.WriteEMailLink
'
'      If Not fOK Then
'        Exit For
'      End If
'    Next objLink
'    Set objLink = Nothing
'  End If
'
'TidyUpAndExit:
'  WriteEmailLinks = fOK
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  Resume TidyUpAndExit
'
'End Function


Public Function CloneColumn(pbAddCopyOfText As Boolean) As Column

    ' JDM - 27/03/01 -  Added parameter whereby we can force "copy of " to appear in front of the column name

  ' Copy the column.
  On Error GoTo ErrorTrap
    
  Dim fOK As Boolean
  Dim objNewColumn As Column
  Dim objDiaryLinks As Collection
  Dim objSourceLink As cDiaryLink
  Dim objLinkCopy As cDiaryLink
  
  'Dim objEmailLinks As Collection
  'Dim objEmailSourceLink As clsEmailLink
  'Dim objEmailLinkCopy As clsEmailLink
  
  
  ' Read the current column definition.
  fOK = ReadColumn
  If fOK Then
    ' Instantiate a column objet to be the clone.
    Set objNewColumn = New SystemMgr.Column
    ' Copy properties from the current column object to the clone.
    With objNewColumn
      Set .Properties = Properties
      .IsNew = True
      .IsChanged = False
      .IsDeleted = False
      .ColumnID = 0
      .ControlValuesString = ControlValuesString

      ' Add "copy of" text in description
      If pbAddCopyOfText Then
        .Properties("ColumnName").value = "Copy_of_" + .Properties("ColumnName").value
      End If

    End With
      
    ' Copy the Diary Link objects.
    Set objDiaryLinks = New Collection
    For Each objSourceLink In mobjDiaryLinks
      Set objLinkCopy = New cDiaryLink
      With objLinkCopy
        .ColumnID = 0
        .Comment = objSourceLink.Comment
        .DiaryLinkId = 0
        .Offset = objSourceLink.Offset
        .Period = objSourceLink.Period
        .Reminder = objSourceLink.Reminder
        .FilterID = objSourceLink.FilterID
        .EffectiveDate = objSourceLink.EffectiveDate
      End With
      objDiaryLinks.Add objLinkCopy
        
      Set objLinkCopy = Nothing
    Next objSourceLink
    
    Set objNewColumn.DiaryLinks = objDiaryLinks
    Set objDiaryLinks = Nothing
    Set objSourceLink = Nothing
  
'    ' Copy the Email Link objects.
'    Set objEmailLinks = New Collection
'    For Each objEmailSourceLink In mobjEmailLinks
'
'      'MH20071016
'      'Set objEmailLinkCopy = New clsEmailLink
'      'Set objEmailLinkCopy = objEmailSourceLink.CloneLink
'      'objEmailLinks.Add objEmailLinkCopy
'      'Set objEmailLinkCopy = New clsEmailLink
'      Set objEmailLinkCopy = objEmailSourceLink.CloneLink
'      objEmailLinkCopy.LinkID = objEmailLinkCopy.GetNewLinkID(objEmailLinks)
'      objEmailLinks.Add objEmailLinkCopy, "ID" & objEmailLinkCopy.LinkID
'
'      Set objEmailLinkCopy = Nothing
'    Next objEmailSourceLink
'
'    Set objNewColumn.EmailLinks = objEmailLinks
'    Set objEmailLinks = Nothing
'    Set objEmailSourceLink = Nothing

  End If
  
TidyUpAndExit:
  If fOK Then
    Set CloneColumn = objNewColumn
  Else
    Set CloneColumn = Nothing
  End If
  ' Disassociate object variables.
  Set objNewColumn = Nothing
  Set objSourceLink = Nothing
  Set objLinkCopy = Nothing
  Set objDiaryLinks = Nothing
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Property Get DiaryLinks() As Collection
  ' Return the collection of diary links.
  Set DiaryLinks = mobjDiaryLinks
End Property
Public Property Set DiaryLinks(pobjNewCollection As Collection)
  ' Set the collection of diary links.
  Set mobjDiaryLinks = pobjNewCollection
End Property


'Public Property Get EmailLinks() As Collection
'  ' Return the collection of Email links.
'  Set EmailLinks = mobjEmailLinks
'End Property
'Public Property Set EmailLinks(pobjNewCollection As Collection)
'  ' Set the collection of Email links.
'  Set mobjEmailLinks = pobjNewCollection
'End Property

Public Function PrintDefinition(piOutputType As SystemMgr.OutputDefintionTypes) As Boolean

  On Error GoTo ErrorTrap

  Dim bOK As Boolean
  Dim frmEdit As frmColEdit

  Set frmEdit = New frmColEdit
  Set frmEdit.Column = Me

  Select Case piOutputType
  
    Case giEXPORT_TO_PRINTER
      frmEdit.PrintDefinition
      
    Case giEXPORT_TO_CLIPBOARD
      frmEdit.CopyDefinitionToClipboard
  
    Case giEXPORT_TO_WORD
  
  End Select

  'JPD 20031008 Fault 7174
  UnLoad frmEdit
  
TidyUpAndExit:
  Set frmEdit = Nothing
  PrintDefinition = bOK
  Exit Function

ErrorTrap:
  bOK = False
  Resume TidyUpAndExit

End Function
  
Public Property Get Locked() As Boolean
  If Properties("Locked") Is Nothing Then
    Locked = False
  Else
    Locked = Properties("Locked").value
  End If
End Property

