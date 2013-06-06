VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Order definition variables.
Private mlngOrderID As Long
Private msOrderName As String
Private mlngTableID As Long
Private miOrderType As Integer
Private mcolOrderItems As Collection

' Object handing variables.
Private mfNew As Boolean
Private mfChanged As Boolean
Private mfDeleted As Boolean

Private mfrmUse As frmUsage

Private mlngUtilityID As Long
Private mactlAllWorkflowElements() As VB.Control

Public Property Get OrderItems() As Collection
  If mcolOrderItems Is Nothing Then
    Set mcolOrderItems = New Collection
  End If
  Set OrderItems = mcolOrderItems
  
End Property

Public Property Set OrderItems(NewItems As Collection)
  Set mcolOrderItems = NewItems
  
End Property


Public Property Get OrderType() As Integer
  ' Return the order type.
  OrderType = miOrderType
  
End Property

Public Property Let AllWorkflowElements(pactlElements As Variant)
  mactlAllWorkflowElements = pactlElements
  
End Property




Public Property Get AllWorkflowElements() As Variant
  AllWorkflowElements = mactlAllWorkflowElements
  
End Property

Public Property Let OrderType(ByVal piNewValue As Integer)
  ' Set the order type.
  miOrderType = piNewValue

End Property

Public Property Get OrderID() As Long

  ' Return the order ID property.
  OrderID = mlngOrderID
  
End Property

Public Property Let OrderID(ID As Long)

  ' Set the order ID property.
  mlngOrderID = ID
  
End Property

Public Property Get OrderName() As String
  OrderName = msOrderName
  
End Property

Public Property Let OrderName(psName As String)
  msOrderName = psName
  
End Property

Public Property Get TableID() As Long
  
  ' Return the table id property.
  TableID = mlngTableID
  
End Property

Public Property Let TableID(pLngNewID As Long)
  
  ' Set the table id property.
  mlngTableID = pLngNewID
  
End Property





Public Property Get IsNew() As Boolean
  IsNew = mfNew
  
End Property

Public Property Get IsChanged() As Boolean
  IsChanged = mfChanged
  
End Property

Public Property Let IsChanged(pfValue As Boolean)
  mfChanged = pfValue
  
End Property

Public Property Let IsNew(pfSetting As Boolean)
  mfNew = pfSetting
  
End Property
Public Property Let IsDeleted(pfSetting As Boolean)
  mfDeleted = pfSetting
  
End Property


Public Property Get IsDeleted() As Boolean
  IsDeleted = mfDeleted
  
End Property

Public Function CopyOrder() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sName As String
  Dim frmEdit As frmOrdEdit
  
  Screen.MousePointer = vbHourglass

  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  fOK = ConstructOrder
  
  'Find & read table record to copy
  If fOK Then
    mfNew = True
    mfChanged = False
    mfDeleted = False
    
    mlngOrderID = 0
    sName = msOrderName
    msOrderName = "Copy_of_" & Trim(sName)
  
    Set frmEdit = New frmOrdEdit
    Set frmEdit.Order = Me
    
    Screen.MousePointer = vbDefault
    
    frmEdit.Show vbModal
    fOK = Not frmEdit.Cancelled
    Set frmEdit = Nothing
    
    If fOK Then
      fOK = WriteOrder
    End If
    
    If Not fOK Then
      mfNew = True
      mfChanged = False
      mfDeleted = False
      mlngOrderID = 0
      msOrderName = sName
    End If
 
 End If
  
ExitCopyOrder:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  CopyOrder = fOK
  Screen.MousePointer = vbDefault
  Exit Function

ErrorTrap:
  fOK = False
  Resume ExitCopyOrder

End Function

Public Function DeleteOrder() As Boolean
  ' Delet the current order object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  fOK = True
  
  'TM19012004
  ' Check that the order is not in use.
  If fOK Then
    Set mfrmUse = New frmUsage
    mfrmUse.ResetList
    If OrderIsUsed(mfrmUse) Then
      Screen.MousePointer = vbDefault
      mfrmUse.ShowMessage Me.OrderName & " Order", "The order cannot be deleted as the order is used by the following:", UsageCheckObject.Order
      fOK = False
    End If
    UnLoad mfrmUse
    Set mfrmUse = Nothing
  End If
  
  If fOK Then
    'Find & read order record to delete
    fOK = ConstructOrder
    
    If fOK Then
      'Check if table is already deleted
      If Not IsDeleted Then
        'Set record deleted flag
        mfDeleted = True
        
        'Write order record
        fOK = WriteOrder
      End If
    End If
  
    If fOK Then
      'Delete order items for this order
      daoDb.Execute "DELETE FROM tmpOrderItems WHERE orderID = " & Trim(Str(mlngOrderID)), dbFailOnError
    End If
  End If
  
TidyUpAndExit:
  If fOK Then
    mfDeleted = True
  End If
  DeleteOrder = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function OrderIsUsed(pfrmUsage As frmUsage) As Boolean
  
  ' Return true if the order is used somewhere and
  ' therefore cannot be deleted.
  '
  ' Orders may be used in the following contexts :
  '
  '   Default Table order
  '   Screen order
  '   Expression field component order
  '   Module definitions.
  '   Mail Merge definitions (no longer used)
  '   Link Column definitions
  On Error GoTo ErrorTrap
  
  Set mfrmUse = pfrmUsage
  
  Dim fUsed As Boolean
  Dim lngExprID As Long
  Dim lngExprBaseTableID As Long
  Dim sSQL As String
  Dim sExprName As String
  Dim sExprType As String
  Dim sExprParentTable As String
  Dim objExpr As CExpression
  Dim objComp As CExprComponent
  Dim rsCheck As DAO.Recordset
  Dim sModuleName As String
  Dim rsUtils1 As ADODB.Recordset
  Dim rsUtils2 As ADODB.Recordset
  Dim rsUtils3 As ADODB.Recordset
  Dim sAccess As String
  Dim strExpressions As String
  Dim lngLoop As Long
  Dim wfTempElement As VB.Control
  Dim iLoop2 As Integer
  Dim iLoop3 As Integer
  Dim asItems() As String
  Dim asWorkflowUsage() As String
  Dim fFound As Boolean
  
  strExpressions = ""
  
  ' Find any tables that use this order as the
  ' default table order.
  sSQL = "SELECT DISTINCT tableName" & _
    " FROM tmpTables" & _
    " WHERE deleted = FALSE" & _
    " AND defaultOrderID = " & Trim(Str(mlngOrderID))
  Set rsCheck = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  If Not (rsCheck.BOF And rsCheck.EOF) Then
    fUsed = True
    Do Until rsCheck.EOF
      mfrmUse.AddToList ("Table Default Order : " & rsCheck!TableName)
      rsCheck.MoveNext
    Loop
  End If
  ' Close the recordset.
  rsCheck.Close
  
  
  ' Check that the order is not used by any screens.
  ' Find any screen that use this order.
  sSQL = "SELECT DISTINCT tmpScreens.name, tmpTables.tableName" & _
    " FROM tmpScreens, tmpTables" & _
    " WHERE tmpScreens.deleted = FALSE" & _
    " AND tmpScreens.orderID = " & Trim(Str(mlngOrderID)) & _
    " AND tmpScreens.tableID = tmpTables.tableID"
  Set rsCheck = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  If Not (rsCheck.BOF And rsCheck.EOF) Then
    fUsed = True
    Do Until rsCheck.EOF
      mfrmUse.AddToList ("Screen Order : " & rsCheck!Name & " <" & rsCheck!TableName & ">")
      rsCheck.MoveNext
    Loop
  End If
  ' Close the recordset.
  rsCheck.Close
  
  
  ' Check that the order is not used by any expression field components.
  ' Find any expression field components that use this order.
  sSQL = "SELECT DISTINCT tmpComponents.componentID" & _
    " FROM tmpComponents, tmpExpressions " & _
    " WHERE (tmpComponents.type = " & Trim(Str(giCOMPONENT_FIELD)) & _
    "   OR tmpComponents.type = " & Trim(Str(giCOMPONENT_WORKFLOWFIELD)) & ")" & _
    "   AND tmpExpressions.exprid = tmpComponents.Exprid " & _
    "   AND (tmpExpressions.type NOT IN (" & giEXPR_RUNTIMECALCULATION & "," & giEXPR_RUNTIMEFILTER & "," & giEXPR_RECORDINDEPENDANTCALC & "," & _
                                             giEXPR_UTILRUNTIMEFILTER & "," & giEXPR_MATCHJOINEXPRESSION & "," & giEXPR_MATCHSCOREEXPRESSION & "," & _
                                             giEXPR_MATCHWHEREEXPRESSION & ")) " & _
    "   AND tmpComponents.fieldSelectionOrderID = " & Trim(Str(mlngOrderID))
  Set rsCheck = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  
  Do Until rsCheck.EOF
    fUsed = True
      
    Set objComp = New CExprComponent
    objComp.ComponentID = rsCheck!ComponentID
    lngExprID = objComp.RootExpressionID
    Set objComp = Nothing
    
    ' Get the expression name and type description.
    Set objExpr = New CExpression
    objExpr.ExpressionID = lngExprID
    
    If objExpr.ReadExpressionDetails Then
      sExprName = objExpr.Name
      sExprType = objExpr.ExpressionTypeName
      lngExprBaseTableID = objExpr.BaseTableID
      
      If (objExpr.ExpressionType = giEXPR_WORKFLOWCALCULATION) _
        Or (objExpr.ExpressionType = giEXPR_WORKFLOWRUNTIMEFILTER) _
        Or (objExpr.ExpressionType = giEXPR_WORKFLOWSTATICFILTER) Then
      
        ' Get the expression's parent table name.
        recWorkflowEdit.Index = "idxWorkflowID"
        recWorkflowEdit.Seek "=", objExpr.UtilityID
           
        If Not recWorkflowEdit.NoMatch Then
          sExprParentTable = recWorkflowEdit!Name
        End If
      Else
        ' Get the expression's parent table name.
        recTabEdit.Index = "idxTableID"
        recTabEdit.Seek "=", lngExprBaseTableID
           
        If Not recTabEdit.NoMatch Then
          sExprParentTable = recTabEdit!TableName
        End If
      End If
    Else
      sExprName = "<unknown>"
      sExprType = "<unknown>"
      sExprParentTable = "<unknown>"
    End If
    ' Disassociate object variables.
    Set objExpr = Nothing
    
    If InStr(1, strExpressions, CStr(" " & lngExprID & " ")) = 0 Then
      mfrmUse.AddToList (sExprType & " : " & sExprName & " <" & sExprParentTable & ">")
      strExpressions = " " & strExpressions & CStr(lngExprID) & " "
    End If
  
    rsCheck.MoveNext
  Loop
  ' Close the recordset.
  rsCheck.Close
  
  
  ' Check that the order is not used in any Module definitions.
  sSQL = "SELECT DISTINCT moduleKey" & _
    " FROM tmpModuleSetup" & _
    " WHERE parameterType = '" & gsPARAMETERTYPE_ORDERID & "'" & _
    " AND parameterValue = '" & Trim(Str(mlngOrderID)) & "'"
  Set rsCheck = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  If Not (rsCheck.BOF And rsCheck.EOF) Then
    fUsed = True
    Do Until rsCheck.EOF
      Select Case rsCheck!moduleKey
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
      rsCheck.MoveNext
    Loop
  End If
  ' Close the recordset.
  rsCheck.Close


  ' Find any tables' columns that use this order as the
  ' link order.
  sSQL = "SELECT tmpTables.tableName, tmpColumns.columnName" & _
    " FROM tmpTables, tmpColumns" & _
    " WHERE tmpTables.deleted = FALSE" & _
    "     AND tmpColumns.deleted = FALSE" & _
    "     AND tmpColumns.linkOrderID = " & Trim(Str(mlngOrderID)) & _
    "     AND tmpColumns.tableID = tmpTables.tableID"
  Set rsCheck = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  If Not (rsCheck.BOF And rsCheck.EOF) Then
    fUsed = True
    Do Until rsCheck.EOF
      mfrmUse.AddToList "Link Order : " & rsCheck!ColumnName & " <" & rsCheck!TableName & ">"
      rsCheck.MoveNext
    Loop
  End If
  ' Close the recordset.
  rsCheck.Close
  
  
  ' Now Check if the order is used in a Record Profile
  sSQL = "SELECT DISTINCT ASRSysRecordProfileName.Name, ASRSysRecordProfileName.Username, ASRSysRecordProfileName.recordProfileID " & _
    "FROM ASRSysRecordProfileName " & _
    "   INNER JOIN ASRSysRecordProfileTables " & _
    "   ON ASRSysRecordProfileName.recordProfileID = ASRSysRecordProfileTables.recordProfileID " & _
    "WHERE ASRSysRecordProfileTables.orderID = " & Trim(Str(mlngOrderID)) & " " & _
    "   OR ASRSysRecordProfileName.orderID = " & Trim(Str(mlngOrderID))
  Set rsUtils1 = New ADODB.Recordset
  rsUtils1.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsUtils1
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Record Profile : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsUtils1 = Nothing


  ' Now Check if the order is used in a CustomReport
  sSQL = "SELECT DISTINCT ASRSysCustomReportsName.Name, ASRSysCustomReportsName.Username, ASRSysCustomReportsName.ID " & _
    "FROM ASRSysCustomReportsName " & _
    "   INNER JOIN ASRSysCustomReportsChildDetails " & _
    "   ON ASRSysCustomReportsName.ID = ASRSysCustomReportsChildDetails.customReportID " & _
    "WHERE ASRSysCustomReportsChildDetails.childOrder = " & Trim(Str(mlngOrderID))
  Set rsUtils2 = New ADODB.Recordset
  rsUtils2.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsUtils2
    If Not (.EOF And .BOF) Then
      fUsed = True
      Do Until .EOF
        mfrmUse.AddToList ("Custom Report : " & !Name)
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsUtils2 = Nothing

  ' Now Check if the order is used in a Runtime Calculation OR Runtime Filter.
  sSQL = "SELECT DISTINCT ASRSysExpressions.exprID, ASRSysExpressions.Type, ASRSysExpressions.Name, ASRSysTables.TableName " & _
        " FROM ASRSysExpressions " & _
        "       LEFT OUTER JOIN  ASRSysExprComponents " & _
        "       ON ASRSysExpressions.ExprID = ASRSysExprComponents.ExprID " & _
        "       LEFT OUTER JOIN ASRSysTables" & _
        "       ON ASRSysExpressions.TableID = ASRSysTables.TableID" & _
        " WHERE ASRSysExpressions.Type IN (" & giEXPR_RUNTIMECALCULATION & "," & giEXPR_RUNTIMEFILTER & ", " & giEXPR_RECORDINDEPENDANTCALC & ") " & _
        "   AND (ASRSysExprComponents.fieldSelectionOrderID = " & Trim(Str(mlngOrderID)) & ")"
  Set rsUtils3 = New ADODB.Recordset
  rsUtils3.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  With rsUtils3
    If Not (rsUtils3.BOF And rsUtils3.EOF) Then
      fUsed = True
      Do Until .EOF
        If InStr(1, strExpressions, CStr(" " & lngExprID & " ")) = 0 Then
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
          strExpressions = " " & strExpressions & CStr(lngExprID) & " "
        End If
        .MoveNext
      Loop
    End If
    .Close
  End With
  Set rsUtils3 = Nothing

  ' Find any saved workflows that use this order as a record selector order.
  ReDim asWorkflowUsage(3, 0)
  sSQL = "SELECT tmpWorkflows.name," & _
    "  tmpWorkflowElements.caption," & _
    "  tmpWorkflowElementItems.ItemType," & _
    "  tmpWorkflowElementItems.identifier" & _
    " FROM tmpWorkflowElementItems, tmpWorkflowElements, tmpWorkflows" & _
    " WHERE tmpWorkflows.deleted = FALSE" & _
    "  AND (tmpWorkflowElementItems.recordOrderID = " & CStr(mlngOrderID) & " OR tmpWorkflowElementItems.LookupOrderID = " & CStr(mlngOrderID) & ")" & _
    "  AND tmpWorkflowElementItems.elementID = tmpWorkflowElements.ID" & _
    "  AND tmpWorkflowElements.workflowID = tmpWorkflows.ID"
  Set rsCheck = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  If Not (rsCheck.BOF And rsCheck.EOF) Then
    fUsed = True
    Do Until rsCheck.EOF
      mfrmUse.AddToList "Workflow : " & rsCheck!Name & " ('" & rsCheck!Caption & "' web form, '" & rsCheck!Identifier & "' " & IIf(rsCheck("ItemType") = 14, "lookup", "record selector") & ") : Record order"
      
      ReDim Preserve asWorkflowUsage(3, UBound(asWorkflowUsage, 2) + 1)
      asWorkflowUsage(0, UBound(asWorkflowUsage, 2)) = UCase(Trim(rsCheck!Name))
      asWorkflowUsage(1, UBound(asWorkflowUsage, 2)) = UCase(Trim(rsCheck!Caption))
      asWorkflowUsage(2, UBound(asWorkflowUsage, 2)) = UCase(Trim(rsCheck!Identifier))
      
      rsCheck.MoveNext
    Loop
  End If
  ' Close the recordset.
  rsCheck.Close
  
  ' Check the current workflow if this order is used as a record selector order.
  If mlngUtilityID > 0 Then
    If IsArray(mactlAllWorkflowElements) Then
      For lngLoop = 1 To UBound(mactlAllWorkflowElements)
        Set wfTempElement = mactlAllWorkflowElements(lngLoop)
    
        Select Case wfTempElement.ElementType
          Case elem_WebForm
            asItems = wfTempElement.Items
            For iLoop2 = 1 To UBound(asItems, 2)
              If ((asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_GRID) And (CLng(asItems(52, iLoop2)) = mlngOrderID)) Or _
              ((asItems(2, iLoop2) = giWFFORMITEM_INPUTVALUE_LOOKUP) And (CLng(asItems(80, iLoop2)) = mlngOrderID)) Then
    
                fUsed = True
                
                fFound = False
                For iLoop3 = 1 To UBound(asWorkflowUsage, 2)
                  If asWorkflowUsage(0, iLoop3) = UCase(Trim(GetWorkflowName(mlngUtilityID))) _
                    And asWorkflowUsage(1, iLoop3) = UCase(Trim(wfTempElement.Caption)) _
                    And asWorkflowUsage(2, iLoop3) = UCase(Trim(asItems(9, iLoop2))) Then
                    
                    fFound = True
                    Exit For
                  End If
                Next iLoop3
                
                If Not fFound Then
                  mfrmUse.AddToList "Workflow : " & GetWorkflowName(mlngUtilityID) & " ('" & wfTempElement.Caption & "' web form, '" & asItems(9, iLoop2) & "' " & IIf(asItems(2, iLoop2) = 14, "lookup", "record selector") & ") : Record order"
                End If
              End If
            Next iLoop2
        End Select
    
        Set wfTempElement = Nothing
      Next lngLoop
    End If
  End If
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set rsCheck = Nothing
  Set objExpr = Nothing
  
  OrderIsUsed = fUsed
  Exit Function

ErrorTrap:
  fUsed = True
  Resume TidyUpAndExit
  
End Function





Public Property Let UtilityID(ByVal plngNewValue As Long)
  mlngUtilityID = plngNewValue
  
End Property


Public Property Get UtilityID() As Long
  UtilityID = mlngUtilityID
  
End Property
Public Function EditOrder() As Boolean
  ' Edit the order.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim frmEdit As frmOrdEdit
  
  Screen.MousePointer = vbHourglass
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  fOK = ConstructOrder
  
  If fOK Then
    Set frmEdit = New frmOrdEdit
    Set frmEdit.Order = Me
    
    Screen.MousePointer = vbDefault
    frmEdit.Show vbModal
    fOK = Not frmEdit.Cancelled
      
    If fOK Then
      fOK = WriteOrder
    End If
    
    Set frmEdit = Nothing
  End If
  
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  EditOrder = fOK
  Screen.MousePointer = vbDefault
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Public Function FindOrder() As Boolean
  On Error GoTo ErrorTrap
  
  With recOrdEdit
    .Index = "idxID"
    .Seek "=", OrderID
    FindOrder = (Not .NoMatch)
  End With
  
  Exit Function
  
ErrorTrap:
  FindOrder = False

End Function

Public Function NewOrder() As Boolean
  ' Create a new order.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim frmEdit As frmOrdEdit
    
  Screen.MousePointer = vbHourglass
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  mfNew = True
  mfChanged = False
  mfDeleted = False
  mlngOrderID = 0
  msOrderName = vbNullString
  fOK = ClearOrderItems
  
  If fOK Then
    Set frmEdit = New frmOrdEdit
    Set frmEdit.Order = Me
    
    Screen.MousePointer = vbDefault
    frmEdit.Show vbModal
    fOK = Not frmEdit.Cancelled
  End If
  
  If fOK Then
    fOK = WriteOrder
    
    If Not fOK Then
      mfNew = True
      mfChanged = False
      mfDeleted = False
      mlngOrderID = 0
      msOrderName = vbNullString
      ClearOrderItems
    End If
  End If
  Set frmEdit = Nothing

TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  NewOrder = fOK
  Screen.MousePointer = vbDefault
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function ConstructOrder() As Boolean
  ' Read the order definition from the local tables.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sIndex As String
  Dim sSQL As String
  Dim rsOrderItems As DAO.Recordset
    
  With recOrdEdit
    'Store current index
    sIndex = .Index
    
    'Find the table record
    fOK = FindOrder
    
    If fOK Then
      'Set table object properties
      msOrderName = !Name
      mlngTableID = !TableID
      mfNew = !New
      mfChanged = !Changed
      mfDeleted = !Deleted
      miOrderType = !Type
    
      fOK = ClearOrderItems
    Else
      'Set table object properties
      msOrderName = vbNullString
      mlngTableID = 0
      mfNew = True
      mfChanged = False
      mfDeleted = False
      miOrderType = giORDERTYPE_STATIC
    End If
  End With
  
  If fOK Then
    sSQL = "SELECT tmpOrderItems.*, tmpColumns.columnName" & _
      " FROM tmpOrderItems, tmpColumns" & _
      " WHERE tmpOrderItems.orderID = " & Trim(Str(mlngOrderID)) & _
      " AND tmpColumns.columnID = tmpOrderItems.columnID" & _
      " ORDER BY tmpOrderItems.Type ASC, tmpOrderItems.Sequence ASC"
    Set rsOrderItems = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    
    With rsOrderItems
      Do While Not .EOF
        AddOrderItem !ColumnID, !Type, !Sequence, !Ascending, !ColumnName
        .MoveNext
      Loop
    End With
    
    rsOrderItems.Close
    Set rsOrderItems = Nothing
  End If
  
ExitConstructOrder:
  'Restore original index
  recOrdEdit.Index = sIndex
  ConstructOrder = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume ExitConstructOrder
  
End Function

Public Function SelectOrder(Optional pvarReadOnly As Variant) As Boolean
  ' Display the order selection form.
  Dim fExit As Boolean
  Dim lngOldOrderID As Long
  Dim frmSelectOrder As frmSelection
  Dim strSQL As String
  Dim fReadOnly As Boolean
  
  If IsMissing(pvarReadOnly) Then
    fReadOnly = False
  Else
    fReadOnly = CBool(pvarReadOnly)
  End If
  
  If TableID > 0 Then
    
    fExit = False
    lngOldOrderID = mlngOrderID
    Set frmSelectOrder = New frmSelection
    
    Do While Not fExit
      ' Set the selection form type.
      ' Set the selection form properties.
      frmSelectOrder.ForcedReadOnly = fReadOnly
      Set frmSelectOrder.Order = Me
            
      frmSelectOrder.Show vbModal
      
      Select Case frmSelectOrder.Action
        
        Case edtAdd
          NewOrder
          
        Case edtEdit
          OrderID = frmSelectOrder.SelectedID
          EditOrder
          
        Case edtCopy
          OrderID = frmSelectOrder.SelectedID
          CopyOrder
          
        Case edtDelete
          OrderID = frmSelectOrder.SelectedID
          
          '02/08/2001 MH Fault 2043
          'DeleteOrder_Transaction
          If MsgBox("Are you sure you want to delete this order ?", _
            vbQuestion + vbYesNo, App.ProductName) = vbYes Then
                DeleteOrder_Transaction
          End If
        
        Case edtPrint
          OrderID = frmSelectOrder.SelectedID
          PrintOrder
        
        Case edtSelect
          OrderID = frmSelectOrder.SelectedID
          SelectOrder = ConstructOrder
          fExit = True
      
        ' Deselect all expressions and return to the calling form.
        Case edtDeselect
          OrderID = -1
          ConstructOrder
          SelectOrder = True
          fExit = True
          
        Case edtCancel
          ' Check if the original order still exists.
          With recOrdEdit
            .Index = "idxID"
            .Seek "=", lngOldOrderID
    
            If .NoMatch Then
              OrderID = 0
            Else
              If !Deleted Then
                OrderID = 0
              End If
            End If
          End With

          SelectOrder = False
          fExit = True
          
      End Select
    Loop
    
    Set frmSelectOrder = Nothing
  End If

End Function

Private Function DeleteOrder_Transaction() As Boolean
  ' Transaction wrapper for the 'DeleteOrder' function.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  fOK = DeleteOrder
  
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  DeleteOrder_Transaction = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Sub PrintOrder()
  ' Print the order definition.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fFound As Boolean
  Dim iItemCount As Integer
  Dim iSequence As Integer
  Dim sSQL As String
  Dim sTableName As String
  Dim rsInfo As DAO.Recordset
  Dim objOrderItem As OrderItem
  Dim objPrinter As clsPrintDef

  ' Read the order to print.
  fOK = ConstructOrder

  If fOK Then
    sSQL = "SELECT tableName" & _
      " FROM tmpTables" & _
      " WHERE tmpTables.tableID = " & Trim(Str(mlngTableID))
    Set rsInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    With rsInfo
      If Not (.EOF And .BOF) Then
        sTableName = !TableName
      Else
        sTableName = "<unknown>"
      End If

      .Close
    End With
    Set rsInfo = Nothing
    
    ' Load the printer object
    Set objPrinter = New clsPrintDef
    With objPrinter
      If .IsOK Then
        If .PrintStart(True) Then
          .PrintHeader "Order Definition : " & Trim(msOrderName)
          .PrintNormal
          .PrintTitle "Table"
          .PrintNonBold Trim(sTableName)
  
          'Print the find items.
          .PrintTitle "Find Window Columns"
          If mcolOrderItems.Count > 0 Then
            For Each objOrderItem In mcolOrderItems
              If (objOrderItem.ItemType = "F") Then
                .PrintNonBold objOrderItem.FullColumnName
              End If
            Next objOrderItem
            Set objOrderItem = Nothing
          Else
            .PrintNonBold "<None>"
          End If
        
          'Print the order items.
          .PrintTitle "Sort Order Columns"
          If mcolOrderItems.Count > 0 Then
            For Each objOrderItem In mcolOrderItems
              If (objOrderItem.ItemType = "O") Then
                .PrintNonBold objOrderItem.FullColumnName
              End If
            Next objOrderItem
            Set objOrderItem = Nothing
          Else
            .PrintNonBold "<None>"
          End If
        
          .PrintEnd
          .PrintConfirm "Order : " & Trim(msOrderName), "Order Definition"
        End If
      End If
    End With
  End If
  
TidyUpAndExit:
  Set objOrderItem = Nothing
  If Not fOK Then
    MsgBox "Unable to print the order '" & msOrderName & "'." & vbCr & vbCr & _
      Err.Description, vbExclamation + vbOKOnly, App.ProductName
  End If
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub


Public Function WriteOrder() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim lngExprID As Long
  Dim sIndex As String
  Dim sSQL As String
  Dim objItem As OrderItem
  Dim objComp As CExprComponent
  Dim rsCheck As DAO.Recordset
  
  With recOrdEdit
    'Store current index
    sIndex = .Index
    
    If (IsNew And Not IsChanged) Then


      'MH20010712
      'Replace the select max from access table with new method....
      'check the SQL database directly (incase anybody in Data Manager
      'has created a new expression.
      '''OrderID = UniqueColumnValue("tmpOrders", "orderID")
      OrderID = GetUniqueID("Orders", "tmpOrders", "orderID")


      'Add a record for new order
      fOK = (OrderID > 0)
      
      If fOK Then
        .AddNew
        !OrderID = OrderID
        !TableID = TableID
      End If
    Else
      'Find existing table record
      fOK = FindOrder
      If fOK Then
        'Edit table record
        .Edit
      End If
    End If
    
    ' Write the order definition to the local tables.
    If fOK Then
      mfChanged = True
      !Name = OrderName
      !New = IsNew
      !Changed = IsChanged
      !Deleted = IsDeleted
      !Type = miOrderType
      .Update
    End If
  End With
  
  If fOK Then
    ' Delete any existing order items for this order.
    daoDb.Execute "DELETE FROM tmpOrderItems WHERE orderID=" & OrderID, dbFailOnError
      
    ' Add any order items for this order.
    With recOrdItemEdit
      For Each objItem In OrderItems
        .AddNew
        !OrderID = OrderID
        !ColumnID = objItem.ColumnID
        !Type = objItem.ItemType
        !Sequence = objItem.Sequence
        !Ascending = objItem.Ascending
        .Update
      Next
      Set objItem = Nothing
    End With
    
    ' Mark as changed any expressions that use this order.
    ' This ensures that the stored procedures are refreshed to reflect the order change.
    sSQL = "SELECT tmpComponents.componentID" & _
      " FROM tmpComponents" & _
      " WHERE tmpComponents.type = " & Trim(Str(giCOMPONENT_FIELD)) & _
      " AND tmpComponents.fieldSelectionOrderID = " & Trim(Str(OrderID))
    Set rsCheck = daoDb.OpenRecordset(sSQL, _
      dbOpenForwardOnly, dbReadOnly)
    Do While Not rsCheck.EOF
      Set objComp = New CExprComponent
      objComp.ComponentID = rsCheck!ComponentID
      lngExprID = objComp.RootExpressionID
      Set objComp = Nothing
    
      sSQL = "UPDATE tmpExpressions" & _
        " SET changed = TRUE" & _
        " WHERE exprID = " & Trim(Str(lngExprID))
        
      daoDb.Execute sSQL, dbFailOnError
        
      rsCheck.MoveNext
    Loop
    rsCheck.Close
    Set rsCheck = Nothing
  End If
      
TidyUpAndExit:
  recOrdEdit.Index = sIndex
  WriteOrder = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Public Function AddOrderItem(plngColumnID As Long, psItemType As String, _
  piSequence As Integer, pfAscending As Boolean, psColumnName As String) As OrderItem
  ' Add the given column to the order definition.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objNewItem As OrderItem
  
  fOK = True
  
  'Create a new order item object
  Set objNewItem = New OrderItem
  
  'Set the properties of the new order item object
  With objNewItem
    .ColumnID = plngColumnID
    .ColumnName = Trim(psColumnName)
    .ItemType = psItemType
    .Sequence = piSequence
    .Ascending = pfAscending
  End With
  
  mcolOrderItems.Add objNewItem, psItemType & "_" & Trim(Str(plngColumnID))
  
TidyUpAndExit:
  'Return the new order item object
  If fOK Then
    Set AddOrderItem = objNewItem
  Else
    Set AddOrderItem = Nothing
  End If
  ' Disassociate object variables.
  Set objNewItem = Nothing
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Function ClearOrderItems() As Boolean
  On Error GoTo ErrorTrap
  
  Do While mcolOrderItems.Count > 0
    mcolOrderItems.Remove 1
  Loop
  Set mcolOrderItems = Nothing
  Set mcolOrderItems = New Collection
  
  ClearOrderItems = True
  Exit Function

ErrorTrap:
  ClearOrderItems = False
  Err = False
  
End Function

Private Sub Class_Initialize()
  'Create the collection
  Set mcolOrderItems = New Collection
  ReDim mactlAllWorkflowElements(0)

End Sub

Private Sub Class_Terminate()
  'Destroy the collection
  Set mcolOrderItems = Nothing
  
End Sub


Public Function CloneOrder(pavCloneRegister As Variant) As Order
  ' Clone the current order.
  ' Used when copying table definitions.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim lngNewColumnID As Long
  Dim objClonedOrder As Order
  Dim objSourceOrderItem As OrderItem
  Dim objCopyOrderItem As OrderItem
  Dim colOrderItems As Collection
  
  ' Read the order definition.
  fOK = ConstructOrder

  If fOK Then
    ' Instantiate the clone order.
    Set objClonedOrder = New Order
    ' Clone the current order's properties.
    With objClonedOrder
      .IsNew = True
      .IsDeleted = False
      .IsChanged = False
      .OrderID = 0
      .OrderName = OrderName
      .OrderType = OrderType
      .ClearOrderItems
      
      ' Instantiate a new collection of order items.
      Set colOrderItems = New Collection
      
      ' Clone the order items, updating the associated column IDs.
      For Each objSourceOrderItem In mcolOrderItems
        ' Find the new column ID for the current order item.
        lngNewColumnID = 0
        For iIndex = 1 To UBound(pavCloneRegister, 2)
          If pavCloneRegister(1, iIndex) = "COLUMN" And _
            pavCloneRegister(2, iIndex) = objSourceOrderItem.ColumnID Then
            lngNewColumnID = pavCloneRegister(3, iIndex)
            Exit For
          End If
        Next iIndex
        
        ' Instantiate the cloned order item.
        Set objCopyOrderItem = New OrderItem
        With objCopyOrderItem
          .ColumnID = lngNewColumnID
          .ItemType = objSourceOrderItem.ItemType
          .Sequence = objSourceOrderItem.Sequence
          .Ascending = objSourceOrderItem.Ascending
        End With
        colOrderItems.Add objCopyOrderItem
        ' Disassociate object variables.
        Set objCopyOrderItem = Nothing
      Next objSourceOrderItem
      
      Set objClonedOrder.OrderItems = colOrderItems
      
      ' Disassociate object variables.
      Set colOrderItems = Nothing
      
    End With
  End If
  
TidyUpAndExit:
  If fOK Then
    Set CloneOrder = objClonedOrder
  Else
    Set CloneOrder = Nothing
  End If
  ' Disassociate object variables.
  Set objClonedOrder = Nothing
  Set objSourceOrderItem = Nothing
  Set objCopyOrderItem = Nothing
  Set colOrderItems = Nothing
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function
