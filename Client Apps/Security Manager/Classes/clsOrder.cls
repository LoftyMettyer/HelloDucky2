VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Order definition variables.
Private mlngOrderID As Long
Private msOrderName As String
Private mlngTableID As Long
Private msTableName As String
Private miOrderType As Integer
Private mcolOrderItems As Collection
Private mlngTimeStamp As Long

' Object handing variables.
Private mfConstructed As Boolean

Public Property Get TableName() As String
  ' Return the order column's table name.
  TableName = msTableName
  
End Property

Public Property Let TableName(ByVal psNewValue As String)
  ' Set the order column's table name.
  msTableName = psNewValue

End Property
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



Public Property Get TableID() As Long
  ' Return the table id property.
  TableID = mlngTableID
  
End Property

Public Function SelectOrder() As Boolean
  ' Display the order selection form.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim fExit As Boolean
  Dim lngOldOrderID As Long
  Dim sSQL As String
  Dim rsOrders As ADODB.Recordset
  'Dim frmSelectOrder As frmExprOrdSelection
  Dim frmSelectOrder As frmDefSel
  
  fExit = False
  lngOldOrderID = mlngOrderID
  
  'Set frmSelectOrder = New frmExprOrdSelection
  Set frmSelectOrder = New frmDefSel
  
  Do While Not fExit
    With frmSelectOrder
      ' Set the selection form properties.
      'Set .Order = Me

      ' JPD 6/6/00 Added edtSelect option so that user can actually select an order.
      '.Options = edtAdd + edtDelete + edtEdit + edtCopy + edtPrint + edtSelect + edtDeselect
      .EnableRun = False
      .TableComboEnabled = False
      .TableComboVisible = False
      .HideDescription = True
      .HelpContextID = 8066

      sSQL = "SELECT name, orderID" & _
               " FROM ASRSysOrders" & _
               " WHERE tableID = " & CStr(TableID) & _
               " AND type = " & CStr(OrderType)
      
      '.ShowList "ORDERS"
      .ShowOrders sSQL, OrderID
      
      ' Display the selection form.
      .Show vbModal
      
      Select Case .Action
        Case edtAdd
          NewOrder
          .SelectedID = OrderID
          
        Case edtEdit
          OrderID = .SelectedID
          If .FromCopy Then
            CopyOrder
            If OrderID > 0 Then
              .SelectedID = OrderID
            End If
          Else
            EditOrder
          End If
          
        'Case edtCopy
        '  OrderID = .SelectedID
        '  CopyOrder
          
        Case edtPrint
          OrderID = .SelectedID
          PrintOrder
        
        Case edtSelect
          OrderID = .SelectedID
          fOK = True
          fExit = True
      
        ' Deselect all expressions and return to the calling form.
        Case edtDeselect
          OrderID = 0
          fOK = True
          fExit = True
          
        Case edtCancel
          ' Check if the original order still exists.
          sSQL = "SELECT name" & _
            " FROM ASRSysOrders" & _
            " WHERE orderID = " & Trim(Str(lngOldOrderID))
          Set rsOrders = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
          If rsOrders.EOF And rsOrders.BOF Then
            OrderID = 0
          Else
            OrderID = lngOldOrderID
          End If
          
          rsOrders.Close
          fOK = False
          fExit = True
          
      End Select
    End With
  Loop
    
  Set frmSelectOrder = Nothing

TidyUpAndExit:
  SelectOrder = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function
Public Sub CopyOrder()
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sName As String
  Dim frmEdit As frmOrder
  
  ' Begin the transaction of data.
  gADOCon.BeginTrans
  
  fOK = ConstructOrder
  
  'Find & read table record to copy
  If fOK Then
    mlngOrderID = 0
    sName = msOrderName
    msOrderName = "Copy_of_" & Trim(sName)
  
    Set frmEdit = New frmOrder
    With frmEdit
      Set .Order = Me
      .Show vbModal
      fOK = Not .Cancelled
    End With
    Set frmEdit = Nothing
    
    If fOK Then
      fOK = WriteOrder
    End If
    
    If Not fOK Then
      mlngOrderID = 0
      msOrderName = sName
    End If
 End If
  
ExitCopyOrder:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    gADOCon.CommitTrans
  Else
    gADOCon.RollbackTrans
  End If
  Exit Sub

ErrorTrap:
  fOK = False
  Resume ExitCopyOrder

End Sub


Private Sub PrintOrder()
  ' Print the order definition.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fFound As Boolean
  Dim iItemCount As Integer
  Dim iSequence As Integer
  Dim objOrderItem As clsOrderItem
  Dim objPrinter As clsPrintDef

  ' Read the order to print.
  fOK = ConstructOrder

  If fOK Then
    ' Load the printer object
    Set objPrinter = New clsPrintDef
    With objPrinter
      If .IsOK Then
        If .PrintStart(True) Then
          .PrintHeader "Order Definition : " & Trim(msOrderName)
          .PrintNormal
          .PrintTitle "Table"
          .PrintNonBold Trim(Me.TableName)
  
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



Private Sub NewOrder()
  ' Create a new order.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim frmOrd As frmOrder
  
  ' Begin the transaction of data to the local database.
  gADOCon.BeginTrans
  
  OrderID = 0
  msOrderName = vbNullString
  fOK = ClearOrderItems
  mfConstructed = True

  If fOK Then
    Set frmOrd = New frmOrder
    Set frmOrd.Order = Me
    frmOrd.Show vbModal
    fOK = Not frmOrd.Cancelled
  End If
  
  If fOK Then
    fOK = WriteOrder
    
    If Not fOK Then
      OrderID = 0
      msOrderName = vbNullString
      ClearOrderItems
      mfConstructed = True
    End If
  End If
  Set frmOrd = Nothing

TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    gADOCon.CommitTrans
  Else
    gADOCon.RollbackTrans
  End If
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub


Private Sub EditOrder()
  ' Edit the order.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim frmOrd As frmOrder
  
  ' Begin the transaction of data.
  gADOCon.BeginTrans
  
  fOK = ConstructOrder
  
  If fOK Then
    Set frmOrd = New frmOrder
    Set frmOrd.Order = Me
    frmOrd.Show vbModal
    
    fOK = Not frmOrd.Cancelled
      
    If fOK Then
      fOK = WriteOrder
    End If
    
  End If
  
TidyUpAndExit:
  Set frmOrd = Nothing
  ' Commit the data transaction if everything was okay.
  If fOK Then
    'TM20011008 Fault 2554
    'Reset the constructed flag.
    mfConstructed = False
    gADOCon.CommitTrans
  Else
    gADOCon.RollbackTrans
  End If
  Exit Sub

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub


Public Sub DeleteOrder()
  ' Delete the current order object.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
   
  ' Begin the transaction of data to the local database.
  gADOCon.BeginTrans
  
  ' Check that the order is not in use.
  fOK = Not OrderIsUsed
  
  If fOK Then
    ' Delete the Order.
    sSQL = "DELETE FROM ASRSysOrders" & _
      " WHERE orderID = " & Trim(Str(mlngOrderID))
    gADOCon.Execute sSQL, , adCmdText
    
    ' Delete the Order Items for this order.
    sSQL = "DELETE FROM ASRSysOrderItems" & _
      " WHERE orderID = " & Trim(Str(mlngOrderID))
    gADOCon.Execute sSQL, , adCmdText
  End If
  
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    gADOCon.CommitTrans
  Else
    gADOCon.RollbackTrans
  End If
  Exit Sub

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub

Private Function OrderIsUsed() As Boolean
  ' Return true if the order is used somewhere and therefore cannot be deleted.
  '
  ' Orders may be used in the following contexts :
  '
  '   Default Table order
  '   Screen order
  '   Expression field component order
  '   Module definitions.
  On Error GoTo ErrorTrap
  
  Dim fUsed As Boolean
  Dim sSQL As String
  Dim sExprType As String
  Dim rsCheck As ADODB.Recordset
  Dim sModuleName As String

  ' Check if any tables use this order as the default table order.
  sSQL = "SELECT tableName" & _
    " FROM ASRSysTables" & _
    " WHERE defaultOrderID = " & Trim(Str(mlngOrderID))
  Set rsCheck = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rsCheck
    fUsed = Not (.EOF And .BOF)
  
    If fUsed Then
      ' Tell the user why the order cannot be deleted.
      MsgBox "This order cannot be deleted." & vbCr & vbCrLf & _
        "It is used as the default order for the '" & !TableName & "' table.", _
        vbExclamation + vbOKOnly, App.ProductName
    End If
  
    .Close
  End With
  Set rsCheck = Nothing
  
  ' Check that the order is not used by any screens.
  If Not fUsed Then
    ' Find any screens that use this order.
    sSQL = "SELECT ASRSysScreens.name" & _
      " FROM ASRSysScreens" & _
      " WHERE ASRSysScreens.orderID = " & Trim(Str(mlngOrderID))
    Set rsCheck = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    With rsCheck
      fUsed = Not (.EOF And .BOF)
    
      If fUsed Then
        ' Tell the user why the order cannot be deleted.
        MsgBox "This order cannot be deleted." & vbCrLf & vbCr & _
          "It is used as the order for the '" & !Name & "' screen.", _
          vbExclamation + vbOKOnly, App.ProductName
      End If
    
      .Close
    End With
    Set rsCheck = Nothing
  End If
  
  ' Check that the order is not used by any expression field components.
  If Not fUsed Then
    ' Find any expressions that use this order.
    sSQL = "SELECT ASRSysExpressions.name, ASRSysExpressions.type, ASRSysTables.tableName" & _
      " FROM ASRSysExpressions" & _
      " JOIN ASRSysTables ON ASRSysExpressions.TableID = ASRSysTables.tableID" & _
      " JOIN ASRSysExprComponents ON ASRSysExpressions.exprID = ASRSysExprComponents.exprID" & _
      " WHERE ASRSysExprComponents.fieldSelectionOrderID = " & Trim(Str(mlngOrderID)) & _
      " AND ASRSysExprComponents.type = " & Trim(Str(giCOMPONENT_FIELD))
    Set rsCheck = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    With rsCheck
      fUsed = Not (.EOF And .BOF)
    
      If fUsed Then
        sExprType = LCase(ExpressionTypeName(!Type))
        
        ' Tell the user why the order cannot be deleted.
        MsgBox "This order cannot be deleted." & vbCr & vbCr & _
          "It is used as the order for a field component in the " & sExprType & " '" & !Name & "'," & vbCrLf & _
          "which is owned by the '" & !TableName & "' table.", _
          vbExclamation + vbOKOnly, App.ProductName
      End If
      
      .Close
    End With
    Set rsCheck = Nothing
  End If
  
  ' Check that the order is not used by any Mail Merge definitions.
  If Not fUsed Then
    ' Find any Mail Merge definitions that use this order.
    sSQL = "SELECT ASRSysMailMergeName.name, ASRSysTables.tableName" & _
      " FROM ASRSysMailMergeName" & _
      " INNER JOIN ASRSysTables ON ASRSysMailMergeName.tableID = ASRSysTables.tableID" & _
      " WHERE ASRSysMailMergeName.orderID = " & Trim(Str(mlngOrderID))

    Set rsCheck = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    With rsCheck
      fUsed = Not (.EOF And .BOF)
    
      If fUsed Then
        ' Tell the user why the order cannot be deleted.
        MsgBox "This order cannot be deleted." & vbCr & vbCr & _
          "It is used in the '" & !Name & "' Mail Merge definition," & vbCrLf & _
          "which is owned by the '" & !TableName & "' table.", _
          vbExclamation + vbOKOnly, App.ProductName
      End If
      
      .Close
    End With
    Set rsCheck = Nothing
  End If
  
  ' Check that the order is not used in any Module definitions.
  If Not fUsed Then
    sSQL = "SELECT moduleKey" & _
      " FROM ASRSysModuleSetup" & _
      " WHERE parameterType = '" & gsPARAMETERTYPE_ORDERID & "'" & _
      " AND parameterValue = '" & Trim(Str(mlngOrderID)) & "'"
    Set rsCheck = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    With rsCheck
    
      If Not (.BOF And .EOF) Then
        fUsed = True
       
        ' Tell the user why the column cannot be deleted.
        MsgBox "This order cannot be deleted." & vbCr & _
          "It is used in the module definition.", _
          vbExclamation + vbOKOnly, App.ProductName
      End If
      
      ' Close the recordset.
      .Close
    End With
    Set rsCheck = Nothing
  End If

TidyUpAndExit:
  ' Disassociate object variables.
  Set rsCheck = Nothing
  
  OrderIsUsed = fUsed
  Exit Function

ErrorTrap:
  fUsed = True
  Resume TidyUpAndExit
  
End Function






Public Function WriteOrder() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fFindColumnsExist As Boolean
  Dim sSQL As String
  Dim objItem As clsOrderItem
  
  fOK = True
  
  If mlngOrderID = 0 Then
    
    
    'MH20010712 Need keep manual record of allocated IDs incase users
    'in SYS MGR have created expressions but not yet saved changes
    'OrderID = UniqueColumnValue("ASRSysOrders", "orderID")
    OrderID = GetUniqueID("Orders", "ASRSysOrders", "orderID")
    
    
    ' Add a record for the new order.
    fOK = (mlngOrderID > 0)
    
    If fOK Then
      sSQL = "INSERT INTO ASRSysOrders" & _
        " (orderID, name, tableID, type)" & _
        " VALUES(" & Trim(Str(mlngOrderID)) & ", " & _
        "'" & Replace(Trim(msOrderName), "'", "''") & "', " & _
        Trim(Str(mlngTableID)) & ", " & _
        Trim(Str(miOrderType)) & ")"
      gADOCon.Execute sSQL, , adCmdText
    End If
  Else
    sSQL = "UPDATE ASRSysOrders" & _
      " SET name = '" & Replace(Trim(msOrderName), "'", "''") & "'" & _
      " WHERE orderID = " & Trim(Str(mlngOrderID))
    gADOCon.Execute sSQL, , adCmdText
  End If
  
  If fOK Then
    ' Delete any existing order items for this order.
    sSQL = "DELETE FROM ASRSysOrderItems" & _
      " WHERE orderID = " & Trim(Str(mlngOrderID))
    gADOCon.Execute sSQL, , adCmdText
    
    ' Add any order items for this order.
    fFindColumnsExist = False
    For Each objItem In mcolOrderItems
      With objItem
        sSQL = "INSERT INTO ASRSysOrderItems" & _
          " (orderID, columnID, type, sequence, ascending)" & _
          " VALUES(" & Trim(Str(mlngOrderID)) & "," & _
          " " & Trim(Str(.ColumnID)) & "," & _
          " '" & Trim(.ItemType) & "'," & _
          " " & Trim(Str(.Sequence)) & "," & _
          " " & IIf(.Ascending, "1", "0") & ")"
        gADOCon.Execute sSQL, , adCmdText
        
        If .ItemType = "F" Then
          fFindColumnsExist = True
        End If
      End With
    Next
    Set objItem = Nothing
    
    ' If no Find columns are defined, then just use the Order columns.
    If Not fFindColumnsExist Then
      For Each objItem In mcolOrderItems
        With objItem
          sSQL = "INSERT INTO ASRSysOrderItems" & _
            " (orderID, columnID, type, sequence, ascending)" & _
            " VALUES(" & Trim(Str(mlngOrderID)) & "," & _
            " " & Trim(Str(.ColumnID)) & "," & _
            " 'F'," & _
            " " & Trim(Str(.Sequence)) & "," & _
            " " & IIf(.Ascending, "1", "0") & ")"
          gADOCon.Execute sSQL, , adCmdText
        End With
      Next
      Set objItem = Nothing
    End If
  End If
      
TidyUpAndExit:
  WriteOrder = fOK
  Exit Function

ErrorTrap:
  MsgBox Err.Description, vbExclamation + vbOKOnly, App.ProductName
  fOK = False
  Resume TidyUpAndExit

End Function


Public Property Get OrderName() As String
  If Not mfConstructed Then
    ConstructOrder
  End If
  
  OrderName = msOrderName
  
End Property

Public Property Get Timestamp() As Long
  ' Return the order's timestamp value.
  Timestamp = mlngTimeStamp
  
End Property

Public Property Let Timestamp(plngNewValue As Long)
  ' Set the order's timestamp property.
  mlngTimeStamp = plngNewValue
  
End Property
Public Property Let OrderName(psName As String)
  msOrderName = psName
  
End Property

Public Function ConstructOrder() As Boolean
  ' Read the order definition into local objects.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sIndex As String
  Dim sSQL As String
  Dim rsOrder As ADODB.Recordset
  Dim rsOrderItems As ADODB.Recordset
    
  fOK = True
  
  If Not mfConstructed Then
    ' Set default values.
    msOrderName = vbNullString
    mlngTableID = 0
    
    ' Get the order definition.
    sSQL = "SELECT ASRSysOrders.*, ASRSysTables.tableName, " & _
      " CONVERT(integer, timestamp) AS intTimestamp" & _
      " FROM ASRSysOrders" & _
      " INNER JOIN ASRSysTables ON ASRSysOrders.tableID = ASRSysTables.tableID" & _
      " WHERE ASRSysOrders.orderID = " & Trim(Str(mlngOrderID))
    Set rsOrder = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    With rsOrder
      fOK = Not (.EOF And .BOF)
    
      If fOK Then
        msOrderName = !Name
        mlngTableID = !TableID
        msTableName = !TableName
        miOrderType = !Type
        mlngTimeStamp = IIf(IsNull(!intTimeStamp), 0, !intTimeStamp)
      
        fOK = ClearOrderItems
      End If
    
      .Close
    End With
    Set rsOrder = Nothing
      
    If fOK Then
      ' Read the order items.
      sSQL = "SELECT ASRSysOrderItems.*, ASRSysColumns.columnName, ASRSysTables.tableName" & _
        " FROM ASRSysOrderItems" & _
        " INNER JOIN ASRSysColumns" & _
        "   INNER JOIN ASRSysTables" & _
        "   ON ASRSysColumns.tableID = ASRSysTables.tableID" & _
        " ON ASRSysOrderItems.columnID = ASRSysColumns.columnID" & _
        " WHERE ASRSysOrderItems.orderID = " & Trim(Str(mlngOrderID) & _
        " ORDER BY ASRSysOrderItems.Type ASC, ASRSysOrderItems.Sequence ASC")
      Set rsOrderItems = modExpression.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
      With rsOrderItems
        Do While Not .EOF
          AddOrderItem !ColumnID, !Type, !Sequence, !Ascending, !ColumnName, !TableName, IIf(IsNull(!Editable), False, !Editable)
          .MoveNext
        Loop
      
        .Close
      End With
      Set rsOrderItems = Nothing
    End If
  End If
  
TidyUpAndExit:
  mfConstructed = fOK
  ConstructOrder = fOK
  Exit Function

ErrorTrap:
  MsgBox "Error constructing order.", _
    vbExclamation + vbOKOnly, App.ProductName
  fOK = False
  Resume TidyUpAndExit
  
End Function
Public Property Get OrderItems() As Collection
  If mcolOrderItems Is Nothing Then
    Set mcolOrderItems = New Collection
  End If
  Set OrderItems = mcolOrderItems
  
End Property

Public Property Set OrderItems(pcolNewItems As Collection)
  Set mcolOrderItems = pcolNewItems
  
End Property

Public Function AddOrderItem(plngColumnID As Long, psItemType As String, _
  piSequence As Integer, pfAscending As Boolean, psColumnName As String, _
  psTableName As String, pbIsEditable As Boolean) As clsOrderItem
  ' Add the given column to the order definition.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objNewItem As clsOrderItem
  
  fOK = True
  
  'Create a new order item object
  Set objNewItem = New clsOrderItem
  
  'Set the properties of the new order item object
  With objNewItem
    .ColumnID = plngColumnID
    .ColumnName = Trim(psColumnName)
    .ItemType = psItemType
    .Sequence = piSequence
    .Ascending = pfAscending
    .TableName = psTableName
    .IsEditable = pbIsEditable
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


Public Property Let TableID(pLngNewID As Long)
  ' Set the table id property.
  mlngTableID = pLngNewID
  
End Property

Public Property Get OrderID() As Long
  ' Return the order ID property.
  OrderID = mlngOrderID
  
End Property

Public Property Let OrderID(plngNewValue As Long)
  ' Set the order ID property.
  If mlngOrderID <> plngNewValue Then
    mlngOrderID = plngNewValue
    mfConstructed = False
  End If
  
End Property

Private Sub Class_Initialize()
  'Create the collection
  Set mcolOrderItems = New Collection
  mfConstructed = False
  
End Sub


Private Sub Class_Terminate()
  'Destroy the collection
  Set mcolOrderItems = Nothing

End Sub

Public Property Get OrderType() As Integer
  ' Return the order type.
  OrderType = miOrderType
  
End Property

Public Property Let OrderType(ByVal piNewValue As Integer)
  ' Set the order type.
  miOrderType = piNewValue

End Property

Public Property Get ContainsEditableColumns() As Boolean

  Dim item As clsOrderItem
  Dim bFound As Boolean
  
  bFound = False
 
  For Each item In mcolOrderItems
    bFound = bFound Or item.IsEditable
  Next item

  ContainsEditableColumns = bFound
End Property
