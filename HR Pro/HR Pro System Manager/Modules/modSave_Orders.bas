Attribute VB_Name = "modSave_Orders"
Option Explicit


Public Function SaveOrders(mfrmUse As frmUsage) As Boolean
  ' Save the new or modified Order definitions to the server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objOrder As Order
  
  fOK = True
  
  With recOrdEdit
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    Do While fOK And Not .EOF
      If !Deleted Then
        Set objOrder = New Order
        objOrder.OrderID = !OrderID
        Set mfrmUse = New frmUsage
        mfrmUse.ResetList
        If objOrder.OrderIsUsed(mfrmUse) Then
          gobjProgress.Visible = False
          Screen.MousePointer = vbNormal
          mfrmUse.ShowMessage !Name & " Order", "The order cannot be deleted as the order is used by the following:", UsageCheckObject.Order
          fOK = False
        End If
        UnLoad mfrmUse
        Set mfrmUse = Nothing
        
        gobjProgress.Visible = True
        
        If fOK Then
          fOK = OrderDelete
        End If
        
      ElseIf !New Then
        fOK = OrderNew
      ElseIf !Changed Then
        fOK = OrderSave
      End If
      
      .MoveNext
    Loop
  End With
  
TidyUpAndExit:
  SaveOrders = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error saving orders"
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function OrderDelete() As Boolean
  ' Delete the current Order definition from the server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  
  fOK = True
  
  ' Delete the existing order definition from the server database.
  sSQL = "DELETE FROM ASRSysOrders" & _
    " WHERE orderID=" & recOrdEdit!OrderID
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords

  ' Delete the existing order item definitions from the server database.
  sSQL = "DELETE FROM ASRSysOrderItems" & _
    " WHERE orderID=" & recOrdEdit!OrderID
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
TidyUpAndExit:
  OrderDelete = fOK
  Exit Function

ErrorTrap:
  fOK = False
  OutputError "Error deleting order"
  Resume TidyUpAndExit
  
End Function

Private Function OrderNew() As Boolean
  ' Write the current Order definition to the server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iColumn As Integer
  Dim sName As String
  Dim rsOrders As ADODB.Recordset
  Dim rsOrderItems As ADODB.Recordset
  
  fOK = True
  Set rsOrders = New ADODB.Recordset
  Set rsOrderItems = New ADODB.Recordset
  
  ' Open the order definition table on the server.
  rsOrders.Open "ASRSysOrders", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect
  
  ' Add the new order definition.
  With rsOrders
    .AddNew
    
    For iColumn = 0 To .Fields.Count - 1
      sName = .Fields(iColumn).Name
      
      If Not UCase$(Trim$(sName)) = "TIMESTAMP" Then
        If Not IsNull(recOrdEdit.Fields(sName).Value) Then
          .Fields(iColumn).Value = recOrdEdit.Fields(sName).Value
        End If
      End If
    Next iColumn
    .Update
  End With
  rsOrders.Close
  
  ' Open the order items definition table on the server.
  rsOrderItems.Open "ASRSysOrderItems", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect
    
  'Add item definitions for this order
  recOrdItemEdit.Index = "idxOrderID"
  recOrdItemEdit.Seek "=", recOrdEdit!OrderID
  
  If Not recOrdItemEdit.NoMatch Then
    Do While Not recOrdItemEdit.EOF
      
      ' If no more items for this order exit loop.
      If recOrdItemEdit!OrderID <> recOrdEdit!OrderID Then
        Exit Do
      End If
      
      ' Add the item definition.
      With rsOrderItems
        .AddNew
        
        For iColumn = 0 To .Fields.Count - 1
          sName = .Fields(iColumn).Name
          If Not IsNull(recOrdItemEdit.Fields(sName).Value) Then
            .Fields(iColumn).Value = recOrdItemEdit.Fields(sName).Value
          End If
        Next iColumn
        
        .Update
      End With
      
      ' Get the next item definition.
      recOrdItemEdit.MoveNext
    Loop
  End If
  rsOrderItems.Close
  
TidyUpAndExit:
  Set rsOrders = Nothing
  Set rsOrderItems = Nothing
  OrderNew = fOK
  Exit Function

ErrorTrap:
  OutputError "Error creating new order"
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function OrderSave() As Boolean
  ' Save the current order to the server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  ' Delete the existing record in the server database.
  fOK = OrderDelete
  
  If fOK Then
    ' Create the new record in the server database.
    fOK = OrderNew
  End If

TidyUpAndExit:
  OrderSave = fOK
  Exit Function
  
ErrorTrap:
  OutputError "Error updating order"
  fOK = False
  Resume TidyUpAndExit
  
End Function
