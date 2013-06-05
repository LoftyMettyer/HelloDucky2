Attribute VB_Name = "modSave_Email"
Option Explicit

Public Function SaveEmailAddrs(mfrmUse As frmUsage) As Boolean
  ' Save the new or modified Email Address definitions to the server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objEmailAddr As clsEmailAddr
  
  fOK = True
  
  With recEmailAddrEdit
    If Not (.BOF And .EOF) Then
      .MoveFirst
    End If
    Do While fOK And Not .EOF
      If !Deleted Then
        Set objEmailAddr = New clsEmailAddr
        objEmailAddr.EmailID = !EmailID
        Set mfrmUse = New frmUsage
        mfrmUse.ResetList
        If objEmailAddr.EmailIsUsed(mfrmUse) Then
          gobjProgress.Visible = False
          Screen.MousePointer = vbDefault
          mfrmUse.ShowMessage !Name & " Email", "The email cannot be deleted as the email is used by the following:", UsageCheckObject.Email
          fOK = False
        End If
        UnLoad mfrmUse
        Set mfrmUse = Nothing
       
        gobjProgress.Visible = True
        
        If fOK Then
          fOK = EmailAddrDelete
        End If
        
      ElseIf !New Then
        fOK = EmailAddrNew
      ElseIf !Changed Then
        fOK = EmailAddrSave
      End If
      
      .MoveNext
    Loop
  End With
  
TidyUpAndExit:
  SaveEmailAddrs = fOK
  Exit Function
  
ErrorTrap:
  'MsgBox "Error creating email addresses" & _
         IIf(Trim(Err.Description) <> vbnullstring, "(" & Err.Description & ")", vbnullstring), vbCritical
  OutputError "Error creating email addresses"
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function EmailAddrDelete() As Boolean
  ' Delete the current Order definition from the server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  
  fOK = True
  
  ' Delete the existing order definition from the server database.
  sSQL = "DELETE FROM ASRSysEmailAddress" & _
    " WHERE EmailID=" & recEmailAddrEdit!EmailID
  gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
TidyUpAndExit:
  EmailAddrDelete = fOK
  Exit Function

ErrorTrap:
  fOK = False
  OutputError "Error Deleting email address"
  Resume TidyUpAndExit
  
End Function

Private Function EmailAddrNew() As Boolean
  ' Write the current Order definition to the server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iColumn As Integer
  Dim sName As String
  Dim rsEmailAddrs As ADODB.Recordset
  
  Set rsEmailAddrs = New ADODB.Recordset
  fOK = True
  
  ' Open the order definition table on the server.
  rsEmailAddrs.Open "ASRSysEmailAddress", gADOCon, adOpenForwardOnly, adLockOptimistic, adCmdTableDirect

  ' Add the new order definition.
  With rsEmailAddrs
    .AddNew
    .Fields("EmailID").Value = recEmailAddrEdit!EmailID

    For iColumn = 0 To .Fields.Count - 1
      sName = .Fields(iColumn).Name
      
      If Not UCase$(Trim$(sName)) = "TIMESTAMP" Then
        If Not IsNull(recEmailAddrEdit.Fields(sName).Value) Then
          .Fields(iColumn).Value = recEmailAddrEdit.Fields(sName).Value
        End If
      End If
    Next iColumn
    .Update
  End With
  rsEmailAddrs.Close
  
TidyUpAndExit:
  Set rsEmailAddrs = Nothing
  EmailAddrNew = fOK
  Exit Function

ErrorTrap:
  OutputError "Error saving new email address"
  fOK = False
  Resume TidyUpAndExit
  
End Function


Private Function EmailAddrSave() As Boolean
  ' Save the current order to the server database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  ' Delete the existing record in the server database.
  fOK = EmailAddrDelete
  
  If fOK Then
    ' Create the new record in the server database.
    fOK = EmailAddrNew
  End If

TidyUpAndExit:
  EmailAddrSave = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

