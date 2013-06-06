Attribute VB_Name = "modAuditAccess"
Option Explicit

Public gsPassword As String
Public gsServerName As String


Public Sub AuditAccess(strAction As String, pstrModuleName As String)

On Error GoTo ErrTrap

Dim pstrSQL As String

  'JPD 20050812 Fault 10166
  pstrSQL = "INSERT INTO AsrSysAuditAccess (DateTimeStamp,UserGroup,UserName,ComputerName,HRProModule,Action) " & _
            "VALUES (GetDate(), " & _
            "'" & gsSecurityGroup & "', " & _
            "'" & Replace(gsUserName, "'", "''") & "', " & _
            "LOWER(HOST_NAME()), " & _
            "'" & pstrModuleName & "', " & _
            "'" & strAction & "')"
  gADOCon.Execute pstrSQL, , adExecuteNoRecords
  
Exit Sub

ErrTrap:

  MsgBox "Error recording module access in the audit log." & vbCrLf & "Contact support stating :" & vbCrLf & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Access Audit Log"
 
End Sub

Private Function GetUserDetails() As String
  ' Return the current user's user group.
  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim sUserGroup As String
  Dim rsUser As New ADODB.Recordset
  
  'JPD 20050812 Fault 10166
  sSQL = "exec sp_helpuser '" & Replace(Trim(gsUserName), "'", "''") & "'"
    
  rsUser.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  
  Do While rsUser!GroupName = "db_owner" _
        Or LCase(Left(rsUser!GroupName, 6)) = "asrsys"
    rsUser.MoveNext
  Loop

  sUserGroup = rsUser!GroupName
  rsUser.Close
  
TidyUpAndExit:
  Set rsUser = Nothing
  GetUserDetails = sUserGroup
  Exit Function
  
ErrorTrap:
  sUserGroup = "<none>"
  Resume TidyUpAndExit

End Function

Public Function GetTableColumnName(lngColumnID As Long) As String
  
  Dim lTableID As Long
  Dim sColName As String
  Dim sTableName As String
  
  On Error GoTo ErrorTrap
  
  sColName = vbNullString
  sTableName = vbNullString
  
  ' Get the tablename
  With recColEdit
    .Index = "idxColumnID"
    .Seek "=", lngColumnID
    If Not .NoMatch Then
      sColName = !ColumnName
      lTableID = !TableID
    End If
  End With
    
  ' Now get the tablename
  With recTabEdit
    .Index = "idxTableID"
    .Seek "=", lTableID
    If Not .NoMatch Then
      sTableName = !TableName
    End If
  End With
  
  If sColName <> vbNullString And sTableName <> vbNullString Then
    GetTableColumnName = sTableName & "." & sColName
  Else
    GetTableColumnName = ""
  End If
  
TidyUpAndExit:
  Exit Function

ErrorTrap:
  
  GetTableColumnName = "<Unknown>"
  Resume TidyUpAndExit

End Function

Public Function ConfigureCustomAuditLog() As Boolean

  Dim fOK As Boolean
  Dim sSQL As String
  Dim lngAuditTableID As Long
  Dim sAuditTableName As String
  
  Dim saryToList() As String
  Dim saryFromList() As String
  Dim iDefined As Integer
  
  Dim sDateColumn As String
  Dim sTimeColumn As String
  Dim sUserColumn As String
  Dim sTableColumn As String
  Dim sColumnColumn As String
  Dim sOldValueColumn As String
  Dim sNewValueColumn As String
  Dim sModuleColumn As String
  Dim sDescriptionColumn As String
  Dim sIDColumn As String

  fOK = True
  iDefined = 0
  ReDim saryFromList(9)
  ReDim saryToList(9)

  ' Get audit module definition
  sAuditTableName = GetModuleSetupValue(gsMODULEKEY_AUDIT, gsPARAMETERKEY_AUDITTABLE, "T")
  If Len(sAuditTableName) > 0 Then
  
    sDateColumn = GetModuleSetupValue(gsMODULEKEY_AUDIT, gsPARAMETERKEY_AUDITDATECOLUMN, "ColumnNameOnly")
    If Len(sDateColumn) > 0 Then
      saryFromList(iDefined) = "DATEADD(D, 0, DateDiff(D, 0, [DateTimeStamp]))" & vbNewLine
      saryToList(iDefined) = sDateColumn
      iDefined = iDefined + 1
    End If
  
    sTimeColumn = GetModuleSetupValue(gsMODULEKEY_AUDIT, gsPARAMETERKEY_AUDITTIMECOLUMN, "ColumnNameOnly")
    If Len(sTimeColumn) > 0 Then
      saryFromList(iDefined) = "SUBSTRING(CONVERT(varchar(24), [datetimestamp], 121), 12, 8)" & vbNewLine
      saryToList(iDefined) = sTimeColumn
      iDefined = iDefined + 1
    End If
  
    sUserColumn = GetModuleSetupValue(gsMODULEKEY_AUDIT, gsPARAMETERKEY_AUDITUSERCOLUMN, "ColumnNameOnly")
    If Len(sUserColumn) > 0 Then
      saryFromList(iDefined) = "CONVERT(varchar(50),[UserName])" & vbNewLine
      saryToList(iDefined) = sUserColumn
      iDefined = iDefined + 1
    End If
  
    sTableColumn = GetModuleSetupValue(gsMODULEKEY_AUDIT, gsPARAMETERKEY_AUDITTABLECOLUMN, "ColumnNameOnly")
    If Len(sTableColumn) > 0 Then
      saryFromList(iDefined) = "[TableName]" & vbNewLine
      saryToList(iDefined) = sTableColumn
      iDefined = iDefined + 1
    End If
  
    sColumnColumn = GetModuleSetupValue(gsMODULEKEY_AUDIT, gsPARAMETERKEY_AUDITCOLUMNCOLUMN, "ColumnNameOnly")
    If Len(sColumnColumn) > 0 Then
      saryFromList(iDefined) = "[ColumnName]" & vbNewLine
      saryToList(iDefined) = sColumnColumn
      iDefined = iDefined + 1
    End If
  
    sOldValueColumn = GetModuleSetupValue(gsMODULEKEY_AUDIT, gsPARAMETERKEY_AUDITOLDVALUECOLUMN, "ColumnNameOnly")
    If Len(sOldValueColumn) > 0 Then
      saryFromList(iDefined) = "[OldValue]" & vbNewLine
      saryToList(iDefined) = sOldValueColumn
      iDefined = iDefined + 1
    End If
  
    sNewValueColumn = GetModuleSetupValue(gsMODULEKEY_AUDIT, gsPARAMETERKEY_AUDITNEWVALUECOLUMN, "ColumnNameOnly")
    If Len(sNewValueColumn) > 0 Then
      saryFromList(iDefined) = "[NewValue]" & vbNewLine
      saryToList(iDefined) = sNewValueColumn
      iDefined = iDefined + 1
    End If
  
    sModuleColumn = GetModuleSetupValue(gsMODULEKEY_AUDIT, gsPARAMETERKEY_AUDITMODULECOLUMN, "ColumnNameOnly")
    If Len(sModuleColumn) > 0 Then
      saryFromList(iDefined) = "CONVERT(varchar(50),APP_NAME())" & vbNewLine
      saryToList(iDefined) = sModuleColumn
      iDefined = iDefined + 1
    End If
     
    sDescriptionColumn = GetModuleSetupValue(gsMODULEKEY_AUDIT, gsPARAMETERKEY_AUDITDESCRIPTIONCOLUMN, "ColumnNameOnly")
    If Len(sDescriptionColumn) > 0 Then
      saryFromList(iDefined) = "[RecordDesc]" & vbNewLine
      saryToList(iDefined) = sDescriptionColumn
      iDefined = iDefined + 1
    End If
     
    sIDColumn = GetModuleSetupValue(gsMODULEKEY_AUDIT, gsPARAMETERKEY_AUDITIDCOLUMN, "ColumnNameOnly")
    If Len(sIDColumn) > 0 Then
      saryFromList(iDefined) = "[ID]" & vbNewLine
      saryToList(iDefined) = sIDColumn
      iDefined = iDefined + 1
    End If
     
     
    If iDefined > 0 Then
    
      ReDim Preserve saryFromList(iDefined - 1)
      ReDim Preserve saryToList(iDefined - 1)
    

'        & "        SELECT " & Join(saryFromList, ", ") & vbNewLine _


      ' Drop the auto generated view
      sSQL = "IF EXISTS (SELECT Name FROM sysobjects " & _
              "WHERE id = object_id('dbo." & sAuditTableName & "') " & _
              "AND sysstat & 0xf = 2) " & _
              "DROP VIEW dbo." & sAuditTableName
      gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
  
  
      ' Recreate using a new view
      sSQL = "CREATE VIEW dbo.[" & sAuditTableName & "]" & vbNewLine _
          & "WITH SCHEMABINDING" & vbNewLine _
          & "AS" & vbNewLine _
          & "SELECT [ID] AS [ID]" & vbNewLine _
          & ", [ID] AS [" & sIDColumn & "]" & vbNewLine _
          & ",[UserName] AS [" & sUserColumn & "]" & vbNewLine _
          & ",[DateTimeStamp] AS [" & sDateColumn & "]" & vbNewLine _
          & ",[Tablename] AS [" & sTableColumn & "]" & vbNewLine _
          & ",[Columnname] AS [" & sColumnColumn & "]" & vbNewLine _
          & ",[OldValue] AS [" & sOldValueColumn & "]" & vbNewLine _
          & ",[NewValue] AS [" & sNewValueColumn & "]" & vbNewLine _
          & ",'' AS [" & sModuleColumn & "]" & vbNewLine _
          & ",[RecordDesc] AS [" & sDescriptionColumn & "]" & vbNewLine _
          & "FROM [dbo].[ASRSysAuditTrail]"
        gADOCon.Execute sSQL, , adCmdText + adExecuteNoRecords
        
    End If
  End If


TidyUpAndExit:
  ConfigureCustomAuditLog = fOK
  Exit Function

ErrorTrap:
  
  ConfigureCustomAuditLog = fOK
  Resume TidyUpAndExit

End Function
