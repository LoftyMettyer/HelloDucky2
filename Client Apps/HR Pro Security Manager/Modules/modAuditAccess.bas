Attribute VB_Name = "modAuditAccess"
Option Explicit

Public Sub AuditAccess(strAction As String, pstrModuleName As String)

  On Error GoTo ErrorTrap

  Dim pstrSQL As String

  pstrSQL = "INSERT INTO AsrSysAuditAccess (DateTimeStamp,UserGroup,UserName,ComputerName,HRProModule,Action) " & _
            "VALUES (GetDate(), " & _
            "'" & gsUserGroup & "', " & _
            "'" & Replace(gsUserName, "'", "''") & "', " & _
            "LOWER(HOST_NAME()), " & _
            "'" & pstrModuleName & "', " & _
            "'" & strAction & "')"
  gADOCon.Execute pstrSQL, , adExecuteNoRecords
  
  Exit Sub

ErrorTrap:

  MsgBox "Error recording module access in the audit log." & vbCrLf & "Contact support stating :" & vbCrLf & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Access Audit Log"
 
End Sub
