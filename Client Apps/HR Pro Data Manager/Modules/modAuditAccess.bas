Attribute VB_Name = "modAuditAccess"
Option Explicit

Public Enum LogType
  iLOGIN = 0
  iLOGOFF = 1
  iRECONNECT = 2
  iDISCONNECTED = 3
End Enum

Public Sub AuditAccess(piLogType As LogType, pstrModuleName As String)

  On Error GoTo ErrTrap

  Dim strSQL As String
  Dim sLogType As String
  
  Select Case piLogType
    Case iLOGIN
      sLogType = "Log In"
    Case iLOGOFF
      sLogType = "Log Out"
    Case iRECONNECT
      sLogType = "Reconnected"
    Case iDISCONNECTED
      sLogType = "Connection Dropped"
  End Select

  strSQL = "INSERT INTO AsrSysAuditAccess (DateTimeStamp,UserGroup,UserName,ComputerName,HRProModule,Action) " & _
            "VALUES (GetDate(), " & _
            "'" & gsUserGroup & "', " & _
            "'" & datGeneral.UserNameForSQL & "', " & _
            "LOWER(HOST_NAME()), " & _
            "'" & pstrModuleName & "', " & _
            "'" & sLogType & "')"
  
  gADOCon.Execute strSQL
  
Exit Sub

ErrTrap:

  MsgBox "Error recording module access in the audit log." & vbNewLine & "Contact support stating :" & vbNewLine & vbNewLine & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Access Audit Log"
 
End Sub



