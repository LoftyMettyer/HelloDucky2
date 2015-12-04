Attribute VB_Name = "modSQLObjects"
Option Explicit


Public Function DropProcedure(strSPName As String) As Boolean

  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  
  fOK = True
  
  sSQL = "IF EXISTS (SELECT Name FROM sysobjects" & _
    "    WHERE id = object_id(N'" & Replace(strSPName, "'", "''") & "')" & _
    "    AND objectproperty(id, N'IsProcedure') = 1)" & _
    "  DROP PROCEDURE " & Replace(strSPName, "'", "''")
  gADOCon.Execute sSQL, , adExecuteNoRecords

TidyUpAndExit:
  DropProcedure = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  OutputError "Error dropping stored procedure '" & strSPName & "'"
  Resume TidyUpAndExit

End Function


Public Function DropFunction(strFnName As String) As Boolean
  
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  
  fOK = True
  
  If glngSQLVersion >= 8 Then
    sSQL = "IF EXISTS" & _
      " (SELECT Name" & _
      "   FROM sysobjects" & _
      "   WHERE id = object_id('" & strFnName & "')" & _
      "     AND sysstat & 0xf = 0)" & _
      " DROP FUNCTION " & strFnName
    gADOCon.Execute sSQL, , adExecuteNoRecords
  End If

TidyUpAndExit:
  DropFunction = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  OutputError "Error dropping function '" & strFnName & "'"
  Resume TidyUpAndExit

End Function

Public Function DropView(viewName As String) As Boolean

  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  
  fOK = True
  
  sSQL = "IF EXISTS" & _
    " (SELECT Name" & _
    "   FROM sysobjects" & _
    "   WHERE id = object_id('" & viewName & "')" & _
    "     AND xtype = 'V')" & _
    " DROP VIEW " & viewName
  gADOCon.Execute sSQL, , adExecuteNoRecords


TidyUpAndExit:
  DropView = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  OutputError "Error dropping function '" & viewName & "'"
  Resume TidyUpAndExit

End Function


Public Function UniqueSQLObjectName(strPrefix As String, intType As Integer) As String

  Dim cmdUniqObj As New ADODB.Command
  Dim pmADO As ADODB.Parameter
  
  With cmdUniqObj
    .CommandText = "sp_ASRUniqueObjectName"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon
              
    Set pmADO = .CreateParameter("UniqueObjectName", adVarChar, adParamOutput, 255)
    .Parameters.Append pmADO
  
    Set pmADO = .CreateParameter("Prefix", adVarChar, adParamInput, 255)
    .Parameters.Append pmADO
    pmADO.value = strPrefix
    
    Set pmADO = .CreateParameter("Type", adInteger, adParamInput)
    .Parameters.Append pmADO
    pmADO.value = intType
    
    Set pmADO = Nothing
    
    .Execute
    
    UniqueSQLObjectName = IIf(IsNull(.Parameters(0).value), vbNullString, .Parameters(0).value)
      
  End With

  Set cmdUniqObj = Nothing
  
End Function


Public Function DropUniqueSQLObject(sSQLObjectName As String, iType As Integer) As Boolean

  On Error GoTo ErrorTrap

  Dim cmdUniqObj As New ADODB.Command
  Dim pmADO As ADODB.Parameter
 
  If Len(sSQLObjectName) > 0 Then
    With cmdUniqObj
      .CommandText = "sp_ASRDropUniqueObject"
      .CommandType = adCmdStoredProc
      .CommandTimeout = 0
      Set .ActiveConnection = gADOCon
                
      Set pmADO = .CreateParameter("UniqueObjectName", adVarChar, adParamInput, 255)
      .Parameters.Append pmADO
      pmADO.value = sSQLObjectName
      
      Set pmADO = .CreateParameter("Type", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.value = iType
      
      Set pmADO = Nothing
      
      .Execute
    End With
  End If
  
  DropUniqueSQLObject = True
  
TidyUpAndExit:
  Set cmdUniqObj = Nothing
  Exit Function
ErrorTrap:
  DropUniqueSQLObject = False

End Function

