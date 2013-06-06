Attribute VB_Name = "modDatabase"
Option Explicit

'Local variable to hold property values
Private mvarServerName As String

Public Property Get MaxColumnNameLength() As Integer
  MaxColumnNameLength = 100
End Property

Public Property Get MaxTableNameLength() As Integer
  MaxTableNameLength = 100
End Property

' Return whether or not the given data type requires a defined number of decimals.
Public Function ColumnHasScale(ByVal DataType As DataTypes) As Boolean
  ColumnHasScale = (DataType = dtNUMERIC)
End Function

' Return whether or not the given data type requires a defined size.
Public Function ColumnHasSize(ByVal DataType As DataTypes) As Boolean

  Select Case DataType
    Case dtVARCHAR, dtNUMERIC
      ColumnHasSize = True
    Case Else
      ColumnHasSize = False
  End Select

End Function

Public Function FormatName(ByVal NameString As String) As String
  On Error GoTo ErrorTrap
  
  Dim strName As String, strChar As String
  Dim intChar As Integer
  Dim blnUpper As Boolean
  
  NameString = Trim(NameString)
  strName = ""
  If Len(NameString) > 0 Then
    blnUpper = True
    For intChar = 1 To Len(NameString)
      strChar = Mid(NameString, intChar, 1)
      strName = strName + strChar
      blnUpper = (strChar = "_")
    Next intChar
  End If

  FormatName = strName
  Exit Function
  
ErrorTrap:
  FormatName = vbNullString
  Err = False
  
End Function
'''
Public Function GetDataDesc(ByVal DataType As DataTypes) As String
  'Returns the datatype description for the specified datatype.
  On Error GoTo ErrorTrap

  Dim sDesc As String

  Select Case DataType
    Case dtVARCHAR
      sDesc = "Character"

    Case dtINTEGER
      sDesc = "Integer"

    Case dtNUMERIC
      sDesc = "Numeric"

    Case dtBIT
      sDesc = "Logic"

    Case dtLONGVARBINARY
      sDesc = "OLE"

    Case dtVARBINARY
      sDesc = "Photo"

    Case dtTIMESTAMP
      sDesc = "Date"

    Case dtLONGVARCHAR
      sDesc = "Working Pattern"

    Case Else
      sDesc = vbNullString
  End Select

  GetDataDesc = sDesc

  Exit Function

ErrorTrap:
  GetDataDesc = vbNullString
  Err = False

End Function

Public Function GetTempTableName(ByVal TempName As String) As String
  Dim strName As String
  Dim intAttempt As Integer
  
  strName = Left(TempName, MaxTableNameLength)
  intAttempt = 0
  Do While intAttempt < 1000 And TableExists(strName)
    intAttempt = intAttempt + 1
    strName = Left(TempName, MaxTableNameLength - 3)
    strName = strName & Right(Str(intAttempt + 1000), 3)
  Loop
  
  If intAttempt < 1000 Then
    GetTempTableName = strName
  Else
    GetTempTableName = vbNullString
    MsgBox "Error creating temporary table name!", vbExclamation + vbOKOnly, Application.Name
  End If
  
End Function

Public Function IsKeyword(ByVal CheckWord As String) As Boolean
  On Error GoTo ErrorTrap
  
  Dim sSQL As String
  Dim rsKeywords As New ADODB.Recordset
    
  ' Open the keywords resultset.
  sSQL = "SELECT keyword FROM ASRSysKeywords" & _
    " WHERE provider='Microsoft SQL Server'" & _
    " AND keyword='" & CheckWord & "'"
  rsKeywords.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  
  ' If the keywords resultset contains any records
  ' then the word is a keyword.
  If Not (rsKeywords.BOF And rsKeywords.EOF) Then
    IsKeyword = True
  Else
    'Now check the word against ODBC keywords
    'IsKeyword = (InStr(1, KeyWords, "," & CheckWord & ",", vbTextCompare) > 0)
  End If
  
  'Close and release keywords recordset
  rsKeywords.Close
  Set rsKeywords = Nothing
  
  Exit Function

ErrorTrap:
  IsKeyword = False
  
  MsgBox ODBC.FormatError(Err.Description), vbOKOnly + vbExclamation, Application.Name
  
  Err = False
  
End Function

Public Function TableExists(ByVal psTableName As String) As Boolean
  Dim sSQL As String
  Dim rsInfo As New ADODB.Recordset
  
  sSQL = "SELECT COUNT(*) AS recCount" & _
    " FROM sysobjects" & _
    " WHERE name = '" & psTableName & "'"
  rsInfo.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

  TableExists = (rsInfo.Fields(0).value > 0)
  rsInfo.Close
  Set rsInfo = Nothing
  
End Function

' Gets the next identity for a metadata table from the database
Public Function GetNextObjectIdentitySeed(ByRef ObjectType As String) As Long

  On Error GoTo ErrorTrap

  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim rstValues As ADODB.Recordset
  
  Dim lngUniqueValue As Long

  Set cmADO = New ADODB.Command
  With cmADO
    .CommandText = "spASRGetNextObjectIdentitySeed"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
    Set .ActiveConnection = gADOCon

    Set pmADO = .CreateParameter("viewname", adVarChar, adParamInput, 255)
    pmADO.value = ObjectType
    .Parameters.Append pmADO

    Set pmADO = .CreateParameter("nextid", adInteger, adParamOutput)
    .Parameters.Append pmADO

    cmADO.Execute

    lngUniqueValue = IIf(IsNull(.Parameters(1).value), 0, .Parameters(1).value)
  End With

TidyUpAndExit:
  GetNextObjectIdentitySeed = lngUniqueValue
  Set pmADO = Nothing
  Set cmADO = Nothing
  Exit Function

ErrorTrap:
  lngUniqueValue = 0
  Resume TidyUpAndExit

End Function

Public Function UniqueFileName(sTableName As String, sFileName As String) As Boolean
  On Error GoTo ErrorTrap

  Dim recUniqueValue As DAO.Recordset
  Dim sSQL As String
  UniqueFileName = False
  
  ' Create a record set of CountNames.
  sSQL = "SELECT COUNT(Name) as CountNames From " & sTableName & " where Name='" & sFileName & "'"
  Set recUniqueValue = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  ' Record value.
  UniqueFileName = recUniqueValue.Fields("CountNames") = 0
  ' Close the temporary recordset.
  recUniqueValue.Close

TidyUpAndExit:
  ' Disassociate object variables.
  Set recUniqueValue = Nothing
  Exit Function

ErrorTrap:
  UniqueFileName = False
  Resume TidyUpAndExit

End Function

Public Function UniqueColumnValue(sTableName As String, sColumnName As String) As Long
  On Error GoTo ErrorTrap

  Dim lngUniqueValue As Long
  Dim recUniqueValue As DAO.Recordset
  Dim sSQL As String

  ' Create a record set with a unique value for the given table and column.
  sSQL = "SELECT MAX(" & sColumnName & ")+1 AS newValue" & _
    " FROM " & sTableName
  Set recUniqueValue = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)



  ' Read the unique column value from temporary recordset.
  If IsNull(recUniqueValue.Fields("newValue")) Then
    lngUniqueValue = 1
  Else
    lngUniqueValue = recUniqueValue.Fields("newValue")
  End If

  ' Close the temporary recordset.
  recUniqueValue.Close

TidyUpAndExit:
  ' Disassociate object variables.
  Set recUniqueValue = Nothing
  'Return the unique column value.
  UniqueColumnValue = lngUniqueValue
  Exit Function

ErrorTrap:
  lngUniqueValue = 1
  Resume TidyUpAndExit

End Function

Public Function ValidNameChar(ByVal piAsciiCode As Integer, ByVal piPosition As Integer) As Integer
  On Error GoTo ErrorTrap
  
  ' Validate the characters used to create table and column names.
  
  If piAsciiCode = Asc(" ") Then
    ' Substitute underscores for spaces.
    If piPosition <> 0 Then
      piAsciiCode = Asc("_")
    Else
      piAsciiCode = 0
    End If
  Else
    ' Allow only pure alpha-numerics and underscores.
    ' Do not allow numerics in the first chracter position.
    If Not (piAsciiCode = 8 Or piAsciiCode = Asc("_") Or _
      (piAsciiCode >= Asc("0") And piAsciiCode <= Asc("9") And piPosition <> 0) Or _
      (piAsciiCode >= Asc("A") And piAsciiCode <= Asc("Z")) Or _
      (piAsciiCode >= Asc("a") And piAsciiCode <= Asc("z"))) Then
      piAsciiCode = 0
    End If
  End If
  
  ValidNameChar = piAsciiCode
  Exit Function
  
ErrorTrap:
  ValidNameChar = 0
  Err = False
  
End Function

Public Function ValidateName(ByVal psName As String) As String
  ' Replace spaces with underscores.
  ' Remove any non-alpha-numeric characters.
  On Error GoTo ErrorTrap
  
  Dim iCounter As Integer
  Dim sNewName As String
  Dim iAsciiCode As Integer
  
  sNewName = ""
  
  For iCounter = 1 To Len(psName)
    iAsciiCode = ValidNameChar(Asc(Mid(psName, iCounter, 1)), Len(sNewName))
    
    If iAsciiCode > 0 Then
      sNewName = sNewName & Chr(iAsciiCode)
    End If
  Next iCounter

  ValidateName = sNewName
  Exit Function

ErrorTrap:
  ValidateName = ""
  Err = False
  
End Function


