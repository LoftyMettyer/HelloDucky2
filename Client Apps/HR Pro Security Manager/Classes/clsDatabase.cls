VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Instantiate internal classes
Private ODBC As New ODBC

'Local variables to hold property values
Private mvarDatabaseName As String
Private mvarKeyWords As String
Private mvarMaxTableNameLength As Integer
Private mvarMaxColumns As Integer
Private mvarMaxColumnNameLength As Integer
Private mvarName As String
Private mvarVersion As String
Private mvarUserName As String
Private mvarServerName As String

'''Public Property Get CurrentUserName() As String
'''  CurrentUserName = rdoEngine.rdoDefaultUser
'''
'''End Property
'''
'''Public Property Get servername() As String
'''  mvarServerName = ODBC.GetInfoString(rdoCon.hdbc, SQL_SERVER_NAME, 128)
'''  servername = mvarServerName
'''End Property
'''
'''Public Property Get UserName() As String
'''  mvarUserName = ODBC.GetInfoString(rdoCon.hdbc, SQL_USER_NAME, 128)
'''  UserName = mvarUserName
'''End Property
'''
'''
'''Public Property Get DatabaseName() As String
'''  mvarDatabaseName = ODBC.GetInfoString(rdoCon.hdbc, SQL_DATABASE_NAME, 128)
'''  DatabaseName = mvarDatabaseName
'''End Property
'''
'''Public Property Get KeyWords() As String
'''  mvarKeyWords = ODBC.GetInfoString(rdoCon.hdbc, SQL_KEYWORDS, 4096)
'''  KeyWords = mvarKeyWords
'''End Property

Public Property Get MaxColumnNameLength() As Integer
'  mvarMaxColumnNameLength = ODBC.GetInfo(rdoCon.hdbc, SQL_MAX_COLUMN_NAME_LEN)
'  MaxColumnNameLength = mvarMaxColumnNameLength
  MaxColumnNameLength = 100
  
End Property

'''Public Property Get MaxColumns() As Integer
'''  mvarMaxColumns = ODBC.GetInfo(rdoCon.hdbc, SQL_MAX_COLUMNS_IN_TABLE)
'''  MaxColumns = mvarMaxColumns
'''End Property

Public Property Get MaxTableNameLength() As Integer
'''  mvarMaxTableNameLength = ODBC.GetInfo(rdoCon.hdbc, SQL_MAX_TABLE_NAME_LEN)
'''  MaxTableNameLength = mvarMaxTableNameLength
  MaxTableNameLength = 100
  
End Property

'''Public Property Get Name() As String
'''  mvarName = ODBC.GetInfoString(rdoCon.hdbc, SQL_DBMS_NAME, 128)
'''  Name = mvarName
'''End Property
'''
'''Public Property Get Version() As String
'''  mvarVersion = ODBC.GetInfoString(rdoCon.hdbc, SQL_DBMS_VER, 128)
'''  Version = mvarVersion
'''End Property
'''
'''Public Sub ClearConnection()
'''  If Not rdoCon Is Nothing Then
'''    If rdoCon.hdbc > 0 Then
'''      rdoCon.Close
'''    End If
'''    Set rdoCon = Nothing
'''  End If
'''End Sub
'''
'''Public Sub ClearEnvironment()
'''  If Not rdoEnv Is Nothing Then
'''    If rdoEnv.henv > 0 Then
'''      rdoEnv.Close
'''    End If
'''    Set rdoEnv = Nothing
'''  End If
'''End Sub



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
      If blnUpper Then
        strName = strName + UCase(strChar)
      Else
        strName = strName + LCase(strChar)
      End If
    
      blnUpper = (strChar = "_")
    Next intChar
  End If

  FormatName = strName
  Exit Function
  
ErrorTrap:
  FormatName = vbNullString
  Err = False
  
End Function

'''Public Function GetTempTableName(ByVal TempName As String) As String
'''  Dim strName As String
'''  Dim intAttempt As Integer
'''
'''  strName = Left(TempName, MaxTableNameLength)
'''  intAttempt = 0
'''  Do While intAttempt < 1000 And TableExists(strName)
'''    intAttempt = intAttempt + 1
'''    strName = Left(TempName, MaxTableNameLength - 3)
'''    strName = strName & Right(Str(intAttempt + 1000), 3)
'''  Loop
'''
'''  If intAttempt < 1000 Then
'''    GetTempTableName = strName
'''  Else
'''    GetTempTableName = vbNullString
'''    MsgBox "Error creating temporary table name", _
'''      vbExclamation + vbOKOnly, Application.Name
'''  End If
'''
'''End Function

Public Function ColumnIsNumeric(ByVal plngColumnID As Long) As Boolean

  Dim sSQL As String
  Dim rsDataType As New ADODB.Recordset

  ColumnIsNumeric = True

  sSQL = "SELECT DataType from AsrSysColumns WHERE ColumnID = " & plngColumnID
  rsDataType.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

  If (rsDataType.BOF = False) And (rsDataType.EOF = False) Then
    Select Case rsDataType.Fields(0).Value
    Case 2, 4
      ColumnIsNumeric = True
    End Select
  End If

  rsDataType.Close
  Set rsDataType = Nothing

End Function

Public Function ColumnIsNumericName(ByVal psColumnName As String) As Boolean

  Dim sSQL As String
  Dim rsDataType As New ADODB.Recordset
  Dim lColumnID As Long
  
  ColumnIsNumericName = True

  sSQL = "SELECT ColumnID FROM AsrSysColumns WHERE ColumnName = '" & psColumnName & "'"
  rsDataType.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

  If rsDataType.BOF And rsDataType.EOF Then
    ColumnIsNumericName = False
  Else
    lColumnID = rsDataType.Fields(0).Value
  End If
  rsDataType.Close
    
  sSQL = "SELECT DataType from AsrSysColumns WHERE ColumnID = " & lColumnID
  rsDataType.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
  
  If (rsDataType.BOF = False) And (rsDataType.EOF = False) Then
    Select Case rsDataType.Fields(0).Value
    Case 2, 4
      ColumnIsNumericName = True
    End Select
  End If

  rsDataType.Close
  Set rsDataType = Nothing

End Function

Public Function IsKeyword(ByVal CheckWord As String) As Boolean
Attribute IsKeyword.VB_Description = "Checks if a specified string is a keyword of the current DBMS"
  On Error GoTo ErrorTrap
  
  Dim SQLcmd As String
  Dim rsKeywords As New ADODB.Recordset
    
  ' JDM - 26/11/01 - Fault 3211 - Causes problems if username has a ' in it
  CheckWord = Replace(CheckWord, "'", "''")
    
  'Build SQL string to extract any keywords matching the specified word
  SQLcmd = "SELECT Keyword FROM ASRSysKeywords" & _
    " WHERE Provider='Microsoft SQL Server'" & _
    " AND Keyword='" & CheckWord & "'"
  
  'Open keywords resultset
  rsKeywords.Open SQLcmd, gADOCon, adOpenKeyset, adLockReadOnly
  
  'If keywords resultset contains any records, then the word is a keyword
  If Not (rsKeywords.BOF And rsKeywords.EOF) Then
    rsKeywords.MoveLast
  End If
  If rsKeywords.RecordCount > 0 Then
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
  
  MsgBox ODBC.FormatError(Err.Description), _
    vbOKOnly + vbExclamation, Application.Name
  
  Err = False
  
End Function

'''Public Function TableExists(ByVal TableName As String) As Boolean
'''  Dim ThisTable As RDO.rdoTable
'''
'''  TableName = UCase(Trim(TableName))
'''
'''  'Refresh RDO tables collection
'''  rdoCon.rdoTables.Refresh
'''
'''  'Loop through RDO tables collection
'''  'to see if this table name exists.
'''  For Each ThisTable In rdoCon.rdoTables
'''    If UCase(ThisTable.Name) = TableName Then
'''      TableExists = True
'''      Exit For
'''    End If
'''  Next
'''
'''End Function

Public Function ValidNameChar(ByVal NameString As String, ByVal AsciiCode As Integer) As Integer
  On Error GoTo ErrorTrap
  
  If AsciiCode = Asc(" ") Then
    If Len(Trim(NameString)) > 1 Then
      AsciiCode = Asc("_")
    Else
      AsciiCode = 0
    End If
  Else
    If Not (AsciiCode = 8 Or AsciiCode = Asc("_") Or _
        (AsciiCode >= Asc("0") And AsciiCode <= Asc("9") And Len(NameString) > 0) Or _
        (AsciiCode >= Asc("A") And AsciiCode <= Asc("Z")) Or _
        (AsciiCode >= Asc("a") And AsciiCode <= Asc("z"))) Then
      AsciiCode = 0
    End If
  End If
  
  ValidNameChar = AsciiCode
  Exit Function
  
ErrorTrap:
  ValidNameChar = 0
  Err = False
  
End Function

Private Sub Class_Terminate()
  'Release internal classes
  Set ODBC = Nothing
End Sub


Public Property Get UserIsSA() As Boolean
  ' Return TRUE if the current user is the System Administrator.
  UserIsSA = UCase(Trim(gsUserName)) = "SA"
  
End Property


'''Public Function UniqueUserName(Optional psName As String) As String
'''  ' Return a unique user name for the current database.
'''  ' Base the user name on the name passed in if there is one.
'''  Dim fOK As Boolean
'''  Dim iCount As Integer
'''  Dim iNextIndex As Integer
'''  Dim sSQL As String
'''  Dim sBaseName As String
'''  Dim sUserName As String
'''  Dim rsUsers As New ADODB.Recordset
'''  Dim asUsers() As String
'''
'''  iCount = 0
'''  ReDim asUsers(0)
'''
'''  ' Get the list of users of the current database.
'''  sSQL = "sp_ASRUserInfo"
'''
'''  rsUsers.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
'''  With rsUsers
'''    Do While Not .EOF
'''      If !UserName <> "" Then
'''        iNextIndex = UBound(asUsers) + 1
'''        ReDim Preserve asUsers(iNextIndex)
'''        asUsers(iNextIndex) = !UserName
'''      End If
'''
'''      .MoveNext
'''    Loop
'''
'''    .Close
'''  End With
'''  Set rsUsers = Nothing
'''
'''  sBaseName = IIf(IsMissing(psName), "User", psName)
'''  sUserName = sBaseName
'''
'''  Do While Not fOK
'''
'''    fOK = True
'''
'''    For iNextIndex = 1 To UBound(asUsers)
'''      If sUserName = asUsers(iNextIndex) Then
'''        iCount = iCount + 1
'''        sUserName = sBaseName & "_" & Trim(Str(iCount))
'''        fOK = False
'''        Exit For
'''      End If
'''    Next iNextIndex
'''
'''  Loop
'''
'''  UniqueUserName = sUserName
'''
'''End Function

Public Function GetPicture(plngPictureID As Long) As ADODB.Recordset
  Dim sSQL As String
  Dim rsPicture As New ADODB.Recordset
       
  sSQL = "SELECT picture" & _
    " FROM ASRSysPictures" & _
    " WHERE pictureID = " & Trim(Str(plngPictureID))

  rsPicture.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly, -1
  Set GetPicture = rsPicture

End Function

Public Sub DeleteRecord(psTableName As String, psIDField As String, plngID As Long)

  ' Delete the given record.
  Dim sSQL As String
    
  sSQL = "DELETE FROM " & psTableName & " WHERE " & psIDField & " = " & plngID
  gADOCon.Execute sSQL, , adExecuteNoRecords

End Sub

Public Property Get AuthenticationType()


End Property


