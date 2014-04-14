Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Text
Imports Microsoft.SqlServer.Server
Imports Assembly.General
Imports System.Transactions
Imports System.Runtime.InteropServices
Imports System.Globalization
Imports System.Threading


Partial Public Class Workflow
  Private Shared _initTrue As Boolean = False
  Private Shared _hiByte As Long = 0
  Private Shared _hiBound As Long = 0
  Private Shared _addTable(255, 255) As Byte
  Private Shared _xTable(255, 255) As Byte

  <Microsoft.SqlServer.Server.SqlFunction(Name:="udfASRNetGetWorkflowQueryString", DataAccess:=DataAccessKind.None)> _
    Public Shared Function GetQueryString(ByVal instanceID As Int32, ByVal stepID As Int32, _
         ByVal userPwd As String, ByVal server As String, ByVal database As String) As SqlString

    Dim key As String = "jmltn"
    Dim encryptedString As String
    Dim sourceString As String
    Dim user As String = ""
    Dim password As String = ""

    ReadWebLogon(userPwd, user, password)

    sourceString = String.Concat( _
                        instanceID, _
                            ControlChars.Tab, _
                        stepID, _
                            ControlChars.Tab, _
                        user, _
                            ControlChars.Tab, _
                        password, _
                            ControlChars.Tab, _
                        server, _
                            ControlChars.Tab, _
                        database)

        Dim sCultureName As String
        sCultureName = Thread.CurrentThread.CurrentCulture.Name

        '' I thought this was required, but not so sure now.
        '' Uncomment out if locales such as Hungarian, Japanese, etc. start getting 'Invalid QueryString' errors.
        'Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-gb")
        'Thread.CurrentThread.CurrentUICulture = CultureInfo.CreateSpecificCulture("en-gb")

    encryptedString = EncryptString(sourceString, key, True)
    encryptedString = CompactString(encryptedString)

        Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture(sCultureName)
        Thread.CurrentThread.CurrentUICulture = CultureInfo.CreateSpecificCulture(sCultureName)

        Return New SqlString(encryptedString)
        
  End Function

  Private Shared Function EncryptString(ByVal text As String, _
      Optional ByVal key As String = "", Optional ByVal outputInHex As Boolean = False) As String

    Dim arrInput() As Byte
    Dim arrKey() As Byte
    Dim arrOutput() As Byte

    text = text & " "
    arrInput = Encoding.Default.GetBytes(text)
    arrKey = Encoding.Default.GetBytes(key)

    arrOutput = EncryptByte(arrInput, arrKey)

    If outputInHex Then
      Return EnHex(arrOutput)
    Else
      Return Encoding.Default.GetString(arrOutput)
    End If

  End Function

  Private Shared Function EncryptByte(ByVal arrText() As Byte, ByVal arrKey() As Byte) As Byte()
    Dim arrTemp() As Byte
    Dim tempCount As Int32 = 0
    Dim loopCount As Int32 = 0
    Dim boundLength As Int32 = 0

    Call InitTbl()

    ReDim arrTemp((arrText.Length) + 3)

    Dim randomNum As New Random()
    arrTemp(0) = CType(randomNum.Next(1, 255), Byte)
    arrTemp(1) = CType(randomNum.Next(1, 255), Byte)
    arrTemp(2) = CType(randomNum.Next(1, 255), Byte)
    arrTemp(3) = CType(randomNum.Next(1, 255), Byte)
    arrTemp(4) = CType(randomNum.Next(1, 255), Byte)

    Encoding.Default.GetBytes(Text.Encoding.Default.GetString(arrText, 0, UBound(arrTemp) - 4)).CopyTo(arrTemp, 5)

    ReDim arrText(arrTemp.Length - 1)
    arrText = arrTemp
    ReDim arrTemp(0)
    boundLength = arrKey.Length - 2

    For loopCount = 0 To arrText.Length - 2
      If tempCount = boundLength Then tempCount = 0
      arrText(loopCount) = _xTable(arrText(loopCount), _addTable(arrText(loopCount + 1), arrKey(tempCount)))
      arrText(loopCount + 1) = _xTable(arrText(loopCount), arrText(loopCount + 1))
      arrText(loopCount) = _xTable(arrText(loopCount), _addTable(arrText(loopCount + 1), arrKey(tempCount + 1)))
      tempCount += 1
    Next loopCount

    Return arrText
  End Function

  Private Shared Sub InitTbl()
    Dim i As Int32 = 0
    Dim j As Int32 = 0

    If _initTrue Then Return

    For i = 0 To 255
      For j = 0 To 255
        _xTable(i, j) = CByte(i Xor j)
        _addTable(i, j) = CByte((i + j) Mod 255)
      Next j
    Next i

    _initTrue = True
  End Sub

  Private Shared Function EnHex(ByVal arrInput() As Byte) As String
    Dim i As Int32 = 0
    Dim output As New Text.StringBuilder()

    For i = 0 To arrInput.Length - 1
      output.Append(arrInput(i).ToString("X2"))
    Next

    Return output.ToString()
  End Function

  Private Shared Function CompactString(ByVal sourceString As String) As String
    ' Compact the encrypted string.
    ' psSourceString is a string of the hexadecimal values of the Ascii codes for each character in the encrypted string.
    ' In this string each character in the encrypted string is represented as 2 hex digits.
    ' As it's a string of hex characters all characters are in the range 0-9, A-F
    ' Valid hypertext link characters are 0-9, A-Z, a-z and some others (we'll be using $ and @).
    ' Take advantage of this by implementing our own base64 encoding as follows:
    Dim compactedString As String = String.Empty
    Dim sSubString As String = String.Empty
    Dim modifiedSourceString As String = String.Empty
    Dim iValue As Int32 = 0
    Dim iTemp As Int32 = 0
    Dim newString As String = String.Empty

    modifiedSourceString = sourceString
    Do While modifiedSourceString.Length > 0
      ' Read the hex characters in chunks of 3 (ie. possible values 0 - 4095)
      ' This chunk of 3 Hex characters can then be translated into 2 base64 characters (ie. still have possible values 0 - 4095)
      ' Woohoo! We've reduced the length of the encrypted string by about one third!
      newString = String.Empty
      sSubString = (modifiedSourceString & "000").Substring(0, 3)
      'sModifiedSourceString = Mid(sModifiedSourceString, 4)
      Try
        modifiedSourceString = modifiedSourceString.Substring(3)
      Catch
        modifiedSourceString = String.Empty
      End Try
      iValue = CInt("&H" & sSubString)

      ' Use our own base64 digit set.
      ' Base64 digit values 0-9 are represented as 0-9
      ' Base64 digit values 10-35 are represented as A-Z
      ' Base64 digit values 36-61 are represented as a-z
      ' Base64 digit value 62 is represented as $
      ' Base64 digit value 63 is represented as @

      iTemp = iValue Mod 64
      If iTemp = 63 Then
        newString = "@"
      ElseIf iTemp = 62 Then
        newString = "$"
      ElseIf iTemp >= 36 Then
        newString = Convert.ToChar(iTemp + 61)
      ElseIf iTemp >= 10 Then
        newString = Convert.ToChar(iTemp + 55)
      Else
        newString = Convert.ToChar(iTemp + 48)
      End If

      iTemp = CInt((iValue - iTemp) / 64)

      If iTemp = 63 Then
        newString = "@" & newString
      ElseIf iTemp = 62 Then
        newString = "$" & newString
      ElseIf iTemp >= 36 Then
        newString = Convert.ToChar(iTemp + 61) & newString
      ElseIf iTemp >= 10 Then
        newString = Convert.ToChar(iTemp + 55) & newString
      Else
        newString = Convert.ToChar(iTemp + 48) & newString
      End If

      compactedString = compactedString & newString
    Loop

    ' Append the number of characters to ignore, to the compacted string
    Return compactedString & CStr((3 - (sourceString.Length Mod 3)) Mod 3)

  End Function

  Private Shared Sub ReadWebLogon(ByVal input As String, ByRef userName As String, ByRef password As String)

    Dim eKey As String
    Dim lens As String
    Dim start As Int32 = 0
    Dim finish As Int32 = 0

    start = input.Length - 12
    eKey = input.Substring(start, 10)
    lens = input.Substring(input.Length - 2)
    input = XOREncript(input.Substring(0, start), eKey)

    start = 0
    finish = Asc(lens.Substring(0, 1)) - 127
    userName = input.Substring(start, finish)

    start = start + finish
    finish = Asc(lens.Substring(1, 1)) - 127
    password = input.Substring(start, finish)

  End Sub

  <Microsoft.SqlServer.Server.SqlProcedure(Name:="spASRStoredDataFileActions")> _
  Public Shared Sub StoredDataFileActions(ByVal piInstanceID As Integer, _
      ByVal piElementID As Integer, _
      ByVal piRecordID As Integer)

    Dim sUserName As String = String.Empty
    Dim sPassword As String = String.Empty
    Dim sDatabaseName As String = String.Empty
    Dim sServerName As String = String.Empty
    Dim sConnectString As String = String.Empty
    Dim sSQL As String = String.Empty
    Dim drColumnsToBeUpdated As System.Data.SqlClient.SqlDataReader
    Dim drGetFile As System.Data.SqlClient.SqlDataReader
    Dim iValueType As Integer
    Dim sHeader As String
    Dim sTemp As String
    Dim fFileOK As Boolean
    Dim conn As System.Data.SqlClient.SqlConnection
    Dim iStoredDataColumnDefnID As Integer
    Dim iColumnID As Integer
    Dim abtReadFile As Byte()
    Dim abtWriteFile As Byte()
    Dim abtTemp As Byte()
    Dim abtHeader As Byte()
    Dim iTemp As Integer
    Dim sFileName As String
    Dim sShortFileName As String
    Dim dtNow As DateTime
    Dim sTableName As String
    Dim sColumnName As String
    Dim iColumnOLEType As Integer
    Dim fImageTypeOLEColumn As Boolean
    Dim sWriteFileName As String

    ' Get the logon details, and create the connection string.
    Dim sSystemLogon As String = GetSystemLogon()

    If sSystemLogon = String.Empty Then
      sConnectString = GetConnectionString("", "", ContextDatabaseName, ContextServerName, "OpenHR Workflow")
    Else
      sSystemLogon = ProcessEncryptedString(sSystemLogon)
      DecryptLogonDetails(sSystemLogon, sUserName, sPassword, sDatabaseName, sServerName)
      sConnectString = GetConnectionString(sUserName, sPassword, ContextDatabaseName, ContextServerName, "OpenHR Workflow")
    End If

    ' Ensure the connection string allows multiple active resultsets.
    Dim builder As New SqlConnectionStringBuilder(sConnectString)
    builder.MultipleActiveResultSets = True
    sConnectString = builder.ConnectionString

    Try
      Using scope As New TransactionScope(TransactionScopeOption.Suppress)
        ' Create the connection
        conn = New SqlClient.SqlConnection(sConnectString)

        Try
          conn.Open()
        Catch ex As SqlException
          Throw New Exception(String.Format("Cannot connect to database {0} on server {1} ", ContextDatabaseName, ContextServerName))
        End Try

        ' Get details of the OLE/photo columns that are to be updated in the given StoredData element.
        Dim cmdColumnsToBeUpdated As New SqlCommand()
        cmdColumnsToBeUpdated.Connection = conn
        cmdColumnsToBeUpdated.CommandType = CommandType.Text

        sSQL = "SELECT ISNULL(EC.ID, 0) AS [ID]," _
            & " ISNULL(SC.columnID, '') AS [columnID]," _
            & " ISNULL(SC.columnName, '') AS [columnName]," _
            & " ISNULL(ST.tableName, '') AS [tableName]," _
            & " ISNULL(EC.valueType, 0) AS [valueType]" _
            & " FROM ASRSysWorkflowElementColumns EC" _
            & " INNER JOIN ASRSysColumns SC ON EC.columnID = SC.columnID" _
            & "   AND (SC.dataType = -3 OR SC.dataType = -4)" _
            & " INNER JOIN ASRSysTables ST ON SC.tableID = ST.tableID" _
            & " WHERE EC.elementID = @piElementID"
        cmdColumnsToBeUpdated.Parameters.AddWithValue("@piElementID", piElementID)
        cmdColumnsToBeUpdated.CommandText = sSQL

        drColumnsToBeUpdated = cmdColumnsToBeUpdated.ExecuteReader

        While (drColumnsToBeUpdated.Read)
          ' Loop through for each OLE/photo column that is to be updated.
          fImageTypeOLEColumn = True
          iStoredDataColumnDefnID = CInt(drColumnsToBeUpdated("ID"))         ' ID of the StoredData column record (not the ID of the column being updated).
          iValueType = CInt(drColumnsToBeUpdated("valueType"))   ' 0=Fixed (not used here), 1=WFValue, 2=DBValue, 3=Calced (not used here)
          iColumnID = CInt(drColumnsToBeUpdated("columnID"))   ' ID of the column to be updated.
          sColumnName = CStr(drColumnsToBeUpdated("columnName")) ' Name of the column to be updated.
          sTableName = CStr(drColumnsToBeUpdated("tableName"))   ' Name of the table of the column to be updated.

          fFileOK = False
          sHeader = ""
          sFileName = ""
          sShortFileName = "<none>"
          sWriteFileName = ""
          ReDim abtReadFile(0)
          ReDim abtWriteFile(0)

          ' Get the file (if WFValue or embedded/linked type OLE) and/or details to be used in the update.
          Dim cmdGetFile As New SqlCommand()
          cmdGetFile = New SqlClient.SqlCommand("spASRWorkflowStoredDataFile", conn)
          cmdGetFile.CommandType = CommandType.StoredProcedure

          cmdGetFile.Parameters.AddWithValue("@piElementColumnID", iStoredDataColumnDefnID)
          cmdGetFile.Parameters.AddWithValue("@piInstanceID", piInstanceID)

          cmdGetFile.Parameters.Add("@piValueType", SqlDbType.Int).Direction = ParameterDirection.Output
          cmdGetFile.Parameters.Add("@psFileName", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
          cmdGetFile.Parameters.Add("@psErrorMessage", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
          cmdGetFile.Parameters.Add("@piOLEType", SqlDbType.Int).Direction = ParameterDirection.Output

          drGetFile = cmdGetFile.ExecuteReader

          drGetFile.Read()

          If drGetFile.HasRows Then
            ' Only interested in this for linked/embedded.
            If (Not drGetFile("file") Is DBNull.Value) Then
              abtReadFile = CType(drGetFile("file"), Byte())
              fFileOK = (abtReadFile.GetLength(0) > 0)
            End If
          End If

          drGetFile.Close()
          drGetFile = Nothing

          If iValueType = 1 Then ' Workflow Value
            If fFileOK Then
              sFileName = NullSafeString(cmdGetFile.Parameters("@psFileName").Value)

              ' OpenHR OLE FORMAT...
              '   8 characters denoting format version (eg "<<V002>>")
              '   2 characters denoting OLE type (02 = Embedded document, 03 = UNC link)
              '   70 characters denoting filename
              '   210 character denoting file path
              '   60 characters for UNC
              '   10 characters for file size
              '   20 characters for file create date (in format dd/MM/yyyy HH:MM:SS)
              '   20 characters for file last modified date (in format dd/MM/yyyy HH:MM:SS)
              '   Remainder of the data is the contents of the embedded document.
              ' Initialise the header stub with version (8 chars - '<<V002>>'), 
              ' and embedded file flag (2 chars - '2 ')
              sHeader = "<<V002>>2 "

              ' Add the file name to the header stub (70 chars)
              sTemp = sFileName
              iTemp = sFileName.LastIndexOf("\")
              If iTemp > 0 Then
                sTemp = sFileName.Substring(iTemp + 1)
              End If
              sHeader = sHeader & Left(sTemp.PadRight(70), 70)
              sShortFileName = sTemp

              ' Add an empty file path to the header stub (210 chars)
              sHeader = sHeader & Space(210)
              ' Add an empty UNC to the header stub (60 chars)
              sHeader = sHeader & Space(60)
              ' Add file size to the header stub (10 chars)
							'sHeader = sHeader & Space(10)
							sHeader = sHeader & Left(abtReadFile.Length.ToString().PadRight(10), 10)

              ' Add the file creation date (in format dd/MM/yyyy HH:MM:SS) to the header stub (20 chars)
              dtNow = Now
              sTemp = dtNow.Day.ToString.PadLeft(2, CChar("0")) _
                  & "/" _
                  & dtNow.Month.ToString.PadLeft(2, CChar("0")) _
                  & "/" _
                  & dtNow.Year.ToString.PadLeft(4, CChar("0")) _
                  & " " _
                  & dtNow.Hour.ToString.PadLeft(2, CChar("0")) _
                  & ":" _
                  & dtNow.Minute.ToString.PadLeft(2, CChar("0")) _
                  & ":" _
                  & dtNow.Second.ToString.PadLeft(2, CChar("0"))
              sHeader = sHeader & Left(sTemp.PadRight(20), 20)

              ' Add the file modified date (in format dd/MM/yyyy HH:MM:SS) to the header stub (20 chars)
              sHeader = sHeader & Left(sTemp.PadRight(20), 20)

              ' Create a new byte array with the header followed by the binary version of the file itself.
							ReDim abtWriteFile(abtReadFile.Length + 400 - 1)

              ' Copy the header into the array.
              sHeader = Left(sHeader.PadRight(400), 400)
              abtHeader = System.Text.Encoding.ASCII.GetBytes(sHeader)
              Array.ConstrainedCopy(abtHeader, 0, abtWriteFile, 0, 400)

              ' Copy the embedded document into the array.
              Array.ConstrainedCopy(abtReadFile, 0, abtWriteFile, 400, abtReadFile.Length)
            End If
          End If

          If iValueType = 2 Then ' DBValue 
            ' Get the OLE Type (0=LocalFolder, 1=ServerFolder, 2=LInked/Emdbedded
            iColumnOLEType = NullSafeInteger(cmdGetFile.Parameters("@piOLEType").Value)

            If (iColumnOLEType = 0) Or (iColumnOLEType = 1) Then
              'LocalFolder (0) or ServerFolder(1), so just need to use the file name, not the file itself.
              fFileOK = True
              fImageTypeOLEColumn = False
              sFileName = NullSafeString(cmdGetFile.Parameters("@psFileName").Value)
              sShortFileName = sFileName
              sWriteFileName = sFileName
            ElseIf fFileOK Then
              'Linked/Embedded(2), so need to use the byte array of the file itself.

              ' Rip the file name out of the header stub of the byte array.
              ReDim abtTemp(70)
              Array.ConstrainedCopy(abtReadFile, 10, abtTemp, 0, 70)
              sTemp = System.Text.Encoding.ASCII.GetString(abtTemp).Trim
              sFileName = Left(sTemp, sTemp.Length - 1).Trim
              sShortFileName = sFileName

              'AE20090402 Fault #13644
              'ReDim abtWriteFile(abtReadFile.Length)
              ReDim abtWriteFile(abtReadFile.Length - 1)
              Array.ConstrainedCopy(abtReadFile, 0, abtWriteFile, 0, abtReadFile.Length)
            End If
          End If

          ' Perform the required record update, using strings or byte arrays as required.
          Dim cmdWFUpdate As New SqlCommand()
          cmdWFUpdate.Connection = conn
          cmdWFUpdate.CommandType = CommandType.Text

          If fFileOK Then
            sSQL = "UPDATE " & sTableName _
                & " SET " & sColumnName & " = @imgFile" _
                & " WHERE ID = " & piRecordID.ToString

            If fImageTypeOLEColumn Then
              cmdWFUpdate.Parameters.Add("@imgFile", _
                SqlDbType.Image, abtWriteFile.Length).Value = abtWriteFile
            Else
              cmdWFUpdate.Parameters.Add("@imgFile", _
                SqlDbType.VarChar, 8000).Value = sWriteFileName
            End If
          Else
            sSQL = "UPDATE " & sTableName _
                & " SET " & sColumnName & " = " & CStr(IIf(fImageTypeOLEColumn, "null", "''")) _
                & " WHERE ID = " & piRecordID.ToString
          End If

          cmdWFUpdate.CommandText = sSQL
          cmdWFUpdate.ExecuteNonQuery()
          cmdWFUpdate.Dispose()

          ' Record the values used in the StoredData step in the InstanceValues table.
          Dim cmdWFUpdateIV As New SqlCommand()
          cmdWFUpdateIV.Connection = conn
          cmdWFUpdateIV.CommandType = CommandType.Text

          sSQL = "DELETE FROM ASRSysWorkflowInstanceValues" _
              & " WHERE instanceID = " & CStr(piInstanceID) _
              & " AND elementID = " & CStr(piElementID) _
              & " AND columnID = " & CStr(iColumnID)
          cmdWFUpdateIV.CommandText = sSQL
          cmdWFUpdateIV.ExecuteNonQuery()

          sSQL = "INSERT INTO ASRSysWorkflowInstanceValues" _
              & " (instanceID, elementID, identifier, columnID, value, emailID)" _
              & " VALUES (" & CStr(piInstanceID) _
              & "," & CStr(piElementID) _
              & ", ''," _
              & CStr(iColumnID) _
              & ", @sFile" _
              & ", 0)"

          cmdWFUpdateIV.CommandText = sSQL
          cmdWFUpdateIV.Parameters.Clear()
          cmdWFUpdateIV.Parameters.Add("@sFile", _
            SqlDbType.VarChar, 8000).Value = sShortFileName
          cmdWFUpdateIV.ExecuteNonQuery()

          cmdWFUpdateIV.Dispose()
        End While

        drColumnsToBeUpdated.Close()
        cmdColumnsToBeUpdated.Dispose()

        conn.Close()
        conn.Dispose()
      End Using
    Catch ex As SqlException
      Throw ex
    Catch ex As Exception
      Throw ex
    End Try

  End Sub

  <Microsoft.SqlServer.Server.SqlProcedure(Name:="spASRWorkflowInsertNewRecord")> _
  Public Shared Sub WorkflowInsertNewRecord(<Out()> ByRef newRecordID As Integer, ByVal tableID As Integer, ByVal insertString As String)

    Dim userName As String = String.Empty
    Dim password As String = String.Empty
    Dim databaseName As String = String.Empty
    Dim serverName As String = String.Empty
    Dim connectString As String = String.Empty

    Dim systemLogon As String = GetSystemLogon()

    If systemLogon = String.Empty Then
      connectString = GetConnectionString("", "", ContextDatabaseName, ContextServerName, "OpenHR Workflow")
    Else
      systemLogon = ProcessEncryptedString(systemLogon)
      DecryptLogonDetails(systemLogon, userName, password, databaseName, serverName)
      connectString = GetConnectionString(userName, password, ContextDatabaseName, ContextServerName, "OpenHR Workflow")
    End If

    Try
      Using scope As New TransactionScope(TransactionScopeOption.Suppress)
        Using conn As New SqlConnection(connectString)

          Using cmd As New SqlCommand("sp_ASRInsertNewRecord", conn)
            cmd.CommandType = CommandType.StoredProcedure

            cmd.Parameters.Add("@piNewRecordID", SqlDbType.Int).Direction = ParameterDirection.Output
            cmd.Parameters.AddWithValue("@psInsertString", insertString)

            Try
              cmd.Connection.Open()
            Catch ex As SqlException
              Throw New Exception(String.Format("Cannot connect to database {0} on server {1} ", ContextDatabaseName, ContextServerName))
            End Try

            cmd.ExecuteNonQuery()

            newRecordID = CType(cmd.Parameters("@piNewRecordID").Value, Integer)
          End Using
        End Using
      End Using
    Catch ex As SqlException
      Throw ex
    Catch ex As Exception
      Throw ex
    End Try

  End Sub

  <Microsoft.SqlServer.Server.SqlProcedure(Name:="spASRWorkflowUpdateRecord")> _
  Public Shared Sub WorkflowUpdateRecord(<Out()> ByRef result As Integer, ByVal tableID As Integer, ByVal updateString As String, _
                                         ByVal realSource As String, ByVal recordID As Integer)

    Dim userName As String = String.Empty
    Dim password As String = String.Empty
    Dim databaseName As String = String.Empty
    Dim serverName As String = String.Empty
    Dim connectString As String = String.Empty

    Dim systemLogon As String = GetSystemLogon()

    If systemLogon = String.Empty Then
      connectString = GetConnectionString("", "", ContextDatabaseName, ContextServerName, "OpenHR Workflow")
    Else
      systemLogon = ProcessEncryptedString(systemLogon)
      DecryptLogonDetails(systemLogon, userName, password, databaseName, serverName)
      connectString = GetConnectionString(userName, password, ContextDatabaseName, ContextServerName, "OpenHR Workflow")
    End If

    Try
      Using scope As New TransactionScope(TransactionScopeOption.Suppress)
        Using conn As New SqlConnection(connectString)

          Using cmd As New SqlCommand("sp_ASRUpdateRecord", conn)
            cmd.CommandType = CommandType.StoredProcedure

            cmd.Parameters.Add("@piResult", SqlDbType.Int).Direction = ParameterDirection.Output
            cmd.Parameters.AddWithValue("@psUpdateString", updateString)
            cmd.Parameters.AddWithValue("@piTableID", tableID)
            cmd.Parameters.AddWithValue("@psRealSource", realSource)
            cmd.Parameters.AddWithValue("@piID", recordID)
            cmd.Parameters.AddWithValue("@piTimestamp", DBNull.Value)

            Try
              cmd.Connection.Open()
            Catch ex As SqlException
              Throw New Exception(String.Format("Cannot connect to database {0} on server {1} ", ContextDatabaseName, ContextServerName))
            End Try

            cmd.ExecuteNonQuery()

            result = CType(cmd.Parameters("@piResult").Value, Integer)
          End Using
        End Using
      End Using
    Catch ex As SqlException
      Throw ex
    Catch ex As Exception
      Throw ex
    End Try
  End Sub

  <Microsoft.SqlServer.Server.SqlProcedure(Name:="spASRWorkflowDeleteRecord")> _
  Public Shared Sub WorkflowDeleteRecord(<Out()> ByRef result As Integer, ByVal tableID As Integer, ByVal realSource As String, _
                                         ByVal recordID As Integer)

    Dim userName As String = String.Empty
    Dim password As String = String.Empty
    Dim databaseName As String = String.Empty
    Dim serverName As String = String.Empty
    Dim connectString As String = String.Empty

    Dim systemLogon As String = GetSystemLogon()

    If systemLogon = String.Empty Then
      connectString = GetConnectionString("", "", ContextDatabaseName, ContextServerName, "OpenHR Workflow")
    Else
      systemLogon = ProcessEncryptedString(systemLogon)
      DecryptLogonDetails(systemLogon, userName, password, databaseName, serverName)
      connectString = GetConnectionString(userName, password, ContextDatabaseName, ContextServerName, "OpenHR Workflow")
    End If

    Try
      Using scope As New TransactionScope(TransactionScopeOption.Suppress)
        Using conn As New SqlConnection(connectString)

          Using cmd As New SqlCommand("sp_ASRDeleteRecord", conn)
            cmd.CommandType = CommandType.StoredProcedure

            cmd.Parameters.Add("@piResult", SqlDbType.Int).Direction = ParameterDirection.Output
            cmd.Parameters.AddWithValue("@piTableID", tableID)
            cmd.Parameters.AddWithValue("@psRealSource", realSource)
            cmd.Parameters.AddWithValue("@piID", recordID)

            Try
              cmd.Connection.Open()
            Catch ex As SqlException
              Throw New Exception(String.Format("Cannot connect to database {0} on server {1} ", ContextDatabaseName, ContextServerName))
            End Try

            cmd.ExecuteNonQuery()

            result = CType(cmd.Parameters("@piResult").Value, Integer)
          End Using
        End Using
      End Using
    Catch ex As SqlException
      Throw ex
    Catch ex As Exception
      Throw ex
    End Try

  End Sub

End Class
