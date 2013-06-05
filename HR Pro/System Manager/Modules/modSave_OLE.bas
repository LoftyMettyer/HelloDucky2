Attribute VB_Name = "modSave_OLE"
Option Explicit

Private Const OLE_VERSION = 2

Private Const adTypeText = 2
Private Const adTypeBinary = 1

Public Function UpgradeOLEDataToV2() As Boolean

  Dim rsColumns As ADODB.Recordset
  Dim rsTempInfo As ADODB.Recordset
  Dim bOK As Boolean
  Dim iCurrentVersion As Integer
  
  bOK = True

  ' Get current version of structures
  iCurrentVersion = GetSystemSetting("Database", "OLEStructure", 1)
  
  If iCurrentVersion < OLE_VERSION Then
   
    ' Turn triggers off
    Set rsTempInfo = New ADODB.Recordset
    rsTempInfo.Open "SELECT @@SPID", gADOCon, adOpenForwardOnly, adLockReadOnly, adCmdText
    SaveSystemSetting "Database", "UpdateLoginColumnSPID", rsTempInfo.Fields(0).Value
    rsTempInfo.Close
  
    Set rsColumns = New ADODB.Recordset
    rsColumns.Open "SELECT c.ColumnID, t.TableName, c.ColumnName" & _
      " FROM ASRSysColumns c" & _
      " INNER JOIN ASRSysTables t ON t.TableID = c.TableID" & _
      " WHERE c.DataType IN (-3,-4) AND c.OLEType IN (2,3)", gADOCon, adOpenDynamic, adLockReadOnly
    
    OutputCurrentProcess2 vbNullString, rsColumns.RecordCount
    
    With rsColumns
      While Not (.EOF Or .BOF)
        OutputCurrentProcess2 .Fields(1).Value & "." & .Fields(2).Value
      
        UpdateOLEColumnData .Fields(0).Value, .Fields(1).Value, .Fields(2).Value
        gobjProgress.UpdateProgress2
        
        .MoveNext
      Wend
    End With
  
    rsColumns.Close
  
    ' Update current OLE version structure
    SaveSystemSetting "Database", "OLEStructure", OLE_VERSION
    
  End If

TidyUpAndExit:
  
  ' Remove bypass on personnel trigger bypass
  SaveSystemSetting "Database", "UpdateLoginColumnSPID", 0
  
  Set rsColumns = Nothing
  Set rsTempInfo = Nothing
  UpgradeOLEDataToV2 = bOK
  Exit Function

ErrorTrap:
  bOK = False
  Resume TidyUpAndExit

End Function

Private Function UpdateOLEColumnData(ByRef plngColumnID As Long, ByRef psTableName As String, ByRef psSQLColumnName As String) As Boolean

  Dim rsADOStream As ADODB.Stream
  Dim rsData As ADODB.Recordset
  Dim sSQL As String
  Dim lngRecordID As Long
  Dim sFileName As String
  Dim bOK As Boolean
  
  On Error GoTo ErrorTrap
  
  bOK = True
  
  Set rsData = New ADODB.Recordset
  sSQL = "SELECT ID, " & psSQLColumnName & " FROM " & psTableName & " WHERE " & psSQLColumnName & " IS NOT NULL"
  rsData.Open sSQL, gADOCon, adOpenDynamic, adLockReadOnly, adCmdText

  With rsData
    While Not (.BOF Or .EOF)

      lngRecordID = .Fields(0).Value

      Set rsADOStream = New ADODB.Stream
      rsADOStream.Type = adTypeBinary
      rsADOStream.Open
      rsADOStream.Write rsData.Fields(1).Value
           
      UpgradeStream rsADOStream
      CommitStream rsADOStream, plngColumnID, lngRecordID

      rsADOStream.Close
    
      .MoveNext

    Wend
  End With

  rsData.Close

TidyUpAndExit:
   
  Set rsData = Nothing
  Set rsADOStream = Nothing
  UpdateOLEColumnData = bOK
  Exit Function

ErrorTrap:
  bOK = False
  Resume TidyUpAndExit

End Function
'
'Private Function GenerateDocumentFromStream(ByRef mobjStream As ADODB.Stream) As String
'
'  Dim objTextStream As TextStream
'
'  Dim strTempFileName As String
'  Dim strProperties As String
'  Dim objDocumentStream As ADODB.Stream
'
'  Set objDocumentStream = New ADODB.Stream
'
'Dim mstrTempPath As String
'mstrTempPath = Space(1024)
'Call GetTempPath(1024, mstrTempPath)
'
'Dim mstrFileName As String
'mstrFileName = GetTmpFName
'
'  ' Save the document information to file to read in.
'  strTempFileName = mstrFileName
'
'  If mobjStream.State = adStateClosed Then
'    mobjStream.Open
'    mobjStream.Type = adTypeBinary
'  End If
'
'  ' Setup new document stream
'  objDocumentStream.Type = adTypeBinary
'  objDocumentStream.Open
'
'  ' Copy out the document part of the stream
'  mobjStream.Position = 300
'  mobjStream.CopyTo objDocumentStream, mobjStream.Size - 300
'  objDocumentStream.SaveToFile strTempFileName, adSaveCreateOverWrite
'
'  GenerateDocumentFromStream = strTempFileName
'
'  objDocumentStream.Close
'  Set objDocumentStream = Nothing
'
'End Function

Private Function UpgradeStream(ByRef pobjStream As ADODB.Stream) As Boolean

  Dim strTempFile As String
  Dim objHeader As ADODB.Stream
  Dim objDocumentStream As ADODB.Stream
  Dim sHeader As String
  Dim sNewHeader As String

  If pobjStream.State = adStateClosed Then
    pobjStream.Open
    pobjStream.Type = adTypeBinary
  End If

  ' Copy out the header part of the stream
  Set objHeader = New ADODB.Stream
  objHeader.Type = adTypeBinary
  objHeader.Open
  pobjStream.Position = 0
  pobjStream.CopyTo objHeader, 300
   
  ' Copy out the document part of the stream
  Set objDocumentStream = New ADODB.Stream
  objDocumentStream.Type = adTypeBinary
  objDocumentStream.Open
  pobjStream.Position = 300
  pobjStream.CopyTo objDocumentStream, pobjStream.Size - 300

  ' Close original stream
  pobjStream.Close

  ' Adjust the header
  objHeader.Position = 0
  sHeader = Stream_BinaryToString(objHeader.Read)
  
  ' Double check, not to re-run
  If Mid(sHeader, 1, 8) = "<<V002>>" Then
    UpgradeStream = True
    Exit Function
  End If
  
  sNewHeader = "<<V002>>" & _
      Mid(sHeader, 9, 2) & _
      Left$(Trim(Mid(sHeader, 11, 50)) & String(70, " "), 70) & _
      Left$(Trim(Mid(sHeader, 61, 100)) & String(210, " "), 210) & _
      Left$(Trim(Mid(sHeader, 161, 50)) & String(50, " "), 60) & _
      Mid(sHeader, 201, 10) & _
      Mid(sHeader, 221, 20) & _
      Mid(sHeader, 241, 20)
  sNewHeader = Left$(Trim$(sNewHeader) & String(400, " "), 400)
  objHeader.Close
   
  ' Generate new stream
  pobjStream.Open
  pobjStream.Write Stream_StringToBinary(sNewHeader)
  objDocumentStream.Position = 0
  objDocumentStream.CopyTo pobjStream, objDocumentStream.Size
  
  objDocumentStream.Close
  Set objDocumentStream = Nothing



End Function


Private Function CommitStream(ByRef pobjStream As ADODB.Stream, _
  ByRef plngColumnID As Long, plngRecordID As Long) As Boolean

  Dim bOK As Boolean
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter

  On Error GoTo ErrorTrap
  bOK = True

    Set cmADO = New ADODB.Command
    With cmADO
    
      .CommandText = "spASRUpdateOLEField_" & plngColumnID
      .CommandType = adCmdStoredProc
      .CommandTimeout = 0
      Set .ActiveConnection = gADOCon
            
      Set pmADO = .CreateParameter("currentID", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = plngRecordID
                      
      Set pmADO = .CreateParameter("UploadFile", adLongVarBinary, adParamInput, -1)
      .Parameters.Append pmADO
    
      If pobjStream.State = adStateClosed Then
        pobjStream.Open
      End If
      
      If pobjStream.Size > 0 Then
        pobjStream.Position = 0
        pmADO.Value = pobjStream.Read
      Else
        pmADO.Value = Null
      End If
    
    End With
            
    cmADO.Execute
    
TidyUpAndExit:
  Set pmADO = Nothing
  Set cmADO = Nothing
  CommitStream = bOK
  Exit Function

ErrorTrap:
  bOK = False
  GoTo TidyUpAndExit

End Function

Function Stream_BinaryToString(ByRef Binary)
  
  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")
  
  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeBinary
  
  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.Write Binary
  
  
  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeText
  
  'Specify charset For the source text (unicode) data.
  BinaryStream.Charset = "us-ascii"
  
  'Open the stream And get binary data from the object
  Stream_BinaryToString = BinaryStream.ReadText
End Function

Function Stream_StringToBinary(ByRef Text As String)
  
  'Create Stream object
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")
  
  'Specify stream type - we want To save text/string data.
  BinaryStream.Type = adTypeText
  
  'Specify charset For the source text (unicode) data.
  BinaryStream.Charset = "us-ascii"
  
  'Open the stream And write text/string data To the object
  BinaryStream.Open
  BinaryStream.WriteText Text
  
  
  'Change stream type To binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeBinary
  
  'Ignore first two bytes - sign of
  BinaryStream.Position = 0
  
  'Open the stream And get binary data from the object
  Stream_StringToBinary = BinaryStream.Read
End Function

