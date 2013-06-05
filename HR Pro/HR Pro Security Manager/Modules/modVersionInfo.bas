Attribute VB_Name = "modVersionInfo"
Option Explicit

''''
''''Private Const mbDisplayAllModules = False
''''
''''Private Enum MODULE_TYPES
''''  giMODULE_DATAMANGER = 1
''''  giMODULE_INTRANET = 4
''''  giMODULE_SECURITYMANAGER = 2
''''  giMODULE_SYSTEMMANAGER = 3
''''End Enum
''''
''''Private mrstVersionChanges As ADODB.Recordset
''''Private mstrFilename As String
''''

'Public Function FrameworkVersionOK(Optional pstrOldVersion As String, _
'    Optional pstrNewVersion As String) As Boolean
'
'  Dim bOK As Boolean
'  Dim fSuccessFlag As Boolean
'  Dim sMessage As String
'
'  If glngSQLVersion < 9 Then
'    FrameworkVersionOK = True
'    Exit Function
'  End If
'
'  On Error GoTo ErrorTrap
'
'  Dim cmADO As ADODB.Command
'  Dim pmADO As ADODB.Parameter
'
'  Set cmADO = New ADODB.Command
'  With cmADO
'    .CommandText = "spASRSysCheckFrameworkVersion"
'    .CommandType = adCmdStoredProc
'    .CommandTimeout = 0
'    Set .ActiveConnection = gADOCon
'
'    Set pmADO = .CreateParameter("@piSuccessFlag", adInteger, adParamOutput)
'    .Parameters.Append pmADO
'    Set pmADO = .CreateParameter("@psMessage", adVarChar, adParamOutput, 8000)
'    .Parameters.Append pmADO
'    Set pmADO = .CreateParameter("@psOldVersion", adVarChar, adParamOutput, 255)
'    .Parameters.Append pmADO
'    Set pmADO = .CreateParameter("@psNewVersion", adVarChar, adParamOutput, 255)
'    .Parameters.Append pmADO
'
'    cmADO.Execute
'
'    fSuccessFlag = CBool(.Parameters(0).Value)
'    sMessage = CStr(.Parameters(1).Value)
'    pstrOldVersion = IIf(IsNull(.Parameters(2).Value), vbNullString, .Parameters(2).Value)
'    pstrNewVersion = IIf(IsNull(.Parameters(3).Value), vbNullString, .Parameters(3).Value)
'  End With
'
'  Set pmADO = Nothing
'  Set cmADO = Nothing
'
'TidyUpAndExit:
'  FrameworkVersionOK = fSuccessFlag
'  Exit Function
'
'ErrorTrap:
'  bOK = False
'  Resume TidyUpAndExit
'
'End Function


''''Private Function GenerateUniqueName() As String
''''
''''  Dim strFileName As String
''''
''''  ' generate unique file
'''''  strFileName = Space(400)
'''''  GetTempFileName App.Path, "HRProVersionInfo", 0, strFileName
'''''  Kill strFileName
'''''  strFileName = Mid(Trim(strFileName), 1, Len(Trim(strFileName)) - 4) & "htm"
'''''  GenerateUniqueName = strFileName
''''
''''  strFileName = App.Path & "\" & Replace(gsUserName, " ", "") & "sec_VersionInfo.htm"
''''  GenerateUniqueName = strFileName
''''
''''End Function
''''
''''
''''Private Function GetVersions() As String()
''''
''''Dim bOK As Boolean
''''Dim astrVersion() As String
''''Dim iRecordCount As Integer
''''Dim rstVersionInfo As ADODB.Recordset
''''
''''ReDim Preserve astrVersion(0)
''''
''''iRecordCount = 0
''''
''''Set rstVersionInfo = modExpression.OpenRecordset("SELECT Distinct Version FROM asrSysVersionInformation ORDER BY Version")
''''
''''If Not (rstVersionInfo.BOF And rstVersionInfo.EOF) Then
''''  rstVersionInfo.MoveFirst
''''  Do While Not rstVersionInfo.EOF
''''    ReDim Preserve astrVersion(iRecordCount)
''''    astrVersion(iRecordCount) = rstVersionInfo.Fields("Version").Value
''''    rstVersionInfo.MoveNext
''''    iRecordCount = iRecordCount + 1
''''  Loop
''''End If
''''
''''GetVersions = astrVersion
''''
''''End Function
''''
''''
''''Private Function GetData(pstrVersion As String, piModule As MODULE_TYPES) As ADODB.Recordset
''''
''''  Dim bOK As Boolean
''''  Dim rstVersionInfo As ADODB.Recordset
''''  Dim strSQL As String
''''
''''  strSQL = "SELECT * FROM ASRSysVersionInformation " _
''''    & "WHERE Version = '" & pstrVersion & "' AND HRPro_Module_Code = " & piModule _
''''    & " ORDER BY Version"
''''  Set rstVersionInfo = modExpression.OpenRecordset(strSQL)
''''
''''  If rstVersionInfo.BOF And rstVersionInfo.EOF Then
''''    bOK = False
''''  End If
''''
''''  Set GetData = rstVersionInfo
''''
''''End Function
''''
''''Public Function VersionInfo_GenerateHTML() As String
''''
''''  Dim bOK As Boolean
''''  Dim iVersionsCount As Integer
''''  Dim iModulesCount As Integer
''''  Dim astrVersions() As String
''''  Dim astrModules() As String
''''  Dim strHTML As String
''''  Dim intFileNo As Integer
''''
''''  ' Get versions
''''  astrVersions = GetVersions
''''
''''  If UBound(astrVersions) = 0 Then
''''    VersionInfo_GenerateHTML = ""
''''    Exit Function
''''  End If
''''
''''  ' generate unique file
''''  mstrFilename = GenerateUniqueName
''''
''''  ' If filename specified already exists then delete it first.
''''  If Len(Dir(mstrFilename)) > 0 Then
''''    Kill mstrFilename
''''  End If
''''
''''  intFileNo = FreeFile
''''  Open mstrFilename For Output As intFileNo
''''
''''  ' Start document
''''  strHTML = "<HTML><BODY>"
''''  Print #intFileNo, strHTML
''''
''''  ' Heading
''''  strHTML = "<H1 style=""FONT-SIZE: 18pt; FONT-FAMILY: Tahoma"">HR Pro Version Information</H1>"
''''  Print #intFileNo, strHTML
''''
''''  For iVersionsCount = LBound(astrVersions) To UBound(astrVersions)
''''
''''    strHTML = "<TABLE cellpadding=12 style=""FONT-SIZE: 10pt; FONT-FAMILY: Tahoma"" cellSpacing=0 cellPadding=1 width=""100%"" border=1>" & vbCrLf _
''''              & "<TR bgcolor=gray style=""COLOR: white"">" _
''''              & "<TD><B>HR Pro Version : <B></B></B></TD>" _
''''              & "<TD><B>" & astrVersions(iVersionsCount) & "</B></TD></TR>"
''''    Print #intFileNo, strHTML
''''
''''    strHTML = "<TR><TD colSpan=2>"
''''    Print #intFileNo, strHTML
''''
''''    If mbDisplayAllModules Then
''''
''''      ' Get changes for the data manager
''''      Set mrstVersionChanges = GetData(astrVersions(iVersionsCount), giMODULE_DATAMANGER)
''''      HTMLOutputChanges "Data Manager", intFileNo
''''
''''    End If
''''
''''    ' Get changes for the security manager
''''    Set mrstVersionChanges = GetData(astrVersions(iVersionsCount), giMODULE_SECURITYMANAGER)
''''    HTMLOutputChanges "Security Manager", intFileNo
''''
''''    If mbDisplayAllModules Then
''''
''''      ' Get changes for the system manager
''''      Set mrstVersionChanges = GetData(astrVersions(iVersionsCount), giMODULE_SYSTEMMANAGER)
''''      HTMLOutputChanges "System Manager", intFileNo
''''
''''      ' Get changes for the Intranet
''''      Set mrstVersionChanges = GetData(astrVersions(iVersionsCount), giMODULE_INTRANET)
''''      HTMLOutputChanges "Intranet Manager", intFileNo
''''
''''    End If
''''
''''    strHTML = "</TD></TR>"
''''    Print #intFileNo, strHTML
''''
''''  Next iVersionsCount
''''
''''
''''  ' End document
''''  strHTML = "</BODY></HTML>"
''''  Print #intFileNo, strHTML
''''
''''  ' Close the final output file
''''  Close #intFileNo
''''
''''  VersionInfo_GenerateHTML = mstrFilename
''''
''''End Function
''''
''''Private Function HTMLOutputChanges(pstrModule As String, pintFileNo As Integer) As Boolean
''''
''''  Dim bChangesInForThisModule As Boolean
''''  Dim strHTML As String
''''  bChangesInForThisModule = Not (mrstVersionChanges.EOF And mrstVersionChanges.BOF)
''''
''''  If bChangesInForThisModule Then
''''    strHTML = "<P><TABLE cellpadding=5 border=0 style=""FONT-SIZE: 10pt; FONT-FAMILY: Tahoma WIDTH: 100%"" cellSpacing=1 cellPadding=1 width=""100%"" background="""" border=1>" & vbCrLf _
''''        & "<TR><TD><B><U>" & pstrModule & "</U></B></TD></TR>"
''''    Print #pintFileNo, strHTML
''''
''''    Do While Not mrstVersionChanges.EOF
''''      strHTML = "<TR><TD><B>" & Replace(mrstVersionChanges.Fields("Area").Value, "&", "&amp;") & " : </B>" _
''''          & mrstVersionChanges.Fields("Description").Value & "</TD></TR>" & vbCrLf
''''      Print #pintFileNo, strHTML
''''      mrstVersionChanges.MoveNext
''''    Loop
''''
''''    strHTML = "</TABLE></P>"
''''    Print #pintFileNo, strHTML
''''  Else
''''    strHTML = "<P><TABLE cellpadding=5 border=0 style=""FONT-SIZE: 10pt; FONT-FAMILY: Tahoma WIDTH: 100%"" cellSpacing=1 cellPadding=1 width=""100%"" background="""" border=1>" & vbCrLf _
''''        & "<TR><TD><B><U>" & pstrModule & "</U></B></TD></TR>" _
''''        & "<TR><TD>No changes in this version</TD></TR></TABLE></P>"
''''    Print #pintFileNo, strHTML
''''  End If
''''
''''End Function
